"""
shared_context.py — Canonical Contract & Enforcement Engine (Phase 1)
Provides:
  - Canonical tool response shape (§2.1)
  - Pre/Postcondition validation decorator
  - Idempotency key caching
  - Excel alert suppression context manager
  - Formula injection escaping & range validators
"""

from __future__ import annotations
import contextlib
import time
import json
import functools
import logging
from typing import Any, Dict, Optional, Callable

logger = logging.getLogger("excel_mcp")

# ── Idempotency Key Cache (Phase 1) ───────────────────────────────────────────
_IDEMPOTENCY_CACHE: Dict[str, Dict[str, Any]] = {}


def get_cached_idempotency_result(key: Optional[str]) -> Optional[dict]:
    if key and key in _IDEMPOTENCY_CACHE:
        logger.info("Idempotency key '%s' hit. Returning cached result.", key)
        return _IDEMPOTENCY_CACHE[key]
    return None


def set_cached_idempotency_result(key: Optional[str], result: dict) -> None:
    if key:
        _IDEMPOTENCY_CACHE[key] = result


def ok(message: str, data: Any = None) -> dict:
    result = {"status": "ok", "code": "SUCCESS", "message": message}
    if data is not None:
        result["data"] = data
    return result


def err(code: str, message: str, details: Any = None) -> dict:
    result = {"status": "error", "code": code, "message": message}
    if details is not None:
        result["details"] = details
    logger.error("[%s] %s | details=%s", code, message, details)
    return result

def canonical_response(
    status: str,  # "success" | "failed" | "partial"
    operation: str,
    workbook_id: Optional[str] = None,
    sheet: Optional[str] = None,
    input_params: Optional[dict] = None,
    result: Optional[dict] = None,
    verification: Optional[dict] = None,
    error: Optional[dict] = None,
    duration_ms: float = 0.0,
) -> dict:
    """
    Constructs the non-negotiable canonical response shape (§2.1).
    """
    return {
        "status": status,
        "operation": operation,
        "workbook_id": workbook_id or "wb_active",
        "sheet": sheet or "N/A",
        "input": input_params or {},
        "result": result or {},
        "verification": verification or {"passed": True, "checks": []},
        "error": error,
        "duration_ms": round(duration_ms, 2),
    }


def ok_canonical(
    operation: str,
    workbook_id: Optional[str] = None,
    sheet: Optional[str] = None,
    input_params: Optional[dict] = None,
    result: Optional[dict] = None,
    verification: Optional[dict] = None,
    duration_ms: float = 0.0,
) -> dict:
    return canonical_response(
        status="success",
        operation=operation,
        workbook_id=workbook_id,
        sheet=sheet,
        input_params=input_params,
        result=result,
        verification=verification or {"passed": True, "checks": ["execution_completed"]},
        error=None,
        duration_ms=duration_ms,
    )


def err_canonical(
    operation: str,
    code: str,
    message: str,
    workbook_id: Optional[str] = None,
    sheet: Optional[str] = None,
    input_params: Optional[dict] = None,
    recoverable: bool = True,
    rollback_available: bool = True,
    details: Any = None,
    duration_ms: float = 0.0,
) -> dict:
    error_obj = {
        "code": code,
        "message": message,
        "recoverable": recoverable,
        "rollback_available": rollback_available,
    }
    if details is not None:
        error_obj["details"] = details

    logger.error("[%s] Operation '%s' failed: %s", code, operation, message)
    return canonical_response(
        status="failed",
        operation=operation,
        workbook_id=workbook_id,
        sheet=sheet,
        input_params=input_params,
        result={},
        verification={"passed": False, "checks": ["execution_failed"]},
        error=error_obj,
        duration_ms=duration_ms,
    )


def result_to_str(r: dict) -> str:
    """Serialize canonical result dictionary to JSON string."""
    return json.dumps(r, default=str)


# ── Pre/Postcondition Decorator (§2.1 / Phase 1) ──────────────────────────────

def validate_tool_contract(
    operation_name: str,
    pre_checks: Optional[Callable[..., None]] = None,
    post_checks: Optional[Callable[..., None]] = None,
):
    """
    Decorator enforcing:
    1. Idempotency replay check
    2. Input pre-validation
    3. Error handling with canonical error output
    4. Duration tracking
    5. Result caching for idempotency
    """
    def decorator(fn):
        @functools.wraps(fn)
        def wrapper(*args, **kwargs):
            start_time = time.perf_counter()

            # Check idempotency
            idem_key = kwargs.get("idempotency_key")
            cached = get_cached_idempotency_result(idem_key)
            if cached:
                return cached

            # Extract common params
            workbook_id = kwargs.get("workbook_id") or kwargs.get("workbook")
            sheet = kwargs.get("sheet") or kwargs.get("source_sheet")

            # Run pre-condition checks if supplied
            if pre_checks:
                try:
                    pre_checks(*args, **kwargs)
                except Exception as pre_err:
                    elapsed = (time.perf_counter() - start_time) * 1000
                    res = err_canonical(
                        operation=operation_name,
                        code="PRECONDITION_FAILED",
                        message=str(pre_err),
                        workbook_id=workbook_id,
                        sheet=sheet,
                        input_params=kwargs,
                        duration_ms=elapsed,
                    )
                    set_cached_idempotency_result(idem_key, res)
                    return res

            try:
                res = fn(*args, **kwargs)
                elapsed = (time.perf_counter() - start_time) * 1000

                # Ensure result is canonical dict
                if isinstance(res, dict) and "status" in res:
                    res["duration_ms"] = round(elapsed, 2)
                    if workbook_id and ("workbook_id" not in res or res["workbook_id"] == "wb_active"):
                        res["workbook_id"] = workbook_id
                else:
                    # Wrap string/legacy returns
                    res = ok_canonical(
                        operation=operation_name,
                        workbook_id=workbook_id,
                        sheet=sheet,
                        input_params=kwargs,
                        result={"message": str(res)},
                        duration_ms=elapsed,
                    )

                # Run post-condition checks if supplied
                if post_checks and res.get("status") == "success":
                    try:
                        post_checks(*args, **kwargs, result=res)
                    except Exception as post_err:
                        res["status"] = "failed"
                        res["verification"] = {"passed": False, "checks": ["postcondition_failed"]}
                        res["error"] = {
                            "code": "POSTCONDITION_FAILED",
                            "message": str(post_err),
                            "recoverable": True,
                            "rollback_available": True,
                        }

                set_cached_idempotency_result(idem_key, res)
                return res

            except Exception as exc:
                elapsed = (time.perf_counter() - start_time) * 1000
                res = err_canonical(
                    operation=operation_name,
                    code="EXECUTION_ERROR",
                    message=f"{type(exc).__name__}: {str(exc)}",
                    workbook_id=workbook_id,
                    sheet=sheet,
                    input_params=kwargs,
                    duration_ms=elapsed,
                )
                set_cached_idempotency_result(idem_key, res)
                return res

        return wrapper
    return decorator


# ── Excel COM Suppression Context Manager ────────────────────────────────────

@contextlib.contextmanager
def excel_suppressed(app):
    old_alerts = None
    old_screen = None
    old_calc = None

    try:
        try:
            old_alerts = app.display_alerts
            app.display_alerts = False
        except Exception:
            pass

        try:
            old_screen = app.screen_updating
            app.screen_updating = False
        except Exception:
            pass

        try:
            old_calc = app.api.Calculation
            app.api.Calculation = -4135  # xlCalculationManual
        except Exception:
            pass

        yield

    finally:
        try:
            if old_alerts is not None:
                app.display_alerts = old_alerts
        except Exception:
            pass
        try:
            if old_screen is not None:
                app.screen_updating = old_screen
        except Exception:
            pass
        try:
            if old_calc is not None:
                app.api.Calculation = old_calc
                app.api.Calculate()
        except Exception:
            pass


# ── Resource Limits & Validation Helpers ─────────────────────────────────────

MAX_CELLS = 500_000

def check_data_size(data: list, max_cells: int = MAX_CELLS) -> None:
    if not data:
        return
    rows = len(data)
    cols = len(data[0]) if isinstance(data[0], (list, dict)) else 1
    if isinstance(data[0], dict):
        cols = len(data[0])
    total = rows * cols
    if total > max_cells:
        raise ValueError(
            f"Data payload has {total:,} cells ({rows} rows × {cols} cols) "
            f"which exceeds the limit of {max_cells:,}."
        )

_INJECTION_CHARS = ('=', '+', '-', '@', '\t', '\r')

def escape_injection(value: Any) -> Any:
    if isinstance(value, str) and value and value[0] in _INJECTION_CHARS:
        return "'" + value
    return value

def escape_data(data: list) -> list:
    result = []
    for row in data:
        if isinstance(row, list):
            result.append([escape_injection(cell) for cell in row])
        elif isinstance(row, dict):
            result.append({k: escape_injection(v) for k, v in row.items()})
        else:
            result.append(escape_injection(row))
    return result

def validate_range_address(address: str) -> str:
    if not address or not isinstance(address, str):
        raise ValueError("Range address must be a non-empty string.")
    address = address.strip()
    if len(address) > 100:
        raise ValueError(f"Range address too long: '{address[:30]}...'")
    return address.upper()

def validate_sheet_name(name: str) -> str:
    if not name or not isinstance(name, str):
        raise ValueError("Sheet name must be a non-empty string.")
    name = name.strip()
    forbidden = set(r'\/:*?[]')
    found = [c for c in name if c in forbidden]
    if found:
        raise ValueError(f"Sheet name contains invalid characters: {found}")
    if len(name) > 31:
        raise ValueError(f"Sheet name too long ({len(name)} chars, max 31): '{name}'")
    return name
