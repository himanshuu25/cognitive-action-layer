"""
logger.py — Phase 3
Structured per-tool logging with rotating file handler.
Every tool call is logged with: tool name, args summary, result status, duration.
"""

from __future__ import annotations
import logging
import logging.handlers
import os
import time
import functools
import json
from pathlib import Path

# ── Log Directory Setup ───────────────────────────────────────────────────────

_LOG_DIR = Path(os.path.dirname(__file__)) / "logs"
_LOG_DIR.mkdir(exist_ok=True)
_LOG_FILE = _LOG_DIR / "excel_mcp.log"

# ── Logger Configuration ──────────────────────────────────────────────────────

def _setup_logger() -> logging.Logger:
    log = logging.getLogger("excel_mcp")
    if log.handlers:
        return log  # Already configured

    log.setLevel(logging.DEBUG)

    fmt = logging.Formatter(
        fmt="%(asctime)s | %(levelname)-8s | %(name)s | %(message)s",
        datefmt="%Y-%m-%dT%H:%M:%S"
    )

    # Rotating file handler — 5 MB per file, keep 5 backups
    fh = logging.handlers.RotatingFileHandler(
        _LOG_FILE,
        maxBytes=5 * 1024 * 1024,
        backupCount=5,
        encoding="utf-8"
    )
    fh.setLevel(logging.DEBUG)
    fh.setFormatter(fmt)
    log.addHandler(fh)

    # Console handler — INFO and above
    ch = logging.StreamHandler()
    ch.setLevel(logging.INFO)
    ch.setFormatter(fmt)
    log.addHandler(ch)

    return log


logger = _setup_logger()


# ── Tool Call Logging Decorator ───────────────────────────────────────────────

def log_tool_call(tool_name: str = None):
    """
    Decorator that wraps an MCP tool function and logs:
    - Tool name and sanitized args at call time
    - Result status and elapsed time on return
    - Exception details on failure
    """
    def decorator(fn):
        name = tool_name or fn.__name__

        @functools.wraps(fn)
        def wrapper(*args, **kwargs):
            start = time.perf_counter()

            # Build a safe args summary (redact password fields)
            safe_kwargs = {
                k: ("***REDACTED***" if "password" in k.lower() else v)
                for k, v in kwargs.items()
            }
            try:
                args_summary = json.dumps(safe_kwargs, default=str)[:300]
            except Exception:
                args_summary = str(safe_kwargs)[:300]

            logger.info("CALL  [%s] args=%s", name, args_summary)

            try:
                result = fn(*args, **kwargs)
                elapsed = (time.perf_counter() - start) * 1000

                # Try to parse status from structured result
                status = "ok"
                if isinstance(result, dict):
                    status = result.get("status", "ok")
                elif isinstance(result, str):
                    try:
                        parsed = json.loads(result)
                        status = parsed.get("status", "ok")
                    except Exception:
                        status = "ok" if "error" not in result.lower()[:20] else "error"

                logger.info("DONE  [%s] status=%s elapsed=%.1fms", name, status, elapsed)
                return result

            except Exception as exc:
                elapsed = (time.perf_counter() - start) * 1000
                logger.error(
                    "FAIL  [%s] elapsed=%.1fms exception=%s: %s",
                    name, elapsed, type(exc).__name__, str(exc),
                    exc_info=True
                )
                raise

        return wrapper
    return decorator


def log_audit(operation: str, details: dict) -> None:
    """Write a structured audit-log entry for critical operations."""
    audit_logger = logging.getLogger("excel_mcp.audit")
    if not audit_logger.handlers:
        audit_file = _LOG_DIR / "audit.log"
        fh = logging.handlers.RotatingFileHandler(
            audit_file, maxBytes=10 * 1024 * 1024, backupCount=10, encoding="utf-8"
        )
        fh.setFormatter(logging.Formatter("%(asctime)s | %(message)s", "%Y-%m-%dT%H:%M:%S"))
        audit_logger.addHandler(fh)
        audit_logger.setLevel(logging.INFO)

    try:
        entry = json.dumps({"operation": operation, **details}, default=str)
    except Exception:
        entry = f"operation={operation} details={details}"

    audit_logger.info(entry)
