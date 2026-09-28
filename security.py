"""
security.py — Phase 5
Security controls:
  - Path sandboxing with allowlist + traversal rejection
  - Secret field redaction
  - Resource limit checks
  - Tamper-evident append-only audit log
"""

from __future__ import annotations
import os
import re
import json
import time
import hashlib
import logging
from pathlib import Path
from typing import Any, List, Optional

from logger import log_audit

logger = logging.getLogger("excel_mcp")

# ── Allowed Directories Allowlist ─────────────────────────────────────────────

# Defaults: Desktop, Documents, Downloads for the current user.
# Override by setting EXCEL_MCP_ALLOWED_DIRS env var (semicolon-separated paths).
_DEFAULT_ALLOWED = [
    str(Path.home() / "Desktop"),
    str(Path.home() / "Documents"),
    str(Path.home() / "Downloads"),
    str(Path.home() / "OneDrive"),
    str(Path.home() / "OneDrive" / "Desktop"),
    str(Path.home() / "OneDrive" / "Documents"),
]


def _get_allowed_dirs() -> List[str]:
    env_val = os.environ.get("EXCEL_MCP_ALLOWED_DIRS", "")
    if env_val.strip():
        dirs = [d.strip() for d in env_val.split(";") if d.strip()]
        return [str(Path(d).resolve()) for d in dirs]
    return [str(Path(d).resolve()) for d in _DEFAULT_ALLOWED if Path(d).exists()]


def validate_path(path: str, operation: str = "access") -> str:
    """
    Validate that a file path is:
    1. Within the allowed directory allowlist
    2. Not a path traversal attempt
    3. Not excessively long

    Returns the resolved absolute path string on success.
    Raises ValueError with a clear message on failure.
    """
    if not path or not isinstance(path, str):
        raise ValueError("File path must be a non-empty string.")

    if len(path) > 500:
        raise ValueError("File path is too long (max 500 chars).")

    # Resolve to absolute, collapsing any ../ traversal
    try:
        resolved = str(Path(path).resolve())
    except Exception as e:
        raise ValueError(f"Cannot resolve path '{path}': {e}")

    # Check for path traversal patterns in the original string
    raw = path.replace("\\", "/")
    if "../" in raw or raw.startswith("../") or "/.." in raw:
        raise ValueError(f"Path traversal detected in: '{path}'")

    # Check against allowlist
    allowed = _get_allowed_dirs()
    if allowed:  # Only enforce if we have an allowlist
        in_allowed = any(
            resolved.lower().startswith(a.lower()) for a in allowed
        )
        if not in_allowed:
            raise ValueError(
                f"Path '{resolved}' is outside the allowed directories.\n"
                f"Allowed: {allowed}\n"
                f"Set EXCEL_MCP_ALLOWED_DIRS env var to override."
            )

    logger.debug("Path validated for %s: %s", operation, resolved)
    return resolved


def is_macro_enabled(path: str) -> bool:
    """Return True if the file extension indicates macro-enabled format."""
    return Path(path).suffix.lower() in (".xlsm", ".xlam", ".xltm", ".xlb")


def warn_macro_enabled(path: str) -> Optional[str]:
    """Return a warning string if the file is macro-enabled, else None."""
    if is_macro_enabled(path):
        return (
            f"⚠️  WARNING: '{Path(path).name}' is a macro-enabled workbook. "
            "Macros will NOT be enabled by the automation server. "
            "If this file requires macros to function correctly, results may be incomplete."
        )
    return None


# ── Secret Redaction ──────────────────────────────────────────────────────────

_SECRET_KEYS = {"password", "passwd", "pwd", "secret", "token", "key", "credential"}


def redact_secrets(args: dict) -> dict:
    """Return a copy of args dict with secret fields replaced by '***REDACTED***'."""
    return {
        k: ("***REDACTED***" if any(s in k.lower() for s in _SECRET_KEYS) else v)
        for k, v in args.items()
    }


# ── Resource Limits ───────────────────────────────────────────────────────────

MAX_CELLS_PER_OP = 500_000
MAX_FILE_SIZE_BYTES = 500 * 1024 * 1024  # 500 MB


def check_file_size(path: str) -> None:
    """Warn (not error) if file exceeds size threshold."""
    try:
        size = os.path.getsize(path)
        if size > MAX_FILE_SIZE_BYTES:
            logger.warning(
                "Workbook '%s' is %.1f MB — operations may be slow.",
                path, size / (1024 * 1024)
            )
    except Exception:
        pass


# ── Tamper-Evident Audit Log ──────────────────────────────────────────────────

_AUDIT_FILE = Path(os.path.dirname(__file__)) / "logs" / "tamper_evident_audit.jsonl"


def _read_last_hash() -> str:
    """Read the hash of the last audit entry for chain linking."""
    if not _AUDIT_FILE.exists():
        return "GENESIS"
    try:
        lines = _AUDIT_FILE.read_text(encoding="utf-8").strip().splitlines()
        if not lines:
            return "GENESIS"
        last = json.loads(lines[-1])
        return last.get("entry_hash", "GENESIS")
    except Exception:
        return "GENESIS"


def audit_write(
    operation: str,
    tool: str,
    args_summary: dict,
    before_state: Any = None,
    after_state: Any = None,
    result_status: str = "ok",
    workbook: str = None,
    sheet: str = None,
    range_address: str = None,
) -> None:
    """
    Append a tamper-evident audit entry to the audit log.
    Each entry includes a SHA-256 hash of its own content chained
    to the previous entry's hash.
    """
    _AUDIT_FILE.parent.mkdir(exist_ok=True)

    timestamp = time.strftime("%Y-%m-%dT%H:%M:%SZ", time.gmtime())
    prev_hash = _read_last_hash()

    entry = {
        "timestamp": timestamp,
        "operation": operation,
        "tool": tool,
        "workbook": workbook,
        "sheet": sheet,
        "range": range_address,
        "args": redact_secrets(args_summary) if isinstance(args_summary, dict) else args_summary,
        "before": before_state,
        "after": after_state,
        "result_status": result_status,
        "prev_hash": prev_hash,
    }

    # Compute entry hash for tamper-evidence
    entry_str = json.dumps(entry, sort_keys=True, default=str)
    entry_hash = hashlib.sha256(entry_str.encode("utf-8")).hexdigest()
    entry["entry_hash"] = entry_hash

    try:
        with open(_AUDIT_FILE, "a", encoding="utf-8") as f:
            f.write(json.dumps(entry, default=str) + "\n")
    except Exception as e:
        logger.error("Failed to write audit log: %s", e)

    # Also write to the plain audit log
    log_audit(operation, {"tool": tool, "status": result_status, "hash": entry_hash})


def verify_audit_chain() -> dict:
    """
    Verify the integrity of the tamper-evident audit log.
    Returns a dict with 'ok': True/False and 'broken_at' entry index if tampered.
    """
    if not _AUDIT_FILE.exists():
        return {"ok": True, "entries": 0, "message": "No audit log found."}

    try:
        lines = _AUDIT_FILE.read_text(encoding="utf-8").strip().splitlines()
    except Exception as e:
        return {"ok": False, "message": f"Cannot read audit log: {e}"}

    prev_hash = "GENESIS"
    for i, line in enumerate(lines):
        try:
            entry = json.loads(line)
            stored_hash = entry.pop("entry_hash", None)
            entry_str = json.dumps(entry, sort_keys=True, default=str)
            computed = hashlib.sha256(entry_str.encode("utf-8")).hexdigest()

            if stored_hash != computed:
                return {
                    "ok": False,
                    "broken_at": i,
                    "message": f"Entry {i} has been tampered with.",
                }
            if entry.get("prev_hash") != prev_hash:
                return {
                    "ok": False,
                    "broken_at": i,
                    "message": f"Chain broken at entry {i}: prev_hash mismatch.",
                }
            prev_hash = stored_hash
        except Exception as e:
            return {"ok": False, "broken_at": i, "message": f"Parse error at entry {i}: {e}"}

    return {"ok": True, "entries": len(lines), "message": "Audit chain is intact."}
