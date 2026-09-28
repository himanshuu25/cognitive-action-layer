"""
snapshot.py — Phase 4
Workbook snapshot (backup), restore, and backup→verify→swap pattern for safe critical edits.
"""

from __future__ import annotations
import os
import shutil
import time
from pathlib import Path
from typing import Callable, Optional

from logger import logger, log_audit


# ── Snapshot Directory ────────────────────────────────────────────────────────

def _get_snapshot_dir(workbook_path: str) -> Path:
    """Return the snapshot directory next to the workbook."""
    wb_path = Path(workbook_path).resolve()
    snap_dir = wb_path.parent / ".excel_mcp_snapshots"
    snap_dir.mkdir(exist_ok=True)
    return snap_dir


def _timestamp() -> str:
    return time.strftime("%Y%m%d_%H%M%S")


# ── Create / Restore Snapshot ─────────────────────────────────────────────────

def create_snapshot(workbook_path: str) -> str:
    """
    Save a timestamped backup copy of the workbook.

    Returns the path of the snapshot file.
    Raises IOError on failure.
    """
    wb_path = Path(workbook_path).resolve()
    if not wb_path.exists():
        raise FileNotFoundError(f"Workbook not found: '{workbook_path}'")

    snap_dir = _get_snapshot_dir(str(wb_path))
    stem = wb_path.stem
    suffix = wb_path.suffix
    snap_name = f"{stem}_snapshot_{_timestamp()}{suffix}"
    snap_path = snap_dir / snap_name

    shutil.copy2(str(wb_path), str(snap_path))
    logger.info("Snapshot created: '%s'", snap_path)
    log_audit(
        "SNAPSHOT_CREATE",
        {"workbook": str(wb_path), "snapshot": str(snap_path)}
    )
    return str(snap_path)


def restore_snapshot(snapshot_path: str, workbook_path: str) -> str:
    """
    Overwrite the workbook with the snapshot.

    Returns a success message.
    Raises IOError on failure.
    """
    snap = Path(snapshot_path).resolve()
    wb = Path(workbook_path).resolve()

    if not snap.exists():
        raise FileNotFoundError(f"Snapshot not found: '{snapshot_path}'")

    shutil.copy2(str(snap), str(wb))
    logger.info("Restored '%s' from snapshot '%s'.", wb, snap)
    log_audit(
        "SNAPSHOT_RESTORE",
        {"workbook": str(wb), "snapshot": str(snap)}
    )
    return f"Restored '{wb.name}' from snapshot '{snap.name}'."


def list_snapshots(workbook_path: str) -> list:
    """Return a list of snapshot paths for the given workbook, newest first."""
    wb_path = Path(workbook_path).resolve()
    snap_dir = _get_snapshot_dir(str(wb_path))
    stem = wb_path.stem
    snaps = sorted(
        snap_dir.glob(f"{stem}_snapshot_*{wb_path.suffix}"),
        key=lambda p: p.stat().st_mtime,
        reverse=True,
    )
    return [str(p) for p in snaps]


def cleanup_old_snapshots(workbook_path: str, keep: int = 10) -> int:
    """Delete old snapshots, keeping only the N most recent. Returns count deleted."""
    snaps = list_snapshots(workbook_path)
    to_delete = snaps[keep:]
    for p in to_delete:
        try:
            os.remove(p)
        except Exception:
            pass
    return len(to_delete)


# ── Backup → Operate → Verify → Swap ─────────────────────────────────────────

def safe_operation(
    workbook_path: str,
    operation: Callable[[str], dict],
    verify: Optional[Callable[[str], dict]] = None,
    keep_snapshot: bool = True,
) -> dict:
    """
    Executes a critical workbook operation with snapshot + rollback safety.

    Workflow:
    1. Save the workbook (caller must have it already saved/closed)
    2. Create a timestamped snapshot
    3. Execute operation(workbook_path) → expected to return a dict with 'ok' key
    4. Optionally run verify(workbook_path) → if it returns {ok: False}, rollback
    5. On failure at any step: restore from snapshot

    Args:
        workbook_path: Absolute path to the target workbook.
        operation:     Callable that performs the operation. Receives workbook_path.
        verify:        Optional callable to verify the result. Receives workbook_path.
        keep_snapshot: If True, keep the snapshot even on success (for audit).

    Returns:
        {
            "success": bool,
            "snapshot": str,
            "operation_result": dict,
            "verify_result": dict or None,
            "rolled_back": bool,
            "message": str,
        }
    """
    snapshot_path = None
    rolled_back = False

    try:
        # Step 1: Create snapshot
        snapshot_path = create_snapshot(workbook_path)

        # Step 2: Run operation
        try:
            op_result = operation(workbook_path)
        except Exception as e:
            logger.error("Operation failed, rolling back: %s", e)
            restore_snapshot(snapshot_path, workbook_path)
            rolled_back = True
            return {
                "success": False,
                "snapshot": snapshot_path,
                "operation_result": {"error": str(e)},
                "verify_result": None,
                "rolled_back": True,
                "message": f"Operation failed and was rolled back: {e}",
            }

        # Step 3: Verify (if provided)
        verify_result = None
        if verify is not None:
            try:
                verify_result = verify(workbook_path)
                if not verify_result.get("ok", True):
                    logger.error("Verification failed, rolling back.")
                    restore_snapshot(snapshot_path, workbook_path)
                    rolled_back = True
                    return {
                        "success": False,
                        "snapshot": snapshot_path,
                        "operation_result": op_result,
                        "verify_result": verify_result,
                        "rolled_back": True,
                        "message": "Verification failed; file restored from snapshot.",
                    }
            except Exception as e:
                logger.error("Verify step raised, rolling back: %s", e)
                restore_snapshot(snapshot_path, workbook_path)
                rolled_back = True
                return {
                    "success": False,
                    "snapshot": snapshot_path,
                    "operation_result": op_result,
                    "verify_result": {"error": str(e)},
                    "rolled_back": True,
                    "message": f"Verification raised an exception; file restored: {e}",
                }

        # Step 4: Success
        if not keep_snapshot and snapshot_path:
            try:
                os.remove(snapshot_path)
                snapshot_path = None
            except Exception:
                pass

        log_audit(
            "SAFE_OPERATION_SUCCESS",
            {
                "workbook": workbook_path,
                "snapshot": snapshot_path,
                "op_result": str(op_result)[:200],
            }
        )

        return {
            "success": True,
            "snapshot": snapshot_path,
            "operation_result": op_result,
            "verify_result": verify_result,
            "rolled_back": False,
            "message": "Operation completed successfully.",
        }

    except Exception as e:
        # Unexpected failure in the harness itself
        if snapshot_path and not rolled_back:
            try:
                restore_snapshot(snapshot_path, workbook_path)
                rolled_back = True
            except Exception:
                pass
        return {
            "success": False,
            "snapshot": snapshot_path,
            "operation_result": None,
            "verify_result": None,
            "rolled_back": rolled_back,
            "message": f"Unexpected harness failure: {e}",
        }
