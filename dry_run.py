"""
dry_run.py — Phase 5
Dry-run diff preview and human confirmation gates for irreversible operations.
"""

from __future__ import annotations
import json
from typing import Any, List, Optional

from logger import logger


# ── Diff Generation ───────────────────────────────────────────────────────────

def _cell_addr(base_row: int, base_col: int, r: int, c: int) -> str:
    """Build a cell address like 'B5' from 1-based base row/col and offsets."""
    col_num = base_col + c
    col_letter = _col_num_to_letter(col_num)
    return f"{col_letter}{base_row + r}"


def _col_num_to_letter(n: int) -> str:
    result = ""
    while n > 0:
        n, rem = divmod(n - 1, 26)
        result = chr(65 + rem) + result
    return result


def preview_write(
    sheet,
    range_address: str,
    intended_data: list,
) -> dict:
    """
    Generate a before/after diff preview without writing anything.

    Returns:
        {
            "changes": [
                {
                    "address": "B3",
                    "before": <current value>,
                    "after": <intended value>,
                    "type": "modify" | "clear" | "new"
                },
                ...
            ],
            "total_cells": int,
            "cells_changing": int,
            "cells_unchanged": int,
        }
    """
    try:
        rng = sheet.range(range_address)
        base_row = rng.row
        base_col = rng.column

        # Read current values
        current = rng.value
        if current is None:
            current = []
        elif not isinstance(current, list):
            current = [[current]]
        elif current and not isinstance(current[0], list):
            current = [current]

        # Normalize intended to 2D list
        if isinstance(intended_data, list) and intended_data:
            if isinstance(intended_data[0], dict):
                headers = list(intended_data[0].keys())
                intended_2d = [headers] + [[row.get(h) for h in headers] for row in intended_data]
            elif isinstance(intended_data[0], list):
                intended_2d = intended_data
            else:
                intended_2d = [intended_data]
        else:
            intended_2d = intended_data or []

        changes = []
        total = 0
        changing = 0
        unchanged = 0

        max_rows = max(len(current), len(intended_2d))

        for r_idx in range(max_rows):
            cur_row = current[r_idx] if r_idx < len(current) else []
            int_row = intended_2d[r_idx] if r_idx < len(intended_2d) else []

            if not isinstance(cur_row, list):
                cur_row = [cur_row]

            max_cols = max(len(cur_row), len(int_row))
            for c_idx in range(max_cols):
                total += 1
                before = cur_row[c_idx] if c_idx < len(cur_row) else None
                after = int_row[c_idx] if c_idx < len(int_row) else None
                addr = _cell_addr(base_row, base_col, r_idx, c_idx)

                if before == after or (before is None and after is None):
                    unchanged += 1
                    continue

                if before is None or before == "":
                    change_type = "new"
                elif after is None or after == "":
                    change_type = "clear"
                else:
                    change_type = "modify"

                changes.append({
                    "address": addr,
                    "before": str(before) if before is not None else None,
                    "after": str(after) if after is not None else None,
                    "type": change_type,
                })
                changing += 1

        return {
            "changes": changes,
            "total_cells": total,
            "cells_changing": changing,
            "cells_unchanged": unchanged,
        }

    except Exception as e:
        return {"error": str(e), "changes": [], "total_cells": 0, "cells_changing": 0}


def format_preview_summary(preview: dict, max_show: int = 20) -> str:
    """
    Format a dry-run preview dict into a human-readable summary string.
    """
    if "error" in preview:
        return f"Preview error: {preview['error']}"

    changes = preview.get("changes", [])
    total = preview.get("total_cells", 0)
    changing = preview.get("cells_changing", 0)
    unchanged = preview.get("cells_unchanged", 0)

    lines = [
        f"📋 DRY-RUN PREVIEW — {changing} of {total} cells will change ({unchanged} unchanged)",
        "",
    ]

    shown = changes[:max_show]
    for ch in shown:
        marker = {"modify": "✏️ ", "new": "➕ ", "clear": "🗑️ "}.get(ch["type"], "   ")
        before = ch.get("before", "—") or "—"
        after = ch.get("after", "—") or "—"
        lines.append(f"  {marker} {ch['address']:8s}  {before!r:30s} → {after!r}")

    if len(changes) > max_show:
        lines.append(f"  ... and {len(changes) - max_show} more changes not shown.")

    if changing == 0:
        lines.append("  (No changes — data matches existing sheet content.)")

    lines.append("")
    lines.append("To apply these changes, call the same tool with confirmed=True.")
    return "\n".join(lines)


# ── Confirmation Gate ─────────────────────────────────────────────────────────

class ConfirmationRequired(Exception):
    """
    Raised when a destructive or irreversible operation is attempted
    without explicit confirmation. The message includes the dry-run preview.
    """
    def __init__(self, preview_summary: str):
        super().__init__(preview_summary)
        self.preview_summary = preview_summary


def require_confirmation(
    sheet,
    range_address: str,
    intended_data: list,
    confirmed: bool = False,
    threshold_cells: int = 1,
) -> None:
    """
    Check if a write requires confirmation.
    If confirmed=False and the operation will change >= threshold_cells cells,
    raises ConfirmationRequired with the dry-run preview as the message.

    Usage in a tool:
        require_confirmation(sheet, "A1", data, confirmed=confirmed)
        # If this doesn't raise, proceed with the actual write.
    """
    if confirmed:
        logger.info("Confirmation provided for write to '%s'.", range_address)
        return

    preview = preview_write(sheet, range_address, intended_data)

    if preview.get("cells_changing", 0) >= threshold_cells:
        summary = format_preview_summary(preview)
        logger.warning(
            "Confirmation required for write to '%s' (%d cells changing).",
            range_address, preview["cells_changing"]
        )
        raise ConfirmationRequired(summary)
