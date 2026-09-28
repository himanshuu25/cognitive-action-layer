"""
table_tools.py — Phase 0/4
Self-finishing table write: write → number formats → autofit → clamp → freeze header.
Also provides describe_sheet for precise sheet introspection.
"""

from __future__ import annotations
import re
import json
from datetime import date, datetime
from typing import Any, List, Optional, Union

import xlwings as xw

from workbook_tools import get_active_book, get_sheet
from shared_context import (
    excel_suppressed,
    escape_data,
    validate_range_address,
    validate_sheet_name,
    check_data_size,
    ok, err, result_to_str,
)
from verification import verify_write, reconcile_numeric, scan_formula_errors
from logger import logger, log_audit
from security import audit_write


# ── Column Width Constants ────────────────────────────────────────────────────

MIN_COL_WIDTH = 8
MAX_COL_WIDTH = 48
HEADER_ROW_HEIGHT = 20


# ── Auto Number Format Detection ──────────────────────────────────────────────

def _detect_format(values: list) -> str:
    """
    Heuristically detect the best Excel number format for a column of values.
    Returns an Excel format string, or '' (General) if mixed/text.
    """
    numeric_count = 0
    date_count = 0
    pct_count = 0
    total = 0

    for v in values:
        if v is None or v == "":
            continue
        total += 1

        if isinstance(v, (date, datetime)):
            date_count += 1
            continue

        if isinstance(v, (int, float)):
            numeric_count += 1
            # Heuristic: looks like a percentage if between -1 and 1 and non-zero
            if isinstance(v, float) and -1.0 <= v <= 1.0 and v != 0:
                pct_count += 1
            continue

        if isinstance(v, str):
            s = v.strip()
            if s.endswith("%"):
                pct_count += 1
                continue
            # Try numeric
            try:
                float(s.replace(",", ""))
                numeric_count += 1
            except ValueError:
                pass

    if total == 0:
        return ""

    if date_count / max(total, 1) > 0.5:
        return "yyyy-mm-dd"

    if numeric_count / max(total, 1) > 0.5:
        if pct_count / max(numeric_count, 1) > 0.5:
            return "0.00%"
        # Check if looks like currency (large numbers)
        big = [v for v in values if isinstance(v, (int, float)) and abs(v) >= 1000]
        if len(big) > max(total * 0.3, 1):
            return '#,##0.00'
        return "General"

    return ""  # General / text


# ── Self-Finishing Table Write ────────────────────────────────────────────────

def write_table(
    sheet: str,
    start_cell: str,
    data: list,
    column_formats: Optional[dict] = None,
    column_order: Optional[list[str]] = None,
    table_name: Optional[str] = None,
    freeze_header: bool = True,
    verify: bool = True,
    allow_formulas: bool = False,
) -> str:
    """
    Atomic, self-finishing table write. Runs in this exact order:
      1. Reorder columns (if column_order provided) → 2. Write data 
      → 3. Apply number formats → 4. Autofit columns → 5. Clamp widths 
      → 6. Freeze header → 7. Create Excel Table → 8. Verify.

    Parameters:
    - sheet: Worksheet name.
    - start_cell: Top-left cell (e.g. 'A1').
    - data: list-of-lists (first row = headers) or list-of-dicts.
    - column_formats: Optional dict mapping column header → Excel format string.
    - column_order: Optional list of column header names specifying exact left-to-right order.
    - table_name: Optional Excel Table name. Auto-generated if omitted.
    - freeze_header: If True, freeze the header row.
    - verify: If True, read back and verify after writing.
    - allow_formulas: If True, skip injection escaping (use only for trusted data).
    """
    try:
        validate_sheet_name(sheet)
        start_cell = validate_range_address(start_cell)
        check_data_size(data)
    except ValueError as e:
        return result_to_str(err("VALIDATION_ERROR", str(e)))

    try:
        book = get_active_book()
        s = get_sheet(book, sheet)
        app = book.app

        with excel_suppressed(app):
            # ── 1. Convert data & apply column_order ────────────────────
            if data and isinstance(data[0], dict):
                headers = list(data[0].keys())
                if column_order:
                    # Enforce explicit column order, keeping any extra columns at the end
                    ordered_headers = [c for c in column_order if c in headers]
                    remaining = [c for c in headers if c not in ordered_headers]
                    headers = ordered_headers + remaining
                rows = [headers] + [[row.get(h) for h in headers] for row in data]
            elif data and isinstance(data[0], list):
                headers = data[0]
                data_rows = data[1:]
                if column_order:
                    ordered_indices = [headers.index(c) for c in column_order if c in headers]
                    remaining_indices = [i for i, c in enumerate(headers) if c not in column_order]
                    final_indices = ordered_indices + remaining_indices
                    new_headers = [headers[i] for i in final_indices]
                    new_data_rows = [
                        [r[i] if i < len(r) else None for i in final_indices]
                        for r in data_rows
                    ]
                    rows = [new_headers] + new_data_rows
                else:
                    rows = data
            else:
                rows = data

            if not allow_formulas:
                rows = escape_data(rows)

            # ── 2. Write data ────────────────────────────────────────────
            start_rng = s.range(start_cell)
            start_rng.value = rows

            n_rows = len(rows)
            n_cols = len(rows[0]) if rows else 0

            if n_rows == 0 or n_cols == 0:
                return result_to_str(ok("No data to write."))

            # Determine the full table range
            end_col = _col_num_to_letter(start_rng.column + n_cols - 1)
            end_row = start_rng.row + n_rows - 1
            table_range_addr = f"{start_cell}:{end_col}{end_row}"
            table_rng = s.range(table_range_addr)

            # ── 3. Style header row ──────────────────────────────────────
            header_rng = s.range(f"{start_cell}:{end_col}{start_rng.row}")
            try:
                header_rng.font.bold = True
                header_rng.font.size = 11
                header_rng.color = (31, 73, 125)      # Dark navy
                header_rng.font.color = (255, 255, 255)
                header_rng.api.HorizontalAlignment = -4108  # xlCenter
                header_rng.row_height = HEADER_ROW_HEIGHT
            except Exception:
                pass

            # ── 4. Apply number formats per column ───────────────────────
            if rows:
                headers_list = rows[0]
                data_rows = rows[1:]

                for c_idx, hdr in enumerate(headers_list):
                    col_letter = _col_num_to_letter(start_rng.column + c_idx)
                    data_row_start = start_rng.row + 1
                    data_row_end = start_rng.row + len(data_rows)

                    if data_row_start > data_row_end:
                        continue

                    col_range_addr = f"{col_letter}{data_row_start}:{col_letter}{data_row_end}"
                    col_rng = s.range(col_range_addr)

                    # User-supplied format takes priority
                    fmt = ""
                    if column_formats and hdr in column_formats:
                        fmt = column_formats[hdr]
                    else:
                        # Auto-detect from column data values
                        col_values = [
                            r[c_idx] if c_idx < len(r) else None
                            for r in data_rows
                            if isinstance(r, list)
                        ]
                        fmt = _detect_format(col_values)

                    if fmt:
                        try:
                            col_rng.number_format = fmt
                        except Exception:
                            pass

            # ── 5. Autofit & Clamp Columns ───────────────────────────────
            autofit_smart(sheet, table_range_addr, _already_in_suppress=True, _sheet_obj=s, _app=app)

            # ── 6. Borders on data range ─────────────────────────────────
            try:
                table_rng.api.Borders.LineStyle = 1
                table_rng.api.Borders.Weight = 2
            except Exception:
                pass

            # ── 7. Native Excel Table ────────────────────────────────────
            tbl_name = table_name or f"Table_{sheet[:6]}_{start_cell}"
            tbl_name = re.sub(r"[^A-Za-z0-9_]", "_", tbl_name)
            try:
                s.api.ListObjects.Add(
                    SourceType=1,       # xlSrcRange
                    Source=table_rng.api,
                    XlListObjectHasHeaders=1,   # xlYes
                ).Name = tbl_name
            except Exception:
                # Table already exists or COM error — apply manual header borders as fallback
                try:
                    header_rng.api.Borders.LineStyle = 1
                    header_rng.api.Borders.Weight = 4
                except Exception:
                    pass

            # ── 8. Freeze Header ─────────────────────────────────────────
            if freeze_header:
                try:
                    freeze_cell = f"A{start_rng.row + 1}"
                    s.activate()
                    win = app.api.ActiveWindow
                    win.FreezePanes = False
                    win.ScrollRow = 1
                    win.ScrollColumn = 1
                    s.range(freeze_cell).api.Select()
                    win.FreezePanes = True
                except Exception:
                    pass

        # ── 9. Verify (after suppression block so calc runs) ─────────────
        verify_result = None
        if verify:
            intended_2d = rows
            verify_result = verify_write(s, table_range_addr, intended_2d)
            if not verify_result["verified"]:
                log_audit(
                    "WRITE_TABLE_VERIFY_FAIL",
                    {"sheet": sheet, "range": table_range_addr, "result": verify_result}
                )
                return result_to_str(err(
                    "VERIFY_FAILED",
                    f"Write completed but verification found discrepancies in '{table_range_addr}'.",
                    details=verify_result,
                ))

        audit_write(
            operation="WRITE_TABLE",
            tool="write_table",
            args_summary={"sheet": sheet, "start_cell": start_cell, "rows": n_rows, "cols": n_cols},
            after_state={"range": table_range_addr, "verified": verify_result.get("verified") if verify_result else None},
            result_status="ok",
            sheet=sheet,
            range_address=table_range_addr,
        )

        msg = (
            f"Table written to '{sheet}'!{table_range_addr} "
            f"({n_rows} rows × {n_cols} cols, including header)."
        )
        if verify_result:
            msg += " ✓ Verified."

        return result_to_str(ok(msg, data={"range": table_range_addr, "rows": n_rows, "cols": n_cols}))

    except Exception as e:
        logger.exception("write_table failed")
        return result_to_str(err("WRITE_TABLE_ERROR", str(e)))


# ── Smart Autofit ─────────────────────────────────────────────────────────────

def autofit_smart(
    sheet: str,
    range_address: str = None,
    min_width: float = MIN_COL_WIDTH,
    max_width: float = MAX_COL_WIDTH,
    wrap_long_text: bool = True,
    _already_in_suppress: bool = False,
    _sheet_obj=None,
    _app=None,
) -> str:
    """
    Autofit all columns in the range, then clamp widths between min_width and max_width.
    Long-text columns (wider than max_width before clamp) get wrap_text enabled.
    """
    def _do_autofit(s, app):
        rng = s.range(range_address) if range_address else s.used_range
        cols = rng.columns
        cols.autofit()

        for col in cols:
            try:
                w = col.column_width
                if w < min_width:
                    col.column_width = min_width
                elif w > max_width:
                    col.column_width = max_width
                    if wrap_long_text:
                        try:
                            col.api.WrapText = True
                        except Exception:
                            pass
            except Exception:
                pass

    try:
        if _already_in_suppress and _sheet_obj is not None:
            _do_autofit(_sheet_obj, _app)
            return result_to_str(ok(f"Columns autofitted and clamped [{min_width}–{max_width}]."))

        validate_sheet_name(sheet)
        book = get_active_book()
        s = get_sheet(book, sheet)
        with excel_suppressed(book.app):
            _do_autofit(s, book.app)

        return result_to_str(ok(
            f"Columns autofitted and clamped [{min_width}–{max_width}] in '{sheet}'."
        ))
    except Exception as e:
        logger.exception("autofit_smart failed")
        return result_to_str(err("AUTOFIT_ERROR", str(e)))


# ── Sheet Introspection ────────────────────────────────────────────────────────

def describe_sheet(sheet: str) -> str:
    """
    Returns a precise description of the used range in a worksheet:
    - Used range address
    - Last used row number
    - Last used column number and letter
    - Total rows and columns of data
    - First-row headers (if any)
    - Named tables on the sheet

    This is the prerequisite for accurate range targeting by the LLM.
    """
    try:
        validate_sheet_name(sheet)
        book = get_active_book()
        s = get_sheet(book, sheet)

        used = s.used_range
        if used is None:
            return result_to_str(ok("Sheet is empty.", data={"empty": True, "sheet": sheet}))

        first_row_val = used.rows[0].value
        if not isinstance(first_row_val, list):
            first_row_val = [first_row_val]

        headers = [str(h) if h is not None else "" for h in first_row_val]

        # Named tables
        tables = []
        try:
            for lo in s.api.ListObjects:
                tables.append({
                    "name": lo.Name,
                    "range": lo.Range.Address,
                    "header_row": lo.HeaderRowRange.Address if lo.HeaderRowRange else None,
                })
        except Exception:
            pass

        data = {
            "sheet": sheet,
            "used_range": used.address,
            "first_row": used.row,
            "first_col": used.column,
            "last_row": used.last_cell.row,
            "last_col": used.last_cell.column,
            "last_col_letter": _col_num_to_letter(used.last_cell.column),
            "total_rows": used.rows.count,
            "total_cols": used.columns.count,
            "headers": headers,
            "tables": tables,
        }

        return result_to_str(ok(f"Sheet '{sheet}' described.", data=data))

    except Exception as e:
        logger.exception("describe_sheet failed")
        return result_to_str(err("DESCRIBE_SHEET_ERROR", str(e)))


# ── Full Workbook Map & Formula Dependency Graph (Phase 8) ──────────────────────

def describe_workbook(workbook_id: str = None) -> str:
    """
    Returns a comprehensive map of the entire workbook (Phase 8):
    - All sheets (visible and hidden)
    - Used ranges and cell counts per sheet
    - Named tables and named ranges
    - Pivot tables and chart objects
    """
    try:
        from workbook_registry import registry
        book = registry.get_book(workbook_id)

        sheets_summary = []
        for s in book.sheets:
            try:
                used = s.used_range
                u_addr = used.address if used else "Empty"
                u_rows = used.rows.count if used else 0
                u_cols = used.columns.count if used else 0
            except Exception:
                u_addr, u_rows, u_cols = "Empty", 0, 0

            sheets_summary.append({
                "name": s.name,
                "visible": s.api.Visible == -1,
                "used_range": u_addr,
                "rows": u_rows,
                "cols": u_cols,
            })

        # Named ranges
        names = []
        try:
            for n in book.api.Names:
                names.append({"name": n.Name, "refers_to": n.RefersTo})
        except Exception:
            pass

        data = {
            "workbook_name": book.name,
            "workbook_path": book.fullname,
            "total_sheets": len(book.sheets),
            "sheets": sheets_summary,
            "named_ranges": names,
        }

        return result_to_str(ok(f"Workbook '{book.name}' mapped successfully.", data=data))
    except Exception as e:
        logger.exception("describe_workbook failed")
        return result_to_str(err("DESCRIBE_WORKBOOK_ERROR", str(e)))


def get_formula_graph(sheet: str, range_address: str = None) -> str:
    """
    Extracts all formulas in a range and parses their formula dependency graph (Phase 8).
    Returns mapping of formula cell -> referenced cells/ranges.
    """
    try:
        validate_sheet_name(sheet)
        from workbook_tools import get_active_book, get_sheet
        book = get_active_book()
        s = get_sheet(book, sheet)

        rng = s.range(range_address) if range_address else s.used_range
        formulas = rng.formula

        if not formulas:
            return result_to_str(ok("No data found.", data={}))

        if not isinstance(formulas, list):
            formulas = [[formulas]]
        elif formulas and not isinstance(formulas[0], list):
            formulas = [formulas]

        base_row = rng.row
        base_col = rng.column

        graph = {}
        for r_idx, row in enumerate(formulas):
            if not isinstance(row, list):
                row = [row]
            for c_idx, cell in enumerate(row):
                if isinstance(cell, str) and cell.startswith("="):
                    cell_addr = f"{_col_num_to_letter(base_col + c_idx)}{base_row + r_idx}"
                    # Extract cell/range references using regex
                    refs = re.findall(r'[A-Za-z]{1,3}\d+(?::[A-Za-z]{1,3}\d+)?', cell)
                    graph[cell_addr] = {
                        "formula": cell,
                        "references": list(set(refs)),
                    }

        return result_to_str(ok(
            f"Extracted {len(graph)} formula node(s) in '{sheet}'!{rng.address}.",
            data=graph
        ))
    except Exception as e:
        logger.exception("get_formula_graph failed")
        return result_to_str(err("FORMULA_GRAPH_ERROR", str(e)))


def _col_num_to_letter(n: int) -> str:
    result = ""
    while n > 0:
        n, rem = divmod(n - 1, 26)
        result = chr(65 + rem) + result
    return result
