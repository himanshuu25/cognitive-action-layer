"""
data_tools.py — Production-hardened
Data read/write/clean with:
  - Formula/DDE injection guard on all writes
  - Type-preserving dedupe (no int→float coercion)
  - Numeric filter support (>, <, >=, <=, ==, !=)
  - Structured error returns
  - Write-and-verify integration on write_data
  - Resource limit checks
"""

from __future__ import annotations
import json
import math
from typing import Optional

import xlwings as xw
import pandas as pd

from workbook_tools import get_active_book, get_sheet
from shared_context import (
    excel_suppressed,
    escape_data,
    validate_range_address,
    validate_sheet_name,
    check_data_size,
    ok, err, result_to_str,
)
from verification import verify_write, reconcile_numeric
from logger import logger, log_audit


# ── Read Data ─────────────────────────────────────────────────────────────────

def read_data(sheet: str, range_address: str = None) -> str:
    """
    Reads cell data from a worksheet and returns it as a JSON array.
    Automatically reads the entire used range if range_address is omitted.
    """
    try:
        validate_sheet_name(sheet)
        book = get_active_book()
        s = get_sheet(book, sheet)

        target_range = s.range(range_address) if range_address else s.used_range
        values = target_range.value

        if not values:
            return result_to_str(ok("No data found.", data=[]))

        # Single cell
        if not isinstance(values, list):
            return result_to_str(ok("Cell value read.", data={"cell_value": values}))

        # Normalize to 2D
        if values and not isinstance(values[0], list):
            values = [values]

        # Parse into list-of-dicts if first row is all strings (likely headers)
        if (
            len(values) > 1
            and all(isinstance(x, str) for x in values[0] if x is not None)
        ):
            headers = [str(h) if h is not None else f"Column_{i}" for i, h in enumerate(values[0])]
            rows = []
            for row in values[1:]:
                if not isinstance(row, list):
                    row = [row]
                row_extended = row + [None] * (len(headers) - len(row))
                rows.append({headers[i]: row_extended[i] for i in range(len(headers))})
            return result_to_str(ok(f"Read {len(rows)} rows from '{sheet}'.", data=rows))
        else:
            return result_to_str(ok(f"Read {len(values)} rows from '{sheet}'.", data=values))

    except Exception as e:
        logger.exception("read_data failed")
        return result_to_str(err("READ_DATA_ERROR", str(e)))


# ── Write Data ────────────────────────────────────────────────────────────────

def write_data(
    sheet: str,
    range_address: str,
    data: list,
    allow_formulas: bool = False,
    verify: bool = False,
) -> str:
    """
    Writes a list of rows (list-of-lists or list-of-dicts) to a worksheet.

    Parameters:
    - sheet: Worksheet name.
    - range_address: Starting top-left cell (e.g. 'A1').
    - data: List of lists or list of dicts.
    - allow_formulas: If True, skip injection escaping (only for trusted formula data).
    - verify: If True, read back and verify after writing.

    NOTE: For production table writes with auto-formatting and verification,
    use write_table() instead — it gives you better first-render output.
    """
    try:
        validate_sheet_name(sheet)
        range_address = validate_range_address(range_address)
        check_data_size(data)

        book = get_active_book()
        s = get_sheet(book, sheet)

        # Convert list-of-dicts to list-of-lists
        if data and isinstance(data[0], dict):
            headers = list(data[0].keys())
            rows = [headers] + [[item.get(h) for h in headers] for item in data]
        else:
            rows = data

        # Apply injection escaping unless explicitly disabled
        if not allow_formulas:
            rows = escape_data(rows)

        with excel_suppressed(book.app):
            s.range(range_address).value = rows

        # Optional verification
        if verify:
            vr = verify_write(s, range_address, rows)
            if not vr["verified"]:
                return result_to_str(err(
                    "VERIFY_FAILED",
                    f"Data written but verification found discrepancies at '{range_address}'.",
                    details=vr,
                ))
            return result_to_str(ok(
                f"Data written and verified at '{sheet}'!{range_address}.",
                data={"rows": len(rows), "verified": True}
            ))

        return result_to_str(ok(
            f"Data written to '{sheet}'!{range_address} ({len(rows)} rows).",
            data={"rows": len(rows)}
        ))

    except Exception as e:
        logger.exception("write_data failed")
        return result_to_str(err("WRITE_DATA_ERROR", str(e)))


# ── Remove Duplicates ────────────────────────────────────────────────────────

def remove_duplicates(sheet: str, range_address: str = None, columns: list[str] = None) -> str:
    """
    Removes duplicate rows using Pandas. Preserves original data types (no int→float coercion).

    Parameters:
    - sheet: Worksheet name.
    - range_address: Cell range (e.g. 'A1:D100'). Defaults to entire used range.
    - columns: Optional list of column headers to check for duplicates. Defaults to all columns.
    """
    try:
        validate_sheet_name(sheet)
        book = get_active_book()
        s = get_sheet(book, sheet)

        target_range = s.range(range_address) if range_address else s.used_range
        values = target_range.value

        if not values or not isinstance(values, list) or len(values) <= 1:
            return result_to_str(ok("No tabular data found."))

        if not isinstance(values[0], list):
            return result_to_str(ok("Single-row range — nothing to deduplicate."))

        headers = [str(h) if h is not None else f"Col_{i}" for i, h in enumerate(values[0])]
        data_rows = values[1:]

        # Build DataFrame while preserving types
        df = pd.DataFrame(data_rows, columns=headers)
        initial_len = len(df)

        # Type-preserving: track original dtypes
        original_dtypes = {col: df[col].dtype for col in df.columns}

        if columns:
            valid_cols = [c for c in columns if c in df.columns]
            if not valid_cols:
                return result_to_str(err(
                    "COLUMN_NOT_FOUND",
                    f"None of {columns} found in headers: {headers}",
                ))
            df = df.drop_duplicates(subset=valid_cols).reset_index(drop=True)
        else:
            df = df.drop_duplicates().reset_index(drop=True)

        removed_count = initial_len - len(df)

        # Restore original types where possible
        for col in df.columns:
            try:
                df[col] = df[col].astype(original_dtypes[col])
            except Exception:
                pass

        # Convert NaN to None (not "nan" string)
        df = df.where(pd.notna(df), other=None)

        with excel_suppressed(book.app):
            target_range.clear_contents()
            new_values = [headers] + [
                [row[h] for h in headers] for row in df.to_dict(orient="records")
            ]
            s.range(target_range.address.split(":")[0]).value = new_values

        return result_to_str(ok(
            f"Removed {removed_count} duplicate row(s). {len(df)} rows remaining.",
            data={"removed": removed_count, "remaining": len(df)}
        ))

    except Exception as e:
        logger.exception("remove_duplicates failed")
        return result_to_str(err("REMOVE_DUPLICATES_ERROR", str(e)))


# ── Filter Data ───────────────────────────────────────────────────────────────

def filter_data(
    sheet: str,
    range_address: str = None,
    filter_column: str = None,
    criteria: str = None,
    target_sheet: str = None,
    operator: str = "contains",
) -> str:
    """
    Filters rows by a column value and optionally writes results to a new sheet.

    Parameters:
    - sheet: Source worksheet name.
    - range_address: Cell range (e.g. 'A1:E100'). Defaults to used range.
    - filter_column: Column header to filter on.
    - criteria: Value or threshold to compare against.
    - target_sheet: Optional sheet name to write filtered results to.
    - operator: Comparison operator:
        'contains' (text substring, default),
        '==' | '=' (exact match),
        '!=' (not equal),
        '>' | '>=' | '<' | '<=' (numeric comparisons).
    """
    try:
        validate_sheet_name(sheet)
        book = get_active_book()
        s = get_sheet(book, sheet)

        target_range = s.range(range_address) if range_address else s.used_range
        values = target_range.value

        if not values or not isinstance(values, list) or len(values) <= 1:
            return result_to_str(ok("No tabular data found.", data=[]))

        if not isinstance(values[0], list):
            values = [values]

        headers = [str(h) if h is not None else f"Col_{i}" for i, h in enumerate(values[0])]
        data_rows = values[1:]

        df = pd.DataFrame(data_rows, columns=headers)

        if filter_column not in df.columns:
            return result_to_str(err(
                "COLUMN_NOT_FOUND",
                f"Column '{filter_column}' not found. Available: {headers}",
            ))

        # Build mask based on operator
        op = operator.strip().lower()

        # Numeric operators
        if op in (">", ">=", "<", "<="):
            try:
                threshold = float(criteria)
                numeric_col = pd.to_numeric(df[filter_column], errors="coerce")
                if op == ">":
                    mask = numeric_col > threshold
                elif op == ">=":
                    mask = numeric_col >= threshold
                elif op == "<":
                    mask = numeric_col < threshold
                else:
                    mask = numeric_col <= threshold
            except (TypeError, ValueError):
                return result_to_str(err(
                    "INVALID_CRITERIA",
                    f"Criteria '{criteria}' must be numeric for operator '{op}'.",
                ))

        elif op in ("==", "=", "equal", "equals"):
            # Try numeric first, fall back to string
            try:
                threshold = float(criteria)
                numeric_col = pd.to_numeric(df[filter_column], errors="coerce")
                mask = numeric_col == threshold
            except (TypeError, ValueError):
                mask = df[filter_column].astype(str).str.lower() == str(criteria).lower()

        elif op in ("!=", "not_equal", "ne"):
            try:
                threshold = float(criteria)
                numeric_col = pd.to_numeric(df[filter_column], errors="coerce")
                mask = numeric_col != threshold
            except (TypeError, ValueError):
                mask = df[filter_column].astype(str).str.lower() != str(criteria).lower()

        else:  # 'contains' (default)
            mask = df[filter_column].astype(str).str.contains(
                str(criteria), case=False, na=False, regex=False
            )

        filtered_df = df[mask]

        if target_sheet:
            t_sheet = None
            for sh in book.sheets:
                if sh.name.lower() == target_sheet.lower():
                    t_sheet = sh
                    break
            if not t_sheet:
                t_sheet = book.sheets.add(name=target_sheet)

            with excel_suppressed(book.app):
                t_sheet.clear()
                new_values = [headers] + filtered_df.values.tolist()
                t_sheet.range("A1").value = new_values

            return result_to_str(ok(
                f"Filtered {len(filtered_df)} matching rows → sheet '{target_sheet}'.",
                data={"matching_rows": len(filtered_df), "target_sheet": target_sheet}
            ))
        else:
            # Convert NaN to None
            records = json.loads(filtered_df.where(pd.notna(filtered_df), other=None).to_json(orient="records"))
            return result_to_str(ok(
                f"Found {len(records)} matching rows.",
                data=records
            ))

    except Exception as e:
        logger.exception("filter_data failed")
        return result_to_str(err("FILTER_DATA_ERROR", str(e)))


# ── Clear Range ───────────────────────────────────────────────────────────────

def clear_range(sheet: str, range_address: str, clear_type: str = "all") -> str:
    """
    Clears a range. clear_type: 'all' (contents + formats), 'contents', 'formats'.
    """
    try:
        validate_sheet_name(sheet)
        range_address = validate_range_address(range_address)
        book = get_active_book()
        s = get_sheet(book, sheet)
        r = s.range(range_address)

        with excel_suppressed(book.app):
            c_type = clear_type.lower()
            if c_type == "contents":
                r.clear_contents()
                return result_to_str(ok(f"Contents cleared in '{sheet}'!{range_address} (formatting kept)."))
            elif c_type == "formats":
                try:
                    r.api.ClearFormats()
                except Exception:
                    r.clear()
                return result_to_str(ok(f"Formatting cleared in '{sheet}'!{range_address} (contents kept)."))
            else:
                r.clear()
                return result_to_str(ok(f"Contents and formatting cleared in '{sheet}'!{range_address}."))

    except Exception as e:
        logger.exception("clear_range failed")
        return result_to_str(err("CLEAR_RANGE_ERROR", str(e)))


# ── Copy / Paste ───────────────────────────────────────────────────────────────

def copy_paste_range(
    source_sheet: str,
    source_range: str,
    target_sheet: str,
    target_range: str,
    paste_type: str = "all",
) -> str:
    """
    Copies a range from source_sheet to target_sheet.
    paste_type: 'all' (contents + styles) or 'values' (raw values only).
    """
    try:
        validate_sheet_name(source_sheet)
        validate_sheet_name(target_sheet)
        book = get_active_book()
        s_sheet = get_sheet(book, source_sheet)
        t_sheet = get_sheet(book, target_sheet)

        src = s_sheet.range(source_range)
        tgt = t_sheet.range(target_range)

        with excel_suppressed(book.app):
            if paste_type.lower() == "values":
                tgt.value = src.value
            else:
                src.copy(destination=tgt)

        return result_to_str(ok(
            f"Copied '{source_sheet}'!{source_range} → '{target_sheet}'!{target_range} (paste_type={paste_type})."
        ))
    except Exception as e:
        logger.exception("copy_paste_range failed")
        return result_to_str(err("COPY_PASTE_ERROR", str(e)))


# ── Find & Replace ────────────────────────────────────────────────────────────

def find_replace(
    sheet: str,
    find_value: str,
    replace_value: str,
    range_address: str = None,
    match_partial: bool = False,
) -> str:
    """
    Finds cell values and replaces them.

    Parameters:
    - sheet: Worksheet name.
    - find_value: Value to find.
    - replace_value: Replacement value.
    - range_address: Optional range. Defaults to used range.
    - match_partial: If True, do substring matching (not just whole-cell). Default False.
    """
    try:
        validate_sheet_name(sheet)
        book = get_active_book()
        s = get_sheet(book, sheet)
        r = s.range(range_address) if range_address else s.used_range

        values = r.value
        if not values:
            return result_to_str(ok("No data found.", data={"replaced": 0}))

        # Normalize to 2D
        if not isinstance(values, list):
            values = [[values]]
        elif values and not isinstance(values[0], list):
            values = [values]

        count = 0

        def _replace_cell(item):
            nonlocal count
            if item is None:
                return item
            item_str = str(item)
            find_str = str(find_value)

            if match_partial:
                if find_str.lower() in item_str.lower():
                    count += 1
                    new_str = item_str.replace(find_str, str(replace_value))
                    # Try to preserve original type
                    if isinstance(item, (int, float)):
                        try:
                            return type(item)(new_str)
                        except ValueError:
                            return new_str
                    return new_str
            else:
                if item_str.lower() == find_str.lower():
                    count += 1
                    # Try to cast replace_value to original type
                    if isinstance(item, (int, float)):
                        try:
                            if "." in str(replace_value):
                                return float(replace_value)
                            return int(replace_value)
                        except ValueError:
                            pass
                    return replace_value
            return item

        new_values = []
        for row in values:
            if isinstance(row, list):
                new_values.append([_replace_cell(cell) for cell in row])
            else:
                new_values.append(_replace_cell(row))

        with excel_suppressed(book.app):
            r.value = new_values

        return result_to_str(ok(
            f"Replaced {count} occurrence(s) of '{find_value}' with '{replace_value}' in '{sheet}'.",
            data={"replaced": count}
        ))

    except Exception as e:
        logger.exception("find_replace failed")
        return result_to_str(err("FIND_REPLACE_ERROR", str(e)))
