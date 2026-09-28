"""
editing_tools.py — Phase 1
Structural editing: insert/delete rows & columns, sort, AutoFilter, named ranges, comments.
"""

from __future__ import annotations
import re
from typing import Any, List, Optional, Union

import xlwings as xw

from workbook_tools import get_active_book, get_sheet
from shared_context import (
    excel_suppressed,
    validate_range_address,
    validate_sheet_name,
    ok, err, result_to_str,
)
from logger import logger


# ── Row Operations ────────────────────────────────────────────────────────────

def insert_rows(sheet: str, row_number: int, count: int = 1) -> str:
    """
    Insert blank rows above row_number.

    Parameters:
    - sheet: Worksheet name.
    - row_number: 1-based row number above which rows are inserted.
    - count: Number of rows to insert (default 1).
    """
    try:
        validate_sheet_name(sheet)
        if row_number < 1:
            return result_to_str(err("INVALID_INPUT", "row_number must be >= 1."))
        if count < 1 or count > 10_000:
            return result_to_str(err("INVALID_INPUT", "count must be between 1 and 10,000."))

        book = get_active_book()
        s = get_sheet(book, sheet)
        with excel_suppressed(book.app):
            # Select the range of rows to insert above
            insert_rng = s.range(f"{row_number}:{row_number + count - 1}")
            insert_rng.api.EntireRow.Insert()

        return result_to_str(ok(
            f"Inserted {count} blank row(s) above row {row_number} in '{sheet}'."
        ))
    except Exception as e:
        logger.exception("insert_rows failed")
        return result_to_str(err("INSERT_ROWS_ERROR", str(e)))


def delete_rows(sheet: str, row_number: int, count: int = 1) -> str:
    """
    Delete rows starting at row_number.

    Parameters:
    - sheet: Worksheet name.
    - row_number: 1-based first row to delete.
    - count: Number of consecutive rows to delete (default 1).
    """
    try:
        validate_sheet_name(sheet)
        if row_number < 1:
            return result_to_str(err("INVALID_INPUT", "row_number must be >= 1."))
        if count < 1 or count > 10_000:
            return result_to_str(err("INVALID_INPUT", "count must be between 1 and 10,000."))

        book = get_active_book()
        s = get_sheet(book, sheet)
        with excel_suppressed(book.app):
            delete_rng = s.range(f"{row_number}:{row_number + count - 1}")
            delete_rng.api.EntireRow.Delete()

        return result_to_str(ok(
            f"Deleted {count} row(s) starting at row {row_number} in '{sheet}'."
        ))
    except Exception as e:
        logger.exception("delete_rows failed")
        return result_to_str(err("DELETE_ROWS_ERROR", str(e)))


# ── Column Operations ─────────────────────────────────────────────────────────

def insert_columns(sheet: str, column_letter: str, count: int = 1) -> str:
    """
    Insert blank columns to the left of column_letter.

    Parameters:
    - sheet: Worksheet name.
    - column_letter: Column letter (e.g. 'C') before which columns are inserted.
    - count: Number of columns to insert (default 1).
    """
    try:
        validate_sheet_name(sheet)
        col = column_letter.strip().upper()
        if not re.match(r'^[A-Z]{1,3}$', col):
            return result_to_str(err("INVALID_INPUT", f"Invalid column letter: '{column_letter}'"))
        if count < 1 or count > 1000:
            return result_to_str(err("INVALID_INPUT", "count must be between 1 and 1,000."))

        book = get_active_book()
        s = get_sheet(book, sheet)

        with excel_suppressed(book.app):
            # Build the column range to insert before
            start_col_num = _col_letter_to_num(col)
            end_col_num = start_col_num + count - 1
            end_col = _col_num_to_letter(end_col_num)
            insert_rng = s.range(f"{col}:{end_col}")
            insert_rng.api.EntireColumn.Insert()

        return result_to_str(ok(
            f"Inserted {count} blank column(s) to the left of '{col}' in '{sheet}'."
        ))
    except Exception as e:
        logger.exception("insert_columns failed")
        return result_to_str(err("INSERT_COLS_ERROR", str(e)))


def delete_columns(sheet: str, column_letter: str, count: int = 1) -> str:
    """
    Delete columns starting at column_letter.

    Parameters:
    - sheet: Worksheet name.
    - column_letter: First column letter to delete (e.g. 'C').
    - count: Number of consecutive columns to delete (default 1).
    """
    try:
        validate_sheet_name(sheet)
        col = column_letter.strip().upper()
        if not re.match(r'^[A-Z]{1,3}$', col):
            return result_to_str(err("INVALID_INPUT", f"Invalid column letter: '{column_letter}'"))
        if count < 1 or count > 1000:
            return result_to_str(err("INVALID_INPUT", "count must be between 1 and 1,000."))

        book = get_active_book()
        s = get_sheet(book, sheet)

        with excel_suppressed(book.app):
            start_col_num = _col_letter_to_num(col)
            end_col = _col_num_to_letter(start_col_num + count - 1)
            delete_rng = s.range(f"{col}:{end_col}")
            delete_rng.api.EntireColumn.Delete()

        return result_to_str(ok(
            f"Deleted {count} column(s) starting at '{col}' in '{sheet}'."
        ))
    except Exception as e:
        logger.exception("delete_columns failed")
        return result_to_str(err("DELETE_COLS_ERROR", str(e)))


# ── Sort ───────────────────────────────────────────────────────────────────────

def sort_range(
    sheet: str,
    range_address: str = None,
    sort_column: str = None,
    ascending: bool = True,
    header: bool = True,
) -> str:
    """
    Sort a range by a column header.

    Parameters:
    - sheet: Worksheet name.
    - range_address: Range to sort (e.g. 'A1:D100'). If omitted, sorts used range.
    - sort_column: Column header name to sort by (e.g. 'Revenue').
    - ascending: True for A→Z / smallest→largest, False for descending.
    - header: True if the first row is a header row (default True).
    """
    try:
        validate_sheet_name(sheet)
        book = get_active_book()
        s = get_sheet(book, sheet)

        with excel_suppressed(book.app):
            rng = s.range(range_address) if range_address else s.used_range

            if sort_column:
                # Find the column index matching sort_column header
                headers_row = rng.rows[0].value
                if not isinstance(headers_row, list):
                    headers_row = [headers_row]
                headers_lower = [str(h).lower() if h is not None else "" for h in headers_row]
                col_name_lower = sort_column.lower()

                if col_name_lower not in headers_lower:
                    return result_to_str(err(
                        "COLUMN_NOT_FOUND",
                        f"Column '{sort_column}' not found. Available: {[h for h in headers_row if h]}",
                    ))

                col_offset = headers_lower.index(col_name_lower)
                sort_key_rng = rng.columns[col_offset]
            else:
                sort_key_rng = rng.columns[0]

            order = 1 if ascending else 2  # xlAscending / xlDescending
            hdr_const = 1 if header else 2  # xlYes / xlNo

            rng.api.Sort(
                Key1=sort_key_rng.api,
                Order1=order,
                Header=hdr_const,
                MatchCase=False,
                Orientation=1,   # xlTopToBottom
            )

        direction = "ascending" if ascending else "descending"
        col_name = sort_column or "first column"
        return result_to_str(ok(
            f"Sorted '{sheet}'!{rng.address} by '{col_name}' ({direction})."
        ))
    except Exception as e:
        logger.exception("sort_range failed")
        return result_to_str(err("SORT_ERROR", str(e)))


# ── AutoFilter ────────────────────────────────────────────────────────────────

def apply_autofilter(sheet: str, range_address: str = None) -> str:
    """
    Enable Excel's native AutoFilter drop-down arrows on a range.

    Parameters:
    - sheet: Worksheet name.
    - range_address: Range to filter (e.g. 'A1:F100'). If omitted, uses the entire used range.
    """
    try:
        validate_sheet_name(sheet)
        book = get_active_book()
        s = get_sheet(book, sheet)

        with excel_suppressed(book.app):
            rng = s.range(range_address) if range_address else s.used_range
            # Toggle off first to avoid "already on" COM errors
            try:
                if s.api.AutoFilterMode:
                    s.api.AutoFilterMode = False
            except Exception:
                pass
            rng.api.AutoFilter()

        return result_to_str(ok(
            f"AutoFilter enabled on '{sheet}'!{rng.address}."
        ))
    except Exception as e:
        logger.exception("apply_autofilter failed")
        return result_to_str(err("AUTOFILTER_ERROR", str(e)))


def remove_autofilter(sheet: str) -> str:
    """
    Remove AutoFilter from the worksheet.

    Parameters:
    - sheet: Worksheet name.
    """
    try:
        validate_sheet_name(sheet)
        book = get_active_book()
        s = get_sheet(book, sheet)
        with excel_suppressed(book.app):
            try:
                if s.api.AutoFilterMode:
                    s.api.AutoFilterMode = False
            except Exception:
                pass
        return result_to_str(ok(f"AutoFilter removed from '{sheet}'."))
    except Exception as e:
        return result_to_str(err("REMOVE_AUTOFILTER_ERROR", str(e)))


# ── Named Ranges ──────────────────────────────────────────────────────────────

def create_named_range(name: str, sheet: str, range_address: str) -> str:
    """
    Create a named range in the workbook.

    Parameters:
    - name: Name for the range (e.g. 'SalesData'). Must be a valid Excel name.
    - sheet: Worksheet name.
    - range_address: Range to name (e.g. 'A1:D100').
    """
    try:
        if not re.match(r'^[A-Za-z_][A-Za-z0-9_.]*$', name):
            return result_to_str(err("INVALID_NAME", f"'{name}' is not a valid Excel name."))

        validate_sheet_name(sheet)
        book = get_active_book()
        s = get_sheet(book, sheet)

        with excel_suppressed(book.app):
            full_ref = f"'{sheet}'!{s.range(range_address).address}"
            book.api.Names.Add(Name=name, RefersTo=f"={full_ref}")

        return result_to_str(ok(f"Named range '{name}' created → {full_ref}."))
    except Exception as e:
        logger.exception("create_named_range failed")
        return result_to_str(err("NAMED_RANGE_ERROR", str(e)))


def delete_named_range(name: str) -> str:
    """
    Delete a named range from the workbook.

    Parameters:
    - name: The named range to delete.
    """
    try:
        book = get_active_book()
        with excel_suppressed(book.app):
            book.api.Names(name).Delete()
        return result_to_str(ok(f"Named range '{name}' deleted."))
    except Exception as e:
        return result_to_str(err("DELETE_NAMED_RANGE_ERROR", str(e)))


def list_named_ranges() -> str:
    """List all named ranges in the active workbook."""
    try:
        book = get_active_book()
        names = []
        for n in book.api.Names:
            try:
                names.append({"name": n.Name, "refers_to": n.RefersTo})
            except Exception:
                pass
        return result_to_str(ok(f"Found {len(names)} named ranges.", data=names))
    except Exception as e:
        return result_to_str(err("LIST_NAMED_RANGES_ERROR", str(e)))


# ── Cell Comments ─────────────────────────────────────────────────────────────

def add_comment(sheet: str, cell_address: str, comment_text: str, author: str = "Excel MCP") -> str:
    """
    Add or update a comment on a cell.

    Parameters:
    - sheet: Worksheet name.
    - cell_address: Single cell address (e.g. 'B5').
    - comment_text: The comment text.
    - author: Optional author name shown in the comment.
    """
    try:
        validate_sheet_name(sheet)
        book = get_active_book()
        s = get_sheet(book, sheet)
        cell = s.range(cell_address)

        with excel_suppressed(book.app):
            # Delete existing comment first
            try:
                cell.api.Comment.Delete()
            except Exception:
                pass
            cell.api.AddComment(comment_text)
            try:
                cell.api.Comment.Author = author
            except Exception:
                pass

        return result_to_str(ok(f"Comment added to '{sheet}'!{cell_address}."))
    except Exception as e:
        return result_to_str(err("ADD_COMMENT_ERROR", str(e)))


def delete_comment(sheet: str, cell_address: str) -> str:
    """
    Delete a comment from a cell.

    Parameters:
    - sheet: Worksheet name.
    - cell_address: Single cell address (e.g. 'B5').
    """
    try:
        validate_sheet_name(sheet)
        book = get_active_book()
        s = get_sheet(book, sheet)
        with excel_suppressed(book.app):
            s.range(cell_address).api.Comment.Delete()
        return result_to_str(ok(f"Comment deleted from '{sheet}'!{cell_address}."))
    except Exception as e:
        return result_to_str(err("DELETE_COMMENT_ERROR", str(e)))


# ── Data Cleaning ─────────────────────────────────────────────────────────────

def clean_data(sheet: str, range_address: str = None) -> str:
    """
    Clean common data quality issues in a range:
    - Trim leading/trailing whitespace from text cells
    - Convert numbers-stored-as-text to actual numbers
    - Remove zero-width spaces and non-breaking spaces

    Parameters:
    - sheet: Worksheet name.
    - range_address: Range to clean (e.g. 'A1:D100'). If omitted, cleans used range.
    """
    try:
        import pandas as pd
        validate_sheet_name(sheet)
        book = get_active_book()
        s = get_sheet(book, sheet)

        with excel_suppressed(book.app):
            rng = s.range(range_address) if range_address else s.used_range
            values = rng.value

            if not values:
                return result_to_str(ok("No data found to clean."))

            # Normalize to 2D
            if not isinstance(values, list):
                values = [[values]]
            elif values and not isinstance(values[0], list):
                values = [values]

            cleaned = 0
            new_values = []
            for row in values:
                new_row = []
                for cell in row:
                    new_cell = _clean_cell(cell)
                    if new_cell != cell:
                        cleaned += 1
                    new_row.append(new_cell)
                new_values.append(new_row)

            rng.value = new_values

        return result_to_str(ok(
            f"Cleaned {cleaned} cell(s) in '{sheet}'!{rng.address}.",
            data={"cells_cleaned": cleaned}
        ))
    except Exception as e:
        logger.exception("clean_data failed")
        return result_to_str(err("CLEAN_DATA_ERROR", str(e)))


def _clean_cell(value) -> any:
    """Clean a single cell value."""
    if not isinstance(value, str):
        return value
    # Remove zero-width spaces, non-breaking spaces
    cleaned = (
        value
        .replace("\u200b", "")   # zero-width space
        .replace("\u00a0", " ")  # non-breaking space → regular space
        .replace("\ufeff", "")   # BOM
        .strip()
    )
    # Try converting numbers-stored-as-text
    if cleaned != value.strip():
        try:
            if "." in cleaned:
                return float(cleaned.replace(",", ""))
            else:
                return int(cleaned.replace(",", ""))
        except ValueError:
            pass
    # Check original value after strip for numeric conversion
    try:
        stripped = value.strip()
        if stripped and stripped != value:
            if "." in stripped:
                return float(stripped.replace(",", ""))
            return int(stripped.replace(",", ""))
    except ValueError:
        pass
    return cleaned


# ── Helpers ────────────────────────────────────────────────────────────────────

def _col_num_to_letter(n: int) -> str:
    result = ""
    while n > 0:
        n, rem = divmod(n - 1, 26)
        result = chr(65 + rem) + result
    return result


def _col_letter_to_num(letters: str) -> int:
    result = 0
    for ch in letters.upper():
        result = result * 26 + (ord(ch) - ord('A') + 1)
    return result
