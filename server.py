"""
server.py — Production-Ready Excel Automation MCP Server
All 5 phases implemented:
  Phase 0: write_table, autofit_smart, describe_sheet, injection guard, workbook-by-name
  Phase 1: insert/delete rows & columns, sort, AutoFilter, named ranges, comments, clean_data
  Phase 2: rich conditional formatting, data validation, chart orientation, filter operators
  Phase 3: structured logging, structured errors throughout
  Phase 4: write-and-verify, reconciliation, formula error scan, snapshot/rollback
  Phase 5: path sandboxing, audit trail, dry-run preview, confirmation gates, resource limits
"""

from mcp.server.fastmcp import FastMCP

import workbook_tools
import data_tools
import format_tools
import chart_tools
import table_tools
import editing_tools
import verification
import snapshot as snapshot_tools
import security
import dry_run

from shared_context import result_to_str, ok, err
from logger import logger

# Initialize FastMCP Server
mcp = FastMCP("Excel-Automation-Server")


# ══════════════════════════════════════════════════════════════════════════════
# WORKBOOK & SHEET MANAGEMENT
# ══════════════════════════════════════════════════════════════════════════════

@mcp.tool()
def open_workbook(file_path: str) -> str:
    """
    Opens an existing Excel workbook on disk. Validates path against allowed directories
    and warns if the file is macro-enabled (.xlsm).

    Parameters:
    - file_path: Absolute or relative path to the Excel file (.xlsx, .xlsm, .csv).
    """
    return workbook_tools.open_workbook(file_path)


@mcp.tool()
def create_workbook(file_path: str = None) -> str:
    """
    Creates a new blank active Excel workbook.

    Parameters:
    - file_path: Optional path to save the workbook immediately.
    """
    return workbook_tools.create_workbook(file_path)


@mcp.tool()
def save_workbook(file_path: str = None) -> str:
    """
    Saves the active workbook.

    Parameters:
    - file_path: Optional new path for 'Save As'. If omitted, saves in place.
    """
    return workbook_tools.save_workbook(file_path)


@mcp.tool()
def list_sheets() -> list[str]:
    """Lists the names of all worksheets in the active workbook."""
    return workbook_tools.list_sheets()


@mcp.tool()
def list_open_workbooks() -> str:
    """
    Lists all currently open Excel workbooks so you can identify the correct target.
    Always call this first if multiple workbooks might be open.
    """
    return workbook_tools.list_open_workbooks()


@mcp.tool()
def add_sheet(name: str) -> str:
    """
    Adds a new worksheet to the active workbook.

    Parameters:
    - name: Name for the new worksheet (max 31 chars, no special characters).
    """
    return workbook_tools.add_sheet(name)


@mcp.tool()
def rename_sheet(sheet_name: str, new_name: str) -> str:
    """
    Renames a worksheet in the active workbook.

    Parameters:
    - sheet_name: Current name of the worksheet.
    - new_name: New name to assign.
    """
    return workbook_tools.rename_sheet(sheet_name, new_name)


@mcp.tool()
def delete_sheet(sheet_name: str) -> str:
    """
    Deletes a worksheet. Refuses to delete the last remaining sheet.

    Parameters:
    - sheet_name: Name of the worksheet to remove.
    """
    return workbook_tools.delete_sheet(sheet_name)


@mcp.tool()
def export_to_pdf(pdf_path: str, sheet_name: str = None) -> str:
    """
    Exports the active workbook or a specific sheet to a PDF file.

    Parameters:
    - pdf_path: Output path for the PDF file.
    - sheet_name: Optional worksheet name. If omitted, exports the entire workbook.
    """
    return workbook_tools.export_to_pdf(pdf_path, sheet_name)


@mcp.tool()
def protect_sheet(sheet_name: str, password: str = None, protect: bool = True) -> str:
    """
    Protects or unprotects a worksheet.
    Passwords are never logged. Note: Excel sheet protection is weak by design.

    Parameters:
    - sheet_name: Name of the worksheet.
    - password: Optional protection password.
    - protect: True to protect, False to unprotect.
    """
    return workbook_tools.protect_sheet(sheet_name, password, protect)


@mcp.tool()
def set_sheet_visibility(sheet_name: str, visible: bool = True) -> str:
    """
    Hides or shows a worksheet.

    Parameters:
    - sheet_name: Name of the worksheet.
    - visible: True to show, False to hide.
    """
    return workbook_tools.set_sheet_visibility(sheet_name, visible)


# ══════════════════════════════════════════════════════════════════════════════
# SHEET INTROSPECTION (Phase 0 — ALWAYS call before writing)
# ══════════════════════════════════════════════════════════════════════════════

@mcp.tool()
def describe_sheet(sheet: str) -> str:
    """
    Returns a precise description of the used range in a worksheet.
    ALWAYS call this before writing to know the exact used range, last row, last column,
    and existing headers — eliminates range-guessing errors.

    Returns: used_range address, last_row, last_col, last_col_letter,
             total_rows, total_cols, headers list, and named tables.

    Parameters:
    - sheet: Name of the worksheet to inspect.
    """
    return table_tools.describe_sheet(sheet)


@mcp.tool()
def describe_workbook(workbook_id: str = None) -> str:
    """
    [PHASE 8 INTROSPECTION] Returns a full structural map of the workbook:
    all sheets (visible & hidden), used range dimensions, named tables, and named ranges.

    Parameters:
    - workbook_id: Optional workbook ID or path. Defaults to active workbook.
    """
    return table_tools.describe_workbook(workbook_id)


@mcp.tool()
def get_formula_graph(sheet: str, range_address: str = None) -> str:
    """
    [PHASE 8 INTROSPECTION] Extracts all formulas in a range and builds a formula dependency graph
    showing which cells reference which targets.

    Parameters:
    - sheet: Worksheet name.
    - range_address: Range to inspect. Defaults to used range.
    """
    return table_tools.get_formula_graph(sheet, range_address)


# ══════════════════════════════════════════════════════════════════════════════
# DATA OPERATIONS
# ══════════════════════════════════════════════════════════════════════════════

@mcp.tool()
def read_data(sheet: str, range_address: str = None) -> str:
    """
    Reads cell data from a worksheet and returns it as JSON.
    Returns list-of-dicts if first row looks like headers, else raw list-of-lists.

    Parameters:
    - sheet: Name of the target worksheet.
    - range_address: Optional range (e.g. 'A1:C50'). Reads entire used range if omitted.
    """
    return data_tools.read_data(sheet, range_address)


@mcp.tool()
def write_table(
    sheet: str,
    start_cell: str,
    data: list,
    column_formats: dict = None,
    column_order: list[str] = None,
    table_name: str = None,
    freeze_header: bool = True,
    verify: bool = True,
    allow_formulas: bool = False,
) -> str:
    """
    [PREFERRED] Self-finishing table write. Runs the only correct order:
    reorder columns → write → number formats → autofit → clamp widths → freeze header → Excel Table → verify.

    Eliminates the '####' first-render bug. Auto-detects date/currency/percentage formats.
    Reads back and verifies after writing.

    Parameters:
    - sheet: Worksheet name.
    - start_cell: Top-left cell (e.g. 'A1').
    - data: list-of-lists (first row = headers) or list-of-dicts.
    - column_formats: Optional dict mapping header name → Excel format string
                      (e.g. {'Revenue': '$#,##0.00', 'Date': 'yyyy-mm-dd'}).
    - column_order: Optional list of column header names specifying the exact left-to-right order
                    (e.g. ['Full Name', 'Father Name', 'Address', 'Phone']).
    - table_name: Optional Excel Table name. Auto-generated if omitted.
    - freeze_header: Freeze the header row (default True).
    - verify: Read back and verify after writing (default True).
    - allow_formulas: Set True only for trusted data with intentional formula cells.
    """
    return table_tools.write_table(
        sheet=sheet,
        start_cell=start_cell,
        data=data,
        column_formats=column_formats,
        column_order=column_order,
        table_name=table_name,
        freeze_header=freeze_header,
        verify=verify,
        allow_formulas=allow_formulas,
    )


@mcp.tool()
def write_data(
    sheet: str,
    range_address: str,
    data: list,
    allow_formulas: bool = False,
    verify: bool = False,
) -> str:
    """
    Writes a list of rows (list-of-lists or list-of-dicts) to a worksheet.
    Applies formula injection escaping by default.

    NOTE: For production table writes with auto-formatting, use write_table() instead.

    Parameters:
    - sheet: Worksheet name.
    - range_address: Starting top-left cell (e.g. 'A1').
    - data: List of lists or list of dicts.
    - allow_formulas: Set True only for trusted formula data.
    - verify: Read back and verify after writing.
    """
    return data_tools.write_data(sheet, range_address, data, allow_formulas, verify)


@mcp.tool()
def autofit_smart(
    sheet: str,
    range_address: str = None,
    min_width: float = 8.0,
    max_width: float = 48.0,
    wrap_long_text: bool = True,
) -> str:
    """
    Autofits all columns in the range, then clamps widths between min and max.
    Long columns get wrap_text enabled automatically.

    Parameters:
    - sheet: Worksheet name.
    - range_address: Optional range (e.g. 'A1:F100'). Defaults to used range.
    - min_width: Minimum column width (default 8).
    - max_width: Maximum column width (default 48).
    - wrap_long_text: Enable word-wrap for columns clamped at max_width (default True).
    """
    return table_tools.autofit_smart(sheet, range_address, min_width, max_width, wrap_long_text)


@mcp.tool()
def remove_duplicates(sheet: str, range_address: str = None, columns: list[str] = None) -> str:
    """
    Removes duplicate rows. Preserves original data types (no int→float coercion).

    Parameters:
    - sheet: Worksheet name.
    - range_address: Range (e.g. 'A1:D100'). Defaults to entire used range.
    - columns: Optional list of column headers to check for duplicates. Defaults to all columns.
    """
    return data_tools.remove_duplicates(sheet, range_address, columns)


@mcp.tool()
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
    - range_address: Cell range. Defaults to used range.
    - filter_column: Column header to filter on (e.g. 'Status').
    - criteria: Value or threshold (e.g. 'Active' or '100').
    - target_sheet: Optional sheet name to write filtered results to.
    - operator: Comparison operator:
        'contains' (text substring, default),
        '==' (exact match), '!=' (not equal),
        '>' | '>=' | '<' | '<=' (numeric comparisons).
    """
    return data_tools.filter_data(sheet, range_address, filter_column, criteria, target_sheet, operator)


@mcp.tool()
def clean_data(sheet: str, range_address: str = None) -> str:
    """
    Cleans common data quality issues:
    - Trims whitespace from text cells
    - Converts numbers-stored-as-text to actual numbers
    - Removes zero-width spaces and non-breaking spaces

    Parameters:
    - sheet: Worksheet name.
    - range_address: Range to clean. Defaults to used range.
    """
    return editing_tools.clean_data(sheet, range_address)


@mcp.tool()
def clear_range(sheet: str, range_address: str, clear_type: str = "all") -> str:
    """
    Clears a cell range.

    Parameters:
    - sheet: Worksheet name.
    - range_address: Cells range (e.g. 'A1:C10').
    - clear_type: 'all' (contents + styles), 'contents', or 'formats'.
    """
    return data_tools.clear_range(sheet, range_address, clear_type)


@mcp.tool()
def copy_paste_range(
    source_sheet: str,
    source_range: str,
    target_sheet: str,
    target_range: str,
    paste_type: str = "all",
) -> str:
    """
    Copies a cell range from one sheet to another.

    Parameters:
    - source_sheet: Source worksheet name.
    - source_range: Range to copy (e.g. 'A1:B10').
    - target_sheet: Target worksheet name.
    - target_range: Starting target cell (e.g. 'D1').
    - paste_type: 'all' (contents + styles) or 'values' (values only).
    """
    return data_tools.copy_paste_range(source_sheet, source_range, target_sheet, target_range, paste_type)


@mcp.tool()
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
    - range_address: Optional range. Defaults to entire used range.
    - match_partial: If True, replaces substring matches (not just whole-cell). Default False.
    """
    return data_tools.find_replace(sheet, find_value, replace_value, range_address, match_partial)


# ══════════════════════════════════════════════════════════════════════════════
# STRUCTURAL EDITING (Phase 1)
# ══════════════════════════════════════════════════════════════════════════════

@mcp.tool()
def insert_rows(sheet: str, row_number: int, count: int = 1) -> str:
    """
    Inserts blank rows above a given row number.

    Parameters:
    - sheet: Worksheet name.
    - row_number: 1-based row number above which to insert.
    - count: Number of rows to insert (default 1).
    """
    return editing_tools.insert_rows(sheet, row_number, count)


@mcp.tool()
def delete_rows(sheet: str, row_number: int, count: int = 1) -> str:
    """
    Deletes rows from the worksheet.

    Parameters:
    - sheet: Worksheet name.
    - row_number: 1-based first row to delete.
    - count: Number of consecutive rows to delete (default 1).
    """
    return editing_tools.delete_rows(sheet, row_number, count)


@mcp.tool()
def insert_columns(sheet: str, column_letter: str, count: int = 1) -> str:
    """
    Inserts blank columns to the left of a given column letter.

    Parameters:
    - sheet: Worksheet name.
    - column_letter: Column letter before which to insert (e.g. 'C').
    - count: Number of columns to insert (default 1).
    """
    return editing_tools.insert_columns(sheet, column_letter, count)


@mcp.tool()
def delete_columns(sheet: str, column_letter: str, count: int = 1) -> str:
    """
    Deletes columns from the worksheet.

    Parameters:
    - sheet: Worksheet name.
    - column_letter: First column letter to delete (e.g. 'C').
    - count: Number of consecutive columns to delete (default 1).
    """
    return editing_tools.delete_columns(sheet, column_letter, count)


@mcp.tool()
def sort_range(
    sheet: str,
    range_address: str = None,
    sort_column: str = None,
    ascending: bool = True,
    header: bool = True,
) -> str:
    """
    Sorts a range by a column header value.

    Parameters:
    - sheet: Worksheet name.
    - range_address: Range to sort. Defaults to used range.
    - sort_column: Column header name to sort by (e.g. 'Revenue').
    - ascending: True for A→Z / smallest→largest, False for descending.
    - header: True if the first row is a header row (default True).
    """
    return editing_tools.sort_range(sheet, range_address, sort_column, ascending, header)


@mcp.tool()
def apply_autofilter(sheet: str, range_address: str = None) -> str:
    """
    Enables Excel's native AutoFilter drop-down arrows on a range.

    Parameters:
    - sheet: Worksheet name.
    - range_address: Range to filter (e.g. 'A1:F100'). Defaults to used range.
    """
    return editing_tools.apply_autofilter(sheet, range_address)


@mcp.tool()
def remove_autofilter(sheet: str) -> str:
    """
    Removes the AutoFilter from a worksheet.

    Parameters:
    - sheet: Worksheet name.
    """
    return editing_tools.remove_autofilter(sheet)


@mcp.tool()
def create_named_range(name: str, sheet: str, range_address: str) -> str:
    """
    Creates a named range in the workbook (e.g. 'SalesData' → 'Sheet1!A1:D100').

    Parameters:
    - name: Range name (e.g. 'SalesData'). Must start with a letter or underscore.
    - sheet: Worksheet name.
    - range_address: Cell range to name (e.g. 'A1:D100').
    """
    return editing_tools.create_named_range(name, sheet, range_address)


@mcp.tool()
def delete_named_range(name: str) -> str:
    """
    Deletes a named range from the workbook.

    Parameters:
    - name: The named range to delete.
    """
    return editing_tools.delete_named_range(name)


@mcp.tool()
def list_named_ranges() -> str:
    """Lists all named ranges in the active workbook."""
    return editing_tools.list_named_ranges()


@mcp.tool()
def add_comment(sheet: str, cell_address: str, comment_text: str, author: str = "Excel MCP") -> str:
    """
    Adds or updates a comment on a cell.

    Parameters:
    - sheet: Worksheet name.
    - cell_address: Single cell address (e.g. 'B5').
    - comment_text: The comment text.
    - author: Optional author name (default 'Excel MCP').
    """
    return editing_tools.add_comment(sheet, cell_address, comment_text, author)


@mcp.tool()
def delete_comment(sheet: str, cell_address: str) -> str:
    """
    Deletes a comment from a cell.

    Parameters:
    - sheet: Worksheet name.
    - cell_address: Single cell address (e.g. 'B5').
    """
    return editing_tools.delete_comment(sheet, cell_address)


# ══════════════════════════════════════════════════════════════════════════════
# FORMATTING & FORMULA TOOLS
# ══════════════════════════════════════════════════════════════════════════════

@mcp.tool()
def apply_formula(sheet: str, range_address: str, formula: str) -> str:
    """
    Writes an Excel formula to a cell or range.

    Parameters:
    - sheet: Worksheet name.
    - range_address: Destination cell or range (e.g. 'C15' or 'D2:D100').
    - formula: Formula starting with '=' (e.g. '=SUM(B2:B14)').
    """
    return format_tools.apply_formula(sheet, range_address, formula)


@mcp.tool()
def format_range(
    sheet: str,
    range_address: str,
    font_bold: bool = None,
    font_color: str = None,
    fill_color: str = None,
    font_size: int = None,
    border_style: str = None,
    number_format: str = None,
    alignment: str = None,
    italic: bool = None,
    underline: bool = None,
    wrap_text: bool = None,
) -> str:
    """
    Applies visual formatting to a range.

    Parameters:
    - sheet: Worksheet name.
    - range_address: Cell range (e.g. 'A1:D1').
    - font_bold: True to bold text.
    - font_color: Hex color for font (e.g. '#FFFFFF').
    - fill_color: Hex color for background (e.g. '#1B365D').
    - font_size: Integer font size (e.g. 11, 14).
    - border_style: 'thin', 'thick', 'double', or 'none'.
    - number_format: Excel format string (e.g. '$#,##0.00', '0.0%', 'yyyy-mm-dd').
    - alignment: 'left', 'center', or 'right'.
    - italic: True for italic.
    - underline: True for underline.
    - wrap_text: True to enable word wrap.
    """
    return format_tools.format_range(
        sheet=sheet, range_address=range_address,
        font_bold=font_bold, font_color=font_color, fill_color=fill_color,
        font_size=font_size, border_style=border_style, number_format=number_format,
        alignment=alignment, italic=italic, underline=underline, wrap_text=wrap_text,
    )


@mcp.tool()
def set_column_width(
    sheet: str,
    column_letter_or_range: str = None,
    width: float = None,
    auto_fit: bool = False,
) -> str:
    """
    Adjusts column widths. Supports specific columns or auto-fitting.

    Parameters:
    - sheet: Worksheet name.
    - column_letter_or_range: Column letter (e.g. 'A') or range (e.g. 'A:D').
    - width: Numeric column width (e.g. 15.0).
    - auto_fit: If True, auto-fits to content.
    """
    return format_tools.set_column_width(sheet, column_letter_or_range, width, auto_fit)


@mcp.tool()
def set_row_height(sheet: str, range_address: str, height: float = None, auto_fit: bool = False) -> str:
    """
    Adjusts row heights.

    Parameters:
    - sheet: Worksheet name.
    - range_address: Cell range covering target rows (e.g. 'A1:A5').
    - height: Numeric row height (e.g. 25.0).
    - auto_fit: If True, auto-fits to content.
    """
    return format_tools.set_row_height(sheet, range_address, height, auto_fit)


@mcp.tool()
def freeze_panes(sheet: str, cell_address: str = "A2") -> str:
    """
    Freezes rows/columns above and to the left of the specified cell.

    Parameters:
    - sheet: Worksheet name.
    - cell_address: Anchor cell (default 'A2' = freeze Row 1). Use 'A1' to unfreeze.
    """
    return format_tools.freeze_panes(sheet, cell_address)


@mcp.tool()
def merge_cells(sheet: str, range_address: str, merge: bool = True) -> str:
    """
    Merges or unmerges a block of cells.

    Parameters:
    - sheet: Worksheet name.
    - range_address: Range of cells (e.g. 'A1:D1').
    - merge: True to merge, False to unmerge.
    """
    return format_tools.merge_cells(sheet, range_address, merge)


@mcp.tool()
def add_conditional_formatting(
    sheet: str,
    range_address: str,
    rule_type: str = "cell_value",
    operator: str = ">",
    value: str = "0",
    value2: str = None,
    fill_color: str = None,
    font_color: str = None,
    formula: str = None,
) -> str:
    """
    Adds conditional formatting to a range. Supports many rule types.

    rule_type options:
    - 'cell_value': Compare using operator + value (e.g. operator='>' value='100').
    - 'data_bar': Blue progress bar.
    - 'color_scale': Green–Yellow–Red 3-color scale.
    - 'icon_set': Traffic-light icon arrows.
    - 'top_n': Highlight top N values (set value='10').
    - 'bottom_n': Highlight bottom N values.
    - 'above_average' / 'below_average': Statistical comparison.
    - 'duplicate' / 'unique': Highlight duplicate or unique cells.
    - 'text_contains': Highlight cells containing value as text.
    - 'formula': Custom formula rule (set formula='=A1>100').

    Parameters:
    - sheet: Worksheet name.
    - range_address: Cell range (e.g. 'C2:C100').
    - rule_type: See above.
    - operator: For cell_value: '>', '<', '>=', '<=', '==', '!=', 'between'.
    - value: Threshold value or top-N count.
    - value2: Second value for 'between' operator.
    - fill_color: Hex background color (e.g. '#FFC7CE').
    - font_color: Hex font color (e.g. '#9C0006').
    - formula: Custom formula for rule_type='formula'.
    """
    return format_tools.add_conditional_formatting(
        sheet=sheet, range_address=range_address,
        rule_type=rule_type, operator=operator, value=value, value2=value2,
        fill_color=fill_color, font_color=font_color, formula=formula,
    )


@mcp.tool()
def add_dropdown_validation(sheet: str, range_address: str, choices: list[str]) -> str:
    """
    Adds an in-cell drop-down list validation to a range.
    Handles lists exceeding Excel's 255-char inline limit automatically.

    Parameters:
    - sheet: Worksheet name.
    - range_address: Cell range (e.g. 'B2:B100').
    - choices: List of valid options (e.g. ['Approved', 'Pending', 'Rejected']).
    """
    return format_tools.add_dropdown_validation(sheet, range_address, choices)


@mcp.tool()
def add_data_validation(
    sheet: str,
    range_address: str,
    validation_type: str = "whole_number",
    operator: str = "between",
    value1: str = "0",
    value2: str = None,
    input_title: str = None,
    input_message: str = None,
    error_title: str = "Invalid Input",
    error_message: str = "The value entered is not valid.",
) -> str:
    """
    Adds numeric, date, text-length, or formula data validation to a range.

    validation_type: 'whole_number', 'decimal', 'date', 'time', 'text_length', 'formula'.
    operator: 'between', 'not_between', '=', '!=', '>', '<', '>=', '<='.

    Parameters:
    - sheet: Worksheet name.
    - range_address: Cell range (e.g. 'A2:A100').
    - validation_type: Type of validation.
    - operator: Comparison operator.
    - value1: First bound or formula.
    - value2: Second bound (for 'between'/'not_between').
    - input_title: Optional tooltip title shown when cell is selected.
    - input_message: Optional tooltip message shown when cell is selected.
    - error_title: Error dialog title (default 'Invalid Input').
    - error_message: Error dialog message.
    """
    return format_tools.add_data_validation(
        sheet=sheet, range_address=range_address,
        validation_type=validation_type, operator=operator,
        value1=value1, value2=value2,
        input_title=input_title, input_message=input_message,
        error_title=error_title, error_message=error_message,
    )


@mcp.tool()
def add_hyperlink(sheet: str, range_address: str, url: str, display_text: str = None) -> str:
    """
    Inserts a clickable hyperlink into a cell.

    Parameters:
    - sheet: Worksheet name.
    - range_address: Destination cell (e.g. 'A10').
    - url: Target URL (e.g. 'https://example.com').
    - display_text: Optional visible link text.
    """
    return format_tools.add_hyperlink(sheet, range_address, url, display_text)


@mcp.tool()
def fill_range(sheet: str, range_address: str, direction: str = "down") -> str:
    """
    Auto-fills cells by repeating a formula or series across a range.

    Parameters:
    - sheet: Worksheet name.
    - range_address: Range to fill (e.g. 'D2:D100').
    - direction: 'down', 'up', 'right', or 'left'.
    """
    return format_tools.fill_range(sheet, range_address, direction)


# ══════════════════════════════════════════════════════════════════════════════
# CHARTS & PIVOT TABLES
# ══════════════════════════════════════════════════════════════════════════════

@mcp.tool()
def create_chart(
    sheet: str,
    chart_type: str,
    source_range: str,
    title: str = None,
    x_axis_title: str = "",
    y_axis_title: str = "",
    target_cell: str = "G2",
    width: int = 420,
    height: int = 260,
    plot_by: str = "columns",
    show_legend: bool = True,
    show_data_labels: bool = False,
) -> str:
    """
    Creates a native Excel chart positioned next to your data.

    Parameters:
    - sheet: Worksheet name.
    - chart_type: 'column', 'bar', 'line', 'pie', 'area', 'scatter', 'doughnut',
                  'column_stacked', 'bar_stacked', 'line_markers', 'radar', etc.
    - source_range: Data range including headers (e.g. 'A1:B10').
    - title: Optional chart title.
    - x_axis_title: Optional X axis label.
    - y_axis_title: Optional Y axis label.
    - target_cell: Top-left cell for chart placement (default 'G2').
    - width: Chart width in points (default 420).
    - height: Chart height in points (default 260).
    - plot_by: 'columns' (default) or 'rows' — controls data series orientation.
    - show_legend: Show chart legend (default True).
    - show_data_labels: Show data labels on series (default False).
    """
    return chart_tools.create_chart(
        sheet=sheet, chart_type=chart_type, source_range=source_range,
        title=title, x_axis_title=x_axis_title, y_axis_title=y_axis_title,
        target_cell=target_cell, width=width, height=height,
        plot_by=plot_by, show_legend=show_legend, show_data_labels=show_data_labels,
    )


# ══════════════════════════════════════════════════════════════════════════════
# WORKFLOW ENGINE & TRANSACTIONS (Phases 1 - 9)
# ══════════════════════════════════════════════════════════════════════════════

@mcp.tool()
def start_workflow(plan: dict) -> str:
    """
    [ENTERPRISE WORKFLOW ENGINE] Executes a multi-step client-authored DAG plan with step dependencies,
    variable resolution (e.g. {{step_1.result.range}}), step transaction snapshots, and automatic rollback.

    Plan Schema (§4.1):
      {
        "workflow_id": "sales_analysis_001",
        "transaction": true,
        "workbook_path": "C:\\path\\to\\sales.xlsx",
        "operations": [
           { "id": "step_1", "operation": "open_workbook", "parameters": { "file_path": "sales.xlsx" } },
           { "id": "step_2", "operation": "clean_data", "parameters": { "sheet": "Sales" }, "depends_on": ["step_1"] },
           { "id": "step_3", "operation": "write_table", "parameters": { "sheet": "Summary", "start_cell": "A1", "data": "..." }, "depends_on": ["step_2"] }
        ]
      }
    """
    import workflow_engine
    result = workflow_engine.start_workflow(plan)
    return result_to_str(result)


@mcp.tool()
def get_workflow_status(workflow_id: str) -> str:
    """
    Returns the step-by-step execution status and output journal of a workflow.

    Parameters:
    - workflow_id: The ID of the workflow (e.g. 'sales_analysis_001').
    """
    import workflow_engine
    result = workflow_engine.get_workflow_status(workflow_id)
    return result_to_str(result)


@mcp.tool()
def resume_workflow(workflow_id: str, from_step: str = None) -> str:
    """
    Resumes a paused or failed workflow from disk state, optionally re-running from a specified step.

    Parameters:
    - workflow_id: The ID of the workflow to resume.
    - from_step: Optional step ID to resume execution from (e.g. 'step_4').
    """
    import workflow_engine
    result = workflow_engine.resume_workflow(workflow_id, from_step)
    return result_to_str(result)


@mcp.tool()
def cancel_workflow(workflow_id: str) -> str:
    """
    Cancels a running or paused workflow and triggers transaction rollback.

    Parameters:
    - workflow_id: The ID of the workflow to cancel.
    """
    import workflow_engine
    result = workflow_engine.cancel_workflow(workflow_id)
    return result_to_str(result)


@mcp.tool()
def run_control_totals(sheet: str, range_address: str, checks: list[dict]) -> str:
    """
    Executes pluggable control total checks (row count, numeric sum, null count, data hash) on a sheet range.

    Parameters:
    - sheet: Worksheet name.
    - range_address: Range address to check (e.g. 'A1:D100').
    - checks: List of check definitions (e.g. [{"type": "row_count", "expected": 100}, {"type": "numeric_sum", "expected": 150000.50}]).
    """
    try:
        from workbook_tools import get_active_book, get_sheet
        import control_totals
        book = get_active_book()
        s = get_sheet(book, sheet)
        res = control_totals.run_control_total_pipeline(s, range_address, checks)
        return result_to_str(ok("Control totals completed.", data=res))
    except Exception as e:
        return result_to_str(err("CONTROL_TOTALS_ERROR", str(e)))


@mcp.tool()
def compare_workbooks(workbook_id_1: str, workbook_id_2: str) -> str:
    """
    Compares two workbooks structurally (sheets added/removed, shape differences).

    Parameters:
    - workbook_id_1: ID or path/name of first workbook.
    - workbook_id_2: ID or path/name of second workbook.
    """
    try:
        import workbook_registry
        res = workbook_registry.registry.compare_workbooks(workbook_id_1, workbook_id_2)
        return result_to_str(ok("Workbook comparison completed.", data=res))
    except Exception as e:
        return result_to_str(err("COMPARE_WORKBOOKS_ERROR", str(e)))


@mcp.tool()
def evaluate_rule(expression: str, context: dict) -> str:
    """
    Safely evaluates a business rule condition against a dictionary context using a sandboxed AST parser.

    Parameters:
    - expression: Logical expression string (e.g. "variance_percent > 10 and status == 'Active'").
    - context: Key-value dictionary containing variable values (e.g. {"variance_percent": 12.5, "status": "Active"}).
    """
    try:
        import rules_evaluator
        val = rules_evaluator.evaluate_expression(expression, context)
        return result_to_str(ok(f"Rule evaluated to {val}", data={"result": val}))
    except Exception as e:
        return result_to_str(err("EVALUATE_RULE_ERROR", str(e)))


@mcp.tool()
def create_pivot_table(
    source_sheet: str,
    source_range: str = None,
    target_sheet: str = "Pivot_Summary",
    target_cell: str = "A3",
    row_fields: list[str] = None,
    col_fields: list[str] = None,
    data_field: str = None,
    agg_func: str = "sum",
) -> str:
    """
    Generates a Pivot Table summarizing data. Falls back to Pandas if COM fails.

    Parameters:
    - source_sheet: Worksheet containing the source data.
    - source_range: Data range (e.g. 'A1:F500'). Auto-detects if omitted.
    - target_sheet: Worksheet to write the Pivot Table to (default 'Pivot_Summary').
    - target_cell: Top-left cell for the pivot (default 'A3').
    - row_fields: Column headers to group by rows.
    - col_fields: Column headers to group by columns (optional).
    - data_field: Column header to aggregate.
    - agg_func: 'sum', 'count', 'average', 'max', 'min' (default 'sum').
    """
    return chart_tools.create_pivot_table(
        source_sheet=source_sheet, source_range=source_range,
        target_sheet=target_sheet, target_cell=target_cell,
        row_fields=row_fields, col_fields=col_fields,
        data_field=data_field, agg_func=agg_func,
    )


# ══════════════════════════════════════════════════════════════════════════════
# VERIFICATION & ACCURACY (Phase 4)
# ══════════════════════════════════════════════════════════════════════════════

@mcp.tool()
def verify_range(sheet: str, range_address: str, expected_data: list) -> str:
    """
    Reads back a range and verifies it element-by-element against expected_data.
    Returns a full mismatch report — use after any critical write.

    Parameters:
    - sheet: Worksheet name.
    - range_address: Range to verify (e.g. 'A1:D100').
    - expected_data: 2D list (list-of-lists) of expected values.
    """
    try:
        from workbook_tools import get_active_book, get_sheet
        book = get_active_book()
        s = get_sheet(book, sheet)
        result = verification.verify_write(s, range_address, expected_data)
        status = "ok" if result["verified"] else "error"
        return result_to_str({
            "status": status,
            "code": "VERIFIED" if result["verified"] else "VERIFY_FAILED",
            "message": "Verification passed." if result["verified"] else f"{len(result['mismatches'])} mismatches found.",
            "data": result,
        })
    except Exception as e:
        return result_to_str(err("VERIFY_RANGE_ERROR", str(e)))


@mcp.tool()
def scan_errors(sheet: str, range_address: str) -> str:
    """
    Scans a range for Excel error values (#REF!, #DIV/0!, #VALUE!, #N/A, etc.)
    after forcing a recalculation. Returns list of error cells.

    Parameters:
    - sheet: Worksheet name.
    - range_address: Range to scan (e.g. 'A1:Z200').
    """
    try:
        from workbook_tools import get_active_book, get_sheet
        book = get_active_book()
        s = get_sheet(book, sheet)
        result = verification.scan_formula_errors(s, range_address)
        status = "ok" if result["clean"] else "error"
        return result_to_str({
            "status": status,
            "code": "NO_ERRORS" if result["clean"] else "FORMULA_ERRORS_FOUND",
            "message": "No formula errors found." if result["clean"] else f"{len(result['error_cells'])} error cell(s) found.",
            "data": result,
        })
    except Exception as e:
        return result_to_str(err("SCAN_ERRORS_ERROR", str(e)))


# ══════════════════════════════════════════════════════════════════════════════
# SNAPSHOT & ROLLBACK (Phase 4)
# ══════════════════════════════════════════════════════════════════════════════

@mcp.tool()
def create_snapshot(workbook_path: str) -> str:
    """
    Saves a timestamped backup copy of a workbook before making critical changes.
    Snapshots are stored in a .excel_mcp_snapshots folder next to the workbook.

    Parameters:
    - workbook_path: Absolute path to the workbook file.
    """
    try:
        snap = snapshot_tools.create_snapshot(workbook_path)
        return result_to_str(ok(f"Snapshot created: '{snap}'", data={"snapshot_path": snap}))
    except Exception as e:
        return result_to_str(err("SNAPSHOT_ERROR", str(e)))


@mcp.tool()
def restore_snapshot(snapshot_path: str, workbook_path: str) -> str:
    """
    Restores a workbook from a previously created snapshot (rollback).

    Parameters:
    - snapshot_path: Path to the snapshot file.
    - workbook_path: Path to the workbook to restore to.
    """
    try:
        msg = snapshot_tools.restore_snapshot(snapshot_path, workbook_path)
        return result_to_str(ok(msg))
    except Exception as e:
        return result_to_str(err("RESTORE_SNAPSHOT_ERROR", str(e)))


@mcp.tool()
def list_snapshots(workbook_path: str) -> str:
    """
    Lists all available snapshots for a workbook, newest first.

    Parameters:
    - workbook_path: Absolute path to the workbook.
    """
    try:
        snaps = snapshot_tools.list_snapshots(workbook_path)
        return result_to_str(ok(f"Found {len(snaps)} snapshot(s).", data=snaps))
    except Exception as e:
        return result_to_str(err("LIST_SNAPSHOTS_ERROR", str(e)))


# ══════════════════════════════════════════════════════════════════════════════
# DRY-RUN PREVIEW (Phase 5)
# ══════════════════════════════════════════════════════════════════════════════

@mcp.tool()
def preview_write(sheet: str, range_address: str, data: list) -> str:
    """
    Generate a before/after diff preview WITHOUT writing anything.
    Shows exactly which cells will change, from what value, to what value.
    Call this before any large write to review changes, then call write_table with confirmed=True.

    Parameters:
    - sheet: Worksheet name.
    - range_address: Starting cell or range to preview (e.g. 'A1').
    - data: The data you plan to write (list-of-lists or list-of-dicts).
    """
    try:
        from workbook_tools import get_active_book, get_sheet
        book = get_active_book()
        s = get_sheet(book, sheet)
        preview = dry_run.preview_write(s, range_address, data)
        summary = dry_run.format_preview_summary(preview)
        return result_to_str(ok(summary, data=preview))
    except Exception as e:
        return result_to_str(err("PREVIEW_WRITE_ERROR", str(e)))


# ══════════════════════════════════════════════════════════════════════════════
# SECURITY AUDIT (Phase 5)
# ══════════════════════════════════════════════════════════════════════════════

@mcp.tool()
def verify_audit_chain() -> str:
    """
    Verifies the integrity of the tamper-evident audit log.
    Checks that no entries have been modified since they were written.
    Returns 'ok' if the chain is intact, 'error' with the broken entry index if tampered.
    """
    try:
        result = security.verify_audit_chain()
        status = "ok" if result["ok"] else "error"
        return result_to_str({
            "status": status,
            "code": "CHAIN_INTACT" if result["ok"] else "CHAIN_BROKEN",
            "message": result["message"],
            "data": result,
        })
    except Exception as e:
        return result_to_str(err("AUDIT_VERIFY_ERROR", str(e)))


# ══════════════════════════════════════════════════════════════════════════════
# ENTRY POINT
# ══════════════════════════════════════════════════════════════════════════════

if __name__ == "__main__":
    logger.info("Excel Automation MCP Server starting (production-ready, all phases).")
    mcp.run(transport="stdio")
