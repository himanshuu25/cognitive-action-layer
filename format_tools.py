"""
format_tools.py — Production-hardened
All formatting operations use the shared excel_suppressed context manager,
return structured results, and validate inputs before touching COM.
Added: rich conditional formatting, real data validation, improved range targeting.
"""

from __future__ import annotations
import re
from typing import Optional, Union

import xlwings as xw

from workbook_tools import get_active_book, get_sheet
from shared_context import (
    excel_suppressed,
    validate_range_address,
    validate_sheet_name,
    ok, err, result_to_str,
)
from logger import logger


# ── Color Helpers ─────────────────────────────────────────────────────────────

def hex_to_rgb(hex_str: str):
    """Converts a hex color string (e.g. '#1B365D') to an (R, G, B) tuple."""
    if not hex_str:
        return None
    hex_str = hex_str.lstrip("#")
    if len(hex_str) == 3:
        hex_str = "".join([c * 2 for c in hex_str])
    if len(hex_str) != 6:
        raise ValueError(f"Invalid hex color: '#{hex_str}'")
    return tuple(int(hex_str[i : i + 2], 16) for i in (0, 2, 4))


def rgb_to_bgr_int(rgb: tuple) -> int:
    """Converts an (R, G, B) tuple to Excel's BGR integer format for COM APIs."""
    r, g, b = rgb
    return r + g * 256 + b * 65536


# ── Formulas ──────────────────────────────────────────────────────────────────

def apply_formula(sheet: str, range_address: str, formula: str) -> str:
    """Applies an Excel formula to a cell or range. Validates formula starts with '='."""
    try:
        validate_sheet_name(sheet)
        range_address = validate_range_address(range_address)
        book = get_active_book()
        s = get_sheet(book, sheet)

        if not formula.startswith("="):
            formula = "=" + formula

        with excel_suppressed(book.app):
            s.range(range_address).formula = formula

        return result_to_str(ok(f"Applied formula '{formula}' to '{sheet}'!{range_address}."))
    except Exception as e:
        logger.exception("apply_formula failed")
        return result_to_str(err("APPLY_FORMULA_ERROR", str(e)))


# ── Cell Formatting ───────────────────────────────────────────────────────────

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
    - range_address: Cell range (e.g. 'A1:D10').
    - font_bold: True to bold text.
    - font_color: Hex color string for font (e.g. '#FFFFFF').
    - fill_color: Hex color string for cell background (e.g. '#1B365D').
    - font_size: Integer font size (e.g. 11, 14, 16).
    - border_style: 'thin', 'thick', 'double', or 'none'.
    - number_format: Excel number format string (e.g. '$#,##0.00', '0.0%', 'yyyy-mm-dd').
    - alignment: 'left', 'center', or 'right'.
    - italic: True for italic text.
    - underline: True for underlined text.
    - wrap_text: True to enable word wrap in the cell.
    """
    try:
        validate_sheet_name(sheet)
        range_address = validate_range_address(range_address)
        book = get_active_book()
        s = get_sheet(book, sheet)

        with excel_suppressed(book.app):
            r = s.range(range_address)

            if font_bold is not None:
                r.font.bold = font_bold
            if italic is not None:
                r.font.italic = italic
            if underline is not None:
                r.font.underline = underline
            if font_color:
                r.font.color = hex_to_rgb(font_color)
            if fill_color:
                r.color = hex_to_rgb(fill_color)
            if font_size:
                r.font.size = font_size
            if number_format:
                r.number_format = number_format
            if wrap_text is not None:
                try:
                    r.api.WrapText = wrap_text
                except Exception:
                    pass
            if alignment:
                align_map = {"left": -4131, "center": -4108, "right": -4152}
                val = align_map.get(alignment.lower())
                if val:
                    try:
                        r.api.HorizontalAlignment = val
                    except Exception:
                        pass
            if border_style:
                _apply_border(r, border_style)

        return result_to_str(ok(f"Formatted '{sheet}'!{range_address}."))
    except Exception as e:
        logger.exception("format_range failed")
        return result_to_str(err("FORMAT_RANGE_ERROR", str(e)))


def _apply_border(rng, border_style: str):
    """Apply border style to a range via COM."""
    b = border_style.lower()
    try:
        if b == "thin":
            rng.api.Borders.LineStyle = 1
            rng.api.Borders.Weight = 2
        elif b == "thick":
            rng.api.Borders.LineStyle = 1
            rng.api.Borders.Weight = 4
        elif b == "double":
            rng.api.Borders.LineStyle = 9
        elif b == "none":
            rng.api.Borders.LineStyle = -4142
    except Exception:
        pass


# ── Column / Row Sizing ───────────────────────────────────────────────────────

def set_column_width(
    sheet: str,
    column_letter_or_range: str = None,
    width: float = None,
    auto_fit: bool = False,
) -> str:
    """Sets column width or auto-fits columns."""
    try:
        validate_sheet_name(sheet)
        book = get_active_book()
        s = get_sheet(book, sheet)

        with excel_suppressed(book.app):
            if column_letter_or_range:
                rng = s.range(column_letter_or_range)
                if auto_fit:
                    rng.columns.autofit()
                    return result_to_str(ok(f"Auto-fitted columns '{column_letter_or_range}' in '{sheet}'."))
                elif width is not None:
                    rng.column_width = width
                    return result_to_str(ok(f"Set column width of '{column_letter_or_range}' to {width} in '{sheet}'."))
            else:
                if auto_fit:
                    s.autofit(axis="columns")
                    return result_to_str(ok(f"Auto-fitted all columns in '{sheet}'."))

        return result_to_str(err("INVALID_INPUT", "Provide column_letter_or_range and either width or auto_fit=True."))
    except Exception as e:
        logger.exception("set_column_width failed")
        return result_to_str(err("SET_COLUMN_WIDTH_ERROR", str(e)))


def set_row_height(
    sheet: str,
    range_address: str,
    height: float = None,
    auto_fit: bool = False,
) -> str:
    """Sets row height or auto-fits rows."""
    try:
        validate_sheet_name(sheet)
        book = get_active_book()
        s = get_sheet(book, sheet)

        with excel_suppressed(book.app):
            r = s.range(range_address)
            if auto_fit:
                r.rows.autofit()
                return result_to_str(ok(f"Auto-fitted rows for '{range_address}' in '{sheet}'."))
            elif height is not None:
                r.row_height = height
                return result_to_str(ok(f"Set row height for '{range_address}' to {height} in '{sheet}'."))

        return result_to_str(err("INVALID_INPUT", "Provide height or auto_fit=True."))
    except Exception as e:
        logger.exception("set_row_height failed")
        return result_to_str(err("SET_ROW_HEIGHT_ERROR", str(e)))


# ── Freeze Panes ──────────────────────────────────────────────────────────────

def freeze_panes(sheet: str, cell_address: str = "A2") -> str:
    """
    Freezes panes at cell_address (rows/cols above and left are locked).
    Pass 'A1' to unfreeze completely.
    """
    try:
        validate_sheet_name(sheet)
        book = get_active_book()
        s = get_sheet(book, sheet)

        with excel_suppressed(book.app):
            s.activate()
            win = book.app.api.ActiveWindow
            win.FreezePanes = False
            if cell_address.upper() != "A1":
                win.ScrollRow = 1
                win.ScrollColumn = 1
                s.api.Activate()
                s.range(cell_address).api.Select()
                win.FreezePanes = True
                return result_to_str(ok(f"Froze panes at '{cell_address}' in '{sheet}'."))
            else:
                return result_to_str(ok(f"Unfroze all panes in '{sheet}'."))

    except Exception as e:
        logger.exception("freeze_panes failed")
        return result_to_str(err("FREEZE_PANES_ERROR", str(e)))


# ── Merge Cells ───────────────────────────────────────────────────────────────

def merge_cells(sheet: str, range_address: str, merge: bool = True) -> str:
    """Merges or unmerges a cell range."""
    try:
        validate_sheet_name(sheet)
        book = get_active_book()
        s = get_sheet(book, sheet)

        with excel_suppressed(book.app):
            r = s.range(range_address)
            if merge:
                r.api.Merge(Across=False)
                return result_to_str(ok(f"Merged '{sheet}'!{range_address}."))
            else:
                r.api.UnMerge()
                return result_to_str(ok(f"Unmerged '{sheet}'!{range_address}."))

    except Exception as e:
        logger.exception("merge_cells failed")
        return result_to_str(err("MERGE_CELLS_ERROR", str(e)))


# ── Conditional Formatting ────────────────────────────────────────────────────

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
    Adds conditional formatting to a range.

    rule_type options:
    - 'cell_value': Compare cell value using operator and value (default).
    - 'data_bar': Blue data bar (no colors needed).
    - 'color_scale': Green–Yellow–Red 3-color scale.
    - 'icon_set': Traffic-light icon set (3 arrows).
    - 'top_n': Highlight top N values (set value='10' for top 10).
    - 'bottom_n': Highlight bottom N values.
    - 'above_average': Highlight above-average cells.
    - 'below_average': Highlight below-average cells.
    - 'duplicate': Highlight duplicate values.
    - 'unique': Highlight unique values.
    - 'text_contains': Highlight cells where text contains value.
    - 'formula': Use a custom formula (set formula='=A1>100').

    Parameters:
    - operator: For cell_value: '>', '<', '>=', '<=', '==', '!=', 'between'.
    - value: Threshold value or formula string.
    - value2: Second value for 'between' operator.
    - fill_color: Background hex color (e.g. '#FFC7CE').
    - font_color: Font hex color (e.g. '#9C0006').
    - formula: Custom formula string for rule_type='formula'.
    """
    try:
        validate_sheet_name(sheet)
        range_address = validate_range_address(range_address)
        book = get_active_book()
        s = get_sheet(book, sheet)

        with excel_suppressed(book.app):
            r = s.range(range_address)

            # Clear existing conditions
            try:
                r.api.FormatConditions.Delete()
            except Exception:
                pass

            rt = rule_type.lower()

            if rt == "cell_value":
                op_map = {
                    "==": 3, "=": 3, "equal": 3,
                    "!=": 4, "not_equal": 4,
                    ">": 5, "greater_than": 5,
                    "<": 6, "less_than": 6,
                    ">=": 7, "greater_equal": 7,
                    "<=": 8, "less_equal": 8,
                    "between": 1,
                }
                op_const = op_map.get(operator.strip().lower(), 5)
                if op_const == 1 and value2:  # between
                    fc = r.api.FormatConditions.Add(Type=1, Operator=1, Formula1=str(value), Formula2=str(value2))
                else:
                    fc = r.api.FormatConditions.Add(Type=1, Operator=op_const, Formula1=str(value))
                _apply_fc_colors(fc, fill_color, font_color)
                fc.StopIfTrue = False

            elif rt == "data_bar":
                try:
                    r.api.FormatConditions.AddDatabar()
                except Exception as de:
                    return result_to_str(err("CONDITIONAL_FORMAT_ERROR", f"DataBar not supported: {de}"))

            elif rt == "color_scale":
                try:
                    cs = r.api.FormatConditions.AddColorScale(ColorScaleType=3)
                    cs.ColorScaleCriteria(1).FormatColor.Color = rgb_to_bgr_int(hex_to_rgb("#FF0000"))  # Red (low)
                    cs.ColorScaleCriteria(2).FormatColor.Color = rgb_to_bgr_int(hex_to_rgb("#FFFF00"))  # Yellow (mid)
                    cs.ColorScaleCriteria(3).FormatColor.Color = rgb_to_bgr_int(hex_to_rgb("#00B050"))  # Green (high)
                except Exception as cse:
                    return result_to_str(err("CONDITIONAL_FORMAT_ERROR", f"ColorScale not supported: {cse}"))

            elif rt == "icon_set":
                try:
                    icons = r.api.FormatConditions.AddIconSetCondition()
                    icons.IconSet = book.app.api.ActiveWorkbook.IconSets(1)  # 3Arrows
                except Exception as ie:
                    return result_to_str(err("CONDITIONAL_FORMAT_ERROR", f"IconSet not supported: {ie}"))

            elif rt in ("top_n", "top"):
                try:
                    n = int(value) if value else 10
                    fc = r.api.FormatConditions.AddTop10()
                    fc.TopBottom = 0   # xlTop10Top
                    fc.Rank = n
                    fc.Percent = False
                    _apply_fc_colors(fc, fill_color, font_color)
                except Exception as te:
                    return result_to_str(err("CONDITIONAL_FORMAT_ERROR", f"Top N not supported: {te}"))

            elif rt in ("bottom_n", "bottom"):
                try:
                    n = int(value) if value else 10
                    fc = r.api.FormatConditions.AddTop10()
                    fc.TopBottom = 1   # xlTop10Bottom
                    fc.Rank = n
                    fc.Percent = False
                    _apply_fc_colors(fc, fill_color, font_color)
                except Exception as be:
                    return result_to_str(err("CONDITIONAL_FORMAT_ERROR", f"Bottom N not supported: {be}"))

            elif rt == "above_average":
                try:
                    fc = r.api.FormatConditions.AddAboveAverage()
                    fc.AboveBelow = 0   # xlAboveAverage
                    _apply_fc_colors(fc, fill_color, font_color)
                except Exception as ae:
                    return result_to_str(err("CONDITIONAL_FORMAT_ERROR", f"AboveAverage not supported: {ae}"))

            elif rt == "below_average":
                try:
                    fc = r.api.FormatConditions.AddAboveAverage()
                    fc.AboveBelow = 1   # xlBelowAverage
                    _apply_fc_colors(fc, fill_color, font_color)
                except Exception as be2:
                    return result_to_str(err("CONDITIONAL_FORMAT_ERROR", f"BelowAverage not supported: {be2}"))

            elif rt == "duplicate":
                try:
                    fc = r.api.FormatConditions.AddUniqueValues()
                    fc.DupeUnique = 1   # xlDuplicate
                    _apply_fc_colors(fc, fill_color, font_color)
                except Exception as due:
                    return result_to_str(err("CONDITIONAL_FORMAT_ERROR", f"Duplicate rule not supported: {due}"))

            elif rt == "unique":
                try:
                    fc = r.api.FormatConditions.AddUniqueValues()
                    fc.DupeUnique = 0   # xlUnique
                    _apply_fc_colors(fc, fill_color, font_color)
                except Exception as uqe:
                    return result_to_str(err("CONDITIONAL_FORMAT_ERROR", f"Unique rule not supported: {uqe}"))

            elif rt == "text_contains":
                try:
                    fc = r.api.FormatConditions.Add(
                        Type=2,   # xlExpression
                        Formula1=f'=ISNUMBER(SEARCH("{value}",{r.address.split(":")[0]}))'
                    )
                    _apply_fc_colors(fc, fill_color, font_color)
                    fc.StopIfTrue = False
                except Exception as txe:
                    return result_to_str(err("CONDITIONAL_FORMAT_ERROR", f"TextContains not supported: {txe}"))

            elif rt == "formula":
                if not formula:
                    return result_to_str(err("INVALID_INPUT", "formula parameter is required for rule_type='formula'."))
                if not formula.startswith("="):
                    formula = "=" + formula
                fc = r.api.FormatConditions.Add(Type=2, Formula1=formula)
                _apply_fc_colors(fc, fill_color, font_color)
                fc.StopIfTrue = False

            else:
                return result_to_str(err(
                    "UNKNOWN_RULE_TYPE",
                    f"Unknown rule_type '{rule_type}'. "
                    "Valid: cell_value, data_bar, color_scale, icon_set, top_n, bottom_n, "
                    "above_average, below_average, duplicate, unique, text_contains, formula."
                ))

        return result_to_str(ok(
            f"Conditional formatting ('{rt}') applied to '{sheet}'!{range_address}."
        ))
    except Exception as e:
        logger.exception("add_conditional_formatting failed")
        return result_to_str(err("CONDITIONAL_FORMAT_ERROR", str(e)))


def _apply_fc_colors(fc, fill_color: str, font_color: str):
    """Apply fill and font colors to a FormatCondition object."""
    try:
        if fill_color:
            fc.Interior.Color = rgb_to_bgr_int(hex_to_rgb(fill_color))
    except Exception:
        pass
    try:
        if font_color:
            fc.Font.Color = rgb_to_bgr_int(hex_to_rgb(font_color))
    except Exception:
        pass


# ── Data Validation ───────────────────────────────────────────────────────────

def add_dropdown_validation(sheet: str, range_address: str, choices: list[str]) -> str:
    """
    Adds an in-cell drop-down list to a range.
    For lists exceeding Excel's 255-char inline limit, automatically creates
    a named hidden range and references it instead.
    """
    try:
        validate_sheet_name(sheet)
        range_address = validate_range_address(range_address)
        book = get_active_book()
        s = get_sheet(book, sheet)

        with excel_suppressed(book.app):
            r = s.range(range_address)
            try:
                r.api.Validation.Delete()
            except Exception:
                pass

            choices_str = ",".join(str(c) for c in choices)

            if len(choices_str) <= 255:
                # Inline list
                r.api.Validation.Add(
                    Type=3,         # xlValidateList
                    AlertStyle=1,   # xlValidAlertStop
                    Operator=1,
                    Formula1=f'"{choices_str}"'
                )
            else:
                # Use a hidden sheet as a reference range to bypass 255-char limit
                helper_sheet_name = "_MCP_Validation_Lists"
                helper = None
                for sh in book.sheets:
                    if sh.name == helper_sheet_name:
                        helper = sh
                        break
                if not helper:
                    helper = book.sheets.add(name=helper_sheet_name)
                    helper.api.Visible = 2  # xlSheetVeryHidden

                # Find next empty column in helper sheet
                try:
                    last_col = helper.used_range.last_cell.column + 1
                except Exception:
                    last_col = 1
                from table_tools import _col_num_to_letter
                col_letter = _col_num_to_letter(last_col)
                for i, choice in enumerate(choices, 1):
                    helper.range(f"{col_letter}{i}").value = str(choice)
                list_range = f"'{helper_sheet_name}'!${col_letter}$1:${col_letter}${len(choices)}"
                r.api.Validation.Add(Type=3, AlertStyle=1, Operator=1, Formula1=f"={list_range}")

            r.api.Validation.IgnoreBlank = True
            r.api.Validation.InCellDropdown = True
            r.api.Validation.ShowError = True
            r.api.Validation.ErrorTitle = "Invalid Entry"
            r.api.Validation.ErrorMessage = f"Please choose from: {', '.join(str(c) for c in choices[:10])}"

        return result_to_str(ok(
            f"Dropdown validation added to '{sheet}'!{range_address}. {len(choices)} choices.",
            data={"choices": choices}
        ))
    except Exception as e:
        logger.exception("add_dropdown_validation failed")
        return result_to_str(err("DROPDOWN_VALIDATION_ERROR", str(e)))


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
    Adds data validation to a range (numeric, date, text length, or formula).

    validation_type options:
    - 'whole_number': Integer validation.
    - 'decimal': Decimal number validation.
    - 'date': Date validation.
    - 'time': Time validation.
    - 'text_length': Limit text length.
    - 'formula': Custom formula validation.

    operator options: 'between', 'not_between', '=', '!=', '>', '<', '>=', '<='
    """
    try:
        validate_sheet_name(sheet)
        range_address = validate_range_address(range_address)
        book = get_active_book()
        s = get_sheet(book, sheet)

        type_map = {
            "whole_number": 1,
            "decimal": 2,
            "date": 4,
            "time": 5,
            "text_length": 6,
            "formula": 8,
        }
        op_map = {
            "between": 1, "not_between": 2,
            "=": 3, "equal": 3,
            "!=": 4, "not_equal": 4,
            ">": 5, "greater_than": 5,
            "<": 6, "less_than": 6,
            ">=": 7, "greater_equal": 7,
            "<=": 8, "less_equal": 8,
        }

        vtype = type_map.get(validation_type.lower())
        if vtype is None:
            return result_to_str(err("INVALID_TYPE", f"Unknown validation_type '{validation_type}'."))

        op_const = op_map.get(operator.lower(), 1)

        with excel_suppressed(book.app):
            r = s.range(range_address)
            try:
                r.api.Validation.Delete()
            except Exception:
                pass

            kwargs = dict(Type=vtype, AlertStyle=1, Operator=op_const, Formula1=str(value1))
            if value2 and op_const in (1, 2):
                kwargs["Formula2"] = str(value2)
            r.api.Validation.Add(**kwargs)
            r.api.Validation.IgnoreBlank = True
            r.api.Validation.ShowError = True
            r.api.Validation.ErrorTitle = error_title
            r.api.Validation.ErrorMessage = error_message

            if input_title or input_message:
                r.api.Validation.ShowInput = True
                r.api.Validation.InputTitle = input_title or ""
                r.api.Validation.InputMessage = input_message or ""

        return result_to_str(ok(
            f"Data validation ({validation_type}, {operator} {value1}"
            + (f"–{value2}" if value2 else "")
            + f") added to '{sheet}'!{range_address}."
        ))
    except Exception as e:
        logger.exception("add_data_validation failed")
        return result_to_str(err("DATA_VALIDATION_ERROR", str(e)))


# ── Hyperlinks ────────────────────────────────────────────────────────────────

def add_hyperlink(sheet: str, range_address: str, url: str, display_text: str = None) -> str:
    """Inserts a clickable hyperlink into a cell."""
    try:
        validate_sheet_name(sheet)
        book = get_active_book()
        s = get_sheet(book, sheet)
        cell = s.range(range_address)
        display = display_text if display_text else url

        with excel_suppressed(book.app):
            s.api.Hyperlinks.Add(Anchor=cell.api, Address=url, ScreenTip=url, TextToDisplay=display)

        return result_to_str(ok(f"Hyperlink '{display}' → '{url}' added to '{sheet}'!{range_address}."))
    except Exception as e:
        logger.exception("add_hyperlink failed")
        return result_to_str(err("ADD_HYPERLINK_ERROR", str(e)))


# ── Fill / Auto-Fill ──────────────────────────────────────────────────────────

def fill_range(sheet: str, range_address: str, direction: str = "down") -> str:
    """Auto-fills formulas or series across a range."""
    try:
        validate_sheet_name(sheet)
        range_address = validate_range_address(range_address)
        book = get_active_book()
        s = get_sheet(book, sheet)

        with excel_suppressed(book.app):
            r = s.range(range_address)
            d = direction.lower()
            if d == "down":
                r.api.FillDown()
            elif d == "up":
                r.api.FillUp()
            elif d == "right":
                r.api.FillRight()
            elif d == "left":
                r.api.FillLeft()
            else:
                return result_to_str(err("INVALID_INPUT", f"Unknown direction '{direction}'. Use: down, up, right, left."))

        return result_to_str(ok(f"Auto-filled '{sheet}'!{range_address} ({direction})."))
    except Exception as e:
        logger.exception("fill_range failed")
        return result_to_str(err("FILL_RANGE_ERROR", str(e)))
