"""
workbook_tools.py — Production-hardened
Workbook and sheet management with:
  - Workbook-by-name targeting (no more "active book" ambiguity)
  - Path sandboxing on open/save/export
  - Macro-enabled format warnings
  - Structured error returns
  - Secret redaction in protect_sheet
"""

from __future__ import annotations
import os
import xlwings as xw

from shared_context import (
    validate_sheet_name,
    ok, err, result_to_str,
)
from security import validate_path, warn_macro_enabled, audit_write
from logger import logger


# ── Active Book Resolution ────────────────────────────────────────────────────

def get_active_book(workbook_name: str = None):
    """
    Return the target xlwings Book.

    If workbook_name is provided, search all open books for that name (case-insensitive).
    If not found, raise ValueError — never silently fall back to a wrong workbook.

    If workbook_name is None, return the currently active book (legacy behaviour,
    safe when only one workbook is open).
    """
    if workbook_name:
        target = workbook_name.lower()
        for book in xw.books:
            if (
                os.path.basename(book.fullname).lower() == target
                or book.name.lower() == target
            ):
                return book
        raise ValueError(
            f"Workbook '{workbook_name}' is not open. "
            f"Open books: {[b.name for b in xw.books]}"
        )

    # Legacy path: return active book, launch Excel if needed
    if len(xw.apps) == 0:
        app = xw.App(visible=True, add_book=False)
        return xw.Book()
    if len(xw.books) == 0:
        return xw.Book()
    return xw.books.active


def get_sheet(book, sheet_name: str):
    """Find a sheet by case-insensitive name. Raises ValueError if not found."""
    for s in book.sheets:
        if s.name.lower() == sheet_name.lower():
            return s
    available = [s.name for s in book.sheets]
    raise ValueError(f"Sheet '{sheet_name}' not found. Available: {available}")


# ── Workbook Operations ───────────────────────────────────────────────────────

def open_workbook(file_path: str) -> str:
    """Opens an existing Excel workbook. Validates path and warns on macro-enabled formats."""
    try:
        # Validate path
        try:
            abs_path = validate_path(file_path, operation="open")
        except ValueError as e:
            return result_to_str(err("PATH_DENIED", str(e)))

        if not os.path.exists(abs_path):
            return result_to_str(err("FILE_NOT_FOUND", f"File '{file_path}' does not exist."))

        # Macro-enabled warning
        macro_warning = warn_macro_enabled(abs_path)

        # Check if already open
        for book in xw.books:
            try:
                if os.path.abspath(book.fullname).lower() == abs_path.lower():
                    book.activate()
                    msg = f"Workbook '{os.path.basename(abs_path)}' is already open and focused."
                    if macro_warning:
                        msg += f"\n{macro_warning}"
                    return result_to_str(ok(msg))
            except Exception:
                pass

        # Ensure Excel is running
        if len(xw.apps) == 0:
            xw.App(visible=True, add_book=False)

        xw.Book(abs_path)
        msg = f"Opened '{os.path.basename(abs_path)}'."
        if macro_warning:
            msg += f"\n{macro_warning}"

        audit_write("OPEN_WORKBOOK", "open_workbook", {"path": abs_path}, result_status="ok")
        return result_to_str(ok(msg, data={"path": abs_path}))

    except Exception as e:
        logger.exception("open_workbook failed")
        return result_to_str(err("OPEN_ERROR", str(e)))


def create_workbook(file_path: str = None) -> str:
    """Creates a new blank workbook. Optionally saves it to a validated path."""
    try:
        save_path = None
        if file_path:
            try:
                save_path = validate_path(file_path, operation="create")
            except ValueError as e:
                return result_to_str(err("PATH_DENIED", str(e)))

        if len(xw.apps) == 0:
            xw.App(visible=True, add_book=False)

        book = xw.Book()
        if save_path:
            book.save(save_path)
            audit_write("CREATE_WORKBOOK", "create_workbook", {"path": save_path}, result_status="ok")
            return result_to_str(ok(f"Created and saved workbook at '{save_path}'.", data={"path": save_path}))

        return result_to_str(ok("Created a new blank workbook."))
    except Exception as e:
        logger.exception("create_workbook failed")
        return result_to_str(err("CREATE_ERROR", str(e)))


def save_workbook(file_path: str = None) -> str:
    """Saves the active workbook. Validates path if Save As is requested."""
    try:
        book = get_active_book()

        if file_path:
            try:
                abs_path = validate_path(file_path, operation="save")
            except ValueError as e:
                return result_to_str(err("PATH_DENIED", str(e)))

            # Protect against accidental overwrites of different files
            if os.path.exists(abs_path):
                try:
                    current = os.path.abspath(book.fullname)
                    if current.lower() != abs_path.lower():
                        return result_to_str(err(
                            "OVERWRITE_DENIED",
                            f"'{abs_path}' already exists and is a different file. "
                            "Use a new filename or open the target workbook first.",
                        ))
                except Exception:
                    pass

            book.save(abs_path)
            audit_write("SAVE_WORKBOOK", "save_workbook", {"path": abs_path}, result_status="ok")
            return result_to_str(ok(f"Saved as '{abs_path}'."))
        else:
            book.save()
            audit_write("SAVE_WORKBOOK", "save_workbook", {"path": book.fullname}, result_status="ok")
            return result_to_str(ok(f"Saved '{book.name}'."))

    except Exception as e:
        logger.exception("save_workbook failed")
        return result_to_str(err("SAVE_ERROR", str(e)))


def list_sheets() -> list[str]:
    """Lists the names of all worksheets in the active workbook."""
    try:
        book = get_active_book()
        return [sheet.name for sheet in book.sheets]
    except Exception as e:
        return [result_to_str(err("LIST_SHEETS_ERROR", str(e)))]


def add_sheet(name: str) -> str:
    """Adds a new worksheet to the active workbook."""
    try:
        validate_sheet_name(name)
        book = get_active_book()
        for sheet in book.sheets:
            if sheet.name.lower() == name.lower():
                return result_to_str(ok(f"Sheet '{name}' already exists."))
        book.sheets.add(name=name)
        return result_to_str(ok(f"Added sheet '{name}'."))
    except Exception as e:
        logger.exception("add_sheet failed")
        return result_to_str(err("ADD_SHEET_ERROR", str(e)))


def rename_sheet(sheet_name: str, new_name: str) -> str:
    """Renames a worksheet."""
    try:
        validate_sheet_name(new_name)
        book = get_active_book()
        s = get_sheet(book, sheet_name)
        s.name = new_name
        return result_to_str(ok(f"Renamed '{sheet_name}' → '{new_name}'."))
    except Exception as e:
        logger.exception("rename_sheet failed")
        return result_to_str(err("RENAME_SHEET_ERROR", str(e)))


def delete_sheet(sheet_name: str) -> str:
    """Deletes a worksheet. Refuses to delete the last remaining sheet."""
    try:
        book = get_active_book()
        if len(book.sheets) <= 1:
            return result_to_str(err(
                "LAST_SHEET",
                "Cannot delete the only remaining worksheet."
            ))
        s = get_sheet(book, sheet_name)
        s.delete()
        return result_to_str(ok(f"Deleted sheet '{sheet_name}'."))
    except Exception as e:
        logger.exception("delete_sheet failed")
        return result_to_str(err("DELETE_SHEET_ERROR", str(e)))


def export_to_pdf(pdf_path: str, sheet_name: str = None) -> str:
    """Exports the workbook or a specific sheet to PDF. Validates output path."""
    try:
        try:
            abs_path = validate_path(pdf_path, operation="export")
        except ValueError as e:
            return result_to_str(err("PATH_DENIED", str(e)))

        book = get_active_book()
        dir_name = os.path.dirname(abs_path)
        if dir_name:
            os.makedirs(dir_name, exist_ok=True)

        if sheet_name:
            s = get_sheet(book, sheet_name)
            s.to_pdf(abs_path)
            audit_write("EXPORT_PDF", "export_to_pdf", {"path": abs_path, "sheet": sheet_name}, result_status="ok")
            return result_to_str(ok(f"Exported sheet '{sheet_name}' to PDF: '{abs_path}'."))
        else:
            book.to_pdf(abs_path)
            audit_write("EXPORT_PDF", "export_to_pdf", {"path": abs_path}, result_status="ok")
            return result_to_str(ok(f"Exported workbook to PDF: '{abs_path}'."))

    except Exception as e:
        logger.exception("export_to_pdf failed")
        return result_to_str(err("EXPORT_PDF_ERROR", str(e)))


def protect_sheet(sheet_name: str, password: str = None, protect: bool = True) -> str:
    """
    Protects or unprotects a worksheet.
    Passwords are NEVER logged (redacted in audit trail).
    Note: Excel sheet protection is weak — do not treat it as a security boundary.
    """
    try:
        book = get_active_book()
        s = get_sheet(book, sheet_name)

        if protect:
            s.api.Protect(Password=password)
            action = "protected"
        else:
            s.api.Unprotect(Password=password)
            action = "unprotected"

        # Log WITHOUT the password
        audit_write(
            "PROTECT_SHEET", "protect_sheet",
            {"sheet": sheet_name, "protect": protect, "password": "***REDACTED***"},
            result_status="ok"
        )
        return result_to_str(ok(f"Sheet '{sheet_name}' {action}."))
    except Exception as e:
        logger.exception("protect_sheet failed")
        return result_to_str(err("PROTECT_SHEET_ERROR", str(e)))


def set_sheet_visibility(sheet_name: str, visible: bool = True) -> str:
    """Hides or shows a worksheet."""
    try:
        book = get_active_book()
        s = get_sheet(book, sheet_name)
        s.api.Visible = -1 if visible else 0
        action = "visible" if visible else "hidden"
        return result_to_str(ok(f"Sheet '{sheet_name}' set to '{action}'."))
    except Exception as e:
        logger.exception("set_sheet_visibility failed")
        return result_to_str(err("VISIBILITY_ERROR", str(e)))


def list_open_workbooks() -> str:
    """Lists all currently open workbooks so the correct target can be identified."""
    try:
        books = []
        for b in xw.books:
            try:
                books.append({
                    "name": b.name,
                    "path": b.fullname,
                    "active": b == xw.books.active,
                })
            except Exception:
                pass
        return result_to_str(ok(f"Found {len(books)} open workbook(s).", data=books))
    except Exception as e:
        return result_to_str(err("LIST_WORKBOOKS_ERROR", str(e)))
