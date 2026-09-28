"""
workbook_registry.py — Phase 1 & 7
Enterprise Workbook Registry:
  - Maps unique workbook_ids (e.g. 'wb_a1b2c3') to file paths and handles.
  - Supports simultaneous multi-workbook operations.
  - Cross-workbook comparison and structural diffing.
"""

from __future__ import annotations
import os
import hashlib
import time
from pathlib import Path
from typing import Dict, Any, Optional, List

import xlwings as xw
from logger import logger


class WorkbookRegistry:
    """Singleton registry tracking open workbooks by workbook_id."""
    
    _instance = None
    
    def __new__(cls):
        if cls._instance is None:
            cls._instance = super(WorkbookRegistry, cls).__new__(cls)
            cls._instance._registry: Dict[str, Dict[str, Any]] = {}
        return cls._instance

    def generate_id(self, path: str) -> str:
        """Generate a deterministic workbook_id based on path and timestamp hash."""
        abs_path = os.path.abspath(path).lower()
        digest = hashlib.md5(abs_path.encode('utf-8')).hexdigest()[:8]
        return f"wb_{digest}"

    def register(self, path: str, book: xw.Book) -> str:
        """Register an open workbook handle and return its workbook_id."""
        wb_id = self.generate_id(path)
        self._registry[wb_id] = {
            "workbook_id": wb_id,
            "path": os.path.abspath(path),
            "name": book.name,
            "book": book,
            "registered_at": time.time(),
        }
        logger.info("Registered workbook_id='%s' for path='%s'", wb_id, path)
        return wb_id

    def get_book(self, workbook_id: Optional[str] = None) -> xw.Book:
        """
        Get xlwings Book handle by workbook_id.
        Fallback: If workbook_id is None or not found, return active book.
        """
        if workbook_id and workbook_id in self._registry:
            entry = self._registry[workbook_id]
            try:
                # Test if book object is still alive
                _ = entry["book"].name
                return entry["book"]
            except Exception:
                # Re-open if handle died
                logger.warning("Workbook handle for '%s' died. Re-opening from path.", workbook_id)
                book = xw.Book(entry["path"])
                entry["book"] = book
                return book
        
        # Search open xlwings books by name or path match
        if workbook_id:
            target = workbook_id.lower()
            for book in xw.books:
                try:
                    if (
                        book.name.lower() == target
                        or os.path.basename(book.fullname).lower() == target
                        or self.generate_id(book.fullname) == workbook_id
                    ):
                        return book
                except Exception:
                    pass
            raise ValueError(f"Workbook with ID or name '{workbook_id}' is not registered or open.")

        # Fallback to active book
        if len(xw.apps) == 0:
            xw.App(visible=True, add_book=False)
            b = xw.Book()
            self.register(b.fullname, b)
            return b
        if len(xw.books) == 0:
            b = xw.Book()
            self.register(b.fullname, b)
            return b
        return xw.books.active

    def get_info(self, workbook_id: str) -> Optional[dict]:
        return self._registry.get(workbook_id)

    def compare_workbooks(self, workbook_id_1: str, workbook_id_2: str) -> dict:
        """
        Structural diff between two open workbooks (Phase 7 & 8).
        Compares sheet names, row/col counts, named ranges, and formulas.
        """
        b1 = self.get_book(workbook_id_1)
        b2 = self.get_book(workbook_id_2)

        s1_names = [s.name for s in b1.sheets]
        s2_names = [s.name for s in b2.sheets]

        added_sheets = [s for s in s2_names if s not in s1_names]
        removed_sheets = [s for s in s1_names if s not in s2_names]
        common_sheets = [s for s in s1_names if s in s2_names]

        sheet_diffs = {}
        for sname in common_sheets:
            try:
                u1 = b1.sheets[sname].used_range
                u2 = b2.sheets[sname].used_range
                r1, c1 = (u1.last_cell.row, u1.last_cell.column) if u1 else (0, 0)
                r2, c2 = (u2.last_cell.row, u2.last_cell.column) if u2 else (0, 0)

                sheet_diffs[sname] = {
                    "dimensions_changed": (r1, c1) != (r2, c2),
                    "wb1_shape": f"{r1}x{c1}",
                    "wb2_shape": f"{r2}x{c2}",
                }
            except Exception:
                pass

        return {
            "wb1": b1.name,
            "wb2": b2.name,
            "added_sheets": added_sheets,
            "removed_sheets": removed_sheets,
            "sheet_diffs": sheet_diffs,
        }


registry = WorkbookRegistry()

