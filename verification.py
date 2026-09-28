"""
verification.py — Phase 4
Write-and-verify, reconciliation, formula error scanning, and schema contracts.
"""

from __future__ import annotations
import json
import math
from decimal import Decimal, ROUND_HALF_UP
from typing import Any, List, Optional

import xlwings as xw

from logger import logger, log_audit

# Excel error cell values (what xlwings returns for error cells)
_EXCEL_ERRORS = {
    -2146826281: "#DIV/0!",
    -2146826246: "#N/A",
    -2146826259: "#NAME?",
    -2146826288: "#NULL!",
    -2146826252: "#NUM!",
    -2146826265: "#REF!",
    -2146826273: "#VALUE!",
}


def _is_error_value(v: Any) -> bool:
    return isinstance(v, int) and v in _EXCEL_ERRORS


def _error_name(v: Any) -> str:
    return _EXCEL_ERRORS.get(v, str(v))


# ── Write-and-Verify ──────────────────────────────────────────────────────────

def verify_write(
    sheet,
    range_address: str,
    intended_data: list,
    tolerance: float = 1e-9,
) -> dict:
    """
    Read back the written range and compare element-by-element against intended_data.

    Returns:
        {
            "verified": bool,
            "written_rows": int,
            "written_cols": int,
            "mismatches": [ { "row": r, "col": c, "expected": ..., "actual": ... } ],
            "error_cells": [ { "row": r, "col": c, "error": "#REF!" } ]
        }
    """
    try:
        actual = sheet.range(range_address).value
    except Exception as e:
        return {
            "verified": False,
            "error": f"Could not read back range '{range_address}': {e}",
        }

    # Normalize to 2D
    if actual is None:
        actual = []
    elif not isinstance(actual, list):
        actual = [[actual]]
    elif actual and not isinstance(actual[0], list):
        actual = [actual]

    if not isinstance(intended_data, list):
        intended_data = [[intended_data]]
    elif intended_data and not isinstance(intended_data[0], list):
        intended_data = [intended_data]

    mismatches = []
    error_cells = []

    for r_idx, intended_row in enumerate(intended_data):
        actual_row = actual[r_idx] if r_idx < len(actual) else []
        for c_idx, expected_val in enumerate(intended_row):
            actual_val = actual_row[c_idx] if c_idx < len(actual_row) else None

            # Check for Excel error values
            if _is_error_value(actual_val):
                error_cells.append({
                    "row": r_idx, "col": c_idx, "error": _error_name(actual_val)
                })
                continue

            # Type-aware comparison
            if not _values_equal(expected_val, actual_val, tolerance):
                mismatches.append({
                    "row": r_idx, "col": c_idx,
                    "expected": str(expected_val),
                    "actual": str(actual_val),
                })

    rows = len(actual)
    cols = max((len(r) for r in actual if isinstance(r, list)), default=0)

    verified = len(mismatches) == 0 and len(error_cells) == 0

    result = {
        "verified": verified,
        "written_rows": rows,
        "written_cols": cols,
        "mismatches": mismatches,
        "error_cells": error_cells,
    }

    if not verified:
        logger.warning(
            "Write verification FAILED: %d mismatches, %d error cells in range '%s'.",
            len(mismatches), len(error_cells), range_address
        )
    else:
        logger.info("Write verification PASSED for range '%s' (%d×%d).", range_address, rows, cols)

    return result


def _values_equal(expected: Any, actual: Any, tolerance: float = 1e-9) -> bool:
    """Type-aware equality check with float tolerance."""
    if expected is None and actual is None:
        return True
    if expected is None or actual is None:
        # Allow empty string == None
        if expected == "" or actual == "":
            return True
        return False

    # Both numeric?
    if isinstance(expected, (int, float)) and isinstance(actual, (int, float)):
        if math.isnan(expected) and math.isnan(actual):
            return True
        return abs(expected - actual) <= tolerance

    # String comparison (case-insensitive for text)
    if isinstance(expected, str) and isinstance(actual, str):
        return expected.strip() == actual.strip()

    # Fallback
    return str(expected).strip() == str(actual).strip()


# ── Reconciliation (Control Totals) ──────────────────────────────────────────

def reconcile_numeric(
    source_data: list,
    sheet,
    range_address: str,
    tolerance: float = 1e-6,
) -> dict:
    """
    Compare control totals (sum + count of numeric values) between source_data
    and what was actually written to the sheet range.

    Returns:
        {
            "reconciled": bool,
            "source_sum": float,
            "dest_sum": float,
            "source_count": int,
            "dest_count": int,
            "delta_sum": float,
            "delta_count": int,
        }
    """
    def _extract_numerics(data):
        total = 0.0
        count = 0
        flat = []
        if isinstance(data, list):
            for row in data:
                if isinstance(row, list):
                    flat.extend(row)
                elif isinstance(row, dict):
                    flat.extend(row.values())
                else:
                    flat.append(row)
        else:
            flat = [data]
        for v in flat:
            try:
                f = float(v)
                if not math.isnan(f) and not math.isinf(f):
                    total += f
                    count += 1
            except (TypeError, ValueError):
                pass
        return total, count

    src_sum, src_count = _extract_numerics(source_data)

    try:
        actual = sheet.range(range_address).value
    except Exception as e:
        return {"reconciled": False, "error": f"Cannot read destination: {e}"}

    dest_sum, dest_count = _extract_numerics(actual)

    delta_sum = abs(src_sum - dest_sum)
    delta_count = abs(src_count - dest_count)
    reconciled = delta_sum <= tolerance and delta_count == 0

    result = {
        "reconciled": reconciled,
        "source_sum": src_sum,
        "dest_sum": dest_sum,
        "source_count": src_count,
        "dest_count": dest_count,
        "delta_sum": delta_sum,
        "delta_count": delta_count,
    }

    if reconciled:
        logger.info(
            "Reconciliation PASSED: sum=%.6f count=%d in '%s'.",
            dest_sum, dest_count, range_address
        )
    else:
        logger.warning(
            "Reconciliation FAILED for '%s': Δsum=%.6f Δcount=%d.",
            range_address, delta_sum, delta_count
        )

    return result


# ── Formula Verification & Error Scanning ────────────────────────────────────

def scan_formula_errors(sheet, range_address: str) -> dict:
    """
    Scan a range for Excel error values (#REF!, #DIV/0!, etc.).
    Also forces a recalculation before scanning.

    Returns:
        {
            "clean": bool,
            "error_cells": [ { "address": "B5", "error": "#REF!" } ]
        }
    """
    try:
        # Force recalculation
        try:
            sheet.book.app.api.Calculate()
        except Exception:
            pass

        rng = sheet.range(range_address)
        values = rng.value

        if values is None:
            return {"clean": True, "error_cells": []}

        # Normalize to 2D
        if not isinstance(values, list):
            values = [[values]]
        elif values and not isinstance(values[0], list):
            values = [values]

        top_row = rng.row
        left_col = rng.column

        error_cells = []
        for r_idx, row in enumerate(values):
            if not isinstance(row, list):
                row = [row]
            for c_idx, cell_val in enumerate(row):
                if _is_error_value(cell_val):
                    # Build cell address
                    col_letter = _col_num_to_letter(left_col + c_idx)
                    cell_addr = f"{col_letter}{top_row + r_idx}"
                    error_cells.append({
                        "address": cell_addr,
                        "error": _error_name(cell_val),
                    })

        clean = len(error_cells) == 0
        if not clean:
            logger.warning(
                "Formula error scan found %d errors in '%s': %s",
                len(error_cells), range_address,
                [e["error"] for e in error_cells[:5]]
            )

        return {"clean": clean, "error_cells": error_cells}

    except Exception as e:
        return {"clean": False, "error": str(e)}


def _col_num_to_letter(n: int) -> str:
    """Convert 1-based column number to Excel column letter (1=A, 26=Z, 27=AA…)."""
    result = ""
    while n > 0:
        n, rem = divmod(n - 1, 26)
        result = chr(65 + rem) + result
    return result


# ── Schema Contracts ──────────────────────────────────────────────────────────

class WriteTableSchema:
    """
    Pre/post-condition schema for a table write operation.
    Validates column names, expected dtypes, nullable constraints, and row counts.
    """

    def __init__(
        self,
        columns: List[str],
        dtypes: Optional[dict] = None,          # {"col_name": "int"|"float"|"str"|"date"}
        nullable: Optional[dict] = None,         # {"col_name": True/False}
        min_rows: Optional[int] = None,
        max_rows: Optional[int] = None,
    ):
        self.columns = columns
        self.dtypes = dtypes or {}
        self.nullable = nullable or {}
        self.min_rows = min_rows
        self.max_rows = max_rows

    def validate_pre(self, data: list) -> dict:
        """
        Validate data before writing.
        Returns {"valid": bool, "violations": [...]}
        """
        violations = []

        if not data:
            return {"valid": False, "violations": ["Data is empty."]}

        # Get headers from first row
        first_row = data[0]
        if isinstance(first_row, dict):
            headers = list(first_row.keys())
            rows = data
        elif isinstance(first_row, list):
            headers = first_row
            rows = data[1:]
        else:
            return {"valid": False, "violations": [f"Unexpected data format: {type(first_row)}"]}

        # Check column presence
        missing_cols = [c for c in self.columns if c not in headers]
        if missing_cols:
            violations.append(f"Missing required columns: {missing_cols}")

        # Check row count
        row_count = len(rows)
        if self.min_rows is not None and row_count < self.min_rows:
            violations.append(
                f"Too few rows: expected >= {self.min_rows}, got {row_count}."
            )
        if self.max_rows is not None and row_count > self.max_rows:
            violations.append(
                f"Too many rows: expected <= {self.max_rows}, got {row_count}."
            )

        # Nullable checks
        for col, must_have_value in self.nullable.items():
            if not must_have_value and col in headers:
                col_idx = headers.index(col)
                for i, row in enumerate(rows):
                    if isinstance(row, dict):
                        val = row.get(col)
                    elif isinstance(row, list):
                        val = row[col_idx] if col_idx < len(row) else None
                    else:
                        val = None
                    if val is None or val == "":
                        violations.append(
                            f"Column '{col}' has null/empty value in row {i + 1} (nullable=False)."
                        )
                        break  # Report first violation per column

        valid = len(violations) == 0
        return {"valid": valid, "violations": violations, "row_count": len(rows), "col_count": len(headers)}

    def validate_post(self, sheet, range_address: str) -> dict:
        """
        Validate the written range shape matches schema expectations.
        """
        try:
            rng = sheet.range(range_address).expand("table")
            actual_rows = rng.rows.count - 1  # Exclude header
            actual_cols = rng.columns.count
        except Exception as e:
            return {"valid": False, "violations": [f"Cannot read back range: {e}"]}

        violations = []
        if len(self.columns) > 0 and actual_cols < len(self.columns):
            violations.append(
                f"Post-write column count {actual_cols} < expected {len(self.columns)}."
            )
        if self.min_rows is not None and actual_rows < self.min_rows:
            violations.append(
                f"Post-write row count {actual_rows} < expected min {self.min_rows}."
            )
        if self.max_rows is not None and actual_rows > self.max_rows:
            violations.append(
                f"Post-write row count {actual_rows} > expected max {self.max_rows}."
            )

        return {
            "valid": len(violations) == 0,
            "violations": violations,
            "actual_rows": actual_rows,
            "actual_cols": actual_cols,
        }
