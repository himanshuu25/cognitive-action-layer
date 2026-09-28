"""
control_totals.py — Phase 6
Pluggable Control Totals & Verification Pipeline:
  - Generates structural, data, and numeric control totals for any sheet/range.
  - Pluggable checks (row_count, numeric_sum, cell_hash, null_count).
  - Mandatory verification pipeline runner.
"""

from __future__ import annotations
import math
import hashlib
from typing import List, Dict, Any, Optional

import xlwings as xw
from logger import logger, log_audit
from shared_context import canonical_response, ok_canonical, err_canonical


class ControlTotalCheck:
    """Represents a single control total check configuration."""
    def __init__(
        self,
        check_type: str,            # "row_count" | "numeric_sum" | "null_count" | "hash"
        target_column: Optional[str] = None,
        expected_value: Optional[Any] = None,
        tolerance: float = 1e-6,
    ):
        self.check_type = check_type
        self.target_column = target_column
        self.expected_value = expected_value
        self.tolerance = tolerance


def compute_range_control_totals(sheet, range_address: str) -> dict:
    """
    Compute comprehensive control metrics for a given sheet range:
      - total_cells, row_count, col_count
      - numeric_sum, numeric_count
      - null_count, empty_string_count
      - data_hash (SHA-256 of contents)
    """
    try:
        rng = sheet.range(range_address)
        values = rng.value
    except Exception as e:
        return {"error": f"Could not read range '{range_address}': {e}"}

    if values is None:
        values = []
    elif not isinstance(values, list):
        values = [[values]]
    elif values and not isinstance(values[0], list):
        values = [values]

    row_count = len(values)
    col_count = max((len(r) for r in values if isinstance(r, list)), default=0)
    total_cells = row_count * col_count

    numeric_sum = 0.0
    numeric_count = 0
    null_count = 0

    flat_str_buf = []

    for row in values:
        if not isinstance(row, list):
            row = [row]
        for cell in row:
            flat_str_buf.append(str(cell) if cell is not None else "")
            if cell is None or cell == "":
                null_count += 1
                continue
            if isinstance(cell, (int, float)):
                if not math.isnan(cell) and not math.isinf(cell):
                    numeric_sum += float(cell)
                    numeric_count += 1
            elif isinstance(cell, str):
                try:
                    val = float(cell.replace(",", ""))
                    numeric_sum += val
                    numeric_count += 1
                except ValueError:
                    pass

    data_hash = hashlib.sha256("|".join(flat_str_buf).encode("utf-8")).hexdigest()

    return {
        "range": range_address,
        "row_count": row_count,
        "col_count": col_count,
        "total_cells": total_cells,
        "numeric_sum": round(numeric_sum, 6),
        "numeric_count": numeric_count,
        "null_count": null_count,
        "data_hash": data_hash,
    }


def run_control_total_pipeline(
    sheet,
    range_address: str,
    checks: List[dict],
) -> dict:
    """
    Executes a list of pluggable control-total assertions (§3 Phase 6).

    Check schema:
      { "type": "row_count", "expected": 100, "tolerance": 0 }
      { "type": "numeric_sum", "column": "Amount", "expected": 150000.50, "tolerance": 0.01 }
      { "type": "null_count", "expected": 0 }
    """
    totals = compute_range_control_totals(sheet, range_address)
    if "error" in totals:
        return {"passed": False, "error": totals["error"], "checks": []}

    passed_all = True
    check_results = []

    for c in checks:
        ctype = c.get("type")
        expected = c.get("expected")
        tolerance = c.get("tolerance", 1e-6)

        check_passed = True
        actual = None
        message = ""

        if ctype == "row_count":
            actual = totals["row_count"]
            if expected is not None:
                check_passed = abs(actual - expected) <= tolerance
            message = f"Row count actual={actual}, expected={expected}"

        elif ctype == "numeric_sum":
            actual = totals["numeric_sum"]
            if expected is not None:
                check_passed = abs(actual - float(expected)) <= tolerance
            message = f"Numeric sum actual={actual}, expected={expected} (Δ={abs(actual - float(expected or 0)):.6f})"

        elif ctype == "null_count":
            actual = totals["null_count"]
            if expected is not None:
                check_passed = actual <= expected
            message = f"Null count actual={actual}, max_allowed={expected}"

        elif ctype == "hash":
            actual = totals["data_hash"]
            if expected is not None:
                check_passed = actual == expected
            message = f"Data hash match={check_passed}"

        if not check_passed:
            passed_all = False

        check_results.append({
            "check_type": ctype,
            "passed": check_passed,
            "actual": actual,
            "expected": expected,
            "message": message,
        })

    logger.info(
        "Control totals verification for '%s': passed=%s (%d checks)",
        range_address, passed_all, len(check_results)
    )

    return {
        "passed": passed_all,
        "metrics": totals,
        "checks": check_results,
    }
