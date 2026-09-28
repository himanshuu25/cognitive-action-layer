"""
tests/test_workflow_engine.py — Enterprise Workflow Engine Test Suite
Tests:
  - Canonical response formatting
  - AST Expression Evaluator (Phase 9)
  - DAG topological sorting and cycle detection (Phase 5)
  - Variable template resolution (Phase 4)
"""

import pytest
import sys
import os

sys.path.insert(0, os.path.dirname(os.path.dirname(__file__)))

from shared_context import ok_canonical, err_canonical, canonical_response
from rules_evaluator import evaluate_expression, RuleEvaluationError
from workflow_engine import _topological_sort, _resolve_variables


class TestCanonicalContracts:
    def test_ok_canonical(self):
        res = ok_canonical(
            operation="remove_duplicates",
            workbook_id="wb_123",
            sheet="Sales",
            result={"rows_removed": 43},
        )
        assert res["status"] == "success"
        assert res["operation"] == "remove_duplicates"
        assert res["workbook_id"] == "wb_123"
        assert res["result"]["rows_removed"] == 43
        assert res["verification"]["passed"] is True

    def test_err_canonical(self):
        res = err_canonical(
            operation="write_table",
            code="RANGE_COLLISION",
            message="Target range is protected.",
            workbook_id="wb_123",
        )
        assert res["status"] == "failed"
        assert res["error"]["code"] == "RANGE_COLLISION"
        assert res["verification"]["passed"] is False


class TestRulesEvaluator:
    def test_simple_comparison(self):
        assert evaluate_expression("revenue > 10000", {"revenue": 15000}) is True
        assert evaluate_expression("revenue <= 10000", {"revenue": 15000}) is False

    def test_logical_and_or(self):
        ctx = {"variance": 12.5, "status": "Active"}
        assert evaluate_expression("variance > 10 and status == 'Active'", ctx) is True
        assert evaluate_expression("variance > 20 or status == 'Active'", ctx) is True

    def test_security_rejection_import(self):
        with pytest.raises(RuleEvaluationError, match="Security Rejection|Invalid expression"):
            evaluate_expression("__import__('os').system('dir')", {})

    def test_security_rejection_eval(self):
        with pytest.raises(RuleEvaluationError, match="Security Rejection|Invalid expression"):
            evaluate_expression("eval('1+1')", {})


class TestWorkflowDAG:
    def test_topological_sort_valid(self):
        ops = [
            {"id": "step_3", "operation": "sort_range", "depends_on": ["step_2"]},
            {"id": "step_1", "operation": "open_workbook"},
            {"id": "step_2", "operation": "clean_data", "depends_on": ["step_1"]},
        ]
        sorted_ops = _topological_sort(ops)
        ids = [op["id"] for op in sorted_ops]
        assert ids == ["step_1", "step_2", "step_3"]

    def test_topological_sort_cycle_detected(self):
        ops = [
            {"id": "step_1", "operation": "clean_data", "depends_on": ["step_2"]},
            {"id": "step_2", "operation": "sort_range", "depends_on": ["step_1"]},
        ]
        with pytest.raises(ValueError, match="Cyclic dependency"):
            _topological_sort(ops)

    def test_variable_resolution(self):
        step_outputs = {
            "step_1": {"result": {"range": "A1:D100", "rows": 100}},
            "step_2": {"data": {"col_name": "Revenue"}},
        }
        params = {
            "target_range": "{{step_1.result.range}}",
            "column": "{{step_2.data.col_name}}",
            "static": "value",
        }
        resolved = _resolve_variables(params, step_outputs)
        assert resolved["target_range"] == "A1:D100"
        assert resolved["column"] == "Revenue"
        assert resolved["static"] == "value"


if __name__ == "__main__":
    pytest.main([__file__, "-v"])
