"""
rules_evaluator.py — Phase 9
Safe Business-Rule Evaluator:
  - Evaluates client-supplied conditional rules (e.g. "variance_percent > 10")
  - Uses an AST allowlist parser — strictly denies eval(), os.system, or code injection
  - Performs safe math and string comparison operations
"""

from __future__ import annotations
import ast
import operator as op
from typing import Any, Dict

# Allowed operators
_SAFE_OPERATORS = {
    ast.Add: op.add,
    ast.Sub: op.sub,
    ast.Mult: op.mul,
    ast.Div: op.truediv,
    ast.Mod: op.mod,
    ast.Pow: op.pow,
    ast.Eq: op.eq,
    ast.NotEq: op.ne,
    ast.Lt: op.lt,
    ast.LtE: op.le,
    ast.Gt: op.gt,
    ast.GtE: op.ge,
    ast.And: lambda a, b: a and b,
    ast.Or: lambda a, b: a or b,
    ast.Not: op.not_,
}


class RuleEvaluationError(Exception):
    pass


def evaluate_expression(expression: str, context: Dict[str, Any]) -> Any:
    """
    Safely evaluate a mathematical/logical condition expression against a dictionary context.

    Example:
      evaluate_expression("revenue > 10000 and status == 'Active'", {"revenue": 15000, "status": "Active"})
      --> Returns True

    Security:
      Rejects arbitrary code, attribute calls, module imports, or statements.
    """
    if not expression or not isinstance(expression, str):
        raise RuleEvaluationError("Expression must be a non-empty string.")

    try:
        parsed_ast = ast.parse(expression, mode='eval')
    except Exception as e:
        raise RuleEvaluationError(f"Invalid expression syntax: {e}")

    return _eval_node(parsed_ast.body, context)


def _eval_node(node: ast.AST, context: Dict[str, Any]) -> Any:
    if isinstance(node, ast.Constant):
        return node.value

    elif isinstance(node, ast.Name):
        if node.id in context:
            return context[node.id]
        elif node.id.lower() in ("true", "false", "none"):
            return {"true": True, "false": False, "none": None}[node.id.lower()]
        raise RuleEvaluationError(f"Undefined variable in condition: '{node.id}'")

    elif isinstance(node, ast.UnaryOp):
        op_type = type(node.op)
        if op_type in _SAFE_OPERATORS:
            return _SAFE_OPERATORS[op_type](_eval_node(node.operand, context))
        raise RuleEvaluationError(f"Unsupported unary operator: {op_type.__name__}")

    elif isinstance(node, ast.BinOp):
        op_type = type(node.op)
        if op_type in _SAFE_OPERATORS:
            left = _eval_node(node.left, context)
            right = _eval_node(node.right, context)
            return _SAFE_OPERATORS[op_type](left, right)
        raise RuleEvaluationError(f"Unsupported binary operator: {op_type.__name__}")

    elif isinstance(node, ast.Compare):
        left = _eval_node(node.left, context)
        for comparator, op_ast in zip(node.comparators, node.ops):
            op_type = type(op_ast)
            if op_type not in _SAFE_OPERATORS:
                raise RuleEvaluationError(f"Unsupported comparison operator: {op_type.__name__}")
            right = _eval_node(comparator, context)
            if not _SAFE_OPERATORS[op_type](left, right):
                return False
            left = right
        return True

    elif isinstance(node, ast.BoolOp):
        op_type = type(node.op)
        if op_type == ast.And:
            return all(_eval_node(val, context) for val in node.values)
        elif op_type == ast.Or:
            return any(_eval_node(val, context) for val in node.values)
        raise RuleEvaluationError(f"Unsupported boolean operator: {op_type.__name__}")

    else:
        raise RuleEvaluationError(f"Security Rejection: Expression node type '{type(node).__name__}' is forbidden.")
