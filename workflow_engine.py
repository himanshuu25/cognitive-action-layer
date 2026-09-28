"""
workflow_engine.py — Phase 2, 3, 4, 5
Durable Multi-Step Workflow Engine:
  - Executes client-supplied DAG plans with step dependencies (Phase 2 & 5)
  - Variable resolution (e.g. {{step_1.result.range}}) (Phase 4)
  - Disk-backed state persistence under .excel_mcp_workflows/ (Phase 2)
  - Pausing, resuming, and step-level transaction rollback (Phase 3)
"""

from __future__ import annotations
import os
import json
import time
import hashlib
import re
from pathlib import Path
from typing import Dict, Any, List, Optional, Set

from logger import logger, log_audit
from workbook_registry import registry
from snapshot import create_snapshot, restore_snapshot
from shared_context import canonical_response, ok_canonical, err_canonical


# ── Workflow State Storage ───────────────────────────────────────────────────

_WORKFLOW_DIR = Path(os.path.dirname(__file__)) / ".excel_mcp_workflows"
_WORKFLOW_DIR.mkdir(exist_ok=True)


def _get_workflow_file(workflow_id: str) -> Path:
    return _WORKFLOW_DIR / f"{workflow_id}.json"


def _save_workflow_state(state: dict) -> None:
    wf_id = state["workflow_id"]
    file_path = _get_workflow_file(wf_id)
    with open(file_path, "w", encoding="utf-8") as f:
        json.dump(state, f, indent=2, default=str)


def _load_workflow_state(workflow_id: str) -> Optional[dict]:
    file_path = _get_workflow_file(workflow_id)
    if not file_path.exists():
        return None
    try:
        with open(file_path, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception as e:
        logger.error("Failed to load workflow state for '%s': %s", workflow_id, e)
        return None


# ── Variable Resolution (Phase 4: {{step_1.result.range}}) ──────────────────

def _resolve_variables(obj: Any, step_outputs: Dict[str, dict]) -> Any:
    """
    Recursively substitutes {{step_id.key.path}} template variables in parameters.
    """
    if isinstance(obj, str):
        pattern = r"\{\{([\w\d_\-\.]+)\}\}"
        matches = re.findall(pattern, obj)
        if not matches:
            return obj
        
        val_str = obj
        for match in matches:
            parts = match.split(".")
            step_id = parts[0]
            if step_id not in step_outputs:
                raise ValueError(f"Variable resolution error: Step '{step_id}' output not found.")
            
            curr = step_outputs[step_id]
            for part in parts[1:]:
                if isinstance(curr, dict) and part in curr:
                    curr = curr[part]
                else:
                    raise ValueError(f"Variable resolution error: Key '{part}' not found in step '{step_id}' output.")
            
            # If the template is the entire string, return raw type (e.g. dict, int)
            if obj.strip() == f"{{{{{match}}}}}":
                return curr
            val_str = val_str.replace(f"{{{{{match}}}}}", str(curr))
        return val_str

    elif isinstance(obj, dict):
        return {k: _resolve_variables(v, step_outputs) for k, v in obj.items()}

    elif isinstance(obj, list):
        return [_resolve_variables(item, step_outputs) for item in obj]

    return obj


# ── DAG Topological Sort & Cycle Detection (Phase 5) ──────────────────────────

def _topological_sort(operations: List[dict]) -> List[dict]:
    """
    Sort steps according to their depends_on declarations using Kahn's algorithm.
    Raises ValueError if a cycle is detected.
    """
    op_map = {op["id"]: op for op in operations}
    in_degree = {op["id"]: 0 for op in operations}
    adj_list = {op["id"]: [] for op in operations}

    for op in operations:
        deps = op.get("depends_on", [])
        for dep in deps:
            if dep not in op_map:
                raise ValueError(f"DAG Error: Step '{op['id']}' depends on non-existent step '{dep}'.")
            adj_list[dep].append(op["id"])
            in_degree[op["id"]] += 1

    queue = [op_id for op_id, deg in in_degree.items() if deg == 0]
    sorted_order = []

    while queue:
        curr = queue.pop(0)
        sorted_order.append(op_map[curr])

        for neighbor in adj_list[curr]:
            in_degree[neighbor] -= 1
            if in_degree[neighbor] == 0:
                queue.append(neighbor)

    if len(sorted_order) != len(operations):
        raise ValueError("DAG Error: Cyclic dependency detected in workflow plan operations.")

    return sorted_order


# ── Workflow Engine API ────────────────────────────────────────────────────────

def start_workflow(plan: dict, tool_runner: Optional[Callable] = None) -> dict:
    """
    Executes a multi-step client-authored DAG workflow plan.

    Plan Schema (§4.1):
      {
        "workflow_id": "wf_12345",  (optional, auto-generated if missing)
        "transaction": true,        (optional, snapshot & rollback on failure)
        "workbook_path": "...",     (optional target file)
        "operations": [
           { "id": "step_1", "operation": "open_workbook", "parameters": { "path": "sales.xlsx" } },
           { "id": "step_2", "operation": "clean_data", "parameters": { "sheet": "Sales" }, "depends_on": ["step_1"] }
        ]
      }
    """
    wf_id = plan.get("workflow_id") or f"wf_{int(time.time())}_{os.urandom(3).hex()}"
    use_transaction = plan.get("transaction", True)
    wb_path = plan.get("workbook_path")

    # Topological sort & cycle check
    try:
        sorted_ops = _topological_sort(plan.get("operations", []))
    except ValueError as dag_err:
        return err_canonical("start_workflow", "DAG_VALIDATION_ERROR", str(dag_err))

    # Initialize workflow state
    state = {
        "workflow_id": wf_id,
        "status": "running",
        "created_at": time.time(),
        "updated_at": time.time(),
        "transaction": use_transaction,
        "workbook_path": wb_path,
        "snapshot_path": None,
        "current_step": None,
        "completed_steps": [],
        "step_outputs": {},
        "operations": [op["id"] for op in sorted_ops],
        "op_specs": {op["id"]: op for op in sorted_ops},
        "error": None,
    }

    # Transaction snapshot (Phase 3)
    if use_transaction and wb_path and os.path.exists(wb_path):
        try:
            state["snapshot_path"] = create_snapshot(wb_path)
            logger.info("Workflow '%s' snapshot created at '%s'", wf_id, state["snapshot_path"])
        except Exception as snap_err:
            logger.warning("Could not create transaction snapshot: %s", snap_err)

    _save_workflow_state(state)
    log_audit("WORKFLOW_START", {"workflow_id": wf_id, "steps": len(sorted_ops)})

    # Execute workflow steps
    return _execute_workflow_loop(state, tool_runner)


def _execute_workflow_loop(state: dict, tool_runner: Optional[Callable] = None) -> dict:
    wf_id = state["workflow_id"]
    step_outputs = state.get("step_outputs", {})
    completed = set(state.get("completed_steps", []))

    for op_id in state["operations"]:
        if op_id in completed:
            continue

        op_spec = state["op_specs"][op_id]
        op_name = op_spec["operation"]

        state["current_step"] = op_id
        state["updated_at"] = time.time()
        _save_workflow_state(state)

        # Resolve template variables {{step_1.result.range}}
        try:
            resolved_params = _resolve_variables(op_spec.get("parameters", {}), step_outputs)
        except ValueError as var_err:
            return _fail_workflow(state, op_id, "VARIABLE_RESOLUTION_FAILED", str(var_err))

        logger.info("Workflow '%s' executing step '%s' (%s)", wf_id, op_id, op_name)

        # Invoke actual tool
        step_result = None
        if tool_runner:
            try:
                step_result = tool_runner(op_name, **resolved_params)
            except Exception as tool_err:
                step_result = err_canonical(op_name, "STEP_EXCEPTED", str(tool_err))
        else:
            # Internal dispatcher
            step_result = _dispatch_internal_tool(op_name, resolved_params)

        # Check step status
        if isinstance(step_result, dict) and step_result.get("status") == "failed":
            error_msg = step_result.get("error", {}).get("message", "Step failed.")
            return _fail_workflow(state, op_id, "STEP_FAILED", f"Step '{op_id}' ({op_name}) failed: {error_msg}")

        # Record step success
        step_outputs[op_id] = step_result if isinstance(step_result, dict) else {"result": step_result}
        completed.add(op_id)
        state["completed_steps"].append(op_id)
        state["step_outputs"] = step_outputs
        _save_workflow_state(state)

    # All steps completed successfully
    state["status"] = "completed"
    state["current_step"] = None
    state["updated_at"] = time.time()
    _save_workflow_state(state)

    log_audit("WORKFLOW_COMPLETED", {"workflow_id": wf_id, "completed_steps": len(completed)})

    return ok_canonical(
        operation="start_workflow",
        workbook_id=wf_id,
        result={
            "workflow_id": wf_id,
            "status": "completed",
            "completed_steps": len(completed),
            "step_outputs": step_outputs,
        }
    )


def _fail_workflow(state: dict, step_id: str, code: str, message: str) -> dict:
    wf_id = state["workflow_id"]
    state["status"] = "failed"
    state["error"] = {"step_id": step_id, "code": code, "message": message}
    state["updated_at"] = time.time()

    # Phase 3 Transaction Rollback
    rolled_back = False
    if state.get("transaction") and state.get("snapshot_path") and state.get("workbook_path"):
        try:
            restore_snapshot(state["snapshot_path"], state["workbook_path"])
            rolled_back = True
            logger.info("Workflow '%s' rolled back cleanly from snapshot.", wf_id)
        except Exception as rb_err:
            logger.error("Workflow '%s' rollback failed: %s", wf_id, rb_err)

    state["rolled_back"] = rolled_back
    _save_workflow_state(state)
    log_audit("WORKFLOW_FAILED", {"workflow_id": wf_id, "failed_step": step_id, "rolled_back": rolled_back})

    return err_canonical(
        operation="workflow_execution",
        code=code,
        message=message,
        workbook_id=wf_id,
        recoverable=True,
        rollback_available=rolled_back,
        details={"failed_step": step_id, "rolled_back": rolled_back},
    )


def resume_workflow(workflow_id: str, from_step: Optional[str] = None, tool_runner: Optional[Callable] = None) -> dict:
    """
    Resumes a paused or failed workflow from disk state (§3 Phase 2).
    """
    state = _load_workflow_state(workflow_id)
    if not state:
        return err_canonical("resume_workflow", "WORKFLOW_NOT_FOUND", f"No workflow found with ID '{workflow_id}'.")

    if from_step:
        # Clear steps from from_step onward
        ops = state["operations"]
        if from_step in ops:
            idx = ops.index(from_step)
            to_remove = set(ops[idx:])
            state["completed_steps"] = [s for s in state["completed_steps"] if s not in to_remove]
            for s in to_remove:
                state["step_outputs"].pop(s, None)

    state["status"] = "running"
    state["error"] = None
    _save_workflow_state(state)

    return _execute_workflow_loop(state, tool_runner)


def get_workflow_status(workflow_id: str) -> dict:
    """Returns the current status and outputs of a workflow."""
    state = _load_workflow_state(workflow_id)
    if not state:
        return err_canonical("get_workflow_status", "WORKFLOW_NOT_FOUND", f"No workflow found with ID '{workflow_id}'.")
    return ok_canonical(operation="get_workflow_status", workbook_id=workflow_id, result=state)


def cancel_workflow(workflow_id: str) -> dict:
    """Cancels a running or paused workflow and restores snapshot if available."""
    state = _load_workflow_state(workflow_id)
    if not state:
        return err_canonical("cancel_workflow", "WORKFLOW_NOT_FOUND", f"No workflow found with ID '{workflow_id}'.")

    state["status"] = "cancelled"
    _save_workflow_state(state)
    return ok_canonical(operation="cancel_workflow", workbook_id=workflow_id, result={"status": "cancelled"})


def _dispatch_internal_tool(op_name: str, params: dict) -> Any:
    """Internal tool lookup table for workflow engine execution."""
    import data_tools
    import table_tools
    import editing_tools
    import format_tools
    import chart_tools
    import verification
    import workbook_tools

    tool_map = {
        "open_workbook": workbook_tools.open_workbook,
        "create_workbook": workbook_tools.create_workbook,
        "save_workbook": workbook_tools.save_workbook,
        "read_data": data_tools.read_data,
        "write_data": data_tools.write_data,
        "write_table": table_tools.write_table,
        "autofit_smart": table_tools.autofit_smart,
        "describe_sheet": table_tools.describe_sheet,
        "clean_data": editing_tools.clean_data,
        "remove_duplicates": data_tools.remove_duplicates,
        "filter_data": data_tools.filter_data,
        "sort_range": editing_tools.sort_range,
        "insert_rows": editing_tools.insert_rows,
        "delete_rows": editing_tools.delete_rows,
        "format_range": format_tools.format_range,
        "create_chart": chart_tools.create_chart,
        "create_dashboard": chart_tools.create_dashboard,
        "create_pivot_table": chart_tools.create_pivot_table,
        "scan_errors": verification.scan_formula_errors,
    }

    if op_name not in tool_map:
        raise ValueError(f"Unknown tool operation: '{op_name}'")

    fn = tool_map[op_name]
    res = fn(**params)
    if isinstance(res, str):
        try:
            return json.loads(res)
        except Exception:
            return {"message": res}
    return res
