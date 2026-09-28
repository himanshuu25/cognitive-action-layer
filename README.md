# Excel Automation MCP Server

An AI-powered Excel Automation Model Context Protocol (MCP) Server that seamlessly bridges LLMs (such as Claude Desktop) with Microsoft Excel utilizing `xlwings`, `FastMCP`, and `pandas`. Built for **critical-grade accuracy, security, deterministic workflow execution, and high-quality first-render UX**.

---

## 🚀 Highlights & Core Capabilities

- **Durable Enterprise Workflow Engine (`workflow_engine.py`)**:
  - Executes multi-step DAG plans with topological sorting and cycle detection.
  - Variable template substitution engine (e.g. `{{step_1.result.range}}`).
  - Disk-backed state persistence (`.excel_mcp_workflows/`) enabling workflow pausing and step-level resumption.
  - Multi-step transactional snapshot boundaries with auto-rollback on failure.
- **Canonical Response & Validation Contract (§2.1)**:
  - Standardized JSON return shape across all tools: `{status, operation, workbook_id, sheet, input, result, verification, error, duration_ms}`.
  - Idempotency key caching (`idempotency_key`) preventing duplicate executions.
  - Pre/post-condition contract validation decorators (`@validate_tool_contract`).
- **Pluggable Control Totals & Reconciliation (`control_totals.py`)**:
  - Runs structural, row count, numeric sum, null count, and SHA-256 hash validation pipelines.
- **Deep Workbook & Formula Introspection (`table_tools.py` & `workbook_registry.py`)**:
  - `describe_workbook`: Generates a full structural map of all visible and hidden sheets, used ranges, named tables, and named ranges.
  - `get_formula_graph`: Extracts formulas in a range and builds an AST-parsed formula dependency graph showing cell references.
  - `compare_workbooks`: Performs structural cross-file diffing between workbooks.
- **Deterministic Business-Rule Evaluator (`rules_evaluator.py`)**:
  - Sandboxed AST expression evaluator (`revenue > 10000 and status == 'Active'`) rejecting code injection attempts (`eval()`, `__import__`).
- **Self-Finishing Table Write (`write_table`)**:
  - Atomic sequence: **Column Reordering → Write Values → Number Formats → Autofit → Width Clamping (8-48) → Text Wrapping → Freeze Header → Native Excel Table → Read-back Verification**.
  - Eliminates the `####` first-render column width bug completely.
  - Accepts explicit left-to-right column sequences (`column_order`).
- **First-Render Chart & Dashboard Engine (`chart_tools.py`)**:
  - **No Chart Overlaps**: Legend auto-positioned at bottom (`xlLegendPositionBottom`).
  - **X-Axis Auto-Rotation**: Category labels auto-angle at `-45°` when long strings are present.
  - **`create_dashboard` Tool**: Generates 2x2 grid or stacked multi-chart dashboards with uniform spacing in a single call.
- **Security & Audit Hardening**:
  - Path sandboxing (directory allowlist + path traversal `../` rejection).
  - Untrusted text formula/DDE injection escaping (`=`, `+`, `-`, `@`).
  - Redaction of passwords/secret arguments in logs.
  - SHA-256 chained, tamper-evident audit log (`logs/tamper_evident_audit.jsonl`).
  - Dry-run diff preview (`preview_write`).

---

## 📁 Repository Structure

```text
excelwithmcp/
├── server.py               # Main FastMCP Server registering all 60+ MCP tools
├── workflow_engine.py      # Durable multi-step DAG workflow engine & step transaction rollback
├── workbook_registry.py    # Unique workbook_id handle registry & cross-workbook comparison
├── control_totals.py       # Pluggable control totals & numeric reconciliation pipelines
├── rules_evaluator.py      # Sandboxed AST expression evaluator for business rules
├── table_tools.py          # write_table, autofit_smart, describe_sheet, describe_workbook, get_formula_graph
├── data_tools.py           # Data read/write, injection escaping, dedupe, filter operators
├── editing_tools.py        # Insert/delete rows & cols, sort, AutoFilter, named ranges, comments
├── format_tools.py         # Visual formatting, conditional formatting, dropdown data validation
├── chart_tools.py          # Native Excel charts, PlotBy orientation, create_dashboard, pivot tables
├── verification.py         # Write-and-verify read-back, control totals, formula error scanner
├── snapshot.py             # Backup snapshot creation, restore, and safe operation harness
├── security.py             # Path sandbox, secret redaction, resource limits, SHA-256 audit log
├── dry_run.py              # Before/after diff preview generator
├── shared_context.py       # Canonical contract builder, idempotency cache, pre/post decorators
├── logger.py               # Rotating file logging & per-tool log decorators
├── pack_installer.py       # Standalone installer script with dynamic Claude path discovery
├── BASELINE_AUDIT.md       # Audit matrix of all 60+ tools across compliance criteria
├── PROJECT_OVERVIEW.md     # Executive project overview & architecture showcase
├── requirements.txt        # Pinned Python dependencies
├── dist/                   # Compiled Standalone Executables
│   ├── install.exe         # 🎯 Standalone 1-click Installer EXE for distribution
│   └── excel_mcp.exe       # Standalone Server Executable (embedded inside install.exe)
└── tests/
    ├── test_helpers.py         # Pytest unit tests (38 tests)
    └── test_workflow_engine.py # Workflow engine & rules evaluator unit tests (9 tests)
```

---

## 📦 Zero-Dependency Single-File Installer (`install.exe`)

For distribution to end-users without Python installed:

1. Copy `dist/install.exe` from this repository.
2. Send `install.exe` to the target user.
3. When double-clicked, `install.exe` automatically:
   - Extracts `excel_mcp.exe` to `%LOCALAPPDATA%\Programs\ExcelMCP\excel_mcp.exe`.
   - Dynamically discovers all Claude Desktop installations (both standard desktop installer and Windows Store AppX packages).
   - Automatically merges the `excel-mcp` configuration into `claude_desktop_config.json`.
4. The user simply restarts Claude Desktop and the server is live!

---

## 🛠️ Developer Setup & Running

### Requirements
- Operating System: Windows (requires Microsoft Excel installed)
- Python: 3.10+

### Installation

```bash
pip install -r requirements.txt
```

### Running the MCP Server Manually

```bash
python server.py
```

### Running Unit Tests

```bash
python -m pytest tests/ -v
```

### Rebuilding the Installer Executables

```bash
# 1. Compile the server executable
pyinstaller --onefile --name excel_mcp server.py

# 2. Compile the single-file installer (bundles excel_mcp.exe inside it)
pyinstaller --onefile --clean --name install --add-data "dist/excel_mcp.exe;." pack_installer.py
```
