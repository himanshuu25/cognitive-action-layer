# Section 1: Baseline Audit Matrix of All 40+ Tools

| Tool Name | Returns Canonical Result? | Validates Preconditions? | Validates Postconditions? | Assumes "Active Workbook"? | Idempotent with Key? | Compliance Action Taken |
|---|---|---|---|---|---|---|
| `open_workbook` | Yes | Path Sandbox | File Existence | No | Yes | Retrofitted to `canonical_response` + `workbook_id` |
| `create_workbook` | Yes | Path Sandbox | File Creation | No | Yes | Retrofitted to `canonical_response` + `workbook_id` |
| `save_workbook` | Yes | Overwrite Safety | Save Check | No | Yes | Retrofitted to `canonical_response` + `workbook_id` |
| `list_sheets` | Yes | Handle Validity | Non-empty List | No | Yes | Uses `workbook_id` handle |
| `list_open_workbooks` | Yes | None | Non-empty List | No | Yes | Queries registry handles |
| `add_sheet` | Yes | Sheet Name Validity | Sheet Presence | No | Yes | Retrofitted to `canonical_response` |
| `rename_sheet` | Yes | Sheet Existence | Name Match | No | Yes | Retrofitted to `canonical_response` |
| `delete_sheet` | Yes | Sheet Count > 1 | Sheet Removal | No | Yes | Protects last sheet |
| `export_to_pdf` | Yes | Path Sandbox | PDF Existence | No | Yes | Validates PDF write |
| `protect_sheet` | Yes | Password Redaction | Protection State | No | Yes | Redacts secrets |
| `set_sheet_visibility` | Yes | Sheet Existence | Visibility State | No | Yes | Sets visibility flag |
| `describe_sheet` | Yes | Sheet Existence | Used Range Map | No | Yes | Full sheet introspection |
| `describe_workbook` | Yes | Workbook Existence | Full Map | No | Yes | **NEW Phase 8**: Deep workbook map |
| `get_formula_graph` | Yes | Sheet Existence | Formula AST | No | Yes | **NEW Phase 8**: Formula dependency graph |
| `read_data` | Yes | Range Bounds | Non-null Output | No | Yes | Converts to JSON rows |
| `write_table` | Yes | Sheet Protection/Bounds | Read-back Match | No | Yes | Self-finishing table engine |
| `write_data` | Yes | Range Bounds | Read-back Match | No | Yes | Injection-escaped write |
| `autofit_smart` | Yes | Range Bounds | Width Boundaries | No | Yes | Clamps widths 8-48 |
| `remove_duplicates` | Yes | Range Bounds | Count Delta | No | Yes | Type-preserving dedupe |
| `filter_data` | Yes | Column Presence | Match Count | No | Yes | Numeric & string operators |
| `clean_data` | Yes | Range Bounds | Clean Count | No | Yes | Strips zero-width/BOMs |
| `clear_range` | Yes | Range Bounds | Content Emptiness | No | Yes | Clears contents/formats |
| `copy_paste_range` | Yes | Source/Target Bounds | Target Contents | No | Yes | Copies values/formats |
| `find_replace` | Yes | Range Bounds | Replacement Count | No | Yes | Whole & partial match |
| `insert_rows` | Yes | Row Bounds | Row Shift | No | Yes | Inserts blank rows |
| `delete_rows` | Yes | Row Bounds | Row Shift | No | Yes | Deletes rows |
| `insert_columns` | Yes | Col Bounds | Col Shift | No | Yes | Inserts blank columns |
| `delete_columns` | Yes | Col Bounds | Col Shift | No | Yes | Deletes columns |
| `sort_range` | Yes | Column Presence | Sort Order | No | Yes | Header sorting |
| `apply_autofilter` | Yes | Range Bounds | Filter Mode | No | Yes | Enables drop-down arrows |
| `remove_autofilter` | Yes | Sheet Existence | Filter Removal | No | Yes | Removes AutoFilter |
| `create_named_range` | Yes | Name Syntax | Registry Check | No | Yes | Validates range name |
| `delete_named_range` | Yes | Name Existence | Removal Check | No | Yes | Removes named range |
| `list_named_ranges` | Yes | Handle Validity | Range List | No | Yes | Lists named ranges |
| `add_comment` | Yes | Cell Bounds | Comment Presence | No | Yes | Adds cell comment |
| `delete_comment` | Yes | Cell Bounds | Comment Removal | No | Yes | Deletes comment |
| `apply_formula` | Yes | Formula Syntax (`=`) | Calculation Check | No | Yes | Enforces formula syntax |
| `format_range` | Yes | Range Bounds | Style State | No | Yes | Hex colors, borders, fonts |
| `set_column_width` | Yes | Col Bounds | Width Check | No | Yes | Sets column width |
| `set_row_height` | Yes | Row Bounds | Height Check | No | Yes | Sets row height |
| `freeze_panes` | Yes | Cell Address | Freeze State | No | Yes | Locks header rows/cols |
| `merge_cells` | Yes | Range Bounds | Merge State | No | Yes | Merges/unmerges block |
| `add_conditional_formatting` | Yes | Range Bounds | Rule Presence | No | Yes | 12 rule types |
| `add_dropdown_validation` | Yes | Range Bounds | Rule Presence | No | Yes | Bypasses 255-char limit |
| `add_data_validation` | Yes | Range Bounds | Rule Presence | No | Yes | Tooltip & alert stops |
| `add_hyperlink` | Yes | Cell Address | Hyperlink Check | No | Yes | Clickable URL |
| `fill_range` | Yes | Direction | Fill Bounds | No | Yes | Directional auto-fill |
| `create_chart` | Yes | Range Bounds | Chart Object | No | Yes | Bottom legend & -45° labels |
| `create_dashboard` | Yes | Config List | Grid Positioning | No | Yes | Multi-chart 2x2 layout |
| `create_pivot_table` | Yes | Source Range | Pivot Table | No | Yes | COM + Pandas fallback |
| `verify_range` | Yes | Range Bounds | Match Report | No | Yes | Read-back verification |
| `scan_errors` | Yes | Range Bounds | Error Cell List | No | Yes | Recalculates & scans errors |
| `create_snapshot` | Yes | File Existence | Snapshot Path | No | Yes | Timestamped backup |
| `restore_snapshot` | Yes | Snapshot Existence | File Overwrite | No | Yes | Reverts to snapshot |
| `list_snapshots` | Yes | File Existence | Snapshot List | No | Yes | Lists snapshots |
| `preview_write` | Yes | Range Bounds | Cell Diff List | No | Yes | Dry-run diff preview |
| `verify_audit_chain` | Yes | Log File | SHA-256 Hash Chain | No | Yes | Tamper-evident verification |
| `start_workflow` | Yes | Plan Schema & DAG | Execution Journal | No | Yes | Multi-step DAG engine |
| `get_workflow_status` | Yes | Workflow ID | State Object | No | Yes | Workflow status query |
| `resume_workflow` | Yes | Workflow ID | Execution Resume | No | Yes | Resumes from step |
| `cancel_workflow` | Yes | Workflow ID | Rollback Execution | No | Yes | Triggers rollback |
| `run_control_totals` | Yes | Range & Checks | Metrics Object | No | Yes | Pluggable control totals |
| `compare_workbooks` | Yes | Both Workbooks | Structural Diff | No | Yes | Cross-file diff |
| `evaluate_rule` | Yes | Expression Syntax | Result Value | No | Yes | Safe AST evaluator |
