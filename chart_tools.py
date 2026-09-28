"""
chart_tools.py — Production-hardened
Chart creation and pivot table with:
  - PlotBy (row vs column orientation) setting
  - Structured error returns
  - All suppression using shared excel_suppressed context manager
"""

from __future__ import annotations
import xlwings as xw
import pandas as pd

from workbook_tools import get_active_book, get_sheet
from shared_context import (
    excel_suppressed,
    validate_sheet_name,
    validate_range_address,
    ok, err, result_to_str,
)
from logger import logger


# xlwings / Excel COM chart type integer constants
CHART_TYPE_MAP = {
    "column":            51,    # xlColumnClustered
    "column_clustered":  51,
    "column_stacked":    52,    # xlColumnStacked
    "column_100":        53,    # xlColumnStacked100
    "bar":               57,    # xlBarClustered
    "bar_clustered":     57,
    "bar_stacked":       58,    # xlBarStacked
    "bar_100":           59,    # xlBarStacked100
    "line":              4,     # xlLine
    "line_markers":      65,    # xlLineMarkers
    "line_stacked":      63,    # xlLineStacked
    "pie":               5,     # xlPie
    "pie_exploded":      69,    # xlPieExploded
    "doughnut":          -4120, # xlDoughnut
    "area":              1,     # xlArea
    "area_stacked":      76,    # xlAreaStacked
    "scatter":           -4169, # xlXYScatter
    "scatter_smooth":    64,    # xlXYScatterSmooth
    "radar":             -4151, # xlRadar
    "combo":             -4169, # fallback to scatter
}


def create_chart(
    sheet: str,
    chart_type: str,
    source_range: str,
    title: str = None,
    x_axis_title: str = "",
    y_axis_title: str = "",
    target_cell: str = "G2",
    width: int = 420,
    height: int = 260,
    plot_by: str = "columns",
    show_legend: bool = True,
    show_data_labels: bool = False,
) -> str:
    """
    Creates a native Excel chart positioned next to data.

    Parameters:
    - sheet: Worksheet name.
    - chart_type: Type of chart ('column', 'bar', 'line', 'pie', 'area', 'scatter', etc.).
    - source_range: Cell range containing data and headers (e.g. 'A1:B10').
    - title: Optional chart title.
    - x_axis_title: Optional label for the X (horizontal) axis.
    - y_axis_title: Optional label for the Y (vertical) axis.
    - target_cell: Top-left cell where the chart will be placed (default 'G2').
    - width: Chart width in points (default 420).
    - height: Chart height in points (default 260).
    - plot_by: 'columns' (default) or 'rows' — sets data orientation.
    - show_legend: Show chart legend (default True).
    - show_data_labels: Show data labels on chart series (default False).
    """
    try:
        validate_sheet_name(sheet)
        source_range = validate_range_address(source_range)

        book = get_active_book()
        s = get_sheet(book, sheet)
        target_rng = s.range(target_cell)

        with excel_suppressed(book.app):
            chart = s.charts.add(
                left=target_rng.left,
                top=target_rng.top,
                width=width,
                height=height,
            )

            chart.set_source_data(s.range(source_range))

            c_type_int = CHART_TYPE_MAP.get(chart_type.lower(), 51)
            try:
                chart.api.ChartType = c_type_int
            except Exception:
                pass

            # Set PlotBy (row vs column orientation)
            try:
                # xlColumns = 2, xlRows = 1
                chart.api.PlotBy = 2 if plot_by.lower() == "columns" else 1
            except Exception:
                pass

            # Title
            try:
                chart.api.HasTitle = True
                chart.api.ChartTitle.Text = title if title else (chart_type.capitalize() + " Chart")
                chart.api.ChartTitle.Font.Size = 12
                chart.api.ChartTitle.Font.Bold = True
            except Exception:
                pass

            # Axis titles and label formatting (not for pie/doughnut)
            c_lower = chart_type.lower()
            if c_lower not in ("pie", "doughnut", "pie_exploded"):
                try:
                    ax = chart.api.Axes(1)   # xlCategory
                    if x_axis_title:
                        ax.HasTitle = True
                        ax.AxisTitle.Text = x_axis_title
                        ax.AxisTitle.Font.Size = 9
                    # Prevent label overlap by tilting category labels slightly if text is present
                    try:
                        ax.TickLabels.Orientation = -45
                    except Exception:
                        pass
                except Exception:
                    pass

                try:
                    if y_axis_title:
                        ay = chart.api.Axes(2)   # xlValue
                        ay.HasTitle = True
                        ay.AxisTitle.Text = y_axis_title
                        ay.AxisTitle.Font.Size = 9
                except Exception:
                    pass

            # Legend position (xlLegendPositionBottom = -4107) to avoid data overlap
            try:
                chart.api.HasLegend = show_legend
                if show_legend:
                    chart.api.Legend.Position = -4107  # xlLegendPositionBottom
                    chart.api.Legend.Font.Size = 9
            except Exception:
                pass

            # Data labels
            if show_data_labels:
                try:
                    for series in chart.api.SeriesCollection():
                        series.HasDataLabels = True
                        series.DataLabels.Font.Size = 8
                except Exception:
                    pass

        return result_to_str(ok(
            f"Created '{chart_type}' chart '{title or 'Chart'}' in '{sheet}' at '{target_cell}'.",
            data={"chart_type": chart_type, "source_range": source_range, "location": target_cell}
        ))

    except Exception as e:
        logger.exception("create_chart failed")
        return result_to_str(err("CREATE_CHART_ERROR", str(e)))


def create_dashboard(
    sheet: str,
    charts_config: list[dict],
    start_cell: str = "G2",
    layout: str = "2x2",
    chart_width: int = 400,
    chart_height: int = 250,
) -> str:
    """
    Generates a clean multi-chart dashboard in a structured grid (e.g. 2x2 or vertical stack)
    with zero chart overlap and uniform styling in a single call.

    Parameters:
    - sheet: Worksheet name.
    - charts_config: List of dicts specifying chart configurations. Each dict can include:
        {"chart_type": "column", "source_range": "A1:B10", "title": "Sales", "x_axis_title": "Month", "y_axis_title": "USD"}
    - start_cell: Top-left cell for the first chart in the dashboard (default 'G2').
    - layout: Layout arrangement ('2x2' grid or 'stacked' vertical column).
    - chart_width: Width of each chart in points (default 400).
    - chart_height: Height of each chart in points (default 250).
    """
    try:
        validate_sheet_name(sheet)
        book = get_active_book()
        s = get_sheet(book, sheet)

        start_rng = s.range(start_cell)
        start_left = start_rng.left
        start_top = start_rng.top

        padding_x = 20
        padding_y = 20

        results = []

        with excel_suppressed(book.app):
            for idx, cfg in enumerate(charts_config):
                # Calculate grid position
                if layout.lower() == "2x2":
                    col_idx = idx % 2
                    row_idx = idx // 2
                else:  # stacked
                    col_idx = 0
                    row_idx = idx

                left_pos = start_left + col_idx * (chart_width + padding_x)
                top_pos = start_top + row_idx * (chart_height + padding_y)

                # Add chart object
                chart = s.charts.add(
                    left=left_pos,
                    top=top_pos,
                    width=chart_width,
                    height=chart_height,
                )

                src_range = validate_range_address(cfg.get("source_range"))
                chart.set_source_data(s.range(src_range))

                ctype = cfg.get("chart_type", "column").lower()
                c_type_int = CHART_TYPE_MAP.get(ctype, 51)
                try:
                    chart.api.ChartType = c_type_int
                except Exception:
                    pass

                # Orientation
                plot_by = cfg.get("plot_by", "columns")
                try:
                    chart.api.PlotBy = 2 if plot_by.lower() == "columns" else 1
                except Exception:
                    pass

                # Title
                title = cfg.get("title", f"Chart {idx + 1}")
                try:
                    chart.api.HasTitle = True
                    chart.api.ChartTitle.Text = title
                    chart.api.ChartTitle.Font.Size = 11
                    chart.api.ChartTitle.Font.Bold = True
                except Exception:
                    pass

                # Legend at bottom
                try:
                    chart.api.HasLegend = cfg.get("show_legend", True)
                    if chart.api.HasLegend:
                        chart.api.Legend.Position = -4107  # xlLegendPositionBottom
                        chart.api.Legend.Font.Size = 9
                except Exception:
                    pass

                # Axes
                if ctype not in ("pie", "doughnut", "pie_exploded"):
                    try:
                        ax = chart.api.Axes(1)
                        if cfg.get("x_axis_title"):
                            ax.HasTitle = True
                            ax.AxisTitle.Text = cfg["x_axis_title"]
                            ax.AxisTitle.Font.Size = 9
                        try:
                            ax.TickLabels.Orientation = -45
                        except Exception:
                            pass
                    except Exception:
                        pass

                    try:
                        if cfg.get("y_axis_title"):
                            ay = chart.api.Axes(2)
                            ay.HasTitle = True
                            ay.AxisTitle.Text = cfg["y_axis_title"]
                            ay.AxisTitle.Font.Size = 9
                    except Exception:
                        pass

                results.append(title)

        return result_to_str(ok(
            f"Created dashboard with {len(results)} chart(s) in '{sheet}' starting at '{start_cell}' ({layout} layout).",
            data={"charts": results, "count": len(results)}
        ))

    except Exception as e:
        logger.exception("create_dashboard failed")
        return result_to_str(err("CREATE_DASHBOARD_ERROR", str(e)))


def create_pivot_table(
    source_sheet: str,
    source_range: str = None,
    target_sheet: str = "Pivot_Summary",
    target_cell: str = "A3",
    row_fields: list[str] = None,
    col_fields: list[str] = None,
    data_field: str = None,
    agg_func: str = "sum",
) -> str:
    """
    Creates a Pivot Table. Tries native COM first; falls back to Pandas aggregation.

    Parameters:
    - source_sheet: Source worksheet name.
    - source_range: Data range (e.g. 'A1:F500'). Auto-detects if omitted.
    - target_sheet: Target worksheet name (defaults to 'Pivot_Summary').
    - target_cell: Target cell position (defaults to 'A3').
    - row_fields: Columns to group by rows.
    - col_fields: Columns to group by columns (optional).
    - data_field: Column to aggregate.
    - agg_func: 'sum', 'count', 'average', 'max', 'min' (defaults to 'sum').
    """
    try:
        validate_sheet_name(source_sheet)

        book = get_active_book()
        s_sheet = get_sheet(book, source_sheet)
        src_range = s_sheet.range(source_range) if source_range else s_sheet.used_range

        # Ensure target sheet exists
        t_sheet = None
        for sh in book.sheets:
            if sh.name.lower() == target_sheet.lower():
                t_sheet = sh
                break
        if not t_sheet:
            t_sheet = book.sheets.add(name=target_sheet)

        # ── Try native COM Pivot ──────────────────────────────────────────
        with excel_suppressed(book.app):
            try:
                t_sheet.clear()

                func_map = {
                    "sum":     -4157,
                    "count":   -4112,
                    "average": -4106,
                    "mean":    -4106,
                    "max":     -4136,
                    "min":     -4139,
                }
                f_const = func_map.get(agg_func.lower(), -4157)
                pt_name = f"PT_{target_sheet[:8]}_{target_cell}".replace(" ", "_")

                pc = book.api.PivotCaches().Create(SourceType=1, SourceData=src_range.api)
                pt = pc.CreatePivotTable(
                    TableDestination=t_sheet.range(target_cell).api,
                    TableName=pt_name,
                )

                if row_fields:
                    for idx, rf in enumerate(row_fields, 1):
                        f = pt.PivotFields(rf)
                        f.Orientation = 1   # xlRowField
                        f.Position = idx

                if col_fields:
                    for idx, cf in enumerate(col_fields, 1):
                        f = pt.PivotFields(cf)
                        f.Orientation = 2   # xlColumnField
                        f.Position = idx

                if data_field:
                    df_field = pt.PivotFields(data_field)
                    pt.AddDataField(df_field, f"{agg_func.capitalize()} of {data_field}", f_const)

                return result_to_str(ok(
                    f"Native Excel Pivot Table created in '{target_sheet}' at '{target_cell}'.",
                    data={"method": "native_com"}
                ))

            except Exception as com_err:
                logger.warning("COM pivot failed, using Pandas fallback: %s", com_err)

                # ── Pandas Fallback ──────────────────────────────────────────
                values = src_range.value
                if not values or not isinstance(values, list) or len(values) <= 1:
                    return result_to_str(err(
                        "PIVOT_ERROR",
                        f"COM pivot failed: {com_err}. Fallback also failed: no tabular data found."
                    ))

                headers = [str(h) if h is not None else f"Col_{i}" for i, h in enumerate(values[0])]
                df = pd.DataFrame(values[1:], columns=headers)

                missing = []
                if row_fields:
                    missing += [r for r in row_fields if r not in df.columns]
                if col_fields:
                    missing += [c for c in col_fields if c not in df.columns]
                if data_field and data_field not in df.columns:
                    missing.append(data_field)
                if missing:
                    return result_to_str(err(
                        "COLUMN_NOT_FOUND",
                        f"Columns not found: {missing}. Available: {list(df.columns)}"
                    ))

                pivot_df = df.pivot_table(
                    values=data_field,
                    index=row_fields,
                    columns=col_fields,
                    aggfunc=agg_func if data_field else "count",
                    fill_value=0,
                ).reset_index()

                t_sheet.clear()
                t_sheet.range("A1").value = (
                    f"Pivot Summary — {agg_func.capitalize()} of {data_field or 'Count'}"
                )
                t_sheet.range("A1").font.bold = True
                t_sheet.range("A1").font.size = 13
                t_sheet.range(target_cell).value = pivot_df

                # Style header
                try:
                    hdr = t_sheet.range(target_cell).expand("right")
                    hdr.font.bold = True
                    hdr.color = (27, 54, 93)
                    hdr.font.color = (255, 255, 255)
                except Exception:
                    pass

                return result_to_str(ok(
                    f"Pivot Summary (Pandas fallback) in '{target_sheet}' at '{target_cell}'. "
                    f"{len(pivot_df)} result rows.",
                    data={"method": "pandas_fallback", "rows": len(pivot_df)}
                ))

    except Exception as e:
        logger.exception("create_pivot_table failed")
        return result_to_str(err("PIVOT_TABLE_ERROR", str(e)))
