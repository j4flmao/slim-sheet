"""`add_chart` use-case: a native Excel chart from a block of cells (PLAN §5.3).

Layout convention (what agents and humans expect): the first row holds series names and
the first column holds categories (x values for scatter); every other column is a series.
openpyxl keeps charts across later edits; rows_ops/sheet_ops rewrite their ranges.
"""

from __future__ import annotations

from typing import Any

from slim_sheet.adapters.writers.chart_refs import build_chart
from slim_sheet.config import Settings
from slim_sheet.core.a1 import index_to_col, parse_range
from slim_sheet.core.errors import ErrorCode, SlimError
from slim_sheet.core.models import OutputResult
from slim_sheet.core.refs import split_sheet
from slim_sheet.core.styles import choice
from slim_sheet.services.editing import Change, edit_workbook, get_sheet

CHART_TYPES = ("column", "bar", "line", "pie", "area", "scatter")


def add_chart(
    path: str,
    sheet: str,
    type: str,
    data_range: str,
    settings: Settings,
    anchor: str | None = None,
    title: str | None = None,
    x_title: str | None = None,
    y_title: str | None = None,
    out_path: str | None = None,
    overwrite: bool = False,
) -> OutputResult:
    kind = choice(type, CHART_TYPES, "chart type")
    data_sheet, address = split_sheet(data_range)
    rng = parse_range(address)
    if rng.n_rows < 2 or rng.n_cols < 2:
        raise SlimError(
            ErrorCode.BAD_ARGUMENT,
            f"data_range {address} needs a header row and a category column plus >= 1 series",
            hint="e.g. data_range='A1:C13' (A = months, B..C = series)",
        )
    if anchor is not None:
        parse_range(anchor)

    def mutate(wb: Any) -> Change:
        host = get_sheet(wb, sheet)
        source = get_sheet(wb, data_sheet) if data_sheet else host
        place = anchor or f"{index_to_col(rng.right + 2)}{rng.top + 1}"
        chart = build_chart(kind, source, address)
        chart.title = title
        if kind != "pie":
            chart.x_axis.title, chart.y_axis.title = x_title, y_title
        host.add_chart(chart, place)
        series = 1 if kind == "pie" else rng.n_cols - 1
        return Change(
            f"{kind} chart of {series} series from {source.title}!{address} at {host.title}!{place}"
        )

    return edit_workbook(path, settings, f"add_chart {kind} {sheet}", mutate, out_path, overwrite)
