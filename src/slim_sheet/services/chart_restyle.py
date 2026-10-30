"""Bring charts already in a workbook up to the report rules (`style_sheet`, PLAN §16.4).

openpyxl reads charts back on load, so a chart made earlier (by Excel, another tool or an older
slim-sheet) can be restyled in place: same data, same anchor, same title, but the brand colours,
number formats from its source cells, clean axes and readable labels of `style_chart`.
"""

from __future__ import annotations

import logging
from typing import Any

from openpyxl.chart import AreaChart, BarChart, DoughnutChart, LineChart, PieChart, ScatterChart

from slim_sheet.adapters.writers.chart_style import ChartText, clean_title, style_chart
from slim_sheet.core.a1 import parse_range
from slim_sheet.core.design import Brand, formats
from slim_sheet.core.errors import SlimError
from slim_sheet.core.refs import split_sheet
from slim_sheet.core.report_spec import infer_format

log = logging.getLogger(__name__)

_KINDS: tuple[tuple[type, str], ...] = (
    (DoughnutChart, "doughnut"),
    (PieChart, "pie"),
    (LineChart, "line"),
    (AreaChart, "area"),
    (ScatterChart, "scatter"),
)
_NO_FORMAT_KEYS = ("date", "year", "text")


def restyle_charts(wb: Any, ws: Any, brand: Brand) -> list[str]:
    """Restyle every chart on `ws`; -> one note (or none when the sheet has no charts)."""
    done, skipped = 0, []
    for chart in ws._charts:
        kind = chart_kind(chart)
        if kind is None:
            skipped.append(type(chart).__name__)
            continue
        _reset(chart)
        style_chart(chart, kind, brand, _chart_text(wb, chart, kind, brand))
        done += 1
    notes = [f"restyled {done} chart(s)"] if done else []
    if skipped:
        notes.append(
            f"left {len(skipped)} chart(s) as they are ({', '.join(sorted(set(skipped)))})"
        )
    return notes


def chart_kind(chart: Any) -> str | None:
    """The slim-sheet kind of a loaded chart; None for kinds it does not style (3-D, radar...)."""
    if isinstance(chart, BarChart):
        return "bar" if chart.type == "bar" else "column"
    for cls, kind in _KINDS:
        if type(chart) is cls:
            return kind
    return None


def _reset(chart: Any) -> None:
    """Drop the old manual plot layout, per-point colours and labels; style_chart sets its
    own (a fixed plot area clips the labels at the end of the longest bar)."""
    chart.layout = None  # openpyxl writes the plot area layout from chart.layout
    if isinstance(chart, BarChart):  # labelled bars need room past a fixed maximum
        chart.y_axis.scaling.min = chart.y_axis.scaling.max = None
    for series in chart.series:
        series.dPt = []
        series.dLbls = None


def _chart_text(wb: Any, chart: Any, kind: str, brand: Brand) -> ChartText:
    series = chart.series[0] if chart.series else None
    y_ref = _ref(series, "yVal" if kind == "scatter" else "val")
    values = _numbers(wb, y_ref)
    header = _series_name(wb, series)
    x_format = None
    if kind == "scatter":
        x_ref = _ref(series, "xVal")  # x has no series name: its axis title says what it is
        x_name = title_text(chart.x_axis.title) or ""
        x_format = _format(wb, x_ref, brand, x_name, _numbers(wb, x_ref))
    across, up = _axis_titles(chart, kind)
    lines = title_lines(chart.title)
    return ChartText(
        title=lines[0] if lines else None,
        subtitle=lines[1] if len(lines) > 1 else None,
        value_format=_format(wb, y_ref, brand, header, values),
        x_format=x_format,
        x_title=across,
        y_title=up,
        values=tuple(values),
    )


def _series_name(wb: Any, series: Any) -> str:
    """The series header ('Net margin') when it comes from a cell: it tells % from amounts."""
    if series is None or series.tx is None or series.tx.strRef is None:
        return ""
    cells = _cells(wb, str(series.tx.strRef.f))
    return str(cells[0].value or "") if cells else ""


def _axis_titles(chart: Any, kind: str) -> tuple[str | None, str | None]:
    """(horizontal, vertical) axis titles as the reader sees them."""
    if kind in ("pie", "doughnut"):
        return None, None
    category, value = title_text(chart.x_axis.title), title_text(chart.y_axis.title)
    return (value, category) if kind == "bar" else (category, value)


def title_lines(title: Any) -> list[str]:
    """The text of each line of a chart title (message, then subtitle); [] if absent."""
    if title is None or title.tx is None or title.tx.rich is None:
        return []
    lines = (
        clean_title("".join(run.t for run in (p.r or []) if run.t is not None))
        for p in title.tx.rich.p
    )
    return [line for line in lines if line]


def title_text(title: Any) -> str | None:
    """An axis title as one line; None if absent or a cell reference."""
    lines = title_lines(title)
    return " ".join(lines) if lines else None


def _ref(series: Any, part: str) -> str | None:
    holder = getattr(series, part, None) if series is not None else None
    if holder is None or holder.numRef is None:
        return None
    return str(holder.numRef.f)


def _cells(wb: Any, ref: str | None) -> list[Any]:
    if not ref:
        return []
    sheet, address = split_sheet(ref)
    if not sheet or sheet not in wb.sheetnames:
        return []
    try:
        rng = parse_range(address.replace("$", ""))
    except SlimError as exc:
        log.debug("Chart reference %s not restyled: %s", ref, exc)
        return []
    ws = wb[sheet]
    return [
        ws.cell(row=r + 1, column=c + 1)
        for r in range(rng.top, rng.bottom + 1)
        for c in range(rng.left, rng.right + 1)
    ]


def _numbers(wb: Any, ref: str | None) -> list[float]:
    return [
        float(c.value)
        for c in _cells(wb, ref)
        if isinstance(c.value, int | float) and not isinstance(c.value, bool)
    ]


def _format(wb: Any, ref: str | None, brand: Brand, header: str, values: list[float]) -> str | None:
    """The source cells' own number format, else one inferred from header and values."""
    for cell in _cells(wb, ref):
        if isinstance(cell.value, int | float) and not isinstance(cell.value, bool):
            if cell.number_format and cell.number_format != "General":
                return str(cell.number_format)
            break
    key, _why = infer_format(header, values)
    return formats(brand)[key] if key and key not in _NO_FORMAT_KEYS else None
