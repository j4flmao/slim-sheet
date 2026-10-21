"""Build charts, and keep existing charts pointing at the right cells.

openpyxl 3.1 reads charts back on load, so charts survive a load/save cycle, but their
series formulas (e.g. 'Data'!$B$2:$B$9) and anchors are left untouched when rows move or
a sheet is renamed. These helpers apply the same rules as formulas (core/shift.py).
"""

from __future__ import annotations

from collections.abc import Callable, Iterator
from typing import Any

from openpyxl.chart import (
    AreaChart,
    BarChart,
    DoughnutChart,
    LineChart,
    PieChart,
    Reference,
    ScatterChart,
    Series,
)

from slim_sheet.core.a1 import parse_range
from slim_sheet.core.shift import Shift, rename_in_formula, shift_address, shift_formula

WIDTH_CM, HEIGHT_CM = 16, 8
_SERIES_PARTS = ("tx", "cat", "val", "xVal", "yVal", "bubbleSize")
_REF_KINDS = ("numRef", "strRef", "multiLvlStrRef")


def build_chart(kind: str, ws: Any, data: str) -> Any:
    """First row = series names, first column = categories (x for scatter)."""
    rng = parse_range(data)
    top, left, bottom, right = rng.top + 1, rng.left + 1, rng.bottom + 1, rng.right + 1
    categories = Reference(ws, min_col=left, min_row=top + 1, max_row=bottom)
    if kind == "scatter":
        chart = ScatterChart()
        chart.style = 13
        for col in range(left + 1, right + 1):
            values = Reference(ws, min_col=col, min_row=top, max_row=bottom)
            series = Series(values, categories, title_from_data=True)
            series.marker.symbol = "circle"
            chart.series.append(series)
    else:
        classes = {
            "column": BarChart,
            "bar": BarChart,
            "line": LineChart,
            "pie": PieChart,
            "doughnut": DoughnutChart,
            "area": AreaChart,
        }
        chart = classes[kind]()
        if kind in ("column", "bar"):
            chart.type = "col" if kind == "column" else "bar"
        last = left + 1 if kind in ("pie", "doughnut") else right
        data_ref = Reference(ws, min_col=left + 1, max_col=last, min_row=top, max_row=bottom)
        chart.add_data(data_ref, titles_from_data=True)
        chart.set_categories(categories)
    chart.width, chart.height = WIDTH_CM, HEIGHT_CM
    return chart


def _refs(chart: Any) -> Iterator[Any]:
    for series in chart.series:
        for part in _SERIES_PARTS:
            holder = getattr(series, part, None)
            for kind in _REF_KINDS:
                ref = getattr(holder, kind, None) if holder is not None else None
                if ref is not None and ref.f:
                    yield ref


def _rewrite(wb: Any, fn: Callable[[str, str], tuple[str, bool]]) -> int:
    changed = 0
    for ws in wb.worksheets:
        for chart in ws._charts:
            for ref in _refs(chart):
                new, did = fn("=" + ref.f, ws.title)
                if did:
                    ref.f = new[1:]
                    changed += 1
    return changed


def shift_charts(wb: Any, sheet: str, shift: Shift) -> int:
    """Shift series ranges that point at `sheet`, and move anchors of charts placed on it."""
    changed = _rewrite(wb, lambda f, home: shift_formula(f, home, sheet, shift))
    for ws in wb.worksheets:
        if ws.title.casefold() != sheet.casefold():
            continue
        for chart in ws._charts:
            changed += _move_anchor(chart, shift)
    return changed


def anchor_cell(obj: Any) -> tuple[int, int] | None:
    """(row, col) of a chart's or image's top-left cell, 0-based; None when it has no anchor."""
    anchor = obj.anchor
    if isinstance(anchor, str):
        rng = parse_range(anchor)
        return rng.top, rng.left
    marker = getattr(anchor, "_from", None)
    return (marker.row, marker.col) if marker is not None else None


def _move_anchor(chart: Any, shift: Shift) -> int:
    anchor = chart.anchor
    if isinstance(anchor, str):
        moved = shift_address(anchor, shift)
        if moved not in (anchor, "#REF!"):
            chart.anchor = moved
            return 1
        return 0
    changed = 0
    for marker in (getattr(anchor, "_from", None), getattr(anchor, "to", None)):
        if marker is None:
            continue
        attr = "row" if shift.axis == "rows" else "col"
        index = shift.index(getattr(marker, attr) + 1)
        new = (index if index is not None else shift.at) - 1
        if new != getattr(marker, attr):
            setattr(marker, attr, new)
            changed = 1
    return changed


def rename_charts(wb: Any, old: str, new: str) -> int:
    return _rewrite(wb, lambda f, _home: rename_in_formula(f, old, new))


def charts_reading(wb: Any, sheet: str) -> int:
    """Charts (on any sheet) whose data comes from `sheet`."""
    return sum(
        1
        for ws in wb.worksheets
        for chart in ws._charts
        if any(rename_in_formula("=" + ref.f, sheet, sheet)[1] for ref in _refs(chart))
    )
