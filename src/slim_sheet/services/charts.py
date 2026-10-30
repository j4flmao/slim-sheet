"""`add_chart` use-case: a native Excel chart from a block of cells (PLAN §5.3, §16.4).

Layout convention (what agents and humans expect): the first row holds series names and
the first column holds categories (x values for scatter); every other column is a series.
openpyxl keeps charts across later edits; rows_ops/sheet_ops rewrite their ranges.
Every chart is styled from the brand (no chart junk, visible axes); type='auto' picks the
chart from the data shape.
"""

from __future__ import annotations

from typing import Any

from slim_sheet.adapters.brand import brand_for, resolve_brand
from slim_sheet.adapters.writers.chart_refs import anchor_cell, build_chart
from slim_sheet.adapters.writers.chart_style import ChartText, style_chart
from slim_sheet.adapters.writers.sparklines import line_group
from slim_sheet.config import Settings
from slim_sheet.core.a1 import index_to_col, parse_range
from slim_sheet.core.contrast import faint_series
from slim_sheet.core.design import Brand, formats
from slim_sheet.core.errors import ErrorCode, SlimError
from slim_sheet.core.models import CellRange, OutputResult
from slim_sheet.core.refs import split_sheet
from slim_sheet.core.report_spec import choose_chart, infer_format
from slim_sheet.core.shift import quote_sheet
from slim_sheet.core.styles import choice
from slim_sheet.services.editing import Change, edit_workbook, get_sheet

MAX_ROUND_PARTS, MAX_PAIRED_SERIES, MAX_LINES = 5, 2, 4
CHART_TYPES = ("auto", "column", "bar", "line", "pie", "doughnut", "area", "scatter", "sparkline")


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
    highlight: str | None = None,
    brand: str | None = None,
    replace: bool = False,
) -> OutputResult:
    """`replace=true` first removes the chart(s) whose top-left cell is the anchor (re-runs)."""
    kind = choice(type, CHART_TYPES, "chart type")
    data_sheet, address = split_sheet(data_range)
    rng = parse_range(address)
    if kind == "sparkline":
        return _sparklines(
            path, sheet, data_sheet, rng, settings, anchor, brand, out_path, overwrite
        )
    if rng.n_rows < 2 or rng.n_cols < 2:
        raise SlimError(
            ErrorCode.BAD_ARGUMENT,
            f"data_range {address} needs a header row and a category column plus >= 1 series",
            hint="e.g. data_range='A1:C13' (A = months, B..C = series)",
        )
    if anchor is not None:
        parse_range(anchor)
    chosen_look, explicit = resolve_brand(brand, settings)

    def mutate(wb: Any) -> Change:
        look = brand_for(wb, chosen_look, explicit)
        host = get_sheet(wb, sheet)
        source = get_sheet(wb, data_sheet) if data_sheet else host
        categories = _column(source, rng, rng.left)
        chosen, why = _resolve_kind(kind, source, rng, categories)
        place = anchor or f"{index_to_col(rng.right + 2)}{rng.top + 1}"
        chart = build_chart(chosen, source, address)
        _check_series(source, rng)
        series_values = _column(source, rng, rng.left + 1)
        text = ChartText(
            title=title,
            highlight=_highlight_index(categories, highlight),
            value_format=_number_format(source, rng, rng.left + 1, look),
            x_format=_number_format(source, rng, rng.left, look) if chosen == "scatter" else None,
            x_title=x_title,
            y_title=y_title,
            values=tuple(_numbers(series_values)),
        )
        style_chart(chart, chosen, look, text)
        removed = _remove_charts_at(host, place) if replace else 0
        host.add_chart(chart, place)
        series = 1 if chosen in ("pie", "doughnut") else rng.n_cols - 1
        summary = (
            f"{chosen} chart of {series} series from {source.title}!{address} "
            f"at {host.title}!{place}"
        )
        summary += f" (auto: {why})" if why else ""
        if replace:
            summary += (
                f"; replaced {removed} chart(s)" if removed else "; no chart there to replace"
            )
        return Change(summary, tuple(chart_warnings(chosen, series, rng.n_rows - 1, look)))

    return edit_workbook(path, settings, f"add_chart {kind} {sheet}", mutate, out_path, overwrite)


def chart_warnings(kind: str, series: int, points: int, brand: Brand) -> list[str]:
    """Readability limits the chart breaks (docs/prompts/design.md §4), as teaching hints."""
    found: list[str] = []
    if kind in ("pie", "doughnut") and points > MAX_ROUND_PARTS:
        found.append(
            f"{points} slices: more than {MAX_ROUND_PARTS} are hard to compare; use type='bar' "
            "(sorted) or keep the top parts and sum the rest as 'Other'"
        )
    if kind in ("column", "bar") and series > MAX_PAIRED_SERIES:
        found.append(
            f"{series} series side by side: more than {MAX_PAIRED_SERIES} are hard to read; "
            "chart the change instead, or one chart per series"
        )
    if kind in ("line", "area", "scatter") and series > MAX_LINES:
        found.append(
            f"{series} lines: more than {MAX_LINES} tangle; highlight one and grey the rest, "
            "or split into small charts"
        )
    faint = faint_series(brand.palette, series if kind not in ("pie", "doughnut") else points)
    if faint:
        found.append(
            f"brand colour(s) {', '.join(faint)} have under 3:1 contrast on white (WCAG 1.4.11); "
            "label the series directly or use brand='slim'"
        )
    return found


def _column(ws: Any, rng: CellRange, col: int) -> list[object]:
    return [ws.cell(row=r + 1, column=col + 1).value for r in range(rng.top + 1, rng.bottom + 1)]


def _check_series(ws: Any, rng: CellRange) -> None:
    """A text column plotted as a series draws nothing but a legend entry: refuse it."""
    for col in range(rng.left + 1, rng.right + 1):
        values = [v for v in _column(ws, rng, col) if v not in (None, "")]
        plottable = [
            v
            for v in values
            if (isinstance(v, int | float) and not isinstance(v, bool))
            or (isinstance(v, str) and v.startswith("="))
        ]
        if values and not plottable:
            header = ws.cell(row=rng.top + 1, column=col + 1).value
            raise SlimError(
                ErrorCode.BAD_ARGUMENT,
                f"Column {index_to_col(col)} ('{header}') holds text, not numbers to plot",
                hint="The first column holds the categories, every other column one numeric "
                "series: move or leave out the text column (e.g. a range starting one column "
                "to the right).",
            )


def _numbers(values: list[object]) -> list[float]:
    return [
        float(v) if isinstance(v, int | float) and not isinstance(v, bool) else 0.0 for v in values
    ]


def _number_format(ws: Any, rng: CellRange, col: int, brand: Brand) -> str | None:
    """The column's own number format, else one inferred from its header and values, so
    axes and labels show 27.2% and 1,810 rather than 0.272 and 1810.0012."""
    for row in range(rng.top + 1, rng.bottom + 1):
        cell = ws.cell(row=row + 1, column=col + 1)
        if isinstance(cell.value, int | float) and not isinstance(cell.value, bool):
            if cell.number_format and cell.number_format != "General":
                return str(cell.number_format)
            break
    header = str(ws.cell(row=rng.top + 1, column=col + 1).value or "")
    key, _why = infer_format(header, _column(ws, rng, col))
    return formats(brand)[key] if key and key not in ("date", "year", "text") else None


def _resolve_kind(kind: str, ws: Any, rng: CellRange, categories: list[object]) -> tuple[str, str]:
    if kind != "auto":
        return kind, ""
    values = _numbers(_column(ws, rng, rng.left + 1))
    picked = choose_chart("auto", categories, values)
    return picked.kind, picked.reason


def _highlight_index(categories: list[object], highlight: str | None) -> int | None:
    if not highlight:
        return None
    wanted = highlight.strip().casefold()
    for index, value in enumerate(categories):
        if str(value).strip().casefold() == wanted:
            return index
    shown = ", ".join(str(c) for c in categories[:10])
    raise SlimError(
        ErrorCode.BAD_ARGUMENT,
        f"highlight '{highlight}' is not one of the chart's categories",
        hint=f"Categories: {shown}",
    )


def _sparklines(
    path: str,
    sheet: str,
    data_sheet: str | None,
    rng: CellRange,
    settings: Settings,
    anchor: str | None,
    brand: str | None,
    out_path: str | None,
    overwrite: bool,
) -> OutputResult:
    """One line sparkline per row of a numbers block, in a column (default: right of it)."""
    if rng.n_cols < 2:
        raise SlimError(
            ErrorCode.BAD_ARGUMENT,
            "A sparkline needs at least two values per row",
            hint="e.g. data_range='B2:M10' (12 months per row), anchor='N2'",
        )
    start = (
        parse_range(anchor) if anchor else CellRange(rng.top, rng.right + 1, rng.top, rng.right + 1)
    )
    chosen_look, explicit = resolve_brand(brand, settings)

    def mutate(wb: Any) -> Change:
        look = brand_for(wb, chosen_look, explicit)
        host = get_sheet(wb, sheet)
        source = get_sheet(wb, data_sheet) if data_sheet else host
        prefix = quote_sheet(source.title)
        pairs = [
            (
                f"{prefix}!{index_to_col(rng.left)}{r + 1}:{index_to_col(rng.right)}{r + 1}",
                f"{index_to_col(start.left)}{start.top + 1 + (r - rng.top)}",
            )
            for r in range(rng.top, rng.bottom + 1)
        ]
        group = line_group(pairs, look.primary, look.bad)
        where = f"{host.title}!{pairs[0][1]}:{pairs[-1][1]}"
        return Change(f"{len(pairs)} sparkline(s) at {where}", sparklines=((host.title, group),))

    return edit_workbook(
        path, settings, f"add_chart sparkline {sheet}", mutate, out_path, overwrite
    )


def _remove_charts_at(ws: Any, place: str) -> int:
    cell = parse_range(place)
    keep = [c for c in ws._charts if anchor_cell(c) != (cell.top, cell.left)]
    removed = len(ws._charts) - len(keep)
    ws._charts[:] = keep
    return removed
