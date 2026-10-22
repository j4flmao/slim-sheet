"""`build_report` use-case: a branded one-screen dashboard from compact strings (PLAN §16.3).

The model says *what* to show ("Revenue=sum:Total", "rank:Region:Total"); the server decides
*how*: the layout grid, KPI tiles with live formulas, the chart type, sort order, colours, an
insight title for every chart, number formats, and a source footer. Chart numbers are
aggregated once into a '<sheet> data' sheet that the charts read.
"""

from __future__ import annotations

import datetime as dt
from dataclasses import dataclass
from typing import Any

from openpyxl.styles import PatternFill
from openpyxl.worksheet.properties import PageSetupProperties

from slim_sheet.adapters.brand import brand_for, resolve_brand
from slim_sheet.adapters.writers.chart_refs import build_chart
from slim_sheet.adapters.writers.chart_style import style_chart
from slim_sheet.adapters.writers.theme import apply_brand, style_cell
from slim_sheet.config import Settings
from slim_sheet.core.a1 import index_to_col
from slim_sheet.core.design import FORMAT_NAMES, Brand, formats, role_style, scaled_format
from slim_sheet.core.errors import ErrorCode, SlimError
from slim_sheet.core.guard import check_formula
from slim_sheet.core.models import OutputResult
from slim_sheet.core.report_spec import (
    ChartSpec,
    KpiSpec,
    aggregate,
    choose_chart,
    infer_format,
    insight_title,
    order,
    parse_chart,
    parse_kpi,
    reduce_values,
)
from slim_sheet.core.shift import quote_sheet
from slim_sheet.services.design import TableBlock, column_values, find_table, style_block
from slim_sheet.services.editing import Change, edit_workbook, get_sheet

MAX_KPIS, MAX_CHARTS = 6, 6
MAX_BARS = 12  # longer rankings keep the top 11 and fold the rest into "Other"
TILE_WIDTH, GUTTER, MARGIN = 22, 2, 2
CHART_WIDTH_CM, CHART_HEIGHT_CM, CHART_ROWS, CHART_TOP = 13.6, 7.6, 16, 8
ROW_HEIGHTS = {1: 32, 2: 18, 3: 8, 4: 20, 5: 36, 6: 18, 7: 12}


@dataclass(frozen=True, slots=True)
class _Data:
    """The source table as the report sees it."""

    ws: Any
    block: TableBlock
    first: int  # 1-based first and last data rows
    last: int

    def col(self, name: str) -> int:
        wanted = name.strip().casefold()
        for offset, header in enumerate(self.block.headers):
            if header.casefold() == wanted:
                return self.block.rng.left + offset
        raise SlimError(
            ErrorCode.BAD_ARGUMENT,
            f"No column '{name}' in {self.ws.title}",
            hint=f"Columns: {', '.join(h for h in self.block.headers if h)}",
        )

    def ref(self, col: int) -> str:
        letter = index_to_col(col)
        return f"{quote_sheet(self.ws.title)}!${letter}${self.first}:${letter}${self.last}"

    def values(self, col: int) -> list[object]:
        return column_values(self.ws, self.block, col)


def build_report(
    path: str,
    data_sheet: str,
    title: str,
    settings: Settings,
    kpis: list[str] | None = None,
    charts: list[str] | None = None,
    subtitle: str | None = None,
    sheet: str = "Dashboard",
    brand: str | None = None,
    out_path: str | None = None,
    overwrite: bool = False,
) -> OutputResult:
    kpi_specs = [parse_kpi(k) for k in kpis or []]
    chart_specs = [parse_chart(c) for c in charts or []]
    _check_counts(kpi_specs, chart_specs, sheet, data_sheet)
    chosen, explicit = resolve_brand(brand, settings)

    def mutate(wb: Any) -> Change:
        look = brand_for(wb, chosen, explicit)
        data = _source(wb, data_sheet)
        style_block(data.ws, data.block, "data", look, "auto")
        board = _fresh_sheet(wb, sheet, 0)
        helper = _fresh_sheet(wb, f"{sheet} data"[:31], 1)
        _layout(board, look, title, subtitle or _default_subtitle(data))
        lines = [_tile(board, i, spec, data, look) for i, spec in enumerate(kpi_specs)]
        row = 1
        for i, spec in enumerate(chart_specs):
            line, row = _chart(board, helper, i, spec, data, look, row)
            lines.append(line)
        _footer(board, look, data, len(chart_specs))
        apply_brand(wb, look, force=explicit)
        wb.active = 0
        rows = data.last - data.first + 1
        head = f"built {board.title} (brand {look.name}) from {data.ws.title} ({rows} rows)"
        return Change(head + "".join(f"\n- {line}" for line in lines))

    return edit_workbook(path, settings, f"build_report {sheet}", mutate, out_path, overwrite)


def _check_counts(kpis: list[KpiSpec], charts: list[ChartSpec], sheet: str, data: str) -> None:
    if not kpis and not charts:
        raise SlimError(
            ErrorCode.BAD_ARGUMENT,
            "build_report needs at least one KPI or chart",
            hint="e.g. kpis=['Revenue=sum:Total', 'Orders=count'], charts=['rank:Region:Total']",
        )
    if len(kpis) > MAX_KPIS or len(charts) > MAX_CHARTS:
        raise SlimError(
            ErrorCode.TOO_LARGE,
            f"A one-screen dashboard holds at most {MAX_KPIS} KPIs and {MAX_CHARTS} charts",
            hint="Keep the most important ones; put the rest in a second build_report sheet.",
        )
    if sheet.strip().casefold() in (data.strip().casefold(), f"{data} data".casefold()):
        raise SlimError(
            ErrorCode.BAD_ARGUMENT,
            f"sheet '{sheet}' would replace the data sheet",
            hint="e.g. sheet='Dashboard'",
        )


def _source(wb: Any, name: str) -> _Data:
    ws = get_sheet(wb, name)
    block = find_table(ws, None, None)
    first = block.rng.top + 2
    # 1-based last data row: the row above a total row (0-based total_row), else the bottom
    last = block.total_row if block.total_row is not None else block.rng.bottom + 1
    if last < first:
        raise SlimError(ErrorCode.BAD_ARGUMENT, f"Sheet '{ws.title}' has a header but no data rows")
    return _Data(ws, block, first, last)


def _fresh_sheet(wb: Any, name: str, index: int) -> Any:
    for title in wb.sheetnames:
        if title.casefold() == name.casefold():
            wb.remove(wb[title])
    return wb.create_sheet(name, index)


def _default_subtitle(data: _Data) -> str:
    today = dt.date.today().isoformat()
    return f"Source: {data.ws.title} ({data.last - data.first + 1} rows) · Built {today}"


def _layout(ws: Any, brand: Brand, title: str, subtitle: str) -> None:
    ws.column_dimensions["A"].width = MARGIN
    for i in range(MAX_KPIS):
        ws.column_dimensions[index_to_col(1 + 2 * i)].width = TILE_WIDTH
        ws.column_dimensions[index_to_col(2 + 2 * i)].width = GUTTER
    for row, height in ROW_HEIGHTS.items():
        ws.row_dimensions[row].height = height
    ws["B1"], ws["B2"] = title, subtitle
    style_cell(ws["B1"], role_style("title", brand), brand)
    style_cell(ws["B2"], role_style("subtitle", brand), brand)
    view = ws.sheet_view
    view.showGridLines, view.zoomScale = False, 100
    ws.sheet_properties.tabColor = f"FF{brand.primary}"
    ws.page_setup.orientation = "landscape"
    ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
    ws.page_setup.fitToWidth, ws.page_setup.fitToHeight = 1, 1  # one screen, one page


# --- KPI tiles ----------------------------------------------------------------------------


def _tile(ws: Any, index: int, spec: KpiSpec, data: _Data, brand: Brand) -> str:
    formula, value, note, key = _kpi_parts(spec, data)
    fmt = formats(brand)[scaled_format(key, value or 0.0)]
    col = index_to_col(1 + 2 * index)
    fill = PatternFill("solid", fgColor=f"FF{brand.band}")
    for row, content, role in (
        (4, spec.label, "kpi_label"),
        (5, formula, "kpi_value"),
        (6, note, "kpi_note"),
    ):
        cell = ws[f"{col}{row}"]
        cell.value = content
        style_cell(cell, role_style(role, brand), brand, fmt if row == 5 else None)
        cell.fill = fill
    shown = "computed in Excel" if value is None else _plain(value)
    return f"KPI {spec.label} = {shown} ({formula}) at {col}5"


def _plain(value: float) -> str:
    """A number for the agent to quote: 7,083,120 / 1,234.5 / 0.2042 (never 7.08e+06)."""
    if abs(value) >= 1000 or value.is_integer():
        return f"{value:,.0f}"
    return f"{value:,.4f}".rstrip("0").rstrip(".") if abs(value) < 1 else f"{value:,.1f}"


def _kpi_parts(spec: KpiSpec, data: _Data) -> tuple[str, float | None, str, str]:
    """-> (formula, value now, description, format key)."""
    if spec.agg == "formula":
        check_formula(spec.formula, f"KPI {spec.label}")
        return spec.formula, None, "Custom formula", _fmt_key(spec.fmt, "integer")
    if spec.agg == "ratio":
        top, bottom = data.col(spec.column), data.col(spec.column2)
        formula = f'=IFERROR(SUM({data.ref(top)})/SUM({data.ref(bottom)}),"n/a")'
        num, den = _total(data.values(top)), _total(data.values(bottom))
        value = num / den if den else None
        return formula, value, f"{spec.column} / {spec.column2}", _fmt_key(spec.fmt, "percent")
    if spec.agg == "count":
        col = data.col(spec.column) if spec.column else data.block.rng.left
        rows = sum(1 for v in data.values(col) if v not in (None, ""))
        what = f"Count of {spec.column}" if spec.column else "Rows"
        return f"=COUNTA({data.ref(col)})", float(rows), what, _fmt_key(spec.fmt, "integer")
    col = data.col(spec.column)
    values = data.values(col)
    nums = [float(v) for v in values if isinstance(v, int | float) and not isinstance(v, bool)]
    func = {"sum": "SUM", "avg": "AVERAGE", "min": "MIN", "max": "MAX"}[spec.agg]
    value = reduce_values(nums, spec.agg) if nums else None
    inferred = infer_format(spec.column, values)[0] or "integer"
    if spec.agg == "avg" and inferred == "integer" and value is not None and abs(value) < 100:
        inferred = "number1"  # an average of whole numbers is rarely whole
    word = {"sum": "Total", "avg": "Average", "min": "Lowest", "max": "Highest"}[spec.agg]
    formula = f"={func}({data.ref(col)})"
    return formula, value, f"{word} {spec.column}", _fmt_key(spec.fmt, inferred)


def _total(values: list[object]) -> float:
    return sum(float(v) for v in values if isinstance(v, int | float) and not isinstance(v, bool))


def _fmt_key(requested: str, fallback: str) -> str:
    key = requested.strip().lower()
    if not key:
        return fallback
    if key not in FORMAT_NAMES:
        raise SlimError(
            ErrorCode.BAD_ARGUMENT,
            f"Unknown KPI format '{requested}'",
            hint=f"Use one of: {', '.join(FORMAT_NAMES)}, e.g. 'Revenue=sum:Total|currency'",
        )
    return key


# --- charts -------------------------------------------------------------------------------


def _chart(
    board: Any, helper: Any, index: int, spec: ChartSpec, data: _Data, brand: Brand, row: int
) -> tuple[str, int]:
    """Aggregate, write the block on the helper sheet, add a styled chart. -> (line, next row)."""
    category = data.col(spec.category)
    label = "Rows" if spec.value == "count" else spec.value
    if spec.intent == "relationship":
        points = _scatter_points(data, category, data.col(spec.value))
    else:
        raw = data.values(data.col(spec.value)) if spec.value != "count" else data.values(category)
        points = aggregate(list(zip(data.values(category), raw, strict=True)), spec.agg)
    choice_ = choose_chart(spec.intent, [p[0] for p in points], [p[1] for p in points])
    points = _fold(order(points, choice_.sort), choice_.kind, spec.agg)
    intent = (
        spec.intent if spec.intent != "auto" else ("trend" if choice_.sort == "asc" else "compare")
    )
    default = f"{label} vs {spec.category}" if intent == "relationship" else None
    title = spec.title or default or insight_title(intent, label, points)
    key = "integer" if spec.agg == "count" else infer_format(label, [p[1] for p in points])[0]
    fmt = formats(brand)[
        scaled_format(key or "integer", max((abs(p[1]) for p in points), default=0))
    ]
    address = _write_block(helper, row, title, (spec.category, label), points, fmt, brand)
    chart = build_chart(choice_.kind, helper, address)
    highlight = 0 if intent in ("compare", "rank") and choice_.kind in ("bar", "column") else None
    style_chart(chart, choice_.kind, brand, title=title, highlight=highlight, value_format=fmt)
    chart.width, chart.height = CHART_WIDTH_CM, CHART_HEIGHT_CM
    anchor = f"{'B' if index % 2 == 0 else 'H'}{CHART_TOP + (index // 2) * CHART_ROWS}"
    board.add_chart(chart, anchor)
    line = (
        f"chart {anchor}: {choice_.kind} ({choice_.reason}), {len(points)} points, title '{title}'"
    )
    return line, row + len(points) + 4


def _scatter_points(data: _Data, x_col: int, y_col: int) -> list[tuple[object, float]]:
    points: list[tuple[object, float]] = []
    for x, y in zip(data.values(x_col), data.values(y_col), strict=True):
        if isinstance(x, int | float) and isinstance(y, int | float) and not isinstance(y, bool):
            points.append((x, float(y)))
    return points[:500]


def _fold(points: list[tuple[object, float]], kind: str, agg: str) -> list[tuple[object, float]]:
    if kind not in ("bar", "column") or len(points) <= MAX_BARS or agg not in ("sum", "count"):
        return points[:MAX_BARS] if kind in ("bar", "column") else points
    head = points[: MAX_BARS - 1]
    return [*head, ("Other", sum(p[1] for p in points[MAX_BARS - 1 :]))]


def _write_block(
    ws: Any,
    row: int,
    title: str,
    headers: tuple[str, str],
    points: list[tuple[object, float]],
    fmt: str,
    brand: Brand,
) -> str:
    ws.column_dimensions["A"].width, ws.column_dimensions["B"].width = 28, 16
    ws.cell(row=row, column=1, value=title)
    style_cell(ws.cell(row=row, column=1), role_style("note", brand), brand)
    for col, text in enumerate(headers, start=1):
        cell = ws.cell(row=row + 1, column=col, value=text)
        style_cell(cell, role_style("header", brand), brand)
    for offset, (category, value) in enumerate(points, start=row + 2):
        ws.cell(row=offset, column=1, value=category)
        ws.cell(row=offset, column=2, value=value).number_format = fmt
    return f"A{row + 1}:B{row + 1 + len(points)}"


def _footer(ws: Any, brand: Brand, data: _Data, n_charts: int) -> None:
    row = CHART_TOP + ((n_charts + 1) // 2) * CHART_ROWS + (1 if n_charts else 0)
    text = (
        f"Source: {data.ws.title}!{index_to_col(data.block.rng.left)}{data.first - 1}:"
        f"{index_to_col(data.block.rng.right)}{data.last}. KPI tiles are live formulas; chart "
        f"numbers are in '{ws.title} data'. Made with slim-sheet."
    )
    ws.cell(row=row, column=2, value=text)
    style_cell(ws.cell(row=row, column=2), role_style("note", brand), brand)
