"""Building blocks of a `build_report` dashboard beyond tiles and charts (PLAN §16.3).

- ReportData: the source table as the report sees it (column lookup, absolute refs)
- kpi_comparison: a KPI note that states the change against a prior column, live and coloured
- write_messages: the 'Key messages' lines under the KPI row (the report's SAY rule)
- table_slot: a ranked top-N table with live SUMIFS formulas, in a chart's grid slot
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import Any

from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Font

from slim_sheet.adapters.writers.theme import style_cell
from slim_sheet.core.a1 import index_to_col
from slim_sheet.core.design import Brand, RoleStyle, formats, role_style
from slim_sheet.core.errors import ErrorCode, SlimError
from slim_sheet.core.file_formula import to_file
from slim_sheet.core.report_spec import (
    ChartSpec,
    KpiSpec,
    aggregate,
    infer_format,
    order,
    reduce_values,
)
from slim_sheet.core.shift import quote_sheet
from slim_sheet.services.design import TableBlock, column_values

MAX_MESSAGES, MAX_MESSAGE_CHARS = 5, 200
TABLE_ROWS = 10
_LOWER_IS_BETTER = {"cost", "costs", "expense", "expenses", "churn", "debt", "loss", "losses",
                    "spend", "spending", "defects", "returns", "complaints", "overdue", "opex",
                    "attrition", "leavers", "lead"}  # fmt: skip
_AGG_FUNCS = {"sum": "SUM", "avg": "AVERAGE", "min": "MIN", "max": "MAX"}
_IFS = {"sum": "SUMIFS", "avg": "AVERAGEIFS", "min": "MINIFS", "max": "MAXIFS"}


@dataclass(frozen=True, slots=True)
class ReportData:
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

    def numbers(self, col: int) -> list[float]:
        return [
            float(v)
            for v in self.values(col)
            if isinstance(v, int | float) and not isinstance(v, bool)
        ]


# --- KPI comparison -----------------------------------------------------------------------


@dataclass(frozen=True, slots=True)
class Comparison:
    label: str  # what the KPI is compared with, e.g. 'FY2024' or 'plan'
    prior: str  # Excel expression of the comparison value (no leading '=')
    prior_value: float | None  # the same value computed now, for the agent's reply
    points: bool  # a ratio changes in percentage points, a value in percent
    attainment: bool = False  # vs a plan/target: 107.3% of plan (100% is on plan)


def kpi_comparison(spec: KpiSpec, data: ReportData, plan: bool = False) -> Comparison:
    """The KPI against `spec.vs` (a change) or, with `plan`, against `spec.plan`."""
    column, label = (spec.plan, spec.plan_label) if plan else (spec.vs, spec.vs_label)
    if spec.agg == "ratio":
        top, _, bottom = column.partition("/")
        num, den = data.col(top.strip()), data.col(bottom.strip())
        prior = f"SUM({data.ref(num)})/SUM({data.ref(den)})"
        total = sum(data.numbers(den))
        value = sum(data.numbers(num)) / total if total else None
        return Comparison(label, prior, value, points=True)
    col = data.col(column)
    nums = data.numbers(col)
    prior = f"{_AGG_FUNCS[spec.agg]}({data.ref(col)})"
    value = reduce_values(nums, spec.agg) if nums else None
    return Comparison(label, prior, value, points=False, attainment=plan)


# (points, change, multiple, attainment) phrasing; Excel formats keep the reader's separators
_PHRASES = {
    "en": ('" pts vs {l}"', '" vs {l}"', '" vs {l}"', '" of {l}"', ""),
    "vi": ('" điểm % so với {l}"', '" so với {l}"', '" so với {l}"', '" {l}"', '"Hoàn thành "&'),
}


def comparison_formula(comparison: Comparison, value_cell: str, locale: str = "en") -> str:
    """A live note: '+15.6% vs FY2024', '2.4x vs FY2024', '+1.8 pts vs FY2024' or
    '107.3% of plan' (Vietnamese: 'so với cùng kỳ', 'Hoàn thành 107,3% KH')."""
    label = comparison.label.replace('"', '""')
    points, change, multiple, attained, lead = (
        p.replace("{l}", label) for p in _PHRASES.get(locale, _PHRASES["en"])
    )
    ratio = f"{value_cell}/({comparison.prior})"
    if comparison.points:
        text = f'TEXT(({value_cell}-({comparison.prior}))*100,"+0.0;-0.0;0.0")&{points}'
    elif comparison.attainment:
        text = f'{lead}TEXT({ratio},"0.0%")&{attained}'
    else:  # a jump of 2x or more reads better as a multiple than as +140%
        text = (
            f'IF({ratio}>=2,TEXT({ratio},"0.0""x""")&{multiple},'
            f'TEXT({ratio}-1,"+0.0%;-0.0%;0.0%")&{change})'
        )
    return f'=IFERROR({text},"n/a")'


def joined_notes(first: str, second: str) -> str:
    """Two note formulas in one cell: '+15.6% vs FY2024 · 107.3% of plan'."""
    return f'{first}&" · "&{second.removeprefix("=")}'


def colour_comparison(
    ws: Any, cell: str, value_cell: str, comparison: Comparison, spec: KpiSpec, brand: Brand
) -> None:
    """Green when the KPI moved the good way, red otherwise ('Cost' going up is bad)."""
    words = set(f"{spec.label} {spec.column}".lower().replace("(", " ").split())
    better, worse = (">=", "<") if not words & _LOWER_IS_BETTER else ("<=", ">")
    for sign, colour in ((better, brand.good), (worse, brand.bad)):
        rule = FormulaRule(
            formula=[f"AND(ISNUMBER({value_cell}),{value_cell}{sign}({comparison.prior}))"],
            font=Font(color=f"FF{colour}", bold=True),
        )
        ws.conditional_formatting.add(cell, rule)


def describe_change(value: float | None, comparison: Comparison) -> str:
    if value is None or comparison.prior_value is None:
        return ""
    if comparison.points:
        return f", {(value - comparison.prior_value) * 100:+.1f} pts vs {comparison.label}"
    if not comparison.prior_value:
        return ""
    ratio = value / comparison.prior_value
    if comparison.attainment:
        return f", {ratio:.1%} of {comparison.label}"
    if ratio >= 2:
        return f", {ratio:.1f}x {comparison.label}"
    return f", {ratio - 1:+.1%} vs {comparison.label}"


# --- key messages -------------------------------------------------------------------------


def check_messages(messages: list[str]) -> list[str]:
    kept = [m.strip() for m in messages if m and m.strip()]
    if len(kept) > MAX_MESSAGES:
        raise SlimError(
            ErrorCode.TOO_LARGE,
            f"{len(kept)} key messages; a dashboard states at most {MAX_MESSAGES}",
            hint="Keep the 3-5 findings a reader must remember; details go on their own sheet.",
        )
    long = [m for m in kept if len(m) > MAX_MESSAGE_CHARS]
    if long:
        raise SlimError(
            ErrorCode.BAD_ARGUMENT,
            f"A key message is {len(long[0])} characters (max {MAX_MESSAGE_CHARS})",
            hint="One sentence per message: the finding and its number.",
        )
    return kept


def write_messages(ws: Any, messages: list[str], top: int, brand: Brand, heading: str) -> int:
    """Heading with a rule across the grid, one line per message; -> the next free row."""
    ws.row_dimensions[top].height = 20
    for col in range(2, 13):
        style_cell(ws.cell(row=top, column=col), RoleStyle(bottom="thin"), brand)
    ws.cell(row=top, column=2, value=heading)
    style_cell(ws.cell(row=top, column=2), role_style("section", brand), brand)
    for offset, message in enumerate(messages, start=1):
        cell = ws.cell(row=top + offset, column=2, value=f"▪  {message}")
        style_cell(cell, RoleStyle(size=10, color=brand.ink), brand)
        ws.row_dimensions[top + offset].height = 17
    return top + len(messages) + 2


# --- top-N table --------------------------------------------------------------------------


def table_slot(
    ws: Any, spec: ChartSpec, data: ReportData, brand: Brand, left: int, top: int
) -> str:
    """A ranked table in a chart slot: title, header, top rows and a total, all live formulas.
    `left` is the 1-based first grid column (B or H), `top` the slot's first row."""
    category = data.col(spec.category)
    counting = spec.value == "count"
    raw = data.values(category) if counting else data.values(data.col(spec.value))
    points = order(aggregate(list(zip(data.values(category), raw, strict=True)), spec.agg), "desc")
    shown = points[:TABLE_ROWS]
    label = "Rows" if counting else spec.value
    key = "integer" if counting else infer_format(label, [p[1] for p in points])[0] or "number1"
    fmt = formats(brand)[key]
    share = spec.agg in ("sum", "count") and all(p[1] >= 0 for p in points)
    title = spec.title or f"Top {len(shown)} {spec.category} by {label}"
    cols = (left, left + 2, left + 4) if share else (left, left + 2)
    heads = (spec.category, label, "Share") if share else (spec.category, label)
    _table_head(ws, top, cols, title, heads, brand)
    value_of = _value_formula(spec, data, category, counting)
    total = _total_formula(spec, data, category, counting)
    for offset, (name, _value) in enumerate(shown, start=top + 2):
        _table_row(ws, offset, cols, name, value_of, total if share else None, fmt, brand)
    end = top + 2 + len(shown)
    _table_total(ws, end, cols, total, share, fmt, brand, spec.agg)
    where = f"{index_to_col(left - 1)}{top}"
    return f"table {where}: top {len(shown)} of {len(points)} {spec.category} by {label}, '{title}'"


def _table_head(
    ws: Any, top: int, cols: tuple[int, ...], title: str, heads: tuple[str, ...], brand: Brand
) -> None:
    ws.cell(row=top, column=cols[0], value=title)
    style_cell(
        ws.cell(row=top, column=cols[0]), RoleStyle(12, True, brand.ink, font="heading"), brand
    )
    for col, text in zip(cols, heads, strict=True):
        cell = ws.cell(row=top + 1, column=col, value=text)
        style_cell(cell, role_style("header", brand), brand)
        if col != cols[0]:
            style_cell(cell, RoleStyle(align="right"), brand)
    style_cell(ws.cell(row=top + 1, column=cols[0] + 1), role_style("header", brand), brand)


def _value_formula(spec: ChartSpec, data: ReportData, category: int, counting: bool) -> str:
    """Formula template for one category's value; '{key}' is replaced by its cell."""
    keys = data.ref(category)
    if counting:
        return f"=COUNTIFS({keys},{{key}})"
    values = data.ref(data.col(spec.value))
    return f"={_IFS[spec.agg]}({values},{keys},{{key}})"


def _total_formula(spec: ChartSpec, data: ReportData, category: int, counting: bool) -> str:
    if counting:
        return f"COUNTA({data.ref(category)})"
    return f"{_AGG_FUNCS[spec.agg]}({data.ref(data.col(spec.value))})"


def _table_row(
    ws: Any,
    row: int,
    cols: tuple[int, ...],
    name: object,
    value_of: str,
    total: str | None,
    fmt: str,
    brand: Brand,
) -> None:
    name_cell = ws.cell(row=row, column=cols[0], value=name)
    key = f"${index_to_col(cols[0] - 1)}${row}"
    value_cell = ws.cell(row=row, column=cols[1], value=to_file(value_of.replace("{key}", key)))
    value_cell.number_format = fmt
    cells = [name_cell, ws.cell(row=row, column=cols[0] + 1), value_cell]  # names overflow
    if total:
        value_ref = f"${index_to_col(cols[1] - 1)}${row}"
        share_cell = ws.cell(row=row, column=cols[2], value=f"=IFERROR({value_ref}/{total},0)")
        share_cell.number_format = "0.0%"
        cells.append(share_cell)
    for cell in cells:
        style_cell(cell, RoleStyle(size=10, color=brand.ink, bottom="hair"), brand)


def _table_total(
    ws: Any,
    row: int,
    cols: tuple[int, ...],
    total: str,
    share: bool,
    fmt: str,
    brand: Brand,
    agg: str,
) -> None:
    word = {"sum": "Total", "count": "Total", "avg": "Average", "min": "Lowest", "max": "Highest"}
    ws.cell(row=row, column=cols[0], value=f"{word[agg]} (all)")
    ws.cell(row=row, column=cols[1], value=f"={total}").number_format = fmt
    if share:
        ws.cell(row=row, column=cols[2], value=1).number_format = "0.0%"
    for col in (cols[0] + 1, *cols):  # the gutter after the name carries the rule too
        style_cell(ws.cell(row=row, column=col), role_style("total", brand), brand)
