"""Full-tier presentation tools: format_range, set_table, add_validation, conditional_format."""

from __future__ import annotations

import datetime as dt
from typing import Any

from openpyxl.formatting.rule import (
    CellIsRule,
    ColorScaleRule,
    DataBarRule,
    FormulaRule,
    Rule,
)
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.styles.differential import DifferentialStyle
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.table import Table, TableStyleInfo

from slim_sheet.config import Settings
from slim_sheet.core.a1 import format_range as a1_range
from slim_sheet.core.a1 import index_to_col, parse_range
from slim_sheet.core.errors import ErrorCode, SlimError
from slim_sheet.core.models import CellRange, OutputResult
from slim_sheet.core.refs import valid_defined_name
from slim_sheet.core.styles import ALIGNMENTS, BORDERS, choice, color
from slim_sheet.core.styles import number_format as format_preset
from slim_sheet.services.editing import Change, edit_workbook, get_sheet

MAX_STYLED_CELLS = 200_000
MIN_WIDTH, MAX_WIDTH = 8, 60
VALIDATION_TYPES = ("list", "whole", "decimal", "date", "text_length", "custom")
CF_RULES = (
    "greater_than",
    "less_than",
    "between",
    "equal",
    "not_equal",
    "text_contains",
    "duplicates",
    "unique",
    "top_n",
    "bottom_n",
    "above_average",
    "color_scale",
    "data_bar",
    "formula",
)
_CELL_IS = {
    "greater_than": "greaterThan",
    "less_than": "lessThan",
    "between": "between",
    "equal": "equal",
    "not_equal": "notEqual",
}


def _area(range_ref: str) -> CellRange:
    rng = parse_range(range_ref)
    if rng.n_rows * rng.n_cols > MAX_STYLED_CELLS:
        raise SlimError(ErrorCode.TOO_LARGE, f"{range_ref} has more than {MAX_STYLED_CELLS} cells")
    return rng


def _cells(ws: Any, rng: CellRange) -> Any:
    return ws.iter_rows(
        min_row=rng.top + 1, max_row=rng.bottom + 1, min_col=rng.left + 1, max_col=rng.right + 1
    )


# --- format_range -------------------------------------------------------------------------


def format_range(
    path: str,
    sheet: str,
    range: str,
    settings: Settings,
    bold: bool | None = None,
    italic: bool | None = None,
    font_color: str | None = None,
    font_size: float | None = None,
    fill_color: str | None = None,
    number_format: str | None = None,
    align: str | None = None,
    wrap: bool | None = None,
    border: str | None = None,
    autofit: bool = False,
    out_path: str | None = None,
    overwrite: bool = False,
) -> OutputResult:
    rng = _area(range)
    font_rgb = color(font_color) if font_color else None
    fill_rgb = color(fill_color) if fill_color else None
    fmt = format_preset(number_format) if number_format else None
    horizontal = choice(align, ALIGNMENTS, "align") if align else None
    edge = choice(border, BORDERS, "border") if border else None

    def mutate(wb: Any) -> Change:
        ws = get_sheet(wb, sheet)
        applied: list[str] = []
        for row in _cells(ws, rng):
            for cell in row:
                _font(cell, bold, italic, font_rgb, font_size)
                if fill_rgb:
                    cell.fill = PatternFill("solid", fgColor=fill_rgb)
                if fmt:
                    cell.number_format = fmt
                if horizontal or wrap is not None:
                    cell.alignment = Alignment(
                        horizontal=horizontal or cell.alignment.horizontal,
                        wrap_text=wrap if wrap is not None else cell.alignment.wrap_text,
                    )
        if edge:
            _borders(ws, rng, edge)
        if autofit:
            applied.append(f"widths for {_autofit(ws, rng)} column(s)")
        flags = {
            "font": bool(bold or italic or font_rgb or font_size),
            "fill": bool(fill_rgb),
            "number format": bool(fmt),
            "alignment": bool(horizontal) or wrap is not None,
            "border": bool(edge),
        }
        styled = [name for name, on in flags.items() if on]
        what = ", ".join(styled + applied) or "nothing (no style arguments given)"
        return Change(f"formatted {ws.title}!{a1_range(rng)}: {what}")

    note = f"format_range {sheet}!{range}"
    return edit_workbook(path, settings, note, mutate, out_path, overwrite)


def _font(
    cell: Any, bold: bool | None, italic: bool | None, rgb: str | None, size: float | None
) -> None:
    if bold is None and italic is None and rgb is None and size is None:
        return
    old = cell.font
    cell.font = Font(
        name=old.name,
        size=size if size is not None else old.size,
        bold=bold if bold is not None else old.bold,
        italic=italic if italic is not None else old.italic,
        underline=old.underline,
        color=rgb if rgb is not None else old.color,
    )


def _borders(ws: Any, rng: CellRange, edge: str) -> None:
    if edge == "none":
        for row in _cells(ws, rng):
            for cell in row:
                cell.border = Border()
        return
    style = "thin" if edge == "outline" else edge
    side = Side(style=style)
    for r, row in enumerate(_cells(ws, rng)):
        for c, cell in enumerate(row):
            if edge != "outline":
                cell.border = Border(left=side, right=side, top=side, bottom=side)
                continue
            cell.border = Border(
                left=side if c == 0 else cell.border.left,
                right=side if c == rng.n_cols - 1 else cell.border.right,
                top=side if r == 0 else cell.border.top,
                bottom=side if r == rng.n_rows - 1 else cell.border.bottom,
            )


def _autofit(ws: Any, rng: CellRange) -> int:
    widths: dict[int, int] = {}
    for row in _cells(ws, rng):
        for cell in row:
            if cell.value is not None:
                widths[cell.column] = max(widths.get(cell.column, 0), len(str(cell.value)))
    for column, width in widths.items():
        ws.column_dimensions[get_column_letter(column)].width = min(
            MAX_WIDTH, max(MIN_WIDTH, width + 2)
        )
    return len(widths)


# --- set_table ----------------------------------------------------------------------------


def set_table(
    path: str,
    sheet: str,
    range: str,
    name: str,
    settings: Settings,
    style: str = "TableStyleMedium9",
    out_path: str | None = None,
    overwrite: bool = False,
) -> OutputResult:
    rng = _area(range)
    if rng.n_rows < 2:
        raise SlimError(
            ErrorCode.BAD_ARGUMENT, "A table needs a header row and at least one data row"
        )
    if not valid_defined_name(name) or "." in name:
        raise SlimError(
            ErrorCode.BAD_ARGUMENT, f"Invalid table name '{name}'", hint="e.g. name='SalesTable'"
        )

    def mutate(wb: Any) -> Change:
        ws = get_sheet(wb, sheet)
        _check_table_conflicts(wb, ws, rng, name)
        header = next(_cells(ws, CellRange(rng.top, rng.left, rng.top, rng.right)))
        labels = [str(c.value).strip() if c.value is not None else "" for c in header]
        if any(not label for label in labels) or len({lb.casefold() for lb in labels}) != len(
            labels
        ):
            raise SlimError(
                ErrorCode.BAD_ARGUMENT,
                f"Header row {rng.top + 1} needs unique, non-empty labels: {labels}",
                hint="Write header labels first with write(...).",
            )
        for cell, label in zip(header, labels, strict=True):
            cell.value = label  # Excel stores table headers as text
        table = Table(displayName=name, ref=a1_range(rng))
        table.tableStyleInfo = TableStyleInfo(name=style, showRowStripes=True)
        ws.add_table(table)
        return Change(f"table {name} on {ws.title}!{a1_range(rng)} ({rng.n_rows - 1} data rows)")

    return edit_workbook(path, settings, f"set_table {name}", mutate, out_path, overwrite)


def _check_table_conflicts(wb: Any, ws: Any, rng: CellRange, name: str) -> None:
    for other_ws in wb.worksheets:
        for table in other_ws.tables.values():
            if table.name.casefold() == name.casefold():
                raise SlimError(
                    ErrorCode.BAD_ARGUMENT, f"A table named '{table.name}' already exists"
                )
            if other_ws is ws and parse_range(table.ref).intersect(rng) is not None:
                raise SlimError(
                    ErrorCode.BAD_ARGUMENT, f"{a1_range(rng)} overlaps table {table.name}"
                )


# --- add_validation -----------------------------------------------------------------------


def add_validation(
    path: str,
    sheet: str,
    range: str,
    type: str,
    settings: Settings,
    values: list[str] | None = None,
    min: str | None = None,
    max: str | None = None,
    formula: str | None = None,
    message: str | None = None,
    allow_blank: bool = True,
    out_path: str | None = None,
    overwrite: bool = False,
) -> OutputResult:
    rng = _area(range)
    kind = choice(type, VALIDATION_TYPES, "validation type")
    dv = _validation(kind, values, min, max, formula)
    dv.allow_blank = allow_blank
    dv.showErrorMessage = True
    if message:
        dv.error, dv.errorTitle = message, "Invalid value"
        dv.prompt, dv.showInputMessage = message, True

    def mutate(wb: Any) -> Change:
        ws = get_sheet(wb, sheet)
        ws.add_data_validation(dv)
        dv.add(a1_range(rng))
        rule = f"{dv.type} {dv.operator or ''} {dv.formula1 or ''} {dv.formula2 or ''}".split()
        return Change(f"validation on {ws.title}!{a1_range(rng)}: {' '.join(rule)}")

    return edit_workbook(
        path, settings, f"add_validation {sheet}!{range}", mutate, out_path, overwrite
    )


def _validation(
    kind: str, values: list[str] | None, low: str | None, high: str | None, formula: str | None
) -> DataValidation:
    if kind == "list":
        if formula:
            return DataValidation(type="list", formula1=formula.lstrip("="))
        if not values:
            raise SlimError(
                ErrorCode.BAD_ARGUMENT,
                "list validation needs values=[...] or formula='=Lists!A1:A9'",
            )
        joined = ",".join(v.replace('"', "") for v in values)
        if len(joined) > 255:
            raise SlimError(
                ErrorCode.TOO_LARGE,
                "List is over 255 characters",
                hint="Put the options in cells and pass formula.",
            )
        return DataValidation(type="list", formula1=f'"{joined}"')
    if kind == "custom":
        if not formula:
            raise SlimError(
                ErrorCode.BAD_ARGUMENT, "custom validation needs formula, e.g. '=LEN(A1)<=10'"
            )
        return DataValidation(type="custom", formula1=formula.lstrip("="))
    if low is None and high is None:
        raise SlimError(ErrorCode.BAD_ARGUMENT, f"{kind} validation needs min and/or max")
    excel_type = "textLength" if kind == "text_length" else kind
    first, second = _bound(kind, low), _bound(kind, high)
    if first is not None and second is not None:
        return DataValidation(type=excel_type, operator="between", formula1=first, formula2=second)
    if first is not None:
        return DataValidation(type=excel_type, operator="greaterThanOrEqual", formula1=first)
    return DataValidation(type=excel_type, operator="lessThanOrEqual", formula1=second)


def _bound(kind: str, value: str | None) -> str | None:
    if value is None:
        return None
    if kind == "date":
        try:
            day = dt.date.fromisoformat(str(value))
        except ValueError as exc:
            raise SlimError(
                ErrorCode.BAD_ARGUMENT, f"Date bound '{value}' must be YYYY-MM-DD"
            ) from exc
        return f"DATE({day.year},{day.month},{day.day})"
    try:
        float(value)
    except ValueError as exc:
        raise SlimError(ErrorCode.BAD_ARGUMENT, f"Bound '{value}' must be a number") from exc
    return str(value)


# --- conditional_format -------------------------------------------------------------------


def conditional_format(
    path: str,
    sheet: str,
    range: str,
    rule: str,
    settings: Settings,
    value: str | None = None,
    value2: str | None = None,
    fill_color: str = "light_red",
    font_color: str = "9C0006",
    out_path: str | None = None,
    overwrite: bool = False,
) -> OutputResult:
    rng = _area(range)
    kind = choice(rule, CF_RULES, "rule")
    built = _cf_rule(kind, rng, value, value2, color(fill_color), color(font_color))

    def mutate(wb: Any) -> Change:
        ws = get_sheet(wb, sheet)
        ws.conditional_formatting.add(a1_range(rng), built)
        detail = f" {value}" + (f"..{value2}" if value2 else "") if value else ""
        return Change(f"conditional format on {ws.title}!{a1_range(rng)}: {kind}{detail}")

    return edit_workbook(
        path, settings, f"conditional_format {sheet}!{range}", mutate, out_path, overwrite
    )


def _need(value: str | None, kind: str) -> str:
    if value is None or str(value).strip() == "":
        raise SlimError(
            ErrorCode.BAD_ARGUMENT, f"rule '{kind}' needs value", hint="e.g. value='100'"
        )
    return str(value).lstrip("=")


def _cf_rule(
    kind: str, rng: CellRange, value: str | None, value2: str | None, fill: str, font: str
) -> Rule:
    style = DifferentialStyle(fill=PatternFill(bgColor=fill), font=Font(color=font))
    fill_obj = PatternFill(start_color=fill, end_color=fill, fill_type="solid")
    top_left = f"{index_to_col(rng.left)}{rng.top + 1}"
    if kind in _CELL_IS:
        args = [_need(value, kind)] + ([_need(value2, kind)] if kind == "between" else [])
        return CellIsRule(
            operator=_CELL_IS[kind], formula=args, fill=fill_obj, font=Font(color=font)
        )
    if kind == "text_contains":
        needle = _need(value, kind).replace('"', '""')
        return FormulaRule(formula=[f'ISNUMBER(SEARCH("{needle}",{top_left}))'], fill=fill_obj)
    if kind == "formula":
        return FormulaRule(formula=[_need(value, kind)], fill=fill_obj, font=Font(color=font))
    if kind in ("duplicates", "unique"):
        return Rule(type="duplicateValues" if kind == "duplicates" else "uniqueValues", dxf=style)
    if kind in ("top_n", "bottom_n"):
        rank = int(_need(value, kind))
        return Rule(type="top10", rank=rank, bottom=kind == "bottom_n", dxf=style)
    if kind == "above_average":
        return Rule(type="aboveAverage", aboveAverage=True, dxf=style)
    if kind == "color_scale":
        return ColorScaleRule(
            start_type="min", start_color="FFF8696B", mid_type="percentile", mid_value=50,
            mid_color="FFFFEB84", end_type="max", end_color="FF63BE7B",
        )  # fmt: skip
    return DataBarRule(start_type="min", end_type="max", color="FF638EC6")
