"""Full-tier presentation tools: format_range, set_table, add_validation, conditional_format."""

from __future__ import annotations

import datetime as dt
import re
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
from openpyxl.worksheet.filters import AutoFilter
from openpyxl.worksheet.table import Table, TableColumn, TableStyleInfo

from slim_sheet.adapters.brand import resolve_brand
from slim_sheet.adapters.writers.theme import style_cell
from slim_sheet.config import Settings
from slim_sheet.core.a1 import format_range as a1_range
from slim_sheet.core.a1 import index_to_col, parse_range
from slim_sheet.core.design import FORMAT_NAMES, Brand, formats, role_style
from slim_sheet.core.errors import ErrorCode, SlimError
from slim_sheet.core.guard import check_formula
from slim_sheet.core.models import CellRange, OutputResult
from slim_sheet.core.refs import valid_defined_name
from slim_sheet.core.styles import ALIGNMENTS, BORDERS, VERTICAL_ALIGNMENTS, choice, color
from slim_sheet.core.styles import number_format as format_preset
from slim_sheet.services.editing import Change, edit_workbook, get_sheet

MAX_STYLED_CELLS = 200_000
MIN_WIDTH, MAX_WIDTH = 8, 60
WRAP_WIDTH = 15  # wrapped text: ~15 chars per line, but never narrower than its longest word
BASE_FONT_SIZE = 11.0
# Estimated display widths (chars) for formula results, which an edit load cannot see.
GENERAL_CHARS, PERCENT_CHARS, NUMBER_CHARS = 10, 8, 12  # 12 fits "$1,234,567.89"
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
    valign: str | None = None,
    role: str | None = None,
    merge: bool | None = None,
) -> OutputResult:
    rng = _area(range)
    look = resolve_brand(None, settings)[0] if role or number_format else None
    role_look = role_style(role, look) if role and look else None
    font_rgb = color(font_color) if font_color else None
    fill_rgb = color(fill_color) if fill_color else None
    fmt = _named_format(number_format, look) if number_format else None
    horizontal = _horizontal(align) if align else None
    vertical = _vertical(valign) if valign else None
    edge = choice(border, BORDERS, "border") if border else None

    def mutate(wb: Any) -> Change:
        ws = get_sheet(wb, sheet)
        applied: list[str] = []
        for row in _cells(ws, rng):
            for cell in row:
                if role_look and look:  # the role first, explicit arguments refine it
                    style_cell(cell, role_look, look)
                _font(cell, bold, italic, font_rgb, font_size)
                if fill_rgb:
                    cell.fill = PatternFill("solid", fgColor=fill_rgb)
                if fmt:
                    cell.number_format = fmt
                if horizontal or vertical or wrap is not None:
                    cell.alignment = Alignment(
                        horizontal=horizontal or cell.alignment.horizontal,
                        vertical=vertical or cell.alignment.vertical,
                        wrap_text=wrap if wrap is not None else cell.alignment.wrap_text,
                    )
        if edge:
            _borders(ws, rng, edge)
        if merge is not None:
            applied.append(_merge(ws, rng, merge))
        if autofit:
            applied.append(f"widths for {autofit_columns(ws, rng)} column(s)")
        flags = {
            f"role {role}": bool(role_look),
            "font": bool(bold or italic or font_rgb or font_size),
            "fill": bool(fill_rgb),
            "number format": bool(fmt),
            "alignment": bool(horizontal or vertical) or wrap is not None,
            "border": bool(edge),
        }
        styled = [name for name, on in flags.items() if on]
        what = ", ".join(styled + applied) or "nothing (no style arguments given)"
        return Change(f"formatted {ws.title}!{a1_range(rng)}: {what}")

    note = f"format_range {sheet}!{range}"
    return edit_workbook(path, settings, note, mutate, out_path, overwrite)


def _named_format(name: str, look: Brand | None) -> str:
    """Design names ('percent', 'growth', 'm'...) follow the brand; others use the old presets."""
    key = name.strip().lower()
    if key == "currency" and (look is None or not look.currency):
        return format_preset(name)  # no brand currency: keep the familiar "$"#,##0.00
    if look is not None and key in FORMAT_NAMES:
        return formats(look)[key]
    return format_preset(name)


def _merge(ws: Any, rng: CellRange, merge: bool) -> str:
    """Merge the range (keeping the top-left value) or unmerge every merge inside it."""
    if merge:
        if rng.n_rows * rng.n_cols == 1:
            raise SlimError(ErrorCode.BAD_ARGUMENT, "merge needs at least two cells")
        for existing in list(ws.merged_cells.ranges):
            if parse_range(existing.coord).intersect(rng):
                ws.unmerge_cells(existing.coord)
        ws.merge_cells(a1_range(rng))
        return f"merged {a1_range(rng)}"
    inside = [m.coord for m in ws.merged_cells.ranges if parse_range(m.coord).intersect(rng)]
    for coord in inside:
        ws.unmerge_cells(coord)
    return f"unmerged {len(inside)} range(s)"


def _horizontal(align: str) -> str:
    if align.strip().lower() in VERTICAL_ALIGNMENTS:
        raise SlimError(
            ErrorCode.BAD_ARGUMENT,
            f"align is horizontal only; '{align}' is a vertical alignment",
            hint=f"Use valign='{align.strip().lower()}'; align takes left|center|right",
        )
    return choice(align, ALIGNMENTS, "align")


def _vertical(valign: str) -> str:
    key = choice(valign, VERTICAL_ALIGNMENTS, "valign")
    return "center" if key == "middle" else key  # openpyxl/Excel spell middle as 'center'


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


def autofit_columns(ws: Any, rng: CellRange) -> int:
    """Size columns from what Excel displays: formula results (not formula text), wrapped
    text at its line width, and characters scaled by font size (Excel widths count
    characters of the default 11pt font, so a 22pt header really needs twice the room)."""
    widths: dict[int, float] = {}
    for row in _cells(ws, rng):
        for cell in row:
            chars = _display_chars(cell)
            if chars:
                scale = (cell.font.size or BASE_FONT_SIZE) / BASE_FONT_SIZE
                widths[cell.column] = max(widths.get(cell.column, 0.0), chars * scale)
    for column, width in widths.items():
        ws.column_dimensions[get_column_letter(column)].width = min(
            MAX_WIDTH, max(MIN_WIDTH, round(width) + 2)
        )
    return len(widths)


def _display_chars(cell: Any) -> int:
    value = cell.value
    if value is None:
        return 0
    if cell.data_type == "f" or isinstance(value, dt.date | dt.time):
        # No cached result in an edit load: estimate the displayed value from its format.
        return _format_chars(cell.number_format)
    if isinstance(value, int | float) and not isinstance(value, bool):
        return _number_chars(float(value), cell.number_format)
    text = str(value)
    longest_line = max((len(line) for line in text.splitlines()), default=0)
    if cell.alignment.wrap_text:
        longest_word = max((len(word) for word in text.split()), default=0)
        return min(longest_line, max(WRAP_WIDTH, longest_word))
    return longest_line


def _number_chars(value: float, number_format: str | None) -> int:
    """Width of a number as its format shows it ('0.2723' as '27.2%'), not as raw digits."""
    fmt = (number_format or "General").strip()
    if fmt == "General":
        return len(f"{value:.10g}")
    section = re.sub(r'"[^"]*"|\\.|\[[^\]]*\]', "", fmt.split(";")[0])
    if any(token in section.lower() for token in ("y", "d", "h", "m", "s")) and "0" not in section:
        return _format_chars(fmt)  # a date format on a serial number
    after_point = section.split(".")[1] if "." in section else ""
    decimals = len(re.match(r"[0#?]*", after_point).group(0))  # type: ignore[union-attr]
    scaling = re.search(r"[0#](,+)(?![0#,])", section)  # '#,##0,,' divides by 1000 per comma
    scaled = value / 1000 ** len(scaling.group(1)) if scaling else value
    if "%" in section:
        scaled *= 100
    shown = f"{scaled:,.{decimals}f}" if "," in section else f"{scaled:.{decimals}f}"
    literals = sum(len(q) for q in re.findall(r'"([^"]*)"', fmt.split(";")[0]))  # "m", "$"
    extras = len(re.sub(r"[0#?.,%]|_.", "", section)) + ("%" in section) + literals + 1
    return len(shown) + extras


def _format_chars(number_format: str | None) -> int:
    fmt = (number_format or "General").strip()
    visible = re.sub(r'"[^"]*"|\\.|_.|\[[^\]]*\]', "", fmt).lower()
    if fmt == "General" or not visible:
        return GENERAL_CHARS
    if "%" in visible:
        return PERCENT_CHARS
    if any(token in visible for token in ("y", "d", "h", "s")):
        return len(re.sub(r'"[^"]*"|\[[^\]]*\]|\\', "", fmt.split(";")[0]))  # date/time
    return NUMBER_CHARS


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
    totals: bool = False,
) -> OutputResult:
    """Make `range` an Excel table; re-running with an existing name resizes that table."""
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
        old = _check_table_conflicts(wb, ws, rng, name, totals)
        labels = _header_labels(ws, rng)
        if old is not None:
            _drop_table(ws, old, rng)
        ref = CellRange(rng.top, rng.left, rng.bottom + (1 if totals else 0), rng.right)
        table = Table(displayName=name, ref=a1_range(ref))
        table.tableStyleInfo = TableStyleInfo(name=style, showRowStripes=True)
        table.tableColumns = [TableColumn(id=i + 1, name=lb) for i, lb in enumerate(labels)]
        table.autoFilter = AutoFilter(ref=a1_range(rng))
        summed = _add_totals(ws, table, rng, labels) if totals else []
        ws.add_table(table)
        verb = "resized table" if old is not None else "table"
        extra = f"; totals row sums {', '.join(summed)}" if summed else ""
        extra = extra or ("; totals row added" if totals else "")
        return Change(
            f"{verb} {name} on {ws.title}!{a1_range(ref)} ({rng.n_rows - 1} data rows){extra}"
        )

    return edit_workbook(path, settings, f"set_table {name}", mutate, out_path, overwrite)


def _header_labels(ws: Any, rng: CellRange) -> list[str]:
    header = next(_cells(ws, CellRange(rng.top, rng.left, rng.top, rng.right)))
    labels = [str(c.value).strip() if c.value is not None else "" for c in header]
    if any(not label for label in labels) or len({lb.casefold() for lb in labels}) != len(labels):
        raise SlimError(
            ErrorCode.BAD_ARGUMENT,
            f"Header row {rng.top + 1} needs unique, non-empty labels: {labels}",
            hint="Write header labels first with write(...).",
        )
    for cell, label in zip(header, labels, strict=True):
        cell.value = label  # Excel stores table headers as text
    return labels


def _drop_table(ws: Any, table: Any, rng: CellRange) -> None:
    """Remove a table before it is re-added at a new size.

    Its old totals row is cleared unless the new range takes that row in as data (someone
    typed a new record over it).
    """
    if table.totalsRowCount:
        old = parse_range(table.ref)
        row = CellRange(old.bottom, old.left, old.bottom, old.right)
        if rng.intersect(row) is None:
            for cell in next(_cells(ws, row)):
                cell.value = None
    del ws.tables[table.name]


def _add_totals(ws: Any, table: Any, rng: CellRange, labels: list[str]) -> list[str]:
    """Fill the row under the data: SUBTOTAL(109) under numeric columns, 'Total' in the first."""
    row = rng.bottom + 1
    summed: list[str] = []
    for offset, (column, label) in enumerate(zip(table.tableColumns, labels, strict=True)):
        cell = ws.cell(row=row + 1, column=rng.left + offset + 1)
        if _numeric_column(ws, rng, rng.left + offset):
            column.totalsRowFunction = "sum"
            cell.value = f"=SUBTOTAL(109,{table.displayName}[{_escape_column(label)}])"
            summed.append(label)
        elif offset == 0:
            column.totalsRowLabel = "Total"
            cell.value = "Total"
    table.totalsRowCount = 1
    return summed


def _numeric_column(ws: Any, rng: CellRange, col: int) -> bool:
    values = [ws.cell(row=r + 1, column=col + 1).value for r in range(rng.top + 1, rng.bottom + 1)]
    present = [v for v in values if v is not None]
    return bool(present) and all(
        isinstance(v, int | float) and not isinstance(v, bool) for v in present
    )


def _escape_column(label: str) -> str:
    """Structured-reference escaping: ' [ ] # are prefixed with a single quote."""
    return re.sub(r"(['\[\]#])", r"'\1", label)


def _check_table_conflicts(wb: Any, ws: Any, rng: CellRange, name: str, totals: bool) -> Any | None:
    """Return the same-named table on this sheet (to resize), or raise on any conflict."""
    same: Any | None = None
    for other_ws in wb.worksheets:
        for table in other_ws.tables.values():
            if table.name.casefold() == name.casefold():
                if other_ws is not ws:
                    raise SlimError(
                        ErrorCode.BAD_ARGUMENT,
                        f"A table named '{table.name}' already exists on sheet {other_ws.title}",
                        hint="Pick another name, or run set_table on that sheet to resize it.",
                    )
                same = table
            elif other_ws is ws and parse_range(table.ref).intersect(rng) is not None:
                raise SlimError(
                    ErrorCode.BAD_ARGUMENT, f"{a1_range(rng)} overlaps table {table.name}"
                )
    if totals:
        _check_totals_row(ws, rng, same)
    return same


def _check_totals_row(ws: Any, rng: CellRange, same: Any | None) -> None:
    """The totals row goes under the data, so that row must be empty (or the old totals row)."""
    if same is not None and same.totalsRowCount and parse_range(same.ref).bottom == rng.bottom + 1:
        return
    row = CellRange(rng.bottom + 1, rng.left, rng.bottom + 1, rng.right)
    if any(c.value is not None for c in next(_cells(ws, row))):
        raise SlimError(
            ErrorCode.BAD_ARGUMENT,
            f"The totals row {a1_range(row)} is not empty",
            hint="Insert a row first: range_ops(op='insert', range='<row>'), or clear it.",
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


# --- rules (validation + conditional formatting in one tool) ------------------------------

VALIDATION_KINDS = ("list", "dropdown", *VALIDATION_TYPES[1:])
RULE_KINDS = (*VALIDATION_KINDS, *(k for k in CF_RULES if k not in VALIDATION_KINDS))


def rules(
    path: str,
    sheet: str,
    range: str,
    kind: str,
    settings: Settings,
    values: list[str] | None = None,
    value: str | None = None,
    value2: str | None = None,
    formula: str | None = None,
    message: str | None = None,
    allow_blank: bool = True,
    fill_color: str = "light_red",
    font_color: str = "9C0006",
    out_path: str | None = None,
    overwrite: bool = False,
) -> OutputResult:
    """Dropdowns and input rules (list/whole/decimal/date/text_length/custom) or highlight
    rules (greater_than ... formula): value/value2 are the bounds of either kind."""
    key = choice(kind, RULE_KINDS, "kind")
    for text in (formula, value if key == "formula" else None):
        if text:
            check_formula(text if text.startswith("=") else f"={text}")
    if key in VALIDATION_KINDS:
        return add_validation(
            path, sheet, range, "list" if key == "dropdown" else key, settings, values,
            value, value2, formula, message, allow_blank, out_path, overwrite,
        )  # fmt: skip
    return conditional_format(
        path, sheet, range, key, settings, value if value is not None else formula, value2,
        fill_color, font_color, out_path, overwrite,
    )  # fmt: skip
