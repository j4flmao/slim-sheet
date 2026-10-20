"""`write` use-case: non-destructive, journaled cell writes (PLAN §3.2, §8).

Target selection, journaling and atomic saves come from the shared edit session
(services/editing.py), so writes keep slim-sheet charts and follow working-copy rules.
"""

from __future__ import annotations

import re
from collections.abc import Sequence
from dataclasses import dataclass
from typing import Any

from slim_sheet.config import Settings
from slim_sheet.core.a1 import format_range, index_to_col, parse_range
from slim_sheet.core.errors import ErrorCode, SlimError
from slim_sheet.core.models import CellRange, CellValue, WriteResult
from slim_sheet.core.validate import validate
from slim_sheet.services.editing import Change, edit_workbook

MAX_CELLS = 10_000
MAX_FORMULA_WARNINGS = 5
_STRING_LITERAL = re.compile(r'"[^"]*"')
_DANGEROUS = re.compile(
    r"\b(WEBSERVICE|FILTERXML|CALL|REGISTER(\.ID)?|EXEC|RTD|IMPORTXML)\s*\("  # web / code
    r"|\[[^\]]*\.(xl\w*|ods|csv)\]|\[\d+\][^!]*!"  # links to other workbooks
    r"|\|",  # DDE, e.g. =cmd|'/c calc'!A0
    re.IGNORECASE,
)


@dataclass(frozen=True, slots=True)
class WriteRequest:
    path: str
    sheet: str
    range: str
    values: Sequence[Sequence[CellValue]]
    out_path: str | None = None
    overwrite: bool = False


def write(request: WriteRequest, settings: Settings) -> WriteResult:
    if not request.sheet.strip():
        raise SlimError(ErrorCode.BAD_ARGUMENT, "sheet is empty", hint="e.g. sheet='Sheet1'")
    area = _area(request.range, request.values)
    warnings = _check_values(request.values, area)
    created: list[bool] = []

    def mutate(wb: Any) -> Change:
        existing = next(
            (t for t in wb.sheetnames if t.casefold() == request.sheet.casefold()), None
        )
        created.append(existing is None)
        ws = wb[existing] if existing else wb.create_sheet(request.sheet)
        for r, row in enumerate(request.values):
            for c, value in enumerate(row):
                ws.cell(row=area.top + r + 1, column=area.left + c + 1, value=value)
        return Change(f"wrote {ws.title}!{format_range(area)}")

    note = f"write {request.sheet}!{format_range(area)}"
    out = edit_workbook(request.path, settings, note, mutate, request.out_path, request.overwrite)
    return WriteResult(
        source=out.source,
        target=out.target,
        sheet=request.sheet,
        written=area,
        cells=sum(len(row) for row in request.values),
        created_sheet=created[0],
        backup=out.backup,
        warnings=(*warnings, *out.warnings),
    )


def _area(range_ref: str, values: Sequence[Sequence[CellValue]]) -> CellRange:
    if not values or not any(values):
        raise SlimError(
            ErrorCode.BAD_ARGUMENT, "values is empty", hint='e.g. values=[["Total", 120]]'
        )
    rows, cols = len(values), max(len(row) for row in values)
    if rows * cols > MAX_CELLS:
        raise SlimError(
            ErrorCode.TOO_LARGE, f"{rows * cols} cells exceeds the {MAX_CELLS}-cell limit"
        )
    asked = parse_range(range_ref)
    anchor = CellRange(asked.top, asked.left, asked.top, asked.left)
    is_anchor = asked == anchor
    if not is_anchor and (asked.n_rows, asked.n_cols) != (rows, cols):
        raise SlimError(
            ErrorCode.BAD_ARGUMENT,
            f"range {format_range(asked)} is {asked.n_rows}x{asked.n_cols} "
            f"but values are {rows}x{cols}",
            hint=f"Pass only the top-left cell, e.g. range='{format_range(anchor)}'",
        )
    return CellRange(asked.top, asked.left, asked.top + rows - 1, asked.left + cols - 1)


def _check_values(values: Sequence[Sequence[CellValue]], area: CellRange) -> list[str]:
    formulas = 0
    problems: list[str] = []
    for r, row in enumerate(values):
        for c, value in enumerate(row):
            if not (isinstance(value, str) and value.startswith("=")):
                continue
            formulas += 1
            if _DANGEROUS.search(_STRING_LITERAL.sub('""', value)):
                raise SlimError(
                    ErrorCode.FORMULA_DENIED,
                    f"Formula reaches outside the workbook: {value[:80]}",
                    hint="External links, DDE and web functions are not allowed.",
                )
            report = validate(value)
            if not report.ok and len(problems) < MAX_FORMULA_WARNINGS:
                where = f"{index_to_col(area.left + c)}{area.top + r + 1}"
                problems.append(
                    f"{where} {value[:60]}: {report.errors[0]} (check validate_formula)"
                )
    if not formulas:
        return []
    return [
        *problems,
        f"{formulas} formula(s) written without cached results; "
        "reads show them empty until recalculated in Excel (evaluate computes them offline)",
    ]
