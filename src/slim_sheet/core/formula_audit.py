"""Model audit and dependency tracing over a workbook's formulas (PLAN §11.9). Pure.

Input is plain data: {sheet: {(row, col): '=formula'}} and {sheet: {(row, col): cached value}}
(0-based). Findings target the errors SpreadsheetBench 2 injects into models: hardcoded numbers
inside formulas, a formula that breaks its row/column pattern, a typed-over value in a block
of formulas, sums that include a subtotal (double counting), broken references, and error
values traced to their root cell.
"""

from __future__ import annotations

import re
from collections import Counter
from collections.abc import Iterator, Mapping
from dataclasses import dataclass

from slim_sheet.core.a1 import col_to_index, format_range, index_to_col, parse_range
from slim_sheet.core.errors import SlimError
from slim_sheet.core.formula import analyze
from slim_sheet.core.models import CellRange
from slim_sheet.core.refs import classify, split_sheet

Formulas = Mapping[str, Mapping[tuple[int, int], str]]
Values = Mapping[str, Mapping[tuple[int, int], object]]
ERRORS = ("#REF!", "#DIV/0!", "#N/A", "#VALUE!", "#NAME?", "#NUM!", "#NULL!", "#SPILL!", "#CALC!")
ALLOWED_CONSTANTS = {0, 1, 2, 3, 4, 7, 10, 12, 24, 52, 60, 100, 365, 1000}
_STRING = re.compile(r'"[^"]*"')
_CELL = re.compile(r"(?<![A-Za-z0-9_.])(\$?)([A-Z]{1,3})(\$?)(\d{1,7})(?![A-Za-z0-9_(])")
_NUMBER = re.compile(r"(?<![A-Za-z0-9_$.!:])\d+(?:\.\d+)?(?:[eE][+-]?\d+)?(?![\d.A-Za-z_(:!])")
_SUMS = re.compile(r"\b(SUM|SUBTOTAL|AGGREGATE)\s*\(", re.IGNORECASE)
MIN_RUN = 3


@dataclass(frozen=True, slots=True)
class AuditFinding:
    sheet: str
    row: int
    col: int
    kind: str
    formula: str
    detail: str

    @property
    def address(self) -> str:
        return f"{self.sheet}!{index_to_col(self.col)}{self.row + 1}"


# --- references ---------------------------------------------------------------------------


def references(formula: str, home: str) -> list[tuple[str, CellRange]]:
    """Cell and range references of a formula as (sheet, range); names/columns are skipped."""
    try:
        refs = analyze(formula).references
    except SlimError:
        return []
    found: list[tuple[str, CellRange]] = []
    for ref in refs:
        sheet, address = split_sheet(ref)
        if classify(address).kind != "cells":
            continue
        try:
            found.append((sheet or home, parse_range(address.replace("$", ""))))
        except SlimError:
            continue
    return found


def relative_form(formula: str, row: int, col: int) -> str:
    """'=B2*C2' in D2 -> '=RC[-2]*RC[-1]': equal forms mean the same formula copied."""

    def swap(match: re.Match[str]) -> str:
        col_abs, letters, row_abs, digits = match.groups()
        c = col_to_index(letters)
        r = int(digits) - 1
        rpart = f"R{r + 1}" if row_abs else f"R[{r - row}]"
        cpart = f"C{c + 1}" if col_abs else f"C[{c - col}]"
        return rpart + cpart

    parts = re.split(r'("[^"]*")', formula)
    return "".join(p if p.startswith('"') else _CELL.sub(swap, p) for p in parts)


# --- audit --------------------------------------------------------------------------------


def audit(formulas: Formulas, values: Values) -> list[AuditFinding]:
    """Every finding, most severe kinds first."""
    found: list[AuditFinding] = []
    for sheet, cells in formulas.items():
        found += _broken_and_hardcoded(sheet, cells)
        found += _pattern_breaks(sheet, cells, values.get(sheet, {}))
        found += _double_counting(sheet, cells)
    found += _error_roots(formulas, values)
    rank = {k: i for i, k in enumerate(KINDS)}
    return sorted(found, key=lambda f: (rank.get(f.kind, 99), f.sheet, f.row, f.col))


_TOTAL_LABEL = re.compile(
    r"\b(?:totals?|sub-?totals?|grand|overall|all|sum|combined)\b", re.IGNORECASE
)
KINDS = (
    "error_root",
    "broken_reference",
    "inconsistent_formula",
    "hardcoded_value",
    "double_count",
    "hardcode_in_formula",
)


def _broken_and_hardcoded(
    sheet: str, cells: Mapping[tuple[int, int], str]
) -> Iterator[AuditFinding]:
    for (row, col), formula in cells.items():
        if "#REF!" in formula.upper():
            yield AuditFinding(
                sheet, row, col, "broken_reference", formula, "formula contains #REF!"
            )
            continue
        bare = _CELL.sub("", _STRING.sub("", formula))
        constants = [n for n in _NUMBER.findall(bare) if float(n) not in ALLOWED_CONSTANTS]
        if constants:
            shown = ", ".join(constants[:3])
            yield AuditFinding(
                sheet, row, col, "hardcode_in_formula", formula,
                f"number(s) {shown} typed into the formula; move them to a labelled input cell",
            )  # fmt: skip


def _runs(cells: Mapping[tuple[int, int], str], vertical: bool) -> Iterator[list[tuple[int, int]]]:
    """Contiguous runs of formula cells down a column (or along a row)."""
    lines: dict[int, list[int]] = {}
    for row, col in cells:
        lines.setdefault(col if vertical else row, []).append(row if vertical else col)
    for line, positions in lines.items():
        run: list[int] = []
        for pos in sorted(positions):
            if run and pos != run[-1] + 1:
                yield [(p, line) if vertical else (line, p) for p in run]
                run = []
            run.append(pos)
        if run:
            yield [(p, line) if vertical else (line, p) for p in run]


def _summary_rows(
    cells: Mapping[tuple[int, int], str], values: Mapping[tuple[int, int], object]
) -> set[int]:
    """Rows labelled like a total ('Total', 'All 130 listings', 'Vietnam total') and the formula
    rows right under them: a summary block whose formulas differ from the body by design."""
    first_col: dict[int, int] = {}
    for row, col in cells:
        first_col[row] = min(col, first_col.get(row, col))
    totals = set()
    for row, col in first_col.items():
        labels = [values.get((row, c)) for c in range(col)]
        text = [v for v in labels if isinstance(v, str) and v.strip()]
        if text and _TOTAL_LABEL.search(text[-1]):
            totals.add(row)
    block = set(totals)
    for row in sorted(totals):
        below = row + 1
        while below in first_col and below not in block:
            block.add(below)
            below += 1
    return block


def _pattern_breaks(
    sheet: str, cells: Mapping[tuple[int, int], str], values: Mapping[tuple[int, int], object]
) -> list[AuditFinding]:
    found: dict[tuple[int, int], AuditFinding] = {}
    summary = _summary_rows(cells, values)
    body = {cell: f for cell, f in cells.items() if cell[0] not in summary}
    for vertical in (True, False):
        for run in _runs(body, vertical):
            if len(run) < MIN_RUN:
                continue
            forms = {cell: relative_form(body[cell], *cell) for cell in run}
            common, count = Counter(forms.values()).most_common(1)[0]
            if count < max(2, (len(run) + 1) // 2) or count == len(run):
                continue
            example = next(body[c] for c in run if forms[c] == common)
            for cell in run:
                if forms[cell] != common and cell not in found:
                    found[cell] = AuditFinding(
                        sheet, *cell, "inconsistent_formula", cells[cell],
                        f"differs from {count} neighbours (they look like {example})",
                    )  # fmt: skip
            found.update(_typed_over(sheet, run, values, vertical, found))
    return list(found.values())


def _typed_over(
    sheet: str,
    run: list[tuple[int, int]],
    values: Mapping[tuple[int, int], object],
    vertical: bool,
    seen: Mapping[tuple[int, int], AuditFinding],
) -> dict[tuple[int, int], AuditFinding]:
    """A plain number right after a run of formulas, with formulas continuing after it."""
    result: dict[tuple[int, int], AuditFinding] = {}
    last = run[-1]
    gap = (last[0] + 1, last[1]) if vertical else (last[0], last[1] + 1)
    value = values.get(gap)
    if isinstance(value, int | float) and not isinstance(value, bool) and gap not in seen:
        result[gap] = AuditFinding(
            sheet, *gap, "hardcoded_value", str(value),
            "a typed number where the formulas around it would calculate one",
        )  # fmt: skip
    return result


def _double_counting(sheet: str, cells: Mapping[tuple[int, int], str]) -> Iterator[AuditFinding]:
    subtotals = {cell for cell, formula in cells.items() if _SUMS.search(formula)}
    for (row, col), formula in cells.items():
        if not _SUMS.search(formula):
            continue
        for ref_sheet, rng in references(formula, sheet):
            if ref_sheet != sheet or rng.n_rows * rng.n_cols < 2:
                continue
            inner = [c for c in subtotals if c != (row, col) and _inside(c, rng)]
            if inner:
                where = ", ".join(f"{index_to_col(c[1])}{c[0] + 1}" for c in sorted(inner)[:3])
                yield AuditFinding(
                    sheet, row, col, "double_count", formula,
                    f"the summed range also contains subtotal(s) at {where}",
                )  # fmt: skip
                break


def _inside(cell: tuple[int, int], rng: CellRange) -> bool:
    return rng.top <= cell[0] <= rng.bottom and rng.left <= cell[1] <= rng.right


def _is_error(value: object) -> bool:
    return isinstance(value, str) and value.strip().upper() in ERRORS


def _error_roots(formulas: Formulas, values: Values) -> list[AuditFinding]:
    """Error cells whose precedents hold no error: where a #DIV/0! or #REF! starts."""
    found: list[AuditFinding] = []
    propagated = 0
    for sheet, cells in values.items():
        for cell, value in cells.items():
            if not _is_error(value):
                continue
            formula = formulas.get(sheet, {}).get(cell, "")
            inputs = references(formula, sheet) if formula else []
            if any(_range_has_error(values, s, r) for s, r in inputs):
                propagated += 1
                continue
            found.append(
                AuditFinding(
                    sheet, *cell, "error_root", formula or str(value), f"{value} starts here"
                )
            )
    if found and propagated:
        first = found[0]
        found[0] = AuditFinding(
            first.sheet, first.row, first.col, first.kind, first.formula,
            f"{first.detail}; {propagated} other error cell(s) only inherit errors",
        )  # fmt: skip
    return found


def _range_has_error(values: Values, sheet: str, rng: CellRange) -> bool:
    cells = values.get(sheet, {})
    if rng.n_rows * rng.n_cols > 50_000:
        return False
    return any(_is_error(v) for c, v in cells.items() if _inside(c, rng))


# --- tracing ------------------------------------------------------------------------------


def precedents(
    formulas: Formulas, sheet: str, cell: tuple[int, int], depth: int, limit: int
) -> list[tuple[int, str, str]]:
    """(level, address, formula or '') for what a cell reads, breadth first."""
    lines: list[tuple[int, str, str]] = []
    queue = [(sheet, cell, 0)]
    seen = {(sheet, cell)}
    while queue and len(lines) < limit:
        here_sheet, here, level = queue.pop(0)
        formula = formulas.get(here_sheet, {}).get(here)
        if not formula or level >= depth:
            continue
        for ref_sheet, rng in references(formula, here_sheet):
            label = _label(ref_sheet, rng)
            inner = [c for c in formulas.get(ref_sheet, {}) if _inside(c, rng)]
            single = rng.n_rows * rng.n_cols == 1
            shown = (
                formulas.get(ref_sheet, {}).get((rng.top, rng.left), "")
                if single
                else f"{len(inner)} formula cell(s)"
            )
            lines.append((level + 1, label, shown))
            for c in inner[:20]:
                if (ref_sheet, c) not in seen:
                    seen.add((ref_sheet, c))
                    queue.append((ref_sheet, c, level + 1))
    return lines[:limit]


def dependents(
    formulas: Formulas, sheet: str, cell: tuple[int, int], depth: int, limit: int
) -> list[tuple[int, str, str]]:
    """(level, address, formula) for every formula that reads the cell, breadth first."""
    readers: list[tuple[str, tuple[int, int], str, list[tuple[str, CellRange]]]] = [
        (s, c, f, references(f, s)) for s, cells in formulas.items() for c, f in cells.items()
    ]
    lines: list[tuple[int, str, str]] = []
    queue = [(sheet, cell, 0)]
    seen = {(sheet, cell)}
    while queue and len(lines) < limit:
        here_sheet, here, level = queue.pop(0)
        if level >= depth:
            continue
        for s, c, f, refs in readers:
            if (s, c) in seen or not any(rs == here_sheet and _inside(here, r) for rs, r in refs):
                continue
            seen.add((s, c))
            lines.append((level + 1, f"{s}!{index_to_col(c[1])}{c[0] + 1}", f))
            queue.append((s, c, level + 1))
    return lines[:limit]


def _label(sheet: str, rng: CellRange) -> str:
    return f"{sheet}!{format_range(rng)}"
