"""Formula audit (`anomalies scope=formulas`) and dependency tracing (`describe_cell trace=`).

Both read every formula of an xlsx/xlsm plus Excel's cached values; nothing is recalculated
and nothing is written.
"""

from __future__ import annotations

from collections import Counter
from pathlib import Path

import openpyxl

from slim_sheet.adapters.xlsx_meta import sheet_formulas
from slim_sheet.config import Settings
from slim_sheet.core.a1 import parse_range
from slim_sheet.core.errors import ErrorCode, SlimError
from slim_sheet.core.formula_audit import audit, dependents, precedents
from slim_sheet.core.models import AnomalyResult, FileFormat, Finding, TraceResult
from slim_sheet.core.refs import split_sheet
from slim_sheet.services.workbook import open_file, pick_sheet

TRACE_DIRECTIONS = ("none", "precedents", "dependents")
MAX_TRACE_LINES = 60


def _xlsx_path(path: str, settings: Settings) -> tuple[Path, list[str], str]:
    with open_file(path, settings) as handle:
        if handle.format != FileFormat.XLSX:
            raise SlimError(
                ErrorCode.UNSUPPORTED_FORMAT,
                f"Formulas are only readable from xlsx/xlsm, not {handle.format.value}",
                hint=f"convert(path='{handle.display}', out_path='copy.xlsx') keeps values only",
            )
        return handle.path, [n for n, _ in handle.reader.sheets()], handle.display


def _cached_values(path: Path) -> dict[str, dict[tuple[int, int], object]]:
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    try:
        values: dict[str, dict[tuple[int, int], object]] = {}
        for ws in wb.worksheets:
            sheet: dict[tuple[int, int], object] = {}
            for row in ws.iter_rows():
                for cell in row:
                    value = getattr(cell, "value", None)
                    if value is not None and value != "":
                        sheet[(cell.row - 1, cell.column - 1)] = value
            values[ws.title] = sheet
        return values
    finally:
        wb.close()


def formula_audit(
    path: str, settings: Settings, sheet: str | None = None, max_results: int = 20
) -> AnomalyResult:
    source, sheets, display = _xlsx_path(path, settings)
    formulas = sheet_formulas(source)
    values = _cached_values(source)
    if sheet:
        name = next((s for s in sheets if s.casefold() == sheet.casefold()), None)
        if name is None:
            raise SlimError(
                ErrorCode.SHEET_NOT_FOUND,
                f"No sheet '{sheet}'",
                hint=f"Sheets: {', '.join(sheets)}",
            )
        formulas = {name: formulas.get(name, {})}
        values = {name: values.get(name, {})}
    found = audit(formulas, values)
    findings = tuple(
        Finding(f.address, f.formula, f.kind, f.detail)
        for f in found[: max(1, min(max_results, 200))]
    )
    counts = tuple(Counter(f.kind for f in found).items())
    scanned = sum(len(cells) for cells in formulas.values())
    return AnomalyResult(display, sheet or "all sheets", "formulas", scanned, findings, counts)


def trace_cell(
    path: str, settings: Settings, cell: str, sheet: str | None, direction: str, depth: int
) -> TraceResult:
    """What a cell reads (precedents) or what reads it (dependents), breadth first."""
    if direction not in TRACE_DIRECTIONS:
        raise SlimError(
            ErrorCode.BAD_ARGUMENT,
            f"Unknown trace '{direction}'",
            hint="Use precedents or dependents",
        )
    source, _sheets, _display = _xlsx_path(path, settings)
    sheet_part, ref = split_sheet(cell)
    with open_file(path, settings) as handle:
        name = pick_sheet(handle, sheet_part or sheet)
    rng = parse_range(ref)
    formulas = sheet_formulas(source)
    levels = max(1, min(depth, 10))
    walk = precedents if direction == "precedents" else dependents
    lines = walk(formulas, name, (rng.top, rng.left), levels, MAX_TRACE_LINES)
    truncated = len(lines) >= MAX_TRACE_LINES
    return TraceResult(direction, f"{name}!{ref}", levels, tuple(lines), truncated)
