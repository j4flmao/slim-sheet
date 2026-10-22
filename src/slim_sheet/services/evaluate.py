"""`evaluate` use-case: compute a formula offline (PLAN §10), with dependency chasing.

Inputs are read from the workbook's cached values (what Excel last calculated). A
precedent cell that holds a formula WITHOUT a cached value (typical for files written by
scripts) is evaluated recursively, with cycle and depth guards. Functions the engine does
not implement fail with `unsupported_function`; nothing is ever guessed.
"""

from __future__ import annotations

import datetime as dt
from contextlib import ExitStack
from typing import Any

from slim_sheet.adapters import formula_engine as engine
from slim_sheet.adapters.xlsx_meta import XlsxMeta
from slim_sheet.config import Settings
from slim_sheet.core.a1 import format_range, index_to_col, parse_range
from slim_sheet.core.budget import cell_text
from slim_sheet.core.diff import same
from slim_sheet.core.errors import ErrorCode, SlimError
from slim_sheet.core.formula import analyze
from slim_sheet.core.functions import VOLATILE
from slim_sheet.core.models import CellRange, CellValue, EvalResult, FileFormat
from slim_sheet.core.refs import classify, split_sheet
from slim_sheet.services.workbook import OpenFile, open_file, pick_sheet

MAX_DEPTH = 30
MAX_INPUT_CELLS = 250_000
DATE_RESULTS = {"DATE", "TODAY", "NOW", "EDATE", "EOMONTH", "WORKDAY", "WORKDAY.INTL", "DATEVALUE"}


class _Evaluator:
    def __init__(
        self,
        handle: OpenFile | None,
        meta: XlsxMeta | None,
        overrides: dict[str, Any] | None = None,
        fresh: bool = False,
    ) -> None:
        self.handle = handle
        self.meta = meta
        # what_if: 'Sheet!B3' -> value replacing the cell; fresh recomputes every formula
        self.overrides = {k.casefold(): v for k, v in (overrides or {}).items()}
        self.fresh = fresh
        self.inputs: list[tuple[str, str]] = []
        self.recomputed: list[str] = []
        self._stack: list[str] = []
        self._done: dict[str, Any] = {}

    def evaluate(self, text: str, sheet: str | None, depth: int = 0) -> Any:
        _precheck(text)
        compiled = engine.compile_formula(text)
        args = [self._input(ref, sheet, depth) for ref in compiled.inputs]
        return engine.run(compiled, args)

    def _input(self, ref: str, context: str | None, depth: int) -> Any:
        if self.handle is None:
            raise SlimError(
                ErrorCode.BAD_ARGUMENT,
                f"The formula reads {ref}; pass path (and sheet) so cells can be read",
                hint="e.g. evaluate(path='book.xlsx', formula='=SUM(A1:A9)', sheet='Data')",
            )
        sheet, rng = self._locate(ref, context)
        cells = rng.n_rows * rng.n_cols
        if cells > MAX_INPUT_CELLS:
            raise SlimError(ErrorCode.TOO_LARGE, f"{ref} covers {cells} cells")
        grid = self.handle.reader.read(sheet, rng)
        values = [
            [self._cell(sheet, rng.top + r, rng.left + c, v, depth) for c, v in enumerate(row)]
            for r, row in enumerate(grid)
        ]
        if depth == 0:
            label = f"{sheet}!{format_range(rng)}"
            shown = f"({cells} cells)"
            if cells == 1:
                if label in self._done:
                    value = engine.from_engine(self._done[label])
                    text = "array" if isinstance(value, tuple) else cell_text(value, 40)
                    shown = f"= {text} (recomputed)"
                else:
                    shown = f"= {cell_text(grid[0][0], 40)}"
            self.inputs.append((label, shown))
        if cells == 1:
            return values[0][0]
        import numpy as np

        return np.array(values, dtype=object)

    def _locate(self, ref: str, context: str | None) -> tuple[str, CellRange]:
        assert self.handle is not None
        sheet_part, address = split_sheet(ref)
        kind = classify(address).kind
        if kind == "name":
            if self.meta is None:
                raise SlimError(
                    ErrorCode.BAD_ARGUMENT, f"Defined name {address} needs an xlsx file"
                )
            sheet_part, address = self.meta.name_target(address)
        elif kind in ("table", "external"):
            raise SlimError(
                ErrorCode.UNSUPPORTED_FUNCTION,
                f"{ref}: {kind} references cannot be evaluated offline",
                hint="Rewrite with a plain range such as Sheet1!B2:B100.",
            )
        sheet = pick_sheet(self.handle, sheet_part or context)
        used = self.handle.reader.used_range(sheet) or CellRange(0, 0, 0, 0)
        if kind == "columns":
            left, _, right = address.replace("$", "").partition(":")
            cols = parse_range(f"{left}1:{right}1")
            return sheet, CellRange(0, cols.left, used.bottom, cols.right)
        if kind == "rows":
            top, _, bottom = address.replace("$", "").partition(":")
            return sheet, CellRange(int(top) - 1, 0, int(bottom) - 1, used.right)
        return sheet, parse_range(address)

    def _cell(self, sheet: str, row: int, col: int, cached: CellValue, depth: int) -> Any:
        key = f"{sheet}!{index_to_col(col)}{row + 1}".casefold()
        if key in self.overrides:
            return engine.to_engine(self.overrides[key])
        if cached is not None and cached != "" and not self.fresh:
            return engine.to_engine(cached)
        formula = self.meta.formula_at(sheet, row, col) if self.meta is not None else None
        if formula is None:
            return engine.empty() if cached is None or cached == "" else engine.to_engine(cached)
        return self._recompute(f"{sheet}!{index_to_col(col)}{row + 1}", sheet, formula, depth + 1)

    def _recompute(self, address: str, sheet: str, formula: str, depth: int) -> Any:
        if address in self._done:
            return self._done[address]
        if address in self._stack:
            chain = " -> ".join([*self._stack, address])
            raise SlimError(ErrorCode.CIRCULAR_REFERENCE, f"Circular reference: {chain}")
        if depth > MAX_DEPTH:
            raise SlimError(
                ErrorCode.TOO_LARGE, f"Dependency chain deeper than {MAX_DEPTH} at {address}"
            )
        self._stack.append(address)
        try:
            result = engine.from_engine(self.evaluate(formula, sheet, depth))
        finally:
            self._stack.pop()
        value = engine.to_engine(result) if not isinstance(result, tuple) else result[0][0]
        self._done[address] = value
        self.recomputed.append(address)
        return value


def _precheck(text: str) -> None:
    report = analyze(text)
    if report.errors:
        raise SlimError(
            ErrorCode.FORMULA_SYNTAX,
            f"{report.formula}: {report.errors[0]}",
            hint="Run explain_formula or validate_formula for details.",
        )
    for raw in report.references:
        kind = classify(raw).kind
        if kind in ("table", "external"):
            raise SlimError(
                ErrorCode.UNSUPPORTED_FUNCTION,
                f"{raw}: {kind} references cannot be evaluated offline",
                hint="Rewrite with a plain range such as Sheet1!B2:B100.",
            )
    missing = sorted({c.name for c in report.calls} - engine.supported_functions())
    if missing:
        raise SlimError(
            ErrorCode.UNSUPPORTED_FUNCTION,
            f"The offline engine does not implement: {', '.join(missing)}",
            hint="Open the file in Excel to calculate it; slim-sheet never guesses values.",
        )


def evaluate(
    settings: Settings,
    formula: str | None = None,
    path: str | None = None,
    cell: str | None = None,
    sheet: str | None = None,
) -> EvalResult:
    if not formula and not (path and cell):
        raise SlimError(
            ErrorCode.BAD_ARGUMENT,
            "Pass formula, or path and cell",
            hint="e.g. evaluate(formula='=ROUND(2.345,2)') or (path='x.xlsx', cell='Calc!C7')",
        )
    with ExitStack() as stack:
        handle = stack.enter_context(open_file(path, settings)) if path else None
        meta = None
        if handle is not None and handle.format == FileFormat.XLSX:
            meta = stack.enter_context(XlsxMeta(handle.path))
        evaluator = _Evaluator(handle, meta)
        if cell and handle is not None:
            return _evaluate_cell(evaluator, handle, meta, cell, sheet)
        context = pick_sheet(handle, sheet) if handle is not None else None
        text = formula if formula and formula.startswith("=") else f"={formula}"
        value = engine.from_engine(evaluator.evaluate(text, context))
        return _result(evaluator, text, None, value, None)


def _evaluate_cell(
    evaluator: _Evaluator, handle: OpenFile, meta: XlsxMeta | None, cell: str, sheet: str | None
) -> EvalResult:
    sheet_part, ref = split_sheet(cell)
    name = pick_sheet(handle, sheet_part or sheet)
    rng = parse_range(ref)
    address = f"{name}!{index_to_col(rng.left)}{rng.top + 1}"
    cached = handle.reader.read(name, CellRange(rng.top, rng.left, rng.top, rng.left))[0][0]
    text = meta.formula_at(name, rng.top, rng.left) if meta is not None else None
    if text is None:
        note = (
            "cell holds a value, not a formula"
            if meta is not None
            else ("formulas are only readable from xlsx/xlsm; showing the stored value")
        )
        return EvalResult("", address, cached, cached, (), (), (note,))
    value = engine.from_engine(evaluator.evaluate(text, name))
    return _result(evaluator, text, address, value, None if cached == "" else cached)


def _result(
    evaluator: _Evaluator,
    text: str,
    address: str | None,
    value: CellValue | tuple[tuple[CellValue, ...], ...],
    cached: CellValue,
) -> EvalResult:
    report = analyze(text)
    names = {c.name for c in report.calls}
    notes = []
    if names & VOLATILE:
        notes.append(f"volatile ({', '.join(sorted(names & VOLATILE))}): result changes over time")
    outer = next((c.name for c in reversed(report.calls) if c.depth == 0), None)
    if outer in DATE_RESULTS and isinstance(value, int | float):
        moment = engine.EXCEL_EPOCH + dt.timedelta(days=float(value))
        shown = (
            moment.date().isoformat()
            if float(value).is_integer()
            else moment.isoformat(" ", "seconds")
        )
        notes.append(f"as a date: {shown}")
    if (
        cached is not None
        and not isinstance(value, tuple)
        and not same(engine.to_engine(cached), value)
    ):
        notes.append(
            f"differs from Excel's cached value {cached!r} "
            "(inputs changed, or the file was not recalculated)"
        )
    return EvalResult(
        text,
        address,
        value,
        cached,
        tuple(evaluator.inputs),
        tuple(evaluator.recomputed),
        tuple(notes),
    )


def evaluate_cells(
    settings: Settings,
    path: str,
    cells: list[str],
    sheet: str | None,
    overrides: dict[str, Any],
) -> list[CellValue]:
    """Recompute `cells` from scratch with some input cells replaced (what_if)."""
    with ExitStack() as stack:
        handle = stack.enter_context(open_file(path, settings))
        if handle.format != FileFormat.XLSX:
            raise SlimError(ErrorCode.UNSUPPORTED_FORMAT, "what_if needs an xlsx/xlsm workbook")
        meta = stack.enter_context(XlsxMeta(handle.path))
        resolved = {_qualify(handle, ref, sheet): v for ref, v in overrides.items()}
        evaluator = _Evaluator(handle, meta, resolved, fresh=True)
        results: list[CellValue] = []
        for cell in cells:
            address = _qualify(handle, cell, sheet)
            if address.casefold() in evaluator.overrides:
                results.append(evaluator.overrides[address.casefold()])
                continue
            name, ref = split_sheet(address)
            rng = parse_range(ref)
            text = meta.formula_at(name or "", rng.top, rng.left)
            if text is None:
                cached = handle.reader.read(
                    name or "", CellRange(rng.top, rng.left, rng.top, rng.left)
                )
                results.append(cached[0][0])
                continue
            value = engine.from_engine(evaluator.evaluate(text, name))
            results.append(value[0][0] if isinstance(value, tuple) else value)
        return results


def _qualify(handle: OpenFile, ref: str, sheet: str | None) -> str:
    """'B3' / 'Inputs!B3' -> 'Inputs!B3' with the sheet's real name."""
    sheet_part, address = split_sheet(ref.strip())
    name = pick_sheet(handle, sheet_part or sheet)
    rng = parse_range(address)
    if rng.n_rows * rng.n_cols != 1:
        raise SlimError(ErrorCode.BAD_ARGUMENT, f"{ref} must be one cell", hint="e.g. 'Inputs!B3'")
    return f"{name}!{index_to_col(rng.left)}{rng.top + 1}"
