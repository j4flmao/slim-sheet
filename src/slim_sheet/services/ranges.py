"""`range_ops` use-case: the everyday editing verbs on a block of cells (PLAN §11.9, §17.2).

op = fill | copy | move | transpose | sort | replace | insert | delete
- fill: copy the first row down (or first column right), translating formulas; two seed
  numbers or dates with blanks below continue as a series.
- copy / move / transpose: to another place (also another sheet); formulas are translated.
- sort: rows of a table by one column (header kept on top).
- replace: text in values (formulas are left alone).
- insert / delete: rows ('5', '5:7') or columns ('C', 'C:E') with Excel's reference rules.
- where (copy/delete): a SQL condition on the table's columns, e.g. "Status = 'Closed'",
  evaluated by the locked DuckDB engine on the cells' cached values.
"""

from __future__ import annotations

import copy as copy_module
import datetime as dt
import logging
import re
from dataclasses import dataclass
from pathlib import Path
from typing import Any

import duckdb
import openpyxl
from openpyxl.formula.translate import Translator

from slim_sheet.adapters.engine.arrow import to_arrow
from slim_sheet.adapters.engine.cache import LoadedTable
from slim_sheet.adapters.engine.duck import engine_from
from slim_sheet.adapters.writers.workbook_refs import shift_workbook
from slim_sheet.config import Settings
from slim_sheet.core.a1 import col_to_index, index_to_col, parse_range
from slim_sheet.core.a1 import format_range as a1
from slim_sheet.core.errors import ErrorCode, SlimError
from slim_sheet.core.models import CellRange, OutputResult
from slim_sheet.core.move_refs import Move, move_references, qualify
from slim_sheet.core.refs import split_sheet
from slim_sheet.core.shift import Axis, Shift
from slim_sheet.core.styles import choice
from slim_sheet.services.design import TableBlock, find_table
from slim_sheet.services.editing import (
    Change,
    edit_workbook,
    get_sheet,
    resolve_target,
    writable_source,
)

LOG = logging.getLogger(__name__)
OPS = ("fill", "copy", "move", "transpose", "sort", "replace", "insert", "delete", "clear", "split")
CLEAR_WHAT = ("contents", "formats", "rules", "all")
MAX_CELLS = 200_000
MAX_SPAN = 10_000
MAX_MATCHES = 5_000
_SPAN = re.compile(r"^\s*(\$?[A-Za-z]{1,3}|\d+)(?:\s*:\s*(\$?[A-Za-z]{1,3}|\d+))?\s*$")


@dataclass(frozen=True, slots=True)
class RangeRequest:
    path: str
    sheet: str
    op: str
    range: str | None = None
    to: str | None = None
    by: str | None = None
    desc: bool = False
    find: str | None = None
    replace: str | None = None
    where: str | None = None
    header: bool = True
    out_path: str | None = None
    overwrite: bool = False
    values_only: bool = False
    formats_only: bool = False  # copy: styles only (Excel's format painter)
    what: str = "contents"  # clear: contents | formats | rules | all
    delimiter: str = ","  # split: text to columns


@dataclass(frozen=True, slots=True)
class Prepared:
    """What is read from the file before editing: rows matching `where`, cached values."""

    matches: list[int] | None = None
    cached: dict[tuple[int, int], Any] | None = None  # 0-based (row, col) -> Excel's last value


def range_ops(request: RangeRequest, settings: Settings) -> OutputResult:
    op = choice(request.op, OPS, "op")
    _check(op, request)
    pre = Prepared(
        _matches(request, settings) if request.where else None,
        _cached_block(request, settings) if request.values_only else None,
    )

    def mutate(wb: Any) -> Change:
        ws = get_sheet(wb, request.sheet)
        handler = _HANDLERS[op]
        return handler(wb, ws, request, pre)

    note = f"range_ops {op} {request.sheet}!{request.range or ''}"
    return edit_workbook(request.path, settings, note, mutate, request.out_path, request.overwrite)


def _check(op: str, req: RangeRequest) -> None:
    needs_range = op in ("fill", "move", "transpose", "insert", "clear", "split") or (
        op in ("copy", "delete") and not req.where
    )
    if needs_range and not req.range:
        raise SlimError(ErrorCode.BAD_ARGUMENT, f"op={op} needs range", hint=_EXAMPLES[op])
    if op in ("copy", "move", "transpose") and not req.to:
        raise SlimError(ErrorCode.BAD_ARGUMENT, f"op={op} needs to", hint=_EXAMPLES[op])
    if op == "sort" and not req.by:
        raise SlimError(ErrorCode.BAD_ARGUMENT, "op=sort needs by", hint=_EXAMPLES[op])
    if op == "replace" and not req.find:
        raise SlimError(ErrorCode.BAD_ARGUMENT, "op=replace needs find", hint=_EXAMPLES[op])
    if req.where and op not in ("copy", "delete"):
        raise SlimError(
            ErrorCode.BAD_ARGUMENT,
            f"where works with op=copy or op=delete, not {op}",
            hint="e.g. op='delete', where=\"Status = 'Closed'\"",
        )
    if req.formats_only and (op != "copy" or req.values_only or req.where):
        raise SlimError(
            ErrorCode.BAD_ARGUMENT,
            "formats_only works with op=copy alone (no values_only, no where)",
            hint="e.g. op='copy', range='B2', to='B3:B20', formats_only=true",
        )
    if op == "clear":
        choice(req.what, CLEAR_WHAT, "what")
    if op == "split" and not req.delimiter:
        raise SlimError(
            ErrorCode.BAD_ARGUMENT, "split needs a delimiter", hint="e.g. delimiter=';'"
        )
    if req.values_only and op not in ("copy", "transpose"):
        raise SlimError(
            ErrorCode.BAD_ARGUMENT,
            f"values_only works with op=copy or op=transpose, not {op}",
            hint="e.g. op='copy', range='D2:D9', to='Report!B2', values_only=true",
        )
    if req.where and ";" in req.where:
        raise SlimError(ErrorCode.BAD_SQL, "where must be one condition, without ';'")


_EXAMPLES = {
    "fill": "e.g. range_ops(path, sheet='Data', op='fill', range='D2:D200')",
    "copy": "e.g. op='copy', range='A1:C10', to='Summary!B2'",
    "move": "e.g. op='move', range='A1:C10', to='E1'",
    "transpose": "e.g. op='transpose', range='A1:C3', to='F1'",
    "sort": "e.g. op='sort', range='A1:F200', by='Revenue', desc=true",
    "replace": "e.g. op='replace', find='N/A', replace=''",
    "insert": "e.g. op='insert', range='5:7' (rows) or range='C' (a column)",
    "delete": "e.g. op='delete', range='5' or op='delete', where=\"Status = 'Closed'\"",
    "clear": "e.g. op='clear', range='B2:D9', what='formats'",
    "split": "e.g. op='split', range='A2:A200', delimiter=';'",
}


# --- insert / delete ----------------------------------------------------------------------


def parse_span(text: str) -> tuple[Axis, int, int]:
    """'5' / '5:7' -> rows; 'C' / 'C:E' -> columns. -> (axis, first index 1-based, count)."""
    match = _SPAN.match(text or "")
    if not match:
        raise SlimError(
            ErrorCode.BAD_RANGE,
            f"'{text}' is not a row or column span",
            hint="Use '5', '5:7' for rows or 'C', 'C:E' for columns",
        )
    first, last = (
        match.group(1).replace("$", ""),
        (match.group(2) or match.group(1)).replace("$", ""),
    )
    if first.isdigit() != last.isdigit():
        raise SlimError(ErrorCode.BAD_RANGE, f"'{text}' mixes rows and columns")
    if first.isdigit():
        lo, hi = sorted((int(first), int(last)))
        axis: Axis = "rows"
    else:
        lo, hi = sorted((col_to_index(first.upper()) + 1, col_to_index(last.upper()) + 1))
        axis = "cols"
    if lo < 1 or hi - lo + 1 > MAX_SPAN:
        raise SlimError(ErrorCode.BAD_RANGE, f"'{text}' must start at 1 and span <= {MAX_SPAN}")
    return axis, lo, hi - lo + 1


def _shift(wb: Any, ws: Any, axis: Axis, index: int, count: int, insert: bool) -> list[str]:
    shift = Shift(axis, index, count if insert else -count)
    method = getattr(ws, f"{'insert' if insert else 'delete'}_{axis}")
    report = shift_workbook(wb, ws, shift, lambda: method(index, count))
    return list(report.broken)


def _insert(wb: Any, ws: Any, req: RangeRequest, _pre: Prepared) -> Change:
    axis, index, count = parse_span(req.range or "")
    _shift(wb, ws, axis, index, count, insert=True)
    return Change(f"inserted {count} {axis[:-1]}(s) at {req.range} in '{ws.title}'")


def _delete(wb: Any, ws: Any, req: RangeRequest, pre: Prepared) -> Change:
    matches = pre.matches
    if matches is not None:
        broken: list[str] = []
        for top, size in _blocks(matches):  # bottom-up so earlier rows keep their numbers
            broken += _shift(wb, ws, "rows", top, size, insert=False)
        warnings = tuple(f"{cell} now contains #REF!" for cell in broken[:10])
        return Change(
            f"deleted {len(matches)} row(s) matching {req.where} in '{ws.title}'", warnings
        )
    axis, index, count = parse_span(req.range or "")
    broken = _shift(wb, ws, axis, index, count, insert=False)
    warnings = tuple(f"{cell} now contains #REF!" for cell in broken[:10])
    return Change(f"deleted {count} {axis[:-1]}(s) at {req.range} in '{ws.title}'", warnings)


def _blocks(rows: list[int]) -> list[tuple[int, int]]:
    """1-based rows -> contiguous (start, size) blocks, last block first."""
    blocks: list[tuple[int, int]] = []
    for row in sorted(set(rows)):
        if blocks and blocks[-1][0] + blocks[-1][1] == row:
            blocks[-1] = (blocks[-1][0], blocks[-1][1] + 1)
        else:
            blocks.append((row, 1))
    return blocks[::-1]


# --- where: rows matching a SQL condition -------------------------------------------------


def _matches(req: RangeRequest, settings: Settings) -> list[int]:
    """1-based sheet rows of the table whose cached values satisfy `where`."""
    source = writable_source(req.path, settings)
    target = resolve_target(source, settings, req.out_path, req.overwrite)
    base = target if target.exists() else source
    headers, rows, first = _cached_table(base, req)
    columns = (*headers, "__row")
    body = [(*row, first + i) for i, row in enumerate(rows)]
    engine = engine_from([LoadedTable("t", req.sheet, to_arrow(columns, body))])
    try:
        found = engine.con.execute(f'SELECT "__row" FROM t WHERE ({req.where})').fetchall()
    except duckdb.Error as exc:
        cols = ", ".join(f'"{h}"' for h in headers)
        raise SlimError(ErrorCode.BAD_SQL, f"where failed: {exc}", hint=f"Columns: {cols}") from exc
    finally:
        engine.close()
    if len(found) > MAX_MATCHES:
        raise SlimError(ErrorCode.TOO_LARGE, f"{len(found)} rows match; the limit is {MAX_MATCHES}")
    return [int(r[0]) for r in found]


def _cached_block(req: RangeRequest, settings: Settings) -> dict[tuple[int, int], Any]:
    """Excel's last calculated values for the source range (for values_only pastes)."""
    source = writable_source(req.path, settings)
    target = resolve_target(source, settings, req.out_path, req.overwrite)
    base = target if target.exists() else source
    rng = _area(req.range)
    wb = openpyxl.load_workbook(base, data_only=True)
    try:
        ws = get_sheet(wb, req.sheet)
        values = {
            (r, c): ws.cell(row=r + 1, column=c + 1).value
            for r in range(rng.top, rng.bottom + 1)
            for c in range(rng.left, rng.right + 1)
        }
        title = ws.title
    finally:
        wb.close()
    return _fill_uncached(values, base, title, settings)


def _fill_uncached(
    values: dict[tuple[int, int], Any], base: Path, sheet: str, settings: Settings
) -> dict[tuple[int, int], Any]:
    """Formulas without a cached result (files saved by scripts, or by an earlier edit) are
    calculated with the offline engine; if it cannot, they stay empty and are reported."""
    from slim_sheet.services.evaluate import evaluate_cells

    wb = openpyxl.load_workbook(base)
    try:
        ws = get_sheet(wb, sheet)
        todo = [
            key
            for key, value in values.items()
            if value is None
            and str(ws.cell(row=key[0] + 1, column=key[1] + 1).value).startswith("=")
        ]
    finally:
        wb.close()
    if not todo:
        return values
    cells = [f"{index_to_col(c)}{r + 1}" for r, c in todo]
    try:
        results = evaluate_cells(settings, str(base), cells, sheet, {})
    except SlimError as err:
        LOG.warning("values_only: could not calculate %d cell(s) of %s: %s", len(cells), sheet, err)
        return values
    for key, result in zip(todo, results, strict=True):
        values[key] = None if result == "" else result
    return values


def _cached_table(
    base: Path, req: RangeRequest
) -> tuple[tuple[str, ...], list[tuple[Any, ...]], int]:
    wb = openpyxl.load_workbook(base, data_only=True)
    try:
        ws = get_sheet(wb, req.sheet)
        block = find_table(ws, req.range, False)
        rng = block.rng
        headers = tuple(h or index_to_col(rng.left + i) for i, h in enumerate(block.headers))
        rows = [
            tuple(ws.cell(row=r + 1, column=c + 1).value for c in range(rng.left, rng.right + 1))
            for r in range(rng.top + 1, rng.bottom + 1)
        ]
        return headers, rows, rng.top + 2
    finally:
        wb.close()


def _copy_where(wb: Any, ws: Any, req: RangeRequest, matches: list[int]) -> Change:
    block = find_table(ws, req.range, False)
    dest_ws, top, left = _destination(wb, ws, req.to or "")
    rows = [block.rng.top + 1, *matches]
    for out, row in enumerate(rows):
        for offset, col in enumerate(range(block.rng.left, block.rng.right + 1)):
            source = ws.cell(row=row, column=col + 1)
            target = dest_ws.cell(row=top + out + 1, column=left + offset + 1)
            target.value = source.value
            target._style = copy_module.copy(source._style)
    where = f"{dest_ws.title}!{index_to_col(left)}{top + 1}"
    return Change(f"copied header + {len(matches)} row(s) matching {req.where} to {where}")


# --- copy / move / transpose --------------------------------------------------------------


def _destination(wb: Any, ws: Any, to: str) -> tuple[Any, int, int]:
    sheet, address = split_sheet(to)
    if sheet and not any(t.casefold() == sheet.casefold() for t in wb.sheetnames):
        dest = wb.create_sheet(sheet)
    else:
        dest = get_sheet(wb, sheet) if sheet else ws
    rng = parse_range(address)
    return dest, rng.top, rng.left


def _area(text: str | None) -> CellRange:
    rng = parse_range(text or "")
    if rng.n_rows * rng.n_cols > MAX_CELLS:
        raise SlimError(ErrorCode.TOO_LARGE, f"{text} has more than {MAX_CELLS} cells")
    return rng


def _put(source: Any, target: Any, translate: bool) -> None:
    value = source.value
    if translate and isinstance(value, str) and value.startswith("="):
        value = Translator(value, origin=source.coordinate).translate_formula(target.coordinate)
    target.value = value
    if source.has_style:
        target._style = copy_module.copy(source._style)


def _copy(wb: Any, ws: Any, req: RangeRequest, pre: Prepared) -> Change:
    if pre.matches is not None:
        return _copy_where(wb, ws, req, pre.matches)
    rng = _area(req.range)
    if req.formats_only:
        return _paint(wb, ws, req, rng)
    dest, top, left = _destination(wb, ws, req.to or "")
    missing = _copy_block(ws, rng, dest, top, left, transpose=False, cached=pre.cached)
    end = CellRange(top, left, top + rng.n_rows - 1, left + rng.n_cols - 1)
    kind = " (values only)" if pre.cached is not None else ""
    return Change(f"copied {ws.title}!{a1(rng)} to {dest.title}!{a1(end)}{kind}", _missing(missing))


def _missing(count: int) -> tuple[str, ...]:
    if not count:
        return ()
    return (
        f"{count} formula cell(s) had no calculated value (file never opened in Excel): pasted "
        "as empty; run recalc first or copy without values_only",
    )


def _copy_block(
    ws: Any,
    rng: CellRange,
    dest: Any,
    top: int,
    left: int,
    transpose: bool,
    cached: dict[tuple[int, int], Any] | None = None,
    translate: bool = True,
) -> int:
    """Copy cells with styles; `cached` pastes Excel's values instead of formulas.
    Returns how many formula cells had no cached value."""
    missing = 0
    cells = [
        [ws.cell(row=r + 1, column=c + 1) for c in range(rng.left, rng.right + 1)]
        for r in range(rng.top, rng.bottom + 1)
    ]
    snapshot = [[(cell.value, cell) for cell in row] for row in cells]  # copy before writing
    values = [[copy_module.copy(cell._style) for cell in row] for row in cells]
    for r, row in enumerate(snapshot):
        for c, (value, cell) in enumerate(row):
            rr, cc = (c, r) if transpose else (r, c)
            target = dest.cell(row=top + rr + 1, column=left + cc + 1)
            is_formula = isinstance(value, str) and value.startswith("=")
            if is_formula and cached is not None:
                value = cached.get((rng.top + r, rng.left + c))
                missing += value is None
            elif is_formula and translate and not transpose:
                value = Translator(value, origin=cell.coordinate).translate_formula(
                    target.coordinate
                )
            target.value = value
            target._style = values[r][c]
    return missing


def _move(wb: Any, ws: Any, req: RangeRequest, _pre: Prepared) -> Change:
    """Excel's cut/paste: the cells move, and every reference to them follows (on any sheet)."""
    rng = _area(req.range)
    dest, top, left = _destination(wb, ws, req.to or "")
    if dest is ws:
        ws.move_range(a1(rng), rows=top - rng.top, cols=left - rng.left, translate=False)
    else:
        _copy_block(ws, rng, dest, top, left, transpose=False, translate=False)
        for r in range(top, top + rng.n_rows):  # moved formulas still read their old sheet
            for c in range(left, left + rng.n_cols):
                cell = dest.cell(row=r + 1, column=c + 1)
                if isinstance(cell.value, str) and cell.value.startswith("="):
                    cell.value = qualify(cell.value, ws.title)
        for row in ws.iter_rows(
            min_row=rng.top + 1, max_row=rng.bottom + 1, min_col=rng.left + 1, max_col=rng.right + 1
        ):
            for cell in row:
                cell.value, cell.style = None, "Normal"
    move = Move(ws.title, rng, dest.title, top - rng.top, left - rng.left)
    repointed = _repoint(wb, move)
    end = CellRange(top, left, top + rng.n_rows - 1, left + rng.n_cols - 1)
    follow = f"; {repointed} formula(s)/name(s) now point at the new place" if repointed else ""
    return Change(f"moved {ws.title}!{a1(rng)} to {dest.title}!{a1(end)}{follow}")


def _repoint(wb: Any, move: Move) -> int:
    changed = 0
    for sheet in wb.worksheets:
        for row in sheet.iter_rows():
            for cell in row:
                value = cell.value
                if isinstance(value, str) and value.startswith("="):
                    new, did = move_references(value, sheet.title, move)
                    if did:
                        cell.value, changed = new, changed + 1
    for name in wb.defined_names.values():
        if name.attr_text:
            new, did = move_references("=" + name.attr_text, move.source_sheet, move)
            if did:
                name.attr_text, changed = new[1:], changed + 1
    return changed


def _transpose(wb: Any, ws: Any, req: RangeRequest, _pre: Prepared) -> Change:
    rng = _area(req.range)
    dest, top, left = _destination(wb, ws, req.to or "")
    missing = _copy_block(ws, rng, dest, top, left, transpose=True, cached=_pre.cached)
    end = CellRange(top, left, top + rng.n_cols - 1, left + rng.n_rows - 1)
    notes = _missing(missing)
    if _pre.cached is None:
        notes += (
            "formulas are copied as written when transposing; values_only=true pastes values",
        )
    return Change(f"transposed {ws.title}!{a1(rng)} to {dest.title}!{a1(end)}", notes)


# --- fill ---------------------------------------------------------------------------------


def _fill(wb: Any, ws: Any, req: RangeRequest, _pre: Prepared) -> Change:
    rng = _area(req.range)
    down = rng.n_rows > 1
    lines = range(rng.left, rng.right + 1) if down else range(rng.top, rng.bottom + 1)
    filled, series = 0, 0
    for line in lines:
        cells = (
            [ws.cell(row=r + 1, column=line + 1) for r in range(rng.top, rng.bottom + 1)]
            if down
            else [ws.cell(row=line + 1, column=c + 1) for c in range(rng.left, rng.right + 1)]
        )
        step = _series_step(cells)
        if step is not None:
            series += 1
            for i, cell in enumerate(cells[2:], start=2):
                cell.value = cells[0].value + step * i
                _style_from(cells[0], cell)
            filled += len(cells) - 2
            continue
        for cell in cells[1:]:
            _put(cells[0], cell, translate=True)
        filled += len(cells) - 1
    kind = f", {series} as a series" if series else ""
    where = "down" if down else "right"
    return Change(f"filled {filled} cell(s) {where} in {ws.title}!{a1(rng)}{kind}")


def _series_step(cells: list[Any]) -> Any:
    """Two numbers (or dates) on top and only blanks below -> the step between them."""
    if len(cells) < 3 or any(c.value not in (None, "") for c in cells[2:]):
        return None
    first, second = cells[0].value, cells[1].value
    numbers = all(isinstance(v, int | float) and not isinstance(v, bool) for v in (first, second))
    dates = all(isinstance(v, dt.date) for v in (first, second))
    if numbers or dates:
        return second - first
    return None


def _style_from(source: Any, target: Any) -> None:
    if source.has_style:
        target._style = copy_module.copy(source._style)


# --- sort ---------------------------------------------------------------------------------


def _sort(wb: Any, ws: Any, req: RangeRequest, _pre: Prepared) -> Change:
    block = find_table(ws, req.range, False)
    rng = block.rng
    if any(parse_range(m.coord).intersect(rng) for m in ws.merged_cells.ranges):
        raise SlimError(ErrorCode.BAD_ARGUMENT, f"{a1(rng)} contains merged cells; unmerge first")
    first = rng.top + 1 if req.header else rng.top
    key = _sort_column(block, req.by or "", req.header)
    rows = [
        [ws.cell(row=r + 1, column=c + 1) for c in range(rng.left, rng.right + 1)]
        for r in range(first, rng.bottom + 1)
    ]
    saved = [
        [(cell.value, copy_module.copy(cell._style), cell.coordinate) for cell in row]
        for row in rows
    ]
    order = sorted(range(len(saved)), key=lambda i: _sort_key(saved[i][key][0]), reverse=req.desc)
    if req.desc:  # keep blanks last in both directions
        order = [i for i in order if saved[i][key][0] not in (None, "")] + [
            i for i in order if saved[i][key][0] in (None, "")
        ]
    for target_row, source_index in zip(rows, order, strict=True):
        for cell, (value, style, origin) in zip(target_row, saved[source_index], strict=True):
            if isinstance(value, str) and value.startswith("="):
                value = Translator(value, origin=origin).translate_formula(cell.coordinate)
            cell.value, cell._style = value, style
    direction = "descending" if req.desc else "ascending"
    return Change(f"sorted {len(rows)} row(s) of {ws.title}!{a1(rng)} by {req.by} {direction}")


def _sort_column(block: TableBlock, by: str, header: bool) -> int:
    wanted = by.strip()
    if header:
        for i, name in enumerate(block.headers):
            if name.casefold() == wanted.casefold():
                return i
    if wanted.isalpha() and len(wanted) <= 3:
        index = int(col_to_index(wanted.upper())) - int(block.rng.left)
        if 0 <= index < block.rng.n_cols:
            return index
    raise SlimError(
        ErrorCode.BAD_ARGUMENT,
        f"No column '{by}' in the table",
        hint=f"Columns: {', '.join(h for h in block.headers if h)} (or a letter like 'C')",
    )


def _sort_key(value: Any) -> tuple[int, Any]:
    if value is None or value == "":
        return (3, "")
    if isinstance(value, bool):
        return (2, str(value))
    if isinstance(value, int | float):
        return (0, value)
    if isinstance(value, dt.datetime | dt.date):
        return (0, dt.datetime(value.year, value.month, value.day).timestamp())
    return (1, str(value).casefold())


# --- replace ------------------------------------------------------------------------------


def _replace(wb: Any, ws: Any, req: RangeRequest, _pre: Prepared) -> Change:
    rng = _area(req.range) if req.range else find_table(ws, None, False).rng
    pattern = re.compile(re.escape(req.find or ""), re.IGNORECASE)
    new = req.replace or ""
    changed = 0
    for row in ws.iter_rows(
        min_row=rng.top + 1, max_row=rng.bottom + 1, min_col=rng.left + 1, max_col=rng.right + 1
    ):
        for cell in row:
            value = cell.value
            if isinstance(value, str) and not value.startswith("=") and pattern.search(value):
                replaced = pattern.sub(lambda _m: new, value)
                cell.value = replaced if replaced != "" else None
                changed += 1
    return Change(
        f"replaced '{req.find}' with '{new}' in {changed} cell(s) of {ws.title}!{a1(rng)}"
    )


def _paint(wb: Any, ws: Any, req: RangeRequest, rng: CellRange) -> Change:
    """Format painter: the source block's styles tile over the target range (values untouched)."""
    sheet, address = split_sheet(req.to or "")
    dest = get_sheet(wb, sheet) if sheet else ws
    target = _area(address)
    if target.n_rows * target.n_cols == 1:  # a single cell means "the size of the source"
        target = CellRange(
            target.top, target.left, target.top + rng.n_rows - 1, target.left + rng.n_cols - 1
        )
    styles = [
        [
            copy_module.copy(ws.cell(row=r + 1, column=c + 1)._style)
            for c in range(rng.left, rng.right + 1)
        ]
        for r in range(rng.top, rng.bottom + 1)
    ]
    for r in range(target.top, target.bottom + 1):
        for c in range(target.left, target.right + 1):
            style = styles[(r - target.top) % rng.n_rows][(c - target.left) % rng.n_cols]
            dest.cell(row=r + 1, column=c + 1)._style = copy_module.copy(style)
    return Change(f"copied formats of {ws.title}!{a1(rng)} onto {dest.title}!{a1(target)}")


def _clear(wb: Any, ws: Any, req: RangeRequest, _pre: Prepared) -> Change:
    rng = _area(req.range)
    what = choice(req.what, CLEAR_WHAT, "what")
    for row in ws.iter_rows(
        min_row=rng.top + 1, max_row=rng.bottom + 1, min_col=rng.left + 1, max_col=rng.right + 1
    ):
        for cell in row:
            if what in ("contents", "all"):
                cell.value = None
            if what in ("formats", "all"):
                cell.style = "Normal"
            if what == "all":
                cell.comment, cell.hyperlink = None, None
    removed = _clear_rules(ws, rng) if what in ("rules", "all") else 0
    extra = f", {removed} rule(s) removed" if removed else ""
    return Change(f"cleared {what} of {ws.title}!{a1(rng)}{extra}")


def _clear_rules(ws: Any, rng: CellRange) -> int:
    """Drop data validations and conditional formats whose ranges lie inside `rng`."""
    removed = 0
    keep = []
    for dv in ws.data_validations.dataValidation:
        if all(_inside(parse_range(str(r)), rng) for r in dv.sqref.ranges):
            removed += 1
        else:
            keep.append(dv)
    ws.data_validations.dataValidation = keep
    for key in list(ws.conditional_formatting._cf_rules):
        if all(_inside(parse_range(str(r)), rng) for r in key.sqref.ranges):
            del ws.conditional_formatting._cf_rules[key]
            removed += 1
    return removed


def _inside(inner: CellRange, outer: CellRange) -> bool:
    return outer.intersect(inner) == inner


def _split(wb: Any, ws: Any, req: RangeRequest, _pre: Prepared) -> Change:
    """Text to columns: each cell's text is split into the cells to its right (overwritten)."""
    rng = _area(req.range)
    if rng.n_cols != 1:
        raise SlimError(ErrorCode.BAD_RANGE, "split takes one column", hint="e.g. range='A2:A200'")
    widest = 0
    replaced: list[str] = []
    for r in range(rng.top, rng.bottom + 1):
        cell = ws.cell(row=r + 1, column=rng.left + 1)
        if not isinstance(cell.value, str) or req.delimiter not in cell.value:
            continue
        parts = [part.strip() for part in cell.value.split(req.delimiter)]
        widest = max(widest, len(parts))
        for offset, part in enumerate(parts):
            target = ws.cell(row=r + 1, column=rng.left + 1 + offset)
            if offset and target.value not in (None, ""):
                replaced.append(target.coordinate)
            target.value = _typed(part)
    if not widest:
        return Change(
            f"nothing to split: no cell in {ws.title}!{a1(rng)} contains '{req.delimiter}'"
        )
    summary = f"split {ws.title}!{a1(rng)} on '{req.delimiter}' into up to {widest} column(s)"
    return Change(summary, warnings=_replaced_warning(replaced))


def _replaced_warning(replaced: list[str]) -> tuple[str, ...]:
    """Excel's Text to Columns asks before it overwrites; the tool reports what it replaced."""
    if not replaced:
        return ()
    shown = ", ".join(replaced[:5]) + (", ..." if len(replaced) > 5 else "")
    return (
        f"split overwrote {len(replaced)} non-empty cell(s) to the right ({shown}); "
        "undo, insert empty columns with range_ops(op='insert', range='C:D'), and split again",
    )


def _typed(text: str) -> Any:
    """'12' -> 12, '1.5' -> 1.5, '' -> None, other text unchanged (Excel's General format)."""
    if text == "":
        return None
    try:
        number = float(text)
    except ValueError:
        return text
    return (
        int(number)
        if number.is_integer() and "." not in text and "e" not in text.lower()
        else number
    )


_HANDLERS = {
    "fill": _fill,
    "copy": _copy,
    "move": _move,
    "transpose": _transpose,
    "sort": _sort,
    "replace": _replace,
    "insert": _insert,
    "delete": _delete,
    "clear": _clear,
    "split": _split,
}
