"""Typed, explainable cleaning operations over one table (PLAN §9). Pure, no I/O.

Ops always run in canonical order regardless of the order requested. Column-level
conversions (numbers, dates) only fire when >= CONVERT_SHARE of a column's non-empty
values parse, so ID-like text columns are left alone.
"""

from __future__ import annotations

import datetime as dt
import re
import unicodedata
from collections.abc import Callable, Sequence
from dataclasses import dataclass, field

from slim_sheet.core.errors import ErrorCode, SlimError
from slim_sheet.core.models import CellValue
from slim_sheet.core.textparse import date_parts, parse_date, parse_number

CLEAN_OPS = (
    "remove_nonprintable",
    "trim",
    "collapse_spaces",
    "unify_null",
    "numbers_from_text",
    "normalize_dates",
    "drop_empty_rows",
    "drop_empty_cols",
    "dedupe",
)
NULL_TOKENS = {"", "n/a", "na", "null", "none", "nil", "-", "--", "nan", "#n/a", "?"}
CONVERT_SHARE = 0.8
MAX_SAMPLES = 8
_SPACES = re.compile(r"\s+")


@dataclass(frozen=True, slots=True)
class CellChange:
    row: int  # index into the ORIGINAL body
    col: int  # index into the ORIGINAL columns
    before: CellValue
    after: CellValue
    op: str


@dataclass
class CleanOutcome:
    columns: list[str]
    rows: list[list[CellValue]]
    counts: dict[str, int] = field(default_factory=dict)
    samples: list[CellChange] = field(default_factory=list)
    notes: list[str] = field(default_factory=list)


def check_ops(ops: Sequence[str]) -> list[str]:
    unknown = [op for op in ops if op not in CLEAN_OPS]
    if unknown or not ops:
        raise SlimError(
            ErrorCode.BAD_ARGUMENT,
            f"Unknown clean op(s): {', '.join(unknown)}" if unknown else "ops is empty",
            hint=f"Choose from: {', '.join(CLEAN_OPS)}",
        )
    return [op for op in CLEAN_OPS if op in ops]


def clean_table(
    columns: Sequence[str],
    body: Sequence[Sequence[CellValue]],
    ops: Sequence[str],
    keys: Sequence[str] | None = None,
) -> CleanOutcome:
    ordered = check_ops(ops)
    out = CleanOutcome(list(columns), [list(r) + [None] * (len(columns) - len(r)) for r in body])
    row_ids = list(range(len(out.rows)))
    col_ids = list(range(len(columns)))
    for op in ordered:
        if op in _CELL_OPS:
            _map_cells(out, row_ids, col_ids, op, _CELL_OPS[op])
        elif op == "numbers_from_text":
            _convert_columns(out, row_ids, col_ids, op, _number_converter)
        elif op == "normalize_dates":
            _convert_columns(out, row_ids, col_ids, op, _date_converter)
        elif op == "drop_empty_rows":
            _drop_rows(out, row_ids, op, lambda i: all(_empty(v) for v in out.rows[i]))
        elif op == "drop_empty_cols":
            _drop_empty_cols(out, col_ids)
        elif op == "dedupe":
            _dedupe(out, row_ids, keys)
    return out


def _empty(value: CellValue) -> bool:
    return value is None or (isinstance(value, str) and value.strip() == "")


def _record(out: CleanOutcome, change: CellChange) -> None:
    out.counts[change.op] = out.counts.get(change.op, 0) + 1
    if len(out.samples) < MAX_SAMPLES:
        out.samples.append(change)


def _map_cells(
    out: CleanOutcome,
    row_ids: list[int],
    col_ids: list[int],
    op: str,
    fn: Callable[[str], CellValue],
) -> None:
    for r, row in enumerate(out.rows):
        for c, value in enumerate(row):
            if isinstance(value, str):
                new = fn(value)
                if new != value:
                    row[c] = new
                    _record(out, CellChange(row_ids[r], col_ids[c], value, new, op))


def _nonprintable(text: str) -> str:
    return "".join(ch for ch in text if ch in "\t\n" or unicodedata.category(ch)[0] != "C")


def _unify_null(text: str) -> CellValue:
    return None if text.strip().casefold() in NULL_TOKENS else text


_CELL_OPS: dict[str, Callable[[str], CellValue]] = {
    "remove_nonprintable": _nonprintable,
    "trim": lambda s: s.strip(),
    "collapse_spaces": lambda s: _SPACES.sub(" ", s),
    "unify_null": _unify_null,
}

Converter = Callable[[list[str]], "tuple[Callable[[str], CellValue], str | None] | None"]


def _convert_columns(
    out: CleanOutcome, row_ids: list[int], col_ids: list[int], op: str, make: Converter
) -> None:
    for c in range(len(out.columns)):
        texts = [v for r in out.rows if isinstance(v := r[c], str) and v.strip()]
        if not texts:
            continue
        built = make(texts)
        if built is None:
            continue
        convert, note = built
        ok = sum(1 for t in texts if convert(t) is not None)
        if ok / len(texts) < CONVERT_SHARE:
            continue
        if note:
            out.notes.append(f"{out.columns[c]}: {note}")
        for r, row in enumerate(out.rows):
            value = row[c]
            if isinstance(value, str) and value.strip():
                new = convert(value)
                if new is not None:
                    row[c] = new
                    _record(out, CellChange(row_ids[r], col_ids[c], value, new, op))


def _number_converter(texts: list[str]) -> tuple[Callable[[str], CellValue], str | None]:
    return parse_number, None


def _date_converter(texts: list[str]) -> tuple[Callable[[str], CellValue], str | None] | None:
    parts = [p for p in (date_parts(t) for t in texts) if p is not None]
    day_first, note = True, None
    if not any(a > 12 for a, _, _ in parts):  # a first part > 12 proves day-first
        if any(b > 12 for _, b, _ in parts):
            day_first = False
        elif parts:
            note = "ambiguous numeric dates assumed day-first (DD/MM/YYYY)"

    def convert(text: str) -> dt.date | dt.datetime | None:
        return parse_date(text, day_first)

    return convert, note


def _drop_rows(
    out: CleanOutcome, row_ids: list[int], op: str, should_drop: Callable[[int], bool]
) -> None:
    keep = [i for i in range(len(out.rows)) if not should_drop(i)]
    dropped = len(out.rows) - len(keep)
    if dropped:
        out.counts[op] = out.counts.get(op, 0) + dropped
    out.rows[:] = [out.rows[i] for i in keep]
    row_ids[:] = [row_ids[i] for i in keep]


def _drop_empty_cols(out: CleanOutcome, col_ids: list[int]) -> None:
    keep = [c for c in range(len(out.columns)) if any(not _empty(r[c]) for r in out.rows)]
    dropped = [out.columns[c] for c in range(len(out.columns)) if c not in keep]
    if dropped:
        out.counts["drop_empty_cols"] = len(dropped)
        out.notes.append("dropped empty columns: " + ", ".join(dropped))
    out.columns[:] = [out.columns[c] for c in keep]
    out.rows[:] = [[r[c] for c in keep] for r in out.rows]
    col_ids[:] = [col_ids[c] for c in keep]


def _dedupe(out: CleanOutcome, row_ids: list[int], keys: Sequence[str] | None) -> None:
    lookup = {name.casefold(): i for i, name in enumerate(out.columns)}
    missing = [k for k in keys or () if k.casefold() not in lookup]
    if missing:
        raise SlimError(
            ErrorCode.BAD_ARGUMENT,
            f"Unknown key column(s): {', '.join(missing)}",
            hint=f"Columns: {', '.join(out.columns)}",
        )
    idx = [lookup[k.casefold()] for k in keys] if keys else list(range(len(out.columns)))
    seen: set[tuple[object, ...]] = set()

    def duplicate(i: int) -> bool:
        key = tuple(_dedupe_key(out.rows[i][c]) for c in idx)
        if key in seen:
            return True
        seen.add(key)
        return False

    _drop_rows(out, row_ids, "dedupe", duplicate)


def _dedupe_key(value: CellValue) -> object:
    return value.strip().casefold() if isinstance(value, str) else value
