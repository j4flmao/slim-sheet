"""Header and column-type detection on a small sample (no I/O)."""

from __future__ import annotations

import datetime as dt
from collections.abc import Sequence

from slim_sheet.core.models import CellValue

HEADER_SCAN_ROWS = 10


def _non_empty(row: Sequence[CellValue]) -> list[CellValue]:
    return [v for v in row if v is not None and v != ""]


def guess_header_row(rows: Sequence[Sequence[CellValue]]) -> int | None:
    """Index (within `rows`) of the likely header row, or None.

    A header is a mostly-filled row of unique text labels that is followed either by
    non-text data or, for all-text tables, is simply the first filled row.
    """
    first_filled: int | None = None
    for index, row in enumerate(rows[:HEADER_SCAN_ROWS]):
        filled = _non_empty(row)
        if not filled:
            continue
        if first_filled is None:
            first_filled = index
        if len(filled) * 2 < len(row) or not _is_label_row(filled):
            continue
        below = [v for r in rows[index + 1 : index + 6] for v in _non_empty(r)]
        if not below:
            continue
        if any(not isinstance(v, str) for v in below) or index == first_filled:
            return index
    return None


def _is_label_row(filled: Sequence[CellValue]) -> bool:
    return all(isinstance(v, str) for v in filled) and len(set(filled)) == len(filled)


def infer_type(values: Sequence[CellValue]) -> str:
    kinds = {kind_of(v) for v in values if v is not None and v != ""}
    if not kinds:
        return "empty"
    if kinds == {"int"}:
        return "int"
    if kinds <= {"int", "float"}:
        return "float"
    if len(kinds) == 1:
        return kinds.pop()
    if kinds <= {"date", "datetime"}:
        return "datetime"
    return "mixed"


def kind_of(value: CellValue) -> str:
    if isinstance(value, bool):
        return "bool"
    if isinstance(value, int):
        return "int"
    if isinstance(value, float):
        return "int" if value.is_integer() else "float"
    if isinstance(value, dt.datetime):
        return "datetime"
    if isinstance(value, dt.date):
        return "date"
    if isinstance(value, dt.time):
        return "time"
    if isinstance(value, dt.timedelta):
        return "duration"
    return "text"
