"""Cell grids -> typed Arrow columns, using the same type inference as `overview`."""

from __future__ import annotations

import datetime as dt
from collections.abc import Sequence

import pyarrow as pa

from slim_sheet.core.budget import cell_text
from slim_sheet.core.detect import infer_type
from slim_sheet.core.models import CellValue

_NO_TRUNCATE = 1_000_000


def to_arrow(names: Sequence[str], rows: Sequence[Sequence[CellValue]]) -> pa.Table:
    columns = {}
    for index, name in enumerate(names):
        values = [row[index] if index < len(row) else None for row in rows]
        values = [None if v == "" else v for v in values]
        columns[name] = _column(values)
    return pa.table(columns)


def _column(values: list[CellValue]) -> pa.Array:
    kind = infer_type(values)
    if kind == "int":
        return pa.array([None if v is None else int(v) for v in values], pa.int64())  # type: ignore[arg-type]
    if kind == "float":
        return pa.array([None if v is None else float(v) for v in values], pa.float64())  # type: ignore[arg-type]
    if kind == "bool":
        return pa.array(values, pa.bool_())
    if kind == "date":
        return pa.array(values, pa.date32())
    if kind == "datetime":
        return pa.array([_as_datetime(v) for v in values], pa.timestamp("us"))
    if kind == "time":
        return pa.array(values, pa.time64("us"))
    if kind == "duration":
        return pa.array(values, pa.duration("us"))
    return pa.array(
        [None if v is None else _text(v) for v in values],
        pa.string(),
    )


def _as_datetime(value: CellValue) -> dt.datetime | None:
    if value is None or isinstance(value, dt.datetime):
        return value
    if isinstance(value, dt.date):
        return dt.datetime.combine(value, dt.time())
    raise TypeError(f"Expected date/datetime, got {type(value).__name__}")


def _text(value: CellValue) -> str:
    return value if isinstance(value, str) else cell_text(value, _NO_TRUNCATE)
