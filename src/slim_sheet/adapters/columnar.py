"""Columnar cache for big sheets: parse once, then answer every call from Parquet (PLAN §11.2).

A workbook sheet with hundreds of thousands of rows costs seconds and gigabytes to parse on
every call. The first time a big sheet is opened it is converted once, in batches of
BATCH_ROWS, into a Parquet file in the user's cache folder; every later read, query, profile,
pivot or export of that sheet reads the Parquet file instead (windows by row group, SQL by
streaming scan). The cache key is the file's path, size and modification time, so an edited
workbook is converted again, and old versions are deleted.

The first HEAD_ROWS rows are kept exactly (titles, header, typed values) next to the Parquet
body, so overviews and reads near the top see the sheet as it is.
"""

from __future__ import annotations

import datetime as dt
import json
import logging
import threading
from collections.abc import Iterator
from dataclasses import dataclass
from pathlib import Path
from typing import Any

import duckdb
import pyarrow as pa
import pyarrow.compute as pc
import pyarrow.dataset as ds
import pyarrow.parquet as pq

from slim_sheet.adapters.facts import cache_dir, version_key
from slim_sheet.adapters.readers.duck_reader import SCAN
from slim_sheet.core.a1 import format_range, index_to_col
from slim_sheet.core.budget import cell_text
from slim_sheet.core.detect import guess_header_row
from slim_sheet.core.errors import ErrorCode, SlimError
from slim_sheet.core.models import CellRange, CellValue, FileFormat
from slim_sheet.core.regions import column_names

log = logging.getLogger(__name__)

BIG_SHEET_CELLS = 500_000  # used-range cells from which a sheet is served from the cache
HEAD_ROWS = 300
BATCH_ROWS = 50_000
FLAT_CACHE_BYTES = 50 * 2**20  # text files from this size are cached as Parquet
MAX_CACHE_BYTES = 20 * 2**30  # oldest cache files are removed past this size
_lock = threading.Lock()


@dataclass(frozen=True, slots=True)
class SheetStore:
    """A big sheet as exact head rows plus a typed Parquet body."""

    sheet: str
    used: CellRange
    head: tuple[tuple[CellValue, ...], ...]  # sheet rows used.top .. used.top + len(head) - 1
    body_top: int  # 0-based sheet row of the first Parquet row
    columns: tuple[str, ...]
    has_header: bool
    rows: int  # Parquet rows
    parquet: Path
    merged: int = 0  # merged ranges in the sheet

    @property
    def source(self) -> str:
        return f"{self.sheet}!{format_range(self.used)}"

    def dataset(self) -> ds.Dataset:
        return ds.dataset(self.parquet, format="parquet")

    def read(self, window: CellRange) -> list[list[CellValue]]:
        """Exactly window.n_rows rows of window.n_cols values, from the head or the body."""
        out: list[list[CellValue]] = []
        head_end = self.used.top + len(self.head)  # first sheet row not in the head
        for row in range(window.top, min(window.bottom + 1, head_end)):
            line = self.head[row - self.used.top] if row >= self.used.top else ()
            out.append(
                [_at(line, c - self.used.left) for c in range(window.left, window.right + 1)]
            )
        first = max(window.top, head_end)
        if first <= window.bottom:
            out += self._body_rows(first - self.body_top, window.bottom - first + 1, window)
        return out

    def _body_rows(self, start: int, count: int, window: CellRange) -> list[list[CellValue]]:
        width = window.n_cols
        if start >= self.rows:
            return [[None] * width for _ in range(count)]
        left = window.left - self.used.left
        wanted = [i for i in range(left, left + width) if 0 <= i < len(self.columns)]
        table = _slice(self.parquet, start, count, [self.columns[i] for i in wanted])
        values = {
            i: table.column(n).to_pylist() for i, n in zip(wanted, table.column_names, strict=True)
        }
        got = table.num_rows
        rows = [[values[i][r] if i in values else None for i in range(left, left + width)]
                for r in range(got)]  # fmt: skip
        return rows + [[None] * width for _ in range(count - got)]


def _at(line: tuple[CellValue, ...], index: int) -> CellValue:
    return line[index] if 0 <= index < len(line) else None


def _slice(path: Path, start: int, count: int, columns: list[str]) -> pa.Table:
    """Rows [start, start+count) reading only the row groups that hold them."""
    file = pq.ParquetFile(path)
    groups: list[int] = []
    first_row = offset = 0
    for index in range(file.num_row_groups):
        size = file.metadata.row_group(index).num_rows
        if first_row + size > start and first_row < start + count:
            if not groups:
                offset = start - first_row
            groups.append(index)
        first_row += size
    if not groups:
        return pa.table({c: pa.array([], pa.null()) for c in columns})
    table = file.read_row_groups(groups, columns=columns)
    return table.slice(offset, count)


# --- cache lookup and build ------------------------------------------------------------------


def _key(path: Path, sheet: str) -> tuple[str, str]:
    return version_key(path, sheet)


def _sheets_dir() -> Path:
    return cache_dir() / "sheets"


def cached_store(path: Path, sheet: str) -> SheetStore | None:
    """The cache of this exact file version, or None (nothing is parsed)."""
    stem, _ = _key(path, sheet)
    meta = _sheets_dir() / f"{stem}.json"
    parquet = meta.with_suffix(".parquet")
    if not (meta.exists() and parquet.exists()):
        return None
    try:
        return _from_meta(json.loads(meta.read_text(encoding="utf-8")), sheet, parquet)
    except (OSError, ValueError, KeyError, TypeError) as exc:
        log.warning("Ignoring unreadable sheet cache %s: %s", meta, exc)
        return None


def build_store(
    path: Path, sheet: str, calamine_sheet: Any, used: CellRange, merged: int = 0
) -> SheetStore:
    """Convert a loaded calamine sheet into the cache, in batches; returns the store."""
    with _lock:
        found = cached_store(path, sheet)
        if found is not None:
            return found
        folder = _sheets_dir()
        folder.mkdir(parents=True, exist_ok=True)
        stem, prefix = _key(path, sheet)
        head = _head(calamine_sheet, used)
        header_idx = guess_header_row(head)
        body_top = used.top + (header_idx + 1 if header_idx is not None else 0)
        letters = [index_to_col(c) for c in range(used.left, used.right + 1)]
        header = head[header_idx] if header_idx is not None else None
        names = column_names(header, letters)
        target = folder / f"{stem}.parquet"
        rows = _write(target, names, _batches(calamine_sheet, used, body_top))
        store = SheetStore(sheet, used, tuple(tuple(r) for r in head), body_top, tuple(names),
                           header_idx is not None, rows, target, merged)  # fmt: skip
        (folder / f"{stem}.json").write_text(json.dumps(_to_meta(store)), encoding="utf-8")
        _prune(folder, prefix, stem)
        log.info("Cached %s!%s: %d rows in %s", path.name, sheet, rows, target)
        return store


def flat_parquet(path: Path, fmt: FileFormat) -> Path:
    """A big CSV/TSV/JSON converted once to Parquet by DuckDB in one streaming pass, so later
    calls scan it in milliseconds instead of parsing the text again (cached per file version)."""
    stem, owner = version_key(path, "flat")
    folder = cache_dir() / "flat"
    target = folder / f"{stem}.parquet"
    if target.exists():
        return target
    folder.mkdir(parents=True, exist_ok=True)
    temp = target.with_suffix(".tmp")
    con = duckdb.connect(":memory:")
    try:
        quoted = str(temp).replace("'", "''")
        con.execute(f"COPY (SELECT * FROM {SCAN[fmt]}) TO '{quoted}' (FORMAT parquet, "
                    "COMPRESSION zstd)", [str(path)])  # fmt: skip
    except duckdb.Error as exc:
        temp.unlink(missing_ok=True)
        raise SlimError(ErrorCode.READ_FAILED, f"Failed to load '{path.name}': {exc}") from exc
    finally:
        con.close()
    temp.replace(target)
    for stale in folder.glob(f"{owner}-*.parquet"):
        if stale != target:
            stale.unlink(missing_ok=True)
    return target


def flat_store(path: Path, fmt: FileFormat) -> SheetStore | None:
    """A big text file as a store over its Parquet cache (one sheet 'data', header in row 1)."""
    if fmt == FileFormat.PARQUET:
        parquet = path
    elif path.stat().st_size >= FLAT_CACHE_BYTES:
        parquet = flat_parquet(path, fmt)
    else:
        return None
    meta = pq.ParquetFile(parquet).metadata
    names = tuple(meta.schema.to_arrow_schema().names)
    rows = meta.num_rows
    if rows * max(1, len(names)) < BIG_SHEET_CELLS:
        return None
    used = CellRange(0, 0, rows, max(0, len(names) - 1))
    return SheetStore("data", used, (names,), 1, names, True, rows, parquet)


def is_big(used: CellRange) -> bool:
    return used.n_rows * used.n_cols >= BIG_SHEET_CELLS


def _head(calamine_sheet: Any, used: CellRange) -> list[list[CellValue]]:
    rows = calamine_sheet.to_python(skip_empty_area=False, nrows=used.top + HEAD_ROWS)
    return [_window_row(r, used) for r in rows[used.top : used.top + HEAD_ROWS]]


def _window_row(row: list[CellValue], used: CellRange) -> list[CellValue]:
    cells = row[used.left : used.right + 1]
    cells += [None] * (used.n_cols - len(cells))
    return [None if v == "" else v for v in cells]


def _batches(calamine_sheet: Any, used: CellRange, body_top: int) -> Iterator[list[pa.Array]]:
    """Body rows as Arrow columns, BATCH_ROWS at a time. iter_rows counts rows from row 1 but
    holds only the used columns. Each column is typed by Arrow in C, not cell by cell."""
    width, rows = used.n_cols, []
    for index, row in enumerate(calamine_sheet.iter_rows()):
        if index >= body_top:
            rows.append(row)
            if len(rows) == BATCH_ROWS:
                yield _columns(rows, width)
                rows = []
    if rows:
        yield _columns(rows, width)


def _columns(rows: list[list[CellValue]], width: int) -> list[pa.Array]:
    padded = rows if all(len(r) == width for r in rows) else [_pad(r, width) for r in rows]
    return [_array(list(column)) for column in zip(*padded, strict=True)]


def _pad(row: list[CellValue], width: int) -> list[CellValue]:
    return list(row[:width]) + [None] * (width - len(row))


def _array(values: list[CellValue]) -> pa.Array:
    """Arrow infers the type; blanks ('') become nulls; mixed columns become text."""
    try:
        return _integral(pa.array(values))
    except (pa.ArrowInvalid, pa.ArrowTypeError):
        pass
    cleaned = [None if v == "" else v for v in values]
    try:
        return _integral(pa.array(cleaned))
    except (pa.ArrowInvalid, pa.ArrowTypeError):
        return pa.array([None if v is None else _as_text(v) for v in cleaned], pa.string())


def _integral(array: pa.Array) -> pa.Array:
    """Excel stores every number as a float: whole-number columns become integers; an all-text
    column keeps '' as null like every other reader."""
    if pa.types.is_floating(array.type) and len(array) and array.null_count < len(array):
        valid = array.drop_null()
        if (
            pc.all(pc.equal(valid, pc.floor(valid))).as_py()
            and pc.max(pc.abs(valid)).as_py() < 2**53
        ):
            return pc.cast(array, pa.int64())
    if pa.types.is_string(array.type) or pa.types.is_large_string(array.type):
        return pc.if_else(pc.equal(array, ""), pa.scalar(None, array.type), array)
    return array


def _as_text(value: CellValue) -> str:
    return value if isinstance(value, str) else cell_text(value, 1_000_000)


def _unified(types: list[pa.DataType]) -> pa.DataType:
    """One type for a column whose batches were typed separately."""
    real = {t for t in types if not pa.types.is_null(t)}
    if not real:
        return pa.string()
    if len(real) == 1:
        return real.pop()
    if all(pa.types.is_integer(t) or pa.types.is_floating(t) for t in real):
        return pa.float64()
    if all(pa.types.is_date(t) or pa.types.is_timestamp(t) for t in real):
        return pa.timestamp("us")
    return pa.string()


def _write(target: Path, names: list[str], batches: Iterator[list[pa.Array]]) -> int:
    """Hold the typed batches (columnar, compact), unify each column's type, write Parquet."""
    held = list(batches)
    if not held:
        held = [[pa.array([], pa.string()) for _ in names]]
    types = [_unified([batch[i].type for batch in held]) for i in range(len(names))]
    schema = pa.schema([pa.field(n, t) for n, t in zip(names, types, strict=True)])
    temp = target.with_suffix(".tmp")
    total = 0
    with pq.ParquetWriter(temp, schema, compression="zstd") as writer:
        while held:
            batch = held.pop(0)
            arrays = [_cast(a, t) for a, t in zip(batch, types, strict=True)]
            record = pa.RecordBatch.from_arrays(arrays, schema=schema)
            writer.write_batch(record, row_group_size=BATCH_ROWS)
            total += record.num_rows
    temp.replace(target)
    return total


def _cast(array: pa.Array, target: pa.DataType) -> pa.Array:
    if array.type == target:
        return array
    if pa.types.is_string(target) and not pa.types.is_null(array.type):
        return pa.array([None if v is None else _as_text(v) for v in array.to_pylist()], target)
    return pc.cast(array, target)


def _prune(folder: Path, prefix: str, keep: str) -> None:
    """Drop older versions of this sheet, then the oldest files past MAX_CACHE_BYTES."""
    for stale in folder.glob(f"{prefix}-*"):
        if not stale.name.startswith(keep):
            stale.unlink(missing_ok=True)
    files = sorted(folder.glob("*.parquet"), key=lambda p: p.stat().st_mtime, reverse=True)
    total = 0
    for file in files:
        total += file.stat().st_size
        if total > MAX_CACHE_BYTES and not file.name.startswith(keep):
            file.unlink(missing_ok=True)
            file.with_suffix(".json").unlink(missing_ok=True)


# --- head rows as JSON (typed) ---------------------------------------------------------------


def _to_meta(store: SheetStore) -> dict[str, object]:
    used = store.used
    return {
        "used": [used.top, used.left, used.bottom, used.right],
        "head": [[_encode(v) for v in row] for row in store.head],
        "body_top": store.body_top,
        "columns": list(store.columns),
        "has_header": store.has_header,
        "rows": store.rows,
        "merged": store.merged,
    }


def _from_meta(meta: dict[str, Any], sheet: str, parquet: Path) -> SheetStore:
    top, left, bottom, right = meta["used"]
    head = tuple(tuple(_decode(v) for v in row) for row in meta["head"])
    return SheetStore(sheet, CellRange(top, left, bottom, right), head, int(meta["body_top"]),
                      tuple(meta["columns"]), bool(meta["has_header"]), int(meta["rows"]),
                      parquet, int(meta.get("merged", 0)))  # fmt: skip


def _encode(value: CellValue) -> object:
    if isinstance(value, dt.datetime):
        return {"dt": value.isoformat()}
    if isinstance(value, dt.date):
        return {"d": value.isoformat()}
    if isinstance(value, dt.time):
        return {"t": value.isoformat()}
    if isinstance(value, dt.timedelta):
        return {"s": value.total_seconds()}
    return value


def _decode(value: object) -> CellValue:
    if isinstance(value, dict):
        if "dt" in value:
            return dt.datetime.fromisoformat(value["dt"])
        if "d" in value:
            return dt.date.fromisoformat(value["d"])
        if "t" in value:
            return dt.time.fromisoformat(value["t"])
        return dt.timedelta(seconds=float(value["s"]))
    return value  # type: ignore[return-value]
