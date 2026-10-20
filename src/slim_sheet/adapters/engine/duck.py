"""Sandboxed DuckDB session over one file (PLAN ADR-3, §8).

Workbook sheets are split into regions; the largest region of a sheet is registered
under the sheet name, further regions as "<sheet>.2", "<sheet>.3", ... Flat files are
loaded as table "data". Parsed tables are cached per file (see cache.py). After loading,
file access and configuration are locked, so SQL cannot read other files or re-enable access.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path

import duckdb

from slim_sheet.adapters.engine.arrow import to_arrow
from slim_sheet.adapters.engine.cache import TABLE_CACHE, LoadedTable, file_key
from slim_sheet.adapters.engine.tables import extract_tables
from slim_sheet.adapters.readers import SheetReader
from slim_sheet.adapters.readers.duck_reader import SCAN, SHEET
from slim_sheet.core.errors import ErrorCode, SlimError
from slim_sheet.core.models import FileFormat, TableSchema


@dataclass
class Engine:
    con: duckdb.DuckDBPyConnection
    tables: list[TableSchema] = field(default_factory=list)
    unloaded: list[str] = field(default_factory=list)

    def table(self, name: str) -> TableSchema:
        for schema in self.tables:
            if schema.name.casefold() == name.casefold():
                return schema
        names = ", ".join(t.name for t in self.tables) or "none"
        raise SlimError(
            ErrorCode.SHEET_NOT_FOUND,
            f"No table '{name}'",
            hint=f"Tables: {names}",
        )

    def close(self) -> None:
        self.con.close()


def load_tables(
    path: Path, fmt: FileFormat, reader: SheetReader, only: str | None = None
) -> tuple[list[LoadedTable], list[str]]:
    """Arrow tables for a file (cached); with `only`, just sheets named in that text."""
    key = file_key(path)
    if not fmt.is_workbook:
        cached = TABLE_CACHE.get(key, SHEET)
        if cached is None:
            cached = [_load_flat(path, fmt)]
            TABLE_CACHE.put(key, SHEET, cached)
        return cached, []
    names = [sheet for sheet, _visible in reader.sheets()]
    wanted = only.casefold() if only else ""
    selected = [s for s in names if s.casefold() in wanted] or names
    loaded: list[LoadedTable] = []
    for sheet in selected:
        tables = TABLE_CACHE.get(key, sheet)
        if tables is None:
            tables = [
                LoadedTable(t.name, t.source, to_arrow(t.columns, t.body))
                for t in extract_tables(reader, sheet)
            ]
            TABLE_CACHE.put(key, sheet, tables)
        loaded.extend(tables)
    return loaded, [s for s in names if s not in selected]


def build_engine(
    path: Path, fmt: FileFormat, reader: SheetReader, only: str | None = None
) -> Engine:
    loaded, unloaded = load_tables(path, fmt, reader, only)
    return engine_from(loaded, unloaded)


def engine_from(tables: list[LoadedTable], unloaded: list[str] | None = None) -> Engine:
    """Register tables (zero-copy) in a fresh connection, then lock file access."""
    engine = Engine(duckdb.connect(":memory:"), unloaded=list(unloaded or []))
    try:
        for loaded in tables:
            engine.con.register(loaded.name, loaded.table)
            engine.tables.append(_describe(engine.con, loaded.name, loaded.source))
        engine.con.execute("SET enable_external_access = false")
        engine.con.execute("SET lock_configuration = true")
    except BaseException:
        engine.close()
        raise
    return engine


def _load_flat(path: Path, fmt: FileFormat) -> LoadedTable:
    con = duckdb.connect(":memory:")
    try:
        table = con.execute(f"SELECT * FROM {SCAN[fmt]}", [str(path)]).to_arrow_table()
    except duckdb.Error as exc:
        raise SlimError(ErrorCode.READ_FAILED, f"Failed to load '{path.name}': {exc}") from exc
    finally:
        con.close()
    return LoadedTable(SHEET, path.name, table)


def _describe(con: duckdb.DuckDBPyConnection, name: str, source: str) -> TableSchema:
    quoted = '"' + name.replace('"', '""') + '"'
    cols = con.execute(f"DESCRIBE {quoted}").fetchall()
    count = con.execute(f"SELECT count(*) FROM {quoted}").fetchone()
    columns = tuple((str(c[0]), str(c[1])) for c in cols)
    return TableSchema(name, source, columns, int(count[0]) if count else 0)


def schema_hint(engine: Engine, limit_chars: int = 900) -> str:
    """Compact table/column listing used in SQL error hints."""
    parts = []
    for t in engine.tables:
        cols = ", ".join(f'"{c}" {ty}' for c, ty in t.columns)
        parts.append(f'"{t.name}" ({t.rows} rows): {cols}')
    text = "; ".join(parts)
    if engine.unloaded:
        text += f"; other sheets (name them in SQL to load): {', '.join(engine.unloaded)}"
    return text if len(text) <= limit_chars else text[:limit_chars] + "..."
