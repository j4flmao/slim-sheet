"""`profile` use-case: per-column statistics via DuckDB SUMMARIZE + top values."""

from __future__ import annotations

from collections.abc import Sequence

import duckdb

from slim_sheet.config import Settings
from slim_sheet.core.budget import cell_text
from slim_sheet.core.errors import ErrorCode, SlimError
from slim_sheet.core.models import ColumnProfile, ProfileResult, TableSchema
from slim_sheet.services.query import open_engine

MAX_COLUMNS = 30
TOP_VALUES = 3
STAT_CHARS = 24


def profile(
    path: str, settings: Settings, sheet: str | None = None, columns: Sequence[str] | None = None
) -> ProfileResult:
    with open_engine(path, settings, only=sheet) as engine:
        if not engine.tables:
            raise SlimError(ErrorCode.READ_FAILED, f"'{path}' has no data to profile")
        table = engine.table(sheet) if sheet else engine.tables[0]
        chosen = _pick_columns(table, columns)
        stats = [
            _column(engine.con, table.name, name, dtype) for name, dtype in chosen[:MAX_COLUMNS]
        ]
        return ProfileResult(
            path=path,
            table=table.name,
            source=table.source,
            rows=table.rows,
            columns=tuple(stats),
            omitted_columns=max(0, len(chosen) - MAX_COLUMNS),
        )


def _pick_columns(table: TableSchema, wanted: Sequence[str] | None) -> list[tuple[str, str]]:
    if not wanted:
        return list(table.columns)
    by_name = {name.casefold(): (name, dtype) for name, dtype in table.columns}
    missing = [w for w in wanted if w.casefold() not in by_name]
    if missing:
        raise SlimError(
            ErrorCode.BAD_ARGUMENT,
            f"Unknown column(s) in '{table.name}': {', '.join(missing)}",
            hint=f"Columns: {', '.join(n for n, _ in table.columns)}",
        )
    return [by_name[w.casefold()] for w in wanted]


def _q(name: str) -> str:
    return '"' + name.replace('"', '""') + '"'


def _column(con: duckdb.DuckDBPyConnection, table: str, name: str, dtype: str) -> ColumnProfile:
    row = con.execute(
        f"SELECT min, max, avg, q50, count, null_percentage, "
        f"(SELECT count(DISTINCT {_q(name)}) FROM {_q(table)}) "  # SUMMARIZE's count is approximate
        f"FROM (SUMMARIZE SELECT {_q(name)} FROM {_q(table)})"
    ).fetchone()
    if row is None:
        raise SlimError(ErrorCode.READ_FAILED, f"Could not summarize column '{name}'")
    count, distinct = int(row[4] or 0), int(row[6] or 0)
    top = []
    if distinct < count:  # an all-unique column has no meaningful "top" values
        top = con.execute(
            f"SELECT CAST({_q(name)} AS VARCHAR) AS v, count(*) AS n FROM {_q(table)} "
            f"WHERE {_q(name)} IS NOT NULL GROUP BY 1 ORDER BY 2 DESC, 1 LIMIT {TOP_VALUES}"
        ).fetchall()
    return ColumnProfile(
        name=name,
        dtype=dtype,
        count=count,
        null_pct=float(row[5] or 0),
        distinct=distinct,
        min=_short(row[0]),
        max=_short(row[1]),
        mean=_short(row[2]),
        median=_short(row[3]),
        top=tuple((cell_text(v, STAT_CHARS), int(n)) for v, n in top),
    )


def _short(value: object) -> str | None:
    if value is None:
        return None
    text = str(value)
    try:
        number = float(text)
    except ValueError:
        return cell_text(text, STAT_CHARS)
    return f"{number:.6g}"
