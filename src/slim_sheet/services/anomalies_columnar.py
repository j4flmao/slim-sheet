"""`anomalies` on a big sheet: the same six checks as core/anomalies.py, run by DuckDB over the
sheet's Parquet cache instead of cell by cell in Python, so a million rows take seconds and
little memory. Rules and thresholds are shared with the in-memory version; only the first
`limit` findings of each kind are fetched, while the counts cover every row.
"""

from __future__ import annotations

from collections.abc import Sequence

import duckdb

from slim_sheet.adapters.columnar import SheetStore
from slim_sheet.core.anomalies import (
    _ID_NAME,
    DOMINANT_SHARE,
    KINDS,
    MAX_CATEGORIES,
    MIN_NUMERIC,
    MIN_VALUES,
    ROBUST_Z,
    Anomaly,
)
from slim_sheet.core.models import CellValue

_NUMERIC = ("TINYINT", "SMALLINT", "INTEGER", "BIGINT", "HUGEINT", "FLOAT", "DOUBLE", "DECIMAL")


def columnar_anomalies(
    store: SheetStore, only: Sequence[int] | None, limit: int
) -> tuple[list[Anomaly], dict[str, int]]:
    """(the first findings in KINDS order, total count per kind)."""
    con = duckdb.connect(":memory:")
    try:
        source = str(store.parquet).replace("'", "''")  # our own cache file, not user input
        con.execute(
            f"CREATE VIEW t AS SELECT * FROM read_parquet('{source}', file_row_number = true)"
        )
        types = {str(r[0]): str(r[1]) for r in con.execute("DESCRIBE t").fetchall()}
        found: list[Anomaly] = []
        counts: dict[str, int] = {}
        for c in only if only is not None else range(len(store.columns)):
            name = store.columns[c]
            for kind, rows, total in _column_checks(con, name, types.get(name, ""), limit):
                counts[kind] = counts.get(kind, 0) + total
                found += [Anomaly(r, c, v, kind, d) for r, v, d in rows]
    finally:
        con.close()
    order = {kind: i for i, kind in enumerate(KINDS)}
    found.sort(key=lambda a: (order[a.kind], a.col, a.row))
    return found, {k: counts[k] for k in KINDS if counts.get(k)}


Finding = tuple[int, CellValue, str]


def _column_checks(
    con: duckdb.DuckDBPyConnection, name: str, sql_type: str, limit: int
) -> list[tuple[str, list[Finding], int]]:
    col = '"' + name.replace('"', '""') + '"'
    filled = f"{col} IS NOT NULL" + (f" AND {col} <> ''" if sql_type == "VARCHAR" else "")
    n = _one(con, f"SELECT count(*) FROM t WHERE {filled}")
    if not n:
        return []
    checks = [_duplicates(con, name, col, filled, n, limit)]
    if sql_type == "VARCHAR":
        checks += [
            _type_mismatch(con, col, filled, n, limit),
            _variants(con, col, filled, n, limit),
            _whitespace(con, col, limit),
        ]
    if sql_type.startswith(_NUMERIC):
        checks.append(_outliers(con, col, n, limit))
    if sql_type in ("DATE", "TIMESTAMP"):
        checks.append(_date_range(con, col, limit))
    return [c for c in checks if c is not None]


def _one(con: duckdb.DuckDBPyConnection, sql: str, params: list[object] | None = None) -> int:
    row = con.execute(sql, params or []).fetchone()
    return int(row[0]) if row and row[0] is not None else 0


def _rows(
    con: duckdb.DuckDBPyConnection, where: str, col: str, detail: str, limit: int
) -> tuple[list[Finding], int]:
    total = _one(con, f"SELECT count(*) FROM t WHERE {where}")
    fetched = con.execute(
        f"SELECT file_row_number, {col}, {detail} FROM t WHERE {where} "
        f"ORDER BY file_row_number LIMIT {limit}"
    ).fetchall()
    return [(int(r), v, str(d)) for r, v, d in fetched], total


def _duplicates(
    con: duckdb.DuckDBPyConnection, name: str, col: str, filled: str, n: int, limit: int
) -> tuple[str, list[Finding], int] | None:
    key = f"lower(trim(regexp_replace(CAST({col} AS VARCHAR), '\\s+', ' ', 'g')))"
    distinct = _one(con, f"SELECT count(DISTINCT {key}) FROM t WHERE {filled}")
    if n < MIN_VALUES or not (_ID_NAME.search(name) or distinct / n >= 0.95):
        return None
    con.execute(
        f"CREATE OR REPLACE TEMP VIEW dup AS SELECT file_row_number, {col} AS v, "
        f"min(file_row_number) OVER (PARTITION BY {key}) AS first FROM t WHERE {filled}"
    )
    where = "file_row_number IN (SELECT file_row_number FROM dup WHERE file_row_number > first)"
    rows, total = _rows(con, where, col, "'x'", limit)
    firsts = dict(con.execute(
        "SELECT file_row_number, first FROM dup WHERE file_row_number > first "
        f"ORDER BY file_row_number LIMIT {limit}"
    ).fetchall())  # fmt: skip
    return "duplicate_key", [(r, v, f"same as body row {firsts[r] + 1}") for r, v, _ in rows], total


def _type_mismatch(
    con: duckdb.DuckDBPyConnection, col: str, filled: str, n: int, limit: int
) -> tuple[str, list[Finding], int] | None:
    numeric = f"TRY_CAST({col} AS DOUBLE) IS NOT NULL"
    numbers = _one(con, f"SELECT count(*) FROM t WHERE {filled} AND {numeric}")
    if n < MIN_VALUES or numbers in (0, n):
        return None
    dominant, other, where = (
        ("number", "text", f"{filled} AND NOT {numeric}")
        if numbers >= n - numbers
        else ("text", "number", f"{filled} AND {numeric}")
    )
    if max(numbers, n - numbers) / n < DOMINANT_SHARE:
        return None
    rows, total = _rows(con, where, col, f"'{other} in a mostly {dominant} column'", limit)
    return "type_mismatch", rows, total


def _outliers(
    con: duckdb.DuckDBPyConnection, col: str, n: int, limit: int
) -> tuple[str, list[Finding], int] | None:
    if n < MIN_NUMERIC:
        return None
    x = f"CAST({col} AS DOUBLE)"
    median = con.execute(f"SELECT median({x}) FROM t").fetchone()[0]  # type: ignore[index]
    mad = con.execute(f"SELECT median(abs({x} - ?)) FROM t", [median]).fetchone()[0]  # type: ignore[index]
    if mad and mad > 0:
        z = f"0.6745 * ({x} - {median!r}) / {mad!r}"
        detail = f"'robust z=' || round({z}, 1) || ', median ' || {median!r}"
        rows, total = _rows(con, f"abs({z}) > {ROBUST_Z}", col, detail, limit)
        return "outlier", rows, total
    q1, q3 = con.execute(
        f"SELECT quantile_cont({x}, 0.25), quantile_cont({x}, 0.75) FROM t"
    ).fetchone()  # type: ignore[misc]
    iqr = (q3 or 0) - (q1 or 0)
    if iqr <= 0:
        return None
    low, high = q1 - 3 * iqr, q3 + 3 * iqr
    rows, total = _rows(con, f"{x} NOT BETWEEN {low!r} AND {high!r}", col,
                        f"'outside [{low:g}, {high:g}]'", limit)  # fmt: skip
    return "outlier", rows, total


def _variants(
    con: duckdb.DuckDBPyConnection, col: str, filled: str, n: int, limit: int
) -> tuple[str, list[Finding], int] | None:
    norm = f"lower(trim(regexp_replace({col}, '\\s+', ' ', 'g')))"
    groups = _one(con, f"SELECT count(DISTINCT {norm}) FROM t WHERE {filled}")
    if n < MIN_VALUES or groups > MAX_CATEGORIES:
        return None
    con.execute(
        f"CREATE OR REPLACE TEMP TABLE canon AS SELECT k, arg_max(v, c) AS main FROM ("
        f"SELECT {norm} AS k, {col} AS v, count(*) AS c FROM t WHERE {filled} GROUP BY 1, 2"
        ") GROUP BY k"
    )
    where = f"{filled} AND trim({col}) <> (SELECT main FROM canon WHERE canon.k = {norm})"
    detail = (
        f"'spelled differently from ''' || (SELECT main FROM canon WHERE canon.k = {norm}) || ''''"
    )
    rows, total = _rows(con, where, col, detail, limit)
    return "variant", rows, total


def _whitespace(
    con: duckdb.DuckDBPyConnection, col: str, limit: int
) -> tuple[str, list[Finding], int]:
    detail = "'leading/trailing spaces (breaks joins and lookups)'"
    rows, total = _rows(con, f"{col} <> trim({col})", col, detail, limit)
    return "whitespace", rows, total


def _date_range(
    con: duckdb.DuckDBPyConnection, col: str, limit: int
) -> tuple[str, list[Finding], int]:
    where = f"year({col}) NOT BETWEEN 1900 AND 2100"
    rows, total = _rows(con, where, col, f"'year ' || year({col}) || ' is implausible'", limit)
    return "date_range", rows, total
