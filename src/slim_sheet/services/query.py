"""`query` use-case: read-only SQL over sheets/regions with offset paging (PLAN §5.1)."""

from __future__ import annotations

from collections.abc import Iterator
from contextlib import contextmanager

import duckdb

from slim_sheet.adapters.engine.duck import Engine, build_engine, schema_hint
from slim_sheet.adapters.readers.duck_reader import to_cell
from slim_sheet.config import Settings
from slim_sheet.core.budget import HEADROOM, clamp_budget, estimate_tokens
from slim_sheet.core.errors import ErrorCode, SlimError
from slim_sheet.core.models import CellValue, QueryResult
from slim_sheet.core.render_results import query_header, query_row
from slim_sheet.services.workbook import open_file

FETCH_ROWS = 500
MAX_SQL_CHARS = 8000


@contextmanager
def open_engine(path: str, settings: Settings, only: str | None = None) -> Iterator[Engine]:
    with open_file(path, settings) as handle:
        engine = build_engine(handle.path, handle.format, handle.reader, only)
    try:
        yield engine
    finally:
        engine.close()


def query(
    path: str, sql: str, settings: Settings, offset: int = 0, max_tokens: int | None = None
) -> QueryResult:
    statement = _check_sql(sql)
    if offset < 0:
        raise SlimError(ErrorCode.BAD_ARGUMENT, "offset must be >= 0")
    budget = clamp_budget(max_tokens, settings.max_tokens)
    with open_engine(path, settings, only=statement) as engine:
        try:
            return run_select(engine, statement, offset, budget, settings.max_cell_chars)
        except duckdb.Error as exc:
            raise SlimError(
                ErrorCode.BAD_SQL,
                _first_line(str(exc)),
                hint=f"Tables: {schema_hint(engine)}. Quote names with spaces: "
                'SELECT "Region", sum("Total") FROM "Orders" GROUP BY 1',
            ) from exc


def _check_sql(sql: str) -> str:
    text = sql.strip().rstrip(";").strip()
    if not text:
        raise SlimError(
            ErrorCode.BAD_SQL, "sql is empty", hint='e.g. SELECT * FROM "Sheet1" LIMIT 10'
        )
    if len(text) > MAX_SQL_CHARS:
        raise SlimError(ErrorCode.TOO_LARGE, f"sql is longer than {MAX_SQL_CHARS} characters")
    try:
        statements = duckdb.connect(":memory:").extract_statements(text)
    except duckdb.Error as exc:
        raise SlimError(ErrorCode.BAD_SQL, _first_line(str(exc))) from exc
    if len(statements) != 1:
        raise SlimError(ErrorCode.BAD_SQL, "Send exactly one SQL statement")
    if statements[0].type != duckdb.StatementType.SELECT:
        raise SlimError(
            ErrorCode.SQL_WRITE_DENIED,
            f"Only SELECT queries are allowed, got {statements[0].type.name}",
            hint="To change cells use the write tool.",
        )
    return text


def run_select(engine: Engine, sql: str, offset: int, budget: int, max_chars: int) -> QueryResult:
    con = engine.con
    total_row = con.execute(f"SELECT count(*) FROM ({sql}) AS q").fetchone()
    total = int(total_row[0]) if total_row else 0
    cursor = con.execute(f"SELECT * FROM ({sql}) AS q LIMIT {FETCH_ROWS} OFFSET {offset}")
    columns = tuple(str(d[0]) for d in cursor.description or [])
    fetched = cursor.fetchall()
    spent = estimate_tokens(query_header(columns)) + 40
    rows: list[tuple[CellValue, ...]] = []
    for raw in fetched:
        row = tuple(to_cell(v) for v in raw)
        cost = estimate_tokens(query_row(row, max_chars)) + 1
        if rows and spent + cost > budget * HEADROOM:
            break
        spent += cost
        rows.append(row)
    end = offset + len(rows)
    return QueryResult(sql, columns, tuple(rows), total, offset, end if end < total else None)


def _first_line(message: str) -> str:
    return message.strip().splitlines()[0] if message.strip() else "SQL error"
