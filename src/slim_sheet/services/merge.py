"""`merge` use-case: join two tables (same or different files) and report what did not match.

Keys compare as trimmed text, so 1001 (number) matches "1001" (text) — the most common
cause of failed lookups between exports. `keys` items are 'Col' or 'LeftCol=RightCol'.
"""

from __future__ import annotations

from collections.abc import Iterator
from contextlib import ExitStack, contextmanager
from pathlib import Path

import duckdb

from slim_sheet.adapters.engine.cache import LoadedTable
from slim_sheet.adapters.engine.duck import Engine, engine_from, load_tables
from slim_sheet.adapters.sandbox import resolve_path
from slim_sheet.adapters.writers.export import export_table
from slim_sheet.config import Settings
from slim_sheet.core.budget import clamp_budget
from slim_sheet.core.errors import ErrorCode, SlimError
from slim_sheet.core.models import MergeResult, OutputResult, TableSchema
from slim_sheet.services.output import journal_before_write, output_path
from slim_sheet.services.query import run_select
from slim_sheet.services.workbook import open_file

HOWS = ("left", "inner", "full", "anti", "reconcile")
STATUSES = ("different", "only_left", "only_right", "matched")


def _q(name: str) -> str:
    return '"' + name.replace('"', '""') + '"'


@contextmanager
def _pair_engine(
    settings: Settings, left: tuple[str, str | None], right: tuple[str, str | None]
) -> Iterator[tuple[Engine, str, str]]:
    labels = []
    chosen: list[LoadedTable] = []
    with ExitStack() as stack:
        for alias, (path, sheet) in zip(("a", "b"), (left, right), strict=True):
            handle = stack.enter_context(open_file(path, settings))
            tables, _ = load_tables(handle.path, handle.format, handle.reader, only=sheet)
            table = next(
                (t for t in tables if not sheet or t.name.casefold() == sheet.casefold()), None
            )
            if table is None:
                names = ", ".join(t.name for t in tables) or "none"
                raise SlimError(
                    ErrorCode.SHEET_NOT_FOUND,
                    f"No table '{sheet}' in {handle.display}",
                    hint=f"Tables: {names}",
                )
            chosen.append(LoadedTable(alias, table.source, table.table))
            labels.append(f"{handle.display} / {table.name}")
    engine = engine_from(chosen)
    try:
        yield engine, labels[0], labels[1]
    finally:
        engine.close()


def merge(
    path: str,
    other_path: str,
    keys: list[str],
    settings: Settings,
    how: str = "left",
    sheet: str | None = None,
    other_sheet: str | None = None,
    out_path: str | None = None,
    overwrite: bool = False,
    max_tokens: int | None = None,
    compare: list[str] | None = None,
    tolerance: float = 0.0,
) -> MergeResult:
    """how='reconcile' lines both tables up by key and labels every row matched / different /
    only_left / only_right, comparing `compare` columns (default: same-named non-key columns;
    numbers within `tolerance` count as equal)."""
    if how not in HOWS:
        raise SlimError(
            ErrorCode.BAD_ARGUMENT, f"Unknown how '{how}'", hint=f"Use one of: {', '.join(HOWS)}"
        )
    if not keys:
        raise SlimError(
            ErrorCode.BAD_ARGUMENT, "keys is empty", hint="e.g. keys=['OrderID'] or ['ID=Order No']"
        )
    base = resolve_path(path, settings).parent
    target = output_path(out_path, settings, overwrite, base) if out_path else None
    with _pair_engine(settings, (path, sheet), (other_path, other_sheet)) as (engine, left, right):
        a, b = engine.table("a"), engine.table("b")
        pairs = [_key_pair(k, a, b) for k in keys]
        if how == "reconcile":
            sql = _reconcile_sql(a, b, pairs, _compared(a, b, pairs, compare), tolerance)
        else:
            sql = _join_sql(a, b, pairs, how)
        cond = " AND ".join(f"{_txt('a', lk)} = {_txt('b', rk)}" for lk, rk in pairs)
        con = engine.con
        try:
            unmatched_a = _count(
                con, f"SELECT count(*) FROM a WHERE NOT EXISTS (SELECT 1 FROM b WHERE {cond})"
            )
            unmatched_b = _count(
                con, f"SELECT count(*) FROM b WHERE NOT EXISTS (SELECT 1 FROM a WHERE {cond})"
            )
            dup_keys = _count(con, _dup_sql(pairs))
            statuses = _statuses(con, sql) if how == "reconcile" else ()
            preview = run_select(
                engine,
                sql,
                0,
                clamp_budget(max_tokens, settings.max_tokens),
                settings.max_cell_chars,
            )
            output = _export(con, sql, target, f"merge {left} + {right}", left) if target else None
        except duckdb.Error as exc:
            raise SlimError(ErrorCode.BAD_SQL, f"Merge failed: {str(exc).splitlines()[0]}") from exc
    warnings: tuple[str, ...] = statuses
    if dup_keys and how != "anti":
        warnings += (
            f"{dup_keys} key value(s) repeat in the right table, so matching rows are multiplied",
        )
    return MergeResult(
        left,
        right,
        how,
        tuple(keys),
        a.rows,
        b.rows,
        unmatched_a,
        unmatched_b,
        preview,
        output,
        warnings,
    )


def _key_pair(raw: str, a: TableSchema, b: TableSchema) -> tuple[str, str]:
    left_raw, _, right_raw = raw.partition("=")
    left_name = _column(a, left_raw.strip(), "left")
    right_name = _column(b, (right_raw or left_raw).strip(), "right")
    return left_name, right_name


def _column(table: TableSchema, name: str, side: str) -> str:
    for column, _ in table.columns:
        if column.casefold() == name.casefold():
            return column
    raise SlimError(
        ErrorCode.BAD_ARGUMENT,
        f"Key '{name}' not found in the {side} table",
        hint=f"{side} columns: {', '.join(c for c, _ in table.columns)}. "
        "Different names: keys=['A=B']",
    )


def _txt(alias: str, column: str) -> str:
    return f"trim(CAST({alias}.{_q(column)} AS VARCHAR))"


def _join_sql(a: TableSchema, b: TableSchema, pairs: list[tuple[str, str]], how: str) -> str:
    cond = " AND ".join(f"{_txt('a', lk)} = {_txt('b', rk)}" for lk, rk in pairs)
    if how == "anti":
        return f"SELECT a.* FROM a WHERE NOT EXISTS (SELECT 1 FROM b WHERE {cond})"
    types_b = dict(b.columns)
    right_keys = {rk for _, rk in pairs}
    select = []
    for name, dtype in a.columns:
        pair = next((rk for lk, rk in pairs if lk == name), None)
        if how == "full" and pair is not None:
            same = dtype == types_b[pair]
            left_expr = f"a.{_q(name)}" if same else _txt("a", name)
            right_expr = f"b.{_q(pair)}" if same else _txt("b", pair)
            select.append(f"COALESCE({left_expr}, {right_expr}) AS {_q(name)}")
        else:
            select.append(f"a.{_q(name)}")
    taken = {name.casefold() for name, _ in a.columns}
    for name, _ in b.columns:
        if name in right_keys:
            continue
        alias, n = name, 2
        while alias.casefold() in taken:
            alias, n = f"{name}_{n}", n + 1
        taken.add(alias.casefold())
        select.append(f"b.{_q(name)} AS {_q(alias)}")
    join = {"left": "LEFT JOIN", "inner": "JOIN", "full": "FULL OUTER JOIN"}[how]
    return f"SELECT {', '.join(select)} FROM a {join} b ON {cond}"


def _dup_sql(pairs: list[tuple[str, str]]) -> str:
    cols = ", ".join(_txt("b", rk) for _, rk in pairs)
    return f"SELECT count(*) FROM (SELECT {cols} FROM b GROUP BY ALL HAVING count(*) > 1)"


def _count(con: duckdb.DuckDBPyConnection, sql: str) -> int:
    row = con.execute(sql).fetchone()
    return int(row[0]) if row else 0


def _export(
    con: duckdb.DuckDBPyConnection, sql: str, target: Path, note: str, source: str
) -> OutputResult:
    table = con.execute(sql).to_arrow_table()
    backup = journal_before_write(target, note)
    export_table(table, target, sheet="merged")
    return OutputResult(source, target.name, f"{table.num_rows} merged rows", backup)


# --- reconcile ----------------------------------------------------------------------------

_NUMERIC = ("TINYINT", "SMALLINT", "INTEGER", "BIGINT", "HUGEINT", "FLOAT", "DOUBLE", "DECIMAL")


def _compared(
    a: TableSchema, b: TableSchema, pairs: list[tuple[str, str]], wanted: list[str] | None
) -> list[tuple[str, str, bool]]:
    """(left column, right column, numeric) to compare."""
    keys_a, keys_b = {lk for lk, _ in pairs}, {rk for _, rk in pairs}
    types_a, types_b = dict(a.columns), dict(b.columns)
    if wanted:
        chosen = [_key_pair(w, a, b) for w in wanted]
    else:
        right = {name.casefold(): name for name in types_b if name not in keys_b}
        chosen = [
            (n, right[n.casefold()]) for n in types_a if n not in keys_a and n.casefold() in right
        ]
    if not chosen:
        raise SlimError(
            ErrorCode.BAD_ARGUMENT,
            "Nothing to compare: the tables share no columns besides the keys",
            hint="Name them: compare=['Amount'] or compare=['Amount=Total']",
        )
    return [(lc, rc, _numeric(types_a[lc]) and _numeric(types_b[rc])) for lc, rc in chosen]


def _numeric(dtype: str) -> bool:
    return dtype.upper().startswith(_NUMERIC)


def _reconcile_sql(
    a: TableSchema,
    b: TableSchema,
    pairs: list[tuple[str, str]],
    compared: list[tuple[str, str, bool]],
    tolerance: float,
) -> str:
    cond = " AND ".join(f"{_txt('a', lk)} = {_txt('b', rk)}" for lk, rk in pairs)
    differs = []
    columns = [f"COALESCE({_txt('a', lk)}, {_txt('b', rk)}) AS {_q(lk)}" for lk, rk in pairs]
    for lc, rc, numeric in compared:
        left, right = f"a.{_q(lc)}", f"b.{_q(rc)}"
        if numeric:
            differs.append(
                f"abs(coalesce({left}, 0) - coalesce({right}, 0)) > {float(tolerance)!r}"
            )
            delta = f"coalesce({left}, 0) - coalesce({right}, 0)"
            columns += [
                f"{left} AS {_q(lc + ' (left)')}",
                f"{right} AS {_q(rc + ' (right)')}",
                f"{delta} AS {_q(lc + ' diff')}",
            ]
        else:
            differs.append(f"coalesce({_txt('a', lc)}, '') <> coalesce({_txt('b', rc)}, '')")
            columns += [f"{left} AS {_q(lc + ' (left)')}", f"{right} AS {_q(rc + ' (right)')}"]
    status = (
        "CASE WHEN a.__in_a IS NULL THEN 'only_right' WHEN b.__in_b IS NULL THEN 'only_left' "
        f"WHEN {' OR '.join(differs)} THEN 'different' ELSE 'matched' END"
    )
    order = " ".join(f"WHEN '{s}' THEN {i}" for i, s in enumerate(STATUSES))
    return (
        f"SELECT {status} AS status, {', '.join(columns)} "
        "FROM (SELECT *, 1 AS __in_a FROM a) a FULL OUTER JOIN (SELECT *, 1 AS __in_b FROM b) b "
        f"ON {cond} ORDER BY CASE status {order} END"
    )


def _statuses(con: duckdb.DuckDBPyConnection, sql: str) -> tuple[str, ...]:
    rows = con.execute(f"SELECT status, count(*) FROM ({sql}) GROUP BY 1").fetchall()
    counts = {str(status): int(n) for status, n in rows}
    return ("reconcile: " + ", ".join(f"{s} {counts.get(s, 0)}" for s in STATUSES),)
