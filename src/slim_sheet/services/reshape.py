"""`pivot` and `split` use-cases: reshape one table with DuckDB (PLAN §5.3).

`pivot` produces a static summary (values, not an Excel PivotTable object); the generated
SQL is returned so the agent can refine it with `query`.
"""

from __future__ import annotations

import re
from collections.abc import Iterator
from contextlib import contextmanager
from pathlib import Path
from typing import Any

from slim_sheet.adapters import journal
from slim_sheet.adapters.engine.cache import LoadedTable
from slim_sheet.adapters.engine.duck import Engine, engine_from, load_tables
from slim_sheet.adapters.sandbox import resolve_output
from slim_sheet.adapters.writers.export import export_grids, export_table
from slim_sheet.adapters.writers.pivot_native import PivotSpec, build_pivot, layout_values
from slim_sheet.config import Settings
from slim_sheet.core.a1 import format_range
from slim_sheet.core.budget import clamp_budget
from slim_sheet.core.errors import ErrorCode, SlimError
from slim_sheet.core.models import CellRange, CellValue, OutputResult, PivotResult, TableSchema
from slim_sheet.services.design import find_table
from slim_sheet.services.editing import Change, edit_workbook, get_sheet
from slim_sheet.services.output import default_output, journal_before_write, output_path
from slim_sheet.services.query import run_select
from slim_sheet.services.workbook import open_file

AGGREGATES = {
    "sum": "sum",
    "count": "count",
    "avg": "avg",
    "min": "min",
    "max": "max",
    "median": "median",
    "count_distinct": "count(DISTINCT {})",
}
MAX_PIVOT_COLUMNS = 50
MAX_GROUPS = 200
_UNSAFE_FILE = re.compile(r'[<>:"/\\|?*\x00-\x1f]')


def _q(name: str) -> str:
    return '"' + name.replace('"', '""') + '"'


@contextmanager
def _table_engine(
    path: str, sheet: str | None, settings: Settings
) -> Iterator[tuple[Engine, str, Any]]:
    with open_file(path, settings) as handle:
        tables, _ = load_tables(handle.path, handle.format, handle.reader, only=sheet)
        chosen = next(
            (t for t in tables if not sheet or t.name.casefold() == sheet.casefold()), None
        )
        if chosen is None:
            names = ", ".join(t.name for t in tables) or "none"
            raise SlimError(
                ErrorCode.SHEET_NOT_FOUND, f"No table '{sheet}'", hint=f"Tables: {names}"
            )
        label, source_path = f"{handle.display} / {chosen.name}", handle.path
    engine = engine_from([LoadedTable("src", chosen.source, chosen.table)])
    try:
        yield engine, label, source_path
    finally:
        engine.close()


def _column(table: TableSchema, name: str) -> str:
    for column, _ in table.columns:
        if column.casefold() == name.strip().casefold():
            return column
    raise SlimError(
        ErrorCode.BAD_ARGUMENT,
        f"Unknown column '{name}'",
        hint=f"Columns: {', '.join(c for c, _ in table.columns)}",
    )


# --- pivot --------------------------------------------------------------------------------


def pivot(
    path: str,
    rows: list[str],
    values: list[str],
    settings: Settings,
    agg: str = "sum",
    columns: str | None = None,
    sheet: str | None = None,
    to_sheet: str | None = None,
    out_path: str | None = None,
    overwrite: bool = False,
    max_tokens: int | None = None,
    native: bool = False,
) -> PivotResult:
    if native and not to_sheet:
        raise SlimError(
            ErrorCode.BAD_ARGUMENT,
            "native=true writes a real PivotTable into the workbook and needs to_sheet",
            hint="e.g. pivot(path, rows=['Region'], values=['Total'], native=true, "
            "to_sheet='Pivot')",
        )
    if agg not in AGGREGATES:
        raise SlimError(
            ErrorCode.BAD_ARGUMENT,
            f"Unknown agg '{agg}'",
            hint=f"Use one of: {', '.join(AGGREGATES)}",
        )
    if not rows or not values:
        raise SlimError(
            ErrorCode.BAD_ARGUMENT,
            "rows and values are required",
            hint="e.g. rows=['Region'], values=['Total']",
        )
    with _table_engine(path, sheet, settings) as (engine, label, _source):
        src = engine.table("src")
        row_cols = [_column(src, r) for r in rows]
        value_cols = [_column(src, v) for v in values]
        pivot_col = _column(src, columns) if columns else None
        sql = _pivot_sql(engine, row_cols, value_cols, agg, pivot_col)
        budget = clamp_budget(max_tokens, settings.max_tokens)
        preview = run_select(engine, sql, 0, budget, settings.max_cell_chars)
        full = engine.con.execute(sql).to_arrow_table() if (to_sheet or out_path) else None
    output = None
    if full is not None and out_path:
        target = output_path(out_path, settings, overwrite)
        backup = journal_before_write(target, f"pivot {label}")
        export_table(full, target, sheet="pivot")
        output = OutputResult(label, target.name, f"{full.num_rows} pivot rows", backup)
    elif full is not None and to_sheet and native:
        output = _native_to_sheet(
            path, sheet, rows, values, agg, columns, to_sheet, settings, overwrite
        )
    elif full is not None and to_sheet:
        output = _pivot_to_sheet(path, to_sheet, full, settings, overwrite)
    return PivotResult(label, sql, preview, output)


def _agg_expr(agg: str, column: str) -> str:
    template = AGGREGATES[agg]
    return template.format(_q(column)) if "{}" in template else f"{template}({_q(column)})"


def _pivot_sql(engine: Engine, rows: list[str], values: list[str], agg: str, on: str | None) -> str:
    group = ", ".join(_q(r) for r in rows)
    if on is None:
        aggs = ", ".join(f"{_agg_expr(agg, v)} AS {_q(f'{agg} of {v}')}" for v in values)
        return f"SELECT {group}, {aggs} FROM src GROUP BY ALL ORDER BY ALL"
    distinct = engine.con.execute(f"SELECT count(DISTINCT {_q(on)}) FROM src").fetchone()
    if distinct and int(distinct[0]) > MAX_PIVOT_COLUMNS:
        raise SlimError(
            ErrorCode.TOO_LARGE,
            f"'{on}' has {distinct[0]} distinct values (limit {MAX_PIVOT_COLUMNS} pivot columns)",
            hint="Pick a column with fewer values, or filter first with query.",
        )
    using = ", ".join(_agg_expr(agg, v) for v in values)
    return f"PIVOT src ON {_q(on)} USING {using} GROUP BY {group} ORDER BY ALL"


def _pivot_to_sheet(
    path: str, to_sheet: str, table: Any, settings: Settings, overwrite: bool
) -> OutputResult:
    rows: list[list[CellValue]] = [list(table.column_names)]
    rows += [[record[c] for c in table.column_names] for record in table.to_pylist()]

    def mutate(wb: Any) -> Change:
        if any(s.casefold() == to_sheet.casefold() for s in wb.sheetnames):
            raise SlimError(
                ErrorCode.BAD_ARGUMENT,
                f"Sheet '{to_sheet}' already exists",
                hint="Choose another to_sheet.",
            )
        ws = wb.create_sheet(to_sheet[:31])
        for row in rows:
            ws.append(row)
        return Change(f"pivot written to new sheet '{ws.title}' ({len(rows) - 1} rows)")

    return edit_workbook(path, settings, f"pivot to {to_sheet}", mutate, None, overwrite)


def _native_to_sheet(
    path: str,
    sheet: str | None,
    rows: list[str],
    values: list[str],
    agg: str,
    columns: str | None,
    to_sheet: str,
    settings: Settings,
    overwrite: bool,
) -> OutputResult:
    """A refreshable Excel PivotTable on a new sheet, with its values already filled in."""
    if len(rows) != 1 or len(values) > 3 or (columns and len(values) > 1):
        raise SlimError(
            ErrorCode.BAD_ARGUMENT,
            "native pivots take one row field and up to 3 values (1 value with columns)",
            hint="For more fields use the static pivot (native=false).",
        )

    def mutate(wb: Any) -> Change:
        if any(s.casefold() == to_sheet.casefold() for s in wb.sheetnames):
            raise SlimError(
                ErrorCode.BAD_ARGUMENT,
                f"Sheet '{to_sheet}' already exists",
                hint="Choose another to_sheet.",
            )
        source = get_sheet(wb, sheet) if sheet else wb.worksheets[0]
        spec = _native_spec(wb, source, rows[0], values, agg, columns)
        target = wb.create_sheet(to_sheet[:31])
        cache_id = 1 + max((p.cacheId for ws in wb.worksheets for p in ws._pivots), default=0)
        target._pivots.append(build_pivot(spec, cache_id, top=2, left=0))
        target["A1"] = f"PivotTable of {source.title} (refreshes when opened in Excel)"
        for r, line in enumerate(layout_values(spec), start=3):
            for c, value in enumerate(line, start=1):
                target.cell(row=r, column=c, value=value)
        where = f"{source.title}!{spec.source_ref}"
        return Change(f"native PivotTable '{spec.name}' on new sheet '{target.title}' from {where}")

    return edit_workbook(path, settings, f"pivot native to {to_sheet}", mutate, None, overwrite)


def _native_spec(
    wb: Any, ws: Any, row: str, values: list[str], agg: str, columns: str | None
) -> PivotSpec:
    block = find_table(ws, None, False)
    # the contiguous headed block: stray cells to the right (notes, totals) are not data
    width = next((i for i, h in enumerate(block.headers) if not h), len(block.headers))
    headers = block.headers[:width]

    def index(name: str) -> int:
        for i, header in enumerate(headers):
            if header.casefold() == name.strip().casefold():
                return i
        raise SlimError(
            ErrorCode.BAD_ARGUMENT,
            f"Unknown column '{name}'",
            hint=f"Columns: {', '.join(h for h in headers if h)}",
        )

    rng = CellRange(block.rng.top, block.rng.left, block.rng.bottom, block.rng.left + width - 1)
    data = tuple(
        tuple(ws.cell(row=r + 1, column=c + 1).value for c in range(rng.left, rng.right + 1))
        for r in range(rng.top + 1, rng.bottom + 1)
    )
    count = sum(len(w._pivots) for w in wb.worksheets)
    return PivotSpec(
        headers=headers, rows=data, source_sheet=ws.title, source_ref=format_range(rng),
        row_fields=(index(row),), col_field=index(columns) if columns else None,
        value_fields=tuple(index(v) for v in values), agg=agg, name=f"PivotTable{count + 1}",
    )  # fmt: skip


# --- split --------------------------------------------------------------------------------


def split(
    path: str,
    by: str,
    settings: Settings,
    sheet: str | None = None,
    mode: str = "sheets",
    out_path: str | None = None,
    overwrite: bool = False,
) -> OutputResult:
    if mode not in ("sheets", "files"):
        raise SlimError(
            ErrorCode.BAD_ARGUMENT, f"Unknown mode '{mode}'", hint="Use sheets or files."
        )
    with _table_engine(path, sheet, settings) as (engine, label, source):
        column = _column(engine.table("src"), by)
        groups = _groups(engine, column)
        parts = [(_group_label(value), _rows_for(engine, column, value)) for value in groups]
    if mode == "sheets":
        return _split_to_sheets(label, source, parts, settings, out_path, overwrite)
    return _split_to_files(label, source, parts, settings, out_path, overwrite)


def _groups(engine: Engine, column: str) -> list[Any]:
    found = engine.con.execute(
        f"SELECT DISTINCT {_q(column)} FROM src ORDER BY 1 NULLS LAST"
    ).fetchall()
    if len(found) > MAX_GROUPS:
        raise SlimError(
            ErrorCode.TOO_LARGE,
            f"'{column}' has {len(found)} distinct values (limit {MAX_GROUPS})",
            hint="Split by a coarser column.",
        )
    return [row[0] for row in found]


def _rows_for(engine: Engine, column: str, value: Any) -> Any:
    sql = f"SELECT * FROM src WHERE {_q(column)} IS NOT DISTINCT FROM ?"
    return engine.con.execute(sql, [value]).to_arrow_table()


def _group_label(value: Any) -> str:
    text = "(blank)" if value is None or value == "" else str(value)
    return _UNSAFE_FILE.sub("_", text).strip() or "(blank)"


def _grid(table: Any) -> list[list[CellValue]]:
    rows: list[list[CellValue]] = [list(table.column_names)]
    return rows + [[record[c] for c in table.column_names] for record in table.to_pylist()]


def _split_to_sheets(
    label: str,
    source: Path,
    parts: list[tuple[str, Any]],
    settings: Settings,
    out_path: str | None,
    overwrite: bool,
) -> OutputResult:
    target = output_path(out_path or default_output(source, "split", ".xlsx"), settings, overwrite)
    names: list[str] = []
    for name, _ in parts:
        base, n, candidate = name[:28], 2, name[:31]
        while candidate.casefold() in {x.casefold() for x in names}:
            candidate, n = f"{base}_{n}", n + 1
        names.append(candidate)
    backup = journal_before_write(target, f"split {label}")
    export_grids([(n, _grid(t)) for n, (_, t) in zip(names, parts, strict=True)], target)
    counts = ", ".join(f"{n}={t.num_rows}" for n, (_, t) in zip(names, parts, strict=True))
    return OutputResult(label, target.name, f"{len(parts)} sheet(s): {counts}", backup)


def _split_to_files(
    label: str,
    source: Path,
    parts: list[tuple[str, Any]],
    settings: Settings,
    out_path: str | None,
    overwrite: bool,
) -> OutputResult:
    raw = out_path or str(source.with_name(f"{source.stem}_split"))
    if Path(raw).suffix:
        raise SlimError(
            ErrorCode.BAD_ARGUMENT,
            "In files mode out_path is a folder",
            hint=f"e.g. out_path='{source.stem}_split'",
        )
    folder = resolve_output(raw, settings)  # inside the roots, parent folder exists
    folder.mkdir(exist_ok=True)
    suffix = (
        ".xlsx"
        if source.suffix.lower() in (".xlsx", ".xlsm", ".xls", ".xlsb", ".ods")
        else source.suffix
    )
    written = []
    for name, table in parts:
        target = folder / f"{source.stem}_{name}{suffix}"
        if target.exists() and not overwrite:
            raise SlimError(
                ErrorCode.WRITE_CONFLICT,
                f"'{target.name}' already exists",
                hint="Pass overwrite=true.",
            )
        journal.record(target, f"split {label}")
        export_table(table, target, sheet=name[:31])
        written.append(f"{target.name} ({table.num_rows})")
    return OutputResult(
        label, folder.name + "/", f"{len(written)} file(s): {', '.join(written)}", None
    )
