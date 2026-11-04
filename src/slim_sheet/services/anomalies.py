"""`anomalies` use-case: suspicious cells with A1 addresses, most severe first."""

from __future__ import annotations

from collections import Counter

from slim_sheet.adapters.columnar import SheetStore
from slim_sheet.config import Settings
from slim_sheet.core.a1 import index_to_col
from slim_sheet.core.anomalies import find_anomalies
from slim_sheet.core.errors import ErrorCode, SlimError
from slim_sheet.core.models import AnomalyResult, Finding
from slim_sheet.services.anomalies_columnar import columnar_anomalies
from slim_sheet.services.tables import big_store, resolve_table
from slim_sheet.services.workbook import open_file

MAX_RESULTS_CAP = 200


def anomalies(
    path: str,
    settings: Settings,
    sheet: str | None = None,
    columns: list[str] | None = None,
    max_results: int = 20,
) -> AnomalyResult:
    limit = max(1, min(max_results, MAX_RESULTS_CAP))
    with open_file(path, settings) as handle:
        store = big_store(handle, sheet)
        if store is not None:  # a big sheet: the checks run as SQL over its Parquet cache
            return _big(handle.display, store, columns, limit)
        table = resolve_table(handle, sheet)
        lookup = {name.casefold(): i for i, name in enumerate(table.columns)}
        missing = [c for c in columns or () if c.casefold() not in lookup]
        if missing:
            raise SlimError(
                ErrorCode.BAD_ARGUMENT,
                f"Unknown column(s): {', '.join(missing)}",
                hint=f"Columns of {table.name}: {', '.join(table.columns)}",
            )
        only = [lookup[c.casefold()] for c in columns] if columns else None
        found = find_anomalies(table.columns, table.body, only)
        findings = tuple(
            Finding(
                address=table.address(a.row, a.col),
                value=a.value,
                kind=a.kind,
                detail=f"{table.columns[a.col]}: {a.detail}",
            )
            for a in found[:limit]
        )
        counts = tuple(Counter(a.kind for a in found).items())
        return AnomalyResult(
            handle.display, table.name, table.source, len(table.body), findings, counts
        )


def _big(display: str, store: SheetStore, columns: list[str] | None, limit: int) -> AnomalyResult:
    only = _column_indexes(store.columns, store.sheet, columns)
    found, counts = columnar_anomalies(store, only, limit)
    findings = tuple(
        Finding(
            address=f"{store.sheet}!{index_to_col(store.used.left + a.col)}"
            f"{store.body_top + a.row + 1}",
            value=a.value,
            kind=a.kind,
            detail=f"{store.columns[a.col]}: {a.detail}",
        )
        for a in found[:limit]
    )
    return AnomalyResult(display, store.sheet, store.source, store.rows, findings,
                         tuple(counts.items()))  # fmt: skip


def _column_indexes(
    names: tuple[str, ...], table: str, wanted: list[str] | None
) -> list[int] | None:
    if not wanted:
        return None
    lookup = {name.casefold(): i for i, name in enumerate(names)}
    missing = [c for c in wanted if c.casefold() not in lookup]
    if missing:
        raise SlimError(
            ErrorCode.BAD_ARGUMENT,
            f"Unknown column(s): {', '.join(missing)}",
            hint=f"Columns of {table}: {', '.join(names)}",
        )
    return [lookup[c.casefold()] for c in wanted]
