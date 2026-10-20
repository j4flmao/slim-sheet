"""`anomalies` use-case: suspicious cells with A1 addresses, most severe first."""

from __future__ import annotations

from collections import Counter

from slim_sheet.config import Settings
from slim_sheet.core.anomalies import find_anomalies
from slim_sheet.core.errors import ErrorCode, SlimError
from slim_sheet.core.models import AnomalyResult, Finding
from slim_sheet.services.tables import resolve_table
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
