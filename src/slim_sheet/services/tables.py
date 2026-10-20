"""Resolve a user-supplied table name ('Orders', 'orders', 'Report.2', None) to a region table."""

from __future__ import annotations

from slim_sheet.adapters.engine.tables import RegionTable, extract_tables, find_table
from slim_sheet.core.errors import ErrorCode, SlimError
from slim_sheet.services.workbook import OpenFile


def resolve_table(handle: OpenFile, name: str | None) -> RegionTable:
    sheets = [n for n, _ in handle.reader.sheets()]
    candidates = _candidate_sheets(sheets, name)
    seen: list[str] = []
    for sheet in candidates:
        tables = extract_tables(handle.reader, sheet)
        seen += [t.name for t in tables]
        table = find_table(
            tables, None if name is None or name.casefold() == sheet.casefold() else name
        )
        if table is not None:
            return table
    raise SlimError(
        ErrorCode.SHEET_NOT_FOUND,
        f"No table '{name}' in {handle.display}" if name else f"{handle.display} has no data",
        hint=f"Tables: {', '.join(seen) or 'none'}; sheets: {', '.join(sheets)}",
    )


def _candidate_sheets(sheets: list[str], name: str | None) -> list[str]:
    if name is None:
        return sheets
    folded = name.casefold()
    exact = [s for s in sheets if s.casefold() == folded]
    prefix = [s for s in sheets if folded.startswith(s.casefold() + ".") and s not in exact]
    return exact + prefix or sheets
