"""Resolve a user-supplied table name ('Orders', 'orders', 'Report.2', None) to a region table."""

from __future__ import annotations

from slim_sheet.adapters.columnar import SheetStore, flat_store
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


def big_store(handle: OpenFile, name: str | None) -> SheetStore | None:
    """The columnar cache of the sheet `name` (or of the first sheet) when it is big; a big
    CSV/TSV/JSON is its Parquet cache seen as one sheet named 'data'."""
    if not handle.format.is_workbook:
        return flat_store(handle.path, handle.format)
    store = getattr(handle.reader, "store", None)
    if store is None:
        return None
    sheets = [n for n, _ in handle.reader.sheets()]
    target = sheets[0] if name is None and sheets else name
    exact = next((s for s in sheets if target and s.casefold() == target.casefold()), None)
    return store(exact) if exact else None


def _candidate_sheets(sheets: list[str], name: str | None) -> list[str]:
    if name is None:
        return sheets
    folded = name.casefold()
    exact = [s for s in sheets if s.casefold() == folded]
    prefix = [s for s in sheets if folded.startswith(s.casefold() + ".") and s not in exact]
    return exact + prefix or sheets
