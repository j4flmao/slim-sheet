"""`convert` use-case: any readable format -> xlsx / csv / tsv / parquet / json.

- workbook -> .xlsx without `sheet`: every sheet copied cell-for-cell (values only)
- workbook -> .xlsx with `sheet`:   that sheet's cells only
- anything -> csv/tsv/parquet/json: one table (detected header), chosen by `sheet`
"""

from __future__ import annotations

from pathlib import Path

from slim_sheet.adapters.engine.duck import load_tables
from slim_sheet.adapters.readers import SheetReader
from slim_sheet.adapters.writers.export import export_grids, export_table
from slim_sheet.config import Settings
from slim_sheet.core.errors import ErrorCode, SlimError
from slim_sheet.core.models import CellRange, CellValue, OutputResult
from slim_sheet.services.output import journal_before_write, output_path
from slim_sheet.services.workbook import OpenFile, open_file, pick_sheet

VALUES_ONLY = "values only: formulas, styles, merges and charts are not copied"


def convert(
    path: str, out_path: str, settings: Settings, sheet: str | None = None, overwrite: bool = False
) -> OutputResult:
    target = output_path(out_path, settings, overwrite)
    with open_file(path, settings) as handle:
        if handle.path == target:
            raise SlimError(ErrorCode.WRITE_CONFLICT, "out_path must differ from path for convert")
        if target.suffix.lower() == ".xlsx" and handle.format.is_workbook:
            return _workbook_to_xlsx(handle, target, sheet)
        return _table_export(handle, target, sheet)


def _workbook_to_xlsx(handle: OpenFile, target: Path, sheet: str | None) -> OutputResult:
    names = [pick_sheet(handle, sheet)] if sheet else [n for n, _ in handle.reader.sheets()]
    grids = [(name, _grid_from_a1(handle.reader, name)) for name in names]
    backup = journal_before_write(target, f"convert {handle.display}")
    export_grids(grids, target)
    cells = sum(len(r) * (len(r[0]) if r else 0) for _, r in grids)
    summary = f"{len(grids)} sheet(s), {cells} cells"
    return OutputResult(handle.display, target.name, summary, backup, (VALUES_ONLY,))


def _grid_from_a1(reader: SheetReader, sheet: str) -> list[list[CellValue]]:
    used = reader.used_range(sheet)
    if used is None:
        return []
    return reader.read(sheet, CellRange(0, 0, used.bottom, used.right))


def _table_export(handle: OpenFile, target: Path, sheet: str | None) -> OutputResult:
    tables, _ = load_tables(handle.path, handle.format, handle.reader, only=sheet)
    chosen = next(
        (t for t in tables if sheet is None or t.name.casefold() == sheet.casefold()), None
    )
    if chosen is None:
        names = ", ".join(t.name for t in tables) or "none"
        raise SlimError(ErrorCode.SHEET_NOT_FOUND, f"No table '{sheet}'", hint=f"Tables: {names}")
    backup = journal_before_write(target, f"convert {handle.display}")
    export_table(chosen.table, target, sheet=chosen.name)
    summary = f"table {chosen.name} ({chosen.source}): {chosen.table.num_rows} rows"
    return OutputResult(handle.display, target.name, summary, backup)
