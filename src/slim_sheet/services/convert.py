"""`convert` use-case: any readable format -> xlsx / csv / tsv / parquet / json.

- workbook -> .xlsx without `sheet`: every sheet copied cell-for-cell (values only)
- workbook -> .xlsx with `sheet`:   that sheet's cells only
- anything -> csv/tsv/parquet/json: one table (detected header), chosen by `sheet`
"""

from __future__ import annotations

import tempfile
from dataclasses import dataclass
from pathlib import Path

import duckdb

from slim_sheet.adapters.engine.duck import load_tables
from slim_sheet.adapters.formats import sniff
from slim_sheet.adapters.readers import SheetReader
from slim_sheet.adapters.sandbox import resolve_path
from slim_sheet.adapters.writers.export import export_grids, export_table
from slim_sheet.config import Settings
from slim_sheet.core.errors import ErrorCode, SlimError
from slim_sheet.core.models import CellRange, CellValue, FileFormat, OutputResult
from slim_sheet.services.output import journal_before_write, output_path
from slim_sheet.services.workbook import OpenFile, open_file, pick_sheet

VALUES_ONLY = "values only: formulas, styles, merges and charts are not copied"


@dataclass(frozen=True, slots=True)
class CsvOptions:
    """How the CSV involved is written: read options when the source is CSV/TSV, write options
    (delimiter, encoding='utf-8-sig' for a BOM) when the target is."""

    delimiter: str | None = None
    encoding: str | None = None
    decimal: str | None = None

    @property
    def given(self) -> bool:
        return any((self.delimiter, self.encoding, self.decimal))


def convert(
    path: str,
    out_path: str,
    settings: Settings,
    sheet: str | None = None,
    overwrite: bool = False,
    csv: CsvOptions | None = None,
) -> OutputResult:
    options = csv or CsvOptions()
    source = resolve_path(path, settings)
    target = output_path(out_path, settings, overwrite, base=source.parent)
    if source == target:
        raise SlimError(ErrorCode.WRITE_CONFLICT, "out_path must differ from path for convert")
    if options.given and sniff(source) in (FileFormat.CSV, FileFormat.TSV):
        # the default reader auto-detects everything, so explicit CSV options parse it here
        return _csv_with_options(source, target, options)
    with open_file(path, settings) as handle:
        if target.suffix.lower() == ".xlsx" and handle.format.is_workbook:
            return _workbook_to_xlsx(handle, target, sheet)
        return _table_export(handle, target, sheet, options)


def _csv_with_options(source: Path, target: Path, options: CsvOptions) -> OutputResult:
    """Parse a CSV with an explicit delimiter / encoding / decimal separator."""
    if options.decimal not in (None, ".", ","):
        raise SlimError(
            ErrorCode.BAD_ARGUMENT, "decimal must be '.' or ','", hint="e.g. decimal=','"
        )
    encoding = (options.encoding or "utf-8").strip()
    try:
        text = source.read_bytes().decode(encoding)
    except (LookupError, UnicodeDecodeError) as exc:
        raise SlimError(
            ErrorCode.READ_FAILED,
            f"'{source.name}' is not valid {encoding}: {exc}",
            hint="Try encoding='utf-8', 'cp1252', 'latin-1', 'cp1258' (Vietnamese) or 'utf-16'.",
        ) from exc
    args = ["header=true", "auto_detect=true"]
    if options.delimiter:
        args.append("delim=?")
    if options.decimal:
        args.append("decimal_separator=?")
    params = [p for p in (options.delimiter, options.decimal) if p]
    with tempfile.TemporaryDirectory(prefix="slim-csv-") as folder:
        staged = Path(folder) / "data.csv"
        staged.write_text(text.lstrip("\ufeff"), encoding="utf-8")
        con = duckdb.connect(":memory:")
        try:
            table = con.execute(
                f"SELECT * FROM read_csv(?, {', '.join(args)})", [str(staged), *params]
            ).to_arrow_table()
        except duckdb.Error as exc:
            raise SlimError(
                ErrorCode.READ_FAILED, f"Failed to parse '{source.name}': {exc}"
            ) from exc
        finally:
            con.close()
    backup = journal_before_write(target, f"convert {source.name}")
    export_table(table, target, sheet="data")
    given = (("delimiter", options.delimiter), ("encoding", encoding), ("decimal", options.decimal))
    used = ", ".join(f"{key}={value!r}" for key, value in given if value)
    summary = f"table data ({used}): {table.num_rows} rows, {table.num_columns} columns"
    return OutputResult(source.name, target.name, summary, backup)


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


def _table_export(
    handle: OpenFile, target: Path, sheet: str | None, options: CsvOptions | None = None
) -> OutputResult:
    tables, _ = load_tables(handle.path, handle.format, handle.reader, only=sheet)
    chosen = next(
        (t for t in tables if sheet is None or t.name.casefold() == sheet.casefold()), None
    )
    if chosen is None:
        names = ", ".join(t.name for t in tables) or "none"
        raise SlimError(ErrorCode.SHEET_NOT_FOUND, f"No table '{sheet}'", hint=f"Tables: {names}")
    backup = journal_before_write(target, f"convert {handle.display}")
    options = options or CsvOptions()
    bom = (options.encoding or "").lower().replace("_", "-") in ("utf-8-sig", "utf8-sig")
    if options.encoding and not bom and options.encoding.lower() not in ("utf-8", "utf8"):
        raise SlimError(
            ErrorCode.BAD_ARGUMENT,
            "CSV output is UTF-8; encoding='utf-8-sig' adds the BOM Excel needs for accents",
            hint="Use encoding='utf-8-sig' (or leave it out)",
        )
    export_table(chosen.table, target, sheet=chosen.name, delimiter=options.delimiter, bom=bom)
    summary = f"table {chosen.name} ({chosen.source}): {chosen.table.num_rows} rows"
    return OutputResult(handle.display, target.name, summary, backup)
