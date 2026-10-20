"""Export Arrow tables / raw grids to xlsx, csv, tsv, parquet or json (atomic, journaled).

CSV/TSV text cells starting with = + - @ TAB CR are prefixed with an apostrophe so
spreadsheet apps do not execute them (OWASP CSV injection guidance, PLAN §8).
"""

from __future__ import annotations

import contextlib
import json
import os
import tempfile
from collections.abc import Callable, Sequence
from pathlib import Path

import openpyxl
import pyarrow as pa
import pyarrow.csv as pacsv
import pyarrow.parquet as pq
from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE

from slim_sheet.core.errors import ErrorCode, SlimError
from slim_sheet.core.models import CellValue

EXPORT_SUFFIXES = (".xlsx", ".csv", ".tsv", ".parquet", ".json", ".jsonl")
_INJECTION_PREFIXES = ("=", "+", "-", "@", "\t", "\r")
MAX_SHEET_NAME = 31


def check_suffix(target: Path) -> str:
    suffix = target.suffix.lower()
    if suffix not in EXPORT_SUFFIXES:
        raise SlimError(
            ErrorCode.UNSUPPORTED_FORMAT,
            f"Cannot write '{target.name}'",
            hint=f"Output must end with one of: {' '.join(EXPORT_SUFFIXES)}",
        )
    return suffix


def export_table(table: pa.Table, target: Path, sheet: str = "data") -> None:
    suffix = check_suffix(target)
    if suffix == ".xlsx":
        _atomic(target, lambda tmp: _xlsx(tmp, [(sheet, _table_rows(table))]))
    elif suffix in (".csv", ".tsv"):
        delimiter = "\t" if suffix == ".tsv" else ","
        safe = _defuse(table)
        options = pacsv.WriteOptions(delimiter=delimiter)
        _atomic(target, lambda tmp: pacsv.write_csv(safe, tmp, write_options=options))
    elif suffix == ".parquet":
        _atomic(target, lambda tmp: pq.write_table(table, tmp))
    else:
        _atomic(target, lambda tmp: _ndjson(table, tmp))


def export_grids(sheets: Sequence[tuple[str, Sequence[Sequence[CellValue]]]], target: Path) -> None:
    """Raw cell grids -> one xlsx with a sheet per grid (values only)."""
    if check_suffix(target) != ".xlsx":
        raise SlimError(ErrorCode.BAD_ARGUMENT, "Whole-workbook export needs an .xlsx out_path")
    _atomic(target, lambda tmp: _xlsx(tmp, sheets))


def _table_rows(table: pa.Table) -> list[list[CellValue]]:
    header: list[CellValue] = list(table.column_names)
    return [header, *([row[c] for c in table.column_names] for row in table.to_pylist())]


def _xlsx(path: Path, sheets: Sequence[tuple[str, Sequence[Sequence[CellValue]]]]) -> None:
    wb = openpyxl.Workbook(write_only=True)
    for name, rows in sheets:
        ws = wb.create_sheet(_sheet_title(name))
        for row in rows:
            ws.append([_xlsx_safe(v) for v in row])
    wb.save(path)


def _xlsx_safe(value: CellValue) -> CellValue:
    """xlsx cannot store ASCII control characters; drop them instead of failing the export."""
    return ILLEGAL_CHARACTERS_RE.sub("", value) if isinstance(value, str) else value


def _sheet_title(name: str) -> str:
    cleaned = "".join("_" if ch in "[]:*?/\\" else ch for ch in name)
    return cleaned[:MAX_SHEET_NAME] or "Sheet1"


def _defuse(table: pa.Table) -> pa.Table:
    columns = []
    for column in table.columns:
        if pa.types.is_string(column.type) or pa.types.is_large_string(column.type):
            values = [
                f"'{v}" if isinstance(v, str) and v.startswith(_INJECTION_PREFIXES) else v
                for v in column.to_pylist()
            ]
            columns.append(pa.array(values, column.type))
        else:
            columns.append(column)
    return pa.Table.from_arrays(columns, names=table.column_names)


def _ndjson(table: pa.Table, path: Path) -> None:
    with path.open("w", encoding="utf-8") as handle:
        for row in table.to_pylist():
            handle.write(json.dumps(row, ensure_ascii=False, default=str) + "\n")


def _atomic(target: Path, write: Callable[[Path], object]) -> None:
    fd, tmp_name = tempfile.mkstemp(prefix=".slim-", suffix=target.suffix, dir=target.parent)
    os.close(fd)
    tmp = Path(tmp_name)
    try:
        write(tmp)
        tmp.replace(target)
    except PermissionError as exc:
        raise SlimError(
            ErrorCode.FILE_LOCKED,
            f"'{target.name}' is locked by another program",
            hint="Close it and retry, or choose another out_path.",
        ) from exc
    finally:
        with contextlib.suppress(FileNotFoundError):
            tmp.unlink()
