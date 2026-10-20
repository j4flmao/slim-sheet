"""Immutable value objects shared by services and interfaces."""

from __future__ import annotations

import datetime as dt
from dataclasses import dataclass
from enum import StrEnum

CellValue = str | int | float | bool | dt.datetime | dt.date | dt.time | dt.timedelta | None


class FileFormat(StrEnum):
    XLSX = "xlsx"
    XLSB = "xlsb"
    XLS = "xls"
    ODS = "ods"
    CSV = "csv"
    TSV = "tsv"
    PARQUET = "parquet"
    JSON = "json"
    NUMBERS = "numbers"

    @property
    def is_workbook(self) -> bool:
        return self in {
            FileFormat.XLSX,
            FileFormat.XLSB,
            FileFormat.XLS,
            FileFormat.ODS,
            FileFormat.NUMBERS,
        }


@dataclass(frozen=True, slots=True)
class CellRange:
    """Zero-based, inclusive rectangle."""

    top: int
    left: int
    bottom: int
    right: int

    @property
    def n_rows(self) -> int:
        return self.bottom - self.top + 1

    @property
    def n_cols(self) -> int:
        return self.right - self.left + 1

    def intersect(self, other: CellRange) -> CellRange | None:
        top, left = max(self.top, other.top), max(self.left, other.left)
        bottom, right = min(self.bottom, other.bottom), min(self.right, other.right)
        if top > bottom or left > right:
            return None
        return CellRange(top, left, bottom, right)


@dataclass(frozen=True, slots=True)
class ColumnInfo:
    letter: str
    header: str | None
    dtype: str


@dataclass(frozen=True, slots=True)
class SheetInfo:
    name: str
    used: CellRange | None
    visible: bool = True
    header_row: int | None = None
    columns: tuple[ColumnInfo, ...] = ()
    merged_count: int = 0


@dataclass(frozen=True, slots=True)
class WorkbookInfo:
    path: str
    format: FileFormat
    size_bytes: int
    sheets: tuple[SheetInfo, ...]


@dataclass(frozen=True, slots=True)
class Grid:
    """A window of cells plus the cursor to continue from, if truncated."""

    path: str
    sheet: str
    window: CellRange
    rows: tuple[tuple[CellValue, ...], ...]
    total: CellRange | None
    next_cursor: str | None = None


@dataclass(frozen=True, slots=True)
class TableSchema:
    """A SQL-queryable table: one detected region of a sheet, or a whole flat file."""

    name: str
    source: str
    columns: tuple[tuple[str, str], ...]
    rows: int


@dataclass(frozen=True, slots=True)
class QueryResult:
    sql: str
    columns: tuple[str, ...]
    rows: tuple[tuple[CellValue, ...], ...]
    total_rows: int
    offset: int
    next_offset: int | None


@dataclass(frozen=True, slots=True)
class FindHit:
    sheet: str
    row: int
    col: int
    value: CellValue


@dataclass(frozen=True, slots=True)
class FindResult:
    path: str
    query: str
    hits: tuple[FindHit, ...]
    truncated: bool
    cells_scanned: int


@dataclass(frozen=True, slots=True)
class ColumnProfile:
    name: str
    dtype: str
    count: int
    null_pct: float
    distinct: int
    min: str | None
    max: str | None
    mean: str | None
    median: str | None
    top: tuple[tuple[str, int], ...]


@dataclass(frozen=True, slots=True)
class ProfileResult:
    path: str
    table: str
    source: str
    rows: int
    columns: tuple[ColumnProfile, ...]
    omitted_columns: int


@dataclass(frozen=True, slots=True)
class WriteResult:
    source: str
    target: str
    sheet: str
    written: CellRange
    cells: int
    created_sheet: bool
    backup: str | None
    warnings: tuple[str, ...]


@dataclass(frozen=True, slots=True)
class OutputResult:
    """A file produced by convert / clean / merge."""

    source: str
    target: str
    summary: str
    backup: str | None
    warnings: tuple[str, ...] = ()


@dataclass(frozen=True, slots=True)
class CleanResult:
    path: str
    table: str
    source: str
    rows_before: int
    rows_after: int
    counts: tuple[tuple[str, int], ...]
    samples: tuple[tuple[str, CellValue, CellValue, str], ...]  # (A1, before, after, op)
    notes: tuple[str, ...]
    output: OutputResult | None


@dataclass(frozen=True, slots=True)
class Finding:
    address: str
    value: CellValue
    kind: str
    detail: str


@dataclass(frozen=True, slots=True)
class AnomalyResult:
    path: str
    table: str
    source: str
    rows: int
    findings: tuple[Finding, ...]
    counts: tuple[tuple[str, int], ...]


@dataclass(frozen=True, slots=True)
class MergeResult:
    left: str
    right: str
    how: str
    keys: tuple[str, ...]
    left_rows: int
    right_rows: int
    left_unmatched: int
    right_unmatched: int
    preview: QueryResult
    output: OutputResult | None
    warnings: tuple[str, ...]


@dataclass(frozen=True, slots=True)
class DiffResult:
    left: str
    right: str
    mode: str  # "cells" | "key:<column>"
    summary: tuple[str, ...]
    changes: tuple[tuple[str, CellValue, CellValue], ...]  # (location, before, after)
    total_changes: int


@dataclass(frozen=True, slots=True)
class EvalResult:
    formula: str
    address: str | None
    value: CellValue | tuple[tuple[CellValue, ...], ...]
    cached: CellValue  # Excel's last calculated value for a cell, when known
    inputs: tuple[tuple[str, str], ...]  # (reference, short description of what was read)
    recomputed: tuple[str, ...]  # precedent cells evaluated because they had no cached value
    notes: tuple[str, ...]


@dataclass(frozen=True, slots=True)
class CellDescription:
    address: str
    value: CellValue
    facts: tuple[tuple[str, str], ...]


@dataclass(frozen=True, slots=True)
class PivotResult:
    source: str
    sql: str
    preview: QueryResult
    output: OutputResult | None
