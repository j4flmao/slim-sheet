"""Sheet -> region tables (header + body with sheet coordinates). Shared by query, clean,
anomalies and diff so every tool sees the same tables."""

from __future__ import annotations

from dataclasses import dataclass

from slim_sheet.adapters.readers import SheetReader
from slim_sheet.core.a1 import format_range, index_to_col
from slim_sheet.core.detect import guess_header_row
from slim_sheet.core.models import CellRange, CellValue
from slim_sheet.core.regions import column_names, detect_regions


@dataclass(frozen=True, slots=True)
class RegionTable:
    name: str  # "Orders" for the largest region, "Orders.2", ... for the rest
    sheet: str
    region: CellRange
    columns: tuple[str, ...]
    body: tuple[tuple[CellValue, ...], ...]
    body_top: int  # sheet row index (0-based) of body[0]
    has_header: bool

    @property
    def source(self) -> str:
        return f"{self.sheet}!{format_range(self.region)}"

    def address(self, body_row: int, col: int) -> str:
        """'Sheet!C7' for body[body_row][col]."""
        return f"{self.sheet}!{index_to_col(self.region.left + col)}{self.body_top + body_row + 1}"


def extract_tables(reader: SheetReader, sheet: str) -> list[RegionTable]:
    used = reader.used_range(sheet)
    if used is None:
        return []
    rows = reader.read(sheet, used)
    tables = []
    for index, region in enumerate(detect_regions(rows, used.top, used.left)):
        block = [
            r[region.left - used.left : region.right - used.left + 1]
            for r in rows[region.top - used.top : region.bottom - used.top + 1]
        ]
        name = sheet if index == 0 else f"{sheet}.{index + 1}"
        tables.append(_table(name, sheet, region, block))
    return tables


def _table(name: str, sheet: str, region: CellRange, block: list[list[CellValue]]) -> RegionTable:
    header_idx = guess_header_row(block)
    letters = [index_to_col(c) for c in range(region.left, region.right + 1)]
    header = block[header_idx] if header_idx is not None else None
    start = 0 if header_idx is None else header_idx + 1
    return RegionTable(
        name=name,
        sheet=sheet,
        region=region,
        columns=tuple(column_names(header, letters)),
        body=tuple(tuple(r) for r in block[start:]),
        body_top=region.top + start,
        has_header=header_idx is not None,
    )


def find_table(tables: list[RegionTable], name: str | None) -> RegionTable | None:
    if not tables:
        return None
    if name is None:
        return tables[0]
    return next((t for t in tables if t.name.casefold() == name.casefold()), None)
