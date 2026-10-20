"""`overview` use-case: a compact map of every sheet (PLAN §5.1)."""

from __future__ import annotations

from slim_sheet.config import Settings
from slim_sheet.core.a1 import index_to_col
from slim_sheet.core.budget import cell_text
from slim_sheet.core.detect import guess_header_row, infer_type
from slim_sheet.core.models import CellRange, CellValue, ColumnInfo, SheetInfo, WorkbookInfo
from slim_sheet.services.workbook import OpenFile, open_file

SAMPLE_ROWS = 60
MAX_COLUMNS_LISTED = 40
HEADER_CHARS = 40


def overview(path: str, settings: Settings) -> WorkbookInfo:
    with open_file(path, settings) as handle:
        sheets = tuple(_describe(handle, name, visible) for name, visible in handle.reader.sheets())
        return WorkbookInfo(
            path=handle.display,
            format=handle.format,
            size_bytes=handle.path.stat().st_size,
            sheets=sheets,
        )


def _describe(handle: OpenFile, name: str, visible: bool) -> SheetInfo:
    used = handle.reader.used_range(name)
    merged = handle.reader.merged_count(name)
    if used is None:
        return SheetInfo(name=name, used=None, visible=visible, merged_count=merged)
    right = min(used.right, used.left + MAX_COLUMNS_LISTED - 1)
    sample_window = CellRange(used.top, used.left, min(used.bottom, used.top + SAMPLE_ROWS), right)
    sample = handle.reader.read(name, sample_window)
    header_idx = guess_header_row(sample)
    columns = _columns(sample, header_idx, used.left)
    header_row = None if header_idx is None else used.top + header_idx
    return SheetInfo(name, used, visible, header_row, columns, merged)


def _columns(
    sample: list[list[CellValue]], header_idx: int | None, left: int
) -> tuple[ColumnInfo, ...]:
    body = sample if header_idx is None else sample[header_idx + 1 :]
    header = sample[header_idx] if header_idx is not None else None
    width = len(sample[0]) if sample else 0
    result = []
    for offset in range(width):
        label = cell_text(header[offset], HEADER_CHARS) if header else None
        dtype = infer_type([row[offset] for row in body])
        result.append(ColumnInfo(index_to_col(left + offset), label or None, dtype))
    return tuple(result)
