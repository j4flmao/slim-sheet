"""`read` use-case: a budgeted window of cells with a continuation cursor (PLAN §7)."""

from __future__ import annotations

from dataclasses import dataclass

from slim_sheet.config import Settings
from slim_sheet.core.a1 import format_range, parse_range
from slim_sheet.core.budget import HEADROOM, clamp_budget, estimate_tokens
from slim_sheet.core.errors import ErrorCode, SlimError
from slim_sheet.core.models import CellRange, CellValue, Grid
from slim_sheet.core.render import OUTPUT_FORMATS, OutputFormat, grid_header, grid_row
from slim_sheet.services.workbook import OpenFile, open_file, pick_sheet

CHUNK_ROWS = 200
TITLE_RESERVE_TOKENS = 60


@dataclass(frozen=True, slots=True)
class ReadRequest:
    path: str
    sheet: str | None = None
    range: str | None = None
    cursor: str | None = None
    max_tokens: int | None = None
    format: str = "markdown"


def read(request: ReadRequest, settings: Settings) -> Grid:
    fmt = parse_format(request.format)
    budget = clamp_budget(request.max_tokens, settings.max_tokens)
    with open_file(request.path, settings) as handle:
        sheet, window = _resolve_window(handle, request)
        total = handle.reader.used_range(sheet)
        if window is None:
            empty = CellRange(0, 0, 0, 0)
            return Grid(handle.display, sheet, empty, (), total)
        rows = _collect_rows(handle, sheet, window, fmt, budget, settings.max_cell_chars)
        shown = CellRange(window.top, window.left, window.top + max(len(rows), 1) - 1, window.right)
        next_cursor = None
        if window.top + len(rows) <= window.bottom:
            remaining = CellRange(window.top + len(rows), window.left, window.bottom, window.right)
            next_cursor = f"{sheet}!{format_range(remaining)}"
        frozen = tuple(tuple(row) for row in rows)
        return Grid(handle.display, sheet, shown, frozen, total, next_cursor)


def parse_format(raw: str) -> OutputFormat:
    for fmt in OUTPUT_FORMATS:
        if raw == fmt:
            return fmt
    raise SlimError(
        ErrorCode.BAD_ARGUMENT,
        f"Unknown format '{raw}'",
        hint=f"Use one of: {', '.join(OUTPUT_FORMATS)}",
    )


def _resolve_window(handle: OpenFile, request: ReadRequest) -> tuple[str, CellRange | None]:
    """Cursor wins over sheet/range. Cursors are plain 'Sheet!A120:F500' strings."""
    if request.cursor:
        sheet_part, sep, range_part = request.cursor.rpartition("!")
        if not sep or not range_part:
            raise SlimError(
                ErrorCode.BAD_CURSOR,
                f"Malformed cursor '{request.cursor}'",
                hint="Pass the exact cursor string returned by the previous read.",
            )
        return pick_sheet(handle, sheet_part), parse_range(range_part)
    sheet = pick_sheet(handle, request.sheet)
    used = handle.reader.used_range(sheet)
    if not request.range:
        return sheet, used
    asked = parse_range(request.range)
    if used is None:
        return sheet, None
    return sheet, asked.intersect(CellRange(0, 0, used.bottom, used.right))


def _collect_rows(
    handle: OpenFile,
    sheet: str,
    window: CellRange,
    fmt: OutputFormat,
    budget: int,
    max_chars: int,
) -> list[list[CellValue]]:
    limit = int(budget * HEADROOM) - TITLE_RESERVE_TOKENS
    spent = estimate_tokens(grid_header(window.left, window.right, "markdown"))
    rows: list[list[CellValue]] = []
    top = window.top
    while top <= window.bottom:
        chunk = CellRange(top, window.left, min(window.bottom, top + CHUNK_ROWS - 1), window.right)
        for offset, values in enumerate(handle.reader.read(sheet, chunk)):
            cost = estimate_tokens(grid_row(top + offset, values, fmt, max_chars)) + 1
            if rows and spent + cost > limit:
                return rows
            spent += cost
            rows.append(values)
        top = chunk.bottom + 1
    return rows
