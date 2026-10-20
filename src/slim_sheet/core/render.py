"""Pure renderers: models -> compact text for agents (PLAN §5.4, ADR-10)."""

from __future__ import annotations

import json
from collections.abc import Sequence
from typing import Literal

from slim_sheet.core.a1 import format_range, index_to_col
from slim_sheet.core.budget import cell_text
from slim_sheet.core.models import CellValue, ColumnInfo, Grid, WorkbookInfo

OutputFormat = Literal["markdown", "tsv", "json"]
OUTPUT_FORMATS: tuple[OutputFormat, ...] = ("markdown", "tsv", "json")


def grid_header(left: int, right: int, fmt: OutputFormat) -> str:
    letters = [index_to_col(c) for c in range(left, right + 1)]
    if fmt == "tsv":
        return "\t".join(["#", *letters])
    head = "| # | " + " | ".join(letters) + " |"
    rule = "|--:|" + "|".join("---" for _ in letters) + "|"
    return f"{head}\n{rule}"


def grid_row(row_index: int, values: Sequence[CellValue], fmt: OutputFormat, max_chars: int) -> str:
    cells = [cell_text(v, max_chars) for v in values]
    if fmt == "tsv":
        return "\t".join([str(row_index + 1), *cells])
    escaped = [c.replace("|", "\\|") for c in cells]
    return f"| {row_index + 1} | " + " | ".join(escaped) + " |"


def render_grid(grid: Grid, fmt: OutputFormat, max_chars: int) -> str:
    if fmt == "json":
        return _grid_json(grid, max_chars)
    total = format_range(grid.total) if grid.total else "empty"
    title = f"## {grid.path} / {grid.sheet} — window {format_range(grid.window)} (used {total})"
    lines = [title, grid_header(grid.window.left, grid.window.right, fmt)]
    lines += [grid_row(grid.window.top + i, row, fmt, max_chars) for i, row in enumerate(grid.rows)]
    if grid.next_cursor:
        lines.append(f'next: cursor="{grid.next_cursor}" (call read again with this cursor)')
    return "\n".join(lines)


def _grid_json(grid: Grid, max_chars: int) -> str:
    payload = {
        "sheet": grid.sheet,
        "window": format_range(grid.window),
        "used": format_range(grid.total) if grid.total else None,
        "columns": [index_to_col(c) for c in range(grid.window.left, grid.window.right + 1)],
        "first_row": grid.window.top + 1,
        "rows": [[cell_text(v, max_chars) or None for v in row] for row in grid.rows],
        "next": grid.next_cursor,
    }
    return json.dumps(payload, ensure_ascii=False, separators=(",", ":"))


def render_overview(info: WorkbookInfo) -> str:
    size_kb = max(1, info.size_bytes // 1024)
    lines = [f"## {info.path} ({info.format.value}, {size_kb} KB, {len(info.sheets)} sheet(s))"]
    for sheet in info.sheets:
        used = format_range(sheet.used) if sheet.used else "empty"
        flags = "" if sheet.visible else " [hidden]"
        if sheet.merged_count:
            flags += f" [{sheet.merged_count} merged ranges]"
        dims = f"{sheet.used.n_rows}x{sheet.used.n_cols}" if sheet.used else "0x0"
        lines.append(f"- **{sheet.name}**{flags}: used {used} ({dims})")
        if sheet.header_row is not None:
            lines.append(f"  header row {sheet.header_row + 1}: " + _columns_line(sheet.columns))
        elif sheet.columns:
            lines.append("  no header detected; types: " + _columns_line(sheet.columns))
    return "\n".join(lines)


def _columns_line(columns: Sequence[ColumnInfo]) -> str:
    return ", ".join(f"{c.letter}={c.header or '?'}:{c.dtype}" for c in columns)
