"""`overview` use-case: a compact map of every sheet (PLAN §5.1)."""

from __future__ import annotations

import os
from pathlib import Path

from slim_sheet.adapters.hidden_info import hidden_content
from slim_sheet.adapters.sandbox import resolve_folder
from slim_sheet.adapters.writers.xlsx_writer import package_warnings
from slim_sheet.config import Settings
from slim_sheet.core.a1 import index_to_col
from slim_sheet.core.budget import cell_text
from slim_sheet.core.detect import guess_header_row, infer_type
from slim_sheet.core.models import (
    CellRange,
    CellValue,
    ColumnInfo,
    FolderListing,
    SheetInfo,
    WorkbookInfo,
)
from slim_sheet.services.macros import macro_warnings
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
            warnings=(*macro_warnings(handle.path), *package_warnings(handle.path)),
            hidden=tuple(hidden_content(handle.path)),
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


SPREADSHEET_SUFFIXES = (
    ".xlsx", ".xlsm", ".xltx", ".xltm", ".xlsb", ".xls", ".ods", ".numbers",
    ".csv", ".tsv", ".txt", ".parquet", ".json", ".jsonl", ".ndjson",
)  # fmt: skip
MAX_LISTED = 200
MAX_ROOT_LISTED = 40
# a whole-roots listing shows workbooks and tabular files only (not every .txt/.json on disk)
ROOT_LISTING_SUFFIXES = (
    ".xlsx", ".xlsm", ".xltx", ".xltm", ".xlsb", ".xls", ".ods", ".numbers",
    ".csv", ".tsv", ".parquet",
)  # fmt: skip
SKIPPED_DIRS = frozenset(
    {"node_modules", "site-packages", "venv", "env", "__pycache__", "dist-info"}
)


def list_workbooks(path: str, settings: Settings) -> FolderListing | None:
    """A folder inside the roots -> its spreadsheets (recursive); a file -> None.

    An empty path lists every allowed root, with full paths the agent can pass back as is.
    """
    if not path or not path.strip():
        return _list_roots(settings)
    folder = resolve_folder(path, settings)
    if folder is None:
        return None
    found = _spreadsheets(folder)
    items = tuple(
        (item.relative_to(folder).as_posix(), item.stat().st_size) for item in found[:MAX_LISTED]
    )
    return FolderListing(folder.name or str(folder), items, len(found))


def _spreadsheets(folder: Path, suffixes: tuple[str, ...] = SPREADSHEET_SUFFIXES) -> list[Path]:
    """Data files under `folder`, skipping tool and environment folders (.venv, node_modules...)."""
    found: list[Path] = []
    for current, dirs, files in os.walk(folder):
        dirs[:] = sorted(d for d in dirs if not d.startswith(".") and d.lower() not in SKIPPED_DIRS)
        found += [
            Path(current) / name
            for name in sorted(files)
            if Path(name).suffix.lower() in suffixes and not name.startswith(("~$", ".slim-"))
        ]
    return found


def _list_roots(settings: Settings) -> FolderListing:
    """Every allowed root at once: real spreadsheets only, most recently changed first."""
    found = [
        item for root in settings.roots if root.is_dir()
        for item in _spreadsheets(root, ROOT_LISTING_SUFFIXES)
    ]  # fmt: skip
    found.sort(key=lambda item: item.stat().st_mtime, reverse=True)
    items = tuple((str(item), item.stat().st_size) for item in found[:MAX_ROOT_LISTED])
    folders = "; ".join(map(str, settings.roots))
    return FolderListing(f"the allowed folders ({folders}), newest first", items, len(found))
