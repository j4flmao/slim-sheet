"""Cell-count guard for xlsx-family packages (PLAN §11.10).

A few-KB zip can declare millions of cells and make any full load (openpyxl, calamine) use
gigabytes of memory. Before a workbook is opened, the sheet XML is streamed and its cell
elements are counted, stopping as soon as the limit is passed. Results are cached per file
version, so the check costs one pass per change.
"""

from __future__ import annotations

import re
import zipfile
from pathlib import Path

from slim_sheet.core.errors import ErrorCode, SlimError

CHUNK = 4 * 1024 * 1024
_SHEET_PART = re.compile(r"^xl/worksheets/[^/]+\.(xml|bin)$")
_CELL = re.compile(rb"<c[ >/]")
_cache: dict[tuple[str, int, int], int] = {}


def count_cells(path: Path, stop_after: int) -> int:
    """Cell elements in all worksheet parts (stops counting once `stop_after` is exceeded)."""
    stat = path.stat()
    key = (str(path), stat.st_mtime_ns, stat.st_size)
    if key in _cache:
        return _cache[key]
    total = 0
    with zipfile.ZipFile(path) as archive:
        for info in archive.infolist():
            if not _SHEET_PART.match(info.filename) or info.filename.endswith(".bin"):
                continue
            total += _count_part(archive, info.filename, stop_after - total)
            if total > stop_after:
                break
    _cache[key] = total
    return total


def _count_part(archive: zipfile.ZipFile, name: str, budget: int) -> int:
    found, tail = 0, b""
    with archive.open(name) as stream:
        while chunk := stream.read(CHUNK):
            data = tail + chunk
            found += len(_CELL.findall(data[: len(data) - 3])) if len(data) > 3 else 0
            tail = data[-3:] if len(data) > 3 else data
            if found > budget:
                return found
    return found + len(_CELL.findall(tail))


def check_cells(path: Path, limit: int) -> None:
    """Refuse a package with more than `limit` cells (0 disables the check)."""
    if limit <= 0 or not zipfile.is_zipfile(path):
        return
    try:
        cells = count_cells(path, limit)
    except (zipfile.BadZipFile, OSError, KeyError):
        return  # not a readable package: the reader reports the real problem
    if cells > limit:
        raise SlimError(
            ErrorCode.TOO_LARGE,
            f"'{path.name}' has more than {limit:,} cells",
            hint="Raise SLIM_SHEET_MAX_CELLS if this file is trusted, or export the sheet you "
            "need to CSV/parquet in Excel first.",
        )
