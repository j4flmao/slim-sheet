"""`find` use-case: locate text across sheets, returning A1 addresses."""

from __future__ import annotations

import re
from collections.abc import Callable

from slim_sheet.config import Settings
from slim_sheet.core.budget import cell_text
from slim_sheet.core.errors import ErrorCode, SlimError
from slim_sheet.core.models import CellRange, CellValue, FindHit, FindResult
from slim_sheet.services.workbook import OpenFile, open_file, pick_sheet

CHUNK_ROWS = 1000
MAX_RESULTS_CAP = 200
MAX_PATTERN_CHARS = 200
_NO_TRUNCATE = 1_000_000


def find(
    path: str,
    text: str,
    settings: Settings,
    sheet: str | None = None,
    regex: bool = False,
    max_results: int = 30,
) -> FindResult:
    match = _matcher(text, regex)
    limit = max(1, min(max_results, MAX_RESULTS_CAP))
    hits: list[FindHit] = []
    scanned = 0
    with open_file(path, settings) as handle:
        sheets = [pick_sheet(handle, sheet)] if sheet else [n for n, _ in handle.reader.sheets()]
        for name in sheets:
            scanned += _scan_sheet(handle, name, match, hits, limit + 1)
            if len(hits) > limit:
                break
        return FindResult(handle.display, text, tuple(hits[:limit]), len(hits) > limit, scanned)


def _matcher(text: str, regex: bool) -> Callable[[str], bool]:
    if not text:
        raise SlimError(
            ErrorCode.BAD_ARGUMENT, "text is empty", hint="e.g. find(path, text='invoice')"
        )
    if not regex:
        needle = text.casefold()
        return lambda value: needle in value.casefold()
    if len(text) > MAX_PATTERN_CHARS:
        raise SlimError(ErrorCode.BAD_ARGUMENT, f"regex longer than {MAX_PATTERN_CHARS} characters")
    try:
        pattern = re.compile(text, re.IGNORECASE)
    except re.error as exc:
        raise SlimError(
            ErrorCode.BAD_ARGUMENT,
            f"Invalid regex: {exc}",
            hint="Set regex=false to search plain text.",
        ) from exc
    return lambda value: pattern.search(value) is not None


def _scan_sheet(
    handle: OpenFile, sheet: str, match: Callable[[str], bool], hits: list[FindHit], stop: int
) -> int:
    used = handle.reader.used_range(sheet)
    if used is None:
        return 0
    scanned = 0
    for top in range(used.top, used.bottom + 1, CHUNK_ROWS):
        window = CellRange(top, used.left, min(used.bottom, top + CHUNK_ROWS - 1), used.right)
        for r, row in enumerate(handle.reader.read(sheet, window)):
            scanned += len(row)
            for c, value in enumerate(row):
                if _is_hit(value, match):
                    hits.append(FindHit(sheet, window.top + r, window.left + c, value))
                    if len(hits) >= stop:
                        return scanned
    return scanned


def _is_hit(value: CellValue, match: Callable[[str], bool]) -> bool:
    if value is None or value == "":
        return False
    return match(value if isinstance(value, str) else cell_text(value, _NO_TRUNCATE))
