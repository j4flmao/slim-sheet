"""A1 notation <-> zero-based coordinates."""

from __future__ import annotations

import re

from slim_sheet.core.errors import ErrorCode, SlimError
from slim_sheet.core.models import CellRange

_CELL = re.compile(r"^\$?([A-Za-z]{1,3})\$?(\d{1,7})$")
_MAX_COL = 16_383  # XFD
_MAX_ROW = 1_048_575


def col_to_index(letters: str) -> int:
    index = 0
    for char in letters.upper():
        if not "A" <= char <= "Z":
            raise SlimError(ErrorCode.BAD_RANGE, f"Invalid column letters '{letters}'")
        index = index * 26 + (ord(char) - 64)
    return index - 1


def index_to_col(index: int) -> str:
    if index < 0:
        raise SlimError(ErrorCode.BAD_RANGE, f"Negative column index {index}")
    letters = ""
    index += 1
    while index:
        index, rem = divmod(index - 1, 26)
        letters = chr(65 + rem) + letters
    return letters


def parse_cell(ref: str) -> tuple[int, int]:
    match = _CELL.match(ref.strip())
    if not match:
        raise SlimError(
            ErrorCode.BAD_RANGE,
            f"Invalid cell reference '{ref}'",
            hint="Use A1 notation, e.g. range='B2:F50'",
        )
    col, row = col_to_index(match.group(1)), int(match.group(2)) - 1
    if row < 0 or row > _MAX_ROW or col > _MAX_COL:
        raise SlimError(ErrorCode.BAD_RANGE, f"Cell '{ref}' is outside the sheet grid")
    return row, col


def parse_range(ref: str) -> CellRange:
    """Parse 'B2' or 'B2:F50' (either corner order). Sheet prefixes are rejected."""
    if "!" in ref:
        raise SlimError(
            ErrorCode.BAD_RANGE,
            f"Range '{ref}' contains a sheet name",
            hint="Pass the sheet separately: sheet='Orders', range='A1:C10'",
        )
    parts = ref.split(":")
    if len(parts) > 2 or not parts[0]:
        raise SlimError(ErrorCode.BAD_RANGE, f"Invalid range '{ref}'", hint="e.g. range='A1:D20'")
    r1, c1 = parse_cell(parts[0])
    r2, c2 = parse_cell(parts[-1])
    return CellRange(min(r1, r2), min(c1, c2), max(r1, r2), max(c1, c2))


def format_range(rng: CellRange) -> str:
    start = f"{index_to_col(rng.left)}{rng.top + 1}"
    end = f"{index_to_col(rng.right)}{rng.bottom + 1}"
    return start if start == end else f"{start}:{end}"
