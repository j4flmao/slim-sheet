"""Classify formula reference operands: cell/range, whole column/row, table, defined name."""

from __future__ import annotations

import re
from dataclasses import dataclass

from slim_sheet.core.a1 import col_to_index

_CELL = r"\$?[A-Za-z]{1,3}\$?\d+"
_CELL_RANGE = re.compile(rf"^{_CELL}(:{_CELL})?$")
_COLS = re.compile(r"^\$?([A-Za-z]{1,3}):\$?([A-Za-z]{1,3})$")
_ROWS = re.compile(r"^\$?(\d+):\$?(\d+)$")
_TABLE = re.compile(r"^[A-Za-z_\\][\w.]*\[.*\]$")
_DIGITS = re.compile(r"\d+")
_LETTERS = re.compile(r"[A-Za-z]+")

XLS_MAX_ROW = 65_536
XLS_MAX_COL = 255  # IV
MAX_ROW = 1_048_576
MAX_COL = 16_383  # XFD


@dataclass(frozen=True, slots=True)
class Ref:
    raw: str
    sheet: str | None
    address: str
    kind: str  # "cells" | "columns" | "rows" | "table" | "name" | "external"


def split_sheet(raw: str) -> tuple[str | None, str]:
    """ "'My Sheet'!A1" -> ("My Sheet", "A1"); "A1" -> (None, "A1")."""
    if "!" not in raw:
        return None, raw
    sheet, _, address = raw.rpartition("!")
    if sheet.startswith("'") and sheet.endswith("'"):
        sheet = sheet[1:-1].replace("''", "'")
    return sheet, address


def classify(raw: str) -> Ref:
    sheet, address = split_sheet(raw)
    if sheet is not None and sheet.startswith("["):
        return Ref(raw, sheet, address, "external")
    if _CELL_RANGE.match(address):
        kind = "cells"
    elif _COLS.match(address):
        kind = "columns"
    elif _ROWS.match(address):
        kind = "rows"
    elif _TABLE.match(address):
        kind = "table"
    else:
        kind = "name"
    return Ref(raw, sheet, address, kind)


def max_extent(ref: Ref) -> tuple[int, int]:
    """Largest (row, col) touched, 1-based row and 0-based col; (0, 0) if not applicable."""
    if ref.kind == "cells":
        rows = [int(d) for d in _DIGITS.findall(ref.address)]
        cols = [col_to_index(c) for c in _LETTERS.findall(ref.address)]
        return max(rows), max(cols)
    if ref.kind == "columns":
        match = _COLS.match(ref.address)
        assert match is not None
        return 0, max(col_to_index(match.group(1)), col_to_index(match.group(2)))
    if ref.kind == "rows":
        match = _ROWS.match(ref.address)
        assert match is not None
        return max(int(match.group(1)), int(match.group(2))), 0
    return 0, 0


def unquoted_sheet_with_space(formula: str) -> str | None:
    """Return the first sheet prefix that contains a space but is not quoted, e.g. My Sheet!A1."""
    text = re.sub(r'"[^"]*"', '""', formula)
    for match in re.finditer(r"!", text):
        start = match.start()
        if start and text[start - 1] == "'":
            continue
        segment_start = max(text.rfind(ch, 0, start) for ch in "=(,;+-*/&<>^{")
        segment = text[segment_start + 1 : start]
        if " " in segment.strip() and not segment.strip().startswith("'"):
            return segment.strip()
    return None


_NAME_RULE = re.compile(r"^[A-Za-z_\\][A-Za-z0-9_.\\]*$")
_CELL_LIKE = re.compile(r"^([A-Za-z]{1,3}\d+|[Rr]\d*[Cc]\d*)$")


def valid_defined_name(name: str | None) -> bool:
    """Excel's rules for defined names and table names (no spaces, not a cell address)."""
    if not name:
        return False
    return bool(_NAME_RULE.match(name)) and not _CELL_LIKE.match(name)
