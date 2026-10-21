"""Re-point references after cells are moved, the way Excel's cut/paste does. Pure.

Every reference (anywhere in the workbook) to a cell or range lying entirely inside the moved
block follows it: same offset, the destination sheet's name when it moved to another sheet,
and the original $-anchoring kept. References that only partly overlap the block are left
alone (Excel does the same). String literals are never touched.
"""

from __future__ import annotations

import re
from dataclasses import dataclass

from slim_sheet.core.a1 import col_to_index, index_to_col
from slim_sheet.core.models import CellRange
from slim_sheet.core.shift import quote_sheet

_REF = re.compile(
    r"(?<![A-Za-z0-9_.!$])"
    r"(?:(?P<sheet>'(?:[^']|'')+'|[A-Za-z_][A-Za-z0-9_.]*)!)?"
    r"(?P<a>\$?[A-Za-z]{1,3}\$?\d+)(?::(?P<b>\$?[A-Za-z]{1,3}\$?\d+))?"
    r"(?![A-Za-z0-9_(!])"
)
_CELL = re.compile(r"(\$?)([A-Za-z]{1,3})(\$?)(\d+)")


@dataclass(frozen=True, slots=True)
class Move:
    source_sheet: str
    block: CellRange  # 0-based, before the move
    dest_sheet: str
    rows: int  # offset
    cols: int


def _cell(text: str) -> tuple[str, int, str, int]:
    match = _CELL.fullmatch(text)
    assert match is not None
    col_abs, letters, row_abs, digits = match.groups()
    return col_abs, col_to_index(letters.upper()), row_abs, int(digits) - 1


def _inside(row: int, col: int, block: CellRange) -> bool:
    return block.top <= row <= block.bottom and block.left <= col <= block.right


def _sheet_name(raw: str | None, home: str) -> str:
    if raw is None:
        return home
    return raw[1:-1].replace("''", "'") if raw.startswith("'") else raw


def move_references(formula: str, home: str, move: Move) -> tuple[str, bool]:
    """`formula` lives on sheet `home`; returns the re-pointed formula and whether it changed."""
    changed = False

    def swap(match: re.Match[str]) -> str:
        nonlocal changed
        sheet = _sheet_name(match.group("sheet"), home)
        if sheet.casefold() != move.source_sheet.casefold():
            return match.group(0)
        ends = [match.group("a")] + ([match.group("b")] if match.group("b") else [])
        parsed = [_cell(e) for e in ends]
        if not all(_inside(r, c, move.block) for _, c, _, r in parsed):
            return match.group(0)
        changed = True
        moved = ":".join(
            f"{ca}{index_to_col(c + move.cols)}{ra}{r + move.rows + 1}" for ca, c, ra, r in parsed
        )
        needs_sheet = (
            match.group("sheet") is not None or move.dest_sheet.casefold() != home.casefold()
        )
        return f"{quote_sheet(move.dest_sheet)}!{moved}" if needs_sheet else moved

    parts = re.split(r'("(?:[^"]|"")*")', formula)
    out = "".join(p if p.startswith('"') else _REF.sub(swap, p) for p in parts)
    return out, changed


def qualify(formula: str, sheet: str) -> str:
    """Give unqualified references an explicit sheet ('=C2*10' -> '=Data!C2*10'), so a formula
    moved to another sheet keeps pointing at the cells it read before."""

    def add(match: re.Match[str]) -> str:
        if match.group("sheet") is not None:
            return match.group(0)
        return f"{quote_sheet(sheet)}!{match.group(0)}"

    parts = re.split(r'("(?:[^"]|"")*")', formula)
    return "".join(p if p.startswith('"') else _REF.sub(add, p) for p in parts)
