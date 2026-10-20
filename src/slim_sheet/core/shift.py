"""Rewrite A1 references when rows/columns are inserted or deleted, or a sheet is renamed.

openpyxl moves cell values but leaves every reference untouched, which silently breaks
workbooks. These pure functions apply Excel's own rules:
- insert at index k: every row/col index >= k moves by +count (absolute `$` refs too)
- delete [k, k+count): refs after the block move back; a range losing only some of its
  rows shrinks; a reference that lies entirely inside the block becomes #REF!
"""

from __future__ import annotations

import re
from collections.abc import Callable
from dataclasses import dataclass
from typing import Literal

from openpyxl.formula.tokenizer import Token, Tokenizer

from slim_sheet.core.a1 import col_to_index, index_to_col
from slim_sheet.core.refs import split_sheet

Axis = Literal["rows", "cols"]
RefRewriter = Callable[
    [str, str], str
]  # (reference token, sheet the formula lives on) -> new token
REF_ERROR = "#REF!"
_POINT = re.compile(r"^(\$?)([A-Za-z]{1,3})?(\$?)(\d+)?$")
_PLAIN_SHEET = re.compile(r"^[A-Za-z_][A-Za-z0-9_.]*$")
_LOOKS_LIKE_CELL = re.compile(r"^[A-Za-z]{1,3}\d+$")


@dataclass(frozen=True, slots=True)
class Shift:
    """Insert (count > 0) or delete (count < 0) `abs(count)` rows/cols at 1-based index `at`."""

    axis: Axis
    at: int
    count: int

    def index(self, value: int) -> int | None:
        """New 1-based index for a single row/col, or None if it was deleted."""
        if self.count > 0:
            return value + self.count if value >= self.at else value
        gone = -self.count
        if value < self.at:
            return value
        if value < self.at + gone:
            return None
        return value - gone

    def span(self, start: int, end: int) -> tuple[int, int] | None:
        if self.count > 0:
            return (self.index(start) or start, self.index(end) or end)
        gone = -self.count
        first, last = self.at, self.at + gone - 1
        if first <= start and end <= last:
            return None
        new_start = start if start < first else (first if start <= last else start - gone)
        new_end = end if end < first else (first - 1 if end <= last else end - gone)
        return new_start, new_end


@dataclass
class _Point:
    col_abs: str
    col: int | None  # 1-based
    row_abs: str
    row: int | None

    @classmethod
    def parse(cls, text: str) -> _Point | None:
        match = _POINT.match(text)
        if not match or (match.group(2) is None and match.group(4) is None):
            return None
        col = col_to_index(match.group(2)) + 1 if match.group(2) else None
        row = int(match.group(4)) if match.group(4) else None
        return cls(match.group(1), col, match.group(3), row)

    def render(self) -> str:
        col = f"{self.col_abs}{index_to_col(self.col - 1)}" if self.col is not None else ""
        row = f"{self.row_abs}{self.row}" if self.row is not None else ""
        return col + row


def shift_address(address: str, shift: Shift) -> str:
    """'B2', 'A1:C9', 'A:C', '3:5' shifted; returns REF_ERROR if wholly deleted."""
    parts = address.split(":")
    points = [_Point.parse(p) for p in parts]
    if not points or any(p is None for p in points) or len(points) > 2:
        return address  # not an A1 address (name, table, error); leave untouched
    a, b = points[0], points[-1]
    assert a is not None and b is not None
    attr = "row" if shift.axis == "rows" else "col"
    start, end = getattr(a, attr), getattr(b, attr)
    if start is None or end is None:
        return address  # e.g. 'A:C' under a row shift: whole columns are unaffected
    moved = shift.span(start, end)
    if moved is None:
        return REF_ERROR
    setattr(a, attr, moved[0])
    setattr(b, attr, moved[1])
    return a.render() if len(points) == 1 else f"{a.render()}:{b.render()}"


def _rewrite_refs(formula: str, home_sheet: str, fn: RefRewriter) -> str:
    if not formula.startswith("="):
        return formula
    tokenizer = Tokenizer(formula)
    changed = False
    for token in tokenizer.items:
        if token.type == Token.OPERAND and token.subtype == Token.RANGE:
            new = fn(str(token.value), home_sheet)
            if new != token.value:
                token.value = new
                changed = True
    return tokenizer.render() if changed else formula


class _Shifter:
    def __init__(self, target_sheet: str, shift: Shift) -> None:
        self.target = target_sheet.casefold()
        self.shift = shift
        self.count = 0

    def __call__(self, ref: str, home_sheet: str) -> str:
        sheet, address = split_sheet(ref)
        if (sheet or home_sheet).casefold() != self.target:
            return ref
        new = shift_address(address, self.shift)
        if new == address:
            return ref
        self.count += 1
        prefix = ref[: len(ref) - len(address)] if sheet is not None else ""
        return REF_ERROR if new == REF_ERROR else prefix + new


class _Renamer:
    def __init__(self, old: str, new: str) -> None:
        self.old = old.casefold()
        self.new_prefix = quote_sheet(new) + "!"
        self.count = 0

    def __call__(self, ref: str, home_sheet: str) -> str:
        sheet, address = split_sheet(ref)
        if sheet is None or sheet.casefold() != self.old:
            return ref
        self.count += 1
        return self.new_prefix + address


def shift_formula(
    formula: str, home_sheet: str, target_sheet: str, shift: Shift
) -> tuple[str, bool]:
    """Shift references to `target_sheet` inside a formula living on `home_sheet`."""
    shifter = _Shifter(target_sheet, shift)
    new = _rewrite_refs(formula, home_sheet, shifter)
    return new, shifter.count > 0


def rename_in_formula(formula: str, old: str, new: str) -> tuple[str, bool]:
    renamer = _Renamer(old, new)
    result = _rewrite_refs(formula, "", renamer)
    return result, renamer.count > 0


def quote_sheet(name: str) -> str:
    if _PLAIN_SHEET.match(name) and not _LOOKS_LIKE_CELL.match(name):
        return name
    return "'" + name.replace("'", "''") + "'"


def shift_sqref(sqref: str, shift: Shift) -> str:
    """Space-separated ranges (data validation / conditional formatting); deleted ones vanish."""
    kept = [r for r in (shift_address(p, shift) for p in sqref.split()) if r != REF_ERROR]
    return " ".join(kept)
