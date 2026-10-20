"""Grid and keyed-table comparison (PLAN §9). Pure, no I/O.

Values compare loosely: 1 == 1.0, None == "", and strings ignore surrounding spaces,
so re-exports do not show up as noise.
"""

from __future__ import annotations

from collections.abc import Iterator, Sequence
from dataclasses import dataclass

from slim_sheet.core.errors import ErrorCode, SlimError
from slim_sheet.core.models import CellValue


@dataclass(frozen=True, slots=True)
class CellDiff:
    row: int
    col: int
    before: CellValue
    after: CellValue


@dataclass(frozen=True, slots=True)
class KeyedDiff:
    added: list[str]
    removed: list[str]
    changed: list[tuple[str, str, CellValue, CellValue]]  # (key, column, before, after)
    columns_added: list[str]
    columns_removed: list[str]
    duplicate_keys: int


def same(a: CellValue, b: CellValue) -> bool:
    a = None if a == "" else a
    b = None if b == "" else b
    if isinstance(a, str) and isinstance(b, str):
        return a.strip() == b.strip()
    if isinstance(a, bool) or isinstance(b, bool):
        return a == b and type(a) is type(b)
    if isinstance(a, int | float) and isinstance(b, int | float):
        return float(a) == float(b)
    return a == b


def diff_grids(
    before: Sequence[Sequence[CellValue]], after: Sequence[Sequence[CellValue]]
) -> Iterator[CellDiff]:
    """Both grids start at A1. Yields differences in row-major order."""
    for r in range(max(len(before), len(after))):
        row_a = before[r] if r < len(before) else ()
        row_b = after[r] if r < len(after) else ()
        for c in range(max(len(row_a), len(row_b))):
            va = row_a[c] if c < len(row_a) else None
            vb = row_b[c] if c < len(row_b) else None
            if not same(va, vb):
                yield CellDiff(r, c, va, vb)


def diff_keyed(
    cols_a: Sequence[str],
    body_a: Sequence[Sequence[CellValue]],
    cols_b: Sequence[str],
    body_b: Sequence[Sequence[CellValue]],
    key: str,
) -> KeyedDiff:
    ka, kb = _key_index(cols_a, key, "first"), _key_index(cols_b, key, "second")
    rows_a, dup_a = _by_key(body_a, ka)
    rows_b, dup_b = _by_key(body_b, kb)
    folded_b = {c.casefold(): i for i, c in enumerate(cols_b)}
    shared = [
        (i, c, folded_b[c.casefold()]) for i, c in enumerate(cols_a) if c.casefold() in folded_b
    ]
    changed = []
    for k in rows_a.keys() & rows_b.keys():
        ra, rb = rows_a[k], rows_b[k]
        for ia, name, ib in shared:
            va = ra[ia] if ia < len(ra) else None
            vb = rb[ib] if ib < len(rb) else None
            if not same(va, vb):
                changed.append((k, name, va, vb))
    changed.sort(key=lambda x: (x[0], x[1]))
    folded_a = {c.casefold() for c in cols_a}
    return KeyedDiff(
        added=sorted(rows_b.keys() - rows_a.keys()),
        removed=sorted(rows_a.keys() - rows_b.keys()),
        changed=changed,
        columns_added=[c for c in cols_b if c.casefold() not in folded_a],
        columns_removed=[c for c in cols_a if c.casefold() not in folded_b],
        duplicate_keys=dup_a + dup_b,
    )


def _key_index(columns: Sequence[str], key: str, side: str) -> int:
    for i, name in enumerate(columns):
        if name.casefold() == key.casefold():
            return i
    raise SlimError(
        ErrorCode.BAD_ARGUMENT,
        f"Key column '{key}' not in the {side} table",
        hint=f"Columns: {', '.join(columns)}",
    )


def _by_key(
    body: Sequence[Sequence[CellValue]], index: int
) -> tuple[dict[str, Sequence[CellValue]], int]:
    rows: dict[str, Sequence[CellValue]] = {}
    duplicates = 0
    for row in body:
        raw = row[index] if index < len(row) else None
        if raw is None or raw == "":
            continue
        key = _key_text(raw)
        if key in rows:
            duplicates += 1
            continue
        rows[key] = row
    return rows, duplicates


def _key_text(value: CellValue) -> str:
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


def diff_formulas(
    before: dict[tuple[int, int], str], after: dict[tuple[int, int], str]
) -> list[tuple[int, int, str | None, str | None]]:
    """Formula changes by cell, ignoring whitespace and function-name case outside strings."""
    changes = []
    for cell in sorted(before.keys() | after.keys()):
        fa, fb = before.get(cell), after.get(cell)
        if fa is None or fb is None or _norm_formula(fa) != _norm_formula(fb):
            changes.append((cell[0], cell[1], fa, fb))
    return changes


def _norm_formula(text: str) -> str:
    parts = text.split('"')
    # even indexes are outside string literals
    return '"'.join(p if i % 2 else "".join(p.split()).upper() for i, p in enumerate(parts))
