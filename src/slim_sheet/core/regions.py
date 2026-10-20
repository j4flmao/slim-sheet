"""Split a sheet into rectangular data regions (PLAN §6.1). Pure, no I/O.

Rows separated by >= ROW_GAP empty rows start a new band; inside a band, columns
separated by >= COL_GAP empty columns start a new region. Regions smaller than
MIN_ROWS x MIN_COLS are dropped (titles, stray notes).
"""

from __future__ import annotations

from collections.abc import Sequence

from slim_sheet.core.budget import cell_text
from slim_sheet.core.models import CellRange, CellValue

ROW_GAP = 2
COL_GAP = 1
MIN_ROWS = 2
MIN_COLS = 1


def _filled(value: CellValue) -> bool:
    return value is not None and value != ""


def _runs(flags: Sequence[bool], gap: int) -> list[tuple[int, int]]:
    """Inclusive (start, end) runs of True, merging runs separated by fewer than `gap` Falses."""
    runs: list[tuple[int, int]] = []
    for index, flag in enumerate(flags):
        if not flag:
            continue
        if runs and index - runs[-1][1] - 1 < gap:
            runs[-1] = (runs[-1][0], index)
        else:
            runs.append((index, index))
    return runs


def detect_regions(rows: Sequence[Sequence[CellValue]], top: int, left: int) -> list[CellRange]:
    """Regions in sheet coordinates; `rows[0][0]` sits at (top, left). Largest region first."""
    width = max((len(r) for r in rows), default=0)
    row_flags = [any(_filled(v) for v in r) for r in rows]
    regions: list[CellRange] = []
    for r0, r1 in _runs(row_flags, ROW_GAP):
        band = rows[r0 : r1 + 1]
        col_flags = [any(c < len(r) and _filled(r[c]) for r in band) for c in range(width)]
        for c0, c1 in _runs(col_flags, COL_GAP):
            block = _trim_rows(band, c0, c1)
            if block is None:
                continue
            b0, b1 = block
            if b1 - b0 + 1 >= MIN_ROWS and c1 - c0 + 1 >= MIN_COLS:
                regions.append(CellRange(top + r0 + b0, left + c0, top + r0 + b1, left + c1))
    regions.sort(key=lambda g: (-(g.n_rows * g.n_cols), g.top, g.left))
    return regions


def _trim_rows(band: Sequence[Sequence[CellValue]], c0: int, c1: int) -> tuple[int, int] | None:
    hits = [i for i, r in enumerate(band) if any(_filled(v) for v in r[c0 : c1 + 1])]
    return (hits[0], hits[-1]) if hits else None


def column_names(header: Sequence[CellValue] | None, letters: Sequence[str]) -> list[str]:
    """Unique, non-empty SQL-friendly column names; falls back to column letters."""
    names: list[str] = []
    seen: set[str] = set()
    for index, letter in enumerate(letters):
        raw = header[index] if header is not None and index < len(header) else None
        base = " ".join(cell_text(raw, 200).split()) if _filled(raw) else letter
        name, n = base, 2
        while name.casefold() in seen:
            name, n = f"{base}_{n}", n + 1
        seen.add(name.casefold())
        names.append(name)
    return names
