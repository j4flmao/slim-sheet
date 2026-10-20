"""Reader protocol shared by every format backend."""

from __future__ import annotations

from types import TracebackType
from typing import Protocol, Self

from slim_sheet.core.models import CellRange, CellValue


class SheetReader(Protocol):
    def sheets(self) -> list[tuple[str, bool]]:
        """(name, visible) in workbook order."""
        ...

    def used_range(self, sheet: str) -> CellRange | None: ...

    def merged_count(self, sheet: str) -> int: ...

    def read(self, sheet: str, window: CellRange) -> list[list[CellValue]]:
        """Exactly window.n_rows rows of window.n_cols values (padded with None)."""
        ...

    def close(self) -> None: ...

    def __enter__(self) -> Self: ...

    def __exit__(
        self,
        exc_type: type[BaseException] | None,
        exc: BaseException | None,
        tb: TracebackType | None,
    ) -> None: ...


def pad_window(
    rows: list[list[CellValue]], window: CellRange, row_offset: int, col_offset: int
) -> list[list[CellValue]]:
    """Cut `window` out of `rows` (whose [0][0] is at row_offset/col_offset), padding gaps."""
    out: list[list[CellValue]] = []
    for r in range(window.top, window.bottom + 1):
        src = rows[r - row_offset] if 0 <= r - row_offset < len(rows) else []
        line: list[CellValue] = []
        for c in range(window.left, window.right + 1):
            idx = c - col_offset
            value = src[idx] if 0 <= idx < len(src) else None
            line.append(None if value == "" else value)
        out.append(line)
    return out
