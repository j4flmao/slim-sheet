"""Apple Numbers (.numbers) via numbers-parser (optional extra `numbers`).

Numbers documents hold several tables per sheet. Each table is exposed as one sheet:
"Sheet" when the sheet has a single table, otherwise "Sheet / Table".
"""

from __future__ import annotations

import datetime as dt
from pathlib import Path
from types import TracebackType
from typing import Any, Self

from slim_sheet.adapters.readers.base import pad_window
from slim_sheet.core.errors import ErrorCode, SlimError
from slim_sheet.core.models import CellRange, CellValue

INSTALL_HINT = (
    "Install the numbers extra: `uv sync --extra numbers` or "
    "`uvx --from 'slim-sheet[numbers]' slim-sheet mcp`."
)


class NumbersReader:
    def __init__(self, path: Path) -> None:
        try:
            from numbers_parser import Document
        except ImportError as exc:
            raise SlimError(
                ErrorCode.MISSING_EXTRA,
                f"'{path.name}' is an Apple Numbers file",
                hint=INSTALL_HINT,
            ) from exc
        try:
            document = Document(str(path))
        except Exception as exc:
            raise SlimError(ErrorCode.READ_FAILED, f"Failed to open '{path.name}': {exc}") from exc
        self._tables: dict[str, Any] = {}
        for sheet in document.sheets:
            tables = list(sheet.tables)
            for table in tables:
                name = sheet.name if len(tables) == 1 else f"{sheet.name} / {table.name}"
                self._tables[name] = table
        self._rows: dict[str, list[list[CellValue]]] = {}

    def _grid(self, sheet: str) -> list[list[CellValue]]:
        if sheet not in self._rows:
            raw = self._tables[sheet].rows(values_only=True)
            self._rows[sheet] = [[_clean(v) for v in row] for row in raw]
        return self._rows[sheet]

    def sheets(self) -> list[tuple[str, bool]]:
        return [(name, True) for name in self._tables]

    def used_range(self, sheet: str) -> CellRange | None:
        filled = [
            (r, c)
            for r, row in enumerate(self._grid(sheet))
            for c, value in enumerate(row)
            if value is not None and value != ""
        ]
        if not filled:
            return None
        rows, cols = [r for r, _ in filled], [c for _, c in filled]
        return CellRange(min(rows), min(cols), max(rows), max(cols))

    def merged_count(self, sheet: str) -> int:
        return len(getattr(self._tables[sheet], "merge_ranges", []) or [])

    def read(self, sheet: str, window: CellRange) -> list[list[CellValue]]:
        return pad_window(self._grid(sheet), window, row_offset=0, col_offset=0)

    def close(self) -> None:
        self._rows.clear()

    def __enter__(self) -> Self:
        return self

    def __exit__(
        self,
        exc_type: type[BaseException] | None,
        exc: BaseException | None,
        tb: TracebackType | None,
    ) -> None:
        self.close()


def _clean(value: Any) -> CellValue:
    """Numbers stores decimals; numbers-parser returns floats like 12.500000000000002."""
    if isinstance(value, float):
        rounded = float(f"{value:.15g}")
        return int(rounded) if rounded.is_integer() and abs(rounded) < 1e15 else rounded
    if value is None or isinstance(value, str | int | bool | dt.datetime | dt.date | dt.timedelta):
        return value
    return str(value)
