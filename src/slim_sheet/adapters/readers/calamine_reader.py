"""xlsx/xlsm/xlsb/xls/ods via python-calamine (PLAN ADR-1)."""

from __future__ import annotations

from pathlib import Path
from types import TracebackType
from typing import Any, Self

from python_calamine import CalamineWorkbook

from slim_sheet.adapters.readers.base import pad_window
from slim_sheet.core.errors import ErrorCode, SlimError
from slim_sheet.core.models import CellRange, CellValue


class CalamineReader:
    def __init__(self, path: Path) -> None:
        self._path = path
        try:
            self._wb = CalamineWorkbook.from_path(str(path))
        except PermissionError as exc:
            raise SlimError(
                ErrorCode.FILE_LOCKED,
                f"'{path.name}' is locked by another program",
                hint="Close it in Excel and retry.",
            ) from exc
        except Exception as exc:
            raise SlimError(ErrorCode.READ_FAILED, f"Failed to open '{path.name}': {exc}") from exc
        self._cache: dict[str, Any] = {}
        self._rows_cache: dict[str, list[list[CellValue]]] = {}
        self._complete: dict[str, bool] = {}

    def sheets(self) -> list[tuple[str, bool]]:
        result = []
        for meta in self._wb.sheets_metadata:
            visible = str(meta.visible).lower().endswith("visible")
            result.append((str(meta.name), visible))
        return result

    def _sheet(self, name: str) -> Any:
        if name not in self._cache:
            self._cache[name] = self._wb.get_sheet_by_name(name)
        return self._cache[name]

    def used_range(self, sheet: str) -> CellRange | None:
        ws = self._sheet(sheet)
        if ws.start is None or ws.end is None or ws.height == 0:
            return None
        (top, left), (bottom, right) = ws.start, ws.end
        return CellRange(int(top), int(left), int(bottom), int(right))

    def merged_count(self, sheet: str) -> int:
        ranges = self._sheet(sheet).merged_cell_ranges
        return len(ranges) if ranges else 0

    def read(self, sheet: str, window: CellRange) -> list[list[CellValue]]:
        rows = self._rows(sheet, window.bottom + 1)
        return pad_window(rows, window, row_offset=0, col_offset=0)

    def _rows(self, sheet: str, needed: int) -> list[list[CellValue]]:
        """Rows from A1 down to at least `needed`; grows geometrically so scans stay linear."""
        cached = self._rows_cache.get(sheet)
        if cached is not None and (len(cached) >= needed or self._complete.get(sheet, False)):
            return cached
        want = max(needed, 2 * len(cached or []))
        rows: list[list[CellValue]] = self._sheet(sheet).to_python(
            skip_empty_area=False, nrows=want
        )
        self._rows_cache[sheet] = rows
        self._complete[sheet] = len(rows) < want
        return rows

    def close(self) -> None:
        self._cache.clear()
        self._rows_cache.clear()
        self._wb.close()

    def __enter__(self) -> Self:
        return self

    def __exit__(
        self,
        exc_type: type[BaseException] | None,
        exc: BaseException | None,
        tb: TracebackType | None,
    ) -> None:
        self.close()
