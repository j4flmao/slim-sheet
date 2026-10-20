"""csv/tsv/parquet/json via DuckDB (PLAN ADR-3). Exposed as one sheet named 'data'.

Row 1 holds the column names, data starts at row 2 — matching how Excel opens a CSV.
"""

from __future__ import annotations

import datetime as dt
import decimal
from pathlib import Path
from types import TracebackType
from typing import Self

import duckdb

from slim_sheet.adapters.readers.base import pad_window
from slim_sheet.core.errors import ErrorCode, SlimError
from slim_sheet.core.models import CellRange, CellValue, FileFormat

SHEET = "data"
SCAN = {
    FileFormat.CSV: "read_csv(?, auto_detect=true)",
    FileFormat.TSV: "read_csv(?, auto_detect=true, delim='\t')",
    FileFormat.PARQUET: "read_parquet(?)",
    FileFormat.JSON: "read_json_auto(?)",
}


class DuckReader:
    def __init__(self, path: Path, fmt: FileFormat) -> None:
        self._path = str(path)
        self._scan = SCAN[fmt]
        self._con = duckdb.connect(":memory:")
        try:
            rel = self._con.execute(f"SELECT * FROM {self._scan} LIMIT 0", [self._path])
            self._columns = [str(d[0]) for d in rel.description or []]
            count = self._con.execute(f"SELECT count(*) FROM {self._scan}", [self._path]).fetchone()
        except duckdb.Error as exc:
            self._con.close()
            raise SlimError(ErrorCode.READ_FAILED, f"Failed to parse '{path.name}': {exc}") from exc
        self._rows = int(count[0]) if count else 0

    def sheets(self) -> list[tuple[str, bool]]:
        return [(SHEET, True)]

    def _check(self, sheet: str) -> None:
        if sheet != SHEET:
            raise SlimError(
                ErrorCode.SHEET_NOT_FOUND,
                f"Flat files have a single sheet named '{SHEET}', not '{sheet}'",
                hint=f"Retry with sheet='{SHEET}' or omit sheet.",
            )

    def used_range(self, sheet: str) -> CellRange | None:
        self._check(sheet)
        if not self._columns:
            return None
        return CellRange(0, 0, self._rows, len(self._columns) - 1)

    def merged_count(self, sheet: str) -> int:
        self._check(sheet)
        return 0

    def read(self, sheet: str, window: CellRange) -> list[list[CellValue]]:
        self._check(sheet)
        header: list[list[CellValue]] = [list(self._columns)] if window.top == 0 else []
        first_data = max(window.top, 1) - 1
        limit = window.bottom - max(window.top, 1) + 1
        data: list[list[CellValue]] = []
        if limit > 0:
            sql = f"SELECT * FROM {self._scan} LIMIT ? OFFSET ?"
            fetched = self._con.execute(sql, [self._path, limit, first_data]).fetchall()
            data = [[to_cell(v) for v in row] for row in fetched]
        offset = window.top if header else first_data + 1
        return pad_window(header + data, window, row_offset=offset, col_offset=0)

    def close(self) -> None:
        self._con.close()

    def __enter__(self) -> Self:
        return self

    def __exit__(
        self,
        exc_type: type[BaseException] | None,
        exc: BaseException | None,
        tb: TracebackType | None,
    ) -> None:
        self.close()


def to_cell(value: object) -> CellValue:
    if value is None or isinstance(value, str | int | float | bool):
        return value
    if isinstance(value, dt.datetime | dt.date | dt.time | dt.timedelta):
        return value
    if isinstance(value, decimal.Decimal):
        return float(value)
    return str(value)
