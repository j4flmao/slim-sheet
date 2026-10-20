"""Reader factory: picks a backend for a sniffed format."""

from __future__ import annotations

from pathlib import Path

from slim_sheet.adapters.readers.base import SheetReader
from slim_sheet.core.models import FileFormat


def open_reader(path: Path, fmt: FileFormat) -> SheetReader:
    if fmt == FileFormat.NUMBERS:
        from slim_sheet.adapters.readers.numbers_reader import NumbersReader

        return NumbersReader(path)
    if fmt.is_workbook:
        from slim_sheet.adapters.readers.calamine_reader import CalamineReader

        return CalamineReader(path)
    from slim_sheet.adapters.readers.duck_reader import DuckReader

    return DuckReader(path, fmt)


__all__ = ["SheetReader", "open_reader"]
