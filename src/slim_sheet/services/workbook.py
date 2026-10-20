"""Shared helpers: open a sandboxed file and resolve sheet names with helpful errors."""

from __future__ import annotations

from collections.abc import Iterator
from contextlib import ExitStack, contextmanager
from dataclasses import dataclass
from pathlib import Path

from slim_sheet.adapters import crypto
from slim_sheet.adapters.formats import sniff
from slim_sheet.adapters.readers import SheetReader, open_reader
from slim_sheet.adapters.sandbox import resolve_path
from slim_sheet.config import Settings
from slim_sheet.core.errors import ErrorCode, SlimError
from slim_sheet.core.models import FileFormat


@dataclass(frozen=True, slots=True)
class OpenFile:
    path: Path
    display: str
    format: FileFormat
    reader: SheetReader


@contextmanager
def open_file(raw_path: str, settings: Settings) -> Iterator[OpenFile]:
    """Open a sandboxed file for reading; encrypted workbooks are decrypted to a temp copy."""
    path = resolve_path(raw_path, settings)
    with ExitStack() as stack:
        readable = path
        if crypto.is_encrypted(path):
            readable = stack.enter_context(crypto.decrypted(path, settings.password))
        fmt = sniff(readable)
        reader = stack.enter_context(open_reader(readable, fmt))
        yield OpenFile(path=readable, display=path.name, format=fmt, reader=reader)


def pick_sheet(handle: OpenFile, requested: str | None) -> str:
    names = [name for name, _ in handle.reader.sheets()]
    if not names:
        raise SlimError(ErrorCode.READ_FAILED, f"'{handle.display}' contains no sheets")
    if requested is None or requested == "":
        visible = [name for name, shown in handle.reader.sheets() if shown]
        return (visible or names)[0]
    if requested in names:
        return requested
    folded = {n.casefold(): n for n in names}
    if requested.casefold() in folded:
        return folded[requested.casefold()]
    retry = f"read(path='{handle.display}', sheet='{names[0]}')"
    raise SlimError(
        ErrorCode.SHEET_NOT_FOUND,
        f"No sheet '{requested}' in {handle.display}",
        hint=f"Sheets: {', '.join(names)}. Retry: {retry}",
    )
