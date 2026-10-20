"""Format sniffing: magic bytes first, extension second (PLAN §6)."""

from __future__ import annotations

import zipfile
from pathlib import Path

from slim_sheet.core.errors import ErrorCode, SlimError
from slim_sheet.core.models import FileFormat

_OLE2 = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"
_ZIP = b"PK\x03\x04"
_PARQUET = b"PAR1"
_TEXT_EXT = {
    ".csv": FileFormat.CSV,
    ".txt": FileFormat.CSV,
    ".psv": FileFormat.CSV,
    ".tsv": FileFormat.TSV,
    ".tab": FileFormat.TSV,
    ".json": FileFormat.JSON,
    ".jsonl": FileFormat.JSON,
    ".ndjson": FileFormat.JSON,
}
_SUPPORTED = "xlsx xlsm xltx xltm xlsb xls ods numbers csv tsv txt parquet json ndjson"


def sniff(path: Path) -> FileFormat:
    with path.open("rb") as handle:
        head = handle.read(8)
    if head.startswith(_ZIP):
        return _sniff_zip(path)
    if head.startswith(_OLE2):
        return FileFormat.XLS
    if head.startswith(_PARQUET):
        return FileFormat.PARQUET
    fmt = _TEXT_EXT.get(path.suffix.lower())
    if fmt is not None:
        return fmt
    raise _unsupported(path, "unrecognised content")


def _sniff_zip(path: Path) -> FileFormat:
    try:
        with zipfile.ZipFile(path) as archive:
            names = set(archive.namelist())
            mimetype = (
                archive.read("mimetype").decode("ascii", "ignore") if "mimetype" in names else ""
            )
    except zipfile.BadZipFile as exc:
        raise SlimError(ErrorCode.READ_FAILED, f"'{path.name}' is a corrupt ZIP container") from exc
    if "xl/workbook.xml" in names:
        return FileFormat.XLSX
    if "xl/workbook.bin" in names:
        return FileFormat.XLSB
    if "opendocument.spreadsheet" in mimetype:
        return FileFormat.ODS
    if any(n.startswith("Index/") for n in names):
        return FileFormat.NUMBERS
    raise _unsupported(path, "ZIP archive without a spreadsheet")


def _unsupported(path: Path, why: str) -> SlimError:
    return SlimError(
        ErrorCode.UNSUPPORTED_FORMAT,
        f"Cannot read '{path.name}': {why}",
        hint=f"Supported: {_SUPPORTED}",
    )
