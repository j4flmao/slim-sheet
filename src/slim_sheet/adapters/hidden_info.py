"""What an xlsx package hides from a casual look (PLAN §11.10): `overview`'s hidden-content report.

Read straight from the zip parts (no workbook load): very-hidden sheets, hidden rows/columns,
hidden defined names, ';;;' number formats, notes, embedded objects and author fields. Sheet
parts are scanned in chunks so a large file costs one streaming pass and no memory spike.
"""

from __future__ import annotations

import logging
import re
import zipfile
from collections.abc import Iterator
from pathlib import Path

from slim_sheet.adapters.writers.sparklines import sheet_parts

log = logging.getLogger(__name__)

CHUNK = 1 << 20
MAX_AUTHOR_CHARS = 60
_VERY_HIDDEN = re.compile(r'<sheet\b[^>]*\bname="([^"]*)"[^>]*\bstate="veryHidden"')
_HIDDEN_ROW = re.compile(rb'<row\b[^>]*\bhidden="(?:1|true)"')
_COL = re.compile(rb"<col\b[^>]*>")
_HIDDEN = re.compile(rb'\bhidden="(?:1|true)"')
_SPAN = (re.compile(rb'\bmin="(\d+)"'), re.compile(rb'\bmax="(\d+)"'))
_HIDDEN_NAME = re.compile(r'<definedName\b[^>]*\bname="([^"]*)"[^>]*\bhidden="(?:1|true)"')
_HIDING_FORMAT = re.compile(r'<numFmt\b[^>]*\bformatCode=";;;?"')
_AUTHOR = re.compile(r"<(dc:creator|cp:lastModifiedBy)>([^<]+)</")


def hidden_content(path: Path) -> list[str]:
    """Plain-language findings; [] for non-xlsx files or packages that cannot be read."""
    try:
        with zipfile.ZipFile(path) as archive:
            names = set(archive.namelist())
            if "xl/workbook.xml" not in names:
                return []
            return [
                *_workbook_findings(archive),
                *_sheet_findings(archive),
                *_part_findings(archive, names),
            ]
    except (zipfile.BadZipFile, KeyError, OSError) as exc:
        log.debug("Hidden-content scan skipped for %s: %s", path.name, exc)
        return []


def _text(archive: zipfile.ZipFile, part: str) -> str:
    try:
        return archive.read(part).decode("utf-8", errors="replace")
    except KeyError:
        return ""


def _workbook_findings(archive: zipfile.ZipFile) -> Iterator[str]:
    workbook = _text(archive, "xl/workbook.xml")
    very = _VERY_HIDDEN.findall(workbook)
    if very:
        yield f"very hidden sheet(s), not listed by Excel's Unhide: {', '.join(very[:5])}"
    names = [n for n in _HIDDEN_NAME.findall(workbook) if not n.startswith("_xlnm.")]
    if names:
        yield f"{len(names)} hidden defined name(s): {', '.join(names[:5])}"
    if _HIDING_FORMAT.search(_text(archive, "xl/styles.xml")):
        yield "a ';;;' number format is defined: cells using it look empty but hold values"


def _sheet_findings(archive: zipfile.ZipFile) -> Iterator[str]:
    for title, part in sheet_parts(archive).items():
        rows, cols = _hidden_rows_cols(archive, part)
        if rows or cols:
            parts = [f"{n} {kind}" for n, kind in ((rows, "row(s)"), (cols, "column(s)")) if n]
            yield f"{title}: {' and '.join(parts)} hidden"


def _hidden_rows_cols(archive: zipfile.ZipFile, part: str) -> tuple[int, int]:
    rows = cols = 0
    tail = b""
    try:
        with archive.open(part) as stream:
            while chunk := stream.read(CHUNK):
                block = tail + chunk
                cut = block.rfind(b">") + 1  # never split a tag between two blocks
                body, tail = block[:cut], block[cut:]
                rows += len(_HIDDEN_ROW.findall(body))
                cols += sum(_span(tag) for tag in _COL.findall(body) if _HIDDEN.search(tag))
    except KeyError:
        return 0, 0
    return rows, cols


def _span(tag: bytes) -> int:
    low, high = (pattern.search(tag) for pattern in _SPAN)
    return int(high.group(1)) - int(low.group(1)) + 1 if low and high else 1


def _part_findings(archive: zipfile.ZipFile, names: set[str]) -> Iterator[str]:
    # Excel names them xl/commentsN.xml, openpyxl xl/comments/commentN.xml
    notes = [n for n in names if n.startswith("xl/comments") and n.endswith(".xml")]
    threads = [n for n in names if n.startswith("xl/threadedComments/")]
    if notes or threads:
        yield f"notes/comments on {len(notes) or len(threads)} sheet(s) (describe_cell shows them)"
    embedded = [n for n in names if n.startswith("xl/embeddings/")]
    if embedded:
        yield f"{len(embedded)} embedded object(s) (OLE/package files inside the workbook)"
    authors = {
        f"{'author' if tag == 'dc:creator' else 'last saved by'}={value.strip()[:MAX_AUTHOR_CHARS]}"
        for tag, value in _AUTHOR.findall(_text(archive, "docProps/core.xml"))
        if value.strip()
    }
    if authors:
        yield "document properties: " + ", ".join(sorted(authors))
