"""Dynamic-array (spilling) formulas through openpyxl saves (PLAN §11.10).

Excel stores a spilling formula as <f t="array" ref=...> plus cm="1" on its <c>, an index into
the XLDAPR cell metadata of xl/metadata.xml. openpyxl drops both: after any edit a spill became
a legacy CSE array, and a new FILTER showed as =@FILTER (one value). `extract` records which
cells were dynamic before the edit; `inject` marks them again after the save, together with the
cells slim-sheet wrote as dynamic, and adds the metadata part (Excel 365's own XML).
"""

from __future__ import annotations

import contextlib
import logging
import os
import re
import tempfile
import zipfile
from pathlib import Path

from slim_sheet.adapters.writers.sparklines import sheet_parts

log = logging.getLogger(__name__)

METADATA_PART = "xl/metadata.xml"
METADATA_TYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml"
METADATA_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sheetMetadata"
METADATA_XML = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    '<metadata xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
    'xmlns:xda="http://schemas.microsoft.com/office/spreadsheetml/2017/dynamicarray">'
    '<metadataTypes count="1"><metadataType name="XLDAPR" minSupportedVersion="120000" copy="1" '
    'pasteAll="1" pasteValues="1" merge="1" splitFirst="1" rowColShift="1" clearFormats="1" '
    'clearComments="1" assign="1" coerce="1" cellMeta="1"/></metadataTypes>'
    '<futureMetadata name="XLDAPR" count="1"><bk><extLst>'
    '<ext uri="{bdbb8cdc-fa1e-496e-a857-3c3f30c029c3}">'
    '<xda:dynamicArrayProperties fDynamic="1" fCollapsed="0"/></ext></extLst></bk>'
    '</futureMetadata><cellMetadata count="1"><bk><rc t="1" v="0"/></bk></cellMetadata></metadata>'
)
_CELL = re.compile(r'<c\b([^>]*?)\br="([A-Z]{1,3}[0-9]+)"([^>]*)>(?=<f\b[^>]*\bt="array")')
_HAS_CM = re.compile(r'\bcm="\d+"')
_CM_CELL = re.compile(r'<c\b[^>]*\bcm="\d+"[^>]*>')
_REF = re.compile(r'\br="([A-Z]{1,3}[0-9]+)"')


def extract(path: Path) -> dict[str, set[str]]:
    """Sheet title -> addresses of cells Excel marked as dynamic arrays (cm attribute)."""
    found: dict[str, set[str]] = {}
    try:
        with zipfile.ZipFile(path) as archive:
            if METADATA_PART not in archive.namelist():
                return found
            for title, part in sheet_parts(archive).items():
                xml = archive.read(part).decode("utf-8")
                cells = {m.group(1) for tag in _CM_CELL.findall(xml) if (m := _REF.search(tag))}
                if cells:
                    found[title] = cells
    except (zipfile.BadZipFile, KeyError, OSError) as exc:
        log.debug("No dynamic-array cells read from %s: %s", path.name, exc)
    return found


def inject(path: Path, cells: dict[str, set[str]]) -> int:
    """Mark `cells` (that still hold an array formula) as dynamic; returns how many were marked."""
    wanted = {title: refs for title, refs in cells.items() if refs}
    if not wanted:
        return 0
    fd, tmp_name = tempfile.mkstemp(prefix=".slim-", suffix=path.suffix, dir=path.parent)
    os.close(fd)
    tmp = Path(tmp_name)
    try:
        marked = _rewrite(path, tmp, wanted)
        if marked:
            tmp.replace(path)
    finally:
        with contextlib.suppress(FileNotFoundError):
            tmp.unlink()
    return marked


def _rewrite(path: Path, tmp: Path, wanted: dict[str, set[str]]) -> int:
    marked = 0
    with zipfile.ZipFile(path) as source, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as out:
        parts = {part: wanted[t] for t, part in sheet_parts(source).items() if t in wanted}
        names = set(source.namelist())
        for item in source.infolist():
            data = source.read(item.filename)
            if item.filename in parts:
                xml, count = _mark(data.decode("utf-8"), parts[item.filename])
                data, marked = xml.encode("utf-8"), marked + count
            elif item.filename == "[Content_Types].xml" and METADATA_PART not in names:
                data = _with_content_type(data.decode("utf-8")).encode("utf-8")
            elif item.filename == "xl/_rels/workbook.xml.rels" and METADATA_PART not in names:
                data = _with_relationship(data.decode("utf-8")).encode("utf-8")
            out.writestr(item, data)
        if METADATA_PART not in names:
            out.writestr(METADATA_PART, METADATA_XML)
    return marked


def _mark(xml: str, refs: set[str]) -> tuple[str, int]:
    count = 0

    def add(match: re.Match[str]) -> str:
        nonlocal count
        before, ref, after = match.groups()
        if ref not in refs or _HAS_CM.search(before + after):
            return match.group(0)
        count += 1
        return f'<c{before}r="{ref}"{after} cm="1">'

    return _CELL.sub(add, xml), count


def _with_content_type(xml: str) -> str:
    override = f'<Override PartName="/{METADATA_PART}" ContentType="{METADATA_TYPE}"/>'
    return xml.replace("</Types>", f"{override}</Types>", 1)


def _with_relationship(xml: str) -> str:
    used = {int(n) for n in re.findall(r'\bId="rId(\d+)"', xml)}
    rid = f"rId{max(used, default=0) + 1}"
    rel = f'<Relationship Id="{rid}" Type="{METADATA_REL}" Target="metadata.xml"/>'
    return xml.replace("</Relationships>", f"{rel}</Relationships>", 1)
