"""Sparklines: keep them through openpyxl saves and write new ones (PLAN §11.9 P2).

openpyxl has no sparkline model and drops the worksheet <extLst> on save, so sparklines made in
Excel vanished after any edit. Before an edit, `extract` reads each sheet's sparkline groups from
the package; after the atomic save, `inject` writes them (plus any new groups) back into the
sheet XML as an x14 extension. Ranges inside kept groups are not shifted by row/column edits.
"""

from __future__ import annotations

import contextlib
import os
import re
import tempfile
import zipfile
from pathlib import Path
from xml.sax.saxutils import escape

X14 = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"
XM = "http://schemas.microsoft.com/office/excel/2006/main"
EXT_URI = "{05C60535-1F16-4fd2-B633-F4F36F0B64E0}"
_GROUP = re.compile(r"<(\w+:)?sparklineGroup\b.*?</(\w+:)?sparklineGroup>", re.DOTALL)
_SHEET = re.compile(r'<sheet\b[^>]*\bname="([^"]*)"[^>]*\br:id="([^"]*)"', re.DOTALL)
_REL = re.compile(r'<Relationship\b[^>]*\bId="([^"]*)"[^>]*\bTarget="([^"]*)"', re.DOTALL)
_REL_ALT = re.compile(r'<Relationship\b[^>]*\bTarget="([^"]*)"[^>]*\bId="([^"]*)"', re.DOTALL)


def sheet_parts(archive: zipfile.ZipFile) -> dict[str, str]:
    """Sheet title -> 'xl/worksheets/sheetN.xml'."""
    workbook = archive.read("xl/workbook.xml").decode("utf-8")
    rels = archive.read("xl/_rels/workbook.xml.rels").decode("utf-8")
    targets = {rid: target for rid, target in _REL.findall(rels)}
    targets.update({rid: target for target, rid in _REL_ALT.findall(rels)})
    parts: dict[str, str] = {}
    for name, rid in _SHEET.findall(workbook):
        target = targets.get(rid, "")
        if target:
            path = target.lstrip("/") if target.startswith("/") else f"xl/{target}"
            parts[_unescape(name)] = path
    return parts


def _unescape(text: str) -> str:
    return (
        text.replace("&amp;", "&").replace("&lt;", "<").replace("&gt;", ">").replace("&quot;", '"')
    )


def extract(path: Path) -> dict[str, list[str]]:
    """Sheet title -> sparklineGroup elements (with x14 prefix) found in the package."""
    try:
        with zipfile.ZipFile(path) as archive:
            parts = sheet_parts(archive)
            found: dict[str, list[str]] = {}
            for title, part in parts.items():
                if part not in archive.namelist():
                    continue
                xml = archive.read(part).decode("utf-8")
                if "sparklineGroup" in xml:
                    found[title] = [_normalise(m.group(0)) for m in _GROUP.finditer(xml)]
            return found
    except (zipfile.BadZipFile, KeyError, OSError):
        return {}


def _normalise(group: str) -> str:
    """Use the x14/xm prefixes the injected extension declares."""
    group = re.sub(r"<(/?)\w+:(sparkline\w*|color\w+|sparklines)\b", r"<\1x14:\2", group)
    group = re.sub(r"<(/?)(sparkline\w*|color\w+|sparklines)\b", r"<\1x14:\2", group)
    return re.sub(r"<(/?)\w+:(f|sqref)>", r"<\1xm:\2>", group)


def line_group(sources: list[tuple[str, str]], colour: str, negative: str) -> str:
    """One line-sparkline group: (source range 'Sheet!B2:M2', target cell 'N2') per sparkline."""
    sparklines = "".join(
        f"<x14:sparkline><xm:f>{escape(src)}</xm:f><xm:sqref>{escape(cell)}</xm:sqref></x14:sparkline>"
        for src, cell in sources
    )
    colours = (
        f'<x14:colorSeries rgb="FF{colour}"/><x14:colorNegative rgb="FF{negative}"/>'
        '<x14:colorAxis rgb="FF000000"/>'
        f'<x14:colorMarkers rgb="FF{colour}"/><x14:colorFirst rgb="FF{colour}"/>'
        f'<x14:colorLast rgb="FF{colour}"/><x14:colorHigh rgb="FF{colour}"/>'
        f'<x14:colorLow rgb="FF{negative}"/>'
    )
    return (
        f'<x14:sparklineGroup displayEmptyCellsAs="gap" lineWeight="1.25">{colours}'
        f"<x14:sparklines>{sparklines}</x14:sparklines></x14:sparklineGroup>"
    )


def inject(path: Path, groups: dict[str, list[str]]) -> int:
    """Write sparkline groups into the saved package; returns how many groups were written."""
    wanted = {title: g for title, g in groups.items() if g}
    if not wanted:
        return 0
    written = 0
    fd, tmp_name = tempfile.mkstemp(prefix=".slim-", suffix=path.suffix, dir=path.parent)
    os.close(fd)
    tmp = Path(tmp_name)
    try:
        with (
            zipfile.ZipFile(path) as source,
            zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as out,
        ):
            parts = {
                part: wanted[title]
                for title, part in sheet_parts(source).items()
                if title in wanted
            }
            for item in source.infolist():
                data = source.read(item.filename)
                if item.filename in parts:
                    data = _with_groups(data.decode("utf-8"), parts[item.filename]).encode("utf-8")
                    written += len(parts[item.filename])
                out.writestr(item, data)
        tmp.replace(path)
    finally:
        with contextlib.suppress(FileNotFoundError):
            tmp.unlink()
    return written


def _with_groups(xml: str, groups: list[str]) -> str:
    ext = (
        f'<ext uri="{EXT_URI}" xmlns:x14="{X14}">'
        f'<x14:sparklineGroups xmlns:xm="{XM}">{"".join(groups)}</x14:sparklineGroups></ext>'
    )
    if "<extLst>" in xml:
        return xml.replace("</extLst>", f"{ext}</extLst>", 1)
    return xml.replace("</worksheet>", f"<extLst>{ext}</extLst></worksheet>", 1)
