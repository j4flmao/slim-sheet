"""Keep what openpyxl drops on save: drawing shapes, custom XML / add-in parts, workbook extensions.

openpyxl only keeps charts and pictures in drawings and silently drops every shape. That
includes the buttons that run Office Scripts (Automate > Add in workbook) or macros, text boxes
and arrows. It also drops `customXml/` parts (SharePoint and add-in metadata), `xl/webextensions/`
(Office add-ins) and the workbook `extLst` (newer Excel features). `extract` records them from
the file before an edit; `inject` writes them back into the saved package:

- shapes: anchors from each sheet's drawing are appended to the new drawing (or a new drawing
  part), with fresh shape ids. Shapes that point at other parts (pictures, links: `r:embed`,
  `r:id`) cannot be re-linked safely and are reported instead;
- parts: copied with their relationships and content types, if the new package lacks them;
- workbook `extLst`: each `ext` the new workbook.xml lacks and that references no part is
  copied back.

Anchors are not moved when rows or columns are inserted or deleted (like sparklines).
"""

from __future__ import annotations

import contextlib
import logging
import os
import re
import tempfile
import zipfile
from dataclasses import dataclass, field
from pathlib import Path, PurePosixPath

from slim_sheet.adapters.writers.sparklines import sheet_parts

log = logging.getLogger(__name__)

XDR = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
DML = "http://schemas.openxmlformats.org/drawingml/2006/main"
REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
DRAWING_REL = REL_NS + "/drawing"
DRAWING_TYPE = "application/vnd.openxmlformats-officedocument.drawing+xml"
KEPT_PREFIXES = ("customXml/", "xl/webextensions/")

_ANCHOR = re.compile(
    r"<(?P<p>\w+:)?(?P<kind>twoCellAnchor|oneCellAnchor|absoluteAnchor)\b.*?</(?P=p)?(?P=kind)>",
    re.DOTALL,
)
_SHAPE = re.compile(r"<(?:\w+:)?(?:sp|cxnSp|grpSp)\b")
_OTHER = re.compile(r"<(?:\w+:)?(?:graphicFrame|pic|contentPart)\b")
_LINKED = re.compile(r'\br:(?:id|embed|link|pict)="')
_XMLNS = re.compile(r'\sxmlns:(\w+)="([^"]*)"')
_NAME = re.compile(r'<(?:\w+:)?cNvPr\b[^>]*\bname="([^"]*)"')
_ID = re.compile(r'(<(?:\w+:)?cNvPr\b[^>]*\bid=")(\d+)(")')
_REL = re.compile(r"<Relationship\b[^>]*/>")
_OVERRIDE = re.compile(r'<Override\b[^>]*PartName="/([^"]+)"[^>]*/>')
_EXT = re.compile(r"<(?:\w+:)?ext\b[^>]*\buri=\"([^\"]+)\".*?</(?:\w+:)?ext>", re.DOTALL)
# where <drawing> goes in a worksheet: before the first of these elements (ECMA-376 order)
_AFTER_DRAWING = ("legacyDrawing", "legacyDrawingHF", "drawingHF", "picture", "oleObjects",
                  "controls", "webPublishItems", "tableParts", "extLst")  # fmt: skip


@dataclass
class Kept:
    shapes: dict[str, list[str]] = field(default_factory=dict)  # sheet title -> anchor XML
    names: dict[str, list[str]] = field(default_factory=dict)  # sheet title -> shape names
    linked: list[str] = field(default_factory=list)  # "Sheet: name" of shapes not restorable
    parts: dict[str, bytes] = field(default_factory=dict)
    overrides: list[str] = field(default_factory=list)
    root_rels: list[str] = field(default_factory=list)
    workbook_rels: list[str] = field(default_factory=list)
    workbook_ext: list[tuple[str, str]] = field(default_factory=list)  # (uri, xml)

    @property
    def empty(self) -> bool:
        return not (self.shapes or self.parts or self.workbook_ext)


def extract(path: Path) -> Kept:
    kept = Kept()
    try:
        with zipfile.ZipFile(path) as archive:
            _extract_shapes(archive, kept)
            _extract_parts(archive, kept)
            _extract_workbook_ext(archive, kept)
    except (zipfile.BadZipFile, KeyError, OSError) as exc:
        log.debug("Nothing to keep from %s: %s", path.name, exc)
    return kept


def _read(archive: zipfile.ZipFile, name: str) -> str:
    try:
        return archive.read(name).decode("utf-8")
    except KeyError:
        return ""


def _rels_name(part: str) -> str:
    p = PurePosixPath(part)
    return str(p.parent / "_rels" / f"{p.name}.rels")


def _resolve(base_part: str, target: str) -> str:
    if target.startswith("/"):
        return target.lstrip("/")
    parts: list[str] = []
    for piece in (str(PurePosixPath(base_part).parent) + "/" + target).split("/"):
        if piece == "..":
            if parts:
                parts.pop()
        elif piece not in ("", "."):
            parts.append(piece)
    return "/".join(parts)


def _drawing_of(archive: zipfile.ZipFile, sheet_part: str) -> str | None:
    for rel in _REL.findall(_read(archive, _rels_name(sheet_part))):
        if f'Type="{DRAWING_REL}"' in rel:
            target = re.search(r'Target="([^"]+)"', rel)
            return _resolve(sheet_part, target.group(1)) if target else None
    return None


def _extract_shapes(archive: zipfile.ZipFile, kept: Kept) -> None:
    for title, sheet_part in sheet_parts(archive).items():
        drawing = _drawing_of(archive, sheet_part)
        xml = _read(archive, drawing) if drawing else ""
        if not xml:
            continue
        root = xml[: xml.index(">", xml.index("wsDr")) + 1]
        namespaces = dict(_XMLNS.findall(root))
        for match in _ANCHOR.finditer(xml):
            anchor = match.group(0)
            if not _SHAPE.search(anchor) or _OTHER.search(anchor):
                continue  # charts and pictures: openpyxl keeps those itself
            found = _NAME.search(anchor)
            name = found.group(1) if found else "shape"
            if _LINKED.search(anchor):
                kept.linked.append(f"{title}: {name}")
                continue
            kept.shapes.setdefault(title, []).append(_with_namespaces(anchor, namespaces))
            kept.names.setdefault(title, []).append(name)


def _with_namespaces(anchor: str, namespaces: dict[str, str]) -> str:
    """Declare the prefixes the anchor uses on the anchor itself, so it stands alone."""
    needed = {"xdr": XDR, "a": DML, **namespaces}
    used = {p for p in needed if f"<{p}:" in anchor or f" {p}:" in anchor}
    end = anchor.index(">")
    head, tail = anchor[:end], anchor[end:]
    declared = set(re.findall(r"\sxmlns:(\w+)=", head))
    extra = "".join(f' xmlns:{p}="{needed[p]}"' for p in sorted(used - declared))
    if head.endswith("/"):
        return head[:-1] + extra + "/" + tail
    return head + extra + tail


def _extract_parts(archive: zipfile.ZipFile, kept: Kept) -> None:
    names = archive.namelist()
    for name in names:
        if name.startswith(KEPT_PREFIXES):
            kept.parts[name] = archive.read(name)
    if not kept.parts:
        return
    types = _read(archive, "[Content_Types].xml")
    kept.overrides = [
        m.group(0) for m in _OVERRIDE.finditer(types) if m.group(1).startswith(KEPT_PREFIXES)
    ]
    kept.root_rels = _rels_to(_read(archive, "_rels/.rels"), "", kept.parts)
    kept.workbook_rels = _rels_to(
        _read(archive, "xl/_rels/workbook.xml.rels"), "xl/workbook.xml", kept.parts
    )


def _rels_to(xml: str, base: str, parts: dict[str, bytes]) -> list[str]:
    found = []
    for rel in _REL.findall(xml):
        target = re.search(r'Target="([^"]+)"', rel)
        if target and _resolve(base or "_", target.group(1)) in parts:
            found.append(rel)
    return found


def _extract_workbook_ext(archive: zipfile.ZipFile, kept: Kept) -> None:
    xml = _read(archive, "xl/workbook.xml")
    block = re.search(r"<(?:\w+:)?extLst>(.*?)</(?:\w+:)?extLst>", xml, re.DOTALL)
    if not block:
        return
    root = xml[: xml.index(">", xml.index("<workbook")) + 1] if "<workbook" in xml else ""
    namespaces = dict(_XMLNS.findall(root))
    for match in _EXT.finditer(block.group(1)):
        ext = match.group(0)
        if not _LINKED.search(ext):
            kept.workbook_ext.append((match.group(1), _ext_namespaces(ext, namespaces)))


def _ext_namespaces(ext: str, namespaces: dict[str, str]) -> str:
    end = ext.index(">")
    head = ext[:end]
    declared = set(re.findall(r"\sxmlns:(\w+)=", head))
    used = {p for p in namespaces if f"<{p}:" in ext or f" {p}:" in ext}
    extra = "".join(f' xmlns:{p}="{namespaces[p]}"' for p in sorted(used - declared))
    return head + extra + ext[end:]


# --- inject ---------------------------------------------------------------------------------


def inject(path: Path, kept: Kept, sheet_titles: list[str]) -> list[str]:
    """Write `kept` back into the saved package; returns notes for the reply."""
    if kept.empty:
        return _linked_note(kept)
    fd, tmp_name = tempfile.mkstemp(prefix=".slim-", suffix=path.suffix, dir=path.parent)
    os.close(fd)
    tmp = Path(tmp_name)
    try:
        with zipfile.ZipFile(path) as source:
            files = {item.filename: source.read(item.filename) for item in source.infolist()}
            infos = {item.filename: item for item in source.infolist()}
        restored = _inject_shapes(files, kept, set(sheet_titles))
        _inject_parts(files, kept)
        _inject_workbook_ext(files, kept)
        with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as out:
            for name, data in files.items():
                out.writestr(infos.get(name, name), data)
        tmp.replace(path)
    finally:
        with contextlib.suppress(FileNotFoundError):
            tmp.unlink()
    return _linked_note(kept) + restored


def _linked_note(kept: Kept) -> list[str]:
    if not kept.linked:
        return []
    shown = ", ".join(kept.linked[:5]) + (", ..." if len(kept.linked) > 5 else "")
    return [f"saved copy drops {len(kept.linked)} shape(s) linked to pictures or links: {shown}"]


def _text(files: dict[str, bytes], name: str) -> str:
    return files[name].decode("utf-8") if name in files else ""


def _inject_shapes(files: dict[str, bytes], kept: Kept, titles: set[str]) -> list[str]:
    archive_like = _MemoryArchive(files)
    parts = sheet_parts(archive_like)  # type: ignore[arg-type]
    count = 0
    for title, anchors in kept.shapes.items():
        if title not in titles or title not in parts:
            continue
        sheet_part = parts[title]
        drawing = _drawing_of(archive_like, sheet_part)  # type: ignore[arg-type]
        if drawing is None:
            drawing = _new_drawing(files, sheet_part)
        xml = _text(files, drawing)
        next_id = max((int(i) for i in re.findall(r'<(?:\w+:)?cNvPr\b[^>]*\bid="(\d+)"', xml)),
                      default=1) + 1  # fmt: skip
        fresh, next_id = _renumber(anchors, next_id)
        close = re.search(r"</(?:\w+:)?wsDr>\s*$", xml)
        if close is None:
            continue
        files[drawing] = (xml[: close.start()] + "".join(fresh) + xml[close.start() :]).encode()
        count += len(anchors)
    if not count:
        return []
    names = [n for t, ns in kept.names.items() if t in titles for n in ns]
    shown = ", ".join(names[:5]) + (", ..." if len(names) > 5 else "")
    return [f"kept {count} shape(s) openpyxl would drop (buttons, text boxes): {shown}"]


def _renumber(anchors: list[str], start: int) -> tuple[list[str], int]:
    out = []
    current = start
    for anchor in anchors:

        def swap(m: re.Match[str]) -> str:
            nonlocal current
            new = f"{m.group(1)}{current}{m.group(3)}"
            current += 1
            return new

        out.append(_ID.sub(swap, anchor))
    return out, current


def _new_drawing(files: dict[str, bytes], sheet_part: str) -> str:
    """Create an empty drawing part for a sheet that has none and link it."""
    taken = {int(n) for n in re.findall(r"xl/drawings/drawing(\d+)\.xml", " ".join(files))}
    number = max(taken, default=0) + 1
    drawing = f"xl/drawings/drawing{number}.xml"
    files[drawing] = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        f'<xdr:wsDr xmlns:xdr="{XDR}" xmlns:a="{DML}" xmlns:r="{REL_NS}"></xdr:wsDr>'
    ).encode()
    rels_name = _rels_name(sheet_part)
    rels = _text(files, rels_name) or (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        "</Relationships>"
    )
    rid = _free_id(rels)
    rel = f'<Relationship Id="{rid}" Type="{DRAWING_REL}" Target="/{drawing}"/>'
    files[rels_name] = rels.replace("</Relationships>", rel + "</Relationships>").encode()
    files[sheet_part] = _with_drawing_element(_text(files, sheet_part), rid).encode()
    types = _text(files, "[Content_Types].xml")
    override = f'<Override PartName="/{drawing}" ContentType="{DRAWING_TYPE}"/>'
    files["[Content_Types].xml"] = types.replace("</Types>", override + "</Types>").encode()
    return drawing


def _with_drawing_element(sheet: str, rid: str) -> str:
    element = f'<drawing xmlns:r="{REL_NS}" r:id="{rid}"/>'
    for name in _AFTER_DRAWING:
        found = re.search(rf"<(?:\w+:)?{name}\b", sheet)
        if found:
            return sheet[: found.start()] + element + sheet[found.start() :]
    close = sheet.rindex("</")
    return sheet[:close] + element + sheet[close:]


def _free_id(rels: str) -> str:
    used = {int(n) for n in re.findall(r'\bId="rId(\d+)"', rels)}
    return f"rId{max(used, default=0) + 1}"


def _inject_parts(files: dict[str, bytes], kept: Kept) -> None:
    missing = {name: data for name, data in kept.parts.items() if name not in files}
    if not missing:
        return
    files.update(missing)
    types = _text(files, "[Content_Types].xml")
    for override in kept.overrides:
        name = re.search(r'PartName="/([^"]+)"', override)
        if name and name.group(1) in missing and f'PartName="/{name.group(1)}"' not in types:
            types = types.replace("</Types>", override + "</Types>")
    files["[Content_Types].xml"] = types.encode()
    for rels_name, wanted in (("_rels/.rels", kept.root_rels),
                              ("xl/_rels/workbook.xml.rels", kept.workbook_rels)):  # fmt: skip
        rels = _text(files, rels_name)
        if not rels:
            continue
        for rel in wanted:
            target = re.search(r'Target="([^"]+)"', rel)
            if target is None or f'Target="{target.group(1)}"' in rels:
                continue
            fresh = re.sub(r'\bId="[^"]+"', f'Id="{_free_id(rels)}"', rel)
            rels = rels.replace("</Relationships>", fresh + "</Relationships>")
        files[rels_name] = rels.encode()


def _inject_workbook_ext(files: dict[str, bytes], kept: Kept) -> None:
    xml = _text(files, "xl/workbook.xml")
    if not kept.workbook_ext or not xml:
        return
    present = set(re.findall(r'<(?:\w+:)?ext\b[^>]*\buri="([^"]+)"', xml))
    adds = "".join(ext for uri, ext in kept.workbook_ext if uri not in present)
    if not adds:
        return
    block = re.search(r"</(?:\w+:)?extLst>", xml)
    if block:
        xml = xml[: block.start()] + adds + xml[block.start() :]
    else:
        close = xml.rindex("</")
        xml = xml[:close] + f"<extLst>{adds}</extLst>" + xml[close:]
    files["xl/workbook.xml"] = xml.encode()


class _MemoryArchive:
    """The two ZipFile methods sheet_parts/_drawing_of need, over an in-memory package."""

    def __init__(self, files: dict[str, bytes]) -> None:
        self._files = files

    def read(self, name: str) -> bytes:
        return self._files[name]

    def namelist(self) -> list[str]:
        return list(self._files)
