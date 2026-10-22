"""Read-only VBA access (PLAN §11.10): find the VBA project in a workbook and decode its modules.

Works without Excel and never runs anything. Supported containers:
- xlsm / xltm / xlam / xlsb: the part whose content type is `application/vnd.ms-office.vbaProject`
  (normally xl/vbaProject.bin), an OLE compound file
- xls: the `_VBA_PROJECT_CUR` storage inside the OLE file
Module source is decompressed from the text part of each module stream (after MODULEOFFSET);
the compiled p-code before it is ignored, which is also how "VBA stomping" is spotted (p-code
present, source missing).
"""

from __future__ import annotations

import io
import re
import sys
import zipfile
from dataclasses import dataclass
from pathlib import Path

import olefile

from slim_sheet.adapters.ovba import MODULE_DOCUMENT, decompress, parse_dir
from slim_sheet.core.errors import ErrorCode, SlimError

VBA_CONTENT_TYPE = "application/vnd.ms-office.vbaproject"  # compared lower-case
MACROSHEET_TYPES = (
    "application/vnd.ms-excel.macrosheet+xml",
    "application/vnd.ms-excel.intlmacrosheet+xml",
)
MAX_PROJECT_BYTES = 64 * 1024 * 1024
_OVERRIDE = re.compile(r'<Override\b[^>]*\bPartName="([^"]+)"[^>]*\bContentType="([^"]+)"', re.I)
_DEFAULT = re.compile(r'<Default\b[^>]*\bExtension="([^"]+)"[^>]*\bContentType="([^"]+)"', re.I)
_DPB = re.compile(rb'DPB="([0-9A-Fa-f]*)"')


@dataclass(frozen=True, slots=True)
class VbaModule:
    name: str
    kind: str  # standard | class | document (sheet / workbook code)
    source: str
    stream_bytes: int  # p-code + source, as stored


@dataclass(frozen=True, slots=True)
class VbaProject:
    name: str
    modules: tuple[VbaModule, ...]
    references: tuple[str, ...]
    signed: bool
    locked: bool  # a VBA project password is set (heuristic on the DPB field)
    macro_sheets: int  # Excel 4.0 (XLM) macro sheets, which carry macros outside VBA
    stomped: tuple[str, ...]  # modules with compiled code but no source text


def _package_part(archive: zipfile.ZipFile) -> tuple[str | None, int, bool]:
    """(vbaProject part name, XLM macro sheets, signed) from [Content_Types].xml."""
    try:
        types = archive.read("[Content_Types].xml").decode("utf-8", "replace")
    except KeyError:
        return None, 0, False
    part = None
    sheets = 0
    for name, ctype in _OVERRIDE.findall(types):
        if ctype.lower() == VBA_CONTENT_TYPE:
            part = name.lstrip("/")
        elif ctype.lower() in MACROSHEET_TYPES:
            sheets += 1
    if part is None:  # Excel usually maps the type by extension: <Default Extension="bin" .../>
        extensions = {
            e.lower() for e, t in _DEFAULT.findall(types) if t.lower() == VBA_CONTENT_TYPE
        }
        part = next(
            (
                n
                for n in archive.namelist()
                if n.rsplit(".", 1)[-1].lower() in extensions and "signature" not in n.lower()
                and n.lower().endswith("vbaproject.bin")
            ),
            None,
        )  # fmt: skip
    signed = any("vbaprojectsignature" in n.lower() for n in archive.namelist())
    return part, sheets, signed


def has_macros(path: Path) -> tuple[bool, int, bool]:
    """(has a VBA project, XLM macro sheets, signed) without decoding anything."""
    if zipfile.is_zipfile(path):
        with zipfile.ZipFile(path) as archive:
            part, sheets, signed = _package_part(archive)
            return part is not None, sheets, signed
    if not olefile.isOleFile(str(path)):
        return False, 0, False
    with olefile.OleFileIO(str(path)) as doc:
        signed = doc.exists("\x05DigitalSignature") or doc.exists("_xmlsignatures")
        return doc.exists("_VBA_PROJECT_CUR/VBA/dir"), 0, bool(signed)


def read_project(path: Path) -> VbaProject | None:
    """The workbook's VBA project, or None when it has none."""
    if zipfile.is_zipfile(path):
        with zipfile.ZipFile(path) as archive:
            part, sheets, signed = _package_part(archive)
            if part is None:
                return None
            info = archive.getinfo(part)
            if info.file_size > MAX_PROJECT_BYTES:
                raise SlimError(ErrorCode.TOO_LARGE, f"VBA project in {path.name} is over 64 MB")
            blob = archive.read(part)
        with olefile.OleFileIO(io.BytesIO(blob)) as doc:
            return _project(doc, "", sheets, signed)
    if not olefile.isOleFile(str(path)):
        return None
    with olefile.OleFileIO(str(path)) as doc:
        if not doc.exists("_VBA_PROJECT_CUR/VBA/dir"):
            return None
        signed = doc.exists("\x05DigitalSignature") or doc.exists("_xmlsignatures")
        return _project(doc, "_VBA_PROJECT_CUR/", 0, bool(signed))


def _project(doc: object, root: str, macro_sheets: int, signed: bool) -> VbaProject:
    read = doc.openstream  # type: ignore[attr-defined]
    exists = doc.exists  # type: ignore[attr-defined]
    info = parse_dir(decompress(read(f"{root}VBA/dir").read()))
    modules: list[VbaModule] = []
    stomped: list[str] = []
    for record in info.modules:
        stream = f"{root}VBA/{record.stream or record.name}"
        if not exists(stream):
            continue
        data = read(stream).read()
        text = decompress(data[record.offset :]) if record.offset < len(data) else b""
        source = _decode(text, info.code_page)
        kind = _kind(record.kind, source)
        # empty sheet/workbook modules are normal; a code module with compiled code but no
        # source text is the "VBA stomping" signature (Excel may run code nobody can read)
        if kind != "document" and record.offset > 0 and not _code_lines(source):
            stomped.append(record.name)
        modules.append(VbaModule(record.name, kind, source, len(data)))
    project = read(f"{root}PROJECT").read() if exists(f"{root}PROJECT") else b""
    dpb = _DPB.search(project)
    locked = bool(dpb and len(dpb.group(1)) > 32)  # unprotected projects carry a short DPB
    return VbaProject(
        info.project,
        tuple(modules),
        tuple(info.references),
        signed,
        locked,
        macro_sheets,
        tuple(stomped),
    )


def _decode(text: bytes, code_page: int) -> str:
    try:
        source = text.decode(f"cp{code_page}")
    except (LookupError, UnicodeDecodeError) as exc:
        raise SlimError(
            ErrorCode.READ_FAILED,
            f"VBA source is not valid code page {code_page}: {exc}",
            hint="The project may be damaged; open it in Excel's VBA editor to check.",
        ) from exc
    return source.replace("\r\n", "\n")


def _code_lines(source: str) -> list[str]:
    return [
        line for line in source.splitlines() if line.strip() and not line.startswith("Attribute ")
    ]


def _kind(record_kind: int, source: str) -> str:
    if record_kind != MODULE_DOCUMENT:
        return "standard"
    # class modules and document modules share the record type; documents predeclare a host
    if re.search(r'^Attribute VB_Base = "0\{0002', source, re.M):
        return "document"
    return "class"


def zone_id(path: Path) -> int | None:
    """Windows Mark of the Web: 3 = internet, 4 = untrusted; None when unmarked or not Windows."""
    if sys.platform != "win32":
        return None
    try:
        marker = Path(f"{path}:Zone.Identifier")  # NTFS alternate data stream
        with marker.open(encoding="utf-8", errors="replace") as stream:
            for line in stream:
                if line.strip().lower().startswith("zoneid="):
                    return int(line.split("=", 1)[1].strip())
    except (OSError, ValueError):
        return None
    return None


def project_part(path: Path) -> tuple[str, bytes] | None:
    """(part name, bytes) of the package's VBA project, or None."""
    if not zipfile.is_zipfile(path):
        return None
    with zipfile.ZipFile(path) as archive:
        part, _sheets, _signed = _package_part(archive)
        return (part, archive.read(part)) if part else None


def restore_project(path: Path, part: str, blob: bytes) -> str:
    """Put the original VBA project back after another program re-saved the file.

    -> 'kept' (unchanged), 'restored' (part rewritten) or 'lost' (the re-save removed the part;
    its content types/relationships are gone too, so it is not re-added blindly)."""
    now = project_part(path)
    if now is None:
        return "lost"
    if now[1] == blob:
        return "kept"
    tmp = path.with_name(f".slim-vba-{path.name}")
    with zipfile.ZipFile(path) as source, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as out:
        for item in source.infolist():
            out.writestr(item, blob if item.filename == now[0] else source.read(item.filename))
    tmp.replace(path)
    return "restored"
