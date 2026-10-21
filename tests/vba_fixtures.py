"""Build macro-enabled test workbooks without Excel: a minimal OLE compound file (MS-CFB v3)
holding a source-only VBA project (MS-OVBA). Test-only code; nothing here ships.

Excel-made projects (e.g. the add-ins in Office's Library folder) were used to check the reader
by hand; they are Microsoft's files and are not committed.
"""

from __future__ import annotations

import io
import itertools
import struct
import zipfile
from dataclasses import dataclass
from pathlib import Path

import openpyxl

from slim_sheet.adapters.ovba import compress

SECTOR, MINI, CUTOFF = 512, 64, 4096
FREE, END, FATSECT, NOSTREAM = 0xFFFFFFFF, 0xFFFFFFFE, 0xFFFFFFFD, 0xFFFFFFFF


@dataclass
class _Entry:
    name: str
    kind: int  # 1 storage, 2 stream, 5 root
    data: bytes = b""
    children: list[int] | None = None
    start: int = END
    left: int = NOSTREAM
    right: int = NOSTREAM
    child: int = NOSTREAM


def _sort_key(name: str) -> tuple[int, str]:
    return len(name), name.upper()


def _tree(streams: dict[str, bytes]) -> list[_Entry]:
    entries = [_Entry("Root Entry", 5, children=[])]
    index: dict[str, int] = {"": 0}
    for path in sorted(streams):
        parts = path.split("/")
        for depth in range(1, len(parts) + 1):
            key = "/".join(parts[:depth])
            if key in index:
                continue
            is_stream = depth == len(parts)
            entry = _Entry(parts[depth - 1], 2 if is_stream else 1, streams.get(key, b""))
            entry.children = None if is_stream else []
            index[key] = len(entries)
            entries.append(entry)
            parent = entries[index["/".join(parts[: depth - 1])]]
            assert parent.children is not None
            parent.children.append(index[key])
    for entry in entries:  # siblings as a sorted right-leaning chain (valid, if unbalanced)
        if entry.children:
            ordered = sorted(entry.children, key=lambda i: _sort_key(entries[i].name))
            entry.child = ordered[0]
            for a, b in itertools.pairwise(ordered):
                entries[a].right = b
    return entries


def build_cfb(streams: dict[str, bytes]) -> bytes:
    """streams: {'VBA/dir': b'...', 'PROJECT': b'...'} -> compound file bytes."""
    entries = _tree(streams)
    mini, minifat = bytearray(), []
    big: list[_Entry] = []
    for entry in entries:
        if entry.kind != 2:
            continue
        if len(entry.data) < CUTOFF:
            entry.start = len(mini) // MINI if entry.data else END
            count = -(-len(entry.data) // MINI)
            minifat += [entry.start + i + 1 for i in range(count - 1)] + ([END] if count else [])
            mini += entry.data + b"\0" * (-len(entry.data) % MINI)
        else:
            big.append(entry)
    dir_bytes = len(entries) * 128
    regions = [("dir", -(-dir_bytes // SECTOR)), ("minifat", -(-len(minifat) * 4 // SECTOR)),
               ("mini", -(-len(mini) // SECTOR))]  # fmt: skip
    regions += [(f"big{i}", -(-len(e.data) // SECTOR)) for i, e in enumerate(big)]
    used = sum(n for _, n in regions)
    fat_sectors = 1
    while fat_sectors * 128 < used + fat_sectors:
        fat_sectors += 1
    fat = [FREE] * (fat_sectors * 128)
    starts: dict[str, int] = {}
    cursor = 0
    for name, count in regions:
        starts[name] = cursor if count else END
        for i in range(count):
            fat[cursor + i] = cursor + i + 1 if i < count - 1 else END
        cursor += count
    for i in range(fat_sectors):
        fat[cursor + i] = FATSECT
    entries[0].start = starts["mini"]
    entries[0].data = bytes(mini)
    for i, entry in enumerate(big):
        entry.start = starts[f"big{i}"]
    body = bytearray()
    body += _pad(b"".join(_dir_entry(e) for e in entries))
    body += _pad(b"".join(struct.pack("<I", v) for v in minifat))
    body += _pad(bytes(mini))
    for entry in big:
        body += _pad(entry.data)
    body += b"".join(struct.pack("<I", v) for v in fat)
    return _header(fat_sectors, cursor, starts, len(minifat)) + bytes(body)


def _pad(data: bytes) -> bytes:
    return data + b"\0" * (-len(data) % SECTOR)


def _dir_entry(entry: _Entry) -> bytes:
    name = entry.name.encode("utf-16-le") + b"\0\0"
    size = len(entry.data) if entry.kind != 1 else 0
    return (
        name.ljust(64, b"\0")
        + struct.pack("<HBB", len(name), entry.kind, 1)
        + struct.pack("<III", entry.left, entry.right, entry.child)
        + b"\0" * 16  # CLSID
        + struct.pack("<I", 0)  # state bits
        + b"\0" * 16  # created / modified
        + struct.pack("<IQ", entry.start if (entry.kind != 1) else 0, size)
    )


def _header(
    fat_sectors: int, first_fat: int, starts: dict[str, int], minifat_entries: int
) -> bytes:
    difat = [first_fat + i for i in range(fat_sectors)] + [FREE] * (109 - fat_sectors)
    minifat_sectors = -(-minifat_entries * 4 // SECTOR)
    head = (
        bytes.fromhex("D0CF11E0A1B11AE1") + b"\0" * 16
        + struct.pack("<HHHHH", 0x003E, 3, 0xFFFE, 9, 6) + b"\0" * 6
        + struct.pack("<IIII", 0, fat_sectors, starts["dir"], 0)
        + struct.pack("<I", CUTOFF)
        + struct.pack("<II", starts["minifat"] if minifat_sectors else END, minifat_sectors)
        + struct.pack("<II", END, 0)
        + b"".join(struct.pack("<I", v) for v in difat)
    )  # fmt: skip
    assert len(head) == SECTOR
    return head


# --- VBA project ----------------------------------------------------------------------------


def _rec(record_id: int, data: bytes = b"") -> bytes:
    return struct.pack("<HI", record_id, len(data)) + data


def _dir(project: str, modules: list[tuple[str, str, int]]) -> bytes:
    """modules: (name, kind 'standard'|'class'|'document', module offset)."""
    out = bytearray()
    out += _rec(0x0001, struct.pack("<I", 1)) + _rec(0x0002, struct.pack("<I", 0x409))
    out += _rec(0x0014, struct.pack("<I", 0x409)) + _rec(0x0003, struct.pack("<H", 1252))
    out += _rec(0x0004, project.encode("cp1252")) + _rec(0x0005) + _rec(0x0040)
    out += _rec(0x0006) + _rec(0x003D) + _rec(0x0007, b"\0" * 4) + _rec(0x0008, b"\0" * 4)
    out += struct.pack("<HIIH", 0x0009, 4, 1, 0)  # PROJECTVERSION: 6 data bytes despite size 4
    out += _rec(0x000C) + _rec(0x003C)
    out += _rec(0x0016, b"stdole") + _rec(0x003E, "stdole".encode("utf-16-le"))
    out += _rec(0x000F, struct.pack("<H", len(modules))) + _rec(0x0013, struct.pack("<H", 0xFFFF))
    for name, kind, offset in modules:
        out += _rec(0x0019, name.encode("cp1252")) + _rec(0x0047, name.encode("utf-16-le"))
        out += _rec(0x001A, name.encode("cp1252")) + _rec(0x0032, name.encode("utf-16-le"))
        out += _rec(0x001C) + _rec(0x0048) + _rec(0x0031, struct.pack("<I", offset))
        out += _rec(0x001E, b"\0" * 4) + _rec(0x002C, struct.pack("<H", 0xFFFF))
        out += _rec(0x0021 if kind == "standard" else 0x0022) + _rec(0x002B)
    out += _rec(0x0010)
    return bytes(out)


NAME_LINE = 'Attribute VB_Name = "{name}"\r\n'
DOCUMENT_HEADER = NAME_LINE + 'Attribute VB_Base = "0{{00020819-0000-0000-C000-000000000046}}"\r\n'
CLASS_HEADER = NAME_LINE + 'Attribute VB_Base = "0{{FCFB3D2A-A0FA-1068-A738-08002B3371B5}}"\r\n'
KIND_LINE = {"document": "Document={name}/&H00000000", "class": "Class={name}"}


def vba_project_bin(
    modules: dict[str, tuple[str, str]],
    project: str = "VBAProject",
    pcode: dict[str, bytes] | None = None,
    password: bool = False,
) -> bytes:
    """modules: {name: (kind, source)}. `pcode` fakes a compiled prefix (and optional empty
    source) per module, to model "VBA stomping"."""
    streams: dict[str, bytes] = {"VBA/_VBA_PROJECT": bytes.fromhex("cc61ffff000000")}
    table = []
    for name, (kind, source) in modules.items():
        header = {"document": DOCUMENT_HEADER, "class": CLASS_HEADER}.get(kind, NAME_LINE)
        text = (header.format(name=name) + source.replace("\n", "\r\n")).encode("cp1252")
        prefix = (pcode or {}).get(name, b"")
        streams[f"VBA/{name}"] = prefix + compress(text)
        table.append((name, kind, len(prefix)))
    streams["VBA/dir"] = compress(_dir(project, table))
    dpb = (
        "0E0CD1ECDFF4E7F5E7F5E7F5E7F5E7F5E7F5E7F5E7F5E7F5E7F5E7F5E7F5" if password else "0507070707"
    )
    lines = ['ID="{00000000-0000-0000-0000-000000000000}"']
    lines += [KIND_LINE.get(k, "Module={name}").format(name=n) for n, (k, _s) in modules.items()]
    lines += [
        f'Name="{project}"',
        'HelpContextID="0"',
        'VersionCompatible32="393222000"',
        f'DPB="{dpb}"',
        "",
    ]
    streams["PROJECT"] = "\r\n".join(lines).encode("cp1252")
    return build_cfb(streams)


VBA_TYPE = '<Default Extension="bin" ContentType="application/vnd.ms-office.vbaProject"/>'
VBA_REL = (
    '<Relationship Id="rIdVBA" Target="vbaProject.bin" '
    'Type="http://schemas.microsoft.com/office/2006/relationships/vbaProject"/>'
)


def macro_workbook(
    path: Path,
    modules: dict[str, tuple[str, str]],
    rows: list[list[object]] | None = None,
    **kwargs: object,
) -> Path:
    """An .xlsm with a Data sheet (`rows`, default a 2-row table) and the given VBA modules."""
    wb = openpyxl.Workbook()
    ws = wb.active
    assert ws is not None
    ws.title = "Data"
    for row in rows or [["Item", "Qty"], ["a", 1]]:
        ws.append(row)
    buffer = io.BytesIO()
    wb.save(buffer)
    blob = vba_project_bin(modules, **kwargs)  # type: ignore[arg-type]
    with zipfile.ZipFile(buffer) as source, zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as out:
        for item in source.infolist():
            data = source.read(item.filename)
            if item.filename == "[Content_Types].xml":
                text = data.decode()
                text = text.replace(
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml",
                    "application/vnd.ms-excel.sheet.macroEnabled.main+xml",
                ).replace("<Default ", VBA_TYPE + "<Default ", 1)
                data = text.encode()
            elif item.filename == "xl/_rels/workbook.xml.rels":
                data = (
                    data.decode().replace("</Relationships>", VBA_REL + "</Relationships>").encode()
                )
            out.writestr(item, data)
        out.writestr("xl/vbaProject.bin", blob)
    return path
