"""MS-OVBA building blocks (PLAN §11.10): RLE-LZ77 compression and the `dir` stream. Pure bytes.

[MS-OVBA] §2.4.1 defines the compression used for the `dir` stream and module source text;
§2.3.4.2 defines the `dir` stream records (project information, references, modules). Only
what a reader needs is parsed; unknown records are skipped by their size field.
"""

from __future__ import annotations

import struct
from dataclasses import dataclass, field

from slim_sheet.core.errors import ErrorCode, SlimError

CHUNK = 4096
MAX_OUTPUT = 16 * 1024 * 1024  # a module or dir stream never needs more; stops zip-bomb style input


def _bad(detail: str) -> SlimError:
    return SlimError(ErrorCode.READ_FAILED, f"Corrupt VBA project: {detail}")


def _bit_count(position: int) -> int:
    """Bits for the offset part of a copy token at `position` bytes into the chunk (>= 4)."""
    bits = 4
    while (1 << bits) < position:
        bits += 1
    return min(bits, 12)


def decompress(data: bytes, limit: int = MAX_OUTPUT) -> bytes:
    """Decompress a CompressedContainer ([MS-OVBA] §2.4.1.3.1)."""
    if not data or data[0] != 0x01:
        raise _bad("compressed container must start with 0x01")
    out = bytearray()
    pos = 1
    while pos < len(data):
        if pos + 2 > len(data):
            break
        header = struct.unpack_from("<H", data, pos)[0]
        size = (header & 0x0FFF) + 3
        compressed = header & 0x8000
        end = min(pos + size, len(data))
        start = len(out)
        pos += 2
        if not compressed:
            out += data[pos : pos + CHUNK]
            pos += CHUNK
        else:
            pos = _chunk(data, pos, end, out, start)
        if len(out) > limit:
            raise _bad(f"decompressed size exceeds {limit} bytes")
    return bytes(out)


def _chunk(data: bytes, pos: int, end: int, out: bytearray, start: int) -> int:
    while pos < end:
        flags = data[pos]
        pos += 1
        for bit in range(8):
            if pos >= end:
                break
            if not flags & (1 << bit):
                out.append(data[pos])
                pos += 1
                continue
            if pos + 2 > end:
                raise _bad("truncated copy token")
            token = struct.unpack_from("<H", data, pos)[0]
            pos += 2
            bits = _bit_count(len(out) - start)
            length = (token & (0xFFFF >> bits)) + 3
            offset = (token >> (16 - bits)) + 1
            if offset > len(out) - start:
                raise _bad("copy token points before the chunk")
            for _ in range(length):
                out.append(out[-offset])
    return pos


def compress(data: bytes) -> bytes:
    """Compress into a CompressedContainer ([MS-OVBA] §2.4.1.3.6), greedy longest match."""
    out = bytearray(b"\x01")
    for chunk_start in range(0, len(data), CHUNK):
        chunk = data[chunk_start : chunk_start + CHUNK]
        body = _compress_chunk(chunk)
        if len(body) > CHUNK:  # incompressible: store raw (only legal for a full 4096-byte chunk)
            if len(chunk) != CHUNK:
                raise _bad("final chunk is incompressible and shorter than 4096 bytes")
            out += struct.pack("<H", 0x3FFF) + chunk
            continue
        out += struct.pack("<H", 0xB000 | (len(body) + 2 - 3)) + body
    return bytes(out)


def _compress_chunk(chunk: bytes) -> bytes:
    body = bytearray()
    pos = 0
    while pos < len(chunk):
        flag_at = len(body)
        body.append(0)
        flags = 0
        for bit in range(8):
            if pos >= len(chunk):
                break
            bits = _bit_count(pos)
            max_len = (0xFFFF >> bits) + 3
            offset, length = _match(chunk, pos, max_len, 1 << bits)
            if length >= 3:
                token = ((offset - 1) << (16 - bits)) | (length - 3)
                body += struct.pack("<H", token)
                flags |= 1 << bit
                pos += length
            else:
                body.append(chunk[pos])
                pos += 1
        body[flag_at] = flags
    return bytes(body)


def _match(chunk: bytes, pos: int, max_len: int, max_offset: int) -> tuple[int, int]:
    best_len, best_off = 0, 0
    for candidate in range(max(0, pos - max_offset), pos):
        length = 0
        while (
            length < max_len
            and pos + length < len(chunk)
            and chunk[candidate + length] == chunk[pos + length]
        ):
            length += 1
        if length > best_len:
            best_len, best_off = length, pos - candidate
    return best_off, best_len


# --- dir stream -----------------------------------------------------------------------------

MODULE_PROCEDURAL, MODULE_DOCUMENT = 0x0021, 0x0022


@dataclass
class ModuleRecord:
    name: str = ""
    stream: str = ""
    offset: int = 0
    kind: int = MODULE_PROCEDURAL
    private: bool = False
    read_only: bool = False


@dataclass
class DirInfo:
    project: str = ""
    code_page: int = 1252
    modules: list[ModuleRecord] = field(default_factory=list)
    references: list[str] = field(default_factory=list)


def parse_dir(raw: bytes) -> DirInfo:
    """Records of the decompressed `dir` stream ([MS-OVBA] §2.3.4.2)."""
    info = DirInfo()
    current: ModuleRecord | None = None
    pos = 0
    while pos + 6 <= len(raw):
        record_id, size = struct.unpack_from("<HI", raw, pos)
        pos += 6
        if record_id == 0x0009:  # PROJECTVERSION: Reserved=4 but 6 bytes of data follow
            size = 6
        if pos + size > len(raw):
            raise _bad(f"record 0x{record_id:04X} runs past the dir stream")
        data = raw[pos : pos + size]
        pos += size
        if record_id == 0x0010:  # dir terminator
            break
        current = _record(info, current, record_id, data)
    return info


def _text(data: bytes, code_page: int) -> str:
    try:
        return data.decode(f"cp{code_page}")
    except (LookupError, UnicodeDecodeError):
        return data.decode("latin-1")


def _record(
    info: DirInfo, current: ModuleRecord | None, rid: int, data: bytes
) -> ModuleRecord | None:
    if rid == 0x0003 and len(data) >= 2:
        info.code_page = struct.unpack_from("<H", data)[0]
    elif rid == 0x0004:
        info.project = _text(data, info.code_page)
    elif rid == 0x0016:
        info.references.append(_text(data, info.code_page))
    elif rid == 0x0019:
        current = ModuleRecord(name=_text(data, info.code_page))
        info.modules.append(current)
    elif current is not None:
        if rid == 0x0047 and data:
            current.name = data.decode("utf-16-le", "replace")
        elif rid == 0x001A:
            current.stream = _text(data, info.code_page)
        elif rid == 0x0032 and data:
            current.stream = data.decode("utf-16-le", "replace")
        elif rid == 0x0031 and len(data) >= 4:
            current.offset = struct.unpack_from("<I", data)[0]
        elif rid in (MODULE_PROCEDURAL, MODULE_DOCUMENT):
            current.kind = rid
        elif rid == 0x0025:
            current.read_only = True
        elif rid == 0x0028:
            current.private = True
        elif rid == 0x002B:  # module terminator
            current = None
    return current
