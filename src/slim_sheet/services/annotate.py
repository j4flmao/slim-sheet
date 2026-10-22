"""`annotate` use-case: notes, hyperlinks and images on cells, and removing objects (PLAN §17.2).

op = note | link | image | remove
- note:   text (empty text removes the note)
- link:   url = https://..., mailto:... or an in-workbook target like 'Summary!A1'; text (optional)
          becomes the cell's label. file:, javascript: and other schemes are refused.
- image:  image_path (png/jpg/gif inside the allowed roots) anchored at cell; needs extra `images`.
- remove: target = note | link | image | chart | all at that cell.
"""

from __future__ import annotations

import re
from dataclasses import dataclass
from typing import Any

from openpyxl.comments import Comment
from openpyxl.styles import Font
from openpyxl.worksheet.hyperlink import Hyperlink

from slim_sheet.adapters.sandbox import resolve_path
from slim_sheet.adapters.writers.chart_refs import anchor_cell
from slim_sheet.config import Settings
from slim_sheet.core.a1 import index_to_col, parse_range
from slim_sheet.core.errors import ErrorCode, SlimError
from slim_sheet.core.models import OutputResult
from slim_sheet.core.styles import choice
from slim_sheet.services.editing import Change, edit_workbook, get_sheet

OPS = ("note", "link", "image", "remove")
TARGETS = ("note", "link", "image", "chart", "all")
IMAGE_SUFFIXES = (".png", ".jpg", ".jpeg", ".gif", ".bmp")
MAX_NOTE = 2_000
_WEB = re.compile(r"^(https?://|mailto:)", re.IGNORECASE)
_INTERNAL = re.compile(r"^#?('?[^'!]+'?!)?\$?[A-Za-z]{1,3}\$?\d+$")


@dataclass(frozen=True, slots=True)
class AnnotateRequest:
    path: str
    sheet: str
    op: str
    cell: str
    text: str | None = None
    url: str | None = None
    image_path: str | None = None
    target: str = "all"
    out_path: str | None = None
    overwrite: bool = False


def annotate(req: AnnotateRequest, settings: Settings) -> OutputResult:
    op = choice(req.op, OPS, "op")
    rng = parse_range(req.cell)
    if rng.n_rows * rng.n_cols != 1:
        raise SlimError(
            ErrorCode.BAD_ARGUMENT, f"cell must be one cell, got {req.cell}", hint="e.g. cell='B2'"
        )
    image = _image(req, settings) if op == "image" else None
    link = _link_target(req.url) if op == "link" else None

    def mutate(wb: Any) -> Change:
        ws = get_sheet(wb, req.sheet)
        cell = ws.cell(row=rng.top + 1, column=rng.left + 1)
        where = f"{ws.title}!{cell.coordinate}"
        if op == "note":
            return _note(cell, req.text or "", where)
        if op == "link":
            return _link(cell, link or "", req.text, where)
        if op == "image":
            ws.add_image(image, cell.coordinate)
            return Change(f"image {req.image_path} placed at {where}")
        return _remove(ws, cell, choice(req.target, TARGETS, "target"), rng.top, rng.left, where)

    return edit_workbook(
        req.path,
        settings,
        f"annotate {op} {req.sheet}!{req.cell}",
        mutate,
        req.out_path,
        req.overwrite,
    )


def _note(cell: Any, text: str, where: str) -> Change:
    if not text.strip():
        cell.comment = None
        return Change(f"note removed from {where}")
    if len(text) > MAX_NOTE:
        raise SlimError(ErrorCode.TOO_LARGE, f"note is longer than {MAX_NOTE} characters")
    note = Comment(text, "slim-sheet")
    note.width, note.height = 240, max(60, min(400, 20 * (text.count("\n") + len(text) // 40 + 1)))
    cell.comment = note
    return Change(f"note set on {where}")


def _link_target(url: str | None) -> str:
    text = (url or "").strip()
    if _WEB.match(text):
        return text
    if _INTERNAL.match(text):
        return text if text.startswith("#") else f"#{text}"
    raise SlimError(
        ErrorCode.BAD_ARGUMENT,
        f"url '{text[:60]}' is not allowed",
        hint="Use https://..., mailto:... or a cell in this workbook like 'Summary!A1'",
    )


def _link(cell: Any, target: str, text: str | None, where: str) -> Change:
    if target.startswith("#"):  # a place in this workbook, stored as a location, not a URL
        cell.hyperlink = Hyperlink(ref=cell.coordinate, location=target[1:])
    else:
        cell.hyperlink = target
    if text is not None:
        cell.value = text
    elif cell.value in (None, ""):
        cell.value = target.lstrip("#")
    old = cell.font
    cell.font = Font(
        name=old.name, size=old.size, bold=old.bold, color="FF0563C1", underline="single"
    )
    return Change(f"link at {where} -> {target}")


def _image(req: AnnotateRequest, settings: Settings) -> Any:
    if not req.image_path:
        raise SlimError(
            ErrorCode.BAD_ARGUMENT, "op=image needs image_path", hint="e.g. image_path='logo.png'"
        )
    source = resolve_path(req.image_path, settings)
    if source.suffix.lower() not in IMAGE_SUFFIXES or not source.is_file():
        raise SlimError(
            ErrorCode.BAD_ARGUMENT,
            f"'{source.name}' is not an existing image file",
            hint=f"Use one of: {' '.join(IMAGE_SUFFIXES)}",
        )
    try:
        from openpyxl.drawing.image import Image
    except ImportError as exc:  # pragma: no cover - openpyxl always ships it
        raise SlimError(ErrorCode.MISSING_EXTRA, "openpyxl image support is missing") from exc
    try:
        return Image(str(source))
    except ImportError as exc:
        raise SlimError(
            ErrorCode.MISSING_EXTRA,
            "Inserting images needs Pillow",
            hint="Install the extra: uvx --from 'slim-sheet[images]' slim-sheet mcp",
        ) from exc


def _remove(ws: Any, cell: Any, target: str, row: int, col: int, where: str) -> Change:
    removed: list[str] = []
    if target in ("note", "all") and cell.comment is not None:
        cell.comment = None
        removed.append("note")
    if target in ("link", "all") and cell.hyperlink is not None:
        cell.hyperlink = None
        removed.append("link")
    for kind, holder in (("image", "_images"), ("chart", "_charts")):
        if target not in (kind, "all"):
            continue
        objects = getattr(ws, holder)
        keep = [obj for obj in objects if anchor_cell(obj) != (row, col)]
        removed += [kind] * (len(objects) - len(keep))
        objects[:] = keep
    if not removed:
        raise SlimError(
            ErrorCode.BAD_ARGUMENT,
            f"Nothing to remove at {where} (target={target})",
            hint=f"Objects are matched by their top-left cell, e.g. {index_to_col(col)}{row + 1}",
        )
    return Change(f"removed {', '.join(removed)} at {where}")
