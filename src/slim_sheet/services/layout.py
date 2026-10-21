"""`sheet_layout` use-case: how a sheet is viewed and printed, in one call (PLAN §17.2).

freeze, column widths, row heights, autofilter, hidden rows/columns, gridlines, zoom,
print area/orientation/fit-to-width and sheet protection. Every argument is optional and
only the given ones change. Protection never takes a password (secrets do not travel through
tool arguments); people add one in Excel if they need it.
"""

from __future__ import annotations

import re
from dataclasses import dataclass
from typing import Any

from openpyxl.workbook.properties import CalcProperties
from openpyxl.workbook.protection import WorkbookProtection
from openpyxl.worksheet.properties import PageSetupProperties

from slim_sheet.config import Settings
from slim_sheet.core.a1 import format_range as a1
from slim_sheet.core.a1 import index_to_col, parse_range
from slim_sheet.core.errors import ErrorCode, SlimError
from slim_sheet.core.models import OutputResult
from slim_sheet.core.styles import choice
from slim_sheet.services.editing import Change, edit_workbook, get_sheet
from slim_sheet.services.ranges import parse_span

MAX_WIDTH, MAX_HEIGHT = 255.0, 409.0
CALC_MODES = {"auto": "auto", "manual": "manual", "auto_no_tables": "autoNoTable"}
_SIZE = re.compile(r"^\s*([^=]+?)\s*=\s*(\d+(?:\.\d+)?)\s*$")


@dataclass(frozen=True, slots=True)
class LayoutRequest:
    path: str
    sheet: str
    freeze: str | None = None
    widths: list[str] | None = None
    heights: list[str] | None = None
    filter_range: str | None = None
    hide: list[str] | None = None
    unhide: list[str] | None = None
    gridlines: bool | None = None
    zoom: int | None = None
    print_area: str | None = None
    landscape: bool | None = None
    fit_width: bool | None = None
    protect: bool | None = None
    calc_mode: str | None = None  # workbook: auto | manual | auto_no_tables
    protect_structure: bool | None = None  # workbook: lock adding/moving/deleting sheets
    out_path: str | None = None
    overwrite: bool = False


def sheet_layout(req: LayoutRequest, settings: Settings) -> OutputResult:
    widths = [_size(item, "widths", MAX_WIDTH) for item in req.widths or []]
    heights = [_size(item, "heights", MAX_HEIGHT) for item in req.heights or []]
    if req.zoom is not None and not 10 <= req.zoom <= 400:
        raise SlimError(ErrorCode.BAD_ARGUMENT, "zoom must be 10..400", hint="e.g. zoom=100")
    if req.calc_mode is not None:
        choice(req.calc_mode, tuple(CALC_MODES), "calc_mode")

    def mutate(wb: Any) -> Change:
        ws = get_sheet(wb, req.sheet)
        done = _view(ws, req)
        done += _sizes(ws, widths, heights)
        done += _visibility(ws, req.hide or [], hidden=True)
        done += _visibility(ws, req.unhide or [], hidden=False)
        done += _printing(ws, req)
        done += _workbook(wb, req)
        if not done:
            raise SlimError(
                ErrorCode.BAD_ARGUMENT,
                "Nothing to change",
                hint="e.g. sheet_layout(path, sheet='Data', freeze='A2', widths=['A=14','B:D=12'])",
            )
        return Change(f"layout of '{ws.title}': " + "; ".join(done))

    return edit_workbook(
        req.path, settings, f"sheet_layout {req.sheet}", mutate, req.out_path, req.overwrite
    )


def _size(item: str, label: str, limit: float) -> tuple[str, float]:
    match = _SIZE.match(item)
    if not match or float(match.group(2)) > limit:
        example = "'A=14', 'B:D=12'" if label == "widths" else "'1=30', '2:5=18'"
        raise SlimError(
            ErrorCode.BAD_ARGUMENT,
            f"{label} item '{item}' is not 'span=size'",
            hint=f"e.g. {example}",
        )
    return match.group(1), float(match.group(2))


def _view(ws: Any, req: LayoutRequest) -> list[str]:
    done: list[str] = []
    if req.freeze is not None:
        text = req.freeze.strip().upper()
        if text in ("", "NONE", "A1"):
            ws.freeze_panes = None
            done.append("panes unfrozen")
        else:
            cell = parse_range(text)
            ws.freeze_panes = f"{index_to_col(cell.left)}{cell.top + 1}"
            done.append(f"frozen at {ws.freeze_panes}")
    if req.filter_range is not None:
        text = req.filter_range.strip()
        ws.auto_filter.ref = a1(parse_range(text)) if text and text.lower() != "none" else None
        done.append(f"autofilter {ws.auto_filter.ref or 'removed'}")
    if req.gridlines is not None:
        ws.sheet_view.showGridLines = req.gridlines
        done.append(f"gridlines {'on' if req.gridlines else 'off'}")
    if req.zoom is not None:
        ws.sheet_view.zoomScale = req.zoom
        done.append(f"zoom {req.zoom}%")
    return done


def _sizes(ws: Any, widths: list[tuple[str, float]], heights: list[tuple[str, float]]) -> list[str]:
    done: list[str] = []
    for span, width in widths:
        axis, first, count = parse_span(span)
        if axis != "cols":
            raise SlimError(
                ErrorCode.BAD_ARGUMENT, f"widths take columns, got '{span}'", hint="e.g. 'B:D=12'"
            )
        for col in range(first, first + count):
            ws.column_dimensions[index_to_col(col - 1)].width = width
        done.append(f"width {span}={width:g}")
    for span, height in heights:
        axis, first, count = parse_span(span)
        if axis != "rows":
            raise SlimError(
                ErrorCode.BAD_ARGUMENT, f"heights take rows, got '{span}'", hint="e.g. '2:5=18'"
            )
        for row in range(first, first + count):
            ws.row_dimensions[row].height = height
        done.append(f"height {span}={height:g}")
    return done


def _visibility(ws: Any, spans: list[str], hidden: bool) -> list[str]:
    for span in spans:
        axis, first, count = parse_span(span)
        for index in range(first, first + count):
            if axis == "rows":
                ws.row_dimensions[index].hidden = hidden
            else:
                ws.column_dimensions[index_to_col(index - 1)].hidden = hidden
    return [f"{'hid' if hidden else 'showed'} {', '.join(spans)}"] if spans else []


def _printing(ws: Any, req: LayoutRequest) -> list[str]:
    done: list[str] = []
    if req.print_area is not None:
        ws.print_area = a1(parse_range(req.print_area))
        done.append(f"print area {req.print_area}")
    if req.landscape is not None:
        ws.page_setup.orientation = "landscape" if req.landscape else "portrait"
        done.append(ws.page_setup.orientation)
    if req.fit_width is not None:
        ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=req.fit_width)
        ws.page_setup.fitToWidth, ws.page_setup.fitToHeight = (
            (1, 0) if req.fit_width else (None, None)
        )
        done.append("fit to page width" if req.fit_width else "fit to width off")
    if req.protect is not None:
        ws.protection.sheet = req.protect
        done.append("sheet protected (no password)" if req.protect else "sheet unprotected")
    return done


def _workbook(wb: Any, req: LayoutRequest) -> list[str]:
    """Workbook-wide settings: calculation mode and structure protection (no password)."""
    done: list[str] = []
    if req.calc_mode is not None:
        mode = choice(req.calc_mode, tuple(CALC_MODES), "calc_mode")
        if wb.calculation is None:
            wb.calculation = CalcProperties()
        wb.calculation.calcMode = CALC_MODES[mode]
        done.append(f"calculation {mode}")
    if req.protect_structure is not None:
        if wb.security is None:
            wb.security = WorkbookProtection()
        wb.security.lockStructure = req.protect_structure
        done.append(
            "workbook structure protected (no password)"
            if req.protect_structure
            else "workbook structure unprotected"
        )
    return done
