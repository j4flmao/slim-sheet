"""Atomic xlsx/xlsm cell writes via openpyxl (PLAN ADR-2)."""

from __future__ import annotations

import contextlib
import os
import re
import tempfile
import zipfile
from pathlib import Path
from typing import Any

import openpyxl

from slim_sheet.core.errors import ErrorCode, SlimError

# Package parts openpyxl never round-trips.
_FRAGILE_PARTS = {
    "xl/slicers/": "slicers",
    "xl/timelines/": "timelines",
    "xl/ctrlProps/": "form controls",
}
_CHART_PART = re.compile(r"^xl/charts/chart\d+\.xml$")
_MEDIA_PART = re.compile(r"^xl/media/")


def lost_features(path: Path, wb: Any) -> list[str]:
    """What saving `wb` (loaded from `path`) drops: charts/images openpyxl could not read
    back, plus parts it never keeps."""
    try:
        with zipfile.ZipFile(path) as archive:
            names = archive.namelist()
    except zipfile.BadZipFile:
        return []
    lost = sorted(
        {
            label
            for prefix, label in _FRAGILE_PARTS.items()
            if any(n.startswith(prefix) for n in names)
        }
    )
    charts = sum(1 for n in names if _CHART_PART.match(n)) - sum(
        len(ws._charts) for ws in wb.worksheets
    )
    images = sum(1 for n in names if _MEDIA_PART.match(n)) - sum(
        len(ws._images) for ws in wb.worksheets
    )
    if charts > 0:
        lost.insert(0, f"{charts} chart(s) openpyxl could not read")
    if images > 0:
        lost.insert(0, f"{images} image(s) (install Pillow to keep images)")
    return lost


def load_for_edit(base: Path) -> openpyxl.Workbook:
    """Full (non read-only) load that keeps VBA for macro-enabled files."""
    keep_vba = base.suffix.lower() in {".xlsm", ".xltm"}
    try:
        return openpyxl.load_workbook(base, keep_vba=keep_vba)
    except PermissionError as exc:
        raise _locked(base) from exc


def save_atomic(wb: openpyxl.Workbook, target: Path) -> None:
    fd, tmp_name = tempfile.mkstemp(prefix=".slim-", suffix=target.suffix, dir=target.parent)
    os.close(fd)
    tmp = Path(tmp_name)
    try:
        wb.save(tmp)
        tmp.replace(target)
    except PermissionError as exc:
        raise _locked(target) from exc
    finally:
        with contextlib.suppress(FileNotFoundError):
            tmp.unlink()


def _locked(path: Path) -> SlimError:
    return SlimError(
        ErrorCode.FILE_LOCKED,
        f"'{path.name}' is locked by another program",
        hint="Close it in Excel, or pass a different out_path.",
    )
