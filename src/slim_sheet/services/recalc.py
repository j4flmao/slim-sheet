"""`recalc` use-case: refresh every formula's cached value with headless LibreOffice."""

from __future__ import annotations

import shutil
from pathlib import Path

from slim_sheet.adapters import journal, libreoffice
from slim_sheet.adapters.readers import open_reader
from slim_sheet.adapters.vba import project_part, restore_project
from slim_sheet.adapters.xlsx_meta import sheet_formulas
from slim_sheet.config import Settings
from slim_sheet.core.models import CellRange, FileFormat, OutputResult
from slim_sheet.services.editing import resolve_target, writable_source

RESAVE_NOTE = (
    "LibreOffice re-saved the workbook; Excel-only features (slicers, Power Query, some chart "
    "styling) may change. Compare with diff if in doubt."
)


def recalc(
    path: str, settings: Settings, out_path: str | None = None, overwrite: bool = False
) -> OutputResult:
    source = writable_source(path, settings)
    soffice = libreoffice.find_soffice(settings.soffice)
    target = resolve_target(source, settings, out_path, overwrite)
    base = target if target.exists() else source
    original = project_part(base)  # LibreOffice may drop UserForms/buttons when re-saving VBA
    produced = libreoffice.recalculate(base, soffice)
    notes = [RESAVE_NOTE]
    try:
        if original is not None:
            outcome = restore_project(produced, *original)
            if outcome == "restored":
                notes.append(
                    "the VBA project was restored byte-for-byte after LibreOffice changed it"
                )
            elif outcome == "lost":
                notes.append(
                    "LibreOffice removed the VBA project; the previous file is in the journal"
                )
        backup = journal.record(target, "recalc")
        tmp = target.with_name(f".slim-recalc-{target.name}")
        shutil.copy2(produced, tmp)
        tmp.replace(target)
    finally:
        shutil.rmtree(produced.parent.parent, ignore_errors=True)
    filled, total = _cached_formula_values(target)
    summary = f"recalculated {total} formula(s); {filled} now have cached values"
    return OutputResult(
        source.name, target.name, summary, backup.name if backup else None, tuple(notes)
    )


def _cached_formula_values(path: Path) -> tuple[int, int]:
    formulas = sheet_formulas(path)
    filled = 0
    with open_reader(path, FileFormat.XLSX) as reader:
        for sheet, cells in formulas.items():
            for row, col in cells:
                value = reader.read(sheet, CellRange(row, col, row, col))[0][0]
                filled += int(value not in (None, ""))
    return filled, sum(len(c) for c in formulas.values())
