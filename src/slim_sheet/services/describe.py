"""`describe_cell` use-case: everything about one cell — value, formula, format, notes."""

from __future__ import annotations

from slim_sheet.adapters.xlsx_meta import cell_details
from slim_sheet.config import Settings
from slim_sheet.core.a1 import index_to_col, parse_range
from slim_sheet.core.errors import ErrorCode, SlimError
from slim_sheet.core.models import CellDescription, CellRange, FileFormat
from slim_sheet.core.refs import split_sheet
from slim_sheet.services.workbook import open_file, pick_sheet


def describe_cell(
    settings: Settings, path: str, cell: str, sheet: str | None = None
) -> CellDescription:
    sheet_part, ref = split_sheet(cell)
    rng = parse_range(ref)
    if rng.n_rows * rng.n_cols != 1:
        raise SlimError(
            ErrorCode.BAD_ARGUMENT,
            f"describe_cell takes one cell, got {ref}",
            hint="Use read for ranges, e.g. read(path, range='A1:D20').",
        )
    with open_file(path, settings) as handle:
        name = pick_sheet(handle, sheet_part or sheet)
        value = handle.reader.read(name, CellRange(rng.top, rng.left, rng.top, rng.left))[0][0]
        address = f"{name}!{index_to_col(rng.left)}{rng.top + 1}"
        if handle.format != FileFormat.XLSX:
            note = (
                "note",
                f"{handle.format.value} files expose values only (no formulas or styles)",
            )
            return CellDescription(address, value, (note,))
        facts = cell_details(handle.path, name, rng.top, rng.left)
        return CellDescription(address, None if value == "" else value, tuple(facts))
