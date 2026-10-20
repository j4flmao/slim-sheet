"""`diff` use-case: what changed between two files (or a file and its .slim working copy).

Without `key`: cell-by-cell per sheet (values; plus formulas when both files are xlsx),
and sheets added/removed.
With `key`: row-level comparison of one table by a key column.
"""

from __future__ import annotations

from slim_sheet.adapters.xlsx_meta import sheet_formulas
from slim_sheet.config import Settings
from slim_sheet.core.a1 import index_to_col
from slim_sheet.core.diff import diff_formulas, diff_grids, diff_keyed
from slim_sheet.core.models import CellRange, CellValue, DiffResult, FileFormat
from slim_sheet.services.tables import resolve_table
from slim_sheet.services.workbook import OpenFile, open_file, pick_sheet

MAX_RESULTS_CAP = 500


def diff(
    path: str,
    other_path: str,
    settings: Settings,
    sheet: str | None = None,
    key: str | None = None,
    max_results: int = 50,
) -> DiffResult:
    limit = max(1, min(max_results, MAX_RESULTS_CAP))
    with open_file(path, settings) as left, open_file(other_path, settings) as right:
        if key:
            return _keyed(left, right, sheet, key, limit)
        return _cells(left, right, sheet, limit)


def _cells(left: OpenFile, right: OpenFile, sheet: str | None, limit: int) -> DiffResult:
    names_a = [n for n, _ in left.reader.sheets()]
    names_b = [n for n, _ in right.reader.sheets()]
    if sheet:
        common = [(pick_sheet(left, sheet), pick_sheet(right, sheet))]
        summary: list[str] = []
    else:
        folded_b = {n.casefold(): n for n in names_b}
        common = [(n, folded_b[n.casefold()]) for n in names_a if n.casefold() in folded_b]
        added = [n for n in names_b if n.casefold() not in {a.casefold() for a in names_a}]
        removed = [n for n in names_a if n.casefold() not in folded_b]
        summary = [f"sheets added: {', '.join(added)}"] if added else []
        summary += [f"sheets removed: {', '.join(removed)}"] if removed else []
    changes: list[tuple[str, CellValue, CellValue]] = []
    total = _value_changes(left, right, common, limit, changes, summary)
    if left.format == right.format == FileFormat.XLSX:
        total += _formula_changes(left, right, common, limit, changes, summary)
    if not summary:
        summary = ["no differences in cell values or formulas"]
    return DiffResult(left.display, right.display, "cells", tuple(summary), tuple(changes), total)


Changes = list[tuple[str, CellValue, CellValue]]


def _value_changes(
    left: OpenFile,
    right: OpenFile,
    common: list[tuple[str, str]],
    limit: int,
    changes: Changes,
    summary: list[str],
) -> int:
    total = 0
    for name_a, name_b in common:
        count = 0
        for d in diff_grids(_grid(left, name_a), _grid(right, name_b)):
            count += 1
            if len(changes) < limit:
                changes.append((f"{name_a}!{index_to_col(d.col)}{d.row + 1}", d.before, d.after))
        if count:
            summary.append(f"{name_a}: {count} changed value(s)")
        total += count
    return total


def _formula_changes(
    left: OpenFile,
    right: OpenFile,
    common: list[tuple[str, str]],
    limit: int,
    changes: Changes,
    summary: list[str],
) -> int:
    formulas_a, formulas_b = sheet_formulas(left.path), sheet_formulas(right.path)
    total = 0
    for name_a, name_b in common:
        found = diff_formulas(formulas_a.get(name_a, {}), formulas_b.get(name_b, {}))
        for row, col, before, after in found:
            if len(changes) < limit:
                label = f"{name_a}!{index_to_col(col)}{row + 1} (formula)"
                changes.append((label, before or "(no formula)", after or "(no formula)"))
        if found:
            summary.append(f"{name_a}: {len(found)} changed formula(s)")
        total += len(found)
    return total


def _grid(handle: OpenFile, sheet: str) -> list[list[CellValue]]:
    used = handle.reader.used_range(sheet)
    if used is None:
        return []
    return handle.reader.read(sheet, CellRange(0, 0, used.bottom, used.right))


def _keyed(left: OpenFile, right: OpenFile, sheet: str | None, key: str, limit: int) -> DiffResult:
    ta, tb = resolve_table(left, sheet), resolve_table(right, sheet)
    result = diff_keyed(ta.columns, ta.body, tb.columns, tb.body, key)
    summary = [
        f"{len(result.added)} row(s) added, {len(result.removed)} removed, "
        f"{len({k for k, *_ in result.changed})} changed (by {key})"
    ]
    if result.added:
        summary.append("added keys: " + ", ".join(result.added[:20]))
    if result.removed:
        summary.append("removed keys: " + ", ".join(result.removed[:20]))
    if result.columns_added or result.columns_removed:
        summary.append(
            f"columns added: {', '.join(result.columns_added) or '-'}; "
            f"removed: {', '.join(result.columns_removed) or '-'}"
        )
    if result.duplicate_keys:
        summary.append(f"warning: {result.duplicate_keys} duplicate key row(s) ignored")
    changes = tuple((f"{key}={k} / {col}", a, b) for k, col, a, b in result.changed[:limit])
    total = len(result.added) + len(result.removed) + len(result.changed)
    return DiffResult(left.display, right.display, f"key:{key}", tuple(summary), changes, total)
