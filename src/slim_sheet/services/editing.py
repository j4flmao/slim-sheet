"""Shared edit session for every tool that modifies a workbook (PLAN §3.2, §8).

Target selection ("working copy" semantics, identical for all editing tools):
- overwrite=true            -> edit `path` in place
- out_path given            -> edit out_path (created from `path` on first edit)
- path already '*.slim.ext' -> edit it in place (it is our working copy)
- otherwise                 -> edit '<stem>.slim<ext>' next to `path`
Every replaced or created file is journaled for `undo`; saves are atomic.
"""

from __future__ import annotations

from collections.abc import Callable
from dataclasses import dataclass
from pathlib import Path
from typing import Any

from slim_sheet.adapters import crypto, journal
from slim_sheet.adapters.formats import sniff
from slim_sheet.adapters.limits import check_cells
from slim_sheet.adapters.sandbox import resolve_output, resolve_path
from slim_sheet.adapters.writers import dynamic_arrays, sparklines
from slim_sheet.adapters.writers.xlsx_writer import load_for_edit, lost_features, save_atomic
from slim_sheet.config import Settings
from slim_sheet.core.errors import ErrorCode, SlimError
from slim_sheet.core.models import FileFormat, OutputResult

SLIM_TAG = ".slim"
WRITABLE_SUFFIXES = {".xlsx", ".xlsm"}


@dataclass(frozen=True, slots=True)
class Change:
    summary: str
    warnings: tuple[str, ...] = ()
    sparklines: tuple[tuple[str, str], ...] = ()  # (sheet title, new sparklineGroup XML)
    dynamic: tuple[tuple[str, str], ...] = ()  # (sheet title, cell) written as spilling


def writable_source(path: str, settings: Settings) -> Path:
    source = resolve_path(path, settings)
    if crypto.is_encrypted(source):
        raise SlimError(
            ErrorCode.ENCRYPTED,
            f"'{source.name}' is password-protected; slim-sheet does not edit encrypted files",
            hint=(
                "Read it as usual, or make an unencrypted copy: "
                f"convert(path='{source.name}', out_path='copy.xlsx')"
            ),
        )
    if sniff(source) != FileFormat.XLSX or source.suffix.lower() not in WRITABLE_SUFFIXES:
        raise SlimError(
            ErrorCode.UNSUPPORTED_FORMAT,
            f"Editing supports .xlsx/.xlsm only, not '{source.name}'",
            hint=f"Convert first: convert(path='{source.name}', out_path='{source.stem}.xlsx')",
        )
    return source


def resolve_target(source: Path, settings: Settings, out_path: str | None, overwrite: bool) -> Path:
    if overwrite and not out_path:
        return source
    if out_path:
        target = resolve_output(out_path, settings, base=source.parent)
        if target.suffix.lower() != source.suffix.lower():
            raise SlimError(
                ErrorCode.BAD_ARGUMENT, f"out_path must keep the '{source.suffix}' extension"
            )
        if target == source and not overwrite:
            raise SlimError(
                ErrorCode.WRITE_CONFLICT,
                "out_path equals path; editing the original needs overwrite=true",
            )
        return target
    if source.stem.endswith(SLIM_TAG):
        return source
    return source.with_name(f"{source.stem}{SLIM_TAG}{source.suffix}")


def edit_workbook(
    path: str,
    settings: Settings,
    note: str,
    mutate: Callable[[Any], Change],
    out_path: str | None = None,
    overwrite: bool = False,
) -> OutputResult:
    """Load the right base, apply `mutate(workbook)`, journal, save atomically."""
    source = writable_source(path, settings)
    target = resolve_target(source, settings, out_path, overwrite)
    base = target if target.exists() else source
    check_cells(base, settings.max_cells)
    wb = load_for_edit(base)
    lost = lost_features(base, wb)
    warnings = [f"saved copy drops {f} (openpyxl limitation)" for f in lost]
    kept = sparklines.extract(base)  # openpyxl drops them on save; written back below
    spills = dynamic_arrays.extract(base)  # likewise the dynamic-array cell metadata
    change = mutate(wb)
    backup = journal.record(target, note)
    save_atomic(wb, target)
    groups = {title: list(found) for title, found in kept.items() if title in wb.sheetnames}
    for title, group in change.sparklines:
        groups.setdefault(title, []).append(group)
    sparklines.inject(target, groups)
    for title, cell in change.dynamic:
        spills.setdefault(title, set()).add(cell)
    dynamic_arrays.inject(target, {t: c for t, c in spills.items() if t in wb.sheetnames})
    return OutputResult(
        source.name,
        target.name,
        change.summary,
        backup.name if backup else None,
        (*change.warnings, *warnings),
    )


def get_sheet(wb: Any, name: str) -> Any:
    for title in wb.sheetnames:
        if title.casefold() == name.casefold():
            return wb[title]
    raise SlimError(
        ErrorCode.SHEET_NOT_FOUND,
        f"No sheet '{name}'",
        hint=f"Sheets: {', '.join(wb.sheetnames)}",
    )
