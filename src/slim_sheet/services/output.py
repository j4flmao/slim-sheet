"""Shared output handling for tools that produce files (convert, clean, merge)."""

from __future__ import annotations

from pathlib import Path

from slim_sheet.adapters import journal
from slim_sheet.adapters.sandbox import resolve_output
from slim_sheet.adapters.writers.export import check_suffix
from slim_sheet.config import Settings
from slim_sheet.core.errors import ErrorCode, SlimError


def output_path(raw: str, settings: Settings, overwrite: bool, base: Path | None = None) -> Path:
    """`base`: folder of the source file; a relative `raw` is written next to it."""
    target = resolve_output(raw, settings, base)
    check_suffix(target)
    if target.exists() and not overwrite:
        raise SlimError(
            ErrorCode.WRITE_CONFLICT,
            f"'{target.name}' already exists",
            hint="Pass overwrite=true (the old file is kept in the journal) or another out_path.",
        )
    return target


def default_output(source: Path, tag: str, suffix: str | None = None) -> str:
    """'sales.xlsx' + 'clean' -> 'sales.clean.xlsx' next to the source."""
    return str(source.with_name(f"{source.stem}.{tag}{suffix or source.suffix}"))


def journal_before_write(target: Path, note: str) -> str | None:
    backup = journal.record(target, note)
    return backup.name if backup else None
