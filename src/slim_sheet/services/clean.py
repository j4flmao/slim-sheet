"""`clean` use-case: preview by default, write a cleaned copy with apply=true (PLAN §9)."""

from __future__ import annotations

from pathlib import Path

from slim_sheet.adapters.engine.arrow import to_arrow
from slim_sheet.adapters.writers.export import export_table
from slim_sheet.config import Settings
from slim_sheet.core.clean import CleanOutcome, check_ops, clean_table
from slim_sheet.core.models import CleanResult, OutputResult
from slim_sheet.services.output import default_output, journal_before_write, output_path
from slim_sheet.services.tables import resolve_table
from slim_sheet.services.workbook import open_file


def clean(
    path: str,
    ops: list[str],
    settings: Settings,
    sheet: str | None = None,
    keys: list[str] | None = None,
    apply: bool = False,
    out_path: str | None = None,
    overwrite: bool = False,
) -> CleanResult:
    check_ops(ops)
    with open_file(path, settings) as handle:
        table = resolve_table(handle, sheet)
        outcome = clean_table(table.columns, table.body, ops, keys)
        output = None
        if apply:
            fallback_suffix = ".xlsx" if handle.format.is_workbook else None
            raw = out_path or default_output(handle.path, "clean", fallback_suffix)
            output = _write(handle.display, raw, outcome, table.name, settings, overwrite)
        samples = tuple(
            (table.address(s.row, s.col), s.before, s.after, s.op) for s in outcome.samples
        )
        return CleanResult(
            path=handle.display,
            table=table.name,
            source=table.source,
            rows_before=len(table.body),
            rows_after=len(outcome.rows),
            counts=tuple(outcome.counts.items()),
            samples=samples,
            notes=tuple(outcome.notes),
            output=output,
        )


def _write(
    source: str, raw: str, outcome: CleanOutcome, table: str, settings: Settings, overwrite: bool
) -> OutputResult:
    target: Path = output_path(raw, settings, overwrite)
    backup = journal_before_write(target, f"clean {source} / {table}")
    export_table(to_arrow(outcome.columns, outcome.rows), target, sheet=table)
    summary = f"{len(outcome.rows)} cleaned rows of table {table}"
    return OutputResult(source, target.name, summary, backup)
