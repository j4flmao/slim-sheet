"""Pure renderers for the Phase 2 data tools: clean, anomalies, merge, diff, convert, undo."""

from __future__ import annotations

from slim_sheet.core.budget import cell_text
from slim_sheet.core.models import (
    AnomalyResult,
    CellValue,
    CleanResult,
    DiffResult,
    MergeResult,
    OutputResult,
    PivotResult,
)
from slim_sheet.core.render_results import render_query

VALUE_CHARS = 60


def _v(value: CellValue) -> str:
    return "(empty)" if value is None or value == "" else cell_text(value, VALUE_CHARS)


def render_output(result: OutputResult) -> str:
    lines = [f"Wrote {result.target} from {result.source}: {result.summary}."]
    if result.target != result.source and result.source.endswith((".xlsx", ".xlsm")):
        lines.append(f"{result.source} is untouched; continue with path='{result.target}'.")
    if result.backup:
        lines.append(f"Previous {result.target} kept as .slim-sheet/journal/{result.backup}.")
    lines += [f"note: {w}" for w in result.warnings]
    lines.append(f"Revert with undo(path='{result.target}').")
    return "\n".join(lines)


def render_clean(result: CleanResult) -> str:
    mode = "applied" if result.output else "PREVIEW (nothing written; call again with apply=true)"
    lines = [
        f"## clean {result.path} / {result.table} ({result.source}) - {mode}",
        f"rows: {result.rows_before} -> {result.rows_after}",
    ]
    if result.counts:
        lines.append("changes: " + ", ".join(f"{op}={n}" for op, n in result.counts))
    else:
        lines.append("changes: none - the table is already clean for these ops")
    lines += [f"- {addr}: {_v(a)} -> {_v(b)} [{op}]" for addr, a, b, op in result.samples]
    lines += [f"note: {n}" for n in result.notes]
    if result.output:
        lines.append(render_output(result.output))
    return "\n".join(lines)


def render_anomalies(result: AnomalyResult) -> str:
    total = sum(n for _, n in result.counts)
    head = f"## anomalies {result.path} / {result.table} ({result.source}, {result.rows} rows)"
    if not total:
        return f"{head}\nNo anomalies found."
    lines = [head, "found: " + ", ".join(f"{kind}={n}" for kind, n in result.counts)]
    lines += [f"- {f.address} [{f.kind}] {_v(f.value)}: {f.detail}" for f in result.findings]
    if total > len(result.findings):
        lines.append(
            f"{total - len(result.findings)} more; raise max_results or pass columns=[...]."
        )
    return "\n".join(lines)


def render_merge(result: MergeResult, max_chars: int) -> str:
    lines = [
        f"## merge ({result.how}) {result.left} + {result.right} on {', '.join(result.keys)}",
        f"left rows {result.left_rows} ({result.left_unmatched} without a match), "
        f"right rows {result.right_rows} ({result.right_unmatched} without a match)",
    ]
    lines += [f"warning: {w}" for w in result.warnings]
    preview = render_query(result.preview, max_chars).replace("## query:", "preview:", 1)
    lines += [line for line in preview.splitlines() if not line.startswith("next:")]
    if result.output:
        lines.append(render_output(result.output))
    elif result.preview.next_offset is not None:
        lines.append("Pass out_path='merged.xlsx' to save the full result.")
    return "\n".join(lines)


def render_diff(result: DiffResult) -> str:
    lines = [
        f"## diff {result.left} -> {result.right} ({result.mode}): {result.total_changes} change(s)"
    ]
    lines += [f"- {s}" for s in result.summary]
    lines += [f"{loc}: {_v(a)} -> {_v(b)}" for loc, a, b in result.changes]
    shown = len(result.changes)
    if result.mode == "cells" and result.total_changes > shown:
        lines.append(f"{result.total_changes - shown} more; pass sheet=... or raise max_results.")
    return "\n".join(lines)


def render_undo(target: str, note: str, restored_from: str | None, kept_as: str) -> str:
    action = (
        "restored the previous version"
        if restored_from
        else "removed it (that write had created it)"
    )
    return (
        f"Undid '{note}' on {target}: {action}. "
        f"The undone state is kept as .slim-sheet/journal/{kept_as}."
    )


def render_pivot(result: PivotResult, max_chars: int) -> str:
    lines = [f"## pivot of {result.source}", f"sql: {result.sql}"]
    table = render_query(result.preview, max_chars).replace("## query:", "result:", 1)
    lines += [line for line in table.splitlines() if not line.startswith("next:")]
    if result.output:
        lines.append(render_output(result.output))
    elif result.preview.next_offset is not None:
        lines.append(
            "Truncated; save with out_path='pivot.xlsx' or to_sheet='Pivot', "
            "or page through the sql with query."
        )
    return "\n".join(lines)


def render_names(path: str, names: list[tuple[str, str]]) -> str:
    if not names:
        return f"{path} has no defined names."
    return "\n".join(
        [f"## defined names in {path}", *(f"- {name} = {target}" for name, target in names)]
    )
