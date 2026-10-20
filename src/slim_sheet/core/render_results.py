"""Pure renderers for query / find / profile / write / formula results."""

from __future__ import annotations

from collections.abc import Sequence

from slim_sheet.core.a1 import format_range, index_to_col
from slim_sheet.core.budget import cell_text
from slim_sheet.core.formula import FormulaReport
from slim_sheet.core.models import (
    CellValue,
    FindResult,
    ProfileResult,
    QueryResult,
    WriteResult,
)

FIND_VALUE_CHARS = 80
ARG_CHARS = 40


def _md(text: str) -> str:
    return text.replace("|", "\\|")


def query_header(columns: Sequence[str]) -> str:
    head = "| " + " | ".join(_md(c) for c in columns) + " |"
    return head + "\n|" + "|".join("---" for _ in columns) + "|"


def query_row(row: Sequence[CellValue], max_chars: int) -> str:
    return "| " + " | ".join(_md(cell_text(v, max_chars)) for v in row) + " |"


def render_query(result: QueryResult, max_chars: int) -> str:
    if not result.columns:
        return "Query returned no columns."
    first, last = result.offset + 1, result.offset + len(result.rows)
    span = (
        f"rows {first}-{last} of {result.total_rows}"
        if result.rows
        else f"0 of {result.total_rows} rows"
    )
    lines = [f"## query: {span}", query_header(result.columns)]
    lines += [query_row(row, max_chars) for row in result.rows]
    if result.next_offset is not None:
        lines.append(
            f"next: offset={result.next_offset} (same sql; add ORDER BY for stable paging)"
        )
    return "\n".join(lines)


def render_find(result: FindResult) -> str:
    if not result.hits:
        scanned = result.cells_scanned
        return f"No cells in {result.path} contain '{result.query}' ({scanned} cells scanned)."
    more = " (more exist; narrow by sheet or raise max_results)" if result.truncated else ""
    lines = [f"## {len(result.hits)} match(es) for '{result.query}' in {result.path}{more}"]
    for hit in result.hits:
        address = f"{hit.sheet}!{index_to_col(hit.col)}{hit.row + 1}"
        lines.append(f"- {address}: {cell_text(hit.value, FIND_VALUE_CHARS)}")
    return "\n".join(lines)


def render_profile(result: ProfileResult) -> str:
    lines = [
        f"## profile {result.path} / {result.table} ({result.source}, {result.rows} rows)",
        "| column | type | nulls% | distinct | min | max | mean | median | top values |",
        "|---|---|--:|--:|---|---|---|---|---|",
    ]
    for c in result.columns:
        top = ", ".join(f"{v} ({n})" for v, n in c.top) or ("all unique" if c.count else "")
        cells = [
            c.name,
            c.dtype,
            f"{c.null_pct:g}",
            str(c.distinct),
            c.min,
            c.max,
            c.mean,
            c.median,
            top,
        ]
        lines.append("| " + " | ".join(_md(x or "") for x in cells) + " |")
    if result.omitted_columns:
        lines.append(
            f"{result.omitted_columns} more column(s) omitted; pass columns=[...] to choose."
        )
    return "\n".join(lines)


def render_write(result: WriteResult) -> str:
    where = f"{result.target} / {result.sheet}!{format_range(result.written)}"
    lines = [f"Wrote {result.cells} cell(s) to {where}."]
    if result.created_sheet:
        lines.append(f"Created sheet '{result.sheet}'.")
    if result.target != result.source:
        lines.append(
            f"Working copy: {result.target} ({result.source} untouched). "
            f"Read it with path='{result.target}'."
        )
    if result.backup:
        lines.append(f"Previous version saved as .slim-sheet/journal/{result.backup}.")
    lines += [f"warning: {w}" for w in result.warnings]
    return "\n".join(lines)


def render_formula(report: FormulaReport, address: str | None) -> str:
    source = f" ({address})" if address else ""
    lines = [f"## {report.formula}{source}"]
    if report.calls:
        lines.append("Steps (innermost first):")
    for index, call in enumerate(report.calls, start=1):
        args = ", ".join(cell_text(a, ARG_CHARS) for a in call.args)
        lines.append(f"{index}. {call.name}({args})")
        if call.info is not None:
            lines.append(f"   {call.name}({call.info.signature}): {call.info.summary}")
    lines.append("References: " + (", ".join(report.references) or "none"))
    lines.append("Errors: " + ("; ".join(report.errors) or "none"))
    if report.warnings:
        lines.append("Warnings: " + "; ".join(report.warnings))
    return "\n".join(lines)
