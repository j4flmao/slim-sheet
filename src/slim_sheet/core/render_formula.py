"""Pure renderers for the Phase 3 formula tools: validate_formula, evaluate, describe_cell."""

from __future__ import annotations

from slim_sheet.core.budget import UNTRUSTED, cell_text
from slim_sheet.core.models import (
    CellDescription,
    CellValue,
    EvalResult,
    FolderListing,
    MacroListing,
    MacroSource,
    TraceResult,
    WhatIfResult,
)
from slim_sheet.core.validate import ValidationReport

VALUE_CHARS = 80
MAX_SPILL_ROWS = 20
MAX_SPILL_COLS = 10


def _v(value: CellValue) -> str:
    return "(empty)" if value is None or value == "" else cell_text(value, VALUE_CHARS)


def render_validation(report: ValidationReport, address: str | None) -> str:
    where = f" ({address})" if address else ""
    verdict = "OK" if report.ok else f"{len(report.errors)} problem(s)"
    lines = [f"## {report.formula}{where} - target Excel {report.target}: {verdict}"]
    lines += [f"error: {e}" for e in report.errors]
    lines += [f"warning: {w}" for w in report.warnings]
    lines += [f"fix: {f}" for f in report.fixes]
    return "\n".join(lines)


def render_eval(result: EvalResult) -> str:
    where = f"{result.address} " if result.address else ""
    head = f"## {where}{result.formula}".rstrip()
    if isinstance(result.value, tuple):
        lines = [
            head,
            f"result: array {len(result.value)}x{len(result.value[0]) if result.value else 0}",
        ]
        for row in result.value[:MAX_SPILL_ROWS]:
            lines.append("| " + " | ".join(_v(v) for v in row[:MAX_SPILL_COLS]) + " |")
    else:
        lines = [head, f"result: {_v(result.value)}"]
    if result.inputs:
        lines.append("inputs: " + "; ".join(f"{ref} {shown}" for ref, shown in result.inputs))
    if result.recomputed:
        lines.append("recomputed (no cached value): " + ", ".join(result.recomputed))
    lines += [f"note: {n}" for n in result.notes]
    return "\n".join(lines)


def render_cell(description: CellDescription) -> str:
    lines = [f"## {description.address}", f"value: {_v(description.value)}"]
    lines += [f"{label}: {text}" for label, text in description.facts]
    return "\n".join(lines)


def render_trace(result: TraceResult) -> str:
    title = f"{result.direction} of {result.address} (depth {result.depth})"
    if not result.rows:
        return f"{title}: none"
    lines = [f"{title}:"]
    lines += [
        f"{'  ' * (level - 1)}- {address}{(' ' + text) if text else ''}"
        for level, address, text in result.rows
    ]
    if result.truncated:
        lines.append(f"(stopped at {len(result.rows)} lines; lower depth or trace another cell)")
    return "\n".join(lines)


def render_what_if(result: WhatIfResult) -> str:
    lines = [f"## {result.title}"]
    if result.columns:
        lines.append("| " + " | ".join(result.columns) + " |")
        lines.append("|" + "---|" * len(result.columns))
        lines += ["| " + " | ".join(_num(v) for v in row) + " |" for row in result.rows]
    lines += list(result.notes)
    return "\n".join(lines)


def _num(value: CellValue) -> str:
    if isinstance(value, float):
        return f"{value:,.6g}" if abs(value) < 1e6 else f"{value:,.0f}"
    if isinstance(value, int) and not isinstance(value, bool):
        return f"{value:,}"
    return cell_text(value, 30)


def render_folder(listing: FolderListing) -> str:
    lines = [f"## {listing.total} spreadsheet(s) in {listing.folder}"]
    lines += [f"- {path} ({max(1, size // 1024)} KB)" for path, size in listing.items]
    if listing.total > len(listing.items):
        lines.append(
            f"{listing.total - len(listing.items)} more; pass a subfolder to narrow the list."
        )
    return "\n".join(lines)


def render_macros(listing: MacroListing) -> str:
    states = (
        ("signed", listing.signed),
        ("password-locked", listing.locked),
        ("from the internet: Excel blocks these macros", listing.internet),
    )
    flags = [label for label, on in states if on]
    head = f"## VBA project '{listing.project}' in {listing.path}" + (
        f" ({', '.join(flags)})" if flags else ""
    )
    lines = [
        head,
        "slim-sheet never runs macros. " + UNTRUSTED,
        "| module | kind | code lines | runs by itself |",
        "|---|---|--:|---|",
    ]
    lines += [
        f"| {m.name} | {m.kind} | {m.lines} | {', '.join(m.auto_exec) or ''} |"
        for m in listing.modules
    ]
    if listing.stomped:
        lines.append(
            f"warning: compiled code without source in {', '.join(listing.stomped)} (VBA stomping)"
        )
    if listing.macro_sheets:
        lines.append(f"warning: {listing.macro_sheets} Excel 4.0 macro sheet(s) also hold macros")
    lines.append(
        "Show code: read(path, sheet='#vba', range='<module>'); "
        "risks: anomalies(path, scope='macros')."
    )
    return "\n".join(lines)


def render_macro_source(source: MacroSource) -> str:
    last = source.first_line + len(source.lines) - 1
    lines = [
        f"## {source.path} / #vba / {source.module} ({source.kind}) - lines "
        f"{source.first_line}-{last} of {source.total_lines} {UNTRUSTED}",
        "```vba",
    ]
    lines += [f"{source.first_line + i:>4}  {text}" for i, text in enumerate(source.lines)]
    lines.append("```")
    if source.next_cursor:
        lines.append(f'next: cursor="{source.next_cursor}" (call read again with this cursor)')
    return "\n".join(lines)
