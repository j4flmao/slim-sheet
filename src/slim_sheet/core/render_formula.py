"""Pure renderers for the Phase 3 formula tools: validate_formula, evaluate, describe_cell."""

from __future__ import annotations

from slim_sheet.core.budget import cell_text
from slim_sheet.core.models import CellDescription, CellValue, EvalResult
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
