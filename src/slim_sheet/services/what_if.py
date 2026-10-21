"""`what_if` use-case: change inputs in memory and see how outputs move (PLAN §17.2).

Read-only: the workbook is never written; apply a result with `write` afterwards.
- op=delta     inputs=["B3=0.08"], outputs=["Model!B20"]         -> base vs new per output
- op=goal_seek inputs=["B3"], outputs=["B20"], target="1000000"   -> the B3 that makes B20 = target
- op=grid      inputs=["B3=5%,6%,7%"] (one or two), outputs=["B20"] -> a sensitivity table
Every formula between inputs and outputs is recomputed by the offline engine (extra
`formula`); functions it does not implement fail loudly instead of being guessed.
"""

from __future__ import annotations

from collections.abc import Callable
from dataclasses import dataclass

from slim_sheet.config import Settings
from slim_sheet.core.budget import cell_text
from slim_sheet.core.errors import ErrorCode, SlimError
from slim_sheet.core.models import CellValue, WhatIfResult
from slim_sheet.core.styles import choice
from slim_sheet.services.evaluate import evaluate_cells

OPS = ("delta", "goal_seek", "grid")
MAX_GRID = 15  # values per axis
MAX_ITER = 60
NOT_WRITTEN = "Nothing was written; apply a result with write(...) if you want to keep it."


@dataclass(frozen=True, slots=True)
class WhatIfRequest:
    path: str
    outputs: list[str]
    inputs: list[str]
    op: str = "delta"
    target: str | None = None
    sheet: str | None = None


def what_if(req: WhatIfRequest, settings: Settings) -> WhatIfResult:
    op = choice(req.op, OPS, "op")
    if not req.outputs:
        raise SlimError(
            ErrorCode.BAD_ARGUMENT, "outputs is empty", hint="e.g. outputs=['Model!B20']"
        )
    if not req.inputs:
        raise SlimError(
            ErrorCode.BAD_ARGUMENT, "inputs is empty", hint="e.g. inputs=['Inputs!B3=0.08']"
        )
    if op == "delta":
        return _delta(req, settings)
    if op == "goal_seek":
        return _goal_seek(req, settings)
    return _grid(req, settings)


def parse_value(text: str) -> CellValue:
    """'8%' -> 0.08, '1,200' -> 1200, 'TRUE' -> True, other text stays text."""
    raw = text.strip()
    if raw.upper() in ("TRUE", "FALSE"):
        return raw.upper() == "TRUE"
    try:
        if raw.endswith("%"):
            return float(raw[:-1].replace(",", "")) / 100
        number = float(raw.replace(",", ""))
        return int(number) if number.is_integer() and "." not in raw else number
    except ValueError:
        return raw


def _assignment(item: str) -> tuple[str, str]:
    cell, eq, value = item.partition("=")
    if not eq or not cell.strip() or not value.strip():
        raise SlimError(
            ErrorCode.BAD_ARGUMENT,
            f"input '{item}' is not 'Cell=value'",
            hint="e.g. 'Inputs!B3=0.08'",
        )
    return cell.strip(), value.strip()


def _delta(req: WhatIfRequest, settings: Settings) -> WhatIfResult:
    changes = {cell: parse_value(value) for cell, value in map(_assignment, req.inputs)}
    base = evaluate_cells(settings, req.path, req.outputs, req.sheet, {})
    new = evaluate_cells(settings, req.path, req.outputs, req.sheet, changes)
    shown = ", ".join(f"{c} = {cell_text(v, 30)}" for c, v in changes.items())
    rows = tuple(
        (cell, before, after, *_change(before, after))
        for cell, before, after in zip(req.outputs, base, new, strict=True)
    )
    columns = ("output", "base", "new", "change", "change %")
    return WhatIfResult("delta", f"what_if delta: {shown}", columns, rows, (NOT_WRITTEN,))


def _change(before: CellValue, after: CellValue) -> tuple[CellValue, CellValue]:
    if not all(isinstance(v, int | float) and not isinstance(v, bool) for v in (before, after)):
        return "n/a", "n/a"
    diff = float(after) - float(before)  # type: ignore[arg-type]
    pct = f"{diff / float(before):+.1%}" if before else "n/a"  # type: ignore[arg-type]
    return f"{diff:+,.6g}", pct


def _goal_seek(req: WhatIfRequest, settings: Settings) -> WhatIfResult:
    if req.target is None:
        raise SlimError(
            ErrorCode.BAD_ARGUMENT, "goal_seek needs target", hint="e.g. target='1000000'"
        )
    goal = parse_value(req.target)
    if not isinstance(goal, int | float) or isinstance(goal, bool):
        raise SlimError(ErrorCode.BAD_ARGUMENT, f"target '{req.target}' must be a number")
    changing, output = req.inputs[0].partition("=")[0].strip(), req.outputs[0]
    start = evaluate_cells(settings, req.path, [changing], req.sheet, {})[0]
    x0 = float(start) if isinstance(start, int | float) else 0.0

    def gap(x: float) -> float:
        value = evaluate_cells(settings, req.path, [output], req.sheet, {changing: x})[0]
        if not isinstance(value, int | float) or isinstance(value, bool):
            raise SlimError(ErrorCode.BAD_ARGUMENT, f"{output} is not numeric for {changing} = {x}")
        return float(value) - float(goal)

    x, residual = secant(gap, x0, x0 * 1.1 if x0 else 1.0)
    solved = abs(residual) <= max(1e-6, abs(float(goal)) * 1e-6)
    title = f"goal seek ({'solved' if solved else 'closest found'})"
    row = (changing, x, output, float(goal) + residual, float(goal))
    columns = ("change", "to", "so that", "equals", "target")
    note = (
        "Nothing was written; apply with "
        f"write(path, sheet, range='{changing}', values=[[{x:.6g}]])."
    )
    return WhatIfResult("goal_seek", title, columns, (row,), (note,))


def secant(f: Callable[[float], float], x0: float, x1: float) -> tuple[float, float]:
    """Root of f near x0..x1 -> (x, f(x)); returns the best point found if it does not converge."""
    y0, y1 = f(x0), f(x1)
    best = (x0, y0) if abs(y0) < abs(y1) else (x1, y1)
    for _ in range(MAX_ITER):
        if y1 == y0:
            break
        x2 = x1 - y1 * (x1 - x0) / (y1 - y0)
        y2 = f(x2)
        if abs(y2) < abs(best[1]):
            best = (x2, y2)
        if abs(y2) <= 1e-9 * max(1.0, abs(x2)):
            break
        x0, y0, x1, y1 = x1, y1, x2, y2
    return best


def _grid(req: WhatIfRequest, settings: Settings) -> WhatIfResult:
    axes = [_axis(item) for item in req.inputs[:2]]
    output = req.outputs[0]
    rows_cell, row_values = axes[0]

    def value(changes: dict[str, CellValue]) -> CellValue:
        return evaluate_cells(settings, req.path, [output], req.sheet, changes)[0]

    if len(axes) == 1:
        rows = tuple((rv, value({rows_cell: rv})) for rv in row_values)
        return WhatIfResult(
            "grid", f"sensitivity of {output} to {rows_cell}", (rows_cell, output), rows
        )
    cols_cell, col_values = axes[1]
    grid = tuple(
        (rv, *(value({rows_cell: rv, cols_cell: cv}) for cv in col_values)) for rv in row_values
    )
    columns = (f"{rows_cell} \\ {cols_cell}", *(cell_text(v, 20) for v in col_values))
    title = f"sensitivity of {output}: rows {rows_cell}, columns {cols_cell}"
    return WhatIfResult("grid", title, columns, grid)


def _axis(item: str) -> tuple[str, list[CellValue]]:
    cell, values = _assignment(item)
    parsed = [parse_value(v) for v in values.split(",") if v.strip()]
    if not 1 <= len(parsed) <= MAX_GRID:
        raise SlimError(
            ErrorCode.TOO_LARGE, f"use 1..{MAX_GRID} values per input, got {len(parsed)}"
        )
    return cell, parsed
