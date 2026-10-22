"""Offline formula evaluation via the optional `formulas` library (extra `formula`)."""

from __future__ import annotations

import datetime as dt
import warnings
from dataclasses import dataclass
from functools import cache
from typing import Any

from slim_sheet.core.errors import ErrorCode, SlimError
from slim_sheet.core.models import CellValue

# `formulas` lazily imports submodules containing `x is 1.0`-style code; keep stderr clean.
warnings.filterwarnings("ignore", message=r".*with .* literal", category=SyntaxWarning)

EXCEL_EPOCH = dt.datetime(1899, 12, 30)
INSTALL_HINT = (
    "Install the formula extra: `uv sync --extra formula`, or run the server with "
    "`uvx --from 'slim-sheet[formula]' slim-sheet mcp`."
)


@cache
def _lib() -> Any:
    try:
        with warnings.catch_warnings():  # the library emits SyntaxWarnings on import
            warnings.simplefilter("ignore")
            import formulas
    except ImportError as exc:
        raise SlimError(
            ErrorCode.MISSING_EXTRA, "evaluate needs the 'formulas' package", hint=INSTALL_HINT
        ) from exc
    _keep_blank_in_index()
    return formulas


def _keep_blank_in_index() -> None:
    """Excel's INDEX returns a reference, so INDEX(...)="" is TRUE on a blank cell; the library
    turns blanks into 0 inside INDEX. Blanks are hidden from its helper and restored after."""
    import numpy as np
    import schedula
    from formulas.functions import look

    original = look._index
    if getattr(original, "slim_patched", False):
        return
    marker = object()
    hide = np.frompyfunc(lambda v: marker if v is schedula.EMPTY else v, 1, 1)

    def _index(arrays: Any, *args: Any) -> Any:
        masked = [
            hide(a).astype(object) if isinstance(a, np.ndarray) and a.dtype == object else a
            for a in arrays
        ]
        value = original(masked, *args)
        return schedula.EMPTY if value is marker else value

    _index.slim_patched = True  # type: ignore[attr-defined]
    look._index = _index


@cache
def supported_functions() -> frozenset[str]:
    from formulas.functions import get_functions

    _lib()
    return frozenset(name.upper() for name in get_functions())


@dataclass(frozen=True, slots=True)
class Compiled:
    func: Any
    inputs: tuple[str, ...]  # upper-case references as the engine names them, e.g. 'SHEET2!A1:B3'


def compile_formula(text: str) -> Compiled:
    lib = _lib()
    try:
        func = lib.Parser().ast(text)[1].compile()
    except Exception as exc:
        raise SlimError(ErrorCode.FORMULA_SYNTAX, f"Cannot compile {text}: {exc}") from exc
    return Compiled(func, tuple(func.inputs))


def empty() -> Any:
    import schedula

    _lib()
    return schedula.EMPTY


def to_engine(value: CellValue) -> Any:
    """Python cell value -> what Excel stores (dates become serial numbers)."""
    if isinstance(value, dt.datetime):
        return (value - EXCEL_EPOCH).total_seconds() / 86_400
    if isinstance(value, dt.date):
        return float((value - EXCEL_EPOCH.date()).days)
    if isinstance(value, dt.time):
        return (value.hour * 3600 + value.minute * 60 + value.second) / 86_400
    if isinstance(value, dt.timedelta):
        return value.total_seconds() / 86_400
    return value


def run(compiled: Compiled, args: list[Any]) -> Any:
    try:
        return compiled.func(*args)
    except Exception as exc:
        message = str(exc).splitlines()[0] if str(exc) else type(exc).__name__
        raise SlimError(ErrorCode.FORMULA_SYNTAX, f"Evaluation failed: {message}") from exc


def from_engine(result: Any) -> CellValue | tuple[tuple[CellValue, ...], ...]:
    """Engine result -> scalar, or a 2-D tuple for spilled arrays. Errors become '#DIV/0!' etc."""
    shape = getattr(result, "shape", None)
    if shape is not None and len(shape) == 2 and (shape[0] > 1 or shape[1] > 1):
        return tuple(tuple(_scalar(v) for v in row) for row in result.tolist())
    if shape is not None:
        flat = result.ravel().tolist() if hasattr(result, "ravel") else [result]
        return _cell_result(flat[0]) if flat else None
    return _cell_result(result)


def _cell_result(value: Any) -> CellValue:
    """A formula that ends on a blank cell shows 0 in Excel (=B2 with B2 empty)."""
    import schedula

    return 0 if value is schedula.EMPTY else _scalar(value)


def _scalar(value: Any) -> CellValue:
    if type(value).__module__ == "numpy" and hasattr(value, "item"):
        value = value.item()  # np.float64 subclasses float and would leak into replies
    if value is None or isinstance(value, bool):
        return value
    if isinstance(value, str):
        return str(value)  # XlError subclasses str; normalise to plain '#N/A'-style text
    if isinstance(value, int | float):
        return int(value) if isinstance(value, float) and value.is_integer() else value
    text = str(value)
    return None if text in ("", "empty") else text
