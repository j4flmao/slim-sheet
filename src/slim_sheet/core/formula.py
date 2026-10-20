"""Static formula analysis: calls (inside-out), references and issues (PLAN §10).

Built on openpyxl's pure tokenizer; we add paren balancing, arity and locale checks,
because the tokenizer itself accepts unbalanced input.
"""

from __future__ import annotations

import difflib
from dataclasses import dataclass, field

from openpyxl.formula.tokenizer import Token, Tokenizer, TokenizerError

from slim_sheet.core.errors import ErrorCode, SlimError
from slim_sheet.core.functions import (
    FUNCTIONS,
    MODERN_VERSIONS,
    VARIADIC,
    VOLATILE,
    FunctionInfo,
)

MAX_FORMULA_CHARS = 8192
_PREFIXES = ("_xlfn._xlws.", "_xlfn.", "_xlws.")


@dataclass(frozen=True, slots=True)
class FormulaCall:
    name: str
    args: tuple[str, ...]
    depth: int
    info: FunctionInfo | None


@dataclass(frozen=True, slots=True)
class FormulaReport:
    formula: str
    calls: tuple[FormulaCall, ...]
    references: tuple[str, ...]
    errors: tuple[str, ...]
    warnings: tuple[str, ...]


@dataclass
class _Frame:
    name: str | None  # None for a plain parenthesis or an array literal
    kind: str
    args: list[list[str]] = field(default_factory=lambda: [[]])


@dataclass
class _State:
    stack: list[_Frame] = field(default_factory=list)
    calls: list[FormulaCall] = field(default_factory=list)
    refs: list[str] = field(default_factory=list)
    errors: list[str] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)
    top: list[str] = field(default_factory=list)


def analyze(formula: str) -> FormulaReport:
    text = formula.strip()
    if not text:
        raise SlimError(
            ErrorCode.BAD_ARGUMENT, "formula is empty", hint="e.g. formula='=SUM(A1:A10)'"
        )
    if len(text) > MAX_FORMULA_CHARS:
        raise SlimError(ErrorCode.TOO_LARGE, f"formula longer than {MAX_FORMULA_CHARS} characters")
    state = _State()
    if not text.startswith("="):
        text = "=" + text
        state.warnings.append("added the leading '='")
    extra_close = _unmatched_close(text)
    if extra_close is not None:
        error = f"unmatched ')' at position {extra_close}"
        return FormulaReport(text, (), (), (error,), tuple(state.warnings))
    try:
        tokens = Tokenizer(text).items
    except (TokenizerError, IndexError, AssertionError) as exc:  # tokenizer has no error type
        error = f"cannot tokenize: {exc or type(exc).__name__}"
        return FormulaReport(text, (), (), (error,), tuple(state.warnings))
    for token in tokens:
        _consume(state, token)
    _finish(state, tokens)
    return FormulaReport(
        text, tuple(state.calls), tuple(state.refs), tuple(state.errors), tuple(state.warnings)
    )


def _emit(state: _State, piece: str) -> None:
    (state.stack[-1].args[-1] if state.stack else state.top).append(piece)


def _consume(state: _State, token: Token) -> None:
    kind, sub, value = token.type, token.subtype, str(token.value)
    if sub == Token.OPEN and kind in (Token.FUNC, Token.PAREN, "ARRAY"):
        name = _clean_name(value[:-1]) if kind == Token.FUNC else None
        state.stack.append(_Frame(name, kind))
    elif sub == Token.CLOSE and kind in (Token.FUNC, Token.PAREN, "ARRAY"):
        _close(state, value)
    elif kind == Token.SEP:
        _separator(state, value, sub)
    else:
        if kind == Token.OPERAND and sub == Token.RANGE and value not in state.refs:
            state.refs.append(value)
        _emit(state, value)


def _separator(state: _State, value: str, sub: str) -> None:
    frame = state.stack[-1] if state.stack else None
    if frame is not None and frame.kind == "ARRAY":
        frame.args[-1].append(value)
        return
    if sub == "ROW" or value == ";":
        state.errors.append("';' used between arguments; Excel (en-US) expects ','")
    if frame is None or frame.name is None:
        state.errors.append(f"'{value}' outside a function call")
        _emit(state, value)
        return
    frame.args.append([])


def _close(state: _State, value: str) -> None:
    if not state.stack:
        state.errors.append(f"unmatched '{value}'")
        return
    frame = state.stack.pop()
    inner = [" ".join(a).strip() for a in frame.args]
    if frame.kind == "ARRAY":
        _emit(state, "{" + "".join("".join(a) for a in frame.args) + "}")
        return
    if frame.name is None:
        _emit(state, "(" + ", ".join(inner) + ")")
        return
    args = () if inner == [""] else tuple(inner)
    info = FUNCTIONS.get(frame.name)
    state.calls.append(FormulaCall(frame.name, args, len(state.stack), info))
    _check_call(state, frame.name, len(args), info)
    _emit(state, f"{frame.name}({', '.join(args)})")


def _check_call(state: _State, name: str, n_args: int, info: FunctionInfo | None) -> None:
    if info is None:
        close = difflib.get_close_matches(name, FUNCTIONS, n=1, cutoff=0.75)
        hint = f"; did you mean {close[0]}?" if close else ""
        state.warnings.append(f"{name} is not in the catalog{hint}")
        return
    too_many = info.max_args != VARIADIC and n_args > info.max_args
    if n_args < info.min_args or too_many:
        state.errors.append(
            f"{name} takes {_arity(info)} argument(s), got {n_args}: {name}({info.signature})"
        )
    if info.since in MODERN_VERSIONS:
        state.warnings.append(f"{name} needs Excel {info.since}+; older Excel shows #NAME?")
    if name in VOLATILE:
        state.warnings.append(f"{name} is volatile: it recalculates on every change")


def _finish(state: _State, tokens: list[Token]) -> None:
    for frame in reversed(state.stack):
        label = f"{frame.name}(" if frame.name else "("
        state.errors.append(f"missing ')' to close {label}")
    if tokens and tokens[-1].type in (Token.OP_IN, Token.OP_PRE):
        state.errors.append(f"formula ends with operator '{tokens[-1].value}'")


def _unmatched_close(text: str) -> int | None:
    """1-based position of the first ')' with no opener, ignoring quoted text and sheet names."""
    depth, quote = 0, ""
    for index, char in enumerate(text, start=1):
        if quote:
            quote = "" if char == quote else quote
        elif char in "\"'":
            quote = char
        elif char == "(":
            depth += 1
        elif char == ")":
            depth -= 1
            if depth < 0:
                return index
    return None


def _clean_name(raw: str) -> str:
    name = raw.upper()
    for prefix in _PREFIXES:
        if name.startswith(prefix.upper()):
            return name[len(prefix) :]
    return name


def _arity(info: FunctionInfo) -> str:
    if info.max_args == VARIADIC:
        return f"at least {info.min_args}"
    if info.min_args == info.max_args:
        return str(info.min_args)
    return f"{info.min_args}-{info.max_args}"
