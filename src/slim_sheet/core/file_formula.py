"""Formula text as Excel stores it in the file (PLAN §8, verified against Excel 365 saves).

Excel writes functions newer than 2007 as `_xlfn.NAME` (FILTER and SORT as `_xlfn._xlws.`),
and LET/LAMBDA parameter names as `_xlpm.name`. A formula saved without these prefixes opens
as #NAME? — or, for FILTER, Excel refuses to open the whole file. Spilling (dynamic-array)
formulas also need array storage plus the XLDAPR cell metadata (writers/dynamic_arrays.py).
"""

from __future__ import annotations

from openpyxl.formula.tokenizer import Token, Tokenizer, TokenizerError

from slim_sheet.core.a1 import parse_range
from slim_sheet.core.errors import SlimError
from slim_sheet.core.functions import FUNCTIONS
from slim_sheet.core.models import CellRange

_FUTURE_VERSIONS = {"2010", "2013", "2016", "2019", "2021", "365"}
FUTURE = frozenset(n for n, info in FUNCTIONS.items() if info.since in _FUTURE_VERSIONS)
XLWS = frozenset({"FILTER", "SORT"})
# functions whose result is an array that spills over neighbouring cells
DYNAMIC = frozenset({
    "FILTER", "SORT", "SORTBY", "UNIQUE", "SEQUENCE", "RANDARRAY", "TRANSPOSE", "TEXTSPLIT",
    "VSTACK", "HSTACK", "TOROW", "TOCOL", "WRAPROWS", "WRAPCOLS", "TAKE", "DROP", "EXPAND",
    "CHOOSEROWS", "CHOOSECOLS", "MAKEARRAY", "MAP", "SCAN", "BYROW", "BYCOL", "GROUPBY",
    "PIVOTBY", "REGEXEXTRACT", "TRIMRANGE", "MUNIT", "MODE.MULT", "FREQUENCY",
})  # fmt: skip
_PARAM_FUNCTIONS = ("LET", "LAMBDA")
_PREFIXES = ("_xlfn._xlws.", "_xlfn.", "_xlws.")
_PARAM = "_xlpm."


def _tokens(formula: str) -> list[Token] | None:
    try:
        return list(Tokenizer(formula).items)
    except (TokenizerError, ValueError, IndexError):
        return None  # left as typed; validate() reports the syntax problem


def _bare(token: Token) -> str:
    name = str(token.value)[:-1]  # "FILTER(" -> "FILTER"
    for prefix in _PREFIXES:
        if name.upper().startswith(prefix.upper()):
            return name[len(prefix) :].upper()
    return name.upper()


def _is_open(token: Token) -> bool:
    return bool(token.type == Token.FUNC and token.subtype == Token.OPEN)


def to_file(formula: str) -> str:
    """'=FILTER(A:A,A:A>1)' -> '=_xlfn._xlws.FILTER(A:A,A:A>1)'; other text unchanged."""
    if not formula.startswith("=") or formula.startswith("=="):
        return formula
    tokens = _tokens(formula)
    if tokens is None:
        return formula
    params = _parameter_names(tokens)
    parts: list[str] = []
    for token in tokens:
        if _is_open(token):
            parts.append(_prefixed(_bare(token)) + "(")
        elif _is_param(token, params):
            parts.append(_PARAM + token.value)
        else:
            parts.append(token.value)
    return "=" + "".join(parts)


def from_file(formula: str) -> str:
    """The way Excel shows a stored formula: '_xlfn.', '_xlws.' and '_xlpm.' removed."""
    if "_xl" not in formula or not formula.startswith("="):
        return formula
    tokens = _tokens(formula)
    if tokens is None:
        return formula
    parts: list[str] = []
    for token in tokens:
        value = str(token.value)
        if _is_open(token):
            value = _bare_case(value) + "("
        elif token.type == Token.OPERAND and value.lower().startswith(_PARAM):
            value = value[len(_PARAM) :]
        parts.append(value)
    return "=" + "".join(parts)


def _bare_case(value: str) -> str:
    """Strip the prefixes from 'FUNC(' keeping the name's own spelling."""
    name = value[:-1]
    for prefix in _PREFIXES:
        if name.upper().startswith(prefix.upper()):
            return name[len(prefix) :]
    return name


def _prefixed(name: str) -> str:
    if name in XLWS:
        return f"_xlfn._xlws.{name}"
    return f"_xlfn.{name}" if name in FUTURE else name


def _is_param(token: Token, params: set[str]) -> bool:
    return (
        token.type == Token.OPERAND
        and token.subtype == Token.RANGE
        and token.value.casefold() in params
    )


def _parameter_names(tokens: list[Token]) -> set[str]:
    """Names bound by LET(name, value, ..., calc) and LAMBDA(p1, ..., body)."""
    names: set[str] = set()
    stack: list[tuple[str, int]] = []  # (function, argument index)
    for i, token in enumerate(tokens):
        if _is_open(token) or (token.type == Token.PAREN and token.subtype == Token.OPEN):
            stack.append((_bare(token) if _is_open(token) else "", 0))
        elif token.subtype == Token.CLOSE and stack:
            stack.pop()
        elif token.type == Token.SEP and token.subtype == Token.ARG and stack:
            name, index = stack[-1]
            stack[-1] = (name, index + 1)
        elif stack and stack[-1][0] in _PARAM_FUNCTIONS and _binds(tokens, i, *stack[-1]):
            names.add(token.value.casefold())
    return names


def _binds(tokens: list[Token], i: int, function: str, index: int) -> bool:
    token = tokens[i]
    if token.type != Token.OPERAND or token.subtype != Token.RANGE:
        return False
    if token.value.lower().startswith(_PARAM) or ":" in token.value or "!" in token.value:
        return False
    nxt = tokens[i + 1] if i + 1 < len(tokens) else None
    last = nxt is None or nxt.type != Token.SEP  # the final argument is the calculation
    prev = tokens[i - 1] if i else None
    alone = prev is not None and (_is_open(prev) or prev.type == Token.SEP)
    return alone and not last and (function == "LAMBDA" or index % 2 == 0)


def is_dynamic(formula: str) -> bool:
    """True when the formula calls a function that returns a spilling array."""
    tokens = _tokens(formula) if formula.startswith("=") else None
    return bool(tokens) and any(_is_open(t) and _bare(t) in DYNAMIC for t in tokens or [])


def spill_extent(formula: str, anchor: CellRange) -> CellRange | None:
    """Largest area the result can cover, when the formula shows it; None when unknown.

    SEQUENCE(rows, cols) with literal sizes is exact; FILTER/SORT/SORTBY/UNIQUE are bounded
    by their first argument and TRANSPOSE by its swapped argument.
    """
    tokens = _tokens(formula) if formula.startswith("=") else None
    if not tokens or not _is_open(tokens[0]):
        return None
    name, args = _bare(tokens[0]), _first_level_args(tokens)
    rows, cols = 0, 0
    if name == "SEQUENCE" and args and all(a.isdigit() for a in args[:2]):
        rows, cols = int(args[0]), int(args[1]) if len(args) > 1 else 1
    elif name in ("FILTER", "SORT", "SORTBY", "UNIQUE", "TRANSPOSE") and args:
        try:
            source = parse_range(args[0])
        except SlimError:  # not a plain A1 range (a nested call, a name, a sheet)
            return None
        rows, cols = source.n_rows, source.n_cols
        if name == "TRANSPOSE":
            rows, cols = cols, rows
    if rows < 1 or cols < 1 or rows * cols > 1_000_000:
        return None
    return CellRange(anchor.top, anchor.left, anchor.top + rows - 1, anchor.left + cols - 1)


def _first_level_args(tokens: list[Token]) -> list[str]:
    """Plain-text arguments of the outermost call (nested calls become '')."""
    args: list[str] = [""]
    depth = 0
    for token in tokens[1:]:
        if _is_open(token) or (token.type == Token.PAREN and token.subtype == Token.OPEN):
            depth += 1
            args[-1] = ""
        elif token.subtype == Token.CLOSE:
            if depth == 0:
                break
            depth -= 1
        elif depth == 0 and token.type == Token.SEP and token.subtype == Token.ARG:
            args.append("")
        elif depth == 0 and token.type != Token.WSPACE:
            args[-1] += token.value
    return [a.strip() for a in args]
