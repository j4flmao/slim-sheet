"""Task loading and answer scoring for the agent eval suite. Pure logic except `check_written`."""

from __future__ import annotations

import json
import math
import re
from collections.abc import Iterable, Mapping, Sequence
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

from jsonschema import Draft202012Validator  # type: ignore[import-untyped]

ANSWER_KINDS = frozenset({"number", "text", "all", "any", "cells"})
CHECK_KINDS = ANSWER_KINDS | {"written"}
DEFAULT_TOLERANCE = 0.005

_ANSWER_MARK = re.compile(r"answer\s*:", re.IGNORECASE)
_NUMBER = re.compile(r"(?<![\w.])-?\d[\d,]*(?:\.\d+)?")
_CELL = re.compile(
    r"(?<![\w$])(?:('[^']+'|[A-Za-z_][\w.]*)!)?\$?([A-Z]{1,3})\$?([1-9]\d{0,6})(?!\w)"
)


@dataclass(frozen=True, slots=True)
class ToolCall:
    tool: str
    args: Mapping[str, Any]


@dataclass(frozen=True, slots=True)
class Check:
    kind: str
    value: Any = None
    values: tuple[str, ...] = ()
    tolerance: float = DEFAULT_TOLERANCE
    file: str = ""
    cells: Mapping[str, Any] = field(default_factory=dict)


@dataclass(frozen=True, slots=True)
class Task:
    id: str
    prompt: str
    tools: tuple[str, ...]
    check: Check
    reference: tuple[ToolCall, ...]


@dataclass(frozen=True, slots=True)
class CallVerdict:
    valid: bool
    args: dict[str, Any]
    reason: str = ""


# ---------------------------------------------------------------- loading


def load_tasks(path: Path) -> list[Task]:
    """Parse tasks.json and fail loudly on any malformed entry."""
    try:
        raw = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as exc:
        raise ValueError(f"Cannot load eval tasks from {path}: {exc}") from exc
    tasks = [parse_task(item) for item in raw.get("tasks", [])]
    if not tasks:
        raise ValueError(f"No tasks found in {path}")
    ids = [t.id for t in tasks]
    duplicates = sorted({i for i in ids if ids.count(i) > 1})
    if duplicates:
        raise ValueError(f"Duplicate task ids in {path}: {', '.join(duplicates)}")
    return tasks


def parse_task(item: Mapping[str, Any]) -> Task:
    task_id = str(item.get("id", "")).strip()
    if not task_id or not str(item.get("prompt", "")).strip():
        raise ValueError(f"Task needs a non-empty id and prompt: {item!r}")
    reference = tuple(ToolCall(str(c["tool"]), dict(c["args"])) for c in item.get("reference", []))
    if not reference:
        raise ValueError(f"Task '{task_id}' has no reference calls")
    return Task(
        id=task_id,
        prompt=str(item["prompt"]),
        tools=tuple(item.get("tools", ())),
        check=parse_check(task_id, item.get("check", {})),
        reference=reference,
    )


def parse_check(task_id: str, raw: Mapping[str, Any]) -> Check:
    kind = raw.get("kind")
    if kind not in CHECK_KINDS:
        raise ValueError(f"Task '{task_id}': check kind {kind!r} not in {sorted(CHECK_KINDS)}")
    if kind in {"number", "text"} and raw.get("value") is None:
        raise ValueError(f"Task '{task_id}': '{kind}' check needs a value")
    if kind in {"all", "any", "cells"} and not raw.get("values"):
        raise ValueError(f"Task '{task_id}': '{kind}' check needs a non-empty values list")
    if kind == "written" and (not raw.get("file") or not raw.get("cells")):
        raise ValueError(f"Task '{task_id}': 'written' check needs file and cells")
    return Check(
        kind=kind,
        value=raw.get("value"),
        values=tuple(str(v) for v in raw.get("values", ())),
        tolerance=float(raw.get("tolerance", DEFAULT_TOLERANCE)),
        file=str(raw.get("file", "")),
        cells=dict(raw.get("cells", {})),
    )


# ---------------------------------------------------------------- answer parsing


def answer_part(text: str) -> str:
    """Text after the last 'ANSWER:' marker, or the whole text when there is none."""
    marks = list(_ANSWER_MARK.finditer(text))
    return text[marks[-1].end() :].strip() if marks else text


def extract_numbers(text: str) -> list[float]:
    numbers: list[float] = []
    for match in _NUMBER.finditer(text):
        token = match.group(0).replace(",", "")
        try:
            numbers.append(float(token))
        except ValueError:
            continue  # e.g. a lone "-" glued to a list separator
    return numbers


def extract_cells(text: str) -> set[tuple[str | None, str]]:
    """A1 addresses as (lower-cased sheet or None, cell) pairs."""
    found: set[tuple[str | None, str]] = set()
    for sheet, col, row in _CELL.findall(text):
        name = sheet.strip("'").lower() if sheet else None
        found.add((name, f"{col}{row}"))
    return found


def _contains(text: str, needle: str, bounded: bool) -> bool:
    pattern = re.escape(needle)
    if bounded and needle[:1].isalnum():
        pattern = r"(?<!\w)" + pattern
    if bounded and needle[-1:].isalnum():
        pattern += r"(?!\w)"
    return re.search(pattern, text, re.IGNORECASE) is not None


def _number_matches(text: str, expected: float, tolerance: float) -> bool:
    allowed = max(tolerance, abs(expected) * 1e-9)
    return any(math.isclose(n, expected, abs_tol=allowed) for n in extract_numbers(text))


def _cells_match(text: str, expected: Sequence[str], strict: bool) -> bool:
    want = extract_cells(" ".join(expected))
    got = extract_cells(text)

    def same(a: tuple[str | None, str], b: tuple[str | None, str]) -> bool:
        return a[1] == b[1] and (a[0] is None or b[0] is None or a[0] == b[0])

    if not all(any(same(w, g) for g in got) for w in want):
        return False
    return not strict or all(any(same(g, w) for w in want) for g in got)


def check_answer(check: Check, text: str, strict: bool = True) -> bool:
    """Score a final answer. strict=False is for raw tool output (extra cells allowed)."""
    body = answer_part(text) if strict else text
    if check.kind == "number":
        return _number_matches(body, float(check.value), check.tolerance)
    if check.kind == "text":
        return _contains(body, str(check.value), bounded=True)
    if check.kind == "all":
        return all(_contains(body, v, bounded=True) for v in check.values)
    if check.kind == "any":
        return any(_contains(body, v, bounded=False) for v in check.values)
    if check.kind == "cells":
        return _cells_match(body, check.values, strict)
    raise ValueError(f"check_answer cannot score kind '{check.kind}'; use check_written")


# ---------------------------------------------------------------- written files


def values_equal(actual: Any, expected: Any, tolerance: float = DEFAULT_TOLERANCE) -> bool:
    if isinstance(expected, int | float) and not isinstance(expected, bool):
        try:
            return math.isclose(float(actual), float(expected), abs_tol=tolerance)
        except (TypeError, ValueError):
            return False
    return str(actual if actual is not None else "").strip().lower() == str(expected).lower()


def split_address(address: str) -> tuple[str, str]:
    sheet, sep, cell = address.rpartition("!")
    if not sep or not sheet or not cell:
        raise ValueError(f"Expected 'Sheet!A1', got '{address}'")
    return sheet.strip("'"), cell


def written_mismatches(check: Check, workdir: Path) -> list[str]:
    """Cells of the written workbook that differ from the expectation (empty = pass)."""
    import openpyxl

    target = workdir / check.file
    if not target.is_file():
        return [f"{check.file} was not written"]
    wb = openpyxl.load_workbook(target, data_only=False, read_only=True)
    try:
        problems: list[str] = []
        for address, expected in check.cells.items():
            sheet, cell = split_address(address)
            if sheet not in wb.sheetnames:
                problems.append(f"{address}: sheet missing")
                continue
            actual = wb[sheet][cell].value
            if not values_equal(actual, expected, check.tolerance):
                problems.append(f"{address}: got {actual!r}, want {expected!r}")
        return problems
    finally:
        wb.close()


def check_written(check: Check, workdir: Path) -> bool:
    return not written_mismatches(check, workdir)


# ---------------------------------------------------------------- tool-call validity


def validate_call(
    name: str, raw_args: Any, schemas: Mapping[str, Mapping[str, Any]]
) -> CallVerdict:
    """A call is valid when the tool exists and its JSON args satisfy the input schema."""
    if name not in schemas:
        return CallVerdict(False, {}, f"unknown tool '{name}'")
    if isinstance(raw_args, str):
        try:
            raw_args = json.loads(raw_args) if raw_args.strip() else {}
        except json.JSONDecodeError as exc:
            return CallVerdict(False, {}, f"arguments are not JSON: {exc.msg}")
    if not isinstance(raw_args, dict):
        return CallVerdict(False, {}, "arguments must be a JSON object")
    errors = sorted(Draft202012Validator(schemas[name]).iter_errors(raw_args), key=str)
    if errors:
        return CallVerdict(False, raw_args, errors[0].message)
    return CallVerdict(True, raw_args)


def _json_objects(text: str) -> list[dict[str, Any]]:
    decoder = json.JSONDecoder()
    found: list[dict[str, Any]] = []
    index = text.find("{")
    while index != -1:
        try:
            obj, end = decoder.raw_decode(text, index)
        except json.JSONDecodeError:
            index = text.find("{", index + 1)
            continue
        if isinstance(obj, dict):
            found.append(obj)
        index = text.find("{", end)
    return found


def parse_text_calls(content: str) -> list[dict[str, Any]]:
    """Recover tool calls a model wrote as JSON text (`{"name": ..., "arguments": {...}}`),
    optionally inside <tool_call> tags or code fences. Returns OpenAI-style tool_calls."""
    calls: list[dict[str, Any]] = []
    for obj in _json_objects(content):
        name = obj.get("name")
        args = obj.get("arguments", obj.get("parameters"))
        if not isinstance(name, str) or args is None:
            continue
        arguments = args if isinstance(args, str) else json.dumps(args)
        calls.append({"type": "function", "function": {"name": name, "arguments": arguments}})
    return calls


def to_openai_schema(schema: Any) -> Any:
    """Simplify an MCP input schema for small models: drop titles, collapse `X | None`."""
    if isinstance(schema, list):
        return [to_openai_schema(s) for s in schema]
    if not isinstance(schema, dict):
        return schema
    options = schema.get("anyOf")
    if isinstance(options, list):
        non_null = [o for o in options if o != {"type": "null"}]
        if len(non_null) == 1:
            rest = {k: v for k, v in schema.items() if k != "anyOf"}
            return to_openai_schema({**rest, **non_null[0]})
    return {k: to_openai_schema(v) for k, v in schema.items() if k != "title"}


def openai_tools(tools: Iterable[tuple[str, str, Mapping[str, Any]]]) -> list[dict[str, Any]]:
    return [
        {
            "type": "function",
            "function": {
                "name": name,
                "description": description,
                "parameters": to_openai_schema(dict(schema)),
            },
        }
        for name, description, schema in tools
    ]
