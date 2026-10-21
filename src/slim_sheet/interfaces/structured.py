"""Structured tool output (MCP `outputSchema` + `structuredContent`) for the data tools.

The Markdown/TSV text stays the primary content (it is what models read cheaply); hosts that
consume results programmatically also get the same result as JSON. Handlers publish the model
behind their text through `PAYLOAD`; the MCP wrapper turns it into `structuredContent`.
"""

from __future__ import annotations

import dataclasses
import datetime as dt
import enum
from contextvars import ContextVar
from typing import Any

from slim_sheet.core.a1 import format_range
from slim_sheet.core.models import CellRange

PAYLOAD: ContextVar[object | None] = ContextVar("slim_sheet_payload", default=None)

_OBJECT: dict[str, Any] = {"type": "object"}
_ROWS = {"type": "array", "items": {"type": "array"}}
SCHEMAS: dict[str, dict[str, Any]] = {
    "overview": {
        "type": "object",
        "properties": {
            "path": {"type": "string"},
            "format": {"type": "string"},
            "sheets": {"type": "array", "items": _OBJECT},
            "warnings": {"type": "array", "items": {"type": "string"}},
        },
    },
    "read": {
        "type": "object",
        "properties": {
            "sheet": {"type": "string"},
            "window": {"type": "string"},
            "rows": _ROWS,
            "next_cursor": {"type": ["string", "null"]},
        },
    },
    "find": {"type": "object", "properties": {"hits": {"type": "array", "items": _OBJECT}}},
    "query": {
        "type": "object",
        "properties": {
            "columns": {"type": "array", "items": {"type": "string"}},
            "rows": _ROWS,
            "total_rows": {"type": "integer"},
            "next_offset": {"type": ["integer", "null"]},
        },
    },
    "profile": {"type": "object", "properties": {"columns": {"type": "array", "items": _OBJECT}}},
    "diff": {"type": "object", "properties": {"changes": {"type": "array"}}},
}
STRUCTURED_TOOLS = frozenset(SCHEMAS)


def to_json(value: Any) -> Any:
    """Dataclass models -> plain JSON values (ranges as A1 text, dates as ISO strings)."""
    if isinstance(value, CellRange):
        return format_range(value)
    if dataclasses.is_dataclass(value) and not isinstance(value, type):
        return {f.name: to_json(getattr(value, f.name)) for f in dataclasses.fields(value)}
    if isinstance(value, dict):
        return {str(k): to_json(v) for k, v in value.items()}
    if isinstance(value, list | tuple):
        return [to_json(v) for v in value]
    if isinstance(value, dt.date | dt.time):
        return value.isoformat()
    if isinstance(value, enum.Enum):
        return value.value
    if value is None or isinstance(value, str | int | float | bool):
        return value
    return str(value)


def publish(model: object) -> None:
    PAYLOAD.set(model)


def take() -> dict[str, Any] | None:
    """The payload published during this call, as a JSON object (and clear it)."""
    model = PAYLOAD.get()
    PAYLOAD.set(None)
    if model is None:
        return None
    data = to_json(model)
    return data if isinstance(data, dict) else {"result": data}
