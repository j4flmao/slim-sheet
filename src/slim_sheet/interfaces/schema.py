"""Export tool definitions for hosts that do not speak MCP (PLAN §2, rule "one core, many faces").

The schemas come from the live MCP server object, so they are exactly what MCP clients see:
- openai:    [{"type": "function", "function": {name, description, parameters}}]
- anthropic: [{name, description, input_schema}]
- mcp:       [{name, description, inputSchema, annotations}]
"""

from __future__ import annotations

import asyncio
import json
from typing import Any

from slim_sheet.config import Settings
from slim_sheet.core.errors import ErrorCode, SlimError

FORMATS = ("openai", "anthropic", "mcp")


def export_tools(settings: Settings, fmt: str) -> list[dict[str, Any]]:
    if fmt not in FORMATS:
        raise SlimError(
            ErrorCode.BAD_ARGUMENT,
            f"Unknown format '{fmt}'",
            hint=f"Use one of: {', '.join(FORMATS)}",
        )
    from slim_sheet.interfaces.mcp_server import build_server

    tools = asyncio.run(build_server(settings).list_tools())
    out: list[dict[str, Any]] = []
    for tool in tools:
        schema = simplify(tool.input_schema)
        if fmt == "openai":
            out.append(
                {
                    "type": "function",
                    "function": {
                        "name": tool.name,
                        "description": tool.description,
                        "parameters": schema,
                    },
                }
            )
        elif fmt == "anthropic":
            out.append({"name": tool.name, "description": tool.description, "input_schema": schema})
        else:
            annotations = (
                tool.annotations.model_dump(by_alias=True, exclude_none=True)
                if tool.annotations
                else {}
            )
            out.append(
                {
                    "name": tool.name,
                    "description": tool.description,
                    "inputSchema": schema,
                    "annotations": annotations,
                }
            )
    return out


def export_json(settings: Settings, fmt: str) -> str:
    return json.dumps(export_tools(settings, fmt), indent=2, ensure_ascii=False)


def simplify(schema: dict[str, Any]) -> dict[str, Any]:
    """Drop pydantic titles and collapse `anyOf: [T, null]` to `T`.

    Optionality is already carried by `required`, and many local function-calling models
    mishandle `anyOf`; plain types are also fewer tokens on every request.
    """
    out: dict[str, Any] = {}
    for key, value in schema.items():
        if key == "title":
            continue
        if key == "properties":
            out[key] = {name: simplify(prop) for name, prop in value.items()}
        elif key == "items" and isinstance(value, dict):
            out[key] = simplify(value)
        elif key == "anyOf":
            options = [opt for opt in value if opt.get("type") != "null"]
            if len(options) == 1:
                out.update(simplify(options[0]))
            else:
                out[key] = [simplify(opt) for opt in options]
        elif key == "default" and value is None:
            continue
        else:
            out[key] = value
    return out
