from __future__ import annotations

import asyncio
import json
from dataclasses import replace
from pathlib import Path

import pytest
from mcp.server.mcpserver.exceptions import ToolError

from slim_sheet.config import Settings, load_settings
from slim_sheet.core.errors import SlimError
from slim_sheet.interfaces.mcp_server import build_server
from slim_sheet.interfaces.toolspec import TIER_CAPS, TOOLS, enabled_tools


def test_tier_caps_are_respected() -> None:
    # Act / Assert
    for tier, cap in TIER_CAPS.items():
        assert len(enabled_tools(Settings(tier=tier))) <= cap


def test_tool_descriptions_are_short() -> None:
    # Act / Assert
    for spec in TOOLS:
        assert len(spec.description.split()) <= 30, spec.name


def test_server_lists_lite_tools_with_flat_args(settings: Settings) -> None:
    # Arrange
    server = build_server(replace(settings, tier="lite"))

    # Act
    tools = asyncio.run(server.list_tools())

    # Assert
    lite = {
        "overview",
        "read",
        "find",
        "query",
        "profile",
        "write",
        "explain_formula",
        "build_report",
    }
    assert {t.name for t in tools} == lite
    read_schema = next(t for t in tools if t.name == "read").input_schema
    expected = ["path", "sheet", "range", "cursor", "max_tokens", "format"]
    assert list(read_schema["properties"]) == expected
    assert read_schema["required"] == ["path"]
    for prop in read_schema["properties"].values():
        assert "properties" not in prop


def test_server_returns_error_contract_as_tool_error(sales_xlsx: Path, settings: Settings) -> None:
    # Arrange
    server = build_server(settings)

    # Act / Assert
    with pytest.raises(ToolError) as info:
        asyncio.run(server.call_tool("read", {"path": "sales.xlsx", "sheet": "Ordrs"}))
    message = str(info.value)  # SDK prefixes "Error executing tool read: "
    payload = json.loads(message[message.index("{") :])
    assert payload["error"] == "sheet_not_found"
    assert "Orders" in payload["hint"]


def test_explicit_tool_list_rejects_unknown_names() -> None:
    # Act / Assert
    with pytest.raises(SlimError):
        enabled_tools(Settings(tools=("overview", "teleport")))


def test_load_settings_rejects_bad_tier() -> None:
    # Act / Assert
    with pytest.raises(SlimError):
        load_settings({"SLIM_SHEET_TIER": "mega"})


def test_every_tool_schema_is_flat(settings: Settings) -> None:
    # Arrange
    server = build_server(replace(settings, tier="full"))

    # Act
    tools = asyncio.run(server.list_tools())

    # Assert
    for tool in tools:
        for name, prop in tool.input_schema["properties"].items():
            assert "properties" not in prop, f"{tool.name}.{name} is a nested object"
            assert "$ref" not in str(prop), f"{tool.name}.{name} uses $ref"


def test_write_tool_is_marked_destructive(settings: Settings) -> None:
    # Act
    tools = {t.name: t for t in asyncio.run(build_server(settings).list_tools())}

    # Assert
    assert tools["write"].annotations is not None
    assert tools["write"].annotations.read_only_hint is False
    assert tools["query"].annotations is not None
    assert tools["query"].annotations.read_only_hint is True


def test_default_tier_is_full_and_placeholders_fall_back() -> None:
    # Act / Assert
    assert load_settings({}).tier == "full"
    assert load_settings({"SLIM_SHEET_TIER": ""}).tier == "full"
    assert load_settings({"SLIM_SHEET_TIER": "${user_config.tier}"}).tier == "full"
    assert load_settings({"SLIM_SHEET_TIER": " Lite "}).tier == "lite"
    assert len(enabled_tools(load_settings({}))) == 30


def test_instructions_name_every_enabled_tool_and_hidden_ones(settings: Settings) -> None:
    # Arrange
    from slim_sheet.interfaces.mcp_server import build_instructions

    full = {t.name for t in enabled_tools(replace(settings, tier="full"))}
    lite = {t.name for t in enabled_tools(replace(settings, tier="lite"))}

    # Act
    full_text = build_instructions(full)
    lite_text = build_instructions(lite)

    # Assert
    assert all(name in full_text for name in full)
    assert "instead of writing Python/openpyxl" in full_text
    assert "Not enabled" not in full_text
    assert "range_ops (fill, copy, move, sort, replace, insert/delete rows or columns)" in full_text
    assert "Not enabled in this tier:" in lite_text and "add_chart" in lite_text
