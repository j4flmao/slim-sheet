from __future__ import annotations

import datetime as dt
from dataclasses import replace
from pathlib import Path

import openpyxl
import pytest

from slim_sheet.config import Settings
from slim_sheet.core.errors import ErrorCode, SlimError
from slim_sheet.interfaces.http_app import build_http_app, is_loopback
from slim_sheet.interfaces.schema import export_tools, simplify
from slim_sheet.services.overview import overview
from slim_sheet.services.query import query


def test_is_loopback() -> None:
    # Act / Assert
    assert is_loopback("127.0.0.1") and is_loopback("localhost") and is_loopback("::1")
    assert not is_loopback("0.0.0.0") and not is_loopback("example.com")


def test_http_refuses_public_bind_without_token(settings: Settings) -> None:
    # Act / Assert
    with pytest.raises(SlimError) as info:
        build_http_app(settings, "0.0.0.0")
    assert "SLIM_SHEET_TOKEN" in info.value.hint


def test_http_bearer_auth(settings: Settings) -> None:
    # Arrange
    from starlette.testclient import TestClient

    app = build_http_app(replace(settings, token="s3cret-token"), "0.0.0.0")
    client_info = {"name": "t", "version": "1"}
    params = {"protocolVersion": "2025-06-18", "capabilities": {}, "clientInfo": client_info}
    init = {"jsonrpc": "2.0", "id": 1, "method": "initialize", "params": params}
    accept = {"accept": "application/json, text/event-stream"}

    # Act
    with TestClient(app) as client:
        denied = client.post("/mcp", json=init, headers=accept)
        wrong = client.post("/mcp", json=init, headers={**accept, "authorization": "Bearer nope"})
        allowed = client.post(
            "/mcp", json=init, headers={**accept, "authorization": "Bearer s3cret-token"}
        )

    # Assert
    assert denied.status_code == 401 and wrong.status_code == 401
    assert allowed.status_code == 200


def test_schema_export_formats_are_flat_and_lean(settings: Settings) -> None:
    # Act
    openai = export_tools(replace(settings, tier="lite"), "openai")
    anthropic = export_tools(replace(settings, tier="full"), "anthropic")

    # Assert
    assert len(openai) == 8 and openai[0]["type"] == "function"
    read = next(t for t in openai if t["function"]["name"] == "read")["function"]["parameters"]
    assert read["properties"]["sheet"] == {"type": "string"}
    assert len(anthropic) == 30 and "input_schema" in anthropic[0]
    assert all(
        "title" not in prop for t in anthropic for prop in t["input_schema"]["properties"].values()
    )
    with pytest.raises(SlimError):
        export_tools(settings, "yaml")


def test_simplify_opens_scalar_unions_and_keeps_complex_ones() -> None:
    # Act: Gemini and OpenAI strict mode reject anyOf, so scalar unions become an open type
    scalar = simplify(
        {"title": "X", "anyOf": [{"type": "string"}, {"type": "integer"}, {"type": "null"}]}
    )
    complex_union = simplify(
        {"anyOf": [{"type": "string"}, {"type": "array", "items": {"type": "string"}}]}
    )

    # Assert
    assert scalar == {"description": "a cell value (string, integer)"}
    assert "anyOf" in complex_union


def test_encrypted_workbook_read_and_edit_refused(sales_xlsx: Path, settings: Settings) -> None:
    # Arrange
    pytest.importorskip("msoffcrypto")
    from msoffcrypto.format.ooxml import OOXMLFile

    locked = sales_xlsx.with_name("locked.xlsx")
    with sales_xlsx.open("rb") as src, locked.open("wb") as out:
        OOXMLFile(src).encrypt("pw123", out)

    # Act / Assert
    with pytest.raises(SlimError) as no_pw:
        overview("locked.xlsx", settings)
    assert no_pw.value.code == ErrorCode.ENCRYPTED
    with pytest.raises(SlimError) as wrong:
        overview("locked.xlsx", replace(settings, password="nope"))
    assert "Wrong password" in wrong.value.message
    opened = replace(settings, password="pw123")
    assert query("locked.xlsx", 'SELECT count(*) FROM "Orders"', opened).rows == ((300,),)
    from slim_sheet.services.write import WriteRequest, write

    with pytest.raises(SlimError) as edit:
        write(WriteRequest("locked.xlsx", "Orders", "A1", [["x"]]), opened)
    assert edit.value.code == ErrorCode.ENCRYPTED


def test_numbers_document(tmp_path: Path, settings: Settings) -> None:
    # Arrange
    numbers_parser = pytest.importorskip("numbers_parser")
    doc = numbers_parser.Document()
    table = doc.sheets[0].tables[0]
    for c, v in enumerate(["name", "amount", "when"]):
        table.write(0, c, v)
    table.write(1, 0, "a")
    table.write(1, 1, 12.5)
    table.write(1, 2, dt.datetime(2026, 1, 2))
    doc.save(str(tmp_path / "t.numbers"))

    # Act
    info = overview("t.numbers", settings)
    result = query("t.numbers", 'SELECT amount FROM "Sheet 1"', settings)

    # Assert
    assert info.format.value == "numbers"
    assert info.sheets[0].used is not None and info.sheets[0].used.n_cols == 3
    assert result.rows == ((12.5,),)
    assert openpyxl  # xlsx library stays importable alongside the extras


def test_generated_manifests_are_fresh_and_consistent() -> None:
    # Arrange
    import json
    import subprocess
    import sys

    root = Path(__file__).resolve().parent.parent

    # Act
    check = subprocess.run(
        [sys.executable, str(root / "scripts" / "build_manifests.py"), "--check"],
        capture_output=True,
        text=True,
        check=False,
    )
    server = json.loads((root / "server.json").read_text(encoding="utf-8"))
    manifest = json.loads((root / "manifest.json").read_text(encoding="utf-8"))
    readme = (root / "README.md").read_text(encoding="utf-8")

    # Assert
    assert check.returncode == 0, check.stderr
    assert f"<!-- mcp-name: {server['name']} -->" in readme
    assert len(server["description"]) <= 100
    assert server["version"] == manifest["version"]
    assert len(manifest["tools"]) == 30
    assert (root / manifest["server"]["entry_point"]).is_file()
