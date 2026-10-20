"""Generate server.json (MCP Registry) and manifest.json (MCP Bundle) from the single sources
of truth: pyproject.toml (name, version, description) and interfaces/toolspec.py (tools).

Pack the bundle with `npx @anthropic-ai/mcpb pack` from the repository root.

Usage:  uv run python scripts/build_manifests.py          # write both files
        uv run python scripts/build_manifests.py --check  # exit 1 if they are stale (CI)
"""

from __future__ import annotations

import argparse
import json
import sys
import tomllib
from pathlib import Path
from typing import Any

from slim_sheet.interfaces.toolspec import TOOLS

ROOT = Path(__file__).resolve().parent.parent
REGISTRY_NAME = "io.github.j4flmao/slim-sheet"
REPOSITORY = "https://github.com/j4flmao/slim-sheet"
SHORT_DESCRIPTION = (
    "Spreadsheet MCP for any agent: reads xlsx/xls/ods/csv, SQL, cleaning, formulas, safe edits"
)
SERVER_SCHEMA = "https://static.modelcontextprotocol.io/schemas/2025-12-11/server.schema.json"
ENV_VARS = [
    (
        "SLIM_SHEET_TIER",
        "Tool tier: lite (7 tools, small models), standard (16) or full (28)",
        "standard",
        False,
    ),
    (
        "SLIM_SHEET_ROOTS",
        "Folders the server may read and write, separated by the OS path separator",
        None,
        False,
    ),
    ("SLIM_SHEET_MAX_TOKENS", "Default response budget in tokens", "1500", False),
    (
        "SLIM_SHEET_PASSWORD",
        "Password for encrypted workbooks (needs the crypto extra)",
        None,
        True,
    ),
]


def _project() -> dict[str, Any]:
    with (ROOT / "pyproject.toml").open("rb") as handle:
        project: dict[str, Any] = tomllib.load(handle)["project"]
    return project


def server_json(project: dict[str, Any]) -> dict[str, Any]:
    env = []
    for name, description, default, secret in ENV_VARS:
        item: dict[str, Any] = {"name": name, "description": description, "isRequired": False}
        if default is not None:
            item["default"] = default
        if secret:
            item["isSecret"] = True
        env.append(item)
    return {
        "$schema": SERVER_SCHEMA,
        "name": REGISTRY_NAME,
        "title": "slim-sheet",
        "description": SHORT_DESCRIPTION,
        "repository": {"url": REPOSITORY, "source": "github"},
        "version": project["version"],
        "packages": [
            {
                "registryType": "pypi",
                "registryBaseUrl": "https://pypi.org",
                "identifier": project["name"],
                "version": project["version"],
                "runtimeHint": "uvx",
                "packageArguments": [{"type": "positional", "value": "mcp"}],
                "transport": {"type": "stdio"},
                "environmentVariables": env,
            }
        ],
    }


def mcpb_manifest(project: dict[str, Any]) -> dict[str, Any]:
    return {
        "manifest_version": "0.4",
        "name": project["name"],
        "display_name": "slim-sheet",
        "version": project["version"],
        "description": project["description"],
        "long_description": (
            "Reads xlsx, xls, xlsb, ods, Numbers, csv, parquet and json locally. "
            "Token-budgeted output, SQL over sheets, cleaning, anomaly detection, "
            "a formula copilot and Excel-correct editing. Edits go to a .slim copy "
            "unless you allow overwriting, and every change can be undone."
        ),
        "author": {"name": project["authors"][0]["name"], "email": project["authors"][0]["email"]},
        "repository": {"type": "git", "url": REPOSITORY},
        "homepage": REPOSITORY,
        "license": "MIT",
        "keywords": project["keywords"],
        "server": {
            "type": "uv",
            "entry_point": "src/slim_sheet/serve.py",
            "mcp_config": {
                "command": "uv",
                "args": ["run", "--directory", "${__dirname}", "src/slim_sheet/serve.py"],
                "env": {
                    "SLIM_SHEET_TIER": "${user_config.tier}",
                    "SLIM_SHEET_ROOTS": "${user_config.data_folder}",
                },
            },
        },
        "tools": [{"name": t.name, "description": t.description} for t in TOOLS],
        "user_config": {
            "data_folder": {
                "type": "directory",
                "title": "Data folder",
                "description": "Folder with the spreadsheets slim-sheet may read and write",
                "required": True,
            },
            "tier": {
                "type": "string",
                "title": "Tool tier",
                "description": "lite (7 tools), standard (16) or full (28)",
                "default": "standard",
                "required": False,
            },
        },
        "compatibility": {
            "platforms": ["darwin", "win32", "linux"],
            "runtimes": {"python": ">=3.11"},
        },
    }


def _render(data: dict[str, Any]) -> str:
    return json.dumps(data, indent=2, ensure_ascii=False) + "\n"


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--check", action="store_true", help="fail if the files are out of date")
    args = parser.parse_args()
    project = _project()
    outputs = {
        ROOT / "server.json": _render(server_json(project)),
        ROOT / "manifest.json": _render(mcpb_manifest(project)),
    }
    stale = [
        p for p, text in outputs.items() if not p.exists() or p.read_text(encoding="utf-8") != text
    ]
    if args.check:
        for path in stale:
            sys.stderr.write(f"stale: {path.relative_to(ROOT)} (run scripts/build_manifests.py)\n")
        return 1 if stale else 0
    for path, text in outputs.items():
        path.parent.mkdir(exist_ok=True)
        path.write_text(text, encoding="utf-8")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
