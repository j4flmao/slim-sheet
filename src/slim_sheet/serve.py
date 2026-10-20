"""Stdio MCP entry point without CLI parsing (used by the .mcpb bundle and `python -m`).

`python src/slim_sheet/serve.py` and `python -m slim_sheet.serve` both start the server
with settings from the environment (SLIM_SHEET_TIER, SLIM_SHEET_ROOTS, ...).
"""

from __future__ import annotations

import logging
import sys


def main() -> None:
    logging.basicConfig(stream=sys.stderr, level=logging.WARNING)
    from slim_sheet.config import load_settings
    from slim_sheet.interfaces.mcp_server import build_server

    build_server(load_settings()).run(transport="stdio")


if __name__ == "__main__":
    main()
