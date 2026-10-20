"""MCP adapter (stdio / streamable-http). Logs go to stderr; stdout is the protocol."""

from __future__ import annotations

import functools
from collections.abc import Callable
from typing import Any

from mcp.server.mcpserver import MCPServer
from mcp.server.mcpserver.exceptions import ToolError
from mcp.types import ToolAnnotations

from slim_sheet import __version__
from slim_sheet.config import Settings
from slim_sheet.core.errors import SlimError
from slim_sheet.interfaces.handlers import Handlers
from slim_sheet.interfaces.toolspec import enabled_tools

INSTRUCTIONS = (
    "Spreadsheet tools. Start with overview(path) to see sheets, columns and types. "
    "Use read for cells, query for SQL (tables are named after sheets), profile for stats, "
    "find to locate text. write never edits the original unless overwrite=true. "
    "Responses are size-limited; follow 'next' cursors or offsets."
)


def _guard(fn: Callable[..., str]) -> Callable[..., str]:
    """Turn SlimError into an MCP tool error whose text is the JSON error contract."""

    @functools.wraps(fn)  # keeps the real signature, which MCP turns into the input schema
    def wrapper(*args: Any, **kwargs: Any) -> str:
        try:
            return fn(*args, **kwargs)
        except SlimError as err:
            raise ToolError(err.to_json()) from err

    return wrapper


def build_server(settings: Settings) -> MCPServer:
    server: MCPServer = MCPServer(
        name="slim-sheet", version=__version__, instructions=INSTRUCTIONS, log_level="WARNING"
    )

    handlers = Handlers(settings)
    for spec in enabled_tools(settings):
        server.add_tool(
            _guard(getattr(handlers, spec.name)),
            name=spec.name,
            description=spec.description,
            annotations=ToolAnnotations(read_only_hint=spec.read_only, open_world_hint=False),
            structured_output=False,
        )
    return server
