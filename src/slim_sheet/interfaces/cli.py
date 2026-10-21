"""CLI adapter: `slim-sheet <tool> ...`, `slim-sheet mcp`, `slim-sheet tools`.

Tool subcommands are generated from the `Handlers` method signatures, so the CLI can
never drift from the MCP tool schemas: required parameters become positionals, optional
ones become `--flags`, `list[str]` takes several values and 2-D arrays take JSON.
"""

from __future__ import annotations

import argparse
import inspect
import json
import logging
import sys
import types
import typing
from collections.abc import Sequence
from dataclasses import replace
from typing import Any

from slim_sheet import __version__
from slim_sheet.config import TIERS, Settings, load_settings
from slim_sheet.core.errors import ErrorCode, SlimError
from slim_sheet.interfaces.handlers import Handlers
from slim_sheet.interfaces.toolspec import TOOLS, enabled_tools

_META = {"command", "tier", "transport", "host", "port", "format"}


def _parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(prog="slim-sheet", description=__doc__)
    parser.add_argument("--version", action="version", version=__version__)
    sub = parser.add_subparsers(dest="command", required=True)
    mcp = sub.add_parser("mcp", help="run the MCP server")
    mcp.add_argument("--transport", choices=["stdio", "streamable-http"], default="stdio")
    mcp.add_argument("--tier", choices=TIERS)
    mcp.add_argument("--host", default="127.0.0.1", help="streamable-http only")
    mcp.add_argument("--port", type=int, default=8765, help="streamable-http only")
    schema = sub.add_parser("schema", help="print tool definitions for non-MCP hosts")
    schema.add_argument("--format", choices=["openai", "anthropic", "mcp"], default="openai")
    schema.add_argument("--tier", choices=TIERS)
    tools = sub.add_parser("tools", help="list tools enabled for a tier")
    tools.add_argument("--tier", choices=TIERS)
    for spec in TOOLS:
        tool = sub.add_parser(spec.name, help=spec.description, description=spec.description)
        _add_parameters(tool, getattr(Handlers, spec.name))
    return parser


def _add_parameters(parser: argparse.ArgumentParser, method: Any) -> None:
    hints = typing.get_type_hints(method)
    for name, param in inspect.signature(method).parameters.items():
        if name == "self":
            continue
        _add_parameter(parser, name, hints[name], param.default)


def _unwrap_optional(hint: Any) -> tuple[Any, bool]:
    args = typing.get_args(hint)
    is_union = typing.get_origin(hint) in (typing.Union, types.UnionType)
    if is_union and type(None) in args:
        rest = [a for a in args if a is not type(None)]
        return rest[0], True
    return hint, False


def _add_parameter(parser: argparse.ArgumentParser, name: str, hint: Any, default: Any) -> None:
    base, _optional = _unwrap_optional(hint)
    required = default is inspect.Parameter.empty
    flag = "--" + name.replace("_", "-")
    if base is bool:
        if default is True:
            parser.add_argument(f"--no-{name.replace('_', '-')}", dest=name, action="store_false")
        elif default is False:
            parser.add_argument(flag, dest=name, action="store_true")
        else:
            parser.add_argument(flag, dest=name, type=_boolean, default=None, metavar="true|false")
        return
    kwargs: dict[str, Any] = {}
    if typing.get_origin(base) is list:
        inner = typing.get_args(base)[0]
        if typing.get_origin(inner) is list:
            kwargs["type"] = _json_grid
        else:
            kwargs["nargs"] = "+"
    elif base in (int, float):
        kwargs["type"] = base
    if required:
        parser.add_argument(name, **kwargs)
    else:
        parser.add_argument(flag, dest=name, default=default, **kwargs)


def _boolean(raw: str) -> bool:
    lowered = raw.strip().lower()
    if lowered in ("true", "1", "yes", "y"):
        return True
    if lowered in ("false", "0", "no", "n"):
        return False
    raise argparse.ArgumentTypeError(f"expected true or false, got '{raw}'")


def _json_grid(raw: str) -> Any:
    try:
        values = json.loads(raw)
    except json.JSONDecodeError as exc:
        raise argparse.ArgumentTypeError(f"values must be JSON: {exc}") from exc
    if not isinstance(values, list) or not all(isinstance(r, list) for r in values):
        raise argparse.ArgumentTypeError("values must be a JSON array of arrays")
    return values


def _run(args: argparse.Namespace, settings: Settings) -> str | None:
    if getattr(args, "tier", None):
        settings = replace(settings, tier=args.tier)
    if args.command == "mcp":
        _serve(args, settings)
        return None
    if args.command == "schema":
        from slim_sheet.interfaces.schema import export_json

        return export_json(settings, args.format)
    if args.command == "tools":
        return "\n".join(f"{t.name} [{t.tier}] {t.description}" for t in enabled_tools(settings))
    if args.command not in {t.name for t in enabled_tools(settings)}:
        raise SlimError(
            ErrorCode.BAD_ARGUMENT,
            f"Tool '{args.command}' is not enabled for tier '{settings.tier}'",
            hint="Set SLIM_SHEET_TIER=full or add it to SLIM_SHEET_TOOLS.",
        )
    kwargs = {k: v for k, v in vars(args).items() if k not in _META}
    result: str = getattr(Handlers(settings), args.command)(**kwargs)
    return result


def _serve(args: argparse.Namespace, settings: Settings) -> None:
    if args.transport == "streamable-http":
        from slim_sheet.interfaces.http_app import serve_http

        serve_http(settings, args.host, args.port)
        return
    from slim_sheet.interfaces.mcp_server import build_server

    build_server(settings).run(transport="stdio")


# Tools merged in Phase 7 (PLAN §17.2): old commands explain where they went.
RENAMED = {
    "validate_formula": "evaluate --check-only",
    "split": "convert --split-by <column>",
    "add_validation": "rules (kind=list|whole|decimal|date|text_length|custom)",
    "conditional_format": "rules (kind=greater_than|between|duplicates|color_scale|...)",
    "rows_ops": "range_ops (op=insert|delete, range='5:7' or 'C:E')",
}


def main(argv: Sequence[str] | None = None) -> int:
    logging.basicConfig(stream=sys.stderr, level=logging.WARNING)
    argv = list(sys.argv[1:] if argv is None else argv)
    if argv and argv[0] in RENAMED:
        error = SlimError(
            ErrorCode.BAD_ARGUMENT,
            f"'{argv[0]}' was merged into another tool",
            hint=f"Use: slim-sheet {RENAMED[argv[0]]}",
        )
        print(error.to_json())
        return 2
    args = _parser().parse_args(argv)
    sys.stdout.reconfigure(encoding="utf-8")  # type: ignore[union-attr]
    try:
        output = _run(args, load_settings())
    except SlimError as err:
        print(err.to_json())
        return 1
    if output is not None:
        print(output)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
