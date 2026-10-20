"""Streamable HTTP transport with bearer-token auth (PLAN §5 Phase 5).

- The token comes only from SLIM_SHEET_TOKEN (never a CLI flag, so it stays out of shell
  history and process listings).
- Binding to anything but loopback without a token is refused.
- Comparison is constant-time; failures return 401 with a JSON error, never a hint.
"""

from __future__ import annotations

import hmac
import ipaddress
import json
from typing import Any

from slim_sheet.config import Settings
from slim_sheet.core.errors import ErrorCode, SlimError

LOOPBACK_NAMES = {"localhost"}
Scope = dict[str, Any]
Receive = Any
Send = Any


def is_loopback(host: str) -> bool:
    if host in LOOPBACK_NAMES:
        return True
    try:
        return ipaddress.ip_address(host).is_loopback
    except ValueError:
        return False


class BearerAuth:
    """Pure ASGI middleware; lifespan and non-HTTP scopes pass straight through."""

    def __init__(self, app: Any, token: str) -> None:
        self.app = app
        self.expected = f"Bearer {token}".encode()

    async def __call__(self, scope: Scope, receive: Receive, send: Send) -> None:
        if scope["type"] != "http":
            await self.app(scope, receive, send)
            return
        headers = dict(scope.get("headers") or [])
        supplied = headers.get(b"authorization", b"")
        if not hmac.compare_digest(supplied, self.expected):
            await _reject(send)
            return
        await self.app(scope, receive, send)


async def _reject(send: Send) -> None:
    body = json.dumps(
        {"error": "unauthorized", "message": "missing or invalid bearer token"}
    ).encode()
    await send(
        {
            "type": "http.response.start",
            "status": 401,
            "headers": [
                (b"content-type", b"application/json"),
                (b"www-authenticate", b'Bearer realm="slim-sheet"'),
            ],
        }
    )
    await send({"type": "http.response.body", "body": body})


def build_http_app(settings: Settings, host: str) -> Any:
    from slim_sheet.interfaces.mcp_server import build_server

    if not settings.token and not is_loopback(host):
        raise SlimError(
            ErrorCode.BAD_ARGUMENT,
            f"Refusing to serve on {host} without authentication",
            hint="Set SLIM_SHEET_TOKEN to a long random value, or bind to 127.0.0.1.",
        )
    app = build_server(settings).streamable_http_app(host=host)
    return BearerAuth(app, settings.token) if settings.token else app


def serve_http(settings: Settings, host: str, port: int) -> None:
    import uvicorn

    uvicorn.run(build_http_app(settings, host), host=host, port=port, log_level="warning")
