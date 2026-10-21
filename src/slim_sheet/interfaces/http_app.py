"""Streamable HTTP transport with bearer-token auth (PLAN §5 Phase 5).

- The token comes only from SLIM_SHEET_TOKEN (never a CLI flag, so it stays out of shell
  history and process listings).
- Binding to anything but loopback without a token is refused.
- Comparison is constant-time; failures return 401 with a JSON error, never a hint.
"""

from __future__ import annotations

import hashlib
import hmac
import ipaddress
import json
import time
from collections import defaultdict, deque
from typing import Any
from urllib.parse import urlsplit

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


class RateLimit:
    """Sliding one-minute window of POSTs (tool calls) per client: 429 + Retry-After beyond it.

    The client is the bearer token when there is one, otherwise the peer address. The MCP spec
    requires servers to rate-limit tool invocations; stdio (one local client) is not limited.
    """

    def __init__(self, app: Any, per_minute: int, clock: Any = time.monotonic) -> None:
        self.app = app
        self.per_minute = per_minute
        self.clock = clock
        self.calls: dict[bytes, deque[float]] = defaultdict(deque)

    async def __call__(self, scope: Scope, receive: Receive, send: Send) -> None:
        if scope["type"] != "http" or scope.get("method") != "POST" or self.per_minute <= 0:
            await self.app(scope, receive, send)
            return
        now = self.clock()
        window = self.calls[_client(scope)]
        while window and now - window[0] >= 60:
            window.popleft()
        if len(window) >= self.per_minute:
            retry = max(1, int(60 - (now - window[0])))
            await _respond(send, 429, "rate_limited", f"over {self.per_minute} calls/minute", retry)
            return
        window.append(now)
        await self.app(scope, receive, send)


class OriginCheck:
    """Browsers send Origin; anything but loopback or an allowed origin gets 403 (DNS rebinding)."""

    def __init__(self, app: Any, allowed: tuple[str, ...]) -> None:
        self.app = app
        self.allowed = {a.rstrip("/").lower() for a in allowed}

    async def __call__(self, scope: Scope, receive: Receive, send: Send) -> None:
        if scope["type"] == "http":
            origin = dict(scope.get("headers") or []).get(b"origin", b"").decode("latin-1")
            if origin and not self._ok(origin):
                await _respond(send, 403, "forbidden_origin", f"origin {origin} is not allowed")
                return
        await self.app(scope, receive, send)

    def _ok(self, origin: str) -> bool:
        host = urlsplit(origin).hostname or ""
        return is_loopback(host) or origin.rstrip("/").lower() in self.allowed


def _client(scope: Scope) -> bytes:
    headers = dict(scope.get("headers") or [])
    token = headers.get(b"authorization")
    if token:
        return hashlib.sha256(token).digest()  # never keep tokens in memory as dict keys
    peer = scope.get("client") or ("unknown", 0)
    return str(peer[0]).encode()


async def _respond(send: Send, status: int, code: str, message: str, retry: int = 0) -> None:
    headers = [(b"content-type", b"application/json")]
    if retry:
        headers.append((b"retry-after", str(retry).encode()))
    body = json.dumps({"error": code, "message": message}).encode()
    await send({"type": "http.response.start", "status": status, "headers": headers})
    await send({"type": "http.response.body", "body": body})


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
    app: Any = build_server(settings).streamable_http_app(host=host)
    app = RateLimit(app, settings.rate_limit)
    if settings.token:
        app = BearerAuth(app, settings.token)
    return OriginCheck(app, settings.allowed_origins)


def serve_http(settings: Settings, host: str, port: int) -> None:
    import uvicorn

    uvicorn.run(build_http_app(settings, host), host=host, port=port, log_level="warning")
