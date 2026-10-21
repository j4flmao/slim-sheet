"""Environment -> frozen settings (PLAN §13)."""

from __future__ import annotations

import os
from collections.abc import Mapping
from dataclasses import dataclass, field
from pathlib import Path

from slim_sheet.core.errors import ErrorCode, SlimError

TIERS = ("lite", "standard", "full")
DEFAULT_TIER = "full"  # every tool unless a host opts into a smaller surface for small models


@dataclass(frozen=True, slots=True)
class Settings:
    tier: str = DEFAULT_TIER
    tools: tuple[str, ...] = ()
    roots: tuple[Path, ...] = ()
    max_tokens: int = 1500
    max_cell_chars: int = 120
    max_file_mb: int = 200
    max_cells: int = 3_000_000  # cells per xlsx package before any full load; 0 = off
    soffice: str | None = None  # LibreOffice binary for recalc; autodetected when unset
    brand: str | None = None  # built-in brand name or brand .json path (PLAN §16)
    audit_log: str | None = None  # JSONL file for per-call audit records (opt-in)
    audit_user: str | None = None  # principal written to the audit log
    rate_limit: int = 120  # streamable-http: tool calls per minute per client, 0 = off
    allowed_origins: tuple[str, ...] = ()  # browser origins accepted besides loopback
    token: str | None = field(default=None, repr=False)  # bearer token for streamable-http
    password: str | None = field(default=None, repr=False)  # opens encrypted workbooks


def _int(env: Mapping[str, str], key: str, default: int) -> int:
    raw = env.get(key)
    if raw is None or raw == "":
        return default
    try:
        return int(raw)
    except ValueError as exc:
        raise SlimError(ErrorCode.BAD_ARGUMENT, f"{key} must be an integer, got '{raw}'") from exc


def _clean(raw: str | None) -> str:
    """Treat unset, blank and unresolved host placeholders (e.g. '${user_config.tier}') as unset.

    Some hosts pass a template literally when the user left a field empty; refusing to start
    would look like "the server has no tools", so such values fall back to defaults.
    """
    value = (raw or "").strip()
    return "" if value.startswith("${") and value.endswith("}") else value


def load_settings(env: Mapping[str, str] | None = None) -> Settings:
    env = os.environ if env is None else env
    tier = _clean(env.get("SLIM_SHEET_TIER")).lower() or DEFAULT_TIER
    if tier not in TIERS:
        raise SlimError(
            ErrorCode.BAD_ARGUMENT,
            f"SLIM_SHEET_TIER='{tier}' is not valid",
            hint=f"Use one of: {', '.join(TIERS)}",
        )
    tools = tuple(t.strip() for t in _clean(env.get("SLIM_SHEET_TOOLS")).split(",") if t.strip())
    raw_roots = _clean(env.get("SLIM_SHEET_ROOTS"))
    roots = tuple(Path(r).expanduser().resolve() for r in raw_roots.split(os.pathsep) if r.strip())
    return Settings(
        tier=tier,
        tools=tools,
        roots=roots or (Path.cwd().resolve(),),
        max_tokens=_int(env, "SLIM_SHEET_MAX_TOKENS", 1500),
        max_cell_chars=_int(env, "SLIM_SHEET_MAX_CELL_CHARS", 120),
        max_file_mb=_int(env, "SLIM_SHEET_MAX_FILE_MB", 200),
        max_cells=_int(env, "SLIM_SHEET_MAX_CELLS", 3_000_000),
        soffice=env.get("SLIM_SHEET_SOFFICE") or None,
        brand=_clean(env.get("SLIM_SHEET_BRAND")) or None,
        audit_log=_clean(env.get("SLIM_SHEET_AUDIT_LOG")) or None,
        audit_user=_clean(env.get("SLIM_SHEET_AUDIT_USER")) or None,
        rate_limit=_int(env, "SLIM_SHEET_RATE_LIMIT", 120),
        allowed_origins=tuple(
            o.strip() for o in _clean(env.get("SLIM_SHEET_ALLOWED_ORIGINS")).split(",") if o.strip()
        ),
        token=env.get("SLIM_SHEET_TOKEN") or None,
        password=env.get("SLIM_SHEET_PASSWORD") or None,
    )
