"""Environment -> frozen settings (PLAN §13)."""

from __future__ import annotations

import os
from collections.abc import Mapping
from dataclasses import dataclass, field
from pathlib import Path

from slim_sheet.core.errors import ErrorCode, SlimError

TIERS = ("lite", "standard", "full")


@dataclass(frozen=True, slots=True)
class Settings:
    tier: str = "standard"
    tools: tuple[str, ...] = ()
    roots: tuple[Path, ...] = ()
    max_tokens: int = 1500
    max_cell_chars: int = 120
    max_file_mb: int = 200
    soffice: str | None = None  # LibreOffice binary for recalc; autodetected when unset
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


def load_settings(env: Mapping[str, str] | None = None) -> Settings:
    env = os.environ if env is None else env
    tier = env.get("SLIM_SHEET_TIER", "standard").strip().lower()
    if tier not in TIERS:
        raise SlimError(
            ErrorCode.BAD_ARGUMENT,
            f"SLIM_SHEET_TIER='{tier}' is not valid",
            hint=f"Use one of: {', '.join(TIERS)}",
        )
    tools = tuple(t.strip() for t in env.get("SLIM_SHEET_TOOLS", "").split(",") if t.strip())
    raw_roots = env.get("SLIM_SHEET_ROOTS", "")
    roots = tuple(Path(r).expanduser().resolve() for r in raw_roots.split(os.pathsep) if r.strip())
    return Settings(
        tier=tier,
        tools=tools,
        roots=roots or (Path.cwd().resolve(),),
        max_tokens=_int(env, "SLIM_SHEET_MAX_TOKENS", 1500),
        max_cell_chars=_int(env, "SLIM_SHEET_MAX_CELL_CHARS", 120),
        max_file_mb=_int(env, "SLIM_SHEET_MAX_FILE_MB", 200),
        soffice=env.get("SLIM_SHEET_SOFFICE") or None,
        token=env.get("SLIM_SHEET_TOKEN") or None,
        password=env.get("SLIM_SHEET_PASSWORD") or None,
    )
