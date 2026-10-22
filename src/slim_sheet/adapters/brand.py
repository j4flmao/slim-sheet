"""Resolve which brand a tool call uses: argument, then SLIM_SHEET_BRAND, then 'slim'.

A brand is a built-in name ('slim', 'office', 'modern', 'mono', 'gov') or a .json file whose
keys follow core/design.py (only 'name' is required). A company sets SLIM_SHEET_BRAND once and
every agent's output carries its colours and fonts without the model ever seeing a hex code.
"""

from __future__ import annotations

import json
from pathlib import Path
from typing import Any

from slim_sheet.adapters.sandbox import resolve_path
from slim_sheet.config import Settings
from slim_sheet.core.design import (
    BRANDS,
    DEFAULT_BRAND,
    Brand,
    brand_from_mapping,
    builtin_brand,
)
from slim_sheet.core.errors import ErrorCode, SlimError


def brand_for(wb: Any, look: Brand, explicit: bool) -> Brand:
    """The brand for an edit: the call's or the environment's when given, otherwise the brand the
    workbook was built with (its '<brand> title' named style), so later calls keep its look."""
    if explicit:
        return look
    names = {str(getattr(s, "name", s)) for s in getattr(wb, "named_styles", [])}
    found = next((b for b in BRANDS if f"{b} title" in names), None)
    return builtin_brand(found) if found and found != look.name else look


def resolve_brand(requested: str | None, settings: Settings) -> tuple[Brand, bool]:
    """-> (brand, explicit). `explicit` is True when the call or the env chose the brand."""
    if requested and requested.strip():
        return _load(requested.strip(), settings, from_env=False), True
    if settings.brand:
        return _load(settings.brand, settings, from_env=True), True
    return builtin_brand(DEFAULT_BRAND), False


def _load(value: str, settings: Settings, from_env: bool) -> Brand:
    if not value.lower().endswith(".json"):
        return builtin_brand(value)
    # The env value is operator configuration; a tool argument follows the normal path rules
    # (any local data folder; system, AppData and hidden folders are refused).
    path = Path(value).expanduser() if from_env else resolve_path(value, settings)
    if not path.is_file():
        raise SlimError(
            ErrorCode.FILE_NOT_FOUND,
            f"Brand file '{value}' does not exist",
            hint="Use a built-in brand: slim, office, modern, mono, gov",
        )
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, UnicodeDecodeError, json.JSONDecodeError) as exc:
        raise SlimError(
            ErrorCode.BAD_ARGUMENT,
            f"Brand file '{path.name}' is not valid JSON: {exc}",
            hint='e.g. {"name": "Acme", "primary": "#C8102E"}',
        ) from exc
    if not isinstance(data, dict):
        raise SlimError(ErrorCode.BAD_ARGUMENT, f"Brand file '{path.name}' must hold an object")
    return brand_from_mapping(data, path.name)
