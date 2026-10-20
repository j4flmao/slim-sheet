"""Path allow-list: every file access is resolved against SLIM_SHEET_ROOTS."""

from __future__ import annotations

from pathlib import Path

from slim_sheet.config import Settings
from slim_sheet.core.errors import ErrorCode, SlimError


def resolve_output(raw: str, settings: Settings) -> Path:
    """Resolve a path that may not exist yet; it must sit inside a root, in an existing folder."""
    if not raw or not raw.strip():
        raise SlimError(ErrorCode.BAD_ARGUMENT, "out_path is empty")
    candidate = Path(raw.strip()).expanduser()
    if not candidate.is_absolute():
        candidate = settings.roots[0] / candidate
    resolved = candidate.resolve()
    if not any(resolved.is_relative_to(root) for root in settings.roots):
        raise SlimError(
            ErrorCode.PATH_DENIED,
            f"'{raw}' is outside the allowed roots",
            hint=f"Write inside: {', '.join(str(r) for r in settings.roots)}",
        )
    if not resolved.parent.is_dir():
        raise SlimError(ErrorCode.FILE_NOT_FOUND, f"Folder '{resolved.parent}' does not exist")
    return resolved


def resolve_path(raw: str, settings: Settings) -> Path:
    if not raw or not raw.strip():
        raise SlimError(ErrorCode.BAD_ARGUMENT, "path is empty", hint="e.g. path='data/sales.xlsx'")
    candidate = Path(raw.strip()).expanduser()
    if not candidate.is_absolute():
        candidate = settings.roots[0] / candidate
    resolved = candidate.resolve()
    if not any(resolved.is_relative_to(root) for root in settings.roots):
        roots = ", ".join(str(r) for r in settings.roots)
        raise SlimError(
            ErrorCode.PATH_DENIED,
            f"'{raw}' is outside the allowed roots",
            hint=f"Allowed roots: {roots}. Set SLIM_SHEET_ROOTS to widen access.",
        )
    if not resolved.is_file():
        raise SlimError(
            ErrorCode.FILE_NOT_FOUND,
            f"No file at '{resolved}'",
            hint=f"Paths are relative to {settings.roots[0]}",
        )
    size_mb = resolved.stat().st_size / (1024 * 1024)
    if size_mb > settings.max_file_mb:
        raise SlimError(
            ErrorCode.TOO_LARGE,
            f"'{resolved.name}' is {size_mb:.0f} MB (limit {settings.max_file_mb} MB)",
            hint="Raise SLIM_SHEET_MAX_FILE_MB if this is intended.",
        )
    return resolved
