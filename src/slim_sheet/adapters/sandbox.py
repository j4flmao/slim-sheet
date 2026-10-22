"""Where tools may read and write (PLAN §17.5).

Default ("open") mode — flexible, no configuration needed:
- an absolute path is used as given, anywhere on a local disk, so data and outputs can live in any
  folder the user names;
- SLIM_SHEET_ROOTS are the default folders: a relative path is looked up in each of them in order
  (the first that has it wins), and overview() with no path lists them;
- a relative output path is written next to the file it comes from (`base`);
- outside the default folders only spreadsheet/tabular files are opened or written, and system and
  secret locations are always refused (OS and program folders, AppData, dot-folders such as .ssh
  or .git, /etc, network shares), so a prompt injected into a cell cannot reach credentials.

Strict mode (SLIM_SHEET_STRICT_ROOTS=1): only paths inside SLIM_SHEET_ROOTS, as before.
"""

from __future__ import annotations

import os
from pathlib import Path, PurePath

from slim_sheet.config import Settings
from slim_sheet.core.errors import ErrorCode, SlimError

# files the tools can open or produce; anything else outside the default folders is refused
DATA_SUFFIXES = frozenset({
    ".xlsx", ".xlsm", ".xltx", ".xltm", ".xlsb", ".xls", ".ods", ".numbers",
    ".csv", ".tsv", ".txt", ".parquet", ".json", ".jsonl", ".ndjson",
})  # fmt: skip
ALLOWED_DOT_DIRS = frozenset({".slim-sheet"})  # slim-sheet's own journal
_POSIX_SYSTEM = ("/etc", "/bin", "/sbin", "/usr", "/var", "/proc", "/sys", "/dev", "/boot",
                 "/root", "/System", "/Library", "/private", "/opt")  # fmt: skip


def _candidate(raw: str, settings: Settings, base: Path | None = None) -> Path:
    """Where a (possibly relative) path points, before the access check."""
    candidate = Path(raw.strip()).expanduser()
    if candidate.is_absolute():
        return candidate
    if base is not None:
        return base / candidate
    found = next((root / candidate for root in settings.roots if (root / candidate).exists()), None)
    return found or settings.roots[0] / candidate


def _in_roots(path: Path, settings: Settings) -> bool:
    return any(path.is_relative_to(root) for root in settings.roots)


def _system_folders() -> list[Path]:
    """OS, program and per-user application-data folders (never data folders)."""
    names = ("SystemRoot", "ProgramFiles", "ProgramFiles(x86)", "ProgramW6432", "ProgramData",
             "APPDATA", "LOCALAPPDATA")  # fmt: skip
    found = [Path(os.environ[n]).resolve() for n in names if os.environ.get(n)]
    found += [Path(p) for p in _POSIX_SYSTEM if os.name != "nt"]
    return found


def _refusal(path: Path, folder: bool) -> str | None:
    """Why an outside-the-roots path is refused, or None when it is fine."""
    if str(path).startswith(("\\\\", "//")):
        return "network shares are not opened"
    if any(p.startswith(".") and p not in ALLOWED_DOT_DIRS for p in PurePath(path).parts[1:-1]):
        return "hidden folders (.ssh, .git, .aws, ...) are not opened"
    if path.name.startswith(".") and path.name not in ALLOWED_DOT_DIRS and folder:
        return "hidden folders (.ssh, .git, .aws, ...) are not opened"
    if any(path == s or path.is_relative_to(s) for s in _system_folders()):
        return "system and application-data folders are not opened"
    if not folder and path.suffix.lower() not in DATA_SUFFIXES:
        return f"'{path.suffix or path.name}' is not a spreadsheet or table file"
    return None


def _check(raw: str, resolved: Path, settings: Settings, folder: bool) -> Path:
    roots = ", ".join(str(r) for r in settings.roots)
    if raw.strip().startswith(("\\\\", "//")):
        # A UNC path means a network share on every OS; on Linux/macOS a backslash is not a
        # separator, so without this check it would resolve to an odd file name in the root.
        raise SlimError(
            ErrorCode.PATH_DENIED,
            f"'{raw}' is refused: network shares are not opened",
            hint=f"Copy the file to a local data folder such as {roots}.",
        )
    if _in_roots(resolved, settings):
        return resolved
    if settings.strict_roots:
        raise SlimError(
            ErrorCode.PATH_DENIED,
            f"'{raw}' is outside the allowed roots",
            hint=f"Allowed roots: {roots} (SLIM_SHEET_STRICT_ROOTS=1 is set).",
        )
    reason = _refusal(resolved, folder)
    if reason is not None:
        raise SlimError(
            ErrorCode.PATH_DENIED,
            f"'{raw}' is refused: {reason}",
            hint=f"Use a data folder such as {roots}, or another folder on a local disk.",
        )
    return resolved


def resolve_output(raw: str, settings: Settings, base: Path | None = None) -> Path:
    """Resolve a path that may not exist yet (a file or, without a suffix, a folder).

    `base` is the folder of the source file: a relative out_path is written next to it.
    """
    if not raw or not raw.strip():
        raise SlimError(ErrorCode.BAD_ARGUMENT, "out_path is empty")
    resolved = _candidate(raw, settings, base).resolve()
    _check(raw, resolved, settings, folder=not resolved.suffix)
    if not resolved.parent.is_dir():
        raise SlimError(
            ErrorCode.FILE_NOT_FOUND,
            f"Folder '{resolved.parent}' does not exist",
            hint="Write into an existing folder, e.g. next to the source file.",
        )
    return resolved


def _inside_roots(raw: str, settings: Settings, folder: bool = False) -> Path:
    if not raw or not raw.strip():
        raise SlimError(ErrorCode.BAD_ARGUMENT, "path is empty", hint="e.g. path='data/sales.xlsx'")
    resolved = _candidate(raw, settings).resolve()
    return _check(raw, resolved, settings, folder=folder or resolved.is_dir())


def resolve_folder(raw: str, settings: Settings) -> Path | None:
    """A readable directory, or None when `raw` is not a directory."""
    resolved = _candidate(raw, settings).resolve() if raw and raw.strip() else None
    if resolved is None or not resolved.is_dir():
        return None
    return _inside_roots(raw, settings, folder=True)


def resolve_path(raw: str, settings: Settings) -> Path:
    resolved = _inside_roots(raw, settings)
    if not resolved.is_file():
        raise SlimError(
            ErrorCode.FILE_NOT_FOUND,
            f"No file at '{resolved}'",
            hint="Give the full path, or a path relative to: "
            + ", ".join(str(r) for r in settings.roots)
            + ". overview() with no path lists the spreadsheets there.",
        )
    size_mb = resolved.stat().st_size / (1024 * 1024)
    if size_mb > settings.max_file_mb:
        raise SlimError(
            ErrorCode.TOO_LARGE,
            f"'{resolved.name}' is {size_mb:.0f} MB (limit {settings.max_file_mb} MB)",
            hint="Raise SLIM_SHEET_MAX_FILE_MB if this is intended.",
        )
    return resolved
