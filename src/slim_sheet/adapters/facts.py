"""Per-file-version cache on disk: the folder big sheets are cached in, and small facts about a
big file (its value-cell count, its hidden-content findings) that otherwise cost a full pass
over hundreds of MB of sheet XML on every call. Keys include size and modification time, so
an edited file is scanned again.
"""

from __future__ import annotations

import hashlib
import json
import logging
import os
import sys
from collections.abc import Callable
from pathlib import Path
from typing import TypeVar

log = logging.getLogger(__name__)

FORMAT_VERSION = 1
FACT_MIN_BYTES = 5 * 2**20  # smaller files are scanned in milliseconds: nothing is stored
T = TypeVar("T")


def cache_dir() -> Path:
    """SLIM_SHEET_CACHE_DIR, else the user's cache folder (LOCALAPPDATA / XDG_CACHE_HOME)."""
    configured = os.environ.get("SLIM_SHEET_CACHE_DIR", "").strip()
    if configured:
        return Path(configured).expanduser()
    if sys.platform == "win32":
        base = Path(os.environ.get("LOCALAPPDATA") or Path.home() / "AppData" / "Local")
    else:
        base = Path(os.environ.get("XDG_CACHE_HOME") or Path.home() / ".cache")
    return base / "slim-sheet"


def version_key(path: Path, part: str = "") -> tuple[str, str]:
    """(name of this file version's entries, prefix shared by every version of `part`)."""
    stat = path.stat()
    owner = hashlib.sha256(f"{path.resolve()}|{part}".encode()).hexdigest()[:20]
    version = hashlib.sha256(
        f"{stat.st_size}|{stat.st_mtime_ns}|{FORMAT_VERSION}".encode()
    ).hexdigest()[:12]
    return f"{owner}-{version}", owner


def remembered(path: Path, name: str, compute: Callable[[], T], keep: Callable[[T], bool]) -> T:
    """`compute()` once per file version; stored when `keep(value)` (e.g. a complete count)."""
    if path.stat().st_size < FACT_MIN_BYTES:
        return compute()
    stem, owner = version_key(path)
    folder = cache_dir() / "facts"
    file = folder / f"{stem}.json"
    facts: dict[str, object] = {}
    try:
        facts = json.loads(file.read_text(encoding="utf-8")) if file.exists() else {}
    except (OSError, ValueError) as exc:
        log.warning("Ignoring unreadable fact cache %s: %s", file, exc)
    if name in facts:
        return facts[name]  # type: ignore[return-value]
    value = compute()
    if keep(value):
        facts[name] = value
        try:
            folder.mkdir(parents=True, exist_ok=True)
            for stale in folder.glob(f"{owner}-*.json"):
                if stale != file:
                    stale.unlink(missing_ok=True)
            file.write_text(json.dumps(facts), encoding="utf-8")
        except OSError as exc:
            log.warning("Could not store facts for %s: %s", path.name, exc)
    return value
