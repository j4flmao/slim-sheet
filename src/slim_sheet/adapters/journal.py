"""Write journal: every file slim-sheet creates or replaces is recorded so `undo` can
restore it (PLAN §8). Nothing is ever deleted: undone states are kept as journal copies.

Layout: <folder>/.slim-sheet/journal/journal.jsonl plus NNNNN-<name> backup files.
"""

from __future__ import annotations

import json
import os
import shutil
import tempfile
import time
from dataclasses import dataclass
from pathlib import Path
from typing import Any

from slim_sheet.core.errors import ErrorCode, SlimError

JOURNAL_DIR = Path(".slim-sheet") / "journal"
INDEX = "journal.jsonl"


@dataclass(frozen=True, slots=True)
class UndoOutcome:
    target: str
    note: str
    restored_from: str | None  # None: the undone write had created the file, so it is removed
    kept_as: str  # journal copy of the state that was undone


def _folder(target: Path) -> Path:
    return target.parent / JOURNAL_DIR


def _entries(folder: Path) -> list[dict[str, Any]]:
    index = folder / INDEX
    if not index.exists():
        return []
    with index.open(encoding="utf-8") as handle:
        return [json.loads(line) for line in handle if line.strip()]


def _append(folder: Path, entry: dict[str, Any]) -> int:
    folder.mkdir(parents=True, exist_ok=True)
    seq = len(_entries(folder)) + 1
    entry = {"seq": seq, **entry, "ts": time.time()}
    with (folder / INDEX).open("a", encoding="utf-8") as handle:
        handle.write(json.dumps(entry, ensure_ascii=False) + "\n")
    return seq


def record(target: Path, note: str) -> Path | None:
    """Call right BEFORE creating or replacing `target`. Returns the backup path, if any."""
    folder = _folder(target)
    folder.mkdir(parents=True, exist_ok=True)
    backup: Path | None = None
    if target.exists():
        backup = folder / f"{len(_entries(folder)) + 1:05d}-{target.name}"
        shutil.copy2(target, backup)
    _append(
        folder,
        {
            "op": "write",
            "target": target.name,
            "backup": backup.name if backup else None,
            "note": note,
        },
    )
    return backup


def undo(target: Path) -> UndoOutcome:
    folder = _folder(target)
    entries = _entries(folder)
    undone = {e["undoes"] for e in entries if e.get("op") == "undo"}
    pending = [
        e
        for e in entries
        if e.get("op") == "write" and e["target"] == target.name and e["seq"] not in undone
    ]
    if not pending:
        raise SlimError(
            ErrorCode.BAD_ARGUMENT,
            f"Nothing to undo for '{target.name}'",
            hint="undo restores files written by write/clean/convert/merge in this folder.",
        )
    last = pending[-1]
    kept = folder / f"{len(entries) + 1:05d}-undone-{target.name}"
    try:
        if target.exists():
            shutil.copy2(target, kept)
        if last["backup"] is None:
            target.unlink(missing_ok=True)
        else:
            _atomic_copy(folder / last["backup"], target)
    except PermissionError as exc:
        # Nothing was restored, so leave no trace: the same undo must work once it is closed.
        kept.unlink(missing_ok=True)
        raise SlimError(
            ErrorCode.FILE_LOCKED,
            f"'{target.name}' is locked by another program; nothing was undone",
            hint=f"Close it in Excel, then retry undo(path='{target.name}').",
        ) from exc
    _append(folder, {"op": "undo", "undoes": last["seq"], "target": target.name, "kept": kept.name})
    return UndoOutcome(target.name, str(last["note"]), last["backup"], kept.name)


def _atomic_copy(source: Path, target: Path) -> None:
    fd, tmp_name = tempfile.mkstemp(prefix=".slim-", suffix=target.suffix, dir=target.parent)
    os.close(fd)
    tmp = Path(tmp_name)
    try:
        shutil.copy2(source, tmp)
        tmp.replace(target)
    finally:
        tmp.unlink(missing_ok=True)
