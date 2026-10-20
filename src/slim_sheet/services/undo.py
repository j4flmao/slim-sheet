"""`undo` use-case: restore a file to its state before the last slim-sheet write."""

from __future__ import annotations

from slim_sheet.adapters import journal
from slim_sheet.adapters.sandbox import resolve_output
from slim_sheet.config import Settings


def undo(path: str, settings: Settings) -> journal.UndoOutcome:
    # resolve_output: the file may legitimately be gone (e.g. undoing twice past its creation)
    return journal.undo(resolve_output(path, settings))
