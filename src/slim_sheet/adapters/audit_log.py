"""Opt-in per-call audit log (PLAN §11.9): one JSON line per tool call.

Enabled with SLIM_SHEET_AUDIT_LOG=<file>. Records who (SLIM_SHEET_AUDIT_USER or the OS user),
which tool, the file/sheet/range it touched, outcome, error code and duration. Cell values,
formulas, SQL text and anything secret-like are never logged — only their sizes — so the log
can be shipped to a SIEM without leaking spreadsheet contents.
"""

from __future__ import annotations

import datetime as dt
import getpass
import json
import logging
import threading
from pathlib import Path
from typing import Any

LOG = logging.getLogger(__name__)
_LOCK = threading.Lock()
# Arguments that locate the work; everything else is reduced to its size.
LOCATING = {
    "path", "other_path", "out_path", "sheet", "range", "cell", "data_range", "to", "op", "kind",
    "type", "preset", "scope", "trace", "overwrite", "brand", "mode", "anchor",
}  # fmt: skip


def redact(arguments: dict[str, Any]) -> dict[str, Any]:
    """Keep locating arguments; replace content (values, sql, formulas, text) by its size."""
    safe: dict[str, Any] = {}
    for key, value in arguments.items():
        if value is None:
            continue
        if key in LOCATING and isinstance(value, str | int | bool):
            safe[key] = value
        elif isinstance(value, list):
            safe[key] = f"<{len(value)} item(s)>"
        else:
            safe[key] = f"<{len(str(value))} chars>"
    return safe


def principal(configured: str | None) -> str:
    if configured:
        return configured
    try:
        return getpass.getuser()
    except (OSError, KeyError):  # no login name (some containers)
        return "unknown"


def record(path: Path, entry: dict[str, Any]) -> None:
    """Append one line; an unwritable log is reported on stderr but never breaks a tool call."""
    entry = {"ts": dt.datetime.now(dt.UTC).isoformat(timespec="milliseconds"), **entry}
    line = json.dumps(entry, ensure_ascii=False, default=str)
    try:
        with _LOCK, path.open("a", encoding="utf-8") as handle:
            handle.write(line + "\n")
    except OSError as exc:
        LOG.error("Failed to write audit log %s for tool %s: %s", path, entry.get("tool"), exc)
