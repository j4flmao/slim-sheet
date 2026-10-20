"""Process-wide LRU cache of loaded Arrow tables, keyed by file identity (PLAN §11.2 perf).

A long-running MCP server answers many calls on the same workbook; parsing it into
Arrow once and re-registering the tables (zero-copy) makes repeat queries cheap.
The key includes mtime and size, so an edited file is reloaded automatically.
"""

from __future__ import annotations

import threading
from collections import OrderedDict
from dataclasses import dataclass
from pathlib import Path

import pyarrow as pa

FileKey = tuple[str, int, int]


@dataclass(frozen=True, slots=True)
class LoadedTable:
    name: str
    source: str
    table: pa.Table


def file_key(path: Path) -> FileKey:
    stat = path.stat()
    return (str(path), stat.st_mtime_ns, stat.st_size)


class TableCache:
    def __init__(self, max_files: int = 4) -> None:
        self._max_files = max_files
        self._files: OrderedDict[FileKey, dict[str, list[LoadedTable]]] = OrderedDict()
        self._lock = threading.Lock()

    def get(self, key: FileKey, sheet: str) -> list[LoadedTable] | None:
        with self._lock:
            sheets = self._files.get(key)
            if sheets is None or sheet not in sheets:
                return None
            self._files.move_to_end(key)
            return sheets[sheet]

    def put(self, key: FileKey, sheet: str, tables: list[LoadedTable]) -> None:
        with self._lock:
            stale = [k for k in self._files if k[0] == key[0] and k != key]
            for k in stale:
                del self._files[k]
            self._files.setdefault(key, {})[sheet] = tables
            self._files.move_to_end(key)
            while len(self._files) > self._max_files:
                self._files.popitem(last=False)

    def clear(self) -> None:
        with self._lock:
            self._files.clear()


TABLE_CACHE = TableCache()
