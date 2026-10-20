"""Closed error vocabulary. Every error carries a hint with a corrected call."""

from __future__ import annotations

import json
from enum import StrEnum


class ErrorCode(StrEnum):
    FILE_NOT_FOUND = "file_not_found"
    PATH_DENIED = "path_denied"
    UNSUPPORTED_FORMAT = "unsupported_format"
    MISSING_EXTRA = "missing_extra"
    SHEET_NOT_FOUND = "sheet_not_found"
    BAD_RANGE = "bad_range"
    BAD_CURSOR = "bad_cursor"
    BAD_ARGUMENT = "bad_argument"
    TOO_LARGE = "too_large"
    READ_FAILED = "read_failed"
    FILE_LOCKED = "file_locked"
    BAD_SQL = "bad_sql"
    SQL_WRITE_DENIED = "sql_write_denied"
    WRITE_CONFLICT = "write_conflict"
    FORMULA_DENIED = "formula_denied"
    FORMULA_SYNTAX = "formula_syntax"
    UNSUPPORTED_FUNCTION = "unsupported_function"
    CIRCULAR_REFERENCE = "circular_reference"
    MISSING_DEPENDENCY = "missing_dependency"
    ENCRYPTED = "encrypted"


class SlimError(Exception):
    """Domain error rendered verbatim to agents."""

    def __init__(self, code: ErrorCode, message: str, hint: str = "") -> None:
        super().__init__(message)
        self.code = code
        self.message = message
        self.hint = hint

    def to_dict(self) -> dict[str, str]:
        payload = {"error": self.code.value, "message": self.message}
        if self.hint:
            payload["hint"] = self.hint
        return payload

    def to_json(self) -> str:
        return json.dumps(self.to_dict(), ensure_ascii=False)
