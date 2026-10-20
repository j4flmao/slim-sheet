"""Small vocabularies that let agents style cells with plain words. Pure."""

from __future__ import annotations

import re

from slim_sheet.core.errors import ErrorCode, SlimError

COLORS = {
    "black": "000000",
    "white": "FFFFFF",
    "red": "FF0000",
    "green": "00B050",
    "blue": "0070C0",
    "yellow": "FFFF00",
    "orange": "FFC000",
    "gray": "808080",
    "grey": "808080",
    "light_red": "FFC7CE",
    "light_green": "C6EFCE",
    "light_yellow": "FFEB9C",
    "light_blue": "DDEBF7",
    "light_gray": "F2F2F2",
}
NUMBER_FORMATS = {
    "integer": "0",
    "decimal": "0.00",
    "thousands": "#,##0",
    "currency": '"$"#,##0.00',
    "vnd": '#,##0 "₫"',
    "eur": '#,##0.00 "€"',
    "percent": "0.00%",
    "date": "yyyy-mm-dd",
    "datetime": "yyyy-mm-dd hh:mm",
    "time": "hh:mm:ss",
    "text": "@",
}
ALIGNMENTS = ("left", "center", "right")
BORDERS = ("none", "thin", "medium", "thick", "outline")
_HEX = re.compile(r"^#?([0-9A-Fa-f]{6})$")


def color(value: str) -> str:
    """'red' / '#ff0000' / 'FF0000' -> 'FFFF0000' (ARGB as openpyxl stores it)."""
    key = value.strip().lower().replace(" ", "_")
    if key in COLORS:
        return "FF" + COLORS[key]
    match = _HEX.match(value.strip())
    if match:
        return "FF" + match.group(1).upper()
    raise SlimError(
        ErrorCode.BAD_ARGUMENT,
        f"Unknown color '{value}'",
        hint=f"Use #RRGGBB or one of: {', '.join(COLORS)}",
    )


def number_format(value: str) -> str:
    return NUMBER_FORMATS.get(value.strip().lower(), value)


def choice(value: str, options: tuple[str, ...], label: str) -> str:
    key = value.strip().lower()
    if key not in options:
        raise SlimError(
            ErrorCode.BAD_ARGUMENT,
            f"Unknown {label} '{value}'",
            hint=f"Use one of: {', '.join(options)}",
        )
    return key
