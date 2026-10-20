"""Parse numbers and dates written as text, the way messy exports write them. Pure."""

from __future__ import annotations

import datetime as dt
import re

_CURRENCY = "$€£¥₫₩₹"
_NUMBER_BODY = re.compile(r"^\d[\d.,\s']*$")  # \s also covers no-break spaces
_LEADING_ZERO_ID = re.compile(r"^0\d")
_TIME = r"(?:[ T](\d{1,2}):(\d{2})(?::(\d{2}))?)?"
_ISO = re.compile(r"^(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})" + _TIME + "$")
_NUMERIC_DMY = re.compile(r"^(\d{1,2})[-/.](\d{1,2})[-/.](\d{2}|\d{4})" + _TIME + "$")
_D_MON_Y = re.compile(r"^(\d{1,2})[- ]([A-Za-z]{3,9})\.?[- ,]+(\d{2}|\d{4})$")
_MON_D_Y = re.compile(r"^([A-Za-z]{3,9})\.? (\d{1,2}),? (\d{4})$")
_MONTHS = {
    m: i
    for i, names in enumerate(
        [("jan", "january"), ("feb", "february"), ("mar", "march"), ("apr", "april"),
         ("may",), ("jun", "june"), ("jul", "july"), ("aug", "august"),
         ("sep", "sept", "september"), ("oct", "october"), ("nov", "november"),
         ("dec", "december")],
        start=1,
    )
    for m in names
}  # fmt: skip


def parse_number(text: str) -> int | float | None:
    """'1,234.50' '1.234,50' '(120)' '-5%' '$ 1 200' -> number; IDs like '00123' -> None."""
    s = text.strip()
    if not s:
        return None
    negative = s.startswith("(") and s.endswith(")")
    s = s.strip("()").strip().strip(_CURRENCY).strip()
    if s.startswith("-"):
        negative, s = not negative, s[1:].strip()
    elif s.startswith("+"):
        s = s[1:].strip()
    percent = s.endswith("%")
    s = s.rstrip("%").strip().strip(_CURRENCY).strip()
    if not _NUMBER_BODY.match(s) or _LEADING_ZERO_ID.match(s):
        return None
    normalized = _normalize_separators(re.sub(r"[\s']", "", s))
    if normalized is None:
        return None
    value: int | float = float(normalized) if "." in normalized else int(normalized)
    if percent:
        value = value / 100
    return -value if negative else value


def _normalize_separators(s: str) -> str | None:
    if "," in s and "." in s:
        decimal = "," if s.rfind(",") > s.rfind(".") else "."
        thousands = "." if decimal == "," else ","
        s = s.replace(thousands, "").replace(decimal, ".")
    elif "," in s:
        s = s.replace(",", "") if re.fullmatch(r"\d{1,3}(,\d{3})+", s) else s.replace(",", ".")
    elif s.count(".") > 1:
        s = s.replace(".", "") if re.fullmatch(r"\d{1,3}(\.\d{3})+", s) else s
    return s if re.fullmatch(r"\d+(\.\d+)?", s) else None


def date_parts(text: str) -> tuple[int, int, int] | None:
    """(first, second, year) for an ambiguous numeric date like 03/04/2026, else None."""
    match = _NUMERIC_DMY.match(text.strip())
    if not match:
        return None
    return int(match.group(1)), int(match.group(2)), _year(match.group(3))


def parse_date(text: str, day_first: bool) -> dt.date | dt.datetime | None:
    s = text.strip()
    if match := _ISO.match(s):
        y, m, d = int(match.group(1)), int(match.group(2)), int(match.group(3))
        return _build(y, m, d, match.groups()[3:])
    if match := _NUMERIC_DMY.match(s):
        a, b, y = int(match.group(1)), int(match.group(2)), _year(match.group(3))
        d, m = (a, b) if day_first else (b, a)
        return _build(y, m, d, match.groups()[3:])
    if match := _D_MON_Y.match(s):
        month = _MONTHS.get(match.group(2).lower())
        return _build(_year(match.group(3)), month, int(match.group(1)), ()) if month else None
    if match := _MON_D_Y.match(s):
        month = _MONTHS.get(match.group(1).lower())
        return _build(int(match.group(3)), month, int(match.group(2)), ()) if month else None
    return None


def _year(raw: str) -> int:
    year = int(raw)
    if len(raw) == 2:
        year += 2000 if year < 70 else 1900
    return year


def _build(
    y: int, m: int, d: int, time_parts: tuple[str | None, ...]
) -> dt.date | dt.datetime | None:
    try:
        if time_parts and time_parts[0] is not None:
            hh, mm, ss = (int(p) if p else 0 for p in time_parts)
            return dt.datetime(y, m, d, hh, mm, ss)
        return dt.date(y, m, d)
    except ValueError:
        return None
