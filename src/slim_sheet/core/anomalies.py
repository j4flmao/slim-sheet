"""Statistical and structural anomaly checks over one table (PLAN §9). Pure, no I/O.

Kinds, most severe first: duplicate_key, type_mismatch, outlier, variant, date_range,
whitespace. Every finding points at a body cell so agents can jump to it.
"""

from __future__ import annotations

import datetime as dt
import re
import statistics
from collections import Counter, defaultdict
from collections.abc import Sequence
from dataclasses import dataclass

from slim_sheet.core.detect import kind_of
from slim_sheet.core.models import CellValue

KINDS = ("duplicate_key", "type_mismatch", "outlier", "variant", "date_range", "whitespace")
DOMINANT_SHARE = 0.8
MIN_VALUES = 5
MIN_NUMERIC = 8
ROBUST_Z = 3.5
MAX_CATEGORIES = 30
_ID_NAME = re.compile(
    r"(^|[\s_\-])(id|code|key|sku|no|number|invoice|ma|mã|so|số)([\s_\-.]|$)", re.I
)
_FAMILY = {"int": "number", "float": "number", "date": "date", "datetime": "date"}


@dataclass(frozen=True, slots=True)
class Anomaly:
    row: int
    col: int
    value: CellValue
    kind: str
    detail: str


def find_anomalies(
    columns: Sequence[str], body: Sequence[Sequence[CellValue]], only: Sequence[int] | None = None
) -> list[Anomaly]:
    found: list[Anomaly] = []
    for c in only if only is not None else range(len(columns)):
        cells = [(r, row[c]) for r, row in enumerate(body) if c < len(row) and _filled(row[c])]
        if not cells:
            continue
        found += _duplicate_keys(columns[c], c, cells)
        found += _type_mismatch(c, cells)
        found += _outliers(c, cells)
        found += _variants(c, cells)
        found += _date_range(c, cells)
        found += _whitespace(c, cells)
    order = {kind: i for i, kind in enumerate(KINDS)}
    return sorted(found, key=lambda a: (order[a.kind], a.col, a.row))


def _filled(value: CellValue) -> bool:
    return value is not None and value != ""


def _family(value: CellValue) -> str:
    kind = kind_of(value)
    return _FAMILY.get(kind, kind)


def _duplicate_keys(name: str, c: int, cells: list[tuple[int, CellValue]]) -> list[Anomaly]:
    keys = [_norm(v) for _, v in cells]
    unique_share = len(set(keys)) / len(keys)
    if len(cells) < MIN_VALUES or not (_ID_NAME.search(name) or unique_share >= 0.95):
        return []
    first: dict[object, int] = {}
    out = []
    for (r, value), key in zip(cells, keys, strict=True):
        if key in first:
            out.append(Anomaly(r, c, value, "duplicate_key", f"same as body row {first[key] + 1}"))
        else:
            first[key] = r
    return out


def _type_mismatch(c: int, cells: list[tuple[int, CellValue]]) -> list[Anomaly]:
    families = Counter(_family(v) for _, v in cells)
    dominant, count = families.most_common(1)[0]
    if len(cells) < MIN_VALUES or count == len(cells) or count / len(cells) < DOMINANT_SHARE:
        return []
    return [
        Anomaly(r, c, v, "type_mismatch", f"{_family(v)} in a mostly {dominant} column")
        for r, v in cells
        if _family(v) != dominant
    ]


def _outliers(c: int, cells: list[tuple[int, CellValue]]) -> list[Anomaly]:
    nums = [(r, float(v)) for r, v in cells if _family(v) == "number"]  # type: ignore[arg-type]
    if len(nums) < MIN_NUMERIC:
        return []
    values = [v for _, v in nums]
    median = statistics.median(values)
    mad = statistics.median(abs(v - median) for v in values)
    if mad > 0:
        return [
            Anomaly(
                r, c, v, "outlier", f"robust z={0.6745 * (v - median) / mad:.1f}, median {median:g}"
            )
            for r, v in nums
            if abs(0.6745 * (v - median) / mad) > ROBUST_Z
        ]
    q1, _, q3 = statistics.quantiles(values, n=4)
    iqr = q3 - q1
    if iqr <= 0:
        return []
    low, high = q1 - 3 * iqr, q3 + 3 * iqr
    return [
        Anomaly(r, c, v, "outlier", f"outside [{low:g}, {high:g}]")
        for r, v in nums
        if not low <= v <= high
    ]


def _variants(c: int, cells: list[tuple[int, CellValue]]) -> list[Anomaly]:
    texts = [(r, v) for r, v in cells if isinstance(v, str)]
    groups: dict[str, Counter[str]] = defaultdict(Counter)
    for _, v in texts:
        groups[_norm_text(v)][v] += 1
    if len(texts) < MIN_VALUES or len(groups) > MAX_CATEGORIES:
        return []
    canonical = {key: spellings.most_common(1)[0][0] for key, spellings in groups.items()}
    out = []
    for r, v in texts:
        main = canonical[_norm_text(v)]
        if v.strip() != main:  # pure whitespace differences are reported as 'whitespace'
            out.append(Anomaly(r, c, v, "variant", f"spelled differently from '{main}'"))
    return out


def _date_range(c: int, cells: list[tuple[int, CellValue]]) -> list[Anomaly]:
    return [
        Anomaly(r, c, v, "date_range", f"year {v.year} is implausible")
        for r, v in cells
        if isinstance(v, dt.date) and not 1900 <= v.year <= 2100
    ]


def _whitespace(c: int, cells: list[tuple[int, CellValue]]) -> list[Anomaly]:
    return [
        Anomaly(r, c, v, "whitespace", "leading/trailing spaces (breaks joins and lookups)")
        for r, v in cells
        if isinstance(v, str) and v != v.strip()
    ]


def _norm_text(text: str) -> str:
    return " ".join(text.split()).casefold()


def _norm(value: CellValue) -> object:
    return _norm_text(value) if isinstance(value, str) else value
