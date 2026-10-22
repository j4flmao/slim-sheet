"""Inference and compact specs that let a small model describe a report in a few words. Pure.

- infer_format: header + sample values -> a named number format (core/design.py) and why
- choose_chart: intent + data shape -> chart kind, sort order and orientation
- parse_kpi / parse_chart: "Revenue=sum:Total|currency", "rank:Region:Total"
- aggregate / insight_title: the numbers behind a chart and a title that states its message
"""

from __future__ import annotations

import datetime as dt
import re
import statistics
from collections.abc import Sequence
from dataclasses import dataclass
from typing import Any

from slim_sheet.core.errors import ErrorCode, SlimError

_WORD = re.compile(r"[a-z0-9%]+")
_PERCENT_WORDS = {"%", "pct", "percent", "percentage", "margin", "rate", "share", "yield",
                  "conversion", "churn", "utilisation", "utilization", "ctr", "ratio"}  # fmt: skip
_GROWTH_WORDS = {"growth", "change", "yoy", "mom", "qoq", "delta", "variance", "var"}
_MULTIPLE_WORDS = {"multiple", "x", "pe"}
_YEAR_WORDS = {"year", "fy", "yr"}
_COUNT_WORDS = {"count", "companies", "customers", "orders", "items", "units", "makers", "listed",
                "employees", "headcount", "qty", "quantity", "number", "rows"}  # fmt: skip
_AMOUNT_WORDS = {"revenue", "income", "profit", "sales", "cost", "costs", "amount", "value",
                 "price", "total", "cap", "usd", "eur", "vnd", "bn", "m", "k", "spend", "budget",
                 "balance", "margin", "ebitda", "net", "gross"}  # fmt: skip
_ID_WORDS = {"id", "code", "zip", "postal", "postcode", "phone", "ticker", "sku", "ref"}
_MONTHS = ("jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "oct", "nov", "dec")
_PERIOD = re.compile(
    r"^(?:(?:19|20)\d{2}(?:[-/ ]?(?:\d{1,2}|q[1-4]|h[12]))?"  # 2024, 2024-03, 2024 Q1
    r"|q[1-4](?:[ -]\d{2,4})?|fy ?\d{2,4}|w\d{1,2})$",  # Q1 2024, FY25, W12
    re.IGNORECASE,
)


def _words(header: str) -> set[str]:
    text = header.lower()
    words = set(_WORD.findall(text))
    return words | ({"%"} if "%" in text else set())


def _from_header(words: set[str]) -> tuple[str | None, str]:
    """Formula columns have no values to look at (they are computed in Excel): use the header."""
    if words & _ID_WORDS:
        return None, "formula column, identifier header"
    if words & _GROWTH_WORDS:
        return "growth", "formula column, growth/change header"
    if words & _PERCENT_WORDS:
        return "percent", "formula column, rate/share header"
    if words & _MULTIPLE_WORDS:
        return "multiple", "formula column, multiple header"
    if words & _COUNT_WORDS:
        return "integer", "formula column, count header"
    if words & _AMOUNT_WORDS:
        return "number1", "formula column, amount header"
    return None, "formula column, header gives no number type"


def _numbers(values: Sequence[object]) -> list[float]:
    return [float(v) for v in values if isinstance(v, int | float) and not isinstance(v, bool)]


def infer_format(header: str, values: Sequence[object]) -> tuple[str | None, str]:
    """Pick a named format for a column; None means 'leave as is' (text, ids, empty)."""
    present = [v for v in values if v is not None and v != ""]
    if not present:
        return None, "empty"
    formulas = [v for v in present if isinstance(v, str) and v.startswith("=")]
    if len(formulas) >= len(present) * 0.8:
        return _from_header(_words(header))
    if all(isinstance(v, dt.date | dt.datetime) for v in present):
        return "date", "dates"
    nums = _numbers(present)
    if len(nums) < len(present) * 0.8:
        return None, "text"
    words = _words(header)
    if words & _ID_WORDS:
        return None, "identifier"
    biggest = max(abs(n) for n in nums)
    typical = statistics.median(abs(n) for n in nums)  # one outlier must not flip the column
    whole = all(n.is_integer() for n in nums)
    if words & _YEAR_WORDS and whole and all(1900 <= n <= 2100 for n in nums):
        return "year", "year numbers"
    if words & _GROWTH_WORDS and typical <= 10:
        return "growth", "growth/change header with fractions"
    if words & _PERCENT_WORDS and typical <= 1.5:
        return "percent", "rate/share header with fractions"
    if words & _PERCENT_WORDS:
        return "number1", "rate header but values look like points, not fractions"
    if words & _MULTIPLE_WORDS:
        return "multiple", "multiple header"
    if whole or biggest >= 1000:
        return "integer", "whole or large numbers"
    return ("number2" if biggest < 10 else "number1"), "small decimals"


# --- charts -------------------------------------------------------------------------------

INTENTS = ("auto", "compare", "rank", "trend", "share", "relationship")
MAX_BAR_LABEL = 12  # longer category labels read better on horizontal bars
MAX_SLICES = 5


@dataclass(frozen=True, slots=True)
class ChartChoice:
    kind: str  # column | bar | line | doughnut | scatter
    sort: str  # desc | asc | none
    reason: str


def looks_like_time(categories: Sequence[object]) -> bool:
    present = [c for c in categories if c is not None and c != ""]
    if not present:
        return False
    if all(isinstance(c, dt.date | dt.datetime) for c in present):
        return True
    if all(isinstance(c, int) and 1900 <= c <= 2100 for c in present):
        return True
    texts = [str(c).strip().lower() for c in present]
    return all(_PERIOD.match(t) or t[:3] in _MONTHS for t in texts)


def choose_chart(intent: str, categories: Sequence[object], values: Sequence[float]) -> ChartChoice:
    key = intent.strip().lower()
    if key not in INTENTS:
        raise SlimError(
            ErrorCode.BAD_ARGUMENT,
            f"Unknown chart intent '{intent}'",
            hint=f"Use one of: {', '.join(INTENTS)}, e.g. 'rank:Region:Total'",
        )
    if key == "auto":
        key = "trend" if looks_like_time(categories) else "compare"
    long_labels = any(len(str(c)) > MAX_BAR_LABEL for c in categories)
    if key == "trend":
        kind = "column" if len(categories) <= 3 else "line"
        return ChartChoice(kind, "asc", "time categories, shown in order")
    if key == "relationship":
        return ChartChoice("scatter", "asc", "two measures against each other")
    if key == "share":
        if 2 <= len(categories) <= MAX_SLICES and all(v >= 0 for v in values):
            return ChartChoice("doughnut", "desc", f"{len(categories)} parts of a whole")
        return ChartChoice(
            "bar", "desc", f"share of {len(categories)} parts: sorted bars read better"
        )
    if key == "rank" or long_labels or len(categories) > 8:
        return ChartChoice("bar", "desc", "ranked categories on horizontal bars")
    return ChartChoice("column", "desc", "a few short categories compared")


# --- compact specs ------------------------------------------------------------------------

KPI_AGGS = ("sum", "avg", "count", "min", "max", "ratio")
CHART_AGGS = ("sum", "avg", "count", "min", "max")
_KPI_EXAMPLE = "e.g. 'Revenue=sum:Total|currency', 'Orders=count', 'Margin=ratio:Profit/Revenue'"


@dataclass(frozen=True, slots=True)
class KpiSpec:
    label: str
    agg: str  # one of KPI_AGGS, or 'formula'
    column: str = ""
    column2: str = ""  # denominator for ratio
    formula: str = ""
    fmt: str = ""  # named format, or '' to infer


@dataclass(frozen=True, slots=True)
class ChartSpec:
    intent: str
    category: str
    value: str  # a column name, or 'count' for row counts
    agg: str = "sum"
    title: str = ""


def parse_kpi(item: str) -> KpiSpec:
    """'Label=agg:Column|fmt', 'Label=ratio:A/B', 'Label==SUM(...)', 'Orders=count', 'Revenue'."""
    body, _, fmt = item.partition("|")
    label, eq, rest = body.partition("=")
    label, rest, fmt = label.strip(), rest.strip(), fmt.strip().lower()
    if not label:
        raise SlimError(ErrorCode.BAD_ARGUMENT, f"KPI '{item}' has no label", hint=_KPI_EXAMPLE)
    if not eq:
        return KpiSpec(label, "sum", column=label, fmt=fmt)
    if rest.startswith("="):
        return KpiSpec(label, "formula", formula=rest, fmt=fmt)
    agg, _, column = rest.partition(":")
    agg = agg.strip().lower()
    if agg not in KPI_AGGS:
        raise SlimError(
            ErrorCode.BAD_ARGUMENT,
            f"KPI '{item}': unknown aggregate '{agg}'",
            hint=f"Use one of: {', '.join(KPI_AGGS)}; {_KPI_EXAMPLE}",
        )
    column = column.strip()
    if agg == "ratio":
        top, slash, bottom = column.partition("/")
        if not slash or not top.strip() or not bottom.strip():
            raise SlimError(
                ErrorCode.BAD_ARGUMENT,
                f"KPI '{item}': ratio needs 'Numerator/Denominator'",
                hint="e.g. 'Margin=ratio:Profit/Revenue|percent'",
            )
        return KpiSpec(label, agg, top.strip(), bottom.strip(), fmt=fmt or "percent")
    if not column and agg != "count":
        raise SlimError(
            ErrorCode.BAD_ARGUMENT, f"KPI '{item}': {agg} needs a column", hint=_KPI_EXAMPLE
        )
    return KpiSpec(label, agg, column, fmt=fmt)


def parse_chart(item: str) -> ChartSpec:
    """'intent:Category:Value[:agg][:Title]'; Value may be 'count'."""
    parts = [p.strip() for p in item.split(":")]
    if len(parts) < 3 or not parts[1] or not parts[2]:
        raise SlimError(
            ErrorCode.BAD_ARGUMENT,
            f"Chart '{item}' needs intent:Category:Value",
            hint="e.g. 'trend:Month:Revenue', 'rank:Region:Total', 'share:Product:Units:sum'",
        )
    intent = parts[0].lower() or "auto"
    if intent not in INTENTS:
        raise SlimError(
            ErrorCode.BAD_ARGUMENT,
            f"Chart '{item}': unknown intent '{parts[0]}'",
            hint=f"Use one of: {', '.join(INTENTS)}",
        )
    agg = "sum"
    rest = item.split(":")[3:]  # unstripped, so a title keeps its own spacing
    if rest and rest[0].strip().lower() in CHART_AGGS:
        agg, rest = rest[0].strip().lower(), rest[1:]
    title = ":".join(rest).strip()
    value = parts[2]
    if value.lower() == "count":
        value, agg = "count", "count"
    return ChartSpec(intent, parts[1], value, agg, title)


# --- numbers behind a chart ---------------------------------------------------------------


def aggregate(pairs: Sequence[tuple[object, object]], agg: str) -> list[tuple[object, float]]:
    """Group (category, value) pairs in first-seen order; blanks in category are skipped."""
    groups: dict[object, list[float]] = {}
    for category, value in pairs:
        if category is None or category == "":
            continue
        bucket = groups.setdefault(category, [])
        if agg == "count":
            bucket.append(1.0)
        elif isinstance(value, int | float) and not isinstance(value, bool):
            bucket.append(float(value))
    return [(category, reduce_values(values, agg)) for category, values in groups.items()]


def reduce_values(values: Sequence[float], agg: str) -> float:
    """sum | count | avg | min | max of a list of numbers; 0.0 for an empty list."""
    if not values:
        return 0.0
    if agg in ("sum", "count"):
        return float(sum(values))
    if agg == "avg":
        return float(sum(values)) / len(values)
    if agg == "min":
        return float(min(values))
    if agg == "max":
        return float(max(values))
    raise SlimError(
        ErrorCode.BAD_ARGUMENT, f"Unknown aggregate '{agg}'", hint="sum|count|avg|min|max"
    )


def order(points: list[tuple[object, float]], sort: str) -> list[tuple[object, float]]:
    if sort == "desc":
        return sorted(points, key=lambda p: -p[1])
    if sort == "asc":
        try:
            return sorted(points, key=lambda p: _time_key(p[0]))
        except TypeError:
            return points
    return points


def _time_key(category: object) -> tuple[int, Any]:
    if isinstance(category, str) and category[:3].lower() in _MONTHS:
        return (0, _MONTHS.index(category[:3].lower()))
    return (1, category)


def insight_title(intent: str, label: str, points: Sequence[tuple[object, float]]) -> str:
    """A title that states the message ('North leads Total with 41% of the total')."""
    if not points:
        return label
    if intent == "trend" and len(points) >= 2:
        first, last = points[0], points[-1]
        if first[1]:
            change = last[1] / first[1] - 1
            verb = "rose" if change >= 0 else "fell"
            return f"{label} {verb} {abs(change):.0%} from {first[0]} to {last[0]}"
        return f"{label} from {first[0]} to {last[0]}"
    top = max(points, key=lambda p: p[1])
    total = sum(p[1] for p in points)
    if total > 0 and all(p[1] >= 0 for p in points):
        return f"{top[0]} leads {label} with {top[1] / total:.0%} of the total"
    return f"{top[0]} has the highest {label}"
