"""Native, refreshable Excel PivotTables written with openpyxl's pivot model (PLAN §11.9 P1).

The cache holds the source rows (shared items for row/column fields, numbers for value fields)
and is marked refreshOnLoad, so Excel rebuilds the layout from the live source range when the
file is opened. Supported: 1-2 row fields, an optional column field, 1-3 value fields with
sum | count | average | min | max.
"""

from __future__ import annotations

import datetime as dt
from collections.abc import Sequence
from dataclasses import dataclass
from typing import Any

from openpyxl.pivot.cache import (
    CacheDefinition,
    CacheField,
    CacheSource,
    SharedItems,
    WorksheetSource,
)
from openpyxl.pivot.fields import Index, Missing, Number, Text
from openpyxl.pivot.record import Record, RecordList
from openpyxl.pivot.table import (
    DataField,
    FieldItem,
    Location,
    PivotField,
    PivotTableStyle,
    RowColField,
    RowColItem,
    TableDefinition,
)

SUBTOTALS = {"sum": "sum", "count": "count", "avg": "average", "min": "min", "max": "max"}


@dataclass(frozen=True, slots=True)
class PivotSpec:
    headers: tuple[str, ...]
    rows: tuple[tuple[Any, ...], ...]  # source data rows (no header)
    source_sheet: str
    source_ref: str  # 'A1:E301' including the header row
    row_fields: tuple[int, ...]
    col_field: int | None
    value_fields: tuple[int, ...]
    agg: str
    name: str = "PivotTable1"


def _key(value: Any) -> str:
    if value is None or value == "":
        return ""
    if isinstance(value, dt.datetime | dt.date):
        return value.isoformat()
    return str(value)


def _uniques(spec: PivotSpec, col: int) -> list[str]:
    seen: dict[str, None] = {}
    for row in spec.rows:
        seen.setdefault(_key(row[col]), None)
    return sorted(seen, key=lambda v: (v == "", v.casefold()))


def build_pivot(spec: PivotSpec, cache_id: int, top: int, left: int) -> TableDefinition:
    """A pivot table anchored at (top, left), 0-based, with its cache attached."""
    numeric = {c for c in range(len(spec.headers)) if _numeric_column(spec, c)}
    axes = set(spec.row_fields) | ({spec.col_field} if spec.col_field is not None else set())
    # every text column is stored as shared items; number columns only when used as an axis
    items = {
        c: _uniques(spec, c) for c in range(len(spec.headers)) if c in axes or c not in numeric
    }
    cache = _cache(spec, items)
    cache.refreshedBy = (
        f"slim-sheet {spec.name}"  # keeps caches distinct (openpyxl dedupes equal ones)
    )
    pivot = TableDefinition(
        name=spec.name, cacheId=cache_id, dataCaption="Values", updatedVersion=6,
        minRefreshableVersion=3, createdVersion=6, outline=True, outlineData=True,
        useAutoFormatting=True, itemPrintTitles=True, indent=0, multipleFieldFilters=False,
        location=_location(spec, items, top, left),
    )  # fmt: skip
    pivot.cache = cache
    pivot.pivotFields = [_pivot_field(spec, col, items) for col in range(len(spec.headers))]
    pivot.rowFields = [RowColField(x=col) for col in spec.row_fields]
    pivot.rowItems = _row_items(spec, items)
    if spec.col_field is not None:
        pivot.colFields = [RowColField(x=spec.col_field)]
        pivot.colItems = [RowColItem(x=[Index(v=i)]) for i in range(len(items[spec.col_field]))]
        pivot.colItems.append(RowColItem(t="grand", x=[Index(v=0)]))
    elif len(spec.value_fields) > 1:
        pivot.colFields = [RowColField(x=-2)]  # the "Values" pseudo-field
        pivot.colItems = [RowColItem(i=i, x=[Index(v=i)]) for i in range(len(spec.value_fields))]
    else:
        pivot.colItems = [RowColItem()]
    word = {"avg": "Average"}.get(spec.agg, spec.agg.capitalize())
    pivot.dataFields = [
        DataField(
            name=f"{word} of {spec.headers[c]}",
            fld=c,
            subtotal=SUBTOTALS[spec.agg],
            baseField=0,
            baseItem=0,
        )
        for c in spec.value_fields
    ]
    pivot.pivotTableStyleInfo = PivotTableStyle(
        name="PivotStyleLight16", showRowHeaders=True, showColHeaders=True,
        showRowStripes=False, showColStripes=False, showLastColumn=True,
    )  # fmt: skip
    return pivot


def _cache(spec: PivotSpec, items: dict[int, list[str]]) -> CacheDefinition:
    fields = []
    for col, header in enumerate(spec.headers):
        if col in items:
            shared = SharedItems(_fields=[Text(v=v) if v else Missing() for v in items[col]],
                                 count=len(items[col]))  # fmt: skip
            if any(v == "" for v in items[col]):
                shared.containsBlank = True
        else:
            numbers = [float(r[col]) for r in spec.rows if _is_number(r[col])]
            shared = SharedItems(
                containsSemiMixedTypes=False, containsString=False, containsNumber=True,
                minValue=min(numbers, default=0.0), maxValue=max(numbers, default=0.0),
            )  # fmt: skip
        fields.append(CacheField(name=header or f"Column{col + 1}", sharedItems=shared, numFmtId=0))
    source = CacheSource(
        type="worksheet",
        worksheetSource=WorksheetSource(ref=spec.source_ref, sheet=spec.source_sheet),
    )
    cache = CacheDefinition(
        refreshOnLoad=True, saveData=True, createdVersion=6, refreshedVersion=6,
        minRefreshableVersion=3, recordCount=len(spec.rows), cacheSource=source, cacheFields=fields,
    )  # fmt: skip
    cache.records = RecordList(r=[_record(spec, row, items) for row in spec.rows])
    return cache


def _record(spec: PivotSpec, row: Sequence[Any], items: dict[int, list[str]]) -> Record:
    values: list[Any] = []
    for col in range(len(spec.headers)):
        if col in items:
            values.append(Index(v=items[col].index(_key(row[col]))))
        elif _is_number(row[col]):
            values.append(Number(v=float(row[col])))
        else:
            values.append(Missing())
    return Record(_fields=values)


def _numeric_column(spec: PivotSpec, col: int) -> bool:
    present = [r[col] for r in spec.rows if r[col] not in (None, "")]
    return bool(present) and all(_is_number(v) for v in present)


def _is_number(value: Any) -> bool:
    return isinstance(value, int | float) and not isinstance(value, bool)


def _pivot_field(spec: PivotSpec, col: int, items: dict[int, list[str]]) -> PivotField:
    if col in spec.row_fields or col == spec.col_field:
        axis = "axisRow" if col in spec.row_fields else "axisCol"
        entries = [FieldItem(x=i) for i in range(len(items[col]))] + [FieldItem(t="default")]
        return PivotField(axis=axis, showAll=False, items=entries)
    if col in spec.value_fields:
        return PivotField(dataField=True, showAll=False)
    return PivotField(showAll=False)


def _row_items(spec: PivotSpec, items: dict[int, list[str]]) -> list[RowColItem]:
    first = spec.row_fields[0]
    rows = [RowColItem(x=[Index(v=i)]) for i in range(len(items[first]))]
    rows.append(RowColItem(t="grand", x=[Index(v=0)]))
    return rows


def _location(spec: PivotSpec, items: dict[int, list[str]], top: int, left: int) -> Location:
    from openpyxl.utils import get_column_letter

    body_rows = len(items[spec.row_fields[0]]) + 1  # + grand total
    header_rows = 2 if spec.col_field is not None or len(spec.value_fields) > 1 else 1
    if spec.col_field is not None:
        width = len(items[spec.col_field]) + 2
    else:
        width = 1 + len(spec.value_fields)
    first = f"{get_column_letter(left + 1)}{top + 1}"
    last = f"{get_column_letter(left + width)}{top + header_rows + body_rows}"
    return Location(
        ref=f"{first}:{last}", firstHeaderRow=1, firstDataRow=header_rows, firstDataCol=1
    )


def layout_values(spec: PivotSpec) -> list[list[Any]]:
    """The grid Excel shows for this pivot (compact layout), so the file reads correctly before
    Excel refreshes it: header rows, one row per item, grand total row."""
    row_col = spec.row_fields[0]
    row_items = _uniques(spec, row_col)
    word = {"avg": "Average"}.get(spec.agg, spec.agg.capitalize())
    names = [f"{word} of {spec.headers[c]}" for c in spec.value_fields]
    if spec.col_field is not None:
        col_items = _uniques(spec, spec.col_field)
        grid: list[list[Any]] = [
            [names[0], "Column Labels"],
            ["Row Labels", *col_items, "Grand Total"],
        ]
        for item in [*row_items, None]:
            cells = [_aggregate(spec, spec.value_fields[0], item, c) for c in col_items]
            total = _aggregate(spec, spec.value_fields[0], item, None)
            grid.append([item if item is not None else "Grand Total", *cells, total])
        return grid
    head: list[list[Any]] = [["", "Values"]] if len(names) > 1 else []
    grid = [*head, ["Row Labels", *names]]
    for item in [*row_items, None]:
        label = item if item is not None else "Grand Total"
        grid.append([label, *(_aggregate(spec, v, item, None) for v in spec.value_fields)])
    return grid


def _aggregate(spec: PivotSpec, value_col: int, row_item: str | None, col_item: str | None) -> Any:
    """Excel's value for one cell of the pivot (None when the group is empty)."""
    rows = [
        r
        for r in spec.rows
        if (row_item is None or _key(r[spec.row_fields[0]]) == row_item)
        and (col_item is None or spec.col_field is None or _key(r[spec.col_field]) == col_item)
    ]
    if spec.agg == "count":
        return float(sum(1 for r in rows if r[value_col] not in (None, ""))) or None
    nums = [float(r[value_col]) for r in rows if _is_number(r[value_col])]
    if not nums:
        return None
    if spec.agg == "sum":
        return sum(nums)
    if spec.agg == "avg":
        return sum(nums) / len(nums)
    return min(nums) if spec.agg == "min" else max(nums)
