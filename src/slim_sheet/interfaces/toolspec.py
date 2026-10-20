"""Single source of truth for tool names, tiers and descriptions (PLAN §5, ADR-5)."""

from __future__ import annotations

from dataclasses import dataclass

from slim_sheet.config import TIERS, Settings
from slim_sheet.core.errors import ErrorCode, SlimError

TIER_CAPS = {"lite": 8, "standard": 18, "full": 30}


@dataclass(frozen=True, slots=True)
class ToolSpec:
    name: str
    tier: str
    description: str
    read_only: bool = True


TOOLS: tuple[ToolSpec, ...] = (
    ToolSpec(
        "overview",
        "lite",
        "Map a spreadsheet file: sheets, used ranges, header row, column names and types. "
        "Call this first. Example: overview(path='sales.xlsx')",
    ),
    ToolSpec(
        "read",
        "lite",
        "Read cells as a table with A1 coordinates. Optional sheet, range like 'A1:F50', "
        "cursor from a previous read, format markdown|tsv|json.",
    ),
    ToolSpec(
        "find",
        "lite",
        "Find cells containing text (or a regex) across sheets; returns Sheet!A1 addresses. "
        "Example: find(path='sales.xlsx', text='refund')",
    ),
    ToolSpec(
        "query",
        "lite",
        "Run one read-only SQL SELECT; tables are sheet names, quote names with spaces. "
        """Example: sql='SELECT "Region", sum("Total") FROM "Orders" GROUP BY 1'""",
    ),
    ToolSpec(
        "profile",
        "lite",
        "Column statistics for a sheet: type, nulls, distinct, min, max, mean, median, top values. "
        "Optional columns list.",
    ),
    ToolSpec(
        "write",
        "lite",
        "Write a 2-D values array starting at range's top-left cell, e.g. values=[['Total', 120]]. "
        "Saves to <name>.slim.xlsx unless overwrite=true.",
        read_only=False,
    ),
    ToolSpec(
        "explain_formula",
        "lite",
        "Explain and check an Excel formula: steps, references, syntax/arity errors. "
        "Pass formula='=...' or path plus cell like 'Sheet1!C7'.",
    ),
    ToolSpec(
        "clean",
        "standard",
        "Clean a table: ops from trim, collapse_spaces, remove_nonprintable, unify_null, "
        "numbers_from_text, normalize_dates, drop_empty_rows, drop_empty_cols, dedupe. "
        "Previews unless apply=true.",
        read_only=False,
    ),
    ToolSpec(
        "anomalies",
        "standard",
        "Find suspicious cells: duplicate keys, wrong types, outliers, spelling variants, "
        "bad dates, stray spaces. Returns Sheet!A1 addresses.",
    ),
    ToolSpec(
        "merge",
        "standard",
        "Join two tables on keys (how=left|inner|full|anti) and count unmatched rows on both "
        "sides. keys=['ID'] or ['ID=Order No']. out_path saves the result.",
        read_only=False,
    ),
    ToolSpec(
        "convert",
        "standard",
        "Convert a file to xlsx, csv, tsv, parquet or json (by out_path extension). "
        "Reads xls, xlsb, ods, csv and more.",
        read_only=False,
    ),
    ToolSpec(
        "diff",
        "standard",
        "Compare two files cell by cell, or row by row with key='ID'. "
        "Example: diff(path='a.xlsx', other_path='a.slim.xlsx')",
    ),
    ToolSpec(
        "undo",
        "standard",
        "Undo the last slim-sheet write to a file (write, clean, convert or merge output).",
        read_only=False,
    ),
    ToolSpec(
        "validate_formula",
        "standard",
        "Check a formula before writing it: syntax, unknown functions, Excel version "
        "(target='2016' or 'xls'), missing sheets/names. Gives fixes.",
    ),
    ToolSpec(
        "evaluate",
        "standard",
        "Calculate a formula offline, alone or against a workbook. "
        "Example: evaluate(path='x.xlsx', cell='Calc!C7') or evaluate(formula='=ROUND(2.5,0)')",
    ),
    ToolSpec(
        "describe_cell",
        "standard",
        "Everything about one cell: value, formula, number format, font, fill, comment, "
        "validation, merge. Example: describe_cell(path='x.xlsx', cell='Sheet1!B2')",
    ),
    ToolSpec(
        "create_workbook",
        "full",
        "Create a new empty .xlsx with the given sheet names. Example: "
        "create_workbook(path='out.xlsx', sheets=['Data','Summary'])",
        read_only=False,
    ),
    ToolSpec(
        "sheet_ops",
        "full",
        "Add, rename, delete, copy, move, hide or unhide a sheet (op=...). Renames also fix "
        "formulas that point to the sheet.",
        read_only=False,
    ),
    ToolSpec(
        "rows_ops",
        "full",
        "Insert or delete rows/columns (op=insert_rows|delete_rows|insert_cols|delete_cols, "
        "at='5' or 'C', count). Formulas, merges, names are shifted like Excel.",
        read_only=False,
    ),
    ToolSpec(
        "name_ops",
        "full",
        "List, add or delete defined names. Example: name_ops(path, op='add', name='TaxRate', "
        "ref='Rates!B2')",
        read_only=False,
    ),
    ToolSpec(
        "format_range",
        "full",
        "Style a range: bold, italic, font_color, fill_color ('yellow' or '#FFEB9C'), "
        "number_format ('currency', 'percent', 'date' or custom), align, wrap, border, autofit.",
        read_only=False,
    ),
    ToolSpec(
        "set_table",
        "full",
        "Turn a range with a header row into a named Excel table. Example: set_table(path, "
        "sheet='Data', range='A1:D50', name='Sales')",
        read_only=False,
    ),
    ToolSpec(
        "add_validation",
        "full",
        "Add a dropdown or rule: type=list (values=[...]), whole, decimal, date, text_length "
        "(min/max) or custom (formula).",
        read_only=False,
    ),
    ToolSpec(
        "conditional_format",
        "full",
        "Highlight cells by rule: greater_than, less_than, between, equal, text_contains, "
        "duplicates, top_n, color_scale, data_bar, formula.",
        read_only=False,
    ),
    ToolSpec(
        "add_chart",
        "full",
        "Insert a chart (column, bar, line, pie, area, scatter) from data_range: header row + "
        "category column + series columns.",
        read_only=False,
    ),
    ToolSpec(
        "pivot",
        "full",
        "Summarise a table: rows=['Region'], values=['Total'], agg=sum|count|avg|min|max, "
        "optional columns='Month'. Save with out_path or to_sheet.",
        read_only=False,
    ),
    ToolSpec(
        "split",
        "full",
        "Split a table by a column's values into sheets of one workbook (mode=sheets) or separate"
        " files (mode=files).",
        read_only=False,
    ),
    ToolSpec(
        "recalc",
        "full",
        "Recalculate every formula with LibreOffice so cached values are fresh (needs LibreOffice"
        " installed).",
        read_only=False,
    ),
)


def enabled_tools(settings: Settings) -> list[ToolSpec]:
    if settings.tools:
        known = {t.name: t for t in TOOLS}
        unknown = [name for name in settings.tools if name not in known]
        if unknown:
            raise SlimError(
                ErrorCode.BAD_ARGUMENT,
                f"Unknown tool(s) in SLIM_SHEET_TOOLS: {', '.join(unknown)}",
                hint=f"Available: {', '.join(known)}",
            )
        return [known[name] for name in settings.tools]
    allowed = TIERS[: TIERS.index(settings.tier) + 1]
    return [t for t in TOOLS if t.tier in allowed]
