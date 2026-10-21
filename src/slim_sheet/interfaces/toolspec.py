"""Single source of truth for tool names, tiers, descriptions and annotations (PLAN §5, §17).

Annotations follow the MCP spec defaults, which assume the worst (destructive, not idempotent,
open world), so every tool states them explicitly: writes go to a journaled working copy and
are not destructive unless they can delete content (sheet_ops, range_ops).
"""

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
    title: str = ""
    destructive: bool = False
    idempotent: bool = True


def _w(name: str, tier: str, description: str, title: str, **flags: bool) -> ToolSpec:
    """A tool that writes: idempotent unless said otherwise."""
    return ToolSpec(name, tier, description, False, title, **flags)


TOOLS: tuple[ToolSpec, ...] = (
    # --- lite -------------------------------------------------------------------------------
    ToolSpec(
        "overview",
        "lite",
        "Map a spreadsheet file (or list the spreadsheets in a folder): sheets, used ranges, "
        "header row, columns and types. Call this first.",
        title="Workbook overview",
    ),
    ToolSpec(
        "read",
        "lite",
        "Read cells as a table with A1 coordinates (sheet, range 'A1:F50', cursor). "
        "sheet='#vba' lists macros; range='Module1' shows its code.",
        title="Read cells",
    ),
    ToolSpec(
        "find",
        "lite",
        "Find cells containing text (or a regex) across sheets; returns Sheet!A1 addresses. "
        "Example: find(path='sales.xlsx', text='refund')",
        title="Find cells",
    ),
    ToolSpec(
        "query",
        "lite",
        "Run one read-only SQL SELECT; tables are sheet names, quote names with spaces. "
        """Example: sql='SELECT "Region", sum("Total") FROM "Orders" GROUP BY 1'""",
        title="SQL query",
    ),
    ToolSpec(
        "profile",
        "lite",
        "Column statistics for a sheet: type, nulls, distinct, min, max, mean, median, top values. "
        "Optional columns list.",
        title="Profile columns",
    ),
    _w(
        "write",
        "lite",
        "Write a 2-D values array starting at range's top-left cell, e.g. values=[['Total', 120]]. "
        "Saves to <name>.slim.xlsx unless overwrite=true.",
        "Write cells",
    ),
    ToolSpec(
        "explain_formula",
        "lite",
        "Explain and check an Excel formula: steps, references, syntax/arity errors. "
        "Pass formula='=...' or path plus cell like 'Sheet1!C7'.",
        title="Explain formula",
    ),
    _w(
        "build_report",
        "lite",
        "Build a branded dashboard sheet from a data sheet in one call. Example: "
        "kpis=['Revenue=sum:Total','Orders=count'], charts=['trend:Month:Total']",
        "Build dashboard",
    ),
    # --- standard ---------------------------------------------------------------------------
    _w(
        "clean",
        "standard",
        "Clean a table: ops from trim, collapse_spaces, remove_nonprintable, unify_null, "
        "numbers_from_text, normalize_dates, drop_empty_rows, drop_empty_cols, dedupe. "
        "Previews unless apply=true.",
        "Clean table",
    ),
    ToolSpec(
        "anomalies",
        "standard",
        "Find suspicious cells (duplicates, types, outliers, typos); scope='formulas' audits "
        "models; scope='macros' rates VBA risks (auto-run, Shell, downloads).",
        title="Find anomalies",
    ),
    _w(
        "merge",
        "standard",
        "Join two tables on keys (how=left|inner|full|anti), or how=reconcile to label rows "
        "matched/different/only_left/only_right. keys=['ID'] or ['ID=Order No'].",
        "Join or reconcile tables",
    ),
    _w(
        "convert",
        "standard",
        "Convert to xlsx/csv/tsv/parquet/json by out_path extension; CSV: delimiter=';', "
        "encoding='cp1252', decimal=','. split_by='Region' splits; a folder exports VBA.",
        "Convert or split",
    ),
    ToolSpec(
        "diff",
        "standard",
        "Compare two files cell by cell, or row by row with key='ID'. "
        "Example: diff(path='a.xlsx', other_path='a.slim.xlsx')",
        title="Compare files",
    ),
    _w(
        "undo",
        "standard",
        "Undo the last slim-sheet write to a file (every editing tool is journaled).",
        "Undo last write",
        idempotent=False,
    ),
    ToolSpec(
        "evaluate",
        "standard",
        "Calculate a formula offline, or check it with check_only=true (target='2016'). Example: "
        "evaluate(path='x.xlsx', cell='Calc!C7')",
        title="Evaluate or check formula",
    ),
    ToolSpec(
        "describe_cell",
        "standard",
        "Everything about one cell: value, formula, format, note, validation; trace='precedents' "
        "or 'dependents' follows formulas across sheets.",
        title="Describe and trace cell",
    ),
    _w(
        "style_sheet",
        "standard",
        "Make a table look professional in one call: header, number formats (auto), totals, "
        "widths, freeze. preset=report|data|finance|striped|minimal.",
        "Style table",
    ),
    ToolSpec(
        "what_if",
        "standard",
        "Change inputs in memory and see outputs move: op=delta|goal_seek|grid. Example: "
        "inputs=['Inputs!B3=8%'], outputs=['Model!B20']. Never writes.",
        title="What-if analysis",
    ),
    # --- full -------------------------------------------------------------------------------
    _w(
        "create_workbook",
        "full",
        "Create a new empty branded .xlsx with the given sheet names. Example: "
        "create_workbook(path='out.xlsx', sheets=['Data','Summary'])",
        "Create workbook",
        idempotent=False,
    ),
    _w(
        "sheet_ops",
        "full",
        "Add, rename, delete, copy, move, hide or unhide a sheet (op=...). Renames also fix "
        "formulas that point to the sheet.",
        "Sheet operations",
        destructive=True,
        idempotent=False,
    ),
    _w(
        "range_ops",
        "full",
        "Edit a block: op=fill|copy|move|transpose|sort|replace|insert|delete|clear|split. "
        "Example: op='fill', range='D2:D200'. where=\"Status='Closed'\" filters rows.",
        "Range operations",
        destructive=True,
        idempotent=False,
    ),
    _w(
        "name_ops",
        "full",
        "List, add or delete defined names. Example: name_ops(path, op='add', name='TaxRate', "
        "ref='Rates!B2')",
        "Defined names",
    ),
    _w(
        "format_range",
        "full",
        "Style a range: role ('title','header','total'...), bold, colors, number_format "
        "('percent','m'...), align, border, merge, autofit.",
        "Format cells",
    ),
    _w(
        "set_table",
        "full",
        "Turn a range with a header row into a named Excel table; same name again resizes it. "
        "Example: set_table(path, sheet='Data', range='A1:D50', name='Sales', totals=true)",
        "Excel table",
        idempotent=False,
    ),
    _w(
        "rules",
        "full",
        "Add a dropdown/input rule (kind=list|whole|decimal|date|text_length|custom) or highlight "
        "rule (greater_than|between|duplicates|top_n|color_scale|data_bar|formula...).",
        "Validation and highlight rules",
        idempotent=False,
    ),
    _w(
        "add_chart",
        "full",
        "Insert a clean branded chart (type=auto picks one) from data_range: header row + "
        "category column + series columns. Optional highlight='North'; replace=true swaps "
        "the chart at anchor.",
        "Add chart",
        idempotent=False,
    ),
    _w(
        "pivot",
        "full",
        "Summarise a table: rows=['Region'], values=['Total'], agg=sum|count|avg|min|max, "
        "columns='Month'. to_sheet saves it; native=true makes a real Excel PivotTable.",
        "Pivot summary",
    ),
    _w(
        "recalc",
        "full",
        "Recalculate every formula with LibreOffice so cached values are fresh (needs LibreOffice"
        " installed).",
        "Recalculate",
    ),
    _w(
        "sheet_layout",
        "full",
        "Set how a sheet looks and prints: freeze, widths=['A=14'], heights, filter_range, hide, "
        "gridlines, zoom, print_area, landscape, fit_width, protect; workbook: calc_mode, "
        "protect_structure.",
        "Sheet layout",
    ),
    _w(
        "annotate",
        "full",
        "Add a note, hyperlink or image to a cell, or remove objects: op=note|link|image|remove. "
        "Example: op='link', cell='B2', url='https://...'",
        "Notes, links, images",
        idempotent=False,
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
