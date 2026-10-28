"""Transport-agnostic tool handlers: args in, rendered text out. Shared by MCP and CLI.

Each public method IS a tool: its name matches `toolspec.TOOLS` and its signature
becomes the MCP input schema, so keep arguments flat and primitive (PLAN §3.4).
"""

from __future__ import annotations

from pathlib import Path
from typing import Annotated

from pydantic import BeforeValidator

from slim_sheet.config import Settings
from slim_sheet.core.errors import ErrorCode, SlimError
from slim_sheet.core.render import render_grid, render_overview
from slim_sheet.core.render_data import (
    render_anomalies,
    render_clean,
    render_diff,
    render_merge,
    render_names,
    render_output,
    render_pivot,
    render_undo,
)
from slim_sheet.core.render_formula import (
    render_cell,
    render_eval,
    render_folder,
    render_macro_source,
    render_macros,
    render_trace,
    render_validation,
    render_what_if,
)
from slim_sheet.core.render_results import (
    render_find,
    render_formula,
    render_profile,
    render_query,
    render_write,
)
from slim_sheet.core.styles import choice
from slim_sheet.interfaces.structured import publish
from slim_sheet.services import (
    audit,
    charts,
    design,
    layout,
    macros,
    ranges,
    recalc,
    report,
    reshape,
    structure,
    styling,
)
from slim_sheet.services.annotate import AnnotateRequest
from slim_sheet.services.annotate import annotate as annotate_cells
from slim_sheet.services.anomalies import anomalies as find_anomalies
from slim_sheet.services.clean import clean as clean_table
from slim_sheet.services.convert import CsvOptions
from slim_sheet.services.convert import convert as convert_file
from slim_sheet.services.describe import describe_cell as describe_one
from slim_sheet.services.diff import diff as diff_files
from slim_sheet.services.evaluate import evaluate as evaluate_formula
from slim_sheet.services.find import find as find_cells
from slim_sheet.services.formula import explain_formula as explain
from slim_sheet.services.formula import validate_formula as validate_one
from slim_sheet.services.layout import LayoutRequest
from slim_sheet.services.merge import merge as merge_tables
from slim_sheet.services.overview import list_workbooks
from slim_sheet.services.overview import overview as describe
from slim_sheet.services.profile import profile as profile_table
from slim_sheet.services.query import query as run_query
from slim_sheet.services.ranges import RangeRequest
from slim_sheet.services.read import ReadRequest, parse_format
from slim_sheet.services.read import read as read_grid
from slim_sheet.services.undo import undo as undo_write
from slim_sheet.services.what_if import WhatIfRequest
from slim_sheet.services.what_if import what_if as run_what_if
from slim_sheet.services.write import WriteRequest
from slim_sheet.services.write import write as write_cells

Scalar = str | int | float | bool | None
ANOMALY_SCOPES = ("data", "formulas", "macros", "all")


def _as_text(value: object) -> object:
    """MCP clients often send "0" or "100" as JSON numbers; accept them where text is expected."""
    if isinstance(value, bool):
        return value  # left for validation to reject: True is not a sheet name or a format
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    if isinstance(value, int | float):
        return str(value)
    return value


# Schema stays {"type": "string"}; numbers are coerced to text before validation.
Text = Annotated[str, BeforeValidator(_as_text)]


class Handlers:
    def __init__(self, settings: Settings) -> None:
        self.settings = settings

    def overview(self, path: str = "") -> str:
        folder = list_workbooks(path, self.settings)
        if folder is not None:
            publish(folder)
            return render_folder(folder)
        info = describe(path, self.settings)
        publish(info)
        return render_overview(info)

    def read(
        self,
        path: str,
        sheet: Text | None = None,
        range: str | None = None,
        cursor: str | None = None,
        max_tokens: int | None = None,
        format: str = "markdown",
    ) -> str:
        if macros.is_vba_sheet(sheet):  # read-only VBA: module list, or one module's source
            if not range:
                listing = macros.list_macros(path, self.settings)
                publish(listing)
                return render_macros(listing)
            code = macros.macro_source(path, self.settings, range, cursor, max_tokens)
            publish(code)
            return render_macro_source(code)
        request = ReadRequest(path, sheet, range, cursor, max_tokens, format)
        grid = read_grid(request, self.settings)
        publish(grid)
        return render_grid(grid, parse_format(format), self.settings.max_cell_chars)

    def find(
        self,
        path: str,
        text: Text,
        sheet: Text | None = None,
        regex: bool = False,
        max_results: int = 30,
    ) -> str:
        found = find_cells(path, text, self.settings, sheet, regex, max_results)
        publish(found)
        return render_find(found)

    def query(self, path: str, sql: str, offset: int = 0, max_tokens: int | None = None) -> str:
        result = run_query(path, sql, self.settings, offset, max_tokens)
        publish(result)
        return render_query(result, self.settings.max_cell_chars)

    def profile(
        self, path: str, sheet: Text | None = None, columns: list[Text] | None = None
    ) -> str:
        profiled = profile_table(path, self.settings, sheet, columns)
        publish(profiled)
        return render_profile(profiled)

    def write(
        self,
        path: str,
        sheet: Text,
        range: str,
        values: list[list[Scalar]],
        out_path: str | None = None,
        overwrite: bool = False,
    ) -> str:
        request = WriteRequest(path, sheet, range, values, out_path, overwrite)
        return render_write(write_cells(request, self.settings))

    def explain_formula(
        self,
        formula: Text | None = None,
        path: str | None = None,
        cell: str | None = None,
        sheet: Text | None = None,
    ) -> str:
        report, address = explain(self.settings, formula, path, cell, sheet)
        return render_formula(report, address)

    # --- standard tier -------------------------------------------------------------

    def clean(
        self,
        path: str,
        ops: list[str],
        sheet: Text | None = None,
        keys: list[Text] | None = None,
        apply: bool = False,
        out_path: str | None = None,
        overwrite: bool = False,
    ) -> str:
        result = clean_table(path, ops, self.settings, sheet, keys, apply, out_path, overwrite)
        return render_clean(result)

    def anomalies(
        self,
        path: str,
        sheet: Text | None = None,
        columns: list[Text] | None = None,
        max_results: int = 20,
        scope: str = "data",
    ) -> str:
        key = choice(scope, ANOMALY_SCOPES, "scope")
        parts = []
        if key in ("data", "all"):
            found = find_anomalies(path, self.settings, sheet, columns, max_results)
            parts.append(render_anomalies(found))
        if key in ("formulas", "all"):
            audited = audit.formula_audit(path, self.settings, sheet, max_results)
            parts.append(render_anomalies(audited))
        if key == "macros" or (key == "all" and macros.has_project(path, self.settings)):
            parts.append(render_anomalies(macros.macro_audit(path, self.settings, max_results)))
        return "\n\n".join(parts)

    def merge(
        self,
        path: str,
        other_path: str,
        keys: list[Text],
        how: str = "left",
        sheet: Text | None = None,
        other_sheet: Text | None = None,
        out_path: str | None = None,
        overwrite: bool = False,
        max_tokens: int | None = None,
        compare: list[Text] | None = None,
        tolerance: float = 0.0,
    ) -> str:
        result = merge_tables(
            path, other_path, keys, self.settings, how, sheet, other_sheet, out_path, overwrite,
            max_tokens, compare, tolerance,
        )  # fmt: skip
        return render_merge(result, self.settings.max_cell_chars)

    def convert(
        self,
        path: str,
        out_path: str | None = None,
        sheet: Text | None = None,
        overwrite: bool = False,
        split_by: Text | None = None,
        mode: str = "sheets",
        delimiter: Text | None = None,
        encoding: str | None = None,
        decimal: Text | None = None,
    ) -> str:
        if split_by:
            result = reshape.split(path, split_by, self.settings, sheet, mode, out_path, overwrite)
            return render_output(result)
        if not out_path:
            raise SlimError(
                ErrorCode.BAD_ARGUMENT,
                "convert needs out_path (or split_by to split a table)",
                hint="e.g. convert(path='a.xls', out_path='a.xlsx') or split_by='Region'",
            )
        if not Path(out_path).suffix:  # a folder: export the VBA modules as .bas/.cls files
            return render_output(macros.export_macros(path, self.settings, out_path, overwrite))
        options = CsvOptions(delimiter, encoding, decimal)
        return render_output(convert_file(path, out_path, self.settings, sheet, overwrite, options))

    def diff(
        self,
        path: str,
        other_path: str,
        sheet: Text | None = None,
        key: Text | None = None,
        max_results: int = 50,
    ) -> str:
        compared = diff_files(path, other_path, self.settings, sheet, key, max_results)
        publish(compared)
        return render_diff(compared)

    def undo(self, path: str) -> str:
        outcome = undo_write(path, self.settings)
        return render_undo(outcome.target, outcome.note, outcome.restored_from, outcome.kept_as)

    def evaluate(
        self,
        formula: Text | None = None,
        path: str | None = None,
        cell: str | None = None,
        sheet: Text | None = None,
        check_only: bool = False,
        target: Text | None = None,
    ) -> str:
        if check_only:
            report, address = validate_one(self.settings, formula, path, cell, sheet, target)
            return render_validation(report, address)
        return render_eval(evaluate_formula(self.settings, formula, path, cell, sheet))

    def describe_cell(
        self,
        path: str,
        cell: str,
        sheet: Text | None = None,
        trace: str = "none",
        depth: int = 3,
    ) -> str:
        text = render_cell(describe_one(self.settings, path, cell, sheet))
        if trace == "none":
            return text
        traced = audit.trace_cell(path, self.settings, cell, sheet, trace, depth)
        return text + "\n" + render_trace(traced)

    def what_if(
        self,
        path: str,
        outputs: list[Text],
        inputs: list[Text],
        op: str = "delta",
        target: Text | None = None,
        sheet: Text | None = None,
    ) -> str:
        request = WhatIfRequest(path, outputs, inputs, op, target, sheet)
        return render_what_if(run_what_if(request, self.settings))

    # --- full tier -----------------------------------------------------------------

    def create_workbook(
        self,
        path: str,
        sheets: list[str] | None = None,
        overwrite: bool = False,
        brand: str | None = None,
    ) -> str:
        return render_output(
            structure.create_workbook(path, self.settings, sheets, overwrite, brand)
        )

    def sheet_ops(
        self,
        path: str,
        op: str,
        sheet: Text,
        new_name: Text | None = None,
        position: int | None = None,
        out_path: str | None = None,
        overwrite: bool = False,
    ) -> str:
        result = structure.sheet_ops(
            path, op, sheet, self.settings, new_name, position, out_path, overwrite
        )
        return render_output(result)

    def range_ops(
        self,
        path: str,
        sheet: Text,
        op: str,
        range: Text | None = None,
        to: str | None = None,
        by: Text | None = None,
        desc: bool = False,
        find: Text | None = None,
        replace: Text | None = None,
        where: str | None = None,
        header: bool = True,
        values_only: bool = False,
        formats_only: bool = False,
        what: str = "contents",
        delimiter: Text = ",",
        out_path: str | None = None,
        overwrite: bool = False,
    ) -> str:
        request = RangeRequest(
            path, sheet, op, range, to, by, desc, find, replace, where, header, out_path,
            overwrite, values_only, formats_only, what, delimiter,
        )  # fmt: skip
        return render_output(ranges.range_ops(request, self.settings))

    def name_ops(
        self,
        path: str,
        op: str,
        name: str | None = None,
        ref: str | None = None,
        out_path: str | None = None,
        overwrite: bool = False,
    ) -> str:
        if op == "list":
            return render_names(path, structure.list_names(path, self.settings))
        result = structure.name_ops(path, op, self.settings, name, ref, out_path, overwrite)
        return render_output(result)

    def format_range(
        self,
        path: str,
        sheet: Text,
        range: str,
        bold: bool | None = None,
        italic: bool | None = None,
        font_color: Text | None = None,
        font_size: float | None = None,
        fill_color: Text | None = None,
        number_format: Text | None = None,
        align: str | None = None,
        valign: str | None = None,
        wrap: bool | None = None,
        border: str | None = None,
        autofit: bool = False,
        role: str | None = None,
        merge: bool | None = None,
        out_path: str | None = None,
        overwrite: bool = False,
    ) -> str:
        result = styling.format_range(
            path, sheet, range, self.settings, bold, italic, font_color, font_size, fill_color,
            number_format, align, wrap, border, autofit, out_path, overwrite, valign, role, merge,
        )  # fmt: skip
        return render_output(result)

    def set_table(
        self,
        path: str,
        sheet: Text,
        range: str,
        name: str,
        style: str = "TableStyleMedium9",
        totals: bool = False,
        out_path: str | None = None,
        overwrite: bool = False,
    ) -> str:
        result = styling.set_table(
            path, sheet, range, name, self.settings, style, out_path, overwrite, totals
        )
        return render_output(result)

    def rules(
        self,
        path: str,
        sheet: Text,
        range: str,
        kind: str,
        values: list[Text] | None = None,
        value: Text | None = None,
        value2: Text | None = None,
        formula: Text | None = None,
        message: Text | None = None,
        allow_blank: bool = True,
        fill_color: Text = "light_red",
        font_color: Text = "9C0006",
        out_path: str | None = None,
        overwrite: bool = False,
    ) -> str:
        result = styling.rules(
            path, sheet, range, kind, self.settings, values, value, value2, formula, message,
            allow_blank, fill_color, font_color, out_path, overwrite,
        )  # fmt: skip
        return render_output(result)

    def add_chart(
        self,
        path: str,
        sheet: Text,
        type: str,
        data_range: str,
        anchor: str | None = None,
        title: Text | None = None,
        x_title: Text | None = None,
        y_title: Text | None = None,
        highlight: Text | None = None,
        brand: str | None = None,
        replace: bool = False,
        out_path: str | None = None,
        overwrite: bool = False,
    ) -> str:
        result = charts.add_chart(
            path, sheet, type, data_range, self.settings, anchor, title, x_title, y_title,
            out_path, overwrite, highlight, brand, replace,
        )  # fmt: skip
        return render_output(result)

    def style_sheet(
        self,
        path: str,
        sheet: Text,
        range: str | None = None,
        preset: str = "report",
        number_format: Text = "auto",
        total_row: bool | None = None,
        brand: str | None = None,
        out_path: str | None = None,
        overwrite: bool = False,
    ) -> str:
        result = design.style_sheet(
            path, sheet, self.settings, range, preset, number_format, total_row, brand,
            out_path, overwrite,
        )  # fmt: skip
        return render_output(result)

    def build_report(
        self,
        path: str,
        data_sheet: Text,
        title: Text,
        kpis: list[Text] | None = None,
        charts: list[Text] | None = None,
        subtitle: Text | None = None,
        sheet: Text = "Dashboard",
        brand: str | None = None,
        out_path: str | None = None,
        overwrite: bool = False,
        messages: list[Text] | None = None,
    ) -> str:
        result = report.build_report(
            path, data_sheet, title, self.settings, kpis, charts, subtitle, sheet, brand,
            out_path, overwrite, messages,
        )  # fmt: skip
        return render_output(result)

    def pivot(
        self,
        path: str,
        rows: list[Text],
        values: list[Text],
        agg: str = "sum",
        columns: Text | None = None,
        sheet: Text | None = None,
        to_sheet: Text | None = None,
        out_path: str | None = None,
        overwrite: bool = False,
        max_tokens: int | None = None,
        native: bool = False,
    ) -> str:
        result = reshape.pivot(
            path, rows, values, self.settings, agg, columns, sheet, to_sheet, out_path,
            overwrite, max_tokens, native,
        )  # fmt: skip
        return render_pivot(result, self.settings.max_cell_chars)

    def sheet_layout(
        self,
        path: str,
        sheet: Text,
        freeze: str | None = None,
        widths: list[str] | None = None,
        heights: list[str] | None = None,
        filter_range: str | None = None,
        hide: list[str] | None = None,
        unhide: list[str] | None = None,
        gridlines: bool | None = None,
        zoom: int | None = None,
        print_area: str | None = None,
        landscape: bool | None = None,
        fit_width: bool | None = None,
        protect: bool | None = None,
        calc_mode: str | None = None,
        protect_structure: bool | None = None,
        out_path: str | None = None,
        overwrite: bool = False,
    ) -> str:
        request = LayoutRequest(
            path, sheet, freeze, widths, heights, filter_range, hide, unhide, gridlines, zoom,
            print_area, landscape, fit_width, protect, calc_mode, protect_structure, out_path,
            overwrite,
        )  # fmt: skip
        return render_output(layout.sheet_layout(request, self.settings))

    def annotate(
        self,
        path: str,
        sheet: Text,
        op: str,
        cell: str,
        text: Text | None = None,
        url: str | None = None,
        image_path: str | None = None,
        target: str = "all",
        out_path: str | None = None,
        overwrite: bool = False,
    ) -> str:
        request = AnnotateRequest(
            path, sheet, op, cell, text, url, image_path, target, out_path, overwrite
        )
        return render_output(annotate_cells(request, self.settings))

    def recalc(self, path: str, out_path: str | None = None, overwrite: bool = False) -> str:
        return render_output(recalc.recalc(path, self.settings, out_path, overwrite))
