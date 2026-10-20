"""Transport-agnostic tool handlers: args in, rendered text out. Shared by MCP and CLI.

Each public method IS a tool: its name matches `toolspec.TOOLS` and its signature
becomes the MCP input schema, so keep arguments flat and primitive (PLAN §3.4).
"""

from __future__ import annotations

from slim_sheet.config import Settings
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
from slim_sheet.core.render_formula import render_cell, render_eval, render_validation
from slim_sheet.core.render_results import (
    render_find,
    render_formula,
    render_profile,
    render_query,
    render_write,
)
from slim_sheet.services import charts, recalc, reshape, structure, styling
from slim_sheet.services.anomalies import anomalies as find_anomalies
from slim_sheet.services.clean import clean as clean_table
from slim_sheet.services.convert import convert as convert_file
from slim_sheet.services.describe import describe_cell as describe_one
from slim_sheet.services.diff import diff as diff_files
from slim_sheet.services.evaluate import evaluate as evaluate_formula
from slim_sheet.services.find import find as find_cells
from slim_sheet.services.formula import explain_formula as explain
from slim_sheet.services.formula import validate_formula as validate_one
from slim_sheet.services.merge import merge as merge_tables
from slim_sheet.services.overview import overview as describe
from slim_sheet.services.profile import profile as profile_table
from slim_sheet.services.query import query as run_query
from slim_sheet.services.read import ReadRequest, parse_format
from slim_sheet.services.read import read as read_grid
from slim_sheet.services.undo import undo as undo_write
from slim_sheet.services.write import WriteRequest
from slim_sheet.services.write import write as write_cells

Scalar = str | int | float | bool | None


class Handlers:
    def __init__(self, settings: Settings) -> None:
        self.settings = settings

    def overview(self, path: str) -> str:
        return render_overview(describe(path, self.settings))

    def read(
        self,
        path: str,
        sheet: str | None = None,
        range: str | None = None,
        cursor: str | None = None,
        max_tokens: int | None = None,
        format: str = "markdown",
    ) -> str:
        request = ReadRequest(path, sheet, range, cursor, max_tokens, format)
        grid = read_grid(request, self.settings)
        return render_grid(grid, parse_format(format), self.settings.max_cell_chars)

    def find(
        self,
        path: str,
        text: str,
        sheet: str | None = None,
        regex: bool = False,
        max_results: int = 30,
    ) -> str:
        return render_find(find_cells(path, text, self.settings, sheet, regex, max_results))

    def query(self, path: str, sql: str, offset: int = 0, max_tokens: int | None = None) -> str:
        result = run_query(path, sql, self.settings, offset, max_tokens)
        return render_query(result, self.settings.max_cell_chars)

    def profile(self, path: str, sheet: str | None = None, columns: list[str] | None = None) -> str:
        return render_profile(profile_table(path, self.settings, sheet, columns))

    def write(
        self,
        path: str,
        sheet: str,
        range: str,
        values: list[list[Scalar]],
        out_path: str | None = None,
        overwrite: bool = False,
    ) -> str:
        request = WriteRequest(path, sheet, range, values, out_path, overwrite)
        return render_write(write_cells(request, self.settings))

    def explain_formula(
        self,
        formula: str | None = None,
        path: str | None = None,
        cell: str | None = None,
        sheet: str | None = None,
    ) -> str:
        report, address = explain(self.settings, formula, path, cell, sheet)
        return render_formula(report, address)

    # --- standard tier -------------------------------------------------------------

    def clean(
        self,
        path: str,
        ops: list[str],
        sheet: str | None = None,
        keys: list[str] | None = None,
        apply: bool = False,
        out_path: str | None = None,
        overwrite: bool = False,
    ) -> str:
        result = clean_table(path, ops, self.settings, sheet, keys, apply, out_path, overwrite)
        return render_clean(result)

    def anomalies(
        self,
        path: str,
        sheet: str | None = None,
        columns: list[str] | None = None,
        max_results: int = 20,
    ) -> str:
        return render_anomalies(find_anomalies(path, self.settings, sheet, columns, max_results))

    def merge(
        self,
        path: str,
        other_path: str,
        keys: list[str],
        how: str = "left",
        sheet: str | None = None,
        other_sheet: str | None = None,
        out_path: str | None = None,
        overwrite: bool = False,
        max_tokens: int | None = None,
    ) -> str:
        result = merge_tables(
            path,
            other_path,
            keys,
            self.settings,
            how,
            sheet,
            other_sheet,
            out_path,
            overwrite,
            max_tokens,
        )
        return render_merge(result, self.settings.max_cell_chars)

    def convert(
        self, path: str, out_path: str, sheet: str | None = None, overwrite: bool = False
    ) -> str:
        return render_output(convert_file(path, out_path, self.settings, sheet, overwrite))

    def diff(
        self,
        path: str,
        other_path: str,
        sheet: str | None = None,
        key: str | None = None,
        max_results: int = 50,
    ) -> str:
        return render_diff(diff_files(path, other_path, self.settings, sheet, key, max_results))

    def undo(self, path: str) -> str:
        outcome = undo_write(path, self.settings)
        return render_undo(outcome.target, outcome.note, outcome.restored_from, outcome.kept_as)

    def validate_formula(
        self,
        formula: str | None = None,
        path: str | None = None,
        cell: str | None = None,
        sheet: str | None = None,
        target: str | None = None,
    ) -> str:
        report, address = validate_one(self.settings, formula, path, cell, sheet, target)
        return render_validation(report, address)

    def evaluate(
        self,
        formula: str | None = None,
        path: str | None = None,
        cell: str | None = None,
        sheet: str | None = None,
    ) -> str:
        return render_eval(evaluate_formula(self.settings, formula, path, cell, sheet))

    def describe_cell(self, path: str, cell: str, sheet: str | None = None) -> str:
        return render_cell(describe_one(self.settings, path, cell, sheet))

    # --- full tier -----------------------------------------------------------------

    def create_workbook(
        self, path: str, sheets: list[str] | None = None, overwrite: bool = False
    ) -> str:
        return render_output(structure.create_workbook(path, self.settings, sheets, overwrite))

    def sheet_ops(
        self,
        path: str,
        op: str,
        sheet: str,
        new_name: str | None = None,
        position: int | None = None,
        out_path: str | None = None,
        overwrite: bool = False,
    ) -> str:
        result = structure.sheet_ops(
            path, op, sheet, self.settings, new_name, position, out_path, overwrite
        )
        return render_output(result)

    def rows_ops(
        self,
        path: str,
        sheet: str,
        op: str,
        at: str,
        count: int = 1,
        out_path: str | None = None,
        overwrite: bool = False,
    ) -> str:
        result = structure.rows_ops(path, sheet, op, at, self.settings, count, out_path, overwrite)
        return render_output(result)

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
        sheet: str,
        range: str,
        bold: bool | None = None,
        italic: bool | None = None,
        font_color: str | None = None,
        font_size: float | None = None,
        fill_color: str | None = None,
        number_format: str | None = None,
        align: str | None = None,
        wrap: bool | None = None,
        border: str | None = None,
        autofit: bool = False,
        out_path: str | None = None,
        overwrite: bool = False,
    ) -> str:
        result = styling.format_range(
            path, sheet, range, self.settings, bold, italic, font_color, font_size, fill_color,
            number_format, align, wrap, border, autofit, out_path, overwrite,
        )  # fmt: skip
        return render_output(result)

    def set_table(
        self,
        path: str,
        sheet: str,
        range: str,
        name: str,
        style: str = "TableStyleMedium9",
        out_path: str | None = None,
        overwrite: bool = False,
    ) -> str:
        result = styling.set_table(
            path, sheet, range, name, self.settings, style, out_path, overwrite
        )
        return render_output(result)

    def add_validation(
        self,
        path: str,
        sheet: str,
        range: str,
        type: str,
        values: list[str] | None = None,
        min: str | None = None,
        max: str | None = None,
        formula: str | None = None,
        message: str | None = None,
        allow_blank: bool = True,
        out_path: str | None = None,
        overwrite: bool = False,
    ) -> str:
        result = styling.add_validation(
            path, sheet, range, type, self.settings, values, min, max, formula, message,
            allow_blank, out_path, overwrite,
        )  # fmt: skip
        return render_output(result)

    def conditional_format(
        self,
        path: str,
        sheet: str,
        range: str,
        rule: str,
        value: str | None = None,
        value2: str | None = None,
        fill_color: str = "light_red",
        font_color: str = "9C0006",
        out_path: str | None = None,
        overwrite: bool = False,
    ) -> str:
        result = styling.conditional_format(
            path, sheet, range, rule, self.settings, value, value2, fill_color, font_color,
            out_path, overwrite,
        )  # fmt: skip
        return render_output(result)

    def add_chart(
        self,
        path: str,
        sheet: str,
        type: str,
        data_range: str,
        anchor: str | None = None,
        title: str | None = None,
        x_title: str | None = None,
        y_title: str | None = None,
        out_path: str | None = None,
        overwrite: bool = False,
    ) -> str:
        result = charts.add_chart(
            path, sheet, type, data_range, self.settings, anchor, title, x_title, y_title,
            out_path, overwrite,
        )  # fmt: skip
        return render_output(result)

    def pivot(
        self,
        path: str,
        rows: list[str],
        values: list[str],
        agg: str = "sum",
        columns: str | None = None,
        sheet: str | None = None,
        to_sheet: str | None = None,
        out_path: str | None = None,
        overwrite: bool = False,
        max_tokens: int | None = None,
    ) -> str:
        result = reshape.pivot(
            path, rows, values, self.settings, agg, columns, sheet, to_sheet, out_path,
            overwrite, max_tokens,
        )  # fmt: skip
        return render_pivot(result, self.settings.max_cell_chars)

    def split(
        self,
        path: str,
        by: str,
        sheet: str | None = None,
        mode: str = "sheets",
        out_path: str | None = None,
        overwrite: bool = False,
    ) -> str:
        return render_output(
            reshape.split(path, by, self.settings, sheet, mode, out_path, overwrite)
        )

    def recalc(self, path: str, out_path: str | None = None, overwrite: bool = False) -> str:
        return render_output(recalc.recalc(path, self.settings, out_path, overwrite))
