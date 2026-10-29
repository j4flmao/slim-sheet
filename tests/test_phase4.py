from __future__ import annotations

import asyncio
import datetime as dt
import shutil
import zipfile
from pathlib import Path

import openpyxl
import pytest
from mcp.server.mcpserver.exceptions import ToolError
from openpyxl.formatting.rule import CellIsRule
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.table import Table

from slim_sheet.adapters import libreoffice
from slim_sheet.config import Settings
from slim_sheet.core.errors import ErrorCode, SlimError
from slim_sheet.core.shift import REF_ERROR, Shift, quote_sheet, shift_address, shift_formula
from slim_sheet.interfaces.cli import main
from slim_sheet.interfaces.handlers import Handlers
from slim_sheet.interfaces.mcp_server import build_server
from slim_sheet.services import charts, recalc, reshape, structure, styling


@pytest.fixture
def refs_xlsx(tmp_path: Path) -> Path:
    """Data!A1:B8 with a table, merge, validation, conditional format, name and formulas."""
    wb = openpyxl.Workbook()
    ws = wb.active
    assert ws is not None
    ws.title = "Data"
    ws.append(["k", "v"])
    for i in range(1, 8):
        ws.append([f"r{i}", i])
    ws["D1"], ws["D2"] = "=SUM(B2:B8)", "=B5*2"
    ws.merge_cells("F5:G6")
    ws.add_table(Table(displayName="T1", ref="A1:B8"))
    dv = DataValidation(type="whole")
    ws.add_data_validation(dv)
    dv.add("B2:B8")
    ws.conditional_formatting.add("B6:B8", CellIsRule(operator="greaterThan", formula=["3"]))
    other = wb.create_sheet("Other")
    other["A1"] = "=Data!B7+Data!B3"
    wb.defined_names["Last"] = DefinedName("Last", attr_text="Data!$B$8")
    path = tmp_path / "refs.xlsx"
    wb.save(path)
    return path


@pytest.fixture
def sales_grid(tmp_path: Path) -> Path:
    wb = openpyxl.Workbook()
    ws = wb.active
    assert ws is not None
    ws.title = "Sales"
    ws.append(["Region", "Month", "Total"])
    rows = [
        ("North", "Jan", 10),
        ("North", "Feb", 20),
        ("South", "Jan", 30),
        ("South", "Feb", 40),
        ("East", "Jan", 5),
    ]
    for row in rows:
        ws.append(list(row))
    path = tmp_path / "grid.xlsx"
    wb.save(path)
    return path


def _slim(path: Path) -> openpyxl.Workbook:
    return openpyxl.load_workbook(path.with_name(f"{path.stem}.slim.xlsx"))


# --- core shift ------------------------------------------------------------------------


@pytest.mark.parametrize(
    ("address", "shift", "expected"),
    [
        ("B3", Shift("rows", 3, 2), "B5"),
        ("B2", Shift("rows", 3, 2), "B2"),
        ("A1:C9", Shift("rows", 3, -2), "A1:C7"),
        ("A3:B4", Shift("rows", 3, -2), REF_ERROR),
        ("A4:A10", Shift("rows", 3, -2), "A3:A8"),
        ("$A$5", Shift("rows", 3, 2), "$A$7"),
        ("A:C", Shift("rows", 1, 5), "A:C"),
        ("B2:D2", Shift("cols", 3, 1), "B2:E2"),
        ("C1", Shift("cols", 3, -1), REF_ERROR),
        ("TaxRate", Shift("rows", 1, 1), "TaxRate"),
    ],
)
def test_shift_address(address: str, shift: Shift, expected: str) -> None:
    # Act / Assert
    assert shift_address(address, shift) == expected


def test_shift_formula_only_touches_target_sheet() -> None:
    # Act
    new, changed = shift_formula(
        "=SUM(A1:A9)+Other!A5+'Data'!B7", "Data", "Data", Shift("rows", 3, 2)
    )

    # Assert
    assert changed
    assert new == "=SUM(A1:A11)+Other!A5+'Data'!B9"


def test_quote_sheet() -> None:
    # Act / Assert
    assert quote_sheet("Data") == "Data"
    assert quote_sheet("Q1 Data") == "'Q1 Data'"
    assert quote_sheet("A1") == "'A1'"
    assert quote_sheet("It's") == "'It''s'"


# --- rows_ops --------------------------------------------------------------------------


def test_insert_rows_repairs_every_reference(refs_xlsx: Path, settings: Settings) -> None:
    # Act
    result = structure.rows_ops("refs.xlsx", "Data", "insert_rows", "4", settings, count=2)

    # Assert
    wb = _slim(refs_xlsx)
    ws = wb["Data"]
    assert ws["D1"].value == "=SUM(B2:B10)"
    assert ws["D2"].value == "=B7*2"
    assert wb["Other"]["A1"].value == "=Data!B9+Data!B3"
    assert [str(r) for r in ws.merged_cells.ranges] == ["F7:G8"]
    assert ws.tables["T1"].ref == "A1:B10"
    assert str(ws.data_validations.dataValidation[0].sqref) == "B2:B10"
    assert [str(c.sqref) for c in ws.conditional_formatting] == ["B8:B10"]
    assert wb.defined_names["Last"].attr_text == "Data!$B$10"
    assert ws["A6"].value == "r3"
    assert "updated 3 formula(s)" in result.summary


def test_delete_rows_shrinks_ranges_and_reports_ref_errors(
    refs_xlsx: Path, settings: Settings
) -> None:
    # Act
    result = structure.rows_ops("refs.xlsx", "Data", "delete_rows", "5", settings, count=3)

    # Assert
    ws = _slim(refs_xlsx)["Data"]
    assert ws["D1"].value == "=SUM(B2:B5)"
    assert ws["D2"].value == "=#REF!*2"
    assert ws.tables["T1"].ref == "A1:B5"
    assert any("Data!D2 now contains #REF!" in w for w in result.warnings)


def test_delete_cols_and_bad_arguments(refs_xlsx: Path, settings: Settings) -> None:
    # Act
    structure.rows_ops("refs.xlsx", "Other", "insert_cols", "A", settings)

    # Assert
    assert _slim(refs_xlsx)["Other"]["B1"].value == "=Data!B7+Data!B3"
    with pytest.raises(SlimError):
        structure.rows_ops("refs.xlsx", "Data", "insert_rows", "C", settings)
    with pytest.raises(SlimError):
        structure.rows_ops("refs.xlsx", "Data", "shuffle", "3", settings)
    with pytest.raises(SlimError) as info:
        structure.rows_ops("refs.xlsx", "Data", "delete_rows", "1", settings, count=9)
    assert "table 'T1'" in info.value.message


# --- sheet_ops / create_workbook / name_ops --------------------------------------------


def test_sheet_ops_lifecycle(refs_xlsx: Path, settings: Settings) -> None:
    # Act
    structure.sheet_ops("refs.xlsx", "rename", "Data", settings, new_name="Q1 Data")
    structure.sheet_ops("refs.slim.xlsx", "add", "Notes", settings, position=1)
    structure.sheet_ops("refs.slim.xlsx", "copy", "Other", settings, new_name="Other2")
    structure.sheet_ops("refs.slim.xlsx", "move", "Other2", settings, position=1)
    structure.sheet_ops("refs.slim.xlsx", "hide", "Notes", settings)
    deleted = structure.sheet_ops("refs.slim.xlsx", "delete", "Q1 Data", settings)

    # Assert
    wb = _slim(refs_xlsx)
    assert wb.sheetnames == ["Other2", "Notes", "Other"]
    assert wb["Notes"].sheet_state == "hidden"
    assert wb["Other"]["A1"].value == "='Q1 Data'!B7+'Q1 Data'!B3"
    assert "will show #REF!" in deleted.warnings[0]


@pytest.mark.parametrize(
    ("op", "sheet", "new_name", "fragment"),
    [
        ("rename", "Data", "Other", "already exists"),
        ("rename", "Data", "a/b", "Invalid sheet name"),
        ("rename", "Nope", "X", "No sheet"),
        ("explode", "Data", None, "Unknown op"),
    ],
)
def test_sheet_ops_errors(
    refs_xlsx: Path, settings: Settings, op: str, sheet: str, new_name: str | None, fragment: str
) -> None:
    # Act / Assert
    with pytest.raises(SlimError) as info:
        structure.sheet_ops("refs.xlsx", op, sheet, settings, new_name=new_name)
    assert fragment in info.value.message


def test_cannot_hide_or_delete_last_visible_sheet(tmp_path: Path, settings: Settings) -> None:
    # Arrange
    structure.create_workbook("one.xlsx", settings, ["Only"])

    # Act / Assert
    with pytest.raises(SlimError):
        structure.sheet_ops("one.xlsx", "hide", "Only", settings)
    with pytest.raises(SlimError):
        structure.sheet_ops("one.xlsx", "delete", "Only", settings)


def test_create_workbook(tmp_path: Path, settings: Settings) -> None:
    # Act
    result = structure.create_workbook("new.xlsx", settings, ["Data", "Summary"])

    # Assert
    assert openpyxl.load_workbook(tmp_path / "new.xlsx").sheetnames == ["Data", "Summary"]
    assert "Data, Summary" in result.summary
    with pytest.raises(SlimError) as info:
        structure.create_workbook("new.xlsx", settings)
    assert info.value.code == ErrorCode.WRITE_CONFLICT
    with pytest.raises(SlimError):
        structure.create_workbook("bad.xlsx", settings, ["A", "a"])


def test_name_ops(refs_xlsx: Path, settings: Settings) -> None:
    # Act
    structure.name_ops("refs.xlsx", "add", settings, name="Rate", ref="Other!b2:c3")
    listed = structure.list_names("refs.slim.xlsx", settings)
    structure.name_ops("refs.slim.xlsx", "delete", settings, name="last")

    # Assert
    assert ("Rate", "Other!$B$2:$C$3") in listed
    assert list(_slim(refs_xlsx).defined_names) == ["Rate"]
    for bad_name in ("A1", "two words", ""):
        with pytest.raises(SlimError):
            structure.name_ops("refs.xlsx", "add", settings, name=bad_name, ref="Data!A1")
    with pytest.raises(SlimError):
        structure.name_ops("refs.xlsx", "add", settings, name="X", ref="A1")


# --- styling ---------------------------------------------------------------------------


def test_format_range_applies_styles(sales_grid: Path, settings: Settings) -> None:
    # Act
    result = styling.format_range(
        "grid.xlsx",
        "Sales",
        "A1:C1",
        settings,
        bold=True,
        fill_color="yellow",
        font_color="#C00000",
        number_format="currency",
        align="center",
        border="outline",
        autofit=True,
    )

    # Assert
    ws = _slim(sales_grid)["Sales"]
    cell = ws["B1"]
    assert cell.font.b is True
    assert cell.font.color.rgb == "FFC00000"
    assert cell.fill.fgColor.rgb == "FFFFFF00"
    assert cell.number_format == '"$"#,##0.00'
    assert cell.alignment.horizontal == "center"
    assert ws["A1"].border.left.style == "thin" and ws["B1"].border.left.style is None
    assert ws.column_dimensions["A"].width >= 8
    assert "font, fill, number format, alignment, border" in result.summary


def test_format_range_rejects_bad_vocab(sales_grid: Path, settings: Settings) -> None:
    # Act / Assert
    with pytest.raises(SlimError) as info:
        styling.format_range("grid.xlsx", "Sales", "A1", settings, fill_color="sparkly")
    assert "#RRGGBB" in info.value.hint
    with pytest.raises(SlimError):
        styling.format_range("grid.xlsx", "Sales", "A1", settings, align="middle")


def test_format_range_valign_and_vertical_align_hint(sales_grid: Path, settings: Settings) -> None:
    # Act
    styling.format_range("grid.xlsx", "Sales", "A1:C1", settings, align="left", valign="middle")
    with pytest.raises(SlimError) as info:
        styling.format_range("grid.xlsx", "Sales", "A1", settings, align="top")

    # Assert
    cell = _slim(sales_grid)["Sales"]["A1"]
    assert (cell.alignment.horizontal, cell.alignment.vertical) == ("left", "center")
    assert "valign='top'" in info.value.hint


def test_autofit_sizes_formulas_from_number_format_not_formula_text(
    sales_grid: Path, settings: Settings
) -> None:
    # Arrange
    wb = openpyxl.load_workbook(sales_grid)
    ws = wb["Sales"]
    ws["E2"] = "=SUM(Sales!C2:C3)/SUM(Sales!C2:C6)"
    ws["E2"].number_format = "0.00%"
    ws["F2"] = "=SUM(Sales!C2:C3)+SUM(Sales!C4:C6)+SUM(Sales!C2:C6)"
    wb.save(sales_grid)

    # Act
    styling.format_range("grid.xlsx", "Sales", "E2:F2", settings, autofit=True)

    # Assert
    widths = _slim(sales_grid)["Sales"].column_dimensions
    assert widths["E"].width == styling.PERCENT_CHARS + 2
    assert widths["F"].width == styling.GENERAL_CHARS + 2


def test_autofit_caps_wrapped_text_and_scales_with_font_size(
    sales_grid: Path, settings: Settings
) -> None:
    # Arrange
    wb = openpyxl.load_workbook(sales_grid)
    ws = wb["Sales"]
    ws["A1"] = "Share of revenue from the top ten companies"
    ws["B1"] = "Antidisestablishmentarianism rules"
    wb.save(sales_grid)
    styling.format_range("grid.xlsx", "Sales", "A1:B1", settings, wrap=True)
    styling.format_range("grid.slim.xlsx", "Sales", "C1", settings, font_size=22)

    # Act
    styling.format_range("grid.slim.xlsx", "Sales", "A1:C6", settings, autofit=True)

    # Assert
    widths = _slim(sales_grid)["Sales"].column_dimensions
    assert widths["A"].width == styling.WRAP_WIDTH + 2
    assert widths["B"].width == len("Antidisestablishmentarianism") + 2
    assert widths["C"].width == len("Total") * 2 + 2


def test_autofit_empty_range_and_dates(sales_grid: Path, settings: Settings) -> None:
    # Arrange
    wb = openpyxl.load_workbook(sales_grid)
    wb["Sales"]["H2"] = dt.date(2026, 10, 20)
    wb["Sales"]["H2"].number_format = "yyyy-mm-dd"
    wb.save(sales_grid)

    # Act
    result = styling.format_range("grid.xlsx", "Sales", "H1:J9", settings, autofit=True)

    # Assert
    assert "widths for 1 column(s)" in result.summary
    assert _slim(sales_grid)["Sales"].column_dimensions["H"].width == len("yyyy-mm-dd") + 2


def test_mcp_accepts_numbers_for_text_params(sales_grid: Path, settings: Settings) -> None:
    # Arrange: MCP clients turn "0" / "5" into JSON numbers
    server = build_server(settings)
    calls = [
        ("rules", {"range": "C2:C6", "kind": "between", "value": 0, "value2": 2.5}),
        ("format_range", {"range": "C2:C6", "number_format": 0}),
        ("rules", {"range": "C2:C6", "kind": "whole", "value": 1, "value2": 100}),
        ("range_ops", {"op": "insert", "range": 8}),
    ]

    # Act
    for tool, args in calls:
        asyncio.run(server.call_tool(tool, {"path": "grid.xlsx", "sheet": "Sales", **args}))

    # Assert
    ws = _slim(sales_grid)["Sales"]
    rule = next(r for cf in ws.conditional_formatting for r in cf.rules)
    assert rule.formula == ["0", "2.5"]
    assert ws["C2"].number_format == "0"
    dv = ws.data_validations.dataValidation[0]
    assert (dv.formula1, dv.formula2) == ("1", "100")


def test_mcp_text_params_still_reject_booleans(sales_grid: Path, settings: Settings) -> None:
    # Arrange
    server = build_server(settings)
    args = {"path": "grid.xlsx", "sheet": "Sales", "range": "A1", "number_format": True}

    # Act / Assert
    with pytest.raises(ToolError):
        asyncio.run(server.call_tool("format_range", args))


def test_set_table_and_conflicts(sales_grid: Path, settings: Settings) -> None:
    # Act
    styling.set_table("grid.xlsx", "Sales", "A1:C6", "SalesT", settings)

    # Assert
    assert _slim(sales_grid)["Sales"].tables["SalesT"].ref == "A1:C6"
    with pytest.raises(SlimError) as overlap:
        styling.set_table("grid.slim.xlsx", "Sales", "B2:C4", "Other", settings)
    assert "overlaps" in overlap.value.message
    with pytest.raises(SlimError):
        styling.set_table("grid.xlsx", "Sales", "A2:C6", "Bad Name", settings)
    with pytest.raises(SlimError) as header:
        styling.set_table("grid.xlsx", "Sales", "E1:F3", "Empty", settings)
    assert "unique, non-empty labels" in header.value.message


@pytest.mark.parametrize(
    ("kwargs", "check"),
    [
        ({"type": "list", "values": ["yes", "no"]}, ("list", '"yes,no"', None)),
        ({"type": "whole", "min": "1", "max": "10"}, ("whole", "1", "10")),
        ({"type": "date", "min": "2026-01-01"}, ("date", "DATE(2026,1,1)", None)),
        ({"type": "custom", "formula": "=LEN(A2)<=10"}, ("custom", "LEN(A2)<=10", None)),
    ],
)
def test_add_validation(
    sales_grid: Path,
    settings: Settings,
    kwargs: dict[str, object],
    check: tuple[str, str, str | None],
) -> None:
    # Act
    styling.add_validation("grid.xlsx", "Sales", "C2:C6", settings=settings, **kwargs)  # type: ignore[arg-type]

    # Assert
    dv = _slim(sales_grid)["Sales"].data_validations.dataValidation[0]
    assert (dv.type, dv.formula1, dv.formula2) == check
    assert str(dv.sqref) == "C2:C6"


@pytest.mark.parametrize(
    "kwargs",
    [
        {"type": "list"},
        {"type": "whole"},
        {"type": "date", "min": "01/02/2026"},
        {"type": "colour"},
        {"type": "custom"},
    ],
)
def test_add_validation_errors(
    sales_grid: Path, settings: Settings, kwargs: dict[str, str]
) -> None:
    # Act / Assert
    with pytest.raises(SlimError):
        styling.add_validation("grid.xlsx", "Sales", "C2:C6", settings=settings, **kwargs)  # type: ignore[arg-type]


@pytest.mark.parametrize(
    ("rule", "value", "value2", "expected_type"),
    [
        ("greater_than", "15", None, "cellIs"),
        ("between", "10", "30", "cellIs"),
        ("text_contains", "No", None, "expression"),
        ("duplicates", None, None, "duplicateValues"),
        ("top_n", "2", None, "top10"),
        ("color_scale", None, None, "colorScale"),
        ("data_bar", None, None, "dataBar"),
    ],
)
def test_conditional_format_rules(
    sales_grid: Path,
    settings: Settings,
    rule: str,
    value: str | None,
    value2: str | None,
    expected_type: str,
) -> None:
    # Act
    styling.conditional_format(
        "grid.xlsx", "Sales", "C2:C6", rule, settings, value=value, value2=value2
    )

    # Assert
    rules = [r for cf in _slim(sales_grid)["Sales"].conditional_formatting for r in cf.rules]
    assert rules[0].type == expected_type


def test_conditional_format_needs_value(sales_grid: Path, settings: Settings) -> None:
    # Act / Assert
    with pytest.raises(SlimError) as info:
        styling.conditional_format("grid.xlsx", "Sales", "C2:C6", "greater_than", settings)
    assert "needs value" in info.value.message


# --- charts ----------------------------------------------------------------------------


@pytest.mark.parametrize("kind", ["column", "bar", "line", "pie", "area", "scatter"])
def test_add_chart_writes_a_chart_part(sales_grid: Path, settings: Settings, kind: str) -> None:
    # Act
    result = charts.add_chart("grid.xlsx", "Sales", kind, "Sales!B1:C6", settings, title="Totals")

    # Assert
    with zipfile.ZipFile(sales_grid.with_name("grid.slim.xlsx")) as archive:
        assert any(n.startswith("xl/charts/chart") for n in archive.namelist())
    assert f"{kind} chart" in result.summary
    assert "at Sales!E1" in result.summary


def test_add_chart_errors(sales_grid: Path, settings: Settings) -> None:
    # Act / Assert
    with pytest.raises(SlimError):
        charts.add_chart("grid.xlsx", "Sales", "donut", "A1:C6", settings)
    with pytest.raises(SlimError):
        charts.add_chart("grid.xlsx", "Sales", "line", "A1:A6", settings)


# --- pivot / split ---------------------------------------------------------------------


def test_pivot_group_and_cross_tab(sales_grid: Path, settings: Settings) -> None:
    # Act
    simple = reshape.pivot("grid.xlsx", ["Region"], ["Total"], settings)
    cross = reshape.pivot(
        "grid.xlsx", ["Region"], ["Total"], settings, columns="Month", to_sheet="Pivot"
    )

    # Assert
    assert simple.preview.rows == (("East", 5), ("North", 30), ("South", 70))
    assert cross.preview.columns == ("Region", "Feb", "Jan")
    assert cross.preview.rows[1] == ("North", 20, 10)
    assert cross.output is not None
    header = next(iter(_slim(sales_grid)["Pivot"].values))
    assert list(header) == ["Region", "Feb", "Jan"]


def test_pivot_out_path_and_errors(sales_grid: Path, settings: Settings) -> None:
    # Act
    result = reshape.pivot(
        "grid.xlsx", ["Region"], ["Total"], settings, agg="count", out_path="p.csv"
    )

    # Assert
    assert result.output is not None
    assert "count of Total" in (sales_grid.parent / "p.csv").read_text(encoding="utf-8")
    with pytest.raises(SlimError):
        reshape.pivot("grid.xlsx", ["Nope"], ["Total"], settings)
    with pytest.raises(SlimError):
        reshape.pivot("grid.xlsx", ["Region"], ["Total"], settings, agg="mode")


def test_split_into_sheets_and_files(sales_grid: Path, settings: Settings) -> None:
    # Act
    sheets = reshape.split("grid.xlsx", "Region", settings)
    files = reshape.split("grid.xlsx", "Month", settings, mode="files")

    # Assert
    wb = openpyxl.load_workbook(sales_grid.with_name("grid.split.xlsx"))
    assert wb.sheetnames == ["East", "North", "South"]
    assert "North=2" in sheets.summary
    assert (sales_grid.parent / "grid_split" / "grid_Jan.xlsx").exists()
    assert "2 file(s)" in files.summary
    with pytest.raises(SlimError):
        reshape.split("grid.xlsx", "Month", settings, mode="files", out_path="x.xlsx")


# --- recalc ----------------------------------------------------------------------------


def test_find_soffice_missing_has_install_hint(monkeypatch: pytest.MonkeyPatch) -> None:
    # Arrange
    monkeypatch.setattr(libreoffice, "_CANDIDATES", ("definitely-not-soffice",))

    # Act / Assert
    with pytest.raises(SlimError) as info:
        libreoffice.find_soffice(None)
    assert info.value.code == ErrorCode.MISSING_DEPENDENCY
    assert "LibreOffice" in info.value.hint


def test_build_command_and_profile(tmp_path: Path) -> None:
    # Act
    profile = libreoffice.write_profile(tmp_path / "profile")
    command = libreoffice.build_command("soffice", profile, tmp_path / "a.xlsx", tmp_path / "out")

    # Assert
    xcu = (profile / "user" / "registrymodifications.xcu").read_text(encoding="utf-8")
    assert "OOXMLRecalcMode" in xcu
    assert command[1].startswith("-env:UserInstallation=file:")
    assert "xlsx:Calc MS Excel 2007 XML" in command
    with pytest.raises(SlimError):
        libreoffice.build_command("soffice", profile, tmp_path / "a.csv", tmp_path)


def test_recalc_service_flow_with_stubbed_libreoffice(
    sales_grid: Path, settings: Settings, monkeypatch: pytest.MonkeyPatch, tmp_path: Path
) -> None:
    # Arrange: the stub plays LibreOffice by returning a copy whose formula has a cached value
    computed = tmp_path / "lo" / "out" / "grid.xlsx"
    computed.parent.mkdir(parents=True)
    shutil.copy(sales_grid, computed)
    monkeypatch.setattr(libreoffice, "find_soffice", lambda configured: "soffice")
    monkeypatch.setattr(libreoffice, "recalculate", lambda source, soffice: computed)

    # Act
    result = recalc.recalc("grid.xlsx", settings)

    # Assert
    assert result.target == "grid.slim.xlsx"
    assert result.summary == "recalculated 0 formula(s); 0 now have cached values"
    assert "LibreOffice re-saved" in result.warnings[0]


# --- interfaces ------------------------------------------------------------------------


def test_full_tier_handlers_and_generated_cli(
    sales_grid: Path,
    settings: Settings,
    monkeypatch: pytest.MonkeyPatch,
    capsys: pytest.CaptureFixture[str],
) -> None:
    # Arrange
    monkeypatch.setenv("SLIM_SHEET_ROOTS", str(sales_grid.parent))
    monkeypatch.setenv("SLIM_SHEET_TIER", "full")

    # Act
    listed = Handlers(settings).name_ops("grid.xlsx", "list")
    code = main(
        [
            "format_range",
            str(sales_grid),
            "Sales",
            "A1:C1",
            "--bold",
            "true",
            "--fill-color",
            "light_blue",
        ]
    )
    out = capsys.readouterr().out
    code_v = main(
        [
            "rules",
            "grid.slim.xlsx",
            "Sales",
            "C2:C6",
            "whole",
            "--value",
            "0",
            "--no-allow-blank",
        ]
    )
    out_v = capsys.readouterr().out

    # Assert
    assert listed == "grid.xlsx has no defined names."
    assert code == 0 and "formatted Sales!A1:C1: font, fill" in out
    assert code_v == 0 and "whole greaterThanOrEqual 0" in out_v
    dv = openpyxl.load_workbook(sales_grid.with_name("grid.slim.xlsx"))[
        "Sales"
    ].data_validations.dataValidation[0]
    assert dv.allow_blank is False


def test_full_tier_hidden_from_standard(
    sales_grid: Path, monkeypatch: pytest.MonkeyPatch, capsys: pytest.CaptureFixture[str]
) -> None:
    # Arrange
    monkeypatch.setenv("SLIM_SHEET_ROOTS", str(sales_grid.parent))
    monkeypatch.setenv("SLIM_SHEET_TIER", "standard")

    # Act
    code = main(["recalc", str(sales_grid)])

    # Assert
    assert code == 1
    assert "not enabled for tier 'standard'" in capsys.readouterr().out


# --- charts survive later edits --------------------------------------------------------


def _chart_parts(path: Path) -> int:
    with zipfile.ZipFile(path) as archive:
        return sum(1 for n in archive.namelist() if n.startswith("xl/charts/chart"))


def test_charts_survive_edits_and_follow_shifts_and_renames(
    sales_grid: Path, settings: Settings
) -> None:
    # Arrange
    from slim_sheet.services.write import WriteRequest, write

    charts.add_chart("grid.xlsx", "Sales", "line", "B1:C6", settings, anchor="H2")
    slim = sales_grid.with_name("grid.slim.xlsx")

    # Act
    written = write(WriteRequest("grid.slim.xlsx", "Sales", "E1", [["note"]]), settings)
    shifted = structure.rows_ops("grid.slim.xlsx", "Sales", "insert_rows", "1", settings, count=2)
    structure.sheet_ops("grid.slim.xlsx", "rename", "Sales", settings, new_name="Q1 Sales")

    # Assert
    assert _chart_parts(slim) == 1
    assert not any("drops" in w for w in written.warnings)
    assert "chart reference(s)" in shifted.summary
    chart = openpyxl.load_workbook(slim)["Q1 Sales"]._charts[0]
    assert chart.series[0].val.numRef.f == "'Q1 Sales'!$C$4:$C$8"
    assert (
        chart.series[0].cat.numRef.f == "'Q1 Sales'!$B$4:$B$8"
        or chart.series[0].cat.strRef.f == "'Q1 Sales'!$B$4:$B$8"
    )
    assert chart.anchor._from.row == 3  # H2 moved down two rows -> H4


def test_deleting_a_data_sheet_warns_about_charts_reading_it(
    sales_grid: Path, settings: Settings
) -> None:
    # Arrange
    structure.sheet_ops("grid.xlsx", "add", "Dash", settings)
    charts.add_chart("grid.slim.xlsx", "Dash", "pie", "Sales!B1:C6", settings)

    # Act
    result = structure.sheet_ops("grid.slim.xlsx", "delete", "Sales", settings)

    # Assert
    assert any("1 chart(s) on other sheets read 'Sales'" in w for w in result.warnings)


def test_lost_features_counts_parts_openpyxl_did_not_load(sales_grid: Path) -> None:
    # Arrange: a file with one chart, compared against a workbook object that loaded none
    from openpyxl.chart import BarChart, Reference

    from slim_sheet.adapters.writers.xlsx_writer import lost_features

    wb = openpyxl.load_workbook(sales_grid)
    chart = BarChart()
    chart.add_data(Reference(wb["Sales"], min_col=3, min_row=1, max_row=6), titles_from_data=True)
    wb["Sales"].add_chart(chart, "H2")
    wb.save(sales_grid)

    # Act
    loaded = openpyxl.load_workbook(sales_grid)
    empty = openpyxl.Workbook()

    # Assert
    assert lost_features(sales_grid, loaded) == []
    assert lost_features(sales_grid, empty) == ["1 chart(s) openpyxl could not read"]
