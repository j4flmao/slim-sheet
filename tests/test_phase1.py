from __future__ import annotations

import json
from pathlib import Path

import openpyxl
import pytest

from slim_sheet.config import Settings
from slim_sheet.core.errors import ErrorCode, SlimError
from slim_sheet.core.formula import analyze
from slim_sheet.core.models import CellRange
from slim_sheet.core.regions import column_names, detect_regions
from slim_sheet.interfaces.cli import main
from slim_sheet.interfaces.handlers import Handlers
from slim_sheet.services.find import find
from slim_sheet.services.profile import profile
from slim_sheet.services.query import query
from slim_sheet.services.write import WriteRequest, write

# --- regions -------------------------------------------------------------------------


def test_detect_regions_splits_side_by_side_tables_and_drops_title() -> None:
    # Arrange
    rows = [
        ["Title", None, None, None, None],
        [None, None, None, None, None],
        [None, None, None, None, None],
        ["a", "b", None, "x", "y"],
        [1, 2, None, 3, 4],
    ]

    # Act
    regions = detect_regions(rows, top=0, left=0)

    # Assert
    assert CellRange(3, 0, 4, 1) in regions
    assert CellRange(3, 3, 4, 4) in regions
    assert all(r.top != 0 for r in regions)


def test_detect_regions_empty_grid() -> None:
    # Act / Assert
    assert detect_regions([], 0, 0) == []
    assert detect_regions([[None, ""]], 0, 0) == []


def test_column_names_dedupe_and_fallback() -> None:
    # Act
    names = column_names(["Qty", None, "qty", 2024.0], ["A", "B", "C", "D"])

    # Assert
    assert names == ["Qty", "B", "qty_2", "2024"]


# --- query ---------------------------------------------------------------------------


def test_query_aggregates_over_sheet(sales_xlsx: Path, settings: Settings) -> None:
    # Act
    result = query(
        "sales.xlsx",
        'SELECT "Region", count(*) AS n FROM "Orders" GROUP BY 1 ORDER BY 1',
        settings,
    )

    # Assert
    assert result.columns == ("Region", "n")
    assert result.rows == (("North", 150), ("South", 150))
    assert result.next_offset is None


def test_query_pages_with_offset(sales_xlsx: Path, settings: Settings) -> None:
    # Arrange
    sql = 'SELECT "OrderID" FROM orders ORDER BY 1'

    # Act
    first = query("sales.xlsx", sql, settings, max_tokens=300)
    assert first.next_offset is not None
    second = query("sales.xlsx", sql, settings, offset=first.next_offset, max_tokens=300)

    # Assert
    assert first.total_rows == 300
    assert second.rows[0][0] == first.rows[-1][0] + 1  # type: ignore[operator]


def test_query_registers_secondary_regions(messy_xlsx: Path, settings: Settings) -> None:
    # Act
    main_table = query("messy.xlsx", 'SELECT sum("Qty") FROM "Report"', settings)
    side_sql = """SELECT "Note" FROM "Report.2" WHERE "Code" = 'A'"""
    side_table = query("messy.xlsx", side_sql, settings)

    # Assert
    assert main_table.rows == ((14,),)
    assert side_table.rows == (("refund issued",),)


def test_query_csv_table_named_data(people_csv: Path, settings: Settings) -> None:
    # Act
    result = query("people.csv", "SELECT max(age) FROM data", settings)

    # Assert
    assert result.rows == ((59,),)


@pytest.mark.parametrize(
    ("sql", "code"),
    [
        ("DELETE FROM data", ErrorCode.SQL_WRITE_DENIED),
        ("COPY data TO 'x.csv'", ErrorCode.SQL_WRITE_DENIED),
        ("SELECT 1; SELECT 2", ErrorCode.BAD_SQL),
        ("", ErrorCode.BAD_SQL),
        ("SELECT * FROM nope", ErrorCode.BAD_SQL),
    ],
)
def test_query_rejects_unsafe_or_invalid_sql(
    people_csv: Path, settings: Settings, sql: str, code: ErrorCode
) -> None:
    # Act / Assert
    with pytest.raises(SlimError) as info:
        query("people.csv", sql, settings)
    assert info.value.code == code


def test_query_cannot_read_other_files(people_csv: Path, settings: Settings) -> None:
    # Arrange
    secret = people_csv.parent / "secret.csv"
    secret.write_text("k\nv\n", encoding="utf-8")

    # Act / Assert
    with pytest.raises(SlimError) as info:
        query("people.csv", f"SELECT * FROM read_csv('{secret.as_posix()}')", settings)
    assert info.value.code == ErrorCode.BAD_SQL


def test_bad_sql_hint_lists_schema(sales_xlsx: Path, settings: Settings) -> None:
    # Act / Assert
    with pytest.raises(SlimError) as info:
        query("sales.xlsx", 'SELECT "Totl" FROM "Orders"', settings)
    assert '"Total" DOUBLE' in info.value.hint


# --- profile -------------------------------------------------------------------------


def test_profile_reports_stats_and_top_values(sales_xlsx: Path, settings: Settings) -> None:
    # Act
    result = profile("sales.xlsx", settings, sheet="Orders", columns=["Region", "Qty"])

    # Assert
    assert result.rows == 300
    region, qty = result.columns
    assert region.distinct == 2
    assert region.top[0][1] == 150
    assert qty.min == "1"
    assert qty.max == "7"
    assert qty.null_pct == 0


def test_profile_distinct_is_exact_and_unique_columns_skip_top(
    sales_xlsx: Path, settings: Settings
) -> None:
    # Act
    (order_id,) = profile("sales.xlsx", settings, sheet="Orders", columns=["OrderID"]).columns

    # Assert
    assert order_id.distinct == 300
    assert order_id.top == ()


def test_profile_unknown_column_hint(sales_xlsx: Path, settings: Settings) -> None:
    # Act / Assert
    with pytest.raises(SlimError) as info:
        profile("sales.xlsx", settings, sheet="Orders", columns=["Nope"])
    assert "OrderID" in info.value.hint


# --- find ----------------------------------------------------------------------------


def test_find_returns_addresses_across_sheets(messy_xlsx: Path, settings: Settings) -> None:
    # Act
    result = find("messy.xlsx", "REFUND", settings)

    # Assert
    assert [(h.sheet, h.row, h.col) for h in result.hits] == [("Report", 3, 5)]


def test_find_regex_and_truncation(sales_xlsx: Path, settings: Settings) -> None:
    # Act
    result = find("sales.xlsx", r"^No(rth)?$", settings, sheet="Orders", regex=True, max_results=5)

    # Assert
    assert len(result.hits) == 5
    assert result.truncated is True


def test_find_rejects_bad_regex(sales_xlsx: Path, settings: Settings) -> None:
    # Act / Assert
    with pytest.raises(SlimError) as info:
        find("sales.xlsx", "(", settings, regex=True)
    assert info.value.code == ErrorCode.BAD_ARGUMENT


# --- write ---------------------------------------------------------------------------


def test_write_goes_to_working_copy_and_accumulates(sales_xlsx: Path, settings: Settings) -> None:
    # Act
    first = write(WriteRequest("sales.xlsx", "Orders", "G1", [["Flag"], [True]]), settings)
    second = write(WriteRequest("sales.xlsx", "Summary", "B2", [["Total", 1]]), settings)

    # Assert
    copy = sales_xlsx.with_name("sales.slim.xlsx")
    assert first.target == second.target == copy.name
    assert first.backup is None
    assert second.backup is not None
    assert second.created_sheet is True
    wb = openpyxl.load_workbook(copy)
    assert wb["Orders"]["G1"].value == "Flag"
    assert wb["Summary"]["C2"].value == 1
    original = openpyxl.load_workbook(sales_xlsx)
    assert original["Orders"]["G1"].value is None


def test_write_in_place_requires_overwrite(sales_xlsx: Path, settings: Settings) -> None:
    # Act / Assert
    with pytest.raises(SlimError) as info:
        write(WriteRequest("sales.xlsx", "Orders", "A1", [["x"]], out_path="sales.xlsx"), settings)
    assert info.value.code == ErrorCode.WRITE_CONFLICT
    result = write(WriteRequest("sales.xlsx", "Orders", "A1", [["x"]], overwrite=True), settings)
    assert result.target == "sales.xlsx"
    assert result.backup is not None


@pytest.mark.parametrize(
    "formula",
    ['=WEBSERVICE("http://x")', "=cmd|' /C calc'!A0", "='[other.xlsx]Sheet1'!A1"],
)
def test_write_blocks_external_formulas(sales_xlsx: Path, settings: Settings, formula: str) -> None:
    # Act / Assert
    with pytest.raises(SlimError) as info:
        write(WriteRequest("sales.xlsx", "Orders", "A1", [[formula]]), settings)
    assert info.value.code == ErrorCode.FORMULA_DENIED


def test_write_allows_safe_formula_and_warns(sales_xlsx: Path, settings: Settings) -> None:
    # Act
    result = write(WriteRequest("sales.xlsx", "Orders", "H2", [['=A2&"|"&C2']]), settings)

    # Assert
    assert any("formula" in w for w in result.warnings)


def test_write_shape_mismatch_and_bad_format(
    sales_xlsx: Path, people_csv: Path, settings: Settings
) -> None:
    # Act / Assert
    with pytest.raises(SlimError) as shape:
        write(WriteRequest("sales.xlsx", "Orders", "A1:B2", [[1, 2, 3]]), settings)
    assert "range='A1'" in shape.value.hint
    with pytest.raises(SlimError) as fmt:
        write(WriteRequest("people.csv", "data", "A1", [[1]]), settings)
    assert fmt.value.code == ErrorCode.UNSUPPORTED_FORMAT


# --- explain_formula -----------------------------------------------------------------


def test_analyze_orders_calls_inside_out() -> None:
    # Act
    report = analyze("=IFERROR(VLOOKUP(A2,Prices!A:B,2,FALSE),0)")

    # Assert
    assert [c.name for c in report.calls] == ["VLOOKUP", "IFERROR"]
    assert report.calls[0].args == ("A2", "Prices!A:B", "2", "FALSE")
    assert report.references == ("A2", "Prices!A:B")
    assert report.errors == ()


@pytest.mark.parametrize(
    ("formula", "fragment"),
    [
        ("=SUM((A1)", "missing ')'"),
        ("=SUM(A1))", "unmatched"),
        ("=VLOOKUP(A1,B:C)", "VLOOKUP takes 3-4"),
        ("=SUM(A1;A2)", "';'"),
        ("=1+", "ends with operator"),
    ],
)
def test_analyze_reports_errors(formula: str, fragment: str) -> None:
    # Act
    report = analyze(formula)

    # Assert
    assert any(fragment in e for e in report.errors), report.errors


def test_analyze_warnings_for_typos_and_modern_functions() -> None:
    # Act
    report = analyze("XLOOKUP(A1,B:B,C:C)+VLOKUP(A1,B:C,2)")

    # Assert
    joined = " ".join(report.warnings)
    assert "added the leading '='" in joined
    assert "XLOOKUP needs Excel 2021+" in joined
    assert "did you mean VLOOKUP" in joined


def test_analyze_handles_arrays_and_xlfn_prefix() -> None:
    # Act
    report = analyze("=_xlfn.XLOOKUP(1,{1,2;3,4},A1:A4)")

    # Assert
    assert report.errors == ()
    assert report.calls[0].name == "XLOOKUP"
    assert report.calls[0].args[1] == "{1,2;3,4}"


def test_explain_formula_reads_cell(messy_xlsx: Path, settings: Settings) -> None:
    # Act
    text = Handlers(settings).explain_formula(path="messy.xlsx", cell="Calc!C1")

    # Assert
    assert "(Calc!C1)" in text
    assert "1. VLOOKUP(A1, Report!A3:C6, 2, FALSE)" in text
    assert "Errors: none" in text


def test_explain_formula_on_value_cell(messy_xlsx: Path, settings: Settings) -> None:
    # Act / Assert
    with pytest.raises(SlimError) as info:
        Handlers(settings).explain_formula(path="messy.xlsx", cell="A1", sheet="Calc")
    assert "holds a value" in info.value.message


# --- CLI -----------------------------------------------------------------------------


def test_cli_query_and_error_exit_code(
    sales_xlsx: Path, monkeypatch: pytest.MonkeyPatch, capsys: pytest.CaptureFixture[str]
) -> None:
    # Arrange
    monkeypatch.setenv("SLIM_SHEET_ROOTS", str(sales_xlsx.parent))

    # Act
    ok = main(["query", str(sales_xlsx), 'SELECT count(*) AS n FROM "Orders"'])
    out = capsys.readouterr().out
    bad = main(["read", str(sales_xlsx), "--sheet", "Nope"])
    err = json.loads(capsys.readouterr().out)

    # Assert
    assert ok == 0
    assert "| 300 |" in out
    assert bad == 1
    assert err["error"] == "sheet_not_found"


def test_cli_write_parses_json_values(
    sales_xlsx: Path, monkeypatch: pytest.MonkeyPatch, capsys: pytest.CaptureFixture[str]
) -> None:
    # Arrange
    monkeypatch.setenv("SLIM_SHEET_ROOTS", str(sales_xlsx.parent))

    # Act
    code = main(["write", str(sales_xlsx), "Orders", "J1", '[["a", 1]]'])

    # Assert
    assert code == 0
    assert "Wrote 2 cell(s) to sales.slim.xlsx" in capsys.readouterr().out
