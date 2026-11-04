"""Big data paths: the value-cell guard, the columnar (Parquet) cache, SQL anomalies and reports
from big sheets or files. Thresholds are lowered so small fixtures take the big path, and every
result is compared with the in-memory path on the same data."""

from __future__ import annotations

import datetime as dt
import os
import zipfile
from dataclasses import replace
from pathlib import Path

import openpyxl
import pytest

from slim_sheet.adapters import columnar, facts, limits
from slim_sheet.adapters.engine.cache import TABLE_CACHE
from slim_sheet.adapters.readers.calamine_reader import CalamineReader
from slim_sheet.config import Settings, load_settings
from slim_sheet.core.errors import ErrorCode, SlimError
from slim_sheet.core.models import CellRange
from slim_sheet.interfaces.handlers import Handlers
from slim_sheet.services.anomalies import anomalies

ROWS = 120


@pytest.fixture(autouse=True)
def fresh_tables() -> None:
    TABLE_CACHE.clear()


def _big_path(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setattr(columnar, "BIG_SHEET_CELLS", 100)
    monkeypatch.setattr(columnar, "BATCH_ROWS", 25)  # several batches, several row groups
    monkeypatch.setattr(columnar, "HEAD_ROWS", 10)


def _sales(path: Path) -> Path:
    """Title rows, a header in row 3, mixed types and planted anomalies."""
    wb = openpyxl.Workbook()
    ws = wb.active
    assert ws is not None
    ws.title = "Sales"
    ws["A1"] = "Sales 2026"
    ws.append([])
    ws.append(["Order ID", "Date", "Region", "Units", "Price", "Note"])
    for i in range(ROWS):
        region = ["North", "South", "East", "West"][i % 4]
        ws.append([1000 + i, dt.date(2026, 1, 1) + dt.timedelta(days=i), region, 5 + i % 7,
                   round(10 + (i % 9) * 1.5, 2), "ok" if i % 2 else 7])  # fmt: skip
    ws["C40"] = " north"  # whitespace + variant
    ws["D50"] = 9000  # outlier
    ws["A60"] = 1010  # duplicate key
    wb.save(path)
    return path


# --- guard -----------------------------------------------------------------------------------


def test_formatted_empty_cells_do_not_count(tmp_path: Path, settings: Settings) -> None:
    # Arrange: 10 values, 5,000 formatted empty cells
    wb = openpyxl.Workbook()
    ws = wb.active
    assert ws is not None
    ws.append(list(range(10)))
    for row in range(2, 102):
        for col in range(1, 51):
            ws.cell(row=row, column=col).number_format = "0.00"
    wb.save(tmp_path / "styled.xlsx")

    # Act
    counted = limits.count_cells(tmp_path / "styled.xlsx", 10**9)

    # Assert
    assert counted == 10
    assert "styled.xlsx" in Handlers(replace(settings, max_cells=100)).overview("styled.xlsx")


def test_edits_have_their_own_lower_limit_with_a_teaching_hint(
    tmp_path: Path, settings: Settings
) -> None:
    # Arrange
    _sales(tmp_path / "s.xlsx")
    tight = replace(settings, max_edit_cells=50)

    # Act / Assert
    assert "Sales" in Handlers(tight).overview("s.xlsx")  # reading is fine
    with pytest.raises(SlimError) as caught:
        Handlers(tight).write("s.xlsx", "Sales", "H1", [["x"]])
    assert caught.value.code == ErrorCode.TOO_LARGE
    assert "SLIM_SHEET_MAX_EDIT_CELLS" in caught.value.hint
    assert "Do not cut the data down" in caught.value.hint
    assert load_settings({"SLIM_SHEET_MAX_EDIT_CELLS": "7"}).max_edit_cells == 7


# --- columnar cache --------------------------------------------------------------------------


@pytest.mark.parametrize(
    "window",
    [
        CellRange(0, 0, 14, 5),
        CellRange(8, 1, 40, 4),
        CellRange(100, 0, 125, 7),
        CellRange(3, 5, 3, 5),
    ],
)
def test_big_sheet_reads_match_the_normal_reader(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch, window: CellRange
) -> None:
    # Arrange
    path = _sales(tmp_path / "s.xlsx")
    with CalamineReader(path) as normal:
        expected = normal.read("Sales", window)
    _big_path(monkeypatch)

    # Act
    with CalamineReader(path) as big:
        assert big.store("Sales") is not None
        got = big.read("Sales", window)

    # Assert: the mixed 'Note' column comes back as text; everything else is identical
    assert len(got) == window.n_rows and all(len(r) == window.n_cols for r in got)
    for want_row, got_row in zip(expected, got, strict=True):
        for want, have in zip(want_row, got_row, strict=True):
            assert have == want or (isinstance(have, str) and str(want).rstrip(".0") == have)


def test_cache_is_reused_and_rebuilt_when_the_file_changes(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    # Arrange
    path = _sales(tmp_path / "s.xlsx")
    _big_path(monkeypatch)
    with CalamineReader(path) as first:
        store = first.store("Sales")
    assert store is not None

    # Act: a second reader must not parse the sheet again
    with monkeypatch.context() as scoped:
        scoped.setattr(CalamineReader, "_sheet", lambda *_: pytest.fail("sheet parsed again"))
        with CalamineReader(path) as second:
            cached = second.store("Sales")
    wb = openpyxl.load_workbook(path)
    wb["Sales"]["D4"] = 1234
    wb.save(path)
    os.utime(path, (1, 2_000_000_000))
    with CalamineReader(path) as third:
        rebuilt = third.store("Sales")

    # Assert
    assert cached is not None and cached.parquet == store.parquet
    assert rebuilt is not None and rebuilt.parquet != store.parquet
    assert not store.parquet.exists()  # the old version was pruned
    assert rebuilt.read(CellRange(3, 3, 3, 3)) == [[1234]]


def test_types_are_unified_across_batches(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    # Arrange: whole numbers first, decimals later, text at the very end
    wb = openpyxl.Workbook()
    ws = wb.active
    assert ws is not None
    ws.append(["Mixed", "Grows"])
    for i in range(60):
        ws.append([i if i < 59 else "n/a", i if i < 30 else i + 0.5])
    wb.save(tmp_path / "t.xlsx")
    _big_path(monkeypatch)

    # Act
    with CalamineReader(tmp_path / "t.xlsx") as reader:
        store = reader.store("Sheet")
    assert store is not None
    schema = store.dataset().schema

    # Assert
    assert str(schema.field("Mixed").type) == "string"
    assert str(schema.field("Grows").type) == "double"


def test_query_profile_and_pivot_agree_on_both_paths(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch, settings: Settings
) -> None:
    # Arrange
    _sales(tmp_path / "s.xlsx")
    sql = 'SELECT "Region", count(*) AS n, sum("Units") AS u FROM "Sales" GROUP BY 1 ORDER BY 1'
    normal = Handlers(settings)
    expected = (normal.query("s.xlsx", sql), normal.pivot("s.xlsx", rows=["Region"],
                values=["Units"]))  # fmt: skip
    TABLE_CACHE.clear()
    _big_path(monkeypatch)

    # Act
    big = Handlers(settings)
    got = (big.query("s.xlsx", sql), big.pivot("s.xlsx", rows=["Region"], values=["Units"]))
    profile = big.profile("s.xlsx")

    # Assert
    assert got == expected
    assert f"{ROWS} rows" in profile


def test_sql_anomalies_match_the_python_checks(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch, settings: Settings
) -> None:
    # Arrange
    _sales(tmp_path / "s.xlsx")
    expected = anomalies("s.xlsx", settings, max_results=200)
    _big_path(monkeypatch)

    # Act
    got = anomalies("s.xlsx", settings, max_results=200)

    # Assert
    def key(result: object) -> set[tuple[str, str]]:
        return {(f.address, f.kind) for f in result.findings if f.kind != "type_mismatch"}  # type: ignore[attr-defined]

    assert key(got) == key(expected)
    assert {"duplicate_key", "outlier", "variant", "whitespace"} <= {k for k, _ in got.counts}


# --- big text files --------------------------------------------------------------------------


def _csv(path: Path, rows: int = 300) -> Path:
    lines = ["Order ID,Region,Units,Revenue"]
    lines += [
        f"{i},{['North', 'South'][i % 2]},{1 + i % 5},{10.5 * (1 + i % 5)}" for i in range(rows)
    ]
    path.write_text("\n".join(lines) + "\n", encoding="utf-8")
    return path


def test_big_csv_is_cached_as_parquet_and_queried(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch, settings: Settings
) -> None:
    # Arrange
    _csv(tmp_path / "o.csv")
    monkeypatch.setattr(columnar, "FLAT_CACHE_BYTES", 1)
    monkeypatch.setattr(columnar, "BIG_SHEET_CELLS", 100)

    # Act
    text = Handlers(settings).query(
        "o.csv", 'SELECT "Region", sum("Units") u FROM data GROUP BY 1 ORDER BY 1'
    )
    audit = Handlers(settings).anomalies("o.csv")

    # Assert
    assert "| North | 450 |" in text and "| South | 450 |" in text
    assert list((facts.cache_dir() / "flat").glob("*.parquet"))
    assert "300 rows" in audit


def test_report_from_a_big_csv_goes_to_a_new_workbook(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch, settings: Settings
) -> None:
    # Arrange
    _csv(tmp_path / "o.csv")
    monkeypatch.setattr(columnar, "FLAT_CACHE_BYTES", 1)
    monkeypatch.setattr(columnar, "BIG_SHEET_CELLS", 100)
    h = Handlers(settings)

    # Act
    kpis = ["Units=sum:Units", "Orders=count:Order ID", "Avg=avg:Revenue|vs:Units:units"]
    charts = ["rank:Region:Revenue", "table:Region:Units"]
    text = h.build_report("o.csv", "data", "Orders", kpis=kpis, charts=charts, overwrite=True)
    with pytest.raises(SlimError) as again:
        h.build_report("o.csv", "data", "Orders", kpis=["Units=sum:Units"])

    # Assert
    board = openpyxl.load_workbook(tmp_path / "o.report.xlsx")["Dashboard"]
    assert board["B5"].value == 900  # a computed value, not a formula into data that is not there
    assert board["D5"].value == 300
    assert board["F6"].value.startswith("=IFERROR(")  # the comparison stays a live formula
    assert board["J10"].value == "=450.0"  # second grid slot: the table's value column
    assert "aggregated" in text and "o.csv is untouched" not in text
    assert again.value.code == ErrorCode.BAD_ARGUMENT and "overwrite" in again.value.hint
    assert any("computed from all rows" in str(c.value) for row in board.iter_rows() for c in row)


def test_scans_of_big_files_are_remembered(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    # Arrange
    path = _sales(tmp_path / "s.xlsx")
    monkeypatch.setattr(facts, "FACT_MIN_BYTES", 0)
    first = limits.count_cells(path, 10**9)
    limits._cache.clear()

    # Act: a fresh process would find the stored count without reading the XML again
    monkeypatch.setattr(limits, "_count", lambda *_: pytest.fail("counted again"))
    second = limits.count_cells(path, 10**9)

    # Assert
    assert first == second > ROWS
    with zipfile.ZipFile(path) as archive:
        assert "xl/worksheets/sheet1.xml" in archive.namelist()
