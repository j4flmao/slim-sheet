from __future__ import annotations

import json
from pathlib import Path

import pytest

from slim_sheet.adapters.formats import sniff
from slim_sheet.config import Settings
from slim_sheet.core.errors import ErrorCode, SlimError
from slim_sheet.core.models import FileFormat
from slim_sheet.interfaces.handlers import Handlers
from slim_sheet.services.overview import overview
from slim_sheet.services.read import ReadRequest, read
from tests.conftest import ORDER_ROWS


def test_overview_describes_sheets_headers_and_types(sales_xlsx: Path, settings: Settings) -> None:
    # Act
    info = overview("sales.xlsx", settings)

    # Assert
    assert info.format == FileFormat.XLSX
    by_name = {s.name: s for s in info.sheets}
    orders = by_name["Orders"]
    assert orders.header_row == 0
    assert [c.header for c in orders.columns] == ["OrderID", "Date", "Region", "Qty", "Total"]
    assert [c.dtype for c in orders.columns] == ["int", "date", "text", "int", "float"]
    assert by_name["Secret"].visible is False
    assert by_name["Empty"].used is None


def test_overview_render_is_compact(sales_xlsx: Path, settings: Settings) -> None:
    # Act
    text = Handlers(settings).overview("sales.xlsx")

    # Assert
    assert "**Orders**: used A1:E301" in text
    assert "[hidden]" in text
    assert len(text) < 1500


def test_read_pages_with_cursor_without_gaps(sales_xlsx: Path, settings: Settings) -> None:
    # Arrange
    seen: list[object] = []
    cursor: str | None = None

    # Act
    for _ in range(50):
        grid = read(ReadRequest("sales.xlsx", sheet="Orders", cursor=cursor), settings)
        seen.extend(row[0] for row in grid.rows)
        cursor = grid.next_cursor
        if cursor is None:
            break

    # Assert
    assert len(seen) == ORDER_ROWS + 1
    assert seen[0] == "OrderID"
    assert seen[-1] == 10_000 + ORDER_ROWS - 1


def test_read_respects_token_budget(sales_xlsx: Path, settings: Settings) -> None:
    # Act
    text = Handlers(settings).read("sales.xlsx", sheet="Orders", max_tokens=400)

    # Assert
    assert len(text) / 3.5 <= 400
    assert 'next: cursor="Orders!' in text


def test_read_range_and_long_cell_truncation(sales_xlsx: Path, settings: Settings) -> None:
    # Act
    text = Handlers(settings).read("sales.xlsx", sheet="notes", range="A1:E3")

    # Assert
    assert "...(+460)" in text
    assert "| D |" in text and "window A1:D3" in text


def test_read_unknown_sheet_hint_lists_sheets(sales_xlsx: Path, settings: Settings) -> None:
    # Act / Assert
    with pytest.raises(SlimError) as info:
        read(ReadRequest("sales.xlsx", sheet="Order"), settings)
    assert info.value.code == ErrorCode.SHEET_NOT_FOUND
    assert "Orders, Notes" in info.value.hint


def test_read_csv_as_single_sheet_json(people_csv: Path, settings: Settings) -> None:
    # Act
    payload = json.loads(Handlers(settings).read("people.csv", range="A1:C3", format="json"))

    # Assert
    assert payload["sheet"] == "data"
    assert payload["rows"][0] == ["name", "age", "city"]
    assert payload["rows"][1] == ["user0", "20", "City 0"]
    assert payload["used"] == "A1:C51"


def test_sandbox_denies_paths_outside_roots(sales_xlsx: Path, settings: Settings) -> None:
    # Act / Assert
    with pytest.raises(SlimError) as info:
        overview(str(Path(sales_xlsx.anchor) / "Windows" / "win.ini"), settings)
    assert info.value.code == ErrorCode.PATH_DENIED


def test_missing_file_and_bad_format(tmp_path: Path, settings: Settings) -> None:
    # Arrange
    junk = tmp_path / "blob.bin"
    junk.write_bytes(b"\x00\x01garbage")

    # Act / Assert
    with pytest.raises(SlimError) as missing:
        overview("nope.xlsx", settings)
    assert missing.value.code == ErrorCode.FILE_NOT_FOUND
    with pytest.raises(SlimError) as bad:
        sniff(junk)
    assert bad.value.code == ErrorCode.UNSUPPORTED_FORMAT


def test_bad_format_argument_is_rejected(sales_xlsx: Path, settings: Settings) -> None:
    # Act / Assert
    with pytest.raises(SlimError) as info:
        Handlers(settings).read("sales.xlsx", format="xml")
    assert info.value.code == ErrorCode.BAD_ARGUMENT
