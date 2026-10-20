from __future__ import annotations

import datetime as dt

import pytest

from slim_sheet.core.a1 import col_to_index, format_range, index_to_col, parse_range
from slim_sheet.core.budget import cell_text, clamp_budget, estimate_tokens
from slim_sheet.core.detect import guess_header_row, infer_type
from slim_sheet.core.errors import ErrorCode, SlimError
from slim_sheet.core.models import CellRange


@pytest.mark.parametrize(("letters", "index"), [("A", 0), ("Z", 25), ("AA", 26), ("XFD", 16383)])
def test_column_letters_round_trip(letters: str, index: int) -> None:
    # Arrange / Act
    parsed = col_to_index(letters)
    rendered = index_to_col(index)

    # Assert
    assert parsed == index
    assert rendered == letters


def test_parse_range_normalises_reversed_corners() -> None:
    # Act
    rng = parse_range("D9:$B$2")

    # Assert
    assert rng == CellRange(1, 1, 8, 3)
    assert format_range(rng) == "B2:D9"


@pytest.mark.parametrize("bad", ["", "A0", "1A", "Sheet1!A1", "A1:B2:C3", "ZZZZ1"])
def test_parse_range_rejects_invalid_input(bad: str) -> None:
    # Act / Assert
    with pytest.raises(SlimError) as info:
        parse_range(bad)
    assert info.value.code == ErrorCode.BAD_RANGE


def test_cell_text_truncates_long_values_with_marker() -> None:
    # Act
    text = cell_text("a" * 50, max_chars=10)

    # Assert
    assert text == "aaaaaaaaaa...(+40)"


@pytest.mark.parametrize(
    ("value", "expected"),
    [
        (None, ""),
        (True, "TRUE"),
        (3.0, "3"),
        (2.5, "2.5"),
        (dt.datetime(2026, 1, 2), "2026-01-02"),
        ("a\nb", "a\\nb"),
    ],
)
def test_cell_text_renders_scalar_types(value: object, expected: str) -> None:
    # Act / Assert
    assert cell_text(value, 40) == expected  # type: ignore[arg-type]


def test_budget_clamps_to_bounds() -> None:
    # Act / Assert
    assert clamp_budget(None, 1500) == 1500
    assert clamp_budget(10, 1500) == 300
    assert clamp_budget(10**9, 1500) == 20_000
    assert estimate_tokens("x" * 35) == 10


def test_guess_header_row_skips_title_rows() -> None:
    # Arrange
    rows = [
        ["Report 2026", None, None],
        [None, None, None],
        ["id", "name", "amount"],
        [1, "a", 2.5],
    ]

    # Act / Assert
    assert guess_header_row(rows) == 2


def test_guess_header_row_none_for_numeric_grid() -> None:
    # Act / Assert
    assert guess_header_row([[1, 2], [3, 4]]) is None
    assert guess_header_row([]) is None


@pytest.mark.parametrize(
    ("values", "expected"),
    [
        ([1, 2.0], "int"),
        ([1, 2.5], "float"),
        (["a", None], "text"),
        ([], "empty"),
        ([1, "a"], "mixed"),
    ],
)
def test_infer_type(values: list[object], expected: str) -> None:
    # Act / Assert
    assert infer_type(values) == expected  # type: ignore[arg-type]
