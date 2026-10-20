"""Unit tests for the eval scoring, task loading and report helpers (no LLM, no server)."""

from __future__ import annotations

import json
from pathlib import Path
from typing import Any

import openpyxl
import pytest

from evals.report import render_results, summarize
from evals.scoring import (
    Check,
    answer_part,
    check_answer,
    check_written,
    extract_cells,
    extract_numbers,
    load_tasks,
    parse_check,
    parse_text_calls,
    to_openai_schema,
    validate_call,
    written_mismatches,
)

READ_SCHEMA: dict[str, Any] = {
    "type": "object",
    "properties": {
        "path": {"type": "string", "title": "Path"},
        "sheet": {"anyOf": [{"type": "string"}, {"type": "null"}], "default": None},
    },
    "required": ["path"],
}


def test_answer_part_uses_last_marker() -> None:
    # Arrange
    text = "I think ANSWER: 1 is wrong.\nAnswer: 42"

    # Act
    result = answer_part(text)

    # Assert
    assert result == "42"


def test_answer_part_without_marker_returns_everything() -> None:
    # Act / Assert
    assert answer_part("") == ""
    assert answer_part("North is 7462.5") == "North is 7462.5"


def test_extract_numbers_handles_commas_currency_and_codes() -> None:
    # Arrange
    text = "Total $14,962.50; invoice INV-1003; delta -3.5"

    # Act
    numbers = extract_numbers(text)

    # Assert
    assert numbers == [14962.5, 1003.0, -3.5]


def test_number_check_passes_within_tolerance_and_fails_outside() -> None:
    # Arrange
    check = Check("number", value=49.875, tolerance=0.01)

    # Act / Assert
    assert check_answer(check, "The average is about 49.88.\nANSWER: 49.88")
    assert not check_answer(check, "ANSWER: 50")
    assert not check_answer(check, "")


def test_text_check_is_case_insensitive_and_word_bounded() -> None:
    # Arrange
    check = Check("text", value="ink")

    # Act / Assert
    assert check_answer(check, "ANSWER: Ink")
    assert not check_answer(check, "ANSWER: I think it is pen")


def test_all_check_requires_every_value_and_respects_word_boundaries() -> None:
    # Arrange
    check = Check("all", values=("SUMM", "SUM"))

    # Act / Assert
    assert check_answer(check, "ANSWER: SUMM is a typo for SUM")
    assert not check_answer(check, "ANSWER: SUMM is unknown")


def test_any_check_matches_substrings() -> None:
    # Arrange
    check = Check("any", values=("parenthes", "unbalanced"))

    # Act / Assert
    assert check_answer(check, "ANSWER: missing closing parentheses")
    assert not check_answer(check, "ANSWER: looks fine")


def test_extract_cells_reads_sheets_quotes_and_dollars() -> None:
    # Act
    cells = extract_cells("Notes!B3, 'Q1 Targets'!$C$2 and A1:A2; not user0 or INV-1003")

    # Assert
    assert cells == {("notes", "B3"), ("q1 targets", "C2"), (None, "A1"), (None, "A2")}


def test_cells_check_strict_rejects_extra_addresses() -> None:
    # Arrange
    check = Check("cells", values=("Notes!B3", "Notes!B6"))

    # Act / Assert
    assert check_answer(check, "ANSWER: Notes!B3, Notes!B6")
    assert check_answer(check, "ANSWER: B3 and B6")
    assert not check_answer(check, "ANSWER: Notes!B3")
    assert not check_answer(check, "ANSWER: Notes!B3, Notes!B6, Notes!D1")
    assert not check_answer(check, "ANSWER: Orders!B3, Orders!B6")


def test_cells_check_non_strict_allows_extra_addresses() -> None:
    # Arrange
    check = Check("cells", values=("Report!F4",))

    # Act / Assert
    assert check_answer(check, "- Report!F4: refund\n- Calc!A1: 1", strict=False)


def test_check_answer_refuses_written_kind() -> None:
    # Act / Assert
    with pytest.raises(ValueError, match="check_written"):
        check_answer(Check("written", file="x.slim.xlsx", cells={"A!A1": 1}), "ANSWER: done")


@pytest.mark.parametrize(
    "raw",
    [
        {"kind": "magic"},
        {"kind": "number"},
        {"kind": "cells", "values": []},
        {"kind": "written", "file": "x.slim.xlsx"},
    ],
)
def test_parse_check_rejects_malformed_checks(raw: dict[str, Any]) -> None:
    # Act / Assert
    with pytest.raises(ValueError):
        parse_check("t1", raw)


def test_load_tasks_rejects_duplicate_ids(tmp_path: Path) -> None:
    # Arrange
    task = {
        "id": "a",
        "prompt": "p",
        "tools": ["query"],
        "check": {"kind": "number", "value": 1},
        "reference": [{"tool": "query", "args": {"path": "x.csv", "sql": "SELECT 1"}}],
    }
    path = tmp_path / "tasks.json"
    path.write_text(json.dumps({"tasks": [task, task]}), encoding="utf-8")

    # Act / Assert
    with pytest.raises(ValueError, match="Duplicate"):
        load_tasks(path)


def test_load_tasks_rejects_empty_file(tmp_path: Path) -> None:
    # Arrange
    path = tmp_path / "tasks.json"
    path.write_text("{}", encoding="utf-8")

    # Act / Assert
    with pytest.raises(ValueError, match="No tasks"):
        load_tasks(path)


def test_shipped_task_file_loads_about_forty_tasks() -> None:
    # Act
    tasks = load_tasks(Path(__file__).with_name("tasks.json"))

    # Assert
    assert 40 <= len(tasks) <= 50
    assert all(call.tool in t.tools for t in tasks for call in t.reference)


def _workbook(path: Path, cells: dict[str, Any]) -> None:
    wb = openpyxl.Workbook()
    ws = wb.active
    assert ws is not None
    ws.title = "Orders"
    for address, value in cells.items():
        ws[address] = value
    wb.save(path)


def test_check_written_compares_values_with_tolerance(tmp_path: Path) -> None:
    # Arrange
    _workbook(tmp_path / "sales.slim.xlsx", {"G1": "grand total ", "H1": 14962.5})
    check = Check(
        "written", file="sales.slim.xlsx", cells={"Orders!G1": "Grand Total", "Orders!H1": 14962.5}
    )

    # Act / Assert
    assert check_written(check, tmp_path)


def test_written_mismatches_reports_missing_file_sheet_and_value(tmp_path: Path) -> None:
    # Arrange
    _workbook(tmp_path / "sales.slim.xlsx", {"H1": "14962.5x"})
    check = Check("written", file="sales.slim.xlsx", cells={"Orders!H1": 14962.5, "Sum!A1": 1})

    # Act
    problems = written_mismatches(check, tmp_path)
    missing = written_mismatches(Check("written", file="nope.xlsx", cells={"A!A1": 1}), tmp_path)

    # Assert
    assert problems == ["Orders!H1: got '14962.5x', want 14962.5", "Sum!A1: sheet missing"]
    assert missing == ["nope.xlsx was not written"]


def test_validate_call_accepts_valid_args_as_json_string() -> None:
    # Act
    verdict = validate_call("read", '{"path": "sales.xlsx", "sheet": null}', {"read": READ_SCHEMA})

    # Assert
    assert verdict.valid
    assert verdict.args == {"path": "sales.xlsx", "sheet": None}


@pytest.mark.parametrize(
    ("name", "raw", "reason"),
    [
        ("reed", "{}", "unknown tool"),
        ("read", "{not json", "not JSON"),
        ("read", "[1]", "JSON object"),
        ("read", '{"sheet": "Orders"}', "'path' is a required property"),
        ("read", '{"path": 3}', "is not of type 'string'"),
    ],
)
def test_validate_call_rejects_bad_calls(name: str, raw: str, reason: str) -> None:
    # Act
    verdict = validate_call(name, raw, {"read": READ_SCHEMA})

    # Assert
    assert not verdict.valid
    assert reason in verdict.reason


def test_validate_call_treats_empty_string_as_no_args() -> None:
    # Act
    verdict = validate_call("read", "", {"read": READ_SCHEMA})

    # Assert
    assert not verdict.valid
    assert "required" in verdict.reason


def test_to_openai_schema_drops_titles_and_collapses_optional() -> None:
    # Act
    converted = to_openai_schema(READ_SCHEMA)

    # Assert
    assert converted["properties"]["path"] == {"type": "string"}
    assert converted["properties"]["sheet"] == {"type": "string", "default": None}
    assert converted["required"] == ["path"]


def _run(success: bool, calls: int, valid: int, tokens: int) -> dict[str, Any]:
    return {
        "task_id": "t",
        "success": success,
        "answer": "x | y",
        "tool_calls": calls,
        "valid_calls": valid,
        "tool_errors": 0,
        "tokens": tokens,
        "tokens_estimated": False,
        "used_tools": [],
        "stop_reason": "answered",
        "detail": "",
    }


def test_summarize_computes_rates() -> None:
    # Arrange
    runs = [_run(True, 2, 2, 100), _run(False, 2, 1, 300)]

    # Act
    metrics = summarize(runs)

    # Assert
    assert metrics["success_rate"] == 0.5
    assert metrics["validity_rate"] == 0.75
    assert metrics["avg_tokens"] == 200


def test_summarize_handles_no_runs_and_no_calls() -> None:
    # Act / Assert
    assert summarize([])["success_rate"] == 0.0
    assert summarize([_run(False, 0, 0, 10)])["validity_rate"] == 0.0


def test_render_results_lists_rows_and_failures() -> None:
    # Arrange
    runs = [_run(True, 1, 1, 50), _run(False, 3, 2, 70)]
    result = {
        "model": "m",
        "tier": "lite",
        "date": "2026-10-20",
        "endpoint": "http://x/v1",
        "runs": runs,
        "metrics": summarize(runs),
    }

    # Act
    text = render_results([result])
    empty = render_results([])

    # Assert
    assert "| m | lite | native | 2026-10-20 | 1/2 (50.0%) | 3/4 (75.0%) | 60 | 2.0 |" in text
    assert "x / y" in text
    assert "No live runs recorded yet" in empty


def test_parse_text_calls_recovers_json_calls_from_prose_and_tags() -> None:
    # Arrange
    content = (
        'Let me check.\n<tool_call>{"name": "query", "arguments": {"path": "a.csv", '
        '"sql": "SELECT 1"}}</tool_call>\n```json\n{"name": "overview", "parameters": '
        '{"path": "a.csv"}}\n```'
    )

    # Act
    calls = parse_text_calls(content)

    # Assert
    assert [c["function"]["name"] for c in calls] == ["query", "overview"]
    assert json.loads(calls[0]["function"]["arguments"]) == {"path": "a.csv", "sql": "SELECT 1"}


def test_parse_text_calls_ignores_plain_answers_and_broken_json() -> None:
    # Act / Assert
    assert parse_text_calls("ANSWER: 300") == []
    assert parse_text_calls('{"name": "query", "arguments": {"path": ') == []
    assert parse_text_calls('{"total": 5}') == []
    assert parse_text_calls("") == []
