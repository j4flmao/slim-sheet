from __future__ import annotations

import importlib.util
from pathlib import Path

import pytest

_SPEC = importlib.util.spec_from_file_location(
    "check_commit_msg", Path(__file__).resolve().parent.parent / "scripts" / "check_commit_msg.py"
)
assert _SPEC is not None and _SPEC.loader is not None
check_commit_msg = importlib.util.module_from_spec(_SPEC)
_SPEC.loader.exec_module(check_commit_msg)


@pytest.mark.parametrize(
    "message",
    [
        "feat(query): add offset paging",
        "fix: handle empty sheets\n\nBody explains why.",
        "refactor(core)!: rename cursor format",
        "Merge branch 'main' into feature",
        "fixup! feat: add x",
        "docs: update hosts guide\n# comment lines are ignored",
    ],
)
def test_valid_commit_messages(message: str) -> None:
    # Act / Assert
    assert check_commit_msg.check(message) == []


@pytest.mark.parametrize(
    ("message", "fragment"),
    [
        ("added stuff", "header must look like"),
        ("feature: add x", "header must look like"),
        ("feat: Add x", "lowercase"),
        ("feat: add x.", "period"),
        ("feat: add x\nbody without blank line", "blank line"),
        ("feat: " + "x" * 80, "max 72"),
        ("", "header must look like"),
    ],
)
def test_invalid_commit_messages(message: str, fragment: str) -> None:
    # Act
    problems = check_commit_msg.check(message)

    # Assert
    assert any(fragment in p for p in problems), problems
