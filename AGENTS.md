# AGENTS.md — slim-sheet

Canonical instructions for every coding agent working in this repo (Claude Code, OpenCode, Codex,
Cursor, Cline, Aider, local models via LM Studio/Ollama). `CLAUDE.md` and other agent files only point here.

## What this project is
A universal, token-frugal spreadsheet **MCP server + CLI** in Python. It must work for small local
models (3B–8B) as well as frontier models. Read [PLAN.md](PLAN.md) before any non-trivial change —
it holds the architecture, tool surface, tiers and roadmap. If you change a decision, update PLAN.md in the same change.

## Toolchain (uv only)
- Python is managed by **uv** (already installed). Never call `pip`, `python -m venv`, or a global `python`.
- Setup: `uv sync --all-extras && uv run pre-commit install` (hooks: ruff, mypy, manifests, Conventional Commit messages, pytest on push)
- Demo data: `uv run python scripts/make_demo.py` writes sample workbooks to `~/Documents/slim-sheet-demo`
- Run server: `uv run slim-sheet mcp` · CLI: `uv run slim-sheet overview <file>`
- Add a dependency: `uv add <pkg>` (dev: `uv add --dev <pkg>`; optional: `uv add --optional <extra> <pkg>`). Never hand-edit versions in `uv.lock`.

## Verification gate (run before saying "done")
```
uv run ruff check .
uv run ruff format --check .
uv run mypy src
uv run pytest -q
```
Report the actual output. Do not claim success on a skipped or failing step.

## Architecture rules
- `src/slim_sheet/core/` — pure logic, **no I/O**, no imports from adapters/services/interfaces.
- `src/slim_sheet/adapters/` — all file, DuckDB, calamine, openpyxl access. Heavy/optional libs are imported lazily inside functions.
- `src/slim_sheet/services/` — one use-case per module; returns `core.models` objects, never strings for transport.
- `src/slim_sheet/interfaces/` — MCP / CLI adapters only: parse args → call service → render. No business logic.
- Tool metadata (name, tier, description, params) lives **only** in `interfaces/toolspec.py`.
- After adding/renaming a tool or changing its description, regenerate the published manifests:
  `uv run python scripts/build_manifests.py` (CI runs it with `--check`).

## MCP tool design rules (small-model first)
- Tier caps are hard limits: `lite` ≤ 8 tools, `standard` ≤ 18, `full` ≤ 30. Adding a tool means justifying its tier in PLAN.md §5.
- Arguments are flat primitives (`str`, `int`, `bool`, `list[str]`, 2-D `list[list]` for values). No nested objects, no unions beyond `X | None`.
- Descriptions ≤ 25 words, start with a verb, include one tiny example when args are non-obvious.
- Every response respects the token budget (`core/budget.py`) and returns a `next` cursor when truncated.
- Every error is a `SlimError(code, message, hint)` and the hint contains a corrected call.
- **stdout is the MCP channel** — never `print()` in library code; log to stderr via `logging`.

## Safety rules
- Never execute model-supplied Python. SQL is read-only unless `SLIM_SHEET_SQL_WRITE=1`.
- All paths go through `adapters/sandbox.py`: `SLIM_SHEET_ROOTS` are default folders; outside them only
  data files in non-system, non-hidden folders are allowed (`SLIM_SHEET_STRICT_ROOTS=1` = roots only).
- Writes are non-destructive by default (`<stem>.slim<suffix>`), atomic (temp + `os.replace`), and journaled.
- No network access in any code path unless behind an explicit opt-in env flag.

## Code style
- Python ≥ 3.11, full type hints, `mypy --strict` clean. `from __future__ import annotations` in every module.
- Functions ≤ ~40 lines; split helpers instead of growing them.
- No `TODO`/placeholder bodies. Fail fast with contextual messages (file, sheet, range).
- Frozen `dataclass(slots=True)` for models; no pandas in `core/` or `services/` (DuckDB/Arrow only).

## Tests
- Every new service or core function ships with tests (Arrange-Act-Assert, edge cases: empty sheet, merged cells, huge cell, bad path).
- Fixtures are generated in `tests/conftest.py` / `tests/fixtures/build.py` into `tmp_path`; do not commit generated xlsx except formats with no Python writer (xls, xlsb, ods), and document their provenance.

## Commits
Conventional Commits (`feat:`, `fix:`, `docs:`, `test:`, `refactor:`, `chore:`), one logical change per commit.
Do not commit or push unless the user asks.
