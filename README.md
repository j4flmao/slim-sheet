# slim-sheet

<!-- mcp-name: io.github.j4flmao/slim-sheet -->

**A spreadsheet MCP server that small local models can actually use.** It reads
`xlsx xlsm xltx xltm xlsb xls ods numbers csv tsv parquet json` locally (no Office, no network),
keeps every response under a token budget, and gives agents SQL, cleaning, anomaly detection, a
formula copilot, Excel-correct editing and **one-call branded dashboards** — through **8, 18 or 30 tools**
depending on how capable the model is.

Works with Claude Desktop / Claude Code (Free or Pro), OpenCode, Codex, Cursor, Cline, Continue,
LM Studio, Ollama (via mcpo or Open WebUI) — and as a plain CLI for agents without MCP.

## Install

```bash
uvx slim-sheet mcp                          # MCP server over stdio
uvx slim-sheet overview sales.xlsx          # or use any tool from the shell
uvx --from "slim-sheet[all]" slim-sheet mcp # + offline formula engine, encrypted files, Apple Numbers
```

Host configuration for every client (Claude, Cursor, OpenCode, Codex, LM Studio, Ollama/mcpo, Open WebUI,
Cline, Continue) is in [docs/hosts.md](https://github.com/j4flmao/slim-sheet/blob/main/docs/hosts.md).

```json
{
  "mcpServers": {
    "slim-sheet": {
      "command": "uvx",
      "args": ["slim-sheet", "mcp"],
      "env": { "SLIM_SHEET_TIER": "lite", "SLIM_SHEET_ROOTS": "/path/to/your/data" }
    }
  }
}
```

## Tiers

| Tier | Tools | For |
|---|---|---|
| `full` (default) | every tool below | any model that should be able to do everything |
| `lite` | `overview`, `read`, `find`, `query`, `profile`, `write`, `explain_formula`, `build_report` | 3B–8B local models |
| `standard` | lite + `clean`, `anomalies` (+formula audit), `merge`, `convert` (+split), `diff`, `undo`, `evaluate` (+check), `describe_cell` (+trace), `style_sheet`, `what_if` | mid-size and frontier models |
| `full` | standard + `create_workbook`, `sheet_ops`, `range_ops`, `name_ops`, `format_range`, `set_table`, `rules`, `add_chart`, `sheet_layout`, `annotate`, `pivot`, `split`, `recalc` | frontier models building workbooks |

Pick exact tools with `SLIM_SHEET_TOOLS=overview,read,query`.

## Design: professional output with almost no description

The model names *what* to show; slim-sheet decides *how* it looks — layout, colours, fonts, number formats, chart
type, sort order and an insight title — from a brand. One call:

```text
build_report(path="sales.xlsx", data_sheet="Sales", title="Sales performance 2025",
             kpis=["Revenue=sum:Revenue", "Margin=ratio:Profit/Revenue", "Orders=count"],
             charts=["trend:Month:Revenue", "rank:Region:Revenue", "share:Product:Revenue"])
```

gives a one-page dashboard: KPI tiles with live formulas (7.1m · 20.4% · 144), a line titled "Revenue rose 58% from
2025-01 to 2025-12", a sorted bar with the leader highlighted, a labelled doughnut, a source footer.
`style_sheet(path, sheet, preset="report")` does the same for any table and reports the number formats it inferred.

**Your company's look, set once.** Built-in brands: `slim` (default), `office`, `modern` (Aptos), `mono` (IBCS-style),
`gov` (UK government statistics style). Or point `SLIM_SHEET_BRAND` at a JSON file — only `name` is required:

```json
{ "name": "Acme", "primary": "#C8102E", "body_font": "Arial", "currency": "$", "profile": "finance" }
```

Every workbook then carries the brand as a real Office theme, so Excel's own table and chart styles follow it too.
Rules behind the defaults (IBCS, finance colour coding, UK spreadsheet accessibility guidance, chart-choice tables):
[PLAN.md §16](https://github.com/j4flmao/slim-sheet/blob/main/PLAN.md).

## What makes it different

- **Reads everything** with one interface, including `.xls`, `.xlsb`, `.ods`, Apple Numbers and password-protected xlsx.
- **Budgeted output**: Markdown tables with A1 coordinates, ~1,500 tokens per response, `next` cursors to continue.
- **Errors that teach**: every error has a code, a message and a corrected example call.
- **SQL over sheets** with DuckDB; messy sheets are split into regions with detected headers (`"Report.2"`).
- **Never destroys data**: edits go to `<name>.slim.xlsx` unless `overwrite=true`, writes are atomic, every replaced file
  is journaled, and `undo` steps back.
- **Excel-correct structure edits**: inserting/deleting rows or renaming a sheet rewrites formulas, names, merges,
  validations, conditional formats, tables and charts the way Excel does.
- **Model checking**: `anomalies scope='formulas'` finds hardcoded numbers, broken patterns, double counting and
  where an error starts; `describe_cell trace=` follows formulas across sheets; `what_if` runs delta, goal seek
  and sensitivity grids without touching the file.
- **Formula copilot**: 514-function catalog, version checks (`XLOOKUP` in Excel 2016 → suggests `INDEX/MATCH`), offline
  evaluation with dependency chasing.
- **Macros, read-only**: `read sheet='#vba'` shows VBA source, `anomalies scope='macros'` rates risk, and agents redo
  a macro's job with tools (prompt `replace_macro`). slim-sheet never runs macros.
- **Excel-exact formulas**: newer functions are stored with Excel's `_xlfn.` prefixes and spilling formulas as real
  dynamic arrays, with a warning when something blocks the spill.
- **Flexible paths, safe by default**: give any folder in chat (no config change); system folders,
  AppData, dot-folders such as `.ssh` and non-spreadsheet files are refused. Read-only SQL with file access locked, no model-written code is executed, no network.

## Configuration

| Variable | Default | Meaning |
|---|---|---|
| `SLIM_SHEET_TIER` | `full` | `full`, `standard` or `lite` |
| `SLIM_SHEET_TOOLS` | — | explicit comma-separated tool list |
| `SLIM_SHEET_ROOTS` | current directory | default folders for relative paths and `overview()` (`;` on Windows, `:` elsewhere) |
| `SLIM_SHEET_STRICT_ROOTS` | off | `1` = only paths inside `SLIM_SHEET_ROOTS` (the old allow-list) |
| `SLIM_SHEET_MAX_TOKENS` | `1500` | default response budget |
| `SLIM_SHEET_TOKEN` | — | bearer token for `--transport streamable-http` (required off loopback) |
| `SLIM_SHEET_PASSWORD` | — | password for encrypted workbooks (never passed through chat) |
| `SLIM_SHEET_SOFFICE` | autodetect | LibreOffice binary for `recalc` |
| `SLIM_SHEET_AUDIT_LOG` | — | JSONL audit log of every tool call (contents redacted) |
| `SLIM_SHEET_RATE_LIMIT` | `120` | HTTP tool calls per minute per client (`0` = off) |
| `SLIM_SHEET_MAX_CELLS` | `3000000` | refuse workbooks declaring more cells (`0` = off) |
| `SLIM_SHEET_BRAND` | `slim` | brand name (`slim`, `office`, `modern`, `mono`, `gov`) or path to a brand `.json` |

Extras: `formula` (offline `evaluate` and `what_if`), `crypto` (encrypted xlsx), `numbers` (Apple Numbers), `images` (insert pictures with `annotate`), `all`.

## Prompt library

Ready-made, data-agnostic prompts — a shared CORE method, thirteen ROLES (nine analyst roles: sales,
finance, marketing, customer, operations, HR, reconciliation, data quality, general; plus model
auditor, workbook editor, macro migrator and file reviewer) and eight OUTPUT patterns (dashboard,
dashboard pack, report workbook, clean dataset, reconciliation report, quick answer, audit report,
edited workbook),
with filled-in example sets, about 50 one-paste task recipes in 8 packs (finance, sales, HR, operations,
data prep, Excel tasks, reporting, Vietnam) and compact blocks for small local models:
[docs/prompts](https://github.com/j4flmao/slim-sheet/blob/main/docs/prompts/README.md).

## Skill packs

[skills/building-client-reports](skills/building-client-reports/SKILL.md) turns the tools into
client-ready reports: a report brief, role presets (CFO, board, sales, marketing, operations,
HR, market research), a standards checklist and evals. Copy it into `~/.claude/skills/` or
attach it to any agent ([skills/README.md](skills/README.md)).

## Non-MCP hosts

```bash
slim-sheet schema --format openai --tier lite      # function-calling schemas for any OpenAI-compatible API
slim-sheet schema --format anthropic --tier full   # Messages API tool definitions
SLIM_SHEET_TOKEN=... slim-sheet mcp --transport streamable-http --host 0.0.0.0 --port 8765
```

## Development

```bash
uv sync --all-extras
uv run pre-commit install                 # ruff, mypy, manifests, commit-message rule, tests on push
uv run python scripts/make_demo.py        # sample workbooks in ~/Documents/slim-sheet-demo
uv run ruff check . && uv run ruff format --check . && uv run mypy src && uv run pytest -q
```

Issues and pull requests use the templates in `.github/`; security reports go through
[SECURITY.md](https://github.com/j4flmao/slim-sheet/blob/main/SECURITY.md).

Design and roadmap: [PLAN.md](https://github.com/j4flmao/slim-sheet/blob/main/PLAN.md) · agent rules:
[AGENTS.md](https://github.com/j4flmao/slim-sheet/blob/main/AGENTS.md) · MIT licensed.
