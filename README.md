# slim-sheet

<!-- mcp-name: io.github.j4flmao/slim-sheet -->

**A spreadsheet MCP server that small local models can actually use.** It reads
`xlsx xlsm xltx xltm xlsb xls ods numbers csv tsv parquet json` locally (no Office, no network),
keeps every response under a token budget, and gives agents SQL, cleaning, anomaly detection, a
formula copilot and Excel-correct editing — through **7, 16 or 28 tools** depending on how capable
the model is.

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
| `lite` | `overview`, `read`, `find`, `query`, `profile`, `write`, `explain_formula` | 3B–8B local models |
| `standard` (default) | lite + `clean`, `anomalies`, `merge`, `convert`, `diff`, `undo`, `validate_formula`, `evaluate`, `describe_cell` | mid-size and frontier models |
| `full` | standard + `create_workbook`, `sheet_ops`, `rows_ops`, `name_ops`, `format_range`, `set_table`, `add_validation`, `conditional_format`, `add_chart`, `pivot`, `split`, `recalc` | frontier models building workbooks |

Pick exact tools with `SLIM_SHEET_TOOLS=overview,read,query`.

## What makes it different

- **Reads everything** with one interface, including `.xls`, `.xlsb`, `.ods`, Apple Numbers and password-protected xlsx.
- **Budgeted output**: Markdown tables with A1 coordinates, ~1,500 tokens per response, `next` cursors to continue.
- **Errors that teach**: every error has a code, a message and a corrected example call.
- **SQL over sheets** with DuckDB; messy sheets are split into regions with detected headers (`"Report.2"`).
- **Never destroys data**: edits go to `<name>.slim.xlsx` unless `overwrite=true`, writes are atomic, every replaced file
  is journaled, and `undo` steps back.
- **Excel-correct structure edits**: inserting/deleting rows or renaming a sheet rewrites formulas, names, merges,
  validations, conditional formats, tables and charts the way Excel does.
- **Formula copilot**: 514-function catalog, version checks (`XLOOKUP` in Excel 2016 → suggests `INDEX/MATCH`), offline
  evaluation with dependency chasing.
- **Safe by default**: path allow-list, read-only SQL with file access locked, no model-written code is executed, no network.

## Configuration

| Variable | Default | Meaning |
|---|---|---|
| `SLIM_SHEET_TIER` | `standard` | `lite`, `standard` or `full` |
| `SLIM_SHEET_TOOLS` | — | explicit comma-separated tool list |
| `SLIM_SHEET_ROOTS` | current directory | allowed folders (`;` on Windows, `:` elsewhere) |
| `SLIM_SHEET_MAX_TOKENS` | `1500` | default response budget |
| `SLIM_SHEET_TOKEN` | — | bearer token for `--transport streamable-http` (required off loopback) |
| `SLIM_SHEET_PASSWORD` | — | password for encrypted workbooks (never passed through chat) |
| `SLIM_SHEET_SOFFICE` | autodetect | LibreOffice binary for `recalc` |

Extras: `formula` (offline `evaluate`), `crypto` (encrypted xlsx), `numbers` (Apple Numbers), `all`.

## Non-MCP hosts

```bash
slim-sheet schema --format openai --tier lite      # function-calling schemas for any OpenAI-compatible API
slim-sheet schema --format anthropic --tier full   # Messages API tool definitions
SLIM_SHEET_TOKEN=... slim-sheet mcp --transport streamable-http --host 0.0.0.0 --port 8765
```

## Development

```bash
uv sync --all-extras
uv run ruff check . && uv run ruff format --check . && uv run mypy src && uv run pytest -q
```

Design and roadmap: [PLAN.md](https://github.com/j4flmao/slim-sheet/blob/main/PLAN.md) · agent rules:
[AGENTS.md](https://github.com/j4flmao/slim-sheet/blob/main/AGENTS.md) · MIT licensed.
