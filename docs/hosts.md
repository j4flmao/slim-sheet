# Connecting slim-sheet to your agent

slim-sheet runs locally. Every host below starts the same command, `uvx slim-sheet mcp`, which needs
[uv](https://docs.astral.sh/uv/) installed. Add extras with `uvx --from "slim-sheet[all]" slim-sheet mcp`.

Always set `SLIM_SHEET_ROOTS` to the folder that holds your spreadsheets. slim-sheet refuses paths outside it.
On Windows use `D:/data` or `D:\\data` in JSON, and separate several folders with `;`. On macOS/Linux use `:`.

## Pick a tier for your model

| Model | Tier | Why |
|---|---|---|
| 3B–8B local (Qwen 2.5 7B, Llama 3.1 8B, Gemma 3 4B…) | `lite` | 7 tools, about 900 schema tokens, small models stay accurate |
| 12B–32B local, Claude Haiku, GPT mini-class | `standard` | adds cleaning, merge/diff, formula checks |
| Claude Sonnet/Opus, GPT-class frontier models | `full` | adds structure edits, styling, charts, pivots |

## Claude Desktop (Free or Pro)

Option 1: install the `.mcpb` bundle (double-click). It asks for the data folder and tier.

Option 2: Settings → Developer → Edit Config, then add to `claude_desktop_config.json`:

```json
{
  "mcpServers": {
    "slim-sheet": {
      "command": "uvx",
      "args": ["slim-sheet", "mcp"],
      "env": { "SLIM_SHEET_TIER": "standard", "SLIM_SHEET_ROOTS": "C:/Users/me/Documents/data" }
    }
  }
}
```

## Claude Code

```bash
claude mcp add slim-sheet -e SLIM_SHEET_TIER=full -e SLIM_SHEET_ROOTS=/path/to/data -- uvx slim-sheet mcp
```

## Cursor

`~/.cursor/mcp.json` (global) or `.cursor/mcp.json` (project). The block is the same `mcpServers` JSON as Claude Desktop.

## OpenCode

`opencode.json`:

```json
{
  "mcp": {
    "slim-sheet": {
      "type": "local",
      "command": ["uvx", "slim-sheet", "mcp"],
      "environment": { "SLIM_SHEET_TIER": "standard", "SLIM_SHEET_ROOTS": "/path/to/data" },
      "enabled": true
    }
  }
}
```

## Codex CLI

`~/.codex/config.toml`:

```toml
[mcp_servers.slim-sheet]
command = "uvx"
args = ["slim-sheet", "mcp"]
env = { SLIM_SHEET_TIER = "standard", SLIM_SHEET_ROOTS = "/path/to/data" }
```

## VS Code (Copilot agent mode), Cline, Continue

- **VS Code**: `.vscode/mcp.json` → `{"servers": {"slim-sheet": {"type": "stdio", "command": "uvx", "args": ["slim-sheet", "mcp"], "env": {...}}}}`
- **Cline**: MCP Servers → Configure → `cline_mcp_settings.json`, using the same `mcpServers` block as Claude Desktop.
- **Continue**: `config.yaml` → `mcpServers: [{name: slim-sheet, command: uvx, args: [slim-sheet, mcp], env: {...}}]`

## LM Studio (local models)

Program tab → Install → Edit `mcp.json`, then add the same `mcpServers` block with `"SLIM_SHEET_TIER": "lite"`.
Load a model that supports tool use (Qwen 2.5 7B Instruct works well) and enable the slim-sheet tools in the chat.

## Ollama

The Ollama CLI is not an MCP client. Pick one of these:

1. **Open WebUI + mcpo** (OpenAPI bridge):
   ```bash
   uvx mcpo --port 8000 -- uvx slim-sheet mcp
   ```
   In Open WebUI: Settings → Tools → add `http://localhost:8000`. Set `SLIM_SHEET_TIER=lite` and
   `SLIM_SHEET_ROOTS` in the environment that runs mcpo.
2. **Open WebUI with native MCP (streamable HTTP)**:
   ```bash
   SLIM_SHEET_TOKEN=<long-random-string> SLIM_SHEET_TIER=lite uvx slim-sheet mcp --transport streamable-http --port 8765
   ```
   Add `http://localhost:8765/mcp` with bearer token `<long-random-string>`.
3. **A terminal MCP host** such as `ollmcp` or `mcphost`, given the same `mcpServers` JSON.

## Any OpenAI-compatible API, no MCP

Export the tool definitions and route the model's tool calls to the CLI:

```bash
uvx slim-sheet schema --format openai --tier lite > tools.json
uvx slim-sheet query sales.xlsx 'SELECT "Region", sum("Total") FROM "Orders" GROUP BY 1'
```

Every tool is a subcommand with the same arguments (`uvx slim-sheet <tool> --help`). Errors print the same
JSON (`{"error", "message", "hint"}`) and exit with status 1.

## Remote / shared server

```bash
SLIM_SHEET_TOKEN=<long-random-string> uvx slim-sheet mcp --transport streamable-http --host 0.0.0.0 --port 8765
```

Without `SLIM_SHEET_TOKEN`, slim-sheet only listens on loopback. Put it behind TLS (a reverse proxy) when you
expose it beyond your machine.
