# Agent eval results

Suite: `evals/tasks.json` (PLAN.md section 12).

| model | tier | calls | date | success | valid calls | tokens/task | calls/task | endpoint |
|---|---|---|---|--:|--:|--:|--:|---|
| qwen2.5-coder:14b | lite | text-calls | 2026-10-20 | 2/3 (66.7%) | 5/5 (100.0%) | 2004 | 1.7 | http://localhost:11434/v1 |

### qwen2.5-coder:14b / lite / text-calls: 1 failed task(s)
- `write-grand-total` (3 calls, Orders!H1: got 'total_revenue', want 14962.5): ANSWER: Done.

## How to run

```bash
uv sync --group eval
uv run --group eval python -m evals.run --dry-run
uv run --group eval python -m evals.run --model qwen2.5-7b-instruct            # LM Studio
uv run --group eval python -m evals.run --base-url http://localhost:11434/v1 --model qwen2.5:7b
```

`--dry-run` builds the fixtures, replays each task's reference tool calls through the real MCP
server over stdio and checks the expected answers, with no LLM. A live run spawns
`slim-sheet mcp --tier <tier>`, drives the chat endpoint with function calling (max 12 tool
calls per task, temperature 0) and stores `evals/results/<model>__<tier>.json`; this file is
regenerated from all stored results.

Metrics: **success** = final answer (text after `ANSWER:`) or written `.slim.xlsx` cells match
the task check; **validity** = tool calls with a known name and JSON-schema-valid arguments;
**tokens** = prompt + completion tokens reported by the endpoint (estimated at 4 chars/token
when it reports none, marked `~`).

**calls** = `native` counts only OpenAI `tool_calls`; `text-calls` (`--text-tool-calls`) also
recovers calls the model wrote as JSON in its message text, as several local hosts do. Models
whose chat template lacks native tool support (e.g. qwen2.5-coder on Ollama) score 0 natively.
