"""Aggregate per-run metrics and render evals/RESULTS.md from evals/results/*.json."""

from __future__ import annotations

import json
from collections.abc import Mapping, Sequence
from pathlib import Path
from typing import Any

HOW_TO_RUN = """\
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
"""


def summarize(runs: Sequence[Mapping[str, Any]]) -> dict[str, Any]:
    total = len(runs)
    calls = sum(int(r["tool_calls"]) for r in runs)
    valid = sum(int(r["valid_calls"]) for r in runs)
    tokens = sum(int(r["tokens"]) for r in runs)
    successes = sum(1 for r in runs if r["success"])
    return {
        "tasks": total,
        "successes": successes,
        "success_rate": successes / total if total else 0.0,
        "tool_calls": calls,
        "valid_calls": valid,
        "validity_rate": valid / calls if calls else 0.0,
        "tokens": tokens,
        "tokens_estimated": any(r["tokens_estimated"] for r in runs),
        "avg_tokens": tokens / total if total else 0.0,
        "avg_calls": calls / total if total else 0.0,
    }


def _pct(x: float) -> str:
    return f"{100 * x:.1f}%"


def _summary_row(result: Mapping[str, Any]) -> str:
    m = result["metrics"]
    approx = "~" if m["tokens_estimated"] else ""
    cells = [
        result["model"],
        result["tier"],
        result.get("mode", "native"),
        result["date"],
        f"{m['successes']}/{m['tasks']} ({_pct(m['success_rate'])})",
        f"{m['valid_calls']}/{m['tool_calls']} ({_pct(m['validity_rate'])})",
        f"{approx}{m['avg_tokens']:.0f}",
        f"{m['avg_calls']:.1f}",
        result["endpoint"],
    ]
    return "| " + " | ".join(str(c) for c in cells) + " |"


def _failures(result: Mapping[str, Any]) -> list[str]:
    failed = [r for r in result["runs"] if not r["success"]]
    mode = result.get("mode", "native")
    lines = [f"### {result['model']} / {result['tier']} / {mode}: {len(failed)} failed task(s)"]
    for r in failed:
        answer = " ".join(str(r["answer"]).split())[:120].replace("|", "/")
        reason = r["detail"] or r["stop_reason"]
        lines.append(f"- `{r['task_id']}` ({r['tool_calls']} calls, {reason}): {answer}")
    return lines


def render_results(results: Sequence[Mapping[str, Any]]) -> str:
    lines = ["# Agent eval results", "", "Suite: `evals/tasks.json` (PLAN.md section 12).", ""]
    if not results:
        lines += ["No live runs recorded yet; see How to run.", ""]
    else:
        lines += [
            "| model | tier | calls | date | success | valid calls | tokens/task | calls/task "
            "| endpoint |",
            "|---|---|---|---|--:|--:|--:|--:|---|",
        ]
        lines += [_summary_row(r) for r in results]
        lines.append("")
        for result in results:
            lines += [*_failures(result), ""]
    return "\n".join(lines) + "\n" + HOW_TO_RUN


def load_results(results_dir: Path) -> list[dict[str, Any]]:
    if not results_dir.is_dir():
        return []
    files = sorted(results_dir.glob("*.json"))
    return [json.loads(f.read_text(encoding="utf-8")) for f in files]


def write_results_md(results_dir: Path, target: Path) -> None:
    target.write_text(render_results(load_results(results_dir)), encoding="utf-8")
