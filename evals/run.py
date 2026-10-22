"""Agent eval runner: slim-sheet MCP server (stdio) x any OpenAI-compatible chat endpoint.

uv run --group eval python -m evals.run --dry-run
uv run --group eval python -m evals.run --base-url http://localhost:11434/v1 --tier lite
"""

from __future__ import annotations

import argparse
import asyncio
import datetime as dt
import json
import logging
import os
import re
import shutil
import sys
import tempfile
from collections.abc import AsyncIterator, Sequence
from contextlib import asynccontextmanager
from pathlib import Path
from typing import Any

from mcp.client.session import ClientSession
from mcp.client.stdio import StdioServerParameters, stdio_client

from evals.agent import AgentContext, TaskRun, result_text, run_task
from evals.build_fixtures import build_all
from evals.llm import LM_STUDIO_URL, ChatClient, list_models
from evals.report import summarize, write_results_md
from evals.scoring import Task, check_answer, load_tasks, openai_tools, written_mismatches

log = logging.getLogger("evals.run")

EVALS_DIR = Path(__file__).resolve().parent
REPO_ROOT = EVALS_DIR.parent
TASKS_FILE = EVALS_DIR / "tasks.json"
RESULTS_DIR = EVALS_DIR / "results"
RESULTS_MD = EVALS_DIR / "RESULTS.md"


def _parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(prog="evals.run", description=__doc__)
    p.add_argument("--dry-run", action="store_true", help="replay reference calls, no LLM")
    p.add_argument("--base-url", default=os.environ.get("EVAL_BASE_URL", LM_STUDIO_URL))
    p.add_argument("--model", help="model id; default = first model the endpoint lists")
    p.add_argument("--api-key-env", default="EVAL_API_KEY", help="env var holding a bearer key")
    p.add_argument("--tier", default="lite", choices=["lite", "standard", "full"])
    p.add_argument("--task", action="append", default=[], help="run only these task ids")
    p.add_argument("--limit", type=int, help="run only the first N tasks")
    p.add_argument("--max-calls", type=int, default=12)
    p.add_argument("--temperature", type=float, default=0.0)
    p.add_argument("--timeout", type=float, default=300.0, help="seconds per chat request")
    p.add_argument(
        "--text-tool-calls",
        action="store_true",
        help="also accept tool calls the model writes as JSON text (recorded as a separate run)",
    )
    p.add_argument("--workdir", type=Path, help="fixture folder (default: a temp dir)")
    return p


def select_tasks(tasks: list[Task], ids: Sequence[str], limit: int | None) -> list[Task]:
    if ids:
        unknown = sorted(set(ids) - {t.id for t in tasks})
        if unknown:
            raise ValueError(f"Unknown task id(s): {', '.join(unknown)}")
        tasks = [t for t in tasks if t.id in ids]
    return tasks[:limit] if limit else tasks


def reset_workdir(workdir: Path) -> None:
    """Fresh fixtures for every task so writes from one task never leak into the next."""
    shutil.rmtree(workdir / ".slim-sheet", ignore_errors=True)
    build_all(workdir)


@asynccontextmanager
async def mcp_session(workdir: Path, tier: str) -> AsyncIterator[tuple[ClientSession, str]]:
    env = {**os.environ, "SLIM_SHEET_ROOTS": str(workdir), "SLIM_SHEET_TIER": tier}
    params = StdioServerParameters(
        command="uv",
        args=["run", "--project", str(REPO_ROOT), "slim-sheet", "mcp", "--tier", tier],
        env=env,
        cwd=str(workdir),
    )
    async with stdio_client(params) as (read, write), ClientSession(read, write) as session:
        init = await session.initialize()
        yield session, init.instructions or ""


# ---------------------------------------------------------------- dry run


async def _dry_check(session: ClientSession, task: Task, workdir: Path, tools: set[str]) -> str:
    """Return '' when the task's reference solution reproduces its expected answer."""
    outputs: list[str] = []
    for call in task.reference:
        if call.tool not in tools:
            return f"reference tool '{call.tool}' is not served in this tier"
        if call.tool not in task.tools:
            return f"reference tool '{call.tool}' missing from expected tools {task.tools}"
        text, is_error = result_text(await session.call_tool(call.tool, dict(call.args)))
        if is_error:
            return f"{call.tool} returned an error: {text[:200]}"
        outputs.append(text)
    if task.check.kind == "written":
        return "; ".join(written_mismatches(task.check, workdir))
    joined = "\n".join(outputs)
    return "" if check_answer(task.check, joined, strict=False) else f"answer not in: {joined!r}"


def servable(tasks: list[Task], tools: set[str]) -> tuple[list[Task], list[str]]:
    """Tasks the tier can serve, and the ids skipped (keeps scores fair per tier)."""
    kept = [t for t in tasks if all(call.tool in tools for call in t.reference)]
    return kept, [t.id for t in tasks if t not in kept]


async def dry_run(tasks: list[Task], workdir: Path, tier: str) -> int:
    failures = 0
    async with mcp_session(workdir, tier) as (session, _):
        tools = {t.name for t in (await session.list_tools()).tools}
        tasks, skipped = servable(tasks, tools)
        if skipped:
            log.info("skipped (tools not in tier %s): %s", tier, ", ".join(skipped))
        for task in tasks:
            reset_workdir(workdir)
            problem = await _dry_check(session, task, workdir, tools)
            failures += bool(problem)
            log.info("%-22s %s", task.id, f"FAIL {problem}" if problem else "ok")
    log.info("dry run: %d/%d tasks verified", len(tasks) - failures, len(tasks))
    return 1 if failures else 0


# ---------------------------------------------------------------- live run


def _slug(text: str) -> str:
    return re.sub(r"[^A-Za-z0-9.-]+", "-", text).strip("-").lower()


def save_result(model: str, tier: str, endpoint: str, mode: str, runs: list[TaskRun]) -> Path:
    payload: dict[str, Any] = {
        "model": model,
        "tier": tier,
        "mode": mode,
        "endpoint": endpoint,
        "date": dt.date.today().isoformat(),
        "runs": [r.to_json() for r in runs],
    }
    payload["metrics"] = summarize(payload["runs"])
    RESULTS_DIR.mkdir(exist_ok=True)
    suffix = "" if mode == "native" else f"__{mode}"
    target = RESULTS_DIR / f"{_slug(model)}__{tier}{suffix}.json"
    target.write_text(json.dumps(payload, indent=2) + "\n", encoding="utf-8")
    write_results_md(RESULTS_DIR, RESULTS_MD)
    return target


async def live_run(tasks: list[Task], workdir: Path, args: argparse.Namespace) -> int:
    api_key = os.environ.get(args.api_key_env, "")
    model = args.model or list_models(args.base_url, api_key)[0]
    client = ChatClient(args.base_url, model, api_key, args.timeout, args.temperature)
    runs: list[TaskRun] = []
    async with mcp_session(workdir, args.tier) as (session, instructions):
        listed = (await session.list_tools()).tools
        tasks, skipped = servable(tasks, {t.name for t in listed})
        if skipped:
            log.info("skipped (tools not in tier %s): %s", args.tier, ", ".join(skipped))
        schemas = {t.name: t.input_schema for t in listed}
        payload = openai_tools((t.name, t.description or "", t.input_schema) for t in listed)
        ctx = AgentContext(
            session,
            client,
            schemas,
            payload,
            instructions,
            workdir,
            args.max_calls,
            args.text_tool_calls,
        )
        for task in tasks:
            reset_workdir(workdir)
            run = await run_task(ctx, task)
            runs.append(run)
            mark = "PASS" if run.success else "FAIL"
            log.info("%-22s %s calls=%d tokens=%d", task.id, mark, run.tool_calls, run.tokens)
    mode = "text-calls" if args.text_tool_calls else "native"
    target = save_result(model, args.tier, args.base_url, mode, runs)
    metrics = summarize([r.to_json() for r in runs])
    log.info(
        "%s / %s: success %.1f%%, valid calls %.1f%% -> %s",
        model,
        args.tier,
        100 * metrics["success_rate"],
        100 * metrics["validity_rate"],
        target,
    )
    return 0


def main(argv: list[str] | None = None) -> int:
    args = _parser().parse_args(argv)
    logging.basicConfig(level=logging.INFO, stream=sys.stderr, format="%(message)s")
    logging.getLogger("evals.fixtures").setLevel(logging.WARNING)
    tasks = select_tasks(load_tasks(TASKS_FILE), args.task, args.limit)
    with tempfile.TemporaryDirectory(prefix="slim-eval-") as tmp:
        workdir = (args.workdir or Path(tmp)).resolve()
        workdir.mkdir(parents=True, exist_ok=True)
        if args.dry_run:
            return asyncio.run(dry_run(tasks, workdir, args.tier))
        return asyncio.run(live_run(tasks, workdir, args))


if __name__ == "__main__":
    raise SystemExit(main())
