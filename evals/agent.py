"""One task = one function-calling conversation between a chat model and the MCP server."""

from __future__ import annotations

import asyncio
import json
import logging
from collections.abc import Mapping
from dataclasses import asdict, dataclass, field
from pathlib import Path
from typing import Any

from mcp.client.session import ClientSession

from evals.llm import ChatClient
from evals.scoring import (
    Task,
    check_answer,
    parse_text_calls,
    validate_call,
    written_mismatches,
)

log = logging.getLogger("evals.agent")

SYSTEM_PROMPT = (
    "You are a spreadsheet assistant with tools. Files live in the current folder; pass paths "
    "exactly like 'sales.xlsx'. Always use the tools to look at the data; never guess. "
    "When you are done, reply in plain text and end with one line: ANSWER: <short answer>."
)


@dataclass(slots=True)
class TaskRun:
    task_id: str
    success: bool = False
    answer: str = ""
    tool_calls: int = 0
    valid_calls: int = 0
    tool_errors: int = 0
    tokens: int = 0
    tokens_estimated: bool = False
    used_tools: list[str] = field(default_factory=list)
    text_calls: int = 0
    stop_reason: str = "answered"
    detail: str = ""

    def to_json(self) -> dict[str, Any]:
        return asdict(self)


@dataclass(frozen=True, slots=True)
class AgentContext:
    session: ClientSession
    client: ChatClient
    schemas: Mapping[str, Mapping[str, Any]]
    tools_payload: list[dict[str, Any]]
    instructions: str
    workdir: Path
    max_calls: int = 12
    text_tool_calls: bool = False  # recover calls written as JSON in the message text


def result_text(result: Any) -> tuple[str, bool]:
    """Flatten an MCP CallToolResult into (text, is_error)."""
    parts = [getattr(c, "text", "") for c in getattr(result, "content", [])]
    return "\n".join(p for p in parts if p), bool(getattr(result, "is_error", False))


async def _execute(ctx: AgentContext, name: str, raw_args: Any, run: TaskRun) -> str:
    verdict = validate_call(name, raw_args, ctx.schemas)
    run.tool_calls += 1
    run.used_tools.append(name)
    run.valid_calls += int(verdict.valid)
    if name not in ctx.schemas:
        run.tool_errors += 1
        known = ", ".join(ctx.schemas)
        return json.dumps({"error": "unknown_tool", "message": verdict.reason, "hint": known})
    if not verdict.valid and not verdict.args:
        run.tool_errors += 1
        return json.dumps({"error": "bad_argument", "message": verdict.reason})
    try:
        text, is_error = result_text(await ctx.session.call_tool(name, verdict.args))
    except Exception as exc:  # protocol failure must not abort the whole eval run
        log.warning("Tool %s failed for task %s with args %r: %s", name, run.task_id, raw_args, exc)
        text, is_error = json.dumps({"error": "call_failed", "message": str(exc)}), True
    run.tool_errors += int(is_error)
    return text


def _assistant_turn(message: Mapping[str, Any], calls: list[dict[str, Any]]) -> dict[str, Any]:
    return {"role": "assistant", "content": message.get("content") or "", "tool_calls": calls}


def _normalize_calls(raw_calls: list[dict[str, Any]], turn: int) -> list[dict[str, Any]]:
    calls = []
    for i, call in enumerate(raw_calls):
        fn = call.get("function") or {}
        args = fn.get("arguments", "{}")
        calls.append(
            {
                "id": call.get("id") or f"call_{turn}_{i}",
                "type": "function",
                "function": {
                    "name": str(fn.get("name", "")),
                    "arguments": args if isinstance(args, str) else json.dumps(args),
                },
            }
        )
    return calls


async def _converse(ctx: AgentContext, task: Task, run: TaskRun) -> None:
    system = f"{SYSTEM_PROMPT}\n\nServer notes: {ctx.instructions}".strip()
    messages: list[dict[str, Any]] = [
        {"role": "system", "content": system},
        {"role": "user", "content": task.prompt},
    ]
    for turn in range(ctx.max_calls + 1):
        message, used, estimated = await asyncio.to_thread(
            ctx.client.complete, messages, ctx.tools_payload
        )
        run.tokens += used
        run.tokens_estimated |= estimated
        raw_calls = message.get("tool_calls") or []
        if not raw_calls and ctx.text_tool_calls:
            raw_calls = parse_text_calls(str(message.get("content") or ""))
            run.text_calls += len(raw_calls)
        calls = _normalize_calls(raw_calls, turn)
        if not calls:
            run.answer = str(message.get("content") or "")
            return
        if run.tool_calls + len(calls) > ctx.max_calls:
            run.stop_reason = "max_tool_calls"
            return
        messages.append(_assistant_turn(message, calls))
        for call in calls:
            fn = call["function"]
            text = await _execute(ctx, fn["name"], fn["arguments"], run)
            messages.append({"role": "tool", "tool_call_id": call["id"], "content": text})
    run.stop_reason = "max_turns"


async def run_task(ctx: AgentContext, task: Task) -> TaskRun:
    run = TaskRun(task.id)
    try:
        await _converse(ctx, task, run)
    except RuntimeError as exc:
        log.error("Chat endpoint failed on task %s: %s", task.id, exc)
        run.stop_reason, run.detail = "llm_error", str(exc)[:300]
        return run
    if run.stop_reason != "answered":
        run.detail = f"stopped: {run.stop_reason}"
        return run
    if task.check.kind == "written":
        problems = written_mismatches(task.check, ctx.workdir)
        run.success, run.detail = not problems, "; ".join(problems)
    else:
        run.success = check_answer(task.check, run.answer)
    return run
