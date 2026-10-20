"""Minimal OpenAI-compatible chat client (LM Studio, Ollama, vLLM, OpenAI) over stdlib urllib."""

from __future__ import annotations

import json
import logging
import urllib.error
import urllib.request
from dataclasses import dataclass
from typing import Any

log = logging.getLogger("evals.llm")

LM_STUDIO_URL = "http://localhost:1234/v1"
OLLAMA_URL = "http://localhost:11434/v1"


def _request(url: str, api_key: str, timeout: float, body: Any = None) -> Any:
    data = None if body is None else json.dumps(body).encode("utf-8")
    headers = {"Content-Type": "application/json"}
    if api_key:
        headers["Authorization"] = f"Bearer {api_key}"
    req = urllib.request.Request(url, data=data, headers=headers)
    try:
        with urllib.request.urlopen(req, timeout=timeout) as resp:
            return json.loads(resp.read().decode("utf-8"))
    except urllib.error.HTTPError as exc:
        detail = exc.read().decode("utf-8", "replace")[:500]
        raise RuntimeError(f"{url} returned HTTP {exc.code}: {detail}") from exc
    except (urllib.error.URLError, TimeoutError) as exc:
        raise RuntimeError(f"Cannot reach {url}: {exc}") from exc


def list_models(base_url: str, api_key: str = "", timeout: float = 5.0) -> list[str]:
    payload = _request(f"{base_url.rstrip('/')}/models", api_key, timeout)
    return [str(m["id"]) for m in payload.get("data", [])]


def endpoint_reachable(base_url: str, api_key: str = "") -> bool:
    try:
        list_models(base_url, api_key, timeout=3.0)
    except RuntimeError as exc:
        log.info("Endpoint %s not reachable: %s", base_url, exc)
        return False
    return True


def estimate_tokens(payload: Any) -> int:
    """Rough fallback (4 chars per token) when the server reports no usage."""
    return len(json.dumps(payload, ensure_ascii=False)) // 4


@dataclass(frozen=True, slots=True)
class ChatClient:
    base_url: str
    model: str
    api_key: str = ""
    timeout: float = 300.0
    temperature: float = 0.0

    def complete(
        self, messages: list[dict[str, Any]], tools: list[dict[str, Any]]
    ) -> tuple[dict[str, Any], int, bool]:
        """Return (assistant message, tokens used, whether tokens were estimated)."""
        body = {
            "model": self.model,
            "messages": messages,
            "tools": tools,
            "tool_choice": "auto",
            "temperature": self.temperature,
        }
        url = f"{self.base_url.rstrip('/')}/chat/completions"
        payload = _request(url, self.api_key, self.timeout, body)
        choices = payload.get("choices") or []
        if not choices:
            raise RuntimeError(f"{url} returned no choices for model {self.model}: {payload!r}")
        message: dict[str, Any] = choices[0].get("message") or {}
        usage = payload.get("usage") or {}
        used = int(usage.get("prompt_tokens", 0)) + int(usage.get("completion_tokens", 0))
        if used:
            return message, used, False
        return message, estimate_tokens(body) + estimate_tokens(message), True
