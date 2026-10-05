from __future__ import annotations

import asyncio
import importlib
import json
import os
import sys
import tempfile
import time
from collections import deque
from datetime import datetime, timezone
from importlib import metadata
from pathlib import Path
from typing import Any, TextIO


MODEL_PROVIDER_ID = "officeval"
PROVIDER_KEY_ENV = "OFFICEVAL_API_KEY"
BASE_INSTRUCTIONS_PATH = Path(__file__).with_name("codex_base_instructions.md")
BASE_INSTRUCTIONS_HEADER_END = "-->\n\n"
# The codex build the vendored base instructions were taken from. Pinned in
# requirements-container.txt.
BASE_INSTRUCTIONS_CODEX_VERSION = "0.144.4"


class _MirroredStderr(deque):
    """Forward every codex stderr line to our own stderr.

    The SDK drains the codex process' stderr into a bounded deque that it only
    reads back when codex dies unexpectedly, so RUST_LOG output would otherwise
    never leave the container. Replacing that deque keeps the SDK's own tail
    lookups working while the runner captures the full log.
    """

    def append(self, line: str) -> None:
        print(line, file=sys.stderr, flush=True)
        super().append(line)


def mirror_codex_stderr(codex: Any) -> None:
    """Stream the codex binary's stderr out of the container."""
    client = codex._client._sync
    client._stderr_lines = _MirroredStderr(client._stderr_lines, maxlen=400)


def utc_now_iso() -> str:
    """Return the current UTC timestamp in millisecond ISO 8601 format."""
    return datetime.now(timezone.utc).isoformat(timespec="milliseconds").replace(
        "+00:00", "Z"
    )


def _emit_json(stream: TextIO, record: dict[str, Any]) -> None:
    stream.write(json.dumps(record, ensure_ascii=False, sort_keys=True) + "\n")
    stream.flush()


def load_base_instructions() -> str:
    """Return the prompt codex applies to models it has no metadata for.

    A model catalog entry has to carry its own base instructions, so the
    fallback prompt is vendored from the pinned codex build. Refusing to run
    against any other build keeps the benchmark prompt from silently changing
    when codex is upgraded.
    """
    installed = metadata.version("openai-codex-cli-bin")
    if installed != BASE_INSTRUCTIONS_CODEX_VERSION:
        raise RuntimeError(
            f"codex {installed} does not match the base instructions vendored "
            f"from codex {BASE_INSTRUCTIONS_CODEX_VERSION}; install the version "
            "pinned in requirements-container.txt"
        )
    # The file opens with an HTML comment recording where the prompt came
    # from; that header is not part of what codex sends.
    header, separator, prompt = BASE_INSTRUCTIONS_PATH.read_text(
        encoding="utf-8"
    ).partition(BASE_INSTRUCTIONS_HEADER_END)
    if not header.startswith("<!--") or not separator:
        raise RuntimeError(
            f"{BASE_INSTRUCTIONS_PATH} is missing its provenance header"
        )
    return prompt


def build_model_catalog(
    model: str,
    provider: dict[str, Any],
    input_modalities: list[str],
    base_instructions: str,
) -> dict[str, Any]:
    """Describe one model to codex instead of letting it guess.

    codex only knows its own built-in models; for anything else it falls back to
    metadata that claims image support, so a `view_image` call ships a base64
    image to a text-only provider. Every field below reproduces that fallback
    metadata, which keeps the request identical to an unlisted model except for
    `input_modalities`: declaring text only makes codex refuse `view_image`
    in-process. `supported_reasoning_levels` stays empty for the same reason —
    filling it in would start sending a `reasoning` field the fallback drops.
    """
    return {
        "models": [
            {
                "slug": model,
                "display_name": model,
                "description": "OfficeVal benchmark model.",
                "base_instructions": base_instructions,
                "input_modalities": list(input_modalities),
                "context_window": provider["context_window"],
                "max_context_window": provider["context_window"],
                "supported_reasoning_levels": [],
                "supports_reasoning_summaries": False,
                "default_reasoning_summary": "none",
                "support_verbosity": False,
                "supports_parallel_tool_calls": False,
                "supports_image_detail_original": False,
                "supports_search_tool": False,
                "web_search_tool_type": "text",
                "shell_type": "unified_exec",
                "multi_agent_version": "v1",
                "truncation_policy": {"mode": "tokens", "limit": 10000},
                "include_skills_usage_instructions": False,
                "use_responses_lite": False,
                "experimental_supported_tools": [],
                "additional_speed_tiers": [],
                "service_tiers": [],
                "availability_nux": None,
                "upgrade": None,
                "visibility": "list",
                "supported_in_api": True,
                "priority": 1,
            }
        ]
    }


def write_model_catalog(catalog: dict[str, Any]) -> Path:
    """Write the catalog where the codex process can read it at startup."""
    path = Path(tempfile.mkdtemp(prefix="officeval-catalog-")) / "catalog.json"
    path.write_text(json.dumps(catalog, ensure_ascii=False), encoding="utf-8")
    return path


def build_thread_config(provider: dict[str, Any]) -> dict[str, Any]:
    config: dict[str, Any] = {
        "model_providers": {
            MODEL_PROVIDER_ID: {
                "name": "OfficeVal Responses provider",
                "base_url": provider["base_url"],
                "env_key": PROVIDER_KEY_ENV,
                "wire_api": "responses",
                "request_max_retries": provider["request_max_retries"],
                "stream_max_retries": provider["stream_max_retries"],
                "stream_idle_timeout_ms": provider["stream_idle_timeout_ms"],
                "supports_standalone_web_search": False,
                "supports_websockets": False,
            }
        },
        # The benchmark measures the Agent without the web, for every model and run.
        "web_search": "disabled",
    }
    if provider["context_window"] is not None:
        config["model_context_window"] = provider["context_window"]
    return config


async def run_worker(
    request: dict[str, Any],
    *,
    stdout: TextIO = sys.stdout,
    workspace: Path = Path("/workspace"),
) -> int:
    started_at = utc_now_iso()
    started_clock = time.perf_counter()
    sdk = importlib.import_module("openai_codex")
    sdk_types = importlib.import_module("openai_codex.types")
    event_index = 0
    agent_message_steps = 0
    max_steps_hit = False
    completed_payload: Any | None = None
    usage: Any | None = None
    final_response: str | None = None
    provider = request["provider"]
    config_overrides: tuple[str, ...] = ()
    if request["input_modalities"]:
        catalog_path = write_model_catalog(
            build_model_catalog(
                request["model"],
                provider,
                request["input_modalities"],
                load_base_instructions(),
            )
        )
        # Only the codex process reads this key. The same key inside the thread
        # config is loaded but never reaches model metadata, so the modality
        # limit would be silently dropped there.
        config_overrides = (f"model_catalog_json={json.dumps(str(catalog_path))}",)
    codex_config = sdk.CodexConfig(
        cwd=str(workspace),
        env={PROVIDER_KEY_ENV: provider["api_key"]},
        config_overrides=config_overrides,
    )
    codex = sdk.AsyncCodex(codex_config)
    if os.environ.get("RUST_LOG"):
        mirror_codex_stderr(codex)
    async with codex:
        thread = await codex.thread_start(
            approval_mode=sdk.ApprovalMode.deny_all,
            config=build_thread_config(provider),
            cwd=str(workspace),
            developer_instructions="Work only in /workspace. Do not spawn subagents.",
            ephemeral=True,
            model=request["model"],
            model_provider=MODEL_PROVIDER_ID,
            sandbox=sdk.Sandbox.full_access,
        )
        turn_kwargs: dict[str, Any] = {
            "approval_mode": sdk.ApprovalMode.deny_all,
            "cwd": str(workspace),
            "sandbox": sdk.Sandbox.full_access,
        }
        if request["effort"] is not None:
            turn_kwargs["effort"] = sdk_types.ReasoningEffort(request["effort"])
        turn = await thread.turn(request["prompt"], **turn_kwargs)
        stream = turn.stream()
        try:
            async for event in stream:
                event_index += 1
                payload = event.payload
                method = str(event.method)
                _emit_json(
                    stdout,
                    {
                        "type": "event",
                        "event_index": event_index,
                        "recorded_at": utc_now_iso(),
                        "method": method,
                        "payload": payload.model_dump(
                            by_alias=True, exclude_none=True, mode="json"
                        ),
                    },
                )
                if method == "item/completed":
                    item = payload.item.root
                    if item.type == "agentMessage":
                        agent_message_steps += 1
                        final_response = item.text
                        if (
                            agent_message_steps >= request["max_steps"]
                            and not max_steps_hit
                        ):
                            max_steps_hit = True
                            await turn.interrupt()
                elif method == "thread/tokenUsage/updated":
                    usage = payload.token_usage
                elif method == "turn/completed":
                    completed_payload = payload
        finally:
            await stream.aclose()

    completed_turn = completed_payload.turn
    terminal_error = (
        completed_turn.error.model_dump(
            by_alias=True, exclude_none=True, mode="json"
        )
        if completed_turn.error is not None
        else None
    )
    _emit_json(
        stdout,
        {
            "type": "result",
            "agent": {
                "native_status": completed_turn.status.value,
                "thread_id": str(thread.id),
                "turn_id": str(turn.id),
                "error": terminal_error,
                "final_response": final_response,
                "agent_message_steps": agent_message_steps,
                "max_steps_hit": max_steps_hit,
                "usage": (
                    usage.total.model_dump(mode="json")
                    if usage is not None
                    else None
                ),
                "duration_ms": completed_turn.duration_ms,
                "started_at": started_at,
                "ended_at": utc_now_iso(),
                "wall_time_ms": round((time.perf_counter() - started_clock) * 1000),
            },
        },
    )
    return 0


def main() -> int:
    return asyncio.run(run_worker(json.loads(sys.stdin.readline())))


if __name__ == "__main__":
    raise SystemExit(main())
