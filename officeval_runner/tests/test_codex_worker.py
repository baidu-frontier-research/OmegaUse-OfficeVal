from __future__ import annotations

import asyncio
import contextlib
import io
import json
import shutil
import tempfile
import unittest
from collections import deque
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import patch

from src.runtime.worker import (
    BASE_INSTRUCTIONS_CODEX_VERSION,
    build_model_catalog,
    build_thread_config,
    load_base_instructions,
    mirror_codex_stderr,
    run_worker,
)


def _request(
    max_steps: int = 2,
    input_modalities: list[str] | None = None,
) -> dict[str, object]:
    return {
        "model": "model-a",
        "prompt": "edit the file",
        "max_steps": max_steps,
        "effort": "high",
        "input_modalities": input_modalities,
        "provider": {
            "base_url": "https://gateway.example/v1",
            "api_key": "real-secret-key",
            "context_window": 200000,
            "request_max_retries": 4,
            "stream_max_retries": 5,
            "stream_idle_timeout_ms": 300000,
        },
    }


class _FakeStream:
    def __init__(self, events: list[object]):
        self.events = iter(events)
        self.closed = False

    def __aiter__(self):
        return self

    async def __anext__(self):
        try:
            return next(self.events)
        except StopIteration as exc:
            raise StopAsyncIteration from exc

    async def aclose(self) -> None:
        self.closed = True


class _FakeTurn:
    def __init__(self, events: list[object]):
        self.id = "turn-123"
        self.stream_instance = _FakeStream(events)
        self.interrupt_count = 0

    def stream(self) -> _FakeStream:
        return self.stream_instance

    async def interrupt(self) -> None:
        self.interrupt_count += 1


class _FakeModel(SimpleNamespace):
    def __init__(self, dumped: dict[str, object], **attributes: object):
        super().__init__(**attributes)
        self.dumped = dumped

    def model_dump(self, **_: object) -> dict[str, object]:
        return self.dumped


class _FakeThread:
    def __init__(self, turn: _FakeTurn, capture: dict[str, object]):
        self.id = "thread-123"
        self.turn_instance = turn
        self.capture = capture

    async def turn(self, prompt: str, **kwargs: object) -> _FakeTurn:
        self.capture["prompt"] = prompt
        self.capture["turn_kwargs"] = kwargs
        return self.turn_instance


class _FakeAsyncCodex:
    def __init__(self, config: object, thread: _FakeThread, capture: dict[str, object]):
        self.thread = thread
        self.capture = capture
        self.capture["codex_config"] = config

    async def __aenter__(self):
        return self

    async def __aexit__(self, *_: object) -> None:
        return None

    async def thread_start(self, **kwargs: object) -> _FakeThread:
        self.capture["thread_start"] = kwargs
        return self.thread


class CodexWorkerTests(unittest.TestCase):
    def test_thread_config_uses_responses_provider(self) -> None:
        provider = _request()["provider"]
        config = build_thread_config(provider)  # type: ignore[arg-type]
        self.assertEqual(
            config["model_providers"]["officeval"]["wire_api"],
            "responses",
        )
        self.assertEqual(
            config["model_providers"]["officeval"]["env_key"],
            "OFFICEVAL_API_KEY",
        )
        self.assertNotIn("api_key", json.dumps(config))
        self.assertEqual(config["model_context_window"], 200000)
        self.assertEqual(config["web_search"], "disabled")

    def test_counts_only_completed_agent_messages_and_interrupts_once(self) -> None:
        def event(method: str, payload: object) -> object:
            return SimpleNamespace(method=method, payload=payload)

        def item(item_type: str, text: str = "", phase: str | None = None) -> object:
            value = SimpleNamespace(
                type=item_type,
                text=text,
                phase=SimpleNamespace(value=phase) if phase is not None else None,
            )
            return _FakeModel(
                {"item": {"type": item_type, "text": text, "phase": phase}},
                item=SimpleNamespace(root=value),
            )

        completed_turn = _FakeModel(
            {"status": "completed", "durationMs": 321},
            status=SimpleNamespace(value="completed"),
            error=None,
            duration_ms=321,
        )
        events = [
            event("item/completed", item("commandExecution")),
            event("item/completed", item("agentMessage", "working", "commentary")),
            event("item/completed", item("agentMessage", "finished", "final_answer")),
            event(
                "thread/tokenUsage/updated",
                _FakeModel(
                    {"tokenUsage": {"total": {"totalTokens": 140}}},
                    token_usage=SimpleNamespace(
                        total=_FakeModel(
                            {
                                "input_tokens": 100,
                                "cached_input_tokens": 20,
                                "output_tokens": 40,
                                "reasoning_output_tokens": 10,
                                "total_tokens": 140,
                            }
                        )
                    ),
                ),
            ),
            event(
                "turn/completed",
                _FakeModel({"turn": completed_turn.dumped}, turn=completed_turn),
            ),
        ]
        capture: dict[str, object] = {}
        turn = _FakeTurn(events)
        thread = _FakeThread(turn, capture)

        class CodexConfig:
            def __init__(self, **kwargs: object):
                self.kwargs = kwargs

        sdk = SimpleNamespace(
            CodexConfig=CodexConfig,
            AsyncCodex=lambda config: _FakeAsyncCodex(config, thread, capture),
            ApprovalMode=SimpleNamespace(deny_all="deny_all"),
            Sandbox=SimpleNamespace(full_access="full_access"),
        )
        sdk_types = SimpleNamespace(ReasoningEffort=lambda value: value)
        output = io.StringIO()

        def import_module(name: str):
            return sdk if name == "openai_codex" else sdk_types

        with (
            tempfile.TemporaryDirectory() as directory,
            patch(
                "src.runtime.worker.importlib.import_module",
                side_effect=import_module,
            ),
        ):
            exit_code = asyncio.run(
                run_worker(
                    _request(),
                    stdout=output,
                    workspace=Path(directory),
                )
            )

        records = [json.loads(line) for line in output.getvalue().splitlines()]
        terminal = records[-1]["agent"]
        self.assertEqual(exit_code, 0)
        self.assertEqual(records[0]["type"], "event")
        self.assertEqual(records[-1]["type"], "result")
        self.assertEqual(terminal["thread_id"], "thread-123")
        self.assertEqual(terminal["turn_id"], "turn-123")
        self.assertEqual(terminal["agent_message_steps"], 2)
        self.assertTrue(terminal["max_steps_hit"])
        self.assertEqual(terminal["final_response"], "finished")
        self.assertEqual(terminal["usage"]["reasoning_output_tokens"], 10)
        self.assertEqual(turn.interrupt_count, 1)
        self.assertTrue(turn.stream_instance.closed)
        self.assertEqual(capture["thread_start"]["ephemeral"], True)
        self.assertEqual(capture["thread_start"]["sandbox"], "full_access")
        self.assertEqual(capture["thread_start"]["approval_mode"], "deny_all")
        self.assertEqual(capture["codex_config"].kwargs["config_overrides"], ())
        self.assertEqual(
            capture["codex_config"].kwargs["env"],
            {"OFFICEVAL_API_KEY": "real-secret-key"},
        )

    def test_text_only_model_gets_a_catalog_the_codex_process_reads(self) -> None:
        capture: dict[str, object] = {}
        completed_turn = _FakeModel(
            {"status": "completed"},
            status=SimpleNamespace(value="completed"),
            error=None,
            duration_ms=1,
        )
        events = [
            SimpleNamespace(
                method="turn/completed",
                payload=_FakeModel(
                    {"turn": completed_turn.dumped}, turn=completed_turn
                ),
            )
        ]
        thread = _FakeThread(_FakeTurn(events), capture)

        class CodexConfig:
            def __init__(self, **kwargs: object):
                self.kwargs = kwargs

        sdk = SimpleNamespace(
            CodexConfig=CodexConfig,
            AsyncCodex=lambda config: _FakeAsyncCodex(config, thread, capture),
            ApprovalMode=SimpleNamespace(deny_all="deny_all"),
            Sandbox=SimpleNamespace(full_access="full_access"),
        )
        sdk_types = SimpleNamespace(ReasoningEffort=lambda value: value)

        def import_module(name: str):
            return sdk if name == "openai_codex" else sdk_types

        with (
            tempfile.TemporaryDirectory() as directory,
            # Patched before import_module, which mock itself needs to resolve
            # the next target.
            patch(
                "src.runtime.worker.load_base_instructions",
                return_value="vendored prompt",
            ),
            patch(
                "src.runtime.worker.importlib.import_module",
                side_effect=import_module,
            ),
        ):
            asyncio.run(
                run_worker(
                    _request(input_modalities=["text"]),
                    stdout=io.StringIO(),
                    workspace=Path(directory),
                )
            )

        overrides = capture["codex_config"].kwargs["config_overrides"]
        self.assertEqual(len(overrides), 1)
        key, _, quoted_path = overrides[0].partition("=")
        # The thread config accepts this key too but never applies it, so it has
        # to reach the codex process itself.
        self.assertEqual(key, "model_catalog_json")
        self.assertNotIn("model_catalog_json", capture["thread_start"]["config"])
        path = Path(json.loads(quoted_path))
        self.addCleanup(shutil.rmtree, path.parent)
        entry = json.loads(path.read_text(encoding="utf-8"))["models"][0]
        self.assertEqual(entry["slug"], "model-a")
        self.assertEqual(entry["input_modalities"], ["text"])
        self.assertEqual(entry["base_instructions"], "vendored prompt")
        self.assertEqual(entry["context_window"], 200000)

    def test_catalog_entry_keeps_the_metadata_codex_would_have_guessed(self) -> None:
        entry = build_model_catalog(
            "model-a",
            {"context_window": None},
            ["text"],
            "vendored prompt",
        )["models"][0]

        self.assertIsNone(entry["context_window"])
        self.assertEqual(entry["shell_type"], "unified_exec")
        self.assertEqual(entry["multi_agent_version"], "v1")
        # A populated level list would start sending a reasoning field that the
        # metadata codex guesses for unknown models drops.
        self.assertEqual(entry["supported_reasoning_levels"], [])

    def test_base_instructions_refuse_to_load_for_another_codex(self) -> None:
        with patch(
            "src.runtime.worker.metadata.version",
            return_value=BASE_INSTRUCTIONS_CODEX_VERSION,
        ):
            self.assertTrue(
                load_base_instructions().startswith("You are a coding agent")
            )
        with (
            patch("src.runtime.worker.metadata.version", return_value="0.0.0"),
            self.assertRaisesRegex(RuntimeError, "requirements-container.txt"),
        ):
            load_base_instructions()

    def test_mirrored_stderr_forwards_lines_and_keeps_the_sdk_tail(self) -> None:
        sync = SimpleNamespace(_stderr_lines=deque(["earlier line"], maxlen=400))
        codex = SimpleNamespace(_client=SimpleNamespace(_sync=sync))
        mirror_codex_stderr(codex)

        captured = io.StringIO()
        with contextlib.redirect_stderr(captured):
            sync._stderr_lines.append("DEBUG response status 413")

        self.assertEqual(captured.getvalue(), "DEBUG response status 413\n")
        self.assertEqual(
            list(sync._stderr_lines), ["earlier line", "DEBUG response status 413"]
        )

    def test_sdk_exception_propagates(self) -> None:
        error = RuntimeError("provider failed")
        sdk = SimpleNamespace(
            CodexConfig=lambda **_: object(),
            AsyncCodex=lambda _: (_ for _ in ()).throw(error),
        )
        sdk_types = SimpleNamespace(ReasoningEffort=lambda value: value)

        def import_module(name: str):
            return sdk if name == "openai_codex" else sdk_types

        with (
            tempfile.TemporaryDirectory() as directory,
            patch(
                "src.runtime.worker.importlib.import_module",
                side_effect=import_module,
            ),
            self.assertRaisesRegex(RuntimeError, "provider failed"),
        ):
            asyncio.run(run_worker(_request(), workspace=Path(directory)))


if __name__ == "__main__":
    unittest.main()
