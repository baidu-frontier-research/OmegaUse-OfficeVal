from __future__ import annotations

import asyncio
import dataclasses
import json
import subprocess
import tempfile
import unittest
from pathlib import Path

from src.config import ModelSettings, RunnerConfig
from src.runtime.docker import (
    _AGENT_HOME,
    DockerRuntime,
    _agent_create_args,
    _consume_worker_stdout,
    _redact_runtime_secrets,
    verify_agent_image,
)

_SECRET = "real-secret-key"


def _config(root: Path, *, timeout_seconds: float = 7200.0) -> RunnerConfig:
    return RunnerConfig(
        dataset_root=root / "dataset",
        output_root=root / "runs",
        model="model-a",
        task_ids=("001",),
        timeout_seconds=timeout_seconds,
        models_path=root / ".models.yaml",
    )


class _FakeStdin:
    def __init__(self) -> None:
        self.data = bytearray()

    def write(self, value: bytes) -> None:
        self.data.extend(value)

    async def drain(self) -> None:
        pass

    def close(self) -> None:
        pass

    async def wait_closed(self) -> None:
        pass


class _FakeAgentProcess:
    def __init__(self, records: list[dict[str, object]] = ()) -> None:
        self.stdin = _FakeStdin()
        self.stdout = asyncio.StreamReader()
        self.stderr = asyncio.StreamReader()
        self.stderr.feed_eof()
        self.returncode: int | None = None
        self._finished = asyncio.Event()
        for record in records:
            self.stdout.feed_data((json.dumps(record) + "\n").encode())

    async def wait(self) -> int:
        await self._finished.wait()
        return self.returncode  # type: ignore[return-value]

    def finish(self, returncode: int) -> None:
        if self.returncode is None:
            self.returncode = returncode
            self.stdout.feed_eof()
            self._finished.set()

    def kill(self) -> None:
        self.finish(-9)


class _FakeCli:
    def __init__(self, process: _FakeAgentProcess, *, oom_killed: bool = False):
        self.process = process
        self.oom_killed = oom_killed
        self.calls: list[tuple[str, ...]] = []

    async def spawn(self, *args: str) -> _FakeAgentProcess:
        self.calls.append(args)
        return self.process

    async def run(
        self, *args: str, **_: object
    ) -> subprocess.CompletedProcess[str]:
        self.calls.append(args)
        stdout = ""
        if args[0] == "stop":
            self.process.finish(143)
        elif args[0] == "inspect":
            stdout = json.dumps(
                {
                    "ExitCode": self.process.returncode,
                    "OOMKilled": self.oom_killed,
                }
            )
        return subprocess.CompletedProcess(args, 0, stdout, "")


class _ImageCli:
    def __init__(self, env: list[str] | None) -> None:
        self.env = env
        self.calls: list[tuple[str, ...]] = []

    async def run(self, *args: str) -> subprocess.CompletedProcess[str]:
        self.calls.append(args)
        return subprocess.CompletedProcess(args, 0, json.dumps(self.env), "")


class DockerRuntimeTests(unittest.TestCase):
    def test_image_home_must_be_the_agents_writable_tmpfs(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            config = _config(Path(directory))
            cli = _ImageCli(["PATH=/usr/bin", "HOME=/home/officeval"])
            asyncio.run(verify_agent_image(config, cli=cli))  # type: ignore[arg-type]
            self.assertEqual(
                cli.calls,
                [("image", "inspect", config.docker_image, "--format", "{{json .Config.Env}}")],
            )

            for env in (["HOME=/home/officebench"], ["PATH=/usr/bin"], None):
                with self.assertRaisesRegex(RuntimeError, "rebuild the image"):
                    asyncio.run(
                        verify_agent_image(config, cli=_ImageCli(env))  # type: ignore[arg-type]
                    )

    def test_dockerfile_home_matches_the_agents_writable_tmpfs(self) -> None:
        dockerfile = (Path(__file__).parents[1] / "Dockerfile").read_text(
            encoding="utf-8"
        )
        self.assertIn(f"HOME={_AGENT_HOME}", dockerfile)
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            args = _agent_create_args(
                config=_config(root),
                name="agent-name",
                workspace=root,
                container_uid=1,
                container_gid=1,
            )
        self.assertTrue(
            any(arg.startswith(f"{_AGENT_HOME}:rw,") for arg in args)
        )

    def test_create_args_use_one_container_and_the_default_network(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            workspace = root / "workspace"
            workspace.mkdir()
            config = _config(root)
            args = _agent_create_args(
                config=config,
                name="agent-name",
                workspace=workspace,
                container_uid=1234,
                container_gid=2345,
            )

        joined = " ".join(args)
        self.assertEqual(args.count("--mount"), 1)
        self.assertIn(f"type=bind,src={workspace.resolve()},dst=/workspace", args)
        self.assertNotIn("--network", args)
        self.assertEqual(args[-1], config.docker_image)
        self.assertIn("--interactive", args)
        self.assertIn("--read-only", args)
        self.assertEqual(args[args.index("--user") + 1], "1234:2345")
        self.assertNotIn("/var/run/docker.sock", joined)
        self.assertNotIn(".models.yaml", joined)
        self.assertNotIn("--env", args)

    def test_debug_filter_reaches_the_container_as_rust_log(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            config = dataclasses.replace(_config(root), debug="codex_core=debug")
            args = _agent_create_args(
                config=config,
                name="agent-name",
                workspace=root,
                container_uid=1,
                container_gid=1,
            )

        self.assertEqual(args[args.index("--env") + 1], "RUST_LOG=codex_core=debug")
        self.assertEqual(args[-1], config.docker_image)

    def test_image_runs_worker_directly_from_source(self) -> None:
        dockerfile = (Path(__file__).parents[1] / "Dockerfile").read_text(
            encoding="utf-8"
        )
        self.assertIn("PYTHONPATH=/opt/officeval-src", dockerfile)
        self.assertIn("-r requirements-container.txt", dockerfile)
        self.assertNotIn("pyproject.toml", dockerfile)
        self.assertIn(
            'ENTRYPOINT ["python", "-P", "-m", "src.runtime.worker"]',
            dockerfile,
        )

    def test_image_keeps_libreoffice_off_the_gpu(self) -> None:
        dockerfile = (Path(__file__).parents[1] / "Dockerfile").read_text(
            encoding="utf-8"
        )
        for setting in (
            "SAL_DISABLEGL=1",
            "SAL_DISABLESKIA=1",
            "SAL_SKIA=raster",
            "SAL_DISABLE_OPENCL=1",
        ):
            self.assertIn(setting, dockerfile)
        # Set before the build-time smoke test, so it runs the same way.
        self.assertLess(
            dockerfile.index("SAL_DISABLE_OPENCL=1"),
            dockerfile.index("soffice --headless --norestore"),
        )

    def test_host_redaction_removes_reflected_key(self) -> None:
        redacted = _redact_runtime_secrets(
            {
                "authorization": f"Bearer {_SECRET}",
                "payload": [f"prefix {_SECRET} suffix"],
            },
            _SECRET,
        )
        self.assertNotIn(_SECRET, json.dumps(redacted))
        self.assertEqual(redacted["authorization"], "[REDACTED]")

    def test_streaming_deltas_are_left_out_of_events_file(self) -> None:
        methods = [
            "item/started",
            "item/agentMessage/delta",
            "item/commandExecution/outputDelta",
            "item/reasoning/textDelta",
            "item/completed",
        ]
        records: list[dict[str, object]] = [
            {
                "type": "event",
                "method": method,
                "payload": {"item": {"type": "commandExecution"}},
            }
            for method in methods
        ]
        records.append({"type": "result", "agent": {}})

        async def consume(events_path: Path) -> None:
            process = _FakeAgentProcess(records)
            process.finish(0)
            await _consume_worker_stdout(
                process,  # type: ignore[arg-type]
                events_path=events_path,
                output_prefix="task",
                secrets=(),
            )

        with tempfile.TemporaryDirectory() as directory:
            events_path = Path(directory) / "events.jsonl"
            asyncio.run(consume(events_path))
            written = [
                json.loads(line)["method"]
                for line in events_path.read_text().splitlines()
            ]
        self.assertEqual(written, ["item/started", "item/completed"])

    def test_worker_json_error_is_not_repackaged(self) -> None:
        class Process:
            def __init__(self) -> None:
                self.stdout = asyncio.StreamReader()
                self.stdout.feed_data(b"not-json\n")
                self.stdout.feed_eof()

            async def wait(self) -> int:
                return 0

        async def consume(events_path: Path) -> None:
            await _consume_worker_stdout(
                Process(),  # type: ignore[arg-type]
                events_path=events_path,
                output_prefix="task",
                secrets=(),
            )

        with tempfile.TemporaryDirectory() as directory:
            with self.assertRaises(json.JSONDecodeError):
                asyncio.run(consume(Path(directory) / "events.jsonl"))

    def test_unknown_worker_record_raises_its_natural_key_error(self) -> None:
        async def consume(events_path: Path) -> None:
            process = _FakeAgentProcess([{"type": "unexpected"}])
            process.finish(0)
            await _consume_worker_stdout(
                process,  # type: ignore[arg-type]
                events_path=events_path,
                output_prefix="task",
                secrets=(),
            )

        with tempfile.TemporaryDirectory() as directory:
            with self.assertRaisesRegex(KeyError, "agent"):
                asyncio.run(consume(Path(directory) / "events.jsonl"))

    def test_timeout_stops_and_removes_the_container(self) -> None:
        async def exercise(
            root: Path,
        ) -> tuple[dict[str, object], _FakeCli, dict[str, object]]:
            process = _FakeAgentProcess()
            cli = _FakeCli(process)
            runtime = DockerRuntime(
                _config(root, timeout_seconds=0.001),
                ModelSettings(
                    base_url="https://gateway.example/v1",
                    api_key=_SECRET,
                ),
                "run-id",
                cli=cli,  # type: ignore[arg-type]
            )
            workspace = root / "workspace"
            workspace.mkdir()
            result = await runtime.execute_agent(
                task_id="001",
                repeat_index=1,
                workspace=workspace,
                events_path=root / "events.jsonl",
                stderr_path=root / "stderr.log",
                prompt="task",
                output_prefix="task",
            )
            request = json.loads(process.stdin.data)
            self.assertEqual(request["provider"]["api_key"], _SECRET)
            self.assertNotIn("protocol_version", request)
            return result, cli, request

        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            result, cli, request = asyncio.run(exercise(root))
            self.assertEqual(request["model"], "model-a")
            self.assertEqual(result["status"], "timeout")
            self.assertEqual(sum(call[0] == "create" for call in cli.calls), 1)
            self.assertTrue(any(call[0] == "stop" for call in cli.calls))
            self.assertTrue(any(call[0] == "rm" for call in cli.calls))
            self.assertNotIn("cleanup", result["container"])

    def test_text_only_modality_reaches_the_worker_request(self) -> None:
        async def exercise(root: Path) -> dict[str, object]:
            process = _FakeAgentProcess()
            runtime = DockerRuntime(
                _config(root, timeout_seconds=0.001),
                ModelSettings(
                    base_url="https://gateway.example/v1",
                    api_key=_SECRET,
                    input_modalities=("text",),
                ),
                "run-id",
                cli=_FakeCli(process),  # type: ignore[arg-type]
            )
            workspace = root / "workspace"
            workspace.mkdir()
            await runtime.execute_agent(
                task_id="001",
                repeat_index=1,
                workspace=workspace,
                events_path=root / "events.jsonl",
                stderr_path=root / "stderr.log",
                prompt="task",
                output_prefix="task",
            )
            return json.loads(process.stdin.data)

        with tempfile.TemporaryDirectory() as directory:
            request = asyncio.run(exercise(Path(directory)))

        self.assertEqual(request["input_modalities"], ["text"])

    def test_unset_modality_leaves_model_metadata_to_codex(self) -> None:
        async def exercise(root: Path) -> dict[str, object]:
            process = _FakeAgentProcess()
            runtime = DockerRuntime(
                _config(root, timeout_seconds=0.001),
                ModelSettings(
                    base_url="https://gateway.example/v1",
                    api_key=_SECRET,
                ),
                "run-id",
                cli=_FakeCli(process),  # type: ignore[arg-type]
            )
            workspace = root / "workspace"
            workspace.mkdir()
            await runtime.execute_agent(
                task_id="001",
                repeat_index=1,
                workspace=workspace,
                events_path=root / "events.jsonl",
                stderr_path=root / "stderr.log",
                prompt="task",
                output_prefix="task",
            )
            return json.loads(process.stdin.data)

        with tempfile.TemporaryDirectory() as directory:
            request = asyncio.run(exercise(Path(directory)))

        self.assertIsNone(request["input_modalities"])

    def test_container_death_fails_only_its_own_run(self) -> None:
        async def exercise(root: Path, exit_code: int) -> dict[str, object]:
            process = _FakeAgentProcess()
            process.finish(exit_code)
            runtime = DockerRuntime(
                _config(root),
                ModelSettings(base_url="https://gateway.example/v1", api_key="key"),
                "run-id",
                cli=_FakeCli(process),  # type: ignore[arg-type]
            )
            workspace = root / "workspace"
            workspace.mkdir()
            return await runtime.execute_agent(
                task_id="001",
                repeat_index=1,
                workspace=workspace,
                events_path=root / "events.jsonl",
                stderr_path=root / "stderr.log",
                prompt="task",
                output_prefix="task",
            )

        cases = (
            ("worker_error", 2, "exited with code 2"),
            ("sigterm", 143, "terminated by signal 15"),
            ("silent_exit", 0, "exited with code 0"),
        )
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            for name, exit_code, detail in cases:
                with self.subTest(name=name):
                    case_root = root / name
                    case_root.mkdir()
                    result = asyncio.run(exercise(case_root, exit_code))
                    self.assertEqual(result["status"], "container_failed")
                    self.assertEqual(result["container"]["exit_code"], exit_code)
                    self.assertEqual(result["error"]["type"], "ContainerFailed")
                    self.assertIn(detail, result["error"]["message"])
                    self.assertIsNone(result["usage"])


if __name__ == "__main__":
    unittest.main()
