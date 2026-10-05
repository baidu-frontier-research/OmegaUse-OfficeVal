from __future__ import annotations

import asyncio
import json
import os
import subprocess
import time
from pathlib import Path
from typing import Any

from ..config import ModelSettings, RunnerConfig


_STREAM_LIMIT = 32 * 1024 * 1024
_SECRET_FIELD_NAMES = {
    "apikey",
    "authorization",
    "experimentalbearertoken",
    "proxyauthorization",
    "xapikey",
}
# The Agent's only writable home: the image's HOME must point here.
_AGENT_HOME = "/home/officeval"


class DockerCommandError(subprocess.CalledProcessError):
    """CalledProcessError that shows what Docker printed on stderr."""

    def __str__(self) -> str:
        detail = (self.stderr or "").strip() or "(no stderr)"
        return f"{super().__str__()}\ndocker stderr: {detail}"


class DockerCli:
    """Run Docker CLI commands without a shell."""

    def __init__(self, context: str):
        self.prefix = ("docker", "--context", context)

    async def run(
        self,
        *args: str,
    ) -> subprocess.CompletedProcess[str]:
        command = (*self.prefix, *args)
        process = await asyncio.create_subprocess_exec(
            *command,
            stdin=asyncio.subprocess.DEVNULL,
            stdout=asyncio.subprocess.PIPE,
            stderr=asyncio.subprocess.PIPE,
            limit=_STREAM_LIMIT,
        )
        stdout, stderr = await process.communicate()
        result = subprocess.CompletedProcess(
            command,
            process.returncode,
            stdout.decode("utf-8", errors="replace"),
            stderr.decode("utf-8", errors="replace"),
        )
        if result.returncode:
            raise DockerCommandError(
                result.returncode, command, result.stdout, result.stderr
            )
        return result

    async def spawn(self, *args: str) -> asyncio.subprocess.Process:
        return await asyncio.create_subprocess_exec(
            *self.prefix,
            *args,
            stdin=asyncio.subprocess.PIPE,
            stdout=asyncio.subprocess.PIPE,
            stderr=asyncio.subprocess.PIPE,
            limit=_STREAM_LIMIT,
        )


def _agent_create_args(
    *,
    config: RunnerConfig,
    name: str,
    workspace: Path,
    container_uid: int,
    container_gid: int,
) -> list[str]:
    user = f"{container_uid}:{container_gid}"
    # --debug carries a RUST_LOG filter. It reaches the codex binary because the
    # SDK spawns it with a copy of this environment; the worker then mirrors its
    # stderr into ours.
    log_env = ["--env", f"RUST_LOG={config.debug}"] if config.debug else []
    return [
        "create",
        "--name",
        name,
        *log_env,
        "--interactive",
        "--read-only",
        "--cap-drop=ALL",
        "--security-opt",
        "no-new-privileges=true",
        "--init",
        "--user",
        user,
        "--pids-limit",
        str(config.container_pids_limit),
        "--memory",
        config.container_memory,
        "--cpus",
        str(config.container_cpus),
        "--mount",
        f"type=bind,src={workspace.resolve()},dst=/workspace",
        "--tmpfs",
        f"/tmp:rw,nosuid,nodev,noexec,mode=700,size=2g,uid={container_uid},gid={container_gid}",
        "--tmpfs",
        f"/run:rw,nosuid,nodev,noexec,mode=700,size=64m,uid={container_uid},gid={container_gid}",
        "--tmpfs",
        f"{_AGENT_HOME}:rw,nosuid,nodev,noexec,mode=700,size=64m,uid={container_uid},gid={container_gid}",
        "--tmpfs",
        f"/codex-home:rw,nosuid,nodev,noexec,mode=700,size=256m,uid={container_uid},gid={container_gid}",
        config.docker_image,
    ]


async def verify_agent_image(
    config: RunnerConfig, *, cli: DockerCli | None = None
) -> None:
    """Refuse an image whose HOME is not the tmpfs the Agent can write to."""
    cli = cli or DockerCli(config.docker_context)
    inspected = await cli.run(
        "image", "inspect", config.docker_image, "--format", "{{json .Config.Env}}"
    )
    env = dict(item.partition("=")[::2] for item in json.loads(inspected.stdout) or ())
    home = env.get("HOME")
    if home != _AGENT_HOME:
        raise RuntimeError(
            f"{config.docker_image} sets HOME={home!r}, but the Agent can only write "
            f"to {_AGENT_HOME}; rebuild the image from the current Dockerfile"
        )


def _redact_runtime_secrets(value: Any, *secrets: str) -> Any:
    """Remove credentials from worker output before writing it to disk."""
    if isinstance(value, dict):
        return {
            str(key): (
                "[REDACTED]"
                if "".join(
                    character
                    for character in str(key).casefold()
                    if character.isalnum()
                )
                in _SECRET_FIELD_NAMES
                else _redact_runtime_secrets(item, *secrets)
            )
            for key, item in value.items()
        }
    if isinstance(value, list):
        return [_redact_runtime_secrets(item, *secrets) for item in value]
    if isinstance(value, str):
        for secret in secrets:
            if secret:
                value = value.replace(secret, "[REDACTED]")
    return value


def _exit_detail(exit_code: int) -> str:
    """Describe a container exit code the way `docker inspect` reports it."""
    if exit_code > 128:
        return f"was terminated by signal {exit_code - 128} (exit code {exit_code})"
    return f"exited with code {exit_code}"


def _is_streaming_delta(method: str) -> bool:
    """Match codex's per-chunk notifications (item/agentMessage/delta,
    item/commandExecution/outputDelta, item/reasoning/textDelta, ...).

    Their content is repeated in full by the matching item/completed event.
    """
    return method.rsplit("/", 1)[-1].casefold().endswith("delta")


def _failed_agent(
    error_type: str, message: str, started_clock: float
) -> dict[str, Any]:
    """Stand in for the result record the worker never got to emit."""
    return {
        "error": {"type": error_type, "message": message},
        "usage": None,
        "wall_time_ms": round((time.perf_counter() - started_clock) * 1000),
        "max_steps_hit": False,
    }


async def _consume_worker_stdout(
    process: asyncio.subprocess.Process,
    *,
    events_path: Path,
    output_prefix: str,
    secrets: tuple[str, ...],
) -> dict[str, Any] | None:
    agent = None
    with events_path.open("w", encoding="utf-8") as event_stream:
        while line := await process.stdout.readline():
            record = _redact_runtime_secrets(json.loads(line), *secrets)
            if record["type"] == "event":
                if _is_streaming_delta(record["method"]):
                    continue
                event_stream.write(json.dumps(record, ensure_ascii=False) + "\n")
                if record["method"] == "item/completed":
                    item = record["payload"]["item"]
                    if item["type"] == "agentMessage":
                        for output_line in item["text"].splitlines() or [""]:
                            print(f"{output_prefix} {output_line}", flush=True)
            else:
                agent = record["agent"]
    await process.wait()
    return agent


class DockerRuntime:
    """Run one Agent container for each benchmark request."""

    def __init__(
        self,
        config: RunnerConfig,
        model_settings: ModelSettings,
        run_id: str,
        *,
        cli: DockerCli | None = None,
    ) -> None:
        self.config = config
        self.model_settings = model_settings
        self.run_id = run_id
        self.cli = cli or DockerCli(config.docker_context)
        self.container_uid = os.getuid()
        self.container_gid = os.getgid()

    async def execute_agent(
        self,
        *,
        task_id: str,
        repeat_index: int,
        workspace: Path,
        events_path: Path,
        stderr_path: Path,
        prompt: str,
        output_prefix: str,
    ) -> dict[str, Any]:
        name = (
            f"officeval-agent-task-{task_id}-repeat-{repeat_index:02d}-"
            f"{self.run_id}"
        )
        create_args = _agent_create_args(
            config=self.config,
            name=name,
            workspace=workspace,
            container_uid=self.container_uid,
            container_gid=self.container_gid,
        )
        process: asyncio.subprocess.Process | None = None
        stderr_task: asyncio.Task[bytes] | None = None
        container_state: dict[str, Any] = {}
        terminal: dict[str, Any] | None = None
        stderr_text = ""
        timed_out = False
        started_clock = time.perf_counter()

        await self.cli.run(*create_args)
        try:
            process = await self.cli.spawn("start", "--attach", "--interactive", name)
            stderr_task = asyncio.create_task(process.stderr.read())
            modalities = self.model_settings.input_modalities
            request = {
                "model": self.config.model,
                "prompt": prompt,
                "max_steps": self.config.max_steps,
                "effort": self.config.effort,
                "input_modalities": list(modalities) if modalities else None,
                "provider": {
                    "base_url": self.model_settings.base_url,
                    "api_key": self.model_settings.api_key,
                    "context_window": self.model_settings.context_window,
                    "request_max_retries": self.model_settings.request_max_retries,
                    "stream_max_retries": self.model_settings.stream_max_retries,
                    "stream_idle_timeout_ms": self.model_settings.stream_idle_timeout_ms,
                },
            }
            process.stdin.write(
                (json.dumps(request, ensure_ascii=False) + "\n").encode("utf-8")
            )
            await process.stdin.drain()
            process.stdin.close()
            await process.stdin.wait_closed()

            try:
                terminal = await asyncio.wait_for(
                    _consume_worker_stdout(
                        process,
                        events_path=events_path,
                        output_prefix=output_prefix,
                        secrets=(self.model_settings.api_key,),
                    ),
                    timeout=self.config.timeout_seconds,
                )
            except TimeoutError:
                timed_out = True
                await self.cli.run("stop", "--time", "10", name)
                await process.wait()
        finally:
            try:
                try:
                    inspected = await self.cli.run(
                        "inspect", name, "--format", "{{json .State}}"
                    )
                    container_state = json.loads(inspected.stdout)
                finally:
                    await self.cli.run("rm", "--force", name)
            finally:
                if process is not None:
                    if process.returncode is None:
                        process.kill()
                    await process.wait()
                if stderr_task is not None:
                    stderr_text = _redact_runtime_secrets(
                        (await stderr_task).decode("utf-8", errors="replace"),
                        self.model_settings.api_key,
                    )
                stderr_path.write_text(stderr_text, encoding="utf-8")

        container = {
            "name": name,
            "exit_code": container_state["ExitCode"],
            "oom_killed": container_state["OOMKilled"],
        }
        if container["oom_killed"]:
            status = "container_oom"
            agent = _failed_agent(
                "ContainerOOM",
                "Docker reported that the Agent container was OOM-killed",
                started_clock,
            )
        elif timed_out:
            status = "timeout"
            agent = _failed_agent(
                "TimeoutError",
                f"agent exceeded {self.config.timeout_seconds:g} seconds",
                started_clock,
            )
        elif container["exit_code"] or terminal is None:
            # A container that dies before reporting a result fails this run
            # alone; raising here would take down every sibling run in the
            # suite's task group. Exit 143 means SIGTERM reached the container
            # from outside the runner, because our own stop sets timed_out.
            status = "container_failed"
            agent = _failed_agent(
                "ContainerFailed",
                f"the Agent container {_exit_detail(container['exit_code'])} "
                f"before reporting a result; see {stderr_path}",
                started_clock,
            )
        else:
            agent = terminal
            if agent["max_steps_hit"]:
                status = "max_steps_exceeded"
            elif agent["native_status"] == "completed":
                status = "success"
            else:
                status = "agent_failed"

        agent.update(
            {
                "status": status,
                "container": container,
            }
        )
        return agent
