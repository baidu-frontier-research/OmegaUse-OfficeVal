from __future__ import annotations

import asyncio
import json
import re
import shutil
import sys
from dataclasses import dataclass
from datetime import datetime
from itertools import count
from pathlib import Path
from typing import Any

from .config import PricingRates, RunnerConfig, load_model_config
from .reporting import calculate_token_cost, write_json, write_summary
from .runtime.docker import DockerRuntime, verify_agent_image
from .workspace import (
    TaskSpec,
    archive_changed_office_files,
    changed_paths,
    discover_task_ids,
    load_task,
    normalize_task_id,
    restore_directory_access,
    snapshot_tree,
    stage_task_files,
)


@dataclass(frozen=True)
class RunRequest:
    task: TaskSpec
    repeat_index: int
    run_dir: Path


def build_task_prompt(task: TaskSpec) -> str:
    return (
        "Complete the following Office-suite task using the input files in "
        f"/workspace:\n\n{task.instruction}\n\n"
        "Create or modify every final deliverable requested by the task. Preserve "
        "the requested formats and keep every final deliverable directly in "
        "/workspace, not in a subdirectory. LibreOffice is already installed."
    )


def _run_failed(result: dict[str, Any]) -> bool:
    return result["agent"]["status"] != "success"


def _prepare_resume(
    requests: list[RunRequest],
) -> tuple[dict[Path, dict[str, Any]], list[RunRequest]]:
    results: dict[Path, dict[str, Any]] = {}
    pending: list[RunRequest] = []
    for request in requests:
        result_path = request.run_dir / "result.json"
        if result_path.exists():
            result = json.loads(result_path.read_text(encoding="utf-8"))
            if not _run_failed(result):
                results[request.run_dir] = result
                continue
        if request.run_dir.exists():
            restore_directory_access(request.run_dir)
            shutil.rmtree(request.run_dir)
        pending.append(request)
    print(f"[resume] reused={len(results)} pending={len(pending)}", flush=True)
    return results, pending


async def _run_request(
    request: RunRequest,
    model: str,
    rates: PricingRates,
    semaphore: asyncio.Semaphore,
    runtime: DockerRuntime,
) -> dict[str, Any]:
    async with semaphore:
        task = request.task
        run_dir = request.run_dir
        workspace = run_dir / "workspace"
        run_dir.mkdir(parents=True, exist_ok=False)
        label = f"(task={task.task_id}, repeat={request.repeat_index:02d})"
        prefix = f"\033[94m{label}\033[0m" if sys.stdout.isatty() else label
        print(f"{prefix} [start] model={model}", flush=True)

        stage_task_files(task, workspace)
        before = snapshot_tree(workspace)
        prompt = build_task_prompt(task)
        (run_dir / "prompt.txt").write_text(prompt + "\n", encoding="utf-8")
        agent = await runtime.execute_agent(
            task_id=task.task_id,
            repeat_index=request.repeat_index,
            workspace=workspace,
            events_path=run_dir / "events.jsonl",
            stderr_path=run_dir / "container.stderr.log",
            prompt=prompt,
            output_prefix=prefix,
        )

        restore_directory_access(workspace)
        after = snapshot_tree(workspace)
        changed_files = changed_paths(before, after)
        output_files = archive_changed_office_files(
            workspace,
            run_dir / "outputs",
            changed_files,
        )
        result = {
            "task_id": task.task_id,
            "model": model,
            "repeat_index": request.repeat_index,
            "agent": agent,
            "calculated_token_cost_usd": calculate_token_cost(rates, agent["usage"]),
            "output_files": output_files,
        }
        write_json(run_dir / "result.json", result)

        status = agent["status"]
        if _run_failed(result):
            error = agent["error"]
            detail = error["message"] if error is not None else "no error detail"
            print(
                f"{prefix} [error] model={model} status={status}: {detail}",
                file=sys.stderr,
                flush=True,
            )
        else:
            print(
                f"{prefix} [done] model={model} status={status} "
                f"outputs={len(output_files)}",
                flush=True,
            )
        return result


def _selected_tasks(config: RunnerConfig) -> tuple[TaskSpec, ...]:
    selected = [
        part.strip()
        for item in config.task_ids
        for part in item.split(",")
        if part.strip()
    ]
    task_ids = (
        discover_task_ids(config.dataset_root, language=config.task_language)
        if not selected or "all" in selected
        else tuple(dict.fromkeys(normalize_task_id(task_id) for task_id in selected))
    )
    tasks = tuple(
        load_task(config.dataset_root, task_id, language=config.task_language)
        for task_id in task_ids
    )
    if not tasks:
        raise ValueError("no OfficeVal tasks selected")
    return tasks


def _resolve_force_task_ids(
    tasks: tuple[TaskSpec, ...],
    force_task_ids: tuple[str, ...],
) -> tuple[str, ...]:
    available = tuple(task.task_id for task in tasks)
    requested = tuple(
        dict.fromkeys(
            task_id if task_id == "all" else normalize_task_id(task_id)
            for task_id in force_task_ids
        )
    )
    if "all" in requested:
        return available
    available_lookup = dict.fromkeys(available)
    for task_id in requested:
        available_lookup[task_id]
    requested_set = set(requested)
    return tuple(task_id for task_id in available if task_id in requested_set)


def _slugify(value: str, max_length: int) -> str:
    """Reduce a label to something safe for a path and a Docker container name."""
    slug = re.sub(r"[^A-Za-z0-9._-]+", "-", value).strip("-._")
    return slug[:max_length].strip("-._")


def _allocate_suite_dir(config: RunnerConfig) -> Path:
    """Create a fresh suite directory named <date>-<time>-<model>[-<tag>]."""
    parts = [
        datetime.now().strftime("%Y%m%d-%H%M"),
        _slugify(config.model, 24) or "model",
    ]
    if config.tag:
        parts.append(_slugify(config.tag, 16))
    base = "-".join(part for part in parts if part)
    output_root = config.output_root.resolve()
    for attempt in count(1):
        suite_dir = output_root / (base if attempt == 1 else f"{base}-{attempt}")
        try:
            suite_dir.mkdir(parents=True, exist_ok=False)
        except FileExistsError:
            continue
        return suite_dir
    raise AssertionError("count() never stops")


async def run_suite(
    config: RunnerConfig,
    *,
    resume_dir: Path | None = None,
    force_task_ids: tuple[str, ...] = (),
) -> Path:
    tasks = _selected_tasks(config)
    forced = _resolve_force_task_ids(tasks, force_task_ids)
    model_config = load_model_config(config.models_path, config.model)
    await verify_agent_image(config)

    if resume_dir is None:
        suite_dir = _allocate_suite_dir(config)
    else:
        suite_dir = resume_dir.resolve()

    runtime = DockerRuntime(config, model_config.settings, suite_dir.name)
    requests = [
        RunRequest(
            task=task,
            repeat_index=repeat_index,
            run_dir=(
                suite_dir
                / f"task-{task.task_id}"
                / f"repeat-{repeat_index:02d}"
            ),
        )
        for task in tasks
        for repeat_index in range(1, config.repeats + 1)
    ]

    suite_path = suite_dir / "suite.json"
    if resume_dir is None:
        write_json(
            suite_path,
            {
                "config": {
                    "dataset_root": str(config.dataset_root),
                    "model": config.model,
                    "task_language": config.task_language,
                    "repeats": config.repeats,
                    "concurrency": config.concurrency,
                    "max_steps": config.max_steps,
                    "timeout_seconds": config.timeout_seconds,
                    "models_path": str(config.models_path),
                    "docker_context": config.docker_context,
                    "docker_image": config.docker_image,
                    "container_cpus": config.container_cpus,
                    "container_memory": config.container_memory,
                    "container_pids_limit": config.container_pids_limit,
                    "effort": config.effort,
                    "debug": config.debug,
                    "tag": config.tag,
                },
                "resolved_task_ids": [task.task_id for task in tasks],
            },
        )

    results_by_run_dir: dict[Path, dict[str, Any]] = {}
    pending_requests = requests
    if resume_dir is not None:
        for task_id in forced:
            task_dir = suite_dir / f"task-{task_id}"
            if task_dir.exists():
                restore_directory_access(task_dir)
                shutil.rmtree(task_dir)
        results_by_run_dir, pending_requests = _prepare_resume(requests)

    semaphore = asyncio.Semaphore(config.concurrency)
    async with asyncio.TaskGroup() as group:
        run_tasks = [
            group.create_task(
                _run_request(
                    request,
                    config.model,
                    model_config.pricing,
                    semaphore,
                    runtime,
                )
            )
            for request in pending_requests
        ]
    for request, run_task in zip(pending_requests, run_tasks):
        results_by_run_dir[request.run_dir] = run_task.result()

    results = [results_by_run_dir[request.run_dir] for request in requests]
    write_summary(suite_dir, results)
    failed_runs = sum(_run_failed(result) for result in results)
    if failed_runs:
        raise RuntimeError(
            f"{failed_runs} of {len(results)} runs failed; "
            f"results were saved to: {suite_dir}"
        )
    return suite_dir
