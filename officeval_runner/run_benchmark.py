#!/usr/bin/env python3
"""Run Office benchmark tasks with Codex inside Docker."""

from __future__ import annotations

import argparse
import asyncio
import dataclasses
import json
from pathlib import Path

from prepare_dataset import resolve_dataset_root
from src.config import RunnerConfig
from src.runner import run_suite


DEFAULT_MODELS_PATH = Path(__file__).resolve().with_name(".models.yaml")
DEFAULT_OUTPUT_ROOT = Path("runs")
# Logs one line per HTTP request with the provider's status code, and little
# else. codex_core=debug adds the turn trace, but that duplicates events.jsonl
# and dumps every reasoning summary, which costs tens of MB per run.
DEFAULT_DEBUG_LOG = "codex_http_client=debug"

# RunnerConfig is the single source of defaults: --resume falls back to them for
# knobs a suite.json lacks, and --help quotes them.
FIELD_DEFAULTS = {
    field.name: field.default for field in dataclasses.fields(RunnerConfig)
}

# Runtime knobs that may differ from the saved suite config on --resume; every
# CLI option defaults to None so an explicit value is distinguishable from an
# omitted one.
RESUME_OVERRIDABLE = (
    "concurrency",
    "max_steps",
    "timeout_seconds",
    "docker_context",
    "docker_image",
    "container_cpus",
    "container_memory",
    "container_pids_limit",
    "effort",
    "debug",
)
# Options that define what the suite is; changing them mid-suite would make the
# saved results and the new ones incomparable.
RESUME_FIXED = (
    "task",
    "model",
    "dataset_root",
    "task_language",
    "repeat",
    "output_root",
    "tag",
)


def _flags(names: tuple[str, ...]) -> str:
    return ", ".join(f"--{name.replace('_', '-')}" for name in names)


def _default(name: str) -> str:
    return f"(default: {FIELD_DEFAULTS[name]})"


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="python run_benchmark.py",
        description="Run isolated Office tasks with Codex in Docker.",
        epilog=(
            "--resume reuses the suite config stored in suite.json. Explicitly "
            f"passing {_flags(RESUME_OVERRIDABLE)} overrides it for that run; "
            f"{_flags(RESUME_FIXED)} cannot be changed."
        ),
    )
    parser.add_argument("--resume", type=Path, help="Continue an existing suite")
    parser.add_argument(
        "--force-task",
        action="append",
        help=(
            "With --resume, delete and rerun these task IDs; repeat, use "
            "comma-separated IDs, or pass 'all'"
        ),
    )
    parser.add_argument(
        "--dataset-root",
        type=Path,
        help=(
            "Dataset snapshot directory; defaults to the pinned Hugging Face "
            "snapshot prepared by prepare_dataset.py"
        ),
    )
    parser.add_argument(
        "--task-language",
        choices=("zh", "en"),
        help=f"Use Chinese or English task definitions {_default('task_language')}",
    )
    parser.add_argument(
        "--task",
        action="append",
        help="Task ID; repeat, use comma-separated IDs, or pass 'all'",
    )
    parser.add_argument("--model", help="Model ID for this runner process; no default")
    parser.add_argument(
        "--output-root",
        type=Path,
        help=f"Where new suites are created (default: {DEFAULT_OUTPUT_ROOT})",
    )
    parser.add_argument(
        "--tag",
        help="Short label appended to the run directory name, e.g. 'baseline'",
    )
    parser.add_argument(
        "--repeat", type=int, help=f"Runs per task {_default('repeats')}"
    )
    parser.add_argument(
        "--concurrency",
        type=int,
        help=f"Tasks running in parallel {_default('concurrency')}",
    )
    parser.add_argument(
        "--max-steps", type=int, help=f"Steps per run {_default('max_steps')}"
    )
    parser.add_argument(
        "--timeout-seconds",
        type=float,
        help=f"Timeout per run in seconds {_default('timeout_seconds')}",
    )
    parser.add_argument(
        "--docker-context", help=f"Docker context {_default('docker_context')}"
    )
    parser.add_argument(
        "--docker-image", help=f"Worker image {_default('docker_image')}"
    )
    parser.add_argument(
        "--container-cpus",
        type=float,
        help=f"CPU limit {_default('container_cpus')}",
    )
    parser.add_argument(
        "--container-memory", help=f"Memory limit {_default('container_memory')}"
    )
    parser.add_argument(
        "--container-pids-limit",
        type=int,
        help=f"PID limit {_default('container_pids_limit')}",
    )
    parser.add_argument(
        "--effort",
        choices=("none", "minimal", "low", "medium", "high", "xhigh"),
    )
    parser.add_argument(
        "--debug",
        nargs="?",
        const=DEFAULT_DEBUG_LOG,
        metavar="FILTER",
        help=(
            "Log what the agent's codex process is doing into each run's "
            f"container.stderr.log; off by default. Bare --debug uses "
            f"'{DEFAULT_DEBUG_LOG}'. FILTER is a RUST_LOG filter for turning "
            "individual targets up, e.g. 'codex_core=trace,hyper=debug'; keep "
            "it narrow, since hyper=trace logs whole request bodies"
        ),
    )
    return parser


def _explicit_options(args: argparse.Namespace) -> dict[str, object]:
    """Runtime options the user actually passed, keyed by RunnerConfig field."""
    return {
        name: getattr(args, name)
        for name in RESUME_OVERRIDABLE
        if getattr(args, name) is not None
    }


def _resume_config(args: argparse.Namespace, resume_dir: Path) -> tuple[RunnerConfig, tuple[str, ...]]:
    changed = tuple(name for name in RESUME_FIXED if getattr(args, name) is not None)
    if changed:
        raise ValueError(
            f"{_flags(changed)} cannot be changed with --resume; start a new suite "
            f"instead. Only {_flags(RESUME_OVERRIDABLE)} may be overridden."
        )
    suite_path = resume_dir / "suite.json"
    suite = json.loads(suite_path.read_text(encoding="utf-8"))
    saved = suite["config"]
    task_ids = tuple(suite["resolved_task_ids"])
    model = saved["model"]
    models_path = saved["models_path"]
    options = {
        name: saved.get(name, FIELD_DEFAULTS[name]) for name in RESUME_OVERRIDABLE
    }
    options.update(_explicit_options(args))
    config = RunnerConfig(
        dataset_root=Path(saved["dataset_root"]).resolve(),
        output_root=resume_dir.parent,
        model=model,
        task_ids=task_ids,
        task_language=saved["task_language"],
        repeats=saved["repeats"],
        models_path=Path(models_path).resolve(),
        tag=saved.get("tag"),
        **options,
    )
    force_task_ids = tuple(
        part.strip()
        for item in args.force_task or ()
        for part in item.split(",")
        if part.strip()
    )
    return config, force_task_ids


def _new_config(args: argparse.Namespace) -> RunnerConfig:
    if not args.task or not args.model:
        raise ValueError("--task and --model are required unless --resume is used")
    model = args.model.strip()
    if "," in model:
        raise ValueError(
            "--model accepts exactly one model; start a separate runner process for each model"
        )
    options = _explicit_options(args)
    if args.task_language is not None:
        options["task_language"] = args.task_language
    if args.repeat is not None:
        options["repeats"] = args.repeat
    return RunnerConfig(
        dataset_root=(
            args.dataset_root.resolve()
            if args.dataset_root
            else resolve_dataset_root()
        ),
        output_root=(args.output_root or DEFAULT_OUTPUT_ROOT).resolve(),
        model=model,
        task_ids=tuple(args.task),
        models_path=DEFAULT_MODELS_PATH,
        tag=args.tag,
        **options,
    )


def _run_command(args: argparse.Namespace) -> int:
    resume_dir = args.resume.resolve() if args.resume else None
    if args.force_task and resume_dir is None:
        raise ValueError("--force-task requires --resume")
    if resume_dir is None:
        config = _new_config(args)
        force_task_ids: tuple[str, ...] = ()
    else:
        config, force_task_ids = _resume_config(args, resume_dir)
    suite_dir = asyncio.run(
        run_suite(config, resume_dir=resume_dir, force_task_ids=force_task_ids)
    )
    print(f"suite saved to: {suite_dir}")
    return 0


def main(argv: list[str] | None = None) -> int:
    return _run_command(build_parser().parse_args(argv))


if __name__ == "__main__":
    raise SystemExit(main())
