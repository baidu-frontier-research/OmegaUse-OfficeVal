from __future__ import annotations

import asyncio
import json
import tempfile
import unittest
from dataclasses import replace
from pathlib import Path
from unittest.mock import patch

from src.config import ModelConfig, ModelSettings, PricingRates, RunnerConfig
from src.runner import (
    _allocate_suite_dir,
    _selected_tasks,
    build_task_prompt,
    run_suite,
)
from src.workspace import TaskSpec

_SECRET = "real-secret-key"
_NEW_SECRET = "new-secret-key"


def _agent() -> dict[str, object]:
    return {
        "status": "success",
        "native_status": "completed",
        "thread_id": "thread-id",
        "turn_id": "turn-id",
        "error": None,
        "final_response": "done",
        "agent_message_steps": 2,
        "max_steps_hit": False,
        "usage": {
            "input_tokens": 10,
            "cached_input_tokens": 2,
            "output_tokens": 4,
            "reasoning_output_tokens": 1,
            "total_tokens": 14,
        },
        "duration_ms": 50,
        "started_at": "2026-01-01T00:00:00.000Z",
        "ended_at": "2026-01-01T00:00:00.100Z",
        "wall_time_ms": 100,
        "container": {
            "name": "agent-container",
            "exit_code": 0,
            "oom_killed": False,
        },
    }


class _FakeRuntime:
    def __init__(self, *, fail_tasks: set[str] | None = None):
        self.fail_tasks = fail_tasks or set()
        self.calls: list[tuple[str, int, Path]] = []

    async def execute_agent(
        self,
        *,
        task_id: str,
        repeat_index: int,
        workspace: Path,
        **_: object,
    ) -> dict[str, object]:
        self.calls.append((task_id, repeat_index, workspace))
        if task_id in self.fail_tasks:
            raise RuntimeError(f"task {task_id} crashed")
        (workspace / f"answer-{task_id}.docx").write_bytes(b"office-output")
        return _agent()


class _SymlinkRuntime(_FakeRuntime):
    def __init__(self, target: Path):
        super().__init__()
        self.target = target

    async def execute_agent(
        self,
        *,
        task_id: str,
        repeat_index: int,
        workspace: Path,
        **_: object,
    ) -> dict[str, object]:
        self.calls.append((task_id, repeat_index, workspace))
        (workspace / "answer.docx").symlink_to(self.target)
        return _agent()


class RunnerTests(unittest.TestCase):
    def setUp(self) -> None:
        verify_image = patch("src.runner.verify_agent_image")
        verify_image.start()
        self.addCleanup(verify_image.stop)

    def _config(self, root: Path, task_ids: tuple[str, ...] = ("001",)) -> RunnerConfig:
        return RunnerConfig(
            dataset_root=root / "dataset",
            output_root=root / "runs",
            model="model-a",
            task_ids=task_ids,
            repeats=1,
            concurrency=2,
            models_path=root / ".models.yaml",
        )

    def _tasks(self, root: Path, task_ids: tuple[str, ...]) -> tuple[TaskSpec, ...]:
        return tuple(
            TaskSpec(
                task_id=task_id,
                instruction="Create an Office document.",
                task_dir=root / f"input-{task_id}",
            )
            for task_id in task_ids
        )

    def _patch_dependencies(
        self,
        tasks: tuple[TaskSpec, ...],
        runtime: _FakeRuntime,
        settings: ModelSettings | None = None,
    ):
        settings = settings or ModelSettings(
            base_url="https://gateway.example/v1", api_key=_SECRET
        )
        model_config = ModelConfig(
            settings=settings,
            pricing=PricingRates(1.0, 2.0, cache_read_per_million_usd=0.1),
        )
        return (
            patch("src.runner._selected_tasks", return_value=tasks),
            patch("src.runner.load_model_config", return_value=model_config),
            patch("src.runner.DockerRuntime", return_value=runtime),
        )

    def test_prompt_contains_only_container_business_constraints(self) -> None:
        prompt = build_task_prompt(
            TaskSpec(
                task_id="001",
                instruction="Update the spreadsheet.",
                task_dir=Path("/unused"),
            )
        )
        self.assertIn("/workspace", prompt)
        self.assertIn("LibreOffice is already installed", prompt)
        self.assertNotIn("internet", prompt)
        self.assertNotIn("subagents", prompt)
        self.assertNotIn("bwrap", prompt)
        self.assertNotIn("host", prompt.casefold())

    def test_suite_dir_name_carries_model_and_tag_and_avoids_collisions(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            config = replace(
                self._config(root), model="anthropic/claude-opus-5", tag="base line!"
            )
            first = _allocate_suite_dir(config)
            second = _allocate_suite_dir(config)

        self.assertRegex(
            first.name, r"^\d{8}-\d{4}-anthropic-claude-opus-5-base-line$"
        )
        self.assertEqual(second.name, f"{first.name}-2")

    def test_all_requires_at_least_one_discovered_task(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            config = self._config(Path(directory), ("all",))
            with self.assertRaisesRegex(ValueError, "no OfficeVal tasks selected"):
                _selected_tasks(config)

    def test_suite_runs_each_repeat_in_fresh_runtime_request_and_writes_results(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            config = self._config(root)
            config = replace(config, repeats=2)
            tasks = self._tasks(root, ("001",))
            runtime = _FakeRuntime()
            dependencies = self._patch_dependencies(tasks, runtime)
            with dependencies[0], dependencies[1], dependencies[2]:
                suite_dir = asyncio.run(run_suite(config))

            suite = json.loads((suite_dir / "suite.json").read_text())
            result_paths = sorted(suite_dir.rglob("result.json"))
            results = [json.loads(path.read_text()) for path in result_paths]

        self.assertEqual(
            [(task, repeat) for task, repeat, _ in runtime.calls],
            [("001", 1), ("001", 2)],
        )
        self.assertEqual(len({workspace for _, _, workspace in runtime.calls}), 2)
        self.assertEqual(set(suite), {"config", "resolved_task_ids"})
        self.assertNotIn(_SECRET, json.dumps(suite))
        self.assertEqual(
            [path.relative_to(suite_dir) for path in result_paths],
            [
                Path("task-001/repeat-01/result.json"),
                Path("task-001/repeat-02/result.json"),
            ],
        )
        self.assertTrue(all(result["agent"]["status"] == "success" for result in results))
        self.assertTrue(all(result["output_files"] for result in results))

    def test_runtime_exception_propagates(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            config = self._config(root, ("001", "002"))
            tasks = self._tasks(root, ("001", "002"))
            runtime = _FakeRuntime(fail_tasks={"001"})
            dependencies = self._patch_dependencies(tasks, runtime)
            with dependencies[0], dependencies[1], dependencies[2]:
                with self.assertRaises(ExceptionGroup) as raised:
                    asyncio.run(run_suite(config))

        self.assertEqual([task for task, _, _ in runtime.calls], ["001", "002"])
        self.assertIsInstance(raised.exception.exceptions[0], RuntimeError)
        self.assertIn("task 001 crashed", str(raised.exception.exceptions[0]))

    def test_agent_symlink_raises_the_original_archive_error(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            outside = root / "host-secret.txt"
            outside.write_text("do-not-archive-this", encoding="utf-8")
            config = self._config(root)
            tasks = self._tasks(root, ("001",))
            runtime = _SymlinkRuntime(outside)
            dependencies = self._patch_dependencies(tasks, runtime)
            with dependencies[0], dependencies[1], dependencies[2]:
                with self.assertRaises(ExceptionGroup) as raised:
                    asyncio.run(run_suite(config))
            outside_contents = outside.read_text()

        self.assertEqual(outside_contents, "do-not-archive-this")
        self.assertIsInstance(raised.exception.exceptions[0], ValueError)
        self.assertIn("symbolic link", str(raised.exception.exceptions[0]))

    def test_runner_does_not_revalidate_saved_suite_metadata(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            config = self._config(root)
            tasks = self._tasks(root, ("001",))
            suite_dir = root / "incomplete-suite"
            suite_dir.mkdir()
            (suite_dir / "suite.json").write_text(json.dumps({}), encoding="utf-8")
            runtime = _FakeRuntime()
            dependencies = self._patch_dependencies(tasks, runtime)
            with dependencies[0], dependencies[1], dependencies[2]:
                resumed = asyncio.run(
                    run_suite(
                        config,
                        resume_dir=suite_dir,
                    )
                )

        self.assertEqual(resumed, suite_dir)
        self.assertEqual([task for task, _, _ in runtime.calls], ["001"])

    def test_resume_uses_current_provider_settings(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            config = self._config(root, ("001", "002"))
            tasks = self._tasks(root, ("001", "002"))
            first = _FakeRuntime()
            dependencies = self._patch_dependencies(tasks, first)
            with dependencies[0], dependencies[1], dependencies[2]:
                suite_dir = asyncio.run(run_suite(config))
            failed_path = next((suite_dir / "task-001").rglob("result.json"))
            failed = json.loads(failed_path.read_text())
            failed["agent"]["status"] = "agent_failed"
            failed_path.write_text(json.dumps(failed), encoding="utf-8")

            second = _FakeRuntime()
            current_settings = ModelSettings(
                base_url="https://other-gateway.example/v1",
                api_key=_NEW_SECRET,
            )
            dependencies = self._patch_dependencies(
                tasks,
                second,
                current_settings,
            )
            with (
                dependencies[0],
                dependencies[1],
                patch(
                    "src.runner.DockerRuntime", return_value=second
                ) as runtime_constructor,
            ):
                resumed = asyncio.run(
                    run_suite(
                        config,
                        resume_dir=suite_dir,
                    )
                )
            resumed_model = json.loads(
                (suite_dir / "task-001" / "repeat-01" / "result.json").read_text()
            )["model"]

        self.assertEqual(resumed, suite_dir)
        self.assertEqual([task for task, _, _ in second.calls], ["001"])
        self.assertIs(runtime_constructor.call_args.args[1], current_settings)
        self.assertEqual(resumed_model, "model-a")

    def test_force_resume_deletes_and_reruns_only_selected_task(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            config = self._config(root, ("001", "002"))
            tasks = self._tasks(root, ("001", "002"))
            first = _FakeRuntime()
            dependencies = self._patch_dependencies(tasks, first)
            with dependencies[0], dependencies[1], dependencies[2]:
                suite_dir = asyncio.run(run_suite(config))
            marker = suite_dir / "task-001" / "old-result.txt"
            marker.write_text("old", encoding="utf-8")

            second = _FakeRuntime()
            dependencies = self._patch_dependencies(tasks, second)
            with dependencies[0], dependencies[1], dependencies[2]:
                resumed = asyncio.run(
                    run_suite(
                        config,
                        resume_dir=suite_dir,
                        force_task_ids=("officeval_001",),
                    )
                )
            marker_exists = marker.exists()

        self.assertEqual(resumed, suite_dir)
        self.assertEqual([task for task, _, _ in second.calls], ["001"])
        self.assertFalse(marker_exists)

    def test_force_resume_also_retries_other_failed_tasks(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            config = self._config(root, ("001", "002"))
            tasks = self._tasks(root, ("001", "002"))
            first = _FakeRuntime()
            dependencies = self._patch_dependencies(tasks, first)
            with dependencies[0], dependencies[1], dependencies[2]:
                suite_dir = asyncio.run(run_suite(config))

            failed_path = suite_dir / "task-002" / "repeat-01" / "result.json"
            failed = json.loads(failed_path.read_text(encoding="utf-8"))
            failed["agent"]["status"] = "agent_failed"
            failed_path.write_text(json.dumps(failed), encoding="utf-8")
            runtime = _FakeRuntime()
            dependencies = self._patch_dependencies(tasks, runtime)
            with dependencies[0], dependencies[1], dependencies[2]:
                asyncio.run(
                    run_suite(
                        config,
                        resume_dir=suite_dir,
                        force_task_ids=("001",),
                    )
                )

        self.assertEqual([task for task, _, _ in runtime.calls], ["001", "002"])


if __name__ == "__main__":
    unittest.main()
