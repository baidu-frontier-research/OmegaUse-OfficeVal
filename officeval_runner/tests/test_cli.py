from __future__ import annotations

import contextlib
import io
import json
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from run_benchmark import _new_config, _run_command, build_parser


def _suite(root: Path) -> dict[str, object]:
    return {
        "config": {
            "dataset_root": str(root / "dataset"),
            "model": "model-a",
            "task_language": "zh",
            "repeats": 1,
            "concurrency": 2,
            "max_steps": 120,
            "timeout_seconds": 7200.0,
            "models_path": str(root / ".models.yaml"),
            "docker_context": "default",
            "docker_image": "officeval-codex:0.4.0",
            "container_cpus": 2.0,
            "container_memory": "4g",
            "container_pids_limit": 512,
            "effort": "high",
        },
        "resolved_task_ids": ["001", "002", "003"],
    }


def _resume(directory: str, *extra: str) -> dict[str, object]:
    """Run --resume against a stub suite and capture what run_suite received."""
    root = Path(directory)
    suite_dir = root / "run-existing"
    suite_dir.mkdir()
    (suite_dir / "suite.json").write_text(json.dumps(_suite(root)), encoding="utf-8")
    observed: dict[str, object] = {}

    async def fake_run_suite(config, *, resume_dir, force_task_ids):
        observed["config"] = config
        observed["resume_dir"] = resume_dir
        observed["force_task_ids"] = force_task_ids
        return resume_dir

    args = build_parser().parse_args(["--resume", str(suite_dir), *extra])
    with patch("run_benchmark.run_suite", new=fake_run_suite):
        observed["exit_code"] = _run_command(args)
    observed["suite_dir"] = suite_dir
    return observed


class RunBenchmarkCliTests(unittest.TestCase):
    def test_new_run_defaults(self) -> None:
        args = build_parser().parse_args(["--task", "001", "--model", "m"])
        with patch("run_benchmark.resolve_dataset_root", return_value=Path("/ds")):
            config = _new_config(args)
        self.assertEqual(config.output_root, Path("runs").resolve())
        self.assertEqual(config.task_language, "zh")
        self.assertEqual(config.repeats, 1)
        self.assertEqual(config.concurrency, 1)
        self.assertEqual(config.max_steps, 200)
        self.assertEqual(config.timeout_seconds, 7200.0)
        self.assertEqual(config.docker_context, "default")
        self.assertEqual(config.docker_image, "officeval-codex:0.4.0")
        self.assertEqual(config.container_cpus, 2.0)
        self.assertEqual(config.container_memory, "4g")
        self.assertEqual(config.container_pids_limit, 512)
        self.assertIsNone(config.effort)

    def test_new_run_options_are_applied(self) -> None:
        args = build_parser().parse_args(
            [
                "--task",
                "001",
                "--model",
                "m",
                "--task-language",
                "en",
                "--repeat",
                "3",
                "--concurrency",
                "5",
                "--effort",
                "high",
            ]
        )
        with patch("run_benchmark.resolve_dataset_root", return_value=Path("/ds")):
            config = _new_config(args)
        self.assertEqual(config.task_language, "en")
        self.assertEqual(config.repeats, 3)
        self.assertEqual(config.concurrency, 5)
        self.assertEqual(config.effort, "high")

    def test_dataset_root_defaults_to_the_pinned_hub_snapshot(self) -> None:
        args = build_parser().parse_args(["--task", "001", "--model", "m"])
        self.assertIsNone(args.dataset_root)
        snapshot = Path("/cache/snapshots/abc")
        with patch(
            "run_benchmark.resolve_dataset_root", return_value=snapshot
        ) as resolve:
            self.assertEqual(_new_config(args).dataset_root, snapshot)
        resolve.assert_called_once_with()

    def test_dataset_root_override_is_resolved(self) -> None:
        args = build_parser().parse_args(
            ["--task", "001", "--model", "m", "--dataset-root", "local-snapshot"]
        )
        with patch("run_benchmark.resolve_dataset_root") as resolve:
            config = _new_config(args)
        resolve.assert_not_called()
        self.assertEqual(config.dataset_root, Path("local-snapshot").resolve())

    def test_subcommands_are_not_accepted(self) -> None:
        for subcommand in ("run", "summarize", "prepare-dataset"):
            with (
                self.subTest(subcommand=subcommand),
                contextlib.redirect_stderr(io.StringIO()),
                self.assertRaises(SystemExit),
            ):
                build_parser().parse_args([subcommand])

    def test_model_has_no_default_and_comma_is_rejected(self) -> None:
        args = build_parser().parse_args(["--task", "001"])
        self.assertIsNone(args.model)
        with self.assertRaisesRegex(ValueError, "required"):
            _run_command(args)

        args = build_parser().parse_args(
            ["--task", "001", "--model", "model-a,model-b"]
        )
        with self.assertRaisesRegex(ValueError, "separate runner process"):
            _run_command(args)

    def test_removed_legacy_flags_are_not_accepted(self) -> None:
        invocations = (
            ["--max-turns", "5"],
            ["--claude-cli", "/bin/false"],
            ["--permission-mode", "default"],
            ["--no-sandbox"],
            ["--allow-sandbox-fallback"],
            ["--skills", "none"],
        )
        for legacy in invocations:
            with self.subTest(flag=legacy[0]), contextlib.redirect_stderr(io.StringIO()):
                with self.assertRaises(SystemExit):
                    build_parser().parse_args(
                        ["--task", "001", "--model", "m", *legacy]
                    )

    def test_resume_reconstructs_config_and_defers_force_task_resolution(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            observed = _resume(directory, "--force-task", "officeval_003,001")

        config = observed["config"]
        self.assertEqual(observed["exit_code"], 0)
        self.assertEqual(config.model, "model-a")
        self.assertEqual(config.concurrency, 2)
        self.assertEqual(config.max_steps, 120)
        self.assertEqual(config.effort, "high")
        self.assertEqual(observed["resume_dir"], observed["suite_dir"].resolve())
        self.assertEqual(observed["force_task_ids"], ("officeval_003", "001"))

    def test_resume_options_override_the_saved_suite_config(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            observed = _resume(
                directory,
                "--concurrency",
                "5",
                "--effort",
                "low",
                "--container-memory",
                "8g",
            )

        config = observed["config"]
        self.assertEqual(config.concurrency, 5)
        self.assertEqual(config.effort, "low")
        self.assertEqual(config.container_memory, "8g")
        self.assertEqual(config.max_steps, 120)
        self.assertEqual(config.docker_image, "officeval-codex:0.4.0")

    def test_debug_is_off_by_default_and_takes_an_optional_filter(self) -> None:
        parser = build_parser()
        base = ["--task", "001", "--model", "m"]
        self.assertIsNone(parser.parse_args(base).debug)
        self.assertEqual(
            parser.parse_args([*base, "--debug"]).debug, "codex_http_client=debug"
        )
        self.assertEqual(
            parser.parse_args([*base, "--debug", "hyper=trace"]).debug, "hyper=trace"
        )
        # A flag following bare --debug must not be swallowed as its value.
        args = parser.parse_args([*base, "--debug", "--concurrency", "5"])
        self.assertEqual(args.debug, "codex_http_client=debug")
        self.assertEqual(args.concurrency, 5)

    def test_resume_turns_debug_on_for_a_suite_that_ran_without_it(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            observed = _resume(directory, "--debug")
        self.assertEqual(observed["config"].debug, "codex_http_client=debug")

    def test_resume_rejects_options_that_define_the_suite(self) -> None:
        for flag, value in (
            ("--model", "model-b"),
            ("--task", "001"),
            ("--repeat", "2"),
            ("--task-language", "en"),
            ("--tag", "baseline"),
            ("--output-root", "elsewhere"),
            ("--dataset-root", "elsewhere"),
        ):
            with (
                self.subTest(flag=flag),
                tempfile.TemporaryDirectory() as directory,
                self.assertRaisesRegex(ValueError, f"{flag} cannot be changed"),
            ):
                _resume(directory, flag, value)

    def test_resume_preserves_native_missing_config_error(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            (root / "suite.json").write_text(json.dumps({}), encoding="utf-8")
            args = build_parser().parse_args(["--resume", str(root)])
            with self.assertRaisesRegex(KeyError, "config"):
                _run_command(args)

    def test_force_task_requires_resume(self) -> None:
        args = build_parser().parse_args(
            [
                "--task",
                "001",
                "--model",
                "model-a",
                "--force-task",
                "001",
            ]
        )
        with self.assertRaisesRegex(ValueError, "requires --resume"):
            _run_command(args)


if __name__ == "__main__":
    unittest.main()
