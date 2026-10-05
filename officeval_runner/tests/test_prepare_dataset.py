from __future__ import annotations

import contextlib
import io
import tempfile
import unittest
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import patch

from prepare_dataset import (
    DEFAULT_DATASET_REPO_ID,
    DEFAULT_DATASET_REVISION,
    build_parser,
    main,
    prepare_dataset,
    resolve_dataset_root,
)


class PrepareDatasetCliTests(unittest.TestCase):
    def test_defaults(self) -> None:
        args = build_parser().parse_args([])
        self.assertEqual(args.repo_id, DEFAULT_DATASET_REPO_ID)
        self.assertEqual(args.revision, DEFAULT_DATASET_REVISION)

    def test_dataset_root_is_not_selectable(self) -> None:
        with (
            contextlib.redirect_stderr(io.StringIO()),
            self.assertRaises(SystemExit),
        ):
            build_parser().parse_args(["--dataset-root", "somewhere"])

    def test_prepares_requested_snapshot(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            dataset_root = Path(directory) / "snapshot"
            with (
                patch(
                    "prepare_dataset.prepare_dataset",
                    return_value=dataset_root,
                ) as prepare,
                contextlib.redirect_stdout(io.StringIO()),
            ):
                status = main(
                    [
                        "--repo-id",
                        "example/dataset",
                        "--revision",
                        "requested-revision",
                        "--force-download",
                    ]
                )

        self.assertEqual(status, 0)
        prepare.assert_called_once_with(
            repo_id="example/dataset",
            revision="requested-revision",
            force_download=True,
        )


class SnapshotLocationTests(unittest.TestCase):
    def test_download_returns_the_hub_snapshot_directory(self) -> None:
        hub = {"snapshot_download": lambda **kwargs: "/cache/snapshots/abc"}
        with patch.dict("sys.modules", {"huggingface_hub": _module(hub)}):
            self.assertEqual(prepare_dataset(), Path("/cache/snapshots/abc"))

    def test_download_never_materializes_a_project_local_copy(self) -> None:
        observed: dict[str, object] = {}

        def snapshot_download(**kwargs: object) -> str:
            observed.update(kwargs)
            return "/cache/snapshots/abc"

        with patch.dict(
            "sys.modules",
            {"huggingface_hub": _module({"snapshot_download": snapshot_download})},
        ):
            prepare_dataset()

        self.assertNotIn("local_dir", observed)
        self.assertEqual(observed["revision"], DEFAULT_DATASET_REVISION)
        self.assertEqual(observed["repo_type"], "dataset")

    def test_resolve_stays_offline(self) -> None:
        observed: dict[str, object] = {}

        def snapshot_download(**kwargs: object) -> str:
            observed.update(kwargs)
            return "/cache/snapshots/abc"

        with patch.dict(
            "sys.modules",
            {"huggingface_hub": _module({"snapshot_download": snapshot_download})},
        ):
            self.assertEqual(resolve_dataset_root(), Path("/cache/snapshots/abc"))

        self.assertTrue(observed["local_files_only"])
        self.assertEqual(observed["revision"], DEFAULT_DATASET_REVISION)


def _module(attributes: dict[str, object]) -> SimpleNamespace:
    return SimpleNamespace(**attributes)


if __name__ == "__main__":
    unittest.main()
