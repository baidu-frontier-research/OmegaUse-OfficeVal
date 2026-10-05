from __future__ import annotations

import json
import os
import shutil
import tempfile
import unittest
from pathlib import Path

from src.workspace import (
    archive_changed_office_files,
    changed_paths,
    discover_task_ids,
    load_task,
    normalize_task_id,
    restore_directory_access,
    snapshot_tree,
    stage_task_files,
)


class HuggingFaceTaskTests(unittest.TestCase):
    def _write_task(self, root: Path, *, language: str = "zh") -> Path:
        definitions = root / ("tasks" if language == "zh" else "task-en")
        definitions.mkdir(parents=True)
        task_file = definitions / "officeval_001.json"
        task_file.write_text(
            json.dumps(
                {
                    "id": "officeval_001",
                    "instruction": "修改当前目录中的文档。",
                    "origin_files": [
                        {
                            "url": (
                                "https://huggingface.co/datasets/"
                                "baidu-frontier-research/OmegaUse-OfficeVal/"
                                "resolve/main/task_files/officeval_001/input.docx"
                            ),
                            "dest": "input.docx",
                        }
                    ],
                },
                ensure_ascii=False,
            ),
            encoding="utf-8",
        )
        inputs = root / "task_files" / "officeval_001"
        inputs.mkdir(parents=True, exist_ok=True)
        (inputs / "input.docx").write_bytes(b"test office input")
        return task_file

    def test_loads_native_task_and_stages_only_declared_inputs(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            self._write_task(root)

            task = load_task(root, "officeval_001")
            workspace = root / "workspace"
            stage_task_files(task, workspace)

            self.assertEqual(discover_task_ids(root), ("001",))
            self.assertEqual(task.task_id, "001")
            self.assertEqual((workspace / "input.docx").read_bytes(), b"test office input")

    def test_stages_only_declared_inputs(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            self._write_task(root)
            (root / "task_files" / "officeval_001" / "unexpected.png").write_bytes(
                b"unexpected"
            )
            task = load_task(root, "001")
            workspace = root / "workspace"

            stage_task_files(task, workspace)

            self.assertFalse((workspace / "unexpected.png").exists())

    def test_normalizes_numeric_and_dataset_task_ids(self) -> None:
        self.assertEqual(normalize_task_id("1"), "001")
        self.assertEqual(normalize_task_id("001"), "001")
        self.assertEqual(normalize_task_id("officeval_001"), "001")
        with self.assertRaises(ValueError):
            normalize_task_id("../001")

    def test_discovery_and_loading_do_not_skip_malformed_task_ids(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            tasks_root = root / "tasks"
            tasks_root.mkdir()
            (tasks_root / "officeval_invalid.json").write_text("{}", encoding="utf-8")

            with self.assertRaises(ValueError):
                discover_task_ids(root)
            with self.assertRaises(ValueError):
                load_task(root, "")

    def test_rejects_agent_control_paths_from_dataset(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            task_file = self._write_task(root)
            payload = json.loads(task_file.read_text(encoding="utf-8"))
            payload["origin_files"][0]["dest"] = ".agents/skills/pwn/SKILL.md"
            task_file.write_text(json.dumps(payload), encoding="utf-8")

            with self.assertRaisesRegex(ValueError, "reserved Agent control path"):
                load_task(root, "001")

        for control_file in ("AGENTS.md", "AGENTS.override.md"):
            with (
                self.subTest(control_file=control_file),
                tempfile.TemporaryDirectory() as directory,
            ):
                root = Path(directory)
                task_file = self._write_task(root)
                payload = json.loads(task_file.read_text(encoding="utf-8"))
                payload["origin_files"][0]["dest"] = control_file
                task_file.write_text(json.dumps(payload), encoding="utf-8")
                source = root / "task_files" / "officeval_001" / "input.docx"
                source.rename(source.with_name(control_file))

                with self.assertRaisesRegex(ValueError, "reserved Agent control path"):
                    load_task(root, "001")

    def test_allows_nested_files_named_like_agent_documents(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            task_file = self._write_task(root)
            payload = json.loads(task_file.read_text(encoding="utf-8"))
            payload["origin_files"][0]["dest"] = "references/SKILL.md"
            task_file.write_text(json.dumps(payload), encoding="utf-8")
            source = root / "task_files" / "officeval_001" / "input.docx"
            nested = source.parent / "references" / "SKILL.md"
            nested.parent.mkdir()
            source.rename(nested)

            task = load_task(root, "001")
            workspace = root / "workspace"
            stage_task_files(task, workspace)

            self.assertEqual(
                (workspace / "references" / "SKILL.md").read_bytes(),
                b"test office input",
            )

    def test_refuses_symlink_leaving_the_dataset_when_staging_workspace(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            dataset_root = root / "dataset"
            self._write_task(dataset_root)
            task = load_task(dataset_root, "001")
            source = dataset_root / "task_files" / "officeval_001" / "input.docx"
            target = root / "outside-secret"
            target.write_bytes(b"secret")
            source.unlink()
            source.symlink_to(target)

            with self.assertRaisesRegex(ValueError, "symbolic-link"):
                stage_task_files(task, root / "workspace")

    def _write_cached_task(self, root: Path) -> Path:
        """Reproduce the Hugging Face cache layout: snapshot links into blobs."""
        repository = root / "hub" / "datasets--org--OfficeVal"
        dataset_root = repository / "snapshots" / "cd6ba6d"
        self._write_task(dataset_root)
        blob = repository / "blobs" / "bbca5e4a"
        blob.parent.mkdir(parents=True)
        blob.write_bytes(b"test office input")
        source = dataset_root / "task_files" / "officeval_001" / "input.docx"
        source.unlink()
        source.symlink_to(Path("../../../../blobs/bbca5e4a"))
        return dataset_root

    def test_stages_hugging_face_cache_symlinks_as_real_files(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            dataset_root = self._write_cached_task(root)

            task = load_task(dataset_root, "001")
            workspace = root / "workspace"
            stage_task_files(task, workspace)

            staged = workspace / "input.docx"
            self.assertFalse(staged.is_symlink())
            self.assertEqual(staged.read_bytes(), b"test office input")

    def test_refuses_cached_snapshot_link_leaving_the_repository(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            dataset_root = self._write_cached_task(root)
            outside = root / "outside-secret"
            outside.write_bytes(b"secret")
            source = dataset_root / "task_files" / "officeval_001" / "input.docx"
            source.unlink()
            source.symlink_to(outside)

            with self.assertRaisesRegex(ValueError, "symbolic-link"):
                stage_task_files(load_task(dataset_root, "001"), root / "workspace")

    def test_missing_staged_input_preserves_file_not_found_error(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            self._write_task(root)
            task = load_task(root, "001")
            (root / "task_files" / "officeval_001" / "input.docx").unlink()

            with self.assertRaises(FileNotFoundError):
                stage_task_files(task, root / "workspace")

    def test_records_agent_created_symlink_without_following_it(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            workspace = root / "workspace"
            workspace.mkdir()
            outside = root / "outside-secret.txt"
            outside.write_text("must not be read", encoding="utf-8")
            (workspace / "deliverable.docx").symlink_to(outside)

            snapshot = snapshot_tree(workspace)

            self.assertEqual(snapshot["deliverable.docx"][0], "symlink")
            self.assertNotIn("must not be read", json.dumps(snapshot))

    def test_snapshot_and_archive_use_file_metadata_without_hashes(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            workspace = root / "workspace"
            workspace.mkdir()
            document = workspace / "answer.docx"
            document.write_bytes(b"before")
            before = snapshot_tree(workspace)

            document.write_bytes(b"after!")
            original_mtime = before["answer.docx"][2]
            os.utime(document, ns=(original_mtime + 1, original_mtime + 1))
            after = snapshot_tree(workspace)
            changes = changed_paths(before, after)
            archived = archive_changed_office_files(
                workspace,
                root / "outputs",
                changes,
            )

        self.assertEqual(changes, ("answer.docx",))
        self.assertEqual(archived, ["answer.docx"])

    @unittest.skipIf(os.geteuid() == 0, "root ignores directory permission bits")
    def test_restores_traversal_bits_stripped_from_agent_directories(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            workspace = root / "workspace"
            extracted = workspace / "work_extract"
            nested = extracted / "word"
            nested.mkdir(parents=True)
            document = nested / "answer.docx"
            document.write_bytes(b"payload")
            before = snapshot_tree(workspace)
            document_mode = document.lstat().st_mode
            for stripped in (nested, extracted):
                stripped.chmod(0o600)

            with self.assertRaises(PermissionError):
                snapshot_tree(workspace)
            restore_directory_access(workspace)
            after = snapshot_tree(workspace)
            restored_mode = document.lstat().st_mode
            shutil.rmtree(workspace)

            self.assertIn("work_extract/word/answer.docx", after)
            self.assertEqual(changed_paths(before, after), ())
            self.assertEqual(restored_mode, document_mode)
            self.assertFalse(workspace.exists())

    def test_archive_rejects_only_selected_symlinks(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            workspace = root / "workspace"
            workspace.mkdir()
            outside = root / "outside-secret.txt"
            outside.write_text("must not be read", encoding="utf-8")
            (workspace / "notes.txt").symlink_to(outside)
            snapshot = snapshot_tree(workspace)
            changes = tuple(snapshot)

            self.assertEqual(
                archive_changed_office_files(workspace, root / "outputs", changes),
                [],
            )
            document = workspace / "answer.docx"
            document.symlink_to(outside)
            with self.assertRaisesRegex(ValueError, "symbolic link"):
                archive_changed_office_files(
                    workspace,
                    root / "outputs",
                    ("answer.docx",),
                )


if __name__ == "__main__":
    unittest.main()
