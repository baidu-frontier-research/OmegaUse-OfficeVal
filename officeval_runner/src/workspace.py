from __future__ import annotations

import json
import re
import shutil
import stat
from dataclasses import dataclass
from pathlib import Path, PurePosixPath
from typing import Iterable


@dataclass(frozen=True)
class TaskSpec:
    task_id: str
    instruction: str
    task_dir: Path
    origin_files: tuple[str, ...] = ()
    content_root: Path | None = None
    """Directory allowed to hold the real bytes of this task's inputs.

    Defaults to ``task_dir``; a Hugging Face cache needs the wider repository
    directory because its snapshots only link to the shared blob store.
    """


_DATASET_ID_PATTERN = re.compile(r"officeval_(\d{3})")
_NUMERIC_ID_PATTERN = re.compile(r"\d{1,3}")
_LANGUAGE_DIRECTORIES = {"zh": "tasks", "en": "task-en"}
_HUGGING_FACE_SNAPSHOTS_DIRECTORY = "snapshots"
_RESERVED_TOP_LEVEL_DIRECTORIES = frozenset({".agents", ".codex", ".git"})
_RESERVED_TOP_LEVEL_FILES = frozenset({"AGENTS.md", "AGENTS.override.md"})
OFFICE_EXTENSIONS = {
    ".doc",
    ".docm",
    ".docx",
    ".dotm",
    ".dotx",
    ".rtf",
    ".odt",
    ".xls",
    ".xlsb",
    ".xlsm",
    ".xlsx",
    ".xltm",
    ".xltx",
    ".csv",
    ".ods",
    ".ppt",
    ".pptm",
    ".pptx",
    ".potm",
    ".potx",
    ".ppsm",
    ".ppsx",
    ".odp",
    ".pdf",
}
IGNORED_NAMES = {".DS_Store"}
SnapshotEntry = tuple[str, str] | tuple[str, int, int, int]


def normalize_task_id(value: str) -> str:
    """Normalize ``001`` and ``officeval_001`` to the runner's ``001`` form."""
    candidate = value.strip()
    dataset_match = _DATASET_ID_PATTERN.fullmatch(candidate)
    if dataset_match:
        return dataset_match.group(1)
    if _NUMERIC_ID_PATTERN.fullmatch(candidate):
        number = int(candidate)
        if 1 <= number <= 999:
            return f"{number:03d}"
    raise ValueError(f"invalid OfficeVal task id: {value!r}")


def _safe_destination(value: str, task_file: Path) -> str:
    if "\\" in value:
        raise ValueError(f"origin_files dest must use POSIX separators: {task_file}")
    destination = PurePosixPath(value)
    if (
        destination.is_absolute()
        or not destination.parts
        or any(part in {"", ".", ".."} for part in destination.parts)
    ):
        raise ValueError(f"unsafe origin_files dest {value!r}: {task_file}")
    if destination.parts[0] in _RESERVED_TOP_LEVEL_DIRECTORIES or (
        len(destination.parts) == 1
        and destination.name in _RESERVED_TOP_LEVEL_FILES
    ):
        raise ValueError(
            f"origin_files dest uses a reserved Agent control path {value!r}: {task_file}"
        )
    return destination.as_posix()


def _dataset_content_root(dataset_root: Path) -> Path:
    """Return the directory that may legitimately hold this dataset's bytes.

    A Hugging Face cache keeps file contents in ``<repo>/blobs`` and exposes
    ``<repo>/snapshots/<revision>`` as a tree of relative symbolic links into
    that store, so a cached snapshot's bytes live above the snapshot directory.
    """
    if dataset_root.parent.name == _HUGGING_FACE_SNAPSHOTS_DIRECTORY:
        return dataset_root.parent.parent
    return dataset_root


def load_task(dataset_root: Path, task_id: str, *, language: str = "zh") -> TaskSpec:
    """Load one native task definition and its local Hugging Face artifacts."""
    dataset_root = dataset_root.resolve()
    definitions_directory = _LANGUAGE_DIRECTORIES[language]

    normalized_id = normalize_task_id(task_id)
    dataset_task_id = f"officeval_{normalized_id}"
    task_file = dataset_root / definitions_directory / f"{dataset_task_id}.json"
    raw = json.loads(task_file.read_text(encoding="utf-8"))

    task_dir = dataset_root / "task_files" / dataset_task_id
    return TaskSpec(
        task_id=normalized_id,
        instruction=raw["instruction"],
        task_dir=task_dir,
        origin_files=tuple(
            _safe_destination(item["dest"], task_file)
            for item in raw["origin_files"]
        ),
        content_root=_dataset_content_root(dataset_root),
    )


def discover_task_ids(dataset_root: Path, *, language: str = "zh") -> tuple[str, ...]:
    """Discover native OmegaUse-OfficeVal task definitions."""
    tasks_root = dataset_root / _LANGUAGE_DIRECTORIES[language]
    return tuple(
        normalize_task_id(path.stem)
        for path in sorted(tasks_root.glob("officeval_*.json"))
    )


def stage_task_files(task: TaskSpec, workspace: Path) -> None:
    """Copy exactly the dataset-declared input artifacts into a fresh workspace."""
    content_root = (task.content_root or task.task_dir).resolve()
    workspace.mkdir(parents=True, exist_ok=False)
    for origin_file in task.origin_files:
        relative_path = Path(*PurePosixPath(origin_file).parts)
        source = task.task_dir / relative_path
        if not source.resolve().is_relative_to(content_root):
            raise ValueError(
                f"refusing symbolic-link task input leaving {content_root}: {source}"
            )
        destination = workspace / relative_path
        destination.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(source, destination)


def restore_directory_access(root: Path) -> None:
    """Give back the traversal bits an agent stripped from its own directories.

    Unpacking an Office file can land the mode stored in the archive on the
    extracted directories, and a directory without its execute bit leaves the
    host unable to walk or delete the run. Only directories are touched:
    `snapshot_tree` compares file ctimes, so chmod on a file would report it as
    modified. Symbolic links are never followed.
    """
    pending = [root]
    while pending:
        directory = pending.pop()
        mode = stat.S_IMODE(directory.lstat().st_mode)
        if mode & 0o700 != 0o700:
            directory.chmod(mode | 0o700)
        for entry in directory.iterdir():
            if stat.S_ISDIR(entry.lstat().st_mode):
                pending.append(entry)


def snapshot_tree(root: Path) -> dict[str, SnapshotEntry]:
    """Record file metadata without following symbolic links."""
    snapshot: dict[str, SnapshotEntry] = {}
    for path in sorted(root.rglob("*")):
        if path.name in IGNORED_NAMES:
            continue
        relative = path.relative_to(root).as_posix()
        entry_stat = path.lstat()
        if stat.S_ISLNK(entry_stat.st_mode):
            snapshot[relative] = ("symlink", str(path.readlink()))
            continue
        if not stat.S_ISREG(entry_stat.st_mode):
            continue
        snapshot[relative] = (
            "file",
            entry_stat.st_size,
            entry_stat.st_mtime_ns,
            entry_stat.st_ctime_ns,
        )
    return snapshot


def changed_paths(
    before: dict[str, SnapshotEntry], after: dict[str, SnapshotEntry]
) -> tuple[str, ...]:
    """Return added and metadata-modified paths."""
    added = after.keys() - before.keys()
    modified = {
        relative
        for relative in before.keys() & after.keys()
        if before[relative] != after[relative]
    }
    return tuple(sorted(added | modified))


def archive_changed_office_files(
    workspace: Path,
    outputs_dir: Path,
    changed_files: Iterable[str],
) -> list[str]:
    """Copy changed Office deliverables into the run's output directory."""
    selected = sorted(
        relative
        for relative in changed_files
        if Path(relative).suffix.lower() in OFFICE_EXTENSIONS
    )
    if not selected:
        return []

    archived: list[str] = []
    outputs_dir.mkdir(parents=True, exist_ok=True)
    for relative in selected:
        relative_path = PurePosixPath(relative)
        source = workspace.joinpath(*relative_path.parts)
        if source.is_symlink():
            raise ValueError(f"refusing to archive a symbolic link: {source}")
        destination = outputs_dir.joinpath(*relative_path.parts)
        destination.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(source, destination)
        archived.append(relative)
    return archived
