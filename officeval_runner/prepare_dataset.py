#!/usr/bin/env python3
"""Download the pinned OmegaUse-OfficeVal dataset into the Hugging Face cache."""

from __future__ import annotations

import argparse
from pathlib import Path

DEFAULT_DATASET_REPO_ID = "baidu-frontier-research/OmegaUse-OfficeVal"
DEFAULT_DATASET_REVISION = "cd6ba6d8fb83b3fb551e24eebc20e1fb0bd154a5"


def prepare_dataset(
    *,
    repo_id: str = DEFAULT_DATASET_REPO_ID,
    revision: str = DEFAULT_DATASET_REVISION,
    force_download: bool = False,
) -> Path:
    """Download the pinned revision and return its immutable snapshot directory."""
    from huggingface_hub import snapshot_download

    return Path(
        snapshot_download(
            repo_id=repo_id,
            repo_type="dataset",
            revision=revision,
            force_download=force_download,
        )
    )


def resolve_dataset_root(
    *,
    repo_id: str = DEFAULT_DATASET_REPO_ID,
    revision: str = DEFAULT_DATASET_REVISION,
) -> Path:
    """Locate an already-prepared snapshot without contacting Hugging Face."""
    from huggingface_hub import snapshot_download

    return Path(
        snapshot_download(
            repo_id=repo_id,
            repo_type="dataset",
            revision=revision,
            local_files_only=True,
        )
    )


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="python prepare_dataset.py",
        description="Download the pinned OmegaUse-OfficeVal dataset.",
    )
    parser.add_argument("--repo-id", default=DEFAULT_DATASET_REPO_ID)
    parser.add_argument("--revision", default=DEFAULT_DATASET_REVISION)
    parser.add_argument("--force-download", action="store_true", help="Redownload files even when the Hugging Face cache has them")
    return parser


def main(argv: list[str] | None = None) -> int:
    args = build_parser().parse_args(argv)
    dataset_root = prepare_dataset(
        repo_id=args.repo_id,
        revision=args.revision,
        force_download=args.force_download,
    )
    print(f"dataset prepared: root={dataset_root} revision={args.revision}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
