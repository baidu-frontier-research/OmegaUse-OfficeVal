# Copyright 2026 Baidu Inc.
# SPDX-License-Identifier: Apache-2.0
"""Installed command-line entry point for Omegause Officeval."""


from __future__ import annotations

from collections.abc import Sequence

__all__ = ["main"]


def main(args: Sequence[str] | None = None) -> int:
    """运行 Officeval 批量评估命令行入口。"""
    from core.__main__ import main as run_core

    run_core(args)
    return 0
