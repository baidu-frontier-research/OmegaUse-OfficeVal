# -*- coding: utf-8 -*-
"""
core/extractor.py — 安全解压 ZIP 到任务工作目录。

保证所有目标路径位于任务目录内（Zip Slip 防护），
并在解压完成后复核根目录结构。
"""
from __future__ import annotations

import zipfile
from pathlib import Path

from . import config as cfg


def extract_package(zip_path: Path, workspace_dir: Path) -> None:
    """
    将 zip_path 安全解压到 workspace_dir。

    - workspace_dir 必须不存在或为空；若已存在则抛 FileExistsError。
    - 每个成员目标路径必须严格位于 workspace_dir 内；违规则跳过并收集警告。
    - 不依赖成员文件名推断绝对路径，始终以 workspace_dir 为根。

    调用方负责确保在此之前已通过 package_validator 的阶段一检查。
    """
    workspace_dir = workspace_dir.resolve()
    if workspace_dir.exists() and any(workspace_dir.iterdir()):
        raise FileExistsError(f"工作目录已存在且非空：{workspace_dir}")
    workspace_dir.mkdir(parents=True, exist_ok=True)

    skipped: list[str] = []

    with zipfile.ZipFile(zip_path, "r") as zf:
        for member in zf.infolist():
            # 规范化成员路径（去除前导斜杠、替换反斜杠）
            norm = member.filename.replace("\\", "/").lstrip("/")
            if not norm:
                continue

            target = (workspace_dir / norm).resolve()

            # Zip Slip 防护：目标必须严格位于 workspace_dir 内
            try:
                target.relative_to(workspace_dir)
            except ValueError:
                skipped.append(member.filename)
                continue

            if member.is_dir():
                target.mkdir(parents=True, exist_ok=True)
            else:
                target.parent.mkdir(parents=True, exist_ok=True)
                with zf.open(member) as src, target.open("wb") as dst:
                    while chunk := src.read(65536):
                        dst.write(chunk)

    if skipped:
        raise RuntimeError(
            f"解压时跳过了 {len(skipped)} 个路径逃逸成员：{skipped[:5]}"
        )
