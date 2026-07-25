# -*- coding: utf-8 -*-
"""
core/verifier_registry.py

扫描 verifiers/ 目录，建立编号 -> 脚本绝对路径映射，并校验完整性。
"""
from __future__ import annotations

from pathlib import Path

from . import config as cfg


class VerifierRegistry:
    """verifier 文件注册表。"""

    def __init__(self, registry: dict[str, Path]) -> None:
        # id_ -> 脚本绝对路径
        self._registry: dict[str, Path] = registry

    def path_for(self, id_: str) -> Path:
        """返回给定编号对应的脚本绝对路径；不存在时抛 KeyError。"""
        return self._registry[id_]

    def ids(self) -> list[str]:
        return sorted(self._registry)

    def __len__(self) -> int:
        return len(self._registry)


def build_registry() -> VerifierRegistry:
    """
    扫描 cfg.VERIFIERS_DIR，校验 001..100 完整性，返回 VerifierRegistry。

    校验失败时抛 RuntimeError。
    """
    verifiers_dir = cfg.VERIFIERS_DIR
    if not verifiers_dir.is_dir():
        raise RuntimeError(f"verifiers/ 目录不存在：{verifiers_dir}")

    found: dict[str, Path] = {}
    duplicates: dict[str, list[Path]] = {}

    for f in verifiers_dir.iterdir():
        if not f.is_file():
            continue
        id_ = cfg.id_from_verifier_filename(f.name)
        if id_ is None:
            continue
        if id_ in found:
            duplicates.setdefault(id_, [found[id_]]).append(f)
        else:
            found[id_] = f.resolve()

    errors: list[str] = []

    if duplicates:
        for id_, paths in duplicates.items():
            errors.append(f"  编号 {id_} 存在重复 verifier：{[str(p) for p in paths]}")

    missing = cfg.VALID_IDS - found.keys()
    if missing:
        errors.append(f"  缺少 {len(missing)} 个 verifier 文件：{sorted(missing)[:10]}"
                      + ("..." if len(missing) > 10 else ""))

    if errors:
        raise RuntimeError("verifier 注册表校验失败：\n" + "\n".join(errors))

    return VerifierRegistry(found)
