# -*- coding: utf-8 -*-
"""
core/package_validator.py — 阶段一：压缩包整体结构与安全检查。

在解压前检查 ZIP 元数据，在解压后复核实际目录结构。
发现 Fatal 后继续检查同阶段其余项（尽量一次报告全部结构问题），
但不提前进入解压或 verifier 执行阶段。
"""
from __future__ import annotations

import zipfile
from pathlib import Path, PurePosixPath

from . import config as cfg
from .validation_issue import ValidationIssue, Severity


def validate_package(zip_path: Path) -> list[ValidationIssue]:
    """
    对 zip_path 执行阶段一静态元数据检查。

    返回 ValidationIssue 列表（Fatal + Warning），不抛出异常。
    """
    issues: list[ValidationIssue] = []

    # --- 1. 文件可读 ---
    try:
        zf = zipfile.ZipFile(zip_path, "r")
    except zipfile.BadZipFile as exc:
        issues.append(ValidationIssue(
            severity=Severity.FATAL,
            stage="package_structure",
            code="ZIP_BAD_FORMAT",
            id_=None,
            path=zip_path.name,
            message=f"ZIP 文件损坏或格式非法：{exc}",
        ))
        return issues
    except Exception as exc:
        issues.append(ValidationIssue(
            severity=Severity.FATAL,
            stage="package_structure",
            code="ZIP_READ_ERROR",
            id_=None,
            path=zip_path.name,
            message=f"无法读取 ZIP 文件：{exc}",
        ))
        return issues

    with zf:
        # --- 2. 加密检查 ---
        for member in zf.infolist():
            if member.flag_bits & 0x1:
                issues.append(ValidationIssue(
                    severity=Severity.FATAL,
                    stage="package_structure",
                    code="ZIP_ENCRYPTED",
                    id_=None,
                    path=member.filename,
                    message="压缩包包含加密成员，拒绝处理",
                ))
                # 加密时无法安全枚举内容，直接返回
                return issues

        members = zf.infolist()

        # --- 3. 成员数量 ---
        if len(members) > cfg.MAX_ZIP_MEMBER_COUNT:
            issues.append(ValidationIssue(
                severity=Severity.FATAL,
                stage="package_structure",
                code="ZIP_TOO_MANY_MEMBERS",
                id_=None,
                path=None,
                message=f"成员数 {len(members)} 超过限制 {cfg.MAX_ZIP_MEMBER_COUNT}",
            ))

        # --- 4. 压缩包总大小 ---
        zip_size = zip_path.stat().st_size
        if zip_size > cfg.MAX_ZIP_SIZE_BYTES:
            issues.append(ValidationIssue(
                severity=Severity.FATAL,
                stage="package_structure",
                code="ZIP_TOO_LARGE",
                id_=None,
                path=zip_path.name,
                message=f"ZIP 大小 {zip_size:,} 字节超过限制 {cfg.MAX_ZIP_SIZE_BYTES:,} 字节",
            ))

        total_uncompressed = 0
        for member in members:
            raw_name = member.filename

            # --- 5. 路径安全：绝对路径、目录逃逸 ---
            try:
                parts = PurePosixPath(raw_name).parts
            except Exception:
                parts = ()
            if raw_name.startswith("/") or raw_name.startswith("\\"):
                issues.append(ValidationIssue(
                    severity=Severity.FATAL,
                    stage="package_structure",
                    code="ZIP_ABSOLUTE_PATH",
                    id_=None,
                    path=raw_name,
                    message="成员路径包含绝对路径，Zip Slip 风险",
                ))
                continue
            if ".." in parts:
                issues.append(ValidationIssue(
                    severity=Severity.FATAL,
                    stage="package_structure",
                    code="ZIP_PATH_TRAVERSAL",
                    id_=None,
                    path=raw_name,
                    message="成员路径包含 '../'，目录逃逸风险",
                ))
                continue

            # --- 6. 单文件大小 ---
            if member.file_size > cfg.MAX_SINGLE_FILE_BYTES:
                issues.append(ValidationIssue(
                    severity=Severity.FATAL,
                    stage="package_structure",
                    code="FILE_TOO_LARGE",
                    id_=None,
                    path=raw_name,
                    message=f"单文件 {member.file_size:,} 字节超过限制 {cfg.MAX_SINGLE_FILE_BYTES:,} 字节",
                ))

            # --- 7. 压缩炸弹检测 ---
            if member.compress_size > 0:
                ratio = member.file_size / member.compress_size
                if ratio > cfg.MAX_COMPRESSION_RATIO:
                    issues.append(ValidationIssue(
                        severity=Severity.FATAL,
                        stage="package_structure",
                        code="ZIP_BOMB",
                        id_=None,
                        path=raw_name,
                        message=f"压缩比 {ratio:.1f} 超过限制 {cfg.MAX_COMPRESSION_RATIO}，疑似压缩炸弹",
                    ))

            total_uncompressed += member.file_size

        # --- 8. 解压后总大小 ---
        if total_uncompressed > cfg.MAX_UNCOMPRESSED_BYTES:
            issues.append(ValidationIssue(
                severity=Severity.FATAL,
                stage="package_structure",
                code="UNCOMPRESSED_TOO_LARGE",
                id_=None,
                path=None,
                message=f"解压后总大小 {total_uncompressed:,} 字节超过限制 {cfg.MAX_UNCOMPRESSED_BYTES:,} 字节",
            ))

        # --- 9. 顶层目录结构 ---
        top_dirs: set[str] = set()
        for member in members:
            parts = PurePosixPath(member.filename).parts
            if parts:
                top_dirs.add(parts[0])

        # 去掉路径中意外混入的文件（顶层文件）
        top_files = {d for d in top_dirs if not d.endswith("/")}
        # zipfile 成员中目录通常以 '/' 结尾，但也可能不带
        # 通过成员是否为目录判断
        top_dir_names: set[str] = set()
        top_file_names: set[str] = set()
        for member in members:
            parts = PurePosixPath(member.filename).parts
            if not parts:
                continue
            top = parts[0]
            if len(parts) == 1 and not member.is_dir():
                top_file_names.add(top)
            else:
                top_dir_names.add(top)

        # 顶层不应有散落文件
        if top_file_names:
            issues.append(ValidationIssue(
                severity=Severity.FATAL,
                stage="package_structure",
                code="ZIP_TOP_LEVEL_FILES",
                id_=None,
                path=None,
                message=f"压缩包根目录包含散落文件（应只有 officeval_xxx/ 子目录）：{sorted(top_file_names)[:5]}",
            ))

        # 期望的 100 个目录名
        expected_dirs = {cfg.dir_name(id_) for id_ in cfg.VALID_IDS}
        missing = expected_dirs - top_dir_names
        extra   = top_dir_names - expected_dirs

        if missing:
            issues.append(ValidationIssue(
                severity=Severity.WARNING,
                stage="package_structure",
                code="ZIP_MISSING_DIRS",
                id_=None,
                path=None,
                message=f"缺少 {len(missing)} 个编号目录"
                        f"（确认继续测评后，这些项将因维度一（交付内容格式要求）"
                        f"未通过而直接判0分）：{sorted(missing)[:10]}"

                        + ("..." if len(missing) > 10 else ""),
            ))
        if extra:
            issues.append(ValidationIssue(
                severity=Severity.FATAL,
                stage="package_structure",
                code="ZIP_EXTRA_DIRS",
                id_=None,
                path=None,
                message=f"出现 {len(extra)} 个非法顶层目录：{sorted(extra)[:10]}"
                        + ("..." if len(extra) > 10 else ""),
            ))

    return issues


def validate_workspace_structure(workspace_dir: Path) -> list[ValidationIssue]:
    """
    解压后复核 workspace_dir 下的实际目录结构（确保路径未逃逸）。
    """
    issues: list[ValidationIssue] = []
    if not workspace_dir.is_dir():
        issues.append(ValidationIssue(
            severity=Severity.FATAL,
            stage="package_structure",
            code="WORKSPACE_MISSING",
            id_=None,
            path=str(workspace_dir),
            message="工作目录不存在，解压可能未完成",
        ))
        return issues

    ws_resolved = workspace_dir.resolve()

    for item in workspace_dir.iterdir():
        real = item.resolve()
        # Zip Slip 后置校验：解压后路径仍需在 workspace 内
        try:
            real.relative_to(ws_resolved)
        except ValueError:
            issues.append(ValidationIssue(
                severity=Severity.FATAL,
                stage="package_structure",
                code="WORKSPACE_PATH_ESCAPE",
                id_=None,
                path=str(item),
                message=f"解压后路径逃出工作目录：{real}",
            ))

    return issues
