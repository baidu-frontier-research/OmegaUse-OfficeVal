# -*- coding: utf-8 -*-
"""
core/document_validator.py — 阶段二：编号目录内容检查。

逐个检查 officeval_001..officeval_100 目录中的文件数量、类型、
可读性、临时文件和缓存文件。
"""
from __future__ import annotations

from pathlib import Path

from . import config as cfg

from .validation_issue import ValidationIssue, Severity


# 脚本/可执行文件/快捷方式扩展名（Fatal）
_FORBIDDEN_EXTENSIONS: frozenset[str] = frozenset({
    ".py", ".exe", ".bat", ".cmd", ".sh", ".ps1", ".vbs", ".js",
    ".lnk", ".url", ".msi", ".com", ".dll", ".so",
})

# 可忽略的缓存/临时目录名
_IGNORABLE_DIRS: frozenset[str] = frozenset({
    "__pycache__", ".git", ".svn", "node_modules", ".DS_Store",
})

_NO_VALID_DOC_WARNING = (
    "如果确认继续测评，该项将直接判0分"
    "（原因是维度一（交付内容格式要求）未通过）"
)


def validate_documents(workspace_dir: Path) -> list[ValidationIssue]:
    """

    对 workspace_dir 下的所有 officeval_xxx 子目录执行阶段二检查。
    返回 ValidationIssue 列表（Fatal + Warning），不抛出异常。
    """
    issues: list[ValidationIssue] = []

    for id_ in sorted(cfg.VALID_IDS):
        dir_name = cfg.dir_name(id_)
        dir_path = workspace_dir / dir_name

        if not dir_path.exists() or not dir_path.is_dir():
            # 结构缺失已由阶段一报告；此处仅跳过
            continue

        issues.extend(_check_one_dir(id_, dir_path))

    return issues


def _check_one_dir(id_: str, dir_path: Path) -> list[ValidationIssue]:

    issues: list[ValidationIssue] = []
    rel_prefix = cfg.dir_name(id_)

    # 收集直接子项
    try:
        items = list(dir_path.iterdir())
    except PermissionError as exc:
        issues.append(ValidationIssue(
            severity=Severity.FATAL,
            stage="document_content",
            code="DIR_NOT_ACCESSIBLE",
            id_=id_,
            path=rel_prefix,
            message=f"目录无法访问：{exc}",
        ))
        return issues

    if not items:
        issues.append(ValidationIssue(
            severity=Severity.WARNING,
            stage="document_content",
            code="DIR_EMPTY",
            id_=id_,
            path=rel_prefix,
            message=f"目录为空，缺少必需文档；{_NO_VALID_DOC_WARNING}",
        ))
        return issues

    has_valid_doc = False

    for item in items:
        rel = f"{rel_prefix}/{item.name}"

        # 可忽略目录
        if item.is_dir():
            if item.name not in _IGNORABLE_DIRS:
                issues.append(ValidationIssue(
                    severity=Severity.WARNING,
                    stage="document_content",
                    code="UNEXPECTED_SUBDIR",
                    id_=id_,
                    path=rel,
                    message=f"发现子目录，不参与评估：{item.name}",
                ))
            continue

        if not item.is_file():
            continue

        name_lower = item.name.lower()
        ext = Path(name_lower).suffix

        # Office 临时锁文件
        if item.name.startswith("~$"):
            issues.append(ValidationIssue(
                severity=Severity.WARNING,
                stage="document_content",
                code="OFFICE_TEMP_FILE",
                id_=id_,
                path=rel,
                message="发现 Office 临时锁文件，将在评估时忽略",
            ))
            continue

        # 禁止的脚本/可执行文件
        if ext in _FORBIDDEN_EXTENSIONS:
            issues.append(ValidationIssue(
                severity=Severity.FATAL,
                stage="document_content",
                code="FORBIDDEN_FILE_TYPE",
                id_=id_,
                path=rel,
                message=f"包含禁止的文件类型 {ext}，可能为脚本或可执行文件",
            ))
            continue

        # 零字节文件
        try:
            if item.stat().st_size == 0:
                issues.append(ValidationIssue(
                    severity=Severity.FATAL,
                    stage="document_content",
                    code="FILE_ZERO_BYTES",
                    id_=id_,
                    path=rel,
                    message="文件大小为零，无法评估",
                ))
                continue
        except OSError:
            issues.append(ValidationIssue(
                severity=Severity.FATAL,
                stage="document_content",
                code="FILE_STAT_ERROR",
                id_=id_,
                path=rel,
                message="无法读取文件属性",
            ))
            continue

        if ext in cfg.ALLOWED_DOC_EXTENSIONS:
            has_valid_doc = True
        else:
            issues.append(ValidationIssue(
                severity=Severity.WARNING,
                stage="document_content",
                code="UNKNOWN_FILE_TYPE",
                id_=id_,
                path=rel,
                message=f"文件类型 {ext!r} 不在白名单中",
            ))

    if not has_valid_doc:
        issues.append(ValidationIssue(
            severity=Severity.WARNING,
            stage="document_content",
            code="NO_VALID_DOC",
            id_=id_,
            path=rel_prefix,
            message="目录中未找到任何合法文档（允许类型："
                    + ", ".join(sorted(cfg.ALLOWED_DOC_EXTENSIONS))
                    + f"）；{_NO_VALID_DOC_WARNING}",
        ))

    return issues
