# -*- coding: utf-8 -*-
"""
core/validation_issue.py — 统一验证问题数据结构。
"""
from __future__ import annotations

import enum
from dataclasses import dataclass, field
from typing import Optional


class Severity(str, enum.Enum):
    FATAL   = "FATAL"
    WARNING = "WARNING"


@dataclass(frozen=True)
class ValidationIssue:
    severity: Severity
    stage: str         # "package_structure" | "document_content"
    code: str          # 大写下划线常量，如 "ZIP_ENCRYPTED"
    id_: Optional[str] # 三位编号字符串，无关编号时为 None
    path: Optional[str]
    message: str

    def to_dict(self) -> dict:
        return {
            "severity": self.severity.value,
            "stage":    self.stage,
            "code":     self.code,
            "id":       self.id_,
            "path":     self.path,
            "message":  self.message,
        }
