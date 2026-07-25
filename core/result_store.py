# -*- coding: utf-8 -*-
"""
core/result_store.py

负责编号级 JSON、job.json、validation_report.json 和 summary.json 的原子写入。
"""
from __future__ import annotations

import json
from datetime import datetime, timezone, timedelta
from pathlib import Path
from typing import Optional

from . import config as cfg
from .submission_service import _write_json, load_job, update_job_status

_TZ_CST = timezone(timedelta(hours=8))


def _now_iso() -> str:
    return datetime.now(_TZ_CST).isoformat()


# ---------------------------------------------------------------------------
# 系统级结果构造
# ---------------------------------------------------------------------------

_REQUIRED_KEYS = {
    "id", "file_name", "status", "error",
    "dim1_pass", "dim1_reason", "dim2_items",
    "total_score", "max_score",
}

_MISSING_DELIVERABLE_MARKERS = (
    "文件不存在",
    "文档不存在",
    "目录不存在",
    "路径不存在",
    "缺少必需文件",
    "缺少必需文档",
    "缺少所需文件",
    "缺少所需文档",
    "未找到必需文件",
    "未找到必需文档",
    "未找到所需文件",
    "未找到所需文档",
)


def _is_missing_deliverable_error(error: object) -> bool:
    """判断 verifier 错误是否明确表示交付文件缺失。"""
    text = str(error or "")
    if any(marker in text for marker in _MISSING_DELIVERABLE_MARKERS):
        return True

    lower_text = text.lower()
    if "no such file or directory" in lower_text:
        return True

    has_deliverable_reference = any(
        keyword in text
        for keyword in ("文件", "文档", ".doc", ".xls", ".ppt", ".pdf")
    )
    if not has_deliverable_reference:
        return False
    if any(keyword in text for keyword in ("未找到", "找不到", "未发现")):
        return True
    if "未" in text and any(keyword in text for keyword in ("找到", "发现")):
        return True
    if "缺少" in text:
        return True
    return False


def validate_verifier_result(result: dict) -> list[str]:
    """
    校验 verifier 返回值结构。返回错误列表；为空则通过。
    """
    errors: list[str] = []
    missing = _REQUIRED_KEYS - result.keys()
    if missing:
        errors.append(f"缺少字段：{sorted(missing)}")
        return errors  # 缺字段时后续检查无意义

    if not isinstance(result["dim2_items"], list):
        errors.append("dim2_items 不是列表")
    else:
        for i, item in enumerate(result["dim2_items"]):
            if not isinstance(item, dict):
                errors.append(f"dim2_items[{i}] 不是字典")
                continue
            for key in ("rule", "max_delta", "delta", "hit", "detail"):
                if key not in item:
                    errors.append(f"dim2_items[{i}] 缺少字段 {key!r}")
            if "hit" in item and not isinstance(item["hit"], bool):
                errors.append(f"dim2_items[{i}].hit 不是 bool（类型：{type(item['hit']).__name__}）")
            if "delta" in item and "max_delta" in item:
                if item["delta"] not in (0, item["max_delta"]):
                    errors.append(
                        f"dim2_items[{i}] delta={item['delta']!r} 不是 0 或 max_delta={item['max_delta']!r}"
                    )

    return errors


def build_system_result(
    id_: str,
    verifier_result: Optional[dict],
    status: str,
    error: Optional[str],
    duration_seconds: float,
    started_at: str,
    retried: bool = False,
) -> dict:
    """
    在 verifier 返回值基础上追加系统字段，构造最终编号级结果。

    verifier_result 为 None 时（超时/子进程异常）构造最小错误结构。
    """
    finished_at = _now_iso()
    base: dict = verifier_result if isinstance(verifier_result, dict) else {}

    struct_errors: list[str] = []
    if verifier_result is not None and status not in ("timeout",):
        struct_errors = validate_verifier_result(verifier_result)
        if struct_errors:
            status = "error"
            error = "返回结构不合规：" + "; ".join(struct_errors)

    missing_deliverable_reason = ""
    if (
        verifier_result is not None
        and not struct_errors
        and status == "error"
        and _is_missing_deliverable_error(error)
    ):
        missing_deliverable_reason = str(error)
        status = "ok"
        error = None

    result: dict = {
        "id":               base.get("id", id_),
        "file_name":        base.get("file_name", ""),
        "status":           status,
        "error":            error,
        "dim1_pass":        False if missing_deliverable_reason else base.get("dim1_pass", False),
        "dim1_reason":      missing_deliverable_reason or base.get("dim1_reason", ""),
        "dim2_items":       base.get("dim2_items", []) if not struct_errors and not missing_deliverable_reason else [],
        "total_score":      base.get("total_score", 0) if status == "ok" and not missing_deliverable_reason else 0,
        "max_score":        base.get("max_score", 0),
        "duration_seconds": duration_seconds,
        "started_at":       started_at,
        "finished_at":      finished_at,
        "retried":          retried,
    }
    return result


# ---------------------------------------------------------------------------
# 写入接口
# ---------------------------------------------------------------------------

def save_id_result(job_id: str, result: dict) -> None:
    """立即原子写入编号级 JSON：results/<job_id>/<id>.json。"""
    id_ = result["id"]
    out = cfg.RESULTS_DIR / job_id / f"{id_}.json"
    _write_json(out, result)


def save_validation_report(job_id: str, issues: list[dict], meta: dict) -> None:
    """
    写入 validation_report.json。

    meta 建议包含 checked_at、phase1_count、phase2_count、
    fatal_count、warning_count、final_status。
    """
    report = {"meta": meta, "issues": issues}
    _write_json(cfg.RESULTS_DIR / job_id / "validation_report.json", report)


def save_summary_json(job_id: str, results: list[dict], batch_meta: dict) -> None:
    """
    写入 summary.json。

    batch_meta 建议包含 job_id、started_at、finished_at、
    total、ok、error、timeout。
    """
    summary = {"meta": batch_meta, "results": results}
    _write_json(cfg.RESULTS_DIR / job_id / "summary.json", summary)
