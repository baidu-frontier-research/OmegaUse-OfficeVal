# -*- coding: utf-8 -*-
"""
core/csv_reporter.py

在所有 worker 结束后，根据编号级结果列表统一生成：
  results/<job_id>/summary.csv
  results/<job_id>/details.csv

UTF-8 BOM 编码，兼容 Windows Excel 和中文分析工具。
"""
from __future__ import annotations

import csv
import io
import pathlib
import tempfile

from . import config as cfg

# ---------------------------------------------------------------------------
# CSV 公式注入防护
# ---------------------------------------------------------------------------
_FORMULA_CHARS = ("=", "+", "-", "@", "\t", "\r")


def _safe_text(value) -> str:
    """
    将值转为字符串，并对可能触发电子表格公式执行的前缀字符进行转义。
    """
    s = "" if value is None else str(value)
    if s and s[0] in _FORMULA_CHARS:
        s = "'" + s
    return s


# ---------------------------------------------------------------------------
# summary.csv
# ---------------------------------------------------------------------------
_SUMMARY_FIELDS = [
    "job_id", "id", "file_name", "status", "error",
    "dim1_pass", "dim1_reason", "total_score", "max_score", "min_score",
    "completion_rate",
    "duration_seconds", "started_at", "finished_at",
]


def _minimum_score(result: dict) -> int | float:
    """正向规则全部记 0、负向规则全部扣除时的理论最低分。"""
    minimum: int | float = 0
    for item in result.get("dim2_items") or []:
        max_delta = item.get("max_delta")
        if (
            isinstance(max_delta, (int, float))
            and not isinstance(max_delta, bool)
            and max_delta < 0
        ):
            minimum += max_delta
    return minimum


def _completion_rate(result: dict) -> str:
    """完成度百分比：max(0, total_score) / max_score，负分按 0 分记。"""
    if result.get("status") == "ok" and result.get("dim1_pass") is False:
        return "0.0%"
    total = result.get("total_score")
    maximum = result.get("max_score")
    if (
        not isinstance(total, (int, float)) or isinstance(total, bool)
        or not isinstance(maximum, (int, float)) or isinstance(maximum, bool)
        or maximum <= 0
    ):
        return ""
    return f"{max(0, total) / maximum * 100:.1f}%"


def _summary_row(job_id: str, r: dict) -> dict:

    return {
        "job_id":           job_id,
        "id":               str(r.get("id", "")),
        "file_name":        _safe_text(r.get("file_name", "")),
        "status":           str(r.get("status", "")),
        "error":            _safe_text(r.get("error") or ""),
        "dim1_pass":        "true" if r.get("dim1_pass") else "false",
        "dim1_reason":      _safe_text(r.get("dim1_reason", "")),
        "total_score":      str(r.get("total_score", 0)),
        "max_score":        str(r.get("max_score", 0)),
        "min_score":        str(_minimum_score(r)),
        "completion_rate":  _completion_rate(r),
        "duration_seconds": str(r.get("duration_seconds", "")),

        "started_at":       str(r.get("started_at", "")),
        "finished_at":      str(r.get("finished_at", "")),
    }


# ---------------------------------------------------------------------------
# details.csv
# ---------------------------------------------------------------------------
_DETAILS_FIELDS = [
    "job_id", "id", "rule_index", "rule", "max_delta", "delta", "hit", "detail",
]


def _detail_rows(job_id: str, r: dict) -> list[dict]:
    rows = []
    for idx, item in enumerate(r.get("dim2_items") or []):
        rows.append({
            "job_id":     job_id,
            "id":         str(r.get("id", "")),
            "rule_index": str(idx),
            "rule":       _safe_text(item.get("rule", "")),
            "max_delta":  str(item.get("max_delta", "")),
            "delta":      str(item.get("delta", "")),
            "hit":        "true" if item.get("hit") else "false",
            "detail":     _safe_text(item.get("detail", "")),
        })
    return rows


# ---------------------------------------------------------------------------
# 原子写入
# ---------------------------------------------------------------------------

def _atomic_write_csv(path: pathlib.Path, fieldnames: list[str], rows: list[dict]) -> None:
    """先写临时文件，再原子替换。使用 utf-8-sig 编码（BOM）。"""
    tmp_path = path.with_suffix(path.suffix + ".tmp")
    with tmp_path.open("w", encoding="utf-8-sig", newline="") as fh:
        writer = csv.DictWriter(fh, fieldnames=fieldnames, extrasaction="ignore")
        writer.writeheader()
        writer.writerows(rows)
    tmp_path.replace(path)


# ---------------------------------------------------------------------------
# 公开接口
# ---------------------------------------------------------------------------

def generate_csv(job_id: str, results: list[dict]) -> None:
    """
    根据编号级结果列表生成 summary.csv 和 details.csv。
    results 已按 id 排序。
    """
    result_dir = cfg.RESULTS_DIR / job_id
    result_dir.mkdir(parents=True, exist_ok=True)

    summary_rows = [_summary_row(job_id, r) for r in results]
    _atomic_write_csv(result_dir / "summary.csv", _SUMMARY_FIELDS, summary_rows)

    detail_rows: list[dict] = []
    for r in results:
        detail_rows.extend(_detail_rows(job_id, r))
    _atomic_write_csv(result_dir / "details.csv", _DETAILS_FIELDS, detail_rows)
