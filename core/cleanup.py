# -*- coding: utf-8 -*-
"""
core/cleanup.py — 工作目录清理程序。

用法：
  python -m core.cleanup --list
  python -m core.cleanup --job-id "GPT5.5_20260715_123456_ab12"
  python -m core.cleanup --older-than-days 30
"""
from __future__ import annotations

import argparse
import json
import logging
import shutil
import sys
from datetime import datetime, timezone, timedelta
from pathlib import Path

from . import config as cfg

_TZ_CST = timezone(timedelta(hours=8))
_TERMINAL_STATUSES = frozenset({
    "completed", "failed", "validation_failed", "cancelled_before_run",
})


def _now() -> datetime:
    return datetime.now(_TZ_CST)


# ---------------------------------------------------------------------------
# 候选目录扫描
# ---------------------------------------------------------------------------

def _load_job_status(job_id: str) -> str | None:
    """读取 job.json 中的 status；找不到或解析失败返回 None。"""
    try:
        p = cfg.RESULTS_DIR / job_id / "job.json"
        data = json.loads(p.read_text(encoding="utf-8"))
        return str(data.get("status", ""))
    except Exception:
        return None


def _dir_size_mb(path: Path) -> float:
    total = sum(f.stat().st_size for f in path.rglob("*") if f.is_file())
    return total / (1024 ** 2)


def _list_candidates(older_than_days: float | None = None) -> list[dict]:
    """
    扫描 workspaces/ 返回可清理候选列表。
    每项：{job_id, path, status, last_modified, size_mb}
    """
    candidates = []
    if not cfg.WORKSPACES_DIR.is_dir():
        return candidates

    cutoff = None
    if older_than_days is not None:
        cutoff = _now() - timedelta(days=older_than_days)

    for d in sorted(cfg.WORKSPACES_DIR.iterdir()):
        if not d.is_dir():
            continue
        job_id = d.name
        status = _load_job_status(job_id)
        if status is None:
            # 状态未知，跳过
            continue
        if status not in _TERMINAL_STATUSES:
            continue

        last_mod = datetime.fromtimestamp(d.stat().st_mtime, tz=_TZ_CST)
        if cutoff and last_mod > cutoff:
            continue

        size_mb = _dir_size_mb(d)
        candidates.append({
            "job_id":        job_id,
            "path":          d,
            "status":        status,
            "last_modified": last_mod.strftime("%Y-%m-%d %H:%M:%S"),
            "size_mb":       round(size_mb, 1),
        })

    return candidates


# ---------------------------------------------------------------------------
# 安全校验
# ---------------------------------------------------------------------------

def _safe_to_delete(path: Path) -> str | None:
    """
    检查路径是否安全可删。
    返回 None 表示安全，返回字符串表示不可删的原因。
    """
    ws = cfg.WORKSPACES_DIR.resolve()
    try:
        real = path.resolve()
    except Exception as exc:
        return f"resolve() 失败：{exc}"

    # 目标必须严格位于 workspaces/ 下
    try:
        real.relative_to(ws)
    except ValueError:
        return f"路径不在 workspaces/ 内：{real}"

    # 不能是 workspaces/ 本身
    if real == ws:
        return "不允许删除 workspaces/ 根目录"

    # 不能是符号链接或目录联接
    if path.is_symlink() or (path.is_dir() and path.stat().st_reparse_tag if hasattr(path.stat(), "st_reparse_tag") else False):
        return "路径是符号链接或目录联接，跳过"

    return None


# ---------------------------------------------------------------------------
# 删除
# ---------------------------------------------------------------------------

def _delete_workspace(job_id: str, path: Path) -> bool:
    reason = _safe_to_delete(path)
    if reason:
        print(f"  [跳过] {job_id}：{reason}")
        return False
    try:
        shutil.rmtree(path)
        print(f"  [删除] {job_id}  ({path})")
        return True
    except Exception as exc:
        print(f"  [失败] {job_id}：{exc}")
        return False


# ---------------------------------------------------------------------------
# 命令实现
# ---------------------------------------------------------------------------

def cmd_list() -> None:
    candidates = _list_candidates()
    if not candidates:
        print("没有可清理的工作目录。")
        return
    total_mb = sum(c["size_mb"] for c in candidates)
    print(f"可清理任务共 {len(candidates)} 个，预计释放 {total_mb:.1f} MB：\n")
    for c in candidates:
        print(f"  {c['job_id']:<60}  {c['status']:<22}  {c['last_modified']}  {c['size_mb']:>8.1f} MB")
    print(f"\n合计：{total_mb:.1f} MB")


def cmd_by_job_id(job_id: str) -> None:
    path = cfg.WORKSPACES_DIR / job_id
    if not path.is_dir():
        print(f"工作目录不存在：{path}")
        return

    status = _load_job_status(job_id)
    if status is None:
        print(f"无法读取 job.json，跳过 {job_id}")
        return
    if status not in _TERMINAL_STATUSES:
        print(f"任务 {job_id} 状态为 {status!r}，不在终态，跳过")
        return

    size_mb = _dir_size_mb(path)
    print(f"目标：{path}")
    print(f"状态：{status}  大小：{size_mb:.1f} MB")
    try:
        answer = input("确认清理？[y/N] ").strip().lower()
    except (EOFError, KeyboardInterrupt):
        print("\n已取消")
        return
    if answer in ("y", "yes"):
        _delete_workspace(job_id, path)
    else:
        print("已取消")


def cmd_older_than(days: float) -> None:
    candidates = _list_candidates(older_than_days=days)
    if not candidates:
        print(f"没有超过 {days} 天的可清理工作目录。")
        return
    total_mb = sum(c["size_mb"] for c in candidates)
    print(f"将清理 {len(candidates)} 个目录，共 {total_mb:.1f} MB：\n")
    for c in candidates:
        print(f"  {c['job_id']}  {c['size_mb']:.1f} MB  ({c['last_modified']})")
    try:
        answer = input("\n确认清理以上全部目录？[y/N] ").strip().lower()
    except (EOFError, KeyboardInterrupt):
        print("\n已取消")
        return
    if answer not in ("y", "yes"):
        print("已取消")
        return
    for c in candidates:
        _delete_workspace(c["job_id"], c["path"])


# ---------------------------------------------------------------------------
# CLI 入口
# ---------------------------------------------------------------------------

def main(argv=None) -> None:
    parser = argparse.ArgumentParser(
        prog="python -m core.cleanup",
        description="Office 批量评估系统 — 工作目录清理工具",
    )
    group = parser.add_mutually_exclusive_group(required=True)
    group.add_argument("--list", action="store_true", help="列出可清理任务")
    group.add_argument("--job-id", metavar="JOB_ID", help="按任务 ID 清理")
    group.add_argument("--older-than-days", type=float, metavar="N",
                       help="清理超过 N 天且已结束的工作目录")

    args = parser.parse_args(argv)

    if args.list:
        cmd_list()
    elif args.job_id:
        cmd_by_job_id(args.job_id)
    elif args.older_than_days is not None:
        cmd_older_than(args.older_than_days)


if __name__ == "__main__":
    main()
