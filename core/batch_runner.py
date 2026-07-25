# -*- coding: utf-8 -*-
"""
core/batch_runner.py

核心批量调度器：
- 汇总两阶段预检结果
- 执行用户确认闸门
- 生成任务列表
- 1 个 COM 槽位 + (MAX_WORKERS-1) 个普通槽位并发执行
- 管理超时和重试
- 收集结果并更新任务状态
"""
from __future__ import annotations

import concurrent.futures
import logging
import sys
import threading
import time
from datetime import datetime, timezone, timedelta
from pathlib import Path
from typing import Optional

from . import config as cfg
from .extractor import extract_package
from .package_validator import validate_package, validate_workspace_structure
from .document_validator import validate_documents


from .verifier_registry import VerifierRegistry, build_registry
from .result_store import (
    save_id_result, save_summary_json, save_validation_report,
    build_system_result,
)
from .submission_service import load_job, update_job_status, _write_json
from .validation_issue import ValidationIssue, Severity
from .worker import run_verifier
from .csv_reporter import generate_csv

_TZ_CST = timezone(timedelta(hours=8))
_log = logging.getLogger("core.batch_runner")


def _now_iso() -> str:
    return datetime.now(_TZ_CST).isoformat()


_DIM1_FORMAT_FAILURE_REASON = "维度一（交付内容格式要求）未通过"


def _build_dim1_format_failure_result(id_: str) -> dict:
    """Build a normal zero-score result for missing or invalid deliverables."""
    started_at = _now_iso()
    verifier_result = {
        "id": id_,
        "file_name": "",
        "status": "ok",
        "error": None,
        "dim1_pass": False,
        "dim1_reason": _DIM1_FORMAT_FAILURE_REASON,
        "dim2_items": [],
        "total_score": 0,
        "max_score": 0,
    }
    return build_system_result(
        id_, verifier_result, "ok", None, 0.0, started_at,
    )


# ---------------------------------------------------------------------------
# 用户确认闸门
# ---------------------------------------------------------------------------

def _display_issues(issues: list[ValidationIssue]) -> None:
    fatals   = [i for i in issues if i.severity == Severity.FATAL]
    warnings = [i for i in issues if i.severity == Severity.WARNING]
    if fatals:
        print(f"\n[FATAL] 发现 {len(fatals)} 个致命问题：")
        for iss in fatals:
            loc = f"  编号 {iss.id_}  " if iss.id_ else "  "
            print(f"  * [{iss.code}]{loc}{iss.message}")
            if iss.path:
                print(f"      路径：{iss.path}")
    if warnings:
        print(f"\n[WARNING] 发现 {len(warnings)} 个警告：")
        for iss in warnings:
            loc = f"  编号 {iss.id_}  " if iss.id_ else "  "
            print(f"  * [{iss.code}]{loc}{iss.message}")
            if iss.path:
                print(f"      路径：{iss.path}")


def _confirm_gate(issues: list[ValidationIssue]) -> bool:
    """
    按计划实现三种情形的确认闸门。
    返回 True 表示用户确认继续，False 表示取消。
    """
    fatals   = [i for i in issues if i.severity == Severity.FATAL]
    warnings = [i for i in issues if i.severity == Severity.WARNING]

    if fatals:
        print("\n验证失败，存在致命问题，无法开始评估。")
        return False
    elif warnings:
        prompt = f"\n存在 {len(warnings)} 个警告。是否忽略警告并继续评估？[y/N] "
    else:
        print("\n✓ 验证通过，未发现任何问题。")
        prompt = "是否开始评估？[y/N] "

    try:
        answer = input(prompt).strip().lower()
    except (EOFError, KeyboardInterrupt):
        print()
        return False
    return answer in ("y", "yes")


# ---------------------------------------------------------------------------
# 并发调度
# ---------------------------------------------------------------------------

def _validate_runtime_options(max_workers: int, timeout_seconds: int) -> None:
    if not isinstance(max_workers, int) or not (1 <= max_workers <= 100):
        raise ValueError(f"max_workers 必须是 1..100 的整数，当前值：{max_workers!r}")
    if not isinstance(timeout_seconds, int) or timeout_seconds < 10:
        raise ValueError(f"timeout_seconds 必须是不小于 10 的整数，当前值：{timeout_seconds!r}")


def _progress_bar(done: int, total: int, width: int = 24) -> str:
    if total <= 0:
        return "[" + "-" * width + "]"
    filled = min(width, int(width * done / total))
    return "[" + "#" * filled + "-" * (width - filled) + "]"


def _execute_all(
    job_id: str,
    registry: VerifierRegistry,
    workspace_dir: Path,
    *,
    max_workers: int,
    timeout_seconds: int,
    com_enabled: bool,
    format_failure_ids: set[str],
) -> list[dict]:

    """

    执行全部任务。启用 COM 时 COM verifier 始终串行；禁用时写入 skipped 结果。
    每项完成后立即落盘，并周期性输出进度与运行中任务。
    """
    all_ids = set(registry.ids())
    required_com_ids = cfg.COM_REQUIRED_VERIFIER_IDS & all_ids
    com_ids = (cfg.COM_VERIFIER_IDS & all_ids) if com_enabled else set()
    skipped_ids = set() if com_enabled else required_com_ids
    other_ids = all_ids - com_ids - skipped_ids

    # 缺失目录或无合法文档：预检警告经确认后，直接按维度一失败判 0 分。
    missing_ids = {
        id_ for id_ in all_ids
        if not (workspace_dir / cfg.dir_name(id_)).is_dir()
    }
    format_failure_ids = (set(format_failure_ids) | missing_ids) & all_ids
    com_ids -= format_failure_ids
    skipped_ids -= format_failure_ids
    other_ids -= format_failure_ids

    com_lock = threading.Lock()
    results: list[dict] = []
    state_lock = threading.Lock()
    active: dict[str, tuple[float, str]] = {}
    done_count = [0]
    total = len(registry)
    stop_progress = threading.Event()

    if format_failure_ids:
        for id_ in sorted(format_failure_ids):
            r = _build_dim1_format_failure_result(id_)
            save_id_result(job_id, r)
            results.append(r)
            done_count[0] += 1
            print(f"  [{done_count[0]:>3}/{total}] {id_}  ok  0.0s", flush=True)

    if not com_enabled:
        skip_reason = "Office COM 已禁用或当前平台不支持，跳过该 verifier"
        for id_ in sorted(skipped_ids):

            r = build_system_result(
                id_, None, "skipped", skip_reason,
                0.0, _now_iso(),
            )
            save_id_result(job_id, r)
            results.append(r)
            done_count[0] += 1
            print(f"  [{done_count[0]:>3}/{total}] {id_}  skipped  0.0s", flush=True)

    def _print_progress() -> None:

        now = time.monotonic()
        with state_lock:
            done = done_count[0]
            running = [
                f"{id_}({channel},{now - started:.0f}s)"
                for id_, (started, channel) in sorted(active.items())
            ]
        running_text = ", ".join(running) if running else "无"
        print(
            f"  进度 {_progress_bar(done, total)} {done}/{total}  运行中：{running_text}",
            flush=True,
        )


    def _progress_loop() -> None:
        _print_progress()
        while not stop_progress.wait(5):
            _print_progress()

    def _run(id_: str) -> dict:
        started_at = _now_iso()
        started_monotonic = time.monotonic()
        channel = "COM" if id_ in com_ids else "普通"
        verifier_path = registry.path_for(id_)
        document_dir = workspace_dir / cfg.dir_name(id_)

        with state_lock:
            active[id_] = (started_monotonic, channel)

        try:
            if id_ in com_ids:
                with com_lock:
                    r = run_verifier(
                        id_, verifier_path, document_dir, started_at,
                        timeout_seconds=timeout_seconds,
                        com_enabled=com_enabled,
                    )

            else:
                r = run_verifier(
                    id_, verifier_path, document_dir, started_at,
                    timeout_seconds=timeout_seconds,
                    com_enabled=com_enabled,
                )

        except Exception as exc:
            elapsed = round(time.monotonic() - started_monotonic, 2)
            r = build_system_result(
                id_, None, "error", f"{type(exc).__name__}: {exc}",
                elapsed, started_at,
            )
            _log.error("task %s raised: %s", id_, exc, exc_info=True)

        save_id_result(job_id, r)
        with state_lock:
            active.pop(id_, None)
            done_count[0] += 1
            cnt = done_count[0]
        print(f"  [{cnt:>3}/{total}] {id_}  {r['status']}  {r.get('duration_seconds', 0):.1f}s",
              flush=True)
        return r

    progress_thread = threading.Thread(target=_progress_loop, daemon=True)
    progress_thread.start()

    try:
        if not com_enabled:
            with concurrent.futures.ThreadPoolExecutor(max_workers=max_workers) as pool:
                futures = [pool.submit(_run, id_) for id_ in sorted(other_ids)]
                for fut in concurrent.futures.as_completed(futures):
                    results.append(fut.result())
        elif max_workers <= cfg.MAX_COM_WORKERS:
            with concurrent.futures.ThreadPoolExecutor(max_workers=max_workers) as pool:
                run_ids = all_ids - format_failure_ids
                futures = [pool.submit(_run, id_) for id_ in sorted(run_ids)]
                for fut in concurrent.futures.as_completed(futures):
                    results.append(fut.result())
        else:
            max_general = max_workers - cfg.MAX_COM_WORKERS
            with (
                concurrent.futures.ThreadPoolExecutor(max_workers=cfg.MAX_COM_WORKERS) as com_pool,
                concurrent.futures.ThreadPoolExecutor(max_workers=max_general) as gen_pool,
            ):
                futures = []
                futures += [com_pool.submit(_run, id_) for id_ in sorted(com_ids)]
                futures += [gen_pool.submit(_run, id_) for id_ in sorted(other_ids)]
                for fut in concurrent.futures.as_completed(futures):
                    results.append(fut.result())

    finally:
        stop_progress.set()
        progress_thread.join(timeout=1)
        _print_progress()

    return sorted(results, key=lambda r: r["id"])


# ---------------------------------------------------------------------------

# 主入口
# ---------------------------------------------------------------------------

def run_batch(
    job_id: str,
    *,
    max_workers: int | None = None,
    timeout_seconds: int | None = None,
    com_mode: str | None = None,
) -> None:

    """

    从已归档的 job（submitted 状态）开始执行完整批量评估流程。

    终端交互：显示预检摘要 → 用户确认 → 显示进度 → 打印汇总。
    """
    effective_max_workers = cfg.MAX_WORKERS if max_workers is None else max_workers
    effective_timeout_seconds = (
        cfg.DEFAULT_TASK_TIMEOUT_SECONDS if timeout_seconds is None else timeout_seconds
    )
    effective_com_mode = cfg.DEFAULT_COM_MODE if com_mode is None else com_mode
    _validate_runtime_options(effective_max_workers, effective_timeout_seconds)
    com_enabled = cfg.resolve_com_enabled(effective_com_mode)

    job = load_job(job_id)

    archive_path = Path(job["archive_path"])
    workspace_dir = Path(job["workspace_dir"])
    result_dir = Path(job["result_dir"])

    # ------------------------------------------------------------------
    # 0. 启动时系统自检
    # ------------------------------------------------------------------
    cfg.validate_config()

    try:
        registry = build_registry()
    except RuntimeError as exc:
        print(f"[系统错误] {exc}", file=sys.stderr)
        update_job_status(job_id, "failed", finished_at=_now_iso())
        return

    print(f"\n任务 ID：{job_id}")
    print(f"归档路径：{archive_path}")

    # ------------------------------------------------------------------
    # 1. 阶段一：压缩包结构与安全检查
    # ------------------------------------------------------------------
    update_job_status(job_id, "validating")
    print("\n──── 阶段一：压缩包结构检查 ────")
    phase1_issues = validate_package(archive_path)
    _display_issues(phase1_issues)

    # 判断是否可以安全解压
    p1_fatals = [i for i in phase1_issues if i.severity == Severity.FATAL]
    can_extract = not any(
        i.code in ("ZIP_BAD_FORMAT", "ZIP_READ_ERROR", "ZIP_ENCRYPTED",
                   "ZIP_BOMB", "ZIP_ABSOLUTE_PATH", "ZIP_PATH_TRAVERSAL",
                   "UNCOMPRESSED_TOO_LARGE")
        for i in p1_fatals
    )

    phase2_issues: list[ValidationIssue] = []

    if can_extract:
        # ------------------------------------------------------------------
        # 2. 解压
        # ------------------------------------------------------------------
        print("\n──── 解压压缩包 ────")
        try:
            extract_package(archive_path, workspace_dir)
            print(f"  已解压到：{workspace_dir}")
        except FileExistsError:
            # 工作目录已存在（之前任务遗留），跳过解压直接复用
            print(f"  工作目录已存在，复用：{workspace_dir}")
        except Exception as exc:
            print(f"  [错误] 解压失败：{exc}", file=sys.stderr)
            update_job_status(job_id, "validation_failed", finished_at=_now_iso())
            _save_validation_report(job_id, phase1_issues + phase2_issues, "validation_failed")
            return

        # 解压后路径安全复核
        postcheck = validate_workspace_structure(workspace_dir)
        phase1_issues += postcheck

        # ------------------------------------------------------------------
        # 3. 阶段二：编号目录内容检查
        # ------------------------------------------------------------------
        print("\n──── 阶段二：文档内容检查 ────")
        phase2_issues = validate_documents(workspace_dir)

        _display_issues(phase2_issues)


    all_issues = phase1_issues + phase2_issues
    fatals = [i for i in all_issues if i.severity == Severity.FATAL]
    format_failure_ids = {
        issue.id_ for issue in phase2_issues
        if issue.id_ is not None and issue.code in {"DIR_EMPTY", "NO_VALID_DOC"}
    }

    # ------------------------------------------------------------------

    # 4. 保存验证报告
    # ------------------------------------------------------------------
    final_val_status = "awaiting_confirmation" if not fatals else "validation_failed"
    _save_validation_report(job_id, all_issues, final_val_status)

    # ------------------------------------------------------------------
    # 5. 用户确认闸门
    # ------------------------------------------------------------------
    update_job_status(job_id, final_val_status)
    if not _confirm_gate(all_issues):
        confirmed_status = "validation_failed" if fatals else "cancelled_before_run"
        update_job_status(job_id, confirmed_status,
                          confirmed_at=_now_iso(), finished_at=_now_iso())
        print("评估已取消。")
        return

    update_job_status(job_id, "running",
                      confirmed_at=_now_iso(), started_at=_now_iso())

    # ------------------------------------------------------------------
    # 6. 批量执行
    # ------------------------------------------------------------------
    if com_enabled:
        com_status = f"{effective_com_mode}/启用，串行槽位 {cfg.MAX_COM_WORKERS}"
    else:
        skipped_count = len(cfg.COM_REQUIRED_VERIFIER_IDS & set(registry.ids()))

        com_status = f"{effective_com_mode}/禁用，跳过 {skipped_count} 项"
    print(
        f"\n──── 开始评估（并发 {effective_max_workers}，"
        f"任务超时 {effective_timeout_seconds} 秒，COM {com_status}）────"
    )

    batch_started = _now_iso()

    t0 = time.monotonic()

    results = _execute_all(
        job_id,
        registry,
        workspace_dir,
        max_workers=effective_max_workers,
        timeout_seconds=effective_timeout_seconds,
        com_enabled=com_enabled,
        format_failure_ids=format_failure_ids,
    )




    elapsed = round(time.monotonic() - t0, 2)
    batch_finished = _now_iso()

    # ------------------------------------------------------------------
    # 7. 汇总
    # ------------------------------------------------------------------
    update_job_status(job_id, "generating_results")
    counts = {"ok": 0, "error": 0, "timeout": 0, "skipped": 0, "other": 0}

    for r in results:
        s = r.get("status", "other")
        counts[s if s in counts else "other"] += 1

    batch_meta = {
        "job_id":         job_id,
        "started_at":     batch_started,
        "finished_at":    batch_finished,
        "elapsed_seconds": elapsed,
        "max_workers":    effective_max_workers,
        "timeout_seconds": effective_timeout_seconds,
        "com_mode":       effective_com_mode,
        "com_enabled":    com_enabled,
        "total":          len(results),
        "ok":             counts["ok"],
        "error":          counts["error"],
        "timeout":        counts["timeout"],
        "skipped":        counts["skipped"],

    }
    save_summary_json(job_id, results, batch_meta)

    # ------------------------------------------------------------------
    # 8. CSV
    # ------------------------------------------------------------------
    generate_csv(job_id, results)

    # ------------------------------------------------------------------
    # 9. 收尾
    # ------------------------------------------------------------------
    update_job_status(job_id, "completed", finished_at=batch_finished)

    print(f"\n──── 评估完成 ({elapsed:.1f}s) ────")
    print(
        f"  总计 {len(results)} 项：ok={counts['ok']}  error={counts['error']}  "
        f"timeout={counts['timeout']}  skipped={counts['skipped']}"
    )

    print(f"  结果目录：{result_dir}")


def _save_validation_report(
    job_id: str,
    issues: list[ValidationIssue],
    final_status: str,
) -> None:
    from datetime import datetime
    from .result_store import save_validation_report as _svr
    fatals   = sum(1 for i in issues if i.severity == Severity.FATAL)
    warnings = sum(1 for i in issues if i.severity == Severity.WARNING)
    meta = {
        "checked_at":    _now_iso(),
        "total_issues":  len(issues),
        "fatal_count":   fatals,
        "warning_count": warnings,
        "final_status":  final_status,
    }
    _svr(job_id, [i.to_dict() for i in sorted(
        issues, key=lambda x: (x.severity.value, x.id_ or "", x.path or "")
    )], meta)
