# -*- coding: utf-8 -*-
"""
core/__main__.py — python -m core 入口。

用法：
  python -m core --package "/absolute/path/to/submission.zip"
  python -m core                    （交互输入）

"""
from __future__ import annotations

import argparse
import logging
import sys

# 修正 Windows 控制台编码，避免中文 print 触发 GBK 错误
if hasattr(sys.stdout, "reconfigure"):
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass
if hasattr(sys.stderr, "reconfigure"):
    try:
        sys.stderr.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

# 配置日志（批量日志由 batch_runner 另写 batch.log）
logging.basicConfig(
    level=logging.WARNING,
    format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
    stream=sys.stderr,
)


def main(argv=None) -> None:
    from . import config as cfg
    from .submission_service import submit_package
    from .batch_runner import run_batch

    parser = argparse.ArgumentParser(
        prog="python -m core",
        description="Office 文档批量评估系统",
    )
    parser.add_argument(
        "--package", "-p",
        metavar="ZIP_PATH",
        help="待评估压缩包的绝对或相对路径，也可以是 submissions/ 中的文件名",
    )
    parser.add_argument(
        "--max-workers",
        type=int,
        default=cfg.MAX_WORKERS,
        metavar="N",
        help=f"同时运行的最大脚本数，默认 {cfg.MAX_WORKERS}",
    )
    parser.add_argument(
        "--timeout-seconds",
        type=int,
        default=cfg.DEFAULT_TASK_TIMEOUT_SECONDS,
        metavar="SECONDS",
        help=f"每个脚本的超时阈值，默认 {cfg.DEFAULT_TASK_TIMEOUT_SECONDS} 秒",
    )
    parser.add_argument(
        "--com-mode",
        choices=cfg.COM_MODE_CHOICES,
        default=cfg.DEFAULT_COM_MODE,
        help=(
            "仅控制 4 个硬依赖项：auto=Windows 启用、macOS/Linux 跳过；"
            "enabled=Windows 强制启用；disabled=全部跳过"
        ),

    )
    args = parser.parse_args(argv)
    try:
        cfg.resolve_com_enabled(args.com_mode)
    except RuntimeError as exc:
        parser.error(str(exc))

    # ------------------------------------------------------------------
    # 取得压缩包路径（参数或交互输入）

    # ------------------------------------------------------------------
    if args.package:
        package_input = args.package
    else:
        try:
            package_input = input(
                "请输入压缩包路径或 submissions 中的文件名：\n> "
            ).strip()
        except (EOFError, KeyboardInterrupt):
            print("\n已取消")
            sys.exit(0)

    if not package_input:
        print("[错误] 未提供压缩包路径", file=sys.stderr)
        sys.exit(1)

    # ------------------------------------------------------------------
    # 提交归档
    # ------------------------------------------------------------------
    try:
        job_id = submit_package(package_input)
    except (FileNotFoundError, ValueError) as exc:
        print(f"[错误] {exc}", file=sys.stderr)
        sys.exit(1)

    print(f"\n压缩包已归档，任务 ID：{job_id}")
    print(
        f"并发设置：{args.max_workers}，任务超时：{args.timeout_seconds} 秒，"
        f"COM 模式：{args.com_mode}"
    )


    # ------------------------------------------------------------------

    # 批量评估（含预检、确认、执行、CSV 生成）
    # ------------------------------------------------------------------
    try:
        run_batch(
            job_id,
            max_workers=args.max_workers,
            timeout_seconds=args.timeout_seconds,
            com_mode=args.com_mode,
        )

    except KeyboardInterrupt:

        print("\n用户中断")
        sys.exit(130)
    except Exception as exc:
        print(f"[系统错误] {exc}", file=sys.stderr)
        import traceback
        traceback.print_exc()
        sys.exit(1)


if __name__ == "__main__":
    main()
