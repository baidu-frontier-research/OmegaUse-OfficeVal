# -*- coding: utf-8 -*-
"""
core/worker.py

在独立子进程中执行单个 verifier，捕获结果和异常。

子进程职责：
- 动态加载 verifier 脚本（importlib）
- 调用 evaluate(document_dir)
- 将结果序列化为 JSON 写入 stdout（唯一通道）
- 将调试信息写入 stderr

主进程职责：
- 启动子进程
- 设置超时
- 解析子进程 stdout 中的 JSON
- 处理 COM 瞬时错误重试
- 清理超时后的 Office 自动化进程
"""
from __future__ import annotations

import json
import os
import subprocess
import sys
import textwrap
import time
from pathlib import Path
from typing import Optional

from . import config as cfg
from .result_store import build_system_result


# ---------------------------------------------------------------------------
# 子进程入口脚本（内联 Python 代码，通过 -c 传入）
# ---------------------------------------------------------------------------
_WORKER_SCRIPT = textwrap.dedent("""\
import importlib.util, json, sys, pathlib, traceback

verifier_path = sys.argv[1]
document_dir  = sys.argv[2]

result = None
try:
    verifier_file = pathlib.Path(verifier_path)
    module_name = verifier_file.stem
    verifier_dir = str(verifier_file.parent)
    if verifier_dir not in sys.path:
        sys.path.insert(0, verifier_dir)

    spec = importlib.util.spec_from_file_location(module_name, verifier_path)
    mod  = importlib.util.module_from_spec(spec)
    sys.modules[spec.name] = mod
    spec.loader.exec_module(mod)


    if not hasattr(mod, "evaluate") or not callable(mod.evaluate):
        raise AttributeError("verifier 缺少可调用的 evaluate 函数")
    result = mod.evaluate(document_dir)
    if not isinstance(result, dict):
        raise TypeError(f"evaluate() 返回值不是 dict：{type(result).__name__!r}")
except Exception as exc:
    result = {
        "id":          pathlib.Path(verifier_path).stem.split("_")[1] if len(pathlib.Path(verifier_path).stem.split("_")) > 1 else "???",
        "file_name":   "",
        "status":      "error",
        "error":       f"{type(exc).__name__}: {exc}",
        "dim1_pass":   False,
        "dim1_reason": "",
        "dim2_items":  [],
        "total_score": 0,
        "max_score":   0,
    }

output = json.dumps(result, ensure_ascii=False) + "\\n"
if hasattr(sys.stdout, "buffer"):
    sys.stdout.buffer.write(output.encode("utf-8"))
    sys.stdout.buffer.flush()
else:
    sys.stdout.write(output)
    sys.stdout.flush()
""")


# ---------------------------------------------------------------------------
# Office 自动化进程清理（Windows）
# ---------------------------------------------------------------------------

def _automation_office_pids() -> set[int]:
    """返回当前所有 Office 自动化进程的 PID 集合（Windows only）。"""
    if os.name != "nt":
        return set()
    cmd = (
        "Get-CimInstance Win32_Process "
        "-Filter \"Name='WINWORD.EXE' OR Name='EXCEL.EXE' "
        "OR Name='POWERPNT.EXE'\" "
        "| Where-Object CommandLine -Like '* /Automation *' "
        "| Select-Object -ExpandProperty ProcessId"
    )
    try:
        proc = subprocess.run(
            ["powershell", "-NoProfile", "-Command", cmd],
            capture_output=True, timeout=10,
        )
        stdout = proc.stdout.decode("utf-8", errors="ignore")
        return {int(x) for x in stdout.splitlines() if x.strip().isdigit()}
    except Exception:
        return set()



def _kill_pids(pids: set[int]) -> None:
    for pid in sorted(pids):
        try:
            subprocess.run(
                ["taskkill", "/PID", str(pid), "/T", "/F"],
                capture_output=True, timeout=5,
            )
        except Exception:
            pass


# ---------------------------------------------------------------------------
# 核心执行函数
# ---------------------------------------------------------------------------

def _run_once(
    verifier_path: Path,
    document_dir: Path,
    timeout: int,
    com_enabled: bool,
) -> tuple[Optional[dict], Optional[str]]:

    """
    执行一次子进程，返回 (result_dict, error_str)。
    timeout 到期后返回 (None, "timeout")。
    """
    env = {
        **os.environ,
        "PYTHONIOENCODING": "utf-8",
        "OFFICEVAL_COM_ENABLED": "1" if com_enabled else "0",
    }

    try:
        proc = subprocess.run(
            [sys.executable, "-c", _WORKER_SCRIPT,
             str(verifier_path), str(document_dir)],
            capture_output=True,
            timeout=timeout,
            env=env,
        )
    except subprocess.TimeoutExpired:
        return None, "timeout"

    stdout = proc.stdout.decode("utf-8", errors="replace").strip()
    stderr = proc.stderr.decode("utf-8", errors="replace")

    if proc.returncode != 0 and not stdout:
        return None, f"子进程退出码 {proc.returncode}\nstderr: {stderr[:500]}"

    # 从 stdout 提取最后一个合法 JSON 对象（verifier 调试输出可能混入其他内容）
    result = _parse_last_json(stdout)
    if result is None:
        return None, f"无法从 stdout 解析 JSON\nstdout: {stdout[:300]}\nstderr: {stderr[:300]}"

    return result, None


def _parse_last_json(text: str) -> Optional[dict]:
    """从文本中提取最后一个 JSON 对象。"""
    dec = json.JSONDecoder()
    keys = {"id", "file_name", "status", "error", "dim1_pass",
            "dim1_reason", "dim2_items", "total_score", "max_score"}
    found: list[dict] = []
    for i, ch in enumerate(text):
        if ch != "{":
            continue
        try:
            val, _ = dec.raw_decode(text[i:])
        except json.JSONDecodeError:
            continue
        if isinstance(val, dict) and keys.issubset(val):
            found.append(val)
    return found[-1] if found else None


def _is_com_retry_error(result: Optional[dict]) -> bool:
    """判断是否为可重试的 COM 瞬时错误。"""
    if result is None:
        return False
    if result.get("status") != "error":
        return False
    err = str(result.get("error") or "")
    return any(kw in err for kw in cfg.COM_RETRY_KEYWORDS)


# ---------------------------------------------------------------------------
# 公开接口
# ---------------------------------------------------------------------------

def run_verifier(
    id_: str,
    verifier_path: Path,
    document_dir: Path,
    started_at: str,
    timeout_seconds: int | None = None,
    com_enabled: bool = True,
) -> dict:


    """
    执行单个 verifier 并返回系统级结果字典。

    - 普通 verifier：执行一次，超时或失败直接返回错误结果。
    - COM verifier：对瞬时 COM 错误执行受控重试。
    """

    timeout = cfg.task_timeout_seconds(id_, timeout_seconds)

    if not com_enabled and id_ in cfg.COM_REQUIRED_VERIFIER_IDS:
        return build_system_result(
            id_, None, "skipped",
            "Office COM 已禁用或当前平台不支持，跳过该 verifier",
            0.0, started_at,
        )

    # 批次级 com_enabled 只表示平台允许 COM；具体 verifier 还必须是硬依赖项。
    # 存在普通路径的脚本始终收到 OFFICEVAL_COM_ENABLED=0。
    verifier_com_enabled = com_enabled and id_ in cfg.COM_REQUIRED_VERIFIER_IDS
    is_com = verifier_com_enabled

    t0 = time.monotonic()
    before_pids = _automation_office_pids() if is_com else set()

    result, error = _run_once(
        verifier_path, document_dir, timeout, verifier_com_enabled,
    )


    elapsed = round(time.monotonic() - t0, 2)

    if is_com:
        new_pids = _automation_office_pids() - before_pids
        _kill_pids(new_pids)

    retried = False
    if error == "timeout":
        return build_system_result(id_, None, "timeout",
                                   f"超过 {timeout} 秒时间限制",
                                   elapsed, started_at)

    max_attempts = max(1, getattr(cfg, "COM_RETRY_MAX_ATTEMPTS", 1)) if is_com else 1
    attempt = 1
    while attempt < max_attempts and _is_com_retry_error(result):
        # 受控重试：清理、退避、重新执行。
        retried = True
        attempt += 1
        time.sleep(cfg.COM_RETRY_DELAY_SECONDS * (attempt - 1))
        before_retry_pids = _automation_office_pids()
        result2, error2 = _run_once(
            verifier_path, document_dir, timeout, verifier_com_enabled,
        )


        elapsed = round(time.monotonic() - t0, 2)
        new_retry_pids = _automation_office_pids() - before_retry_pids
        _kill_pids(new_retry_pids)

        if error2 == "timeout":
            return build_system_result(id_, None, "timeout",
                                       f"第 {attempt} 次尝试超过 {timeout} 秒时间限制",
                                       elapsed, started_at, retried=True)
        result, error = result2, error2

    if error:
        return build_system_result(id_, None, "error", error,
                                   elapsed, started_at, retried=retried)

    return build_system_result(id_, result, result.get("status", "error"),

                               result.get("error"), elapsed, started_at,
                               retried=retried)
