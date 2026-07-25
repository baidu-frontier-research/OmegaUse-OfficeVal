# -*- coding: utf-8 -*-
"""
core/config.py — 全局配置。

所有路径、并发参数、超时、文件限制和各编号输入要求均在此声明。
系统启动时由 batch_runner 或 __main__ 调用 validate_config() 一次性校验；
配置非法时明确报错，不静默回退。
"""
from __future__ import annotations

import os
import pathlib
import re

# ---------------------------------------------------------------------------
# 项目根目录（以本文件所在的 core/ 上一级为准）
# ---------------------------------------------------------------------------
PROJECT_ROOT: pathlib.Path = pathlib.Path(__file__).resolve().parent.parent

# ---------------------------------------------------------------------------
# 各功能目录
# ---------------------------------------------------------------------------
VERIFIERS_DIR:   pathlib.Path = PROJECT_ROOT / "verifiers"
SUBMISSIONS_DIR: pathlib.Path = PROJECT_ROOT / "submissions"
WORKSPACES_DIR:  pathlib.Path = PROJECT_ROOT / "workspaces"
RESULTS_DIR:     pathlib.Path = PROJECT_ROOT / "results"

# ---------------------------------------------------------------------------
# 并发配置
# ---------------------------------------------------------------------------
# 总 worker 槽位数（含 COM 槽位）
MAX_WORKERS: int = 4
# Office COM verifier 专用串行通道槽位数（≥1；多余为非 COM 槽位）
MAX_COM_WORKERS: int = 1
# auto：Windows 启用、macOS/Linux 跳过；enabled：强制启用；disabled：强制跳过。
COM_MODE_CHOICES: tuple[str, ...] = ("auto", "enabled", "disabled")
DEFAULT_COM_MODE: str = "auto"


# ---------------------------------------------------------------------------
# 任务超时（秒）
# ---------------------------------------------------------------------------
DEFAULT_TASK_TIMEOUT_SECONDS: int = 600
# 编号级覆盖；key 为三位字符串编号。默认不做编号级覆盖，所有任务统一使用全局超时。
TASK_TIMEOUT_OVERRIDES: dict[str, int] = {}


# ---------------------------------------------------------------------------
# 压缩包安全限制
# ---------------------------------------------------------------------------
MAX_ZIP_SIZE_BYTES:        int = 2 * 1024 ** 3   # 2 GB
MAX_SINGLE_FILE_BYTES:     int = 500 * 1024 ** 2  # 500 MB
MAX_ZIP_MEMBER_COUNT:      int = 20_000
MAX_UNCOMPRESSED_BYTES:    int = 5 * 1024 ** 3    # 5 GB
MAX_COMPRESSION_RATIO:     float = 100.0            # 解压/压缩 > 此值视为炸弹

# ---------------------------------------------------------------------------
# 文件类型白名单（小写扩展名）
# ---------------------------------------------------------------------------
ALLOWED_DOC_EXTENSIONS: frozenset[str] = frozenset({
    ".docx",          # Word
    ".xlsx", ".xlsm",  # Excel
    ".pptx",          # PowerPoint
    ".pdf",           # PDF
})

# ---------------------------------------------------------------------------
# 编号命名规则
# ---------------------------------------------------------------------------
# 子文件夹前缀（压缩包内子目录和 workspaces 内子目录均采用此前缀）
DIR_PREFIX  = "officeval_"
# verifier 文件名模式
VERIFIER_SUFFIX = "_verifier.py"

# 合法编号集合（001..100 三位字符串）
VALID_IDS: frozenset[str] = frozenset(f"{i:03d}" for i in range(1, 101))

def dir_name(id_: str) -> str:
    """返回给定编号对应的目录名，如 'officeval_042'。"""
    return f"{DIR_PREFIX}{id_}"

def verifier_filename(id_: str) -> str:
    """返回给定编号对应的 verifier 文件名，如 'officeval_042_verifier.py'。"""
    return f"{DIR_PREFIX}{id_}{VERIFIER_SUFFIX}"

def id_from_dir_name(name: str) -> str | None:
    """从目录名提取三位编号字符串；不合法则返回 None。"""
    if not name.startswith(DIR_PREFIX):
        return None
    suffix = name[len(DIR_PREFIX):]
    if re.fullmatch(r"\d{3}", suffix) and suffix in VALID_IDS:
        return suffix
    return None

def id_from_verifier_filename(name: str) -> str | None:
    """从 verifier 文件名提取三位编号；不合法则返回 None。"""
    pattern = re.fullmatch(
        r"officeval_(\d{3})_verifier\.py", name, re.IGNORECASE
    )
    if pattern and pattern.group(1) in VALID_IDS:
        return pattern.group(1)
    return None


def task_timeout_seconds(id_: str, default_timeout_seconds: int | None = None) -> int:
    """返回编号任务的超时秒数。"""
    base_timeout = DEFAULT_TASK_TIMEOUT_SECONDS if default_timeout_seconds is None else default_timeout_seconds
    return TASK_TIMEOUT_OVERRIDES.get(id_, base_timeout)


def resolve_com_enabled(com_mode: str | None = None) -> bool:
    """根据用户模式和当前操作系统决定是否执行 Office COM verifier。"""
    mode = DEFAULT_COM_MODE if com_mode is None else com_mode
    if mode not in COM_MODE_CHOICES:
        raise ValueError(f"com_mode 必须是 {COM_MODE_CHOICES} 之一，当前值：{mode!r}")
    if mode == "disabled":
        return False
    if mode == "auto":
        return os.name == "nt"
    if os.name != "nt":
        raise RuntimeError("Office COM 只能在 Windows 上启用；请使用 --com-mode auto 或 disabled")
    return True


# ---------------------------------------------------------------------------
# 哪些编号的 verifier 使用 Office COM（需走串行通道）

# ---------------------------------------------------------------------------

# 强依赖 COM：禁用 COM 时整项跳过；启用时进入 COM 串行通道。
COM_REQUIRED_VERIFIER_IDS: frozenset[str] = frozenset({
    "011", "023", "039", "081",
})


# 存在普通解析路径的 COM 相关 verifier：在所有平台强制使用普通模式。
# 即使批次启用了 COM，worker 也会向这些脚本传入 OFFICEVAL_COM_ENABLED=0。
COM_FORCED_NORMAL_VERIFIER_IDS: frozenset[str] = frozenset({
    "002", "004", "010", "019", "024",
    "026", "031", "035", "037", "092",
})


# 只有硬依赖 COM 的编号进入 COM 串行通道。
COM_VERIFIER_IDS: frozenset[str] = COM_REQUIRED_VERIFIER_IDS




# ---------------------------------------------------------------------------
# COM 重试关键字（出现时允许受控重试）
# ---------------------------------------------------------------------------

COM_RETRY_KEYWORDS: tuple[str, ...] = (
    "被呼叫方拒绝接收呼叫",
    "对象已被删除",
    "RPC_E_CALL_REJECTED",
    "RPC_E_SERVERCALL_RETRYLATER",
    "-2147418111",
)
COM_RETRY_DELAY_SECONDS: float = 2.0
COM_RETRY_MAX_ATTEMPTS: int = 3


# ---------------------------------------------------------------------------
# job_id 生成参数
# ---------------------------------------------------------------------------
JOB_ID_MAX_BASE_LENGTH: int = 60   # base 部分（来自文件名）最大字符数
JOB_ID_RANDOM_CHARS:    int = 4    # 随机短串长度

# Windows 保留设备名（不区分大小写）
WINDOWS_RESERVED_NAMES: frozenset[str] = frozenset({
    "CON", "PRN", "AUX", "NUL",
    "COM1", "COM2", "COM3", "COM4", "COM5", "COM6", "COM7", "COM8", "COM9",
    "LPT1", "LPT2", "LPT3", "LPT4", "LPT5", "LPT6", "LPT7", "LPT8", "LPT9",
})

# ---------------------------------------------------------------------------
# 配置校验
# ---------------------------------------------------------------------------
def validate_config() -> None:
    """
    系统启动时调用，对配置值进行一次性检查。
    发现非法值时抛出 ValueError，附带描述信息。
    """
    if not isinstance(MAX_WORKERS, int) or not (1 <= MAX_WORKERS <= 100):
        raise ValueError(
            f"MAX_WORKERS 必须是 1..100 的整数，当前值：{MAX_WORKERS!r}"
        )
    if not isinstance(MAX_COM_WORKERS, int) or not (1 <= MAX_COM_WORKERS <= MAX_WORKERS):
        raise ValueError(
            f"MAX_COM_WORKERS 必须是 1..MAX_WORKERS 的整数，"
            f"当前值：{MAX_COM_WORKERS!r}，MAX_WORKERS={MAX_WORKERS}"
        )
    if DEFAULT_COM_MODE not in COM_MODE_CHOICES:
        raise ValueError(
            f"DEFAULT_COM_MODE 必须是 {COM_MODE_CHOICES} 之一，"
            f"当前值：{DEFAULT_COM_MODE!r}"
        )
    overlap = COM_REQUIRED_VERIFIER_IDS & COM_FORCED_NORMAL_VERIFIER_IDS
    if overlap:
        raise ValueError(f"COM 强依赖与强制普通集合不能重叠：{sorted(overlap)}")
    invalid_com_ids = (
        COM_VERIFIER_IDS | COM_FORCED_NORMAL_VERIFIER_IDS
    ) - VALID_IDS
    if invalid_com_ids:
        raise ValueError(f"COM 相关 verifier 包含非法编号：{sorted(invalid_com_ids)}")

    if not isinstance(DEFAULT_TASK_TIMEOUT_SECONDS, int) or DEFAULT_TASK_TIMEOUT_SECONDS < 10:


        raise ValueError(
            f"DEFAULT_TASK_TIMEOUT_SECONDS 必须 ≥ 10，当前值：{DEFAULT_TASK_TIMEOUT_SECONDS!r}"
        )
    for id_, timeout in TASK_TIMEOUT_OVERRIDES.items():
        if id_ not in VALID_IDS:
            raise ValueError(f"TASK_TIMEOUT_OVERRIDES 包含非法编号：{id_!r}")
        if not isinstance(timeout, int) or timeout < 10:
            raise ValueError(
                f"TASK_TIMEOUT_OVERRIDES[{id_!r}] 必须 ≥ 10，当前值：{timeout!r}"
            )
    if not VERIFIERS_DIR.is_dir():
        raise ValueError(f"verifiers/ 目录不存在：{VERIFIERS_DIR}")
