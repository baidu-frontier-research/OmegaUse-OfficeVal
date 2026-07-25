# -*- coding: utf-8 -*-
"""
core/submission_service.py

统一处理压缩包提交：路径解析、job_id 生成、SHA-256 校验、原子归档。
"""
from __future__ import annotations

import hashlib
import json
import os
import pathlib
import random
import re
import shutil
import string
import time
from datetime import datetime, timezone, timedelta

from . import config as cfg

# 本地时区偏移（CST UTC+8）
_TZ_CST = timezone(timedelta(hours=8))


# ---------------------------------------------------------------------------
# 路径解析
# ---------------------------------------------------------------------------

def resolve_package_path(user_input: str) -> pathlib.Path:
    """
    将用户输入解析为 ZIP 文件的绝对路径。

    解析顺序：
    1. 剥离首尾空白和成对引号。
    2. 若为绝对路径，直接解析。
    3. 若为相对路径，相对于当前工作目录解析。
    4. 若仍未找到，在 submissions/ 目录下按文件名查找。
    5. 文件不存在、不是普通文件或扩展名不是 .zip 时明确报错。
    """
    text = user_input.strip()
    # 剥离成对引号（Windows 路径拖入终端会附带引号）
    if len(text) >= 2 and text[0] in ('"', "'") and text[-1] == text[0]:
        text = text[1:-1].strip()

    candidates: list[pathlib.Path] = []

    p = pathlib.Path(text)
    if p.is_absolute():
        candidates.append(p)
    else:
        candidates.append(pathlib.Path.cwd() / p)
        candidates.append(cfg.SUBMISSIONS_DIR / text)

    for candidate in candidates:
        if candidate.exists():
            if not candidate.is_file():
                raise ValueError(f"路径存在但不是普通文件：{candidate}")
            if candidate.suffix.lower() != ".zip":
                raise ValueError(f"文件扩展名不是 .zip：{candidate}")
            return candidate.resolve()

    raise FileNotFoundError(
        f"找不到压缩包文件：{text!r}\n"
        f"  已尝试：{', '.join(str(c) for c in candidates)}"
    )


# ---------------------------------------------------------------------------
# job_id 生成
# ---------------------------------------------------------------------------

def _sanitize_base(name: str) -> str:
    """
    将文件名清理为适合用作目录名的安全字符串：
    - 移除扩展名
    - 替换路径分隔符和 Windows 非法字符为 _
    - 规避保留设备名
    - 限制长度
    """
    stem = pathlib.Path(name).stem
    # 替换非字母数字、连字符、下划线的字符
    safe = re.sub(r'[^\w\-]', '_', stem, flags=re.UNICODE)
    # 折叠连续下划线
    safe = re.sub(r'_+', '_', safe).strip('_')
    if not safe:
        safe = "package"
    # 规避 Windows 保留设备名（大小写不敏感）
    if safe.upper().split('.')[0] in cfg.WINDOWS_RESERVED_NAMES:
        safe = f"pkg_{safe}"
    # 限制 base 长度
    safe = safe[:cfg.JOB_ID_MAX_BASE_LENGTH]
    return safe


def generate_job_id(original_name: str) -> str:
    """生成唯一任务 ID：<安全文件名>_<YYYYMMDD_HHMMSS>_<随机短串>"""
    base = _sanitize_base(original_name)
    ts = datetime.now(_TZ_CST).strftime("%Y%m%d_%H%M%S")
    rand = ''.join(
        random.choices(string.ascii_lowercase + string.digits, k=cfg.JOB_ID_RANDOM_CHARS)
    )
    return f"{base}_{ts}_{rand}"


# ---------------------------------------------------------------------------
# SHA-256 / 文件复制
# ---------------------------------------------------------------------------

def _sha256(path: pathlib.Path) -> tuple[str, int]:
    """返回 (hex_digest, size_bytes)。"""
    h = hashlib.sha256()
    size = 0
    with path.open("rb") as fh:
        while chunk := fh.read(65536):
            h.update(chunk)
            size += len(chunk)
    return h.hexdigest(), size


def _atomic_copy(src: pathlib.Path, dest_dir: pathlib.Path) -> tuple[pathlib.Path, str, int]:
    """
    安全地将 src 复制为 dest_dir/original.zip：
    1. 先写到 original.zip.part
    2. 写完后核对 SHA-256 和大小
    3. 原子重命名为 original.zip

    返回 (final_path, sha256_hex, size_bytes)。
    """
    src_sha, src_size = _sha256(src)

    dest_dir.mkdir(parents=True, exist_ok=True)
    part_path = dest_dir / "original.zip.part"
    final_path = dest_dir / "original.zip"

    # 防止残留 part 文件干扰
    if part_path.exists():
        part_path.unlink()

    shutil.copyfile(src, part_path)

    # 核对
    copy_sha, copy_size = _sha256(part_path)
    if copy_sha != src_sha or copy_size != src_size:
        part_path.unlink(missing_ok=True)
        raise IOError(
            f"复制后校验失败：源 {src_sha}/{src_size} vs 副本 {copy_sha}/{copy_size}"
        )

    part_path.rename(final_path)
    return final_path, src_sha, src_size


# ---------------------------------------------------------------------------
# 主入口
# ---------------------------------------------------------------------------

def submit_package(package_path: str) -> str:
    """
    将用户提供的压缩包路径提交为一个新任务，返回 job_id。

    副作用：
    - 在 submissions/<job_id>/ 创建 original.zip 和 source.json
    - 在 results/<job_id>/ 创建 job.json（状态 submitted）
    - workspaces/<job_id>/ 此时尚不创建（由 extractor 负责）
    """
    src = resolve_package_path(package_path)
    job_id = generate_job_id(src.name)

    submission_dir = cfg.SUBMISSIONS_DIR / job_id
    result_dir = cfg.RESULTS_DIR / job_id

    # 归档原始 ZIP
    archived, sha256, size = _atomic_copy(src, submission_dir)

    # 写 source.json
    submitted_at = datetime.now(_TZ_CST).isoformat()
    source_info: dict = {
        "original_name": src.name,
        "sha256": sha256,
        "size_bytes": size,
        "submitted_at": submitted_at,
    }
    _write_json(submission_dir / "source.json", source_info)

    # 写 job.json（初始状态）
    result_dir.mkdir(parents=True, exist_ok=True)
    job_info: dict = {
        "job_id": job_id,
        "status": "submitted",
        "submission_dir": str(submission_dir),
        "workspace_dir": str(cfg.WORKSPACES_DIR / job_id),
        "result_dir": str(result_dir),
        "archive_path": str(archived),
        "sha256": sha256,
        "size_bytes": size,
        "submitted_at": submitted_at,
        "confirmed_at": None,
        "started_at": None,
        "finished_at": None,
    }
    _write_json(result_dir / "job.json", job_info)

    return job_id


# ---------------------------------------------------------------------------
# 内部工具
# ---------------------------------------------------------------------------

def _write_json(path: pathlib.Path, data: dict) -> None:
    """原子写入 JSON（先写 .tmp，再重命名）。"""
    tmp = path.with_suffix(path.suffix + ".tmp")
    tmp.write_text(
        json.dumps(data, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    tmp.replace(path)


def load_job(job_id: str) -> dict:
    """读取并返回 job.json 字典；找不到时抛 FileNotFoundError。"""
    p = cfg.RESULTS_DIR / job_id / "job.json"
    if not p.is_file():
        raise FileNotFoundError(f"找不到 job.json：{p}")
    return json.loads(p.read_text(encoding="utf-8"))


def update_job_status(job_id: str, status: str, **extra_fields) -> None:
    """更新 job.json 中的 status 字段及任意附加字段（原子写入）。"""
    p = cfg.RESULTS_DIR / job_id / "job.json"
    data = json.loads(p.read_text(encoding="utf-8"))
    data["status"] = status
    data.update(extra_fields)
    _write_json(p, data)
