#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# pyright: reportAny=false, reportExplicitAny=false, reportUnknownVariableType=false, reportUnknownMemberType=false, reportUnknownArgumentType=false
"""三列蛇形自动跳转录入表自动评分脚本。

评分对象是学生交付的 .xlsm / .xlsx 文件。脚本会先做可用性门槛检查；门槛不通过时
直接 0 分并跳过完成度评分。门槛通过后，使用 Excel COM 自动化实际写入单元格，
检查宏触发后的活动单元格位置，并逐项累计完成度分数。
"""

from __future__ import annotations

import hashlib
import json
import multiprocessing as mp
import posixpath
import queue
import re
import shutil
import subprocess
import sys
import tempfile
import time
import traceback
import zipfile
from dataclasses import asdict, dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Any
from xml.etree import ElementTree as ET

MAX_SCORE = 19
DEFAULT_SHEET = "蛇形录入"
DEFAULT_RANGE = "B4:D23"
TEMPLATE_NAME = "三列蛇形自动跳转录入表_模板.xlsx"
DEFAULT_CANDIDATE_NAMES = (
    "三列蛇形自动跳转录入表_模板.xlsm",
    "三列蛇形自动跳转录入表_模板.xlsx",
)

NS = {
    "main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "rel": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "pkgrel": "http://schemas.openxmlformats.org/package/2006/relationships",
    "xdr": "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
}


@dataclass
class EvaluatorConfig:
    sheet_name: str = DEFAULT_SHEET
    entry_range: str = DEFAULT_RANGE
    timeout_seconds: int = 120
    visible: bool = False
    keep_temp: bool = False


@dataclass
class ScoreItem:
    id: str
    name: str
    points: int
    awarded: int
    passed: bool
    evidence: list[str] = field(default_factory=list)

    def to_dict(self) -> dict[str, Any]:
        return asdict(self)


class ValueGenerator:
    def __init__(self) -> None:
        self.value = 900000

    def next(self) -> int:
        self.value += 1
        return self.value


# ---------------------------------------------------------------------------
# Address helpers


def col_to_num(col: str) -> int:
    total = 0
    for ch in col.upper():
        if not ("A" <= ch <= "Z"):
            raise ValueError(f"非法列名: {col}")
        total = total * 26 + ord(ch) - ord("A") + 1
    return total


def num_to_col(num: int) -> str:
    if num < 1:
        raise ValueError(f"非法列号: {num}")
    out = []
    while num:
        num, rem = divmod(num - 1, 26)
        out.append(chr(ord("A") + rem))
    return "".join(reversed(out))


def split_addr(addr: str) -> tuple[int, int]:
    cleaned = addr.replace("$", "").split("!")[-1]
    match = re.fullmatch(r"([A-Za-z]+)(\d+)", cleaned)
    if not match:
        raise ValueError(f"非法单元格地址: {addr}")
    return int(match.group(2)), col_to_num(match.group(1))


def make_addr(row: int, col: int) -> str:
    return f"{num_to_col(col)}{row}"


def parse_range(range_addr: str) -> tuple[int, int, int, int]:
    parts = range_addr.replace("$", "").split(":")
    if len(parts) != 2:
        raise ValueError(f"录入区域必须是 A1:B2 形式: {range_addr}")
    row1, col1 = split_addr(parts[0])
    row2, col2 = split_addr(parts[1])
    return min(row1, row2), min(col1, col2), max(row1, row2), max(col1, col2)


def range_overlaps(a: tuple[int, int, int, int], b: tuple[int, int, int, int]) -> bool:
    ar1, ac1, ar2, ac2 = a
    br1, bc1, br2, bc2 = b
    return not (ar2 < br1 or br2 < ar1 or ac2 < bc1 or bc2 < ac1)


def normalize_addr(addr: str) -> str:
    row, col = split_addr(addr)
    return make_addr(row, col)


def build_expected_path(entry_range: str) -> list[str]:
    row1, col1, row2, col2 = parse_range(entry_range)
    path: list[str] = []
    for row_offset, row in enumerate(range(row1, row2 + 1)):
        cols = range(col1, col2 + 1) if row_offset % 2 == 0 else range(col2, col1 - 1, -1)
        for col in cols:
            path.append(make_addr(row, col))
    return path


def expected_next_map(path: list[str]) -> dict[str, str]:
    return {addr: path[(idx + 1) % len(path)] for idx, addr in enumerate(path)}


def addresses_in_range(entry_range: str) -> list[str]:
    row1, col1, row2, col2 = parse_range(entry_range)
    for row in range(row1, row2 + 1):
        for col in range(col1, col2 + 1):
            yield make_addr(row, col)


# ---------------------------------------------------------------------------
# Static OOXML preflight


def file_sha256(path: Path) -> str | None:
    if not path.exists() or not path.is_file():
        return None
    h = hashlib.sha256()
    with path.open("rb") as f:
        for chunk in iter(lambda: f.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest()


def read_zip_xml(zf: zipfile.ZipFile, name: str) -> ET.Element | None:
    try:
        return ET.fromstring(zf.read(name))
    except Exception:
        return None


def workbook_sheet_paths(zf: zipfile.ZipFile) -> dict[str, str]:
    workbook = read_zip_xml(zf, "xl/workbook.xml")
    rels = read_zip_xml(zf, "xl/_rels/workbook.xml.rels")
    if workbook is None or rels is None:
        return {}

    rel_map: dict[str, str] = {}
    for rel in rels.findall("pkgrel:Relationship", NS):
        rid = rel.attrib.get("Id")
        target = rel.attrib.get("Target", "")
        if rid and target:
            target = target.lstrip("/")
            target_path = target if target.startswith("xl/") else posixpath.normpath("xl/" + target)
            rel_map[rid] = target_path

    result: dict[str, str] = {}
    for sheet in workbook.findall("main:sheets/main:sheet", NS):
        name = sheet.attrib.get("name")
        rid = sheet.attrib.get(f"{{{NS['rel']}}}id")
        if name and rid in rel_map:
            result[name] = rel_map[rid]
    return result


def sheet_drawing_paths(zf: zipfile.ZipFile, sheet_path: str) -> list[str]:
    """列出工作表引用的 drawing xml 路径。用于静态检查图片/形状是否覆盖录入区域。"""
    sheet_dir = posixpath.dirname(sheet_path)
    sheet_base = posixpath.basename(sheet_path)
    rels_path = f"{sheet_dir}/_rels/{sheet_base}.rels"
    rels = read_zip_xml(zf, rels_path)
    if rels is None:
        return []
    drawing_paths: list[str] = []
    for rel in rels.findall("pkgrel:Relationship", NS):
        type_ = rel.attrib.get("Type", "")
        target = rel.attrib.get("Target", "")
        if not target:
            continue
        if "drawing" not in type_.lower() and "oleObject" not in type_:
            continue
        if target.startswith("/"):
            resolved = target.lstrip("/")
        else:
            resolved = posixpath.normpath(posixpath.join(sheet_dir, target))
        drawing_paths.append(resolved)
    return drawing_paths


def drawing_overlaps(zf: zipfile.ZipFile, drawing_path: str, input_box: tuple[int, int, int, int]) -> list[str]:
    """解析 drawing xml，返回与录入区域重叠的锚点地址列表。"""
    hits: list[str] = []
    drawing_xml = read_zip_xml(zf, drawing_path)
    if drawing_xml is None:
        return hits

    def read_anchor(elem: ET.Element) -> tuple[int, int] | None:
        col_elem = elem.find("xdr:col", NS)
        row_elem = elem.find("xdr:row", NS)
        if col_elem is None or row_elem is None or col_elem.text is None or row_elem.text is None:
            return None
        try:
            # OOXML 的行列锚点是 0 基，转换为 1 基以匹配 A1 地址。
            return int(row_elem.text) + 1, int(col_elem.text) + 1
        except ValueError:
            return None

    for anchor_tag in ("xdr:twoCellAnchor", "xdr:oneCellAnchor"):
        for anchor in drawing_xml.findall(anchor_tag, NS):
            frm = anchor.find("xdr:from", NS)
            if frm is None:
                continue
            frm_pos = read_anchor(frm)
            if frm_pos is None:
                continue
            to = anchor.find("xdr:to", NS)
            to_pos = read_anchor(to) if to is not None else frm_pos
            if to_pos is None:
                to_pos = frm_pos
            r1, c1 = frm_pos
            r2, c2 = to_pos
            box = (min(r1, r2), min(c1, c2), max(r1, r2), max(c1, c2))
            if range_overlaps(box, input_box):
                hits.append(f"{make_addr(box[0], box[1])}:{make_addr(box[2], box[3])}")
    return hits


def try_openpyxl_open(candidate_path: Path) -> tuple[bool, str | None]:
    """用 openpyxl 尝试打开工作簿，验证"文件可正常打开"。

    返回 (是否可用于验证, 错误消息)。若 openpyxl 未安装则返回 (False, None)，
    调用方应视为"跳过"，不作为失败依据；打开异常则返回 (True, 错误描述)。
    """
    try:
        import openpyxl  # type: ignore
    except ImportError:
        return False, None
    keep_vba = candidate_path.suffix.lower() == ".xlsm"
    try:
        wb = openpyxl.load_workbook(
            filename=str(candidate_path),
            read_only=True,
            keep_vba=keep_vba,
            data_only=False,
        )
        try:
            _ = list(wb.sheetnames)
        finally:
            wb.close()
        return True, None
    except Exception as exc:
        return True, f"openpyxl 无法正常打开工作簿：{type(exc).__name__}: {exc}"


def static_preflight(candidate_path: Path, config: EvaluatorConfig) -> dict[str, Any]:
    reasons: list[str] = []
    evidence: dict[str, Any] = {
        "path": str(candidate_path),
        "suffix": candidate_path.suffix,
        "entry_range": config.entry_range,
    }

    if not candidate_path.exists():
        reasons.append("文件不存在")
        return {"passed": False, "reasons": reasons, "evidence": evidence}
    if not candidate_path.is_file():
        reasons.append("路径不是文件")
        return {"passed": False, "reasons": reasons, "evidence": evidence}
    if candidate_path.suffix.lower() not in (".xlsm", ".xlsx"):
        reasons.append("交付文件格式不是 Excel 工作簿（.xlsm/.xlsx）")

    if not zipfile.is_zipfile(candidate_path):
        reasons.append("文件不是可解析的 Office Open XML 工作簿")
        return {"passed": False, "reasons": reasons, "evidence": evidence}

    try:
        with zipfile.ZipFile(candidate_path) as zf:
            names = set(zf.namelist())
            evidence["has_vba_project"] = "xl/vbaProject.bin" in names

            workbook = read_zip_xml(zf, "xl/workbook.xml")
            if workbook is None:
                reasons.append("无法解析 xl/workbook.xml")
            else:
                if workbook.find("main:workbookProtection", NS) is not None:
                    reasons.append("工作簿结构存在保护设置")

            sheet_paths = workbook_sheet_paths(zf)
            evidence["sheets"] = list(sheet_paths.keys())
            sheet_path = sheet_paths.get(config.sheet_name)
            if not sheet_path:
                reasons.append(f"缺少工作表：{config.sheet_name}")
            else:
                evidence["sheet_xml"] = sheet_path
                sheet_xml = read_zip_xml(zf, sheet_path)
                if sheet_xml is None:
                    reasons.append(f"无法解析 {config.sheet_name} 的 XML")
                else:
                    if sheet_xml.find("main:sheetProtection", NS) is not None:
                        reasons.append(f"工作表 {config.sheet_name} 存在保护设置")
                    input_box = parse_range(config.entry_range)
                    overlapping_merges: list[str] = []
                    for merge_cell in sheet_xml.findall("main:mergeCells/main:mergeCell", NS):
                        ref = merge_cell.attrib.get("ref")
                        if ref and ":" in ref and range_overlaps(parse_range(ref), input_box):
                            overlapping_merges.append(ref)
                    evidence["overlapping_merges"] = overlapping_merges
                    if overlapping_merges:
                        reasons.append(f"录入区域与合并单元格重叠：{', '.join(overlapping_merges)}")
    except Exception as exc:
        reasons.append(f"静态预检异常：{exc}")

    # 单独走 openpyxl 兜底"文件可正常打开"这一门槛，避免仅靠 zipfile 通过就下结论。
    can_verify, open_error = try_openpyxl_open(candidate_path)
    evidence["openpyxl_verified"] = can_verify
    if can_verify and open_error:
        reasons.append(open_error)

    return {"passed": not reasons, "reasons": reasons, "evidence": evidence}


# ---------------------------------------------------------------------------
# COM evaluation wrapper


def run_com_with_timeout(candidate_path: Path, config: EvaluatorConfig, original_hash: str | None) -> dict[str, Any]:
    ctx = mp.get_context("spawn")
    q: mp.Queue = ctx.Queue()
    proc = ctx.Process(target=com_worker, args=(str(candidate_path), asdict(config), original_hash, q))
    proc.start()

    excel_pid: int | None = None
    result: dict[str, Any | None] = None
    deadline = time.time() + config.timeout_seconds

    while time.time() < deadline:
        try:
            msg = q.get(timeout=0.2)
        except queue.Empty:
            msg = None

        if isinstance(msg, dict):
            if msg.get("type") == "pid":
                excel_pid = msg.get("pid")
            elif msg.get("type") == "result":
                result = msg.get("report")
                break

        if not proc.is_alive():
            break

    if result is None:
        # Drain any late result after process exit.
        try:
            while True:
                msg = q.get_nowait()
                if isinstance(msg, dict) and msg.get("type") == "result":
                    result = msg.get("report")
                    break
                if isinstance(msg, dict) and msg.get("type") == "pid":
                    excel_pid = msg.get("pid")
        except queue.Empty:
            pass

    if result is None and proc.is_alive():
        proc.terminate()
        if excel_pid:
            kill_process_tree(excel_pid)
        proc.join(timeout=5)
        return gate_failed_report(
            candidate_path,
            config,
            original_hash,
            ["Excel COM 评估超时，可能存在宏运行时错误、弹窗阻塞或无法正常保存/关闭"],
            {"excel_pid": excel_pid, "timeout_seconds": config.timeout_seconds},
        )

    proc.join(timeout=5)

    if result is None:
        return gate_failed_report(
            candidate_path,
            config,
            original_hash,
            ["Excel COM 评估子进程未返回结果"],
            {"process_exitcode": proc.exitcode, "excel_pid": excel_pid},
        )

    return result


def kill_process_tree(pid: int) -> None:
    try:
        subprocess.run(
            ["taskkill", "/PID", str(pid), "/T", "/F"],
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
            check=False,
        )
    except Exception:
        pass


def com_worker(candidate_path_str: str, config_dict: dict[str, Any], original_hash: str | None, q: mp.Queue) -> None:
    config = EvaluatorConfig(**config_dict)
    candidate_path = Path(candidate_path_str)
    temp_dir: str | None = None
    excel = None
    workbook = None

    try:
        try:
            import pythoncom  # type: ignore
            import win32com.client  # type: ignore
            import win32process  # type: ignore
        except Exception as exc:
            q.put({
                "type": "result",
                "report": gate_failed_report(
                    candidate_path,
                    config,
                    original_hash,
                    [f"无法导入 Excel COM 依赖 pywin32：{exc}"],
                    {},
                ),
            })
            return

        pythoncom.CoInitialize()
        temp_dir = tempfile.mkdtemp(prefix="snake_eval_")
        temp_path = Path(temp_dir) / candidate_path.name
        shutil.copy2(candidate_path, temp_path)

        excel = win32com.client.DispatchEx("Excel.Application")
        try:
            _, pid = win32process.GetWindowThreadProcessId(excel.Hwnd)
            q.put({"type": "pid", "pid": pid})
        except Exception:
            pass

        excel.Visible = bool(config.visible)
        excel.DisplayAlerts = False
        excel.EnableEvents = True
        excel.AskToUpdateLinks = False
        try:
            excel.AutomationSecurity = 1  # msoAutomationSecurityLow
        except Exception:
            pass

        workbook = excel.Workbooks.Open(
            str(temp_path),
            UpdateLinks=0,
            ReadOnly=False,
            IgnoreReadOnlyRecommended=True,
            AddToMru=False,
        )

        report = run_com_evaluation(excel, workbook, candidate_path, config, original_hash)
        q.put({"type": "result", "report": report})
    except Exception as exc:
        tb = traceback.format_exc(limit=12)
        q.put({
            "type": "result",
            "report": gate_failed_report(
                candidate_path,
                config,
                original_hash,
                [f"Excel COM 动态评估异常：{exc}"],
                {"traceback": tb},
            ),
        })
    finally:
        try:
            if workbook is not None:
                workbook.Close(SaveChanges=False)
        except Exception:
            pass
        try:
            if excel is not None:
                excel.Quit()
        except Exception:
            pass
        try:
            import pythoncom  # type: ignore
            pythoncom.CoUninitialize()
        except Exception:
            pass
        if temp_dir and not config.keep_temp:
            shutil.rmtree(temp_dir, ignore_errors=True)


# ---------------------------------------------------------------------------
# COM checks


def run_com_evaluation(excel: Any, workbook: Any, candidate_path: Path, config: EvaluatorConfig, original_hash: str | None) -> dict[str, Any]:
    dimension1 = run_dimension1_gate(excel, workbook, config)
    if not dimension1["passed"]:
        return build_report(candidate_path, config, original_hash, dimension1, [], skipped=True)

    dimension2 = run_dimension2_score(excel, workbook, config)
    return build_report(candidate_path, config, original_hash, dimension1, dimension2, skipped=False)


def get_sheet(workbook: Any, sheet_name: str) -> Any:
    try:
        return workbook.Worksheets(sheet_name)
    except Exception as exc:
        raise RuntimeError(f"缺少工作表：{sheet_name}") from exc


def com_address(obj: Any) -> str:
    """Return a COM object's address across Office/WPS COM variants.

    Some COM bindings expose Address as a callable method, while others expose
    it as a string property. The evaluator only needs a normalized A1 address.
    """
    address = getattr(obj, "Address")
    if callable(address):
        try:
            return str(address(False, False))
        except TypeError:
            return str(address())
    return str(address)


def run_dimension1_gate(excel: Any, workbook: Any, config: EvaluatorConfig) -> dict[str, Any]:
    reasons: list[str] = []
    evidence: dict[str, Any] = {}

    try:
        ws = get_sheet(workbook, config.sheet_name)
        evidence["sheet_found"] = True
    except Exception as exc:
        return {"passed": False, "reasons": [str(exc)], "evidence": evidence}

    try:
        evidence["workbook_protect_structure"] = bool(workbook.ProtectStructure)
        evidence["workbook_protect_windows"] = bool(workbook.ProtectWindows)
        if workbook.ProtectStructure or workbook.ProtectWindows:
            reasons.append("工作簿结构或窗口被保护")
    except Exception as exc:
        evidence["workbook_protection_check_error"] = str(exc)

    try:
        evidence["sheet_protected"] = bool(ws.ProtectContents)
        if ws.ProtectContents:
            reasons.append(f"工作表 {config.sheet_name} 被保护，可能无法编辑录入区域")
    except Exception as exc:
        evidence["sheet_protection_check_error"] = str(exc)

    try:
        rng = ws.Range(config.entry_range)
        evidence["entry_range_address"] = com_address(rng)
        if bool(rng.MergeCells):
            reasons.append(f"录入区域 {config.entry_range} 包含合并单元格")
    except Exception as exc:
        reasons.append(f"无法访问录入区域 {config.entry_range}：{exc}")
        return {"passed": False, "reasons": reasons, "evidence": evidence}

    object_hits: list[str] = []
    evidence["overlapping_objects"] = object_hits

    try:
        old_events = excel.EnableEvents
        excel.EnableEvents = False
        ws.Activate()
        ws.Range(config.entry_range).ClearContents()
        ws.Range("B4" if config.entry_range.upper() == DEFAULT_RANGE else config.entry_range.split(":")[0]).Value2 = 12345
        excel.EnableEvents = old_events
        evidence["editable_smoke"] = True
    except Exception as exc:
        try:
            excel.EnableEvents = True
        except Exception:
            pass
        reasons.append(f"录入区域无法写入数字：{exc}")

    return {"passed": not reasons, "reasons": reasons, "evidence": evidence}


def find_overlapping_objects(ws: Any, entry_range: str) -> list[str]:
    hits: list[str] = []
    input_box = parse_range(entry_range)

    def object_box(obj: Any) -> tuple[int, int, int, int | None]:
        try:
            tl = com_address(obj.TopLeftCell)
            br = com_address(obj.BottomRightCell)
            r1, c1 = split_addr(tl)
            r2, c2 = split_addr(br)
            return min(r1, r2), min(c1, c2), max(r1, r2), max(c1, c2)
        except Exception:
            return None

    try:
        for idx in range(1, int(ws.Shapes.Count) + 1):
            shape = ws.Shapes.Item(idx)
            box = object_box(shape)
            if box and range_overlaps(box, input_box):
                name = getattr(shape, "Name", f"Shape{idx}")
                hits.append(f"Shape:{name}@{make_addr(box[0], box[1])}:{make_addr(box[2], box[3])}")
    except Exception:
        pass

    try:
        ole_objects = ws.OLEObjects()
        for idx in range(1, int(ole_objects.Count) + 1):
            obj = ole_objects.Item(idx)
            box = object_box(obj)
            if box and range_overlaps(box, input_box):
                name = getattr(obj, "Name", f"OLEObject{idx}")
                hits.append(f"OLE:{name}@{make_addr(box[0], box[1])}:{make_addr(box[2], box[3])}")
    except Exception:
        pass

    return hits


def clear_entry_range(excel: Any, ws: Any, entry_range: str) -> None:
    old_events = excel.EnableEvents
    excel.EnableEvents = False
    try:
        ws.Activate()
        ws.Range(entry_range).ClearContents()
    finally:
        excel.EnableEvents = old_events


def set_value_without_events(excel: Any, ws: Any, addr: str, value: Any) -> None:
    old_events = excel.EnableEvents
    excel.EnableEvents = False
    try:
        ws.Range(addr).Value2 = value
    finally:
        excel.EnableEvents = old_events


def write_value_and_get_active(excel: Any, ws: Any, addr: str, value: Any) -> str:
    ws.Activate()
    target = ws.Range(addr)
    target.Select()
    excel.EnableEvents = True
    target.Value2 = value
    pump_excel_messages()
    return normalize_addr(com_address(excel.ActiveCell))


def type_value_press_enter_and_get_active(excel: Any, ws: Any, addr: str, value: Any) -> str:
    """写入单元格并读取宏处理后的活动单元格。

    不使用 SendKeys：SendKeys 依赖前台焦点，可能把测试数字输入到聊天窗口、终端
    或其他应用。这里通过 Excel/WPS COM 在工作簿内部写入数值，并保持事件开启，
    让 Worksheet_Change 等办公软件事件宏自行处理跳转。
    """
    ws.Activate()
    target = ws.Range(addr)
    target.Select()
    excel.EnableEvents = True
    target.Value2 = value
    pump_excel_messages()
    return normalize_addr(com_address(excel.ActiveCell))


def pump_excel_messages() -> None:
    try:
        import pythoncom  # type: ignore
        pythoncom.PumpWaitingMessages()
    except Exception:
        pass
    time.sleep(0.05)


def run_dimension2_score(excel: Any, workbook: Any, config: EvaluatorConfig) -> list[ScoreItem]:
    ws = get_sheet(workbook, config.sheet_name)
    path = build_expected_path(config.entry_range)
    next_map = expected_next_map(path)

    items: list[ScoreItem] = []
    item7: ScoreItem | None = None

    for func in (
        lambda: score_odd_rows_ltr(excel, ws, config),
        lambda: score_even_rows_rtl(excel, ws, config),
        lambda: score_complete_path(excel, ws, path, config),
        lambda: score_modify_existing(excel, ws, next_map, config),
    ):
        items.append(score_guard(func))

    # 先评格式，无需保存重开测试。
    item7 = score_guard(lambda: score_number_format(excel, ws, config))
    items.append(item7)
    return items


def score_guard(func: Any) -> ScoreItem:
    try:
        return func()
    except Exception as exc:
        name = getattr(func, "__name__", "评分项")
        return ScoreItem("ERROR", name, 0, 0, False, [f"评分异常：{exc}"])


def pass_or_fail(item_id: str, name: str, points: int, failures: list[str], evidence: list[str | None] = None) -> ScoreItem:
    ev = list(evidence or [])
    if failures:
        ev.extend(failures)
    return ScoreItem(item_id, name, points, points if not failures else 0, not failures, ev)



def score_odd_rows_ltr(excel: Any, ws: Any, config: EvaluatorConfig) -> ScoreItem:
    """D2-2：奇数录入行按从左往右顺序录入。

    细则逐条对应：
      1. "按照从左往右顺序依次录入" — 遍历每一奇数录入行（row_offset 为偶数），
         逐列从左向右写入并确认跳转。
      2. "在左侧单元格输入完成后按下Enter可以跳转到右侧相邻单元格" — 检查
         col1→col1+1、col1+1→col2 每一步的 Enter 后落点是否恰好是右侧相邻列。
      3. "在奇数行右侧最后一个单元格完成输入后，按Enter跳转到相同列的下一行单元格"
         — 检查 col2（最右列）的 Enter 后落点是否恰好是 (row+1, col2)，即同列下一行。
    """
    row1, col1, row2, col2 = parse_range(config.entry_range)
    failures: list[str] = []
    evidence: list[str] = []
    gen = ValueGenerator()
    clear_entry_range(excel, ws, config.entry_range)
    for row_offset, row in enumerate(range(row1, row2 + 1)):
        # 奇数录入行 = row_offset 为偶数（第1行、第3行……）
        if row_offset % 2 != 0:
            continue
        # 细则第2点：左侧单元格 Enter → 右侧相邻单元格（中间列的逐步跳转）
        for col in range(col1, col2):
            src = make_addr(row, col)
            expected = make_addr(row, col + 1)
            actual = type_value_press_enter_and_get_active(excel, ws, src, gen.next())
            if actual != expected:
                failures.append(f"{src} 期望 {expected}（右侧相邻），实际 {actual}")
        # 细则第3点：右侧最后一个单元格 Enter → 相同列的下一行单元格
        if row < row2:
            src = make_addr(row, col2)
            expected = make_addr(row + 1, col2)   # 同列（col2）下一行
            actual = type_value_press_enter_and_get_active(excel, ws, src, gen.next())
            if actual != expected:
                failures.append(f"{src} 期望 {expected}（同列下一行），实际 {actual}")
    if not failures:
        evidence.append("所有奇数录入行均满足：左→右逐格跳转，行末跳同列下一行")
    return pass_or_fail("D2-2", "奇数录入行从左到右", 5, failures, evidence)


def score_even_rows_rtl(excel: Any, ws: Any, config: EvaluatorConfig) -> ScoreItem:
    """D2-3：偶数录入行按从右往左顺序录入。

    细则逐条对应：
      1. "按照从右往左顺序依次录入" — 遍历每一偶数录入行（row_offset 为奇数），
         逐列从右向左写入并确认跳转。
      2. "在右侧单元格输入完成后按下Enter可以跳转到左侧相邻单元格" — 检查
         col2→col2-1、…→col1 每一步的 Enter 后落点是否恰好是左侧相邻列。
      3. "在偶数行左侧最后一个单元格完成输入后，按Enter跳转到相同列的下一行单元格"
         — 检查 col1（最左列）的 Enter 后落点是否恰好是 (row+1, col1)，即同列下一行。
    """
    row1, col1, row2, col2 = parse_range(config.entry_range)
    failures: list[str] = []
    evidence: list[str] = []
    gen = ValueGenerator()
    clear_entry_range(excel, ws, config.entry_range)
    for row_offset, row in enumerate(range(row1, row2 + 1)):
        # 偶数录入行 = row_offset 为奇数（第2行、第4行……）
        if row_offset % 2 == 0:
            continue
        # 细则第2点：右侧单元格 Enter → 左侧相邻单元格（逐列从右向左）
        for col in range(col2, col1, -1):
            src = make_addr(row, col)
            expected = make_addr(row, col - 1)
            actual = type_value_press_enter_and_get_active(excel, ws, src, gen.next())
            if actual != expected:
                failures.append(f"{src} 期望 {expected}（左侧相邻），实际 {actual}")
        # 细则第3点：左侧最后一个单元格 Enter → 相同列的下一行单元格
        if row < row2:
            src = make_addr(row, col1)
            expected = make_addr(row + 1, col1)   # 同列（col1）下一行
            actual = type_value_press_enter_and_get_active(excel, ws, src, gen.next())
            if actual != expected:
                failures.append(f"{src} 期望 {expected}（同列下一行），实际 {actual}")
    if not failures:
        evidence.append("所有偶数录入行均满足：右→左逐格跳转，行末跳同列下一行")
    return pass_or_fail("D2-3", "偶数录入行从右到左", 5, failures, evidence)


def score_complete_path(excel: Any, ws: Any, path: list[str], config: EvaluatorConfig) -> ScoreItem:
    """D2-4：完整蛇形路径不跳出区域。

    细则逐条对应：
      1. "按奇数录入行从左到右、偶数录入行从右到左的顺序跳转" — 使用
         build_expected_path 生成的完整路径（row_offset 偶数行 LTR、奇数行 RTL），
         逐格按下 Enter 检验每一步的实际落点是否与期望路径吻合。
      2. "中途不跳出三列n行的输入区域" — 每次 Enter 后落点必须仍在
         entry_range 内（path_set 覆盖所有合法单元格），一旦跳出立即记录失败。
      3. 触发方式为办公软件的真实 Enter 按键（SendKeys），而非 COM 赋值，
         确保检测与实际用户操作等价。
    """
    failures: list[str] = []
    evidence: list[str] = []
    path_set = set(path)
    gen = ValueGenerator()
    clear_entry_range(excel, ws, config.entry_range)
    ws.Activate()
    ws.Range(path[0]).Select()
    for idx, addr in enumerate(path):
        # 确认当前活动单元格与期望路径吻合（检测上一步跳转是否准确）
        current = normalize_addr(com_address(excel.ActiveCell))
        if current != addr:
            failures.append(f"第 {idx + 1} 步开始位置错误：期望 {addr}，实际 {current}")
            break
        # 用真实 Enter 按键触发跳转，与办公软件用户操作等价
        actual = type_value_press_enter_and_get_active(excel, ws, addr, gen.next())
        if idx < len(path) - 1:
            expected = path[idx + 1]
            # 细则第1点：顺序正确
            if actual != expected:
                failures.append(f"第 {idx + 1} 步 {addr} 期望跳到 {expected}，实际 {actual}")
                break
        else:
            # 细则第2点：最后一格 Enter 后仍不跳出区域
            if actual not in path_set:
                failures.append(f"最后一个单元格 {addr} 输入后跳出录入区域，实际 {actual}")
    if not failures:
        evidence.append(f"完整检查 {len(path)} 个路径单元格，按蛇形顺序跳转且未跳出 {config.entry_range}")
    return pass_or_fail("D2-4", "完整蛇形路径不跳出区域", 5, failures, evidence)


def score_modify_existing(excel: Any, ws: Any, next_map: dict[str, str], config: EvaluatorConfig) -> ScoreItem:
    """D2-5：修改三列n行中已有数字时，修改完成后仍能跳转到下一个蛇形位置。

    细则逐条对应：
      1. "修改三列n行中已有数字" — 遍历 entry_range 内**全部**单元格，先用
         无事件方式预填数字，确保每一格都"已有内容"；随后逐格测试修改。
         这样能覆盖奇偶行、左右边界（含行末跳同列下一行的位置）以及最后
         一格"回到起点"等所有情况，避免因抽样漏判某些单元格修改失效。
      2. "修改完成后仍能跳转" — 用 type_value_press_enter_and_get_active
         保持事件开启后写入数值，模拟用户在办公软件中覆写数字后按 Enter，
         而非 COM 赋值，确保 Worksheet_Change 宏能被触发。
      3. "跳转到该单元格对应的下一个蛇形位置" — 按 Enter 后的活动单元格
         必须与 next_map 中该格对应的蛇形下一格完全一致。
    """
    failures: list[str] = []
    evidence: list[str] = []
    gen = ValueGenerator()

    # 覆盖 entry_range 内全部单元格（含行末跳同列下一行、末格回到起点等边界）。
    probes: list[str] = [addr for addr in addresses_in_range(config.entry_range) if addr in next_map]

    clear_entry_range(excel, ws, config.entry_range)
    # 预先给整个区域填一遍值，确保每一格在被"修改"前都"已有内容"，
    # 而不是被视作首次录入。写入过程屏蔽事件，避免宏在预填阶段抢先跳转。
    for addr in probes:
        set_value_without_events(excel, ws, addr, gen.next())

    checked = 0
    for src in probes:
        expected = next_map[src]
        # 覆写为新值并触发宏跳转
        actual = type_value_press_enter_and_get_active(excel, ws, src, gen.next())
        checked += 1
        if actual != expected:
            failures.append(f"修改已有值 {src} 后期望跳转到 {expected}，实际 {actual}")

    if not failures:
        evidence.append(f"遍历 {config.entry_range} 全部 {checked} 个单元格，修改已有值后均正确跳转到蛇形下一格")
    else:
        evidence.append(f"共检查 {checked} 个单元格，其中 {len(failures)} 个修改后跳转不符")
    return pass_or_fail("D2-5", "修改已有数字仍跳转", 3, failures, evidence)

def score_number_format(excel: Any, ws: Any, config: EvaluatorConfig) -> ScoreItem:
    """D2-7：录入区域为常规或数值格式，输入数字后显示为普通数字。

    细则逐条对应：
      1. "三列n行区域单元格格式为常规或数值格式" — 检查 entry_range 内
         每个单元格的 NumberFormat / NumberFormatLocal，不只检查固定 B4:D4。
      2. "输入1、2、3等数字后显示为普通数字" — 在录入区域首行三列分别
         用真实键盘输入 1、2、3 并按 Enter，再读取办公软件实际显示文本。
      3. "不能显示为日期、科学计数法或乱码" — 显示文本不得包含日期、科学
         计数法、井号溢出、百分号、货币符号等非普通数字特征，且必须能解析为数值。
    """
    row1, col1, _row2, col2 = parse_range(config.entry_range)
    failures: list[str] = []
    evidence: list[str] = []
    allowed_exact = {"general", "常规", "0", "0.0", "0.00", "#,##0", "#,##0.0", "#,##0.00"}
    forbidden_markers = [
        "@", "%", "yy", "yyyy", "dd", "hh", "ss", "年", "月", "日",
        "e+", "e-", "¥", "￥", "$", "€", "£", "元",
    ]

    def is_general_or_numeric_format(fmt: str, fmt_local: str) -> bool:
        fmt_norm = fmt.strip().lower()
        fmt_local_norm = fmt_local.strip().lower()
        if fmt_norm in allowed_exact or fmt_local_norm in allowed_exact:
            return True
        if any(marker in fmt_norm or marker in fmt_local_norm for marker in forbidden_markers):
            return False
        return any(token in fmt_norm or token in fmt_local_norm for token in ("0", "#"))

    def is_plain_number_text(text: str) -> bool:
        stripped = text.strip()
        if not stripped:
            return False
        upper = stripped.upper()
        bad_markers = ["#", "/", "年", "月", "日", "E+", "E-", "%", "¥", "￥", "$", "€", "£", "元"]
        if any(marker in upper for marker in bad_markers):
            return False
        try:
            float(stripped.replace(",", ""))
        except ValueError:
            return False
        return True

    for addr in addresses_in_range(config.entry_range):
        cell = ws.Range(addr)
        fmt = str(cell.NumberFormat)
        fmt_local = str(cell.NumberFormatLocal)
        if not is_general_or_numeric_format(fmt, fmt_local):
            failures.append(f"{addr} 不是常规或普通数值格式：NumberFormat={fmt!r}, Local={fmt_local!r}")

    # 关闭事件后直接赋值读显示文本，D2-7 只检查格式显示，不涉及跳转
    sample_addrs = [make_addr(row1, col) for col in range(col1, col2 + 1)]
    old_events = excel.EnableEvents
    excel.EnableEvents = False
    try:
        for addr, value in zip(sample_addrs, (1, 2, 3)):
            cell = ws.Range(addr)
            cell.Value2 = value
            text = str(cell.Text)
            if not is_plain_number_text(text):
                failures.append(f"{addr} 输入 {value} 后未显示为普通数字：{text!r}")
    finally:
        excel.EnableEvents = old_events

    if not failures:
        evidence.append(f"{config.entry_range} 均为常规或数值格式，输入 1、2、3 后显示为普通数字")
    return pass_or_fail("D2-7", "录入格式为常规或普通数值", 1, failures, evidence)


# ---------------------------------------------------------------------------
# Report output


def gate_failed_report(candidate_path: Path, config: EvaluatorConfig, original_hash: str | None, reasons: list[str], evidence: dict[str, Any]) -> dict[str, Any]:
    dimension1 = {"passed": False, "reasons": reasons, "evidence": evidence}
    return build_report(candidate_path, config, original_hash, dimension1, [], skipped=True)


def build_report(candidate_path: Path, config: EvaluatorConfig, original_hash: str | None, dimension1: dict[str, Any], items: list[ScoreItem], skipped: bool) -> dict[str, Any]:
    current_hash = file_sha256(candidate_path)
    item_dicts = [item.to_dict() for item in items]
    d2_score = sum(item.awarded for item in items) if not skipped else 0
    return {
        "candidate": str(candidate_path),
        "template_reference": str(Path.cwd() / TEMPLATE_NAME),
        "timestamp": datetime.now().isoformat(timespec="seconds"),
        "config": {
            "sheet_name": config.sheet_name,
            "entry_range": config.entry_range,
            "max_score": MAX_SCORE,
            "timeout_seconds": config.timeout_seconds,
        },
        "original_hash_before": original_hash,
        "original_hash_after": current_hash,
        "original_file_unchanged": bool(original_hash and current_hash and original_hash == current_hash),
        "dimension1": dimension1,
        "dimension2": {
            "skipped": skipped,
            "score": d2_score,
            "max_score": MAX_SCORE,
            "items": item_dicts,
        },
        "final_score": 0 if skipped or not dimension1.get("passed") else d2_score,
        "max_score": MAX_SCORE,
    }


RUBRIC_LINES: dict[str, str] = {
    "D2-2": '+5："蛇形录入"工作表蛇形路径奇数行：按照从左往右顺序依次录入，在左侧单元格输入完成后按下Enter可以跳转到右侧相邻单元格，在奇数行右侧最后一个单元格完成输入后，按Enter跳转到相同列的下一行单元格。',
    "D2-3": '+5："蛇形录入"工作表蛇形路径偶数行：按照从右往左顺序依次录入，在右侧单元格输入完成后按下Enter可以跳转到左侧相邻单元格，在偶数行左侧最后一个单元格完成输入后，按Enter跳转到相同列的下一行单元格。',
    "D2-4": '+5："蛇形录入"工作表完整蛇形路径：按奇数录入行从左到右、偶数录入行从右到左的顺序跳转，中途不跳出三列n行（n为大于3的整数）的输入区域。',
    "D2-5": '+3："蛇形录入"工作表撤销或修改逻辑：修改三列n行中已有数字时，修改完成后仍能跳转到该单元格对应的下一个蛇形位置，不因单元格已有内容而宏失效。',
    "D2-7": '+1："蛇形录入"工作表录入格式：三列n行区域单元格格式为常规或数值格式，输入1、2、3等数字后显示为普通数字，不能显示为日期、科学计数法或乱码。',
}

DIMENSION1_RUBRIC = (
    "维度1 可用与可修改性（前置门槛）：\n"
    "  · 交付文件为 .xlsm 或 .xlsx 格式，文件可正常打开。"
)


# ---------------------------------------------------------------------------
# Unified entry


SCRIPT_ID = "081"


def evaluate_candidate(candidate_path: Path, config: EvaluatorConfig) -> dict[str, Any]:
    candidate_path = candidate_path.expanduser().resolve()
    original_hash = file_sha256(candidate_path)

    static = static_preflight(candidate_path, config)
    if not static["passed"]:
        return gate_failed_report(candidate_path, config, original_hash, static["reasons"], {"static_preflight": static["evidence"]})

    report = run_com_with_timeout(candidate_path, config, original_hash)
    report.setdefault("dimension1", {}).setdefault("evidence", {})["static_preflight"] = static["evidence"]
    return report


def find_candidate_in_dir(dir_path: Path) -> Path | None:
    for name in DEFAULT_CANDIDATE_NAMES:
        path = dir_path / name
        if path.exists() and not path.name.startswith("~$"):
            return path
    excel_files = sorted(
        path for pattern in ("*.xlsm", "*.xlsx")
        for path in dir_path.glob(pattern)
        if not path.name.startswith("~$")
    )
    if excel_files:
        return excel_files[0]
    return None


def _empty_result(status: str, error: str | None, file_name: str = "") -> dict[str, Any]:
    return {
        "id": SCRIPT_ID,
        "file_name": file_name,
        "status": status,
        "error": error,
        "dim1_pass": False,
        "dim1_reason": "",
        "dim2_items": [],
        "total_score": 0,
        "max_score": MAX_SCORE,
    }


def _to_unified_schema(report: dict[str, Any], candidate_path: Path) -> dict[str, Any]:
    d1 = report.get("dimension1", {}) or {}
    d1_passed = bool(d1.get("passed"))
    reasons = d1.get("reasons", []) or []
    dim1_reason = "；".join(str(r) for r in reasons) if isinstance(reasons, list) else str(reasons)

    d2 = report.get("dimension2", {}) or {}
    raw_items = d2.get("items", []) or []
    dim2_items: list[dict[str, Any]] = []
    total_score = 0
    max_score = 0
    for item in raw_items:
        if not isinstance(item, dict):
            continue
        item_id = str(item.get("id", ""))
        item_name = str(item.get("name", ""))
        rule = RUBRIC_LINES.get(item_id, item_name)
        max_delta = int(item.get("points", 0))
        delta = int(item.get("awarded", 0))
        hit = bool(item.get("passed"))
        evidence = item.get("evidence", []) or []
        # 保留 evidence 的读取路径以维持结构不变，但对外输出的 detail 字段置空
        _ = "；".join(str(e) for e in evidence) if isinstance(evidence, list) else str(evidence)
        dim2_items.append({
            "rule": rule,
            "max_delta": max_delta,
            "delta": delta,
            "hit": hit,
            "detail": "",
        })
        if max_delta > 0:
            max_score += max_delta
        total_score += delta

    if not max_score:
        max_score = int(report.get("max_score", MAX_SCORE))
    if not d1_passed or d2.get("skipped"):
        total_score = 0
        max_score = 0
        # 维度一未通过时，dim2_items 置空（保留 dim1_reason 承载失败原因）
        dim2_items = []

    return {
        "id": SCRIPT_ID,
        "file_name": candidate_path.name,
        "status": "ok",
        "error": None,
        "dim1_pass": d1_passed,
        "dim1_reason": "" if d1_passed else dim1_reason,
        "dim2_items": dim2_items,
        "total_score": total_score,
        "max_score": max_score,
    }


def evaluate(dir_path: str) -> dict[str, Any]:
    """统一入口：接收脚本所在目录路径，脚本自行在该目录中定位并打开待评估的 Excel 文档。"""
    try:
        directory = Path(dir_path).expanduser().resolve()
        if not directory.exists() or not directory.is_dir():
            return _empty_result("error", f"目录不存在或不是目录：{dir_path}")

        candidate = find_candidate_in_dir(directory)
        if candidate is None:
            return _empty_result("error", f"目录中未找到可评估的 Excel 文件：{directory}")

        config = EvaluatorConfig()
        # 提前验证录入区域格式，避免子进程中才报错。
        row1, col1, row2, col2 = parse_range(config.entry_range)
        if col2 - col1 + 1 != 3 or row2 - row1 + 1 <= 3:
            return _empty_result("error", "录入区域必须是三列且行数 n 大于 3", candidate.name)

        report = evaluate_candidate(candidate, config)
        return _to_unified_schema(report, candidate)
    except Exception as exc:
        return _empty_result("error", f"{type(exc).__name__}: {exc}")


if __name__ == "__main__":
    # Windows spawn 模式需要保护入口。
    mp.freeze_support()
    target_dir = sys.argv[1] if len(sys.argv) > 1 else str(Path(__file__).parent)
    print(json.dumps(evaluate(target_dir), ensure_ascii=False, indent=2))
