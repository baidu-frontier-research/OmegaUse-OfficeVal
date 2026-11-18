#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# pyright: reportUnknownMemberType=false, reportUnknownVariableType=false, reportUnknownArgumentType=false, reportUnknownParameterType=false, reportMissingParameterType=false, reportPrivateUsage=false, reportUnknownLambdaType=false, reportGeneralTypeIssues=false

import os
import re
import sys
import json
import math
import posixpath
import zipfile
import warnings
import traceback
import xml.etree.ElementTree as ET

warnings.filterwarnings("ignore")
import openpyxl

# 说明：按接口统一约定，本模块不修改全局 sys.stdout，也不 print 主结果。
# 主结果统一由 evaluate(dir_path) -> dict 返回；入参为“脚本所在目录的路径”，
# 由脚本自己在该目录内定位并打开被评估的 .xlsx / .xlsm（忽略以 ~$ 打头的临时锁文件）。
# 本文件仅在 __main__ 下作为本地调试入口，才把 dict 序列化为 JSON 打印。
#
# 与 008 号脚本的差异：本任务（教学资源管理台账·登记/核验对比堆叠柱状图）细则同时
# 含“加分项”（指令要求的图表制作任务）与“减分项”（指令未涉及、却改动已有台账与汇总
# 区而造成破坏或麻烦的多余修改）。故 dim2 项的 max_delta 允许为负：正值=加分项
# （命中记满分，未命中记 0）；负值=减分项（触发“破坏”记该负分，未触发记 0）。
# total_score 为全部 delta 之和，取值区间 [-29, +45]；max_score 仍为所有加分项满分
# 之和（+45），作为完成度上限。图表“做了但没做好”的质量问题（占位 0、穿柱网格线、
# 拼接标签、位图截图、工作表锁定等）只体现为相应加分项不得分，不另计减分项；
# 维度1 只判“交付物是否为一个能正常打开的 Excel 工作簿（.xlsx/.xlsm）”，
# 文件损坏或需修复时直接判 0，其余问题一律回落到 dim2 计分，不重复判 0。
#
# 图表侧一律直接解析包内 chart XML（含第三方写出的 xl/drawings/charts/*.xml）：
# openpyxl 对某些写法的 c:externalData 关系会整张图表跳过，不能作为判据；
# 单元格取值、字体填充边框、合并、列宽行高、视图与保护仍走 openpyxl。

SCRIPT_ID = "100"

# —— 加分项按“结果追加顺序”的 max_delta（正值）——
_ADD_DELTAS_ORDER = [
    3,  # A01 新增 1 个柱形图图表对象，用于登记/核验对比
    5,  # A02 登记侧为“综管科/服务科”上下堆叠的堆积柱形图
    3,  # A03 横轴 5 个资源类别，每类别只 1 个分类标签
    3,  # A04 “核验数量”为独立柱体，与登记堆叠柱同刻度左右并排
    3,  # A05 综管科 6/22/18/12/5、服务科 9/8/6/3/2
    3,  # A06 核验数量 13/27/22/14/6
    1,  # A07 堆叠柱总高 = 15/30/24/15/7
    5,  # A08 系列名/分类轴/数值均为单元格区域引用（非常量数组）
    3,  # A09 规整绘图数据表（5 行 × 类别/综管科/服务科/核验数量）
    1,  # A10 各系列取值均为真实数据，不含拼并列用的占位 0
    1,  # A11 可逐系列调整、非单一图形对象，且工作表未被保护锁定
    1,  # A12 图表标题含登记（综管科+服务科）与核验对比信息
    1,  # A13 纵坐标轴标题“数量（台）”
    3,  # A14 综管科/服务科两段数据标签居中显示于各段柱体内部
    3,  # A15 堆叠柱顶显示登记合计标签、核验柱顶显示核验标签
    1,  # A16 图例含三项且不显示承载合计标签的辅助系列名
    1,  # A17 综管科/服务科/核验数量三色明显互不相同
    1,  # A18 纵轴最大刻度不小于 30，柱体与标签无截断
    1,  # A19 同类别两根柱紧邻成组，组内间距明显小于类别间距
    1,  # A20 图表摆放不遮挡台账主表与汇总辅助区
    1,  # A21 网格线只在数值轴，柱体不被竖向网格线穿过，标签不重叠
]

# —— 减分项按“结果追加顺序”的 max_delta（负值）——
_DED_DELTAS_ORDER = [
    -3,  # D01 台账主表数据被删除/改写/移位
    -3,  # D02 主表区域插入或删除行/列致数据与合并错位
    -1,  # D03 A1 标题格式
    -1,  # D04 表头行 A2:F2 格式
    -1,  # D05 数据行 A3:F16 格式
    -1,  # D06 “合计”行（5/8/11/14/17 行）格式
    -1,  # D07 主表 20 处纵向/跨列合并被取消或范围被改
    -1,  # D08 列宽 A/B/C/D/F 或行高（1/2/3 行及以下）被改动
    -1,  # D09 主表表头文字被改写为图表字段名
    -3,  # D10 H2:M7 汇总表被删除或数值被改写
    -3,  # D11 H10:M15 转置汇总表被删除或数值被改写
    -1,  # D12 H2:M2 与 H10:M10 表头格式
    -1,  # D13 H13:M14“合计”“核验数量”两行格式
    -1,  # D14 H15:M15“辅助列”行被删除或填充/边框被改
    -1,  # D15 Sheet1 被改名/隐藏/删除/非首个工作表
    -1,  # D16 绘图数据表覆盖写入 A1:F17 或 H2:M15
    -1,  # D17 残留与任务无关的多余工作表或空白工作表
    -1,  # D18 残留未完成的中间图表或功能重复的第二个对比图
    -1,  # D19 绘图数据表出现指令未要求的自造内容
    -1,  # D20 默认打开视图被改动（冻结/网格线/缩放/活动表）
    -1,  # D21 台账与汇总区新增公式/批注/条件格式/数据验证
]

MAX_SCORE_TOTAL = sum(_ADD_DELTAS_ORDER)               # +45：完成度上限
_MIN_TOTAL = MAX_SCORE_TOTAL + sum(_DED_DELTAS_ORDER)  # +16：全部命中且全部踩坑
assert MAX_SCORE_TOTAL == 45
assert sum(_DED_DELTAS_ORDER) == -29

# 结果追加顺序 = 先加分项、后减分项
_MAX_DELTAS_ORDER = _ADD_DELTAS_ORDER + _DED_DELTAS_ORDER
_ADD_COUNT = len(_ADD_DELTAS_ORDER)
assert len(_MAX_DELTAS_ORDER) == 42

# ================= 任务基准数据（细则中逐项写明的、应保持不变/应正确的取值）=================

_TITLE_TEXT = "教学资源管理汇总台账"
_CATS = ["便携学习终端", "智能显示设备", "网络接入设备", "多功能文印设备", "移动采集设备"]
_ZG = [6.0, 22.0, 18.0, 12.0, 5.0]     # 综管科
_FW = [9.0, 8.0, 6.0, 3.0, 2.0]        # 服务科
_HY = [13.0, 27.0, 22.0, 14.0, 6.0]    # 核验数量
_TOT = [15.0, 30.0, 24.0, 15.0, 7.0]   # 登记合计

# 台账主表 A1:F17 应有取值（键为单元格坐标）
_MAIN_TITLE_CELL = "A1"
_MAIN_HEADER = {
    "A2": "序号", "B2": "资源类别", "C2": "所属单元",
    "D2": "登记数量（件）", "E2": "核验数量（件）", "F2": "备注",
}
_MAIN_DATA = {
    "A3": 1, "B3": "便携学习终端", "C3": "综管科", "D3": 6, "E3": 13,
    "C4": "服务科", "D4": 9,
    "A5": "合计", "D5": 15,
    "A6": 2, "B6": "智能显示设备", "C6": "综管科", "D6": 22, "E6": 27,
    "C7": "服务科", "D7": 8,
    "A8": "合计", "D8": 30,
    "A9": 3, "B9": "网络接入设备", "C9": "综管科", "D9": 18, "E9": 22,
    "C10": "服务科", "D10": 6,
    "A11": "合计", "D11": 24,
    "A12": 4, "B12": "多功能文印设备", "C12": "综管科", "D12": 12, "E12": 14,
    "C13": "服务科", "D13": 3,
    "A14": "合计", "D14": 15,
    "A15": 5, "B15": "移动采集设备", "C15": "综管科", "D15": 5, "E15": 6,
    "C16": "服务科", "D16": 2,
    "A17": "合计", "D17": 7,
}
# 右侧汇总辅助区 H2:M7 应有取值
_HELP_TOP = {
    "H2": "资源类别", "I2": "综管科", "J2": "服务科", "K2": "合计",
    "L2": "核验数量", "M2": "辅助列",
    "H3": "便携学习终端", "I3": 6, "J3": 9, "K3": 15, "L3": 13,
    "H4": "智能显示设备", "I4": 22, "J4": 8, "K4": 30, "L4": 27,
    "H5": "网络接入设备", "I5": 18, "J5": 6, "K5": 24, "L5": 22,
    "H6": "多功能文印设备", "I6": 12, "J6": 3, "K6": 15, "L6": 14,
    "H7": "移动采集设备", "I7": 5, "J7": 2, "K7": 7, "L7": 6,
}

# 右侧转置汇总表 H10:M15 应有取值
_HELP_BOT = {
    "H10": "资源类别", "I10": "便携学习终端", "J10": "智能显示设备",
    "K10": "网络接入设备", "L10": "多功能文印设备", "M10": "移动采集设备",
    "H11": "综管科", "I11": 6, "J11": 22, "K11": 18, "L11": 12, "M11": 5,
    "H12": "服务科", "I12": 9, "J12": 8, "K12": 6, "L12": 3, "M12": 2,
    "H13": "合计", "I13": 15, "J13": 30, "K13": 24, "L13": 15, "M13": 7,
    "H14": "核验数量", "I14": 13, "J14": 27, "K14": 22, "L14": 14, "M14": 6,
    "H15": "辅助列",
}

# 主表应保持的 20 处合并（A1:F1 归入 A1 标题项判定）
_MAIN_MERGES = (
    "A3:A4", "A6:A7", "A9:A10", "A12:A13", "A15:A16",
    "B3:B4", "B6:B7", "B9:B10", "B12:B13", "B15:B16",
    "A5:C5", "A8:C8", "A11:C11", "A14:C14", "A17:C17",
    "E3:E5", "E6:E8", "E9:E11", "E12:E14", "E15:E17",
)
# 列宽与行高（细则给出的显示值；比较时留 0.1 字符 / 0.6 磅容差）
_COL_WIDTHS = {"A": 4.875, "B": 16.125, "C": 21.5, "D": 16.375, "F": 9.0,
               "H": 12.0, "I": 9.0, "K": 13.875, "L": 9.0}
_ROW_HEIGHTS = {1: 28.5, 2: 27.0}
_ROW_HEIGHT_BODY = 26.0
_SHEET_NAME = "Sheet1"

# 配色（细则逐项写明）
_C_TITLE_FILL = "1F4E5F"
_C_HEAD_FILL = "6FA8A0"
_C_BODY_FILL = "E9F4F2"
_C_SUM_FILL = "D8EBE8"
_C_SUM_FONT = "1F2933"
_C_WHITE = "FFFFFF"
_TITLE_FONT = "方正小标宋简体"

# 受保护矩形（1 起的 行1,列1,行2,列2）：台账主表 A1:F17、汇总辅助区 H2:M15
_RECT_MAIN = (1, 1, 17, 6)
_RECT_HELP = (2, 8, 15, 13)

# 绘图数据表允许出现的文字（超出即视为自造内容）
_ALLOWED_TEXT = set(_CATS) | {
    "资源类别", "类别", "序号", "所属单元", "备注", "辅助列",
    "综管科", "服务科", "合计", "核验数量", "登记数量", "登记合计",
    "综管科+服务科", "登记数量（件）", "核验数量（件）", "数量（台）",
    "绘图数据", "绘图数据表", "图表数据源", "数据源", "登记合计（综管科+服务科）",
}
_ALLOWED_NUMS = set(_ZG) | set(_FW) | set(_HY) | set(_TOT) | {1.0, 2.0, 3.0, 4.0, 5.0}
# ================= 通用小工具 =================

NS_C = "http://schemas.openxmlformats.org/drawingml/2006/chart"
NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main"
NS_XDR = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_S = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"


def _q(ns, tag):
    return "{%s}%s" % (ns, tag)


def _norm(s):
    """文字规范化：去空白、半角括号转全角，便于与细则字面比对。"""
    if s is None:
        return ""
    t = str(s).replace("(", "（").replace(")", "）").replace("＋", "+")
    return re.sub(r"\s+", "", t).strip()


def _ext_of(path):
    """取小写扩展名（含点），用于维度1 说明中回显交付格式。"""
    return os.path.splitext(path or "")[1].lower()


def _num(v):
    """取数值；文本型数字一并接受，非数字返回 None。"""
    if isinstance(v, bool):
        return None
    if isinstance(v, (int, float)):
        return float(v)
    if isinstance(v, str):
        try:
            return float(v.strip())
        except Exception:
            return None
    return None


def _same_value(got, want):
    """单元格取值比对：数值按数值比，文字按规范化后比。"""
    if isinstance(want, (int, float)) and not isinstance(want, bool):
        g = _num(got)
        return g is not None and abs(g - float(want)) < 1e-9
    return _norm(got) == _norm(want)


def _nz(seq):
    """去掉 0 值占位后的数值序列（用于容忍“靠 0 值拼并列”的实现）。"""
    return [v for v in seq if v is not None and abs(v) > 1e-9]
def _rgb6(color):
    """openpyxl 颜色 -> 6 位十六进制大写；解析不出返回 None。"""
    try:
        if color is None:
            return None
        rgb = getattr(color, "rgb", None)
        if isinstance(rgb, str) and re.fullmatch(r"[0-9A-Fa-f]{6,8}", rgb):
            return rgb[-6:].upper()
    except Exception:
        pass
    return None


def _fill6(cell):
    """单元格实心填充色（6 位十六进制），非实心填充返回 None。"""
    try:
        f = cell.fill
        if f is None or f.fill_type != "solid":
            return None
        return _rgb6(f.start_color)
    except Exception:
        return None


def _borders(cell):
    """四边线型元组 (左,右,上,下)。"""
    b = cell.border
    if b is None:
        return (None, None, None, None)
    return tuple(
        (getattr(b, side).style if getattr(b, side) is not None else None)
        for side in ("left", "right", "top", "bottom")
    )


def _thin4(cell):
    return all(s == "thin" for s in _borders(cell))


def _centered(cell):
    a = cell.alignment
    return (a is not None and a.horizontal == "center" and a.vertical == "center")


def _wrap_ok(cell):
    """自动换行：显式关闭才算被改动（重存后 None 与 True 视觉一致）。"""
    a = cell.alignment
    return not (a is not None and a.wrap_text is False)
def _merged_map(ws):
    """返回 (合并区字符串集合, 被合并覆盖但非左上角的坐标集合)。"""
    ranges = set()
    covered = set()
    for mr in ws.merged_cells.ranges:
        ranges.add(str(mr))
        for row in ws[str(mr)]:
            for c in row:
                if c.coordinate != mr.coord.split(":")[0]:
                    covered.add(c.coordinate)
    return ranges, covered


def _iter_rect(ws, rect):
    """遍历矩形区域 (r1,c1,r2,c2) 内的单元格对象。"""
    r1, c1, r2, c2 = rect
    for row in ws.iter_rows(min_row=r1, max_row=r2, min_col=c1, max_col=c2):
        for c in row:
            yield c


def _in_rect(row, col, rect):
    r1, c1, r2, c2 = rect
    return r1 <= row <= r2 and c1 <= col <= c2


def _rect_overlap(a, b):
    return not (a[2] < b[0] or b[2] < a[0] or a[3] < b[1] or b[3] < a[1])


_REF_RE = re.compile(r"^(?:'([^']+)'|([^!]+))!\$?([A-Z]{1,3})\$?(\d+)"
                     r"(?::\$?([A-Z]{1,3})\$?(\d+))?$")


def _parse_ref(f):
    """解析 c:f 形如 Sheet1!$B$20:$B$24 -> (sheet, r1, c1, r2, c2)；失败返回 None。"""
    if not f:
        return None
    m = _REF_RE.match(f.strip())
    if not m:
        return None
    sheet = m.group(1) or m.group(2)
    from openpyxl.utils import column_index_from_string
    c1 = column_index_from_string(m.group(3))
    r1 = int(m.group(4))
    c2 = column_index_from_string(m.group(5)) if m.group(5) else c1
    r2 = int(m.group(6)) if m.group(6) else r1
    return (sheet, min(r1, r2), min(c1, c2), max(r1, r2), max(c1, c2))
# ================= 包内 XML：图表 / 绘图 / 关系 =================


def _val_of(el, default=None):
    if el is None:
        return default
    v = el.get("val")
    return v if v is not None else default


def _rich_text(el):
    if el is None:
        return ""
    return "".join(t.text or "" for t in el.iter(_q(NS_A, "t"))).strip()


def _ser_fill(sp):
    """系列填充：'none' 表示无填充（承载合计标签的辅助系列），否则 6 位色值。"""
    if sp is None:
        return None
    if sp.find(_q(NS_A, "noFill")) is not None:
        return "none"
    sf = sp.find(_q(NS_A, "solidFill"))
    if sf is None:
        return None
    c = sf.find(_q(NS_A, "srgbClr"))
    if c is not None:
        return (c.get("val") or "").upper()
    c = sf.find(_q(NS_A, "schemeClr"))
    if c is not None:
        return "scheme:" + (c.get("val") or "")
    return "other"


def _dlbls(el):
    """返回 (showVal, dLblPos 集合)：合并 c:dLbls 与逐点 c:dLbl 的设置。"""
    d = el.find(_q(NS_C, "dLbls"))
    if d is None:
        return (None, set())
    show = _val_of(d.find(_q(NS_C, "showVal")))
    for sub in d.findall(_q(NS_C, "dLbl")):
        s2 = sub.find(_q(NS_C, "showVal"))
        if s2 is not None and s2.get("val") == "1":
            show = "1"
    poss = set(p.get("val") for p in d.iter(_q(NS_C, "dLblPos")))
    return (show, set(p for p in poss if p))
def _pts_num(holder):
    """c:val -> (引用公式 or None, 数值列表, 是否常量数组)。"""
    if holder is None:
        return (None, [], False)
    ref = holder.find(_q(NS_C, "numRef"))
    lit = holder.find(_q(NS_C, "numLit"))
    src, is_lit = (ref, False) if ref is not None else (lit, True)
    if src is None:
        return (None, [], False)
    f = src.find(_q(NS_C, "f"))
    formula = f.text if f is not None else None
    cache = src if is_lit else src.find(_q(NS_C, "numCache"))
    vals = []
    if cache is not None:
        for pt in cache.findall(_q(NS_C, "pt")):
            v = pt.find(_q(NS_C, "v"))
            vals.append(_num(v.text) if v is not None else None)
    return (formula, vals, is_lit)


def _pts_str(holder):
    """c:cat -> (引用公式 or None, 文字列表, 是否常量数组)；数值型分类回退到 _pts_num。"""
    if holder is None:
        return (None, [], False)
    ref = holder.find(_q(NS_C, "strRef"))
    lit = holder.find(_q(NS_C, "strLit"))
    if ref is None and lit is None:
        f, vals, is_lit = _pts_num(holder)
        return (f, ["" if v is None else str(v) for v in vals], is_lit)
    src, is_lit = (ref, False) if ref is not None else (lit, True)
    f = src.find(_q(NS_C, "f"))
    formula = f.text if f is not None else None
    cache = src if is_lit else src.find(_q(NS_C, "strCache"))
    vals = []
    if cache is not None:
        for pt in cache.findall(_q(NS_C, "pt")):
            v = pt.find(_q(NS_C, "v"))
            vals.append((v.text or "").strip() if v is not None else "")
    return (formula, vals, is_lit)


def _ser_name(ser):
    """系列名称 -> (名称文字, 引用公式 or None)。"""
    tx = ser.find(_q(NS_C, "tx"))
    if tx is None:
        return ("", None)
    sref = tx.find(_q(NS_C, "strRef"))
    if sref is not None:
        f = sref.find(_q(NS_C, "f"))
        name = ""
        cache = sref.find(_q(NS_C, "strCache"))
        if cache is not None:
            p = cache.find(_q(NS_C, "pt"))
            if p is not None and p.find(_q(NS_C, "v")) is not None:
                name = (p.find(_q(NS_C, "v")).text or "").strip()
        return (name, f.text if f is not None else None)
    v = tx.find(_q(NS_C, "v"))
    return (((v.text or "").strip() if v is not None else ""), None)
def _int_of(el, default):
    """取 val 属性并转 int；缺失或非数字返回默认值（0 值不可被 or 吞掉）。"""
    v = _num(_val_of(el))
    return int(v) if v is not None else default


def _parse_chart(xml_bytes):
    """解析一个 chartSpace，返回结构化图表信息（不依赖 openpyxl 的图表读取）。"""
    info = {
        "title": "", "auto_title_deleted": None, "groups": [], "other_types": [],
        "catAx": [], "valAx": [], "legend": False, "legend_pos": None,
        "legend_deleted": [], "n_series": 0, "all_have_spPr": True,
    }
    root = ET.fromstring(xml_bytes)
    chart = root.find(_q(NS_C, "chart"))
    if chart is None:
        return info
    info["title"] = _rich_text(chart.find(_q(NS_C, "title")))
    info["auto_title_deleted"] = _val_of(chart.find(_q(NS_C, "autoTitleDeleted")))
    pa = chart.find(_q(NS_C, "plotArea"))
    if pa is None:
        return info
    for child in pa:
        tag = child.tag.split("}")[-1]
        if tag == "barChart":
            grp = {
                "grouping": _val_of(child.find(_q(NS_C, "grouping"))),
                "barDir": _val_of(child.find(_q(NS_C, "barDir"))),
                "gapWidth": _int_of(child.find(_q(NS_C, "gapWidth")), 150),
                "overlap": _int_of(child.find(_q(NS_C, "overlap")), 0),
                "axIds": [a.get("val") for a in child.findall(_q(NS_C, "axId"))],
                "series": [],
            }
            g_show, g_pos = _dlbls(child)
            for ser in child.findall(_q(NS_C, "ser")):
                name, name_ref = _ser_name(ser)
                cat_ref, cats, cat_lit = _pts_str(ser.find(_q(NS_C, "cat")))
                val_ref, vals, val_lit = _pts_num(ser.find(_q(NS_C, "val")))
                s_show, s_pos = _dlbls(ser)
                sp = ser.find(_q(NS_C, "spPr"))
                if sp is None:
                    info["all_have_spPr"] = False
                grp["series"].append({
                    "name": name, "name_ref": name_ref,
                    "cats": cats, "cat_ref": cat_ref, "cat_lit": cat_lit,
                    "vals": vals, "val_ref": val_ref, "val_lit": val_lit,
                    "fill": _ser_fill(sp),
                    "idx": _int_of(ser.find(_q(NS_C, "idx")), 0),
                    "order": _int_of(ser.find(_q(NS_C, "order")), 0),
                    "show": s_show if s_show is not None else g_show,
                    "pos": s_pos or g_pos,
                })
                info["n_series"] += 1
            info["groups"].append(grp)
        elif tag.endswith("Chart"):
            info["other_types"].append(tag)
        elif tag in ("catAx", "valAx"):
            ax = {
                "axId": _val_of(child.find(_q(NS_C, "axId"))),
                "delete": _val_of(child.find(_q(NS_C, "delete"))),
                "grid": child.find(_q(NS_C, "majorGridlines")) is not None,
                "title": _rich_text(child.find(_q(NS_C, "title"))),
                "max": None,
            }
            sc = child.find(_q(NS_C, "scaling"))
            if sc is not None:
                mx = sc.find(_q(NS_C, "max"))
                ax["max"] = _num(mx.get("val")) if mx is not None else None
            info[tag].append(ax)
    lg = chart.find(_q(NS_C, "legend"))
    if lg is not None:
        info["legend"] = True
        info["legend_pos"] = _val_of(lg.find(_q(NS_C, "legendPos")))
        for le in lg.findall(_q(NS_C, "legendEntry")):
            if _val_of(le.find(_q(NS_C, "delete"))) == "1":
                info["legend_deleted"].append(_int_of(le.find(_q(NS_C, "idx")), -1))
    return info


def _resolve(part, target):
    """按 OPC 规则把关系 Target 解析成包内绝对部件名。"""
    if target.startswith("/"):
        return target.lstrip("/")
    return posixpath.normpath(posixpath.join(posixpath.dirname(part), target))


def _rels_of(z, part):
    """读取某部件的关系表：返回 [(Id, Type, 解析后的部件名 or 外部 URI, 是否外部)]。"""
    rels_name = posixpath.join(posixpath.dirname(part), "_rels",
                               posixpath.basename(part) + ".rels")
    out = []
    if rels_name not in z.namelist():
        return out
    try:
        root = ET.fromstring(z.read(rels_name))
    except Exception:
        return out
    for r in root:
        rid = r.get("Id")
        rtype = r.get("Type") or ""
        tgt = r.get("Target") or ""
        ext = (r.get("TargetMode") == "External")
        out.append((rid, rtype, tgt if ext else _resolve(part, tgt), ext))
    return out
def _anchor_pos(el):
    """读取 xdr:from / xdr:to 的 (行, 列)，均转成 1 起。"""
    if el is None:
        return None
    c = el.find(_q(NS_XDR, "col"))
    r = el.find(_q(NS_XDR, "row"))
    try:
        return (int(r.text) + 1, int(c.text) + 1)
    except Exception:
        return None


def _anchor_info(an, drels):
    """解析一个锚点：返回 (rect, 是否图表, 是否图片, 图表部件名 or None)。
       rect = (r1, c1, r2, c2)，oneCellAnchor 依据 xdr:ext 估算占位大小。"""
    frm = _anchor_pos(an.find(_q(NS_XDR, "from")))
    to = _anchor_pos(an.find(_q(NS_XDR, "to")))
    if to is None and frm is not None:
        ext = an.find(_q(NS_XDR, "ext"))
        dr, dc = 15, 8
        if ext is not None:
            try:
                dc = max(1, int(round(int(ext.get("cx")) / 640080.0)))
                dr = max(1, int(round(int(ext.get("cy")) / 190500.0)))
            except Exception:
                pass
        to = (frm[0] + dr, frm[1] + dc)
    rect = (frm[0], frm[1], to[0], to[1]) if (frm and to) else None
    is_pic = an.find(".//" + _q(NS_XDR, "pic")) is not None
    cpart = None
    gf = an.find(".//" + _q(NS_XDR, "graphicFrame"))
    if gf is not None:
        for ch in gf.iter():
            rid = ch.get(_q(NS_R, "id"))
            if rid and rid in drels:
                cpart = drels[rid]
                break
    return (rect, gf is not None, is_pic, cpart)


def _sheet_parts(z):
    """按 workbook.xml 中的顺序返回 [(工作表名, 部件名)]。"""
    out = []
    try:
        wbx = ET.fromstring(z.read("xl/workbook.xml"))
    except Exception:
        return out
    rels = {rid: tgt for rid, _t, tgt, ext in _rels_of(z, "xl/workbook.xml") if not ext}
    sheets = wbx.find(_q(NS_S, "sheets"))
    if sheets is None:
        return out
    for sh in sheets:
        rid = sh.get(_q(NS_R, "id"))
        out.append((sh.get("name") or "", rels.get(rid, "")))
    return out


def _collect_charts(z):
    """收集包内全部图表：
       返回 (charts, anchors, pic_count)
       charts  = [{part, info, sheet}]
       anchors = [{sheet, rect(行列 1 起), is_chart, is_pic}]"""
    charts = []
    anchors = []
    pic_count = 0
    names = set(z.namelist())
    chart_sheet = {}
    for sname, spart in _sheet_parts(z):
        if not spart or spart not in names:
            continue
        for _rid, rtype, tgt, ext in _rels_of(z, spart):
            if ext or not rtype.endswith("/drawing"):
                continue
            if tgt not in names:
                continue
            drels = {rid: t for rid, _ty, t, e in _rels_of(z, tgt) if not e}
            try:
                droot = ET.fromstring(z.read(tgt))
            except Exception:
                continue
            for an in list(droot):
                rect, is_chart, is_pic, cpart = _anchor_info(an, drels)
                if is_pic:
                    pic_count += 1
                anchors.append({"sheet": sname, "rect": rect,
                                "is_chart": is_chart, "is_pic": is_pic})
                if cpart:
                    chart_sheet[cpart] = sname
    for n in sorted(names):
        if re.search(r"chart\d*\.xml$", n) and "/colors" not in n and "/style" not in n:
            try:
                charts.append({"part": n, "info": _parse_chart(z.read(n)),
                               "sheet": chart_sheet.get(n)})
            except Exception:
                charts.append({"part": n, "info": None, "sheet": chart_sheet.get(n)})
    return charts, anchors, pic_count
def _package_broken(z):
    """包结构完整性：返回问题描述列表（空表示结构正常、能被 Office 正常打开）。"""
    bad = []
    try:
        if z.testzip() is not None:
            bad.append("压缩包存在损坏条目")
    except Exception as exc:
        bad.append("压缩包无法校验：%s" % exc)
    names = set(z.namelist())
    if "[Content_Types].xml" not in names:
        bad.append("缺少 [Content_Types].xml")
        return bad
    try:
        ct = ET.fromstring(z.read("[Content_Types].xml"))
    except Exception as exc:
        bad.append("[Content_Types].xml 无法解析：%s" % exc)
        return bad
    declared = set()
    for ov in ct.iter():
        pn = ov.get("PartName")
        if pn:
            declared.add(pn.lstrip("/"))
    defaults = set()
    for de in ct.iter():
        ex = de.get("Extension")
        if ex:
            defaults.add(ex.lower())
    for n in sorted(names):
        if n.endswith("/") or "_rels/" in n:
            continue
        ext = n.rsplit(".", 1)[-1].lower() if "." in n else ""
        if n not in declared and ext not in defaults:
            bad.append("部件 %s 未在 [Content_Types].xml 中声明" % n)
    for n in sorted(names):
        if not n.endswith(".rels"):
            continue
        base = posixpath.dirname(posixpath.dirname(n))
        try:
            rroot = ET.fromstring(z.read(n))
        except Exception:
            continue
        for rel in rroot:
            tgt = rel.get("Target") or ""
            if (rel.get("TargetMode") or "") == "External" or "://" in tgt:
                continue
            full = tgt.lstrip("/") if tgt.startswith("/") \
                else posixpath.normpath(posixpath.join(base, tgt))
            if full not in names:
                bad.append("关系 %s 指向缺失部件 %s" % (rel.get("Id"), full))
    for n in sorted(names):
        if n.endswith(".xml") or n.endswith(".rels"):
            try:
                ET.fromstring(z.read(n))
            except Exception as exc:
                bad.append("部件 %s XML 非法：%s" % (n, exc))
    return bad
# ================= 入口 =================

def evaluate(dir_path):
    """评测指定目录下的输出文件，返回统一结构的评分结果字典。

    未找到待评文件（目录不存在或目录内没有工作簿）时，返回 status='error'、
    file_name 与 dim1_pass/dim1_reason 均留空、dim2_items 为空、分值全 0；
    其他运行异常则给出“脚本运行异常：…”摘要，并按维度一未通过计 0 分。"""
    result = {
        "id": SCRIPT_ID,
        "file_name": "",
        "status": "ok",
        "error": None,
        "dim1_pass": False,
        "dim1_reason": "",
        "dim2_items": [],
        "total_score": 0,
        "max_score": MAX_SCORE_TOTAL,
    }
    try:
        if not dir_path or not os.path.isdir(dir_path):
            raise FileNotFoundError("目录不存在或不是目录：%s" % dir_path)
        names = sorted(f for f in os.listdir(dir_path)
                       if not f.startswith("~$")
                       and os.path.isfile(os.path.join(dir_path, f)))
        # 交付格式：.xlsx（首选）与 .xlsm（启用宏的工作簿）均按 OOXML 解析
        cands = ([f for f in names if f.lower().endswith(".xlsx")]
                 + [f for f in names if f.lower().endswith(".xlsm")])
        if not cands:
            legacy = [f for f in names if f.lower().endswith(".xls")]
            if legacy:
                result["file_name"] = legacy[0]
                raise ValueError("交付文件为 .xls 二进制格式，无法按 OOXML 解析：%s" % legacy[0])
            raise FileNotFoundError("目录中未找到 .xlsx 文件：%s" % dir_path)
        file_path = os.path.join(dir_path, cands[0])
        result["file_name"] = cands[0]
        return _evaluate_impl(file_path, result)
    except FileNotFoundError as exc:
        # 未找到待评文件：不进入维度判定，dim1_pass 留空，分值一并归零
        result["file_name"] = ""
        result["status"] = "error"
        result["error"] = str(exc)
        result["dim1_pass"] = ""
        result["dim1_reason"] = ""
        result["dim2_items"] = []
        result["total_score"] = 0
        result["max_score"] = 0
        return result
    except Exception as exc:
        result["status"] = "error"
        result["error"] = "脚本运行异常：%s: %s" % (type(exc).__name__, exc)
        result["dim1_pass"] = False
        # dim1_reason 统一带上「维度一未通过：…」前缀与计分后果，便于直接作为错误摘要展示
        result["dim1_reason"] = ("维度一未通过：文件无法打开或解析（%s: %s），完成度总分记 0"
                                % (type(exc).__name__, exc))
        result["dim2_items"] = []
        result["total_score"] = 0
        return result


# ================= 证据采集 =================

def _find_main_ws(wb):
    """定位台账主表所在工作表：优先按 A1 标题命中，其次按原名，最后取第一张。"""
    for ws in wb.worksheets:
        if _norm(ws["A1"].value) == _TITLE_TEXT:
            return ws
    if _SHEET_NAME in wb.sheetnames:
        return wb[_SHEET_NAME]
    return wb.worksheets[0]


def _bar_groups(charts):
    """把所有图表的柱形图组摊平：[(chart_index, info, group)]。"""
    out = []
    for i, ch in enumerate(charts):
        info = ch.get("info")
        if not info:
            continue
        for g in info.get("groups", []):
            out.append((i, info, g))
    return out


def _main_chart(charts):
    """挑选“主图表”：优先柱形图系列数最多者，其次部件名靠前者。"""
    best = None
    for ch in charts:
        info = ch.get("info")
        if not info:
            continue
        n = sum(len(g.get("series", [])) for g in info.get("groups", []))
        if n <= 0:
            continue
        key = (n, -len(info.get("other_types", [])))
        if best is None or key > best[0]:
            best = (key, ch)
    if best:
        return best[1]
    return charts[0] if charts else None


def _series_by_name(charts, *keys):
    """在全部柱形图组中按系列名关键字查找：返回 [(group, series)]。"""
    hits = []
    for _i, _info, g in _bar_groups(charts):
        for s in g.get("series", []):
            nm = _norm(s.get("name"))
            if any(k in nm for k in keys):
                hits.append((g, s))
    return hits
def _evaluate_impl(file_path, result):
    """按细则逐项判定；results 内每条形如 "+N：说明" / "-N：说明"。"""
    results = []

    # ---------- 证据采集：包内 XML（图表、绘图锚点、包结构） ----------
    with zipfile.ZipFile(file_path) as z:
        broken = _package_broken(z)
        charts, anchors, pic_count = _collect_charts(z) if not broken else ([], [], 0)
    if broken:
        result["dim1_pass"] = False
        # dim1_reason 统一带上「维度一未通过：…」前缀与计分后果，便于直接作为错误摘要展示
        result["dim1_reason"] = ("维度一未通过：文件包结构异常，Excel 打开时会提示需要修复"
                                "（%s），完成度总分记 0" % broken[0])
        result["dim2_items"] = []
        result["total_score"] = 0
        return result
    live = [c for c in charts if c.get("info")]
    groups = _bar_groups(charts)
    mc = _main_chart(charts)
    minfo = mc.get("info") if mc else None
    n_series = sum(len(g.get("series", [])) for _i, _f, g in groups)
    # 有效图表：至少含 1 个系列（柱形或其它类型）
    real_charts = [c for c in live
                   if sum(len(g.get("series", [])) for g in c["info"]["groups"]) > 0
                   or c["info"]["other_types"]]
    has_chart_frame = any(a["is_chart"] for a in anchors) or bool(real_charts)

    # ---------- 证据采集：单元格 ----------
    try:
        wb = openpyxl.load_workbook(file_path)
        ws = _find_main_ws(wb)
    except Exception as exc:
        result["dim1_pass"] = False
        result["dim1_reason"] = ("维度一未通过：工作簿无法正常读取（%s: %s），完成度总分记 0"
                                % (type(exc).__name__, exc))
        result["dim2_items"] = []
        result["total_score"] = 0
        return result
    merges, covered = _merged_map(ws)
    base_coords = set(_MAIN_DATA) | set(_MAIN_HEADER) | set(_HELP_TOP) | set(_HELP_BOT)
    base_coords.add(_MAIN_TITLE_CELL)

    def cell(coord):
        return ws[coord]

    def ok_value(coord, want):
        return _same_value(ws[coord].value, want)

    # 主表 / 汇总区取值一致性
    main_bad = [c for c, v in _MAIN_DATA.items() if not ok_value(c, v)]
    if not _same_value(ws[_MAIN_TITLE_CELL].value, _TITLE_TEXT):
        main_bad.append(_MAIN_TITLE_CELL)
    head_bad = [c for c, v in _MAIN_HEADER.items() if not ok_value(c, v)]
    htop_bad = [c for c, v in _HELP_TOP.items() if not ok_value(c, v)]
    hbot_bad = [c for c, v in _HELP_BOT.items() if not ok_value(c, v)]

    # ---------- 维度1：可用与可修改性 ----------
    # 只判“交付物是否为一个能正常打开的 Excel 工作簿”；图表可编辑性、版面错乱等
    # 已分别由加分项（A01/A02/A11 等）与减分项（D01/D02 等）计分，此处不重复判 0。
    if broken:
        result["dim1_pass"] = False
        result["dim1_reason"] = ("维度一未通过：文件包结构异常，Excel 打开时会提示需要修复"
                                "（%s），完成度总分记 0" % broken[0])
        result["dim2_items"] = []
        result["total_score"] = 0
        wb.close()
        return result
    result["dim1_pass"] = True
    result["dim1_reason"] = "交付物为 1 个可正常打开的 Excel 工作簿（%s）" % _ext_of(file_path)
    # ---------- 图表侧派生事实 ----------
    cats_n = [_norm(c) for c in _CATS]
    col_groups = [g for _i, _f, g in groups if (g.get("barDir") or "col") == "col"]
    stacked = [g for g in col_groups if (g.get("grouping") or "") == "stacked"]
    all_ser = [s for _i, _f, g in groups for s in g.get("series", [])]
    minfo_ser = [s for g in ((minfo or {}).get("groups") or []) for s in g.get("series", [])]
    zg = _series_by_name(charts, "综管")
    fw = _series_by_name(charts, "服务")
    hy = _series_by_name(charts, "核验")
    gaps = [g.get("gapWidth") for _i, _f, g in groups]

    def cats_of(s):
        return [_norm(c) for c in (s.get("cats") or [])]

    def vals_of(s):
        return _nz(s.get("vals") or [])

    def shown(s):
        return s.get("show") == "1"

    def pos_of(s):
        return set(p for p in (s.get("pos") or set()) if p)

    def _tidy_source(info):
        """图表引用是否落在一张规整的绘图数据表上（5 行 × 每系列 1 列，表头紧邻上一行）。"""
        rows, cols, cat_cols, heads = set(), set(), set(), []
        for g in info.get("groups", []):
            for s in g.get("series", []):
                vr = _parse_ref(s.get("val_ref"))
                cr = _parse_ref(s.get("cat_ref"))
                nr = _parse_ref(s.get("name_ref"))
                if not vr or not cr or not nr or vr[2] != vr[4]:
                    return False
                rows.add((vr[0], vr[1], vr[3]))
                cols.add((vr[0], vr[2]))
                cat_cols.add((cr[0], cr[2], cr[1], cr[3]))
                heads.append((nr, vr))
        if len(rows) != 1 or len(cols) < 3 or len(cat_cols) != 1:
            return False
        sh, r1, r2 = rows.pop()
        if r2 - r1 + 1 != len(_CATS):
            return False
        csh, ccol, cr1, cr2 = cat_cols.pop()
        if csh != sh or (cr1, cr2) != (r1, r2) or (sh, ccol) in cols:
            return False
        for nr, vr in heads:
            if nr[0] != sh or nr[1] != r1 - 1 or nr[2] != vr[2]:
                return False
        return True
    def add(delta, hit, text):
        """加分项：命中记满分，未命中记 0。"""
        results.append("+%d：%s" % (delta if hit else 0, text))

    def ded(delta, trig, text):
        """减分项：触发破坏记该负分，未触发记 0。"""
        results.append("-%d：%s" % (delta if trig else 0, text))

    # ========== 维度2：完成度评分 —— 一、加分项 ==========

    # A01 新增柱形图图表对象
    add(3, bool(real_charts) and any(g.get("series") for g in col_groups),
        "工作簿中新增 1 个柱形图图表对象，用于登记数量与核验数量的对比展示")

    # A02 登记侧为综管科/服务科上下堆叠
    a02 = any(
        any("综管" in _norm(s.get("name")) for s in g.get("series", []))
        and any("服务" in _norm(s.get("name")) for s in g.get("series", []))
        for g in stacked)
    add(5, a02, "登记侧柱体为“综管科”“服务科”两段上下堆叠的堆积柱形图"
                "（非簇状并列、非百分比堆积、非条形图/折线图）")

    # A03 横轴 5 个类别、每类别 1 个标签
    #     要求所有带分类的系列都用这 5 项；出现拼接标签或重复类别文字即不得分
    ser_cats = [cats_of(s) for s in all_ser if cats_of(s)]
    a03 = bool(ser_cats) and all(cl == cats_n for cl in ser_cats)
    add(3, a03, "横轴分类为 5 个资源类别（便携学习终端/智能显示设备/网络接入设备/"
                "多功能文印设备/移动采集设备），每个类别只对应 1 个分类标签")

    # A04 核验数量为同刻度内并排的独立柱体
    a04 = bool(stacked) and any(
        all(g is not t for t in stacked)
        and (g.get("barDir") or "col") == "col"
        and (g.get("grouping") or "") != "stacked"
        and cats_of(s) == cats_n
        for g, s in hy)
    add(3, a04, "“核验数量”为独立柱体，与该类别的登记堆叠柱位于同一分类刻度内左右并排，"
                "且不叠加进堆叠柱内部")

    # A05 / A06 / A07 数据正确性
    a05 = (any(vals_of(s) == _ZG for _g, s in zg)
           and any(vals_of(s) == _FW for _g, s in fw))
    add(3, a05, "堆叠柱“综管科”段取值依次为 6、22、18、12、5，“服务科”段取值依次为 9、8、6、3、2")
    add(3, any(vals_of(s) == _HY for _g, s in hy),
        "“核验数量”柱取值依次为 13、27、22、14、6")
    a07 = False
    for g in stacked:
        cols5 = [vals_of(s) for s in g.get("series", [])
                 if ("综管" in _norm(s.get("name")) or "服务" in _norm(s.get("name")))]
        if len(cols5) >= 2 and all(len(c) == len(_CATS) for c in cols5):
            if [sum(c[i] for c in cols5) for i in range(len(_CATS))] == _TOT:
                a07 = True
    add(1, a07, "每根堆叠柱的总高度等于该类别登记合计 15、30、24、15、7")
    # A08 全部取数来自单元格区域引用
    a08 = bool(all_ser) and all(
        s.get("val_ref") and not s.get("val_lit")
        and s.get("cat_ref") and not s.get("cat_lit") and s.get("name_ref")
        for s in all_ser)
    add(5, a08, "图表的系列名称、分类轴、数值均通过单元格区域引用取数，修改数据源单元格后"
                "图表随之更新（系列数据以常量数组写死在图表内部则不得分）")

    # A09 规整绘图数据表
    a09 = bool(minfo) and _tidy_source(minfo)
    add(3, a09, "图表数据源为一张规整的绘图数据表（5 行数据 × 类别/综管科/服务科/核验数量各列），"
                "行列与图表系列一一对应，便于用户增删类别或系列")

    # A10 各系列取值均为真实数据，无拼并列用的占位 0
    a10 = bool(real_charts) and not any(
        sum(1 for v in (s.get("vals") or []) if v is not None and abs(v) < 1e-9) >= 2
        for s in all_ser)
    add(1, a10, "绘图数据源中各系列取值均为真实数据，不含为拼出并列效果而填入的占位 0，"
                "图例与数据标签中不出现无意义的 0")

    # A11 可逐系列调整（位图截图、单一图形对象、工作表锁定均不得分）
    locked = any(bool(w.protection.sheet) for w in wb.worksheets)
    a11 = has_chart_frame and n_series >= 2 and pic_count == 0 and not locked
    add(1, a11, "图表可直接选中并逐系列调整（可单独设置某一系列的填充色、标签、图表类型），"
                "非合并为单一图形对象，且工作表未被保护/锁定，用户可随时改动图表与数据源")

    # A12 图表标题
    title = _norm(minfo.get("title")) if minfo else ""
    a12 = bool(title) and "登记" in title and "核验" in title \
        and ("对比" in title or "对照" in title)
    add(1, a12, "图表标题为“固定资产登记数量（综管科+服务科）与核验数量对比”或含相同信息的中文标题")

    # A13 纵坐标轴标题
    ax_titles = [_norm(a.get("title")) for a in ((minfo or {}).get("valAx") or [])]
    add(1, any("数量" in t and "台" in t for t in ax_titles), "纵坐标轴带轴标题“数量（台）”")

    # A14 两段标签居中显示于柱体内部
    a14 = bool(zg) and bool(fw) and all(
        shown(s) and "ctr" in pos_of(s) for _g, s in zg + fw)
    add(3, a14, "“综管科”“服务科”两段均显示数据标签且标签位于各段柱体内部"
                "（居中显示 6/22/18/12/5 与 9/8/6/3/2）")

    # A15 堆叠柱顶合计标签 + 核验柱顶标签
    a15 = (any(shown(s) and vals_of(s) == _TOT for s in all_ser)
           and any(shown(s) and vals_of(s) == _HY for _g, s in hy))
    add(3, a15, "堆叠柱顶端显示登记合计标签 15、30、24、15、7，"
                "“核验数量”柱顶端显示 13、27、22、14、6")

    # A16 图例三项、不显示辅助系列
    #     c:legendEntry 的 idx 取“图例项出现顺序”，与 c:ser 的 idx 不一定一致，两者都接受
    deleted = set(minfo.get("legend_deleted") or []) if minfo else set()
    helper_hidden = True
    for pos, s in enumerate(minfo_ser):
        if s.get("fill") == "none" or "合计" in _norm(s.get("name")):
            if pos not in deleted and s.get("idx") not in deleted:
                helper_hidden = False
    a16 = (bool(minfo) and minfo.get("legend") and bool(zg) and bool(fw) and bool(hy)
           and helper_hidden)
    add(1, a16, "图例包含“综管科”“服务科”“核验数量”三项，"
                "且不显示仅用于承载合计标签的辅助系列名称")

    # A17 三色互不相同
    tri_fills = []
    for pairs in (zg, fw, hy):
        tri_fills.append(next((s.get("fill") for _g, s in pairs
                               if s.get("fill") not in (None, "none", "?")), None))
    a17 = all(tri_fills) and len(set(tri_fills)) == 3
    add(1, a17, "“综管科”“服务科”“核验数量”三者填充色明显互不相同，同类别两根柱可一眼区分")

    # A18 纵轴最大刻度
    ax_max = [a.get("max") for a in ((minfo or {}).get("valAx") or [])]
    a18 = bool(real_charts) and all(m is None or m >= 30 for m in ax_max)
    add(1, a18, "纵坐标轴最大刻度不小于 30（建议 40），所有柱体与数据标签完整显示、无截断")

    # A19 组内间距明显小于类别间距
    a19 = a04 and len(gaps) >= 2 and max(gaps) >= 100 and min(gaps) <= max(gaps) / 2.0
    add(1, a19, "同一类别的两根柱紧邻成组，组内间距明显小于类别之间的间距，成对比对关系可辨")

    # A20 摆放不遮挡台账与汇总区
    a20 = bool(real_charts)
    for an in anchors:
        if an["sheet"] != ws.title or not an["is_chart"] or not an["rect"]:
            continue
        if _rect_overlap(an["rect"], _RECT_MAIN) or _rect_overlap(an["rect"], _RECT_HELP):
            a20 = False
    add(1, a20, "图表尺寸与摆放位置不遮挡台账主表与汇总区的单元格内容，图表整体可完整查看")

    # A21 网格线只在数值轴（分类轴网格线在柱形图中是穿过柱体的竖线）
    a21 = bool(real_charts) and not any(
        ax.get("grid") for c in real_charts for ax in c["info"]["catAx"])
    add(1, a21, "网格线仅保留数值轴的横向刻度线用于读数，柱体不被分类轴的竖向网格线穿过遮挡，"
                "数据标签互不重叠、不压住柱体")

    # ========== 维度2：完成度评分 —— 二、减分项 ==========

    sum_rows = (5, 8, 11, 14, 17)
    idx_rows = (3, 6, 9, 12, 15)

    # D01 主表数据被删除/改写/移位
    ded(3, bool(main_bad), "主表中任一数据被删除、改写或移位（序号 1—5、5 个资源类别名称、"
                           "“综管科/服务科/合计”所属单元行、登记数量 6/9/22/8/18/6/12/3/5/2、"
                           "各类合计 15/30/24/15/7、核验数量 13/27/22/14/6）")

    # D02 主表区域插入或删除行/列（按行列位置形态判定，与表头文字改写互不重复）
    struct_ok = (all(_norm(ws.cell(r, 1).value) == "合计" for r in sum_rows)
                 and all(_num(ws.cell(r, 1).value) == float(i + 1)
                         for i, r in enumerate(idx_rows))
                 and all(_norm(ws.cell(r, 3).value) == "综管科" for r in idx_rows)
                 and all(_norm(ws.cell(r + 1, 3).value) == "服务科" for r in idx_rows)
                 and all(ws.cell(2, ci).value is not None for ci in range(1, 7)))
    ded(3, not struct_ok, "在主表区域插入或删除行/列，导致数据与合并单元格错位")

    # D03 A1 标题格式
    a1 = ws["A1"]
    d03 = not ((a1.font.name or "") == _TITLE_FONT and _num(a1.font.sz) == 22.0
               and bool(a1.font.b) and _rgb6(a1.font.color) == _C_WHITE
               and _fill6(a1) == _C_TITLE_FILL and "A1:F1" in merges and _centered(a1))
    ded(1, d03, "A1 标题“教学资源管理汇总台账”不满足——方正小标宋简体、22 磅、加粗、白色字、"
                "填充色 1F4E5F、A1:F1 合并且水平垂直居中")

    # D04 表头行 A2:F2 格式
    d04 = False
    for coord in sorted(_MAIN_HEADER):
        c = ws[coord]
        if not (bool(c.font.b) and _rgb6(c.font.color) == _C_WHITE
                and _fill6(c) == _C_HEAD_FILL and _centered(c)
                and _wrap_ok(c) and _thin4(c)):
            d04 = True
    ded(1, d04, "表头行 A2:F2（序号/资源类别/所属单元/登记数量（件）/核验数量（件）/备注）"
                "不满足——加粗白色字、填充色 6FA8A0、水平垂直居中、自动换行、四周细实线边框")

    # D05 数据行 A3:F16 格式（合计行另计）
    d05 = False
    for c in _iter_rect(ws, (3, 1, 16, 6)):
        if c.row in sum_rows or c.value is None or c.coordinate in covered:
            continue
        sz = _num(c.font.sz)
        if not (not bool(c.font.b) and (sz is None or sz in (11.0, 12.0))
                and _centered(c) and _fill6(c) == _C_BODY_FILL and _thin4(c)):
            d05 = True
    ded(1, d05, "数据行 A3:F16 不满足——11 磅常规字（登记数量列 12 磅）、水平垂直居中、"
                "填充色 E9F4F2、四周细实线边框")
    # D06 “合计”行格式
    d06 = False
    for r in sum_rows:
        for c in _iter_rect(ws, (r, 1, r, 6)):
            if c.value is None or c.coordinate in covered:
                continue
            if not (bool(c.font.b) and _rgb6(c.font.color) == _C_SUM_FONT
                    and _fill6(c) == _C_SUM_FILL and _centered(c) and _thin4(c)):
                d06 = True
    ded(1, d06, "“合计”行（第 5、8、11、14、17 行）不满足——加粗、字体颜色 1F2933、"
                "填充色 D8EBE8、居中对齐、四周细实线边框")

    # D07 主表 20 处合并
    ded(1, not set(_MAIN_MERGES) <= merges,
        "A3:A4、A6:A7、A9:A10、A12:A13、A15:A16 序号纵向合并，B3:B4、B6:B7、B9:B10、"
        "B12:B13、B15:B16 类别纵向合并，A5:C5、A8:C8、A11:C11、A14:C14、A17:C17 合计跨列合并，"
        "E3:E5、E6:E8、E9:E11、E12:E14、E15:E17 核验数量纵向合并被取消或范围被改变")

    # D08 列宽 / 行高
    d08 = False
    for col, w in _COL_WIDTHS.items():
        got = ws.column_dimensions[col].width if col in ws.column_dimensions else None
        if got is None or abs(float(got) - w) > 0.1:
            d08 = True
    for r, h in _ROW_HEIGHTS.items():
        got = ws.row_dimensions[r].height if r in ws.row_dimensions else None
        if got is None or abs(float(got) - h) > 0.6:
            d08 = True
    for r in range(3, 18):
        got = ws.row_dimensions[r].height if r in ws.row_dimensions else None
        if got is None or abs(float(got) - _ROW_HEIGHT_BODY) > 0.6:
            d08 = True
    ded(1, d08, "列宽 A 4.88、B 16.13、C 21.5、D 16.38、F 9.0 或行高"
                "（第 1 行 28.5、第 2 行 27.0、第 3 行及以下 26.0）被改动")

    # D09 主表表头文字被改写
    ded(1, bool(head_bad), "主表表头文字被改写为图表字段名（如把“登记数量（件）”“核验数量（件）”"
                           "改成“登记数量”“核验数量”或英文字段）")

    # D10 / D11 右侧汇总区取值
    ded(3, bool(htop_bad), "H2:M7 汇总表被删除或数值被改写（表头资源类别/综管科/服务科/合计/"
                           "核验数量/辅助列，及 5 个类别对应的 6·9·15·13、22·8·30·27、"
                           "18·6·24·22、12·3·15·14、5·2·7·6）")
    ded(3, bool(hbot_bad), "H10:M15 转置汇总表被删除或数值被改写（H10:M10 类别行，"
                           "综管科 6/22/18/12/5、服务科 9/8/6/3/2、合计 15/30/24/15/7、"
                           "核验数量 13/27/22/14/6 各行及“辅助列”行）")
    # D12 H2:M2 与 H10:M10 表头格式
    d12 = False
    for ch in "HIJKLM":
        for coord, want_fill in (("%s2" % ch, _C_HEAD_FILL), ("%s10" % ch, _C_BODY_FILL)):
            c = ws[coord]
            if not (bool(c.font.b) and _rgb6(c.font.color) == _C_WHITE
                    and _centered(c) and _wrap_ok(c) and _thin4(c)
                    and _fill6(c) == want_fill):
                d12 = True
    ded(1, d12, "H2:M2 与 H10:M10 表头不满足——加粗白色字、水平垂直居中、自动换行、"
                "四周细实线边框（H2:M2 填充色 6FA8A0，H10:M10 填充色 E9F4F2）")

    # D13 H13:M14 两行格式
    d13 = False
    for c in _iter_rect(ws, (13, 8, 14, 13)):
        if c.value is None:
            continue
        if not (bool(c.font.b) and _rgb6(c.font.color) == _C_SUM_FONT
                and _fill6(c) == _C_SUM_FILL and _centered(c)):
            d13 = True
    ded(1, d13, "H13:M14“合计”“核验数量”两行不满足——加粗、字体颜色 1F2933、"
                "填充色 D8EBE8、居中对齐")

    # D14 H15:M15“辅助列”行
    h15 = ws["H15"]
    d14 = not (_norm(h15.value) == "辅助列" and _fill6(h15) == _C_BODY_FILL and _thin4(h15))
    ded(1, d14, "H15:M15“辅助列”行被删除，或其 E9F4F2 填充与四周细实线边框被改动")

    # D15 承载台账的工作表
    d15 = not (ws.title == _SHEET_NAME and ws.sheet_state == "visible"
               and wb.sheetnames.index(ws.title) == 0)
    ded(1, d15, "承载台账的工作表 Sheet1 被改名、隐藏、删除或调整为非首个工作表")

    # D16 绘图数据表覆盖写入受保护区域
    d16 = False
    for rect in (_RECT_MAIN, _RECT_HELP):
        for c in _iter_rect(ws, rect):
            if c.value is not None and c.coordinate not in base_coords:
                d16 = True
    ded(1, d16, "新增的绘图数据表覆盖写入 A1:F17 或 H2:M15 所占单元格，挤占已有内容")

    # D17 残留多余/空白工作表
    used_sheets = set(an["sheet"] for an in anchors)
    d17 = False
    for w in wb.worksheets:
        if w is ws:
            continue
        if w.title in used_sheets:
            continue
        if not any(c.value is not None for row in w.iter_rows() for c in row):
            d17 = True
    ded(1, d17, "交付文件中残留与本次任务无关的多余工作表或空白工作表")
    # D18 残留未完成的中间图表或重复的第二个对比图
    d18 = len(real_charts) >= 2 or any(
        _norm(c["info"].get("title")) in ("图表标题", "ChartTitle") for c in real_charts)
    ded(1, d18, "交付文件中残留未完成的中间图表（如标题仍为“图表标题”的默认簇状柱形图）"
                "或功能重复的第二个对比图表")

    # D19 自造内容
    d19 = False
    for w in wb.worksheets:
        for row in w.iter_rows():
            for c in row:
                if c.value is None or _num(c.value) is not None:
                    continue
                if w is ws and c.coordinate in base_coords:
                    continue
                t = _norm(c.value)
                if t in _ALLOWED_TEXT:
                    continue
                if "登记" in t and "核验" in t and "对比" in t:   # 图表标题文字
                    continue
                d19 = True
    ded(1, d19, "绘图数据表中出现指令未要求的自造内容（如“差异”列、虚构科室、虚构数值、"
                "拼接出的组合标签列）")

    # D20 默认打开视图
    sv = ws.sheet_view
    zoom = sv.zoomScale
    d20 = bool(ws.freeze_panes) or (sv.showGridLines is False) \
        or (zoom is not None and int(zoom) != 100)
    try:
        if wb.views and wb.views[0].activeTab not in (None, 0):
            d20 = True
    except Exception:
        pass
    ded(1, d20, "工作簿的默认打开视图被改动（新增冻结窗格、隐藏网格线、修改缩放比例或"
                "改变活动工作表定位）")

    # D21 新增公式/批注/条件格式/数据验证
    d21 = False
    for rect in (_RECT_MAIN, _RECT_HELP):
        for c in _iter_rect(ws, rect):
            if getattr(c, "data_type", None) == "f" or getattr(c, "comment", None):
                d21 = True
    for holder in (ws.data_validations.dataValidation, list(ws.conditional_formatting)):
        for item in holder:
            sq = getattr(item, "sqref", None)
            for rng in (getattr(sq, "ranges", []) or []):
                rr = (rng.min_row, rng.min_col, rng.max_row, rng.max_col)
                if _rect_overlap(rr, _RECT_MAIN) or _rect_overlap(rr, _RECT_HELP):
                    d21 = True
    ded(1, d21, "在台账主表或汇总辅助区单元格中新增指令未要求的公式、批注、条件格式或数据验证")

    wb.close()

    # ---------- 汇总 ----------
    if len(results) != len(_MAX_DELTAS_ORDER):
        raise RuntimeError("结果条目数(%d)与 max_delta 序列(%d)不一致"
                           % (len(results), len(_MAX_DELTAS_ORDER)))
    total = 0
    dim2_items = []
    for raw, max_delta in zip(results, _MAX_DELTAS_ORDER):
        head, _sep, rule = raw.partition("：")
        got = int(head.replace("+", "").replace("-", "").strip())
        hit = got > 0
        delta = max_delta if hit else 0
        total += delta
        dim2_items.append({"rule": rule, "max_delta": max_delta, "delta": delta,
                           "hit": hit})
    result["dim2_items"] = dim2_items
    result["total_score"] = total
    return result


if __name__ == "__main__":
    _dir = sys.argv[1] if len(sys.argv) > 1 else os.path.dirname(os.path.abspath(__file__))
    print(json.dumps(evaluate(_dir), ensure_ascii=False, indent=2))
