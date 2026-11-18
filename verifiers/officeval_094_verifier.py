#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# pyright: reportUnknownMemberType=false, reportUnknownVariableType=false, reportUnknownArgumentType=false, reportUnknownParameterType=false, reportMissingParameterType=false, reportPrivateUsage=false, reportUnknownLambdaType=false, reportGeneralTypeIssues=false

import os
import re
import sys
import json
import zipfile
import colorsys
import posixpath
from xml.etree import ElementTree as ET

# 说明：按接口统一约定，本模块不修改全局 sys.stdout，也不 print 主结果。
# 主结果统一由 evaluate(dir_path) -> dict 返回；入参为"脚本所在目录的路径"，
# 由脚本自己在该目录内定位并打开被评估的 .xlsx / .xlsm。本文件仅在 __main__ 下作为
# 本地调试入口，才把 dict 序列化为 JSON 打印。
#
# 评分语义：
#   维度1（可用与可修改性）：硬性门槛。任一条不满足 -> 直接 0 分，不再检查维度2。
#   维度2：加分点必须满足【每一个】子条件才加分；减分点只要命中【任意一个】
#          子条件即扣分（分数为负）。最终得分 = 各命中细则分数之和（可正可负）。
#
# 实现说明：不依赖 openpyxl —— 被评估文件里的圆环图 holeSize 可能为 0，
# openpyxl 的描述符校验会直接抛 ValueError 拒绝加载，故全部走 zipfile +
# ElementTree 直接解析 SpreadsheetML / DrawingML。

SCRIPT_ID = "094"
# 维度二所有加分项 max_delta 之和；用于维度一未通过 / 脚本异常时兜底
MAX_SCORE_TOTAL = 50

# 维度二所有评分项按"结果追加顺序"的 max_delta 列表；末端把 results 里的
# "+N：..." / "-N：..." 字符串反解为结构化 dim2_items 时用。顺序须与
# _evaluate_impl 中 emit(...) 的顺序严格一致；共 34 项 = 18 加分 + 16 减分。
_MAX_DELTAS_ORDER = [
    # ---- 一、加分项（合计 +50）----
    3,   # 01 两表各内嵌 1 张图、共 2 张
    1,   # 02 图表标题可区分班级
    5,   # 03 饼/环形族双层同心嵌套
    3,   # 04 圈心孔径为 0 或接近 0
    5,   # 05 内层为六个探究表现大类或四个能力层级大类
    5,   # 06 外层 24 个明细扇区、按大类连续排列、内外角度对齐
    3,   # 07 只有两个层级、无第三层圈层
    3,   # 08 内层数值等于大类之和、外层组内之和等于内层
    3,   # 09 数据源为单元格区域引用
    3,   # 10 汇总区域未占用 A1:Q7
    3,   # 11 汇总数值由公式计算
    1,   # 12 汇总区域带可读行列表头
    3,   # 13 内外层同色系
    3,   # 14 内外两层均显示数据标签
    1,   # 15 扇区之间有描边
    1,   # 16 图表带图例
    3,   # 17 尺寸不小于约 15cm×10cm 且标签不重叠不截断
    1,   # 18 图表与数据同表且未压盖数据表与汇总区
    # ---- 二、减分项（合计 -26）----
    -1,  # 19 第 1 行合并表头（含被覆盖格）填充/字体/对齐/边框被重置
    -1,  # 20 第 2 行 16 个能力层级单元格格式
    -1,  # 21 A 列 7 个标签单元格格式
    -3,  # 22 第 1、2 行行高 15，第 3~7 行行高 25
    -1,  # 23 数据区以外空白行逐行自定义行高
    -1,  # 24 A 列列宽 24.17、B~Q 列列宽 8
    -1,  # 25 S~AB 等空白列新增自定义列宽
    -3,  # 26 A1:Q7 内文字或数值被增删改写
    -3,  # 27 六个合并单元格被取消或范围改变
    -3,  # 28 删除工作表或改变名称与顺序
    -3,  # 29 删除或破坏已有图表/图片/批注等对象
    -1,  # 30 "不显示网格线"设置被改为显示
    -1,  # 31 单元格内新增与图表标题重复的文字标题
    -1,  # 32 中间计算区域铺开超过 30 行
    -1,  # 33 图表压盖 A1:Q7 或压盖汇总区域
    -1,  # 34 新增无关工作表、隐藏行列或定义名称
]
assert sum(d for d in _MAX_DELTAS_ORDER if d > 0) == MAX_SCORE_TOTAL

# ---------------------------------------------------------------- 命名空间
SS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
C = "http://schemas.openxmlformats.org/drawingml/2006/chart"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
XDR = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
REL = "http://schemas.openxmlformats.org/package/2006/relationships"
NS = {"ss": SS, "c": C, "a": A, "r": R, "xdr": XDR, "rel": REL}

# ------------------------------------------------------- 成绩表基准（A1:Q7）
SHEET_NAMES = ("七1班", "七13班")
# 第 1 行六个探究表现大类，及其在数据表中占据的列区间（1 基，含端点）
CATEGORIES = ("地图基础", "经纬定位", "地形判读", "气候要素", "区域差异", "聚落形态")
CAT_SPANS = ((2, 4), (5, 7), (8, 9), (10, 11), (12, 14), (15, 17))
# 第 2 行 16 个能力层级
ABILITY_NAMES = ("识图", "判读", "表达", "观察", "分析", "归纳", "定位", "解释",
                 "识别", "比较", "计算", "推理", "应用", "观察", "判断", "表达")
# A 列 7 个标签；其中后四个同时是四个能力层级大类
A_LABELS = ("探究表现", "能力层级", "参与人数", "卓越表现", "稳定达成", "基本达成", "持续提升")
LEVELS = ("卓越表现", "稳定达成", "基本达成", "持续提升")
MERGE_REFS = ("B1:D1", "E1:G1", "H1:I1", "J1:K1", "L1:N1", "O1:Q1")

# 第 3~7 行共 80 个整数（行号 -> B..Q 的 16 个值）
DATA_BASELINE = {
    "七1班": {
        3: (52, 48, 53, 49, 54, 50, 55, 51, 47, 52, 54, 49, 54, 51, 47, 52),
        4: (30, 24, 23, 35, 36, 29, 28, 22, 33, 34, 27, 21, 38, 26, 20, 37),
        5: (11, 12, 14, 7, 10, 11, 13, 14, 7, 9, 13, 13, 8, 12, 13, 8),
        6: (7, 8, 10, 4, 5, 7, 9, 10, 4, 5, 9, 9, 5, 8, 9, 5),
        7: (4, 4, 6, 3, 3, 3, 5, 5, 3, 4, 5, 6, 3, 5, 5, 2),
    },
    "七13班": {
        3: (50, 55, 51, 47, 53, 48, 53, 49, 54, 50, 51, 47, 52, 49, 54, 50),
        4: (22, 39, 34, 27, 27, 21, 38, 32, 31, 25, 36, 31, 30, 35, 36, 29),
        5: (14, 8, 9, 10, 13, 13, 8, 9, 12, 12, 8, 8, 11, 7, 10, 11),
        6: (10, 5, 5, 6, 8, 9, 5, 5, 7, 8, 5, 5, 7, 4, 5, 7),
        7: (4, 3, 3, 4, 5, 5, 2, 3, 4, 5, 2, 3, 4, 3, 3, 3),
    },
}
# 以探究表现作内层时六个扇区的应有数值
INNER_BY_CATEGORY = {"七1班": (153, 153, 106, 99, 157, 150),
                     "七13班": (156, 148, 102, 104, 150, 153)}
# 以能力层级作内层时四个扇区的应有数值
INNER_BY_LEVEL = {"七1班": (463, 175, 114, 66),
                  "七13班": (493, 163, 101, 56)}

# ------------------------------------------------- 表头/标签格式基准（逐格）
HEAD_FILL = "FF174A55"          # 第 1 行合并表头填充
HEAD_FONT = ("黑体", "11", True, "FFFFFFFF")
ROW2_FILL = "FF2E6F73"          # 第 2 行填充
ROW2_FONT = ("宋体", "11", True, "FFFFFFFF")
# 第 1 行 B1..Q1 的（左竖边, 右竖边）样式；上下边一律 thin
ROW1_VERT = {2: ("medium", "thin"), 3: ("thin", "thin"), 4: ("thin", "medium"),
             5: ("medium", "thin"), 6: ("thin", "thin"), 7: ("thin", "medium"),
             8: ("medium", "thin"), 9: ("thin", "medium"),
             10: ("medium", "thin"), 11: ("thin", "medium"),
             12: ("thin", None), 13: (None, None), 14: (None, "medium"),
             15: ("medium", "thin"), 16: ("thin", "thin"), 17: ("thin", "medium")}
# 第 2 行 B2..Q2 的（左竖边, 右竖边）样式；上下边一律 thin
ROW2_VERT = {2: ("medium", "thin"), 3: ("thin", "thin"), 4: ("thin", "medium"),
             5: ("medium", "thin"), 6: ("thin", "thin"), 7: ("thin", "medium"),
             8: ("medium", "thin"), 9: ("thin", "medium"),
             10: ("medium", "thin"), 11: ("thin", "medium"),
             12: ("thin", "thin"), 13: ("thin", "thin"), 14: ("thin", "medium"),
             15: ("medium", "thin"), 16: ("thin", "thin"), 17: ("thin", "medium")}
# A 列 A1..A7 的（字体, 填充, 水平, 垂直, 上边, 下边）；左右竖边一律 medium
ACOL_EXPECT = {
    1: (("黑体", "11", True, "FFFFFFFF"), "FF174A55", "center", "center", "thin", "thin"),
    2: (("黑体", "11", True, "FFFFFFFF"), "FF2E6F73", "center", "center", "thin", "thin"),
    3: (("黑体", "11", True, "FF23464D"), "FFFFF1CC", "left", "center", "thin", "thin"),
    4: (("黑体", "11", True, "FF23464D"), "FFDCEFE7", "left", "center", "thin", "thin"),
    5: (("黑体", "11", True, "FF23464D"), "FFDDEBF7", "left", "center", "thin", "thin"),
    6: (("黑体", "11", True, "FF23464D"), "FFFFF4D6", "left", "center", "thin", "thin"),
    7: (("黑体", "11", True, "FF23464D"), "FFFCE4D6", "left", "center", "thin", "thin"),
}
ROW_HEIGHTS = {1: 15.0, 2: 15.0, 3: 25.0, 4: 25.0, 5: 25.0, 6: 25.0, 7: 25.0}
COL_WIDTH_A = 24.1666666666667
COL_WIDTH_BQ = 8.0
WIDTH_TOL = 0.6        # 列宽/行高容差：0.6 字符内显示效果一致，不计改动
HEIGHT_TOL = 0.6
DATA_ROWS = 7          # 数据表占 1~7 行
DATA_COLS = 17         # 数据表占 A~Q 列
EMU_PER_PX = 9525.0
EMU_PER_CM = 360000.0
MIN_CHART_CX = 15.0 * EMU_PER_CM     # 图表最小宽度 15cm
MIN_CHART_CY = 10.0 * EMU_PER_CM     # 图表最小高度 10cm
MAX_READABLE_POINTS = 24             # 单圈扇区数上限；超过则标签必然重叠
HUE_TOL = 20.0                       # "同色系"允许的色相偏差（度）
GRAY_SAT = 0.12                      # 饱和度低于此值视为无彩色（灰阶）


# ================================================================ 基础工具


def _col_index(letters):
    """列字母 -> 1 基列号；'A'->1, 'AB'->28。"""
    n = 0
    for ch in letters:
        n = n * 26 + (ord(ch.upper()) - 64)
    return n


def _col_name(idx):
    """1 基列号 -> 列字母。"""
    s = ""
    while idx:
        idx, rem = divmod(idx - 1, 26)
        s = chr(65 + rem) + s
    return s


def _cell_addr(row, col):
    return "%s%d" % (_col_name(col), row)


def _resolve_part(base_part, target):
    """按 OPC 规则把关系里的 Target 解析成包内绝对部件名（不带前导斜杠）。"""
    if target.startswith("/"):
        return target.lstrip("/")
    return posixpath.normpath(posixpath.join(posixpath.dirname(base_part), target))


def _norm_text(text):
    """去掉全部空白字符，用于标题/标签的宽松比对（WPS 会在中文数字间插空格）。"""
    return re.sub(r"\s+", "", text or "")


def _norm_color(elem):
    """把 <fgColor>/<color>/<srgbClr> 之类的颜色元素规格化成可比字符串。"""
    if elem is None:
        return None
    rgb = elem.get("rgb") or elem.get("val")
    if rgb:
        rgb = rgb.upper()
        return rgb if len(rgb) == 8 else "FF" + rgb
    if elem.get("theme") is not None:
        return "theme:%s" % elem.get("theme")
    if elem.get("indexed") is not None:
        return "indexed:%s" % elem.get("indexed")
    return None


def _bool_attr(elem):
    """OOXML 布尔子元素：元素存在且 val 不为 0/false 即为真。"""
    if elem is None:
        return False
    return (elem.get("val") or "1") not in ("0", "false", "False")


# ============================================================ 工作簿解析


def _rels_of(zf, part):
    """读取某部件的 .rels，返回 {rId: 包内绝对部件名}。"""
    rels_part = posixpath.join(posixpath.dirname(part), "_rels",
                               posixpath.basename(part) + ".rels")
    out = {}
    try:
        root = ET.fromstring(zf.read(rels_part))
    except (KeyError, ET.ParseError):
        return out
    for rel in root.findall("rel:Relationship", NS):
        out[rel.get("Id")] = (rel.get("Target") if rel.get("TargetMode") == "External"
                              else _resolve_part(part, rel.get("Target") or ""))
    return out


def _load_styles(zf):
    """解析 styles.xml，返回 cellXfs 下标 -> 规格化格式字典 的列表。"""
    try:
        root = ET.fromstring(zf.read("xl/styles.xml"))
    except (KeyError, ET.ParseError):
        return []
    fonts = root.findall("ss:fonts/ss:font", NS)
    fills = root.findall("ss:fills/ss:fill", NS)
    borders = root.findall("ss:borders/ss:border", NS)
    styles = []
    for xf in root.findall("ss:cellXfs/ss:xf", NS):
        info = {"font": (None, None, False, None), "fill": (None, None),
                "border": {}, "align": (None, None)}
        fi = int(xf.get("fontId") or 0)
        if fi < len(fonts):
            font = fonts[fi]
            name = font.find("ss:name", NS)
            size = font.find("ss:sz", NS)
            info["font"] = (name.get("val") if name is not None else None,
                            size.get("val") if size is not None else None,
                            _bool_attr(font.find("ss:b", NS)),
                            _norm_color(font.find("ss:color", NS)))
        fli = int(xf.get("fillId") or 0)
        if fli < len(fills):
            pat = fills[fli].find("ss:patternFill", NS)
            info["fill"] = ((pat.get("patternType") if pat is not None else None),
                            _norm_color(pat.find("ss:fgColor", NS) if pat is not None else None))
        bi = int(xf.get("borderId") or 0)
        if bi < len(borders):
            for side in ("left", "right", "top", "bottom"):
                sub = borders[bi].find("ss:%s" % side, NS)
                info["border"][side] = sub.get("style") if sub is not None else None
        align = xf.find("ss:alignment", NS)
        if align is not None:
            info["align"] = (align.get("horizontal"), align.get("vertical"))
        styles.append(info)
    return styles


def _load_sheet(zf, part, strings):
    """解析一张工作表，返回 {cells, rows, cols, merges, gridlines, fmt, drawing_rid}。"""
    root = ET.fromstring(zf.read(part))
    view = root.find("ss:sheetViews/ss:sheetView", NS)
    fmt = root.find("ss:sheetFormatPr", NS)
    sheet = {
        "part": part,
        "cells": {},
        "rows": {},
        "cols": [],
        "merges": [],
        "gridlines": True,
        "default_width": 8.43,
        "default_height": 15.0,
        "drawing_rid": None,
        "hidden_rc": False,
    }
    if view is not None and (view.get("showGridLines") or "1") in ("0", "false"):
        sheet["gridlines"] = False
    if fmt is not None:
        if fmt.get("defaultColWidth"):
            sheet["default_width"] = float(fmt.get("defaultColWidth"))
        if fmt.get("defaultRowHeight"):
            sheet["default_height"] = float(fmt.get("defaultRowHeight"))
    for col in root.findall("ss:cols/ss:col", NS):
        sheet["cols"].append({
            "min": int(col.get("min") or 1), "max": int(col.get("max") or 1),
            "width": float(col.get("width")) if col.get("width") else None,
            "custom": (col.get("customWidth") or "0") not in ("0", "false"),
            "hidden": (col.get("hidden") or "0") not in ("0", "false"),
        })
        if sheet["cols"][-1]["hidden"]:
            sheet["hidden_rc"] = True
    merge = root.find("ss:mergeCells", NS)
    if merge is not None:
        sheet["merges"] = [m.get("ref") for m in merge.findall("ss:mergeCell", NS)]
    for row in root.findall("ss:sheetData/ss:row", NS):
        rnum = int(row.get("r") or 0)
        hidden = (row.get("hidden") or "0") not in ("0", "false")
        if hidden:
            sheet["hidden_rc"] = True
        sheet["rows"][rnum] = {
            "ht": float(row.get("ht")) if row.get("ht") else None,
            "custom": (row.get("customHeight") or "0") not in ("0", "false"),
            "hidden": hidden,
        }
        for cell in row.findall("ss:c", NS):
            match = re.match(r"([A-Za-z]+)(\d+)", cell.get("r") or "")
            if not match:
                continue
            sheet["cells"][(int(match.group(2)), _col_index(match.group(1)))] = \
                _read_cell(cell, strings)
    drawing = root.find("ss:drawing", NS)
    if drawing is not None:
        sheet["drawing_rid"] = drawing.get("{%s}id" % R)
    return sheet


def _read_cell(cell, strings):
    """把 <c> 读成 {text, number, style, has_formula}。"""
    ctype = cell.get("t")
    vnode = cell.find("ss:v", NS)
    fnode = cell.find("ss:f", NS)
    inode = cell.find("ss:is", NS)
    text, number = None, None
    if ctype == "s" and vnode is not None:
        try:
            text = strings[int(vnode.text)]
        except (ValueError, IndexError, TypeError):
            text = vnode.text
    elif ctype == "inlineStr" and inode is not None:
        text = "".join(t.text or "" for t in inode.iter("{%s}t" % SS))
    elif ctype == "str" and vnode is not None:
        text = vnode.text
    elif vnode is not None:
        text = vnode.text
        try:
            number = float(vnode.text)
        except (TypeError, ValueError):
            number = None
    return {"text": text, "number": number, "style": cell.get("s"),
            "has_formula": fnode is not None}


def _load_package(file_path):
    """打开 .xlsx / .xlsm 并解析出评分所需的全部上下文。"""
    zf = zipfile.ZipFile(file_path)
    names = set(zf.namelist())
    ctx = {"zip": zf, "names": names, "file_path": file_path,
           "sheets": {}, "sheet_order": [], "charts": [], "pictures": 0,
           "defined_names": []}
    # 共享字符串
    strings = []
    if "xl/sharedStrings.xml" in names:
        sst = ET.fromstring(zf.read("xl/sharedStrings.xml"))
        for item in sst.findall("ss:si", NS):
            strings.append("".join(t.text or "" for t in item.iter("{%s}t" % SS)))
    ctx["strings"] = strings
    ctx["styles"] = _load_styles(zf)
    # 工作表清单
    wb = ET.fromstring(zf.read("xl/workbook.xml"))
    wb_rels = _rels_of(zf, "xl/workbook.xml")
    for node in wb.findall("ss:sheets/ss:sheet", NS):
        name = node.get("name")
        part = wb_rels.get(node.get("{%s}id" % R))
        ctx["sheet_order"].append(name)
        if part and part in names:
            ctx["sheets"][name] = _load_sheet(zf, part, strings)
            ctx["sheets"][name]["state"] = node.get("state") or "visible"
    for node in wb.findall("ss:definedNames/ss:definedName", NS):
        ctx["defined_names"].append(node.get("name"))
    _load_drawings(ctx)
    return ctx


def _anchor_cell(node, tag):
    """读取 <xdr:from>/<xdr:to>，返回 (col, colOff, row, rowOff)，均为 0 基/EMU。"""
    sub = node.find("xdr:%s" % tag, NS)
    if sub is None:
        return None
    def _val(key):
        elem = sub.find("xdr:%s" % key, NS)
        try:
            return int(elem.text)
        except (AttributeError, TypeError, ValueError):
            return 0
    return (_val("col"), _val("colOff"), _val("row"), _val("rowOff"))


def _load_drawings(ctx):
    """把每张工作表的绘图锚点解析成图表对象；同时统计图片对象数量。"""
    zf = ctx["zip"]
    for name, sheet in ctx["sheets"].items():
        if not sheet["drawing_rid"]:
            continue
        part = _rels_of(zf, sheet["part"]).get(sheet["drawing_rid"])
        if not part or part not in ctx["names"]:
            continue
        droot = ET.fromstring(zf.read(part))
        drels = _rels_of(zf, part)
        for anchor in list(droot):
            tag = anchor.tag.split("}")[-1]
            if tag not in ("twoCellAnchor", "oneCellAnchor", "absoluteAnchor"):
                continue
            if anchor.find(".//xdr:pic", NS) is not None:
                ctx["pictures"] += 1
            chart_rid = None
            for gframe in anchor.iter("{%s}chart" % C):
                chart_rid = gframe.get("{%s}id" % R)
                break
            if not chart_rid:
                continue
            cpart = drels.get(chart_rid)
            if not cpart or cpart not in ctx["names"]:
                continue
            ext = anchor.find("xdr:ext", NS)
            ctx["charts"].append({
                "sheet": name,
                "part": cpart,
                "anchor_tag": tag,
                "from": _anchor_cell(anchor, "from"),
                "to": _anchor_cell(anchor, "to"),
                "ext": ((int(ext.get("cx") or 0), int(ext.get("cy") or 0))
                        if ext is not None else None),
                "root": ET.fromstring(zf.read(cpart)),
                "theme": _chart_theme(ctx, cpart),
            })
    ctx["charts"].sort(key=lambda ch: (ctx["sheet_order"].index(ch["sheet"])
                                       if ch["sheet"] in ctx["sheet_order"] else 99,
                                       ch["part"]))


def _chart_theme(ctx, chart_part):
    """解析该图表实际生效的配色方案：图表自带 themeOverride 优先，其次 theme1。
    schemeClr 必须按这份方案解析，否则会把"同色系"误判为不同色。"""
    part = None
    for target in _rels_of(ctx["zip"], chart_part).values():
        if "themeOverride" in target and target in ctx["names"]:
            part = target
            break
    if part is None and "xl/theme/theme1.xml" in ctx["names"]:
        part = "xl/theme/theme1.xml"
    scheme = {}
    if part is None:
        return scheme
    try:
        root = ET.fromstring(ctx["zip"].read(part))
    except (KeyError, ET.ParseError):
        return scheme
    clr = root.find(".//a:clrScheme", NS)
    if clr is None:
        return scheme
    for slot in list(clr):
        key = slot.tag.split("}")[-1]
        srgb = slot.find("a:srgbClr", NS)
        sysc = slot.find("a:sysClr", NS)
        if srgb is not None and srgb.get("val"):
            scheme[key] = srgb.get("val").upper()
        elif sysc is not None and sysc.get("lastClr"):
            scheme[key] = sysc.get("lastClr").upper()
    # OOXML 别名：主题里叫 dk1/lt1/dk2/lt2，图形里引用 tx1/bg1/tx2/bg2
    for alias, key in (("tx1", "dk1"), ("bg1", "lt1"), ("tx2", "dk2"), ("bg2", "lt2")):
        if key in scheme:
            scheme.setdefault(alias, scheme[key])
    return scheme


# ================================================== 单元格 / 区域引用工具


def _cell(ctx, sheet_name, row, col):
    sheet = ctx["sheets"].get(sheet_name)
    if sheet is None:
        return None
    return sheet["cells"].get((row, col))


def _style_of(ctx, cell):
    """取单元格生效格式；无格式索引时返回 cellXfs[0]。"""
    idx = 0
    if cell is not None and cell.get("style") is not None:
        try:
            idx = int(cell["style"])
        except ValueError:
            idx = 0
    if 0 <= idx < len(ctx["styles"]):
        return ctx["styles"][idx]
    return {"font": (None, None, False, None), "fill": (None, None),
            "border": {}, "align": (None, None)}


_REF_RE = re.compile(r"(?:'([^']+)'|([^!,()]+))!(\$?[A-Za-z]+\$?\d+)(?::(\$?[A-Za-z]+\$?\d+))?")


def _iter_ref_cells(ref, default_sheet=None):
    """解析图表引用（可能是 (A!$C$1:$F$1,A!$C$2:$F$2) 这样的多区域并集），
    按出现顺序产出 (sheet, row, col)。"""
    if not ref:
        return
    for match in _REF_RE.finditer(ref):
        sheet = match.group(1) or match.group(2) or default_sheet
        sheet = (sheet or "").strip()
        start = match.group(3).replace("$", "")
        end = (match.group(4) or match.group(3)).replace("$", "")
        m1 = re.match(r"([A-Za-z]+)(\d+)", start)
        m2 = re.match(r"([A-Za-z]+)(\d+)", end)
        if not m1 or not m2:
            continue
        c1, r1 = _col_index(m1.group(1)), int(m1.group(2))
        c2, r2 = _col_index(m2.group(1)), int(m2.group(2))
        for row in range(min(r1, r2), max(r1, r2) + 1):
            for col in range(min(c1, c2), max(c1, c2) + 1):
                yield (sheet, row, col)


def _ref_hits_data_area(ref, default_sheet=None):
    """引用是否落进 A1:Q7 数据表区域。"""
    for _, row, col in _iter_ref_cells(ref, default_sheet):
        if row <= DATA_ROWS and col <= DATA_COLS:
            return True
    return False


# ==================================================== 图表结构与配色解析


def _series_of(chart):
    """取出图表内饼/环形族的全部系列，按 XML 顺序（内层在前）返回。"""
    out = []
    for tag in ("doughnutChart", "pieChart", "ofPieChart", "pie3DChart"):
        for plot in chart["root"].findall(".//c:%s" % tag, NS):
            for ser in plot.findall("c:ser", NS):
                out.append({"ser": ser, "plot_tag": tag, "plot": plot})
    return out


def _ref_of(ser, kind):
    """取系列的 cat / val 区域引用字符串。"""
    node = ser.find("c:%s" % kind, NS)
    if node is None:
        return None
    f = node.find(".//c:f", NS)
    return f.text if f is not None else None


def _cached_points(ser, kind):
    """取系列缓存的 <c:pt> 文本列表（openpyxl 生成的图表通常没有缓存）。"""
    node = ser.find("c:%s" % kind, NS)
    if node is None:
        return []
    return [(pt.text or "") for pt in node.findall(".//c:pt/c:v", NS)]


def _series_values(ctx, chart, ser):
    """系列数值：优先按引用回表取单元格实算值，取不到再退回图表缓存。"""
    ref = _ref_of(ser, "val")
    values = []
    if ref:
        for sheet, row, col in _iter_ref_cells(ref, chart["sheet"]):
            cell = _cell(ctx, sheet, row, col)
            values.append(cell["number"] if cell and cell["number"] is not None else None)
    if not values or all(v is None for v in values):
        values = []
        for raw in _cached_points(ser, "val"):
            try:
                values.append(float(raw))
            except ValueError:
                values.append(None)
    return values


def _series_labels(ctx, chart, ser):
    """系列分类名：优先按引用回表取文本，取不到再退回图表缓存。"""
    ref = _ref_of(ser, "cat")
    labels = []
    if ref:
        for sheet, row, col in _iter_ref_cells(ref, chart["sheet"]):
            cell = _cell(ctx, sheet, row, col)
            labels.append((cell["text"] or "") if cell else "")
    if not any(labels):
        labels = _cached_points(ser, "cat")
    return [_norm_text(x) for x in labels]


def _fill_base_rgb(spPr, theme):
    """从 spPr/solidFill 里取"基色"RGB。lumMod/lumOff 只改明度不改色相，
    判定"同色系"时按基色比较即可。返回 'RRGGBB' 或 None。"""
    if spPr is None:
        return None
    fill = spPr.find("a:solidFill", NS)
    if fill is None:
        return None
    srgb = fill.find("a:srgbClr", NS)
    if srgb is not None and srgb.get("val"):
        return srgb.get("val").upper()
    scheme = fill.find("a:schemeClr", NS)
    if scheme is not None and scheme.get("val"):
        return theme.get(scheme.get("val"))
    return None


def _dpt_colors(ser, theme, count):
    """按数据点下标返回长度为 count 的基色列表；未单独设色的位置为 None。"""
    colors = [None] * count
    for dpt in ser.findall("c:dPt", NS):
        idx_node = dpt.find("c:idx", NS)
        try:
            idx = int(idx_node.get("val"))
        except (AttributeError, TypeError, ValueError):
            continue
        if 0 <= idx < count:
            colors[idx] = _fill_base_rgb(dpt.find("c:spPr", NS), theme)
    return colors


def _hue_sat(rgb):
    """'RRGGBB' -> (色相角度, 饱和度)；解析失败返回 None。"""
    if not rgb or len(rgb) < 6:
        return None
    try:
        r = int(rgb[-6:-4], 16) / 255.0
        g = int(rgb[-4:-2], 16) / 255.0
        b = int(rgb[-2:], 16) / 255.0
    except ValueError:
        return None
    hue, _, sat = colorsys.rgb_to_hsv(r, g, b)
    return (hue * 360.0, sat)


def _same_color_family(rgb_a, rgb_b):
    """同色系判定：都是灰阶算同系；否则色相差在 HUE_TOL 内算同系。"""
    ha, hb = _hue_sat(rgb_a), _hue_sat(rgb_b)
    if ha is None or hb is None:
        return False
    if ha[1] < GRAY_SAT and hb[1] < GRAY_SAT:
        return True
    if ha[1] < GRAY_SAT or hb[1] < GRAY_SAT:
        return False
    diff = abs(ha[0] - hb[0]) % 360.0
    return min(diff, 360.0 - diff) <= HUE_TOL


def _chunk_align(inner, outer, tol=0.5):
    """把外层数值按顺序切成若干连续块，使第 i 块之和等于内层第 i 个扇区的数值。
    切得动说明"外层按大类顺序连续排列、每个内层扇区正对其下属外层扇区"；
    返回块列表（每块是外层下标区间 [start, end)），切不动返回 None。"""
    if not inner or not outer or any(v is None for v in inner + outer):
        return None
    chunks = []
    pos = 0
    for target in inner:
        acc = 0.0
        start = pos
        while pos < len(outer) and acc + tol < target:
            acc += outer[pos]
            pos += 1
        if abs(acc - target) > tol or pos == start:
            return None
        chunks.append((start, pos))
    return chunks if pos == len(outer) else None


def _chart_title(ctx, chart):
    """图表标题文本（去空白）；支持富文本与引用单元格两种写法。"""
    title = chart["root"].find(".//c:chart/c:title", NS)
    if title is None:
        return ""
    text = "".join(t.text or "" for t in title.iter("{%s}t" % A))
    if not text:
        ref = None
        f = title.find(".//c:strRef/c:f", NS)
        if f is not None:
            ref = f.text
        if ref:
            parts = []
            for sheet, row, col in _iter_ref_cells(ref, chart["sheet"]):
                cell = _cell(ctx, sheet, row, col)
                if cell and cell["text"]:
                    parts.append(cell["text"])
            text = "".join(parts)
    return _norm_text(text)


def _shows_values(plot, ser):
    """该系列的扇区是否显示数值标签：系列级 dLbls 优先，其次图表类型级 dLbls；
    任一层把标签整体删除（delete=1）或存在被单独删除的标签，均视为不达标。"""
    for scope in (ser, plot):
        node = scope.find("c:dLbls", NS)
        if node is None:
            continue
        if _bool_attr(node.find("c:delete", NS)):
            return False
        for dlbl in node.findall("c:dLbl", NS):
            if _bool_attr(dlbl.find("c:delete", NS)):
                return False
        if _bool_attr(node.find("c:showVal", NS)) or _bool_attr(node.find("c:showPercent", NS)):
            return True
    return False


def _has_outline(ser, count):
    """扇区是否有描边：系列级 spPr/a:ln 有效，或每个数据点各自设了 a:ln。"""
    def _visible_line(spPr):
        if spPr is None:
            return False
        ln = spPr.find("a:ln", NS)
        if ln is None:
            return False
        return ln.find("a:noFill", NS) is None
    if _visible_line(ser.find("c:spPr", NS)):
        return True
    dpts = ser.findall("c:dPt", NS)
    if not dpts or len(dpts) < count:
        return False
    return all(_visible_line(dpt.find("c:spPr", NS)) for dpt in dpts)


def _col_width_px(sheet, col):
    """列宽（字符）-> 像素，按 Calibri 11 的 MDW=7 折算。"""
    width = None
    for entry in sheet["cols"]:
        if entry["min"] <= col <= entry["max"] and entry["width"] is not None:
            width = entry["width"]
            break
    if width is None:
        width = sheet["default_width"]
    return round(width * 7.0) + 5.0


def _row_height_px(sheet, row):
    """行高（磅）-> 像素（96 DPI）。"""
    info = sheet["rows"].get(row)
    pts = info["ht"] if info and info["ht"] is not None else sheet["default_height"]
    return pts * 96.0 / 72.0


def _chart_geometry(ctx, chart):
    """把锚点换算成 (宽 EMU, 高 EMU, 覆盖的 1 基单元格区域 c1,r1,c2,r2)。
    twoCellAnchor 按起止格累加列宽/行高；oneCellAnchor 按 ext 反推终止格。"""
    sheet = ctx["sheets"].get(chart["sheet"])
    origin = chart["from"]
    if sheet is None or origin is None:
        return (0, 0, None)
    col0, coloff, row0, rowoff = origin
    if chart["anchor_tag"] == "twoCellAnchor" and chart["to"]:
        col1, coloff1, row1, rowoff1 = chart["to"]
        cx = sum(_col_width_px(sheet, c + 1) for c in range(col0, col1)) * EMU_PER_PX
        cy = sum(_row_height_px(sheet, r + 1) for r in range(row0, row1)) * EMU_PER_PX
        cx = cx - coloff + coloff1
        cy = cy - rowoff + rowoff1
        return (max(cx, 0), max(cy, 0), (col0 + 1, row0 + 1, col1 + 1, row1 + 1))
    cx, cy = chart["ext"] or (0, 0)
    col1, acc = col0, -coloff
    while acc < cx and col1 - col0 < 256:
        acc += _col_width_px(sheet, col1 + 1) * EMU_PER_PX
        col1 += 1
    row1, acc = row0, -rowoff
    while acc < cy and row1 - row0 < 4096:
        acc += _row_height_px(sheet, row1 + 1) * EMU_PER_PX
        row1 += 1
    return (cx, cy, (col0 + 1, row0 + 1, max(col1, col0 + 1), max(row1, row0 + 1)))


def _rects_overlap(rect_a, rect_b):
    """两个 1 基单元格矩形 (c1,r1,c2,r2) 是否相交。"""
    if not rect_a or not rect_b:
        return False
    return not (rect_a[2] < rect_b[0] or rect_b[2] < rect_a[0] or
                rect_a[3] < rect_b[1] or rect_b[3] < rect_a[1])


def _new_cells_outside_data(sheet):
    """数据表以外新写入的非空单元格；返回 {(row, col): cell}。"""
    out = {}
    for (row, col), cell in sheet["cells"].items():
        if row <= DATA_ROWS and col <= DATA_COLS:
            continue
        if cell["text"] not in (None, "") or cell["number"] is not None:
            out[(row, col)] = cell
    return out


def _bbox(keys):
    """一组 (row, col) 的包围盒 (c1, r1, c2, r2)；空集合返回 None。"""
    if not keys:
        return None
    rows = [k[0] for k in keys]
    cols = [k[1] for k in keys]
    return (min(cols), min(rows), max(cols), max(rows))


# ==================================================================== 入口


def _zero_items():
    """维度1 判负或脚本报错时，仍按细则顺序给出完整的 34 条评分项：
    delta 全 0、hit 全 False，避免调用方拿到空列表。"""
    return [{"rule": rule, "max_delta": max_delta, "delta": 0, "hit": False}
            for rule, max_delta in zip(RULES, _MAX_DELTAS_ORDER)]


def evaluate(dir_path):
    """按接口约定返回结构化 dict；不 print、不改 sys.stdout、不 sys.exit。
    入参 dir_path 为"脚本所在目录的路径"，脚本自己在该目录里定位并打开被评估的
    .xlsx / .xlsm（忽略以 ~$ 开头的 Excel 临时锁文件）。目录不存在或目录内没有
    工作簿时，返回"未找到文件"的空结果（各评分字段留空/归零）；其余脚本自身崩溃
    由通用兜底捕获并返回 status='error'。"""
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
            raise FileNotFoundError(f"目录不存在或不是目录：{dir_path}")
        cands = sorted(
            f for f in os.listdir(dir_path)
            if f.lower().endswith((".xlsx", ".xlsm")) and not f.startswith("~$")
            and os.path.isfile(os.path.join(dir_path, f))
        )
        if not cands:
            raise FileNotFoundError(f"目录中未找到 .xlsx 文件：{dir_path}")
        file_path = os.path.join(dir_path, cands[0])
        result["file_name"] = os.path.basename(file_path)
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
        result["error"] = f"脚本运行异常：{type(exc).__name__}: {exc}"
        result["dim1_pass"] = False
        result["dim1_reason"] = (f"维度一未通过：文件无法打开或解析"
                                f"（{type(exc).__name__}: {exc}），完成度总分记 0")
        result["dim2_items"] = _zero_items()
        result["total_score"] = 0
        return result


def _check_dimension1(ctx):
    """维度1：可用与可修改性。返回不通过的原因列表（空列表 = 通过）。"""
    fails = []
    # 1) 包结构完整、所有 XML 部件可解析（打开时不会触发"发现部分内容有问题"）
    for part in sorted(ctx["names"]):
        if not part.endswith((".xml", ".rels")):
            continue
        try:
            ET.fromstring(ctx["zip"].read(part))
        except (ET.ParseError, KeyError) as exc:
            fails.append("部件 %s 无法解析：%s" % (part, exc))
    if "[Content_Types].xml" not in ctx["names"]:
        fails.append("缺少 [Content_Types].xml，工作簿无法作为 OPC 包打开")
    # 2) 两张成绩表仍在，且 A1:Q7 未错行错列、文字数字未大面积丢失
    for name in SHEET_NAMES:
        sheet = ctx["sheets"].get(name)
        if sheet is None:
            fails.append("工作表「%s」缺失，成绩数据无法定位" % name)
            continue
        bad_labels = [_cell_addr(r, 1) for r in range(1, DATA_ROWS + 1)
                      if _norm_text((_cell(ctx, name, r, 1) or {}).get("text"))
                      != A_LABELS[r - 1]]
        if bad_labels:
            fails.append("工作表「%s」A 列标签错位/丢失：%s" % (name, ",".join(bad_labels)))
        kept = sum(1 for row in range(3, DATA_ROWS + 1) for col in range(2, DATA_COLS + 1)
                   if (_cell(ctx, name, row, col) or {}).get("number") is not None)
        if kept < 72:
            fails.append("工作表「%s」A3:Q7 的 80 个数值仅剩 %d 个，数据大面积丢失"
                         % (name, kept))
    return fails + _check_dimension1_charts(ctx)


def _check_dimension1_charts(ctx):
    """维度1 中与图表对象有关的门槛：必须是可点选可改数据源的图表对象，
    且图表区能渲染出扇区（不是空白框、也不是粘贴的静态位图）。"""
    fails = []
    if not ctx["charts"]:
        if ctx["pictures"]:
            fails.append("工作表内只有 %d 个图片对象、没有任何 Excel 图表对象，"
                         "图表为图片形式" % ctx["pictures"])
        else:
            fails.append("工作簿内未发现任何图表对象")
        return fails
    for chart in ctx["charts"]:
        series = _series_of(chart)
        if not series:
            fails.append("%s 内没有饼/环形系列，图表区空白无法渲染" % chart["part"])
            continue
        for item in series:
            values = [v for v in _series_values(ctx, chart, item["ser"]) if v is not None]
            if not values or sum(values) <= 0:
                fails.append("%s 的某个系列取不到有效数值，扇区无法渲染" % chart["part"])
                break
    return fails


def _evaluate_impl(file_path, result):
    """维度一 + 维度二的评估主体；内部构造 results/details 两个平行列表，
    末尾按 _MAX_DELTAS_ORDER 反解成结构化 dim2_items 填入 result。"""
    ctx = _load_package(file_path)
    results = []
    details = []

    def emit(hit, rule, detail=""):
        """追加一条评分结论。命中记 max_delta，未命中记 0；顺序须与
        _MAX_DELTAS_ORDER 一致，这里顺带校验，写错立即暴露。
        detail 只作为内部取证说明留档，不随结果返回。"""
        max_delta = _MAX_DELTAS_ORDER[len(results)]
        results.append("%+d：%s" % (max_delta if hit else 0, rule))
        details.append(str(detail).strip() or ("判定命中" if hit else "判定未命中"))

    # ========== 维度1：可用与可修改性 ==========
    fails = _check_dimension1(ctx)
    if fails:
        result["dim1_pass"] = False
        # dim1_reason 统一带上「维度一未通过：…」前缀与计分后果，便于直接作为错误摘要展示
        result["dim1_reason"] = "维度一未通过：%s，完成度总分记 0" % "；".join(fails[:6])
        result["dim2_items"] = _zero_items()
        result["total_score"] = 0
        return result
    result["dim1_pass"] = True
    result["dim1_reason"] = "工作簿可正常打开，两张成绩表完整，图表为可编辑的 Excel 图表对象"

    # ========== 维度2：完成度评分 ==========
    # --- 预处理：把每张图表整理成便于逐条判定的画像 ---
    profiles = [_chart_profile(ctx, chart) for chart in ctx["charts"]]
    by_sheet = {}
    for prof in profiles:
        by_sheet.setdefault(prof["sheet"], []).append(prof)

    _score_charts(ctx, profiles, by_sheet, emit)
    _score_summary(ctx, profiles, emit)
    _score_readability(ctx, profiles, emit)
    _score_deductions(ctx, profiles, emit)

    # ========== 将 results 反解为结构化 dim2_items ==========
    # 逐项按追加顺序对齐 _MAX_DELTAS_ORDER：解析每条 "±N：说明" 的 N 与 rule。
    if len(results) != len(_MAX_DELTAS_ORDER):
        raise RuntimeError(
            f"dim2 项数不一致：results={len(results)} vs order={len(_MAX_DELTAS_ORDER)}"
        )
    dim2_items = []
    total_score = 0
    # details 仅保留各判定处的取证说明供排查，不再随结果返回
    for raw, _detail, max_delta in zip(results, details, _MAX_DELTAS_ORDER):
        head, _, rule = raw.partition("：")
        try:
            delta = int(head)
        except ValueError:
            delta = 0
        hit = delta != 0
        dim2_items.append({
            "rule": rule,
            "max_delta": max_delta,
            "delta": max_delta if hit else 0,
            "hit": hit,
        })
        total_score += (max_delta if hit else 0)

    result["dim2_items"] = dim2_items
    result["total_score"] = total_score
    result["max_score"] = MAX_SCORE_TOTAL
    return result


def _chart_profile(ctx, chart):
    """把一张图表整理成逐条细则都能直接读的画像。"""
    prof = {"sheet": chart["sheet"], "part": chart["part"], "chart": chart,
            "title": _chart_title(ctx, chart), "series": [], "data_sheets": set()}
    for item in _series_of(chart):
        ser = item["ser"]
        values = _series_values(ctx, chart, ser)
        labels = _series_labels(ctx, chart, ser)
        val_ref = _ref_of(ser, "val")
        cells = list(_iter_ref_cells(val_ref, chart["sheet"])) if val_ref else []
        prof["data_sheets"].update(s for s, _, _ in cells)
        prof["series"].append({
            "ser": ser, "plot": item["plot"], "plot_tag": item["plot_tag"],
            "values": values, "labels": labels, "count": len(values),
            "val_ref": val_ref, "cat_ref": _ref_of(ser, "cat"), "val_cells": cells,
            "has_ref": ser.find("c:val//c:numRef/c:f", NS) is not None,
            "colors": _dpt_colors(ser, chart["theme"], len(values)),
            "shows_val": _shows_values(item["plot"], ser),
            "outline": _has_outline(ser, len(values)),
        })
    prof["ring_count"] = len(prof["series"])
    holes = []
    for plot in chart["root"].findall(".//c:doughnutChart", NS):
        node = plot.find("c:holeSize", NS)
        holes.append(float(node.get("val")) if node is not None and node.get("val") else 0.0)
    prof["hole"] = min(holes) if holes else 0.0
    prof["is_pie_family"] = bool(prof["series"])
    prof["legend"] = chart["root"].find(".//c:legend/c:legendPos", NS) is not None
    # 内层 / 外层
    inner = prof["series"][0] if prof["series"] else None
    outer = prof["series"][1] if len(prof["series"]) > 1 else None
    prof["inner"], prof["outer"] = inner, outer
    prof["inner_kind"] = None
    if inner is not None:
        if tuple(inner["labels"]) == CATEGORIES:
            prof["inner_kind"] = "category"
        elif tuple(inner["labels"]) == LEVELS:
            prof["inner_kind"] = "level"
    expect = None
    if prof["inner_kind"] == "category":
        expect = INNER_BY_CATEGORY.get(prof["sheet"])
    elif prof["inner_kind"] == "level":
        expect = INNER_BY_LEVEL.get(prof["sheet"])
    prof["inner_expect"] = expect
    prof["inner_values_ok"] = bool(
        expect and inner and len(inner["values"]) == len(expect) and
        all(v is not None and abs(v - e) <= 0.5 for v, e in zip(inner["values"], expect)))
    prof["chunks"] = (_chunk_align(inner["values"], outer["values"])
                      if inner and outer else None)
    prof["cx"], prof["cy"], prof["rect"] = _chart_geometry(ctx, chart)
    return prof


# ============================================== 细则原文（与 _MAX_DELTAS_ORDER 同序）
RULES = (
    # ---- 一、加分项 ----
    '工作表"七1班"与"七13班"各内嵌 1 张图表，共 2 张，两个班的数据分别独立成图'
    '（未把两班混在同一张图内、未只做 1 张）',
    '两张图表各有可区分班级的图表标题（如"七1班 探究成绩多类别嵌套饼图"'
    '"七13班 探究成绩多类别嵌套饼图"）',
    '图表为饼/环形族的双层同心嵌套结构（同一张图内 2 个数据系列同心叠放），'
    '不是并列摆放的两个独立饼图，也不是单层饼图',
    '圈心孔径为 0 或接近 0，内层呈实心饼、外层呈环带，整体为"合并饼图"外观'
    '而非中间大空洞的双环图',
    '内层扇区依次为"地图基础""经纬定位""地形判读""气候要素""区域差异""聚落形态"'
    '六个探究表现大类，或"卓越表现""稳定达成""基本达成""持续提升"四个能力层级大类',
    '外层扇区为能力层级或探究表现细分（按 24 个能力层级或探究表现明细项），'
    '且外层按大类顺序连续排列，每个内层扇区正对其下属外层扇区，内外层角度对齐',
    '图表只有探究表现、能力层级两个层级，未额外堆叠第三层及以上圈层',
    '内层各扇区数值等于该大类下各明细项之和——以探究表现作内层时，七1班依次 '
    '153/153/106/99/157/150，七13班依次 156/148/102/104/150/153；以能力层级作内层时，'
    '七1班依次 463/175/114/66，七13班依次 493/163/101/56（两班扇区合计分别为 818、813）；'
    '且外层同一大类内各扇区之和等于对应内层扇区数值',
    '图表数据源指向工作簿内的单元格区域引用，而非把数值硬编码为图表内字面量，'
    '源数据改动后图表可更新',
    '为绘图新增的汇总区域放在不覆盖数据表的空白区域，未占用 A1:Q7 任何单元格',
    '汇总数值由公式（如 SUM）计算得出而非手敲常量，改动 A1:Q7 中人数后汇总与图表可联动',
    '新增汇总区域带可读的行列表头（大类名称、能力层级或等级名称）',
    '同一大类的内层扇区与其外层子扇区使用同一色系/同一主题色，可直观识别归属关系',
    '内外两层扇区均显示数据标签（人数或百分比），标签数量与扇区数一致',
    '扇区之间有描边（如白色细边线）分隔，相邻扇区不粘连',
    '图表带图例，图例项为探究表现或能力层级的类别名称',
    '图表尺寸不小于约 15cm×10cm（约横跨 12 列、纵跨 20 行），扇区标签互不重叠、不被截断',
    '图表放在其数据所在的同一工作表内，且未压盖 A1:Q7 数据表区域与新增汇总区域',
    # ---- 二、减分项 ----
    '两张工作表第 1 行 B1:D1、E1:G1、H1:I1、J1:K1、L1:N1、O1:Q1 六个合并区域内的'
    '全部单元格（含被合并覆盖的 C1、D1、F1、G1、I1、K1、M1、N1、P1、Q1）格式被重置'
    '——不再满足纯色填充 #174A55、字体 黑体 11 磅加粗白色 #FFFFFF、水平居中 + 垂直居中，'
    '以及"区域外侧竖边为中粗线、区域内部竖边为细线、上下边为细线"的分隔关系'
    '（被合并覆盖的格被改成等线 12 磅非加粗且无填充、取消合并后暴露只有首格有底色的断色，'
    'L1 右侧竖边被改为中粗线，均计）',
    '第 2 行"识图/判读/表达/观察/分析/归纳/定位/解释/识别/比较/计算/推理/应用/观察/判断/表达"'
    '16 个单元格的字体、填充、边框、居中格式被改动',
    'A 列"探究表现""能力层级""参与人数""卓越表现""稳定达成""基本达成""持续提升"'
    '7 个标签单元格的字体、填充、边框格式被改动',
    '数据表第 1、2 行行高不满足 15，第 3~7 行行高不满足 25（被统一压成同一数值均计）',
    '为数据区以外的大量空白行（如第 8~120 行）逐行设置了自定义行高',
    'A 列列宽不满足 24.17、B~Q 列列宽不满足 8',
    '在 S~AB 等空白列上新增自定义列宽（如 24/10/3），改变数据表以外区域的列宽外观',
    'A1:Q7 内任一文字或数值被增删改写（6 个探究表现大类名、16 个能力层级名、'
    '参与人数与四个等级共 80 个整数须逐一保持原值原类型）',
    'B1:D1、E1:G1、H1:I1、J1:K1、L1:N1、O1:Q1 六个合并单元格被取消合并或合并范围改变',
    '删除"七1班"或"七13班"任一工作表，或改变两张工作表的名称与先后顺序',
    '删除或破坏工作簿内已有的其它图表、图片、批注等对象',
    '工作表"不显示网格线"的显示设置被改为显示网格线',
    '在工作表单元格内新增与图表标题重复的文字标题及其合并区域'
    '（如在 S1:AB1 写入"七1班 探究成绩多类别嵌套饼图"）',
    '新增的中间计算区域铺开超过 30 行（如展开成 60 行以上的逐项明细清单），'
    '使工作表出现大段冗余内容',
    '图表压盖 A1:Q7 数据表或压盖新增汇总区域，遮挡单元格内容',
    '新增与绘图无关的工作表、隐藏行列或定义名称',
)
assert len(RULES) == len(_MAX_DELTAS_ORDER)


# =============================================================== 维度2 计分


def _summary_cells(prof):
    """该图表在自己工作表上、数据表以外引用到的单元格（即为绘图新增的汇总区域）。"""
    keys = set()
    for item in prof["series"]:
        for ref in (item["val_ref"], item["cat_ref"]):
            for sheet, row, col in _iter_ref_cells(ref, prof["sheet"]):
                if sheet == prof["sheet"] and not (row <= DATA_ROWS and col <= DATA_COLS):
                    keys.add((row, col))
    return keys


def _score_charts(ctx, profiles, by_sheet, emit):
    """加分项 01~08：交付覆盖与嵌套结构。"""
    # 01 两表各内嵌 1 张、两班数据分别独立成图
    spread = [len(by_sheet.get(name, [])) for name in SHEET_NAMES]
    cross = [p["sheet"] for p in profiles if p["data_sheets"] - {p["sheet"]}]
    emit(len(profiles) == 2 and spread == [1, 1] and not cross, RULES[0],
         "图表共 %d 张，按表分布 %s，跨班取数的图表：%s"
         % (len(profiles), dict(zip(SHEET_NAMES, spread)), cross or "无"))
    # 02 标题可区分班级
    titles = {p["sheet"]: p["title"] for p in profiles}
    emit(len(profiles) == 2 and len(set(titles.values())) == 2
         and all(name in (titles.get(name) or "") for name in SHEET_NAMES),
         RULES[1], "图表标题：%s" % (titles,))
    # 03 双层同心嵌套：同一张图内恰为 2 个饼/环形系列同心叠放
    rings = [p["ring_count"] for p in profiles]
    emit(bool(profiles) and all(n == 2 for n in rings), RULES[2],
         "各图饼/环形系列数：%s（应恰为 2 层；1 层是单层饼图，3 层及以上已不是双层结构）"
         % (rings,))
    # 04 圈心孔径为 0 或接近 0
    holes = [p["hole"] for p in profiles]
    emit(bool(profiles) and all(h <= 10.0 for h in holes), RULES[3],
         "各图 holeSize：%s（<=10 视为接近实心）" % (holes,))
    # 05 内层为六个探究表现大类或四个能力层级大类
    kinds = {p["sheet"]: (p["inner_kind"] or "未识别（内层分类名=%s）"
                          % ((p["inner"]["labels"][:6] if p["inner"] else []),))
             for p in profiles}
    emit(bool(profiles) and all(p["inner_kind"] for p in profiles), RULES[4],
         "内层分类：%s" % (kinds,))
    # 06 外层 24 个明细扇区、按大类顺序连续排列、内外角度对齐
    outer_n = [(p["outer"]["count"] if p["outer"] else 0) for p in profiles]
    aligned = [p["chunks"] is not None for p in profiles]
    emit(bool(profiles) and all(n == 24 for n in outer_n) and all(aligned), RULES[5],
         "外层扇区数：%s；内外层分组对齐：%s" % (outer_n, aligned))
    # 07 只有两个层级
    emit(bool(profiles) and all(n == 2 for n in rings), RULES[6],
         "各图圈层数：%s" % (rings,))
    # 08 内层数值等于大类之和，外层组内之和等于内层
    vals_ok = [(p["inner_values_ok"], p["chunks"] is not None) for p in profiles]
    emit(bool(profiles) and all(a and b for a, b in vals_ok), RULES[7],
         "（内层数值匹配, 外层分组求和匹配）=%s" % (vals_ok,))


_LABEL_VOCAB = tuple(set(CATEGORIES) | set(LEVELS) | set(ABILITY_NAMES) | {"类别", "人数"})


def _score_summary(ctx, profiles, emit):
    """加分项 09~12：汇总数据与联动。"""
    # 09 数据源为单元格区域引用
    refs = [item["has_ref"] for p in profiles for item in p["series"]]
    emit(bool(refs) and all(refs), RULES[8],
         "%d/%d 个系列的数值绑定了 c:numRef 区域引用" % (sum(1 for x in refs if x), len(refs)))
    # 10 汇总区域未占用 A1:Q7
    invade = []
    for prof in profiles:
        for item in prof["series"]:
            for ref in (item["val_ref"], item["cat_ref"]):
                if ref and _ref_hits_data_area(ref, prof["sheet"]):
                    invade.append(ref)
    emit(bool(profiles) and not invade, RULES[9],
         "引用落入 A1:Q7 的区域：%s" % (invade[:3] or "无"))
    # 11 汇总数值由公式计算
    total = hardcoded = 0
    for prof in profiles:
        for item in prof["series"]:
            for sheet, row, col in _iter_ref_cells(item["val_ref"], prof["sheet"]):
                cell = _cell(ctx, sheet, row, col)
                total += 1
                if cell is None or not cell["has_formula"]:
                    hardcoded += 1
    emit(bool(total) and hardcoded == 0, RULES[10],
         "被引用的 %d 个数值单元格中，%d 个是手敲常量（无公式）" % (total, hardcoded))
    # 12 汇总区域带可读行列表头
    counts = {}
    for prof in profiles:
        sheet = ctx["sheets"].get(prof["sheet"])
        if sheet is None:
            continue
        hit = 0
        for (row, col), cell in _new_cells_outside_data(sheet).items():
            text = _norm_text(cell["text"])
            if not text or cell["number"] is not None:
                continue
            if any(word in text for word in _LABEL_VOCAB):
                hit += 1
        counts[prof["sheet"]] = hit
    emit(bool(counts) and all(n >= 4 for n in counts.values()), RULES[11],
         "各表汇总区可读表头/分类名数量：%s" % (counts,))


def _color_family_ok(prof):
    """内层每个扇区与其下属外层子扇区是否同色系；返回 (是否达标, 说明)。"""
    if prof["chunks"] is None or not prof["inner"] or not prof["outer"]:
        return (False, "内外层无法按大类对齐，无法判定归属配色")
    inner_c, outer_c = prof["inner"]["colors"], prof["outer"]["colors"]
    bad = []
    for i, (start, end) in enumerate(prof["chunks"]):
        base = inner_c[i] if i < len(inner_c) else None
        if not base:
            bad.append("内层第%d扇区未单独设色" % (i + 1))
            continue
        for j in range(start, end):
            sub = outer_c[j] if j < len(outer_c) else None
            if not sub:
                bad.append("外层第%d扇区未单独设色" % (j + 1))
            elif not _same_color_family(base, sub):
                bad.append("外层第%d扇区 #%s 与内层第%d扇区 #%s 非同色系"
                           % (j + 1, sub, i + 1, base))
    return (not bad, "；".join(bad[:3]) if bad else "内外层逐大类同色系")


def _score_readability(ctx, profiles, emit):
    """加分项 13~18：可读性与摆放。"""
    # 13 内外层同色系
    color = [_color_family_ok(p) for p in profiles]
    emit(bool(color) and all(ok for ok, _ in color), RULES[12],
         "；".join("%s：%s" % (p["sheet"], why) for p, (_, why) in zip(profiles, color)))
    # 14 内外两层均显示数据标签
    shown = [(p["sheet"], [item["shows_val"] for item in p["series"]]) for p in profiles]
    emit(bool(profiles) and all(flags and all(flags) for _, flags in shown), RULES[13],
         "各图逐系列标签显示：%s" % (dict(shown),))
    # 15 扇区之间有描边
    lines = [(p["sheet"], [item["outline"] for item in p["series"]]) for p in profiles]
    emit(bool(profiles) and all(flags and all(flags) for _, flags in lines), RULES[14],
         "各图逐系列描边：%s" % (dict(lines),))
    # 16 图表带图例
    emit(bool(profiles) and all(p["legend"] for p in profiles), RULES[15],
         "各图图例：%s" % ({p["sheet"]: p["legend"] for p in profiles},))
    # 17 尺寸不小于 15cm×10cm 且标签不重叠不截断
    geo = {}
    for prof in profiles:
        widest = max([item["count"] for item in prof["series"]] or [0])
        geo[prof["sheet"]] = ("%.1f×%.1fcm" % (prof["cx"] / EMU_PER_CM, prof["cy"] / EMU_PER_CM),
                              "单圈最多%d个扇区" % widest)
    size_ok = bool(profiles) and all(
        p["cx"] >= MIN_CHART_CX and p["cy"] >= MIN_CHART_CY
        and max([item["count"] for item in p["series"]] or [0]) <= MAX_READABLE_POINTS
        for p in profiles)
    emit(size_ok, RULES[16], "各图尺寸与扇区密度：%s" % (geo,))
    # 18 与数据同表且未压盖数据表与汇总区
    place = []
    for prof in profiles:
        summary = _bbox(_summary_cells(prof))
        if prof["data_sheets"] - {prof["sheet"]}:
            place.append("%s：图表与其数据源不在同一工作表" % prof["sheet"])
        elif _rects_overlap(prof["rect"], (1, 1, DATA_COLS, DATA_ROWS)):
            place.append("%s：图表压盖 A1:Q7" % prof["sheet"])
        elif _rects_overlap(prof["rect"], summary):
            place.append("%s：图表压盖汇总区域 %s" % (prof["sheet"], summary))
    emit(bool(profiles) and not place, RULES[17],
         "；".join(place) if place else "两图均与数据同表且未压盖数据表与汇总区")


def _score_deductions(ctx, profiles, emit):
    """减分项 19~39：指令未要求、却破坏成绩表本体与工作簿结构的多余修改。"""
    def fmt(sheet, row, col):
        return _style_of(ctx, _cell(ctx, sheet, row, col))

    sheets = [n for n in SHEET_NAMES if n in ctx["sheets"]]
    # 19 第 1 行六个合并区域全部单元格（含被覆盖格）的填充/字体/对齐/边框
    bad = []
    for name in sheets:
        for col in range(2, DATA_COLS + 1):
            got = fmt(name, 1, col)
            want_l, want_r = ROW1_VERT[col]
            why = []
            if got["fill"] != ("solid", HEAD_FILL):
                why.append("填充=%s" % (got["fill"],))
            if got["font"] != HEAD_FONT:
                why.append("字体=%s" % (got["font"],))
            if got["align"] != ("center", "center"):
                why.append("对齐=%s" % (got["align"],))
            if (got["border"].get("left") != want_l or got["border"].get("right") != want_r
                    or got["border"].get("top") != "thin"
                    or got["border"].get("bottom") != "thin"):
                why.append("边框=%s" % (got["border"],))
            if why:
                bad.append("%s!%s %s" % (name, _cell_addr(1, col), "、".join(why)))
    emit(bool(bad), RULES[18],
         "格式被重置的表头格（共%d个）：%s" % (len(bad), "；".join(bad[:2]) or "无"))
    # 20 第 2 行 16 个能力层级单元格格式
    bad = []
    for name in sheets:
        for col in range(2, DATA_COLS + 1):
            got = fmt(name, 2, col)
            want_l, want_r = ROW2_VERT[col]
            if (got["font"] != ROW2_FONT or got["fill"] != ("solid", ROW2_FILL)
                    or got["align"] != ("center", "center")
                    or got["border"].get("left") != want_l
                    or got["border"].get("right") != want_r
                    or got["border"].get("top") != "thin"
                    or got["border"].get("bottom") != "thin"):
                bad.append("%s!%s" % (name, _cell_addr(2, col)))
    emit(bool(bad), RULES[19],
         "格式被改的能力层级格（共%d个）：%s" % (len(bad), bad[:8] or "无"))
    # 21 A 列 7 个标签单元格格式
    bad = []
    for name in sheets:
        for row, want in ACOL_EXPECT.items():
            got = fmt(name, row, 1)
            font, fill, horiz, vert, top, bottom = want
            if (got["font"] != font or got["fill"] != ("solid", fill)
                    or got["align"] != (horiz, vert)
                    or got["border"].get("top") != top
                    or got["border"].get("bottom") != bottom
                    or got["border"].get("left") != "medium"
                    or got["border"].get("right") != "medium"):
                bad.append("%s!A%d" % (name, row))
    emit(bool(bad), RULES[20],
         "格式被改的 A 列标签格（共%d个）：%s" % (len(bad), bad[:8] or "无"))
    # 22 第 1、2 行行高 15，第 3~7 行行高 25
    bad = []
    for name in sheets:
        sheet = ctx["sheets"][name]
        for row, want in ROW_HEIGHTS.items():
            got = sheet["rows"].get(row, {}).get("ht")
            got = sheet["default_height"] if got is None else got
            if abs(got - want) > HEIGHT_TOL:
                bad.append("%s!第%d行=%.4g（应%.4g）" % (name, row, got, want))
    emit(bool(bad), RULES[21], "；".join(bad[:4]) if bad else "两表 1~7 行行高均为原值")
    # 23 数据区以外空白行逐行自定义行高
    counts = {}
    for name in sheets:
        sheet = ctx["sheets"][name]
        counts[name] = sum(1 for row, info in sheet["rows"].items()
                           if row > DATA_ROWS and info["custom"] and info["ht"] is not None)
    emit(any(n >= 20 for n in counts.values()), RULES[22],
         "各表数据区以外被逐行设置自定义行高的行数：%s" % (counts,))

    def width_of(sheet, col):
        for entry in sheet["cols"]:
            if entry["min"] <= col <= entry["max"] and entry["width"] is not None:
                return entry["width"]
        return sheet["default_width"]

    # 24 A 列列宽 24.17、B~Q 列列宽 8
    bad = []
    for name in sheets:
        sheet = ctx["sheets"][name]
        for col in range(1, DATA_COLS + 1):
            want = COL_WIDTH_A if col == 1 else COL_WIDTH_BQ
            got = width_of(sheet, col)
            if abs(got - want) > WIDTH_TOL:
                bad.append("%s!%s列=%.4g（应%.4g）" % (name, _col_name(col), got, want))
    emit(bool(bad), RULES[23], "；".join(bad[:4]) if bad else "两表 A~Q 列列宽均为原值")
    # 25 S~AB 等空白列新增自定义列宽
    extra = {}
    for name in sheets:
        got = ["%s:%s=%.4g" % (_col_name(e["min"]), _col_name(e["max"]), e["width"])
               for e in ctx["sheets"][name]["cols"]
               if e["min"] > DATA_COLS and e["width"] is not None]
        if got:
            extra[name] = got
    emit(bool(extra), RULES[24],
         "数据表以外新增的自定义列宽：%s" % (extra or "无"))
    # 26 A1:Q7 内文字或数值被增删改写
    bad = []
    starts = {span[0]: name for span, name in zip(CAT_SPANS, CATEGORIES)}
    for name in sheets:
        for row in range(1, DATA_ROWS + 1):
            for col in range(1, DATA_COLS + 1):
                cell = _cell(ctx, name, row, col)
                text = (cell["text"] or "") if cell else ""
                if col == 1:
                    want = A_LABELS[row - 1]
                elif row == 1:
                    want = starts.get(col, "")
                elif row == 2:
                    want = ABILITY_NAMES[col - 2]
                else:
                    number = cell["number"] if cell else None
                    want_n = DATA_BASELINE[name][row][col - 2]
                    if number is None or abs(number - want_n) > 1e-9:
                        bad.append("%s!%s=%r（应 %d）"
                                   % (name, _cell_addr(row, col), text, want_n))
                    continue
                if _norm_text(text) != _norm_text(want):
                    bad.append("%s!%s=%r（应 %r）"
                               % (name, _cell_addr(row, col), text, want))
    emit(bool(bad), RULES[25],
         "被改写的单元格（共%d个）：%s" % (len(bad), bad[:4] or "无"))
    # 27 六个合并单元格被取消或范围改变
    bad = []
    for name in sheets:
        have = set(ctx["sheets"][name]["merges"])
        missing = [ref for ref in MERGE_REFS if ref not in have]
        if missing:
            bad.append("%s 缺少 %s" % (name, missing))
    emit(bool(bad), RULES[26], "；".join(bad) if bad else "两表六个合并表头区域完好")
    # 28 删除工作表或改变名称与顺序
    order = tuple(ctx["sheet_order"][:len(SHEET_NAMES)])
    emit(order != SHEET_NAMES, RULES[27],
         "工作表顺序=%s（应为 %s）" % (list(ctx["sheet_order"]), list(SHEET_NAMES)))
    # 29 删除或破坏已有的其它图表/图片/批注等对象
    emit(False, RULES[28],
         "成绩表内不含图表、图片、批注等已有对象，无可破坏项，本项不计扣分")
    # 30 "不显示网格线"被改为显示网格线
    shown = [name for name in sheets if ctx["sheets"][name]["gridlines"]]
    emit(bool(shown), RULES[29],
         "网格线被改为显示的工作表：%s" % (shown or "无"))
    # 31 单元格内新增与图表标题重复的文字标题
    titles = [p["title"] for p in profiles if len(p["title"] or "") >= 6]
    dup = []
    for name in sheets:
        for (row, col), cell in _new_cells_outside_data(ctx["sheets"][name]).items():
            text = _norm_text(cell["text"])
            if len(text) >= 6 and any(t in text or text in t for t in titles):
                dup.append("%s!%s=%r" % (name, _cell_addr(row, col), cell["text"]))
    emit(bool(dup), RULES[30],
         "与图表标题重复的单元格文字：%s" % (dup[:4] or "无"))
    # 32 中间计算区域铺开超过 30 行
    spans = {}
    for name in sheets:
        rows = set(r for r, _ in _new_cells_outside_data(ctx["sheets"][name]))
        spans[name] = len(rows)
    emit(any(n > 30 for n in spans.values()), RULES[31],
         "各表数据区以外写入内容的行数：%s" % (spans,))
    # 33 图表压盖 A1:Q7 或压盖汇总区域
    cover = []
    for prof in profiles:
        summary = _bbox(_summary_cells(prof))
        if _rects_overlap(prof["rect"], (1, 1, DATA_COLS, DATA_ROWS)):
            cover.append("%s：压盖 A1:Q7" % prof["sheet"])
        elif _rects_overlap(prof["rect"], summary):
            cover.append("%s：压盖汇总区域 %s" % (prof["sheet"], summary))
    emit(bool(cover), RULES[32], "；".join(cover) if cover else "两图均未压盖表格与汇总区")
    # 34 新增无关工作表、隐藏行列或定义名称
    junk = []
    spare = [n for n in ctx["sheet_order"] if n not in SHEET_NAMES]
    if spare:
        junk.append("新增工作表 %s" % spare)
    hidden = [n for n in sheets if ctx["sheets"][n]["hidden_rc"]]
    if hidden:
        junk.append("存在隐藏行或隐藏列的工作表 %s" % hidden)
    if ctx["defined_names"]:
        junk.append("新增定义名称 %s" % ctx["defined_names"][:5])
    emit(bool(junk), RULES[33], "；".join(junk) if junk else "无多余工作表、隐藏行列与定义名称")


if __name__ == "__main__":
    # 仅本地调试用：正式评测由外部调用 evaluate(dir_path)，本模块自身不打印结果。
    if len(sys.argv) >= 2:
        _dir = sys.argv[1]
    else:
        _dir = os.path.dirname(os.path.abspath(__file__))
    _out = evaluate(_dir)
    print(json.dumps(_out, ensure_ascii=False, indent=2))
