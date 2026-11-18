#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# pyright: reportUnknownMemberType=false, reportUnknownVariableType=false, reportUnknownArgumentType=false, reportUnknownParameterType=false, reportMissingParameterType=false, reportPrivateUsage=false, reportUnknownLambdaType=false, reportGeneralTypeIssues=false

import os
import re
import sys
import json
import zipfile
import traceback
from xml.etree import ElementTree as ET
import openpyxl

# 说明：按接口统一约定，本模块不修改全局 sys.stdout，也不 print 主结果。
# 主结果统一由 evaluate(dir_path) -> dict 返回，入参为"脚本所在目录的路径"，
# 由脚本自己在该目录内定位并打开被评估的工作簿。本文件仅在 __main__ 下作为
# 本地调试入口，把那个 dict 序列化为 JSON 打印。

SCRIPT_ID = "098"
# 维度一门禁接受的交付物扩展名：仅 .xlsx / .xlsm
_ACCEPTED_EXTS = (".xlsx", ".xlsm")
# 维度二加分项 max_delta 之和，用于维度一未通过 / 脚本异常时响应
MAX_SCORE_TOTAL = 45
# 维度二减分项 max_delta 之和（负值）
MIN_PENALTY_TOTAL = -32

# 维度二评分项按"结果追加顺序"的 max_delta 列表；末端把 results 里的
# "+N：..." / "-N：..." 字符串拆解为结构化 dim2_items 时用。顺序须与 _evaluate_impl 中
# results.append(...) 的顺序严格一致，共 31 项：前 17 项加分累加 = MAX_SCORE_TOTAL，
# 后 14 项减分累加 = MIN_PENALTY_TOTAL。
_MAX_DELTAS_ORDER = [
    1,   # 01 图表类型为 XY 散点图
    5,   # 02 横轴引用 C 列体验、纵轴引用 B 列关注度（不得引用副本）
    1,   # 03 20 个维度全部入图，综合均值不作为数据点
    5,   # 04 数据点标注可直接识别的维度名称
    1,   # 05 图表标题为「重要性-满意度象限图」
    3,   # 06 关注度 4.11 横向参考线，贯穿全幅
    3,   # 07 体验 3.63 纵向参考线，贯穿全幅
    1,   # 08 参考线非手绘、可见可辨
    5,   # 09 四象限成立（两轴在 3.63/4.11 相交划分四象限）
    3,   # 10 两轴上下限手动设定且量程收紧
    5,   # 11 图面洁净（除两条坐标轴外无其他可见图形元素）
    3,   # 12 中文宋体、英文与数字 Times New Roman
    3,   # 13 两轴刻度标签 Times New Roman 且 9 磅
    3,   # 14 数据标签全部 9 磅
    1,   # 15 两个坐标轴标题字号均 9 磅
    1,   # 16 横轴有轴标题，指向体验均值
    1,   # 17 纵轴有轴标题，指向关注度均值
    -5,  # 18 数据表内容
    -3,  # 19 维度行序
    -1,  # 20 字体与字号
    -5,  # 21 填充底纹
    -3,  # 22 字体颜色
    -1,  # 23 加粗状态
    -1,  # 24 对齐方式
    -1,  # 25 自动换行
    -3,  # 26 边框
    -1,  # 27 数字格式
    -1,  # 28 合并单元格
    -5,  # 29 工作表名称/位置/可见状态
    -1,  # 30 数据表内容范围
    -1,  # 31 默认视图设置
]
assert sum(d for d in _MAX_DELTAS_ORDER if d > 0) == MAX_SCORE_TOTAL
assert sum(d for d in _MAX_DELTAS_ORDER if d < 0) == MIN_PENALTY_TOTAL

# ---------------------------------------------------------------- 命名空间
_NS_C = "http://schemas.openxmlformats.org/drawingml/2006/chart"
_NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main"
_NS_C15 = "http://schemas.microsoft.com/office/drawing/2012/chart"
_NS_XDR = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
_NS_SS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"


def _cq(tag):
    return "{%s}%s" % (_NS_C, tag)


def _aq(tag):
    return "{%s}%s" % (_NS_A, tag)


def _c15q(tag):
    return "{%s}%s" % (_NS_C15, tag)


def _xdrq(tag):
    return "{%s}%s" % (_NS_XDR, tag)


def _ssq(tag):
    return "{%s}%s" % (_NS_SS, tag)


# ---------------------------------------------------------------- 应有状态基线
# 减分项的判定不依赖"原文件"，全部写成自描述的应有状态：表头文案、20 个维度名称、
# 40 个数值、字体字号、填充色、字体色、加粗、对齐、换行、边框、数字格式都在此固化。
_HEADER = ("评估维度", "关注度均值", "体验均值")

_DIM_NAMES = [
    "Q6公共空间清洁", "Q7便民设施完好", "Q8通行秩序维护", "Q9服务人员仪态",
    "Q10事项按时办结", "Q11维修响应效率", "Q12费用说明清晰", "Q13意见受理反馈",
    "Q14线上客服响应", "Q15突发情况处置", "Q16进度主动提醒", "Q17信息公开频次",
    "Q18岗位专业能力", "Q19沟通礼貌程度", "Q20出入安全保障", "Q21服务口碑信任",
    "Q22个性需求关注", "Q23居民建议采纳", "Q24邻里活动组织", "Q25公共权益维护",
]
# 关注度均值（B2:B21）
_ATTN = [4.18, 4.05, 4.11, 3.72, 4.24, 4.31, 4.16, 4.28, 4.02, 4.20,
         4.07, 3.91, 4.09, 4.13, 4.25, 4.19, 3.96, 4.08, 3.85, 4.22]
# 体验均值（C2:C21）
_EXP = [3.74, 3.58, 3.49, 3.81, 3.62, 3.69, 3.55, 3.66, 3.57, 3.83,
        3.51, 3.64, 3.72, 3.78, 3.70, 3.67, 3.48, 3.60, 3.45, 3.42]
_SUM_LABEL = "综合均值"
_SUM_ATTN = 4.105
_SUM_EXP = 3.625
# 参考线位置（显示值 4.11 / 3.63，由 4.105 / 3.625 两位小数取整得到）
_REF_Y = 4.11
_REF_X = 3.63
_CHART_TITLE = "重要性-满意度象限图"

# 字体、填充、字体色、边框、数字格式的应有状态
_FONT_NAME = "微软雅黑"
_FONT_SIZE = 10.0
_COLOR_HEADER = "FFFFFFFF"
_COLOR_BODY = "FF263238"
_COLOR_SUM = "FF1B3A3A"
_FILL_HEADER = "FF315E72"
_FILL_EVEN = "FFEEF5F2"      # 行号为偶数的维度行
_FILL_ODD = "FFFAFCFB"       # 行号为奇数的维度行
_FILL_SUM = "FFD7E7E2"
# A1:C22 的对齐与换行应有状态
_ALIGN_H = "center"
_ALIGN_V = "center"
_NUMFMT_VAL = "0.00"
_NUMFMT_TXT = "General"
# 全表只应有 5 条中等粗细边框，其余单元格四边无边框
_EXPECTED_BORDERS = {
    ("B2", "top"): "medium",
    ("C2", "top"): "medium",
    ("B21", "bottom"): "medium",
    ("C21", "bottom"): "medium",
    ("C22", "bottom"): "medium",
}

# 各减分项的抽检坐标；内容量大的项按细则只抽检其中若干小项
_SAMPLE_CONTENT = ["B1", "A9", "B14", "C20", "B22"]
_SAMPLE_ORDER = ["A2", "A11", "A21"]
_SAMPLE_FONT = ["B1", "A4", "C11", "A18", "B22"]
_SAMPLE_FILL = ["B1", "A4", "A5", "B22"]
_SAMPLE_COLOR = ["B1", "C11", "B22"]
_SAMPLE_BOLD = ["A1", "B1", "C1", "A22", "B22", "C22", "A5", "C15"]
_SAMPLE_ALIGN = ["A1", "A9", "B15", "A22"]
_SAMPLE_WRAP = ["A1", "A11", "A22"]
_SAMPLE_NUMFMT = ["B2", "C15", "B22"]


# ---------------------------------------------------------------- 通用小工具
def _split_coord(coord):
    """把 "B14" 拆成 ("B", 14)。"""
    m = re.match(r"^([A-Za-z]+)(\d+)$", coord.strip())
    if not m:
        raise ValueError("非法单元格坐标：%s" % coord)
    return m.group(1).upper(), int(m.group(2))


def _col_to_idx(col):
    """列字母转 1 起序号，"A"->1、"AA"->27。"""
    n = 0
    for ch in col.upper():
        n = n * 26 + (ord(ch) - 64)
    return n


def _num_eq(a, b, tol=5e-4):
    """数值比较；单元格里可能是 int/float/字符串数字，统一按浮点比。"""
    try:
        return abs(float(a) - float(b)) <= tol
    except (TypeError, ValueError):
        return False


def _as_text(v):
    if v is None:
        return ""
    return str(v).strip()


def _has_cjk(s):
    return any("一" <= ch <= "鿿" for ch in (s or ""))


def _has_latin(s):
    return any(("a" <= ch.lower() <= "z") or ch.isdigit() for ch in (s or ""))


def _norm_title(s):
    """标题比对前归一化：去空白、各种连字符统一为半角 "-"。"""
    s = re.sub(r"\s+", "", s or "")
    return re.sub(r"[‐‑‒–—―−－]", "-", s)


def _expected_cell_value(col, row):
    """给出 A1:C23 内任一坐标的应有内容，超出范围返回 None（应为空）。"""
    if col not in ("A", "B", "C"):
        return None
    ci = _col_to_idx(col) - 1
    if row == 1:
        return _HEADER[ci]
    if 2 <= row <= 21:
        i = row - 2
        return (_DIM_NAMES[i], _ATTN[i], _EXP[i])[ci]
    if row == 22:
        return (_SUM_LABEL, _SUM_ATTN, _SUM_EXP)[ci]
    return None


def _cell_value_ok(ws, coord):
    """抽检单元格内容：文本按去空白全等，数值按 5e-4 容差比较。"""
    col, row = _split_coord(coord)
    exp = _expected_cell_value(col, row)
    got = ws[coord].value
    if isinstance(exp, str):
        return _as_text(got) == exp
    return _num_eq(got, exp)


# ---------------------------------------------------------------- 打开与定位
def _open_excel_workbook(file_path):
    """同时取样式态与取值态：data_only=True 用于读公式缓存值。"""
    wb = openpyxl.load_workbook(file_path, data_only=False)
    try:
        wbv = openpyxl.load_workbook(file_path, data_only=True)
    except Exception:
        wbv = wb
    return wb, wbv


def _read_parts(file_path):
    """把评分需要的 XML 分部一次性读进内存，键为 zip 内路径。"""
    out = {"__names__": []}
    with zipfile.ZipFile(file_path) as zf:
        bad = zf.testzip()
        out["__bad__"] = bad
        out["__names__"] = list(zf.namelist())
        for name in out["__names__"]:
            low = name.lower()
            if (low.startswith("xl/charts/") or low.startswith("xl/drawings/")
                    or low.startswith("xl/worksheets/") or low.startswith("xl/theme/")
                    or low in ("xl/workbook.xml", "[content_types].xml")
                    or low.startswith("xl/externallinks/")):
                try:
                    out[name] = zf.read(name)
                except Exception:
                    out[name] = b""
    return out


def _xml(parts, name):
    """按分部名解析成 Element，失败返回 None。"""
    raw = parts.get(name)
    if not raw:
        return None
    try:
        return ET.fromstring(raw)
    except ET.ParseError:
        return None


def _find_data_sheet(wbv):
    """按内容定位数据表：表头三列 + 综合均值行；表名被改过也能找到。"""
    for ws in wbv.worksheets:
        head = tuple(_as_text(ws.cell(row=1, column=c).value) for c in (1, 2, 3))
        if head == _HEADER:
            return ws
    for ws in wbv.worksheets:
        col_a = {_as_text(ws.cell(row=r, column=1).value) for r in range(1, 40)}
        if len(col_a & set(_DIM_NAMES)) >= 15:
            return ws
    return None


# ---------------------------------------------------------------- 引用解析
_REF_RE = re.compile(
    r"^(?:'((?:[^']|'')+)'|([^!]+))!\$?([A-Za-z]+)\$?(\d+)"
    r"(?::\$?([A-Za-z]+)\$?(\d+))?$")


def _parse_ref(f):
    """"Sheet1!$C$2:$C$21" -> ("Sheet1", "C", 2, "C", 21)；解析失败返回 None。"""
    if not f:
        return None
    m = _REF_RE.match(f.strip())
    if not m:
        return None
    sheet = (m.group(1) or m.group(2) or "").replace("''", "'")
    c1, r1 = m.group(3).upper(), int(m.group(4))
    c2 = (m.group(5) or m.group(3)).upper()
    r2 = int(m.group(6) or m.group(4))
    return sheet, c1, r1, c2, r2


def _resolve_ref(wbv, f):
    """把引用按行优先展平成值列表；工作表不存在返回 None。"""
    ref = _parse_ref(f)
    if ref is None:
        return None
    sheet, c1, r1, c2, r2 = ref
    if sheet not in wbv.sheetnames:
        return None
    ws = wbv[sheet]
    i1, i2 = _col_to_idx(c1), _col_to_idx(c2)
    vals = []
    for r in range(min(r1, r2), max(r1, r2) + 1):
        for c in range(min(i1, i2), max(i1, i2) + 1):
            vals.append(ws.cell(row=r, column=c).value)
    return vals


def _cache_vals(cache):
    """读 numCache / strCache / numLit 里的 c:pt，按 idx 排好序。"""
    if cache is None:
        return []
    pts = []
    for pt in cache.findall(_cq("pt")):
        try:
            idx = int(pt.get("idx") or "0")
        except ValueError:
            idx = 0
        pts.append((idx, _as_text(pt.findtext(_cq("v")))))
    pts.sort(key=lambda t: t[0])
    return [v for _, v in pts]


def _read_num_data(el):
    """读 c:xVal / c:yVal：返回 (公式, 缓存值列表, 是否手工录入常量)。"""
    if el is None:
        return None, [], False
    nr = el.find(_cq("numRef"))
    if nr is not None:
        f = _as_text(nr.findtext(_cq("f")))
        return f, _cache_vals(nr.find(_cq("numCache"))), False
    nl = el.find(_cq("numLit"))
    if nl is not None:
        return None, _cache_vals(nl), True
    return None, [], False


def _to_floats(vals):
    out = []
    for v in vals or []:
        try:
            out.append(float(v))
        except (TypeError, ValueError):
            pass
    return out


# ---------------------------------------------------------------- 图表结构
def _pick_chart(parts):
    """取图表分部：优先含散点图的那个。返回 (分部名, 根节点)。"""
    names = sorted(n for n in parts.get("__names__", [])
                   if re.match(r"^xl/charts/chart\d*\.xml$", n.lower()))
    cands = []
    for n in names:
        root = _xml(parts, n)
        if root is not None:
            cands.append((n, root))
    for n, root in cands:
        if root.find(".//" + _cq("scatterChart")) is not None:
            return n, root
    return cands[0] if cands else (None, None)


def _plot_area(root):
    if root is None:
        return None
    return root.find(".//" + _cq("plotArea"))


def _all_series(root):
    """按 XML 顺序取全部 c:ser（含散点图之外的图种）。"""
    pa = _plot_area(root)
    if pa is None:
        return []
    return [s for s in pa.iter(_cq("ser"))]


def _ref_count(f):
    ref = _parse_ref(f)
    if ref is None:
        return 0
    _, c1, r1, c2, r2 = ref
    return (abs(r2 - r1) + 1) * (abs(_col_to_idx(c2) - _col_to_idx(c1)) + 1)


def _series_info(ser):
    xf, xc, xlit = _read_num_data(ser.find(_cq("xVal")))
    yf, yc, ylit = _read_num_data(ser.find(_cq("yVal")))
    n = max(_ref_count(xf), _ref_count(yf), len(xc), len(yc))
    return {"el": ser, "xf": xf, "yf": yf, "xcache": xc, "ycache": yc,
            "xlit": xlit, "ylit": ylit, "n": n}


def _pick_data_series(infos, sheet_name):
    """数据系列 = 点数 20 的那条；有多条时优先落在数据表上的那条。"""
    twenty = [s for s in infos if s["n"] == 20]
    for s in twenty:
        ref = _parse_ref(s["yf"] or "")
        if ref and ref[0] == sheet_name:
            return s
    if twenty:
        return twenty[0]
    return max(infos, key=lambda s: s["n"]) if infos else None


def _pick_axes(root):
    """按 c:axPos 区分横轴（b/t）与纵轴（l/r）。"""
    pa = _plot_area(root)
    if pa is None:
        return None, None, []
    axes = list(pa.findall(_cq("valAx"))) + list(pa.findall(_cq("catAx"))) \
        + list(pa.findall(_cq("dateAx")))
    ax_b = ax_l = None
    for ax in axes:
        pos = (ax.find(_cq("axPos")).get("val") if ax.find(_cq("axPos")) is not None else "")
        if pos in ("b", "t") and ax_b is None:
            ax_b = ax
        elif pos in ("l", "r") and ax_l is None:
            ax_l = ax
    return ax_b, ax_l, axes


def _ax_val(ax, tag, sub=None):
    """取 c:scaling/c:min|c:max 或 c:crossesAt 的 val，返回 float 或 None。"""
    if ax is None:
        return None
    node = ax.find(_cq(sub) + "/" + _cq(tag)) if sub else ax.find(_cq(tag))
    if node is None:
        return None
    try:
        return float(node.get("val"))
    except (TypeError, ValueError):
        return None


def _ax_scaling(ax):
    return _ax_val(ax, "min", "scaling"), _ax_val(ax, "max", "scaling")


def _ax_crosses(ax):
    """返回 (crossesAt 数值, crosses 关键字)；两者只会有一个生效。"""
    if ax is None:
        return None, None
    kw = ax.find(_cq("crosses"))
    return _ax_val(ax, "crossesAt"), (kw.get("val") if kw is not None else None)


# ---------------------------------------------------------------- 可见性判定
# 口径：存在但不可见的元素视为不存在（用户在成品图上看不到即不计）。
def _ln_state(spPr):
    """线条可见性：True 可见 / False 明确无填充 / None 未指定（继承）。"""
    if spPr is None:
        return None
    ln = spPr.find(_aq("ln"))
    if ln is None:
        return None
    if ln.find(_aq("noFill")) is not None:
        return False
    for t in ("solidFill", "gradFill", "pattFill", "blipFill"):
        if ln.find(_aq(t)) is not None:
            return True
    return None


def _line_visible(spPr, default=True):
    st = _ln_state(spPr)
    return default if st is None else st


def _fill_visible(spPr, default=False):
    """区域填充可见性，用于绘图区底色一类判断。"""
    if spPr is None:
        return default
    if spPr.find(_aq("noFill")) is not None:
        return False
    for t in ("solidFill", "gradFill", "pattFill", "blipFill"):
        if spPr.find(_aq(t)) is not None:
            return True
    return default


# ---------------------------------------------------------------- 字体与文本
def _theme_hans(parts):
    """中文未显式指定 a:ea 时，实际生效的是主题里 script="Hans" 的字体。"""
    for name in sorted(n for n in parts.get("__names__", [])
                       if n.lower().startswith("xl/theme/")):
        root = _xml(parts, name)
        if root is None:
            continue
        for scheme in ("minorFont", "majorFont"):
            node = root.find(".//" + _aq("fontScheme") + "/" + _aq(scheme))
            if node is None:
                continue
            for f in node.findall(_aq("font")):
                if f.get("script") == "Hans":
                    return f.get("typeface") or "宋体"
    return "宋体"


def _font_tuple(rpr):
    """(字号磅值, 西文字体, 中文字体)；未指定的返回 None。"""
    sz = None
    try:
        if rpr.get("sz"):
            sz = float(rpr.get("sz")) / 100.0
    except (TypeError, ValueError):
        sz = None
    lat = rpr.find(_aq("latin"))
    ea = rpr.find(_aq("ea"))
    return (sz,
            lat.get("typeface") if lat is not None else None,
            ea.get("typeface") if ea is not None else None)


def _collect_fonts(el):
    """收集 el 子树内所有 a:defRPr / a:rPr 的字体三元组。"""
    out = []
    if el is None:
        return out
    for node in el.iter():
        if node.tag in (_aq("defRPr"), _aq("rPr")):
            out.append(_font_tuple(node))
    return out


def _texts(el):
    if el is None:
        return ""
    return "".join(t.text or "" for t in el.iter(_aq("t")))


def _title_of(owner):
    """取 c:title 节点（坐标轴标题或图表标题）。"""
    if owner is None:
        return None
    return owner.find(_cq("title"))


def _title_text(owner):
    t = _title_of(owner)
    if t is None:
        return ""
    return _texts(t.find(_cq("tx"))) or _texts(t)


def _label_texts(ser, wbv, n):
    """还原每个数据点实际显示的标签文本；无标签的位置为 None。"""
    labels = [None] * max(n, 0)
    dLbls = ser.find(_cq("dLbls")) if ser is not None else None
    if dLbls is None:
        return labels
    rng_f, rng_cache = None, []
    for dr in dLbls.iter(_c15q("datalabelsRange")):
        f = _as_text(dr.findtext(_c15q("f")))
        if f:
            rng_f = f
        rng_cache = _cache_vals(dr.find(_c15q("dlblRangeCache")))
    if rng_f:
        vals = _resolve_ref(wbv, rng_f)
        if not vals:
            vals = rng_cache
        for i, v in enumerate(vals[:len(labels)]):
            labels[i] = _as_text(v) or None
    for dl in dLbls.findall(_cq("dLbl")):
        idx_node = dl.find(_cq("idx"))
        try:
            idx = int(idx_node.get("val"))
        except (AttributeError, TypeError, ValueError):
            continue
        if not (0 <= idx < len(labels)):
            continue
        dele = dl.find(_cq("delete"))
        if dele is not None and (dele.get("val") or "1") in ("1", "true"):
            labels[idx] = None
            continue
        rich = dl.find(_cq("tx") + "/" + _cq("rich"))
        txt = _texts(rich)
        if txt.strip():
            labels[idx] = txt.strip()
    return labels


def _chart_default_size(root):
    """图表整体默认字号（chartSpace 级 c:txPr）。"""
    if root is None:
        return None
    node = root.find(_cq("txPr"))
    for sz, _lat, _ea in _collect_fonts(node):
        if sz:
            return sz
    return None


def _label_sizes(ser, root):
    """数据标签实际生效的字号集合；一处未指定则回落到图表默认字号。"""
    dLbls = ser.find(_cq("dLbls")) if ser is not None else None
    sizes = {sz for sz, _l, _e in _collect_fonts(dLbls) if sz}
    if not sizes:
        d = _chart_default_size(root)
        sizes = {d if d else 10.0}
    return sizes


# ---------------------------------------------------------------- 对外接口
def evaluate(dir_path):
    """入口：在 dir_path 内定位被评估的 .xlsx/.xlsm 并返回结构化评分结果。"""
    result = {
        "id": SCRIPT_ID,
        "file_name": "",
        "status": "ok",
        "error": None,
        "dim1_pass": True,
        "dim1_reason": "",
        "dim2_items": [],
        "total_score": 0,
        "max_score": MAX_SCORE_TOTAL,
    }
    try:
        if not dir_path or not os.path.isdir(dir_path):
            raise FileNotFoundError("目录不存在：%s" % dir_path)
        cands = sorted(f for f in os.listdir(dir_path)
                       if f.lower().endswith(_ACCEPTED_EXTS)
                       and not f.startswith("~") and not f.startswith("."))
        if not cands:
            raise FileNotFoundError("目录内未找到 .xlsx/.xlsm 文件：%s" % dir_path)
        file_path = os.path.join(dir_path, cands[0])
        result["file_name"] = cands[0]
        return _evaluate_impl(file_path, result)
    except Exception as exc:  # noqa: BLE001
        result["status"] = "error"
        result["error"] = "%s: %s" % (type(exc).__name__, exc)
        # status 为 error 时未能完成判定，dim1_pass 留空，dim1_reason 也为空；
        # 异常摘要只放 error，不在 dim1_reason 里重复。
        result["dim1_pass"] = ""
        result["dim1_reason"] = ""
        result["dim2_items"] = []
        result["total_score"] = 0
        result["max_score"] = MAX_SCORE_TOTAL
        return result


def _marker_visible(ser):
    """散点标记是否可见：symbol=none 或标记本身无填充无线条即不可见。"""
    m = ser.find(_cq("marker")) if ser is not None else None
    if m is None:
        return True
    sym = m.find(_cq("symbol"))
    if sym is not None and sym.get("val") == "none":
        return False
    spPr = m.find(_cq("spPr"))
    if spPr is not None and spPr.find(_aq("noFill")) is not None and _ln_state(spPr) is False:
        return False
    return True


def _find_part(parts, low_name):
    for n in parts.get("__names__", []):
        if n.lower() == low_name:
            return n
    return None


def _is_tnr(name):
    return (name or "").strip().lower() == "times new roman"


def _is_song(name):
    return (name or "").strip() in ("宋体", "SimSun", "NSimSun", "宋体-简")


def _to_edge(coord, side):
    """把"某格某条边"归一化为网格线段，避免边框记在相邻格上被误判。"""
    col, row = _split_coord(coord)
    ci = _col_to_idx(col)
    if side == "top":
        return ("h", ci, row)
    if side == "bottom":
        return ("h", ci, row + 1)
    if side == "left":
        return ("v", ci, row)
    return ("v", ci + 1, row)


def _dim1_fail(result, reason):
    """维度一门禁未通过：完成度按 0 处理，不再逐项计分。"""
    result["dim1_pass"] = False
    result["dim1_reason"] = reason
    result["dim2_items"] = []
    result["total_score"] = 0
    result["max_score"] = MAX_SCORE_TOTAL
    return result


def _plus(results, hit, delta, rule):
    """加分项：命中记 +N，未命中记 +0。"""
    results.append("+%d：%s" % (delta if hit else 0, rule))


def _minus(results, violated, delta, rule):
    """减分项：违反记 -N，符合应有状态记 -0。"""
    results.append("-%d：%s" % (abs(delta) if violated else 0, rule))


def _evaluate_impl(file_path, result):
    results = []

    # ============ 维度一：可用与可修改性（门禁项，本身不计分）============
    # T1 文件为有效的 .xlsx / .xlsm 包且未损坏：扩展名在受支持范围内、zip 分部
    #    CRC 全通过、含 [Content_Types].xml、openpyxl 能完整加载。
    #    任一不成立则后续所有评分无意义。
    if not file_path.lower().endswith(_ACCEPTED_EXTS):
        return _dim1_fail(result, "交付物不是受支持的 Excel 文件类型（仅接受 .xlsx/.xlsm）")
    parts = _read_parts(file_path)
    names = parts.get("__names__", [])
    if parts.get("__bad__"):
        return _dim1_fail(result, "压缩包分部损坏：%s" % parts["__bad__"])
    if _find_part(parts, "[content_types].xml") is None:
        return _dim1_fail(result, "缺少 [Content_Types].xml，不是有效的 Excel（OOXML）包")
    try:
        wb, wbv = _open_excel_workbook(file_path)
    except Exception as exc:  # noqa: BLE001
        return _dim1_fail(result, "工作簿无法被正常加载：%s: %s" % (type(exc).__name__, exc))

    # T2 打开时不触发修复提示、图表能正常渲染：以"读到的每个 XML 分部都能解析成功、
    #    且存在可解析的图表分部"作为可渲染的等价判据（无渲染器时的稳定替代口径）。
    broken = [n for n in names if n in parts and n.lower().endswith(".xml")
              and _xml(parts, n) is None]
    if broken:
        return _dim1_fail(result, "XML 分部无法解析，打开时会触发修复：%s" % "、".join(broken[:3]))
    _chart_part, chart = _pick_chart(parts)
    if chart is None:
        return _dim1_fail(result, "工作簿内不存在可解析的图表分部")

    # T3 工作簿内存在完整的数据表：按内容定位（表头三列 + 20 个维度名），表名被改也能找到。
    ws_v = _find_data_sheet(wbv)
    if ws_v is None:
        return _dim1_fail(result, "工作簿内找不到含表头与 20 个维度的数据表")
    ws = wb[ws_v.title]
    sheet_name = ws_v.title

    # T4 图表是可编辑的图表对象而非图片：绘图分部里必须有 graphicFrame 承载图表。
    frames = 0
    for n in names:
        if re.match(r"^xl/drawings/drawing\d*\.xml$", n.lower()):
            d = _xml(parts, n)
            if d is not None:
                frames += len(list(d.iter(_xdrq("graphicFrame"))))
    if frames < 1:
        return _dim1_fail(result, "图表不是可编辑的图表对象（绘图分部内无 graphicFrame）")

    # T5 图表数据引用指向本工作簿内的有效区域：无 #REF!、无外部工作簿链接、
    #    引用里出现的工作表名都真实存在。
    if any(n.lower().startswith("xl/externallinks/") for n in names):
        return _dim1_fail(result, "工作簿存在外部链接，图表引用不在本簿内")
    bad_refs = []
    for f in chart.iter(_cq("f")):
        txt = _as_text(f.text)
        if not txt:
            continue
        if "#REF" in txt.upper():
            bad_refs.append(txt)
            continue
        for g1, g2 in re.findall(r"(?:'((?:[^']|'')+)'|([^\s,'!()+\-*/]+))!", txt):
            nm = (g1 or g2).replace("''", "'")
            if nm and nm not in wbv.sheetnames:
                bad_refs.append(txt)
                break
    if bad_refs:
        return _dim1_fail(result, "图表存在无效引用：%s" % "、".join(bad_refs[:3]))

    result["dim1_pass"] = True
    # 约定：dim1_reason 只承载"未通过原因摘要"，通过时留空。
    result["dim1_reason"] = ""

    # ================ 维度二：完成度（先加分项 17 条，后减分项 14 条）================
    # 图表事实一次性抽取，后面各项只做判定，不再重复解析 XML。
    infos = [_series_info(s) for s in _all_series(chart)]
    data_ser = _pick_data_series(infos, sheet_name)
    ax_b, ax_l, all_axes = _pick_axes(chart)
    x_min, x_max = _ax_scaling(ax_b)          # 横轴 = 体验
    y_min, y_max = _ax_scaling(ax_l)          # 纵轴 = 关注度
    cross_b, _kw_b = _ax_crosses(ax_b)        # 横轴画在纵轴的哪个取值上
    cross_l, _kw_l = _ax_crosses(ax_l)        # 纵轴画在横轴的哪个取值上
    st_node = chart.find(".//" + _cq("scatterStyle"))
    _st = st_node.get("val") if st_node is not None else "lineMarker"
    ser_line_default = _st in ("line", "lineMarker", "smooth", "smoothMarker")
    hans = _theme_hans(parts)

    def _sv(info, key):
        """取系列某一维的实际数值：优先按引用回表取，取不到用缓存。"""
        f = info[key + "f"]
        vals = _resolve_ref(wbv, f) if f else None
        if not vals:
            vals = info[key + "cache"]
        return _to_floats(vals)

    xs = _sv(data_ser, "x") if data_ser else []
    ys = _sv(data_ser, "y") if data_ser else []
    labels = _label_texts(data_ser["el"], wbv, 20) if data_ser else []

    # 数据表里实际躺着的 20 组（体验, 关注度）与维度名，用于点数、象限、标签比对
    sheet_pairs, sheet_names = [], []
    for r in range(2, 22):
        nm = _as_text(ws_v.cell(row=r, column=1).value)
        try:
            sheet_pairs.append((float(ws_v.cell(row=r, column=3).value),
                                float(ws_v.cell(row=r, column=2).value)))
        except (TypeError, ValueError):
            pass
        if nm:
            sheet_names.append(nm)
    sheet_exp = sorted(round(p[0], 4) for p in sheet_pairs)
    sheet_attn = sorted(round(p[1], 4) for p in sheet_pairs)

    # 01 (+1) 图表类型：必须是 XY 散点图，折线/柱形/气泡等均不算。
    _plus(results, chart.find(".//" + _cq("scatterChart")) is not None, 1,
          "图表类型为 XY 散点图")

    # 02 (+5) 轴绑定：横轴取数据表 C 列体验均值、纵轴取 B 列关注度均值，
    #    且必须是对数据表本体的单元格引用——引用副本表、另建数据块、
    #    numLit 手工录入数值都不给分。
    x_ref = _parse_ref(data_ser["xf"] or "") if data_ser else None
    y_ref = _parse_ref(data_ser["yf"] or "") if data_ser else None
    ok02 = bool(data_ser) and not data_ser["xlit"] and not data_ser["ylit"] \
        and x_ref is not None and y_ref is not None \
        and x_ref[0] == sheet_name and y_ref[0] == sheet_name \
        and x_ref[1] == x_ref[3] == "C" and y_ref[1] == y_ref[3] == "B" \
        and min(x_ref[2], x_ref[4]) == 2 and max(x_ref[2], x_ref[4]) == 21 \
        and min(y_ref[2], y_ref[4]) == 2 and max(y_ref[2], y_ref[4]) == 21
    _plus(results, ok02, 5,
          "横轴引用数据表 C 列体验均值、纵轴引用数据表 B 列关注度均值，"
          "且为对数据表本体 C2:C21 / B2:B21 的单元格引用")

    # 03 (+1) 入图范围：恰好 20 个数据点，与数据表 20 个维度一一对应，
    #    综合均值行不得作为数据点混入。
    sum_x, sum_y = ws_v["C22"].value, ws_v["B22"].value
    ok03 = len(xs) == 20 and len(ys) == 20 \
        and sorted(round(v, 4) for v in xs) == sheet_exp \
        and sorted(round(v, 4) for v in ys) == sheet_attn \
        and not any(_num_eq(a, sum_x) and _num_eq(b, sum_y) for a, b in zip(xs, ys))
    _plus(results, ok03, 1, "20 个维度全部入图且无多余点，综合均值不作为数据点")

    # 04 (+5) 数据点标注：每个点旁的标签要能直接读出维度名称；
    #    只标 Q 编号、只标坐标值、或标签缺失都不给分。
    hit04 = 0
    for nm in sheet_names:
        pure = re.sub(r"^Q\d+", "", nm).strip()
        if pure and any(pure in (lb or "") for lb in labels):
            hit04 += 1
    _plus(results, hit04 == 20 and len(sheet_names) == 20, 5,
          "20 个数据点均标注可直接识别的维度名称")

    # 05 (+1) 图表标题文字，连字符全角半角不计。
    atd = chart.find(".//" + _cq("autoTitleDeleted"))
    title_gone = atd is not None and (atd.get("val") in ("1", "true"))
    ok05 = (not title_gone) and \
        _norm_title(_title_text(chart.find(_cq("chart")))) == _norm_title(_CHART_TITLE)
    _plus(results, ok05, 1, "图表标题为「%s」" % _CHART_TITLE)

    # 06/07/08 参考线：两种正规做法都认——① 让坐标轴本身交叉到均值处（crossesAt），
    # ② 用引用单元格的辅助系列画一条贯穿全幅的线。返回四元组：
    # (做法, 是否可见, 是否单元格引用而非手工常量, 是否贯穿全幅)。
    def _ref_line(direction):
        if direction == "h":
            cross, want, ax = cross_b, _REF_Y, ax_b
            lo, hi = x_min, x_max
        else:
            cross, want, ax = cross_l, _REF_X, ax_l
            lo, hi = y_min, y_max
        if cross is not None and _num_eq(cross, want, 6e-3):
            vis = _line_visible(ax.find(_cq("spPr")), True) if ax is not None else False
            return "axis", vis, True, True
        for s in infos:
            if data_ser is not None and s is data_ser:
                continue
            xv, yv = _sv(s, "x"), _sv(s, "y")
            fixed, free = (yv, xv) if direction == "h" else (xv, yv)
            if len(fixed) >= 2 and all(_num_eq(v, want, 6e-3) for v in fixed):
                full = bool(free) and lo is not None and hi is not None \
                    and _num_eq(min(free), lo, 1e-6) and _num_eq(max(free), hi, 1e-6)
                vis = _line_visible(s["el"].find(_cq("spPr")), ser_line_default)
                return "series", vis, not (s["xlit"] or s["ylit"]), full
        return None, False, False, False

    m_h, vis_h, ref_h, full_h = _ref_line("h")
    m_v, vis_v, ref_v, full_v = _ref_line("v")

    # 06 (+3) 关注度 4.11 处的横向参考线，取值到位且贯穿整幅图宽。
    _plus(results, bool(m_h) and vis_h and full_h, 3,
          "关注度 %.2f 处有横向参考线且贯穿全幅" % _REF_Y)

    # 07 (+3) 体验 3.63 处的纵向参考线，取值到位且贯穿整幅图高。
    _plus(results, bool(m_v) and vis_v and full_v, 3,
          "体验 %.2f 处有纵向参考线且贯穿全幅" % _REF_X)

    # 08 (+1) 两条参考线的生成方式：坐标轴交叉或引用单元格的辅助系列，
    #    不能是画上去的线条/形状；且线条本身颜色粗细可辨。
    drawn_shapes = 0
    for n in names:
        if re.match(r"^xl/drawings/drawing\d*\.xml$", n.lower()):
            d = _xml(parts, n)
            if d is None:
                continue
            for t in ("sp", "cxnSp", "grpSp"):
                drawn_shapes += len(list(d.iter(_xdrq(t))))
    ok08 = bool(m_h) and bool(m_v) and ref_h and ref_v and vis_h and vis_v \
        and drawn_shapes == 0
    _plus(results, ok08, 1, "两条参考线由坐标轴交叉或引用单元格的辅助系列生成、非手绘且可见可辨")

    # 09 (+5) 四象限：必须是两条坐标轴本身在体验 3.63、关注度 4.11 处相交，
    #    交点落在两轴量程内部，并且四个象限里都真的有数据点。
    #    坐标轴贴边（crosses=autoZero/min/max）+ 辅助线分区不算四象限。
    quad_axes = (cross_b is not None and _num_eq(cross_b, _REF_Y, 6e-3)
                 and cross_l is not None and _num_eq(cross_l, _REF_X, 6e-3))
    inside = (None not in (x_min, x_max, y_min, y_max)
              and x_min < _REF_X < x_max and y_min < _REF_Y < y_max)
    quad = [0, 0, 0, 0]
    for ex, at in sheet_pairs:
        if ex > _REF_X and at > _REF_Y:
            quad[0] += 1
        elif ex < _REF_X and at > _REF_Y:
            quad[1] += 1
        elif ex < _REF_X and at < _REF_Y:
            quad[2] += 1
        elif ex > _REF_X and at < _REF_Y:
            quad[3] += 1
    _plus(results, quad_axes and inside and all(v > 0 for v in quad), 5,
          "两条坐标轴在体验 %.2f、关注度 %.2f 处相交划分四象限，四个象限均有数据点"
          % (_REF_X, _REF_Y))

    # 10 (+3) 量程：四个端点都手动设过（scaling 里同时写了 min 与 max），
    #    且横轴量程不超过体验极差的 3 倍、纵轴不超过关注度极差的 3 倍，
    #    保证点群铺开而不是挤成一团。
    x_limit = 3 * (max(_EXP) - min(_EXP))
    y_limit = 3 * (max(_ATTN) - min(_ATTN))
    manual = None not in (x_min, x_max, y_min, y_max)
    ok10 = manual and (x_max - x_min) <= x_limit + 1e-9 \
        and (y_max - y_min) <= y_limit + 1e-9
    _plus(results, ok10, 3,
          "两轴上下限均手动设定，横轴量程不超过 %.2f、纵轴量程不超过 %.2f" % (x_limit, y_limit))

    # 11 (+5) 图面洁净：除横纵两条坐标轴外不应再有可见图形元素。
    #    统一口径——存在但不可见（noFill / symbol=none）视为不存在。
    extras = []
    for ax in all_axes:
        for tag, nm in (("majorGridlines", "网格线"), ("minorGridlines", "网格线")):
            g = ax.find(_cq(tag))
            if g is not None and _line_visible(g.find(_cq("spPr")), True):
                extras.append(nm)
    pa = _plot_area(chart)
    if pa is not None and _line_visible(pa.find(_cq("spPr")), False):
        extras.append("绘图区边框")
    # 注：图表对象最外层那圈浅灰描边是图表容器自带的默认外框，不属于图内元素，不计。
    if chart.find(".//" + _cq("legend")) is not None:
        extras.append("图例")
    if data_ser is not None and \
            _line_visible(data_ser["el"].find(_cq("spPr")), ser_line_default):
        extras.append("数据点连线")
    for s in infos:
        if data_ser is not None and s is data_ser:
            continue
        if _line_visible(s["el"].find(_cq("spPr")), ser_line_default) \
                or _marker_visible(s["el"]):
            extras.append("附加系列图形")
            break
    for t in chart.iter(_cq("trendline")):
        if _line_visible(t.find(_cq("spPr")), True):
            extras.append("趋势线")
    for e in chart.iter(_cq("errBars")):
        if _line_visible(e.find(_cq("spPr")), True):
            extras.append("误差线")
    if len(all_axes) > 2:
        extras.append("次坐标轴")
    if drawn_shapes:
        extras.append("手绘线条或形状")
    extras = sorted(set(extras))
    _plus(results, not extras, 5,
          "图面洁净，除两条坐标轴外无其他可见图形元素")

    # 12~15 字体与字号。判定口径：某元素文字里没有中文就不追究中文字体，
    # 没有西文/数字就不追究西文字体；中文未显式指定 a:ea 时按主题 Hans 字体生效。
    def _sizes_of(el):
        s = {sz for sz, _l, _e in _collect_fonts(el) if sz}
        if not s:
            d = _chart_default_size(chart)
            s = {d if d else 10.0}
        return s

    def _font_check(text, el):
        fonts = _collect_fonts(el)
        lats = {l for _s, l, _e in fonts if l}
        eas = {e for _s, _l, e in fonts if e}
        ok = True
        if _has_latin(text):
            ok = ok and bool(lats) and all(_is_tnr(l) for l in lats)
        if _has_cjk(text):
            eff = eas or {l for l in lats if _is_song(l)} or {hans}
            ok = ok and all(_is_song(e) for e in eff)
        return ok

    tick_txpr = [(ax, ax.find(_cq("txPr"))) for ax in (ax_b, ax_l) if ax is not None]
    label_text = "".join(lb or "" for lb in labels)
    dlbls_el = data_ser["el"].find(_cq("dLbls")) if data_ser is not None else None
    titles = [(ax, _title_of(ax), _title_text(ax)) for ax in (ax_b, ax_l) if ax is not None]

    # 12 (+3) 中文宋体、英文与数字 Times New Roman（覆盖刻度标签、数据标签、轴标题）。
    ok12 = len(tick_txpr) == 2 and len(titles) == 2
    for _ax, tp in tick_txpr:
        ok12 = ok12 and _font_check("0123456789", tp)
    ok12 = ok12 and _font_check(label_text, dlbls_el) and bool(label_text)
    for _ax, tel, ttxt in titles:
        ok12 = ok12 and bool(ttxt) and _font_check(ttxt, tel)
    _plus(results, ok12, 3, "刻度标签、数据标签、轴标题的中文为宋体，英文与数字为 Times New Roman")

    # 13 (+3) 两轴刻度标签：Times New Roman 且 9 磅（小五）。
    ok13 = len(tick_txpr) == 2
    for _ax, tp in tick_txpr:
        lats = {l for _s, l, _e in _collect_fonts(tp) if l}
        ok13 = ok13 and _sizes_of(tp) == {9.0} and bool(lats) and all(_is_tnr(l) for l in lats)
    _plus(results, ok13, 3, "两轴刻度标签全部为 Times New Roman 且字号 9 磅")

    # 14 (+3) 数据标签字号全部 9 磅，不能有漏改的。
    _plus(results, data_ser is not None and _label_sizes(data_ser["el"], chart) == {9.0},
          3, "20 个数据标签字号全部为 9 磅")

    # 15 (+1) 两个坐标轴标题字号均 9 磅。
    ok15 = len(titles) == 2 and all(_sizes_of(t) == {9.0} for _a, t, _x in titles)
    _plus(results, ok15, 1, "两个坐标轴标题字号均为 9 磅")

    # 16/17 (+1/+1) 轴标题按实际绑定判断：横轴绑体验，纵轴绑关注度。
    t_b = next((x for a, _t, x in titles if a is ax_b), "")
    t_l = next((x for a, _t, x in titles if a is ax_l), "")
    _plus(results, "体验" in t_b, 1, "横轴有轴标题且内容指向体验均值")
    _plus(results, "关注" in t_l, 1, "纵轴有轴标题且内容指向关注度均值")

    # ---------------- 减分项：数据表被动了不该动的地方 ----------------
    # 全部写成"应有状态"，不依赖原文件比对；内容量大的项按细则抽检其中若干小项，
    # 抽中任意一处不符即整项扣分。

    # 18 (-5) 表内文字与数值。
    bad18 = [c for c in _SAMPLE_CONTENT if not _cell_value_ok(ws_v, c)]
    _minus(results, bool(bad18), -5,
           "数据表内容应为：A1「评估维度」、B1「关注度均值」、C1「体验均值」；"
           "A2:A21 为 Q6公共空间清洁 至 Q25公共权益维护 共 20 个维度名称，"
           "B2:B21 为各维度关注度均值、C2:C21 为各维度体验均值；"
           "A22「综合均值」、B22 为 4.105、C22 为 3.625")

    # 19 (-3) 维度出现的先后顺序。
    bad19 = []
    for c in _SAMPLE_ORDER:
        col, row = _split_coord(c)
        if _as_text(ws_v[c].value) != _expected_cell_value(col, row):
            bad19.append(c)
    _minus(results, bool(bad19), -3,
           "20 个维度自第 2 行起按 Q6、Q7…Q25 的编号顺序逐行排列，"
           "即 A2 为 Q6公共空间清洁、A11 为 Q15突发情况处置、A21 为 Q25公共权益维护")

    # 20 (-1) 字体与字号。
    bad20 = []
    for c in _SAMPLE_FONT:
        f = ws[c].font
        if (f.name or "") != _FONT_NAME or not _num_eq(f.size or 0, _FONT_SIZE, 0.01):
            bad20.append(c)
    _minus(results, bool(bad20), -1,
           "A1:C22 全部单元格字体为微软雅黑、字号 10 磅")

    # 21 (-5) 填充底纹：表头深色、维度行隔行浅色、汇总行独立色。
    def _exp_fill(row):
        if row == 1:
            return _FILL_HEADER
        if 2 <= row <= 21:
            return _FILL_EVEN if row % 2 == 0 else _FILL_ODD
        if row == 22:
            return _FILL_SUM
        return None

    def _got_fill(coord):
        fl = ws[coord].fill
        if fl is None or fl.fill_type != "solid":
            return None
        rgb = getattr(fl.fgColor, "rgb", None)
        return rgb.upper() if isinstance(rgb, str) else None

    bad21 = [c for c in _SAMPLE_FILL
             if _got_fill(c) != _exp_fill(_split_coord(c)[1])]
    _minus(results, bool(bad21), -5,
           "填充底纹为：第 1 行 %s，第 2 至 21 行按行号偶数 %s、奇数 %s 隔行交替，"
           "第 22 行 %s"
           % (_FILL_HEADER, _FILL_EVEN, _FILL_ODD, _FILL_SUM))

    # 22 (-3) 字体颜色：表头白字，维度行深灰蓝，汇总行深墨绿。
    def _exp_color(row):
        if row == 1:
            return _COLOR_HEADER
        if 2 <= row <= 21:
            return _COLOR_BODY
        if row == 22:
            return _COLOR_SUM
        return None

    def _got_color(coord):
        col = ws[coord].font.color
        rgb = getattr(col, "rgb", None) if col is not None else None
        return rgb.upper() if isinstance(rgb, str) else None

    bad22 = [c for c in _SAMPLE_COLOR
             if _got_color(c) != _exp_color(_split_coord(c)[1])]
    _minus(results, bool(bad22), -3,
           "字体颜色为：第 1 行 %s，第 2 至 21 行 %s，第 22 行 %s"
           % (_COLOR_HEADER, _COLOR_BODY, _COLOR_SUM))

    # 23 (-1) 加粗状态：只有表头行与汇总行加粗。
    bad23 = []
    for c in _SAMPLE_BOLD:
        row = _split_coord(c)[1]
        if bool(ws[c].font.bold) != (row in (1, 22)):
            bad23.append(c)
    _minus(results, bool(bad23), -1,
           "第 1 行与第 22 行加粗，第 2 至 21 行不加粗")

    # 24 (-1) 对齐方式：水平居中、垂直居中。
    bad24 = []
    for c in _SAMPLE_ALIGN:
        al = ws[c].alignment
        if (al.horizontal or "") != _ALIGN_H or (al.vertical or "") != _ALIGN_V:
            bad24.append(c)
    _minus(results, bool(bad24), -1,
           "A1:C22 单元格水平居中、垂直居中")

    # 25 (-1) 自动换行：A1:C22 全部开启。
    bad25 = [c for c in _SAMPLE_WRAP if not ws[c].alignment.wrap_text]
    _minus(results, bool(bad25), -1,
           "A1:C22 单元格开启自动换行")

    # 26 (-3) 边框：整表只有 5 段中等粗细横线，其余各边无边框。
    #    边框可能记在相邻单元格上，先归一化成网格线段再比对，避免误判。
    exp_edges = {}
    for (coord, side), style in _EXPECTED_BORDERS.items():
        exp_edges[_to_edge(coord, side)] = style
    got_edges = {}
    max_r = min(max(ws.max_row, 23), 60)
    max_c = min(max(ws.max_column, 3), 12)
    for row in ws.iter_rows(min_row=1, max_row=max_r, min_col=1, max_col=max_c):
        for cell in row:
            bd = cell.border
            if bd is None:
                continue
            for side in ("left", "right", "top", "bottom"):
                st = getattr(getattr(bd, side), "style", None)
                if st:
                    got_edges[_to_edge(cell.coordinate, side)] = st
    _minus(results, got_edges != exp_edges, -3,
           "整表只有 5 段中等粗细横线边框：B2 与 C2 的上边、B21 与 C21 的下边、"
           "C22 的下边；其余单元格四边均无边框")

    # 27 (-1) 数字格式：两列均值保留两位小数，维度名列为常规。
    bad27 = []
    for c in _SAMPLE_NUMFMT:
        col, row = _split_coord(c)
        want = _NUMFMT_VAL if (col in ("B", "C") and 2 <= row <= 22) else _NUMFMT_TXT
        if (ws[c].number_format or "") != want:
            bad27.append(c)
    _minus(results, bool(bad27), -1,
           "B2:C22 数字格式为 0.00（两位小数），A 列为常规")

    # 28 (-1) 合并单元格：数据表内一个都不应有。
    _minus(results, len(list(ws.merged_cells.ranges)) > 0, -1,
           "数据表内无合并单元格")

    # 29 (-5) 工作表名称、位置与可见状态。
    ok29 = (sheet_name == "Sheet1" and wb.sheetnames[0] == sheet_name
            and (wb[sheet_name].sheet_state or "visible") == "visible")
    _minus(results, not ok29, -5,
           "数据表工作表名为 Sheet1，位于工作簿第 1 个位置，状态为可见")

    # 30 (-1) 数据表内容范围：只有 A1:C22 有内容，第 23 行及 D 列以后为空。
    #    图表浮于单元格之上、不占单元格，不计入本项。
    out_range = []
    for row in ws_v.iter_rows(min_row=1, max_row=min(max(ws_v.max_row, 23), 200),
                              min_col=1, max_col=min(max(ws_v.max_column, 3), 30)):
        for cell in row:
            v = cell.value
            if v is None or (isinstance(v, str) and not v.strip()):
                continue
            if cell.column > 3 or cell.row > 22:
                out_range.append(cell.coordinate)
    _minus(results, bool(out_range), -1,
           "数据表仅 A1:C22 有内容，第 23 行及 D 列以后无任何文字、数值或备注")

    # 31 (-1) 视图设置保持默认：显示网格线、无冻结/拆分窗格、无自动筛选、
    #    无工作表标签颜色、无定义名称；覆盖工作簿内所有工作表。
    #    选中状态、光标位置、窗口尺寸这类不影响内容的痕迹不计。
    view_bad = []
    for n in sorted(nm for nm in names
                    if re.match(r"^xl/worksheets/sheet\d+\.xml$", nm.lower())):
        raw = (parts.get(n) or b"").decode("utf-8", "ignore")
        if re.search(r'showGridLines="(0|false)"', raw):
            view_bad.append("网格线被隐藏")
        if re.search(r"<[A-Za-z0-9]*:?pane[ />]", raw):
            view_bad.append("存在冻结或拆分窗格")
        if re.search(r"<[A-Za-z0-9]*:?autoFilter[ >]", raw):
            view_bad.append("存在自动筛选")
        if re.search(r"<[A-Za-z0-9]*:?tabColor[ />]", raw):
            view_bad.append("工作表标签被着色")
    wbxml_name = _find_part(parts, "xl/workbook.xml")
    wbxml_raw = (parts.get(wbxml_name) or b"").decode("utf-8", "ignore") if wbxml_name else ""
    if re.search(r"<[A-Za-z0-9]*:?definedName[ >]", wbxml_raw):
        view_bad.append("存在定义名称")
    view_bad = sorted(set(view_bad))
    _minus(results, bool(view_bad), -1,
           "工作簿内各工作表视图为默认状态：显示网格线、无冻结或拆分窗格、无自动筛选、"
           "无标签颜色、无定义名称")

    # ---------------- 把 results 拆解为结构化 dim2_items ----------------
    if len(results) != len(_MAX_DELTAS_ORDER):
        raise RuntimeError("dim2 项数不一致：results=%d vs order=%d"
                           % (len(results), len(_MAX_DELTAS_ORDER)))
    dim2_items = []
    total_score = 0
    for raw, max_delta in zip(results, _MAX_DELTAS_ORDER):
        head, _, rule = raw.partition("：")
        try:
            delta = int(head.replace("+", "").strip())
        except ValueError:
            delta = 0
        hit = delta != 0
        dim2_items.append({"rule": rule, "max_delta": max_delta,
                           "delta": max_delta if hit else 0,
                           "hit": hit})
        total_score += (max_delta if hit else 0)
    result["dim2_items"] = dim2_items
    # 加分减完之后不设负分，完成度分母仍为加分项上限 MAX_SCORE_TOTAL。
    result["total_score"] = max(0, total_score)
    result["max_score"] = MAX_SCORE_TOTAL
    return result


if __name__ == "__main__":
    _target = sys.argv[1] if len(sys.argv) > 1 else os.path.dirname(os.path.abspath(__file__))
    try:
        _out = evaluate(_target)
    except Exception:  # noqa: BLE001
        _out = {"id": SCRIPT_ID, "file_name": "", "status": "error",
                "error": traceback.format_exc(limit=3),
                "dim1_pass": "", "dim1_reason": "", "dim2_items": [],
                "total_score": 0, "max_score": MAX_SCORE_TOTAL}
    print(json.dumps(_out, ensure_ascii=False, indent=2))
