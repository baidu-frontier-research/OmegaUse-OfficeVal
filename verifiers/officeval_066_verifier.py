#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# pyright: reportUnknownMemberType=false, reportUnknownVariableType=false
# pyright: reportUnknownArgumentType=false, reportMissingImports=false
# pyright: reportOptionalMemberAccess=false, reportGeneralTypeIssues=false
#
# 任务 066：在"表格图表关联数据.pptx"（原 4 页）末尾新增第 5 页（10 行 × 8 列指标表格）
# 与第 6 页（7 系列带数据标记折线图），并保持第 1 至第 4 页原样。
#
# 维度1（可用与可修改性）：交付 1 个可打开的 .pptx，6 页，新增表格为可逐格编辑的原生
#   DrawingML 表格，新增折线图为原生 chart 部件（含 ppt/embeddings 内嵌工作簿），
#   版面不越界。任一不满足则 dim2 全部不计分、total_score = 0。
# 维度2（完成度）：19 个加分项（合计 65 分）+ 10 个减分项（触发则扣分）。
#
# 细则去重原则：同一处缺陷只在一处记账。为此做了四组合并/收窄——
#   ①"套用表格样式"与"表头深色底纹"合并为一项（本主题 accent1=4F81BD 已在深色白名单内，
#     套样式即必然产生深色表头，两项无法独立判别）；"行列分隔"另有显式框线通路，保留独立项。
#     细则分值只用 1/3/5 三档，合并项按最高档记 5 分（原两项各 3 分）而不是取和记 6 分。
#   ②"前4页标题文案"只由 M06 记账，不再进入 P01 与 M01；M01 收窄为只判页面尺寸
#     （页数与新增页位置由维度1 与 P01 覆盖）。
#   ③M09 收窄为只判"数值追加单位符号"——取整、单位换算、重排年份在 P03/P15 已必然失分；
#     M10 收窄为只扫数据区，表头格的手动换行由 P08 以"列宽足以容纳表头文字"独家记账。
#   ④P14 不再复述维度1 的门槛（有 chart 部件/lineChart/内嵌工作簿），改判维度1 未覆盖的
#     "单一图表类型 + 数据标记 + 数值经 numRef 引用工作簿而非 numLit 硬编码"。
#
# 按接口统一约定，本模块不修改全局 sys.stdout，也不 print 主结果。主结果统一由
# evaluate(dir_path) -> dict 返回；入参为"脚本所在目录的路径"，模块自行在该目录下
# 定位待评测的 .pptx 文档。本文件仅在 __main__ 下作为本地调试入口，才把 dict
# 序列化为 JSON 打印。
#
# 说明：本机不具备 pptx→PDF 的渲染能力（仅有 WPS，无头转换与 COM 自动化均不可用），
# 因此所有判定均落在 OOXML 层面（python-pptx + lxml），并在必要处按"办公软件的实际
# 渲染效果"放宽/收紧阈值，逐项在注释中标注"与办公软件对齐"。

import os
import re
import sys
import json

from lxml import etree
from pptx import Presentation

SCRIPT_ID = "066"
MAX_SCORE_TOTAL = 65

# 每项最大增/减分，顺序必须与 _evaluate_impl 中 results 的追加顺序完全一致。
_MAX_DELTAS_ORDER = [
    3,   # P01 新增第5、6页并置于末尾，表格页在前、折线图页在后，前4页版面构成不变
    5,   # P02 第5页 10行×8列，8个表头列名与列序完整
    5,   # P03 首列9个连续年份 + 9×7=63 个数值逐格一致
    5,   # P04 套用与第1、3页相同的表格样式（firstRow/bandRow）并因此呈深色表头底纹
    3,   # P05 表头白色加粗11磅、数据微软雅黑≥9.5磅
    5,   # P06 行列分隔清晰可辨
    3,   # P07 全部单元格水平居中且垂直居中
    3,   # P08 表格宽度≈10972800EMU、左右留白均衡、指标列等宽且不折行
    1,   # P09 第5、6页页面标题文本框（微软雅黑24磅加粗233142，题名点题）
    1,   # P10 第5页左下角量纲说明（7–8磅，说明百分比与综合指数单位）
    1,   # P11 第6页左下角极值说明（智慧门店响应率 2024 年 78.9）
    3,   # P12 第5、6页左上/右下平行四边形装饰图形
    5,   # P13 单一 lineChart + 各系列带数据标记 + 数值经 numRef 引用内嵌工作簿
    5,   # P14 7条系列名称一一对应 + 63个数据点逐点一致
    3,   # P15 分类轴9年升序、数值轴0–100且主刻度等距、数据点不贴顶贴底
    3,   # P16 图表标题关闭、图例在下方且不覆盖绘图区
    3,   # P17 折线线宽≥2.25磅、坐标轴刻度文字9磅
    3,   # P18 7条折线颜色两两不同且沿用蓝/橙/紫主色板
    5,   # P19 图表画幅（宽≥10000000EMU、高≥4700000EMU）
    -3,  # M01 页面尺寸（16:9 版式基准）被改动
    -5,  # M02 第1、3页表格的行列数、表头或九年数值被改动
    -1,  # M03 第1、3页表格的位置尺寸、框线、底纹、字号被改动
    -3,  # M04 第2、4页折线图的系列构成或数值轴刻度被改动
    -1,  # M05 第2、4页折线图的位置尺寸、系列颜色、标记、线宽、图例、数据标签被改动
    -1,  # M06 第1至第4页标题文字被改写（或第1、2页标题格式被改动）
    -1,  # M07 第1至第4页左下角说明文字被改写、字号或颜色被改动
    -1,  # M08 第1至第4页平行四边形装饰图形被删除、改色、移位或改尺寸
    -3,  # M09 新增表格数值单元格追加 %／％ 等单位符号
    -1,  # M10 新增表格的数据单元格混入制表符、连续空格或手动换行符（表头折行归 P08）
]
assert sum(d for d in _MAX_DELTAS_ORDER if d > 0) == MAX_SCORE_TOTAL

# ---------------------------------------------------------------- 期望数据 ----
# 指令给定的数据表：8 个表头 + 9 年 × 7 项指标。
NEW_HEADERS = [
    "年份",
    "智慧门店响应率（%）",
    "线上订单占比（%）",
    "自助结算覆盖率（%）",
    "绿色到店客流占比（%）",
    "低碳配送完成率（%）",
    "末端接驳便利度（%）",
    "运营协同指数（0–100）",
]
NEW_YEARS = [2016, 2017, 2018, 2019, 2020, 2021, 2022, 2023, 2024]
NEW_ROWS = [
    [31.8, 18.9, 24.6, 27.3, 22.6, 33.4, 29.7],
    [34.9, 23.7, 29.8, 31.5, 26.1, 36.8, 34.1],
    [42.6, 27.9, 35.2, 34.4, 30.6, 39.6, 39.8],
    [45.7, 33.8, 40.9, 37.2, 35.9, 43.8, 44.7],
    [53.4, 41.7, 49.6, 42.8, 44.3, 50.6, 52.9],
    [58.6, 48.8, 55.7, 47.6, 49.7, 55.1, 58.4],
    [66.7, 56.9, 62.6, 55.8, 57.4, 63.7, 66.2],
    [71.6, 62.4, 69.8, 61.6, 64.9, 70.2, 72.1],
    [78.9, 68.1, 76.7, 67.3, 72.6, 74.9, 78.2],
]

# 第 1 页 / 第 3 页表格（须保持不变）
BASE_T1_HEADERS = ["年份", "北岸片区", "云岭片区", "星湾片区"]
BASE_T1_COLS = [
    [48.2, 41.6, 55.4, 49.8, 63.1, 68.5, 73.2, 68.9, 82.4],
    [35.7, 40.3, 52.1, 47.0, 60.8, 54.6, 70.5, 65.1, 69.6],
    [22.8, 36.5, 30.3, 33.9, 38.7, 41.2, 56.4, 51.0, 56.3],
]
BASE_T3_HEADERS = ["年份", "区域均值", "区域均衡度", "运营转化指数"]
BASE_T3_COLS = [
    [35.6, 39.5, 45.9, 43.6, 54.2, 54.8, 66.7, 61.7, 69.4],
    [74.6, 94.9, 74.9, 84.1, 75.6, 72.7, 83.2, 82.1, 73.9],
    [48.1, 57.2, 55.2, 56.6, 61.0, 60.5, 72.0, 68.2, 70.8],
]

# 第 1 至第 4 页的标题与左下角说明文字（须保持不变）
BASE_TITLES = [
    "门店数字化运营指数",
    "门店数字化运营指数趋势",
    "门店数字化运营关联指标",
    "门店数字化运营关联指标趋势",
]
BASE_NOTES = [
    "单位：综合指数（0–100）；",
    "最高点：北岸片区在2024年达到 82.4",
    "单位：综合指数（0–100）；关联数据由第一个表格计算。",
    "最高点：区域均衡度在2017年达到 94.9",
]
BASE_NOTE_SZ = [800, 900, 700, 700]           # 依次 8磅、9磅、7磅、7磅
BASE_TITLE_STRICT = {1: True, 2: True, 3: False, 4: False}  # 仅第1、2页标题带显式字体属性

# 第 1 至第 4 页的几何基准（EMU）
BASE_TBL_OFF = (594360, 1078992)
BASE_TBL_EXT = (10972800, 4736592)
BASE_CHT_OFF = (676656, 1033271)
BASE_CHT_EXT = (10744200, 4800600)
DECO_TL_OFF = (-228600, -182880)
DECO_TL_EXT = (2834640, 502920)
DECO_BR_OFF = (9784080, 5989320)
DECO_BR_EXT = (2926080, 347472)
DECO_TL_FILL = "E2EFFA"
DECO_BR_FILL = "E1F4F1"
BASE_SER_COLORS = ["2F66A4", "D9962A", "7453AA"]

# 表格样式：Medium Style 2 - Accent 1（第1、3页表格所用，新增表格须沿用）
TABLE_STYLE_ID = "{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}"
# 内置表格样式在办公软件中为首行提供 accent1 底纹 + 白色加粗文字、镶边行提供
# accent1 浅色/白色隔行带与白色框线；这些效果不写进 slide XML，需在判定时补齐。
BUILTIN_BANDED_STYLES = {
    TABLE_STYLE_ID: {"header_fill": "accent1", "header_color": "FFFFFF", "header_bold": True},
}
HEADER_DARK_OK = {"2F66A4", "269B91", "4F81BD"}   # 第1页蓝 / 第3页青 / 样式 Accent1 蓝

EMU_PER_INCH = 914400
SLIDE_W = 12192000
SLIDE_H = 6858000

A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
C_NS = "http://schemas.openxmlformats.org/drawingml/2006/chart"
P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"

# ------------------------------------------------------------ 通用 OOXML 工具 ----


def _a(tag):
    return "{%s}%s" % (A_NS, tag)


def _c(tag):
    return "{%s}%s" % (C_NS, tag)


_WS_RE = re.compile(r"[\s 　]+")
_NUM_RE = re.compile(r"-?\d+(?:\.\d+)?")
_DASH_RE = re.compile("[‐-―−]")   # 各类连字符/短横/负号统一成 '-'


def _norm(text):
    """文本归一化：去掉全部空白（含制表符/全角空格）、半角括号转全角、各类连字符统一。

    与办公软件对齐：单元格里的多余空格、制表符、软换行在阅读时并不改变"列名是否正确"，
    因此列名/文案比对在归一化后进行；空格与制表符本身另由 M10 单独扣分。
    """
    if text is None:
        return ""
    s = text.replace("(", "（").replace(")", "）")
    s = _DASH_RE.sub("-", s)
    return _WS_RE.sub("", s)


def _num(text):
    """从单元格文本中取第一个数值；取不到返回 None。"""
    if text is None:
        return None
    m = _NUM_RE.search(text.replace(",", "").replace("，", ""))
    return float(m.group()) if m else None


def _same(v1, v2, eps=1e-6):
    return v1 is not None and v2 is not None and abs(v1 - v2) <= eps


def _near(v1, v2, tol):
    return v1 is not None and v2 is not None and abs(v1 - v2) <= tol


def _theme_colors(prs):
    """读取 ppt/theme/theme1.xml 的配色，用于把 schemeClr 解析成实际 RGB。"""
    mp = {}
    try:
        parts = list(prs.part.package.iter_parts())
    except Exception:
        parts = []
    for part in parts:
        if not str(part.partname).endswith("theme1.xml"):
            continue
        try:
            root = etree.fromstring(part.blob)
        except Exception:
            continue
        scheme = root.find(".//" + _a("clrScheme"))
        if scheme is None:
            continue
        for child in scheme:
            name = etree.QName(child).localname
            srgb = child.find(_a("srgbClr"))
            sys_c = child.find(_a("sysClr"))
            if srgb is not None and srgb.get("val"):
                mp[name] = srgb.get("val").upper()
            elif sys_c is not None and sys_c.get("lastClr"):
                mp[name] = sys_c.get("lastClr").upper()
        break
    for alias, real in (("tx1", "dk1"), ("bg1", "lt1"), ("tx2", "dk2"), ("bg2", "lt2")):
        if real in mp:
            mp.setdefault(alias, mp[real])
    return mp


def _rgb_to_hsl(red, green, blue):
    """0–255 三通道 → (h, s, l)，三者均为 0–1。"""
    r, g, b = red / 255.0, green / 255.0, blue / 255.0
    hi, lo = max(r, g, b), min(r, g, b)
    lum = (hi + lo) / 2.0
    if hi == lo:
        return 0.0, 0.0, lum
    span = hi - lo
    sat = span / (2.0 - hi - lo) if lum > 0.5 else span / (hi + lo)
    if hi == r:
        hue = ((g - b) / span) % 6.0
    elif hi == g:
        hue = (b - r) / span + 2.0
    else:
        hue = (r - g) / span + 4.0
    return hue / 6.0, sat, lum


def _hsl_to_rgb(hue, sat, lum):
    """(h, s, l) 各 0–1 → 0–255 三通道。"""
    if sat == 0:
        chan = int(round(lum * 255))
        return chan, chan, chan
    q = lum * (1.0 + sat) if lum < 0.5 else lum + sat - lum * sat
    p = 2.0 * lum - q

    def _chan(t):
        t %= 1.0
        if t < 1 / 6.0:
            return p + (q - p) * 6.0 * t
        if t < 0.5:
            return q
        if t < 2 / 3.0:
            return p + (q - p) * (2 / 3.0 - t) * 6.0
        return p

    return tuple(int(round(_chan(hue + off) * 255)) for off in (1 / 3.0, 0.0, -1 / 3.0))


def _color_xfrm(clr_node, rgb):
    """按 a:srgbClr / a:schemeClr 下的颜色变换子元素折算出实际 RGB。

    accent1 加 lumMod="60000"（办公软件里的"深色 40%"）与裸 accent1 是肉眼可分的两种蓝，
    只读 val 不折算变换，就会把"7 条颜色各不相同"误判成两条系列共用同一颜色。
    没有可识别的变换时原样返回，避免 HSL 往返的 ±1 误差破坏 E2EFFA 一类的等值比较。
    """
    if clr_node is None or not rgb or len(rgb) != 6:
        return rgb
    try:
        red, green, blue = (int(rgb[0:2], 16), int(rgb[2:4], 16), int(rgb[4:6], 16))
    except ValueError:
        return rgb
    hue, sat, lum, touched = _rgb_to_hsl(red, green, blue) + (False,)
    for child in clr_node:
        name = etree.QName(child).localname
        if name not in ("lumMod", "lumOff", "satMod", "shade", "tint"):
            continue
        try:
            val = int(child.get("val") or 0) / 100000.0
        except ValueError:
            continue
        if name == "lumMod" or name == "shade":          # shade 视为向黑压暗
            lum *= val
        elif name == "lumOff":
            lum += val
        elif name == "tint":                             # 向白混合
            lum = 1.0 - (1.0 - lum) * val
        else:
            sat *= val
        touched = True
    if not touched:
        return rgb
    red, green, blue = _hsl_to_rgb(hue, min(1.0, max(0.0, sat)), min(1.0, max(0.0, lum)))
    return "%02X%02X%02X" % (red, green, blue)


def _solid_color(holder, theme):
    """解析 solidFill/ln 等容器下的颜色，返回 'RRGGBB'；解析不出返回 None。"""
    if holder is None:
        return None
    srgb = holder.find(".//" + _a("srgbClr"))
    if srgb is not None and srgb.get("val"):
        return _color_xfrm(srgb, srgb.get("val").upper())
    sch = holder.find(".//" + _a("schemeClr"))
    if sch is not None and sch.get("val"):
        base = (theme or {}).get(sch.get("val"))
        return "SCHEME:" + sch.get("val") if base is None else _color_xfrm(sch, base)
    return None


def _rect(shape):
    return (shape.left, shape.top, shape.width, shape.height)


def _prst(shape):
    geom = shape._element.find(".//" + _a("prstGeom"))
    return geom.get("prst") if geom is not None else None


def _shape_fill(shape, theme):
    """形状填充色：p:sp 的填充写在 p:spPr 下，图片/图表另有容器，两种命名空间都兼容。"""
    element = shape._element
    for ns in (P_NS, A_NS):
        sp_pr = element.find(".//{%s}spPr" % ns)
        if sp_pr is not None:
            return _solid_color(sp_pr.find(_a("solidFill")), theme)
    return None


def _run_props(run, para, tx_body, theme):
    """按 rPr → 段落 defRPr → txBody/lstStyle 的顺序解析有效字符属性。"""
    out = {"sz": None, "b": None, "latin": None, "ea": None, "color": None}
    sources = []
    if run is not None:
        r_pr = run.find(_a("rPr"))
        if r_pr is not None:
            sources.append(r_pr)
    if para is not None:
        p_pr = para.find(_a("pPr"))
        if p_pr is not None and p_pr.find(_a("defRPr")) is not None:
            sources.append(p_pr.find(_a("defRPr")))
    if tx_body is not None:
        lst = tx_body.find(_a("lstStyle"))
        if lst is not None:
            lvl = lst.find(_a("lvl1pPr"))
            if lvl is not None and lvl.find(_a("defRPr")) is not None:
                sources.append(lvl.find(_a("defRPr")))
    for src in sources:
        if out["sz"] is None and src.get("sz"):
            out["sz"] = int(src.get("sz"))
        if out["b"] is None and src.get("b") is not None:
            out["b"] = src.get("b") in ("1", "true")
        if out["latin"] is None and src.find(_a("latin")) is not None:
            out["latin"] = src.find(_a("latin")).get("typeface")
        if out["ea"] is None and src.find(_a("ea")) is not None:
            out["ea"] = src.find(_a("ea")).get("typeface")
        if out["color"] is None:
            out["color"] = _solid_color(src.find(_a("solidFill")), theme)
    return out


def _tx_body(shape):
    """形状的文字容器：p:sp 用 p:txBody，表格单元格用 a:txBody，两种命名空间都兼容。"""
    element = shape._element
    for ns in (P_NS, A_NS):
        node = element.find(".//{%s}txBody" % ns)
        if node is not None:
            return node
    return None


def _sp_runs(shape, theme):
    """返回文本框内的 [(文本, 有效字符属性)] 列表。"""
    out = []
    tx_body = _tx_body(shape)
    if tx_body is None:
        return out
    for para in tx_body.findall(_a("p")):
        for run in para.findall(_a("r")):
            node = run.find(_a("t"))
            out.append(((node.text or "") if node is not None else "",
                        _run_props(run, para, tx_body, theme)))
    return out


def _sp_text(shape):
    tx_body = _tx_body(shape)
    if tx_body is None:
        return ""
    return "".join((node.text or "") for node in tx_body.iter(_a("t")))


def _sp_algn(shape):
    """首个段落的水平对齐（未写则返回 None，DrawingML 默认左对齐）。"""
    tx_body = _tx_body(shape)
    if tx_body is None:
        return None
    for para in tx_body.findall(_a("p")):
        p_pr = para.find(_a("pPr"))
        return p_pr.get("algn") if p_pr is not None else None
    return None



def _text_shapes(slide):
    """带文字的非表格、非图表形状（标题、说明文字等）。

    slide 传 None（交付文件页数不足，对应页根本不存在）时返回空列表：维度1 只判文件
    能否打开，缺页要让相关加分项自然不得分，而不是抛异常把整份评分打断。
    """
    if slide is None:
        return []
    out = []
    for shape in slide.shapes:
        if shape.has_text_frame and _sp_text(shape).strip():
            out.append(shape)
    return out


def _table_shapes(slide):
    if slide is None:                     # 缺页：交给各加分项不得分，见 _text_shapes
        return []
    return [sh for sh in slide.shapes if getattr(sh, "has_table", False)]


def _tbl_el(shape):
    return shape._element.find(".//" + _a("tbl"))


def _grid(tbl):
    """返回 a:tc 元素的二维数组。"""
    return [tr.findall(_a("tc")) for tr in tbl.findall(_a("tr"))]


def _tc_paras(tc):
    tx_body = tc.find(_a("txBody"))
    return tx_body.findall(_a("p")) if tx_body is not None else []


def _tc_text(tc):
    """单元格文本（保留原始空格/制表符，供 M10 判定；比对时再归一化）。"""
    parts = []
    for para in _tc_paras(tc):
        for child in para:
            if child.tag == _a("r"):
                node = child.find(_a("t"))
                parts.append((node.text or "") if node is not None else "")
            elif child.tag == _a("br"):
                parts.append("\n")
        parts.append("\n")
    return "".join(parts).strip("\n")


def _tc_manual_break(tc):
    """单元格内是否存在手动换行：多于 1 个有文字的 a:p，或出现 a:br。"""
    paras = _tc_paras(tc)
    filled = 0
    for para in paras:
        if para.find(_a("br")) is not None:
            return True
        if any((r.find(_a("t")) is not None and (r.find(_a("t")).text or "").strip())
               for r in para.findall(_a("r"))):
            filled += 1
    return filled > 1


def _tc_pr(tc):
    return tc.find(_a("tcPr"))


def _tc_fill(tc, theme):
    pr = _tc_pr(tc)
    if pr is None:
        return None
    if pr.find(_a("noFill")) is not None:
        return "NOFILL"
    return _solid_color(pr.find(_a("solidFill")), theme)


def _tc_border_visible(tc, theme):
    """单元格是否有肉眼可见的框线（宽度>0 且颜色非纯白）。"""
    pr = _tc_pr(tc)
    if pr is None:
        return False
    for tag in ("lnL", "lnR", "lnT", "lnB"):
        ln = pr.find(_a(tag))
        if ln is None or ln.find(_a("noFill")) is not None:
            continue
        width = int(ln.get("w") or 0)
        color = _solid_color(ln.find(_a("solidFill")), theme)
        if width > 0 and color and color != "FFFFFF":
            return True
    return False


def _tc_border_specs(tc, theme):
    """返回四边框线的 (宽度, 颜色) 集合，用于 M03 比对基准表格是否被改动。"""
    pr = _tc_pr(tc)
    specs = {}
    if pr is None:
        return specs
    for tag in ("lnL", "lnR", "lnT", "lnB"):
        ln = pr.find(_a(tag))
        if ln is None:
            continue
        specs[tag] = (int(ln.get("w") or 0), _solid_color(ln.find(_a("solidFill")), theme))
    return specs


def _tc_runs(tc, theme):
    out = []
    tx_body = tc.find(_a("txBody"))
    if tx_body is None:
        return out
    for para in tx_body.findall(_a("p")):
        for run in para.findall(_a("r")):
            node = run.find(_a("t"))
            out.append(((node.text or "") if node is not None else "",
                        _run_props(run, para, tx_body, theme)))
    return out


def _tbl_style(tbl):
    tbl_pr = tbl.find(_a("tblPr"))
    if tbl_pr is None:
        return None, False, False
    node = tbl_pr.find(_a("tableStyleId"))
    style_id = (node.text or "").strip().upper() if node is not None else None
    first_row = tbl_pr.get("firstRow") in ("1", "true")
    band_row = tbl_pr.get("bandRow") in ("1", "true")
    return style_id, first_row, band_row


def _chart_shapes(slide):
    if slide is None:                     # 缺页：交给各加分项不得分，见 _text_shapes
        return []
    return [sh for sh in slide.shapes if getattr(sh, "has_chart", False)]


def _pt_values(holder):
    """按 idx 升序取 c:pt/c:v 文本列表。"""
    if holder is None:
        return []
    items = []
    for pt in holder.findall(".//" + _c("pt")):
        node = pt.find(_c("v"))
        items.append((int(pt.get("idx") or 0), node.text if node is not None else None))
    items.sort(key=lambda x: x[0])
    return [t[1] for t in items]


def _holder_ref(holder):
    """c:val / c:cat 是否通过 numRef/strRef 引用内嵌工作簿；返回引用式（如 Sheet1!$B$2:$B$10）。

    numLit（把常量直接写死在图表 XML 里）返回 None——这种图表在办公软件里双击虽能打开
    工作簿，但改数不会回写到图表，等同于不可编辑，故单列为判据。
    """
    if holder is None:
        return None
    for tag in ("numRef", "strRef", "multiLvlStrRef"):
        ref = holder.find(_c(tag))
        if ref is None:
            continue
        f_node = ref.find(_c("f"))
        return (f_node.text or "").strip() if f_node is not None else None
    return None


def _ser_name(ser):
    node = ser.find(_c("tx") + "/" + _c("strRef") + "/" + _c("strCache") + "/"
                    + _c("pt") + "/" + _c("v"))
    if node is None:
        node = ser.find(_c("tx") + "/" + _c("v"))
    return (node.text or "") if node is not None else ""


def _axis_txsz(ax):
    if ax is None:
        return None
    node = ax.find(".//" + _a("defRPr"))
    if node is None or not node.get("sz"):
        return None
    return int(node.get("sz"))


def _chart_info(shape, theme):
    """把图表信息抽成一个 dict，避免各判定项重复解析 chart XML。

    shape 传 None（第6页没有原生图表部件，例如折线图是贴的静态图片）时返回空骨架，
    让各图表判定项自然不得分——维度1 只判文件能否打开，不再拦这种交付。
    """
    info = {
        "kinds": [],
        "is_line": False,
        "series": [],
        "cats": [],
        "has_workbook": False,
        "auto_title_deleted": False,
        "legend_pos": None,
        "legend_overlay": None,
        "legend_sz": None,
        "val_min": None,
        "val_max": None,
        "major_unit": None,
        "val_txsz": None,
        "cat_txsz": None,
        "rect": _rect(shape) if shape is not None else (None, None, None, None),
    }
    if shape is None:
        return info
    chart = shape.chart
    root = chart._chartSpace
    plot = root.find(".//" + _c("plotArea"))
    if plot is not None:
        for child in plot:
            name = etree.QName(child).localname
            if name.endswith("Chart"):
                info["kinds"].append(name)
    info["is_line"] = "lineChart" in info["kinds"]
    node = root.find(".//" + _c("autoTitleDeleted"))
    info["auto_title_deleted"] = node is not None and node.get("val") in (None, "1", "true")
    legend = root.find(".//" + _c("legend"))
    if legend is not None:
        pos = legend.find(_c("legendPos"))
        ovl = legend.find(_c("overlay"))
        info["legend_pos"] = pos.get("val") if pos is not None else "r"
        info["legend_overlay"] = (ovl.get("val") in ("1", "true")) if ovl is not None else False
        info["legend_sz"] = _axis_txsz(legend)
    val_ax = root.find(".//" + _c("valAx"))
    if val_ax is not None:
        scaling = val_ax.find(_c("scaling"))
        if scaling is not None:
            for key, tag in (("val_min", "min"), ("val_max", "max")):
                sub = scaling.find(_c(tag))
                if sub is not None and sub.get("val") is not None:
                    info[key] = float(sub.get("val"))
        unit = val_ax.find(_c("majorUnit"))
        if unit is not None and unit.get("val") is not None:
            info["major_unit"] = float(unit.get("val"))
        info["val_txsz"] = _axis_txsz(val_ax)
    cat_ax = root.find(".//" + _c("catAx"))
    if cat_ax is not None:
        info["cat_txsz"] = _axis_txsz(cat_ax)
    for ser in root.findall(".//" + _c("ser")):
        sp_pr = ser.find(_c("spPr"))
        ln = sp_pr.find(_a("ln")) if sp_pr is not None else None
        marker = ser.find(_c("marker"))
        sym = marker.find(_c("symbol")) if marker is not None else None
        msz = marker.find(_c("size")) if marker is not None else None
        d_lbls = ser.find(_c("dLbls"))
        delete = d_lbls.find(_c("delete")) if d_lbls is not None else None
        info["series"].append({
            "name": _ser_name(ser),
            "vals": _pt_values(ser.find(_c("val"))),
            "cats": _pt_values(ser.find(_c("cat"))),
            "val_ref": _holder_ref(ser.find(_c("val"))),
            "cat_ref": _holder_ref(ser.find(_c("cat"))),
            "ln_w": int(ln.get("w")) if (ln is not None and ln.get("w")) else None,
            "color": _solid_color(ln.find(_a("solidFill")), theme) if ln is not None else None,
            "symbol": sym.get("val") if sym is not None else None,
            "msize": int(msz.get("val")) if (msz is not None and msz.get("val")) else None,
            "dlbl_deleted": (d_lbls is None) or (delete is not None
                                                 and delete.get("val") in (None, "1", "true")),
        })
    if info["series"]:
        info["cats"] = info["series"][0]["cats"]
    try:
        for rel in chart.part.rels.values():
            if rel.reltype.endswith("/package") and not rel.is_external:
                info["has_workbook"] = True
                break
    except Exception:
        info["has_workbook"] = False
    return info


def _find_shape_by_rect(slide, off, ext, tol_off, tol_ext, skip=()):
    """按几何位置找装饰图形：命中返回 shape，否则 None。

    与办公软件对齐：装饰图形既可以是自动图形（prstGeom=parallelogram + solidFill），
    也可以是把同一装饰导出后贴上的图片；后者位置/尺寸会有零点几厘米的出入，因此
    这里按容差匹配，颜色只在自动图形上校验。
    """
    if slide is None:                     # 缺页：装饰判定直接不成立，见 _text_shapes
        return None
    for shape in slide.shapes:
        if shape in skip:
            continue
        left, top, width, height = _rect(shape)
        if None in (left, top, width, height):
            continue
        if (_near(left, off[0], tol_off) and _near(top, off[1], tol_off)
                and _near(width, ext[0], tol_ext) and _near(height, ext[1], tol_ext)):
            return shape
    return None


def _deco_ok(slide, theme, tol_off, tol_ext, need_fill=True):
    """检查左上浅蓝 + 右下浅青两个平行四边形装饰是否都在位。"""
    top_left = _find_shape_by_rect(slide, DECO_TL_OFF, DECO_TL_EXT, tol_off, tol_ext)
    bottom_right = _find_shape_by_rect(slide, DECO_BR_OFF, DECO_BR_EXT, tol_off, tol_ext)
    if top_left is None or bottom_right is None:
        return False
    if not need_fill:
        return True
    for shape, want in ((top_left, DECO_TL_FILL), (bottom_right, DECO_BR_FILL)):
        if etree.QName(shape._element).localname == "pic":
            continue                      # 图片形式的装饰：不校验 prstGeom 与填充色
        if _prst(shape) != "parallelogram" or _shape_fill(shape, theme) != want:
            return False
    return True


# ------------------------------------------------------------------ 主入口 ----


def evaluate(dir_path):
    """入参为目录路径；在该目录下定位唯一的 .pptx 交付文件并评分。"""
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
        names = os.listdir(dir_path) if os.path.isdir(dir_path) else []
        candidates = sorted(
            f for f in names
            if f.lower().endswith(".pptx") and not f.startswith("~$")
        )
        if not candidates:
            # 目录里没有交付文件：根本没有进入维度1 判定，dim1_pass 留空字符串以区别于
            # "判过且不通过"（False），dim1_reason 同样留空，满分也记 0。
            result["status"] = "error"
            result["error"] = "目录中未找到 .pptx 文件：%s" % dir_path
            result["dim1_pass"] = ""
            result["dim1_reason"] = ""
            result["dim2_items"] = []
            result["total_score"] = 0
            result["max_score"] = 0
            return result
        file_name = candidates[0]
        result["file_name"] = file_name
        _evaluate_impl(os.path.join(dir_path, file_name), result)
    except Exception as exc:                                  # noqa: BLE001
        # 其余异常属于"维度1 判定失败"：dim1_pass 记 False，并把错误摘要写进 dim1_reason。
        summary = "%s: %s" % (type(exc).__name__, exc)
        result["status"] = "error"
        result["error"] = summary
        result["dim1_pass"] = False
        result["dim1_reason"] = "评测过程异常，无法确认文件可用与可修改性：%s" % summary
        result["dim2_items"] = []
        result["total_score"] = 0
    return result


def _family(rgb):
    """把颜色粗分成 blue / orange / purple / other，用于判断色板是否沿用蓝橙紫。"""
    if not rgb or len(rgb) != 6:
        return "other"
    try:
        red, green, blue = (int(rgb[0:2], 16), int(rgb[2:4], 16), int(rgb[4:6], 16))
    except ValueError:
        return "other"
    if blue >= red + 20 and red >= green + 20:
        return "purple"
    if blue >= red + 24 and blue >= green:
        return "blue"
    if red >= green + 24 and green >= blue + 12:
        return "orange"
    return "other"


def _evaluate_impl(file_path, result):
    # ============================== 维度1：可用与可修改性 ==============================
    # 只判一件事：交付文件是否为可正常打开的 .pptx。页数是否齐、新增表格/折线图是否原生
    # 可编辑、版面是否越界，全部下沉到维度2 逐项记账（页数与页序归 P01，原生表格归
    # P02/P03，原生 lineChart 与内嵌工作簿归 P13，画幅归 P08/P19），不再作为归零门槛。
    try:
        prs = Presentation(file_path)
    except Exception as exc:                                  # noqa: BLE001
        result["dim1_pass"] = False
        result["dim1_reason"] = "文件损坏或无法作为 PowerPoint 文档打开：%s" % exc
        result["total_score"] = 0
        return

    theme = _theme_colors(prs)
    slides = list(prs.slides)
    result["dim1_reason"] = "文件可正常打开为 PowerPoint 文档（.pptx 未损坏），共 %d 页" % len(slides)

    # ---- 定位新增页对象；缺失时留 None/空，由维度2 各项自然不得分 ----
    slide5 = slides[4] if len(slides) >= 5 else None
    slide6 = slides[5] if len(slides) >= 6 else None
    tbl_shapes5 = _table_shapes(slide5)
    cht_shapes6 = _chart_shapes(slide6)
    tbl_shape = tbl_shapes5[0] if tbl_shapes5 else None
    cht_shape = cht_shapes6[0] if cht_shapes6 else None
    cht = _chart_info(cht_shape, theme)

    # ================================= 维度2：完成度 =================================
    results = []

    def _plus(delta, text, hit):
        results.append("+%d：%s" % (delta if hit else 0, text))

    def _minus(delta, text, hit):
        """hit=True 表示"破坏性修改被检出"，按该项最大扣分记账。"""
        results.append("-%d：%s" % (abs(delta) if hit else 0, text))

    # ---- 预取新增页对象；第5页没有原生表格时全部留空，相关项自然不得分 ----
    tbl = _tbl_el(tbl_shape) if tbl_shape is not None else None
    grid = _grid(tbl) if tbl is not None else []
    n_rows = len(grid)
    n_cols = min((len(r) for r in grid), default=0)
    hdr_cells = grid[0] if grid else []
    data_rows = grid[1:]
    grid_el = tbl.find(_a("tblGrid")) if tbl is not None else None
    col_w = ([int(gc.get("w") or 0) for gc in grid_el.findall(_a("gridCol"))]
             if grid_el is not None else [])
    style_id, first_row, band_row = _tbl_style(tbl) if tbl is not None else (None, False, False)
    style_def = BUILTIN_BANDED_STYLES.get(style_id)
    all_cells = [tc for row in grid for tc in row]

    def _page_texts(slide):
        return sorted(_text_shapes(slide), key=lambda sh: (sh.top if sh.top is not None else 0))

    def _title_of(slide):
        texts = _page_texts(slide)
        return texts[0] if texts else None

    def _note_of(slide):
        texts = _page_texts(slide)
        return texts[-1] if len(texts) >= 2 else None

    # ---- 预取第1至第4页基准对象 ----
    # 页数不足时对应位置留 None：维度1 只判文件能否打开，缺页由 P01 与各基准减分项记账。
    base = []
    for idx in range(4):
        slide = slides[idx] if idx < len(slides) else None
        tables = _table_shapes(slide)
        charts = _chart_shapes(slide)
        base.append({
            "slide": slide,
            "table": tables[0] if tables else None,
            "chart_shape": charts[0] if charts else None,
            "chart": _chart_info(charts[0], theme) if charts else None,
            "title": _title_of(slide),
            "note": _note_of(slide),
        })

    base_titles_ok = all(
        item["title"] is not None and _norm(_sp_text(item["title"])) == _norm(BASE_TITLES[i])
        for i, item in enumerate(base))
    base_layout_ok = (base[0]["table"] is not None and base[1]["chart"] is not None
                      and base[2]["table"] is not None and base[3]["chart"] is not None)

    # [P01 +3] 新增页的位置与页序：必须是 6 页，第5页表格、第6页折线图，前4页版面构成不变。
    # 去重：前4页的标题文案由 M06 单独记账，此处不再复查 base_titles_ok，避免改一个标题
    # 同时打掉 P01、M01、M06 三项。
    p01 = (len(slides) == 6
           and bool(_table_shapes(slide5)) and not _chart_shapes(slide5)
           and bool(_chart_shapes(slide6)) and not _table_shapes(slide6)
           and base_layout_ok)
    _plus(3, "在第4页之后依次新增第5、6页并置于文稿末尾，表格页为第5页、折线图页为第6页，"
             "第1至第4页的版面构成（表格/折线图/表格/折线图）保持不变", p01)

    # [P02 +5] 表格规模与表头：10 行 × 8 列，8 个列名与列序完整。
    # 与办公软件对齐：列名比对前先归一化（去空格/制表符、半角括号转全角、连字符统一），
    # 因为这些差异不影响阅读者认定"列名正确"，空格/制表符本身由 M10 单独扣分。
    hdr_texts = [_norm(_tc_text(tc)) for tc in hdr_cells]
    p02 = (n_rows == 10 and n_cols == 8 and hdr_texts == [_norm(h) for h in NEW_HEADERS])
    _plus(5, "第5页为 10 行 × 8 列表格，表头8列自左至右为"
             "年份/智慧门店响应率（%）/线上订单占比（%）/自助结算覆盖率（%）/绿色到店客流占比（%）/"
             "低碳配送完成率（%）/末端接驳便利度（%）/运营协同指数（0–100），列名与列序完整无错序", p02)

    # [P03 +5] 首列 9 个连续年份 + 9×7=63 个数值逐格与数据表一致。
    years = [_num(_tc_text(row[0])) for row in data_rows if row]
    p03 = years == [float(y) for y in NEW_YEARS] and n_cols >= 8 and len(data_rows) >= 9
    if p03:
        for ri in range(9):
            for ci in range(7):
                if not _same(_num(_tc_text(data_rows[ri][ci + 1])), NEW_ROWS[ri][ci]):
                    p03 = False
                    break
            if not p03:
                break
    _plus(5, "表格首列自上而下为 2016 至 2024 共9个连续年份，9行×7列共63个数值与数据表逐格一致"
             "（2016行 31.8/18.9/24.6/27.3/22.6/33.4/29.7，2020行 53.4/41.7/49.6/42.8/44.3/50.6/52.9，"
             "2024行 78.9/68.1/76.7/67.3/72.6/74.9/78.2），无错值、漏值、串行、串列", p03)

    # [P04 +5] 套用与第1、3页相同的表格样式（含首行与镶边行），并因此使表头呈深色底纹。
    # 合并说明：本主题 accent1 = 4F81BD 已在深色白名单内，只要套上该样式并开启 firstRow，
    # 办公软件必然渲染出深蓝表头——"套样式"与"表头深色底纹"无法独立判别，故并为一项。
    # 细则分值只用 1/3/5 三档，合并项按最高档 5 分记（原两项各 3 分，合并后总分 63）。
    # 手工刷底纹的交付走 explicit_hdr 通路，仍能单独判定。
    p04_style = (style_id == TABLE_STYLE_ID and first_row and band_row)
    hdr_fills = [_tc_fill(tc, theme) for tc in hdr_cells]
    explicit_hdr = sorted({f for f in hdr_fills if f and f != "NOFILL"})
    if explicit_hdr:
        p04_dark = len(explicit_hdr) == 1 and explicit_hdr[0] in HEADER_DARK_OK
    else:
        eff_hdr = theme.get(style_def["header_fill"]) if (style_def and first_row) else None
        p04_dark = eff_hdr in HEADER_DARK_OK
    p04 = p04_style and p04_dark
    _plus(5, "表格套用与第1页、第3页表格相同的表格样式（tableStyleId "
             "{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}）并开启首行样式与镶边行，数据区呈隔行斑马纹；"
             "表头行由此呈深色底纹（样式提供的 Accent1 蓝 4F81BD，或显式设为与第1页表头蓝 2F66A4、"
             "第3页表头青 269B91 同类的深色），未改为浅灰、浅黄等与数据区无法区分的浅色", p04)

    # [P05 +3] 表头白色加粗11磅；数据单元格微软雅黑且不小于9.5磅。
    # 与办公软件对齐：白色+加粗可以由内置表格样式的首行样式提供（slide XML 里没有显式
    # rPr），这种写法在办公软件中同样渲染为白色加粗，故按样式折算后再判定。
    def _eff_hdr_prop(props, key):
        if props[key] is not None:
            return props[key]
        if style_def and first_row:
            return style_def.get("header_color" if key == "color" else "header_bold")
        return None

    hdr_props = [pr for tc in hdr_cells for _, pr in _tc_runs(tc, theme)]
    data_props = [pr for row in data_rows for tc in row for _, pr in _tc_runs(tc, theme)]
    p05 = (bool(hdr_props) and bool(data_props)
           and all(_eff_hdr_prop(pr, "color") == "FFFFFF" and _eff_hdr_prop(pr, "b") is True
                   and pr["sz"] == 1100 for pr in hdr_props)
           and all(pr["sz"] is not None and pr["sz"] >= 950
                   and "微软雅黑" in ((pr["latin"] or "") + (pr["ea"] or ""))
                   for pr in data_props))
    _plus(3, "表头文字为白色、加粗、字号11磅，数据单元格文字为微软雅黑、字号9.5磅及以上，"
             "未被压小到9磅以下", p05)

    # [P06 +5] 行列分隔清晰可辨。
    # 两条通路任一成立即可：①单元格自带可见框线（宽度>0 且颜色非纯白，如 D6E0E8）；
    # ②未写显式底纹而由内置镶边样式提供浅色底纹带 + 白色框线。若单元格被手工刷成
    # 白/近白底又未设任何框线，则列与列之间没有可见分隔，判不通过。
    # 与 P04 的区别：套上样式后再把每个单元格手工刷成白底（显式底纹压掉镶边带）仍会判不通过，
    # 故本项不被 P04 蕴含，保留为独立项。
    n_border = sum(1 for tc in all_cells if _tc_border_visible(tc, theme))
    n_expl_fill = sum(1 for tc in all_cells if _tc_fill(tc, theme) not in (None, "NOFILL"))
    p06 = bool(all_cells) and (
        n_border >= int(len(all_cells) * 0.9)
        or (bool(style_def) and band_row and n_expl_fill <= len(all_cells) * 0.1))
    _plus(5, "表格具备清晰可辨的行列分隔（单元格设 D6E0E8 一类浅色框线，或由表格样式的"
             "浅色底纹带配白色框线形成可见分隔），不出现整片白底且行列之间无任何可见分隔的情形", p06)

    # [P07 +3] 全部单元格水平居中 + 垂直居中。
    anchors = [(_tc_pr(tc).get("anchor") if _tc_pr(tc) is not None else None) for tc in all_cells]
    algns = []
    for tc in all_cells:
        for para in _tc_paras(tc):
            p_pr = para.find(_a("pPr"))
            algns.append(p_pr.get("algn") if p_pr is not None else None)
    p07 = bool(all_cells) and all(x == "ctr" for x in anchors) and all(x == "ctr" for x in algns)
    _plus(3, "所有单元格文字水平居中且垂直居中（单元格属性 anchor=\"ctr\"），数字未贴在单元格上沿", p07)

    # [P08 +3] 表格宽度≈10972800EMU、左右留白均衡、指标列等宽且表头不折行。
    # 与办公软件对齐：本机无渲染能力，无法量出表头实际排版宽度，故"列宽足以容纳表头文字"
    # 以"表头单元格内没有手动换行（多段落/a:br）"作为等价判据。
    # 去重：表头格的手动换行只在本项记账，M10 只扫数据区，避免同一个回车被扣两次。
    tbl_left, _tbl_top, tbl_width, _tbl_height = (
        _rect(tbl_shape) if tbl_shape is not None else (None, None, None, None))
    tol_w = int(EMU_PER_INCH * 0.1)
    margin_right = prs.slide_width - (tbl_left or 0) - (tbl_width or 0)
    metric_w = col_w[1:8] if len(col_w) >= 8 else []
    p08 = (_near(tbl_width, BASE_TBL_EXT[0], tol_w)
           and abs((tbl_left or 0) - margin_right) <= tol_w
           and len(metric_w) == 7 and (max(metric_w) - min(metric_w)) <= 9144
           and not any(_tc_manual_break(tc) for tc in hdr_cells))
    _plus(3, "表格宽度约 10972800 EMU（约11.43英寸）、左右留白之差不超过0.1英寸，且除年份列外的"
             "各指标列列宽彼此相等、列宽足以容纳表头文字而无需手动折行或插入换行符", p08)

    # [P09 +1] 第5、6页左上角页面标题文本框：微软雅黑24磅加粗233142，且题名点题。
    def _fmt_title_ok(shape):
        if shape is None:
            return False
        runs = _sp_runs(shape, theme)
        return bool(runs) and all(
            pr["sz"] == 2400 and pr["b"] is True and pr["color"] == "233142"
            and "微软雅黑" in ((pr["latin"] or "") + (pr["ea"] or ""))
            for _txt, pr in runs)

    def _in_top_left(shape):
        return (shape is not None and (shape.left or 0) < prs.slide_width * 0.45
                and (shape.top or 0) < prs.slide_height * 0.3)

    def _in_bottom_left(shape):
        return (shape is not None and (shape.left or 0) < prs.slide_width * 0.45
                and (shape.top or 0) > prs.slide_height * 0.55)

    title5, title6 = _title_of(slide5), _title_of(slide6)
    txt5 = _norm(_sp_text(title5)) if title5 is not None else ""
    txt6 = _norm(_sp_text(title6)) if title6 is not None else ""
    p09 = (_fmt_title_ok(title5) and _fmt_title_ok(title6)
           and _in_top_left(title5) and _in_top_left(title6)
           and any(k in txt5 for k in ("指标", "数据", "指数")) and "趋势" not in txt5
           and "趋势" in txt6)
    _plus(1, "第5页与第6页左上角各设页面标题文本框，文字均为微软雅黑、24磅、加粗、颜色 233142，"
             "题名分别点明本页7项指标主题与该组指标的趋势", p09)

    # [P10 +1] 第5页左下角量纲说明：7–8磅，同时说明百分比与综合指数（0–100）两种单位。
    note5 = _note_of(slide5)
    note5_txt = _norm(_sp_text(note5)) if note5 is not None else ""
    note5_props = [pr for _txt, pr in _sp_runs(note5, theme)] if note5 is not None else []
    p10 = (_in_bottom_left(note5) and bool(note5_props)
           and all(pr["sz"] is not None and 700 <= pr["sz"] <= 800 for pr in note5_props)
           and "单位" in note5_txt and "0-100" in note5_txt
           and ("%" in note5_txt or "百分比" in note5_txt))
    _plus(1, "第5页左下角设量纲说明文本框，文字7–8磅，说明百分比与综合指数（0–100）的单位", p10)

    # [P11 +1] 第6页左下角极值说明：指标名与数值都要正确。
    note6 = _note_of(slide6)
    note6_txt = _norm(_sp_text(note6)) if note6 is not None else ""
    p11 = (_in_bottom_left(note6) and "最高点" in note6_txt
           and "智慧门店响应率" in note6_txt and "2024" in note6_txt and "78.9" in note6_txt)
    _plus(1, "第6页左下角设极值说明文本框，内容为\"最高点：智慧门店响应率在2024年达到 78.9\"，"
             "指标名与数值均正确无误", p11)

    # [P12 +3] 新增两页沿用第1至第4页的两个平行四边形装饰图形。
    # 与办公软件对齐：装饰件若是把原装饰导出为图片再贴回，位置/尺寸会有零点几厘米出入且
    # 无 prstGeom/solidFill，视觉上等效，故按容差匹配、仅在自动图形上校验填充色。
    tol_off = int(EMU_PER_INCH * 0.125)
    tol_ext = int(EMU_PER_INCH * 0.17)
    p12 = (_deco_ok(slide5, theme, tol_off, tol_ext) and _deco_ok(slide6, theme, tol_off, tol_ext))
    _plus(3, "第5页与第6页左上角均有浅蓝 E2EFFA、尺寸约 2834640×502920 EMU 的平行四边形装饰图形，"
             "右下角均有浅青 E1F4F1、尺寸约 2926080×347472 EMU 的平行四边形装饰图形，"
             "位置与第1至第4页相同", p12)

    # [P13 +5] 折线图是"带数据标记、且数据真能改"的原生图表。
    # 去重：维度1 已经把"有 chart 部件 / 是 lineChart / 有 ppt/embeddings 内嵌工作簿"作为门槛
    # （不满足则总分归零），本项不再复述，改判维度1 未覆盖的三件事：
    #   ①图表部件内只有 lineChart 一种绘图类型，未与柱形/面积/散点/雷达混排；
    #   ②各系列都设了数据标记（marker 的 symbol 非 none）；
    #   ③各系列的数值与分类都通过 numRef/strRef 引用工作簿单元格区域（如 Sheet1!$B$2:$B$10），
    #     而不是 numLit 把常量写死在图表 XML 里——后者在办公软件里改工作簿不会回写到图表。
    p13 = (cht["kinds"] == ["lineChart"] and bool(cht["series"])
           and all(s["symbol"] not in (None, "none") for s in cht["series"])
           and all(s["val_ref"] and s["cat_ref"] for s in cht["series"]))
    _plus(5, "第6页图表部件内只有 lineChart 一种绘图类型（未与柱形图、面积图、散点图或雷达图混排），"
             "各系列均设数据标记（marker 的 symbol 非 none），且各系列的数值与分类均通过 "
             "numRef/strRef 引用内嵌工作簿的单元格区域（如 Sheet1!$B$2:$B$10）而非 numLit 硬编码，"
             "双击图表回到工作簿改数即可更新图形", p13)

    # [P14 +5] 7 条系列与表格 7 项指标一一对应，63 个数据点逐点一致。
    ser_names = [_norm(s["name"]) for s in cht["series"]]
    p14 = (len(cht["series"]) == 7 and ser_names == [_norm(h) for h in NEW_HEADERS[1:]])
    if p14:
        for si, ser in enumerate(cht["series"]):
            vals = [_num(v) for v in ser["vals"]]
            if len(vals) != 9 or any(not _same(vals[ri], NEW_ROWS[ri][si]) for ri in range(9)):
                p14 = False
                break
    _plus(5, "图表含7条数据系列，系列名与表格7项指标一一对应（不漏系列、不多系列、未把年份列"
             "当作系列），且7×9共63个数据点与表格数值逐点一致"
             "（如智慧门店响应率为 31.8/34.9/42.6/45.7/53.4/58.6/66.7/71.6/78.9）", p14)

    # [P15 +3] 分类轴9年升序、数值轴 0–100 且主刻度等距、数据点不贴顶贴底。
    # 与办公软件对齐：majorUnit 未写时为"自动"，办公软件按等距刻度渲染，视为满足等距。
    cats = [_num(v) for v in cht["cats"]]
    data_all = [_num(v) for ser in cht["series"] for v in ser["vals"] if _num(v) is not None]
    p15 = (cats == [float(y) for y in NEW_YEARS]
           and _same(cht["val_min"], 0.0) and _same(cht["val_max"], 100.0)
           and (cht["major_unit"] is None or cht["major_unit"] > 0)
           and bool(data_all) and min(data_all) >= 2.0 and max(data_all) <= 98.0)
    _plus(3, "分类轴为 2016 至 2024 共9个年份并按升序排列，数值轴最小值0、最大值100且主刻度等距，"
             "全部数据点落在可视范围内且不贴顶贴底", p15)

    # [P16 +3] 图表标题关闭；图例在绘图区下方且不覆盖绘图区。
    p16 = (cht["auto_title_deleted"] and cht["legend_pos"] == "b"
           and cht["legend_overlay"] is False)
    _plus(3, "图表标题关闭（autoTitleDeleted=\"1\"），图例位于绘图区下方（legendPos=\"b\"）"
             "且不覆盖绘图区（overlay=\"0\"）", p16)

    # [P17 +3] 折线线宽 ≥ 2.25 磅；坐标轴刻度文字 9 磅。
    p17 = (bool(cht["series"])
           and all(s["ln_w"] is not None and s["ln_w"] >= 28575 for s in cht["series"])
           and (cht["val_txsz"] or 0) >= 900 and (cht["cat_txsz"] or 0) >= 900)
    _plus(3, "各系列折线线宽不小于 28575 EMU（2.25磅）且未被细化到2磅以下，"
             "坐标轴刻度文字字号9磅且未被压小到8磅及以下", p17)

    # [P18 +3] 7 条折线颜色两两不同，且整体沿用蓝/橙/紫为主的色板。
    # 注意：系列颜色写成 schemeClr 时按主题配色折算实际 RGB，并且必须把 lumMod/lumOff 一类
    # 颜色变换一起折算（见 _color_xfrm）——accent1 与 accent1+lumMod 60%（"深色 40%"）在
    # 办公软件里是两种可分的蓝，只读 val 会把它们误判成共用同一颜色；反之两条系列裸引用同一个
    # accent 才是真正的颜色重复。
    ser_colors = [s["color"] for s in cht["series"]]
    families = {_family(c) for c in ser_colors}
    p18 = (len(ser_colors) == 7 and all(ser_colors) and len(set(ser_colors)) == 7
           and {"blue", "orange", "purple"} <= families)
    _plus(3, "7条折线颜色两两不同、可清晰区分，无两条系列共用同一颜色，"
             "且整体沿用蓝/橙/紫为主的色板", p18)

    # [P19 +3] 图表画幅与第2、4页图表相当。
    p19 = ((cht["rect"][2] or 0) >= 10000000 and (cht["rect"][3] or 0) >= 4700000)
    _plus(5, "图表区宽度不小于 10000000 EMU、高度不小于 4700000 EMU，占据页面主体画幅，"
             "与第2页、第4页图表的画幅规模相当", p19)

    # [M01 -3] 页面尺寸（整份文稿的版式基准）。
    # 去重：总页数不足 6 页已在维度1 归零，页数多于 6 页或新增页未置于末尾由 P01 记账，
    # 前4页标题文案由 M06 记账，故本项只保留页面尺寸这一项独立信息。
    m01 = not (prs.slide_width == SLIDE_W and prs.slide_height == SLIDE_H)
    _minus(3, "页面尺寸不满足 12192000×6858000 EMU（33.87cm×19.05cm，16:9），"
              "整份文稿的版式基准被改动", m01)

    # [M02 -5] 第1、3页表格的行列数、表头与九年数值。
    def _base_tbl_data_ok(shape, headers, cols):
        if shape is None:
            return False
        cells = _grid(_tbl_el(shape))
        if len(cells) != 10 or min((len(r) for r in cells), default=0) != 4:
            return False
        if [_norm(_tc_text(tc)) for tc in cells[0]] != [_norm(h) for h in headers]:
            return False
        for ri in range(9):
            if not _same(_num(_tc_text(cells[ri + 1][0])), float(NEW_YEARS[ri])):
                return False
            for ci in range(3):
                if not _same(_num(_tc_text(cells[ri + 1][ci + 1])), cols[ci][ri]):
                    return False
        return True

    m02 = not (_base_tbl_data_ok(base[0]["table"], BASE_T1_HEADERS, BASE_T1_COLS)
               and _base_tbl_data_ok(base[2]["table"], BASE_T3_HEADERS, BASE_T3_COLS))
    _minus(5, "第1页表格不满足10行4列、表头为年份/北岸片区/云岭片区/星湾片区、数据为2016年 "
              "48.2/35.7/22.8 至2024年 82.4/69.6/56.3 的九年三片区数值；或第3页表格不满足10行4列、"
              "表头为年份/区域均值/区域均衡度/运营转化指数、数据为2016年 35.6/74.6/48.1 至2024年 "
              "69.4/73.9/70.8 的九年数值", m02)

    # [M03 -1] 第1、3页表格的位置尺寸、框线、底纹、字号。
    def _base_tbl_style_ok(shape, header_fill, check_sz):
        if shape is None:
            return False
        if _rect(shape) != (BASE_TBL_OFF[0], BASE_TBL_OFF[1], BASE_TBL_EXT[0], BASE_TBL_EXT[1]):
            return False
        cells = _grid(_tbl_el(shape))
        for row in cells:
            for tc in row:
                specs = _tc_border_specs(tc, theme)
                if len(specs) != 4 or any(sp != (6350, "D6E0E8") for sp in specs.values()):
                    return False
        if any(_tc_fill(tc, theme) != header_fill for tc in cells[0]):
            return False
        band = []
        for ri in range(1, len(cells)):
            if _tc_fill(cells[ri][0], theme) != "EBF1F7":
                return False
            row_fills = {_tc_fill(cells[ri][ci], theme) for ci in range(1, len(cells[ri]))}
            if len(row_fills) != 1:
                return False
            band.append(row_fills.pop())
        if any(f not in ("FFFFFF", "F2F7F9") for f in band):
            return False
        if any(band[i] == band[i + 1] for i in range(len(band) - 1)):
            return False
        if check_sz:
            hdr = [pr for tc in cells[0] for _txt, pr in _tc_runs(tc, theme)]
            dat = [pr for ri in range(1, len(cells)) for tc in cells[ri]
                   for _txt, pr in _tc_runs(tc, theme)]
            if not hdr or any(pr["sz"] != 1100 for pr in hdr):
                return False
            if not dat or any(pr["sz"] != 950 for pr in dat):
                return False
        return True

    m03 = not (_base_tbl_style_ok(base[0]["table"], "2F66A4", True)
               and _base_tbl_style_ok(base[2]["table"], "269B91", False))
    _minus(1, "第1页、第3页表格不满足位置 (594360, 1078992) EMU、尺寸 10972800×4736592 EMU，"
              "或其单元格框线（6350、D6E0E8）、底纹（第1页表头 2F66A4、第3页表头 269B91、"
              "年份列 EBF1F7、数值区 FFFFFF 与 F2F7F9 交替）、第1页表格字号（表头11磅、数据9.5磅）"
              "被改动", m03)

    # [M04 -3] 第2、4页折线图的系列构成与数值轴刻度。
    def _base_cht_data_ok(info, names, vmin, vmax, unit):
        if info is None or len(info["series"]) != 3:
            return False
        if [_norm(s["name"]) for s in info["series"]] != [_norm(n) for n in names]:
            return False
        if not (_same(info["val_min"], vmin) and _same(info["val_max"], vmax)):
            return False
        if unit is not None and not _same(info["major_unit"], unit):
            return False
        return True

    m04 = not (_base_cht_data_ok(base[1]["chart"], BASE_T1_HEADERS[1:], 20.0, 90.0, None)
               and _base_cht_data_ok(base[3]["chart"], BASE_T3_HEADERS[1:], 0.0, 100.0, 20.0))
    _minus(3, "第2页折线图不满足3条系列（北岸片区、云岭片区、星湾片区）与数值轴最小值20最大值90；"
              "或第4页折线图不满足3条系列（区域均值、区域均衡度、运营转化指数）与数值轴最小值0"
              "最大值100且主刻度20", m04)

    # [M05 -1] 第2、4页折线图的位置尺寸、系列颜色、菱形标记、线宽、图例与数据标签。
    def _base_cht_style_ok(shape, info):
        if shape is None or info is None:
            return False
        if _rect(shape) != (BASE_CHT_OFF[0], BASE_CHT_OFF[1], BASE_CHT_EXT[0], BASE_CHT_EXT[1]):
            return False
        if [s["color"] for s in info["series"]] != BASE_SER_COLORS:
            return False
        if any(s["symbol"] != "diamond" for s in info["series"]):
            return False
        if any(s["ln_w"] != 28575 for s in info["series"]):
            return False
        if not all(s["dlbl_deleted"] for s in info["series"]):
            return False
        return (info["legend_pos"] == "b" and info["legend_overlay"] is False
                and info["legend_sz"] == 1800)

    m05 = not (_base_cht_style_ok(base[1]["chart_shape"], base[1]["chart"])
               and _base_cht_style_ok(base[3]["chart_shape"], base[3]["chart"]))
    _minus(1, "第2页、第4页折线图不满足位置 (676656, 1033271) EMU、尺寸 10744200×4800600 EMU，"
              "或其系列颜色（2F66A4/D9962A/7453AA）、菱形数据标记、2.25磅线宽、18磅图例文字、"
              "图例位于绘图区下方被改动，或各系列数据标签被额外开启造成数值堆叠遮挡折线", m05)

    # [M06 -1] 第1至第4页标题文字（第1、2页另含显式字体属性）。
    m06 = not base_titles_ok
    for idx in (0, 1):
        runs = _sp_runs(base[idx]["title"], theme) if base[idx]["title"] is not None else []
        if not runs or not all(pr["sz"] == 2400 and pr["b"] is True and pr["color"] == "233142"
                               and "微软雅黑" in ((pr["latin"] or "") + (pr["ea"] or ""))
                               for _txt, pr in runs):
            m06 = True
    _minus(1, "第1页标题门店数字化运营指数、第2页标题门店数字化运营指数趋势、第3页标题"
              "门店数字化运营关联指标、第4页标题门店数字化运营关联指标趋势任一被改写，"
              "或第1页、第2页标题不满足微软雅黑、24磅、加粗、颜色 233142", m06)

    # [M07 -1] 第1至第4页左下角说明文字（文案、字号，第1、2页另含颜色 607080）。
    m07 = False
    for idx in range(4):
        note = base[idx]["note"]
        if note is None or _norm(_sp_text(note)) != _norm(BASE_NOTES[idx]):
            m07 = True
            continue
        props = [pr for _txt, pr in _sp_runs(note, theme)]
        if not props or any(pr["sz"] != BASE_NOTE_SZ[idx] for pr in props):
            m07 = True
        if idx in (0, 1) and any(pr["color"] != "607080" for pr in props):
            m07 = True
    _minus(1, "第1至第4页左下角的说明文字单位：综合指数（0–100）；／最高点：北岸片区在2024年达到 "
              "82.4／单位：综合指数（0–100）；关联数据由第一个表格计算。／最高点：区域均衡度在2017年"
              "达到 94.9 任一被改写，或不满足对应字号（依次为8磅、9磅、7磅、7磅），"
              "或第1页、第2页说明文字颜色不为 607080", m07)

    # [M08 -1] 第1至第4页的平行四边形装饰图形（位置、尺寸、填充色均须原样）。
    m08 = not all(_deco_ok(base[idx]["slide"], theme, 0, 0) for idx in range(4))
    _minus(1, "第1至第4页左上角浅蓝 E2EFFA、右下角浅青 E1F4F1 的平行四边形装饰图形"
              "被删除、改色、移位或改尺寸", m08)

    # [M09 -3] 新增表格的数值单元格追加了单位符号。
    # 去重：四舍五入取整、单位换算、按数值大小重排年份都会让 P03（表格63值）与 P14（图表63点）
    # 直接失分，若这里再扣一次，同一处缺陷会被记账两遍。而"31.8%"这种写法 _num 仍能取出
    # 31.8、P03/P14 照样通过，是本项唯一独立于加分项的形态，故 M09 只保留它。
    raw_vals = []
    if n_cols >= 8 and len(data_rows) >= 9:
        raw_vals = [_tc_text(data_rows[ri][ci + 1]) for ri in range(9) for ci in range(7)]
    m09 = any("%" in t or "％" in t for t in raw_vals)
    _minus(3, "新增表格的数值单元格追加了 %／％ 等单位符号（数值应为纯数字，单位只出现在表头与"
              "说明文字中），使单元格内容偏离数据表的原始数值", m09)

    # [M10 -1] 新增表格的数据单元格混入制表符、连续空格或手动换行符。
    # 去重：只扫数据区（grid[1:]），表头格的手动换行由 P08 以"列宽足以容纳表头文字"记账，
    # 否则在表头插一个回车会同时打掉 P08（-3）与触发本项（-1），同一处缺陷记两次账。
    m10 = False
    for tc in [tc for row in data_rows for tc in row]:
        text = _tc_text(tc)
        if "\t" in text or "  " in text or "　　" in text or _tc_manual_break(tc):
            m10 = True
            break
    _minus(1, "新增表格的数据单元格内混入制表符、连续空格或手动换行符，造成同列数字忽左忽右、"
              "行高被撑开而各行不齐", m10)

    # =============================== 汇总（与 008 一致） ===============================
    if len(results) != len(_MAX_DELTAS_ORDER):
        raise RuntimeError("dim2 项数不一致：results=%d vs order=%d"
                           % (len(results), len(_MAX_DELTAS_ORDER)))
    dim2_items = []
    total_score = 0
    for raw, max_delta in zip(results, _MAX_DELTAS_ORDER):
        head, _sep, rule = raw.partition("：")
        try:
            delta = int(head.replace("+", "").strip())
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
    result["total_score"] = max(0, total_score)


if __name__ == "__main__":
    # 本地调试入口：默认评测脚本所在目录，也可通过命令行参数指定目录。
    _dir = sys.argv[1] if len(sys.argv) >= 2 else os.path.dirname(os.path.abspath(__file__))
    print(json.dumps(evaluate(_dir), ensure_ascii=False, indent=2))
