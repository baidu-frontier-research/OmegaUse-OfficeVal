#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# pyright: reportUnknownMemberType=false, reportUnknownVariableType=false, reportUnknownArgumentType=false, reportUnknownParameterType=false, reportMissingParameterType=false, reportPrivateUsage=false, reportUnknownLambdaType=false, reportGeneralTypeIssues=false

import os
import re
import sys
import json
import zipfile
import statistics
import math
import xml.etree.ElementTree as ET
from docx import Document
from docx.oxml.ns import qn

# 说明：按接口统一约定，本模块不修改全局 sys.stdout，也不 print 主结果。
# 主结果统一由 evaluate(dir_path) -> dict 返回；入参为“脚本所在目录的路径”，
# 由脚本自己在该目录内定位并打开被评估的 .docx。本文件仅在 __main__ 下作为
# 本地调试入口，才把 dict 序列化为 JSON 打印。
#
# 本脚本对应任务：把文档中的技术路线流程图改为可编辑图形、删除“资源能力盘点/
# 风险问题排查/合规性初步审查”整行、并把七个窄条节点的文字移入方框内部。
# 判定一律基于 OOXML 中形状的几何与文本（DrawingML wps 形状 + VML w:pict 形状
# 双通道），不依赖人工目测。

SCRIPT_ID = "010"
# 维度二所有评分项 max_delta 之和；用于维度一未通过 / 脚本异常时兜底
MAX_SCORE_TOTAL = 72

# 维度二所有评分项按“结果追加顺序”的 max_delta 列表；末端把 results 里的
# "+N：..." 字符串反解为结构化 dim2_items 时用。顺序须与 _evaluate_impl 中
# results.append(...) 的顺序严格一致；共 20 项，累加 = MAX_SCORE_TOTAL。
_MAX_DELTAS_ORDER = [
    5,  # 01 流程图为可编辑图形、无承载流程图的图片
    5,  # 02 连线为可单独选中的线条/箭头对象
    3,  # 03 形状颗粒度：一节点一形状、可单独移动缩放
    3,  # 04 删除指定整行 + 重接 + 清残留 + 无空洞
    5,  # 05 七个窄条节点文字落在方框边框以内
    5,  # 06 七个窄条节点文字与方框一一对应不串行
    3,  # 07 窄条节点文字单列排布、上下左右不裁切
    3,  # 08 其余各节点文字双向居中于框内
    3,  # 09 标题“技术路线图”+ 下方横向分隔线
    3,  # 10 左侧四个阶段标签
    5,  # 11 首层分四路 + 四路汇聚接入
    5,  # 12 判定分支闭环 + 不通过/通过标注
    3,  # 13 运行环境综合研判分两路 + 3/4 路分叉
    3,  # 14 EFE/IFE 汇入适配矩阵 + 结尾纵向串接
    3,  # 15 节点文字无错字漏字
    5,  # 16 A4 纵向 + 单页 + 横向居中 + 无截断
    3,  # 17 同层等宽等高对齐 + 主干中轴对齐
    3,  # 18 连线端点与节点边框相接
    3,  # 19 样式统一（填充/边框/箭头同色/字号）
    1,  # 20 相邻节点框留有清晰间距
]
assert sum(_MAX_DELTAS_ORDER) == MAX_SCORE_TOTAL

# ========== 细则中出现的节点文本清单 ==========
STAGE_LABELS = ["导入", "核验", "研判", "形成"]
ROW2_LABELS = ["政策信息梳理", "资料归集校核", "参考案例对照", "业务流程梳理"]
NARROW_OUTER = ["政策环境扫描", "需求场景识别", "协同网络分析"]
NARROW_INNER = ["资源储备评估", "流程执行能力", "技术支撑能力", "运营协调机制"]
NARROW_LABELS = NARROW_OUTER + NARROW_INNER
TRUNK_LABELS = [
    "研究导入：明确目标与范围", "需求要点识别与目标分解", "可行性确认",
    "运行环境综合研判", "适配矩阵", "实施路径定位与方案设计",
    "组织保障与质量闭环", "归纳与展望",
]
OTHER_LABELS = [
    "完善资料与方案", "外部条件研判", "内部能力评阅",
    "EFE评阅清单", "IFE评阅清单",
]
ALL_NODE_LABELS = (
    STAGE_LABELS + ROW2_LABELS + NARROW_LABELS + TRUNK_LABELS + OTHER_LABELS
)
DELETED_ROW_LABELS = ["资源能力盘点", "风险问题排查", "合规性初步审查"]
BRANCH_LABELS = ["不通过", "通过"]

# 形状几何分类
NODE_GEOMS = ("rect", "roundRect", "flowChartDecision", "diamond",
              "flowChartProcess", "flowChartAlternateProcess", "textBox")
ARROW_GEOMS = ("downArrow", "upArrow", "leftArrow", "rightArrow", "arrow")
CONN_GEOMS = ARROW_GEOMS + (
    "line", "straightConnector1", "bentConnector2", "bentConnector3",
    "bentConnector4", "bentConnector5", "curvedConnector2", "curvedConnector3",
    "curvedConnector4", "curvedConnector5", "elbowConnector",
)

# 估算用常数
LINE_H_FACTOR = 1.35   # 中文单倍行距的行高 / 字号
DEFAULT_FONT_PT = 10.5  # 未显式设定字号时按五号估算
POS_TOL = 3.0           # 同层顶边对齐容差（磅）
TRUNK_TOL = 15.0        # 主干中轴对齐容差（磅）
TOUCH_TOL = 8.0         # 连线端点与节点边框相接容差（磅）

NS = {
    'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
    'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
    'wp': 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
    'wps': 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape',
    'wpg': 'http://schemas.microsoft.com/office/word/2010/wordprocessingGroup',
    'pic': 'http://schemas.openxmlformats.org/drawingml/2006/picture',
    'mc': 'http://schemas.openxmlformats.org/markup-compatibility/2006',
    'v': 'urn:schemas-microsoft-com:vml',
    'o': 'urn:schemas-microsoft-com:office:office',
}

# ========== 基础工具：命名空间、单位与颜色 ==========
def _q(prefix, tag):
    return "{%s}%s" % (NS[prefix], tag)


def _local(tag):
    return tag.split("}")[-1] if "}" in str(tag) else str(tag)


def _pt_from_emu(v):
    try:
        return float(v) / 12700.0
    except (TypeError, ValueError):
        return 0.0


def _pt_from_twip(v):
    try:
        return float(v) / 20.0
    except (TypeError, ValueError):
        return 0.0


def _hex_to_rgb(h):
    h = str(h or "").strip().lstrip("#")
    if len(h) != 6:
        return None
    try:
        return (int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))
    except ValueError:
        return None


_VML_NAMED_COLORS = {
    "white": (255, 255, 255), "black": (0, 0, 0), "silver": (192, 192, 192),
    "gray": (128, 128, 128), "grey": (128, 128, 128), "red": (255, 0, 0),
    "lime": (0, 255, 0), "green": (0, 128, 0), "blue": (0, 0, 255),
    "yellow": (255, 255, 0), "aqua": (0, 255, 255), "fuchsia": (255, 0, 255),
    "maroon": (128, 0, 0), "navy": (0, 0, 128), "olive": (128, 128, 0),
    "purple": (128, 0, 128), "teal": (0, 128, 128), "whitesmoke": (245, 245, 245),
    "lightgray": (211, 211, 211), "lightgrey": (211, 211, 211),
}

def _parse_color_token(token):
    """解析 VML 的 fillcolor / strokecolor 取值。"""
    if not token:
        return None
    tok = str(token).strip().lower()
    if tok in _VML_NAMED_COLORS:
        return _VML_NAMED_COLORS[tok]
    return _hex_to_rgb(tok)


def _rgb_to_hsl(rgb):
    r, g, b = [c / 255.0 for c in rgb]
    mx, mn = max(r, g, b), min(r, g, b)
    lum = (mx + mn) / 2.0
    if abs(mx - mn) < 1e-9:
        return (0.0, 0.0, lum)
    d = mx - mn
    sat = d / (2.0 - mx - mn) if lum > 0.5 else d / (mx + mn)
    if mx == r:
        hue = ((g - b) / d) % 6.0
    elif mx == g:
        hue = (b - r) / d + 2.0
    else:
        hue = (r - g) / d + 4.0
    return (hue / 6.0, sat, lum)


def _hsl_to_rgb(hsl):
    hue, sat, lum = hsl
    if sat <= 1e-9:
        v = int(round(lum * 255))
        return (v, v, v)
    q = lum * (1 + sat) if lum < 0.5 else lum + sat - lum * sat
    p = 2 * lum - q

    def _ch(t):
        t = t % 1.0
        if t < 1 / 6.0:
            return p + (q - p) * 6 * t
        if t < 0.5:
            return q
        if t < 2 / 3.0:
            return p + (q - p) * (2 / 3.0 - t) * 6
        return p
    return tuple(int(round(max(0.0, min(1.0, _ch(hue + off))) * 255))
                 for off in (1 / 3.0, 0.0, -1 / 3.0))

def _load_theme_palette(zf):
    """读取 word/theme/theme1.xml 的 clrScheme，返回 {scheme_name: (r,g,b)}。"""
    palette = {}
    try:
        raw = zf.read("word/theme/theme1.xml")
    except KeyError:
        return palette
    try:
        root = ET.fromstring(raw)
    except ET.ParseError:
        return palette
    scheme = None
    for node in root.iter(_q('a', 'clrScheme')):
        scheme = node
        break
    if scheme is None:
        return palette
    for child in list(scheme):
        name = _local(child.tag)
        rgb = None
        for sub in list(child):
            sub_name = _local(sub.tag)
            if sub_name == "srgbClr":
                rgb = _hex_to_rgb(sub.get("val"))
            elif sub_name == "sysClr":
                rgb = _hex_to_rgb(sub.get("lastClr")) or (
                    (255, 255, 255) if sub.get("val") == "window" else (0, 0, 0))
            if rgb:
                break
        if rgb:
            palette[name] = rgb
    # Word 里 bg1/tx1/bg2/tx2 是对 lt1/dk1/lt2/dk2 的映射别名
    for alias, real in (("bg1", "lt1"), ("tx1", "dk1"),
                        ("bg2", "lt2"), ("tx2", "dk2")):
        if real in palette:
            palette[alias] = palette[real]
    return palette


def _resolve_fill_like(container, palette):
    """在 spPr / ln 等容器下解析 solidFill，返回 (r,g,b) 或 None。"""
    if container is None:
        return None
    for node in list(container):
        if _local(node.tag) != "solidFill":
            continue
        return _resolve_color_node(node, palette)
    return None

def _resolve_color_node(fill_node, palette):
    """解析 a:solidFill 下的 srgbClr / schemeClr，并套用 lumMod / lumOff。"""
    for child in list(fill_node):
        name = _local(child.tag)
        if name == "srgbClr":
            base = _hex_to_rgb(child.get("val"))
        elif name == "schemeClr":
            base = palette.get(child.get("val"))
            if base is None and child.get("val") in ("bg1", "lt1"):
                base = (255, 255, 255)
            if base is None and child.get("val") in ("tx1", "dk1"):
                base = (0, 0, 0)
        elif name == "sysClr":
            base = _hex_to_rgb(child.get("lastClr"))
        else:
            continue
        if base is None:
            return None
        lum_mod = lum_off = None
        for mod in list(child):
            mod_name = _local(mod.tag)
            try:
                val = int(mod.get("val")) / 100000.0
            except (TypeError, ValueError):
                continue
            if mod_name == "lumMod":
                lum_mod = val
            elif mod_name == "lumOff":
                lum_off = val
        if lum_mod is None and lum_off is None:
            return base
        hue, sat, lum = _rgb_to_hsl(base)
        if lum_mod is not None:
            lum *= lum_mod
        if lum_off is not None:
            lum += lum_off
        return _hsl_to_rgb((hue, sat, max(0.0, min(1.0, lum))))
    return None


def _is_near_white(rgb, floor=232):
    return bool(rgb) and min(rgb) >= floor


def _is_light_gray(rgb):
    if not rgb:
        return False
    return (max(rgb) - min(rgb) <= 12) and 195 <= min(rgb) <= 252 and min(rgb) < 255


def _is_greenish(rgb):
    if not rgb:
        return False
    r, g, b = rgb
    return g > r + 12 and g > b + 12

# ========== VML 样式解析 ==========
def _style_dict(style_text):
    out = {}
    for item in str(style_text or "").split(";"):
        if ":" not in item:
            continue
        key, _, val = item.partition(":")
        out[key.strip().lower()] = val.strip()
    return out


def _len_pt(text, default=None):
    """把 VML/CSS 长度串转成磅值；无单位按磅处理。"""
    if text is None:
        return default
    s = str(text).strip().lower()
    if not s:
        return default
    match = re.match(r"^(-?\d+(?:\.\d+)?)\s*(pt|px|in|cm|mm|pc|em)?$", s)
    if not match:
        return default
    val = float(match.group(1))
    unit = match.group(2) or "pt"
    factor = {"pt": 1.0, "px": 0.75, "in": 72.0, "cm": 28.3465,
              "mm": 2.83465, "pc": 12.0, "em": 12.0}[unit]
    return val * factor


def _vml_point(text):
    """解析 v:line 的 from/to 坐标串。"""
    if not text:
        return None
    parts = str(text).replace(" ", "").split(",")
    if len(parts) != 2:
        return None
    x = _len_pt(parts[0])
    y = _len_pt(parts[1])
    if x is None or y is None:
        return None
    return (x, y)


def _base_x(rel, ctx):
    if rel in ("page", "left-margin-area", "outer-margin-area"):
        return 0.0
    return ctx["left_margin"]


def _base_y(rel, ctx):
    if rel in ("page", "top-margin-area"):
        return 0.0
    if rel in ("margin", "topMargin", "top-margin-area"):
        return ctx["top_margin"]
    return ctx.get("para_base", ctx["top_margin"])

# ========== 统一形状模型 ==========
class Shape(object):
    """DrawingML(wps) 与 VML(w:pict) 两条通道归一化后的形状。"""

    def __init__(self, src, geom, x, y, w, h):
        self.src = src
        self.geom = geom or ""
        self.x = float(x)
        self.y = float(y)
        self.w = float(w)
        self.h = float(h)
        self.rot = 0.0
        self.flip_h = False
        self.flip_v = False
        self.text = ""
        self.font_pts = []
        self.jc = set()
        self.anchor = None
        self.insets = (7.2, 3.6, 7.2, 3.6)
        self.vert = False
        self.fill = None
        self.line_color = None
        self.line_pt = None
        self.has_arrow = False
        self.grouped = False
        self.pts = None          # 连线两端点 [(x1,y1),(x2,y2)]

    # 便捷几何属性
    @property
    def x2(self):
        return self.x + self.w

    @property
    def y2(self):
        return self.y + self.h

    @property
    def cx(self):
        return self.x + self.w / 2.0

    @property
    def cy(self):
        return self.y + self.h / 2.0

    @property
    def font_pt(self):
        return max(self.font_pts) if self.font_pts else DEFAULT_FONT_PT

    def is_node(self):
        return self.geom in NODE_GEOMS

    def is_conn(self):
        return self.geom in CONN_GEOMS

def _fill_text_props(shape, container):
    """从形状内部的 w:txbxContent 提取文字、字号、对齐。"""
    if container is None:
        return
    chunks = []
    for para in container.iter(_q('w', 'p')):
        line = "".join(t.text or "" for t in para.iter(_q('w', 't')))
        if line:
            chunks.append(line)
        for jc in para.iter(_q('w', 'jc')):
            shape.jc.add(jc.get(_q('w', 'val')))
    if not chunks:
        chunks = ["".join(t.text or "" for t in container.iter(_q('w', 't')))]
    shape.text = "".join(chunks).strip()
    for sz in container.iter(_q('w', 'sz')):
        try:
            shape.font_pts.append(int(sz.get(_q('w', 'val'))) / 2.0)
        except (TypeError, ValueError):
            continue


def _compute_endpoints(shape):
    """给出连线/箭头的两个端点（已考虑 flip 与 rot）。"""
    if shape.pts:
        return shape.pts
    horizontal = shape.w >= shape.h
    if shape.geom in ("leftArrow", "rightArrow"):
        horizontal = True
    elif shape.geom in ("downArrow", "upArrow"):
        horizontal = False
    if horizontal:
        p1 = (-shape.w / 2.0, 0.0)
        p2 = (shape.w / 2.0, 0.0)
    else:
        p1 = (0.0, -shape.h / 2.0)
        p2 = (0.0, shape.h / 2.0)
    if shape.geom not in ARROW_GEOMS:
        # 折线/直线连接符：端点取包围盒对角，flip 决定走向
        p1 = (-shape.w / 2.0, -shape.h / 2.0)
        p2 = (shape.w / 2.0, shape.h / 2.0)
        if shape.flip_h != shape.flip_v:
            p1, p2 = (-p1[0], p1[1]), (-p2[0], p2[1])
    if abs(shape.rot) > 1e-6:
        rad = math.radians(shape.rot)
        cos_r, sin_r = math.cos(rad), math.sin(rad)
        p1 = (p1[0] * cos_r - p1[1] * sin_r, p1[0] * sin_r + p1[1] * cos_r)
        p2 = (p2[0] * cos_r - p2[1] * sin_r, p2[0] * sin_r + p2[1] * cos_r)
    shape.pts = [(shape.cx + p1[0], shape.cy + p1[1]),
                 (shape.cx + p2[0], shape.cy + p2[1])]
    return shape.pts

# ========== DrawingML 通道 ==========
def _dml_anchor_box(anchor, ctx):
    ext = anchor.find(_q('wp', 'extent'))
    w = _pt_from_emu(ext.get('cx')) if ext is not None else 0.0
    h = _pt_from_emu(ext.get('cy')) if ext is not None else 0.0
    x = y = None
    for tag, setter in (('positionH', 'x'), ('positionV', 'y')):
        node = anchor.find(_q('wp', tag))
        if node is None:
            continue
        rel = node.get('relativeFrom') or ''
        base = _base_x(rel, ctx) if setter == 'x' else _base_y(rel, ctx)
        off = node.find(_q('wp', 'posOffset'))
        if off is not None and off.text:
            val = base + _pt_from_emu(off.text)
        else:
            align = node.find(_q('wp', 'align'))
            word = (align.text or '').strip() if align is not None else ''
            if setter == 'x':
                span = ctx['page_w'] if rel == 'page' else ctx['body_w']
                val = base + (span - w) / 2.0 if word == 'center' else base
            else:
                val = base
        if setter == 'x':
            x = val
        else:
            y = val
    if x is None:
        x = ctx['left_margin']
    if y is None:
        y = ctx.get('para_base', ctx['top_margin'])
    return x, y, w, h


def _xfrm_of(elem):
    for node in elem.iter(_q('a', 'xfrm')):
        return node
    return None

def _group_frame(grp, ax, ay):
    """解析 wpg 组合的子坐标映射，返回 (基点x, 基点y, 子原点x, 子原点y, sx, sy)。"""
    grp_pr = grp.find(_q('wpg', 'grpSpPr'))
    xf = grp_pr.find(_q('a', 'xfrm')) if grp_pr is not None else None
    if xf is None:
        return None
    ext = xf.find(_q('a', 'ext'))
    ch_ext = xf.find(_q('a', 'chExt'))
    ch_off = xf.find(_q('a', 'chOff'))
    if ext is None or ch_ext is None:
        return None
    cw = _pt_from_emu(ch_ext.get('cx')) or 1.0
    ch = _pt_from_emu(ch_ext.get('cy')) or 1.0
    sx = (_pt_from_emu(ext.get('cx')) or cw) / cw
    sy = (_pt_from_emu(ext.get('cy')) or ch) / ch
    ox = _pt_from_emu(ch_off.get('x')) if ch_off is not None else 0.0
    oy = _pt_from_emu(ch_off.get('y')) if ch_off is not None else 0.0
    return (ax, ay, ox, oy, sx, sy)


def _parse_dml_anchor(anchor, ctx, shapes, images):
    ax, ay, aw, ah = _dml_anchor_box(anchor, ctx)
    gdata = None
    for node in anchor.iter(_q('a', 'graphicData')):
        gdata = node
        break
    if gdata is None:
        return
    pics = list(gdata.iter(_q('pic', 'pic')))
    wsps = list(gdata.iter(_q('wps', 'wsp')))
    for _pic in pics:
        images.append(Shape('dml', 'picture', ax, ay, aw, ah))
    if not wsps:
        return
    grp = None
    for node in gdata.iter(_q('wpg', 'wgp')):
        grp = node
        break
    frame = _group_frame(grp, ax, ay) if grp is not None else None
    for wsp in wsps:
        shape = _build_wsp(wsp, ax, ay, aw, ah, frame, ctx)
        if shape is None:
            continue
        shape.grouped = bool(grp is not None or len(wsps) > 1)
        shapes.append(shape)

def _build_wsp(wsp, ax, ay, aw, ah, frame, ctx):
    palette = ctx["palette"]
    sp_pr = wsp.find(_q('wps', 'spPr'))
    geom = ""
    if sp_pr is not None:
        for node in sp_pr.iter(_q('a', 'prstGeom')):
            geom = node.get('prst') or ""
            break
        if not geom and sp_pr.find(_q('a', 'custGeom')) is not None:
            geom = "custGeom"
    xf = sp_pr.find(_q('a', 'xfrm')) if sp_pr is not None else None
    x, y, w, h = ax, ay, aw, ah
    if xf is not None:
        ext = xf.find(_q('a', 'ext'))
        if ext is not None:
            w = _pt_from_emu(ext.get('cx'))
            h = _pt_from_emu(ext.get('cy'))
        off = xf.find(_q('a', 'off'))
        if frame is not None and off is not None:
            gx, gy, cox, coy, sx, sy = frame
            x = gx + (_pt_from_emu(off.get('x')) - cox) * sx
            y = gy + (_pt_from_emu(off.get('y')) - coy) * sy
            w *= sx
            h *= sy
    shape = Shape('dml', geom or 'rect', x, y, w, h)
    if xf is not None:
        shape.flip_h = xf.get('flipH') in ('1', 'true')
        shape.flip_v = xf.get('flipV') in ('1', 'true')
        try:
            shape.rot = int(xf.get('rot')) / 60000.0
        except (TypeError, ValueError):
            shape.rot = 0.0
    if sp_pr is not None:
        shape.fill = _resolve_fill_like(sp_pr, palette)
        if sp_pr.find(_q('a', 'noFill')) is not None:
            shape.fill = None
        ln = sp_pr.find(_q('a', 'ln'))
        if ln is not None:
            shape.line_color = _resolve_fill_like(ln, palette)
            if ln.get('w'):
                shape.line_pt = _pt_from_emu(ln.get('w'))
            for end in list(ln):
                if _local(end.tag) in ('headEnd', 'tailEnd') and \
                        (end.get('type') or 'none') != 'none':
                    shape.has_arrow = True
    if shape.geom in ARROW_GEOMS:
        shape.has_arrow = True
    _apply_body_pr(shape, wsp)
    return shape

def _apply_body_pr(shape, wsp):
    body = wsp.find(_q('wps', 'bodyPr'))
    if body is not None:
        shape.anchor = body.get('anchor')
        shape.vert = (body.get('vert') or '') in (
            'eaVert', 'vert', 'vert270', 'mongolianVert')
        ins = []
        for key, dflt in (('lIns', 7.2), ('tIns', 3.6),
                          ('rIns', 7.2), ('bIns', 3.6)):
            val = body.get(key)
            ins.append(_pt_from_emu(val) if val is not None else dflt)
        shape.insets = tuple(ins)
    txbx = wsp.find(_q('wps', 'txbx'))
    if txbx is not None:
        _fill_text_props(shape, txbx.find(_q('w', 'txbxContent')))


# ========== VML 通道 ==========
_VML_SHAPE_TAGS = ('rect', 'roundrect', 'oval', 'line', 'shape',
                   'polyline', 'curve', 'arc')
_VML_GEOM_MAP = {'rect': 'rect', 'roundrect': 'roundRect', 'oval': 'ellipse',
                 'line': 'line', 'polyline': 'line', 'curve': 'line',
                 'arc': 'line'}


_VML_TYPE_GEOM = {
    '1': 'rect', '2': 'roundRect', '3': 'ellipse', '4': 'flowChartDecision',
    '13': 'rightArrow', '32': 'line', '33': 'bentConnector3',
    '34': 'bentConnector3', '35': 'bentConnector3', '37': 'curvedConnector3',
    '38': 'curvedConnector3', '66': 'leftArrow', '67': 'downArrow',
    '68': 'upArrow', '75': 'picture', '202': 'textBox',
}


def _vml_geom(elem, style):
    tag = _local(elem.tag)
    if tag in _VML_GEOM_MAP:
        return _VML_GEOM_MAP[tag]
    if elem.get(_q('o', 'connectortype')):
        return 'line'
    match = re.search(r"_x0000_t(\d+)", str(elem.get('type') or ''))
    if match:
        return _VML_TYPE_GEOM.get(match.group(1), 'rect')
    if 'layout-flow' in style or elem.find(_q('v', 'textbox')) is not None:
        return 'textBox'
    return 'rect'

def _parse_vml_shape(elem, ctx, shapes, images):
    style = _style_dict(elem.get('style'))
    geom = _vml_geom(elem, style)
    rel_h = (style.get('mso-position-horizontal-relative') or 'text').lower()
    rel_v = (style.get('mso-position-vertical-relative') or 'text').lower()
    base_x = _base_x('page' if rel_h == 'page' else rel_h, ctx)
    base_y = _base_y('page' if rel_v == 'page' else
                     ('margin' if rel_v == 'margin' else 'paragraph'), ctx)
    # Word 只认 margin-left / margin-top；CSS 的 left / top 不参与定位
    off_x = _len_pt(style.get('margin-left'), 0.0) or 0.0
    off_y = _len_pt(style.get('margin-top'), 0.0) or 0.0
    width = _len_pt(style.get('width'), 0.0) or 0.0
    height = _len_pt(style.get('height'), 0.0) or 0.0
    p_from = _vml_point(elem.get('from'))
    p_to = _vml_point(elem.get('to'))
    if p_from and p_to:
        x1, y1 = base_x + off_x + p_from[0], base_y + off_y + p_from[1]
        x2, y2 = base_x + off_x + p_to[0], base_y + off_y + p_to[1]
        shape = Shape('vml', 'line', min(x1, x2), min(y1, y2),
                      abs(x2 - x1), abs(y2 - y1))
        shape.pts = [(x1, y1), (x2, y2)]
    else:
        pos_x = base_x + off_x
        if (style.get('mso-position-horizontal') or '').lower() == 'center':
            span = ctx['page_w'] if rel_h == 'page' else ctx['body_w']
            pos_x = base_x + (span - width) / 2.0
        shape = Shape('vml', geom, pos_x, base_y + off_y, width, height)
    if elem.find(_q('v', 'imagedata')) is not None:
        shape.geom = 'picture'
        images.append(shape)
        return
    if str(style.get('flip') or '').lower() in ('x', 'x y', 'xy'):
        shape.flip_h = True
    if str(style.get('flip') or '').lower() in ('y', 'x y', 'xy'):
        shape.flip_v = True
    rot_raw = str(style.get('rotation') or '').strip().lower()
    if rot_raw:
        try:
            if rot_raw.endswith('fd') or rot_raw.endswith('f'):
                shape.rot = float(rot_raw.rstrip('fd')) / 65536.0
            else:
                shape.rot = float(rot_raw)
        except ValueError:
            shape.rot = 0.0
    _apply_vml_style(shape, elem, style)
    shapes.append(shape)

def _apply_vml_style(shape, elem, style):
    filled = str(elem.get('filled') or 'true').lower() not in ('f', 'false')
    if filled:
        shape.fill = _parse_color_token(elem.get('fillcolor')) or (255, 255, 255)
    stroked = str(elem.get('stroked') or 'true').lower() not in ('f', 'false')
    if stroked:
        shape.line_color = _parse_color_token(elem.get('strokecolor')) or (0, 0, 0)
    shape.line_pt = _len_pt(elem.get('strokeweight'), None)
    v_fill = elem.find(_q('v', 'fill'))
    if v_fill is not None and v_fill.get('color'):
        shape.fill = _parse_color_token(v_fill.get('color'))
    v_stroke = elem.find(_q('v', 'stroke'))
    if v_stroke is not None:
        if v_stroke.get('color'):
            shape.line_color = _parse_color_token(v_stroke.get('color'))
        if v_stroke.get('weight'):
            shape.line_pt = _len_pt(v_stroke.get('weight'), shape.line_pt)
        for key in ('endarrow', 'startarrow'):
            if str(v_stroke.get(key) or 'none').lower() != 'none':
                shape.has_arrow = True
    if shape.geom in ARROW_GEOMS:
        shape.has_arrow = True
    box = elem.find(_q('v', 'textbox'))
    if box is None:
        return
    box_style = _style_dict(box.get('style'))
    if 'vertical' in str(box_style.get('layout-flow') or ''):
        shape.vert = True
    if str(box_style.get('v-text-anchor') or '').startswith('middle'):
        shape.anchor = 'ctr'
    inset = str(box.get('inset') or '').split(',')
    if len(inset) == 4:
        dflt = (7.2, 3.6, 7.2, 3.6)
        shape.insets = tuple(_len_pt(inset[i], dflt[i]) or dflt[i]
                             for i in range(4))
    _fill_text_props(shape, box.find(_q('w', 'txbxContent')))
    if not shape.text:
        shape.text = "".join(t.text or "" for t in box.iter(_q('w', 't'))).strip()

# ========== 遍历文档：两条通道合并采集 ==========
def _walk(elem, ctx, shapes, images):
    if elem.tag == _q('mc', 'Fallback'):
        return  # 与 wps 的 Choice 等价，跳过以免形状/文字重复计数
    tag = _local(elem.tag)
    if tag in ('anchor', 'inline') and elem.tag.startswith('{%s}' % NS['wp']):
        _parse_dml_anchor(elem, ctx, shapes, images)
        return
    if elem.tag.startswith('{%s}' % NS['v']):
        if tag == 'shapetype':
            return
        if tag == 'group':
            before = len(shapes)
            for child in list(elem):
                _walk(child, ctx, shapes, images)
            for shape in shapes[before:]:
                shape.grouped = True
            return
        if tag in _VML_SHAPE_TAGS:
            _parse_vml_shape(elem, ctx, shapes, images)
            return
    for child in list(elem):
        _walk(child, ctx, shapes, images)


def _para_direct_text(para):
    out = []
    for run in para.findall(_q('w', 'r')):
        out.append("".join(t.text or "" for t in run.findall(_q('w', 't'))))
    return "".join(out)


def _para_direct_sizes(para):
    sizes = []
    holders = [para.find(_q('w', 'pPr'))] + para.findall(_q('w', 'r'))
    for holder in holders:
        if holder is None:
            continue
        rpr = holder.find(_q('w', 'rPr'))
        if rpr is None:
            continue
        sz = rpr.find(_q('w', 'sz'))
        if sz is not None:
            try:
                sizes.append(int(sz.get(_q('w', 'val'))) / 2.0)
            except (TypeError, ValueError):
                continue
    return sizes

def _para_info(para, index):
    ppr = para.find(_q('w', 'pPr'))
    jc = None
    bottom_border = False
    if ppr is not None:
        jc_node = ppr.find(_q('w', 'jc'))
        if jc_node is not None:
            jc = jc_node.get(_q('w', 'val'))
        pbdr = ppr.find(_q('w', 'pBdr'))
        if pbdr is not None:
            for side in ('bottom', 'top'):
                node = pbdr.find(_q('w', side))
                if node is not None and (node.get(_q('w', 'val')) or 'none') \
                        not in ('none', 'nil'):
                    bottom_border = True
    sizes = _para_direct_sizes(para)
    return {
        'index': index,
        'text': _para_direct_text(para).strip(),
        'jc': jc,
        'sizes': sizes,
        'font_pt': max(sizes) if sizes else DEFAULT_FONT_PT,
        'border': bottom_border,
    }


def _estimate_para_height(info):
    return max(info['font_pt'], DEFAULT_FONT_PT) * LINE_H_FACTOR


def _scan_body(body, ctx, shapes, images, paras):
    """按正文流顺序扫描，边推进 para_base 边采集浮动形状。"""
    cursor = ctx['top_margin']
    index = 0
    for child in list(body):
        tag = _local(child.tag)
        if tag == 'sectPr':
            continue
        ctx['para_base'] = cursor
        if tag == 'p':
            info = _para_info(child, index)
            paras.append(info)
            _walk(child, ctx, shapes, images)
            cursor += _estimate_para_height(info)
            index += 1
        else:
            _walk(child, ctx, shapes, images)
            cursor += DEFAULT_FONT_PT * LINE_H_FACTOR

def _all_text(elem, out=None):
    """收集全文文字，跳过 mc:Fallback 以免与 Choice 重复。"""
    if out is None:
        out = []
    if elem.tag == _q('mc', 'Fallback'):
        return out
    if elem.tag == _q('w', 't') and elem.text:
        out.append(elem.text)
    for child in list(elem):
        _all_text(child, out)
    return out


def _load_document(file_path):
    with zipfile.ZipFile(file_path) as zf:
        raw = zf.read('word/document.xml')
        palette = _load_theme_palette(zf)
    root = ET.fromstring(raw)
    body = root.find(_q('w', 'body'))
    sect = None
    for node in root.iter(_q('w', 'sectPr')):
        sect = node
    page_w, page_h = 595.3, 841.9
    left = right = 90.0
    top = 51.0
    bottom = 51.0
    if sect is not None:
        pg_sz = sect.find(_q('w', 'pgSz'))
        if pg_sz is not None:
            page_w = _pt_from_twip(pg_sz.get(_q('w', 'w'))) or page_w
            page_h = _pt_from_twip(pg_sz.get(_q('w', 'h'))) or page_h
        pg_mar = sect.find(_q('w', 'pgMar'))
        if pg_mar is not None:
            left = _pt_from_twip(pg_mar.get(_q('w', 'left')))
            right = _pt_from_twip(pg_mar.get(_q('w', 'right')))
            top = _pt_from_twip(pg_mar.get(_q('w', 'top')))
            bottom = _pt_from_twip(pg_mar.get(_q('w', 'bottom')))
    ctx = {
        'palette': palette, 'page_w': page_w, 'page_h': page_h,
        'left_margin': left, 'right_margin': right,
        'top_margin': top, 'bottom_margin': bottom,
        'body_w': page_w - left - right, 'body_h': page_h - top - bottom,
        'para_base': top,
    }
    shapes, images, paras = [], [], []
    if body is not None:
        _scan_body(body, ctx, shapes, images, paras)
    return {'ctx': ctx, 'shapes': shapes, 'images': images, 'paras': paras,
            'root': root, 'raw': raw.decode('utf-8', 'ignore'),
            'full_text': "".join(_all_text(root))}

# ========== 几何工具 ==========
def _axis_gap(a_lo, a_hi, b_lo, b_hi):
    if a_hi < b_lo:
        return b_lo - a_hi
    if b_hi < a_lo:
        return a_lo - b_hi
    return -min(a_hi, b_hi) + max(a_lo, b_lo)  # 负值表示重叠长度的相反数


def _rect_gap(s1, s2):
    """两个矩形的最短间距；重叠时返回 0。"""
    dx = _axis_gap(s1.x, s1.x2, s2.x, s2.x2)
    dy = _axis_gap(s1.y, s1.y2, s2.y, s2.y2)
    if dx <= 0 and dy <= 0:
        return 0.0
    if dx > 0 and dy > 0:
        return math.hypot(dx, dy)
    return max(dx, dy)


def _overlap_area(s1, s2):
    ow = min(s1.x2, s2.x2) - max(s1.x, s2.x)
    oh = min(s1.y2, s2.y2) - max(s1.y, s2.y)
    if ow <= 0 or oh <= 0:
        return 0.0
    return ow * oh


def _point_to_rect_dist(pt, shape):
    """点到矩形（含内部记 0）的距离。"""
    dx = max(shape.x - pt[0], 0.0, pt[0] - shape.x2)
    dy = max(shape.y - pt[1], 0.0, pt[1] - shape.y2)
    return math.hypot(dx, dy)


def _point_to_seg_dist(pt, seg_a, seg_b):
    ax, ay = seg_a
    bx, by = seg_b
    dx, dy = bx - ax, by - ay
    if abs(dx) < 1e-9 and abs(dy) < 1e-9:
        return math.hypot(pt[0] - ax, pt[1] - ay)
    t = ((pt[0] - ax) * dx + (pt[1] - ay) * dy) / (dx * dx + dy * dy)
    t = max(0.0, min(1.0, t))
    return math.hypot(pt[0] - (ax + t * dx), pt[1] - (ay + t * dy))

def _norm(text):
    return re.sub(r"[\s　]+", "", str(text or ""))


def _find_by_text(shapes, label, exact_only=False):
    target = _norm(label)
    if not target:
        return None
    for shape in shapes:
        if _norm(shape.text) == target:
            return shape
    if exact_only:
        return None
    for shape in shapes:
        if target in _norm(shape.text):
            return shape
    return None


def _touches_edge(conn, node, side, tol=TOUCH_TOL):
    for px, py in _compute_endpoints(conn):
        if side == 'top' and abs(py - node.y) <= tol and \
                node.x - tol <= px <= node.x2 + tol:
            return True
        if side == 'bottom' and abs(py - node.y2) <= tol and \
                node.x - tol <= px <= node.x2 + tol:
            return True
        if side == 'left' and abs(px - node.x) <= tol and \
                node.y - tol <= py <= node.y2 + tol:
            return True
        if side == 'right' and abs(px - node.x2) <= tol and \
                node.y - tol <= py <= node.y2 + tol:
            return True
    return False


def _arrows_into_top(conns, node, tol=TOUCH_TOL):
    """自上而下落在 node 顶边的箭头。"""
    out = []
    for conn in conns:
        if conn.cy > node.y + tol:
            continue
        if _touches_edge(conn, node, 'top', tol) and conn.has_arrow:
            out.append(conn)
    return out


def _conn_down_between(conns, upper, lower, tol=TOUCH_TOL):
    for conn in conns:
        if _touches_edge(conn, upper, 'bottom', tol) and \
                _touches_edge(conn, lower, 'top', tol):
            return conn
    return None

# ========== 结构化模型 ==========
def _build_model(doc):
    shapes = doc['shapes']
    nodes = [s for s in shapes if s.is_node()]
    conns = [s for s in shapes if s.is_conn()]
    text_nodes = [s for s in nodes if _norm(s.text)]
    label_map = {}
    for shape in text_nodes:
        label_map.setdefault(_norm(shape.text), []).append(shape)
    named = {}
    for label in ALL_NODE_LABELS:
        hit = label_map.get(_norm(label))
        if hit:
            named[label] = hit[0]
    model = dict(doc)
    model.update({
        'shapes': shapes, 'nodes': nodes, 'conns': conns,
        'text_nodes': text_nodes, 'label_map': label_map, 'named': named,
        'narrow': [named[k] for k in NARROW_LABELS if k in named],
        'row2': [named[k] for k in ROW2_LABELS if k in named],
        'stage': [named[k] for k in STAGE_LABELS if k in named],
        'trunk': [named[k] for k in TRUNK_LABELS if k in named],
    })
    return model


def _text_fits(shape, chars=None, single_column=True):
    """按字号估算文字块是否落在框线以内。"""
    text = _norm(shape.text)
    count = chars if chars is not None else len(text)
    if count <= 0:
        return True, 0.0, 0.0
    font = shape.font_pt
    border = (shape.line_pt or 1.0)
    if single_column:
        need_w = font
        need_h = count * font * LINE_H_FACTOR
    else:
        need_w = count * font
        need_h = font * LINE_H_FACTOR
    return (need_w <= shape.w - border and need_h <= shape.h - border,
            need_w, need_h)

def _check_dim1(model):
    """可用与可修改性：交付文件为 1 个可正常打开的 Word 文档（.docx）。

    文件损坏或非法 docx 时，_open_word_document / _load_document 会直接抛异常，
    由 evaluate() 兜底为维度一不通过、总分 0；能走到这里即说明文档可正常打开。
    """
    if not model or 'ctx' not in model:
        return False, "文档正文结构异常，无法作为 Word 文档（.docx）正常打开"
    return True, ""

# ========== 定位与打开 ==========
class DocxNotFoundError(Exception):
    """脚本所在目录内未找到可评估的 .docx 文件。"""


def _open_word_document(dir_path):
    """在脚本所在目录内定位并打开被评估的 .docx。"""
    if os.path.isfile(dir_path) and dir_path.lower().endswith('.docx'):
        return dir_path, Document(dir_path)
    names = sorted(name for name in os.listdir(dir_path)
                   if name.lower().endswith('.docx')
                   and not name.startswith('~$'))
    if not names:
        raise DocxNotFoundError("目录中未找到 .docx 文件：%s" % dir_path)
    path = os.path.join(dir_path, names[0])
    return path, Document(path)


def _layers_and_gaps(nodes):
    """按顶边把主流程节点聚成层，并给出相邻层的纵向净间距。"""
    boxes = sorted(nodes, key=lambda s: s.y)
    layers = []
    for box in boxes:
        if layers and abs(box.y - layers[-1][0].y) <= 6.0:
            layers[-1].append(box)
        else:
            layers.append([box])
    gaps = []
    for i in range(len(layers) - 1):
        gaps.append(min(s.y for s in layers[i + 1]) -
                    max(s.y2 for s in layers[i]))
    return layers, gaps


def _conn_color(conn):
    """连线的可见颜色：箭头形状看填充，直线/折线看描边。"""
    if conn.geom in ARROW_GEOMS:
        return conn.fill or conn.line_color
    return conn.line_color or conn.fill


def _main_flow_nodes(model):
    keys = (ROW2_LABELS + NARROW_LABELS + TRUNK_LABELS +
            ["外部条件研判", "内部能力评阅", "EFE评阅清单", "IFE评阅清单"])
    named = model['named']
    return [named[k] for k in keys if k in named]

# ========== 维度二：20 项逐条判定 ==========
def _evaluate_impl(file_path, result):
    doc = _load_document(file_path)
    model = _build_model(doc)
    dim1_pass, dim1_reason = _check_dim1(model)
    result["dim1_pass"] = dim1_pass
    result["dim1_reason"] = dim1_reason
    if not dim1_pass:
        result["dim2_items"] = []
        result["total_score"] = 0
        return result

    results = []

    def add(hit, rule, detail=""):
        # 约定：输出中不含 detail 字段（detail 参数仅保留调用形态，不参与输出）
        max_delta = _MAX_DELTAS_ORDER[len(results)]
        results.append("+%d：%s" % (max_delta if hit else 0, rule))

    ctx = model['ctx']
    nodes = model['nodes']
    conns = model['conns']
    named = model['named']
    narrow = model['narrow']
    row2 = model['row2']
    stage = model['stage']
    trunk = model['trunk']
    text_nodes = model['text_nodes']
    full_text = _norm(model['full_text'])
    body_center = ctx['left_margin'] + ctx['body_w'] / 2.0
    non_title = [n for n in nodes if _norm(n.text) != "技术路线图"]
    top_ref = min([n.y for n in non_title] or [0.0])
    # 标题下方的横向分隔线属装饰，不计入流程连线
    flow_conns = [c for c in conns if c.y2 > top_ref + 2.0]

    # 01 流程图为可编辑图形、无承载流程图的图片
    body_area = max(ctx['body_w'] * ctx['body_h'], 1.0)
    big_images = [im for im in model['images'] if im.w * im.h >= 0.2 * body_area]
    hit = len(text_nodes) >= 20 and not big_images
    add(hit, "流程图为可编辑图形，全部节点是 Word 形状或文本框，文档内无承载流程图的图片",
        "带文字形状 %d 个，图片 %d 张（大幅 %d 张）"
        % (len(text_nodes), len(model['images']), len(big_images)))
    # 02 连线为可单独选中的线条/箭头对象，不用字符拼凑
    arrow_conns = [c for c in conns if c.has_arrow]
    char_fake = re.findall(r"[↓↑←→│┃|‖—─━＿_]{2,}", model['full_text'])
    hit = len(conns) >= 20 and len(arrow_conns) >= 12 and not char_fake
    add(hit, "节点间连接为可单独选中拖动的线条与箭头对象，自上而下贯通，未用字符拼凑",
        "连线对象 %d 个（其中带箭头 %d 个），字符拼凑串 %d 处"
        % (len(conns), len(arrow_conns), len(char_fake)))

    # 03 形状颗粒度：一节点一形状、可单独移动缩放
    grouped = [s for s in model['shapes'] if s.grouped]
    dup_labels = [k for k, v in model['label_map'].items() if len(v) > 1]
    mixed = []
    for shape in text_nodes:
        txt = _norm(shape.text)
        matched = [_norm(lbl) for lbl in ALL_NODE_LABELS if _norm(lbl) in txt]
        # 阶段标签（如“导入”“研判”）本身是长节点文字的子串，不算合并
        independent = [a for a in matched
                       if not any(a != b and a in b for b in matched)]
        if len(independent) >= 2:
            mixed.append(txt[:12])
    locked = re.findall(r'no(?:Move|Resize|Select|Drilldown)="(?:1|true)"',
                        model['raw'])
    # 无可逐个改动的节点形状时，各项“无违规”均为空集，不能据此给分
    hit = (len(text_nodes) >= 12 and not grouped and not dup_labels
           and not mixed and not locked)
    add(hit, "一个节点一个形状、文字装在该形状内，每个形状可单独移动缩放，无锁定与不可拆分的组合",
        "带文字形状 %d 个，组合内形状 %d 个，重复文字 %d 组，多节点合并形状 %d 个，锁定属性 %d 处"
        % (len(text_nodes), len(grouped), len(dup_labels), len(mixed), len(locked)))

    # 04 删除指定整行 + 重接 + 清残留 + 无空洞
    remain = [lbl for lbl in DELETED_ROW_LABELS if _norm(lbl) in full_text]
    need_node = named.get("需求要点识别与目标分解")
    conf_node = named.get("可行性确认")
    direct = (_conn_down_between(conns, need_node, conf_node)
              if need_node is not None and conf_node is not None else None)
    _layers, gaps = _layers_and_gaps(_main_flow_nodes(model))
    gap_ok = False
    band = []
    if need_node is not None and conf_node is not None and gaps:
        this_gap = conf_node.y - need_node.y2
        gap_ok = this_gap <= 2.0 * statistics.median(gaps)
        band = [c for c in conns
                if c.y >= need_node.y2 - 2.0 and c.y2 <= conf_node.y + 2.0]
    hit = (not remain and direct is not None and gap_ok and len(band) <= 2)
    add(hit, "删除“资源能力盘点”“风险问题排查”“合规性初步审查”整行，"
             "“需求要点识别与目标分解”一条箭头直连“可行性确认”，三分叉残留清除、纵向无空洞",
        "残留文字 %s，直连箭头 %s，层间距 %s，区间内连线 %d 条"
        % (remain or "无", "有" if direct is not None else "无",
           "正常" if gap_ok else "偏大", len(band)))
    # 05 七个窄条节点文字落在方框边框以内
    out_of_box = []
    for lbl in NARROW_LABELS:
        shape = named.get(lbl)
        if shape is None:
            out_of_box.append("%s(未找到承载形状)" % lbl)
            continue
        if shape.h < 2.0 * shape.w:
            out_of_box.append("%s(承载形状非窄条竖框)" % lbl)
            continue
        fits, need_w, need_h = _text_fits(shape)
        if not fits:
            out_of_box.append("%s(需%.1f×%.1f 磅>框%.1f×%.1f 磅)"
                              % (lbl, need_w, need_h, shape.w, shape.h))
    hit = not out_of_box
    add(hit, "七个窄条节点的文字均落在各自方框边框以内，无一处留在框外",
        "越界节点：%s" % ("、".join(out_of_box) if out_of_box else "无"))

    # 06 七个窄条节点文字与方框一一对应不串行
    order_detail = "窄条节点仅检出 %d 个" % len(narrow)
    hit = False
    if len(narrow) == 7:
        by_x = sorted(narrow, key=lambda s: s.x)
        got = [_norm(s.text) for s in by_x]
        want = [_norm(lbl) for lbl in NARROW_LABELS]
        outer_node = named.get("外部条件研判")
        inner_node = named.get("内部能力评阅")
        order_ok = got == want
        belong_ok = False
        if outer_node is not None and inner_node is not None:
            outer_cx = statistics.mean([s.cx for s in by_x[:3]])
            inner_cx = statistics.mean([s.cx for s in by_x[3:]])
            belong_ok = (abs(outer_cx - outer_node.cx) <= 40.0 and
                         abs(inner_cx - inner_node.cx) <= 40.0)
        hit = order_ok and belong_ok
        order_detail = "自左至右为 %s；分组归属%s" % (
            "、".join(got), "正确" if belong_ok else "错位")
    add(hit, "七个窄条节点文字与方框一一对应不串行，"
             "“外部条件研判”下为三个、“内部能力评阅”下为四个且自左至右顺序正确",
        order_detail)

    # 07 窄条节点文字单列排布、上下左右不裁切
    clip = []
    for lbl in NARROW_LABELS:
        shape = named.get(lbl)
        if shape is None:
            clip.append("%s(缺失)" % lbl)
            continue
        font = shape.font_pt
        inner_w = shape.w - shape.insets[0] - shape.insets[2]
        inner_h = shape.h - shape.insets[1] - shape.insets[3]
        if inner_w >= 2.0 * font:
            clip.append("%s(框内宽 %.1f 磅可容两列)" % (lbl, inner_w))
        elif len(_norm(shape.text)) * font * LINE_H_FACTOR > inner_h + 1.0:
            clip.append("%s(单列文字高 %.1f 磅>框内高 %.1f 磅)"
                        % (lbl, len(_norm(shape.text)) * font * LINE_H_FACTOR,
                           inner_h))
    hit = not clip
    add(hit, "窄条节点文字每行一字自上而下排成一列，上下不被框顶框底裁切、左右不被框线压字",
        "异常节点：%s" % ("、".join(clip) if clip else "无"))
    # 08 其余各节点文字双向居中于框内
    others = [named[lbl] for lbl in
              (TRUNK_LABELS + ROW2_LABELS + OTHER_LABELS) if lbl in named]
    off_center = []
    for shape in others:
        h_ok = bool(shape.jc) and shape.jc.issubset({"center", "both", "distribute"}) \
            and "center" in shape.jc
        fits, need_w, need_h = _text_fits(shape, single_column=False)
        inner_h = shape.h - shape.insets[1] - shape.insets[3]
        # anchor 非 ctr 时，若文字高度已几乎填满框内高度，视觉上仍是垂直居中
        v_ok = shape.anchor == "ctr" or (inner_h - need_h) <= 4.0
        if not (h_ok and v_ok and fits):
            off_center.append("%s(水平%s/垂直%s/容纳%s)"
                              % (_norm(shape.text)[:10],
                                 "居中" if h_ok else "非居中",
                                 "居中" if v_ok else "非居中",
                                 "正常" if fits else "溢出"))
    hit = bool(others) and not off_center
    add(hit, "其余各节点文字水平与垂直双向居中于框内，不压框线、不溢出框外",
        "异常节点：%s" % ("、".join(off_center) if off_center else "无"))

    # 09 标题“技术路线图”+ 下方横向分隔线
    title_paras = [p for p in model['paras'] if _norm(p['text']) == "技术路线图"]
    title_shape = _find_by_text(text_nodes, "技术路线图", exact_only=True)
    title_pt = 0.0
    title_center = False
    if title_paras:
        title_pt = max(p['font_pt'] for p in title_paras)
        title_center = any(p['jc'] == 'center' for p in title_paras)
    elif title_shape is not None:
        title_pt = title_shape.font_pt
        title_center = abs(title_shape.cx - body_center) <= 30.0
    node_font_max = max([n.font_pt for n in non_title if _norm(n.text)] or [0.0])
    separator = None
    for conn in conns:
        if conn.w >= 80.0 and conn.h <= 4.0 and conn.y2 <= top_ref + 4.0:
            separator = conn
            break
    if separator is None and any(p['border'] for p in title_paras):
        separator = "pBdr"
    hit = (title_pt >= 20.0 and title_center and title_pt > node_font_max
           and separator is not None)
    add(hit, "标题“技术路线图”保留，单行居中、字号≥20 磅且明显大于节点文字，标题下方保留横向分隔线",
        "标题字号 %.1f 磅（节点最大 %.1f 磅），居中%s，分隔线%s"
        % (title_pt, node_font_max, "是" if title_center else "否",
           "有" if separator is not None else "无"))
    # 10 左侧四个阶段标签
    stage_detail = "阶段标签检出 %d 个" % len(stage)
    hit = False
    if len(stage) == 4:
        xs = [s.x for s in stage]
        rest_x = min([n.x for n in nodes if n not in stage] or [0.0])
        vertical = all(s.h >= 2.0 * s.w for s in stage)
        rounded = all(s.geom.startswith("roundRect") for s in stage)
        leftmost = max(s.x2 for s in stage) <= rest_x + 6.0
        aligned = (max(xs) - min(xs)) <= 8.0
        ordered = all(stage[i].y2 <= stage[i + 1].y + 2.0 for i in range(3))
        hit = vertical and rounded and leftmost and aligned and ordered
        stage_detail = ("圆角框%s，文字纵向%s，位于最左%s，左边缘偏差 %.1f 磅，"
                        "自上而下不重叠%s"
                        % ("是" if rounded else "否", "是" if vertical else "否",
                           "是" if leftmost else "否", max(xs) - min(xs),
                           "是" if ordered else "否"))
    add(hit, "左侧“导入”“核验”“研判”“形成”四个阶段标签齐全，各为独立圆角框、文字纵向排布并纵向对齐其所辖层级",
        stage_detail)

    # 11 首层分四路 + 四路汇聚接入
    root_node = named.get("研究导入：明确目标与范围")
    hit = False
    detail11 = "首层框或第二层框缺失"
    if root_node is not None and len(row2) == 4 and need_node is not None:
        into_row2 = sum(1 for s in row2 if _arrows_into_top(conns, s))
        span = max(s.x2 for s in row2) - min(s.x for s in row2)
        split = [c for c in conns if c.h <= 4.0 and c.w >= 0.45 * span
                 and root_node.y2 - 6.0 <= c.cy <= min(s.y for s in row2) + 6.0]
        merge = [c for c in conns if c.h <= 4.0 and c.w >= 0.45 * span
                 and max(s.y2 for s in row2) - 6.0 <= c.cy <= need_node.y + 6.0]
        into_need = _arrows_into_top(conns, need_node)
        hit = into_row2 == 4 and bool(split) and bool(merge) and bool(into_need)
        detail11 = ("落在第二层四框顶部的箭头 %d 条，首层分叉线 %d 条，"
                    "汇聚线 %d 条，接入“需求要点识别与目标分解”箭头 %d 条"
                    % (into_row2, len(split), len(merge), len(into_need)))
    add(hit, "“研究导入：明确目标与范围”向下分四路箭头分别接第二层四框，"
             "四路再经汇聚线收拢向下接入“需求要点识别与目标分解”", detail11)
    # 12 判定分支闭环 + 不通过/通过标注
    fix_node = named.get("完善资料与方案")
    flow_node = named.get("业务流程梳理")
    right_arrow = None
    loop_conn = None
    back_arrow = False
    if conf_node is not None and fix_node is not None:
        for conn in conns:
            if conn.has_arrow and _touches_edge(conn, conf_node, 'right') \
                    and _touches_edge(conn, fix_node, 'left'):
                right_arrow = conn
                break
    if fix_node is not None and flow_node is not None:
        edge = min(fix_node.x2, flow_node.x2)
        for conn in conns:
            if conn.x >= edge - 12.0 and conn.h >= 40.0 and \
                    conn.y - 8.0 <= flow_node.cy and conn.y2 + 8.0 >= fix_node.cy:
                loop_conn = conn
                break
        back_arrow = any(conn.has_arrow and
                         (_touches_edge(conn, flow_node, 'right') or
                          _touches_edge(conn, flow_node, 'top'))
                         for conn in conns)
    branch_labels = [lbl for lbl in BRANCH_LABELS if _norm(lbl) in full_text]
    over_node = []
    over_conn = []
    for lbl in BRANCH_LABELS:
        mark = _find_by_text(text_nodes, lbl, exact_only=True)
        if mark is None:
            continue
        if any(_overlap_area(mark, other) > 0 for other in nodes if other is not mark):
            over_node.append(lbl)
        if any(_overlap_area(mark, conn) > 0 for conn in conns):
            over_conn.append(lbl)
    label_clean = not over_node and not over_conn
    hit = (right_arrow is not None and loop_conn is not None and back_arrow
           and len(branch_labels) == 2 and label_clean)
    add(hit, "“可行性确认”菱形右侧以箭头接“完善资料与方案”并经折线回流接回“业务流程梳理”形成闭环，"
             "两条分支旁保留“不通过”“通过”标注且不压连线与框体",
        "右向箭头%s，回流折线%s，回接箭头%s，分支标注 %s，标注压盖%s"
        % ("有" if right_arrow is not None else "无",
           "有" if loop_conn is not None else "无",
           "有" if back_arrow else "无",
           "、".join(branch_labels) if branch_labels else "缺失",
           "无" if label_clean else "有（压框体：%s；压连线：%s）"
           % ("、".join(over_node) or "无", "、".join(over_conn) or "无")))
    # 13 运行环境综合研判分两路 + 3/4 路分叉
    env_node = named.get("运行环境综合研判")
    outer_node = named.get("外部条件研判")
    inner_node = named.get("内部能力评阅")
    hit = False
    detail13 = "关键节点缺失"
    if env_node is not None and outer_node is not None and inner_node is not None:
        to_outer = _arrows_into_top(conns, outer_node)
        to_inner = _arrows_into_top(conns, inner_node)
        outer_boxes = [named[l] for l in NARROW_OUTER if l in named]
        inner_boxes = [named[l] for l in NARROW_INNER if l in named]
        n_outer = sum(1 for s in outer_boxes if _arrows_into_top(conns, s))
        n_inner = sum(1 for s in inner_boxes if _arrows_into_top(conns, s))
        hit = bool(to_outer) and bool(to_inner) and n_outer == 3 and n_inner == 4
        detail13 = ("接入“外部条件研判”%d 条、“内部能力评阅”%d 条；"
                    "下发落在窄条框顶部的箭头 %d/3 条与 %d/4 条"
                    % (len(to_outer), len(to_inner), n_outer, n_inner))
    add(hit, "“运行环境综合研判”向下分两路接“外部条件研判”“内部能力评阅”，"
             "并分别下发 3 条与 4 条箭头落在对应窄条框顶部", detail13)

    # 14 EFE/IFE 汇入适配矩阵 + 结尾纵向串接
    efe_node = named.get("EFE评阅清单")
    ife_node = named.get("IFE评阅清单")
    matrix = named.get("适配矩阵")
    hit = False
    detail14 = "关键节点缺失"
    if efe_node is not None and ife_node is not None and matrix is not None:
        efe_in = any(c.has_arrow and _touches_edge(c, efe_node, 'right')
                     and _touches_edge(c, matrix, 'left') for c in conns)
        ife_in = any(c.has_arrow and _touches_edge(c, ife_node, 'left')
                     and _touches_edge(c, matrix, 'right') for c in conns)
        chain = ["适配矩阵", "实施路径定位与方案设计",
                 "组织保障与质量闭环", "归纳与展望"]
        broken = []
        for i in range(len(chain) - 1):
            upper, lower = named.get(chain[i]), named.get(chain[i + 1])
            if upper is None or lower is None or \
                    _conn_down_between(conns, upper, lower) is None:
                broken.append("%s→%s" % (chain[i], chain[i + 1]))
        hit = efe_in and ife_in and not broken
        detail14 = ("EFE 右向箭头%s，IFE 左向箭头%s，纵向串接缺失段：%s"
                    % ("有" if efe_in else "无", "有" if ife_in else "无",
                       "、".join(broken) if broken else "无"))
    add(hit, "“EFE评阅清单”向右、“IFE评阅清单”向左箭头均指向“适配矩阵”，"
             "并由“适配矩阵”依次纵向串接至“归纳与展望”，各段箭头齐全", detail14)
    # 15 节点文字无错字漏字
    missing = [lbl for lbl in ALL_NODE_LABELS if lbl not in named]
    hit = not missing
    add(hit, "节点文字无错字漏字，各节点表述完整，无改字、缺字、同音错字",
        "缺失或写错的节点文字：%s" % ("、".join(missing) if missing else "无"))

    # 16 A4 纵向 + 单页 + 横向居中 + 无截断
    a4_ok = (abs(ctx['page_w'] - 595.3) <= 10.0 and
             abs(ctx['page_h'] - 841.9) <= 10.0 and ctx['page_h'] > ctx['page_w'])
    all_shapes = model['shapes'] or nodes
    if all_shapes:
        min_x = min(s.x for s in all_shapes)
        max_x = max(s.x2 for s in all_shapes)
        min_y = min(s.y for s in all_shapes)
        max_y = max(s.y2 for s in all_shapes)
    else:  # 文档内无任何形状（如流程图仍以整张图片承载）
        min_x = max_x = min_y = max_y = 0.0
    inside = bool(all_shapes) and (min_x >= -1.0 and max_x <= ctx['page_w'] + 1.0 and
                                   min_y >= -1.0 and max_y <= ctx['page_h'] + 1.0)
    center_off = abs((min_x + max_x) / 2.0 - body_center)
    single = (max_y - min_y) <= ctx['body_h'] and max_y <= ctx['page_h'] - 5.0
    page_note = "按几何估算%s单页" % ("为" if single else "非")
    hit = a4_ok and inside and center_off <= 20.0 and single
    add(hit, "页面为纵向 A4，整幅流程图在 1 页内完整呈现并横向居中于版心，不跨页、无节点或连线被页边界截断",
        "页面 %.0f×%.0f 磅，%s，横向中心偏差 %.1f 磅，图幅 x∈[%.1f, %.1f] y∈[%.1f, %.1f]"
        % (ctx['page_w'], ctx['page_h'], page_note, center_off,
           min_x, max_x, min_y, max_y))
    # 17 同层等宽等高对齐 + 主干中轴对齐
    layer_bad = []
    layer_groups = (
        ("第二层四框", row2),
        ("窄条层", narrow),
        ("矩阵层", [named[l] for l in ("EFE评阅清单", "适配矩阵", "IFE评阅清单")
                 if l in named]),
        ("研判两框", [named[l] for l in ("外部条件研判", "内部能力评阅") if l in named]),
    )
    for name, group in layer_groups:
        if len(group) < 2:
            continue
        if max(s.w for s in group) - min(s.w for s in group) > 1.0:
            layer_bad.append("%s宽度不齐" % name)
        if max(s.h for s in group) - min(s.h for s in group) > 1.0:
            layer_bad.append("%s高度不齐" % name)
        spread = max(s.y for s in group) - min(s.y for s in group)
        if spread > POS_TOL:
            layer_bad.append("%s顶边偏差 %.1f 磅" % (name, spread))
    trunk_spread = None
    if len(trunk) >= 4:
        cxs = [s.cx for s in trunk]
        trunk_spread = max(cxs) - min(cxs)
    hit = (not layer_bad and trunk_spread is not None
           and trunk_spread <= TRUNK_TOL)
    add(hit, "同一层级各框等宽等高、顶边纵坐标偏差≤3 磅；纵向主干各节点居中于同一中轴、中心横坐标偏差≤15 磅",
        "同层问题：%s；主干中轴偏差 %s"
        % ("、".join(layer_bad) if layer_bad else "无",
           "%.1f 磅" % trunk_spread if trunk_spread is not None else "主干节点不足"))

    # 18 连线端点与节点边框相接
    dangling = []
    crossing = []
    counted = []
    for conn in conns:
        if conn.y2 <= top_ref + 2.0:
            continue  # 标题下方的装饰分隔线不计入连线
        counted.append(conn)
        ends = _compute_endpoints(conn)
        for point in ends:
            if any(_point_to_rect_dist(point, n) <= TOUCH_TOL for n in nodes):
                continue
            linked = False
            for other in conns:
                if other is conn:
                    continue
                seg = _compute_endpoints(other)
                if _point_to_seg_dist(point, seg[0], seg[1]) <= TOUCH_TOL:
                    linked = True
                    break
            if not linked:
                dangling.append(conn)
                break
        for node in nodes:
            if any(_point_to_rect_dist(p, node) <= TOUCH_TOL for p in ends):
                continue
            over_w = min(conn.x2, node.x2) - max(conn.x, node.x)
            over_h = min(conn.y2, node.y2) - max(conn.y, node.y)
            if over_w > 4.0 and over_h > 4.0:
                crossing.append(conn)
                break
    # 文档内没有连线对象时同样不给分，避免“零违规”被判为达标
    hit = bool(counted) and bool(nodes) and not dangling and not crossing
    add(hit, "连线端点与节点边框相接，不悬空在框外空白处、不穿越框体、不压住框内文字",
        "参与判定的连线 %d 条，悬空连线 %d 条，穿越框体连线 %d 条"
        % (len(counted), len(dangling), len(crossing)))
    # 19 样式统一（填充/边框/箭头同色/字号）
    main_nodes = [n for n in text_nodes if n not in stage]
    white_bad = [_norm(n.text)[:8] for n in main_nodes if not _is_near_white(n.fill)]
    widths = [n.line_pt for n in main_nodes if n.line_pt]
    thin_ok = bool(widths) and max(widths) <= 2.0
    stage_ok = len(stage) == 4 and all(
        _is_light_gray(s.fill) and s.geom.startswith("roundRect") for s in stage)
    conn_colors = [_conn_color(c) for c in flow_conns]
    conn_colors = [c for c in conn_colors if c]
    color_set = set(conn_colors)
    color_ok = len(color_set) == 1 and _is_greenish(conn_colors[0])
    conn_widths = {round(c.line_pt, 2) for c in flow_conns
                   if c.line_pt and c.geom not in ARROW_GEOMS}
    width_ok = len(conn_widths) <= 1
    fonts = [n.font_pt for n in text_nodes]
    font_ok = bool(fonts) and min(fonts) >= 8.0
    layer_font_ok = all(len({s.font_pt for s in g}) <= 1
                        for _n, g in layer_groups if g)
    hit = (not white_bad and thin_ok and stage_ok and color_ok and width_ok
           and font_ok and layer_font_ok)
    add(hit, "样式统一——节点为白色填充配细边框，连线与箭头同为绿色系且线宽一致，"
             "四个阶段标签为浅灰填充圆角框，同层级节点字号一致且全部节点文字≥8 磅",
        "非白填充节点 %d 个，边框最大 %.2f 磅，阶段标签浅灰圆角%s，"
        "连线颜色 %d 种（绿色系%s），连线线宽 %d 种，最小字号 %.1f 磅，同层字号%s"
        % (len(white_bad), max(widths) if widths else 0.0,
           "是" if stage_ok else "否", len(color_set),
           "是" if (conn_colors and _is_greenish(conn_colors[0])) else "否",
           len(conn_widths), min(fonts) if fonts else 0.0,
           "一致" if layer_font_ok else "不一致"))

    # 20 相邻节点框留有清晰间距
    # “不通过”“通过”是分支标注而非节点框，其压盖问题由第 12 条判定，此处排除
    mark_ids = {id(_find_by_text(text_nodes, lbl, exact_only=True))
                for lbl in BRANCH_LABELS}
    boxes = [n for n in text_nodes
             if n.w > 1.0 and n.h > 1.0 and id(n) not in mark_ids]
    min_gap = None
    worst_pair = ""
    for i, one in enumerate(boxes):
        for two in boxes[i + 1:]:
            gap = _rect_gap(one, two)
            if min_gap is None or gap < min_gap:
                min_gap = gap
                worst_pair = "%s/%s" % (_norm(one.text)[:6], _norm(two.text)[:6])
    hit = min_gap is not None and min_gap >= 3.0
    add(hit, "相邻节点框之间留有清晰间距，不粘连、不重叠、不共用边框线",
        "最小净间距 %.1f 磅（%s）" % (min_gap if min_gap is not None else -1.0,
                                worst_pair))
    return _finalize(result, results)


def _finalize(result, results):
    """把 results 里的 "+N：..." 反解为结构化 dim2_items 并汇总总分。"""
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
        hit = delta > 0
        dim2_items.append({
            "rule": rule,
            "max_delta": max_delta,
            "delta": max_delta if hit else 0,
            "hit": hit,
        })
        total_score += (max_delta if hit else 0)
    result["dim2_items"] = dim2_items
    result["total_score"] = total_score
    return result


def evaluate(dir_path):
    """入参为脚本所在目录的路径，自行在该目录内定位并打开被评估的 .docx。"""
    result = {
        "id": SCRIPT_ID,
        "file_name": "",
        "status": "ok",
        "error": "",
        "dim1_pass": False,
        "dim1_reason": "",
        "dim2_items": [],
        "total_score": 0,
        "max_score": MAX_SCORE_TOTAL,
    }
    try:
        file_path, _doc = _open_word_document(dir_path)
        result["file_name"] = os.path.basename(file_path)
        _evaluate_impl(file_path, result)
    except DocxNotFoundError as exc:
        # 目录中没有待评估文件：摘要只写 error，两个维度一字段留空，满分记 0
        result["status"] = "error"
        result["error"] = str(exc)
        result["dim1_pass"] = ""
        result["dim1_reason"] = ""
        result["dim2_items"] = []
        result["total_score"] = 0
        result["max_score"] = 0
    except Exception as exc:  # 文档包损坏按维度一失败处理；脚本异常才标记 error
        message = "%s: %s" % (type(exc).__name__, exc)
        text = str(exc)
        damaged = (
            isinstance(exc, KeyError)
            and ("There is no item named" in text
                 or "no relationship of type" in text)
        ) or type(exc).__name__ in (
            "XMLSyntaxError", "ParseError", "BadZipFile", "InvalidXmlError"
        )
        result["status"] = "ok" if damaged else "error"
        result["error"] = None if damaged else message
        result["dim1_pass"] = False
        result["dim1_reason"] = message
        result["dim2_items"] = []
        result["total_score"] = 0
    return result


if __name__ == "__main__":
    _target = sys.argv[1] if len(sys.argv) > 1 else os.path.dirname(
        os.path.abspath(__file__))
    _out = evaluate(_target)
    print(json.dumps(_out, ensure_ascii=False, indent=2))

