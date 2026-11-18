#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# pyright: reportUnknownMemberType=false, reportUnknownVariableType=false, reportUnknownArgumentType=false, reportMissingImports=false, reportMissingModuleSource=false, reportAttributeAccessIssue=false, reportOptionalMemberAccess=false, reportIndexIssue=false, reportGeneralTypeIssues=false

# 按接口统一约定，本模块不修改全局 sys.stdout，也不 print 主结果。主结果统一由
# evaluate(dir_path) -> dict 返回；入参为“脚本所在目录的路径”，即存放待评测 .pptx
# 的目录。本文件仅在 __main__ 下作为本地调试入口，把返回值以 JSON 形式打印出来。
#
# 返回字段：id / file_name / status / error / dim1_pass / dim1_reason /
#           dim2_items（每项只含 rule / max_delta / delta / hit）/
#           total_score / max_score
# status='error' 表示这份提交没被真正评过（目录不存在、目录内无 .pptx、脚本自身
# 异常），此时 dim1_pass 为空串（维度一未判定）、dim2_items 为空表、total_score
# 与 max_score 一律为 0；dim1_pass 为 False 时 dim1_reason 必须给出错误摘要，
# 通过时为空串。
#
# 本脚本对应任务：把一张流程图复刻成“一模一样并且可编辑”的单页 PPT。
# 维度1 为可用与可修改性硬门槛，不满足则完成度总分记 0；
# 维度2 为 38 条完成度加分项，满分 82 分，逐条严格对齐《评测细则.txt》。
#
# 判定所依赖的三类事实来源：
#   ① OOXML 几何与样式：p:spTree 递归展开（含 a:chOff/a:chExt 组内坐标变换），
#      取 a:off/a:ext（保留有符号 cx/cy 以还原箭头指向）、a:prstGeom@prst、
#      a:ln（w/solidFill/prstDash/headEnd/tailEnd）、a:solidFill、a:xfrm@flipH/@flipV；
#   ② 主题色解析：p:clrMap → theme1.xml/a:clrScheme，把 schemeClr（tx1/bg2/accent1…）
#      还原成 srgb，否则 scheme 引用色会被误判为“无颜色”；
#   ③ 折行行数：按 AFM 字宽表估算——可用排版宽度 = 文本框宽度 - 左右内边距，
#      逐段累加字宽折行，留 3% 宽度容差。
#
# 与办公软件对齐：EMU→cm 用 360000 EMU/cm；线宽 EMU→磅用 /12700；字号
# 磅→cm 用 ×2.54/72。位置一律换算成“占页面宽/高的百分比”后再与细则区间比对，
# 使不同页面尺寸的交付文件可以用同一套锚点判定。

import os
import re
import sys
import json
import math
import zipfile
from xml.etree import ElementTree as ET

SCRIPT_ID = "051"
MAX_SCORE_TOTAL = 82

_MAX_DELTAS_ORDER = [
    1,  # 01 左侧极浅青色圆角矩形
    1,  # 02 Multivariate / Time Series 两行标签
    3,  # 03 折线图竖直网格线
    5,  # 04 蓝橙紫绿四条连贯数据曲线
    1,  # 05 青绿色圆角矩形（Sequence Encoder 容器）
    1,  # 06 “1. Sequence Encoder”标题
    3,  # 07 上下并列两个 GRU 圆角矩形
    1,  # 08 紫色圆角矩形（Pattern Encoder 容器）
    1,  # 09 “2. Pattern Encoder”标题
    3,  # 10 紫色系小方块矩阵
    1,  # 11 Temporal / Embedding 两行标签
    1,  # 12 橙色圆角矩形（Fusion 容器）
    1,  # 13 “3. Fusion & Representation Learning”标题
    3,  # 14 橙色容器内横向并排四个子矩形及其标签
    3,  # 15 Concatenate 网格状表格图形
    3,  # 16 注意力示意图形
    3,  # 17 神经网络示意图形
    3,  # 18 曲线示意图形
    1,  # 19 绿色圆角矩形（Anomaly Scoring 容器）
    1,  # 20 “4. Anomaly Scoring”标题
    3,  # 21 Reconstruction Module 框及框内四个小矩形
    1,  # 22 Anomaly Score 框及单行文本
    3,  # 23 概率分布钟形曲线示意图形
    1,  # 24 粉红色圆角矩形（Decision 容器）
    1,  # 25 “5. Decision”标题
    3,  # 26 Thresholding / Normal-Anomaly 两框及箭头
    1,  # 27 极浅蓝色圆角矩形（Training Objective 容器）
    1,  # 28 “6. Training Objective”标题
    3,  # 29 左右并排两个损失项圆角矩形
    3,  # 30 Lrec / Lreg 下标数学符号及“+”
    1,  # 31 损失区水平分隔单实线
    3,  # 32 “Total Loss: L = Lrec + λLreg”汇总公式
    3,  # 33 底部图例框及其七个内部元素
    5,  # 34 七处黑色带箭头连接线且走向正确
    3,  # 35 紫色与蓝色两条虚线折线箭头
    1,  # 36 橙色框至蓝色框的黑色折线箭头
    1,  # 37 六处区块标题单行完整显示且不压边线
    5,  # 38 虚线边框由线型实现、顶层形状数量与成组情况
]
assert sum(_MAX_DELTAS_ORDER) == MAX_SCORE_TOTAL

A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

EMU_PER_CM = 360000.0
EMU_PER_PT = 12700.0
PT_TO_CM = 2.54 / 72.0

_LINE_PRST = frozenset([
    "line", "straightConnector1",
    "bentConnector2", "bentConnector3", "bentConnector4", "bentConnector5",
    "curvedConnector2", "curvedConnector3", "curvedConnector4", "curvedConnector5",
])
_SHAPE_KINDS = ("sp", "pic", "cxnSp", "grpSp", "graphicFrame")
_DASH_PRST = frozenset([
    "dash", "sysDash", "dot", "sysDot", "lgDash", "dashDot",
    "sysDashDot", "lgDashDot", "lgDashDotDot", "sysDashDotDot",
])


def _q(ns, tag):
    return "{%s}%s" % (ns, tag)


def _fnum(v, default=None):
    try:
        return float(v)
    except (TypeError, ValueError):
        return default


def _truthy(v):
    return v in ("1", "true", "on", "True", True, 1)


def _rng(v, lo, hi):
    return v is not None and lo <= v <= hi


def _hex2rgb(c):
    """把解析后的颜色字符串还原成 (r, g, b)；非具体色值（none/grad/blip）返回 None。"""
    if not c or not isinstance(c, str):
        return None
    s = c.strip().lstrip("#")
    if len(s) != 6:
        return None
    try:
        return (int(s[0:2], 16), int(s[2:4], 16), int(s[4:6], 16))
    except ValueError:
        return None


def _hsv(rgb):
    """返回 (hue 0-360, sat 0-1, val 0-1)；rgb 为 0-255 整数三元组。"""
    r, g, b = (rgb[0] / 255.0, rgb[1] / 255.0, rgb[2] / 255.0)
    mx, mn = max(r, g, b), min(r, g, b)
    d = mx - mn
    if d <= 1e-9:
        h = 0.0
    elif mx == r:
        h = 60.0 * (((g - b) / d) % 6.0)
    elif mx == g:
        h = 60.0 * (((b - r) / d) + 2.0)
    else:
        h = 60.0 * (((r - g) / d) + 4.0)
    return (h, 0.0 if mx <= 1e-9 else d / mx, mx)


def _lum(c):
    rgb = _hex2rgb(c)
    if rgb is None:
        return None
    return 0.299 * rgb[0] + 0.587 * rgb[1] + 0.114 * rgb[2]


def _hue_in(c, lo, hi, smin=0.05):
    rgb = _hex2rgb(c)
    if rgb is None:
        return False
    h, s, _v = _hsv(rgb)
    if s < smin:
        return False
    if lo <= hi:
        return lo <= h <= hi
    return h >= lo or h <= hi


# 色系判定：全部走 HSV 色相区间，而不是逐个枚举十六进制值，
# 这样同一族里深浅不同的实现（如 #0F9F9A 与 #37ADA0）都能被同一条规则接受。
def _is_teal(c):
    return _hue_in(c, 150.0, 200.0, 0.03)


def _is_purple(c):
    return _hue_in(c, 250.0, 300.0, 0.03)


def _is_orange(c):
    return _hue_in(c, 12.0, 48.0, 0.05)


def _is_paleorange(c):
    return _hue_in(c, 8.0, 60.0, 0.008)


def _is_green(c):
    return _hue_in(c, 70.0, 150.0, 0.03)


def _is_pink(c):
    return _hue_in(c, 315.0, 8.0, 0.03)


def _is_blue(c):
    return _hue_in(c, 200.0, 250.0, 0.05)


def _is_brownorange(c):
    rgb = _hex2rgb(c)
    if rgb is None:
        return False
    h, s, v = _hsv(rgb)
    return 8.0 <= h <= 45.0 and s >= 0.30 and v <= 0.72


def _is_lightgray(c):
    rgb = _hex2rgb(c)
    if rgb is None:
        return False
    _h, s, v = _hsv(rgb)
    return s <= 0.25 and 0.45 <= v <= 0.97


def _is_black(c):
    rgb = _hex2rgb(c)
    if rgb is None:
        return False
    return _hsv(rgb)[2] <= 0.25


def _is_darkgrayblue(c):
    rgb = _hex2rgb(c)
    if rgb is None:
        return False
    h, s, v = _hsv(rgb)
    return v <= 0.72 and (s <= 0.25 or 190.0 <= h <= 260.0)


def _min_ch(c):
    rgb = _hex2rgb(c)
    return None if rgb is None else min(rgb)


def _is_verylight(c):
    m = _min_ch(c)
    return m is not None and m >= 228


def _is_light(c):
    m = _min_ch(c)
    return m is not None and m >= 185


def _is_white(c):
    m = _min_ch(c)
    return m is not None and m >= 246


def _near_rgb(c, lo_hex, hi_hex, tol=26):
    """逐通道落在给定两个端点色所围成的区间内（含 tol 容差）。"""
    rgb, lo, hi = _hex2rgb(c), _hex2rgb(lo_hex), _hex2rgb(hi_hex)
    if rgb is None or lo is None or hi is None:
        return False
    for k in range(3):
        a, b = min(lo[k], hi[k]) - tol, max(lo[k], hi[k]) + tol
        if not (a <= rgb[k] <= b):
            return False
    return True


def _nat_key(name):
    return [int(t) if t.isdigit() else t for t in re.split(r"(\d+)", name)]


def _theme_lookup(zf):
    """构造 schemeClr 名 → srgb 的映射：p:clrMap 先把 tx1/bg2 等映射到
    dk1/lt2 等主题槽位，再到 theme1.xml 的 a:clrScheme 里取具体色值。"""
    names = zf.namelist()
    cmap = {}
    masters = sorted(
        [n for n in names if re.match(r"ppt/slideMasters/slideMaster\d+\.xml$", n)],
        key=_nat_key)
    if masters:
        try:
            root = ET.fromstring(zf.read(masters[0]))
            node = root.find(_q(P_NS, "clrMap"))
            if node is not None:
                cmap = dict(node.attrib)
        except ET.ParseError:
            cmap = {}
    scheme = {}
    themes = sorted(
        [n for n in names if re.match(r"ppt/theme/theme\d+\.xml$", n)], key=_nat_key)
    if themes:
        try:
            root = ET.fromstring(zf.read(themes[0]))
            cs = root.find(".//" + _q(A_NS, "clrScheme"))
            for child in (cs if cs is not None else []):
                slot = child.tag.split("}")[-1]
                srgb = child.find(_q(A_NS, "srgbClr"))
                sysc = child.find(_q(A_NS, "sysClr"))
                if srgb is not None:
                    scheme[slot] = srgb.get("val")
                elif sysc is not None:
                    scheme[slot] = sysc.get("lastClr") or sysc.get("val")
        except ET.ParseError:
            scheme = {}
    out = {}
    for k in ("bg1", "tx1", "bg2", "tx2", "accent1", "accent2", "accent3",
              "accent4", "accent5", "accent6", "hlink", "folHlink"):
        out[k] = scheme.get(cmap.get(k, k))
    for k, v in scheme.items():
        out.setdefault(k, v)
    out.setdefault("phClr", None)
    return out


def _first_slide_xml(zf):
    names = sorted([n for n in zf.namelist()
                    if re.match(r"ppt/slides/slide\d+\.xml$", n)], key=_nat_key)
    if not names:
        raise ValueError("包内未发现 ppt/slides/slideN.xml")
    return zf.read(names[0]), len(names)


def _slide_size(zf):
    try:
        root = ET.fromstring(zf.read("ppt/presentation.xml"))
    except (KeyError, ET.ParseError):
        return (33.87, 19.05, 1)
    sz = root.find(_q(P_NS, "sldSz"))
    cx = _fnum(sz.get("cx"), 12192000.0) if sz is not None else 12192000.0
    cy = _fnum(sz.get("cy"), 6858000.0) if sz is not None else 6858000.0
    lst = root.find(_q(P_NS, "sldIdLst"))
    n = len(list(lst)) if lst is not None else 0
    return (cx / EMU_PER_CM, cy / EMU_PER_CM, n)


def _color_of(node, theme):
    """从一个填充/线条容器节点里取颜色，schemeClr 用主题表还原成 srgb。"""
    if node is None:
        return None
    srgb = node.find(_q(A_NS, "srgbClr"))
    if srgb is not None:
        return srgb.get("val")
    sch = node.find(_q(A_NS, "schemeClr"))
    if sch is not None:
        return theme.get(sch.get("val"))
    sysc = node.find(_q(A_NS, "sysClr"))
    if sysc is not None:
        return sysc.get("lastClr") or sysc.get("val")
    return None


def _own_xfrm(el):
    for name in ("spPr", "grpSpPr", "xfrm"):
        child = el.find(_q(P_NS, name))
        if child is not None:
            if child.tag == _q(A_NS, "xfrm"):
                return child
            return child.find(_q(A_NS, "xfrm"))
    return None


def _off_ext(xf):
    if xf is None:
        return None
    off, ext = xf.find(_q(A_NS, "off")), xf.find(_q(A_NS, "ext"))
    if off is None or ext is None:
        return None
    return (_fnum(off.get("x"), 0.0), _fnum(off.get("y"), 0.0),
            _fnum(ext.get("cx"), 0.0), _fnum(ext.get("cy"), 0.0))


def _fill_of(el, theme):
    spPr = el.find(_q(P_NS, "spPr"))
    if spPr is None:
        return None
    if spPr.find(_q(A_NS, "noFill")) is not None:
        return "none"
    sf = spPr.find(_q(A_NS, "solidFill"))
    if sf is not None:
        return _color_of(sf, theme)
    if spPr.find(_q(A_NS, "gradFill")) is not None:
        return "grad"
    if spPr.find(_q(A_NS, "blipFill")) is not None:
        return "blip"
    return None


def _line_of(el, theme):
    out = {"color": None, "w": None, "dash": None,
           "head": None, "tail": None, "nofill": False}
    spPr = el.find(_q(P_NS, "spPr"))
    if spPr is None:
        return out
    ln = spPr.find(_q(A_NS, "ln"))
    if ln is None:
        return out
    wv = _fnum(ln.get("w"))
    out["w"] = None if wv is None else wv / EMU_PER_PT
    out["nofill"] = ln.find(_q(A_NS, "noFill")) is not None
    out["color"] = _color_of(ln.find(_q(A_NS, "solidFill")), theme)
    dash = ln.find(_q(A_NS, "prstDash"))
    out["dash"] = dash.get("val") if dash is not None else None
    head, tail = ln.find(_q(A_NS, "headEnd")), ln.find(_q(A_NS, "tailEnd"))
    out["head"] = head.get("type") if head is not None else None
    out["tail"] = tail.get("type") if tail is not None else None
    return out


def _text_of(el, theme):
    tx = el.find(_q(P_NS, "txBody"))
    if tx is None:
        return None
    paras = []
    for p in tx.findall(_q(A_NS, "p")):
        runs = []
        for ch in p:
            if ch.tag == _q(A_NS, "r"):
                t = ch.find(_q(A_NS, "t"))
                rPr = ch.find(_q(A_NS, "rPr"))
                d = {"t": t.text if t is not None and t.text else "",
                     "sz": None, "b": None, "i": None, "base": None,
                     "color": None, "font": None}
                if rPr is not None:
                    sz = _fnum(rPr.get("sz"))
                    d["sz"] = None if sz is None else sz / 100.0
                    d["b"] = rPr.get("b")
                    d["i"] = rPr.get("i")
                    d["base"] = rPr.get("baseline")
                    d["color"] = _color_of(rPr.find(_q(A_NS, "solidFill")), theme)
                    lat = rPr.find(_q(A_NS, "latin"))
                    d["font"] = lat.get("typeface") if lat is not None else None
                runs.append(d)
            elif ch.tag == _q(A_NS, "br"):
                runs.append({"t": "\n", "sz": None, "b": None, "i": None,
                             "base": None, "color": None, "font": None})
        pPr = p.find(_q(A_NS, "pPr"))
        paras.append({"algn": pPr.get("algn") if pPr is not None else None,
                      "runs": runs})
    bodyPr = tx.find(_q(A_NS, "bodyPr"))
    body = {"lIns": None, "rIns": None, "tIns": None, "bIns": None,
            "wrap": None, "anchor": None}
    if bodyPr is not None:
        for k in ("lIns", "rIns", "tIns", "bIns", "wrap", "anchor"):
            body[k] = bodyPr.get(k)
    return {"paras": paras, "body": body}


def _nv_name(el):
    for tag in ("nvSpPr", "nvGrpSpPr", "nvCxnSpPr", "nvPicPr", "nvGraphicFramePr"):
        nv = el.find(_q(P_NS, tag))
        if nv is None:
            continue
        cnv = nv.find(_q(P_NS, "cNvPr"))
        if cnv is not None:
            return cnv.get("name") or ""
    return ""


def _walk(spTree, theme):
    """递归展开 spTree，返回展平后的形状列表。

    组内子形状必须先经过 a:chOff/a:chExt 的坐标变换才能得到页面绝对坐标，
    否则组合内的图标、连接线位置会全部错位。sw/sh 保留 a:ext 的原始符号，
    箭头指向依赖它——一旦提前把负宽高规范化，就再也分不清箭头在头还是在尾。
    """
    out = []

    def rec(container, tf, top_ref, depth):
        for child in container:
            kind = child.tag.split("}")[-1]
            if kind not in _SHAPE_KINDS:
                continue
            xf = _own_xfrm(child)
            oe = _off_ext(xf)
            ax, bx, ay, by = tf
            if oe is None:
                x = y = w = h = sw = sh = None
            else:
                X, Y, CX, CY = oe
                x = (ax + X * bx) / EMU_PER_CM
                y = (ay + Y * by) / EMU_PER_CM
                w = (CX * bx) / EMU_PER_CM
                h = (CY * by) / EMU_PER_CM
                sw, sh = w, h
                if w < 0:
                    x += w
                    w = -w
                if h < 0:
                    y += h
                    h = -h
            idx = len(out)
            top = idx if depth == 0 else top_ref
            spPr = child.find(_q(P_NS, "spPr"))
            prst = None
            if spPr is not None:
                geom = spPr.find(_q(A_NS, "prstGeom"))
                if geom is not None:
                    prst = geom.get("prst")
                elif spPr.find(_q(A_NS, "custGeom")) is not None:
                    prst = "custGeom"
            out.append({
                "i": idx, "top": top, "depth": depth, "kind": kind,
                "name": _nv_name(child), "prst": prst,
                "x": x, "y": y, "w": w, "h": h, "sw": sw, "sh": sh,
                "fill": _fill_of(child, theme), "ln": _line_of(child, theme),
                "flipH": _truthy(xf.get("flipH")) if xf is not None else False,
                "flipV": _truthy(xf.get("flipV")) if xf is not None else False,
                "rot": _fnum(xf.get("rot")) if xf is not None else None,
                "tx": _text_of(child, theme),
            })
            if kind == "grpSp":
                ntf = tf
                if xf is not None and w:
                    cho, che = xf.find(_q(A_NS, "chOff")), xf.find(_q(A_NS, "chExt"))
                    if cho is not None and che is not None:
                        ccx = _fnum(che.get("cx"), 1.0) or 1.0
                        ccy = _fnum(che.get("cy"), 1.0) or 1.0
                        kx = (w * EMU_PER_CM) / ccx
                        ky = (h * EMU_PER_CM) / ccy
                        ntf = (x * EMU_PER_CM - _fnum(cho.get("x"), 0.0) * kx, kx,
                               y * EMU_PER_CM - _fnum(cho.get("y"), 0.0) * ky, ky)
                rec(child, ntf, top, depth + 1)

    rec(spTree, (0.0, 1.0, 0.0, 1.0), -1, 0)
    return out


def _load(file_path):
    with zipfile.ZipFile(file_path) as zf:
        pw, ph, n_slides = _slide_size(zf)
        theme = _theme_lookup(zf)
        blob, n_xml = _first_slide_xml(zf)
    root = ET.fromstring(blob)
    spTree = root.find(".//" + _q(P_NS, "spTree"))
    if spTree is None:
        raise ValueError("首页幻灯片缺少 p:spTree")
    return {"pw": pw, "ph": ph,
            "n_slides": max(n_slides, n_xml),
            "shapes": _walk(spTree, theme), "theme": theme}


# ---------------------------------------------------------------- 文本工具


def _flat(tx):
    if not tx:
        return ""
    return "\n".join("".join(r.get("t") or "" for r in p["runs"])
                     for p in tx["paras"])


def _paras_text(tx):
    if not tx:
        return []
    return [re.sub(r"\s+", " ", "".join(r.get("t") or "" for r in p["runs"])).strip()
            for p in tx["paras"]]


def _nonempty_paras(tx):
    return [t for t in _paras_text(tx) if t]


def _norm(s):
    return re.sub(r"\s+", " ", (s or "")).strip()


def _squash(s):
    return re.sub(r"\s+", "", (s or ""))


def _runs(tx):
    if not tx:
        return []
    return [r for p in tx["paras"] for r in p["runs"]]


def _text_runs(tx):
    return [r for r in _runs(tx) if (r.get("t") or "").strip()]


def _sizes(tx):
    return [r["sz"] for r in _text_runs(tx) if r.get("sz")]


def _dom_sz(tx, default=18.0):
    """取形状里“字数最多的字号”，避免被行首占位空格的小字号带偏。"""
    tally = {}
    for r in _text_runs(tx):
        if r.get("sz"):
            tally[r["sz"]] = tally.get(r["sz"], 0) + len((r.get("t") or "").strip())
    if not tally:
        return default
    return sorted(tally.items(), key=lambda kv: (-kv[1], -kv[0]))[0][0]


def _all_bold(tx):
    rs = _text_runs(tx)
    return bool(rs) and all(_truthy(r.get("b")) for r in rs)


def _all_italic(tx):
    rs = _text_runs(tx)
    return bool(rs) and all(_truthy(r.get("i")) for r in rs)


def _all_ctr(tx):
    if not tx:
        return False
    ps = [p for p in tx["paras"] if "".join(r.get("t") or "" for r in p["runs"]).strip()]
    return bool(ps) and all(p.get("algn") == "ctr" for p in ps)


def _run_colors(tx):
    return [r.get("color") for r in _text_runs(tx) if r.get("color")]


# ---------------------------------------------------------------- 几何工具


class _Page(object):
    """把绝对 cm 坐标换算成占页面宽/高的百分比，让细则锚点与页面尺寸无关。"""

    def __init__(self, pw, ph):
        self.pw = pw or 33.87
        self.ph = ph or 19.05

    def px(self, v):
        return None if v is None else v * 100.0 / self.pw

    def py(self, v):
        return None if v is None else v * 100.0 / self.ph

    def at(self, s, x1, x2, y1, y2):
        if s.get("x") is None or s.get("y") is None:
            return False
        return (x1 <= self.px(s["x"]) <= x2) and (y1 <= self.py(s["y"]) <= y2)

    def at_pt(self, x, y, x1, x2, y1, y2):
        return (x1 <= self.px(x) <= x2) and (y1 <= self.py(y) <= y2)


def _x2(s):
    return None if s.get("x") is None else s["x"] + s["w"]


def _y2(s):
    return None if s.get("y") is None else s["y"] + s["h"]


def _cx(s):
    return None if s.get("x") is None else s["x"] + s["w"] / 2.0


def _cy(s):
    return None if s.get("y") is None else s["y"] + s["h"] / 2.0


def _is_conn(s):
    return s["kind"] == "cxnSp" or s.get("prst") in _LINE_PRST


def _has_arrow(s, which):
    v = s["ln"].get(which)
    return bool(v) and v != "none"


def _ends(s):
    """还原线条的起点与终点。

    a:ext 的 cx/cy 允许为负，负值表示“从右往左/从下往上”画；再叠加
    flipH/flipV 才是最终朝向。丢掉符号会让所有箭头看起来方向一致。
    """
    if s.get("sw") is None or s.get("x") is None:
        return None
    x1, x2 = (s["x"], s["x"] + s["w"]) if s["sw"] >= 0 else (s["x"] + s["w"], s["x"])
    y1, y2 = (s["y"], s["y"] + s["h"]) if s["sh"] >= 0 else (s["y"] + s["h"], s["y"])
    if s["flipH"]:
        x1, x2 = x2, x1
    if s["flipV"]:
        y1, y2 = y2, y1
    return ((x1, y1), (x2, y2))


def _tip(s):
    """带箭头端的坐标：tailEnd 在终点，headEnd 在起点；无箭头返回 None。"""
    e = _ends(s)
    if e is None:
        return None
    if _has_arrow(s, "tail"):
        return e[1]
    if _has_arrow(s, "head"):
        return e[0]
    return None


def _tail_pt(s):
    e = _ends(s)
    if e is None:
        return None
    if _has_arrow(s, "tail"):
        return e[0]
    if _has_arrow(s, "head"):
        return e[1]
    return None


def _bbox(shapes):
    xs = [s["x"] for s in shapes if s.get("x") is not None]
    if not xs:
        return None
    ys = [s["y"] for s in shapes]
    return (min(xs), min(ys),
            max(_x2(s) for s in shapes) - min(xs),
            max(_y2(s) for s in shapes) - min(ys))


def _box_dist(px, py, s):
    """点到形状外接矩形的距离；点在矩形内返回 0。"""
    dx = max(s["x"] - px, 0.0, px - _x2(s))
    dy = max(s["y"] - py, 0.0, py - _y2(s))
    return math.hypot(dx, dy)


def _clusters(vals, gap):
    """一维聚类：排序后相邻差值超过 gap 就切一刀，返回每簇的成员索引。"""
    order = sorted(range(len(vals)), key=lambda k: vals[k])
    groups, cur = [], []
    for k in order:
        if cur and vals[k] - vals[cur[-1]] > gap:
            groups.append(cur)
            cur = []
        cur.append(k)
    if cur:
        groups.append(cur)
    return groups


def _interval_clusters(spans, gap):
    """按区间 [x1, x2] 聚类：区间间空隙小于 gap 视为同一簇。"""
    order = sorted(range(len(spans)), key=lambda k: spans[k][0])
    groups, cur, right = [], [], None
    for k in order:
        a, b = spans[k]
        if cur and a - right > gap:
            groups.append(cur)
            cur, right = [], None
        cur.append(k)
        right = b if right is None else max(right, b)
    if cur:
        groups.append(cur)
    return groups


def _icon_group(shapes, gap=0.12):
    """从一个取点窗口里切出单个小图标的本体。窗口按百分比给定，边界上难免会把
    相邻图标之间的横向箭头连线一起圈进来；这些连线与图标本体之间存在明显的横向
    空隙，因此按横向区间聚类后取成员最多的一簇，即为图标本体。"""
    live = [s for s in shapes if s.get("x") is not None]
    if not live:
        return []
    groups = _interval_clusters([(s["x"], _x2(s)) for s in live], gap)
    best = max(groups, key=lambda g: (len(g), -min(live[k]["x"] for k in g)))
    return [live[k] for k in best]


def _connected(shapes, tol):
    """判断一组线段是否首尾相接成一条链：端点两两距离不超过 tol 即连通，
    最终必须落在同一个连通分量里。单个形状天然连通。"""
    n = len(shapes)
    if n <= 1:
        return n == 1
    pts = []
    for s in shapes:
        e = _ends(s)
        if e is None:
            return False
        pts.append(e)
    parent = list(range(n))

    def find(a):
        while parent[a] != a:
            parent[a] = parent[parent[a]]
            a = parent[a]
        return a

    for a in range(n):
        for b in range(a + 1, n):
            hit = False
            for pa in pts[a]:
                for pb in pts[b]:
                    if math.hypot(pa[0] - pb[0], pa[1] - pb[1]) <= tol:
                        hit = True
                        break
                if hit:
                    break
            if hit:
                ra, rb = find(a), find(b)
                if ra != rb:
                    parent[ra] = rb
    return len(set(find(a) for a in range(n))) == 1


def _hruns(shapes, ytol=0.06, xgap=0.28):
    """把同一水平线上、彼此间隙很小的多段短线合并成“视觉上的一条线”。

    交付文件可能用一串独立实线短段去模拟虚线箭头，渲染出来与真虚线无异；
    本条只负责还原“看起来是一条多长的线、末端有没有箭头”，是否用线型实现
    在虚线边框那一条里单独判定，避免同一处缺陷被重复计分。
    """
    cand = [s for s in shapes
            if s.get("x") is not None and s["h"] <= ytol and s["w"] > 0.02]
    runs = []
    for grp in _clusters([_cy(s) for s in cand], ytol):
        members = sorted([cand[k] for k in grp], key=lambda s: s["x"])
        cur = [members[0]]
        for s in members[1:]:
            if s["x"] - max(_x2(m) for m in cur) <= xgap:
                cur.append(s)
            else:
                runs.append(cur)
                cur = [s]
        runs.append(cur)
    out = []
    for run in runs:
        x1 = min(s["x"] for s in run)
        x2 = max(_x2(s) for s in run)
        out.append({"x1": x1, "x2": x2, "y": sum(_cy(s) for s in run) / len(run),
                    "span": x2 - x1, "n": len(run), "parts": run,
                    "arrow_right": any(_has_arrow(s, "tail") or _has_arrow(s, "head")
                                       for s in run)})
    return out


def _dash_like(run):
    """一条“看起来是虚线”的水平线：本身线型为虚线，或由三段以上短实线拼出。"""
    if run["n"] >= 3:
        return True
    return any((s["ln"].get("dash") or "") in _DASH_PRST for s in run["parts"])


def _solid_like(run):
    if run["n"] >= 3:
        return False
    return all((s["ln"].get("dash") or "solid") not in _DASH_PRST
               for s in run["parts"])


def _pick(shapes, pred):
    return [s for s in shapes if pred(s)]


def _find_rr(shapes, page, x1, x2, y1, y2, w, h, ln_pred=None,
             fill_pred=None, prsts=("roundRect",)):
    """按“位置百分比 + 宽高区间 + 边线/填充色系”定位一个圆角矩形。"""
    hits = []
    for s in shapes:
        if s.get("prst") not in prsts or s.get("x") is None:
            continue
        if not page.at(s, x1, x2, y1, y2):
            continue
        if not (_rng(s["w"], w[0], w[1]) and _rng(s["h"], h[0], h[1])):
            continue
        if ln_pred is not None and not ln_pred(s["ln"].get("color")):
            continue
        if fill_pred is not None and not fill_pred(s.get("fill")):
            continue
        hits.append(s)
    return hits


_AFM_ARIAL = {
    " ": 278, "!": 278, "\"": 355, "#": 556, "$": 556, "%": 889, "&": 667,
    "'": 191, "(": 333, ")": 333, "*": 389, "+": 584, ",": 278, "-": 333,
    ".": 278, "/": 278, "0": 556, "1": 556, "2": 556, "3": 556, "4": 556,
    "5": 556, "6": 556, "7": 556, "8": 556, "9": 556, ":": 278, ";": 278,
    "<": 584, "=": 584, ">": 584, "?": 556, "@": 1015, "A": 667, "B": 667,
    "C": 722, "D": 722, "E": 667, "F": 611, "G": 778, "H": 722, "I": 278,
    "J": 500, "K": 667, "L": 556, "M": 833, "N": 722, "O": 778, "P": 667,
    "Q": 778, "R": 722, "S": 667, "T": 611, "U": 722, "V": 667, "W": 944,
    "X": 667, "Y": 667, "Z": 611, "[": 278, "\\": 278, "]": 278, "^": 469,
    "_": 556, "`": 333, "a": 556, "b": 556, "c": 500, "d": 556, "e": 556,
    "f": 278, "g": 556, "h": 556, "i": 222, "j": 222, "k": 500, "l": 222,
    "m": 833, "n": 556, "o": 556, "p": 556, "q": 556, "r": 333, "s": 500,
    "t": 278, "u": 556, "v": 500, "w": 722, "x": 500, "y": 500, "z": 500,
    "{": 334, "|": 260, "}": 334, "~": 584, "λ": 500,
}

_AFM_ARIAL_BOLD = {
    " ": 278, "-": 333, ".": 278, "/": 278, ":": 333, "+": 584, "=": 584,
    "0": 556, "1": 556, "2": 556, "3": 556, "4": 556, "5": 556, "6": 556,
    "7": 556, "8": 556, "9": 556, "A": 722, "B": 722, "C": 722, "D": 722,
    "E": 667, "F": 611, "G": 778, "H": 722, "I": 278, "J": 556, "K": 722,
    "L": 611, "M": 833, "N": 722, "O": 778, "P": 667, "Q": 778, "R": 722,
    "S": 667, "T": 611, "U": 722, "V": 667, "W": 944, "X": 667, "Y": 667,
    "Z": 611, "a": 556, "b": 611, "c": 556, "d": 611, "e": 556, "f": 333,
    "g": 611, "h": 611, "i": 278, "j": 278, "k": 556, "l": 278, "m": 889,
    "n": 611, "o": 611, "p": 611, "q": 611, "r": 389, "s": 556, "t": 333,
    "u": 611, "v": 556, "w": 778, "x": 556, "y": 556, "z": 500,
    "λ": 556,
}


def _adv(ch, bold):
    tbl = _AFM_ARIAL_BOLD if bold else _AFM_ARIAL
    if ch in tbl:
        return tbl[ch]
    if ord(ch) > 0x2E80:
        return 1000
    return 556 if bold else 500


# 字宽表按 Arial 标定，其余西文字体按整体字宽相对 Arial 的比例折算。
_FONT_W = {
    "calibri": 0.855, "segoe ui": 0.895, "candara": 0.865, "corbel": 0.880,
    "constantia": 0.930, "cambria": 0.945, "georgia": 1.010,
    "times new roman": 0.900, "garamond": 0.855, "tahoma": 0.975,
    "verdana": 1.090, "trebuchet ms": 0.930, "arial narrow": 0.820,
}


def _font_factor(name):
    if not name:
        return 1.0
    return _FONT_W.get(str(name).strip().lower(), 1.0)


def _est_lines(sh):
    """按 AFM 字宽表估算文本框内每个段落实际排出的行数。

    与办公软件对齐：可用排版宽度 = 文本框宽度 - 左内边距 - 右内边距；
    未显式给出内边距时按默认 91440 EMU（0.254 厘米）计。字宽表以 Arial
    标定，按 run 的西文字体名折算字宽比例；估算与真实排版存在细微差异，
    因此判断是否折行时留 3% 宽度容差。
    """
    tx = sh.get("tx")
    if not tx or sh.get("w") is None:
        return 1
    body = tx.get("body") or {}

    def _ins(key):
        v = body.get(key)
        if v is None:
            return 0.254
        return _fnum(v, 0.0) / EMU_PER_CM

    avail = sh["w"] - _ins("lIns") - _ins("rIns")
    if avail <= 0.05:
        return 1
    avail *= 1.03
    total = 0
    for para in tx["paras"]:
        segs = [""]
        bolds = [False]
        fonts = [None]
        for run in para["runs"]:
            t = run.get("t") or ""
            if t == "\n":
                segs.append("")
                bolds.append(False)
                fonts.append(None)
                continue
            segs[-1] += t
            if _truthy(run.get("b")):
                bolds[-1] = True
            if fonts[-1] is None:
                fonts[-1] = run.get("font")
        for seg, bold, font in zip(segs, bolds, fonts):
            words = seg.split()
            if not words:
                total += 1
                continue
            sz = _dom_sz(tx)
            unit = sz * PT_TO_CM / 1000.0 * _font_factor(font)
            lines = 1
            cur = 0.0
            space = _adv(" ", bold) * unit
            for w in words:
                wid = sum(_adv(c, bold) for c in w) * unit
                if cur <= 0.0:
                    cur = wid
                elif cur + space + wid <= avail + 1e-6:
                    cur += space + wid
                else:
                    lines += 1
                    cur = wid
                if cur > avail + 1e-6:
                    # 单个词本身超过一行可排宽度时，办公软件会把它断开续排
                    extra = int(math.ceil(cur / avail)) - 1
                    lines += extra
                    cur -= extra * avail
            total += lines
    return max(1, total)


def _para_lines(sh):
    """文本框内的段落数（含换行符产生的行），不含空段。"""
    tx = sh.get("tx")
    if not tx:
        return 0
    n = 0
    for para in tx["paras"]:
        txt = "".join(r.get("t") or "" for r in para["runs"])
        if not txt.strip():
            continue
        n += 1 + txt.count("\n")
    return max(1, n)


def _no_wrap(sh):
    """估算行数与文本自身的行数一致，即没有因为框太窄而额外折行。"""
    return _est_lines(sh) == _para_lines(sh)


def _check_dim1(file_path, data):
    """维度1 硬门槛：格式可用、单页、内容可编辑。任一条不满足则完成度记 0 分。"""
    # ① 交付文件必须是 .pptx，并且能被正常解包解析。
    #    解析失败会在 evaluate 里落到 status=error，不会走到这里。
    if not os.path.basename(file_path).lower().endswith(".pptx"):
        return (False, "交付文件不是 .pptx 格式")

    # ② 只允许 1 页幻灯片。
    n = data["n_slides"]
    if n != 1:
        return (False, "交付 PPT 包含 %d 页幻灯片，不是 1 页" % n)

    shapes = data["shapes"]

    # ③ 流程图整体或其中任一区块都不能是图片：既不能有 p:pic 对象，
    #    也不能有用图片填充（a:blipFill）去铺一大块版面的形状。
    pics = [s for s in shapes if s["kind"] == "pic"]
    if pics:
        return (False, "页面存在 %d 个图片对象，流程图或其中区块以图片形式呈现" % len(pics))
    page_area = max(data["pw"] * data["ph"], 1e-6)
    for s in shapes:
        if s.get("fill") != "blip" or s.get("w") is None:
            continue
        if s["w"] * s["h"] >= 0.05 * page_area:
            return (False, "页面存在以图片填充的大面积形状，流程图以图片形式呈现")

    # ④ 文字必须落在可编辑文本框内：页面上要真实存在带 a:t 文本的 p:txBody。
    txt = [s for s in shapes if s.get("tx") and _flat(s["tx"]).strip()]
    if not txt:
        return (False, "页面未发现任何可编辑文本框，文字不可编辑")

    # ⑤ 整页可编辑：顶层不能只有孤零零一个对象，否则等同于整页一块死图。
    tops = [s for s in shapes if s["depth"] == 0]
    if len(tops) < 2:
        return (False, "页面顶层仅 %d 个对象，出现整页不可编辑的情况" % len(tops))

    # ⑥ 图形必须是可编辑的形状或线条：至少要有带预置几何或自定义几何的形状。
    geo = [s for s in shapes if s["kind"] in ("sp", "cxnSp") and s.get("prst")]
    if not geo:
        return (False, "页面未发现可编辑的形状或线条对象")

    return (True, "")


def _evaluate_impl(file_path, result):
    data = _load(file_path)
    page = _Page(data["pw"], data["ph"])
    shapes = data["shapes"]
    tops = [s for s in shapes if s["depth"] == 0]

    ok1, why1 = _check_dim1(file_path, data)
    result["dim1_pass"] = bool(ok1)
    # 约定：dim1_pass 为 False 时 dim1_reason 必须带错误摘要，通过时必须为空串。
    if ok1:
        result["dim1_reason"] = ""
    else:
        result["dim1_reason"] = why1 or "维度一未通过（未给出具体原因）"
    if not ok1:
        result["dim2_items"] = []
        result["total_score"] = 0
        result["max_score"] = MAX_SCORE_TOTAL
        return result

    texts = [s for s in shapes if s.get("tx") and _flat(s["tx"]).strip()]

    def tf(*keys):
        """按“去掉全部空白后依次包含这些片段”找文本形状，兼容多 run 拆词与填充空格。"""
        out = []
        for s in texts:
            k = _squash(_flat(s["tx"])).lower()
            pos, ok = 0, True
            for frag in keys:
                f = _squash(frag).lower()
                idx = k.find(f, pos)
                if idx < 0:
                    ok = False
                    break
                pos = idx + len(f)
            if ok:
                out.append(s)
        return out

    def one(hits):
        return hits[0] if hits else None

    conns = [s for s in shapes if _is_conn(s) and s.get("x") is not None]
    blacks = [s for s in conns if _is_black(s["ln"].get("color"))]

    # 六个区块框、图例框、两个评分子框、四个并排子矩形：多条规则共用的位置基准，
    # 这里用比细则本身稍宽的窗口先定位一次，各条规则再按自己的严格区间复核。
    a_teal = one(_find_rr(shapes, page, 0, 10, 18, 32, (4.2, 5.8), (5.2, 6.8),
                          ln_pred=_is_teal))
    a_seq = one(_find_rr(shapes, page, 16, 30, 2, 14, (5.2, 6.8), (4.0, 5.2),
                         ln_pred=_is_teal))
    a_pat = one(_find_rr(shapes, page, 16, 30, 36, 47, (4.8, 6.2), (4.6, 5.6),
                         ln_pred=_is_purple))
    a_fus = one(_find_rr(shapes, page, 38, 48, 2, 14, (12.5, 15.5), (6.8, 8.0),
                         ln_pred=_is_orange))
    a_ano = one(_find_rr(shapes, page, 80, 90, 8, 17, (4.3, 5.6), (7.0, 8.2),
                         ln_pred=_is_green))
    a_dec = one(_find_rr(shapes, page, 80, 90, 55, 64, (4.2, 5.4), (4.3, 5.2),
                         ln_pred=_is_pink))
    a_trn = one(_find_rr(shapes, page, 40, 48, 53, 62, (7.8, 9.5), (4.3, 5.2),
                         ln_pred=_is_blue))
    a_lgd = one(_find_rr(shapes, page, 20, 29, 86, 94, (17.5, 20.5), (1.2, 1.8),
                         ln_pred=_is_lightgray, fill_pred=_is_white))
    a_rec = one(_find_rr(shapes, page, 82, 91, 12, 20, (3.3, 4.3), (2.4, 3.2),
                         ln_pred=_is_green))
    a_scr = one(_find_rr(shapes, page, 82, 91, 31, 40, (3.3, 4.3), (2.4, 3.2),
                         ln_pred=_is_green))
    inner4 = sorted(_find_rr(shapes, page, 40, 80, 12, 18, (2.4, 3.2), (4.6, 5.3),
                             ln_pred=_is_orange), key=lambda s: s["x"])

    results = []

    # 01 左侧数据输入区块框。严格对齐细则，仅踩细则明确写出的点：
    #    ① 位置 0%–8% / 20%–30%；② 宽 4.5–5.5、高 5.5–6.5；③ 填充为极浅青色；
    #    ④ 边线为青绿色且是单实线（线型不能是虚线）；⑤ 边线色逐通道落在
    #    #0F9F9A–#37ADA0 之间。与办公软件对齐：未写 a:prstDash 时按实线计。
    hit = False
    for s in _find_rr(shapes, page, 0, 8, 20, 30, (4.5, 5.5), (5.5, 6.5),
                      ln_pred=lambda c: _near_rgb(c, "0F9F9A", "37ADA0"),
                      fill_pred=lambda c: _is_verylight(c) and (_is_teal(c) or _is_white(c))):
        if (s["ln"].get("dash") or "solid") not in _DASH_PRST:
            hit = True
    results.append(("+1：" if hit else "+0：") + "页面从左至右0%–8%、从上至下20%–30%出现一个圆角矩形，宽4.5–5.5cm，高5.5–6.5cm，填充为极浅青色，边线为青绿色单实线，边线颜色RGB接近#0F9F9A–#37ADA0")

    # 02 输入标签两行文本：位置 0%–12% / 22%–32%，两行、居中、加粗、13–16 磅，
    #    字色逐通道落在 #0A6F74–#1A787D 之间。换行既可能来自 a:br，也可能来自
    #    两个 a:p，这里统一按“去空白后依次包含两个词”＋“非空行数为 2”判定。
    hit = False
    for s in tf("Multivariate", "TimeSeries"):
        if not page.at(s, 0.0, 12.0, 22.0, 32.0):
            continue
        lines = [t for t in _flat(s["tx"]).split("\n") if t.strip()]
        if len(lines) != 2:
            continue
        if not (_all_ctr(s["tx"]) and _all_bold(s["tx"])):
            continue
        if not _rng(_dom_sz(s["tx"]), 13.0, 16.0):
            continue
        cols = _run_colors(s["tx"])
        if cols and all(_near_rgb(c, "0A6F74", "1A787D") for c in cols):
            hit = True
    results.append(("+1：" if hit else "+0：") + "页面从左至右0%–12%、从上至下22%–32%出现“Multivariate”与“Time Series”两行文本，居中对齐，加粗，字号13–16磅，颜色RGB接近#0A6F74–#1A787D")

    # 03 折线图竖直网格线。逐条踩点：① 竖直细线（宽≤0.08）且长 3.0–3.6；
    #    ② 顶端 34%–38%、最左端 2%–6%；③ 数量≥6；④ 等间距（相邻中心距与均值
    #    偏差≤0.2cm）；⑤ 线型为点线或短划线；⑥ 网格内部不出现贯穿整体宽度的
    #    横向网格线——只把“浅灰色、宽度≥网格宽 80%、且纵向严格落在网格内部”的
    #    横线算作贯穿，网格上下边缘处的坐标轴线不计入。
    grid = [s for s in shapes
            if s.get("x") is not None and s["w"] <= 0.08
            and _rng(s["h"], 3.0, 3.6)
            and _is_lightgray(s["ln"].get("color"))
            and _rng(page.py(s["y"]), 34.0, 38.0)]
    hit = False
    if len(grid) >= 6:
        xs = sorted(_cx(s) for s in grid)
        gaps = [xs[k + 1] - xs[k] for k in range(len(xs) - 1)]
        mean = sum(gaps) / len(gaps)
        even = all(abs(g - mean) <= 0.2 for g in gaps)
        dashed = all((s["ln"].get("dash") or "") in _DASH_PRST for s in grid)
        left_ok = _rng(page.px(min(s["x"] for s in grid)), 2.0, 6.0)
        gx1, gx2 = xs[0], xs[-1]
        gy1 = min(s["y"] for s in grid)
        gy2 = max(_y2(s) for s in grid)
        cross = False
        for s in shapes:
            if s.get("x") is None or s["h"] > 0.08:
                continue
            if s["w"] < 0.8 * max(gx2 - gx1, 0.01):
                continue
            if not _is_lightgray(s["ln"].get("color")):
                continue
            if s["x"] > gx1 + 0.2 or _x2(s) < gx2 - 0.2:
                continue
            if gy1 + 0.15 <= _cy(s) <= gy2 - 0.15:
                cross = True
        hit = even and dashed and left_ok and not cross
    results.append(("+3：" if hit else "+0：") + "页面出现折线图竖直网格线，顶端位于从上至下34%–38%、左端位于从左至右2%–6%，数量不少于6条且等间距排列，长度3.0–3.6cm，颜色为浅灰色，线型为点线或短划线，该网格范围内不出现贯穿其整体宽度的横向网格线")

    # 04 四条数据曲线。逐条踩点：① 区域 2%–16% / 35%–55%；② 蓝、橙、紫、绿四
    #    种颜色各自都要有；③ 每种颜色要么是单个形状，要么其全部线段首尾相接
    #    （端点间距≤0.05cm 视为相接），互不相接的独立短线段拼接不算；
    #    ④ 四条曲线纵向中心两两相距≥0.25cm，体现上下错开；⑤ 合并外框宽
    #    3.0–4.0、高 2.0–3.5。近黑色的坐标轴/引线不参与配色归类。
    cand = [s for s in shapes
            if s.get("x") is not None and page.at(s, 2.0, 16.0, 35.0, 55.0)
            and s["ln"].get("color") and s.get("fill") in (None, "none")
            and not _is_black(s["ln"].get("color"))]
    groups = []
    for pred in (_is_blue, _is_orange, _is_purple, _is_green):
        groups.append([s for s in cand if pred(s["ln"].get("color"))])
    hit = False
    if all(groups) and all(_connected(g, 0.05) for g in groups):
        boxes = [_bbox(g) for g in groups]
        ys = sorted(b[1] + b[3] / 2.0 for b in boxes)
        sep = all(ys[k + 1] - ys[k] >= 0.25 for k in range(len(ys) - 1))
        uw = max(b[0] + b[2] for b in boxes) - min(b[0] for b in boxes)
        uh = max(b[1] + b[3] for b in boxes) - min(b[1] for b in boxes)
        hit = sep and _rng(uw, 3.0, 4.0) and _rng(uh, 2.0, 3.5)
    results.append(("+5：" if hit else "+0：") + "页面从左至右2%–16%、从上至下35%–55%范围内出现蓝色、橙色、紫色、绿色四条数据曲线，每种颜色各为一条自左至右连贯的折线或曲线，为单个曲线形状或首尾相接的连续线段，不由互不相接的独立短线段拼接，四条曲线上下错开分布，整体宽3.0–4.0cm，高2.0–3.5cm")

    def inside(t, r, pad=0.35, ctol=0.45):
        """文本框是否落在目标矩形内部：纵向落在矩形高度范围内，横向中心基本对齐。"""
        if t is None or r is None:
            return False
        return (r["y"] - pad <= t["y"] <= r["y"] + r["h"] + pad
                and abs(_cx(t) - _cx(r)) <= ctol)

    # 05 序列编码区块框：位置 18%–28% / 4%–12%，宽 5.5–6.5、高 4.2–5.0，边线青绿。
    #    本条只判外框本身，不涉及线型，虚线与实线都算。
    hit = bool(_find_rr(shapes, page, 18, 28, 4, 12, (5.5, 6.5), (4.2, 5.0),
                        ln_pred=_is_teal))
    results.append(("+1：" if hit else "+0：") + "页面从左至右18%–28%、从上至下4%–12%出现一个圆角矩形，宽5.5–6.5cm，高4.2–5.0cm，边线为青绿色")

    # 06 序列编码标题：位置 20%–28% / 0%–8%，字号 13–16、加粗、青绿色。
    #    “1. Sequence Encoder”与“1.Sequence Encoder”去空白后完全一致，统一匹配。
    hit = False
    for s in tf("1.SequenceEncoder"):
        if not page.at(s, 20.0, 28.0, 0.0, 8.0):
            continue
        if not _all_bold(s["tx"]) or not _rng(_dom_sz(s["tx"]), 13.0, 16.0):
            continue
        cols = _run_colors(s["tx"])
        if cols and all(_is_teal(c) for c in cols):
            hit = True
    results.append(("+1：" if hit else "+0：") + "页面从左至右20%–28%、从上至下0%–8%出现“1. Sequence Encoder”或“1.Sequence Encoder”文本，字号13–16磅，加粗，颜色为青绿色")

    # 07 两个 GRU 圆角矩形。逐条踩点：① 横向 22%–28%；② 上框顶端 9%–13%、下框
    #    顶端 19%–23%；③ 各宽 4.4–5.2、高 1.5–1.9；④ 垂直间距 0.3–0.7；
    #    ⑤ 填充浅青、边线青绿；⑥ 两框内各有“Gated Recurrent”“Unit”两行居中
    #    文本，字号 12–15。居中按几何居中判定：文本框中心与矩形中心横向对齐，
    #    这样段落对齐属性缺省（继承）时也不会误判。
    def _gru(y1, y2):
        return _find_rr(shapes, page, 22, 28, y1, y2, (4.4, 5.2), (1.5, 1.9),
                        ln_pred=_is_teal,
                        fill_pred=lambda c: _is_light(c) and (_is_teal(c) or _is_white(c)))

    hit = False
    for up in _gru(9, 13):
        for lo in _gru(19, 23):
            if not _rng(lo["y"] - (up["y"] + up["h"]), 0.3, 0.7):
                continue
            good = 0
            for r in (up, lo):
                for t in tf("GatedRecurrent", "Unit"):
                    lines = [z for z in _flat(t["tx"]).split("\n") if z.strip()]
                    if len(lines) == 2 and _rng(_dom_sz(t["tx"]), 12.0, 15.0) and inside(t, r):
                        good += 1
                        break
            if good == 2:
                hit = True
    results.append(("+3：" if hit else "+0：") + "页面从左至右22%–28%出现上下并列的两个圆角矩形，上方矩形顶端位于从上至下9%–13%、下方矩形顶端位于从上至下19%–23%，各宽4.4–5.2cm、高1.5–1.9cm，垂直间距0.3–0.7cm，填充为浅青色，边线为青绿色，两者内部各出现居中的“Gated Recurrent”与“Unit”两行文本，字号12–15磅")

    # 08 模式编码区块框：位置 18%–28% / 38%–45%，宽 5.0–6.0、高 4.8–5.4，
    #    边线色逐通道落在 #8A65C5–#9877C7 之间。
    hit = bool(_find_rr(shapes, page, 18, 28, 38, 45, (5.0, 6.0), (4.8, 5.4),
                        ln_pred=lambda c: _near_rgb(c, "8A65C5", "9877C7")))
    results.append(("+1：" if hit else "+0：") + "页面从左至右18%–28%、从上至下38%–45%出现一个圆角矩形，宽5.0–6.0cm，高4.8–5.4cm，边线为紫色，边线颜色RGB接近#8A65C5–#9877C7")

    # 09 模式编码标题：位置 21%–27% / 38%–45%，字号 13–16、加粗、紫色。
    #    文本可能被拆成多个 run（如“2. Pattern Encode”＋“r”），按去空白后包含判定。
    hit = False
    for s in tf("2.PatternEncoder"):
        if not page.at(s, 21.0, 27.0, 38.0, 45.0):
            continue
        if not _all_bold(s["tx"]) or not _rng(_dom_sz(s["tx"]), 13.0, 16.0):
            continue
        cols = _run_colors(s["tx"])
        if cols and all(_is_purple(c) for c in cols):
            hit = True
    results.append(("+1：" if hit else "+0：") + "页面从左至右21%–27%、从上至下38%–45%出现“2. Pattern Encoder”文本，字号13–16磅，加粗，颜色为紫色")

    def _txt_bottom(t):
        """文本末行底端位置（cm）：按文本框底端取，不依赖办公软件实测的文本块高度。"""
        return t["y"] + t["h"]

    # 10 紫色方块矩阵。逐条踩点：① 区域 22%–30% / 44%–54%；② 小方块数量≥15；
    #    ③ 行列数为 3×5 或 4×4；④ 整体宽 2.8–3.4、高 1.7–2.2；⑤ 各方块填充为
    #    深浅不同的紫色系（至少两种不同填充色）、边线为紫色；⑥ 无缝拼接——整体
    #    宽高应约等于“单块尺寸 × 列数/行数”，块间留白会使乘积明显小于整体尺寸。
    #    注意矩阵底衬矩形用的是主题色（解析后为蓝色），不参与本条判定。
    cells = [s for s in shapes
             if s.get("x") is not None and page.at(s, 20.0, 36.0, 42.0, 58.0)
             and _rng(s["w"], 0.5, 0.75) and _rng(s["h"], 0.5, 0.75)
             and _is_purple(s.get("fill")) and _is_purple(s["ln"].get("color"))]
    hit = False
    if len(cells) >= 15:
        rows = _clusters([_cy(s) for s in cells], 0.2)
        cols = _clusters([_cx(s) for s in cells], 0.2)
        bx, by, bw, bh = _bbox(cells)
        aw = sum(s["w"] for s in cells) / len(cells)
        ah = sum(s["h"] for s in cells) / len(cells)
        seamless = (abs(bw - aw * len(cols)) <= 0.3
                    and abs(bh - ah * len(rows)) <= 0.3)
        hit = ((len(rows), len(cols)) in ((3, 5), (4, 4))
               and len(set(s.get("fill") for s in cells)) >= 2
               and seamless and _rng(bw, 2.8, 3.4) and _rng(bh, 1.7, 2.2)
               and _rng(page.px(bx), 22.0, 30.0) and _rng(page.py(by), 44.0, 54.0))
    results.append(("+3：" if hit else "+0：") + "页面从左至右22%–30%、从上至下44%–54%范围内出现由不少于15个小方块无缝拼接成的矩阵，排列为3行5列或4行4列，整体宽2.8–3.4cm、高1.7–2.2cm，各方块填充为深浅不同的紫色系颜色，边线为紫色")

    # 11 Temporal / Embedding 两行文本。逐条踩点：① 区域 24%–30% / 56%–62%；
    #    ② 居中、加粗、字号 12–15；③ 文本仅占两行且无跨行断词——文本自身两行，
    #    实际排出来也必须正好两行，框太窄折成三行即不满足；④ 末行底端≤67%。
    #    ③④ 的行数按字宽表估算，末行底端按文本框底端取。
    hit = False
    for s in tf("Temporal", "Embedding"):
        if not page.at(s, 24.0, 30.0, 56.0, 62.0):
            continue
        if not (_all_ctr(s["tx"]) and _all_bold(s["tx"])):
            continue
        if not _rng(_dom_sz(s["tx"]), 12.0, 15.0):
            continue
        if _para_lines(s) != 2 or _est_lines(s) != 2:
            continue
        if page.py(_txt_bottom(s)) <= 67.0:
            hit = True
    results.append(("+1：" if hit else "+0：") + "页面从左至右24%–30%、从上至下56%–62%出现“Temporal”与“Embedding”两行居中文本，字号12–15磅，加粗，文本仅占两行且无跨行断词，末行底端位于从上至下67%以内")

    # 12 融合区块框：位置 40%–46% / 4%–12%，宽 13.0–15.0、高 7.0–7.8，填充极浅
    #    橙色，边线色逐通道落在 #FF6B1A–#FBAB63 之间。
    hit = bool(_find_rr(shapes, page, 40, 46, 4, 12, (13.0, 15.0), (7.0, 7.8),
                        ln_pred=lambda c: _near_rgb(c, "FF6B1A", "FBAB63"),
                        fill_pred=lambda c: _is_verylight(c) and _is_paleorange(c)))
    results.append(("+1：" if hit else "+0：") + "页面从左至右40%–46%、从上至下4%–12%出现一个圆角矩形，宽13.0–15.0cm，高7.0–7.8cm，填充为极浅橙色，边线为橙色，边线颜色RGB接近#FF6B1A–#FBAB63")

    # 13 融合区块标题：位置 45%–52% / 0%–8%，字号 15–17、加粗、橙色。
    hit = False
    for s in tf("3.Fusion", "RepresentationLearning"):
        if not page.at(s, 45.0, 52.0, 0.0, 8.0):
            continue
        if not _all_bold(s["tx"]) or not _rng(_dom_sz(s["tx"]), 15.0, 17.0):
            continue
        cols = _run_colors(s["tx"])
        if cols and all(_is_orange(c) for c in cols):
            hit = True
    results.append(("+1：" if hit else "+0：") + "页面从左至右45%–52%、从上至下0%–8%出现“3. Fusion & Representation Learning”文本，字号15–17磅，加粗，颜色为橙色")

    # 14 融合区块内横向并排的四个圆角矩形。逐条踩点：① 顶端均在 13%–17% 的同一
    #    水平线上（顶端极差≤0.15cm）；② 最左者左端 41%–46%；③ 各宽 2.5–3.1、高
    #    4.7–5.2；④ 相邻间距 0.5–0.8；⑤ 填充浅橙且颜色深于外围橙色框的填充色
    #    （用亮度比较，浅于外框即不满足）；⑥ 边线橙色；⑦ 自左至右四处文本齐全、
    #    字号 11–14；⑧ “Layer Normalization”必须以两行完整显示，排成三行即
    #    出现跨行断词，不满足。
    quad = sorted(_find_rr(shapes, page, 41, 78, 13, 17, (2.5, 3.1), (4.7, 5.2),
                           ln_pred=_is_orange,
                           fill_pred=lambda c: _is_paleorange(c) and _is_light(c)),
                  key=lambda s: s["x"])
    hit = False
    if len(quad) == 4 and a_fus is not None:
        ytops = [s["y"] for s in quad]
        gaps = [quad[k + 1]["x"] - _x2(quad[k]) for k in range(3)]
        base_lum = _lum(a_fus.get("fill"))
        darker = base_lum is not None and all(
            _lum(s.get("fill")) is not None and _lum(s["fill"]) < base_lum
            for s in quad)
        labels = (("Concatenate",), ("Self-Attention", "Block"),
                  ("Feed-Forward", "Network"), ("Layer", "Normalization"))
        ok_txt = True
        for r, keys in zip(quad, labels):
            found = None
            for t in tf(*keys):
                if inside(t, r) and _rng(_dom_sz(t["tx"]), 11.0, 14.0):
                    found = t
                    break
            if found is None:
                ok_txt = False
                break
            if keys[0] == "Layer" and not (_para_lines(found) == 2
                                           and _est_lines(found) == 2):
                ok_txt = False
                break
        hit = (max(ytops) - min(ytops) <= 0.15
               and _rng(page.px(quad[0]["x"]), 41.0, 46.0)
               and all(_rng(g, 0.5, 0.8) for g in gaps)
               and darker and ok_txt)
    results.append(("+3：" if hit else "+0：") + "页面出现横向并排的四个圆角矩形，顶端均位于从上至下13%–17%的同一水平线上，最左者左端位于从左至右41%–46%，各宽2.5–3.1cm、高4.7–5.2cm，相邻间距0.5–0.8cm，填充为浅橙色且深于其外围橙色圆角矩形的填充色，边线为橙色，四者内部自左至右分别出现居中的“Concatenate”、“Self-Attention Block”、“Feed-Forward Network”、“Layer Normalization”文本，字号11–14磅，其中“Layer Normalization”以“Layer”“Normalization”两行完整显示且无跨行断词")

    # 15 网格状表格图形。逐条踩点：① 区域 43%–48% / 23%–29%；② 线条橙棕色；
    #    ③ 拆分为左右两组且两组之间存在间隙——把成员矩形按横向区间聚类，必须
    #    聚出不少于 2 组；④ 每组内可辨识出不少于 2 列 3 行的网格，故要求存在
    #    独立的橙棕色横向分隔线不少于 2 条；⑤ 整体宽 1.3–2.0、高 2.2–2.8；
    #    ⑥ 填充为浅橙色或无填充。仅靠竖向切分成多列、没有横向分隔线的画法，
    #    辨识不出行，不满足。
    mem = [s for s in shapes
           if s.get("x") is not None and page.at(s, 43.0, 48.0, 23.0, 29.0)
           and _is_brownorange(s["ln"].get("color"))]
    hit = False
    if mem:
        cellr = [s for s in mem if s["h"] >= 0.3 and s["w"] >= 0.2]
        divs = [s for s in mem if s["h"] <= 0.06 and s["w"] >= 0.15]
        grp = _interval_clusters([(s["x"], _x2(s)) for s in cellr], 0.15)
        bx, by, bw, bh = _bbox(mem)
        fills_ok = all(s.get("fill") in (None, "none")
                       or (_is_light(s.get("fill")) and _is_paleorange(s.get("fill")))
                       for s in mem)
        rows_ok = True
        for g in grp:
            gx1 = min(cellr[k]["x"] for k in g)
            gx2 = max(_x2(cellr[k]) for k in g)
            n = sum(1 for d in divs
                    if d["x"] >= gx1 - 0.1 and _x2(d) <= gx2 + 0.1)
            if n < 2:
                rows_ok = False
        hit = (len(grp) >= 2 and len(divs) >= 2 and rows_ok and fills_ok
               and _rng(bw, 1.3, 2.0) and _rng(bh, 2.2, 2.8))
    results.append(("+3：" if hit else "+0：") + "页面从左至右43%–48%、从上至下23%–29%出现网格状表格图形，由矩形与横竖分隔线或多个小矩形拼接构成，拆分为左右两组且两组之间存在间隙，每组可辨识出不少于2列3行的网格，整体宽1.3–2.0cm、高2.2–2.8cm，填充为浅橙色，线条为橙棕色")

    # 16 注意力示意图形。逐条踩点：① 整体外框左上角 52%–58% / 25%–29%；
    #    ② 底部一个横向矩形条（宽≥1.2、高≤0.4）；③ 其上方≥4 个圆形，直径
    #    0.2–0.45，纵向聚成上下两层；④ 上下层之间≥3 条连线；⑤ 各连线两端都要
    #    落在圆形的圆心或边缘处（端点到圆心距离≤半径+0.15），出现穿出圆形之外
    #    或反向延伸的连线即不满足；⑥ 整体宽 1.5–2.2、高 1.6–2.2。
    #    圆形到底部矩形条的引线属于层与底条之间的连线，不计入上下层连线。
    reg = _icon_group([s for s in shapes
                       if s.get("x") is not None and s["kind"] != "grpSp"
                       and page.at(s, 48.0, 62.0, 20.0, 40.0)])
    circ = [s for s in reg if s.get("prst") == "ellipse"
            and _rng(s["w"], 0.2, 0.45) and _rng(s["h"], 0.2, 0.45)]
    bars = [s for s in reg if not _is_conn(s) and s["w"] >= 1.2 and s["h"] <= 0.4]
    hit = False
    if len(circ) >= 4 and bars:
        yc = _clusters([_cy(s) for s in circ], 0.2)
        if len(yc) >= 2:
            layers = sorted(yc, key=lambda g: sum(_cy(circ[k]) for k in g) / len(g))
            btop = min(circ[k]["y"] for k in layers[-1])
            good_lk = bad_lk = 0
            for s in reg:
                if not _is_conn(s) or s["h"] < 0.2:
                    continue
                e = _ends(s)
                if e is None or min(e[0][1], e[1][1]) >= btop + 0.05:
                    continue
                ok = True
                for (qx, qy) in e:
                    if not any(math.hypot(qx - _cx(c), qy - _cy(c))
                               <= max(c["w"], c["h"]) / 2.0 + 0.15 for c in circ):
                        ok = False
                if ok:
                    good_lk += 1
                else:
                    bad_lk += 1
            bx, by, bw, bh = _bbox(circ + bars)
            hit = (good_lk >= 3 and bad_lk == 0
                   and _rng(bw, 1.5, 2.2) and _rng(bh, 1.6, 2.2)
                   and _rng(page.px(bx), 52.0, 58.0)
                   and _rng(page.py(by), 25.0, 29.0))
    results.append(("+3：" if hit else "+0：") + "页面从左至右52%–58%、从上至下25%–29%出现注意力示意图形，底部为一个横向矩形条，其上方有不少于4个圆形分上下两层排布，圆形直径0.2–0.45cm，上下层圆形之间由不少于3条连线相连，各连线两端均落在圆形的圆心或边缘处且无穿出圆形之外或反向延伸的连线，整体宽1.5–2.2cm、高1.6–2.2cm")

    # 17 神经网络示意图形。逐条踩点：① 节点外框左上角 62%–68% / 25%–29%；
    #    ② ≥6 个圆形节点，直径 0.25–0.5，横向聚成不少于两列；③ ≥6 条连线；
    #    ④ 每条连线的两个端点都必须落在某个节点上（到节点外框距离≤0.22），
    #    出现穿出节点之外或方向错误的连线即不满足；⑤ 每个节点至少被 2 个连线
    #    端点连上，孤立的多余节点即不满足；⑥ 整体宽 1.4–2.0、高 1.5–2.1。
    reg = _icon_group([s for s in shapes
                       if s.get("x") is not None and s["kind"] != "grpSp"
                       and page.at(s, 62.0, 72.0, 20.0, 40.0)])
    nodes = [s for s in reg if s.get("prst") == "ellipse"
             and _rng(s["w"], 0.25, 0.5) and _rng(s["h"], 0.25, 0.5)]
    links = [s for s in reg if _is_conn(s)]
    hit = False
    if len(nodes) >= 6 and len(links) >= 6:
        touch = [0] * len(nodes)
        ok = True
        for s in links:
            e = _ends(s)
            if e is None:
                ok = False
                continue
            for (qx, qy) in e:
                best, bi = 1e9, -1
                for k, nd in enumerate(nodes):
                    d = _box_dist(qx, qy, nd)
                    if d < best:
                        best, bi = d, k
                if best > 0.22:
                    ok = False
                else:
                    touch[bi] += 1
        bx, by, bw, bh = _bbox(nodes)
        hit = (ok and len(_clusters([_cx(s) for s in nodes], 0.2)) >= 2
               and all(t >= 2 for t in touch)
               and _rng(bw, 1.4, 2.0) and _rng(bh, 1.5, 2.1)
               and _rng(page.px(bx), 62.0, 68.0)
               and _rng(page.py(by), 25.0, 29.0))
    results.append(("+3：" if hit else "+0：") + "页面从左至右62%–68%、从上至下25%–29%出现神经网络示意图形，不少于6个圆形节点分两列排布，节点直径0.25–0.5cm，节点之间由不少于6条连线交叉相连，各连线两端均分别落在左右两列节点上且无穿出节点之外或方向错误的连线，两列之外不出现多余的孤立节点，整体宽1.4–2.0cm、高1.5–2.1cm")

    # 18 曲线示意图形。逐条踩点：① 整体外框左上角 72%–78% / 26%–32%；
    #    ② 一条带右向箭头的水平坐标轴线（高≤0.1、宽 1.0–1.8、箭头端在右）；
    #    ③ 一条自左至右连贯的单峰隆起曲线：单个形状，或全部线段首尾相接
    #    （端点间距≤0.05cm）；互不相接的独立短线段拼接不满足；
    #    ④ 整体宽 1.3–1.8、高 1.3–1.8。纵向细线属于坐标轴，不计入曲线。
    reg = _icon_group([s for s in shapes
                       if s.get("x") is not None and s["kind"] != "grpSp"
                       and page.at(s, 70.0, 82.0, 22.0, 40.0)])
    axis = []
    for s in reg:
        if s["h"] > 0.1 or not _rng(s["w"], 1.0, 1.8):
            continue
        tp, tl = _tip(s), _tail_pt(s)
        if tp is not None and tl is not None and tp[0] > tl[0]:
            axis.append(s)
    hit = False
    if axis:
        curve = [s for s in reg if s not in axis and s["w"] > 0.1]
        if curve and (len(curve) == 1 or _connected(curve, 0.05)):
            bx, by, bw, bh = _bbox(reg)
            hit = (_rng(bw, 1.3, 1.8) and _rng(bh, 1.3, 1.8)
                   and _rng(page.px(bx), 72.0, 78.0)
                   and _rng(page.py(by), 26.0, 32.0))
    results.append(("+3：" if hit else "+0：") + "页面从左至右72%–78%、从上至下26%–32%出现曲线示意图形，包含一条自左至右连贯的单峰隆起曲线，为单个曲线形状或首尾相接的连续线段，不由互不相接的独立短线段拼接，并包含一条带右向箭头的水平坐标轴线，整体宽1.3–1.8cm、高1.3–1.8cm")

    # 19 异常评分区块框：位置 82%–88% / 10%–15%，宽 4.5–5.4、高 7.2–8.0，
    #    边线色逐通道落在 #6CB447–#82C358 之间。
    hit = bool(_find_rr(shapes, page, 82, 88, 10, 15, (4.5, 5.4), (7.2, 8.0),
                        ln_pred=lambda c: _near_rgb(c, "6CB447", "82C358")))
    results.append(("+1：" if hit else "+0：") + "页面从左至右82%–88%、从上至下10%–15%出现一个圆角矩形，宽4.5–5.4cm，高7.2–8.0cm，边线为绿色，边线颜色RGB接近#6CB447–#82C358")

    # 20 异常评分标题：位置 82%–88% / 4%–9%，字号 15–17、加粗、绿色。
    hit = False
    for s in tf("4.AnomalyScoring"):
        if not page.at(s, 82.0, 88.0, 4.0, 9.0):
            continue
        if not _all_bold(s["tx"]) or not _rng(_dom_sz(s["tx"]), 15.0, 17.0):
            continue
        cols = _run_colors(s["tx"])
        if cols and all(_is_green(c) for c in cols):
            hit = True
    results.append(("+1：" if hit else "+0：") + "页面从左至右82%–88%、从上至下4%–9%出现“4. Anomaly Scoring”或“4.Anomaly Scoring”文本，字号15–17磅，加粗，颜色为绿色")

    def _lightgreen(c):
        return _is_light(c) and (_is_green(c) or _is_white(c))

    # 21 重构模块矩形＋其下 4 个小矩形。逐条踩点：① 矩形位置 84%–89% / 14%–18%，
    #    宽 3.5–4.1、高 2.6–3.0，填充浅绿、边线绿色；② 内上部有居中的
    #    “Reconstruction”“Module”两行文本，字号 11–14；③ 该文本必须两行完整
    #    显示，排成三行即出现跨行断词，不满足；④ 其下并排的小矩形数量正好 4 个，
    #    顶端 24%–27%，各宽 0.35–0.55、高 0.4–0.6，填充浅绿、边线绿色；
    #    ⑤ 4 个小矩形整体宽 1.8–2.6。多出一个方块会同时破坏数量与整体宽度。
    rec_r = one(_find_rr(shapes, page, 84, 89, 14, 18, (3.5, 4.1), (2.6, 3.0),
                         ln_pred=_is_green, fill_pred=_lightgreen))
    hit = False
    if rec_r is not None:
        t = None
        for z in tf("Reconstruction", "Module"):
            if inside(z, rec_r) and _rng(_dom_sz(z["tx"]), 11.0, 14.0):
                t = z
                break
        sq = [s for s in shapes
              if s.get("x") is not None and _rng(page.py(s["y"]), 24.0, 27.0)
              and _rng(s["w"], 0.35, 0.55) and _rng(s["h"], 0.4, 0.6)
              and _is_green(s["ln"].get("color")) and _lightgreen(s.get("fill"))
              and s["x"] >= rec_r["x"] - 0.2 and _x2(s) <= _x2(rec_r) + 0.2]
        if t is not None and len(sq) == 4:
            hit = (_para_lines(t) == 2 and _est_lines(t) == 2
                   and _rng(_bbox(sq)[2], 1.8, 2.6))
    results.append(("+3：" if hit else "+0：") + "页面从左至右84%–89%、从上至下14%–18%出现一个圆角矩形，宽3.5–4.1cm、高2.6–3.0cm，填充为浅绿色，边线为绿色，其内上部出现居中的“Reconstruction”与“Module”两行文本，字号11–14磅，两行完整显示且无跨行断词，该文本下方并排出现4个小矩形，顶端位于从上至下24%–27%，各宽0.35–0.55cm、高0.4–0.6cm，填充为浅绿色，边线为绿色，4个小矩形整体宽1.8–2.6cm")

    # 22 异常分数矩形：位置 84%–89% / 33%–38%，宽 3.5–4.1、高 2.6–3.0，填充浅绿、
    #    边线绿色，内上部居中的“Anomaly Score”字号 11–14，且必须单行完整显示。
    #    行数按字宽表估算，框宽不足折成两行即不满足。
    scr_r = one(_find_rr(shapes, page, 84, 89, 33, 38, (3.5, 4.1), (2.6, 3.0),
                         ln_pred=_is_green, fill_pred=_lightgreen))
    hit = False
    if scr_r is not None:
        for z in tf("AnomalyScore"):
            if not inside(z, scr_r) or not _rng(_dom_sz(z["tx"]), 11.0, 14.0):
                continue
            if _para_lines(z) == 1 and _est_lines(z) == 1:
                hit = True
    results.append(("+1：" if hit else "+0：") + "页面从左至右84%–89%、从上至下33%–38%出现一个圆角矩形，宽3.5–4.1cm、高2.6–3.0cm，填充为浅绿色，边线为绿色，其内上部出现居中的“Anomaly Score”文本，字号11–14磅，单行完整显示且未折行为两行")

    # 23 概率分布示意图形。逐条踩点：① 水平基线顶端 45%–49%、左端 86%–92%，
    #    长度 2.0–2.6；② 基线上方一条自左至右连贯的钟形曲线：单个形状，或全部
    #    线段首尾相接（端点间距≤0.05cm），互不相接的独立短线段拼接不满足；
    #    ③ 钟形曲线右侧尾部（横向中心位于基线中点右侧）存在橙色或浅红色填充的
    #    小面积区域，面积 0.05–1.0 平方厘米。纵向细线为分位线，不计入曲线。
    base = [s for s in shapes
            if s.get("x") is not None and s["h"] <= 0.1
            and _rng(s["w"], 2.0, 2.6)
            and _rng(page.py(s["y"]), 45.0, 49.0)
            and _rng(page.px(s["x"]), 86.0, 92.0)]
    hit = False
    if base:
        b0 = base[0]
        near = [s for s in shapes
                if s.get("x") is not None and page.at(s, 84.0, 100.0, 38.0, 50.0)]
        bell = [s for s in near if s is not b0 and s["w"] >= 0.5 and s["h"] >= 0.15
                and s.get("fill") in (None, "none")]
        tail = [s for s in near
                if 0.05 <= s["w"] * s["h"] <= 1.0 and _cx(s) > _cx(b0)
                and (_is_paleorange(s.get("fill")) or _is_pink(s.get("fill")))]
        hit = bool(bell) and bool(tail) and (len(bell) == 1 or _connected(bell, 0.05))
    results.append(("+3：" if hit else "+0：") + "页面出现概率分布示意图形，水平基线顶端位于从上至下45%–49%、左端位于从左至右86%–92%，基线长度2.0–2.6cm，其上方出现一条自左至右连贯的钟形曲线，为单个曲线形状或首尾相接的连续线段，不由互不相接的独立短线段拼接，钟形曲线右侧尾部存在橙色或浅红色填充的小面积区域")

    # 24 决策区块框：位置 82%–88% / 57%–62%，宽 4.4–5.2、高 4.5–5.0，边线色逐通道
    #    落在 #D95F8B–#DA7292 之间。
    hit = bool(_find_rr(shapes, page, 82, 88, 57, 62, (4.4, 5.2), (4.5, 5.0),
                        ln_pred=lambda c: _near_rgb(c, "D95F8B", "DA7292")))
    results.append(("+1：" if hit else "+0：") + "页面从左至右82%–88%、从上至下57%–62%出现一个圆角矩形，宽4.4–5.2cm，高4.5–5.0cm，边线为粉红色，边线颜色RGB接近#D95F8B–#DA7292")

    # 25 决策区块标题：位置 85%–90% / 60%–64%，字号 15–17、加粗、粉红色。
    hit = False
    for s in tf("5.Decision"):
        if not page.at(s, 85.0, 90.0, 60.0, 64.0):
            continue
        if not _all_bold(s["tx"]) or not _rng(_dom_sz(s["tx"]), 15.0, 17.0):
            continue
        cols = _run_colors(s["tx"])
        if cols and all(_is_pink(c) for c in cols):
            hit = True
    results.append(("+1：" if hit else "+0：") + "页面从左至右85%–90%、从上至下60%–64%出现“5. Decision”或“5.Decision”文本，字号15–17磅，加粗，颜色为粉红色")

    # 26 阈值判定与结论两个圆角矩形。逐条踩点：① 横向 85%–90%；② 上框高 1.2–1.5
    #    且顶端 65%–69%，下框高 0.8–1.1 且顶端 75%–78%，均宽 3.4–4.0；
    #    ③ 填充接近白色、边线粉红；④ 上框内“Thresholding”、下框内
    #    “Normal / Anomaly”居中，字号 11–15，且两处都必须单行完整显示，
    #    框宽不足折成两行即不满足；⑤ 两框之间一条长度 0.3–0.6cm 的竖直向下箭头。
    up = one(_find_rr(shapes, page, 85, 90, 65, 69, (3.4, 4.0), (1.2, 1.5),
                      ln_pred=_is_pink, fill_pred=_is_verylight))
    lo = one(_find_rr(shapes, page, 85, 90, 75, 78, (3.4, 4.0), (0.8, 1.1),
                      ln_pred=_is_pink, fill_pred=_is_verylight))
    hit = False
    if up is not None and lo is not None:
        t1 = t2 = None
        for z in tf("Thresholding"):
            if inside(z, up) and _rng(_dom_sz(z["tx"]), 11.0, 15.0):
                t1 = z
        for z in tf("Normal", "Anomaly"):
            if inside(z, lo) and _rng(_dom_sz(z["tx"]), 11.0, 15.0):
                t2 = z
        arw = False
        for s in shapes:
            if s.get("x") is None or s["w"] > 0.15 or not _rng(s["h"], 0.3, 0.6):
                continue
            tp, tl = _tip(s), _tail_pt(s)
            if tp is None or tl is None or tp[1] <= tl[1]:
                continue
            if (min(tp[1], tl[1]) >= _y2(up) - 0.25
                    and max(tp[1], tl[1]) <= lo["y"] + 0.25
                    and abs(_cx(s) - _cx(up)) <= 0.6):
                arw = True
        if t1 is not None and t2 is not None and arw:
            hit = (_para_lines(t1) == 1 and _est_lines(t1) == 1
                   and _para_lines(t2) == 1 and _est_lines(t2) == 1)
    results.append(("+3：" if hit else "+0：") + "页面从左至右85%–90%出现上下两个圆角矩形，均宽3.4–4.0cm，上方矩形高1.2–1.5cm且顶端位于从上至下65%–69%，下方矩形高0.8–1.1cm且顶端位于从上至下75%–78%，填充为接近白色，边线为粉红色，上方矩形内出现“Thresholding”、下方矩形内出现“Normal / Anomaly”或“Normal/Anomaly”居中文本，字号11–15磅且均单行完整显示，两矩形之间出现一条长度0.3–0.6cm的竖直向下箭头线")

    # 27 训练目标区块框：位置 42%–46% / 55%–60%，宽 8.0–9.3、高 4.5–5.0，填充极浅
    #    蓝色，边线色逐通道落在 #136AC7–#2B7AD8 之间。极浅蓝填充的饱和度很低，
    #    按色相区间 195°–250° 判定，不套用一般蓝色的饱和度门槛。
    hit = bool(_find_rr(shapes, page, 42, 46, 55, 60, (8.0, 9.3), (4.5, 5.0),
                        ln_pred=lambda c: _near_rgb(c, "136AC7", "2B7AD8"),
                        fill_pred=lambda c: _is_verylight(c) and _hue_in(c, 195, 250, 0.01)))
    results.append(("+1：" if hit else "+0：") + "页面从左至右42%–46%、从上至下55%–60%出现一个圆角矩形，宽8.0–9.3cm，高4.5–5.0cm，填充为极浅蓝色，边线为蓝色，边线颜色RGB接近#136AC7–#2B7AD8")

    # 28 训练目标标题：位置 44%–51% / 57%–61%，字号 14–17、加粗、蓝色。
    hit = False
    for s in tf("6.TrainingObjective"):
        if not page.at(s, 44.0, 51.0, 57.0, 61.0):
            continue
        if not _all_bold(s["tx"]) or not _rng(_dom_sz(s["tx"]), 14.0, 17.0):
            continue
        cols = _run_colors(s["tx"])
        if cols and all(_is_blue(c) for c in cols):
            hit = True
    results.append(("+1：" if hit else "+0：") + "页面从左至右44%–51%、从上至下57%–61%出现“6. Training Objective”或“6.Training Objective”文本，字号14–17磅，加粗，颜色为蓝色")

    # 29 两项损失的左右并排矩形。逐条踩点：① 顶端均在 62%–65% 的同一水平线上
    #    （顶端极差≤0.15cm）；② 左者左端 43%–47%、右者左端 55%–59%；
    #    ③ 各宽 3.3–3.9、高 1.8–2.1；④ 边线蓝色；⑤ 左框内“Reconstruction Loss”、
    #    右框内“Regularization Loss”，字号 9–11.5。
    lft = _find_rr(shapes, page, 43, 47, 62, 65, (3.3, 3.9), (1.8, 2.1), ln_pred=_is_blue)
    rgt = _find_rr(shapes, page, 55, 59, 62, 65, (3.3, 3.9), (1.8, 2.1), ln_pred=_is_blue)
    hit = False
    for a in lft:
        for b in rgt:
            if abs(a["y"] - b["y"]) > 0.15:
                continue
            ok = True
            for r, key in ((a, "ReconstructionLoss"), (b, "RegularizationLoss")):
                if not any(inside(z, r) and _rng(_dom_sz(z["tx"]), 9.0, 11.5)
                           for z in tf(key)):
                    ok = False
            if ok:
                hit = True
    results.append(("+3：" if hit else "+0：") + "页面出现左右并排的两个圆角矩形，顶端均位于从上至下62%–65%的同一水平线上，左侧者左端位于从左至右43%–47%、右侧者左端位于从左至右55%–59%，各宽3.3–3.9cm、高1.8–2.1cm，边线为蓝色，左侧矩形内上部出现“Reconstruction Loss”文本、右侧矩形内上部出现“Regularization Loss”文本，字号9–11.5磅")

    # 30 两处带下标的数学符号。逐条踩点：① 区域 46%–62% / 66%–70%；② 去空白后
    #    恰为“Lrec”“Lreg”；③ 均为斜体；④ “rec”“reg”真正以下标格式呈现，即存在
    #    baseline 为负值的 run，靠字号缩小或位置摆放模拟的不算；⑤ 文本中不含
    #    下划线字符“_”；⑥ 不折行；⑦ 53%–57% / 65%–69% 另有“+”文本。
    def _sub_sym(word):
        for s in tf(word):
            if not page.at(s, 46.0, 62.0, 66.0, 70.0):
                continue
            flat = _flat(s["tx"])
            if "_" in flat or _squash(flat).lower() != word.lower():
                continue
            if not _all_italic(s["tx"]):
                continue
            if not any(_fnum(r.get("base"), 0.0) < 0 for r in _runs(s["tx"])):
                continue
            if _no_wrap(s):
                return True
        return False

    plus = any(page.at(s, 53.0, 57.0, 65.0, 69.0) and _squash(_flat(s["tx"])) == "+"
               for s in tf("+"))
    hit = plus and _sub_sym("Lrec") and _sub_sym("Lreg")
    results.append(("+3：" if hit else "+0：") + "页面从左至右46%–62%、从上至下66%–70%出现“Lrec”与“Lreg”两处数学符号文本，均为斜体，其中“rec”“reg”以下标格式呈现，基线降低，文本中不含下划线字符“_”且不折行，页面从左至右53%–57%、从上至下65%–69%出现“+”文本")

    # 31 总损失公式上方的水平分隔线：左端 43%–47%、顶端 75%–77%，长度 7.2–8.0，
    #    必须是单实线（线型不能是虚线），颜色为深灰或深蓝。
    hit = False
    for s in shapes:
        if s.get("x") is None or s["h"] > 0.06:
            continue
        if not _rng(s["w"], 7.2, 8.0):
            continue
        if not (_rng(page.px(s["x"]), 43.0, 47.0) and _rng(page.py(s["y"]), 75.0, 77.0)):
            continue
        if (s["ln"].get("dash") or "solid") in _DASH_PRST:
            continue
        if _is_darkgrayblue(s["ln"].get("color")):
            hit = True
    results.append(("+1：" if hit else "+0：") + "页面出现一条水平单实线，左端位于从左至右43%–47%，顶端位于从上至下75%–77%，长度7.2–8.0cm，颜色为深灰色或深蓝色")

    # 32 总损失公式文本。逐条踩点：① 区域 44%–50% / 76%–80%；② 去空白后包含
    #    “L=Lrec+λLreg”；③ 含希腊字母“λ”；④ 字号 12–17；⑤ “rec”“reg”真正以下标
    #    格式呈现，即存在 baseline 为负值的 run；⑥ 文本中不含下划线字符“_”。
    #    用“L_rec”这类下划线写法虽然肉眼可读，但不是下标格式，不满足。
    hit = False
    for s in tf("TotalLoss"):
        if not page.at(s, 44.0, 50.0, 76.0, 80.0):
            continue
        flat = _flat(s["tx"])
        if "_" in flat or "λ" not in flat:
            continue
        if "l=lrec+λlreg" not in _squash(flat).lower():
            continue
        if not _rng(_dom_sz(s["tx"]), 12.0, 17.0):
            continue
        if any(_fnum(r.get("base"), 0.0) < 0 for r in _runs(s["tx"])):
            hit = True
    results.append(("+3：" if hit else "+0：") + "页面从左至右44%–50%、从上至下76%–80%出现“Total Loss: L = Lrec + λLreg”文本，包含希腊字母“λ”，字号12–17磅，其中“rec”“reg”以下标格式呈现，文本中不含下划线字符“_”")

    # 33 底部图例框及其内部七项内容。逐条踩点：① 圆角矩形位置 22%–27% / 88%–92%，
    #    宽 18.0–20.0、高 1.3–1.6，白色填充、灰色边线；② 框内自左至右依次为
    #    “Legend:”、一条 1.0–1.4cm 带箭头水平实线、“Data Flow”、一条 1.0–1.4cm
    #    带箭头水平虚线、“Auxiliary Flow”、一个 0.9–1.2×0.4–0.6 的圆角矩形色块、
    #    “Vector / Feature”；③ “Legend:”加粗且字号 11–15；④ 后三处说明文本字号
    #    9–11.5；⑤ 七项的横向中心必须严格自左至右递增。本条按渲染出来的样子判定
    #    实线与虚线，是否由短线段拼出在虚线边框那一条里单独判定。
    lgd = one(_find_rr(shapes, page, 22, 27, 88, 92, (18.0, 20.0), (1.3, 1.6),
                       ln_pred=_is_lightgray, fill_pred=_is_white))
    hit = False
    if lgd is not None:
        ins = [s for s in shapes
               if s.get("x") is not None and s is not lgd
               and s["x"] >= lgd["x"] - 0.1 and _x2(s) <= _x2(lgd) + 0.1
               and s["y"] >= lgd["y"] - 0.35 and _y2(s) <= _y2(lgd) + 0.35]
        ids = set(id(s) for s in ins)
        lines = sorted([r for r in _hruns(ins)
                        if _rng(r["span"], 1.0, 1.4) and r["arrow_right"]],
                       key=lambda r: r["x1"])
        solid = [r for r in lines if _solid_like(r)]
        dashr = [r for r in lines if _dash_like(r)]
        swat = [s for s in ins if s.get("prst") == "roundRect"
                and _rng(s["w"], 0.9, 1.2) and _rng(s["h"], 0.4, 0.6)]

        def _pickt(keys, lo_sz, hi_sz, need_bold=False):
            for z in tf(*keys):
                if id(z) not in ids or not _rng(_dom_sz(z["tx"]), lo_sz, hi_sz):
                    continue
                if need_bold and not _all_bold(z["tx"]):
                    continue
                return z
            return None

        t_lg = _pickt(("Legend:",), 11.0, 15.0, True)
        t_df = _pickt(("DataFlow",), 9.0, 11.5)
        t_af = _pickt(("AuxiliaryFlow",), 9.0, 11.5)
        t_vf = _pickt(("Vector", "Feature"), 9.0, 11.5)
        if solid and dashr and swat and t_lg and t_df and t_af and t_vf:
            seq = [_cx(t_lg), (solid[0]["x1"] + solid[0]["x2"]) / 2.0, _cx(t_df),
                   (dashr[0]["x1"] + dashr[0]["x2"]) / 2.0, _cx(t_af),
                   _cx(swat[0]), _cx(t_vf)]
            hit = all(seq[k] < seq[k + 1] for k in range(6))
    results.append(("+3：" if hit else "+0：") + "页面从左至右22%–27%、从上至下88%–92%出现一个横向圆角矩形，宽18.0–20.0cm，高1.3–1.6cm，填充为白色，边线为灰色，其内部自左至右依次出现“Legend:”文本、一条长度1.0–1.4cm的带右向箭头水平实线、“Data Flow”文本、一条长度1.0–1.4cm的带右向箭头水平虚线、“Auxiliary Flow”文本、一个宽0.9–1.2cm高0.4–0.6cm的圆角矩形、“Vector / Feature”或“Vector/Feature”文本，其中“Legend:”加粗且字号11–15磅，后三处说明文本字号9–11.5磅")

    # 34 七处黑色带箭头连接线及其走向。逐条踩点，任一处缺失或方向反了都不给分：
    #    ①②③ 三处折线箭头，用“横段贴着源矩形右边线引出 + 竖段落在两矩形之间 +
    #    箭头端落在目标矩形左边线上”三要素判定，并用箭头端纵坐标相对源矩形中心
    #    的上下关系区分“向右上”与“向右下”；④ 融合区块内并排矩形之间三条长度
    #    0.5–0.8cm 的水平右向箭头且位于同一水平线上；⑤ 一条长度 1.0–1.6cm 的水平
    #    右向箭头指向绿色框左边线；⑥ 一条长度 0.6–1.0cm 的竖直向下箭头连接两个
    #    绿色小矩形；⑦ 一条长度 1.2–1.6cm 的竖直向下箭头由绿色框下边线指向粉红
    #    色框上边线。箭头方向按 a:ext 的正负号叠加 flipH/flipV 还原，不取绝对值。
    def _elbow(src, dst, down):
        if src is None or dst is None:
            return False
        tipped = False
        for s in blacks:
            tp = _tip(s)
            if tp is None or abs(tp[0] - dst["x"]) > 0.25:
                continue
            if not (dst["y"] - 0.25 <= tp[1] <= _y2(dst) + 0.25):
                continue
            if down and tp[1] <= _cy(src):
                continue
            if not down and tp[1] >= _cy(src):
                continue
            tipped = True
        vert = any(s["w"] <= 0.12 and s["h"] >= 0.2
                   and _x2(src) - 0.3 <= _cx(s) <= dst["x"] + 0.1 for s in blacks)
        horz = any(s["h"] <= 0.12 and s["w"] >= 0.2
                   and abs(s["x"] - _x2(src)) <= 0.8
                   and src["y"] <= _cy(s) <= _y2(src) for s in blacks)
        return tipped and vert and horz

    trio = [s for s in blacks
            if s["h"] <= 0.12 and _rng(s["w"], 0.5, 0.8) and _tip(s) is not None]
    d_ok = any(len(g) >= 3 for g in _clusters([_cy(s) for s in trio], 0.1))
    e_ok = False
    for s in blacks:
        tp = _tip(s)
        if s["h"] <= 0.12 and _rng(s["w"], 1.0, 1.6) and tp is not None \
                and a_ano is not None and abs(tp[0] - a_ano["x"]) <= 0.25:
            e_ok = True

    def _vdown(h_lo, h_hi, y_from, y_to, tol_from, tol_to):
        if y_from is None or y_to is None:
            return False
        for s in blacks:
            if s["w"] > 0.15 or not _rng(s["h"], h_lo, h_hi):
                continue
            tp, tl = _tip(s), _tail_pt(s)
            if tp is None or tl is None or tp[1] <= tl[1]:
                continue
            if abs(tl[1] - y_from) <= tol_from and abs(tp[1] - y_to) <= tol_to:
                return True
        return False

    r_up = rec_r if rec_r is not None else a_rec
    r_dn = scr_r if scr_r is not None else a_scr
    f_ok = _vdown(0.6, 1.0, _y2(r_up) if r_up else None,
                  r_dn["y"] if r_dn else None, 0.35, 0.25)
    g_ok = _vdown(1.2, 1.6, _y2(a_ano) if a_ano else None,
                  a_dec["y"] if a_dec else None, 0.35, 0.35)
    hit = (_elbow(a_teal, a_seq, False) and _elbow(a_teal, a_pat, True)
           and _elbow(a_seq, a_fus, True) and d_ok and e_ok and f_ok and g_ok)
    results.append(("+5：" if hit else "+0：") + "页面出现七处黑色带箭头连接线且走向正确，分别为自左侧青色圆角矩形右侧向右上折入青绿色圆角矩形左侧，自左侧青色圆角矩形右侧向右下折入紫色圆角矩形左侧，自青绿色圆角矩形右侧向右下折入橙色圆角矩形左侧，橙色圆角矩形内四个并排矩形之间的三条水平右向箭头且长度0.5–0.8cm并位于同一水平线上，自橙色圆角矩形右侧一条长度1.0–1.6cm的水平右向箭头指向绿色圆角矩形左侧，“Reconstruction”所在矩形与“Anomaly Score”所在矩形之间一条长度0.6–1.0cm的竖直向下箭头，自绿色圆角矩形下方一条长度1.2–1.6cm的竖直向下箭头指向粉红色圆角矩形")

    # 35 两条虚线折线箭头。紫色一条：整体自紫色框右侧引出（合并外框左端与紫色框
    #    右边线相距≤0.5），纵向跨度≥1.5，箭头端落在橙色框左边线上（横向≤0.3）
    #    且纵向位于橙色框下半部。蓝色一条：整体右端与蓝色框左边线相距≤0.4，
    #    纵向跨度≥1.0，存在一条向上的箭头且箭头端落在紫色框底边（≤0.4）。
    #    本条判定的是走向与连通位置，是否用线型画虚线在最后一条里单独判定。
    #    区块框自身的边框线段（含用短线段拼出来的边框）落在框的边界上，不属于
    #    连接线，按“整体落在框内”剔除；右侧的水平分隔线也是深蓝色，用“整体位于
    #    蓝色框左边线以左”把它排除掉。
    def _on_frame(s, f, pad=0.12):
        if f is None:
            return False
        return (s["x"] >= f["x"] - pad and _x2(s) <= _x2(f) + pad
                and s["y"] >= f["y"] - pad and _y2(s) <= _y2(f) + pad)

    pur = [s for s in conns if _is_purple(s["ln"].get("color"))
           and page.at(s, 33.0, 48.0, 30.0, 58.0) and not _on_frame(s, a_pat)]
    p_ok = False
    if pur and a_fus is not None and a_pat is not None:
        bx, by, bw, bh = _bbox(pur)
        tipped = False
        for s in pur:
            tp = _tip(s)
            if tp is None or abs(tp[0] - a_fus["x"]) > 0.3:
                continue
            if _cy(a_fus) <= tp[1] <= _y2(a_fus) + 0.3:
                tipped = True
        p_ok = tipped and abs(bx - _x2(a_pat)) <= 0.5 and bh >= 1.5
    blu = [s for s in conns if _is_blue(s["ln"].get("color"))
           and page.at(s, 25.0, 50.0, 60.0, 80.0)
           and a_trn is not None and _x2(s) <= a_trn["x"] + 0.15]
    b_ok = False
    if blu and a_trn is not None and a_pat is not None:
        bx, by, bw, bh = _bbox(blu)
        upward = False
        for s in blu:
            tp, tl = _tip(s), _tail_pt(s)
            if tp is None or tl is None or tp[1] >= tl[1]:
                continue
            if abs(tp[1] - _y2(a_pat)) <= 0.4:
                upward = True
        b_ok = upward and abs((bx + bw) - a_trn["x"]) <= 0.4 and bh >= 1.0
    hit = p_ok and b_ok
    results.append(("+3：" if hit else "+0：") + "页面出现两条虚线折线箭头，一条紫色虚线自紫色圆角矩形右上方向右上折行并以右向箭头指向橙色圆角矩形左侧下部，一条蓝色虚线自蓝色圆角矩形左侧向左延伸后向上折行并以向上箭头指向紫色圆角矩形底部")

    # 36 橙色框到蓝色框的黑色折线箭头：全部线段落在 52%–64% / 45%–58%，整体顶端
    #    贴着橙色框下边线（≤0.25），纵向跨度≥1.5，末端向下的箭头落在蓝色框上边线
    #    （≤0.25）。
    mid = [s for s in blacks if page.at(s, 52.0, 64.0, 45.0, 58.0)]
    hit = False
    if mid and a_fus is not None and a_trn is not None:
        bx, by, bw, bh = _bbox(mid)
        dn = False
        for s in mid:
            tp, tl = _tip(s), _tail_pt(s)
            if tp is None or tl is None or tp[1] <= tl[1]:
                continue
            if abs(tp[1] - a_trn["y"]) <= 0.25:
                dn = True
        hit = dn and abs(by - _y2(a_fus)) <= 0.25 and bh >= 1.5
    results.append(("+1：" if hit else "+0：") + "页面从左至右52%–64%、从上至下45%–58%范围内出现一条黑色折线箭头，自上而下连接橙色圆角矩形下边线与蓝色圆角矩形上边线，末端带向下箭头")

    # 37 六处区块标题的排版。逐条踩点：① 六处标题都要能找到；② 每处都必须单行
    #    完整显示，框宽不足折成两行即不满足；③ 不压住其下方圆角矩形的上边线、也
    #    不侵入矩形内部——要么文本末行底端在矩形上边线之上（含 0.15cm 容差），
    #    要么文本整体在上边线之下，正好横跨上边线的情况不满足。
    #    行数按字宽表估算，末行底端按文本框底端取。
    pairs = ((("1.SequenceEncoder",), a_seq), (("2.PatternEncoder",), a_pat),
             (("3.Fusion", "RepresentationLearning"), a_fus),
             (("4.AnomalyScoring",), a_ano), (("5.Decision",), a_dec),
             (("6.TrainingObjective",), a_trn))
    hit = True
    for keys, r in pairs:
        t = one(tf(*keys))
        if t is None or r is None or _est_lines(t) != 1:
            hit = False
            break
        if not (_txt_bottom(t) <= r["y"] + 0.15 or t["y"] >= r["y"] - 0.15):
            hit = False
            break
    results.append(("+1：" if hit else "+0：") + "页面上“1. Sequence Encoder”“2. Pattern Encoder”“3. Fusion & Representation Learning”“4. Anomaly Scoring”“5. Decision”“6. Training Objective”六处文本均在各自文本框内单行完整显示，未因文本框宽度不足折行为两行，未压住其下方圆角矩形的上边线，未侵入矩形内部")

    # 38 虚线边框的实现方式与整页结构。逐条踩点：① 青绿、紫、橙、绿、粉红、蓝六个
    #    区块框以及底部图例框，七处都必须在形状边线上写有虚线线型 a:prstDash；
    #    用一串独立实线短段逐段拼出的“虚线”不满足，因为整条边框无法被整体选中、
    #    移动和调整尺寸；② 单页顶层形状总数不超过 300 个；③ 图标类图形已成组，
    #    即顶层存在不少于 4 个组合对象。短线段拼虚线会同时把顶层形状数量顶到
    #    数百个并使成组数量归零，三个条件相互印证。
    frames = (a_seq, a_pat, a_fus, a_ano, a_dec, a_trn, a_lgd)
    dashed = all(f is not None and (f["ln"].get("dash") or "") in _DASH_PRST
                 for f in frames)
    ngrp = sum(1 for s in tops if s["kind"] == "grpSp")
    hit = dashed and len(tops) <= 300 and ngrp >= 4
    results.append(("+5：" if hit else "+0：") + "页面上青绿色、紫色、橙色、绿色、粉红色、蓝色六处区块框及底部图例框的虚线边框均通过形状边线的虚线线型实现，未由多条独立实线短线段逐段拼接模拟，单页幻灯片顶层形状总数不超过300个，图标类图形已成组，各虚线边框可被整体选中、移动并调整尺寸")

    # 汇总：results 的顺序必须与 _MAX_DELTAS_ORDER 完全一致，逐项反解出得分。
    if len(results) != len(_MAX_DELTAS_ORDER):
        raise RuntimeError(
            "dim2 项数不一致：results=%d vs order=%d" % (len(results), len(_MAX_DELTAS_ORDER)))
    dim2_items = []
    total_score = 0
    for raw, max_delta in zip(results, _MAX_DELTAS_ORDER):
        head, _, rule = raw.partition("：")
        try:
            delta = int(head.replace("+", "").strip())
        except ValueError:
            delta = 0
        hit_item = delta > 0
        dim2_items.append({"rule": rule, "max_delta": max_delta,
                           "delta": max_delta if hit_item else 0,
                           "hit": hit_item})
        total_score += (max_delta if hit_item else 0)
    result["dim2_items"] = dim2_items
    result["total_score"] = total_score
    result["max_score"] = MAX_SCORE_TOTAL
    return result


def evaluate(dir_path):
    """对目录下的交付 PPTX 打分，返回统一约定的结果字典。

    status='error' 只在“这份提交没被真正评过”时出现——目录不存在、目录内没有
    .pptx、或脚本自身抛异常；此时 dim1_pass 为空串（未判定）、dim1_reason 为
    空串、dim2_items 为空表、total_score 与 max_score 一律为 0。
    文件能定位到但门槛不达标，仍按 status='ok' + 维度一未通过处理，此时
    dim1_pass=False 且 dim1_reason 带有未通过的错误摘要。"""
    result = {
        "id": SCRIPT_ID,
        "file_name": "",
        "status": "ok",
        "error": None,
        # 空串表示“维度一尚未判定”；status='error' 的三条出口都保持空串，
        # 只有真正跑过 _check_dim1 才会写入 True / False。
        "dim1_pass": "",
        "dim1_reason": "",
        "dim2_items": [],
        "total_score": 0,
        "max_score": 0,
    }
    try:
        if not dir_path or not os.path.isdir(dir_path):
            result["status"] = "error"
            result["error"] = "目录不存在：%s" % dir_path
            return result
        names = sorted(f for f in os.listdir(dir_path)
                       if f.lower().endswith(".pptx") and not f.startswith("~"))
        if not names:
            result["status"] = "error"
            result["error"] = "目录中未找到 .pptx 文件：%s" % dir_path
            return result
        result["file_name"] = names[0]
        _evaluate_impl(os.path.join(dir_path, names[0]), result)
    except Exception as exc:
        # _evaluate_impl 是就地写 result 的，异常可能发生在它写了一半之后，
        # 因此这里把所有承载评分结论的字段显式复位，保证 error 的返回值是干净的全 0；
        # dim1_pass 复位为空串而不是 False——异常意味着维度一的结论作废，不是判定为不通过。
        result["status"] = "error"
        result["error"] = "%s: %s" % (type(exc).__name__, exc)
        result["dim1_pass"] = ""
        result["dim1_reason"] = ""
        result["dim2_items"] = []
        result["total_score"] = 0
        result["max_score"] = 0
    return result


if __name__ == "__main__":
    _dir = sys.argv[1] if len(sys.argv) >= 2 else os.path.dirname(os.path.abspath(__file__))
    print(json.dumps(evaluate(_dir), ensure_ascii=False, indent=2))
