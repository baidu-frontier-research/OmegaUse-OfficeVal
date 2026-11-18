#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# pyright: reportUnknownMemberType=false, reportUnknownVariableType=false, reportUnknownArgumentType=false, reportMissingImports=false, reportOptionalMemberAccess=false, reportGeneralTypeIssues=false, reportAttributeAccessIssue=false, reportIndexIssue=false, reportArgumentType=false, reportOptionalSubscript=false, reportOptionalIterable=false

# =============================================================================
# 接口约定（与 officeval_008_verifier.py 完全一致）：
#   1. 本脚本不修改全局 sys.stdout，不打印主结果；
#   2. 主结果由 evaluate(dir_path) -> dict 返回；
#   3. 入参为“脚本所在目录的路径”，由脚本自行在该目录下定位并打开被测文件；
#   4. __main__ 仅作为本地调试入口。
#
# 评分对象：口算教学课件（.pptx），第 10 页“乘号 + 双转盘 + 控制按钮”交互页。
# 评分逻辑：维度1（可用或可修改性）为闸门；维度2（完成度）= 25 条加分（满分 63）
#           与 20 条减分（扣分上限 42），完成度 = max(0, 加分合计 - 扣分合计) / 63。
# 所有判定常量均已内联，脚本不依赖原始文件与客户指令。
# =============================================================================

import os
import re
import sys
import json
import math
import shutil
import hashlib
import zipfile
import tempfile
import xml.etree.ElementTree as ET

try:
    from pptx import Presentation
except ImportError:
    Presentation = None

SCRIPT_ID = "065"
MAX_SCORE_TOTAL = 63
MAX_DEDUCT_TOTAL = 42

_MAX_DELTAS_ORDER = [
    # ---------------- 加分项 25 条，合计 63 分 ----------------
    1,   # 01 乘号存在且是可单独选中的对象
    3,   # 02 乘号中心落在 (17.03, 9.71) cm ±0.3
    1,   # 03 乘号高 2.4—4.8cm、亮色、未被遮挡超 20%
    3,   # 04 两个转盘左右分列、离中线≥2.5cm、中心 y 在 3—16cm、直径≥4cm
    3,   # 05 每盘扇形≥6 块、颜色≥4 种、相邻两块不同色
    3,   # 06 每盘纯数字≥6 个、与扇形块数一致、盘内不重复
    1,   # 07 一盘 1—9 一位数，另一盘 10—90 整十数
    3,   # 08 每盘各 1 个指针共 2 个、中心 x 差≤0.3cm、面积≤15%、不随盘转
    1,   # 09 底板包住转盘、面积 1.05—2.5 倍、数字≥12 磅
    3,   # 10 两盘宽差、高差均≤0.1cm
    3,   # 11 两盘中心 y 差≤0.1cm
    3,   # 12 两盘中心 x 的中点与 17.03cm 差≤0.3cm
    1,   # 13 新增对象不出界且两个转盘单元互不重叠
    3,   # 14 有 1 个带文字的按钮并由它触发转盘旋转
    1,   # 15 按钮中心 x 与 17.03cm 差≤0.5cm 且位于转盘中心之下
    1,   # 16 按钮为带填色的圆角矩形类形状、字号≥14 磅、明暗差明显
    5,   # 17 单击按钮两盘同时启动、无延迟
    5,   # 18 无限重复、单圈整 360 度
    5,   # 19 再次单击两盘同时停下、可反复切换
    3,   # 20 暂停停在当前角度（相对量 by）且保持结束状态
    3,   # 21 只由按钮控制，非任意处单击、非自动播放
    3,   # 22 只有两个转盘在动（按钮占位动画≤0.1 秒且 100% 不算）
    3,   # 23 绕圆心旋转：宽高比 0.95—1.05 且圆心偏差≤0.2cm
    1,   # 24 单圈耗时 0.3—1 秒
    1,   # 25 每盘 1 个组合、嵌套≤4 层、非图片或视频承担
    # ---------------- 减分项 20 条，合计 -42 分 ----------------
    -3,  # 26 背景图被删除、移动、改尺寸或不在最底层
    -3,  # 27 多出覆盖页面≥60% 的图片或视频对象
    -1,  # 28 4 个小图标任一被删除、移动或改尺寸
    -3,  # 29 椭圆区域被单个新增对象盖住过半
    -1,  # 30 残留“同步/双转盘/示例/demo/水印/模板/参考图”或 >15% 无文字纯色块
    -5,  # 31 页数≠11 或页面尺寸≠33.87×19.05cm
    -5,  # 32 其余 10 页文字指纹不符
    -3,  # 33 任意页出现自动缩小字号或压缩行距
    -1,  # 34 存在隐藏页或切换效果
    -1,  # 35 母版≠2 或版式≠3 或主题字体≠Microsoft YaHei
    -1,  # 36 备注缺少以“课堂备注：”开头且长度 40—90 字的段落
    -3,  # 37 含宏代码
    -3,  # 38 存在外部链接的媒体
    -1,  # 39 体积 >18MB 或存在冗余未引用素材
    -1,  # 40 文件名≠口算教学课件.pptx
    -1,  # 41 第 10 页存在空形状、页外对象或重复对象
    -3,  # 42 第 1—9、11 页被加了动画
    -1,  # 43 第 10 页存在音频或视频对象
    -1,  # 44 第 10 页存在单击跳转超链接
    -1,  # 45 文档属性被清空或改错
]

assert sum(d for d in _MAX_DELTAS_ORDER if d > 0) == MAX_SCORE_TOTAL
assert -sum(d for d in _MAX_DELTAS_ORDER if d < 0) == MAX_DEDUCT_TOTAL
assert len(_MAX_DELTAS_ORDER) == 45

# --------------------------------------------------------------------------
# 内联常量（全部为绝对值，不依赖原始文件）
# --------------------------------------------------------------------------
EMU_PER_CM = 360000.0
PAGE_W_CM = 33.87                 # 页宽（p:sldSz cx=12192000 EMU）
PAGE_H_CM = 19.05                 # 页高（p:sldSz cy=6858000 EMU）
SLIDE_COUNT = 11                  # 应有页数
TARGET_INDEX = 10                 # 交互页页码（1 基）
CENTER_X_CM = 17.03               # 页面竖向中线（乘号 / 按钮的对齐基准）
CROSS_CENTER_CM = (17.03, 9.71)   # 乘号应在的中心点
ELLIPSE_RECT_CM = (11.54, 3.73, 22.52, 15.68)   # 原页粉色椭圆区域 left/top/right/bottom
FILE_NAME_EXPECT = "口算教学课件.pptx"
SIZE_LIMIT_BYTES = 18 * 1024 * 1024
THEME_FONT_EXPECT = "Microsoft YaHei"
MASTER_COUNT = 2
LAYOUT_COUNT = 3
DOC_TITLE_EXPECT = "10乘两位数、整十数乘整十数的口算教学课件"
DOC_SUBJECT_EXPECT = "小学数学口算教学课件"
NOTE_PREFIX = "课堂备注："
NOTE_LEN_RANGE = (40, 90)
BANNED_WORDS = ("同步", "双转盘", "示例", "demo", "水印", "模板", "参考图")

# 第 10 页 4 个小图标基线：left/top/边长（cm），容差 ±0.30cm
ICONS_CM = (
    (6.70, 4.59, 1.13),
    (25.75, 4.94, 0.85),
    (25.40, 10.58, 0.99),
    (7.06, 10.94, 0.71),
)

# 第 1—9、11 页文字指纹（SHA-1）。归一化方式：递归遍历 spTree（含组合内部），
# 每个形状按段落用 \n 连接其 a:t 文本并顺序拼接全部形状，再删除所有空白
# （含 　 与不换行空格），对 UTF-8 字节取 SHA-1。
TEXT_SHA1 = {
    1: "788f69604ac64fbb49743a67f324e680d5ec81fb",
    2: "326d32698d9a69096344189cf608895f357aa386",
    3: "5b49e6ac31ab3c98604ed295dfa08708b1c419f6",
    4: "ac91cfec311feb0dfa7281b23539c7c343aa1be2",
    5: "c7a818d582ffa7282d43b32062b47fed96b302b3",
    6: "c7a8c99e5f86b2cbab01bd00f75da60686ac7453",
    7: "81a10cae5ec1d05528b2831805bf84d01ad29d23",
    8: "ff3f7c3b216e6d46a4d0fb220236e92ee69854cc",
    9: "2575c3442dc9cb7c8442534c0e4ec8bf8fc567c1",
    11: "2e53dec0ec7ad6f00a7a2a1cc7ee4d9cd6fa92a9",
}

NS = {
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "cp": "http://schemas.openxmlformats.org/package/2006/metadata/core-properties",
    "dc": "http://purl.org/dc/elements/1.1/",
}
_A = NS["a"]
_P = NS["p"]

# --------------------------------------------------------------------------
# 基础工具
# --------------------------------------------------------------------------
def _cm(emu):
    """EMU -> 厘米。"""
    try:
        return float(emu) / EMU_PER_CM
    except (TypeError, ValueError):
        return 0.0


def _xml(z, name):
    """读取包内 XML 部件，缺失或损坏时返回 None。"""
    try:
        return ET.fromstring(z.read(name))
    except (KeyError, ET.ParseError):
        return None


def _rels(z, part_name):
    """返回某部件的 rels 映射：rId -> (Target, TargetMode)。"""
    d = os.path.dirname(part_name)
    rel_name = "%s/_rels/%s.rels" % (d, os.path.basename(part_name))
    root = _xml(z, rel_name)
    out = {}
    if root is None:
        return out
    for rel in root:
        out[rel.get("Id")] = (rel.get("Target") or "", rel.get("TargetMode") or "Internal")
    return out


def _slide_parts(z):
    """按 p:sldIdLst 的顺序返回幻灯片部件名列表（真实放映顺序）。"""
    pres = _xml(z, "ppt/presentation.xml")
    if pres is None:
        return []
    rels = _rels(z, "ppt/presentation.xml")
    out = []
    lst = pres.find("p:sldIdLst", NS)
    if lst is None:
        return out
    for sld in lst.findall("p:sldId", NS):
        rid = sld.get("{%s}id" % NS["r"])
        target = rels.get(rid, ("", ""))[0]
        if not target:
            continue
        target = target.replace("../", "")
        if not target.startswith("ppt/"):
            target = "ppt/" + target
        out.append(target)
    return out


def _shape_text(el):
    """形状内文本：按段落用 \\n 连接（只取 a:t）。"""
    body = el.find("p:txBody", NS)
    if body is None:
        return ""
    paras = []
    for para in body.findall("a:p", NS):
        paras.append("".join(t.text or "" for t in para.iter("{%s}t" % _A)))
    return "\n".join(paras)


def _strip_ws(s):
    """删除所有空白（含全角空格与不换行空格）。"""
    return re.sub(r"[\s　 ﻿]", "", s or "")


def _text_sha1(sp_tree):
    """整页文字指纹：递归拼接所有形状文本后去空白取 SHA-1。"""
    chunks = []

    def walk(node):
        for ch in list(node):
            tag = _local(ch)
            if tag == "grpSp":
                walk(ch)
            elif tag in ("sp", "pic", "graphicFrame", "cxnSp"):
                chunks.append(_shape_text(ch))
                for tbl_t in ch.iter("{%s}t" % _A):
                    if ch.find("p:txBody", NS) is None:
                        chunks.append(tbl_t.text or "")

    if sp_tree is not None:
        walk(sp_tree)
    raw = _strip_ws("".join(chunks))
    return hashlib.sha1(raw.encode("utf-8")).hexdigest(), len(raw)


def _local(el):
    """去掉命名空间的标签名。"""
    tag = el.tag
    return tag.rsplit("}", 1)[-1] if "}" in tag else tag


def _cnv_pr(el):
    """返回 (id, name)。"""
    for path in ("p:nvSpPr/p:cNvPr", "p:nvGrpSpPr/p:cNvPr", "p:nvPicPr/p:cNvPr",
                 "p:nvCxnSpPr/p:cNvPr", "p:nvGraphicFramePr/p:cNvPr"):
        c = el.find(path, NS)
        if c is not None:
            return c.get("id") or "", c.get("name") or ""
    return "", ""


def _attr_num(el, name, default=0.0):
    """读取 OOXML 数值属性。

    部分生成工具会写出非法值（如 cy="D"），此时按 default 处理，
    避免整脚本因单个坏属性而中断评估。
    """
    if el is None:
        return default
    raw = el.get(name)
    if raw is None:
        return default
    try:
        return float(raw)
    except (TypeError, ValueError):
        return default


def _collect(sp_tree):
    """
    递归收集形状，返回记录列表。每条记录带页面绝对坐标（cm）：
      tag/sid/name/x/y/w/h/depth/el/parent/z
    组合内部子元素坐标按 a:off/a:ext 与 a:chOff/a:chExt 做线性映射还原到页面坐标，
    这与办公软件里“组合内对象的实际显示位置”一致。
    """
    out = []
    if sp_tree is None:
        return out

    def walk(node, tf, depth, parent, top_z):
        ax, ay, sx, sy = tf
        for ch in list(node):
            tag = _local(ch)
            if tag in ("nvGrpSpPr", "grpSpPr", "extLst"):
                continue
            if depth == 0:
                top_z = len(out)
            if tag == "grpSp":
                xf = ch.find("p:grpSpPr/a:xfrm", NS)
                if xf is None:
                    continue
                off, ext = xf.find("a:off", NS), xf.find("a:ext", NS)
                choff, chext = xf.find("a:chOff", NS), xf.find("a:chExt", NS)
                if off is None or ext is None:
                    continue
                ox, oy = _attr_num(off, "x"), _attr_num(off, "y")
                cx, cy = _attr_num(ext, "cx"), _attr_num(ext, "cy")
                gx = _attr_num(choff, "x") if choff is not None else 0.0
                gy = _attr_num(choff, "y") if choff is not None else 0.0
                gcx = _attr_num(chext, "cx", cx) if chext is not None else cx
                gcy = _attr_num(chext, "cy", cy) if chext is not None else cy
                bx, by = ax + ox * sx, ay + oy * sy
                sid, name = _cnv_pr(ch)
                out.append({"tag": tag, "sid": sid, "name": name, "el": ch,
                            "x": _cm(bx), "y": _cm(by), "w": _cm(cx * sx), "h": _cm(cy * sy),
                            "depth": depth, "parent": parent, "z": top_z})
                nsx = (cx / gcx * sx) if gcx else sx
                nsy = (cy / gcy * sy) if gcy else sy
                walk(ch, (bx - gx * nsx, by - gy * nsy, nsx, nsy), depth + 1, sid, top_z)
            elif tag in ("sp", "pic", "cxnSp", "graphicFrame"):
                xf = ch.find(".//a:xfrm", NS) if tag != "graphicFrame" else ch.find("p:xfrm", NS)
                if xf is None:
                    continue
                off, ext = xf.find("a:off", NS), xf.find("a:ext", NS)
                if off is None or ext is None:
                    continue
                sid, name = _cnv_pr(ch)
                out.append({"tag": tag, "sid": sid, "name": name, "el": ch,
                            "x": _cm(ax + _attr_num(off, "x") * sx),
                            "y": _cm(ay + _attr_num(off, "y") * sy),
                            "w": _cm(_attr_num(ext, "cx") * sx),
                            "h": _cm(_attr_num(ext, "cy") * sy),
                            "depth": depth, "parent": parent, "z": top_z})

    walk(sp_tree, (0.0, 0.0, 1.0, 1.0), 0, None, -1)
    return out


def _rect(s):
    return (s["x"], s["y"], s["x"] + s["w"], s["y"] + s["h"])


def _center(s):
    return (s["x"] + s["w"] / 2.0, s["y"] + s["h"] / 2.0)


def _area(s):
    return max(0.0, s["w"]) * max(0.0, s["h"])


def _inter_area(r1, r2):
    dx = min(r1[2], r2[2]) - max(r1[0], r2[0])
    dy = min(r1[3], r2[3]) - max(r1[1], r2[1])
    return dx * dy if dx > 0 and dy > 0 else 0.0


def _contains(outer, inner, tol=0.05):
    return (outer[0] <= inner[0] + tol and outer[1] <= inner[1] + tol
            and outer[2] >= inner[2] - tol and outer[3] >= inner[3] - tol)


def _descendants(shapes, sid):
    """某组合下的全部后代记录。"""
    kids = [s for s in shapes if s["parent"] == sid]
    out = list(kids)
    for k in kids:
        if k["tag"] == "grpSp":
            out += _descendants(shapes, k["sid"])
    return out


def _max_depth(shapes, sid, base=1):
    """组合嵌套层数：自身算 1 层。"""
    depth = base
    for k in [s for s in shapes if s["parent"] == sid and s["tag"] == "grpSp"]:
        depth = max(depth, _max_depth(shapes, k["sid"], base + 1))
    return depth


def _theme_palette(z):
    """主题 clrScheme + 母版 clrMap，用于把 schemeClr 解析成 srgb 十六进制。"""
    palette, clrmap = {}, {}
    theme = _xml(z, "ppt/theme/theme1.xml")
    if theme is not None:
        scheme = theme.find(".//a:clrScheme", NS)
        for node in list(scheme) if scheme is not None else []:
            key = _local(node)
            srgb = node.find("a:srgbClr", NS)
            sysc = node.find("a:sysClr", NS)
            if srgb is not None:
                palette[key] = (srgb.get("val") or "").upper()
            elif sysc is not None:
                palette[key] = (sysc.get("lastClr") or "").upper()
    master = _xml(z, "ppt/slideMasters/slideMaster1.xml")
    if master is not None:
        cm = master.find("p:clrMap", NS)
        if cm is not None:
            clrmap = dict(cm.attrib)
    return palette, clrmap


def _hex_of(node, palette, clrmap):
    """把 a:srgbClr / a:schemeClr / a:sysClr 解析为 RRGGBB。"""
    if node is None:
        return None
    tag = _local(node)
    if tag == "srgbClr":
        return (node.get("val") or "").upper() or None
    if tag == "sysClr":
        return (node.get("lastClr") or "").upper() or None
    if tag == "schemeClr":
        val = node.get("val") or ""
        val = clrmap.get(val, val)
        if val in ("phClr",):
            return None
        return palette.get(val)
    return None


def _first_color(parent, palette, clrmap):
    for name in ("srgbClr", "schemeClr", "sysClr"):
        node = parent.find("a:%s" % name, NS)
        if node is not None:
            return _hex_of(node, palette, clrmap)
    return None


def _fill_hex(el, palette, clrmap):
    """
    形状填充色：solidFill 直接取；gradFill 取各 gs 色的平均值；
    noFill / 无 spPr 填充返回 None（与办公软件“无填充”一致）。
    """
    spPr = el.find("p:spPr", NS)
    if spPr is None:
        return None
    if spPr.find("a:noFill", NS) is not None:
        return None
    solid = spPr.find("a:solidFill", NS)
    if solid is not None:
        return _first_color(solid, palette, clrmap)
    grad = spPr.find("a:gradFill", NS)
    if grad is not None:
        cols = []
        for gs in grad.findall("a:gsLst/a:gs", NS):
            hx = _first_color(gs, palette, clrmap)
            if hx:
                cols.append(hx)
        if cols:
            r = sum(int(c[0:2], 16) for c in cols) // len(cols)
            g = sum(int(c[2:4], 16) for c in cols) // len(cols)
            b = sum(int(c[4:6], 16) for c in cols) // len(cols)
            return "%02X%02X%02X" % (r, g, b)
    return None


def _lum(hx):
    """相对亮度 0—1。"""
    if not hx or len(hx) < 6:
        return None
    try:
        r, g, b = int(hx[0:2], 16), int(hx[2:4], 16), int(hx[4:6], 16)
    except ValueError:
        return None
    return (0.2126 * r + 0.7152 * g + 0.0722 * b) / 255.0


def _font_pts(el, default=18.0):
    """
    形状内每个文本 run 的字号（磅）。rPr@sz 未写时取演示文稿默认正文 18 磅，
    与办公软件“新建文本框默认 18 号”一致。
    """
    out = []
    body = el.find("p:txBody", NS)
    if body is None:
        return out
    for para in body.findall("a:p", NS):
        p_sz = None
        pPr = para.find("a:pPr", NS)
        if pPr is not None:
            d = pPr.find("a:defRPr", NS)
            if d is not None and d.get("sz"):
                p_sz = float(d.get("sz")) / 100.0
        for run in para.findall("a:r", NS):
            txt = "".join(t.text or "" for t in run.iter("{%s}t" % _A))
            if not txt.strip():
                continue
            rPr = run.find("a:rPr", NS)
            sz = None
            if rPr is not None and rPr.get("sz"):
                sz = float(rPr.get("sz")) / 100.0
            out.append((txt, sz if sz else (p_sz if p_sz else default)))
    return out


def _text_hex(el, palette, clrmap):
    """文字颜色：优先 run 的 solidFill，其次 p:style/a:fontRef 的主题色。"""
    body = el.find("p:txBody", NS)
    if body is not None:
        for run in body.iter("{%s}r" % _A):
            rPr = run.find("a:rPr", NS)
            if rPr is None:
                continue
            solid = rPr.find("a:solidFill", NS)
            if solid is not None:
                hx = _first_color(solid, palette, clrmap)
                if hx:
                    return hx
    ref = el.find("p:style/a:fontRef", NS)
    if ref is not None:
        hx = _first_color(ref, palette, clrmap)
        if hx:
            return hx
    return "000000"


def _prst(el):
    g = el.find("p:spPr/a:prstGeom", NS)
    return (g.get("prst") or "") if g is not None else ""


def _has_custgeom(el):
    return el.find("p:spPr/a:custGeom", NS) is not None


_BEHAVIOR_TAGS = ("anim", "animClr", "animEffect", "animMotion", "animRot",
                  "animScale", "set", "cmd", "audio", "video")


def _parse_timing(root):
    """
    解析 p:timing。返回：
      present / has_main_seq / seqs / behaviors / onclick_spids
    behaviors 每项：kind, spid, dur, by, to, eff_nodeType, eff_repeat, eff_fill,
                    eff_delay, in_seq_restart, scale_by
    这些字段对应办公软件“动画窗格”里的：触发方式、开始（单击时/与上一动画同时）、
    持续时间、重复、播完后保持结束状态。
    """
    info = {"present": False, "has_main_seq": False, "seqs": [],
            "behaviors": [], "onclick_spids": []}
    timing = root.find("p:timing", NS)
    if timing is None:
        return info
    info["present"] = True
    parent = {}
    for node in timing.iter():
        for ch in list(node):
            parent[ch] = node

    def spid_of(cond):
        t = cond.find("p:tgtEl/p:spTgt", NS)
        return t.get("spid") if t is not None else None

    for cond in timing.iter("{%s}cond" % _P):
        if cond.get("evt") == "onClick":
            up = parent.get(cond)
            if up is not None and _local(up) == "stCondLst":
                info["onclick_spids"].append(spid_of(cond))

    for seq in timing.iter("{%s}seq" % _P):
        ctn = seq.find("p:cTn", NS)
        trig, nxt = [], []
        for cond in (ctn.findall("p:stCondLst/p:cond", NS) if ctn is not None else []):
            if cond.get("evt") == "onClick":
                trig.append(spid_of(cond))
        for cond in seq.findall("p:nextCondLst/p:cond", NS):
            if cond.get("evt") == "onClick":
                nxt.append(spid_of(cond))
        indefinite = False
        for cond in seq.iter("{%s}cond" % _P):
            up = parent.get(cond)
            if (cond.get("delay") == "indefinite" and cond.get("evt") is None
                    and up is not None and _local(up) == "stCondLst"):
                indefinite = True
        node_type = ctn.get("nodeType") if ctn is not None else None
        if node_type == "mainSeq":
            info["has_main_seq"] = True
        info["seqs"].append({
            "el": seq, "node_type": node_type,
            "restart": ctn.get("restart") if ctn is not None else None,
            "concurrent": seq.get("concurrent"), "next_ac": seq.get("nextAc"),
            "trigger_spids": trig, "next_spids": nxt,
            "has_indefinite_start": indefinite,
        })

    for kind in _BEHAVIOR_TAGS:
        for beh in timing.iter("{%s}%s" % (_P, kind)):
            cbhvr = beh.find("p:cBhvr", NS)
            inner = cbhvr.find("p:cTn", NS) if cbhvr is not None else None
            tgt = cbhvr.find("p:tgtEl/p:spTgt", NS) if cbhvr is not None else None
            by_node = beh.find("p:by", NS)
            scale_by = None
            if by_node is not None and by_node.get("x") is not None:
                scale_by = (by_node.get("x"), by_node.get("y"))
            eff = {"node_type": None, "repeat": None, "fill": None, "delay": None,
                   "end_cond": False, "seq_restart": None}
            up = parent.get(beh)
            while up is not None:
                if _local(up) == "cTn" and up.get("nodeType"):
                    eff["node_type"] = up.get("nodeType")
                    eff["repeat"] = up.get("repeatCount")
                    eff["fill"] = up.get("fill")
                    cond = up.find("p:stCondLst/p:cond", NS)
                    eff["delay"] = cond.get("delay") if cond is not None else None
                    eff["end_cond"] = up.find("p:endCondLst", NS) is not None
                    break
                up = parent.get(up)
            up = parent.get(beh)
            while up is not None:
                if _local(up) == "seq":
                    c = up.find("p:cTn", NS)
                    eff["seq_restart"] = c.get("restart") if c is not None else None
                    break
                up = parent.get(up)
            info["behaviors"].append({
                "kind": kind,
                "spid": tgt.get("spid") if tgt is not None else None,
                "dur": inner.get("dur") if inner is not None else None,
                "inner_fill": inner.get("fill") if inner is not None else None,
                "by": beh.get("by"), "to": beh.get("to"),
                "scale_by": scale_by,
                "attr": ",".join(a.text or "" for a in beh.iter("{%s}attrName" % _P)),
                "eff": eff,
            })

    return info


def _wheel_profile(shapes, wheel, palette, clrmap):
    """统计某个转盘候选内部的扇形块与纯数字文本。"""
    kids = _descendants(shapes, wheel["sid"]) if wheel["tag"] == "grpSp" else []
    sectors, numbers = [], []
    for k in kids:
        if k["tag"] != "sp":
            continue
        el = k["el"]
        is_sector = _has_custgeom(el) or _prst(el) in ("pie", "chord", "blockArc", "arc")
        if is_sector:
            hx = _fill_hex(el, palette, clrmap)
            lum = _lum(hx)
            # 近白色（亮度≥0.95）的窄块视为扇形之间的分隔线，不计为扇形
            if hx and lum is not None and lum < 0.95:
                sectors.append({"hex": hx, "center": _center(k)})
        txt = _shape_text(el).strip()
        if txt.isdigit():
            numbers.append({"num": int(txt), "pts": [p for _, p in _font_pts(el)]})
    return sectors, numbers


def _identify(shapes, palette, clrmap, timing):
    """识别第 10 页的各类角色对象。"""
    top = [s for s in shapes if s["depth"] == 0]
    page_area = PAGE_W_CM * PAGE_H_CM
    pics = [s for s in top if s["tag"] == "pic"]
    bg = None
    for s in pics:
        if (abs(s["x"]) <= 0.15 and abs(s["y"]) <= 0.15
                and abs(s["w"] - PAGE_W_CM) <= 0.35 and abs(s["h"] - PAGE_H_CM) <= 0.35):
            bg = s
            break
    icons = [s for s in pics if s is not bg]

    cross = None
    for s in top:
        if s["tag"] != "sp":
            continue
        if _prst(s["el"]) == "mathMultiply" or _shape_text(s["el"]).strip() in ("×", "✕", "x", "X", "*"):
            cross = s
            break

    cands = []
    for s in shapes:
        if s["tag"] not in ("grpSp", "sp"):
            continue
        if min(s["w"], s["h"]) < 4.0:
            continue
        ratio = (s["w"] / s["h"]) if s["h"] else 0.0
        if not (0.90 <= ratio <= 1.11):
            continue
        sec, num = _wheel_profile(shapes, s, palette, clrmap)
        if len(sec) >= 6 or len(num) >= 6:
            cands.append(s)
    cand_sids = set(c["sid"] for c in cands)

    def has_cand_ancestor(s):
        p = s["parent"]
        guard = 0
        while p and guard < 20:
            if p in cand_sids:
                return True
            up = [q for q in shapes if q["sid"] == p]
            p = up[0]["parent"] if up else None
            guard += 1
        return False

    cands = [c for c in cands if not has_cand_ancestor(c)]
    cands.sort(key=_area, reverse=True)
    wheels = sorted(cands[:2], key=lambda s: _center(s)[0])

    wheel_sids = set(w["sid"] for w in wheels)
    inside_wheel = set()
    for w in wheels:
        for k in _descendants(shapes, w["sid"]):
            inside_wheel.add(k["sid"])

    # 底板：顶层、有填充、完整包住某个转盘、面积为转盘的 1.05—2.5 倍
    boards = []
    for w in wheels:
        pick = None
        for s in top:
            if s["tag"] != "sp" or s["sid"] in wheel_sids or (cross and s is cross):
                continue
            if _fill_hex(s["el"], palette, clrmap) is None:
                continue
            if not _contains(_rect(s), _rect(w)):
                continue
            ratio = _area(s) / _area(w) if _area(w) else 0.0
            if 1.05 <= ratio <= 2.5:
                if pick is None or _area(s) < _area(pick):
                    pick = s
        boards.append(pick)
    board_sids = set(b["sid"] for b in boards if b)

    # 按钮：优先取动画触发对象；否则取带文字、面积 < 页面 20% 的顶层形状
    trig = set(x for seq in timing["seqs"] for x in seq["trigger_spids"] if x)
    button = None
    for s in top:
        if s["tag"] != "sp" or s["sid"] in wheel_sids or s["sid"] in board_sids:
            continue
        if cross and s is cross:
            continue
        if not _shape_text(s["el"]).strip():
            continue
        if _area(s) > 0.20 * page_area:
            continue
        if s["sid"] in trig:
            button = s
            break
        if button is None:
            button = s

    # 指针：顶层、不在转盘组内、面积≤转盘 15%、中心 x 与转盘中心对齐、位于转盘中心之上
    pointers = []
    for w in wheels:
        wcx, wcy = _center(w)
        pick = None
        for s in top:
            if s["tag"] not in ("sp", "pic", "cxnSp"):
                continue
            if s["sid"] in wheel_sids or s["sid"] in board_sids or s["sid"] in inside_wheel:
                continue
            if (cross and s is cross) or (button and s is button):
                continue
            if _area(w) and _area(s) > 0.15 * _area(w):
                continue
            cx, cy = _center(s)
            if abs(cx - wcx) <= 0.3 and cy < wcy:
                if pick is None or abs(_center(pick)[0] - wcx) > abs(cx - wcx):
                    pick = s
        pointers.append(pick)

    return {"bg": bg, "bg_first": bool(bg) and bg["z"] == 0, "icons": icons,
            "cross": cross, "wheels": wheels, "boards": boards,
            "pointers": pointers, "button": button, "page_area": page_area}


def _open_presentation(file_path):
    """
    打开演示文稿包。返回 (zipfile, dim1_reason)；reason 非空表示维度1 不通过。
    仅做“能否打开、是不是 .pptx 演示文稿、是否被加密或锁定”这类格式层判定，
    不涉及文件里的内容。
    """
    if not file_path.lower().endswith(".pptx"):
        return None, "交付物不是 .pptx 演示文稿：%s" % os.path.basename(file_path)
    try:
        z = zipfile.ZipFile(file_path)
    except zipfile.BadZipFile:
        return None, "文件损坏或被加了打开密码，无法作为 .pptx 打开"
    except Exception as exc:
        return None, "文件无法打开：%s" % exc
    names = set(z.namelist())
    if "ppt/presentation.xml" not in names:
        z.close()
        return None, "包内缺少 ppt/presentation.xml，不是可编辑的 .pptx 演示文稿"
    ct = _xml(z, "[Content_Types].xml")
    if ct is not None:
        macro_ct = any("macroEnabled.main+xml" in (o.get("ContentType") or "")
                       for o in ct.iter() if o.get("ContentType"))
        if macro_ct:
            z.close()
            return None, "含宏代码却存成 .pptx，打开时会报错"
    custom = _xml(z, "docProps/custom.xml")
    if custom is not None:
        for prop in custom:
            if (prop.get("name") or "") == "_MarkAsFinal":
                vals = [(v.text or "").strip().lower() for v in list(prop)]
                if "true" in vals or "1" in vals:
                    z.close()
                    return None, "文件被设为“标记为最终状态”，不可继续编辑"
    if os.path.exists(file_path) and not os.access(file_path, os.W_OK):
        z.close()
        return None, "文件为只读，不可继续编辑"
    return z, ""


def evaluate(dir_path):
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
                       if f.lower().endswith(".pptx") and not f.startswith("~$"))
        if not cands:
            raise FileNotFoundError("目录内未找到 .pptx 文件：%s" % dir_path)
        file_path = os.path.join(dir_path, cands[0])
        result["file_name"] = os.path.basename(file_path)
        return _evaluate_impl(file_path, result)
    except Exception as exc:
        # status 为 error 时未能完成判定，dim1_pass 留空，dim1_reason 也为空
        result["status"] = "error"
        result["error"] = "%s: %s" % (type(exc).__name__, exc)
        result["dim1_pass"] = ""
        result["dim1_reason"] = ""
        result["dim2_items"] = []
        result["total_score"] = 0
        return result


def _evaluate_impl(file_path, result):
    results = []

    # ---------------- 维度1：可用或可修改性（仅格式层闸门） ----------------
    z, reason = _open_presentation(file_path)
    if reason:
        result["dim1_pass"] = False
        result["dim1_reason"] = reason
        result["dim2_items"] = []
        result["total_score"] = 0
        result["max_score"] = MAX_SCORE_TOTAL
        return result
    result["dim1_pass"] = True
    # dim1_reason 只记录"维度一不通过的原因"，通过时保持为空字符串。
    result["dim1_reason"] = ""

    try:
        names = z.namelist()
        pres = _xml(z, "ppt/presentation.xml")
        sld_sz = pres.find("p:sldSz", NS) if pres is not None else None
        page_w = _cm(sld_sz.get("cx")) if sld_sz is not None else 0.0
        page_h = _cm(sld_sz.get("cy")) if sld_sz is not None else 0.0
        slide_parts = _slide_parts(z)
        slide_roots = [_xml(z, n) for n in slide_parts]
        palette, clrmap = _theme_palette(z)

        idx10 = TARGET_INDEX - 1
        root10 = slide_roots[idx10] if len(slide_roots) > idx10 else None
        if root10 is None and slide_roots:
            root10 = slide_roots[-1]
        tree10 = root10.find("p:cSld/p:spTree", NS) if root10 is not None else None
        shapes = _collect(tree10)
        top10 = [s for s in shapes if s["depth"] == 0]
        timing = _parse_timing(root10) if root10 is not None else _parse_timing(ET.Element("x"))
        ident = _identify(shapes, palette, clrmap, timing)
        wheels = ident["wheels"]
        boards = ident["boards"]
        pointers = ident["pointers"]
        cross = ident["cross"]
        button = ident["button"]
        page_area = PAGE_W_CM * PAGE_H_CM
        rots = [b for b in timing["behaviors"] if b["kind"] == "animRot"]
        wheel_sids = [w["sid"] for w in wheels]

        # 01 +1 乘号必须是能单独点选的对象（形状/文本框/图片/组合），画在背景像素里不得分
        results.append(("+1：第10页存在一个可单独选中的乘号对象"
                        if cross is not None else "+0：第10页存在一个可单独选中的乘号对象"))

        # 02 +3 乘号中心落在 (17.03, 9.71) cm，容差 0.3cm（对应“对象位置/大小”对话框）
        ok = False
        if cross is not None:
            ccx, ccy = _center(cross)
            ok = (abs(ccx - CROSS_CENTER_CM[0]) <= 0.3 and abs(ccy - CROSS_CENTER_CM[1]) <= 0.3)
        results.append(("+3：乘号中心位于横向17.03厘米、纵向9.71厘米处，偏差不超过0.3厘米"
                        if ok else "+0：乘号中心位于横向17.03厘米、纵向9.71厘米处，偏差不超过0.3厘米"))

        # 03 +1 乘号高 2.4—4.8cm、填充为亮色（亮度≥0.6）、被上层对象遮挡不超过 20%
        ok = False
        if cross is not None:
            lum = _lum(_fill_hex(cross["el"], palette, clrmap))
            cover = 0.0
            for s in top10:
                if s is cross or s["z"] <= cross["z"]:
                    continue
                cover += _inter_area(_rect(cross), _rect(s))
            ok = (2.4 <= cross["h"] <= 4.8 and lum is not None and lum >= 0.6
                  and (cover / _area(cross) if _area(cross) else 1.0) <= 0.20)
        results.append(("+1：乘号高度在2.4至4.8厘米之间、填充为亮色，且被其他对象遮挡的面积不超过两成"
                        if ok else "+0：乘号高度在2.4至4.8厘米之间、填充为亮色，且被其他对象遮挡的面积不超过两成"))

        # 04 +3 两个转盘一左一右、中心离 17.03cm 竖线≥2.5cm、中心 y 在 3—16cm、直径≥4cm
        ok = False
        if len(wheels) == 2:
            c1, c2 = _center(wheels[0]), _center(wheels[1])
            ok = (c1[0] < CENTER_X_CM < c2[0]
                  and abs(c1[0] - CENTER_X_CM) >= 2.5 and abs(c2[0] - CENTER_X_CM) >= 2.5
                  and all(3.0 <= c[1] <= 16.0 for c in (c1, c2))
                  and all(min(w["w"], w["h"]) >= 4.0 for w in wheels))
        results.append(("+3：页面上有两个转盘，一个在左、一个在右，中心距横向17.03厘米竖线各不少于2.5厘米，中心纵向都在3至16厘米之间，直径都不小于4厘米"
                        if ok else "+0：页面上有两个转盘，一个在左、一个在右，中心距横向17.03厘米竖线各不少于2.5厘米，中心纵向都在3至16厘米之间，直径都不小于4厘米"))

        # 05 +3 每盘扇形≥6 块、颜色≥4 种、按角度排序后相邻两块不同色
        ok = len(wheels) == 2
        for w in wheels:
            sec, _num = _wheel_profile(shapes, w, palette, clrmap)
            if len(sec) < 6 or len(set(s["hex"] for s in sec)) < 4:
                ok = False
                break
            wc = _center(w)
            order = sorted(sec, key=lambda s: math.atan2(s["center"][1] - wc[1],
                                                         s["center"][0] - wc[0]))
            for i in range(len(order)):
                if order[i]["hex"] == order[(i + 1) % len(order)]["hex"]:
                    ok = False
                    break
        results.append(("+3：每个转盘都被分成不少于6块扇形，用了不少于4种颜色，且相邻两块颜色不同"
                        if ok else "+0：每个转盘都被分成不少于6块扇形，用了不少于4种颜色，且相邻两块颜色不同"))

        # 06 +3 每盘纯数字文本≥6 个、个数与扇形块数一致、盘内不重复
        ok = len(wheels) == 2
        wheel_nums = []
        for w in wheels:
            sec, num = _wheel_profile(shapes, w, palette, clrmap)
            vals = [n["num"] for n in num]
            wheel_nums.append(num)
            if len(vals) < 6 or len(vals) != len(sec) or len(set(vals)) != len(vals):
                ok = False
        results.append(("+3：每个转盘的扇形里各写着一个数字，数量不少于6个、与扇形块数一一对应且同盘内不重复"
                        if ok else "+0：每个转盘的扇形里各写着一个数字，数量不少于6个、与扇形块数一一对应且同盘内不重复"))

        # 07 +1 一盘全是 1—9 的一位数，另一盘全是 10—90 的整十数
        ok = False
        if len(wheel_nums) == 2 and all(wheel_nums):
            a = [n["num"] for n in wheel_nums[0]]
            b = [n["num"] for n in wheel_nums[1]]
            single = lambda v: all(1 <= x <= 9 for x in v)
            tens = lambda v: all(10 <= x <= 90 and x % 10 == 0 for x in v)
            ok = (single(a) and tens(b)) or (single(b) and tens(a))
        results.append(("+1：一个转盘上的数字全是1到9的一位数，另一个转盘上的数字全是10到90的整十数"
                        if ok else "+0：一个转盘上的数字全是1到9的一位数，另一个转盘上的数字全是10到90的整十数"))

        # 08 +3 两个指针各归一盘、中心 x 与转盘中心差≤0.3cm、面积≤转盘 15%、不在旋转组内
        ok = (len(wheels) == 2 and all(p is not None for p in pointers)
              and pointers[0] is not pointers[1])
        if ok:
            for w, p in zip(wheels, pointers):
                if abs(_center(p)[0] - _center(w)[0]) > 0.3:
                    ok = False
                if _area(w) and _area(p) > 0.15 * _area(w):
                    ok = False
                if p["sid"] in [k["sid"] for k in _descendants(shapes, w["sid"])]:
                    ok = False
        results.append(("+3：两个转盘各配一个指针、共2个，指针中心与所属转盘中心的横向偏差不超过0.3厘米，面积不超过转盘的15%，且不随转盘一起旋转"
                        if ok else "+0：两个转盘各配一个指针、共2个，指针中心与所属转盘中心的横向偏差不超过0.3厘米，面积不超过转盘的15%，且不随转盘一起旋转"))

        # 09 +1 底板完整包住转盘、面积为转盘的 1.05—2.5 倍，且盘上数字字号≥12 磅
        ok = (len(wheels) == 2 and all(b is not None for b in boards))
        if ok:
            for num in wheel_nums:
                for n in num:
                    if any(p < 12.0 for p in n["pts"]) or not n["pts"]:
                        ok = False
        results.append(("+1：每个转盘下方都有一块把转盘完整包住的底板，底板面积是转盘的1.05至2.5倍，盘上数字字号不小于12磅"
                        if ok else "+0：每个转盘下方都有一块把转盘完整包住的底板，底板面积是转盘的1.05至2.5倍，盘上数字字号不小于12磅"))

        # 10 +3 两盘宽差、高差均≤0.1cm（对应“大小”面板的宽度/高度）
        ok = (len(wheels) == 2 and abs(wheels[0]["w"] - wheels[1]["w"]) <= 0.1
              and abs(wheels[0]["h"] - wheels[1]["h"]) <= 0.1)
        results.append(("+3：两个转盘的宽度差和高度差都不超过0.1厘米"
                        if ok else "+0：两个转盘的宽度差和高度差都不超过0.1厘米"))

        # 11 +3 两盘中心纵坐标差≤0.1cm（左右齐平）
        ok = (len(wheels) == 2
              and abs(_center(wheels[0])[1] - _center(wheels[1])[1]) <= 0.1)
        results.append(("+3：两个转盘的中心纵坐标差不超过0.1厘米，左右齐平"
                        if ok else "+0：两个转盘的中心纵坐标差不超过0.1厘米，左右齐平"))

        # 12 +3 两盘中心横坐标的中点与 17.03cm 差≤0.3cm（整体关于乘号左右对称）
        ok = False
        if len(wheels) == 2:
            mid = (_center(wheels[0])[0] + _center(wheels[1])[0]) / 2.0
            ok = abs(mid - CENTER_X_CM) <= 0.3
        results.append(("+3：两个转盘中心横坐标的中点与横向17.03厘米的偏差不超过0.3厘米，整体左右对称"
                        if ok else "+0：两个转盘中心横坐标的中点与横向17.03厘米的偏差不超过0.3厘米，整体左右对称"))

        # 13 +1 新增对象全部落在页面内，且左右两个转盘单元（转盘+底板+指针）互不重叠
        added = [s for s in [cross, button] + list(wheels) + list(boards) + list(pointers) if s]
        ok = bool(added)
        for s in added:
            r = _rect(s)
            if r[0] < -0.05 or r[1] < -0.05 or r[2] > PAGE_W_CM + 0.05 or r[3] > PAGE_H_CM + 0.05:
                ok = False
        if len(wheels) == 2:
            units = []
            for i in range(2):
                parts = [wheels[i]] + ([boards[i]] if boards[i] else []) \
                        + ([pointers[i]] if pointers[i] else [])
                rs = [_rect(p) for p in parts]
                units.append((min(r[0] for r in rs), min(r[1] for r in rs),
                              max(r[2] for r in rs), max(r[3] for r in rs)))
            if _inter_area(units[0], units[1]) > 0:
                ok = False
        else:
            ok = False
        results.append(("+1：新增的对象都完整落在页面范围内，左右两个转盘单元互不重叠"
                        if ok else "+0：新增的对象都完整落在页面范围内，左右两个转盘单元互不重叠"))

        # 14 +3 有且只有一个带文字的按钮，并且它是两个转盘旋转动画的触发对象
        trig_all = set(x for seq in timing["seqs"] for x in seq["trigger_spids"] if x)
        rot_sids = set(b["spid"] for b in rots if b["spid"])
        ok = (button is not None and _shape_text(button["el"]).strip() != ""
              and button["sid"] in trig_all and len(wheel_sids) == 2
              and rot_sids == set(wheel_sids))
        results.append(("+3：页面上有一个写着文字的控制按钮，单击它就能让两个转盘转起来"
                        if ok else "+0：页面上有一个写着文字的控制按钮，单击它就能让两个转盘转起来"))

        # 15 +1 按钮中心 x 与 17.03cm 差≤0.5cm，且位于两个转盘中心连线之下
        ok = False
        if button is not None and len(wheels) == 2:
            bcx, bcy = _center(button)
            ok = (abs(bcx - CENTER_X_CM) <= 0.5
                  and bcy > max(_center(w)[1] for w in wheels))
        results.append(("+1：按钮中心与横向17.03厘米的偏差不超过0.5厘米，且位于两个转盘中心的下方"
                        if ok else "+0：按钮中心与横向17.03厘米的偏差不超过0.5厘米，且位于两个转盘中心的下方"))

        # 16 +1 按钮是带填色的矩形/圆角矩形/椭圆类形状，字号≥14 磅，底色与文字明暗差≥0.25
        ok = False
        if button is not None:
            prst = _prst(button["el"])
            shape_ok = (prst in ("roundRect", "round1Rect", "round2SameRect", "round2DiagRect",
                                 "rect", "ellipse", "snip1Rect", "snipRoundRect", "flowChartTerminator")
                        or _has_custgeom(button["el"]))
            fill = _fill_hex(button["el"], palette, clrmap)
            pts = [p for _, p in _font_pts(button["el"])]
            l1 = _lum(fill)
            l2 = _lum(_text_hex(button["el"], palette, clrmap))
            ok = (shape_ok and fill is not None and pts and min(pts) >= 14.0
                  and l1 is not None and l2 is not None and abs(l1 - l2) >= 0.25)
        results.append(("+1：按钮是带填充色的圆角矩形一类形状，文字字号不小于14磅，底色与文字明暗对比明显"
                        if ok else "+0：按钮是带填充色的圆角矩形一类形状，文字字号不小于14磅，底色与文字明暗对比明显"))

        # 17 +5 单击按钮后两盘同时开始转，没有先后延迟
        #    对应动画窗格：两条旋转效果的“开始”分别为“单击时”和“与上一动画同时”，延迟 0 秒
        ok = False
        if len(rots) == 2 and button is not None and rot_sids == set(wheel_sids):
            delays = [(b["eff"]["delay"] in (None, "0", "0ms")) for b in rots]
            types = [b["eff"]["node_type"] for b in rots]
            in_inter = any(seq["node_type"] == "interactiveSeq"
                           and button["sid"] in seq["trigger_spids"] for seq in timing["seqs"])
            ok = all(delays) and "withEffect" in types and in_inter
        results.append(("+5：单击按钮后两个转盘同时开始旋转，两者之间没有先后延迟"
                        if ok else "+0：单击按钮后两个转盘同时开始旋转，两者之间没有先后延迟"))

        # 18 +5 两条旋转都是无限重复，单圈整 360 度（by=±21600000，即 ±360×60000）
        ok = len(rots) == 2
        for b in rots:
            if b["eff"]["repeat"] != "indefinite":
                ok = False
            try:
                if abs(int(b["by"])) != 21600000:
                    ok = False
            except (TypeError, ValueError):
                ok = False
        results.append(("+5：两个转盘都是无限重复地转，单圈刚好转满360度"
                        if ok else "+0：两个转盘都是无限重复地转，单圈刚好转满360度"))

        # 19 +5 再次单击同一按钮两盘同时停下，并可反复启停
        #    对应：p:seq 带 nextCondLst（onClick 指向按钮）+ 另一组 delay="indefinite" 的触发分支
        ok = False
        if button is not None:
            for seq in timing["seqs"]:
                if button["sid"] in seq["next_spids"] and seq["has_indefinite_start"]:
                    ok = True
        results.append(("+5：再次单击同一个按钮，两个转盘同时停下，并且可以反复启动和停止"
                        if ok else "+0：再次单击同一个按钮，两个转盘同时停下，并且可以反复启动和停止"))

        # 20 +3 暂停后停在当前角度：用相对旋转量 by（而非绝对 to），且效果保持结束状态 fill=hold
        ok = len(rots) == 2
        for b in rots:
            if not b["by"] or b["to"]:
                ok = False
            if b["eff"]["fill"] != "hold" or b["inner_fill"] not in ("hold", None):
                ok = False
        results.append(("+3：暂停后转盘停在当前角度而不回到起始位置，动画用相对旋转量并保持结束状态"
                        if ok else "+0：暂停后转盘停在当前角度而不回到起始位置，动画用相对旋转量并保持结束状态"))

        # 21 +3 只由按钮控制：没有“单击任意处/自动播放”的主序列，所有单击触发都指向按钮
        ok = (not timing["has_main_seq"] and button is not None
              and bool(timing["onclick_spids"])
              and all(sp == button["sid"] for sp in timing["onclick_spids"]))
        results.append(("+3：旋转只能由这个按钮控制，既不是单击页面任意处触发，也不是打开页面就自动播放"
                        if ok else "+0：旋转只能由这个按钮控制，既不是单击页面任意处触发，也不是打开页面就自动播放"))

        # 22 +3 只有两个转盘在动；按钮自身的占位动画（≤0.1 秒且缩放 100%）不算多余动画
        ok = len(rots) == 2 and rot_sids == set(wheel_sids)
        for b in timing["behaviors"]:
            if b["spid"] in wheel_sids:
                continue
            try:
                dur = int(b["dur"])
            except (TypeError, ValueError):
                dur = 99999
            placeholder = (b["kind"] == "animScale" and dur <= 100
                           and b["scale_by"] == ("100000", "100000"))
            if not placeholder:
                ok = False
        results.append(("+3：整页只有两个转盘在动，没有给其他对象额外加动画"
                        if ok else "+0：整页只有两个转盘在动，没有给其他对象额外加动画"))

        # 23 +3 绕自身圆心转：旋转对象外框宽高比 0.95—1.05，且外框中心与扇形分布中心差≤0.2cm
        ok = len(rots) == 2 and len(wheels) == 2
        for w in wheels:
            if w["sid"] not in rot_sids:
                ok = False
                continue
            ratio = (w["w"] / w["h"]) if w["h"] else 0.0
            if not (0.95 <= ratio <= 1.05):
                ok = False
            sec, _num = _wheel_profile(shapes, w, palette, clrmap)
            if sec:
                mx = sum(s["center"][0] for s in sec) / len(sec)
                my = sum(s["center"][1] for s in sec) / len(sec)
                wc = _center(w)
                if math.hypot(wc[0] - mx, wc[1] - my) > 0.2:
                    ok = False
            else:
                ok = False
        results.append(("+3：转盘是绕自己的圆心转的，外框接近正方形，旋转中心与扇形圆心偏差不超过0.2厘米"
                        if ok else "+0：转盘是绕自己的圆心转的，外框接近正方形，旋转中心与扇形圆心偏差不超过0.2厘米"))

        # 24 +1 单圈耗时 0.3—1 秒（动画窗格“持续时间”300—1000 毫秒）
        ok = len(rots) == 2
        for b in rots:
            try:
                if not (300 <= int(b["dur"]) <= 1000):
                    ok = False
            except (TypeError, ValueError):
                ok = False
        results.append(("+1：转盘转一圈用0.3至1秒，快慢适合课堂抢答"
                        if ok else "+0：转盘转一圈用0.3至1秒，快慢适合课堂抢答"))

        # 25 +1 每盘由一个组合承担旋转、嵌套不超过 4 层，且不是图片或视频在转
        ok = len(wheels) == 2
        for w in wheels:
            if w["tag"] != "grpSp" or w["sid"] not in rot_sids:
                ok = False
                continue
            if _max_depth(shapes, w["sid"]) > 4:
                ok = False
            for k in _descendants(shapes, w["sid"]):
                if k["tag"] == "pic" or k["el"].find(".//p:nvPr/a:videoFile", NS) is not None:
                    ok = False
        results.append(("+1：每个转盘由一个组合整体承担旋转，组合嵌套不超过4层，转的不是图片或视频"
                        if ok else "+0：每个转盘由一个组合整体承担旋转，组合嵌套不超过4层，转的不是图片或视频"))

        # 26 −3 背景图必须还在原位：铺满整页、位于 (0,0)、且是 spTree 里的第一个对象（最底层）
        bad = (ident["bg"] is None) or (not ident["bg_first"])
        results.append(("−3：第10页铺满整页的背景图被删除、移动、改了尺寸，或不再位于最底层"
                        if bad else "−0：第10页铺满整页的背景图被删除、移动、改了尺寸，或不再位于最底层"))

        # 27 −3 不得再压一层覆盖页面 60% 以上的图片或视频（整页贴图、整页录屏视频）
        bad = False
        for s in top10:
            if s is ident["bg"]:
                continue
            is_media = (s["tag"] == "pic") or (s["el"].find(".//a:videoFile", NS) is not None) \
                or (s["el"].find(".//p:nvPr/p:audioFile", NS) is not None)
            if is_media and _area(s) >= 0.60 * page_area:
                bad = True
        results.append(("−3：第10页多出了一个覆盖页面六成以上的图片或视频对象"
                        if bad else "−0：第10页多出了一个覆盖页面六成以上的图片或视频对象"))

        # 28 −1 4 个小图标位置与大小基线（cm）：见 ICONS_CM，位置容差 0.30、尺寸容差 0.15
        bad = len(ident["icons"]) != 4
        if not bad:
            used = []
            for exp in ICONS_CM:
                match = None
                for s in ident["icons"]:
                    if s in used:
                        continue
                    if (abs(s["x"] - exp[0]) <= 0.30 and abs(s["y"] - exp[1]) <= 0.30
                            and abs(s["w"] - exp[2]) <= 0.15 and abs(s["h"] - exp[2]) <= 0.15):
                        match = s
                        break
                if match is None:
                    bad = True
                    break
                used.append(match)
        results.append(("−1：第10页原有的4个小图标中有被删除、被移动或被改了大小的"
                        if bad else "−0：第10页原有的4个小图标中有被删除、被移动或被改了大小的"))

        # 29 −3 原粉色椭圆区域（11.54,3.73)-(22.52,15.68) 不得被单个新增对象盖住过半
        er = ELLIPSE_RECT_CM
        er_area = (er[2] - er[0]) * (er[3] - er[1])
        bad = False
        for s in top10:
            if s is ident["bg"] or _area(s) >= 0.95 * page_area:
                continue
            if er_area and _inter_area(er, _rect(s)) / er_area > 0.50:
                bad = True
        results.append(("−3：原来放粉色椭圆的那片区域被某个新增对象盖住了一半以上"
                        if bad else "−0：原来放粉色椭圆的那片区域被某个新增对象盖住了一半以上"))

        # 30 −1 页面上残留提示字样，或出现面积超过页面 15% 且没有文字的纯色块
        page_text = "".join(_shape_text(s["el"]) for s in shapes if s["tag"] == "sp")
        low = page_text.lower()
        bad = any(w.lower() in low for w in BANNED_WORDS)
        for s in top10:
            if s["tag"] != "sp" or _shape_text(s["el"]).strip():
                continue
            if _fill_hex(s["el"], palette, clrmap) and _area(s) > 0.15 * page_area:
                bad = True
        results.append(("−1：第10页残留了“同步/双转盘/示例/demo/水印/模板/参考图”之类的说明文字，或多出面积超过页面15%的无文字纯色块"
                        if bad else "−0：第10页残留了“同步/双转盘/示例/demo/水印/模板/参考图”之类的说明文字，或多出面积超过页面15%的无文字纯色块"))

        # 31 −5 页数必须是 11 页，页面尺寸必须是 33.87×19.05cm（容差 0.1cm）
        bad = (len(slide_parts) != SLIDE_COUNT
               or abs(page_w - PAGE_W_CM) > 0.1 or abs(page_h - PAGE_H_CM) > 0.1)
        results.append(("−5：演示文稿不是11页，或页面尺寸不是宽33.87厘米、高19.05厘米"
                        if bad else "−0：演示文稿不是11页，或页面尺寸不是宽33.87厘米、高19.05厘米"))

        # 32 −5 第 1—9、11 页文字指纹必须与内联 SHA-1 完全一致（改字、删框、加框都会变）
        bad = False
        for page, want in TEXT_SHA1.items():
            i = page - 1
            if i >= len(slide_roots) or slide_roots[i] is None:
                bad = True
                continue
            got, _n = _text_sha1(slide_roots[i].find("p:cSld/p:spTree", NS))
            if got != want:
                bad = True
        results.append(("−5：除第10页以外的其余10页文字被改动了，与原有文字指纹不符"
                        if bad else "−0：除第10页以外的其余10页文字被改动了，与原有文字指纹不符"))

        # 33 −3 任意页出现自动缩小字号或压缩行距（a:normAutofit 带 fontScale/lnSpcReduction）
        bad = False
        for root in slide_roots:
            if root is None:
                continue
            for af in root.iter("{%s}normAutofit" % _A):
                if af.get("fontScale") or af.get("lnSpcReduction"):
                    bad = True
        results.append(("−3：有页面的文本框出现了自动缩小字号或压缩行距，说明文字装不下了"
                        if bad else "−0：有页面的文本框出现了自动缩小字号或压缩行距，说明文字装不下了"))

        # 34 −1 不得有隐藏页（p:sld@show="0"）或页面切换效果（p:transition）
        bad = False
        for root in slide_roots:
            if root is None:
                continue
            if root.get("show") == "0":
                bad = True
            if root.find(".//p:transition", NS) is not None:
                bad = True
        results.append(("−1：演示文稿里出现了隐藏页或页面切换效果"
                        if bad else "−0：演示文稿里出现了隐藏页或页面切换效果"))

        # 35 −1 母版应为 2 个、版式 3 个，主题中英文字体均为 Microsoft YaHei
        masters = len([n for n in names if re.match(r"ppt/slideMasters/slideMaster\d+\.xml$", n)])
        layouts = len([n for n in names if re.match(r"ppt/slideLayouts/slideLayout\d+\.xml$", n)])
        theme = _xml(z, "ppt/theme/theme1.xml")
        fonts = []
        if theme is not None:
            for tag in ("majorFont", "minorFont"):
                node = theme.find(".//a:fontScheme/a:%s/a:latin" % tag, NS)
                fonts.append((node.get("typeface") or "") if node is not None else "")
        bad = (masters != MASTER_COUNT or layouts != LAYOUT_COUNT
               or fonts != [THEME_FONT_EXPECT, THEME_FONT_EXPECT])
        results.append(("−1：母版数量不是2个、版式数量不是3个，或主题字体不再是 Microsoft YaHei"
                        if bad else "−0：母版数量不是2个、版式数量不是3个，或主题字体不再是 Microsoft YaHei"))

        # 36 −1 每页备注都应保留一段以“课堂备注：”开头、长度 40—90 字的段落
        #    注意：备注页里还有日期与页码占位符，必须逐段匹配，不能拿整页备注文本判断
        bad = False
        for part in slide_parts:
            rels = _rels(z, part)
            note_part = None
            for _rid, (target, _mode) in rels.items():
                if "notesSlide" in target:
                    note_part = "ppt/notesSlides/" + os.path.basename(target)
            hit_note = False
            note_root = _xml(z, note_part) if note_part else None
            if note_root is not None:
                for para in note_root.iter("{%s}p" % _A):
                    txt = "".join(t.text or "" for t in para.iter("{%s}t" % _A)).strip()
                    if txt.startswith(NOTE_PREFIX) and NOTE_LEN_RANGE[0] <= len(txt) <= NOTE_LEN_RANGE[1]:
                        hit_note = True
            if not hit_note:
                bad = True
        results.append(("−1：有页面的备注丢了或被改写，缺少以“课堂备注：”开头、长度40到90字的段落"
                        if bad else "−0：有页面的备注丢了或被改写，缺少以“课堂备注：”开头、长度40到90字的段落"))

        # 37 −3 不得夹带宏代码（vbaProject.bin）
        bad = any(n.lower().endswith("vbaproject.bin") for n in names)
        results.append(("−3：文件里夹带了宏代码"
                        if bad else "−0：文件里夹带了宏代码"))

        # 38 −3 不得用外部链接的媒体（换台机器就播不出来）
        bad = False
        for part in [n for n in names if n.endswith(".rels")]:
            root = _xml(z, part)
            if root is None:
                continue
            for rel in root:
                if (rel.get("TargetMode") or "") == "External":
                    tgt = (rel.get("Target") or "").lower()
                    if re.search(r"\.(mp4|avi|wmv|mov|mkv|flv|mp3|wav|wma|m4a|gif)($|\?)", tgt) \
                            or "video" in (rel.get("Type") or "").lower() \
                            or "audio" in (rel.get("Type") or "").lower() \
                            or "image" in (rel.get("Type") or "").lower():
                        bad = True
        results.append(("−3：页面上的音视频或图片用的是外部链接，不是嵌入到文件里"
                        if bad else "−0：页面上的音视频或图片用的是外部链接，不是嵌入到文件里"))

        # 39 −1 体积不得超过 18MB，且不得留下没被任何页面引用的冗余素材
        size = os.path.getsize(file_path)
        referenced = set()
        for part in [n for n in names if n.endswith(".rels")]:
            root = _xml(z, part)
            if root is None:
                continue
            for rel in root:
                tgt = (rel.get("Target") or "").replace("../", "")
                if tgt:
                    referenced.add(os.path.basename(tgt))
        media = [n for n in names if n.startswith("ppt/media/") and not n.endswith("/")]
        unused = [n for n in media if os.path.basename(n) not in referenced]
        bad = (size > SIZE_LIMIT_BYTES) or bool(unused)
        results.append(("−1：文件体积超过18MB，或包里残留了没被任何页面引用的冗余素材"
                        if bad else "−0：文件体积超过18MB，或包里残留了没被任何页面引用的冗余素材"))

        # 40 −1 文件名必须是“口算教学课件.pptx”
        bad = os.path.basename(file_path) != FILE_NAME_EXPECT
        results.append(("−1：文件名不是“口算教学课件.pptx”"
                        if bad else "−0：文件名不是“口算教学课件.pptx”"))

        # 41 −1 第10页不得留下空形状（无文字且无填充无边框）、页面外对象或完全重复的对象
        bad = False
        for s in shapes:
            if s["tag"] == "sp" and not _shape_text(s["el"]).strip():
                spPr = s["el"].find("p:spPr", NS)
                no_fill = spPr is not None and spPr.find("a:noFill", NS) is not None
                ln = spPr.find("a:ln", NS) if spPr is not None else None
                no_line = (ln is None) or (ln.find("a:noFill", NS) is not None)
                if no_fill and no_line:
                    bad = True
            r = _rect(s)
            if r[2] <= 0 or r[3] <= 0 or r[0] >= PAGE_W_CM or r[1] >= PAGE_H_CM:
                bad = True
        sig = [(s["name"], round(s["x"], 2), round(s["y"], 2), round(s["w"], 2), round(s["h"], 2))
               for s in top10]
        if len(sig) != len(set(sig)):
            bad = True
        results.append(("−1：第10页留下了空形状、跑到页面外的对象，或完全重复的对象"
                        if bad else "−0：第10页留下了空形状、跑到页面外的对象，或完全重复的对象"))

        # 42 −3 第 1—9、11 页不得被加动画（这些页原本没有任何动画）
        bad = False
        for i, root in enumerate(slide_roots):
            if root is None or i == idx10:
                continue
            t = _parse_timing(root)
            if t["behaviors"]:
                bad = True
        results.append(("−3：第1到9页或第11页被加上了动画"
                        if bad else "−0：第1到9页或第11页被加上了动画"))

        # 43 −1 第10页不得放音频或视频对象（旋转应由动画实现）
        bad = False
        if root10 is not None:
            for tag in ("videoFile", "audioFile", "wavAudioFile", "audioCd"):
                if root10.find(".//a:%s" % tag, NS) is not None:
                    bad = True
            if root10.find(".//p:nvPr/p:media", NS) is not None:
                bad = True
        results.append(("−1：第10页放了音频或视频对象"
                        if bad else "−0：第10页放了音频或视频对象"))

        # 44 −1 第10页不得用单击跳转超链接（会跳走而不是原地转）
        bad = False
        if root10 is not None:
            for hl in root10.iter("{%s}hlinkClick" % _A):
                if hl.get("{%s}id" % NS["r"]) or hl.get("action"):
                    bad = True
        results.append(("−1：第10页的对象被加了单击跳转的超链接"
                        if bad else "−0：第10页的对象被加了单击跳转的超链接"))

        # 45 −1 文档属性（标题、主题、作者）不得被清空或改错
        core = _xml(z, "docProps/core.xml")
        title = subject = creator = ""
        if core is not None:
            for tag, ns_key in (("title", "dc"), ("subject", "dc"), ("creator", "dc")):
                node = core.find("%s:%s" % (ns_key, tag), NS)
                val = (node.text or "").strip() if node is not None else ""
                if tag == "title":
                    title = val
                elif tag == "subject":
                    subject = val
                else:
                    creator = val
        bad = (title != DOC_TITLE_EXPECT or subject != DOC_SUBJECT_EXPECT or not creator)
        results.append(("−1：文档属性里的标题、主题或作者被清空或改错了"
                        if bad else "−0：文档属性里的标题、主题或作者被清空或改错了"))









    finally:
        z.close()

    if len(results) != len(_MAX_DELTAS_ORDER):
        raise RuntimeError("dim2 项数不一致：results=%d vs order=%d"
                           % (len(results), len(_MAX_DELTAS_ORDER)))
    dim2_items = []
    total_score = 0
    for raw, max_delta in zip(results, _MAX_DELTAS_ORDER):
        head, _, rule = raw.partition("：")
        token = head.replace("+", "").replace("−", "-").replace("－", "-").strip()
        try:
            delta = int(token)
        except ValueError:
            delta = 0
        hit = delta != 0
        dim2_items.append({"rule": rule, "max_delta": max_delta,
                           "delta": max_delta if hit else 0,
                           "hit": hit})
        total_score += (max_delta if hit else 0)
    result["dim2_items"] = dim2_items
    result["total_score"] = max(0, total_score)
    result["max_score"] = MAX_SCORE_TOTAL
    return result


if __name__ == "__main__":
    if len(sys.argv) >= 2:
        _dir = sys.argv[1]
    else:
        _dir = os.path.dirname(os.path.abspath(__file__))
    _out = evaluate(_dir)
    print(json.dumps(_out, ensure_ascii=False, indent=2))












