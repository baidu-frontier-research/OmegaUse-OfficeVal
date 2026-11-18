#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# pyright: reportUnknownMemberType=false, reportUnknownVariableType=false, reportUnknownArgumentType=false, reportUnknownParameterType=false, reportMissingParameterType=false, reportPrivateUsage=false, reportUnknownLambdaType=false, reportGeneralTypeIssues=false

import os
import re
import sys
import json
import math
import statistics
import unicodedata

# PyMuPDF 已从本项目移除（AGPL 与本项目许可证不兼容），运行环境不安装 fitz。
# 改用同目录的 pymupdf_compat：用 pdfplumber + pypdfium2 实现本脚本用到的
# PyMuPDF 子集。verifier 由 core/worker.py 按文件路径加载（其所在目录已被插入
# sys.path），作为 verifiers 包导入时下面这两行会补齐目录，两种方式都成立。
_VERIFIER_DIR = os.path.dirname(os.path.abspath(__file__))
if _VERIFIER_DIR not in sys.path:
    sys.path.insert(0, _VERIFIER_DIR)
import pymupdf_compat as pymupdf

# 说明：按接口统一约定，本模块不修改全局 sys.stdout，也不 print 主结果。
# 主结果统一由 evaluate(dir_path) -> dict 返回；入参为“脚本所在目录的路径”，
# 由脚本自己在该目录内定位并打开被评估的 .pdf。本文件仅在 __main__ 下作为
# 本地调试入口，才把 dict 序列化为 JSON 打印。
#
# 任务背景：源文件为 16 页《七年级地理导学案》PDF，指令要求“pdf文档需要添加
# 目录，目录样式参考图片”，即在封面页之后插入 1 个独立整页目录，全册变 17 页。
# 本脚本严格对齐细则，仅踩细则写出的 35 项（未写出的不额外约束）。

SCRIPT_ID = "017"
# 维度二所有“加分项”max_delta 之和；用于维度一未通过 / 脚本异常时兜底。
# 细则含 19 个加分项（合计 +65）与 16 个扣分项（合计 -23）；
# max_score 只统计加分项上限，扣分项以负 max_delta 形式出现在 dim2_items 中。
MAX_SCORE_TOTAL = 65

# 维度二所有评分项按“结果追加顺序”的 max_delta 列表；末端把 results 里的
# "+N：..." / "-N：..." 字符串反解为结构化 dim2_items 时用。顺序须与
# _evaluate_impl 中 _emit(...) 的顺序严格一致；共 35 项。
_MAX_DELTAS_ORDER = [
    5,   # 01 封面后新增独立目录页、16→17页、其余各页内容与次序不变
    1,   # 02 “目录”标题：左上、绿色、加粗、≥30磅
    3,   # 03 标题左侧紧贴页边的绿色斜条纹装饰（深色块+2-3条渐浅斜条）
    3,   # 04 标题下方横向绿色分隔线/色带，横贯至版心右边界、线宽≥2pt
    5,   # 05 目录共 18 条条目、无遗漏/重复/杜撰
    3,   # 06 条目名称与其所指页面标题文字一致
    5,   # 07 四个单元序号位于绿色五边形旗标内、白色加粗
    3,   # 08 单元名称在旗标右侧、黑色加粗
    3,   # 09 任务1-10 相对旗标缩进≥30pt、常规体、15-17磅
    3,   # 10 三个项目条目前有绿色三角/箭头符号、加粗、≥17磅且大于任务
    1,   # 11 “学习结构说明”首条、加粗、与单元同级左对齐不缩进
    5,   # 12 18 条名称与页码之间点状前导引导线，不重叠/不断裂/不越过页码
    3,   # 13 18 条页码右对齐并与版心右边界齐平
    5,   # 14 目录所标页码与条目实际所在页完全对应
    5,   # 15 全册页脚“第N页”与物理页序一致（目录页第2页、末页第17页）
    3,   # 16 目录页页眉：灰色#646464 8-9磅、右551pt/上36pt，与其余页一致
    3,   # 17 目录页页脚：居中、灰色 8-10.5磅、纵向约790-806pt（页码内容归 15）
    5,   # 18 18 条全部排在同一页且落在版心内（上≥36pt、下≤790pt、不侵页脚）
    1,   # 19 各绿色元素同一色系，全页未混入第三种强调色
    -3,  # 20 文件体积≤5MB（无未子集化的巨型字体流导致膨胀至20MB以上）
    -5,  # 21 目录与正文中文字体同属一款简体中文黑体字族、未混入日文字形字体
    -1,  # 22 目录条目“任务”与序号之间无多余空格
    -1,  # 23 文档属性 标题/作者/主题
    -2,  # 24 第1页主标题/副标题/学科行/导语四行的色号、字号、字重与居中【合并】
    -1,  # 25 第1页九宫格 3行3列 #404040 加粗11磅 + 9 种浅色底纹
    -1,  # 26 全册各级标题条（说明页/单元/任务/项目）通栏底纹(左右容差8pt)+深蓝加粗15磅【合并】
    -1,  # 27 单行四格表（第1页底部 + 十条任务标题条下方）标签深蓝加粗8.5磅【合并】
    -1,  # 28 第3页三个小节标题 #1F4E79 加粗12磅并按序排列
    -1,  # 29 第3页三张表格行列结构与表头样式（四步学习链/项目地图/评价方式）【合并】
    -1,  # 30 小节标题 深蓝或绿色 加粗 11磅
    -1,  # 31 六处示意图 居中、宽约354pt、高度符合、无拉伸变形【合并】
    -1,  # 32 第1页及第3-17页页眉 #646464 8.5磅、右551pt、上36pt
    -1,  # 33 第1页及第3-17页版心 左≈44pt、右≈551.5pt、正文首行上≈36pt（目录页归 18）
    -1,  # 34 各页表格细实线全框、单元格文字左对齐且无截断压盖
    -1,  # 35 各处“□”空心方框、字号合规、与所在行文字基线对齐
]
assert len(_MAX_DELTAS_ORDER) == 35
assert sum(d for d in _MAX_DELTAS_ORDER if d > 0) == MAX_SCORE_TOTAL
assert sum(d for d in _MAX_DELTAS_ORDER if d < 0) == -23

# ========== 期望值表（来源：原始 16 页输入件 + 细则明文） ==========
A4_W, A4_H = 595.0, 842.0
EXPECT_PAGES = 17
TOC_PIDX = 1                      # 目录页 0-based 页索引（= 第2页）
BODY_LEFT, BODY_RIGHT = 44.0, 551.5   # 版心左右边界（pt）
BODY_TOP, BODY_BOTTOM = 36.0, 790.0   # 版心上边界 / 页脚区上沿（pt）
PAGE_CX = A4_W / 2.0
# 标题条通栏底纹左右边界与版心边界的容差（pt）：口径从宽，只要看起来仍是
# 一条与版心齐平的通栏色带即可，允许由多段色块拼接而成。
BAR_EDGE_TOL = 8.0

C_DARKBLUE = (0x1F, 0x4E, 0x79)   # 深蓝 #1F4E79
C_GREEN = (0x54, 0x82, 0x35)      # 绿 #548235
C_DARKGRAY = (0x40, 0x40, 0x40)   # 深灰 #404040
C_HEADGRAY = (0x64, 0x64, 0x64)   # 页眉灰 #646464

HEADER_TEXT = "七年级地理导学案·项目化学习任务单"

# 目录 18 条条目：(种类, 前缀标签, 名称, 应标注页码)
# 种类 unit=单元旗标行 / task=任务行 / project=项目行 / intro=说明行
TOC_ENTRIES = [
    ("intro",   "学习结构说明", "学习结构说明",           3),
    ("unit",    "第一单元",     "地球与地图",             4),
    ("task",    "任务1",        "认识地球与经纬网",       4),
    ("task",    "任务2",        "地图三要素与方向",       5),
    ("task",    "任务3",        "等高线地形图",           6),
    ("unit",    "第二单元",     "陆地与海洋",             8),
    ("task",    "任务4",        "七大洲与四大洋",         8),
    ("task",    "任务5",        "海陆变迁与板块运动",     8),
    ("unit",    "第三单元",     "天气与气候",            10),
    ("task",    "任务6",        "天气预报与空气质量",    10),
    ("task",    "任务7",        "气温曲线与降水柱状图",  11),
    ("task",    "任务8",        "世界气候类型与生活",    12),
    ("unit",    "第四单元",     "居民与聚落",            13),
    ("task",    "任务9",        "人口分布与聚落形成",    13),
    ("task",    "任务10",       "聚落发展与可持续生活",  14),
    ("project", "期中项目",     "校园导览图制作任务",    15),
    ("project", "期末项目",     "家乡聚落观察报告",      16),
    ("project", "全册学习档案袋", "阶段复盘与成长记录",  17),
]
assert len(TOC_ENTRIES) == 18

# 各正文页应出现的标题文字（1-based 输出页号 -> 归一化后的必含片段）
BODY_TITLES = {
    3:  ["学习结构说明"],
    4:  ["单元一", "地球与地图", "任务1", "认识地球与经纬网"],
    5:  ["任务2", "地图三要素与方向"],
    6:  ["任务3", "等高线地形图"],
    8:  ["单元二", "陆地与海洋", "任务4", "七大洲与四大洋", "任务5", "海陆变迁与板块运动"],
    10: ["单元三", "天气与气候", "任务6", "天气预报与空气质量"],
    11: ["任务7", "气温曲线与降水柱状图"],
    12: ["任务8", "世界气候类型与生活"],
    13: ["单元四", "居民与聚落", "任务9", "人口分布与聚落形成"],
    14: ["任务10", "聚落发展与可持续生活"],
    15: ["期中项目", "校园导览图制作任务"],
    16: ["期末项目", "家乡聚落观察报告"],
    17: ["全册学习档案袋", "阶段复盘与成长记录"],
}

UNIT_BAR_PAGES = [4, 8, 10, 13]          # 单元标题条所在输出页
TASK_BARS = [(4, "任务1"), (5, "任务2"), (6, "任务3"), (8, "任务4"), (8, "任务5"),
             (10, "任务6"), (11, "任务7"), (12, "任务8"), (13, "任务9"), (14, "任务10")]
PROJ_BAR_PAGES = [15, 16, 17]            # 项目/档案袋标题条所在输出页

# 示意图：输出页号 -> (宽, 高) pt
DIAGRAMS = {4: (354.2, 331.6), 5: (354.2, 180.7), 7: (354.2, 238.5),
            10: (354.2, 173.5), 11: (354.2, 224.0), 13: (354.2, 180.7)}

# “□”在各输出页的数量（源自输入件原有分布，合计 108）
CHECKBOX_PER_PAGE = {3: 8, 4: 7, 5: 6, 6: 6, 7: 3, 8: 10, 9: 6, 10: 7,
                     11: 6, 12: 9, 13: 7, 14: 9, 15: 12, 16: 12}
CHECKBOX_TOTAL = 108
assert sum(CHECKBOX_PER_PAGE.values()) == CHECKBOX_TOTAL

NINE_CELLS = ["观天识云", "读图定位", "认识地球", "辨析地形", "探究气候",
              "观察聚落", "解释分布", "表达观点", "完成项目"]
FOUR_LABELS = ["学校：", "班级：", "姓名：", "使用时间："]
P3_SECTIONS = ["一、学习方式", "二、全册项目主题", "三、评价方式"]
SEC_TITLES = {"真实情境", "学习目标", "图表示意图", "图表/示意图", "活动任务单",
              "分层练习", "核心概念整理", "单元学习任务板"}
# ========== 通用小工具 ==========
_WS_RE = re.compile(r"[\s 　   ]+")
_DOT_RE = re.compile(r"^[.．·・…‧∙•·․⋯\s]+$")
_CJK_RE = re.compile(r"[㐀-䶿一-鿿豈-﫿]")


def _norm(text):
    """去掉全部空白（含 NBSP / 全角空格）后的紧凑串，用于文字比对。

    先做 NFC 规范化：部分文档嵌入日文 CJK 字体，取字时会落到“兼容汉字”
    区（如 落=U+F918、量=U+F97E），NFC 可折回标准汉字，避免误判为缺字。
    """
    s = unicodedata.normalize("NFC", text or "")
    return _WS_RE.sub("", s)


def _rgb(color_int):
    """PyMuPDF span color（int）-> (r, g, b) 0-255。"""
    c = int(color_int or 0)
    return ((c >> 16) & 0xFF, (c >> 8) & 0xFF, c & 0xFF)


def _f2rgb(seq):
    """drawings 的 0-1 浮点色 -> (r, g, b) 0-255。"""
    if not seq:
        return None
    v = list(seq)
    if len(v) == 1:                      # 灰度
        v = [v[0], v[0], v[0]]
    if len(v) == 4:                      # CMYK
        c, m, y, k = v
        return (int(round(255 * (1 - min(1.0, c + k)))),
                int(round(255 * (1 - min(1.0, m + k)))),
                int(round(255 * (1 - min(1.0, y + k)))))
    return tuple(max(0, min(255, int(round(x * 255)))) for x in v[:3])


def _cdist(a, b):
    """两色的最大分量差，用于“接近某个指定色”判断。"""
    if not a or not b:
        return 999
    return max(abs(int(a[i]) - int(b[i])) for i in range(3))


def _is_bold(span):
    f = span.get("font") or ""
    if (span.get("flags") or 0) & 16:
        return True
    low = f.lower()
    return ("bold" in low or "heavy" in low or "black" in low
            or "黑" in f or low.endswith("-bd") or ",bold" in low)


def _is_dark(rgb, thr=60):
    return rgb is not None and max(rgb) <= thr


def _is_white(rgb, thr=232):
    return rgb is not None and min(rgb) >= thr


def _is_greenish(rgb):
    """绿色系：G 为最大分量且明显高于 R/B（含浅绿网点色）。"""
    if rgb is None:
        return False
    r, g, b = rgb
    return g >= 100 and g > r and g > b


def _is_neutral(rgb, tol=22):
    return rgb is not None and (max(rgb) - min(rgb)) <= tol


def _is_lightblue(rgb):
    if rgb is None:
        return False
    r, g, b = rgb
    return min(rgb) >= 185 and b >= g >= r and (b - r) >= 8


def _is_lightgreen(rgb):
    if rgb is None:
        return False
    r, g, b = rgb
    return min(rgb) >= 185 and g > r and g > b


def _is_tint(rgb):
    """任意“浅色底纹”：足够亮、又不是纯白（浅蓝/浅绿/浅黄/浅灰皆可）。"""
    if rgb is None:
        return False
    if max(rgb) < 235 or min(rgb) < 180:
        return False
    return (max(rgb) - min(rgb)) >= 8 or max(rgb) <= 250


def _union_span(rects, gap=3.0):
    """把一组矩形按 x 方向并集合并，返回最宽的连续区间 (x0, x1)。

    通栏底纹在部分生成器里会被拆成几十个按 run 断开的小色块，
    只看“最宽的单块”会误判，必须先做并集再量总宽。
    """
    iv = sorted((float(r[0]), float(r[1])) for r in rects if r[1] > r[0])
    if not iv:
        return None
    best = cur = list(iv[0])
    for x0, x1 in iv[1:]:
        if x0 <= cur[1] + gap:
            cur[1] = max(cur[1], x1)
        else:
            if cur[1] - cur[0] > best[1] - best[0]:
                best = cur
            cur = [x0, x1]
    if cur[1] - cur[0] > best[1] - best[0]:
        best = cur
    return (best[0], best[1])


def _centered(x0, x1, tol=10.0, cx=PAGE_CX):
    return abs((x0 + x1) / 2.0 - cx) <= tol


def _near(a, b, tol):
    return a is not None and b is not None and abs(a - b) <= tol


def _spread(vals):
    vals = [v for v in vals if v is not None]
    if len(vals) < 2:
        return 0.0
    return max(vals) - min(vals)
# ========== 页面元素提取 ==========
def _spans(page):
    """页面全部非空文本 span，附归一化文本、坐标、字号、颜色、粗体标记。"""
    out = []
    try:
        raw = page.get_text("dict")
    except Exception:
        return out
    for blk in raw.get("blocks", []):
        if blk.get("type") == 1:
            continue
        for line in blk.get("lines", []):
            for sp in line.get("spans", []):
                txt = sp.get("text") or ""
                if not txt.strip():
                    continue
                bb = sp.get("bbox") or (0, 0, 0, 0)
                out.append({
                    "raw": txt, "norm": _norm(txt),
                    "x0": float(bb[0]), "y0": float(bb[1]),
                    "x1": float(bb[2]), "y1": float(bb[3]),
                    "cy": (float(bb[1]) + float(bb[3])) / 2.0,
                    "size": float(sp.get("size") or 0.0),
                    "font": sp.get("font") or "",
                    "flags": int(sp.get("flags") or 0),
                    "rgb": _rgb(sp.get("color")),
                    "bold": False,
                })
    for s in out:
        s["bold"] = _is_bold(s)
    out.sort(key=lambda s: (round(s["cy"], 1), s["x0"]))
    return out


def _image_blocks(page):
    """页面图片块：版面 bbox + 原始像素宽高（用于判断是否拉伸变形）。"""
    out = []
    try:
        raw = page.get_text("dict")
    except Exception:
        return out
    for blk in raw.get("blocks", []):
        if blk.get("type") != 1:
            continue
        bb = blk.get("bbox") or (0, 0, 0, 0)
        out.append({"x0": float(bb[0]), "y0": float(bb[1]),
                    "x1": float(bb[2]), "y1": float(bb[3]),
                    "w": float(bb[2]) - float(bb[0]),
                    "h": float(bb[3]) - float(bb[1]),
                    "pw": float(blk.get("width") or 0.0),
                    "ph": float(blk.get("height") or 0.0)})
    return out


def _drawings(page):
    """矢量图元：区分填充(fill)与描边(stroke)，并统计顶点数以识别多边形。"""
    fills, strokes = [], []
    try:
        raw = page.get_drawings()
    except Exception:
        return fills, strokes
    seen = set()
    for dr in raw:
        rect = dr.get("rect")
        if rect is None:
            continue
        segs = [it[0] for it in (dr.get("items") or [])]
        npts = sum(4 if k == "re" else (2 if k == "l" else 3) for k in segs)
        base = {"x0": float(rect.x0), "y0": float(rect.y0),
                "x1": float(rect.x1), "y1": float(rect.y1),
                "w": float(rect.width), "h": float(rect.height),
                "segs": segs, "nseg": len(segs), "npts": npts,
                "width": float(dr.get("width") or 0.0),
                "dashes": dr.get("dashes") or ""}
        typ = dr.get("type") or ""
        if dr.get("fill") is not None and typ in ("f", "fs"):
            item = dict(base)
            item["rgb"] = _f2rgb(dr.get("fill"))
            key = ("f", round(item["x0"], 1), round(item["y0"], 1),
                   round(item["x1"], 1), round(item["y1"], 1), item["rgb"])
            if key not in seen:
                seen.add(key)
                fills.append(item)
        if dr.get("color") is not None and typ in ("s", "fs"):
            item = dict(base)
            item["rgb"] = _f2rgb(dr.get("color"))
            key = ("s", round(item["x0"], 1), round(item["y0"], 1),
                   round(item["x1"], 1), round(item["y1"], 1), item["rgb"],
                   round(item["width"], 2))
            if key not in seen:
                seen.add(key)
                strokes.append(item)
    return fills, strokes
def _render(page, scale=2.0):
    """渲染页面位图；目录页装饰可能是位图，需要按区域取色。失败返回 None。"""
    try:
        return page.get_pixmap(matrix=pymupdf.Matrix(scale, scale), alpha=False)
    except Exception:
        return None


def _pix_colors(pix, scale, x0, y0, x1, y1, samples=26):
    """在给定 pt 矩形内采样，返回 [(色, 计数), ...]（仅保留有彩度的像素）。"""
    if pix is None:
        return []
    px0, py0 = max(0, int(x0 * scale)), max(0, int(y0 * scale))
    px1 = min(pix.width, int(math.ceil(x1 * scale)))
    py1 = min(pix.height, int(math.ceil(y1 * scale)))
    if px1 <= px0 or py1 <= py0:
        return []
    sx = max(1, (px1 - px0) // samples)
    sy = max(1, (py1 - py0) // samples)
    buckets = {}
    for yy in range(py0, py1, sy):
        for xx in range(px0, px1, sx):
            try:
                c = pix.pixel(xx, yy)[:3]
            except Exception:
                continue
            if max(c) - min(c) < 25 or max(c) < 45:
                continue
            key = tuple(v // 16 for v in c)
            slot = buckets.setdefault(key, [0, [0, 0, 0]])
            slot[0] += 1
            for i in range(3):
                slot[1][i] += c[i]
    out = []
    for _, (n, tot) in buckets.items():
        out.append((tuple(int(round(v / n)) for v in tot), n))
    out.sort(key=lambda kv: -kv[1])
    return out


def _pix_main(pix, scale, x0, y0, x1, y1):
    """区域内最主要的彩色，用于给位图装饰/色带定色。"""
    cols = _pix_colors(pix, scale, x0, y0, x1, y1)
    return cols[0][0] if cols else None


def _overlap(a0, a1, b0, b1):
    return max(0.0, min(a1, b1) - max(a0, b0))


def _in_band(el, y0, y1, ratio=0.4):
    """图元与 [y0,y1] 纵向重叠达到自身高度的 ratio 即视为落在该带内。"""
    h = max(0.5, el["y1"] - el["y0"])
    return _overlap(el["y0"], el["y1"], y0, y1) >= min(h, (y1 - y0)) * ratio


def _footer_spans(spans):
    """页脚区文本（y0 > 782 且横向居于版心中部），排除正文溢出。"""
    return [s for s in spans if s["y0"] > 782.0]


def _header_spans(spans):
    """页眉区文本（y1 < 50）。"""
    return [s for s in spans if s["y1"] < 50.0]


def _body_spans(spans):
    return [s for s in spans if s["y1"] >= 50.0 and s["y0"] <= 782.0]


def _footer_pageno(spans):
    """从页脚区还原“第N页”，允许被拆成多个 span 或夹杂空格。返回 int 或 None。"""
    cand = [s for s in _footer_spans(spans) if s["x1"] >= 200.0 and s["x0"] <= 400.0]
    if not cand:
        return None
    cand.sort(key=lambda s: s["x0"])
    txt = _norm("".join(s["raw"] for s in cand))
    m = re.search(r"第(\d{1,3})页", txt)
    return int(m.group(1)) if m else None


def _toc_rows(spans):
    """把目录页文本聚成条目行；返回按纵向次序排列的行列表。"""
    cand = []
    for s in spans:
        if s["y1"] < 50.0:                                     # 页眉
            continue
        if re.fullmatch(r"第\d{1,3}页", s["norm"]):             # 完整页脚串
            continue
        if s["y0"] > 782.0 and 255.0 <= s["x0"] <= 335.0 and (
                s["norm"] in ("第", "页") or s["norm"].isdigit()):
            continue                                            # 被拆开的页脚
        if s["size"] >= 24.0 and s["norm"].startswith("目录"):  # “目录”大标题
            continue
        cand.append(s)

    texts = [s for s in cand if not _DOT_RE.match(s["raw"])]
    dots = [s for s in cand if _DOT_RE.match(s["raw"])]
    texts.sort(key=lambda s: (s["cy"], s["x0"]))

    rows = []
    for s in texts:
        if rows and abs(s["cy"] - rows[-1]["cy"]) <= 14.0:
            grp = rows[-1]
            grp["texts"].append(s)
            grp["cy"] = sum(t["cy"] for t in grp["texts"]) / len(grp["texts"])
        else:
            rows.append({"texts": [s], "cy": s["cy"], "dots": []})
    for d in dots:
        if not rows:
            break
        best = min(rows, key=lambda r: abs(r["cy"] - d["cy"]))
        if abs(best["cy"] - d["cy"]) <= 18.0:
            best["dots"].append(d)

    for r in rows:
        ts = sorted(r["texts"], key=lambda s: s["x0"])
        r["texts"] = ts
        r["y0"] = min(s["y0"] for s in ts)
        r["y1"] = max(s["y1"] for s in ts)
        num = None
        for s in ts:
            if s["norm"].isdigit() and s["x0"] >= 430.0:
                if num is None or s["x1"] > num["x1"]:
                    num = s
        r["num_span"] = num
        r["num"] = int(num["norm"]) if num else None
        name = [s for s in ts if s is not num]
        r["name_spans"] = name
        r["name"] = "".join(s["norm"] for s in name)
        r["name_x0"] = min((s["x0"] for s in name), default=None)
        r["name_x1"] = max((s["x1"] for s in name), default=None)
        r["max_size"] = max((s["size"] for s in name), default=0.0)
    rows.sort(key=lambda r: r["cy"])
    return rows


_FONT_SUFFIX_RE = re.compile(
    r"(?:[-,_ ]?(?:Bold|Regular|Italic|Oblique|Light|Medium|Heavy|Black|"
    r"Semibold|DemiBold|Thin|Roman|MT|PS|Std|Pro))+$", re.I)


def _font_family(name):
    """剥掉子集前缀与字重后缀，得到字族名，用于“是否同一款字体”比对。"""
    n = name or ""
    if len(n) > 7 and n[6] == "+":
        n = n[7:]
    prev = None
    while prev != n:
        prev = n
        n = _FONT_SUFFIX_RE.sub("", n)
    return n.strip("-,_ ") or (name or "")


def _is_jp_font(name):
    low = (name or "").lower().replace(" ", "")
    return bool(re.search(r"cjk(jp|kr|tc|hk)", low)) or any(
        k in low for k in ("japan", "meiryo", "mincho", "msgothic", "mspgothic",
                           "yugothic", "hiragino", "kozuka", "batang", "gulim"))


def _tables(page):
    """按“划线”策略识别表格，过滤掉行列数不合理的伪表。"""
    out = []
    try:
        finder = page.find_tables(strategy="lines")
    except Exception:
        return out
    for tb in (getattr(finder, "tables", None) or []):
        try:
            nrow = int(getattr(tb, "row_count", 0) or len(tb.rows))
            ncol = int(getattr(tb, "col_count", 0) or 0)
            bb = tb.bbox
        except Exception:
            continue
        if not (2 <= nrow <= 12 and 2 <= ncol <= 8):
            continue
        out.append({"rows": nrow, "cols": ncol, "obj": tb,
                    "x0": float(bb[0]), "y0": float(bb[1]),
                    "x1": float(bb[2]), "y1": float(bb[3])})
    out.sort(key=lambda t: t["y0"])
    return out
# ========== 维度一：可用性硬门槛 ==========
def _check_dim1(pages):
    """细则维度一：交付文件为 1 个可正常打开的 PDF 文档。

    页数、页面尺寸、版式细节等均由维度二逐项计分，此处不再重复设卡；
    只要 PDF 能被正常打开且至少含 1 页有效页面即视为通过。
    返回 (是否通过, 说明)。不通过 -> 总分 0。
    按输出约定：通过时说明恒为空串，只有不通过才给出摘要文字。"""
    if not pages:
        return False, "PDF 可打开但页数为 0，无有效页面"
    return True, ""


def _error_result(result, message):
    """脚本异常/找不到文件：摘要文字统一落在 error 字段。

    此时未进入维度一判定，dim1_pass 置为空串（既非 True 也非 False），
    dim1_reason 留空，max_score 归 0。
    """
    result["status"] = "error"
    result["error"] = message
    result["dim1_pass"] = ""
    result["dim1_reason"] = ""
    result["dim2_items"] = []
    result["total_score"] = 0
    result["max_score"] = 0
    return result


def evaluate(dir_path):
    result = {"id": SCRIPT_ID, "file_name": "", "status": "ok", "error": None,
              "dim1_pass": True, "dim1_reason": "", "dim2_items": [],
              "total_score": 0, "max_score": MAX_SCORE_TOTAL}
    try:
        if not dir_path or not os.path.isdir(dir_path):
            raise FileNotFoundError(f"目录不存在：{dir_path}")
        cands = sorted(f for f in os.listdir(dir_path)
                       if f.lower().endswith('.pdf') and not f.startswith('~'))
        if not cands:
            raise FileNotFoundError(f"目录中未找到 .pdf 文件：{dir_path}")
        file_path = os.path.join(dir_path, cands[0])
        result["file_name"] = os.path.basename(file_path)
        return _evaluate_impl(file_path, result)
    except FileNotFoundError as exc:
        # 定位不到被评估文件：error 直接给出可读摘要，不带异常类型前缀。
        return _error_result(result, str(exc))
    except Exception as exc:
        return _error_result(result, f"{type(exc).__name__}: {exc}")
# ========== 维度二：逐项打分 ==========
def _evaluate_impl(file_path, result):
    doc = pymupdf.open(file_path)
    try:
        pages = []
        for i in range(doc.page_count):
            pg = doc[i]
            fills, strokes = _drawings(pg)
            pages.append({"pg": pg, "no": i + 1,
                          "w": float(pg.rect.width), "h": float(pg.rect.height),
                          "spans": _spans(pg), "imgs": _image_blocks(pg),
                          "fills": fills, "strokes": strokes,
                          "text": pg.get_text("text")})

        ok1, why1 = _check_dim1(pages)
        result["dim1_pass"] = ok1
        result["dim1_reason"] = why1
        if not ok1:
            result["dim2_items"] = []
            result["total_score"] = 0
            result["max_score"] = MAX_SCORE_TOTAL
            return result

        results, details = [], []

        def _emit(max_delta, ok, rule, detail=""):
            """加分项：命中记 +max_delta；扣分项（max_delta<0）：违规记 -|max_delta|。"""
            if max_delta >= 0:
                results.append("+%d：%s" % (max_delta if ok else 0, rule))
            else:
                results.append("-%d：%s" % (0 if ok else -max_delta, rule))
            details.append(detail)

        def _pg(no):
            return pages[no - 1] if 1 <= no <= len(pages) else None

        def _ptext(no):
            p = _pg(no)
            return _norm(p["text"]) if p else ""

        # ---------- 全册公共量 ----------
        # 版心左右边界用“通栏图元”测量（文字不一定顶到右边界）
        wide = [(f["x0"], f["x1"]) for p in pages for f in (p["fills"] + p["strokes"])
                if f["w"] >= 400.0]
        m_left = statistics.median([a for a, _ in wide]) if wide else None
        m_right = statistics.median([b for _, b in wide]) if wide else None

        toc = pages[TOC_PIDX]
        tspans, tfills, tstrokes = toc["spans"], toc["fills"], toc["strokes"]
        timgs = toc["imgs"]
        rows = _toc_rows(tspans)
        PIXS = 2.0
        pix = _render(toc["pg"], PIXS)

        # “目录”大标题
        title = None
        for s in tspans:
            if s["norm"].startswith("目录") and len(s["norm"]) <= 4 and s["size"] >= 18.0:
                if title is None or s["size"] > title["size"]:
                    title = s

        # 四个单元序号（旗标内白字）
        unit_labels = sorted([s for s in tspans if re.fullmatch(r"第[一二三四]单元", s["norm"])],
                             key=lambda s: s["cy"])
        def _flag_for(lab):
            """找出包住单元序号的绿色旗标；判断是否带右端尖角（五边形/矩形+三角）。"""
            lh = max(1.0, lab["y1"] - lab["y0"])
            cands = [f for f in tfills if _is_greenish(f["rgb"])
                     and _overlap(f["y0"], f["y1"], lab["y0"], lab["y1"]) >= 0.45 * lh
                     and f["x0"] <= lab["x1"] + 45.0 and f["x1"] >= lab["x0"] - 45.0
                     and f["w"] >= 6.0 and f["h"] >= 8.0]
            if not cands:
                return None
            ux0 = min(f["x0"] for f in cands)
            ux1 = max(f["x1"] for f in cands)
            tip = any(f["nseg"] >= 5 and "re" not in f["segs"] for f in cands)
            if not tip:
                rects = [f for f in cands if "re" in f["segs"] or f["nseg"] >= 4]
                tris = [f for f in cands if f["nseg"] == 3 and "re" not in f["segs"]]
                for rc in rects:
                    for tr in tris:
                        if abs(tr["x0"] - rc["x1"]) <= 3.5 and tr["x1"] > rc["x1"]:
                            tip = True
            return {"x0": ux0, "x1": ux1, "n": len(cands), "tip": tip,
                    "covers": ux0 <= lab["x0"] + 2.0 and ux1 >= lab["x1"] - 2.0,
                    "rgbs": [f["rgb"] for f in cands]}

        flags = [_flag_for(s) for s in unit_labels]
        flag_left = min([f["x0"] for f in flags if f], default=None)
        if flag_left is None:
            flag_left = min([s["x0"] for s in unit_labels], default=BODY_LEFT)

        # 行 -> 条目 的对应关系（严格按纵向次序与细则顺序配对）
        matched = []
        if len(rows) == len(TOC_ENTRIES):
            for idx, (kind, label, name, pno) in enumerate(TOC_ENTRIES):
                rn = rows[idx]["name"]
                hit = (_norm(name) in rn) and (_norm(label) in rn)
                matched.append(hit)
        row_by_entry = {}
        if matched and all(matched):
            for idx, ent in enumerate(TOC_ENTRIES):
                row_by_entry[idx] = rows[idx]

        def _rows_of(kind):
            return [rows[i] for i in row_by_entry if TOC_ENTRIES[i][0] == kind]

        task_rows = _rows_of("task")
        unit_rows = _rows_of("unit")
        proj_rows = _rows_of("project")
        intro_rows = _rows_of("intro")
        task_size = max([r["max_size"] for r in task_rows], default=0.0)

        # 单元名称 span（旗标右侧的黑色加粗名称）
        unit_name_spans = []
        for i in row_by_entry:
            if TOC_ENTRIES[i][0] != "unit":
                continue
            key = _norm(TOC_ENTRIES[i][2])
            got = [s for s in rows[i]["name_spans"] if key[:3] in s["norm"]]
            if got:
                unit_name_spans.append(min(got, key=lambda s: s["x0"]))

        # 页眉基准（取第3页页眉，用于“与其余各页一致”比对）
        def _header_of(no):
            p = _pg(no)
            if not p:
                return None
            hs = _header_spans(p["spans"])
            if not hs:
                return None
            return {"spans": hs, "norm": _norm("".join(s["raw"] for s in
                                                      sorted(hs, key=lambda s: s["x0"]))),
                    "x1": max(s["x1"] for s in hs), "y0": min(s["y0"] for s in hs),
                    "size": max(s["size"] for s in hs),
                    "rgbs": [s["rgb"] for s in hs]}

        ref_header = _header_of(3) or _header_of(4)
        # ---- 01 封面后新增独立目录页、16→17页、其余各页内容与次序不变 ----
        cover_ok = any(s["norm"].startswith("七年级地理导学案") and s["size"] >= 26.0
                       for s in pages[0]["spans"])
        toc_is_toc = title is not None and len(rows) >= 14
        order_bad = []
        for no in sorted(BODY_TITLES):
            txt = _ptext(no)
            miss = [k for k in BODY_TITLES[no] if _norm(k) not in txt]
            if miss:
                order_bad.append("第%d页缺「%s」" % (no, "/".join(miss)))
        toc_pure = not any(k in _norm(toc["text"])
                           for k in ("真实情境", "分层练习", "核心概念整理"))
        ok01 = (len(pages) == EXPECT_PAGES and cover_ok and toc_is_toc
                and not order_bad and toc_pure)
        _emit(5, ok01,
              "封面页之后新增 1 个独立整页目录，全册由 16 页变为 17 页，其余各页内容与先后次序保持不变",
              "页数=%d；封面主标题=%s；第2页目录行=%d；正文页序校验=%s"
              % (len(pages), "有" if cover_ok else "无", len(rows),
                 "全部匹配" if not order_bad else "；".join(order_bad[:3])))

        # ---- 02 “目录”标题：左上、绿色、加粗、≥30磅 ----
        if title is None:
            _emit(1, False, "“目录”标题位于页面左上部、绿色、加粗、字号不小于30磅", "未找到“目录”标题")
        else:
            tr, tg, tb = title["rgb"]
            c_ok = (10 <= tr <= 60 and 140 <= tg <= 160 and 50 <= tb <= 80)
            ok02 = (c_ok and title["bold"] and title["size"] >= 29.5
                    and title["x0"] <= 200.0 and title["y0"] <= 130.0)
            _emit(1, ok02, "“目录”标题位于页面左上部、绿色、加粗、字号不小于30磅",
                  "字号=%.1f磅；颜色=#%02X%02X%02X(%s)；加粗=%s；位置=(%.1f,%.1f)"
                  % (title["size"], tr, tg, tb, "在绿色区间" if c_ok else "不在RGB(10-60,140-160,50-80)",
                     title["bold"], title["x0"], title["y0"]))

        # ---- 03 标题左侧紧贴页边的绿色斜条纹装饰 ----
        d03 = "未找到“目录”标题"
        ok03 = False
        if title is not None:
            zx1 = title["x0"] + 40.0
            zy0, zy1 = title["y0"] - 24.0, title["y1"] + 14.0
            deco = []
            for f in tfills:
                if f["x1"] <= zx1 and f["w"] >= 3.0 \
                        and _overlap(f["y0"], f["y1"], zy0, zy1) >= 4.0:
                    deco.append((f["rgb"], f["x0"], f["x1"], f["y0"], f["y1"], "矢量"))
            for im in timgs:
                if im["x1"] <= zx1 \
                        and _overlap(im["y0"], im["y1"], zy0, zy1) >= 4.0:
                    c = _pix_main(pix, PIXS, im["x0"], im["y0"], im["x1"], im["y1"])
                    deco.append((c, im["x0"], im["x1"], im["y0"], im["y1"], "位图"))
            greens = [d for d in deco if _is_greenish(d[0])]
            shades = {tuple(v // 24 for v in d[0]) for d in greens}
            touch = min([d[1] for d in greens], default=99.0)
            ok03 = (3 <= len(greens) <= 8 and len(shades) >= 2 and touch <= 12.0)
            d03 = ("绿色装饰元素=%d个（%s）；深浅层次=%d种；最左边界=%.1fpt"
                   % (len(greens), "/".join("%s#%02X%02X%02X" % (d[5], d[0][0], d[0][1], d[0][2])
                                            for d in greens[:5]), len(shades), touch))
        _emit(3, ok03, "“目录”标题左侧紧贴页面左边缘有绿色斜条纹装饰（1个深色块+2-3条渐浅斜条）", d03)

        # ---- 04 标题下方横向绿色分隔线/色带 ----
        d04 = "未找到“目录”标题"
        ok04 = False
        if title is not None:
            by0, by1 = title["y1"] - 8.0, title["y1"] + 34.0
            bars = []
            for f in tfills:
                if f["w"] >= 200.0 and f["h"] >= 1.0 and _in_band(f, by0, by1, 0.5) \
                        and _is_greenish(f["rgb"]):
                    bars.append((f["x0"], f["x1"], f["h"], f["rgb"]))
            for st in tstrokes:
                if st["w"] >= 200.0 and _in_band(st, by0, by1, 0.5) and _is_greenish(st["rgb"]):
                    bars.append((st["x0"], st["x1"], max(st["width"], st["h"]), st["rgb"]))
            for im in timgs:
                if im["w"] >= 200.0 and _in_band(im, by0, by1, 0.5):
                    c = _pix_main(pix, PIXS, im["x0"], im["y0"], im["x1"], im["y1"])
                    if _is_greenish(c):
                        bars.append((im["x0"], im["x1"], im["h"], c))
            need_x1 = BODY_LEFT + 0.90 * (BODY_RIGHT - BODY_LEFT)
            reach = max([b[1] for b in bars], default=0.0)
            thick = max([b[2] for b in bars], default=0.0)
            start = min([b[0] for b in bars], default=999.0)
            ok04 = (bool(bars) and reach >= need_x1 and thick >= 2.0
                    and start <= title["x0"] + 22.0)
            d04 = ("绿色横带=%d条；起点x=%.1f，终点x=%.1f（需≥%.1f）；最大线宽=%.2fpt"
                   % (len(bars), start, reach, need_x1, thick))
        _emit(3, ok04, "“目录”标题下方有横向绿色分隔线/色带，自装饰块右侧横贯至版心右边界、线宽不小于2pt", d04)
        # ---- 05 目录共 18 条条目、无遗漏/重复/杜撰 ----
        d05 = []
        if len(rows) != len(TOC_ENTRIES):
            d05.append("条目行数=%d，应为18" % len(rows))
        else:
            for idx, (kind, label, name, pno) in enumerate(TOC_ENTRIES):
                if not matched[idx]:
                    d05.append("第%d行应为「%s %s」，实为「%s」"
                               % (idx + 1, label, name, rows[idx]["name"][:22]))
        allnames = [r["name"] for r in rows]
        dup = sorted({n for n in allnames if allnames.count(n) > 1})
        if dup:
            d05.append("重复条目：%s" % "、".join(x[:14] for x in dup[:3]))
        ok05 = (len(rows) == len(TOC_ENTRIES) and bool(matched) and all(matched) and not dup)
        _emit(5, ok05, "目录共 18 条条目（1条说明+4个单元+10个任务+3个项目/档案袋），无遗漏、重复或杜撰",
              "识别到 %d 条；%s" % (len(rows), "与细则清单逐条一致" if ok05 else "；".join(d05[:3])))

        # ---- 06 条目名称与其所指页面标题文字一致 ----
        d06 = []
        for idx, (kind, label, name, pno) in enumerate(TOC_ENTRIES):
            if idx not in row_by_entry:
                continue
            if _norm(name) not in _ptext(pno):
                d06.append("「%s」未出现在第%d页" % (name, pno))
        ok06 = bool(row_by_entry) and not d06
        _emit(3, ok06, "目录中每条条目的名称文字与其所指页面上的标题文字完全一致，无错字漏字",
              "18 条名称全部能在对应页面找到" if ok06 else
              ("条目未全部匹配，无法逐条核对" if not row_by_entry else "；".join(d06[:3])))

        # ---- 07 四个单元序号位于绿色五边形旗标内、白色加粗 ----
        d07 = []
        for lab, fg in zip(unit_labels, flags):
            if fg is None:
                d07.append("「%s」外无绿色旗标" % lab["norm"])
                continue
            if not fg["covers"]:
                d07.append("「%s」未被旗标包住" % lab["norm"])
            if not fg["tip"]:
                d07.append("「%s」旗标无右端三角尖角" % lab["norm"])
            if not _is_white(lab["rgb"]):
                d07.append("「%s」非白色(#%02X%02X%02X)" % ((lab["norm"],) + lab["rgb"]))
            if not lab["bold"]:
                d07.append("「%s」未加粗" % lab["norm"])
        ok07 = (len(unit_labels) == 4 and all(f is not None for f in flags) and not d07)
        _emit(5, ok07, "四个单元序号位于绿色五边形旗标（矩形+右端三角尖角）内，文字白色加粗",
              "单元序号=%d个，旗标齐备且白色加粗" % len(unit_labels) if ok07
              else "单元序号=%d个；%s" % (len(unit_labels), "；".join(d07[:3])))

        # ---- 08 单元名称在旗标右侧、黑色加粗 ----
        d08 = []
        if len(unit_name_spans) != 4:
            d08.append("仅识别到 %d 个单元名称" % len(unit_name_spans))
            ok08 = False
        else:
            xs = [s["x0"] for s in unit_name_spans]
            usize = min(s["size"] for s in unit_name_spans)
            right_ok = all(any(fg and s["x0"] >= fg["x1"] - 3.0 for fg in flags)
                           for s in unit_name_spans)
            dark_ok = all(_is_dark(s["rgb"]) for s in unit_name_spans)
            bold_ok = all(s["bold"] for s in unit_name_spans)
            ok08 = (right_ok and dark_ok and bold_ok)
            # 左边界与字号仅作为信息展示，不参与判定
            d08.append("左边界=%s（极差%.1fpt，仅供参考）"
                       % (["%.1f" % v for v in xs], _spread(xs)))
            d08.append("字号=%.1f 磅（任务 %.1f 磅，仅供参考）" % (usize, task_size))
            if not right_ok:
                d08.append("存在名称未位于旗标右侧")
            if not dark_ok:
                d08.append("存在非黑色名称")
            if not bold_ok:
                d08.append("存在未加粗名称")
        _emit(3, ok08, "单元名称紧随旗标右侧，黑色加粗",
              "；".join(d08[:4]))

        # ---- 09 任务 1-10 缩进≥30pt、常规体、15-17磅 ----
        d09 = []
        if len(task_rows) != 10:
            d09.append("仅识别到 %d 条任务" % len(task_rows))
            ok09 = False
        else:
            inds = [r["name_x0"] - flag_left for r in task_rows]
            sizes = [s["size"] for r in task_rows for s in r["name_spans"]]
            bolds = [s["norm"] for r in task_rows for s in r["name_spans"] if s["bold"]]
            ok09 = (min(inds) >= 30.0 and not bolds
                    and 15.0 <= min(sizes) and max(sizes) <= 17.0)
            d09.append("缩进=%.1f~%.1fpt（限≥30）" % (min(inds), max(inds)))
            d09.append("字号=%.1f~%.1f磅（限15-17）" % (min(sizes), max(sizes)))
            d09.append("加粗字段=%d处" % len(bolds))
        _emit(3, ok09, "任务1-10 条目相对单元旗标向右缩进不少于30pt，字体为常规体（不加粗），字号15-17磅",
              "；".join(d09[:3]))
        # ---- 10 三个项目条目前有绿色三角/箭头符号、加粗、≥17磅且大于任务 ----
        d10 = []
        if len(proj_rows) != 3:
            ok10 = False
            d10.append("仅识别到 %d 条项目/档案袋条目" % len(proj_rows))
        else:
            marks = []
            for r in proj_rows:
                rh = max(1.0, r["y1"] - r["y0"])
                got = [f for f in tfills if _is_greenish(f["rgb"])
                       and 4.0 <= f["w"] <= 34.0 and 4.0 <= f["h"] <= 34.0
                       and f["x1"] <= r["name_x0"] + 3.0
                       and _overlap(f["y0"], f["y1"], r["y0"], r["y1"]) >= 0.30 * rh]
                glyph = [s for s in tspans if _is_greenish(s["rgb"])
                         and s["norm"] in ("▶", "►", "▸", "➤", "→", "▲", "◆", "»")
                         and s["x1"] <= r["name_x0"] + 3.0
                         and _overlap(s["y0"], s["y1"], r["y0"], r["y1"]) >= 0.30 * rh]
                marks.append(len(got) + len(glyph))
            sizes = [s["size"] for r in proj_rows for s in r["name_spans"]]
            nb = [s["norm"] for r in proj_rows for s in r["name_spans"] if not s["bold"]]
            ok10 = (all(m > 0 for m in marks) and not nb and min(sizes) >= 16.5
                    and task_size > 0 and min(sizes) > task_size + 0.2)
            d10.append("前置绿色符号=%s" % marks)
            d10.append("字号=%.1f~%.1f磅（限≥17且>任务%.1f）" % (min(sizes), max(sizes), task_size))
            if nb:
                d10.append("未加粗=%d处" % len(nb))
        _emit(3, ok10, "三个项目/档案袋条目前有绿色三角或箭头符号，文字加粗、字号不小于17磅且大于任务条目",
              "；".join(d10[:3]))

        # ---- 11 “学习结构说明”首条、加粗、与单元同级不缩进 ----
        if not intro_rows:
            _emit(1, False, "“学习结构说明”位于目录第一条，加粗，与单元条目同级左对齐不缩进",
                  "未识别到“学习结构说明”条目")
        else:
            r = intro_rows[0]
            is_first = bool(rows) and rows[0] is r
            bold_ok = all(s["bold"] for s in r["name_spans"])
            align = abs(r["name_x0"] - flag_left)
            ok11 = is_first and bold_ok and align <= 6.0
            _emit(1, ok11, "“学习结构说明”位于目录第一条，加粗，与单元条目同级左对齐不缩进",
                  "序位=%d/18；加粗=%s；左边界=%.1f vs 旗标%.1f（差%.1fpt）"
                  % (rows.index(r) + 1, bold_ok, r["name_x0"], flag_left, align))

        # ---- 12 点状前导引导线 ----
        def _leader(r):
            """定位该行的引导点：优先点字符 span，其次成排小圆点填充。"""
            if r["dots"]:
                dd = r["dots"]
                return (min(d["x0"] for d in dd), max(d["x1"] for d in dd),
                        True, "点字符%d段" % len(dd))
            rh = max(1.0, r["y1"] - r["y0"])
            dots = [f for f in tfills if f["w"] <= 5.0 and f["h"] <= 5.0
                    and f["x0"] >= (r["name_x0"] or 0.0)
                    and _overlap(f["y0"], f["y1"], r["y0"] - 2.0, r["y1"] + 2.0) >= 0.2 * min(rh, f["h"] + 1.0)]
            if len(dots) < 5:
                return (None, None, False, "引导点=%d个" % len(dots))
            cx = sorted((f["x0"] + f["x1"]) / 2.0 for f in dots)
            gaps = [cx[i + 1] - cx[i] for i in range(len(cx) - 1)]
            med = statistics.median(gaps) if gaps else 0.0
            cont = bool(med) and max(gaps) <= max(3.0 * med, med + 2.0)
            return (min(f["x0"] for f in dots), max(f["x1"] for f in dots), cont,
                    "圆点%d个/间距%.1fpt/最大断口%.1fpt" % (len(dots), med, max(gaps) if gaps else 0.0))

        d12 = []
        if len(rows) != len(TOC_ENTRIES):
            ok12 = False
            d12.append("条目行未能完整识别（%d 条）" % len(rows))
        else:
            ok12 = True
            covs = []
            for i, r in enumerate(rows):
                lx0, lx1, cont, why = _leader(r)
                nx1, px0 = r["name_x1"], (r["num_span"]["x0"] if r["num_span"] else None)
                if lx0 is None or nx1 is None or px0 is None:
                    ok12 = False
                    d12.append("第%d行无引导线（%s）" % (i + 1, why))
                    continue
                gap = px0 - nx1
                cov = (lx1 - lx0) / gap if gap > 0 else 0.0
                covs.append(cov)
                if gap <= 0:
                    ok12 = False
                    d12.append("第%d行名称与页码重叠" % (i + 1))
                if lx0 < nx1 - 2.0:
                    ok12 = False
                    d12.append("第%d行引导线压住名称" % (i + 1))
                if lx1 > px0 + 2.0:
                    ok12 = False
                    d12.append("第%d行引导线越过页码" % (i + 1))
                if cov < 0.55:
                    ok12 = False
                    d12.append("第%d行引导线仅覆盖%.0f%%" % (i + 1, 100 * cov))
                if not cont:
                    ok12 = False
                    d12.append("第%d行引导线断裂（%s）" % (i + 1, why))
            if covs:
                d12.insert(0, "覆盖率%.0f%%~%.0f%%" % (100 * min(covs), 100 * max(covs)))
        _emit(5, ok12, "全部 18 条名称与页码之间有点状前导引导线，不与文字重叠、不断裂、不越过页码",
              "；".join(d12[:4]) if d12 else "18 条引导线齐全")
        # ---- 13 页码右对齐并与版心右边界齐平 ----
        nums = [r["num_span"] for r in rows if r["num_span"]]
        if len(nums) != len(TOC_ENTRIES):
            _emit(3, False, "18 条页码右对齐且与版心右边界齐平",
                  "仅识别到 %d 个页码" % len(nums))
        else:
            xs = [s["x1"] for s in nums]
            edge = statistics.median(xs)
            ok13 = abs(edge - BODY_RIGHT) <= 6.0
            # 各条右边界极差仅作为信息展示，不参与判定
            _emit(3, ok13, "18 条页码右对齐且与版心右边界齐平",
                  "右边界 %.1f~%.1f（极差%.1fpt，仅供参考）；中位数%.1f vs 版心右%.1f"
                  % (min(xs), max(xs), _spread(xs), edge, BODY_RIGHT))

        # ---- 14 目录所标页码与条目实际所在页完全对应 ----
        d14 = []
        for idx, (kind, label, name, pno) in enumerate(TOC_ENTRIES):
            if idx not in row_by_entry:
                continue
            got = rows[idx]["num"]
            if got != pno:
                d14.append("「%s」标%s应为%d" % (name, got, pno))
        ok14 = bool(row_by_entry) and not d14
        _emit(5, ok14, "目录标注的页码与各条目在新文档中的实际所在页完全对应（3,4,4,5,6,8,8,8,10,10,11,12,13,13,14,15,16,17）",
              "18 条页码全部正确" if ok14 else
              ("条目未全部匹配，无法逐条核对" if not row_by_entry
               else "错%d条：%s" % (len(d14), "；".join(d14[:4]))))

        # ---- 15 全册页脚“第N页”与物理页序一致 ----
        d15 = []
        for p in pages:
            got = _footer_pageno(p["spans"])
            if got != p["no"]:
                d15.append("第%d页页脚=%s" % (p["no"], ("第%d页" % got) if got else "缺失"))
        ok15 = not d15
        _emit(5, ok15, "全册页脚“第N页”与物理页序一致（目录页为第2页、末页为第17页），插入目录后完成重新编号",
              "17 页页脚编号连续正确" if ok15 else "错%d页：%s" % (len(d15), "；".join(d15[:5])))

        # ---- 16 目录页页眉 ----
        th = _header_of(TOC_PIDX + 1)
        if th is None:
            _emit(3, False, "目录页顶部保留与其余各页一致的页眉（灰色#646464、8-9磅、右边界≈551pt、上边界≈36pt）",
                  "目录页无页眉")
        else:
            same = True
            if ref_header:
                same = (th["norm"] == ref_header["norm"]
                        and abs(th["x1"] - ref_header["x1"]) <= 2.5
                        and abs(th["y0"] - ref_header["y0"]) <= 2.5)
            gray = all(_cdist(c, C_HEADGRAY) <= 24 for c in th["rgbs"])
            ok16 = (_norm(HEADER_TEXT) in th["norm"] and gray
                    and 8.0 <= th["size"] <= 9.0
                    and abs(th["x1"] - 551.0) <= 4.0 and abs(th["y0"] - 36.0) <= 4.0
                    and same)
            _emit(3, ok16, "目录页顶部保留与其余各页一致的页眉（灰色#646464、8-9磅、右边界≈551pt、上边界≈36pt）",
                  "文字=「%s」；字号%.1f磅；灰色=%s；右%.1f/上%.1f；与第3页页眉一致=%s"
                  % (th["norm"][:22], th["size"], gray, th["x1"], th["y0"], same))

        # ---- 17 目录页页脚样式与位置（页码内容归第 15 项）----
        fcand = [s for s in _footer_spans(tspans) if 200.0 <= s["x1"] and s["x0"] <= 400.0]
        if not fcand:
            _emit(3, False, "目录页底部有居中页脚，灰色 8-10.5 磅，纵向位置约 790-806pt", "目录页无页脚")
        else:
            ftxt = _norm("".join(s["raw"] for s in sorted(fcand, key=lambda s: s["x0"])))
            fx0 = min(s["x0"] for s in fcand)
            fx1 = max(s["x1"] for s in fcand)
            fy0 = min(s["y0"] for s in fcand)
            fsz = [s["size"] for s in fcand]
            fgray = any(_cdist(s["rgb"], C_HEADGRAY) <= 30 or _is_neutral(s["rgb"]) for s in fcand)
            ok17 = (fgray and 8.0 <= min(fsz) and max(fsz) <= 10.5
                    and _centered(fx0, fx1, 14.0) and 788.0 <= fy0 <= 808.0)
            _emit(3, ok17, "目录页底部有居中页脚，灰色 8-10.5 磅，纵向位置约 790-806pt",
                  "文字=「%s」；字号%.1f~%.1f磅；灰色=%s；居中偏移%.1fpt；上边界%.1fpt"
                  % (ftxt[:12], min(fsz), max(fsz), fgray,
                     abs((fx0 + fx1) / 2.0 - PAGE_CX), fy0))

        # ---- 18 18 条全部排在同一页且落在版心内 ----
        d18 = []
        if len(rows) != len(TOC_ENTRIES):
            ok18 = False
            d18.append("条目未全部落在第2页（识别到 %d 条）" % len(rows))
        else:
            top = min(r["y0"] for r in rows)
            bot = max(r["y1"] for r in rows)
            lft = min(r["name_x0"] for r in rows)
            rgt = max(r["num_span"]["x1"] if r["num_span"] else r["name_x1"] for r in rows)
            if title is not None:
                top = min(top, title["y0"])
                lft = min(lft, title["x0"])
            ok18 = (top >= BODY_TOP - 0.5 and bot <= BODY_BOTTOM
                    and lft >= BODY_LEFT - 3.0 and rgt <= BODY_RIGHT + 4.0)
            d18.append("上%.1f（限≥36）/下%.1f（限≤790）/左%.1f/右%.1f" % (top, bot, lft, rgt))
        _emit(5, ok18, "18 条条目全部排在同一页内，全部文字与图形落在版心内（上边界≥36pt、下边界≤790pt、不侵入页脚区）",
              "；".join(d18[:2]))

        # ---- 19 各绿色元素同一色系，全页无第三种强调色 ----
        gset, others = [], []
        if title is not None and _is_greenish(title["rgb"]):
            gset.append(title["rgb"])
        for fg in flags:
            if fg:
                gset.extend([c for c in fg["rgbs"] if _is_greenish(c)])
        for el in tfills + tstrokes:
            if _is_greenish(el["rgb"]):
                gset.append(el["rgb"])
            elif not _is_neutral(el["rgb"]) and el["rgb"] is not None:
                others.append(("图形", el["rgb"]))
        for s in tspans:
            if _is_greenish(s["rgb"]):
                gset.append(s["rgb"])
            elif not _is_neutral(s["rgb"]):
                others.append(("文字「%s」" % s["norm"][:6], s["rgb"]))
        for im in timgs:
            for c, n in _pix_colors(pix, PIXS, im["x0"], im["y0"], im["x1"], im["y1"])[:4]:
                if _is_greenish(c):
                    gset.append(c)
                elif n >= 3:
                    others.append(("位图", c))
        strong = [c for c in gset if (max(c) - min(c)) >= 40]
        hue_ok = all(abs(c[0] - c[2]) <= 0.55 * c[1] for c in strong)
        ok20 = bool(gset) and hue_ok and not others
        _emit(1, ok20, "目录页各绿色元素（标题、装饰、分隔带、旗标、项目符号）同属一个绿色色系，全页未混入第三种强调色",
              "绿色元素%d处（其中饱和%d处，同色系=%s）；非绿非中性色=%s"
              % (len(gset), len(strong), hue_ok,
                 "无" if not others else "、".join("%s#%02X%02X%02X" % (o[0], o[1][0], o[1][1], o[1][2])
                                                  for o in others[:3])))

        # ---- 20 文件体积 ≤5MB ----
        fsize = os.path.getsize(file_path)
        maxfont, seen_xref = 0, set()
        for p in pages:
            try:
                flist = p["pg"].get_fonts(full=True)
            except Exception:
                flist = []
            for fi in flist:
                xr = fi[0] if fi else 0
                if not xr or xr in seen_xref:
                    continue
                seen_xref.add(xr)
                try:
                    ff = doc.extract_font(xr)
                    buf = ff[3] if ff and len(ff) > 3 else b""
                    maxfont = max(maxfont, len(buf or b""))
                except Exception:
                    pass
        ok21 = fsize <= 5 * 1024 * 1024
        _emit(-3, ok21, "文件体积不超过 5MB（未因嵌入未子集化的巨型字体流而膨胀到 20MB 以上）",
              "文件 %.2fMB（限5MB）；最大单个嵌入字体流 %.2fMB"
              % (fsize / 1048576.0, maxfont / 1048576.0))

        # ---- 21 中文字体同族、未混入日文字形字体 ----
        fam_cnt, jp_fonts = {}, set()
        for p in pages:
            for s in _body_spans(p["spans"]):
                if not _CJK_RE.search(s["raw"]):
                    continue
                fam_cnt[_font_family(s["font"])] = fam_cnt.get(_font_family(s["font"]), 0) + 1
                if _is_jp_font(s["font"]):
                    jp_fonts.add(s["font"])
        ok22 = len(fam_cnt) == 1 and not jp_fonts
        _emit(-5, ok22, "目录与正文中文字体同属一款简体中文黑体字族，未混入日文字形字体",
              "中文字族=%s；日文/其他地区字形=%s"
              % ("、".join("%s(%d处)" % (k, v) for k, v in
                          sorted(fam_cnt.items(), key=lambda kv: -kv[1])[:3]) or "无",
                 "、".join(sorted(jp_fonts)[:2]) or "无"))

        # ---- 22 “任务”与序号之间无多余空格 ----
        bad23 = []
        for r in task_rows:
            for s in r["name_spans"]:
                if re.search(r"任[\s 　]+务", s["raw"]) or \
                        re.search(r"务[\s 　]+\d", s["raw"]):
                    bad23.append(repr(s["raw"][:12]))
        ok23 = bool(task_rows) and not bad23
        _emit(-1, ok23, "目录条目中“任务”二字与其后序号之间没有被多余空格拆开",
              "10 条任务标签无多余空格" if ok23 else
              ("未识别到任务条目" if not task_rows else "异常%d处：%s" % (len(bad23), "、".join(bad23[:3]))))

        # ---- 23 文档属性 ----
        md = doc.metadata or {}
        mt = (md.get("title") or "").strip()
        ma = (md.get("author") or "").strip()
        ms = (md.get("subject") or "").strip()
        ok24 = (mt == "七年级地理导学案" and ma == "地理教研组" and ms == "项目化学习任务单")
        _emit(-1, ok24, "文档属性中标题为“七年级地理导学案”、作者为“地理教研组”、主题为“项目化学习任务单”",
              "标题=「%s」/作者=「%s」/主题=「%s」" % (mt, ma, ms))
        # ---- 25~30 第1页封面 ----
        p1 = pages[0]

        def _lines_of(spans, tol=3.0):
            """把 span 列表按纵向聚成行，行内按 x 排序。"""
            out = []
            for s in sorted(spans, key=lambda z: (z["cy"], z["x0"])):
                if out and abs(s["cy"] - out[-1][0]) <= tol:
                    out[-1][1].append(s)
                else:
                    out.append([s["cy"], [s]])
            return [sorted(g, key=lambda z: z["x0"]) for _, g in out]

        # ---- 24 第1页主标题/副标题/学科行/导语 四行文字层级样式（合并项）----
        d25 = []

        t25 = [s for s in p1["spans"]
               if s["norm"].startswith("七年级地理导学案") and s["size"] >= 20.0]
        if not t25:
            d25.append("主标题未找到")
        else:
            s = max(t25, key=lambda z: z["size"])
            if not (_cdist(s["rgb"], C_DARKBLUE) <= 22 and s["bold"]
                    and _near(s["size"], 30.0, 1.0) and _centered(s["x0"], s["x1"], 10.0)):
                d25.append("主标题#%02X%02X%02X/%.1f磅/加粗%s/偏移%.1fpt"
                           % (s["rgb"][0], s["rgb"][1], s["rgb"][2], s["size"], s["bold"],
                              abs((s["x0"] + s["x1"]) / 2.0 - PAGE_CX)))

        c26 = [s for s in p1["spans"] if s["y0"] > 110.0 and s["size"] >= 12.0
               and ("项目化学习任务单" in s["norm"] or "学生用" in s["norm"])]
        if not c26:
            d25.append("副标题未找到")
        else:
            ln = _lines_of(c26)[0]
            x0 = min(s["x0"] for s in ln)
            x1 = max(s["x1"] for s in ln)
            sz = max(s["size"] for s in ln)
            if not (all(_cdist(s["rgb"], C_GREEN) <= 22 for s in ln)
                    and all(s["bold"] for s in ln) and _near(sz, 16.0, 1.0)
                    and _centered(x0, x1, 12.0)):
                d25.append("副标题#%02X%02X%02X/%.1f磅/加粗%s/偏移%.1fpt"
                           % (ln[0]["rgb"][0], ln[0]["rgb"][1], ln[0]["rgb"][2], sz,
                              all(s["bold"] for s in ln), abs((x0 + x1) / 2.0 - PAGE_CX)))

        c27 = [s for s in p1["spans"] if s["y0"] > 110.0
               and any(k in s["norm"] for k in ("学科：", "年级：", "课型："))]
        if not c27:
            d25.append("学科/年级/课型行未找到")
        else:
            lns = _lines_of(c27)
            ln = lns[0]
            x0 = min(s["x0"] for s in ln)
            x1 = max(s["x1"] for s in ln)
            szs = [s["size"] for s in ln]
            if not (len(lns) == 1 and all(_is_dark(s["rgb"], 45) for s in ln)
                    and not any(s["bold"] for s in ln)
                    and _near(min(szs), 12.0, 0.7) and _near(max(szs), 12.0, 0.7)
                    and _centered(x0, x1, 12.0)):
                d25.append("学科行%d行/%.1f磅/加粗%s/偏移%.1fpt"
                           % (len(lns), min(szs), any(s["bold"] for s in ln),
                              abs((x0 + x1) / 2.0 - PAGE_CX)))

        nine_keys = {_norm(x) for x in NINE_CELLS}
        c28 = [s for s in p1["spans"] if 212.0 < s["y0"] < 282.0
               and 9.4 <= s["size"] <= 11.6 and not s["bold"] and s["norm"] not in nine_keys]
        lns28 = _lines_of(c28)
        if len(lns28) != 2:
            d25.append("导语识别到%d行（应为2行）" % len(lns28))
        else:
            szs = [s["size"] for ln in lns28 for s in ln]
            offs = [abs((min(s["x0"] for s in ln) + max(s["x1"] for s in ln)) / 2.0 - PAGE_CX)
                    for ln in lns28]
            if not (all(_is_dark(s["rgb"], 45) for ln in lns28 for s in ln)
                    and _near(min(szs), 10.5, 0.8) and _near(max(szs), 10.5, 0.8)
                    and max(offs) <= 16.0):
                d25.append("导语%.1f~%.1f磅/偏移%.1f、%.1fpt"
                           % (min(szs), max(szs), offs[0], offs[1]))

        _emit(-2, not d25,
              "第1页主标题(#1F4E79 加粗30磅)、副标题(#548235 加粗16磅)、"
              "学科/年级/课型行(黑色常规12磅单行)、导语(黑色常规10.5磅两行)均居中且样式正确",
              "四行文字层级样式均正确" if not d25 else "；".join(d25[:4]))

        # ---- 25 第1页九宫格 ----
        cells = [s for s in p1["spans"] if s["norm"] in nine_keys]
        d29 = []
        if len(cells) != 9:
            ok29 = False
            d29.append("九宫格文字识别到 %d 个" % len(cells))
        else:
            rws = _lines_of(cells, 6.0)
            cols = len({round((s["x0"] + s["x1"]) / 2.0 / 40.0) for s in cells})
            shades = []
            for c in cells:
                cover = [f for f in p1["fills"]
                         if f["x0"] <= c["x0"] + 1.0 and f["x1"] >= c["x1"] - 1.0
                         and f["y0"] <= c["y0"] + 1.0 and f["y1"] >= c["y1"] - 1.0
                         and f["w"] <= 220.0 and f["rgb"] and min(f["rgb"]) >= 185]
                if cover:
                    shades.append(min(cover, key=lambda z: z["w"] * z["h"])["rgb"])
            nshade = len({tuple(v // 8 for v in c) for c in shades})
            ok29 = (len(rws) == 3 and all(len(r) == 3 for r in rws) and cols == 3
                    and all(_cdist(s["rgb"], C_DARKGRAY) <= 22 for s in cells)
                    and all(s["bold"] for s in cells)
                    and all(_near(s["size"], 11.0, 0.7) for s in cells)
                    and len(shades) == 9 and nshade == 9)
            d29.append("%d行×%d列 / %.1f磅 / 底纹%d块共%d种色"
                       % (len(rws), cols, cells[0]["size"], len(shades), nshade))
        _emit(-1, ok29, "第1页九宫格为 3 行 3 列、文字 #404040 加粗 11 磅，9 个格子各有 1 种浅色底纹",
              "；".join(d29[:2]))

        # 第1页底部四格：先算出结论，与任务页四格表合并为一项（见新 28 项）
        d_four = []
        lab30 = [s for s in p1["spans"] if s["y0"] > 300.0
                 and any(k[:2] in s["norm"] for k in FOUR_LABELS)]
        lab30 = _lines_of(lab30, 4.0)[-1] if lab30 else []
        if len(lab30) != 4:
            d_four.append("第1页底部标签识别到 %d 个（应为4个）" % len(lab30))
        else:
            und = 0
            for s in lab30:
                if "_" in s["raw"] or "＿" in s["raw"]:
                    und += 1
                    continue
                nb = [t for t in p1["spans"] if t is not s and abs(t["cy"] - s["cy"]) <= 6.0
                      and s["x1"] - 2.0 <= t["x0"] <= s["x1"] + 30.0
                      and set(_norm(t["raw"])) <= set("_＿")]
                if nb:
                    und += 1
                    continue
                ln = [st for st in (p1["strokes"] + p1["fills"])
                      if st["w"] >= 25.0 and st["h"] <= 3.5
                      and s["y1"] - 5.0 <= st["y0"] <= s["y1"] + 15.0
                      and st["x0"] <= s["x1"] + 30.0 and st["x1"] >= s["x1"] + 10.0]
                if ln:
                    und += 1
            if not (all(_cdist(s["rgb"], C_DARKBLUE) <= 22 for s in lab30)
                    and all(s["bold"] for s in lab30)
                    and all(_near(s["size"], 8.5, 0.6) for s in lab30) and und == 4):
                d_four.append("第1页底部四格：%.1f磅/加粗%s/下划线%d处"
                              % (lab30[0]["size"], all(s["bold"] for s in lab30), und))
        # ---- 标题条 / 表格 通用判定 ----
        def _bar(no, key, shade_pred):
            """标题条：主标题深蓝加粗15磅 + 深灰10.5磅副标题 + 通栏底纹。

            通栏底纹按“并集宽度”判定，且口径从宽：底纹允许由多段色块拼接
            （部分生成器会把段落底纹拆成按文字 run 的碎块），只要拼接后的
            左右边界与版心基本齐平（误差≤8pt）、视觉上仍是一条通栏色带即可。
            """
            p = _pg(no)
            if not p:
                return None
            mains = [s for s in p["spans"] if _norm(key) in s["norm"] and s["size"] >= 13.0]
            if not mains:
                return {"ok": False, "why": "第%d页未找到「%s」标题条" % (no, key)}
            m = min(mains, key=lambda s: s["y0"])
            b0, b1 = m["y0"] - 6.0, m["y1"] + 6.0
            line = [s for s in p["spans"] if abs(s["cy"] - m["cy"]) <= 6.0 and s is not m]
            subs = [s for s in line if _cdist(s["rgb"], C_DARKGRAY) <= 28]
            cand = [f for f in p["fills"] if _is_tint(f["rgb"]) and _in_band(f, b0, b1, 0.45)]
            uni = _union_span([(f["x0"], f["x1"]) for f in cand])
            ux0, ux1 = uni if uni else (0.0, 0.0)
            full = (bool(uni) and ux0 <= BODY_LEFT + BAR_EDGE_TOL
                    and ux1 >= BODY_RIGHT - BAR_EDGE_TOL)
            hue = any(shade_pred(f["rgb"]) for f in cand)
            ok = (_cdist(m["rgb"], C_DARKBLUE) <= 24 and m["bold"]
                  and _near(m["size"], 15.0, 0.9) and full and hue
                  and bool(subs) and all(_near(s["size"], 10.5, 1.0) for s in subs))
            return {"ok": ok, "why": "第%d页：主标题#%02X%02X%02X/%.1f磅/加粗%s，副标题%d个(%.1f磅)，"
                                     "底纹%d块并集%.1f-%.1fpt(通栏%s/色系%s)"
                    % (no, m["rgb"][0], m["rgb"][1], m["rgb"][2], m["size"], m["bold"],
                       len(subs), subs[0]["size"] if subs else 0.0,
                       len(cand), ux0, ux1, full, hue),
                    "y0": m["y0"], "y1": m["y1"], "span": m}

        def _table_with(no, headers):
            p = _pg(no)
            if not p:
                return None, [], ""
            for t in _tables(p["pg"]):
                inside = [s for s in p["spans"]
                          if s["x0"] >= t["x0"] - 2.0 and s["x1"] <= t["x1"] + 2.0
                          and s["y0"] >= t["y0"] - 2.0 and s["y1"] <= t["y1"] + 2.0]
                txt = "".join(s["norm"] for s in inside)
                if all(_norm(h) in txt for h in headers):
                    return t, inside, txt
            return None, [], ""

        # ---- 26 全册各级标题条：通栏底纹 + 深蓝加粗15磅主标题 + 深灰10.5磅副标题（合并项）----
        # 说明页 / 单元 / 任务 / 项目四类标题条判据完全相同，仅底纹色与所在页不同，
        # 合并为一项；detail 里按类别报出异常条目。
        b37 = [_bar(no, k, _is_lightblue) for no, k in TASK_BARS]
        bar_groups = [
            ("说明页", [_bar(3, "学习结构说明", _is_lightblue)]),
            ("单元", [_bar(no, "单元" + k, _is_lightgreen)
                      for no, k in zip(UNIT_BAR_PAGES, ["一", "二", "三", "四"])]),
            ("任务", b37),
            ("项目", [_bar(no, k, _is_tint) for no, k in
                      zip(PROJ_BAR_PAGES, ["期中项目", "期末项目", "全册学习档案袋"])]),
        ]
        nbar = sum(len(v) for _, v in bar_groups)
        d26 = ["%s%s" % (g, (b["why"] if b else "缺页"))
               for g, v in bar_groups for b in v if not (b and b["ok"])]
        _emit(-1, not d26,
              "全册各级标题条（第3页说明页、四个单元页、十条任务、第15-17页项目）看起来均为"
              "与版心齐平的通栏底纹（左右误差≤8pt，允许由多段色块拼接）"
              " + #1F4E79 加粗 15 磅主标题 + 深灰 10.5 磅副标题",
              "%d 条标题条样式一致" % nbar if not d26
              else "异常 %d/%d 条：%s" % (len(d26), nbar, "；".join(d26[:3])[:240]))

        # ---- 27 单行四格表：第1页底部 + 十条任务标题条下方（合并项）----
        for (no, k), b in zip(TASK_BARS, b37):
            if not b or "y1" not in b:
                d_four.append("第%d页%s无标题条" % (no, k))
                continue
            zone = [s for s in _pg(no)["spans"] if b["y1"] <= s["y0"] <= b["y1"] + 34.0]
            labs = [s for s in zone
                    if re.match(r"^(学科|年级|课时|学习方式)[：:]", s["norm"])]
            if len(labs) < 3:
                d_four.append("第%d页%s下方四格表缺失(%d标签)" % (no, k, len(labs)))
                continue
            lab_bad = [s for s in labs if not (_cdist(s["rgb"], C_DARKBLUE) <= 28 and s["bold"]
                                              and _near(s["size"], 8.5, 0.8))]
            vals = [s for s in zone if s not in labs and abs(s["cy"] - labs[0]["cy"]) <= 6.0]
            val_bad = [s for s in vals if not (_is_dark(s["rgb"], 60) and not s["bold"]
                                              and _near(s["size"], 8.5, 0.8))]
            if lab_bad or val_bad:
                d_four.append("第%d页%s标签异常%d/取值异常%d" % (no, k, len(lab_bad), len(val_bad)))
        _emit(-1, not d_four,
              "单行四格表（第1页底部“学校/班级/姓名/使用时间”与十条任务标题条下方"
              "“学科/年级/课时/学习方式”）标签为深蓝加粗 8.5 磅、取值为黑色常规 8.5 磅、后跟可书写下划线",
              "11 处四格表样式一致" if not d_four else "；".join(d_four[:3]))

        # ---- 28 第3页三个小节标题 ----
        got32 = []
        p3_for_sec = _pg(3)
        for k in P3_SECTIONS:
            hit = ([s for s in p3_for_sec["spans"] if s["norm"].startswith(_norm(k))]
                   if p3_for_sec else [])
            got32.append(min(hit, key=lambda s: s["y0"]) if hit else None)
        if not all(got32):
            _emit(-1, False, "第3页三个小节标题为 #1F4E79 加粗 12 磅并按“一、二、三”顺序排列",
                  "缺失小节标题：%s" % "、".join(k for k, g in zip(P3_SECTIONS, got32) if not g))
        else:
            ys = [s["y0"] for s in got32]
            ok32 = (all(_cdist(s["rgb"], C_DARKBLUE) <= 24 for s in got32)
                    and all(s["bold"] for s in got32)
                    and all(_near(s["size"], 12.0, 0.7) for s in got32)
                    and ys == sorted(ys))
            _emit(-1, ok32, "第3页三个小节标题为 #1F4E79 加粗 12 磅并按“一、二、三”顺序排列",
                  "字号=%s；顺序=%s" % (["%.1f" % s["size"] for s in got32],
                                       "正确" if ys == sorted(ys) else "错乱"))

        # ---- 29 第3页三张表格的行列结构与表头样式（合并项）----
        d30 = []
        t33, in33, _x = _table_with(3, ["步骤"])
        if t33 is None:
            d30.append("未识别到“四步学习链”表")
        else:
            hd = _lines_of(in33, 4.0)[0]
            shade = [f for f in _pg(3)["fills"] if _is_lightblue(f["rgb"]) and f["w"] >= 60.0
                     and _in_band(f, t33["y0"] - 1.0, hd[-1]["y1"] + 3.0, 0.4)]
            if not (t33["rows"] == 5 and t33["cols"] == 3
                    and all(_cdist(s["rgb"], C_DARKBLUE) <= 26 for s in hd)
                    and all(s["bold"] for s in hd)
                    and all(_near(s["size"], 9.0, 0.9) for s in hd) and bool(shade)):
                d30.append("四步学习链表%d行×%d列/表头%.1f磅加粗%s/浅蓝底纹%d块"
                           % (t33["rows"], t33["cols"], hd[0]["size"],
                              all(s["bold"] for s in hd), len(shade)))
        t34, in34, _x = _table_with(3, ["项目主题", "成果形式"])
        if t34 is None:
            d30.append("未识别到“项目地图”表")
        elif not (t34["rows"] == 5 and t34["cols"] == 4):
            d30.append("项目地图表%d行×%d列（应5行×4列）" % (t34["rows"], t34["cols"]))
        t35, in35, _x = _table_with(3, ["我会做"])
        if t35 is None:
            d30.append("未识别到评价方式表")
        elif not (t35["rows"] == 5 and t35["cols"] == 3):
            d30.append("评价方式表%d行×%d列（应5行×3列）" % (t35["rows"], t35["cols"]))
        _emit(-1, not d30,
              "第3页三张表格行列结构与表头样式正确（四步学习链 3列5行且表头浅蓝底纹#1F4E79加粗9磅、"
              "项目地图 4列5行、评价方式 3列5行）",
              "三张表结构与表头正确" if not d30 else "；".join(d30[:3]))
        # ---- 30 小节标题 深蓝或绿色 加粗 11 磅 ----
        seckeys = sorted({_norm(k) for k in SEC_TITLES}, key=len, reverse=True)
        sec_found, sec_bad = [], []
        for p in pages[2:]:
            for s in p["spans"]:
                if len(s["norm"]) > 12:
                    continue
                if not any(s["norm"].startswith(k) for k in seckeys):
                    continue
                sec_found.append(s)
                if not (s["bold"] and _near(s["size"], 11.0, 0.8)
                        and (_cdist(s["rgb"], C_DARKBLUE) <= 26 or _cdist(s["rgb"], C_GREEN) <= 32)):
                    sec_bad.append("第%d页「%s」%.1f磅#%02X%02X%02X"
                                   % (p["no"], s["norm"][:6], s["size"],
                                      s["rgb"][0], s["rgb"][1], s["rgb"][2]))
        ok31 = len(sec_found) >= 10 and not sec_bad
        _emit(-1, ok31, "各页小节标题（真实情境/学习目标/示意图/活动任务单/分层练习/核心概念整理）为深蓝或绿色、加粗、11 磅",
              "识别 %d 处；异常 %d 处%s" % (len(sec_found), len(sec_bad),
                                          ("：" + "；".join(sec_bad[:2])) if sec_bad else ""))

        # ---- 31 六处示意图居中、尺寸正确、无拉伸变形（合并项）----
        def _diag(no, ew, eh):
            p = _pg(no)
            if not p or not p["imgs"]:
                return False, "第%d页无图片" % no
            im = min(p["imgs"], key=lambda z: abs(z["w"] - ew))
            cen = abs((im["x0"] + im["x1"]) / 2.0 - PAGE_CX)
            ar_ok = True
            if im["pw"] > 0 and im["ph"] > 0 and im["h"] > 0:
                a1, a2 = im["w"] / im["h"], im["pw"] / im["ph"]
                ar_ok = abs(a1 - a2) / a2 <= 0.06
            ok = (_near(im["w"], ew, 8.0) and _near(im["h"], eh, 8.0)
                  and cen <= 6.0 and ar_ok)
            return ok, ("第%d页 %.1f×%.1fpt（期望%.0f×%.0f）/居中偏移%.1fpt/变形=%s"
                        % (no, im["w"], im["h"], ew, eh, cen, "无" if ar_ok else "有"))

        r32 = [_diag(no, *DIAGRAMS[no]) for no in (4, 5, 7, 10, 11, 13)]
        ok32b = all(x[0] for x in r32)
        _emit(-1, ok32b,
              "第4/5/7/10/11/13页示意图均居中、宽约 354pt、高度依次约 332/181/239/174/225/181pt，无拉伸变形",
              "六幅示意图尺寸与位置正确" if ok32b
              else "；".join(x[1] for x in r32 if not x[0])[:240])

        # ---- 32 第1页及第3-17页页眉 ----
        d42 = []
        for p in pages:
            if p["no"] == TOC_PIDX + 1:
                continue
            h = _header_of(p["no"])
            if h is None:
                d42.append("第%d页无页眉" % p["no"])
            elif _norm(HEADER_TEXT) not in h["norm"]:
                d42.append("第%d页页眉文字为「%s」" % (p["no"], h["norm"][:16]))
            elif not all(_cdist(c, C_HEADGRAY) <= 24 for c in h["rgbs"]):
                d42.append("第%d页页眉非#646464" % p["no"])
            elif not _near(h["size"], 8.5, 0.6):
                d42.append("第%d页页眉%.1f磅" % (p["no"], h["size"]))
            elif not (_near(h["x1"], 551.0, 4.0) and _near(h["y0"], 36.0, 4.0)):
                d42.append("第%d页页眉位置(%.1f,%.1f)" % (p["no"], h["x1"], h["y0"]))
        ok42 = not d42
        _emit(-1, ok42, "第1页及第3-17页页眉为 #646464、8.5 磅、右边界≈551pt、上边界≈36pt",
              "16 页页眉一致" if ok42 else "；".join(d42[:3]))

        # ---- 33 第1页及第3-17页版心边界（目录页版心归第18项）----
        tops = [min(s["y0"] for s in p["spans"]) for p in pages
                if p["spans"] and p["no"] != TOC_PIDX + 1]
        min_top = min(tops) if tops else 0.0
        ok43 = (m_left is not None and m_right is not None
                and _near(m_left, BODY_LEFT, 3.0) and _near(m_right, BODY_RIGHT, 3.5)
                and min_top >= 34.0)
        _emit(-1, ok43, "第1页及第3-17页版心左边界≈44pt、右边界≈551.5pt，正文首行上边界≈36pt",
              "通栏图元左%.1f/右%.1f；正文页最高文字上边界%.1fpt（限≥34）"
              % (m_left or -1, m_right or -1, min_top))

        # ---- 34 表格细实线全框、单元格文字左对齐且无截断压盖 ----
        d45 = []
        ntab = 0
        for p in pages:
            for t in _tables(p["pg"]):
                ntab += 1
                seg = [st for st in p["strokes"]
                       if st["x0"] >= t["x0"] - 3.0 and st["x1"] <= t["x1"] + 3.0
                       and st["y0"] >= t["y0"] - 3.0 and st["y1"] <= t["y1"] + 3.0]
                if [st for st in seg if st["width"] > 1.2]:
                    d45.append("第%d页表格框线过粗" % p["no"])
                if [st for st in seg if (st["dashes"] or "").strip() not in ("", "[] 0", "[]0")]:
                    d45.append("第%d页表格框线为虚线" % p["no"])
                inside = [s for s in p["spans"]
                          if t["x0"] <= (s["x0"] + s["x1"]) / 2.0 <= t["x1"]
                          and t["y0"] <= s["cy"] <= t["y1"]]
                if [s for s in inside if s["x0"] < t["x0"] - 1.5 or s["x1"] > t["x1"] + 1.5
                        or s["y0"] < t["y0"] - 1.5 or s["y1"] > t["y1"] + 1.5]:
                    d45.append("第%d页表格文字越出边框" % p["no"])
                try:
                    cells = [c for c in (t["obj"].cells or []) if c]
                except Exception:
                    cells = []
                cols = {}
                for c in cells:
                    inn = [s for s in inside
                           if c[0] - 1.0 <= s["x0"] and s["x1"] <= c[2] + 1.0
                           and c[1] - 1.0 <= s["cy"] <= c[3] + 1.0]
                    if not inn:
                        continue
                    cols.setdefault(round(c[0], 1), []).append(
                        min(s["x0"] for s in inn) - c[0])
                worst = max((_spread(v) for v in cols.values() if len(v) >= 2), default=0.0)
                if worst > 8.0:
                    d45.append("第%d页表格列内文字未左对齐(缩进极差%.1fpt)" % (p["no"], worst))
        ok45 = ntab > 0 and not d45
        _emit(-1, ok45, "各页表格为细实线全框，单元格文字左对齐、无截断或压盖",
              "共 %d 个表格，框线与对齐正常" % ntab if ok45
              else "共 %d 个表格；%s" % (ntab, "；".join(sorted(set(d45))[:3])))

        # ---- 35 “□”空心方框（含第3页评价表两列×4行分布）----
        cnt46, box46 = {}, []
        for p in pages:
            n = sum(s["raw"].count("□") for s in p["spans"])
            if n:
                cnt46[p["no"]] = n
            for s in p["spans"]:
                if "□" in s["raw"]:
                    box46.append((p, s))
        total46 = sum(cnt46.values())
        dist_ok = (cnt46 == CHECKBOX_PER_PAGE)
        # 第3页评价方式表内的“□”应为两列、每列 4 行各 1 个（表头行不含）
        p3_box = [s for s in in35 if "□" in s["raw"]]
        p3_nb = sum(s["raw"].count("□") for s in p3_box)
        p3_xcl = len({round(s["x0"] / 40.0) for s in p3_box})
        p3_ycl = len(_lines_of(p3_box, 5.0)) if p3_box else 0
        p3_ok = (p3_nb == 8 and p3_xcl == 2 and p3_ycl == 4)
        sz_bad = [s["size"] for _, s in box46 if not (7.8 <= s["size"] <= 10.6)]
        al_bad = 0
        for p, s in box46:
            nb = [t for t in p["spans"] if t is not s and abs(t["cy"] - s["cy"]) <= 6.0
                  and t["x0"] > s["x1"] - 2.0 and t["x0"] < s["x1"] + 60.0]
            if not nb:
                continue
            t = min(nb, key=lambda z: z["x0"])
            hmin = min(s["y1"] - s["y0"], t["y1"] - t["y0"])
            if _overlap(s["y0"], s["y1"], t["y0"], t["y1"]) < 0.5 * max(1.0, hmin):
                al_bad += 1
        ok46 = (total46 == CHECKBOX_TOTAL and dist_ok and not sz_bad and al_bad == 0
                and p3_ok)
        _emit(-1, ok46, "各处“□”为可逐个勾选的空心方框，字号合规、与所在行文字基线对齐，"
                        "且第3页评价方式表中呈两列×4行分布",
              "共%d个（应%d）；分页分布=%s；字号越界%d处；基线错位%d处；第3页评价表分布%d列×%d行"
              % (total46, CHECKBOX_TOTAL, "一致" if dist_ok else "不一致",
                 len(sz_bad), al_bad, p3_xcl, p3_ycl))
        # ========== 将 results 反解为结构化 dim2_items ==========
        # 逐项按追加顺序对齐 _MAX_DELTAS_ORDER：解析每条 "+N：说明" / "-N：说明"。
        # max_delta ≥ 0 为加分项（命中得分）；max_delta < 0 为扣分项（违规扣分）。
        # 按输出约定：dim2_items 不含 detail 字段（_emit 收集的 details 仅用于本地调试）。
        if len(results) != len(_MAX_DELTAS_ORDER):
            raise RuntimeError(
                f"dim2 项数不一致：results={len(results)} vs order={len(_MAX_DELTAS_ORDER)}"
            )
        dim2_items = []
        total_score = 0
        for raw, max_delta in zip(results, _MAX_DELTAS_ORDER):
            head, _, rule = raw.partition("：")
            try:
                val = int(head.replace("+", "").replace("-", "").strip())
            except ValueError:
                val = 0
            if max_delta >= 0:
                hit = val > 0
                delta = max_delta if hit else 0
            else:
                violated = val > 0
                hit = not violated
                delta = max_delta if violated else 0
            dim2_items.append({"rule": rule, "max_delta": max_delta,
                               "delta": delta, "hit": hit})
            total_score += delta

        result["dim2_items"] = dim2_items
        result["total_score"] = max(0, total_score)
        result["max_score"] = MAX_SCORE_TOTAL
        return result
    finally:
        try:
            doc.close()
        except Exception:
            pass


if __name__ == "__main__":
    if len(sys.argv) >= 2:
        _dir = sys.argv[1]
    else:
        _dir = os.path.dirname(os.path.abspath(__file__))
    _out = evaluate(_dir)
    print(json.dumps(_out, ensure_ascii=False, indent=2))
