#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# pyright: reportUnknownMemberType=false, reportUnknownVariableType=false, reportUnknownArgumentType=false, reportUnknownParameterType=false, reportMissingParameterType=false, reportPrivateUsage=false, reportUnknownLambdaType=false, reportGeneralTypeIssues=false

import os
import re
import sys
import json
import math
import traceback
# PyMuPDF 已从本项目移除（AGPL 与本项目许可证不兼容），运行环境不安装 fitz。
# 改用同目录的 pymupdf_compat：用 pdfplumber + pypdfium2 实现本脚本用到的
# PyMuPDF 子集。verifier 由 core/worker.py 按文件路径加载（其所在目录已被插入
# sys.path），作为 verifiers 包导入时下面这两行会补齐目录，两种方式都成立。
_VERIFIER_DIR = os.path.dirname(os.path.abspath(__file__))
if _VERIFIER_DIR not in sys.path:
    sys.path.insert(0, _VERIFIER_DIR)
try:
    import pymupdf_compat as fitz
except ImportError:
    fitz = None
try:
    import numpy as np  # 位图测量（整体倾角、墨迹边界、底色、坐标轴）依赖
except ImportError:
    np = None
try:
    from PIL import Image  # 旋转-投影锐度法测整体倾角依赖
except ImportError:
    Image = None

# 说明：按接口统一约定，本模块不修改全局 sys.stdout，也不 print 主结果。
# 主结果统一由 evaluate(dir_path) -> dict 返回，入参为"脚本所在目录的路径"，
# 由脚本自己在该目录内定位并打开被评估的 .pdf。本文件仅在 __main__ 下作为
# 本地调试入口，把该 dict 序列化为 JSON 打印。

SCRIPT_ID = "027"
# 维度二加分项 max_delta 之和，用于维度一未通过 / 脚本异常时兜底
MAX_SCORE_TOTAL = 22

# 维度二评分项按"结果追加顺序"的 max_delta 列表：加分项为正，扣分项为负。
# 末端把 results 里的 "±N：说明" 字符串拆解为结构化 dim2_items 时用。顺序须与
# _evaluate_impl 中 results.append(...) 的顺序严格一致，共 22 项；其中正项
# 累加 = MAX_SCORE_TOTAL = 22。本细则无扣分下限，负项只表示该条的扣分幅度。
_MAX_DELTAS_ORDER = [
    5,   # 01 正文行水平（首尾字中心连线与纸张上边夹角 ≤0.3°）
    5,   # 02 文字整体正立（整体倾斜角 ≤0.4°，无斜体/伪斜体/倾斜变形）
    5,   # 03 题间距介于 0.5–1.5 行高
    3,   # 04 表格框线与图表坐标轴与纸张边夹角 ≤0.4°
    3,   # 05 A4 纵向 + 顶部留白 ≤3cm 且各页极差 ≤1cm
    1,   # 06 各页左边界互差 ≤5mm 且均 ≥5mm
    -5,  # 07 既非矢量文本又低于 500dpi
    -5,  # 08 任一题目选项不足 A、B、C、D 四项
    -3,  # 09 图表数据标签 200% 放大下不可辨识
    -3,  # 10 分数/上下标错位、重叠或脱开基线
    -3,  # 11 表格/图表标题缺失
    -3,  # 12 表头错列、框线中断缺失、同列数字对齐不一致
    -3,  # 13 表格或图跨页且后页未重复表头、未标"续表"
    -3,  # 14 各材料段落首行缩进极差 >0.5 字宽
    -3,  # 15 文字非等比缩放（同字号汉字宽高比偏离 1:1 >3%）
    -3,  # 16 最大字号行偏心/未加粗、标题起始不齐、次级标题与正文无区分
    -3,  # 17 水印、"由 XX 转换/生成"字样、软件标识、批注框
    -3,  # 18 版面完全重复页或完全空白页
    -3,  # 19 小于整页的图片底色不一致/带外框线/无内容
    -1,  # 20 同行同字号汉字字心间距极差 >0.3 字宽或字框重叠
    -1,  # 21 同题四选项相邻起始位置之差极差 >6mm
    -1,  # 22 题干末行与首选项间距 >0.5 行高
]
assert sum(d for d in _MAX_DELTAS_ORDER if d > 0) == MAX_SCORE_TOTAL
assert len(_MAX_DELTAS_ORDER) == 22

# ---------------------------------------------------------------- 度量常量
MM = 25.4 / 72.0          # 1pt 折合毫米
DARK = 205                # 灰度暗像素阈值：图表网格线为浅灰，取 150 会漏检
EDGE_DARK = 235           # 量"内容上边缘/左边缘"用更松的阈值，抗锯齿的笔画尖端也算内容
MIN_INK_PX = 8            # "孤立杂点不计"：一行/一列至少这么多暗像素才算内容
MIN_EDGE_PX = 4           # 量边缘时的同口径下限（阈值放松后相应收紧计数）
A4_W, A4_H = 210.0, 297.0  # A4 毫米
CJK_RE = re.compile(r"[一-鿿]")
SENT_PUNCT = "。；"
OPT_LETTERS = "ABCD"
WATERMARK_RE = re.compile(
    r"(由[^\s，。]{1,20}(转换|生成|导出|创建)|试用版|评估版|未注册|demo\s*version"
    r"|watermark|水印|created\s+with|converted\s+by|generated\s+by|powered\s+by)",
    re.IGNORECASE)


# ---------------------------------------------------------------- 基础读取
def _open_pdf(file_path):
    """打开 PDF。仅此一处负责建立 fitz 文档对象，其余度量函数只接受已打开的对象。"""
    if fitz is None:
        raise RuntimeError("缺少依赖 PyMuPDF（import fitz 失败），无法读取 PDF")
    return fitz.open(file_path)


def _doc_lines(doc):
    """全文文本行（rawdict 一次读全，后续所有文本度量都基于它，避免反复解析）。

    每行含 pno / x0,y0,x1,y1 / text / dir / spans（span 内保留 chars）。
    CambriaMath 的 % ~ 等字形自报约 55pt 高的 bbox，会污染行高与行边界，
    故在计算行几何时先剔除该字体的 span（若整行都是它则保留原样）。
    """
    out = []
    for pno in range(doc.page_count):
        for bl in doc[pno].get_text("rawdict")["blocks"]:
            if bl.get("type"):
                continue
            for ln in bl.get("lines", []):
                sp = [s for s in ln["spans"] if s.get("font") != "CambriaMath"] or ln["spans"]
                if not sp:
                    continue
                txt = "".join(c["c"] for s in sp for c in s.get("chars", []))
                if not txt.strip():
                    continue
                out.append({
                    "pno": pno,
                    "x0": min(s["bbox"][0] for s in sp),
                    "y0": min(s["bbox"][1] for s in sp),
                    "x1": max(s["bbox"][2] for s in sp),
                    "y1": max(s["bbox"][3] for s in sp),
                    "text": txt,
                    "dir": ln.get("dir", (1.0, 0.0)),
                    "spans": sp,
                })
    out.sort(key=lambda r: (r["pno"], round(r["y0"], 1), r["x0"]))
    return out


def _line_chars(ln):
    """行内字符扁平表：[(字, bbox, 字号, span)]，按书写顺序。"""
    res = []
    for s in ln["spans"]:
        for c in s.get("chars", []):
            res.append((c["c"], c["bbox"], s.get("size", 0.0), s))
    return res


def _visual_rows(lines, y_tol=1.5):
    """把 PDF 文本层拆开的同一视觉行重新合并。

    兼容后端可能把同一基线上的年份、正文片段或左右相邻文本返回成多条 line。
    直接把这些碎片当成段首、标题或完整文字行会产生假缩进和假标题。
    """
    rows = []
    for ln in sorted(lines, key=lambda r: (r["pno"], r["y0"], r["x0"])):
        cy = (ln["y0"] + ln["y1"]) / 2.0
        row = next(
            (r for r in reversed(rows)
             if r["pno"] == ln["pno"] and abs(r["_cy"] - cy) <= y_tol),
            None,
        )
        if row is None:
            row = dict(ln)
            row["spans"] = list(ln["spans"])
            row["_parts"] = [(ln["x0"], ln["text"])]
            row["_cy"] = cy
            rows.append(row)
            continue
        row["x0"] = min(row["x0"], ln["x0"])
        row["y0"] = min(row["y0"], ln["y0"])
        row["x1"] = max(row["x1"], ln["x1"])
        row["y1"] = max(row["y1"], ln["y1"])
        row["spans"].extend(ln["spans"])
        row["_parts"].append((ln["x0"], ln["text"]))
    for row in rows:
        row["spans"].sort(key=lambda s: (s.get("bbox") or (0,))[0])
        row["text"] = "".join(text for _x, text in sorted(row.pop("_parts")))
        row.pop("_cy", None)
    return rows


def _body_size(lines):
    """正文字号：取承载汉字最多的那个字号（四舍五入到 0.1pt）。"""
    tally = {}
    for ln in lines:
        for s in ln["spans"]:
            n = sum(1 for c in s.get("chars", []) if CJK_RE.match(c["c"]))
            if n:
                tally[round(s.get("size", 0.0), 1)] = tally.get(round(s.get("size", 0.0), 1), 0) + n
    return max(tally, key=lambda k: tally[k]) if tally else 0.0


def _body_left(lines):
    """版口左边界：取全文行左边缘的最小值簇（最小的 5% 分位），躲开孤立左突行。"""
    xs = sorted(ln["x0"] for ln in lines)
    if not xs:
        return 0.0
    return xs[max(0, int(len(xs) * 0.02))]
def _page_pitch(lines, pno, page_h):
    """该页行高：同页各行 y0 排序后相邻差的中位数，只取 10–25pt 的正常行距。
    页内文本不足时退回 0.35 倍字号经验值由调用方处理（返回 None）。"""
    ys = sorted({round(ln["y0"], 2) for ln in lines if ln["pno"] == pno})
    d = [b - a for a, b in zip(ys, ys[1:]) if 10.0 <= (b - a) <= 25.0]
    if not d:
        return None
    d.sort()
    return d[len(d) // 2]


# ---------------------------------------------------------------- 位图度量
def _render_gray(page, dpi=150, clip=None):
    """把页面（或页面上的一块 clip）渲染成灰度 numpy 数组。"""
    if np is None:
        return None
    px = page.get_pixmap(dpi=dpi, clip=clip, colorspace=fitz.csGRAY, annots=False)
    if px.width <= 0 or px.height <= 0:
        return None
    return np.frombuffer(px.samples, dtype=np.uint8).reshape(px.height, px.width)


def _ink_edges(page, dpi=200):
    """该页墨迹的上边缘与左边缘（pt，相对纸张左上角）。

    用渲染像素而不是文本/图像 bbox：整页位图的 bbox 等于整页，说明不了顶部留白；
    文本 bbox 的上边是字体 ascent 线而非笔画实际起点，会把留白少报。像素法对矢量页
    与位图页口径一致。一行/一列至少 MIN_EDGE_PX 个暗像素才算内容，对应"孤立杂点不计"。
    """
    a = _render_gray(page, dpi=dpi)
    if a is None:
        return None, None
    dark = a < EDGE_DARK
    rows = dark.sum(axis=1)
    cols = dark.sum(axis=0)
    ri = np.nonzero(rows >= MIN_EDGE_PX)[0]
    ci = np.nonzero(cols >= MIN_EDGE_PX)[0]
    k = 72.0 / dpi
    top = float(ri[0]) * k if ri.size else None
    left = float(ci[0]) * k if ci.size else None
    return top, left


def _projection_angle(arr, axis, lo=-3.0, hi=3.0, coarse=0.10, fine=0.01):
    """旋转-投影锐度法测整体倾角，返回版面倾角（度，正值表示内容顺时针歪）。

    axis=0 用行投影，测的是"基线是否水平"；axis=1 用列投影，测的是"字的竖向
    是否正立"。刚性旋转时两者一致；两者背离说明页面被做了水平错切——基线平了、
    字还是斜的，必须两个都测才判得出"文字整体正立"。
    """
    if Image is None or np is None or arr is None or arr.size == 0:
        return None
    im = Image.fromarray(arr)

    def score(th):
        r = np.asarray(im.rotate(float(th), resample=Image.BICUBIC, fillcolor=255))
        # 先在 uint8 上截断，再只对一维投影使用 float32，避免每次候选角度
        # 同时分配两张页面大小的 float32 数组而触发内存峰值。
        w = np.where(r < DARK, DARK - r, 0).astype(np.uint8)
        p = w.sum(axis=1) if axis == 0 else w.sum(axis=0)
        return float(np.var(np.diff(p.astype(np.float32))))

    grid = [lo + coarse * i for i in range(int(round((hi - lo) / coarse)) + 1)]
    best = max(grid, key=score)
    grid = [best - coarse + fine * i for i in range(int(round(2 * coarse / fine)) + 1)]
    best = max(grid, key=score)
    return -float(best)


def _text_region(page, lines, pno):
    """本页"正文文字区"的裁剪框：把图表区域排除掉。

    投影锐度法会被折线图的数据折线带偏（实测能读出 +2.5° 的假值），所以测倾角前
    必须把图片块和大面积矢量图形块从测量区里剔除。有文本层时用文本行 bbox 求并集；
    无文本层（纯扫描件）时退回整页减去图片以外无从判断，只能用整页。
    """
    r = page.rect
    imgs = []
    for bl in page.get_text("dict")["blocks"]:
        if bl.get("type") == 1:
            imgs.append(fitz.Rect(bl["bbox"]))
    own = [ln for ln in lines if ln["pno"] == pno]
    if own:
        x0 = min(ln["x0"] for ln in own)
        x1 = max(ln["x1"] for ln in own)
        bands = []
        for ln in sorted(own, key=lambda v: v["y0"]):
            box = fitz.Rect(ln["x0"], ln["y0"], ln["x1"], ln["y1"])
            if any((im & box).get_area() > 0.3 * box.get_area() for im in imgs):
                continue
            bands.append((ln["y0"], ln["y1"]))
        if bands:
            return fitz.Rect(x0, bands[0][0] - 1, x1, bands[-1][1] + 1)
    # 无文本层：取整页，但若整页里图片占比过半则取图片之外最大的横条
    if imgs and max(im.get_area() for im in imgs) < 0.5 * r.get_area():
        return r
    return r


def _runs(mask, min_len):
    """布尔序列里长度 ≥min_len 的连续 True 段，返回 [(起, 止)]（闭区间）。"""
    res = []
    s = None
    for i, v in enumerate(mask):
        if v and s is None:
            s = i
        elif not v and s is not None:
            if i - s >= min_len:
                res.append((s, i - 1))
            s = None
    if s is not None and len(mask) - s >= min_len:
        res.append((s, len(mask) - 1))
    return res


def _fit_edge(arr, fixed, q0, q1, horiz, half=6, bands=6, min_n=40):
    """灰度加权亚像素质心 + 分段拟合，测一条直线边（表格框线 / 坐标轴）的倾角。

    horiz=True 沿 x 扫、每列求纵向质心；False 沿 y 扫、每行求横向质心。质心用
    w=clip(DARK-灰度,0,None) 加权，比二值化稳。拟合前把序列切成 bands 段取段均值，
    并要求段均值单调——真实刚性倾斜是单调的，抖动的是污染（JPEG 块、柱形边、
    数据折线），不能当倾斜。返回 (角度°, 样本数, 是否单调)。
    """
    if np is None or arr is None:
        return None, 0, False
    h, w = arr.shape
    xs, cs = [], []
    c = float(fixed)
    for t in range(int(q0), int(q1) + 1):
        if horiz:
            if not (0 <= t < w):
                continue
            lo = int(max(0, c - half)); hi = int(min(h - 1, c + half))
            seg = arr[lo:hi + 1, t]
        else:
            if not (0 <= t < h):
                continue
            lo = int(max(0, c - half)); hi = int(min(w - 1, c + half))
            seg = arr[t, lo:hi + 1]
        wt = np.clip(DARK - seg.astype(np.float32), 0, None)
        if wt.sum() < 12:
            continue
        cc = lo + float((wt * np.arange(len(wt))).sum() / wt.sum())
        if abs(cc - c) > half * 0.9:
            continue
        c = 0.5 * c + 0.5 * cc          # 跟踪漂移：倾斜线不会停在同一像素列/行
        xs.append(float(t)); cs.append(cc)
    if len(xs) < min_n:
        return None, len(xs), False
    idx = np.array_split(np.arange(len(xs)), bands)
    bx = [float(np.mean([xs[i] for i in g])) for g in idx if len(g)]
    bc = [float(np.mean([cs[i] for i in g])) for g in idx if len(g)]
    if len(bx) < 3:
        return None, len(xs), False
    slope = float(np.polyfit(bx, bc, 1)[0])
    dif = [b - a for a, b in zip(bc, bc[1:])]
    mono = all(d > 0 for d in dif) or all(d < 0 for d in dif)
    ang = math.degrees(math.atan(slope))
    return ang, len(xs), mono


def _native_images(doc):
    """取出全部嵌入图片的原生像素（不重采样）。

    角度必须在原生像素上量：把 220dpi 的图表 JPEG 渲染到 600dpi 再拟合质心，
    双三次插值会把亚像素梯度抹平，实测会把 0.39° 少报成 0.25°。
    返回 [(pno, xref, 灰度数组, 放置矩形, 折算 dpi)]。
    """
    res = []
    if np is None or Image is None:
        return res
    import io
    seen = set()
    for pno in range(doc.page_count):
        page = doc[pno]
        for info in page.get_images(full=True):
            xref = info[0]
            try:
                rects = page.get_image_rects(xref)
            except Exception:
                rects = []
            if not rects:
                continue
            key = (pno, xref)
            if key in seen:
                continue
            seen.add(key)
            try:
                d = doc.extract_image(xref)
                im = Image.open(io.BytesIO(d["image"])).convert("L")
            except Exception:
                continue
            a = np.asarray(im, dtype=np.uint8)
            r = rects[0]
            dpi = (a.shape[1] / (r.width / 72.0)) if r.width > 1 else 0.0
            res.append((pno, xref, a, r, dpi))
    return res


def _is_thin_line(a, pos, q0, q1, horiz, slope=0.0, half=8, need=0.90, maxw=4):
    """沿一条候选线取 24 个采样点，看穿过它的暗像素连续段是否始终很细。

    坐标轴与框线是细线，任何采样点上穿过它的暗段都只有几个像素；柱形图的柱边
    穿过去是整根实心柱，暗段很宽。这一条把"柱边"从"坐标轴"里剔出去。
    slope 是该线的斜率（每前进 1px 的横移量），采样窗口必须跟着它走：0.6° 的线
    在 850px 上要漂 9px，窗口不跟就会在后半段脱靶，把细线误判成"不细"。
    """
    if np is None:
        return False
    n = ok = 0
    step = max(1, int((q1 - q0) / 24))
    for t in range(int(q0), int(q1) + 1, step):
        c = pos + slope * (t - q0)
        if horiz:
            if not (0 <= t < a.shape[1]):
                continue
            lo = int(max(0, c - half)); hi = int(min(a.shape[0] - 1, c + half))
            seg = a[lo:hi + 1, t]
        else:
            if not (0 <= t < a.shape[0]):
                continue
            lo = int(max(0, c - half)); hi = int(min(a.shape[1] - 1, c + half))
            seg = a[t, lo:hi + 1]
        if seg.size < 3:
            continue
        n += 1
        rs = _runs(seg < DARK, 1)
        if rs and min(e - s + 1 for s, e in rs) <= maxw:
            ok += 1
    return n >= 8 and ok / float(n) >= need


def _line_candidates(a, horiz, min_frac=0.30, span=1.6, step=0.08):
    """错切累加找候选直线（简化 Hough）：返回 [(暗像素数, 起点位置, 斜率角°)]。

    直接按行/列计数只能找到水平线——0.6° 的线在 900px 上跨 9 行，每行只落到
    1/9 的墨，任何阈值都筛不出来。这里把图按一组试探角错切后再计数，倾斜线在
    对应角度上会重新聚成一条满格的峰。同一条线会在相邻角度上重复命中，按线在
    幅面中点的位置归并，只留计数最高的那个角。
    """
    if np is None:
        return []
    d = (a < DARK) if horiz else (a < DARK).T
    ext, other = d.shape
    need = other * min_frac
    edge = max(3, int(ext * 0.04))
    xs = np.arange(other)
    best = {}
    n_ang = int(round(2.0 * span / step))
    for i in range(n_ang + 1):
        ang = -span + step * i
        k = math.tan(math.radians(ang))
        acc = np.zeros(ext, dtype=np.int32)
        off = np.rint(k * xs).astype(int)
        for x in range(other):
            o = int(off[x]); col = d[:, x]
            if o > 0:
                acc[:ext - o] += col[o:]
            elif o < 0:
                acc[-o:] += col[:ext + o]
            else:
                acc += col
        for r0, r1 in _runs(acc >= need, 1):
            p = (r0 + r1) // 2
            if p < edge or p > ext - 1 - edge:
                continue          # 贴图幅边的暗线是截图边界或扫描黑边，不是坐标轴
            v = int(acc[r0:r1 + 1].max())
            key = int(round((p + k * other / 2.0) / 3.0))
            if key not in best or v > best[key][0]:
                best[key] = (v, p, ang)
    res, cen = [], []
    for (v, p, ang) in sorted(best.values(), reverse=True):
        c = p + math.tan(math.radians(ang)) * other / 2.0
        if any(abs(c - c2) <= 4.0 for c2 in cen):
            continue              # 同一条线会在相邻试探角上重复命中，按幅面中点归并
        cen.append(c)
        res.append((v, p, ang))
    return res


def _bitmap_edges(a, min_frac=0.30, min_len=0.55):
    """在一张位图里找长直边（图表外框、坐标轴、表格框线），逐条量倾角。

    候选由 _line_candidates 用错切累加给出（能找到倾斜线），再用 _fit_edge 亚像素
    拟合并要求分段单调、有效样本覆盖 ≥min_len 倍幅面。最后一道是"确认"：要么这条
    线穿过去始终很细（独立的坐标轴、外框），要么它和同向的另外两条以上平行（表格
    的成排框线、图表的网格线）——刚性倾斜必然是一族平行线，而数据折线、柱形边、
    JPEG 块噪不会既长、又单调、又互相平行。
    返回 [(kind, 角度°, 样本数)]，kind ∈ {"h","v"}。
    """
    out = []
    if np is None or a is None or a.ndim != 2:
        return out
    h, w = a.shape
    if h < 60 or w < 60:
        return out
    cand = []
    for horiz in (True, False):
        ext = w if horiz else h
        q0, q1 = int(ext * 0.06), int(ext * 0.94)
        for (_v, pos, ang0) in _line_candidates(a, horiz, min_frac=min_frac)[:12]:
            k = math.tan(math.radians(ang0))
            start = pos + k * q0            # 候选位置是 x=0 处的截距，挪到扫描起点
            ang, n, mono = _fit_edge(a, start, q0, q1, horiz)
            if ang is None or not mono or n < min_len * ext:
                continue
            cand.append((horiz, ang, n,
                         _is_thin_line(a, start, q0, q1, horiz, slope=k)))
    for i, (horiz, ang, n, thin) in enumerate(cand):
        par = sum(1 for j, t in enumerate(cand)
                  if j != i and t[0] == horiz and abs(t[1] - ang) <= 0.10)
        if not (thin or par >= 2):
            continue
        out.append(("h" if horiz else "v", ang if horiz else -ang, n))
        # 竖线：质心随 y 右移即顺时针，取反与横线对齐口径
    return out


# ---------------------------------------------------------------- 题目结构
def _option_marks(lines):
    """选项标记的字符级定位：字符属 A/B/C/D、其后紧跟句点、且它是词首。

    "词首"判据是它与前一字符的水平间隙 >2pt，用来排除正文里夹带的字母。
    返回 [(pno, y0, x0, letter, line)]。
    """
    res = []
    for ln in lines:
        chs = _line_chars(ln)
        for j, (c, bb, _sz, _sp) in enumerate(chs):
            if c not in OPT_LETTERS:
                continue
            if j + 1 >= len(chs) or chs[j + 1][0] not in ".．、":
                continue
            if j > 0 and bb[0] - chs[j - 1][1][2] <= 2.0:
                continue
            res.append((ln["pno"], bb[1], bb[0], c, ln))
    return res


def _questions(lines):
    """按题号切分题目。

    题号判据说死为：行首三位数字后紧跟句点或空白、行左边缘不超过版口左边界 +45pt、
    且号码落在 101–130 区间（本题本的题号域）。不按"点后是否为非数字"过滤——
    材料里的"2010.12"这类不在行首且左边缘不满足，已被前两条排除。
    """
    if not lines:
        return []
    bl = _body_left(lines)
    stems = []
    for i, ln in enumerate(lines):
        m = re.match(r"^(\d{3})[.．\s]", ln["text"])
        if m and ln["x0"] <= bl + 45 and 101 <= int(m.group(1)) <= 130:
            stems.append((i, int(m.group(1))))
    marks = _option_marks(lines)
    qs = []
    for k, (i, no) in enumerate(stems):
        j = stems[k + 1][0] if k + 1 < len(stems) else len(lines)
        seg = lines[i:j]
        key0 = (lines[i]["pno"], lines[i]["y0"])
        key1 = (lines[j]["pno"], lines[j]["y0"]) if j < len(lines) else (10 ** 6, 0.0)
        own = [m for m in marks if key0 < (m[0], m[1]) < key1]
        opt_lines, letters = [], []
        for m in own:
            letters.append(m[3])
            if m[4] not in opt_lines:
                opt_lines.append(m[4])
        # 题干 = 首个选项行之前的那些行；不能用"非选项行"取末行，选项之后还挂着
        # 下一组材料时会把题干末行取到选项下面去，间距就成了荒谬的几十行高。
        if opt_lines:
            fk = min((ln["pno"], ln["y0"]) for ln in opt_lines)
            stem_lines = [ln for ln in seg if (ln["pno"], ln["y0"]) < fk]
        else:
            stem_lines = list(seg)
        qs.append({
            "no": no,
            "pno": lines[i]["pno"],
            "stem_first": lines[i],
            "stem_last": stem_lines[-1] if stem_lines else lines[i],
            "opt_lines": opt_lines,
            "marks": own,
            "letters": letters,
        })
    return qs


def _tables(doc):
    """逐页取表格（PyMuPDF 的 find_tables）。老版本或识别失败时返回空表，
    对应的扣分项即无从触发——细则的扣分项宁可漏判也不误判。"""
    res = []
    for pno in range(doc.page_count):
        try:
            tf = doc[pno].find_tables()
            for t in tf.tables:
                res.append((pno, t))
        except Exception:
            continue
    return res
def _is_caption_str(t):
    """表格自身首行的文字是否构成标题：不含句末标点、长度 ≤50、且不是纯数字表头。"""
    t = (t or "").strip().replace("\n", "")
    if not t or len(t) > 50:
        return False
    if any(p in t for p in SENT_PUNCT):
        return False
    if re.fullmatch(r"[\d\s.,%－\-—~～]+", t):
        return False
    return bool(CJK_RE.search(t))


def _is_caption(ln, page_w, ref_c=None):
    """标题行判据（说死，不依赖指令与原稿）：不含句末标点「。；」、长度 ≤50 字、
    且墨迹水平中心距纸张水平中心不超过 10mm（居中排布）——本题本的表题图题一律居中。
    图表标题也允许对其所属图表居中（ref_c 传该图的水平中心），图窄于版口时
    题居图心与题居页心同为正确排法。表格自身首行的文字另行单独判定，不走这里。"""
    t = ln["text"].strip()
    if not t or len(t) > 50:
        return False
    if any(p in t for p in SENT_PUNCT):
        return False
    if re.match(r"^\d{3}[.．\s]", t) or re.match(r"^[A-D][.．]", t):
        return False
    if re.match(r"^(表|图)\s*\d", t):
        return True
    c = (ln["x0"] + ln["x1"]) / 2.0
    if abs(c - page_w / 2.0) * MM <= 10.0:
        return True
    return ref_c is not None and abs(c - ref_c) * MM <= 10.0


def _content_edges(page, lines, pno):
    """该页"最上方内容的上边缘"与"最靠左内容的左边缘"（pt）。

    口径：取页面上内容对象（文本行、矢量图形、非整页图片）的外框上边与左边——
    这就是尺子量到的版口位置。整页位图页没有有效的内容外框（其外框等于整页），
    此时退回渲染墨迹的实测边缘。孤立杂点由 _ink_edges 的像素计数下限排除。
    """
    tops, lefts = [], []
    has_txt = False
    for ln in lines:
        if ln["pno"] == pno:
            has_txt = True
            tops.append(ln["y0"]); lefts.append(ln["x0"])
    try:
        for d in page.get_drawings():
            r = fitz.Rect(d["rect"])
            if r.height <= 0.0 and r.width <= 0.0:
                continue
            tops.append(r.y0); lefts.append(r.x0)
    except Exception:
        pass
    big = 0.9 * page.rect.get_area()
    if has_txt:                            # 无文字层的页，图片就是整页内容本身，
        for info in page.get_images(full=True):   # 它的外框说明不了版口，只能量墨迹
            try:
                for r in page.get_image_rects(info[0]):
                    if r.get_area() < big:
                        tops.append(r.y0); lefts.append(r.x0)
            except Exception:
                pass
    if tops and lefts:
        return min(tops), min(lefts)
    return _ink_edges(page)


# ---------------------------------------------------------------- 维度一
def _check_dim1(doc, file_path):
    """维度一：文件是否正常。返回 (是否通过, 未通过原因)；通过时原因为空串。"""
    if os.path.splitext(file_path)[1].lower() != ".pdf":
        return False, "被评估文件不是 .pdf"
    if getattr(doc, "needs_pass", False) or getattr(doc, "is_encrypted", False):
        return False, "文件带打开口令或处于加密未解状态"
    try:
        perm = int(doc.permissions)
        if not (perm & fitz.PDF_PERM_PRINT):
            return False, "文件禁止打印"
    except Exception:
        pass
    if doc.page_count <= 0:
        return False, "文件无可显示页面"
    for pno in range(doc.page_count):
        page = doc[pno]
        try:
            a = _render_gray(page, dpi=48)
        except Exception as exc:
            return False, "第 %d 页无法渲染：%s" % (pno + 1, exc)
        if a is None:
            continue
        if float(a.mean()) < 60.0:
            return False, "第 %d 页为黑页" % (pno + 1)
        txt = page.get_text("text")
        ink = float((a < DARK).mean())
        has_obj = bool(page.get_images(full=True)) or bool(page.get_drawings())
        if ink < 5e-4 and not txt.strip() and not has_obj:
            return False, "第 %d 页为纯空白页" % (pno + 1)
        if txt and (txt.count("�") + txt.count("\x00")) > max(3, len(txt) * 0.02):
            return False, "第 %d 页文本大量不可映射，疑为乱码页" % (pno + 1)
    return True, ""


# ---------------------------------------------------------------- 对外接口
def evaluate(dir_path):
    """入参为目录路径，在该目录内定位并打开被评估的 .pdf，返回结构化结果字典。"""
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
        cands = sorted(
            f for f in os.listdir(dir_path)
            if f.lower().endswith(".pdf") and not f.startswith("~$")
        )
        if not cands:
            raise FileNotFoundError("目录内未找到 .pdf 文件：%s" % dir_path)
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
    """维度一门槛 + 维度二 22 条逐条判定。results 里每条形如 "±N：条文｜实测"，
    N 为 0 表示该加分项未得分 / 该扣分项未触发。追加顺序须与 _MAX_DELTAS_ORDER 一致。"""
    doc = _open_pdf(file_path)
    try:
        ok, why = _check_dim1(doc, file_path)
        result["dim1_pass"] = bool(ok)
        # dim1_reason 只在维度一未通过时给出原因，通过时留空
        result["dim1_reason"] = "" if ok else why
        if not ok:
            result["dim2_items"] = []
            result["total_score"] = 0
            result["max_score"] = MAX_SCORE_TOTAL
            return result
        return _score_dim2(doc, result)
    finally:
        try:
            doc.close()
        except Exception:
            pass


def _shrink(arr, maxw=760):
    """把待测数组等比缩到不超过 maxw 宽，纯粹为了控制旋转搜索的耗时。
    倾角是角量，等比缩放不改变它；但不要用放大——放大会抹平亚像素梯度、少报倾角。"""
    if arr is None or Image is None or arr.shape[1] <= maxw:
        return arr
    k = maxw / float(arr.shape[1])
    im = Image.fromarray(arr).resize(
        (maxw, max(1, int(arr.shape[0] * k))), Image.LANCZOS)
    return np.asarray(im, dtype=np.uint8)


def _cell_box(lines, pno, rect):
    """落在给定单元格矩形内的字符的墨迹外框与文本，用于表格对齐类判定。"""
    x0 = y0 = 1e9
    x1 = y1 = -1e9
    buf = []
    for ln in lines:
        if ln["pno"] != pno:
            continue
        for c, bb, _sz, _sp in _line_chars(ln):
            cx, cy = (bb[0] + bb[2]) / 2.0, (bb[1] + bb[3]) / 2.0
            if rect[0] <= cx <= rect[2] and rect[1] <= cy <= rect[3]:
                if not c.strip():
                    continue
                x0 = min(x0, bb[0]); y0 = min(y0, bb[1])
                x1 = max(x1, bb[2]); y1 = max(y1, bb[3])
                buf.append(c)
    if not buf:
        return None
    return (x0, y0, x1, y1, "".join(buf))


def _seg_cover(segs, a, b):
    """线段集合 segs=[(起,止)] 对区间 [a,b] 的覆盖比例，用于判"框线中断或缺失"。"""
    if b <= a:
        return 0.0
    hit = []
    for s, e in segs:
        s, e = max(s, a), min(e, b)
        if e > s:
            hit.append((s, e))
    hit.sort()
    tot, cur = 0.0, None
    for s, e in hit:
        if cur is None or s > cur[1]:
            if cur:
                tot += cur[1] - cur[0]
            cur = [s, e]
        else:
            cur[1] = max(cur[1], e)
    if cur:
        tot += cur[1] - cur[0]
    return tot / (b - a)


def _vector_rules(page):
    """页面里的矢量直线，按近水平/近竖直分两组返回。
    近轴判据为偏离横向或纵向不超过 5 度——真正斜排的装饰线不是"框线"，不参与判定。"""
    hs, vs = [], []
    try:
        drawings = page.get_drawings()
    except Exception:
        return hs, vs
    for d in drawings:
        for it in d.get("items", []):
            pts = []
            if it[0] == "l":
                pts = [(it[1], it[2])]
            elif it[0] == "re":
                r = it[1]
                pts = [(fitz.Point(r.x0, r.y0), fitz.Point(r.x1, r.y0)),
                       (fitz.Point(r.x0, r.y1), fitz.Point(r.x1, r.y1)),
                       (fitz.Point(r.x0, r.y0), fitz.Point(r.x0, r.y1)),
                       (fitz.Point(r.x1, r.y0), fitz.Point(r.x1, r.y1))]
            for p0, p1 in pts:
                dx, dy = p1.x - p0.x, p1.y - p0.y
                L = math.hypot(dx, dy)
                if L < 20.0:
                    continue
                a = math.degrees(math.atan2(dy, dx)) % 180.0
                if min(a, 180.0 - a) <= 5.0:
                    hs.append({"a": min(a, 180.0 - a), "x0": min(p0.x, p1.x),
                               "x1": max(p0.x, p1.x), "y": (p0.y + p1.y) / 2.0})
                elif abs(a - 90.0) <= 5.0:
                    vs.append({"a": abs(a - 90.0), "y0": min(p0.y, p1.y),
                               "y1": max(p0.y, p1.y), "x": (p0.x + p1.x) / 2.0})
    return hs, vs


# ---------------------------------------------------------------- 维度二
def _score_dim2(doc, result):
    results = []
    npg = doc.page_count
    pages = [doc[i] for i in range(npg)]
    lines = _doc_lines(doc)
    page_w = pages[0].rect.width
    body = _body_size(lines)
    bleft = _body_left(lines)
    has_text = len(lines) >= npg          # 平均每页不足一行文本即视为无文字层
    imgs = _native_images(doc)
    full_bitmap_pages = [p for p in range(npg)
                         if not [ln for ln in lines if ln["pno"] == p]]
    # 图表 = 排在有文字的版面里、幅面小于整页、且长宽均不小于 20mm 的图片。无文字层
    # 页面上的位图不是"图表"，它承载的就是这一页的全部正文，按图表去判标题、底色、
    # 标签分辨率都是无的放矢；行内嵌的十几 pt 小图是分数、符号等字符件，同样不是图表。
    sub_imgs = [t for t in imgs
                if t[0] not in full_bitmap_pages
                and t[3].get_area() < 0.9 * pages[t[0]].rect.get_area()
                and min(t[3].width, t[3].height) >= 20.0 / MM]
    qs = _questions(lines) if has_text else []
    pitch = {p: _page_pitch(lines, p, pages[p].rect.height) for p in range(npg)}
    vr = [_vector_rules(pg) for pg in pages]
    proj = {}                             # 页级双向投影倾角缓存 {pno: (行投影, 列投影)}

    def page_proj(p):
        if p not in proj:
            reg = _text_region(pages[p], lines, p)
            arr = _shrink(_render_gray(pages[p], dpi=120, clip=reg))
            proj[p] = (_projection_angle(arr, 0, coarse=0.25, fine=0.02),
                       _projection_angle(arr, 1, coarse=0.25, fine=0.02))
        return proj[p]

    # 01 +5 正文行水平：有文字层时直接读行方向向量；纯扫描件用行投影锐度法测基线。
    a01, miss01 = [], []
    for p in range(npg):
        own = [ln for ln in lines if ln["pno"] == p and len(ln["text"].strip()) >= 6]
        if own:
            v = max(abs(math.degrees(math.atan2(ln["dir"][1], ln["dir"][0]))) for ln in own)
        else:
            r, _c = page_proj(p)
            v = abs(r) if r is not None else None
        if v is None:
            miss01.append(p + 1)
        else:
            a01.append(v)
    ok01 = (not miss01) and len(a01) == npg and max(a01) <= 0.3
    results.append("%s：每一页任取一整行正文文字，其首尾字中心连线与纸张上边的夹角不超过 0.3 度"
                   "｜逐页最大夹角 %s"
                   % ("+5" if ok01 else "+0",
                      ("%.3f°" % max(a01)) if a01 else "无法测量"))

    # 02 +5 文字整体正立：行投影测基线、列投影测字竖向，两者都要 ≤0.4°。
    # 二者背离即说明页面做过水平错切——基线摆平了、字还是斜的，只测基线判不出来。
    a02, sh02, it02 = [], 0.0, 0
    for ln in lines:
        for s in ln["spans"]:
            fn = s.get("font") or ""
            if (int(s.get("flags", 0)) & 2) or re.search(r"(italic|oblique|斜)", fn, re.I):
                it02 += len(s.get("chars", []))
    for p in range(npg):
        r, c = page_proj(p)
        for v in (r, c):
            if v is not None:
                a02.append(abs(v))
        if r is not None and c is not None:
            sh02 = max(sh02, abs(r - c))
    ok02 = bool(a02) and max(a02) <= 0.4 and sh02 <= 0.4 and it02 == 0
    results.append("%s：正文文字整体正立，每一页正文文字区的整体倾斜角不超过 0.4 度，"
                   "且全文无斜体、伪斜体或倾斜变形的字形｜整体倾角最大 %s，行/列投影最大背离 %.3f°，斜体字符 %d 个"
                   % ("+5" if ok02 else "+0",
                      ("%.3f°" % max(a02)) if a02 else "无法测量", sh02, it02))
    # 03 +5 题间距：上题末选项文字下边缘到下题题号文字上边缘，折算成该页行高的倍数。
    bynum = {q["no"]: q for q in qs}

    def _gap_ratio(a, b):
        if a["pno"] != b["pno"]:
            return None
        pt = pitch.get(a["pno"])
        if not pt:
            return None
        bot = (max(ln["y1"] for ln in a["opt_lines"]) if a["opt_lines"]
               else a["stem_last"]["y1"])
        return (b["stem_first"]["y0"] - bot) / pt

    pairs = []
    for x, y in ((113, 114), (118, 120)):
        if x in bynum and y in bynum:
            r = _gap_ratio(bynum[x], bynum[y])
            if r is not None:
                pairs.append(("%d-%d" % (x, y), r))
    if len(pairs) < 2:                     # 抽检两处被分置两页时，改取全文最先出现的两对同页相邻题
        pairs = []
        for a, b in zip(qs, qs[1:]):
            r = _gap_ratio(a, b)
            if r is not None:
                pairs.append(("%d-%d" % (a["no"], b["no"]), r))
            if len(pairs) == 2:
                break
    ok03 = len(pairs) == 2 and all(0.5 <= r <= 1.5 for _n, r in pairs)
    results.append("%s：上一道题最后一个选项的文字下边缘，与下一道题题号的文字上边缘之间的距离，"
                   "介于该页行高的 0.5 至 1.5 倍，抽检 113 题与 114 题之间、118 题与 120 题之间｜%s"
                   % ("+5" if ok03 else "+0",
                      ("，".join("%s 为 %.2f 行高" % (n, r) for n, r in pairs)
                       if pairs else "无文字层，无法定位题号")))

    # 04 +3 框线与坐标轴：矢量线按坐标精确算；位图内的线用原生像素质心分段拟合，
    # 并要求分段单调（真实刚性倾斜是单调的），排除 JPEG 块噪声与柱形边的假值；
    # 整页位图退回页级双向投影——该页内容是原稿的刚性栅格，框线与页面同倾。
    d04 = []
    for p in range(npg):
        hs, vs = vr[p]
        d04 += [x["a"] for x in hs] + [x["a"] for x in vs]
    for (p, _xr, arr, _rect, _dpi) in sub_imgs:
        # 位图图表只认拟合出来的细长直线；一条都没找到就是这张图里没有可测的坐标轴，
        # 不能退回整图投影——投影会被折线图的数据折线带偏，实测能读出 0.88° 的假值。
        for _k, a, _n in _bitmap_edges(arr, min_frac=0.30):
            d04.append(abs(a))
    for p in full_bitmap_pages:
        # 整页位图页：页内容是原稿的刚性栅格，框线与页面同倾，故以页级双向投影
        # 代表其框线倾角（已由实测验证：A 卷竖框线中位 +0.656/+0.709°，与列投影
        # 读到的错切 0.69/0.90° 吻合到 0.02°）。
        r, c = page_proj(p)
        for v in (r, c):
            if v is not None:
                d04.append(abs(v))
    ok04 = (not d04) or max(d04) <= 0.4
    results.append("%s：全部表格的横竖框线，以及全部图表的横纵坐标轴，与纸张边的夹角不超过 0.4 度"
                   "｜实测最大夹角 %s（受检线 %d 条）"
                   % ("+3" if ok04 else "+0",
                      ("%.3f°" % max(d04)) if d04 else "无表格与图表", len(d04)))

    # 05 +3 A4 纵向 + 顶部留白：留白用渲染像素量，整页位图的对象 bbox 等于整页，说明不了留白。
    sz_ok, tops = True, []
    for p in range(npg):
        r = pages[p].rect
        w_mm, h_mm = r.width * MM, r.height * MM
        if abs(w_mm - A4_W) > 2.0 or abs(h_mm - A4_H) > 2.0:
            sz_ok = False
        t, _l = _content_edges(pages[p], lines, p)
        if t is not None:
            tops.append(t * MM / 10.0)     # cm
    top_ok = (len(tops) == npg and max(tops) <= 3.0
              and (max(tops) - min(tops)) <= 1.0)
    ok05 = sz_ok and top_ok
    results.append("%s：全部页面纸张尺寸为 A4 210×297mm、方向纵向，且每页最上方内容的上边缘距纸张上边"
                   "不超过 3cm，各页该距离的最大差不超过 1cm"
                   "｜纸张 %.1f×%.1fmm，顶部留白 %s"
                   % ("+3" if ok05 else "+0", pages[0].rect.width * MM,
                      pages[0].rect.height * MM,
                      ("%.2f–%.2fcm，极差 %.2fcm" % (min(tops), max(tops), max(tops) - min(tops))
                       if tops else "无法测量")))

    # 06 +1 各页左边界：同样用渲染像素，一列至少 MIN_INK_PX 个暗像素才算内容（孤立杂点不计）。
    lefts = []
    for p in range(npg):
        _t, l = _content_edges(pages[p], lines, p)
        if l is not None:
            lefts.append(l * MM)
    ok06 = (len(lefts) == npg and (max(lefts) - min(lefts)) <= 5.0
            and min(lefts) >= 5.0)
    results.append("%s：各页最靠左的文字或框线距纸张左边的距离，各页之间相差不超过 5mm，"
                   "且均不小于 5mm｜%s"
                   % ("+1" if ok06 else "+0",
                      ("%.2f–%.2fmm，极差 %.2fmm" % (min(lefts), max(lefts), max(lefts) - min(lefts))
                       if lefts else "无法测量")))
    # 07 −5 既非矢量文本又低于 500dpi：两个条件同时成立才扣。
    dpis = [t[4] for t in imgs if t[4] > 0]
    lowdpi = bool(dpis) and max(dpis) < 500.0
    hit07 = (not has_text) and lowdpi
    results.append("%s：正文文字既非可正常选中复制的矢量文本，且按页面实际尺寸折算的成像分辨率低于 500dpi"
                   "｜文本行 %d 行，图像折算分辨率 %s"
                   % ("−5" if hit07 else "−0", len(lines),
                      ("最高 %.0fdpi" % max(dpis)) if dpis else "无嵌入图像"))

    # 08 −5 选项不足四项：无文字层时无从核验，按不触发处理（扣分项宁可漏判不误判）。
    bad08 = [q["no"] for q in qs if set(q["letters"]) != set(OPT_LETTERS)]
    hit08 = bool(bad08)
    results.append("%s：任一题目的选项数量不足 A、B、C、D 四项｜共识别 %d 道题，%s"
                   % ("−5" if hit08 else "−0", len(qs),
                      ("缺项题号 " + "、".join(str(n) for n in bad08[:8])) if bad08 else "均齐备"))

    # 09 −3 图表数据标签可辨识度：以原生像素折算的成像分辨率为准。
    # 200% 放大后仍要看得清，对应原生不低于 150dpi；低于此值标签必然糊连成马赛克。
    bad09 = [(t[0] + 1, t[4]) for t in sub_imgs if 0 < t[4] < 150.0]
    hit09 = bool(bad09)
    results.append("%s：长宽均不小于 20mm 的图表，其内部数据标签在 200%% 放大下存在笔画糊连、"
                   "马赛克块或无法辨识的字符"
                   "｜图表 %d 张，最低折算 %s"
                   % ("−3" if hit09 else "−0", len(sub_imgs),
                      ("%.0fdpi（第 %d 页）" % (bad09[0][1], bad09[0][0])) if bad09
                      else ("%.0fdpi" % min([t[4] for t in sub_imgs] or [0]))))

    # 10 −3 分数与上下标：分数以短横线为锚，检查上下两部分居中对齐、不重叠；
    # 上下标检查其外框是否整体脱出所在行的行框。
    bad10 = []
    for p in range(npg):
        hs, _vs = vr[p]
        for h in hs:
            wdt = h["x1"] - h["x0"]
            if not (4.0 <= wdt <= 40.0):
                continue
            up, dn = [], []
            for ln in lines:
                if ln["pno"] != p:
                    continue
                for c, bb, sz, _sp in _line_chars(ln):
                    if not c.strip() or bb[2] < h["x0"] - 3 or bb[0] > h["x1"] + 3:
                        continue
                    if 0 < h["y"] - bb[3] <= 14.0:
                        up.append(bb)
                    elif 0 < bb[1] - h["y"] <= 14.0:
                        dn.append(bb)
            if not up or not dn:
                continue
            cu = (min(b[0] for b in up) + max(b[2] for b in up)) / 2.0
            cd = (min(b[0] for b in dn) + max(b[2] for b in dn)) / 2.0
            szu = max(b[3] - b[1] for b in up)
            if abs(cu - cd) > 0.25 * max(szu, 1.0):
                bad10.append("第 %d 页分数上下错位 %.2fpt" % (p + 1, abs(cu - cd)))
            if max(b[3] for b in up) > min(b[1] for b in dn):
                bad10.append("第 %d 页分数上下重叠" % (p + 1))
    for ln in lines:
        chs = [t for t in _line_chars(ln) if t[0].strip()]
        if len(chs) < 3:
            continue
        big = max(t[2] for t in chs)
        for c, bb, sz, _sp in chs:
            if sz >= big * 0.8:
                continue
            if bb[1] > ln["y1"] - 0.05 or bb[3] < ln["y0"] + 0.05:
                bad10.append("第 %d 页上下标脱开基线（'%s'）" % (ln["pno"] + 1, c))
                break
    hit10 = bool(bad10)
    results.append("%s：分数、上下标等叠置结构的上下两部分左右错位、相互重叠，或与所在行基线脱开｜%s"
                   % ("−3" if hit10 else "−0",
                      "；".join(bad10[:3]) if bad10 else "未见异常"))

    # 11 −3 表格与图表标题：表格标题可在自身首行，或自上边缘向上追溯五行文字之内
    # （不足则顺延至上一页页末）；图表标题在其上、下各五行文字之内任一侧即可。
    tabs = _tables(doc)
    bad11 = []
    for pno, t in tabs:
        bb = t.bbox
        own = False
        try:
            rows = t.extract()
            if rows and any(_is_caption_str(x) for x in (rows[0] or []) if x):
                own = True
        except Exception:
            pass
        if own:
            continue
        above = [ln for ln in lines if (ln["pno"], ln["y1"]) < (pno, bb[1] + 0.5)]
        if any(_is_caption(ln, pages[ln["pno"]].rect.width) for ln in above[-5:]):
            continue
        bad11.append("第 %d 页表格上方五行内无标题" % (pno + 1))
    for (p, _xr, _arr, rect, _dpi) in sub_imgs:
        above = [ln for ln in lines if (ln["pno"], ln["y1"]) < (p, rect.y0 + 0.5)]
        below = [ln for ln in lines if (ln["pno"], ln["y0"]) > (p, rect.y1 - 0.5)]
        cand = above[-5:] + below[:5]
        rc = (rect.x0 + rect.x1) / 2.0
        if not any(_is_caption(ln, pages[ln["pno"]].rect.width, rc) for ln in cand):
            bad11.append("第 %d 页图表上下五行内均无标题" % (p + 1))
    hit11 = bool(bad11)
    results.append("%s：表格的标题既不在其自身首行，也不在自表格上边缘按阅读顺序向上追溯的五行文字之内；"
                   "或长宽均不小于 20mm 的图表自其上、下边缘各向外追溯"
                   "五行文字，两侧均无该图表的标题"
                   "｜表格 %d 个、图表 %d 张，%s"
                   % ("−3" if hit11 else "−0", len(tabs), len(sub_imgs),
                      "；".join(bad11[:3]) if bad11 else "标题均可追溯"))

    # 12 −3 表格内部：表头与数据列错位、框线中断缺失、同列数字对齐方式不一致。
    # 三项都取宽松口径：无框表不判"框线缺失"，只有本就成网格的表才校验外框覆盖率。
    bad12 = []
    for pno, t in tabs:
        bb = t.bbox
        hs, vs = vr[pno]
        thin_h = [(x["x0"], x["x1"]) for x in hs if bb[1] - 2 <= x["y"] <= bb[3] + 2]
        thin_v = [(x["y0"], x["y1"]) for x in vs if bb[0] - 2 <= x["x"] <= bb[2] + 2]
        if len(thin_h) >= 2 and len(thin_v) >= 2:
            for x in (bb[0], bb[2]):
                segs = [(v["y0"], v["y1"]) for v in vs if abs(v["x"] - x) <= 2.0]
                if _seg_cover(segs, bb[1], bb[3]) < 0.85:
                    bad12.append("第 %d 页表格竖向外框覆盖不足" % (pno + 1))
                    break
            for y in (bb[1], bb[3]):
                segs = [(h["x0"], h["x1"]) for h in hs if abs(h["y"] - y) <= 2.0]
                if _seg_cover(segs, bb[0], bb[2]) < 0.85:
                    bad12.append("第 %d 页表格横向外框覆盖不足" % (pno + 1))
                    break
        try:
            cells = [c for c in t.cells if c]
        except Exception:
            cells = []
        cols = {}
        for c in cells:
            cols.setdefault(round(c[0], 1), []).append(c)
        for _k, cl in cols.items():
            cl.sort(key=lambda r: r[1])
            boxes = [_cell_box(lines, pno, r) for r in cl]
            boxes = [b for b in boxes if b]
            if len(boxes) < 3:
                continue
            wid = max(1.0, cl[0][2] - cl[0][0])
            hdr, dat = boxes[0], boxes[1:]
            nums = [b for b in dat if re.fullmatch(r"[\d.,%－\-—~～]+", b[4] or "")]
            if len(nums) >= 3:
                sl = max(b[0] for b in nums) - min(b[0] for b in nums)
                sr = max(b[2] for b in nums) - min(b[2] for b in nums)
                sc = (max((b[0] + b[2]) / 2 for b in nums)
                      - min((b[0] + b[2]) / 2 for b in nums))
                if min(sl, sr, sc) > 0.5 * max(body, 1.0):
                    bad12.append("第 %d 页表格同列数字对齐不一致 %.2fpt" % (pno + 1, min(sl, sr, sc)))
            ch = (hdr[0] + hdr[2]) / 2.0
            cc = sum((b[0] + b[2]) / 2.0 for b in dat) / len(dat)
            if abs(ch - cc) > 0.5 * wid:
                bad12.append("第 %d 页表头与数据列错位 %.2fpt" % (pno + 1, abs(ch - cc)))
    hit12 = bool(bad12)
    results.append("%s：表格存在表头单元格与其数据列左右错位、框线中断或缺失、同列数字的对齐方式不一致｜%s"
                   % ("−3" if hit12 else "−0",
                      "；".join(sorted(set(bad12))[:3]) if bad12 else "未见异常"))

    # 13 −3 表格或图跨页：前页表格贴到版口底部、后页顶部又起一表，才算被拆开；
    # 此时后页首行须与前表首行一致（重复表头），或其上三行内出现"续表"。
    bad13 = []
    by_page = {}
    for pno, t in tabs:
        by_page.setdefault(pno, []).append(t)
    for pno in sorted(by_page):
        ph = pages[pno].rect.height
        for t in by_page[pno]:
            if t.bbox[3] < ph - 60.0:
                continue
            nxt = by_page.get(pno + 1) or []
            nxt = [x for x in nxt if x.bbox[1] < 140.0]
            if not nxt:
                continue
            n = nxt[0]
            try:
                r0 = [x for x in (t.extract()[0] or []) if x]
                r1 = [x for x in (n.extract()[0] or []) if x]
            except Exception:
                r0, r1 = [], []
            same = bool(r0) and r0 == r1
            head = [ln for ln in lines
                    if ln["pno"] == pno + 1 and ln["y1"] <= n.bbox[1] + 0.5][-3:]
            cont = any("续表" in ln["text"] for ln in head)
            if not (same or cont):
                bad13.append("第 %d–%d 页表格跨页未重复表头且未标续表" % (pno + 1, pno + 2))
    hit13 = bool(bad13)
    results.append("%s：表格或图被拆到两页，且后一页未重复表头、未标注\"续表\"｜%s"
                   % ("−3" if hit13 else "−0",
                      "；".join(bad13[:3]) if bad13 else "未见被拆开的表格与图"))

    # 14 −3 材料段落首行缩进：只把“根据……材料/资料，回答……”分组标题后的
    # 第一条视觉行认作材料段首。PDF 后端常把同一行拆成年份、正文、数字等碎片，
    # 仅凭“相对版口缩进”会把表格单元格、选项续行和正文残行误认成段首。
    tab_boxes = [(pno, fitz.Rect(t.bbox)) for pno, t in tabs]
    firsts = []
    visual_rows = _visual_rows(lines)
    material_heading = re.compile(
        r"^[一二三四五六七八九十百]+、.*(?:材料|资料).*回答.*题")
    for i, ln in enumerate(visual_rows[:-1]):
        if not material_heading.search(re.sub(r"\s+", "", ln["text"])):
            continue
        first = visual_rows[i + 1]
        if first["pno"] != ln["pno"] or len(first["text"].strip()) < 8:
            continue
        # 图表型材料紧接的是居中的图表标题，不属于“材料段落首行”；
        # 表格内部首行同样不参与段落缩进比较。
        if _is_caption(first, pages[first["pno"]].rect.width):
            continue
        center = fitz.Point((first["x0"] + first["x1"]) / 2,
                            (first["y0"] + first["y1"]) / 2)
        if any(pno == first["pno"] and rect.contains(center)
               for pno, rect in tab_boxes):
            continue
        firsts.append(first["x0"])
    spread14 = (max(firsts) - min(firsts)) / body if len(firsts) >= 2 and body > 0 else 0.0
    hit14 = spread14 > 0.5
    results.append("%s：各材料段落首行文字左边缘之间的最大差超过 0.5 个汉字宽｜首行 %d 处，极差 %.2f 字宽"
                   % ("−3" if hit14 else "−0", len(firsts), spread14))

    # 15 −3 非等比缩放：汉字的字面宽应等于字号（全角方框），宽高比偏离 1:1 超 3% 即变形。
    # 逐字统计，允许零星异形字，超过 2% 的字符越界才判为"文字被非等比缩放"。
    tot15 = bad15 = 0
    worst15 = 0.0
    for ln in lines:
        for c, bb, sz, _sp in _line_chars(ln):
            if sz <= 0 or not CJK_RE.match(c):
                continue
            tot15 += 1
            r = (bb[2] - bb[0]) / sz
            worst15 = max(worst15, abs(r - 1.0))
            if abs(r - 1.0) > 0.03:
                bad15 += 1
    hit15 = tot15 >= 50 and bad15 > 0.02 * tot15
    results.append("%s：文字被非等比缩放，同一字号下汉字的宽高比偏离 1:1 超过 3%%，呈拉长或压扁"
                   "｜汉字 %d 个，越界 %d 个，最大偏离 %.2f%%"
                   % ("−3" if hit15 else "−0", tot15, bad15, worst15 * 100.0))

    # 16 −3 标题层级：最大字号行的居中度与字重、同字号标题行起始位置、次级标题与正文的区分度。
    bad16 = []
    if lines:
        smax = max((s.get("size", 0.0) for ln in lines for s in ln["spans"]), default=0.0)
        tl = [ln for ln in lines
              if any(abs(s.get("size", 0.0) - smax) < 0.05 for s in ln["spans"])]
        for ln in tl:
            cx = (ln["x0"] + ln["x1"]) / 2.0
            off = abs(cx - pages[ln["pno"]].rect.width / 2.0) * MM
            if off > 3.0:
                bad16.append("第 %d 页最大字号行偏心 %.2fmm" % (ln["pno"] + 1, off))
            bold = any((int(s.get("flags", 0)) & 20) == 20
                       or re.search(r"(bold|black|heavy|黑|粗)", s.get("font") or "", re.I)
                       for s in ln["spans"])
            if not bold:
                bad16.append("第 %d 页最大字号行未加粗" % (ln["pno"] + 1))
        by_size = {}
        for ln in lines:
            k = round(max(s.get("size", 0.0) for s in ln["spans"]), 1)
            # 居中排布的标题行不参与"起始位置"比较：居中时左边缘随字数变化，
            # 比 x0 等于在比标题长短，这一条针对的是左对齐的同级标题不齐。
            if _is_caption(ln, pages[ln["pno"]].rect.width):
                continue
            # 表格内的单元格文字也不参与：单元格文字的 x0 由列位置决定，跨半页的
            # 列间距会被读成"标题起始位置相差 150mm"，那不是标题层级问题。
            if any(pno == ln["pno"] and r.contains(fitz.Point((ln["x0"] + ln["x1"]) / 2,
                                                              (ln["y0"] + ln["y1"]) / 2))
                   for pno, r in tab_boxes):
                continue
            # 标题要独占一行：右侧还压着同一水平带上的别的文字块，说明它是表头、
            # 图例或分栏标签，不是标题行。
            if any(o is not ln and o["pno"] == ln["pno"] and o["x0"] > ln["x1"] + 2
                   and min(o["y1"], ln["y1"]) - max(o["y0"], ln["y0"]) > 0.4 * (ln["y1"] - ln["y0"])
                   for o in lines):
                continue
            if k > body + 0.05 and len(ln["text"].strip()) <= 30:
                by_size.setdefault(k, []).append(ln["x0"])
        for k, xs in by_size.items():
            if len(xs) >= 2 and (max(xs) - min(xs)) * MM > 3.0:
                bad16.append("%.1fpt 标题行起始位置相差 %.2fmm" % (k, (max(xs) - min(xs)) * MM))
        body_bold = None
        for ln in lines:
            for s in ln["spans"]:
                if abs(s.get("size", 0.0) - body) < 0.05 and len(s.get("chars", [])) > 12:
                    body_bold = (int(s.get("flags", 0)) & 20) == 20
                    break
            if body_bold is not None:
                break
        # 次级标题：短行、无句末标点、不是题号行也不是选项行，上方留出 ≥1.3 倍行高的
        # 空档（标题总要与上文拉开距离），且下方紧接一整行正文。只靠"短行"会把段末
        # 残行和表题都当标题，误判成"次级标题与正文无区分"。
        subs = []
        heading_rows = visual_rows
        for i, ln in enumerate(heading_rows):
            t = ln["text"].strip()
            if not t or len(t) > 30 or any(p in t for p in SENT_PUNCT):
                continue
            if re.match(r"^\d{3}[.．\s]", t) or re.match(r"^[A-D][.．]", t):
                continue
            k = round(max(s.get("size", 0.0) for s in ln["spans"]), 1)
            if abs(k - smax) < 0.05:
                continue
            nx = heading_rows[i + 1] if i + 1 < len(heading_rows) else None
            if nx is None or nx["pno"] != ln["pno"] or len(nx["text"].strip()) < 15:
                continue
            pv = heading_rows[i - 1] if i > 0 else None
            pt = pitch.get(ln["pno"]) or (body * 1.5 if body else 0.0)
            if pv is None or pv["pno"] != ln["pno"] or not pt:
                continue
            # 上方空档要落在 0.55–2.5 倍行高之间：标题总与上文拉开一点距离，但
            # 空档大到十几行说明中间隔着图表——那是图题，不是次级标题。
            if not (0.55 * pt <= (ln["y0"] - pv["y1"]) <= 2.5 * pt):
                continue
            subs.append(ln)
        for ln in subs:
            k = round(max(s.get("size", 0.0) for s in ln["spans"]), 1)
            bd = any((int(s.get("flags", 0)) & 20) == 20
                     or re.search(r"(bold|black|heavy|黑|粗)", s.get("font") or "", re.I)
                     for s in ln["spans"])
            if abs(k - body) < 0.05 and body_bold is not None and bd == body_bold:
                bad16.append("第 %d 页次级标题与正文字号字重完全相同" % (ln["pno"] + 1))
                break
    hit16 = bool(bad16)
    results.append("%s：字号最大的文字行其墨迹水平中心与纸张水平中心相差超过 3mm，或该行未加粗，"
                   "或同一字号的标题行之间起始位置相差超过 3mm，"
                   "或次级标题与正文的字号及字重完全相同而无任何区分｜%s"
                   % ("−3" if hit16 else "−0",
                      "；".join(list(dict.fromkeys(bad16))[:3]) if bad16 else "未见异常"))

    # 17 −3 版面无关的附加元素：水印/转换生成字样/软件标识按正则命中，批注框按注释对象计数。
    bad17 = []
    for ln in lines:
        m = WATERMARK_RE.search(ln["text"])
        if m:
            bad17.append("第 %d 页出现\"%s\"" % (ln["pno"] + 1, m.group(0)[:20]))
    nann = 0
    for p in range(npg):
        try:
            nann += sum(1 for _a in pages[p].annots())
        except Exception:
            pass
    if nann:
        bad17.append("批注等注释对象 %d 个" % nann)
    hit17 = bool(bad17)
    results.append("%s：页面存在与版面无关的附加元素，含水印、\"由 XX 转换\"或\"由 XX 生成\"字样、"
                   "软件标识、批注框｜%s"
                   % ("−3" if hit17 else "−0",
                      "；".join(sorted(set(bad17))[:3]) if bad17 else "未见附加元素"))
    # 18 −3 重复页与空白页：重复以"页内文字 + 图形对象数 + 墨迹分布指纹"比对；
    # 空白以渲染墨迹占比判定，避免把只有图形没有文字的页误判成空白。
    bad18, sig = [], {}
    for p in range(npg):
        txt = re.sub(r"\s+", "", pages[p].get_text("text"))
        a = _render_gray(pages[p], dpi=48)
        ink = float((a < DARK).mean()) if a is not None else 0.0
        nobj = len(pages[p].get_drawings()) + len(pages[p].get_images(full=True))
        if ink < 5e-4 and not txt and nobj == 0:
            bad18.append("第 %d 页为空白页" % (p + 1))
            continue
        key = (txt, nobj)
        if txt and key in sig:
            bad18.append("第 %d 页与第 %d 页版面完全重复" % (sig[key] + 1, p + 1))
        else:
            sig[key] = p
    hit18 = bool(bad18)
    results.append("%s：存在版面内容完全重复的页，或完全空白页｜%s"
                   % ("−3" if hit18 else "−0",
                      "；".join(bad18[:3]) if bad18 else "%d 页均有独立内容" % npg))

    # 19 −3 小于整页的图片：底色须与页面底色一致、不带外框线、须承载内容。
    # 底色取图内灰度的 90 分位（纸白占多数）与页面同口径值相比，容差 25 级——
    # 扫描件本身略灰不算问题，套一层灰底才算。外框线要求四边同时压暗且图内底色
    # 也异常：图表自身的绘图区外框是内容，不能算"给图片描了个边"。
    bad19 = []
    for (p, _xr, arr, rect, _dpi) in sub_imgs:
        base = _render_gray(pages[p], dpi=72)
        pbg = float(np.percentile(base, 90)) if base is not None else 255.0
        h, w = arr.shape
        if h < 16 or w < 16:
            continue
        ibg = float(np.percentile(arr, 90))
        tinted = abs(ibg - pbg) > 25.0
        if tinted:
            bad19.append("第 %d 页图片底色与页面底色相差 %.0f 级" % (p + 1, abs(ibg - pbg)))
        sides = [arr[0:2, :], arr[-2:, :], arr[:, 0:2], arr[:, -2:]]
        if tinted and all(float(np.median(s)) < pbg - 60.0 for s in sides):
            bad19.append("第 %d 页图片带外框线" % (p + 1))
        if float((arr < DARK).mean()) < 1e-3:
            bad19.append("第 %d 页图片不承载内容" % (p + 1))
    hit19 = bool(bad19)
    results.append("%s：幅面小于整页且长宽均不小于 20mm 的图片，其矩形范围内的底色与页面底色不一致，"
                   "或带有外框线，或不承载任何内容｜图片 %d 张，%s"
                   % ("−3" if hit19 else "−0", len(sub_imgs),
                      "；".join(sorted(set(bad19))[:3]) if bad19 else "未见异常"))
    # 20 −1 字距与字框：同行内相邻的同字号汉字，字心间距极差 >0.3 字宽即视为字距不匀；
    # 字框重叠留 0.15 字宽容差——字符前进宽度上的记账式微量交叠不是可见的压字。
    worst20, ov20 = 0.0, 0
    for ln in visual_rows:
        chs = sorted(_line_chars(ln), key=lambda item: (
            (item[1][0] + item[1][2]) / 2.0, item[1][1]))
        # 只取"在字符序列上真正相邻"的同字号汉字对：中间夹了数字、括号、标点的两个
        # 汉字不是相邻字，按相邻处理会把一整段跳跃当成字距，读出上千分之一的假极差。
        for sz in {round(t[2], 1) for t in chs if CJK_RE.match(t[0])}:
            if sz <= 0:
                continue
            runs20, cur = [], []
            for (c0, b0, s0, _p0), (c1, b1, s1, _p1) in zip(chs, chs[1:]):
                ok = (CJK_RE.match(c0) and CJK_RE.match(c1)
                      and abs(round(s0, 1) - sz) <= 0.05
                      and abs(round(s1, 1) - sz) <= 0.05)
                # 用字符左边缘的推进距离，不用字形 bbox 中心；不同汉字的左右
                # 侧承会改变 bbox 宽度，但不代表排版字距发生变化。
                gap = (b1[0] - b0[0]) if ok else -1.0
                # 间距超过 1.4 字宽的不是相邻字：表格里同一基线上的两个单元格会被
                # 合成一"行"，跨单元格取差会把列间空白当成字距。遇到就断开成新串。
                if not ok or gap <= 0 or gap > 1.4 * sz:
                    if cur:
                        runs20.append(cur)
                    cur = []
                    continue
                cur.append(gap)
                if b0[2] - b1[0] > 0.15 * sz:
                    ov20 += 1
            if cur:
                runs20.append(cur)
            for d in runs20:
                if len(d) >= 5:
                    # 单个字符缺失、侧承异常或文本层碎片会制造一个孤立大间距；
                    # 去掉两端各一个样本后仍超阈值，才认为整行字距确实不均。
                    ordered = sorted(d)
                    robust_spread = ordered[-2] - ordered[1]
                    worst20 = max(worst20, robust_spread / sz)
    hit20 = worst20 > 0.3 or ov20 > 0
    results.append("%s：同一行内相邻的同字号汉字，其字心间距的最大差超过 0.3 个汉字宽，"
                   "或字符外框相互重叠｜字心间距极差最大 %.3f 字宽，重叠 %d 处"
                   % ("−1" if hit20 else "−0", worst20, ov20))

    # 21 −1 选项起始位置：先按 y 把四个标记聚成行，只取"行内"相邻标记的水平间距，
    # 跨行取差会把换行当成间距，得出无意义的巨值。
    worst21, who21 = 0.0, None
    for q in qs:
        rows = {}
        for (_p, y, x, _c, _ln) in q["marks"]:
            k = next((r for r in rows if abs(r - y) < 4.0), None)
            rows.setdefault(k if k is not None else y, []).append(x)
        gaps = []
        for _k, xs in rows.items():
            xs.sort()
            gaps += [b - a for a, b in zip(xs, xs[1:])]
        if len(gaps) >= 2:
            r = (max(gaps) - min(gaps)) * MM
            if r > worst21:
                worst21, who21 = r, q["no"]
    hit21 = worst21 > 6.0
    results.append("%s：同一道题目的四个选项，相邻两个选项起始位置之差的极差超过 6mm｜最大极差 %.2fmm%s"
                   % ("−1" if hit21 else "−0", worst21,
                      ("（第 %d 题）" % who21) if who21 else ""))

    # 22 −1 题干末行与首选项：两者之间不应留出半行以上的空档。
    # 空档里夹着图片（分数、算式等行内小图）时，那段纵向空间是被图占着的，
    # 不算"留白"，故以"图的上边缘"作为下一处内容的起点。
    img_boxes = {}
    for (p, _xr, _a, r, _d) in imgs:
        img_boxes.setdefault(p, []).append(r)
    worst22, who22 = 0.0, None
    for q in qs:
        if not q["opt_lines"]:
            continue
        pt = pitch.get(q["pno"])
        if not pt:
            continue
        first = min(q["opt_lines"], key=lambda ln: (ln["pno"], ln["y0"]))
        if first["pno"] != q["stem_last"]["pno"]:
            continue                       # 题干与首选项分置两页时不构成"行间空档"
        top = first["y0"]
        for r in img_boxes.get(q["pno"], []):
            if r.y1 > q["stem_last"]["y1"] + 0.5 and r.y0 < first["y0"] - 0.5:
                top = min(top, max(r.y0, q["stem_last"]["y1"]))
        r22 = (top - q["stem_last"]["y1"]) / pt
        if r22 > worst22:
            worst22, who22 = r22, q["no"]
    hit22 = worst22 > 0.5
    results.append("%s：任一题目的题干末行文字下边缘，与其后第一处内容上边缘之间的距离，"
                   "超过该页行高的 0.5 倍｜最大 %.2f 行高%s"
                   % ("−1" if hit22 else "−0", worst22,
                      ("（第 %d 题）" % who22) if who22 else ""))

    # ---- 汇总：把 results 里的 "±N：条文｜实测" 拆成结构化 dim2_items ----
    if len(results) != len(_MAX_DELTAS_ORDER):
        raise RuntimeError("dim2 项数不一致：results=%d vs order=%d"
                           % (len(results), len(_MAX_DELTAS_ORDER)))
    dim2_items = []
    total_score = 0
    for raw, max_delta in zip(results, _MAX_DELTAS_ORDER):
        head, _, body_txt = raw.partition("：")
        rule, _, _detail = body_txt.partition("｜")   # 明细只用于本地调试，不进结果
        head = head.strip()
        sign = -1 if head[:1] in ("-", "−") else 1
        digits = re.sub(r"[^0-9]", "", head)
        mag = int(digits) if digits else 0
        hit = mag > 0
        if hit and mag != abs(max_delta):
            raise RuntimeError("dim2 分值与 max_delta 不符：%s vs %d" % (head, max_delta))
        if hit and (sign > 0) != (max_delta > 0):
            raise RuntimeError("dim2 正负号与 max_delta 不符：%s vs %d" % (head, max_delta))
        dim2_items.append({
            "rule": rule,
            "max_delta": max_delta,
            "delta": max_delta if hit else 0,
            "hit": hit,
        })
        total_score += (max_delta if hit else 0)
    result["dim2_items"] = dim2_items
    result["total_score"] = max(0, total_score)   # 完成度分母固定为加分项满分，得分不低于 0
    result["max_score"] = MAX_SCORE_TOTAL
    return result


if __name__ == "__main__":
    _dir = sys.argv[1] if len(sys.argv) >= 2 else os.path.dirname(os.path.abspath(__file__))
    try:
        print(json.dumps(evaluate(_dir), ensure_ascii=False, indent=2))
    except Exception:
        traceback.print_exc()
