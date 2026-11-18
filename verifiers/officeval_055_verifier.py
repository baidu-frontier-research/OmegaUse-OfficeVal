# -*- coding: utf-8 -*-
"""officeval_055_verifier —— “网站故障处理流程图（可编辑 PPT）”交付件完成度评测。

评测对象：单页 .pptx，横向 16:9，自左至右 5 个流程步骤组（编号 / 浅蓝圆形图标底 /
深色圆角矩形阶段标签 / 说明文本框），左上角图例框，下部橙色虚线情绪折线与情绪标签，
底部主题文本。

实现要点：
1) 直接解析 ppt/slides/slide1.xml，不依赖 python-pptx；填充/线条颜色均读取 a:alpha，
   alpha=0 的白色不计为“白色填充”。
2) 组合（p:grpSp）内的子形状按 a:xfrm 的 off/ext/chOff/chExt 做坐标变换，换算为幻灯片
   绝对坐标，因此组合内的图标零件同样可被量测。
3) 细则允许 5 个步骤图标“由形状组合或小尺寸图片实现”。因此图标造型类条目先按形状
   组合逐件核验；若图标区域内存在具备图标尺寸的位图（p:pic），则该造型条目按“图片
   实现”记为达成——位图内部构造无法静态拆解，这是细则允许范围内的必然放宽；徽标、
   对勾等叠放件若本身为位图，其内部（白色感叹号 / 对勾连续性）同样不再拆解，但位置、
   尺寸、叠放关系仍照常量测。
4) “文字完整显示未被裁切”一项仅做结构性核验（文本框尺寸为正、位于页内、未被其后
   绘制的不透明形状整体覆盖），不模拟字体排版换行。

返回 evaluate(dir_path) -> dict，格式与 officeval_008_verifier 一致。
"""

from __future__ import annotations

import json
import math
import os
import sys
import zipfile
from xml.etree import ElementTree as ET

SCRIPT_ID = "055"
# 维度二所有"加分项"max_delta 之和；也是维度一未通过时回填的 max_score
# （脚本没评上、即 status 为 error 时 max_score 为 0）。
MAX_SCORE_TOTAL = 90
SUPPORTED_EXTS = (".pptx",)

A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
P = "{http://schemas.openxmlformats.org/presentationml/2006/main}"

EMU_CM = 360000.0
EMU_PT = 12700.0

SIZE_TOL = 0.06          # cm，尺寸区间容差
COLOR_TOL = 6            # 每通道容差
POS_TOL = 0.4            # 百分比位置容差（个别条目单独指定）

STAGE_LABELS = ("start", "check", "problem peak", "solution point", "finish")

BODY_TEXTS = (
    "at 8:40 a.m., the team noticed that the online store was no longer loading "
    "for customers.",
    "the support lead checked alerts, confirmed the outage, and opened an incident "
    "channel for developers and operations.",
    "orders failed repeatedly because the database connection pool was exhausted "
    "during the launch traffic surge.",
    "the team rolled back the latest configuration and switched traffic to a "
    "healthy backup instance.",
    "the site was back up and stable within 10 minutes, order submissions resumed, "
    "and a post-incident review has been scheduled for 4:00 pm this afternoon.",
)

EMOTION_WORDS = ("concerned", "tense and focused", "hopeful", "relieved")
THEME_TEXT = "theme: managing a website outage before a product launch"
LEGEND_WORDS = ("action", "emotion")

# 编号可写作阿拉伯数字，或带圈数字符号
NUM_GLYPHS = {}
for _i, _chars in enumerate(("1➊❶①⑴𝟏", "2➋❷②⑵", "3➌❸③⑶", "4➍❹④⑷", "5➎❺⑤⑸"), start=1):
    for _c in _chars:
        NUM_GLYPHS[_c] = _i

# 颜色区间（细则给出的“近似”区间，逐通道取上下界）
C_CIRCLE = ("DCEEF5", "E3EEF3")      # 浅蓝灰圆形底
C_PILL = ("003341", "003E50")        # 深青蓝阶段标签
C_BOX_LINE = ("86B9DD", "8FBBDC")    # 说明框边线
C_BOX_FILL = ("EDF3F8", "F5F8FB")    # 说明框极浅蓝灰填充
C_BODY_TEXT = ("151B2A", "0B2F45")   # 说明文字深色
C_WARN_TRI = ("F2551A", "FF6B2C")    # 屏幕内橙色三角
C_CURVE = ("FC5E13", "FF4B10")       # 情绪折线橙色
C_GREEN = ("18934E", "20A45B")       # 对勾底绿色
C_THEME = ("0A2E90", "1E5AB4")       # 主题文字蓝色

# 5 个步骤组水平中心所在的页面百分比区间
GROUP_BANDS = ((6.0, 13.0), (26.0, 33.0), (46.0, 53.0), (67.0, 74.0), (86.0, 93.0))
# 4 个情绪标签左端所在的页面百分比区间
EMOTION_BANDS = ((3.0, 13.0), (26.0, 36.0), (64.0, 74.0), (84.0, 95.0))

_MAX_DELTAS_ORDER = [
    5, 1, 3, 3, 3, 1, 3, 1, 3,          # 版面骨架与文本
    5, 3, 5, 3, 5, 3, 3, 3, 5, 3,       # 5 个图标造型与叠放件
    3, 3, 3,                            # 图例框与两行图例
    5, 3, 3, 3,                         # 情绪折线、折线实现、情绪标签、主题文字
    1, 1, 1, 1, 1, 1,                   # 版式与规范性
]
assert sum(_MAX_DELTAS_ORDER) == MAX_SCORE_TOTAL

def _hex2rgb(h):
    h = (h or "").strip().lstrip("#")
    if len(h) != 6:
        return None
    try:
        return (int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))
    except ValueError:
        return None


def _in_band(rgb, lo_hex, hi_hex, tol=COLOR_TOL):
    """逐通道判断 rgb 是否落在两个端点色构成的区间内（含容差）。"""
    if not rgb:
        return False
    lo, hi = _hex2rgb(lo_hex), _hex2rgb(hi_hex)
    for v, a, b in zip(rgb, lo, hi):
        if not (min(a, b) - tol <= v <= max(a, b) + tol):
            return False
    return True


def _lum(rgb):
    return 0.299 * rgb[0] + 0.587 * rgb[1] + 0.114 * rgb[2] if rgb else 0.0


def _is_dark(rgb):
    return bool(rgb) and _lum(rgb) <= 110


def _is_white(rgb, tol=18):
    return bool(rgb) and min(rgb) >= 255 - tol


def _is_light(rgb):
    return bool(rgb) and _lum(rgb) >= 190


def _is_orange(rgb):
    if not rgb:
        return False
    r, g, b = rgb
    return r >= 180 and 30 <= g <= 165 and b <= 110 and (r - g) >= 55 and (r - b) >= 90


def _is_green(rgb):
    if not rgb:
        return False
    r, g, b = rgb
    return g >= 110 and (g - r) >= 35 and (g - b) >= 25


def _is_lightblue(rgb):
    if not rgb:
        return False
    r, g, b = rgb
    return b >= 170 and b >= r and g >= r and (b - r) >= 15


def _between(v, lo, hi, tol=SIZE_TOL):
    return (lo - tol) <= v <= (hi + tol)


def _norm_text(s):
    if not s:
        return ""
    out = []
    for ch in s:
        if ch in "    \t\r\n":
            out.append(" ")
        elif ch in "‘’ʼ":
            out.append("'")
        elif ch in "“”":
            out.append('"')
        elif ch in "‐‑‒–—−":
            out.append("-")
        else:
            out.append(ch)
    return " ".join("".join(out).split()).strip()


def _key(s):
    return _norm_text(s).lower()


class El(object):
    """幻灯片上的一个绘图对象（已换算为幻灯片绝对坐标，单位 cm）。"""

    __slots__ = ("tag", "prst", "x", "y", "w", "h", "rw", "rh", "flip_h", "flip_v",
                 "rot", "fill_kind", "fill_rgb", "fill_alpha", "ln_none", "ln_w",
                 "ln_rgb", "ln_alpha", "ln_dash", "ln_head", "ln_tail", "text",
                 "runs", "align", "anchor", "autofit", "depth", "order", "geom_ok")

    def __init__(self):
        self.tag = ""
        self.prst = ""
        self.x = self.y = self.w = self.h = 0.0
        self.rw = self.rh = 0.0
        self.flip_h = self.flip_v = False
        self.rot = 0.0
        self.fill_kind = "inherit"
        self.fill_rgb = None
        self.fill_alpha = 100.0
        self.ln_none = False
        self.ln_w = None
        self.ln_rgb = None
        self.ln_alpha = 100.0
        self.ln_dash = ""
        self.ln_head = ""
        self.ln_tail = ""
        self.text = ""
        self.runs = []
        self.align = ""
        self.anchor = ""
        self.autofit = ""
        self.depth = 0
        self.order = 0
        self.geom_ok = False

    # --- 便捷属性 ---
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
    def area(self):
        return self.w * self.h

    @property
    def is_leaf(self):
        return self.tag in ("sp", "pic", "cxnSp", "graphicFrame")

    @property
    def filled(self):
        return self.fill_kind in ("solid", "grad", "blip") and self.fill_alpha > 8.0

    def bbox(self):
        return (self.x, self.y, self.x2, self.y2)

    def contains_point(self, px, py, pad=0.0):
        return (self.x - pad <= px <= self.x2 + pad) and (self.y - pad <= py <= self.y2 + pad)

    def contains(self, other, pad=0.0):
        return (self.x - pad <= other.x and other.x2 <= self.x2 + pad and
                self.y - pad <= other.y and other.y2 <= self.y2 + pad)

    def overlaps(self, other, pad=0.0):
        return not (other.x > self.x2 + pad or other.x2 < self.x - pad or
                    other.y > self.y2 + pad or other.y2 < self.y - pad)

    def endpoints(self):
        """线条类对象的起点/终点（考虑翻转，并容忍非法的负尺寸）。"""
        ascending = (self.rh < 0) != bool(self.flip_v)
        if ascending:
            p1, p2 = (self.x, self.y2), (self.x2, self.y)
        else:
            p1, p2 = (self.x, self.y), (self.x2, self.y2)
        if self.flip_h:
            p1, p2 = p2, p1
        return p1, p2

    @property
    def seg_len(self):
        (x1, y1), (x2, y2) = self.endpoints()
        return math.hypot(x2 - x1, y2 - y1)

    @property
    def arrow_right(self):
        """箭头（线端属性）是否位于线段右端。"""
        (x1, _), (x2, _) = self.endpoints()
        tail_at_right = x2 >= x1
        if self.ln_tail and self.ln_tail != "none":
            return tail_at_right
        if self.ln_head and self.ln_head != "none":
            return not tail_at_right
        return False

    @property
    def has_arrow(self):
        return bool((self.ln_tail and self.ln_tail != "none") or
                    (self.ln_head and self.ln_head != "none"))


def _theme_colors(zf):
    """读取主题配色，并按常规 clrMap 映射 bg1/tx1 等别名。"""
    table = {}
    try:
        root = ET.fromstring(zf.read("ppt/theme/theme1.xml"))
    except Exception:
        root = None
    if root is not None:
        scheme = root.find(".//" + A + "clrScheme")
        for node in (scheme if scheme is not None else []):
            name = node.tag.split("}")[-1]
            srgb = node.find(A + "srgbClr")
            sysc = node.find(A + "sysClr")
            val = srgb.get("val") if srgb is not None else (
                sysc.get("lastClr") if sysc is not None else None)
            if val:
                table[name] = val.upper()
    table.setdefault("dk1", "000000")
    table.setdefault("lt1", "FFFFFF")
    table.setdefault("dk2", "000000")
    table.setdefault("lt2", "FFFFFF")
    for alias, target in (("tx1", "dk1"), ("bg1", "lt1"), ("tx2", "dk2"), ("bg2", "lt2")):
        table[alias] = table.get(target, "000000")
    return table


def _color_of(container, theme):
    """从 solidFill 容器读出 (rgb, alpha%)；无 solidFill 返回 (None, None)。"""
    if container is None:
        return None, None
    sf = container.find(A + "solidFill")
    if sf is None:
        return None, None
    rgb, alpha = None, 100.0
    node = sf.find(A + "srgbClr")
    if node is None:
        node = sf.find(A + "schemeClr")
        if node is not None:
            rgb = _hex2rgb(theme.get(node.get("val") or "", ""))
    else:
        rgb = _hex2rgb(node.get("val"))
    if node is not None:
        al = node.find(A + "alpha")
        if al is not None:
            try:
                alpha = int(al.get("val")) / 1000.0
            except (TypeError, ValueError):
                alpha = 100.0
        lm = node.find(A + "lumMod")
        lo = node.find(A + "lumOff")
        if rgb and (lm is not None or lo is not None):
            f = int(lm.get("val")) / 100000.0 if lm is not None else 1.0
            off = int(lo.get("val")) / 100000.0 if lo is not None else 0.0
            rgb = tuple(int(max(0, min(255, v * f + 255 * off))) for v in rgb)
    return rgb, alpha


def _shape_props(node):
    for tag in (P + "spPr", A + "spPr", P + "grpSpPr", A + "grpSpPr"):
        found = node.find(tag)
        if found is not None:
            return found
    return None


def _read_text(node, el, theme):
    body = node.find(P + "txBody")
    if body is None:
        body = node.find(A + "txBody")
    if body is None:
        return
    bpr = body.find(A + "bodyPr")
    if bpr is not None:
        el.anchor = bpr.get("anchor") or ""
        if bpr.find(A + "normAutofit") is not None:
            el.autofit = "normAutofit"
        elif bpr.find(A + "spAutoFit") is not None:
            el.autofit = "spAutoFit"
        else:
            el.autofit = "none"
    paras = []
    for para in body.findall(A + "p"):
        ppr = para.find(A + "pPr")
        if ppr is not None and not el.align:
            el.align = ppr.get("algn") or ""
        chunks = []
        for child in para:
            name = child.tag.split("}")[-1]
            if name == "r":
                t = child.find(A + "t")
                chunks.append(t.text or "" if t is not None else "")
                el.runs.append(_read_run(child, theme))
            elif name == "br":
                chunks.append(" ")
            elif name == "fld":
                t = child.find(A + "t")
                chunks.append(t.text or "" if t is not None else "")
        paras.append("".join(chunks))
    el.text = "\n".join(paras)


def _read_run(run_node, theme):
    pr = run_node.find(A + "rPr")
    info = {"size": None, "bold": False, "italic": False, "rgb": None, "alpha": 100.0}
    if pr is None:
        return info
    sz = pr.get("sz")
    if sz:
        try:
            info["size"] = int(sz) / 100.0
        except ValueError:
            pass
    info["bold"] = pr.get("b") == "1"
    info["italic"] = pr.get("i") == "1"
    rgb, alpha = _color_of(pr, theme)
    info["rgb"] = rgb
    if alpha is not None:
        info["alpha"] = alpha
    return info


def _read_geom(node, el, theme):
    pr = _shape_props(node)
    if pr is None:
        return None
    geom = pr.find(A + "prstGeom")
    if geom is not None:
        el.prst = geom.get("prst") or ""
    elif pr.find(A + "custGeom") is not None:
        el.prst = "custGeom"
    if pr.find(A + "noFill") is not None:
        el.fill_kind = "none"
    elif pr.find(A + "solidFill") is not None:
        el.fill_kind = "solid"
        el.fill_rgb, alpha = _color_of(pr, theme)
        el.fill_alpha = 100.0 if alpha is None else alpha
    elif pr.find(A + "gradFill") is not None:
        el.fill_kind = "grad"
    elif pr.find(A + "blipFill") is not None:
        el.fill_kind = "blip"
    if el.tag == "pic":
        el.fill_kind = "blip"
    ln = pr.find(A + "ln")
    if ln is not None:
        if ln.get("w"):
            try:
                el.ln_w = int(ln.get("w")) / EMU_PT
            except ValueError:
                el.ln_w = None
        el.ln_none = ln.find(A + "noFill") is not None
        el.ln_rgb, alpha = _color_of(ln, theme)
        el.ln_alpha = 100.0 if alpha is None else alpha
        dash = ln.find(A + "prstDash")
        el.ln_dash = (dash.get("val") or "") if dash is not None else "solid"
        head = ln.find(A + "headEnd")
        tail = ln.find(A + "tailEnd")
        el.ln_head = (head.get("type") or "") if head is not None else ""
        el.ln_tail = (tail.get("type") or "") if tail is not None else ""
    return pr


def _xfrm_of(pr):
    if pr is None:
        return None
    xf = pr.find(A + "xfrm")
    return xf


def _walk(container, tf, depth, out, theme, counter):
    """遍历形状树；tf = (ox, oy, sx, sy, cox, coy)，把子空间坐标换算到绝对坐标。"""
    ox, oy, sx, sy, cox, coy = tf
    for node in container:
        tag = node.tag.split("}")[-1]
        if tag not in ("sp", "pic", "cxnSp", "grpSp", "graphicFrame"):
            continue
        el = El()
        el.tag = tag
        el.depth = depth
        el.order = counter[0]
        counter[0] += 1
        pr = _read_geom(node, el, theme)
        xf = _xfrm_of(pr)
        if xf is None and tag == "graphicFrame":
            xf = node.find(P + "xfrm")
        lx = ly = lw = lh = 0.0
        if xf is not None:
            off = xf.find(A + "off")
            ext = xf.find(A + "ext")
            if off is not None:
                lx = float(off.get("x") or 0) / EMU_CM
                ly = float(off.get("y") or 0) / EMU_CM
            if ext is not None:
                lw = float(ext.get("cx") or 0) / EMU_CM
                lh = float(ext.get("cy") or 0) / EMU_CM
            el.flip_h = xf.get("flipH") == "1"
            el.flip_v = xf.get("flipV") == "1"
            try:
                el.rot = float(xf.get("rot") or 0) / 60000.0
            except ValueError:
                el.rot = 0.0
            el.geom_ok = True
        el.rw, el.rh = lw * sx, lh * sy
        el.w, el.h = abs(el.rw), abs(el.rh)
        el.x = ox + (lx - cox) * sx
        el.y = oy + (ly - coy) * sy
        if el.rw < 0:
            el.x += el.rw
        if el.rh < 0:
            el.y += el.rh
        _read_text(node, el, theme)
        out.append(el)
        if tag == "grpSp":
            chx = chy = 0.0
            nsx, nsy = sx, sy
            if xf is not None:
                choff = xf.find(A + "chOff")
                chext = xf.find(A + "chExt")
                if choff is not None:
                    chx = float(choff.get("x") or 0) / EMU_CM
                    chy = float(choff.get("y") or 0) / EMU_CM
                if chext is not None:
                    ccx = float(chext.get("cx") or 0) / EMU_CM
                    ccy = float(chext.get("cy") or 0) / EMU_CM
                    nsx = (el.rw / ccx) if ccx else sx
                    nsy = (el.rh / ccy) if ccy else sy
            _walk(node, (el.x, el.y, nsx, nsy, chx, chy), depth + 1, out, theme, counter)


class Deck(object):
    """一份 pptx 的解析结果。"""

    def __init__(self):
        self.page_w = 0.0
        self.page_h = 0.0
        self.slide_count = 0
        self.els = []
        self.leaves = []
        self.tops = []
        self.texts = []
        self.theme = {}
        self.bg_fill = None
        self.has_shadow = False


def _load(path):
    deck = Deck()
    with zipfile.ZipFile(path) as zf:
        names = set(zf.namelist())
        pres = ET.fromstring(zf.read("ppt/presentation.xml"))
        sz = pres.find(P + "sldSz")
        if sz is not None:
            deck.page_w = float(sz.get("cx") or 0) / EMU_CM
            deck.page_h = float(sz.get("cy") or 0) / EMU_CM
        slides = sorted(n for n in names
                        if n.startswith("ppt/slides/slide") and n.endswith(".xml"))
        deck.slide_count = len(slides)
        if not slides:
            return deck
        deck.theme = _theme_colors(zf)
        raw = zf.read(slides[0])
        text = raw.decode("utf-8", "ignore")
        deck.has_shadow = ("outerShdw" in text or "innerShdw" in text
                           or "prstShdw" in text)
        root = ET.fromstring(raw)
        csld = root.find(P + "cSld")
        if csld is not None:
            bg = csld.find(P + "bg")
            if bg is not None:
                bgpr = bg.find(P + "bgPr")
                rgb, alpha = _color_of(bgpr, deck.theme)
                if rgb and (alpha or 100.0) > 8.0:
                    deck.bg_fill = rgb
            tree = csld.find(P + "spTree")
            if tree is not None:
                _walk(tree, (0.0, 0.0, 1.0, 1.0, 0.0, 0.0), 0, deck.els,
                      deck.theme, [0])
    deck.leaves = [e for e in deck.els if e.is_leaf]
    deck.tops = [e for e in deck.els if e.depth == 0]
    deck.texts = [e for e in deck.els if _norm_text(e.text)]
    return deck


class Ctx(object):
    """把幻灯片上的对象归类为版面语义部件。"""

    def __init__(self, deck):
        self.deck = deck
        self.pw = deck.page_w or 33.87
        self.ph = deck.page_h or 19.05
        self.leaves = deck.leaves
        self.circles = self._find_circles()
        self.legend = self._find_legend()
        self.numbers = self._find_numbers()
        self.num_circles = self._find_num_circles()
        self.pills = self._find_pills()
        self.labels = self._match_labels()
        self.desc_round = False
        self.descs = self._find_descs()
        self.bodies = self._match_bodies()
        self.arrows = self._find_step_arrows()
        self.curve = self._find_curve()
        self.emotions = self._find_emotions()
        self.theme_el = self._find_theme()
        self.icons = [self._icon_set(c) for c in self.circles]

    def px(self, v):
        return v / self.pw * 100.0

    def py(self, v):
        return v / self.ph * 100.0

    def _find_circles(self):
        cand = [e for e in self.leaves
                if e.prst == "ellipse" and e.filled and 2.4 <= e.w <= 3.8
                and abs(e.w - e.h) <= 0.3 and _is_light(e.fill_rgb)
                and 20.0 <= self.py(e.cy) <= 50.0]
        cand.sort(key=lambda e: e.x)
        return cand

    def _find_legend(self):
        cand = [e for e in self.leaves
                if e.prst in ("roundRect", "rect") and 4.5 <= e.w <= 7.5
                and 1.5 <= e.h <= 3.0 and self.px(e.x) <= 10.0
                and self.py(e.y) <= 10.0]
        cand.sort(key=lambda e: (e.y, e.x))
        return cand[0] if cand else None

    def _find_numbers(self):
        out = []
        for e in self.deck.texts:
            t = _norm_text(e.text)
            if len(t) == 1 and t in NUM_GLYPHS:
                out.append((NUM_GLYPHS[t], e))
        out.sort(key=lambda pair: pair[1].x)
        return out


    def _find_num_circles(self):
        res = []
        for _digit, ne in self.numbers:
            best = None
            for e in self.leaves:
                if e.prst != "ellipse" or not e.filled or not _is_dark(e.fill_rgb):
                    continue
                if not (0.4 <= e.w <= 1.8) or abs(e.w - e.h) > 0.25:
                    continue
                if not e.contains_point(ne.cx, ne.cy, 0.15):
                    continue
                if best is None or e.area < best.area:
                    best = e
            res.append(best)
        return res

    def _find_pills(self):
        cand = [e for e in self.leaves
                if e.prst in ("roundRect", "rect") and e.filled and _is_dark(e.fill_rgb)
                and 3.0 <= e.w <= 5.6 and 0.6 <= e.h <= 1.5
                and 30.0 <= self.py(e.cy) <= 52.0]
        cand.sort(key=lambda e: e.x)
        return cand

    def _match_labels(self):
        out = []
        for p in self.pills:
            found = None
            for e in self.deck.texts:
                if _key(e.text) not in STAGE_LABELS:
                    continue
                if abs(e.cy - p.cy) <= 0.7 and abs(e.cx - p.cx) <= 0.9:
                    found = e
                    break
            out.append(found)
        return out

    def _find_descs(self):
        def cand(prsts):
            got = [e for e in self.leaves
                   if e.prst in prsts and 4.2 <= e.w <= 6.0 and 3.8 <= e.h <= 5.6
                   and 40.0 <= self.py(e.y) <= 56.0]
            got.sort(key=lambda e: e.x)
            return got
        rounded = cand(("roundRect",))
        self.desc_round = len(rounded) >= 5
        if len(rounded) >= 5:
            return rounded
        extra = [e for e in cand(("rect",)) if e not in rounded]
        return sorted(rounded + extra, key=lambda e: e.x)


    def _match_bodies(self):
        out = []
        for d in self.descs[:5]:
            if len(_norm_text(d.text)) > 30:
                out.append(d)
                continue
            best = None
            for e in self.deck.texts:
                if e is d or len(_norm_text(e.text)) <= 30:
                    continue
                if not d.contains_point(e.cx, e.cy, 0.25):
                    continue
                if best is None or e.area > best.area:
                    best = e
            out.append(best)
        return out

    def _find_step_arrows(self):
        out = []
        for e in self.leaves:
            if e.prst not in ("line", "straightConnector1", "bentConnector3"):
                continue
            if e.ln_none or not e.has_arrow or not _is_dark(e.ln_rgb):
                continue
            if e.h > 0.25 or not (1.5 <= e.w <= 3.4):
                continue
            if not (26.0 <= self.py(e.cy) <= 42.0):
                continue
            out.append(e)
        out.sort(key=lambda e: e.x)
        return out

    def _find_curve(self):
        out = []
        for e in self.leaves:
            if e.prst not in ("line", "straightConnector1", "bentConnector3",
                              "curvedConnector3", "curvedConnector4", "custGeom",
                              "freeform"):
                continue
            if e.ln_none or e.ln_alpha <= 8 or not _is_orange(e.ln_rgb):
                continue
            if self.py(e.cy) < 62.0:
                continue
            out.append(e)
        out.sort(key=lambda e: e.x)
        return out

    def _find_emotions(self):
        out = [e for e in self.deck.texts if _key(e.text) in EMOTION_WORDS]
        out.sort(key=lambda e: e.x)
        return out


    def _find_theme(self):
        for e in self.deck.texts:
            if _key(e.text).startswith("theme:"):
                return e
        return None

    def _icon_set(self, circle):
        skip = set([id(circle)])
        for group in (self.pills, self.descs, self.arrows, self.curve,
                      self.emotions, self.labels, self.bodies, self.num_circles):
            for e in group:
                if e is not None:
                    skip.add(id(e))
        for _digit, e in self.numbers:
            skip.add(id(e))
        if self.legend is not None:
            skip.add(id(self.legend))
        if self.theme_el is not None:
            skip.add(id(self.theme_el))
        out = []
        for e in self.leaves:
            if id(e) in skip:
                continue
            if not circle.contains_point(e.cx, e.cy, 0.12):
                continue
            if e.w > circle.w * 1.2 or e.h > circle.h * 1.3:
                continue
            out.append(e)
        return out


def _pic_impl(els, lo=1.3, hi=3.6):
    """图标区域内具备图标尺寸的位图（细则允许图标由小尺寸图片实现）。"""
    return [e for e in els if e.tag == "pic" and min(e.w, e.h) >= lo
            and max(e.w, e.h) <= hi]


def _white_bits(els, host):
    """host 内部的白色零件（填充为白或线条为白）。"""
    out = []
    for e in els:
        if e is host or not host.contains_point(e.cx, e.cy, 0.05):
            continue
        if e.filled and _is_white(e.fill_rgb):
            out.append(e)
        elif not e.ln_none and e.ln_alpha > 8 and _is_white(e.ln_rgb):
            out.append(e)
    return out


def _has_exclaim(els, host):
    """host 内是否有白色感叹号（位图内部不再拆解，直接认定）。"""
    if host.tag == "pic":
        return True
    for e in els:
        t = _norm_text(e.text)
        if t and t.strip("!！i") == "" and host.contains_point(e.cx, e.cy, 0.1):
            return True
    for e in _white_bits(els, host):
        if e.h >= 0.12 and e.w <= 0.22:
            return True
    return False


def _check_mark_ok(els, host):
    """host（绿色圆）内白色对勾是否为两笔相接、连续完整。"""
    if host.tag == "pic":
        return True
    strokes = [e for e in _white_bits(els, host)
               if (e.w + e.h) > 0.12 and min(e.w, e.h) <= 0.5]
    if not strokes:
        return False
    for e in strokes:
        if e.rw < 0 or e.rh < 0:
            return False              # 负尺寸为非法几何，渲染为断裂笔画
    if len(strokes) == 1:
        return strokes[0].prst in ("custGeom", "freeform")
    for i, a in enumerate(strokes):
        near = None
        for j, b in enumerate(strokes):
            if i == j:
                continue
            for pa in a.endpoints():
                for pb in b.endpoints():
                    d = math.hypot(pa[0] - pb[0], pa[1] - pb[1])
                    if near is None or d < near:
                        near = d
        if near is None or near > 0.12:
            return False
    return True


def _union_box(els):
    box = El()
    box.tag = "sp"
    box.x = min(e.x for e in els)
    box.y = min(e.y for e in els)
    box.w = max(e.x2 for e in els) - box.x
    box.h = max(e.y2 for e in els) - box.y
    return box


def _laptop(els):
    """返回 (是否为笔记本造型, 屏幕外框元素)。"""
    frames = [e for e in els
              if e.prst in ("rect", "roundRect", "custGeom") and e.tag == "sp"
              and 1.2 <= e.w <= 3.2 and 0.8 <= e.h <= 2.6 and e.w >= e.h
              and (_is_dark(e.fill_rgb)
                   or (not e.ln_none and e.ln_alpha > 8 and _is_dark(e.ln_rgb)))]
    frames.sort(key=lambda e: -e.area)
    screen = frames[0] if frames else None
    pics = _pic_impl(els)
    if pics:
        return True, (screen or max(pics, key=lambda e: e.area))
    if screen is None:
        return False, None
    bases = [e for e in els
             if e is not screen and e.h <= 0.3 and e.w >= 0.5
             and screen.y2 - 0.15 <= e.cy <= screen.y2 + 0.8
             and abs(e.cx - screen.cx) <= screen.w * 0.6]
    if not bases:
        return False, screen
    base = max(bases, key=lambda e: e.w)
    stems = [e for e in els
             if e.w <= 0.2 and 0.08 <= e.h <= 0.9
             and screen.y2 - 0.15 <= e.cy <= base.cy + 0.15
             and abs(e.cx - screen.cx) <= 0.5]
    if stems:
        return False, screen              # 立柱→台式支架造型，非笔记本
    if base.w < screen.w * 0.85:
        return False, screen
    top, bot = min(screen.y, base.y), max(screen.y2, base.y2)
    if not (_between(max(screen.w, base.w), 2.0, 3.2)
            and _between(bot - top, 1.6, 3.2)):
        return False, screen
    return True, screen


def _screen_triangle(els, screen):
    tris = []
    for e in els:
        if e.prst == "triangle" and e.filled and (_is_orange(e.fill_rgb)
                                                  or _in_band(e.fill_rgb, *C_WARN_TRI)):
            tris.append(e)
        elif (e.tag == "pic" and abs(e.w - e.h) <= 0.35 and 0.9 <= e.w <= 1.9
              and screen is not None and screen.contains_point(e.cx, e.cy, 0.25)):
            tris.append(e)
    for t in tris:
        if not (_between(t.w, 1.0, 1.8) and _between(t.h, 1.0, 1.8)):
            continue
        if not _has_exclaim(els, t):
            continue
        if screen is not None and not screen.contains(t, 0.05):
            continue
        return True
    return False


def _magnifier(els):
    """返回 (是否为放大镜造型, 镜片元素)。"""
    lens, handle_ok = None, False
    for e in els:
        if (e.prst == "ellipse" and abs(e.w - e.h) <= 0.3 and _between(e.w, 1.2, 2.0)
                and not e.ln_none and e.ln_alpha > 8 and _is_dark(e.ln_rgb)):
            lens = e
            break
    if lens is not None:
        handle_ok = any(e is not lens and 0.25 <= e.w <= 1.3 and 0.25 <= e.h <= 1.3
                        and e.cx > lens.cx and e.cy > lens.cy for e in els)
    if lens is None or not handle_ok:
        for e in _pic_impl(els, 1.15, 2.1):
            if abs(e.w - e.h) <= 0.3 and _between(e.w, 1.2, 2.0):
                lens, handle_ok = e, True
                break
    bars = [e for e in els
            if e.filled and _is_lightblue(e.fill_rgb) and e.h > e.w
            and 0.05 <= e.w <= 0.45 and e.h >= 0.15]
    heights = sorted(set(round(e.h, 2) for e in bars))
    panel = [e for e in els
             if e is not lens and max(e.w, e.h) >= 0.8
             and ((not e.ln_none and e.ln_alpha > 8 and _is_lightblue(e.ln_rgb))
                  or (e.filled and _is_lightblue(e.fill_rgb)))]
    ok = bool(lens) and handle_ok and len(bars) >= 3 and len(heights) >= 2 and bool(panel)
    return ok, lens


def _badge(els, lens):
    cands = []
    for e in els:
        if e is lens:
            continue
        if e.prst == "ellipse" and e.filled and (_is_orange(e.fill_rgb)
                                                 or _in_band(e.fill_rgb, *C_WARN_TRI)):
            cands.append(e)
        elif e.tag == "pic" and abs(e.w - e.h) <= 0.2 and _between(e.w, 0.6, 1.2):
            cands.append(e)
    for b in cands:
        if not (_between(b.w, 0.6, 1.2) and _between(b.h, 0.6, 1.2)):
            continue
        if not (b.cx > lens.cx and b.cy > lens.cy):
            continue
        if not _has_exclaim(els, b):
            continue
        if math.hypot(b.cx - lens.cx, b.cy - lens.cy) > (lens.w + b.w) / 2.0 + 0.05:
            continue                      # 与镜片边缘不相叠
        return True
    return False


def _server(els):
    """返回 (是否为服务器造型, 机架整体外接框)。"""
    racks = [e for e in els
             if e.prst in ("rect", "roundRect") and e.filled and _is_dark(e.fill_rgb)
             and 1.2 <= e.w <= 2.6 and 0.25 <= e.h <= 0.65]
    racks.sort(key=lambda e: e.y)
    box = _union_box(racks) if racks else None
    if len(racks) < 3:
        pics = _pic_impl(els, 1.4)
        if pics:
            return True, (box or max(pics, key=lambda e: e.area))
        return False, box
    ok = True
    widths = [r.w for r in racks[:3]]
    if max(widths) - min(widths) > 0.2 or not _between(max(widths), 1.5, 2.5):
        ok = False
    for r in racks[:3]:
        bits = _white_bits(els, r)
        lamps = [e for e in bits if e.cx < r.cx and max(e.w, e.h) <= 0.4
                 and abs(e.w - e.h) <= 0.12]
        lines = [e for e in bits if e.cx > r.cx and e.w >= 0.2 and e.h <= 0.15]
        if not lamps or not lines:
            ok = False
    return ok, box


def _corner_triangle(els, box):
    """box 右下方的橙色三角警示标志。"""
    tris = []
    for e in els:
        if e.prst == "triangle" and e.filled and (_is_orange(e.fill_rgb)
                                                  or _in_band(e.fill_rgb, *C_WARN_TRI)):
            tris.append(e)
        elif e.tag == "pic" and abs(e.w - e.h) <= 0.3 and _between(e.w, 1.0, 1.8):
            tris.append(e)
    for t in tris:
        if not (_between(t.w, 1.0, 1.8) and _between(t.h, 1.0, 1.8)):
            continue
        if box is not None and not (t.cx > box.cx and t.cy > box.cy):
            continue
        if not _has_exclaim(els, t):
            continue
        return True
    return False


def _cloud(els):
    """返回 (云朵造型是否达标, 云朵元素)。"""
    for e in els:
        if e.prst in ("cloud", "cloudCallout"):
            ok = (_between(e.w, 1.6, 2.6) and _between(e.h, 1.0, 1.8)
                  and not e.ln_none and e.ln_alpha > 8 and _is_dark(e.ln_rgb)
                  and e.filled and e.fill_alpha > 50.0 and _is_white(e.fill_rgb))
            return ok, e
    pics = [e for e in els if e.tag == "pic" and min(e.w, e.h) >= 0.9
            and max(e.w, e.h) <= 3.4]
    if pics:
        return True, max(pics, key=lambda e: e.area)
    frees = [e for e in els if e.prst in ("custGeom", "freeform") and e.w >= 1.4]
    if frees:
        e = max(frees, key=lambda x: x.area)
        ok = (_between(e.w, 1.6, 2.6) and _between(e.h, 1.0, 1.8)
              and not e.ln_none and e.ln_alpha > 8 and _is_dark(e.ln_rgb)
              and e.filled and e.fill_alpha > 50.0 and _is_white(e.fill_rgb))
        return ok, e
    return False, None


def _monitor(els):
    """返回 (是否为显示器造型, 屏幕外框元素)。"""
    frames = [e for e in els
              if e.prst in ("rect", "roundRect", "custGeom") and e.tag == "sp"
              and 1.2 <= e.w <= 2.8 and 0.8 <= e.h <= 2.2 and e.w >= e.h * 0.9
              and (_is_dark(e.fill_rgb)
                   or (not e.ln_none and e.ln_alpha > 8 and _is_dark(e.ln_rgb)))]
    frames.sort(key=lambda e: -e.area)
    screen = frames[0] if frames else None
    pics = _pic_impl(els, 1.3)
    if pics:
        return True, (screen or max(pics, key=lambda e: e.area))
    if screen is None:
        return False, None
    stems = [e for e in els if e.w <= 0.25 and 0.08 <= e.h <= 0.9
             and screen.y2 - 0.2 <= e.cy <= screen.y2 + 1.0
             and abs(e.cx - screen.cx) <= 0.6]
    bases = [e for e in els if e.h <= 0.3 and e.w >= 0.4
             and screen.y2 <= e.cy <= screen.y2 + 1.2
             and abs(e.cx - screen.cx) <= screen.w * 0.6]
    if not stems or not bases:
        return False, screen
    bot = max(max(b.y2 for b in bases), screen.y2)
    ok = _between(screen.w, 1.6, 2.6) and _between(bot - screen.y, 1.2, 2.2)
    return ok, screen


def _green_badge(els, host, mode):
    """host 右下方（corner）或中央（center）的绿色对勾圆。"""
    cands = []
    for e in els:
        if e is host:
            continue
        if e.prst == "ellipse" and e.filled and (_is_green(e.fill_rgb)
                                                 or _in_band(e.fill_rgb, *C_GREEN)):
            cands.append(e)
        elif e.tag == "pic" and abs(e.w - e.h) <= 0.2 and _between(e.w, 0.7, 1.7):
            cands.append(e)
    for g in cands:
        if not (_between(g.w, 0.8, 1.6) and _between(g.h, 0.8, 1.6)):
            continue
        if mode == "corner":
            if not (g.cx > host.cx and g.cy > host.cy):
                continue
        elif abs(g.cx - host.cx) > 0.45 or abs(g.cy - host.cy) > 0.45:
            continue
        if not _check_mark_ok(els, g):
            continue
        return True
    return False


def _covered(el, leaves):
    """el 是否被其后绘制的不透明形状整体覆盖。"""
    for other in leaves:
        if other is el or other.order <= el.order:
            continue
        if not (other.filled and other.fill_alpha > 85.0):
            continue
        if other.contains(el):
            return True
    return False


def _sizes_of(el):
    return [r["size"] for r in el.runs if r.get("size")]


def _colors_of(el):
    return [r["rgb"] for r in el.runs if r.get("rgb")]


def _check_layout(ctx):
    """细则第 1–9 条：版面骨架与文本。"""
    f = [False] * 9
    circles, pills, descs = ctx.circles, ctx.pills, ctx.descs
    nums, labels, bodies = ctx.numbers, ctx.labels, ctx.bodies
    full = (len(circles) == 5 and len(pills) >= 5 and len(descs) >= 5
            and len(nums) == 5)
    if full:
        ok = True
        for i in range(5):
            ne, c, p, d = nums[i][1], circles[i], pills[i], descs[i]
            cs = [ne.cx, c.cx, p.cx, d.cx]
            if max(cs) - min(cs) > 1.0:
                ok = False
            if not (ne.cy < c.cy < p.cy < d.cy):
                ok = False
            if ne.cy > c.y + 0.6 or p.y < c.y2 - 0.6 or d.y < p.y2 - 0.6:
                ok = False
            lo, hi = GROUP_BANDS[i]
            if not (lo - POS_TOL <= ctx.px(sum(cs) / 4.0) <= hi + POS_TOL):
                ok = False
        f[0] = ok
    if len(nums) == 5:
        ok = [d for d, _e in nums] == [1, 2, 3, 4, 5]
        for idx, (_dg, e) in enumerate(nums):
            top = e.y
            nc = ctx.num_circles[idx] if idx < len(ctx.num_circles) else None
            if nc is not None:
                top = min(top, nc.y)
            if not (14.0 - POS_TOL <= ctx.py(top) <= 26.0 + POS_TOL):
                ok = False
            sizes = _sizes_of(e)
            if not sizes or not all(14.0 <= s <= 24.0 for s in sizes):
                ok = False
            cols = _colors_of(e)
            if _norm_text(e.text) in "12345":
                if nc is None or not cols or not all(_is_white(c) for c in cols):
                    ok = False
            elif not cols or not all(_is_dark(c) for c in cols):
                ok = False
        f[1] = ok
    if len(circles) == 5:
        f[2] = all(_in_band(c.fill_rgb, *C_CIRCLE) and _between(c.w, 2.6, 3.5)
                   and _between(c.h, 2.6, 3.5)
                   and 22.0 - POS_TOL <= ctx.py(c.y) <= 28.0 + POS_TOL
                   for c in circles)
    if len(pills) >= 5:
        f[3] = all(_in_band(p.fill_rgb, *C_PILL) and _between(p.h, 0.8, 1.3)
                   and _between(p.w, 3.5, 5.0)
                   and 36.0 - POS_TOL <= ctx.py(p.y) <= 43.0 + POS_TOL
                   for p in pills[:5])
    if len(pills) >= 5 and all(labels[:5]):
        ok = True
        for i, e in enumerate(labels[:5]):
            if _key(e.text) != STAGE_LABELS[i]:
                ok = False
            sizes = _sizes_of(e)
            if not sizes or not all(13.0 <= s <= 18.0 for s in sizes):
                ok = False
            if not e.runs or not all(r["bold"] for r in e.runs):
                ok = False
            cols = _colors_of(e)
            if not cols or not all(_is_white(c) for c in cols):
                ok = False
            if abs(e.cx - pills[i].cx) > 0.4 and e.align != "ctr":
                ok = False
        f[4] = ok
    if len(circles) == 5:
        hit = 0
        for i in range(4):
            left, right = circles[i], circles[i + 1]
            for a in ctx.arrows:
                if not (left.x2 - 0.6 <= a.x and a.x2 <= right.x + 0.6):
                    continue
                if not _between(a.w, 1.8, 3.0, 0.1) or not a.arrow_right:
                    continue
                if not (31.0 - POS_TOL <= ctx.py(a.cy) <= 38.0 + POS_TOL):
                    continue
                hit += 1
                break
        f[5] = hit == 4
    if len(descs) >= 5:
        ok = ctx.desc_round
        for d in descs[:5]:
            if not (_between(d.w, 4.6, 5.6) and _between(d.h, 4.2, 5.2)):
                ok = False
            if not (44.0 - POS_TOL <= ctx.py(d.y) <= 50.0 + POS_TOL):
                ok = False
            if d.ln_none or d.ln_alpha <= 8 or not _in_band(d.ln_rgb, *C_BOX_LINE):
                ok = False
            if d.ln_dash not in ("", "solid"):
                ok = False
            if not (d.filled and _in_band(d.fill_rgb, *C_BOX_FILL)):
                ok = False
            if d.fill_rgb and min(d.fill_rgb) >= 250:
                ok = False                # 纯白填充
        f[6] = ok
    if len(bodies) >= 5 and all(bodies[:5]):
        ok = True
        for b in bodies[:5]:
            sizes = _sizes_of(b)
            if not sizes or not all(10.0 <= s <= 14.0 for s in sizes):
                ok = False
            if b.align != "ctr" or b.anchor != "ctr":
                ok = False
            for c in _colors_of(b):
                if not (_is_dark(c) or _in_band(c, *C_BODY_TEXT)):
                    ok = False
        f[7] = ok
        f[8] = all(_key(bodies[i].text) == BODY_TEXTS[i] for i in range(5))
    return f


def _check_icons(ctx):
    """细则第 10–19 条：5 个步骤图标造型与其叠放件。"""
    f = [False] * 10
    if len(ctx.circles) != 5:
        return f
    ic = ctx.icons
    laptop_ok, screen = _laptop(ic[0])
    f[0] = laptop_ok
    if laptop_ok:
        f[1] = _screen_triangle(ic[0], screen)
    mag_ok, lens = _magnifier(ic[1])
    f[2] = mag_ok
    if lens is not None:
        f[3] = _badge(ic[1], lens)
    srv_ok, srv_box = _server(ic[2])
    f[4] = srv_ok
    if srv_box is not None:
        f[5] = _corner_triangle(ic[2], srv_box)
    cloud_ok, cloud = _cloud(ic[3])
    f[6] = cloud_ok
    if cloud is not None:
        f[7] = _green_badge(ic[3], cloud, "corner")
    mon_ok, mon_screen = _monitor(ic[4])
    f[8] = mon_ok
    if mon_screen is not None:
        f[9] = _green_badge(ic[4], mon_screen, "center")
    return f


def _check_legend(ctx):
    """细则第 20–22 条：图例框与两行图例。"""
    f = [False] * 3
    box = ctx.legend
    if box is None:
        return f
    fill_ok = (not box.filled) or box.fill_alpha <= 8 or _is_white(box.fill_rgb)
    line_ok = (not box.ln_none and box.ln_alpha > 8 and _is_dark(box.ln_rgb)
               and box.ln_dash in ("", "solid")
               and (box.ln_w is None or 0.85 <= box.ln_w <= 3.15))
    f[0] = (box.prst == "roundRect" and fill_ok and line_ok
            and _between(box.w, 5.0, 7.0) and _between(box.h, 1.8, 2.8)
            and ctx.px(box.x) <= 8.0 + POS_TOL and ctx.py(box.y) <= 8.0 + POS_TOL)
    inner = [e for e in ctx.leaves if e is not box and box.contains_point(e.cx, e.cy, 0.35)]
    lines = [e for e in inner
             if e.prst in ("line", "straightConnector1", "bentConnector3")
             and not e.ln_none and e.ln_alpha > 8]
    dark = [e for e in lines if _is_dark(e.ln_rgb)]
    orange = [e for e in lines if _is_orange(e.ln_rgb)]
    texts = dict((_key(e.text), e) for e in inner if _key(e.text) in LEGEND_WORDS)
    ae = texts.get("action")
    if ae is not None and dark:
        arrow = None
        for e in dark:
            if (_between(e.seg_len, 2.0, 3.0, 0.1) and e.arrow_right
                    and e.ln_dash in ("", "solid")):
                arrow = e
                break
        sizes = _sizes_of(ae)
        f[1] = (arrow is not None and ae.cx > arrow.cx and bool(sizes)
                and all(12.0 <= s <= 18.0 for s in sizes)
                and bool(ae.runs) and all(r["bold"] for r in ae.runs)
                and all(_is_dark(c) for c in _colors_of(ae)))
    ee = texts.get("emotion")
    segs = [e for e in orange if e.seg_len > 0.05]
    if ee is not None and segs:
        main = [e for e in segs
                if _between(e.seg_len, 2.0, 3.0, 0.1) and e.arrow_right
                and e.ln_dash not in ("", "solid")]
        sizes = _sizes_of(ee)
        f[2] = (len(segs) <= 2 and bool(main)
                and ee.cx > max(e.cx for e in segs) and bool(sizes)
                and all(12.0 <= s <= 18.0 for s in sizes)
                and bool(ee.runs) and all(r["bold"] for r in ee.runs)
                and all(_is_orange(c) or _in_band(c, *C_CURVE)
                        for c in _colors_of(ee)))
    return f


def _trend_ok(ctx, pts):
    """折线走势：先下降后上升、再下降、再上升，波峰/波谷位置在指定区间内。"""
    if len(pts) < 4:
        return False

    def window(lo, hi):
        return [p for p in pts if lo <= ctx.px(p[0]) <= hi]

    left, mid = window(18.0, 50.0), window(44.0, 70.0)
    if not left or not mid:
        return False
    peak = min(left, key=lambda p: p[1])
    valley = max(mid, key=lambda p: p[1])
    if not (32.0 - POS_TOL <= ctx.px(peak[0]) <= 42.0 + POS_TOL):
        return False
    if not (52.0 - POS_TOL <= ctx.px(valley[0]) <= 62.0 + POS_TOL):
        return False
    before = [p for p in pts if p[0] < peak[0]]
    if not before or max(p[1] for p in before) <= pts[0][1] + 0.08:
        return False                      # 起始段未先下降
    if pts[-1][1] >= valley[1] - 0.08:
        return False                      # 末段未再上升
    return peak[1] < valley[1] - 0.3


def _check_curve(ctx):
    """细则第 23–26 条：情绪折线、折线实现方式、情绪标签与主题文字。"""
    f = [False] * 4
    segs = ctx.curve
    if segs:
        pts = []
        for e in segs:
            p1, p2 = e.endpoints()
            pts.append(p1)
            pts.append(p2)
        pts.sort(key=lambda p: p[0])
        xs = [p[0] for p in pts]
        ys = [p[1] for p in pts]
        span_ok = (ctx.px(min(xs)) <= 6.0 + POS_TOL
                   and ctx.px(max(xs)) >= 92.0 - POS_TOL)
        band_ok = (ctx.py(min(ys)) >= 72.0 - POS_TOL
                   and ctx.py(max(ys)) <= 90.0 + POS_TOL)
        arrow_ok = any(e.has_arrow for e in segs if e.x2 >= max(xs) - 0.6)
        f[0] = span_ok and band_ok and arrow_ok and _trend_ok(ctx, pts)
        f[1] = (len(segs) <= 30 and all(e.seg_len >= 2.0 - 0.05 for e in segs)
                and all(e.ln_dash not in ("", "solid") for e in segs)
                and any(e.has_arrow for e in segs))
    ems = ctx.emotions
    if len(ems) == 4:
        ok = True
        for i, e in enumerate(ems):
            if _key(e.text) != EMOTION_WORDS[i]:
                ok = False
            sizes = _sizes_of(e)
            if not sizes or not all(12.0 <= s <= 18.0 for s in sizes):
                ok = False
            if not e.runs or not all(r["italic"] for r in e.runs):
                ok = False
            for c in _colors_of(e):
                if not (_is_orange(c) or _in_band(c, *C_CURVE)):
                    ok = False
            if not (78.0 - POS_TOL <= ctx.py(e.cy) <= 95.0 + POS_TOL):
                ok = False
            lo, hi = EMOTION_BANDS[i]
            if not (lo - POS_TOL <= ctx.px(e.x) <= hi + POS_TOL):
                ok = False
        f[2] = ok
    te = ctx.theme_el
    if te is not None:
        sizes = _sizes_of(te)
        f[3] = (_key(te.text) == THEME_TEXT and bool(sizes)
                and all(13.0 <= s <= 18.0 for s in sizes)
                and bool(te.runs) and all(r["italic"] for r in te.runs)
                and all(_in_band(c, *C_THEME) for c in _colors_of(te))
                and 90.0 - POS_TOL <= ctx.py(te.y) <= 97.0 + POS_TOL
                and 45.0 - POS_TOL <= ctx.px(te.cx) <= 55.0 + POS_TOL)
    return f


def _check_norms(ctx):
    """细则第 27–32 条：版式与规范性。"""
    f = [False] * 6
    deck = ctx.deck
    page_ok = (32.9 <= ctx.pw <= 34.1 and 18.4 <= ctx.ph <= 19.6)
    inside = True
    for e in ctx.leaves:
        if e.w <= 0.0 and e.h <= 0.0:
            continue
        if (e.x < -0.2 or e.y < -0.2 or e.x2 > ctx.pw + 0.2
                or e.y2 > ctx.ph + 0.2):
            inside = False
            break
    f[0] = page_ok and inside
    if (len(ctx.circles) == 5 and len(ctx.pills) >= 5 and len(ctx.descs) >= 5
            and len(ctx.numbers) == 5):
        ok = True
        for i in range(5):
            cs = [ctx.numbers[i][1].cx, ctx.circles[i].cx, ctx.pills[i].cx,
                  ctx.descs[i].cx]
            if max(cs) - min(cs) > 0.5:
                ok = False
        f[1] = ok
    bodies = [b for b in ctx.bodies[:5] if b is not None]
    if len(bodies) == 5:
        sizes = []
        for b in bodies:
            sizes.extend(_sizes_of(b))
        f[2] = bool(sizes) and (max(sizes) - min(sizes) <= 0.2)
    ok = True
    for idx, nc in enumerate(ctx.num_circles):
        if nc is None:
            continue
        digit = ctx.numbers[idx][1]
        for e in ctx.leaves:
            if e is nc or e is digit or not e.filled:
                continue
            if not nc.contains_point(e.cx, e.cy, 0.02):
                continue
            if max(e.w, e.h) > nc.w * 0.9:
                continue
            if _is_white(e.fill_rgb) or _is_light(e.fill_rgb):
                ok = False                # 编号圆内多余的白色/浅色小圆点
    if len(ctx.icons) >= 2:
        _mag_ok, lens = _magnifier(ctx.icons[1])
        if (lens is not None and lens.tag != "pic" and lens.filled
                and lens.fill_alpha > 20.0):
            ok = False                    # 镜片为不透明填充，遮挡其后柱条
    f[3] = ok
    allowed = set(BODY_TEXTS) | set(STAGE_LABELS) | set(EMOTION_WORDS) \
        | set(LEGEND_WORDS) | set([THEME_TEXT])
    ok = True
    for e in deck.texts:
        t = _norm_text(e.text)
        if e.w <= 0.05 or e.h <= 0.05:
            ok = False
        if (e.x < -0.2 or e.y < -0.2 or e.x2 > ctx.pw + 0.2
                or e.y2 > ctx.ph + 0.2):
            ok = False
        if _covered(e, ctx.leaves):
            ok = False
        if _key(t) in allowed:
            continue
        if len(t) == 1 and t in NUM_GLYPHS:
            continue
        if all(not ch.isalnum() for ch in t):
            continue                      # 图标内的感叹号、对勾等符号
        ok = False
    f[4] = ok
    ok = True
    if deck.bg_fill is not None and not _is_white(deck.bg_fill, 12):
        ok = False
    if deck.has_shadow:
        ok = False
    for e in ctx.leaves:
        if e.w >= ctx.pw * 0.9 and e.h >= ctx.ph * 0.9:
            if e.filled or (not e.ln_none and e.ln_alpha > 8):
                ok = False
    f[5] = ok
    return f


def _check_dim1(deck, path):
    reasons = []
    if not path.lower().endswith(SUPPORTED_EXTS):
        reasons.append("交付文件格式不是.pptx")
    if deck.slide_count != 1:
        reasons.append("幻灯片页数为%d页，不等于1页" % deck.slide_count)
    for e in deck.leaves:
        if (e.tag == "pic" and deck.page_w and e.w >= deck.page_w * 0.85
                and e.h >= deck.page_h * 0.85):
            reasons.append("版面由整张位图充当，图形与文字均不可编辑")
            break
    if len(deck.texts) < 18:
        reasons.append("可单独选中修改的文本对象仅%d个，编号/阶段标签/说明文字/"
                       "图例文字/情绪标签/主题文字未全部以文本实现" % len(deck.texts))
    if len(deck.leaves) < 30:
        reasons.append("可单独点选、单独编辑的对象仅%d个，少于30个" % len(deck.leaves))
    if reasons:
        return False, "；".join(reasons)
    # 约定：维度一通过时 dim1_reason 保持空串，只有未通过才写错误摘要。
    return True, ""


RULE_TEXTS = [
    "页面自左至右并列出现5个流程步骤组，每组自上而下依次为“编号—浅蓝色圆形图标底—深色圆角矩形阶段标签—说明文本框”四层结构；5组的水平中心分别位于页面从左至右6%–13%、26%–33%、46%–53%、67%–74%、86%–93%",
    "5个编号自左至右依次为1、2、3、4、5，位于页面从上至下14%–26%，表现为深色实心圆形内白色数字或深色圆形数字符号，字号14–24磅",
    "5个浅蓝色圆形填充为浅蓝灰色（近似#DCEEF5–#E3EEF3），直径均为2.6–3.5cm，顶端位于页面从上至下22%–28%",
    "5个深色圆角矩形阶段标签填充为深青蓝色（近似#003341–#003E50），高0.8–1.3cm、宽3.5–5.0cm，顶端位于页面从上至下36%–43%",
    "5个阶段标签内自左至右依次出现白色加粗文本“start”“check”“problem peak”“solution point”“finish”，字号13–18磅，水平居中显示于所在标签之内",
    "相邻两个浅蓝色圆形之间各出现1条水平深色实线箭头，共4条，位于页面从上至下31%–38%，长度1.8–3.0cm，箭头指向右侧",
    "5个说明文本框为圆角矩形，宽4.6–5.6cm、高4.2–5.2cm，顶端位于页面从上至下44%–50%，边线为浅蓝色实线（近似#86B9DD–#8FBBDC），填充为极浅蓝灰色（近似#EDF3F8–#F5F8FB），非纯白",
    "5个说明文本框内文字均为10–14磅、水平居中、垂直居中，颜色为黑色或深色（近似#151B2A–#0B2F45）",
    "5个说明文本框内自左至右依次为“At 8:40 a.m., the team noticed that the online store was no longer loading for customers.”“The support lead checked alerts, confirmed the outage, and opened an incident channel for developers and operations.”“Orders failed repeatedly because the database connection pool was exhausted during the launch traffic surge.”“The team rolled back the latest configuration and switched traffic to a healthy backup instance.”“The site was back up and stable within 10 minutes, order submissions resumed, and a post-incident review has been scheduled for 4:00 PM this afternoon.”，文字完整无缺字、无错拼",
    "第1个浅蓝色圆形内出现深色笔记本电脑造型（由形状组合或图片实现）：由屏幕外框矩形与其正下方的横向宽底座条构成，非立柱加圆盘的台式支架造型，整体宽2.0–3.2cm、高1.6–3.2cm",
    "第1个浅蓝色圆形内的笔记本电脑屏幕区域内叠放橙色实心三角形警示标志（近似#F2551A–#FF6B2C，由形状或图片实现）：宽高均为1.0–1.8cm，三角形内有白色感叹号，且三角形完整位于屏幕外框的边线之内、未越出下边线或右边线",
    "第2个浅蓝色圆形内出现深色放大镜造型（由形状组合或图片实现）：圆环镜片加右下斜向手柄，镜片外径1.2–2.0cm；镜片右后方有浅蓝色柱状图面板，面板内有3根及以上高度不等的竖直柱条",
    "第2个浅蓝色圆形内放大镜镜片的右下方出现橙色实心圆形（由形状或图片实现）：直径0.6–1.2cm，圆内为白色感叹号，且该圆与镜片边缘相叠",
    "第3个浅蓝色圆形内出现深色服务器造型（由形状组合或图片实现）：自上而下3个等宽横向机架（矩形或圆角矩形，单个高0.3–0.6cm），每个机架内左侧有白色小圆形指示灯、右侧有白色横线条，整体宽1.5–2.5cm",
    "第3个浅蓝色圆形内服务器的右下方叠放橙色实心三角形警示标志（由形状或图片实现）：宽高均为1.0–1.8cm，三角形内有白色感叹号",
    "第4个浅蓝色圆形内出现云朵造型（由形状或图片实现）：深色描边、白色填充，宽1.6–2.6cm、高1.0–1.8cm",
    "第4个浅蓝色圆形内云朵的右下方叠放绿色实心圆形（近似#18934E–#20A45B，由形状或图片实现）：直径0.8–1.6cm，圆内为一个两笔相接、连续完整的白色对勾（✓），不呈两段互不相接的斜线",
    "第5个浅蓝色圆形内出现深色显示器造型（由形状组合或图片实现）：由屏幕外框矩形与其下方的立柱及横向底座构成，整体宽1.6–2.6cm、高1.2–2.2cm",
    "第5个浅蓝色圆形内屏幕区域中央叠放绿色实心圆形（由形状或图片实现）：直径0.8–1.6cm，圆内为一个两笔相接、连续完整的白色对勾（✓），不呈两段互不相接的斜线",
    "页面左上角出现图例框：圆角矩形，无填充或白色填充、深色实线边线（线宽1–3磅），宽5.0–7.0cm、高1.8–2.8cm，左边线位于页面从左至右0%–8%、上边线位于从上至下0%–8%",
    "图例框内上行为深色实线箭头（长2.0–3.0cm、右端带箭头），其右侧为深色加粗文本“action”，字号12–18磅",
    "图例框内下行为1条橙色虚线箭头（单条线长2.0–3.0cm、右端带箭头、线型为虚线或点线，非由多段短线段拼接），其右侧为橙色加粗文本“emotion”，字号12–18磅",
    "页面下部出现1条橙色（近似#FC5E13–#FF4B10）虚线折线：自页面从左至右0%–6%横向贯通至92%–100%，整体位于页面从上至下72%–90%，右端带箭头；走势为先下降后上升、再下降、再上升，波峰位于页面从左至右32%–42%、波谷位于52%–62%",
    "页面下部的橙色虚线折线由连续的线条或连接符以虚线线型实现，单条线段长度均不小于2cm，不存在由30段以上互不相连的独立短线段拼接充当虚线效果的情况，末端箭头由线条端点箭头属性实现",
    "页面下部出现4个橙色倾斜文本，自左至右依次为“concerned”“tense and focused”“hopeful”“relieved”，字号12–18磅，均位于页面从上至下78%–95%，水平位置分别位于页面从左至右3%–13%、26%–36%、64%–74%、84%–95%",
    "页面底部出现倾斜蓝色文本“theme: managing a website outage before a product launch”（近似#0A2E90–#1E5AB4）：字号13–18磅，顶端位于页面从上至下90%–97%，整体水平中心位于页面从左至右45%–55%",
    "页面为16:9横向版式（宽33–34cm、高18.5–19.5cm），全部元素完整位于页面之内，无越出页面边界的元素",
    "同一步骤组内的编号、浅蓝色圆形、阶段标签、说明文本框上下居中对齐，各元素水平中心的相互偏差不超过0.5cm",
    "5个说明文本框的正文字号完全一致，不存在其中某一框字号被单独调小或调大的情况",
    "5个编号圆形内除白色数字外无多余的白色或浅色小圆点、小圆形等形状；放大镜镜片为无填充或透明，其后方柱状图面板的柱条完整可见未被遮挡",
    "页面内所有文字均完整显示在所属形状的可视区域之内，未被相邻形状遮挡、未被裁切；页面内除编号1–5、5个阶段标签文本、5段英文说明文字、图例文字“action”“emotion”、4个情绪标签文本与底部主题文本外，未出现其他文字（如标题、页码、日期、水印、生成说明等）",
    "页面内无覆盖整页的背景色块与整页外边框，所有形状与文本均未设置阴影效果",
]

assert len(RULE_TEXTS) == len(_MAX_DELTAS_ORDER)


# 可能被当作交付件提交的格式。.pptx 之外的这几种也认作交付件，以便维度一按
# "格式不为 .pptx"记 0 分；除此以外的文件（脚本自身、说明文本等）不是交付件。
DELIVERABLE_EXTS = (".pptx", ".ppt", ".pptm", ".ppsx", ".odp", ".pdf",
                    ".png", ".jpg", ".jpeg")


def _locate_document(dir_path):
    """在 dir_path 内定位被评估的交付文件；目录内没有交付件时返回 None。

    优先 .pptx；若目录内只有其他演示或图片格式（.ppt/.pdf/.png 等），也返回该文件，
    以便维度一按"格式不为 .pptx"记 0 分而不是报"没评上"。目录内只有脚本、说明
    文档一类非交付件时返回 None，由 evaluate 按"没评上"处理。
    """
    names = sorted(f for f in os.listdir(dir_path)
                   if not f.startswith("~$") and not f.startswith(".")
                   and os.path.isfile(os.path.join(dir_path, f)))
    pptx = [f for f in names if f.lower().endswith(".pptx")]
    if pptx:
        return os.path.join(dir_path, pptx[0])
    other = [f for f in names
             if os.path.splitext(f)[1].lower() in DELIVERABLE_EXTS]
    if other:
        return os.path.join(dir_path, other[0])
    return None


def evaluate(dir_path):
    """按接口约定返回结构化 dict；不 print、不改 sys.stdout、不 sys.exit。

    入参 dir_path 为"脚本所在目录的路径"，脚本自己在该目录里定位并打开被评估的
    .pptx。

    返回信息约定：
      status="error"（目录不存在、目录中没有交付文件、脚本自身异常）表示这份作业
      根本没评上，此时 dim1_pass 为空串（未判定）、dim1_reason 为空串、
      dim2_items 为 []、total_score 与 max_score 均为 0；
      维度一未通过时 status 仍为 "ok"，dim1_pass 为 False、dim1_reason 写错误摘要、
      dim2_items 为 []（未真正逐条评过）、total_score 为 0、max_score 为满分；
      维度一通过时 dim1_pass 为 True、dim1_reason 保持空串。
      dim2_items 每条只含 rule / max_delta / delta / hit 四个键。
    """
    result = {
        "id": SCRIPT_ID,
        "file_name": "",
        "status": "ok",
        "error": None,
        "dim1_pass": "",
        "dim1_reason": "",
        "dim2_items": [],
        "total_score": 0,
        "max_score": 0,
    }
    try:
        if not dir_path or not os.path.isdir(dir_path):
            raise FileNotFoundError("目录不存在：%s" % dir_path)
        file_path = _locate_document(dir_path)
        if file_path is None:
            # 目录里一个交付物都没有：这份作业没评上，而不是维度一判不通过。
            raise FileNotFoundError("目录内未找到 .pptx 文件：%s" % dir_path)
        result["file_name"] = os.path.basename(file_path)
        return _evaluate_impl(file_path, result)
    except Exception as exc:
        # status 为 error 时未能完成判定，dim1_pass 留空，dim1_reason 也为空，
        # 两项分数一并归零。
        result["status"] = "error"
        result["error"] = "%s: %s" % (type(exc).__name__, exc)
        result["dim1_pass"] = ""
        result["dim1_reason"] = ""
        result["dim2_items"] = []
        result["total_score"] = 0
        result["max_score"] = 0
        return result


def _dim1_fail(result, reason):
    """维度一未通过：完成度总分记 0，不再逐条罗列（未真正逐条评过）。"""
    result["dim1_pass"] = False
    result["dim1_reason"] = reason or "维度一未通过（未给出具体原因）"
    result["dim2_items"] = []
    result["total_score"] = 0
    result["max_score"] = MAX_SCORE_TOTAL
    return result


def _build_dim2_items(results):
    """把 "+N：说明" 逐条反解为结构化 dim2_items（顺序对齐 _MAX_DELTAS_ORDER）。

    返回的每条只含 rule / max_delta / delta / hit 四个键。"""
    if len(results) != len(_MAX_DELTAS_ORDER):
        raise RuntimeError("dim2 项数不一致：results=%d vs order=%d"
                           % (len(results), len(_MAX_DELTAS_ORDER)))
    items = []
    for raw, max_delta in zip(results, _MAX_DELTAS_ORDER):
        head, _, rule = raw.partition("：")
        try:
            delta = int(head.replace("+", "").strip())
        except ValueError:
            delta = 0
        hit = delta > 0
        items.append({
            "rule": rule,
            "max_delta": max_delta,
            "delta": max_delta if hit else 0,
            "hit": hit,
        })
    return items


def _evaluate_impl(file_path, result):
    """维度一门禁 + 维度二 32 条加分项的评估主体。"""
    if not file_path.lower().endswith(SUPPORTED_EXTS):
        return _dim1_fail(result, "交付文件格式不是.pptx（实际为%s）"
                                  % (os.path.splitext(file_path)[1] or "无扩展名"))
    try:
        deck = _load(file_path)
    except Exception as exc:
        return _dim1_fail(result, "交付文件无法正常打开（损坏或需修复）：%s: %s"
                                  % (type(exc).__name__, exc))
    dim1_ok, dim1_reason = _check_dim1(deck, file_path)
    if not dim1_ok:
        return _dim1_fail(result, dim1_reason)
    result["dim1_pass"] = True
    # 约定：维度一通过时 dim1_reason 保持空串，只有未通过才写错误摘要。
    result["dim1_reason"] = ""

    ctx = Ctx(deck)
    flags = []
    flags.extend(_check_layout(ctx))
    flags.extend(_check_icons(ctx))
    flags.extend(_check_legend(ctx))
    flags.extend(_check_curve(ctx))
    flags.extend(_check_norms(ctx))
    if len(flags) != len(_MAX_DELTAS_ORDER):
        raise RuntimeError("检查项数不一致：flags=%d vs order=%d"
                           % (len(flags), len(_MAX_DELTAS_ORDER)))

    results = []
    for hit, delta, text in zip(flags, _MAX_DELTAS_ORDER, RULE_TEXTS):
        results.append("+%d：%s" % (delta if hit else 0, text))

    items = _build_dim2_items(results)
    result["dim2_items"] = items
    result["total_score"] = sum(it["delta"] for it in items)
    result["max_score"] = MAX_SCORE_TOTAL
    return result


if __name__ == "__main__":
    # 本地调试入口：evaluate 只接收"脚本所在目录的路径"，由脚本自己在该目录里
    # 定位并打开被评估的 .pptx。命令行第 1 个参数可覆盖目录路径。
    if len(sys.argv) >= 2:
        _dir = sys.argv[1]
    else:
        _dir = os.path.dirname(os.path.abspath(__file__))
    _out = evaluate(_dir)
    print(json.dumps(_out, ensure_ascii=False, indent=2))
