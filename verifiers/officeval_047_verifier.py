#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# pyright: reportUnknownMemberType=false, reportUnknownVariableType=false, reportUnknownArgumentType=false, reportUnknownParameterType=false, reportMissingParameterType=false, reportPrivateUsage=false, reportUnknownLambdaType=false, reportGeneralTypeIssues=false
"""
PPT 自动评估脚本 —— 一份代码，适配两种不同的制作方式。

两个待评文件由两人以不同方式完成，本脚本对差异做归一化后用同一套规则评估：
  方式 A：节点文字放在圆角矩形“内部”；箭头 = 连接线 + 任意多边形箭头；颜色用 srgbClr。
  方式 B：节点文字用“独立文本框”叠在圆角矩形上；箭头 = 连接线 + 等腰三角形箭头；
          颜色用主题色 schemeClr + lumMod/lumOff。

归一化要点：
  - 递归展开分组(group)，把所有形状换算成幻灯片绝对坐标(cm)。
  - 节点文字：既读矩形内部文字，也把落在矩形内的独立文本框文字并入该节点。
  - 颜色：schemeClr 通过主题色表解析，并套用 lumMod/lumOff 得到近似 RGB，再按色相归类。
  - 箭头：把邻近的“线段 + 箭头头部(三角形/多边形)”聚合成一条逻辑箭头，
          分析其拐点、方向、长度、首尾锚点判断连接关系。

评估逻辑：
  维度一（门槛项）：任一不通过 -> 总分 0，且不再检查维度二。
  维度二（完成度）：逐条评分；加分项需“全部子条件满足”才加分，
                    扣分项“任一子条件满足”即扣分。最后打印命中项与总分。

依赖：python-pptx, lxml  (pip install python-pptx lxml)

接口约定：主结果统一由 evaluate(dir_path) -> dict 返回，入参为"脚本所在目录的
路径"，由脚本自己在该目录内定位并打开被评估的 .pptx（忽略以 ~$ 开头的临时锁
文件）。本模块不 print 主结果、不修改全局 sys.stdout、不 sys.exit；仅在
__main__ 下作为本地调试入口，把返回 dict 序列化为 JSON 打印。
用法：
  python evaluate_ppt.py             # 评估脚本所在目录下的 .pptx，打印 JSON
  python evaluate_ppt.py <目录路径>   # 评估指定目录下的 .pptx，打印 JSON
"""

import os
import sys
import json
import math

from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.oxml.ns import qn

# 脚本编号（与外部评测调度约定一致，可按实际任务号调整）。
SCRIPT_ID = "047"

CM = 360000.0          # 每厘米的 EMU 数
PT = 12700.0           # 每磅的 EMU 数

# 维度二完成度最大可得分：所有加分项满分之和（扣分项最好情形贡献 0）。
MAX_SCORE_TOTAL = 60

# 维度二评分项"结果追加顺序"的 max_delta 列表：与 _evaluate_impl 中 rules
# （build_rules 返回）逐条 append 的顺序严格一致，共 40 项。加分项为其满分，
# 扣分项在最好情形下贡献 0，累加恰为 MAX_SCORE_TOTAL。
_MAX_DELTAS_ORDER = [
    5,  # 01 流程图整体
    1,  # 02 标题文本框
    3,  # 03 Q1 节点
    3,  # 04 Q2 节点
    3,  # 05 Q3 节点
    3,  # 06 Q4 节点
    1,  # 07 标签 Y=1
    3,  # 08 折线箭头 Q4→Q1
    1,  # 09 标签 N=0
    3,  # 10 折线箭头 Q1→Q4
    1,  # 11 标签 X=1
    3,  # 12 折线箭头 Q1→Q2
    1,  # 13 标签 M=0
    1,  # 14 直线箭头 Q2→Q4
    1,  # 15 标签 T↑
    3,  # 16 折线箭头 Q2→Q3
    1,  # 17 标签 R=1
    3,  # 18 折线箭头 Q4→Q3
    1,  # 19 标签 P=0
    3,  # 20 回环箭头 Q3
    1,  # 21 标签 Z=1
    3,  # 22 回环箭头 Q4
    1,  # 23 无"项目原则"标题/正文
    5,  # 24 大标题字号一致（按达标页数计分）
    3,  # 25 目录五条目
    1,  # 26 目录字号一致
    1,  # 27 汇报人
    1,  # 28 邮箱
    0,  # 29 [扣分]第2页目录说明文字非浅色 (-5)
    0,  # 30 [扣分]第12/24页原有内容缺失 (-3)
    0,  # 31 [扣分]第1页缺主标题或指导老师行 (-3)
    0,  # 32 [扣分]第2页目录卡片填充/版式异常 (-3)
    0,  # 33 [扣分]第2页缺说明文字或右上角装饰图片 (-3)
    0,  # 34 [扣分]第3页缺正文 (-3)
    0,  # 35 [扣分]第9页缺正文 (-3)
    0,  # 36 [扣分]第19页导航条异常 (-5)
    0,  # 37 [扣分]第19页 Logo 缺失/移位 (-3)
    0,  # 38 [扣分]第1页 Research presentation 被压 (-1)
    0,  # 39 [扣分]第33页缺汇报人/指导老师行 (-1)
    0,  # 40 [扣分]备注页数或主题配色字体被改 (-1)
]
assert sum(_MAX_DELTAS_ORDER) == MAX_SCORE_TOTAL


# ===========================================================================
# 一、形状归一化：展开分组，换算绝对坐标
# ===========================================================================
class Box:
    """一个形状的归一化视图：绝对坐标(cm) + 文本 + 原始 shape 引用。"""

    def __init__(self, shape, left, top, width, height):
        self.shape = shape
        self.left = left
        self.top = top
        self.width = width
        self.height = height
        self.name = getattr(shape, "name", "")
        self.shape_type = shape.shape_type
        try:
            self.auto = str(shape.auto_shape_type)
        except Exception:
            self.auto = ""
        try:
            self.text = shape.text_frame.text if shape.has_text_frame else ""
        except Exception:
            self.text = ""

    @property
    def cx(self):
        return self.left + self.width / 2.0

    @property
    def cy(self):
        return self.top + self.height / 2.0

    @property
    def right(self):
        return self.left + self.width

    @property
    def bottom(self):
        return self.top + self.height

    @property
    def is_rounded_rect(self):
        return "ROUNDED_RECTANGLE" in self.auto

    @property
    def is_line(self):
        return self.shape_type == MSO_SHAPE_TYPE.LINE

    @property
    def is_picture(self):
        return self.shape_type == MSO_SHAPE_TYPE.PICTURE


def _child_abs(group, child):
    """把分组内子形状的坐标换算为幻灯片绝对坐标(EMU)。"""
    el = group._element
    xfrm = el.find(qn("p:grpSpPr")).find(qn("a:xfrm"))
    off = xfrm.find(qn("a:off"))
    ext = xfrm.find(qn("a:ext"))
    choff = xfrm.find(qn("a:chOff"))
    chext = xfrm.find(qn("a:chExt"))
    gx, gy = int(off.get("x")), int(off.get("y"))
    gcx, gcy = int(ext.get("cx")), int(ext.get("cy"))
    cox, coy = int(choff.get("x")), int(choff.get("y"))
    ccx, ccy = int(chext.get("cx")), int(chext.get("cy"))
    sx = gcx / ccx if ccx else 1.0
    sy = gcy / ccy if ccy else 1.0
    al = gx + (child.left - cox) * sx
    at = gy + (child.top - coy) * sy
    return al, at, child.width * sx, child.height * sy


def flatten(shapes):
    """递归展开所有形状（含分组子形状），返回绝对坐标(cm)的 Box 列表。"""
    out = []

    def rec(shape, ox, oy, sx, sy):
        try:
            al = ox + shape.left * sx
            at = oy + shape.top * sy
            aw = shape.width * sx
            ah = shape.height * sy
        except Exception:
            al = at = aw = ah = 0
        out.append(Box(shape, al / CM, at / CM, aw / CM, ah / CM))
        if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
            el = shape._element
            xfrm = el.find(qn("p:grpSpPr")).find(qn("a:xfrm"))
            g_off = xfrm.find(qn("a:off"))
            g_ext = xfrm.find(qn("a:ext"))
            c_off = xfrm.find(qn("a:chOff"))
            c_ext = xfrm.find(qn("a:chExt"))
            gx, gy = int(g_off.get("x")), int(g_off.get("y"))
            gcx, gcy = int(g_ext.get("cx")), int(g_ext.get("cy"))
            cox, coy = int(c_off.get("x")), int(c_off.get("y"))
            ccx, ccy = int(c_ext.get("cx")), int(c_ext.get("cy"))
            nsx = (gcx / ccx if ccx else 1.0) * sx
            nsy = (gcy / ccy if ccy else 1.0) * sy
            # 子形状原点：先减去子坐标原点，再按缩放映射到组的绝对位置
            base_x = al - cox * nsx
            base_y = at - coy * nsy
            for child in shape.shapes:
                rec(child, base_x, base_y, nsx, nsy)

    for sh in shapes:
        rec(sh, 0.0, 0.0, 1.0, 1.0)
    return out


# ===========================================================================
# 二、颜色解析：主题色 + lumMod/lumOff -> 近似 RGB -> 色相归类
# ===========================================================================
def theme_color_map(prs, slide):
    """读取幻灯片对应主题的配色方案 {name: 'RRGGBB'}。"""
    try:
        from lxml import etree
        master = slide.slide_layout.slide_master
        rel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme"
        tp = master.part.part_related_by(rel)
        root = etree.fromstring(tp.blob)
    except Exception:
        return {}
    ns = {"a": "http://schemas.openxmlformats.org/drawingml/2006/main"}
    scheme = root.find(".//a:clrScheme", ns)
    m = {}
    if scheme is None:
        return m
    for c in scheme:
        tag = c.tag.split("}")[-1]
        child = c[0]
        val = child.get("val") or child.get("lastClr")
        m[tag] = val
    m.setdefault("tx1", m.get("dk1"))
    m.setdefault("bg1", m.get("lt1"))
    m.setdefault("tx2", m.get("dk2"))
    m.setdefault("bg2", m.get("lt2"))
    # windowText/window 等系统色兜底
    if m.get("dk1") in (None, "windowText"):
        m["dk1"] = "000000"; m["tx1"] = "000000"
    if m.get("lt1") in (None, "window"):
        m["lt1"] = "FFFFFF"; m["bg1"] = "FFFFFF"
    return m


def _apply_lum(rgb, lummod, lumoff):
    """在 RGB 上近似套用 lumMod/lumOff（PowerPoint 的常见近似做法）。"""
    r = int(rgb[0:2], 16)
    g = int(rgb[2:4], 16)
    b = int(rgb[4:6], 16)

    def f(c):
        c = c / 255.0
        if lummod is not None:
            c *= lummod
        if lumoff is not None:
            c += lumoff
        return max(0, min(255, round(c * 255)))

    return "%02X%02X%02X" % (f(r), f(g), f(b))


def resolve_clr_element(clr, tmap):
    """把 <a:srgbClr>/<a:schemeClr> 元素解析成 'RRGGBB'（含 lum 调整）。"""
    if clr is None:
        return None
    tag = clr.tag.split("}")[-1]
    if tag == "srgbClr":
        base = clr.get("val")
    elif tag == "schemeClr":
        base = tmap.get(clr.get("val"))
    elif tag == "sysClr":
        base = clr.get("lastClr") or clr.get("val")
    else:
        base = None
    if base is None or len(base) != 6:
        return None
    lm = clr.find(qn("a:lumMod"))
    lo = clr.find(qn("a:lumOff"))
    lummod = int(lm.get("val")) / 100000.0 if lm is not None else None
    lumoff = int(lo.get("val")) / 100000.0 if lo is not None else None
    if lummod is not None or lumoff is not None:
        base = _apply_lum(base, lummod, lumoff)
    return base


def _first_child_clr(parent):
    """在给定 xml 父节点下找到第一个颜色元素。"""
    if parent is None:
        return None
    for tag in ("srgbClr", "schemeClr", "sysClr"):
        e = parent.find(qn("a:" + tag))
        if e is not None:
            return e
    return None


def shape_fill_rgb(box, tmap):
    """解析形状填充色 -> 'RRGGBB'。"""
    spPr = box.shape._element.find(qn("p:spPr"))
    if spPr is None:
        return None
    sf = spPr.find(qn("a:solidFill"))
    return resolve_clr_element(_first_child_clr(sf), tmap)


def shape_line_info(box, tmap):
    """解析形状边线 -> (线宽pt, 'RRGGBB')。"""
    spPr = box.shape._element.find(qn("p:spPr"))
    if spPr is None:
        return None, None
    ln = spPr.find(qn("a:ln"))
    if ln is None:
        return None, None
    w = ln.get("w")
    wpt = int(w) / PT if w else None
    return wpt, resolve_clr_element(_first_child_clr(ln.find(qn("a:solidFill"))), tmap)


def rgb_to_hsv(rgb):
    r = int(rgb[0:2], 16) / 255.0
    g = int(rgb[2:4], 16) / 255.0
    b = int(rgb[4:6], 16) / 255.0
    mx, mn = max(r, g, b), min(r, g, b)
    v = mx
    d = mx - mn
    s = 0 if mx == 0 else d / mx
    if d == 0:
        h = 0
    elif mx == r:
        h = (60 * ((g - b) / d) + 360) % 360
    elif mx == g:
        h = (60 * ((b - r) / d) + 120) % 360
    else:
        h = (60 * ((r - g) / d) + 240) % 360
    return h, s, v


def hue_family(rgb):
    """把 RGB 归类为色系名，用于跨两种配色方式统一判断。

    浅色系(浅蓝/浅黄/浅粉/浅绿)填充饱和度本就很低，故用较低阈值，
    只有近乎纯灰(饱和度极低)才判为 gray，其余按主色相归类。
    """
    if rgb is None:
        return "unknown"
    h, s, v = rgb_to_hsv(rgb)
    if s < 0.03:
        return "gray"
    if h < 15 or h >= 345:
        return "red"
    if h < 45:
        return "orange"
    if h < 70:
        return "yellow"
    if h < 165:
        return "green"
    if h < 200:
        return "cyan"
    if h < 255:
        return "blue"
    if h < 290:
        return "purple"
    if h < 345:
        return "pink"       # 品红/紫红/粉
    return "unknown"


# 每个节点“填充浅色系 / 边线深色系”的可接受色族集合（兼容两种做法）
NODE_COLOR_RULES = {
    "Q1": {"fill": {"blue", "cyan", "gray"}, "line": {"blue", "cyan"}},       # 浅蓝 / 深蓝
    "Q2": {"fill": {"yellow", "orange"}, "line": {"orange", "red"}},          # 浅黄 / 橙红
    "Q3": {"fill": {"pink", "purple", "red"}, "line": {"pink", "purple", "red"}},  # 浅粉 / 紫红
    "Q4": {"fill": {"green"}, "line": {"green"}},                             # 浅绿 / 墨绿
}


# ===========================================================================
# 三、图元语义层：节点、标签、箭头（对两种做法做归一化）
# ===========================================================================
def _norm(s):
    """去掉空白/换行并统一冒号，便于做子串匹配。

    两份文件里“主背景：/线框：”既有全角“：”也有半角“:”，统一去掉冒号后再匹配，
    这样节点文字判定不受书写习惯影响。
    """
    return "".join((s or "").split()).replace("：", "").replace(":", "")


def collect_nodes(boxes):
    """在流程图区域找出 4 个圆角矩形节点，并把落在框内的独立文本框文字并入。

    方式 A：文字在独立文本框 -> 需要并入；方式 B：文字已在框内。
    返回 {'Q1': node_box, ...}，node_box 附带属性 merged_text。
    """
    rects = [b for b in boxes if b.is_rounded_rect and b.top > 2.4]
    textboxes = [b for b in boxes
                 if b.shape_type == MSO_SHAPE_TYPE.TEXT_BOX and b.top > 2.4]
    for r in rects:
        merged = _norm(r.text)
        for tb in textboxes:
            if r.left - 0.2 <= tb.cx <= r.right + 0.2 and \
               r.top - 0.2 <= tb.cy <= r.bottom + 0.2:
                merged += _norm(tb.text)
        r.merged_text = merged
    nodes = {}
    for key in ("Q1", "Q2", "Q3", "Q4"):
        for r in rects:
            if key in getattr(r, "merged_text", ""):
                nodes[key] = r
                break
    return nodes


def find_label(boxes, text):
    """查找文本恰为 text（归一化后相等或以其为主）的独立标签文本框，返回 Box 或 None。"""
    target = _norm(text)
    cands = []
    for b in boxes:
        if b.shape_type != MSO_SHAPE_TYPE.TEXT_BOX:
            continue
        t = _norm(b.text)
        if t == target or (target in t and len(t) <= len(target) + 2):
            cands.append(b)
    return cands[0] if cands else None


# --- 箭头：把线段聚合成逻辑箭头 -------------------------------------------
def _seg_endpoints(box):
    """返回连接线的两端点(cm)，考虑 flipH/flipV。"""
    el = box.shape._element
    xfrm = el.find(".//" + qn("a:xfrm"))
    fH = xfrm is not None and xfrm.get("flipH") == "1"
    fV = xfrm is not None and xfrm.get("flipV") == "1"
    x0, y0, x1, y1 = box.left, box.top, box.right, box.bottom
    p0 = [x0, y0]
    p1 = [x1, y1]
    if fH:
        p0[0], p1[0] = x1, x0
    if fV:
        p0[1], p1[1] = y1, y0
    return (p0[0], p0[1]), (p1[0], p1[1])


def _is_head(box):
    """箭头头部：任意多边形(方式A) 或 等腰三角形(方式B)，且尺寸较小。"""
    if box.top <= 2.4:
        return False
    small = box.width < 1.2 and box.height < 1.2
    if box.shape_type == MSO_SHAPE_TYPE.FREEFORM and small:
        return True
    if "TRIANGLE" in box.auto and small:
        return True
    return False


class Arrow:
    """一条逻辑箭头（由若干首尾相接的线段 + 头部聚合而成）。"""

    def __init__(self, segs, heads):
        self.segs = segs                       # [((x,y),(x,y)), ...]
        self.heads = heads                     # [(cx,cy), ...]
        xs, ys = [], []
        for a, b in segs:
            xs += [a[0], b[0]]
            ys += [a[1], b[1]]
        self.minx, self.maxx = min(xs), max(xs)
        self.miny, self.maxy = min(ys), max(ys)
        self.span_x = self.maxx - self.minx
        self.span_y = self.maxy - self.miny
        # 各线段方向与长度
        self.hsegs = []   # 水平段 (x0,x1,y,len)
        self.vsegs = []   # 竖直段 (y0,y1,x,len)
        for a, b in segs:
            dx = abs(a[0] - b[0])
            dy = abs(a[1] - b[1])
            if dx >= dy:
                self.hsegs.append((a[0], b[0], (a[1] + b[1]) / 2, dx))
            else:
                self.vsegs.append((a[1], b[1], (a[0] + b[0]) / 2, dy))
        self.max_h = max([s[3] for s in self.hsegs], default=0.0)
        self.max_v = max([s[3] for s in self.vsegs], default=0.0)

    def free_ends(self):
        """出现次数为 1 的端点（箭头的“头/尾”自由端）。"""
        from collections import defaultdict
        cnt = defaultdict(int)
        pts = []
        for a, b in self.segs:
            for p in (a, b):
                key = (round(p[0], 1), round(p[1], 1))
                cnt[key] += 1
                pts.append((key, p))
        ends = []
        seen = set()
        for key, p in pts:
            if cnt[key] == 1 and key not in seen:
                ends.append(p)
                seen.add(key)
        return ends

    def head_end(self):
        """离任一头部最近的自由端（视为箭头终点）。"""
        if not self.heads:
            return None
        best, bd = None, 1e9
        for e in self.free_ends():
            for h in self.heads:
                d = math.hypot(e[0] - h[0], e[1] - h[1])
                if d < bd:
                    bd, best = d, e
        return best

    def tail_end(self):
        """远离头部的自由端（视为箭头起点）。"""
        he = self.head_end()
        ends = self.free_ends()
        if he is None or not ends:
            return ends[0] if ends else None
        best, bd = None, -1
        for e in ends:
            d = math.hypot(e[0] - he[0], e[1] - he[1])
            if d > bd:
                bd, best = d, e
        return best


def build_arrows(boxes, tol=0.6):
    """把流程图内的连接线按端点邻近聚合为逻辑箭头，并挂接头部。"""
    lines = [b for b in boxes if b.shape_type == MSO_SHAPE_TYPE.LINE and b.top > 2.4]
    heads = [(b.cx, b.cy) for b in boxes if _is_head(b)]
    segs = [_seg_endpoints(b) for b in lines]
    n = len(segs)
    par = list(range(n))

    def find(x):
        while par[x] != x:
            par[x] = par[par[x]]
            x = par[x]
        return x

    def near(p, q):
        return abs(p[0] - q[0]) < tol and abs(p[1] - q[1]) < tol

    for i in range(n):
        for j in range(i + 1, n):
            linked = any(near(a, b) for a in segs[i] for b in segs[j])
            if linked:
                par[find(i)] = find(j)

    from collections import defaultdict
    comp = defaultdict(list)
    for i in range(n):
        comp[find(i)].append(i)

    arrows = []
    for idxs in comp.values():
        csegs = [segs[i] for i in idxs]
        # 收集落在该组附近的头部
        chs = []
        for h in heads:
            for a, b in csegs:
                if (abs(h[0] - a[0]) < 0.8 and abs(h[1] - a[1]) < 0.8) or \
                   (abs(h[0] - b[0]) < 0.8 and abs(h[1] - b[1]) < 0.8):
                    chs.append(h)
                    break
        arrows.append(Arrow(csegs, chs))
    return arrows


def _near_edge(pt, node, edge, tol=0.9):
    """判断点 pt 是否位于 node 的指定边(top/bottom/left/right)附近。"""
    x, y = pt
    if node is None:
        return False
    if edge == "top":
        return abs(y - node.top) < tol and node.left - tol <= x <= node.right + tol
    if edge == "bottom":
        return abs(y - node.bottom) < tol and node.left - tol <= x <= node.right + tol
    if edge == "left":
        return abs(x - node.left) < tol and node.top - tol <= y <= node.bottom + tol
    if edge == "right":
        return abs(x - node.right) < tol and node.top - tol <= y <= node.bottom + tol
    return False


# ===========================================================================
# 四、文字 / 字体 辅助
# ===========================================================================
def iter_runs(shape):
    """产出 (text, size_pt, font_name) —— 遍历形状内所有 run。"""
    if not getattr(shape, "has_text_frame", False):
        return
    for para in shape.text_frame.paragraphs:
        for run in para.runs:
            sz = run.font.size.pt if run.font.size is not None else None
            yield run.text, sz, run.font.name


def biggest_title(slide):
    """返回该页“最大字号”的一段文字信息 (text, size_pt, font_name)。"""
    best = None
    for sh in slide.shapes:
        for text, sz, name in iter_runs(sh):
            if not text.strip():
                continue
            if sz is None:
                continue
            if best is None or sz > best[1]:
                best = (text, sz, name)
    return best


def slide_text(slide):
    """把一页（含分组）所有文字拼接为一个规范化串。"""
    boxes = flatten(slide.shapes)
    return "".join(_norm(b.text) for b in boxes if b.text)


def slide_has_text(slide, needle):
    return _norm(needle) in slide_text(slide)


def count_pictures(slide):
    return sum(1 for b in flatten(slide.shapes) if b.is_picture)


def is_yahei(name):
    """字体是否为微软雅黑或黑体（None 视为不确定 -> 交由主题字体兜底判断）。"""
    if not name:
        return False
    return ("雅黑" in name) or ("YaHei" in name) or ("黑体" in name) \
        or ("SimHei" in name) or ("Microsoft YaHei" in name)


def theme_font_names(prs, slide):
    """读取主题字体方案的中英文字体名 (major_latin, major_ea, minor_latin, minor_ea)。

    run 未显式设置字体时会继承主题字体，因此需要这个兜底。
    """
    try:
        from lxml import etree
        master = slide.slide_layout.slide_master
        rel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme"
        tp = master.part.part_related_by(rel)
        root = etree.fromstring(tp.blob)
    except Exception:
        return []
    ns = {"a": "http://schemas.openxmlformats.org/drawingml/2006/main"}
    fs = root.find(".//a:fontScheme", ns)
    if fs is None:
        return []
    names = []
    for kind in ("a:majorFont", "a:minorFont"):
        fe = fs.find(kind, ns)
        if fe is None:
            continue
        for sub in ("a:latin", "a:ea"):
            e = fe.find(sub, ns)
            if e is not None and e.get("typeface"):
                names.append(e.get("typeface"))
    return names


def font_ok(name, theme_names):
    """判断字体是否微软雅黑/黑体：显式字体优先，未设置则用主题字体兜底。"""
    if name:
        return is_yahei(name)
    # 未显式设置 -> 继承主题字体，只要主题任一默认字体是雅黑/黑体即算达标
    return any(is_yahei(n) for n in (theme_names or []))


# ===========================================================================
# 五、维度一：门槛检查项
# ===========================================================================
def check_gate(prs):
    """返回 (passed:bool, details:list[(desc, ok)])。任一 False -> 总分 0。"""
    details = []

    # 1) .pptx 可正常打开：能构造 Presentation 即视为通过
    details.append((".pptx 格式且可正常打开", True))

    # 2) 共 33 页
    n = len(prs.slides)
    details.append(("共 33 页幻灯片（实际 %d 页）" % n, n == 33))

    # 3) 页面尺寸 33.87cm × 19.05cm
    w = prs.slide_width / CM
    h = prs.slide_height / CM
    size_ok = abs(w - 33.87) < 0.2 and abs(h - 19.05) < 0.2
    details.append(("页面尺寸 33.87×19.05cm（实际 %.2f×%.2f）" % (w, h), size_ok))

    passed = all(ok for _, ok in details)
    return passed, details


# ===========================================================================
# 六、维度二：完成度评分细则
# ===========================================================================
class Rule:
    """一条评分细则。points 可正(加分)可负(扣分)。"""

    def __init__(self, points, desc, fn):
        self.points = points
        self.desc = desc
        self.fn = fn      # fn(ctx) -> bool（是否命中该细则）


class Ctx:
    """一次评估的上下文缓存，避免重复解析。"""

    def __init__(self, prs):
        self.prs = prs
        self.slides = list(prs.slides)
        self.s19 = self.slides[18] if len(self.slides) > 18 else None
        self.boxes19 = flatten(self.s19.shapes) if self.s19 else []
        self.tmap = theme_color_map(prs, self.s19) if self.s19 else {}
        self.theme_fonts = theme_font_names(prs, self.s19) if self.s19 else []
        self.nodes = collect_nodes(self.boxes19) if self.s19 else {}
        self.arrows = build_arrows(self.boxes19) if self.s19 else []

    def arrow_between(self, tail_key, tail_edges, head_key, head_edges):
        """找一条 tail 端锚在 tail_key 的某条边、head 端锚在 head_key 的某条边的箭头。"""
        for ar in self.arrows:
            te, he = ar.tail_end(), ar.head_end()
            if not te or not he:
                continue
            t_ok = any(_near_edge(te, self.nodes.get(tail_key), e) for e in tail_edges)
            h_ok = any(_near_edge(he, self.nodes.get(head_key), e) for e in head_edges)
            if t_ok and h_ok:
                return ar
            # 头尾方向也可能被识别反（头部检测失败时），做一次对称兜底
            t_ok2 = any(_near_edge(he, self.nodes.get(tail_key), e) for e in tail_edges)
            h_ok2 = any(_near_edge(te, self.nodes.get(head_key), e) for e in head_edges)
            if t_ok2 and h_ok2:
                return ar
        return None


def _leg_at(arrow, pt):
    """返回离 pt 最近的那条“有效”线段的 (方向'H'/'V', 长度cm)。

    箭头端部常有 <0.4cm 的极短线段（连接箭头图形的小接头），会误判为末段方向，
    因此优先忽略这些极短线段，只在没有更长线段可选时才回退使用它们。
    """
    def _pick(min_len):
        best, bd = None, 1e9
        for a, b in arrow.segs:
            seg_len = math.hypot(a[0] - b[0], a[1] - b[1])
            if seg_len < min_len:
                continue
            for e in (a, b):
                d = math.hypot(e[0] - pt[0], e[1] - pt[1])
                if d < bd:
                    bd, best = d, (a, b)
        return best

    best = _pick(0.4) or _pick(0.0)
    if best is None:
        return None, 0.0
    a, b = best
    dx, dy = abs(a[0] - b[0]), abs(a[1] - b[1])
    return ("V" if dy > dx else "H"), max(dx, dy)


def _in(v, lo, hi):
    return lo <= v <= hi


# --- 节点细则通用实现 ------------------------------------------------------
def _node_rule(ctx, key, need_texts, lo_L, hi_L, lo_T, hi_T,
               fill_fams, line_fams):
    """圆角矩形节点综合检查：文字、位置、尺寸、线宽、填充/边线色系。"""
    nd = ctx.nodes.get(key)
    if nd is None:
        return False
    merged = getattr(nd, "merged_text", _norm(nd.text))
    if not all(_norm(t) in merged for t in need_texts):
        return False
    if not (_in(nd.left, lo_L, hi_L) and _in(nd.top, lo_T, hi_T)):
        return False
    if not (_in(nd.width, 3.0, 5.0) and _in(nd.height, 2.3, 3.7)):
        return False
    lw, lc = shape_line_info(nd, ctx.tmap)
    if lw is None or not _in(lw, 2.0, 3.0):
        return False
    fill = shape_fill_rgb(nd, ctx.tmap)
    if hue_family(fill) not in fill_fams:
        return False
    if hue_family(lc) not in line_fams:
        return False
    return True


# --- 标签细则通用实现 ------------------------------------------------------
def _label_rule(ctx, text, lo_L, hi_L, lo_T, hi_T):
    b = find_label(ctx.boxes19, text)
    if b is None:
        return False
    return _in(b.left, lo_L, hi_L) and _in(b.top, lo_T, hi_T)


# --- 折线箭头细则通用实现 --------------------------------------------------
def _elbow_rule(ctx, tail_key, tail_edges, head_key, head_edges,
                first_dir, first_lo, first_hi,
                last_dir, last_lo, last_hi):
    """折线箭头：首段方向/长度 + 末段方向/长度 + 首尾锚点。"""
    ar = ctx.arrow_between(tail_key, tail_edges, head_key, head_edges)
    if ar is None:
        return False
    te, he = ar.tail_end(), ar.head_end()
    # 保证 te 锚在 tail、he 锚在 head（arrow_between 已保证其一，规范化方向）
    if not any(_near_edge(te, ctx.nodes.get(tail_key), e) for e in tail_edges):
        te, he = he, te
    fd, fl = _leg_at(ar, te)
    ld, ll = _leg_at(ar, he)
    if fd != first_dir or not _in(fl, first_lo, first_hi):
        return False
    if ld != last_dir or not _in(ll, last_lo, last_hi):
        return False
    return True


def _loop_rule(ctx, key, side):
    """回环箭头：起终点都在同一节点的同一侧，短促伸出后折返，主干范围较小。

    side='bottom'：位于节点正下方，向下引出；side='left'：位于节点左侧，向左引出。
    """
    nd = ctx.nodes.get(key)
    if nd is None:
        return False
    for ar in ctx.arrows:
        te, he = ar.tail_end(), ar.head_end()
        if not te or not he:
            # 退化情形：用整体范围判断
            ends = ar.free_ends()
            pts = ends if ends else [p for s in ar.segs for p in s]
        else:
            pts = [te, he]
        # 两端是否都贴着该节点（回环会向外伸出，容差需覆盖伸出距离）
        near_node = all(
            (nd.left - 2.6 <= p[0] <= nd.right + 2.6 and
             nd.top - 2.6 <= p[1] <= nd.bottom + 2.6)
            for p in pts)
        if not near_node:
            continue
        # 整体回环尺寸不超过约 3.5cm
        if ar.span_x > 3.5 or ar.span_y > 3.5:
            continue
        # 位于期望的一侧
        cx = (ar.minx + ar.maxx) / 2.0
        cy = (ar.miny + ar.maxy) / 2.0
        if side == "bottom" and cy < nd.bottom - 0.3:
            continue
        if side == "left" and cx > nd.left + 0.5:
            continue
        # 短促伸出（0.3–2.5cm）
        protr = ar.max_v if side == "bottom" else ar.max_h
        if not _in(protr, 0.3, 2.5):
            continue
        return True
    return False


# --- 大标题字号一致性（+5，按达标页数计分）--------------------------------
def _big_title_pages_ok(ctx):
    """第 9/15/22/28/33 页大标题字号 50–56、字体微软雅黑；返回达标页数。"""
    targets = [(8, ""), (14, ""), (21, ""), (27, ""), (32, "")]
    ok = 0
    for idx, _ in targets:
        if idx >= len(ctx.slides):
            continue
        info = biggest_title(ctx.slides[idx])
        if info is None:
            continue
        _, sz, name = info
        if sz is not None and _in(sz, 50, 56) and font_ok(name, ctx.theme_fonts):
            ok += 1
    return ok


# --- 目录条目（+3）--------------------------------------------------------
def _toc_ok(ctx):
    """第 2 页存在 -01-~-05- 五个编号，且各自对应正确标题，按编号顺序排列。"""
    if len(ctx.slides) < 2:
        return False
    boxes = flatten(ctx.slides[1].shapes)
    nums = ["-01-", "-02-", "-03-", "-04-", "-05-"]
    titles = ["研究背景", "研究思路", "项目进展", "问题讨论", "研究结论"]
    # 收集编号框及其 x 坐标
    num_pos = {}
    for b in boxes:
        t = _norm(b.text)
        for nseq in nums:
            if _norm(nseq) == t or _norm(nseq) in t and len(t) <= 6:
                num_pos.setdefault(nseq, b.left)
    if any(n not in num_pos for n in nums):
        return False
    # 按 x 排序应与 01..05 顺序一致
    order = sorted(nums, key=lambda n: num_pos[n])
    if order != nums:
        return False
    # 标题存在
    alltext = "".join(_norm(b.text) for b in boxes)
    if not all(_norm(t) in alltext for t in titles):
        return False
    return True


def _toc_font_ok(ctx):
    """第 2 页 -05- 编号 26–28 磅、研究结论标题 22–24 磅，且与其余条目一致。"""
    if len(ctx.slides) < 2:
        return False
    num_sizes = []
    title_sizes = []
    for sh in ctx.slides[1].shapes:
        for text, sz, name in iter_runs(sh):
            t = _norm(text)
            if t in ("-01-", "-02-", "-03-", "-04-", "-05-") and sz:
                num_sizes.append(sz)
            if t in ("研究背景", "研究思路", "项目进展", "问题讨论", "研究结论") and sz:
                title_sizes.append(sz)
    if len(num_sizes) < 5 or len(title_sizes) < 5:
        return False
    num_ok = all(_in(s, 26, 28) for s in num_sizes) and (max(num_sizes) - min(num_sizes) < 0.6)
    title_ok = all(_in(s, 22, 24) for s in title_sizes) and (max(title_sizes) - min(title_sizes) < 0.6)
    return num_ok and title_ok


# --- 第2页“研究结论”条目下方说明文字（+1）--------------------------------
# 注：该加分项已按需求移除。


# --- 流程图整体（+5）------------------------------------------------------
def _flowchart_overall(ctx):
    """存在由可编辑对象构成的流程图：4 个圆角矩形节点 + 多条箭头，
    环状(上下左右)排列，整体位于指定范围，且水平居中。"""
    if len(ctx.nodes) < 4:
        return False
    nds = ctx.nodes
    # 4 节点整体范围
    lefts = [n.left for n in nds.values()]
    tops = [n.top for n in nds.values()]
    rights = [n.right for n in nds.values()]
    bottoms = [n.bottom for n in nds.values()]
    if min(lefts) < 2 or max(rights) > 31:
        return False
    if min(tops) < 2 or max(bottoms) > 19:
        return False
    # 至少 3 条箭头
    real_arrows = [a for a in ctx.arrows if a.tail_end() and a.head_end()]
    if len(real_arrows) < 3:
        return False
    # 环状：上下左右——Q1 在上、Q3 在下、且左右两侧各一个(Q2/Q4)
    q1, q2, q3, q4 = nds.get("Q1"), nds.get("Q2"), nds.get("Q3"), nds.get("Q4")
    if not all([q1, q2, q3, q4]):
        return False
    ring = (q1.cy < q3.cy and                       # Q1 在 Q3 上方
            min(q2.cx, q4.cx) < q1.cx < max(q2.cx, q4.cx) and  # 左右夹住上节点
            min(q2.cy, q4.cy) < q3.cy)               # 左右在下节点上方附近
    if not ring:
        return False
    # 水平居中：所有节点整体的横向中点落在 11–22cm
    mid_x = (min(lefts) + max(rights)) / 2.0
    if not _in(mid_x, 11, 22):
        return False
    return True


# --- 标题文本框（+1）------------------------------------------------------
def _title_box_ok(ctx):
    """第 19 页标题“设备指示状态切换流程图”文本框：位置/尺寸/字体/字号。"""
    for b in ctx.boxes19:
        if b.shape_type != MSO_SHAPE_TYPE.TEXT_BOX:
            continue
        if _norm(b.text) != _norm("设备指示状态切换流程图"):
            continue
        if not (_in(b.left, 7, 25) and _in(b.top, 1, 5)):
            return False
        if not (_in(b.width, 4, 18) and _in(b.height, 0.5, 2)):
            return False
        # 字体/字号
        for text, sz, name in iter_runs(b.shape):
            if not text.strip():
                continue
            if sz is None or not _in(sz, 16, 28):
                return False
            if not font_ok(name, ctx.theme_fonts):
                return False
        return True
    return False


# ===========================================================================
# 六之二、扣分项判定辅助（只判定“惹麻烦”的破坏性改动，主元素与加分项不重叠）
# ===========================================================================
# 汇报人/指导老师行首的符号图标（Wingdings 私用区字符）
_ICON = ""

# “浅色文字”的亮度下限：白色 1.00、浅灰 E7E6E6 约 0.91、白色减淡 A6A6A6 约 0.65
# 都算浅色；黑色 0.00、深蓝 143852 约 0.19 不算。原文件导航标签即 A6A6A6，
# 因此阈值取 0.55，既容得下白色减淡，又能判出黑字/深蓝字压深底的情形。
_LIGHT_MIN = 0.55


def _luma(rgb):
    """近似相对亮度 0–1，用于区分“浅色文字/深色文字”。"""
    if not rgb or len(rgb) != 6:
        return None
    r = int(rgb[0:2], 16) / 255.0
    g = int(rgb[2:4], 16) / 255.0
    b = int(rgb[4:6], 16) / 255.0
    return 0.299 * r + 0.587 * g + 0.114 * b


def _close(rgb, target, tol=24):
    """两个 'RRGGBB' 是否逐通道接近（容差 tol）。"""
    if not rgb or not target or len(rgb) != 6 or len(target) != 6:
        return False
    for i in (0, 2, 4):
        if abs(int(rgb[i:i + 2], 16) - int(target[i:i + 2], 16)) > tol:
            return False
    return True


def _run_color(run, tmap):
    """解析一个 run 的显式文字颜色 -> 'RRGGBB'；未显式设置返回 None。

    未显式设置时会继承演示文稿默认文字样式（本模板解析结果为黑色），因此
    None 在“文字必须是浅色”的判定里等同于不达标。
    """
    rPr = run._r.find(qn("a:rPr"))
    if rPr is None:
        return None
    return resolve_clr_element(_first_child_clr(rPr.find(qn("a:solidFill"))), tmap)


def _shape_runs(box):
    """取一个 Box 内的所有非空 run。"""
    sh = box.shape
    if not getattr(sh, "has_text_frame", False):
        return []
    return [r for p in sh.text_frame.paragraphs for r in p.runs if r.text.strip()]


def _find_text_box(slide, needle):
    """在一页里找到文字含 needle 的第一个形状（含分组内）。"""
    for b in flatten(slide.shapes):
        if b.text and _norm(needle) in _norm(b.text):
            return b
    return None


def _grad_fill_has(box, target, tol=24):
    """形状是否为渐变填充且渐变色含接近 target 的颜色。"""
    spPr = box.shape._element.find(qn("p:spPr"))
    if spPr is None:
        return False
    gf = spPr.find(qn("a:gradFill"))
    if gf is None:
        return False
    for c in gf.iter(qn("a:srgbClr")):
        if _close(c.get("val"), target, tol):
            return True
    return False


# --- -5 第2页目录卡片内说明文字非浅色 --------------------------------------
_TOC_DESC_SAMPLES = [
    "完成样本整理、模型迭代与初步验证，形成可复用流程。",
    "聚焦技术难点、落地边界和后续优化方向。",
]


def _toc_desc_color_bad(ctx):
    """抽检两条说明文字：颜色须为白色或浅灰 E7E6E6 等浅色。"""
    if len(ctx.slides) < 2:
        return True
    s2 = ctx.slides[1]
    tmap = theme_color_map(ctx.prs, s2)
    for needle in _TOC_DESC_SAMPLES:
        b = _find_text_box(s2, needle)
        if b is None:
            return True
        runs = _shape_runs(b)
        if not runs:
            return True
        for r in runs:
            lum = _luma(_run_color(r, tmap))
            if lum is None or lum < _LIGHT_MIN:
                return True
    return False


# --- -3 第12/24页原有内容被破坏 --------------------------------------------
def _anchor_pages_bad(ctx):
    if len(ctx.slides) < 24:
        return True
    s12, s24 = ctx.slides[11], ctx.slides[23]
    if not slide_has_text(s12, "闭环"):
        return True
    if not slide_has_text(s12, "需求拆解：明确设备、工况、维护节点和预警目标。"):
        return True
    if not slide_has_text(s24, "风险一"):
        return True
    if not slide_has_text(s24, "数据缺口会影响模型对早期异常的识别能力，需要建立缺失提示。"):
        return True
    return False


# --- -3 第1页主标题 + 指导老师行 -------------------------------------------
def _p1_title_teacher_bad(ctx):
    s1 = ctx.slides[0]
    b = _find_text_box(s1, "设备状态监测与预测维护研究")
    if b is None:
        return True
    runs = _shape_runs(b)
    szs = [r.font.size.pt for r in runs if r.font.size is not None]
    if not szs or not _in(max(szs), 50, 56):
        return True
    if not any(font_ok(r.font.name, ctx.theme_fonts) for r in runs):
        return True

    b = _find_text_box(s1, "指导老师：周沐辞")
    if b is None:
        return True
    runs = _shape_runs(b)
    if _ICON not in "".join(r.text for r in runs):      # 行首符号图标
        return True
    szs = [r.font.size.pt for r in runs if r.font.size is not None]
    if not szs or not _in(max(szs), 12, 14):
        return True
    if not any(font_ok(r.font.name, ctx.theme_fonts) for r in runs):
        return True
    return False


# --- -3 第2页目录卡片填充与版式 --------------------------------------------
def _toc_cards_bad(ctx):
    """卡片须为深蓝渐变 082290；各卡片高度/纵向起点/宽度须彼此接近。"""
    if len(ctx.slides) < 2:
        return True
    cards = [b for b in flatten(ctx.slides[1].shapes)
             if not b.is_picture and not b.is_line
             and _in(b.width, 4, 9) and _in(b.height, 6, 10) and _in(b.top, 5, 9)]
    if not cards:
        return True
    for b in cards:
        if not _grad_fill_has(b, "082290"):
            return True
    if len(cards) >= 2:
        hs = [b.height for b in cards]
        ts = [b.top for b in cards]
        ws = [b.width for b in cards]
        if max(hs) - min(hs) > 0.3 or max(ts) - min(ts) > 0.3 or max(ws) - min(ws) > 0.5:
            return True
    return False


# --- -3 第2页说明文字与右上角装饰图片 --------------------------------------
def _p2_desc_pic_bad(ctx):
    if len(ctx.slides) < 2:
        return True
    s2 = ctx.slides[1]
    if not slide_has_text(s2, "围绕设备运行数据、故障预警需求与应用场景展开。"):
        return True
    for b in flatten(s2.shapes):
        if b.is_picture and b.left >= 24 and b.top <= 5:
            return False
    return True


# --- -3 正文文字（第3页、第9页）通用判定 -----------------------------------
def _body_text_bad(ctx, idx, needle, lo=11, hi=15):
    """指定页须出现该段正文，且其文字号在 lo–hi 磅（13 磅上下）。"""
    if idx >= len(ctx.slides):
        return True
    b = _find_text_box(ctx.slides[idx], needle)
    if b is None:
        return True
    szs = [r.font.size.pt for r in _shape_runs(b) if r.font.size is not None]
    if szs and not _in(max(szs), lo, hi):
        return True
    return False


# --- -5 第19页顶部导航条 ---------------------------------------------------
_NAV_LABELS = ["研究背景", "研究思路", "项目进展", "问题讨论", "研究结论"]


def _nav_bad(ctx):
    """五个标签齐全；“项目进展”深蓝 071769、其余四个浅色；导航条顶部通栏。"""
    if ctx.s19 is None:
        return True
    tmap = ctx.tmap or theme_color_map(ctx.prs, ctx.s19)
    got = {}
    for b in ctx.boxes19:
        if not b.text or not _in(b.top, -0.6, 2.5):
            continue
        t = _norm(b.text)
        for lb in _NAV_LABELS:
            if _norm(lb) in t and lb not in got:
                got[lb] = b
    if len(got) < 5:
        return True
    for lb, b in got.items():
        cols = [_run_color(r, tmap) for r in _shape_runs(b)]
        if lb == "项目进展":
            if not any(_close(c, "071769", 30) for c in cols):
                return True
        else:
            if not cols:
                return True
            for c in cols:
                lum = _luma(c)
                if lum is None or lum < _LIGHT_MIN:
                    return True
    # 导航条本体：距上 0–1 厘米起、覆盖页面整宽、高约 2.2 厘米的通栏
    for b in ctx.boxes19:
        if b.width >= 30 and _in(b.height, 1.5, 3.2) and _in(b.top, -0.6, 1.0):
            return False
    return True


# --- -3 第19页左上角 Logo --------------------------------------------------
def _logo_bad(ctx):
    """Logo 图片须落在距左 1.76–2.76、距上 0–0.99，且约 3.0×1.9 厘米。"""
    for b in ctx.boxes19:
        if b.is_picture and _in(b.left, 1.76, 2.76) and _in(b.top, 0, 0.99) \
           and _in(b.width, 2.6, 3.4) and _in(b.height, 1.5, 2.3):
            return False
    return True


# --- -1 第1页 Research presentation 被其他文字横向压住 ---------------------
def _p1_rp_covered(ctx):
    boxes = [b for b in flatten(ctx.slides[0].shapes) if b.text and b.text.strip()]
    tgt = None
    for b in boxes:
        if "research presentation" in b.text.strip().lower() and _in(b.top, 11.5, 14):
            tgt = b
            break
    if tgt is None:
        return False        # 该文字本身不在本条判定范围
    for b in boxes:
        if b is tgt:
            continue
        if abs(b.cy - tgt.cy) > 0.6:          # 不在同一行
            continue
        if min(b.right, tgt.right) - max(b.left, tgt.left) >= 0.5:
            return True
    return False


# --- -1 第33页汇报人/指导老师行 --------------------------------------------
def _p33_bad(ctx):
    if len(ctx.slides) < 33:
        return True
    s33 = ctx.slides[32]
    b = _find_text_box(s33, "汇报人：林泽宇")
    if b is None:
        return True
    if _ICON not in "".join(r.text for r in _shape_runs(b)):
        return True
    if not slide_has_text(s33, "指导老师：周沐辰"):
        return True
    return False


# --- -1 备注页数量与主题配色/字体 ------------------------------------------
def _package_meta_bad(ctx):
    prs = ctx.prs
    if sum(1 for s in prs.slides if s.has_notes_slide) != 32:
        return True
    tmap = theme_color_map(prs, prs.slides[0])
    if not _close(tmap.get("accent1"), "9B0D14", 12):
        return True
    if not _close(tmap.get("accent2"), "002060", 12):
        return True
    names = theme_font_names(prs, prs.slides[0])   # [大标题西文, 大标题中文, 正文西文, 正文中文]
    if len(names) < 4:
        return True
    if "微软雅黑" not in (names[2] or ""):
        return True
    minor_ea = names[3] or ""
    if "微软雅黑" not in minor_ea or "Light" not in minor_ea:
        return True
    return False


# ===========================================================================
# 七、评分细则表（维度二）
# ===========================================================================
def build_rules():
    R = []

    # +5 流程图整体
    R.append(Rule(5, "第19页存在由可编辑对象构成的流程图：4 个圆角矩形节点 + 多条箭头，"
                     "环状(上下左右)排列，整体位于指定范围且水平居中",
                  _flowchart_overall))

    # +1 标题文本框
    R.append(Rule(1, "第19页存在标题“设备指示状态切换流程图”文本框(位置/尺寸/字体/字号达标)",
                  _title_box_ok))

    # +3 四个节点
    R.append(Rule(3, "Q1节点：含“Q1/00/主背景：浅蓝/线框：深蓝”，位置尺寸线宽达标，浅蓝填充+深蓝边线",
                  lambda c: _node_rule(c, "Q1", ["Q1", "00", "主背景：浅蓝", "线框：深蓝"],
                                       12, 20, 2.5, 9,
                                       NODE_COLOR_RULES["Q1"]["fill"], NODE_COLOR_RULES["Q1"]["line"])))
    R.append(Rule(3, "Q2节点：含“Q2/01/主背景：浅黄/线框：橙红”，位置尺寸线宽达标，浅黄填充+橙红边线",
                  lambda c: _node_rule(c, "Q2", ["Q2", "01", "主背景：浅黄", "线框：橙红"],
                                       17, 25, 7.5, 14,
                                       NODE_COLOR_RULES["Q2"]["fill"], NODE_COLOR_RULES["Q2"]["line"])))
    R.append(Rule(3, "Q3节点：含“Q3/11/主背景：浅粉/线框：紫红”，位置尺寸线宽达标，浅粉填充+紫红边线",
                  lambda c: _node_rule(c, "Q3", ["Q3", "11", "主背景：浅粉", "线框：紫红"],
                                       12, 20, 12.5, 19,
                                       NODE_COLOR_RULES["Q3"]["fill"], NODE_COLOR_RULES["Q3"]["line"])))
    R.append(Rule(3, "Q4节点：含“Q4/10/主背景：浅绿/线框：墨绿”，位置尺寸线宽达标，浅绿填充+墨绿边线",
                  lambda c: _node_rule(c, "Q4", ["Q4", "10", "主背景：浅绿", "线框：墨绿"],
                                       7.5, 15.5, 7.5, 14,
                                       NODE_COLOR_RULES["Q4"]["fill"], NODE_COLOR_RULES["Q4"]["line"])))

    # +1 标签 Y=1
    R.append(Rule(1, "标签“Y=1”文本框(距左6.5–15.5, 距上1.5–8.5)",
                  lambda c: _label_rule(c, "Y=1", 6.5, 15.5, 1.5, 8.5)))
    # +3 折线箭头 Q4顶 -> Q1左
    R.append(Rule(3, "折线箭头：Q4顶边→Q1左边，先竖直上行1.5–5，再水平右行1.5–5",
                  lambda c: _elbow_rule(c, "Q4", ["top"], "Q1", ["left"],
                                        "V", 1.5, 5, "H", 1.5, 5)))
    # +1 标签 N=0
    R.append(Rule(1, "标签“N=0”文本框(距左6–15, 距上4–11)",
                  lambda c: _label_rule(c, "N=0", 6, 15, 4, 11)))
    # +3 折线箭头 Q1 -> Q4顶
    R.append(Rule(3, "折线箭头：Q1→Q4顶边，先水平左行1.5–5，再竖直下行1.5–5",
                  lambda c: _elbow_rule(c, "Q1", ["left", "top", "bottom", "right"], "Q4", ["top"],
                                        "H", 1.5, 5, "V", 1.5, 5)))
    # +1 标签 X=1
    R.append(Rule(1, "标签“X=1”文本框(距左16.5–26, 距上3–11)",
                  lambda c: _label_rule(c, "X=1", 16.5, 26, 3, 11)))
    # +3 折线箭头 Q1右 -> Q2顶
    R.append(Rule(3, "折线箭头：Q1右边→Q2顶边，先水平右行1.5–5，再竖直下行1.5–5",
                  lambda c: _elbow_rule(c, "Q1", ["right"], "Q2", ["top"],
                                        "H", 1.5, 5, "V", 1.5, 5)))
    # +1 标签 M=0
    R.append(Rule(1, "标签“M=0”文本框(距左11–21, 距上7–14)",
                  lambda c: _label_rule(c, "M=0", 11, 21, 7, 14)))
    # +1 直线箭头 Q2左 -> Q4右
    R.append(Rule(1, "直线箭头：Q2左边→Q4右边，单段水平无拐点，长3.5–7",
                  lambda c: _straight_rule(c, "Q2", "Q4", 3.5, 7)))
    # +1 标签 T↑
    R.append(Rule(1, "标签“T↑”文本框(距左15–25, 距上11–18.5)",
                  lambda c: _label_rule(c, "T↑", 15, 25, 11, 18.5)))
    # +3 折线箭头 Q2底 -> Q3右
    R.append(Rule(3, "折线箭头：Q2底边→Q3右边，先竖直下行1.5–5，再水平左行1.5–5",
                  lambda c: _elbow_rule(c, "Q2", ["bottom"], "Q3", ["right"],
                                        "V", 1.5, 5, "H", 1.5, 5)))
    # +1 标签 R=1
    R.append(Rule(1, "标签“R=1”文本框(距左6.5–16, 距上11–18.5)",
                  lambda c: _label_rule(c, "R=1", 6.5, 16, 11, 18.5)))
    # +3 折线箭头 Q4底 -> Q3左
    R.append(Rule(3, "折线箭头：Q4底边→Q3左边，先竖直下行1.5–5，再水平右行1.5–5",
                  lambda c: _elbow_rule(c, "Q4", ["bottom"], "Q3", ["left"],
                                        "V", 1.5, 5, "H", 1.5, 5)))
    # +1 标签 P=0
    R.append(Rule(1, "标签“P=0”文本框(距左10.5–20, 距上14.5–19)",
                  lambda c: _label_rule(c, "P=0", 10.5, 20, 14.5, 19)))
    # +3 回环箭头 Q3 下方
    R.append(Rule(3, "回环箭头：位于Q3正下方，起终点均在Q3底边，向下引出后折返，回环尺寸≤约3.5cm",
                  lambda c: _loop_rule(c, "Q3", "bottom")))
    # +1 标签 Z=1
    R.append(Rule(1, "标签“Z=1”文本框(距左3.5–12, 距上7–14)",
                  lambda c: _label_rule(c, "Z=1", 3.5, 12, 7, 14)))
    # +3 回环箭头 Q4 左侧
    R.append(Rule(3, "回环箭头：位于Q4左侧，起终点均在Q4左边，向左引出后折返，回环尺寸≤约3.5cm",
                  lambda c: _loop_rule(c, "Q4", "left")))

    # +1 不存在“项目原则”标题与相应正文
    R.append(Rule(1, "第19页不存在“项目原则”标题及“项目推进以…为原则：”开头正文",
                  _no_project_principle))

    # +5 大标题字号一致（按达标页数计分，特殊处理）
    R.append(Rule(5, "第9/15/22/28/33页大标题字号50–56且微软雅黑（按达标页数×1，封顶5）",
                  None, ))  # 特殊：见 evaluate 中的按比例计分

    # +3 目录五条目
    R.append(Rule(3, "第2页目录五条目 -01-~-05- 顺序正确且标题对应",
                  _toc_ok))
    # +1 目录字号一致
    R.append(Rule(1, "第2页目录 -05- 编号26–28磅、研究结论标题22–24磅且与其余条目一致",
                  _toc_font_ok))
    # +1 汇报人 陈前康
    R.append(Rule(1, "第1页“汇报人：”之后为“陈前康”",
                  lambda c: _norm("汇报人：陈前康") in slide_text(c.slides[0])))
    # +1 邮箱
    R.append(Rule(1, "第1页汇报人行包含“邮箱:896700205@qq.com”",
                  lambda c: _norm("896700205@qq.com") in slide_text(c.slides[0])))

    # ---- 扣分项（12 条，最多扣 34 分）----
    R.append(Rule(-5, "第2页目录卡片内说明文字颜色不是白色或浅灰E7E6E6等浅色"
                      "(抽检“完成样本整理…”“聚焦技术难点…”两处)",
                  _toc_desc_color_bad))
    R.append(Rule(-3, "第12页缺“闭环”或“需求拆解：…”；第24页缺“风险一”或其段落文字",
                  _anchor_pages_bad))
    R.append(Rule(-3, "第1页缺54磅微软雅黑主标题“设备状态监测与预测维护研究”，"
                      "或缺13磅微软雅黑、行首带符号图标的“指导老师：周沐辞”",
                  _p1_title_teacher_bad))
    R.append(Rule(-3, "第2页目录卡片填充不是深蓝渐变082290，或五个卡片高度/纵向起点/宽度不齐",
                  _toc_cards_bad))
    R.append(Rule(-3, "第2页缺说明文字“围绕设备运行数据…”，或缺右上角装饰图片",
                  _p2_desc_pic_bad))
    R.append(Rule(-3, "第3页缺13磅正文“围绕多源设备数据治理…”",
                  lambda c: _body_text_bad(c, 2,
                      "围绕多源设备数据治理、异常识别与预测维护需求，说明研究的现实基础与应用价值。")))
    R.append(Rule(-3, "第9页缺13磅正文“采用分阶段推进方式…”",
                  lambda c: _body_text_bad(c, 8,
                      "采用分阶段推进方式，将需求分析、数据治理、模型构建和验证反馈整合为一条可落地路径。")))
    R.append(Rule(-5, "第19页缺顶部导航条或五个标签，或“项目进展”非深蓝071769、其余四个非白色，"
                      "或导航条不在距上0–1起的顶部通栏位置",
                  _nav_bad))
    R.append(Rule(-3, "第19页缺Logo图片，或其不在距左1.76–2.76、距上0–0.99，或不是约3.0×1.9厘米",
                  _logo_bad))
    R.append(Rule(-1, "第1页“Research presentation”被其他文字横向重叠遮挡",
                  _p1_rp_covered))
    R.append(Rule(-1, "第33页缺“汇报人：林泽宇”(行首符号图标)或“指导老师：周沐辰”",
                  _p33_bad))
    R.append(Rule(-1, "备注页不是32个，或主题配色第一/第二强调色、主题正文中英文字体被改",
                  _package_meta_bad))

    return R


def _straight_rule(ctx, tail_key, head_key, lo_len, hi_len):
    """直线箭头：起终点分别锚在两节点左右边，单段水平、无拐点、长度达标。"""
    ar = ctx.arrow_between(tail_key, ["left", "right"], head_key, ["left", "right"])
    if ar is None:
        return False
    # 单段水平：只有一条主要水平段、竖直跨度很小
    if ar.span_y > 0.6:
        return False
    if len(ar.hsegs) < 1:
        return False
    if not _in(ar.max_h, lo_len, hi_len):
        return False
    return True


def _no_project_principle(ctx):
    """第19页不存在“项目原则”标题，也不存在以指定文字开头的正文段落。"""
    text = slide_text(ctx.s19) if ctx.s19 else ""
    if _norm("项目原则") in text:
        return False
    if _norm("项目推进以") in text and _norm("可验证、可解释、可维护") in text:
        return False
    return True


# ===========================================================================
# 八、维度一/维度二评估主体 + 结果封装
# ===========================================================================
def _open_presentation(file_path):
    """打开 .pptx 演示文稿（python-pptx 仅支持 OOXML/.pptx）。"""
    return Presentation(file_path)


def evaluate(dir_path):
    """按接口统一约定返回结构化 dict；不 print、不改 sys.stdout、不 sys.exit。

    入参 dir_path 为"脚本所在目录的路径"，脚本自己在该目录内定位并打开被评估的
    .pptx（忽略以 ~$ 开头的 Office 临时锁文件）。目录不存在、未发现文档等异常
    由本函数捕获并返回 status='error'。"""
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
            raise FileNotFoundError(f"目录不存在：{dir_path}")
        cands = sorted(
            f for f in os.listdir(dir_path)
            if f.lower().endswith(".pptx") and not f.startswith("~$")
        )
        if not cands:
            raise FileNotFoundError(f"目录内未找到 .pptx 文件：{dir_path}")
        file_path = os.path.join(dir_path, cands[0])
        result["file_name"] = os.path.basename(file_path)
        return _evaluate_impl(file_path, result)
    except Exception as exc:
        # status 为 error 时未能完成判定，dim1_pass 留空，dim1_reason 也为空
        result["status"] = "error"
        result["error"] = f"{type(exc).__name__}: {exc}"
        result["dim1_pass"] = ""
        result["dim1_reason"] = ""
        result["dim2_items"] = []
        result["total_score"] = 0
        return result

def _evaluate_impl(file_path, result):
    """维度一（门槛）+ 维度二（完成度）评估主体：沿用原打分逻辑，末尾按
    _MAX_DELTAS_ORDER 拆解为结构化 dim2_items 填入 result。"""
    # ========== 维度1：可用与可修改性（门槛） ==========
    # 任一门槛项未过 -> 总分 0，且不再检查维度二（与原逻辑一致）。
    try:
        prs = _open_presentation(file_path)
    except Exception as exc:
        result["dim1_pass"] = False
        result["dim1_reason"] = "无法打开文件：%s" % exc
        result["dim2_items"] = []
        result["total_score"] = 0
        result["max_score"] = MAX_SCORE_TOTAL
        return result

    passed, details = check_gate(prs)
    if not passed:
        result["dim1_pass"] = False
        result["dim1_reason"] = "；".join(desc for desc, ok in details if not ok)
        result["dim2_items"] = []
        result["total_score"] = 0
        result["max_score"] = MAX_SCORE_TOTAL
        return result
    result["dim1_pass"] = True
    result["dim1_reason"] = ""

    # ========== 维度2：完成度评分 ==========
    # 沿用 build_rules() 的细则与 Ctx 上下文；逐条计算"实际得分 delta"：
    #   - 加分项：命中得满分 points，否则 0
    #   - 扣分项：触发得 points（负值），否则 0
    #   - 大标题特殊项：按达标页数比例得分（0–5）
    ctx = Ctx(prs)
    rules = build_rules()
    results = []
    # 每条细则的"分值上限/下限"：加分项为其满分(正)，扣分项为其扣分(负)，
    # 用作 dim2_items 中 max_delta 的取值——扣分项即展示为该扣分值(如 -3)。
    max_deltas = []
    for r in rules:
        # +5 大标题按比例计分的特殊项（fn 为 None）
        if r.points == 5 and r.fn is None:
            ok_pages = _big_title_pages_ok(ctx)
            gained = min(round(ok_pages / 5.0 * 5.0, 2), 5.0)
            results.append("%+d：%s" % (int(gained), r.desc))
            max_deltas.append(r.points)
            continue
        try:
            hit = bool(r.fn(ctx))
        except Exception:
            hit = False
        earned = r.points if hit else 0
        results.append("%+d：%s" % (earned, r.desc))
        max_deltas.append(r.points)
    # ========== 将 results 拆解为结构化 dim2_items ==========
    # 逐项对齐 _MAX_DELTAS_ORDER，解析每条 "+N：说明" 的 N 为实际 delta。
    # 注：因本细则含扣分项与比例计分项，delta/total 采用"实际得分"而非满分，
    #     hit 以 delta != 0 判定（触发扣分亦记为 hit）。
    if len(results) != len(_MAX_DELTAS_ORDER):
        raise RuntimeError(
            f"dim2 项数不一致：results={len(results)} vs order={len(_MAX_DELTAS_ORDER)}"
        )
    dim2_items = []
    total_score = 0
    for raw, max_delta in zip(results, max_deltas):
        head, _, rule = raw.partition("：")
        try:
            delta = int(head.replace("+", "").strip())
        except ValueError:
            delta = 0
        total_score += delta
        if max_delta and delta not in (0, max_delta):
            # 比例计分项允许部分得分，但统一结果契约要求每个条目的 delta
            # 只能是 0 或 max_delta，因此拆成 1 分（或 -1 分）的原子计分单元。
            step = 1 if max_delta > 0 else -1
            units = abs(max_delta)
            gained = min(abs(delta), units)
            for unit in range(units):
                hit_unit = unit < gained
                dim2_items.append({
                    "rule": "%s（计分单元 %d/%d）" % (rule, unit + 1, units),
                    "max_delta": step,
                    "delta": step if hit_unit else 0,
                    "hit": hit_unit,
                })
            continue
        dim2_items.append({
            "rule": rule,
            "max_delta": max_delta,
            "delta": delta,
            "hit": delta != 0,
        })

    result["dim2_items"] = dim2_items
    result["total_score"] = total_score
    result["max_score"] = MAX_SCORE_TOTAL
    return result


if __name__ == "__main__":
    # 本地调试入口：evaluate 只接收"脚本所在目录的路径"，由脚本自己在该目录内
    # 定位并打开被评估的 .pptx。命令行第 1 个参数可覆盖该目录，未指定时使用脚本
    # 自身所在目录。
    if len(sys.argv) >= 2:
        _dir = sys.argv[1]
    else:
        _dir = os.path.dirname(os.path.abspath(__file__))
    _out = evaluate(_dir)
    print(json.dumps(_out, ensure_ascii=False, indent=2))
