#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# pyright: reportUnknownMemberType=false, reportUnknownVariableType=false
# pyright: reportUnknownArgumentType=false, reportMissingImports=false
# pyright: reportOptionalMemberAccess=false, reportGeneralTypeIssues=false
"""
officeval_056_verifier —— “客户投诉闭环处理流程”流程图 PPT 完成度评分器

评分口径完全依照同目录《细则.txt》：
    维度1  交付文件为1个可正常打开的 .pptx（门禁，不通过则总分直接记 0）
    维度2  加分项共 24 条、合计 92 分，与《细则.txt》维度2 的 24 行按序一一对应
           （细则第 4-27 行 ←→ dim2_items[0..23]，条目合并沿革见 _MAX_DELTAS_ORDER 上方注释）

返回结构与 officeval_008_verifier 完全一致：
    evaluate(dir_path) -> {
        id, file_name, status, error,
        dim1_pass, dim1_reason,
        dim2_items[{rule, max_delta, delta, hit}],
        total_score, max_score
    }
字段口径：
    status="ok"  且 dim1_pass=True  -> error=None、dim1_reason=""，
                                       dim2_items 为 24 条、max_score=92
    status="ok"  且 dim1_pass=False -> error=None、dim1_reason 必为不通过原因摘要
                                       （维度一为 False 时该字段不得为空），
                                       dim2_items=[]、total_score=0、max_score=92
    status="error"（脚本自身异常，含 -> error 为错误摘要、dim1_pass=""（未判定，
    目录不存在、目录中未找到 .pptx）   既非 True 也非 False）、dim1_reason=""、
                                       dim2_items=[]、total_score=0、max_score=0

本模块只用 zipfile + lxml 直接读取 OPC 包与 DrawingML，不依赖 python-pptx；
模块内不打印结果、不改动 sys.stdout、不调用 sys.exit。
"""

import json
import math
import os
import sys
import zipfile

from lxml import etree

SCRIPT_ID = "056"
MAX_SCORE_TOTAL = 92

# 与下方 results.append / emit 的调用顺序严格一一对应，且与《细则.txt》维度2 的 24 行按序对齐
# （细则第 4-27 行 ←→ dim2_items[0..23]，各条 max_delta 即细则该行行首的 +N）。
# 沿革：细则最初写为 30 条，其中度量同一属性、判定必然同步的 6 组已两两合并为 6 条，
# 总分仍为 92（合并不改变任何一条的判定口径，只是把两组判据并入同一条）；细则原文亦已同步合并：
#   原L14 描边规格统一 + 原L28 三段配色分区      -> 形状描边规格与三段配色分区统一（1+3=4）
#   原L15 顶端对齐高度一致 + 原L17 纵向定位      -> 两排纵向对齐与定位准确（3+1=4）
#   原L19 "是"分支竖直 + 原L20 两条回流折线      -> 分支与回流走线正确（3+5=8）
#   原L23 图标齐备 + 原L24 图标规格             -> 环节图标齐备与规格（5+3=8）
#   原L29 图例内容 + 原L31 圆形底衬与标题色      -> 图例块存在与内容（3+1=4）
#   原L30 图例尺寸位置 + 原L32 图例文字         -> 图例块尺寸与排版得当（3+1=4）
_MAX_DELTAS_ORDER = [
    5, 3, 1, 1, 3, 3,      # 幅面 / 居中 / 洁净 / 主标题 / 标题字号 / 装饰带
    5, 5, 3, 5,            # 上排顺序 / 下排顺序 / 形状类型 / 尺寸
    4, 4, 3,               # 描边+配色 / 纵向对齐+定位 / 间隙均匀
    5, 8, 3, 3,            # 11箭头 / 分支与回流走线 / 标注齐备 / 标注字号
    8, 1, 5, 3,            # 图标齐备+规格 / 菱形徽标 / 文字不溢出 / 文字字号
    4, 4, 3,               # 图例存在与内容 / 图例尺寸与排版 / 互不压盖
]
assert sum(_MAX_DELTAS_ORDER) == MAX_SCORE_TOTAL

A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
P = 'http://schemas.openxmlformats.org/presentationml/2006/main'
R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
NS = {'a': A, 'p': P, 'r': R}

EMU_CM = 360000.0           # 1cm = 360000 EMU
PT_CM = 2.54 / 72.0         # 1磅 = 0.03528cm

# PowerPoint 文本框默认内边距：左右 91440EMU=0.254cm，上下 45720EMU=0.127cm
DEF_LR_INS = 91440 / EMU_CM
DEF_TB_INS = 45720 / EMU_CM
# 全角汉字前进宽度（相对字号的比例）。中文字库标称 1em，实际正文/艺术字普遍带负side-bearing，
# 取 0.85 作为估算系数；行高按字号 1.2 倍计。文本范围仅用于“是否溢出/是否压盖/是否居中”的
# 量级判断，故所有相关判定都另给 0.25-0.45cm 容差。
CJK_ADV = 0.85
LINE_H = 1.2

UPPER_TEXTS = ('接收投诉', '信息登记', '问题分级', '是否重大问题？',
               '客服处理', '客户确认', '结案归档')
LOWER_TEXTS = ('指派专员', '原因分析', '制定整改方案', '执行整改',
               '质量复核', '复核通过？')
NODE_TEXTS = UPPER_TEXTS + LOWER_TEXTS
DECISION_TEXTS = ('是否重大问题？', '复核通过？')
TITLE_TEXT = '客户投诉闭环处理流程'
LEGEND_TITLE = '相关记录：'
LEGEND_ITEMS = ('沟通记录', '证据材料', '处理报告')
BULLET_CHARS = '•●▪·◦‣∙◆·-–'
ROUND_GEOMS = ('roundRect', 'round1Rect', 'round2SameRect', 'round2DiagRect')


# ---------------------------------------------------------------------------
# 通用小工具
# ---------------------------------------------------------------------------
def _loc(el, name):
    """取直接子元素（忽略命名空间前缀差异）"""
    if el is None:
        return []
    return el.xpath('./*[local-name()="%s"]' % name)


def _deep(el, name):
    if el is None:
        return []
    return el.xpath('.//*[local-name()="%s"]' % name)


def _norm(s):
    """文本归一化：去掉所有空白/换行，半角问号统一为全角，便于与细则字面比对"""
    if not s:
        return ''
    out = []
    for ch in s:
        if ch.isspace():
            continue
        if ch == '?':
            ch = '？'
        elif ch == ':':
            ch = '：'
        out.append(ch)
    return ''.join(out)


def _rng(vals):
    """极差"""
    vals = [v for v in vals if v is not None]
    if not vals:
        return 0.0
    return max(vals) - min(vals)


def _ov1(a0, a1, b0, b1):
    """一维重叠长度（负数表示相离）"""
    return min(a1, b1) - max(a0, b0)


def _fmt(v):
    return ('%.2f' % v) if isinstance(v, float) else str(v)


# ---------------------------------------------------------------------------
# 颜色：主题解析 + 色系判定
# 与办公软件对齐：PowerPoint 里 schemeClr 显示为主题色，取 theme1.xml 的 clrScheme
# 映射为具体 RGB；lumMod/shade 等微调不影响“蓝/青绿/橙”的色相归类，故忽略。
# ---------------------------------------------------------------------------
def _theme_map(pkg):
    out = {}
    names = [n for n in pkg.namelist() if n.startswith('ppt/theme/') and n.endswith('.xml')]
    for n in sorted(names):
        try:
            root = etree.fromstring(pkg.read(n))
        except Exception:
            continue
        sc = root.find('.//a:clrScheme', NS)
        if sc is None:
            continue
        for ch in sc:
            tag = etree.QName(ch).localname
            s = ch.find('a:srgbClr', NS)
            y = ch.find('a:sysClr', NS)
            if s is not None:
                out[tag] = s.get('val')
            elif y is not None:
                out[tag] = y.get('lastClr')
        break
    out.setdefault('dk1', '000000')
    out.setdefault('lt1', 'FFFFFF')
    out.setdefault('tx1', out.get('dk1'))
    out.setdefault('bg1', out.get('lt1'))
    out.setdefault('tx2', out.get('dk2', '44546A'))
    out.setdefault('bg2', out.get('lt2', 'E7E6E6'))
    return out


def _color_of(container, tm):
    """从 a:ln / a:spPr / a:rPr 等容器里取 solidFill 颜色，返回 6 位 hex / 'NONE' / None"""
    if container is None:
        return None
    if _loc(container, 'noFill'):
        return 'NONE'
    sf = _loc(container, 'solidFill')
    if not sf:
        return None
    f = sf[0]
    s = _loc(f, 'srgbClr')
    if s:
        return (s[0].get('val') or '').upper()
    sc = _loc(f, 'schemeClr')
    if sc:
        v = sc[0].get('val')
        if v in ('phClr',):
            return None
        return (tm.get(v) or '').upper() or None
    return None


def _rgb(hexs):
    if not hexs or hexs == 'NONE' or len(hexs) != 6:
        return None
    try:
        return (int(hexs[0:2], 16), int(hexs[2:4], 16), int(hexs[4:6], 16))
    except ValueError:
        return None


def _hsl(hexs):
    """返回 (色相0-360, 饱和度0-1, 明度0-1)；无法解析返回 None"""
    v = _rgb(hexs)
    if v is None:
        return None
    r, g, b = (c / 255.0 for c in v)
    mx, mn = max(r, g, b), min(r, g, b)
    l = (mx + mn) / 2.0
    d = mx - mn
    if d < 1e-9:
        return (0.0, 0.0, l)
    s = d / (2.0 - mx - mn) if l > 0.5 else d / (mx + mn)
    if mx == r:
        h = 60.0 * (((g - b) / d) % 6.0)
    elif mx == g:
        h = 60.0 * (2.0 + (b - r) / d)
    else:
        h = 60.0 * (4.0 + (r - g) / d)
    return (h % 360.0, s, l)


def _is_blue(hexs):
    h = _hsl(hexs)
    return bool(h) and 195.0 <= h[0] <= 255.0 and h[1] >= 0.25


def _is_dark_blue(hexs):
    h = _hsl(hexs)
    return _is_blue(hexs) and h[2] < 0.36


def _is_light_blue(hexs):
    h = _hsl(hexs)
    return bool(h) and 188.0 <= h[0] <= 220.0 and h[1] >= 0.30 and h[2] >= 0.36


def _is_teal(hexs):
    h = _hsl(hexs)
    return bool(h) and 155.0 <= h[0] < 195.0 and h[1] >= 0.20


def _is_orange(hexs):
    h = _hsl(hexs)
    return bool(h) and 8.0 <= h[0] <= 48.0 and h[1] >= 0.35


def _light(hexs, thr=0.80):
    h = _hsl(hexs)
    return bool(h) and h[2] >= thr


def _dark(hexs, thr=0.62):
    h = _hsl(hexs)
    return h is None or h[2] <= thr


def _cdist(h1, h2):
    a, b = _rgb(h1), _rgb(h2)
    if a is None or b is None:
        return 0.0
    return math.sqrt(sum((a[i] - b[i]) ** 2 for i in range(3)))


# ---------------------------------------------------------------------------
# 形状记录：所有坐标已展平到幻灯片坐标系并换算为 cm
# ---------------------------------------------------------------------------
class Shape(object):
    __slots__ = ('name', 'kind', 'geom', 'txbox', 'rot', 'flipH', 'flipV',
                 'L', 'T', 'W', 'H', 'vL', 'vT', 'vW', 'vH',
                 'lnw', 'dash', 'lnc', 'tail', 'head', 'fill',
                 'text', 'runs', 'algns', 'anchor', 'wrap', 'af', 'ins')

    @property
    def R(self):
        return self.L + self.W

    @property
    def B(self):
        return self.T + self.H

    @property
    def cx(self):
        return self.L + self.W / 2.0

    @property
    def cy(self):
        return self.T + self.H / 2.0

    @property
    def vR(self):
        return self.vL + self.vW

    @property
    def vB(self):
        return self.vT + self.vH

    @property
    def vcx(self):
        return self.vL + self.vW / 2.0

    @property
    def vcy(self):
        return self.vT + self.vH / 2.0

    @property
    def area(self):
        return self.W * self.H

    @property
    def size(self):
        """最大字号（磅）"""
        v = [r[1] for r in self.runs if r[1]]
        return max(v) if v else None

    @property
    def font(self):
        for r in self.runs:
            if r[2] or r[3]:
                return r[2] or r[3]
        return None

    @property
    def color(self):
        for r in self.runs:
            if r[4]:
                return r[4]
        return None

    def is_diamond(self):
        return self.geom == 'diamond'

    def __repr__(self):
        return '<%s %s %.2f,%.2f %.2fx%.2f>' % (self.kind, self.name, self.L, self.T, self.W, self.H)


def _vis_box(L, T, W, H, rot):
    """旋转后的可视外接框（与办公软件中肉眼看到的选择框一致）"""
    if not rot:
        return (L, T, W, H)
    a = math.radians(rot)
    nw = abs(W * math.cos(a)) + abs(H * math.sin(a))
    nh = abs(W * math.sin(a)) + abs(H * math.cos(a))
    return (L + W / 2.0 - nw / 2.0, T + H / 2.0 - nh / 2.0, nw, nh)


def _in_shape(sh, x, y, tol=0.0):
    """点是否落在形状的可视区域内；菱形按真实棱形判定，其余按外接框"""
    if sh.is_diamond():
        hw, hh = sh.W / 2.0 + tol, sh.H / 2.0 + tol
        if hw <= 0 or hh <= 0:
            return False
        return abs(x - sh.cx) / hw + abs(y - sh.cy) / hh <= 1.0
    return (sh.vL - tol <= x <= sh.vR + tol) and (sh.vT - tol <= y <= sh.vB + tol)


def _xspan(sh, y0, y1):
    """形状在 [y0,y1] 纵向区间内可能达到的最大水平范围。
    菱形上下窄、中间宽，用外接框判压盖会误判，这里按棱形的真实轮廓取该区间内最宽处。"""
    if not sh.is_diamond():
        return (sh.vL, sh.vR)
    if y1 < sh.T or y0 > sh.B:
        return (sh.cx, sh.cx)
    ny = min(max(sh.cy, y0), y1)          # 区间内距中心最近的 y
    hh = sh.H / 2.0
    k = 1.0 - (abs(ny - sh.cy) / hh if hh > 0 else 1.0)
    k = max(0.0, min(1.0, k))
    hw = sh.W / 2.0 * k
    return (sh.cx - hw, sh.cx + hw)


def _overlap(a, b):
    """两形状可视范围的重叠量 (ox, oy)；菱形按棱形轮廓修正水平范围"""
    ay0, ay1 = a.vT, a.vB
    by0, by1 = b.vT, b.vB
    oy = _ov1(ay0, ay1, by0, by1)
    y0, y1 = max(ay0, by0), min(ay1, by1)
    if oy <= 0:
        al, ar = a.vL, a.vR
        bl, br = b.vL, b.vR
    else:
        al, ar = _xspan(a, y0, y1)
        bl, br = _xspan(b, y0, y1)
    return (_ov1(al, ar, bl, br), oy)


# ---------------------------------------------------------------------------
# OPC / DrawingML 解析
# ---------------------------------------------------------------------------
def _cnv(el):
    q = ('./*[local-name()="nvSpPr" or local-name()="nvPicPr" or '
         'local-name()="nvCxnSpPr" or local-name()="nvGraphicFramePr" or '
         'local-name()="nvGrpSpPr"]/*[local-name()="cNvPr"]')
    r = el.xpath(q)
    return r[0] if r else None


def _read_text(txBody, tm):
    """返回 (整体文本, runs, 段落对齐列表)"""
    runs = []
    algns = []
    segs = []
    for p in _loc(txBody, 'p'):
        pieces = []
        pPr = _loc(p, 'pPr')
        algns.append(pPr[0].get('algn') if pPr else None)
        for node in p:
            nm = etree.QName(node).localname
            if nm == 'r':
                t = ''.join(node.itertext())
                pieces.append(t)
                rPr = _loc(node, 'rPr')
                sz = la = ea = col = None
                bold = False
                if rPr:
                    rp = rPr[0]
                    s = rp.get('sz')
                    sz = int(s) / 100.0 if s else None
                    bold = rp.get('b') == '1'
                    l1 = _loc(rp, 'latin')
                    e1 = _loc(rp, 'ea')
                    la = l1[0].get('typeface') if l1 else None
                    ea = e1[0].get('typeface') if e1 else None
                    col = _color_of(rp, tm)
                runs.append((t, sz, la, ea, col, bold))
            elif nm == 'br':
                pieces.append('\n')
            elif nm == 'fld':
                pieces.append(''.join(node.itertext()))
        segs.append(''.join(pieces))
    return ('\n'.join(segs), runs, algns)


def _read_shape(el, tag, tm, ox, oy, sx, sy):
    sh = Shape()
    c = _cnv(el)
    sh.name = (c.get('name') if c is not None else '') or ''
    sh.kind = {'sp': 'sp', 'pic': 'pic', 'cxnSp': 'cxn', 'graphicFrame': 'gfx'}.get(tag, 'other')
    cnvsp = el.xpath('./*[local-name()="nvSpPr"]/*[local-name()="cNvSpPr"]')
    sh.txbox = bool(cnvsp) and cnvsp[0].get('txBox') == '1'

    sp = _loc(el, 'spPr')
    sp = sp[0] if sp else None
    xf = _loc(sp, 'xfrm') if sp is not None else []
    if not xf:
        xf = _loc(el, 'xfrm')
    L = T = W = H = 0
    sh.rot = 0.0
    sh.flipH = sh.flipV = False
    if xf:
        x = xf[0]
        sh.rot = float(x.get('rot') or 0) / 60000.0
        sh.flipH = x.get('flipH') == '1'
        sh.flipV = x.get('flipV') == '1'
        off = _loc(x, 'off')
        ext = _loc(x, 'ext')
        if off:
            L, T = int(off[0].get('x') or 0), int(off[0].get('y') or 0)
        if ext:
            W, H = int(ext[0].get('cx') or 0), int(ext[0].get('cy') or 0)
    sh.L = (L * sx + ox) / EMU_CM
    sh.T = (T * sy + oy) / EMU_CM
    sh.W = W * sx / EMU_CM
    sh.H = H * sy / EMU_CM
    sh.vL, sh.vT, sh.vW, sh.vH = _vis_box(sh.L, sh.T, sh.W, sh.H, sh.rot)

    sh.geom = None
    if sp is not None:
        pg = _loc(sp, 'prstGeom')
        if pg:
            sh.geom = pg[0].get('prst')
        elif _loc(sp, 'custGeom'):
            sh.geom = 'custom'

    # 描边与填充；未显式写 ln/fill 时回落到 p:style 的 lnRef/fillRef（与软件取值一致）
    sh.lnw = sh.dash = sh.lnc = sh.tail = sh.head = None
    sh.fill = None
    if sp is not None:
        ln = _loc(sp, 'ln')
        if ln:
            l0 = ln[0]
            w = l0.get('w')
            sh.lnw = round(int(w) / 12700.0, 2) if w else None
            d = _loc(l0, 'prstDash')
            sh.dash = d[0].get('val') if d else None
            sh.lnc = _color_of(l0, tm)
            t = _loc(l0, 'tailEnd')
            sh.tail = t[0].get('type') if t else None
            h = _loc(l0, 'headEnd')
            sh.head = h[0].get('type') if h else None
        sh.fill = _color_of(sp, tm)
    st = _loc(el, 'style')
    if st:
        if sh.lnc is None:
            sh.lnc = _color_of(_loc(st[0], 'lnRef')[0] if _loc(st[0], 'lnRef') else None, tm)
        if sh.fill is None:
            sh.fill = _color_of(_loc(st[0], 'fillRef')[0] if _loc(st[0], 'fillRef') else None, tm)

    # 文本
    sh.text, sh.runs, sh.algns = '', [], []
    sh.anchor = sh.wrap = sh.af = None
    sh.ins = (DEF_LR_INS, DEF_LR_INS, DEF_TB_INS, DEF_TB_INS)
    tb = _loc(el, 'txBody')
    if tb:
        sh.text, sh.runs, sh.algns = _read_text(tb[0], tm)
        bp = _loc(tb[0], 'bodyPr')
        if bp:
            b = bp[0]
            sh.anchor = b.get('anchor')
            sh.wrap = b.get('wrap')
            sh.af = ('spAf' if _loc(b, 'spAutoFit')
                     else ('norm' if _loc(b, 'normAutofit') else None))

            def _i(k, dv):
                v = b.get(k)
                return int(v) / EMU_CM if v is not None else dv
            sh.ins = (_i('lIns', DEF_LR_INS), _i('rIns', DEF_LR_INS),
                      _i('tIns', DEF_TB_INS), _i('bIns', DEF_TB_INS))
    return sh


def _walk(parent, tm, out, ox=0.0, oy=0.0, sx=1.0, sy=1.0, groups=None):
    """递归展平 spTree；组合按 off/ext/chOff/chExt 做坐标换算，等价于软件里进组后看到的位置"""
    for el in parent:
        tag = etree.QName(el).localname
        if tag == 'grpSp':
            gsp = _loc(el, 'grpSpPr')
            gsp = gsp[0] if gsp else None
            xf = _loc(gsp, 'xfrm') if gsp is not None else []
            if not xf:
                _walk(el, tm, out, ox, oy, sx, sy, groups)
                continue
            x = xf[0]
            off, ext = _loc(x, 'off'), _loc(x, 'ext')
            cho, che = _loc(x, 'chOff'), _loc(x, 'chExt')
            gx = int(off[0].get('x') or 0) if off else 0
            gy = int(off[0].get('y') or 0) if off else 0
            gcx = int(ext[0].get('cx') or 0) if ext else 0
            gcy = int(ext[0].get('cy') or 0) if ext else 0
            cx0 = int(cho[0].get('x') or 0) if cho else 0
            cy0 = int(cho[0].get('y') or 0) if cho else 0
            ccx = int(che[0].get('cx') or 0) if che else 0
            ccy = int(che[0].get('cy') or 0) if che else 0
            nsx = sx * (gcx / float(ccx) if ccx else 1.0)
            nsy = sy * (gcy / float(ccy) if ccy else 1.0)
            _walk(el, tm, out, ox + sx * gx - cx0 * nsx, oy + sy * gy - cy0 * nsy,
                  nsx, nsy, groups)
        elif tag in ('sp', 'pic', 'cxnSp', 'graphicFrame'):
            out.append(_read_shape(el, tag, tm, ox, oy, sx, sy))


def _open_pkg(file_path):
    if not zipfile.is_zipfile(file_path):
        return (None, '文件不是有效的 OPC(zip) 包，无法作为 .pptx 打开')
    try:
        pkg = zipfile.ZipFile(file_path, 'r')
    except Exception as exc:
        return (None, '演示文稿包无法打开：%s' % exc)
    bad = pkg.testzip()
    if bad:
        pkg.close()
        return (None, '包内部件已损坏：%s' % bad)
    if 'ppt/presentation.xml' not in pkg.namelist():
        pkg.close()
        return (None, '缺少 ppt/presentation.xml，不是 PowerPoint 演示文稿')
    return (pkg, '')


def _slide_parts(pkg):
    """按 sldIdLst 顺序返回幻灯片部件名列表"""
    pres = etree.fromstring(pkg.read('ppt/presentation.xml'))
    rels = {}
    try:
        rr = etree.fromstring(pkg.read('ppt/_rels/presentation.xml.rels'))
        for rel in rr:
            rels[rel.get('Id')] = rel.get('Target')
    except Exception:
        pass
    names = []
    lst = pres.find('p:sldIdLst', NS)
    if lst is not None:
        for sid in lst:
            rid = sid.get('{%s}id' % R)
            tgt = rels.get(rid)
            if not tgt:
                continue
            tgt = tgt.replace('../', '')
            if not tgt.startswith('ppt/'):
                tgt = 'ppt/' + tgt
            names.append(tgt)
    if not names:
        names = sorted(n for n in pkg.namelist()
                       if n.startswith('ppt/slides/slide') and n.endswith('.xml'))
    sz = pres.find('p:sldSz', NS)
    pw = int(sz.get('cx')) / EMU_CM if sz is not None else 0.0
    ph = int(sz.get('cy')) / EMU_CM if sz is not None else 0.0
    return names, pw, ph


def _cw(ch, size):
    """单字符前进宽度（cm）：全角按 CJK_ADV×字号，半角折半"""
    o = ord(ch)
    full = (0x1100 <= o <= 0x115F or 0x2E80 <= o <= 0xA4CF or 0xAC00 <= o <= 0xD7A3
            or 0xF900 <= o <= 0xFAFF or 0xFE30 <= o <= 0xFE6F or 0xFF00 <= o <= 0xFF60
            or 0xFFE0 <= o <= 0xFFE6 or o > 0xFFFF)
    return size * PT_CM * CJK_ADV * (1.0 if full else 0.5)


def _wrap(hard, size, inner_w):
    out = []
    for line in hard:
        if not line:
            out.append('')
            continue
        cur, w = '', 0.0
        for ch in line:
            cv = _cw(ch, size)
            if inner_w > 0 and cur and w + cv > inner_w + 1e-6:
                out.append(cur)
                cur, w = ch, cv
            else:
                cur += ch
                w += cv
        out.append(cur)
    return out or ['']


def _txt_lines(sh):
    """估算渲染后的每一行文字。
    与办公软件对齐：形状开启“根据文字调整形状大小”(spAutoFit) 时，生成端已把排版结果
    写进了形状高度，用高度反推行数比按字宽硬算更贴近软件实际排版，因此优先采信高度。"""
    hard = (sh.text or '').split('\n')
    if not any(hard):
        return ['']
    size = sh.size or 18.0
    inner_w = max(0.0, sh.W - sh.ins[0] - sh.ins[1])
    if sh.wrap == 'none':
        return hard
    est = _wrap(hard, size, inner_w)
    lh = size * PT_CM * LINE_H
    if sh.af == 'spAf' and lh > 0:
        n = max(1, int(round(max(0.0, sh.H - sh.ins[2] - sh.ins[3]) / lh)))
        if n == len(hard):
            return hard
    return est


def _txt_rect(sh):
    """估算文字实际占据的范围 (L, T, R, B)；空文本返回 None"""
    if not (sh.text or '').strip():
        return None
    lines = _txt_lines(sh)
    size = sh.size or 18.0
    lh = size * PT_CM * LINE_H
    il = sh.L + sh.ins[0]
    ir = sh.R - sh.ins[1]
    inner_w = max(0.0, ir - il)
    ws = []
    for ln in lines:
        w = sum(_cw(c, size) for c in ln)
        if sh.wrap != 'none' and inner_w > 0:
            w = min(w, inner_w)
        ws.append(w)
    th = lh * len(lines)
    anchor = sh.anchor or 't'
    if anchor == 'ctr':
        t = sh.cy - th / 2.0
    elif anchor == 'b':
        t = sh.B - sh.ins[3] - th
    else:
        t = sh.T + sh.ins[2]
    algn = None
    for a in sh.algns:
        if a:
            algn = a
            break
    xs = []
    for w in ws:
        if algn == 'ctr':
            l = (il + ir) / 2.0 - w / 2.0
        elif algn in ('r', 'just'):
            l = ir - w
        else:
            l = il
        xs.append((l, l + w))
    return (min(x[0] for x in xs), t, max(x[1] for x in xs), t + th)


def _picto(s):
    """文本里是否含有图形化符号（emoji / dingbat），用于识别“以字符充当图标”"""
    for ch in s or '':
        o = ord(ch)
        if o >= 0x1F300 or 0x2190 <= o <= 0x27BF or 0x2B00 <= o <= 0x2BFF:
            return True
    return False


def _pick_arrow(a, b, cands, used):
    """在 a、b 两形状的间隙里挑一个向右的箭头；返回 (箭头, 侵入深度) 或 (None, None)"""
    gl, gr = a.vR, b.vL
    ymid = (a.cy + b.cy) / 2.0
    for s in cands:
        if id(s) in used:
            continue
        if s.vW <= s.vH:                        # 需为横向（向右）
            continue
        if s.vL < gl - 0.35 or s.vR > gr + 0.35:
            continue
        if abs(s.vcy - ymid) > 1.6:
            continue
        pen_a = _xspan(a, s.vT, s.vB)[1] - s.vL
        pen_b = s.vR - _xspan(b, s.vT, s.vB)[0]
        used.add(id(s))
        return (s, max(pen_a, pen_b))
    return (None, None)


# ---------------------------------------------------------------------------
# 结构识别：把幻灯片还原成“标题/装饰带/13环节/箭头/回流/分支标注/图例”
# ---------------------------------------------------------------------------
class Ctx(object):
    pass


def _analyze(pkg):
    c = Ctx()
    c.tm = _theme_map(pkg)
    names, c.pw, c.ph = _slide_parts(pkg)
    c.nslides = len(names)
    c.shapes = []
    if names:
        root = etree.fromstring(pkg.read(names[0]))
        tree = root.find('.//p:cSld/p:spTree', NS)
        if tree is not None:
            _walk(tree, c.tm, c.shapes)
    page_area = max(1e-6, c.pw * c.ph)
    c.page_area = page_area

    # 整页底板/边框（不参与幅面与压盖计算）
    c.backdrops = [s for s in c.shapes
                   if s.vW >= 0.85 * c.pw and s.vH >= 0.85 * c.ph
                   and s.vW * s.vH >= 0.75 * page_area]
    bset = set(id(s) for s in c.backdrops)

    # 残留文本框：零尺寸或空内容
    c.residuals = [s for s in c.shapes if s.txbox
                   and (s.W < 0.35 or s.H < 0.35 or not (s.text or '').strip())]

    texts = [s for s in c.shapes if (s.text or '').strip()]
    c.texts = texts

    # 主标题
    c.title = None
    for s in texts:
        if _norm(s.text) == TITLE_TEXT and s.vcy <= c.ph * 0.30:
            if c.title is None or (s.size or 0) > (c.title.size or 0):
                c.title = s
    c.title_exact = c.title is not None
    if c.title is None:                     # 文字有错漏时退化为顶部最大字号文本，供幅面/字号项使用
        top = [s for s in texts if s.vcy <= c.ph * 0.30]
        if top:
            c.title = max(top, key=lambda s: (s.size or 0, s.vW))

    # ---- 13 个环节：先认文字，再认承载形状 ----
    cands = [s for s in c.shapes
             if s.kind == 'sp' and not s.txbox and id(s) not in bset
             and s.geom in (ROUND_GEOMS + ('diamond', 'rect', 'hexagon', 'flowChartDecision'))
             and s.W >= 1.5 and s.H >= 1.0 and s.W * s.H < 0.55 * page_area]
    c.cands = cands
    c.nodes = {}            # text -> {'label':Shape, 'node':Shape}
    used_node = {}
    for want in NODE_TEXTS:
        lab = None
        for s in texts:
            if _norm(s.text) == want:
                lab = s
                break
        if lab is None:
            continue
        owner = lab if (lab in cands) else None
        if owner is None:
            hit = [s for s in cands if _in_shape(s, lab.vcx, lab.vcy, 0.35)]
            if hit:
                hit.sort(key=lambda s: abs(s.cx - lab.vcx) + abs(s.cy - lab.vcy))
                owner = hit[0]
        if owner is None or id(owner) in used_node:
            continue
        used_node[id(owner)] = want
        c.nodes[want] = {'label': lab, 'node': owner}
    c.node_set = set(id(v['node']) for v in c.nodes.values())

    # 两排：按纵向中心分层，再按水平中心排序
    c.up_found = [t for t in NODE_TEXTS if t in c.nodes and c.nodes[t]['node'].cy < c.ph / 2.0]
    c.lo_found = [t for t in NODE_TEXTS if t in c.nodes and c.nodes[t]['node'].cy >= c.ph / 2.0]
    c.up_order = sorted(c.up_found, key=lambda t: c.nodes[t]['node'].cx)
    c.lo_order = sorted(c.lo_found, key=lambda t: c.nodes[t]['node'].cx)
    c.up_shapes = [c.nodes[t]['node'] for t in c.up_order]
    c.lo_shapes = [c.nodes[t]['node'] for t in c.lo_order]
    c.up_rects = [s for s in c.up_shapes if not s.is_diamond()]
    c.lo_rects = [s for s in c.lo_shapes if not s.is_diamond()]
    c.d1 = c.nodes[DECISION_TEXTS[0]]['node'] if DECISION_TEXTS[0] in c.nodes else None
    c.d2 = c.nodes[DECISION_TEXTS[1]]['node'] if DECISION_TEXTS[1] in c.nodes else None

    # ---- 箭头候选：带末端箭头的连接符 / 箭头形状 / 横向小图片（图标化箭头）----
    def _dir(s):
        return (s.tail not in (None, 'none')) or (s.head not in (None, 'none'))

    c.arrow_cands = []
    for s in c.shapes:
        if id(s) in bset or id(s) in c.node_set:
            continue
        if s.kind == 'pic':
            if (s.vW > s.vH and 0.35 <= s.vW <= 2.8 and s.vH <= 1.7
                    and s.vW * s.vH < 0.3 * page_area):
                c.arrow_cands.append(s)
        elif s.kind in ('cxn', 'sp'):
            if _dir(s) or (s.geom and 'Arrow' in s.geom):
                c.arrow_cands.append(s)
    c.dir_lines = [s for s in c.shapes if s.kind in ('cxn', 'sp') and _dir(s)]

    # 11 处相邻箭头逐一配对
    c.arrow_hits = []
    used = set()
    pairs = ([(c.up_shapes[i], c.up_shapes[i + 1]) for i in range(len(c.up_shapes) - 1)] +
             [(c.lo_shapes[i], c.lo_shapes[i + 1]) for i in range(len(c.lo_shapes) - 1)])
    c.pairs = pairs
    for a, b in pairs:
        c.arrow_hits.append(_pick_arrow(a, b, c.arrow_cands, used))

    # ---- “是”分支：菱形下顶点垂直落到“制定整改方案”上方 ----
    c.yes_branch = None
    tgt = c.nodes.get('制定整改方案', {}).get('node')
    if c.d1 is not None and tgt is not None:
        for s in c.dir_lines:
            if s.vW > 0.30 or s.vH < 0.8:
                continue
            if abs(s.vcx - c.d1.cx) > 0.6:
                continue
            if s.vT < c.d1.B - 0.5 or s.vB > tgt.T + 0.5:
                continue
            if not (tgt.L - 0.35 <= s.vcx <= tgt.R + 0.35):
                continue
            c.yes_branch = s
            break
    # 兜底：记录一条“从菱形下方连到下排”的走线，供未命中时给出取证说明
    c.yes_like = None
    if c.d1 is not None and tgt is not None:
        for s in c.shapes:
            if s.kind not in ('cxn', 'sp') or (s.text or '').strip():
                continue
            if s.vH < 0.8 or s.vT < c.d1.cy or s.vB > tgt.cy:
                continue
            if abs(s.vcx - c.d1.cx) > 2.5:
                continue
            c.yes_like = s
            break

    # ---- 两条回流折线（肘形连接符）----
    bents = [s for s in c.shapes
             if s.geom and (s.geom.startswith('bentConnector') or s.geom.startswith('curvedConnector'))]
    c.bents = bents
    c.dashed_lines = [s for s in c.shapes
                      if s.kind in ('cxn', 'sp') and not (s.text or '').strip()
                      and s.dash and s.dash != 'solid']

    def _dirp(s):
        return (s.tail not in (None, 'none')) or (s.head not in (None, 'none'))

    c.reflow_solid = None
    c.reflow_dash = None
    end_up = c.nodes.get('结案归档', {}).get('node')
    end_lo = c.nodes.get('执行整改', {}).get('node')
    for s in bents:
        if not _dirp(s):
            continue
        dashed = bool(s.dash) and s.dash != 'solid'
        if (not dashed) and end_up is not None and c.reflow_solid is None:
            if (_ov1(s.vL, s.vR, end_up.L, end_up.R) > -0.35
                    and end_up.B - 0.35 <= s.vT <= end_up.B + 1.8):
                c.reflow_solid = s
                continue
        if dashed and end_lo is not None and c.reflow_dash is None:
            if (_ov1(s.vL, s.vR, end_lo.L, end_lo.R) > -0.35
                    and s.vT >= end_lo.B - 0.35):
                c.reflow_dash = s

    # ---- 5 处分支标注 ----
    c.marks = [s for s in c.texts if _norm(s.text) in ('是', '否', '返工')]
    c.mk = {'d1_no': None, 'd1_yes': None, 'd2_yes': None, 'd2_no': None, 'rework': None}
    for s in c.marks:
        t = _norm(s.text)
        if c.d1 is not None and t == '否' and c.mk['d1_no'] is None:
            if s.vcx > c.d1.R - 0.45 and c.d1.T - 0.35 <= s.vcy <= c.d1.B + 0.35:
                c.mk['d1_no'] = s
                continue
        if c.d1 is not None and t == '是' and c.mk['d1_yes'] is None:
            if s.vcy > c.d1.B - 0.45 and abs(s.vcx - c.d1.cx) <= 1.8:
                c.mk['d1_yes'] = s
                continue
        if c.d2 is not None and t == '是' and c.mk['d2_yes'] is None:
            if s.vcx > c.d2.R - 0.7 and s.vcy < c.d2.cy:
                c.mk['d2_yes'] = s
                continue
        if c.d2 is not None and t == '否' and c.mk['d2_no'] is None:
            if s.vcy > c.d2.B - 0.45 and abs(s.vcx - c.d2.cx) <= 2.2:
                c.mk['d2_no'] = s
                continue
        if t == '返工' and c.mk['rework'] is None:
            c.mk['rework'] = s

    # ---- 菱形圆形徽标 ----
    ells = [s for s in c.shapes if s.kind == 'sp' and s.geom == 'ellipse'
            and max(s.W, s.H) <= 1.35 and min(s.W, s.H) >= 0.15]
    c.badges = {}
    for key, d in (('up', c.d1), ('low', c.d2)):
        if d is None:
            continue
        for e in ells:
            if not _in_shape(d, e.cx, e.cy):
                continue
            mark = None
            if _norm(e.text) in ('？', '✓', '√', '✔') or _picto(e.text):
                mark = e
            else:
                for p in c.shapes:
                    if p.kind == 'pic' and e.L <= p.cx <= e.R and e.T <= p.cy <= e.B:
                        mark = p
                        break
            c.badges[key] = (e, mark)
            break
    badge_ids = set()
    for e, m in c.badges.values():
        badge_ids.add(id(e))
        if m is not None:
            badge_ids.add(id(m))

    # ---- 图例块 ----
    c.legend = None
    leg_hint = [s for s in c.texts
                if LEGEND_TITLE in _norm(s.text) or any(k in _norm(s.text) for k in LEGEND_ITEMS)]
    for s in c.cands:
        if id(s) in c.node_set or s.cx > c.pw / 2.0 or s.cy < c.ph / 2.0:
            continue
        if s.W < 3.0 or s.H < 1.8:
            continue
        if any(_in_shape(s, h.vcx, h.vcy, 0.5) for h in leg_hint):
            if c.legend is None or s.W * s.H > c.legend.W * c.legend.H:
                c.legend = s
    c.leg_paras = []       # [(shape, para_text, size, color, y_center)]
    if c.legend is not None:
        for s in c.texts:
            if not _in_shape(c.legend, s.vcx, s.vcy, 0.6):
                continue
            lines = (s.text or '').split('\n')
            size = s.size or 18.0
            lh = size * PT_CM * LINE_H
            r = _txt_rect(s)
            top = r[1] if r else s.T
            for i, ln in enumerate(lines):
                if not ln.strip():
                    continue
                rs = [x for x in s.runs if x[0] and x[0].strip() and ln.strip()[:1] in x[0]]
                sz = rs[0][1] if rs else size
                col = rs[0][4] if rs else s.color
                c.leg_paras.append((s, ln, sz or size, col, top + lh * (i + 0.5)))
    c.leg_paras.sort(key=lambda x: x[4])
    c.leg_title = None
    c.leg_items = []
    for it in c.leg_paras:
        n = _norm(it[1])
        if LEGEND_TITLE.rstrip('：') in n and not c.leg_items:
            if c.leg_title is None:
                c.leg_title = it
                continue
        if any(k in n for k in LEGEND_ITEMS):
            c.leg_items.append(it)

    # 圆形底衬 + 记录图标（图标可以是嵌入图片，也可以是象形字符）
    c.leg_circle = None
    c.leg_icon = None
    if c.legend is not None:
        for s in c.shapes:
            if s.geom != 'ellipse' or not _in_shape(c.legend, s.cx, s.cy, 0.3):
                continue
            if max(s.W, s.H) < 0.8:
                continue
            if c.leg_circle is None or s.area > c.leg_circle.area:
                c.leg_circle = s
        if c.leg_circle is not None:
            e = c.leg_circle
            if _picto(e.text) or (e.text or '').strip():
                c.leg_icon = e
            for p in c.shapes:
                if p.kind == 'pic' and e.L <= p.cx <= e.R and e.T <= p.cy <= e.B:
                    c.leg_icon = p
                    break

    # 每条记录前的小圆点：独立小图形，或段首的项目符号字符
    c.leg_bullets = []
    if c.legend is not None:
        dots = [s for s in c.shapes
                if s.geom in ('ellipse', 'roundRect', 'rect')
                and 0.05 <= max(s.W, s.H) <= 0.5
                and _in_shape(c.legend, s.cx, s.cy, 0.2)]
        for it in c.leg_items:
            n = (it[1] or '').strip()
            if n[:1] in BULLET_CHARS:
                c.leg_bullets.append(('char', it))
                continue
            hit = [d for d in dots if abs(d.cy - it[4]) <= 0.45 and d.cx < it[0].R]
            if hit:
                c.leg_bullets.append(('dot', hit[0]))

    # ---- 装饰带：标题与上排之间的两段横线 + 4 个圆点 ----
    band_hi = (min(s.vT for s in c.up_shapes) - 0.2) if c.up_shapes else c.ph * 0.40
    c.band_lines = [s for s in c.shapes
                    if s.kind in ('cxn', 'sp') and not (s.text or '').strip()
                    and s.vW >= 4.0 and s.vH <= 0.35 and s.vcy <= band_hi
                    and (_is_blue(s.lnc) or _is_teal(s.lnc))]
    c.band_lines.sort(key=lambda s: s.vcx)
    band_y = (sum(s.vcy for s in c.band_lines) / len(c.band_lines)) if c.band_lines else None
    c.band_dots = []
    if band_y is not None:
        for s in c.shapes:
            if s.geom != 'ellipse' or (s.text or '').strip():
                continue
            if not (0.08 <= max(s.W, s.H) <= 0.60):
                continue
            if abs(s.vcy - band_y) > 1.2:
                continue
            c.band_dots.append(s)
        c.band_dots.sort(key=lambda s: s.cx)

    # ---- 环节图标：落在环节形状内的图片（或象形字符形状）；菱形以其徽标标记充当图标 ----
    arrow_used = set(id(a) for a, _ in c.arrow_hits if a is not None)
    leg_ids = set()
    for s in (c.legend, c.leg_circle, c.leg_icon):
        if s is not None:
            leg_ids.add(id(s))
    c.icons = {}
    for t, v in c.nodes.items():
        nd, lab = v['node'], v['label']
        best = None
        for s in c.shapes:
            if id(s) in bset or id(s) in arrow_used or id(s) in c.node_set or id(s) in leg_ids:
                continue
            if s is lab or s is nd:
                continue
            if not (s.kind == 'pic' or (s.kind == 'sp' and _picto(s.text))):
                continue
            if not _in_shape(nd, s.cx, s.cy):
                continue
            if best is None or s.area > best.area:
                best = s
        c.icons[t] = best

    # ---- 参与幅面/居中计算的内容集合：排除整页底板与退化形状 ----
    c.content = [s for s in c.shapes
                 if id(s) not in bset and not (s.vW < 0.35 and s.vH < 0.35)]
    if c.content:
        c.bbox = (min(s.vL for s in c.content), min(s.vT for s in c.content),
                  max(s.vR for s in c.content), max(s.vB for s in c.content))
    else:
        c.bbox = None
    c.badge_ids = badge_ids
    return c


# ---------------------------------------------------------------------------
# 维度2：24 个加分项（对应《细则.txt》第 4-27 行，合计 92 分）
# ---------------------------------------------------------------------------
def _score(c):
    """按细则第 4-27 行逐条打分（每行一条，共 24 条），返回 (results, details)；
    两张表的顺序与 _MAX_DELTAS_ORDER 严格一一对应。
    details 仅作本地排查用，不进入 evaluate 的返回值（dim2_items 不含 detail 字段）。"""
    results = []
    details = []

    def emit(delta, rule, detail=""):
        results.append("+%d：%s" % (delta, rule))
        details.append(detail)

    def fam(sh):
        """字体家族：主题字体占位符（+mn-lt 等）归一为空串，视作同一家族"""
        f = sh.font or ''
        return '' if f.startswith('+') else f

    def cont(sh, r):
        """文本范围 r 越出形状 sh 的最大量（菱形按棱形真实轮廓）"""
        if r is None:
            return 0.0
        l, t, rr, b = r
        if sh.is_diamond():
            xs = _xspan(sh, t, b)
            over = max(xs[0] - l, rr - xs[1], sh.T - t, b - sh.B)
        else:
            over = max(sh.L - l, rr - sh.R, sh.T - t, b - sh.B)
        return max(0.0, over)

    nodes = [v['node'] for v in c.nodes.values()]
    labels = dict((t, v['label']) for t, v in c.nodes.items())
    nmap = dict((id(v['node']), t2) for t2, v in c.nodes.items())

    # +5：流程图幅面与给定图幅相当——标题、13个环节形状、连线、分支标注、图例框在内的
    #     所有元素外接范围宽度≥32cm、高度≥17cm，未整体缩成小图偏居一隅
    # 严格对齐细则，仅踩细则明确写出的点，不额外约束：
    #   ① 外接范围取全部内容对象的可视范围并集，宽≥32cm 且高≥17cm 即命中；
    #   ② 包裹整页的底板/边框属细则第6行单独扣的"额外装饰"，不能替流程图撑幅面，故剔除；
    #   ③ 零尺寸残留对象（两向都<0.35cm）不计入，避免残留框把范围虚撑到页角。
    # 与办公软件对齐：这里的外接范围等价于在 PowerPoint 里全选后看到的选择框，
    #   旋转对象按旋转后的可视框计（"肘形连接符"带 rot=90 时尤其明显）。
    bb = c.bbox
    bw = (bb[2] - bb[0]) if bb else 0.0
    bh = (bb[3] - bb[1]) if bb else 0.0
    emit(5 if (bw >= 32.0 and bh >= 17.0) else 0,
         "流程图幅面与给定图幅相当——所有元素外接范围宽度≥32cm、高度≥17cm",
         "外接范围 %.2f×%.2fcm（x %.2f-%.2f，y %.2f-%.2f）"
         % (bw, bh, bb[0], bb[2], bb[1], bb[3]) if bb else "页面内未发现内容对象")

    # +3：流程图整体在页面内居中放置——上述元素外接范围的水平中心、垂直中心与页面中心的
    #     偏差均≤0.5cm
    dx = abs((bb[0] + bb[2]) / 2.0 - c.pw / 2.0) if bb else 99.0
    dy = abs((bb[1] + bb[3]) / 2.0 - c.ph / 2.0) if bb else 99.0
    emit(3 if (dx <= 0.5 and dy <= 0.5) else 0,
         "流程图整体在页面内居中放置——外接范围的水平中心、垂直中心与页面中心偏差均≤0.5cm",
         "水平偏差 %.2fcm，垂直偏差 %.2fcm（页面 %.2f×%.2fcm）" % (dx, dy, c.pw, c.ph))

    # +1：页面洁净无冗余——仅1页幻灯片，无零尺寸/空内容的残留文本框，且未额外添加包裹整页的
    #     边框、底板等给定内容之外的装饰对象
    # 严格对齐细则，仅踩细则明确写出的点，不额外约束：
    #   ① 演示文稿只能有 1 页幻灯片（维度1门禁只判"能否打开"，页数由本条按 1 分考核）；
    #   ② 不得存在零尺寸（任一边<0.35cm）或空内容的文本框；
    #   ③ 不得存在覆盖整页（宽高均≥页面85%且面积≥75%）的边框/底板对象。
    # 与办公软件对齐：页数按 presentation.xml 的 p:sldIdLst 计，即 PowerPoint 左侧
    #   幻灯片窗格里可见的页数；空白页也算 1 页。
    clean = []
    if c.nslides != 1:
        clean.append('幻灯片共 %d 页（应仅 1 页）' % c.nslides)
    if c.residuals:
        clean.append('残留文本框 %d 个（%s）'
                     % (len(c.residuals), '、'.join(s.name for s in c.residuals[:3])))
    if c.backdrops:
        clean.append('整页装饰对象 %d 个（%s）'
                     % (len(c.backdrops), '、'.join(s.name for s in c.backdrops[:3])))
    emit(0 if clean else 1,
         "页面洁净无冗余——仅1页幻灯片，无零尺寸/空内容残留文本框，无包裹整页的边框或底板装饰",
         '；'.join(clean) or '仅 1 页幻灯片、无残留框、无整页装饰')

    # +1：顶部设有主标题"客户投诉闭环处理流程"，单行完整不换行、无错字漏字，
    #     水平居中（与页面中心偏差≤1.5cm），标题顶端距页顶≤1.5cm
    t = c.title
    tr = _txt_rect(t) if t is not None else None
    tcx = ((tr[0] + tr[2]) / 2.0) if tr else (t.vcx if t is not None else 0.0)
    t_off = abs(tcx - c.pw / 2.0) if t is not None else 99.0
    t_lines = len(_txt_lines(t)) if t is not None else 0
    ok_title = (c.title_exact and t_lines == 1 and t_off <= 1.5 and t.vT <= 1.5)
    emit(1 if ok_title else 0,
         "顶部设有主标题\"客户投诉闭环处理流程\"，单行完整不换行、无错字漏字，水平居中（偏差≤1.5cm），顶端距页顶≤1.5cm",
         ("标题=%r 行数%d 水平偏差%.2fcm 顶端%.2fcm"
          % (_norm(t.text), t_lines, t_off, t.vT)) if t is not None else '未发现主标题文本')

    # +3：标题字号≥45磅，且不小于环节文字字号的2.5倍，视觉上明显为全页主标题
    tsz = (t.size if t is not None else None) or 0.0
    nsz = [labels[k].size for k in labels if labels[k].size]
    nmax = max(nsz) if nsz else 0.0
    emit(3 if (tsz >= 45.0 and nmax > 0 and tsz >= 2.5 * nmax) else 0,
         "标题字号≥45磅，且不小于环节文字字号的2.5倍，视觉上明显为全页主标题",
         "标题 %.1f磅，环节文字最大 %.1f磅（需≥%.1f磅）" % (tsz, nmax, 2.5 * nmax))

    # +3：标题下方设有装饰带——左右各一段水平短横线（单段长度≥4cm，分别取蓝色系与
    #     青绿色系）、中间一行4个小圆点分别为深蓝/亮蓝/橙/青绿四色，圆点直径0.2-0.35cm、
    #     横向等间距（相邻间距极差≤0.1cm）并落在两段横线之间
    # 严格对齐细则，仅踩细则明确写出的点，不额外约束：
    #   ① 两段横线各自长度≥4cm，且一段为蓝色系、一段为青绿色系；
    #   ② 圆点恰 4 个、直径 0.2-0.35cm、相邻横向间距极差≤0.1cm；
    #   ③ 四个圆点颜色分别覆盖深蓝、亮蓝、橙、青绿四个色系，且横向落在两段横线之间。
    # 与办公软件对齐：色系按 HSL 色相归类（蓝 195-255°、青绿 155-195°、橙 8-48°），
    #   深蓝/亮蓝以明度 0.36 为界，主题色先展开为 RGB 再判色系。
    bl = c.band_lines
    dots = c.band_dots
    dcol = [(s.fill if s.fill not in (None, 'NONE') else s.lnc) for s in dots]
    gaps = [dots[i + 1].cx - dots[i].cx for i in range(len(dots) - 1)]
    band = []
    if len(bl) < 2:
        band.append('横线 %d 段（需2段≥4cm的蓝/青绿横线）' % len(bl))
    elif not (any(_is_blue(s.lnc) for s in bl) and any(_is_teal(s.lnc) for s in bl)):
        band.append('两段横线未分别为蓝色系与青绿色系（%s）'
                    % '、'.join(str(s.lnc) for s in bl[:2]))
    if len(dots) != 4:
        band.append('圆点 %d 个（需4个）' % len(dots))
    else:
        dia = [max(s.W, s.H) for s in dots]
        if not all(0.2 <= v <= 0.35 for v in dia):
            band.append('圆点直径 %s（需0.2-0.35cm）' % '/'.join('%.2f' % v for v in dia))
        if _rng(gaps) > 0.1:
            band.append('圆点间距极差 %.2fcm（需≤0.1cm）' % _rng(gaps))
        need = (_is_dark_blue, _is_light_blue, _is_orange, _is_teal)
        miss = [n for n, f in zip(('深蓝', '亮蓝', '橙', '青绿'), need)
                if not any(f(v) for v in dcol)]
        if miss:
            band.append('圆点缺%s色（实际 %s）' % ('/'.join(miss), '、'.join(str(v) for v in dcol)))
        if len(bl) >= 2:
            lo, hi = bl[0].vR, bl[-1].vL
            if not all(lo - 0.3 <= s.cx <= hi + 0.3 for s in dots):
                band.append('圆点未落在两段横线之间（横线内侧 %.2f-%.2fcm）' % (lo, hi))
    emit(0 if band else 3,
         "标题下方设有装饰带——左右各一段≥4cm的蓝色系/青绿色系短横线，中间4个直径0.2-0.35cm的深蓝/亮蓝/橙/青绿圆点等间距落在两线之间",
         '；'.join(band) or ('横线 %d 段（%s）、圆点 %d 个间距极差%.2fcm'
                            % (len(bl), '、'.join(str(s.lnc) for s in bl[:2]),
                               len(dots), _rng(gaps))))

    # +5：上排完整呈现7个环节且顺序正确：接收投诉→信息登记→问题分级→是否重大问题？（判定）
    #     →客服处理→客户确认→结案归档，文字无错字、漏字、并字
    # 严格对齐细则，仅踩细则明确写出的点，不额外约束：
    #   ① 7 段文字逐字命中（归一化只去空白/换行、统一全角问号，错字漏字并字都会落空）；
    #   ② 按承载形状的水平中心从左到右排序后须与给定次序完全一致。
    emit(5 if list(c.up_order) == list(UPPER_TEXTS) else 0,
         "上排完整呈现7个环节且顺序正确：接收投诉→信息登记→问题分级→是否重大问题？→客服处理→客户确认→结案归档，文字无错漏并字",
         "实际上排（左→右）：%s" % ('→'.join(c.up_order) or '未识别到上排环节'))

    # +5：下排完整呈现6个环节且顺序正确：指派专员→原因分析→制定整改方案→执行整改→
    #     质量复核→复核通过？（判定），文字无错字、漏字、并字
    emit(5 if list(c.lo_order) == list(LOWER_TEXTS) else 0,
         "下排完整呈现6个环节且顺序正确：指派专员→原因分析→制定整改方案→执行整改→质量复核→复核通过？，文字无错漏并字",
         "实际下排（左→右）：%s" % ('→'.join(c.lo_order) or '未识别到下排环节'))

    # +3：13个环节各由独立形状承载——11个圆角矩形节点＋2个菱形判定框，形状类型与语义匹配
    #     （判定环节必须为菱形，不得用矩形代替）
    # 严格对齐细则，仅踩细则明确写出的点，不额外约束：
    #   ① 13 个环节各自有独立承载形状；② 其中 11 个为圆角矩形（roundRect 系）；
    #   ③ 两个判定环节的形状必须是菱形（diamond）。
    rects = c.up_rects + c.lo_rects
    shp = []
    if len(c.nodes) != 13:
        shp.append('识别到 %d 个独立环节形状（需13个）' % len(c.nodes))
    bad_r = [s.name for s in rects if s.geom not in ROUND_GEOMS]
    if len(rects) != 11 or bad_r:
        shp.append('圆角矩形 %d 个%s' % (len(rects), ('，非圆角：' + '、'.join(bad_r[:3])) if bad_r else ''))
    for k, d in zip(DECISION_TEXTS, (c.d1, c.d2)):
        if d is None:
            shp.append('缺判定框"%s"' % k)
        elif not d.is_diamond():
            shp.append('"%s"用了 %s 而非菱形' % (k, d.geom))
    emit(0 if shp else 3,
         "13个环节各由独立形状承载——11个圆角矩形节点＋2个菱形判定框，判定环节必须为菱形",
         '；'.join(shp) or '11个圆角矩形＋2个菱形，形状类型与语义匹配')

    # +5：13个环节形状尺寸与幅面相称——上排6个矩形节点宽3.4-4.6cm、高2.3-2.8cm；
    #     下排5个矩形节点宽2.6-3.6cm、高2.9-3.5cm；2个菱形判定框均≥3.5cm×3.5cm
    #     且高度不低于同排相邻矩形节点
    sz = []
    if len(c.up_rects) != 6 or len(c.lo_rects) != 5:
        sz.append('上排矩形 %d 个、下排矩形 %d 个（需6+5）' % (len(c.up_rects), len(c.lo_rects)))
    for s in c.up_rects:
        if not (3.4 <= s.W <= 4.6 and 2.3 <= s.H <= 2.8):
            sz.append('上排%s %.2f×%.2f（需3.4-4.6×2.3-2.8）' % (s.name, s.W, s.H))
    for s in c.lo_rects:
        if not (2.6 <= s.W <= 3.6 and 2.9 <= s.H <= 3.5):
            sz.append('下排%s %.2f×%.2f（需2.6-3.6×2.9-3.5）' % (s.name, s.W, s.H))
    for k, d, rr in ((DECISION_TEXTS[0], c.d1, c.up_rects), (DECISION_TEXTS[1], c.d2, c.lo_rects)):
        if d is None:
            continue
        hmax = max([x.H for x in rr] or [0.0])
        if not (d.W >= 3.5 and d.H >= 3.5):
            sz.append('菱形"%s" %.2f×%.2f（需≥3.5×3.5）' % (k, d.W, d.H))
        elif d.H < hmax - 0.02:
            sz.append('菱形"%s"高 %.2f＜同排矩形 %.2f' % (k, d.H, hmax))
    emit(0 if sz else 5,
         "13个环节形状尺寸与幅面相称——上排矩形宽3.4-4.6cm高2.3-2.8cm、下排矩形宽2.6-3.6cm高2.9-3.5cm、菱形≥3.5×3.5cm且高不低于同排矩形",
         '；'.join(sz[:4]) or '上下两排矩形与2个菱形尺寸均在区间内')

    # +4：形状描边规格与三段配色分区统一（细则第14行；由原细则的描边规格 +1 与三段配色 +3 合并
    #     而来，原描边条里"所有环节框为圆角矩形"与第12行重复，形状类型由第12行单独考核，
    #     此处不再重复判定）
    #     ——描边为实线、线宽1.5-2.5磅，填充为白色或浅色底且文字为深色对比清晰；上排前3个环节
    #     蓝色系、上排后3个环节与菱形"是否重大问题？"青绿色系、下排6个环节橙色系描边，
    #     且同一配色分区内各环节的描边颜色与线宽完全一致
    # 严格对齐细则，仅踩细则明确写出的点，不额外约束：
    #   ① 描边为实线（无 prstDash 即默认实线）、线宽 1.5-2.5 磅；
    #   ② 填充为白色/浅色（未显式填充或 noFill 时随浅色页底，按浅色计），文字为深色；
    #   ③ 三个分区的描边色分别属蓝/青绿/橙色系；
    #   ④ 各分区内部的描边颜色与线宽必须完全一致。
    # 与办公软件对齐：线宽取 a:ln/@w÷12700 得磅值，未写 @w 时软件按 1 磅细线渲染。
    uni = []
    for s in nodes:
        if s.dash and s.dash != 'solid':
            uni.append('%s 描边为%s（需实线）' % (s.name, s.dash))
        w = s.lnw if s.lnw is not None else 1.0
        if not (1.5 <= w <= 2.5):
            uni.append('%s 线宽 %.2f磅（需1.5-2.5磅）' % (s.name, w))
        if s.fill not in (None, 'NONE') and not _light(s.fill, 0.72):
            uni.append('%s 填充 %s 非浅色' % (s.name, s.fill))
    for t2, lab in labels.items():
        if lab.color and not _dark(lab.color):
            uni.append('"%s"文字色 %s 不够深' % (t2, lab.color))
    zones = (
        ('上排前3（蓝）', ('接收投诉', '信息登记', '问题分级'), _is_blue),
        ('上排后3+菱形（青绿）', ('客服处理', '客户确认', '结案归档', DECISION_TEXTS[0]), _is_teal),
        ('下排6（橙）', LOWER_TEXTS, _is_orange),
    )
    for zn, keys, fn in zones:
        zs = [c.nodes[k]['node'] for k in keys if k in c.nodes]
        if len(zs) != len(keys):
            uni.append('%s仅识别到 %d/%d 个环节' % (zn, len(zs), len(keys)))
        bad_c = ['%s(%s)' % (nmap.get(id(s), s.name), s.lnc) for s in zs if not fn(s.lnc)]
        if bad_c:
            uni.append('%s描边色不符：%s' % (zn, '、'.join(bad_c[:3])))
        cols = set(str(s.lnc) for s in zs)
        wids = set(round(s.lnw if s.lnw is not None else 1.0, 2) for s in zs)
        if len(cols) > 1:
            uni.append('%s描边色不一致（%s）' % (zn, '、'.join(sorted(cols))))
        if len(wids) > 1:
            uni.append('%s线宽不一致（%s磅）' % (zn, '、'.join('%.2f' % v for v in sorted(wids))))
    emit(0 if uni else 4,
         "形状描边规格与三段配色分区统一——描边实线、线宽1.5-2.5磅、浅色底深色字；上排前3蓝色系、上排后3与菱形\"是否重大问题？\"青绿色系、下排6橙色系，且同一分区内描边颜色与线宽完全一致",
         '；'.join(uni[:4]) or '描边实线1.5-2.5磅、浅底深字，蓝/青绿/橙三段分区正确且各分区内描边色与线宽完全一致')

    # +4：两排纵向对齐与定位准确（细则第15行；由原细则的顶端对齐 +3 与纵向定位 +1 合并而来——
    #     同一组 y 坐标测量，顶端齐＋高度齐即中心齐，定位条只多"落在分层带内"与
    #     "菱形与同排中心对齐"两句）
    #     ——上排6个矩形节点顶端对齐（极差≤0.2cm）且高度一致；下排5个矩形节点顶端对齐
    #     （极差≤0.2cm）且高度一致；上排环节垂直中心落在5.5-7.0cm、下排落在11.5-13.0cm，
    #     2个菱形分别与同排矩形节点垂直中心对齐（偏差≤0.3cm）
    aln = []
    for nm, rr, cnt in (('上排', c.up_rects, 6), ('下排', c.lo_rects, 5)):
        if len(rr) != cnt:
            aln.append('%s矩形 %d 个（需%d个）' % (nm, len(rr), cnt))
            continue
        rt = _rng([s.T for s in rr])
        rh = _rng([s.H for s in rr])
        if rt > 0.2 or rh > 0.2:
            aln.append('%s顶端极差 %.2fcm、高度极差 %.2fcm' % (nm, rt, rh))
    for nm, rr, d, lo2, hi2 in (('上排', c.up_rects, c.d1, 5.5, 7.0),
                                ('下排', c.lo_rects, c.d2, 11.5, 13.0)):
        if not rr:
            aln.append('%s未识别到矩形节点' % nm)
            continue
        cys = [s.cy for s in rr]
        out = ['%.2f' % v for v in cys if not (lo2 <= v <= hi2)]
        if out:
            aln.append('%s有节点垂直中心 %s（需%.1f-%.1fcm）' % (nm, '、'.join(out), lo2, hi2))
        if d is not None:
            dv = abs(d.cy - sum(cys) / len(cys))
            if dv > 0.3:
                aln.append('%s菱形与矩形垂直中心偏差 %.2fcm（需≤0.3cm）' % (nm, dv))
    emit(0 if aln else 4,
         "两排纵向对齐与定位——上排6个矩形节点、下排5个矩形节点各自顶端对齐（极差≤0.2cm）且高度一致；上排垂直中心落在5.5-7.0cm、下排落在11.5-13.0cm，2个菱形与同排矩形垂直中心偏差≤0.3cm",
         '；'.join(aln[:4]) or ('上排顶端极差%.2f/高度极差%.2f、中心%.2fcm；下排 %.2f/%.2f、中心%.2fcm；菱形对齐偏差均≤0.3cm'
                               % (_rng([s.T for s in c.up_rects]), _rng([s.H for s in c.up_rects]),
                                  sum(s.cy for s in c.up_rects) / max(1, len(c.up_rects)),
                                  _rng([s.T for s in c.lo_rects]), _rng([s.H for s in c.lo_rects]),
                                  sum(s.cy for s in c.lo_rects) / max(1, len(c.lo_rects)))))

    # +3：两排相邻矩形节点之间的水平间隙均匀，各排内间隙极差≤0.5cm，不出现某一处间隙被
    #     明显拉大（如"客户确认→结案归档"间距远大于其余节点间距）
    # 严格对齐细则，仅踩细则明确写出的点，不额外约束：
    #   ① 只比"相邻矩形节点之间"的间隙——跨菱形的那两处间隙按细则字面不计入；
    #   ② 每排各自算极差，任一排＞0.5cm 即落空。
    gapinfo = []
    gap_bad = []
    for nm, seq in (('上排', c.up_order), ('下排', c.lo_order)):
        gs = []
        for i in range(len(seq) - 1):
            a, b = c.nodes[seq[i]]['node'], c.nodes[seq[i + 1]]['node']
            if a.is_diamond() or b.is_diamond():
                continue
            gs.append((b.L - a.R, seq[i], seq[i + 1]))
        gapinfo.append('%s间隙 %s（极差%.2fcm）'
                       % (nm, '/'.join('%.2f' % g[0] for g in gs), _rng([g[0] for g in gs])))
        if _rng([g[0] for g in gs]) > 0.5:
            wide = max(gs, key=lambda g: g[0])
            gap_bad.append('%s间隙极差 %.2fcm，最宽处"%s→%s"达 %.2fcm'
                           % (nm, _rng([g[0] for g in gs]), wide[1], wide[2], wide[0]))
    emit(0 if gap_bad else 3,
         "两排相邻矩形节点之间的水平间隙均匀，各排内间隙极差≤0.5cm，不出现某一处间隙被明显拉大",
         '；'.join(gap_bad) or '；'.join(gapinfo))

    # +5：相邻环节之间共设11处向右的箭头（上排6处、下排5处），每处箭头均落在两形状的间隙内、
    #     不穿入形状内部或与边线交叠，末端箭头端点明确、指向下一环节方向正确
    # 严格对齐细则，仅踩细则明确写出的点，不额外约束：
    #   ① 11 处相邻环节之间逐一都要有一个横向对象；
    #   ② 该对象须落在两形状的水平间隙内，穿入形状的深度≤0.3cm（含与边线交叠）；
    #   ③ 方向向右——按左右相邻形状配对，横向跨度大于纵向跨度即视为指向下一环节。
    # 与办公软件对齐：箭头既可能是带 tailEnd 的连接符，也可能是插入的箭头图片/箭头形状，
    #   三者在软件里都是"一个可单独点选的箭头"，故一并接受；菱形按棱形轮廓算侵入深度。
    arr = []
    if len(c.pairs) != 11:
        arr.append('相邻环节对 %d 处（需11处）' % len(c.pairs))
    pen_max = 0.0
    for (a, b), (s, pen) in zip(c.pairs, c.arrow_hits):
        tag = '%s→%s' % (nmap.get(id(a), a.name), nmap.get(id(b), b.name))
        if s is None:
            arr.append('"%s"缺向右箭头' % tag)
        else:
            pen_max = max(pen_max, pen)
            if pen > 0.3:
                arr.append('"%s"箭头穿入形状 %.2fcm' % (tag, pen))
    emit(0 if arr else 5,
         "相邻环节之间共设11处向右的箭头（上排6处、下排5处），均落在两形状间隙内、不穿入形状或与边线交叠，末端箭头指向下一环节",
         '；'.join(arr[:4]) or '11处向右箭头齐备，最大侵入 %.2fcm' % pen_max)

    # +8：分支与回流走线正确（细则第18行；由原细则的"是"分支竖直 +3 与两条回流折线 +5
    #     合并而来——同为"非相邻环节之间的走线规格"）
    #     ——菱形"是否重大问题？"的"是"分支为竖直向下的箭头（起止点水平偏差≤0.3cm，自菱形下顶点
    #     垂直落到下排"制定整改方案"上方，未画成斜线）；两条回流走线均为折线（肘形）且末端带箭头，
    #     "复核通过？→是"实线折回指向"结案归档"下方、"复核通过？→否"虚线折回指向"执行整改"下方，
    #     两条走线均不穿越其他环节框
    # 严格对齐细则，仅踩细则明确写出的点，不额外约束：
    #   ① "是"分支横向跨度≤0.3cm 且自菱形下顶点落到"制定整改方案"上方；
    #   ② 两条回流都必须是折线——几何为 bentConnector/curvedConnector 系（直线段拼出来的不算）；
    #   ③ 一条实线折回到"结案归档"下方，一条虚线折回到"执行整改"下方，且都带末端箭头；
    #   ④ 不穿越其他环节框——按"走线范围内不含其他环节中心、且与其重叠面积≤该框25%"判定。
    # 与办公软件对齐：肘形连接符的外接框远大于其实际墨迹（带 rot 时更明显），
    #   故不以外接框相交直接判穿越，只在明显覆盖环节框时才认定。
    yb = c.yes_branch
    if yb is not None:
        yd = '%s 横向跨度 %.2fcm、纵向 %.2fcm（自菱形下顶点垂直落到"制定整改方案"上方）' \
             % (yb.name, yb.vW, yb.vH)
    elif c.yes_like is not None:
        yd = '%s 横向跨度 %.2fcm（>0.3cm，为倾斜斜线而非竖直向下）' % (c.yes_like.name, c.yes_like.vW)
    else:
        yd = '未发现自"是否重大问题？"下方连到"制定整改方案"的走线'
    def crosses(s, keep):
        out = []
        for nd in nodes:
            if id(nd) in keep:
                continue
            if s.vL <= nd.cx <= s.vR and s.vT <= nd.cy <= s.vB:
                out.append(nd.name)
                continue
            ox = _ov1(s.vL, s.vR, nd.vL, nd.vR)
            oy = _ov1(s.vT, s.vB, nd.vT, nd.vB)
            if ox > 0 and oy > 0 and nd.area > 0 and ox * oy > 0.25 * nd.area:
                out.append(nd.name)
        return out

    ref = []
    if yb is None:
        ref.append(yd)
    for nm, s, tgt in (('实线回流（→"结案归档"下方）', c.reflow_solid, '结案归档'),
                       ('虚线回流（→"执行整改"下方）', c.reflow_dash, '执行整改')):
        if s is None:
            ref.append('缺%s' % nm)
            continue
        keep = set(id(x) for x in (c.d2, c.nodes.get(tgt, {}).get('node')) if x is not None)
        cr = crosses(s, keep)
        if cr:
            ref.append('%s穿越 %s' % (nm, '、'.join(cr[:3])))
    emit(0 if ref else 8,
         "分支与回流走线正确——\"是否重大问题？\"的\"是\"分支为竖直向下的箭头（水平偏差≤0.3cm、落到\"制定整改方案\"上方）；两条回流走线均为肘形折线且末端带箭头，一条实线折回指向\"结案归档\"下方、一条虚线折回指向\"执行整改\"下方，均不穿越其他环节框",
         '；'.join(ref[:3]) or ('%s；实线 %s + 虚线 %s，均为肘形带箭头且不穿越其他环节框'
                               % (yd, c.reflow_solid.name, c.reflow_dash.name)))

    # +3：分支标注文字齐备且就位——菱形"是否重大问题？"右侧标"否"、下方标"是"；
    #     菱形"复核通过？"右上方标"是"、下方标"否"；虚线回流走线上标"返工"，共5处
    # 严格对齐细则，仅踩细则明确写出的点，不额外约束：
    #   ① 5 段文字齐备；② 各自方位就位（右侧/下方/右上方/下方）；
    #   ③ "返工"须贴着虚线回流走线（到虚线可视范围距离≤2cm）。
    def pdist(sh, x, y):
        return math.hypot(max(sh.vL - x, x - sh.vR, 0.0), max(sh.vT - y, y - sh.vB, 0.0))

    mk = c.mk
    need_mk = (('d1_no', '"是否重大问题？"右侧的"否"'), ('d1_yes', '"是否重大问题？"下方的"是"'),
               ('d2_yes', '"复核通过？"右上方的"是"'), ('d2_no', '"复核通过？"下方的"否"'),
               ('rework', '虚线回流走线上的"返工"'))
    mkbad = ['缺%s' % lbl for k, lbl in need_mk if mk.get(k) is None]
    dsh = ([c.reflow_dash] if c.reflow_dash is not None else []) + c.dashed_lines
    if mk.get('rework') is not None:
        r = mk['rework']
        dd = min([pdist(s, r.vcx, r.vcy) for s in dsh] or [99.0])
        if dd > 2.0:
            mkbad.append('"返工"距虚线回流走线 %.2fcm（未标在走线上）' % dd)
    emit(0 if mkbad else 3,
         "分支标注文字齐备且就位——\"是否重大问题？\"右侧\"否\"、下方\"是\"；\"复核通过？\"右上方\"是\"、下方\"否\"；虚线回流走线上\"返工\"，共5处",
         '；'.join(mkbad) or '5处分支标注齐备且方位正确')

    # +3：分支标注文字（是／否／返工）字号≥20磅，与环节文字为同一字体家族
    mks = [s for s in (mk.get(k) for k, _ in need_mk) if s is not None]
    fsz = []
    small = ['%s %.1f磅' % (_norm(s.text), s.size or 0.0) for s in mks if (s.size or 0.0) < 20.0]
    if len(mks) < 5:
        fsz.append('仅识别到 %d 处分支标注' % len(mks))
    if small:
        fsz.append('字号不足20磅：%s' % '、'.join(small))
    mfam = set(fam(s) for s in mks)
    nfam = set(fam(labels[k]) for k in labels)
    if len(mfam | nfam) > 1:
        fsz.append('字体家族不一致（标注 %s / 环节 %s）'
                   % ('、'.join(sorted(x or '主题字体' for x in mfam)),
                      '、'.join(sorted(x or '主题字体' for x in nfam))))
    emit(0 if fsz else 3,
         "分支标注文字（是／否／返工）字号≥20磅，与环节文字为同一字体家族",
         '；'.join(fsz) or ('5处标注字号 %s磅，与环节文字同一字体家族'
                           % '/'.join('%.1f' % (s.size or 0.0) for s in mks)))

    # +8：环节图标齐备与规格（细则第21行；由原细则的图标齐备 +5 与图标规格 +3 合并而来——
    #     同一批 13 个图标对象，"有没有"与"合不合规"是单向依赖的两段判据）
    #     ——13个环节形状内均配有与该环节含义对应的图标，逐个到位不缺失、不以纯文字或空白代替；
    #     图标尺寸统一（长边0.7-1.6cm），完整落在所属形状内部、不越出边框，与该环节文字并排或
    #     上下排布而不相互压盖
    # 严格对齐细则，仅踩细则明确写出的点，不额外约束：
    #   ① 13 个环节形状内各有一个图标对象（嵌入图片，或以象形符号成形的独立形状）；
    #   ② "不以纯文字或空白代替"——环节文字本身、以及空形状都不计作图标；
    #   ③ 图标须落在所属环节形状内部（菱形按棱形轮廓判定）；
    #   ④ 长边落在 0.7-1.6cm、完整落在形状内部、与该环节文字不相互压盖。
    # 与办公软件对齐：图标在软件里就是环节框内可单独点选的一张图片/一个符号形状；
    #   两个菱形的圆形徽标标记即其语义图标，其规格由细则第22行单独约束，
    #   故 ④ 的尺寸/越界/压盖只考核 11 个矩形节点的图标。
    ico_miss = [t2 for t2 in NODE_TEXTS if c.nodes.get(t2) and c.icons.get(t2) is None]
    ico_n = sum(1 for t2 in c.icons if c.icons[t2] is not None)
    ic = []
    if len(c.nodes) != 13 or ico_miss:
        ic.append('%d/13 个环节有图标，缺：%s' % (ico_n, '、'.join(ico_miss[:6])))
    for t2 in list(UPPER_TEXTS) + list(LOWER_TEXTS):
        v = c.nodes.get(t2)
        if not v or v['node'].is_diamond():
            continue
        g = c.icons.get(t2)
        if g is None:
            ic.append('"%s"无图标' % t2)
            continue
        lg = max(g.vW, g.vH)
        if not (0.7 <= lg <= 1.6):
            ic.append('"%s"图标长边 %.2fcm（需0.7-1.6cm）' % (t2, lg))
        nd = v['node']
        if min(g.vL - nd.L, nd.R - g.vR, g.vT - nd.T, nd.B - g.vB) < -0.02:
            ic.append('"%s"图标越出形状边框' % t2)
        r = _txt_rect(v['label'])
        if r is not None:
            ox = _ov1(g.vL, g.vR, r[0], r[2])
            oy = _ov1(g.vT, g.vB, r[1], r[3])
            if ox > 0.25 and oy > 0.25:
                ic.append('"%s"图标与文字压盖 %.2f×%.2fcm' % (t2, ox, oy))
    emit(0 if ic else 8,
         "13个环节形状内均配有与该环节含义对应的图标，逐个到位不缺失、不以纯文字或空白代替；图标尺寸统一（长边0.7-1.6cm），完整落在所属形状内部不越出边框，与环节文字并排或上下排布而不压盖",
         '；'.join(ic[:4]) or '13个环节图标齐备，11个矩形节点图标长边均在0.7-1.6cm、完整在框内且不与文字压盖')

    # +1：2个菱形判定框内各设一个圆形徽标——上排菱形内为青绿色圆圈＋"？"标记，
    #     下排菱形内为橙色圆圈＋"✓"标记
    bg = []
    for key, dnm, fn, cname, glyphs in (
            ('up', DECISION_TEXTS[0], _is_teal, '青绿色', ('？',)),
            ('low', DECISION_TEXTS[1], _is_orange, '橙色', ('✓', '√', '✔'))):
        b = c.badges.get(key)
        if b is None:
            bg.append('菱形"%s"内未见圆形徽标' % dnm)
            continue
        e, m = b
        col = e.lnc if fn(e.lnc) else e.fill
        if not fn(col):
            bg.append('菱形"%s"徽标圆圈色 %s 非%s' % (dnm, col, cname))
        if m is None:
            bg.append('菱形"%s"徽标内缺标记' % dnm)
        elif m.kind != 'pic' and _norm(m.text) not in glyphs:
            bg.append('菱形"%s"徽标标记为 %r' % (dnm, _norm(m.text)))
    emit(0 if bg else 1,
         "2个菱形判定框内各设一个圆形徽标——上排菱形内为青绿色圆圈＋\"？\"标记，下排菱形内为橙色圆圈＋\"✓\"标记",
         '；'.join(bg) or '上排青绿圈＋"？"、下排橙圈＋"✓"均到位')

    # +5：13个环节文字均完整落在所属形状内部——不越出形状边界、不被形状边线切断、
    #     不与形状描边或圆形徽标相互压盖，且不出现三行及以上的挤压换行
    # 严格对齐细则，仅踩细则明确写出的点，不额外约束：
    #   ① 文字范围完整落在所属形状内（越界量≤0.3cm 的估算容差）；
    #   ② 不与圆形徽标压盖；③ 渲染行数≤2 行。
    # 与办公软件对齐：字宽按全角 0.85×字号估算，开启 spAutoFit 的文本框以其高度反推行数，
    #   与软件里实际看到的排版一致（"是否重大问题？"因此判为两行）。
    ov = []
    for t2, v in c.nodes.items():
        nd, lab = v['node'], v['label']
        r = _txt_rect(lab)
        ln = len(_txt_lines(lab))
        if ln >= 3:
            ov.append('"%s"挤压为 %d 行' % (t2, ln))
        o = cont(nd, r)
        if o > 0.3:
            ov.append('"%s"文字越出形状 %.2fcm' % (t2, o))
        if r is not None:
            for e, m in c.badges.values():
                if not _in_shape(nd, e.cx, e.cy):
                    continue
                ox = _ov1(e.vL, e.vR, r[0], r[2])
                oy = _ov1(e.vT, e.vB, r[1], r[3])
                if ox > 0.2 and oy > 0.2:
                    ov.append('"%s"文字与圆形徽标压盖 %.2f×%.2fcm' % (t2, ox, oy))
    emit(0 if (ov or len(c.nodes) != 13) else 5,
         "13个环节文字均完整落在所属形状内部——不越出边界、不被边线切断、不与描边或圆形徽标压盖，且无三行及以上的挤压换行",
         '；'.join(ov[:4]) or '13处文字均在形状内、不压徽标、行数≤2行')

    # +3：环节文字字号≥16磅（未为塞进过小的框体而压缩字号），13处字号统一（极差≤2磅）
    #     并在所属形状内居中对齐
    # 严格对齐细则，仅踩细则明确写出的点，不额外约束：
    #   ① 每处字号≥16磅；② 13 处字号极差≤2磅；③ 在所属形状内居中。
    # 与办公软件对齐："居中"看的是文字连同其图标构成的内容块相对形状中心的偏移
    #   （图标+文字左右并排时，文本框本身必然偏在一侧，故按内容块整体判定，容差0.4cm）。
    fs = []
    szs = [(t2, labels[t2].size or 0.0) for t2 in labels]
    low = ['"%s"%.1f磅' % (k, v) for k, v in szs if v < 16.0]
    if low:
        fs.append('字号不足16磅：%s' % '、'.join(low[:4]))
    if _rng([v for _, v in szs]) > 2.0:
        fs.append('13处字号极差 %.1f磅（需≤2磅）' % _rng([v for _, v in szs]))
    for t2, v in c.nodes.items():
        nd = v['node']
        r = _txt_rect(v['label'])
        if r is None:
            continue
        bl2, bt2, br2, bb2 = r
        g = c.icons.get(t2)
        if g is not None:
            bl2, bt2 = min(bl2, g.vL), min(bt2, g.vT)
            br2, bb2 = max(br2, g.vR), max(bb2, g.vB)
        off = max(abs((bl2 + br2) / 2.0 - nd.cx), abs((bt2 + bb2) / 2.0 - nd.cy))
        if off > 0.4:
            fs.append('"%s"内容块偏离形状中心 %.2fcm' % (t2, off))
    emit(0 if (fs or len(c.nodes) != 13) else 3,
         "环节文字字号≥16磅，13处字号统一（极差≤2磅）并在所属形状内居中对齐",
         '；'.join(fs[:4]) or ('13处字号 %.1f-%.1f磅（极差%.1f磅）、均在形状内居中'
                              % (min(v for _, v in szs), max(v for _, v in szs),
                                 _rng([v for _, v in szs]))
                              if szs else '未识别到环节文字'))

    # +4：图例块存在与内容（细则第25行；由原细则的图例内容 +3 与圆形底衬及标题色 +1 合并而来
    #     ——同为"图例块里该有什么"）
    #     ——左下角设有"相关记录"图例块：圆角矩形描边框内含标题"相关记录："及三条记录
    #     "沟通记录""证据材料""处理报告"，每条前带小圆点项目符号，文字无错漏；块左侧设有
    #     圆形底衬＋记录类图标（剪贴板/文档），圆形直径1.4-2.2cm、图标完整落在圆内；
    #     标题"相关记录："使用蓝色并与三条记录正文的文字颜色明显区分
    # 严格对齐细则，仅踩细则明确写出的点，不额外约束：
    #   ① 有圆角矩形描边框；② 框内含标题与三条记录且文字无错漏；
    #   ③ 每条记录前带小圆点项目符号——独立小圆点图形或段首项目符号字符均可；
    #   ④ 圆形底衬直径 1.4-2.2cm 且记录类图标完整落在圆内；
    #   ⑤ 标题为蓝色且与正文颜色明显区分（RGB 距离≥60）。
    lg = []
    if c.legend is None:
        lg.append('未发现左下角"相关记录"图例块')
    else:
        if c.legend.geom not in ROUND_GEOMS:
            lg.append('图例框为 %s 而非圆角矩形' % c.legend.geom)
        if c.leg_title is None:
            lg.append('缺标题"相关记录："')
        got = [_norm(it[1]) for it in c.leg_items]
        miss = [k for k in LEGEND_ITEMS if not any(k in g for g in got)]
        if miss:
            lg.append('缺记录条目：%s' % '、'.join(miss))
        if len(c.leg_bullets) < 3:
            lg.append('仅 %d/3 条记录带小圆点项目符号' % len(c.leg_bullets))
    if c.leg_circle is None:
        lg.append('图例块内未见圆形底衬')
    else:
        e = c.leg_circle
        dia = max(e.W, e.H)
        if not (1.4 <= dia <= 2.2):
            lg.append('圆形底衬直径 %.2fcm（需1.4-2.2cm）' % dia)
        g = c.leg_icon
        if g is None:
            lg.append('圆形底衬内缺记录类图标')
        elif g is not e and not (e.L - 0.02 <= g.vL and g.vR <= e.R + 0.02
                                 and e.T - 0.02 <= g.vT and g.vB <= e.B + 0.02):
            lg.append('记录图标未完整落在圆内')
    if c.leg_title is not None:
        tc = c.leg_title[3]
        ics = [it[3] for it in c.leg_items if it[3]]
        if not _is_blue(tc):
            lg.append('标题色 %s 非蓝色' % tc)
        elif ics and min(_cdist(tc, v) for v in ics) < 60.0:
            lg.append('标题色 %s 与正文 %s 区分不明显' % (tc, ics[0]))
    emit(0 if lg else 4,
         "左下角设有\"相关记录\"图例块——圆角矩形描边框内含标题\"相关记录：\"及\"沟通记录\"\"证据材料\"\"处理报告\"三条记录、每条前带小圆点项目符号；左侧设圆形底衬（直径1.4-2.2cm）＋记录类图标且图标完整落在圆内；标题使用蓝色并与正文颜色明显区分",
         '；'.join(lg[:4]) or ('圆角矩形图例框内标题与三条记录齐备、每条前带小圆点，圆形底衬 %.2fcm＋记录图标到位，标题色 %s 与正文明显区分'
                              % (max(c.leg_circle.W, c.leg_circle.H), c.leg_title[3])))

    # +4：图例块尺寸与排版得当（细则第26行；由原细则的图例尺寸位置 +3 与图例文字 +1 合并而来
    #     ——同为"图例块摆得对不对"）
    #     ——图例块尺寸与幅面相称（宽≥5.2cm、高≥3.9cm），且位于版面左下角（左边距≤1.2cm、
    #     底端≥16.5cm）；块内文字字号≥16磅，与图例块尺寸相称、四周留白匀齐不贴边
    # 严格对齐细则，仅踩细则明确写出的点，不额外约束：
    #   ① 宽≥5.2cm 且高≥3.9cm；② 左边距≤1.2cm 且底端≥16.5cm；
    #   ③ 标题与三条记录的字号均≥16磅；④ 文字四周不贴边（到图例框内侧留白≥0.1cm）。
    # 与办公软件对齐：圆形底衬这类装饰件常有意压在框边上，留白只考核"文字"。
    lgp = []
    if c.legend is None:
        lgp.append('未发现图例块')
    else:
        s = c.legend
        if s.W < 5.2 or s.H < 3.9:
            lgp.append('图例块 %.2f×%.2fcm（需≥5.2×3.9cm）' % (s.W, s.H))
        if s.L > 1.2:
            lgp.append('左边距 %.2fcm（需≤1.2cm）' % s.L)
        if s.B < 16.5:
            lgp.append('底端 %.2fcm（需≥16.5cm）' % s.B)
    lps = ([c.leg_title] if c.leg_title else []) + c.leg_items
    ssm = ['%.1f磅' % it[2] for it in lps if (it[2] or 0.0) < 16.0]
    if not lps:
        lgp.append('图例块内未识别到文字')
    if ssm:
        lgp.append('字号不足16磅：%s' % '、'.join(ssm[:4]))
    if c.legend is not None:
        for it in lps:
            r = _txt_rect(it[0])
            if r is None:
                continue
            m = min(r[0] - c.legend.L, c.legend.R - r[2], r[1] - c.legend.T, c.legend.B - r[3])
            if m < 0.1:
                lgp.append('%r 处文字贴边（留白 %.2fcm）' % (_norm(it[1])[:6], m))
                break
    emit(0 if lgp else 4,
         "图例块尺寸与排版得当——尺寸与幅面相称（宽≥5.2cm、高≥3.9cm）且位于版面左下角（左边距≤1.2cm、底端≥16.5cm）；块内文字字号≥16磅，四周留白匀齐不贴边",
         '；'.join(lgp[:4]) or ('图例块 %.2f×%.2fcm，左边距 %.2fcm，底端 %.2fcm，文字 %s磅且四周留白匀齐'
                               % (c.legend.W, c.legend.H, c.legend.L, c.legend.B,
                                  '/'.join('%.1f' % it[2] for it in lps))))

    # +3：整页元素互不压盖——环节框、连线、分支标注文字、图例块四类对象的可视范围彼此不重叠，
    #     分支标注文字未压在走线或形状边线上
    # 严格对齐细则，仅踩细则明确写出的点，不额外约束：
    #   ① 环节框彼此、环节框与图例块不重叠（两向重叠均>0.3cm 才算压盖，避开估算误差）；
    #   ② 走线不压在环节框上（同回流走线项的判法：不含框中心且重叠面积≤该框25%）；
    #   ③ 分支标注文字不压在环节框、图例块或走线上。
    # 与办公软件对齐：菱形按棱形真实轮廓算重叠；本项只判前面各项未覆盖的跨类配对——
    #   相邻环节之间的 11 个箭头本就贴着框边留在间隙里（"11处箭头"项已单独考核）、
    #   两条回流走线的穿越已在"分支与回流走线"项考核，均不重复计入；
    #   "返工"按细则第19行要求就标在虚线走线上，故不与该虚线互判压盖。
    ol = []
    au = set(id(a) for a, _ in c.arrow_hits if a is not None)
    au |= set(id(s) for s in (c.reflow_solid, c.reflow_dash) if s is not None)
    lines_all = [s for s in c.shapes
                 if id(s) not in au and not (s.text or '').strip()
                 and (s.kind == 'cxn' or (s.kind == 'sp' and s.geom
                      and ('Connector' in s.geom or s.geom == 'line')))]
    for i in range(len(nodes)):
        for j in range(i + 1, len(nodes)):
            ox, oy = _overlap(nodes[i], nodes[j])
            if ox > 0.3 and oy > 0.3:
                ol.append('环节框"%s"与"%s"压盖 %.2f×%.2fcm'
                          % (nmap.get(id(nodes[i]), ''), nmap.get(id(nodes[j]), ''), ox, oy))
    if c.legend is not None:
        for nd in nodes:
            ox, oy = _overlap(c.legend, nd)
            if ox > 0.3 and oy > 0.3:
                ol.append('图例块与环节框"%s"压盖 %.2f×%.2fcm' % (nmap.get(id(nd), ''), ox, oy))
    for s in lines_all:
        cr = crosses(s, set())
        if cr:
            ol.append('走线 %s 压在环节框 %s 上' % (s.name, '、'.join(cr[:2])))
    for k, lbl in need_mk:
        s = mk.get(k)
        r = _txt_rect(s) if s is not None else None
        if r is None:
            continue
        for nd in nodes + ([c.legend] if c.legend is not None else []):
            xs = _xspan(nd, r[1], r[3])
            if _ov1(r[0], r[2], xs[0], xs[1]) > 0.3 and _ov1(r[1], r[3], nd.vT, nd.vB) > 0.3:
                ol.append('分支标注"%s"压在 %s 边线上' % (_norm(s.text), nd.name))
        if k == 'rework':
            continue
        for s2 in lines_all:
            if _ov1(r[0], r[2], s2.vL, s2.vR) > 0.3 and _ov1(r[1], r[3], s2.vT, s2.vB) > 0.3:
                ol.append('分支标注"%s"压在走线 %s 上' % (_norm(s.text), s2.name))
    emit(0 if ol else 3,
         "整页元素互不压盖——环节框、连线、分支标注文字、图例块四类对象的可视范围彼此不重叠，分支标注文字未压在走线或形状边线上",
         '；'.join(ol[:3]) or '四类对象可视范围彼此不重叠，分支标注未压走线与边线')

    return results, details


# ---------------------------------------------------------------------------
# 维度1 门禁 + 对外接口
# ---------------------------------------------------------------------------
def _dim1_gate(file_path):
    """维度1：交付文件为1个可正常打开的 PowerPoint 文档（.pptx）。
    返回 (pkg, ctx, reason)，reason 非空即门禁不通过。

    门禁只判"能不能打开"这一件事——页面幅面、幻灯片页数、各元素是否可单独点选编辑、
    版面是否错乱等，均不在门禁内考核（对应内容由维度2各条目按分值扣减）。

    与办公软件对齐：PowerPoint 打开 .pptx 时先校验 OPC 包结构与 presentation.xml，
      再按 sldIdLst 读取幻灯片；此处按同一链路检查——包能打开、且页内容能正常解析出来，
      即视为"可正常打开的 .pptx"。
    """
    pkg, err = _open_pkg(file_path)
    if pkg is None:
        return (None, None, err)
    try:
        c = _analyze(pkg)
    except Exception as exc:
        return (pkg, None, '演示文稿内容无法正常解析：%s: %s' % (type(exc).__name__, exc))
    return (pkg, c, '')


def _evaluate_impl(file_path, result):
    """维度一 + 维度二的评估主体；末尾按 _MAX_DELTAS_ORDER 反解成结构化 dim2_items。"""
    pkg, c, reason = _dim1_gate(file_path)
    try:
        if reason:
            # 维度一不通过：dim1_reason 必须写入不通过原因摘要，不得留空。
            result["dim1_pass"] = False
            result["dim1_reason"] = reason
            result["dim2_items"] = []
            result["total_score"] = 0
            result["max_score"] = MAX_SCORE_TOTAL
            return result
        # 门禁通过时 dim1_reason 恒为空字符串——reason 只用于说明"为什么不通过"，
        # 通过即无需说明；不通过时才写入具体原因。
        result["dim1_pass"] = True
        result["dim1_reason"] = ""
        results, _details = _score(c)
    finally:
        if pkg is not None:
            pkg.close()

    # ========== 将 results 反解为结构化 dim2_items ==========
    # 逐项按追加顺序对齐 _MAX_DELTAS_ORDER：解析每条 "+N：说明" 的 N 与 rule；
    # 命中 N=max_delta，未命中 N=0。每项只输出 rule / max_delta / delta / hit 四个字段。
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
    result["max_score"] = MAX_SCORE_TOTAL
    return result


class _InputError(Exception):
    """入参层面的可预期错误（目录不存在、目录中未找到 .pptx）。
    这类错误的 message 本身已足够自解释，写入 error 字段时不加异常类型前缀。"""


def evaluate(dir_path):
    """按约定返回结构化 dict；不 print、不改 sys.stdout、不 sys.exit。
    入参 dir_path 为"脚本所在目录的路径"，脚本自己在该目录里定位并打开被评估的
    .pptx（忽略以 ~$ 开头的 PowerPoint 临时锁文件）。脚本自身崩溃（含目录不存在、
    未找到文档等）由本函数捕获并返回 status='error'，此时 dim1_pass 记为 ""（未判定）、
    total_score 与 max_score 均记 0。"""
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
            raise _InputError("目录不存在：%s" % dir_path)
        cands = sorted(
            f for f in os.listdir(dir_path)
            if f.lower().endswith('.pptx') and not f.startswith('~$')
        )
        if not cands:
            raise _InputError("目录中未找到 .pptx 文件：%s" % dir_path)
        file_path = os.path.join(dir_path, cands[0])
        result["file_name"] = os.path.basename(file_path)
        return _evaluate_impl(file_path, result)
    except Exception as exc:
        # 脚本自身异常（目录不存在、目录中未找到 .pptx、评分过程中途崩溃等）：
        # status=error 且 error 写入摘要。此时门禁根本没跑过，维度一既非通过也非不通过，
        # 故 dim1_pass 记空串表示"未判定"、dim1_reason 留空（原因已在 error 里）；
        # 既然一条都没评，max_score 也一并记 0，不再声称有 92 分可拿。
        result["status"] = "error"
        if isinstance(exc, _InputError):
            result["error"] = str(exc)
        else:
            result["error"] = "%s: %s" % (type(exc).__name__, exc)
        result["dim1_pass"] = ""
        result["dim1_reason"] = ""
        result["dim2_items"] = []
        result["total_score"] = 0
        result["max_score"] = 0
        return result


if __name__ == "__main__":
    # 本地调试入口：evaluate 只接收"脚本所在目录的路径"，由脚本自己在该目录里
    # 定位并打开被评估的 .pptx。命令行第 1 个参数可覆盖目录路径；未指定时使用
    # 脚本自身所在目录。
    # 注：evaluate() 的接口约定就是"入参为目录路径"——路径始终由外部显式传入；
    #     这里的默认值仅为本地调试便利，不影响 evaluate 的接口约定。
    if len(sys.argv) >= 2:
        _dir = sys.argv[1]
    else:
        _dir = os.path.dirname(os.path.abspath(__file__))
    _out = evaluate(_dir)
    print(json.dumps(_out, ensure_ascii=False, indent=2))
