#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# pyright: reportUnknownMemberType=false, reportUnknownVariableType=false, reportUnknownArgumentType=false, reportUnknownParameterType=false, reportMissingParameterType=false, reportPrivateUsage=false, reportUnknownLambdaType=false, reportGeneralTypeIssues=false

import os
import re
import io
import sys
import json
import math
import zlib
import base64
import zipfile
import traceback
from pptx import Presentation
try:
    from PIL import Image
except ImportError:
    Image = None

# 说明：按接口统一约定，本模块不修改全局 sys.stdout，也不 print 主结果。
# 主结果统一由 evaluate(dir_path) -> dict 返回；入参为"脚本所在目录的路径"，
# 由脚本自己在该目录内定位并打开被评估的 .pptx。本文件仅在 __main__ 下作为
# 本地调试入口，才把 dict 序列化为 JSON 打印。
#
# 本脚本对应任务：把 12 张照片按 1—12 的顺序精准放进"核心技术"演示文稿两页
# 四个浅色内容卡片下方原本空着的区域（第 1 页每卡"上一大下两小"、第 2 页每卡
# "左二小右一大"），并给两页的页面主标题加浅蓝色边框。
# 维度二分为 9 条加分项（+，满分 35）与 19 条减分项（-，合计 -33）；加分项来自
# 指令下达的任务，减分项来自指令未涉及却被改动、破坏交付演示文稿既有内容的
# 多余修改。

SCRIPT_ID = "064"
# 维度二加分项 max_delta 之和；用于维度一未通过 / 脚本异常时兜底
MAX_SCORE_TOTAL = 35
# 维度二减分项 max_delta 之和（负值），仅用于自检
MIN_PENALTY_TOTAL = -33

# 维度二所有评分项按"结果追加顺序"的 max_delta 列表；末端把 results 里的
# "+N：..." / "-N：..." 字符串反解为结构化 dim2_items 时用。顺序须与
# _evaluate_impl 中 results.append(...) 的顺序严格一致；共 28 项。
_MAX_DELTAS_ORDER = [
    5,   # 01 12 张照片按 1—12 顺序分入四个卡片
    3,   # 02 每卡三张照片的卡内落位顺序与编号一致
    5,   # 03 第 1 页"上一大下两小"
    5,   # 04 第 2 页"左二小右一大"
    3,   # 05 照片不与角标/卡片小标题/卡片正文文字压叠
    3,   # 06 照片全部落在所属卡片浅色区块边界内
    5,   # 07 两页主标题加浅蓝色矩形边框
    3,   # 08 标题边框贴合标题文字尺寸
    3,   # 09 照片与卡片边缘、照片之间留白均匀且对齐
    -3,  # 10 照片被物理裁切后重新编码写入
    -3,  # 11 新增浅色边框颜色不是浅蓝色/四处不统一
    -3,  # 12 新增边框线条压穿文字
    -1,  # 13 新增或遗留指令未要求的文字
    -1,  # 14 交付形态不正确（文件名被改写/压缩包/PDF）
    -3,  # 15 顶部深色标题条或装饰细条被改动
    -1,  # 16 页面主标题文字或格式被改动
    -1,  # 17 右上角提示语被删改
    -1,  # 18 导语段被删改
    -3,  # 19 四个内容卡片底色/宽度/描边/并排布局被改动
    -1,  # 20 卡片角标被改动
    -1,  # 21 四个卡片小标题被改动
    -3,  # 22 四段卡片正文被删改或格式不一致
    -1,  # 23 卡片底部 1.7 磅彩色横线被改动
    -3,  # 24 8 组 KPI 数字与说明被删改
    -1,  # 25 页面尺寸或幻灯片数量被改动
    -1,  # 26 备注页被删除或被写入排版说明
    -1,  # 27 照片被加旋转/描边/阴影/发光/圆角裁剪/动画/超链接
    -1,  # 28 文档属性标题被改写或新增无关自定义属性
]
assert sum(d for d in _MAX_DELTAS_ORDER if d > 0) == MAX_SCORE_TOTAL
assert sum(d for d in _MAX_DELTAS_ORDER if d < 0) == MIN_PENALTY_TOTAL

TARGET_NAME = "核心技术.pptx"

EMU_CM = 360000.0          # 1 cm = 360000 EMU
EMU_PT = 12700.0           # 1 pt  = 12700 EMU
SLIDE_W = 12191365         # 33.87 cm
SLIDE_H = 6858000          # 19.05 cm

NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main"
NS_P = "http://schemas.openxmlformats.org/presentationml/2006/main"
NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


def _a(tag):
    return "{%s}%s" % (NS_A, tag)


def _p(tag):
    return "{%s}%s" % (NS_P, tag)


# ---------------------------------------------------------------------------
# 交付演示文稿既有内容的基线（用于全部减分项的"是否被改动"判定）
# 键为幻灯片序号（1 基）。坐标/尺寸单位 EMU，字号单位磅，颜色为 srgbClr 值。
# ---------------------------------------------------------------------------
BAR_OFF = (0, 0)
BAR_EXT = (12191695, 786384)
STRIP_OFF = (0, 786384)
STRIP_EXT = (12191695, 50292)
TITLE_OFF = (438912, 164592)
TITLE_FONT = "Noto Sans CJK SC"
TAGLINE_LEFT = 8229600
TAGLINE_EXT = (3429000, 256032)
LEAD_EXT = (10972800, 475488)
CARD_LEFT = (640080, 6236335)
CARD_CX = 5321935
BADGE_EXT = (1143000, 256032)
HEAD_EXT = (4809744, 402336)
BODY_CX = 4809744
LINE_CX = 4809744
LINE_W_EMU = 21590        # 1.7 磅
KPI_CX = 1828800
KPI_LEFT = (1078865, 3703320, 6327775, 8952230)

EXPECT = {
    1: {
        "strip": "5EEAD4",
        "title": "高效动力链方案",
        "tagline": "结构减重 · 传动效率 · 耐久验证",
        "lead": ("以更轻的传动单元和更稳定的材料连接方式，提升小型赛车平台的动力响应与可靠性。",
                 16.0, "DDE9F6"),
        "cards": ("E0F2FE", "FEF3C7"),
        "badges": (("MODULE A", "38BDF8"), ("MODULE B", "F59E0B")),
        "heads": ("智能变速总成", "高强复合半轴"),
        "bodies": (
            "将传统滑动导向改为低阻尼滚柱导轨，并采用复合包覆传动带。新结构降低局部发热，扩大速比调节窗口，使起步和中高速巡航的动力输出更加平顺。",
            "采用碳纤维多向铺层与分段固化工艺，优化花键端过渡区域的受力路径。连接面通过韧性胶膜与金属嵌件协同承载，兼顾轻量化和扭转载荷稳定性。",
        ),
        "line_y": 5913501,
        "line_clr": ("38BDF8", "F59E0B"),
        "kpi": (("18%", "5EEAD4", "传动损耗下降"), ("52%", "38BDF8", "总成质量减轻"),
                ("32%", "F59E0B", "接合强度提升"), ("5.5h", "A78BFA", "耐久循环增加")),
    },
    2: {
        "strip": "F59E0B",
        "title": "整车性能协同优化",
        "tagline": "气动布局 · 人机空间 · 安全体验",
        "lead": ("围绕尾部气流、驾驶姿态与防护结构进行协同设计，让车辆在高速稳定性和操作舒适度之间取得更好的平衡。",
                 15.5, "E2E8F0"),
        "cards": ("DCFCE7", "F3E8FF"),
        "badges": (("AERO", "22C55E"), ("COCKPIT", "A78BFA")),
        "heads": ("尾部导流组件", "座舱轻量布局"),
        "bodies": (
            "以双曲率翼面、短端板和尾部整流片形成连续导流通道，减少分离涡的能量损失。高速工况下车尾压力恢复更稳定，方向修正更轻。",
            "将踏板组、转向柱和防滚架交汇点重新布置，保留关键防护空间的同时压缩冗余支架。驾驶员腿部活动范围更大，进出车动作更自然。",
        ),
        "line_y": 5970524,
        "line_clr": ("22C55E", "A78BFA"),
        "kpi": (("24%", "22C55E", "尾涡强度降低"), ("0.34", "38BDF8", "目标阻力系数"),
                ("120mm", "F59E0B", "腿部空间增加"), ("15%", "A78BFA", "框架利用率提升")),
    },
}

# 卡片正文统一字号/颜色；角标、小标题、KPI 的字号
BODY_SZ, BODY_CLR = 12.2, "475569"
BADGE_SZ, HEAD_SZ, HEAD_CLR = 9.5, 18.0, "172033"
KPI_NUM_SZ, KPI_CAP_SZ, KPI_CAP_CLR = 19.0, 8.5, "94A3B8"
BAR_CLR = "0B1220"
DOC_TITLE = "高效移动平台技术方案"
CUSTOM_PROPS = ("KSOProductBuildVer", "ICV")

# 12 张待插入照片的基线：像素尺寸一致，均为 1448x1086
PHOTO_PX = (1448, 1086)
PHOTO_AR = 1448 / 1086.0

# 12 张待插入照片的灰度指纹：每张按 48x36 网格重采样后取高 4 位，两像素合一字节
# （864 字节/张，共 10368 字节），zlib 压缩后 base64 存放。用于在不依赖素材目录的
# 情况下判断"插入的是第几张照片"，并容忍缩放、重新编码与居中裁切。
_SIG_GW, _SIG_GH = 48, 36
_PHOTO_SIG_B64 = """
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
"""

# ---------------------------------------------------------------------------
# 基础工具：单位换算与容差
# ---------------------------------------------------------------------------


def _cm(emu):
    return emu / EMU_CM


def _near(a, b, tol_cm=0.05):
    return abs(a - b) <= tol_cm * EMU_CM


def _pt_of(emu_w):
    return emu_w / EMU_PT


# ---------------------------------------------------------------------------
# 基础工具：DrawingML 属性读取（直接读元素，避免不同 python-pptx 版本的封装差异）
# ---------------------------------------------------------------------------


def _sp_pr(el):
    """取形状的 spPr（p:sp / p:pic 通用）。"""
    for tag in ("spPr", "grpSpPr"):
        node = el.find(_p(tag))
        if node is not None:
            return node
    return None


def _prst_geom(el):
    """返回 (prst 名称, adj 值或 None)；无 prstGeom 返回 (None, None)。"""
    sppr = _sp_pr(el)
    if sppr is None:
        return None, None
    geom = sppr.find(_a("prstGeom"))
    if geom is None:
        return None, None
    adj = None
    node = geom.find(_a("avLst"))
    if node is not None:
        gd = node.find(_a("gd"))
        if gd is not None:
            fmla = gd.get("fmla") or ""
            m = re.search(r"val\s+(-?\d+)", fmla)
            if m:
                adj = int(m.group(1))
    return geom.get("prst"), adj


def _solid_rgb(node):
    """从任意容器节点里取 a:solidFill/a:srgbClr 的 (颜色, alpha 百分比)。"""
    if node is None:
        return None, 100.0
    fill = node.find(_a("solidFill"))
    if fill is None:
        return None, 100.0
    clr = fill.find(_a("srgbClr"))
    if clr is None:
        return None, 100.0
    alpha = 100.0
    a_node = clr.find(_a("alpha"))
    if a_node is not None and a_node.get("val"):
        alpha = int(a_node.get("val")) / 1000.0
    return (clr.get("val") or "").upper(), alpha


def _fill_rgb(el):
    """形状填充色；无填充或非纯色返回 None。"""
    sppr = _sp_pr(el)
    if sppr is None:
        return None
    if sppr.find(_a("noFill")) is not None:
        return None
    rgb, alpha = _solid_rgb(sppr)
    if rgb is None or alpha < 5.0:
        return None
    return rgb


def _line_of(el):
    """形状轮廓：返回 dict(w=EMU 宽度, rgb=颜色, alpha=不透明度, visible=是否可见)。"""
    sppr = _sp_pr(el)
    out = {"w": None, "rgb": None, "alpha": 0.0, "visible": False}
    if sppr is None:
        return out
    ln = sppr.find(_a("ln"))
    if ln is None:
        return out
    out["w"] = int(ln.get("w")) if ln.get("w") else None
    if ln.find(_a("noFill")) is not None:
        return out
    rgb, alpha = _solid_rgb(ln)
    out["rgb"], out["alpha"] = rgb, alpha
    out["visible"] = bool(rgb) and alpha >= 5.0
    return out


def _has_effect(el):
    """是否带阴影/发光/映像等效果。"""
    sppr = _sp_pr(el)
    if sppr is None:
        return False
    for tag in ("effectLst", "effectDag"):
        node = sppr.find(_a(tag))
        if node is not None and len(node) > 0:
            return True
    return False


def _rot_of(el):
    sppr = _sp_pr(el)
    if sppr is None:
        return 0
    xfrm = sppr.find(_a("xfrm"))
    if xfrm is None:
        return 0
    return int(xfrm.get("rot") or 0)

# ---------------------------------------------------------------------------
# 基础工具：文本读取
# ---------------------------------------------------------------------------


def _tx_body(el):
    return el.find(_p("txBody"))


def _paras(el):
    body = _tx_body(el)
    if body is None:
        return []
    return body.findall(_a("p"))


def _para_text(para):
    out = []
    for node in para.iter():
        if node.tag == _a("t"):
            out.append(node.text or "")
        elif node.tag == _a("br"):
            out.append("\n")
    return "".join(out)


def _text_of(el):
    return "\n".join(_para_text(x) for x in _paras(el))


def _runs(el):
    return [r for para in _paras(el) for r in para.findall(_a("r"))]


def _run_font(run):
    """返回 (字号磅, 是否加粗, 颜色, 中文字体名)。缺省返回 None 占位。"""
    rpr = run.find(_a("rPr"))
    if rpr is None:
        return None, None, None, None
    sz = int(rpr.get("sz")) / 100.0 if rpr.get("sz") else None
    bold = rpr.get("b")
    bold = None if bold is None else bold in ("1", "true")
    rgb, _alpha = _solid_rgb(rpr)
    face = None
    for tag in ("ea", "latin"):
        node = rpr.find(_a(tag))
        if node is not None and node.get("typeface"):
            face = node.get("typeface")
            if tag == "ea":
                break
    return sz, bold, rgb, face


def _first_font(el):
    for run in _runs(el):
        return _run_font(run)
    return None, None, None, None


def _max_sz(el, default=18.0):
    best = None
    for run in _runs(el):
        sz = _run_font(run)[0]
        if sz is not None and (best is None or sz > best):
            best = sz
    return default if best is None else best


def _algn_of(el):
    for para in _paras(el):
        ppr = para.find(_a("pPr"))
        if ppr is not None and ppr.get("algn"):
            return ppr.get("algn")
    return "l"


def _line_spacing_pct(el):
    """段落行距百分比；未设置返回 None。"""
    for para in _paras(el):
        ppr = para.find(_a("pPr"))
        if ppr is None:
            continue
        lnspc = ppr.find(_a("lnSpc"))
        if lnspc is None:
            continue
        pct = lnspc.find(_a("spcPct"))
        if pct is not None and pct.get("val"):
            return int(pct.get("val")) / 1000.0
    return None


def _body_pr(el):
    """返回 dict(anchor, lIns, tIns, rIns, bIns, wrap, vert)。"""
    body = _tx_body(el)
    out = {"anchor": "t", "lIns": 91440, "tIns": 45720, "rIns": 91440,
           "bIns": 45720, "wrap": "square", "vert": None}
    if body is None:
        return out
    bpr = body.find(_a("bodyPr"))
    if bpr is None:
        return out
    if bpr.get("anchor"):
        out["anchor"] = bpr.get("anchor")
    for key in ("lIns", "tIns", "rIns", "bIns"):
        if bpr.get(key) is not None:
            out[key] = int(bpr.get(key))
    if bpr.get("wrap"):
        out["wrap"] = bpr.get("wrap")
    out["vert"] = bpr.get("vert")
    return out

# ---------------------------------------------------------------------------
# 形状包装：把（可能嵌在组合里的）形状统一成带绝对坐标的轻量对象
# ---------------------------------------------------------------------------


class _Sh(object):
    __slots__ = ("el", "name", "kind", "l", "t", "w", "h", "grouped", "part")

    def __init__(self, el, name, kind, l, t, w, h, grouped, part):
        self.el, self.name, self.kind = el, name, kind
        self.l, self.t, self.w, self.h = int(l), int(t), int(w), int(h)
        self.grouped, self.part = grouped, part

    @property
    def r(self):
        return self.l + self.w

    @property
    def b(self):
        return self.t + self.h

    @property
    def cx(self):
        return self.l + self.w // 2

    @property
    def cy(self):
        return self.t + self.h // 2

    @property
    def box(self):
        return (self.l, self.t, self.r, self.b)

    @property
    def text(self):
        return _text_of(self.el)


def _xfrm_of(el):
    sppr = _sp_pr(el)
    if sppr is None:
        return None
    return sppr.find(_a("xfrm"))


def _off_ext(xfrm):
    off = xfrm.find(_a("off")) if xfrm is not None else None
    ext = xfrm.find(_a("ext")) if xfrm is not None else None
    if off is None or ext is None:
        return None
    return (int(off.get("x") or 0), int(off.get("y") or 0),
            int(ext.get("cx") or 0), int(ext.get("cy") or 0))


def _nv_name(el):
    for tag in ("nvSpPr", "nvPicPr", "nvGrpSpPr", "nvCxnSpPr"):
        nv = el.find(_p(tag))
        if nv is None:
            continue
        cnv = nv.find(_p("cNvPr"))
        if cnv is not None:
            return cnv.get("name") or ""
    return ""


def _walk(container, part, out, ox=0, oy=0, sx=1.0, sy=1.0, grouped=False):
    """深度遍历 spTree，展开组合并把子形状坐标换算到幻灯片绝对坐标系。"""
    for el in container:
        tag = el.tag
        if tag == _p("grpSp"):
            xfrm = _xfrm_of(el)
            oe = _off_ext(xfrm)
            if oe is None:
                _walk(el, part, out, ox, oy, sx, sy, True)
                continue
            gx, gy, gcx, gcy = oe
            choff = xfrm.find(_a("chOff"))
            chext = xfrm.find(_a("chExt"))
            cox = int(choff.get("x") or 0) if choff is not None else 0
            coy = int(choff.get("y") or 0) if choff is not None else 0
            ccx = int(chext.get("cx") or 0) if chext is not None else gcx
            ccy = int(chext.get("cy") or 0) if chext is not None else gcy
            nsx = sx * (gcx / float(ccx) if ccx else 1.0)
            nsy = sy * (gcy / float(ccy) if ccy else 1.0)
            nox = ox + sx * gx - nsx * cox
            noy = oy + sy * gy - nsy * coy
            _walk(el, part, out, nox, noy, nsx, nsy, True)
            continue
        if tag == _p("pic"):
            kind = "pic"
        elif tag in (_p("sp"), _p("cxnSp")):
            kind = "sp"
        elif tag == _p("graphicFrame"):
            kind = "frame"
        else:
            continue
        oe = _off_ext(_xfrm_of(el))
        if oe is None and tag == _p("graphicFrame"):
            xfrm = el.find(_p("xfrm"))
            oe = _off_ext(xfrm)
        if oe is None:
            oe = (0, 0, 0, 0)
        x, y, cx_, cy_ = oe
        out.append(_Sh(el, _nv_name(el), kind,
                       ox + sx * x, oy + sy * y, sx * cx_, sy * cy_,
                       grouped, part))


def _shapes_of(slide):
    """返回该页所有形状（按 z 序、组合已展开）。"""
    tree = slide.shapes._spTree
    out = []
    _walk(tree, slide.part, out)
    return out

# ---------------------------------------------------------------------------
# 矩形运算与"文字实际占位"估算
# ---------------------------------------------------------------------------


def _inter(a, b):
    """两矩形交集 (dx, dy)；不相交时为 0。"""
    dx = min(a[2], b[2]) - max(a[0], b[0])
    dy = min(a[3], b[3]) - max(a[1], b[1])
    return max(0, dx), max(0, dy)


def _overlap_area(a, b):
    dx, dy = _inter(a, b)
    return dx * dy


def _hit_overlap(a, b, tol_cm=0.1):
    """双向重叠都超过 tol_cm 才算压叠，避免边缘相切被误判。"""
    dx, dy = _inter(a, b)
    return dx > tol_cm * EMU_CM and dy > tol_cm * EMU_CM


def _char_units(text):
    """按"全角=1、半角=0.5"折算一行的字符宽度单位。"""
    total = 0.0
    for ch in text:
        if ch in ("\r", "\n"):
            continue
        total += 0.5 if ord(ch) < 0x2E80 else 1.0
    return total


def _text_rect(sh):
    """估算文字在页面上实际占据的矩形（而非文本框外框）。

    版式类减分/压叠类加分都以文字实际占位为准：模型为腾位置把文本框改矮或
    照片压到空文本框的下半部分，都不应被判成"压住文字"。
    """
    text = sh.text
    if not text.strip():
        return None
    bpr = _body_pr(sh.el)
    sz = _max_sz(sh.el)
    usable_w = max(1, sh.w - bpr["lIns"] - bpr["rIns"])
    per_char = sz * EMU_PT
    cpl = max(1.0, usable_w / per_char)
    lines, widest = 0, 0.0
    for para in _paras(sh.el):
        units = _char_units(_para_text(para))
        if bpr["wrap"] == "none":
            n = 1
            widest = max(widest, units)
        else:
            n = max(1, int(math.ceil(units / cpl))) if units else 1
            widest = max(widest, min(units, cpl))
        lines += n
    spc = _line_spacing_pct(sh.el)
    factor = 1.5 if spc is None else max(1.0, 1.5 * spc / 100.0)
    est_h = lines * per_char * factor
    est_w = min(widest * per_char, usable_w) if bpr["wrap"] != "none" \
        else widest * per_char
    anchor = bpr["anchor"]
    if anchor in ("ctr", "mid"):
        y0 = sh.t + (sh.h - est_h) / 2.0
    elif anchor == "b":
        y0 = sh.b - bpr["bIns"] - est_h
    else:
        y0 = sh.t + bpr["tIns"]
    algn = _algn_of(sh.el)
    if algn == "ctr":
        x0 = sh.l + bpr["lIns"] + (usable_w - est_w) / 2.0
    elif algn == "r":
        x0 = sh.r - bpr["rIns"] - est_w
    else:
        x0 = sh.l + bpr["lIns"]
    return (int(x0), int(y0), int(x0 + est_w), int(y0 + est_h))


def _stroke_points(sh, step_cm=0.04):
    """采样形状轮廓线的路径点；矩形与圆角矩形（含圆角圆弧）均支持。

    用于判定"新增边框线条是否压穿了文字"：包围式边框的描边永远不会落在被
    包围文字的矩形内，只有位置/圆角不当的边框才会。
    """
    prst, adj = _prst_geom(sh.el)
    l, t, r, b = sh.box
    rad = 0
    if prst in ("roundRect", "round2SameRect", "round1Rect", "snipRoundRect"):
        val = 16667 if adj is None else max(0, min(50000, adj))
        rad = int(min(sh.w, sh.h) * val / 100000.0)
    step = max(1, int(step_cm * EMU_CM))
    pts = []
    for x in range(l + rad, r - rad + 1, step):
        pts.append((x, t))
        pts.append((x, b))
    for y in range(t + rad, b - rad + 1, step):
        pts.append((l, y))
        pts.append((r, y))
    if rad > 0:
        for k in range(0, 13):
            ang = math.pi / 2.0 * k / 12.0
            dx, dy = rad * math.cos(ang), rad * math.sin(ang)
            pts.append((l + rad - dx, t + rad - dy))
            pts.append((r - rad + dx, t + rad - dy))
            pts.append((l + rad - dx, b - rad + dy))
            pts.append((r - rad + dx, b - rad + dy))
    return pts


def _pt_in_rect(pt, rect, margin=0):
    return (rect[0] + margin <= pt[0] <= rect[2] - margin
            and rect[1] + margin <= pt[1] <= rect[3] - margin)

# ---------------------------------------------------------------------------
# 照片识别：判断每张插入的图片对应素材 1—12 中的哪一张
# ---------------------------------------------------------------------------

_OUT_W, _OUT_H = 16, 12         # 比对时统一降采样到的小网格
_SIG_ACCEPT = 0.62              # 判为"同一张素材"的距离上限（实测最差 0.43）
_SIG_CACHE = []


def _resample(im, w, h):
    flt = getattr(getattr(Image, "Resampling", Image), "BILINEAR")
    return im.convert("L").resize((w, h), flt)


def _ref_sigs():
    """解出 12 张素材的 48x36 灰度指纹（4 位量化值列表）。"""
    if _SIG_CACHE:
        return _SIG_CACHE[0]
    raw = zlib.decompress(base64.b64decode(re.sub(r"\s+", "", _PHOTO_SIG_B64)))
    per = _SIG_GW * _SIG_GH // 2
    sigs = {}
    for i in range(12):
        chunk = raw[i * per:(i + 1) * per]
        vals = []
        for byte in chunk:
            vals.append(byte >> 4)
            vals.append(byte & 0x0F)
        sigs[i + 1] = vals
    _SIG_CACHE.append(sigs)
    return sigs


def _sub_sig(grid, fw, fh):
    """从 48x36 参考网格中取中心裁切区域，重采样成 16x12，用于对齐候选图。"""
    cw = max(1, int(round(_SIG_GW * fw)))
    ch = max(1, int(round(_SIG_GH * fh)))
    x0, y0 = (_SIG_GW - cw) // 2, (_SIG_GH - ch) // 2
    out = []
    for j in range(_OUT_H):
        sy = min(_SIG_GH - 1, y0 + int((j + 0.5) * ch / _OUT_H))
        for i in range(_OUT_W):
            sx = min(_SIG_GW - 1, x0 + int((i + 0.5) * cw / _OUT_W))
            out.append(grid[sy * _SIG_GW + sx])
    return out


def _norm(vec):
    mean = sum(vec) / float(len(vec))
    std = (sum((v - mean) ** 2 for v in vec) / float(len(vec))) ** 0.5 or 1.0
    return [(v - mean) / std for v in vec]


def _sig_dist(a, b):
    a, b = _norm(a), _norm(b)
    return sum(abs(x - y) for x, y in zip(a, b)) / len(a)


def _pic_blob(sh):
    """取图片二进制；取不到返回 None。"""
    embed = None
    blip = sh.el.find(_p("blipFill"))
    if blip is not None:
        node = blip.find(_a("blip"))
        if node is not None:
            embed = node.get("{%s}embed" % NS_R)
    if not embed:
        return None
    try:
        return sh.part.related_part(embed).blob
    except Exception:
        pass
    try:
        return sh.part.related_parts[embed].blob
    except Exception:
        return None


def _pic_info(sh):
    """返回 dict(px=(w,h), cand=16x12 指纹, blob_len)；失败时 px 为 None。"""
    info = {"px": None, "cand": None, "blob_len": 0, "crop": False}
    blob = _pic_blob(sh)
    if not blob or Image is None:
        return info
    info["blob_len"] = len(blob)
    try:
        im = Image.open(io.BytesIO(blob))
        im.load()
    except Exception:
        return info
    info["px"] = im.size
    small = _resample(im, _OUT_W, _OUT_H)
    info["cand"] = [v >> 4 for v in small.tobytes()]
    return info


def _assign_photos(pics):
    """把一批图片与素材 1—12 做一对一匹配（贪心取全局最小代价）。

    返回 {图片下标: 素材编号}；代价大于阈值的图片视为"不属于这 12 张素材"，
    不参与匹配。
    """
    sigs = _ref_sigs()
    cost = {}
    for idx, info in enumerate(pics):
        if not info.get("cand"):
            continue
        pw, ph = info["px"]
        ar = pw / float(ph)
        if ar >= PHOTO_AR:
            fw, fh = 1.0, PHOTO_AR / ar
        else:
            fw, fh = ar / PHOTO_AR, 1.0
        for k in sigs:
            cost[(idx, k)] = _sig_dist(info["cand"], _sub_sig(sigs[k], fw, fh))
    result = {}
    used_pic, used_src = set(), set()
    for (idx, k), c in sorted(cost.items(), key=lambda kv: kv[1]):
        if idx in used_pic or k in used_src or c > _SIG_ACCEPT:
            continue
        result[idx] = k
        used_pic.add(idx)
        used_src.add(k)
    return result

# ---------------------------------------------------------------------------
# 版面定位：把每页形状对应到既有的 23 个角色，其余视为"新增形状"
# ---------------------------------------------------------------------------

def _norm_text(s):
    return re.sub(r"\s+", " ", (s or "")).strip()


def _locate_slide(idx, shapes):
    """返回 dict：角色名 -> _Sh；另含 pics / added / all。

    先按文字精确匹配（文字未被改写时最可靠，能容忍模型为腾位置做的大幅移动），
    再对文字已被改写的角色按"宽度 + 左右半区"补配，剩下的就是新增形状。
    """
    exp = EXPECT[idx]
    loc = {"pics": [sh for sh in shapes if sh.kind == "pic"],
           "all": shapes, "added": []}
    rest = [sh for sh in shapes if sh.kind != "pic"]
    taken = set()

    def claim(role, sh):
        if sh is None or role in loc:
            return
        loc[role] = sh
        taken.add(id(sh))

    # 1) 文字精确匹配
    by_text = {}
    for sh in rest:
        by_text.setdefault(_norm_text(sh.text), []).append(sh)

    def pick(text, left_half=None):
        for sh in by_text.get(_norm_text(text), []):
            if id(sh) in taken:
                continue
            if left_half is True and sh.cx > SLIDE_W // 2:
                continue
            if left_half is False and sh.cx < SLIDE_W // 2:
                continue
            return sh
        return None

    claim("title", pick(exp["title"]))
    claim("tagline", pick(exp["tagline"]))
    claim("lead", pick(exp["lead"][0]))
    for k, side in ((0, True), (1, False)):
        suffix = "L" if k == 0 else "R"
        claim("badge" + suffix, pick(exp["badges"][k][0], side))
        claim("head" + suffix, pick(exp["heads"][k], side))
        claim("body" + suffix, pick(exp["bodies"][k], side))
    for k in range(4):
        num, _clr, cap = exp["kpi"][k]
        claim("kpiN%d" % k, pick(num))
        claim("kpiC%d" % k, pick(cap))

    # 2) 几何补配：条形/细条/卡片/横线按位置与尺寸认定，文本框按宽度认定
    def geom(pred, key=None):
        cands = [sh for sh in rest if id(sh) not in taken and pred(sh)]
        if not cands:
            return None
        return sorted(cands, key=key)[0] if key else cands[0]

    claim("bar", geom(lambda s: _near(s.l, BAR_OFF[0], 0.2)
                      and _near(s.t, BAR_OFF[1], 0.2)
                      and _near(s.h, BAR_EXT[1], 0.3)
                      and s.w > SLIDE_W * 0.8))
    claim("strip", geom(lambda s: _near(s.t, STRIP_OFF[1], 0.3)
                        and s.h <= 0.4 * EMU_CM and s.w > SLIDE_W * 0.8))
    for suffix, x in (("L", CARD_LEFT[0]), ("R", CARD_LEFT[1])):
        claim("card" + suffix,
              geom(lambda s, x=x: not s.text.strip() and _near(s.w, CARD_CX, 0.3)
                   and _near(s.l, x, 0.4) and s.h > 3.0 * EMU_CM))
    for suffix, x in (("L", CARD_LEFT[0]), ("R", CARD_LEFT[1])):
        claim("line" + suffix,
              geom(lambda s, x=x: s.h <= 0.1 * EMU_CM and _near(s.w, LINE_CX, 0.4)
                   and abs(s.cx - (x + CARD_CX // 2)) < 2.0 * EMU_CM))
    claim("title", geom(lambda s: s.text.strip() and s.t < 700000
                        and s.l < 2000000,
                        key=lambda s: -_max_sz(s.el)))
    claim("tagline", geom(lambda s: s.text.strip()
                          and _near(s.w, TAGLINE_EXT[0], 0.5) and s.t < 900000))
    claim("lead", geom(lambda s: s.text.strip()
                       and _near(s.w, LEAD_EXT[0], 0.5)))
    for suffix, side in (("L", True), ("R", False)):
        half = (lambda s: s.cx < SLIDE_W // 2) if side else \
            (lambda s: s.cx > SLIDE_W // 2)
        claim("badge" + suffix,
              geom(lambda s, h=half: s.text.strip() and h(s)
                   and _near(s.w, BADGE_EXT[0], 0.5)))
        claim("head" + suffix,
              geom(lambda s, h=half: s.text.strip() and h(s)
                   and _near(s.w, HEAD_EXT[0], 0.5)
                   and _near(s.h, HEAD_EXT[1], 0.4)))
        claim("body" + suffix,
              geom(lambda s, h=half: s.text.strip() and h(s)
                   and _near(s.w, BODY_CX, 0.5),
                   key=lambda s: -len(s.text)))
    for k in range(4):
        x = KPI_LEFT[k]
        for role, top_first in (("kpiN%d" % k, True), ("kpiC%d" % k, False)):
            claim(role, geom(lambda s, x=x: s.text.strip()
                             and _near(s.w, KPI_CX, 0.5) and _near(s.l, x, 0.6)
                             and s.t > 15.0 * EMU_CM,
                             key=(lambda s: s.t) if top_first else
                             (lambda s: -s.t)))

    loc["added"] = [sh for sh in rest if id(sh) not in taken]
    return loc

CARD_TOP = {1: 1783080, 2: 1847215}
CARD_CY = {1: 4250055, 2: 4322445}


def _card_box(loc, idx, side):
    """卡片浅色区块的矩形；卡片形状缺失时用标定位置兜底。"""
    sh = loc.get("card" + side)
    if sh is not None:
        return sh.box
    x = CARD_LEFT[0 if side == "L" else 1]
    return (x, CARD_TOP[idx], x + CARD_CX, CARD_TOP[idx] + CARD_CY[idx])


def _group_photos(prs):
    """把两页的图片按"所属卡片"归组，并识别每张图对应的素材编号与卡内几何角色。

    返回 dict：(页序, 'L'/'R') -> dict(pics=[...], seq=[...], ok_shape=bool)
    其中 pics 为 (照片 _Sh, 素材编号或 None) 列表，seq 为按版式要求的落位顺序。
    """
    slides = list(prs.slides)
    locs, all_pics = [], []
    for i, slide in enumerate(slides[:2], 1):
        loc = _locate_slide(i, _shapes_of(slide))
        locs.append(loc)
        for sh in loc["pics"]:
            all_pics.append((i, sh))
    infos = [_pic_info(sh) for _i, sh in all_pics]
    src_map = _assign_photos(infos)
    groups = {}
    for i in (1, 2):
        loc = locs[i - 1] if len(locs) >= i else {}
        for side in ("L", "R"):
            groups[(i, side)] = {"pics": [], "seq": [], "shape_ok": False,
                                 "box": _card_box(loc, i, side)}
    for pos, (i, sh) in enumerate(all_pics):
        boxes = [(side, _card_box(locs[i - 1], i, side)) for side in ("L", "R")]
        hit = None
        for side, box in boxes:
            if box[0] <= sh.cx <= box[2] and box[1] <= sh.cy <= box[3]:
                hit = side
                break
        if hit is None:
            hit = min(boxes, key=lambda kv: abs(sh.cx - (kv[1][0] + kv[1][2]) / 2))[0]
        groups[(i, hit)]["pics"].append((sh, src_map.get(pos), infos[pos]))
    # 卡内几何角色：第 1 页"上一大下两小"，第 2 页"左二小右一大"
    for (i, side), grp in groups.items():
        items = grp["pics"]
        if len(items) != 3:
            continue
        big = max(items, key=lambda it: it[0].w * it[0].h)
        others = [it for it in items if it is not big]
        if i == 1:
            others.sort(key=lambda it: it[0].l)
            gap_ok = _inter(others[0][0].box, others[1][0].box)[0] == 0
            row_ok = _inter(others[0][0].box, others[1][0].box)[1] > \
                0.5 * min(others[0][0].h, others[1][0].h)
            above = big[0].b <= min(o[0].t for o in others) + 0.05 * EMU_CM
            grp["shape_ok"] = bool(gap_ok and row_ok and above)
            grp["seq"] = [big, others[0], others[1]]
        else:
            others.sort(key=lambda it: it[0].t)
            gap_ok = _inter(others[0][0].box, others[1][0].box)[1] == 0
            col_ok = _inter(others[0][0].box, others[1][0].box)[0] > \
                0.5 * min(others[0][0].w, others[1][0].w)
            right = big[0].l >= max(o[0].r for o in others) - 0.05 * EMU_CM
            grp["shape_ok"] = bool(gap_ok and col_ok and right)
            grp["seq"] = [others[0], others[1], big]
    return locs, groups, infos, src_map

# ---------------------------------------------------------------------------
# 颜色判定与标题边框识别
# ---------------------------------------------------------------------------


def _hsl(rgb):
    r, g, b = (int(rgb[0:2], 16) / 255.0, int(rgb[2:4], 16) / 255.0,
               int(rgb[4:6], 16) / 255.0)
    mx, mn = max(r, g, b), min(r, g, b)
    lum, diff = (mx + mn) / 2.0, mx - mn
    if diff < 1e-6:
        return 0.0, 0.0, lum
    sat = diff / (2.0 - mx - mn) if lum > 0.5 else diff / (mx + mn)
    if mx == r:
        hue = 60.0 * (((g - b) / diff) % 6)
    elif mx == g:
        hue = 60.0 * (2 + (b - r) / diff)
    else:
        hue = 60.0 * (4 + (r - g) / diff)
    return hue, sat, lum


def _is_light_blue(rgb):
    """浅蓝色判定：色相落在青蓝区间、明度偏亮。

    可接受 87CEFA / ADD8E6 / 9EDAF5 / 38BDF8 一类浅蓝，排除粉色 F6B6D2、
    纯蓝 0000FF、青绿 5EEAD4 与灰白。
    """
    if not rgb or not re.match(r"^[0-9A-Fa-f]{6}$", rgb):
        return False
    hue, sat, lum = _hsl(rgb.upper())
    return 185.0 <= hue <= 255.0 and sat >= 0.15 and lum >= 0.55


def _title_border(loc):
    """识别页面主标题的边框：可以是标题形状自带轮廓，也可以是外套的边框形状。

    返回 dict(kind, box, rgb, w, sh) 或 None。
    """
    title = loc.get("title")
    if title is None:
        return None
    ln = _line_of(title.el)
    prst = _prst_geom(title.el)[0]
    if ln["visible"]:
        return {"kind": "ln", "box": title.box, "rgb": ln["rgb"],
                "w": ln["w"] or 9525, "prst": prst or "rect", "sh": title}
    trect = _text_rect(title) or title.box
    best = None
    for sh in loc["added"]:
        if sh.text.strip():
            continue
        line = _line_of(sh.el)
        if not line["visible"]:
            continue
        if not (sh.l <= trect[0] and sh.t <= trect[1]
                and sh.r >= trect[2] and sh.b >= trect[3]):
            continue
        if sh.w * sh.h > 0.25 * SLIDE_W * SLIDE_H:
            continue
        if best is None or sh.w * sh.h < best.w * best.h:
            best = sh
    if best is None:
        return None
    return {"kind": "shape", "box": best.box, "rgb": _line_of(best.el)["rgb"],
            "w": _line_of(best.el)["w"] or 9525,
            "prst": _prst_geom(best.el)[0] or "rect", "sh": best}


def _added_borders(loc, title_bd):
    """指令未要求、模型自行添加的"边框类"形状（无文字、无填充或透明、有可见轮廓）。"""
    skip = id(title_bd["sh"]) if title_bd and title_bd["kind"] == "shape" else None
    out = []
    for sh in loc["added"]:
        if sh.text.strip() or id(sh) == skip:
            continue
        if not _line_of(sh.el)["visible"]:
            continue
        out.append(sh)
    return out

# ---------------------------------------------------------------------------
# 维度一：可用与可修改性（不通过则维度二整体为 0）
# ---------------------------------------------------------------------------


def _check_dim1(file_path):
    try:
        with zipfile.ZipFile(file_path) as zf:
            names = zf.namelist()
    except Exception as exc:
        return False, "文件无法作为 .pptx 包打开：%s" % exc
    if "ppt/presentation.xml" not in names:
        return False, "缺少 ppt/presentation.xml，不是可正常打开的演示文稿"
    try:
        Presentation(file_path)
    except Exception as exc:
        return False, "演示文稿内容损坏，无法正常打开：%s" % exc
    return True, ""


# ---------------------------------------------------------------------------
# 对外入口
# ---------------------------------------------------------------------------


class _TargetMissing(Exception):
    """目录不存在或目录中没有 .pptx：连待评测对象都没有，不进入维度判定。"""


def evaluate(dir_path):
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
        if not dir_path or not os.path.isdir(dir_path):
            raise _TargetMissing("目录不存在：%s" % dir_path)
        cands = sorted(f for f in os.listdir(dir_path)
                       if f.lower().endswith(".pptx") and not f.startswith("~"))
        if not cands:
            raise _TargetMissing("目录中未找到 .pptx 文件：%s" % dir_path)
        target = TARGET_NAME if TARGET_NAME in cands else cands[0]
        result["file_name"] = target
        _evaluate_impl(os.path.join(dir_path, target), dir_path, result)
    except _TargetMissing as exc:
        # 找不到待评测文件：维度1 未参与判定，dim1_pass / dim1_reason 一律留空，满分也置 0
        result["file_name"] = ""
        result["status"] = "error"
        result["error"] = str(exc)
        result["dim1_pass"] = ""
        result["dim1_reason"] = ""
        result["dim2_items"] = []
        result["total_score"] = 0
        result["max_score"] = 0
    except Exception as exc:
        # 文件存在但打不开 / 解析失败：维度1 判不通过，dim1_reason 给出错误摘要
        result["status"] = "error"
        result["error"] = "%s: %s" % (type(exc).__name__, exc)
        result["dim1_pass"] = False
        result["dim1_reason"] = "文件无法解析或无法打开：%s: %s" % (type(exc).__name__, exc)
        result["dim2_items"] = []
        result["total_score"] = 0
    return result

def _evaluate_impl(file_path, dir_path, result):
    ok1, reason = _check_dim1(file_path)
    result["dim1_pass"] = ok1
    result["dim1_reason"] = reason if not ok1 else \
        "交付物为 1 个可正常打开的 .pptx 演示文稿"
    if not ok1:
        result["status"] = "error"
        result["error"] = reason
        result["dim2_items"] = []
        result["total_score"] = 0
        return

    prs = Presentation(file_path)
    locs, groups, _infos, _src_map = _group_photos(prs)

    results = []
    tol = 0.02 * EMU_CM
    order_target = {(1, "L"): (1, 2, 3), (1, "R"): (4, 5, 6),
                    (2, "L"): (7, 8, 9), (2, "R"): (10, 11, 12)}
    n_pic = sum(len(loc["pics"]) for loc in locs)
    zidx = {}
    for loc in locs:
        for k, sh in enumerate(loc["all"]):
            zidx[id(sh)] = k
    bds = [_title_border(loc) for loc in locs]

    # 01 12 张照片按 1—12 顺序分入四个卡片（+5）
    ok = n_pic == 12 and len(locs) >= 2
    if ok:
        for key, want in order_target.items():
            got = [src for _sh, src, _in in groups[key]["pics"]]
            if len(got) != 3 or sorted(s for s in got if s) != list(want):
                ok = False
                break
    results.append(("+5：" if ok else "+0：") +
                   "12张照片一张不缺、无重复地全部放入两页的四个浅色内容卡片中，并按 1.png—12.png 的先后顺序对位："
                   "第1页左侧浅蓝卡片放第1、2、3张，第1页右侧浅黄卡片放第4、5、6张，"
                   "第2页左侧浅绿卡片放第7、8、9张，第2页右侧浅紫卡片放第10、11、12张")

    # 02 每卡三张照片的卡内落位顺序与编号一致（+3）
    ok = len(locs) >= 2
    for key in order_target:
        seq = groups[key]["seq"]
        if len(seq) != 3:
            ok = False
            continue
        idxs = [src for _sh, src, _in in seq]
        if any(s is None for s in idxs) or not (idxs[0] < idxs[1] < idxs[2]):
            ok = False
    results.append(("+3：" if ok else "+0：") +
                   "每个卡片内三张照片的落位顺序与编号一致——第1页为\"上方大图=组内第1张、左下小图=第2张、"
                   "右下小图=第3张\"，第2页为\"左上小图=第1张、左下小图=第2张、右侧大图=第3张\"")

    # 03 第 1 页"上一大下两小"（+5）
    ok = len(locs) >= 1
    for side in ("L", "R"):
        grp = groups[(1, side)]
        if len(grp["seq"]) != 3 or not grp["shape_ok"]:
            ok = False
            continue
        big, sl, sr = [x[0] for x in grp["seq"]]
        card = grp["box"]
        usable = (card[2] - card[0]) - 2 * 0.2 * EMU_CM
        if not (12.8 <= _cm(big.w) <= 13.8 and 2.6 <= _cm(big.h) <= 3.4
                and big.w >= 0.85 * usable
                and all(6.2 <= _cm(x.w) <= 7.0 for x in (sl, sr))
                and all(2.4 <= _cm(x.h) <= 3.2 for x in (sl, sr))
                and abs(sl.h - sr.h) <= 0.2 * EMU_CM
                and abs(sl.w - sr.w) <= 0.35 * EMU_CM
                and big.w >= 1.5 * max(sl.w, sr.w)
                and abs((sr.r - sl.l) - big.w) <= 0.6 * EMU_CM):
            ok = False
    results.append(("+5：" if ok else "+0：") +
                   "第1页两个卡片内各排成\"上一大下两小\"：上方1张大图单独占满卡片可用宽度"
                   "（约13.3cm宽、2.9-3.1cm高），下方2张小图等高并排（各约6.5-6.7cm宽、2.7-2.9cm高）")

    # 04 第 2 页"左二小右一大"（+5）
    ok = len(locs) >= 2
    for side in ("L", "R"):
        grp = groups[(2, side)]
        if len(grp["seq"]) != 3 or not grp["shape_ok"]:
            ok = False
            continue
        top, bottom, big = [x[0] for x in grp["seq"]]
        if not (all(4.8 <= _cm(x.w) <= 5.6 for x in (top, bottom))
                and all(2.6 <= _cm(x.h) <= 3.4 for x in (top, bottom))
                and abs(top.w - bottom.w) <= 0.2 * EMU_CM
                and abs(top.l - bottom.l) <= 0.12 * EMU_CM
                and 7.4 <= _cm(big.w) <= 8.6 and 5.8 <= _cm(big.h) <= 6.7
                and abs(big.t - top.t) <= 0.2 * EMU_CM
                and big.b >= bottom.b - 0.25 * EMU_CM
                and big.w >= 1.3 * max(top.w, bottom.w)):
            ok = False
    results.append(("+5：" if ok else "+0：") +
                   "第2页两个卡片内各排成\"左二小右一大\"：左列2张小图上下堆叠（各约5.1-5.3cm宽、2.9-3.1cm高），"
                   "右侧1张大图（约7.9-8.0cm宽、6.2-6.3cm高）纵向跨过左列两张的总高度，左右两组顶边对齐")

    # 05 照片不与角标/卡片小标题/卡片正文文字压叠（+3）
    ok = n_pic == 12 and len(locs) >= 2
    for i, loc in enumerate(locs, 1):
        for side in ("L", "R"):
            rects = []
            for role in ("badge", "head", "body"):
                sh = loc.get(role + side)
                if sh is None:
                    ok = False
                    continue
                rect = _text_rect(sh)
                if rect:
                    rects.append(rect)
            for pic, _src, _in in groups[(i, side)]["pics"]:
                if any(_hit_overlap(pic.box, r) for r in rects):
                    ok = False
            head, body = loc.get("head" + side), loc.get("body" + side)
            if head is not None and body is not None:
                hr, br = _text_rect(head), _text_rect(body)
                if hr and br and _hit_overlap(hr, br, 0.02):
                    ok = False
            if body is not None:
                br = _text_rect(body)
                if br and br[3] > body.b + 0.1 * EMU_CM:
                    ok = False
    results.append(("+3：" if ok else "+0：") +
                   "12张照片整体落在卡片内部原本空着的浅色区域内，不与卡片内已有的"
                   "\"MODULE A/MODULE B/AERO/COCKPIT\"角标、卡片小标题、卡片正文段落文字发生压叠；"
                   "若为腾出位置而调整了小标题与正文的纵向位置，调整后各段文字仍完整可读、互不重叠")

    # 06 照片全部落在所属卡片浅色区块边界内（+3）
    ok = n_pic == 12 and len(locs) >= 2
    for i, loc in enumerate(locs, 1):
        bounds = [sh.t for sh in (loc.get("lineL"), loc.get("lineR")) if sh]
        bounds += [loc["kpiN%d" % k].t for k in range(4)
                   if loc.get("kpiN%d" % k) is not None]
        ymax = min(bounds) if bounds else CARD_TOP[i] + CARD_CY[i]
        for side in ("L", "R"):
            grp = groups[(i, side)]
            card = grp["box"]
            for pic, _src, _in in grp["pics"]:
                if not (pic.l >= card[0] - tol and pic.t >= card[1] - tol
                        and pic.r <= card[2] + tol and pic.b <= card[3] + tol):
                    ok = False
                if pic.b > ymax + tol:
                    ok = False
    results.append(("+3：" if ok else "+0：") +
                   "12张照片全部落在所属卡片的浅色区块边界内（第1页卡片区约 x1.78-16.56cm 与 "
                   "x17.32-32.10cm、y4.9-16.9cm；第2页卡片区约 y5.1-17.2cm），未越出卡片边缘，"
                   "未压住卡片底部的1.7磅彩色横线，也未压到卡片下方的KPI数字带")

    # 07 两页主标题加浅蓝色矩形边框（+5）
    ok = (len(bds) >= 2 and all(bd is not None for bd in bds[:2])
          and all(_is_light_blue(bd["rgb"]) for bd in bds[:2])
          and all(bd["w"] and 6350 <= bd["w"] <= 19050 for bd in bds[:2])
          and all(bd["prst"] in ("rect", "roundRect", "snipRoundRect")
                  for bd in bds[:2])
          and bds[0]["kind"] == bds[1]["kind"]
          and (bds[0]["rgb"] or "") == (bds[1]["rgb"] or ""))
    results.append(("+5：" if ok else "+0：") +
                   "两页的页面主标题\"高效动力链方案\"\"整车性能协同优化\"各自加上了浅蓝色矩形边框"
                   "（如 87CEFA、ADD8E6、9EDAF5 一类浅蓝色，线宽0.5-1.5磅），四周闭合、两页做法一致")

    # 08 标题边框贴合标题文字尺寸（+3）
    ok = len(bds) >= 2 and all(bd is not None for bd in bds[:2])
    if ok:
        for bd in bds[:2]:
            box = bd["box"]
            if not (5.4 <= _cm(box[2] - box[0]) <= 7.6
                    and 0.95 <= _cm(box[3] - box[1]) <= 1.15):
                ok = False
    results.append(("+3：" if ok else "+0：") +
                   "标题边框贴合标题文字尺寸（框宽约5.5-7.5cm、框高约1.0-1.1cm），"
                   "未直接沿用横跨半页的19.05cm空文本框宽度形成大片空框，"
                   "也未把边框套在照片区、卡片或整页上代替标题")

    # 09 照片与卡片边缘、照片之间留白均匀且对齐（+3）
    ok = len(locs) >= 2
    for i, loc in enumerate(locs, 1):
        for side in ("L", "R"):
            grp = groups[(i, side)]
            if len(grp["seq"]) != 3:
                ok = False
                continue
            first, second, third = [x[0] for x in grp["seq"]]
            card = grp["box"]
            for pic in (first, second, third):
                if (pic.l - card[0] < 0.2 * EMU_CM - tol
                        or card[2] - pic.r < 0.2 * EMU_CM - tol):
                    ok = False
            if i == 1:
                gaps = [third.l - second.r]
                if (abs(second.t - third.t) > 0.15 * EMU_CM
                        or abs(second.b - third.b) > 0.15 * EMU_CM):
                    ok = False
            else:
                gaps = [second.t - first.b, third.l - max(first.r, second.r)]
                if abs(first.l - second.l) > 0.12 * EMU_CM:
                    ok = False
            if any(not (0.2 * EMU_CM - tol <= g <= 0.6 * EMU_CM + tol)
                   for g in gaps):
                ok = False
            for pic in (first, second, third):
                for sh in loc["all"]:
                    if sh.kind != "sp" or zidx[id(sh)] <= zidx[id(pic)]:
                        continue
                    if _fill_rgb(sh.el) is None:
                        continue
                    if _overlap_area(pic.box, sh.box) > 0.2 * pic.w * pic.h:
                        ok = False
    results.append(("+3：" if ok else "+0：") +
                   "照片与卡片边缘、照片与照片之间留有均匀留白（距卡片左右边缘内缩不少于0.2cm、"
                   "相邻照片间距0.2-0.6cm），同排照片顶边与底边对齐，无相互压叠或参差，"
                   "也不出现照片被上层形状遮挡只露出一部分")

    # 10 照片被物理裁切后重新编码写入（-3）
    cropped = [info for grp in groups.values() for (_sh, src, info) in grp["pics"]
               if src and info["px"] and tuple(info["px"]) != PHOTO_PX]
    bad = bool(cropped)
    results.append(("-3：" if bad else "-0：") +
                   "照片画面被物理裁切后重新编码写入（如宽幅位置的照片只保留 1448×330 像素、"
                   "上下各被切掉约35%画面，并排小图被切成 1448×617，纵向小图被切成 1448×842，"
                   "大图被切成 1383×1086；第1张体积由 1,177,492 字节降为 463,800 字节，"
                   "像素不再是完整的 1448×1086），被切掉的画面内容在演示文稿内无法恢复")

    # 11 新增浅色边框颜色不是浅蓝色/四处不统一（-3）
    border_clrs = []
    for loc, bd in zip(locs, bds):
        for sh in _added_borders(loc, bd):
            rgb = _line_of(sh.el)["rgb"]
            if rgb:
                border_clrs.append(rgb.upper())
    bad = bool(border_clrs) and (len(set(border_clrs)) > 1
                                 or any(not _is_light_blue(c) for c in border_clrs))
    results.append(("-3：" if bad else "-0：") +
                   "新增的浅色边框中有2处使用粉色 F6B6D2（第1页右侧浅黄卡片、第2页右侧浅紫卡片各1处），"
                   "不是要求的浅蓝色，四处边框颜色不统一")

    # 12 新增边框线条压穿文字（-3）
    bad = False
    for loc, bd in zip(locs, bds):
        borders = _added_borders(loc, bd)
        if not borders:
            continue
        rects = [r for r in (_text_rect(sh) for sh in loc["all"]
                             if sh.kind == "sp") if r]
        for sh in borders:
            for pt in _stroke_points(sh):
                if any(_pt_in_rect(pt, r) for r in rects):
                    bad = True
                    break
            if bad:
                break
    results.append(("-3：" if bad else "-0：") +
                   "新增的浅色边框上边线横穿\"实物展示\"一行文字，线条与文字笔画重叠压字，圆角处还切在文字左半边")

    # 13 新增或遗留指令未要求的文字（-1）
    ph_re = re.compile(r"待插入照片|图片占位|此处放图|见附件|待补充|占位|TODO", re.I)
    bad = False
    for loc in locs:
        if any(sh.text.strip() for sh in loc["added"]):
            bad = True
        if any(ph_re.search(sh.text or "") for sh in loc["all"]):
            bad = True
    results.append(("-1：" if bad else "-0：") +
                   "新增或遗留了指令未要求的文字——如\"实物展示\"一类说明文字共4处"
                   "（Microsoft YaHei 11磅加粗、字色 9EDAF5／F6B6D2，与两页统一使用的 "
                   "Noto Sans CJK SC 字体体系不一致），或\"待插入照片\"\"图片占位\"\"此处放图\""
                   "\"见附件\"等未清理的提示文字")

    # 14 交付形态不正确（-1）
    extra = [f for f in os.listdir(dir_path)
             if os.path.splitext(f)[1].lower() in (".zip", ".rar", ".7z", ".pdf")]
    bad = result["file_name"] != TARGET_NAME or bool(extra)
    results.append(("-1：" if bad else "-0：") +
                   "交付形态不正确——文件名被改写（如追加\"_照片排版版\"\"_图片版\"等后缀）"
                   "而未保持\"核心技术.pptx\"，或交付物为 .zip／.rar 压缩包、PDF、导出图片，"
                   "需另行转换才能得到可编辑的演示文稿")

    # 15 顶部深色标题条或装饰细条被改动（-3）
    bad = False
    for i, loc in enumerate(locs, 1):
        bar, strip = loc.get("bar"), loc.get("strip")
        if bar is None or strip is None:
            bad = True
            continue
        if _fill_rgb(bar.el) != BAR_CLR:
            bad = True
        if not (_near(bar.l, BAR_OFF[0], 0.1) and _near(bar.t, BAR_OFF[1], 0.1)
                and _near(bar.w, BAR_EXT[0], 0.1) and _near(bar.h, BAR_EXT[1], 0.1)):
            bad = True
        if _fill_rgb(strip.el) != EXPECT[i]["strip"]:
            bad = True
        if not (_near(strip.t, STRIP_OFF[1], 0.05)
                and _near(strip.w, STRIP_EXT[0], 0.1)
                and _near(strip.h, STRIP_EXT[1], 0.05)):
            bad = True
    results.append(("-3：" if bad else "-0：") +
                   "两页顶部深色标题条（色值 0B1220、33.87cm×2.18cm、位于页面最上端）或其下方"
                   "0.14cm装饰细条（第1页青色 5EEAD4、第2页橙色 F59E0B）的颜色、尺寸、位置被改动或删除")

    # 16 页面主标题文字或格式被改动（-1）
    bad = False
    for i, loc in enumerate(locs, 1):
        sh = loc.get("title")
        if sh is None or _norm_text(sh.text) != EXPECT[i]["title"]:
            bad = True
            continue
        sz, bold, clr, face = _first_font(sh.el)
        if not (sz == 24.0 and bold and clr == "FFFFFF" and face == TITLE_FONT):
            bad = True
        if not (_near(sh.l, TITLE_OFF[0]) and _near(sh.t, TITLE_OFF[1])):
            bad = True
        if _body_pr(sh.el)["anchor"] != "ctr":
            bad = True
    results.append(("-1：" if bad else "-0：") +
                   "第1页主标题\"高效动力链方案\"、第2页主标题\"整车性能协同优化\"的文字被改写，"
                   "或不满足 Noto Sans CJK SC 24磅 加粗 白色 FFFFFF、左1.22cm 上0.46cm 竖直居中")

    # 17 右上角提示语被删改（-1）
    bad = False
    for i, loc in enumerate(locs, 1):
        sh = loc.get("tagline")
        if sh is None or _norm_text(sh.text) != EXPECT[i]["tagline"]:
            bad = True
            continue
        sz, _b, clr, face = _first_font(sh.el)
        if not (sz == 10.5 and clr == "CBD5E1" and face == TITLE_FONT):
            bad = True
        if _algn_of(sh.el) != "r" or not _near(sh.l, TAGLINE_LEFT):
            bad = True
    results.append(("-1：" if bad else "-0：") +
                   "两页右上角的\"结构减重 · 传动效率 · 耐久验证\"\"气动布局 · 人机空间 · 安全体验\"被删改，"
                   "或不满足 Noto Sans CJK SC 10.5磅、字色 CBD5E1、右对齐、右侧起点 x22.86cm")

    # 18 导语段被删改（-1）
    bad = False
    for i, loc in enumerate(locs, 1):
        sh = loc.get("lead")
        text, want_sz, want_clr = EXPECT[i]["lead"]
        if sh is None or _norm_text(sh.text) != text:
            bad = True
            continue
        sz, _b, clr, _f = _first_font(sh.el)
        if sz != want_sz or clr != want_clr:
            bad = True
    results.append(("-1：" if bad else "-0：") +
                   "两页导语段\"以更轻的传动单元和更稳定的材料连接方式，提升小型赛车平台的动力响应与可靠性。\""
                   "\"围绕尾部气流、驾驶姿态与防护结构进行协同设计，让车辆在高速稳定性和操作舒适度之间取得更好的平衡。\""
                   "的文字被删改，或不满足第1页16磅 字色 DDE9F6、第2页15.5磅 字色 E2E8F0")

    # 19 四个内容卡片底色/宽度/描边/并排布局被改动（-3）
    bad = False
    for i, loc in enumerate(locs, 1):
        for k, side in ((0, "L"), (1, "R")):
            sh = loc.get("card" + side)
            if sh is None:
                bad = True
                continue
            if _fill_rgb(sh.el) != EXPECT[i]["cards"][k]:
                bad = True
            if not (_near(sh.w, CARD_CX) and _near(sh.l, CARD_LEFT[k])):
                bad = True
            rgb = _line_of(sh.el)["rgb"]
            if rgb is not None and rgb != "FFFFFF":
                bad = True
    results.append(("-3：" if bad else "-0：") +
                   "四个内容卡片的底色被改动（第1页浅蓝 E0F2FE、浅黄 FEF3C7；第2页浅绿 DCFCE7、"
                   "浅紫 F3E8FF），或卡片宽度14.78cm、白色描边、左侧卡片左1.78cm与右侧卡片左17.32cm"
                   "的并排布局被改动")

    # 20 卡片角标被改动（-1）
    bad = False
    for i, loc in enumerate(locs, 1):
        for k, side in ((0, "L"), (1, "R")):
            sh = loc.get("badge" + side)
            text, want_clr = EXPECT[i]["badges"][k]
            if sh is None or _norm_text(sh.text) != text:
                bad = True
                continue
            sz, bold, clr, _f = _first_font(sh.el)
            if sz != BADGE_SZ or not bold or clr != want_clr:
                bad = True
    results.append(("-1：" if bad else "-0：") +
                   "卡片角标\"MODULE A\"（38BDF8）、\"MODULE B\"（F59E0B）、\"AERO\"（22C55E）、"
                   "\"COCKPIT\"（A78BFA）的文字、9.5磅加粗格式或字色被改动")

    # 21 四个卡片小标题被改动（-1）
    bad = False
    for i, loc in enumerate(locs, 1):
        for k, side in ((0, "L"), (1, "R")):
            sh = loc.get("head" + side)
            if sh is None or _norm_text(sh.text) != EXPECT[i]["heads"][k]:
                bad = True
                continue
            sz, bold, clr, _f = _first_font(sh.el)
            if sz != HEAD_SZ or not bold or clr != HEAD_CLR:
                bad = True
    results.append(("-1：" if bad else "-0：") +
                   "四个卡片小标题\"智能变速总成\"\"高强复合半轴\"\"尾部导流组件\"\"座舱轻量布局\""
                   "的文字被改写，或不满足 18磅 加粗 字色 172033")

    # 22 四段卡片正文被删改或格式不一致（-3）
    bad = False
    for i, loc in enumerate(locs, 1):
        sizes = []
        for k, side in ((0, "L"), (1, "R")):
            sh = loc.get("body" + side)
            if sh is None or _norm_text(sh.text) != EXPECT[i]["bodies"][k]:
                bad = True
                continue
            sz, _b, clr, _f = _first_font(sh.el)
            sizes.append(sz)
            if sz != BODY_SZ or clr != BODY_CLR:
                bad = True
            if _body_pr(sh.el)["vert"] not in (None, "horz"):
                bad = True
            spc = _line_spacing_pct(sh.el)
            if spc is not None and spc < 90.0:
                bad = True
        if len(set(sizes)) > 1:
            bad = True
    results.append(("-3：" if bad else "-0：") +
                   "四段卡片正文（\"将传统滑动导向改为低阻尼滚柱导轨…\"\"采用碳纤维多向铺层与分段固化工艺…\""
                   "\"以双曲率翼面、短端板和尾部整流片…\"\"将踏板组、转向柱和防滚架交汇点重新布置…\"）"
                   "的文字被删除、精简、改写或调序，或不满足 12.2磅 字色 475569"
                   "（含为给照片腾位而调小字号、压缩行距、改为竖排，致同页四段正文格式不再一致）")

    # 23 卡片底部 1.7 磅彩色横线被改动（-1）
    bad = False
    for i, loc in enumerate(locs, 1):
        for k, side in ((0, "L"), (1, "R")):
            sh = loc.get("line" + side)
            if sh is None:
                bad = True
                continue
            ln = _line_of(sh.el)
            if not (ln["visible"] and ln["rgb"] == EXPECT[i]["line_clr"][k]):
                bad = True
            if ln["w"] is None or abs(ln["w"] - LINE_W_EMU) > 1270:
                bad = True
            if not (_near(sh.w, LINE_CX, 0.1) and _near(sh.t, EXPECT[i]["line_y"])
                    and sh.h <= 0.02 * EMU_CM):
                bad = True
    results.append(("-1：" if bad else "-0：") +
                   "四个卡片底部的1.7磅彩色横线（依次为 38BDF8、F59E0B、22C55E、A78BFA，长13.36cm）"
                   "被删除、改色、加粗或上下移位")

    # 24 8 组 KPI 数字与说明被删改（-3）
    bad = False
    for i, loc in enumerate(locs, 1):
        for k in range(4):
            num, want_clr, cap = EXPECT[i]["kpi"][k]
            shn, shc = loc.get("kpiN%d" % k), loc.get("kpiC%d" % k)
            if shn is None or _norm_text(shn.text) != num:
                bad = True
            else:
                sz, bold, clr, _f = _first_font(shn.el)
                if sz != KPI_NUM_SZ or not bold or clr != want_clr:
                    bad = True
            if shc is None or _norm_text(shc.text) != cap:
                bad = True
            else:
                sz, _b, clr, _f = _first_font(shc.el)
                if sz != KPI_CAP_SZ or clr != KPI_CAP_CLR \
                        or _algn_of(shc.el) != "ctr":
                    bad = True
    results.append(("-3：" if bad else "-0：") +
                   "8组KPI数字与说明（18%／传动损耗下降、52%／总成质量减轻、32%／接合强度提升、"
                   "5.5h／耐久循环增加、24%／尾涡强度降低、0.34／目标阻力系数、120mm／腿部空间增加、"
                   "15%／框架利用率提升）的文字被删改，或数字不满足19磅加粗（配色 5EEAD4／38BDF8／"
                   "F59E0B／A78BFA）、说明文字不满足8.5磅 字色 94A3B8 居中")

    # 25 页面尺寸或幻灯片数量被改动（-1）
    slides = list(prs.slides)
    bad = not (_near(prs.slide_width or 0, SLIDE_W, 0.1)
               and _near(prs.slide_height or 0, SLIDE_H, 0.1)) or len(slides) != 2
    results.append(("-1：" if bad else "-0：") +
                   "页面尺寸被改动（不再是 33.87cm×19.05cm 的16:9屏幕演示），"
                   "或幻灯片数量不再是2页（新增照片页、封面页、目录页、说明页，或删页）")

    # 26 备注页被删除或被写入排版说明（-1）
    bad = len(slides) < 2
    for slide in slides[:2]:
        if not slide.has_notes_slide:
            bad = True
            continue
        if (slide.notes_slide.notes_text_frame.text or "").strip():
            bad = True
    results.append(("-1：" if bad else "-0：") +
                   "两张幻灯片各自附带的备注页被删除，或备注中被写入排版说明文字")

    # 27 照片被加旋转/描边/阴影/发光/圆角裁剪/动画/超链接（-1）
    bad = False
    for loc in locs:
        for pic in loc["pics"]:
            ln = _line_of(pic.el)
            if _rot_of(pic.el) != 0 or _has_effect(pic.el):
                bad = True
            if ln["visible"] and (ln["w"] or 0) > 0:
                bad = True
            if (_prst_geom(pic.el)[0] or "rect") != "rect":
                bad = True
            if pic.el.find(".//" + _a("hlinkClick")) is not None:
                bad = True
    with zipfile.ZipFile(file_path) as zf:
        names = zf.namelist()
        for name in names:
            if not re.match(r"ppt/slides/slide\d+\.xml$", name):
                continue
            xml = zf.read(name).decode("utf-8", "ignore")
            if re.search(r"<p:(anim|animEffect|animMotion|animRot|animScale|"
                         r"animClr)\b", xml):
                bad = True
        core = zf.read("docProps/core.xml").decode("utf-8", "ignore") \
            if "docProps/core.xml" in names else ""
        custom = zf.read("docProps/custom.xml").decode("utf-8", "ignore") \
            if "docProps/custom.xml" in names else ""
    results.append(("-1：" if bad else "-0：") +
                   "给照片加了指令未要求的旋转角度、厚描边、阴影、外发光、艺术效果、圆角裁剪"
                   "或动画、超链接、触发动作")

    # 28 文档属性标题被改写或新增无关自定义属性（-1）
    m = re.search(r"<dc:title>(.*?)</dc:title>", core, re.S)
    bad = (m.group(1).strip() if m else "") != DOC_TITLE
    props = set(re.findall(r'name="([^"]+)"', custom))
    if props - set(CUSTOM_PROPS):
        bad = True
    results.append(("-1：" if bad else "-0：") +
                   "文档属性中的标题\"高效移动平台技术方案\"被改写，或新增了与本次排版无关的自定义属性")

    # ------------------------------------------------------------------
    # 把 results 反解为结构化 dim2_items
    # ------------------------------------------------------------------
    if len(results) != len(_MAX_DELTAS_ORDER):
        raise RuntimeError("dim2 项数不一致：results=%d vs order=%d"
                           % (len(results), len(_MAX_DELTAS_ORDER)))
    dim2_items = []
    total = 0
    for raw, max_delta in zip(results, _MAX_DELTAS_ORDER):
        head, _sep, rule = raw.partition("：")
        value = int(head)
        if max_delta > 0:
            hit = value > 0
        else:
            hit = value != 0
        delta = max_delta if hit else 0
        total += delta
        dim2_items.append({"rule": rule, "max_delta": max_delta,
                           "delta": delta, "hit": hit})
    result["dim2_items"] = dim2_items
    result["total_score"] = total


if __name__ == "__main__":
    _dir = sys.argv[1] if len(sys.argv) >= 2 \
        else os.path.dirname(os.path.abspath(__file__))
    try:
        print(json.dumps(evaluate(_dir), ensure_ascii=False, indent=2))
    except Exception:
        traceback.print_exc()
