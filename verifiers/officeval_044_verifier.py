#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# pyright: reportUnknownMemberType=false, reportUnknownVariableType=false
# pyright: reportUnknownArgumentType=false, reportMissingImports=false
# pyright: reportOptionalMemberAccess=false, reportGeneralTypeIssues=false
"""
044 号任务评分脚本：演示文稿"新增目录页 + 第9页嵌入可单击播放的竖版视频"。

被评估对象：一个 .pptx 演示文稿。
评分口径完全对齐同目录 细则.txt：
    维度一 可用与可修改性：能否作为一个正常的 PowerPoint 演示文稿被打开；
    维度二 完成度：16 条加分项（合计 66 分）+ 16 条扣分项（合计 -32 分）。

说明：按接口统一约定，本模块不修改全局 sys.stdout，也不 print 主结果。
主结果统一由 evaluate(dir_path) -> dict 返回；入参为"脚本所在目录的路径"，
由脚本自己在该目录内定位并打开被评估的 .pptx。本文件仅在 __main__ 下作为
本地调试入口，才把 dict 序列化为 JSON 打印。

维度二每一项在 results 中都表现为一条字符串：
    "+N：规则文字"  表示该加分项拿到分；"+0：规则文字" 表示没拿到；
    "-N：规则文字"  表示该扣分项被触发；"-0：规则文字" 表示没被触发。
字符串顺序必须与 _MAX_DELTAS_ORDER 严格一一对应。
"""

import os
import re
import sys
import json
import struct
import hashlib
import zipfile
import traceback

from lxml import etree

SCRIPT_ID = "044"

# 满分只统计加分项（扣分项为负向惩罚，不计入满分）。
MAX_SCORE_TOTAL = 66

_MAX_DELTAS_ORDER = [
    5,    # 01 第8页原有五处文字全部删除
    5,    # 02 新增"目录"二字作为页面标题，位于三条篇章标题上方
    3,    # 03 "目录"二字为宋体、96号（字号与中文字体同时达标）
    3,    # 04 "目录"二字水平居中
    5,    # 05 目录正文仅列出三条篇章大标题且文字正确
    3,    # 06 三条篇章标题均为仿宋、44号（字号与中文字体同时达标）
    3,    # 07 三条篇章标题位于"目录"下方、按序纵向排列、间距均匀无重叠未溢出
    5,    # 08 三条篇章标题均设幻灯片内部跳转超链接
    5,    # 09 三条跳转目标分别正确指向第10、16、21页
    5,    # 10 第9页插入 mp4 并随文稿嵌入包内且声明内容类型
    5,    # 11 嵌入视频与所提供素材字节完全一致（未重编码、未拉伸补边）
    5,    # 12 视频对象宽高落在指定范围
    5,    # 13 视频左上角与插画区域四边重合，且未越界压盖右侧卡片
    5,    # 14 视频单击触发并连续自动播放至结束
    3,    # 15 视频停止时不显示（showWhenStopped="0"）
    1,    # 16 视频封面首帧为该视频自身画面
    -3,   # 17 第8页右侧深色木桌静物摄影配图未保留
    -3,   # 18 第8页左侧深蓝色竖向色块未保留
    -1,   # 19 第8页右下角"01"页码文字或其右侧浅灰短横线未保留
    -1,   # 20 第8页新增了未要求的装饰图形或说明性文字
    -1,   # 21 第8页目录文字块压到右侧摄影配图之上
    -3,   # 22 第9页帆船码头插画或其上白色圆角描边框被删除、缩放或移位
    -1,   # 23 第9页视频对象层级过低被遮挡
    -3,   # 24 第9页顶部四项导航文字或金色分隔线被改动
    -3,   # 25 第9页标题、副标题或褐色粗短线被改动、移位或删除
    -3,   # 26 第9页米色圆角卡片或卡内文字被改动、遮挡或删除
    -1,   # 27 第9页右下角"09"页码及短横线被删除
    -3,   # 28 除第8、9页以外的页面（含三页手写书法标题图片）被改动
    -3,   # 29 演示文稿总页数不再是25页
    -1,   # 30 页面尺寸不再是33.87×19.05厘米
    -1,   # 31 演讲者备注内容被清空或改写
    -1,   # 32 三条篇章标题被加下划线或改为超链接主题色
]

assert sum(d for d in _MAX_DELTAS_ORDER if d > 0) == MAX_SCORE_TOTAL

NS = {
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "p14": "http://schemas.microsoft.com/office/powerpoint/2010/main",
    "ct": "http://schemas.openxmlformats.org/package/2006/content-types",
    "pr": "http://schemas.openxmlformats.org/package/2006/relationships",
}
R_ID = "{%s}id" % NS["r"]
R_LINK = "{%s}link" % NS["r"]
R_EMBED = "{%s}embed" % NS["r"]

EMU_PER_CM = 360000.0
TOL_01CM = 36000        # 0.1 厘米
TOL_05CM = 180000       # 0.5 厘米
TOL_1CM = 360000        # 1.0 厘米
TOL_GEOM = 36000        # 判定"位置/尺寸未被改动"的容差：0.1 厘米

# ---- 版面基线：页面尺寸与页数 ----
SLD_SZ = (12192000, 6858000)          # 33.87 × 19.05 厘米
SLIDE_COUNT = 25

# ---- 第8页需要保留的元素（细则以尺寸/位置/颜色/字号描述，不依赖任何外部文件）----
S8_PHOTO_OFF = (6601968, 347472)      # 距左 18.34 厘米、距上 0.97 厘米
S8_PHOTO_EXT = (5212080, 6080760)     # 宽 14.48 厘米、高 16.89 厘米
S8_NAVY_OFF = (0, 0)
S8_NAVY_EXT = (6446520, 6858000)      # 宽 17.91 厘米、高 19.05 厘米
S8_NAVY_FILL = "173B57"
S8_PAGENO_TEXT = "01"
S8_PAGENO_OFF = (11292840, 6446520)   # 距左 31.37 厘米、距上 17.91 厘米
S8_PAGENO_EXT = (384048, 228600)
S8_PAGENO_SZ = "950"
S8_PAGENO_CLR = "697987"
S8_DASH_OFF = (11722608, 6556248)
S8_DASH_EXT = (228600, 0)             # 长 0.64 厘米
S8_DASH_CLR = "697987"
S8_DASH_ALPHA = "70000"               # 70% 透明

# 第8页原有的五处文字（指令要求全部删除）
S8_OLD_TEXTS = (
    "第6课第一子目",
    "新航路开辟的动因",
    "历史课堂说课设计",
    "以“为什么出发”为问题线索，重构近代早期欧洲远航的动力图景。",
    "澜川市星商中学·林沐辰",
)
# 第8页原有非文字图形的外框尺寸集合（用于识别"新增装饰图形"）
S8_BASE_GRAPHIC_EXTS = (
    (6446520, 6858000),   # 深蓝色竖向色块
    (5212080, 6080760),   # 深色木桌静物摄影配图
    (2286000, 0),         # 标题下方金色短线
    (228600, 0),          # 页码右侧浅灰短横线
)

# ---- 第9页需要保留的元素 ----
S9_NAV_TEXTS = ("教学理念与立意", "教学前端分析", "教学过程设计", "教学评价")
S9_NAV_OFFS = ((1645920, 91440), (4069080, 91440),
               (6492240, 91440), (8915400, 91440))
S9_NAV_EXT = (2148840, 228600)
S9_NAV_SZ = "2000"
S9_NAV_CLR = "FFFFFF"
S9_GOLD_OFF = (0, 420624)
S9_GOLD_EXT = (12191695, 0)
S9_GOLD_CLR = "D6A642"
S9_TITLE_TEXT = "情境导入：一封来自港口的委托信"
S9_TITLE_OFF = (676656, 694944)
S9_TITLE_EXT = (7132320, 457200)
S9_TITLE_SZ = "2800"
S9_TITLE_CLR = "173B57"
S9_SUB_TEXT = "左侧图片、右侧文字，形成观察—设问—入课的图文结构"
S9_SUB_OFF = (713232, 1170432)
S9_SUB_EXT = (8869680, 256032)
S9_SUB_SZ = "1350"
S9_SUB_CLR = "697987"
S9_BROWN_OFF = (713232, 1499616)
S9_BROWN_EXT = (1828800, 0)           # 长 5.08 厘米
S9_BROWN_CLR = "B36B2C"
S9_ILLUS_OFF = (749808, 1755648)      # 左侧帆船码头插画，也是视频应覆盖的区域
S9_ILLUS_EXT = (5120640, 4343400)     # 宽 14.22 厘米、高 12.06 厘米
S9_FRAME_LN_CLR = "FFFFFF"
S9_FRAME_LN_ALPHA = "80000"
S9_CARD_OFF = (6217920, 1755648)
S9_CARD_EXT = (5166360, 4343400)      # 宽 14.35 厘米、高 12.06 厘米
S9_CARD_FILL = "FFF8EC"
S9_CARD_TITLE_TEXT = "致年轻的航海学徒："
S9_CARD_TITLE_OFF = (6565392, 2057400)
S9_CARD_TITLE_SZ = "1800"
S9_BODY_OFF = (6583680, 2606040)
S9_BODY_EXT = (4343400, 2468880)
S9_BODY_SZ = "1550"
S9_BODY_CLR = "27313A"
S9_BODY_PARA_COUNT = 3
S9_INTENT_TEXT = "导入意图：以港口场景激活学生经验，提出本课核心问题。"
S9_INTENT_OFF = (6583680, 5559552)
S9_INTENT_SZ = "1250"
S9_PAGENO_TEXT = "09"
S9_PAGENO_OFF = (11292840, 6446520)
S9_PAGENO_EXT = (384048, 228600)
S9_PAGENO_SZ = "950"
S9_DASH_OFF = (11722608, 6556248)
S9_DASH_EXT = (228600, 0)             # 长 0.64 厘米
# 第9页原有各元素的外框（用于识别"新增元素是否遮挡米色卡片"）
S9_BASE_RECTS = (
    ((0, 0), (12191695, 420624)),
    ((1645920, 91440), (2148840, 228600)),
    ((4069080, 91440), (2148840, 228600)),
    ((6492240, 91440), (2148840, 228600)),
    ((8915400, 91440), (2148840, 228600)),
    ((0, 420624), (12191695, 0)),
    ((676656, 694944), (7132320, 457200)),
    ((713232, 1170432), (8869680, 256032)),
    ((713232, 1499616), (1828800, 0)),
    ((749808, 1755648), (5120640, 4343400)),
    ((6217920, 1755648), (5166360, 4343400)),
    ((6565392, 2057400), (4297680, 301752)),
    ((6583680, 2606040), (4343400, 2468880)),
    ((6565392, 5413248), (1508760, 0)),
    ((6583680, 5559552), (4297680, 301752)),
    ((11292840, 6446520), (384048, 228600)),
    ((11722608, 6556248), (228600, 0)),
)

# ---- 目录页应有的内容 ----
CATALOG_TITLE = "目录"
CATALOG_TITLE_SZ_MIN = 9500            # 允许 95—97 号
CATALOG_TITLE_SZ_MAX = 9700
CATALOG_TITLE_FONT = "宋体"
ENTRY_TEXTS = ("第一篇探索！", "第二篇讨论", "第三篇作业")
ENTRY_SZ = "4400"
ENTRY_FONT = "仿宋"
ENTRY_TARGETS = (10, 16, 21)           # 三条条目应跳转到的页序号
HLINK_SLIDE_JUMP = "ppaction://hlinksldjump"
REL_TYPE_SLIDE = ("http://schemas.openxmlformats.org/officeDocument/"
                  "2006/relationships/slide")
REL_TYPE_VIDEO = ("http://schemas.openxmlformats.org/officeDocument/"
                  "2006/relationships/video")
REL_TYPE_MEDIA = "http://schemas.microsoft.com/office/2007/relationships/media"
# 目录页不应混入的栏目名（非篇章标题）
NON_CHAPTER_WORDS = ("教学理念与立意", "教学前端分析", "教学过程设计", "教学评价",
                     "教学重难点", "教学反思", "板书设计", "作业设计")

# ---- 视频对象的目标几何 ----
VIDEO_CX_MIN, VIDEO_CX_MAX = 5040000, 5220000   # 14.0—14.5 厘米
VIDEO_CY_MIN, VIDEO_CY_MAX = 4284000, 4392000   # 11.9—12.2 厘米
VIDEO_X_MIN, VIDEO_X_MAX = 720000, 792000       # 距左 2.0—2.2 厘米
VIDEO_Y_MIN, VIDEO_Y_MAX = 1728000, 1800000     # 距上 4.8—5.0 厘米
VIDEO_RIGHT_LIMIT = 6228000                     # 页面左起 17.3 厘米
MP4_W, MP4_H = 720, 960                         # 竖版素材分辨率
POSTER_RATIO = MP4_W / float(MP4_H)             # 0.75
POSTER_RATIO_TOL = 0.03
PLAY_CMDS = ("togglepause", "play", "playfrom(0.0)")

# ---- 输入基线：25 页结构签名（形状类型/位置/尺寸/几何/填充/文字/字号/字体/颜色）----
BASELINE_SLIDE_FP = {
    1: "ecb4d485653f650cd5ecb4dfd72406cd",
    2: "efe8f1700ec48ceab9367920d71c510f",
    3: "0672c6a458e65ae2970d8e7a720178c5",
    4: "cb7bc7e9d99e6b56fd56927e8d72a1fd",
    5: "f110ae1ae6fef5bfcfce4bbff0736873",
    6: "2ae1d68459273dfa7705b5d2b18dd137",
    7: "7217b0b6bfde8ac200852963952370c0",
    8: "f9c44420e384520f31775f5a62e948a4",
    9: "3087a45e8bee4c9ebe7efb264e97ffaa",
    10: "fc69dde131e599428549da5691ee0a5d",
    11: "8438b5709f8e6ca446b8df434ad90488",
    12: "1e6fe9dc684412bd4796dbcae9902a24",
    13: "9cba0fd490e5daf33d0e59b4399de563",
    14: "70af7870c757a239485d405916504213",
    15: "f64e84384a73bd80a69cd9e7341d326d",
    16: "aa6ac3762f552bed9593d08b40b5b676",
    17: "cd7eaa8e87c793f7855d3bde32b704cb",
    18: "0ab999c845ebf00de2f87905b50c001d",
    19: "aee3d36952ece2200fb120b96e19c8f2",
    20: "e39006311b3682a359bf906fe200dc19",
    21: "8fd37467eb9ee7b132e1857d8d2bc79e",
    22: "15b1f644cc481af374c956dd72b973da",
    23: "1c31454a73da0d0a5d67f98d832bd1a5",
    24: "2922fb362488fc02b02bf24b98bf3341",
    25: "b66593b5581ca76a754495d1b7d1d72d",
}
# 25 页备注文字的 md5（本任务基线各页备注均无文字，md5 为空串的 md5）
EMPTY_MD5 = "d41d8cd98f00b204e9800998ecf8427e"
BASELINE_NOTES_FP = dict((i, EMPTY_MD5) for i in range(1, SLIDE_COUNT + 1))

# ---- 第10、16、21页手写书法体篇章标题图片：图片字节 md5 + 位置 + 尺寸 ----
BASELINE_CALLIGRAPHY = {
    10: ("f0275b910c1f595f13324bd950c45e10", (106680, 1306830), (6668135, 3834765)),
    16: ("b77ce47524b33827c20e064d5dc00a25", (-86995, 1360170), (6762750, 3889375)),
    21: ("2a04f29e4acc7f8c08e2dec1e8d7715e", (-215265, 909955), (7396480, 4645660)),
}

# ---- 所提供 720×960 竖版 mp4 素材的字节 md5 ----
BASELINE_MP4_MD5 = "a7dc423f8befaeceb3d5737ba4ae0188"

# ---------------------------------------------------------------------------
# 通用工具
# ---------------------------------------------------------------------------

_PUNCT_MAP = {ord("!"): "！", ord("："): "："}


def _norm(s):
    """去掉全部空白并统一半角感叹号，便于文字比对（"篇"后有无空格均可）。"""
    return "".join((s or "").split()).translate(_PUNCT_MAP)


def _ln(el):
    """返回元素本地名；注释/处理指令等非字符串 tag 不是形状节点。"""
    if not isinstance(getattr(el, "tag", None), str):
        return ""
    return etree.QName(el).localname


def _md5(data):
    return hashlib.md5(data).hexdigest()


def _cm(v):
    return v / EMU_PER_CM


def _near(a, b, tol):
    return a is not None and b is not None and abs(a - b) <= tol


def _near_pt(p, q, tol):
    if p is None or q is None:
        return False
    return abs(p[0] - q[0]) <= tol and abs(p[1] - q[1]) <= tol


def _rect(off, ext):
    """(off, ext) -> (x0, y0, x1, y1)。"""
    if off is None or ext is None:
        return None
    return (off[0], off[1], off[0] + ext[0], off[1] + ext[1])


def _intersect(r1, r2):
    """两个矩形是否有正面积重叠（相切不算）。"""
    if r1 is None or r2 is None:
        return False
    return (r1[0] < r2[2] and r2[0] < r1[2]
            and r1[1] < r2[3] and r2[1] < r1[3])


def _off_ext(el):
    xfrm = el.find(".//a:xfrm", NS)
    if xfrm is None:
        return (None, None)
    o = xfrm.find("a:off", NS)
    e = xfrm.find("a:ext", NS)
    off = (int(o.get("x")), int(o.get("y"))) if o is not None else None
    ext = (int(e.get("cx")), int(e.get("cy"))) if e is not None else None
    return (off, ext)


def _shape_text(el):
    return _norm("".join(t.text or "" for t in el.iter("{%s}t" % NS["a"])))


def _runs(el):
    """收集形状内所有非空文字 run 的排版属性。"""
    out = []
    for r_ in el.iter("{%s}r" % NS["a"]):
        t = r_.find("a:t", NS)
        txt = _norm(t.text if t is not None else "")
        if not txt:
            continue
        rPr = r_.find("a:rPr", NS)
        info = {"text": txt, "sz": None, "b": None, "u": None, "srgb": None,
                "scheme": None, "latin": None, "ea": None,
                "hl_action": None, "hl_rid": None}
        if rPr is not None:
            info["sz"] = rPr.get("sz")
            info["b"] = rPr.get("b")
            info["u"] = rPr.get("u")
            c = rPr.find(".//a:srgbClr", NS)
            info["srgb"] = c.get("val").upper() if c is not None else None
            sc = rPr.find(".//a:schemeClr", NS)
            info["scheme"] = sc.get("val") if sc is not None else None
            l_ = rPr.find("a:latin", NS)
            e_ = rPr.find("a:ea", NS)
            info["latin"] = l_.get("typeface") if l_ is not None else None
            info["ea"] = e_.get("typeface") if e_ is not None else None
            hl = rPr.find("a:hlinkClick", NS)
            if hl is not None:
                info["hl_action"] = hl.get("action")
                info["hl_rid"] = hl.get(R_ID)
        out.append(info)
    return out


def _paras(el):
    """返回 [(algn, [run, ...]), ...]，只保留有文字的段落。"""
    out = []
    for p_ in el.iter("{%s}p" % NS["a"]):
        runs = []
        for r_ in p_.iter("{%s}r" % NS["a"]):
            t = r_.find("a:t", NS)
            txt = _norm(t.text if t is not None else "")
            if txt:
                rPr = r_.find("a:rPr", NS)
                runs.append((txt, rPr))
        if not runs:
            continue
        pPr = p_.find("a:pPr", NS)
        out.append(((pPr.get("algn") if pPr is not None else None), runs))
    return out


def _shape_hlink(el):
    """形状级超链接（挂在 p:cNvPr 上）。"""
    nv = el.find(".//p:cNvPr", NS)
    if nv is None:
        return (None, None)
    hl = nv.find("a:hlinkClick", NS)
    if hl is None:
        return (None, None)
    return (hl.get("action"), hl.get(R_ID))


def _solid_fill(el):
    spPr = el.find("p:spPr", NS)
    if spPr is None:
        return None
    fill = spPr.find("a:solidFill/a:srgbClr", NS)
    return fill.get("val").upper() if fill is not None else None


def _line_clr(el):
    spPr = el.find("p:spPr", NS)
    if spPr is None:
        return (None, None)
    c = spPr.find("a:ln//a:srgbClr", NS)
    if c is None:
        return (None, None)
    alpha = c.find("a:alpha", NS)
    return (c.get("val").upper(),
            alpha.get("val") if alpha is not None else None)


def _prst(el):
    g = el.find(".//a:prstGeom", NS)
    return g.get("prst") if g is not None else None


# ---------------------------------------------------------------------------
# 页面结构签名：用于"某一页是否被改动"的整体比对
# 覆盖形状类型、位置、尺寸、预设几何、填充/描边颜色、段落对齐、
# 以及每个文字 run 的文字、字号、加粗、倾斜、下划线、颜色、中英文字体。
# 空段落（无实际文字）不参与签名，避免仅新增空 txBody 被误判为改动。
# ---------------------------------------------------------------------------


def _ws(s):
    return "".join((s or "").split())


def _shape_sig(el):
    parts = [_ln(el)]
    off, ext = _off_ext(el)
    parts.append("off=%s ext=%s" % (off, ext))
    prst = el.find(".//a:prstGeom", NS)
    parts.append("prst=%s" % (prst.get("prst") if prst is not None else None))
    spPr = el.find("p:spPr", NS)
    if spPr is not None:
        parts.append("clr=" + ",".join(
            "%s:%s" % (_ln(c), c.get("val")) for c in spPr.iter()
            if _ln(c) in ("srgbClr", "schemeClr", "alpha")))
    for p_ in el.iter("{%s}p" % NS["a"]):
        rows = []
        has_text = False
        for r_ in p_.iter("{%s}r" % NS["a"]):
            rPr = r_.find("a:rPr", NS)
            sz = b = i = u = clr = lat = ea = None
            if rPr is not None:
                sz, b = rPr.get("sz"), rPr.get("b")
                i, u = rPr.get("i"), rPr.get("u")
                c = rPr.find(".//a:srgbClr", NS)
                clr = c.get("val") if c is not None else None
                l_ = rPr.find("a:latin", NS)
                e_ = rPr.find("a:ea", NS)
                lat = l_.get("typeface") if l_ is not None else None
                ea = e_.get("typeface") if e_ is not None else None
            t = r_.find("a:t", NS)
            txt = _ws(t.text if t is not None else "")
            if txt:
                has_text = True
            rows.append("R %s|%s|%s|%s|%s|%s|%s|%s"
                        % (txt, sz, b, i, u, clr, lat, ea))
        if not has_text:
            continue
        pPr = p_.find("a:pPr", NS)
        parts.append("P algn=%s" % (pPr.get("algn") if pPr is not None else None))
        parts.extend(rows)
    return "\n".join(parts)


_SHAPE_TAGS = ("sp", "pic", "graphicFrame", "cxnSp", "grpSp")


def _top_shapes(root):
    """spTree 下的顶层形状，按文档顺序（顺序即叠放层级，靠后者在上）。"""
    spTree = root.find(".//p:cSld/p:spTree", NS)
    if spTree is None:
        return []
    return [el for el in spTree if _ln(el) in _SHAPE_TAGS]


def _all_shapes(root):
    """展开组合后的全部形状，仅用于"某段文字是否出现在本页"这类判断。"""
    out = []

    def walk(node):
        for el in node:
            if _ln(el) in _SHAPE_TAGS:
                out.append(el)
                if _ln(el) == "grpSp":
                    walk(el)
    spTree = root.find(".//p:cSld/p:spTree", NS)
    if spTree is not None:
        walk(spTree)
    return out


def _slide_fingerprint(root):
    sigs = [_shape_sig(el) for el in _all_shapes(root)]
    return _md5("\n--\n".join(sigs).encode("utf-8"))


# ---------------------------------------------------------------------------
# 媒体字节解析：mp4 轨道分辨率、位图像素尺寸
# ---------------------------------------------------------------------------


def _mp4_track_dims(data):
    """从 moov/trak/tkhd 读取视频轨道显示分辨率（16.16 定点数）。"""
    dims = []

    def walk(off, end):
        while off + 8 <= end:
            try:
                size = struct.unpack(">I", data[off:off + 4])[0]
            except struct.error:
                return
            typ = data[off + 4:off + 8]
            hdr = 8
            if size == 1:
                if off + 16 > end:
                    return
                size = struct.unpack(">Q", data[off + 8:off + 16])[0]
                hdr = 16
            elif size == 0:
                size = end - off
            if size < hdr:
                return
            body, bend = off + hdr, min(off + size, end)
            if typ == b"tkhd" and body < bend:
                ver = data[body]
                pos = body + 4
                pos += 8 if ver == 0 else 16     # creation/modification time
                pos += 4                          # track_ID
                pos += 4                          # reserved
                pos += 4 if ver == 0 else 8       # duration
                pos += 8 + 2 + 2 + 2 + 2 + 36     # reserved/layer/group/vol/matrix
                if pos + 8 <= bend:
                    w = struct.unpack(">I", data[pos:pos + 4])[0] >> 16
                    h = struct.unpack(">I", data[pos + 4:pos + 8])[0] >> 16
                    if w > 0 and h > 0:
                        dims.append((w, h))
            if typ in (b"moov", b"trak", b"mdia"):
                walk(body, bend)
            off += size

    walk(0, len(data))
    return dims[0] if dims else None


def _image_dims(data):
    """返回位图像素尺寸；支持 PNG / JPEG / GIF / BMP，识别不了返回 None。"""
    if data[:8] == b"\x89PNG\r\n\x1a\n" and len(data) >= 24:
        return struct.unpack(">II", data[16:24])
    if data[:2] == b"BM" and len(data) >= 26:
        w, h = struct.unpack("<ii", data[18:26])
        return (abs(w), abs(h))
    if data[:6] in (b"GIF87a", b"GIF89a") and len(data) >= 10:
        return struct.unpack("<HH", data[6:10])
    if data[:2] == b"\xff\xd8":
        i, n = 2, len(data)
        while i + 9 < n:
            if data[i] != 0xFF:
                i += 1
                continue
            marker = data[i + 1]
            if marker in (0xD8, 0x01) or 0xD0 <= marker <= 0xD7:
                i += 2
                continue
            seglen = struct.unpack(">H", data[i + 2:i + 4])[0]
            if 0xC0 <= marker <= 0xCF and marker not in (0xC4, 0xC8, 0xCC):
                h, w = struct.unpack(">HH", data[i + 5:i + 9])
                return (w, h)
            i += 2 + seglen
    return None


# ---------------------------------------------------------------------------
# OOXML 包读取
# ---------------------------------------------------------------------------


def _resolve(base_dir, target):
    """把关系中的相对 Target 解析为包内绝对部件名。"""
    parts = [s for s in base_dir.split("/") if s]
    for seg in target.replace("\\", "/").split("/"):
        if seg in ("", "."):
            continue
        if seg == "..":
            if parts:
                parts.pop()
        else:
            parts.append(seg)
    return "/".join(parts)


class _Pkg(object):
    """只读方式打开 .pptx，按需解析部件。"""

    def __init__(self, path):
        self.zf = zipfile.ZipFile(path, "r")
        self.names = set(self.zf.namelist())
        self._xml_cache = {}
        self._rels_cache = {}
        self.pres = self.xml("ppt/presentation.xml")
        prels = self.rels("ppt/presentation.xml")
        self.slide_parts = []
        lst = self.pres.find("p:sldIdLst", NS)
        if lst is not None:
            for sid in lst:
                rid = sid.get(R_ID)
                if rid in prels:
                    self.slide_parts.append(prels[rid][1])

    def read(self, part):
        return self.zf.read(part)

    def has(self, part):
        return part in self.names

    def xml(self, part):
        if part not in self._xml_cache:
            self._xml_cache[part] = etree.fromstring(self.zf.read(part))
        return self._xml_cache[part]

    def rels(self, part):
        """返回 {rId: (Type, 绝对部件名或外部地址, TargetMode)}。"""
        if part in self._rels_cache:
            return self._rels_cache[part]
        head, _, tail = part.rpartition("/")
        rp = "%s/_rels/%s.rels" % (head, tail)
        out = {}
        if rp in self.names:
            for rel in self.xml(rp):
                tgt = rel.get("Target") or ""
                mode = rel.get("TargetMode") or "Internal"
                abs_t = tgt if mode == "External" else _resolve(head, tgt)
                out[rel.get("Id")] = (rel.get("Type") or "", abs_t, mode)
        self._rels_cache[part] = out
        return out

    def slide_count(self):
        return len(self.slide_parts)

    def slide(self, n):
        """第 n 页（1 起）的 XML 根节点；越界返回 None。"""
        if 1 <= n <= len(self.slide_parts):
            return self.xml(self.slide_parts[n - 1])
        return None

    def slide_rels(self, n):
        if 1 <= n <= len(self.slide_parts):
            return self.rels(self.slide_parts[n - 1])
        return {}

    def sld_sz(self):
        el = self.pres.find("p:sldSz", NS)
        if el is None:
            return None
        return (int(el.get("cx")), int(el.get("cy")))

    def notes_text(self, n):
        """第 n 页演讲者备注的纯文字（去空白）。"""
        txt = ""
        for _t, tgt, mode in self.slide_rels(n).values():
            if mode == "External" or not _t.endswith("/notesSlide"):
                continue
            if self.has(tgt):
                root = self.xml(tgt)
                txt = "".join(_ws(t.text) for t in root.iter("{%s}t" % NS["a"]))
        return txt

    def mp4_content_type_declared(self):
        """[Content_Types].xml 是否声明了 mp4 内容类型。"""
        if not self.has("[Content_Types].xml"):
            return False
        root = self.xml("[Content_Types].xml")
        for el in root:
            tag = _ln(el)
            if tag == "Default" and (el.get("Extension") or "").lower() == "mp4":
                return "video" in (el.get("ContentType") or "").lower()
            if tag == "Override" and (el.get("PartName") or "").lower().endswith(".mp4"):
                return "video" in (el.get("ContentType") or "").lower()
        return False


def _pic_image_md5(pkg, pic, rels):
    """取 p:pic 的填充图片字节 md5（取不到返回 None）。"""
    blip = pic.find(".//p:blipFill/a:blip", NS)
    if blip is None:
        blip = pic.find(".//a:blip", NS)
    if blip is None:
        return None
    rid = blip.get(R_EMBED)
    if not rid or rid not in rels:
        return None
    tgt = rels[rid][1]
    if rels[rid][2] == "External" or not pkg.has(tgt):
        return None
    return _md5(pkg.read(tgt))


def _is_video_pic(el):
    return _ln(el) == "pic" and el.find(".//a:videoFile", NS) is not None


# ---------------------------------------------------------------------------
# 形状匹配：按细则给出的尺寸/位置/几何/颜色/字号等属性寻找目标形状
# ---------------------------------------------------------------------------


def _match(shapes, kind=None, text=None, off=None, ext=None, prst=None,
           fill=None, ln_clr=None, ln_alpha=None, sz=None, bold=None,
           color=None, algn=None, para_count=None, tol=TOL_GEOM,
           exclude=None):
    for el in shapes:
        if exclude is not None and exclude(el):
            continue
        if kind is not None and _ln(el) not in kind:
            continue
        if text is not None and _shape_text(el) != _norm(text):
            continue
        o, e = _off_ext(el)
        if off is not None and not _near_pt(o, off, tol):
            continue
        if ext is not None and not _near_pt(e, ext, tol):
            continue
        if prst is not None and _prst(el) != prst:
            continue
        if fill is not None and _solid_fill(el) != fill.upper():
            continue
        if ln_clr is not None or ln_alpha is not None:
            c, alpha = _line_clr(el)
            if ln_clr is not None and c != ln_clr.upper():
                continue
            if ln_alpha is not None and alpha != ln_alpha:
                continue
        rs = _runs(el)
        if sz is not None and (not rs or any(r["sz"] != sz for r in rs)):
            continue
        if bold is not None and (not rs or any((r["b"] or "0") != bold for r in rs)):
            continue
        if color is not None and (not rs or any(r["srgb"] != color.upper() for r in rs)):
            continue
        ps = _paras(el)
        if algn is not None and (not ps or any(p[0] != algn for p in ps)):
            continue
        if para_count is not None and len(ps) != para_count:
            continue
        return el
    return None


def _entry_link(el, rels):
    """取目录条目的跳转设置：优先 run 级 a:hlinkClick，其次形状级 p:cNvPr。"""
    action = rid = None
    for r in _runs(el):
        if r["hl_action"] or r["hl_rid"]:
            action, rid = r["hl_action"], r["hl_rid"]
            break
    if action is None and rid is None:
        action, rid = _shape_hlink(el)
    if action is None and rid is None:
        return None
    rel = rels.get(rid) if rid else None
    return {"action": action or "", "rid": rid,
            "type": rel[0] if rel else None,
            "target": rel[1] if rel else None,
            "mode": rel[2] if rel else None}


def _timing_info(slide, spid):
    """解析 p:timing 中与指定媒体对象相关的播放设置。"""
    info = {"media_node": False, "show_when_stopped": None, "media_dur": None,
            "click_trigger": False, "play_cmd": None}
    if slide is None or not spid:
        return info
    tm = slide.find("p:timing", NS)
    if tm is None:
        return info
    for cmn in tm.iter("{%s}cMediaNode" % NS["p"]):
        tgt = cmn.find(".//p:tgtEl/p:spTgt", NS)
        if tgt is None or tgt.get("spid") != spid:
            continue
        info["media_node"] = True
        info["show_when_stopped"] = cmn.get("showWhenStopped")
        ctn = cmn.find("p:cTn", NS)
        if ctn is not None:
            info["media_dur"] = ctn.get("dur")
        break
    for seq in tm.iter("{%s}seq" % NS["p"]):
        ctn = seq.find("p:cTn", NS)
        if ctn is None or ctn.get("nodeType") != "interactiveSeq":
            continue
        cond_ok = False
        for cond in ctn.findall("p:stCondLst/p:cond", NS):
            if cond.get("evt") != "onClick":
                continue
            st = cond.find("p:tgtEl/p:spTgt", NS)
            if st is not None and st.get("spid") == spid:
                cond_ok = True
        if not cond_ok:
            continue
        for sub in ctn.iter("{%s}cTn" % NS["p"]):
            if (sub.get("presetClass") != "mediacall"
                    or sub.get("nodeType") != "clickEffect"):
                continue
            st = sub.find(".//p:tgtEl/p:spTgt", NS)
            if st is None or st.get("spid") != spid:
                continue
            info["click_trigger"] = True
            cmd = sub.find(".//p:cmd", NS)
            if cmd is not None:
                info["play_cmd"] = (cmd.get("cmd") or "").lower()
    return info


def _evaluate_impl(file_path, result):
    results = []

    # -----------------------------------------------------------------------
    # 维度一：可用与可修改性
    # 交付文件必须是一个能被 PowerPoint 正常打开的 .pptx：
    # ①是合法 OPC 包；②含 ppt/presentation.xml 且能解析；
    # ③幻灯片清单非空、每一页 XML 均可解析；④声明了页面尺寸。
    # -----------------------------------------------------------------------
    try:
        pkg = _Pkg(file_path)
        if pkg.slide_count() <= 0:
            raise ValueError("演示文稿中不含任何幻灯片")
        if pkg.sld_sz() is None:
            raise ValueError("演示文稿缺少页面尺寸定义 p:sldSz")
        for i in range(1, pkg.slide_count() + 1):
            pkg.slide(i)
    except Exception as exc:
        result["dim1_pass"] = False
        result["dim1_reason"] = ("交付文件无法作为 PowerPoint 演示文稿正常打开：%s: %s"
                                % (type(exc).__name__, exc))
        result["dim2_items"] = []
        result["total_score"] = 0
        result["max_score"] = MAX_SCORE_TOTAL
        return result

    result["dim1_pass"] = True
    result["dim1_reason"] = "交付文件为 1 个可正常打开的 PowerPoint 演示文稿（.pptx）"

    # -----------------------------------------------------------------------
    # 公共上下文
    # -----------------------------------------------------------------------
    s8 = pkg.slide(8)
    s9 = pkg.slide(9)
    s8_top = _top_shapes(s8) if s8 is not None else []
    s8_all = _all_shapes(s8) if s8 is not None else []
    s9_top = _top_shapes(s9) if s9 is not None else []
    s8_rels = pkg.slide_rels(8)
    s9_rels = pkg.slide_rels(9)

    # 目录标题与三条篇章条目
    title = None
    for el in s8_all:
        if _shape_text(el) == CATALOG_TITLE:
            title = el
            break
    title_off, title_ext = _off_ext(title) if title is not None else (None, None)
    title_rect = _rect(title_off, title_ext)
    title_runs = _runs(title) if title is not None else []
    title_paras = _paras(title) if title is not None else []

    entry_hits = {}
    for el in s8_all:
        t = _shape_text(el)
        if t in ENTRY_TEXTS:
            entry_hits.setdefault(t, []).append(el)
    entries = [entry_hits.get(t, [None])[0] for t in ENTRY_TEXTS]
    entry_boxes = []
    for el in entries:
        if el is None:
            entry_boxes.append((None, None))
        else:
            entry_boxes.append(_off_ext(el))
    entry_centers = []
    for off, ext in entry_boxes:
        if off is None:
            entry_centers.append(None)
        else:
            entry_centers.append(off[1] + (ext[1] if ext else 0) / 2.0)
    all_entries_found = all(el is not None for el in entries)

    # 第9页视频对象
    video = None
    video_idx = -1
    for i, el in enumerate(s9_top):
        if _is_video_pic(el):
            video, video_idx = el, i
            break
    video_off, video_ext = _off_ext(video) if video is not None else (None, None)
    video_rect = _rect(video_off, video_ext)
    video_spid = None
    if video is not None:
        nv = video.find(".//p:cNvPr", NS)
        video_spid = nv.get("id") if nv is not None else None

    # 视频所引用的媒体部件：优先取 p14:media 的 r:embed（包内嵌入副本）
    media_part = None
    media_internal = False
    if video is not None:
        rids = []
        m = video.find(".//p14:media", NS)
        if m is not None and m.get(R_EMBED):
            rids.append(m.get(R_EMBED))
        vf = video.find(".//a:videoFile", NS)
        if vf is not None and vf.get(R_LINK):
            rids.append(vf.get(R_LINK))
        for rid in rids:
            rel = s9_rels.get(rid)
            if rel is None:
                continue
            if rel[2] != "External" and pkg.has(rel[1]):
                media_part = rel[1]
                media_internal = True
                break
            if media_part is None:
                media_part = rel[1]
    media_bytes = pkg.read(media_part) if (media_part and media_internal) else None
    media_md5 = _md5(media_bytes) if media_bytes else None
    media_dims = _mp4_track_dims(media_bytes) if media_bytes else None
    timing = _timing_info(s9, video_spid)

    # -----------------------------------------------------------------------
    # 维度二：完成度（加分项 01—16）
    # -----------------------------------------------------------------------

    # +5：第8页原有的五处文字全部删除
    # 严格对齐细则，仅踩细则明确写出的点：①第8页任何形状（含组合内形状）的文字
    # 里都不得再出现细则列出的五条文字中的任何一条；②比对前统一去除空白字符，
    # 避免换行、空格造成的假差异。
    # 与办公软件对齐：遍历 p:sld//a:t 的文字内容，等价于 PowerPoint 中该页可见
    # 的全部文本框文字。
    page8_text = "".join(_shape_text(el) for el in s8_all)
    left_old = [t for t in S8_OLD_TEXTS if _norm(t) in page8_text]
    if s8 is not None and not left_old:
        results.append("+5：第8页原有的五处文字全部删除")
    else:
        results.append("+0：第8页原有的五处文字全部删除（残留 %d 处）" % len(left_old))

    # +5：第8页新增"目录"二字作为页面标题，位于三条篇章标题上方
    # 严格对齐细则，仅踩细则明确写出的点：①第8页存在文字为"目录"的文字块；
    # ②该文字块垂直中心高于三条篇章标题的垂直中心，且其顶边不低于三条标题的顶边。
    # 与办公软件对齐：读 a:t 文字与 a:off/@y、a:ext/@cy 计算版面上下关系。
    hit02 = False
    if title is not None and title_off is not None:
        known = [c for c in entry_centers if c is not None]
        tops = [b[0][1] for b in entry_boxes if b[0] is not None]
        t_center = title_off[1] + (title_ext[1] if title_ext else 0) / 2.0
        if known and tops:
            hit02 = t_center < min(known) and title_off[1] <= min(tops)
    if hit02:
        results.append("+5：第8页新增\"目录\"二字作为页面标题，位于三条篇章标题上方")
    else:
        results.append("+0：第8页新增\"目录\"二字作为页面标题，位于三条篇章标题上方")

    # +3："目录"二字为宋体、96号（sz=9600，允许95—97号；latin 与 ea 的 typeface
    #      均为"宋体"）
    # 严格对齐细则，仅踩细则明确写出的点：①"目录"文字块内每一个文字 run 的 sz 都
    # 落在 9500—9700，未显式声明字号（继承占位符）不算达标；②每个 run 的
    # a:latin/@typeface 为"宋体"，且 a:ea/@typeface 也为"宋体"——中文字形由 a:ea
    # 决定，缺少 a:ea 时中文会回落到主题的中日韩字体，只写 a:latin 不算达标；
    # ③字号与中文字体两个条件同时满足才计分。
    # 与办公软件对齐：读 a:rPr/@sz（单位为百分之一磅）、a:rPr/a:latin/@typeface
    # 与 a:rPr/a:ea/@typeface。
    title_sz_ok = bool(title_runs) and all(
        r["sz"] is not None and CATALOG_TITLE_SZ_MIN <= int(r["sz"]) <= CATALOG_TITLE_SZ_MAX
        for r in title_runs)
    title_font_ok = bool(title_runs) and all(
        r["latin"] == CATALOG_TITLE_FONT and r["ea"] == CATALOG_TITLE_FONT
        for r in title_runs)
    results.append(("+3：" if (title_sz_ok and title_font_ok) else "+0：")
                   + "\"目录\"二字为宋体、96号（sz=9600，允许95—97号；latin 与 ea 的 "
                     "typeface 均为\"宋体\"）")

    # +3："目录"二字水平居中显示（段落 algn="ctr"，且文字块水平中心落在其所处
    #      版面区域水平中线 ±1 厘米内）
    # 严格对齐细则，仅踩细则明确写出的点：①"目录"的有字段落 algn 均为 "ctr"；
    # ②文字块水平中心与"其所处版面区域"的水平中线相差不超过 1 厘米。版面区域
    # 取横向覆盖该文字块中心的纯色填充色块（取最贴合的那一块），页面整幅也作为
    # 一个候选区域，任一候选达标即算居中。
    # 与办公软件对齐：读 a:pPr/@algn、a:off/@x、a:ext/@cx 及色块 a:solidFill。
    hit04 = False
    if title is not None and title_rect is not None:
        algn_ok = bool(title_paras) and all(p[0] == "ctr" for p in title_paras)
        h_center = (title_rect[0] + title_rect[2]) / 2.0
        regions = [(SLD_SZ[0], 0, SLD_SZ[0])]
        for el in s8_top:
            if el is title or _shape_text(el) or _solid_fill(el) is None:
                continue
            o, e = _off_ext(el)
            if o is None or e is None or e[0] <= 0:
                continue
            if o[0] - TOL_GEOM <= h_center <= o[0] + e[0] + TOL_GEOM:
                regions.append((e[0], o[0], o[0] + e[0]))
        hit04 = algn_ok and any(abs(h_center - (x0 + x1) / 2.0) <= TOL_1CM
                                for _w, x0, x1 in regions)
    if hit04:
        results.append("+3：\"目录\"二字水平居中显示")
    else:
        results.append("+0：\"目录\"二字水平居中显示")

    # +5：目录正文列出且仅列出三条篇章大标题，无错字、无遗漏、无多余条目，
    #      未把非篇章标题混入
    # 严格对齐细则，仅踩细则明确写出的点：①"第一篇 探索！""第二篇 讨论"
    # "第三篇 作业"三条文字各出现且仅出现一次（"篇"字后有无空格均可）；
    # ②全页形如"第N篇"的条目总数正好是 3，没有多出来的篇章条目；
    # ③页面上不得混入"教学理念与立意""教学过程设计"等栏目名。
    # 与办公软件对齐：比对各文字块 a:t 拼接后的文字（去空白）。
    chapter_pat = re.compile(r"第[一二三四五六七八九十0-9]+篇")
    chapter_like = [el for el in s8_all if chapter_pat.search(_shape_text(el))]
    mixed_in = [el for el in s8_all
                if any(w in _shape_text(el) for w in NON_CHAPTER_WORDS)]
    hit05 = (all(len(entry_hits.get(t, [])) == 1 for t in ENTRY_TEXTS)
             and len(chapter_like) == 3 and not mixed_in)
    if hit05:
        results.append("+5：目录正文列出且仅列出三条篇章大标题，无错字、无遗漏、无多余条目")
    else:
        results.append("+0：目录正文列出且仅列出三条篇章大标题，无错字、无遗漏、无多余条目")

    # +3：三条篇章标题文字均为仿宋、44号（sz=4400；latin 与 ea 的 typeface 均为"仿宋"）
    # 严格对齐细则，仅踩细则明确写出的点：①三条条目都存在；②每条内每个 run 的
    # a:rPr/@sz 均为 4400；③每条内每个 run 的 a:latin 为"仿宋"，且 a:ea 也为
    # "仿宋"——只写 a:latin 时中文实际按主题的中日韩字体渲染，不算达标；
    # ④字号与中文字体两个条件同时满足才计分。
    # 与办公软件对齐：读 a:rPr/@sz、a:rPr/a:latin/@typeface 与 a:rPr/a:ea/@typeface。
    entry_runs = [(_runs(el) if el is not None else []) for el in entries]
    entry_sz_ok = all_entries_found and all(
        rs and all(r["sz"] == ENTRY_SZ for r in rs) for rs in entry_runs)
    entry_font_ok = all_entries_found and all(
        rs and all(r["latin"] == ENTRY_FONT and r["ea"] == ENTRY_FONT for r in rs)
        for rs in entry_runs)
    results.append(("+3：" if (entry_sz_ok and entry_font_ok) else "+0：")
                   + "三条篇章标题文字均为仿宋、44号（sz=4400；latin 与 ea 的 "
                     "typeface 均为\"仿宋\"）")

    # +3：三条篇章标题位于"目录"下方，按"第一篇—第二篇—第三篇"自上而下纵向排列，
    #      间距均匀（相邻条目垂直中心间距互差 ≤0.5 厘米），条目之间及与标题之间
    #      无文字重叠，且未溢出页面边界
    # 严格对齐细则，仅踩细则明确写出的点：①三条条目的垂直中心严格递增，即顺序为
    # 第一篇—第二篇—第三篇；②三条条目的垂直中心都低于"目录"文字块的垂直中心；
    # ③两段相邻间距之差不超过 0.5 厘米；④三条条目彼此、以及条目与"目录"标题的
    # 外框互不重叠；⑤四个文字块都落在页面矩形内。
    # 与办公软件对齐：读 a:off/@x、@y 与 a:ext/@cx、@cy 做矩形运算，页面矩形取
    # ppt/presentation.xml 的 p:sldSz。
    hit07 = False
    if all_entries_found and all(c is not None for c in entry_centers):
        order_ok = entry_centers[0] < entry_centers[1] < entry_centers[2]
        below_ok = False
        if title_off is not None:
            t_center = title_off[1] + (title_ext[1] if title_ext else 0) / 2.0
            below_ok = all(c > t_center for c in entry_centers)
        even_ok = abs((entry_centers[1] - entry_centers[0])
                      - (entry_centers[2] - entry_centers[1])) <= TOL_05CM
        rects = [_rect(off, ext) for off, ext in entry_boxes]
        if title_rect is not None:
            rects.append(title_rect)
        rects = [r for r in rects if r is not None]
        overlap = any(_intersect(rects[a], rects[b])
                      for a in range(len(rects)) for b in range(a + 1, len(rects)))
        page_sz = pkg.sld_sz() or SLD_SZ
        inside = all(r[0] >= -TOL_GEOM and r[1] >= -TOL_GEOM
                     and r[2] <= page_sz[0] + TOL_GEOM
                     and r[3] <= page_sz[1] + TOL_GEOM
                     for r in rects)
        hit07 = order_ok and below_ok and even_ok and not overlap and inside
    results.append(("+3：" if hit07 else "+0：")
                   + "三条篇章标题位于\"目录\"下方，按\"第一篇—第二篇—第三篇\"自上而下"
                     "纵向排列，间距均匀，条目之间及与标题之间无文字重叠，"
                     "且未溢出页面边界")

    # +5：三条篇章标题均设有幻灯片内部跳转超链接，放映时单击一次即可跳转
    # 严格对齐细则，仅踩细则明确写出的点：①三条条目都带 a:hlinkClick（挂在文字
    # run 上或挂在形状 p:cNvPr 上都算，两者放映时都是单击一次触发）；
    # ②action 为 "ppaction://hlinksldjump"；③其 r:id 指向的关系类型为
    # .../relationships/slide，且目标是包内的幻灯片部件。
    # 与办公软件对齐：读 a:hlinkClick/@action 与 @r:id，并在
    # ppt/slides/_rels/slideN.xml.rels 中查该关系的 Type/Target/TargetMode。
    links = [(_entry_link(el, s8_rels) if el is not None else None) for el in entries]

    def _link_ok(lk):
        return (lk is not None
                and lk["action"] == HLINK_SLIDE_JUMP
                and lk["type"] == REL_TYPE_SLIDE
                and lk["mode"] != "External"
                and bool(lk["target"]) and pkg.has(lk["target"]))

    hit08 = all_entries_found and all(_link_ok(lk) for lk in links)
    results.append(("+5：" if hit08 else "+0：")
                   + "三条篇章标题均设有幻灯片内部跳转超链接，放映时单击一次即可跳转")

    # +5：三条跳转目标分别正确指向第10页、第16页、第21页（slide10.xml /
    #      slide16.xml / slide21.xml），无错页、无错位对应
    # 严格对齐细则，仅踩细则明确写出的点：①每条条目的跳转关系必须是合规的幻灯片
    # 内部跳转（见上一项）；②Target 解析出的目标页序号分别等于 10、16、21，三条
    # 全对才计分——若被设成网页地址、打开文件、"下一页/上一页"等非指定页面的动作，
    # 或三条互相错位，本项即不得分。
    # 与办公软件对齐：把 rels 中的 Target 解析为包内绝对部件名，再在
    # ppt/presentation.xml 的 p:sldIdLst 顺序中定位其页序号。
    def _target_page(lk):
        if not _link_ok(lk):
            return None
        for i, part in enumerate(pkg.slide_parts, 1):
            if part == lk["target"]:
                return i
        m = re.search(r"slide(\d+)\.xml$", lk["target"])
        return int(m.group(1)) if m else None

    hit09 = all(_target_page(links[i]) == want
                for i, want in enumerate(ENTRY_TARGETS))
    results.append(("+5：" if hit09 else "+0：")
                   + "三条跳转目标分别正确指向第10页、第16页、第21页，"
                     "无错页、无错位对应")

    # +5：第9页插入了 mp4 视频媒体对象，视频文件本体随演示文稿一并嵌入包内
    #      （ppt/media 下新增 mp4，且 [Content_Types].xml 已声明 mp4 内容类型）
    # 严格对齐细则，仅踩细则明确写出的点：①第9页存在带 a:videoFile 的 p:pic 媒体
    # 对象；②它引用的媒体部件是包内 ppt/media 下的 .mp4（不是外部链接）；
    # ③[Content_Types].xml 中声明了 mp4 的 video 内容类型。
    # 与办公软件对齐：读 p:pic//a:videoFile/@r:link 与 p14:media/@r:embed，
    # 在 slide9 的 rels 中解析 Target/TargetMode，并检查 [Content_Types].xml。
    ct_ok = pkg.mp4_content_type_declared()
    embed_ok = bool(media_part) and media_internal \
        and media_part.lower().startswith("ppt/media/") \
        and media_part.lower().endswith(".mp4")
    hit10 = video is not None and embed_ok and ct_ok
    results.append(("+5：" if hit10 else "+0：")
                   + "第9页插入了 mp4 视频媒体对象，视频文件本体随演示文稿一并嵌入包内")

    # +5：嵌入的视频与所提供的 720×960 竖版 mp4 素材字节完全一致，未被重新编码为
    #      其他分辨率、未改变画面宽高、未裁剪画面、未加模糊填充边条、未横向拉伸变形
    # 严格对齐细则，仅踩细则明确写出的点：包内媒体部件的字节 md5 必须与所提供素材
    # 的字节 md5 完全一致——只要重编码成别的分辨率（如 1126×960 并在左右补模糊
    # 填充条）、裁剪画面或横向拉伸变形，字节必然改变，本项即不得分。
    # 与办公软件对齐：直接取包内 ppt/media/*.mp4 的原始字节做 md5，并解析 mp4 的
    # moov/trak/tkhd 显示宽高（16.16 定点数）核对画面仍为 720×960。
    dims_ok = media_dims is None or media_dims == (MP4_W, MP4_H)
    hit11 = media_md5 is not None and media_md5 == BASELINE_MP4_MD5 and dims_ok
    results.append(("+5：" if hit11 else "+0：")
                   + "嵌入的视频与所提供的 720×960 竖版 mp4 素材字节完全一致，"
                     "未被重新编码为其他分辨率、未加模糊填充边条、未横向拉伸变形")

    # +5：视频对象宽 14.0—14.5 厘米、高 11.9—12.2 厘米
    # 严格对齐细则，仅踩细则明确写出的点：a:ext/@cx 落在 5040000—5220000 EMU、
    # @cy 落在 4284000—4392000 EMU。
    # 与办公软件对齐：读 p:pic/p:spPr/a:xfrm/a:ext。
    hit12 = (video_ext is not None
             and VIDEO_CX_MIN <= video_ext[0] <= VIDEO_CX_MAX
             and VIDEO_CY_MIN <= video_ext[1] <= VIDEO_CY_MAX)
    results.append(("+5：" if hit12 else "+0：")
                   + "视频对象宽 14.0—14.5 厘米、高 11.9—12.2 厘米")

    # +5：视频对象左上角距页面左边界 2.0—2.2 厘米、距上边界 4.8—5.0 厘米，与左侧
    #      插画区域四边重合、各边偏差 ≤0.1 厘米，完整覆盖该区域不露底、不留空白边；
    #      且右边界不超过页面左起 17.3 厘米，未压盖右侧米色圆角文字卡片、
    #      未超出页面范围
    # 严格对齐细则，仅踩细则明确写出的点：①a:off/@x 落在 2.0—2.2 厘米、@y 落在
    # 4.8—5.0 厘米；②视频外框四条边与左侧插画区域（距左 2.08 厘米、距上 4.88
    # 厘米、宽 14.22 厘米、高 12.06 厘米）逐边相差不超过 0.1 厘米，即完整覆盖该
    # 区域、既不露底也不留空白边；③视频右边界 ≤ 6228000 EMU（页面左起 17.3
    # 厘米）；④视频外框与右侧米色卡片区域（距左 17.27 厘米、宽 14.35 厘米、
    # 高 12.06 厘米）无重叠；⑤视频外框完整落在页面矩形内。
    # 与办公软件对齐：读 a:off/a:ext 计算左、上、右、下四边坐标后逐边比对，
    # 页面矩形取 ppt/presentation.xml 的 p:sldSz。
    hit13 = False
    if video_off is not None and video_rect is not None:
        pos_ok = (VIDEO_X_MIN <= video_off[0] <= VIDEO_X_MAX
                  and VIDEO_Y_MIN <= video_off[1] <= VIDEO_Y_MAX)
        target = _rect(S9_ILLUS_OFF, S9_ILLUS_EXT)
        edge_ok = all(abs(video_rect[k] - target[k]) <= TOL_01CM for k in range(4))
        page_sz = pkg.sld_sz() or SLD_SZ
        bound_ok = (video_rect[2] <= VIDEO_RIGHT_LIMIT
                    and not _intersect(video_rect, _rect(S9_CARD_OFF, S9_CARD_EXT))
                    and video_rect[0] >= -TOL_GEOM and video_rect[1] >= -TOL_GEOM
                    and video_rect[2] <= page_sz[0] + TOL_GEOM
                    and video_rect[3] <= page_sz[1] + TOL_GEOM)
        hit13 = pos_ok and edge_ok and bound_ok
    results.append(("+5：" if hit13 else "+0：")
                   + "视频对象左上角位置与左侧插画区域四边重合、各边偏差 ≤0.1 厘米，"
                     "且右边界不超过页面左起 17.3 厘米，未压盖右侧米色圆角文字卡片、"
                     "未超出页面范围")

    # +5：为视频设置了单击触发，放映时单击该区域一次即触发视频并连续自动播放至
    #      结束，无需再次点击播放按钮或媒体控制条
    # 严格对齐细则，仅踩细则明确写出的点：①p:timing 中存在 nodeType
    # ="interactiveSeq" 的 p:seq；②其 p:cond evt="onClick" 的 p:spTgt 指向该视频
    # 对象；③其内部存在 presetClass="mediacall" 且 nodeType="clickEffect" 的节点，
    # 其 p:spTgt 同样指向该视频对象；④该视频存在对应的 p:cMediaNode 播放节点；
    # ⑤播放节点未被有限 dur 截断（dur 缺省或 indefinite 表示播到媒体自然结束）；
    # ⑥单击效果下发的是播放类命令（togglePause / play / playFrom），而非暂停或
    # 停止——单击即播且一次播到底，两段条件都满足才计分。
    # 与办公软件对齐：读 p:timing//p:seq/p:cTn 的 nodeType、p:stCondLst/p:cond
    # 的 @evt 与 p:tgtEl/p:spTgt/@spid，clickEffect 节点下的 p:spTgt/@spid 与
    # p:cmd/@cmd，以及 p:cMediaNode/p:cTn/@dur。
    click_ok = video is not None and timing["click_trigger"]
    autoplay_ok = (timing["media_node"]
                   and timing["media_dur"] in (None, "indefinite")
                   and (timing["play_cmd"] is None or timing["play_cmd"] in PLAY_CMDS))
    hit14 = click_ok and autoplay_ok
    results.append(("+5：" if hit14 else "+0：")
                   + "为视频设置了单击触发，放映时单击该区域一次即触发视频并连续"
                     "自动播放至结束，无需再次点击播放按钮或媒体控制条")

    # +3：视频设置为停止时不显示（p:cMediaNode 的 showWhenStopped="0"）
    # 严格对齐细则，仅踩细则明确写出的点：该视频对应的 p:cMediaNode 上
    # showWhenStopped 属性显式为 "0"。
    # 与办公软件对齐：读 p:timing//p:video/p:cMediaNode/@showWhenStopped，
    # 并以 p:tgtEl/p:spTgt/@spid 确认它对应的是该视频对象。
    hit15 = video is not None and timing["show_when_stopped"] == "0"
    results.append(("+3：" if hit15 else "+0：")
                   + "视频设置为停止时不显示（p:cMediaNode 的 showWhenStopped=\"0\"）")

    # +1：视频封面首帧为该视频自身画面（宽高比与 720×960 素材一致），
    #      非黑屏、非灰底占位图、非无关图片
    # 严格对齐细则，仅踩细则明确写出的点：①视频对象带封面图（p:blipFill/a:blip）；
    # ②封面图在包内存在；③封面图像素宽高比与 720:960（0.75）相差不超过 0.03，
    # 即确为该竖版视频自身画面的截帧，而非横版占位图或无关图片。
    # 与办公软件对齐：读 p:pic/p:blipFill/a:blip/@r:embed 定位图片部件，
    # 再从图片字节头解析像素宽高。
    hit16 = False
    if video is not None:
        blip = video.find(".//p:blipFill/a:blip", NS)
        rid = blip.get(R_EMBED) if blip is not None else None
        rel = s9_rels.get(rid) if rid else None
        if rel is not None and rel[2] != "External" and pkg.has(rel[1]):
            dims = _image_dims(pkg.read(rel[1]))
            if dims and dims[1]:
                hit16 = abs(dims[0] / float(dims[1]) - POSTER_RATIO) <= POSTER_RATIO_TOL
    results.append(("+1：" if hit16 else "+0：")
                   + "视频封面首帧为该视频自身画面（宽高比与 720×960 素材一致）")

    # -----------------------------------------------------------------------
    # 维度二：扣分项（17—32）
    # 约定："-N：…" 表示该扣分项被触发（真的扣了 N 分）；"-0：…" 表示未触发。
    # -----------------------------------------------------------------------

    # -3：第8页右侧未保留宽14.48厘米、高16.89厘米、左上角距页面左边界18.34厘米、
    #      上边界0.97厘米的深色木桌静物摄影配图
    # 严格对齐细则，仅踩细则明确写出的点：第8页必须仍有一个图片对象，其左上角在
    # (18.34, 0.97) 厘米、尺寸为 14.48×16.89 厘米（各值容差 0.1 厘米）。
    # 与办公软件对齐：读 p:pic/p:spPr/a:xfrm 的 a:off 与 a:ext。
    photo = _match(s8_top, kind=("pic",), off=S8_PHOTO_OFF, ext=S8_PHOTO_EXT)
    results.append(("-0：" if photo is not None else "-3：")
                   + "第8页右侧未保留宽14.48厘米、高16.89厘米、左上角距页面左边界"
                     "18.34厘米、上边界0.97厘米的深色木桌静物摄影配图")

    # -3：第8页左侧未保留宽17.91厘米、高19.05厘米、填充色 #173B57 的深蓝色竖向
    #      色块（被横贯整页的满版矩形或纯色背景替代）
    # 严格对齐细则，仅踩细则明确写出的点：第8页必须仍有一个左上角在 (0,0)、
    # 尺寸 17.91×19.05 厘米、纯色填充为 173B57 的矩形；把它换成横贯整页的满版
    # 矩形（宽 33.87 厘米）时，尺寸不符，本项即触发。
    # 与办公软件对齐：读 a:off/a:ext 与 p:spPr/a:solidFill/a:srgbClr/@val。
    navy = _match(s8_top, kind=("sp",), off=S8_NAVY_OFF, ext=S8_NAVY_EXT,
                  fill=S8_NAVY_FILL)
    navy_rect = _rect(*_off_ext(navy)) if navy is not None else None
    results.append(("-0：" if navy is not None else "-3：")
                   + "第8页左侧未保留宽17.91厘米、高19.05厘米、填充色 #173B57 的"
                     "深蓝色竖向色块")

    # -1：第8页右下角未保留"01"页码文字（微软雅黑9.5磅、颜色 #697987、右对齐、
    #      位于距左边界31.37厘米、上边界17.91厘米处）及其右侧长0.64厘米的浅灰色
    #      （#697987，70%透明）短横线
    # 严格对齐细则，仅踩细则明确写出的点：①存在文字为"01"、字号 950、字色
    # 697987、段落右对齐、左上角在 (31.37, 17.91) 厘米的文字块；②存在左上角在
    # (32.56, 18.21) 厘米、长 0.64 厘米（cx=228600、cy=0）、描边色 697987 且透明度
    # 70% 的线条；③页码与短横线属同一处页脚装饰，缺任何一个即触发本项一次。
    # 与办公软件对齐：读 a:t、a:rPr/@sz、a:rPr//a:srgbClr/@val、a:pPr/@algn、
    # a:off/a:ext 以及 p:spPr/a:ln//a:srgbClr 及其 a:alpha/@val。
    pageno8 = _match(s8_top, text=S8_PAGENO_TEXT, off=S8_PAGENO_OFF,
                     ext=S8_PAGENO_EXT, sz=S8_PAGENO_SZ, color=S8_PAGENO_CLR,
                     algn="r")
    dash8 = _match(s8_top, off=S8_DASH_OFF, ext=S8_DASH_EXT,
                   ln_clr=S8_DASH_CLR, ln_alpha=S8_DASH_ALPHA)
    results.append(("-0：" if (pageno8 is not None and dash8 is not None) else "-1：")
                   + "第8页右下角未保留\"01\"页码文字（微软雅黑9.5磅、颜色 #697987、"
                     "右对齐、位于距左边界31.37厘米、上边界17.91厘米处）及其右侧长"
                     "0.64厘米的浅灰色（#697987，70%透明）短横线")

    # -1：第8页新增了指令未要求的装饰图形或说明性文字（如横贯版面的金色矩形色条、
    #      新增边框色块，或"点击篇章标题可跳转"之类操作提示语）
    # 严格对齐细则，仅踩细则明确写出的点：①统计第8页无文字的非图片形状，若其外框
    # 尺寸不属于该页原有图形的尺寸集合（深蓝色块、摄影配图、标题下金色短线、页码
    # 短横线），即为新增装饰图形——满版矩形（几乎铺满整页）不在本项范围内，它由
    # "深蓝色竖向色块未保留"单独计分，避免同一处改动被重复扣分；②第8页有文字的
    # 形状里，除"目录"、三条篇章标题、页码"01"以及该页原有的五处文字之外，出现任何
    # 其它文字即为新增说明性文字；③两类新增同属"擅自加东西"，任一出现即触发一次。
    # 与办公软件对齐：读各形状 a:ext、a:t 是否为空及其拼接后的文字（去空白）。
    page = pkg.sld_sz() or SLD_SZ
    new_decor = []
    for el in s8_top:
        if _shape_text(el) or _ln(el) == "pic":
            continue
        _o, e = _off_ext(el)
        if e is None:
            continue
        if any(_near_pt(e, b, TOL_GEOM) for b in S8_BASE_GRAPHIC_EXTS):
            continue
        if e[0] >= page[0] * 0.95 and e[1] >= page[1] * 0.95:
            continue
        new_decor.append(el)
    allowed_texts = set(ENTRY_TEXTS) | set([CATALOG_TITLE, S8_PAGENO_TEXT])
    allowed_texts |= set(_norm(t) for t in S8_OLD_TEXTS)
    extra_texts = [el for el in s8_all
                   if _shape_text(el) and _shape_text(el) not in allowed_texts]
    results.append(("-1：" if (new_decor or extra_texts) else "-0：")
                   + "第8页新增了指令未要求的装饰图形或说明性文字（如横贯版面的金色"
                     "矩形色条、新增边框色块，或\"点击篇章标题可跳转\"之类操作提示语）")

    # -1：第8页目录文字块横向超出深蓝色块范围，压到右侧摄影配图之上
    # 严格对齐细则，仅踩细则明确写出的点：仅当该页仍保留深色木桌摄影配图时才谈得上
    # "压到配图之上"；此时若"目录"文字块右边界越过深蓝色块右边界、且其外框与配图
    # 外框发生重叠，即触发。
    # 与办公软件对齐：读三者的 a:off/a:ext 做矩形运算。
    hit21 = False
    if title_rect is not None and photo is not None:
        photo_rect = _rect(*_off_ext(photo))
        right_limit = navy_rect[2] if navy_rect is not None else page[0]
        hit21 = (title_rect[2] > right_limit + TOL_GEOM
                 and _intersect(title_rect, photo_rect))
    results.append(("-1：" if hit21 else "-0：")
                   + "第8页目录文字块横向超出深蓝色块范围，压到右侧摄影配图之上")

    # -3：第9页左侧宽14.22厘米、高12.06厘米的帆船码头插画，或其上等大的白色半透明
    #      描边圆角矩形边框被删除、缩放或移位
    # 严格对齐细则，仅踩细则明确写出的点：①第9页必须仍有一个非视频图片对象，左上角
    # 在 (2.08, 4.88) 厘米、尺寸 14.22×12.06 厘米（容差 0.1 厘米）；②必须仍有一个
    # 与插画同位同尺寸、预设几何为 roundRect、描边色为白色且透明度 80% 的形状；
    # ③两者同属"触发前呈现的插画区域"，缺任何一个即触发本项一次。
    # 与办公软件对齐：读 p:pic 与 p:sp 的 a:off/a:ext、a:prstGeom/@prst 与
    # p:spPr/a:ln//a:srgbClr 及 a:alpha/@val，并用 a:videoFile 的有无区分插画与视频。
    illus = _match(s9_top, kind=("pic",), off=S9_ILLUS_OFF, ext=S9_ILLUS_EXT,
                   exclude=_is_video_pic)
    frame = _match(s9_top, kind=("sp",), off=S9_ILLUS_OFF, ext=S9_ILLUS_EXT,
                   prst="roundRect", ln_clr=S9_FRAME_LN_CLR,
                   ln_alpha=S9_FRAME_LN_ALPHA)
    results.append(("-0：" if (illus is not None and frame is not None) else "-3：")
                   + "第9页左侧宽14.22厘米、高12.06厘米的帆船码头插画，或其上等大的"
                     "白色半透明描边圆角矩形边框被删除、缩放或移位")

    # -1：第9页视频对象被置于插画或白色圆角描边框之下（层级过低），
    #      放映时被其他元素遮挡
    # 严格对齐细则，仅踩细则明确写出的点：spTree 的文档顺序即叠放层级（越靠后越
    # 靠上）；若插画或白色描边框排在视频对象之后，视频就会被压在下面。
    # 与办公软件对齐：比较 p:spTree 子元素的文档顺序。
    hit23 = False
    if video is not None:
        for other in (illus, frame):
            if other is not None and other in s9_top and s9_top.index(other) > video_idx:
                hit23 = True
    results.append(("-1：" if hit23 else "-0：")
                   + "第9页视频对象被置于插画或白色圆角描边框之下（层级过低），"
                     "放映时被其他元素遮挡")

    # -3：第9页顶部四项导航文字或其下方金色（#D6A642）分隔线被改动、移位或删除
    # 严格对齐细则，仅踩细则明确写出的点：①"教学理念与立意""教学前端分析"
    # "教学过程设计""教学评价"四个文字块都在原位（各自左上角、尺寸 5.97×0.64
    # 厘米）、字号 2000、字色白色、段落居中；②其下方 y=1.17 厘米处的金色
    # （D6A642）横线仍在。
    # 与办公软件对齐：读 a:t、a:off/a:ext、a:rPr/@sz、a:rPr//a:srgbClr、a:pPr/@algn
    # 以及线条的 p:spPr/a:ln//a:srgbClr。
    nav_ok = all(_match(s9_top, kind=("sp",), text=t, off=o, ext=S9_NAV_EXT,
                        sz=S9_NAV_SZ, color=S9_NAV_CLR, algn="ctr") is not None
                 for t, o in zip(S9_NAV_TEXTS, S9_NAV_OFFS))
    gold_line = _match(s9_top, off=S9_GOLD_OFF, ext=S9_GOLD_EXT,
                       ln_clr=S9_GOLD_CLR)
    results.append(("-0：" if (nav_ok and gold_line is not None) else "-3：")
                   + "第9页顶部四项导航文字或其下方金色（#D6A642）分隔线被改动、"
                     "移位或删除")

    # -3：第9页左上角标题"情境导入：一封来自港口的委托信"（微软雅黑28磅加粗、颜色
    #      #173B57）、其下副标题"左侧图片、右侧文字，形成观察—设问—入课的图文结构"
    #      （微软雅黑13.5磅、颜色 #697987）或副标题下方长5.08厘米的褐色（#B36B2C）
    #      粗短线被改动、移位或删除
    # 严格对齐细则，仅踩细则明确写出的点：①标题文字完全一致、左上角在 (1.88, 1.93)
    # 厘米、尺寸 19.81×1.27 厘米、字号 2800、加粗、字色 173B57；②副标题文字完全
    # 一致、左上角在 (1.98, 3.25) 厘米、字号 1350、字色 697987；③存在左上角在
    # (1.98, 4.17) 厘米、长 5.08 厘米（cx=1828800、cy=0）、描边色 B36B2C 的线条；
    # ④三者同属该页左上角标题组，任一被改动即触发本项一次。
    # 与办公软件对齐：读 a:t、a:off/a:ext、a:rPr 的 @sz/@b 与 a:srgbClr/@val，
    # 以及线条的 p:spPr/a:ln//a:srgbClr/@val。
    s9_title = _match(s9_top, kind=("sp",), text=S9_TITLE_TEXT, off=S9_TITLE_OFF,
                      ext=S9_TITLE_EXT, sz=S9_TITLE_SZ, bold="1",
                      color=S9_TITLE_CLR)
    s9_sub = _match(s9_top, kind=("sp",), text=S9_SUB_TEXT, off=S9_SUB_OFF,
                    ext=S9_SUB_EXT, sz=S9_SUB_SZ, color=S9_SUB_CLR)
    s9_brown = _match(s9_top, off=S9_BROWN_OFF, ext=S9_BROWN_EXT,
                      ln_clr=S9_BROWN_CLR)
    head_ok = (s9_title is not None and s9_sub is not None
               and s9_brown is not None)
    results.append(("-0：" if head_ok else "-3：")
                   + "第9页左上角标题\"情境导入：一封来自港口的委托信\"、其下副标题"
                     "\"左侧图片、右侧文字，形成观察—设问—入课的图文结构\"或副标题下方"
                     "长5.08厘米的褐色（#B36B2C）粗短线被改动、移位或删除")

    # -3：第9页右侧宽14.35厘米、高12.06厘米、填充色 #FFF8EC 的米色圆角卡片，或卡片
    #      内"致年轻的航海学徒："、三段正文、"导入意图：以港口场景激活学生经验，
    #      提出本课核心问题。"被改动、遮挡或删除
    # 严格对齐细则，仅踩细则明确写出的点：①米色圆角卡片在原位、原尺寸、填充
    # FFF8EC；②卡内标题文字"致年轻的航海学徒："在原位、字号 1800、加粗；
    # ③卡内正文仍为 3 段、字号 1550、字色 27313A；④"导入意图：…"说明文字完全
    # 一致、左上角在 (18.29, 15.44) 厘米、字号 1250；⑤卡片区域未被新增元素压盖
    # （新增元素指外框不在该页原有元素外框集合中的形状，例如放歪了的视频）。
    # 与办公软件对齐：读 a:prstGeom、a:solidFill/a:srgbClr、a:off/a:ext、a:t、
    # a:rPr 的 @sz/@b 与 a:srgbClr，段落数取有文字的 a:p 个数。
    card = _match(s9_top, kind=("sp",), off=S9_CARD_OFF, ext=S9_CARD_EXT,
                  prst="roundRect", fill=S9_CARD_FILL)
    card_title = _match(s9_top, kind=("sp",), text=S9_CARD_TITLE_TEXT,
                        off=S9_CARD_TITLE_OFF, sz=S9_CARD_TITLE_SZ, bold="1")
    card_body = _match(s9_top, kind=("sp",), off=S9_BODY_OFF, ext=S9_BODY_EXT,
                       sz=S9_BODY_SZ, color=S9_BODY_CLR,
                       para_count=S9_BODY_PARA_COUNT)
    card_rect = _rect(S9_CARD_OFF, S9_CARD_EXT)
    card_covered = False
    for el in s9_top:
        o, e = _off_ext(el)
        if o is None or e is None:
            continue
        if any(_near_pt(o, br[0], TOL_GEOM) and _near_pt(e, br[1], TOL_GEOM)
               for br in S9_BASE_RECTS):
            continue
        if _intersect(_rect(o, e), card_rect):
            card_covered = True
    card_ok = (card is not None and card_title is not None
               and card_body is not None and not card_covered)
    s9_intent = _match(s9_top, kind=("sp",), text=S9_INTENT_TEXT,
                       off=S9_INTENT_OFF, sz=S9_INTENT_SZ)
    results.append(("-0：" if (card_ok and s9_intent is not None) else "-3：")
                   + "第9页右侧宽14.35厘米、高12.06厘米、填充色 #FFF8EC 的米色圆角卡片，"
                     "或卡片内\"致年轻的航海学徒：\"、三段正文、\"导入意图：以港口场景"
                     "激活学生经验，提出本课核心问题。\"被改动、遮挡或删除")

    # -1：第9页右下角"09"页码（微软雅黑9.5磅）及其右侧0.64厘米短横线被删除
    # 严格对齐细则，仅踩细则明确写出的点：①存在文字"09"、字号 950、左上角在
    # (31.37, 17.91) 厘米的文字块；②存在其右侧长 0.64 厘米的短横线。
    # 与办公软件对齐：读 a:t、a:rPr/@sz、a:off/a:ext。
    pageno9 = _match(s9_top, text=S9_PAGENO_TEXT, off=S9_PAGENO_OFF,
                     ext=S9_PAGENO_EXT, sz=S9_PAGENO_SZ)
    dash9 = _match(s9_top, off=S9_DASH_OFF, ext=S9_DASH_EXT)
    results.append(("-0：" if (pageno9 is not None and dash9 is not None) else "-1：")
                   + "第9页右下角\"09\"页码（微软雅黑9.5磅）及其右侧0.64厘米短横线被删除")

    # -3：除第8、9页以外的任意页面，其文字内容、字体、字号、颜色、元素位置或图片被
    #      改动（含第10、16、21页手写书法体篇章标题"第一篇 探索！""第二篇 讨论"
    #      "第三篇 作业"被替换、删除或改变尺寸与位置）
    # 严格对齐细则，仅踩细则明确写出的点：①逐页计算结构签名并与基线比对，签名覆盖
    # 形状类型、位置、尺寸、预设几何、填充与描边颜色、段落对齐，以及每个文字 run
    # 的文字、字号、加粗、倾斜、下划线、颜色和中英文字体——正好是细则点到的
    # "文字内容、字体、字号、颜色或元素位置"；仅新增空文本框这类无语义差异不计。
    # ②第10、16、21页那三张手写书法标题图片必须仍在——图片字节 md5 不变
    # （未被替换）、左上角与尺寸不变（未移位、未缩放）。任一条不满足即触发本项一次。
    # 与办公软件对齐：签名的每一个字段都直接取自 PowerPoint 读取的同名 XML 属性；
    # 图片经 p:pic/p:blipFill/a:blip/@r:embed 定位部件后取字节 md5，位置尺寸读
    # 同一 p:pic 的 a:off/a:ext。
    changed_pages = []
    for idx in range(1, pkg.slide_count() + 1):
        if idx in (8, 9):
            continue
        base = BASELINE_SLIDE_FP.get(idx)
        if base is None:
            continue
        root = pkg.slide(idx)
        if root is None or _slide_fingerprint(root) != base:
            changed_pages.append(idx)
    calli_bad = []
    for pno, (want_md5, want_off, want_ext) in sorted(BASELINE_CALLIGRAPHY.items()):
        root = pkg.slide(pno)
        found = False
        if root is not None:
            rels = pkg.slide_rels(pno)
            for pic in root.iter("{%s}pic" % NS["p"]):
                if _pic_image_md5(pkg, pic, rels) != want_md5:
                    continue
                o, e = _off_ext(pic)
                if _near_pt(o, want_off, TOL_GEOM) and _near_pt(e, want_ext, TOL_GEOM):
                    found = True
                    break
        if not found:
            calli_bad.append(pno)
    results.append(("-3：" if (changed_pages or calli_bad) else "-0：")
                   + "除第8、9页以外的任意页面，其文字内容、字体、字号、颜色、元素位置"
                     "或图片被改动（含第10、16、21页手写书法体篇章标题被替换、删除或"
                     "改变尺寸与位置）")

    # -3：演示文稿总页数不再是25页（新增目录页、拆分页面或删除页面）
    # 严格对齐细则，仅踩细则明确写出的点：p:sldIdLst 中的幻灯片条目数必须仍是 25。
    # 与办公软件对齐：读 ppt/presentation.xml 的 p:sldIdLst 子元素个数。
    results.append(("-0：" if pkg.slide_count() == SLIDE_COUNT else "-3：")
                   + "演示文稿总页数不再是25页（新增目录页、拆分页面或删除页面）")

    # -1：页面尺寸不再是33.87×19.05厘米（12192000×6858000 EMU）
    # 严格对齐细则，仅踩细则明确写出的点：p:sldSz 的 cx/cy 必须精确等于
    # 12192000×6858000。
    # 与办公软件对齐：读 ppt/presentation.xml 的 p:sldSz/@cx 与 @cy。
    results.append(("-0：" if pkg.sld_sz() == SLD_SZ else "-1：")
                   + "页面尺寸不再是33.87×19.05厘米（12192000×6858000 EMU）")

    # -1：任意页面的演讲者备注内容被清空或改写
    # 严格对齐细则，仅踩细则明确写出的点：逐页取备注页的纯文字做 md5，与基线比对，
    # 任何一页不一致（被清空或被改写）即触发。
    # 与办公软件对齐：沿 slideN.xml.rels 的 notesSlide 关系找到备注页，
    # 读其 a:t 文字。
    notes_changed = []
    for idx in range(1, pkg.slide_count() + 1):
        base = BASELINE_NOTES_FP.get(idx)
        if base is None:
            continue
        if _md5(pkg.notes_text(idx).encode("utf-8")) != base:
            notes_changed.append(idx)
    results.append(("-1：" if notes_changed else "-0：")
                   + "任意页面的演讲者备注内容被清空或改写")

    # -1：为实现跳转而给三条篇章标题文字添加了下划线、改变字色为超链接主题色等
    #      未经要求的样式改动
    # 严格对齐细则，仅踩细则明确写出的点：三条条目的文字 run 上出现 u（下划线，
    # u="none" 除外），或字色被设为主题的超链接色（a:schemeClr val="hlink"
    # 或 "folHlink"），即触发。
    # 与办公软件对齐：读 a:rPr/@u 与 a:rPr//a:schemeClr/@val。
    hit32 = False
    for rs in entry_runs:
        for r in rs:
            if r["u"] not in (None, "none"):
                hit32 = True
            if r["scheme"] in ("hlink", "folHlink"):
                hit32 = True
    results.append(("-1：" if hit32 else "-0：")
                   + "为实现跳转而给三条篇章标题文字添加了下划线、改变字色为超链接"
                     "主题色等未经要求的样式改动")


    # -----------------------------------------------------------------------
    # 汇总：results 顺序必须与 _MAX_DELTAS_ORDER 一一对应
    # -----------------------------------------------------------------------
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
        hit = delta != 0
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


def evaluate(dir_path):
    result = {
        "id": SCRIPT_ID,
        "file_name": "",
        "status": "ok",
        "error": None,
        "dim1_pass": "",
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
            raise FileNotFoundError("目录中未找到 .pptx 文件：%s" % dir_path)
        file_path = os.path.join(dir_path, cands[0])
        result["file_name"] = os.path.basename(file_path)
        return _evaluate_impl(file_path, result)
    except Exception as exc:
        # 找不到文件/无法定位被评估对象：维度一未进入评判，dim1_pass 留空，
        # dim1_reason 留空，满分与得分均为 0。
        result["status"] = "error"
        result["error"] = "%s" % exc
        result["dim1_pass"] = ""
        result["dim1_reason"] = ""
        result["dim2_items"] = []
        result["total_score"] = 0
        result["max_score"] = 0
        return result


if __name__ == "__main__":
    if len(sys.argv) >= 2:
        _dir = sys.argv[1]
    else:
        _dir = os.path.dirname(os.path.abspath(__file__))
    _out = evaluate(_dir)
    print(json.dumps(_out, ensure_ascii=False, indent=2))
