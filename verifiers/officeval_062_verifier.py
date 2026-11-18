#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# pyright: reportUnknownMemberType=false, reportUnknownVariableType=false, reportUnknownArgumentType=false, reportUnknownParameterType=false, reportMissingParameterType=false, reportPrivateUsage=false, reportUnknownLambdaType=false, reportGeneralTypeIssues=false

import os
import re
import sys
import json
import hashlib
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.enum.text import PP_ALIGN

# 说明：按接口统一约定，本模块不修改全局 sys.stdout，也不 print 主结果。
# 主结果统一由 evaluate(dir_path) -> dict 返回；入参为"脚本所在目录的路径"，
# 由脚本自己在该目录内定位并打开被评估的 .pptx。本文件仅在 __main__ 下作为
# 本地调试入口，才把 dict 序列化为 JSON 打印。
#
# 本脚本只读取被评估的 .pptx，全部判定都在该文件内部自洽完成（坐标、字体、
# 颜色、层级、媒体关系均取自文件本身），不与任何输入文件或原始素材做比对。
#
# 返回信息约定：
#   status="error"（目录不存在、目录中没有 .pptx、脚本自身异常）表示这份作业
#   根本没评上，此时 dim1_pass 为空串（未判定）、dim1_reason 为空串、
#   dim2_items 为 []、total_score 与 max_score 均为 0；
#   维度一未通过时 status 仍为 "ok"，dim1_pass 为 False、dim1_reason 写错误摘要、
#   dim2_items 为 []（未真正逐条评过）、total_score 为 0、max_score 为满分；
#   维度一通过时 dim1_pass 为 True、dim1_reason 保持空串。
#   dim2_items 每条只含 rule / max_delta / delta / hit 四个键。

SCRIPT_ID = "062"
# 维度二所有"加分项"max_delta 之和；用于维度一未通过 / 脚本异常时兜底
MAX_SCORE_TOTAL = 67

# 每一项的满分（加分项为正、减分项为负），顺序必须与下方 results.append 严格一致。
_MAX_DELTAS_ORDER = [
    5,   # 01 2-15 页顶部装饰组构成与摆放位置一致
    5,   # 02 2-15 页顶部平行四边形数字块
    1,   # 03 数字块蓝色系填充、白色数字且不小于 24 磅
    3,   # 04 数字块内数字与页码对应递增 02-15
    3,   # 05 数字块右侧衔接紧凑横条底板/蓝色细线框
    3,   # 06 装饰整体外接矩形宽高比 2.5:1-7:1 且宽度不超过 16.9 厘米
    3,   # 07 装饰整体外接矩形位于页面左上角
    3,   # 08 中文标题文本框紧邻数字块右侧、落在横幅高度内
    3,   # 09 中文标题文字与页码对应
    1,   # 10 中文标题不小于 30 磅且加粗
    3,   # 11 中文标题下方/右侧的英文文字 10-13 磅蓝色系
    3,   # 12 英文文字为对应译名且 14 页互不相同
    3,   # 13 横幅右端不少于 9 个小点呈 3 行 3 列点阵
    1,   # 14 3 个平行小斜条 + 1 个独立小斜块
    3,   # 15 第 16 页含独立"谢谢观看"文本对象
    3,   # 16 "谢谢观看"中西文字体均为宋体
    1,   # 17 "谢谢观看"加粗
    1,   # 18 "谢谢观看"蓝色系
    3,   # 19 "谢谢观看"段落居中且文本框水平居中
    5,   # 20 第 16 页两个嵌入到文件内部的可播放视频
    3,   # 21 两个视频指向两个不同的媒体文件
    5,   # 22 两个视频左右并列、文本在上、完整位于页内
    1,   # 23 两个视频宽度差与高度差不超过 0.5 厘米
    1,   # 24 两个视频宽高比 0.71-0.79
    1,   # 25 两个视频均带首帧预览画面
    -5,  # 26 顶部标题栏装饰形状覆盖页面已有元素
    -1,  # 27 第 2 页页面之外的 0×0 空白形状
    -1,  # 28 第 14 页两处"风险管理"且一处超出页面右边界
    -1,  # 29 第 15 页正文段落位置
    -1,  # 30 第 16 页圆角矩形边框对象
    -1,  # 31 第 16 页页码内容不为"16"
    -3,  # 32 第 1 页主标题格式与位置
    -1,  # 33 第 1 页副标题格式与位置
    -1,  # 34 第 1 页右侧照片位置与尺寸
    -3,  # 35 2-15 页左下角页脚格式与位置
    -1,  # 36 2-15 页右下角页码格式与位置
    -3,  # 37 2-15 页导语段落格式
    -3,  # 38 2-15 页条目正文段落格式
    -3,  # 39 第 4/6/7/8/12/13 页各含 1 个可编辑图表
    -3,  # 40 2-15 页照片位置与尺寸
    -1,  # 41 第 3 页品牌关键词色块填充色与边线色
    -1,  # 42 第 9 页"协同重点"文字块
    -1,  # 43 第 2/10/14 页条目左侧深灰色小竖条
    -1,  # 44 页面尺寸 33.86×19.05 厘米
]
assert sum(d for d in _MAX_DELTAS_ORDER if d > 0) == MAX_SCORE_TOTAL

# ---------------------------------------------------------------- 常量与词表

EMU_PER_CM = 360000.0
PAGE_W_CM = 33.86          # 页面宽（厘米）
PAGE_H_CM = 19.05          # 页面高（厘米）
TOP_BAND_CM = 2.6          # "顶部区域"：距页面上边界 2.6 厘米以内
HALF_W_CM = 16.9           # 页面宽度的一半：紧凑横幅右边界上限
CENTER_X_CM = 16.93        # 页面水平中心

# 第 2 页至第 15 页依次对应的中文标题
TITLES = [
    "内容框架", "品牌定位", "消费场景", "产品矩阵", "数字化运营", "会员运营",
    "门店模型", "供应链协同", "品质控制", "增长路径", "营销策略", "效率提升",
    "风险管理", "总结展望",
]

# 各页英文文字须为该页中文标题的英文译名：以译名关键词（小写子串）判定，
# 兼容 "brand positioning" / "Brand Positioning" / "Brand Position" 等写法。
EN_KEYWORDS = {
    "内容框架": ("content", "framework", "outline"),
    "品牌定位": ("brand", "position"),
    "消费场景": ("consum", "scenario", "scene", "occasion"),
    "产品矩阵": ("product", "matrix", "portfolio"),
    "数字化运营": ("digital", "operation"),
    "会员运营": ("member", "loyalty", "operation"),
    "门店模型": ("store", "shop", "model"),
    "供应链协同": ("supply", "chain", "collaborat", "synerg", "coordinat"),
    "品质控制": ("quality", "control"),
    "增长路径": ("growth", "path", "route"),
    "营销策略": ("marketing", "strateg"),
    "效率提升": ("efficien", "improve", "enhance", "upgrad"),
    "风险管理": ("risk", "manage", "control"),
    "总结展望": ("summary", "outlook", "conclusion", "prospect", "future"),
}

# 第 2 页至第 15 页照片的"距左/距上/宽/高"（厘米）
PHOTO_SPECS = {
    2: (20.95, 2.54, 11.30, 13.97), 3: (1.78, 2.67, 14.48, 13.84),
    4: (18.41, 2.41, 13.59, 14.73), 5: (1.65, 2.54, 13.21, 14.48),
    6: (2.03, 2.41, 13.46, 14.99), 7: (18.29, 2.41, 13.72, 15.11),
    8: (18.29, 2.41, 13.72, 15.11), 9: (1.78, 2.41, 15.11, 14.86),
    10: (17.65, 2.41, 14.35, 15.11), 11: (18.41, 2.54, 13.21, 14.48),
    12: (1.78, 2.41, 14.61, 14.86), 13: (1.65, 2.41, 13.59, 15.11),
    14: (18.16, 2.41, 13.59, 15.11), 15: (18.03, 0.00, 15.82, 19.05),
}

CHART_PAGES = (4, 6, 7, 8, 12, 13)     # 须各含 1 个可编辑图表对象的页
BAR_PAGES = (2, 10, 14)                # 条目左侧须含深灰色小竖条的页

BODY_FONT = "Noto Sans CJK SC"         # 正文/标题/页脚统一字体
SONG = ("宋体", "SimSun", "NSimSun")   # 宋体及其等价西文族名

P15_TEXT = ("晨岚咖啡的增长关键，不是追逐单一爆品或短期声量，"
            "而是把稳定出品、便利触达和轻量会员关系长期做细。")
P9_TEXT = "协同重点：少压货、快流转、可追踪"
DOC_TITLE = "晨岚咖啡品牌运营研究"
SUB_TITLE = "城市轻饮场景的产品、渠道与运营重构"

# 蓝色系：命中蓝色主题色或常见蓝色色值（用于装饰形状填充/线条的粗判）
BLUE_XML_RE = re.compile(
    r'srgbClr val="(?:2962FF|347BDC|4472C4|5590E2|6CA1EB|2360B9|88B5F5'
    r'|BCD7FB|BCD6FA|C6DEFF|00B0F0|1F4E79|305496)"'
    r'|schemeClr val="accent(?:1|5)"'
)
CJK_RE = re.compile(r'[一-鿿]')
EN_RE = re.compile(r"^[A-Za-z][A-Za-z0-9 &\-/',\.]*$")

# ---------------------------------------------------------------- 通用取值工具


def _cm(v):
    """EMU -> 厘米，保留 2 位小数（与办公软件"位置和大小"面板显示口径一致）。"""
    return round((v or 0) / EMU_PER_CM, 2)


def _box(sh):
    """返回 (左, 上, 右, 下)，单位厘米。"""
    left = _cm(sh.left)
    top = _cm(sh.top)
    return (left, top, left + _cm(sh.width), top + _cm(sh.height))


def _wh(sh):
    return (_cm(sh.width), _cm(sh.height))


def _inter_area(a, b):
    """两个矩形的重叠面积（平方厘米）。"""
    x1 = max(a[0], b[0])
    y1 = max(a[1], b[1])
    x2 = min(a[2], b[2])
    y2 = min(a[3], b[3])
    if x2 <= x1 or y2 <= y1:
        return 0.0
    return (x2 - x1) * (y2 - y1)

def _area(b):
    return max(0.0, b[2] - b[0]) * max(0.0, b[3] - b[1])


def _xml(sh):
    try:
        return sh._element.xml
    except Exception:
        return ""


def _txt(sh):
    """形状内文字（去空白）。无文本框返回空串。"""
    try:
        if not sh.has_text_frame:
            return ""
    except Exception:
        return ""
    try:
        return (sh.text_frame.text or "").strip()
    except Exception:
        return ""


def _prst(sh):
    """预置几何类型（parallelogram / roundRect / rect / diagStripe ...）。
    自由曲线（custGeom）返回 'custGeom'，取不到返回 None。"""
    x = _xml(sh)
    m = re.search(r'<a:prstGeom[^>]+prst="([^"]+)"', x)
    if m:
        return m.group(1)
    if "<a:custGeom" in x:
        return "custGeom"
    return None


def _stype(sh):
    try:
        return sh.shape_type
    except Exception:
        return None


def _is_pic(sh):
    return _stype(sh) == MSO_SHAPE_TYPE.PICTURE


def _has_chart(sh):
    try:
        return bool(sh.has_chart)
    except Exception:
        return False

def _runs(sh):
    out = []
    try:
        for p in sh.text_frame.paragraphs:
            out.extend(list(p.runs))
    except Exception:
        pass
    return out


def _sig_runs(sh):
    """"有效 run"：文本长度大于 1 个字符。
    与办公软件对齐：段落末尾常被切分出单独一个标点 run（如"。"），
    其字号/字体沿用段落属性且不承载可视格式，逐字符比对易误判，故不计。"""
    rs = [r for r in _runs(sh) if len(r.text.strip()) > 1]
    if rs:
        return rs
    return [r for r in _runs(sh) if r.text.strip()]


def _run_size(run):
    try:
        return run.font.size.pt if run.font.size is not None else None
    except Exception:
        return None


def _run_bold(run):
    try:
        return bool(run.font.bold)
    except Exception:
        return False


def _run_latin(run):
    m = re.search(r'<a:latin[^>]+typeface="([^"]*)"', run._r.xml)
    return m.group(1) if m else None


def _run_ea(run):
    m = re.search(r'<a:ea[^>]+typeface="([^"]*)"', run._r.xml)
    return m.group(1) if m else None


def _run_color(run):
    """run 的文字颜色：直接色值返回 6 位大写十六进制，主题色返回 'scheme:xxx'。"""
    x = run._r.xml
    m = re.search(r'<a:solidFill>\s*<a:srgbClr val="([0-9A-Fa-f]{6})"', x)
    if m:
        return m.group(1).upper()
    m = re.search(r'<a:solidFill>\s*<a:schemeClr val="([A-Za-z0-9]+)"', x)
    if m:
        return "scheme:" + m.group(1)
    return None

def _hex_is_blue(h):
    """蓝色系：蓝通道足够高且明显强于红通道、不低于绿通道。"""
    try:
        r = int(h[0:2], 16)
        g = int(h[2:4], 16)
        b = int(h[4:6], 16)
    except Exception:
        return False
    return b >= 0x80 and b > r + 0x20 and b >= g


def _is_blue_color(c):
    if not c:
        return False
    if c.startswith("scheme:"):
        return c.split(":", 1)[1] in ("accent1", "accent5", "hlink")
    return _hex_is_blue(c)


def _is_near_hex(c, refs, tol=0x50):
    """色值近似判定：与任一参考色的各通道偏差都在 tol 以内。"""
    if not c or c.startswith("scheme:"):
        return False
    try:
        cur = [int(c[i:i + 2], 16) for i in (0, 2, 4)]
    except Exception:
        return False
    for ref in refs:
        rv = [int(ref[i:i + 2], 16) for i in (0, 2, 4)]
        if all(abs(a - b) <= tol for a, b in zip(cur, rv)):
            return True
    return False


def _is_white(c):
    if not c:
        return False
    if c.startswith("scheme:"):
        return c.split(":", 1)[1] in ("bg1", "lt1")
    try:
        return all(int(c[i:i + 2], 16) >= 0xF0 for i in (0, 2, 4))
    except Exception:
        return False


def _is_song(v):
    return bool(v) and any(k in v for k in SONG)

def _shape_blue(sh):
    """形状自身（填充或线条）是否为蓝色系。"""
    return bool(BLUE_XML_RE.search(_xml(sh)))


def _shape_light(sh):
    """形状是否为浅色填充（任一显式色值的平均亮度不低于 0xD0）。"""
    for h in re.findall(r'srgbClr val="([0-9A-Fa-f]{6})"', _xml(sh)):
        try:
            v = (int(h[0:2], 16) + int(h[2:4], 16) + int(h[4:6], 16)) / 3.0
        except Exception:
            continue
        if v >= 0xD0:
            return True
    return False


def _line_w_pt(sh):
    """线宽（磅）。未显式设置返回 None（沿用主题默认细线）。"""
    m = re.search(r'<a:ln[^>]*\sw="(\d+)"', _xml(sh))
    return round(int(m.group(1)) / 12700.0, 2) if m else None


def _clusters(vals, tol=0.1):
    """把一维坐标按容差聚成若干簇（用于识别点阵的行/列）。"""
    groups = []
    for v in sorted(vals):
        if groups and abs(v - groups[-1][-1]) <= tol:
            groups[-1].append(v)
        else:
            groups.append([v])
    return groups

# --------------------------------------------------------- 顶部区域形状集合


def _band_shapes(slide):
    """顶部区域形状：整体（含高度）都落在距上边界 2.6 厘米以内。
    与办公软件对齐：只有完整位于顶部标题栏内的对象才算标题栏构件，
    仅"上边界在 2.6 厘米内"会把正文照片（上边界 2.41 厘米）误纳入。"""
    out = []
    for sh in slide.shapes:
        left, top, _, bottom = _box(sh)
        if left > 100 or top > 100:      # 页面之外的幽灵形状不参与
            continue
        if bottom <= TOP_BAND_CM:
            out.append(sh)
    return out

def _decor(slide):
    """顶部装饰形状：顶部区域内不承载文字的对象（排除标题、英文、数字文本）。"""
    return [sh for sh in _band_shapes(slide) if not _txt(sh)]


def _decor_bbox(slide):
    """装饰图形整体外接矩形；不计宽高均为 0 厘米的空形状。"""
    boxes = []
    for sh in _decor(slide):
        w, h = _wh(sh)
        if w == 0 and h == 0:
            continue
        boxes.append(_box(sh))
    if not boxes:
        return None
    return (min(b[0] for b in boxes), min(b[1] for b in boxes),
            max(b[2] for b in boxes), max(b[3] for b in boxes))


def _num_shapes(slide):
    """顶部区域内文字为两位数字的对象（数字块，不限几何类型）。"""
    return [sh for sh in _band_shapes(slide)
            if re.fullmatch(r"\d{2}", _txt(sh))]


def _num_para(slide):
    """顶部区域内几何类型为 parallelogram 且文字为两位数字的对象。"""
    return [sh for sh in _num_shapes(slide) if _prst(sh) == "parallelogram"]


def _num_right(slide):
    """数字块右边界（无数字块时返回 0）。"""
    return max([_box(sh)[2] for sh in _num_shapes(slide)], default=0.0)


def _title_shape(slide, expected):
    """按预期标题文字定位中文标题文本框，优先取落在顶部区域的那个。"""
    cands = [sh for sh in slide.shapes if _txt(sh) == expected]
    band = [sh for sh in cands if _cm(sh.top) <= TOP_BAND_CM]
    if band:
        return band[0]
    return cands[0] if cands else None


def _en_shape(slide):
    """标题栏英文文字对象：页面上部纯英文（含空格/连字符）的最上方文本对象。"""
    best = None
    for sh in slide.shapes:
        t = _txt(sh)
        if len(t) < 4 or not EN_RE.fullmatch(t):
            continue
        if _cm(sh.top) > 6.0:
            continue
        if best is None or _cm(sh.top) < _cm(best.top):
            best = sh
    return best

def _fmt_ok(sh, font=None, size=None, bold=None, color=None,
            left=None, top=None, pos_tol=0.05, size_tol=0.6):
    """逐项校验一个文本对象的字体族、字号、加粗、颜色与页面坐标。
    返回 (是否全部满足, 差异说明)。字号容差 0.6 磅，坐标容差 0.05 厘米。"""
    if sh is None:
        return (False, "对象缺失")
    bad = []
    if left is not None and abs(_cm(sh.left) - left) > pos_tol:
        bad.append("距左边界%.2f≠%.2f" % (_cm(sh.left), left))
    if top is not None and abs(_cm(sh.top) - top) > pos_tol:
        bad.append("距上边界%.2f≠%.2f" % (_cm(sh.top), top))
    rs = _sig_runs(sh)
    if not rs:
        bad.append("无文本内容")
    for r in rs:
        if font is not None:
            la, ea = _run_latin(r), _run_ea(r)
            hit = (la == font or ea == font)
            clean = (la is None or la == font) and (ea is None or ea == font)
            if not (hit and clean):
                bad.append("字体%s/%s≠%s" % (la or "-", ea or "-", font))
        if size is not None:
            sz = _run_size(r)
            if sz is None or abs(sz - size) > size_tol:
                bad.append("字号%s≠%s" % (sz, size))
        if bold is not None and _run_bold(r) != bold:
            bad.append("加粗%s≠%s" % (_run_bold(r), bold))
        if color is not None and _run_color(r) != color:
            bad.append("颜色%s≠%s" % (_run_color(r) or "-", color))
    return (not bad, "；".join(list(dict.fromkeys(bad))[:3]))


def _para_align(sh, keyword=None):
    """取（含关键字的）段落的对齐方式。"""
    try:
        for p in sh.text_frame.paragraphs:
            if keyword is None or keyword in (p.text or ""):
                return p.alignment
    except Exception:
        pass
    return None

def _shape_has_fill(sh):
    """形状自身是否设置了非"无填充"的填充（用于识别色块类对象）。
    形状填充在 spPr 中先于线条出现，取 spPr 内第一个填充节点判定。"""
    m = re.search(r'<p:spPr[ >].*?</p:spPr>', _xml(sh), re.S)
    if not m:
        return False
    f = re.search(r'<a:(solidFill|gradFill|blipFill|pattFill|grpFill|noFill)',
                  m.group(0))
    return bool(f) and f.group(1) != "noFill"


def _body_paras(slide):
    """正文区中文段落（导语 + 条目）。
    过滤掉顶部标题栏（上边界 2.2 厘米以内）、页脚与页码（上边界 17 厘米以下），
    以及带填充的色块类对象——色块承载的是关键词标签而非正文段落。"""
    out = []
    for sh in slide.shapes:
        if not _txt(sh):
            continue
        left, top, _, _ = _box(sh)
        if left > 100 or top < 2.2 or top > 17.0:
            continue
        if _stype(sh) not in (MSO_SHAPE_TYPE.TEXT_BOX, MSO_SHAPE_TYPE.AUTO_SHAPE):
            continue
        if _shape_has_fill(sh):
            continue
        try:
            paras = list(sh.text_frame.paragraphs)
        except Exception:
            continue
        for p in paras:
            t = (p.text or "").strip()
            if not t or not CJK_RE.search(t):
                continue
            rs = [r for r in p.runs if len(r.text.strip()) > 1]
            if not rs:
                rs = [r for r in p.runs if r.text.strip()]
            if not rs:
                continue
            out.append((sh, t, rs))
    return out


def _para_fmt_bad(rs, font, size, color, bold=None, size_tol=0.6):
    """段落 run 级格式校验，返回差异说明（空串表示全部满足）。"""
    bad = []
    for r in rs:
        la, ea = _run_latin(r), _run_ea(r)
        hit = (la == font or ea == font)
        clean = (la is None or la == font) and (ea is None or ea == font)
        if not (hit and clean):
            bad.append("字体%s/%s≠%s" % (la or "-", ea or "-", font))
        sz = _run_size(r)
        if sz is None or abs(sz - size) > size_tol:
            bad.append("字号%s≠%s" % (sz, size))
        if bold is not None and _run_bold(r) != bold:
            bad.append("加粗%s≠%s" % (_run_bold(r), bold))
        if _run_color(r) != color:
            bad.append("颜色%s≠%s" % (_run_color(r) or "-", color))
    return "；".join(list(dict.fromkeys(bad))[:3])

# ------------------------------------------------------------- 媒体对象工具


def _is_media(sh):
    """可播放视频媒体对象：媒体形状，或图形内含 videoFile / p14:media 引用。"""
    if _stype(sh) == MSO_SHAPE_TYPE.MEDIA:
        return True
    x = _xml(sh)
    return ("<a:videoFile" in x) or ("p14:media" in x)


def _media_rids(sh):
    """返回 (视频关系 id 列表, 嵌入媒体数据关系 id 列表)。"""
    x = _xml(sh)
    links = re.findall(r'<a:videoFile[^>]+r:link="([^"]+)"', x)
    embeds = re.findall(r'<p14:media[^>]+r:embed="([^"]+)"', x)
    return (links, embeds)


def _media_parts(slide, sh):
    """解析媒体形状引用到的内部媒体部件，返回 [(部件名, 数据摘要, 字节数)]。"""
    out = []
    links, embeds = _media_rids(sh)
    for rid in list(dict.fromkeys(links + embeds)):
        try:
            rel = slide.part.rels[rid]
        except Exception:
            continue
        try:
            if rel.is_external:
                continue
        except Exception:
            pass
        try:
            part = rel.target_part
            blob = part.blob or b""
        except Exception:
            continue
        out.append((str(part.partname),
                    hashlib.sha1(blob).hexdigest(), len(blob)))
    return out


def _has_poster(sh):
    """是否带首帧预览画面：形状内含图片填充（p:blipFill / a:blipFill）
    且能解析到图片关系。纯黑框或空白框不会写入 blip 引用。"""
    x = _xml(sh)
    if "blipFill" not in x:
        return False
    return bool(re.search(r'<a:blip[^>]+r:(?:embed|link)="[^"]+"', x))

# ------------------------------------------------------------------- 维度一


def _dim1_check(prs):
    """维度一：可用与可修改性。
    ① 交付文件为 .pptx 且能被正常打开（能构造 Presentation 即视为未损坏）；
    ② 总页数为 16 页；
    ③ 不存在铺满整页的位图（宽>33 厘米且高>19 厘米的图片对象），
       即页面内容仍为可编辑对象而非整页贴图。
    维度一未通过时维度二完成度总分记 0，且不再返回逐条明细。"""
    reasons = []
    slides = list(prs.slides)
    if len(slides) != 16:
        reasons.append("总页数为%d页，不为16页" % len(slides))
    for idx, s in enumerate(slides, 1):
        for sh in s.shapes:
            if not _is_pic(sh):
                continue
            w, h = _wh(sh)
            if w > 33.0 and h > 19.0:
                reasons.append("第%d页存在铺满整页的位图（%.2f×%.2f厘米）"
                               % (idx, w, h))
                break
    return (not reasons), "；".join(reasons)


# ------------------------------------------------------------------- 主入口

# 维度一只接受 .pptx；目录里若只有下列格式，属于"交付物格式不被接受"，
# 交给维度一记 0，而不是当作"这份作业没评上"。
_EXT_REJECT = (".ppt", ".pps", ".ppsx", ".odp", ".dps", ".dpt", ".pdf", ".key")


def evaluate(dir_path):
    # 约定：status 为 "error" 表示"这份作业根本没评上"（目录不存在、目录里没有
    # 待评文件、脚本自身异常），此时 dim1_pass 为空串（未判定）、dim1_reason 为
    # 空串，dim2_items 为空列表，total_score 与 max_score 均为 0。
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
            result["status"] = "error"
            result["error"] = "目录不存在：%s" % dir_path
            return result
        cands = sorted(f for f in os.listdir(dir_path)
                       if f.lower().endswith(".pptx") and not f.startswith("~"))
        if not cands:
            rejects = sorted(f for f in os.listdir(dir_path)
                             if f.lower().endswith(_EXT_REJECT)
                             and not f.startswith("~"))
            if rejects:
                # 目录里有交付物但格式不被接受：已有交付动作，交给维度一记 0
                result["file_name"] = rejects[0]
                return _dim1_fail(result, "交付物格式为 %s，维度一仅接受 .pptx"
                                          % os.path.splitext(rejects[0])[1].lower())
            result["status"] = "error"
            result["error"] = "目录中未找到 .pptx 文件：%s" % dir_path
            return result
        file_path = os.path.join(dir_path, cands[0])
        result["file_name"] = os.path.basename(file_path)
        return _evaluate_impl(file_path, result)
    except Exception as exc:
        # 脚本异常同样视为"没评上"：维度一未判定，两项分数归零。
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


def _evaluate_impl(file_path, result):
    try:
        prs = Presentation(file_path)
    except Exception as exc:
        # 文件在但打不开（损坏或只是改了扩展名）：属于维度一不通过，
        # 不是"这份作业没评上"，所以走 _dim1_fail 而非 status="error"。
        return _dim1_fail(result, "演示文稿无法打开：%s: %s"
                                  % (type(exc).__name__, exc))
    slides = list(prs.slides)

    dim1_pass, dim1_reason = _dim1_check(prs)
    if not dim1_pass:
        return _dim1_fail(result, dim1_reason)
    result["dim1_pass"] = True
    # 约定：维度一通过时 dim1_reason 保持空串，只有未通过才写错误摘要。
    result["dim1_reason"] = ""

    results = []          # 形如 ("+5", "规则说明", "实测明细")

    def add(hit, delta, rule, detail=""):
        """delta 为该项满分：加分项传正数、减分项传负数。
        hit 表示"该项成立"——加分项为达成、减分项为违规命中。"""
        sign = "%+d" % delta if hit else ("+0" if delta > 0 else "-0")
        results.append((sign, rule, detail))

    # 页面尺寸（多项判定共用；不足 16 页时用默认页面尺寸兜底避免越界）
    pw = _cm(prs.slide_width)
    ph = _cm(prs.slide_height)
    body = slides[1:15]                     # 第 2 页至第 15 页
    s16 = slides[15] if len(slides) >= 16 else None
    s1 = slides[0] if slides else None

    # ============================================================ 加分项 01
    # +5：14 个页面顶部均含不少于 6 个形状组成的装饰图形，且构成与摆放位置一致
    # ①每页顶部装饰形状（不含标题/英文/数字文本框）数量不少于 6；
    # ②各页装饰形状按坐标排序后一一对应，几何类型相同；
    # ③对应形状左边界、上边界偏差均不超过 0.2 厘米。
    sigs = []
    counts = []
    ok01 = len(body) == 14
    for s in body:
        d = _decor(s)
        counts.append(len(d))
        if len(d) < 6:
            ok01 = False
        sig = sorted((round(_cm(sh.left), 2), round(_cm(sh.top), 2),
                      _prst(sh) or "none") for sh in d)
        sigs.append(sig)
    diff01 = ""
    if ok01 and sigs:
        ref = sigs[0]
        for i, sig in enumerate(sigs[1:], 3):
            if len(sig) != len(ref):
                ok01 = False
                diff01 = "第%d页装饰形状数量%d≠第2页%d" % (i, len(sig), len(ref))
                break
            for (l1, t1, g1), (l2, t2, g2) in zip(ref, sig):
                if g1 != g2 or abs(l1 - l2) > 0.2 or abs(t1 - t2) > 0.2:
                    ok01 = False
                    diff01 = "第%d页存在偏差超 0.2 厘米的对应形状" % i
                    break
            if not ok01:
                break
    add(ok01, 5,
        "第2页至第15页共14个页面的顶部区域（距页面上边界2.6厘米以内）均含有一组由不少于6个形状对象组成的装饰图形，"
        "且14个页面中该组装饰形状（不含承载标题与英文文字的文本框）的构成与摆放位置一致，"
        "各页对应形状的左边界、上边界坐标偏差不超过0.2厘米",
        diff01 or ("各页顶部装饰形状 %d–%d 个" % (min(counts or [0]), max(counts or [0]))))

    # ============================================================ 加分项 02
    # +5：每页顶部均含一个几何类型为 parallelogram 的对象，内部文字为两位数字
    miss02 = [i for i, s in enumerate(body, 2) if not _num_para(s)]
    ok02 = (len(body) == 14) and not miss02
    add(ok02, 5,
        "第2页至第15页每页顶部区域均含有一个平行四边形（几何类型 parallelogram）对象，其内部文字为两位数字",
        "全部14页均命中" if ok02 else "缺失页：%s" % (miss02 or "-"))

    # ============================================================ 加分项 03
    # +1：数字块蓝色系填充（纯色或渐变），块内数字为白色且字号不小于 24 磅
    ok03 = ok02
    diff03 = ""
    for i, s in enumerate(body, 2):
        ps = _num_para(s)
        if not ps:
            ok03 = False
            diff03 = "第%d页无平行四边形数字块" % i
            break
        sh = ps[0]
        if not _shape_blue(sh):
            ok03 = False
            diff03 = "第%d页数字块填充非蓝色系" % i
            break
        rs = _sig_runs(sh)
        for r in rs:
            sz = _run_size(r)
            if not _is_white(_run_color(r)):
                ok03 = False
                diff03 = "第%d页数字颜色%s非白色" % (i, _run_color(r) or "-")
            elif sz is None or sz < 24:
                ok03 = False
                diff03 = "第%d页数字字号%s小于24磅" % (i, sz)
            if not ok03:
                break
        if not ok03:
            break
    add(ok03, 1,
        "第2页至第15页顶部的平行四边形数字块为蓝色系填充（纯色或渐变，色值近似 #2962FF、#347BDC、#4472C4），"
        "块内数字为白色且字号不小于24磅",
        diff03 or "14页数字块均为蓝色系填充、白色数字且字号达标")

    # ============================================================ 加分项 04
    # +3：数字块内数字与页码对应递增，第 2 页 "02" ... 第 15 页 "15"
    seq = []
    for i, s in enumerate(body, 2):
        ps = _num_para(s)
        seq.append(_txt(ps[0]) if ps else "-")
    exp_seq = ["%02d" % i for i in range(2, 16)]
    ok04 = (seq == exp_seq)
    add(ok04, 3,
        "第2页至第15页顶部平行四边形数字块内的数字与页码对应递增：第2页为\"02\"、第3页为\"03\"，依次到第15页为\"15\"",
        "实测序列 %s" % ",".join(seq))

    # ============================================================ 加分项 05
    # +3：数字块右侧衔接一条横向长条形浅色底板或蓝色细线框
    # ①长度不小于 3 厘米、宽高比不小于 3:1（线条高度按 0.01 厘米计）；
    # ②浅色填充或蓝色系线条，显式线宽不超过 3.5 磅；
    # ③左边界不超过数字块右边界，右边界距页面左边界不超过 16.9 厘米。
    ok05 = len(body) == 14
    diff05 = ""
    for i, s in enumerate(body, 2):
        nr = _num_right(s)
        found = None
        for sh in _decor(s):
            left, top, right, bottom = _box(sh)
            w = right - left
            h = bottom - top
            if w < 3.0 or w < 3.0 * max(h, 0.01):
                continue
            if not (_shape_light(sh) or _shape_blue(sh)):
                continue
            lw = _line_w_pt(sh)
            if lw is not None and lw > 3.5:
                continue
            if left > nr + 0.01 or right > HALF_W_CM:
                continue
            found = (left, right, w, h)
            break
        if found is None:
            ok05 = False
            diff05 = "第%d页未找到合规横条（数字块右边界%.2f厘米）" % (i, nr)
            break
    add(ok05, 3,
        "第2页至第15页每页数字块右侧衔接一条横向长条形浅色填充底板或蓝色细线框（线宽不超过3.5磅），"
        "该底板或线框长度不小于3厘米、宽高比不小于3:1，其左边界不超过数字块右边界、"
        "右边界距页面左边界不超过16.9厘米，与数字块拼接成一条紧凑横幅而非横贯整页的通栏",
        diff05 or "14页均命中紧凑横条且右边界在16.9厘米以内")

    # ============================================================ 加分项 06
    # +3：装饰整体外接矩形宽高比在 2.5:1 至 7:1 之间，且宽度不超过 16.9 厘米
    ok06 = len(body) == 14
    diff06 = ""
    for i, s in enumerate(body, 2):
        bb = _decor_bbox(s)
        if bb is None:
            ok06 = False
            diff06 = "第%d页无装饰形状" % i
            break
        w = bb[2] - bb[0]
        h = bb[3] - bb[1]
        ratio = w / h if h > 0 else 999.0
        if not (2.5 <= ratio <= 7.0) or w > HALF_W_CM:
            ok06 = False
            diff06 = "第%d页外接矩形%.2f×%.2f厘米、宽高比%.2f:1" % (i, w, h, ratio)
            break
    add(ok06, 3,
        "第2页至第15页顶部装饰图形整体的外接矩形（不计宽高均为0厘米的空形状）宽高比在2.5:1至7:1之间，"
        "且宽度不超过16.9厘米（不超过页面宽度的一半），呈紧凑横幅形态而非横贯整页的细长条",
        diff06 or "14页外接矩形均为紧凑横幅")

    # ============================================================ 加分项 07
    # +3：装饰整体外接矩形左边界不超过 2 厘米、下边界不超过 2.6 厘米
    ok07 = len(body) == 14
    diff07 = ""
    for i, s in enumerate(body, 2):
        bb = _decor_bbox(s)
        if bb is None or bb[0] > 2.0 or bb[3] > TOP_BAND_CM:
            ok07 = False
            diff07 = ("第%d页无装饰形状" % i) if bb is None else \
                ("第%d页左边界%.2f、下边界%.2f厘米" % (i, bb[0], bb[3]))
            break
    add(ok07, 3,
        "第2页至第15页顶部装饰图形整体的外接矩形左边界距页面左边界不超过2厘米、"
        "下边界距页面上边界不超过2.6厘米，整体位于页面左上角",
        diff07 or "14页装饰整体均位于页面左上角")

    # ============================================================ 加分项 08
    # +3：中文标题文本框左边界在 1 至 5 厘米之间，紧邻数字块右侧且落在横幅高度内
    ok08 = len(body) == 14
    diff08 = ""
    for i, s in enumerate(body, 2):
        sh = _title_shape(s, TITLES[i - 2])
        if sh is None:
            ok08 = False
            diff08 = "第%d页未找到中文标题文本框" % i
            break
        left, top, _, bottom = _box(sh)
        nr = _num_right(s)
        bb = _decor_bbox(s)
        if not (1.0 <= left <= 5.0):
            ok08 = False
            diff08 = "第%d页标题左边界%.2f厘米不在1-5厘米之间" % (i, left)
            break
        if left > nr + 2.0:
            ok08 = False
            diff08 = "第%d页标题左边界%.2f厘米未紧邻数字块（右边界%.2f）" % (i, left, nr)
            break
        if bb is not None and not (top >= bb[1] - 0.5 and bottom <= bb[3] + 0.5):
            ok08 = False
            diff08 = "第%d页标题垂直范围%.2f-%.2f超出横幅%.2f-%.2f" % (
                i, top, bottom, bb[1], bb[3])
            break
    add(ok08, 3,
        "第2页至第15页的中文标题文本框左边界距页面左边界在1至5厘米之间，"
        "紧邻数字块右侧且在垂直方向落在顶部横幅高度范围内，未处于页面中部或右侧",
        diff08 or "14页标题均紧邻数字块右侧且落在横幅高度内")

    # ============================================================ 加分项 09
    # +3：中文标题文字与页码对应（须位于顶部区域，避免命中正文里的同名文字）
    miss09 = []
    for i, s in enumerate(body, 2):
        sh = _title_shape(s, TITLES[i - 2])
        if sh is None or _cm(sh.top) > TOP_BAND_CM:
            miss09.append(i)
    ok09 = (len(body) == 14) and not miss09
    add(ok09, 3,
        "第2页至第15页的中文标题文字与页码对应，依次为\"内容框架\"\"品牌定位\"\"消费场景\"\"产品矩阵\""
        "\"数字化运营\"\"会员运营\"\"门店模型\"\"供应链协同\"\"品质控制\"\"增长路径\"\"营销策略\""
        "\"效率提升\"\"风险管理\"\"总结展望\"",
        "14页标题文字全部对应" if ok09 else "不匹配页：%s" % (miss09 or "-"))

    # ============================================================ 加分项 10
    # +1：中文标题字号不小于 30 磅且为加粗
    ok10 = len(body) == 14
    diff10 = ""
    for i, s in enumerate(body, 2):
        sh = _title_shape(s, TITLES[i - 2])
        rs = _sig_runs(sh) if sh is not None else []
        if not rs:
            ok10 = False
            diff10 = "第%d页未取到标题文本" % i
            break
        for r in rs:
            sz = _run_size(r)
            if sz is None or sz < 30 or not _run_bold(r):
                ok10 = False
                diff10 = "第%d页标题字号%s、加粗%s" % (i, sz, _run_bold(r))
                break
        if not ok10:
            break
    add(ok10, 1,
        "第2页至第15页的中文标题字号不小于30磅且为加粗",
        diff10 or "14页标题均不小于30磅且加粗")

    # ============================================================ 加分项 11
    # +3：中文标题下方或右侧含一行英文文字，字号 10-13 磅，颜色蓝色系
    ok11 = len(body) == 14
    diff11 = ""
    en_texts = []
    for i, s in enumerate(body, 2):
        en = _en_shape(s)
        en_texts.append(_txt(en) if en is not None else "")
        if en is None:
            ok11 = False
            diff11 = "第%d页未找到英文文字对象" % i
            break
        title = _title_shape(s, TITLES[i - 2])
        el, et, _, _ = _box(en)
        if title is not None:
            tl, tt, tr, _ = _box(title)
            if not (et >= tt - 0.1 or el >= tr - 0.5):
                ok11 = False
                diff11 = "第%d页英文既不在标题下方也不在右侧" % i
                break
        for r in _sig_runs(en):
            sz = _run_size(r)
            c = _run_color(r)
            if sz is None or not (10.0 <= sz <= 13.0):
                ok11 = False
                diff11 = "第%d页英文字号%s不在10-13磅之间" % (i, sz)
            elif not _is_blue_color(c):
                ok11 = False
                diff11 = "第%d页英文颜色%s非蓝色系" % (i, c or "-")
            if not ok11:
                break
        if not ok11:
            break
    add(ok11, 3,
        "第2页至第15页每页在中文标题下方或右侧含有一行英文文字，字号在10至13磅之间，颜色为蓝色系",
        diff11 or "14页英文行均为10-13磅蓝色系")

    # ============================================================ 加分项 12
    # +3：英文文字为该页中文标题对应的英文译名，且 14 页互不相同
    ok12 = (len(en_texts) == 14) and all(en_texts)
    diff12 = ""
    if ok12 and len(set(t.lower() for t in en_texts)) != 14:
        ok12 = False
        diff12 = "14页英文内容存在重复（去重后%d种）" % len(
            set(t.lower() for t in en_texts))
    if ok12:
        for i, t in enumerate(en_texts, 2):
            kws = EN_KEYWORDS[TITLES[i - 2]]
            low = t.lower()
            if len(t) > 48 or not any(k in low for k in kws):
                ok12 = False
                diff12 = "第%d页英文\"%s\"非\"%s\"的译名" % (i, t, TITLES[i - 2])
                break
    add(ok12, 3,
        "第2页至第15页的英文文字分别为该页中文标题对应的英文译名（如\"品牌定位\"页为 brand positioning、"
        "\"门店模型\"页为 Store Model、\"增长路径\"页为 Growth Path），14个页面的英文内容互不相同",
        diff12 or "14页英文均为对应译名且互不相同")

    # ============================================================ 加分项 13
    # +3：横幅右端不少于 9 个尺寸不超过 0.3 厘米的小点，呈不少于 3 行 3 列点阵
    # ①小点为装饰形状中宽高均大于 0 且不超过 0.3 厘米者，右边界不超过 16.9 厘米；
    # ②按 0.1 厘米容差聚簇后，至少 3 个"含 3 个以上小点"的列簇与行簇；
    # ③点阵重心落在装饰整体外接矩形的右半侧（位于横幅右端）。
    ok13 = len(body) == 14
    diff13 = ""
    for i, s in enumerate(body, 2):
        dots = []
        for sh in _decor(s):
            w, h = _wh(sh)
            left, top, right, _ = _box(sh)
            if 0 < w <= 0.3 and 0 < h <= 0.3 and right <= HALF_W_CM:
                dots.append((left, top))
        if len(dots) < 9:
            ok13 = False
            diff13 = "第%d页16.9厘米内不超过0.3厘米的小点仅%d个" % (i, len(dots))
            break
        cols = [g for g in _clusters([d[0] for d in dots]) if len(g) >= 3]
        rows = [g for g in _clusters([d[1] for d in dots]) if len(g) >= 3]
        if len(cols) < 3 or len(rows) < 3:
            ok13 = False
            diff13 = "第%d页点阵仅%d列%d行" % (i, len(cols), len(rows))
            break
        grid = [d for d in dots
                if any(abs(d[0] - v) <= 0.1 for g in cols for v in g)]
        bb = _decor_bbox(s)
        gx = sum(d[0] for d in grid) / float(len(grid))
        if bb is not None and gx < bb[0] + (bb[2] - bb[0]) / 2.0:
            ok13 = False
            diff13 = "第%d页点阵重心%.2f厘米未落在横幅右端" % (i, gx)
            break
    add(ok13, 3,
        "第2页至第15页顶部装饰图形中含有不少于9个尺寸不超过0.3厘米的小圆点或小方点，"
        "在横幅右端呈不少于3行3列的网格状排布，且该点阵整体的右边界距页面左边界不超过16.9厘米"
        "（位于左上角横幅右端，而非页面右侧边缘）",
        diff13 or "14页均含3行3列以上点阵且位于16.9厘米以内")

    # ============================================================ 加分项 14
    # +1：含 3 个平行排列的小斜条（平行四边形/斜纹）以及 1 个独立的小斜块，
    #     且斜条与斜块的右边界距页面左边界均不超过 16.9 厘米。
    ok14 = len(body) == 14
    diff14 = ""
    for i, s in enumerate(body, 2):
        slants = []
        for sh in _decor(s):
            if _prst(sh) not in ("parallelogram", "diagStripe"):
                continue
            w, h = _wh(sh)
            left, top, right, _ = _box(sh)
            if 0 < w <= 1.2 and 0 < h <= 1.2 and right <= HALF_W_CM:
                slants.append((left, top))
        groups = [g for g in _clusters([v[1] for v in slants]) if len(g) >= 3]
        if not groups or len(slants) < 4:
            ok14 = False
            diff14 = "第%d页16.9厘米内小斜条%d个、平行组%d组" % (
                i, len(slants), len(groups))
            break
    add(ok14, 1,
        "第2页至第15页顶部装饰图形中含有3个平行排列的小斜条（平行四边形或斜纹形状），以及1个独立的小斜块，"
        "且这些斜条与斜块的右边界距页面左边界均不超过16.9厘米",
        diff14 or "14页均含3个平行小斜条与1个独立小斜块")

    # ---------------------------------------------- 第 16 页结束页相关对象定位
    thx = []
    vids = []
    if s16 is not None:
        thx = [sh for sh in s16.shapes if _txt(sh) == "谢谢观看"]
        vids = [sh for sh in s16.shapes if _is_media(sh)]
        vids.sort(key=lambda sh: _cm(sh.left))

    # ============================================================ 加分项 15
    # +3：第 16 页含独立的"谢谢观看"文本对象（单独一个文本对象，文字即"谢谢观看"）
    ok15 = len(thx) >= 1
    add(ok15, 3, "第16页含有独立的\"谢谢观看\"文本对象",
        "命中%d个\"谢谢观看\"文本对象" % len(thx) if ok15 else "第16页未找到\"谢谢观看\"文本对象")

    # ============================================================ 加分项 16
    # +3："谢谢观看"的中文字体（东亚字体）与西文字体均为宋体
    ok16 = False
    diff16 = "对象缺失"
    if ok15:
        ok16 = True
        for r in _sig_runs(thx[0]):
            la, ea = _run_latin(r), _run_ea(r)
            if not (_is_song(la) and _is_song(ea)):
                ok16 = False
                diff16 = "中文字体%s、西文字体%s" % (ea or "-", la or "-")
                break
        if ok16:
            diff16 = "中西文字体均为宋体"
    add(ok16, 3, "第16页\"谢谢观看\"的中文字体与西文字体均为宋体", diff16)

    # ============================================================ 加分项 17
    # +1："谢谢观看"为加粗
    ok17 = bool(ok15) and all(_run_bold(r) for r in _sig_runs(thx[0]))
    add(ok17, 1, "第16页\"谢谢观看\"为加粗",
        "已加粗" if ok17 else "未加粗或对象缺失")

    # ============================================================ 加分项 18
    # +1："谢谢观看"为蓝色系（色值近似 #2360B9、#4472C4，或引用蓝色主题色）
    ok18 = False
    diff18 = "对象缺失"
    if ok15:
        cs = [_run_color(r) for r in _sig_runs(thx[0])]
        ok18 = bool(cs) and all(
            _is_near_hex(c, ("2360B9", "4472C4")) or
            (c or "").startswith("scheme:accent") for c in cs)
        diff18 = "文字颜色 %s" % ",".join(c or "-" for c in cs)
    add(ok18, 1, "第16页\"谢谢观看\"为蓝色系（色值近似 #2360B9、#4472C4）", diff18)

    # ============================================================ 加分项 19
    # +3："谢谢观看"所在段落居中对齐，且文本框水平中心与页面水平中心偏差不超过 0.5 厘米
    ok19 = False
    diff19 = "对象缺失"
    if ok15:
        sh = thx[0]
        left, _, right, _ = _box(sh)
        cx = (left + right) / 2.0
        al = _para_align(sh, "谢谢观看")
        ok19 = (al == PP_ALIGN.CENTER) and abs(cx - CENTER_X_CM) <= 0.5
        diff19 = "段落对齐%s、文本框水平中心%.2f厘米" % (
            getattr(al, "name", al), cx)
    add(ok19, 3,
        "第16页\"谢谢观看\"所在段落为居中对齐，且其文本框水平中心与页面水平中心"
        "（距页面左边界16.93厘米处）偏差不超过0.5厘米", diff19)

    # ============================================================ 加分项 20
    # +5：第 16 页含两个嵌入到文件内部的可播放视频媒体对象
    # ①媒体对象数量为 2；②各自含 videoFile 视频关系与 p14:media 嵌入关系；
    # ③嵌入关系能解析到包内非空媒体部件（真正的视频数据，而非静态图片）。
    parts20 = []
    ok20 = (len(vids) == 2)
    diff20 = "媒体对象数量%d" % len(vids)
    if ok20:
        for sh in vids:
            links, embeds = _media_rids(sh)
            mp = _media_parts(s16, sh)
            parts20.append(mp)
            if not links or not embeds or not any(p[2] > 0 for p in mp):
                ok20 = False
        if ok20:
            diff20 = "2个媒体对象，嵌入媒体数据 %s" % "、".join(
                "%d字节" % max(p[2] for p in mp) for mp in parts20)
    add(ok20, 5,
        "第16页含有两个嵌入到文件内部的可播放视频媒体对象（含视频关系与嵌入的媒体数据），而非静态图片",
        diff20)

    # ============================================================ 加分项 21
    # +3：两个视频对象指向两个不同的媒体文件（部件名不同且数据摘要不同）
    ok21 = False
    diff21 = "媒体对象不足2个"
    if len(parts20) == 2 and all(parts20):
        n1 = set(p[0] for p in parts20[0])
        n2 = set(p[0] for p in parts20[1])
        d1 = set(p[1] for p in parts20[0])
        d2 = set(p[1] for p in parts20[1])
        ok21 = bool(n1) and bool(n2) and not (n1 & n2) and not (d1 & d2)
        diff21 = "引用媒体部件 %s / %s" % (
            ",".join(os.path.basename(x) for x in sorted(n1)) or "-",
            ",".join(os.path.basename(x) for x in sorted(n2)) or "-")
    add(ok21, 3, "第16页两个视频对象指向两个不同的媒体文件，未重复引用同一段媒体数据", diff21)

    # ============================================================ 加分项 22
    # +5：两个视频左右并列排布
    # ①"谢谢观看"文本框顶端高于两个视频顶端；②两视频顶端高度差不超过 0.5 厘米；
    # ③水平投影互不重叠；④均完整位于页面内。
    ok22 = False
    diff22 = "媒体对象不足2个或文本对象缺失"
    if len(vids) == 2 and ok15:
        b1 = _box(vids[0])
        b2 = _box(vids[1])
        bt = _box(thx[0])
        cond_a = bt[1] < min(b1[1], b2[1])
        cond_b = abs(b1[1] - b2[1]) <= 0.5
        cond_c = (b1[2] <= b2[0] + 0.001) or (b2[2] <= b1[0] + 0.001)
        cond_d = all(b[0] >= -0.001 and b[2] <= PAGE_W_CM + 0.001
                     and b[3] <= PAGE_H_CM + 0.001 for b in (b1, b2))
        ok22 = cond_a and cond_b and cond_c and cond_d
        diff22 = ("文本顶端%.2f，视频顶端%.2f/%.2f，横向%.2f-%.2f 与 %.2f-%.2f，"
                  "下边界%.2f/%.2f" % (bt[1], b1[1], b2[1], b1[0], b1[2],
                                    b2[0], b2[2], b1[3], b2[3]))
    add(ok22, 5,
        "第16页两个视频左右并列排布：\"谢谢观看\"文本框顶端高于两个视频顶端，两个视频顶端高度差不超过0.5厘米，"
        "水平投影互不重叠，且均完整位于页面内（左边界不小于0厘米、右边界不超过33.86厘米、下边界不超过19.05厘米）",
        diff22)

    # ============================================================ 加分项 23
    # +1：两个视频的宽度差与高度差均不超过 0.5 厘米
    ok23 = False
    diff23 = "媒体对象不足2个"
    if len(vids) == 2:
        w1, h1 = _wh(vids[0])
        w2, h2 = _wh(vids[1])
        ok23 = abs(w1 - w2) <= 0.5 and abs(h1 - h2) <= 0.5
        diff23 = "尺寸 %.2f×%.2f 与 %.2f×%.2f 厘米" % (w1, h1, w2, h2)
    add(ok23, 1, "第16页两个视频的宽度差与高度差均不超过0.5厘米", diff23)

    # ============================================================ 加分项 24
    # +1：两个视频的宽高比均在 0.71 至 0.79 之间（3:4 竖版，画面未横向变形）
    ok24 = False
    diff24 = "媒体对象不足2个"
    if len(vids) == 2:
        rs = []
        for sh in vids:
            w, h = _wh(sh)
            rs.append(w / h if h > 0 else 0.0)
        ok24 = all(0.71 <= r <= 0.79 for r in rs)
        diff24 = "宽高比 %.3f / %.3f" % (rs[0], rs[1])
    add(ok24, 1, "第16页两个视频的宽高比均在0.71至0.79之间（3:4竖版比例），画面未横向变形", diff24)

    # ============================================================ 加分项 25
    # +1：两个视频均带首帧预览画面（含预览图填充），不显示为纯黑框或空白框
    ok25 = (len(vids) == 2) and all(_has_poster(sh) for sh in vids)
    add(ok25, 1, "第16页两个视频均带首帧预览画面（含预览图填充），不显示为纯黑框或空白框",
        "两个视频均含预览图填充" if ok25 else "存在未设置预览图填充的视频对象")

    # ============================================================ 减分项 26
    # -5：顶部标题栏的装饰形状位于页面已有元素之上层并与其重叠，任一页重叠面积
    #     超过 1 平方厘米。层级按 spTree 顺序判定（下标越大越靠上层）。
    bad26 = []
    for i, s in enumerate(body, 2):
        shapes = list(s.shapes)
        order = {id(sh._element): k for k, sh in enumerate(shapes)}
        band = _band_shapes(s)
        band_ids = set(id(sh._element) for sh in band)
        targets = []
        for sh in shapes:
            if id(sh._element) in band_ids:
                continue
            left, top, _, _ = _box(sh)
            if left > 100 or top > 100:
                continue
            if _is_pic(sh) or _has_chart(sh) or _txt(sh):
                targets.append(sh)
        worst = 0.0
        for bsh in band:
            for tsh in targets:
                if order[id(bsh._element)] <= order[id(tsh._element)]:
                    continue
                a = _inter_area(_box(bsh), _box(tsh))
                if a > worst:
                    worst = a
        if worst > 1.0:
            bad26.append((i, round(worst, 2)))
    hit26 = bool(bad26)
    add(hit26, -5,
        "第2页至第15页中，顶部标题栏的装饰形状（浅色底板、线框、文字框等）位于页面已有元素"
        "（照片、图表、文字块等）之上层并与其重叠，任一页重叠面积超过1平方厘米",
        ("命中页/最大重叠面积：%s" % "，".join("第%d页%.2f平方厘米" % b for b in bad26))
        if hit26 else "各页标题栏均未压盖页面已有元素")

    # ============================================================ 减分项 27
    # -1：第 2 页含宽高均为 0 厘米、左边界与上边界均大于 5000 厘米的页外空白形状
    ghost = []
    if len(slides) >= 2:
        for sh in slides[1].shapes:
            w, h = _wh(sh)
            left, top, _, _ = _box(sh)
            if w == 0 and h == 0 and left > 5000 and top > 5000:
                ghost.append((left, top))
    hit27 = bool(ghost)
    add(hit27, -1,
        "第2页含有一个宽度与高度均为0厘米、左边界与上边界均大于5000厘米、位于页面之外的空白形状对象",
        ("命中%d个（左%.0f/上%.0f厘米）" % (len(ghost), ghost[0][0], ghost[0][1]))
        if hit27 else "未发现页外空白形状")

    # ============================================================ 减分项 28
    # -1：第 14 页含两处"风险管理"文字，且其中一处文本框右边界超出页面右边界
    hit28 = False
    diff28 = "第14页缺失"
    if len(slides) >= 14:
        hits = [sh for sh in slides[13].shapes if "风险管理" in _txt(sh)]
        out = [sh for sh in hits if _box(sh)[2] > PAGE_W_CM + 0.01]
        hit28 = len(hits) >= 2 and bool(out)
        diff28 = "\"风险管理\"文字%d处，超出页面右边界%d处" % (len(hits), len(out))
    add(hit28, -1,
        "第14页含有两处\"风险管理\"文字，且其中一处文本框右边界超出页面右边界（大于33.86厘米）",
        diff28)

    # ============================================================ 减分项 29
    # -1：第 15 页正文段落须距页面左边界 1.98 厘米、距页面上边界 5.97 厘米
    hit29 = True
    diff29 = "第15页未找到该正文段落"
    if len(slides) >= 15:
        for sh in slides[14].shapes:
            if P15_TEXT in _txt(sh):
                left, top, _, _ = _box(sh)
                diff29 = "实测距左%.2f、距上%.2f厘米" % (left, top)
                if abs(left - 1.98) <= 0.05 and abs(top - 5.97) <= 0.05:
                    hit29 = False
                    break
    add(hit29, -1,
        "第15页正文段落\"晨岚咖啡的增长关键，不是追逐单一爆品或短期声量，而是把稳定出品、便利触达和轻量会员关系长期做细。\""
        "不满足距页面左边界1.98厘米、距页面上边界5.97厘米", diff29)

    # ============================================================ 减分项 30
    # -1：第 16 页含包围"谢谢观看"文字或包围视频画面的圆角矩形边框对象
    hit30 = False
    diff30 = "未发现圆角矩形边框对象"
    if s16 is not None:
        frames = [sh for sh in s16.shapes
                  if _prst(sh) in ("roundRect", "round1Rect", "round2SameRect")
                  and not _txt(sh)]
        tgts = [_box(sh) for sh in thx] + [_box(sh) for sh in vids]
        worst = 0.0
        for f in frames:
            for tb in tgts:
                a = _inter_area(_box(f), tb)
                cov = a / _area(tb) if _area(tb) > 0 else 0.0
                if cov > worst:
                    worst = cov
        hit30 = worst >= 0.5
        if frames:
            diff30 = "圆角矩形对象%d个，最大包围比例%.0f%%" % (len(frames), worst * 100)
    add(hit30, -1,
        "第16页含有包围\"谢谢观看\"文字的圆角矩形边框对象，或含有包围视频画面的圆角矩形边框对象",
        diff30)

    # ============================================================ 减分项 31
    # -1：第 16 页含页码文字且其内容不为"16"（如与第 15 页重复显示为"15"）
    pn16 = []
    if s16 is not None:
        pn16 = [_txt(sh) for sh in s16.shapes
                if re.fullmatch(r"\d{1,2}", _txt(sh))]
    hit31 = bool(pn16) and all(t != "16" for t in pn16)
    add(hit31, -1, "第16页含有页码文字且其内容不为\"16\"（如与第15页重复显示为\"15\"）",
        "页码文字：%s" % (",".join(pn16) if pn16 else "无页码文字"))

    # ============================================================ 减分项 32
    # -3：第 1 页主标题格式与位置
    t32 = None
    if s1 is not None:
        cands = [sh for sh in s1.shapes if _txt(sh) == DOC_TITLE]
        t32 = cands[0] if cands else None
    ok32, diff32 = _fmt_ok(t32, font=BODY_FONT, size=40, bold=True,
                           color="143A4A", left=1.83, top=5.84)
    add(not ok32, -3,
        "第1页主标题\"晨岚咖啡品牌运营研究\"不满足字体 Noto Sans CJK SC、40磅、加粗、颜色 143A4A，"
        "距页面左边界1.83厘米、距页面上边界5.84厘米",
        diff32 or "字体、字号、加粗、颜色与位置均满足")

    # ============================================================ 减分项 33
    # -1：第 1 页副标题格式与位置
    t33 = None
    if s1 is not None:
        cands = [sh for sh in s1.shapes if _txt(sh) == SUB_TITLE]
        t33 = cands[0] if cands else None
    ok33, diff33 = _fmt_ok(t33, font=BODY_FONT, size=20,
                           color="123B63", left=1.91, top=8.08)
    add(not ok33, -1,
        "第1页副标题\"城市轻饮场景的产品、渠道与运营重构\"不满足字体 Noto Sans CJK SC、20磅、颜色 123B63，"
        "距页面左边界1.91厘米、距页面上边界8.08厘米",
        diff33 or "字体、字号、颜色与位置均满足")

    # ============================================================ 减分项 34
    # -1：第 1 页右侧照片位置与尺寸（18.67/0/15.19×19.05）
    hit34 = True
    diff34 = "第1页未找到照片对象"
    if s1 is not None:
        pics = [sh for sh in s1.shapes
                if _is_pic(sh) and _wh(sh)[0] >= 5 and _wh(sh)[1] >= 5]
        for sh in pics:
            left, top, _, _ = _box(sh)
            w, h = _wh(sh)
            diff34 = "实测 %.2f/%.2f/%.2f×%.2f" % (left, top, w, h)
            if (abs(left - 18.67) <= 0.06 and abs(top - 0.0) <= 0.06
                    and abs(w - 15.19) <= 0.06 and abs(h - 19.05) <= 0.06):
                hit34 = False
                break
    add(hit34, -1,
        "第1页右侧照片不满足距页面左边界18.67厘米、距页面上边界0厘米、宽15.19厘米、高19.05厘米",
        diff34)

    # ============================================================ 减分项 35
    # -3：第 2 页至第 15 页左下角页脚格式与位置
    hit35 = False
    diff35 = "14页页脚字体、字号、颜色与位置均满足"
    for i, s in enumerate(body, 2):
        cands = [sh for sh in s.shapes
                 if _txt(sh) == DOC_TITLE and _cm(sh.top) > 15.0]
        sh = cands[0] if cands else None
        ok, d = _fmt_ok(sh, font=BODY_FONT, size=8.5, color="7F99A8",
                        left=1.40, top=17.91, size_tol=0.3)
        if not ok:
            hit35 = True
            diff35 = "第%d页：%s" % (i, d)
            break
    add(hit35, -3,
        "第2页至第15页左下角页脚\"晨岚咖啡品牌运营研究\"不满足字体 Noto Sans CJK SC、8.5磅、颜色 7F99A8，"
        "距页面左边界1.40厘米、距页面上边界17.91厘米", diff35)

    # ============================================================ 减分项 36
    # -1：第 2 页至第 15 页右下角两位页码格式、右对齐与位置
    # 页码与顶部数字块文字相同，按"上边界大于 15 厘米"限定在页面右下角取值。
    hit36 = False
    diff36 = "14页页码字体、字号、右对齐、颜色与位置均满足"
    for i, s in enumerate(body, 2):
        cands = [sh for sh in s.shapes
                 if re.fullmatch(r"\d{2}", _txt(sh)) and _cm(sh.top) > 15.0]
        sh = cands[0] if cands else None
        ok, d = _fmt_ok(sh, font=BODY_FONT, size=9, color="7F99A8",
                        left=31.11, top=17.83)
        al = _para_align(sh) if sh is not None else None
        if not ok or al != PP_ALIGN.RIGHT:
            hit36 = True
            diff36 = "第%d页：%s" % (i, d or ("对齐方式%s≠RIGHT"
                                            % getattr(al, "name", al)))
            break
    add(hit36, -1,
        "第2页至第15页右下角两位页码不满足字体 Noto Sans CJK SC、9磅、右对齐、颜色 7F99A8，"
        "距页面左边界31.11厘米、距页面上边界17.83厘米", diff36)

    # ============================================================ 减分项 37/38
    # -3：导语段落须为 Noto Sans CJK SC、20磅、加粗、颜色 143A4A；
    # -3：条目正文段落须为 Noto Sans CJK SC、20磅、颜色 5E7280。
    # 正文区段落按"是否加粗"区分导语与条目：加粗者为导语，非加粗者为条目；
    # 段落末尾单个标点字符不参与比对（见 _sig_runs 说明）。
    hit37 = False
    hit38 = False
    diff37 = "14页导语段落格式均满足"
    diff38 = "14页条目正文段落格式均满足"
    for i, s in enumerate(body, 2):
        leads = []
        items = []
        for sh, t, rs in _body_paras(s):
            if all(_run_bold(r) for r in rs):
                leads.append((t, rs))
            else:
                items.append((t, rs))
        if not hit37:
            if not leads:
                hit37 = True
                diff37 = "第%d页未识别到导语段落" % i
            else:
                for t, rs in leads:
                    d = _para_fmt_bad(rs, BODY_FONT, 20, "143A4A", bold=True)
                    if d:
                        hit37 = True
                        diff37 = "第%d页\"%s\"：%s" % (i, t[:12], d)
                        break
        if not hit38:
            if not items:
                hit38 = True
                diff38 = "第%d页未识别到条目正文段落" % i
            else:
                for t, rs in items:
                    d = _para_fmt_bad(rs, BODY_FONT, 20, "5E7280")
                    if d:
                        hit38 = True
                        diff38 = "第%d页\"%s\"：%s" % (i, t[:12], d)
                        break
        if hit37 and hit38:
            break
    add(hit37, -3,
        "第2页至第15页的导语段落（如第3页\"晨岚咖啡定位为\"高频轻饮补给站\"，面向通勤、办公、学习和轻社交人群。\"）"
        "不满足字体 Noto Sans CJK SC、20磅、加粗、颜色 143A4A（不计段落末尾单个标点字符）", diff37)
    add(hit38, -3,
        "第2页至第15页的条目正文段落（如第3页\"价格表达采用区间化策略，减少单点价格依赖。\"）"
        "不满足字体 Noto Sans CJK SC、20磅、颜色 5E7280", diff38)

    # ============================================================ 减分项 39
    # -3：第 4、6、7、8、12、13 页各含 1 个可编辑图表对象（未被移除或图片化）
    bad39 = []
    for n in CHART_PAGES:
        if len(slides) < n:
            bad39.append("第%d页缺失" % n)
            continue
        c = sum(1 for sh in slides[n - 1].shapes if _has_chart(sh))
        if c != 1:
            bad39.append("第%d页图表%d个" % (n, c))
    hit39 = bool(bad39)
    add(hit39, -3,
        "第4、6、7、8、12、13页不满足各含有1个可编辑的图表对象（图表被移除或以图片形式呈现）",
        "；".join(bad39) if hit39 else "6个页面均含1个可编辑图表对象")

    # ============================================================ 减分项 40
    # -3：第 2 页至第 15 页照片位置与尺寸（逐页比对，容差 0.06 厘米）
    # 顶部装饰用的小图点尺寸远小于 5 厘米，不纳入照片比对。
    bad40 = []
    for n, (el, et, ew, eh) in sorted(PHOTO_SPECS.items()):
        if len(slides) < n:
            bad40.append("第%d页缺失" % n)
            continue
        pics = [sh for sh in slides[n - 1].shapes
                if _is_pic(sh) and _wh(sh)[0] >= 5 and _wh(sh)[1] >= 5]
        ok = False
        for sh in pics:
            left, top, _, _ = _box(sh)
            w, h = _wh(sh)
            if (abs(left - el) <= 0.06 and abs(top - et) <= 0.06
                    and abs(w - ew) <= 0.06 and abs(h - eh) <= 0.06):
                ok = True
                break
        if not ok:
            got = "、".join("%.2f/%.2f/%.2f×%.2f" % (_box(p)[0], _box(p)[1],
                                                   _wh(p)[0], _wh(p)[1])
                           for p in pics) or "无照片"
            bad40.append("第%d页应为%.2f/%.2f/%.2f×%.2f，实测%s"
                         % (n, el, et, ew, eh, got))
    hit40 = bool(bad40)
    add(hit40, -3,
        "第2页至第15页的照片不满足以下位置与尺寸：第2页 20.95/2.54/11.30×13.97，第3页 1.78/2.67/14.48×13.84，"
        "第4页 18.41/2.41/13.59×14.73，第5页 1.65/2.54/13.21×14.48，第6页 2.03/2.41/13.46×14.99，"
        "第7页 18.29/2.41/13.72×15.11，第8页 18.29/2.41/13.72×15.11，第9页 1.78/2.41/15.11×14.86，"
        "第10页 17.65/2.41/14.35×15.11，第11页 18.41/2.54/13.21×14.48，第12页 1.78/2.41/14.61×14.86，"
        "第13页 1.65/2.41/13.59×15.11，第14页 18.16/2.41/13.59×15.11，第15页 18.03/0/15.82×19.05"
        "（依次为距页面左边界厘米/距页面上边界厘米/宽厘米×高厘米）",
        "；".join(bad40[:2]) if hit40 else "14页照片位置与尺寸均满足")

    # ============================================================ 减分项 41
    # -1：第 3 页"品牌关键词"色块须为填充色 FFF9EF、边线色 D5E8F6
    hit41 = True
    diff41 = "第3页未找到品牌关键词色块"
    if len(slides) >= 3:
        for sh in slides[2].shapes:
            if "品牌关键词" not in _txt(sh):
                continue
            x = _xml(sh).upper()
            has_fill = 'SRGBCLR VAL="FFF9EF"' in x
            has_line = 'SRGBCLR VAL="D5E8F6"' in x
            diff41 = "填充色FFF9EF%s、边线色D5E8F6%s" % (
                "命中" if has_fill else "缺失", "命中" if has_line else "缺失")
            if has_fill and has_line:
                hit41 = False
                break
    add(hit41, -1,
        "第3页\"品牌关键词：稳定出品｜轻松下单｜快速取杯\"色块不满足填充色 FFF9EF、边线色 D5E8F6",
        diff41)

    # ============================================================ 减分项 42
    # -1：第 9 页须含"协同重点：少压货、快流转、可追踪"文字块
    hit42 = True
    if len(slides) >= 9:
        hit42 = not any(P9_TEXT in _txt(sh) for sh in slides[8].shapes)
    add(hit42, -1, "第9页不满足含有\"协同重点：少压货、快流转、可追踪\"文字块",
        "未找到该文字块" if hit42 else "已命中该文字块")

    # ============================================================ 减分项 43
    # -1：第 2、10、14 页各条目左侧须含宽 0.13、高 0.89 厘米、颜色 333333 的小竖条
    bad43 = []
    for n in BAR_PAGES:
        if len(slides) < n:
            bad43.append("第%d页缺失" % n)
            continue
        cnt = 0
        for sh in slides[n - 1].shapes:
            w, h = _wh(sh)
            if abs(w - 0.13) <= 0.03 and abs(h - 0.89) <= 0.03 \
                    and "333333" in _xml(sh).upper():
                cnt += 1
        if cnt < 1:
            bad43.append("第%d页小竖条%d个" % (n, cnt))
    hit43 = bool(bad43)
    add(hit43, -1,
        "第2、10、14页各条目左侧不满足含有宽0.13厘米、高0.89厘米、颜色 333333 的深灰色小竖条",
        "；".join(bad43) if hit43 else "3个页面均含深灰色小竖条")

    # ============================================================ 减分项 44
    # -1：页面尺寸须为宽 33.86 厘米、高 19.05 厘米
    hit44 = not (abs(pw - PAGE_W_CM) <= 0.05 and abs(ph - PAGE_H_CM) <= 0.05)
    add(hit44, -1, "页面尺寸不满足宽33.86厘米、高19.05厘米",
        "实测页面 %.2f×%.2f 厘米" % (pw, ph))

    # ------------------------------------------------- results -> dim2_items
    if len(results) != len(_MAX_DELTAS_ORDER):
        raise RuntimeError("dim2 项数不一致：results=%d vs order=%d"
                           % (len(results), len(_MAX_DELTAS_ORDER)))
    dim2_items = []
    total_score = 0
    # detail 只用于本地排查，不写入返回信息：命中与未命中登记的规则原文完全一致，
    # 返回的 rule 不随判定结果变化。
    for (head, rule, detail), max_delta in zip(results, _MAX_DELTAS_ORDER):
        try:
            delta_raw = int(head)
        except ValueError:
            delta_raw = 0
        hit = (delta_raw != 0)
        delta = max_delta if hit else 0
        dim2_items.append({
            "rule": rule,
            "max_delta": max_delta,
            "delta": delta,
            "hit": hit,
        })
        # 加分项命中累加正分，减分项命中累加负分
        total_score += delta

    if total_score < 0:
        total_score = 0

    result["dim2_items"] = dim2_items
    result["total_score"] = total_score
    result["max_score"] = MAX_SCORE_TOTAL
    return result


if __name__ == "__main__":
    if len(sys.argv) >= 2:
        _dir = sys.argv[1]
    else:
        _dir = os.path.dirname(os.path.abspath(__file__))
    _out = evaluate(_dir)
    print(json.dumps(_out, ensure_ascii=False, indent=2))

