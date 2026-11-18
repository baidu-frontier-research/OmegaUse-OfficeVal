#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# pyright: reportAttributeAccessIssue=false, reportOptionalMemberAccess=false, reportMissingImports=false

"""
按接口统一约定，本模块不修改全局 sys.stdout，也不 print 主结果。
主结果统一由 evaluate(dir_path) -> dict 返回，入参为“脚本所在目录的路径”，
由脚本自己在该目录内定位并打开被评估的 .pptx。
本文件仅在 __main__ 下作为本地调试入口，把 dict 序列化为 JSON 打印。
"""

import os
import re
import sys
import json
import math
import shutil
import zipfile
import tempfile

from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.oxml.ns import qn

SCRIPT_ID = "063"

# 完成度分母 = 加分项合计 46 分；减分项不进入分母，只从实得分里扣。
MAX_SCORE_TOTAL = 46

# 减分项合计（21 条），仅用于自检 RULE_MAX_DELTA 写全了没有。
MAX_PENALTY_TOTAL = -59

EMU_PER_CM = 360000.0

# 统一允差 ±0.3 cm，另有注明的从其注明。
TOL = 0.3

PAGE_W_CM = 33.87
PAGE_H_CM = 19.05

YAHEI_NAMES = ("Microsoft YaHei", "微软雅黑")

# 页序：封面／岗位职责／协作关系／核心品类／工作成效／问题复盘／能力成长／团队协作／后续规划／感谢聆听
PAGE_KEYS = [
    "直播运营阶段复盘", "岗位职责", "协作关系", "核心品类", "工作成效",
    "问题复盘", "能力成长", "团队协作", "后续规划", "感谢聆听",
]

# 各页背景纯色色值；工作成效页（索引 4）由加分项单独评，此处置空。
BG_COLORS = [
    "EAF8F9", "F3FAFB", "F6FBFC", "EFF9FA", None,
    "FFFDF8", "F3FAFB", "F6FBFC", "F0FAFB", "EAF8F9",
]

# 工作成效页背景允许区间（浅青白）：EAF8F9–F6FBFC
BG_WORK_LO = (0xEA, 0xF8, 0xF9)
BG_WORK_HI = (0xF6, 0xFB, 0xFC)

TITLE_COLOR = "17323A"        # 页面标题色
LINE_TEAL = "29A7A5"          # 标题下方装饰短线色
LINE_LIGHT = "B9E1E4"         # 右下角装饰横线色

# 图表数据
CHART_CATS = ["1月", "2月", "3月", "4月", "5月"]
CHART_REVENUE = [68000.0, 352000.0, 486000.0, 423000.0, 579000.0]
CHART_REVENUE_ALT = [6.8, 35.2, 48.6, 42.3, 57.9]
CHART_RATIO = [1.42, 1.56, 1.71, 1.63, 1.85]
CHART_SER_NAMES = ("营收数值", "盈利比值")

# 感谢聆听页原有 7 个对象（文字用于定位，图片/底板/短线用几何定位）
THANKS_TEXTS = [
    ("感谢聆听", 2.4, 5.8),
    ("营养品自播项目·内容运营组", 2.5, 8.3),
    ("许泠瞻", 2.5, 9.5),
    ("以更稳的流程、更清晰的表达、更快的反馈完成下一阶段工作。", 2.5, 12.1),
]

# 五页原状清单：封面／协作关系／核心品类／能力成长／团队协作
COVER_SPEC = {
    "count": 10,
    "texts": [
        ("直播运营阶段复盘", 2.0, 4.3),
        ("营养品自播项目·内容运营组", 2.1, 8.9),
        ("许泠瞻·2026年春季", 2.1, 10.0),
        ("围绕主播执行、互动场控与客户服务的阶段性整理。", 2.0, 11.8),
        ("阶段复盘", 2.3, 2.3),
    ],
    "pic": (16.1, 0.0, 17.7, 19.05),
}

COLLAB_SPEC = {
    "count": 18,
    "bars": [
        ("销售管理中心", 2.1, 3.9),
        ("营养健康事业部", 3.0, 5.8),
        ("社群与内容业务部", 4.0, 7.7),
        ("短视频运营组", 5.0, 9.6),
        ("自播执行小组", 5.9, 11.5),
        ("内容运营专员", 6.9, 13.3),
    ],
    "note": ("虚构组织名称用于公开汇报表达，层级关系保留“中心—部门—小组—岗位”的业务逻辑。", 21.6, 12.6),
    "pic": (22.0, 3.0, 9.5, 7.9),
}

CATEGORY_SPEC = {
    "count": 15,
    "sub": ("复合营养片与矿物质补充场景", 5.1, 1.3),
    "cards": [3.4, 6.9, 10.4],
    "card_left": 16.6,
    "card_size": (14.9, 2.7),
    "titles": ["钙质支持", "铁元素关注", "锌元素补充"],
    "bodies": [
        ["骨骼与牙齿维护", "运动后恢复辅助", "家庭日常营养补给"],
        ["活力状态管理", "特殊阶段营养补充", "搭配饮食建议"],
        ["成长发育支持", "食欲与代谢相关沟通", "亲子人群常见咨询"],
    ],
    "bottom": ("表达原则：突出适用场景，弱化绝对化承诺；用通俗语言解释配方逻辑。", 16.9, 14.6),
    "pic": (1.8, 3.4, 13.8, 11.0),
}

GROWTH_SPEC = {
    "count": 17,
    "axis": (5.6, 7.6, 21.3),
    "dots": [(4.7, 6.6), (15.2, 6.7), (25.8, 6.7)],
    "dot_size": (1.8, 1.8),
    "cards": [3.3, 13.0, 22.8],
    "card_top": 9.2,
    "card_size": (6.9, 4.1),
    "titles": ["熟悉期", "提升期", "稳定期"],
    "bodies": [
        ["了解产品基础信息", "掌握开播前准备", "完成基础讲解任务"],
        ["能处理常见咨询", "独立配合场控节奏", "开始参与复盘整理"],
        ["形成个人话术库", "协助新人熟悉流程", "支持活动型直播"],
    ],
    "bottom": ("核心变化：少一点“被动执行”，多一点“主动发现与提前准备”", 6.0, 15.4),
}

TEAM_SPEC = {
    "count": 15,
    "cards": [(15.9, 3.4), (24.0, 3.4)],
    "card_size": (7.2, 3.2),
    "titles": ["主播协作", "运营协作"],
    "bodies": ["共享话术、提醒节奏、复盘高频问题", "同步活动重点、反馈用户关注点、配合调整机制"],
    "keep_card": (15.9, 7.7, 15.4, 3.6),
    "bar": (15.9, 12.6, 15.4, 1.9),
    "bar_text": "同一条信息在直播前、直播中、直播后反复确认，减少遗漏。",
    "pic": (1.8, 3.2, 13.0, 10.8),
    "keep_body": (16.5, 9.4, 14.1, 1.6),
    "keep_lines": ["复盘清单每日更新；复盘清单每日更新", "用户问题及时记录；用户问题及时记录"],
}

DUTY_SPEC = {
    "count": 15,
    "cards": [2.1, 11.7, 21.4],
    "card_top": 3.7,
    "card_size": (8.3, 11.0),
    "arrows": [(10.6, 8.3), (20.3, 8.3)],
    "arrow_size": (0.9, 1.2),
    "sub": ("多角色协同，保障直播链路稳定", 5.1, 1.3),
    "bottom": ("一人多岗，重点在“信息准确、节奏顺畅、反馈及时”", 5.4, 15.7),
    "t1": ("主播执行", 2.7, 4.2),
    "b1": (["按脚本完成产品讲解与互动引导", "根据实时反馈调整表达节奏", "沉淀常见问题与标准回应"], 2.5, 5.3),
    "t2": ("互动场控", 12.6, 4.1),
    "b2": (["关注评论氛围与咨询集中点", "提醒主播补充关键信息", "记录异常情况并同步运营"], 12.2, 5.4),
    "b3": ["协助排班、物料确认与备注核对", "跟进用户咨询与售后反馈", "整理复盘要点，推动流程闭环"],
}

WORK_SPEC = {
    "card_size": (7.0, 3.4),
    # 原文件卡底为白色系填充：两张 F9FEFE、两张 FFFFFF
    "card_fills": ("F9FEFE", "FFFFFF"),
    "groups": [
        ("直播执行", "按计划完成日常场次，讲解流程更顺滑"),
        ("互动维护", "评论区咨询响应更快，氛围更可控"),
        ("服务跟进", "备注核对、订单反馈与售后转交更清晰"),
        ("排班协调", "高峰时段人手安排更均衡"),
    ],
    "title_pt": 14.0,
    "title_color": "0F5963",
    "body_pt": 11.2,
    "body_color": "203942",
}

REVIEW_SPEC = {
    "tri": (16.2, 7.6, 1.4, 1.4),
    "cards": [2.3, 18.9],
    "card_top": 3.7,
    "card_size": (12.6, 10.3),
    "bottom": ("把一次问题变成一套流程", 11.0, 15.2, 28.0, "0F5963"),
    "left_title": ("风险表现", 2.9, 4.2, 32.0, "A85E1C"),
    "left_body_pos": (2.9, 5.3),
    "left_body": ["直播话术中出现表达边界不够清晰的情况",
                  "部分用户问题未能在第一时间形成标准回复",
                  "活动节奏变化时，信息同步偶有滞后"],
    "opt_lines": ["建立开播前话术检查表", "将敏感表达改为场景化、说明式表达",
                  "每日收场后进行小结并同步给运营"],
    "opt_dup": "将敏感表达改为场景化、说明式表达",
}

PLAN_SPEC = {
    "card_size": (8.5, 6.5),
    "card_top": 5.3,
    "near_title": "近期",
    "near_body": ["完善话术卡片", "稳定开播前检查", "提升互动响应质量"],
    "long_title": "长期",
    "long_body": ["搭建标准化复盘范式", "强化岗位协同配合效能", "赋能团队稳步长效发展"],
    "long_body_size": (7.2, 4.5),
    "mid_title": "中期",
    "mid_body": ["沉淀品类知识库", "参与活动直播策划", "协助新人熟悉流程"],
    "removed_line": "稳住基本盘，再把经验变成流程。",
    "bg_pic": (0.0, 2.2, 33.9, 15.2),
}

# 封面、协作关系、核心品类、团队协作、感谢聆听 5 页各 1 张图片（页索引, x, y, w, h）
PIC_SPEC = [
    (0, 16.1, 0.0, 17.7, 19.05),
    (2, 22.0, 3.0, 9.5, 7.9),
    (3, 1.8, 3.4, 13.8, 11.0),
    (7, 1.8, 3.2, 13.0, 10.8),
    (9, 17.7, 0.0, 16.2, 19.05),
]

# 封面页既有的空文本框（正常，不计入空对象）
COVER_EMPTY_BOX = (1.8, 17.9, 14.0, 0.5)

# 临时批注类文字
TEMP_TEXT_KEYS = ("已改", "待确认", "待修改", "待核对", "待补充", "待定", "修改说明",
                  "请审核", "已完成修改", "TODO", "todo", "批注")

# 新增说明性文字 / 水印 / 页码 / Logo 类文字
EXTRA_TEXT_KEYS = ("数据来源", "资料来源", "来源：", "注：", "说明：", "备注：", "制表",
                   "制图", "水印", "已按要求", "本页由", "logo", "LOGO", "Logo")

# 常见标注类几何图形（原文件只用 rect/roundRect/line/chevron/triangle/ellipse）
ANNOT_GEOMS = ("wedgeRectCallout", "wedgeEllipseCallout", "wedgeRoundRectCallout",
               "cloudCallout", "borderCallout1", "borderCallout2", "borderCallout3",
               "accentCallout1", "accentCallout2", "accentCallout3",
               "rightArrow", "leftArrow", "upArrow", "downArrow", "leftRightArrow",
               "bentArrow", "curvedRightArrow", "curvedLeftArrow", "notchedRightArrow",
               "stripedRightArrow", "quadArrow", "circularArrow")


def _cm(value):
    """EMU → 厘米。1 cm = 360000 EMU。"""
    return None if value is None else float(value) / EMU_PER_CM


def _decimals(tol):
    """取容差写出的有效位数：0.3 → 1 位，0.1 → 1 位，1 → 0 位。"""
    text = ("%f" % float(tol)).rstrip("0")
    return len(text.split(".")[1]) if "." in text and text.split(".")[1] else 0


def _le(diff, tol):
    """容差按写出的有效位比较：量测值先四舍五入到该精度再判，避免刀锋判定。"""
    return round(abs(float(diff)), _decimals(tol)) <= float(tol) + 1e-9


def _near(value, target, tol=TOL):
    return False if value is None else _le(value - target, tol)


def _pos_ok(shape, x, y, tol=TOL):
    """(X, Y) 指对象左上角，即 a:off；旋转不改变 a:off。"""
    if shape is None:
        return False
    return _near(_cm(shape.left), x, tol) and _near(_cm(shape.top), y, tol)


def _size_ok(shape, w, h, tol=TOL):
    """尺寸即 a:ext（cx, cy）。"""
    if shape is None:
        return False
    ok = True
    if w is not None:
        ok = ok and _near(_cm(shape.width), w, tol)
    if h is not None:
        ok = ok and _near(_cm(shape.height), h, tol)
    return ok


def _geom_ok(shape, x, y, w, h, tol=TOL):
    return _pos_ok(shape, x, y, tol) and _size_ok(shape, w, h, tol)


def _iter_shapes(container):
    """展开组合，便于文字/字体一类的全篇扫描；对象计数仍按顶层形状。"""
    out = []
    for shape in container:
        out.append(shape)
        if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
            out.extend(_iter_shapes(shape.shapes))
    return out


def _text_of(shape):
    try:
        return shape.text_frame.text if shape.has_text_frame else ""
    except Exception:
        return ""


def _norm(text):
    """比文字只比字形：去掉所有空白（行尾多余空格不扣）。"""
    return re.sub(r"\s+", "", text or "")


def _plain(text):
    """再去掉项目符号前缀，用于比较正文逐字内容。"""
    return _norm(text).replace("•", "").replace("·", "")


def _find_exact(shapes, text):
    key = _norm(text)
    for shape in shapes:
        if _norm(_text_of(shape)) == key:
            return shape
    return None


def _find_contains(shapes, text):
    key = _norm(text)
    if not key:
        return None
    for shape in shapes:
        if key in _norm(_text_of(shape)):
            return shape
    return None


def _find_all_contains(shapes, text):
    key = _norm(text)
    return [s for s in shapes if key and key in _norm(_text_of(s))]


def _find_geom(shapes, x=None, y=None, w=None, h=None, tol=TOL, exclude=()):
    """按几何定位对象：不同办公软件写出的形状名不一致，只认 a:off / a:ext。"""
    for shape in shapes:
        if any(shape is e for e in exclude):
            continue
        if shape.left is None or shape.top is None:
            continue
        if x is not None and not _near(_cm(shape.left), x, tol):
            continue
        if y is not None and not _near(_cm(shape.top), y, tol):
            continue
        if w is not None and not _near(_cm(shape.width), w, tol):
            continue
        if h is not None and not _near(_cm(shape.height), h, tol):
            continue
        return shape
    return None


def _pictures(slide):
    return [s for s in slide.shapes if s.shape_type == MSO_SHAPE_TYPE.PICTURE]


def _rot(shape):
    try:
        return float(shape.rotation) % 360.0
    except Exception:
        return 0.0


def _sp_pr(shape):
    return shape._element.find(qn("p:spPr"))


def _fill_srgb(shape):
    """形状自身填充色：只认 spPr 下直挂的 a:solidFill/a:srgbClr。"""
    sp_pr = _sp_pr(shape)
    if sp_pr is None:
        return None
    fill = sp_pr.find(qn("a:solidFill"))
    if fill is None:
        return None
    clr = fill.find(qn("a:srgbClr"))
    return clr.get("val").upper() if clr is not None and clr.get("val") else None


def _is_no_fill(shape):
    sp_pr = _sp_pr(shape)
    return sp_pr is not None and sp_pr.find(qn("a:noFill")) is not None


def _line_srgb(shape):
    """线条色：a:ln/a:solidFill/a:srgbClr。"""
    sp_pr = _sp_pr(shape)
    if sp_pr is None:
        return None
    ln = sp_pr.find(qn("a:ln"))
    if ln is None:
        return None
    fill = ln.find(qn("a:solidFill"))
    if fill is None:
        return None
    clr = fill.find(qn("a:srgbClr"))
    return clr.get("val").upper() if clr is not None and clr.get("val") else None


def _has_visible_line(shape):
    sp_pr = _sp_pr(shape)
    if sp_pr is None:
        return False
    ln = sp_pr.find(qn("a:ln"))
    if ln is None:
        return False
    return ln.find(qn("a:noFill")) is None


def _has_shadow(shape):
    return "outerShdw" in shape._element.xml


def _prst_geom(shape):
    sp_pr = _sp_pr(shape)
    if sp_pr is None:
        return None
    geom = sp_pr.find(qn("a:prstGeom"))
    return geom.get("prst") if geom is not None else None


def _slide_bg(slide):
    """页面背景：p:cSld/p:bg。返回 ('srgb', 'EAF8F9') / ('scheme', 'accent4') / ('other', None) / (None, None)。"""
    c_sld = slide._element.find(qn("p:cSld"))
    bg = c_sld.find(qn("p:bg")) if c_sld is not None else None
    if bg is None:
        return (None, None)
    xml = bg.xml
    m = re.search(r"<a:srgbClr val=\"([0-9A-Fa-f]{6})\"", xml)
    if m and "<a:gradFill" not in xml and "<a:blipFill" not in xml:
        return ("srgb", m.group(1).upper())
    m = re.search(r"<a:schemeClr val=\"([^\"]+)\"", xml)
    if m and "<a:gradFill" not in xml and "<a:blipFill" not in xml:
        return ("scheme", m.group(1))
    return ("other", None)


def _hex_in_range(value, lo, hi):
    if not value or len(value) != 6:
        return False
    try:
        rgb = (int(value[0:2], 16), int(value[2:4], 16), int(value[4:6], 16))
    except ValueError:
        return False
    return all(lo[i] <= rgb[i] <= hi[i] for i in range(3))


def _paragraphs(shape):
    if shape is None or not shape.has_text_frame:
        return []
    return list(shape.text_frame.paragraphs)


def _nonempty_lines(shape, plain=True):
    conv = _plain if plain else _norm
    return [conv(p.text) for p in _paragraphs(shape) if _norm(p.text)]


def _r_pr(run):
    return run._r.find(qn("a:rPr")) if run is not None else None


def _first_run(paragraph):
    runs = paragraph.runs
    return runs[0] if runs else None


def _run_size_pt(run):
    r_pr = _r_pr(run)
    if r_pr is None or not r_pr.get("sz"):
        return None
    return int(r_pr.get("sz")) / 100.0


def _run_bold(run):
    r_pr = _r_pr(run)
    return r_pr is not None and r_pr.get("b") in ("1", "true", "on")


def _run_color(run):
    r_pr = _r_pr(run)
    if r_pr is None:
        return None
    fill = r_pr.find(qn("a:solidFill"))
    if fill is None:
        return None
    clr = fill.find(qn("a:srgbClr"))
    return clr.get("val").upper() if clr is not None and clr.get("val") else None


def _def_r_pr_chain(shape, paragraph):
    """字体继承链的中间两级：段落 a:pPr/a:defRPr → 形状 a:lstStyle/a:lvl1pPr/a:defRPr。"""
    chain = []
    p_pr = paragraph._pPr
    if p_pr is not None:
        node = p_pr.find(qn("a:defRPr"))
        if node is not None:
            chain.append(node)
    if shape is not None and shape.has_text_frame:
        tx_body = shape.text_frame._txBody
        lst = tx_body.find(qn("a:lstStyle"))
        if lst is not None:
            lvl1 = lst.find(qn("a:lvl1pPr"))
            if lvl1 is not None:
                node = lvl1.find(qn("a:defRPr"))
                if node is not None:
                    chain.append(node)
    return chain


def _typeface(node, tag):
    if node is None:
        return None
    child = node.find(qn("a:" + tag))
    if child is None:
        return None
    name = child.get("typeface")
    return name if name else None


def _resolve_effective_eastasia(shape, paragraph, run, theme_ea):
    """中文字体：run 的 a:ea → 段落 defRPr → 形状 lstStyle → 主题 minorFont 的 a:ea。
    主题 a:ea 为空时中文会回落到系统字体，此处返回 None，视为未设置。"""
    for node in [_r_pr(run)] + _def_r_pr_chain(shape, paragraph):
        name = _typeface(node, "ea")
        if name:
            return name
    return theme_ea or None


def _resolve_effective_ascii(shape, paragraph, run, theme_latin):
    """西文字体：run 的 a:latin → 段落 defRPr → 形状 lstStyle → 主题 minorFont 的 a:latin。"""
    for node in [_r_pr(run)] + _def_r_pr_chain(shape, paragraph):
        name = _typeface(node, "latin")
        if name:
            return name
    return theme_latin or None


def _is_yahei(name):
    """“微软雅黑”与“Microsoft YaHei”是同一字体的中英文名称，两种写法均可。"""
    return (name or "").strip() in YAHEI_NAMES


def _line_spacing_pct(shape, paragraph):
    """段落行距：a:pPr/a:lnSpc/a:spcPct（200000 = 200%）；段落未写时回落形状 lstStyle。"""
    for holder in ([paragraph._pPr] if paragraph._pPr is not None else []):
        node = holder.find(qn("a:lnSpc"))
        if node is not None:
            pct = node.find(qn("a:spcPct"))
            return int(pct.get("val")) / 1000.0 if pct is not None else None
    if shape is not None and shape.has_text_frame:
        lst = shape.text_frame._txBody.find(qn("a:lstStyle"))
        lvl1 = lst.find(qn("a:lvl1pPr")) if lst is not None else None
        node = lvl1.find(qn("a:lnSpc")) if lvl1 is not None else None
        if node is not None:
            pct = node.find(qn("a:spcPct"))
            return int(pct.get("val")) / 1000.0 if pct is not None else None
    return None


def _theme_fonts(parts):
    """主题 majorFont / minorFont 的 latin 与 ea。"""
    xml = parts.get("ppt/theme/theme1.xml", "")
    out = {}
    for tag in ("majorFont", "minorFont"):
        m = re.search(r"<a:%s>(.*?)</a:%s>" % (tag, tag), xml, re.S)
        block = m.group(1) if m else ""
        latin = re.search(r"<a:latin typeface=\"([^\"]*)\"", block)
        ea = re.search(r"<a:ea typeface=\"([^\"]*)\"", block)
        out[tag] = (latin.group(1) if latin else None, ea.group(1) if ea else None)
    return out


def _rect(shape):
    left = _cm(shape.left) or 0.0
    top = _cm(shape.top) or 0.0
    return (left, top, left + (_cm(shape.width) or 0.0), top + (_cm(shape.height) or 0.0))


def _overlap(a, b, slack=0.05):
    """两矩形是否重叠（留 0.05 cm 贴边余量）。"""
    return (a[0] < b[2] - slack and b[0] < a[2] - slack
            and a[1] < b[3] - slack and b[1] < a[3] - slack)


def _inside_page(shape, slack=0.05):
    left, top, right, bottom = _rect(shape)
    return (left >= -slack and top >= -slack
            and right <= PAGE_W_CM + slack and bottom <= PAGE_H_CM + slack)


def _find_chart_shape(prs):
    """原生图表对象：p:graphicFrame 内挂 c:chart 关系（python-pptx 的 has_chart）。
    静态图片/截图贴图没有图表部件，此处一律取不到。"""
    for idx, slide in enumerate(prs.slides):
        for shape in _iter_shapes(slide.shapes):
            if getattr(shape, "has_chart", False):
                return (idx, shape)
    return (-1, None)


def _chart_xml(shape):
    try:
        return shape.chart.part.blob.decode("utf-8", "ignore")
    except Exception:
        return ""


def _chart_sers(xml):
    """返回 [(系列名, [数值...]), ...]。"""
    out = []
    for block in re.findall(r"<c:ser>(.*?)</c:ser>", xml, re.S):
        name = ""
        m = re.search(r"<c:tx>.*?<c:v>(.*?)</c:v>", block, re.S)
        if m:
            name = _norm(m.group(1))
        vals = []
        m = re.search(r"<c:val>(.*?)</c:val>", block, re.S)
        if m:
            for raw in re.findall(r"<c:v>([^<]*)</c:v>", m.group(1)):
                try:
                    vals.append(float(raw))
                except ValueError:
                    pass
        out.append((name, vals))
    return out


def _chart_cats(xml):
    m = re.search(r"<c:cat>(.*?)</c:cat>", xml, re.S)
    if not m:
        return []
    return [_norm(v) for v in re.findall(r"<c:v>([^<]*)</c:v>", m.group(1))]


def _chart_smooth(xml):
    """线条是否为平滑曲线：每个系列的 c:smooth 必须为 1（普通折线为 0 或缺失）。"""
    blocks = re.findall(r"<c:ser>(.*?)</c:ser>", xml, re.S)
    if not blocks:
        return False
    for block in blocks:
        m = re.search(r"<c:smooth val=\"([^\"]+)\"", block)
        if m is None or m.group(1) not in ("1", "true"):
            return False
    return True


def _chart_has_secondary_axis(xml):
    """次坐标轴：plotArea 内出现 4 个不同 c:axId（两组坐标轴）。"""
    return len(set(re.findall(r"<c:axId val=\"(\d+)\"", xml))) >= 4


def _values_match(vals, expect, rel=0.005):
    if len(vals) != len(expect):
        return False
    for got, want in zip(vals, expect):
        if abs(got - want) > max(abs(want) * rel, 1e-6):
            return False
    return True


def _pic_distortion(pic):
    """宽高比失真：显示框宽高比 ÷ 原图宽高比 − 1。整体等比缩放为 0。"""
    try:
        px, py = pic.image.size
        dpi_x, dpi_y = pic.image.dpi
        native = (px / float(dpi_x or 72)) / (py / float(dpi_y or 72))
        box = float(pic.width) / float(pic.height)
        return abs(box / native - 1.0)
    except Exception:
        return None


def _pic_cropped(pic):
    try:
        return max(abs(pic.crop_left), abs(pic.crop_right),
                   abs(pic.crop_top), abs(pic.crop_bottom)) > 0.001
    except Exception:
        return False


def _open_presentation(file_path):
    """打开 .pptx；加密只读、结构损坏、实际为 .ppt/.pdf/图片改名的，这里就会抛异常。"""
    return Presentation(file_path)


def _read_parts(file_path):
    """读出包内所有 xml / rels 文本，供主题、母版、动画、切换、批注等整包检查使用。"""
    parts = {}
    with zipfile.ZipFile(file_path) as zf:
        for name in zf.namelist():
            if name.endswith(".xml") or name.endswith(".rels"):
                parts[name] = zf.read(name).decode("utf-8", "ignore")
            else:
                parts[name] = ""
    return parts


def _find_lines(shapes, lines):
    want = [_plain(x) for x in lines]
    for shape in shapes:
        if _nonempty_lines(shape) == want:
            return shape
    return None


def _check_cover(slide):
    shapes = list(slide.shapes)
    if len(shapes) != COVER_SPEC["count"]:
        return False
    for text, x, y in COVER_SPEC["texts"]:
        shape = _find_exact(shapes, text)
        if not _pos_ok(shape, x, y):
            return False
    pics = _pictures(slide)
    x, y, w, h = COVER_SPEC["pic"]
    return len(pics) == 1 and _geom_ok(pics[0], x, y, w, h)


def _check_collab(slide):
    shapes = list(slide.shapes)
    if len(shapes) != COLLAB_SPEC["count"]:
        return False
    names = [t for t, _, _ in COLLAB_SPEC["bars"]]
    for text, x, y in COLLAB_SPEC["bars"]:
        if _find_geom(shapes, x=x, y=y, h=1.2) is None:
            return False
    labels = [s for s in shapes if _norm(_text_of(s)) in [_norm(n) for n in names]]
    if len(labels) != len(names):
        return False
    order = [_norm(_text_of(s)) for s in sorted(labels, key=lambda s: s.top)]
    if order != [_norm(n) for n in names]:
        return False
    text, x, y = COLLAB_SPEC["note"]
    if not _pos_ok(_find_exact(shapes, text), x, y):
        return False
    pics = _pictures(slide)
    x, y, w, h = COLLAB_SPEC["pic"]
    return len(pics) == 1 and _geom_ok(pics[0], x, y, w, h)


def _check_category(slide):
    shapes = list(slide.shapes)
    if len(shapes) != CATEGORY_SPEC["count"]:
        return False
    text, x, y = CATEGORY_SPEC["sub"]
    if not _pos_ok(_find_exact(shapes, text), x, y):
        return False
    w, h = CATEGORY_SPEC["card_size"]
    for top in CATEGORY_SPEC["cards"]:
        if _find_geom(shapes, x=CATEGORY_SPEC["card_left"], y=top, w=w, h=h) is None:
            return False
    titles = []
    for name in CATEGORY_SPEC["titles"]:
        shape = _find_exact(shapes, name)
        if shape is None:
            return False
        titles.append(shape)
    if [t for t in sorted(titles, key=lambda s: s.top)] != titles:
        return False
    for lines in CATEGORY_SPEC["bodies"]:
        if _find_lines(shapes, lines) is None:
            return False
    text, x, y = CATEGORY_SPEC["bottom"]
    if not _pos_ok(_find_exact(shapes, text), x, y):
        return False
    pics = _pictures(slide)
    x, y, w, h = CATEGORY_SPEC["pic"]
    return len(pics) == 1 and _geom_ok(pics[0], x, y, w, h)


def _check_growth(slide):
    shapes = list(slide.shapes)
    if len(shapes) != GROWTH_SPEC["count"]:
        return False
    x, y, length = GROWTH_SPEC["axis"]
    if _find_geom(shapes, x=x, y=y, w=length) is None:
        return False
    dw, dh = GROWTH_SPEC["dot_size"]
    for x, y in GROWTH_SPEC["dots"]:
        if _find_geom(shapes, x=x, y=y, w=dw, h=dh) is None:
            return False
    cw, ch = GROWTH_SPEC["card_size"]
    for x in GROWTH_SPEC["cards"]:
        if _find_geom(shapes, x=x, y=GROWTH_SPEC["card_top"], w=cw, h=ch) is None:
            return False
    for name in GROWTH_SPEC["titles"]:
        if _find_exact(shapes, name) is None:
            return False
    for lines in GROWTH_SPEC["bodies"]:
        if _find_lines(shapes, lines) is None:
            return False
    text, x, y = GROWTH_SPEC["bottom"]
    return _pos_ok(_find_exact(shapes, text), x, y)


def _check_team(slide):
    shapes = list(slide.shapes)
    if len(shapes) != TEAM_SPEC["count"]:
        return False
    w, h = TEAM_SPEC["card_size"]
    for x, y in TEAM_SPEC["cards"]:
        if _find_geom(shapes, x=x, y=y, w=w, h=h) is None:
            return False
    for name in TEAM_SPEC["titles"]:
        if _find_exact(shapes, name) is None:
            return False
    for body in TEAM_SPEC["bodies"]:
        if _find_exact(shapes, body) is None:
            return False
    x, y, w, h = TEAM_SPEC["keep_card"]
    if _find_geom(shapes, x=x, y=y, w=w, h=h) is None:
        return False
    x, y, w, h = TEAM_SPEC["bar"]
    if _find_geom(shapes, x=x, y=y, w=w, h=h) is None:
        return False
    if not _find_exact(shapes, TEAM_SPEC["bar_text"]):
        return False
    pics = _pictures(slide)
    x, y, w, h = TEAM_SPEC["pic"]
    return len(pics) == 1 and _geom_ok(pics[0], x, y, w, h)


def _check_team_keep_body(slide):
    """“重复沉淀”卡内那两句重复写法属正常内容，须原样保留。"""
    shapes = list(slide.shapes)
    body = _find_lines(shapes, TEAM_SPEC["keep_lines"])
    if body is None:
        return False
    x, y, w, h = TEAM_SPEC["keep_body"]
    if not _geom_ok(body, x, y, w, h):
        return False
    return len([p for p in _paragraphs(body) if _norm(p.text)]) == 2


def _bansu_shapes(slide):
    """岗位职责页新增的“排班岗”标题对象（可能被写成文本框或形状）。"""
    return [s for s in slide.shapes if _norm(_text_of(s)) == "排班岗"]


def _check_duty(slide):
    bansu = [s._element for s in _bansu_shapes(slide)]
    shapes = [s for s in slide.shapes if s._element not in bansu]
    if len(shapes) != DUTY_SPEC["count"]:
        return False
    w, h = DUTY_SPEC["card_size"]
    for x in DUTY_SPEC["cards"]:
        if _find_geom(shapes, x=x, y=DUTY_SPEC["card_top"], w=w, h=h) is None:
            return False
    aw, ah = DUTY_SPEC["arrow_size"]
    for x, y in DUTY_SPEC["arrows"]:
        if _find_geom(shapes, x=x, y=y, w=aw, h=ah) is None:
            return False
    for key in ("sub", "bottom", "t1", "t2"):
        text, x, y = DUTY_SPEC[key]
        if not _pos_ok(_find_exact(shapes, text), x, y):
            return False
    for key in ("b1", "b2"):
        lines, x, y = DUTY_SPEC[key]
        body = _find_lines(shapes, lines)
        if body is None or not _pos_ok(body, x, y):
            return False
    # 第三卡正文的 Y 坐标允许上下微调以容纳标题，只查文字逐字一致与左边缘。
    body3 = _find_lines(shapes, DUTY_SPEC["b3"])
    return body3 is not None


def _duty_title_baseline(slide):
    """返回 (前两卡标题顶端列表, 前两卡“标题左边缘相对卡底左边缘”偏移均值)。
    前两卡该偏移本身有 0.2 cm 上下，故按均值比，避免刀锋判定。"""
    shapes = list(slide.shapes)
    tops, offsets = [], []
    for key, card_x in (("t1", DUTY_SPEC["cards"][0]), ("t2", DUTY_SPEC["cards"][1])):
        text = DUTY_SPEC[key][0]
        title = _find_exact(shapes, text)
        card = _find_geom(shapes, x=card_x, y=DUTY_SPEC["card_top"],
                          w=DUTY_SPEC["card_size"][0], h=DUTY_SPEC["card_size"][1])
        if title is None or card is None:
            return (None, None)
        tops.append(_cm(title.top))
        offsets.append(_cm(title.left) - _cm(card.left))
    return (tops, sum(offsets) / len(offsets))


def _work_cards(slide):
    """工作成效页 4 张卡底：按圆角矩形 + 7.0×3.4 尺寸认（原文件 2 张 F9FEFE、2 张 FFFFFF）。"""
    w, h = WORK_SPEC["card_size"]
    return [s for s in slide.shapes
            if _prst_geom(s) == "roundRect" and _size_ok(s, w, h)]


def _work_triples(slide):
    """把 4 组“卡底 + 小标题 + 正文”配起来；标题/正文按文字内容认，归属按最近卡底。"""
    cards = _work_cards(slide)
    shapes = list(slide.shapes)
    triples = []
    for title_text, body_text in WORK_SPEC["groups"]:
        title = _find_exact(shapes, title_text)
        body = _find_exact(shapes, body_text)
        if title is None or body is None or not cards:
            return []
        card = min(cards, key=lambda c: (abs(_cm(c.left) - _cm(title.left))
                                         + abs(_cm(c.top) - _cm(title.top))))
        triples.append((card, title, body))
    if len({id(c) for c, _, _ in triples}) != 4:
        return []
    return triples


def _check_work_cards(slide):
    cards = _work_cards(slide)
    if len(cards) != 4:
        return False
    for card in cards:
        if _prst_geom(card) != "roundRect" or not _has_shadow(card):
            return False
        if _fill_srgb(card) not in WORK_SPEC["card_fills"]:
            return False
    triples = _work_triples(slide)
    if len(triples) != 4:
        return False
    for _, title, body in triples:
        t_run = _first_run(_paragraphs(title)[0])
        b_run = _first_run(_paragraphs(body)[0])
        if not _near(_run_size_pt(t_run) or 0, WORK_SPEC["title_pt"], 0.1):
            return False
        if _run_color(t_run) != WORK_SPEC["title_color"]:
            return False
        if not _near(_run_size_pt(b_run) or 0, WORK_SPEC["body_pt"], 0.1):
            return False
        if _run_color(b_run) != WORK_SPEC["body_color"]:
            return False
    for getter in (lambda c, s: _cm(s.left) - _cm(c.left), lambda c, s: _cm(s.top) - _cm(c.top)):
        for pick in (1, 2):
            spread = [getter(t[0], t[pick]) for t in triples]
            if not _le(max(spread) - min(spread), 0.3):
                return False
    return True


def _check_review_layout(slide):
    shapes = list(slide.shapes)
    tris = [s for s in shapes if _prst_geom(s) == "triangle"]
    if len(tris) != 1:
        return False
    x, y, w, h = REVIEW_SPEC["tri"]
    if not _geom_ok(tris[0], x, y, w, h) or not _le(_rot(tris[0]) - 90.0, 1.0):
        return False
    w, h = REVIEW_SPEC["card_size"]
    for x in REVIEW_SPEC["cards"]:
        if _find_geom(shapes, x=x, y=REVIEW_SPEC["card_top"], w=w, h=h) is None:
            return False
    text, x, y, pt, color = REVIEW_SPEC["bottom"]
    shape = _find_exact(shapes, text)
    if not _pos_ok(shape, x, y):
        return False
    run = _first_run(_paragraphs(shape)[0])
    return _near(_run_size_pt(run) or 0, pt, 0.1) and _run_color(run) == color


def _check_review_left(slide):
    shapes = list(slide.shapes)
    text, x, y, pt, color = REVIEW_SPEC["left_title"]
    title = _find_exact(shapes, text)
    if not _pos_ok(title, x, y):
        return False
    run = _first_run(_paragraphs(title)[0])
    if not _near(_run_size_pt(run) or 0, pt, 0.1) or _run_color(run) != color:
        return False
    body = _find_lines(shapes, REVIEW_SPEC["left_body"])
    if body is None or not _pos_ok(body, *REVIEW_SPEC["left_body_pos"]):
        return False
    paras = [p for p in _paragraphs(body) if _norm(p.text)]
    if len(paras) != 3:
        return False
    for para in paras:
        run = _first_run(para)
        if not _near(_line_spacing_pct(body, para) or 0, 200.0, 1.0):
            return False
        if not _near(_run_size_pt(run) or 0, 16.0, 0.1):
            return False
        if _run_color(run) != "203942" or not para.text.strip().startswith("•"):
            return False
    return True


def _plan_cards(slide):
    """后续规划页的卡底：8.5×6.5、顶端 5.3；左右位置不查（删卡后重新居中属允许）。"""
    w, h = PLAN_SPEC["card_size"]
    return [s for s in slide.shapes
            if _size_ok(s, w, h) and _near(_cm(s.top), PLAN_SPEC["card_top"])]


def _card_of(cards, shape):
    for card in cards:
        left, right = _cm(card.left), _cm(card.left) + _cm(card.width)
        if left - 0.3 <= _cm(shape.left) and _cm(shape.left) + _cm(shape.width) <= right + 0.3:
            return card
    return None


def _check_plan_cards(slide):
    shapes = list(slide.shapes)
    cards = _plan_cards(slide)
    if len(cards) < 2:
        return False
    near_title = _find_exact(shapes, PLAN_SPEC["near_title"])
    long_title = _find_exact(shapes, PLAN_SPEC["long_title"])
    if near_title is None or long_title is None:
        return False
    near_card = _card_of(cards, near_title)
    long_card = _card_of(cards, long_title)
    if near_card is None or long_card is None or near_card is long_card:
        return False
    if not _inside_page(near_card) or not _inside_page(long_card):
        return False
    body = _find_lines(shapes, PLAN_SPEC["near_body"])
    if body is None:
        return False
    paras = [p for p in _paragraphs(body) if _norm(p.text)]
    if len(paras) != 3:
        return False
    for para in paras:
        run = _first_run(para)
        if not _near(_run_size_pt(run) or 0, 16.0, 0.1):
            return False
        if not _near(_line_spacing_pct(body, para) or 0, 150.0, 1.0):
            return False
        if _run_color(run) != "203942":
            return False
    return True


def _check_plan_bg(slide):
    x, y, w, h = PLAN_SPEC["bg_pic"]
    for pic in _pictures(slide):
        if _geom_ok(pic, x, y, w, h, tol=0.1) and not _pic_cropped(pic):
            return True
    return False


def _plan_residual_boxes(slide):
    """后续规划页上除卡底之外的空白框（残留的空卡底 / 空文本框）。"""
    keep = [c._element for c in _plan_cards(slide)]
    out = []
    for shape in _iter_shapes(slide.shapes):
        if shape.shape_type == MSO_SHAPE_TYPE.PICTURE or shape._element in keep:
            continue
        if _norm(_text_of(shape)):
            continue
        if _cm(shape.height) >= 0.3 and _cm(shape.width) >= 2.0:
            out.append(shape)
    return out


def _check_plan_mid_removed(slide):
    """+3：只剩近期／长期 2 张卡，"中期"标题与其 3 行正文均不存在，无残留空卡底。"""
    shapes = list(slide.shapes)
    if _find_exact(shapes, PLAN_SPEC["mid_title"]) is not None:
        return False
    for line in PLAN_SPEC["mid_body"]:
        if _find_contains(shapes, line) is not None:
            return False
    if len(_plan_cards(slide)) != 2:
        return False
    return not _plan_residual_boxes(slide)


def _check_plan_line_removed(slide):
    """+1：句子"稳住基本盘…"不存在，也没留下对应的空文本框。"""
    shapes = list(slide.shapes)
    if _find_contains(shapes, PLAN_SPEC["removed_line"]) is not None:
        return False
    if _find_contains(shapes, "稳住基本盘") is not None:
        return False
    return not _plan_residual_boxes(slide)


def _plan_long_body(slide):
    return _find_lines(list(slide.shapes), PLAN_SPEC["long_body"])


def _align_of(shape):
    """段落对齐；未写 algn 时按 PowerPoint 默认左对齐处理（显式 l 与缺省视为一致）。"""
    algn = []
    for para in _paragraphs(shape):
        p_pr = para._pPr
        algn.append((None if p_pr is None else p_pr.get("algn")) or "l")
    return algn


def _check_plan_long_format(slide, theme_latin, theme_ea):
    """+3：长期卡正文 16pt / 行距 150% / 203942 / 中英文微软雅黑 / 7.2×4.5，对齐与近期卡一致。"""
    body = _plan_long_body(slide)
    if body is None:
        return False
    w, h = PLAN_SPEC["long_body_size"]
    if not _size_ok(body, w, h):
        return False
    near = _find_lines(list(slide.shapes), PLAN_SPEC["near_body"])
    if near is not None and set(_align_of(body)) != set(_align_of(near)):
        return False
    for para in _paragraphs(body):
        if not _norm(para.text):
            continue
        if not _near(_line_spacing_pct(body, para) or 0, 150.0, 1.0):
            return False
        for run in para.runs:
            if not _near(_run_size_pt(run) or 0, 16.0, 0.1):
                return False
            if _run_color(run) != "203942":
                return False
            latin = _resolve_effective_ascii(body, para, run, theme_latin)
            ea = _resolve_effective_eastasia(body, para, run, theme_ea)
            if not _is_yahei(latin) or not _is_yahei(ea):
                return False
    return True


def _check_thanks(slide):
    """−3：感谢聆听页原有 7 个对象齐全且在位（新增图表及承载对象不计入个数）。"""
    shapes = list(slide.shapes)
    for text, x, y in THANKS_TEXTS:
        if not _pos_ok(_find_exact(shapes, text), x, y):
            return False
    pics = _pictures(slide)
    if not any(_geom_ok(p, 17.7, 0.0, 16.2, PAGE_H_CM) for p in pics):
        return False
    if _find_geom(shapes, x=0.0, y=0.0, w=17.7, h=PAGE_H_CM) is None:
        return False
    if _find_geom(shapes, x=2.5, y=11.3, w=3.2) is None:
        return False
    return True


def _check_chart_on_thanks(prs, chart_shape):
    """+1：图表落在感谢聆听页（不是单独占一页、也不在其他页）。"""
    for slide in prs.slides:
        for shape in _iter_shapes(slide.shapes):
            if shape._element is chart_shape._element:
                return _slide_key_ok(slide, PAGE_KEYS[9])
    return False


def _slide_key_ok(slide, key):
    text = _norm("".join(_text_of(s) for s in _iter_shapes(slide.shapes)))
    return _plain(key) in _plain(text)


def _page_by_key(prs, key):
    """按页内关键字定位页面：页数/页序出问题时元素级检查仍能落到正确的页上。"""
    for slide in prs.slides:
        if _slide_key_ok(slide, key):
            return slide
    return None


def _check_chart_geometry(prs, chart_shape):
    """+1：图表完整落在页面内、宽高比正常无拉伸，且不遮挡感谢聆听页的四处文字。"""
    if not _inside_page(chart_shape):
        return False
    w, h = _cm(chart_shape.width), _cm(chart_shape.height)
    if not w or not h or not (0.7 <= w / h <= 3.0):
        return False
    slide = _page_by_key(prs, PAGE_KEYS[9])
    if slide is None:
        return False
    for text, _x, _y in THANKS_TEXTS:
        shape = _find_exact(list(slide.shapes), text)
        if shape is not None and _overlap(_rect(shape), _rect(chart_shape)):
            return False
    return True


def _check_titles_and_lines(prs):
    """−3：岗位职责～后续规划每页有页面标题与青色短线；封面～后续规划每页有右下角横线。"""
    for key in PAGE_KEYS[1:9]:
        slide = _page_by_key(prs, key)
        if slide is None:
            return False
        shapes = list(slide.shapes)
        title = _find_geom(shapes, x=1.8, y=1.1, w=21.3, h=1.1)
        if title is None or _plain(_text_of(title)) != _plain(key):
            return False
        run = _first_run(_paragraphs(title)[0]) if _paragraphs(title) else None
        if not _near(_run_size_pt(run) or 0, 24.0, 0.1):
            return False
        if _find_geom(shapes, x=1.8, y=2.6, w=2.9) is None:
            return False
    for key in PAGE_KEYS[0:9]:
        slide = _page_by_key(prs, key)
        if slide is None or _find_geom(list(slide.shapes), x=27.1, y=18.1, w=4.8) is None:
            return False
    return True


def _check_page_order(prs):
    """−5：10 页，且页序为 封面／岗位职责／…／感谢聆听页。"""
    slides = list(prs.slides)
    if len(slides) != 10:
        return False
    for idx, key in enumerate(PAGE_KEYS):
        if not _slide_key_ok(slides[idx], key):
            return False
    return True


def _check_backgrounds(prs):
    """−3：各页背景纯色色值，页面标题色 17323A、青色短线 29A7A5、右下横线 B9E1E4。"""
    for key, expect in zip(PAGE_KEYS, BG_COLORS):
        if expect is None:          # 工作成效页背景由加分项单独评
            continue
        slide = _page_by_key(prs, key)
        if slide is None:
            return False
        kind, value = _slide_bg(slide)
        if kind != "srgb" or (value or "").upper() != expect:
            return False
    for key in PAGE_KEYS[1:9]:
        shapes = list(_page_by_key(prs, key).shapes)
        title = _find_geom(shapes, x=1.8, y=1.1, w=21.3, h=1.1)
        if title is not None and _paragraphs(title):
            if _run_color(_first_run(_paragraphs(title)[0])) != TITLE_COLOR:
                return False
        line = _find_geom(shapes, x=1.8, y=2.6, w=2.9)
        if line is not None and _line_srgb(line) != LINE_TEAL:
            return False
    for key in PAGE_KEYS[0:9]:
        line = _find_geom(list(_page_by_key(prs, key).shapes), x=27.1, y=18.1, w=4.8)
        if line is not None and _line_srgb(line) != LINE_LIGHT:
            return False
    return True


def _check_pictures(prs):
    """−3：封面、协作关系、核心品类、团队协作、感谢聆听 5 页各 1 张图片且位置尺寸不变。"""
    for idx, x, y, w, h in PIC_SPEC:
        slide = _page_by_key(prs, PAGE_KEYS[idx])
        if slide is None:
            return False
        pics = _pictures(slide)
        # 新增的图表/贴图不计入张数，只认原图是否仍在原位（与感谢聆听页口径一致）
        if not any(_geom_ok(p, x, y, w, h) for p in pics):
            return False
    return True


def _check_theme_and_fonts(prs, parts):
    """−5：主题 majorFont/minorFont 的 latin 为微软雅黑、ea 为空；单 master 单 layout；全篇正文微软雅黑。"""
    fonts = _theme_fonts(parts)
    if not fonts:
        return False
    for tag in ("majorFont", "minorFont"):
        latin, ea = fonts.get(tag, (None, None))
        if not _is_yahei(latin) or (ea or "").strip():
            return False
    masters = [n for n in parts if re.match(r"ppt/slideMasters/slideMaster\d+\.xml$", n)]
    layouts = [n for n in parts if re.match(r"ppt/slideLayouts/slideLayout\d+\.xml$", n)]
    if len(masters) != 1 or len(layouts) != 1:
        return False
    for slide in prs.slides:
        for shape in _iter_shapes(slide.shapes):
            for para in _paragraphs(shape):
                for run in para.runs:
                    r_pr = _r_pr(run)
                    for tag in ("latin", "ea"):
                        name = _typeface(r_pr, tag)
                        if name and not _is_yahei(name):
                            return False
    return True


def _check_no_animation(parts):
    """−1：全篇无动画、无页面切换效果。"""
    for name, xml in parts.items():
        if not re.match(r"ppt/slides/slide\d+\.xml$", name):
            continue
        if "<p:transition" in xml or "<mc:AlternateContent" in xml and "transition" in xml:
            return False
        if "<p:childTnLst" in xml:
            return False
    return True


def _check_no_annotations(prs, parts):
    """−3：无标注图形／批注／荧光笔标记，也无"已改""待确认"等临时文字。"""
    for name in parts:
        low = name.lower()
        if "comment" in low or "ink" in low.replace("linked", ""):
            return False
    for slide in prs.slides:
        if "<a:highlight" in slide._element.xml:
            return False
        for shape in _iter_shapes(slide.shapes):
            text = _norm(_text_of(shape))
            if any(key in text for key in TEMP_TEXT_KEYS):
                return False
            geom = _prst_geom(shape)
            if geom in ANNOT_GEOMS:
                return False
            if geom == "ellipse" and _is_no_fill(shape) and _has_visible_line(shape):
                return False
    return True


def _check_no_empty_objects(prs):
    """−1：除封面页既有的空文本框外，无空占位符、空文本框及无填充无线条的多余对象。"""
    for idx, slide in enumerate(prs.slides):
        for shape in _iter_shapes(slide.shapes):
            if shape.shape_type in (MSO_SHAPE_TYPE.PICTURE, MSO_SHAPE_TYPE.CHART,
                                    MSO_SHAPE_TYPE.TABLE, MSO_SHAPE_TYPE.GROUP):
                continue
            text = _text_of(shape)
            if "单击此处" in text or "click to add" in text.lower():
                return False
            if _norm(text):
                continue
            if shape.is_placeholder:
                return False
            if idx == 0 and _geom_ok(shape, *COVER_EMPTY_BOX):
                continue
            # 空对象：面积可见（长宽均成形）且无填充、无线条、无阴影 → 多余对象
            if _cm(shape.height) < 0.3 or _cm(shape.width) < 0.3:
                continue        # 装饰短线（高度为 0）等，起装饰作用
            if _is_no_fill(shape) and not _has_visible_line(shape) and not _has_shadow(shape):
                return False
    return True


def _check_no_extra_text(prs):
    """−1：无新增说明性文字、水印、页码、Logo。"""
    for slide in prs.slides:
        if 'type="slidenum"' in slide._element.xml:
            return False
        for shape in _iter_shapes(slide.shapes):
            text = _text_of(shape)
            if any(key in text for key in EXTRA_TEXT_KEYS):
                return False
    return True


def _work_objects(slide):
    """工作成效页 4 张卡片的 12 个对象：4 个卡底 + 4 个小标题 + 4 段正文（按文字认）。"""
    shapes = list(slide.shapes)
    out = _work_cards(slide)
    if len(out) != 4:
        return []
    for title_text, body_text in WORK_SPEC["groups"]:
        for text in (title_text, body_text):
            shape = _find_exact(shapes, text)
            if shape is None:
                return []
            out.append(shape)
    return out


def _check_work_rotation(slide):
    """+3：4 张卡片的卡底、小标题、正文共 12 个对象旋转角全部为 0°。"""
    objects = _work_objects(slide)
    if len(objects) != 12:
        return False
    return all(_le(_rot(shape), 0.01) for shape in objects)


def _check_work_pic_rotation(slide):
    """+1：右侧图片旋转角为 0°。"""
    pics = _pictures(slide)
    return len(pics) == 1 and _le(_rot(pics[0]), 0.01)


def _check_work_align(slide):
    """+3：左右两列卡底左边缘对齐、上下两行卡底顶端对齐，同列/同行偏差 ≤0.1 cm。"""
    cards = _work_cards(slide)
    if len(cards) != 4:
        return False
    lefts = sorted(_cm(c.left) for c in cards)
    tops = sorted(_cm(c.top) for c in cards)
    for pairs in (lefts, tops):
        if not _le(pairs[1] - pairs[0], 0.1) or not _le(pairs[3] - pairs[2], 0.1):
            return False
    if _le(lefts[2] - lefts[1], 0.1) or _le(tops[2] - tops[1], 0.1):
        return False        # 两列/两行须真正分开，避免 4 张挤成一列被判"对齐"
    return True


def _check_work_bg_range(slide):
    """+3：背景为浅青白，色值落在 EAF8F9–F6FBFC 区间。"""
    kind, value = _slide_bg(slide)
    return kind == "srgb" and _hex_in_range(value, BG_WORK_LO, BG_WORK_HI)


def _check_work_bg_pure(slide):
    """+1：背景为 srgbClr 纯色填充（非渐变、非图片、非主题色变体）。"""
    kind, _value = _slide_bg(slide)
    return kind == "srgb"


def _check_work_picture(slide):
    """−3：本页 1 张图片，画面完整、未被裁剪或替换，宽高比失真 ≤2%。"""
    pics = _pictures(slide)
    if len(pics) != 1:
        return False
    pic = pics[0]
    if _pic_cropped(pic):
        return False
    dist = _pic_distortion(pic)
    return dist is None or dist <= 0.02


def _check_work_no_overlap(slide):
    """−1：所有对象完整落在页面内，卡片之间、图片与卡片之间无重叠。"""
    cards = _work_cards(slide)
    pics = _pictures(slide)
    for shape in _iter_shapes(slide.shapes):
        if not _inside_page(shape):
            return False
    boxes = [_rect(s) for s in cards + pics]
    for i in range(len(boxes)):
        for j in range(i + 1, len(boxes)):
            if _overlap(boxes[i], boxes[j]):
                return False
    return True


def _opt_body(slide):
    """问题复盘页右卡"优化动作"下方的正文文本框：按含"检查表"一句定位。"""
    shapes = list(slide.shapes)
    header = _find_exact(shapes, "优化动作")
    cands = _find_all_contains(shapes, REVIEW_SPEC["opt_lines"][0])
    if not cands:
        return None
    if header is not None:
        cands = [c for c in cands if _cm(c.top) >= _cm(header.top) - 0.5] or cands
    return cands[0]


def _check_opt_dup(slide):
    """+3："将敏感表达改为场景化、说明式表达"只出现 1 次（整页统计）。"""
    key = _plain(REVIEW_SPEC["opt_dup"])
    hits = 0
    for shape in _iter_shapes(slide.shapes):
        for para in _paragraphs(shape):
            hits += _plain(para.text).count(key)
    return hits == 1


def _check_opt_paras(slide):
    """+1：优化动作正文无空段落，共 3 段且顺序正确。"""
    body = _opt_body(slide)
    if body is None:
        return False
    paras = _paragraphs(body)
    if any(not _norm(p.text) for p in paras):
        return False
    return [_plain(p.text) for p in paras] == [_plain(x) for x in REVIEW_SPEC["opt_lines"]]


def _check_opt_format(slide, theme_latin, theme_ea):
    """+1：优化动作 3 段格式一致：行距 200%、16pt、203942、中英文微软雅黑、前缀"• "。"""
    body = _opt_body(slide)
    if body is None:
        return False
    paras = [p for p in _paragraphs(body) if _norm(p.text)]
    if len(paras) != 3:
        return False
    for para in paras:
        if not _norm(para.text).startswith("•"):
            return False
        if not _near(_line_spacing_pct(body, para) or 0, 200.0, 1.0):
            return False
        for run in para.runs:
            if not _near(_run_size_pt(run) or 0, 16.0, 0.1):
                return False
            if _run_color(run) != "203942":
                return False
            latin = _resolve_effective_ascii(body, para, run, theme_latin)
            ea = _resolve_effective_eastasia(body, para, run, theme_ea)
            if not _is_yahei(latin) or not _is_yahei(ea):
                return False
    return True


def _bansu_title(slide):
    """第三卡（卡底左边缘 21.4、宽 8.3）内的"排班岗"标题；不在卡内则不算。"""
    shapes = list(slide.shapes)
    card = _find_geom(shapes, x=DUTY_SPEC["cards"][2], y=DUTY_SPEC["card_top"],
                      w=DUTY_SPEC["card_size"][0], h=DUTY_SPEC["card_size"][1])
    if card is None:
        return None
    cx1, cx2 = _cm(card.left), _cm(card.left) + _cm(card.width)
    cy1, cy2 = _cm(card.top), _cm(card.top) + _cm(card.height)
    for shape in _bansu_shapes(slide):
        x1, y1 = _cm(shape.left), _cm(shape.top)
        if cx1 - TOL <= x1 <= cx2 + TOL and cy1 - TOL <= y1 <= cy2 + TOL:
            return shape
    return None


def _check_bansu_title(slide):
    """+3：第三张卡片内有标题"排班岗"，文字准确、无错字、无多余字符。"""
    title = _bansu_title(slide)
    if title is None:
        return False
    return [_norm(p.text) for p in _paragraphs(title) if _norm(p.text)] == ["排班岗"]


def _check_bansu_format(slide, theme_latin, theme_ea):
    """+1：24pt、加粗、0F5963、中英文均为微软雅黑（只设 latin 不得分）。"""
    title = _bansu_title(slide)
    if title is None:
        return False
    paras = [p for p in _paragraphs(title) if _norm(p.text)]
    if not paras:
        return False
    for para in paras:
        for run in para.runs:
            if not _near(_run_size_pt(run) or 0, 24.0, 0.1) or not _run_bold(run):
                return False
            if _run_color(run) != "0F5963":
                return False
            latin = _resolve_effective_ascii(title, para, run, theme_latin)
            ea = _resolve_effective_eastasia(title, para, run, theme_ea)
            if not _is_yahei(latin) or not _is_yahei(ea):
                return False
    return True


def _check_bansu_baseline(slide):
    """+1：与前两卡标题同一基准：顶端偏差 ≤0.3 cm、卡内左偏移与前两卡均值差 ≤0.3 cm、未超卡片边界。"""
    title = _bansu_title(slide)
    if title is None:
        return False
    tops, offset_avg = _duty_title_baseline(slide)
    if not tops or offset_avg is None:
        return False
    if not all(_le(_cm(title.top) - t, TOL) for t in tops):
        return False
    card = _find_geom(list(slide.shapes), x=DUTY_SPEC["cards"][2], y=DUTY_SPEC["card_top"],
                      w=DUTY_SPEC["card_size"][0], h=DUTY_SPEC["card_size"][1])
    if card is None:
        return False
    if not _le((_cm(title.left) - _cm(card.left)) - offset_avg, TOL):
        return False
    right = _cm(title.left) + _cm(title.width)
    return right <= _cm(card.left) + _cm(card.width) + TOL


def _check_bansu_box(slide):
    """−1："排班岗"文本框无边框、无填充、无阴影，单行不折行，且不与第三卡正文重叠、不溢出卡片。"""
    title = _bansu_title(slide)
    if title is None:
        return True             # 标题不存在属加分项 0 分，本条不重复扣
    if not _is_no_fill(title) or _has_visible_line(title) or _has_shadow(title):
        return False
    if len([p for p in _paragraphs(title) if _norm(p.text)]) != 1:
        return False
    body3 = _find_lines(list(slide.shapes), DUTY_SPEC["b3"])
    if body3 is not None and _overlap(_rect(title), _rect(body3)):
        return False
    card = _find_geom(list(slide.shapes), x=DUTY_SPEC["cards"][2], y=DUTY_SPEC["card_top"],
                      w=DUTY_SPEC["card_size"][0], h=DUTY_SPEC["card_size"][1])
    if card is None:
        return False
    x1, y1 = _cm(title.left), _cm(title.top)
    x2, y2 = x1 + _cm(title.width), y1 + _cm(title.height)
    cx1, cy1 = _cm(card.left), _cm(card.top)
    cx2, cy2 = cx1 + _cm(card.width), cy1 + _cm(card.height)
    return (x1 >= cx1 - TOL and y1 >= cy1 - TOL
            and x2 <= cx2 + TOL and y2 <= cy2 + TOL)


# ============================================================
# 维度一：可用性 / 可修改性（门槛项，不计分）
#   点1：交付物是可正常打开的 .pptx（非 .ppt/.pdf/图片/压缩包，未加密只读，包结构完整）
#   点2：无“PowerPoint 已修复”迹象；若含图表对象，图表部件与关系完整、可正常显示
#   点3：页面尺寸为 33.87×19.05 cm（16:9）
#   任一条不满足 → 完成度直接判 0 分
# ============================================================


def check_dimension_one(prs, pptx_path):
    issues = []

    # 点1：扩展名与包结构
    if not pptx_path.lower().endswith(".pptx"):
        issues.append("交付物不是 .pptx 文件: %s" % os.path.basename(pptx_path))
    try:
        with zipfile.ZipFile(pptx_path) as z:
            names = z.namelist()
            if "ppt/presentation.xml" not in names:
                issues.append("包内缺少 ppt/presentation.xml，文件可能已损坏")
            bad = z.testzip()
            if bad:
                issues.append("包内文件损坏: %s" % bad)
            if not [n for n in names if re.match(r"ppt/slides/slide\d+\.xml$", n)]:
                issues.append("包内没有任何幻灯片，无法评估")
    except Exception as e:
        issues.append("文件无法作为 .pptx 包读取（可能损坏或加密）: %s" % e)

    # 点1：文件系统只读属性
    try:
        if os.path.exists(pptx_path) and not os.access(pptx_path, os.W_OK):
            issues.append("文件为只读，无法修改")
    except Exception:
        pass

    # 点2：图表对象完整可显示（图表部件存在、关系可解析、XML 可读）
    try:
        for slide in prs.slides:
            for shape in _iter_shapes(slide.shapes):
                if not getattr(shape, "has_chart", False):
                    continue
                if not _chart_xml(shape):
                    issues.append("存在无法读取的图表对象（可能链接已断开）")
    except Exception as e:
        issues.append("幻灯片内容无法解析（文件可能需要修复）: %s" % e)

    # 点3：页面尺寸 33.87×19.05 cm
    try:
        w, h = _cm(prs.slide_width), _cm(prs.slide_height)
        if not _near(w, PAGE_W_CM) or not _near(h, PAGE_H_CM):
            issues.append("页面尺寸为 %.2f×%.2f cm，不是 33.87×19.05 cm" % (w, h))
    except Exception:
        issues.append("无法读取页面尺寸")

    return (len(issues) == 0), issues


# ============================================================
# 维度二：完成度
# ============================================================


class Evaluator(object):
    """逐条比对评分细则；加分项满足记满分，减分项不满足记扣分。"""

    def __init__(self, prs, pptx_path):
        self.prs = prs
        self.path = pptx_path
        self.parts = _read_parts(pptx_path)
        theme = _theme_fonts(self.parts)
        self.theme_latin, self.theme_ea = theme.get("minorFont", (None, None))
        self.chart_idx, self.chart = _find_chart_shape(prs)
        self.chart_xml = _chart_xml(self.chart) if self.chart is not None else ""
        self.chart_sers = _chart_sers(self.chart_xml) if self.chart_xml else []
        # 原生平滑曲线图：既是真正的 PPT 图表对象，线条又是平滑曲线。
        # 该项不成立 → 图表相关的 7 条一律 0 分（质量项挂在核心项之下）。
        self.chart_native = (bool(self.chart_sers) and "<c:lineChart" in self.chart_xml
                             and _chart_smooth(self.chart_xml))
        self.hits = []
        self.details = {}

    # --------------------------------------------------------
    # 通用工具
    # --------------------------------------------------------
    def page(self, idx):
        """按页内关键字取页；取不到返回 None（页数/页序问题由对应减分项单独扣）。"""
        return _page_by_key(self.prs, PAGE_KEYS[idx])

    def rep(self, label, max_delta, ok, detail=""):
        """记录一条检查项：加分项达标记满分；减分项不达标记扣分（负值）"""
        if max_delta >= 0:
            self.hits.append((label, max_delta if ok else 0))
        else:
            self.hits.append((label, 0 if ok else max_delta))
        self.details[label] = detail

    def guard(self, label, max_delta, precondition, ok, detail=""):
        """质量项：核心项不成立则整项记 0，不做二次判定。"""
        self.rep(label, max_delta, bool(precondition) and bool(ok),
                 detail if precondition else "前置核心项不成立，本项记 0")

    # --------------------------------------------------------
    # 逐条检查（顺序与评分细则一致：先 22 条加分项，再 21 条减分项）
    # --------------------------------------------------------
    def check_all(self):
        self.eval_chart_native()
        self.eval_chart_two_series()
        self.eval_chart_values()
        self.eval_chart_categories()
        self.eval_chart_legend()
        self.eval_chart_on_thanks()
        self.eval_chart_placement()
        self.eval_bansu_title()
        self.eval_bansu_format()
        self.eval_bansu_baseline()
        self.eval_work_rotation()
        self.eval_work_pic_rotation()
        self.eval_work_align()
        self.eval_work_bg_range()
        self.eval_work_bg_pure()
        self.eval_opt_dup()
        self.eval_opt_paras()
        self.eval_opt_format()
        self.eval_plan_mid_removed()
        self.eval_plan_line_removed()
        self.eval_plan_long_body()
        self.eval_plan_long_format()
        self.eval_page_order()
        self.eval_five_pages_intact()
        self.eval_team_keep_body()
        self.eval_theme_master()
        self.eval_backgrounds()
        self.eval_no_annotations()
        self.eval_duty_page()
        self.eval_bansu_box()
        self.eval_work_cards()
        self.eval_work_picture()
        self.eval_work_no_overlap()
        self.eval_review_layout()
        self.eval_review_left()
        self.eval_plan_cards()
        self.eval_plan_bg()
        self.eval_thanks_page()
        self.eval_titles_and_lines()
        self.eval_pictures()
        self.eval_no_empty_objects()
        self.eval_no_animation()
        self.eval_no_extra_text()

    # --------------------------------------------------------
    # 分值：+5
    # 规则原文：文件末尾有原生可编辑的平滑曲线图（真正的 PPT 图表对象 + 平滑曲线）
    # 定位口径：全篇的 p:graphicFrame/c:chart 图表部件
    # 两个条件都满足 → +5；缺任何一个 → 图表相关 7 条一律 0 分
    # --------------------------------------------------------
    def eval_chart_native(self):
        LABEL = "文件末尾有原生可编辑的平滑曲线图"
        self.rep(LABEL, 5, self.chart_native,
                 "图表对象 %s，平滑曲线 %s"
                 % ("有" if self.chart is not None else "无",
                    "是" if self.chart_native else "否"))

    # --------------------------------------------------------
    # 分值：+3
    # 规则原文：图表含"营收数值"与"盈利比值"两个系列，两条线都清晰可见
    # 定位口径：c:ser 的系列名与数值；次坐标轴看 c:axId 是否分两组
    # 全部满足 → +3；被压成贴底直线 → 0
    # --------------------------------------------------------
    def eval_chart_two_series(self):
        LABEL = "图表含营收数值与盈利比值两个系列且两条线都清晰可见"
        ok, detail = False, "图表非原生平滑曲线图"
        if self.chart_native:
            sers = self.chart_sers
            names = [n for n, _v in sers]
            named = all(any(k in n for k in CHART_SER_NAMES) for n in names)
            peaks = [max([abs(x) for x in v] or [0.0]) for _n, v in sers]
            visible = (_chart_has_secondary_axis(self.chart_xml)
                       or (min(peaks) > 0 and max(peaks) / min(peaks) <= 10.0))
            ok = len(sers) == 2 and named and visible
            detail = "系列 %s，次坐标轴 %s" % (names, "有" if
                                              _chart_has_secondary_axis(self.chart_xml) else "无")
        self.guard(LABEL, 3, self.chart_native, ok, detail)

    # --------------------------------------------------------
    # 分值：+3
    # 规则原文：营收 68000/352000/486000/423000/579000（或等值 6.8/35.2/48.6/42.3/57.9 万），
    #          盈利比值 1.42/1.56/1.71/1.63/1.85，无错值、无漏点、无多余点
    # 定位口径：按系列名取 c:val 数值
    # 全部满足 → +3；否则 0
    # --------------------------------------------------------
    def eval_chart_values(self):
        LABEL = "图表数据为指定的营收数值与盈利比值，无错值漏点多余点"
        ok, detail = False, "图表非原生平滑曲线图"
        if self.chart_native:
            got = {}
            for name, vals in self.chart_sers:
                key = "rev" if CHART_SER_NAMES[0] in name else (
                    "ratio" if CHART_SER_NAMES[1] in name else name)
                got[key] = vals
            rev, ratio = got.get("rev", []), got.get("ratio", [])
            ok = ((_values_match(rev, CHART_REVENUE) or _values_match(rev, CHART_REVENUE_ALT))
                  and _values_match(ratio, CHART_RATIO))
            detail = "营收 %s，盈利比值 %s" % (rev, ratio)
        self.guard(LABEL, 3, self.chart_native, ok, detail)

    # --------------------------------------------------------
    # 分值：+1
    # 规则原文：横轴为 1月–5月共 5 个类别且顺序正确
    # 定位口径：c:cat 内的类别文字
    # --------------------------------------------------------
    def eval_chart_categories(self):
        LABEL = "图表横轴为1月至5月共5个类别且顺序正确"
        cats = _chart_cats(self.chart_xml) if self.chart_native else []
        ok = cats == CHART_CATS
        self.guard(LABEL, 1, self.chart_native, ok,
                   "类别 %s" % (cats or "无"))

    # --------------------------------------------------------
    # 分值：+1
    # 规则原文：有图例或系列标识可区分两条线；双轴时两侧刻度能对应到各自系列
    # 定位口径：c:legend 与 c:ser/c:tx；双轴看 c:valAx 是否两个
    # --------------------------------------------------------
    def eval_chart_legend(self):
        LABEL = "图表有图例或系列标识可区分两条线且双轴刻度对应各自系列"
        ok, detail = False, "图表非原生平滑曲线图"
        if self.chart_native:
            legend = "<c:legend>" in self.chart_xml
            named = all(n for n, _v in self.chart_sers)
            val_ax = len(re.findall(r"<c:valAx>", self.chart_xml))
            dual = _chart_has_secondary_axis(self.chart_xml)
            ok = (legend or named) and (not dual or val_ax >= 2)
            detail = "图例 %s，系列名 %s，valAx %d" % ("有" if legend else "无",
                                                     "全有" if named else "缺", val_ax)
        self.guard(LABEL, 1, self.chart_native, ok, detail)

    # --------------------------------------------------------
    # 分值：+1
    # 规则原文：图表落在"感谢聆听"页上（单独占一页或落在其他页面 → 0 分）
    # 定位口径：图表所在页是否含"感谢聆听"
    # --------------------------------------------------------
    def eval_chart_on_thanks(self):
        LABEL = "图表落在感谢聆听页上"
        ok = self.chart_native and _check_chart_on_thanks(self.prs, self.chart)
        self.guard(LABEL, 1, self.chart_native, ok,
                   "图表所在页索引 %d" % self.chart_idx)

    # --------------------------------------------------------
    # 分值：+1
    # 规则原文：图表完整落在页面内，宽高比正常无拉伸变形，且不遮挡该页四处文字
    # 定位口径：图表外框与该页文字框的相交判定
    # --------------------------------------------------------
    def eval_chart_placement(self):
        LABEL = "图表完整落在页面内且不遮挡感谢聆听页文字"
        ok = self.chart_native and _check_chart_geometry(self.prs, self.chart)
        self.guard(LABEL, 1, self.chart_native, ok, "")

    # --------------------------------------------------------
    # 分值：+3
    # 规则原文：岗位职责页第三张卡片（卡底左边缘 21.4 cm、宽 8.3 cm）内有标题"排班岗"
    # 定位口径：第三张卡底范围内、文字为"排班岗"的对象
    # --------------------------------------------------------
    def eval_bansu_title(self):
        LABEL = "岗位职责页第三张卡片内有标题排班岗且文字准确"
        slide = self.page(1)
        ok = slide is not None and _check_bansu_title(slide)
        self.rep(LABEL, 3, ok, "")

    # --------------------------------------------------------
    # 分值：+1
    # 规则原文："排班岗"标题 24pt、加粗、0F5963、中英文字体均为微软雅黑
    # 定位口径：run 级 sz/b/solidFill 与 latin/ea 字体（只设 latin 不得分）
    # 仅在标题存在时评分
    # --------------------------------------------------------
    def eval_bansu_format(self):
        LABEL = "排班岗标题格式与同页前两卡标题一致"
        slide = self.page(1)
        exists = slide is not None and _bansu_title(slide) is not None
        ok = exists and _check_bansu_format(slide, self.theme_latin, self.theme_ea)
        self.guard(LABEL, 1, exists, ok, "")

    # --------------------------------------------------------
    # 分值：+1
    # 规则原文："排班岗"标题与前两卡标题同一基准（顶端 ≤0.3 cm、卡内左偏移与均值差 ≤0.3 cm）
    # 定位口径：标题顶端与"标题左边缘−卡底左边缘"偏移
    # 仅在标题存在时评分
    # --------------------------------------------------------
    def eval_bansu_baseline(self):
        LABEL = "排班岗标题与前两卡标题同一基准且未超出卡片"
        slide = self.page(1)
        exists = slide is not None and _bansu_title(slide) is not None
        ok = exists and _check_bansu_baseline(slide)
        self.guard(LABEL, 1, exists, ok, "")

    # --------------------------------------------------------
    # 分值：+3
    # 规则原文：工作成效页 4 张卡片的卡底、小标题、正文共 12 个对象旋转角全部为 0°
    # 定位口径：a:xfrm@rot（任意一个带旋转即不得分）
    # --------------------------------------------------------
    def eval_work_rotation(self):
        LABEL = "工作成效页4张卡片12个对象旋转角全部为0度"
        slide = self.page(4)
        ok = slide is not None and _check_work_rotation(slide)
        self.rep(LABEL, 3, ok, "")

    # --------------------------------------------------------
    # 分值：+1
    # 规则原文：工作成效页右侧图片旋转角为 0°
    # --------------------------------------------------------
    def eval_work_pic_rotation(self):
        LABEL = "工作成效页右侧图片旋转角为0度"
        slide = self.page(4)
        ok = slide is not None and _check_work_pic_rotation(slide)
        self.rep(LABEL, 1, ok, "")

    # --------------------------------------------------------
    # 分值：+3
    # 规则原文：4 张卡片排列整齐，左右两列卡底左边缘对齐、上下两行卡底顶端对齐，偏差 ≤0.1 cm
    # 定位口径：旋转已归零的前提下量 a:off（旋转不改变坐标）
    # 仅在 12 个对象旋转已归零时评分
    # --------------------------------------------------------
    def eval_work_align(self):
        LABEL = "工作成效页4张卡片两列两行对齐偏差不超过0.1厘米"
        slide = self.page(4)
        zeroed = slide is not None and _check_work_rotation(slide)
        ok = zeroed and _check_work_align(slide)
        self.guard(LABEL, 3, zeroed, ok, "")

    # --------------------------------------------------------
    # 分值：+3
    # 规则原文：工作成效页背景为浅青白，色值落在 EAF8F9–F6FBFC 区间
    # 定位口径：p:cSld/p:bg 的纯色色值
    # --------------------------------------------------------
    def eval_work_bg_range(self):
        LABEL = "工作成效页背景为浅青白且色值在EAF8F9至F6FBFC区间"
        slide = self.page(4)
        ok = slide is not None and _check_work_bg_range(slide)
        kind, value = _slide_bg(slide) if slide is not None else (None, None)
        self.rep(LABEL, 3, ok, "背景 %s %s" % (kind, value))

    # --------------------------------------------------------
    # 分值：+1
    # 规则原文：工作成效页背景为 srgbClr 纯色填充（非渐变、非图片、非主题色变体）
    # 仅在背景为浅青白时评分
    # --------------------------------------------------------
    def eval_work_bg_pure(self):
        LABEL = "工作成效页背景为srgbClr纯色填充"
        slide = self.page(4)
        in_range = slide is not None and _check_work_bg_range(slide)
        ok = in_range and _check_work_bg_pure(slide)
        self.guard(LABEL, 1, in_range, ok, "")

    # --------------------------------------------------------
    # 分值：+3
    # 规则原文：问题复盘页右卡"优化动作"下方正文中，
    #          "将敏感表达改为场景化、说明式表达"只出现 1 次
    # 定位口径：整页段落文字中该句出现次数
    # --------------------------------------------------------
    def eval_opt_dup(self):
        LABEL = "优化动作正文中重复句只出现1次"
        slide = self.page(5)
        ok = slide is not None and _check_opt_dup(slide)
        self.rep(LABEL, 3, ok, "")

    # --------------------------------------------------------
    # 分值：+1
    # 规则原文："优化动作"正文无空段落/空行，共 3 段且顺序为检查表／场景化表达／小结同步
    # 仅在该句只出现 1 次时评分
    # --------------------------------------------------------
    def eval_opt_paras(self):
        LABEL = "优化动作正文无空段落且3段顺序正确"
        slide = self.page(5)
        dedup = slide is not None and _check_opt_dup(slide)
        ok = dedup and _check_opt_paras(slide)
        self.guard(LABEL, 1, dedup, ok, "")

    # --------------------------------------------------------
    # 分值：+1
    # 规则原文："优化动作"3 段格式一致：行距 200%、16pt、203942、中英文微软雅黑、前缀"• "
    # 仅在该句只出现 1 次时评分
    # --------------------------------------------------------
    def eval_opt_format(self):
        LABEL = "优化动作3段格式一致行距200%且16pt微软雅黑"
        slide = self.page(5)
        dedup = slide is not None and _check_opt_dup(slide)
        ok = dedup and _check_opt_format(slide, self.theme_latin, self.theme_ea)
        self.guard(LABEL, 1, dedup, ok, "")

    # --------------------------------------------------------
    # 分值：+3
    # 规则原文：后续规划页只剩 2 张卡（近期、长期）：无"中期"标题与其 3 行正文，无残留空卡底
    # --------------------------------------------------------
    def eval_plan_mid_removed(self):
        LABEL = "后续规划页只剩近期长期2张卡且无残留空卡底"
        slide = self.page(8)
        ok = slide is not None and _check_plan_mid_removed(slide)
        self.rep(LABEL, 3, ok, "")

    # --------------------------------------------------------
    # 分值：+1
    # 规则原文：后续规划页不存在句子"稳住基本盘，再把经验变成流程。"，也没留下空文本框
    # --------------------------------------------------------
    def eval_plan_line_removed(self):
        LABEL = "后续规划页不存在稳住基本盘一句且无空文本框残留"
        slide = self.page(8)
        ok = slide is not None and _check_plan_line_removed(slide)
        self.rep(LABEL, 1, ok, "")

    # --------------------------------------------------------
    # 分值：+3
    # 规则原文："长期"卡正文为 3 个独立段落：搭建标准化复盘范式／强化岗位协同配合效能／
    #          赋能团队稳步长效发展；文字准确无错字、无漏行（行尾多余空格不扣）
    # --------------------------------------------------------
    def eval_plan_long_body(self):
        LABEL = "长期卡正文为3个独立段落且文字准确"
        slide = self.page(8)
        ok = slide is not None and _plan_long_body(slide) is not None
        self.rep(LABEL, 3, ok, "")

    # --------------------------------------------------------
    # 分值：+3
    # 规则原文："长期"卡正文 16pt、行距 150%、203942、中英文微软雅黑、文本框 7.2×4.5 cm，
    #          对齐方式与"近期"卡正文一致（左右位置不查）
    # 仅在"长期"三行已为上述新内容时评分
    # --------------------------------------------------------
    def eval_plan_long_format(self):
        LABEL = "长期卡正文格式为16pt行距150%且文本框7.2乘4.5厘米"
        slide = self.page(8)
        renewed = slide is not None and _plan_long_body(slide) is not None
        ok = renewed and _check_plan_long_format(slide, self.theme_latin, self.theme_ea)
        self.guard(LABEL, 3, renewed, ok, "")

    # --------------------------------------------------------
    # 分值：−5
    # 规则原文：文件为 10 页，页面顺序为封面／岗位职责／协作关系／核心品类／工作成效／
    #          问题复盘／能力成长／团队协作／后续规划／感谢聆听页
    # 页数不是 10 页（含为放图表另起一页）或页序不符 → −5
    # --------------------------------------------------------
    def eval_page_order(self):
        LABEL = "文件为10页且页面顺序正确"
        ok = _check_page_order(self.prs)
        self.rep(LABEL, -5, ok, "共 %d 页" % len(list(self.prs.slides)))

    # --------------------------------------------------------
    # 分值：−5
    # 规则原文：封面、协作关系、核心品类、能力成长、团队协作 5 页须为原状（对象个数、
    #          文字内容、位置尺寸），任一项不符即 −5
    # --------------------------------------------------------
    def eval_five_pages_intact(self):
        LABEL = "封面协作关系核心品类能力成长团队协作5页保持原状"
        pairs = ((0, _check_cover), (2, _check_collab), (3, _check_category),
                 (6, _check_growth), (7, _check_team))
        bad = []
        for idx, fn in pairs:
            slide = self.page(idx)
            if slide is None or not fn(slide):
                bad.append(PAGE_KEYS[idx])
        self.rep(LABEL, -5, not bad, "不符页：%s" % (bad or "无"))

    # --------------------------------------------------------
    # 分值：−3
    # 规则原文：团队协作页"重复沉淀"卡内正文在 (16.5, 9.4)、14.1×1.6，为 2 段，
    #          逐字为两句"…；…"重复写法（须原样保留）
    # --------------------------------------------------------
    def eval_team_keep_body(self):
        LABEL = "团队协作页重复沉淀卡内两句重复写法原样保留"
        slide = self.page(7)
        ok = slide is not None and _check_team_keep_body(slide)
        self.rep(LABEL, -3, ok, "")

    # --------------------------------------------------------
    # 分值：−5
    # 规则原文：主题 majorFont/minorFont 的 latin 均为 Microsoft YaHei、ea 均为空；
    #          只有 1 个 slideMaster、1 个 slideLayout；全篇正文字体为微软雅黑
    # --------------------------------------------------------
    def eval_theme_master(self):
        LABEL = "主题字体与母版版式保持原状且全篇正文为微软雅黑"
        ok = _check_theme_and_fonts(self.prs, self.parts)
        self.rep(LABEL, -5, ok, "主题 %s" % (_theme_fonts(self.parts) or "无"))

    # --------------------------------------------------------
    # 分值：−3
    # 规则原文：各页背景为纯色且色值依次为规定值；页面标题色 17323A、
    #          标题下装饰短线色 29A7A5、右下角装饰横线色 B9E1E4
    # --------------------------------------------------------
    def eval_backgrounds(self):
        LABEL = "各页背景色与标题装饰线色值保持原状"
        ok = _check_backgrounds(self.prs)
        self.rep(LABEL, -3, ok, "")

    # --------------------------------------------------------
    # 分值：−3
    # 规则原文：文件中不存在红圈/黄圈/箭头等标注图形、批注、荧光笔标记，
    #          以及"已改""待确认"等临时文字
    # --------------------------------------------------------
    def eval_no_annotations(self):
        LABEL = "无标注图形批注荧光笔标记及临时文字"
        ok = _check_no_annotations(self.prs, self.parts)
        self.rep(LABEL, -3, ok, "")

    # --------------------------------------------------------
    # 分值：−3
    # 规则原文：岗位职责页除"排班岗"标题外为 15 个对象，三张卡底、两个箭头、副标题、
    #          底部标语、卡 1/卡 2 标题与正文、第三卡正文逐字均须在位
    # --------------------------------------------------------
    def eval_duty_page(self):
        LABEL = "岗位职责页除排班岗外15个对象及其位置尺寸不变"
        slide = self.page(1)
        ok = slide is not None and _check_duty(slide)
        self.rep(LABEL, -3, ok, "")

    # --------------------------------------------------------
    # 分值：−1
    # 规则原文："排班岗"文本框无边框、无填充、无阴影，文字单行不折行，
    #          且与第三卡正文不重叠、不遮挡正文、不溢出卡片边界
    # --------------------------------------------------------
    def eval_bansu_box(self):
        LABEL = "排班岗文本框无边框填充阴影且不折行不重叠不溢出"
        slide = self.page(1)
        ok = slide is None or _check_bansu_box(slide)
        self.rep(LABEL, -1, ok, "")

    # --------------------------------------------------------
    # 分值：−3
    # 规则原文：工作成效页 4 张卡片各由卡底 7.0×3.4 + 小标题（14pt、0F5963）+
    #          正文（11.2pt、203942）三件组成，四组文字不变；卡底圆角、F9FEFE、带阴影；
    #          4 张卡"标题相对卡底""正文相对卡底"的偏移量彼此偏差 ≤0.3 cm
    # --------------------------------------------------------
    def eval_work_cards(self):
        LABEL = "工作成效页4张卡片三件套构成文字与卡内偏移不变"
        slide = self.page(4)
        ok = slide is not None and _check_work_cards(slide)
        self.rep(LABEL, -3, ok, "")

    # --------------------------------------------------------
    # 分值：−3
    # 规则原文：工作成效页有 1 张图片，画面完整、未被裁剪或替换，宽高比失真不超过 2%
    # --------------------------------------------------------
    def eval_work_picture(self):
        LABEL = "工作成效页图片完整未裁剪替换且宽高比失真不超过2%"
        slide = self.page(4)
        ok = slide is not None and _check_work_picture(slide)
        self.rep(LABEL, -3, ok, "")

    # --------------------------------------------------------
    # 分值：−1
    # 规则原文：工作成效页所有对象完整落在页面内，卡片之间、图片与卡片之间无重叠
    # 仅在该页已摆正的情况下适用
    # --------------------------------------------------------
    def eval_work_no_overlap(self):
        LABEL = "工作成效页对象不越界且卡片与图片互不重叠"
        slide = self.page(4)
        zeroed = slide is not None and _check_work_rotation(slide)
        ok = (not zeroed) or _check_work_no_overlap(slide)
        self.rep(LABEL, -1, ok, "" if zeroed else "该页未摆正，本条不适用")

    # --------------------------------------------------------
    # 分值：−3
    # 规则原文：问题复盘页有青色三角形 1 个、1.4×1.4、在 (16.2, 7.6)、旋转 90°；
    #          两张卡底各 12.6×10.3、顶端 3.7、左边缘 2.3/18.9；
    #          底部标语"把一次问题变成一套流程"（28pt、0F5963）在 (11.0, 15.2)
    # --------------------------------------------------------
    def eval_review_layout(self):
        LABEL = "问题复盘页三角形两张卡底与底部标语在位"
        slide = self.page(5)
        ok = slide is not None and _check_review_layout(slide)
        self.rep(LABEL, -3, ok, "")

    # --------------------------------------------------------
    # 分值：−3
    # 规则原文：问题复盘页左卡标题"风险表现"（32pt、A85E1C）在 (2.9, 4.2)，
    #          其下方正文在 (2.9, 5.3)、共 3 段、行距 200%、16pt、203942、前缀"• "
    # --------------------------------------------------------
    def eval_review_left(self):
        LABEL = "问题复盘页左卡风险表现标题与3段正文不变"
        slide = self.page(5)
        ok = slide is not None and _check_review_left(slide)
        self.rep(LABEL, -3, ok, "")

    # --------------------------------------------------------
    # 分值：−3
    # 规则原文：后续规划页"近期"卡完整（卡底 8.5×6.5、顶端 5.3，标题"近期"，正文 3 段
    #          16pt、行距 150%、203942）；"长期"卡卡底 8.5×6.5、顶端 5.3、标题"长期"；
    #          两卡均完整落在页面内（左右位置不查）
    # --------------------------------------------------------
    def eval_plan_cards(self):
        LABEL = "后续规划页近期卡完整且长期卡卡底与标题在位"
        slide = self.page(8)
        ok = slide is not None and _check_plan_cards(slide)
        self.rep(LABEL, -3, ok, "")

    # --------------------------------------------------------
    # 分值：−3
    # 规则原文：后续规划页通栏背景图在 (0, 2.2)、尺寸 33.9×15.2（允差 ±0.1 cm），画面完整
    # --------------------------------------------------------
    def eval_plan_bg(self):
        LABEL = "后续规划页通栏背景图位置尺寸不变且画面完整"
        slide = self.page(8)
        ok = slide is not None and _check_plan_bg(slide)
        self.rep(LABEL, -3, ok, "")

    # --------------------------------------------------------
    # 分值：−3
    # 规则原文：感谢聆听页原有 7 个对象齐全且在位（右侧图片、左侧底板、主标题、
    #          组名、姓名、装饰短线、寄语）；新增的图表及承载对象不计入对象个数
    # --------------------------------------------------------
    def eval_thanks_page(self):
        LABEL = "感谢聆听页原有7个对象齐全且在位"
        slide = self.page(9)
        ok = slide is not None and _check_thanks(slide)
        self.rep(LABEL, -3, ok, "")

    # --------------------------------------------------------
    # 分值：−3
    # 规则原文：岗位职责页至后续规划页每页有页面标题（24pt、17323A）在 (1.8, 1.1)、21.3×1.1
    #          及标题下青色短线在 (1.8, 2.6)、长 2.9；封面至后续规划页每页有右下角
    #          浅色装饰横线在 (27.1, 18.1)、长 4.8
    # --------------------------------------------------------
    def eval_titles_and_lines(self):
        LABEL = "各页页面标题与装饰短线横线齐全且位置不变"
        ok = _check_titles_and_lines(self.prs)
        self.rep(LABEL, -3, ok, "")

    # --------------------------------------------------------
    # 分值：−3
    # 规则原文：封面、协作关系、核心品类、团队协作、感谢聆听 5 页各有 1 张图片，
    #          位置尺寸分别为规定值，任一张缺失或不符即扣分
    # --------------------------------------------------------
    def eval_pictures(self):
        LABEL = "5页配图各1张且位置尺寸不变"
        ok = _check_pictures(self.prs)
        self.rep(LABEL, -3, ok, "")

    # --------------------------------------------------------
    # 分值：−1
    # 规则原文：除封面页 (1.8, 17.9)、14.0×0.5 处的空文本框外，不存在空占位符、
    #          空文本框，以及无填充、无线条、无文字且不起装饰作用的多余对象
    # --------------------------------------------------------
    def eval_no_empty_objects(self):
        LABEL = "无空占位符空文本框及多余空对象"
        ok = _check_no_empty_objects(self.prs)
        self.rep(LABEL, -1, ok, "")

    # --------------------------------------------------------
    # 分值：−1
    # 规则原文：全篇无动画、无页面切换效果
    # --------------------------------------------------------
    def eval_no_animation(self):
        LABEL = "全篇无动画无页面切换效果"
        ok = _check_no_animation(self.parts)
        self.rep(LABEL, -1, ok, "")

    # --------------------------------------------------------
    # 分值：−1
    # 规则原文：不存在新增的说明性文字（如"数据来源：…"），也无水印、页码、Logo
    # --------------------------------------------------------
    def eval_no_extra_text(self):
        LABEL = "无新增说明性文字水印页码Logo"
        ok = _check_no_extra_text(self.prs)
        self.rep(LABEL, -1, ok, "")














# ============================================================
# 各检查项满分（加分项 22 条 / 46 分；减分项 21 条 / 最多 −59 分）
#   加分项：达标记满分，hit = (delta == max_delta)
#   减分项：达标记 0、不达标记负值，hit = (delta == 0)
#   实得分 = Σ加分项 + Σ减分项，下限 0，上限 46；完成度 = 实得分 ÷ 46
# ============================================================

RULE_MAX_DELTA = {
    "文件末尾有原生可编辑的平滑曲线图": 5,
    "图表含营收数值与盈利比值两个系列且两条线都清晰可见": 3,
    "图表数据为指定的营收数值与盈利比值，无错值漏点多余点": 3,
    "图表横轴为1月至5月共5个类别且顺序正确": 1,
    "图表有图例或系列标识可区分两条线且双轴刻度对应各自系列": 1,
    "图表落在感谢聆听页上": 1,
    "图表完整落在页面内且不遮挡感谢聆听页文字": 1,
    "岗位职责页第三张卡片内有标题排班岗且文字准确": 3,
    "排班岗标题格式与同页前两卡标题一致": 1,
    "排班岗标题与前两卡标题同一基准且未超出卡片": 1,
    "工作成效页4张卡片12个对象旋转角全部为0度": 3,
    "工作成效页右侧图片旋转角为0度": 1,
    "工作成效页4张卡片两列两行对齐偏差不超过0.1厘米": 3,
    "工作成效页背景为浅青白且色值在EAF8F9至F6FBFC区间": 3,
    "工作成效页背景为srgbClr纯色填充": 1,
    "优化动作正文中重复句只出现1次": 3,
    "优化动作正文无空段落且3段顺序正确": 1,
    "优化动作3段格式一致行距200%且16pt微软雅黑": 1,
    "后续规划页只剩近期长期2张卡且无残留空卡底": 3,
    "后续规划页不存在稳住基本盘一句且无空文本框残留": 1,
    "长期卡正文为3个独立段落且文字准确": 3,
    "长期卡正文格式为16pt行距150%且文本框7.2乘4.5厘米": 3,
    "文件为10页且页面顺序正确": -5,
    "封面协作关系核心品类能力成长团队协作5页保持原状": -5,
    "团队协作页重复沉淀卡内两句重复写法原样保留": -3,
    "主题字体与母版版式保持原状且全篇正文为微软雅黑": -5,
    "各页背景色与标题装饰线色值保持原状": -3,
    "无标注图形批注荧光笔标记及临时文字": -3,
    "岗位职责页除排班岗外15个对象及其位置尺寸不变": -3,
    "排班岗文本框无边框填充阴影且不折行不重叠不溢出": -1,
    "工作成效页4张卡片三件套构成文字与卡内偏移不变": -3,
    "工作成效页图片完整未裁剪替换且宽高比失真不超过2%": -3,
    "工作成效页对象不越界且卡片与图片互不重叠": -1,
    "问题复盘页三角形两张卡底与底部标语在位": -3,
    "问题复盘页左卡风险表现标题与3段正文不变": -3,
    "后续规划页近期卡完整且长期卡卡底与标题在位": -3,
    "后续规划页通栏背景图位置尺寸不变且画面完整": -3,
    "感谢聆听页原有7个对象齐全且在位": -3,
    "各页页面标题与装饰短线横线齐全且位置不变": -3,
    "5页配图各1张且位置尺寸不变": -3,
    "无空占位符空文本框及多余空对象": -1,
    "全篇无动画无页面切换效果": -1,
    "无新增说明性文字水印页码Logo": -1,
}

assert sum(v for v in RULE_MAX_DELTA.values() if v > 0) == MAX_SCORE_TOTAL
assert sum(v for v in RULE_MAX_DELTA.values() if v < 0) == MAX_PENALTY_TOTAL

def _find_pptx_in_dir(dir_path):
    """在目录下查找待评估的 .pptx（跳过 PowerPoint 临时文件）"""
    try:
        names = os.listdir(dir_path)
    except Exception:
        return None
    cands = [n for n in names
             if n.lower().endswith(".pptx") and not n.startswith("~$")]
    cands.sort()
    for n in cands:
        full = os.path.join(dir_path, n)
        if os.path.isfile(full):
            return full
    return None


def evaluate(dir_path: str) -> dict:
    result = {
        "id": SCRIPT_ID,
        "file_name": "",
        "status": "ok",
        "error": None,
        "dim1_pass": False,
        "dim1_reason": "",
        "dim2_items": [],
        "total_score": 0,
        "max_score": sum(v for v in RULE_MAX_DELTA.values() if v > 0),
    }

    try:
        if not dir_path or not os.path.isdir(dir_path):
            raise FileNotFoundError("目录不存在：%s" % dir_path)
        pptx_path = _find_pptx_in_dir(dir_path)
        if not pptx_path:
            raise FileNotFoundError("目录内未找到 .pptx 文件：%s" % dir_path)
        result["file_name"] = os.path.basename(pptx_path)

        prs = _open_presentation(pptx_path)

        ok, issues = check_dimension_one(prs, pptx_path)
        result["dim1_pass"] = ok
        # dim1_reason 仅在维度一未通过时给出原因摘要；通过则为空
        result["dim1_reason"] = "" if ok else "; ".join(issues)
        if not ok:
            # 维度一未通过：完成度直接判 0，不再检查维度二
            return result

        ev = Evaluator(prs, pptx_path)
        ev.check_all()

        total = 0
        for label, delta in ev.hits:
            max_delta = RULE_MAX_DELTA.get(label, 0)
            total += delta
            result["dim2_items"].append({
                "rule": label,
                "max_delta": max_delta,
                "delta": delta,
                # 加分项达标为命中；减分项未被扣分为命中
                "hit": (delta == max_delta) if max_delta > 0 else (delta == 0),
            })
        # 实得分下限 0、上限 46
        result["total_score"] = max(0, min(total, MAX_SCORE_TOTAL))
    except Exception as exc:
        # status 为 error 时未能完成判定，dim1_pass 留空，dim1_reason 也为空
        result["status"] = "error"
        result["error"] = "%s: %s" % (type(exc).__name__, exc)
        result["dim1_pass"] = ""
        result["dim1_reason"] = ""
        result["dim2_items"] = []
        result["total_score"] = 0
        return result

    return result


if __name__ == "__main__":
    _dir = sys.argv[1] if len(sys.argv) > 1 else os.path.dirname(os.path.abspath(__file__))
    _payload = json.dumps(evaluate(_dir), ensure_ascii=False, indent=2)
    try:
        sys.stdout.buffer.write((_payload + "\n").encode("utf-8"))
    except AttributeError:
        print(_payload)




































