#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# pyright: reportUnknownMemberType=false, reportUnknownVariableType=false
# pyright: reportUnknownArgumentType=false, reportMissingImports=false
# pyright: reportOptionalMemberAccess=false, reportGeneralTypeIssues=false
"""
艺术漆北区销售额提升方案 PPT 完成度校验器。

说明：按接口统一约定，本模块不修改全局 sys.stdout，也不 print 主结果。
主结果统一由 evaluate(dir_path) -> dict 返回；入参为"脚本所在目录的路径"，
函数在该目录内自行定位待评测的 .pptx 文件。__main__ 仅用于本地自测输出。
"""

import os
import re
import sys
import json
import math
import zipfile
import colorsys
import difflib
import traceback

try:
    from pptx import Presentation
except Exception:  # pragma: no cover - 依赖缺失时由 evaluate 统一兜底
    Presentation = None

SCRIPT_ID = "076"

# 维度2 满分 = 20 条加分项之和；减分项最好情况为 0，因此不计入满分。
MAX_SCORE_TOTAL = 72
MIN_SCORE_TOTAL = -27

# 与 _evaluate_impl 中 results.append 的顺序严格一一对应。
# 正数 = 加分项（命中得满分，未命中得 0）；负数 = 减分项（命中即扣分，保持原样为 0）。
_MAX_DELTAS_ORDER = [
    5,   # 01 拆分为30页且顺序为 首页/目录/9标题页/18正文页/尾页
    5,   # 02 九个一级标题各自单独成标题页
    5,   # 03 18个二级标题各自单独成页且同页放置对应正文
    5,   # 04 全文中文文字字体均为宋体
    5,   # 05 所有页面背景均为绿色渐变且由浅到深
    5,   # 06 18个正文页正文外围有橙色矩形边框且正文在框内
    5,   # 07 二级标题与正文各有进入动画且效果不同
    5,   # 08 首个动画单击触发、其后动画自动承接
    3,   # 09 首页标题文字为"艺术漆北区销售额提升方案"
    3,   # 10 首页与9个一级标题页标题加粗且单行不换行
    3,   # 11 首页标题横向宽度占页面宽度四分之三以上
    3,   # 12 目录页存在且完整列出九个一级标题
    3,   # 13 绿色渐变整体明度偏浅，最深处为中浅绿
    3,   # 14 橙色矩形边框线宽不小于2磅
    3,   # 15 18个正文页二级标题位于页面最上方且靠左
    3,   # 16 18个正文页正文与素材内容一致
    3,   # 17 结尾板块内容放置在最后一页
    3,   # 18 尾页沿用深蓝色底、金色笔触画面的整版版式
    1,   # 19 同类页面的字号与加粗设置彼此一致
    1,   # 20 无空白页、无空文本框、正文无多余空行与行首空格
    -5,  # 21 幻灯片母版与版式沿用既有版面
    -3,  # 22 右下角 AkzoNobel 标识与"内部培训请勿传播"及图片素材齐全
    -3,  # 23 页面尺寸为 12192000 x 6858000 EMU
    -3,  # 24 除绿色渐变外没有其他背景装饰
    -3,  # 25 橙色框为无填充直角矩形
    -3,  # 26 一级标题页只显示该一级标题文字
    -1,  # 27 主题字体方案的东亚字体为宋体
    -1,  # 28 首页没有横条色块
    -1,  # 29 首页除标题外没有其他文字
    -1,  # 30 目录页只放置"目录"标题与九个条目
    -1,  # 31 目录页条目文字保留顿号
    -1,  # 32 正文页最上方靠左只显示二级标题
    -1,  # 33 正文页没有额外标注
]
assert sum(d for d in _MAX_DELTAS_ORDER if d > 0) == MAX_SCORE_TOTAL
assert sum(d for d in _MAX_DELTAS_ORDER if d < 0) == MIN_SCORE_TOTAL

# ---------------------------------------------------------------- 素材文本常量
_COVER_TITLE = "艺术漆北区销售额提升方案"

_L1_TITLES = [
    "一、市场趋势与消费升级",
    "二、重点城市增长机会",
    "三、行业竞争与品牌机遇",
    "四、年度发展目标",
    "五、核心产品体系",
    "六、产品组合与差异化卖点",
    "七、多渠道拓展策略",
    "八、整合营销与本地转化",
    "九、服务体系与执行保障",
]

_L2_TITLES = [
    "1.1 市场扩容趋势",
    "1.2 消费需求变化",
    "2.1 改善型住房需求释放",
    "2.2 存量房翻新潜力",
    "3.1 竞争格局变化",
    "3.2 品牌突破方向",
    "4.1 经营增长目标",
    "4.2 品牌认知目标",
    "5.1 主推系列定位",
    "5.2 消费决策简化",
    "6.1 产品矩阵设计",
    "6.2 核心卖点提炼",
    "7.1 零售终端升级",
    "7.2 家装与工程渠道拓展",
    "8.1 线上内容种草",
    "8.2 线下活动转化",
    "9.1 服务体系升级",
    "9.2 经销商赋能与组织保障",
]

# 版面自带的附属文字：保持原样出现不视为"新增文字"。
_TEMPLATE_KEEP_TEXTS = (
    "艺术漆产品资料",
    "质感墙面 · 色彩灵感 · 空间焕新",
    "鎏金天鹅绒艺术漆",
    "建议零售价",
    "内部培训请勿传播",
)

_CLOSING_WORDS = ("谢谢", "感谢", "thanks", "thank you")

_L2_BODIES = {
    "1.1 市场扩容趋势":
        "2026年，艺术涂料市场仍处于快速扩容阶段。随着居民家装需求从“基础居住”转向“个"
        "性表达、质感升级和环保安心”，艺术漆在内墙涂料中的渗透率仍有较大提升空间。",
    "1.2 消费需求变化":
        "年轻消费群体逐渐成为家装决策主力，他们更关注空间氛围、色彩搭配、材质表现和整体设"
        "计感。这一变化推动艺术漆从小众装饰材料逐步走向主流家装品类，也为墙面美学消费打开"
        "了更大的增长空间。",
    "2.1 改善型住房需求释放":
        "华北重点城市的市场机会主要来自改善型住房需求和存量房翻新需求。一线核心城市在楼市"
        "政策调整后，二手房交易和改善型装修需求有所释放，城市更新项目也带动了墙面翻新、精"
        "装升级和局部改造需求。",
    "2.2 存量房翻新潜力":
        "环渤海重点城市在老旧小区改造、存量住宅更新和家庭局部焕新方面具备较大潜力。家居建"
        "材行业的增长重心正在从新房增量逐步转向存量房精细化改造，艺术漆可借此切入局部焕新"
        "、整屋翻新和品质升级等消费场景。",
    "3.1 竞争格局变化":
        "艺术漆赛道已进入品牌集中发力阶段。行业头部企业正在通过环保体系、产品质感、终端体"
        "验和设计师渠道抢占消费者心智，部分品牌在零售端和工程端已形成一定先发优势。",
    "3.2 品牌突破方向":
        "涂料行业受原材料、环保标准和产品升级影响，价格体系可能继续调整。对于“澜庭漆艺”"
        "而言，应借助环保标准升级和审美消费提升的窗口期，强化其在环保、美学和整体空间解决"
        "方案上的差异化优势。",
    "4.1 经营增长目标":
        "本方案建议将2026年北区艺术漆业务目标设定为销售额同比增长不低于20%，重点城市市场"
        "份额提升3—5个百分点，新增艺术漆专卖店或店中店约30家，核心经销商数量提升约20%。",
    "4.2 品牌认知目标":
        "品牌层面应集中塑造“澜庭漆艺=墙面美学解决方案专家”的认知，在重点城市实现目标客"
        "群认知度明显提升，并通过终端体验、设计师推荐和用户口碑形成持续转化。",
    "5.1 主推系列定位":
        "核心策略上，应以“艺境家”艺术漆系列为主引擎，以“色感共生”搭配体系作为消费沟通"
        "纽带。该体系可围绕不同生活氛围建立色盘组合，将乳胶漆色彩、艺术漆质感和空间风格整"
        "合在同一套选择逻辑中。",
    "5.2 消费决策简化":
        "相比单纯销售单品，更应向用户传递“已经搭配好的整体墙面方案”。通过统一色彩、质感"
        "和风格表达，降低消费者在艺术漆选购过程中的搭配难度，让用户像选择成套软装或妆容方"
        "案一样选择墙面效果。",
    "6.1 产品矩阵设计":
        "产品策略应形成高端引流、中端走量和工程适配三类组合。高端系列可主打通透光感、细腻"
        "肌理和静奢风格，用于旗舰店展示和高预算客户转化；中端系列则围绕不同氛围色盘，提供"
        "多价位、多风格的艺术漆组合；工程系列则面向精装房、旧改项目和商业空间，推出更适合"
        "批量施工、成本控制和效果稳定的产品包装。",
    "6.2 核心卖点提炼":
        "在产品传播中，应重点突出环保领先、色质统一、专业美学支持和施工标准化等核心卖点。"
        "通过“环保安心+美学搭配+稳定施工”的组合表达，建立产品与普通墙面材料之间的差异。",
    "7.1 零售终端升级":
        "渠道策略应围绕零售终端、家装公司、工长群体和工程项目四条路径展开。重点城市旗舰店"
        "可进行“艺境家”主题化升级，设置沉浸式样板区和色盘体验区，让消费者直观看到不同墙"
        "面效果在真实空间中的呈现；建材市场可布局店中店或艺术漆专区，配套样板、选色设备和"
        "导购培训。",
    "7.2 家装与工程渠道拓展":
        "家装公司与设计师渠道可通过战略合作、设计沙龙、样板间植入和推荐激励提升转化效率；"
        "工长渠道可通过培训认证、标杆评选和施工支持，稳定最终落地效果；工程渠道则重点跟进"
        "城市更新、精装房和商业空间项目，形成零售与项目双轮驱动。",
    "8.1 线上内容种草":
        "营销层面应以内容种草、本地化投放和城市主题活动为主线。线上可在生活方式内容平台、"
        "短视频平台和私域社群持续输出艺术漆案例、施工过程、前后对比和搭配攻略，形成从认知"
        "到咨询再到到店体验的闭环。",
    "8.2 线下活动转化":
        "线下可围绕核心城市打造“墙面美学季”“旧房焕新节”“双城联动年终礼”等主题活动。"
        "促销方式上，可结合老带新、小区团购、完工晒家评选、设计师联合案例等方式提升裂变传"
        "播和本地转化效率。",
    "9.1 服务体系升级":
        "服务与执行保障是方案落地的关键。建议推出“艺术漆焕新”专属服务包，涵盖免费上门勘"
        "测、数字化选色、专业施工、质量验收和售后回访，并针对存量房用户提供“住着改”、局"
        "部翻新、旧墙快速评估和样板小区服务。",
    "9.2 经销商赋能与组织保障":
        "经销商端应提供产品培训、销售培训、线上获客培训、样品支持、门店补贴和销售激励。执"
        "行上可成立北区“艺境家”专项工作小组，由销售、市场、渠道、培训和工程团队协同推进"
        "，按月复盘门店升级、渠道拓展、活动转化和销售达成情况，确保全年目标稳步落地。",
}
assert len(_L2_BODIES) == 18 and all(t in _L2_BODIES for t in _L2_TITLES)

# ------------------------------------------------------------------ 基础工具层
_NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main"
_NS_P = "http://schemas.openxmlformats.org/presentationml/2006/main"

_EMU_CM = 360000.0          # 1 厘米 = 360000 EMU
_EMU_PT = 12700.0           # 1 磅 = 12700 EMU
_STD_W = 12192000           # 标准 16:9 页宽（33.87 厘米）
_STD_H = 6858000            # 标准 16:9 页高（19.05 厘米）
_MIN_LINE_W = 2 * _EMU_PT   # 2 磅线宽

_SONGTI = ("宋体", "SimSun", "NSimSun", "Songti", "STSong", "SimSun-ExtB")


def _a(tag):
    return "{%s}%s" % (_NS_A, tag)


def _p(tag):
    return "{%s}%s" % (_NS_P, tag)


def _norm(text):
    """去掉所有空白（含全角空格）后的文本，用于结构匹配。"""
    if not text:
        return ""
    return re.sub(r"[\s　\xa0]+", "", text)


def _norm_cmp(text):
    """用于正文一致性比对：统一空白、去掉零宽字符，保留标点。"""
    if not text:
        return ""
    text = text.replace("　", "").replace("\xa0", "")
    text = re.sub(r"[​-\u200F﻿]", "", text)
    return re.sub(r"\s+", "", text)


def _is_cjk(ch):
    o = ord(ch)
    return (0x3400 <= o <= 0x4DBF or 0x4E00 <= o <= 0x9FFF
            or 0xF900 <= o <= 0xFAFF or 0x3000 <= o <= 0x303F
            or 0xFF00 <= o <= 0xFF65)


def _has_cjk(text):
    return any(_is_cjk(c) for c in text or "")


def _theme_colors(file_path):
    """从 ppt/theme/theme1.xml 读取主题配色，返回 {'accent3': (r,g,b), ...}。"""
    out = {}
    try:
        with zipfile.ZipFile(file_path) as zf:
            xml = zf.read("ppt/theme/theme1.xml").decode("utf-8", "ignore")
    except Exception:
        return out
    m = re.search(r"<a:clrScheme.*?</a:clrScheme>", xml, re.S)
    if not m:
        return out
    for mm in re.finditer(r"<a:(\w+)>\s*<a:(?:srgbClr val=\"([0-9A-Fa-f]{6})\""
                          r"|sysClr[^>]*?lastClr=\"([0-9A-Fa-f]{6})\")", m.group(0)):
        hexv = mm.group(2) or mm.group(3)
        out[mm.group(1)] = _hex_rgb(hexv)
    # 常见 clrMap 别名，便于 schemeClr val="bg1"/"tx1" 直接取色
    for alias, real in (("bg1", "lt1"), ("tx1", "dk1"), ("bg2", "lt2"), ("tx2", "dk2")):
        if real in out:
            out.setdefault(alias, out[real])
    return out


def _theme_fonts(file_path):
    """返回 [(kind, latin, ea, hans), ...]，kind 为 major/minor。"""
    out = []
    try:
        with zipfile.ZipFile(file_path) as zf:
            xml = zf.read("ppt/theme/theme1.xml").decode("utf-8", "ignore")
    except Exception:
        return out
    for m in re.finditer(r"<a:(major|minor)Font>(.*?)</a:\1Font>", xml, re.S):
        body = m.group(2)
        lat = re.search(r"<a:latin[^>]*typeface=\"([^\"]*)\"", body)
        ea = re.search(r"<a:ea[^>]*typeface=\"([^\"]*)\"", body)
        hans = re.search(r"<a:font script=\"Hans\" typeface=\"([^\"]*)\"", body)
        out.append((m.group(1),
                    lat.group(1) if lat else "",
                    ea.group(1) if ea else "",
                    hans.group(1) if hans else ""))
    return out


def _hex_rgb(hexv):
    hexv = (hexv or "").strip().lstrip("#")
    if len(hexv) != 6:
        return None
    try:
        return (int(hexv[0:2], 16), int(hexv[2:4], 16), int(hexv[4:6], 16))
    except ValueError:
        return None


def _pct(el, tag):
    """读取 a:lumMod/a:lumOff/a:shade/a:tint 等百分比子元素，返回 0~1 或 None。"""
    child = el.find(_a(tag))
    if child is None:
        return None
    val = child.get("val")
    if val is None:
        return None
    try:
        return float(val) / 100000.0
    except ValueError:
        return None


def _resolve_color(el, theme):
    """把 a:srgbClr / a:schemeClr / a:sysClr（含亮度修饰）解析为 (r,g,b)。"""
    if el is None:
        return None
    tag = el.tag.split("}")[-1]
    if tag == "srgbClr":
        rgb = _hex_rgb(el.get("val"))
    elif tag == "sysClr":
        rgb = _hex_rgb(el.get("lastClr"))
    elif tag == "schemeClr":
        rgb = theme.get(el.get("val") or "")
    else:
        return None
    if rgb is None:
        return None
    r, g, b = [c / 255.0 for c in rgb]
    h, l, s = colorsys.rgb_to_hls(r, g, b)
    lum_mod = _pct(el, "lumMod")
    lum_off = _pct(el, "lumOff")
    if lum_mod is not None:
        l *= lum_mod
    if lum_off is not None:
        l += lum_off
    sat_mod = _pct(el, "satMod")
    if sat_mod is not None:
        s = min(1.0, s * sat_mod)
    l = min(1.0, max(0.0, l))
    r, g, b = colorsys.hls_to_rgb(h, l, s)
    shade = _pct(el, "shade")
    if shade is not None:
        r, g, b = r * shade, g * shade, b * shade
    tint = _pct(el, "tint")
    if tint is not None:
        r = r * tint + (1.0 - tint)
        g = g * tint + (1.0 - tint)
        b = b * tint + (1.0 - tint)
    return (int(round(r * 255)), int(round(g * 255)), int(round(b * 255)))


def _hls(rgb):
    """返回 (色相角度 0~360, 明度 0~1, 饱和度 0~1)。"""
    if not rgb:
        return None
    h, l, s = colorsys.rgb_to_hls(rgb[0] / 255.0, rgb[1] / 255.0, rgb[2] / 255.0)
    return (h * 360.0, l, s)


def _is_green(rgb):
    hls = _hls(rgb)
    if not hls:
        return False
    h, l, s = hls
    return 55.0 <= h <= 175.0 and s >= 0.10 and 0.10 <= l <= 0.99


def _is_orange(rgb):
    hls = _hls(rgb)
    if not hls:
        return False
    h, l, s = hls
    return 12.0 <= h <= 48.0 and s >= 0.35 and 0.25 <= l <= 0.80


def _iter_shapes(shapes):
    """展开组合形状，逐个返回底层形状。"""
    for sh in shapes:
        yield sh
        sub = getattr(sh, "shapes", None)
        if sub is not None:
            for inner in _iter_shapes(sub):
                yield inner


def _box(sh):
    """返回 (left, top, right, bottom)，单位 EMU。"""
    l = int(sh.left or 0)
    t = int(sh.top or 0)
    w = int(sh.width or 0)
    h = int(sh.height or 0)
    return (l, t, l + w, t + h)


def _sp_pr(sh):
    el = getattr(sh, "_element", None)
    if el is None:
        return None
    for tag in ("spPr", "grpSpPr"):
        found = el.find(_p(tag))
        if found is not None:
            return found
    return None


def _geom(sh):
    """返回预设几何名称，如 rect / roundRect / ellipse；无预设几何返回 None。"""
    sp_pr = _sp_pr(sh)
    if sp_pr is None:
        return None
    prst = sp_pr.find(_a("prstGeom"))
    return prst.get("prst") if prst is not None else None


def _fill_kind(sh):
    """返回形状自身填充类型（noFill/solidFill/gradFill/...），未显式设置返回 None。"""
    sp_pr = _sp_pr(sh)
    if sp_pr is None:
        return None
    for child in sp_pr:
        tag = child.tag.split("}")[-1]
        if tag.endswith("Fill"):
            return tag
    return None


def _fill_color(sh, theme):
    sp_pr = _sp_pr(sh)
    if sp_pr is None:
        return None
    solid = sp_pr.find(_a("solidFill"))
    if solid is None:
        return None
    for child in solid:
        rgb = _resolve_color(child, theme)
        if rgb:
            return rgb
    return None


def _line_info(sh, theme):
    """返回 (线宽EMU 或 None, 线色 rgb 或 None, 是否显式 noLine)。"""
    sp_pr = _sp_pr(sh)
    if sp_pr is None:
        return (None, None, False)
    ln = sp_pr.find(_a("ln"))
    if ln is None:
        return (None, None, False)
    width = None
    if ln.get("w"):
        try:
            width = int(ln.get("w"))
        except ValueError:
            width = None
    no_line = ln.find(_a("noFill")) is not None
    color = None
    solid = ln.find(_a("solidFill"))
    if solid is not None:
        for child in solid:
            color = _resolve_color(child, theme)
            if color:
                break
    return (width, color, no_line)


def _run_ea_font(run, theme_fonts):
    """解析单个 run 的东亚字体：优先 a:ea，其次 a:latin，最后回退主题字体。"""
    r_pr = run._r.find(_a("rPr"))
    if r_pr is not None:
        for tag in ("ea", "latin"):
            node = r_pr.find(_a(tag))
            if node is not None:
                face = (node.get("typeface") or "").strip()
                if face and not face.startswith("+"):
                    return face
                if face.startswith("+mj") or face.startswith("+mn"):
                    break
    # 未在 run 上显式指定时，回退到主题字体方案的次要（正文）字体
    for k, latin, ea, hans in theme_fonts:
        if k != "minor":
            continue
        for face in (ea, hans, latin):
            if face:
                return face
    return ""


def _is_songti(face):
    face = (face or "").strip()
    if not face:
        return False
    return any(face == s or face.startswith(s) for s in _SONGTI)


def _resolve_bg(slide):
    """按 幻灯片 -> 版式 -> 母版 的顺序取到生效的 p:bg 元素。"""
    for part in (slide, slide.slide_layout, slide.slide_layout.slide_master):
        c_sld = part._element.find(_p("cSld"))
        if c_sld is None:
            continue
        bg = c_sld.find(_p("bg"))
        if bg is not None:
            return bg
    return None


def _grad_stops(el, theme):
    """从任意含 a:gradFill 的元素中取出 [(pos 0~1, rgb), ...]，按 pos 升序。"""
    if el is None:
        return []
    grad = el.find(".//" + _a("gradFill"))
    if grad is None:
        return []
    stops = []
    for gs in grad.iter(_a("gs")):
        try:
            pos = float(gs.get("pos") or 0) / 100000.0
        except ValueError:
            pos = 0.0
        rgb = None
        for child in gs:
            rgb = _resolve_color(child, theme)
            if rgb:
                break
        if rgb:
            stops.append((pos, rgb))
    stops.sort(key=lambda x: x[0])
    return stops


def _full_page_grad(slide, theme, slide_w, slide_h):
    """页面上铺满整页的渐变填充形状同样视为页面背景（形状化背景）。"""
    for sh in _iter_shapes(slide.shapes):
        if int(sh.width or 0) < slide_w * 0.95 or int(sh.height or 0) < slide_h * 0.95:
            continue
        if _fill_kind(sh) != "gradFill":
            continue
        stops = _grad_stops(_sp_pr(sh), theme)
        if len(stops) >= 2:
            return stops
    return []


_RECT_GEOMS = ("rect", "roundRect", "round1Rect", "round2SameRect", "round2DiagRect",
               "snip1Rect", "snip2SameRect", "snip2DiagRect", "snipRoundRect", "flowChartProcess")


def _full_page_picture(slide, slide_w, slide_h):
    """返回铺满整页（≥90% 宽高）的图片形状，没有则返回 None。"""
    if slide is None:
        return None
    for sh in _iter_shapes(slide.shapes):
        if sh.shape_type is None and not hasattr(sh, "image"):
            continue
        if not hasattr(sh, "image"):
            continue
        if int(sh.width or 0) >= slide_w * 0.9 and int(sh.height or 0) >= slide_h * 0.9:
            return sh
    return None


def _dark_gold_cover(pic):
    """判断整版画面是否为"深蓝底 + 金色笔触"：统计像素中深蓝与金色的占比。"""
    if pic is None:
        return False, "尾页无整版画面"
    try:
        blob = pic.image.blob
    except Exception:
        return False, "整版画面不可读取"
    try:
        from PIL import Image     # pyright: ignore[reportMissingModuleSource]
        import io
        img = Image.open(io.BytesIO(blob)).convert("RGB")
        img = img.resize((60, 34))
        px = list(img.getdata())
    except Exception:
        return True, "尾页存在整版画面（%.0fKB）" % (len(blob) / 1024.0)
    blue = gold = 0
    for rgb in px:
        h, l, s = _hls(rgb)
        if 190 <= h <= 260 and l <= 0.40 and s >= 0.15:
            blue += 1
        elif 30 <= h <= 60 and s >= 0.25 and 0.25 <= l <= 0.85:
            gold += 1
    total = float(len(px)) or 1.0
    ok = blue / total >= 0.35 and gold / total >= 0.01
    return ok, "深蓝占比%.0f%%、金色占比%.0f%%" % (blue / total * 100, gold / total * 100)


def _pkg_info(file_path, prs):
    """统计包内母版/版式/媒体等资源情况，用于判断版面是否被整体重建。"""
    info = {"masters": 0, "layouts": 0, "media": 0, "used": [], "notice": False,
            "logo_masters": 0, "bgref_masters": 0, "plain_masters": 0}
    try:
        with zipfile.ZipFile(file_path) as zf:
            names = zf.namelist()
            info["masters"] = len([n for n in names
                                   if re.match(r"ppt/slideMasters/slideMaster\d+\.xml$", n)])
            info["layouts"] = len([n for n in names
                                   if re.match(r"ppt/slideLayouts/slideLayout\d+\.xml$", n)])
            info["media"] = len([n for n in names if n.startswith("ppt/media/")])
            used = []
            for slide in prs.slides:
                part = slide.slide_layout.slide_master.part.partname
                base = os.path.basename(str(part))
                if base not in used:
                    used.append(base)
            info["used"] = used
            notice = _norm("内部培训请勿传播")
            for base in used:
                try:
                    xml = zf.read("ppt/slideMasters/" + base).decode("utf-8", "ignore")
                except Exception:
                    continue
                if notice in _norm(re.sub(r"<[^>]+>", "", xml)):
                    info["notice"] = True
                if re.search(r"name=\"[^\"]*logo[^\"]*\"", xml, re.I):
                    info["logo_masters"] += 1
                if "<p:bgRef" in xml:
                    info["bgref_masters"] += 1
                if xml.count("<p:pic>") == 0 and xml.count("<p:grpSp>") == 0:
                    info["plain_masters"] += 1
    except Exception:
        pass
    return info


def _orange_frames(page, theme):
    """返回本页所有"橙色矩形边框"形状记录：[(shape, 线宽, 几何, 填充类型, box)]。"""
    out = []
    for sh in _iter_shapes(page["slide"].shapes):
        geom = _geom(sh)
        if geom not in _RECT_GEOMS:
            continue
        width, color, no_line = _line_info(sh, theme)
        if no_line or color is None or not _is_orange(color):
            continue
        out.append({"sh": sh, "w": width, "geom": geom,
                    "fill": _fill_kind(sh), "box": _box(sh),
                    "id": str(sh.shape_id)})
    return out


def _area(box):
    return max(0, box[2] - box[0]) * max(0, box[3] - box[1])


def _contains(outer, inner, tol=None):
    if tol is None:
        tol = int(0.15 * _EMU_CM)
    return (inner[0] >= outer[0] - tol and inner[1] >= outer[1] - tol
            and inner[2] <= outer[2] + tol and inner[3] <= outer[3] + tol)


def _text_insets(rec):
    body_pr = rec["sh"].text_frame._txBody.find(_a("bodyPr"))
    ins = {"lIns": 91440, "rIns": 91440, "tIns": 45720, "bIns": 45720}
    if body_pr is not None:
        for k in list(ins):
            if body_pr.get(k):
                try:
                    ins[k] = int(body_pr.get(k))
                except ValueError:
                    pass
    return ins


def _est_text_height(rec):
    """按字号与文本框宽度估算文字实际占用高度（EMU）。"""
    size = rec["max_size"]
    if not size:
        return 0
    ins = _text_insets(rec)
    inner_w = (rec["box"][2] - rec["box"][0]) - ins["lIns"] - ins["rIns"]
    per_char = size * _EMU_PT
    cpl = max(1, int(inner_w // per_char)) if per_char else 1
    lines = 0
    for para in rec["paras"]:
        txt = para.text or ""
        units = sum(1.0 if _is_cjk(c) else 0.52 for c in txt)
        lines += max(1, int(math.ceil(units / cpl)))
    return int(lines * size * 1.2 * _EMU_PT) + ins["tIns"] + ins["bIns"]


# -------------------------------------------------------------- 动画（p:timing）
_EFFECT_TAGS = ("animEffect", "anim", "animMotion", "animScale", "animRot", "set")


def _main_seq(slide):
    timing = slide._element.find(_p("timing"))
    if timing is None:
        return None
    for seq in timing.iter(_p("seq")):
        c_tn = seq.find(_p("cTn"))
        if c_tn is not None and c_tn.get("nodeType") == "mainSeq":
            return seq
    return None


def _cond_delays(el, stop_tag=_p("seq")):
    """收集元素自身及其祖先 cTn 上 stCondLst/cond 的 delay 值。"""
    out = []
    node = el
    while node is not None:
        if node.tag == _p("cTn"):
            st = node.find(_p("stCondLst"))
            if st is not None:
                for cond in st.findall(_p("cond")):
                    if cond.get("delay"):
                        out.append(cond.get("delay"))
        if node.tag == stop_tag:
            break
        node = node.getparent()
    return out


def _slide_effects(slide):
    """
    按文档顺序抽取本页动画效果，每项为 dict：
      spid      目标形状 id（字符串）
      node_type 所属 cTn 的 nodeType（clickEffect/afterEffect/withEffect 或 None）
      ident     效果标识（presetID 优先，其次 animEffect@filter，最后行为标签名）
      entr      是否为进入类动画（presetClass="entr" 或 filter 属于常见进入效果）
      delays    自身与祖先上的 cond delay 列表
    """
    seq = _main_seq(slide)
    if seq is None:
        return []
    effects = []
    owners = {}
    for beh in seq.iter():
        tag = beh.tag.split("}")[-1]
        if tag not in _EFFECT_TAGS or not beh.tag.startswith("{" + _NS_P):
            continue
        sp_tgt = beh.find(".//" + _p("spTgt"))
        if sp_tgt is None:
            continue
        node_type = None
        preset_id = None
        preset_class = None
        owner = None
        node = beh.getparent()
        while node is not None and node.tag != _p("seq"):
            if node.tag == _p("cTn"):
                n_type = node.get("nodeType")
                if n_type in ("mainSeq", "tmRoot"):
                    break          # 序列级容器不代表单个动画效果
                if owner is None:
                    owner = node
                if n_type or node.get("presetID"):
                    node_type = node_type or n_type
                    preset_id = preset_id or node.get("presetID")
                    preset_class = preset_class or node.get("presetClass")
                    owner = node
                    break
            node = node.getparent()
        filt = beh.get("filter") if tag == "animEffect" else None
        ident = ("preset:%s" % preset_id) if preset_id else (filt or tag)
        # 用 cTn 的 id 属性作键：lxml 元素代理的内存地址会被回收复用，不能用 id() 去重
        if owner is not None:
            key = ("cTn", owner.get("id"))
        else:
            inner = beh.find(".//" + _p("cTn"))
            key = ("beh", inner.get("id") if inner is not None else None)
        if key in owners:
            rec = owners[key]
            if filt and rec["ident"] in _EFFECT_TAGS:
                rec["ident"] = filt
            rec["entr"] = rec["entr"] or bool(filt)
            continue
        rec = {
            "spid": sp_tgt.get("spid"),
            "node_type": node_type,
            "ident": ident,
            "entr": (preset_class == "entr") or bool(filt) or tag == "set",
            "delays": _cond_delays(beh),
        }
        owners[key] = rec
        effects.append(rec)
    return effects


# ------------------------------------------------------------------ 页面结构识别
def _est_text_width(text, size_pt):
    """按字符宽度估算单行文本的渲染宽度（EMU）：中日韩全角按 1 em，西文按 0.52 em。"""
    if not text or not size_pt:
        return 0
    units = 0.0
    for ch in text:
        if _is_cjk(ch):
            units += 1.0
        elif ch == " ":
            units += 0.35
        else:
            units += 0.52
    return int(units * size_pt * _EMU_PT)


def _shape_records(slide, theme_fonts):
    """把一页内所有含文字的形状整理为便于判定的记录列表。"""
    recs = []
    for sh in _iter_shapes(slide.shapes):
        if not getattr(sh, "has_text_frame", False):
            continue
        tf = sh.text_frame
        raw = tf.text or ""
        runs = [r for para in tf.paragraphs for r in para.runs]
        sizes = [r.font.size.pt for r in runs if r.font.size is not None]
        bolds = [r.font.bold for r in runs if (r.text or "").strip()]
        body_pr = tf._txBody.find(_a("bodyPr"))
        recs.append({
            "sh": sh,
            "id": str(sh.shape_id),
            "name": sh.name or "",
            "raw": raw,
            "text": raw.strip(),
            "norm": _norm(raw),
            "cmp": _norm_cmp(raw),
            "box": _box(sh),
            "runs": runs,
            "sizes": sizes,
            "max_size": max(sizes) if sizes else None,
            "sz": max(sizes) if sizes else None,
            "bold": (True if bolds and all(b is True for b in bolds) else False),
            "paras": [p for p in tf.paragraphs],
            "n_paras": len([p for p in tf.paragraphs if (p.text or "").strip()]),
            "wrap": body_pr.get("wrap") if body_pr is not None else None,
            "empty": not raw.strip(),
            # 仅文本框与占位符算"文本框"；带文字框属性的装饰形状不算
            "is_txbox": ("TEXT_BOX" in str(sh.shape_type)
                         or bool(getattr(sh, "is_placeholder", False))),
        })
    return recs


def _text_left_edge(rec):
    """按段落对齐方式估算文字实际渲染的左边缘（EMU），居中时按框内居中折算。"""
    box = rec["box"]
    box_w = box[2] - box[0]
    ins = _text_insets(rec)
    lins, rins = ins["lIns"], ins["rIns"]
    inner = max(0, box_w - lins - rins)
    algn = str(rec["paras"][0].alignment) if rec["paras"] else "None"
    text_w = _est_text_width(rec["text"].split("\n")[0], rec["sz"])
    if algn.startswith("CENTER"):
        return box[0] + lins + max(0, (inner - text_w)) / 2.0
    if algn.startswith("RIGHT"):
        return box[0] + lins + max(0, inner - text_w)
    return box[0] + lins


def _match_title(norm_text, titles):
    """在标准标题列表中找到与文本对应的标题（允许多余尾随字符不超过 2 个）。"""
    for t in titles:
        nt = _norm(t)
        if norm_text == nt:
            return t
        if norm_text.startswith(nt) and len(norm_text) - len(nt) <= 2:
            return t
    return None


def _l1_core(title):
    """去掉"一、"之类序号前缀，返回一级标题正文部分。"""
    return _norm(title.split("、", 1)[-1])


def _is_keep_text(norm_text):
    return any(_norm(k) and _norm(k) in norm_text for k in _TEMPLATE_KEEP_TEXTS)


def _classify(prs, theme_fonts):
    """
    识别页面角色，返回 pages 列表，每项 dict：
      idx      1 起页码
      slide    幻灯片对象
      recs     含文字形状记录
      kind     cover / toc / section / body / tail / other
      l1       section 页对应的一级标题
      blocks   body 页的 [(二级标题记录, 正文记录)]
    """
    pages = []
    n = len(prs.slides)
    for i, slide in enumerate(prs.slides, 1):
        recs = _shape_records(slide, theme_fonts)
        filled = [r for r in recs if not r["empty"]]
        l2_hits = []
        for r in filled:
            t = _match_title(r["norm"], _L2_TITLES)
            if t is not None:
                l2_hits.append((t, r))
        l1_hits = []
        for r in filled:
            t = _match_title(r["norm"], _L1_TITLES)
            if t is not None:
                l1_hits.append((t, r))
        joined = "".join(r["norm"] for r in filled)
        page = {"idx": i, "slide": slide, "recs": recs, "kind": "other",
                "l1": None, "blocks": [], "l2": [t for t, _ in l2_hits]}
        if l2_hits:
            page["kind"] = "body"
            page["blocks"] = _pair_blocks(l2_hits, filled)
            if l1_hits:
                page["extra_l1"] = [r for _, r in l1_hits]
        elif "目录" in joined and sum(1 for t in _L1_TITLES if _l1_core(t) in joined) >= 5:
            page["kind"] = "toc"
        elif l1_hits and len(l1_hits) == 1:
            page["kind"] = "section"
            page["l1"] = l1_hits[0][0]
            page["l1_rec"] = l1_hits[0][1]
        elif i == 1 and _norm(_COVER_TITLE) in joined:
            page["kind"] = "cover"
        elif i == n:
            page["kind"] = "tail"
        pages.append(page)
    # 首页与尾页优先按位置确定，避免被 other 吞掉
    if pages:
        if pages[0]["kind"] in ("other", "cover"):
            pages[0]["kind"] = "cover"
        if pages[-1]["kind"] == "other":
            pages[-1]["kind"] = "tail"
    return pages


def _pair_blocks(l2_hits, filled):
    """为每个二级标题匹配同页最近的下方正文记录。"""
    title_ids = {id(r) for _, r in l2_hits}
    cands = [r for r in filled
             if id(r) not in title_ids
             and len(r["cmp"]) >= 30
             and not _is_keep_text(r["norm"])
             and _match_title(r["norm"], _L1_TITLES) is None]
    blocks = []
    used = set()
    for t, tr in sorted(l2_hits, key=lambda x: x[1]["box"][1]):
        best, best_d = None, None
        for c in cands:
            if id(c) in used:
                continue
            d = c["box"][1] - tr["box"][1]
            score = d if d >= 0 else (abs(d) + 10 ** 9)
            if best_d is None or score < best_d:
                best, best_d = c, score
        if best is not None:
            used.add(id(best))
        blocks.append((tr, best))
    return blocks


# ------------------------------------------------------------------------ 主入口
def _is_damaged_package(exc):
    """判断异常是否源于交付文件自身结构损坏，而非脚本逻辑缺陷。"""
    text = str(exc)
    if isinstance(exc, AttributeError) and (
        "slide_layout" in text or "slide_master" in text
    ):
        return True
    if isinstance(exc, KeyError) and (
        "There is no item named" in text or "no relationship of type" in text
    ):
        return True
    return type(exc).__name__ in ("XMLSyntaxError", "ParseError", "BadZipFile",
                                  "InvalidXmlError")


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
        "max_score": 0,
    }
    try:
        if Presentation is None:
            raise ImportError("缺少依赖 python-pptx，无法解析 .pptx")
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
        if _is_damaged_package(exc):
            # 交付文件结构损坏（缺版式/母版关系、XML 非法等）属于维度一不通过，
            # 按正常 0 分处理，不记为脚本运行错误。
            result["status"] = "ok"
            result["error"] = None
            result["dim1_pass"] = False
            result["dim1_reason"] = "交付文件结构不完整，无法正常解析：%s: %s" % (
                type(exc).__name__, exc,
            )
            result["dim2_items"] = []
            result["total_score"] = 0
            return result
        # 找不到文件／无法解析：dim1 未进入评测，dim1_pass 与 dim1_reason 留空
        result["status"] = "error"
        result["error"] = str(exc) or type(exc).__name__
        result["dim1_pass"] = ""
        result["dim1_reason"] = ""
        result["dim2_items"] = []
        result["total_score"] = 0
        result["max_score"] = 0
        return result


def _check_dim1(prs, pages):
    """
    维度1：可用性。
    ① 交付文件为 1 个可正常打开的 PowerPoint 演示文稿（.pptx）——能走到这里说明已解析成功；
    ② 演示文稿内至少有 1 页幻灯片。
    与办公软件对齐：以 PowerPoint 能否正常打开该文件为准。
    """
    if not len(prs.slides):
        return (False, "演示文稿内没有任何幻灯片")
    return (True, "交付文件为1个可正常打开的PowerPoint演示文稿（.pptx），共%d页"
            % len(prs.slides))


def _evaluate_impl(file_path, result):
    prs = Presentation(file_path)
    theme = _theme_colors(file_path)
    theme_fonts = _theme_fonts(file_path)
    slide_w = int(prs.slide_width or 0)
    slide_h = int(prs.slide_height or 0)
    pages = _classify(prs, theme_fonts)
    n_slides = len(pages)

    dim1_pass, dim1_reason = _check_dim1(prs, pages)
    result["dim1_pass"] = dim1_pass
    result["dim1_reason"] = dim1_reason
    if not dim1_pass:
        result["dim2_items"] = []
        result["total_score"] = 0
        result["max_score"] = MAX_SCORE_TOTAL
        return result

    cover = pages[0] if pages else None
    toc_pages = [p for p in pages if p["kind"] == "toc"]
    toc = toc_pages[0] if toc_pages else None
    section_pages = [p for p in pages if p["kind"] == "section"]
    body_pages = [p for p in pages if p["kind"] == "body"]
    tail = pages[-1] if pages else None

    # 正文页上的所有 (二级标题, 正文) 块
    all_blocks = []
    for pg in body_pages:
        for tr, br in pg["blocks"]:
            all_blocks.append((pg, tr, br))
    covered_l2 = set()
    for pg in body_pages:
        covered_l2.update(pg["l2"])
    l2_all_covered = (len(covered_l2) == 18)

    results = []

    # ---- 01 +5：按"一个标题板块一页"拆分为 30 页且顺序正确
    # ① 总页数为 30；② 顺序为 首页1 + 目录1 + (标题页1 + 正文页2) x 9 + 尾页1；
    # ③ 每个正文页只承载 1 个二级标题，不得把两个二级标题合并在同一页。
    # 与办公软件对齐：以 PowerPoint 幻灯片浏览视图中的页序与每页承载的标题数量为准。
    expect_kinds = ["cover", "toc"] + ["section", "body", "body"] * 9 + ["tail"]
    actual_kinds = [p["kind"] for p in pages]
    order_ok = (n_slides == 30 and actual_kinds == expect_kinds)
    if order_ok:
        seq_l1 = [p["l1"] for p in pages if p["kind"] == "section"]
        seq_l2 = [t for p in pages if p["kind"] == "body" for t in p["l2"]]
        order_ok = (seq_l1 == _L1_TITLES and seq_l2 == _L2_TITLES
                    and all(len(p["l2"]) == 1 for p in body_pages))
    if order_ok:
        results.append("+5：演示文稿按一个标题板块一页拆分为30页，"
                       "顺序为首页1页、目录页1页、9个一级标题页、18个二级标题正文页、尾页1页")
    else:
        results.append("+0：演示文稿未按一个标题板块一页拆分为30页并保持"
                       "首页/目录/9标题页/18正文页/尾页的顺序（实际%d页，"
                       "标题页%d个、正文页%d个）" % (n_slides, len(section_pages), len(body_pages)))

    # ---- 02 +5：九个一级标题各自单独成为一页标题页
    # ① 9 个一级标题全部出现；② 每个各占独立一页；③ 该页不承载二级标题与正文。
    # 与办公软件对齐：逐页查看是否只有一个一级标题作为该页主标题。
    sec_map = {}
    for pg in section_pages:
        sec_map.setdefault(pg["l1"], []).append(pg["idx"])
    sec_ok = (len(sec_map) == 9 and all(len(v) == 1 for v in sec_map.values())
              and all(t in sec_map for t in _L1_TITLES))
    if sec_ok:
        results.append("+5：九个一级标题各自单独成为一页标题页（9/9）")
    else:
        results.append("+0：九个一级标题未各自单独成为一页标题页（识别到%d个标题页，"
                       "覆盖%d个一级标题）" % (len(section_pages), len(sec_map)))

    # ---- 03 +5：18 个二级标题各自单独成页且同页放置对应正文
    # ① 18 个二级标题全部出现；② 每个二级标题独占一页；
    # ③ 该页同时放置与之对应的正文段落（正文与素材同源，相似度不低于 0.6）。
    # 与办公软件对齐：逐页确认标题与正文成对出现，且一页只讲一个二级标题。
    l2_pages = {}
    for pg in body_pages:
        for t in pg["l2"]:
            l2_pages.setdefault(t, []).append(pg["idx"])
    solo_ok = 0
    for pg, tr, br in all_blocks:
        if len(pg["l2"]) != 1 or br is None:
            continue
        src = _L2_BODIES.get(_match_title(tr["norm"], _L2_TITLES) or "", "")
        if src and difflib.SequenceMatcher(None, _norm_cmp(src), br["cmp"]).ratio() >= 0.6:
            solo_ok += 1
    l2_ok = (len(l2_pages) == 18 and all(len(v) == 1 for v in l2_pages.values())
             and solo_ok == 18)
    if l2_ok:
        results.append("+5：18个二级标题各自单独成页，且每页同时放置对应的正文内容（18/18）")
    else:
        results.append("+0：18个二级标题未各自单独成页并放置对应正文（覆盖%d个二级标题，"
                       "独立成页且正文对应的有%d个）" % (len(l2_pages), solo_ok))

    # ---- 04 +5：全文所有中文文字字体均为宋体
    # ① 逐个 run 解析东亚字体（a:ea，缺省时回退 a:latin 与主题字体方案）；
    # ② 只统计含中文字符的 run；③ 全部命中宋体（含 SimSun 等等价名）才得分。
    # 与办公软件对齐：选中文字后"字体"框显示的中文字体名称。
    font_total = 0
    font_ok = 0
    bad_faces = []
    for pg in pages:
        for rec in pg["recs"]:
            for run in rec["runs"]:
                txt = run.text or ""
                if not _has_cjk(txt):
                    continue
                font_total += 1
                face = _run_ea_font(run, theme_fonts)
                if _is_songti(face):
                    font_ok += 1
                elif face not in bad_faces:
                    bad_faces.append(face or "(未指定)")
    if font_total and font_ok == font_total:
        results.append("+5：全文所有中文文字字体均为宋体（%d/%d）" % (font_ok, font_total))
    else:
        results.append("+0：全文中文文字字体未统一为宋体（%d/%d，出现%s）"
                       % (font_ok, font_total, "、".join(bad_faces[:3]) or "无可判定字体"))

    # ---- 05 +5：所有页面背景均为绿色渐变，且渐变由浅到深
    # ① 逐页解析生效背景（幻灯片 -> 版式 -> 母版），或整页铺满的渐变填充形状；
    # ② 渐变色标均为绿色系；③ 沿渐变方向明度递减，即由浅到深。
    # 与办公软件对齐："设置背景格式"中看到的渐变填充色标与方向。
    bg_ok_pages = []
    bg_bad_pages = []
    darkest_l = None
    for pg in pages:
        stops = _grad_stops(_resolve_bg(pg["slide"]), theme)
        if len(stops) < 2:
            stops = _full_page_grad(pg["slide"], theme, slide_w, slide_h)
        ok = len(stops) >= 2 and all(_is_green(c) for _, c in stops)
        if ok:
            lums = [_hls(c)[1] for _, c in stops]
            ok = lums[0] - lums[-1] > 0.02 and all(
                lums[i] >= lums[i + 1] - 0.02 for i in range(len(lums) - 1))
            cur = min(lums)
            darkest_l = cur if darkest_l is None else min(darkest_l, cur)
        (bg_ok_pages if ok else bg_bad_pages).append(pg["idx"])
    if bg_ok_pages and not bg_bad_pages:
        results.append("+5：所有页面背景均为绿色渐变且呈由浅到深的变化（%d/%d页）"
                       % (len(bg_ok_pages), n_slides))
    else:
        results.append("+0：页面背景未全部为由浅到深的绿色渐变（合规%d/%d页）"
                       % (len(bg_ok_pages), n_slides))

    # ---- 06 +5：18 个正文页的正文外围各有一个橙色矩形边框，正文完整位于框内
    # ① 每个正文块都能找到一个带橙色边线的矩形类形状（文本框自身描边同样算）；
    # ② 正文文本框在该矩形范围之内；③ 正文估算高度不溢出该矩形。
    # 与办公软件对齐：以页面上看到的橙色方框及框内文字是否完整为准。
    frames_by_page = {pg["idx"]: _orange_frames(pg, theme) for pg in body_pages}
    frame_of_block = {}
    frame_ok = 0
    for pg, tr, br in all_blocks:
        if br is None:
            continue
        hit = None
        for fr in frames_by_page.get(pg["idx"], []):
            if _contains(fr["box"], br["box"]):
                if hit is None or _area(fr["box"]) < _area(hit["box"]):
                    hit = fr
        if hit is None:
            continue
        frame_of_block[(pg["idx"], tr["id"])] = hit
        inner_h = (hit["box"][3] - hit["box"][1])
        if _est_text_height(br) <= inner_h * 1.25:
            frame_ok += 1
    if all_blocks and frame_ok == len(all_blocks) and l2_all_covered:
        results.append("+5：正文页的正文文字外围均有一个橙色矩形边框，且正文文字完整位于框内"
                       "（%d/%d个正文块）" % (frame_ok, len(all_blocks)))
    else:
        results.append("+0：正文页的正文文字外围未全部套上橙色矩形边框或文字未完整落在框内"
                       "（%d/%d个正文块）" % (frame_ok, len(all_blocks)))

    # ---- 07 +5：正文页的二级标题与正文各自设置了进入动画，且两者效果不相同
    # ① 二级标题形状与正文形状都是动画目标；② 两者均为进入类动画；
    # ③ 两者效果标识（presetID 或动画滤镜）不同。
    # 与办公软件对齐："动画窗格"中两个对象各有一条进入动画且动画名称不同。
    anim_ok = 0
    anim_detail = []
    for pg, tr, br in all_blocks:
        if br is None:
            continue
        eff = _slide_effects(pg["slide"])
        t_eff = [e for e in eff if e["spid"] == tr["id"]]
        b_eff = [e for e in eff if e["spid"] == br["id"]]
        if not t_eff or not b_eff:
            anim_detail.append("第%d页缺动画" % pg["idx"])
            continue
        if not (any(e["entr"] for e in t_eff) and any(e["entr"] for e in b_eff)):
            anim_detail.append("第%d页非进入动画" % pg["idx"])
            continue
        if {e["ident"] for e in t_eff} & {e["ident"] for e in b_eff}:
            anim_detail.append("第%d页动画效果相同" % pg["idx"])
            continue
        anim_ok += 1
    if all_blocks and anim_ok == len(all_blocks) and l2_all_covered:
        results.append("+5：正文页的二级标题与正文各自设置了进入动画，且两者动画效果不相同"
                       "（%d/%d个正文块）" % (anim_ok, len(all_blocks)))
    else:
        results.append("+0：正文页的二级标题与正文未各自设置不同的进入动画（%d/%d个正文块，%s）"
                       % (anim_ok, len(all_blocks), "；".join(anim_detail[:3]) or "未识别到动画"))

    # ---- 08 +5：同一页面动画自动出来、换页之后动画通过单击召唤
    # ① 该页第一个动画为单击触发（clickEffect 且触发条件 delay=indefinite）；
    # ② 其后动画均为承接上一动画自动播放（afterEffect 或 withEffect）；
    # ③ 逐个正文页判定，全部满足才得分。
    # 与办公软件对齐："动画窗格"中第一条为"单击时"，其余为"上一动画之后/与上一动画同时"。
    click_ok = 0
    for pg in body_pages:
        eff = _slide_effects(pg["slide"])
        if len(eff) < 2:
            continue
        first = eff[0]
        if first["node_type"] != "clickEffect" or "indefinite" not in first["delays"]:
            continue
        if all(e["node_type"] in ("afterEffect", "withEffect") for e in eff[1:]):
            click_ok += 1
    if body_pages and click_ok == len(body_pages) and l2_all_covered:
        results.append("+5：正文页动画满足同页自动播放、换页后单击召唤"
                       "（首个动画为单击触发，其后为自动承接，%d/%d页）"
                       % (click_ok, len(body_pages)))
    else:
        results.append("+0：正文页动画未满足首个动画单击触发、其后自动承接"
                       "（%d/%d页）" % (click_ok, len(body_pages)))
    # ---- 09 +3：首页标题文字为"艺术漆北区销售额提升方案"
    # ① 首页存在文本对象；② 其文字与指定标题完全一致（忽略空白差异）。
    # 与办公软件对齐：打开第 1 页看到的主标题文字。
    cover_recs = [r for r in (cover["recs"] if cover else []) if not r["empty"]]
    cover_title = None
    for r in cover_recs:
        if r["norm"] == _norm(_COVER_TITLE):
            if cover_title is None or (r["max_size"] or 0) > (cover_title["max_size"] or 0):
                cover_title = r
    if cover_title is not None:
        results.append("+3：首页标题文字为“%s”" % _COVER_TITLE)
    else:
        results.append("+0：首页标题文字不是“%s”" % _COVER_TITLE)

    # ---- 10 +3：首页标题与 9 个一级标题页的标题均加粗，且各自排在同一行内不换行
    # ① 10 个标题的所有文字 run 都设置了加粗；② 只有一个段落；
    # ③ 文本不换行（bodyPr wrap="none"，或按字号估算的单行宽度不超过文本框可用宽度）。
    # 与办公软件对齐：标题呈粗体且视觉上未折到第二行。
    title_recs = ([cover_title] if cover_title is not None else [])
    title_recs += [pg["l1_rec"] for pg in section_pages if pg.get("l1_rec") is not None]
    bold_ok = 0
    line_ok = 0
    for r in title_recs:
        if r["bold"]:
            bold_ok += 1
        ins = _text_insets(r)
        avail = (r["box"][2] - r["box"][0]) - ins["lIns"] - ins["rIns"]
        est = _est_text_width(r["text"], r["max_size"] or 0)
        if r["n_paras"] <= 1 and (r["wrap"] == "none" or est <= avail):
            line_ok += 1
    need_titles = 10
    if len(title_recs) == need_titles and bold_ok == need_titles and line_ok == need_titles:
        results.append("+3：首页标题与9个一级标题页标题均加粗，且每个标题排在同一行内不换行"
                       "（%d/%d）" % (bold_ok, need_titles))
    else:
        results.append("+0：首页标题与9个一级标题页标题未做到均加粗且单行不换行"
                       "（识别到%d个标题，加粗%d个、单行%d个）"
                       % (len(title_recs), bold_ok, line_ok))

    # ---- 11 +3：首页标题横向宽度占页面宽度四分之三以上
    # ① 按字号与字符数估算标题文字的实际渲染宽度；
    # ② 该宽度不小于页面宽度的 75%，即"尽量占满一行"。
    # 与办公软件对齐：首页标题在版面上横向铺开的实际占位，而非文本框空框的宽度。
    if cover_title is not None:
        est_w = _est_text_width(cover_title["text"], cover_title["max_size"] or 0)
        ratio = (float(est_w) / slide_w) if slide_w else 0.0
    else:
        est_w, ratio = 0, 0.0
    if ratio >= 0.75:
        results.append("+3：首页标题文字横向宽度占页面宽度的%.0f%%，达到占满一行的效果"
                       % (ratio * 100))
    else:
        results.append("+0：首页标题文字横向宽度仅占页面宽度的%.0f%%，未达到四分之三"
                       % (ratio * 100))
    # ---- 12 +3：目录页存在，且完整列出九个一级标题的文字
    # ① 存在一页可识别为目录页；② 页面文字中九个一级标题的正文部分全部出现。
    # 与办公软件对齐：目录页上逐条比对九个章节名称是否齐全。
    toc_joined = "".join(r["norm"] for r in (toc["recs"] if toc else []) if not r["empty"])
    toc_hit = sum(1 for t in _L1_TITLES if _l1_core(t) in toc_joined)
    if toc is not None and toc_hit == 9:
        results.append("+3：目录页存在，且完整列出九个一级标题的文字（9/9）")
    else:
        results.append("+0：目录页缺失或未完整列出九个一级标题的文字（%d/9）" % toc_hit)

    # ---- 13 +3：绿色渐变整体明度偏浅，最深处仍为中浅绿
    # ① 取全篇绿色渐变中最深色标的明度；② 明度不低于 0.35 视为中浅绿；
    # ③ 明度过低（深绿/墨绿）则不得分。
    # 与办公软件对齐：观感上背景为浅绿到中浅绿，文字与背景对比清晰可辨。
    if darkest_l is not None and darkest_l >= 0.35:
        results.append("+3：绿色渐变整体明度偏浅，最深处明度为%.2f，仍为中浅绿，"
                       "文字与背景对比清晰可辨" % darkest_l)
    else:
        results.append("+0：绿色渐变明度不满足偏浅要求（最深处明度%s）"
                       % ("%.2f" % darkest_l if darkest_l is not None else "未识别"))

    # ---- 14 +3：正文页橙色矩形边框线宽不小于 2 磅
    # ① 取所有正文块对应的橙色框；② 线宽均不小于 25400 EMU（2 磅）。
    # 与办公软件对齐："设置形状格式-线条-宽度"读到的磅值。
    widths = [fr["w"] for fr in frame_of_block.values()]
    width_ok = [w for w in widths if w is not None and w >= _MIN_LINE_W]
    if widths and len(width_ok) == len(widths):
        results.append("+3：正文页橙色矩形边框线宽均不小于2磅（最小%.1f磅）"
                       % (min(widths) / _EMU_PT))
    else:
        results.append("+0：正文页橙色矩形边框线宽未全部达到2磅（合规%d/%d）"
                       % (len(width_ok), len(widths)))
    # ---- 15 +3：18 个正文页的二级标题均位于页面最上方区域且靠左对齐
    # ① 二级标题是本页位置最靠上的文字对象（版面自带附属文字不计）；
    # ② 顶端落在页面上部四分之一以内；③ 左端靠左（页面左侧四分之一以内）且段落左对齐。
    # 与办公软件对齐：肉眼看二级标题压在页面顶部并贴左边距。
    top_ok = 0
    for pg, tr, br in all_blocks:
        others = [r for r in pg["recs"] if not r["empty"] and id(r) != id(tr)
                  and not _is_keep_text(r["norm"])]
        is_top = all(r["box"][1] >= tr["box"][1] - int(0.1 * _EMU_CM) for r in others)
        left = _text_left_edge(tr)
        if is_top and tr["box"][1] <= slide_h * 0.25 and left <= slide_w * 0.25:
            top_ok += 1
    if all_blocks and top_ok == len(all_blocks) and l2_all_covered:
        results.append("+3：正文页的二级标题均位于页面最上方区域且靠左对齐（%d/%d）"
                       % (top_ok, len(all_blocks)))
    else:
        results.append("+0：正文页的二级标题未全部位于页面最上方区域并靠左对齐（%d/%d）"
                       % (top_ok, len(all_blocks)))

    # ---- 16 +3：18 个正文页的正文文字与所给素材内容一致
    # ① 与素材段落的相似度不低于 0.98；② 结尾未被截断（素材末尾 8 字须出现）；
    # ③ 无语句重复（同一 12 字片段不得出现两次）。
    # 与办公软件对齐：逐页比对正文是否缺句、半句截断或整句重复。
    text_ok = 0
    text_bad = []
    for pg, tr, br in all_blocks:
        key = _match_title(tr["norm"], _L2_TITLES)
        src = _norm_cmp(_L2_BODIES.get(key or "", ""))
        got = br["cmp"] if br is not None else ""
        if not src or not got:
            text_bad.append("第%d页无正文" % pg["idx"])
            continue
        ratio = difflib.SequenceMatcher(None, src, got).ratio()
        truncated = src[-8:] not in got
        dup = any(got.count(got[i:i + 12]) > 1 for i in range(0, max(0, len(got) - 12), 4))
        if ratio >= 0.98 and not truncated and not dup:
            text_ok += 1
        else:
            text_bad.append("第%d页%s" % (pg["idx"], "内容重复" if dup else
                                        ("结尾截断" if truncated else "与素材不一致")))
    if all_blocks and text_ok == len(all_blocks) and l2_all_covered:
        results.append("+3：正文页的正文文字与所给素材内容一致，无缺句、无中途截断、无语句重复"
                       "（%d/%d）" % (text_ok, len(all_blocks)))
    else:
        results.append("+0：正文页的正文文字与素材内容不一致（%d/%d，%s）"
                       % (text_ok, len(all_blocks), "；".join(text_bad[:3])))

    # ---- 17 +3：结尾板块内容放置在最后一页
    # ① 最后一页不承载一级/二级标题正文；② 该页有结尾性文字或整版结尾画面；
    # ③ 结尾内容没有被放到中间页。
    # 与办公软件对齐：最后一页呈现结尾致谢版面。
    tail_joined = "".join(r["norm"] for r in (tail["recs"] if tail else []) if not r["empty"])
    tail_pic = _full_page_picture(tail["slide"], slide_w, slide_h) if tail else None
    tail_is_closing = bool(tail) and tail["kind"] == "tail" and (
        any(w in tail_joined.lower() for w in _CLOSING_WORDS) or tail_pic is not None)
    if tail_is_closing and tail_joined:
        results.append("+3：结尾板块内容放置在最后一页（第%d页）" % tail["idx"])
    else:
        results.append("+0：结尾板块内容未放置在最后一页")

    # ---- 18 +3：尾页沿用深蓝色底、金色笔触画面的整版版式呈现
    # ① 最后一页存在铺满整页的图片对象；② 该画面主色调为深蓝并带金色笔触。
    # 与办公软件对齐：尾页整版铺一张深蓝底、金色笔触的画面。
    tail_style_ok, tail_style_note = _dark_gold_cover(tail_pic)
    if tail_style_ok:
        results.append("+3：尾页沿用深蓝色底、金色笔触画面的整版版式呈现（%s）" % tail_style_note)
    else:
        results.append("+0：尾页未沿用深蓝色底、金色笔触画面的整版版式（%s）" % tail_style_note)
    # ---- 19 +1：同类页面的文字字号与加粗状态保持一致
    # ① 九个一级标题页的标题字号一致；② 全部二级标题字号与加粗状态一致；
    # ③ 全部正文字号一致。
    # 与办公软件对齐：翻页时同类文字大小不跳变。
    sec_sizes = {r["sz"] for r in (p["l1_rec"] for p in section_pages) if r is not None
                 and r["sz"] is not None}
    l2_sizes = {tr["sz"] for _, tr, _ in all_blocks if tr["sz"] is not None}
    l2_bolds = {tr["bold"] for _, tr, _ in all_blocks}
    bd_sizes = {br["sz"] for _, _, br in all_blocks if br is not None and br["sz"] is not None}
    consist_bad = []
    if len(sec_sizes) > 1:
        consist_bad.append("一级标题页字号%s" % sorted(sec_sizes))
    if len(l2_sizes) > 1:
        consist_bad.append("二级标题字号%s" % sorted(l2_sizes))
    if len(l2_bolds) > 1:
        consist_bad.append("二级标题加粗不一致")
    if len(bd_sizes) > 1:
        consist_bad.append("正文字号%s" % sorted(bd_sizes))
    if not consist_bad and sec_sizes and l2_sizes and bd_sizes:
        results.append("+1：同类页面的文字字号与加粗状态保持一致（一级标题%gpt／二级标题%gpt／"
                       "正文%gpt）" % (list(sec_sizes)[0], list(l2_sizes)[0], list(bd_sizes)[0]))
    else:
        results.append("+0：同类页面的文字字号或加粗状态不一致（%s）"
                       % ("；".join(consist_bad) if consist_bad else "同类文字缺失"))

    # ---- 20 +1：全文无空白页、无空文本框、无多余空行与行首空格
    # ① 每一页都有文字或整版画面；② 不存在没有文字的文本框；
    # ③ 正文段落无空段、段首无空格。
    # 与办公软件对齐：没有翻到空页，也没有点开却是空的文本框。
    blank_pages = [p["idx"] for p in pages
                   if not any(not r["empty"] for r in p["recs"])
                   and _full_page_picture(p["slide"], slide_w, slide_h) is None]
    empty_boxes = [(p["idx"], r["name"]) for p in pages for r in p["recs"]
                   if r["empty"] and r["is_txbox"]]
    space_bad = []
    for pg, tr, br in all_blocks:
        for rec in (tr, br):
            if rec is None:
                continue
            for para in rec["paras"]:
                raw = para.text or ""
                if not raw.strip():
                    space_bad.append("第%d页空行" % pg["idx"])
                elif raw[:1] in (" ", "　", "\t"):
                    space_bad.append("第%d页段首空格" % pg["idx"])
    if not blank_pages and not empty_boxes and not space_bad:
        results.append("+1：全文无空白页、无空文本框、无多余空行与行首空格")
    else:
        note = []
        if blank_pages:
            note.append("空白页%s" % blank_pages[:3])
        if empty_boxes:
            note.append("空文本框%s" % ["第%d页%s" % e for e in empty_boxes[:3]])
        if space_bad:
            note.append("；".join(dict.fromkeys(space_bad))[:40])
        results.append("+0：存在空白页／空文本框／多余空行或行首空格（%s）" % "，".join(note))
    # ================================================================ 减分项 21-33
    pkg = _pkg_info(file_path, prs)

    # ---- 21 -5：幻灯片母版与版式沿用既有版面
    # ① 母版/版式数量未被压缩到"只剩一套"；② 所用母版未被换成纯色主题背景 bgRef；
    # ③ 所用母版仍保留页脚栏、分隔线、辅助数据形状等版面元素（pic / grpSp）。
    # 与办公软件对齐：在"视图-幻灯片母版"中仍能看到原有的多套母版及其版面元素。
    m_bad = []
    if pkg["layouts"] < 20:
        m_bad.append("版式仅剩%d个（母版%d套）" % (pkg["layouts"], pkg["masters"]))
        if pkg["bgref_masters"]:
            m_bad.append("母版背景改为纯色主题背景")
    if pkg["plain_masters"]:
        m_bad.append("%d套所用母版丢失页脚栏与辅助形状" % pkg["plain_masters"])
    if m_bad:
        results.append("-5：幻灯片母版与版式未沿用既有版面（%s）" % "，".join(m_bad))
    else:
        results.append("-0：幻灯片母版与版式沿用既有版面（母版%d套、版式%d个，"
                       "所用%d套母版均保留版面元素）"
                       % (pkg["masters"], pkg["layouts"], len(pkg["used"])))

    # ---- 22 -3：右下角标识图形、上方提示文字与随包图片素材保持齐全
    # ① 所用母版中仍有名称含 Logo 的标识图形；② 仍有"内部培训请勿传播"提示文字；
    # ③ 随包媒体文件未被大幅删减到仅剩个别文件。
    # 与办公软件对齐：每页右下角能看到蓝色标识及其上方的提示文字。
    logo_ok = pkg["logo_masters"] >= max(1, len(pkg["used"]))
    media_ok = pkg["media"] >= 20
    if logo_ok and pkg["notice"] and media_ok:
        results.append("-0：每页右下角保留标识图形与「内部培训请勿传播」提示文字，"
                       "随包图片素材齐全（%d个媒体文件）" % pkg["media"])
    else:
        note = []
        if not logo_ok:
            note.append("标识图形缺失（%d/%d套母版含标识）" % (pkg["logo_masters"], len(pkg["used"])))
        if not pkg["notice"]:
            note.append("「内部培训请勿传播」提示文字缺失")
        if not media_ok:
            note.append("媒体文件仅剩%d个" % pkg["media"])
        results.append("-3：右下角标识图形／提示文字／随包图片素材未保持齐全（%s）" % "，".join(note))

    # ---- 23 -3：页面尺寸保持 12192000 x 6858000 EMU
    # ① 宽度等于 12192000 EMU；② 高度等于 6858000 EMU；③ 未被改成非标准近似值。
    # 与办公软件对齐："设计-幻灯片大小"中宽 33.87 厘米、高 19.05 厘米。
    if slide_w == 12192000 and slide_h == 6858000:
        results.append("-0：页面尺寸保持宽12192000 EMU、高6858000 EMU（33.87×19.05厘米）")
    else:
        results.append("-3：页面尺寸被改为宽%d EMU、高%d EMU（%.2f×%.2f厘米），"
                       "非12192000×6858000标准值"
                       % (slide_w, slide_h, slide_w / _EMU_CM, slide_h / _EMU_CM))

    # ---- 24 -3：页面上除绿色渐变背景外没有其他背景装饰
    # ① 页面上没有额外添加的圆形装饰；② 装饰不出现在页面右上角或左下角；
    # ③ 不以半透明浅色形状叠加在背景上。
    # 与办公软件对齐：页面四角干净，只有渐变底色。
    deco = []
    for p in pages:
        for sh in _iter_shapes(p["slide"].shapes):
            if _geom(sh) not in ("ellipse", "circle", "pie", "chord", "blob"):
                continue
            if _fill_kind(sh) in ("noFill", "none", None):
                continue
            deco.append("第%d页%s" % (p["idx"], sh.name))
    if not deco:
        results.append("-0：页面上除绿色渐变背景外没有其他背景装饰形状")
    else:
        results.append("-3：页面上额外添加了背景装饰形状（共%d处，如%s）"
                       % (len(deco), "、".join(deco[:3])))

    # ---- 25 -3：正文外围的橙色框为无填充的直角矩形
    # ① 几何形状为直角矩形 rect；② 填充为无填充，不遮盖页面渐变底色；
    # ③ 未改为白色等实色填充的圆角矩形。
    # 与办公软件对齐：橙色框内能看到与页面一致的绿色渐变底色。
    frames_all = [fr for p in body_pages for fr in _orange_frames(p, theme)]
    bad_geom = [fr for fr in frames_all if fr["geom"] != "rect"]
    bad_fill = [fr for fr in frames_all if fr["fill"] not in ("noFill", "none", None)]
    if not bad_geom and not bad_fill:
        results.append("-0：正文外围的橙色框为无填充的直角矩形（%d个）" % len(frames_all))
    else:
        results.append("-3：正文外围的橙色框不是无填充的直角矩形（共%d个，非直角矩形%d个、"
                       "带填充%d个）" % (len(frames_all), len(bad_geom), len(bad_fill)))

    # ---- 26 -3：一级标题页只显示该一级标题文字
    # ① 页面上除该一级标题外没有章节编号数字；② 没有英文副标；
    # ③ 没有额外的橙色短线等装饰形状。
    # 与办公软件对齐：标题页画面上只有一行中文标题。
    sec_extra = []
    for p in section_pages:
        for r in p["recs"]:
            if r["empty"] or _is_keep_text(r["norm"]):
                continue
            if p["l1_rec"] is not None and id(r) == id(p["l1_rec"]):
                continue
            sec_extra.append("第%d页「%s」" % (p["idx"], r["norm"][:12]))
        for sh in _iter_shapes(p["slide"].shapes):
            box = _box(sh)
            w, h = box[2] - box[0], box[3] - box[1]
            if h <= 0 or w <= 0 or h > 0.03 * slide_h or w < 0.02 * slide_w:
                continue
            col = _fill_color(sh, theme)
            if col is not None and _is_orange(col):
                sec_extra.append("第%d页橙色短线装饰" % p["idx"])
    if not sec_extra:
        results.append("-0：九个一级标题页均只显示该一级标题文字，无章节编号、英文副标与装饰短线")
    else:
        results.append("-3：一级标题页除标题外还显示了其他内容（共%d处，如%s）"
                       % (len(sec_extra), "、".join(dict.fromkeys(sec_extra))[:60]))
    # ---- 27 -1：主题字体方案的东亚字体设置为宋体
    # ① 主题中 majorFont/minorFont 的东亚字体项写明宋体；
    # ② 不是东亚字体项为空的默认主题；③ 中文因此不会回落到默认黑体系字体。
    # 与办公软件对齐："设计-字体"中中文标题与正文字体显示为宋体。
    ea_faces = [(k, ea) for k, lat, ea, hans in theme_fonts]
    ea_ok = bool(ea_faces) and all(_is_songti(ea) for _, ea in ea_faces)
    if ea_ok:
        results.append("-0：主题字体方案的东亚字体设置为宋体（%s）"
                       % "、".join("%s=%s" % (k, ea) for k, ea in ea_faces))
    else:
        results.append("-1：主题字体方案的东亚字体未设置为宋体（%s）"
                       % "、".join("%s=%s" % (k, ea or "（空）") for k, ea in ea_faces))

    # ---- 28 -1：首页标题区域下方没有横条色块
    # ① 首页没有实色填充的宽扁形状；② 该形状不横跨页面且高度不占小半页；
    # ③ 首页底色为完整的渐变，未被色块压住。
    # 与办公软件对齐：首页画面上没有一条白色长横条。
    bars = []
    if cover is not None:
        for sh in _iter_shapes(cover["slide"].shapes):
            if _geom(sh) not in _RECT_GEOMS:
                continue
            if _fill_kind(sh) not in ("solidFill",):
                continue
            if sh.has_text_frame and _norm(sh.text_frame.text):
                continue
            box = _box(sh)
            w, h = box[2] - box[0], box[3] - box[1]
            if w >= 0.5 * slide_w and 0.03 * slide_h <= h <= 0.30 * slide_h:
                bars.append("%s（%.2f×%.2f厘米）" % (sh.name, w / _EMU_CM, h / _EMU_CM))
    if not bars:
        results.append("-0：首页标题区域下方没有横条色块")
    else:
        results.append("-1：首页增加了横条色块（%s）" % "、".join(bars[:2]))

    # ---- 29 -1：首页除标题外没有其他新增文字
    # ① 首页没有新增副标题句；② 没有新增底部年度说明；
    # ③ 版面自带的固定文字不计入新增。
    # 与办公软件对齐：首页画面上只有主标题一句话是新写上去的。
    cover_extra = []
    if cover is not None:
        for r in cover["recs"]:
            if r["empty"] or _is_keep_text(r["norm"]):
                continue
            if _norm(_COVER_TITLE) in r["norm"]:
                continue
            cover_extra.append("「%s」" % r["norm"][:16])
    if not cover_extra:
        results.append("-0：首页除标题外没有其他新增文字")
    else:
        results.append("-1：首页除标题外增加了其他文字（%d处：%s）"
                       % (len(cover_extra), "、".join(cover_extra[:3])))

    # ---- 30 -1：目录页只放置目录标题与九个一级标题条目
    # ① 没有新增导语说明句；② 除标题与条目外没有其他成句文字；
    # ③ 条目文字与一级标题一一对应。
    # 与办公软件对齐：目录页上没有一句多余的介绍语。
    toc_extra = []
    if toc is not None:
        for r in toc["recs"]:
            if r["empty"] or _is_keep_text(r["norm"]):
                continue
            t = r["norm"]
            if "目录" in t or _match_title(t, _L1_TITLES) is not None:
                continue
            if any(_l1_core(x) in t for x in _L1_TITLES):
                continue
            if len(t) <= 3:
                continue
            toc_extra.append("「%s」" % t[:20])
    if not toc_extra:
        results.append("-0：目录页只放置目录标题与九个一级标题条目，无多余说明句")
    else:
        results.append("-1：目录页增加了多余说明文字（%d处：%s）"
                       % (len(toc_extra), "、".join(toc_extra[:2])))

    # ---- 31 -1：目录页条目文字保留顿号
    # ① 九个条目均保留"一、"至"九、"的顿号写法；
    # ② 未以空格替代顿号；③ 序号与标题之间没有被改成其他分隔符。
    # 与办公软件对齐：目录条目读作"一、市场趋势与消费升级"。
    dun_hit = sum(1 for t in _L1_TITLES if t[:2] in toc_joined) if toc is not None else 0
    if toc is None or dun_hit == 9:
        results.append("-0：目录页九个条目均保留顿号写法（%d/9）" % dun_hit)
    else:
        results.append("-1：目录页条目文字未保留顿号（%d/9条保留），序号与标题之间被改为空格等写法"
                       % dun_hit)

    # ---- 32 -1：正文页最上方靠左位置只显示二级标题
    # ① 二级标题上方没有另加一条一级标题文字；
    # ② 页面顶部只有一行标题；③ 不重复出现所属章节名。
    # 与办公软件对齐：正文页顶部只有"1.1 市场扩容趋势"这类二级标题。
    extra_l1_pages = [p["idx"] for p in body_pages if p.get("extra_l1")]
    if not extra_l1_pages:
        results.append("-0：正文页最上方靠左位置只显示二级标题，未另加一级标题文字")
    else:
        results.append("-1：正文页在二级标题上方另加了一级标题文字（共%d页，如第%s页）"
                       % (len(extra_l1_pages), "、".join(str(i) for i in extra_l1_pages[:3])))

    # ---- 33 -1：正文页没有额外标注
    # ① 没有"自动呈现"等动画提示文字；② 页面右上角没有新增标签色块与标签文字；
    # ③ 页面上除二级标题与正文外没有其他文字。
    # 与办公软件对齐：正文页画面上只有标题与正文，没有小标签。
    mark_extra = []
    for p in body_pages:
        known = set()
        for tr, br in p["blocks"]:
            known.add(id(tr))
            if br is not None:
                known.add(id(br))
        for r in p["recs"]:
            if r["empty"] or _is_keep_text(r["norm"]) or id(r) in known:
                continue
            if _match_title(r["norm"], _L1_TITLES) is not None:
                continue          # 一级标题另由第 32 条判定
            mark_extra.append("第%d页「%s」" % (p["idx"], r["norm"][:10]))
    if not mark_extra:
        results.append("-0：正文页没有额外标注，页面上只有二级标题与正文")
    else:
        results.append("-1：正文页增加了额外标注文字或标签（共%d处，如%s）"
                       % (len(mark_extra), "、".join(mark_extra[:3])))

    # ------------------------------------------------------------ 汇总为返回结构
    assert len(results) == len(_MAX_DELTAS_ORDER), \
        "细则条目数不匹配：%d vs %d" % (len(results), len(_MAX_DELTAS_ORDER))
    dim2_items = []
    total = 0
    for raw, max_delta in zip(results, _MAX_DELTAS_ORDER):
        head, _, rule = raw.partition("：")
        try:
            delta = int(head.strip())
        except ValueError:
            delta = 0
        hit = (delta > 0) if max_delta >= 0 else (delta < 0)
        got = max_delta if hit else 0
        total += got
        dim2_items.append({"rule": rule, "max_delta": max_delta,
                           "delta": got, "hit": hit})
    result["dim2_items"] = dim2_items
    result["total_score"] = total
    result["max_score"] = MAX_SCORE_TOTAL
    return result


if __name__ == "__main__":
    _dir = sys.argv[1] if len(sys.argv) > 1 else os.path.dirname(os.path.abspath(__file__))
    _out = evaluate(_dir)
    print(json.dumps(_out, ensure_ascii=False, indent=2))









