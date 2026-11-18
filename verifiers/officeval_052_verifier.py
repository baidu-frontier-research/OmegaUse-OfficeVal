#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# pyright: reportUnknownMemberType=false, reportUnknownVariableType=false
# pyright: reportUnknownArgumentType=false, reportMissingImports=false
# pyright: reportOptionalMemberAccess=false, reportAttributeAccessIssue=false
"""
officeval_052_verifier —「数字化校园阅读项目实施的三维协同模型」PPT 评测脚本

用法：
    python officeval_052_verifier.py <存放交付 .pptx 的目录>
不传参数时，默认扫描脚本自身所在目录。

返回：evaluate(dir_path) -> dict
    {
      "id": "052", "file_name": ..., "status": "ok"|"error", "error": None|str,
      "dim1_pass": ""|bool, "dim1_reason": str,
      "dim2_items": [{"rule","max_delta","delta","hit"}, ...],
      "total_score": int, "max_score": int
    }

约定：
  * status="error" 表示这份交付根本没能进入评分（目录不存在、目录中无 .pptx、
    脚本异常），此时 dim1_pass 为空串 ""（未判定）、dim2_items 为 []、
    total_score 与 max_score 均为 0。
  * dim1_pass 为 False 时，dim1_reason 必须带错误摘要；为 True 时 dim1_reason 为空串。
  * dim2_items 每项只输出 rule / max_delta / delta / hit 四个键，判定说明由 rule 原文承载；
    delta 只取 0 或 max_delta 两个值。
"""

import os
import re
import sys
import json
import colorsys

from pptx import Presentation
from pptx.util import Emu  # noqa: F401  （保持与办公软件单位换算的显式依赖）

SCRIPT_ID = "052"
MAX_SCORE_TOTAL = 119

# 维度2 细则原文，行序与《评测细则.txt》严格一致，逐条对应 results 的追加顺序。
# 返回信息里的 rule 一律取自本表，不随判定结果变化、不附带实测值。
_DIM2_RULES = [
    (5, '画面全部元素的整体外接矩形宽高比在1.20–1.50之间'),
    (3, '画面整体外接矩形的左右空白之差≤1.5cm，上下空白之差≤1.5cm'),
    (5, '承担主色的形状填充与边线均为橙色系（色相20°–45°，近似#E07000–#FAD7A8）'),
    (5, '页面上部出现一个朝上箭头形状，下部为竖直箭身、上部为三角形头部，位于页面从左至右30%–70%、从上至下0%–40%，宽6–12cm、高5–8cm，橙色系填充'),
    (3, '页面上部朝上箭头内出现"目标层"文本：白色、加粗、18–28磅，水平中心位于页面从左至右45%–55%，顶端位于从上至下3%–20%'),
    (3, '页面上部朝上箭头的箭身内自上而下依次出现"核心素养导向""阅读兴趣培养""信息获取与表达""校园文化共建"四行文本：各占一行、白色、12–18磅、水平居中，整体位于页面从上至下12%–36%'),
    (1, '页面上部朝上箭头内"目标层"与四行条目文本之间出现一条白色横线，长3–7cm，线的中部有一个白色实心小圆点，直径0.15–0.4cm'),
    (5, '页面下部出现一个朝下箭头形状，上部为竖直箭身、下部为三角形头部，位于页面从左至右30%–70%、从上至下58%–99%，宽6–12cm、高5–8cm，橙色系填充'),
    (3, '页面下部朝下箭头内出现"成效层"文本：白色、加粗、18–28磅，水平中心位于页面从左至右45%–55%，顶端位于从上至下62%–72%'),
    (3, '页面下部朝下箭头的箭身内自上而下依次出现"阅读习惯形成""思维能力提升""协作交流增强""自主学习深化"四行文本：各占一行、白色、12–18磅、水平居中，整体位于页面从上至下69%–92%'),
    (1, '页面下部朝下箭头内"成效层"与四行条目文本之间出现一条白色横线，长3–7cm，线的中部有一个白色实心小圆点，直径0.15–0.4cm'),
    (3, '页面上部朝上箭头与页面下部朝下箭头的箭身内，四行条目文本的行间各出现3条白色横向分隔线，分别位于第1行、第2行、第3行文本的下方，两处箭头合计6条：均为虚线或点线样式，线宽≤1.5磅，长度3–5cm，水平中心位于页面从左至右45%–55%，且长度小于所在箭身的宽度'),
    (5, '页面中部出现一个圆角矩形或胶囊形形状，位于页面从左至右25%–75%、从上至下38%–60%，宽8–16cm、高2.5–4cm，填充浅橙黄色（R≥240、G在200–245之间、B在140–235之间），边线为橙色实线且线宽≥1.5磅'),
    (3, '页面中部圆角矩形内上部出现"实施枢纽"文本：深橙色或棕色（近似#8A4600–#C87000）、加粗、18–26磅，水平中心位于页面从左至右45%–55%'),
    (1, '页面中部圆角矩形内下部出现单行文本"课程内容·技术平台·活动组织"，三项之间以间隔点"·"分隔，12–18磅、水平居中'),
    (1, '页面中部圆角矩形内"实施枢纽"与"课程内容·技术平台·活动组织"之间出现一条橙色横线，长6–13cm，线的中部有一个橙色实心小圆点，直径0.15–0.4cm'),
    (3, '页面竖直中轴上出现两个朝下的箭头形状，水平中心均位于页面从左至右45%–55%，顶端分别位于从上至下32%–40%和53%–60%，宽1–2cm、高0.8–2cm，橙色系填充'),
    (3, '页面从左至右0%–32%、从上至下15%–45%范围内出现一张信息卡片形状：宽5–8cm、高4–5cm，白色或接近白色填充，橙色边线且线宽≥1.5磅，轮廓四角为圆角'),
    (3, '页面从左至右68%–100%、从上至下15%–45%范围内出现一张信息卡片形状：宽5–8cm、高4–5cm，白色或接近白色填充，橙色边线且线宽≥1.5磅，轮廓四角为圆角'),
    (3, '页面从左至右0%–32%、从上至下55%–88%范围内出现一张信息卡片形状：宽5–8cm、高4–5cm，白色或接近白色填充，橙色边线且线宽≥1.5磅，轮廓四角为圆角'),
    (3, '页面从左至右68%–100%、从上至下55%–88%范围内出现一张信息卡片形状：宽5–8cm、高4–5cm，白色或接近白色填充，橙色边线且线宽≥1.5磅，轮廓四角为圆角'),
    (3, '四张信息卡片内分别出现"资源整合""教学支持""活动设计""评价反馈"标题文本：橙色系、加粗、14–20磅'),
    (1, '四张信息卡片的标题文本下方各出现一条橙色横线，长3–6cm，线的中部有一个橙色实心小圆点，直径0.15–0.4cm'),
    (1, '"资源整合"卡片内出现"图书资源""社区支持""数字素材"三行条目文本：深灰色或黑色、11–16磅'),
    (1, '"教学支持"卡片内出现"分层任务""情境引导""跨学科融合"三行条目文本：深灰色或黑色、11–16磅'),
    (1, '"活动设计"卡片内出现"晨读打卡""主题书展""读书分享"三行条目文本：深灰色或黑色、11–16磅'),
    (1, '"评价反馈"卡片内出现"过程记录""成果展示""数据改进"三行条目文本：深灰色或黑色、11–16磅'),
    (3, '"资源整合"卡片内出现一个橙色系实心圆形，直径1.2–1.8cm，圆内为白色书本图标，图标由图片或形状构成而非文字字符'),
    (3, '"教学支持"卡片内出现一个橙色系实心圆形，直径1.2–1.8cm，圆内为白色学士帽图标，图标由图片或形状构成而非文字字符'),
    (3, '"活动设计"卡片内出现一个橙色系实心圆形，直径1.2–1.8cm，圆内为白色日历图标，图标由图片或形状构成而非文字字符'),
    (3, '"评价反馈"卡片内出现一个橙色系实心圆形，直径1.2–1.8cm，圆内为白色柱状图图标，图标由图片或形状构成而非文字字符'),
    (5, '页面中部围绕圆角矩形出现6–8条椭圆弧线，均为无填充或白色填充、橙色系边线的开口弧段而非闭合椭圆：弧线左右两两成对，或由左上、左下、右上、右下四条四分之一弧拼合，共同构成3条完整的扁宽椭圆环，长轴水平、短轴竖直；每条拼合后的椭圆宽高比≥2.0，宽度大于圆角矩形的宽度且左右两端伸出圆角矩形的左右边线之外，上下边缘越过圆角矩形的上下边线；3条椭圆环中心的横坐标相互偏差≤1cm、纵坐标相互偏差≤1cm，呈同心嵌套；全部弧线的整体外接矩形位于页面从左至右18%–82%、从上至下28%–70%'),
    (3, '页面中部的椭圆环线中有2条为虚线或点线边线'),
    (3, '页面左右两侧各出现2条水平方向的双向箭头线，两端均带箭头、中间的连接线为虚线或点线样式，长度均在1.5–4cm之间，4条均完整可见、未被其他形状遮挡'),
    (3, '页面左右两侧各出现1条竖直方向的双向箭头线，两端均带箭头、中间的连接线为虚线或点线样式，长度均在2–4cm之间'),
    (3, '四张信息卡片各向页面中部引出一条带转折的折线，橙色系细线，与卡片和椭圆环均相接'),
    (3, '四张信息卡片引出的折线与椭圆环相接的一端各出现一个橙色系实心小圆点，直径0.15–0.5cm'),
    (3, '页面底部出现文本"图3-2 数字化校园阅读项目实施的三维协同模型"：黑色、12–18磅，水平居中，顶端位于页面从上至下90%–99%'),
    (3, '页面中所有形状的宽度与高度均为正值，无悬空、断开、与任何元素都不相接的孤立线段'),
    (3, '页面下部朝下箭头为一个完整闭合的箭头轮廓，箭身内部无横贯的深色接缝线'),
    (3, '页面中部的椭圆环线均无旋转，环线之间无斜向交叉'),
    (1, '页面中所有形状均未添加外部阴影效果'),
    (1, '页面中所有文本框均未启用"缩放文字以适应形状"，所有文字字号均为整数磅值'),
]
_MAX_DELTAS_ORDER = [d for d, _ in _DIM2_RULES]
assert sum(_MAX_DELTAS_ORDER) == MAX_SCORE_TOTAL

# ---------------------------------------------------------------- 基础常量
EMU_PER_CM = 360000.0
_A = "http://schemas.openxmlformats.org/drawingml/2006/main"
_P = "http://schemas.openxmlformats.org/presentationml/2006/main"


def _aq(tag):
    return "{%s}%s" % (_A, tag)


def _pq(tag):
    return "{%s}%s" % (_P, tag)


# 圆角类预设几何（用于"轮廓四角为圆角"判定）
_ROUND_PRESETS = {
    "roundRect", "round1Rect", "round2SameRect", "round2DiagRect",
    "flowChartTerminator", "flowChartAlternateProcess", "ellipse",
    "flowChartConnector",
}
# 虚线/点线类 prstDash 取值
_DASH_VALUES = {
    "dash", "dashDot", "lgDash", "lgDashDot", "lgDashDotDot",
    "sysDash", "sysDot", "sysDashDot", "sysDashDotDot", "dot",
}
# 有效箭头端点类型（a:headEnd/a:tailEnd 的 type）
_ARROW_ENDS = {"triangle", "stealth", "arrow", "diamond", "oval"}


# ---------------------------------------------------------------- 颜色工具
def _theme_scheme(prs):
    """读取主题 clrScheme，返回 {'accent2': 'EE822F', ...}（大写十六进制）。

    与办公软件对齐：主题色板存放在 ppt/theme/themeN.xml 的 a:clrScheme 中，
    经 slide_master 的 /theme 关系可达；sysClr 取其 lastClr 缓存值。
    """
    out = {}
    try:
        from lxml import etree
        master = prs.slides[0].slide_layout.slide_master
        for rel in master.part.rels.values():
            if not rel.reltype.endswith("/theme"):
                continue
            root = etree.fromstring(rel.target_part.blob)
            cs = root.find(".//" + _aq("clrScheme"))
            if cs is None:
                break
            for child in cs:
                name = child.tag.split("}")[-1]
                srgb = child.find(_aq("srgbClr"))
                sysc = child.find(_aq("sysClr"))
                if srgb is not None:
                    out[name] = (srgb.get("val") or "000000").upper()
                elif sysc is not None:
                    out[name] = (sysc.get("lastClr") or "000000").upper()
            break
    except Exception:
        return out
    return out


def _apply_mods(rgb, node):
    """把 lumMod/lumOff/satMod/satOff（HSL 空间）与 shade/tint（RGB 空间）作用到基色上。"""
    try:
        r, g, b = (int(rgb[i:i + 2], 16) / 255.0 for i in (0, 2, 4))
    except Exception:
        return rgb
    h, l, s = colorsys.rgb_to_hls(r, g, b)

    def _val(tag):
        n = node.find(_aq(tag))
        if n is None:
            return None
        v = n.get("val")
        return (int(v) / 100000.0) if v is not None else None

    lum_mod, lum_off = _val("lumMod"), _val("lumOff")
    sat_mod, sat_off = _val("satMod"), _val("satOff")
    shade, tint = _val("shade"), _val("tint")
    if lum_mod is not None:
        l *= lum_mod
    if lum_off is not None:
        l += lum_off
    if sat_mod is not None:
        s *= sat_mod
    if sat_off is not None:
        s += sat_off
    l = max(0.0, min(1.0, l))
    s = max(0.0, min(1.0, s))
    r, g, b = colorsys.hls_to_rgb(h, l, s)
    if shade is not None:
        r, g, b = r * shade, g * shade, b * shade
    if tint is not None:
        r = r * tint + (1 - tint)
        g = g * tint + (1 - tint)
        b = b * tint + (1 - tint)

    def _c(x):
        return max(0, min(255, int(round(x * 255))))

    return "%02X%02X%02X" % (_c(r), _c(g), _c(b))


def _color_of(node, scheme):
    """解析 a:srgbClr / a:schemeClr 容器节点为十六进制串；解析不出返回 None。"""
    if node is None:
        return None
    srgb = node.find(_aq("srgbClr"))
    if srgb is not None:
        return _apply_mods((srgb.get("val") or "000000").upper(), srgb)
    sch = node.find(_aq("schemeClr"))
    if sch is not None:
        val = sch.get("val") or ""
        key = {"tx1": "dk1", "bg1": "lt1", "tx2": "dk2", "bg2": "lt2"}.get(val, val)
        base = scheme.get(key) or scheme.get(val) or "000000"
        return _apply_mods(base, sch)
    return None


def _rgb(hexstr):
    return (int(hexstr[0:2], 16), int(hexstr[2:4], 16), int(hexstr[4:6], 16))


def _hue_sat(hexstr):
    """返回 (色相角度 0–360, 饱和度 0–1)；灰阶返回 (None, 0.0)。"""
    r, g, b = (c / 255.0 for c in _rgb(hexstr))
    h, l, s = colorsys.rgb_to_hls(r, g, b)
    if s <= 0.0:
        return (None, 0.0)
    return (h * 360.0, s)


def _is_orange(hexstr, h_lo=18.0, h_hi=46.0, s_min=0.20):
    """橙色系：HSL 色相在 h_lo–h_hi 之间且饱和度足够（覆盖 #E07000–#FAD7A8 一族）。"""
    if not hexstr or len(hexstr) != 6:
        return False
    hue, sat = _hue_sat(hexstr)
    if hue is None or sat < s_min:
        return False
    return h_lo <= hue <= h_hi


def _is_whiteish(hexstr, min_ch=235, max_spread=22):
    """白色或接近白色。"""
    if not hexstr or len(hexstr) != 6:
        return False
    r, g, b = _rgb(hexstr)
    return min(r, g, b) >= min_ch and (max(r, g, b) - min(r, g, b)) <= max_spread


def _is_darkish(hexstr, max_ch=110):
    """深灰色或黑色。"""
    if not hexstr or len(hexstr) != 6:
        return False
    r, g, b = _rgb(hexstr)
    return max(r, g, b) <= max_ch


def _is_deep_orange(hexstr):
    """深橙色或棕色，近似 #8A4600–#C87000。"""
    if not hexstr or len(hexstr) != 6:
        return False
    r, g, b = _rgb(hexstr)
    hue, sat = _hue_sat(hexstr)
    if hue is None or not (15.0 <= hue <= 45.0) or sat < 0.30:
        return False
    return 125 <= r <= 215 and 55 <= g <= 125 and b <= 50


# ---------------------------------------------------------------- 形状读取
def _sp_pr(el):
    """<p:spPr>（PresentationML 命名空间），图片为 <p:spPr> 亦同。"""
    node = el.find(_pq("spPr"))
    if node is None:
        node = el.find(_aq("spPr"))
    return node


def _style_ref(el, tag):
    st = el.find(_pq("style"))
    return st.find(_aq(tag)) if st is not None else None


def _grad_colors(grad, scheme):
    """收集 a:gradFill 各 a:gs 色标解析后的颜色。"""
    out = []
    gs_lst = grad.find(_aq("gsLst"))
    if gs_lst is None:
        return out
    for gs in gs_lst.findall(_aq("gs")):
        col = _color_of(gs, scheme)
        if col:
            out.append(col)
    return out


def _read_fill(el, scheme):
    """返回 (kind, colors)：kind ∈ none/solid/grad/blip/style/inherit。

    与办公软件对齐：优先读 p:spPr 的 a:noFill/a:solidFill/a:gradFill/a:blipFill，
    缺省时回退到 p:style/a:fillRef 的主题填充样式引用。
    """
    sp_pr = _sp_pr(el)
    if sp_pr is not None:
        if sp_pr.find(_aq("noFill")) is not None:
            return ("none", [])
        solid = sp_pr.find(_aq("solidFill"))
        if solid is not None:
            col = _color_of(solid, scheme)
            return ("solid", [col] if col else [])
        grad = sp_pr.find(_aq("gradFill"))
        if grad is not None:
            return ("grad", _grad_colors(grad, scheme))
        if sp_pr.find(_aq("blipFill")) is not None:
            return ("blip", [])
        if sp_pr.find(_aq("pattFill")) is not None:
            return ("patt", [])
    ref = _style_ref(el, "fillRef")
    if ref is not None:
        if (ref.get("idx") or "") == "0":
            return ("none", [])
        col = _color_of(ref, scheme)
        return ("style", [col] if col else [])
    return ("inherit", [])


def _read_line(el, scheme):
    """返回 dict：kind/colors/width_pt/dash/head/tail。

    与办公软件对齐：a:ln 下的 a:noFill/a:solidFill/a:gradFill 决定线色，
    w 属性为 EMU（1pt = 12700 EMU），a:prstDash 决定虚实，a:headEnd/a:tailEnd 决定箭头。
    """
    info = {"kind": "inherit", "colors": [], "width_pt": None,
            "dash": None, "head": None, "tail": None}
    sp_pr = _sp_pr(el)
    ln = sp_pr.find(_aq("ln")) if sp_pr is not None else None
    if ln is not None:
        raw_w = ln.get("w")
        if raw_w:
            try:
                info["width_pt"] = round(int(raw_w) / 12700.0, 3)
            except ValueError:
                info["width_pt"] = None
        dash = ln.find(_aq("prstDash"))
        if dash is not None:
            info["dash"] = dash.get("val")
        head = ln.find(_aq("headEnd"))
        tail = ln.find(_aq("tailEnd"))
        if head is not None:
            info["head"] = head.get("type")
        if tail is not None:
            info["tail"] = tail.get("type")
        if ln.find(_aq("noFill")) is not None:
            info["kind"] = "none"
            return info
        solid = ln.find(_aq("solidFill"))
        grad = ln.find(_aq("gradFill"))
        if solid is not None:
            col = _color_of(solid, scheme)
            info["kind"] = "solid"
            info["colors"] = [col] if col else []
            return info
        if grad is not None:
            info["kind"] = "grad"
            info["colors"] = _grad_colors(grad, scheme)
            return info
    ref = _style_ref(el, "lnRef")
    if ref is not None:
        if (ref.get("idx") or "") == "0":
            info["kind"] = "none"
            return info
        col = _color_of(ref, scheme)
        info["kind"] = "style"
        info["colors"] = [col] if col else []
    return info


class _Shape(object):
    """把一个可点选对象归约成评测所需的度量集合（长度单位 cm，字号单位 pt）。"""

    __slots__ = ("idx", "kind", "el", "left", "top", "width", "height",
                 "right", "bottom", "mid_x", "mid_y", "rot", "prst", "custom",
                 "fill_kind", "fill_colors", "line", "texts", "para_texts",
                 "has_shadow", "has_autofit", "curved", "path_pts",
                 "raw_cx", "raw_cy")

    def __init__(self, idx, shape, scheme):
        el = shape._element
        self.idx = idx
        self.el = el
        self.kind = str(getattr(shape, "shape_type", "")) or ""
        try:
            self.raw_cx = int(shape.width)
            self.raw_cy = int(shape.height)
            self.left = float(shape.left) / EMU_PER_CM
            self.top = float(shape.top) / EMU_PER_CM
            self.width = float(shape.width) / EMU_PER_CM
            self.height = float(shape.height) / EMU_PER_CM
        except Exception:
            self.raw_cx = self.raw_cy = 0
            self.left = self.top = self.width = self.height = 0.0
        # 负值 ext 在渲染中会被折向反侧，取其绝对范围用于位置判定
        self.right = self.left + abs(self.width)
        self.bottom = self.top + abs(self.height)
        if self.width < 0:
            self.left, self.right = self.left + self.width, self.left
        if self.height < 0:
            self.top, self.bottom = self.top + self.height, self.top
        self.mid_x = (self.left + self.right) / 2.0
        self.mid_y = (self.top + self.bottom) / 2.0
        try:
            self.rot = float(getattr(shape, "rotation", 0.0) or 0.0) % 360.0
        except Exception:
            self.rot = 0.0

        prst = el.find(".//" + _aq("prstGeom"))
        self.prst = prst.get("prst") if prst is not None else None
        cust = el.find(".//" + _aq("custGeom"))
        self.custom = cust is not None
        self.curved = False
        self.path_pts = 0
        if cust is not None:
            path_lst = cust.find(_aq("pathLst"))
            if path_lst is not None:
                for node in path_lst.iter():
                    tag = node.tag.split("}")[-1]
                    if tag in ("cubicBezTo", "quadBezTo", "arcTo"):
                        self.curved = True
                    if tag in ("moveTo", "lnTo", "cubicBezTo", "quadBezTo", "arcTo"):
                        self.path_pts += 1

        self.fill_kind, self.fill_colors = _read_fill(el, scheme)
        self.line = _read_line(el, scheme)
        self.has_shadow = el.find(".//" + _aq("outerShdw")) is not None
        self.has_autofit = el.find(".//" + _aq("normAutofit")) is not None

        self.texts = []       # [(text, size_pt, bold, color_hex_or_None)]
        self.para_texts = []  # 每段拼接后的文本
        if getattr(shape, "has_text_frame", False):
            for para in shape.text_frame.paragraphs:
                buf = []
                for run in para.runs:
                    size = run.font.size.pt if run.font.size is not None else None
                    rpr = run._r.find(_aq("rPr"))
                    color = None
                    if rpr is not None:
                        sf = rpr.find(_aq("solidFill"))
                        if sf is not None:
                            color = _color_of(sf, scheme)
                    self.texts.append((run.text, size, bool(run.font.bold), color))
                    buf.append(run.text)
                self.para_texts.append("".join(buf))

    # ---- 便捷判定 ----
    def fill_hexes(self):
        return [c for c in self.fill_colors if c]

    def line_hexes(self):
        return [c for c in self.line["colors"] if c]

    def is_orange_fill(self):
        return any(_is_orange(c) for c in self.fill_hexes())

    def is_orange_line(self):
        if self.line["kind"] in ("none", "inherit"):
            return False
        return any(_is_orange(c) for c in self.line_hexes())

    def is_white_fill(self):
        """noFill 落在白色版面上与白色填充视觉等价，一并视为"白色或接近白色"。"""
        if self.fill_kind == "none":
            return True
        return any(_is_whiteish(c) for c in self.fill_hexes())

    def text_all(self):
        return "".join(t[0] for t in self.texts)

    def text_flat(self):
        return re.sub(r"\s+", "", self.text_all())

    def sizes(self):
        return [t[1] for t in self.texts if t[1] is not None]

    def is_dashed(self):
        return (self.line.get("dash") or "") in _DASH_VALUES

    def both_ends_arrow(self):
        head = self.line.get("head") or "none"
        tail = self.line.get("tail") or "none"
        return head in _ARROW_ENDS and tail in _ARROW_ENDS

    def is_thin_line(self, thick=0.25):
        return min(abs(self.width), abs(self.height)) <= thick

    def overlaps(self, other, pad=0.0):
        return (self.left - pad <= other.right and other.left - pad <= self.right
                and self.top - pad <= other.bottom and other.top - pad <= self.bottom)

    def contains_point(self, x, y, pad=0.0):
        return (self.left - pad <= x <= self.right + pad
                and self.top - pad <= y <= self.bottom + pad)


def _flatten(shapes, scheme, bag, counter):
    """展开组合，逐个收集可单独点选的对象（组合本身不计入可编辑对象数）。"""
    for shape in shapes:
        rec = _Shape(counter[0], shape, scheme)
        counter[0] += 1
        if "GROUP" in rec.kind:
            _flatten(shape.shapes, scheme, bag, counter)
            continue
        bag.append(rec)


# ---------------------------------------------------------------- 主入口
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
        if not dir_path or not os.path.isdir(dir_path):
            result["status"] = "error"
            result["error"] = "目录不存在：%s" % dir_path
            return result
        cands = sorted(
            f for f in os.listdir(dir_path)
            if f.lower().endswith(".pptx") and not f.startswith("~$")
        )
        if not cands:
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


def _zero_dim2(result, reason):
    """维度1 未通过：完成度总分记 0，不再逐条罗列（明细已无判定信息）。"""
    result["dim1_pass"] = False
    result["dim1_reason"] = reason or "维度一未通过（未给出具体原因）"
    result["dim2_items"] = []
    result["total_score"] = 0
    result["max_score"] = MAX_SCORE_TOTAL
    return result


# ---------------------------------------------------------------- 结构定位
def _find_arrow(shapes, pct_x, pct_y, y_lo, y_hi, presets):
    """在指定纵向区间内定位箭头：优先单一形状，其次同色块拼合。

    返回 (ok, single, bbox, members)；bbox=(left, top, right, bottom)。
    """
    single = None
    for s in shapes:
        if not s.is_orange_fill():
            continue
        geo_ok = (s.prst in presets) or s.custom
        if not geo_ok:
            continue
        if not (6.0 <= abs(s.width) <= 12.0 and 5.0 <= abs(s.height) <= 8.0):
            continue
        if not (30.0 <= pct_x(s.left) and pct_x(s.right) <= 70.0):
            continue
        if not (y_lo <= pct_y(s.top) and pct_y(s.bottom) <= y_hi):
            continue
        if single is None or abs(s.width * s.height) > abs(single.width * single.height):
            single = s
    if single is not None:
        return (True, True, (single.left, single.top, single.right, single.bottom), [single])

    members = []
    for s in shapes:
        if not s.is_orange_fill():
            continue
        if abs(s.width * s.height) < 4.0:
            continue
        if not (30.0 <= pct_x(s.mid_x) <= 70.0):
            continue
        if not (y_lo <= pct_y(s.mid_y) <= y_hi):
            continue
        members.append(s)
    if not members:
        return (False, False, None, [])
    left = min(s.left for s in members)
    top = min(s.top for s in members)
    right = max(s.right for s in members)
    bottom = max(s.bottom for s in members)
    ok = (6.0 <= right - left <= 12.0 and 5.0 <= bottom - top <= 8.0
          and 30.0 <= pct_x(left) and pct_x(right) <= 70.0
          and y_lo <= pct_y(top) and pct_y(bottom) <= y_hi)
    return (ok, False, (left, top, right, bottom), members)


def _find_capsule(shapes, pct_x, pct_y):
    """中部圆角矩形/胶囊形：位置、尺寸、浅橙黄填充、橙色实线边框。"""
    best = None
    for s in shapes:
        rounded = (s.prst in _ROUND_PRESETS) or (s.custom and s.curved)
        if not rounded:
            continue
        if not (8.0 <= abs(s.width) <= 16.0 and 2.5 <= abs(s.height) <= 4.0):
            continue
        if not (25.0 <= pct_x(s.left) and pct_x(s.right) <= 75.0):
            continue
        if not (38.0 <= pct_y(s.top) and pct_y(s.bottom) <= 60.0):
            continue
        if best is None or abs(s.width * s.height) > abs(best.width * best.height):
            best = s
    return best


def _find_cards(shapes, pct_x, pct_y):
    """按四个象限定位信息卡片：宽5–8cm、高4–5cm、白色系填充、橙色边线。"""
    regions = [(0.0, 32.0, 15.0, 45.0), (68.0, 100.0, 15.0, 45.0),
               (0.0, 32.0, 55.0, 88.0), (68.0, 100.0, 55.0, 88.0)]
    found = []
    for x_lo, x_hi, y_lo, y_hi in regions:
        best = None
        for s in shapes:
            if not (5.0 <= abs(s.width) <= 8.0 and 4.0 <= abs(s.height) <= 5.0):
                continue
            if not s.is_white_fill() or not s.is_orange_line():
                continue
            if not (x_lo <= pct_x(s.left) and pct_x(s.right) <= x_hi):
                continue
            if not (y_lo <= pct_y(s.top) and pct_y(s.bottom) <= y_hi):
                continue
            if best is None or abs(s.width * s.height) > abs(best.width * best.height):
                best = s
        found.append(best)
    return found


def _find_rings(shapes, capsule, page_w, pct_x, pct_y, card_ids):
    """把开口弧段两两/四四拼合成扁宽椭圆环。

    与办公软件对齐：开口弧段在 OOXML 中为 a:custGeom 且路径不闭合（含 a:cubicBezTo），
    填充为 a:noFill；左右成对的弧段其 a:off/a:ext 关于版心左右镜像。
    返回 (rings, arcs)：rings 为 [(bbox, [成员...]), ...]，arcs 为全部被采用的弧段。
    """
    cands = [s for s in shapes
             if s.idx not in card_ids
             and s.fill_kind in ("none",)
             and s.is_orange_line()
             and ((s.custom and s.curved) or s.prst in ("arc", "blockArc"))
             and 15.0 <= pct_x(s.mid_x) <= 85.0
             and 25.0 <= pct_y(s.mid_y) <= 75.0]

    # 第一步：左右镜像配对
    used = set()
    pairs = []
    for i, a in enumerate(cands):
        if a.idx in used:
            continue
        best, best_err = None, None
        for b in cands[i + 1:]:
            if b.idx in used:
                continue
            err = (abs((page_w - a.right) - b.left)
                   + abs(a.top - b.top) + abs(abs(a.height) - abs(b.height)))
            if err > 1.2:
                continue
            if best_err is None or err < best_err:
                best, best_err = b, err
        if best is None:
            continue
        used.add(a.idx)
        used.add(best.idx)
        pairs.append({
            "left": min(a.left, best.left), "top": min(a.top, best.top),
            "right": max(a.right, best.right), "bottom": max(a.bottom, best.bottom),
            "members": [a, best], "merged": False,
        })

    # 第二步：上下两段（四分之一弧）纵向拼合成一条完整椭圆
    rings = []
    for i, up in enumerate(pairs):
        if up["merged"]:
            continue
        for down in pairs[i + 1:]:
            if down["merged"]:
                continue
            hi, lo = (up, down) if up["top"] <= down["top"] else (down, up)
            if hi["bottom"] > lo["top"] + 0.05:
                continue
            if abs(hi["left"] - lo["left"]) > 0.20 or abs(hi["right"] - lo["right"]) > 0.20:
                continue
            up["merged"] = True
            down["merged"] = True
            rings.append({
                "left": min(hi["left"], lo["left"]), "top": hi["top"],
                "right": max(hi["right"], lo["right"]), "bottom": lo["bottom"],
                "members": hi["members"] + lo["members"],
            })
            break
    for pair in pairs:
        if not pair["merged"]:
            rings.append({k: pair[k] for k in ("left", "top", "right", "bottom", "members")})

    # 第三步：筛掉不构成"包裹圆角矩形的扁宽椭圆环"的拼合结果
    kept = []
    for ring in rings:
        w = ring["right"] - ring["left"]
        h = ring["bottom"] - ring["top"]
        if h <= 0 or w / h < 2.0:
            continue
        if capsule is not None:
            if w <= abs(capsule.width):
                continue
            if not (ring["left"] < capsule.left and ring["right"] > capsule.right):
                continue
            if not (ring["top"] < capsule.top and ring["bottom"] > capsule.bottom):
                continue
        kept.append(ring)

    arcs = []
    for ring in kept:
        arcs.extend(ring["members"])
    return (kept, arcs)


def _find_line_with_dot(shapes, color_test, len_lo, len_hi, box, y_lo, y_hi):
    """在 box 范围、y_lo–y_hi 之间找一条横线，且其中点处有同色实心小圆点。

    与办公软件对齐：横线为极扁的 a:ln 形状，圆点为 prstGeom="ellipse" 且有 solidFill。
    返回 (line, dot) 或 (None, None)。
    """
    lines = []
    for s in shapes:
        if abs(s.height) > 0.25 or not (len_lo <= abs(s.width) <= len_hi):
            continue
        if s.line["kind"] in ("none", "inherit"):
            continue
        if not any(color_test(c) for c in s.line_hexes()):
            continue
        if box is not None and not (s.left >= box[0] - 0.3 and s.right <= box[2] + 0.3):
            continue
        if not (y_lo <= s.mid_y <= y_hi):
            continue
        lines.append(s)
    for line in lines:
        for d in shapes:
            if d.prst != "ellipse" or d.fill_kind not in ("solid", "grad", "style"):
                continue
            if not (0.15 <= abs(d.width) <= 0.40 and 0.15 <= abs(d.height) <= 0.40):
                continue
            if not any(color_test(c) for c in d.fill_hexes()):
                continue
            if abs(d.mid_x - line.mid_x) <= 0.45 and abs(d.mid_y - line.mid_y) <= 0.45:
                return (line, d)
    return (None, None)


def _text_shapes(shapes, flat_key):
    """返回文本（去空白后）包含 flat_key 的对象列表。"""
    return [s for s in shapes if flat_key in s.text_flat()]


def _first_text_shape(shapes, flat_key):
    hits = _text_shapes(shapes, flat_key)
    return hits[0] if hits else None


def _run_of(shape, flat_key):
    """取形状内命中 flat_key 的文本属性：(size, bold, color)。

    这些文字均放在专用文本框中，故按整形状聚合：字号取最大值，
    加粗取"存在加粗run"，颜色取首个显式指定的颜色（未指定表示继承，返回 None）。
    """
    if shape is None:
        return (None, False, None)
    if flat_key not in shape.text_flat():
        return (None, False, None)
    sizes = [t[1] for t in shape.texts if t[1] is not None and t[0].strip()]
    bolds = [t[2] for t in shape.texts if t[0].strip()]
    colors = [t[3] for t in shape.texts if t[3] and t[0].strip()]
    return (max(sizes) if sizes else None, any(bolds),
            colors[0] if colors else None)


# ---------------------------------------------------------------- 评测实现
_UP_ITEMS = ["核心素养导向", "阅读兴趣培养", "信息获取与表达", "校园文化共建"]
_DOWN_ITEMS = ["阅读习惯形成", "思维能力提升", "协作交流增强", "自主学习深化"]
_CARD_TITLES = ["资源整合", "教学支持", "活动设计", "评价反馈"]
_CARD_ITEMS = {
    "资源整合": ["图书资源", "社区支持", "数字素材"],
    "教学支持": ["分层任务", "情境引导", "跨学科融合"],
    "活动设计": ["晨读打卡", "主题书展", "读书分享"],
    "评价反馈": ["过程记录", "成果展示", "数据改进"],
}
_CAPTION_KEY = "数字化校园阅读项目实施的三维协同模型"


def _evaluate_impl(file_path, result):
    # 文件已找到但打不开（损坏 / 需修复 / 实际不是 pptx 包）属于维度一未通过，
    # 而不是"没评上"：status 仍为 ok，由 _zero_dim2 记 0 分并写错误摘要。
    try:
        prs = Presentation(file_path)
    except Exception as exc:
        return _zero_dim2(result, "交付文件无法正常打开（损坏或需修复）：%s: %s"
                          % (type(exc).__name__, exc))
    scheme = _theme_scheme(prs)
    page_w = float(prs.slide_width) / EMU_PER_CM
    page_h = float(prs.slide_height) / EMU_PER_CM

    def pct_x(v):
        return v / page_w * 100.0

    def pct_y(v):
        return v / page_h * 100.0

    # ============================== 维度1：可用与可修改性 ==============================
    if len(prs.slides) != 1:
        return _zero_dim2(result, "交付PPT包含%d页幻灯片，未满足只包含1页" % len(prs.slides))

    shapes = []
    _flatten(prs.slides[0].shapes, scheme, shapes, [0])
    if len(shapes) < 50:
        return _zero_dim2(result, "可单独点选、单独编辑的对象仅%d个，少于50个" % len(shapes))

    need_text = ["目标层", "成效层", "实施枢纽"] + _CARD_TITLES
    missing = [k for k in need_text if not _text_shapes(shapes, k)]
    if missing:
        return _zero_dim2(result, "缺少可编辑文本：%s" % "、".join(missing))

    page_area = page_w * page_h
    big_pic = [s for s in shapes
               if "PICTURE" in s.kind and abs(s.width * s.height) >= page_area * 0.60]
    if big_pic:
        return _zero_dim2(result, "页面主体为整张位图（图片对象覆盖版面%.0f%%）"
                          % (abs(big_pic[0].width * big_pic[0].height) / page_area * 100.0))

    result["dim1_pass"] = True
    # 约定：维度一通过时 dim1_reason 保持空串，只有未通过才写错误摘要。
    result["dim1_reason"] = ""

    # ============================== 结构定位（供多条细则复用）==============================
    up_ok, up_single, up_box, up_members = _find_arrow(
        shapes, pct_x, pct_y, 0.0, 40.0, ("upArrow",))
    dn_ok, dn_single, dn_box, dn_members = _find_arrow(
        shapes, pct_x, pct_y, 58.0, 99.0, ("downArrow",))
    capsule = _find_capsule(shapes, pct_x, pct_y)
    cards = _find_cards(shapes, pct_x, pct_y)
    card_ids = set(c.idx for c in cards if c is not None)
    rings, ring_arcs = _find_rings(shapes, capsule, page_w, pct_x, pct_y, card_ids)
    ring_ids = set(a.idx for a in ring_arcs)
    if ring_arcs:
        ring_box = (min(a.left for a in ring_arcs), min(a.top for a in ring_arcs),
                    max(a.right for a in ring_arcs), max(a.bottom for a in ring_arcs))
    else:
        ring_box = None

    results = []

    # 01 +5 画面全部元素的整体外接矩形宽高比在 1.20–1.50 之间
    # ①遍历全部可点选对象的 a:off/a:ext；②取整体外接矩形；③宽高比落区间。
    # 与办公软件对齐：a:off/a:ext 即选择窗格中每个对象的位置与大小。
    all_left = min(s.left for s in shapes)
    all_top = min(s.top for s in shapes)
    all_right = max(s.right for s in shapes)
    all_bottom = max(s.bottom for s in shapes)
    bbox_w = all_right - all_left
    bbox_h = all_bottom - all_top
    ratio = (bbox_w / bbox_h) if bbox_h > 0 else 0.0
    if 1.20 <= ratio <= 1.50:
        results.append("+5：画面全部元素的整体外接矩形宽高比在1.20–1.50之间（实测%.2f）" % ratio)
    else:
        results.append("+0：画面全部元素的整体外接矩形宽高比在1.20–1.50之间（实测%.2f）" % ratio)

    # 02 +3 整体外接矩形左右空白之差 ≤1.5cm，上下空白之差 ≤1.5cm
    dx = abs(all_left - (page_w - all_right))
    dy = abs(all_top - (page_h - all_bottom))
    if dx <= 1.5 and dy <= 1.5:
        results.append("+3：整体外接矩形左右空白之差≤1.5cm、上下空白之差≤1.5cm"
                       "（实测%.2f/%.2fcm）" % (dx, dy))
    else:
        results.append("+0：整体外接矩形左右空白之差≤1.5cm、上下空白之差≤1.5cm"
                       "（实测%.2f/%.2fcm）" % (dx, dy))

    # 03 +5 承担主色的形状填充与边线均为橙色系（色相20°–45°）
    # ①取上下两个箭头的主体形状与中部胶囊；②填充与边线颜色均需落在橙色色相区间。
    # 与办公软件对齐：a:solidFill/a:gradFill 为填充，a:ln/a:solidFill 为边线颜色。
    main_shapes = []
    if up_members:
        main_shapes.append(max(up_members, key=lambda s: abs(s.width * s.height)))
    if dn_members:
        main_shapes.append(max(dn_members, key=lambda s: abs(s.width * s.height)))
    if capsule is not None:
        main_shapes.append(capsule)
    main_ok = bool(main_shapes) and len(main_shapes) >= 3
    for s in main_shapes:
        if not s.is_orange_line():
            main_ok = False
        if s is not capsule and not s.is_orange_fill():
            main_ok = False
    hues = []
    for s in main_shapes:
        for c in s.fill_hexes() + s.line_hexes():
            hue, sat = _hue_sat(c)
            if hue is not None and sat >= 0.20:
                hues.append(hue)
    hue_txt = ("%.0f°–%.0f°" % (min(hues), max(hues))) if hues else "无有彩色"
    results.append(("+5：" if main_ok else "+0：")
                   + "承担主色的形状填充与边线均为橙色系（实测色相%s）" % hue_txt)

    # 04 +5 页面上部朝上箭头：位置 30%–70% / 0%–40%，宽6–12cm、高5–8cm，橙色填充
    if up_ok and up_box is not None:
        results.append("+5：页面上部朝上箭头位置与尺寸达标（%.2f×%.2fcm，x%%%.1f–%.1f，"
                       "y%%%.1f–%.1f）" % (up_box[2] - up_box[0], up_box[3] - up_box[1],
                                          pct_x(up_box[0]), pct_x(up_box[2]),
                                          pct_y(up_box[1]), pct_y(up_box[3])))
    else:
        results.append("+0：页面上部未找到符合位置与尺寸要求的橙色朝上箭头")

    # 05 +3 朝上箭头内的目标层文本：白色、加粗、18–28磅，中心45%–55%，顶端3%–20%
    goal = _first_text_shape(shapes, "目标层")
    g_size, g_bold, g_color = _run_of(goal, "目标层")
    goal_ok = (goal is not None and g_size is not None and 18.0 <= g_size <= 28.0
               and g_bold and g_color is not None and _is_whiteish(g_color)
               and 45.0 <= pct_x(goal.mid_x) <= 55.0
               and 3.0 <= pct_y(goal.top) <= 20.0)
    results.append(("+3：" if goal_ok else "+0：")
                   + "朝上箭头内目标层文本为白色加粗18–28磅且位置达标（实测%s磅，中心x%%%.1f，"
                     "顶端y%%%.1f）" % (g_size, pct_x(goal.mid_x) if goal else 0.0,
                                     pct_y(goal.top) if goal else 0.0))

    # 06 +3 朝上箭身内自上而下四行条目：白色、12–18磅、水平居中，整体 y% 12–36
    up_item_shapes = []
    for key in _UP_ITEMS:
        hit = None
        for s in _text_shapes(shapes, key):
            if up_box is None or (s.left >= up_box[0] - 2.0 and s.top >= up_box[1] - 0.5):
                hit = s
                break
        up_item_shapes.append(hit)
    up_items_ok = all(s is not None for s in up_item_shapes)
    if up_items_ok:
        tops = [s.top for s in up_item_shapes]
        up_items_ok = tops == sorted(tops)
        for s, key in zip(up_item_shapes, _UP_ITEMS):
            size, _bold, color = _run_of(s, key)
            if size is None or not (12.0 <= size <= 18.0):
                up_items_ok = False
            if color is None or not _is_whiteish(color):
                up_items_ok = False
            if not (s.left <= page_w * 0.52 and s.right >= page_w * 0.48):
                up_items_ok = False
        band = (pct_y(min(s.top for s in up_item_shapes)),
                pct_y(max(s.bottom for s in up_item_shapes)))
        if not (12.0 <= band[0] and band[1] <= 36.0):
            up_items_ok = False
        detail = "整体y%%%.1f–%.1f" % band
    else:
        detail = "条目文本缺失"
    results.append(("+3：" if up_items_ok else "+0：")
                   + "朝上箭身内四行条目各占一行、白色12–18磅、水平居中且整体位于"
                     "y%%12–36（%s）" % detail)

    # 07 +1 目标层与四行条目之间：白色横线长3–7cm，线中部白色实心小圆点直径0.15–0.4cm
    if goal is not None and up_item_shapes[0] is not None:
        y_lo = goal.bottom - 0.35
        y_hi = up_item_shapes[0].top + 0.35
        w_line, w_dot = _find_line_with_dot(shapes, _is_whiteish, 3.0, 7.0,
                                            up_box, y_lo, y_hi)
    else:
        w_line = w_dot = None
    results.append(("+1：" if (w_line is not None and w_dot is not None) else "+0：")
                   + "目标层与四行条目之间有白色横线（长3–7cm）且线中部有白色实心小圆点"
                     "（直径0.15–0.4cm）")

    # 08 +5 页面下部朝下箭头：位置 30%–70% / 58%–99%，宽6–12cm、高5–8cm，橙色填充
    if dn_ok and dn_box is not None:
        results.append("+5：页面下部朝下箭头位置与尺寸达标（%.2f×%.2fcm，x%%%.1f–%.1f，"
                       "y%%%.1f–%.1f）" % (dn_box[2] - dn_box[0], dn_box[3] - dn_box[1],
                                          pct_x(dn_box[0]), pct_x(dn_box[2]),
                                          pct_y(dn_box[1]), pct_y(dn_box[3])))
    else:
        results.append("+0：页面下部未找到符合位置与尺寸要求的橙色朝下箭头")

    # 09 +3 朝下箭头内的成效层文本：白色、加粗、18–28磅，中心45%–55%，顶端62%–72%
    eff = _first_text_shape(shapes, "成效层")
    e_size, e_bold, e_color = _run_of(eff, "成效层")
    eff_ok = (eff is not None and e_size is not None and 18.0 <= e_size <= 28.0
              and e_bold and e_color is not None and _is_whiteish(e_color)
              and 45.0 <= pct_x(eff.mid_x) <= 55.0
              and 62.0 <= pct_y(eff.top) <= 72.0)
    results.append(("+3：" if eff_ok else "+0：")
                   + "朝下箭头内成效层文本为白色加粗18–28磅且位置达标（实测%s磅，中心x%%%.1f，"
                     "顶端y%%%.1f）" % (e_size, pct_x(eff.mid_x) if eff else 0.0,
                                     pct_y(eff.top) if eff else 0.0))

    # 10 +3 朝下箭身内自上而下四行条目：白色、12–18磅、水平居中，整体 y% 69–92
    dn_item_shapes = []
    for key in _DOWN_ITEMS:
        hit = None
        for s in _text_shapes(shapes, key):
            if dn_box is None or (s.left >= dn_box[0] - 2.0 and s.bottom <= dn_box[3] + 0.5):
                hit = s
                break
        dn_item_shapes.append(hit)
    dn_items_ok = all(s is not None for s in dn_item_shapes)
    if dn_items_ok:
        tops = [s.top for s in dn_item_shapes]
        dn_items_ok = tops == sorted(tops)
        for s, key in zip(dn_item_shapes, _DOWN_ITEMS):
            size, _bold, color = _run_of(s, key)
            if size is None or not (12.0 <= size <= 18.0):
                dn_items_ok = False
            if color is None or not _is_whiteish(color):
                dn_items_ok = False
            if not (s.left <= page_w * 0.52 and s.right >= page_w * 0.48):
                dn_items_ok = False
        band = (pct_y(min(s.top for s in dn_item_shapes)),
                pct_y(max(s.bottom for s in dn_item_shapes)))
        if not (69.0 <= band[0] and band[1] <= 92.0):
            dn_items_ok = False
        detail = "整体y%%%.1f–%.1f" % band
    else:
        detail = "条目文本缺失"
    results.append(("+3：" if dn_items_ok else "+0：")
                   + "朝下箭身内四行条目各占一行、白色12–18磅、水平居中且整体位于"
                     "y%%69–92（%s）" % detail)

    # 11 +1 成效层与四行条目之间：白色横线长3–7cm，线中部白色实心小圆点直径0.15–0.4cm
    if eff is not None and dn_item_shapes[0] is not None:
        y_lo = eff.bottom - 0.35
        y_hi = dn_item_shapes[0].top + 0.35
        w_line2, w_dot2 = _find_line_with_dot(shapes, _is_whiteish, 3.0, 7.0,
                                             dn_box, y_lo, y_hi)
    else:
        w_line2 = w_dot2 = None
    results.append(("+1：" if (w_line2 is not None and w_dot2 is not None) else "+0：")
                   + "成效层与四行条目之间有白色横线（长3–7cm）且线中部有白色实心小圆点"
                     "（直径0.15–0.4cm）")

    # 12 +3 两处箭身内四行条目行间各3条白色虚线分隔线，合计6条
    # ①虚线/点线 a:prstDash；②线宽≤1.5磅；③长度3–5cm且小于箭身宽度；④中心 x% 45–55。
    def _seps(box):
        out = []
        if box is None:
            return out
        shaft_w = box[2] - box[0]
        for s in shapes:
            if abs(s.height) > 0.25 or not (3.0 <= abs(s.width) <= 5.0):
                continue
            if not s.is_dashed():
                continue
            wpt = s.line.get("width_pt")
            if wpt is not None and wpt > 1.5:
                continue
            if not any(_is_whiteish(c) for c in s.line_hexes()):
                continue
            if not (45.0 <= pct_x(s.mid_x) <= 55.0):
                continue
            if not (box[1] <= s.mid_y <= box[3]):
                continue
            if abs(s.width) >= shaft_w:
                continue
            out.append(s)
        return out

    sep_up = _seps(up_box)
    sep_dn = _seps(dn_box)
    sep_ok = len(sep_up) == 3 and len(sep_dn) == 3
    results.append(("+3：" if sep_ok else "+0：")
                   + "两处箭身内四行条目行间各有3条白色虚线/点线分隔线，合计6条"
                     "（实测上%d条、下%d条）" % (len(sep_up), len(sep_dn)))

    # 13 +5 中部圆角矩形/胶囊形：位置、尺寸、浅橙黄填充、橙色实线边框≥1.5磅
    cap_ok = False
    cap_detail = "未找到符合位置与尺寸的圆角矩形/胶囊形"
    if capsule is not None:
        fills = capsule.fill_hexes()
        fill_ok = False
        for c in fills:
            r, g, b = _rgb(c)
            if r >= 240 and 200 <= g <= 245 and 140 <= b <= 235:
                fill_ok = True
                break
        wpt = capsule.line.get("width_pt")
        line_ok = (capsule.is_orange_line()
                   and (capsule.line.get("dash") in (None, "solid"))
                   and wpt is not None and wpt >= 1.5)
        cap_ok = fill_ok and line_ok
        cap_detail = "%.2f×%.2fcm，填充%s，边线%s/%s磅" % (
            abs(capsule.width), abs(capsule.height), ",".join(fills) or "无",
            ",".join(capsule.line_hexes()) or "无", wpt)
    results.append(("+5：" if cap_ok else "+0：")
                   + "中部圆角矩形/胶囊形位置尺寸达标、浅橙黄填充、橙色实线边框≥1.5磅"
                     "（%s）" % cap_detail)

    # 14 +3 圆角矩形内上部的实施枢纽文本：深橙或棕色、加粗、18–26磅，中心 45%–55%
    hub = None
    for s in _text_shapes(shapes, "实施枢纽"):
        if capsule is None or s.overlaps(capsule, 0.6):
            hub = s
            break
    h_size, h_bold, h_color = _run_of(hub, "实施枢纽")
    hub_ok = (hub is not None and h_size is not None and 18.0 <= h_size <= 26.0
              and h_bold and h_color is not None and _is_deep_orange(h_color)
              and 45.0 <= pct_x(hub.mid_x) <= 55.0)
    results.append(("+3：" if hub_ok else "+0：")
                   + "圆角矩形内实施枢纽文本为深橙/棕色加粗18–26磅且水平中心45%%–55%%"
                     "（实测%s磅，%s）" % (h_size, h_color or "未显式设色"))

    # 15 +1 圆角矩形内下部单行文本：课程内容·技术平台·活动组织，12–18磅、水平居中
    sub = None
    for s in shapes:
        flat = s.text_flat()
        if "课程内容" in flat and "技术平台" in flat and "活动组织" in flat:
            if capsule is None or s.overlaps(capsule, 0.8):
                sub = s
                break
    sub_ok = False
    sub_detail = "未找到该单行文本"
    if sub is not None:
        flat = re.sub(r"[\s]+", "", sub.text_all())
        sizes = sub.sizes()
        lines = [p for p in sub.para_texts if p.strip()]
        sub_ok = (flat == "课程内容·技术平台·活动组织" and len(lines) == 1
                  and bool(sizes) and all(12.0 <= v <= 18.0 for v in sizes)
                  and 42.0 <= pct_x(sub.mid_x) <= 58.0)
        sub_detail = "%s，%s磅" % (flat, sizes[0] if sizes else "未设")
    results.append(("+1：" if sub_ok else "+0：")
                   + "圆角矩形内下部单行文本为课程内容·技术平台·活动组织且12–18磅居中"
                     "（%s）" % sub_detail)

    # 16 +1 实施枢纽与下部单行文本之间：橙色横线长6–13cm，线中部橙色实心小圆点
    if hub is not None and sub is not None:
        o_line, o_dot = _find_line_with_dot(shapes, _is_orange, 6.0, 13.0, None,
                                           hub.bottom - 0.35, sub.top + 0.35)
    else:
        o_line = o_dot = None
    results.append(("+1：" if (o_line is not None and o_dot is not None) else "+0：")
                   + "圆角矩形内实施枢纽与单行文本之间有橙色横线（长6–13cm）且线中部有"
                     "橙色实心小圆点（直径0.15–0.4cm）")

    # 17 +3 竖直中轴两个朝下小箭头：中心 x% 45–55，顶端 y% 32–40 与 53–60
    mid_arrows = []
    for s in shapes:
        if s.prst != "downArrow" and not (s.custom and abs(s.width) <= 2.0):
            continue
        if not (1.0 <= abs(s.width) <= 2.0 and 0.8 <= abs(s.height) <= 2.0):
            continue
        if not (45.0 <= pct_x(s.mid_x) <= 55.0):
            continue
        if not s.is_orange_fill():
            continue
        mid_arrows.append(s)
    tops_pct = sorted(pct_y(s.top) for s in mid_arrows)
    mid_ok = (len(mid_arrows) == 2
              and 32.0 <= tops_pct[0] <= 40.0 and 53.0 <= tops_pct[1] <= 60.0)
    results.append(("+3：" if mid_ok else "+0：")
                   + "竖直中轴上有两个朝下小箭头（宽1–2cm、高0.8–2cm、橙色填充，顶端"
                     "y%%32–40与53–60；实测%d个%s）"
                   % (len(mid_arrows), ["%.1f" % v for v in tops_pct]))

    # 18–21 +3×4 四张信息卡片：宽5–8cm、高4–5cm、白色系填充、橙色边线≥1.5磅、四角圆角
    # 与办公软件对齐：圆角由 prstGeom 的圆角预设或 custGeom 中的 a:cubicBezTo 拐角实现。
    card_names = ["左上", "右上", "左下", "右下"]
    card_zones = ["x%0–32/y%15–45", "x%68–100/y%15–45",
                  "x%0–32/y%55–88", "x%68–100/y%55–88"]
    card_pass = []
    for card, name, zone in zip(cards, card_names, card_zones):
        ok = False
        detail = "未找到符合位置与尺寸的卡片形状"
        if card is not None:
            rounded = (card.prst in _ROUND_PRESETS) or (card.custom and card.curved)
            wpt = card.line.get("width_pt")
            ok = bool(rounded) and wpt is not None and wpt >= 1.5
            detail = "%.2f×%.2fcm，几何%s，边线%s磅，%s" % (
                abs(card.width), abs(card.height), card.prst or "custGeom", wpt,
                "四角圆角" if rounded else "四角非圆角")
        card_pass.append(ok)
        results.append(("+3：" if ok else "+0：")
                       + "%s范围内有一张信息卡片（宽5–8cm、高4–5cm、白色系填充、橙色边线"
                         "≥1.5磅、四角圆角）——%s（%s）" % (zone, name, detail))

    # 卡片与标题的归属关系（供 22–31 复用）
    title_card = {}
    title_shape = {}
    for key in _CARD_TITLES:
        shp = None
        for s in _text_shapes(shapes, key):
            if len(s.text_flat()) <= 12:
                shp = s
                break
        title_shape[key] = shp
        owner = None
        if shp is not None:
            for card in cards:
                if card is not None and card.contains_point(shp.mid_x, shp.mid_y, 0.6):
                    owner = card
                    break
        title_card[key] = owner

    # 22 +3 四张卡片标题文本：橙色系、加粗、14–20磅
    title_ok = True
    t_detail = []
    for key in _CARD_TITLES:
        shp = title_shape[key]
        size, bold, color = _run_of(shp, key)
        good = (shp is not None and size is not None and 14.0 <= size <= 20.0
                and bold and color is not None and _is_orange(color))
        if not good:
            title_ok = False
        t_detail.append("%s:%s磅/%s" % (key, size, color or "未显式设色"))
    results.append(("+3：" if title_ok else "+0：")
                   + "四张卡片标题文本均为橙色系、加粗、14–20磅（%s）" % "；".join(t_detail))

    # 23 +1 四张卡片标题下方各有橙色横线长3–6cm，线中部橙色实心小圆点
    line_hits = 0
    for key in _CARD_TITLES:
        card = title_card[key]
        shp = title_shape[key]
        if card is None or shp is None:
            continue
        box = (card.left, card.top, card.right, card.bottom)
        l_line, l_dot = _find_line_with_dot(shapes, _is_orange, 3.0, 6.0, box,
                                           shp.bottom - 0.35, card.bottom)
        if l_line is not None and l_dot is not None:
            line_hits += 1
    results.append(("+1：" if line_hits == 4 else "+0：")
                   + "四张卡片标题文本下方各有一条橙色横线（长3–6cm）且线中部有橙色实心"
                     "小圆点（直径0.15–0.4cm；实测%d/4）" % line_hits)

    # 24–27 +1×4 四张卡片内各三行条目：深灰色或黑色、11–16磅
    for key in _CARD_TITLES:
        card = title_card[key]
        items = _CARD_ITEMS[key]
        hit, sizes_seen = 0, []
        for item in items:
            for s in _text_shapes(shapes, item):
                if card is not None and not card.contains_point(s.left, s.top, 0.8):
                    continue
                size, _bold, color = _run_of(s, item)
                if size is None or not (11.0 <= size <= 16.0):
                    continue
                if color is not None and not _is_darkish(color):
                    continue
                sizes_seen.append(size)
                hit += 1
                break
        ok = (hit == 3)
        results.append(("+1：" if ok else "+0：")
                       + "%s卡片内有%s三行条目文本（深灰或黑色、11–16磅；实测%d/3，%s磅）"
                       % (key, "/".join(items), hit,
                          sizes_seen[0] if sizes_seen else "未设"))

    # 28–31 +3×4 四张卡片内橙色实心圆（直径1.2–1.8cm）+ 圆内白色图标（图片或形状）
    icon_labels = {"资源整合": "书本", "教学支持": "学士帽",
                   "活动设计": "日历", "评价反馈": "柱状图"}
    for key in _CARD_TITLES:
        card = title_card[key]
        circle = None
        if card is not None:
            for s in shapes:
                if s.prst != "ellipse" or s.idx == card.idx:
                    continue
                if not (1.2 <= abs(s.width) <= 1.8 and 1.2 <= abs(s.height) <= 1.8):
                    continue
                if not s.is_orange_fill():
                    continue
                if card.contains_point(s.mid_x, s.mid_y, 0.3):
                    circle = s
                    break
        icon = None
        if circle is not None:
            for s in shapes:
                if s.idx == circle.idx:
                    continue
                if s.text_flat():
                    continue
                is_pic = "PICTURE" in s.kind
                is_shape = (s.prst is not None or s.custom)
                if not (is_pic or is_shape):
                    continue
                if abs(s.width) > abs(circle.width) or abs(s.height) > abs(circle.height):
                    continue
                if circle.contains_point(s.mid_x, s.mid_y, 0.05):
                    icon = s
                    break
        ok = circle is not None and icon is not None
        results.append(("+3：" if ok else "+0：")
                       + "%s卡片内有橙色系实心圆（直径1.2–1.8cm）且圆内白色%s图标由图片或"
                         "形状构成而非文字字符（%s）"
                       % (key, icon_labels[key],
                          ("圆径%.2fcm/图标%s" % (abs(circle.width), icon.kind.split()[0]))
                          if ok else "未同时找到实心圆与非文字图标"))

    # 32 +5 中部 6–8 条开口弧线拼合成 3 条同心扁宽椭圆环
    ring_ok = False
    ring_detail = "未拼合出合格椭圆环"
    if len(rings) == 3 and 6 <= len(ring_arcs) <= 8 and ring_box is not None:
        cxs = [(r["left"] + r["right"]) / 2.0 for r in rings]
        cys = [(r["top"] + r["bottom"]) / 2.0 for r in rings]
        concentric = (max(cxs) - min(cxs) <= 1.0) and (max(cys) - min(cys) <= 1.0)
        in_zone = (18.0 <= pct_x(ring_box[0]) and pct_x(ring_box[2]) <= 82.0
                   and 28.0 <= pct_y(ring_box[1]) and pct_y(ring_box[3]) <= 70.0)
        ring_ok = concentric and in_zone
        ring_detail = ("%d条弧线拼成3条椭圆环，宽高比%s，中心偏差%.2f/%.2fcm，"
                       "整体x%%%.1f–%.1f、y%%%.1f–%.1f"
                       % (len(ring_arcs),
                          "/".join("%.2f" % ((r["right"] - r["left"])
                                             / max(r["bottom"] - r["top"], 1e-6))
                                   for r in rings),
                          max(cxs) - min(cxs), max(cys) - min(cys),
                          pct_x(ring_box[0]), pct_x(ring_box[2]),
                          pct_y(ring_box[1]), pct_y(ring_box[3])))
    elif rings:
        ring_detail = "仅拼合出%d条椭圆环（弧线%d条）" % (len(rings), len(ring_arcs))
    results.append(("+5：" if ring_ok else "+0：")
                   + "中部6–8条无填充橙色开口弧线两两/四四拼合成3条同心扁宽椭圆环并包裹"
                     "圆角矩形（%s）" % ring_detail)

    # 33 +3 椭圆环线中有 2 条为虚线或点线边线
    dashed_arcs = [a for a in ring_arcs if a.is_dashed()]
    results.append(("+3：" if len(dashed_arcs) >= 2 else "+0：")
                   + "中部椭圆环线中有2条为虚线或点线边线（实测%d条）" % len(dashed_arcs))

    # 34 +3 左右各 2 条水平双向箭头线：两端箭头、虚线连接线、长1.5–4cm、完整可见
    def _covered(target):
        for s in shapes:
            if s.idx <= target.idx or s.fill_kind not in ("solid", "grad", "blip", "style"):
                continue
            if (s.left <= target.left and s.right >= target.right
                    and s.top <= target.top and s.bottom >= target.bottom):
                return True
        return False

    h_arrows = [s for s in shapes
                if s.both_ends_arrow() and s.is_dashed()
                and abs(s.height) <= 0.25 and 1.5 <= abs(s.width) <= 4.0]
    h_left = [s for s in h_arrows if pct_x(s.mid_x) < 50.0]
    h_right = [s for s in h_arrows if pct_x(s.mid_x) >= 50.0]
    h_ok = (len(h_left) == 2 and len(h_right) == 2
            and not any(_covered(s) for s in h_arrows))
    results.append(("+3：" if h_ok else "+0：")
                   + "页面左右两侧各有2条水平双向虚线箭头线（长1.5–4cm且完整可见；"
                     "实测左%d条、右%d条）" % (len(h_left), len(h_right)))

    # 35 +3 左右各 1 条竖直双向箭头线：两端箭头、虚线连接线、长2–4cm
    v_arrows = [s for s in shapes
                if s.both_ends_arrow() and s.is_dashed()
                and abs(s.width) <= 0.25 and 2.0 <= abs(s.height) <= 4.0]
    v_left = [s for s in v_arrows if pct_x(s.mid_x) < 50.0]
    v_right = [s for s in v_arrows if pct_x(s.mid_x) >= 50.0]
    v_ok = len(v_left) == 1 and len(v_right) == 1
    results.append(("+3：" if v_ok else "+0：")
                   + "页面左右两侧各有1条竖直双向虚线箭头线（长2–4cm；实测左%d条、右%d条）"
                   % (len(v_left), len(v_right)))

    # 36 +3 四张卡片各向页面中部引出一条带转折的橙色细折线，与卡片和椭圆环均相接
    connectors = []
    for card in cards:
        if card is None or ring_box is None:
            connectors.append(None)
            continue
        hit = None
        for s in shapes:
            if s.idx in card_ids or s.idx in ring_ids:
                continue
            if not (s.custom and s.path_pts >= 3):
                continue
            if s.fill_kind not in ("none", "inherit") or not s.is_orange_line():
                continue
            if not card.overlaps(s, 0.3):
                continue
            if not (s.left <= ring_box[2] + 0.3 and s.right >= ring_box[0] - 0.3
                    and s.top <= ring_box[3] + 0.3 and s.bottom >= ring_box[1] - 0.3):
                continue
            hit = s
            break
        connectors.append(hit)
    conn_n = sum(1 for c in connectors if c is not None)
    results.append(("+3：" if conn_n == 4 else "+0：")
                   + "四张信息卡片各向页面中部引出一条带转折的橙色细折线且与卡片和椭圆环"
                     "均相接（实测%d/4）" % conn_n)

    # 37 +3 四条折线与椭圆环相接的一端各有一个橙色实心小圆点（直径0.15–0.5cm）
    dot_n = 0
    for conn in connectors:
        if conn is None:
            continue
        for s in shapes:
            if s.prst != "ellipse" or s.fill_kind not in ("solid", "grad", "style"):
                continue
            if not (0.15 <= abs(s.width) <= 0.5 and 0.15 <= abs(s.height) <= 0.5):
                continue
            if not s.is_orange_fill():
                continue
            if conn.overlaps(s, 0.3):
                dot_n += 1
                break
    results.append(("+3：" if dot_n == 4 else "+0：")
                   + "四条折线与椭圆环相接的一端各有一个橙色实心小圆点（直径0.15–0.5cm；"
                     "实测%d/4）" % dot_n)

    # 38 +3 底部图注文本：黑色、12–18磅、水平居中，顶端 y% 90–99
    cap = None
    for s in _text_shapes(shapes, _CAPTION_KEY):
        cap = s
        break
    cap_ok2 = False
    cap_detail2 = "未找到底部图注文本"
    if cap is not None:
        flat = cap.text_flat()
        sizes = cap.sizes()
        _size, _bold, color = _run_of(cap, _CAPTION_KEY)
        cap_ok2 = (flat.startswith("图3-2") and bool(sizes)
                   and all(12.0 <= v <= 18.0 for v in sizes)
                   and (color is None or _is_darkish(color))
                   and 42.0 <= pct_x(cap.mid_x) <= 58.0
                   and 90.0 <= pct_y(cap.top) <= 99.0)
        cap_detail2 = "%s磅，中心x%%%.1f，顶端y%%%.1f" % (
            sizes[0] if sizes else "未设", pct_x(cap.mid_x), pct_y(cap.top))
    results.append(("+3：" if cap_ok2 else "+0：")
                   + "底部图注文本为图3-2 数字化校园阅读项目实施的三维协同模型且黑色"
                     "12–18磅居中、顶端y%%90–99（%s）" % cap_detail2)

    # 39 +3 所有形状宽高为正值，无与任何元素都不相接的孤立线段
    neg = [s for s in shapes if s.raw_cx < 0 or s.raw_cy < 0]
    isolated = []
    for s in shapes:
        if not s.is_thin_line():
            continue
        if not any(o.idx != s.idx and s.overlaps(o, 1.0) for o in shapes):
            isolated.append(s)
    size_ok = not neg and not isolated
    results.append(("+3：" if size_ok else "+0：")
                   + "所有形状宽高均为正值且无悬空断开的孤立线段（负尺寸%d个、孤立线段%d条）"
                   % (len(neg), len(isolated)))

    # 40 +3 朝下箭头为一个完整闭合轮廓，箭身内无横贯的深色接缝线
    seam = None
    if dn_box is not None:
        for s in shapes:
            if s.idx in ring_ids or not s.is_thin_line():
                continue
            if abs(s.height) > 0.2 or abs(s.width) < (dn_box[2] - dn_box[0]) * 0.6:
                continue
            if not (dn_box[0] - 0.2 <= s.mid_x <= dn_box[2] + 0.2
                    and dn_box[1] <= s.mid_y <= dn_box[3]):
                continue
            hexes = s.line_hexes()
            if any(_is_darkish(h, 150) or _is_orange(h) for h in hexes):
                seam = s
                break
    dn_solid = bool(dn_ok) and bool(dn_single) and seam is None
    results.append(("+3：" if dn_solid else "+0：")
                   + "页面下部朝下箭头为一个完整闭合的箭头轮廓且箭身内无横贯的深色接缝线"
                     "（%s）" % ("单一闭合形状" if dn_solid
                                 else ("由%d个形状拼接" % len(dn_members)
                                       if not dn_single and dn_members else "存在横贯接缝线")))

    # 41 +3 椭圆环线均无旋转，环线之间无斜向交叉
    ring_pool = list(ring_arcs)
    if not ring_pool:
        for s in shapes:
            if s.prst != "ellipse" and not (s.custom and s.curved):
                continue
            if not (s.is_white_fill() or s.fill_kind in ("none", "inherit")):
                continue
            if abs(s.width) < 3.0 or not (25.0 <= pct_y(s.mid_y) <= 75.0):
                continue
            if not (35.0 <= pct_x(s.mid_x) <= 65.0):
                continue
            ring_pool.append(s)
    rotated = [s for s in ring_pool if abs(((s.rot or 0.0) + 180.0) % 360.0 - 180.0) > 1.0]
    rot_ok = bool(ring_pool) and not rotated
    results.append(("+3：" if rot_ok else "+0：")
                   + "页面中部的椭圆环线均无旋转、环线之间无斜向交叉（%s）"
                   % ("环线%d条旋转角均为0" % len(ring_pool) if rot_ok
                      else ("%d条环线存在旋转（%s）"
                            % (len(rotated), "/".join("%.0f°" % s.rot for s in rotated[:4]))
                            if rotated else "未找到椭圆环线")))

    # 42 +1 所有形状均未添加外部阴影效果
    shadowed = [s for s in shapes if s.has_shadow]
    results.append(("+1：" if not shadowed else "+0：")
                   + "页面中所有形状均未添加外部阴影效果（带阴影%d个）" % len(shadowed))

    # 43 +1 未启用缩放文字以适应形状，且所有字号为整数磅值
    autofit = [s for s in shapes if s.has_autofit]
    frac = []
    for s in shapes:
        for v in s.sizes():
            if abs(v - round(v)) > 1e-6:
                frac.append(v)
    results.append(("+1：" if not autofit and not frac else "+0：")
                   + "所有文本框均未启用缩放文字以适应形状且字号均为整数磅值"
                     "（自动缩放%d处、非整数字号%d处）" % (len(autofit), len(frac)))

    if len(results) != len(_DIM2_RULES):
        raise RuntimeError("dim2 项数不一致：results=%d vs rules=%d"
                           % (len(results), len(_DIM2_RULES)))
    dim2_items = []
    total_score = 0
    # results 里的文字只用于本地排查；返回信息中的 rule 一律取细则原文，
    # 只从 "+N" 前缀取命中与否，保证同一条细则命中/未命中读起来完全一致。
    for raw, (max_delta, rule) in zip(results, _DIM2_RULES):
        head, _, _ = raw.partition("：")
        try:
            delta = int(head.replace("+", "").strip())
        except ValueError:
            delta = 0
        hit = delta > 0
        dim2_items.append({"rule": rule, "max_delta": max_delta,
                           "delta": max_delta if hit else 0, "hit": hit})
        total_score += (max_delta if hit else 0)
    result["dim2_items"] = dim2_items
    result["total_score"] = total_score
    result["max_score"] = MAX_SCORE_TOTAL
    return result


if __name__ == "__main__":
    target = sys.argv[1] if len(sys.argv) > 1 else os.path.dirname(os.path.abspath(__file__))
    print(json.dumps(evaluate(target), ensure_ascii=False, indent=2))
