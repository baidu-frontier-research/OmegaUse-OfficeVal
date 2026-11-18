#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# pyright: reportUnknownMemberType=false, reportUnknownVariableType=false, reportUnknownArgumentType=false, reportMissingTypeStubs=false, reportPrivateUsage=false, reportUnknownParameterType=false, reportUnusedVariable=false
"""053 电化学分析系统原理图 PPT 完成度评测。

按接口统一约定，本模块不修改全局 sys.stdout，也不 print 主结果。主结果统一由
evaluate(dir_path) -> dict 返回；入参为“脚本所在目录的路径”，由脚本自己在该目录内
定位并打开被评估的 .pptx。本文件仅在 __main__ 下作为本地调试入口，才把 dict 序列化
为 JSON 打印。

返回信息约定：
  * status="error" 表示这份交付根本没能进入评分（目录不存在、目录中无 .pptx、脚本异常），
    此时 dim1_pass 为空串 ""（未判定）、dim2_items 为 []、total_score 与 max_score 均为 0。
  * dim1_pass 为 False 时 dim1_reason 必须带错误摘要；为 True 时 dim1_reason 为空串。
  * dim2_items 每项只输出 rule / max_delta / delta / hit 四个键，rule 取细则原文，
    delta 只取 0 或 max_delta 两个值。
"""

from __future__ import annotations

import json
import os
import re
import sys
import zipfile
from dataclasses import dataclass, field
from typing import Any, Callable, Optional

SCRIPT_ID = "053"
MAX_SCORE_TOTAL = 86

EMU_PER_CM = 360000.0
EMU_PER_PT = 12700.0
# <a:ln> 未显式给出 w 时，PowerPoint 按 1 磅（细线）渲染，评分须按 1.0 磅判定。
DEFAULT_LINE_WIDTH_PT = 1.0
# 未显式给出 sz 的 <a:rPr>，PowerPoint 按 18 磅默认字号渲染。
DEFAULT_FONT_SIZE_PT = 18.0
# 百分比区间容差：0.2 个百分点以内视为落在区间内的取整误差。
PCT_TOL = 0.2
# 颜色“近似”判定：RGB 三通道最大偏差阈值。
RGB_TOL = 36
RGB_TOL_LOOSE = 48

NS = {
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
}
A = "{%s}" % NS["a"]
P = "{%s}" % NS["p"]

# 维度2 细则：(满分, 细则原文)，顺序与《评测细则.txt》完全一致。
_DIM2_RULES: list[tuple[int, str]] = [
    (5, "幻灯片上全部图形与文字构成的整体内容区域宽高比为1.63–1.73（即近似1.68:1），未被拉伸为与16:9页面一致的1.78:1"),
    (5, "幻灯片上全部图形与文字构成的整体内容区域在页面上居中放置：左右两侧留白差值不超过整体宽度的3%，上下两侧留白差值不超过整体高度的5%"),
    (3, "幻灯片出现一个无填充的矩形边框：边线为蓝色单实线，线宽2–3.5磅；左边线位于页面从左至右1%–5%、右边线位于95%–99%，上边线位于页面从上至下0%–3%、下边线位于97%–101%"),
    (3, "幻灯片在页面从左至右14%–36%、从上至下4%–14%的位置出现“分析物输入”文本：颜色为蓝色（近似#557CB6），字号19–23磅"),
    (1, "幻灯片在页面从左至右19%–27%、从上至下10%–21%的位置出现一条竖直向下的蓝色虚线箭头：线型为虚线，末端带箭头"),
    (5, "幻灯片在页面从左至右9%–33%、从上至下39%–74%的位置出现一个烧杯图形：由杯身轮廓、内部液体色块与气泡组成，整体宽度5.5–7.5cm，高度4.5–6.5cm"),
    (3, "幻灯片在页面从左至右11%–30%、从上至下47%–72%的位置出现一块浅蓝色液体色块：为渐变或纯色浅蓝（近似#B6E2FD），顶部呈水平液面形态"),
    (1, "幻灯片在页面从左至右13%–28%、从上至下56%–69%的范围内出现7个及以上圆形气泡形状：直径0.15–0.55cm，填充为白色至浅蓝的渐变或浅蓝色，位置分散不重叠堆叠"),
    (3, "幻灯片在页面从左至右11%–19%、从上至下38%–62%的位置出现一根绿色电极杆图形：杆体高度3–4.5cm、宽度0.3–1.5cm；沿长度方向分为上、中、下三段及以上——上段为绿色帽套（近似#5EAD5F），中段为白色或灰白色的细杆（肉眼呈白色玻璃管质感，不是绿色的浅色调），下段末端收窄为尖端而非平直矩形端"),
    (3, "幻灯片在页面从左至右17%–25%、从上至下23%–66%的位置出现一根竖直柱体形状：柱身为白色或浅色，顶端带深色帽盖（近似#2E3F58），帽盖位于页面从上至下23%–32%的位置"),
    (3, "幻灯片在页面从左至右22%–30%、从上至下38%–62%的位置出现一根红色电极杆图形：杆体高度3–4.5cm、宽度0.3–1.5cm；沿长度方向分为上、中、下三段及以上——上段为红色帽套（近似#B90303或#EA0000），中段为白色或灰白色的细杆（肉眼呈白色玻璃管质感，不是红色的浅色调），下段末端收窄为尖端而非平直矩形端"),
    (1, "幻灯片在页面从左至右2%–15%、从上至下47%–57%的位置出现“参比电极”文本：字号13–17磅"),
    (1, "幻灯片在页面从左至右16%–30%、从上至下44%–53%的位置出现“工作电极”文本：字号13–17磅"),
    (1, "幻灯片在页面从左至右28%–39%、从上至下47%–57%的位置出现“对电极”文本：字号13–17磅"),
    (1, "幻灯片在页面从左至右12%–28%、从上至下62%–71%的位置出现“样品溶液”文本：字号13–17磅"),
    (1, "幻灯片在页面从上至下49%–56%的位置出现两条水平的黑色实线箭头，长度均为0.5–1.5cm：一条位于页面从左至右10%–16%、箭头头部在其右端（自左向右指向电极），另一条位于页面从左至右26%–32%、箭头头部在其左端（自右向左指向电极）"),
    (5, "幻灯片在页面从左至右37%–66%、从上至下44%–66%的位置出现一个斜视立体外观的机箱图形：由四块色块拼出厚度——浅灰色顶面（近似#C7D0E0）为高度不超过1.2cm的扁平斜置四边形，深蓝色前壳（近似#233959）为正面块，前壳之内叠放亮蓝色面板（近似#1777D5），前壳右边线以右另有一块颜色比顶面更深的灰蓝色右侧面色块、其高度在2.5cm以上；顶面右端与右侧面色块右端均比前壳右边线向右多伸出1cm以上"),
    (3, "幻灯片在页面从左至右39%–48%、从上至下51%–62%的位置出现一个小屏幕矩形：为白色填充或白色边框，内部有波形折线"),
    (1, "幻灯片在页面从左至右53%–59%、从上至下49%–62%的位置出现3个圆形指示灯：直径0.15–0.45cm，自上而下填充分别为绿色、深灰蓝色、红色"),
    (1, "幻灯片在页面从左至右44%–58%、从上至下52%–60%的位置出现“电化学分析仪”文本：白色，字号11–13磅，非加粗"),
    (3, "幻灯片在页面从左至右46%–65%、从上至下30%–40%的位置出现一个无填充、蓝色虚线边框的圆角矩形，其内为“信号采集与处理”文本：文本字号13–16磅"),
    (3, "幻灯片在页面从左至右41%–58%、从上至下68%–79%的位置出现一个无填充、蓝色虚线边框的圆角矩形，其内为“电位/电流控制”文本：文本字号14–18磅"),
    (3, "幻灯片在页面从左至右73%–93%、从上至下18%–31%的位置出现一个无填充、蓝色虚线边框的圆角矩形，其内为“数据分析与可视化”文本：文本字号13–19磅"),
    (1, "幻灯片出现3条竖直的蓝色虚线箭头：分别位于页面从左至右47%–53%且从上至下63%–71%、从左至右53%–59%且从上至下38%–48%、从左至右80%–86%且从上至下26%–35%的位置；线型均为虚线且末端带箭头"),
    (5, "幻灯片在页面从左至右9%–48%、从上至下21%–51%的区域内出现三条平滑弯曲的曲线：绿色曲线（近似#5EAD5F）位于页面从左至右9%–43%、从上至下32%–51%，红色曲线（近似#B90303）位于页面从左至右26%–46%、从上至下31%–49%，黑色曲线（近似#151310）位于页面从左至右20%–48%、从上至下21%–48%；三条曲线线宽均为1.5–3磅"),
    (1, "幻灯片在页面从左至右62%–74%、从上至下50%–58%的位置出现一条水平的黑色线段及位于其右端的小矩形块：矩形块宽度0.4–1cm"),
    (5, "幻灯片在页面从左至右65%–95%、从上至下32%–68%的位置出现一个笔记本电脑图形：由深蓝色屏幕外框（近似#2B415C）、白色屏幕区、下方底座与触控板构成，底座上出现键盘按键线条或按键形状"),
    (3, "幻灯片在页面从左至右72%–83%、从上至下39%–54%的位置出现一条蓝色平滑曲线或折线（近似#3276C3）：该线由一个连续形状实现、可整体选中，未由多段各自独立的直线段首尾拼接而成，线条本身不出现断开的缺口；其外围有浅色矩形图表边框"),
    (3, "幻灯片在页面从左至右81%–92%、从上至下40%–55%的位置出现一条红色平滑曲线或折线（近似#E56F6E或#E10D0D）：该线由一个连续形状实现、可整体选中，未由多段各自独立的直线段首尾拼接而成，线条本身不出现断开的缺口；其外围有浅色矩形图表边框"),
    (1, "幻灯片在页面从左至右73%–85%、从上至下64%–72%的位置出现“计算机”文本：字号13–17磅"),
    (3, "幻灯片在页面从左至右9%–91%、从上至下80%–96%的位置出现一个无填充的蓝色虚线圆角矩形：宽度25–30cm，高度2–3cm"),
    (1, "幻灯片在页面从左至右12%–18%、从上至下84%–88%的位置出现一条水平的黑色短线段（线宽1.5–2.5磅），其右侧出现“工作电极连接线（信号输出）”文本：文本字号12–15磅"),
    (1, "幻灯片在页面从左至右12%–18%、从上至下89%–93%的位置出现一条水平的绿色短线段（线宽1.5–2.5磅），其右侧出现“参比电极连接线（电位参照）”文本：文本字号12–15磅"),
    (1, "幻灯片在页面从左至右41%–47%、从上至下85%–89%的位置出现一条水平的红色短线段（线宽1.5–2.5磅），其右侧出现“对电极连接线（电流通路）”文本：文本字号12–15磅"),
    (1, "幻灯片在页面从左至右67%–74%、从上至下85%–89%的位置出现一条水平的蓝色虚线箭头（线宽2–3磅），其右侧出现“信号/控制/数据流向”文本：文本字号12–15磅"),
    (1, "幻灯片内文字均完整显示在所属形状的可视区域之内，未被其他形状遮挡、未被裁切；未添加标题、页码、日期、说明性注释等要求之外的文字"),
]

_MAX_DELTAS_ORDER: list[int] = [d for d, _ in _DIM2_RULES]
assert sum(_MAX_DELTAS_ORDER) == MAX_SCORE_TOTAL, (
    "细则满分之和与 MAX_SCORE_TOTAL 不一致：%d vs %d" % (sum(_MAX_DELTAS_ORDER), MAX_SCORE_TOTAL)
)

# ---------------------------------------------------------------- 颜色工具 ----

def _hex2rgb(h: str) -> tuple[int, int, int]:
    h = (h or "").strip().lstrip("#")
    if len(h) != 6:
        return (0, 0, 0)
    try:
        return (int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))
    except ValueError:
        return (0, 0, 0)


def _rgb2hex(rgb: tuple[float, float, float]) -> str:
    return "%02X%02X%02X" % tuple(max(0, min(255, int(round(c)))) for c in rgb)


def _near(color: Optional[str], target: str, tol: int = RGB_TOL) -> bool:
    """三通道最大偏差不超过 tol 即视为“近似”该颜色。"""
    if not color:
        return False
    a, b = _hex2rgb(color), _hex2rgb(target)
    return max(abs(a[i] - b[i]) for i in range(3)) <= tol


def _near_any(colors: list[str], target: str, tol: int = RGB_TOL) -> bool:
    return any(_near(c, target, tol) for c in colors)


def _lum(color: str) -> float:
    r, g, b = _hex2rgb(color)
    return 0.299 * r + 0.587 * g + 0.114 * b


def _is_blue(color: Optional[str]) -> bool:
    if not color:
        return False
    r, g, b = _hex2rgb(color)
    return b >= 90 and b > r + 30 and b > g + 10


def _is_green(color: Optional[str]) -> bool:
    if not color:
        return False
    r, g, b = _hex2rgb(color)
    return g > r + 15 and g > b + 15


def _is_red(color: Optional[str]) -> bool:
    if not color:
        return False
    r, g, b = _hex2rgb(color)
    return r >= 120 and r > g + 50 and r > b + 50


def _is_black(color: Optional[str]) -> bool:
    if not color:
        return False
    r, g, b = _hex2rgb(color)
    return max(r, g, b) <= 70


def _is_whiteish(color: Optional[str]) -> bool:
    """肉眼呈白色/灰白：三通道都很亮且彼此接近（浅色调彩色不算）。"""
    if not color:
        return False
    r, g, b = _hex2rgb(color)
    return min(r, g, b) >= 190 and (max(r, g, b) - min(r, g, b)) <= 45


def _is_light(color: Optional[str]) -> bool:
    if not color:
        return False
    return min(_hex2rgb(color)) >= 195


def _is_dark_slate(color: Optional[str]) -> bool:
    """深灰蓝：整体偏暗且蓝分量不低于红分量。"""
    if not color:
        return False
    r, g, b = _hex2rgb(color)
    return max(r, g, b) <= 150 and b >= r


class ThemeColors:
    """解析 theme1.xml 的 clrScheme 与 slideMaster 的 clrMap，用于还原 schemeClr。"""

    def __init__(self, pptx_path: str) -> None:
        self.scheme: dict[str, str] = {}
        self.clr_map: dict[str, str] = {}
        try:
            with zipfile.ZipFile(pptx_path) as zf:
                names = zf.namelist()
                themes = sorted(n for n in names if re.match(r"ppt/theme/theme\d+\.xml$", n))
                if themes:
                    self._parse_scheme(zf.read(themes[0]).decode("utf-8", "ignore"))
                masters = sorted(n for n in names if re.match(r"ppt/slideMasters/slideMaster\d+\.xml$", n))
                if masters:
                    self._parse_map(zf.read(masters[0]).decode("utf-8", "ignore"))
        except Exception:
            pass

    def _parse_scheme(self, xml: str) -> None:
        m = re.search(r"<a:clrScheme.*?</a:clrScheme>", xml, re.S)
        if not m:
            return
        for blk in re.finditer(r"<a:(\w+)>(.*?)</a:\1>", m.group(0), re.S):
            key, body = blk.group(1), blk.group(2)
            srgb = re.search(r'<a:srgbClr val="([0-9A-Fa-f]{6})"', body)
            if srgb:
                self.scheme[key] = srgb.group(1).upper()
                continue
            sysc = re.search(r'<a:sysClr[^>]*lastClr="([0-9A-Fa-f]{6})"', body)
            if sysc:
                self.scheme[key] = sysc.group(1).upper()

    def _parse_map(self, xml: str) -> None:
        m = re.search(r"<p:clrMap([^>]*)/>", xml)
        if not m:
            return
        for k, v in re.findall(r'(\w+)="(\w+)"', m.group(1)):
            self.clr_map[k] = v

    def resolve(self, val: str) -> Optional[str]:
        key = self.clr_map.get(val, val)
        if key == "phClr":
            key = "accent1"
        if key in ("tx1", "dk1"):
            return self.scheme.get("dk1", "000000")
        if key in ("bg1", "lt1"):
            return self.scheme.get("lt1", "FFFFFF")
        if key in ("tx2", "dk2"):
            return self.scheme.get("dk2", "44546A")
        if key in ("bg2", "lt2"):
            return self.scheme.get("lt2", "E7E6E6")
        return self.scheme.get(key)


def _apply_mods(base: str, node: Any) -> str:
    """按 lumMod/lumOff/shade/tint 近似换算最终 RGB（用于色系判定，非精确色管理）。"""
    r, g, b = (float(c) for c in _hex2rgb(base))
    for child in list(node):
        tag = child.tag.split("}")[-1]
        raw = child.get("val")
        if raw is None:
            continue
        try:
            v = int(raw) / 100000.0
        except ValueError:
            continue
        if tag in ("lumMod", "shade"):
            r, g, b = r * v, g * v, b * v
        elif tag == "lumOff":
            r = r + 255.0 * v
            g = g + 255.0 * v
            b = b + 255.0 * v
        elif tag == "tint":
            r = r * v + 255.0 * (1 - v)
            g = g * v + 255.0 * (1 - v)
            b = b * v + 255.0 * (1 - v)
    return _rgb2hex((r, g, b))


# ------------------------------------------------------------ 填充/线条解析 ----

def _colors_of(node: Any, theme: ThemeColors) -> tuple[list[str], str, int]:
    """返回 (颜色列表, 类型, 最小 alpha 千分比)。类型取 solid/grad/none/other/inherit。"""
    if node is None:
        return ([], "inherit", 100000)
    grad = node.find(A + "gradFill")
    if grad is not None:
        cols: list[str] = []
        alpha = 100000
        for gs in grad.iter(A + "gs"):
            c, a = _one_color(gs, theme)
            if c:
                cols.append(c)
            alpha = min(alpha, a)
        return (cols, "grad", alpha)
    solid = node.find(A + "solidFill")
    if solid is None and node.tag.endswith("solidFill"):
        solid = node
    if solid is not None:
        c, a = _one_color(solid, theme)
        return ([c] if c else [], "solid", a)
    if node.find(A + "noFill") is not None:
        return ([], "none", 0)
    if node.find(A + "grpFill") is not None:
        return ([], "grpFill", 100000)
    if node.find(A + "blipFill") is not None or node.find(A + "pattFill") is not None:
        return ([], "other", 100000)
    return ([], "inherit", 100000)


def _one_color(holder: Any, theme: ThemeColors) -> tuple[Optional[str], int]:
    alpha = 100000
    srgb = holder.find(".//" + A + "srgbClr")
    scheme = holder.find(".//" + A + "schemeClr")
    node = srgb if srgb is not None else scheme
    if node is None:
        return (None, alpha)
    an = node.find(A + "alpha")
    if an is not None and an.get("val"):
        try:
            alpha = int(an.get("val"))
        except ValueError:
            alpha = 100000
    if srgb is not None:
        return (_apply_mods(srgb.get("val") or "000000", srgb).upper(), alpha)
    base = theme.resolve(scheme.get("val") or "")
    if base is None:
        return (None, alpha)
    return (_apply_mods(base, scheme).upper(), alpha)


def _style_ref_color(el: Any, ref_tag: str, theme: ThemeColors) -> Optional[str]:
    """<p:style> 里 lnRef/fillRef 引用的主题色，作为 spPr 未显式给色时的回退。"""
    style = el.find(P + "style")
    if style is None:
        return None
    ref = style.find(A + ref_tag)
    if ref is None:
        return None
    c, _ = _one_color(ref, theme)
    return c


# ---------------------------------------------------------------- 形状模型 ----

@dataclass
class RunInfo:
    size_pt: float
    bold: bool
    color: Optional[str]


@dataclass
class ShapeInfo:
    z: int = 0
    name: str = ""
    prst: str = "-"
    x0: float = 0.0
    y0: float = 0.0
    x1: float = 0.0
    y1: float = 0.0
    raw_cx: float = 0.0
    raw_cy: float = 0.0
    rot: float = 0.0
    flip_h: bool = False
    flip_v: bool = False
    fill_colors: list[str] = field(default_factory=list)
    fill_kind: str = "inherit"
    fill_alpha: int = 100000
    line_colors: list[str] = field(default_factory=list)
    line_kind: str = "inherit"
    line_w_pt: float = DEFAULT_LINE_WIDTH_PT
    dash: str = "solid"
    head_end: Optional[str] = None
    tail_end: Optional[str] = None
    seg_line: int = 0
    seg_bez: int = 0
    seg_move: int = 0
    is_text: bool = False
    text: str = ""
    runs: list[RunInfo] = field(default_factory=list)
    is_picture: bool = False

    # --- 便捷属性（单位换算与形态判定）---
    @property
    def w(self) -> float:
        return self.x1 - self.x0

    @property
    def h(self) -> float:
        return self.y1 - self.y0

    @property
    def w_cm(self) -> float:
        return self.w / EMU_PER_CM

    @property
    def h_cm(self) -> float:
        return self.h / EMU_PER_CM

    @property
    def fill_main(self) -> Optional[str]:
        return self.fill_colors[0] if self.fill_colors else None

    @property
    def line_main(self) -> Optional[str]:
        return self.line_colors[0] if self.line_colors else None

    @property
    def opaque_fill(self) -> bool:
        return self.fill_kind in ("solid", "grad", "other") and self.fill_alpha >= 60000

    @property
    def filled(self) -> bool:
        return self.fill_kind in ("solid", "grad") and self.fill_alpha >= 20000

    @property
    def is_line_like(self) -> bool:
        return self.prst in ("line", "straightConnector1", "bentConnector3", "curvedConnector3")

    @property
    def has_arrow_end(self) -> bool:
        return bool(self._end(self.head_end)) or bool(self._end(self.tail_end))

    @staticmethod
    def _end(v: Optional[str]) -> bool:
        return bool(v) and v != "none"


@dataclass
class SlideModel:
    page_w: float = 12192000.0
    page_h: float = 6858000.0
    shapes: list[ShapeInfo] = field(default_factory=list)
    group_count: int = 0
    slide_count: int = 0

    def px(self, v: float) -> float:
        return v / self.page_w * 100.0

    def py(self, v: float) -> float:
        return v / self.page_h * 100.0

    def texts(self) -> list[ShapeInfo]:
        return [s for s in self.shapes if s.is_text and s.text.strip()]


def _grp_xfrm(el: Any) -> Optional[tuple[float, ...]]:
    xf = el.find(P + "grpSpPr")
    xf = xf.find(A + "xfrm") if xf is not None else None
    if xf is None:
        return None
    off, ext = xf.find(A + "off"), xf.find(A + "ext")
    coff, cext = xf.find(A + "chOff"), xf.find(A + "chExt")
    if off is None or ext is None or coff is None or cext is None:
        return None
    return (
        float(off.get("x") or 0), float(off.get("y") or 0),
        float(ext.get("cx") or 0), float(ext.get("cy") or 0),
        float(coff.get("x") or 0), float(coff.get("y") or 0),
        float(cext.get("cx") or 1), float(cext.get("cy") or 1),
    )


def _compose(parent: Callable[..., tuple[float, float, float, float]], el: Any):
    g = _grp_xfrm(el)
    if g is None:
        return parent
    ox, oy, cx, cy, hx, hy, hcx, hcy = g
    sx = cx / (hcx or 1.0)
    sy = cy / (hcy or 1.0)

    def mapper(x: float, y: float, w: float, h: float):
        return parent(ox + (x - hx) * sx, oy + (y - hy) * sy, w * sx, h * sy)

    return mapper


def _identity(x: float, y: float, w: float, h: float):
    return (x, y, w, h)


def _spPr_of(el: Any) -> Any:
    node = el.find(P + "spPr")
    if node is None:
        node = el.find(P + "grpSpPr")
    return node


def _read_shape(sh: Any, mapper: Any, theme: ThemeColors, grp_fill: tuple[list[str], str, int],
                z: int) -> ShapeInfo:
    el = sh._element
    info = ShapeInfo(z=z, name=str(sh.name or ""))
    left = float(sh.left or 0)
    top = float(sh.top or 0)
    width = float(sh.width or 0)
    height = float(sh.height or 0)
    x, y, w, h = mapper(left, top, width, height)
    info.raw_cx, info.raw_cy = w, h
    # ext 为负值不符合 ST_PositiveCoordinate，PowerPoint 会按绝对值修复后渲染，
    # 故此处统一取规范化后的包围盒。
    info.x0, info.x1 = min(x, x + w), max(x, x + w)
    info.y0, info.y1 = min(y, y + h), max(y, y + h)

    geom = el.find(".//" + A + "prstGeom")
    if geom is not None:
        info.prst = geom.get("prst") or "-"
    elif el.find(".//" + A + "custGeom") is not None:
        info.prst = "cust"
    cg = el.find(".//" + A + "custGeom")
    if cg is not None:
        info.seg_line = len(list(cg.iter(A + "lnTo")))
        info.seg_bez = len(list(cg.iter(A + "cubicBezTo"))) + len(list(cg.iter(A + "quadBezTo")))
        info.seg_move = len(list(cg.iter(A + "moveTo")))

    spPr = _spPr_of(el)
    xf = spPr.find(A + "xfrm") if spPr is not None else None
    if xf is not None:
        if xf.get("rot"):
            info.rot = int(xf.get("rot")) / 60000.0
        info.flip_h = xf.get("flipH") in ("1", "true")
        info.flip_v = xf.get("flipV") in ("1", "true")

    cols, kind, alpha = _colors_of(spPr, theme)
    if kind == "grpFill":
        cols, kind, alpha = grp_fill[0], grp_fill[1], grp_fill[2]
    if kind == "inherit":
        ref = _style_ref_color(el, "fillRef", theme)
        if ref:
            cols, kind, alpha = [ref], "solid", 100000
    info.fill_colors, info.fill_kind, info.fill_alpha = cols, kind, alpha

    ln = spPr.find(A + "ln") if spPr is not None else None
    if ln is not None:
        lc, lk, la = _colors_of(ln, theme)
        info.line_colors, info.line_kind = lc, lk
        if ln.get("w"):
            try:
                info.line_w_pt = int(ln.get("w")) / EMU_PER_PT
            except ValueError:
                pass
        d = ln.find(A + "prstDash")
        if d is not None:
            info.dash = d.get("val") or "solid"
        for tag in ("headEnd", "tailEnd"):
            node = ln.find(A + tag)
            if node is not None:
                if tag == "headEnd":
                    info.head_end = node.get("type") or "none"
                else:
                    info.tail_end = node.get("type") or "none"
    if info.line_kind in ("inherit",) and not info.line_colors:
        ref = _style_ref_color(el, "lnRef", theme)
        if ref:
            info.line_colors, info.line_kind = [ref], "solid"

    info.is_picture = el.tag == P + "pic"
    if sh.has_text_frame:
        info.is_text = True
        info.text = sh.text_frame.text or ""
        for para in sh.text_frame.paragraphs:
            for run in para.runs:
                rPr = run._r.find(A + "rPr")
                size = DEFAULT_FONT_SIZE_PT
                bold = False
                color: Optional[str] = None
                if rPr is not None:
                    if rPr.get("sz"):
                        size = int(rPr.get("sz")) / 100.0
                    bold = rPr.get("b") in ("1", "true")
                    rc, _, _ = _colors_of(rPr, theme)
                    color = rc[0] if rc else None
                info.runs.append(RunInfo(size, bold, color))
    return info


def _walk(shapes: Any, mapper: Any, theme: ThemeColors, grp_fill: tuple[list[str], str, int],
          model: SlideModel) -> None:
    for sh in shapes:
        el = sh._element
        if el.tag == P + "grpSp":
            model.group_count += 1
            own = _colors_of(_spPr_of(el), theme)
            if own[1] in ("inherit", "grpFill"):
                own = grp_fill
            try:
                _walk(sh.shapes, _compose(mapper, el), theme, own, model)
            except Exception:
                pass
            continue
        try:
            model.shapes.append(_read_shape(sh, mapper, theme, grp_fill, len(model.shapes)))
        except Exception:
            continue


def build_model(pptx_path: str) -> SlideModel:
    from pptx import Presentation  # 延迟导入，便于在缺依赖时由 evaluate 统一兜底报错

    prs = Presentation(pptx_path)
    model = SlideModel()
    model.page_w = float(prs.slide_width or 12192000)
    model.page_h = float(prs.slide_height or 6858000)
    model.slide_count = len(prs.slides)
    if model.slide_count:
        theme = ThemeColors(pptx_path)
        _walk(prs.slides[0].shapes, _identity, theme, ([], "inherit", 100000), model)
    return model


# ------------------------------------------------------------------ 判定工具 ----

def _rng(v: float, lo: float, hi: float, tol: float = PCT_TOL) -> bool:
    return (lo - tol) <= v <= (hi + tol)


def _span(lo_v: float, hi_v: float, lo: float, hi: float, tol: float = PCT_TOL) -> bool:
    """整段（起止两端）都落在 [lo, hi] 内。"""
    return _rng(lo_v, lo, hi, tol) and _rng(hi_v, lo, hi, tol)


def _bbox(items: list[ShapeInfo]) -> tuple[float, float, float, float]:
    return (min(s.x0 for s in items), min(s.y0 for s in items),
            max(s.x1 for s in items), max(s.y1 for s in items))


def _inter_area(a: ShapeInfo, b: ShapeInfo) -> float:
    dx = min(a.x1, b.x1) - max(a.x0, b.x0)
    dy = min(a.y1, b.y1) - max(a.y0, b.y0)
    return dx * dy if dx > 0 and dy > 0 else 0.0


def _inside(inner: ShapeInfo, outer: ShapeInfo, tol: float = 0.15 * EMU_PER_CM) -> bool:
    return (inner.x0 >= outer.x0 - tol and inner.x1 <= outer.x1 + tol
            and inner.y0 >= outer.y0 - tol and inner.y1 <= outer.y1 + tol)


def _sizes(s: ShapeInfo) -> list[float]:
    return [r.size_pt for r in s.runs] or [DEFAULT_FONT_SIZE_PT]


def _norm(t: str) -> str:
    """比较文本时忽略空白（“电位 / 电流控制”与“电位/电流控制”视为同一文本）。"""
    return re.sub(r"\s+", "", t or "")


def _find_text(model: SlideModel, needle: str) -> list[ShapeInfo]:
    key = _norm(needle)
    return [s for s in model.texts() if key in _norm(s.text)]


def _text_hit(model: SlideModel, needle: str, xr: tuple[float, float], yr: tuple[float, float],
              sr: tuple[float, float]) -> tuple[bool, str]:
    """文本内容 + 位置区间 + 字号区间 的通用判定。"""
    cands = _find_text(model, needle)
    if not cands:
        return (False, "未找到包含“%s”的文本" % needle)
    notes: list[str] = []
    for s in cands:
        x0, x1 = model.px(s.x0), model.px(s.x1)
        y0, y1 = model.py(s.y0), model.py(s.y1)
        szs = _sizes(s)
        pos_ok = _span(x0, x1, xr[0], xr[1]) and _span(y0, y1, yr[0], yr[1])
        sz_ok = any(sr[0] <= v <= sr[1] for v in szs)
        if pos_ok and sz_ok:
            return (True, "%s：x%.2f%%–%.2f%% y%.2f%%–%.2f%% 字号%s磅"
                    % (s.name, x0, x1, y0, y1, "/".join("%.1f" % v for v in szs)))
        notes.append("%s：x%.2f%%–%.2f%% y%.2f%%–%.2f%% 字号%s磅（位置%s/字号%s）"
                     % (s.name, x0, x1, y0, y1, "/".join("%.1f" % v for v in szs),
                        "符合" if pos_ok else "不符", "符合" if sz_ok else "不符"))
    return (False, "; ".join(notes))


def _dashed(s: ShapeInfo) -> bool:
    return s.dash not in ("solid", "") and s.dash is not None


def _solid_line(s: ShapeInfo) -> bool:
    return s.dash in ("solid", "")


def _labeled(s: ShapeInfo) -> bool:
    return bool(s.text.strip())


def _content_items(m: SlideModel) -> list[ShapeInfo]:
    return [s for s in m.shapes if s.w > 0 or s.h > 0]


# =============================================================== 维度2 检查 ====

def check_01_aspect(m: SlideModel) -> tuple[bool, str]:
    items = _content_items(m)
    if not items:
        return (False, "幻灯片无任何形状")
    x0, y0, x1, y1 = _bbox(items)
    w_cm, h_cm = (x1 - x0) / EMU_PER_CM, (y1 - y0) / EMU_PER_CM
    if h_cm <= 0:
        return (False, "整体内容区域高度为 0")
    ratio = w_cm / h_cm
    ok = 1.63 <= ratio <= 1.73
    return (ok, "整体内容区域 %.2fcm×%.2fcm，宽高比 %.3f（要求 1.63–1.73）" % (w_cm, h_cm, ratio))


def check_02_centered(m: SlideModel) -> tuple[bool, str]:
    items = _content_items(m)
    if not items:
        return (False, "幻灯片无任何形状")
    x0, y0, x1, y1 = _bbox(items)
    left, right = x0, m.page_w - x1
    top, bottom = y0, m.page_h - y1
    w, h = x1 - x0, y1 - y0
    dx_ok = abs(left - right) <= 0.03 * w
    dy_ok = abs(top - bottom) <= 0.05 * h
    return (dx_ok and dy_ok,
            "左留白%.2fcm/右留白%.2fcm（差%.2fcm，限%.2fcm）；上留白%.2fcm/下留白%.2fcm（差%.2fcm，限%.2fcm）"
            % (left / EMU_PER_CM, right / EMU_PER_CM, abs(left - right) / EMU_PER_CM,
               0.03 * w / EMU_PER_CM, top / EMU_PER_CM, bottom / EMU_PER_CM,
               abs(top - bottom) / EMU_PER_CM, 0.05 * h / EMU_PER_CM))


def check_03_outer_frame(m: SlideModel) -> tuple[bool, str]:
    notes: list[str] = []
    for s in m.shapes:
        if s.prst != "rect" or _labeled(s):
            continue
        x0, x1 = m.px(s.x0), m.px(s.x1)
        y0, y1 = m.py(s.y0), m.py(s.y1)
        if not (_rng(x0, 1, 5) and _rng(x1, 95, 99) and _rng(y0, 0, 3) and _rng(y1, 97, 101)):
            continue
        reasons: list[str] = []
        if s.fill_kind != "none":
            reasons.append("非无填充(%s)" % s.fill_kind)
        if not _is_blue(s.line_main):
            reasons.append("边线非蓝色(%s)" % (s.line_main or "无"))
        if not _solid_line(s):
            reasons.append("边线非单实线(%s)" % s.dash)
        if not (2.0 <= s.line_w_pt <= 3.5):
            reasons.append("线宽%.2f磅不在2–3.5" % s.line_w_pt)
        desc = "%s：x%.2f%%–%.2f%% y%.2f%%–%.2f%% 边线%s %.2f磅 %s" % (
            s.name, x0, x1, y0, y1, s.line_main or "无", s.line_w_pt, s.dash)
        if not reasons:
            return (True, desc)
        notes.append(desc + "（" + "、".join(reasons) + "）")
    return (False, "; ".join(notes) or "未找到位置符合的矩形边框")


def check_04_inlet_text(m: SlideModel) -> tuple[bool, str]:
    ok, detail = _text_hit(m, "分析物输入", (14, 36), (4, 14), (19, 23))
    if not ok:
        return (False, detail)
    for s in _find_text(m, "分析物输入"):
        if any(_near(r.color, "557CB6", RGB_TOL_LOOSE) for r in s.runs):
            return (True, detail + "，颜色近似#557CB6")
    cols = [r.color or "-" for s in _find_text(m, "分析物输入") for r in s.runs]
    return (False, detail + "，但文字颜色为 %s，非蓝色#557CB6" % "/".join(cols))


def check_05_inlet_arrow(m: SlideModel) -> tuple[bool, str]:
    notes: list[str] = []
    for s in m.shapes:
        if _labeled(s) or s.h <= s.w:
            continue
        x0, x1 = m.px(s.x0), m.px(s.x1)
        y0, y1 = m.py(s.y0), m.py(s.y1)
        if not (_span(x0, x1, 19, 27) and _span(y0, y1, 10, 21)):
            continue
        reasons: list[str] = []
        if not _dashed(s):
            reasons.append("线型非虚线(%s)" % s.dash)
        if not s.has_arrow_end:
            reasons.append("末端无箭头")
        if not _is_blue(s.line_main):
            reasons.append("线条非蓝色(%s)" % (s.line_main or "无"))
        desc = "%s：x%.2f%%–%.2f%% y%.2f%%–%.2f%% %s 箭头%s/%s 颜色%s" % (
            s.name, x0, x1, y0, y1, s.dash, s.head_end or "-", s.tail_end or "-",
            s.line_main or "无")
        if not reasons:
            return (True, desc)
        notes.append(desc + "（" + "、".join(reasons) + "）")
    return (False, "; ".join(notes[:4]) or "该区域未找到竖直线条")


def _window(m: SlideModel, s: ShapeInfo, xr: tuple[float, float], yr: tuple[float, float]) -> bool:
    return (_span(m.px(s.x0), m.px(s.x1), xr[0], xr[1])
            and _span(m.py(s.y0), m.py(s.y1), yr[0], yr[1]))


def _bubbles(m: SlideModel, xr: tuple[float, float], yr: tuple[float, float]) -> list[ShapeInfo]:
    out: list[ShapeInfo] = []
    for s in m.shapes:
        if s.prst != "ellipse" or _labeled(s):
            continue
        d = max(s.w_cm, s.h_cm)
        if not (0.15 <= d <= 0.55) or abs(s.w_cm - s.h_cm) > 0.15:
            continue
        if _window(m, s, xr, yr):
            out.append(s)
    return out


def check_06_beaker(m: SlideModel) -> tuple[bool, str]:
    parts = [s for s in m.shapes
             if not _labeled(s) and not s.is_line_like and s.prst != "triangle"
             and s.w_cm >= 3.0 and _window(m, s, (7, 35), (36, 77))]
    if not parts:
        return (False, "该区域未找到烧杯轮廓部件")
    x0, y0, x1, y1 = _bbox(parts)
    w_cm, h_cm = (x1 - x0) / EMU_PER_CM, (y1 - y0) / EMU_PER_CM
    liquid = [s for s in parts if _near_any(s.fill_colors, "B6E2FD", RGB_TOL_LOOSE)]
    body = [s for s in parts if s.h_cm >= 3.0]
    bubbles = _bubbles(m, (13, 28), (56, 69))
    reasons: list[str] = []
    if not _span(m.px(x0), m.px(x1), 9, 33) or not _span(m.py(y0), m.py(y1), 39, 74):
        reasons.append("整体位置越界")
    if not (5.5 <= w_cm <= 7.5):
        reasons.append("整体宽度%.2fcm不在5.5–7.5" % w_cm)
    if not (4.5 <= h_cm <= 6.5):
        reasons.append("整体高度%.2fcm不在4.5–6.5" % h_cm)
    if not body:
        reasons.append("无杯身轮廓（高度≥3cm的部件）")
    if not liquid:
        reasons.append("无内部浅蓝液体色块")
    if len(bubbles) < 5:
        reasons.append("气泡仅%d个" % len(bubbles))
    desc = "烧杯部件%d个：x%.2f%%–%.2f%% y%.2f%%–%.2f%% %.2fcm×%.2fcm，液体%d块、气泡%d个" % (
        len(parts), m.px(x0), m.px(x1), m.py(y0), m.py(y1), w_cm, h_cm, len(liquid), len(bubbles))
    return (not reasons, desc + ("（" + "、".join(reasons) + "）" if reasons else ""))


def check_07_liquid(m: SlideModel) -> tuple[bool, str]:
    notes: list[str] = []
    flats = [s for s in m.shapes if s.w_cm >= 3.0 and s.h_cm <= 1.5
             and _near_any(s.fill_colors, "B6E2FD", RGB_TOL_LOOSE)]
    for s in m.shapes:
        if _labeled(s) or s.w_cm < 3.0:
            continue
        if not _near_any(s.fill_colors, "B6E2FD", RGB_TOL_LOOSE):
            continue
        x0, x1 = m.px(s.x0), m.px(s.x1)
        y0, y1 = m.py(s.y0), m.py(s.y1)
        desc = "%s：x%.2f%%–%.2f%% y%.2f%%–%.2f%% 填充%s" % (
            s.name, x0, x1, y0, y1, "/".join(s.fill_colors) or "无")
        if not (_span(x0, x1, 11, 30) and _span(y0, y1, 47, 72)):
            notes.append(desc + "（位置越界）")
            continue
        # 顶部水平液面：柱体(can)自带水平液面，或顶端另有一条扁平的浅蓝液面块
        top_flat = s.prst in ("can", "rect", "roundRect")
        if not top_flat:
            top_flat = any(abs(f.y0 - s.y0) <= 1.0 * EMU_PER_CM and f is not s for f in flats)
        if top_flat:
            return (True, desc)
        notes.append(desc + "（顶部未见水平液面形态）")
    return (False, "; ".join(notes[:3]) or "未找到浅蓝色液体色块")


def check_08_bubbles(m: SlideModel) -> tuple[bool, str]:
    cands = _bubbles(m, (13, 28), (56, 69))
    ok_fill = [s for s in cands
               if _is_light(s.fill_main) or _near_any(s.fill_colors, "B6E2FD", RGB_TOL_LOOSE)
               or s.fill_kind in ("inherit", "grpFill")]
    stacked = 0
    for i, a in enumerate(ok_fill):
        for b in ok_fill[i + 1:]:
            small = min(a.w * a.h, b.w * b.h)
            if small > 0 and _inter_area(a, b) / small >= 0.5:
                stacked += 1
    reasons: list[str] = []
    if len(ok_fill) < 7:
        reasons.append("符合条件的气泡仅%d个（需≥7）" % len(ok_fill))
    if stacked:
        reasons.append("存在%d对气泡重叠堆叠" % stacked)
    return (not reasons, "区域内圆形气泡%d个（尺寸/填充合规%d个）%s" % (
        len(cands), len(ok_fill), "（" + "、".join(reasons) + "）" if reasons else ""))


def _rod_parts(m: SlideModel, xr: tuple[float, float], yr: tuple[float, float],
               w_max: float) -> list[ShapeInfo]:
    """电极杆部件：有填充、非线条、非椭圆气泡，且完整落在给定窗口内。"""
    out: list[ShapeInfo] = []
    for s in m.shapes:
        if _labeled(s) or s.is_line_like or s.prst == "ellipse":
            continue
        if not s.filled or s.w_cm > w_max or s.h_cm > 4.6:
            continue
        if _window(m, s, xr, yr):
            out.append(s)
    return out


def _rod_check(m: SlideModel, xr: tuple[float, float], yr: tuple[float, float],
               box_x: tuple[float, float], box_y: tuple[float, float],
               w_rng: tuple[float, float], h_rng: tuple[float, float],
               cap_hexes: tuple[str, ...], cap_test: Callable[[Optional[str]], bool],
               cap_name: str) -> tuple[bool, str]:
    parts = _rod_parts(m, xr, yr, 1.6)
    if len(parts) < 3:
        return (False, "该区域仅找到%d个电极部件（需≥3段：帽/杆/尖端）" % len(parts))
    x0, y0, x1, y1 = _bbox(parts)
    w_cm, h_cm = (x1 - x0) / EMU_PER_CM, (y1 - y0) / EMU_PER_CM
    span_y = m.py(y1) - m.py(y0)
    cap = [s for s in parts
           if m.py(s.y1) <= m.py(y0) + 0.40 * span_y + PCT_TOL
           and any(_near_any(s.fill_colors, h, RGB_TOL_LOOSE) for h in cap_hexes)
           and any(cap_test(c) for c in s.fill_colors)]
    mid_lo, mid_hi = m.py(y0) + 0.30 * span_y, m.py(y0) + 0.80 * span_y
    shaft = [s for s in parts
             if m.py(s.y1) >= mid_lo and m.py(s.y0) <= mid_hi
             and any(_is_whiteish(c) for c in s.fill_colors)]
    tip = [s for s in parts
           if m.py(s.y1) >= m.py(y0) + 0.70 * span_y
           and (s.prst in ("cust", "triangle") or s.seg_bez > 0 or s.seg_line >= 3)]
    reasons: list[str] = []
    if not _span(m.px(x0), m.px(x1), box_x[0], box_x[1]) or not _span(m.py(y0), m.py(y1), box_y[0], box_y[1]):
        reasons.append("整体位置越界")
    if not (w_rng[0] <= w_cm <= w_rng[1]):
        reasons.append("宽度%.2fcm不在%.1f–%.1f" % (w_cm, w_rng[0], w_rng[1]))
    if not (h_rng[0] <= h_cm <= h_rng[1]):
        reasons.append("高度%.2fcm不在%.1f–%.1f" % (h_cm, h_rng[0], h_rng[1]))
    if not cap:
        reasons.append("顶部无%s帽（近似#%s）" % (cap_name, "或#".join(cap_hexes)))
    if not shaft:
        reasons.append("中段杆身非白色玻璃管质感")
    if not tip:
        reasons.append("底端无渐细尖端造型")
    desc = "电极部件%d段：x%.2f%%–%.2f%% y%.2f%%–%.2f%% %.2fcm×%.2fcm，帽%d/白杆%d/尖端%d" % (
        len(parts), m.px(x0), m.px(x1), m.py(y0), m.py(y1), w_cm, h_cm, len(cap), len(shaft), len(tip))
    return (not reasons, desc + ("（" + "、".join(reasons) + "）" if reasons else ""))


def check_09_green_electrode(m: SlideModel) -> tuple[bool, str]:
    return _rod_check(m, (10, 20), (37, 63), (11, 19), (38, 62),
                      (0.3, 1.5), (3.0, 4.5), ("5EAD5F",), _is_green, "绿色")


def check_10_working_column(m: SlideModel) -> tuple[bool, str]:
    parts = _rod_parts(m, (17, 25), (23, 66), 2.0)
    if len(parts) < 2:
        return (False, "该区域仅找到%d个柱状部件" % len(parts))
    x0, y0, x1, y1 = _bbox(parts)
    w_cm, h_cm = (x1 - x0) / EMU_PER_CM, (y1 - y0) / EMU_PER_CM
    shaft = [s for s in parts if s.h_cm >= 1.5 and any(_is_whiteish(c) for c in s.fill_colors)]
    cap = [s for s in parts
           if _near_any(s.fill_colors, "2E3F58", RGB_TOL_LOOSE)
           and _span(m.py(s.y0), m.py(s.y1), 23, 32)]
    reasons: list[str] = []
    if not _span(m.px(x0), m.px(x1), 17, 25) or not _span(m.py(y0), m.py(y1), 23, 66):
        reasons.append("整体位置越界")
    if not shaft:
        reasons.append("柱身非白色半透明质感")
    if not cap:
        reasons.append("顶端无位于23%–32%高度的深蓝灰接头（近似#2E3F58）")
    desc = "柱状部件%d段：x%.2f%%–%.2f%% y%.2f%%–%.2f%% %.2fcm×%.2fcm，白柱%d/深蓝灰帽%d" % (
        len(parts), m.px(x0), m.px(x1), m.py(y0), m.py(y1), w_cm, h_cm, len(shaft), len(cap))
    return (not reasons, desc + ("（" + "、".join(reasons) + "）" if reasons else ""))


def check_11_red_electrode(m: SlideModel) -> tuple[bool, str]:
    return _rod_check(m, (21, 31), (37, 63), (22, 30), (38, 62),
                      (0.3, 1.5), (3.0, 4.5), ("B90303", "EA0000"), _is_red, "红色")


def check_12_ref_label(m: SlideModel) -> tuple[bool, str]:
    return _text_hit(m, "参比电极", (2, 15), (47, 57), (13, 17))


def check_13_work_label(m: SlideModel) -> tuple[bool, str]:
    return _text_hit(m, "工作电极", (16, 30), (44, 53), (13, 17))


def check_14_counter_label(m: SlideModel) -> tuple[bool, str]:
    return _text_hit(m, "对电极", (28, 39), (47, 57), (13, 17))


def check_15_sample_label(m: SlideModel) -> tuple[bool, str]:
    return _text_hit(m, "样品溶液", (12, 28), (62, 71), (13, 17))


def check_16_label_arrows(m: SlideModel) -> tuple[bool, str]:
    def pick(xr: tuple[float, float]) -> list[ShapeInfo]:
        out: list[ShapeInfo] = []
        for s in m.shapes:
            if _labeled(s) or s.w <= s.h:
                continue
            if not (_span(m.px(s.x0), m.px(s.x1), xr[0], xr[1])
                    and _span(m.py(s.y0), m.py(s.y1), 49, 56)):
                continue
            out.append(s)
        return out

    def judge(cands: list[ShapeInfo], want_right: bool, side: str) -> tuple[bool, str]:
        if not cands:
            return (False, "%s未找到水平线段" % side)
        notes: list[str] = []
        for s in cands:
            # 端点箭头随 flipH 一同镜像，故按规范化后的渲染结果判断箭头朝向。
            tail, head = ShapeInfo._end(s.tail_end), ShapeInfo._end(s.head_end)
            at_right = (tail and not s.flip_h) or (head and s.flip_h)
            at_left = (head and not s.flip_h) or (tail and s.flip_h)
            reasons: list[str] = []
            if not s.has_arrow_end:
                reasons.append("无端点箭头")
            elif not (at_right if want_right else at_left):
                reasons.append("箭头头部不在%s端" % ("右" if want_right else "左"))
            if not _is_black(s.line_main):
                reasons.append("线条非黑色(%s)" % (s.line_main or "无"))
            if not _solid_line(s):
                reasons.append("线型非实线(%s)" % s.dash)
            if not (0.5 <= s.w_cm <= 1.5):
                reasons.append("长度%.2fcm不在0.5–1.5" % s.w_cm)
            desc = "%s：x%.2f%%–%.2f%% y%.2f%% 长%.2fcm 箭头%s/%s flipH=%s" % (
                s.name, m.px(s.x0), m.px(s.x1), m.py(s.y0), s.w_cm,
                s.head_end or "-", s.tail_end or "-", s.flip_h)
            if not reasons:
                return (True, desc)
            notes.append(desc + "（" + "、".join(reasons) + "）")
        return (False, "%s%s" % (side, "; ".join(notes[:2])))

    ok_l, d_l = judge(pick((10, 16)), True, "左侧箭头")
    ok_r, d_r = judge(pick((26, 32)), False, "右侧箭头")
    return (ok_l and ok_r, "%s | %s" % (d_l, d_r))


def check_17_chassis(m: SlideModel) -> tuple[bool, str]:
    parts = [s for s in m.shapes
             if not _labeled(s) and not s.is_line_like and s.filled and _window(m, s, (37, 66), (44, 66))]
    if not parts:
        return (False, "该区域未找到机箱色块")
    top_col = [s for s in parts if _near_any(s.fill_colors, "C7D0E0", RGB_TOL_LOOSE) and s.w_cm >= 4.0]
    top = [s for s in top_col if s.h_cm <= 1.2]
    front = [s for s in parts if _near_any(s.fill_colors, "233959", RGB_TOL_LOOSE) and s.w_cm >= 3.0]
    reasons: list[str] = []
    if not top_col:
        reasons.append("无浅灰色顶面色块（近似#C7D0E0）")
    elif not top:
        reasons.append("顶面高度%.2fcm超过1.2cm，非扁平斜置四边形" % min(s.h_cm for s in top_col))
    if not front:
        reasons.append("无深蓝色前壳色块（近似#233959）")
    if reasons:
        return (False, "机箱色块%d块（%s）" % (len(parts), "、".join(reasons)))
    fr = max(front, key=lambda s: s.w)
    panel = [s for s in parts if _near_any(s.fill_colors, "1777D5", RGB_TOL_LOOSE) and _inside(s, fr)]
    side = [s for s in parts
            if all(s is not t for t in top) and s is not fr and s.h_cm >= 2.5 and s.fill_colors
            and s.x1 >= fr.x1 + 1.0 * EMU_PER_CM
            and min(_lum(c) for c in s.fill_colors) < min(_lum(c) for c in top[0].fill_colors)]
    top_over = (max(s.x1 for s in top) - fr.x1) / EMU_PER_CM
    if not panel:
        reasons.append("前壳内未叠放亮蓝色面板（近似#1777D5）")
    if not side:
        reasons.append("前壳右侧无更深的灰蓝色侧面色块（高度≥2.5cm）")
    if top_over < 1.0:
        reasons.append("顶面右端仅比前壳右边线外伸%.2fcm（需≥1cm）" % top_over)
    if side:
        side_over = (max(s.x1 for s in side) - fr.x1) / EMU_PER_CM
        if side_over < 1.0:
            reasons.append("侧面右端仅比前壳右边线外伸%.2fcm（需≥1cm）" % side_over)
    desc = "顶面%s(%.2fcm高)/前壳%s(右边线%.2f%%)/面板%d块/侧面%d块，顶面外伸%.2fcm" % (
        top[0].name, top[0].h_cm, fr.name, m.px(fr.x1), len(panel), len(side), top_over)
    return (not reasons, desc + ("（" + "、".join(reasons) + "）" if reasons else ""))


def check_18_screen(m: SlideModel) -> tuple[bool, str]:
    notes: list[str] = []
    for s in m.shapes:
        if _labeled(s) or s.is_line_like:
            continue
        if not _window(m, s, (39, 48), (51, 62)) or s.w_cm < 0.8 or s.h_cm < 0.5:
            continue
        white = _is_whiteish(s.fill_main) or _is_whiteish(s.line_main)
        inner = [t for t in m.shapes if t is not s and _inside(t, s, 0.10 * EMU_PER_CM)]
        wave_one = [t for t in inner if t.seg_bez > 0 or t.seg_line >= 2]
        wave_seg = [t for t in inner if t.is_line_like or (t.w > 0 and t.h_cm <= 0.20)]
        desc = "%s：x%.2f%%–%.2f%% y%.2f%%–%.2f%% 填充%s 边框%s，内部波形%d段" % (
            s.name, m.px(s.x0), m.px(s.x1), m.py(s.y0), m.py(s.y1),
            s.fill_main or s.fill_kind, s.line_main or "无", len(wave_one) + len(wave_seg))
        reasons: list[str] = []
        if not white:
            reasons.append("既非白色填充也非白色边框")
        if not wave_one and len(wave_seg) < 3:
            reasons.append("内部无波形折线")
        if not reasons:
            return (True, desc)
        notes.append(desc + "（" + "、".join(reasons) + "）")
    return (False, "; ".join(notes[:3]) or "该区域未找到小屏幕矩形")


def check_19_leds(m: SlideModel) -> tuple[bool, str]:
    cands = [s for s in m.shapes
             if s.prst == "ellipse" and not _labeled(s)
             and 0.15 <= max(s.w_cm, s.h_cm) <= 0.45 and _window(m, s, (53, 59), (49, 62))]
    cands.sort(key=lambda s: s.y0)
    if len(cands) != 3:
        return (False, "该区域圆形指示灯为%d个（需3个，直径0.15–0.45cm）" % len(cands))
    tests: list[tuple[Callable[[Optional[str]], bool], str]] = [
        (_is_green, "绿色"), (_is_dark_slate, "深灰蓝色"), (_is_red, "红色")]
    reasons: list[str] = []
    for s, (fn, nm) in zip(cands, tests):
        if not any(fn(c) for c in s.fill_colors):
            reasons.append("%s填充%s非%s" % (s.name, s.fill_main or "无", nm))
    desc = "指示灯自上而下：%s" % "、".join(
        "%s(%.2fcm,%s)" % (s.name, max(s.w_cm, s.h_cm), s.fill_main or "无") for s in cands)
    return (not reasons, desc + ("（" + "、".join(reasons) + "）" if reasons else ""))


def check_20_analyzer_text(m: SlideModel) -> tuple[bool, str]:
    ok, detail = _text_hit(m, "电化学分析仪", (44, 58), (52, 60), (11, 13))
    if not ok:
        return (False, detail)
    notes: list[str] = []
    for s in _find_text(m, "电化学分析仪"):
        white = any(_is_whiteish(r.color) for r in s.runs)
        bold = any(r.bold for r in s.runs)
        if white and not bold:
            return (True, detail + "，白色非加粗")
        bad: list[str] = []
        if not white:
            bad.append("文字颜色%s非白色" % "/".join(r.color or "-" for r in s.runs))
        if bold:
            bad.append("文字加粗")
        notes.append("%s（%s）" % (s.name, "、".join(bad)))
    return (False, detail + "，但 " + "; ".join(notes))


def _panel_check(m: SlideModel, needle: str, xr: tuple[float, float], yr: tuple[float, float],
                 sr: tuple[float, float]) -> tuple[bool, str]:
    """无填充蓝色虚线圆角矩形 + 框内指定文本 + 文本字号。"""
    boxes = [s for s in m.shapes
             if s.prst in ("roundRect", "round2SameRect", "round1Rect", "rect", "cust")
             and s.w_cm >= 1.5 and s.h_cm >= 0.5 and _window(m, s, xr, yr)]
    txts = _find_text(m, needle)
    if not txts:
        return (False, "未找到“%s”文本" % needle)
    if not boxes:
        return (False, "该区域未找到框形（“%s”文本位于 x%.2f%%–%.2f%% y%.2f%%–%.2f%%）"
                % (needle, m.px(txts[0].x0), m.px(txts[0].x1), m.py(txts[0].y0), m.py(txts[0].y1)))
    notes: list[str] = []
    for b in boxes:
        reasons: list[str] = []
        if b.fill_kind != "none":
            reasons.append("非无填充(%s)" % b.fill_kind)
        if not _dashed(b):
            reasons.append("边框非虚线(%s)" % b.dash)
        if not _is_blue(b.line_main):
            reasons.append("边框非蓝色(%s)" % (b.line_main or "无"))
        # 文本框自身高度常大于可见文字，故按“文本主体落在框内”判定（面积过半或中心在框内）
        inner: list[ShapeInfo] = []
        for t in txts:
            if t is b:
                inner.append(t)
                continue
            area = t.w * t.h
            cx, cy = (t.x0 + t.x1) / 2.0, (t.y0 + t.y1) / 2.0
            center_in = (b.x0 <= cx <= b.x1) and (b.y0 <= cy <= b.y1)
            if (area > 0 and _inter_area(t, b) / area >= 0.5) or center_in:
                inner.append(t)
        if not inner:
            reasons.append("框内无“%s”文本" % needle)
            szs: list[float] = []
        else:
            szs = _sizes(inner[0])
            if not any(sr[0] <= v <= sr[1] for v in szs):
                reasons.append("文本字号%s磅不在%g–%g" % ("/".join("%.1f" % v for v in szs), sr[0], sr[1]))
        desc = "%s：x%.2f%%–%.2f%% y%.2f%%–%.2f%% 填充%s 边框%s %s 文本字号%s磅" % (
            b.name, m.px(b.x0), m.px(b.x1), m.py(b.y0), m.py(b.y1), b.fill_kind,
            b.line_main or "无", b.dash, "/".join("%.1f" % v for v in szs) or "-")
        if not reasons:
            return (True, desc)
        notes.append(desc + "（" + "、".join(reasons) + "）")
    return (False, "; ".join(notes[:3]))


def check_21_acq_box(m: SlideModel) -> tuple[bool, str]:
    return _panel_check(m, "信号采集与处理", (46, 65), (30, 40), (13, 16))


def check_22_ctrl_box(m: SlideModel) -> tuple[bool, str]:
    return _panel_check(m, "电位/电流控制", (41, 58), (68, 79), (14, 18))


def check_23_viz_box(m: SlideModel) -> tuple[bool, str]:
    return _panel_check(m, "数据分析与可视化", (73, 93), (18, 31), (13, 19))


def check_24_flow_arrows(m: SlideModel) -> tuple[bool, str]:
    wins = [((47.0, 53.0), (63.0, 71.0)), ((53.0, 59.0), (38.0, 48.0)), ((80.0, 86.0), (26.0, 35.0))]
    hits: list[str] = []
    bad: list[str] = []
    for xr, yr in wins:
        found = ""
        notes: list[str] = []
        for s in m.shapes:
            if _labeled(s) or s.h <= s.w or not _window(m, s, xr, yr):
                continue
            desc = "%s(x%.2f%%–%.2f%% y%.2f%%–%.2f%% %s 箭头%s/%s %s)" % (
                s.name, m.px(s.x0), m.px(s.x1), m.py(s.y0), m.py(s.y1), s.dash,
                s.head_end or "-", s.tail_end or "-", s.line_main or "无")
            reasons: list[str] = []
            if not _dashed(s):
                reasons.append("非虚线")
            if not s.has_arrow_end:
                reasons.append("末端无箭头")
            if not _is_blue(s.line_main):
                reasons.append("非蓝色")
            if not reasons:
                found = desc
                break
            notes.append(desc + "（" + "、".join(reasons) + "）")
        if found:
            hits.append(found)
        else:
            bad.append("x%g%%–%g%%/y%g%%–%g%% 处：%s"
                       % (xr[0], xr[1], yr[0], yr[1], "; ".join(notes[:2]) or "未找到竖直线条"))
    return (len(hits) == 3, "命中%d/3条虚线箭头%s" % (
        len(hits), "（" + "；".join(bad) + "）" if bad else "：" + "、".join(hits)))


def check_25_curves(m: SlideModel) -> tuple[bool, str]:
    specs = [("5EAD5F", (9.0, 43.0), (32.0, 51.0), "绿色"),
             ("B90303", (26.0, 46.0), (31.0, 49.0), "红色"),
             ("151310", (20.0, 48.0), (21.0, 48.0), "黑色")]
    hits: list[str] = []
    bad: list[str] = []
    for hexc, xr, yr, nm in specs:
        found = ""
        notes: list[str] = []
        for s in m.shapes:
            if _labeled(s) or not _near_any(s.line_colors, hexc, RGB_TOL):
                continue
            if s.w_cm < 1.0 or s.h_cm < 0.3:
                continue
            desc = "%s(x%.2f%%–%.2f%% y%.2f%%–%.2f%% %s %.2f磅 曲线段%d)" % (
                s.name, m.px(s.x0), m.px(s.x1), m.py(s.y0), m.py(s.y1),
                s.line_main or "无", s.line_w_pt, s.seg_bez)
            reasons: list[str] = []
            if not _window(m, s, xr, yr):
                reasons.append("位置越界")
            if s.seg_bez <= 0:
                reasons.append("非平滑弯曲曲线")
            if not (1.5 <= s.line_w_pt <= 3.0):
                reasons.append("线宽%.2f磅不在1.5–3" % s.line_w_pt)
            if not reasons:
                found = desc
                break
            notes.append(desc + "（" + "、".join(reasons) + "）")
        if found:
            hits.append(nm + found)
        else:
            bad.append("%s曲线（近似#%s）：%s" % (nm, hexc, "; ".join(notes[:2]) or "未找到"))
    return (len(hits) == 3, "命中%d/3条曲线%s" % (
        len(hits), "（" + "；".join(bad) + "）" if bad else "：" + "、".join(hits)))


def check_26_host_link(m: SlideModel) -> tuple[bool, str]:
    cands = [s for s in m.shapes
             if not _labeled(s) and s.w >= s.h and _is_black(s.line_main)
             and _window(m, s, (62, 74), (50, 58))]
    lines = [s for s in cands if s.w_cm >= 1.0]
    blocks = [s for s in m.shapes
              if not _labeled(s) and s.filled and 0.4 <= s.w_cm <= 1.0
              and _window(m, s, (62, 74), (50, 58))]
    if not lines:
        if len(cands) >= 3:
            return (False, "该处黑色水平线由%d段碎片拼接（最长%.2fcm），未由一条线段实现"
                    % (len(cands), max(s.w_cm for s in cands)))
        return (False, "该区域未找到长度≥1cm的黑色水平线段")
    ln = max(lines, key=lambda s: s.w)
    right = [b for b in blocks if b.x0 >= ln.x1 - 0.30 * EMU_PER_CM]
    reasons: list[str] = []
    if len(cands) >= 3:
        reasons.append("该处另有%d段同色碎片" % (len(cands) - 1))
    if not right:
        reasons.append("线段右端无宽度0.4–1cm的小矩形块（该区域候选块%d个）" % len(blocks))
    desc = "%s：x%.2f%%–%.2f%% y%.2f%% 长%.2fcm %s%s" % (
        ln.name, m.px(ln.x0), m.px(ln.x1), m.py(ln.y0), ln.w_cm, ln.line_main or "无",
        "，右端矩形块%s宽%.2fcm" % (right[0].name, right[0].w_cm) if right else "")
    return (not reasons, desc + ("（" + "、".join(reasons) + "）" if reasons else ""))


def _cluster(m: SlideModel, xr: tuple[float, float], yr: tuple[float, float],
             frac: float = 0.5) -> list[ShapeInfo]:
    """取“主体落在给定区域内”的有填充部件：超出区域的部件也会被纳入，便于发现越界。"""
    ax0, ax1 = xr[0] / 100.0 * m.page_w, xr[1] / 100.0 * m.page_w
    ay0, ay1 = yr[0] / 100.0 * m.page_h, yr[1] / 100.0 * m.page_h
    out: list[ShapeInfo] = []
    for s in m.shapes:
        if _labeled(s) or not s.filled:
            continue
        ox = min(s.x1, ax1) - max(s.x0, ax0)
        oy = min(s.y1, ay1) - max(s.y0, ay0)
        if ox < 0 or oy < 0:
            continue
        fx = ox / s.w if s.w > 0 else (1.0 if ax0 <= s.x0 <= ax1 else 0.0)
        fy = oy / s.h if s.h > 0 else (1.0 if ay0 <= s.y0 <= ay1 else 0.0)
        if fx >= frac and fy >= frac:
            out.append(s)
    return out


def check_27_laptop(m: SlideModel) -> tuple[bool, str]:
    parts = _cluster(m, (65, 95), (32, 68))
    if len(parts) < 4:
        return (False, "该区域仅找到%d个部件，未构成笔记本电脑图形" % len(parts))
    x0, y0, x1, y1 = _bbox(parts)
    frame = [s for s in parts
             if _near_any(s.fill_colors, "2B415C", RGB_TOL_LOOSE) and s.w_cm >= 3.0 and s.h_cm >= 1.5]
    fr = max(frame, key=lambda s: s.w * s.h) if frame else None
    below = (fr.y1 - 0.5 * EMU_PER_CM) if fr is not None else y0
    screen = [s for s in parts
              if any(_is_whiteish(c) for c in s.fill_colors) and s.w_cm >= 2.0 and s.h_cm >= 1.0]
    base = [s for s in parts if s.w_cm >= 4.0 and s.y0 >= (below if fr is None else (fr.y0 + fr.y1) / 2)]
    pad = [s for s in parts if 1.2 <= s.w_cm <= 4.5 and s.h_cm <= 1.5 and s.y0 >= below]
    keys = [s for s in parts if s.y0 >= below and s.w_cm <= 2.5 and s.h_cm <= 0.9
            and all(s is not p for p in pad)]
    key_lines = [s for s in m.shapes
                 if not _labeled(s) and s.line_colors and s.y0 >= below
                 and s.w_cm <= 2.5 and s.h_cm <= 0.9 and _window(m, s, (65, 95), (32, 68))]
    reasons: list[str] = []
    if not _span(m.px(x0), m.px(x1), 65, 95) or not _span(m.py(y0), m.py(y1), 32, 68):
        reasons.append("整体位置越界")
    if not frame:
        reasons.append("无深蓝色屏幕外框（近似#2B415C）")
    if not screen:
        reasons.append("无白色屏幕区")
    if not base:
        reasons.append("屏幕下方无底座色块")
    if not pad:
        reasons.append("底座上无触控板色块")
    if len(keys) + len(key_lines) < 5:
        reasons.append("底座上键盘按键线条/形状仅%d个（需≥5）" % (len(keys) + len(key_lines)))
    desc = "笔记本部件%d个：x%.2f%%–%.2f%% y%.2f%%–%.2f%%，外框%d/屏幕%d/底座%d/触控板%d/按键%d" % (
        len(parts), m.px(x0), m.px(x1), m.py(y0), m.py(y1),
        len(frame), len(screen), len(base), len(pad), len(keys) + len(key_lines))
    return (not reasons, desc + ("（" + "、".join(reasons) + "）" if reasons else ""))


def _chart_frame(m: SlideModel, cur: ShapeInfo) -> tuple[bool, str]:
    pad = 1.5 * EMU_PER_CM
    rects = [s for s in m.shapes
             if s is not cur and s.prst in ("rect", "roundRect")
             and (_is_light(s.line_main) or _is_light(s.fill_main))
             and _inside(cur, s, 0.5 * EMU_PER_CM)]
    if rects:
        return (True, "浅色矩形边框%s" % rects[0].name)
    near = [s for s in m.shapes
            if s is not cur and _is_light(s.line_main)
            and s.x0 <= cur.x1 + pad and s.x1 >= cur.x0 - pad
            and s.y0 <= cur.y1 + pad and s.y1 >= cur.y0 - pad]
    vert = [s for s in near if s.w_cm <= 0.15 and s.h_cm >= 0.5]
    horiz = [s for s in near if s.h_cm <= 0.15 and s.w_cm >= 0.5]
    if vert and horiz:
        return (True, "浅色边框线%d竖%d横" % (len(vert), len(horiz)))
    return (False, "外围无浅色矩形图表边框（浅色边线%d竖%d横）" % (len(vert), len(horiz)))


def _chart_curve(m: SlideModel, hexes: tuple[str, ...], xr: tuple[float, float],
                 yr: tuple[float, float], name: str) -> tuple[bool, str]:
    same = [s for s in m.shapes
            if not _labeled(s)
            and any(_near_any(s.line_colors, h, RGB_TOL_LOOSE) for h in hexes)
            and _window(m, s, xr, yr)]
    cands = [s for s in same if s.w_cm >= 0.5 and s.h_cm >= 0.3]
    if not cands:
        if len(same) >= 3:
            return (False, "该处%s线由%d段各自独立的线段首尾拼接而成，未由一个连续形状实现"
                    % (name, len(same)))
        return (False, "该区域未找到%s曲线（近似#%s）" % (name, "或#".join(hexes)))
    if len(cands) > 1:
        return (False, "该处%s线由%d段独立形状拼接而成，未由一个连续形状实现" % (name, len(cands)))
    c = cands[0]
    reasons: list[str] = []
    if len(same) > 1:
        reasons.append("同色另有%d段独立线段" % (len(same) - 1))
    if c.seg_move > 1:
        reasons.append("路径含%d个起点，线条存在断开缺口" % c.seg_move)
    if c.seg_bez <= 0 and c.seg_line < 2:
        reasons.append("非连续曲线或折线形状")
    ok_frame, fdesc = _chart_frame(m, c)
    if not ok_frame:
        reasons.append(fdesc)
    desc = "%s：x%.2f%%–%.2f%% y%.2f%%–%.2f%% 线%s 路径(直线%d/曲线%d/起点%d)%s" % (
        c.name, m.px(c.x0), m.px(c.x1), m.py(c.y0), m.py(c.y1), c.line_main or "无",
        c.seg_line, c.seg_bez, c.seg_move, "，" + fdesc if ok_frame else "")
    return (not reasons, desc + ("（" + "、".join(reasons) + "）" if reasons else ""))


def check_28_blue_chart(m: SlideModel) -> tuple[bool, str]:
    return _chart_curve(m, ("3276C3",), (72, 83), (39, 54), "蓝色")


def check_29_red_chart(m: SlideModel) -> tuple[bool, str]:
    return _chart_curve(m, ("E56F6E", "E10D0D"), (81, 92), (40, 55), "红色")


def check_30_computer_label(m: SlideModel) -> tuple[bool, str]:
    return _text_hit(m, "计算机", (73, 85), (64, 72), (13, 17))


def check_31_legend_box(m: SlideModel) -> tuple[bool, str]:
    notes: list[str] = []
    for s in m.shapes:
        if s.prst not in ("roundRect", "round1Rect", "round2SameRect"):
            continue
        if not _window(m, s, (9, 91), (80, 96)) or s.w_cm < 10.0:
            continue
        reasons: list[str] = []
        if s.fill_kind != "none":
            reasons.append("非无填充(%s)" % s.fill_kind)
        if not _dashed(s):
            reasons.append("边框非虚线(%s)" % s.dash)
        if not _is_blue(s.line_main):
            reasons.append("边框非蓝色(%s)" % (s.line_main or "无"))
        if not (25.0 <= s.w_cm <= 30.0):
            reasons.append("宽度%.2fcm不在25–30" % s.w_cm)
        if not (2.0 <= s.h_cm <= 3.0):
            reasons.append("高度%.2fcm不在2–3" % s.h_cm)
        desc = "%s：x%.2f%%–%.2f%% y%.2f%%–%.2f%% %.2fcm×%.2fcm 填充%s 边框%s %s" % (
            s.name, m.px(s.x0), m.px(s.x1), m.py(s.y0), m.py(s.y1), s.w_cm, s.h_cm,
            s.fill_kind, s.line_main or "无", s.dash)
        if not reasons:
            return (True, desc)
        notes.append(desc + "（" + "、".join(reasons) + "）")
    return (False, "; ".join(notes[:3]) or "该区域未找到宽度≥10cm的圆角矩形")


def _legend_row(m: SlideModel, xr: tuple[float, float], yr: tuple[float, float],
                ctest: Callable[[Optional[str]], bool], cname: str, needle: str,
                sr: tuple[float, float], w_rng: tuple[float, float],
                want_dash: bool = False) -> tuple[bool, str]:
    lines = [s for s in m.shapes
             if not _labeled(s) and s.w > s.h and s.line_colors and _window(m, s, xr, yr)]
    txts = _find_text(m, needle)
    if not txts:
        return (False, "未找到“%s”文本" % needle)
    if not lines:
        return (False, "该区域未找到水平线段（“%s”文本已出现）" % needle)
    notes: list[str] = []
    for ln in lines:
        reasons: list[str] = []
        if not ctest(ln.line_main):
            reasons.append("线段非%s(%s)" % (cname, ln.line_main or "无"))
        if not (w_rng[0] <= ln.line_w_pt <= w_rng[1]):
            reasons.append("线宽%.2f磅不在%g–%g" % (ln.line_w_pt, w_rng[0], w_rng[1]))
        if want_dash and not _dashed(ln):
            reasons.append("线型非虚线(%s)" % ln.dash)
        if want_dash and not ln.has_arrow_end:
            reasons.append("末端无箭头")
        cy = (m.py(ln.y0) + m.py(ln.y1)) / 2.0
        right = [t for t in txts
                 if t.x0 >= ln.x1 - 0.5 * EMU_PER_CM
                 and abs((m.py(t.y0) + m.py(t.y1)) / 2.0 - cy) <= 4.0]
        szs = _sizes(right[0]) if right else []
        if not right:
            reasons.append("线段右侧未见“%s”文本" % needle)
        elif not any(sr[0] <= v <= sr[1] for v in szs):
            reasons.append("文本字号%s磅不在%g–%g" % ("/".join("%.1f" % v for v in szs), sr[0], sr[1]))
        desc = "%s：x%.2f%%–%.2f%% y%.2f%% %s %.2f磅 %s，右侧文本字号%s磅" % (
            ln.name, m.px(ln.x0), m.px(ln.x1), cy, ln.line_main or "无", ln.line_w_pt,
            ln.dash, "/".join("%.1f" % v for v in szs) or "-")
        if not reasons:
            return (True, desc)
        notes.append(desc + "（" + "、".join(reasons) + "）")
    return (False, "; ".join(notes[:2]))


def check_32_legend_work(m: SlideModel) -> tuple[bool, str]:
    return _legend_row(m, (12, 18), (84, 88), _is_black, "黑色",
                       "工作电极连接线（信号输出）", (12, 15), (1.5, 2.5))


def check_33_legend_ref(m: SlideModel) -> tuple[bool, str]:
    return _legend_row(m, (12, 18), (89, 93), _is_green, "绿色",
                       "参比电极连接线（电位参照）", (12, 15), (1.5, 2.5))


def check_34_legend_counter(m: SlideModel) -> tuple[bool, str]:
    return _legend_row(m, (41, 47), (85, 89), _is_red, "红色",
                       "对电极连接线（电流通路）", (12, 15), (1.5, 2.5))


def check_35_legend_flow(m: SlideModel) -> tuple[bool, str]:
    return _legend_row(m, (67, 74), (85, 89), _is_blue, "蓝色",
                       "信号/控制/数据流向", (12, 15), (2.0, 3.0), want_dash=True)


_EXPECTED_TEXTS: list[str] = [
    "分析物输入", "参比电极", "工作电极", "对电极", "样品溶液", "电化学分析仪",
    "信号采集与处理", "电位/电流控制", "数据分析与可视化", "计算机",
    "工作电极连接线（信号输出）", "参比电极连接线（电位参照）",
    "对电极连接线（电流通路）", "信号/控制/数据流向",
]


def _expected_like(txt: str) -> bool:
    """是否属于要求之内的文字。错字/多空格等按“写错”处理，不在本条重复计分。"""
    t = _norm(txt)
    if not t:
        return True
    for e in _EXPECTED_TEXTS:
        k = _norm(e)
        if k in t or t in k:
            return True
        if len(set(t) & set(k)) >= max(2, int(0.5 * len(k))):
            return True
    return False


def check_36_text_integrity(m: SlideModel) -> tuple[bool, str]:
    problems: list[str] = []
    for t in m.texts():
        label = t.text.strip().replace("\n", " ")
        area = t.w * t.h
        cov, worst = 0.0, None
        for s in m.shapes:
            # 注意：自动图形即使没有文字也带文本框，故此处按“是否叠放在文字之上且不透明”判断
            if s is t or s.z <= t.z or not s.opaque_fill:
                continue
            ia = _inter_area(t, s)
            if ia > cov:
                cov, worst = ia, s
        if area > 0 and cov / area >= 0.15 and worst is not None:
            problems.append("“%s”被%s遮挡约%.1f%%" % (label[:12], worst.name, cov / area * 100.0))
        tol = 0.2 * EMU_PER_CM
        if t.x0 < -tol or t.y0 < -tol or t.x1 > m.page_w + tol or t.y1 > m.page_h + tol:
            problems.append("“%s”超出页面边界被裁切" % label[:12])
        if not _expected_like(t.text):
            problems.append("出现要求之外的文字“%s”" % label[:16])
    if problems:
        return (False, "; ".join(problems[:5]))
    return (True, "共%d处文本，均完整显示且无多余文字" % len(m.texts()))


_CHECKS: list[Callable[[SlideModel], tuple[bool, str]]] = [
    check_01_aspect,            # +5 整体宽高比 1.63–1.73
    check_02_centered,          # +5 整体居中放置
    check_03_outer_frame,       # +3 蓝色实线外框
    check_04_inlet_text,        # +3 “分析物输入”文本
    check_05_inlet_arrow,       # +1 竖直蓝色虚线箭头
    check_06_beaker,            # +5 烧杯整体
    check_07_liquid,            # +3 浅蓝液体色块
    check_08_bubbles,           # +1 气泡≥7个
    check_09_green_electrode,   # +3 绿色电极杆
    check_10_working_column,    # +3 竖直柱体与深色帽盖
    check_11_red_electrode,     # +3 红色电极杆
    check_12_ref_label,         # +1 “参比电极”
    check_13_work_label,        # +1 “工作电极”
    check_14_counter_label,     # +1 “对电极”
    check_15_sample_label,      # +1 “样品溶液”
    check_16_label_arrows,      # +1 两条黑色实线箭头
    check_17_chassis,           # +5 斜视立体机箱
    check_18_screen,            # +3 小屏幕与波形
    check_19_leds,              # +1 三个指示灯
    check_20_analyzer_text,     # +1 “电化学分析仪”
    check_21_acq_box,           # +3 “信号采集与处理”虚线框
    check_22_ctrl_box,          # +3 “电位/电流控制”虚线框
    check_23_viz_box,           # +3 “数据分析与可视化”虚线框
    check_24_flow_arrows,       # +1 三条竖直虚线箭头
    check_25_curves,            # +5 三条平滑曲线
    check_26_host_link,         # +1 主机连线与小矩形块
    check_27_laptop,            # +5 笔记本电脑图形
    check_28_blue_chart,        # +3 蓝色图表曲线
    check_29_red_chart,         # +3 红色图表曲线
    check_30_computer_label,    # +1 “计算机”
    check_31_legend_box,        # +3 图例虚线框
    check_32_legend_work,       # +1 图例·工作电极连接线
    check_33_legend_ref,        # +1 图例·参比电极连接线
    check_34_legend_counter,    # +1 图例·对电极连接线
    check_35_legend_flow,       # +1 图例·信号/控制/数据流向
    check_36_text_integrity,    # +1 文字完整可见且无多余文字
]

assert len(_CHECKS) == len(_DIM2_RULES), (
    "检查函数个数与细则条数不一致：%d vs %d" % (len(_CHECKS), len(_DIM2_RULES))
)


# ================================================================ 维度1 判定 ====

def _dim1_gate(file_path: str, m: SlideModel) -> list[str]:
    fails: list[str] = []
    if not file_path.lower().endswith(".pptx"):
        fails.append("交付文件格式为%s，不是.pptx" % (os.path.splitext(file_path)[1] or "无扩展名"))
    if m.slide_count != 1:
        fails.append("幻灯片页数为%d页（要求1页）" % m.slide_count)
    page = m.page_w * m.page_h
    pics = [s for s in m.shapes if s.is_picture]
    editable = [s for s in m.shapes if not s.is_picture]
    big = [p for p in pics if page > 0 and (p.w * p.h) / page >= 0.60]
    if big and len(editable) < 10:
        fails.append("以整张位图充当版面（位图%s覆盖页面%.0f%%，可编辑对象仅%d个）"
                     % (big[0].name, (big[0].w * big[0].h) / page * 100.0, len(editable)))
    return fails


# ================================================================== 主流程 ====

def _dim1_fail(result: dict[str, Any], reason: str) -> dict[str, Any]:
    """维度一未通过：完成度总分记 0，不再逐条罗列（未真正逐条评过）。

    status 保持 "ok"（文件确实找到了，只是维度一不达标），dim1_reason 必须非空。"""
    result["dim1_pass"] = False
    result["dim1_reason"] = reason or "维度一未通过（未给出具体原因）"
    result["dim2_items"] = []
    result["total_score"] = 0
    result["max_score"] = MAX_SCORE_TOTAL
    return result


def _evaluate_impl(file_path: str, result: dict[str, Any]) -> dict[str, Any]:
    # 文件已找到但打不开（损坏 / 需修复 / 实际不是 pptx 包）属于维度一未通过，
    # 而不是"没评上"：status 仍为 ok，按维度一记 0 分并写错误摘要。
    try:
        model = build_model(file_path)
    except Exception as exc:
        return _dim1_fail(result, "交付文件无法正常打开（损坏或需修复）：%s: %s"
                                  % (type(exc).__name__, exc))

    gate = _dim1_gate(file_path, model)
    if gate:
        # 维度一未通过：完成度总分记 0，不再逐条罗列（未真正逐条评过）。
        return _dim1_fail(result, "; ".join(gate))

    result["dim1_pass"] = True
    # 约定：维度一通过时 dim1_reason 保持空串，只有未通过才写错误摘要。
    result["dim1_reason"] = ""

    results: list[str] = []
    for fn, (md, rule) in zip(_CHECKS, _DIM2_RULES):
        try:
            hit, _evidence = fn(model)
        except Exception:  # 单条检查异常不影响整体评测，按未命中计
            hit = False
        results.append("+%d：%s" % (md if hit else 0, rule))

    if len(results) != len(_MAX_DELTAS_ORDER):
        raise RuntimeError("dim2 项数不一致：results=%d vs order=%d"
                           % (len(results), len(_MAX_DELTAS_ORDER)))

    dim2_items: list[dict[str, Any]] = []
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


def _locate_pptx(dir_path: str) -> str:
    """返回目录内首个 .pptx 的完整路径；目录中没有时返回空串。"""
    names = [f for f in os.listdir(dir_path)
             if f.lower().endswith(".pptx") and not f.startswith("~$")]
    if not names:
        return ""
    return os.path.join(dir_path, sorted(names)[0])


def evaluate(dir_path: str) -> dict[str, Any]:
    result: dict[str, Any] = {
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
        file_path = _locate_pptx(dir_path)
        if not file_path:
            result["status"] = "error"
            result["error"] = "目录中未找到 .pptx 文件：%s" % dir_path
            return result
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


if __name__ == "__main__":
    _dir = sys.argv[1] if len(sys.argv) >= 2 else os.path.dirname(os.path.abspath(__file__))
    print(json.dumps(evaluate(_dir), ensure_ascii=False, indent=2))
















