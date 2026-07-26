#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
自动评估“求学申请材料_林绯绯_志愿服务版.pdf”。

按批量评估统一约定，本脚本对外只暴露一个函数：
    evaluate(dir_path: str) -> dict
- 入参 dir_path 为脚本所在目录路径，脚本自行在该目录下定位并打开被评估文档。
- 返回结构化字典（含维度一通过与否、维度二逐项得分、总分），主结果不走 print。

评分逻辑：
1. 先检查维度1（PDF 格式、可正常打开）。维度1不通过，直接以 0 分记结果。
2. 维度1通过后，按维度2的所有得分点/扣分点逐条自动检测并累计分数。

依赖：pdfplumber（经 pdf_backend 适配层访问）。
"""

from __future__ import annotations

import json
import math
import os
import re
import sys
from dataclasses import dataclass
from typing import Iterable, Optional

try:
    try:
        import pdf_backend
    except ImportError:
        from verifiers import pdf_backend
except Exception as exc:  # pragma: no cover
    print("缺少依赖 pdfplumber：请先安装 pip install pdfplumber", file=sys.stderr)
    raise


class _Page:
    """单页访问句柄：封装 pdf_backend 文档的页级接口。"""

    def __init__(self, doc, index: int):
        self._doc = doc
        self.index = index

    @property
    def rect(self):
        w, h = self._doc.page_size(self.index)
        return pdf_backend.PdfRect(0.0, 0.0, w, h)

    def get_text(self) -> str:
        return self._doc.page_text(self.index)

    def span_lines(self):
        return self._doc.extract_span_lines(self.index)

    def get_drawings(self):
        return [{"rect": d.rect, "fill": d.fill}
                for d in self._doc.extract_drawings(self.index)]

SCRIPT_ID = "032"
PREFERRED_PDF_NAME = "求学申请材料_林绯绯_志愿服务版.pdf"

DARK_GREEN = (0.141176, 0.368627, 0.360784)
GOLD = (0.831373, 0.650980, 0.345098)

MAIN_TITLES = [
    "教育背景",
    "实践经历",
    "技能证书与获奖",
    "志愿服务",
    "研修规划",
]

VOLUNTEER_1 = "2022年1-2月在社区做抗击疫情先进志愿者"
VOLUNTEER_2 = "2023年7-9月在山底村做志愿者"
SKILL_ADDED = "2021年11月-2023年9月参加青年马克思主义者培养工程，培训合格"
BANNED_PRACTICE = "项目内容聚焦设备运行状态识别与远程诊断，未保留原始业务数据。"
FOOTER_TEXT = "求学申请材料 · 2026"

RUBRIC_TEXT = {
    "+3 实践经历未出现禁用文本": "第1页实践经历段落：实践经历模块中未出现文本“项目内容聚焦设备运行状态识别与远程诊断，未保留原始业务数据。”",
    "+3 志愿服务模块标题条": "第1页志愿服务模块标题条：志愿服务模块位于右侧主内容区，放置在“技能证书与获奖”模块之后、“研修规划”模块之前，或在不破坏一页排版的情况下位于右侧内容区合适位置；新增“志愿服务”模块，标题条样式与“教育背景”“实践经历”“技能证书与获奖”“研修规划”模块一致，使用深绿色圆角矩形条，左侧有金色竖向小装饰块，标题文字为宋体或近似宋体，字号为五号或近似五号大小，白色加粗。",
    "+5 志愿服务第一/第二行文本": "第1页志愿服务第一行文本：新增文本“2022年1-2月在社区做抗击疫情先进志愿者”，作为志愿服务模块正文第一行，字体为宋体或近似宋体，字号为小五号或近似小五号大小，颜色为黑色或深灰色。第二行文本：新增文本“2023年7-9月在山底村做志愿者”，作为志愿服务模块正文第二行，与第一行分行显示，字体、字号、颜色与第一行一致。",
    "+3 志愿服务正文格式": "第1页志愿服务正文格式：两段志愿服务文本左对齐，行距与“技能证书与获奖”正文一致，段前段后为0磅或与原简历正文一致，不出现项目符号混乱或过大行距。",
    "+3 志愿服务模块间距": "第1页志愿服务模块间距：志愿服务模块标题与上方模块、下方模块之间保持单倍行间距或近似单倍间距，正文与标题条之间间距与正文间间距一样，不与其他模块重叠。",
    "+5 技能证书与获奖新增行": "第1页技能证书与获奖模块新增行：在“技能证书与获奖”模块中另起一行添加文本“2021年11月-2023年9月参加青年马克思主义者培养工程，培训合格”字体为宋体或近似宋体，字号为小五号或近似小五号大小。位于技能证书与获奖模块原有条目之后、下一模块之前，不能插入到标题条内部或覆盖原有条目。",
    "+3 技能证书新增行项目符号": "第1页技能证书新增行项目符号：若技能证书与获奖模块原有条目使用金色圆点项目符号，新增行也使用相同金色圆点；若原有条目无项目符号，新增行也保持无项目符号，格式一致。",
    "-5 实践经历下方大空白/多页": "交付简历“实践经历”下方不可出现超过页面50%的空白导致简历变成2页或更多页。",
    "-3 志愿服务两段未分行": "志愿服务两段文字没有分行显示。",
    "-3 志愿服务标题样式不一致": "“志愿服务”模块标题样式与其他模块标题条明显不一致。",
    "-3 志愿服务内容位置错误": "志愿服务内容被放入左侧信息栏、页眉、页脚或页面空白处，未作为右侧正文模块呈现。",
    "-3 技能证书新增内容覆盖原文": "技能证书新增内容覆盖原有“英语能力”“工具能力”或“综合表现”文本。",
    "-3 研修规划正文截断/越界": "新增内容后“研修规划”模块正文被截断或超出页面底部。",
    "-3 成绩表数字缺失": "新增内容后成绩表中的64、71、118、96、349任意一个文本缺失。",
    "-3 页脚消失或被覆盖": "第1页底部页脚“求学申请材料 · 2026”被正文内容覆盖或完全消失。",
    "-3 研修规划移到教育背景前": "新增内容后“研修规划”不可被移动到教育背景前。",
}


@dataclass
class Line:
    text: str
    x0: float
    y0: float
    x1: float
    y1: float
    size: float
    font: str
    color: int
    flags: int = 0
    duplicate_count: int = 1

    @property
    def cx(self) -> float:
        return (self.x0 + self.x1) / 2

    @property
    def cy(self) -> float:
        return (self.y0 + self.y1) / 2

    @property
    def height(self) -> float:
        return self.y1 - self.y0


@dataclass
class Hit:
    code: str
    score: int
    matched: bool
    detail: str


def normalize_text(text: str) -> str:
    """去除空白，保留中文标点，便于跨行文本检测。"""
    return re.sub(r"\s+", "", text or "")


def color_to_rgb(color: int) -> tuple[int, int, int]:
    return ((color >> 16) & 255, (color >> 8) & 255, color & 255)


def rgb_distance_int(color: int, target: tuple[int, int, int]) -> float:
    rgb = color_to_rgb(color)
    return math.sqrt(sum((rgb[i] - target[i]) ** 2 for i in range(3)))


def rgb_distance_float(color: Optional[tuple[float, float, float]], target: tuple[float, float, float]) -> float:
    if color is None:
        return 999.0
    return math.sqrt(sum((color[i] - target[i]) ** 2 for i in range(3)))


def rects_overlap(a: tuple[float, float, float, float], b: tuple[float, float, float, float], pad: float = 0) -> bool:
    ax0, ay0, ax1, ay1 = a
    bx0, by0, bx1, by1 = b
    return not (ax1 + pad < bx0 or bx1 + pad < ax0 or ay1 + pad < by0 or by1 + pad < ay0)


def line_rect(line: Line) -> tuple[float, float, float, float]:
    return (line.x0, line.y0, line.x1, line.y1)


def extract_lines(page) -> tuple[list[Line], list[Line]]:
    """返回去重后的行、原始行。该 PDF 中加粗/阴影常以微小位移重复文本呈现。"""
    raw: list[Line] = []
    for sp in page.span_lines():
        text = sp.text.strip()
        if not text:
            continue
        raw.append(Line(text, sp.bbox.x0, sp.bbox.y0, sp.bbox.x1, sp.bbox.y1,
                        sp.size, sp.font, sp.color, sp.flags))

    deduped: list[Line] = []
    for ln in sorted(raw, key=lambda l: (round(l.y0, 1), round(l.x0, 1), l.text)):
        found: Optional[Line] = None
        for old in deduped:
            if (
                old.text == ln.text
                and abs(old.x0 - ln.x0) <= 0.8
                and abs(old.y0 - ln.y0) <= 0.8
                and abs(old.x1 - ln.x1) <= 0.8
                and abs(old.y1 - ln.y1) <= 0.8
            ):
                found = old
                break
        if found:
            found.duplicate_count += 1
        else:
            deduped.append(ln)
    return deduped, raw


def page_text(lines: Iterable[Line]) -> str:
    return "\n".join(l.text for l in sorted(lines, key=lambda l: (l.y0, l.x0)))


def find_lines(lines: list[Line], needle: str, *, exact: bool = False, min_x: Optional[float] = None) -> list[Line]:
    n = normalize_text(needle)
    out = []
    for line in lines:
        if min_x is not None and line.x0 < min_x:
            continue
        t = normalize_text(line.text)
        if (exact and t == n) or (not exact and n in t):
            out.append(line)
    return sorted(out, key=lambda l: (l.y0, l.x0))


def first_line(lines: list[Line], needle: str, *, exact: bool = False, min_x: Optional[float] = None) -> Optional[Line]:
    matches = find_lines(lines, needle, exact=exact, min_x=min_x)
    return matches[0] if matches else None


def module_region(lines: list[Line], title: str, next_titles: Iterable[str], *, min_x: float = 180) -> tuple[Optional[Line], list[Line]]:
    title_line = first_line(lines, title, exact=True, min_x=min_x)
    if not title_line:
        return None, []
    next_candidates = []
    for nt in next_titles:
        ln = first_line(lines, nt, exact=True, min_x=min_x)
        if ln and ln.y0 > title_line.y0 + 1:
            next_candidates.append(ln)
    y_end = min((ln.y0 for ln in next_candidates), default=10_000)
    # 没有下一个模块时，不把页脚误判为模块正文；常规页脚位于页面最底部约 25pt 区域。
    if y_end == 10_000:
        y_end = float(getattr(title_line, "page_height", 10_000)) if hasattr(title_line, "page_height") else 10_000
    body = [
        l
        for l in lines
        if l.y0 > title_line.y1 + 1 and l.y0 < y_end - 1 and l.x0 >= min_x and "求学申请材料 · 2026" not in l.text
    ]
    return title_line, sorted(body, key=lambda l: (l.y0, l.x0))


def drawings(page) -> list[dict]:
    return page.get_drawings()


def drawing_rect(d: dict) -> Optional[tuple[float, float, float, float]]:
    r = d.get("rect")
    if not r:
        return None
    return (float(r.x0), float(r.y0), float(r.x1), float(r.y1))


def has_dark_green_title_bar(page, title_line: Line) -> bool:
    """检测标题文字背后的深绿色圆角矩形条：用若干块深绿色绘图拼成也算。"""
    cy = title_line.cy
    candidates = []
    for d in drawings(page):
        r = drawing_rect(d)
        if not r:
            continue
        x0, y0, x1, y1 = r
        fill = d.get("fill")
        if rgb_distance_float(fill, DARK_GREEN) <= 0.08 and y0 <= cy <= y1 and x0 <= title_line.x0 <= x1 + 15:
            candidates.append(r)
    if not candidates:
        return False
    min_x = min(r[0] for r in candidates)
    max_x = max(r[2] for r in candidates)
    min_y = min(r[1] for r in candidates)
    max_y = max(r[3] for r in candidates)
    return min_x <= title_line.x0 - 3 and max_x >= max(title_line.x1 + 20, 400) and 14 <= max_y - min_y <= 26


def has_gold_decorator(page, title_line: Line) -> bool:
    cy = title_line.cy
    for d in drawings(page):
        r = drawing_rect(d)
        if not r:
            continue
        x0, y0, x1, y1 = r
        fill = d.get("fill")
        if rgb_distance_float(fill, GOLD) <= 0.08 and y0 - 2 <= cy <= y1 + 2 and title_line.x0 - 18 <= x0 <= title_line.x0:
            if 2 <= x1 - x0 <= 8 and 4 <= y1 - y0 <= 14:
                return True
    return False


def near_song_font(font: str) -> bool:
    f = (font or "").lower()
    # PDF 子集字体名不一定直接叫 SimSun/宋体；Bousung/GBSong/serif 通常可视为宋体或近似宋体。
    return any(k in f for k in ["song", "sung", "bousung", "serif", "ming", "stkaiti", "kai"])


def is_dark_text(color: int) -> bool:
    r, g, b = color_to_rgb(color)
    return max(r, g, b) <= 90


def is_white_text(color: int) -> bool:
    return rgb_distance_int(color, (255, 255, 255)) <= 35


def is_bold_text(line: Line) -> bool:
    f = (line.font or "").lower()
    return "bold" in f or "black" in f or "heavy" in f or bool(line.flags & 16) or line.duplicate_count >= 2


def title_style_ok(page, title_line: Optional[Line]) -> bool:
    if not title_line:
        return False
    return (
        has_dark_green_title_bar(page, title_line)
        and has_gold_decorator(page, title_line)
        and near_song_font(title_line.font)
        and 9.5 <= title_line.size <= 11.5
        and is_white_text(title_line.color)
        and is_bold_text(title_line)
    )


def body_text_style_ok(line: Optional[Line]) -> bool:
    if not line:
        return False
    return near_song_font(line.font) and 8.0 <= line.size <= 10.2 and is_dark_text(line.color)


def gold_bullets_near_line(page, line: Line, *, x_min: float = 180, x_max: Optional[float] = None) -> list[tuple[float, float, float, float]]:
    if x_max is None:
        x_max = line.x0 + 2
    out = []
    for d in drawings(page):
        r = drawing_rect(d)
        if not r:
            continue
        x0, y0, x1, y1 = r
        fill = d.get("fill")
        if rgb_distance_float(fill, GOLD) <= 0.08 and x_min <= x0 <= x_max:
            if abs(((y0 + y1) / 2) - line.cy) <= 5 and 2 <= x1 - x0 <= 8 and 2 <= y1 - y0 <= 8:
                out.append(r)
    return out


def lines_overlap(a: Line, b: Line) -> bool:
    return rects_overlap(line_rect(a), line_rect(b), pad=0)


def _evaluate_pdf(pdf_path: str) -> tuple[bool, list[Hit], int, list[str]]:
    notes: list[str] = []

    # 维度1：PDF 格式且可正常打开。
    if not os.path.isfile(pdf_path):
        return False, [Hit("维度1", 0, False, f"文件不存在：{pdf_path}")], 0, notes
    if os.path.splitext(pdf_path)[1].lower() != ".pdf":
        return False, [Hit("维度1", 0, False, "交付文件不是 .pdf 格式")], 0, notes
    try:
        doc = pdf_backend.open_pdf(pdf_path)
        page_count = doc.page_count
        if page_count <= 0:
            return False, [Hit("维度1", 0, False, "PDF可打开但页数为0")], 0, notes
        first_page = _Page(doc, 0)
        # 触发一次解析，避免损坏 PDF 到维度2才报错。
        _ = first_page.get_text()
    except Exception as exc:
        return False, [Hit("维度1", 0, False, f"PDF无法正常打开或解析：{exc}")], 0, notes

    page = _Page(doc, 0)
    lines, raw_lines = extract_lines(page)
    text = page_text(lines)
    compact_text = normalize_text(text)
    width, height = float(page.rect.width), float(page.rect.height)
    notes.append(f"PDF可打开，共 {page_count} 页；第1页尺寸约 {width:.1f}×{height:.1f} pt。")

    hits: list[Hit] = [Hit("维度1", 0, True, "交付文件为PDF格式，且可正常打开。")]

    # 常用定位。
    education_title = first_line(lines, "教育背景", exact=True, min_x=180)
    practice_title, practice_body = module_region(lines, "实践经历", ["技能证书与获奖", "志愿服务", "研修规划"], min_x=180)
    skill_title, skill_body = module_region(lines, "技能证书与获奖", ["志愿服务", "研修规划"], min_x=180)
    volunteer_title, volunteer_body = module_region(lines, "志愿服务", ["研修规划"], min_x=180)
    plan_title, plan_body = module_region(lines, "研修规划", [], min_x=180)

    v1 = first_line(lines, VOLUNTEER_1, min_x=None)
    v2 = first_line(lines, VOLUNTEER_2, min_x=None)
    skill_added = first_line(lines, SKILL_ADDED, min_x=None)

    # +3：实践经历模块（标题行+正文）中未出现指定禁用文本。
    practice_all_lines = ([practice_title] if practice_title else []) + practice_body
    practice_region_text = normalize_text(page_text(practice_all_lines)) if practice_title else ""
    if practice_title and normalize_text(BANNED_PRACTICE) not in practice_region_text:
        hits.append(Hit("+3 实践经历未出现禁用文本", 3, True, "实践经历模块内未检测到指定禁用文本。"))
    else:
        detail = "未能定位实践经历模块。" if not practice_title else "实践经历模块内出现了指定禁用文本。"
        hits.append(Hit("+3 实践经历未出现禁用文本", 3, False, detail))

    # +3：志愿服务标题条位置与样式。
    volunteer_title_position_ok = False
    if volunteer_title:
        in_right = volunteer_title.x0 >= 180
        after_skill = skill_title is None or volunteer_title.y0 > skill_title.y0
        before_plan = plan_title is None or volunteer_title.y0 < plan_title.y0
        standard_order = after_skill and before_plan
        suitable_right_position = in_right and page_count == 1 and volunteer_title.y1 < height - 35
        volunteer_title_position_ok = in_right and (standard_order or suitable_right_position)
    volunteer_title_style_checks = {
        "深绿色圆角矩形条": bool(volunteer_title and has_dark_green_title_bar(page, volunteer_title)),
        "金色竖向小装饰块": bool(volunteer_title and has_gold_decorator(page, volunteer_title)),
        "宋体或近似宋体": bool(volunteer_title and near_song_font(volunteer_title.font)),
        "五号或近似五号大小": bool(volunteer_title and 9.5 <= volunteer_title.size <= 11.5),
        "白色": bool(volunteer_title and is_white_text(volunteer_title.color)),
        "加粗": bool(volunteer_title and is_bold_text(volunteer_title)),
    }
    volunteer_title_style = all(volunteer_title_style_checks.values())
    if volunteer_title and volunteer_title_position_ok and volunteer_title_style:
        hits.append(Hit("+3 志愿服务模块标题条", 3, True, "志愿服务位于右侧主内容区，放置在技能证书与获奖之后、研修规划之前，或在不破坏一页排版的情况下位于右侧内容区合适位置；标题条为深绿色圆角矩形条，左侧有金色竖向小装饰块，标题文字为近宋体五号白色加粗。"))
    else:
        reasons = []
        if not volunteer_title:
            reasons.append("未检测到右侧主内容区的“志愿服务”标题")
        elif not volunteer_title_position_ok:
            reasons.append("未位于右侧主内容区，或未满足技能证书与获奖之后、研修规划之前/一页排版内右侧合适位置")
        missing_style = [name for name, ok in volunteer_title_style_checks.items() if not ok]
        if missing_style:
            reasons.append("标题条样式缺少" + "、".join(missing_style))
        hits.append(Hit("+3 志愿服务模块标题条", 3, False, "；".join(reasons)))

    # +5：志愿服务第一、第二行正文文本。
    # 第一行：v1 存在，字体宋体/近似宋体，字号小五号/近似小五号（8~10.2pt），颜色黑色或深灰色。
    v1_exists = bool(v1)
    v1_font_ok = bool(v1 and near_song_font(v1.font))
    v1_size_ok = bool(v1 and 8.0 <= v1.size <= 10.2)
    v1_color_ok = bool(v1 and is_dark_text(v1.color))
    v1_style_ok = v1_font_ok and v1_size_ok and v1_color_ok
    # 第二行：v2 存在，与第一行分行显示（y坐标不同），字体/字号/颜色与第一行一致。
    v2_exists = bool(v2)
    v_lines_separate = bool(v1 and v2 and abs(v1.y0 - v2.y0) > 3)
    v2_font_consistent = bool(v1 and v2 and near_song_font(v2.font))
    v2_size_consistent = bool(v1 and v2 and abs(v1.size - v2.size) <= 0.5)
    v2_color_consistent = bool(v1 and v2 and rgb_distance_int(v1.color, color_to_rgb(v2.color)) <= 20)
    v2_style_consistent = v2_font_consistent and v2_size_consistent and v2_color_consistent
    # 第一行须是志愿服务模块正文第一行，第二行须是第二行。
    volunteer_body_lines = [l for l in volunteer_body] if volunteer_title else []
    v1_is_first_body_line = bool(
        v1 and volunteer_body_lines
        and normalize_text(VOLUNTEER_1) in normalize_text(volunteer_body_lines[0].text)
    )
    v2_is_second_body_line = bool(
        v2 and len(volunteer_body_lines) >= 2
        and normalize_text(VOLUNTEER_2) in normalize_text(volunteer_body_lines[1].text)
    )
    all_ok = (
        v1_exists and v1_style_ok and v1_is_first_body_line
        and v2_exists and v_lines_separate and v2_style_consistent and v2_is_second_body_line
    )
    if all_ok:
        hits.append(Hit("+5 志愿服务第一/第二行文本", 5, True,
            "第一行文本存在，字体宋体/近似宋体，字号小五号/近似小五号，颜色黑色或深灰色，作为志愿服务正文第一行；"
            "第二行文本存在，与第一行分行显示，字体/字号/颜色与第一行一致，作为正文第二行。"))
    else:
        reasons = []
        if not v1_exists:
            reasons.append('缺少第一行指定文本‘2022年1-2月在社区做抗击疫情先进志愿者’')
        else:
            if not v1_font_ok:
                reasons.append("第一行字体不是宋体或近似宋体")
            if not v1_size_ok:
                reasons.append(f"第一行字号 {v1.size:.1f}pt 不符合小五号/近似小五号（8~10.2pt）")
            if not v1_color_ok:
                reasons.append("第一行颜色不是黑色或深灰色")
            if not v1_is_first_body_line:
                reasons.append("第一行指定文本不是志愿服务模块正文第一行")
        if not v2_exists:
            reasons.append(f"缺少第二行指定文本‘{VOLUNTEER_2}’")
        else:
            if v1 and not v_lines_separate:
                reasons.append("第二行与第一行未分行显示")
            if not v2_font_consistent:
                reasons.append("第二行字体与第一行不一致")
            if not v2_size_consistent:
                reasons.append(f"第二行字号与第一行不一致（差值超过0.5pt）")
            if not v2_color_consistent:
                reasons.append("第二行颜色与第一行不一致")
            if not v2_is_second_body_line:
                reasons.append("第二行指定文本不是志愿服务模块正文第二行")
        hits.append(Hit("+5 志愿服务第一/第二行文本", 5, False, "；".join(reasons)))

    # +3：志愿服务正文格式。
    volunteer_format_ok = False
    volunteer_format_reasons: list[str] = []
    if v1 and v2:
        # 1. 两段文本左对齐。
        left_aligned = abs(v1.x0 - v2.x0) <= 3

        # 2. 行距与”技能证书与获奖”正文一致。
        line_gap = v2.y0 - v1.y0
        skill_y_positions = sorted({round(l.y0, 2) for l in skill_body})
        skill_line_gaps = [b - a for a, b in zip(skill_y_positions, skill_y_positions[1:]) if b - a > 3]
        reference_gap = sorted(skill_line_gaps)[len(skill_line_gaps) // 2] if skill_line_gaps else None
        line_gap_matches_skill = bool(reference_gap is not None and abs(line_gap - reference_gap) <= 1.5)

        # 3. 过大行距：行距不超过参考行距的2倍且不超过40pt绝对上限。
        no_excessive_gap = line_gap <= (reference_gap * 2 if reference_gap else 40) and line_gap <= 40

        # 4. 段前段后为0磅或与原简历正文一致：检测v1之前、v2之后与相邻行的间距是否在正文正常范围内。
        # 取技能证书与获奖正文第一行与其标题底部的间距作为”正文与标题间距”（段前）参照。
        skill_title_to_body_gap: Optional[float] = None
        if skill_title and skill_body:
            first_skill_body = min(skill_body, key=lambda l: l.y0)
            skill_title_to_body_gap = first_skill_body.y0 - skill_title.y1
        # 志愿服务标题底部到第一行正文的间距。
        v1_para_before = (v1.y0 - volunteer_title.y1) if volunteer_title else None
        # 容忍范围：与技能参照间距相差不超过3pt，或落在[0, 20]pt合理区间内。
        if v1_para_before is not None and skill_title_to_body_gap is not None:
            para_before_ok = abs(v1_para_before - skill_title_to_body_gap) <= 3 or 0 <= v1_para_before <= 20
        else:
            para_before_ok = v1_para_before is None or (0 <= v1_para_before <= 20)

        # 4b. 段后：v2 下沿到下一模块/下方最近正文内容的间距。
        # 参照标准：技能证书正文最后一行到"志愿服务"标题条的间距（即简历中"正文→下一模块标题"的实际段后）。
        skill_body_to_next_title_gap: Optional[float] = None
        if skill_body and volunteer_title:
            last_skill_body = max(skill_body, key=lambda l: l.y1)
            skill_body_to_next_title_gap = volunteer_title.y0 - last_skill_body.y1
        # 志愿服务下方最近的内容 y0：优先取"研修规划"标题；若不存在，取整页中位于 v2 之下的最高非志愿服务行。
        next_after_v2_y0: Optional[float] = None
        if plan_title is not None and plan_title.y0 > v2.y1:
            next_after_v2_y0 = plan_title.y0
        else:
            volunteer_line_ids = {id(v1), id(v2)} | ({id(volunteer_title)} if volunteer_title else set())
            below: list[float] = []
            for line in lines:
                if id(line) in volunteer_line_ids:
                    continue
                if line.y0 > v2.y1:
                    below.append(line.y0)
            if below:
                next_after_v2_y0 = min(below)
        v2_para_after: Optional[float] = None
        if next_after_v2_y0 is not None:
            v2_para_after = next_after_v2_y0 - v2.y1
        # 容忍：与参照相差不超过3pt，或落在[0, 20]pt区间；找不到下方内容视为满足（v2 已到页尾）。
        if v2_para_after is None:
            para_after_ok = True
        elif skill_body_to_next_title_gap is not None:
            para_after_ok = abs(v2_para_after - skill_body_to_next_title_gap) <= 3 or 0 <= v2_para_after <= 20
        else:
            para_after_ok = 0 <= v2_para_after <= 20

        # 5. 不出现项目符号混乱：文本型项目符号检测 + 绘制型金色圆点一致性。
        text_bullets_ok = (
            not v1.text.strip().startswith(("•", "·", "-"))
            and not v2.text.strip().startswith(("•", "·", "-"))
        )
        drawn_bullets = [bool(gold_bullets_near_line(page, l, x_min=180, x_max=l.x0)) for l in (v1, v2)]
        drawn_bullets_consistent = drawn_bullets[0] == drawn_bullets[1]
        no_bullet_chaos = text_bullets_ok and drawn_bullets_consistent

        volunteer_format_ok = (
            left_aligned
            and line_gap_matches_skill
            and no_excessive_gap
            and para_before_ok
            and para_after_ok
            and no_bullet_chaos
        )
        if not left_aligned:
            volunteer_format_reasons.append("两段文本未左对齐")
        if not line_gap_matches_skill:
            if reference_gap is None:
                volunteer_format_reasons.append("无法取得技能证书与获奖正文行距作为参照")
            else:
                volunteer_format_reasons.append(f"志愿服务行距 {line_gap:.1f}pt 与技能证书与获奖正文行距 {reference_gap:.1f}pt 不一致")
        if not no_excessive_gap:
            volunteer_format_reasons.append(f"行距 {line_gap:.1f}pt 过大")
        if not para_before_ok:
            volunteer_format_reasons.append("段前间距与原简历正文不一致")
        if not para_after_ok:
            ref_txt = f"，参照 {skill_body_to_next_title_gap:.1f}pt" if skill_body_to_next_title_gap is not None else ""
            volunteer_format_reasons.append(
                f"段后间距 {v2_para_after:.1f}pt 既不接近0磅也不与原简历正文一致{ref_txt}"
                if v2_para_after is not None else "段后间距未通过校验"
            )
        if not no_bullet_chaos:
            volunteer_format_reasons.append("存在项目符号混乱")
    else:
        volunteer_format_reasons.append("缺少两段志愿服务文本，无法判断正文格式")
    if volunteer_format_ok:
        hits.append(Hit("+3 志愿服务正文格式", 3, True, "两段志愿服务文本左对齐，行距与技能证书与获奖正文一致，段前段后与原简历正文一致，未检测到项目符号混乱或过大行距。"))
    else:
        hits.append(Hit("+3 志愿服务正文格式", 3, False, "；".join(volunteer_format_reasons)))

    # +3：志愿服务模块间距。
    spacing_ok = False
    spacing_reasons: list[str] = []
    if volunteer_title and skill_body and plan_title and v1 and v2:
        # 用技能证书与获奖正文的实际行距作为"原简历单倍行距"参照。
        skill_y_positions = sorted({round(l.y0, 2) for l in skill_body})
        skill_gaps = [b - a for a, b in zip(skill_y_positions, skill_y_positions[1:]) if b - a > 3]
        ref_line_gap = sorted(skill_gaps)[len(skill_gaps) // 2] if skill_gaps else (v2.y0 - v1.y0)

        # 近似单倍间距：允许偏差为参照行距的50%（上下模块间距通常略大于行距）。
        def near_single_gap(gap: float) -> bool:
            return gap > 0 and gap <= ref_line_gap * 2.0

        upper_lines = [l for l in skill_body if l.y1 < volunteer_title.y0]
        gap_prev = volunteer_title.y0 - max(l.y1 for l in upper_lines) if upper_lines else None
        gap_title_body = v1.y0 - volunteer_title.y1
        gap_body_next = plan_title.y0 - v2.y1

        # 1. 标题与上方模块间距：近似单倍间距。
        upper_gap_ok = bool(gap_prev is not None and near_single_gap(gap_prev))
        # 2. 标题与下方模块间距：近似单倍间距。
        lower_gap_ok = near_single_gap(gap_body_next)
        # 3. 正文与标题条之间间距与正文行间间距一样：与参考行距相差不超过参照的50%。
        title_body_gap_ok = gap_title_body > 0 and abs(gap_title_body - ref_line_gap) <= ref_line_gap * 0.5
        # 4. 不与其他模块重叠。
        volunteer_rects = [line_rect(l) for l in (volunteer_title, v1, v2)]
        other_module_lines = skill_body + [plan_title] + plan_body
        no_overlap = not any(
            rects_overlap(vr, line_rect(other), pad=0)
            for vr in volunteer_rects
            for other in other_module_lines
        )

        spacing_ok = upper_gap_ok and lower_gap_ok and title_body_gap_ok and no_overlap
        if not upper_gap_ok:
            if gap_prev is None:
                spacing_reasons.append("无法取得志愿服务模块与上方模块之间的间距")
            else:
                spacing_reasons.append(f"志愿服务标题与上方模块间距 {gap_prev:.1f}pt 不是单倍或近似单倍行距（参照 {ref_line_gap:.1f}pt）")
        if not lower_gap_ok:
            spacing_reasons.append(f"志愿服务正文与下方模块间距 {gap_body_next:.1f}pt 不是单倍或近似单倍行距（参照 {ref_line_gap:.1f}pt）")
        if not title_body_gap_ok:
            spacing_reasons.append(f"正文与标题条之间间距 {gap_title_body:.1f}pt 与正文行间间距 {ref_line_gap:.1f}pt 不一致")
        if not no_overlap:
            spacing_reasons.append("志愿服务模块与其他模块重叠")
    else:
        spacing_reasons.append("缺少志愿服务标题、两段正文、上方模块或下方模块，无法判断模块间距")
    if spacing_ok:
        hits.append(Hit("+3 志愿服务模块间距", 3, True, "志愿服务模块标题与上方模块、下方模块之间保持单倍或近似单倍间距，正文与标题条之间间距与正文间距一致，且未与其他模块重叠。"))
    else:
        hits.append(Hit("+3 志愿服务模块间距", 3, False, "；".join(spacing_reasons)))

    # +5：技能证书与获奖新增行。
    original_skill_lines = [
        first_line(lines, "英语能力", min_x=180),
        first_line(lines, "工具能力", min_x=180),
        first_line(lines, "综合表现", min_x=180),
    ]
    original_skill_lines = [l for l in original_skill_lines if l]
    skill_added_ok = False
    skill_added_reasons: list[str] = []
    if skill_added and skill_title:
        # 1. 新增行须位于“技能证书与获奖”模块正文区。
        in_skill_module = skill_added.x0 >= 180 and skill_added.y0 > skill_title.y1
        # 2. 另起一行：不能与原有技能证书条目位于同一行。
        separate_line = not any(abs(skill_added.y0 - l.y0) <= 3 for l in original_skill_lines)
        # 3. 字体为宋体或近似宋体。
        font_ok = near_song_font(skill_added.font)
        # 4. 字号为小五号或近似小五号大小。
        size_ok = 8.0 <= skill_added.size <= 10.2
        # 5. 位于技能证书与获奖模块原有条目之后。
        after_original = bool(original_skill_lines) and skill_added.y0 > max(l.y0 for l in original_skill_lines)
        # 6. 位于下一模块之前。
        before_next = (volunteer_title is None or skill_added.y1 < volunteer_title.y0) and (plan_title is None or skill_added.y1 < plan_title.y0)
        # 7. 不能插入到标题条内部。
        not_in_title = skill_added.y0 > skill_title.y1 + 2
        # 8. 不能覆盖原有条目。
        not_cover_original = not any(rects_overlap(line_rect(skill_added), line_rect(l), pad=0) for l in original_skill_lines)

        skill_added_ok = (
            in_skill_module
            and separate_line
            and font_ok
            and size_ok
            and after_original
            and before_next
            and not_in_title
            and not_cover_original
        )
        if not in_skill_module:
            skill_added_reasons.append("新增文本不在技能证书与获奖模块正文区")
        if not separate_line:
            skill_added_reasons.append("新增文本未另起一行")
        if not font_ok:
            skill_added_reasons.append("新增文本字体不是宋体或近似宋体")
        if not size_ok:
            skill_added_reasons.append(f"新增文本字号 {skill_added.size:.1f}pt 不符合小五号或近似小五号大小")
        if not after_original:
            skill_added_reasons.append("新增文本未位于原有技能证书条目之后")
        if not before_next:
            skill_added_reasons.append("新增文本未位于下一模块之前")
        if not not_in_title:
            skill_added_reasons.append("新增文本插入到技能证书与获奖标题条内部")
        if not not_cover_original:
            skill_added_reasons.append("新增文本覆盖原有技能证书条目")
    else:
        if not skill_title:
            skill_added_reasons.append("未检测到技能证书与获奖模块标题")
        if not skill_added:
            skill_added_reasons.append(f"未检测到新增文本‘{SKILL_ADDED}’")
    if skill_added_ok:
        hits.append(Hit("+5 技能证书与获奖新增行", 5, True, "指定青年马克思主义者培养工程文本位于技能证书与获奖模块中，另起一行，字体为宋体或近似宋体，字号为小五号或近似小五号，位于原有条目之后、下一模块之前，未插入标题条内部且未覆盖原有条目。"))
    else:
        hits.append(Hit("+5 技能证书与获奖新增行", 5, False, "；".join(skill_added_reasons)))

    # +3：技能证书新增行项目符号一致。
    skill_bullet_ok = False
    skill_bullet_reasons: list[str] = []
    if skill_added and original_skill_lines:
        original_has_bullets = [bool(gold_bullets_near_line(page, l, x_min=180, x_max=l.x0)) for l in original_skill_lines]
        added_has_bullet = bool(gold_bullets_near_line(page, skill_added, x_min=180, x_max=skill_added.x0))
        originals_all_have = all(original_has_bullets)
        originals_none_have = not any(original_has_bullets)
        if originals_all_have:
            # 原有条目使用金色圆点，新增行也须使用金色圆点。
            skill_bullet_ok = added_has_bullet
            if not added_has_bullet:
                skill_bullet_reasons.append("原有技能证书条目使用金色圆点项目符号，但新增行未检测到金色圆点")
        elif originals_none_have:
            # 原有条目无项目符号，新增行也须无项目符号。
            skill_bullet_ok = not added_has_bullet
            if added_has_bullet:
                skill_bullet_reasons.append("原有技能证书条目无项目符号，但新增行出现了金色圆点")
        else:
            # 原有条目项目符号不统一，无法按细则判断一致性。
            skill_bullet_ok = False
            skill_bullet_reasons.append("原有技能证书条目项目符号不统一，无法判断新增行是否一致")
    else:
        if not skill_added:
            skill_bullet_reasons.append("未检测到新增文本，无法判断项目符号")
        if not original_skill_lines:
            skill_bullet_reasons.append("未检测到原有技能证书条目，无法判断项目符号")
    if skill_bullet_ok:
        hits.append(Hit("+3 技能证书新增行项目符号", 3, True, "新增行项目符号与原有技能证书条目保持一致。"))
    else:
        hits.append(Hit("+3 技能证书新增行项目符号", 3, False, "；".join(skill_bullet_reasons)))

    # 扣分项 -5：实践经历下方出现超过页面50%的空白导致简历变成2页或更多页。
    # 修正：不再用「页面高度 - 实践经历正文底部」笼统当作空白，那样会把正常排在
    # 实践经历下方的技能证书/志愿服务/研修规划模块也算进空白。这里改为计算
    # 实践经历正文底部以下「连续无文本、无绘图占用」的真实最大空白高度。
    excessive_blank_or_multipage = False
    excessive_reasons: list[str] = []
    if page_count >= 2:
        # 条件一：简历变成2页或更多页。
        excessive_blank_or_multipage = True
        excessive_reasons.append(f"PDF共 {page_count} 页，超过1页")
    if practice_body:
        practice_bottom = max(b.y1 for b in practice_body)

        # 收集第1页上位于实践经历正文底部以下的所有「内容占用」纵向区间：
        # 文本行 + 绘图矩形（标题条、装饰块、项目符号、分隔线等都算内容）。
        # 页脚版权行不算模块内容，但作为页面底部内容仍可界定空白下界。
        footer = first_line(lines, FOOTER_TEXT)
        content_intervals: list[tuple[float, float]] = []
        for l in lines:
            if l.y1 > practice_bottom + 0.5:
                content_intervals.append((max(l.y0, practice_bottom), l.y1))
        for d in drawings(page):
            r = drawing_rect(d)
            if r is None:
                continue
            dy0, dy1 = r[1], r[3]
            # 跳过近似整页的背景填充矩形（会覆盖全页，使空白判断失真）。
            if (dy1 - dy0) >= height * 0.9:
                continue
            if dy1 > practice_bottom + 0.5:
                content_intervals.append((max(dy0, practice_bottom), dy1))

        # 空白下界：优先取页脚顶部，否则取页面底部（去掉常规下边距）。
        region_bottom = footer.y0 if footer else height
        # 在 [practice_bottom, region_bottom] 内，合并内容区间后求最大连续空白高度。
        max_blank = 0.0
        blank_span_desc = ""
        spans = sorted(
            (iv for iv in content_intervals if iv[0] < region_bottom),
            key=lambda iv: iv[0],
        )
        cursor = practice_bottom
        for s0, s1 in spans:
            if s0 > cursor:
                gap = s0 - cursor
                if gap > max_blank:
                    max_blank = gap
                    blank_span_desc = f"{cursor:.1f}→{s0:.1f}pt"
            cursor = max(cursor, s1)
        # 末尾到空白下界之间的空白（实践经历下方内容结束后到页脚/页面底部）。
        if region_bottom - cursor > max_blank:
            max_blank = region_bottom - cursor
            blank_span_desc = f"{cursor:.1f}→{region_bottom:.1f}pt"

        if max_blank > height * 0.5:
            # 条件二：实践经历下方存在超过页面50%的连续真实空白。
            excessive_blank_or_multipage = True
            excessive_reasons.append(
                f"实践经历下方连续空白 {max_blank:.1f}pt（{blank_span_desc}），超过页面高度 {height:.1f}pt 的50%"
            )
    if excessive_blank_or_multipage:
        hits.append(Hit("-5 实践经历下方大空白/多页", -5, True, "；".join(excessive_reasons)))
    else:
        hits.append(Hit("-5 实践经历下方大空白/多页", -5, False, "未检测到实践经历下方超过页面50%的连续真实空白，且简历为单页。"))

    # 扣分项 -3：技能证书新增内容覆盖原有"英语能力""工具能力"或"综合表现"文本。
    missing_originals = [name for name in ["英语能力", "工具能力", "综合表现"] if not first_line(lines, name, min_x=180)]
    if missing_originals:
        hits.append(Hit("-3 技能证书新增内容覆盖原文", -3, True, "原有文本缺失：" + "、".join(missing_originals)))
    else:
        hits.append(Hit("-3 技能证书新增内容覆盖原文", -3, False, "英语能力、工具能力、综合表现均存在，未被覆盖。"))

    # 扣分项 -3：新增内容后“研修规划”模块正文被截断或超出页面底部。
    plan_truncated_or_overflow = False
    plan_reasons: list[str] = []
    if not plan_title or not plan_body:
        plan_truncated_or_overflow = True
        plan_reasons.append("研修规划模块标题或正文缺失，判定为正文被截断")
    else:
        footer = first_line(lines, FOOTER_TEXT)
        footer_top = footer.y0 if footer else height - 25
        last_plan_bottom = max(l.y1 for l in plan_body)
        if last_plan_bottom >= footer_top - 3 or last_plan_bottom > height - 35:
            plan_truncated_or_overflow = True
            plan_reasons.append("研修规划正文超出页面底部或进入页脚区域")
    if plan_truncated_or_overflow:
        hits.append(Hit("-3 研修规划正文截断/越界", -3, True, "；".join(plan_reasons)))
    else:
        hits.append(Hit("-3 研修规划正文截断/越界", -3, False, "研修规划正文未被截断，且未超出页面底部。"))

    # 扣分项 -3：成绩表数字缺失。
    missing_scores = [s for s in ["64", "71", "118", "96", "349"] if s not in compact_text]
    if missing_scores:
        hits.append(Hit("-3 成绩表数字缺失", -3, True, "缺失成绩文本：" + "、".join(missing_scores)))
    else:
        hits.append(Hit("-3 成绩表数字缺失", -3, False, "成绩表中的 64、71、118、96、349 均可检测到。"))

    # 扣分项 -3：第1页底部页脚"求学申请材料 · 2026"被正文内容覆盖或完全消失。
    footer = first_line(lines, FOOTER_TEXT)
    footer_bad = False
    footer_reason = ""
    if not footer:
        # 完全消失。
        footer_bad = True
        footer_reason = "页脚文本未检测到，判定为完全消失"
    else:
        # 被正文内容覆盖：其他行与页脚行在几何上重叠。
        for l in lines:
            if l is footer or normalize_text(l.text) == normalize_text(FOOTER_TEXT):
                continue
            if rects_overlap(line_rect(l), line_rect(footer), pad=0):
                footer_bad = True
                footer_reason = "检测到正文内容与页脚文本位置重叠"
                break
    if footer_bad:
        hits.append(Hit("-3 页脚消失或被覆盖", -3, True, footer_reason))
    else:
        hits.append(Hit("-3 页脚消失或被覆盖", -3, False, "页脚文本存在且未被正文内容覆盖。"))

    total = sum(h.score for h in hits if h.matched)
    return True, hits, total, notes


def _locate_pdf(dir_path: str) -> str | None:
    """在给定目录里定位待评估 PDF：优先精确名称，其次任一 .pdf 文件。"""
    if not os.path.isdir(dir_path):
        return None
    preferred = os.path.join(dir_path, PREFERRED_PDF_NAME)
    if os.path.isfile(preferred):
        return preferred
    for name in sorted(os.listdir(dir_path)):
        if name.lower().endswith(".pdf"):
            candidate = os.path.join(dir_path, name)
            if os.path.isfile(candidate):
                return candidate
    return None


def evaluate(dir_path: str) -> dict[str, object]:
    """按统一约定的入口：接收脚本所在目录路径，返回结构化评估结果。"""
    result: dict[str, object] = {
        "id": SCRIPT_ID,
        "file_name": "",
        "status": "ok",
        "error": None,
        "dim1_pass": False,
        "dim1_reason": "",
        "dim2_items": [],
        "total_score": 0,
        "max_score": 0,
    }

    try:
        pdf_path = _locate_pdf(dir_path)
        if not pdf_path:
            result["status"] = "error"
            result["error"] = f"目录中未找到待评估的 PDF 文件：{dir_path}"
            return result
        result["file_name"] = os.path.basename(pdf_path)

        dimension1_ok, hits, total, _notes = _evaluate_pdf(pdf_path)

        # 维度一
        dim1_hit = next((h for h in hits if h.code == "维度1"), None)
        result["dim1_pass"] = bool(dimension1_ok)
        result["dim1_reason"] = "" if dimension1_ok else (dim1_hit.detail if dim1_hit else "维度一不通过")

        # 维度二：命中与未命中项都返回
        dim2_items: list[dict[str, object]] = []
        max_score = 0
        for h in hits:
            if h.code == "维度1":
                continue
            max_delta = h.score  # +N 或 -N，均为该项的满额影响
            delta = h.score if h.matched else 0
            dim2_items.append({
                "rule": h.code,
                "max_delta": max_delta,
                "delta": delta,
                "hit": bool(h.matched),
                "detail": "",
            })
            if max_delta > 0:
                max_score += max_delta

        result["dim2_items"] = dim2_items
        result["total_score"] = total if dimension1_ok else 0
        result["max_score"] = max_score
        return result
    except Exception as exc:  # 兜底：脚本自身异常 -> status=error
        result["status"] = "error"
        result["error"] = f"{type(exc).__name__}: {exc}"
        result["dim1_pass"] = False
        result["dim2_items"] = []
        result["total_score"] = 0
        return result


if __name__ == "__main__":
    # 本地调试：默认取脚本所在目录；也支持通过 argv[1] 覆盖。
    _dir = sys.argv[1] if len(sys.argv) > 1 else os.path.dirname(os.path.abspath(__file__))
    print(json.dumps(evaluate(_dir), ensure_ascii=False, indent=2))
