#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""自动评估“七年级地理导学案_添加目录版.pdf”的目录页完成度。

依赖：pdfplumber（经 pdf_backend 适配层）。本脚本不调用人工判断；对于 PDF 中
字体替换、坐标抽取误差等情况，使用可解释的容差来贴近评分意图。
"""
from __future__ import annotations

import json
import math
import os
import re
import sys
import unicodedata
from dataclasses import dataclass
from typing import Any, Iterable

try:
    try:
        import pdf_backend
    except ImportError:
        from verifiers import pdf_backend
except ImportError as exc:  # pragma: no cover
    raise ImportError("缺少依赖 pdfplumber：请先安装 pip install pdfplumber") from exc


class _Pt:
    """fitz.Point 兼容的轻量点（drawing items 用）。"""

    __slots__ = ("x", "y")

    def __init__(self, x: float, y: float):
        self.x = x
        self.y = y


class _Page:
    """页面句柄：rect / 文本 span / 矢量绘图（替代 fitz.Page）。"""

    def __init__(self, doc: "pdf_backend.PdfDocument", index: int):
        self._doc = doc
        self.index = index
        w, h = doc.page_size(index)
        self.rect = pdf_backend.PdfRect(0.0, 0.0, w, h)
        self._drawings: list[dict] | None = None
        self._spans: "list[Span] | None" = None
        self._text: str | None = None

    def page_text(self) -> str:
        if self._text is None:
            self._text = self._doc.page_text(self.index)
        return self._text

    def raw_spans(self):
        # gap_chars=1.2 使 span 切分粒度与历史 rawdict 对齐（已逐页对照验证）
        return self._doc.extract_raw_spans(self.index, gap_chars=1.2)

    def images(self):
        return self._doc.extract_images(self.index)

    def get_drawings(self) -> list[dict]:
        """返回与历史 get_drawings 结构兼容的 drawing 字典列表。"""
        if self._drawings is not None:
            return self._drawings
        out: list[dict] = []
        for p in self._doc.extract_paths(self.index):
            items: list[tuple] = []
            for it in p.items:
                if it[0] == "re":
                    items.append(("re", it[1]))
                elif it[0] == "l":
                    items.append(("l", _Pt(*it[1]), _Pt(*it[2])))
            if p.fill is not None and p.stroke is not None:
                dtype = "fs"
            elif p.fill is not None:
                dtype = "f"
            else:
                dtype = "s"
            out.append({
                "rect": (p.rect.x0, p.rect.y0, p.rect.x1, p.rect.y1),
                "fill": p.fill,
                "color": p.stroke,
                "width": p.line_width,
                "type": dtype,
                "items": items,
            })
        self._drawings = out
        return out


class _Doc:
    """文档句柄：页数 / 页访问 / 迭代 / 关闭（替代 fitz.Document）。"""

    def __init__(self, path: str):
        self._doc = pdf_backend.open_pdf(path)
        self.page_count = self._doc.page_count
        self._pages = [_Page(self._doc, i) for i in range(self.page_count)]

    def __len__(self) -> int:
        return self.page_count

    def __getitem__(self, index: int) -> _Page:
        return self._pages[index]

    def __iter__(self):
        return iter(self._pages)

    def close(self) -> None:
        self._doc.close()

PT_PER_CM = 72 / 2.54
SCRIPT_ID = "017"
PREFERRED_PDF_NAME = "七年级地理导学案_添加目录版.pdf"
DIM2_MAX_SCORE = 95  # 维度二所有正向评分项 max_delta 之和


@dataclass
class Rect:
    x0: float
    y0: float
    x1: float
    y1: float

    @property
    def width(self) -> float:
        return max(0.0, self.x1 - self.x0)

    @property
    def height(self) -> float:
        return max(0.0, self.y1 - self.y0)

    @property
    def area(self) -> float:
        return self.width * self.height

    @property
    def cx(self) -> float:
        return (self.x0 + self.x1) / 2

    @property
    def cy(self) -> float:
        return (self.y0 + self.y1) / 2

    def intersect(self, other: "Rect") -> "Rect | None":
        x0, y0 = max(self.x0, other.x0), max(self.y0, other.y0)
        x1, y1 = min(self.x1, other.x1), min(self.y1, other.y1)
        if x1 <= x0 or y1 <= y0:
            return None
        return Rect(x0, y0, x1, y1)

    def to_tuple(self) -> tuple[float, float, float, float]:
        return (self.x0, self.y0, self.x1, self.y1)


def to_rect(obj: Any) -> Rect:
    return Rect(float(obj[0]), float(obj[1]), float(obj[2]), float(obj[3]))


@dataclass
class Span:
    text: str
    norm: str
    compact: str
    rect: Rect
    size: float
    font: str
    color: int
    flags: int
    origin: tuple[float, float] | None = None
    chars: list[dict[str, Any]] | None = None

    @property
    def baseline_y(self) -> float:
        if self.origin:
            return float(self.origin[1])
        return self.rect.y1


@dataclass
class Row:
    key: str
    label: str
    page_no: str
    kind: str
    parts: list[Span]
    number_span: Span | None

    @property
    def rect(self) -> Rect:
        return union_bbox([s.rect for s in self.parts])

    @property
    def y(self) -> float:
        return median([s.rect.cy for s in self.parts])

    @property
    def text_right(self) -> float:
        return max(s.rect.x1 for s in self.parts)

    @property
    def text_left(self) -> float:
        return min(s.rect.x0 for s in self.parts)


@dataclass
class ScoreResult:
    score: int
    desc: str
    matched: bool
    evidence: str


@dataclass
class GateResult:
    name: str
    passed: bool
    evidence: str


def normalize_text(text: str) -> str:
    text = unicodedata.normalize("NFKC", text or "")
    # 常见 CJK 兼容字形归一化后通常已经转换；这里补充少数 PDF 提取差异。
    text = text.replace("落", "落").replace("量", "量")
    text = text.replace("　", " ").replace("\xa0", " ")
    text = re.sub(r"\s+", " ", text).strip()
    return text


def compact(text: str) -> str:
    return re.sub(r"\s+", "", normalize_text(text))


def nearly(value: float, low: float, high: float, tol: float = 0.0) -> bool:
    return low - tol <= value <= high + tol


def median(values: Iterable[float]) -> float:
    vals = sorted(values)
    if not vals:
        return 0.0
    mid = len(vals) // 2
    if len(vals) % 2:
        return vals[mid]
    return (vals[mid - 1] + vals[mid]) / 2


def union_bbox(rects: list[Rect]) -> Rect:
    return Rect(
        min(r.x0 for r in rects),
        min(r.y0 for r in rects),
        max(r.x1 for r in rects),
        max(r.y1 for r in rects),
    )


def union_area(rects: list[Rect]) -> float:
    """计算矩形并集面积，矩形数量较小，使用 x 方向扫描即可。"""
    rects = [r for r in rects if r.area > 0]
    if not rects:
        return 0.0
    xs = sorted({x for r in rects for x in (r.x0, r.x1)})
    total = 0.0
    for a, b in zip(xs, xs[1:]):
        if b <= a:
            continue
        intervals: list[tuple[float, float]] = []
        for r in rects:
            if r.x0 < b and r.x1 > a:
                intervals.append((r.y0, r.y1))
        if not intervals:
            continue
        intervals.sort()
        merged: list[list[float]] = []
        for y0, y1 in intervals:
            if not merged or y0 > merged[-1][1]:
                merged.append([y0, y1])
            else:
                merged[-1][1] = max(merged[-1][1], y1)
        height = sum(y1 - y0 for y0, y1 in merged)
        total += (b - a) * height
    return total


def rgb_from_int(color: int) -> tuple[int, int, int]:
    return ((int(color) >> 16) & 255, (int(color) >> 8) & 255, int(color) & 255)


def color_tuple_to_rgb255(color: tuple[float, float, float] | None) -> tuple[int, int, int] | None:
    if color is None:
        return None
    return tuple(max(0, min(255, int(round(v * 255)))) for v in color[:3])  # type: ignore[return-value]


def is_black_int(color: int, max_channel: int = 60) -> bool:
    r, g, b = rgb_from_int(color)
    return r <= max_channel and g <= max_channel and b <= max_channel


def is_green_int(color: int) -> bool:
    r, g, b = rgb_from_int(color)
    return g >= 90 and g > r * 1.6 and g > b * 1.15


def is_white_int(color: int) -> bool:
    r, g, b = rgb_from_int(color)
    return r >= 245 and g >= 245 and b >= 245


def is_green_tuple(color: tuple[float, float, float] | None) -> bool:
    rgb = color_tuple_to_rgb255(color)
    if not rgb:
        return False
    r, g, b = rgb
    return g >= 90 and g > r * 1.4 and g > b * 1.05


def is_green_decoration_tuple(color: tuple[float, float, float] | None) -> bool:
    rgb = color_tuple_to_rgb255(color)
    if not rgb:
        return False
    r, g, b = rgb
    # 装饰斜条可能是浅绿色/渐变绿色，RGB 中 G 仍应占优。
    return g >= 120 and g >= r * 1.15 and g >= b * 1.05


def is_black_tuple(color: tuple[float, float, float] | None) -> bool:
    rgb = color_tuple_to_rgb255(color)
    if not rgb:
        return False
    return max(rgb) <= 70


def is_white_tuple(color: tuple[float, float, float] | None) -> bool:
    rgb = color_tuple_to_rgb255(color)
    if not rgb:
        return False
    return min(rgb) >= 245


def luminance_tuple(color: tuple[float, float, float] | None) -> float:
    rgb = color_tuple_to_rgb255(color)
    if not rgb:
        return 0.0
    r, g, b = rgb
    return 0.2126 * r + 0.7152 * g + 0.0722 * b


def is_bold(span: Span) -> bool:
    return bool(span.flags & 16) or "bold" in span.font.lower() or "black" in span.font.lower()


def normalized_font_name(span: Span) -> str:
    """归一化字体名：剥离 PDF 子集前缀（形如 "ABCDEF+FontName"），再去除大小写/分隔符。

    办公软件（Word/WPS）导出 PDF 时通常给嵌入字体加 6 位随机子集前缀，例如
    `BCDEEE+SimHei`。剥离后按包含关键字判断中文字族。
    """
    font = span.font
    if "+" in font:
        font = font.split("+", 1)[1]
    return font.lower().replace("-", "").replace("_", "").replace(" ", "")


def is_cjk_sans_or_yahei(span: Span) -> bool:
    # 严格按评分细则：只认可黑体或微软雅黑，不把 Noto/SourceHan 等替代字体自动等同。
    font = normalized_font_name(span)
    markers = ["simhei", "stheiti", "heiti", "microsoftyahei", "msyh", "yahei", "黑体", "微软雅黑"]
    return any(m in font for m in markers)


def is_song_or_yahei_or_equivalent(span: Span) -> bool:
    # 严格按评分细则：只认可宋体或微软雅黑，不把 Noto/SourceHan 等替代字体自动等同。
    font = normalized_font_name(span)
    markers = ["simsun", "simsong", "songti", "song", "microsoftyahei", "msyh", "yahei", "宋体", "微软雅黑"]
    return any(m in font for m in markers)


def extract_spans(page: _Page) -> list[Span]:
    if page._spans is not None:
        return page._spans
    spans: list[Span] = []
    for sp in page.raw_spans():
        text = sp.text
        if not normalize_text(text):
            continue
        font = sp.font
        # 与历史 flags 语义对齐：bit4（加粗）由字体名合成（本迁移已验证
        # 样本中 flags&16 与字体名含 Bold/Black 完全一致），其余位不使用。
        flags = 16 if ("bold" in font.lower() or "black" in font.lower()) else 0
        spans.append(
            Span(
                text=text,
                norm=normalize_text(text),
                compact=compact(text),
                rect=Rect(sp.bbox.x0, sp.bbox.y0, sp.bbox.x1, sp.bbox.y1),
                size=float(sp.size),
                font=font,
                color=int(sp.color),
                flags=flags,
                origin=None,
                chars=[
                    {"c": ch.c, "bbox": (ch.bbox.x0, ch.bbox.y0, ch.bbox.x1, ch.bbox.y1)}
                    for ch in sp.chars
                ],
            )
        )
    page._spans = spans
    return spans


def find_span(spans: list[Span], phrase: str, *, y_range: tuple[float, float] | None = None) -> Span | None:
    target = compact(phrase)
    candidates = []
    for sp in spans:
        if target and target in sp.compact:
            if y_range and not (y_range[0] <= sp.rect.cy <= y_range[1]):
                continue
            candidates.append(sp)
    if not candidates:
        return None
    return sorted(candidates, key=lambda s: (s.rect.y0, s.rect.x0))[0]


def find_exact_number(spans: list[Span], number: str, y: float, tolerance: float = 8) -> Span | None:
    nums = [s for s in spans if s.compact == str(number) and abs(s.rect.cy - y) <= tolerance]
    if not nums:
        return None
    return max(nums, key=lambda s: s.rect.x0)


def drawing_rect(d: dict[str, Any]) -> Rect | None:
    rect = d.get("rect")
    if not rect:
        return None
    return to_rect(rect)


def drawing_items_have_polygon(d: dict[str, Any]) -> bool:
    items = d.get("items", [])
    line_count = sum(1 for item in items if item and item[0] == "l")
    return line_count >= 2


def line_segments_from_drawing(d: dict[str, Any]) -> list[tuple[float, float, float, float]]:
    segs: list[tuple[float, float, float, float]] = []
    last: tuple[float, float] | None = None
    for item in d.get("items", []):
        if not item:
            continue
        op = item[0]
        if op == "l" and len(item) >= 3:
            p0, p1 = item[1], item[2]
            segs.append((float(p0.x), float(p0.y), float(p1.x), float(p1.y)))
            last = (float(p1.x), float(p1.y))
        elif op == "re" and len(item) >= 2:
            r = item[1]
            segs.extend(
                [
                    (r.x0, r.y0, r.x1, r.y0),
                    (r.x1, r.y0, r.x1, r.y1),
                    (r.x1, r.y1, r.x0, r.y1),
                    (r.x0, r.y1, r.x0, r.y0),
                ]
            )
        elif op == "m" and len(item) >= 2:
            p = item[1]
            last = (float(p.x), float(p.y))
        elif op == "c" and last and len(item) >= 4:
            p3 = item[3]
            segs.append((last[0], last[1], float(p3.x), float(p3.y)))
            last = (float(p3.x), float(p3.y))
    return segs


def small_black_dots(page: _Page) -> list[Rect]:
    """收集页面上所有“黑色小点”候选（用作点状引导线的点）。

    面向办公软件（Word/WPS）导出的 PDF，引导点有两种来源：
    A) 段落使用“制表符-前导符-点”（Tab Leader）：PDF 中表现为一串
       “英文句点 . / 中文句号 。 / 中点 ・”字符（文本流）。
    B) 手动插入的形状小圆点：PDF 中表现为黑色填充的极小 drawing（矢量图形）。
    这里两种都作为候选点收集。
    """
    dots: list[Rect] = []

    # A. 文本流中的引导点字符
    for sp in extract_spans(page):
        if not is_black_int(sp.color):
            continue
        # 一个 span 里可能一次性抽出多个句点；用逐字符 bbox 展开
        for ch in sp.chars or []:
            c = ch.get("c", "")
            if c in {".", "．", "·", "・", "•", "。"}:
                r = to_rect(ch["bbox"])
                dots.append(r)

    # B. 矢量绘制的黑色小点
    for d in page.get_drawings():
        r = drawing_rect(d)
        if not r:
            continue
        fill = d.get("fill")
        if is_black_tuple(fill) and 0.4 <= r.width <= 4.0 and 0.4 <= r.height <= 4.0:
            dots.append(r)
    return dots


def has_dotted_leader(dots: list[Rect], y: float, x_left: float, x_right: float) -> tuple[bool, str]:
    """判断某行是否存在“黑色点状引导线”。

    参数:
        dots    页面所有候选黑色小点/句点字符
        y       该目录条目行的中心 y
        x_left  条目文字最右侧 x（引导线应从此右侧开始）
        x_right 页码数字最左侧 x（引导线应止于此左侧）

    返回 True 需同时满足：
    - 点位于 (x_left, x_right) 严格区间内 —— 不穿过条目文字或页码
    - 点的 y 与行 y 近似（同一水平行）
    - 至少 2 个点，且横向排列（x 单调递增，非纵向堆叠）
    """
    row_dots = [
        d for d in dots
        if abs(d.cy - y) <= 4 and d.x0 > x_left and d.x1 < x_right
    ]
    if len(row_dots) < 2:
        return False, f"y≈{y:.1f} 的点数不足（{len(row_dots)} 个，需≥2）"

    xs = sorted(d.cx for d in row_dots)
    span_x = xs[-1] - xs[0]
    ys = [d.cy for d in row_dots]
    span_y = max(ys) - min(ys)
    horizontal_ok = span_x >= span_y * 2  # 横向排列（横跨 >> 纵跨）

    return horizontal_ok, (
        f"点数={len(row_dots)}，横向跨度={span_x:.1f}pt，纵向跨度={span_y:.1f}pt，"
        f"横向排列={'✓' if horizontal_ok else '✗'}"
    )


def open_pdf_gate(path: str) -> tuple["_Doc | None", list[GateResult]]:
    gates: list[GateResult] = []
    if not path.lower().endswith(".pdf"):
        gates.append(GateResult("交付文件为 PDF 格式", False, f"扩展名不是 .pdf：{path}"))
        return None, gates
    gates.append(GateResult("交付文件为 PDF 格式", True, os.path.basename(path)))
    try:
        doc = _Doc(path)
        # 强制读取第一页，以便发现损坏或加密打不开的 PDF。
        if len(doc) > 0:
            _ = doc[0].page_text()
        gates.append(GateResult("文件可正常打开", True, f"pdfplumber 成功打开，页数 {len(doc)}"))
        return doc, gates
    except Exception as exc:
        gates.append(GateResult("文件可正常打开", False, f"打开失败：{exc}"))
        return None, gates


def page_is_blank(page: _Page) -> bool:
    text = compact(page.page_text())
    if len(text) >= 8:
        return False
    if page.images():
        return False
    page_area = page.rect.width * page.rect.height
    draw_area = 0.0
    for d in page.get_drawings():
        r = drawing_rect(d)
        if r:
            draw_area += min(r.area, page_area)
    return draw_area < page_area * 0.01


def text_overlap_area(page: _Page) -> float:
    spans = [s for s in extract_spans(page) if s.rect.area > 8]
    inters: list[Rect] = []
    for i, a in enumerate(spans):
        for b in spans[i + 1 :]:
            # 相邻字符/同一文本行偶有微小 bbox 接触，不计入重叠。
            inter = a.rect.intersect(b.rect)
            if not inter or inter.area < 4:
                continue
            if abs(a.rect.cy - b.rect.cy) < max(a.size, b.size) * 0.35:
                continue
            inters.append(inter)
    return union_area(inters)


def dimension1_gates(path: str) -> tuple["_Doc | None", list[GateResult]]:
    doc, gates = open_pdf_gate(path)
    if doc is None:
        return None, gates

    pages_ok = len(doc) >= 17
    gates.append(GateResult("交付 PDF 页数不少于 17 页", pages_ok, f"实际页数：{len(doc)}"))

    max_blank_run = 0
    cur = 0
    blank_pages: list[int] = []
    overlap_bad: list[str] = []
    for idx, page in enumerate(doc, start=1):
        if page_is_blank(page):
            cur += 1
            blank_pages.append(idx)
            max_blank_run = max(max_blank_run, cur)
        else:
            cur = 0
        area = text_overlap_area(page)
        page_area = page.rect.width * page.rect.height
        if area > page_area / 3:
            overlap_bad.append(f"第{idx}页重叠面积约 {area:.0f}/{page_area:.0f}pt²")

    gates.append(
        GateResult(
            "交付 PDF 无连续 2 页以上空白页",
            max_blank_run < 2,
            f"空白页：{blank_pages or '无'}，最大连续 {max_blank_run} 页",
        )
    )
    gates.append(
        GateResult(
            "交付 PDF 无超过 1/3 页面面积文字重叠",
            not overlap_bad,
            "; ".join(overlap_bad) if overlap_bad else "未发现超过 1/3 页面面积的文字重叠",
        )
    )
    return doc, gates


def build_rows(spans: list[Span]) -> dict[str, Row]:
    specs: list[tuple[str, str, str, str, list[str]]] = [
        ("learning", "学习结构说明", "2", "learning", ["学习结构说明"]),
        ("unit1", "第一单元 地球与地图", "3", "unit", ["第一单元", "地球与地图"]),
        ("task1", "任务 1 认识地球与经纬网", "3", "task", ["任务 1", "认识地球与经纬网"]),
        ("task2", "任务 2 地图三要素与方向", "4", "task", ["任务 2", "地图三要素与方向"]),
        ("task3", "任务 3 等高线地形图", "5", "task", ["任务 3", "等高线地形图"]),
        ("unit2", "第二单元 陆地与海洋", "7", "unit", ["第二单元", "陆地与海洋"]),
        ("task4", "任务 4 七大洲与四大洋", "7", "task", ["任务 4", "七大洲与四大洋"]),
        ("task5", "任务 5 海陆变迁与板块运动", "8", "task", ["任务 5", "海陆变迁与板块运动"]),
        ("unit3", "第三单元 天气与气候", "9", "unit", ["第三单元", "天气与气候"]),
        ("task6", "任务 6 天气预报与空气质量", "9", "task", ["任务 6", "天气预报与空气质量"]),
        ("task7", "任务 7 气温曲线与降水柱状图", "10", "task", ["任务 7", "气温曲线与降水柱状图"]),
        ("task8", "任务 8 世界气候类型与生活", "11", "task", ["任务 8", "世界气候类型与生活"]),
        ("unit4", "第四单元 居民与聚落", "12", "unit", ["第四单元", "居民与聚落"]),
        ("task9", "任务 9 人口分布与聚落形成", "12", "task", ["任务 9", "人口分布与聚落形成"]),
        ("task10", "任务 10 聚落发展与可持续生活", "13", "task", ["任务 10", "聚落发展与可持续生活"]),
        ("mid_project", "期中项目 校园导览图制作任务", "14", "project", ["期中项目", "校园导览图制作任务"]),
        ("final_project", "期末项目 家乡聚落观察报告", "15", "project", ["期末项目", "家乡聚落观察报告"]),
        ("portfolio", "全册学习档案袋 阶段复盘与成长记录", "16", "project", ["全册学习档案袋", "阶段复盘与成长记录"]),
    ]
    rows: dict[str, Row] = {}
    for key, label, page_no, kind, phrases in specs:
        parts: list[Span] = []
        for phrase in phrases:
            sp = find_span(spans, phrase)
            if sp:
                parts.append(sp)
        if len(parts) != len(phrases):
            continue
        y = median([p.rect.cy for p in parts])
        num = find_exact_number(spans, page_no, y)
        rows[key] = Row(key, label, page_no, kind, parts, num)
    return rows


def get_title_span(spans: list[Span]) -> Span | None:
    return find_span(spans, "目录", y_range=(0, 100))


def title_char_gap_pt(title: Span) -> float | None:
    """返回“目”和“录”两字之间的净水平间距（磅）。

    面向办公软件（Word/WPS）：两字符经过“字符间距”设置后，PDF 抽取时会体现为
    上一字 bbox 右边到下一字 bbox 左边的空白。若某个字被拆到不同 span 里，也从
    title.chars 中按字符定位取值。
    """
    chars = title.chars or []
    mu = [c for c in chars if normalize_text(c.get("c", "")) == "目"]
    lu = [c for c in chars if normalize_text(c.get("c", "")) == "录"]
    if not mu or not lu:
        return None
    m = to_rect(mu[0]["bbox"])
    l = to_rect(lu[0]["bbox"])
    return max(0.0, l.x0 - m.x1)


def check_title_position(title: Span | None) -> tuple[bool, str]:
    """“目录”两字距离页面上边界55-65磅，距离页面左边界85-95磅，
    且“目”和“录”之间空0.3-1字符。

    严格对应细则三点，面向办公软件（Word/WPS）导出的 PDF：
    1) 距页面上边界 55-65 磅：以“目录”文字 bbox 顶部 (y0) 到页面顶部的距离衡量，
       对应 Word 中“页边距-上”+行前空白后的实际排版起始位置。
    2) 距页面左边界 85-95 磅：以“目录”文字 bbox 左侧 (x0) 到页面左侧的距离衡量，
       对应 Word 中“页边距-左”+段落缩进后的实际排版位置。
    3) “目”和“录”之间空 0.3-1 字符：以两字符之间的净水平间距 ÷ 一字符宽度衡量。
       办公软件中“1 字符”对于 CJK 汉字 = 1 个中文字宽 = 字号（em 宽）；此处以字号
       作为“一字符”的参考宽度（36pt 小初 → 1 字符 = 36pt）。
    """
    if not title:
        return False, "未找到“目录”标题"

    top_distance = title.rect.y0
    left_distance = title.rect.x0
    gap_pt = title_char_gap_pt(title)
    # 办公软件中“1 字符”对 CJK = 一个汉字宽 = 字号（em）。
    char_width = title.size if title.size > 0 else 36.0
    gap_ratio = (gap_pt / char_width) if gap_pt is not None else None

    top_ok = 55 <= top_distance <= 65
    left_ok = 85 <= left_distance <= 95
    gap_ok = gap_ratio is not None and 0.3 <= gap_ratio <= 1.0

    ok = top_ok and left_ok and gap_ok
    gap_text = "无法计算" if gap_ratio is None else f"{gap_ratio:.2f}字符({gap_pt:.1f}pt)"
    return ok, (
        f"距上边界={top_distance:.1f}pt({'✓55-65' if top_ok else '✗超出55-65'})，"
        f"距左边界={left_distance:.1f}pt({'✓85-95' if left_ok else '✗超出85-95'})，"
        f"目/录间距={gap_text}({'✓0.3-1字符' if gap_ok else '✗超出0.3-1字符'})"
    )


def is_heiti_or_yahei_font(span: Span) -> bool:
    """办公软件（Word/WPS）中“黑体”或“微软雅黑”的字体名匹配。

    PDF 嵌入字体常带 6 位随机子集前缀（如 BCDEEE+SimHei），需先剥离；
    Windows Office：黑体→SimHei，微软雅黑→Microsoft YaHei / MSYH。
    macOS Office：黑体→STHeiti / Heiti SC。
    """
    font = span.font
    # 剥离 PDF 子集前缀（形如 "ABCDEF+FontName"）
    if "+" in font:
        font = font.split("+", 1)[1]
    font = font.lower().replace("-", "").replace("_", "").replace(" ", "")
    markers = ["simhei", "stheiti", "heiti", "microsoftyahei", "msyh", "yahei", "黑体", "微软雅黑"]
    return any(marker in font for marker in markers)


def check_title_style(title: Span | None) -> tuple[bool, str]:
    """“目录”两字字体字号为黑体或微软雅黑、小初、加粗，颜色为绿色。

    严格对应细则四点，且面向办公软件（Word/WPS）导出的 PDF：
    1) 字体：办公软件中“黑体”对应 SimHei（Windows）或 STHeiti（macOS）；
       “微软雅黑”对应 Microsoft YaHei / MSYH / YaHei。字体名可能带子集前缀
       （如 BCDEEE+SimHei），去除前缀后按包含匹配。
    2) 字号：小初 = 36 磅。允许 ±0.5pt 抽取误差。
    3) 加粗：办公软件按“加粗”按钮后，PDF span flags 第 4 位置位；
       或字体名带 Bold/Black 后缀（如 MSYH-Bold）。
    4) 颜色：绿色（不限定具体 RGB，只要 G 通道占优）。
    """
    if not title:
        return False, "未找到“目录”标题"

    font_ok = is_heiti_or_yahei_font(title)
    size_ok = nearly(title.size, 36, 36, 0.5)
    bold_ok = is_bold(title)
    color_ok = is_green_int(title.color)

    ok = font_ok and size_ok and bold_ok and color_ok
    rgb = rgb_from_int(title.color)
    return ok, (
        f"字体={title.font}({'✓黑体/微软雅黑' if font_ok else '✗非黑体/微软雅黑'})，"
        f"字号={title.size:.1f}pt({'✓小初' if size_ok else '✗非小初36pt'})，"
        f"加粗={'✓' if bold_ok else '✗'}，"
        f"颜色RGB={rgb}({'✓绿色' if color_ok else '✗非绿色'})"
    )


def check_title_decorations(page: _Page, title: Span | None) -> tuple[bool, str]:
    """“目录”两字左侧有一个绿色的实心矩形块和两条绿色平行的斜条装饰，
    斜条从左上向右下倾斜，与文字顶部齐平。

    严格对应细则五点，面向办公软件（Word/WPS）导出的 PDF：
    1) 位置：装饰位于“目录”两字左侧（元素的右边缘 x1 ≤ 标题 x0），
       且与标题在同一垂直范围（drawing 与标题 bbox 在 y 方向有重叠），
       以排除页面其它位置（例如页脚项目箭头）的绿色形状干扰。
    2) 一个绿色实心矩形块：Word/WPS 中“插入-矩形”并填充绿色，导出 PDF 后表现为
       type='f' 且 items 含 're' 操作的绿色 drawing，取至少 1 个。
    3) 两条绿色平行斜条：Word/WPS 中“插入-直线”并旋转，导出后表现为含 'l' 线段
       的绿色 drawing（或线宽较粗的绿色 stroke），取至少 2 条。
    4) 从左上向右下倾斜：在 PDF 坐标系（y 向下增大）中，视觉“左上→右下”对应
       dx > 0 且 dy > 0 —— x 增大时 y 也增大。
    5) 与文字顶部齐平：两条斜条的顶部（min y0）与标题文字顶部（title.rect.y0）
       近似对齐（容差 ≈ 半个字号，用于抵消办公软件字形上升部导致的 bbox 偏差）。
    """
    if not title:
        return False, "未找到“目录”标题"

    rects: list[Rect] = []
    slants: list[tuple[Rect, float]] = []  # (rect, signed_angle_deg)

    for d in page.get_drawings():
        r = drawing_rect(d)
        if not r:
            continue
        # 颜色：填充或描边为绿色即视为绿色装饰
        fill_green = is_green_decoration_tuple(d.get("fill"))
        stroke_green = is_green_decoration_tuple(d.get("color"))
        if not (fill_green or stroke_green):
            continue

        # 点1：位于“目录”两字左侧（装饰左边缘位于标题左边缘之前），
        # 且与标题在同一垂直范围（drawing 与标题 bbox 在 y 方向有重叠），
        # 以排除页面其它位置（例如页脚项目箭头）的绿色形状干扰。
        # 使用 x0 而非 x1 作为“位于左侧”的判据：斜向平行四边形的尾部可能
        # 略探入标题字的 x 范围，但只要其起点位于标题左边缘之前即视为“左侧装饰”。
        if r.x0 >= title.rect.x0:
            continue
        if r.y1 < title.rect.y0 or r.y0 > title.rect.y1:
            continue

        items = d.get("items", [])
        has_rect_op = any(item and item[0] == "re" for item in items)
        has_line_op = any(item and item[0] == "l" for item in items)

        # 点2：绿色实心矩形块 —— 填充型 + 含 're'
        if d.get("type") == "f" and has_rect_op and fill_green and not has_line_op:
            rects.append(r)
            continue

        # 点3+4：绿色斜条 —— 遍历线段，取斜向段并记录“有符号角度”
        for x0, y0, x1, y1 in line_segments_from_drawing(d):
            dx, dy = x1 - x0, y1 - y0
            if abs(dx) < 6 or abs(dy) < 6:
                continue
            # 方向规范化：让 dx 恒为正（视觉从左向右扫描）
            if dx < 0:
                dx, dy = -dx, -dy
            # 视觉“左上→右下” 在 PDF 坐标（y 向下增大）中要求 dy > 0
            angle = math.degrees(math.atan2(dy, dx))  # 正值 = 左上→右下
            slants.append((r, angle))
            break  # 每个 drawing 只贡献一条斜条

    # 判定
    rects_ok = len(rects) >= 1
    count_ok = len(slants) >= 2
    direction_ok = count_ok and all(35 <= a <= 75 for _, a in slants)  # 明显向右下
    parallel_ok = False
    if count_ok:
        angs = [a for _, a in slants]
        parallel_ok = max(angs) - min(angs) <= 8

    top_aligned = False
    if slants:
        top_aligned = abs(min(r.y0 for r, _ in slants) - title.rect.y0) <= title.size * 0.5

    ok = rects_ok and count_ok and direction_ok and parallel_ok and top_aligned
    return ok, (
        f"绿色实心矩形块={len(rects)}个({'✓≥1' if rects_ok else '✗<1'})，"
        f"斜条={len(slants)}条({'✓≥2' if count_ok else '✗<2'})，"
        f"左上→右下倾斜={'✓' if direction_ok else '✗'}，"
        f"平行={'✓' if parallel_ok else '✗'}，"
        f"与文字顶部齐平={'✓' if top_aligned else '✗'}"
    )


def title_line_candidates(page: _Page, title: Span | None) -> list[tuple[Rect, tuple[float, float, float] | None]]:
    """收集“目录”下方所有可能的绿色横线片段（含填充块与描边线段）。

    面向办公软件（Word/WPS）：
    - “插入-形状-矩形”做渐变填充：导出为多个绿色 fill 小矩形拼接（PDF 抽取层
      无法直接读取 shading，通常拆为色阶片段）；
    - “插入-形状-直线”做渐变描边：导出为绿色 stroke 线段；
    这里都作为候选，交由上层判定粗细/位置/渐变方向。
    """
    if not title:
        return []
    out: list[tuple[Rect, tuple[float, float, float] | None]] = []
    for d in page.get_drawings():
        r = drawing_rect(d)
        if not r:
            continue
        fill = d.get("fill")
        color = d.get("color")
        # 只要填充或描边为绿色即算候选
        green = None
        if is_green_tuple(fill):
            green = fill
        elif is_green_tuple(color):
            green = color
        else:
            continue
        # 位置粗筛：在“目录”正下方（不要求紧贴，容纳办公软件段前段后间距）
        below_title = r.y0 >= title.rect.y1 - 3
        # 形态粗筛：横向（明显宽>高），且高度落在“线条粗细”合理量级
        horizontal = r.width >= r.height * 4 and r.width >= 30
        thin = 0.5 <= r.height <= 12
        if below_title and horizontal and thin:
            out.append((r, green))
    return out


def _line_gradient_left_darker(cands: list[tuple[Rect, tuple[float, float, float] | None]]) -> bool:
    """判断“左深右浅”：按候选片段 x 中心排序，比较左半均值与右半均值亮度。"""
    valid = [(r, c) for r, c in cands if c is not None]
    if len(valid) < 2:
        return False
    valid.sort(key=lambda rc: rc[0].cx)
    mid = len(valid) // 2
    left_lum = sum(luminance_tuple(c) for _, c in valid[:mid]) / max(1, mid)
    right_lum = sum(luminance_tuple(c) for _, c in valid[-mid:]) / max(1, mid)
    # 左深右浅 ⇔ 左侧亮度低于右侧
    return right_lum - left_lum >= 5


def check_title_gradient_line(page: _Page, title: Span | None) -> tuple[bool, str]:
    """“目录”两字下方有一条粗0.1-0.2cm的横向绿色渐变线条，左深右浅且与目字
    左边缘对齐，从目字右侧延伸至页面右侧约三分之二处。

    严格对应细则七点，面向办公软件（Word/WPS）导出的 PDF：
    1) 位于“目录”两字下方：线条整体 y0 ≥ 标题 y1。
    2) 粗 0.1-0.2 cm：线条 bbox 高度 ∈ [0.1, 0.2] cm（≈2.83-5.67pt）。
    3) 横向：bbox 宽 >> 高（横向排列）。
    4) 绿色：颜色为绿色。
    5) 渐变（左深右浅）：横向多个绿色片段呈现颜色渐变，且左侧亮度低于右侧。
    6) 与“目”字左边缘对齐：线条 x0 ≈ 标题 x0。
    7) 延伸至页面右侧约三分之二处：线条 x1 ≈ 页面宽度 × 2/3。
    """
    if not title:
        return False, "未找到“目录”标题"
    cands = title_line_candidates(page, title)
    if not cands:
        return False, "“目录”下方未找到绿色横线"

    line_bbox = union_bbox([r for r, _ in cands])
    height_cm = line_bbox.height / PT_PER_CM
    page_w = page.rect.width

    below_ok = line_bbox.y0 >= title.rect.y1 - 3
    thickness_ok = 0.1 <= height_cm <= 0.2
    horizontal_ok = line_bbox.width >= line_bbox.height * 4
    left_align_ok = abs(line_bbox.x0 - title.rect.x0) <= 8
    # “约三分之二处”：允许 2/3 处 ±5% 页宽的偏差
    target_right = page_w * (2 / 3)
    right_ok = abs(line_bbox.x1 - target_right) <= page_w * 0.05
    gradient_ok = _line_gradient_left_darker(cands)
    # 颜色已在候选阶段筛过，这里再兜底一次
    green_ok = any(is_green_tuple(c) for _, c in cands)

    ok = below_ok and thickness_ok and horizontal_ok and left_align_ok and right_ok and gradient_ok and green_ok
    return ok, (
        f"位于标题下方={'✓' if below_ok else '✗'}，"
        f"粗细={height_cm:.2f}cm({'✓0.1-0.2' if thickness_ok else '✗超出0.1-0.2'})，"
        f"横向={'✓' if horizontal_ok else '✗'}，"
        f"绿色={'✓' if green_ok else '✗'}，"
        f"左深右浅渐变={'✓' if gradient_ok else '✗'}，"
        f"左端={line_bbox.x0:.1f}pt/目字左={title.rect.x0:.1f}pt({'✓对齐' if left_align_ok else '✗未对齐'})，"
        f"右端={line_bbox.x1:.1f}pt/页面2/3={target_right:.1f}pt({'✓' if right_ok else '✗'})"
    )


def all_numbers(rows: dict[str, Row]) -> list[Span]:
    return [r.number_span for r in rows.values() if r.number_span is not None]


def check_page_numbers_basic(page: _Page, rows: dict[str, Row]) -> tuple[bool, str]:
    """所有目录页码位于页面右侧同一纵向区域；页码为黑色阿拉伯数字；
    页码与对应条目位于同一水平行。右对齐且距页面右边线 1.5-3 字符之间。

    严格对应细则五点，面向办公软件（Word/WPS）导出的 PDF：

    点1｜位于页面右侧同一纵向区域：所有页码的水平中心 cx 都位于页面右半侧
      （cx > 页宽/2），且各页码 cx 的极差在同一“纵向区域”内（≤ 一字符）。
    点2｜黑色阿拉伯数字：颜色为黑色（RGB 通道均低）且文本仅由数字字符 0-9 构成。
    点3｜页码与对应条目位于同一水平行：页码 bbox 的 y 中心 ≈ 对应条目行 y。
    点4｜右对齐：所有页码 bbox 的右边缘 x1 一致（Word/WPS 中“右对齐”导出为
      各行 x1 严格对齐，与内容长度无关）。
    点5｜距页面右边线 1.5-3 字符：右边距 (page_w - x1) ∈ [1.5em, 3em]，其中
      “1 字符”按页码自身字号计算（办公软件里“1 字符”对 CJK/数字段落均 = em）。
    """
    # 页码总数：所有条目应各有一个页码
    nums_with_row: list[tuple[Row, Span]] = [
        (r, r.number_span) for r in rows.values() if r.number_span is not None
    ]
    expected_count = 18
    if len(nums_with_row) < expected_count:
        return False, f"找到页码 {len(nums_with_row)}/{expected_count} 个（部分条目缺页码）"

    nums = [n for _, n in nums_with_row]
    page_w = page.rect.width

    # 点1：页面右侧同一纵向区域
    centers = [n.rect.cx for n in nums]
    on_right_half = all(cx > page_w / 2 for cx in centers)
    # “同一纵向区域”容差取一字符（以页码字号中位数）
    num_size = median([n.size for n in nums]) or 16.0
    same_zone_ok = on_right_half and (max(centers) - min(centers) <= num_size)

    # 点2：黑色阿拉伯数字
    black_ok = all(is_black_int(n.color) for n in nums)
    digit_ok = all(n.compact.isdigit() for n in nums)

    # 点3：与对应条目同一水平行
    same_line_ok = all(abs(n.rect.cy - r.y) <= max(3.0, n.size * 0.3) for r, n in nums_with_row)

    # 点4：右对齐（x1 严格对齐；办公软件“右对齐”导出容差通常在 1pt 内，
    # 放宽到 2pt 抵消 PDF 抽取误差）
    right_edges = [n.rect.x1 for n in nums]
    right_align_ok = max(right_edges) - min(right_edges) <= 2.0

    # 点5：距页面右边线 1.5-3 字符
    # “1 字符”按目录页主字号（页码中出现的最大字号，即主层级—— learning/单元/
    # 项目——的字号，Tasks 属于下一层级不作为主字号基准）
    em = max((n.size for n in nums), default=num_size)
    right_margins = [page_w - x1 for x1 in right_edges]
    min_margin_pt = 1.5 * em
    max_margin_pt = 3.0 * em
    margin_ok = all(min_margin_pt <= m <= max_margin_pt for m in right_margins)

    ok = same_zone_ok and black_ok and digit_ok and same_line_ok and right_align_ok and margin_ok
    return ok, (
        f"页码数={len(nums)}；"
        f"右侧同纵向区域={'✓' if same_zone_ok else '✗'}(cx范围{min(centers):.1f}~{max(centers):.1f}pt)；"
        f"黑色={'✓' if black_ok else '✗'}；"
        f"阿拉伯数字={'✓' if digit_ok else '✗'}；"
        f"同水平行={'✓' if same_line_ok else '✗'}；"
        f"右对齐={'✓' if right_align_ok else '✗'}(x1范围{min(right_edges):.1f}~{max(right_edges):.1f}pt)；"
        f"距右边线={min(right_margins):.1f}~{max(right_margins):.1f}pt="
        f"{min(right_margins)/em:.2f}~{max(right_margins)/em:.2f}字符(基准主字号{em:.1f}pt)"
        f"({'✓1.5-3' if margin_ok else '✗超出1.5-3字符'})"
    )


#: 目录应有的全部条目 key（与 build_rows 的 specs 一一对应），用于点状引导线
#: 检查——即便某条目未被 build_rows 识别（文字匹配失败等），也必须计为失败，
#: 不能因为 rows 里没有这个 key 就被跳过检查而"免于扣分"。
ALL_ROW_KEYS: list[str] = [
    "learning",
    "unit1", "task1", "task2", "task3",
    "unit2", "task4", "task5",
    "unit3", "task6", "task7", "task8",
    "unit4", "task9", "task10",
    "mid_project", "final_project", "portfolio",
]


def check_all_dotted_leaders(page: _Page, rows: dict[str, Row], keys: list[str]) -> tuple[bool, str]:
    """对每个目录条目行独立验证“点状引导线”是否满足细则三点：
    1) 目录条目文字与页码之间使用黑色点状引导线（点为黑色，位于条目文字右侧、页码左侧）
    2) 引导线横向排列（点在同一水平行，横向跨度 >> 纵向跨度）
    3) 引导线不穿过条目文字和页码数字（点严格位于 (文字x1, 页码x0) 区间内）

    `keys` 须传入固定的期望条目列表（如 ALL_ROW_KEYS），而不是 `rows.keys()`：
    若某条目未被 build_rows 识别（文字匹配失败、页码缺失等），它不会出现在
    rows 中——若以 rows.keys() 作为遍历对象，这类缺失条目会被直接跳过检查，
    导致本项在条目缺失时仍可能得分。这里显式将缺失条目计为失败。
    """
    dots = small_black_dots(page)  # 只收集黑色候选点，覆盖“颜色为黑色”这一点
    failures: list[str] = []
    details: list[str] = []
    for key in keys:
        row = rows.get(key)
        if not row or not row.number_span:
            failures.append(key)
            details.append(f"{key}:缺失条目或页码")
            continue
        ok, ev = has_dotted_leader(dots, row.y, row.text_right, row.number_span.rect.x0)
        details.append(f"{key}:{ev}")
        if not ok:
            failures.append(key)
    return not failures, "；".join(details[:4]) + (f"；失败 {failures}" if failures else "")


def check_text_alignment_and_spacing(rows: dict[str, Row]) -> tuple[bool, str]:
    """“目录”横线下方的文字保持左对齐。

    严格对应细则，面向办公软件（Word/WPS）导出的 PDF：

    左对齐：Word/WPS 中“左对齐”是段落属性——同层级（同缩进级别）的多行
      文字，其左边缘 x0 一致；若是“居中”则每行 x0 随内容长度变化，若是“右对齐”
      则每行右边缘一致而 x0 各异。因此按 单元 / 任务 / 项目 三个层级分别检查
      每组内 x0 是否稳定。
    """
    needed = [
        "learning",
        "unit1", "task1", "task2", "task3",
        "unit2", "task4", "task5",
        "unit3", "task6", "task7", "task8",
        "unit4", "task9", "task10",
        "mid_project", "final_project", "portfolio",
    ]
    missing = [k for k in needed if k not in rows]
    if missing:
        return False, f"缺失目录行：{missing}"

    # 左对齐——按层级分组，检查每组内 x0 一致
    task_lefts = [rows[f"task{i}"].text_left for i in range(1, 11)]
    unit_lefts = [rows[f"unit{i}"].text_left for i in range(1, 5)]
    project_lefts = [rows[k].text_left for k in ["mid_project", "final_project", "portfolio"]]
    task_align_ok = max(task_lefts) - min(task_lefts) <= 4
    unit_align_ok = max(unit_lefts) - min(unit_lefts) <= 15
    project_align_ok = max(project_lefts) - min(project_lefts) <= 6
    align_ok = task_align_ok and unit_align_ok and project_align_ok

    ok = align_ok
    return ok, (
        f"左对齐-单元x0={min(unit_lefts):.1f}~{max(unit_lefts):.1f}({'✓' if unit_align_ok else '✗'})，"
        f"任务x0={min(task_lefts):.1f}~{max(task_lefts):.1f}({'✓' if task_align_ok else '✗'})，"
        f"项目x0={min(project_lefts):.1f}~{max(project_lefts):.1f}({'✓' if project_align_ok else '✗'})"
    )


def check_learning_style_and_position(page: _Page, rows: dict[str, Row], title_line_bbox: Rect | None) -> tuple[bool, str]:
    """“学习结构说明”字体字号为黑体或微软雅黑小二、加粗，距离目录下方的
    横线大约 30-35 磅，距离页面左边界 45-50 磅。

    严格对应细则五点，面向办公软件（Word/WPS）导出的 PDF：
    1) 字体：黑体 或 微软雅黑（含 SimHei / STHeiti / Microsoft YaHei / MSYH，
       允许 PDF 子集前缀）。
    2) 字号：小二 = 18 磅（±0.5pt 抽取误差）。
    3) 加粗：办公软件“加粗”按钮 → PDF flags 位 4 置位或字体名含 Bold/Black。
    4) 距目录下方横线大约 30-35 磅：文字 bbox 顶部 (y0) 到横线 bbox 底部 (y1)
       的垂直距离 ∈ [30, 35] 磅。
    5) 距页面左边界 45-50 磅：文字 bbox 左侧 (x0) ∈ [45, 50] 磅。
    """
    row = rows.get("learning")
    if not row:
        return False, "未找到“学习结构说明”"
    sp = row.parts[0]

    font_ok = is_cjk_sans_or_yahei(sp)
    size_ok = nearly(sp.size, 18, 18, 0.5)   # 小二 = 18pt
    bold_ok = is_bold(sp)
    left_ok = 45 <= sp.rect.x0 <= 50

    if title_line_bbox is None:
        gap = None
        gap_ok = False
    else:
        gap = sp.rect.y0 - title_line_bbox.y1
        gap_ok = 30 <= gap <= 35

    ok = font_ok and size_ok and bold_ok and left_ok and gap_ok
    gap_text = "无横线" if gap is None else f"{gap:.1f}pt"
    return ok, (
        f"字体={sp.font}({'✓黑体/微软雅黑' if font_ok else '✗'})，"
        f"字号={sp.size:.1f}pt({'✓小二' if size_ok else '✗非小二18pt'})，"
        f"加粗={'✓' if bold_ok else '✗'}，"
        f"距横线={gap_text}({'✓30-35' if gap_ok else '✗超出30-35磅'})，"
        f"距左边界={sp.rect.x0:.1f}pt({'✓45-50' if left_ok else '✗超出45-50磅'})"
    )


def check_learning_page_and_leader(page: _Page, rows: dict[str, Row]) -> tuple[bool, str]:
    """“学习结构说明”右侧页码为数字文本2，中间用黑色点状引导线连接。

    严格对应细则三点，面向办公软件（Word/WPS）导出的 PDF：
    1) 右侧页码位于“学习结构说明”文字右侧：页码 bbox.x0 > 文字 bbox.x1。
    2) 页码为数字文本"2"：页码 span 归一化后等于字符串 "2"。
       （办公软件里可能被输入成“2”“２”等，compact 归一化后再比对）
    3) 中间用黑色点状引导线连接：文字右边到页码左边之间存在黑色点状引导线
       （复用现有 has_dotted_leader）。
    """
    row = rows.get("learning")
    if not row:
        return False, "未找到“学习结构说明”"
    num = row.number_span
    if not num:
        return False, "“学习结构说明”右侧未识别到页码"

    # 点1：页码位于文字右侧
    right_ok = num.rect.x0 > row.text_right

    # 点2：页码为数字文本 "2"
    is_two = num.compact == "2"

    # 点3：中间用黑色点状引导线连接
    dots = small_black_dots(page)
    leader_ok, leader_ev = has_dotted_leader(dots, row.y, row.text_right, num.rect.x0)

    ok = right_ok and is_two and leader_ok
    return ok, (
        f"页码位置={'✓在右侧' if right_ok else '✗未在右侧'}(文字x1={row.text_right:.1f}, 页码x0={num.rect.x0:.1f})，"
        f"页码文本=\"{num.compact}\"({'✓为2' if is_two else '✗非2'})，"
        f"黑色点状引导线={'✓' if leader_ok else '✗'}({leader_ev})"
    )


def check_units_style(rows: dict[str, Row]) -> tuple[bool, str]:
    """“第一单元”“第二单元”“第三单元”“第四单元”的字体字号为黑体或微软雅黑
    小二，加粗，颜色为白色。

    严格对应细则四点，面向办公软件（Word/WPS）导出的 PDF：
    1) 字体：黑体 或 微软雅黑（SimHei/STHeiti/Microsoft YaHei/MSYH，允许 PDF 子集前缀）。
    2) 字号：小二 = 18 磅（±0.5pt 抽取误差）。
    3) 加粗：办公软件“加粗”按钮 → PDF flags 位 4 置位或字体名含 Bold/Black。
    4) 颜色为白色：RGB 三通道均 ≥ 245（抵消办公软件颜色管理导致的极小偏差）。

    注意：本项只考察“第 X 单元”这四个字标签本身；单元主题名（如“地球与地图”）
    的样式在其它打分项里，本函数不涉及。
    """
    failures: list[str] = []
    details: list[str] = []
    for i in range(1, 5):
        row = rows.get(f"unit{i}")
        if not row:
            failures.append(f"unit{i}(缺失)")
            details.append(f"第{i}单元: 缺失")
            continue
        sp = row.parts[0]  # “第X单元”字段
        font_ok = is_cjk_sans_or_yahei(sp)
        size_ok = nearly(sp.size, 18, 18, 0.5)
        bold_ok = is_bold(sp)
        color_ok = is_white_int(sp.color)
        rgb = rgb_from_int(sp.color)
        row_ok = font_ok and size_ok and bold_ok and color_ok
        if not row_ok:
            failures.append(row.label)
        details.append(
            f"{sp.norm}: 字体={sp.font}({'✓' if font_ok else '✗'})，"
            f"字号={sp.size:.1f}({'✓小二' if size_ok else '✗非18'})，"
            f"加粗={'✓' if bold_ok else '✗'}，颜色RGB={rgb}({'✓白' if color_ok else '✗非白'})"
        )
    return not failures, "；".join(details)


def unit_arrow_rects(page: _Page) -> list[Rect]:
    """收集页面上属于“单元箭头”的绿色横向箭头形状。

    面向办公软件（Word/WPS）导出的 PDF：
    Word 里的“形状-箭头总汇-右箭头”导出后有两种表现：
    - A. 单个 drawing 内含多条线段（多边形路径），构成箭头一体形状；
    - B. 拆成“矩形（re）+ 尖角三角形（多条 l 组成）”两个 drawing；
    两种情况都作为候选，最后合并同一水平行、几何相邻的片段。

    过滤条件（仅几何形态，非评分尺寸）：
    - 绿色填充
    - 形态“横向”：宽 > 高
    """
    parts: list[Rect] = []
    for d in page.get_drawings():
        r = drawing_rect(d)
        if not r or not is_green_tuple(d.get("fill")):
            continue
        # 排除标题下方渐变横线（很扁），保留至少像“形状”的元素
        if r.height < 10:
            continue
        # 横向优先：宽应显著大于高（一体箭头或较宽的矩形段）；尖角段常宽 < 高
        rect_like = r.width >= r.height * 1.5
        tip_like = drawing_items_have_polygon(d) and r.height * 0.3 <= r.width <= r.height * 2.5
        if rect_like or tip_like:
            parts.append(r)

    # 合并同一水平行、相邻的片段（矩形 + 尖角）
    merged: list[Rect] = []
    used = [False] * len(parts)
    for i, r in enumerate(parts):
        if used[i]:
            continue
        group = [r]
        used[i] = True
        for j, other in enumerate(parts[i + 1:], start=i + 1):
            same_row = abs(other.cy - r.cy) <= 3
            gbb = union_bbox(group)
            touches = other.x0 <= gbb.x1 + 3 and other.x1 >= gbb.x0 - 3
            if not used[j] and same_row and touches:
                group.append(other)
                used[j] = True
        arrow = union_bbox(group)
        # 最终形态：横向（宽 > 高）
        if arrow.width > arrow.height:
            merged.append(arrow)
    return sorted(merged, key=lambda r: r.y0)


def _arrow_has_right_tip(page: _Page, arrow: Rect) -> bool:
    """判断合并后的箭头形状“右侧带尖角，左边是矩形”。

    办公软件的“右箭头”形状 PDF 抽取后，右侧尖角表现为：
    - 右侧存在指向右方的多边形/多段线（若干条斜线交汇于一个 x 极大点），且
    - 该尖角段的高度不超过左侧矩形段（尖角在右侧“缩窄”）。
    这里通过检测：在箭头 bbox 内是否存在带非水平/非垂直线段（斜线）的绿色
    drawing 且其 x 右边界接近 arrow.x1；同时是否存在含 're' 的矩形填充位于左侧。
    """
    has_tip = False
    has_rect = False
    for d in page.get_drawings():
        r = drawing_rect(d)
        if not r or not is_green_tuple(d.get("fill")):
            continue
        if r.y1 < arrow.y0 - 2 or r.y0 > arrow.y1 + 2:
            continue
        if r.x1 < arrow.x0 - 2 or r.x0 > arrow.x1 + 2:
            continue
        items = d.get("items", [])
        has_rect_op = any(item and item[0] == "re" for item in items)
        # 检测斜线段（尖角）
        has_slant = False
        for x0, y0, x1, y1 in line_segments_from_drawing(d):
            dx, dy = abs(x1 - x0), abs(y1 - y0)
            if dx >= 3 and dy >= 3:  # 斜向线段
                has_slant = True
                break
        # 位于右半侧且含斜线 → 视为尖角段
        if has_slant and r.cx >= (arrow.x0 + arrow.x1) / 2:
            has_tip = True
        # 位于左半侧且含 're' → 视为矩形段
        if has_rect_op and r.cx <= (arrow.x0 + arrow.x1) / 2 + arrow.width * 0.1:
            has_rect = True
        # 一体箭头：单一 drawing 既有 're' 又有斜线，也算“左矩形右尖角”
        if has_rect_op and has_slant:
            has_tip = True
            has_rect = True
    return has_tip and has_rect


def check_unit_arrows(page: _Page, rows: dict[str, Row]) -> tuple[bool, str]:
    """“第一单元”“第二单元”“第三单元”“第四单元”文本底部均有一个右侧带尖角的
    绿色横向箭头，左边是矩形，右边有一个尖角，文本位于形状内部居中排列，
    背景填充为绿色。

    严格对应细则六点（“绿色横向 + 背景填充为绿色”合并为“绿色填充横向”一点）：
    1) 每个单元均有一个箭头（4 个箭头 —— “均”）。
    2) 箭头位于该单元“第X单元”文本底部（覆盖文本 y 范围）。
    3) 箭头是绿色横向（绿色填充 + 宽 > 高；同时背景填充为绿色）。
    4) 左边是矩形：左半侧存在含 're' 的矩形绿色填充。
    5) 右边有一个尖角：右半侧存在带斜向线段的绿色多边形。
    6) 文本位于形状内部居中排列：文本 bbox 完全被箭头 bbox 包含，
       且文本水平中心 ≈ 箭头水平中心，垂直中心 ≈ 箭头垂直中心。
    """
    arrows = unit_arrow_rects(page)
    count_ok = len(arrows) >= 4
    if not count_ok:
        return False, f"找到绿色横向箭头 {len(arrows)}/4 个（应有4个）"

    ok = True
    details: list[str] = []
    for i in range(1, 5):
        row = rows.get(f"unit{i}")
        if not row:
            ok = False
            details.append(f"unit{i}: 缺失")
            continue
        arrow = min(arrows, key=lambda a: abs(a.cy - row.y))
        sp = row.parts[0]  # “第X单元”文本

        # 点2：箭头位于文本底部（覆盖文本 y 范围）
        under_text = arrow.y0 <= sp.rect.y0 + sp.rect.height * 0.2 and arrow.y1 >= sp.rect.y1 - sp.rect.height * 0.2
        # 点3：绿色横向（在采集阶段已保证），此处只再校验形态
        horizontal = arrow.width > arrow.height
        # 点4+5：左矩形 + 右尖角
        tip_and_rect = _arrow_has_right_tip(page, arrow)
        # 点6：文本位于形状内部居中排列
        inside = arrow.x0 <= sp.rect.x0 and arrow.x1 >= sp.rect.x1 and arrow.y0 <= sp.rect.y0 and arrow.y1 >= sp.rect.y1
        h_centered = abs(sp.rect.cx - (arrow.x0 + arrow.x1) / 2) <= arrow.width * 0.15
        v_centered = abs(sp.rect.cy - arrow.cy) <= max(3.0, arrow.height * 0.2)
        centered = inside and h_centered and v_centered

        row_ok = under_text and horizontal and tip_and_rect and centered
        if not row_ok:
            ok = False
        details.append(
            f"{sp.norm}: 箭头=({arrow.x0:.1f},{arrow.y0:.1f},{arrow.x1:.1f},{arrow.y1:.1f})，"
            f"覆盖文本={'✓' if under_text else '✗'}，"
            f"横向={'✓' if horizontal else '✗'}，"
            f"左矩右尖={'✓' if tip_and_rect else '✗'}，"
            f"文本居中={'✓' if centered else '✗'}"
        )
    return ok, "；".join(details)


def check_unit_arrow_left(page: _Page) -> tuple[bool, str]:
    """四个带尖角的绿色横向箭头最左侧距离页面左边框 45-50 磅。

    严格对应细则两点，面向办公软件（Word/WPS）导出的 PDF：
    1) 存在四个带尖角的绿色横向箭头（依赖 `unit_arrow_rects` 已保证形态与颜色，
       本函数在此基础上再要求四个箭头齐备）。
    2) 每一个箭头的“最左侧”x0 ∈ [45, 50] 磅（页面左边框 = 页面 x=0）。
       办公软件中该边距对应页边距-左 + 段落左缩进 + 形状左边距的组合。
    """
    arrows = unit_arrow_rects(page)
    if len(arrows) < 4:
        return False, f"找到绿色横向箭头 {len(arrows)}/4 个（应为4个）"

    arrows = arrows[:4]
    lefts = [a.x0 for a in arrows]
    per_arrow_ok = [45 <= x <= 50 for x in lefts]
    ok = all(per_arrow_ok)
    detail = "，".join(
        f"箭头{i+1}最左={x:.1f}pt({'✓' if pass_ else '✗'})"
        for i, (x, pass_) in enumerate(zip(lefts, per_arrow_ok))
    )
    return ok, detail


def check_unit_names_position(rows: dict[str, Row], page: _Page) -> tuple[bool, str]:
    """“地球与地图”“陆地与海洋”“天气与气候”“居民与聚落”位于绿色箭头右侧，
    按顺序分别和“第一单元”“第二单元”“第三单元”“第四单元”在同一水平线上，
    距离左侧绿色箭头尖角大约 20-25 磅。

    严格对应细则四点，面向办公软件（Word/WPS）导出的 PDF：
    1) 四个主题名（“地球与地图/陆地与海洋/天气与气候/居民与聚落”）依序对应
       “第一单元/第二单元/第三单元/第四单元”。
    2) 位于绿色箭头**右侧**：主题名 x0 > 对应箭头 x1（箭头包含尖角，x1 即尖角尖端）。
    3) 与“第X单元”在同一水平线上：主题名文字中心 y ≈ 单元标签文字中心 y。
    4) 距离左侧绿色箭头尖角大约 20-25 磅：主题名 x0 − 箭头 x1 ∈ [20, 25]。
       办公软件中该间距对应主题段落左缩进与形状右边距的组合。
    """
    arrows = unit_arrow_rects(page)
    if len(arrows) < 4:
        return False, f"绿色箭头 {len(arrows)}/4 个（应为4个）"

    # 依 y 排序，保证第 i 个箭头对应第 i 个单元行
    arrows = sorted(arrows[:4], key=lambda a: a.cy)
    expected = ["地球与地图", "陆地与海洋", "天气与气候", "居民与聚落"]

    ok = True
    details: list[str] = []
    for idx, name in enumerate(expected):
        i = idx + 1
        row = rows.get(f"unit{i}")
        if not row or len(row.parts) < 2:
            ok = False
            details.append(f"第{i}单元: 缺失单元或主题名")
            continue

        theme = row.parts[1]  # 主题名 span

        # 点1：主题名文本匹配
        name_ok = compact(theme.norm) == compact(name)
        # 依据“第X单元”这一行的中心 y 挑最接近的箭头（按顺序 idx 应与按 y 排序一致）
        unit_label = row.parts[0]
        arrow = min(arrows, key=lambda a: abs(a.cy - unit_label.rect.cy))

        # 点2：位于箭头右侧（arrow.x1 即尖角尖端的最右 x）
        right_of_arrow = theme.rect.x0 >= arrow.x1
        # 点3：与“第X单元”同一水平线
        same_y = abs(theme.rect.cy - unit_label.rect.cy) <= max(3.0, theme.size * 0.3)
        # 点4：距箭头尖角 20-25 磅
        gap = theme.rect.x0 - arrow.x1
        gap_ok = 20 <= gap <= 25

        row_ok = name_ok and right_of_arrow and same_y and gap_ok
        if not row_ok:
            ok = False
        details.append(
            f"{name}: 主题名匹配={'✓' if name_ok else '✗'}({theme.norm})，"
            f"箭头右侧={'✓' if right_of_arrow else '✗'}，"
            f"与第{i}单元同水平={'✓' if same_y else '✗'}，"
            f"距尖角={gap:.1f}pt({'✓20-25' if gap_ok else '✗超出20-25磅'})"
        )
    return ok, "；".join(details)


def check_unit_pages_and_dots(page: _Page, rows: dict[str, Row]) -> tuple[bool, str]:
    """“地球与地图”“陆地与海洋”“天气与气候”“居民与聚落”右侧页码分别为
    数字文本 3、7、9、12，中间用黑色点状引导线连接。

    严格对应细则三点，面向办公软件（Word/WPS）导出的 PDF，对四个主题名分别判定：
    1) 右侧页码：页码位于主题名右侧（page.x0 > 主题名 x1）。
    2) 页码为数字文本 3 / 7 / 9 / 12（按顺序）：`compact` 归一化后严格等于目标数字。
       办公软件里可能被输入成全角 “３” 等，compact 归一化后再比较。
    3) 中间用黑色点状引导线连接：主题名右边到页码左边之间存在黑色点状引导线
       （复用 has_dotted_leader，兼容制表符前导符与手动矢量点两种绘制方式）。
    """
    keys = ["unit1", "unit2", "unit3", "unit4"]
    themes = ["地球与地图", "陆地与海洋", "天气与气候", "居民与聚落"]
    expected_pages = ["3", "7", "9", "12"]

    dots = small_black_dots(page)
    ok = True
    details: list[str] = []
    for key, theme_name, want in zip(keys, themes, expected_pages):
        row = rows.get(key)
        if not row or len(row.parts) < 2:
            ok = False
            details.append(f"{theme_name}: 单元行缺失")
            continue

        theme = row.parts[1]      # 主题名 span
        num = row.number_span     # 页码 span
        if num is None:
            ok = False
            details.append(f"{theme_name}: 未识别到右侧页码")
            continue

        # 点1：页码位于主题名右侧
        right_ok = num.rect.x0 > theme.rect.x1
        # 点2：页码为数字文本（且等于目标值）
        page_text_ok = num.compact == want
        # 点3：中间以黑色点状引导线连接（引导线区间取主题名右端到页码左端）
        leader_ok, leader_ev = has_dotted_leader(dots, row.y, theme.rect.x1, num.rect.x0)

        row_ok = right_ok and page_text_ok and leader_ok
        if not row_ok:
            ok = False
        details.append(
            f"{theme_name}: 页码=\"{num.compact}\"({'✓=' + want if page_text_ok else '✗需=' + want})，"
            f"在主题右侧={'✓' if right_ok else '✗'}，"
            f"黑色点状引导线={'✓' if leader_ok else '✗'}({leader_ev})"
        )
    return ok, "；".join(details)


def check_task_font_and_left(rows: dict[str, Row]) -> tuple[bool, str]:
    """文本“任务 1 …”至“任务 10 …”字体字号宋体或微软雅黑三号，其中的数字与
    右侧文本相隔 1-2 字符，和页面左边界相距 85-95 磅。

    严格对应细则三点，面向办公软件（Word/WPS）导出的 PDF，对 10 个任务行分别判定：
    1) 字体字号：宋体 或 微软雅黑，三号（16 磅，允许 ±0.5pt 抽取误差）。
       “其中的数字”不是独立字段——办公软件里“任务 N”整体是一段文字，数字仅是
       其中的字符；因此字号与字体判定覆盖整行（“任务N”标签 + 右侧标题）。
    2) 数字与右侧文本相隔 1-2 字符：数字 N 的字符 bbox 右端 到 右侧标题 bbox 左端
       的净水平距离 / 一字符宽度 ∈ [1.0, 2.0]。办公软件中“1 字符”对 CJK/数字段落
       均 = em = 字号。
    3) “任务 N”左端与页面左边界相距 85-95 磅：标签 x0 ∈ [85, 95]。
    """
    ok = True
    details: list[str] = []
    for i in range(1, 11):
        row = rows.get(f"task{i}")
        if not row or len(row.parts) < 2:
            ok = False
            details.append(f"任务{i}: 缺失")
            continue
        label_sp, title_sp = row.parts[0], row.parts[1]  # 标签“任务 N”与右侧标题

        # 点1：字体（宋体 或 微软雅黑）与 字号（三号=16pt）
        font_ok = all(is_song_or_yahei_or_equivalent(sp) for sp in (label_sp, title_sp))
        size_ok = all(nearly(sp.size, 16, 16, 0.5) for sp in (label_sp, title_sp))

        # 点2：数字 N 与右侧文本相隔 1-2 字符
        # 数字 N 的右端：从 label_sp 的字符流中找到最后一个属于目标数字的字符 bbox 右端；
        # 若无字符级信息，退化为 label_sp.rect.x1（“任务 N”整体右端）
        digit_right = _digit_char_right_edge(label_sp, str(i))
        if digit_right is None:
            digit_right = label_sp.rect.x1
        gap_pt = title_sp.rect.x0 - digit_right
        # 1 字符 = em = 字号（办公软件对 CJK/数字段落均如此）
        em = title_sp.size if title_sp.size > 0 else 16.0
        gap_chars = gap_pt / em
        gap_ok = 1.0 <= gap_chars <= 2.0

        # 点3：距页面左边界 85-95 磅
        left_ok = 85 <= label_sp.rect.x0 <= 95

        row_ok = font_ok and size_ok and gap_ok and left_ok
        if not row_ok:
            ok = False
        details.append(
            f"任务{i}: 字体={label_sp.font}/{title_sp.font}"
            f"({'✓宋体/微软雅黑' if font_ok else '✗'})，"
            f"字号={label_sp.size:.1f}/{title_sp.size:.1f}pt({'✓三号' if size_ok else '✗非三号16pt'})，"
            f"数字后间距={gap_chars:.2f}字符({gap_pt:.1f}pt)({'✓1-2字符' if gap_ok else '✗超出1-2字符'})，"
            f"左距={label_sp.rect.x0:.1f}pt({'✓85-95' if left_ok else '✗超出85-95磅'})"
        )
    return ok, "；".join(details[:5]) + ("；..." if len(details) > 5 else "")


def _digit_char_right_edge(span: Span, digit_text: str) -> float | None:
    """在 span 的字符流中定位目标数字字符串（如 "1"、"10"）最后一个字符的 x1。

    办公软件里“任务 N”导出为一段文字，PDF 抽取层会给出每个字符的 bbox；
    我们据此取“数字 N”末位字符的右端，作为“数字与右侧文本”的分界。
    若字符流不可用或未找到，返回 None。
    """
    chars = span.chars or []
    if not chars or not digit_text:
        return None
    text = "".join(normalize_text(c.get("c", "")) for c in chars)
    idx = text.rfind(digit_text)
    if idx < 0:
        return None
    last_char = chars[idx + len(digit_text) - 1]
    return to_rect(last_char["bbox"]).x1


def check_task_group(page: _Page, rows: dict[str, Row], unit_key: str, task_keys: list[str], pages: list[str]) -> tuple[bool, str]:
    """任务组通用检查：面向办公软件（Word/WPS）导出的 PDF，对每个任务行独立判定：
    1) 位于对应“第 X 单元 …”下方：任务行的 y > 单元行的 y。
    2) 右侧页码为指定数字文本（compact 归一化后严格等于目标值）。
    3) 标题与右侧页码中间用黑色点状引导线连接：任务标题右端到页码左端的区间
       内存在黑色点状引导线（复用 has_dotted_leader，支持制表符前导符 `.` 字符
       与手动矢量点两种绘制方式）。
    """
    unit = rows.get(unit_key)
    if not unit:
        return False, f"缺失 {unit_key}"
    ok = True
    details: list[str] = []
    dots = small_black_dots(page)
    for key, page_no in zip(task_keys, pages):
        row = rows.get(key)
        if not row:
            ok = False
            details.append(f"{key}: 缺失任务行")
            continue
        if not row.number_span:
            ok = False
            details.append(f"{key}: 未识别到右侧页码")
            continue

        # 点1：位于对应单元下方
        below = row.y > unit.y
        # 点2：页码为目标数字
        page_text_ok = row.number_span.compact == page_no
        # 点3：标题与右侧页码中间用黑色点状引导线连接
        dots_ok, dot_ev = has_dotted_leader(
            dots, row.y, row.text_right, row.number_span.rect.x0
        )

        row_ok = below and page_text_ok and dots_ok
        if not row_ok:
            ok = False
        details.append(
            f"{key}: 位于{unit_key}下方={'✓' if below else '✗'}(任务y={row.y:.1f}/单元y={unit.y:.1f})，"
            f"页码=\"{row.number_span.compact}\"({'✓=' + page_no if page_text_ok else '✗需=' + page_no})，"
            f"黑色点状引导线={'✓' if dots_ok else '✗'}({dot_ev})"
        )
    return ok, "；".join(details)


def check_projects_position_and_pages(page: _Page, rows: dict[str, Row]) -> tuple[bool, str]:
    """“期中项目 校园导览图制作任务”“期末项目 家乡聚落观察报告”“全册学习档案袋
    阶段复盘与成长记录”位于“任务 10 聚落发展与可持续生活”下方，右侧页码分别为
    14、15、16，文本与右侧页码中间用黑色点状引导线连接。

    严格对应细则三点，面向办公软件（Word/WPS）导出的 PDF，对三条项目行分别判定：
    1) 位于“任务 10 聚落发展与可持续生活”下方：项目行 y > 任务10 y。
    2) 右侧页码为数字文本 14 / 15 / 16（按顺序）：`compact` 归一化后严格等于目标值。
    3) 文本与右侧页码中间用黑色点状引导线连接：文本右端 → 页码左端 的区间内
       存在黑色点状引导线（复用 has_dotted_leader，支持制表符前导符 `.` 字符与
       手动矢量点两种绘制方式）。
    """
    task10 = rows.get("task10")
    if not task10:
        return False, "缺失任务10"

    keys = ["mid_project", "final_project", "portfolio"]
    labels = [
        "期中项目 校园导览图制作任务",
        "期末项目 家乡聚落观察报告",
        "全册学习档案袋 阶段复盘与成长记录",
    ]
    pages = ["14", "15", "16"]

    ok = True
    details: list[str] = []
    dots = small_black_dots(page)
    for key, label, page_no in zip(keys, labels, pages):
        row = rows.get(key)
        if not row:
            ok = False
            details.append(f"{label}: 缺失该行")
            continue
        if not row.number_span:
            ok = False
            details.append(f"{label}: 未识别到右侧页码")
            continue

        # 点1：位于“任务 10 …”下方
        below = row.y > task10.y
        # 点2：右侧页码为目标数字
        page_text_ok = row.number_span.compact == page_no
        # 点3：文本与右侧页码中间用黑色点状引导线连接
        dots_ok, dot_ev = has_dotted_leader(
            dots, row.y, row.text_right, row.number_span.rect.x0
        )

        row_ok = below and page_text_ok and dots_ok
        if not row_ok:
            ok = False
        details.append(
            f"{label}: 位于任务10下方={'✓' if below else '✗'}(y={row.y:.1f}/任务10 y={task10.y:.1f})，"
            f"页码=\"{row.number_span.compact}\"({'✓=' + page_no if page_text_ok else '✗需=' + page_no})，"
            f"黑色点状引导线={'✓' if dots_ok else '✗'}({dot_ev})"
        )
    return ok, "；".join(details)


def check_project_font(rows: dict[str, Row]) -> tuple[bool, str]:
    """细则："期中项目：校园导览图制作任务""期末项目：家乡聚落观察报告"
    "全册学习档案袋   阶段复盘与成长记录" 字体字号为黑体或微软雅黑小二、加粗。

    对三条项目行独立判定，每条行的每一 span 都须同时满足：
      1) 字体：黑体或微软雅黑（含 SimHei / STHeiti / Microsoft YaHei / MSYH，
         并剥离办公软件 PDF 子集前缀 "XXXXXX+"）
      2) 字号：小二 = 18pt
      3) 加粗（span.flags bit4 或字体名含 Bold/Black）
    """
    keys = ["mid_project", "final_project", "portfolio"]
    labels = [
        "期中项目：校园导览图制作任务",
        "期末项目：家乡聚落观察报告",
        "全册学习档案袋   阶段复盘与成长记录",
    ]
    ok = True
    details: list[str] = []
    for key, label in zip(keys, labels):
        row = rows.get(key)
        if not row:
            ok = False
            details.append(f"{label}: 缺失该行")
            continue
        fonts = [normalized_font_name(sp) for sp in row.parts]
        sizes = [round(sp.size, 2) for sp in row.parts]
        bolds = [is_bold(sp) for sp in row.parts]
        font_ok = all(is_cjk_sans_or_yahei(sp) for sp in row.parts)
        size_ok = all(nearly(sp.size, 18, 18, 0.5) for sp in row.parts)
        bold_ok = all(bolds)
        row_ok = font_ok and size_ok and bold_ok
        if not row_ok:
            ok = False
        details.append(
            f"{label}: 字体={fonts}(黑体/微软雅黑={'✓' if font_ok else '✗'})，"
            f"字号={sizes}(小二18pt={'✓' if size_ok else '✗'})，"
            f"加粗={bolds}({'✓' if bold_ok else '✗'})"
        )
    return ok, "；".join(details)


def check_project_spacing(rows: dict[str, Row]) -> tuple[bool, str]:
    """细则："期中项目"与"校园导览图制作任务"、"期末项目"与"家乡聚落观察报告"、
    "全册学习档案袋"与"阶段复盘与成长记录" 分别之间空 1-2 字符。

    对三对分别独立判定，每对一个点：
      · 空白距离 = 右段左边 x0 − 左段右边 x1
      · 1 字符 = 该行字号（em），办公软件里"空一个字"即空一个字身宽
      · 闭区间 1.0 ≤ 空白/字号 ≤ 2.0
    """
    keys = ["mid_project", "final_project", "portfolio"]
    lefts_labels = ["期中项目", "期末项目", "全册学习档案袋"]
    rights_labels = ["校园导览图制作任务", "家乡聚落观察报告", "阶段复盘与成长记录"]
    ok = True
    details: list[str] = []
    for key, ll, rl in zip(keys, lefts_labels, rights_labels):
        row = rows.get(key)
        if not row or len(row.parts) < 2:
            ok = False
            details.append(f"“{ll}”/“{rl}”: 缺失该行或未识别到两段文本")
            continue
        left, right = row.parts[0], row.parts[1]
        em = max(left.size, right.size)  # 一字符 = 字号
        gap_pt = right.rect.x0 - left.rect.x1
        gap_chars = gap_pt / max(1.0, em)
        gap_ok = 1.0 <= gap_chars <= 2.0
        if not gap_ok:
            ok = False
        details.append(
            f"“{ll}”/“{rl}”: 空白={gap_pt:.1f}pt≈{gap_chars:.2f}字符"
            f"(期望1-2字符={'✓' if gap_ok else '✗'})"
        )
    return ok, "；".join(details)


def project_arrow_rects(page: _Page) -> list[Rect]:
    """收集页面上的“小型绿色右箭头”候选形状。

    面向办公软件（Word/WPS）导出 PDF：
    - Word 中的“形状-右箭头”导出后可能是单个 drawing 内含矩形+多段线，
      也可能拆成“矩形（re）+ 尖角三角形（多条 l）”两个 drawing；
    - 这里以“绿色填充 + 具备形状高度”作为宽松候选，之后按“同一行且相邻”合并。
    评分阈值（宽/高/x0）不在此函数中收紧，交由调用者按细则闭区间判定。
    """
    parts: list[Rect] = []
    for d in page.get_drawings():
        r = drawing_rect(d)
        if not r or not is_green_tuple(d.get("fill")):
            continue
        if r.height < 3:
            continue
        parts.append(r)

    merged: list[Rect] = []
    used = [False] * len(parts)
    for i, r in enumerate(parts):
        if used[i]:
            continue
        group = [r]
        used[i] = True
        for j, other in enumerate(parts[i + 1:], start=i + 1):
            if used[j]:
                continue
            gbb = union_bbox(group)
            same_row = abs(other.cy - gbb.cy) <= 4
            touches = other.x0 <= gbb.x1 + 3 and other.x1 >= gbb.x0 - 3
            if same_row and touches:
                group.append(other)
                used[j] = True
        merged.append(union_bbox(group))
    return sorted(merged, key=lambda r: r.y0)


def _project_arrow_has_rect_and_tip(page: _Page, arrow: Rect) -> tuple[bool, bool]:
    """判断合并后形状“左边是矩形，右侧是三角尖角”。返回 (has_rect_left, has_tip_right)。"""
    has_rect_left = False
    has_tip_right = False
    mid_x = (arrow.x0 + arrow.x1) / 2
    for d in page.get_drawings():
        r = drawing_rect(d)
        if not r or not is_green_tuple(d.get("fill")):
            continue
        # 落在这个箭头的 bbox 内
        if r.y1 < arrow.y0 - 2 or r.y0 > arrow.y1 + 2:
            continue
        if r.x1 < arrow.x0 - 2 or r.x0 > arrow.x1 + 2:
            continue
        items = d.get("items", [])
        has_rect_op = any(item and item[0] == "re" for item in items)
        has_slant = False
        for x0, y0, x1, y1 in line_segments_from_drawing(d):
            dx, dy = abs(x1 - x0), abs(y1 - y0)
            if dx >= 2 and dy >= 2:
                has_slant = True
                break
        # 位于左半 → 视为矩形段
        if has_rect_op and r.cx <= mid_x:
            has_rect_left = True
        # 位于右半且含斜线 → 视为尖角段
        if has_slant and r.cx >= mid_x:
            has_tip_right = True
        # 一体箭头：单个 drawing 同时含矩形 + 斜线
        if has_rect_op and has_slant:
            has_rect_left = True
            has_tip_right = True
    return has_rect_left, has_tip_right


def check_project_arrows(page: _Page, rows: dict[str, Row]) -> tuple[bool, str]:
    """细则："期中项目...""期末项目...""全册学习档案袋..."的左侧 1 字符的位置
    均有一个小型绿色右箭头，左侧为矩形，右侧为三角尖角，宽 20-25 磅，
    高 30-35 磅，距离页面左边界 52-57 磅。

    对三条项目行独立判定，每条行 6 个点：
      1) 左侧存在一个绿色右箭头（且水平位置在文本左侧 1 字符处，闭区间 [0.5,1.5]）
      2) 左侧为矩形
      3) 右侧为三角尖角
      4) 宽 ∈ [20, 25] pt
      5) 高 ∈ [30, 35] pt
      6) 距页面左边界（arrow.x0）∈ [52, 57] pt
    """
    keys = ["mid_project", "final_project", "portfolio"]
    labels = [
        "期中项目 校园导览图制作任务",
        "期末项目 家乡聚落观察报告",
        "全册学习档案袋 阶段复盘与成长记录",
    ]
    all_arrows = project_arrow_rects(page)
    ok = True
    details: list[str] = []
    for key, label in zip(keys, labels):
        row = rows.get(key)
        if not row:
            ok = False
            details.append(f"{label}: 缺失该行")
            continue
        # 选取"位于文本左侧、垂直上与该行同一水平"的最近箭头
        em = max((sp.size for sp in row.parts), default=18.0)
        row_cy = row.y
        candidates = [a for a in all_arrows if a.x1 <= row.text_left and abs(a.cy - row_cy) <= em]
        if not candidates:
            ok = False
            details.append(f"{label}: 未在该行左侧找到绿色形状")
            continue
        arrow = min(candidates, key=lambda a: row.text_left - a.x1)

        gap_chars = (row.text_left - arrow.x1) / max(1.0, em)
        pos_ok = 0.5 <= gap_chars <= 1.5   # 细则"左侧 1 字符的位置"
        has_rect_left, has_tip_right = _project_arrow_has_rect_and_tip(page, arrow)
        width_ok = 20 <= arrow.width <= 25
        height_ok = 30 <= arrow.height <= 35
        left_ok = 52 <= arrow.x0 <= 57

        row_ok = pos_ok and has_rect_left and has_tip_right and width_ok and height_ok and left_ok
        if not row_ok:
            ok = False
        details.append(
            f"{label}: 位于文本左侧{gap_chars:.2f}字符(1字符位置={'✓' if pos_ok else '✗'})，"
            f"左矩形={'✓' if has_rect_left else '✗'}，右尖角={'✓' if has_tip_right else '✗'}，"
            f"宽={arrow.width:.1f}pt(20-25={'✓' if width_ok else '✗'})，"
            f"高={arrow.height:.1f}pt(30-35={'✓' if height_ok else '✗'})，"
            f"距左边界={arrow.x0:.1f}pt(52-57={'✓' if left_ok else '✗'})"
        )
    return ok, "；".join(details)


def check_all_page_numbers_alignment(page: _Page, rows: dict[str, Row]) -> tuple[bool, str]:
    """细则：所有页码上下对齐，距离页面右边界 50-60 磅。

    两个点独立判定：
      1) 所有页码上下对齐——所有页码 span 的右边缘 x1 位于同一纵向直线上
         （办公软件右对齐时，视觉基准是右边缘）
      2) 距页面右边界 50-60 磅——每个页码的右边距 ∈ [50, 60] pt
    """
    nums = all_numbers(rows)
    if not nums:
        return False, "未找到页码"

    right_edges = [n.rect.x1 for n in nums]
    right_margins = [page.rect.width - n.rect.x1 for n in nums]

    # 点 1：上下对齐——右边缘落在同一纵向线上（容差 1pt 供办公软件浮点舍入）
    edge_span = max(right_edges) - min(right_edges)
    align_ok = edge_span <= 1.0
    # 点 2：右边距每一个都 ∈ [50, 60] pt
    margin_ok = all(50 <= m <= 60 for m in right_margins)

    ok = align_ok and margin_ok
    evidence = (
        f"共{len(nums)}个页码；"
        f"右边缘 x1={min(right_edges):.1f}-{max(right_edges):.1f}pt(跨度{edge_span:.2f}pt，"
        f"上下对齐={'✓' if align_ok else '✗'})；"
        f"距页面右边界={min(right_margins):.1f}-{max(right_margins):.1f}pt"
        f"(期望50-60={'✓' if margin_ok else '✗'})"
    )
    return ok, evidence


def large_white_occlusion(page: _Page) -> tuple[bool, str]:
    """细则："任意一页出现超过 1/3 页面面积的空白遮挡"。

    两个点独立判定：
      1) "空白遮挡"——页面上存在白色填充的绘制形状（办公软件 Word/WPS 里的
         白色文本框、白色矩形等，导出 PDF 后表现为 fill=白色 的 drawing）
      2) 面积超过 1/3 页面面积——按所有白色填充形状的**并集面积**判定
         （避免多个白色块拼接后绕过阈值）

    返回 True 表示违规成立（扣 -3）。
    """
    page_area = page.rect.width * page.rect.height
    threshold = page_area / 3
    white_rects: list[Rect] = []
    for d in page.get_drawings():
        r = drawing_rect(d)
        if not r or r.area <= 0:
            continue
        if is_white_tuple(d.get("fill")):
            white_rects.append(r)
    if not white_rects:
        return False, "未发现白色填充形状"
    union = union_area(white_rects)
    max_single = max(r.area for r in white_rects)
    hit = union > threshold
    return hit, (
        f"白色填充形状 {len(white_rects)} 个；单个最大 {max_single:.0f}pt²；"
        f"并集面积 {union:.0f}pt²/页面 {page_area:.0f}pt²（占比 {union/page_area*100:.1f}%）；"
        f"阈值1/3={threshold:.0f}pt²，超过={'是' if hit else '否'}"
    )


def suspicious_garble_area(page: _Page) -> tuple[bool, str]:
    """细则："任意一页出现超过 1/3 页面面积的乱码"。

    两个点独立判定：
      1) "乱码"字符——办公软件（Word/WPS）导出 PDF 时字体缺失/编码丢失的典型呈现：
         · U+FFFD 替换字符 �、U+FFFC 对象替换字符
         · 方框系列（.notdef 常见呈现）：□ ■ ▯ ▫ ▪ ⬜ ⬛
         · Unicode 类别 Co（私有区，办公软件缺字符常回落到 PUA）
           / Cs（代理项）
      2) 面积超过 1/3 页面面积——按逐字符 bbox 的并集面积判定
         （逐字符更贴近"乱码面积"字面语义，一个 span 内有乱码只累加乱码字身面积）

    返回 True 表示违规成立（扣 -3）。
    """
    garble_chars = {"�", "￼", "□", "■", "▯", "▫", "▪", "⬜", "⬛"}
    bad_rects: list[Rect] = []
    for sp in extract_spans(page):
        for ch in (sp.chars or []):
            c = ch.get("c", "")
            if not c:
                continue
            cat = unicodedata.category(c)
            is_garble = (c in garble_chars) or (cat in {"Co", "Cs"})
            if not is_garble:
                continue
            bbox = ch.get("bbox")
            if bbox:
                bad_rects.append(to_rect(bbox))
    page_area = page.rect.width * page.rect.height
    threshold = page_area / 3
    area = union_area(bad_rects)
    hit = area > threshold
    return hit, (
        f"乱码字符 {len(bad_rects)} 个；"
        f"并集面积 {area:.0f}pt²/页面 {page_area:.0f}pt²（占比 {area/max(page_area,1)*100:.2f}%）；"
        f"阈值1/3={threshold:.0f}pt²，超过={'是' if hit else '否'}"
    )


def image_obstructs_text(page: _Page) -> tuple[bool, str]:
    """细则："任意一页图片遮挡文本"。

    独立判定：
      · 图片：PDF 图片块（Word/WPS 中插入的图片，导出 PDF 后为 block type=1）
      · 文本：页面上的所有文字 span（由 extract_spans 提供）
      · 遮挡：图片 bbox 与任一文字 bbox 存在实质性几何重叠（相交面积 > 0）

    办公软件里的"环绕方式"为"衬于文字下方 / 浮于文字上方 / 四周型 / 紧密型"时，
    图片可能与文字 bbox 相交；细则字面为"遮挡"，不设面积阈值，只要图片 bbox
    与文字 bbox 有真实相交面积即视为遮挡。图片与文字恰好相邻（上下/左右贴边）
    时，抽取坐标存在亚 pt 舍入，可能产生宽或高不足 0.5pt 的细缝"相交"，这不是
    真实遮挡，按与本脚本其它检查一致的 0.5pt 舍入容差排除。

    返回 True 表示违规成立（扣 -3）。
    """
    text_spans = extract_spans(page)
    image_rects: list[Rect] = []
    for im in page.images():
        image_rects.append(Rect(im.rect.x0, im.rect.y0, im.rect.x1, im.rect.y1))
    if not image_rects:
        return False, "页面上未发现图片块"
    if not text_spans:
        return False, f"图片块 {len(image_rects)} 个，页面无文字"

    TOL = 0.5  # 亚 pt 抽取舍入容差：相交宽和高都要超过它才算真实遮挡
    hit_pairs: list[str] = []
    for img in image_rects:
        for sp in text_spans:
            inter = img.intersect(sp.rect)
            if inter and inter.width > TOL and inter.height > TOL:
                hit_pairs.append(sp.norm[:8])
                break
    hit = bool(hit_pairs)
    return hit, (
        f"图片块 {len(image_rects)} 个；文字 span {len(text_spans)} 个；"
        f"图片与文字相交处 {len(hit_pairs)} 处，遮挡={'是' if hit else '否'}"
        + (f"（示例：{hit_pairs[:3]}）" if hit_pairs else "")
    )


def text_clipped_by_boundary(page: _Page) -> tuple[bool, str]:
    """细则："任意一页文本被页面边界裁切"。

    独立判定：
      · 文本：页面上所有文字 span（含逐字符 bbox）
      · 页面边界：page.rect（[0, 0, width, height]）
      · 被裁切：文字 bbox 有部分落在页面边界之外（x0<0 / y0<0 /
        x1>width / y1>height），意味着渲染时该部分不可见

    办公软件（Word/WPS）中"文本被页面边界裁切"的典型成因：
      · 段落缩进 / 负缩进超出页边距
      · 文本框位置越过页边界
      · 页边距被人为设成负值 / 内容溢出
      · 字号过大且行不换行

    判定粒度使用逐字符 bbox（适配层 span.chars[i].bbox）：
    只要页面上任何一个字符的可见 bbox 越过页面边界，即视为该页文本被裁切。
    这比 span bbox 更精准——span 可能横跨页边界，但只有右侧字符溢出的那几个
    才是真正被裁切的字符。

    返回 True 表示违规成立（扣 -3）。
    """
    w, h = page.rect.width, page.rect.height
    clipped_chars: list[str] = []
    tol = 0.5  # 亚 pt 浮点渲染舍入容差
    for sp in extract_spans(page):
        for ch in (sp.chars or []):
            c = ch.get("c", "")
            if not c or c.isspace():
                continue
            bbox = ch.get("bbox")
            if not bbox:
                continue
            cx0, cy0, cx1, cy1 = bbox
            if cx0 < -tol or cy0 < -tol or cx1 > w + tol or cy1 > h + tol:
                clipped_chars.append(c)
        # 兜底：字符级 bbox 缺失时使用 span bbox
        if not sp.chars:
            r = sp.rect
            if r.x0 < -tol or r.y0 < -tol or r.x1 > w + tol or r.y1 > h + tol:
                clipped_chars.append(sp.norm[:8])

    hit = bool(clipped_chars)
    return hit, (
        f"页面尺寸 {w:.0f}×{h:.0f}pt；越界字符 {len(clipped_chars)} 个，"
        f"裁切={'是' if hit else '否'}"
        + (f"（示例：{clipped_chars[:5]}）" if clipped_chars else "")
    )


def negative_page_scan(doc: _Doc) -> dict[str, tuple[bool, str]]:
    checks = {
        "blank_occlusion": ("任意一页出现超过1/3页面面积的空白遮挡", large_white_occlusion),
        "garble": ("任意一页出现超过1/3页面面积的乱码", suspicious_garble_area),
        "image_obstruct": ("任意一页图片遮挡文本", image_obstructs_text),
        "text_clipped": ("任意一页文本被页面边界裁切", text_clipped_by_boundary),
    }
    results: dict[str, tuple[bool, str]] = {}
    for key, (desc, fn) in checks.items():
        hit_pages = []
        evidences = []
        for idx, page in enumerate(doc, start=1):
            bad, ev = fn(page)
            if bad:
                hit_pages.append(idx)
                evidences.append(f"第{idx}页：{ev}")
        results[key] = (bool(hit_pages), "；".join(evidences) if hit_pages else "全册未发现")
    return results


def check_catalog_position(doc: _Doc) -> tuple[bool, str]:
    """目录页位置：目录页位于文件第2页，在“学习结构说明”页面之前。

    严格对应细则两点：
    1) 目录页位于文件第2页（即物理页2，索引1，页顶存在“目录”标题）。
    2) 目录页在“学习结构说明”页面之前（“学习结构说明”作为独立页面出现在页2之后的任意页）。
    办公软件（Word/WPS）中“目录”通常以页顶大标题形式出现，故仅检查页顶区域是否含“目录”即可。
    """
    if len(doc) < 2:
        return False, f"文件页数不足2页，实际{len(doc)}页"

    # 要求1：目录页位于文件第2页
    page2 = doc[1]
    catalog_title = find_span(extract_spans(page2), "目录", y_range=(0, 100))
    if not catalog_title:
        return False, "文件第2页顶部未找到“目录”标题"

    # 要求2：目录页在“学习结构说明”页面之前——在第2页之后的任意页存在该页面
    learning_page_no: int | None = None
    for idx in range(2, len(doc)):
        if "学习结构说明" in compact(doc[idx].page_text()):
            learning_page_no = idx + 1
            break
    if learning_page_no is None:
        return False, "第2页之后未找到“学习结构说明”页面"

    return True, f"目录页=第2页（页顶含“目录”标题）；“学习结构说明”页=第{learning_page_no}页（在目录页之后）"


def score_dimension2(doc: _Doc) -> list[ScoreResult]:
    page = doc[1]
    spans = extract_spans(page)
    rows = build_rows(spans)
    title = get_title_span(spans)
    line_cands = title_line_candidates(page, title)
    title_line_bbox = union_bbox([r for r, _ in line_cands]) if line_cands else None

    results: list[ScoreResult] = []

    def add(score: int, desc: str, check: tuple[bool, str]) -> None:
        matched, evidence = check
        results.append(ScoreResult(score, desc, matched, evidence))

    add(5, "目录页位置：目录页位于文件第2页，在“学习结构说明”页面之前", check_catalog_position(doc))
    add(1, "“目录”两字字体字号为黑体或微软雅黑、小初、加粗，颜色为绿色", check_title_style(title))
    add(3, "“目录”两字距离页面上边界55-65磅，距离页面左边界85-95磅字符且“目”和“录”之间空0.3-1字符", check_title_position(title))
    add(3, "“目录”两字左侧有一个绿色的实心矩形块和两条绿色平行的斜条装饰，斜条从左上向右下倾斜，与文字顶部齐平", check_title_decorations(page, title))
    add(3, "“目录”两字下方有粗0.1-0.2cm的横向绿色渐变线条", check_title_gradient_line(page, title))
    add(5, "“目录”横线下方的文字保持左对齐", check_text_alignment_and_spacing(rows))
    add(3, "所有目录页码位于页面右侧同一纵向区域；页码为黑色阿拉伯数字；页码与对应条目位于同一水平行。右对齐且距页面右边线1.5-3字符之间", check_page_numbers_basic(page, rows))
    add(3, "目录页点状引导线：目录条目文字与页码之间使用黑色点状引导线；引导线横向排列；引导线不穿过条目文字和页码数字", check_all_dotted_leaders(page, rows, ALL_ROW_KEYS))
    add(3, "“学习结构说明”字体字号为黑体或微软雅黑小二、加粗，距离目录下方的横线大约30-35磅，距离页面左边界45-50磅", check_learning_style_and_position(page, rows, title_line_bbox))
    add(3, "“学习结构说明”右侧页码为数字文本2，中间用黑色点状引导线连接", check_learning_page_and_leader(page, rows))
    add(3, "“第一单元”“第二单元”“第三单元”“第四单元”的字体字号为黑体或微软雅黑小二，加粗，颜色为白色", check_units_style(rows))
    add(5, "“第一单元”“第二单元”“第三单元”“第四单元”文本底部均有一个右侧带尖角的绿色横向箭头，左边是矩形，右边有一个尖角，文本位于形状内部居中排列，背景填充为绿色", check_unit_arrows(page, rows))
    add(5, "四个带尖角的绿色横向箭头最左侧距离页面左边框45-50磅", check_unit_arrow_left(page))
    add(5, "“地球与地图”“陆地与海洋”“天气与气候”“居民与聚落”位于绿色箭头右侧，按顺序分别和“第一单元”“第二单元”“第三单元”“第四单元”在同一水平线上，距离左侧绿色箭头尖角大约20-25磅", check_unit_names_position(rows, page))
    add(5, "“地球与地图”“陆地与海洋”“天气与气候”“居民与聚落”右侧页码分别为数字文本3、7、9、12，中间用黑色点状引导线连接", check_unit_pages_and_dots(page, rows))
    add(5, "文本“任务 1 认识地球与经纬网”“任务 2 地图三要素与方向”“任务 3 等高线地形图”“任务 4 七大洲与四大洋”“任务 5 海陆变迁与板块运动”“任务 6 天气预报与空气质量”“任务 7 气温曲线与降水柱状图”“任务 8 世界气候类型与生活”“任务 9 人口分布与聚落形成”“任务 10 聚落发展与可持续生活”字体字号宋体或微软雅黑三号，其中的数字与右侧文本相隔1-2字符，和页面左边界相距85-95磅", check_task_font_and_left(rows))
    add(5, "标题“任务 1 认识地球与经纬网”“任务 2 地图三要素与方向”“任务 3 等高线地形图”位于“第一单元 地球与地图”下方；右侧页码分别为3、4、5，标题与右侧页码中间用黑色点状引导线连接", check_task_group(page, rows, "unit1", ["task1", "task2", "task3"], ["3", "4", "5"]))
    add(5, "标题“任务 4 七大洲与四大洋”“任务 5 海陆变迁与板块运动”位于“第二单元 陆地与海洋”下方；右侧页码分别为7、8，标题与右侧页码中间用黑色点状引导线连接", check_task_group(page, rows, "unit2", ["task4", "task5"], ["7", "8"]))
    add(5, "标题“任务 6 天气预报与空气质量”“任务 7 气温曲线与降水柱状图”“任务 8 世界气候类型与生活”位于“第三单元 天气与气候”下方；右侧页码分别为9、10、11，标题与右侧页码中间用黑色点状引导线连接", check_task_group(page, rows, "unit3", ["task6", "task7", "task8"], ["9", "10", "11"]))
    add(5, "标题“任务 9 人口分布与聚落形成”“任务 10 聚落发展与可持续生活”位于“第四单元 居民与聚落”下方；右侧页码分别为12、13，标题与右侧页码中间用黑色点状引导线连接", check_task_group(page, rows, "unit4", ["task9", "task10"], ["12", "13"]))
    add(3, "“期中项目   校园导览图制作任务”“期末项目   家乡聚落观察报告”“全册学习档案袋   阶段复盘与成长记录”位于“任务 10 聚落发展与可持续生活”下方，右侧页码分别为14、15、16，文本与右侧页码中间用黑色点状引导线连接", check_projects_position_and_pages(page, rows))
    add(3, "“期中项目：校园导览图制作任务”“期末项目：家乡聚落观察报告”“全册学习档案袋   阶段复盘与成长记录”字体字号为黑体或微软雅黑小二、加粗", check_project_font(rows))
    add(3, "“期中项目”“期末项目”“全册学习档案袋”分别与“校园导览图制作任务”“家乡聚落观察报告”“阶段复盘与成长记录”之间空1-2字符", check_project_spacing(rows))
    add(3, "“期中项目   校园导览图制作任务”“期末项目   家乡聚落观察报告”“全册学习档案袋   阶段复盘与成长记录”的左侧1字符的位置均有一个小型绿色右箭头，左侧为矩形，右侧为三角尖角，宽20-25磅，高30-35磅，距离页面左边界52-57磅", check_project_arrows(page, rows))
    add(3, "所有页码上下对齐，距离页面右边界50-60磅", check_all_page_numbers_alignment(page, rows))

    neg = negative_page_scan(doc)
    neg_descs = {
        "blank_occlusion": "任意一页出现超过1/3页面面积的空白遮挡",
        "garble": "任意一页出现超过1/3页面面积的乱码",
        "image_obstruct": "任意一页图片遮挡文本",
        "text_clipped": "任意一页文本被页面边界裁切",
    }
    for key in ["blank_occlusion", "garble", "image_obstruct", "text_clipped"]:
        matched, evidence = neg[key]
        results.append(ScoreResult(-3, neg_descs[key], matched, evidence))

    return results


def print_report(path: str, gates: list[GateResult], score_items: list[ScoreResult] | None) -> None:
    print(f"评估文件：{path}")
    print("\n维度1：可用与可修改性")
    for g in gates:
        mark = "通过" if g.passed else "不通过"
        print(f"- [{mark}] {g.name}：{g.evidence}")
    if not all(g.passed for g in gates):
        print("\n维度1未全部满足，按规则直接判为 0 分，不再检查维度2。")
        print("最终得分：0")
        return

    assert score_items is not None
    total = sum(item.score for item in score_items if item.matched)
    print("\n维度2：完成度评分细则")
    for item in score_items:
        if item.matched:
            sign = "+" if item.score > 0 else ""
            print(f"{sign}{item.score}：{item.desc}")
    print(f"\n最终得分：{total}")


def _locate_pdf(dir_path: str) -> str | None:
    """在给定目录内定位待评估的 PDF。

    优先使用文件名 `七年级地理导学案_添加目录版.pdf`；若未命中，则回退到目录内
    任意一个 .pdf 文件。目录不存在或没有 PDF 时返回 None。
    """
    if not dir_path or not os.path.isdir(dir_path):
        return None
    preferred = os.path.join(dir_path, PREFERRED_PDF_NAME)
    if os.path.isfile(preferred):
        return preferred
    for name in sorted(os.listdir(dir_path)):
        if name.lower().endswith(".pdf"):
            return os.path.join(dir_path, name)
    return None


def _build_dim1_reason(gates: list[GateResult]) -> str:
    fails = [g for g in gates if not g.passed]
    if not fails:
        return ""
    return "；".join(f"{g.name}：{g.evidence}" for g in fails)


def _score_items_to_dim2(items: list[ScoreResult]) -> list[dict[str, Any]]:
    out: list[dict[str, Any]] = []
    for it in items:
        # 正向项：max_delta 为该项满分；未命中 delta=0，命中 delta=max_delta。
        # 负向项（score<0，扣分项）：最佳情形为不触发违规、delta=0，
        #                       触发违规则 delta=score（负值）。
        if it.score >= 0:
            max_delta = it.score
            delta = it.score if it.matched else 0
        else:
            max_delta = it.score
            delta = max_delta if it.matched else 0
        out.append({
            "rule": it.desc,
            "max_delta": max_delta,
            "delta": delta,
            "hit": bool(it.matched),
            "detail": "",
        })
    return out


def evaluate(dir_path: str) -> dict:
    """统一入口：接收脚本所在目录的路径，返回结构化评估结果。

    脚本自己负责在 `dir_path` 内定位并打开待评估的 PDF 文档。
    """
    result: dict[str, Any] = {
        "id": SCRIPT_ID,
        "file_name": "",
        "status": "ok",
        "error": None,
        "dim1_pass": False,
        "dim1_reason": "",
        "dim2_items": [],
        "total_score": 0,
        "max_score": DIM2_MAX_SCORE,
    }
    try:
        pdf_path = _locate_pdf(dir_path)
        if pdf_path is None:
            if os.path.isdir(dir_path):
                candidates = sorted(
                    name for name in os.listdir(dir_path)
                    if not name.startswith(('~$', '.~'))
                )
                result["file_name"] = ", ".join(candidates)
            result["dim1_reason"] = "交付文件格式不符合要求：未找到 PDF 文件"
            return result
        result["file_name"] = os.path.basename(pdf_path)

        doc, gates = dimension1_gates(pdf_path)
        dim1_pass = bool(gates) and all(g.passed for g in gates)
        result["dim1_pass"] = dim1_pass
        result["dim1_reason"] = _build_dim1_reason(gates)

        if not dim1_pass or doc is None:
            result["dim2_items"] = []
            result["total_score"] = 0
            return result

        items = score_dimension2(doc)
        result["dim2_items"] = _score_items_to_dim2(items)
        result["total_score"] = sum(it.score for it in items if it.matched)
        return result
    except Exception as exc:  # 兜底：脚本自身异常
        result["status"] = "error"
        result["error"] = f"{type(exc).__name__}: {exc}"
        return result


if __name__ == "__main__":
    target_dir = sys.argv[1] if len(sys.argv) > 1 else os.path.dirname(os.path.abspath(__file__))
    # 本地调试用：避免 Windows 默认 cp1252 控制台无法编码中文时抛出 UnicodeEncodeError，
    # 直接以 UTF-8 字节写入 stdout.buffer；不修改全局 sys.stdout。
    payload = json.dumps(evaluate(target_dir), ensure_ascii=False, indent=2) + "\n"
    try:
        sys.stdout.buffer.write(payload.encode("utf-8"))
    except AttributeError:
        print(payload)
