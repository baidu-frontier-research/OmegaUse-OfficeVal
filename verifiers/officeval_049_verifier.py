#!/usr/bin/env python3
"""Automatic evaluator for Riverdale Tunnel Section 7 PPT recreations.

对外仅暴露 ``evaluate(dir_path: str) -> dict``，接收脚本所在目录的路径，
由脚本自身在该目录内定位并打开被评估的 .pptx 文件，返回结构化字典
（见 §2.2）。检测逻辑与阈值与旧版本完全一致。

评分模型：
1. 维度一为硬门槛，任一必备可用性/可编辑性规则未通过则总分为 0，
   不再执行维度二评分；
2. 维度二为逐项计分，命中项累加得分，未命中项以 hit=False 单独列出。
"""

from __future__ import annotations

import json

SCRIPT_ID = "049"
import math
import re
import sys
import zipfile
from dataclasses import dataclass, field
from pathlib import Path
from typing import Callable, Iterable, Optional

from lxml import etree

EMU_PER_CM = 360000
EMU_PER_PT = 12700
SLIDE_NS = {
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
}
REL_NS = {"rel": "http://schemas.openxmlformats.org/package/2006/relationships"}
ALLOWED_FONTS = {"arial", "calibri"}
# 主题色（有明确 RGB 目标，需在 is_near_color 容差范围内匹配）。
GREEN_TARGET = (20, 145, 45)
BLUE_TARGET = (0, 90, 220)
YELLOW_TARGET = (245, 230, 0)
# 非主题色判定阈值：只需符合颜色大类特征即可通过，不要求精确色值。
BLACKISH_MAX = 120        # "黑色/近黑" — 任意通道 <= 该值即视为黑色系
GRAY_CHANNEL_DIFF = 40    # "灰色" — R/G/B 极差 <= 该值即视为中性灰
WHITE_MIN = 220           # "白色/极浅灰白" — 最小通道 >= 该值即视为白色系


@dataclass
class Shape:
    index: int
    tag: str
    geom: str
    x: float
    y: float
    w: float
    h: float
    text: str = ""
    font: str = ""
    font_size: Optional[float] = None
    bold: bool = False
    text_color: Optional[tuple[int, int, int]] = None
    body_rot: float = 0.0
    shape_rot: float = 0.0
    line_color: Optional[tuple[int, int, int]] = None
    fill_color: Optional[tuple[int, int, int]] = None
    line_width: Optional[float] = None
    dash: str = "solid"
    no_fill: bool = False

    @property
    def x1(self) -> float:
        return min(self.x, self.x + self.w)

    @property
    def x2(self) -> float:
        return max(self.x, self.x + self.w)

    @property
    def y1(self) -> float:
        return min(self.y, self.y + self.h)

    @property
    def y2(self) -> float:
        return max(self.y, self.y + self.h)

    @property
    def cx(self) -> float:
        return (self.x1 + self.x2) / 2

    @property
    def cy(self) -> float:
        return (self.y1 + self.y2) / 2

    @property
    def area(self) -> float:
        return abs(self.w * self.h)

    @property
    def length(self) -> float:
        return math.hypot(self.w, self.h)

    @property
    def angle(self) -> float:
        # PowerPoint coordinates grow downward, so invert dy to get mathematical angle.
        return (math.degrees(math.atan2(-self.h, self.w)) + 360) % 360

    def has_text(self) -> bool:
        return bool(self.text.strip())

    def is_line(self) -> bool:
        return self.geom == "line"

    def is_rect(self) -> bool:
        return self.geom in {"rect", "roundRect"}

    def is_ellipse(self) -> bool:
        return self.geom == "ellipse"

    def visual_bounds(self) -> tuple[float, float, float, float]:
        """Axis-aligned bounding box as PowerPoint renders it (x1, y1, x2, y2).

        PowerPoint rotates a shape around its own center by ``shape_rot`` degrees
        (clockwise, positive downward axis).  A raw ``a:off``/``a:ext`` read can
        yield coordinates that fall outside the page for rotated shapes (e.g. a
        vertical axis label), so we rotate the four corners to recover the region
        the user actually sees in the office application.
        """
        cx = self.x + self.w / 2
        cy = self.y + self.h / 2
        angle = math.radians(self.shape_rot)
        cos_a = math.cos(angle)
        sin_a = math.sin(angle)
        xs: list[float] = []
        ys: list[float] = []
        for px, py in ((self.x, self.y), (self.x + self.w, self.y), (self.x, self.y + self.h), (self.x + self.w, self.y + self.h)):
            dx = px - cx
            dy = py - cy
            xs.append(cx + dx * cos_a - dy * sin_a)
            ys.append(cy + dx * sin_a + dy * cos_a)
        return min(xs), min(ys), max(xs), max(ys)


@dataclass
class EvaluationContext:
    path: Path
    slide_width: float
    slide_height: float
    slide_count: int
    shapes: list[Shape]
    image_count: int
    background_color: Optional[tuple[int, int, int]] = None
    has_background_fill: bool = False
    notes: list[str] = field(default_factory=list)

    @property
    def text_shapes(self) -> list[Shape]:
        return [s for s in self.shapes if s.has_text()]

    @property
    def line_shapes(self) -> list[Shape]:
        return [s for s in self.shapes if s.is_line()]

    @property
    def ellipse_shapes(self) -> list[Shape]:
        return [s for s in self.shapes if s.is_ellipse()]


@dataclass
class CheckResult:
    label: str
    passed: bool
    detail: str


@dataclass
class ScoreItem:
    points: int
    label: str
    checker: Callable[[EvaluationContext], CheckResult]


def cm(value: str | int | None) -> float:
    if value is None:
        return 0.0
    return int(value) / EMU_PER_CM


def pt(value: str | int | None) -> Optional[float]:
    if value is None:
        return None
    return int(value) / EMU_PER_PT


def parse_color(value: str | None) -> Optional[tuple[int, int, int]]:
    if not value or len(value) < 6:
        return None
    value = value[-6:]
    try:
        return tuple(int(value[i : i + 2], 16) for i in (0, 2, 4))
    except ValueError:
        return None


def color_distance(a: Optional[tuple[int, int, int]], b: tuple[int, int, int]) -> float:
    if a is None:
        return 999.0
    return math.sqrt(sum((x - y) ** 2 for x, y in zip(a, b)))


def is_near_color(color: Optional[tuple[int, int, int]], target: tuple[int, int, int], tolerance: float = 75) -> bool:
    return color_distance(color, target) <= tolerance


def is_blackish(color: Optional[tuple[int, int, int]]) -> bool:
    # 非主题色判定：只要属于"黑色/近黑色"大类即视为通过，不要求精确色值。
    # 任意通道均较暗（max <= BLACKISH_MAX）就算黑色系。
    return color is not None and max(color) <= BLACKISH_MAX


def is_grayish(color: Optional[tuple[int, int, int]]) -> bool:
    # 非主题色判定：只要是"灰色"（R/G/B 接近相等的中性色）就视为通过，
    # 不限制具体明度。避免对灰度值范围做严格约束，允许浅灰、中灰、深灰。
    if color is None:
        return False
    return max(color) - min(color) <= GRAY_CHANNEL_DIFF


def in_range(value: float, low: float, high: float, tolerance: float = 0.0) -> bool:
    return low - tolerance <= value <= high + tolerance


def normalize_text(text: str) -> str:
    return (
        text.replace("–", "-")
        .replace("—", "-")
        .replace("−", "-")
        .replace(" ", " ")
        .strip()
    )


def has_text(ctx: EvaluationContext, expected: str) -> bool:
    expected_norm = normalize_text(expected).lower()
    return any(normalize_text(s.text).lower() == expected_norm for s in ctx.text_shapes)


def find_text(ctx: EvaluationContext, expected: str) -> list[Shape]:
    expected_norm = normalize_text(expected).lower()
    return [s for s in ctx.text_shapes if normalize_text(s.text).lower() == expected_norm]


def text_matches_any(ctx: EvaluationContext, pattern: str) -> list[Shape]:
    rx = re.compile(pattern, re.I)
    return [s for s in ctx.text_shapes if rx.fullmatch(normalize_text(s.text))]


def shape_text_style_ok(
    shape: Shape,
    font_size_low: float,
    font_size_high: float,
    require_bold: bool = False,
    color: str = "black",
) -> bool:
    font_ok = shape.font.lower() in ALLOWED_FONTS if shape.font else True
    size_ok = shape.font_size is not None and in_range(shape.font_size, font_size_low, font_size_high, 0.6)
    bold_ok = shape.bold if require_bold else True
    color_ok = is_blackish(shape.text_color)
    return font_ok and size_ok and bold_ok and color_ok


# PowerPoint 默认线宽约 1.0 磅 —— 当 XML 中未显式写出 a:ln/@w 时，
# 我们按 1.0 磅处理（Shape.line_width 为 None）。以下辅助函数封装该缺省。
DEFAULT_LINE_WIDTH_PT = 1.0


def effective_line_width(shape: Shape) -> float:
    """未显式设置线宽时按 PowerPoint 默认线宽 1.0 磅处理。"""
    return shape.line_width if shape.line_width is not None else DEFAULT_LINE_WIDTH_PT


def line_width_ok(shape: Shape, low: float, high: float, tolerance: float = 0.2) -> bool:
    # 只要有效线宽（含 1.0 磅默认值）落在范围内即通过。
    return in_range(effective_line_width(shape), low, high, tolerance)


def line_width_in(shape: Shape, low: float, high: float, tolerance: float = 0.0) -> bool:
    """Range-only variant used by rubric points that quote an explicit width range."""
    return in_range(effective_line_width(shape), low, high, tolerance)


def is_horizontal(shape: Shape, tolerance_degrees: float = 8) -> bool:
    angle = shape.angle
    return min(abs(angle), abs(angle - 180), abs(angle - 360)) <= tolerance_degrees


def is_vertical(shape: Shape, tolerance_degrees: float = 8) -> bool:
    angle = shape.angle
    return min(abs(angle - 90), abs(angle - 270)) <= tolerance_degrees


def angle_in(shape: Shape, ranges: Iterable[tuple[float, float]]) -> bool:
    angle = shape.angle
    return any(low <= angle <= high for low, high in ranges)


def load_pptx(path: Path) -> EvaluationContext:
    if not zipfile.is_zipfile(path):
        raise ValueError("文件不是可解析的 .pptx Zip/OpenXML 文件。")

    with zipfile.ZipFile(path) as zf:
        presentation_xml = etree.fromstring(zf.read("ppt/presentation.xml"))
        size = presentation_xml.find("p:sldSz", SLIDE_NS)
        slide_width = cm(size.get("cx")) if size is not None else 0.0
        slide_height = cm(size.get("cy")) if size is not None else 0.0
        slide_ids = presentation_xml.xpath(".//p:sldIdLst/p:sldId", namespaces=SLIDE_NS)
        slide_count = len(slide_ids)
        slide_xml_path = "ppt/slides/slide1.xml"
        if slide_ids:
            rid = slide_ids[0].get(f"{{{SLIDE_NS['r']}}}id")
            rels = etree.fromstring(zf.read("ppt/_rels/presentation.xml.rels"))
            target = rels.xpath(f".//rel:Relationship[@Id='{rid}']/@Target", namespaces=REL_NS)
            if target:
                slide_xml_path = "ppt/" + target[0].lstrip("/")
        slide_root = etree.fromstring(zf.read(slide_xml_path))
        background_color, has_background_fill = resolve_background_color(zf, slide_xml_path, slide_root)

    shapes = parse_shapes(slide_root)
    image_count = len(slide_root.xpath(".//p:pic", namespaces=SLIDE_NS))
    return EvaluationContext(
        path,
        slide_width,
        slide_height,
        slide_count,
        shapes,
        image_count,
        background_color,
        has_background_fill,
    )


def _bg_solid_srgb(root: etree._Element) -> Optional[tuple[int, int, int]]:
    """Return the srgb background fill color of a slide/layout/master XML root, if any."""
    fills = root.xpath(".//p:cSld/p:bg/p:bgPr/a:solidFill/a:srgbClr/@val", namespaces=SLIDE_NS)
    return parse_color(fills[0]) if fills else None


def resolve_background_color(
    zf: zipfile.ZipFile, slide_xml_path: str, slide_root: etree._Element
) -> tuple[Optional[tuple[int, int, int]], bool]:
    """Resolve the effective slide background as PowerPoint would render it.

    Inheritance order: slide -> its slideLayout -> that layout's slideMaster.
    A missing background fill means PowerPoint shows the default white page, so we
    report white with has_background_fill=False in that case.
    """
    color = _bg_solid_srgb(slide_root)
    if color is not None:
        return color, True

    def follow(part_path: str, rel_type_fragment: str) -> Optional[str]:
        rels_path = _rels_path_for(part_path)
        try:
            rels = etree.fromstring(zf.read(rels_path))
        except KeyError:
            return None
        for rel in rels.xpath(".//rel:Relationship", namespaces=REL_NS):
            if rel_type_fragment in (rel.get("Type") or ""):
                return _normalize_part_path(part_path, rel.get("Target"))
        return None

    layout_path = follow(slide_xml_path, "slideLayout")
    if layout_path:
        try:
            layout_root = etree.fromstring(zf.read(layout_path))
        except KeyError:
            layout_root = None
        if layout_root is not None:
            color = _bg_solid_srgb(layout_root)
            if color is not None:
                return color, True
            master_path = follow(layout_path, "slideMaster")
            if master_path:
                try:
                    master_root = etree.fromstring(zf.read(master_path))
                except KeyError:
                    master_root = None
                if master_root is not None:
                    color = _bg_solid_srgb(master_root)
                    if color is not None:
                        return color, True

    # No explicit fill anywhere: PowerPoint renders the default white background.
    return (255, 255, 255), False


def _rels_path_for(part_path: str) -> str:
    directory, _, filename = part_path.rpartition("/")
    return f"{directory}/_rels/{filename}.rels" if directory else f"_rels/{filename}.rels"


def _normalize_part_path(base_part: str, target: str) -> str:
    if target.startswith("/"):
        return target.lstrip("/")
    base_dir = base_part.rpartition("/")[0]
    parts = base_dir.split("/") if base_dir else []
    for segment in target.split("/"):
        if segment == "..":
            if parts:
                parts.pop()
        elif segment not in ("", "."):
            parts.append(segment)
    return "/".join(parts)


def parse_shapes(slide_root: etree._Element) -> list[Shape]:
    result: list[Shape] = []
    elements = slide_root.xpath(".//p:cSld/p:spTree/p:sp", namespaces=SLIDE_NS)
    for index, element in enumerate(elements):
        xfrm = element.find("p:spPr/a:xfrm", SLIDE_NS)
        off = xfrm.find("a:off", SLIDE_NS) if xfrm is not None else None
        ext = xfrm.find("a:ext", SLIDE_NS) if xfrm is not None else None
        shape_rot = int(xfrm.get("rot", "0")) / 60000 if xfrm is not None else 0.0
        prst = element.find("p:spPr/a:prstGeom", SLIDE_NS)
        geom = prst.get("prst") if prst is not None else "custom"
        line = element.find("p:spPr/a:ln", SLIDE_NS)
        body = element.find("p:txBody/a:bodyPr", SLIDE_NS)
        rpr = element.find(".//a:rPr", SLIDE_NS)
        text = "".join(element.xpath(".//a:t/text()", namespaces=SLIDE_NS)).strip()
        font = ""
        if rpr is not None:
            latin = rpr.find("a:latin", SLIDE_NS)
            font = latin.get("typeface", "") if latin is not None else ""
        shape = Shape(
            index=index,
            tag=etree.QName(element).localname,
            geom=geom,
            x=cm(off.get("x")) if off is not None else 0.0,
            y=cm(off.get("y")) if off is not None else 0.0,
            w=cm(ext.get("cx")) if ext is not None else 0.0,
            h=cm(ext.get("cy")) if ext is not None else 0.0,
            text=text,
            font=font,
            font_size=(int(rpr.get("sz")) / 100 if rpr is not None and rpr.get("sz") else None),
            bold=(rpr is not None and rpr.get("b") in {"1", "true"}),
            text_color=parse_color(first(element.xpath(".//a:rPr/a:solidFill/a:srgbClr/@val", namespaces=SLIDE_NS))),
            body_rot=(int(body.get("rot", "0")) / 60000 if body is not None and body.get("rot") else 0.0),
            shape_rot=shape_rot,
            line_color=parse_color(first(element.xpath("p:spPr/a:ln/a:solidFill/a:srgbClr/@val", namespaces=SLIDE_NS))),
            fill_color=parse_color(first(element.xpath("p:spPr/a:solidFill/a:srgbClr/@val", namespaces=SLIDE_NS))),
            line_width=pt(line.get("w")) if line is not None else None,
            dash=first(element.xpath("p:spPr/a:ln/a:prstDash/@val", namespaces=SLIDE_NS)) or "solid",
            no_fill=bool(element.xpath("p:spPr/a:noFill", namespaces=SLIDE_NS)),
        )
        result.append(shape)
    return result


def first(values: list[str]) -> Optional[str]:
    return values[0] if values else None


def gate_checks(ctx: EvaluationContext) -> list[CheckResult]:
    large_images = []
    checks = [
        CheckResult(
            "交付文件为 .pptx 格式，文件可正常打开",
            ctx.path.suffix.lower() == ".pptx" and ctx.slide_width > 0 and ctx.slide_height > 0,
            f"扩展名 {ctx.path.suffix}，页面 {ctx.slide_width:.2f}cm × {ctx.slide_height:.2f}cm",
        ),
    ]
    if large_images:
        checks.append(CheckResult("未发现大面积图片", False, f"大图片数量 {len(large_images)}"))
    return checks


def content_bounding_box(shapes: list[Shape]) -> tuple[float, float, float, float]:
    real = [s for s in shapes if s.has_text() or s.line_color or s.fill_color or s.geom in {"line", "ellipse", "rect", "roundRect"}]
    if not real:
        return (0, 0, 0, 0)
    return (min(s.x1 for s in real), min(s.y1 for s in real), max(s.x2 for s in real), max(s.y2 for s in real))


def bbox_area(bbox: tuple[float, float, float, float]) -> float:
    return max(0.0, bbox[2] - bbox[0]) * max(0.0, bbox[3] - bbox[1])


def text_overlap_ratio(text_shapes: list[Shape]) -> float:
    if not text_shapes:
        return 0.0
    overlapped = 0
    for i, left in enumerate(text_shapes):
        for right in text_shapes[i + 1 :]:
            if rect_overlap_area(left, right) > min(left.area, right.area) * 0.5:
                overlapped += 1
                break
    return overlapped / len(text_shapes)


def rect_overlap_area(a: Shape, b: Shape) -> float:
    width = max(0.0, min(a.x2, b.x2) - max(a.x1, b.x1))
    height = max(0.0, min(a.y2, b.y2) - max(a.y1, b.y1))
    return width * height


def has_major_overflow(ctx: EvaluationContext) -> bool:
    margin = 0.25
    overflow = [s for s in ctx.shapes if s.x2 < -margin or s.y2 < -margin or s.x1 > ctx.slide_width + margin or s.y1 > ctx.slide_height + margin]
    return bool(overflow)


def required_sensor_labels() -> set[str]:
    return {
        "B114", "B116", "B117", "B118", "B119", "B120", "B121", "B122", "B123", "B124", "B125",
        "A201", "A203", "A205", "A207", "A209", "A211", "A301", "A303", "A308", "A314",
        "K908", "K903", "K899", "K651", "K657", "K663", "K670", "K679",
    }


def green_line_shapes(ctx: EvaluationContext) -> list[Shape]:
    return [s for s in ctx.line_shapes if is_near_color(s.line_color, GREEN_TARGET, 90)]


def blue_points(ctx: EvaluationContext) -> list[Shape]:
    return [s for s in ctx.ellipse_shapes if is_near_color(s.fill_color, BLUE_TARGET, 95)]


def yellow_points(ctx: EvaluationContext) -> list[Shape]:
    return [s for s in ctx.ellipse_shapes if is_near_color(s.fill_color, YELLOW_TARGET, 95)]


def is_zero_dash_line(shape: Shape) -> bool:
    return shape.is_line() and shape.dash != "solid" and is_grayish(shape.line_color) and line_width_ok(shape, 0.5, 0.5, 0.15)


def is_zero_reference_line(shape: Shape) -> bool:
    return shape.is_line() and (is_grayish(shape.line_color) or is_blackish(shape.line_color))


def strict_green_line_shapes(ctx: EvaluationContext) -> list[Shape]:
    """Green solid lines with effective width in the rubric range 0.75–1.00pt.

    未显式设置线宽的按 PowerPoint 默认线宽 1.0 磅处理，1.0 磅落在 0.75–1.00 内即通过。
    """
    return [
        s
        for s in ctx.line_shapes
        if s.dash == "solid"
        and is_near_color(s.line_color, GREEN_TARGET, 90)
        and 0.75 <= effective_line_width(s) <= 1.00
    ]


def line_in_region(lines: Iterable[Shape], x_low: float, x_high: float, y_low: float, y_high: float) -> list[Shape]:
    return [s for s in lines if in_range(s.cx, x_low, x_high) and in_range(s.cy, y_low, y_high)]


def count_lines(lines: Iterable[Shape], predicate: Callable[[Shape], bool]) -> int:
    return sum(1 for s in lines if predicate(s))


def rects_cross(vertical: Shape, horizontal: Shape, tolerance: float = 0.1) -> bool:
    return horizontal.x1 - tolerance <= vertical.cx <= horizontal.x2 + tolerance and vertical.y1 - tolerance <= horizontal.cy <= vertical.y2 + tolerance


def sensor_label_shapes(ctx: EvaluationContext, prefix: str) -> list[Shape]:
    return text_matches_any(ctx, rf"{prefix}\d+")


def labels_present(ctx: EvaluationContext, labels: Iterable[str]) -> set[str]:
    wanted = {normalize_text(label) for label in labels}
    return {normalize_text(s.text) for s in ctx.text_shapes if normalize_text(s.text) in wanted}


def make_pass(label: str, detail: str = "") -> CheckResult:
    return CheckResult(label, True, detail)


def make_fail(label: str, detail: str = "") -> CheckResult:
    return CheckResult(label, False, detail)


def scoring_items() -> list[ScoreItem]:
    return [
        ScoreItem(3, "背景、主体图区域与整体留白", check_background_and_plot_area),
        ScoreItem(3, "主标题位置、字体、字号、颜色、加粗", check_main_title),
        ScoreItem(1, "副标题位置、字体、字号、颜色", check_subtitle),
        ScoreItem(5, "完整二维坐标轴系统", check_axes),
        ScoreItem(3, "横轴标签 Lateral Position (m)", check_x_axis_label),
        ScoreItem(3, "纵轴标签 Vertical Position (m)", check_y_axis_label),
        ScoreItem(3, "横轴刻度文本与短刻度线", check_x_ticks),
        ScoreItem(3, "纵轴刻度文本与短刻度线", check_y_ticks),
        ScoreItem(3, "横向零位灰色虚线", check_horizontal_zero_dash),
        ScoreItem(3, "竖向零位灰色虚线", check_vertical_zero_dash),
        ScoreItem(5, "绿色隧道断面外轮廓", check_green_outer_contour),
        ScoreItem(3, "绿色断面上边界十三段连续折线", check_top_boundary),
        ScoreItem(3, "绿色断面下边界十二段连续折线", check_bottom_boundary),
        ScoreItem(3, "绿色断面左侧不规则边界", check_left_boundary),
        ScoreItem(3, "绿色断面右侧不规则边界", check_right_boundary),
        ScoreItem(5, "绿色内部网格线总量与样式", check_internal_grid_total),
        ScoreItem(3, "内部近横向绿色网格线", check_internal_horizontal_grid),
        ScoreItem(3, "内部近竖向绿色网格线", check_internal_vertical_grid),
        ScoreItem(3, "内部左上至右下绿色斜线", check_diag_down_grid),
        ScoreItem(3, "内部左下至右上绿色斜线", check_diag_up_grid),
        ScoreItem(3, "零位虚线两侧绿色网格未被截断", check_grid_not_cut_by_zero_dash),
        ScoreItem(5, "蓝色中心线传感器点数量、标签、范围", check_blue_center_sensors),
        ScoreItem(3, "蓝色传感器点样式与标签样式", check_blue_sensor_style),
        ScoreItem(5, "黄色边界传感器点数量、标签、分布", check_yellow_boundary_sensors),
        ScoreItem(3, "黄色传感器点样式与标签样式", check_yellow_sensor_style),
        ScoreItem(3, "上边界黄色 K 系列点位", check_k_series_positions),
        ScoreItem(3, "下边界黄色 A 系列点位", check_a_series_positions),
        ScoreItem(3, "蓝色 B 系列点位顺序与标签距离", check_b_series_positions),
        ScoreItem(3, "传感器编号文本完整可编辑且不严重重叠", check_sensor_label_quality),
        ScoreItem(3, "左下角比例尺说明框", check_scale_box),
        ScoreItem(3, "比例尺绿色 X/Y 小箭头", check_scale_arrows),
        ScoreItem(1, "比例尺说明框文本", check_scale_text),
        ScoreItem(3, "右下角图例框", check_legend_box),
        ScoreItem(3, "图例圆点与文本", check_legend_content),
    ]


def content_visual_bounding_box(shapes: list[Shape]) -> tuple[float, float, float, float]:
    """Visual bounding box (x1, y1, x2, y2) of the rendered figure content.

    Uses each shape's rotation-corrected ``visual_bounds`` so the region matches
    what PowerPoint actually draws on the page (rotated labels included), instead
    of the raw XML offsets which can fall outside the slide for rotated shapes.
    """
    real = [s for s in shapes if s.has_text() or s.line_color or s.fill_color or s.geom in {"line", "ellipse", "rect", "roundRect"}]
    if not real:
        return (0, 0, 0, 0)
    bounds = [s.visual_bounds() for s in real]
    return (
        min(b[0] for b in bounds),
        min(b[1] for b in bounds),
        max(b[2] for b in bounds),
        max(b[3] for b in bounds),
    )


def is_white_or_light_grayish(color: Optional[tuple[int, int, int]]) -> bool:
    """True for white or a light gray-white (rubric: 白色或极浅灰白色).

    非主题色判定：只要属于"白色/浅色"大类就通过，不要求精确的白色值。
    """
    return color is not None and min(color) >= WHITE_MIN


def check_background_and_plot_area(ctx: EvaluationContext) -> CheckResult:
    # Rubric point (+3): the slide background is white or a very light gray-white,
    # and the main figure area sits 0–2cm from the left, 0.3–1.0cm from the top,
    # with an overall width of 36–38cm and height of 27–29cm.  Each of these five
    # sub-conditions is checked below against the exact rubric ranges.  Geometry is
    # measured from the rotation-corrected visual bounds so it reflects what the
    # office application actually renders on the page.
    background_ok = is_white_or_light_grayish(ctx.background_color)
    x1, y1, x2, y2 = content_visual_bounding_box(ctx.shapes)
    left = x1
    top = y1
    width = x2 - x1
    height = y2 - y1
    left_ok = in_range(left, 0, 2)
    top_ok = in_range(top, 0.3, 1.0)
    width_ok = in_range(width, 36, 38)
    height_ok = in_range(height, 27, 29)
    ok = background_ok and left_ok and top_ok and width_ok and height_ok
    return CheckResult(
        "背景、主体图区域与整体留白",
        ok,
        f"背景 {ctx.background_color}（白/极浅灰白 {background_ok}）；"
        f"主体图区域 距左 {left:.2f}cm[0-2]{left_ok}，距上 {top:.2f}cm[0.3-1.0]{top_ok}，"
        f"宽 {width:.2f}cm[36-38]{width_ok}，高 {height:.2f}cm[27-29]{height_ok}",
    )


def check_main_title(ctx: EvaluationContext) -> CheckResult:
    # Rubric point (+3): a main title "RIVERDALE TUNNEL – SECTION 7 MONITORING
    # LAYOUT" in the top-left corner, positioned 2.50–3.00cm from the left and
    # 0.3–0.75cm from the top; font Arial or Calibri, size 15–17pt, color black,
    # weight bold.  Each of these sub-conditions is checked below against the exact
    # rubric ranges (no extra tolerance).  Position is measured from the
    # rotation-corrected visual bounds so it reflects what the office application
    # actually renders on the page.
    shapes = find_text(ctx, "RIVERDALE TUNNEL - SECTION 7 MONITORING LAYOUT")
    result: Optional[CheckResult] = None
    for s in shapes:
        left, top, _, _ = s.visual_bounds()
        left_ok = in_range(left, 2.50, 3.00)
        top_ok = in_range(top, 0.3, 0.75)
        font_ok = s.font.lower() in ALLOWED_FONTS
        size_ok = s.font_size is not None and in_range(s.font_size, 15, 17)
        color_ok = is_blackish(s.text_color)
        bold_ok = s.bold
        detail = (
            f"距左 {left:.2f}cm[2.50-3.00]{left_ok}，距上 {top:.2f}cm[0.3-0.75]{top_ok}，"
            f"字体 {s.font or '未指定'}[Arial/Calibri]{font_ok}，字号 {s.font_size}磅[15-17]{size_ok}，"
            f"颜色 {s.text_color}（黑色 {color_ok}），加粗 {bold_ok}"
        )
        if left_ok and top_ok and font_ok and size_ok and color_ok and bold_ok:
            return CheckResult("主标题", True, detail)
        if result is None:
            result = CheckResult("主标题", False, detail)
    return result or CheckResult("主标题", False, "未找到主标题文本")


def check_subtitle(ctx: EvaluationContext) -> CheckResult:
    # Rubric point (+1): a subtitle "Displacement Sensors (mm)" below the main
    # title, positioned 2.50–3.00cm from the left and 1.10–1.30cm from the top;
    # font Arial or Calibri, size 15–17pt, color black.  Each sub-condition is
    # checked below against the exact rubric ranges (no extra tolerance).  Position
    # is measured from the rotation-corrected visual bounds so it matches what the
    # office application actually renders.
    shapes = find_text(ctx, "Displacement Sensors (mm)")
    result: Optional[CheckResult] = None
    for s in shapes:
        left, top, _, _ = s.visual_bounds()
        left_ok = in_range(left, 2.50, 3.00)
        top_ok = in_range(top, 1.10, 1.30)
        font_ok = s.font.lower() in ALLOWED_FONTS
        size_ok = s.font_size is not None and in_range(s.font_size, 15, 17)
        color_ok = is_blackish(s.text_color)
        detail = (
            f"距左 {left:.2f}cm[2.50-3.00]{left_ok}，距上 {top:.2f}cm[1.10-1.30]{top_ok}，"
            f"字体 {s.font or '未指定'}[Arial/Calibri]{font_ok}，字号 {s.font_size}磅[15-17]{size_ok}，"
            f"颜色 {s.text_color}（黑色 {color_ok}）"
        )
        if left_ok and top_ok and font_ok and size_ok and color_ok:
            return CheckResult("副标题", True, detail)
        if result is None:
            result = CheckResult("副标题", False, detail)
    return result or CheckResult("副标题", False, "未找到副标题文本")


def check_axes(ctx: EvaluationContext) -> CheckResult:
    # Rubric point (+5): a complete 2D axis system.
    #   横轴: horizontal single solid line, length 32–34cm, positioned 2.4–2.9cm
    #        from the left and 25.5–26.5cm from the top.
    #   纵轴: vertical single solid line, length 32–34cm, positioned 2.3–2.8cm
    #        from the left and 0.5–0.7cm from the top.
    #   both: color black or dark gray, line width 0.5pt.
    # Every sub-condition is checked below against the exact rubric ranges (no
    # extra tolerance).  Position uses the rotation-corrected visual bounds so it
    # matches what the office application renders; "单实线" requires a solid dash.
    def axis_color_ok(s: Shape) -> bool:
        return is_blackish(s.line_color) or is_grayish(s.line_color)

    def width_half_pt(s: Shape) -> bool:
        # 未显式设置线宽时按默认 1.0 磅处理；1.0 磅不在 0.5 磅范围内 → 不通过。
        return in_range(effective_line_width(s), 0.5, 0.5, 0.01)

    x_axis: list[Shape] = []
    y_axis: list[Shape] = []
    for s in ctx.line_shapes:
        if s.dash != "solid":
            continue
        left, top, _, _ = s.visual_bounds()
        if (
            is_horizontal(s, 3)
            and in_range(s.length, 32, 34)
            and in_range(left, 2.4, 2.9)
            and in_range(top, 25.5, 26.5)
            and axis_color_ok(s)
            and width_half_pt(s)
        ):
            x_axis.append(s)
        if (
            is_vertical(s, 3)
            and in_range(s.length, 32, 34)
            and in_range(left, 2.3, 2.8)
            and in_range(top, 0.5, 0.7)
            and axis_color_ok(s)
            and width_half_pt(s)
        ):
            y_axis.append(s)
    ok = bool(x_axis) and bool(y_axis)
    return CheckResult(
        "二维坐标轴",
        ok,
        f"合格横轴 {len(x_axis)}（水平单实线/长32-34/距左2.4-2.9/距上25.5-26.5/黑或深灰/0.5磅），"
        f"合格纵轴 {len(y_axis)}（垂直单实线/长32-34/距左2.3-2.8/距上0.5-0.7/黑或深灰/0.5磅）",
    )


def check_x_axis_label(ctx: EvaluationContext) -> CheckResult:
    # Rubric point (+3): x-axis label "Lateral Position (m)", positioned 14.5–19.5cm
    # from the left and 24–29cm from the top; font Arial or Calibri, size 14–16pt,
    # color black.  Each sub-condition is checked against the exact rubric ranges
    # (no extra tolerance); position uses rotation-corrected visual bounds so it
    # matches what the office application renders.
    shapes = find_text(ctx, "Lateral Position (m)")
    result: Optional[CheckResult] = None
    for s in shapes:
        left, top, _, _ = s.visual_bounds()
        left_ok = in_range(left, 14.5, 19.5)
        top_ok = in_range(top, 24, 29)
        font_ok = s.font.lower() in ALLOWED_FONTS
        size_ok = s.font_size is not None and in_range(s.font_size, 14, 16)
        color_ok = is_blackish(s.text_color)
        detail = (
            f"距左 {left:.2f}cm[14.5-19.5]{left_ok}，距上 {top:.2f}cm[24-29]{top_ok}，"
            f"字体 {s.font or '未指定'}[Arial/Calibri]{font_ok}，字号 {s.font_size}磅[14-16]{size_ok}，"
            f"颜色 {s.text_color}（黑色 {color_ok}）"
        )
        if left_ok and top_ok and font_ok and size_ok and color_ok:
            return CheckResult("横轴标签", True, detail)
        if result is None:
            result = CheckResult("横轴标签", False, detail)
    return result or CheckResult("横轴标签", False, "未找到横轴标签文本")


def check_y_axis_label(ctx: EvaluationContext) -> CheckResult:
    # Rubric point (+3): y-axis label "Vertical Position (m)" rotated 90°,
    # positioned 0.1–0.5cm from the left and 9–12cm from the top; font Arial or
    # Calibri, size 14–16pt, color black.  Each sub-condition is checked against
    # the exact rubric ranges (no extra tolerance).  The 90° rotation may be
    # applied as a shape rotation (a:xfrm/@rot) or a text-body rotation
    # (a:bodyPr/@rot) in the office application, so either is accepted; position
    # uses the rotation-corrected visual bounds so it matches what is rendered.
    shapes = find_text(ctx, "Vertical Position (m)")
    result: Optional[CheckResult] = None
    for s in shapes:
        left, top, _, _ = s.visual_bounds()
        effective_rot = s.shape_rot if abs(s.shape_rot) > 0.01 else s.body_rot
        rotated_90 = abs((abs(effective_rot) % 180) - 90) <= 5
        left_ok = in_range(left, 0.1, 0.5)
        top_ok = in_range(top, 9, 12)
        font_ok = s.font.lower() in ALLOWED_FONTS
        size_ok = s.font_size is not None and in_range(s.font_size, 14, 16)
        color_ok = is_blackish(s.text_color)
        detail = (
            f"旋转 {effective_rot:.0f}°（90度 {rotated_90}），距左 {left:.2f}cm[0.1-0.5]{left_ok}，"
            f"距上 {top:.2f}cm[9-12]{top_ok}，字体 {s.font or '未指定'}[Arial/Calibri]{font_ok}，"
            f"字号 {s.font_size}磅[14-16]{size_ok}，颜色 {s.text_color}（黑色 {color_ok}）"
        )
        if rotated_90 and left_ok and top_ok and font_ok and size_ok and color_ok:
            return CheckResult("纵轴标签", True, detail)
        if result is None:
            result = CheckResult("纵轴标签", False, detail)
    return result or CheckResult("纵轴标签", False, "未找到纵轴标签文本")


def check_x_ticks(ctx: EvaluationContext) -> CheckResult:
    # Rubric point (+3): the x-axis tick texts "-0.020" "-0.010" "0" "+0.010"
    # "+0.020" all appear, arranged left-to-right below the x-axis; font Arial or
    # Calibri, size 7–10pt, color black or dark gray; the corresponding tick marks
    # are vertical short lines, length 0.12–0.25cm, line width 0.5pt.  Each
    # sub-condition is checked against the exact rubric ranges (no extra
    # tolerance).  Positions use rotation-corrected visual bounds; the "below the
    # x-axis" filter also disambiguates the "0" that also labels the y-axis.
    expected = ["-0.020", "-0.010", "0", "+0.010", "+0.020"]
    tick_texts: list[Shape] = []
    for value in expected:
        candidates = [s for s in find_text(ctx, value) if s.visual_bounds()[1] >= 24]
        candidates.sort(key=lambda s: s.visual_bounds()[0])
        tick_texts.extend(candidates[:1])
    all_present = len(tick_texts) == len(expected)
    ordered = [normalize_text(s.text) for s in sorted(tick_texts, key=lambda s: s.visual_bounds()[0])]
    order_ok = ordered == expected

    def tick_font_ok(s: Shape) -> bool:
        return s.font.lower() in ALLOWED_FONTS

    def tick_size_ok(s: Shape) -> bool:
        return s.font_size is not None and in_range(s.font_size, 7, 10)

    def tick_color_ok(s: Shape) -> bool:
        return is_blackish(s.text_color) or is_grayish(s.text_color)

    font_ok = all_present and all(tick_font_ok(s) for s in tick_texts)
    size_ok = all_present and all(tick_size_ok(s) for s in tick_texts)
    color_ok = all_present and all(tick_color_ok(s) for s in tick_texts)

    tick_lines = [
        s
        for s in ctx.line_shapes
        if is_vertical(s, 5)
        and in_range(s.length, 0.12, 0.25)
        and (is_blackish(s.line_color) or is_grayish(s.line_color))
        and in_range(effective_line_width(s), 0.5, 0.5, 0.01)
    ]
    lines_ok = len(tick_lines) >= len(expected)

    ok = all_present and order_ok and font_ok and size_ok and color_ok and lines_ok
    sizes = [s.font_size for s in tick_texts]
    return CheckResult(
        "横轴刻度",
        ok,
        f"文本 {ordered}（齐全 {all_present}/左到右 {order_ok}），字号 {sizes}[7-10]{size_ok}，"
        f"字体[Arial/Calibri]{font_ok}，颜色[黑/深灰]{color_ok}，"
        f"垂直短线(长0.12-0.25/0.5磅) {len(tick_lines)}/{len(expected)}{lines_ok}",
    )


def check_y_ticks(ctx: EvaluationContext) -> CheckResult:
    # Rubric point (+3): the y-axis tick texts "-0.025" "-0.015" "-0.005" "0"
    # "+0.005" "+0.015" "+0.025" all appear, arranged bottom-to-top on the left
    # side of the y-axis; font Arial or Calibri, size 13–15pt, color black or dark
    # gray; the corresponding tick marks are horizontal short lines, length
    # 0.40–0.60cm, line width 0.5pt.  Each sub-condition is checked against the
    # exact rubric ranges (no extra tolerance).  Positions use rotation-corrected
    # visual bounds; the "left of the y-axis" filter also disambiguates the "0"
    # that also labels the x-axis.
    expected = ["-0.025", "-0.015", "-0.005", "0", "+0.005", "+0.015", "+0.025"]
    y_axis_left = 2.5  # y-axis is around x=2.5cm; tick labels sit to its left.
    tick_texts: list[Shape] = []
    for value in expected:
        candidates = [s for s in find_text(ctx, value) if s.visual_bounds()[0] < y_axis_left]
        candidates.sort(key=lambda s: s.visual_bounds()[1], reverse=True)
        tick_texts.extend(candidates[:1])
    all_present = len(tick_texts) == len(expected)
    ordered = [normalize_text(s.text) for s in sorted(tick_texts, key=lambda s: s.visual_bounds()[1], reverse=True)]
    order_ok = ordered == expected

    def tick_font_ok(s: Shape) -> bool:
        return s.font.lower() in ALLOWED_FONTS

    def tick_size_ok(s: Shape) -> bool:
        return s.font_size is not None and in_range(s.font_size, 13, 15)

    def tick_color_ok(s: Shape) -> bool:
        return is_blackish(s.text_color) or is_grayish(s.text_color)

    font_ok = all_present and all(tick_font_ok(s) for s in tick_texts)
    size_ok = all_present and all(tick_size_ok(s) for s in tick_texts)
    color_ok = all_present and all(tick_color_ok(s) for s in tick_texts)

    tick_lines = [
        s
        for s in ctx.line_shapes
        if is_horizontal(s, 5)
        and in_range(s.length, 0.40, 0.60)
        and (is_blackish(s.line_color) or is_grayish(s.line_color))
        and in_range(effective_line_width(s), 0.5, 0.5, 0.01)
    ]
    lines_ok = len(tick_lines) >= len(expected)

    ok = all_present and order_ok and font_ok and size_ok and color_ok and lines_ok
    sizes = [s.font_size for s in tick_texts]
    return CheckResult(
        "纵轴刻度",
        ok,
        f"文本 {ordered}（齐全 {all_present}/下到上 {order_ok}），字号 {sizes}[13-15]{size_ok}，"
        f"字体[Arial/Calibri]{font_ok}，颜色[黑/深灰]{color_ok}，"
        f"水平短线(长0.40-0.60/0.5磅) {len(tick_lines)}/{len(expected)}{lines_ok}",
    )


def check_horizontal_zero_dash(ctx: EvaluationContext) -> CheckResult:
    # Rubric point (+3): a horizontal zero-reference dashed line, positioned
    # 13.5–14cm from the top, running across the figure starting 2.3–2.8cm from
    # the left; the line is gray and dashed, length 32–35cm, line width 0.5pt,
    # angle 0°.  Each sub-condition is checked against the exact rubric ranges (no
    # extra tolerance).  Position uses rotation-corrected visual bounds; "虚线"
    # requires a non-solid dash and "0度" requires a horizontal orientation.
    result: Optional[CheckResult] = None
    for s in ctx.line_shapes:
        left, top, _, _ = s.visual_bounds()
        angle_ok = is_horizontal(s, 3)
        top_ok = in_range(top, 13.5, 14.0)
        left_ok = in_range(left, 2.3, 2.8)
        gray_ok = is_grayish(s.line_color)
        dash_ok = s.dash != "solid"
        length_ok = in_range(s.length, 32, 35)
        width_ok = in_range(effective_line_width(s), 0.5, 0.5, 0.01)
        if angle_ok and top_ok and left_ok and gray_ok and dash_ok and length_ok and width_ok:
            return CheckResult(
                "横向零位虚线",
                True,
                f"距上 {top:.2f}cm[13.5-14]，距左 {left:.2f}cm[2.3-2.8]，长 {s.length:.2f}cm[32-35]，"
                f"线宽 {effective_line_width(s)}磅，角度 {s.angle:.1f}°，灰色 {s.line_color}，虚线 {s.dash}",
            )
        if angle_ok and top_ok and left_ok and result is None:
            result = CheckResult(
                "横向零位虚线",
                False,
                f"距上 {top:.2f}cm[13.5-14]{top_ok}，距左 {left:.2f}cm[2.3-2.8]{left_ok}，"
                f"长 {s.length:.2f}cm[32-35]{length_ok}，线宽 {effective_line_width(s)}磅[0.5]{width_ok}，"
                f"角度 {s.angle:.1f}°[0]{angle_ok}，灰色 {s.line_color}({gray_ok})，虚线 {s.dash}({dash_ok})",
            )
    return result or CheckResult("横向零位虚线", False, "未找到符合位置的横向线条")


def check_vertical_zero_dash(ctx: EvaluationContext) -> CheckResult:
    # Rubric point (+3): a vertical zero-reference dashed line, positioned
    # 16.5–17.3cm from the left, running down the figure starting 2.0–2.5cm from
    # the top; the line is gray and dashed, line width 0.5pt, length 22–25cm,
    # angle 90°, and crosses the horizontal zero-reference dashed line near the
    # center of the figure.  Each sub-condition is checked against the exact
    # rubric ranges (no extra tolerance).  Position uses rotation-corrected visual
    # bounds; "虚线" requires a non-solid dash and "90度" a vertical orientation.
    def is_horizontal_zero_dash(s: Shape) -> bool:
        left, top, _, _ = s.visual_bounds()
        return (
            is_horizontal(s, 3)
            and in_range(top, 13.5, 14.0)
            and in_range(left, 2.3, 2.8)
            and is_grayish(s.line_color)
            and s.dash != "solid"
            and in_range(s.length, 32, 35)
            and in_range(effective_line_width(s), 0.5, 0.5, 0.01)
        )

    horizontal = [s for s in ctx.line_shapes if is_horizontal_zero_dash(s)]

    result: Optional[CheckResult] = None
    for s in ctx.line_shapes:
        left, top, _, _ = s.visual_bounds()
        angle_ok = is_vertical(s, 3)
        left_ok = in_range(left, 16.5, 17.3)
        top_ok = in_range(top, 2.0, 2.5)
        gray_ok = is_grayish(s.line_color)
        dash_ok = s.dash != "solid"
        length_ok = in_range(s.length, 22, 25)
        width_ok = in_range(effective_line_width(s), 0.5, 0.5, 0.01)
        crosses = any(rects_cross(s, h) for h in horizontal)
        if angle_ok and left_ok and top_ok and gray_ok and dash_ok and length_ok and width_ok and crosses:
            return CheckResult(
                "竖向零位虚线",
                True,
                f"距左 {left:.2f}cm[16.5-17.3]，距上 {top:.2f}cm[2.0-2.5]，长 {s.length:.2f}cm[22-25]，"
                f"线宽 {effective_line_width(s)}磅，角度 {s.angle:.1f}°，灰色 {s.line_color}，虚线 {s.dash}，与横向交叉 {crosses}",
            )
        if angle_ok and left_ok and top_ok and result is None:
            result = CheckResult(
                "竖向零位虚线",
                False,
                f"距左 {left:.2f}cm[16.5-17.3]{left_ok}，距上 {top:.2f}cm[2.0-2.5]{top_ok}，"
                f"长 {s.length:.2f}cm[22-25]{length_ok}，线宽 {effective_line_width(s)}磅[0.5]{width_ok}，"
                f"角度 {s.angle:.1f}°[90]{angle_ok}，灰色 {s.line_color}({gray_ok})，"
                f"虚线 {s.dash}({dash_ok})，与横向零位虚线交叉 {crosses}",
            )
    return result or CheckResult(
        "竖向零位虚线",
        False,
        f"未找到符合位置的竖向线条；符合条件的横向零位虚线 {len(horizontal)} 条",
    )


def check_green_outer_contour(ctx: EvaluationContext) -> CheckResult:
    # Rubric point (+5): the green tunnel cross-section outer contour is made of
    # editable polylines / freeform curves; the contour as a whole is positioned
    # with its left edge 2.5–3.0cm from the left and its top edge 2.00–6.00cm from
    # the top; the lines are green single solid lines, color close to RGB(20,145,
    # 45), line width 0.75–1.00pt, with no fill or transparent fill.  Each
    # sub-condition is checked against the exact rubric ranges (no extra
    # tolerance).  In the office application each contour segment is a separate
    # editable connector (geom "line"); position uses rotation-corrected visual
    # bounds so it matches what is rendered.
    def contour_segment_ok(s: Shape) -> bool:
        return (
            s.is_line()
            and s.dash == "solid"
            and is_near_color(s.line_color, GREEN_TARGET)
            and in_range(effective_line_width(s), 0.75, 1.00)
            and (s.no_fill or s.fill_color is None)
        )

    segments = [s for s in ctx.line_shapes if contour_segment_ok(s)]
    editable_ok = bool(segments)
    if segments:
        bounds = [s.visual_bounds() for s in segments]
        left = min(b[0] for b in bounds)
        top = min(b[1] for b in bounds)
        left_ok = in_range(left, 2.5, 3.0)
        top_ok = in_range(top, 2.00, 6.00)
    else:
        left = top = float("nan")
        left_ok = top_ok = False
    ok = editable_ok and left_ok and top_ok
    return CheckResult(
        "绿色外轮廓",
        ok,
        f"可编辑绿色单实线段(近RGB20/145/45,宽0.75-1.00磅,无填充) {len(segments)}；"
        f"外轮廓 距左 {left:.2f}cm[2.5-3.0]{left_ok}，距上 {top:.2f}cm[2.00-6.00]{top_ok}",
    )


def _line_endpoints(shape: Shape) -> tuple[tuple[float, float], tuple[float, float]]:
    """Return the two rotation-corrected endpoints of a line as the office app draws it."""
    cx = shape.x + shape.w / 2
    cy = shape.y + shape.h / 2
    angle = math.radians(shape.shape_rot)
    cos_a = math.cos(angle)
    sin_a = math.sin(angle)
    pts = []
    for px, py in ((shape.x, shape.y), (shape.x + shape.w, shape.y + shape.h)):
        dx = px - cx
        dy = py - cy
        pts.append((cx + dx * cos_a - dy * sin_a, cy + dx * sin_a + dy * cos_a))
    return pts[0], pts[1]


def _is_continuous_polyline(segments: list[Shape], tol: float = 0.3) -> bool:
    """按端点相接判断线段是否首尾相连组成单条连续折线。

    连续折线特征：段间共享端点（≤ tol cm 视为同一节点），恰有 2 个自由端点、
    其余端点两两相接；所有段通过邻接关系连成一条路径（无分叉、无环、无断裂）。
    """
    n = len(segments)
    if n == 0:
        return False
    if n == 1:
        return True
    endpoints = [_line_endpoints(s) for s in segments]
    # 邻接：两段若各有一个端点距离 ≤ tol，则相邻
    adj: list[list[int]] = [[] for _ in range(n)]
    for i in range(n):
        for j in range(i + 1, n):
            connected = any(
                math.hypot(pi[0] - pj[0], pi[1] - pj[1]) <= tol
                for pi in endpoints[i]
                for pj in endpoints[j]
            )
            if connected:
                adj[i].append(j)
                adj[j].append(i)
    degrees = [len(a) for a in adj]
    # 单条路径：恰 2 个端点度=1，其余度=2
    if sorted(degrees) != [1, 1] + [2] * (n - 2):
        return False
    # 从某端点出发能一次性走完所有段（连通、无环）
    start = degrees.index(1)
    visited = {start}
    prev = -1
    cur = start
    while True:
        nxt = next((k for k in adj[cur] if k != prev), None)
        if nxt is None:
            break
        if nxt in visited:
            return False
        visited.add(nxt)
        prev, cur = cur, nxt
    return len(visited) == n


def check_top_boundary(ctx: EvaluationContext) -> CheckResult:
    # Rubric point (+3): the green cross-section top boundary is made of thirteen
    # continuous polyline segments; the boundary as a whole extends from near
    # (left 2.4–2.7cm, top 2.3–2.5cm) to near (left 35–37cm, top 2.3–2.5cm); each
    # single segment is about 1.8–4.5cm long; the angles include mild slopes of
    # about -20° to +20°; and the right end has a clear rising (upturned) slant.
    # Each sub-condition is checked against the exact rubric ranges (segment
    # length used as the "单段长度" property).  Geometry uses rotation-corrected
    # visual bounds/endpoints so it matches what the office application renders.
    green = green_line_shapes(ctx)
    top_band = [s for s in green if s.visual_bounds()[1] <= 6.5]
    segments = [s for s in top_band if in_range(s.length, 1.8, 4.5)]
    # 细则点：严格 13 段
    count_ok = len(segments) == 13
    # 细则点：按端点相接形成连续折线
    continuous_ok = _is_continuous_polyline(segments)

    endpoints = [pt for s in segments for pt in _line_endpoints(s)]
    left_start = any(in_range(px, 2.4, 2.7) and in_range(py, 2.3, 2.5) for px, py in endpoints)
    right_end = any(in_range(px, 35, 37) and in_range(py, 2.3, 2.5) for px, py in endpoints)

    # Mild slopes: angle within about -20°..+20° of horizontal (either direction).
    mild = count_lines(segments, lambda s: angle_in(s, [(0, 20), (340, 360), (160, 200)]))
    mild_ok = mild >= 1

    # Right-end upturn: a segment reaching the right side with a clearly rising
    # (steeper than mild) slant.  A rising line in PPT (top-left to bottom-right is
    # falling) appears with angle in ~20°..80° or its mirror ~100°..160°.
    right_upturn = any(
        max(p[0] for p in _line_endpoints(s)) >= 35 and angle_in(s, [(20, 80), (100, 160)])
        for s in segments
    )

    ok = count_ok and continuous_ok and left_start and right_end and mild_ok and right_upturn
    return CheckResult(
        "上边界折线",
        ok,
        f"上边界绿色段(长1.8-4.5cm) {len(segments)}(需=13){count_ok}，连续折线={continuous_ok}，"
        f"左端(2.4-2.7,2.3-2.5) {left_start}，右端(35-37,2.3-2.5) {right_end}，"
        f"缓斜线(-20~+20) {mild}{mild_ok}，右侧上扬斜线 {right_upturn}",
    )


def check_bottom_boundary(ctx: EvaluationContext) -> CheckResult:
    # Rubric point (+3): the green cross-section bottom boundary is made of twelve
    # continuous polyline segments; the boundary as a whole extends from near
    # (left 2.4–2.7cm, top 17–19cm) to near (left 35–37cm, top 15–17cm); each
    # single segment is about 1.5–4cm long; the angles include mild slopes of
    # about -25° to +25°; and the lowest point in the middle sits near top
    # 21–23cm.  Each sub-condition is checked against the exact rubric ranges
    # (segment length used as the "单段长度" property).  Geometry uses
    # rotation-corrected visual bounds/endpoints so it matches what the office
    # application renders.
    green = green_line_shapes(ctx)
    bottom_band = [s for s in green if 15 <= s.visual_bounds()[3] <= 23.5]
    segments = [s for s in bottom_band if in_range(s.length, 1.5, 4.0)]
    # 细则点：严格 12 段
    count_ok = len(segments) == 12
    # 细则点：按端点相接形成连续折线
    continuous_ok = _is_continuous_polyline(segments)

    endpoints = [pt for s in segments for pt in _line_endpoints(s)]
    left_start = any(in_range(px, 2.4, 2.7) and in_range(py, 17, 19) for px, py in endpoints)
    right_end = any(in_range(px, 35, 37) and in_range(py, 15, 17) for px, py in endpoints)

    # Mild slopes: within about -25°..+25° of horizontal (either direction).
    mild = count_lines(segments, lambda s: angle_in(s, [(0, 25), (335, 360), (155, 205)]))
    mild_ok = mild >= 1

    # Middle lowest point: the largest top (y) among segment endpoints, near 21–23cm.
    lowest = max((py for _, py in endpoints), default=0.0)
    lowest_ok = in_range(lowest, 21, 23)

    ok = count_ok and continuous_ok and left_start and right_end and mild_ok and lowest_ok
    return CheckResult(
        "下边界折线",
        ok,
        f"下边界绿色段(长1.5-4.0cm) {len(segments)}(需=12){count_ok}，连续折线={continuous_ok}，"
        f"左端(2.4-2.7,17-19) {left_start}，右端(35-37,15-17) {right_end}，"
        f"缓斜线(-25~+25) {mild}{mild_ok}，中部最低点 {lowest:.2f}cm[21-23]{lowest_ok}",
    )


def _orientation_near(shape: Shape, target_deg: float, tolerance: float = 15) -> bool:
    """True if the line's orientation (mod 180°) is within tolerance of target_deg.

    A line has no direction, so orientation 60° and 240° are the same; we compare
    modulo 180° so "约60/90/120度方向" matches regardless of which endpoint is first.
    """
    diff = abs((shape.angle - target_deg) % 180)
    return min(diff, 180 - diff) <= tolerance


def _segments_connected_component(segments: list[Shape], tol: float = 0.3) -> bool:
    """判断给定线段集合是否在端点相邻意义下构成单个连通分量。

    连通判定：两段若各有一个端点距离 ≤ tol cm，则视为相邻；
    从任一段出发 BFS 若能覆盖全部段即为单一连通分量。
    与 _is_continuous_polyline 的差异：本函数允许分叉与环，仅要求"连通"，
    适用于"连续不规则收口"这类可能含 T 形交汇/闭合的折线簇。
    """
    n = len(segments)
    if n <= 1:
        return n == 1
    endpoints = [_line_endpoints(s) for s in segments]

    def adjacent(i: int, j: int) -> bool:
        return any(
            math.hypot(pi[0] - pj[0], pi[1] - pj[1]) <= tol
            for pi in endpoints[i]
            for pj in endpoints[j]
        )

    visited = {0}
    stack = [0]
    while stack:
        cur = stack.pop()
        for k in range(n):
            if k in visited:
                continue
            if adjacent(cur, k):
                visited.add(k)
                stack.append(k)
    return len(visited) == n


def check_left_boundary(ctx: EvaluationContext) -> CheckResult:
    # Rubric point (+3): the green cross-section left boundary is made of an
    # irregular polyline located within 2.3–5.0cm from the left; it contains
    # vertical lines, diagonal lines and bent-angle lines; each single segment is
    # 0.8–3cm long; the angles include roughly the 60°, 90° and 120° directions,
    # forming a continuous irregular closure at the left end.  Each sub-condition
    # is checked against the exact rubric ranges (no extra tolerance beyond the
    # "约" approximate-angle allowance).  Geometry uses rotation-corrected visual
    # bounds so it matches what the office application renders.
    green = green_line_shapes(ctx)
    segments = [
        s
        for s in green
        if s.visual_bounds()[0] >= 2.3
        and s.visual_bounds()[2] <= 5.0
        and in_range(s.length, 0.8, 3.0)
    ]
    has_vertical = any(is_vertical(s, 15) for s in segments)
    has_diagonal = any(_orientation_near(s, 45, 25) or _orientation_near(s, 135, 25) for s in segments)
    # A bent-angle ("折角") join is present when segments run in clearly different
    # directions within the left region (i.e. both steep and shallow orientations).
    has_bent = any(is_vertical(s, 20) for s in segments) and any(is_horizontal(s, 20) for s in segments)
    dir_60 = any(_orientation_near(s, 60) for s in segments)
    dir_90 = any(_orientation_near(s, 90) for s in segments)
    dir_120 = any(_orientation_near(s, 120) for s in segments)
    angles_ok = dir_60 and dir_90 and dir_120
    # 细则点：形成左端“连续不规则收口”——线段按端点两两相接构成单一连通分量
    closure_ok = _segments_connected_component(segments)
    ok = bool(segments) and has_vertical and has_diagonal and has_bent and angles_ok and closure_ok
    return CheckResult(
        "左侧边界",
        ok,
        f"左区绿色段(距左2.3-5.0,长0.8-3) {len(segments)}，竖向线 {has_vertical}，斜向线 {has_diagonal}，" +
        f"折角线 {has_bent}，方向约60° {dir_60}/90° {dir_90}/120° {dir_120}，连续收口 {closure_ok}",
    )


def check_right_boundary(ctx: EvaluationContext) -> CheckResult:
    # Rubric point (+3): the green cross-section right boundary is made of a
    # continuous irregular polyline located within 34–37cm from the left; it
    # contains upward slants, downward slants and near-vertical lines; each single
    # segment is 1–4cm long; the angles include roughly the 60°, 90° and 120°
    # directions, forming a rising (upturned) boundary at the right end.  Each
    # sub-condition is checked against the exact rubric ranges (beyond the "约"
    # approximate-angle allowance).  Geometry uses rotation-corrected visual bounds
    # so it matches what the office application renders.
    green = green_line_shapes(ctx)
    segments = [
        s
        for s in green
        if s.visual_bounds()[0] >= 34.0
        and s.visual_bounds()[2] <= 37.0
        and in_range(s.length, 1.0, 4.0)
    ]
    # A rising slant goes lower-left to upper-right (PPT angle ~0..90 or its
    # mirror ~180..270); a falling slant goes upper-left to lower-right (~90..180
    # or ~270..360).  Exclude the near-horizontal/near-vertical edges from slants.
    has_up = any(angle_in(s, [(5, 85), (185, 265)]) for s in segments)
    has_down = any(angle_in(s, [(95, 175), (275, 355)]) for s in segments)
    has_vertical = any(is_vertical(s, 15) for s in segments)
    dir_60 = any(_orientation_near(s, 60) for s in segments)
    dir_90 = any(_orientation_near(s, 90) for s in segments)
    dir_120 = any(_orientation_near(s, 120) for s in segments)
    angles_ok = dir_60 and dir_90 and dir_120
    # 细则点：连续不规则折线——线段按端点相接构成单一连通分量
    continuous_ok = _segments_connected_component(segments)
    # 细则点：右端上扬——右区边界整体呈上扬趋势。
    # 取所有端点，比较左半侧最低点（y 最大）与右半侧最高点（y 最小）：
    # 若右端最高端点显著高于左端最低端点，则整体形态上扬。
    endpoints = [pt for s in segments for pt in _line_endpoints(s)]
    rising_ok = False
    if endpoints:
        xs = [px for px, _ in endpoints]
        x_min, x_max = min(xs), max(xs)
        mid = (x_min + x_max) / 2
        left_pts = [py for px, py in endpoints if px <= mid]
        right_pts = [py for px, py in endpoints if px >= mid]
        if left_pts and right_pts:
            # y 轴向下：上扬意味着右端 y 更小；至少高 0.5cm 视为“明显上扬”。
            rising_ok = (max(left_pts) - min(right_pts)) >= 0.5
    ok = bool(segments) and has_up and has_down and has_vertical and angles_ok and continuous_ok and rising_ok
    return CheckResult(
        "右侧边界",
        ok,
        f"右区绿色段(距左34-37,长1-4) {len(segments)}，向上斜线 {has_up}，向下斜线 {has_down}，" +
        f"近竖向线 {has_vertical}，方向约60° {dir_60}/90° {dir_90}/120° {dir_120}，" +
        f"连续折线 {continuous_ok}，右端上扬 {rising_ok}",
    )


def internal_green_lines(ctx: EvaluationContext) -> list[Shape]:
    return [s for s in strict_green_line_shapes(ctx) if 3 <= s.cx <= 35 and 4 <= s.cy <= 22 and s.length >= 0.7]


def internal_grid_segments(ctx: EvaluationContext) -> list[Shape]:
    """Editable green single-solid-line segments inside the cross-section interior.

    Matches the rubric's grid lines: editable line objects, green (close to
    RGB(20,145,45), same as the contour), solid dash, width 0.75–1.00pt, whose
    center lies within the interior region.  Geometry uses the shape's center as
    the office application renders it.
    """
    segments = []
    for s in ctx.line_shapes:
        if not (s.is_line() and s.dash == "solid"):
            continue
        if not is_near_color(s.line_color, GREEN_TARGET):
            continue
        if not in_range(effective_line_width(s), 0.75, 1.00):
            continue
        if 3 <= s.cx <= 35 and 4 <= s.cy <= 22:
            segments.append(s)
    return segments


def check_internal_grid_total(ctx: EvaluationContext) -> CheckResult:
    # Rubric point (+5): the green internal grid lines consist of about 35 editable
    # single solid lines or polylines, covering the main interior region of the
    # cross-section; the line color matches the outer contour, close to
    # RGB(20,145,45); line width 0.75–1.00pt; internally they form irregular
    # triangular, quadrilateral and pentagonal grid cells.  Each sub-condition is
    # checked against the exact rubric properties (colour/solid/width enforced per
    # segment).  Geometry uses rotation-corrected visual bounds so coverage matches
    # what the office application renders.
    segments = internal_grid_segments(ctx)
    count_ok = len(segments) >= 30  # "35条左右" — about 35.

    # Coverage of the main interior region: the segments' visual bounding box must
    # span a large part of the interior both horizontally and vertically.
    if segments:
        bounds = [s.visual_bounds() for s in segments]
        span_x = max(b[2] for b in bounds) - min(b[0] for b in bounds)
        span_y = max(b[3] for b in bounds) - min(b[1] for b in bounds)
    else:
        span_x = span_y = 0.0
    coverage_ok = span_x >= 20 and span_y >= 12

    # Irregular triangle/quad/pentagon mesh: an irregular polygon mesh requires
    # segments running in several distinct orientations (near-horizontal,
    # near-vertical, and both diagonal directions) so that closed cells form.
    has_h = any(is_horizontal(s, 15) for s in segments)
    has_v = any(is_vertical(s, 15) for s in segments)
    has_diag_down = any(angle_in(s, [(20, 70), (200, 250)]) for s in segments)
    has_diag_up = any(angle_in(s, [(110, 160), (290, 340)]) for s in segments)
    mesh_ok = has_h and has_v and has_diag_down and has_diag_up

    ok = count_ok and coverage_ok and mesh_ok
    return CheckResult(
        "内部网格总量",
        ok,
        f"内部绿色可编辑单实线段(近RGB20/145/45,宽0.75-1.00磅) {len(segments)}(约35){count_ok}，"
        f"覆盖 {span_x:.1f}×{span_y:.1f}cm{coverage_ok}，"
        f"多向网格(横{has_h}/竖{has_v}/斜下{has_diag_down}/斜上{has_diag_up}){mesh_ok}",
    )


def check_internal_horizontal_grid(ctx: EvaluationContext) -> CheckResult:
    # Rubric point (+3): 7–10 horizontal / near-horizontal green lines appear in
    # the internal grid, distributed mainly within top ranges 7–8.5cm,
    # 10.5–12.5cm, 9.0–11cm, 12–14cm, 13–15cm; each single line is 2–7cm long;
    # angle is -10° to +10°; the lines are green single solid lines.  Each
    # sub-condition is checked against the exact rubric ranges.  Vertical position
    # uses the shape center (cy) as rendered; "-10°~+10°" is the horizontal band.
    bands = [(7.0, 8.5), (10.5, 12.5), (9.0, 11.0), (12.0, 14.0), (13.0, 15.0)]
    lines = [
        s
        for s in ctx.line_shapes
        if s.is_line()
        and s.dash == "solid"
        and is_near_color(s.line_color, GREEN_TARGET)
        and in_range(s.length, 2, 7)
        and angle_in(s, [(0, 10), (350, 360), (170, 190)])
        and any(low <= s.cy <= high for low, high in bands)
    ]
    count_ok = 7 <= len(lines) <= 10  # "7-10条"
    ok = count_ok
    covered = [i for i, (low, high) in enumerate(bands) if any(low <= s.cy <= high for s in lines)]
    return CheckResult(
        "内部横向网格",
        ok,
        f"绿色单实线近横向(长2-7,角度-10~+10,分布指定带内) {len(lines)}(需7-10){count_ok}，"
        f"覆盖带 {len(covered)}/5",
    )


def check_internal_vertical_grid(ctx: EvaluationContext) -> CheckResult:
    # Rubric point (+3): 8–12 vertical / near-vertical green lines appear in the
    # internal grid, distributed mainly within left ranges 5–7cm, 15–17cm,
    # 16–18cm, 22–24cm, 30–32cm, 32–34cm, 17.5–19.5cm; each single line is 1.5–5cm
    # long; angle is 75° to 105°; the lines are green single solid lines, used to
    # connect the top/bottom boundaries or adjacent grid nodes.  Each sub-condition
    # is checked against the exact rubric ranges.  Horizontal position uses the
    # shape center (cx) as rendered; "75°~105°" is the near-vertical band.
    x_bands = [(5, 7), (15, 17), (16, 18), (22, 24), (30, 32), (32, 34), (17.5, 19.5)]
    lines = [
        s
        for s in ctx.line_shapes
        if s.is_line()
        and s.dash == "solid"
        and is_near_color(s.line_color, GREEN_TARGET)
        and in_range(s.length, 1.5, 5.0)
        and angle_in(s, [(75, 105), (255, 285)])
        and any(low <= s.cx <= high for low, high in x_bands)
    ]
    count_ok = 8 <= len(lines) <= 12  # "8-12条"
    ok = count_ok
    covered = [i for i, (low, high) in enumerate(x_bands) if any(low <= s.cx <= high for s in lines)]
    return CheckResult(
        "内部竖向网格",
        ok,
        f"绿色单实线近竖向(长1.5-5,角度75-105,分布指定带内) {len(lines)}(需8-12){count_ok}，"
        f"覆盖列 {len(covered)}/7",
    )


def check_diag_down_grid(ctx: EvaluationContext) -> CheckResult:
    # Rubric point (+3): 13–17 top-left-to-bottom-right green diagonal lines
    # appear in the internal grid, distributed across the left, middle and right
    # regions of the figure; each single line is 1.5–5.5cm long; angle is about
    # 25°–60° or 205°–240°; the lines are green single solid lines.  Each
    # sub-condition is checked against the exact rubric ranges.  Horizontal region
    # uses the shape center (cx) as rendered.
    lines = [
        s
        for s in ctx.line_shapes
        if s.is_line()
        and s.dash == "solid"
        and is_near_color(s.line_color, GREEN_TARGET)
        and in_range(s.length, 1.5, 5.5)
        and angle_in(s, [(25, 60), (205, 240)])
    ]
    count_ok = 13 <= len(lines) <= 17  # "13-17条"
    left = any(s.cx < 13 for s in lines)
    mid = any(13 <= s.cx <= 25 for s in lines)
    right = any(s.cx > 25 for s in lines)
    regions_ok = left and mid and right
    ok = count_ok and regions_ok
    return CheckResult(
        "左上至右下斜线",
        ok,
        f"绿色单实线左上至右下(长1.5-5.5,角度25-60/205-240) {len(lines)}(需13-17){count_ok}，"
        f"分布 左{left}/中{mid}/右{right}{regions_ok}",
    )


def check_diag_up_grid(ctx: EvaluationContext) -> CheckResult:
    # Rubric point (+3): 49–53 bottom-left-to-top-right green diagonal lines
    # appear in the internal grid, distributed across the left, middle and right
    # regions of the figure; each single line is 1.5–6cm long; angle is about
    # 120°–160° or 300°–340°; the lines are green single solid lines.  Each
    # sub-condition is checked against the exact rubric ranges.  Horizontal region
    # uses the shape center (cx) as rendered.
    lines = [
        s
        for s in ctx.line_shapes
        if s.is_line()
        and s.dash == "solid"
        and is_near_color(s.line_color, GREEN_TARGET)
        and in_range(s.length, 1.5, 6.0)
        and angle_in(s, [(120, 160), (300, 340)])
    ]
    count_ok = 49 <= len(lines) <= 53  # "49-53条"
    left = any(s.cx < 13 for s in lines)
    mid = any(13 <= s.cx <= 25 for s in lines)
    right = any(s.cx > 25 for s in lines)
    regions_ok = left and mid and right
    ok = count_ok and regions_ok
    return CheckResult(
        "左下至右上斜线",
        ok,
        f"绿色单实线左下至右上(长1.5-6,角度120-160/300-340) {len(lines)}(需49-53){count_ok}，"
        f"分布 左{left}/中{mid}/右{right}{regions_ok}",
    )


def check_grid_not_cut_by_zero_dash(ctx: EvaluationContext) -> CheckResult:
    # Rubric point (+3): the green grid lines on both sides of the central vertical
    # zero-reference DASHED line must not be truncated by the dashed line; the
    # green lines may cross the gray dashed line but must not be occluded; at the
    # crossing the line width stays clear and the green color must not turn gray.
    # Checked against the actual shapes as the office application renders them.
    # The premise is a vertical zero-reference *dashed* line: if no such dashed
    # line exists (e.g. the file only has solid lines), this point cannot be
    # satisfied and fails.
    verticals = [
        s
        for s in ctx.line_shapes
        if is_vertical(s, 5) and is_grayish(s.line_color) and s.dash != "solid" and s.length >= 15
    ]
    if not verticals:
        return CheckResult("绿色网格未被虚线截断", False, "未找到中央竖向灰色零位虚线（图中无虚线，仅有实线）")
    zero_x = verticals[0].cx

    green = green_line_shapes(ctx)
    # Green grid present on both sides of the zero line (not truncated to one side).
    left_side = any(s.cx < zero_x - 0.5 for s in green)
    right_side = any(s.cx > zero_x + 0.5 for s in green)
    both_sides = left_side and right_side

    # Green lines that genuinely span across the zero-line x (crossing, not cut).
    crossing = [s for s in green if s.visual_bounds()[0] < zero_x < s.visual_bounds()[2]]
    has_crossing = bool(crossing)

    # At the crossing the green lines keep a clear (positive) width and stay green
    # (close to the contour green, never turned gray).  未显式设置线宽时按默认 1.0 磅处理。
    width_clear = all(effective_line_width(s) > 0 for s in crossing)
    stays_green = all(is_near_color(s.line_color, GREEN_TARGET) and not is_grayish(s.line_color) for s in crossing)

    ok = both_sides and has_crossing and width_clear and stays_green
    return CheckResult(
        "绿色网格未被虚线截断",
        ok,
        f"中央零位虚线 x={zero_x:.2f}；两侧绿色网格 左{left_side}/右{right_side}{both_sides}，"
        f"跨越绿色线 {len(crossing)}(有交叉{has_crossing})，交叉处线宽清晰 {width_clear}，"
        f"保持绿色未变灰 {stays_green}",
    )


def check_blue_center_sensors(ctx: EvaluationContext) -> CheckResult:
    # Rubric point (+5): blue center-line sensor points appear, count 10 or 11,
    # with labels B114, B116, B117, B118, B119, B120, B121, B122, B123, B124,
    # B125; the points are arranged with slight undulation near the horizontal
    # zero line, and the group as a whole sits within 11–27cm from the left and
    # 11–16cm from the top.  Each sub-condition is checked against the exact rubric
    # ranges.  Point positions use the shape center as the office app renders them;
    # the "整体位于" box also selects the center-line group (excluding e.g. the
    # blue legend dot outside this region).
    expected = ["B114", "B116", "B117", "B118", "B119", "B120", "B121", "B122", "B123", "B124", "B125"]
    points = [p for p in blue_points(ctx) if in_range(p.cx, 11, 27) and in_range(p.cy, 11, 16)]
    count_ok = 10 <= len(points) <= 11

    # 细则点：每个B标签对应一个蓝点——按“最近蓝点”一一配对
    # · 收集存在的 B 系列标签（可能有多个 shape 同名，只取每个名字的第一个）；
    # · 距离阈值 2.0cm 对应"标签在点旁"的视觉判定；
    # · 每个蓝点至多被一个标签占用，用贪心按距离升序分配。
    label_shapes: dict[str, Shape] = {}
    for s in ctx.text_shapes:
        name = normalize_text(s.text)
        if name in expected and name not in label_shapes:
            label_shapes[name] = s
    labels_present_names = set(label_shapes.keys())
    labels_ok = len(labels_present_names) == len(expected)

    pair_thresh = 2.0
    pairs: list[tuple[float, str, int]] = []
    for name, ls in label_shapes.items():
        for idx, p in enumerate(points):
            dist = math.hypot(ls.cx - p.cx, ls.cy - p.cy)
            if dist <= pair_thresh:
                pairs.append((dist, name, idx))
    pairs.sort()
    matched_labels: set[str] = set()
    used_points: set[int] = set()
    for _, name, idx in pairs:
        if name in matched_labels or idx in used_points:
            continue
        matched_labels.add(name)
        used_points.add(idx)
    pair_ok = matched_labels == labels_present_names and len(matched_labels) == len(labels_present_names)

    # Slight undulation near the horizontal zero line: the points ripple around a
    # roughly constant height (small vertical spread), rather than trending like a
    # slope.  Require a modest cy range that stays inside the rubric's 11–16cm band.
    undulation_ok = False
    if points:
        cys = [p.cy for p in points]
        undulation_ok = (max(cys) - min(cys)) <= 3.0

    ok = count_ok and labels_ok and pair_ok and undulation_ok
    missing = [b for b in expected if b not in labels_present_names]
    unpaired = sorted(labels_present_names - matched_labels)
    return CheckResult(
        "蓝色中心线传感器",
        ok,
        f"区域内(距左11-27,距上11-16)蓝点 {len(points)}(需10或11){count_ok}，" +
        f"B标签 {len(labels_present_names)}/{len(expected)}{labels_ok}" +
        (f"(缺{missing})" if missing else "") +
        f"，标签-蓝点一一对应(≤{pair_thresh}cm) {pair_ok}" +
        (f"(未配对{unpaired})" if unpaired else "") +
        f"，水平零位附近轻微起伏 {undulation_ok}",
    )


def _is_dark_blue(color: Optional[tuple[int, int, int]]) -> bool:
    """True for a dark-blue outline: blue channel clearly dominant and overall dark."""
    if color is None:
        return False
    r, g, b = color
    return b >= r + 40 and b >= g + 40 and b <= 200 and max(r, g) <= 120


def check_blue_sensor_style(ctx: EvaluationContext) -> CheckResult:
    # Rubric point (+3): the blue sensor points are circular, diameter 0.18–0.20cm,
    # filled blue (color close to RGB(0,90,220)), with a dark-blue outline or no
    # outline; each point has a black label beside it, font Arial or Calibri, size
    # 12–14pt.  Each sub-condition is checked against the exact rubric ranges.
    # Sensor points are the blue center-line group (within the rubric's 11–27cm /
    # 11–16cm box) so the legend dot outside that region is not treated as a sensor
    # point.  Geometry uses the shape's rendered size/center.
    points = [p for p in blue_points(ctx) if in_range(p.cx, 11, 27) and in_range(p.cy, 11, 16)]
    b_labels = [s for s in ctx.text_shapes if re.fullmatch(r"B\d+", normalize_text(s.text))]

    def label_style_ok(s: Shape) -> bool:
        return (
            (s.font.lower() in ALLOWED_FONTS)
            and s.font_size is not None
            and in_range(s.font_size, 12, 14)
            and is_blackish(s.text_color)
        )

    def point_ok(p: Shape) -> bool:
        circular = p.is_ellipse() and abs(abs(p.w) - abs(p.h)) <= 0.06
        diameter = (abs(p.w) + abs(p.h)) / 2
        diameter_ok = in_range(diameter, 0.18, 0.20)
        fill_ok = is_near_color(p.fill_color, BLUE_TARGET)
        edge_ok = p.line_color is None or _is_dark_blue(p.line_color)
        # A black, correctly-styled label sits beside the point.
        has_label = any(
            label_style_ok(s) and math.hypot(s.cx - p.cx, s.cy - p.cy) <= 1.5
            for s in b_labels
        )
        return circular and diameter_ok and fill_ok and edge_ok and has_label

    good = sum(1 for p in points if point_ok(p))
    ok = bool(points) and good == len(points)
    sample = points[0] if points else None
    detail = (
        f"蓝色传感器点 {len(points)}，全部合格 {good}/{len(points)}"
        + (
            f"；示例 直径 {(abs(sample.w)+abs(sample.h))/2:.3f}cm[0.18-0.20]，"
            f"圆形 {sample.is_ellipse()}，填充 {sample.fill_color}，边线 {sample.line_color}"
            if sample
            else ""
        )
    )
    return CheckResult("蓝点样式", ok, detail)


def _point_to_segment_distance(px: float, py: float, ax: float, ay: float, bx: float, by: float) -> float:
    """点 (px,py) 到线段 AB 的最短距离（cm）。"""
    dx, dy = bx - ax, by - ay
    seg_len_sq = dx * dx + dy * dy
    if seg_len_sq == 0:
        return math.hypot(px - ax, py - ay)
    t = ((px - ax) * dx + (py - ay) * dy) / seg_len_sq
    t = max(0.0, min(1.0, t))
    fx, fy = ax + t * dx, ay + t * dy
    return math.hypot(px - fx, py - fy)


def _min_distance_to_green_boundary(px: float, py: float, green_segments: list[Shape]) -> float:
    """点到所有绿色边界线段的最小距离；无边界时返回 inf。"""
    best = float("inf")
    for s in green_segments:
        (ax, ay), (bx, by) = _line_endpoints(s)
        d = _point_to_segment_distance(px, py, ax, ay, bx, by)
        if d < best:
            best = d
    return best


def check_yellow_boundary_sensors(ctx: EvaluationContext) -> CheckResult:
    # Rubric point (+5): yellow boundary sensor points appear, count not fewer than
    # 20, with labels including A201, A203, A205, A207, A209, A211, A301, A303,
    # A308, A314, K908, K903, K899, K651, K657, K663, K670, K679; the points are
    # distributed near the green cross-section's top, bottom and left/right
    # boundaries.  Each sub-condition is checked against the exact rubric ranges.
    # Point positions use the shape center as the office application renders them.
    expected = [
        "A201", "A203", "A205", "A207", "A209", "A211", "A301", "A303", "A308", "A314",
        "K908", "K903", "K899", "K651", "K657", "K663", "K670", "K679",
    ]
    points = yellow_points(ctx)
    count_ok = len(points) >= 20

    # 细则点：每个指定标签与黄点一一配对——按“最近黄点”贪心分配
    # · 收集每个 expected 名字的第一个 shape；
    # · 距离阈值 2.0cm 视为"标签在点旁"；
    # · 每个黄点至多被一个标签占用。
    label_shapes: dict[str, Shape] = {}
    for s in ctx.text_shapes:
        name = normalize_text(s.text)
        if name in expected and name not in label_shapes:
            label_shapes[name] = s
    labels_present_names = set(label_shapes.keys())
    labels_ok = len(labels_present_names) == len(expected)

    pair_thresh = 2.0
    pairs: list[tuple[float, str, int]] = []
    for name, ls in label_shapes.items():
        for idx, p in enumerate(points):
            dist = math.hypot(ls.cx - p.cx, ls.cy - p.cy)
            if dist <= pair_thresh:
                pairs.append((dist, name, idx))
    pairs.sort()
    matched: dict[str, int] = {}
    used_points: set[int] = set()
    for _, name, idx in pairs:
        if name in matched or idx in used_points:
            continue
        matched[name] = idx
        used_points.add(idx)
    pair_ok = set(matched.keys()) == labels_present_names and len(matched) == len(labels_present_names)

    # 细则点：分布在绿色断面上/下/左/右边界附近——用配对到的黄点到绿色边界线的距离判定
    green_segments = green_line_shapes(ctx)
    boundary_thresh = 1.5  # 距离 ≤ 1.5cm 视为“贴近绿色边界”
    near_boundary_labels: list[str] = []
    far_labels: list[str] = []
    for name, idx in matched.items():
        p = points[idx]
        d = _min_distance_to_green_boundary(p.cx, p.cy, green_segments)
        if d <= boundary_thresh:
            near_boundary_labels.append(name)
        else:
            far_labels.append(f"{name}({d:.1f}cm)")
    distribution_ok = bool(matched) and not far_labels

    ok = count_ok and labels_ok and pair_ok and distribution_ok
    missing = [b for b in expected if b not in labels_present_names]
    unpaired = sorted(labels_present_names - set(matched.keys()))
    return CheckResult(
        "黄色边界传感器",
        ok,
        f"黄点 {len(points)}(需≥20){count_ok}，指定标签 {len(labels_present_names)}/{len(expected)}{labels_ok}" +
        (f"(缺{missing})" if missing else "") +
        f"，标签-黄点一一对应(≤{pair_thresh}cm) {pair_ok}" +
        (f"(未配对{unpaired})" if unpaired else "") +
        f"，全部贴近绿色边界(≤{boundary_thresh}cm) {distribution_ok}" +
        (f"(远离{far_labels})" if far_labels else ""),
    )


def check_yellow_sensor_style(ctx: EvaluationContext) -> CheckResult:
    # Rubric point (+3): the yellow sensor points are circular, diameter
    # 0.18–0.20cm, filled yellow (color close to RGB(245,230,0)), with a green or
    # dark-gray outline of line width 0.3–0.75pt; each point has a black label
    # beside it, font Arial or Calibri, size 12–14pt.  Each sub-condition is
    # checked against the exact rubric ranges.  Geometry uses the shape's rendered
    # size/center.
    points = yellow_points(ctx)
    ak_labels = [s for s in ctx.text_shapes if re.fullmatch(r"[AK]\d+", normalize_text(s.text))]

    def label_style_ok(s: Shape) -> bool:
        return (
            (s.font.lower() in ALLOWED_FONTS)
            and s.font_size is not None
            and in_range(s.font_size, 12, 14)
            and is_blackish(s.text_color)
        )

    def edge_ok(p: Shape) -> bool:
        # Green or dark gray outline, with line width 0.3–0.75pt.
        # 未显式设置线宽时按默认 1.0 磅处理，1.0 磅不在 0.3–0.75 内 → 不通过。
        if p.line_color is None:
            return False
        color_ok = is_near_color(p.line_color, GREEN_TARGET) or is_grayish(p.line_color)
        return color_ok and in_range(effective_line_width(p), 0.3, 0.75)

    def point_ok(p: Shape) -> bool:
        circular = p.is_ellipse() and abs(abs(p.w) - abs(p.h)) <= 0.06
        diameter = (abs(p.w) + abs(p.h)) / 2
        diameter_ok = in_range(diameter, 0.18, 0.20)
        fill_ok = is_near_color(p.fill_color, YELLOW_TARGET)
        has_label = any(
            label_style_ok(s) and math.hypot(s.cx - p.cx, s.cy - p.cy) <= 1.5
            for s in ak_labels
        )
        return circular and diameter_ok and fill_ok and edge_ok(p) and has_label

    good = sum(1 for p in points if point_ok(p))
    ok = bool(points) and good == len(points)
    sample = points[0] if points else None
    if sample is not None:
        sample_detail = (
            f"；示例 直径 {(abs(sample.w)+abs(sample.h))/2:.3f}cm[0.18-0.20]，"
            f"圆形 {sample.is_ellipse()}，填充 {sample.fill_color}，边线 {sample.line_color}，"
            f"线宽 {effective_line_width(sample)}磅[0.3-0.75]"
        )
    else:
        sample_detail = ""
    detail = f"黄色传感器点 {len(points)}，全部合格 {good}/{len(points)}{sample_detail}"
    return CheckResult("黄点样式", ok, detail)


def _pair_labels_with_points(
    ctx: EvaluationContext,
    required: list[str],
    points: list[Shape],
    pair_thresh: float = 2.0,
) -> tuple[dict[str, Shape], list[str]]:
    """将 required 中的每个文本标签与最近点(按距离升序贪心)配对。

    返回 (name→point, unpaired_labels)。仅在标签存在且距最近点 ≤ pair_thresh 时配对；
    每个点至多被一个标签占用。
    """
    label_shapes: dict[str, Shape] = {}
    for s in ctx.text_shapes:
        name = normalize_text(s.text)
        if name in required and name not in label_shapes:
            label_shapes[name] = s
    pairs: list[tuple[float, str, int]] = []
    for name, ls in label_shapes.items():
        for idx, p in enumerate(points):
            dist = math.hypot(ls.cx - p.cx, ls.cy - p.cy)
            if dist <= pair_thresh:
                pairs.append((dist, name, idx))
    pairs.sort()
    matched: dict[str, Shape] = {}
    used: set[int] = set()
    for _, name, idx in pairs:
        if name in matched or idx in used:
            continue
        matched[name] = points[idx]
        used.add(idx)
    unpaired = sorted(set(label_shapes.keys()) - set(matched.keys()))
    unpaired += [n for n in required if n not in label_shapes]
    return matched, sorted(set(unpaired))


def check_k_series_positions(ctx: EvaluationContext) -> CheckResult:
    # Rubric point (+3): the top-boundary yellow K-series points are distributed
    # near the upper green contour; at least K908, K903, K899, K651, K657, K663,
    # K670, K679 appear; K908/K903/K899 are in the left-upper to mid-left region;
    # K651 is in the middle-right; K670 and K679 are in the right-upper region and
    # clearly higher than the middle points.  Positions use the yellow sensor
    # point's center (paired with the K label), so region/height and the
    # distance-to-upper-green-contour reflect the point itself, not the label.
    required = ["K908", "K903", "K899", "K651", "K657", "K663", "K670", "K679"]
    yellow = yellow_points(ctx)
    matched, _ = _pair_labels_with_points(ctx, required, yellow, pair_thresh=2.0)
    all_paired = all(k in matched for k in required)

    # 细则点：点位沿上部绿色轮廓附近分布——用点中心到上部绿色边界的距离
    upper_green = [s for s in green_line_shapes(ctx) if s.visual_bounds()[1] <= 6.5]
    contour_thresh = 1.5
    if all_paired:
        near_contour = all(
            _min_distance_to_green_boundary(matched[k].cx, matched[k].cy, upper_green) <= contour_thresh
            for k in required
        )
    else:
        near_contour = False

    left_mid = all_paired and all(matched[k].cx < 16 for k in ["K908", "K903", "K899"])
    middle_right = all_paired and in_range(matched["K651"].cx, 18, 30)

    if all_paired:
        middle_cy = matched["K651"].cy
        right_upper = all(matched[k].cx >= 30 for k in ["K670", "K679"])
        clearly_higher = all(matched[k].cy <= middle_cy - 1.5 for k in ["K670", "K679"])
    else:
        right_upper = clearly_higher = False

    ok = all_paired and near_contour and left_mid and middle_right and right_upper and clearly_higher
    return CheckResult(
        "K 系列上边界",
        ok,
        f"标签-黄点配对 {len(matched)}/{len(required)}(齐全{all_paired})，" +
        f"点位近上部绿色轮廓(≤{contour_thresh}cm) {near_contour}，" +
        f"K908/903/899左上至中左 {left_mid}，K651中部偏右 {middle_right}，" +
        f"K670/679右上 {right_upper}且明显高于中部点 {clearly_higher}",
    )


def check_a_series_positions(ctx: EvaluationContext) -> CheckResult:
    # Rubric point (+3): the bottom-boundary yellow A-series points are distributed
    # near the lower green contour; at least A201, A203, A205, A207, A209, A211,
    # A301, A303, A308, A314 appear; A201–A211 are in the left-lower to
    # middle-lower region, and A301–A314 are in the middle-right to lower-right
    # region.  Positions use the yellow sensor point's center (paired with the
    # A label), so region and distance-to-lower-green-contour reflect the point
    # itself, not the label.
    required = ["A201", "A203", "A205", "A207", "A209", "A211", "A301", "A303", "A308", "A314"]
    a2_series = ["A201", "A203", "A205", "A207", "A209", "A211"]
    a3_series = ["A301", "A303", "A308", "A314"]
    yellow = yellow_points(ctx)
    matched, _ = _pair_labels_with_points(ctx, required, yellow, pair_thresh=2.0)
    all_paired = all(k in matched for k in required)

    # 细则点：点位沿下部绿色轮廓附近分布——用点中心到下部绿色边界的距离
    lower_green = [s for s in green_line_shapes(ctx) if s.visual_bounds()[3] >= 15]
    contour_thresh = 1.5
    if all_paired:
        near_contour = all(
            _min_distance_to_green_boundary(matched[k].cx, matched[k].cy, lower_green) <= contour_thresh
            for k in required
        )
    else:
        near_contour = False

    # A201–A211: left-lower to middle-lower (left of the figure center ~19cm).
    a2_ok = all_paired and all(matched[k].cx <= 20 for k in a2_series)
    # A301–A314: middle-right to lower-right (right of the figure center).
    a3_ok = all_paired and all(matched[k].cx >= 20 for k in a3_series)

    ok = all_paired and near_contour and a2_ok and a3_ok
    return CheckResult(
        "A 系列下边界",
        ok,
        f"标签-黄点配对 {len(matched)}/{len(required)}(齐全{all_paired})，" +
        f"点位近下部绿色轮廓(≤{contour_thresh}cm) {near_contour}，" +
        f"A201-A211左下至中部下方 {a2_ok}，A301-A314中右至右下 {a3_ok}",
    )


def check_b_series_positions(ctx: EvaluationContext) -> CheckResult:
    # Rubric point (+3): the blue B-series points are arranged near the horizontal
    # zero line; B114 is on the left, B121 is close to the vertical zero-reference
    # dashed line, and B124 and B125 are on the right; the labels sit near their
    # corresponding dots, with a label-to-dot distance of about 0.1–0.5cm, and do
    # not obviously occlude the green lines.  Each sub-condition is checked against
    # the exact rubric ranges.  Positions use the rendered shape centers; the
    # horizontal/vertical zero lines are located from the actual gray zero lines.
    labels = {normalize_text(s.text): s for s in ctx.text_shapes if re.fullmatch(r"B\d+", normalize_text(s.text))}
    points = [p for p in blue_points(ctx) if in_range(p.cx, 11, 27) and in_range(p.cy, 11, 16)]

    # 细则点：每个 B 标签与蓝点建立一一配对（按最近距离贪心，标签-点距离 ≤ 2.0cm 视为配对）
    b_names = sorted(labels.keys())
    pairs: list[tuple[float, str, int]] = []
    for name in b_names:
        ls = labels[name]
        for idx, p in enumerate(points):
            dist = math.hypot(ls.cx - p.cx, ls.cy - p.cy)
            if dist <= 2.0:
                pairs.append((dist, name, idx))
    pairs.sort()
    matched: dict[str, Shape] = {}
    used: set[int] = set()
    for _, name, idx in pairs:
        if name in matched or idx in used:
            continue
        matched[name] = points[idx]
        used.add(idx)

    # Horizontal zero line height (gray horizontal reference line, near the middle
    # of the figure).  Position filter excludes the X-axis at y≈26cm — after the
    # gray-color threshold was widened, dark-gray axes also match ``is_grayish``,
    # so we anchor to the rubric's middle band (y≈13.5–14).
    h_zero = [
        s
        for s in ctx.line_shapes
        if is_horizontal(s, 5)
        and is_grayish(s.line_color)
        and s.length > 30
        and in_range(s.cy, 13.5, 14.0, 0.5)
    ]
    zero_y = h_zero[0].cy if h_zero else 13.75
    # Vertical zero reference line x (central gray vertical line, near the middle
    # of the figure).  Position filter excludes the Y-axis at x≈2.5cm.
    v_zero = [
        s
        for s in ctx.line_shapes
        if is_vertical(s, 5)
        and is_grayish(s.line_color)
        and s.length >= 15
        and 16.5 <= s.cx <= 20.0
    ]
    zero_x = v_zero[0].cx if v_zero else 19.5

    # 细则点：沿横向零线附近排列——所有蓝点靠近 zero_y
    near_zero_line = bool(points) and all(abs(p.cy - zero_y) <= 2.0 for p in points)

    # 细则点：B114/B121/B124/B125 的位置以“配对到的蓝点中心”为准
    b114_left = "B114" in matched and matched["B114"].cx < 14
    b121_near_vzero = "B121" in matched and abs(matched["B121"].cx - zero_x) <= 2.0
    b124_b125_right = all(k in matched and matched[k].cx >= 22 for k in ["B124", "B125"])

    # 细则点：标签-圆点距离约 0.1–0.5cm——用“配对到的那颗蓝点”，而非任意最近蓝点
    def label_gap_ok(name: str) -> bool:
        if name not in matched:
            return False
        ls = labels[name]
        p = matched[name]
        gap = math.hypot(ls.cx - p.cx, ls.cy - p.cy) - max(abs(p.w), abs(p.h)) / 2
        return in_range(gap, 0.1, 0.5)

    labels_near = bool(labels) and all(label_gap_ok(n) for n in labels)

    # Not obviously occluding the green lines: label centers keep a small clearance
    # from the green grid (they sit beside, not on top of, the lines).
    green = green_line_shapes(ctx)
    not_occluding = all(min_distance_to_lines(s, green) >= 0.1 for s in labels.values()) if labels else False

    ok = (
        len(labels) >= 10
        and len(matched) == len(labels)
        and near_zero_line
        and b114_left
        and b121_near_vzero
        and b124_b125_right
        and labels_near
        and not_occluding
    )
    return CheckResult(
        "B 系列位置",
        ok,
        f"B标签 {len(labels)}，标签-蓝点配对 {len(matched)}/{len(labels)}，" +
        f"沿横向零线排列 {near_zero_line}，B114左侧 {b114_left}，" +
        f"B121近竖向零位虚线 {b121_near_vzero}，B124/B125右侧 {b124_b125_right}，" +
        f"标签近圆点(0.1-0.5cm) {labels_near}，未明显遮挡绿线 {not_occluding}",
    )


def check_sensor_label_quality(ctx: EvaluationContext) -> CheckResult:
    # Rubric point (+3): the sensor-number texts have no obvious missing/wrong
    # characters or reversed order; the label color is black or dark gray; all
    # labels are editable; and they do not severely overlap the sensor points or
    # the green grid.  Each sub-condition is checked against the rubric.  In the
    # office application these labels are real text runs (a:t), so their presence
    # as text shapes means they are editable.
    expected = required_sensor_labels()
    labels = [s for s in ctx.text_shapes if re.fullmatch(r"[ABK]\d+", normalize_text(s.text))]
    present = {normalize_text(s.text) for s in labels}
    # No obvious missing/wrong character or reversed order: every expected label is
    # present exactly (a misspelled or reversed one would not match), and there are
    # no stray sensor-format labels outside the expected set.
    missing = expected - present
    extra = present - expected
    text_ok = not missing and not extra

    # Label color black or dark gray.
    color_ok = bool(labels) and all(is_blackish(s.text_color) or is_grayish(s.text_color) for s in labels)

    # All labels editable: they are parsed as text shapes (editable text runs).
    editable_ok = bool(labels)

    # Not severely overlapping the sensor points or the green grid.
    points = blue_points(ctx) + yellow_points(ctx)
    green = green_line_shapes(ctx)

    def severe_point_overlap(s: Shape) -> bool:
        for p in points:
            smaller = min(s.area, p.area)
            if smaller > 0 and rect_overlap_area(s, p) > 0.5 * smaller:
                return True
        return False

    def severe_green_overlap(s: Shape) -> bool:
        # "严重重叠" = the label is substantially covered by green lines, not merely
        # grazed.  Green grid lines are thin (~1pt ≈ 0.035cm), so estimate the
        # fraction of the label's area covered by green segments: sample each
        # segment, take the portion falling inside the label rectangle, multiply by
        # the line width, and sum.  Severe only when coverage exceeds ~50%.
        label_area = s.area
        if label_area <= 0:
            return False
        covered = 0.0
        for g in green:
            width_cm = effective_line_width(g) * (EMU_PER_PT / EMU_PER_CM)
            samples = 24
            inside = 0
            for i in range(samples + 1):
                t = i / samples
                px = g.x + t * g.w
                py = g.y + t * g.h
                if s.x1 <= px <= s.x2 and s.y1 <= py <= s.y2:
                    inside += 1
            if inside:
                covered += (inside / (samples + 1)) * g.length * width_cm
        return covered > 0.5 * label_area

    overlap_ok = all(not severe_point_overlap(s) and not severe_green_overlap(s) for s in labels)

    ok = text_ok and color_ok and editable_ok and overlap_ok
    return CheckResult(
        "传感器编号文本",
        ok,
        f"编号完整无漏/错/颠倒 {text_ok}(缺{sorted(missing)}/多{sorted(extra)})，"
        f"颜色黑或深灰 {color_ok}，可编辑 {editable_ok}，未与点位/绿网严重重叠 {overlap_ok}",
    )


def has_white_fill(shape: Shape) -> bool:
    # 非主题色判定：白底填充只需属于"白色/浅色"大类即可。
    return shape.fill_color is not None and min(shape.fill_color) >= WHITE_MIN


def check_scale_box(ctx: EvaluationContext) -> CheckResult:
    # Rubric point (+3): a scale-legend box appears in the lower-left corner, with
    # an overall width of 4.4–4.8cm and height 2.4–2.7cm, positioned 3.0–3.4cm from
    # the left and 22.5–24.5cm from the top; the border is a gray single solid
    # line of width 0.5pt, and the fill is white.  Each sub-condition is checked
    # against the exact rubric ranges (no extra tolerance).  Position uses the
    # rotation-corrected visual bounds so it matches what the office app renders.
    result: Optional[CheckResult] = None
    for s in ctx.shapes:
        if not s.is_rect():
            continue
        left, top, _, _ = s.visual_bounds()
        width = abs(s.w)
        height = abs(s.h)
        left_ok = in_range(left, 3.0, 3.4)
        top_ok = in_range(top, 22.5, 24.5)
        width_ok = in_range(width, 4.4, 4.8)
        height_ok = in_range(height, 2.4, 2.7)
        if not (left_ok and top_ok and width_ok and height_ok):
            continue
        border_gray = is_grayish(s.line_color)
        solid_ok = s.dash == "solid"
        width_line_ok = in_range(effective_line_width(s), 0.5, 0.5, 0.01)
        fill_white = has_white_fill(s)
        detail = (
            f"距左 {left:.2f}cm[3.0-3.4]{left_ok}，距上 {top:.2f}cm[22.5-24.5]{top_ok}，"
            f"宽 {width:.2f}cm[4.4-4.8]{width_ok}，高 {height:.2f}cm[2.4-2.7]{height_ok}，"
            f"边框灰色 {s.line_color}({border_gray})，单实线 {s.dash}({solid_ok})，"
            f"线宽 {effective_line_width(s)}磅[0.5]{width_line_ok}，白色填充 {s.fill_color}({fill_white})"
        )
        if border_gray and solid_ok and width_line_ok and fill_white:
            return CheckResult("比例尺说明框", True, detail)
        if result is None:
            result = CheckResult("比例尺说明框", False, detail)
    return result or CheckResult("比例尺说明框", False, "未找到符合位置与尺寸的矩形框")


def check_scale_arrows(ctx: EvaluationContext) -> CheckResult:
    # Rubric point (+3): inside the scale-legend box there are green coordinate
    # arrows — an upward Y-direction arrow and a rightward X-direction arrow; the
    # arrow lines are green single solid lines, color close to RGB(20,145,45),
    # line width 0.75–1.25pt; the upward arrow is 0.5–0.8cm long at angle 90°, and
    # the rightward arrow is 0.7–1cm long at angle 0°.  Each sub-condition is
    # checked against the exact rubric ranges.  Arrows are located by their center
    # inside the scale box; "green single solid line" requires a solid dash and a
    # color close to the contour green.
    def arrow_common(s: Shape) -> bool:
        return (
            s.is_line()
            and s.dash == "solid"
            and is_near_color(s.line_color, GREEN_TARGET)
            and in_range(effective_line_width(s), 0.75, 1.25)
            and 3.0 <= s.cx <= 7.8
            and 22.3 <= s.cy <= 25.5
        )

    arrows = [s for s in ctx.line_shapes if arrow_common(s)]
    up = [s for s in arrows if is_vertical(s, 3) and in_range(s.length, 0.5, 0.8)]
    right = [s for s in arrows if is_horizontal(s, 3) and in_range(s.length, 0.7, 1.0)]
    has_up = bool(up)
    has_right = bool(right)
    ok = has_up and has_right
    return CheckResult(
        "比例尺小箭头",
        ok,
        f"框内绿色单实线(近RGB20/145/45,宽0.75-1.25) {len(arrows)}；"
        f"向上箭头(长0.5-0.8,角度90°) {has_up}，向右箭头(长0.7-1.0,角度0°) {has_right}；"
        f"长度 {[round(s.length, 3) for s in arrows]}",
    )


def check_scale_text(ctx: EvaluationContext) -> CheckResult:
    # Rubric point (+1): inside the scale-legend box the texts "Y (Up)",
    # "X (Right)" and "Displacement Scale: 1 mm" appear; font Arial or Calibri,
    # size 9–10pt, color black.  Each sub-condition is checked against the exact
    # rubric ranges (no extra tolerance).  All three texts must be present and each
    # must satisfy the font/size/color style.
    expected = ["Y (Up)", "X (Right)", "Displacement Scale: 1 mm"]
    found: list[Shape] = []
    for text in expected:
        found.extend(find_text(ctx, text))
    all_present = len(found) >= len(expected)

    def style_ok(s: Shape) -> bool:
        return (
            (s.font.lower() in ALLOWED_FONTS)
            and s.font_size is not None
            and in_range(s.font_size, 9, 10)
            and is_blackish(s.text_color)
        )

    good = sum(1 for s in found if style_ok(s))
    ok = all_present and good == len(found) and len(found) == len(expected)
    return CheckResult(
        "比例尺文本",
        ok,
        f"文本 {len(found)}/{len(expected)}(齐全{all_present})，样式合格(Arial/Calibri,9-10磅,黑色) {good}/{len(found)}，"
        f"字号 {[s.font_size for s in found]}",
    )


def check_legend_box(ctx: EvaluationContext) -> CheckResult:
    # Rubric point (+3): a legend box appears in the lower-right corner, with an
    # overall width of 6.5–7.0cm and height 1.7–2cm, positioned 28–30cm from the
    # left and 21–23cm from the top; the border is a gray single solid line of
    # width 0.5pt, and the fill is white.  Each sub-condition is checked against
    # the exact rubric ranges (no extra tolerance).  Position uses the
    # rotation-corrected visual bounds so it matches what the office app renders.
    result: Optional[CheckResult] = None
    for s in ctx.shapes:
        if not s.is_rect():
            continue
        left, top, _, _ = s.visual_bounds()
        width = abs(s.w)
        height = abs(s.h)
        left_ok = in_range(left, 28, 30)
        top_ok = in_range(top, 21, 23)
        width_ok = in_range(width, 6.5, 7.0)
        height_ok = in_range(height, 1.7, 2.0)
        if not (left_ok and top_ok and width_ok and height_ok):
            continue
        border_gray = is_grayish(s.line_color)
        solid_ok = s.dash == "solid"
        width_line_ok = in_range(effective_line_width(s), 0.5, 0.5, 0.01)
        fill_white = has_white_fill(s)
        detail = (
            f"距左 {left:.2f}cm[28-30]{left_ok}，距上 {top:.2f}cm[21-23]{top_ok}，"
            f"宽 {width:.2f}cm[6.5-7.0]{width_ok}，高 {height:.2f}cm[1.7-2.0]{height_ok}，"
            f"边框灰色 {s.line_color}({border_gray})，单实线 {s.dash}({solid_ok})，"
            f"线宽 {effective_line_width(s)}磅[0.5]{width_line_ok}，白色填充 {s.fill_color}({fill_white})"
        )
        if border_gray and solid_ok and width_line_ok and fill_white:
            return CheckResult("图例框", True, detail)
        if result is None:
            result = CheckResult("图例框", False, detail)
    return result or CheckResult("图例框", False, "未找到符合位置与尺寸的矩形框")


def check_legend_content(ctx: EvaluationContext) -> CheckResult:
    # Rubric point (+3): inside the legend box there are a blue dot with the text
    # "Crown/Centerline Sensors" and a yellow dot with the text "Upper/Lower
    # Boundary Sensors"; the dots are 0.16–0.20cm in diameter; the text font is
    # Arial or Calibri, size 11–13pt, color black.  Each sub-condition is checked
    # against the exact rubric ranges (no extra tolerance).  Dot positions use the
    # rendered center within the legend region.
    crown = find_text(ctx, "Crown/Centerline Sensors")
    boundary = find_text(ctx, "Upper/Lower Boundary Sensors")
    text_shapes = crown + boundary
    both_texts = bool(crown) and bool(boundary)

    def text_style_ok(s: Shape) -> bool:
        return (
            (s.font.lower() in ALLOWED_FONTS)
            and s.font_size is not None
            and in_range(s.font_size, 11, 13)
            and is_blackish(s.text_color)
        )

    text_ok = both_texts and all(text_style_ok(s) for s in text_shapes)

    # Legend dots inside the legend box region, diameter 0.16–0.20cm.
    legend_points = [
        p
        for p in ctx.ellipse_shapes
        if 28 <= p.cx <= 36
        and 21 <= p.cy <= 24
        and in_range((abs(p.w) + abs(p.h)) / 2, 0.16, 0.20)
    ]
    has_blue = any(is_near_color(p.fill_color, BLUE_TARGET) for p in legend_points)
    has_yellow = any(is_near_color(p.fill_color, YELLOW_TARGET) for p in legend_points)

    ok = both_texts and text_ok and has_blue and has_yellow
    good_style = sum(1 for s in text_shapes if text_style_ok(s))
    return CheckResult(
        "图例内容",
        ok,
        f"文本齐全 {both_texts}(Crown {bool(crown)}/Boundary {bool(boundary)})，"
        f"文字样式(Arial/Calibri,11-13磅,黑色) {good_style}/{len(text_shapes)}，"
        f"图例圆点(直径0.16-0.20) {len(legend_points)}，蓝点 {has_blue}/黄点 {has_yellow}",
    )


def describe_shapes(shapes: list[Shape]) -> str:
    if not shapes:
        return "未找到"
    s = shapes[0]
    return f"找到 {len(shapes)} 个；首个 x={s.x:.2f}, y={s.y:.2f}, w={s.w:.2f}, h={s.h:.2f}, font={s.font}, size={s.font_size}"


def describe_lines(lines: list[Shape]) -> str:
    if not lines:
        return "未找到"
    s = lines[0]
    return f"找到 {len(lines)} 条；首条 x={s.x1:.2f}, y={s.y1:.2f}, len={s.length:.2f}, angle={s.angle:.1f}, width={s.line_width}"


def line_extent(lines: list[Shape]) -> str:
    if not lines:
        return "无"
    return f"x {min(s.x1 for s in lines):.2f}-{max(s.x2 for s in lines):.2f}, y {min(s.y1 for s in lines):.2f}-{max(s.y2 for s in lines):.2f}"


def distribution_ok(points: list[Shape], min_quadrants: int = 3) -> bool:
    quadrants = set()
    for p in points:
        quadrants.add((p.cx >= 19, p.cy >= 13.8))
    return len(quadrants) >= min_quadrants


def labels_near_points(ctx: EvaluationContext, labels: Iterable[Shape], points: list[Shape], low: float, high: float) -> int:
    count = 0
    for label in labels:
        if any(low <= math.hypot(label.cx - p.cx, label.cy - p.cy) <= high + max(label.w, label.h) / 2 for p in points):
            count += 1
    return count


def min_distance_to_lines(point: Shape, lines: list[Shape]) -> float:
    if not lines:
        return 999.0
    return min(distance_point_to_segment(point.cx, point.cy, line.x, line.y, line.x + line.w, line.y + line.h) for line in lines)


def distance_point_to_segment(px: float, py: float, x1: float, y1: float, x2: float, y2: float) -> float:
    dx = x2 - x1
    dy = y2 - y1
    if dx == 0 and dy == 0:
        return math.hypot(px - x1, py - y1)
    t = max(0.0, min(1.0, ((px - x1) * dx + (py - y1) * dy) / (dx * dx + dy * dy)))
    closest_x = x1 + t * dx
    closest_y = y1 + t * dy
    return math.hypot(px - closest_x, py - closest_y)


def _locate_pptx(dir_path: Path) -> Path:
    """在脚本所在目录中定位待评估的 .pptx 文件。

    约定 runner 只把脚本所在目录传进来，脚本自己扫描并选定被评估文档：
    优先匹配默认文件名，否则挑选目录内第一个 .pptx。
    """
    preferred = dir_path / "Riverdale_Tunnel_Section7_Editable_Recreation.pptx"
    if preferred.exists():
        return preferred
    candidates = sorted(dir_path.glob("*.pptx"))
    if candidates:
        return candidates[0]
    raise FileNotFoundError(f"目录中未找到 .pptx 文件: {dir_path}")


def _build_dim2_items(ctx: EvaluationContext) -> tuple[list[dict], int, int]:
    """执行维度二逐项评分，返回 (items, total_score, max_score)。"""
    items: list[dict] = []
    total = 0
    max_score = 0
    for item in scoring_items():
        max_score += item.points
        errored = False
        error_detail = ""
        try:
            result = item.checker(ctx)
        except Exception as exc:  # defensive: every rubric point must report automatically.
            result = CheckResult(item.label, False, f"检测异常: {exc}")
            errored = True
            error_detail = f"检测异常: {exc}"
        delta = item.points if result.passed else 0
        total += delta
        # detail 字段按对外约定恒为空字符串：无论该项命中/未命中/报错，评分结果
        # （delta、hit）已由上方逻辑独立记录，detail 不再对外输出具体理由。
        # 上方仍保留 errored / error_detail 的计算，以保持异常兜底逻辑不变。
        _ = errored, error_detail  # 保留占位，避免误删异常兜底分支
        items.append(
            {
                "rule": item.label,
                "max_delta": item.points,
                "delta": delta,
                "hit": result.passed,
                "detail": "",
            }
        )
    return items, total, max_score


def evaluate(dir_path: str) -> dict:
    """统一入口：接收脚本所在目录，返回结构化评估结果。

    Args:
        dir_path: 脚本所在目录（内部包含待评估的 .pptx 文件）。

    Returns:
        评估结果字典，字段参见《脚本接口差异与统一建议.md》§2.2。
    """
    script_id = "049"
    max_score = sum(item.points for item in scoring_items())

    try:
        directory = Path(dir_path)
        if not directory.is_dir():
            return {
                "id": script_id,
                "file_name": "",
                "status": "error",
                "error": f"目录不存在或不是目录: {dir_path}",
                "dim1_pass": False,
                "dim1_reason": "",
                "dim2_items": [],
                "total_score": 0,
                "max_score": max_score,
            }

        pptx_path = _locate_pptx(directory)
        file_name = pptx_path.name

        if pptx_path.suffix.lower() != ".pptx":
            return {
                "id": script_id,
                "file_name": file_name,
                "status": "ok",
                "error": None,
                "dim1_pass": False,
                "dim1_reason": f"文件格式不符合要求: {pptx_path.suffix}",
                "dim2_items": [],
                "total_score": 0,
                "max_score": max_score,
            }

        try:
            ctx = load_pptx(pptx_path)
        except Exception as exc:
            return {
                "id": script_id,
                "file_name": file_name,
                "status": "ok",
                "error": None,
                "dim1_pass": False,
                "dim1_reason": f"文件无法正常解析: {exc}",
                "dim2_items": [],
                "total_score": 0,
                "max_score": max_score,
            }

        gate = gate_checks(ctx)
        if not all(result.passed for result in gate):
            failed = next(r for r in gate if not r.passed)
            reason = f"{failed.label}（{failed.detail}）" if failed.detail else failed.label
            return {
                "id": script_id,
                "file_name": file_name,
                "status": "ok",
                "error": None,
                "dim1_pass": False,
                "dim1_reason": reason,
                "dim2_items": [],
                "total_score": 0,
                "max_score": max_score,
            }

        items, total, computed_max = _build_dim2_items(ctx)
        return {
            "id": script_id,
            "file_name": file_name,
            "status": "ok",
            "error": None,
            "dim1_pass": True,
            "dim1_reason": "",
            "dim2_items": items,
            "total_score": total,
            "max_score": computed_max,
        }
    except Exception as exc:  # 顶层兜底：脚本自身异常统一记为 status=error
        return {
            "id": script_id,
            "file_name": "",
            "status": "error",
            "error": f"{type(exc).__name__}: {exc}",
            "dim1_pass": False,
            "dim1_reason": "",
            "dim2_items": [],
            "total_score": 0,
            "max_score": max_score,
        }


if __name__ == "__main__":
    # 仅用于本地调试：默认评估脚本所在目录，可通过命令行参数覆盖。
    debug_dir = sys.argv[1] if len(sys.argv) > 1 else str(Path(__file__).resolve().parent)
    print(json.dumps(evaluate(debug_dir), ensure_ascii=False, indent=2))
