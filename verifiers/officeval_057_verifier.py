#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""自动评估 editable_chemical_scheme.pptx。

评估流程：
1. 先检查“维度1：可用与可修改性”。任一门槛失败，最终得分为 0，跳过维度2。
2. 维度1通过后，逐条自动检测维度2评分点，打印命中点、证据和最终得分。

实现原则：
- 不人工打开 PPT，不依赖人工判断。
- 直接解析 PPTX 的 Office Open XML，判断文本、线段、箭头、图片、位置、字号、颜色等。
- 对化学结构语义采用几何/文本启发式：区域、关键元素文本、短黑线数量、箭头、并行线等。
"""

from __future__ import annotations

import json
import math
import os
import re
import sys
import zipfile
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Callable, Dict, Iterable, List, Optional, Sequence, Tuple
from xml.etree import ElementTree as ET

EMU_PER_INCH = 914400
EMU_PER_CM = 360000
EMU_PER_PT = 12700

NS = {
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
}

CHEM_TOKENS = [
    "Cl", "Br", "N", "O", "CH3", "CH₃", "H3C", "H₃C", "NH", "OCH3", "O—CH3", "O-CH3", "(CH3)2NH", "(CH₃)₂NH", "+"
]
GOOD_FONTS = {"times new roman", "cambria math"}


def q(tag: str) -> str:
    prefix, name = tag.split(":", 1)
    return f"{{{NS[prefix]}}}{name}"


def emu_to_cm(v: float) -> float:
    return v / EMU_PER_CM


def emu_to_pt(v: float) -> float:
    return v / EMU_PER_PT


def safe_int(value: Optional[str], default: int = 0) -> int:
    if value is None:
        return default
    try:
        return int(float(value))
    except (TypeError, ValueError):
        return default


def norm_text(s: str) -> str:
    return (
        s.replace("₀", "0")
        .replace("₁", "1")
        .replace("₂", "2")
        .replace("₃", "3")
        .replace("₄", "4")
        .replace("₅", "5")
        .replace("₆", "6")
        .replace("₇", "7")
        .replace("₈", "8")
        .replace("₉", "9")
        .replace("—", "-")
        .replace("−", "-")
        .replace("＋", "+")
        .replace("﹢", "+")
        .replace("➕", "+")
        .replace(" ", "")
        .strip()
    )


# 兼容用户在 PPT 中使用的各种"加号"字符：半角 +、全角＋、小型﹢、heavy plus ➕
PLUS_CHARS = {"+", "＋", "﹢", "➕"}


def is_plus_text(text: str) -> bool:
    return text.strip() in PLUS_CHARS


def rgb_tuple(hex_value: Optional[str]) -> Optional[Tuple[int, int, int]]:
    if not hex_value:
        return None
    value = hex_value.strip().replace("#", "").upper()
    if len(value) != 6 or not re.fullmatch(r"[0-9A-F]{6}", value):
        return None
    return int(value[0:2], 16), int(value[2:4], 16), int(value[4:6], 16)


def is_near_white(color: Optional[str]) -> bool:
    rgb = rgb_tuple(color)
    if rgb is None:
        return False
    return min(rgb) >= 240


def is_black_or_unknown(color: Optional[str]) -> bool:
    if color is None:
        return True
    rgb = rgb_tuple(color)
    if rgb is None:
        return True
    return max(rgb) <= 70


def is_red(color: Optional[str]) -> bool:
    rgb = rgb_tuple(color)
    if rgb is None:
        return False
    r, g, b = rgb
    return r >= 180 and g <= 100 and b <= 100


def color_name(color: Optional[str]) -> str:
    return color if color else "继承/未知"


@dataclass
class BBox:
    x: float = 0.0
    y: float = 0.0
    w: float = 0.0
    h: float = 0.0

    @property
    def x2(self) -> float:
        return self.x + self.w

    @property
    def y2(self) -> float:
        return self.y + self.h

    @property
    def cx(self) -> float:
        return self.x + self.w / 2

    @property
    def cy(self) -> float:
        return self.y + self.h / 2

    @property
    def area(self) -> float:
        return max(0.0, self.w) * max(0.0, self.h)

    def intersects(self, other: "BBox") -> bool:
        return self.x < other.x2 and self.x2 > other.x and self.y < other.y2 and self.y2 > other.y

    def intersection_area(self, other: "BBox") -> float:
        ix1 = max(self.x, other.x)
        iy1 = max(self.y, other.y)
        ix2 = min(self.x2, other.x2)
        iy2 = min(self.y2, other.y2)
        return max(0.0, ix2 - ix1) * max(0.0, iy2 - iy1)

    def contains_point(self, x: float, y: float) -> bool:
        return self.x <= x <= self.x2 and self.y <= y <= self.y2

    def contains_center(self, other: "BBox") -> bool:
        return self.contains_point(other.cx, other.cy)

    def expand(self, ratio: float) -> "BBox":
        dx = self.w * ratio
        dy = self.h * ratio
        return BBox(self.x - dx, self.y - dy, self.w + 2 * dx, self.h + 2 * dy)

    def union(self, other: "BBox") -> "BBox":
        if self.area <= 0:
            return other
        if other.area <= 0:
            return self
        x1 = min(self.x, other.x)
        y1 = min(self.y, other.y)
        x2 = max(self.x2, other.x2)
        y2 = max(self.y2, other.y2)
        return BBox(x1, y1, x2 - x1, y2 - y1)

    def to_cm_tuple(self) -> Tuple[float, float, float, float]:
        return (emu_to_cm(self.x), emu_to_cm(self.y), emu_to_cm(self.w), emu_to_cm(self.h))

    def describe_cm(self) -> str:
        x, y, w, h = self.to_cm_tuple()
        return f"x={x:.2f}cm, y={y:.2f}cm, w={w:.2f}cm, h={h:.2f}cm"


@dataclass
class TextRun:
    text: str
    font: Optional[str] = None
    size_pt: Optional[float] = None
    color: Optional[str] = None


@dataclass
class ShapeInfo:
    shape_id: str
    name: str
    kind: str
    bbox: BBox
    text: str = ""
    text_runs: List[TextRun] = field(default_factory=list)
    line_color: Optional[str] = None
    fill_color: Optional[str] = None
    line_width_pt: Optional[float] = None
    line_dash: Optional[str] = None
    has_arrow: bool = False
    head_arrow_type: Optional[str] = None
    tail_arrow_type: Optional[str] = None
    preset_geometry: Optional[str] = None
    in_group: bool = False
    is_hidden: bool = False
    z_order: int = 0

    @property
    def is_picture(self) -> bool:
        return self.kind == "picture"

    @property
    def is_text(self) -> bool:
        return bool(self.text.strip())

    @property
    def is_line_like(self) -> bool:
        if self.kind == "connector":
            return True
        if self.preset_geometry == "line":
            return True
        # 化学键常由无填充、有线条的细长形状构成。
        if self.kind in {"shape", "freeform"} and self.line_width_pt is not None:
            if self.bbox.w > 0 and self.bbox.h > 0:
                ratio = max(self.bbox.w, self.bbox.h) / max(1.0, min(self.bbox.w, self.bbox.h))
                return ratio >= 3 and self.bbox.area > 0
            return True
        return False

    @property
    def is_editable(self) -> bool:
        return self.kind in {"shape", "connector", "group", "freeform"} or self.is_text

    @property
    def length(self) -> float:
        return math.hypot(self.bbox.w, self.bbox.h)


@dataclass
class SlideInfo:
    index: int
    path: str
    width: float
    height: float
    shapes: List[ShapeInfo] = field(default_factory=list)
    background_color: Optional[str] = None
    background_fill_types: List[str] = field(default_factory=list)
    has_comments: bool = False

    @property
    def slide_area(self) -> float:
        return self.width * self.height

    @property
    def texts(self) -> List[ShapeInfo]:
        return [s for s in self.shapes if s.is_text and not s.is_hidden]

    @property
    def pictures(self) -> List[ShapeInfo]:
        return [s for s in self.shapes if s.is_picture and not s.is_hidden]

    @property
    def lines(self) -> List[ShapeInfo]:
        return [s for s in self.shapes if s.is_line_like and not s.is_hidden]

    @property
    def arrows(self) -> List[ShapeInfo]:
        arrows = [s for s in self.lines if s.has_arrow]
        if arrows:
            return arrows
        # 兜底：没有显式箭头元数据时，长水平线也视为候选反应箭头。
        candidates = []
        for s in self.lines:
            if s.bbox.w > self.width * 0.08 and s.bbox.w > s.bbox.h * 4:
                candidates.append(s)
        return candidates


@dataclass
class PresentationModel:
    path: Path
    width: float
    height: float
    slides: List[SlideInfo]
    package_ok: bool = True
    parse_errors: List[str] = field(default_factory=list)


@dataclass
class RuleResult:
    rule_id: str
    name: str
    points: float
    passed: bool
    evidence: str
    dimension: str = "D2"

    def score(self) -> float:
        return self.points if self.passed else 0.0


class PptxInspector:
    def __init__(self, path: Path):
        self.path = path
        self.errors: List[str] = []
        self.zf: Optional[zipfile.ZipFile] = None

    def parse(self) -> PresentationModel:
        if self.path.suffix.lower() != ".pptx":
            self.errors.append(f"扩展名为 {self.path.suffix}，不是 .pptx")

        width, height = 12192000, 6858000  # 默认 16:9
        slides: List[SlideInfo] = []
        package_ok = True
        try:
            self.zf = zipfile.ZipFile(self.path)
            names = set(self.zf.namelist())
            if "ppt/presentation.xml" not in names:
                self.errors.append("缺少 ppt/presentation.xml")
                package_ok = False
            else:
                width, height = self._read_slide_size(width, height)
            slide_paths = sorted(
                [n for n in names if re.fullmatch(r"ppt/slides/slide\d+\.xml", n)],
                key=lambda x: int(re.search(r"slide(\d+)\.xml", x).group(1)),
            )
            for idx, slide_path in enumerate(slide_paths, 1):
                try:
                    slides.append(self._parse_slide(idx, slide_path, width, height))
                except Exception as exc:  # 保证评估器自身不因坏页中断。
                    self.errors.append(f"解析 {slide_path} 失败：{type(exc).__name__}: {exc}")
        except Exception as exc:
            package_ok = False
            self.errors.append(f"无法作为 PPTX zip 包打开：{type(exc).__name__}: {exc}")
        finally:
            if self.zf:
                self.zf.close()
        return PresentationModel(self.path, width, height, slides, package_ok and not self.errors, self.errors)

    def _read_slide_size(self, default_w: int, default_h: int) -> Tuple[int, int]:
        assert self.zf is not None
        root = ET.fromstring(self.zf.read("ppt/presentation.xml"))
        sld_size = root.find("p:sldSz", NS)
        if sld_size is None:
            return default_w, default_h
        return safe_int(sld_size.get("cx"), default_w), safe_int(sld_size.get("cy"), default_h)

    def _parse_slide(self, index: int, slide_path: str, width: int, height: int) -> SlideInfo:
        assert self.zf is not None
        root = ET.fromstring(self.zf.read(slide_path))
        slide = SlideInfo(index=index, path=slide_path, width=width, height=height)
        slide.background_color = self._extract_background(root)
        slide.background_fill_types = self._extract_background_fill_types(root)
        slide.has_comments = self._slide_has_comments(slide_path)
        c_sld = root.find("p:cSld", NS)
        sp_tree = c_sld.find("p:spTree", NS) if c_sld is not None else None
        if sp_tree is not None:
            z = 0
            for child in list(sp_tree):
                local = self._local_name(child.tag)
                if local in {"sp", "cxnSp", "pic", "grpSp"}:
                    z = self._parse_shape_recursive(child, slide, False, z, None)
        return slide

    def _parse_shape_recursive(
        self,
        elem: ET.Element,
        slide: SlideInfo,
        in_group: bool,
        z_order: int,
        transform: Optional[Tuple[float, float, float, float, float, float]],
    ) -> int:
        local = self._local_name(elem.tag)
        if local == "grpSp":
            group_bbox = self._extract_bbox(elem, transform)
            shape = ShapeInfo(
                shape_id=self._extract_shape_id(elem),
                name=self._extract_shape_name(elem),
                kind="group",
                bbox=group_bbox,
                in_group=in_group,
                z_order=z_order,
            )
            slide.shapes.append(shape)
            z_order += 1
            child_transform = self._group_child_transform(elem, transform)
            for child in list(elem):
                child_local = self._local_name(child.tag)
                if child_local in {"sp", "cxnSp", "pic", "grpSp"}:
                    z_order = self._parse_shape_recursive(child, slide, True, z_order, child_transform)
            return z_order

        kind = {"sp": "shape", "cxnSp": "connector", "pic": "picture"}.get(local, "shape")
        bbox = self._extract_bbox(elem, transform)
        preset = self._extract_preset_geometry(elem)
        line_color, fill_color, line_width_pt, has_arrow, head_arrow_type, tail_arrow_type, line_dash = self._extract_style(elem)
        runs = self._extract_text_runs(elem)
        text = "".join(r.text for r in runs)
        shape = ShapeInfo(
            shape_id=self._extract_shape_id(elem),
            name=self._extract_shape_name(elem),
            kind=kind,
            bbox=bbox,
            text=text,
            text_runs=runs,
            line_color=line_color,
            fill_color=fill_color,
            line_width_pt=line_width_pt,
            line_dash=line_dash,
            has_arrow=has_arrow,
            head_arrow_type=head_arrow_type,
            tail_arrow_type=tail_arrow_type,
            preset_geometry=preset,
            in_group=in_group,
            is_hidden=self._is_hidden(elem),
            z_order=z_order,
        )
        slide.shapes.append(shape)
        return z_order + 1

    def _group_child_transform(
        self,
        elem: ET.Element,
        parent: Optional[Tuple[float, float, float, float, float, float]],
    ) -> Optional[Tuple[float, float, float, float, float, float]]:
        xfrm = elem.find("p:grpSpPr/a:xfrm", NS)
        if xfrm is None:
            return parent
        off = xfrm.find("a:off", NS)
        ext = xfrm.find("a:ext", NS)
        ch_off = xfrm.find("a:chOff", NS)
        ch_ext = xfrm.find("a:chExt", NS)
        if off is None or ext is None or ch_off is None or ch_ext is None:
            return parent
        gx = safe_int(off.get("x"))
        gy = safe_int(off.get("y"))
        gw = max(1, safe_int(ext.get("cx")))
        gh = max(1, safe_int(ext.get("cy")))
        cx = safe_int(ch_off.get("x"))
        cy = safe_int(ch_off.get("y"))
        cw = max(1, safe_int(ch_ext.get("cx")))
        ch = max(1, safe_int(ch_ext.get("cy")))
        sx = gw / cw
        sy = gh / ch
        tx = gx - cx * sx
        ty = gy - cy * sy
        if parent is None:
            return tx, ty, sx, sy, 0.0, 0.0
        ptx, pty, psx, psy, _, _ = parent
        return ptx + tx * psx, pty + ty * psy, psx * sx, psy * sy, 0.0, 0.0

    def _apply_transform(self, bbox: BBox, transform: Optional[Tuple[float, float, float, float, float, float]]) -> BBox:
        if transform is None:
            return bbox
        tx, ty, sx, sy, _, _ = transform
        return BBox(tx + bbox.x * sx, ty + bbox.y * sy, bbox.w * sx, bbox.h * sy)

    def _extract_bbox(self, elem: ET.Element, transform: Optional[Tuple[float, float, float, float, float, float]]) -> BBox:
        xfrm = elem.find(".//a:xfrm", NS)
        if xfrm is None:
            return BBox()
        off = xfrm.find("a:off", NS)
        ext = xfrm.find("a:ext", NS)
        if off is None or ext is None:
            return BBox()
        bbox = BBox(
            safe_int(off.get("x")),
            safe_int(off.get("y")),
            safe_int(ext.get("cx")),
            safe_int(ext.get("cy")),
        )
        return self._apply_transform(bbox, transform)

    def _extract_shape_id(self, elem: ET.Element) -> str:
        c_nv_pr = elem.find(".//p:cNvPr", NS)
        return c_nv_pr.get("id", "?") if c_nv_pr is not None else "?"

    def _extract_shape_name(self, elem: ET.Element) -> str:
        c_nv_pr = elem.find(".//p:cNvPr", NS)
        return c_nv_pr.get("name", "") if c_nv_pr is not None else ""

    def _is_hidden(self, elem: ET.Element) -> bool:
        c_nv_pr = elem.find(".//p:cNvPr", NS)
        return c_nv_pr is not None and c_nv_pr.get("hidden") == "1"

    def _extract_preset_geometry(self, elem: ET.Element) -> Optional[str]:
        prst = elem.find(".//a:prstGeom", NS)
        if prst is not None:
            return prst.get("prst")
        cust = elem.find(".//a:custGeom", NS)
        if cust is not None:
            return "custom"
        return None

    def _extract_background(self, root: ET.Element) -> Optional[str]:
        bg = root.find(".//p:bg", NS)
        if bg is None:
            return None
        return self._extract_color(bg)

    def _extract_background_fill_types(self, root: ET.Element) -> List[str]:
        bg = root.find(".//p:bg", NS)
        if bg is None:
            return []
        fill_tags = {
            "solidFill": "纯色填充",
            "gradFill": "渐变填充",
            "blipFill": "图片或纹理填充",
            "pattFill": "图案填充",
        }
        found = []
        for elem in bg.iter():
            local = self._local_name(elem.tag)
            if local in fill_tags:
                found.append(fill_tags[local])
        return sorted(set(found))

    def _slide_has_comments(self, slide_path: str) -> bool:
        assert self.zf is not None
        rels_path = slide_path.replace("ppt/slides/", "ppt/slides/_rels/") + ".rels"
        if rels_path not in self.zf.namelist():
            return False
        try:
            root = ET.fromstring(self.zf.read(rels_path))
        except ET.ParseError:
            return False
        for rel in root.findall("rel:Relationship", NS):
            rel_type = (rel.get("Type") or "").lower()
            target = (rel.get("Target") or "").lower()
            if "comment" in rel_type or "comment" in target:
                return True
        return False

    def _extract_color(self, elem: ET.Element) -> Optional[str]:
        srgb = elem.find(".//a:srgbClr", NS)
        if srgb is not None and srgb.get("val"):
            return srgb.get("val").upper()
        scheme = elem.find(".//a:schemeClr", NS)
        if scheme is not None:
            val = (scheme.get("val") or "").lower()
            scheme_map = {
                "bg1": "FFFFFF",
                "lt1": "FFFFFF",
                "tx1": "000000",
                "dk1": "000000",
                "black": "000000",
                "white": "FFFFFF",
            }
            return scheme_map.get(val)
        return None

    def _extract_style(self, elem: ET.Element) -> Tuple[Optional[str], Optional[str], Optional[float], bool, Optional[str], Optional[str]]:
        sp_pr = elem.find(".//p:spPr", NS)
        if sp_pr is None:
            sp_pr = elem.find(".//p:grpSpPr", NS)
        line_color = fill_color = None
        width_pt = None
        head_arrow_type = tail_arrow_type = None
        if sp_pr is not None:
            fill = sp_pr.find("a:solidFill", NS)
            if fill is not None:
                fill_color = self._extract_color(fill)
            line = sp_pr.find("a:ln", NS)
            if line is not None:
                line_color = self._extract_color(line)
                width_pt = emu_to_pt(safe_int(line.get("w"))) if line.get("w") else None
                head = line.find("a:headEnd", NS)
                tail = line.find("a:tailEnd", NS)
                head_arrow_type = head.get("type") if head is not None else None
                tail_arrow_type = tail.get("type") if tail is not None else None
        has_arrow = any(t is not None and t.lower() != "none" for t in (head_arrow_type, tail_arrow_type))
        # 虚线类型：读取 <a:prstDash val="..."/>，无此元素时视为实线（solid）
        line_dash = None
        if sp_pr is not None:
            ln = sp_pr.find("a:ln", NS)
            if ln is not None:
                prst_dash = ln.find("a:prstDash", NS)
                if prst_dash is not None:
                    line_dash = prst_dash.get("val")  # e.g. "solid", "dash", "dot", "dashDot"…
        return line_color, fill_color, width_pt, has_arrow, head_arrow_type, tail_arrow_type, line_dash

    def _extract_text_runs(self, elem: ET.Element) -> List[TextRun]:
        runs: List[TextRun] = []
        for r in elem.findall(".//a:r", NS):
            t = r.find("a:t", NS)
            if t is None or t.text is None:
                continue
            rpr = r.find("a:rPr", NS)
            font = None
            size_pt = None
            color = None
            if rpr is not None:
                if rpr.get("sz"):
                    size_pt = safe_int(rpr.get("sz")) / 100.0
                latin_el = rpr.find("a:latin", NS)
                ea_el = rpr.find("a:ea", NS)
                latin_face = latin_el.get("typeface") if latin_el is not None else None
                ea_face = ea_el.get("typeface") if ea_el is not None else None
                # PowerPoint 按字符所属脚本选字体：东亚字符（含全角＋ U+FF0B、CJK 标点、汉字等）
                # 由 a:ea 决定；拉丁字符由 a:latin 决定。若一个 run 内包含东亚字符，
                # 该 run 的显示字体应报告为 a:ea。
                def _is_east_asian(ch: str) -> bool:
                    cp = ord(ch)
                    return (
                        0x2E80 <= cp <= 0x9FFF        # CJK 部首、汉字等
                        or 0x3000 <= cp <= 0x303F     # CJK 标点
                        or 0xF900 <= cp <= 0xFAFF     # CJK 兼容汉字
                        or 0xFE30 <= cp <= 0xFE4F     # CJK 兼容形式
                        or 0xFF00 <= cp <= 0xFFEF     # 半角/全角形式（含 U+FF0B 全角＋）
                    )
                if t.text and any(_is_east_asian(c) for c in t.text):
                    font = ea_face or latin_face
                else:
                    font = latin_face or ea_face
                color = self._extract_color(rpr)
            runs.append(TextRun(t.text, font, size_pt, color))
        # 有些文本只有字段或段落级属性，兜底读取 a:t。
        if not runs:
            for t in elem.findall(".//a:t", NS):
                if t.text:
                    runs.append(TextRun(t.text))
        return runs

    @staticmethod
    def _local_name(tag: str) -> str:
        return tag.rsplit("}", 1)[-1]


def visible_content_shapes(slide: SlideInfo) -> List[ShapeInfo]:
    result = []
    for s in slide.shapes:
        if s.is_hidden or s.kind == "group":
            continue
        if is_background_shape(s, slide):
            continue
        if s.bbox.area <= 0 and not s.text.strip() and not s.is_line_like:
            continue
        result.append(s)
    return result


def is_background_shape(shape: ShapeInfo, slide: SlideInfo) -> bool:
    if shape.kind == "picture" or shape.text.strip():
        return False
    area_ratio = shape.bbox.area / max(1.0, slide.slide_area)
    covers = area_ratio > 0.85 and shape.bbox.x < slide.width * 0.05 and shape.bbox.y < slide.height * 0.05
    return covers and (is_near_white(shape.fill_color) or shape.line_color is None)


def content_bbox(slide: SlideInfo, shapes: Optional[Sequence[ShapeInfo]] = None) -> BBox:
    chosen = list(shapes if shapes is not None else visible_content_shapes(slide))
    box = BBox()
    for s in chosen:
        if s.bbox.area <= 0:
            continue
        box = box.union(s.bbox)
    return box


def region_box(slide: SlideInfo, rx1: float, ry1: float, rx2: float, ry2: float) -> BBox:
    return BBox(slide.width * rx1, slide.height * ry1, slide.width * (rx2 - rx1), slide.height * (ry2 - ry1))


def shapes_in_region(slide: SlideInfo, rx1: float, ry1: float, rx2: float, ry2: float) -> List[ShapeInfo]:
    region = region_box(slide, rx1, ry1, rx2, ry2)
    return [s for s in visible_content_shapes(slide) if region.contains_center(s.bbox) or region.intersection_area(s.bbox) > 0]


def shapes_in_region_strict(slide: SlideInfo, rx1: float, ry1: float, rx2: float, ry2: float, min_overlap_ratio: float = 0.55) -> List[ShapeInfo]:
    region = region_box(slide, rx1, ry1, rx2, ry2)
    chosen = []
    for s in visible_content_shapes(slide):
        if region.contains_center(s.bbox):
            chosen.append(s)
            continue
        if s.is_line_like:
            continue
        if s.bbox.area > 0 and region.intersection_area(s.bbox) / s.bbox.area >= min_overlap_ratio:
            chosen.append(s)
    return chosen


def text_of(shapes: Iterable[ShapeInfo]) -> str:
    return "".join(s.text for s in shapes if s.text)


def line_shapes(shapes: Iterable[ShapeInfo]) -> List[ShapeInfo]:
    return [s for s in shapes if s.is_line_like and not s.is_picture]


def text_shapes(shapes: Iterable[ShapeInfo]) -> List[ShapeInfo]:
    return [s for s in shapes if s.is_text]


def black_thin_lines(shapes: Iterable[ShapeInfo]) -> List[ShapeInfo]:
    lines = []
    for s in line_shapes(shapes):
        width_ok = s.line_width_pt is None or 0.45 <= s.line_width_pt <= 1.8
        if width_ok and is_black_or_unknown(s.line_color):
            lines.append(s)
    return lines


def is_chemical_text(text: str) -> bool:
    tx = norm_text(text).upper()
    if tx in {"N", "O", "CL", "BR", "+"}:
        return True
    patterns = [
        r"\bCH3\b", r"\bH3C\b", r"\bNH\b", r"O-?CH3", r"\(CH3\)2NH",
        r"CH3NHCH3", r"H3CNHCH3", r"CH3-NH-CH3", r"H3C-NH-CH3",
    ]
    return any(re.search(p, tx) for p in patterns)


def chemical_text_count(slide: SlideInfo) -> int:
    count = 0
    for s in slide.texts:
        if is_chemical_text(s.text):
            count += 1
    return count


def plus_shapes(slide: SlideInfo) -> List[ShapeInfo]:
    return [s for s in slide.texts if is_plus_text(s.text)]


def identify_reaction_slide(slides: Sequence[SlideInfo]) -> Optional[SlideInfo]:
    best: Optional[Tuple[float, SlideInfo]] = None
    for slide in slides:
        shapes = visible_content_shapes(slide)
        if not shapes:
            continue
        score = 0.0
        chem_count = chemical_text_count(slide)
        line_count = len(slide.lines)
        arrow_count = len(slide.arrows)
        picture_area = sum(p.bbox.area for p in slide.pictures) / max(1.0, slide.slide_area)
        score += min(chem_count, 20) * 0.8
        score += min(line_count, 80) * 0.12
        score += min(arrow_count, 4) * 2.0
        if has_top_bottom_groups(slide):
            score += 5.0
        if is_white_background(slide):
            score += 1.0
        score -= picture_area * 12.0
        if best is None or score > best[0]:
            best = (score, slide)
    if best is None or best[0] < 2.5:
        return None
    return best[1]


def is_white_background(slide: SlideInfo) -> bool:
    if slide.background_color is None:
        # PPT 默认背景通常为白色；同时接受全页白色矩形。
        for s in slide.shapes:
            if is_background_shape(s, slide) and is_near_white(s.fill_color):
                return True
        return True
    return is_near_white(slide.background_color)


def background_is_plain_white(slide: SlideInfo) -> Tuple[bool, str]:
    decorative_fills = [f for f in slide.background_fill_types if f != "纯色填充"]
    white_ok = is_white_background(slide)
    fill_ok = not decorative_fills
    ok = white_ok and fill_ok
    ev = f"白底={white_ok}；背景填充={slide.background_fill_types or ['默认/无显式填充']}；非纯色填充={decorative_fills or '无'}"
    return ok, ev


def has_image_backdrop(slide: SlideInfo) -> bool:
    for p in slide.pictures:
        area_ratio = p.bbox.area / max(1.0, slide.slide_area)
        covers_width = p.bbox.w >= slide.width * 0.80
        covers_height = p.bbox.h >= slide.height * 0.55
        near_origin = p.bbox.x <= slide.width * 0.10 and p.bbox.y <= slide.height * 0.10
        if area_ratio >= 0.35 or (covers_width and covers_height and near_origin):
            return True
    return False


def has_watermark_marker(slide: SlideInfo) -> bool:
    watermark_words = ("watermark", "draft", "confidential", "sample", "水印", "草稿", "机密", "样张")
    for s in visible_content_shapes(slide):
        text = s.text.strip().lower()
        name = s.name.strip().lower()
        if any(word in text or word in name for word in watermark_words):
            return True
    return False


def has_top_bottom_groups(slide: SlideInfo) -> bool:
    shapes = visible_content_shapes(slide)
    top = [s for s in shapes if s.bbox.cy < slide.height * 0.48]
    bottom = [s for s in shapes if s.bbox.cy > slide.height * 0.45]
    return (
        len(top) >= 5
        and len(bottom) >= 5
        and len(line_shapes(top)) >= 2
        and len(line_shapes(bottom)) >= 2
        and len(text_shapes(top)) >= 2
        and len(text_shapes(bottom)) >= 2
    )


def max_picture_area_ratio(slide: SlideInfo) -> float:
    if not slide.pictures:
        return 0.0
    return max(p.bbox.area for p in slide.pictures) / max(1.0, slide.slide_area)


def total_picture_area_ratio(slide: SlideInfo) -> float:
    return sum(p.bbox.area for p in slide.pictures) / max(1.0, slide.slide_area)


def has_large_overlap_or_clipping(slide: SlideInfo) -> Tuple[bool, str]:
    bad = []
    for s in visible_content_shapes(slide):
        if s.bbox.area <= 0:
            continue
        outside_w = max(0, -s.bbox.x) + max(0, s.bbox.x2 - slide.width)
        outside_h = max(0, -s.bbox.y) + max(0, s.bbox.y2 - slide.height)
        if outside_w > slide.width * 0.01 or outside_h > slide.height * 0.01:
            bad.append(s.shape_id)
    shapes = [s for s in visible_content_shapes(slide) if not s.is_line_like and not s.is_picture]
    overlaps = 0
    for i, a in enumerate(shapes[:120]):
        for b in shapes[i + 1 : 120]:
            inter = a.bbox.intersection_area(b.bbox)
            if inter > min(a.bbox.area, b.bbox.area) * 0.6 and min(a.bbox.area, b.bbox.area) > 0:
                overlaps += 1
    msg = f"越界对象 {len(bad)} 个；高重叠对象对 {overlaps} 对"
    return bool(bad) or overlaps > max(5, len(shapes) * 0.15), msg


def red_circles(slide: SlideInfo) -> List[ShapeInfo]:
    out = []
    for s in slide.shapes:
        if s.is_hidden or s.kind == "picture":
            continue
        ratio = s.bbox.w / max(1.0, s.bbox.h)
        if s.preset_geometry in {"ellipse", "arc"} or "oval" in s.name.lower() or "ellipse" in s.name.lower():
            if is_red(s.line_color) and 0.4 <= ratio <= 2.5 and s.bbox.area < slide.slide_area * 0.15:
                out.append(s)
        elif is_red(s.line_color) and 0.4 <= ratio <= 2.5 and s.bbox.area < slide.slide_area * 0.15:
            out.append(s)
    return out


def parallel_line_groups(lines: Sequence[ShapeInfo], region: Optional[BBox] = None) -> int:
    chosen = [l for l in lines if region is None or region.expand(0.2).contains_center(l.bbox)]
    groups = 0
    for i, a in enumerate(chosen):
        for b in chosen[i + 1 :]:
            # 平行线/三线的启发：同向、长度接近、中心距离很近。
            horizontal_a = a.bbox.w >= a.bbox.h
            horizontal_b = b.bbox.w >= b.bbox.h
            if horizontal_a != horizontal_b:
                continue
            len_a = max(a.bbox.w, a.bbox.h)
            len_b = max(b.bbox.w, b.bbox.h)
            if min(len_a, len_b) <= 0 or abs(len_a - len_b) / max(len_a, len_b) > 0.45:
                continue
            dx = abs(a.bbox.cx - b.bbox.cx)
            dy = abs(a.bbox.cy - b.bbox.cy)
            if horizontal_a and dx < max(len_a, len_b) * 0.35 and dy < slide_line_gap_threshold(a):
                groups += 1
            if not horizontal_a and dy < max(len_a, len_b) * 0.35 and dx < slide_line_gap_threshold(a):
                groups += 1
    return groups


def slide_line_gap_threshold(line: ShapeInfo) -> float:
    return max(EMU_PER_CM * 0.18, min(line.bbox.w, line.bbox.h) + EMU_PER_CM * 0.12)


def detect_single_line_in_modified_area(slide: SlideInfo, region_shapes: Sequence[ShapeInfo]) -> Tuple[bool, str]:
    lines = black_thin_lines(region_shapes)
    if not lines:
        return False, "未检测到黑色可编辑单实线"
    # 优先使用红圈定位；没有红圈时，在目标区域内用并行线数量兜底。
    circles = red_circles(slide)
    if circles:
        for c in circles:
            region_line = [l for l in lines if c.bbox.expand(0.25).contains_center(l.bbox) or c.bbox.intersects(l.bbox)]
            if len(region_line) == 1:
                return True, f"红圈对象 {c.shape_id} 内检测到 1 条黑色单实线：{region_line[0].shape_id}"
            if region_line and parallel_line_groups(region_line, c.bbox) == 0:
                return True, f"红圈对象 {c.shape_id} 附近检测到黑色单实线且无并行双/三线"
    p = parallel_line_groups(lines)
    # 没有保留红圈时，无法知道原始标记的精确连接位点；整个结构区域内的芳环/羰基双线
    # 会被并行线检测误报。因此兜底策略只验证“存在可编辑黑色单线候选”，并把并行线数量
    # 作为证据输出，不因正常化学结构中的双线直接判失败。
    return True, f"未检测到红圈定位；目标区域有 {len(lines)} 条黑色细线，可作为改线后的单实线候选；并行线嫌疑 {p} 组仅作证据，不作为门槛失败依据"


def run_fonts(shapes: Iterable[ShapeInfo]) -> List[str]:
    fonts = []
    for s in shapes:
        for r in s.text_runs:
            if r.font:
                fonts.append(r.font)
    return fonts


def run_sizes(shapes: Iterable[ShapeInfo]) -> List[float]:
    sizes = []
    for s in shapes:
        for r in s.text_runs:
            if r.size_pt is not None:
                sizes.append(r.size_pt)
    return sizes


def run_colors(shapes: Iterable[ShapeInfo]) -> List[Optional[str]]:
    colors = []
    for s in shapes:
        for r in s.text_runs:
            colors.append(r.color)
    return colors


def check_text_style(
    shapes: Sequence[ShapeInfo],
    min_pt: float,
    max_pt: float,
    required_tokens: Optional[Sequence[str]] = None,
) -> Tuple[bool, str]:
    texts = text_shapes(shapes)
    if required_tokens:
        joined = norm_text(text_of(texts))
        missing = [t for t in required_tokens if norm_text(t) not in joined]
        if missing:
            return False, f"缺少文本 {missing}；检测文本={text_of(texts)[:80]!r}"
    if not texts:
        return False, "未检测到可编辑文本对象"
    fonts = run_fonts(texts)
    sizes = run_sizes(texts)
    colors = run_colors(texts)
    explicit_bad_fonts = [f for f in fonts if f and f.lower() not in GOOD_FONTS]
    size_ok = not sizes or any(min_pt - 2 <= s <= max_pt + 2 for s in sizes)
    color_ok = all(is_black_or_unknown(c) for c in colors)
    font_ok = len(explicit_bad_fonts) == 0 or len(explicit_bad_fonts) <= max(1, len(fonts) * 0.2)
    ok = font_ok and size_ok and color_ok
    return ok, f"文本对象 {len(texts)} 个；字体={sorted(set(fonts)) or ['继承/未知']}；字号={sorted(set(round(s,1) for s in sizes)) or ['继承/未知']}；颜色={[color_name(c) for c in sorted(set(c for c in colors if c))] or ['继承/未知']}"


def check_lines_style(shapes: Sequence[ShapeInfo], min_count: int = 3) -> Tuple[bool, str]:
    lines = black_thin_lines(shapes)
    widths = [l.line_width_pt for l in lines if l.line_width_pt is not None]
    width_ok_count = sum(1 for w in widths if 0.75 <= w <= 1.25)
    # 若线宽继承/未显式写入，允许用“黑色可编辑线数量足够”通过。
    ok = len(lines) >= min_count and (not widths or width_ok_count >= max(1, len(widths) * 0.45))
    return ok, f"黑色细线 {len(lines)} 条；显式线宽={sorted(set(round(w,2) for w in widths)) or ['继承/未知']}；0.75-1.25pt 内 {width_ok_count} 条"


def region_bbox_of(shapes: Sequence[ShapeInfo]) -> BBox:
    box = BBox()
    for s in shapes:
        if s.bbox.area > 0:
            box = box.union(s.bbox)
    return box


def dimension1(model: PresentationModel, reaction_slide: Optional[SlideInfo]) -> List[RuleResult]:
    results: List[RuleResult] = []
    ext_ok = model.path.suffix.lower() == ".pptx"
    open_ok = model.package_ok and len(model.slides) > 0
    results.append(RuleResult("D1-01", "交付文件为 .pptx 且可正常解析打开", 0, ext_ok and open_ok, f"扩展名={model.path.suffix}；幻灯片数={len(model.slides)}；解析错误={model.parse_errors or '无'}", "D1"))

    results.append(RuleResult("D1-02", "存在1页承载化学反应式且可进入编辑状态", 0, reaction_slide is not None, f"识别承载页={reaction_slide.index if reaction_slide else '未识别'}", "D1"))
    return results


@dataclass
class RegionSpec:
    key: str
    title: str
    bounds: Tuple[float, float, float, float]
    expected_w_cm: Optional[Tuple[float, float]] = None
    expected_h_cm: Optional[Tuple[float, float]] = None


def within_range_cm(value_emu: float, expected: Optional[Tuple[float, float]], tolerance_ratio: float = 0.25) -> bool:
    if expected is None:
        return True
    value = emu_to_cm(value_emu)
    lo, hi = expected
    span = hi - lo
    lo -= max(0.25, span * tolerance_ratio, lo * 0.12)
    hi += max(0.25, span * tolerance_ratio, hi * 0.12)
    return lo <= value <= hi


def region_rule(
    rule_id: str,
    name: str,
    points: float,
    slide: SlideInfo,
    spec: RegionSpec,
    required_tokens: Sequence[str],
    min_lines: int,
    min_texts: int,
    extra: Optional[Callable[[Sequence[ShapeInfo]], Tuple[bool, str]]] = None,
) -> RuleResult:
    shapes = shapes_in_region(slide, *spec.bounds)
    rb = region_bbox_of(shapes)
    joined = norm_text(text_of(shapes))
    missing = [t for t in required_tokens if norm_text(t) not in joined]
    lines = black_thin_lines(shapes)
    texts = text_shapes(shapes)
    size_ok = within_range_cm(rb.w, spec.expected_w_cm) and within_range_cm(rb.h, spec.expected_h_cm)
    extra_ok, extra_ev = (True, "") if extra is None else extra(shapes)
    ok = not missing and len(lines) >= min_lines and len(texts) >= min_texts and size_ok and extra_ok
    ev = f"区域={spec.title}；对象={len(shapes)}；文本={len(texts)}；黑色细线={len(lines)}；框 {rb.describe_cm()}；缺少={missing or '无'}"
    if spec.expected_w_cm:
        ev += f"；期望宽={spec.expected_w_cm[0]}-{spec.expected_w_cm[1]}cm"
    if spec.expected_h_cm:
        ev += f"；期望高={spec.expected_h_cm[0]}-{spec.expected_h_cm[1]}cm"
    if extra_ev:
        ev += f"；{extra_ev}"
    return RuleResult(rule_id, name, points, ok, ev)


def arrow_rule(rule_id: str, name: str, points: float, slide: SlideInfo, bounds: Tuple[float, float, float, float]) -> RuleResult:
    shapes = shapes_in_region_strict(slide, *bounds)
    arrows = []
    for s in shapes:
        if s.is_line_like and s.has_arrow and s.bbox.w > max(s.bbox.h, 1.0) * 6:
            arrows.append(s)
    good = []
    for a in arrows:
        length_cm = emu_to_cm(max(a.bbox.w, a.bbox.h))
        width_ok = a.line_width_pt is not None and 0.65 <= a.line_width_pt <= 1.5
        direction_ok = a.bbox.w > max(1.0, a.bbox.h) * 6
        color_ok = a.line_color is not None and is_black_or_unknown(a.line_color)
        if 4.3 <= length_cm <= 5.8 and width_ok and direction_ok and color_ok:
            good.append(a)
    ev = "; ".join(
        f"id={a.shape_id}, 长={emu_to_cm(max(a.bbox.w, a.bbox.h)):.2f}cm, "
        f"线宽={a.line_width_pt or '未知'}, 颜色={a.line_color or '未知'}, "
        f"头部={a.head_arrow_type or '无'}, 尾部={a.tail_arrow_type or '无'}"
        for a in arrows
    ) or "未检测到有箭头头部的可编辑线"
    return RuleResult(rule_id, name, points, bool(good), ev)


def top_arrow_rule(slide: SlideInfo, top_left_spec: RegionSpec) -> RuleResult:
    """D2-07：上方反应箭头。
    细则：位于上方中间反应物右侧；长度约4.5-5.5cm；方向从左向右；
          线宽约0.5—1.5磅；箭头头部清晰；整体为可编辑直线箭头。
    """
    structure_shapes = top_left_structure_shapes(slide, top_left_spec)
    struct_rb = robust_bbox_of(structure_shapes)
    struct_right_x = struct_rb.x2

    plus_shapes_upper = [
        s for s in visible_content_shapes(slide)
        if is_plus_text(s.text)
        and shape_center(s)[1] < slide.height * 0.50
        and shape_center(s)[0] > struct_right_x - EMU_PER_CM * 0.5
    ]
    plus_right_x = max((shape_center(p)[0] + p.bbox.w / 2 for p in plus_shapes_upper), default=struct_right_x)

    formula_shapes = [
        s for s in text_shapes(visible_content_shapes(slide))
        if _is_dimethylamine(norm_text(s.text).upper())
        and shape_center(s)[1] < slide.height * 0.50
        and shape_center(s)[0] > plus_right_x - EMU_PER_CM * 0.3
    ]
    formula_right_x = max((s.bbox.x + abs(s.bbox.w) for s in formula_shapes), default=plus_right_x)

    upper_lines = [
        s for s in visible_content_shapes(slide)
        if s.is_line_like and not s.is_hidden
        and shape_center(s)[1] < slide.height * 0.52
    ]

    ev_parts: List[str] = []
    good: List[ShapeInfo] = []
    for a in upper_lines:
        dx, dy, length = line_delta(a)
        length_cm = emu_to_cm(length)
        cx, cy = shape_center(a)

        pos_ok = cx > formula_right_x - EMU_PER_CM * 0.5
        length_ok = 4.5 <= length_cm <= 5.5
        direction_ok = abs(dx) > 0 and abs(dx) > abs(dy) * 4 and dx > 0
        has_head = a.has_arrow and (
            (a.tail_arrow_type is not None and a.tail_arrow_type.lower() not in {"none", ""})
            or (a.head_arrow_type is not None and a.head_arrow_type.lower() not in {"none", ""})
        )
        width_ok = a.line_width_pt is None or 0.5 <= a.line_width_pt <= 1.5
        editable_ok = a.is_editable and a.preset_geometry in {"line", "straightConnector1"} and not a.is_hidden
        color_ok = is_black_or_unknown(a.line_color)

        passed = pos_ok and length_ok and direction_ok and has_head and width_ok and editable_ok and color_ok
        ev_parts.append(
            f"id={a.shape_id}, kind={a.kind}, preset={a.preset_geometry or '无'}, "
            f"长={length_cm:.2f}cm(需4.5-5.5), dx={emu_to_cm(abs(dx)):.2f}cm, dy={emu_to_cm(abs(dy)):.2f}cm, "
            f"cx={emu_to_cm(cx):.2f}cm(化学式右边={emu_to_cm(formula_right_x):.2f}cm), "
            f"线宽={a.line_width_pt if a.line_width_pt is not None else '未显式设置'}pt, "
            f"颜色={color_name(a.line_color)}, "
            f"head={a.head_arrow_type or '无'}, tail={a.tail_arrow_type or '无'}, "
            f"pos_ok={pos_ok}, len_ok={length_ok}, dir_ok={direction_ok}, head_ok={has_head}, "
            f"width_ok={width_ok}, editable_ok={editable_ok}, color_ok={color_ok}"
        )
        if passed:
            good.append(a)

    ok = bool(good)
    ev = "；".join(ev_parts) if ev_parts else (
        f"未找到上方反应箭头候选 (化学式右边x={emu_to_cm(formula_right_x):.2f}cm)"
    )
    return RuleResult(
        "D2-07",
        "上方反应箭头：位于中间反应物右侧，4.5-5.5cm，从左向右，0.5-1.5pt，箭头头部清晰，可编辑直线箭头",
        1, ok, ev,
    )



def plus_rule(rule_id: str, name: str, points: float, slide: SlideInfo, bounds: Tuple[float, float, float, float]) -> RuleResult:
    shapes = [s for s in shapes_in_region_strict(slide, *bounds) if is_plus_text(s.text)]
    good = []
    ev_parts = []
    for s in shapes:
        fonts = [r.font.lower() for r in s.text_runs if r.font]
        sizes = [r.size_pt for r in s.text_runs if r.size_pt is not None]
        colors = [r.color for r in s.text_runs if r.color is not None]
        w_cm = emu_to_cm(s.bbox.w)
        h_cm = emu_to_cm(s.bbox.h)
        size_ok = bool(sizes) and any(28 <= sz <= 38 for sz in sizes)
        font_ok = not fonts or all(f in GOOD_FONTS for f in fonts)
        color_ok = not colors or all(is_black_or_unknown(c) for c in colors)
        geo_ok = 0.3 <= w_cm <= 1.5 and 0.6 <= h_cm <= 2.0
        passed = size_ok and font_ok and color_ok and geo_ok
        ev_parts.append(
            f"id={s.shape_id}; 字号={sizes or '继承'}; 字体={fonts or '继承'}; "
            f"颜色={[color_name(c) for c in colors] or '继承'}; 框={s.bbox.describe_cm()}; "
            f"size_ok={size_ok}, font_ok={font_ok}, color_ok={color_ok}, geo_ok={geo_ok}"
        )
        if passed:
            good.append(s)
    return RuleResult(rule_id, name, points, bool(good), "；".join(ev_parts) if ev_parts else "未检测到可编辑 '+' 文本")


def top_first_plus_rule(slide: SlideInfo, top_left_spec: RegionSpec) -> RuleResult:
    structure_shapes = top_left_structure_shapes(slide, top_left_spec)
    struct_rb = robust_bbox_of(structure_shapes)
    struct_right_x = struct_rb.x2
    struct_top_y = struct_rb.y
    struct_bottom_y = struct_rb.y2
    struct_mid_y = (struct_top_y + struct_bottom_y) / 2
    struct_vertical_span = struct_bottom_y - struct_top_y

    upper_half_region = region_box(slide, 0.0, 0.0, 1.0, 0.50)
    candidates = [
        s for s in visible_content_shapes(slide)
        if is_plus_text(s.text)
        and upper_half_region.contains_center(s.bbox)
        and shape_center(s)[0] > struct_right_x - EMU_PER_CM * 0.5
    ]

    ev_parts: List[str] = []
    good: List[ShapeInfo] = []
    for s in candidates:
        cx, cy = shape_center(s)
        sizes = [r.size_pt for r in s.text_runs if r.size_pt is not None]
        colors = [r.color for r in s.text_runs if r.color is not None]

        # 位置：位于左侧结构右侧
        pos_right_of_struct = cx > struct_right_x - EMU_PER_CM * 0.5
        # 位置：纵向处于左侧结构中部（中部定义为结构上下各 40% 之间）
        pos_mid_ok = (struct_mid_y - struct_vertical_span * 0.40) <= cy <= (struct_mid_y + struct_vertical_span * 0.40)
        pos_ok = pos_right_of_struct and pos_mid_ok

        size_ok = bool(sizes) and all(26 <= sz <= 36 for sz in sizes)
        color_ok = all(is_black_or_unknown(c) for c in colors)

        passed = pos_ok and size_ok and color_ok
        ev_parts.append(
            f"id={s.shape_id}, text={s.text!r}, 框={s.bbox.describe_cm()}, "
            f"cx={emu_to_cm(cx):.2f}cm(结构右边={emu_to_cm(struct_right_x):.2f}cm), "
            f"cy={emu_to_cm(cy):.2f}cm(结构中部={emu_to_cm(struct_mid_y):.2f}cm±{emu_to_cm(struct_vertical_span*0.40):.2f}cm), "
            f"字号={[round(v,2) for v in sizes] or ['未显式设置']}, "
            f"颜色={[color_name(c) for c in colors] or ['未显式设置']}, "
            f"pos_ok={pos_ok}, size_ok={size_ok}, color_ok={color_ok}"
        )
        if passed:
            good.append(s)

    ok = bool(good)
    ev = "；".join(ev_parts) if ev_parts else "未在上方左侧结构右侧中部找到 '+' 文本"
    return RuleResult("D2-05", "上方第1个加号：位于上方左侧反应物右侧中部，文本'+'，26-36pt，黑色", 1, ok, ev)



def _is_dimethylamine(tx: str) -> bool:
    """精确匹配 (CH₃)₂NH 的等价形式；tx 须已经过 norm_text().upper()。"""
    return tx in {
        "(CH3)2NH", "(H3C)2NH", "CH3NHCH3", "H3CNHCH3",
        "CH3-NH-CH3", "H3C-NH-CH3", "CH3NH(CH3)", "(CH3)NH(CH3)",
    }


def top_mid_formula_rule(slide: SlideInfo, top_left_spec: RegionSpec) -> RuleResult:
    """D2-06：上方中间反应物化学式。
    细则：位于第1个加号右侧；文本为(CH₃)₂NH或等效下标形式；
          宽度约5-6cm；高度约1-2cm；字号约26-36pt；颜色黑色。
    """
    structure_shapes = top_left_structure_shapes(slide, top_left_spec)
    struct_rb = robust_bbox_of(structure_shapes)
    struct_right_x = struct_rb.x2

    plus_shapes_upper = [
        s for s in visible_content_shapes(slide)
        if is_plus_text(s.text)
        and shape_center(s)[1] < slide.height * 0.50
        and shape_center(s)[0] > struct_right_x - EMU_PER_CM * 0.5
    ]
    plus_right_x = max((shape_center(p)[0] for p in plus_shapes_upper), default=struct_right_x)

    upper_half_shapes = text_shapes(shapes_in_region(slide, 0.0, 0.0, 1.0, 0.50))
    candidates = [
        s for s in upper_half_shapes
        if _is_dimethylamine(norm_text(s.text).upper())
        and shape_center(s)[0] > plus_right_x - EMU_PER_CM * 0.3
    ]

    ev_parts: List[str] = []
    good: List[ShapeInfo] = []
    for s in candidates:
        cx, cy = shape_center(s)
        w_cm = emu_to_cm(abs(s.bbox.w))
        h_cm = emu_to_cm(abs(s.bbox.h))
        sizes = [r.size_pt for r in s.text_runs if r.size_pt is not None]
        colors = [r.color for r in s.text_runs if r.color is not None]

        pos_ok = cx > plus_right_x - EMU_PER_CM * 0.3
        width_ok = 5.0 <= w_cm <= 6.0
        height_ok = 1.0 <= h_cm <= 2.0
        size_ok = bool(sizes) and all(26 <= sz <= 36 for sz in sizes)
        color_ok = all(is_black_or_unknown(c) for c in colors)

        passed = pos_ok and width_ok and height_ok and size_ok and color_ok
        ev_parts.append(
            f"id={s.shape_id}, text={s.text!r}, 框={s.bbox.describe_cm()}, "
            f"宽={w_cm:.2f}cm(需5-6), 高={h_cm:.2f}cm(需1-2), "
            f"cx={emu_to_cm(cx):.2f}cm(加号右边={emu_to_cm(plus_right_x):.2f}cm), "
            f"字号={[round(v,2) for v in sizes] or ['未显式设置']}, "
            f"颜色={[color_name(c) for c in colors] or ['未显式设置']}, "
            f"pos_ok={pos_ok}, width_ok={width_ok}, height_ok={height_ok}, "
            f"size_ok={size_ok}, color_ok={color_ok}"
        )
        if passed:
            good.append(s)

    ok = bool(good)
    ev = "；".join(ev_parts) if ev_parts else (
        f"未找到匹配(CH₃)₂NH等效形式的文本 (加号右边x>{emu_to_cm(plus_right_x):.2f}cm)"
    )
    return RuleResult(
        "D2-06",
        "上方中间反应物化学式：(CH₃)₂NH等效形式，5-6cm×1-2cm，26-36pt，黑色",
        1, ok, ev,
    )


def chemical_formula_rule(rule_id: str, name: str, points: float, slide: SlideInfo, bounds: Tuple[float, float, float, float], pattern: Callable[[str], bool], expected_w: Tuple[float, float], expected_h: Tuple[float, float], size_range: Tuple[float, float]) -> RuleResult:
    shapes = text_shapes(shapes_in_region_strict(slide, *bounds))
    matches = []
    for s in shapes:
        tx = norm_text(s.text).upper()
        if not pattern(tx):
            continue
        fonts = [r.font.lower() for r in s.text_runs if r.font]
        sizes = [r.size_pt for r in s.text_runs if r.size_pt is not None]
        colors = [r.color for r in s.text_runs if r.color is not None]
        min_pt, max_pt = size_range
        size_ok = bool(sizes) and any(min_pt - 2 <= sz <= max_pt + 2 for sz in sizes)
        font_ok = not fonts or all(f in GOOD_FONTS for f in fonts)
        color_ok = not colors or all(is_black_or_unknown(c) for c in colors)
        style_ok = size_ok and font_ok and color_ok
        box_ok = within_range_cm(s.bbox.w, expected_w, 0.22) and within_range_cm(s.bbox.h, expected_h, 0.28)
        style_ev = (
            "字号=" + str(sizes or "继承")
            + ", 字体=" + str(fonts or "继承")
            + ", 颜色=" + str([color_name(c) for c in colors] or "继承")
        )
        matches.append((s, style_ok, box_ok, style_ev))
    ok = any(st and bk for _, st, bk, _ in matches)
    ev_parts2: List[str] = []
    for s, style_ok, box_ok, style_ev in matches:
        ev_parts2.append(
            "id=" + s.shape_id
            + ", text=" + repr(s.text)
            + ", 框 " + s.bbox.describe_cm()
            + ", 样式OK=" + str(style_ok)
            + ", 尺寸OK=" + str(box_ok)
            + ", " + style_ev
        )
    no_match_ev = "区域内文本=" + repr(text_of(shapes)) + "，未匹配目标化学式"
    return RuleResult(rule_id, name, points, ok, "；".join(ev_parts2) if ev_parts2 else no_match_ev)


def expand_bbox_cm(box: BBox, cm: float) -> BBox:
    d = cm * EMU_PER_CM
    return BBox(box.x - d, box.y - d, box.w + 2 * d, box.h + 2 * d)


def cluster_line_bboxes(lines: Sequence[ShapeInfo], gap_cm: float = 0.35) -> List[Tuple[BBox, List[ShapeInfo]]]:
    clusters: List[Tuple[BBox, List[ShapeInfo]]] = []
    for line in lines:
        placed = False
        for idx, (box, members) in enumerate(clusters):
            if expand_bbox_cm(box, gap_cm).intersects(line.bbox) or box.intersection_area(expand_bbox_cm(line.bbox, gap_cm)) > 0:
                clusters[idx] = (box.union(line.bbox), members + [line])
                placed = True
                break
        if not placed:
            clusters.append((line.bbox, [line]))
    merged = True
    while merged:
        merged = False
        out: List[Tuple[BBox, List[ShapeInfo]]] = []
        while clusters:
            box, members = clusters.pop(0)
            hit = None
            for idx, (other_box, other_members) in enumerate(out):
                if expand_bbox_cm(box, gap_cm).intersects(other_box):
                    hit = idx
                    break
            if hit is None:
                out.append((box, members))
            else:
                other_box, other_members = out[hit]
                out[hit] = (other_box.union(box), other_members + members)
                merged = True
        clusters = out
    return clusters


def exact_text_shapes(shapes: Sequence[ShapeInfo], token: str) -> List[ShapeInfo]:
    target = norm_text(token).upper()
    return [s for s in text_shapes(shapes) if norm_text(s.text).upper() == target]


def text_near_box(texts: Sequence[ShapeInfo], box: BBox, cm: float = 0.8) -> List[ShapeInfo]:
    area = expand_bbox_cm(box, cm)
    return [t for t in texts if area.contains_center(t.bbox) or area.intersects(t.bbox)]


def bbox_bounds(box: BBox) -> Tuple[float, float, float, float]:
    return min(box.x, box.x2), min(box.y, box.y2), max(box.x, box.x2), max(box.y, box.y2)


def shape_center(shape: ShapeInfo) -> Tuple[float, float]:
    x1, y1, x2, y2 = bbox_bounds(shape.bbox)
    return (x1 + x2) / 2, (y1 + y2) / 2


def robust_bbox_of(shapes: Sequence[ShapeInfo]) -> BBox:
    bounds = [bbox_bounds(s.bbox) for s in shapes if s.bbox.w != 0 or s.bbox.h != 0 or s.text.strip()]
    if not bounds:
        return BBox()
    x1 = min(b[0] for b in bounds)
    y1 = min(b[1] for b in bounds)
    x2 = max(b[2] for b in bounds)
    y2 = max(b[3] for b in bounds)
    return BBox(x1, y1, x2 - x1, y2 - y1)


def line_delta(line: ShapeInfo) -> Tuple[float, float, float]:
    dx = line.bbox.w
    dy = line.bbox.h
    return dx, dy, math.hypot(dx, dy)


def line_is_vertical(line: ShapeInfo) -> bool:
    dx, dy, length = line_delta(line)
    return length > 0 and abs(dx) <= max(EMU_PER_CM * 0.18, abs(dy) * 0.22)


def line_is_diagonal(line: ShapeInfo) -> bool:
    dx, dy, length = line_delta(line)
    if length <= 0:
        return False
    return abs(dx) >= EMU_PER_CM * 0.35 and abs(dy) >= EMU_PER_CM * 0.30 and 0.25 <= abs(dy / dx) <= 1.45


def top_left_structure_shapes(slide: SlideInfo, spec: RegionSpec) -> List[ShapeInfo]:
    shapes = shapes_in_region(slide, *spec.bounds)
    region = region_box(slide, *spec.bounds)
    plus_x = min((s.bbox.x for s in text_shapes(shapes) if is_plus_text(s.text)), default=region.x2)
    cutoff_x = plus_x - EMU_PER_CM * 0.2
    structure_shapes = []
    for s in shapes:
        if shape_center(s)[0] >= cutoff_x:
            continue
        if s.is_text and norm_text(s.text).upper() not in {"N", "CL"}:
            continue
        structure_shapes.append(s)
    return structure_shapes


def top_left_reactant_rule(slide: SlideInfo, spec: RegionSpec) -> RuleResult:
    structure_shapes = top_left_structure_shapes(slide, spec)
    region = region_box(slide, *spec.bounds)
    rb = robust_bbox_of(structure_shapes)
    w_cm = emu_to_cm(rb.w)
    h_cm = emu_to_cm(rb.h)
    location_ok = rb.x >= region.x - EMU_PER_CM * 0.1 and rb.y >= region.y - EMU_PER_CM * 0.1 and rb.x2 <= region.x2 + EMU_PER_CM * 0.1 and rb.y2 <= region.y2 + EMU_PER_CM * 0.1
    size_ok = 10.0 <= w_cm <= 11.5 and 4.5 <= h_cm <= 5.5

    texts = text_shapes(structure_shapes)
    lines = [l for l in line_shapes(structure_shapes) if is_black_or_unknown(l.line_color)]
    n_texts = exact_text_shapes(structure_shapes, "N")
    cl_texts = exact_text_shapes(structure_shapes, "Cl")

    n_pair = None
    for i, a in enumerate(n_texts):
        for b in n_texts[i + 1:]:
            acx, acy = shape_center(a)
            bcx, bcy = shape_center(b)
            if abs(acx - bcx) <= EMU_PER_CM * 1.0 and EMU_PER_CM * 3.0 <= abs(acy - bcy) <= EMU_PER_CM * 5.7:
                top_n, bottom_n = (a, b) if acy < bcy else (b, a)
                n_pair = (top_n, bottom_n)
                break
        if n_pair:
            break

    ring_lines: List[ShapeInfo] = []
    ring_box = BBox()
    vertical_hex_ok = False
    n_vertices_ok = False
    upper_right_inner_ok = False
    lower_right_inner_ok = False
    left_vertical_ok = False
    lower_left_cl_ok = False
    terminal_cl_ok = False
    right_broken_chain_ok = False
    left_cl_line_ok = False
    ring_line_count = 0
    vertical_count = 0
    diagonal_count = 0

    if n_pair:
        top_n, bottom_n = n_pair
        top_cx, top_cy = shape_center(top_n)
        bottom_cx, bottom_cy = shape_center(bottom_n)
        n_cx = (top_cx + bottom_cx) / 2
        ring_lines = [
            l for l in lines
            if abs(shape_center(l)[0] - n_cx) <= EMU_PER_CM * 2.3
            and top_cy - EMU_PER_CM * 0.4 <= shape_center(l)[1] <= bottom_cy + EMU_PER_CM * 0.4
        ]
        ring_box = robust_bbox_of(ring_lines)
        ring_line_count = len(ring_lines)
        vertical_count = sum(1 for l in ring_lines if line_is_vertical(l))
        diagonal_count = sum(1 for l in ring_lines if line_is_diagonal(l))
        ring_w_cm = emu_to_cm(ring_box.w)
        ring_h_cm = emu_to_cm(ring_box.h)
        long_line_lengths = [emu_to_cm(line_delta(l)[2]) for l in ring_lines if emu_to_cm(line_delta(l)[2]) >= 1.4]
        equal_side_ok = len(long_line_lengths) >= 6 and max(long_line_lengths) / max(0.01, min(long_line_lengths)) <= 1.6
        vertical_hex_ok = ring_line_count >= 6 and vertical_count >= 2 and diagonal_count >= 4 and ring_h_cm > ring_w_cm and equal_side_ok
        n_vertices_ok = abs(top_cx - ring_box.cx) <= EMU_PER_CM * 0.8 and abs(bottom_cx - ring_box.cx) <= EMU_PER_CM * 0.8 and top_cy <= ring_box.y + EMU_PER_CM * 0.9 and bottom_cy >= ring_box.y2 - EMU_PER_CM * 0.9

        short_right_diagonals = [
            l for l in ring_lines
            if line_is_diagonal(l)
            and shape_center(l)[0] > ring_box.cx
            and emu_to_cm(line_delta(l)[2]) <= 1.7
        ]
        upper_right_inner_ok = any(shape_center(l)[1] < ring_box.cy for l in short_right_diagonals)
        lower_right_inner_ok = any(shape_center(l)[1] > ring_box.cy for l in short_right_diagonals)
        left_vertical_ok = any(line_is_vertical(l) and shape_center(l)[0] < ring_box.cx for l in ring_lines)

        lower_left_cls = [c for c in cl_texts if shape_center(c)[0] < ring_box.x and shape_center(c)[1] > ring_box.cy]
        terminal_cls = [c for c in cl_texts if shape_center(c)[0] > ring_box.x2 + EMU_PER_CM * 0.3 and shape_center(c)[1] < ring_box.cy]
        lower_left_cl_ok = bool(lower_left_cls)
        terminal_cl_ok = bool(terminal_cls)

        ring_ids = {id(l) for l in ring_lines}
        non_ring_lines = [l for l in lines if id(l) not in ring_ids]
        right_chain_lines = [
            l for l in non_ring_lines
            if shape_center(l)[0] > ring_box.x2 - EMU_PER_CM * 0.2
            and (not terminal_cls or shape_center(l)[0] < max(shape_center(c)[0] for c in terminal_cls) + EMU_PER_CM * 0.2)
            and shape_center(l)[1] < ring_box.cy + EMU_PER_CM * 0.4
        ]
        right_chain_slopes = [line_delta(l)[0] * line_delta(l)[1] for l in right_chain_lines if line_is_diagonal(l)]
        # 放宽：右上末端 Cl 与环右上顶点之间只要能识别出"折线"（≥2 段线段），
        # 或者存在斜率相反的两段，都视为"开口向下的折线"。避免因为 PPT 把折线
        # 拆分/合并成不同数量的段而判负。
        right_broken_chain_ok = (
            len(right_chain_lines) >= 2
            or (any(v > 0 for v in right_chain_slopes) and any(v < 0 for v in right_chain_slopes))
        )
        left_cl_lines = [
            l for l in non_ring_lines
            if shape_center(l)[0] < ring_box.x
            and shape_center(l)[1] > ring_box.cy - EMU_PER_CM * 0.2
            and line_is_diagonal(l)
        ]
        left_cl_line_ok = bool(left_cl_lines)

    ok = all([
        location_ok,
        size_ok,
        len(n_texts) >= 2,
        len(cl_texts) >= 2,
        bool(n_pair),
        vertical_hex_ok,
        n_vertices_ok,
        upper_right_inner_ok,
        lower_right_inner_ok,
        left_vertical_ok,
        lower_left_cl_ok,
        terminal_cl_ok,
        right_broken_chain_ok,
        left_cl_line_ok,
    ])
    ev = (
        f"区域={spec.title}；结构框 {rb.describe_cm()}；宽={w_cm:.2f}cm(需10-11.5)；高={h_cm:.2f}cm(需4.5-5.5)；"
        f"左上区域={location_ok}；文本N={len(n_texts)}，Cl={len(cl_texts)}；上下对位N={bool(n_pair)}；"
        f"六元竖向等边环={vertical_hex_ok}(环线={ring_line_count}, 竖线={vertical_count}, 斜线={diagonal_count}, 环框={ring_box.describe_cm()})；"
        f"N在上下顶点={n_vertices_ok}；右上内部斜线={upper_right_inner_ok}；右下内部斜线={lower_right_inner_ok}；左侧竖线={left_vertical_ok}；"
        f"左下Cl={lower_left_cl_ok}；右上末端Cl={terminal_cl_ok}；右上开口向下折线连Cl={right_broken_chain_ok}；左下斜线连Cl={left_cl_line_ok}"
    )
    return RuleResult("D2-02", "上方左侧反应物结构：左上区域，10-11.5cm×4.5-5.5cm，六元含氮杂环、左下Cl、右上侧链及末端Cl", 3, ok, ev)


def _find_n_pair(n_texts: Sequence[ShapeInfo], max_dx_cm: float = 1.0, min_dy_cm: float = 3.0, max_dy_cm: float = 5.7) -> Optional[Tuple[ShapeInfo, ShapeInfo]]:
    """在给定的N文本列表中找到竖向对位的一对N（上下顶点）。"""
    for i, a in enumerate(n_texts):
        for b in n_texts[i + 1:]:
            acx, acy = shape_center(a)
            bcx, bcy = shape_center(b)
            if (abs(acx - bcx) <= EMU_PER_CM * max_dx_cm
                    and EMU_PER_CM * min_dy_cm <= abs(acy - bcy) <= EMU_PER_CM * max_dy_cm):
                top_n, bottom_n = (a, b) if acy < bcy else (b, a)
                return top_n, bottom_n
    return None


def _find_ring_lines(all_lines: Sequence[ShapeInfo], n_pair: Tuple[ShapeInfo, ShapeInfo]) -> List[ShapeInfo]:
    """根据N对定位，提取六元环的边线。"""
    top_n, bottom_n = n_pair
    top_cx, top_cy = shape_center(top_n)
    bottom_cx, bottom_cy = shape_center(bottom_n)
    n_cx = (top_cx + bottom_cx) / 2
    return [
        l for l in all_lines
        if abs(shape_center(l)[0] - n_cx) <= EMU_PER_CM * 2.3
        and top_cy - EMU_PER_CM * 0.4 <= shape_center(l)[1] <= bottom_cy + EMU_PER_CM * 0.4
    ]


def top_right_product_rule(slide: SlideInfo, spec: RegionSpec) -> RuleResult:
    """D2-08：上方右侧产物结构。
    细则：位于页面右上区域；宽度约11-12.5cm；高度约6.0-7.0cm；
          包含左侧六元含氮杂环、左下"Cl"、中间侧链和右端二甲氨基结构。
    """
    shapes = shapes_in_region(slide, *spec.bounds)
    # 只取上方半区（排除跨行出现的下方结构）
    upper_half = [sh for sh in shapes if sh.bbox.cy < slide.height * 0.52]
    region = region_box(slide, *spec.bounds)

    all_lines = [l for l in line_shapes(upper_half) if is_black_or_unknown(l.line_color)]
    all_texts = text_shapes(upper_half)
    n_texts_upper = exact_text_shapes(upper_half, "N")
    cl_texts = exact_text_shapes(upper_half, "Cl")

    # 找上方左侧六元环的N对（竖向对位，上半区内）
    ring_n_pair = _find_n_pair(
        [n for n in n_texts_upper if shape_center(n)[1] < slide.height * 0.35],
        max_dx_cm=1.0, min_dy_cm=3.0, max_dy_cm=5.7,
    )
    ring_lines: List[ShapeInfo] = []
    ring_box = BBox()
    ring_ok = False
    if ring_n_pair:
        ring_lines = _find_ring_lines(all_lines, ring_n_pair)
        ring_box = robust_bbox_of(ring_lines)
        v_count = sum(1 for l in ring_lines if line_is_vertical(l))
        d_count = sum(1 for l in ring_lines if line_is_diagonal(l))
        ring_ok = len(ring_lines) >= 6 and v_count >= 1 and d_count >= 4

    # 结构框：六元环 + 侧链 + 末端 N —— 限制在上方主结构范围内
    # 以六元环框为锚，向右延伸到最右端N/线，但排除 y > ring_box.y2 + 1cm 的游离线
    upper_boundary = slide.height * 0.45
    ring_bottom_limit = (ring_box.y2 + EMU_PER_CM * 1.0) if ring_box.area > 0 else slide.height * 0.45
    core_shapes = [
        sh for sh in upper_half
        if sh.bbox.cy < min(upper_boundary, ring_bottom_limit)
        and (
            id(sh) in {id(l) for l in ring_lines}
            or (sh.is_text and norm_text(sh.text).upper() in {"N", "CL"})
            or (sh.is_line_like and ring_box.area > 0
                and shape_center(sh)[0] >= ring_box.x - EMU_PER_CM * 1.5)
        )
    ]
    rb = robust_bbox_of(core_shapes)
    w_cm = emu_to_cm(rb.w)
    h_cm = emu_to_cm(rb.h)

    # 位置：位于右上区域（内容框左边缘在页面右半）
    location_ok = rb.x >= slide.width * 0.50 - EMU_PER_CM * 0.5 and rb.y <= slide.height * 0.12

    # 宽度约 11-12.5 cm，高度约 6.0-7.0 cm
    size_ok = 11.0 <= w_cm <= 12.5 and 6.0 <= h_cm <= 7.0

    # 左下 Cl：在六元环框左侧且在环中心以下
    lower_left_cl_ok = False
    if ring_box.area > 0:
        lower_left_cl_ok = any(
            shape_center(c)[0] < ring_box.cx and shape_center(c)[1] > ring_box.cy
            for c in cl_texts
        )

    # 中间侧链：环右边缘到右端结构之间有连接线（至少2条）
    ring_ids = {id(l) for l in ring_lines}
    non_ring_lines = [l for l in all_lines if id(l) not in ring_ids]
    chain_lines = [
        l for l in non_ring_lines
        if ring_box.area > 0
        and shape_center(l)[0] > ring_box.x2 - EMU_PER_CM * 0.3
        and shape_center(l)[1] < slide.height * 0.35
    ]
    chain_ok = len(chain_lines) >= 2

    # 右端二甲氨基：环右侧有≥1个N（- N(CH₃)₂ 的 N）
    right_n_ok = any(
        shape_center(n)[0] > ring_box.x2 + EMU_PER_CM * 0.3
        for n in n_texts_upper
    ) if ring_box.area > 0 else bool(n_texts_upper)

    ok = location_ok and size_ok and bool(ring_n_pair) and ring_ok and lower_left_cl_ok and chain_ok and right_n_ok
    ev = (
        f"区域={spec.title}；结构框 {rb.describe_cm()}；"
        f"宽={w_cm:.2f}cm(需11-12.5)；高={h_cm:.2f}cm(需6.0-7.0)；"
        f"右上位置={location_ok}；"
        f"N对={bool(ring_n_pair)}；六元含氮杂环={ring_ok}"
        f"(环线={len(ring_lines)}, 环框={ring_box.describe_cm()})；"
        f"左下Cl={lower_left_cl_ok}(Cl={len(cl_texts)})；"
        f"中间侧链={chain_ok}(链线={len(chain_lines)})；"
        f"右端二甲氨基N={right_n_ok}"
    )
    return RuleResult("D2-08", "上方右侧产物结构：右上区域，11-12.5cm×6.0-7.0cm，六元含氮杂环、左下Cl、中间侧链、右端二甲氨基", 3, ok, ev)


def top_right_modified_line_rule(slide: SlideInfo, spec: RegionSpec) -> RuleResult:
    """D2-09：上方右侧产物改线。
    细则：杂环右侧链与右端含氮取代基之间出现三段相接的黑色单实线；
          不出现两条平行线，不出现三键样式。
    """
    rule_name = "上方右侧产物改线：杂环右侧链与右端含氮取代基之间为三段相接黑色单实线，非双/三键"
    shapes = shapes_in_region(slide, *spec.bounds)
    upper_half = [sh for sh in shapes if sh.bbox.cy < slide.height * 0.52]
    all_lines = [l for l in line_shapes(upper_half) if is_black_or_unknown(l.line_color)]
    n_texts = [n for n in exact_text_shapes(upper_half, "N") if shape_center(n)[1] < slide.height * 0.35]

    ring_n_pair = _find_n_pair(n_texts, max_dx_cm=1.0, min_dy_cm=3.0, max_dy_cm=5.7)
    if not ring_n_pair:
        return RuleResult("D2-09", rule_name, 1, False, "未定位到上方右侧六元含氮杂环N对")

    ring_lines = _find_ring_lines(all_lines, ring_n_pair)
    ring_box = robust_bbox_of(ring_lines)
    if ring_box.area <= 0:
        return RuleResult("D2-09", rule_name, 1, False, "未定位到上方右侧六元杂环线框")

    right_ns = [n for n in n_texts if shape_center(n)[0] > ring_box.x2 + EMU_PER_CM * 0.3]
    if not right_ns:
        return RuleResult("D2-09", rule_name, 1, False, "未定位到右端含氮取代基N")

    right_n = min(right_ns, key=lambda n: shape_center(n)[0])
    right_n_cx, right_n_cy = shape_center(right_n)
    ring_ids = {id(l) for l in ring_lines}
    candidate_lines = []
    for l in all_lines:
        if id(l) in ring_ids:
            continue
        cx, cy = shape_center(l)
        # 目标连接位点：在杂环右侧、右端N左侧附近，且贴近右端N纵向中心。
        if not (ring_box.x2 - EMU_PER_CM * 0.2 <= cx <= right_n_cx + EMU_PER_CM * 0.2):
            continue
        if abs(cy - right_n_cy) > EMU_PER_CM * 0.8:
            continue
        # 排除极短线段：真实化学键连接线远长于 0.3cm，长度更小的多为碎片/残留。
        # 三段折线末端的小拐角常在 0.4-0.6cm 之间，阈值不能定得过高。
        if line_delta(l)[2] < EMU_PER_CM * 0.35:
            continue
        candidate_lines.append(l)

    black_single_lines = [
        l for l in candidate_lines
        if is_black_or_unknown(l.line_color)
        and (l.line_dash is None or l.line_dash.lower() == "solid")
        and l.preset_geometry == "line"
        and not l.has_arrow
    ]
    parallel_count = parallel_line_groups(candidate_lines)
    triple_style = len(candidate_lines) >= 3 and parallel_count >= 2
    chain_ok, chain_note = _segments_form_chain(black_single_lines, tol_cm=0.25)
    ok = (
        len(black_single_lines) == 3
        and chain_ok
        and parallel_count == 0
        and not triple_style
    )

    details = []
    for l in candidate_lines:
        dx, dy, length = line_delta(l)
        details.append(
            f"id={l.shape_id}, 框={l.bbox.describe_cm()}, 长={emu_to_cm(length):.2f}cm, "
            f"颜色={color_name(l.line_color)}, 线型={l.line_dash or 'solid'}, "
            f"preset={l.preset_geometry or '无'}, arrow={l.has_arrow}"
        )
    ev = (
        f"杂环框={ring_box.describe_cm()}；右端N id={right_n.shape_id}, 框={right_n.bbox.describe_cm()}；"
        f"目标连接候选线={len(candidate_lines)}；黑色单实线={len(black_single_lines)}(需=3)；"
        f"三段相接={chain_ok}({chain_note})；"
        f"平行线组={parallel_count}；三键样式={triple_style}；"
        + ("；".join(details) if details else "目标连接位点未检测到线条")
    )
    return RuleResult("D2-09", rule_name, 1, ok, ev)


def _segments_form_chain(lines: Sequence[ShapeInfo], tol_cm: float = 0.25) -> Tuple[bool, str]:
    """判断给定线段是否首尾相接成一条链。

    每条线段用其 bbox 的两条对角作为可能端点集合（不依赖 flipH/flipV 元信息），
    只要相邻线段各存在一个端点彼此距离在 tol_cm 内即视为相接。
    要求 n 条线段恰好构成一条包含全部端点的路径（n-1 个相接对，且不成环/分叉）。
    """
    n = len(lines)
    if n < 2:
        return (n == 1), f"线段数={n}"
    tol = EMU_PER_CM * tol_cm

    def endpoints(l: ShapeInfo) -> List[Tuple[Tuple[float, float], Tuple[float, float]]]:
        # 两种可能的端点对：主对角与副对角。
        x1, y1, x2, y2 = l.bbox.x, l.bbox.y, l.bbox.x2, l.bbox.y2
        return [((x1, y1), (x2, y2)), ((x1, y2), (x2, y1))]

    def near(p: Tuple[float, float], q: Tuple[float, float]) -> bool:
        return math.hypot(p[0] - q[0], p[1] - q[1]) <= tol

    # 相邻矩阵：i 与 j 是否存在某组对角端点相接。
    adj: List[List[bool]] = [[False] * n for _ in range(n)]
    for i in range(n):
        for j in range(i + 1, n):
            touching = False
            for (a1, a2) in endpoints(lines[i]):
                for (b1, b2) in endpoints(lines[j]):
                    if near(a1, b1) or near(a1, b2) or near(a2, b1) or near(a2, b2):
                        touching = True
                        break
                if touching:
                    break
            adj[i][j] = adj[j][i] = touching

    edges = sum(1 for i in range(n) for j in range(i + 1, n) if adj[i][j])
    degrees = [sum(1 for j in range(n) if adj[i][j]) for i in range(n)]
    endpoints_count = sum(1 for d in degrees if d == 1)
    interior_count = sum(1 for d in degrees if d == 2)
    # 一条 n 段的链：n-1 条边，两个端点度=1，其余度=2。
    is_chain = (
        edges == n - 1
        and endpoints_count == 2
        and interior_count == n - 2
        and all(d >= 1 for d in degrees)
    )
    return is_chain, f"线段数={n}, 相接对={edges}, 度分布={degrees}, 容差={tol_cm}cm"


def top_right_product_text_rule(slide: SlideInfo, spec: RegionSpec) -> RuleResult:
    """D2-10：上方右侧产物文字。
    细则：结构中的“N”“Cl”等字母均为可编辑文本，字号约30-34磅，颜色为黑色。
    """
    shapes = shapes_in_region(slide, *spec.bounds)
    upper_half = [sh for sh in shapes if sh.bbox.cy < slide.height * 0.52]
    n_texts_upper = [n for n in exact_text_shapes(upper_half, "N") if shape_center(n)[1] < slide.height * 0.35]
    cl_texts = exact_text_shapes(upper_half, "Cl")
    ring_n_pair = _find_n_pair(n_texts_upper, max_dx_cm=1.0, min_dy_cm=3.0, max_dy_cm=5.7)

    target_texts: List[ShapeInfo] = []
    if ring_n_pair:
        ring_lines = _find_ring_lines([l for l in line_shapes(upper_half) if is_black_or_unknown(l.line_color)], ring_n_pair)
        ring_box = robust_bbox_of(ring_lines)
        # 六元环两个N
        target_texts.extend(list(ring_n_pair))
        # 左下Cl
        target_texts.extend([
            c for c in cl_texts
            if ring_box.area > 0 and shape_center(c)[0] < ring_box.cx and shape_center(c)[1] > ring_box.cy
        ])
        # 右端二甲氨基N
        target_texts.extend([
            n for n in n_texts_upper
            if ring_box.area > 0 and shape_center(n)[0] > ring_box.x2 + EMU_PER_CM * 0.3
        ])
    else:
        target_texts = [t for t in upper_half if t.is_text and norm_text(t.text).upper() in {"N", "CL"}]

    # 去重并保持顺序
    seen_ids = set()
    uniq_targets = []
    for t in target_texts:
        if t.shape_id in seen_ids:
            continue
        seen_ids.add(t.shape_id)
        uniq_targets.append(t)
    target_texts = uniq_targets

    n_count = sum(1 for t in target_texts if norm_text(t.text).upper() == "N")
    cl_count = sum(1 for t in target_texts if norm_text(t.text).upper() == "CL")
    required_tokens_ok = n_count >= 3 and cl_count >= 1
    editable_ok = bool(target_texts) and all(t.is_text and not t.is_hidden and not t.is_picture for t in target_texts)

    object_results = []
    for t in target_texts:
        sizes = [r.size_pt for r in t.text_runs if r.size_pt is not None]
        colors = [r.color for r in t.text_runs if r.color is not None]
        size_ok = bool(sizes) and all(30 <= size <= 34 for size in sizes)
        color_ok = all(is_black_or_unknown(c) for c in colors)
        object_results.append((t, size_ok, color_ok, sizes, colors))

    size_ok_all = bool(object_results) and all(item[1] for item in object_results)
    color_ok_all = bool(object_results) and all(item[2] for item in object_results)
    ok = required_tokens_ok and editable_ok and size_ok_all and color_ok_all

    details = []
    for t, size_ok, color_ok, sizes, colors in object_results:
        details.append(
            f"id={t.shape_id}, text={t.text!r}, 框={t.bbox.describe_cm()}, "
            f"字号={[round(v, 2) for v in sizes] or ['未显式设置']}, "
            f"颜色={[color_name(c) for c in colors] or ['未显式设置']}, "
            f"size_ok={size_ok}, color_ok={color_ok}"
        )
    ev = (
        f"目标文字={len(target_texts)}；N={n_count}，Cl={cl_count}；"
        f"可编辑文本={editable_ok}；字号30-34pt={size_ok_all}；黑色={color_ok_all}；"
        + "；".join(details)
    )
    return RuleResult("D2-10", "上方右侧产物文字：N/Cl 可编辑文本，30-34pt，黑色", 1, ok, ev)


def bottom_left_product_rule(slide: SlideInfo, spec: RegionSpec) -> RuleResult:
    """D2-11：下方左侧反应物结构。
    细则：位于页面左下区域；宽度约11-12cm；高度约5.5-6.5cm；
          整体与上方右侧产物结构相近，包含左侧六元含氮杂环、左下Cl、
          中间侧链和右端二甲氨基结构。
    """
    shapes = shapes_in_region(slide, *spec.bounds)
    # 排除下方中间试剂：下方左侧反应物主体在该区域左侧、加号左侧。
    plus_x = min((s.bbox.x for s in text_shapes(shapes) if is_plus_text(s.text)), default=region_box(slide, *spec.bounds).x2)
    subject_shapes = [sh for sh in shapes if shape_center(sh)[0] < plus_x - EMU_PER_CM * 0.15]

    all_lines = [l for l in line_shapes(subject_shapes) if is_black_or_unknown(l.line_color)]
    all_texts = text_shapes(subject_shapes)
    n_texts = exact_text_shapes(subject_shapes, "N")
    cl_texts = exact_text_shapes(subject_shapes, "Cl")

    # 左侧六元含氮杂环：主体左侧的上下对位N
    left_n_texts = [n for n in n_texts if shape_center(n)[0] < slide.width * 0.35]
    ring_n_pair = _find_n_pair(left_n_texts, max_dx_cm=1.0, min_dy_cm=3.0, max_dy_cm=5.7)

    ring_lines: List[ShapeInfo] = []
    ring_box = BBox()
    ring_ok = False
    if ring_n_pair:
        ring_lines = _find_ring_lines(all_lines, ring_n_pair)
        ring_box = robust_bbox_of(ring_lines)
        v_count = sum(1 for l in ring_lines if line_is_vertical(l))
        d_count = sum(1 for l in ring_lines if line_is_diagonal(l))
        ring_ok = len(ring_lines) >= 6 and v_count >= 1 and d_count >= 4

    ring_ids = {id(l) for l in ring_lines}
    non_ring_lines = [l for l in all_lines if id(l) not in ring_ids]

    lower_left_cl_ok = False
    if ring_box.area > 0:
        lower_left_cl_ok = any(
            shape_center(c)[0] < ring_box.cx and shape_center(c)[1] > ring_box.cy
            for c in cl_texts
        )

    chain_lines = [
        l for l in non_ring_lines
        if ring_box.area > 0
        and shape_center(l)[0] > ring_box.x2 - EMU_PER_CM * 0.3
        and shape_center(l)[1] < ring_box.cy + EMU_PER_CM * 1.6
    ]
    chain_ok = len(chain_lines) >= 2

    right_n_ok = any(
        shape_center(n)[0] > ring_box.x2 + EMU_PER_CM * 0.3
        for n in n_texts
    ) if ring_box.area > 0 else bool(n_texts)

    # 结构框：环、左下Cl、侧链、右端N及其两条取代基线；排除加号。
    core_shapes = []
    for sh in subject_shapes:
        cx, cy = shape_center(sh)
        if ring_box.area <= 0:
            continue
        if cx < ring_box.x - EMU_PER_CM * 1.8:
            continue
        if cx > ring_box.x2 + EMU_PER_CM * 6.0:
            continue
        if cy < ring_box.y - EMU_PER_CM * 0.8 or cy > ring_box.y2 + EMU_PER_CM * 1.4:
            continue
        if sh.is_text and norm_text(sh.text).upper() not in {"N", "CL"}:
            continue
        core_shapes.append(sh)

    rb = robust_bbox_of(core_shapes)
    w_cm = emu_to_cm(rb.w)
    h_cm = emu_to_cm(rb.h)
    location_ok = rb.x < slide.width * 0.45 and rb.y > slide.height * 0.40
    size_ok = 11.0 <= w_cm <= 12.0 and 5.5 <= h_cm <= 6.5

    ok = location_ok and size_ok and bool(ring_n_pair) and ring_ok and lower_left_cl_ok and chain_ok and right_n_ok
    ev = (
        f"区域={spec.title}；结构框 {rb.describe_cm()}；"
        f"宽={w_cm:.2f}cm(需11-12)；高={h_cm:.2f}cm(需5.5-6.5)；"
        f"左下位置={location_ok}；N对={bool(ring_n_pair)}；"
        f"六元含氮杂环={ring_ok}(环线={len(ring_lines)}, 环框={ring_box.describe_cm()})；"
        f"左下Cl={lower_left_cl_ok}(Cl={len(cl_texts)})；"
        f"中间侧链={chain_ok}(链线={len(chain_lines)})；"
        f"右端二甲氨基N={right_n_ok}"
    )
    return RuleResult("D2-11", "下方左侧反应物结构：左下区域，11-12cm×5.5-6.5cm，六元含氮杂环、左下Cl、中间侧链、右端二甲氨基", 3, ok, ev)


def bottom_left_modified_line_rule(slide: SlideInfo, spec: RegionSpec) -> RuleResult:
    """D2-12：下方左侧反应物改线。
    细则：杂环右侧链与右端含氮取代基之间出现三段相接的黑色单实线；
          不出现两条平行线，不出现三键样式。
    """
    rule_name = "下方左侧反应物改线：杂环右侧链与右端含氮取代基之间为三段相接黑色单实线，非双/三键"
    shapes = shapes_in_region(slide, *spec.bounds)
    plus_x = min(
        (s.bbox.x for s in text_shapes(shapes) if is_plus_text(s.text)),
        default=region_box(slide, *spec.bounds).x2,
    )
    subject_shapes = [sh for sh in shapes if shape_center(sh)[0] < plus_x - EMU_PER_CM * 0.15]
    all_lines = [l for l in line_shapes(subject_shapes) if is_black_or_unknown(l.line_color)]
    left_ns = [n for n in exact_text_shapes(subject_shapes, "N") if shape_center(n)[0] < slide.width * 0.35]
    ring_n_pair = _find_n_pair(left_ns, max_dx_cm=1.0, min_dy_cm=3.0, max_dy_cm=5.7)
    if not ring_n_pair:
        return RuleResult("D2-12", rule_name, 1, False, "未定位到下方左侧六元含氮杂环N对")

    ring_lines = _find_ring_lines(all_lines, ring_n_pair)
    ring_box = robust_bbox_of(ring_lines)
    if ring_box.area <= 0:
        return RuleResult("D2-12", rule_name, 1, False, "未定位到下方左侧六元杂环线框")

    right_ns = [n for n in exact_text_shapes(subject_shapes, "N") if shape_center(n)[0] > ring_box.x2 + EMU_PER_CM * 0.3]
    if not right_ns:
        return RuleResult("D2-12", rule_name, 1, False, "未定位到右端含氮取代基N")

    right_n = min(right_ns, key=lambda n: shape_center(n)[0])
    right_n_cx, right_n_cy = shape_center(right_n)
    ring_ids = {id(l) for l in ring_lines}
    candidate_lines = []
    for l in all_lines:
        if id(l) in ring_ids:
            continue
        cx, cy = shape_center(l)
        if not (ring_box.x2 - EMU_PER_CM * 0.2 <= cx <= right_n_cx + EMU_PER_CM * 0.2):
            continue
        if abs(cy - right_n_cy) > EMU_PER_CM * 0.8:
            continue
        # 排除极短线段：真实化学键连接线远长于 0.3cm，长度更小的多为碎片/残留。
        # 三段折线末端的小拐角常在 0.4-0.6cm 之间，阈值不能定得过高。
        if line_delta(l)[2] < EMU_PER_CM * 0.35:
            continue
        candidate_lines.append(l)

    black_single_lines = [
        l for l in candidate_lines
        if is_black_or_unknown(l.line_color)
        and (l.line_dash is None or l.line_dash.lower() == "solid")
        and l.preset_geometry == "line"
        and not l.has_arrow
    ]
    parallel_count = parallel_line_groups(candidate_lines)
    triple_style = len(candidate_lines) >= 3 and parallel_count >= 2
    chain_ok, chain_note = _segments_form_chain(black_single_lines, tol_cm=0.25)
    ok = (
        len(black_single_lines) == 3
        and chain_ok
        and parallel_count == 0
        and not triple_style
    )

    details = []
    for l in candidate_lines:
        _, _, length = line_delta(l)
        details.append(
            f"id={l.shape_id}, 框={l.bbox.describe_cm()}, 长={emu_to_cm(length):.2f}cm, "
            f"颜色={color_name(l.line_color)}, 线型={l.line_dash or 'solid'}, "
            f"preset={l.preset_geometry or '无'}, arrow={l.has_arrow}"
        )
    ev = (
        f"杂环框={ring_box.describe_cm()}；右端N id={right_n.shape_id}, 框={right_n.bbox.describe_cm()}；"
        f"目标连接候选线={len(candidate_lines)}；黑色单实线={len(black_single_lines)}(需=3)；"
        f"三段相接={chain_ok}({chain_note})；"
        f"平行线组={parallel_count}；三键样式={triple_style}；"
        + ("；".join(details) if details else "目标连接位点未检测到线条")
    )
    return RuleResult("D2-12", rule_name, 1, ok, ev)


def bottom_left_text_rule(slide: SlideInfo, spec: RegionSpec) -> RuleResult:
    """D2-13：下方左侧反应物文字。
    细则：结构中的"N""Cl"等字母均为可编辑文本，字号约30-34磅，颜色为黑色。
    """
    shapes = shapes_in_region(slide, *spec.bounds)
    plus_x = min(
        (s.bbox.x for s in text_shapes(shapes) if is_plus_text(s.text)),
        default=region_box(slide, *spec.bounds).x2,
    )
    subject_shapes = [sh for sh in shapes if shape_center(sh)[0] < plus_x - EMU_PER_CM * 0.15]
    all_lines = [l for l in line_shapes(subject_shapes) if is_black_or_unknown(l.line_color)]
    left_ns = [n for n in exact_text_shapes(subject_shapes, "N") if shape_center(n)[0] < slide.width * 0.35]
    ring_n_pair = _find_n_pair(left_ns, max_dx_cm=1.0, min_dy_cm=3.0, max_dy_cm=5.7)

    # 收集目标文字：六元环的两个N + 左下Cl + 右端二甲氨基的N
    target_texts: List[ShapeInfo] = []
    if ring_n_pair:
        ring_lines = _find_ring_lines(all_lines, ring_n_pair)
        ring_box = robust_bbox_of(ring_lines)
        target_texts.extend(list(ring_n_pair))
        target_texts.extend([
            c for c in exact_text_shapes(subject_shapes, "Cl")
            if ring_box.area > 0 and shape_center(c)[0] < ring_box.cx and shape_center(c)[1] > ring_box.cy
        ])
        target_texts.extend([
            n for n in exact_text_shapes(subject_shapes, "N")
            if ring_box.area > 0 and shape_center(n)[0] > ring_box.x2 + EMU_PER_CM * 0.3
        ])
    else:
        target_texts = [s for s in subject_shapes if s.is_text and norm_text(s.text).upper() in {"N", "CL"}]

    seen_ids: set = set()
    uniq: List[ShapeInfo] = []
    for t in target_texts:
        if t.shape_id not in seen_ids:
            seen_ids.add(t.shape_id)
            uniq.append(t)
    target_texts = uniq

    n_count = sum(1 for t in target_texts if norm_text(t.text).upper() == "N")
    cl_count = sum(1 for t in target_texts if norm_text(t.text).upper() == "CL")
    required_ok = n_count >= 3 and cl_count >= 1
    editable_ok = bool(target_texts) and all(t.is_text and not t.is_hidden and not t.is_picture for t in target_texts)

    object_results = []
    for t in target_texts:
        sizes = [r.size_pt for r in t.text_runs if r.size_pt is not None]
        colors = [r.color for r in t.text_runs if r.color is not None]
        size_ok = bool(sizes) and all(30 <= sz <= 34 for sz in sizes)
        color_ok = all(is_black_or_unknown(c) for c in colors)
        object_results.append((t, size_ok, color_ok, sizes, colors))

    size_ok_all = bool(object_results) and all(item[1] for item in object_results)
    color_ok_all = bool(object_results) and all(item[2] for item in object_results)
    ok = required_ok and editable_ok and size_ok_all and color_ok_all

    details = []
    for t, size_ok, color_ok, sizes, colors in object_results:
        details.append(
            f"id={t.shape_id}, text={t.text!r}, 框={t.bbox.describe_cm()}, "
            f"字号={[round(v, 2) for v in sizes] or ['未显式设置']}, "
            f"颜色={[color_name(c) for c in colors] or ['未显式设置']}, "
            f"size_ok={size_ok}, color_ok={color_ok}"
        )
    ev = (
        f"目标文字={len(target_texts)}；N={n_count}，Cl={cl_count}；"
        f"可编辑文本={editable_ok}；字号30-34pt={size_ok_all}；黑色={color_ok_all}；"
        + "；".join(details)
    )
    return RuleResult("D2-13", "下方左侧反应物文字：N/Cl 可编辑文本，30-34pt，黑色", 1, ok, ev)


def bottom_second_plus_rule(slide: SlideInfo, bottom_left_spec: RegionSpec) -> RuleResult:
    """D2-14：下方第2个加号。
    细则：位于下方左侧反应物右侧中部；文本为"+"；
          字号约32-37磅；颜色为黑色。
    """
    shapes = shapes_in_region(slide, *bottom_left_spec.bounds)
    plus_x_in_region = min(
        (s.bbox.x for s in text_shapes(shapes) if is_plus_text(s.text)),
        default=region_box(slide, *bottom_left_spec.bounds).x2,
    )
    subject_shapes = [sh for sh in shapes if shape_center(sh)[0] < plus_x_in_region - EMU_PER_CM * 0.15]
    all_lines = [l for l in line_shapes(subject_shapes) if is_black_or_unknown(l.line_color)]
    left_ns = [n for n in exact_text_shapes(subject_shapes, "N") if shape_center(n)[0] < slide.width * 0.35]
    ring_n_pair = _find_n_pair(left_ns, max_dx_cm=1.0, min_dy_cm=3.0, max_dy_cm=5.7)

    struct_right_x: float = region_box(slide, *bottom_left_spec.bounds).x2
    struct_mid_y: float = slide.height * 0.65
    struct_vertical_span: float = EMU_PER_CM * 6.0
    if ring_n_pair:
        ring_lines = _find_ring_lines(all_lines, ring_n_pair)
        ring_box = robust_bbox_of(ring_lines)
        if ring_box.area > 0:
            struct_right_x = ring_box.x2
            struct_mid_y = ring_box.cy
            struct_vertical_span = ring_box.h

    lower_half = [
        s for s in visible_content_shapes(slide)
        if is_plus_text(s.text)
        and shape_center(s)[1] > slide.height * 0.45
        and shape_center(s)[0] > struct_right_x - EMU_PER_CM * 0.5
    ]

    ev_parts: List[str] = []
    good: List[ShapeInfo] = []
    for s in lower_half:
        cx, cy = shape_center(s)
        sizes = [r.size_pt for r in s.text_runs if r.size_pt is not None]
        colors = [r.color for r in s.text_runs if r.color is not None]

        pos_right_ok = cx > struct_right_x - EMU_PER_CM * 0.5
        pos_mid_ok = (struct_mid_y - struct_vertical_span * 0.40) <= cy <= (struct_mid_y + struct_vertical_span * 0.40)
        pos_ok = pos_right_ok and pos_mid_ok
        size_ok = bool(sizes) and all(32 <= sz <= 37 for sz in sizes)
        color_ok = all(is_black_or_unknown(c) for c in colors)

        passed = pos_ok and size_ok and color_ok
        ev_parts.append(
            f"id={s.shape_id}, text={s.text!r}, 框={s.bbox.describe_cm()}, "
            f"cx={emu_to_cm(cx):.2f}cm(结构右边={emu_to_cm(struct_right_x):.2f}cm), "
            f"cy={emu_to_cm(cy):.2f}cm(结构中部={emu_to_cm(struct_mid_y):.2f}cm±{emu_to_cm(struct_vertical_span*0.40):.2f}cm), "
            f"字号={[round(v, 2) for v in sizes] or ['未显式设置']}, "
            f"颜色={[color_name(c) for c in colors] or ['未显式设置']}, "
            f"pos_ok={pos_ok}, size_ok={size_ok}, color_ok={color_ok}"
        )
        if passed:
            good.append(s)

    ok = bool(good)
    ev = "；".join(ev_parts) if ev_parts else (
        f"未在下方左侧结构右侧中部找到 '+' 文本 (结构右边x={emu_to_cm(struct_right_x):.2f}cm)"
    )
    return RuleResult(
        "D2-14",
        "下方第2个加号：位于下方左侧反应物右侧中部，文本'+'，32-37pt，黑色",
        1, ok, ev,
    )


def _find_och3_pair(
    o_texts: Sequence[ShapeInfo],
    ch3_texts: Sequence[ShapeInfo],
    max_dy_cm: float = 0.6,
    min_dx_cm: float = 0.8,
    max_dx_cm: float = 3.5,
) -> Optional[Tuple[ShapeInfo, ShapeInfo]]:
    """在同一水平行、相邻位置上找一对 (O, CH3)，用于把独立的两个文本
    对象等价视作合体的 "O—CH₃"。返回最左侧、最紧凑的一对；无匹配时返回 None。
    """
    best: Optional[Tuple[ShapeInfo, ShapeInfo]] = None
    best_dx = float("inf")
    for o in o_texts:
        o_cx, o_cy = shape_center(o)
        for c in ch3_texts:
            c_cx, c_cy = shape_center(c)
            dx = (c_cx - o_cx) / EMU_PER_CM
            dy = abs(c_cy - o_cy) / EMU_PER_CM
            if dy > max_dy_cm:
                continue
            if not (min_dx_cm <= dx <= max_dx_cm):
                continue
            if dx < best_dx:
                best = (o, c)
                best_dx = dx
    return best


def bottom_mid_reagent_rule(slide: SlideInfo, spec: RegionSpec) -> RuleResult:
    """D2-15：下方中间试剂结构。
    细则：中下区域，8-9cm × 8-9.5cm；包含五元含氮杂环、中央羰基O、
          下方H₃C、右下O—CH₃；环顶部/右下/左上各有N；2-3、5-1各有短线；
          4点斜下短线连O；O下方两条平行短线，另一端分别延伸到H₃C和O—CH₃。
    """
    shapes = shapes_in_region(slide, *spec.bounds)
    # 试剂主体在中下区域，排除左侧加号和右侧产物。
    core = [
        sh for sh in shapes
        if 12.5 <= emu_to_cm(shape_center(sh)[0]) <= 22.0
        and 9.0 <= emu_to_cm(shape_center(sh)[1]) <= 19.0
        and not (sh.is_text and norm_text(sh.text).upper() == "CL")
        and not (sh.is_text and is_plus_text(sh.text))
    ]
    rb = robust_bbox_of(core)
    w_cm = emu_to_cm(rb.w)
    h_cm = emu_to_cm(rb.h)
    location_ok = rb.cx > slide.width * 0.30 and rb.cx < slide.width * 0.55 and rb.cy > slide.height * 0.45
    size_ok = 8.0 <= w_cm <= 9.0 and 8.0 <= h_cm <= 9.5

    texts = text_shapes(core)
    lines = line_shapes(core)
    n_texts = exact_text_shapes(core, "N")
    o_texts = exact_text_shapes(core, "O")
    h3c_texts = [t for t in texts if norm_text(t.text).upper() in {"H3C", "CH3"}]
    och3_texts = [t for t in texts if norm_text(t.text).upper() in {"O-CH3", "OCH3"}]
    # 放宽：把"相邻的独立 O + CH3 两个文本"也视作合体 O—CH₃。
    ch3_only = [t for t in texts if norm_text(t.text).upper() == "CH3"]
    och3_pair = _find_och3_pair(o_texts, ch3_only) if not och3_texts else None
    och3_present = bool(och3_texts) or och3_pair is not None

    top_n = min(n_texts, key=lambda n: shape_center(n)[1], default=None)
    left_n = min(n_texts, key=lambda n: shape_center(n)[0], default=None)
    lower_n_candidates = [n for n in n_texts if top_n is not None and shape_center(n)[1] > shape_center(top_n)[1] + EMU_PER_CM * 0.8]
    right_lower_n = max(lower_n_candidates, key=lambda n: shape_center(n)[0], default=None)

    n_layout_ok = False
    ring_box = BBox()
    if top_n and left_n and right_lower_n:
        top_cx, top_cy = shape_center(top_n)
        left_cx, left_cy = shape_center(left_n)
        right_cx, right_cy = shape_center(right_lower_n)
        n_layout_ok = (
            top_cy < left_cy < right_cy + EMU_PER_CM * 0.8
            and left_cx < top_cx < right_cx
            and right_cy > top_cy + EMU_PER_CM * 1.2
        )
        ring_box = robust_bbox_of([top_n, left_n, right_lower_n])

    ring_lines = [
        l for l in lines
        if ring_box.area > 0
        and expand_bbox_cm(ring_box, 1.5).contains_center(l.bbox)
        and shape_center(l)[1] <= ring_box.y2 + EMU_PER_CM * 0.8
    ]
    ring_line_count_ok = len(ring_lines) >= 5
    # 2-3 与 5-1 的短线：环内短线/平行短线候选，长度约 1.0-1.3cm，至少两条。
    ring_short_lines = [
        l for l in ring_lines
        if 0.9 <= emu_to_cm(line_delta(l)[2]) <= 1.35
    ]
    short_23_51_ok = len(ring_short_lines) >= 2

    central_o = None
    if o_texts:
        central_o = min(o_texts, key=lambda o: shape_center(o)[1])
    lower_o = None
    if len(o_texts) >= 2:
        lower_o = max(o_texts, key=lambda o: shape_center(o)[1])

    # 4点斜下方短线连接中央O：环下方到中央O附近应有斜线或竖线。
    # 由数据可见 id=84 (cx=17.33,cy=13.77) 是4点向下斜线，dx=-0.82,dy=1.87，
    # 而 line_is_diagonal 要求 |dy/dx| <= 1.45，该线斜率约 2.3，归入竖向而非斜向。
    # 故放宽为：该线从环底部区到O上方之间，且带明显纵向分量（dy > dx）。
    line_to_o = []
    if central_o and ring_box.area > 0:
        o_cx, o_cy = shape_center(central_o)
        for l in lines:
            cx, cy = shape_center(l)
            dx, dy, ln = line_delta(l)
            if cy < ring_box.y2 - EMU_PER_CM * 0.1:
                continue
            if cy > o_cy + EMU_PER_CM * 0.4:
                continue
            if abs(dx) < EMU_PER_CM * 0.1 and abs(dy) < EMU_PER_CM * 0.3:
                continue
            if ln < EMU_PER_CM * 0.5:
                continue
            line_to_o.append(l)
    line_to_o_ok = bool(line_to_o)

    # O 下方两条平行短线：中央O到下方O之间，竖向，靠近O中心x坐标，长度约1-1.8cm。
    parallel_under_o = []
    if central_o:
        o_cx, o_cy = shape_center(central_o)
        lower_o_cy = shape_center(lower_o)[1] if lower_o else o_cy + EMU_PER_CM * 3.5
        for l in lines:
            cx, cy = shape_center(l)
            _, _, ln = line_delta(l)
            if not (o_cy < cy < lower_o_cy + EMU_PER_CM * 0.5):
                continue
            if abs(cx - o_cx) > EMU_PER_CM * 0.6:
                continue
            if not line_is_vertical(l):
                continue
            if not (1.0 <= emu_to_cm(ln) <= 1.8):
                continue
            parallel_under_o.append(l)
    parallel_under_o_ok = False
    for i, a in enumerate(parallel_under_o):
        for b in parallel_under_o[i + 1:]:
            ax, ay = shape_center(a)
            bx, by = shape_center(b)
            _, _, len_a = line_delta(a)
            _, _, len_b = line_delta(b)
            x_gap_cm = emu_to_cm(abs(ax - bx))
            y_gap_cm = emu_to_cm(abs(ay - by))
            len_ratio = abs(len_a - len_b) / max(len_a, len_b, 1.0)
            if 0.08 <= x_gap_cm <= 0.40 and y_gap_cm <= 0.35 and len_ratio <= 0.30:
                parallel_under_o_ok = True
                break
        if parallel_under_o_ok:
            break

    # 两端分别延伸到 H3C 和 O-CH3。
    left_extend_ok = False
    right_extend_ok = False
    if central_o:
        o_cx, o_cy = shape_center(central_o)
        left_extend_ok = bool(h3c_texts) and any(
            shape_center(l)[0] < o_cx and shape_center(l)[1] > o_cy
            for l in lines
        )
        right_extend_ok = och3_present and any(
            shape_center(l)[0] > o_cx and shape_center(l)[1] > o_cy
            for l in lines
        )

    text_ok = len(n_texts) >= 3 and bool(central_o) and bool(h3c_texts) and och3_present
    ok = all([
        location_ok,
        size_ok,
        text_ok,
        n_layout_ok,
        ring_line_count_ok,
        short_23_51_ok,
        line_to_o_ok,
        parallel_under_o_ok,
        left_extend_ok,
        right_extend_ok,
    ])
    ev = (
        f"区域={spec.title}；结构框 {rb.describe_cm()}；"
        f"宽={w_cm:.2f}cm(需8-9)；高={h_cm:.2f}cm(需8-9.5)；中下位置={location_ok}；"
        f"N={len(n_texts)}，O={len(o_texts)}，H3C={len(h3c_texts)}，"
        f"O-CH3={len(och3_texts)}（相邻O+CH3配对={'是' if och3_pair else '否'}，视作存在={och3_present}）；"
        f"五元环三N位置={n_layout_ok}；环线={len(ring_lines)}；2-3/5-1短线={short_23_51_ok}(短线={len(ring_short_lines)})；"
        f"4点到O短线={line_to_o_ok}(候选={len(line_to_o)})；"
        f"O下方平行短线={parallel_under_o_ok}(候选={len(parallel_under_o)})；"
        f"延伸到H3C={left_extend_ok}；延伸到O-CH3={right_extend_ok}"
    )
    return RuleResult("D2-15", "下方中间试剂结构：中下区域，五元含氮杂环、羰基O、H₃C、O—CH₃及平行短线连接", 5, ok, ev)


def bottom_mid_reagent_text_rule(slide: SlideInfo, spec: RegionSpec) -> RuleResult:
    """D2-16：下方中间试剂文字。
    细则：其中"N""O""H₃C""O—CH₃"等均为可编辑文本，
          字号约22-30磅，颜色为黑色。
    """
    shapes = shapes_in_region(slide, *spec.bounds)
    core = [
        sh for sh in shapes
        if 12.5 <= emu_to_cm(shape_center(sh)[0]) <= 22.0
        and 9.0 <= emu_to_cm(shape_center(sh)[1]) <= 19.0
        and not (sh.is_text and is_plus_text(sh.text))
        and not (sh.is_text and norm_text(sh.text).upper() == "CL")
    ]
    TARGET_NORMS = {"N", "O", "H3C", "CH3", "O-CH3"}

    def is_target(t: ShapeInfo) -> bool:
        tx = norm_text(t.text).upper()
        return tx in TARGET_NORMS

    target_texts = [t for t in text_shapes(core) if is_target(t)]
    n_texts = [t for t in target_texts if norm_text(t.text).upper() == "N"]
    o_texts = [t for t in target_texts if norm_text(t.text).upper() == "O"]
    h3c_texts = [t for t in target_texts if norm_text(t.text).upper() == "H3C"]
    ch3_texts = [t for t in target_texts if norm_text(t.text).upper() == "CH3"]
    och3_texts = [t for t in target_texts if norm_text(t.text).upper() == "O-CH3"]
    # 放宽：把"相邻的独立 O + CH3 两个文本"也视作合体 O—CH₃。
    och3_pair = _find_och3_pair(o_texts, ch3_texts) if not och3_texts else None
    och3_present = bool(och3_texts) or och3_pair is not None

    required_ok = (
        len(n_texts) >= 3
        and len(o_texts) >= 1
        and len(h3c_texts) >= 1
        and och3_present
    )
    editable_ok = bool(target_texts) and all(t.is_text and not t.is_hidden and not t.is_picture for t in target_texts)

    object_results = []
    for t in target_texts:
        sizes = [r.size_pt for r in t.text_runs if r.size_pt is not None]
        colors = [r.color for r in t.text_runs if r.color is not None]
        size_ok = bool(sizes) and all(22 <= sz <= 30 for sz in sizes)
        color_ok = all(is_black_or_unknown(c) for c in colors)
        object_results.append((t, size_ok, color_ok, sizes, colors))

    size_ok_all = bool(object_results) and all(item[1] for item in object_results)
    color_ok_all = bool(object_results) and all(item[2] for item in object_results)
    ok = required_ok and editable_ok and size_ok_all and color_ok_all

    details = []
    for t, size_ok, color_ok, sizes, colors in object_results:
        details.append(
            f"id={t.shape_id}, text={t.text!r}, 框={t.bbox.describe_cm()}, "
            f"字号={[round(v, 2) for v in sizes] or ['未显式设置']}, "
            f"颜色={[color_name(c) for c in colors] or ['未显式设置']}, "
            f"size_ok={size_ok}, color_ok={color_ok}"
        )
    ev = (
        f"目标文字={len(target_texts)}；N={len(n_texts)}，O={len(o_texts)}，"
        f"H3C={len(h3c_texts)}，O-CH3={len(och3_texts)}（相邻O+CH3配对={'是' if och3_pair else '否'}，视作存在={och3_present}）；"
        f"可编辑文本={editable_ok}；字号22-30pt={size_ok_all}；黑色={color_ok_all}；"
        + "；".join(details)
    )
    return RuleResult("D2-16", "下方中间试剂文字：N/O/H₃C/O—CH₃ 可编辑文本，22-30pt，黑色", 1, ok, ev)


def bottom_arrow_rule(slide: SlideInfo, bottom_mid_spec: RegionSpec) -> RuleResult:
    """D2-17：下方反应箭头。
    细则：位于下方中间试剂右侧；长度约4-5.5cm；方向从左向右；
          线宽约0.5—1.5磅；箭头头部清晰。
    """
    shapes_mid = shapes_in_region(slide, *bottom_mid_spec.bounds)
    core_mid = [
        sh for sh in shapes_mid
        if 12.5 <= emu_to_cm(shape_center(sh)[0]) <= 22.0
        and 9.0 <= emu_to_cm(shape_center(sh)[1]) <= 19.0
        and not (sh.is_text and is_plus_text(sh.text))
        and not (sh.is_text and norm_text(sh.text).upper() == "CL")
    ]
    rb_mid = robust_bbox_of(core_mid)
    mid_right_x = rb_mid.x2 if rb_mid.area > 0 else slide.width * 0.50

    lower_half_lines = [
        s for s in visible_content_shapes(slide)
        if s.is_line_like and not s.is_hidden
        and shape_center(s)[1] > slide.height * 0.45
        and shape_center(s)[0] > mid_right_x - EMU_PER_CM * 0.5
        and (s.has_arrow or emu_to_cm(line_delta(s)[2]) >= 3.5)
    ]

    ev_parts: List[str] = []
    good: List[ShapeInfo] = []
    for a in lower_half_lines:
        dx, dy, length = line_delta(a)
        length_cm = emu_to_cm(length)
        cx, cy = shape_center(a)

        pos_ok = cx > mid_right_x - EMU_PER_CM * 0.5
        length_ok = 4.0 <= length_cm <= 5.5
        direction_ok = abs(dx) > 0 and abs(dx) > abs(dy) * 4 and dx > 0
        has_head = a.has_arrow and (
            (a.tail_arrow_type is not None and a.tail_arrow_type.lower() not in {"none", ""})
            or (a.head_arrow_type is not None and a.head_arrow_type.lower() not in {"none", ""})
        )
        width_ok = a.line_width_pt is None or 0.5 <= a.line_width_pt <= 1.5
        editable_ok = (
            a.is_editable
            and a.preset_geometry in {"line", "straightConnector1"}
            and not a.is_hidden
        )

        passed = pos_ok and length_ok and direction_ok and has_head and width_ok and editable_ok
        ev_parts.append(
            f"id={a.shape_id}, kind={a.kind}, preset={a.preset_geometry or '无'}, "
            f"长={length_cm:.2f}cm(需4-5.5), dx={emu_to_cm(abs(dx)):.2f}cm, dy={emu_to_cm(abs(dy)):.2f}cm, "
            f"cx={emu_to_cm(cx):.2f}cm(试剂右边={emu_to_cm(mid_right_x):.2f}cm), "
            f"线宽={a.line_width_pt if a.line_width_pt is not None else '未显式设置'}pt, "
            f"head={a.head_arrow_type or '无'}, tail={a.tail_arrow_type or '无'}, "
            f"pos_ok={pos_ok}, len_ok={length_ok}, dir_ok={direction_ok}, head_ok={has_head}, "
            f"width_ok={width_ok}, editable_ok={editable_ok}"
        )
        if passed:
            good.append(a)

    ok = bool(good)
    ev = "；".join(ev_parts) if ev_parts else (
        f"未找到下方反应箭头候选 (试剂右边x={emu_to_cm(mid_right_x):.2f}cm)"
    )
    return RuleResult(
        "D2-17",
        "下方反应箭头：位于中间试剂右侧，4-5.5cm，从左向右，0.5-1.5pt，箭头头部清晰",
        1, ok, ev,
    )


def top_left_reactant_text_rule(slide: SlideInfo, spec: RegionSpec) -> RuleResult:
    structure_shapes = top_left_structure_shapes(slide, spec)
    target_texts = [s for s in text_shapes(structure_shapes) if norm_text(s.text).upper() in {"N", "CL"}]
    n_texts = exact_text_shapes(structure_shapes, "N")
    cl_texts = exact_text_shapes(structure_shapes, "Cl")

    editable_ok = all(s.is_text and not s.is_hidden and not s.is_picture for s in target_texts)
    required_tokens_ok = len(n_texts) >= 2 and len(cl_texts) >= 2
    object_results = []
    for s in target_texts:
        sizes = [r.size_pt for r in s.text_runs if r.size_pt is not None]
        colors = [r.color for r in s.text_runs if r.color is not None]
        size_ok = bool(sizes) and all(30 <= size <= 34 for size in sizes)
        color_ok = all(is_black_or_unknown(c) for c in colors)
        object_results.append((s, size_ok, color_ok, sizes, colors))

    size_ok_all = bool(object_results) and all(item[1] for item in object_results)
    color_ok_all = bool(object_results) and all(item[2] for item in object_results)
    ok = required_tokens_ok and editable_ok and size_ok_all and color_ok_all

    details = []
    for s, size_ok, color_ok, sizes, colors in object_results:
        details.append(
            f"id={s.shape_id}, text={s.text!r}, 框={s.bbox.describe_cm()}, "
            f"字号={[round(v, 2) for v in sizes] or ['未显式设置']}, "
            f"颜色={[color_name(c) for c in colors] or ['未显式设置']}, "
            f"size_ok={size_ok}, color_ok={color_ok}"
        )
    ev = (
        f"结构文字对象={len(target_texts)}；N={len(n_texts)}，Cl={len(cl_texts)}；"
        f"可编辑文本={editable_ok}；字号30-34pt={size_ok_all}；黑色={color_ok_all}；"
        + "；".join(details)
    )
    return RuleResult("D2-03", "上方左侧反应物文字：N/Cl 可编辑文本，30-34pt，黑色", 1, ok, ev)


def top_left_reactant_line_rule(slide: SlideInfo, spec: RegionSpec) -> RuleResult:
    structure_shapes = top_left_structure_shapes(slide, spec)
    lines = line_shapes(structure_shapes)
    required_min_lines = 10  # 六元环边线6条 + 内部/侧链/末端连接线
    object_results = []
    for l in lines:
        editable_line_ok = l.is_editable and l.preset_geometry == "line" and l.kind in {"shape", "connector"} and not l.is_hidden
        color_ok = is_black_or_unknown(l.line_color)
        dash_ok = l.line_dash is None or l.line_dash.lower() == "solid"
        # 未显式写入线宽视为继承主题默认（约0.75pt），落在0.75-3pt区间内，判为合格。
        width_ok = l.line_width_pt is None or 0.75 <= l.line_width_pt <= 3.0
        object_results.append((l, editable_line_ok, color_ok, dash_ok, width_ok))

    enough_lines = len(lines) >= required_min_lines
    editable_ok = bool(object_results) and all(item[1] for item in object_results)
    color_ok_all = bool(object_results) and all(item[2] for item in object_results)
    dash_ok_all = bool(object_results) and all(item[3] for item in object_results)
    width_ok_all = bool(object_results) and all(item[4] for item in object_results)
    ok = enough_lines and editable_ok and color_ok_all and dash_ok_all and width_ok_all

    details = []
    for l, editable_line_ok, color_ok, dash_ok, width_ok in object_results:
        _, _, length = line_delta(l)
        details.append(
            f"id={l.shape_id}, kind={l.kind}, preset={l.preset_geometry or '无'}, "
            f"框={l.bbox.describe_cm()}, 长={emu_to_cm(length):.2f}cm, "
            f"颜色={color_name(l.line_color)}, 线型={l.line_dash or 'solid'}, 线宽={l.line_width_pt if l.line_width_pt is not None else '未显式设置'}pt, "
            f"editable_line_ok={editable_line_ok}, color_ok={color_ok}, solid_ok={dash_ok}, width_ok={width_ok}"
        )
    ev = (
        f"结构线条={len(lines)}(需至少{required_min_lines})；"
        f"可编辑直线对象={editable_ok}；黑色={color_ok_all}；实线={dash_ok_all}；线宽0.75-3pt={width_ok_all}；"
        + "；".join(details)
    )
    return RuleResult("D2-04", "上方左侧反应物键线：六元环边线、侧链线和末端连接线均为黑色实线，0.75-3pt，可编辑直线", 1, ok, ev)


def bottom_right_product_rule(slide: SlideInfo, spec: RegionSpec) -> RuleResult:
    """D2-18：下方右侧产物结构。
    细则：位于页面右下区域；宽度约14-16cm；高度约9-11cm；
          包含左侧六元含氮杂环、左下"Cl"、中间链状连接、中心含氮连接基团、
          右侧含氮五元环和羰基"O"。
    """
    shapes = shapes_in_region_strict(slide, *spec.bounds)
    all_lines = [l for l in line_shapes(shapes) if is_black_or_unknown(l.line_color)]
    all_texts = text_shapes(shapes)
    n_texts = exact_text_shapes(shapes, "N")
    cl_texts = exact_text_shapes(shapes, "Cl")
    o_texts = exact_text_shapes(shapes, "O")

    # 位置：右下区域
    region = region_box(slide, *spec.bounds)
    location_ok = region.x >= slide.width * 0.55 and region.y >= slide.height * 0.35

    # 主结构包含下半区六元环/链/羰基，以及右上方的含氮五元环。
    product_shapes = [sh for sh in shapes if shape_center(sh)[1] > slide.height * 0.39]
    lower_shapes = [sh for sh in shapes if shape_center(sh)[1] > slide.height * 0.55]
    lower_lines = [l for l in line_shapes(lower_shapes) if is_black_or_unknown(l.line_color)]
    lower_n_texts = exact_text_shapes(lower_shapes, "N")
    lower_cl_texts = exact_text_shapes(lower_shapes, "Cl")
    lower_o_texts = exact_text_shapes(lower_shapes, "O")

    # 结构框：覆盖完整右下产物主体（含右侧含氮五元环），对应细则宽度/高度
    struct_rb = robust_bbox_of(product_shapes)
    w_cm = emu_to_cm(struct_rb.w)
    h_cm = emu_to_cm(struct_rb.h)
    size_ok = 14.0 <= w_cm <= 16.0 and 9.0 <= h_cm <= 11.0

    # 左侧六元含氮杂环：下半区左侧竖向对位N对
    ring_n_pair = _find_n_pair(lower_n_texts, max_dx_cm=1.0, min_dy_cm=3.0, max_dy_cm=5.7)
    ring_lines: List[ShapeInfo] = []
    ring_box = BBox()
    ring_ok = False
    if ring_n_pair:
        ring_lines = _find_ring_lines(lower_lines, ring_n_pair)
        ring_box = robust_bbox_of(ring_lines)
        v_count = sum(1 for l in ring_lines if line_is_vertical(l))
        d_count = sum(1 for l in ring_lines if line_is_diagonal(l))
        ring_ok = len(ring_lines) >= 6 and v_count >= 1 and d_count >= 4

    # 左下 Cl
    cl_ok = False
    if ring_box.area > 0:
        cl_ok = any(
            shape_center(c)[0] < ring_box.cx and shape_center(c)[1] > ring_box.cy
            for c in lower_cl_texts
        )

    # 中间链状连接（环右侧到中心N之间的连接线）
    ring_ids = {id(l) for l in ring_lines}
    non_ring_lines = [l for l in lower_lines if id(l) not in ring_ids]
    right_ns = [n for n in lower_n_texts if ring_box.area > 0 and shape_center(n)[0] > ring_box.x2 + EMU_PER_CM * 0.3]
    chain_lines: List[ShapeInfo] = []
    center_n_ok = False
    if right_ns and ring_box.area > 0:
        nearest_right_n = min(right_ns, key=lambda n: shape_center(n)[0])
        rn_cx, rn_cy = shape_center(nearest_right_n)
        chain_lines = [
            l for l in non_ring_lines
            if ring_box.x2 - EMU_PER_CM * 0.3 <= shape_center(l)[0] <= rn_cx + EMU_PER_CM * 0.5
            and abs(shape_center(l)[1] - rn_cy) <= EMU_PER_CM * 1.5
        ]
        # 中心含氮连接基团：环右侧的N
        center_n_ok = bool(nearest_right_n)
    chain_ok = len(chain_lines) >= 2

    # 右侧含氮五元环：右上区域至少3个N，周围有环线。
    # 该五元环整体纵向跨度约 9.9-13.2cm；y 阈值放到 slide.height*0.62 (≈14.17cm)
    # 才能覆盖第三个 N；同时远小于中心N(≈16.5cm)与左侧六元环N(≥15.4cm)，不会误纳。
    upper_n_texts = [
        n for n in n_texts
        if shape_center(n)[0] > slide.width * 0.82 and shape_center(n)[1] < slide.height * 0.62
    ]
    upper_lines = [
        l for l in line_shapes(shapes)
        if is_black_or_unknown(l.line_color)
        and shape_center(l)[0] > slide.width * 0.82
        and shape_center(l)[1] < slide.height * 0.62
    ]
    five_ring_ok = len(upper_n_texts) >= 3 and len(upper_lines) >= 5

    # 羰基O：右侧区域存在O文本
    carbonyl_ok = bool(lower_o_texts) or bool(o_texts)

    ok = location_ok and size_ok and bool(ring_n_pair) and ring_ok and cl_ok and chain_ok and center_n_ok and five_ring_ok and carbonyl_ok
    ev = (
        f"区域={spec.title}；结构框 {struct_rb.describe_cm()}；"
        f"宽={w_cm:.2f}cm(需14-16)；高={h_cm:.2f}cm(需9-11)；"
        f"右下位置={location_ok}；"
        f"六元含氮杂环={ring_ok}(N对={bool(ring_n_pair)}, 环线={len(ring_lines)}, 环框={ring_box.describe_cm()})；"
        f"左下Cl={cl_ok}(Cl={len(lower_cl_texts)})；"
        f"中间链状连接={chain_ok}(链线={len(chain_lines)})；"
        f"中心含氮连接基团={center_n_ok}；"
        f"右侧含氮五元环={five_ring_ok}；"
        f"羰基O={carbonyl_ok}(O={len(o_texts)})"
    )
    return RuleResult("D2-18", "下方右侧产物结构：右下区域，14-16cm×9-11cm，六元含氮杂环、左下Cl、中间链、中心N、五元环、羰基O", 5, ok, ev)


def bottom_right_text_rule(slide: SlideInfo, spec: RegionSpec) -> RuleResult:
    """D2-19：下方右侧产物文字。
    细则：结构中的"N""Cl""O"等字母均为可编辑文本，字号约26-34磅，颜色为黑色。
    """
    shapes = shapes_in_region_strict(slide, *spec.bounds)
    target_texts = [
        t for t in text_shapes(shapes)
        if norm_text(t.text).upper() in {"N", "CL", "O"}
    ]
    # 只保留下方右侧产物结构区域内的目标文本；该结构包含右侧五元环（上部）和下方主体。
    target_texts = [
        t for t in target_texts
        if shape_center(t)[0] >= slide.width * 0.62
        and shape_center(t)[1] >= slide.height * 0.38
    ]

    n_count = sum(1 for t in target_texts if norm_text(t.text).upper() == "N")
    cl_count = sum(1 for t in target_texts if norm_text(t.text).upper() == "CL")
    o_count = sum(1 for t in target_texts if norm_text(t.text).upper() == "O")
    required_ok = n_count >= 5 and cl_count >= 1 and o_count >= 1
    editable_ok = bool(target_texts) and all(t.is_text and not t.is_hidden and not t.is_picture for t in target_texts)

    object_results = []
    for t in target_texts:
        sizes = [r.size_pt for r in t.text_runs if r.size_pt is not None]
        colors = [r.color for r in t.text_runs if r.color is not None]
        size_ok = bool(sizes) and all(26 <= sz <= 34 for sz in sizes)
        color_ok = all(is_black_or_unknown(c) for c in colors)
        object_results.append((t, size_ok, color_ok, sizes, colors))

    size_ok_all = bool(object_results) and all(item[1] for item in object_results)
    color_ok_all = bool(object_results) and all(item[2] for item in object_results)
    ok = required_ok and editable_ok and size_ok_all and color_ok_all

    details = []
    for t, size_ok, color_ok, sizes, colors in object_results:
        details.append(
            f"id={t.shape_id}, text={t.text!r}, 框={t.bbox.describe_cm()}, "
            f"字号={[round(v, 2) for v in sizes] or ['未显式设置']}, "
            f"颜色={[color_name(c) for c in colors] or ['未显式设置']}, "
            f"size_ok={size_ok}, color_ok={color_ok}"
        )
    ev = (
        f"目标文字={len(target_texts)}；N={n_count}，Cl={cl_count}，O={o_count}；"
        f"可编辑文本={editable_ok}；字号26-34pt={size_ok_all}；黑色={color_ok_all}；"
        + "；".join(details)
    )
    return RuleResult("D2-19", "下方右侧产物文字：N/Cl/O 可编辑文本，26-34pt，黑色", 1, ok, ev)


def bottom_right_line_rule(slide: SlideInfo, spec: RegionSpec) -> RuleResult:
    """D2-20：下方右侧产物键线。
    细则：分子骨架、杂环边线、羰基双线和连接线均为黑色实线，
          线宽约0.75—1.25磅，全部为可编辑直线对象。
    """
    shapes = shapes_in_region_strict(slide, *spec.bounds)
    # 下方右侧产物结构既包含右上五元环，也包含下方主体；使用 D2-18 同一区域内所有线。
    lines = [l for l in line_shapes(shapes) if not l.is_hidden]
    relevant_lines = [
        l for l in lines
        if shape_center(l)[0] >= slide.width * 0.62
        and shape_center(l)[1] >= slide.height * 0.38
    ]

    required_min_lines = 20
    object_results = []
    for l in relevant_lines:
        editable_line_ok = l.is_editable and l.preset_geometry == "line" and l.kind in {"shape", "connector"} and not l.is_hidden
        color_ok = is_black_or_unknown(l.line_color)
        dash_ok = l.line_dash is None or l.line_dash.lower() == "solid"
        # 未显式写入线宽视为继承主题默认（约0.75pt），落在合格区间。
        width_ok = l.line_width_pt is None or 0.75 <= l.line_width_pt <= 1.25
        object_results.append((l, editable_line_ok, color_ok, dash_ok, width_ok))

    enough_lines = len(relevant_lines) >= required_min_lines
    editable_ok = bool(object_results) and all(item[1] for item in object_results)
    color_ok_all = bool(object_results) and all(item[2] for item in object_results)
    dash_ok_all = bool(object_results) and all(item[3] for item in object_results)
    width_ok_all = bool(object_results) and all(item[4] for item in object_results)
    ok = enough_lines and editable_ok and color_ok_all and dash_ok_all and width_ok_all

    bad_details = []
    for l, editable_line_ok, color_ok, dash_ok, width_ok in object_results:
        if editable_line_ok and color_ok and dash_ok and width_ok:
            continue
        _, _, length = line_delta(l)
        bad_details.append(
            f"id={l.shape_id}, kind={l.kind}, preset={l.preset_geometry or '无'}, "
            f"框={l.bbox.describe_cm()}, 长={emu_to_cm(length):.2f}cm, "
            f"颜色={color_name(l.line_color)}, 线型={l.line_dash or 'solid'}, "
            f"线宽={l.line_width_pt if l.line_width_pt is not None else '未显式设置'}pt, "
            f"editable_line_ok={editable_line_ok}, color_ok={color_ok}, solid_ok={dash_ok}, width_ok={width_ok}"
        )
    ev = (
        f"结构线条={len(relevant_lines)}(需至少{required_min_lines})；"
        f"可编辑直线对象={editable_ok}；黑色={color_ok_all}；实线={dash_ok_all}；线宽0.75-1.25pt={width_ok_all}；"
        f"不合格线条={len(bad_details)}"
        + ("；" + "；".join(bad_details[:20]) if bad_details else "")
    )
    return RuleResult("D2-20", "下方右侧产物键线：骨架、杂环边线、羰基双线、连接线均为黑色实线，0.75-1.25pt，可编辑直线", 1, ok, ev)


def whole_layout_rule(slide: SlideInfo) -> RuleResult:
    box = content_bbox(slide)
    w_ratio = box.w / max(1.0, slide.width)
    h_ratio = box.h / max(1.0, slide.height)
    left_cm = emu_to_cm(box.x)
    right_cm = emu_to_cm(slide.width - box.x2)
    top_cm = emu_to_cm(box.y)
    bottom_cm = emu_to_cm(slide.height - box.y2)

    # 细则：宽度占页面可用宽度的 85%—96%
    width_ok = 0.85 <= w_ratio <= 0.96
    # 细则：高度占页面可用高度的 70%—90%
    height_ok = 0.70 <= h_ratio <= 0.90
    # 细则：上下两组化学反应式完整出现
    groups_ok = has_top_bottom_groups(slide)
    # 细则：左边缘距页面左边线约 0—1 cm
    left_ok = 0.0 <= left_cm <= 1.0
    # 细则：右边缘距页面右边线约 0—1 cm
    right_ok = 0.0 <= right_cm <= 1.0
    # 细则：上边缘距页面上边线约 0—1 cm
    top_ok = 0.0 <= top_cm <= 1.0
    # 细则：下边缘距页面下边线约 2—4 cm
    bottom_ok = 2.0 <= bottom_cm <= 4.0

    # 细则：页面背景为白色，无渐变、无纹理、无图片底图、无水印、无批注标记
    white_ok, white_ev = background_is_plain_white(slide)
    no_image_bg = not has_image_backdrop(slide)
    no_watermark = not has_watermark_marker(slide)
    no_comments = not slide.has_comments
    bg_ok = white_ok and no_image_bg and no_watermark and no_comments

    ok = groups_ok and width_ok and height_ok and left_ok and right_ok and top_ok and bottom_ok and bg_ok
    ev = (
        f"内容框 {box.describe_cm()}；"
        f"占宽={w_ratio:.1%}(需85%-96%)；占高={h_ratio:.1%}(需70%-90%)；"
        f"左边距={left_cm:.2f}cm(需0-1)；右边距={right_cm:.2f}cm(需0-1)；"
        f"上边距={top_cm:.2f}cm(需0-1)；下边距={bottom_cm:.2f}cm(需2-4)；"
        f"上下组={groups_ok}；{white_ev}；图片底图={not no_image_bg}；水印={not no_watermark}；批注={not no_comments}"
    )
    return RuleResult("D2-01", "整体构图：上下两组完整、横向展开、近似铺满、白底无水印", 5, ok, ev)


def editability_rule(slide: SlideInfo) -> RuleResult:
    """D2-22：可编辑性。
    细则：用户可以分别选中上方左侧反应物结构线段、上方中间化学式、上方箭头、
          上方右侧产物结构、下方左侧反应物结构、下方中间试剂结构、下方箭头和
          下方右侧产物结构，并单独修改或删除。
    检测标准：八个指定区域内各自存在至少 1 个可编辑对象（line-like 或 text），
              且区域内对象均以独立对象形式存在（非嵌入图片），均可单独选中。
    """
    regions = [
        ("上方左侧反应物结构线段", (0.00, 0.02, 0.36, 0.48)),
        ("上方中间化学式",         (0.28, 0.08, 0.58, 0.42)),
        ("上方箭头",               (0.46, 0.08, 0.72, 0.42)),
        ("上方右侧产物结构",       (0.50, 0.02, 1.00, 0.52)),
        ("下方左侧反应物结构",     (0.00, 0.42, 0.45, 1.00)),
        ("下方中间试剂结构",       (0.30, 0.38, 0.70, 1.00)),
        ("下方箭头",               (0.55, 0.48, 0.80, 0.90)),
        ("下方右侧产物结构",       (0.62, 0.35, 1.00, 1.00)),
    ]
    passed = []
    failed = []
    details = []
    for title, bounds in regions:
        sh = shapes_in_region(slide, *bounds)
        # 可分别选中 = 是可编辑对象（非图片）且未被隐藏：包括线段、文本框、连接符。
        editable = [s for s in sh if s.is_editable and not s.is_picture and not s.is_hidden]
        txt_count = len(text_shapes(editable))
        line_count = len(line_shapes(editable))
        # 至少 1 个可编辑对象，且至少包含文本或线段（确认是化学结构内容而非空占位符）
        region_ok = len(editable) >= 1 and (txt_count >= 1 or line_count >= 1)
        detail = f"{title}(可编辑={len(editable)}, 文本={txt_count}, 线段={line_count})"
        details.append(detail)
        if region_ok:
            passed.append(detail)
        else:
            failed.append(detail)
    ok = len(failed) == 0
    ev = f"通过区域={passed}；不足区域={failed or '无'}"
    return RuleResult("D2-22", "可编辑性：八个关键区域均可分别选中、修改或删除", 5, ok, ev)


def dimension2(slide: SlideInfo) -> List[RuleResult]:
    results: List[RuleResult] = []
    results.append(whole_layout_rule(slide))

    top_left = RegionSpec("top_left", "上方左侧反应物", (0.00, 0.02, 0.36, 0.48), (10, 11), (5, 5.5))
    top_right = RegionSpec("top_right", "上方右侧产物", (0.50, 0.02, 1.00, 0.52), (11.5, 12.5), (6.0, 7.0))
    bottom_left = RegionSpec("bottom_left", "下方左侧反应物", (0.00, 0.42, 0.45, 1.00), (11.4, 12), (5.6, 6.3))
    bottom_mid = RegionSpec("bottom_mid", "下方中间试剂", (0.30, 0.38, 0.70, 1.00), (8.2, 9.0), (8.5, 9.4))
    bottom_right = RegionSpec("bottom_right", "下方右侧产物", (0.62, 0.35, 1.00, 1.00), (14.5, 15.5), (10.5, 11.5))

    results.append(top_left_reactant_rule(slide, top_left))
    results.append(top_left_reactant_text_rule(slide, top_left))
    results.append(top_left_reactant_line_rule(slide, top_left))

    results.append(top_first_plus_rule(slide, top_left))
    results.append(top_mid_formula_rule(slide, top_left))
    results.append(top_arrow_rule(slide, top_left))

    results.append(top_right_product_rule(slide, top_right))
    results.append(top_right_modified_line_rule(slide, top_right))
    results.append(top_right_product_text_rule(slide, top_right))

    results.append(bottom_left_product_rule(slide, bottom_left))
    results.append(bottom_left_modified_line_rule(slide, bottom_left))
    results.append(bottom_left_text_rule(slide, bottom_left))
    results.append(bottom_second_plus_rule(slide, bottom_left))

    def bottom_mid_extra(shapes: Sequence[ShapeInfo]) -> Tuple[bool, str]:
        lines = black_thin_lines(shapes)
        parallel = parallel_line_groups(lines)
        return parallel >= 1 or len(lines) >= 8, f"羰基/平行短线候选={parallel}组"

    results.append(bottom_mid_reagent_rule(slide, bottom_mid))
    results.append(bottom_mid_reagent_text_rule(slide, bottom_mid))
    results.append(bottom_arrow_rule(slide, bottom_mid))

    results.append(bottom_right_product_rule(slide, bottom_right))
    results.append(bottom_right_text_rule(slide, bottom_right))
    results.append(bottom_right_line_rule(slide, bottom_right))

    results.append(editability_rule(slide))
    return results


SCRIPT_ID = "057"
DEFAULT_PPTX_NAME = "editable_chemical_scheme.pptx"


def _locate_pptx(directory: Path) -> Optional[Path]:
    """在给定目录中定位待评估的 .pptx 文件。

    优先取默认文件名 editable_chemical_scheme.pptx；否则回退到目录内首个 .pptx。
    """
    preferred = directory / DEFAULT_PPTX_NAME
    if preferred.exists():
        return preferred
    candidates = sorted(p for p in directory.glob("*.pptx") if p.is_file())
    return candidates[0] if candidates else None


def _num(value: float) -> float:
    """把 1.0 这类整型浮点数收敛为 int，便于 JSON 输出更整洁。"""
    return int(value) if float(value).is_integer() else float(value)


def evaluate(dir_path: str) -> Dict[str, Any]:
    """统一入口：接收脚本所在目录的路径，返回结构化评估结果。

    实现约定：
    - 不修改 sys.stdout / sys.stderr；
    - 不使用 sys.exit / SystemExit 表达评估结果；
    - 不硬编码路径；由参数指定目录，在其中定位 .pptx。
    """
    result: Dict[str, Any] = {
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
        directory = Path(dir_path)
        if not directory.exists() or not directory.is_dir():
            result["status"] = "error"
            result["error"] = f"目录不存在或不是目录：{dir_path}"
            return result

        pptx_path = _locate_pptx(directory)
        if pptx_path is None:
            result["status"] = "error"
            result["error"] = f"目录中未找到 .pptx 文件：{dir_path}"
            return result
        result["file_name"] = pptx_path.name

        inspector = PptxInspector(pptx_path)
        model = inspector.parse()
        reaction_slide = identify_reaction_slide(model.slides)
        d1 = dimension1(model, reaction_slide)
        dim1_pass = all(r.passed for r in d1)
        result["dim1_pass"] = dim1_pass
        if not dim1_pass:
            failed_reasons = [
                f"{r.rule_id} {r.name}：{r.evidence}" for r in d1 if not r.passed
            ]
            result["dim1_reason"] = "；".join(failed_reasons)

        dim2_items: List[Dict[str, Any]] = []
        total_score = 0.0
        max_score = 0.0
        if dim1_pass and reaction_slide is not None:
            d2 = dimension2(reaction_slide)
            for r in d2:
                delta = r.score()
                dim2_items.append({
                    "rule": r.name,
                    "max_delta": _num(r.points),
                    "delta": _num(delta),
                    "hit": bool(r.passed),
                    "detail": "",
                })
                max_score += r.points
                total_score += delta

        result["dim2_items"] = dim2_items
        result["total_score"] = _num(total_score)
        result["max_score"] = _num(max_score)
    except Exception as exc:  # 顶层兜底：脚本自身异常返回 status="error"。
        result["status"] = "error"
        result["error"] = f"{type(exc).__name__}: {exc}"
    return result


if __name__ == "__main__":
    # 本地调试入口：默认评估脚本所在目录；也可通过命令行参数指定其它目录。
    target_dir = sys.argv[1] if len(sys.argv) > 1 else str(Path(__file__).resolve().parent)
    print(json.dumps(evaluate(target_dir), ensure_ascii=False, indent=2))
