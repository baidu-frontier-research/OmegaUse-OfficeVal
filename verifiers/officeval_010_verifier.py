#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
自动评估“技术路线图”Word交付物。

对外只暴露 evaluate(dir_path: str) -> dict：
- 入参为脚本所在目录路径，脚本自行在该目录内定位并打开被评估的 .docx；
- 返回统一结构 dict（字段见批量运行接口约定：id / file_name / status / error /
  dim1_pass / dim1_reason / dim2_items / total_score / max_score）。

说明：
- 仅支持 .docx（OOXML）；不再支持 .doc。
- Word 本身不在 OOXML 中可靠保存“页数/可打开渲染后页面”等信息；脚本会优先读取 docProps/app.xml
  的 Pages，读不到时用“无分页符、单 section、所有对象落在第一页范围内”的方式自动近似判定。
- 对于“约”“大约”的尺寸项，脚本使用少量容差，以自动评估意图为准。
"""

from __future__ import annotations

import json
import math
import re
import sys
import zipfile
from dataclasses import dataclass, field
from pathlib import Path
from typing import Callable, Iterable, Optional

from lxml import etree

PT_PER_INCH = 72.0
CM_PER_INCH = 2.54
PT_PER_CM = PT_PER_INCH / CM_PER_INCH
EMU_PER_PT = 12700.0
TWIP_PER_PT = 20.0

NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "v": "urn:schemas-microsoft-com:vml",
    "o": "urn:schemas-microsoft-com:office:office",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "wps": "http://schemas.microsoft.com/office/word/2010/wordprocessingShape",
    "ep": "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties",
}


@dataclass
class TextStyle:
    font: str = ""
    size_pt: Optional[float] = None
    bold: bool = False
    align: str = ""


@dataclass
class Shape:
    id: str
    kind: str
    text: str
    x: float
    y: float
    w: float
    h: float
    fill: str = ""
    stroke: str = ""
    weight_pt: Optional[float] = None
    filled: bool = True
    stroked: bool = True
    style: TextStyle = field(default_factory=TextStyle)
    source: str = "vml"
    # 是否真正具备被 Word/WPS 按 left/top 绝对定位渲染的条件。
    # VML v:rect/v:roundrect 等形状只有当 style 同时含
    # `mso-position-horizontal:absolute` 与 `mso-position-vertical:absolute`
    # 时，办公软件才会把它按声明坐标绘制到页面上；否则会被当作内联对象
    # 塌陷到段落起点，视觉上不出现在"应该"的位置。
    # DrawingML/wps 走 <wp:anchor> + <wp:positionH/V>，天然绝对定位，故默认 True。
    position_absolute: bool = True

    @property
    def cx(self) -> float:
        return self.x + self.w / 2

    @property
    def cy(self) -> float:
        return self.y + self.h / 2

    @property
    def right(self) -> float:
        return self.x + self.w

    @property
    def bottom(self) -> float:
        return self.y + self.h

    def cm_box(self) -> tuple[float, float, float, float]:
        return tuple(pt_to_cm(v) for v in (self.x, self.y, self.w, self.h))


@dataclass
class Line:
    id: str
    x1: float
    y1: float
    x2: float
    y2: float
    color: str = ""
    weight_pt: Optional[float] = None
    end_arrow: str = ""
    start_arrow: str = ""
    dash: str = ""

    @property
    def length(self) -> float:
        return math.hypot(self.x2 - self.x1, self.y2 - self.y1)

    @property
    def horizontal_len(self) -> float:
        return abs(self.x2 - self.x1)

    @property
    def vertical_len(self) -> float:
        return abs(self.y2 - self.y1)

    @property
    def is_horizontal(self) -> bool:
        return abs(self.y2 - self.y1) <= 2.0

    @property
    def is_vertical(self) -> bool:
        return abs(self.x2 - self.x1) <= 2.0

    @property
    def arrow_tip(self) -> tuple[float, float]:
        # VML endarrow 箭头在 to=(x2,y2) 端。
        return (self.x2, self.y2)


@dataclass
class EvaluationResult:
    dimension1_passed: bool
    score: int
    max_score: int
    dimension1_items: list[tuple[str, bool, str]]
    hits: list[tuple[int, str, str]]
    misses: list[tuple[int, str, str]]


class DocxRoadmap:
    def __init__(self, path: Path):
        self.path = path
        self.root: Optional[etree._Element] = None
        self.app_root: Optional[etree._Element] = None
        self.shapes: list[Shape] = []
        self.lines: list[Line] = []
        self.has_images = False
        self.image_rel_count = 0
        self.page_w_pt = 595.0   # A4 default
        self.page_h_pt = 842.0
        self.margin_top_pt = 72.0
        self.margin_left_pt = 72.0
        self.margin_right_pt = 72.0
        self.margin_bottom_pt = 72.0
        self.section_count = 0
        self.explicit_pages: Optional[int] = None
        self.open_error = ""
        self.loaded = False

    def load(self) -> bool:
        try:
            with zipfile.ZipFile(self.path, "r") as z:
                xml = z.read("word/document.xml")
                self.root = etree.fromstring(xml)
                try:
                    self.app_root = etree.fromstring(z.read("docProps/app.xml"))
                except Exception:
                    self.app_root = None
                self.image_rel_count = len([n for n in z.namelist() if n.startswith("word/media/")])
        except Exception as exc:
            self.open_error = str(exc)
            return False

        self.loaded = True
        self._parse_page_info()
        self._parse_images()
        self._parse_vml_lines()
        self._parse_vml_shapes()
        self._parse_wps_shapes()
        self._parse_drawingml_connectors()
        return True

    def _parse_page_info(self) -> None:
        assert self.root is not None
        sects = self.root.xpath(".//w:sectPr", namespaces=NS)
        self.section_count = len(sects)
        if sects:
            sect = sects[-1]
            pg = first(sect.xpath("./w:pgSz", namespaces=NS))
            mar = first(sect.xpath("./w:pgMar", namespaces=NS))
            if pg is not None:
                self.page_w_pt = twip_to_pt(get_float(pg, qn("w", "w"), self.page_w_pt * TWIP_PER_PT))
                self.page_h_pt = twip_to_pt(get_float(pg, qn("w", "h"), self.page_h_pt * TWIP_PER_PT))
            if mar is not None:
                self.margin_top_pt = twip_to_pt(get_float(mar, qn("w", "top"), self.margin_top_pt * TWIP_PER_PT))
                self.margin_left_pt = twip_to_pt(get_float(mar, qn("w", "left"), self.margin_left_pt * TWIP_PER_PT))
                self.margin_right_pt = twip_to_pt(get_float(mar, qn("w", "right"), self.margin_right_pt * TWIP_PER_PT))
                self.margin_bottom_pt = twip_to_pt(get_float(mar, qn("w", "bottom"), self.margin_bottom_pt * TWIP_PER_PT))

        if self.app_root is not None:
            pages = first(self.app_root.xpath(".//*[local-name()='Pages']/text()"))
            if pages and str(pages).strip().isdigit():
                self.explicit_pages = int(str(pages).strip())

    def _parse_images(self) -> None:
        assert self.root is not None
        blips = self.root.xpath(".//a:blip | .//v:imagedata", namespaces=NS)
        self.has_images = bool(blips or self.image_rel_count)

    def _parse_vml_lines(self) -> None:
        assert self.root is not None
        for el in self.root.xpath(".//v:line", namespaces=NS):
            x1, y1 = parse_pair(el.get("from"), (0.0, 0.0))
            x2, y2 = parse_pair(el.get("to"), (0.0, 0.0))
            stroke = first(el.xpath("./v:stroke", namespaces=NS))
            end_arrow = (stroke.get("endarrow") if stroke is not None else "") or el.get("endarrow", "")
            start_arrow = (stroke.get("startarrow") if stroke is not None else "") or el.get("startarrow", "")
            dash = (stroke.get("dashstyle") if stroke is not None else "") or el.get("dashstyle", "")
            self.lines.append(Line(
                id=el.get("id", ""),
                x1=x1,
                y1=y1,
                x2=x2,
                y2=y2,
                color=norm_color(el.get("strokecolor", "")),
                weight_pt=parse_length_to_pt(el.get("strokeweight", "")),
                end_arrow=end_arrow,
                start_arrow=start_arrow,
                dash=dash,
            ))

    def _parse_vml_shapes(self) -> None:
        assert self.root is not None
        query = ".//v:rect | .//v:roundrect | .//v:oval | .//v:shape"
        for el in self.root.xpath(query, namespaces=NS):
            style = parse_style(el.get("style", ""))
            x = parse_length_to_pt(style.get("left", "0")) or 0.0
            y = parse_length_to_pt(style.get("top", "0")) or 0.0
            w = parse_length_to_pt(style.get("width", "0")) or 0.0
            h = parse_length_to_pt(style.get("height", "0")) or 0.0
            kind = etree.QName(el).localname
            text = extract_text(el)
            # 只有当 VML style 里显式声明了 mso-position-horizontal:absolute
            # 与 mso-position-vertical:absolute 时，办公软件（Word/WPS）才会
            # 按 left/top 把形状绝对定位到页面上；否则会被当作内联对象堆到
            # 段落起点，位置类评分项无从谈起。这里把这一"可见前提"记录到 Shape 上，
            # 供位置/对齐相关的 self.item(...) 判断使用。
            pos_abs = (
                style.get("mso-position-horizontal", "").lower() == "absolute"
                and style.get("mso-position-vertical", "").lower() == "absolute"
            )
            shape = Shape(
                id=el.get("id", ""),
                kind=kind,
                text=text,
                x=x,
                y=y,
                w=w,
                h=h,
                fill=norm_color(el.get("fillcolor", "")),
                stroke=norm_color(el.get("strokecolor", "")),
                weight_pt=parse_length_to_pt(el.get("strokeweight", "")),
                filled=not attr_false(el.get("filled", "")),
                stroked=not attr_false(el.get("stroked", "")),
                style=parse_text_style(el),
                source="vml",
                position_absolute=pos_abs,
            )
            self.shapes.append(shape)

    def _parse_wps_shapes(self) -> None:
        assert self.root is not None
        for wsp in self.root.xpath(".//wps:wsp", namespaces=NS):
            anchor = ancestor(wsp, qn("wp", "anchor"))
            if anchor is None:
                anchor = ancestor(wsp, qn("wp", "inline"))
            x = y = w = h = 0.0
            docpr_name = ""
            if anchor is not None:
                x = emu_to_pt(get_text_float(anchor.xpath("./wp:positionH/wp:posOffset/text()", namespaces=NS), 0.0))
                y = emu_to_pt(get_text_float(anchor.xpath("./wp:positionV/wp:posOffset/text()", namespaces=NS), 0.0))
                extent = first(anchor.xpath("./wp:extent", namespaces=NS))
                if extent is not None:
                    w = emu_to_pt(get_float(extent, "cx", 0.0))
                    h = emu_to_pt(get_float(extent, "cy", 0.0))
                docpr = first(anchor.xpath("./wp:docPr", namespaces=NS))
                if docpr is not None:
                    docpr_name = docpr.get("name", "")
            geom = first(wsp.xpath(".//a:prstGeom", namespaces=NS))
            kind = geom.get("prst", "wps") if geom is not None else "wps"
            fill = first(wsp.xpath(".//wps:spPr/a:solidFill/a:srgbClr/@val", namespaces=NS)) or ""
            stroke = first(wsp.xpath(".//wps:spPr/a:ln/a:solidFill/a:srgbClr/@val", namespaces=NS)) or ""
            ln = first(wsp.xpath(".//wps:spPr/a:ln", namespaces=NS))
            weight = emu_to_pt(get_float(ln, "w", 0.0)) if ln is not None else None
            self.shapes.append(Shape(
                id=docpr_name,
                kind=kind,
                text=extract_text(wsp),
                x=x,
                y=y,
                w=w,
                h=h,
                fill=norm_color(fill),
                stroke=norm_color(stroke),
                weight_pt=weight,
                filled=True,
                stroked=True,
                style=parse_text_style(wsp),
                source="wps",
            ))

    def _parse_drawingml_connectors(self) -> None:
        """解析 DrawingML 的直线/连接符/箭头（<wps:cxnSp> 连接符，以及
        <wps:wsp> 里 prstGeom prst="line" 的直线形状），统一转成 Line 加入
        self.lines，与 VML v:line 保持同一套后续判定逻辑（颜色/箭头/端点）。

        背景：合格 Word 文档如果用"插入形状→直线/箭头"（DrawingML 而非
        VML）画技术路线图上的连接线，document.xml 里就不会出现 v:line，
        只解析 v:line 会让 doc.lines 为空/不完整，被误判为"没有引出线条"。

        端点换算（对齐 Word/WPS 实际渲染语义）：
        - <a:xfrm> 的 <a:off x y> 是外接矩形左上角，<a:ext cx cy> 是宽高；
          直线/连接符默认端点为矩形的左上角→右下角（即 (x,y)→(x+cx,y+cy)）。
        - flipH="1" 时水平翻转：起止点的 x 互换；flipV="1" 时垂直翻转：
          起止点的 y 互换。Word/WPS 按 flip 属性对角线端点做镜像绘制，
          这样才能画出"从右上到左下"等其他三个方向的直线。
        - 坐标基准：与 wps:wsp 矩形解析一致，取 <wp:anchor>/<wp:inline> 的
          positionH/positionV 作为整体偏移（EMU→pt），再叠加 xfrm 内的
          局部坐标（同样 EMU→pt），得到页面绝对坐标。
        """
        assert self.root is not None
        # <wps:cxnSp>：连接符（连线/箭头的专用形状），以及
        # <wps:wsp> 里 prstGeom prst 为直线族的，都按"直线"解析。
        line_prst = {"line", "straightConnector1", "bentConnector2", "bentConnector3"}
        nodes = self.root.xpath(".//wps:cxnSp | .//wps:wsp", namespaces=NS)
        for node in nodes:
            tag = etree.QName(node).localname
            geom = first(node.xpath(".//a:prstGeom", namespaces=NS))
            prst = geom.get("prst", "") if geom is not None else ""
            if tag != "cxnSp" and prst not in line_prst:
                continue

            anchor = ancestor(node, qn("wp", "anchor"))
            if anchor is None:
                anchor = ancestor(node, qn("wp", "inline"))
            base_x = base_y = 0.0
            if anchor is not None:
                base_x = emu_to_pt(get_text_float(anchor.xpath("./wp:positionH/wp:posOffset/text()", namespaces=NS), 0.0))
                base_y = emu_to_pt(get_text_float(anchor.xpath("./wp:positionV/wp:posOffset/text()", namespaces=NS), 0.0))

            xfrm = first(node.xpath(".//a:xfrm", namespaces=NS))
            if xfrm is None:
                continue
            off = first(xfrm.xpath("./a:off", namespaces=NS))
            ext = first(xfrm.xpath("./a:ext", namespaces=NS))
            ox = emu_to_pt(get_float(off, "x", 0.0)) if off is not None else 0.0
            oy = emu_to_pt(get_float(off, "y", 0.0)) if off is not None else 0.0
            cx = emu_to_pt(get_float(ext, "cx", 0.0)) if ext is not None else 0.0
            cy = emu_to_pt(get_float(ext, "cy", 0.0)) if ext is not None else 0.0

            x1, y1 = base_x + ox, base_y + oy
            x2, y2 = base_x + ox + cx, base_y + oy + cy
            if attr_false(xfrm.get("flipH", "0")) is False and xfrm.get("flipH") in ("1", "true"):
                x1, x2 = x2, x1
            if xfrm.get("flipV") in ("1", "true"):
                y1, y2 = y2, y1

            ln = first(node.xpath(".//a:ln", namespaces=NS))
            color = ""
            weight_pt = None
            dash = ""
            head_type = ""
            tail_type = ""
            if ln is not None:
                srgb = first(ln.xpath("./a:solidFill/a:srgbClr/@val", namespaces=NS))
                if srgb:
                    color = norm_color(srgb)
                else:
                    # 主题色配色方案（a:schemeClr）：常见流程图绿色多落在 accent 系；
                    # 没有可靠的通用 RGB 映射表时，保留主题色名，交给 is_green 里
                    # 的宽松兜底判断（схemeClr 场景較少，这里不强行伪造 RGB）。
                    scheme = first(ln.xpath("./a:solidFill/a:schemeClr/@val", namespaces=NS))
                    if scheme:
                        color = scheme
                weight_pt = emu_to_pt(get_float(ln, "w", 0.0)) or None
                prst_dash = first(ln.xpath("./a:prstDash/@val", namespaces=NS))
                dash = prst_dash or ""
                head = first(ln.xpath("./a:headEnd", namespaces=NS))
                tail = first(ln.xpath("./a:tailEnd", namespaces=NS))
                head_type = head.get("type", "") if head is not None else ""
                tail_type = tail.get("type", "") if tail is not None else ""

            docpr = first(node.xpath(".//wp:docPr", namespaces=NS))
            line_id = docpr.get("name", "") if docpr is not None else (node.get("id", "") or "")

            self.lines.append(Line(
                id=line_id,
                x1=x1,
                y1=y1,
                x2=x2,
                y2=y2,
                color=color,
                weight_pt=weight_pt,
                # DrawingML 里 tailEnd 对应线条终点(to)的箭头，headEnd 对应起点(from)；
                # 与 VML 的 endarrow(终点)/startarrow(起点) 语义一致，直接对应存放，
                # 供 arrow_is_triangle() 统一判断（triangle/stealth/oval 等均按“有箭头”处理，
                # 这里只在类型非空/非 none 时记为三角箭头同义词）。
                end_arrow=_drawingml_arrow_kind(tail_type),
                start_arrow=_drawingml_arrow_kind(head_type),
                dash=dash,
            ))

    @property
    def text(self) -> str:
        assert self.root is not None
        return extract_text(self.root)

    def text_shapes(self) -> list[Shape]:
        return [s for s in self.shapes if norm_text(s.text)]

    def visible_rects(self) -> list[Shape]:
        return [s for s in self.shapes if s.kind in {"rect", "roundrect"} and s.filled and s.stroked and s.w > 0 and s.h > 0]

    def flow_rects(self) -> list[Shape]:
        # 排除标题/通过不通过等无边框文本框，排除左侧灰色阶段标签，只统计主流程中的白底矩形框。
        return [s for s in self.visible_rects() if is_white(s.fill) and s.text and s.w >= s.h]

    def green_lines(self) -> list[Line]:
        return [l for l in self.lines if is_green(l.color)]

    def green_arrow_lines(self) -> list[Line]:
        return [l for l in self.green_lines() if l.end_arrow or l.start_arrow]

    def find_shape_text(self, expected: str, *, exact: bool = True) -> Optional[Shape]:
        exp = norm_text(expected)
        for s in self.shapes:
            got = norm_text(s.text)
            if (got == exp) if exact else (exp in got):
                return s
        return None

    def find_all_texts(self, expected: Iterable[str]) -> list[Optional[Shape]]:
        return [self.find_shape_text(t) for t in expected]


class Scorer:
    def __init__(self, doc: DocxRoadmap):
        self.doc = doc
        self.d1: list[tuple[str, bool, str]] = []
        self.hits: list[tuple[int, str, str]] = []
        self.misses: list[tuple[int, str, str]] = []
        self.max_score = 0

    def d1_check(self, name: str, ok: bool, detail: str = "") -> None:
        self.d1.append((name, bool(ok), detail))

    def item(self, points: int, desc: str, ok: bool, detail: str = "",
             *, pos_shapes: Iterable["Shape | None"] = ()) -> None:
        self.max_score += points
        # 位置类判定项的可见性前置：VML v:rect / v:roundrect 只有在 style 里同时声明
        # `mso-position-horizontal:absolute` 与 `mso-position-vertical:absolute` 时，
        # 办公软件（Word/WPS）才会按 left/top 把形状绝对定位到页面上；否则会被当作
        # 内联对象塌陷到段落起点（左上角），"对齐 / 正下方 / 居中 / 间距"等所有
        # 依赖矩形可见位置的评价，在用户实际打开的文档里都不成立，本项不给分。
        # DrawingML/wps 走 <wp:anchor>，天然绝对定位，Shape.position_absolute 默认 True。
        if pos_shapes:
            for s in pos_shapes:
                if s is None or not getattr(s, "position_absolute", True):
                    ok = False
                    extra = "位置类判定：矩形未按声明坐标定位，办公软件未渲染到应在位置"
                    detail = f"{detail} ｜ {extra}" if detail else extra
                    break
        if ok:
            self.hits.append((points, desc, detail))
        else:
            self.misses.append((points, desc, detail))

    def evaluate(self) -> EvaluationResult:
        self.evaluate_dimension1()
        if not all(ok for _, ok, _ in self.d1):
            return EvaluationResult(False, 0, self.max_score, self.d1, [], [])
        self.evaluate_dimension2()
        return EvaluationResult(True, sum(p for p, _, _ in self.hits), self.max_score, self.d1, self.hits, self.misses)

    def evaluate_dimension1(self) -> None:
        d = self.doc
        ext_ok = d.path.suffix.lower() == ".docx"
        self.d1_check("交付物为 .docx Word 文档", ext_ok, f"扩展名：{d.path.suffix}")
        self.d1_check("文件可以正常打开并解析", d.loaded, d.open_error or "可读取 word/document.xml")
        if not d.loaded:
            return

        # 当前脚本不再把页数是否为1页、连续空白页/乱码/文字重叠、流程图是否为
        # 可编辑对象作为维度一门禁。

    def evaluate_dimension2(self) -> None:
        d = self.doc
        title = d.find_shape_text("技术路线图")
        top_line = self._title_rule_line(title)
        first = self._first_row_box(top_line)
        second = self._second_row_boxes(first)
        goal = d.find_shape_text("需求要点识别与目标分解")
        diamond = d.find_shape_text("可行性确认")
        improve = d.find_shape_text("完善资料与方案")
        judge = d.find_shape_text("运行环境综合研判")
        external = d.find_shape_text("外部条件研判")
        internal = d.find_shape_text("内部能力评阅")
        efe = d.find_shape_text("EFE评阅清单")
        matrix = d.find_shape_text("适配矩阵")
        ife = d.find_shape_text("IFE评阅清单")
        path = d.find_shape_text("实施路径定位与方案设计")
        org = d.find_shape_text("组织保障与质量闭环")
        summary = d.find_shape_text("归纳与展望")

        # 细则两点，逐点校验（针对办公软件 Word/WPS 的实际渲染语义）：
        # (1) "位于页面顶端"：页面上没有其他文字对象位于标题的正上方（标题是最上方的正文对象）。
        # (2) "距离页面上边距 1.8-2.5cm"：VML 的 mso-position-vertical-relative 为 page 时，
        #     title.y（即 style 里的 top）就是办公软件相对于页面的垂直绝对距离，直接换算 cm 判断。
        title_at_top = bool(title) and not any(
            s is not title and norm_text(s.text) and s.y < title.y - 1 for s in d.shapes
        )
        title_dist_ok = bool(title) and cm_between(title.y, 1.8, 2.5)
        self.item(
            3,
            "“技术路线图”文本位于页面顶端，距离页面上边距 1.8-2.5cm",
            title_at_top and title_dist_ok,
            detail_pos(title),
            pos_shapes=[title],
        )
        # 细则四点，逐点校验（针对办公软件 Word/WPS 的实际渲染语义）：
        # (1) 字体为"黑体"：w:rFonts 的 eastAsia（中文渲染字体）为 SimHei / 黑体，
        #     Word/WPS 在正文渲染时以 eastAsia 字体绘制汉字。
        # (2) 字号为"小一"：Word/WPS 中"小一"= 24pt，w:sz 存半点值 48（parse 后 24.0）。
        # (3) 加粗：w:rPr 下存在 <w:b/>（办公软件据此显示加粗）。
        # (4) 居中显示：段落 w:pPr 的 <w:jc w:val="center"/>。
        title_font_ok = bool(title) and is_heiti_strict(title.style.font)
        title_size_ok = bool(title) and title.style.size_pt == 24.0
        title_bold_ok = bool(title) and title.style.bold
        title_center_ok = bool(title) and title.style.align == "center"
        self.item(
            3,
            "“技术路线图”字体字号为黑体、小一、加粗，居中显示",
            title_font_ok and title_size_ok and title_bold_ok and title_center_ok,
            detail_style(title),
        )
        # 细则四点，逐点校验（针对办公软件 Word/WPS 的实际渲染语义）：
        # (1) 颜色为灰色：VML v:line 的 strokecolor 为灰阶（R≈G≈B，非黑非白）。
        # (2) 距离"技术路线图"文本 0.5-0.7cm：即横线的 y 与标题文本底边（title.y + title.h）
        #     的垂直差；两者都带 mso-position-vertical-relative:page，办公软件按页面坐标渲染。
        # (3) 粗细约 1.0-1.5 磅：strokeweight 换算为 pt 后落在 [1.0, 1.5]。
        # (4) 长度约 17.5-18 cm：横线两端点距离（因是水平线，等于 |x2-x1|）换算 cm 判断。
        line_grey_ok = bool(top_line) and is_grey(top_line.color)
        line_dist_ok = bool(title and top_line) and cm_between(top_line.y1 - title.bottom, 0.5, 0.7)
        line_weight_ok = bool(top_line) and between(top_line.weight_pt, 1.0, 1.5)
        line_len_ok = bool(top_line) and cm_between(top_line.length, 17.5, 18.0, tol=0.1)
        self.item(
            1,
            "“技术路线图”下方的横向线条颜色为灰色，距离“技术路线图”文本0.5-0.7cm，粗细约1.0-1.5磅，长度约17.5-18cm",
            line_grey_ok and line_dist_ok and line_weight_ok and line_len_ok,
            detail_line(top_line),
        )

        # 细则三点，逐点校验（针对办公软件 Word/WPS 的实际渲染语义）：
        # (1) 评分对象：页面上"除'导入/核验/研判/形成/技术路线图'外"的其余矩形框——
        #     即主流程横向矩形（含标签、条件、连接框），显式按矩形内文本排除左侧灰色
        #     阶段标签（导入/核验/研判/形成）与标题文本框（技术路线图），使评分实现
        #     与评分标准原文的表述一一对应，避免依赖"竖向/无描边"的间接特征。
        # (2) 填充颜色为白色：fillcolor 为白色（RGB≈255）；办公软件按此渲染填充色。
        # (3) 边框为黑色或深灰色：strokecolor 的 RGB 三通道最大值 ≤ 90
        #     （对应 is_blackish：包含纯黑 #000000 与深灰 #444444/#555555 等），
        #     办公软件按此渲染描边颜色，肉眼可辨为"黑或深灰"。
        # (4) 细线：strokeweight ≤ 1.5pt（Word/WPS 中"细线"的直观上限）。
        excluded_texts = {norm_text(t) for t in ("导入", "核验", "研判", "形成", "技术路线图")}
        main_rects = [
            s for s in d.visible_rects()
            if s.text and norm_text(s.text) not in excluded_texts
        ]
        rect_fill_ok = bool(main_rects) and all(is_white(s.fill) for s in main_rects)
        rect_stroke_ok = bool(main_rects) and all(is_blackish(s.stroke) for s in main_rects)
        rect_thin_ok = bool(main_rects) and all(s.weight_pt is not None and s.weight_pt <= 1.5 for s in main_rects)
        self.item(
            5,
            "页面上除“导入”、“核验”、“研判”、“形成”、“技术路线图”文本矩形框外其余矩形框填充颜色为白色，边框为黑色细线",
            rect_fill_ok and rect_stroke_ok and rect_thin_ok,
            f"检测矩形框 {len(main_rects)} 个",
        )
        flow_rects = d.flow_rects()
        # 细则三点，逐点校验（针对办公软件 Word/WPS 的实际渲染语义）：
        # (1) 字体为"黑体"：w:rFonts 的 eastAsia（中文渲染字体）为 SimHei / 黑体，
        #     办公软件对汉字用 eastAsia 字体渲染。
        # (2) 字号为"小五"：Word/WPS 中"小五" = 9pt，w:sz 存半点值 18（parse 后 9.0）。
        # (3) 居中对齐：段落 w:pPr 的 <w:jc w:val="center"/>。
        text_font_ok = bool(flow_rects) and all(is_heiti_strict(s.style.font) for s in flow_rects)
        text_size_ok = bool(flow_rects) and all(s.style.size_pt == 9.0 for s in flow_rects)
        text_center_ok = bool(flow_rects) and all(s.style.align == "center" for s in flow_rects)
        self.item(
            5,
            "页面上所有矩形框内文本的字体字号为黑体小五，居中对齐",
            text_font_ok and text_size_ok and text_center_ok,
            summarize_styles(flow_rects),
        )

        # 细则两点，逐点校验（针对办公软件 Word/WPS 的实际渲染语义）：
        # (1) "位于分隔横线的正下方"：第一行矩形框水平中心与分隔横线水平中心对齐，
        #     且矩形框在横线下方（即 first.y > top_line.y1）。VML 位置属性
        #     mso-position-*-relative:page 都相对页面，办公软件按同一页面坐标渲染，
        #     故直接比较对象的 x/y 中心即可反映办公软件里的"正下方"关系。
        # (2) 距离分隔横线大约 0.3-0.7cm：矩形框上边到横线的垂直差（first.y - top_line.y1）
        #     换算 cm 后落在 [0.3, 0.7]，办公软件按页面坐标渲染此间距。
        line_cx = (top_line.x1 + top_line.x2) / 2 if top_line else 0.0
        first_below_ok = bool(first and top_line) and first.y > top_line.y1
        first_aligned_ok = bool(first and top_line) and abs(first.cx - line_cx) <= pt_from_cm(0.2)
        first_dist_ok = bool(first and top_line) and cm_between(first.y - top_line.y1, 0.3, 0.7)
        self.item(
            1,
            "第一行横向矩形框位于标题“技术路线图”分隔横线的正下方，距离“技术路线图”分隔横线大约0.3-0.7cm",
            first_below_ok and first_aligned_ok and first_dist_ok,
            detail_pos(first),
            pos_shapes=[first],
        )
        # 细则三点，逐点校验（针对办公软件 Word/WPS 的实际渲染语义）：
        # (1) "分隔横线下方大约0.3-0.7cm"：first.y - top_line.y1 换算 cm 落在 [0.3, 0.7]；
        #     两者位置都相对页面（mso-position-*-relative:page），办公软件按页面坐标渲染。
        # (2) 高度约 0.6-1cm：VML style 的 height 换算 cm 落在 [0.6, 1.0]；
        #     Word/WPS 按此值绘制矩形高度。
        # (3) 宽度约 5-6cm：VML style 的 width 换算 cm 落在 [5.0, 6.0]；
        #     Word/WPS 按此值绘制矩形宽度。
        first_dist_range_ok = bool(first and top_line) and cm_between(first.y - top_line.y1, 0.3, 0.7)
        first_h_ok = bool(first) and cm_between(first.h, 0.6, 1.0)
        first_w_ok = bool(first) and cm_between(first.w, 5.0, 6.0)
        self.item(
            1,
            "分隔横线下方大约0.3-0.7cm的第一行横向矩形框高度约0.6-1cm，宽度约5-6cm",
            first_dist_range_ok and first_h_ok and first_w_ok,
            detail_size(first),
            pos_shapes=[first],
        )
        # 细则两点，逐点校验（针对办公软件 Word/WPS 的实际渲染语义）：
        # (1) "分隔横线下方大约0.3-0.7cm"：first.y - top_line.y1 换算 cm 落在 [0.3, 0.7]；
        #     两者位置都相对页面（mso-position-*-relative:page），办公软件按页面坐标渲染。
        # (2) 矩形框内文本为"研究导入：明确目标与范围"：
        #     即 first 的 w:txbxContent 内的可见文本等于该字符串（忽略空白与中英文冒号差异）。
        first_dist_range_ok2 = bool(first and top_line) and cm_between(first.y - top_line.y1, 0.3, 0.7)
        first_text_ok = bool(first) and norm_text(first.text) == norm_text("研究导入:明确目标与范围")
        self.item(
            1,
            "分隔横线下方大约0.3-0.7cm的第一行横向矩形框内文本为“研究导入:明确目标与范围”",
            first_dist_range_ok2 and first_text_ok,
            f"实际第一行文本:{first.text if first else '未找到'}",
            pos_shapes=[first],
        )

        self.item(
            3,
            "在“研究导入:明确目标与范围”矩形框底部中心向下引出一条绿色无箭头竖线，"
            "连接到其下方约0.35cm处的水平绿色主干线；水平主干线从“政策信息梳理”框顶部"
            "中心正上方延伸至“业务流程梳理”框顶部中心正上方；再从主干线分别向四个子框"
            "顶部中心引出四条垂直向下的绿色箭头线",
            self._check_top_green_structure(first, second),
            "检查首框下接竖线、主干线范围、四条向下绿色箭头",
            pos_shapes=[first, *second],
        )
        self.item(5, "所有绿色箭头线线宽1.0pt，实线，末端小三角箭头", self._check_green_arrows(), self._green_arrow_detail())
        self.item(
            5,
            "第二行四个横向的矩形框位于同一水平线上，四条垂直向下的绿色箭头线下方，"
            "箭头的最顶端指向每个矩形框上边框的中心点",
            self._check_second_row_geometry(second),
            f"第二行候选={len(second)}",
            pos_shapes=list(second),
        )
        self.item(
            5,
            "第二行四个横向的矩形框内填充文本从左至右分别为“政策信息梳理”“资料归集校核”"
            "“参照案例对照”“业务流程梳理”，宽为2.5-2.7cm，高为0.6-0.8cm，各间距宽为0.6-0.75cm",
            self._check_second_row_text_size_spacing(),
            "政策信息梳理/资料归集校核/参照案例对照/业务流程梳理",
            pos_shapes=list(second),
        )
        self.item(
            5,
            "四个矩形框的底部中心点向下引出四条0.4-0.55cm线条，连接到其下方宽约"
            "9.90-10.5cm水平绿色主干线，绿色主干线中心点向下引出高度为0.3-0.4cm的绿色箭头线",
            self._check_second_bottom_bus(second),
            "检查四条下引线、水平主干线、中心下引箭头",
            pos_shapes=list(second),
        )

        # 细则两点，逐点校验（针对办公软件 Word/WPS 的实际渲染语义）：
        # (1) 第三行横向的矩形框位于垂直向下的绿色箭头线的正下方：
        #     存在一条绿色 v:line，垂直（x1==x2）且方向向下（y2>y1），带末端箭头（endarrow），
        #     且箭头所在的 x 与矩形上边框中心 goal.cx 精确对齐 —— 这才叫办公软件里的"正下方"。
        # (2) 箭头的最顶端指向矩形框上边框的中心点：
        #     该绿色向下箭头线的终点 (x2, y2) 与 (goal.cx, goal.y) 重合（±2pt 只吸收换算浮点误差）。
        # 位置类前置：由 self.item(..., pos_shapes=[goal]) 统一拦截。
        self.item(
            1,
            "第三行横向的矩形框位于垂直向下的绿色箭头线的正下方，箭头的最顶端指向矩形框上边框的中心点",
            bool(goal) and self._has_top_arrow_from_above(goal),
            detail_pos(goal),
            pos_shapes=[goal],
        )
        # 细则三点，逐点校验（针对办公软件 Word/WPS 的实际渲染语义）：
        # (1) 矩形框内填充文本为"需求要点识别与目标分解"：
        #     goal 的 w:txbxContent 内可见文本（归一化后）等于该字符串。
        # (2) 宽为 6-7cm：VML style 的 width 换算 cm 落在 [6.0, 7.0]。
        #     Word/WPS 按 width 直接绘制矩形宽度。
        # (3) 高为 0.7-1.7cm：VML style 的 height 换算 cm 落在 [0.7, 1.7]。
        #     Word/WPS 按 height 直接绘制矩形高度。
        goal_text_ok = bool(goal) and norm_text(goal.text) == norm_text("需求要点识别与目标分解")
        goal_w_ok = bool(goal) and cm_between(goal.w, 6.0, 7.0)
        goal_h_ok = bool(goal) and cm_between(goal.h, 0.7, 1.7)
        self.item(
            1,
            "第三行横向的矩形框内填充文本为“需求要点识别与目标分解”，宽为6-7cm，高为0.7-1.7cm",
            goal_text_ok and goal_w_ok and goal_h_ok,
            detail_size(goal),
            pos_shapes=[goal],
        )
        # 细则四点，逐点校验（针对办公软件 Word/WPS 的实际渲染语义）：
        # (1) 从"矩形框的下框线中心点"向下引出：绿色 v:line，垂直（x1==x2）、方向向下（y2>y1），
        #     起点 (x1,y1) 精确落在 (goal.cx, goal.bottom)（±2pt 只吸收 pt/cm 换算浮点误差）。
        # (2) 绿色箭头线：为绿色 v:line 且带箭头（endarrow 存在，Word/WPS 据此渲染箭头）。
        # (3) 高约 0.4-1.4cm：线条竖直长度 (y2-y1) 换算 cm 落在 [0.4, 1.4]。
        # (4) 宽约 0.03-0.1cm：v:line 的 strokeweight（换算 cm）落在 [0.03, 0.1]。
        #     Word/WPS 按 strokeweight 绘制线粗，是办公软件里可见的线条"宽度"。
        self.item(
            1,
            "第三行横向的矩形框的下框线中心点的位置向下引出一条高约0.4-1.4cm、宽约0.03-0.1cm的绿色箭头线",
            bool(goal) and self._has_down_arrow_from_bottom_center(goal, 0.4, 1.4, 0.03, 0.1),
            "检查从第三行框底部中心向下的绿色箭头",
            pos_shapes=[goal],
        )

        # 细则两点，逐点校验（针对办公软件 Word/WPS 的实际渲染语义）：
        # (1) 菱形框位于垂直向下的绿色箭头线的正下方：
        #     存在一条绿色 v:line，垂直（x1==x2）、方向向下（y2>y1）、末端带三角箭头，
        #     且该箭头线的 x 与菱形上顶点水平位置 diamond.cx 精确对齐。
        #     Word/WPS 里"菱形"由其外接矩形定义，上顶点位于外接矩形上边中点 (cx, y)，
        #     故"正下方"即 x 对齐 cx。
        # (2) 箭头的最顶端指向菱形框上顶点：
        #     绿色向下箭头线的终点 (x2, y2) 精确落在 (diamond.cx, diamond.y)
        #     （±2pt 只吸收 pt/EMU→cm 的浮点换算误差）。
        # VML 位置属性 mso-position-*-relative:page，Word/WPS 按同一页面坐标系渲染。
        self.item(
            1,
            "第四行横向的菱形框位于垂直向下的绿色箭头线的正下方，箭头的最顶端指向菱形框上顶点",
            bool(diamond) and self._has_top_arrow_from_above(diamond),
            detail_pos(diamond),
            pos_shapes=[diamond],
        )
        # 细则三点，逐点校验（针对办公软件 Word/WPS 的实际渲染语义）：
        # (1) 菱形框内填充文本为"可行性确认"：diamond 的 w:txbxContent 内可见文本
        #     （归一化后）等于该字符串。Word/WPS 渲染菱形内文本即取自该文本节点。
        # (2) 高为 1.5-2cm：菱形外接矩形高度（wps:spPr/a:xfrm/a:ext cy 或 VML style height）
        #     换算 cm 落在 [1.5, 2.0]。Word/WPS 按该 ext 绘制菱形高度。
        # (3) 宽为 2.7-3cm：菱形外接矩形宽度换算 cm 落在 [2.7, 3.0]。
        diamond_text_ok = bool(diamond) and norm_text(diamond.text) == norm_text("可行性确认")
        diamond_h_ok = bool(diamond) and cm_between(diamond.h, 1.5, 2.0)
        diamond_w_ok = bool(diamond) and cm_between(diamond.w, 2.7, 3.0)
        self.item(
            1,
            "第四行菱形框内填充文本为“可行性确认”，高为1.5-2cm，宽为2.7-3cm",
            diamond_text_ok and diamond_h_ok and diamond_w_ok,
            detail_size(diamond),
            pos_shapes=[diamond],
        )
        # 细则五点，逐点校验（针对办公软件 Word/WPS 的实际渲染语义）：
        # (1) 从菱形右顶点向右引出：绿色 v:line，水平（y1==y2）、方向向右（x2>x1）、末端带箭头；
        #     起点 (x1,y1) 精确落在菱形右顶点 (diamond.right, diamond.cy)。
        #     Word/WPS 里菱形由外接矩形定义，右顶点位于外接矩形右边中点 (right, cy)。
        # (2) 该箭头线高（strokeweight）0.03-0.1cm：办公软件按 strokeweight 绘制线粗，
        #     即线条视觉上的"高度/粗细"。
        # (3) 宽（水平长度）2-3cm：|x2-x1| 换算 cm 落在 [2, 3]。
        # (4) 箭头右顶点（终点）连接"完善资料与方案"矩形左边框中心点：
        #     线的终点 (x2,y2) 精确落在 (improve.x, improve.cy)。
        # (5) "完善资料与方案"矩形高 0.7-1cm、宽 2.7-3cm：improve 的 h/w 换算 cm 分别落在
        #     [0.7, 1.0] 与 [2.7, 3.0]。VML style 的 width/height 就是 Word/WPS 绘制的尺寸。
        self.item(
            1,
            "第四行菱形框右顶点处向右引出一条高为0.03-0.1cm、宽为2-3cm的绿色箭头线，"
            "箭头右顶点连接填充文本为“完善资料与方案”矩形框的左边框中心点的位置，"
            "高为0.7-1cm，宽为2.7-3cm",
            bool(diamond) and bool(improve)
            and self._has_right_arrow_from_diamond_to_box(diamond, improve, 2.0, 3.0, 0.03, 0.1)
            and cm_between(improve.h, 0.7, 1.0)
            and cm_between(improve.w, 2.7, 3.0),
            "检查右向连接箭头 + 目标矩形尺寸",
            pos_shapes=[diamond, improve],
        )
        # 细则四点，逐点校验（针对办公软件 Word/WPS 的实际渲染语义）：
        # (1) 菱形框右侧有一矩形框：improve 与 diamond 都相对页面定位
        #     （mso-position-*-relative:page），improve.x >= diamond.right 即办公软件里
        #     该矩形位于菱形右侧（视觉上左边缘不越过菱形右边缘）。
        # (2) 高为 0.6-1cm：improve.h（VML style 的 height）换算 cm 落在 [0.6, 1.0]。
        # (3) 宽为 2.5-3cm：improve.w（VML style 的 width）换算 cm 落在 [2.5, 3.0]。
        # (4) 内填充文本为"完善资料与方案"：improve 的 w:txbxContent 内可见文本
        #     （归一化后）等于该字符串。
        improve_right_ok = bool(improve and diamond) and improve.x >= diamond.right
        improve_h_ok = bool(improve) and cm_between(improve.h, 0.6, 1.0)
        improve_w_ok = bool(improve) and cm_between(improve.w, 2.5, 3.0)
        improve_text_ok = bool(improve) and norm_text(improve.text) == norm_text("完善资料与方案")
        self.item(
            1,
            "第四行菱形框右侧有一矩形框，高为0.6-1cm，宽为2.5-3cm，内填充文本为“完善资料与方案”",
            improve_right_ok and improve_h_ok and improve_w_ok and improve_text_ok,
            detail_size(improve),
            pos_shapes=[improve, diamond],
        )
        self.item(
            3,
            "“完善资料与方案”矩形框位于“可行性确认”矩形框右侧，框内文字居中，"
            "并通过绿色回路箭头返回上方“业务流程梳理”矩形框，箭头指向“业务流程梳理”"
            "矩形右侧中心位置，箭头水平方向长度约为0.8-1.2cm，竖直方向长度约为5.2-5.6cm",
            self._check_return_loop(improve, diamond),
            "检查右侧位置、居中、回路箭头(水平0.8-1.2cm、竖直5.2-5.6cm)、箭头指向业务流程梳理右侧中心",
            pos_shapes=[improve, diamond],
        )

        # 细则五点，逐点校验（针对办公软件 Word/WPS 的实际渲染语义）：
        # (1) 第五行横向矩形框居中放置：judge.cx 等于页面水平中心 page_w_pt/2；
        #     VML 用 mso-position-horizontal-relative:page，办公软件按 left+width/2
        #     绘制矩形中心，等于页面中心即视觉居中。
        # (2) 内填充文本为"运行环境综合研判"：judge 的 w:txbxContent 归一化后精确匹配。
        # (3) 高为 0.6-1cm：judge.h（VML height）换算 cm 落在 [0.6, 1.0]。
        # (4) 宽为 6.3-6.7cm：judge.w（VML width）换算 cm 落在 [6.3, 6.7]。
        # (5) 第四行菱形框下有一绿色箭头指向第五行矩形框中央，长度约 0.6-0.8cm：
        #     绿色 v:line，垂直、方向向下、末端带三角箭头；起点 (x1,y1) 精确落在
        #     菱形下顶点 (diamond.cx, diamond.bottom)，终点 (x2,y2) 精确落在
        #     矩形上边框中心 (judge.cx, judge.y)；竖直长度换算 cm 落在 [0.6, 0.8]。
        judge_center_ok = bool(judge) and abs(judge.cx - self.doc.page_w_pt / 2) <= 2
        judge_text_ok = bool(judge) and norm_text(judge.text) == norm_text("运行环境综合研判")
        judge_h_ok = bool(judge) and cm_between(judge.h, 0.6, 1.0)
        judge_w_ok = bool(judge) and cm_between(judge.w, 6.3, 6.7)
        judge_arrow_ok = bool(judge and diamond) and self._has_down_arrow_from_diamond_to_top(diamond, judge, 0.6, 0.8)
        self.item(
            3,
            "第五行横向矩形框居中放置，内填充文本为“运行环境综合研判”，高为0.6-1cm，宽为6.3-6.7cm，"
            "第四行菱形框下有一绿色箭头指向第五行矩形框中央，长度约0.6-0.8cm",
            judge_center_ok and judge_text_ok and judge_h_ok and judge_w_ok and judge_arrow_ok,
            detail_size(judge),
            pos_shapes=[judge, diamond],
        )
        self.item(
            3,
            "“运行环境综合研判”矩形框下方有一绿色箭头分别连接到第六行两个矩形框，"
            "水平方向长度约为7.5-8cm，竖直方向长度约为0.6-1cm",
            self._check_judge_to_sixth(judge, external, internal),
            "检查从judge底部中心下引、水平主干7.5-8cm、两条向下竖箭头0.6-1cm指向external/internal上边框中心",
            pos_shapes=[judge, external, internal],
        )
        # 细则五点，逐点校验（针对办公软件 Word/WPS 的实际渲染语义）：
        # (1) 第六行出现两个横向文本框：external 与 internal 均已定位到；y 相等表示两者
        #     位于同一水平行（VML top 决定办公软件里的上边框位置）。
        # (2) 从左往右数第一个文本框内填充文本为"外部条件研判"：按 x 升序，第一个的
        #     w:txbxContent 归一化后精确匹配"外部条件研判"。
        # (3) 第二个文本框内填充文本为"内部能力评阅"：同上，第二个精确匹配"内部能力评阅"。
        # (4) 文本框宽为 3-3.5cm、高为 0.6-1cm：VML width / height 换算 cm 分别落在
        #     [3.0, 3.5] 与 [0.6, 1.0]。Word/WPS 按 width/height 直接绘制矩形尺寸。
        # (5) 两个文本框间距为 4.5-5cm：即右矩形 x 减左矩形 right（left+width）
        #     换算 cm 落在 [4.5, 5.0]；两者位置都相对页面（mso-position-*-relative:page），
        #     办公软件按同一页面坐标渲染此间距。
        sixth_ok = False
        if external and internal:
            left_box, right_box = (external, internal) if external.x <= internal.x else (internal, external)
            first_text_ok = norm_text(left_box.text) == norm_text("外部条件研判")
            second_text_ok = norm_text(right_box.text) == norm_text("内部能力评阅")
            dims_ok = (
                cm_between(external.w, 3.0, 3.5) and cm_between(internal.w, 3.0, 3.5)
                and cm_between(external.h, 0.6, 1.0) and cm_between(internal.h, 0.6, 1.0)
            )
            gap_ok = cm_between(right_box.x - left_box.right, 4.5, 5.0)
            sixth_ok = first_text_ok and second_text_ok and dims_ok and gap_ok
        self.item(
            3,
            "第六行出现两个横向文本框，从左往右数第一个文本框内填充文本为“外部条件研判”，"
            "第二个文本框内填充文本为“内部能力评阅”，文本框宽为3-3.5cm，高为0.6-1cm，"
            "两个文本框间距为4.5-5cm",
            sixth_ok,
            f"外部={detail_size(external)}；内部={detail_size(internal)}",
            pos_shapes=[external, internal],
        )
        self.item(
            5,
            "第六行左侧横向文本框下方有一绿色箭头分别连接到第七行左侧三个竖向文本框（水平2.5-3cm，竖直0.6-1cm）；"
            "第六行右侧横向文本框下方有一绿色箭头分别连接到第七行右侧四个竖向文本框（水平4-4.5cm，竖直0.6-1cm）",
            self._check_sixth_to_seventh(external, internal),
            "检查两组分支线水平/竖直长度及各箭头对准第七行竖向框顶部中心",
            pos_shapes=[external, internal],
        )
        self.item(
            5,
            "第七行左侧三个竖向文本框从左往右依次为“政策环境扫描”“需求场景识别”“协同路径分析”，"
            "右侧四个竖向文本框从左往右依次为“资源储备评估”“流程执行能力”“技术支撑能力”“运营协调机制”；"
            "第3与第4个文本框间隔3.5-4cm，其余间隔0.8-1cm；文本框宽0.6-1cm、高2.5-3cm；文字都在框内，竖向排布",
            self._check_seventh_row(),
            "检查7个竖向文本框",
            pos_shapes=self.doc.find_all_texts([
                "政策环境扫描", "需求场景识别", "协同路径分析",
                "资源储备评估", "流程执行能力", "技术支撑能力", "运营协调机制",
            ]),
        )

        self.item(
            3,
            "第八行有三个横向文本框，从左往右依次为“EFE评阅清单”“适配矩阵”“IFE评阅清单”，"
            "文本框宽2.5-3cm、高0.6-1cm；中间文本框左侧和右侧出现箭头分别指向两侧文本框，箭头长度0.2-0.5cm",
            self._check_eighth_row(efe, matrix, ife),
            "EFE评阅清单/适配矩阵/IFE评阅清单",
            pos_shapes=[efe, matrix, ife],
        )
        self.item(
            1,
            "“需求场景识别”文本框下方有一倾斜箭头指向“EFE评阅清单”文本框中心；"
            "“技术支撑能力”文本框下方有一倾斜箭头指向“IFE评阅清单”文本框中心，箭头长度约0.8-1.2cm",
            self._check_diagonal_arrows(),
            "检查两条倾斜箭头",
            pos_shapes=self.doc.find_all_texts(["需求场景识别", "技术支撑能力", "EFE评阅清单", "IFE评阅清单"]),
        )
        # 第九行细则逐点校验（针对办公软件 Word/WPS 的 VML 渲染语义）：
        #  (1) 第九行横向矩形框存在，内填充文本为"实施路径定位与方案设计"
        #      —— 由 find_shape_text 精确匹配 w:txbxContent。
        #  (2) 居中放置 —— shape 中心 x 与页面中心 x 相等 (±2pt)；
        #      办公软件按 mso-position-horizontal-relative:page 的 left+w/2 判断。
        #  (3) 高为 0.6-1cm —— VML height 换算 cm ∈ [0.6, 1.0]（无 tol）。
        #  (4) 宽为 6.3-6.7cm —— VML width 换算 cm ∈ [6.3, 6.7]（无 tol）。
        #  (5) "适配矩阵"文本框下方有一绿色箭头指向第九行横向文本框中心 ——
        #      存在一条绿色垂直 v:line，起点在 matrix 底部中心 (matrix.cx, matrix.bottom)
        #      ±2pt，tip (x2, y2) 命中 (path.cx, path.y) ±2pt，end_arrow≠""。
        #  (6) 长度为 0.6-1cm —— 上述垂直线 vertical_len ∈ [0.6, 1.0]cm（无 tol）。
        ninth_ok = False
        if path and matrix:
            page_center = self.doc.page_w_pt / 2
            centered = abs(path.cx - page_center) <= 2
            size_ok = cm_between(path.h, 0.6, 1.0) and cm_between(path.w, 6.3, 6.7)
            arrow_ok = False
            for l in self.doc.green_lines():
                if not (l.is_vertical and l.end_arrow):
                    continue
                if abs(l.x1 - matrix.cx) > 2 or abs(l.y1 - matrix.bottom) > 2:
                    continue
                if abs(l.x2 - path.cx) > 2 or abs(l.y2 - path.y) > 2:
                    continue
                if not cm_between(l.vertical_len, 0.6, 1.0):
                    continue
                arrow_ok = True
                break
            ninth_ok = centered and size_ok and arrow_ok
        self.item(
            3,
            "第九行横向矩形框居中放置，内填充文本为“实施路径定位与方案设计”，高0.6-1cm、宽6.3-6.7cm；"
            "“适配矩阵”文本框下方有一绿色箭头指向第九行横向文本框中心，长度0.6-1cm",
            ninth_ok,
            detail_size(path),
            pos_shapes=[path, matrix],
        )
        # 第十行细则逐点校验（针对办公软件 Word/WPS 的 VML 渲染语义）：
        #  (1) 第十行横向矩形框存在，内填充文本为"组织保障与质量闭环"
        #      —— 由 find_shape_text 精确匹配 w:txbxContent。
        #  (2) 居中放置 —— shape 中心 x 与页面中心 x 相等 (±2pt)；
        #      办公软件按 mso-position-horizontal-relative:page 的 left+w/2 判断。
        #  (3) 高为 0.6-1cm —— VML height 换算 cm ∈ [0.6, 1.0]（无 tol）。
        #  (4) 宽为 6.3-6.7cm —— VML width 换算 cm ∈ [6.3, 6.7]（无 tol）。
        #  (5) "实施路径定位与方案设计"文本框下方有一绿色箭头指向第十行横向文本框中心
        #      —— 存在一条绿色垂直 v:line，起点 = (path.cx, path.bottom) ±2pt，
        #      尖端 (x2, y2) 命中 (org.cx, org.y) ±2pt，end_arrow≠""。
        #  (6) 长度为 0.5-0.8cm —— 上述垂直线 vertical_len ∈ [0.5, 0.8]cm（无 tol）。
        tenth_ok = False
        if org and path:
            page_center = self.doc.page_w_pt / 2
            centered = abs(org.cx - page_center) <= 2
            size_ok = cm_between(org.h, 0.6, 1.0) and cm_between(org.w, 6.3, 6.7)
            arrow_ok = False
            for l in self.doc.green_lines():
                if not (l.is_vertical and l.end_arrow):
                    continue
                if abs(l.x1 - path.cx) > 2 or abs(l.y1 - path.bottom) > 2:
                    continue
                if abs(l.x2 - org.cx) > 2 or abs(l.y2 - org.y) > 2:
                    continue
                if not cm_between(l.vertical_len, 0.5, 0.8):
                    continue
                arrow_ok = True
                break
            tenth_ok = centered and size_ok and arrow_ok
        self.item(
            3,
            "第十行横向矩形框居中放置，内填充文本为“组织保障与质量闭环”，高0.6-1cm、宽6.3-6.7cm；"
            "“实施路径定位与方案设计”文本框下方有一绿色箭头指向第十行横向文本框中心，长度0.5-0.8cm",
            tenth_ok,
            detail_size(org),
            pos_shapes=[org, path],
        )
        # 第十一行细则逐点校验（针对办公软件 Word/WPS 的 VML 渲染语义）：
        #  (1) 第十一行横向矩形框存在，内填充文本为"归纳与展望"
        #      —— 由 find_shape_text 精确匹配 w:txbxContent。
        #  (2) 居中放置 —— shape 中心 x 与页面中心 x 相等 (±2pt)；
        #      办公软件按 mso-position-horizontal-relative:page 的 left+w/2 判断。
        #  (3) 高为 0.6-1cm —— VML height 换算 cm ∈ [0.6, 1.0]（无 tol）。
        #  (4) 宽为 6.3-6.7cm —— VML width 换算 cm ∈ [6.3, 6.7]（无 tol）。
        #  (5) "组织保障与质量闭环"文本框下方有一绿色箭头指向第十一行横向文本框中心
        #      —— 存在一条绿色垂直 v:line，起点 = (org.cx, org.bottom) ±2pt，
        #      尖端 (x2, y2) 命中 (summary.cx, summary.y) ±2pt，end_arrow≠""。
        #  (6) 长度为 0.5-0.8cm —— 上述垂直线 vertical_len ∈ [0.5, 0.8]cm（无 tol）。
        eleventh_ok = False
        if summary and org:
            page_center = self.doc.page_w_pt / 2
            centered = abs(summary.cx - page_center) <= 2
            size_ok = cm_between(summary.h, 0.6, 1.0) and cm_between(summary.w, 6.3, 6.7)
            arrow_ok = False
            for l in self.doc.green_lines():
                if not (l.is_vertical and l.end_arrow):
                    continue
                if abs(l.x1 - org.cx) > 2 or abs(l.y1 - org.bottom) > 2:
                    continue
                if abs(l.x2 - summary.cx) > 2 or abs(l.y2 - summary.y) > 2:
                    continue
                if not cm_between(l.vertical_len, 0.5, 0.8):
                    continue
                arrow_ok = True
                break
            eleventh_ok = centered and size_ok and arrow_ok
        self.item(
            3,
            "第十一行横向矩形框居中放置，内填充文本为“归纳与展望”，高0.6-1cm、宽6.3-6.7cm；"
            "“组织保障与质量闭环”文本框下方有一绿色箭头指向第十一行横向文本框中心，长度0.5-0.8cm",
            eleventh_ok,
            detail_size(summary),
            pos_shapes=[summary, org],
        )
        self.item(
            5,
            "页面左侧出现四个灰色背景填充的竖向文本框，高为2.3-2.6cm,宽为0.6-1cm,从上往下填充文本依次为“导入”、“核验”、“研判”、“形成”，第一个文本框和第二个文本框的间距为1.2-1.4cm,第二个文本框和第三个文本框的间距为2.4-2.7cm，第三个文本框和第四个文本框的间距为5.8-6.2cm，字体加粗，距离页面左边界2.3-2.7cm",
            self._check_stage_boxes(),
            "导入/核验/研判/形成",
            pos_shapes=self.doc.find_all_texts(["导入", "核验", "研判", "形成"]),
        )

    def _title_rule_line(self, title: Optional[Shape]) -> Optional[Line]:
        candidates = [l for l in self.doc.lines if l.is_horizontal and not is_green(l.color)]
        if title:
            candidates = [l for l in candidates if l.y1 > title.y]
            candidates.sort(key=lambda l: abs(l.y1 - title.bottom))
        else:
            candidates.sort(key=lambda l: l.y1)
        return first(candidates)

    def _first_row_box(self, top_line: Optional[Line]) -> Optional[Shape]:
        rects = [s for s in self.doc.flow_rects() if s.w >= s.h]
        if top_line:
            rects = [s for s in rects if s.y > top_line.y1]
        rects.sort(key=lambda s: s.y)
        return first(rects)

    def _second_row_boxes(self, first: Optional[Shape]) -> list[Shape]:
        """按 rubric 明确写出的四个文本定位第二行矩形框，不依赖固定 y 坐标：
        "政策信息梳理""资料归集校核""参照案例对照""业务流程梳理"。
        为了让"位于同一水平线上""从主干线向下引箭头"等相对关系判定依旧成立，
        额外要求这四个框 y 一致、且位于第一行框（研究导入:明确目标与范围）下方——
        这两点本身就是 rubric 里"第二行"的语义，不是外加的坐标假设。"""
        texts = ["政策信息梳理", "资料归集校核", "参照案例对照", "业务流程梳理"]
        found = self.doc.find_all_texts(texts)
        if any(b is None for b in found):
            return []
        boxes = [b for b in found if b is not None]
        if max(b.y for b in boxes) - min(b.y for b in boxes) > 2:
            return []
        if first is not None and not all(b.y > first.bottom for b in boxes):
            return []
        return sorted(boxes, key=lambda s: s.x)


    def _check_top_green_structure(self, first: Optional[Shape], second: list[Shape]) -> bool:
        # 细则四点，逐点校验（针对办公软件 Word/WPS 的实际渲染语义）：
        # (1) 在"研究导入:明确目标与范围"矩形框底部中心向下引出一条绿色无箭头竖线：
        #     绿色 v:line，垂直，起点 (x1,y1) ≈ (first.cx, first.bottom)，无 endarrow/startarrow。
        # (2) 该竖线连接到其下方约 0.35cm 处的水平绿色主干线：
        #     竖线的终点 y2 与水平主干线的 y1 一致，且 (y2 - first.bottom) 约 0.35cm。
        # (3) 水平主干线从"政策信息梳理"框顶部中心正上方延伸至"业务流程梳理"框顶部中心正上方：
        #     水平绿色 v:line，x1 ≈ second[0].cx，x2 ≈ second[3].cx。
        # (4) 从主干线分别向四个子框顶部中心引出四条垂直向下的绿色箭头线：
        #     每个子框对应一条绿色 v:line，垂直，x ≈ box.cx，y1 ≈ 主干线 y，y2 ≈ box.y，
        #     且带箭头（end_arrow 为三角）。
        # 位置属性均相对页面（mso-position-*-relative:page），Word/WPS 按同一页面坐标渲染。
        if not first or len(second) != 4:
            return False
        if norm_text(first.text) != norm_text("研究导入:明确目标与范围"):
            return False

        lines = self.doc.green_lines()

        # (1)+(2): 底部中心向下的无箭头竖线，且末端落在 ≈0.35cm 下方
        down_line = None
        for l in lines:
            if not l.is_vertical:
                continue
            if l.end_arrow or l.start_arrow:
                continue
            if abs(l.x1 - first.cx) > 2 or abs(l.y1 - first.bottom) > 2:
                continue
            if l.y2 <= l.y1:
                continue
            if cm_between(l.y2 - first.bottom, 0.30, 0.40):
                down_line = l
                break
        if down_line is None:
            return False

        # (3): 水平主干线从 second[0].cx 到 second[3].cx，y 与 down_line.y2 一致
        boxes = sorted(second, key=lambda s: s.x)
        bus = None
        for l in lines:
            if not l.is_horizontal:
                continue
            if abs(l.y1 - down_line.y2) > 2:
                continue
            lx1, lx2 = sorted([l.x1, l.x2])
            if abs(lx1 - boxes[0].cx) > 2 or abs(lx2 - boxes[3].cx) > 2:
                continue
            bus = l
            break
        if bus is None:
            return False

        # (4): 从主干线到每个子框顶部中心的四条绿色向下箭头
        for b in boxes:
            hit = False
            for l in lines:
                if not l.is_vertical:
                    continue
                if not arrow_is_triangle(l.end_arrow):
                    continue
                if l.y2 <= l.y1:
                    continue
                if abs(l.x1 - b.cx) > 2 or abs(l.x2 - b.cx) > 2:
                    continue
                if abs(l.y1 - bus.y1) > 2 or abs(l.y2 - b.y) > 2:
                    continue
                hit = True
                break
            if not hit:
                return False
        return True

    def _check_green_arrows(self) -> bool:
        # 细则三点，逐点校验（针对办公软件 Word/WPS 的实际渲染语义）：
        # (1) 线宽 1.0pt：VML v:line 的 strokeweight 精确等于 1.0pt。
        #     Word/WPS 按 strokeweight 直接绘制线宽。
        # (2) 实线：v:stroke 的 dashstyle 为空或 "solid"（未设置时默认实线）。
        #     Word/WPS 按 dashstyle 决定线型。
        # (3) 末端小三角箭头：v:stroke 的 endarrow 为 block/triangle/classic 之一。
        #     这三个值在 Word/WPS 中都渲染为末端三角形箭头。
        arrows = self.doc.green_arrow_lines()
        if not arrows:
            return False
        for l in arrows:
            if l.weight_pt != 1.0:
                return False
            if str(l.dash).strip().lower() not in {"", "solid"}:
                return False
            if not arrow_is_triangle(l.end_arrow):
                return False
        return True

    def _green_arrow_detail(self) -> str:
        arrows = self.doc.green_arrow_lines()
        bad = [l.id for l in arrows if not (l.weight_pt == 1.0 and arrow_is_triangle(l.end_arrow or l.start_arrow))]
        return f"绿色箭头线 {len(arrows)} 条，不合格：{bad[:8]}"

    def _check_second_row_geometry(self, boxes: list[Shape]) -> bool:
        # 细则三点，逐点校验（针对办公软件 Word/WPS 的实际渲染语义）：
        # (1) 第二行四个横向矩形框位于同一水平线上：四个矩形的 y（上边框位置）一致；
        #     VML v:rect 的 top 决定办公软件里矩形的上边框位置，y 相等即可见上边缘对齐。
        # (2) 位于四条垂直向下的绿色箭头线下方：每个矩形上方存在一条对应的绿色 v:line，
        #     垂直（x1==x2），方向向下（y2>y1），末端带箭头（endarrow 为三角）。
        # (3) 箭头的最顶端指向每个矩形框上边框的中心点：
        #     该绿色箭头线的终点 (x2,y2) 与矩形上边框中心 (box.cx, box.y) 重合。
        #     位置属性均相对页面（mso-position-*-relative:page），Word/WPS 按同一页面坐标渲染。
        if len(boxes) != 4:
            return False
        same_line = max(b.y for b in boxes) == min(b.y for b in boxes)
        if not same_line:
            return False
        for b in boxes:
            hit = False
            for l in self.doc.green_lines():
                if not l.is_vertical:
                    continue
                if l.y2 <= l.y1:
                    continue
                if not arrow_is_triangle(l.end_arrow):
                    continue
                if abs(l.x2 - b.cx) > 2 or abs(l.y2 - b.y) > 2:
                    continue
                hit = True
                break
            if not hit:
                return False
        return True

    def _has_top_arrow_from_above(self, box: Shape) -> bool:
        # 细则口径的"位于垂直向下的绿色箭头线的正下方，箭头最顶端指向上边框中心点"：
        # 存在一条绿色 v:line，垂直、方向向下、末端带三角箭头，
        # 且箭头端点 (x2, y2) 精确落在 (box.cx, box.y)（±2pt 只吸收 pt/cm 换算浮点误差）。
        # 另外，该箭头必须"从上游可见对象延伸而来"——即起点 (x1, y1) 精确落在
        # 某个 position_absolute=True（即办公软件里真正被绘制到应在位置）的
        # shape 底部中心（±2pt）。否则该箭头在实际渲染中是一段与流程脱节的
        # 悬空碎片，"位于……正下方"的流程语义不成立，本项不应给分。
        for l in self.doc.green_lines():
            if not l.is_vertical or l.y2 <= l.y1:
                continue
            if not arrow_is_triangle(l.end_arrow):
                continue
            if abs(l.x2 - box.cx) > 2 or abs(l.y2 - box.y) > 2:
                continue
            # 起点必须锚定在一个"办公软件里真正可见"的 shape 的底部中心。
            anchored = False
            for src in self.doc.shapes:
                if src is box:
                    continue
                if not getattr(src, "position_absolute", True):
                    continue
                if abs(l.x1 - src.cx) <= 2 and abs(l.y1 - src.bottom) <= 2:
                    anchored = True
                    break
            if not anchored:
                continue
            return True
        return False

    def _has_qualified_arrow_to_top(self, box: Shape) -> bool:
        for line in self.doc.green_arrow_lines():
            if (
                line.is_vertical
                and line.y2 > line.y1
                and line.weight_pt == 1.0
                and line.dash in {"", "solid"}
                and arrow_is_triangle(line.end_arrow)
                and abs(line.x2 - box.cx) <= 2
                and abs(line.y2 - box.y) <= 2
            ):
                return True
        return False

    def _check_second_row_text_size_spacing(self) -> bool:
        # 细则四点，逐点校验（针对办公软件 Word/WPS 的实际渲染语义）：
        # (1) 文本从左至右分别为"政策信息梳理""资料归集校核""参照案例对照""业务流程梳理"：
        #     按矩形 x（VML style 的 left，即办公软件里左边缘位置）升序排列后，
        #     四个矩形内 w:txbxContent 的可见文本依次等于目标字符串。
        # (2) 宽为 2.5-2.7cm：VML style 的 width 换算 cm 落在 [2.5, 2.7]。
        # (3) 高为 0.6-0.8cm：VML style 的 height 换算 cm 落在 [0.6, 0.8]。
        # (4) 各间距宽为 0.6-0.75cm：相邻两个矩形之间的间距，即右矩形 x 减左矩形右边缘
        #     （left+width）换算 cm 落在 [0.6, 0.75]。位置属性均相对页面
        #     （mso-position-*-relative:page），办公软件按同一页面坐标渲染这些距离。
        texts = ["政策信息梳理", "资料归集校核", "参照案例对照", "业务流程梳理"]
        # 依据"矩形框内文本"来定位：从主流程可见矩形里找文本匹配的四个框。
        candidates = [s for s in self.doc.visible_rects() if s.text]
        found: list[Shape] = []
        for t in texts:
            hit = next((s for s in candidates if norm_text(s.text) == norm_text(t)), None)
            if hit is None:
                return False
            found.append(hit)
        boxes2 = sorted(found, key=lambda s: s.x)
        if [norm_text(b.text) for b in boxes2] != [norm_text(t) for t in texts]:
            return False
        if not all(cm_between(b.w, 2.5, 2.7) for b in boxes2):
            return False
        if not all(cm_between(b.h, 0.6, 0.8) for b in boxes2):
            return False
        gaps = [boxes2[i + 1].x - boxes2[i].right for i in range(3)]
        if not all(cm_between(g, 0.6, 0.75) for g in gaps):
            return False
        return True

    def _check_second_bottom_bus(self, boxes: list[Shape]) -> bool:
        # 细则三点，逐点校验（针对办公软件 Word/WPS 的实际渲染语义）：
        # (1) 四个矩形框的底部中心点向下引出四条 0.4-0.55cm 线条：
        #     对每个矩形，存在一条绿色 v:line，垂直（x1==x2）、方向向下（y2>y1），
        #     起点 (x1,y1) == (b.cx, b.bottom)，长度 (y2-y1) 换算 cm 落在 [0.4, 0.55]。
        # (2) 连接到其下方宽约 9.90-10.5cm 水平绿色主干线：
        #     存在一条绿色水平 v:line，位于四个矩形下方，且四条下引线的终点 y2
        #     与主干线的 y1 相等；主干线长度 |x2-x1| 换算 cm 落在 [9.90, 10.5]。
        # (3) 绿色主干线中心点向下引出高度为 0.3-0.4cm 的绿色箭头线：
        #     存在一条绿色 v:line，垂直、方向向下，末端带三角箭头，
        #     起点 (x1,y1) == (主干线中心, 主干线 y)，长度 (y2-y1) 换算 cm 落在 [0.3, 0.4]。
        # 位置属性均相对页面（mso-position-*-relative:page），Word/WPS 按同一页面坐标渲染。
        if len(boxes) != 4:
            return False
        lines = self.doc.green_lines()

        # (1): 四条 0.4-0.55cm 的下引竖线，起点在各矩形底部中心
        down_lines: list[Line] = []
        for b in boxes:
            hit = None
            for l in lines:
                if not l.is_vertical or l.y2 <= l.y1:
                    continue
                if abs(l.x1 - b.cx) > 2 or abs(l.y1 - b.bottom) > 2:
                    continue
                if not cm_between(l.vertical_len, 0.4, 0.55):
                    continue
                hit = l
                break
            if hit is None:
                return False
            down_lines.append(hit)

        # (2): 水平绿色主干线，长度 9.90-10.5cm，位于四矩形下方，
        #      且四条下引线的终点 y2 与主干线 y1 一致
        bus_y = down_lines[0].y2
        if any(abs(l.y2 - bus_y) > 2 for l in down_lines):
            return False
        bus = None
        for l in lines:
            if not l.is_horizontal:
                continue
            if abs(l.y1 - bus_y) > 2:
                continue
            if l.y1 <= max(b.bottom for b in boxes):
                continue
            if not cm_between(l.horizontal_len, 9.90, 10.5):
                continue
            bus = l
            break
        if bus is None:
            return False

        # (3): 主干线中心点向下的绿色箭头线，高 0.3-0.4cm
        bus_cx = (bus.x1 + bus.x2) / 2
        for l in lines:
            if not l.is_vertical or l.y2 <= l.y1:
                continue
            if not arrow_is_triangle(l.end_arrow):
                continue
            if abs(l.x1 - bus_cx) > 2 or abs(l.y1 - bus.y1) > 2:
                continue
            if cm_between(l.vertical_len, 0.3, 0.4):
                return True
        return False

    def _has_arrow_to_top(self, box: Shape, *, lines: Optional[list[Line]] = None, tip: str = "top_edge") -> bool:
        lines = lines or self.doc.green_lines()
        target_y = box.y
        target_x = box.cx if tip == "top_edge" else box.cx
        for l in lines:
            if not l.end_arrow:
                continue
            if l.is_vertical and l.y2 >= l.y1 and near(l.x2, target_x, 8) and abs(l.y2 - target_y) <= 8:
                return True
        return False

    def _has_down_arrow_from_bottom(self, box: Shape, lo_cm: float, hi_cm: float) -> bool:
        return any(l.is_vertical and l.end_arrow and near(l.x1, box.cx, 8) and abs(l.y1 - box.bottom) <= 8 and cm_between(l.vertical_len, lo_cm, hi_cm, tol=0.1) for l in self.doc.green_lines())

    def _has_down_arrow_from_bottom_center(self, box: Shape, lo_h_cm: float, hi_h_cm: float, lo_w_cm: float, hi_w_cm: float) -> bool:
        # 从矩形下框线中心向下引出的绿色带箭头竖线，高在 [lo_h_cm, hi_h_cm]，
        # 线粗 strokeweight 换算 cm 在 [lo_w_cm, hi_w_cm]。
        for l in self.doc.green_lines():
            if not l.is_vertical or l.y2 <= l.y1:
                continue
            if not l.end_arrow:
                continue
            if abs(l.x1 - box.cx) > 2 or abs(l.y1 - box.bottom) > 2:
                continue
            if not cm_between(l.vertical_len, lo_h_cm, hi_h_cm):
                continue
            if l.weight_pt is None or not cm_between(l.weight_pt, lo_w_cm, hi_w_cm):
                continue
            return True
        return False

    def _has_qualified_down_arrow_from_bottom(self, box: Shape, lo_cm: float, hi_cm: float, min_width_cm: float, max_width_cm: float) -> bool:
        for line in self.doc.green_arrow_lines():
            if (
                line.is_vertical
                and line.y2 > line.y1
                and line.weight_pt == 1.0
                and line.dash in {"", "solid"}
                and arrow_is_triangle(line.end_arrow)
                and abs(line.x1 - box.cx) <= 2
                and abs(line.y1 - box.bottom) <= 2
                and cm_between(line.vertical_len, lo_cm, hi_cm)
                and cm_between(line.weight_pt or 0, min_width_cm, max_width_cm)
            ):
                return True
        return False

    def _has_right_arrow_between(self, left: Shape, right: Shape, lo_cm: float, hi_cm: float) -> bool:
        return any(l.is_horizontal and l.end_arrow and l.x2 > l.x1 and abs(l.x1 - left.right) <= 15 and abs(l.x2 - right.x) <= 15 and abs(l.y2 - right.cy) <= 12 and cm_between(l.length, lo_cm, hi_cm, tol=0.15) for l in self.doc.green_lines())

    def _has_right_arrow_from_diamond_to_box(self, diamond: Shape, box: Shape, lo_w_cm: float, hi_w_cm: float, lo_h_cm: float, hi_h_cm: float) -> bool:
        # 从菱形右顶点 (diamond.right, diamond.cy) 出发、终点连到矩形左边框中心
        # (box.x, box.cy) 的水平绿色带箭头线；strokeweight 换算 cm 在 [lo_h_cm, hi_h_cm]，
        # 水平长度换算 cm 在 [lo_w_cm, hi_w_cm]。
        for l in self.doc.green_lines():
            if not l.is_horizontal or l.x2 <= l.x1:
                continue
            if not l.end_arrow:
                continue
            if abs(l.x1 - diamond.right) > 2 or abs(l.y1 - diamond.cy) > 2:
                continue
            if abs(l.x2 - box.x) > 2 or abs(l.y2 - box.cy) > 2:
                continue
            if not cm_between(l.horizontal_len, lo_w_cm, hi_w_cm):
                continue
            if l.weight_pt is None or not cm_between(l.weight_pt, lo_h_cm, hi_h_cm):
                continue
            return True
        return False

    def _has_qualified_right_arrow_between(self, left: Shape, right: Shape, lo_cm: float, hi_cm: float, min_height_cm: float, max_height_cm: float) -> bool:
        for line in self.doc.green_arrow_lines():
            if (
                line.is_horizontal
                and line.x2 > line.x1
                and line.weight_pt == 1.0
                and line.dash in {"", "solid"}
                and arrow_is_triangle(line.end_arrow)
                and abs(line.x1 - left.right) <= 2
                and abs(line.y1 - left.cy) <= 2
                and abs(line.x2 - right.x) <= 2
                and abs(line.y2 - right.cy) <= 2
                and cm_between(line.length, lo_cm, hi_cm)
                and cm_between(line.weight_pt or 0, min_height_cm, max_height_cm)
            ):
                return True
        return False

    def _check_return_loop(self, improve: Optional[Shape], diamond: Optional[Shape]) -> bool:
        # 细则五点，逐点校验（针对办公软件 Word/WPS 的实际渲染语义）：
        # (1) "完善资料与方案"矩形框位于"可行性确认"矩形框右侧：
        #     improve.x >= diamond.right（两者位置都相对页面 mso-position-*-relative:page，
        #     办公软件按同一页面坐标渲染，improve 左边缘不越过 diamond 右边缘即视觉上右侧）。
        # (2) 框内文字居中：段落 w:pPr 的 <w:jc w:val="center"/>。
        # (3) 通过绿色回路箭头返回上方"业务流程梳理"矩形框，且指向其矩形右侧中心：
        #     回路由三段构成，均为绿色 v:line —— 水平段 h1、竖直段 v、水平段 h2；
        #     h1 起点 (x1,y1)==(improve.right, improve.cy)，向右延伸；
        #     v 起点 y1 == h1 终点 y2，方向向上（y2<y1）；
        #     h2 起点 y1 == v 终点 y2，方向向左（x2<x1），末端带箭头，
        #     终点 (x2,y2) == (process.right, process.cy)（业务流程梳理矩形右侧中心）。
        # (4) 箭头水平方向长度约为 0.8-1.2cm：即 h1、h2 两段水平线的水平长度都落在 [0.8, 1.2]。
        # (5) 箭头竖直方向长度约为 5.2-5.6cm：v 段竖直长度 |y1-y2| 落在 [5.2, 5.6]。
        process = self.doc.find_shape_text("业务流程梳理")
        if not improve or not process or not diamond:
            return False
        if improve.x < diamond.right:
            return False
        if improve.style.align != "center":
            return False

        lines = self.doc.green_lines()

        # 找 h1：从 improve 右侧中心出发向右的水平绿色 v:line，长 0.8-1.2cm
        h1 = None
        for l in lines:
            if not l.is_horizontal or l.x2 <= l.x1:
                continue
            if abs(l.x1 - improve.right) > 2 or abs(l.y1 - improve.cy) > 2:
                continue
            if cm_between(l.horizontal_len, 0.8, 1.2):
                h1 = l
                break
        if h1 is None:
            return False

        # 找 v：从 h1 终点向上的竖直绿色 v:line，长 5.2-5.6cm
        v = None
        for l in lines:
            if not l.is_vertical or l.y2 >= l.y1:
                continue
            if abs(l.x1 - h1.x2) > 2 or abs(l.y1 - h1.y2) > 2:
                continue
            if cm_between(l.vertical_len, 5.2, 5.6):
                v = l
                break
        if v is None:
            return False

        # 找 h2：从 v 终点向左的水平绿色 v:line，长 0.8-1.2cm，末端带箭头，
        # 终点连接到 process 矩形右侧中心 (process.right, process.cy)
        for l in lines:
            if not l.is_horizontal or l.x2 >= l.x1:
                continue
            if not l.end_arrow:
                continue
            if abs(l.x1 - v.x2) > 2 or abs(l.y1 - v.y2) > 2:
                continue
            if abs(l.x2 - process.right) > 2 or abs(l.y2 - process.cy) > 2:
                continue
            if cm_between(l.horizontal_len, 0.8, 1.2):
                return True
        return False

    def _has_down_arrow_from_diamond_to_top(self, diamond: Shape, box: Shape, lo_cm: float, hi_cm: float) -> bool:
        # 从菱形下顶点 (diamond.cx, diamond.bottom) 出发、终点连到矩形上边框中心
        # (box.cx, box.y) 的绿色垂直向下带三角箭头 v:line；竖直长度在 [lo_cm, hi_cm]。
        for l in self.doc.green_lines():
            if not l.is_vertical or l.y2 <= l.y1:
                continue
            if not arrow_is_triangle(l.end_arrow):
                continue
            if abs(l.x1 - diamond.cx) > 2 or abs(l.y1 - diamond.bottom) > 2:
                continue
            if abs(l.x2 - box.cx) > 2 or abs(l.y2 - box.y) > 2:
                continue
            if cm_between(l.vertical_len, lo_cm, hi_cm):
                return True
        return False

    def _is_centered(self, box: Shape) -> bool:
        page_center = self.doc.page_w_pt / 2
        return abs(box.cx - page_center) <= 18

    def _has_arrow_between_y(self, src: Shape, dst: Shape, lo_cm: float, hi_cm: float, *, tol: float = 0.1) -> bool:
        return any(l.is_vertical and l.end_arrow and near(l.x1, src.cx, 12) and near(l.x2, dst.cx, 12) and abs(l.y1 - src.bottom) <= 12 and abs(l.y2 - dst.y) <= 12 and cm_between(l.vertical_len, lo_cm, hi_cm, tol=tol) for l in self.doc.green_lines())

    def _check_judge_to_sixth(self, judge: Optional[Shape], external: Optional[Shape], internal: Optional[Shape]) -> bool:
        # 细则三点，逐点校验（针对办公软件 Word/WPS 的实际渲染语义）：
        # (1) "运行环境综合研判"矩形框下方有一绿色箭头分别连接到第六行两个矩形框：
        #     由三段绿色 v:line 构成的分支结构 —— 从 judge 底部中心向下的短竖线，
        #     接一条水平主干线，主干线两端各向下引一条带箭头竖线，分别指向 external
        #     与 internal 的上边框中心 (cx, y)。VML v:line 的 from/to 端点即办公软件
        #     渲染时的实际端点，端点重合即视觉相接。
        # (2) 水平方向长度约为 7.5-8cm：主干线水平长度 |x2-x1| 换算 cm 落在 [7.5, 8.0]。
        # (3) 竖直方向长度约为 0.6-1cm：两条向下箭头的竖直长度 |y2-y1| 换算 cm
        #     都落在 [0.6, 1.0]。
        if not judge or not external or not internal:
            return False
        lines = self.doc.green_lines()

        # 从 judge 底部中心向下的短竖线，接主干线（不带箭头）
        down_stub = None
        for l in lines:
            if not l.is_vertical or l.y2 <= l.y1:
                continue
            if abs(l.x1 - judge.cx) > 2 or abs(l.y1 - judge.bottom) > 2:
                continue
            down_stub = l
            break
        if down_stub is None:
            return False

        # 水平主干线：y == down_stub.y2，两端 x 与 external.cx / internal.cx 精确对齐，
        # 水平长度 7.5-8cm
        bus = None
        left_cx = min(external.cx, internal.cx)
        right_cx = max(external.cx, internal.cx)
        for l in lines:
            if not l.is_horizontal:
                continue
            if abs(l.y1 - down_stub.y2) > 2:
                continue
            lx1, lx2 = sorted([l.x1, l.x2])
            if abs(lx1 - left_cx) > 2 or abs(lx2 - right_cx) > 2:
                continue
            if not cm_between(l.horizontal_len, 7.5, 8.0):
                continue
            bus = l
            break
        if bus is None:
            return False

        # 两条向下箭头，起点位于主干线两端，终点精确落在 external/internal 上边框中心，
        # 竖直长度 0.6-1cm
        for target in (external, internal):
            hit = False
            for l in lines:
                if not l.is_vertical or l.y2 <= l.y1:
                    continue
                if not arrow_is_triangle(l.end_arrow):
                    continue
                if abs(l.x1 - target.cx) > 2 or abs(l.y1 - bus.y1) > 2:
                    continue
                if abs(l.x2 - target.cx) > 2 or abs(l.y2 - target.y) > 2:
                    continue
                if not cm_between(l.vertical_len, 0.6, 1.0):
                    continue
                hit = True
                break
            if not hit:
                return False
        return True

    def _check_sixth_to_seventh(self, external: Optional[Shape], internal: Optional[Shape]) -> bool:
        # 细则逐点校验（针对办公软件 Word/WPS 的 VML 渲染语义）：
        # 左侧：
        #  (L1) 第六行左侧横向文本框("外部条件研判")下方有一绿色箭头 ——
        #       即存在一条 y 位于 external.bottom 下方的水平绿色主干线；
        #       办公软件按 VML from/to 直接绘制该水平线段。
        #  (L2) 分别连接到第七行左侧三个竖向文本框 ——
        #       主干线下方要有 3 条 end_arrow≠"" 的绿色 down-arrow(v:line vertical)，
        #       箭头下端(x2, y2)分别对准第七行左侧 3 个竖向文本框(w<h)顶部中心
        #       (top edge center，即 (b.cx, b.y))；Word/WPS 中 end_arrow=block/
        #       triangle/classic 均渲染为三角箭头。
        #  (L3) 水平方向长度为 2.5-3cm —— 主干线 horizontal_len/PT_PER_CM
        #       落在 [2.5, 3.0]。
        #  (L4) 竖直方向长度为 0.6-1cm —— 3 条 down-arrow 的 vertical_len 均
        #       落在 [0.6, 1.0] cm。
        # 右侧同理：4 条 down-arrow 分别指向第七行右侧 4 个竖向文本框，
        #       主干线水平 4-4.5cm，各下箭头竖直 0.6-1cm。
        if not external or not internal:
            return False
        lines = self.doc.green_lines()
        # 第七行竖向文本框：位于第六行下方、宽 < 高。按 x 排序取左 3 右 4。
        seventh = [s for s in self.doc.shapes if s.y > external.bottom + 6 and s.h > s.w]
        seventh.sort(key=lambda s: s.x)
        if len(seventh) < 7:
            return False
        left_boxes = seventh[:3]
        right_boxes = seventh[3:7]

        def check_side(anchor: Shape, boxes: list[Shape], h_lo: float, h_hi: float, v_lo: float, v_hi: float) -> bool:
            xs = [b.cx for b in boxes]
            x_lo, x_hi = min(xs), max(xs)
            # 定位该侧水平主干线：y 严格位于 anchor.bottom 之下，且水平覆盖该侧所有 box 中心 x。
            bus: Optional[Line] = None
            for l in lines:
                if not l.is_horizontal:
                    continue
                if l.y1 <= anchor.bottom:
                    continue
                l_x_lo, l_x_hi = sorted([l.x1, l.x2])
                if l_x_lo <= x_lo + 2 and l_x_hi >= x_hi - 2:
                    bus = l
                    break
            if bus is None:
                return False
            # (L3/R3) 水平方向长度 —— 严格区间，无 tol；±2pt 已由 is_horizontal 吸收浮点误差。
            if not cm_between(bus.horizontal_len, h_lo, h_hi):
                return False
            # (L2+L4 / R2+R4) 各分支下箭头端点对准 box 顶部中心，且竖直长度落在区间。
            for b in boxes:
                hit = False
                for l in lines:
                    if not (l.is_vertical and l.end_arrow):
                        continue
                    if l.y2 < l.y1:  # 需向下
                        continue
                    if abs(l.y1 - bus.y1) > 2:  # 顶端在主干线上
                        continue
                    if abs(l.x2 - b.cx) > 2 or abs(l.y2 - b.y) > 2:  # 尖端指向顶部中心
                        continue
                    if not cm_between(l.vertical_len, v_lo, v_hi):
                        continue
                    hit = True
                    break
                if not hit:
                    return False
            return True

        left_ok = check_side(external, left_boxes, 2.5, 3.0, 0.6, 1.0)
        right_ok = check_side(internal, right_boxes, 4.0, 4.5, 0.6, 1.0)
        return left_ok and right_ok

    def _check_seventh_row(self) -> bool:
        # 细则逐点校验（针对办公软件 Word/WPS 的 VML 渲染语义）：
        #  (1) 左侧三个竖向文本框 + 右侧四个竖向文本框 = 共 7 个；y 相同表示同一水平行。
        #  (2) 左三从左往右依次："政策环境扫描"、"需求场景识别"、"协同路径分析"（精确匹配）。
        #  (3) 右四从左往右依次："资源储备评估"、"流程执行能力"、"技术支撑能力"、"运营协调机制"。
        #  (4) 第 3 个与第 4 个文本框中间间隔 3.5-4cm ——
        #       gap = boxes[3].x - boxes[2].right（VML 页面坐标），换算 cm 落在 [3.5, 4.0]。
        #  (5) 其余文本框中间间隔（1-2、2-3、4-5、5-6、6-7 共 5 处）均为 0.8-1cm。
        #  (6) 文本框宽为 0.6-1cm —— VML width 换算 cm 落在 [0.6, 1.0]。
        #  (7) 文本框高为 2.5-3cm —— VML height 换算 cm 落在 [2.5, 3.0]。
        #  (8) 文字都在框内 —— 使用 v:rect/wsp 的 w:txbxContent（Word/WPS 将文本渲染在
        #       shape 内部，w:t 非空即视为 "在框内"），且文字长度不超过按 9pt(小五) 估算的
        #       竖向可容纳字符数 floor(h / 12pt)，超过则办公软件里会被裁切/溢出。
        #  (9) 竖向排布 —— 竖向文本框在办公软件里的渲染签名为 h > w（宽为一字宽，
        #       高足以叠放 N 个字），此时 Word/WPS 逐字向下堆叠。
        texts = ["政策环境扫描", "需求场景识别", "协同路径分析",
                 "资源储备评估", "流程执行能力", "技术支撑能力", "运营协调机制"]
        boxes = self.doc.find_all_texts(texts)
        if any(b is None for b in boxes):
            return False
        # 按 x 升序，验证从左往右顺序与文本一一对应（同时覆盖 (2)(3)）。
        bs = sorted([b for b in boxes if b], key=lambda s: s.x)
        if [norm_text(b.text) for b in bs] != [norm_text(t) for t in texts]:
            return False
        # (6)(7) 宽高严格区间（无 tol）。
        if not all(cm_between(b.w, 0.6, 1.0) and cm_between(b.h, 2.5, 3.0) for b in bs):
            return False
        # (4)(5) 间距：仅第 3-4 间为大间距，其余为小间距，均严格区间。
        gaps = [bs[i + 1].x - bs[i].right for i in range(6)]
        if not cm_between(gaps[2], 3.5, 4.0):
            return False
        if not all(cm_between(g, 0.8, 1.0) for i, g in enumerate(gaps) if i != 2):
            return False
        # (9) 竖向排布：h > w，办公软件里逐字向下堆叠。
        if not all(b.h > b.w for b in bs):
            return False
        # (8) 文字都在框内：文本非空，且字符数 <= 按 9pt(小五) 估算可容纳的竖向字数
        #     floor(box.h / 12pt)；小五行高约 12pt。
        for b in bs:
            chars = len(norm_text(b.text))
            if chars == 0:
                return False
            if chars > int(b.h // 12.0):
                return False
        return True

    def _check_eighth_row(self, efe: Optional[Shape], matrix: Optional[Shape], ife: Optional[Shape]) -> bool:
        # 细则逐点校验（针对办公软件 Word/WPS 的 VML 渲染语义）：
        #  (1) 第八行有三个横向文本框（三个 shape 均已由 find_shape_text 精确文本命中）。
        #  (2) 从左往右依次为 "EFE评阅清单"、"适配矩阵"、"IFE评阅清单" ——
        #       文本已由 find_shape_text 精确匹配；再校验 x 从小到大顺序。
        #  (3) 文本框宽为 2.5-3cm —— VML width 换算 cm 落在 [2.5, 3.0]，办公软件按此绘制。
        #  (4) 高为 0.6-1cm —— VML height 换算 cm 落在 [0.6, 1.0]。
        #  (5) 中间文本框左侧出现箭头指向左侧文本框 ——
        #       存在一条水平绿色 v:line，起点 (x1,y1) 在 matrix.left ±2pt、tip (x2,y2)
        #       在 EFE 一侧（x2 < x1，即向左），y 位于 matrix 垂直范围内 [matrix.y, matrix.bottom]，
        #       end_arrow ∈ {block, triangle, classic}（Word/WPS 均渲染三角箭头）。
        #  (6) 中间文本框右侧出现箭头指向右侧文本框 —— 对称：起点在 matrix.right ±2pt、
        #       tip 在 IFE 一侧（x2 > x1，即向右），其余同上。
        #  (7) 箭头长度为 0.2-0.5cm —— 上述两条箭头 horizontal_len/PT_PER_CM ∈ [0.2, 0.5]。
        if not efe or not matrix or not ife:
            return False
        if not (efe.x < matrix.x < ife.x):
            return False
        if not all(cm_between(b.w, 2.5, 3.0) for b in (efe, matrix, ife)):
            return False
        if not all(cm_between(b.h, 0.6, 1.0) for b in (efe, matrix, ife)):
            return False

        lines = self.doc.green_lines()

        def side_arrow(origin_x: float, direction: int) -> bool:
            # direction: -1 向左指向 EFE；+1 向右指向 IFE。
            for l in lines:
                if not l.is_horizontal or not l.end_arrow:
                    continue
                if abs(l.x1 - origin_x) > 2:
                    continue
                if direction < 0 and not (l.x2 < l.x1):
                    continue
                if direction > 0 and not (l.x2 > l.x1):
                    continue
                if not (matrix.y - 2 <= l.y1 <= matrix.bottom + 2):
                    continue
                if not cm_between(l.horizontal_len, 0.2, 0.5):
                    continue
                return True
            return False

        return side_arrow(matrix.x, -1) and side_arrow(matrix.right, +1)

    def _check_diagonal_arrows(self) -> bool:
        # 细则逐点校验（针对办公软件 Word/WPS 的 VML 渲染语义）：
        #  (1) "需求场景识别"文本框下方有一倾斜箭头 ——
        #       存在一条既非水平也非垂直的绿色 v:line，起点 (x1,y1) 位于
        #       "需求场景识别"下方（y1 ≥ demand.bottom - 2），且 end_arrow≠""；
        #       办公软件按 VML from/to 直接绘制这条斜线，末端渲染为三角箭头。
        #  (2) 指向"EFE评阅清单"文本框中心 —— tip (x2, y2) 命中 EFE 的中心
        #       (cx, cy) ±2pt。
        #  (3) "技术支撑能力"文本框下方有一倾斜箭头指向"IFE评阅清单"文本框中心
        #       —— 与 (1)(2) 对称，源为 tech.bottom 之下、tip 命中 IFE 中心。
        #  (4) 箭头长度约为 0.8-1.2cm —— l.length/PT_PER_CM ∈ [0.8, 1.2]。
        demand = self.doc.find_shape_text("需求场景识别")
        tech = self.doc.find_shape_text("技术支撑能力")
        efe = self.doc.find_shape_text("EFE评阅清单")
        ife = self.doc.find_shape_text("IFE评阅清单")
        if not all([demand, tech, efe, ife]):
            return False

        def has_diag(src: Shape, dst: Shape) -> bool:
            for l in self.doc.green_lines():
                if l.is_horizontal or l.is_vertical:
                    continue
                if not l.end_arrow:
                    continue
                # 起点在源文本框下方（y1 ≥ src.bottom - 2）
                if l.y1 < src.bottom - 2:
                    continue
                # 尖端命中目标文本框中心 ±2pt
                if abs(l.x2 - dst.cx) > 2 or abs(l.y2 - dst.cy) > 2:
                    continue
                # 长度 0.8-1.2cm（无 tol）
                if not cm_between(l.length, 0.8, 1.2):
                    continue
                return True
            return False

        return has_diag(demand, efe) and has_diag(tech, ife)

    def _check_stage_boxes(self) -> bool:
        # 细则逐点校验（针对办公软件 Word/WPS 的 VML 渲染语义）：
        #  (1) 页面左侧出现四个灰色背景填充的竖向文本框 ——
        #      w:txbxContent 分别为 "导入"/"核验"/"研判"/"形成"，均为 4 个 shape，
        #      每个 fill 为浅灰（RGB 三通道接近相等且亮度 180-255；办公软件里渲染为灰底）。
        #  (2) 高为 2.3-2.6cm —— VML height 换算 cm ∈ [2.3, 2.6]（无 tol）。
        #  (3) 宽为 0.6-1cm —— VML width 换算 cm ∈ [0.6, 1.0]（无 tol）。
        #  (4) 从上往下填充文本依次为"导入""核验""研判""形成"——
        #      按 y 升序，w:txbxContent 精确匹配。
        #  (5) 第 1-2 个间距 1.2-1.4cm —— gap = bs[1].y - bs[0].bottom（VML 页面坐标），
        #      换算 cm ∈ [1.2, 1.4]。
        #  (6) 第 2-3 个间距 2.4-2.7cm —— 换算 cm ∈ [2.4, 2.7]。
        #  (7) 第 3-4 个间距 5.8-6.2cm —— 换算 cm ∈ [5.8, 6.2]。
        #      （细则原文两处均写"第二个和第三个"，第二处按上下文与"依次三段"的表述
        #        实为"第三个与第四个"的笔误。）
        #  (8) 字体加粗 —— w:rPr/<w:b/> 为真。
        #  (9) 距离页面左边界 2.3-2.7cm —— shape.x（相对页面左边）换算 cm ∈ [2.3, 2.7]。
        texts = ["导入", "核验", "研判", "形成"]
        boxes = self.doc.find_all_texts(texts)
        if any(b is None for b in boxes):
            return False
        bs = sorted([b for b in boxes if b], key=lambda s: s.y)
        # (4)
        if [norm_text(b.text) for b in bs] != [norm_text(t) for t in texts]:
            return False
        # (1) 灰底填充；(2) 高；(3) 宽。
        if not all(is_greyish_fill(b.fill) for b in bs):
            return False
        if not all(cm_between(b.h, 2.3, 2.6) for b in bs):
            return False
        if not all(cm_between(b.w, 0.6, 1.0) for b in bs):
            return False
        # (5)(6)(7) 三段间距。
        gaps = [bs[i + 1].y - bs[i].bottom for i in range(3)]
        if not cm_between(gaps[0], 1.2, 1.4):
            return False
        if not cm_between(gaps[1], 2.4, 2.7):
            return False
        if not cm_between(gaps[2], 5.8, 6.2):
            return False
        # (8) 加粗。
        if not all(b.style.bold for b in bs):
            return False
        # (9) 距离页面左边界。
        if not all(cm_between(b.x, 2.3, 2.7) for b in bs):
            return False
        return True


# ---------- parsing helpers ----------

def qn(prefix: str, local: str) -> str:
    return f"{{{NS[prefix]}}}{local}"


def first(seq):
    return seq[0] if seq else None


def ancestor(el: etree._Element, tag: str) -> Optional[etree._Element]:
    p = el.getparent()
    while p is not None:
        if p.tag == tag:
            return p
        p = p.getparent()
    return None


def get_float(el: Optional[etree._Element], attr: str, default: float = 0.0) -> float:
    if el is None:
        return default
    try:
        return float(el.get(attr, default))
    except Exception:
        return default


def get_text_float(values: list[str], default: float = 0.0) -> float:
    try:
        return float(first(values) or default)
    except Exception:
        return default


def twip_to_pt(v: float) -> float:
    return v / TWIP_PER_PT


def emu_to_pt(v: float) -> float:
    return v / EMU_PER_PT


def pt_to_cm(v: float) -> float:
    return v / PT_PER_CM


def pt_from_cm(v: float) -> float:
    return v * PT_PER_CM


def pt2_to_cm2(v: float) -> float:
    return v / (PT_PER_CM * PT_PER_CM)


def parse_style(style: str) -> dict[str, str]:
    out: dict[str, str] = {}
    for part in style.split(";"):
        if ":" in part:
            k, v = part.split(":", 1)
            out[k.strip().lower()] = v.strip()
    return out


def parse_length_to_pt(value: str | None) -> Optional[float]:
    if value is None or value == "":
        return None
    s = str(value).strip().lower()
    m = re.match(r"(-?\d+(?:\.\d+)?)(pt|cm|in|mm|px)?", s)
    if not m:
        return None
    num = float(m.group(1))
    unit = m.group(2) or "pt"
    if unit == "pt":
        return num
    if unit == "cm":
        return num * PT_PER_CM
    if unit == "mm":
        return num / 10.0 * PT_PER_CM
    if unit == "in":
        return num * PT_PER_INCH
    if unit == "px":
        return num * 0.75
    return num


def parse_pair(value: str | None, default: tuple[float, float]) -> tuple[float, float]:
    if not value or "," not in value:
        return default
    a, b = value.split(",", 1)
    return parse_length_to_pt(a) or 0.0, parse_length_to_pt(b) or 0.0


def extract_text(el: etree._Element) -> str:
    parts = el.xpath(".//w:t/text()", namespaces=NS)
    return "".join(parts).strip()


def parse_text_style(el: etree._Element) -> TextStyle:
    rpr = first(el.xpath(".//w:rPr", namespaces=NS))
    font = ""
    size_pt = None
    bold = False
    if rpr is not None:
        fonts = first(rpr.xpath("./w:rFonts", namespaces=NS))
        if fonts is not None:
            font = fonts.get(qn("w", "eastAsia")) or fonts.get(qn("w", "ascii")) or fonts.get(qn("w", "hAnsi")) or ""
        sz = first(rpr.xpath("./w:sz", namespaces=NS))
        if sz is not None:
            try:
                size_pt = float(sz.get(qn("w", "val"))) / 2.0
            except Exception:
                size_pt = None
        bold = bool(rpr.xpath("./w:b", namespaces=NS))
    jc = first(el.xpath(".//w:pPr/w:jc", namespaces=NS))
    align = jc.get(qn("w", "val"), "") if jc is not None else ""
    return TextStyle(font=font, size_pt=size_pt, bold=bold, align=align)


def attr_false(value: str) -> bool:
    return str(value).strip().lower() in {"f", "false", "0"}


def norm_color(value: str) -> str:
    s = str(value or "").strip().upper()
    if not s:
        return ""
    if s.startswith("#"):
        s = s[1:]
    if len(s) == 3:
        s = "".join(ch * 2 for ch in s)
    return s


def rgb(value: str) -> Optional[tuple[int, int, int]]:
    c = norm_color(value)
    if not re.fullmatch(r"[0-9A-F]{6}", c):
        return None
    return int(c[0:2], 16), int(c[2:4], 16), int(c[4:6], 16)


def norm_text(s: str) -> str:
    # 去掉空白、换行和中英文引号差异；保留冒号等语义标点。
    s = (s or "").replace("：", ":")
    s = re.sub(r"[\s　'‘’\"“”]", "", s)
    return s


def is_white(c: str) -> bool:
    val = rgb(c)
    return bool(val and min(val) >= 245)


def is_blackish(c: str) -> bool:
    val = rgb(c)
    return bool(val and max(val) <= 90)


def is_black(c: str) -> bool:
    # 严格判定"黑色"：RGB 都很低（≈#000000），办公软件里视觉为黑。
    val = rgb(c)
    return bool(val and max(val) <= 40)


def is_grey(c: str) -> bool:
    val = rgb(c)
    return bool(val and max(val) - min(val) <= 18 and 80 <= sum(val) / 3 <= 220)


def is_greyish_fill(c: str) -> bool:
    val = rgb(c)
    return bool(val and max(val) - min(val) <= 20 and 180 <= sum(val) / 3 <= 255)


def is_green(c: str) -> bool:
    val = rgb(c)
    return bool(val and val[1] >= 100 and val[1] > val[0] * 1.2 and val[1] > val[2] * 1.2)


def is_heiti(font: str) -> bool:
    f = (font or "").lower()
    return any(k.lower() in f for k in ["SimHei", "黑体", "Microsoft YaHei", "微软雅黑", "Heiti"])


def is_heiti_strict(font: str) -> bool:
    # 严格判定"黑体"：仅接受 SimHei / 黑体 / Heiti；不把"微软雅黑"当作黑体。
    f = (font or "").strip()
    return any(k in f for k in ["SimHei", "黑体"]) or f.lower() in {"heiti", "heiti sc", "heiti tc"}


def between(value: Optional[float], lo: float, hi: float) -> bool:
    return value is not None and lo <= value <= hi


def cm_between(pt_value: Optional[float], lo_cm: float, hi_cm: float, *, tol: float = 0.0) -> bool:
    return pt_value is not None and (lo_cm - tol) <= pt_to_cm(pt_value) <= (hi_cm + tol)


def near(a: float, b: float, tol_pt: float) -> bool:
    return abs(a - b) <= tol_pt


def arrow_is_triangle(value: str) -> bool:
    return str(value).strip().lower() in {"block", "triangle", "classic"}


def _drawingml_arrow_kind(head_or_tail_type: str) -> str:
    """把 DrawingML <a:headEnd>/<a:tailEnd> 的 type（triangle/stealth/arrow/
    diamond/oval/none 等）归一化成本文件既有的 VML 箭头词汇（block/triangle/
    classic/空字符串），使 arrow_is_triangle() 等既有判定逻辑无需改动即可
    同时兼容 VML 与 DrawingML 两种来源的箭头。Word/WPS 对 triangle/stealth/
    arrow/oval/diamond 均渲染为可见的实心箭头（视觉上都是"末端小三角/箭头"
    的同义表达）；未设置或显式 "none" 表示无箭头。"""
    t = str(head_or_tail_type or "").strip().lower()
    if t in ("", "none"):
        return ""
    if t in ("triangle", "stealth", "arrow", "diamond", "oval"):
        return "triangle"
    return t


def detail_pos(s: Optional[Shape]) -> str:
    if not s:
        return "未找到"
    return f"{s.id or s.text}: x={pt_to_cm(s.x):.2f}cm, y={pt_to_cm(s.y):.2f}cm, w={pt_to_cm(s.w):.2f}cm, h={pt_to_cm(s.h):.2f}cm"


def detail_size(s: Optional[Shape]) -> str:
    if not s:
        return "未找到"
    return f"{s.id or s.text}: w={pt_to_cm(s.w):.2f}cm, h={pt_to_cm(s.h):.2f}cm, text={s.text}"


def detail_style(s: Optional[Shape]) -> str:
    if not s:
        return "未找到"
    return f"font={s.style.font}, size={s.style.size_pt}, bold={s.style.bold}, align={s.style.align}"


def detail_line(l: Optional[Line]) -> str:
    if not l:
        return "未找到"
    return f"{l.id}: color=#{l.color}, weight={l.weight_pt}pt, len={pt_to_cm(l.length):.2f}cm, y={pt_to_cm(l.y1):.2f}cm"


def summarize_styles(shapes: list[Shape]) -> str:
    if not shapes:
        return "未找到矩形框"
    bad = [f"{s.id or s.text}(font={s.style.font},size={s.style.size_pt},align={s.style.align})" for s in shapes if not (is_heiti(s.style.font) and between(s.style.size_pt, 8.5, 9.5) and s.style.align == "center")]
    return "全部符合" if not bad else "不符合示例：" + "; ".join(bad[:5])


# ---------- file handling / reporting ----------

SCRIPT_ID = "010"


def _find_target_doc(dir_path: Path) -> Optional[Path]:
    """在脚本目录中定位待评估的 .docx 文件（忽略 Office 临时文件）。"""
    candidates: list[Path] = []
    for p in sorted(dir_path.glob("*.docx")):
        if p.name.startswith("~$"):
            continue
        candidates.append(p)
    return candidates[0] if candidates else None


def _build_result(
    file_name: str,
    status: str,
    error: Optional[str],
    dim1_pass: bool,
    dim1_reason: str,
    dim2_items: list[dict],
    total_score: int,
    max_score: int,
) -> dict:
    return {
        "id": SCRIPT_ID,
        "file_name": file_name,
        "status": status,
        "error": error,
        "dim1_pass": dim1_pass,
        "dim1_reason": dim1_reason,
        "dim2_items": dim2_items,
        "total_score": total_score,
        "max_score": max_score,
    }


def evaluate(dir_path: str) -> dict:
    """统一入口：接收“脚本所在目录路径”，自动在该目录里定位并评估文档。"""
    try:
        base = Path(dir_path)
        if not base.exists() or not base.is_dir():
            return _build_result("", "error", f"目录不存在：{dir_path}", False, "", [], 0, 0)

        target = _find_target_doc(base)
        if target is None:
            return _build_result("", "error", f"目录中未找到 .docx 文件：{dir_path}", False, "", [], 0, 0)

        file_name = target.name

        doc = DocxRoadmap(target)
        doc.load()
        result = Scorer(doc).evaluate()

        dim1_reason = "" if result.dimension1_passed else "；".join(
            f"{name}：{detail}" for name, ok, detail in result.dimension1_items if not ok
        )

        dim2_items: list[dict] = []
        if result.dimension1_passed:
            for points, desc, detail in result.hits:
                dim2_items.append({
                    "rule": desc,
                    "max_delta": points,
                    "delta": points,
                    "hit": True,
                    "detail": "",
                })
            for points, desc, detail in result.misses:
                dim2_items.append({
                    "rule": desc,
                    "max_delta": points,
                    "delta": 0,
                    "hit": False,
                    "detail": "",
                })

        return _build_result(
            file_name,
            "ok",
            None,
            result.dimension1_passed,
            dim1_reason,
            dim2_items,
            result.score if result.dimension1_passed else 0,
            result.max_score if result.dimension1_passed else sum(i["max_delta"] for i in dim2_items),
        )
    except Exception as exc:
        return _build_result("", "error", f"{type(exc).__name__}: {exc}", False, "", [], 0, 0)


if __name__ == "__main__":
    target_dir = sys.argv[1] if len(sys.argv) > 1 else str(Path(__file__).resolve().parent)
    print(json.dumps(evaluate(target_dir), ensure_ascii=False, indent=2))
