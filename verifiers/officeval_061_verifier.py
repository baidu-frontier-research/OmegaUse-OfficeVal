#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
自动评估《公司简介_动画合并版.pptx》。

对外接口：
    evaluate(dir_path: str) -> dict
    接收脚本所在目录路径，脚本自身在该目录内定位并打开被评估的 .pptx 文件，
    返回结构化字典（含维度一通过与否、维度二逐项得分、总分等）。

说明：
- 仅依赖 Python 标准库；如环境存在 Pillow，会额外检查图片清晰度/比例。
- PPTX 中“图片内容语义”（例如是否为书本发光、课堂举手）无法在不接入视觉模型/OCR 的情况下可靠判断，
  本脚本采用可自动化的代理指标：图片对象数量、位置、尺寸、圆角裁剪、可编辑性、比例、动画对象与路径等。
- 维度 1 不通过时不再评估维度 2，total_score 记为 0。
"""

from __future__ import annotations

import hashlib
import re
import sys
import zipfile
import xml.etree.ElementTree as ET
from dataclasses import dataclass, field
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Sequence, Tuple

SCRIPT_ID = "061"

try:
    from PIL import Image
except Exception:  # pragma: no cover - Pillow 是可选依赖
    Image = None

EMU_PER_INCH = 914400
CM_PER_INCH = 2.54

NS = {
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
}


def qn(prefix: str, tag: str) -> str:
    return f"{{{NS[prefix]}}}{tag}"


def emu_to_in(value: Optional[str]) -> float:
    if value is None:
        return 0.0
    return int(value) / EMU_PER_INCH


def cm_to_in(cm: float) -> float:
    return cm / CM_PER_INCH


def almost(value: float, low: float, high: float, tol: float = 0.0) -> bool:
    return low - tol <= value <= high + tol


@dataclass
class Bounds:
    x: float
    y: float
    w: float
    h: float

    @property
    def r(self) -> float:
        return self.x + self.w

    @property
    def b(self) -> float:
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

    def intersect_area(self, other: "Bounds") -> float:
        x1, y1 = max(self.x, other.x), max(self.y, other.y)
        x2, y2 = min(self.r, other.r), min(self.b, other.b)
        return max(0.0, x2 - x1) * max(0.0, y2 - y1)

    def inside_percent(self, slide_w: float, slide_h: float, left: float, right: float, top: float, bottom: float) -> bool:
        return self.x >= slide_w * left and self.r <= slide_w * right and self.y >= slide_h * top and self.b <= slide_h * bottom

    def union(self, other: "Bounds") -> "Bounds":
        x1, y1 = min(self.x, other.x), min(self.y, other.y)
        x2, y2 = max(self.r, other.r), max(self.b, other.b)
        return Bounds(x1, y1, x2 - x1, y2 - y1)

    def fmt(self) -> str:
        return f"x={self.x:.2f}, y={self.y:.2f}, w={self.w:.2f}, h={self.h:.2f} 英寸"


@dataclass
class TextRun:
    text: str
    size_pt: Optional[float]
    bold: Optional[bool]
    color: Optional[str]


@dataclass
class Shape:
    slide_index: int
    kind: str  # sp / pic / grp
    shape_id: str
    name: str
    bounds: Bounds
    z_index: int
    text: str = ""
    runs: List[TextRun] = field(default_factory=list)
    geom: str = ""
    media_target: str = ""
    media_size_px: Optional[Tuple[int, int]] = None
    has_pic_fill: bool = False  # sp 形状是否带图片填充（a:blipFill），WPS/PPT 中同样显示为“图片”。

    @property
    def is_picture_like(self) -> bool:
        # pic 对象，或带图片填充的形状，用户视角都是“图片”。
        return self.kind == "pic" or (self.kind == "sp" and self.has_pic_fill)

    @property
    def display_aspect(self) -> float:
        return self.bounds.w / self.bounds.h if self.bounds.h else 0.0

    @property
    def media_aspect(self) -> Optional[float]:
        if not self.media_size_px or self.media_size_px[1] == 0:
            return None
        return self.media_size_px[0] / self.media_size_px[1]

    @property
    def ppi(self) -> Optional[Tuple[float, float]]:
        if not self.media_size_px or self.bounds.w <= 0 or self.bounds.h <= 0:
            return None
        return (self.media_size_px[0] / self.bounds.w, self.media_size_px[1] / self.bounds.h)


@dataclass
class AnimationInfo:
    target_ids: List[str]
    motion_target_ids: List[str]
    horizontal_motion_ids: List[str]
    durations_ms: List[int]
    delays_ms: List[int]
    click_trigger_count: int
    paths: List[str]
    raw_anim_count: int
    vertical_motion_ids: List[str] = field(default_factory=list)
    rotation_anim_count: int = 0
    fly_in_count: int = 0
    # spid -> 该卡片所有 animMotion 的 x 位移量列表（单位：幻灯片宽度的比例，PowerPoint motion path 坐标为 0-1）。
    motion_x_offsets: Dict[str, List[float]] = field(default_factory=dict)
    # spid -> 该卡片所有 animMotion 的 y 位移量列表（单位：幻灯片高度的比例）。
    motion_y_offsets: Dict[str, List[float]] = field(default_factory=dict)
    # spid -> 该卡片所有 animMotion 的 (x_off, y_off) 有向序列（保留正负号，用于判断能否"轮到中心"）。
    motion_signed_offsets: Dict[str, List[Tuple[float, float]]] = field(default_factory=dict)
    # spid -> 该卡片 animScale 上的最大缩放比例列表（x, y），单位：100000 = 100%。
    scale_by_shape: Dict[str, List[Tuple[float, float]]] = field(default_factory=dict)
    # spid -> 该卡片透明度动画中出现过的所有 alpha 值（0-100 归一化）。
    alpha_values_by_shape: Dict[str, List[float]] = field(default_factory=dict)
    # spid -> 该卡片所属 animMotion/animScale/anim 节点的 delay 集合（ms），用于判定"依次"播放。
    delays_by_shape: Dict[str, List[int]] = field(default_factory=dict)
    # 时间线中出现 spTgt 的顺序（一维列表），用于判断卡片是否按顺序被触发。
    timeline_order: List[str] = field(default_factory=list)


@dataclass
class Slide:
    index: int
    path: str
    xml: ET.Element
    rels: Dict[str, str]
    shapes: List[Shape]
    animation: AnimationInfo

    @property
    def text(self) -> str:
        return "".join(s.text for s in self.shapes if s.text)


@dataclass
class PointResult:
    score: int
    name: str
    hit: bool
    evidence: str


class PptxAnalyzer:
    def __init__(self, path: Path):
        self.path = path
        self.slide_w = 0.0
        self.slide_h = 0.0
        self.slide_paths: List[str] = []
        self.slides: List[Slide] = []
        self.media_sizes: Dict[str, Tuple[int, int]] = {}
        # 主题色映射：scheme 名（dk1/lt1/dk2/.../accent1..6/tx1/bg1/tx2/bg2/hlink/folHlink）→ 6 位十六进制。
        self.theme_colors: Dict[str, str] = {}
        self._zip: Optional[zipfile.ZipFile] = None

    def __enter__(self) -> "PptxAnalyzer":
        self._zip = zipfile.ZipFile(self.path)
        self._load_media_sizes()
        self._load_theme_colors()
        self._load_presentation()
        self._load_slides()
        return self

    def __exit__(self, exc_type, exc, tb):
        if self._zip:
            self._zip.close()

    @property
    def z(self) -> zipfile.ZipFile:
        assert self._zip is not None
        return self._zip

    def _read_xml(self, name: str) -> ET.Element:
        return ET.fromstring(self.z.read(name))

    def _load_media_sizes(self) -> None:
        if Image is None:
            return
        for name in self.z.namelist():
            if not name.startswith("ppt/media/"):
                continue
            suffix = Path(name).suffix.lower()
            if suffix not in {".png", ".jpg", ".jpeg", ".gif", ".bmp", ".tif", ".tiff"}:
                continue
            try:
                with self.z.open(name) as f:
                    img = Image.open(f)
                    self.media_sizes[name] = img.size
            except Exception:
                pass

    def _load_theme_colors(self) -> None:
        # 从 ppt/theme/theme*.xml 解析 a:clrScheme，把主题色（dk1/lt1/dk2/lt2/accent1..6/hlink/folHlink）
        # 映射到 6 位十六进制。用于把 schemeClr 保守还原到 RGB，避免误判为深色。
        theme_names = [n for n in self.z.namelist() if n.startswith("ppt/theme/") and n.endswith(".xml")]
        if not theme_names:
            return
        # 只解析第一个 theme（本次任务默认使用单一主题）；如需多主题，可按 slide→master→theme 依次解析。
        theme_names.sort()
        try:
            root = self._read_xml(theme_names[0])
        except Exception:
            return
        clr_scheme = root.find(".//a:clrScheme", NS)
        if clr_scheme is None:
            return
        # 常见 scheme 别名：tx1/bg1/tx2/bg2 分别对应 dk1/lt1/dk2/lt2。
        aliases = {"dk1": "tx1", "lt1": "bg1", "dk2": "tx2", "lt2": "bg2"}
        for child in clr_scheme:
            # tag 形如 {..drawingml..}dk1、accent1 等。
            local = child.tag.split("}", 1)[-1]
            srgb = child.find("a:srgbClr", NS)
            sys_clr = child.find("a:sysClr", NS)
            rgb: Optional[str] = None
            if srgb is not None and srgb.attrib.get("val"):
                rgb = srgb.attrib["val"].lower()
            elif sys_clr is not None:
                # sysClr 通常带 lastClr 作为回退 RGB。
                rgb = (sys_clr.attrib.get("lastClr") or "").lower() or None
            if rgb and re.fullmatch(r"[0-9a-f]{6}", rgb):
                self.theme_colors[local] = rgb
                if local in aliases:
                    self.theme_colors[aliases[local]] = rgb
                # 反向别名：tx1→dk1 等，方便双向查询。
                for k, v in aliases.items():
                    if v == local:
                        self.theme_colors[k] = rgb

    def _load_rels(self, rels_path: str) -> Dict[str, str]:
        if rels_path not in self.z.namelist():
            return {}
        root = self._read_xml(rels_path)
        rels = {}
        base_dir = str(Path(rels_path).parent.parent).replace("\\", "/")
        for rel in root:
            rid = rel.attrib.get("Id", "")
            target = rel.attrib.get("Target", "")
            if not rid or not target:
                continue
            if target.startswith("../"):
                full = "ppt/" + target[3:]
            elif target.startswith("/"):
                full = target.lstrip("/")
            elif base_dir and base_dir != ".":
                full = f"{base_dir}/{target}"
            else:
                full = target
            rels[rid] = full.replace("\\", "/")
        return rels

    def _load_presentation(self) -> None:
        root = self._read_xml("ppt/presentation.xml")
        sld_sz = root.find("p:sldSz", NS)
        if sld_sz is not None:
            self.slide_w = emu_to_in(sld_sz.attrib.get("cx"))
            self.slide_h = emu_to_in(sld_sz.attrib.get("cy"))

        prs_rels = self._load_rels("ppt/_rels/presentation.xml.rels")
        sld_id_lst = root.find("p:sldIdLst", NS)
        if sld_id_lst is None:
            self.slide_paths = []
            return
        paths = []
        for sld_id in sld_id_lst:
            rid = sld_id.attrib.get(qn("r", "id"), "")
            target = prs_rels.get(rid, "")
            if target.startswith("ppt/"):
                paths.append(target)
            elif target:
                paths.append("ppt/" + target.lstrip("/"))
        self.slide_paths = paths

    def _load_slides(self) -> None:
        self.slides = []
        for idx, slide_path in enumerate(self.slide_paths, start=1):
            xml = self._read_xml(slide_path)
            rels_path = f"ppt/slides/_rels/{Path(slide_path).name}.rels"
            rels = self._load_rels(rels_path)
            shapes = self._extract_shapes(idx, xml, rels)
            anim = self._extract_animation(xml)
            self.slides.append(Slide(idx, slide_path, xml, rels, shapes, anim))

    def _extract_bounds(self, elem: ET.Element) -> Optional[Bounds]:
        # 普通形状和图片一般在 a:xfrm；组合形状在 p:grpSpPr/a:xfrm。
        xfrm = elem.find(".//a:xfrm", NS)
        if elem.tag == qn("p", "grpSp"):
            group_xfrm = elem.find("p:grpSpPr/a:xfrm", NS)
            if group_xfrm is not None:
                xfrm = group_xfrm
        if xfrm is None:
            return None
        off = xfrm.find("a:off", NS)
        ext = xfrm.find("a:ext", NS)
        if off is None or ext is None:
            return None
        return Bounds(
            emu_to_in(off.attrib.get("x")),
            emu_to_in(off.attrib.get("y")),
            emu_to_in(ext.attrib.get("cx")),
            emu_to_in(ext.attrib.get("cy")),
        )

    def _shape_name_id(self, elem: ET.Element) -> Tuple[str, str]:
        c_nv_pr = elem.find(".//p:cNvPr", NS)
        if c_nv_pr is None:
            return "", ""
        return c_nv_pr.attrib.get("name", ""), c_nv_pr.attrib.get("id", "")

    def _extract_text_runs(self, elem: ET.Element) -> Tuple[str, List[TextRun]]:
        runs: List[TextRun] = []
        all_text = []
        for r in elem.findall(".//a:r", NS):
            txt = "".join(t.text or "" for t in r.findall("a:t", NS))
            if not txt:
                continue
            all_text.append(txt)
            r_pr = r.find("a:rPr", NS)
            size_pt = None
            bold = None
            color = None
            if r_pr is not None:
                if r_pr.attrib.get("sz"):
                    try:
                        size_pt = int(r_pr.attrib["sz"]) / 100
                    except ValueError:
                        pass
                if r_pr.attrib.get("b") is not None:
                    bold = r_pr.attrib.get("b") in {"1", "true", "True"}
                srgb = r_pr.find(".//a:srgbClr", NS)
                scheme = r_pr.find(".//a:schemeClr", NS)
                if srgb is not None:
                    color = (srgb.attrib.get("val") or "").lower() or None
                elif scheme is not None:
                    name = (scheme.attrib.get("val") or "").lower()
                    if not name:
                        color = None
                    elif name in self.theme_colors:
                        # 主题色成功解析到实际 RGB。
                        color = self.theme_colors[name]
                    elif name in {"tx1", "dk1", "dk2", "black"}:
                        # OOXML 规范中 tx1/dk1/dk2 语义上就是深色，允许作为“深色”回退。
                        color = name
                    else:
                        # 未能解析的主题色（如 accent1/lt1/hlink 等）：保留原样，判定时按“未确认”处理。
                        color = f"scheme:{name}"
            runs.append(TextRun(txt, size_pt, bold, color))
        # 有些文本尺寸写在默认属性中，给没有 size 的 run 补一次。
        default_size = None
        default_bold = None
        for r_pr in elem.findall(".//a:defRPr", NS):
            if r_pr.attrib.get("sz") and default_size is None:
                try:
                    default_size = int(r_pr.attrib["sz"]) / 100
                except ValueError:
                    pass
            if r_pr.attrib.get("b") is not None and default_bold is None:
                default_bold = r_pr.attrib.get("b") in {"1", "true", "True"}
        if default_size is not None or default_bold is not None:
            for run in runs:
                if run.size_pt is None:
                    run.size_pt = default_size
                if run.bold is None:
                    run.bold = default_bold
        return "".join(all_text), runs

    def _extract_geom(self, elem: ET.Element) -> str:
        prst = elem.find(".//a:prstGeom", NS)
        if prst is not None:
            return prst.attrib.get("prst", "")
        if elem.find(".//a:custGeom", NS) is not None:
            return "custGeom"
        return ""

    def _extract_media_target(self, elem: ET.Element, rels: Dict[str, str]) -> str:
        blip = elem.find(".//a:blip", NS)
        if blip is None:
            return ""
        rid = blip.attrib.get(qn("r", "embed"), "") or blip.attrib.get(qn("r", "link"), "")
        return rels.get(rid, "")

    def _extract_shapes(self, slide_index: int, root: ET.Element, rels: Dict[str, str]) -> List[Shape]:
        shapes: List[Shape] = []
        sp_tree = root.find("p:cSld/p:spTree", NS)
        iterable = list(sp_tree) if sp_tree is not None else list(root)
        z_counter = 0

        def add_shape(elem: ET.Element, kind: str) -> None:
            nonlocal z_counter
            bounds = self._extract_bounds(elem)
            if bounds is None:
                return
            name, shape_id = self._shape_name_id(elem)
            text, runs = self._extract_text_runs(elem)
            geom = self._extract_geom(elem)
            # sp 形状可能通过 a:blipFill 承载图片（WPS/PPT 中同样显示为图片）。
            has_pic_fill = kind == "sp" and elem.find(".//p:spPr/a:blipFill", NS) is not None
            if has_pic_fill and elem.find(".//p:spPr//a:blip", NS) is None:
                has_pic_fill = False
            media_target = self._extract_media_target(elem, rels) if (kind == "pic" or has_pic_fill) else ""
            media_size = self.media_sizes.get(media_target)
            z_counter += 1
            shapes.append(
                Shape(
                    slide_index=slide_index,
                    kind=kind,
                    shape_id=shape_id,
                    name=name,
                    bounds=bounds,
                    z_index=z_counter,
                    text=text,
                    runs=runs,
                    geom=geom,
                    media_target=media_target,
                    media_size_px=media_size,
                    has_pic_fill=has_pic_fill,
                )
            )

        # 先按 spTree 的直接顺序记录，再补充嵌套对象，保证 z_index 大体可用于遮挡判断。
        for elem in iterable:
            tag = elem.tag
            if tag == qn("p", "sp"):
                add_shape(elem, "sp")
            elif tag == qn("p", "pic"):
                add_shape(elem, "pic")
            elif tag == qn("p", "grpSp"):
                add_shape(elem, "grp")
                for sub in elem.findall(".//p:sp", NS):
                    add_shape(sub, "sp")
                for sub in elem.findall(".//p:pic", NS):
                    add_shape(sub, "pic")

        # 兜底：如果直接遍历没拿到嵌套 pic/sp，用全局 XPath 补齐。
        known = {(s.kind, s.shape_id) for s in shapes}
        for elem, kind in [(e, "sp") for e in root.findall(".//p:sp", NS)] + [(e, "pic") for e in root.findall(".//p:pic", NS)]:
            name, sid = self._shape_name_id(elem)
            if (kind, sid) not in known:
                add_shape(elem, kind)
                known.add((kind, sid))
        return shapes

    def _extract_animation(self, root: ET.Element) -> AnimationInfo:
        target_ids: List[str] = []
        motion_ids: List[str] = []
        horizontal_ids: List[str] = []
        vertical_ids: List[str] = []
        durations: List[int] = []
        delays: List[int] = []
        paths: List[str] = []
        raw_anim_count = 0
        rotation_anim_count = 0
        fly_in_count = 0

        timing = root.find("p:timing", NS)
        if timing is None:
            return AnimationInfo([], [], [], [], [], 0, [], 0)

        for ctn in timing.findall(".//p:cTn", NS):
            for attr, out in [("dur", durations), ("delay", delays), ("stDelay", delays)]:
                value = ctn.attrib.get(attr)
                if value and value.isdigit():
                    out.append(int(value))

        for sp_tgt in timing.findall(".//p:spTgt", NS):
            sid = sp_tgt.attrib.get("spid")
            if sid:
                target_ids.append(sid)

        # 时间线顺序：以 spTgt 在 XML 文档中的出现顺序作为触发顺序的代理。
        # PPTX 的 p:seq/p:childTnLst 是线性组织的，spTgt 的文档序基本等价于播放顺序。
        timeline_order: List[str] = []
        for sp_tgt in timing.iter(qn("p", "spTgt")):
            sid = sp_tgt.attrib.get("spid")
            if sid:
                timeline_order.append(sid)

        for node in list(timing.findall(".//p:anim", NS)) + list(timing.findall(".//p:animEffect", NS)) + list(timing.findall(".//p:animScale", NS)) + list(timing.findall(".//p:set", NS)):
            raw_anim_count += 1

        # 旋转动画：animRot 节点，或 anim 中对 r/rotation 属性做动画。
        for rot in timing.findall(".//p:animRot", NS):
            rotation_anim_count += 1
        for anim in timing.findall(".//p:anim", NS):
            attr_names = [a.text or "" for a in anim.findall(".//p:attrName", NS)]
            if any("r" == n or "rotation" in n.lower() for n in attr_names):
                rotation_anim_count += 1

        # 飞入/随机进入效果：animEffect 的 transition/filter 含 fly/random/in。
        for eff in timing.findall(".//p:animEffect", NS):
            transition = (eff.attrib.get("transition", "") or "").lower()
            filt = (eff.attrib.get("filter", "") or "").lower()
            blob = transition + " " + filt
            if "fly" in blob or "random" in blob or transition == "in":
                fly_in_count += 1

        motion_x_offsets: Dict[str, List[float]] = {}
        motion_y_offsets: Dict[str, List[float]] = {}
        motion_signed_offsets: Dict[str, List[Tuple[float, float]]] = {}
        scale_by_shape: Dict[str, List[Tuple[float, float]]] = {}
        alpha_values_by_shape: Dict[str, List[float]] = {}
        delays_by_shape: Dict[str, List[int]] = {}

        def _delay_of(node: ET.Element) -> Optional[int]:
            # 就近查询节点自身及父级 cTn 的 delay/stDelay，用于把 delay 关联到具体 spid。
            for target in (node, *list(node.iter())):
                cur = target.find("p:cTn", NS) if target is node else None
                if cur is None and target.tag.endswith("}cTn"):
                    cur = target
                if cur is None:
                    continue
                for attr in ("delay", "stDelay"):
                    v = cur.attrib.get(attr)
                    if v and v.isdigit():
                        return int(v)
            return None

        for motion in timing.findall(".//p:animMotion", NS):
            raw_anim_count += 1
            sid = ""
            sp_tgt = motion.find(".//p:spTgt", NS)
            if sp_tgt is not None:
                sid = sp_tgt.attrib.get("spid", "")
            if sid:
                motion_ids.append(sid)
            path = " ".join(str(motion.attrib.get(k, "")) for k in ("path", "by", "from", "to", "rCtr"))
            paths.append(path.strip())
            if self._looks_horizontal_motion(path):
                if sid:
                    horizontal_ids.append(sid)
            elif self._looks_vertical_motion(path):
                if sid:
                    vertical_ids.append(sid)
            if sid:
                x_off, y_off = self._motion_offsets(motion, path)
                if x_off is not None:
                    motion_x_offsets.setdefault(sid, []).append(x_off)
                if y_off is not None:
                    motion_y_offsets.setdefault(sid, []).append(y_off)
                sx, sy = self._motion_signed_offsets(motion, path)
                if sx is not None or sy is not None:
                    motion_signed_offsets.setdefault(sid, []).append((sx or 0.0, sy or 0.0))
                d = _delay_of(motion)
                if d is not None:
                    delays_by_shape.setdefault(sid, []).append(d)

        # animScale：解析 by/from/to 上的 x/y 缩放比例（100000 = 100%）。
        for scale in timing.findall(".//p:animScale", NS):
            sp_tgt = scale.find(".//p:spTgt", NS)
            sid = sp_tgt.attrib.get("spid", "") if sp_tgt is not None else ""
            if not sid:
                continue
            for attr in ("by", "from", "to"):
                node = scale.find(f"p:{attr}", NS)
                if node is None:
                    continue
                x_val = node.attrib.get("x")
                y_val = node.attrib.get("y")
                try:
                    xv = float(x_val) / 100000.0 if x_val is not None else 1.0
                    yv = float(y_val) / 100000.0 if y_val is not None else 1.0
                except ValueError:
                    continue
                scale_by_shape.setdefault(sid, []).append((xv, yv))
            d = _delay_of(scale)
            if d is not None:
                delays_by_shape.setdefault(sid, []).append(d)

        # 透明度动画：anim 对 style.opacity / ppt_alpha 做动画，或 set/anim 的 tavLst 里出现 alpha。
        for anim in timing.findall(".//p:anim", NS):
            attr_names = [a.text or "" for a in anim.findall(".//p:attrName", NS)]
            is_alpha = any(("opacity" in n.lower() or "alpha" in n.lower()) for n in attr_names)
            if not is_alpha:
                continue
            sp_tgt = anim.find(".//p:spTgt", NS)
            sid = sp_tgt.attrib.get("spid", "") if sp_tgt is not None else ""
            if not sid:
                continue
            for tav in anim.findall(".//p:tav", NS):
                val = tav.find(".//p:strVal", NS)
                raw = val.attrib.get("val") if val is not None else None
                if raw is None:
                    fval = tav.find(".//p:fltVal", NS)
                    raw = fval.attrib.get("val") if fval is not None else None
                if raw is None:
                    continue
                m = re.search(r"[-+]?\d*\.?\d+", raw)
                if not m:
                    continue
                try:
                    v = float(m.group(0))
                except ValueError:
                    continue
                # 归一化到 0-100（部分文档以 0-1 存 alpha，部分以 0-100000）。
                if v <= 1.5:
                    v *= 100
                elif v > 1000:
                    v /= 1000.0
                alpha_values_by_shape.setdefault(sid, []).append(v)

        click_count = 0
        for cond in timing.findall(".//p:cond", NS):
            evt = cond.attrib.get("evt", "")
            if "click" in evt.lower() or evt in {"onBegin", "onClick"}:
                click_count += 1

        return AnimationInfo(
            target_ids=dedupe(target_ids),
            motion_target_ids=dedupe(motion_ids),
            horizontal_motion_ids=dedupe(horizontal_ids),
            durations_ms=durations,
            delays_ms=delays,
            click_trigger_count=click_count,
            paths=paths,
            raw_anim_count=raw_anim_count,
            vertical_motion_ids=dedupe(vertical_ids),
            rotation_anim_count=rotation_anim_count,
            fly_in_count=fly_in_count,
            motion_x_offsets=motion_x_offsets,
            motion_y_offsets=motion_y_offsets,
            motion_signed_offsets=motion_signed_offsets,
            scale_by_shape=scale_by_shape,
            alpha_values_by_shape=alpha_values_by_shape,
            delays_by_shape=delays_by_shape,
            timeline_order=timeline_order,
        )

    @staticmethod
    def _motion_signed_offsets(motion: ET.Element, path: str) -> Tuple[Optional[float], Optional[float]]:
        # 保留正负号版本：用于判断能否把卡片"向中心"移动。
        by = motion.attrib.get("by") or motion.attrib.get("to") or motion.attrib.get("from")
        if by:
            m = re.findall(r"[-+]?\d*\.?\d+", by)
            x_off = float(m[0]) if len(m) >= 1 else None
            y_off = float(m[1]) if len(m) >= 2 else None
            if x_off is not None or y_off is not None:
                return x_off, y_off
        nums = [float(x) for x in re.findall(r"[-+]?\d*\.?\d+", path)]
        if len(nums) >= 2:
            xs = nums[0::2]
            ys = nums[1::2]
            # 有向：从起点到终点的净位移（末点 - 首点）。
            x_off = xs[-1] - xs[0] if len(xs) >= 2 else None
            y_off = ys[-1] - ys[0] if len(ys) >= 2 else None
            return x_off, y_off
        return None, None

    @staticmethod
    def _motion_offsets(motion: ET.Element, path: str) -> Tuple[Optional[float], Optional[float]]:
        # 取该 animMotion 在 x、y 方向的最大偏移量（x 为幻灯片宽度比例、y 为高度比例，motion path 坐标系 0-1）。
        by = motion.attrib.get("by") or motion.attrib.get("to") or motion.attrib.get("from")
        if by:
            m = re.findall(r"[-+]?\d*\.?\d+", by)
            x_off = abs(float(m[0])) if len(m) >= 1 else None
            y_off = abs(float(m[1])) if len(m) >= 2 else None
            if x_off is not None or y_off is not None:
                return x_off, y_off
        nums = [float(x) for x in re.findall(r"[-+]?\d*\.?\d+", path)]
        if len(nums) >= 2:
            xs = nums[0::2]
            ys = nums[1::2]
            x_off = max(abs(x) for x in xs) if xs else None
            y_off = max(abs(y) for y in ys) if ys else None
            return x_off, y_off
        return None, None

    @staticmethod
    def _looks_horizontal_motion(path: str) -> bool:
        if not path:
            return False
        nums = [float(x) for x in re.findall(r"[-+]?\d*\.?\d+", path)]
        if len(nums) < 2:
            # PowerPoint motion path 文本中 H/L 常见为横向线段，作为弱判断。
            return bool(re.search(r"\b[HhLl]\b", path))
        # 常见 path 数字按 x,y 成对出现；只要 x 变化明显、y 变化较小，就视为水平/近似水平。
        pts = list(zip(nums[0::2], nums[1::2]))
        if len(pts) < 2:
            return False
        xs = [p[0] for p in pts]
        ys = [p[1] for p in pts]
        return (max(xs) - min(xs)) > 0.05 and (max(ys) - min(ys)) <= max(0.03, (max(xs) - min(xs)) * 0.25)

    @staticmethod
    def _looks_vertical_motion(path: str) -> bool:
        # 上下跳动：y 变化明显且大于 x 变化，视为纵向运动（细则禁止）。
        if not path:
            return False
        nums = [float(x) for x in re.findall(r"[-+]?\d*\.?\d+", path)]
        if len(nums) < 2:
            return bool(re.search(r"\b[Vv]\b", path))
        pts = list(zip(nums[0::2], nums[1::2]))
        if len(pts) < 2:
            return False
        xs = [p[0] for p in pts]
        ys = [p[1] for p in pts]
        return (max(ys) - min(ys)) > 0.05 and (max(ys) - min(ys)) > (max(xs) - min(xs))


def dedupe(values: Iterable[str]) -> List[str]:
    seen = set()
    out = []
    for v in values:
        if v and v not in seen:
            seen.add(v)
            out.append(v)
    return out


def normalize_text(text: str) -> str:
    return re.sub(r"\s+", "", text or "")


def dark_or_unknown(color: Optional[str]) -> Optional[bool]:
    # 严格三态判定：True=深色可通过，False=非深色，None=无法确认。
    # 原实现把 None/未知色都视为可接受（默认深色），会把主题色误判为合格；
    # 现改为：颜色缺失或主题色无法解析时返回 None，由调用方按“不通过”处理。
    if color is None:
        return None
    color = color.lower()
    if color in {"tx1", "dk1", "dk2", "black"}:
        return True
    if re.fullmatch(r"[0-9a-f]{6}", color):
        # 判定标准：R/G/B 最大值 ≤ 105，视为黑色或深灰色。
        r, g, b = int(color[0:2], 16), int(color[2:4], 16), int(color[4:6], 16)
        return max(r, g, b) <= 105
    # scheme:xxx / bg1/lt1 等无法确认的色标，返回 None 让调用方保守判为不通过。
    return None


def first_size(shape: Shape) -> Optional[float]:
    for run in shape.runs:
        if run.size_pt is not None:
            return run.size_pt
    return None


def is_bold(shape: Shape) -> bool:
    return any(run.bold is True for run in shape.runs)


def is_dark(shape: Shape) -> bool:
    # 严格判定：所有 run 都能被解析为深色才通过；有任意 run 为非深色或“无法确认”则不通过。
    # 这样把之前把未知/主题色默认视作深色的漏洞堵住。
    if not shape.runs:
        return False
    for run in shape.runs:
        verdict = dark_or_unknown(run.color)
        if verdict is not True:
            return False
    return True


def slide_is_blank(slide: Slide, slide_w: float, slide_h: float) -> bool:
    text_len = len(normalize_text(slide.text))
    visible_area = sum(s.bounds.area for s in slide.shapes if s.kind in {"pic", "grp"})
    return text_len < 5 and visible_area < slide_w * slide_h * 0.05


def slide_fingerprint(slide: Slide) -> str:
    parts = [normalize_text(slide.text)]
    parts += sorted(s.media_target for s in slide.shapes if s.media_target)
    return hashlib.sha1("|".join(parts).encode("utf-8", "ignore")).hexdigest()


def find_text_shapes(slide: Slide, *keywords: str) -> List[Shape]:
    result = []
    for s in slide.shapes:
        text = normalize_text(s.text)
        if text and all(k in text for k in keywords):
            result.append(s)
    return result


def is_background_picture(shape: Shape, slide_w: float, slide_h: float) -> bool:
    if shape.kind != "pic":
        return False
    return shape.bounds.area >= slide_w * slide_h * 0.75 and shape.bounds.x <= slide_w * 0.05 and shape.bounds.y <= slide_h * 0.05


def union_bounds(shapes: Sequence[Shape]) -> Optional[Bounds]:
    if not shapes:
        return None
    b = shapes[0].bounds
    for s in shapes[1:]:
        b = b.union(s.bounds)
    return b


def shape_aspect_ok(shape: Shape, tolerance: float = 0.20) -> bool:
    media_aspect = shape.media_aspect
    if media_aspect is None:
        return True
    if media_aspect == 0:
        return False
    return abs(shape.display_aspect / media_aspect - 1.0) <= tolerance


# 左图应为“书本发光/学习成长”相关；右图应为“课堂举手/教学实践”相关。
# 无视觉/OCR 能力，退化为文件名与形状名的元数据关键词匹配（简繁与英文常见词）。
LEFT_SEMANTIC_KEYWORDS = (
    "书", "书本", "書", "学习", "學習", "成长", "成長", "发光", "發光",
    "book", "study", "learn", "grow", "glow", "light",
)
RIGHT_SEMANTIC_KEYWORDS = (
    "课堂", "課堂", "举手", "舉手", "教学", "教學", "实践", "實踐", "课", "課",
    "class", "classroom", "hand", "raise", "teach", "practice", "lesson",
)


def _picture_semantic_tags(shape: Shape) -> Tuple[bool, bool, str]:
    # 返回 (是否命中左侧语义, 是否命中右侧语义, 用于取证的原始标识串)。
    hay_parts = [shape.name or "", shape.media_target or ""]
    hay = " ".join(hay_parts).lower()
    left_hit = any(kw.lower() in hay for kw in LEFT_SEMANTIC_KEYWORDS)
    right_hit = any(kw.lower() in hay for kw in RIGHT_SEMANTIC_KEYWORDS)
    return left_hit, right_hit, " | ".join(p for p in hay_parts if p)


def slide1_picture_pair(slide: Slide, slide_w: float, slide_h: float) -> Tuple[bool, List[Shape], str]:
    target_w = (cm_to_in(6.5), cm_to_in(7.0))
    target_h = (cm_to_in(4.0), cm_to_in(4.5))
    candidates = []
    aspect_failures: List[str] = []
    for s in slide.shapes:
        # pic 对象或带图片填充的圆角矩形形状都视为“图片”候选。
        if not s.is_picture_like or is_background_picture(s, slide_w, slide_h):
            continue
        b = s.bounds
        # 宽 6.5-7cm、高 4-4.5cm，严格按细则区间（零容差）。
        if not almost(b.w, *target_w):
            continue
        if not almost(b.h, *target_h):
            continue
        # 保持圆角矩形裁剪：仅接受圆角矩形 geom（含其常见自定义裁剪存储形式 custGeom）。
        if s.geom not in {"roundRect", "custGeom"}:
            continue
        # 保持原始横纵比：显示比例与媒体原始比例偏差 ≤ 5%（无原始尺寸时视为通过，避免误伤）。
        if not shape_aspect_ok(s, tolerance=0.05):
            aspect_failures.append(
                f"{s.name or s.shape_id}(显示比={s.display_aspect:.3f}, 原始比={s.media_aspect})"
            )
            continue
        candidates.append(s)

    best_pair: List[Shape] = []
    for i, a in enumerate(candidates):
        for b in candidates[i + 1 :]:
            left, right = sorted([a, b], key=lambda x: x.bounds.x)
            pair_union = left.bounds.union(right.bounds)
            gap = right.bounds.x - left.bounds.r
            # 整体位于页面宽度 55%-93%、高度 25%-72% 范围内。
            if not pair_union.inside_percent(slide_w, slide_h, 0.55, 0.93, 0.25, 0.72):
                continue
            # 左右并排，水平间距约 1.3-1.5cm（零容差）。
            if not almost(gap, cm_to_in(1.3), cm_to_in(1.5)):
                continue
            # 顶部基本对齐，高度差不超过 0.4cm。
            if abs(left.bounds.y - right.bounds.y) > cm_to_in(0.4):
                continue
            if abs(left.bounds.h - right.bounds.h) > cm_to_in(0.4):
                continue
            best_pair = [left, right]
            break
        if best_pair:
            break

    if not best_pair:
        parts = [f"找到 {len(candidates)} 个符合尺寸/圆角/比例的候选图片，未形成合规左右图片组。"]
        if aspect_failures:
            parts.append("被比例校验剔除的候选：" + "; ".join(aspect_failures))
        return False, candidates[:2], "；".join(parts)

    left, right = best_pair
    # 图片内容语义：脚本无视觉/OCR 能力，只能依据文件名与形状名元数据做保守判断。
    # 任一图片缺失可识别的语义关键词，或左右语义交叉/错位，均判定不通过。
    l_left, l_right, l_ev = _picture_semantic_tags(left)
    r_left, r_right, r_ev = _picture_semantic_tags(right)
    semantic_ok = l_left and r_right and not l_right and not r_left
    geom_evidence = "; ".join(
        f"{s.name or s.shape_id}({s.bounds.fmt()}, geom={s.geom or '未知'}, "
        f"显示比={s.display_aspect:.3f}, 原始比={s.media_aspect})"
        for s in best_pair
    )
    semantic_evidence = (
        f"左图元数据=[{l_ev or '空'}] 左侧关键词命中={l_left}、右侧关键词命中={l_right}；"
        f"右图元数据=[{r_ev or '空'}] 左侧关键词命中={r_left}、右侧关键词命中={r_right}"
    )
    if not semantic_ok:
        return False, best_pair, geom_evidence + "；" + semantic_evidence + (
            "；无法通过文件名/形状名确认图片内容为‘书本发光/学习成长’与‘课堂举手/教学实践’，"
            "且脚本无视觉/OCR 能力，故按细则保守判定不通过。"
        )
    return True, best_pair, geom_evidence + "；" + semantic_evidence + "；图片语义由文件名/形状名元数据确认。"


def text_group_under_picture(slide: Slide, title: str, subtitle: str, picture: Optional[Shape]) -> Tuple[bool, str]:
    title_shapes = find_text_shapes(slide, title)
    subtitle_shapes = find_text_shapes(slide, subtitle)
    if not title_shapes or not subtitle_shapes:
        return False, f"未同时找到可编辑文字“{title}”和“{subtitle}”。"
    t = title_shapes[0]
    st = subtitle_shapes[0]
    if picture is None:
        return False, f"找到文字但未找到对应上方图片，无法确认其位于图片下方。"
    pic_bottom = picture.bounds.b
    # 文字组位于图片下方：标题与副标题都在图片下边之下。
    below_ok = t.bounds.y >= pic_bottom and st.bounds.y >= pic_bottom
    # 文字组顶部（标题在上）距图片下边约 0.4-1.0cm，严格按区间（零容差）。
    gap = t.bounds.y - pic_bottom
    distance_ok = almost(gap, cm_to_in(0.4), cm_to_in(1.0))
    # 与图片垂直对齐：细则表述为“近似对齐”，允许一定容差。
    # 采用两级容差：绝对偏差 ≤ 0.5cm，或相对图片宽度 ≤ 10%，取更宽松一项。
    align_tol = max(cm_to_in(0.5), picture.bounds.w * 0.10)
    title_delta = abs(t.bounds.cx - picture.bounds.cx)
    subtitle_delta = abs(st.bounds.cx - picture.bounds.cx)
    align_ok = title_delta <= align_tol and subtitle_delta <= align_tol
    if below_ok and distance_ok and align_ok:
        return True, (
            f"{title}/{subtitle} 位于图片 {picture.name or picture.shape_id} 下方并与其垂直对齐，"
            f"距下边 {gap:.3f} 英寸；标题中心偏差 {title_delta:.3f} 英寸、副标题中心偏差 {subtitle_delta:.3f} 英寸"
            f"（容差 {align_tol:.3f} 英寸）。"
        )
    return False, (
        f"找到文字，但位置不满足：标题 {t.bounds.fmt()}（中心偏差 {title_delta:.3f} 英寸），"
        f"副标题 {st.bounds.fmt()}（中心偏差 {subtitle_delta:.3f} 英寸），"
        f"图片 {picture.bounds.fmt()}，距下边 {gap:.3f} 英寸，对齐容差 {align_tol:.3f} 英寸。"
    )


def slide1_text_format(slide: Slide) -> Tuple[bool, str]:
    titles = [find_text_shapes(slide, "持续精进"), find_text_shapes(slide, "实践见效")]
    subtitles = [find_text_shapes(slide, "学习成长与迭代创新"), find_text_shapes(slide, "把方法落实到课堂与服务")]
    flat_titles = [x[0] for x in titles if x]
    flat_subtitles = [x[0] for x in subtitles if x]
    if len(flat_titles) < 2 or len(flat_subtitles) < 2:
        return False, "未找到两组标题/副标题可编辑文本。"
    title_ok = []
    subtitle_ok = []
    details = []
    for s in flat_titles:
        size = first_size(s)
        # 标题：黑色或深灰色、加粗、字号约 18-22 磅（严格按区间，零容差）。
        ok = size is not None and almost(size, 18, 22) and is_bold(s) and is_dark(s)
        title_ok.append(ok)
        details.append(f"标题“{s.text}”字号={size}, 加粗={is_bold(s)}, 深色={is_dark(s)}")
    for s in flat_subtitles:
        size = first_size(s)
        # 副标题：黑色或深灰色、字号约 12-16 磅（严格按区间，零容差）。
        ok = size is not None and almost(size, 12, 16) and is_dark(s)
        subtitle_ok.append(ok)
        details.append(f"副标题“{s.text}”字号={size}, 深色={is_dark(s)}")
    return all(title_ok) and all(subtitle_ok), "；".join(details)


def teacher_card_candidates(slide: Slide, slide_w: float, slide_h: float) -> List[Shape]:
    candidates: List[Shape] = []
    for s in slide.shapes:
        if s.kind not in {"pic", "grp"}:
            continue
        if s.kind == "pic" and is_background_picture(s, slide_w, slide_h):
            continue
        b = s.bounds
        aspect = b.w / b.h if b.h else 0
        # 标准卡片约 8-8.2cm × 14.3-14.5cm，即 3.15-3.23 × 5.63-5.71 英寸；
        # 初始轮播两侧可能缩小，因此给予较宽容的自动化阈值。
        if 2.2 <= b.w <= 3.7 and 4.0 <= b.h <= 6.2 and 0.45 <= aspect <= 0.70:
            candidates.append(s)
    # 去重：组合和其内部图片都像卡片时，保留更外层/面积更大的一个。
    candidates.sort(key=lambda s: s.bounds.area, reverse=True)
    kept: List[Shape] = []
    for s in candidates:
        contained = False
        for k in kept:
            inter = s.bounds.intersect_area(k.bounds)
            if inter >= s.bounds.area * 0.85:
                contained = True
                break
        if not contained:
            kept.append(s)
    return sorted(kept, key=lambda s: s.bounds.x)


# 卡片来源团队关键词：出现在卡片内文本或组名/图片名时，用于判定该卡片属于哪一团队。
TEACHING_TEAM_KEYWORDS = ("教学团队", "教学老师", "教学", "教师团队")
SUPPORT_TEAM_KEYWORDS = ("学习支持团队", "学习支持", "支持团队", "学习顾问", "学管")


def _shapes_inside(card: Shape, slide: Slide, kinds: Tuple[str, ...]) -> List[Shape]:
    # 找出卡片区域内的子形状：面积交集 ≥ 子形状面积 85% 视为“属于该卡”。
    # 用于定位卡片内部的头像图片（pic）或文本框（sp）。
    inside: List[Shape] = []
    for s in slide.shapes:
        if s is card or s.kind not in kinds:
            continue
        if s.bounds.area <= 0:
            continue
        inter = s.bounds.intersect_area(card.bounds)
        if inter >= s.bounds.area * 0.85:
            inside.append(s)
    return inside


def _card_text(card: Shape, slide: Slide) -> str:
    # 汇总卡片自身文本 + 卡片区域内所有子文本框文本，作为该卡片的“可见文字”。
    parts = [card.text or ""]
    for s in _shapes_inside(card, slide, ("sp",)):
        if s.text:
            parts.append(s.text)
    return normalize_text("".join(parts))


def _card_metadata_blob(card: Shape, slide: Slide) -> str:
    # 卡片可用于识别来源的元数据：卡片名 + 卡片内所有 pic 的 media_target + 内部 sp 名称。
    blob_parts = [card.name or "", card.media_target or ""]
    for s in _shapes_inside(card, slide, ("pic", "sp")):
        if s.name:
            blob_parts.append(s.name)
        if s.media_target:
            blob_parts.append(s.media_target)
    return " ".join(blob_parts).lower()


def _classify_card_source(card: Shape, slide: Slide) -> str:
    # 返回 "teaching" / "support" / "unknown"。
    # 优先看卡片内可见文字（作者常在卡内标注“教学”/“学习支持”）；
    # 其次看卡片及内部子对象的形状名/图片文件名（作者常按团队命名素材文件夹）。
    text = _card_text(card, slide)
    meta = _card_metadata_blob(card, slide)
    haystacks = [text, meta]
    for hay in haystacks:
        if any(kw in hay for kw in SUPPORT_TEAM_KEYWORDS):
            # 支持团队关键词更具区分度（“学习支持”不会被“教学”前缀命中），优先判定。
            return "support"
        if any(kw in hay for kw in TEACHING_TEAM_KEYWORDS):
            return "teaching"
    return "unknown"


def _card_headshot(card: Shape, slide: Slide) -> Optional[Shape]:
    # 定位卡片内“头像”图片：卡内所有 pic 中，选面积最大且带原始像素尺寸的一张。
    # 找不到时返回 None（用于把“头像拉伸”与“整卡显示比例”解耦）。
    if card.kind == "pic":
        return card if card.media_size_px else None
    pics = [s for s in _shapes_inside(card, slide, ("pic",)) if s.media_size_px]
    if not pics:
        return None
    return max(pics, key=lambda s: s.bounds.area)


def card_quantity_ok(cards: List[Shape], slide: Slide) -> Tuple[bool, str]:
    if len(cards) != 5:
        return False, f"检测到 {len(cards)} 个独立教师卡片候选，要求 5 个。"
    dims_ok: List[bool] = []
    headshot_ok: List[bool] = []
    details: List[str] = []
    sources: List[str] = []
    for s in cards:
        b = s.bounds
        # 宽约 8-8.2cm、高约 14.3-14.5cm（严格按区间，零容差）。
        dim_ok = almost(b.w, cm_to_in(8.0), cm_to_in(8.2)) and almost(b.h, cm_to_in(14.3), cm_to_in(14.5))
        # 头像不拉伸：定位卡片内头像图片，比较头像自身的显示比例与原始比例（容差 5%）。
        # 若卡片本身即 pic 图片，则退化为卡片整体比例（保留原逻辑，但收紧容差）。
        head = _card_headshot(s, slide)
        if head is None:
            head_ok = False
            head_desc = "未定位到卡片内头像图片，无法确认头像未被拉伸/压缩"
        else:
            head_ok = shape_aspect_ok(head, tolerance=0.05)
            head_desc = (
                f"头像={head.name or head.shape_id}, 显示比={head.display_aspect:.3f}, "
                f"原始比={head.media_aspect}, 保持比例={head_ok}"
            )
        # 卡片来源分类。
        src = _classify_card_source(s, slide)
        sources.append(src)
        dims_ok.append(dim_ok)
        headshot_ok.append(head_ok)
        details.append(f"{s.name or s.shape_id}: {b.fmt()}, 来源={src}; {head_desc}")

    # rubric 要求 3 个来自“教学团队介绍”、2 个来自“学习支持团队介绍”，不能缺少任意一个。
    teaching_count = sources.count("teaching")
    support_count = sources.count("support")
    unknown_count = sources.count("unknown")
    # 严格判定：必须恰好 3 教学 + 2 学习支持；有任何“未知来源”卡片都视为不通过（无法确认齐全）。
    source_ok = teaching_count == 3 and support_count == 2 and unknown_count == 0

    # 页面整体也需同时出现两个团队关键词，作为对逐卡分类的兜底佐证。
    page_text = normalize_text(slide.text)
    has_teaching = "教学团队" in page_text
    has_support = "学习支持团队" in page_text
    page_ok = has_teaching and has_support

    ok = all(dims_ok) and all(headshot_ok) and source_ok and page_ok
    return ok, (
        "；".join(details)
        + f"；来源计数：教学={teaching_count}(要求3)、学习支持={support_count}(要求2)、未知={unknown_count}(要求0)；"
        + f"页面关键词：教学团队={has_teaching}、学习支持团队={has_support}。"
    )


def card_layout_ok(cards: List[Shape], slide_w: float, slide_h: float) -> Tuple[bool, str]:
    if len(cards) != 5:
        return False, "未检测到 5 个独立卡片，无法判定横向轮播队列布局。"
    overall = union_bounds(cards)
    assert overall is not None
    # 整体位于页面宽度 5%-95%、高度 18%-92% 范围内。
    bounds_ok = overall.inside_percent(slide_w, slide_h, 0.05, 0.95, 0.18, 0.92)
    # 横向轮播队列排列：卡片按中心 x 从左到右依次排列。
    ordered = sorted(cards, key=lambda c: c.bounds.cx)
    horizontal_order = [c.bounds.cx for c in cards] == [c.bounds.cx for c in ordered]

    # 中间卡片：最靠近页面水平中线的一张。
    center_card = min(cards, key=lambda c: abs(c.bounds.cx - slide_w / 2))
    side_cards = [c for c in cards if c is not center_card]
    avg_side_area = sum(c.bounds.area for c in side_cards) / 4 if side_cards else 0.0

    # 中间卡片突出显示：面积 ≥ 两侧均值的 1.05 倍（保留原判据，作为“突出”的必要条件）。
    prominent = avg_side_area > 0 and center_card.bounds.area >= avg_side_area * 1.05

    # rubric 关键补强：两侧卡片必须“部分可见或缩小显示”。
    # 逐张两侧卡片满足以下任一条件即视为“非等大满铺”，符合轮播观感：
    #   (a) 缩小显示：面积 ≤ 中间卡片的 90%（宽或高任一 ≤ 中间卡片对应值的 95%）；
    #   (b) 部分可见：与页面可见区 [0, slide_w] 存在裁切（左/右侧超出页面），
    #       或被相邻卡片遮挡（与更靠内的卡片存在水平重叠）。
    side_visibility_details: List[str] = []
    side_visibility_flags: List[bool] = []
    for c in side_cards:
        b = c.bounds
        shrunk = (
            b.area <= center_card.bounds.area * 0.90
            or b.w <= center_card.bounds.w * 0.95
            or b.h <= center_card.bounds.h * 0.95
        )
        # 页面裁切：卡片左边超出 0（或右边超出 slide_w），可视为“被页面裁切，部分可见”。
        clipped_by_page = b.x < -1e-3 or b.r > slide_w + 1e-3
        # 相邻遮挡：与中心卡片或另一侧相邻卡片在水平方向重叠（右边界越过邻居左边界）。
        # 这里放宽为“与中心卡片水平重叠”即视为轮播叠放结构。
        overlaps_neighbor = b.intersect_area(center_card.bounds) > 0
        partial_visible = clipped_by_page or overlaps_neighbor
        ok = shrunk or partial_visible
        side_visibility_flags.append(ok)
        side_visibility_details.append(
            f"{c.name or c.shape_id}(缩小={shrunk}, 裁切={clipped_by_page}, "
            f"与中心卡水平重叠={overlaps_neighbor}, 面积比={b.area / center_card.bounds.area:.2f})"
        )
    sides_partial_or_shrunk = all(side_visibility_flags)

    # 轮播结构：非等距完全并排。等距完全并排的判据 = 相邻卡片间距接近相等且无重叠、无裁切、无缩小。
    # 若同时出现“至少 1 张两侧卡片被裁切/被遮挡”或“两侧卡与中心卡尺寸有明显差异”，即视为轮播结构。
    xs = [c.bounds.cx for c in ordered]
    gaps = [xs[i + 1] - xs[i] for i in range(len(xs) - 1)]
    even_spacing = False
    if gaps and min(gaps) > 0:
        even_spacing = (max(gaps) - min(gaps)) / max(gaps) <= 0.10  # 相邻间距差 ≤ 10% 视为等距
    widths = [c.bounds.w for c in cards]
    similar_size = (max(widths) - min(widths)) / max(widths) <= 0.05 if max(widths) > 0 else True
    # 检测“完整等距等大”的反面模式：等距 + 等大 + 无裁切 + 无遮挡。
    any_clip_or_overlap = any(
        (c.bounds.x < -1e-3 or c.bounds.r > slide_w + 1e-3 or c.bounds.intersect_area(center_card.bounds) > 0)
        for c in side_cards
    )
    is_full_even_row = even_spacing and similar_size and not any_clip_or_overlap
    carousel_structure = not is_full_even_row

    ok = bounds_ok and horizontal_order and prominent and sides_partial_or_shrunk and carousel_structure
    return ok, (
        f"整体范围 {overall.fmt()}，范围合规={bounds_ok}，横向排序={horizontal_order}，"
        f"中间卡片突出={prominent}(中心面积={center_card.bounds.area:.2f} vs 两侧均值={avg_side_area:.2f})，"
        f"两侧卡片部分可见或缩小={sides_partial_or_shrunk}[{'; '.join(side_visibility_details)}]，"
        f"轮播结构（非等距等大满铺）={carousel_structure}(等距={even_spacing}, 等大={similar_size}, "
        f"裁切或遮挡={any_clip_or_overlap})。"
    )


def animation_effect_ok(slide: Slide, cards: List[Shape]) -> Tuple[bool, str]:
    if len(cards) != 5:
        return False, "未检测到 5 个独立卡片，无法确认每张卡片的轮播动画。"
    card_ids = {c.shape_id for c in cards if c.shape_id}
    anim = slide.animation
    anim_targets = set(anim.target_ids)
    motion_targets = set(anim.motion_target_ids)
    horizontal_targets = set(anim.horizontal_motion_ids)

    # === 1) 依次横向移动：每张卡片都要出现在时间线中，且触发顺序覆盖所有卡片 ===
    individually_targeted = len(card_ids & anim_targets) >= 5 or len(card_ids & motion_targets) >= 5
    # 时间线顺序：至少 5 张不同卡片作为动画目标依序出现，认为“依次”播放。
    seen_in_order: List[str] = []
    for sid in anim.timeline_order:
        if sid in card_ids and sid not in seen_in_order:
            seen_in_order.append(sid)
    sequential_ok = len(seen_in_order) >= 5
    # 每卡都有 delay 或都在时间线中，认为不是同时触发。
    delays_per_card = [len(anim.delays_by_shape.get(c.shape_id, [])) for c in cards]
    staggered_ok = sum(1 for d in delays_per_card if d > 0) >= 4 or sequential_ok

    # === 2) 水平移动为主，禁用纵向、飞入、旋转 ===
    horizontal_ok = (
        len(card_ids & horizontal_targets) >= 4  # 至少 4 张卡片有明确水平位移
        or (len(anim.paths) >= 5 and len(horizontal_targets) >= 4)
    )
    has_enough_animation = anim.raw_anim_count >= 5
    no_rotation = anim.rotation_anim_count == 0
    no_fly_in = anim.fly_in_count == 0
    no_vertical = len(set(anim.vertical_motion_ids) & card_ids) == 0 and len(anim.vertical_motion_ids) == 0
    forbidden_clear = no_rotation and no_fly_in and no_vertical

    # === 3) 逐卡验证“轮到中间时”的突出（尺寸 1.1-1.4 倍 或 层级最上层） ===
    # 初始布局中面积最大者作为“当前中心卡”的参考；两侧卡片取剩余 4 张的均值。
    init_center = max(cards, key=lambda c: c.bounds.area)
    init_sides = [c for c in cards if c is not init_center]
    avg_side_w = sum(c.bounds.w for c in init_sides) / len(init_sides) if init_sides else 0.0
    avg_side_h = sum(c.bounds.h for c in init_sides) / len(init_sides) if init_sides else 0.0

    per_card_prominent: List[bool] = []
    per_card_detail: List[str] = []
    for c in cards:
        # (a) 尺寸判据：卡片本身尺寸 ≥ 两侧均值 1.1-1.4 倍；或 animScale 上出现放大到 1.1-1.4 的关键帧。
        w_ratio = c.bounds.w / avg_side_w if avg_side_w else 0.0
        h_ratio = c.bounds.h / avg_side_h if avg_side_h else 0.0
        static_scale_ok = almost(w_ratio, 1.1, 1.4) and almost(h_ratio, 1.1, 1.4)
        scale_keyframes = anim.scale_by_shape.get(c.shape_id, [])
        dyn_scale_ok = any(1.10 <= sx <= 1.40 and 1.10 <= sy <= 1.40 for sx, sy in scale_keyframes)
        size_ok = static_scale_ok or dyn_scale_ok
        # (b) 层级判据：卡片 z_index 最高（作为“最上层”的近似），说明存在明显层级突出。
        max_z = max(x.z_index for x in cards)
        layer_top = c.z_index == max_z
        # rubric：“尺寸或层级明显突出” —— 二者之一即可。
        prominent = size_ok or layer_top
        per_card_prominent.append(prominent)
        per_card_detail.append(
            f"{c.name or c.shape_id}: w比={w_ratio:.2f}, h比={h_ratio:.2f}, "
            f"静态突出={static_scale_ok}, 动态缩放突出={dyn_scale_ok}, 层级最上={layer_top}"
        )
    # 至少每张卡片“轮到中间时”都能被突出。用两个信号的组合退让：
    #   - 若每卡都有 dyn_scale_ok（放大关键帧命中）→ 严格通过；
    #   - 否则要求：至少一张卡片满足尺寸/层级突出（初始中心卡符合 1.1-1.4 倍），
    #     且时间线里 5 张卡片按序都被触发（sequential_ok）——
    #     即“通过依次成为中心时的尺寸/位置轮换”呈现突出效果。
    dyn_all = all(
        any(1.10 <= sx <= 1.40 and 1.10 <= sy <= 1.40 for sx, sy in anim.scale_by_shape.get(c.shape_id, []))
        for c in cards
    )
    prominence_ok = dyn_all or (any(per_card_prominent) and sequential_ok)

    # === 4) 中心区域：至少一张卡片中心在页面水平 [35%, 65%] 范围内 ===
    slide_w = 0.0
    for c in cards:
        # 从卡片父 slide 借用宽度不方便；从 slide.shapes 推断页面宽度：max 右边界。
        slide_w = max(slide_w, c.bounds.r)
    # 用最靠近水平中线的卡片检验；页面宽度从 slide.shapes 推断难免有误差，改用参数化：
    # 这里 slide_w 只是兜底，函数入口未显式传入 slide_w；用整页文本无关的近似。
    # 更稳的做法：直接使用初始中心卡的中心横坐标是否位于所有卡片 x 极值的中间 30%。
    xs = [c.bounds.cx for c in cards]
    x_min, x_max = min(xs), max(xs)
    center_band_low = x_min + (x_max - x_min) * 0.35
    center_band_high = x_min + (x_max - x_min) * 0.65
    in_center_region = center_band_low <= init_center.bounds.cx <= center_band_high

    # === 5) 透明度：当前中心卡片透明度 100%，两侧卡片可以降低 ===
    center_alpha_vals = anim.alpha_values_by_shape.get(init_center.shape_id, [])
    # 100% 视为“达到 95 及以上”即通过；若无 alpha 动画视为始终 100%。
    center_alpha_ok = (not center_alpha_vals) or any(a >= 95.0 for a in center_alpha_vals)

    # === 6) 两侧卡片：向左右移动 或 缩小 或 透明度降低（rubric 明确列出三种）===
    side_effect_ok_list: List[bool] = []
    side_effect_details: List[str] = []
    for c in init_sides:
        moved = c.shape_id in horizontal_targets or c.shape_id in motion_targets
        scale_kf = anim.scale_by_shape.get(c.shape_id, [])
        shrunk_dyn = any(sx < 0.98 or sy < 0.98 for sx, sy in scale_kf)
        shrunk_static = c.bounds.area < init_center.bounds.area * 0.95
        alpha_vals = anim.alpha_values_by_shape.get(c.shape_id, [])
        alpha_down = any(a < 95.0 for a in alpha_vals)
        ok_this = moved or shrunk_dyn or shrunk_static or alpha_down
        side_effect_ok_list.append(ok_this)
        side_effect_details.append(
            f"{c.name or c.shape_id}(移动={moved}, 动态缩小={shrunk_dyn}, "
            f"静态缩小={shrunk_static}, 透明度降低={alpha_down})"
        )
    sides_effect_ok = all(side_effect_ok_list)

    # === 7) 非当前展示卡片不遮挡中间卡片“主要头像和信息区” ===
    # 用中心卡片的中央 60% × 60% 作为“主要头像+信息区”核心带；
    # 中心卡片 z_index 应 ≥ 各侧边卡片 z_index，两侧任一张 z_index 高于中心且与该核心带存在重叠即视为遮挡。
    core = Bounds(
        init_center.bounds.x + init_center.bounds.w * 0.20,
        init_center.bounds.y + init_center.bounds.h * 0.20,
        init_center.bounds.w * 0.60,
        init_center.bounds.h * 0.60,
    )
    occluders: List[str] = []
    for c in init_sides:
        if c.z_index > init_center.z_index and c.bounds.intersect_area(core) > 0:
            occluders.append(c.name or c.shape_id)
    no_occlusion = not occluders

    ok = (
        individually_targeted
        and sequential_ok
        and staggered_ok
        and horizontal_ok
        and has_enough_animation
        and forbidden_clear
        and prominence_ok
        and in_center_region
        and center_alpha_ok
        and sides_effect_ok
        and no_occlusion
    )
    return ok, (
        f"依次触发顺序={seen_in_order}(需≥5)，逐卡 delay 出现数={delays_per_card}；"
        f"水平motion目标={sorted(horizontal_targets)}，动画节点数={anim.raw_anim_count}；"
        f"逐卡突出：[{'; '.join(per_card_detail)}]，动态放大覆盖全部={dyn_all}，突出通过={prominence_ok}；"
        f"中心卡在中心区域={in_center_region}，中心卡α值={center_alpha_vals}(≥95%即视为100%)={center_alpha_ok}；"
        f"两侧卡效应：[{'; '.join(side_effect_details)}]，两侧通过={sides_effect_ok}；"
        f"遮挡卡片={occluders}，不遮挡={no_occlusion}；"
        f"旋转数={anim.rotation_anim_count}、飞入数={anim.fly_in_count}、纵向位移目标={anim.vertical_motion_ids}。"
    )


def animation_duration_ok(slide: Slide, cards: List[Shape]) -> Tuple[bool, str]:
    durations = [d for d in slide.animation.durations_ms if 100 <= d <= 60000]
    if not durations:
        return False, "未找到可解析的动画持续时间。"
    # 只统计显式动画 dur。复杂时间线可能有并行节点，此处采用总和代理“一轮总时长”。
    total_sec = sum(durations) / 1000.0
    # 若 XML 包含很多嵌套容器 dur，sum 会偏大；同时用最大 dur 作为兜底。
    max_sec = max(durations) / 1000.0
    plausible_total = total_sec if total_sec <= 30 else max_sec
    # 总时长约 8-18 秒（严格按区间，零容差）。
    total_ok = almost(plausible_total, 8, 18)
    # 单个卡片居中停留约 1-3 秒（零容差）。
    per_card = plausible_total / 5 if cards else 0
    per_card_ok = almost(per_card, 1, 3)
    ok = total_ok and per_card_ok
    return ok, (
        f"解析到持续时间 {durations} ms，估算一轮 {plausible_total:.1f}s，单卡约 {per_card:.1f}s，"
        f"总时长合规={total_ok}，单卡合规={per_card_ok}。"
    )


def animation_trigger_ok(slide: Slide, cards: List[Shape], slide_w: float) -> Tuple[bool, str]:
    if len(cards) != 5:
        return False, "未检测到 5 个独立卡片，无法确认连续触发与路径边界。"
    anim = slide.animation
    anim_targets = set(anim.target_ids)
    card_ids = {c.shape_id for c in cards if c.shape_id}
    horizontal_targets = set(anim.horizontal_motion_ids)

    # 单击开始或自动开始；开始后连续播放，不需手动点击 5 次：点击触发条件 ≤1。
    click_ok = anim.click_trigger_count <= 1
    # 开始后卡片按顺序连续播放：5 个卡片都进入动画目标且动画节点足量。
    continuous_ok = len(card_ids & anim_targets) >= 5 and anim.raw_anim_count >= 5
    # 每个教师卡片移动路径为水平或近似水平直线：5 个卡片路径都水平。
    paths_horizontal = len(horizontal_targets & card_ids) >= 5
    # 路径不超出幻灯片左右边界：解析每个卡片路径终点 x，校验初始位置叠加位移后仍在 [0, slide_w] 内。
    inside_count = 0
    for c in cards:
        x_offsets = anim.motion_x_offsets.get(c.shape_id, [])
        if not x_offsets:
            # 无可解析位移：仅按初始位置判断是否在页面内。
            in_bounds = c.bounds.x >= -0.05 and c.bounds.r <= slide_w + 0.05
        else:
            # motion path 偏移以幻灯片宽度比例表示，换算为英寸后叠加到卡片左/右极端位置。
            max_shift = max(x_offsets) * slide_w
            left_most = c.bounds.x - max_shift
            right_most = c.bounds.r + max_shift
            in_bounds = left_most >= -0.05 and right_most <= slide_w + 0.05
        if in_bounds:
            inside_count += 1
    path_inside = inside_count == 5

    # 中间卡片位于最上层，不被两侧卡片遮挡：中心卡片 z_index 高于所有两侧卡片。
    center_card = min(cards, key=lambda c: abs(c.bounds.cx - slide_w / 2))
    top_most = all(center_card.z_index >= c.z_index for c in cards)

    ok = click_ok and continuous_ok and paths_horizontal and path_inside and top_most
    return ok, (
        f"点击触发数={anim.click_trigger_count}，连续目标足够={continuous_ok}，"
        f"每卡水平路径={paths_horizontal}，路径不越界卡片数={inside_count}/5，中心卡片最上层={top_most}。"
    )


def has_merged_teacher_cards(slide: Slide, cards: List[Shape], slide_w: float, slide_h: float) -> Tuple[bool, str]:
    # 5 个教师卡片被合并成一张图片：卡片区域呈现为单张横向大图（宽度大、高度中等），
    # 且独立可编辑卡片候选少于 5 个。
    large_pics = []
    for s in slide.shapes:
        if s.kind != "pic" or is_background_picture(s, slide_w, slide_h):
            continue
        b = s.bounds
        # 横向合并图典型为“宽大高中”：宽度 ≥ 页面 60%、高度 ≥ 页面 40%。
        if b.w >= slide_w * 0.60 and b.h >= slide_h * 0.40:
            large_pics.append(s)
    if large_pics and len(cards) < 5:
        return True, "检测到单张横向大图（宽大高中）覆盖卡片区域，且独立卡片候选少于 5 个：" + "; ".join(f"{s.name or s.shape_id}({s.bounds.fmt()})" for s in large_pics)
    return False, "未发现单张大图替代 5 个独立教师卡片的明显迹象。"


def animation_pane_not_individual(slide: Slide, cards: List[Shape], merged: bool) -> Tuple[bool, str]:
    # 无法单独修改每个卡片的“动画顺序”或“持续时间”即扣分。
    if merged:
        return True, "教师卡片疑似合并为单张图片，动画窗格无法单独修改每个卡片的顺序/持续时间。"
    if len(cards) < 5:
        return True, f"独立卡片候选仅 {len(cards)} 个，无法保证 5 个卡片动画可单独修改。"
    card_ids = {c.shape_id for c in cards if c.shape_id}
    anim = slide.animation
    anim_targets = set(anim.target_ids)
    # 顺序可逐卡调整：每个卡片都有独立动画条目（进入动画目标）。
    targeted = len(card_ids & anim_targets)
    order_editable = targeted >= 5
    # 时长可逐卡调整：可解析的动画持续时间条目至少 5 个（每卡一条 dur）。
    duration_count = len([d for d in anim.durations_ms if d > 0])
    duration_editable = duration_count >= 5
    bad = not (order_editable and duration_editable)
    return bad, (
        f"5 个卡片中有 {targeted} 个出现在动画目标中（顺序可调={order_editable}），"
        f"可解析动画时长条目数={duration_count}（时长可调={duration_editable}）。"
    )


def first_slide_moved_after_seventh(slides: List[Slide]) -> Tuple[bool, str]:
    first_has = "公司简介" in normalize_text(slides[0].text) if slides else False
    later = [s.index for s in slides[7:] if "公司简介" in normalize_text(s.text)]
    if not first_has and later:
        return True, f"第 1 页未检测到“公司简介”，但第 {later} 页检测到该标题。"
    return False, "第 1 页包含“公司简介”，未发现被移动到第 7 页之后。"


def dimension1(analyzer: PptxAnalyzer) -> Tuple[bool, List[str]]:
    # 交付文件为 .pptx 格式，文件可正常打开。
    # 判定要点：
    #   1) 扩展名必须为 .pptx（不再接受 .ppt）。
    #   2) 文件可作为 PPTX 正常打开（能读取 presentation.xml 中的幻灯片列表）。
    reasons = []
    suffix = analyzer.path.suffix.lower()
    if suffix != ".pptx":
        reasons.append(f"文件扩展名为 {suffix}，不是 .pptx。")
    else:
        try:
            if not analyzer.slide_paths:
                reasons.append("无法读取 presentation.xml 中的幻灯片列表，文件无法正常打开。")
        except Exception as exc:
            reasons.append(f"PPTX 文件无法正常打开：{exc}")
    if len(analyzer.slide_paths) != 11:
        reasons.append(f"最终页数为 {len(analyzer.slide_paths)}，不是 11 页。")
    return not reasons, reasons


def evaluate_dimension2(analyzer: PptxAnalyzer) -> List[PointResult]:
    slides = analyzer.slides
    slide1 = slides[0]
    slide7 = slides[6] if len(slides) >= 7 else None
    w, h = analyzer.slide_w, analyzer.slide_h
    results: List[PointResult] = []

    # 正向得分点
    pic_pair_ok, pic_pair, pic_evidence = slide1_picture_pair(slide1, w, h)
    results.append(PointResult(+5, "第1页右侧图片组：两张图片作为组合整体移动到第1页右侧居中区域，整体位于页面宽度55%—93%、页面高度25%—72%范围内。两张图片宽度均为6.5-7厘米，高度均为4-4.5厘米，保持圆角矩形裁剪。两张图片左右并排，水平间距约1.3-1.5厘米，顶部基本对齐，高度差不超过0.4厘米。", pic_pair_ok, pic_evidence))

    left_pic = pic_pair[0] if len(pic_pair) >= 2 else None
    right_pic = pic_pair[1] if len(pic_pair) >= 2 else None
    ok, ev = text_group_under_picture(slide1, "持续精进", "学习成长与迭代创新", left_pic)
    results.append(PointResult(+1, "第1页“持续精进”文字组：位于左侧新增图片下方，包含“持续精进”和“学习成长与迭代创新”，与图片垂直对齐，距图片下边约0.4—1.0厘米。", ok, ev))
    ok, ev = text_group_under_picture(slide1, "实践见效", "把方法落实到课堂与服务", right_pic)
    results.append(PointResult(+1, "第1页“实践见效”文字组：位于右侧新增图片下方，包含“实践见效”和“把方法落实到课堂与服务”，与图片垂直对齐，距图片下边约0.4—1.0厘米。", ok, ev))
    ok, ev = slide1_text_format(slide1)
    results.append(PointResult(+1, "第1页新增文字格式：两组标题文字为黑色或深灰色加粗，字号约18-22磅；副标题文字为黑色或深灰色，字号约12-16磅，均可编辑。", ok, ev))

    if slide7 is None:
        dummy = "缺少第 7 页。"
        for score, name in [
            (+5, "第7页教师卡片数量：第7页包含5个老师信息介绍卡片，其中3个来自“教学团队介绍”，2个来自“学习支持团队介绍”，不能缺少任意一个。5个教师卡片均保持原始横纵比，宽度约8-8.2厘米，高度约14.3-14.5厘米，人物头像没有横向拉伸或纵向压缩。"),
            (+3, "第7页5个教师卡片初始布局：播放前5个卡片以横向轮播队列形式排列，中间卡片突出显示，两侧卡片部分可见或缩小显示，整体位于页面宽度5%—95%、页面高度18%—92%范围内。"),
            (+5, "第7页轮播动画效果：5个教师卡片播放时依次横向移动，形成类似视频中的卡片轮播效果；每次至少有1个卡片位于中间突出位置，两侧卡片向左右移动或缩小。每个教师卡片轮到中间时，尺寸或层级明显突出，宽度和高度约为两侧卡片的1.1—1.4倍，透明度为100%，位于页面中心区域。非当前展示卡片位于左右两侧时，尺寸缩小或透明度降低，且不遮挡中间卡片主要头像和信息区。卡片整体采用水平左右移动方式，不能使用随机飞入、旋转散开或与视频不相符的上下跳动效果。"),
            (+1, "第7页动画时长：5个教师卡片完整展示一轮的总时长约8—18秒，单个卡片居中停留约1—3秒。"),
            (+3, "第7页动画触发方式：动画可设置为单击开始或自动开始；开始后卡片按顺序连续播放，不需要手动点击5次才能看完整轮。每个教师卡片的移动路径为水平或近似水平直线路径，路径不超出幻灯片左右边界。轮播动画中当前中间卡片位于最上层，两侧卡片位于其后方或边侧，不出现中间卡片被侧边卡片遮挡的情况。"),
        ]:
            results.append(PointResult(score, name, False, dummy))
        return results

    cards = teacher_card_candidates(slide7, w, h)
    ok, ev = card_quantity_ok(cards, slide7)
    results.append(PointResult(+5, "第7页教师卡片数量：第7页包含5个老师信息介绍卡片，其中3个来自“教学团队介绍”，2个来自“学习支持团队介绍”，不能缺少任意一个。5个教师卡片均保持原始横纵比，宽度约8-8.2厘米，高度约14.3-14.5厘米，人物头像没有横向拉伸或纵向压缩。", ok, ev))
    ok, ev = card_layout_ok(cards, w, h)
    results.append(PointResult(+3, "第7页5个教师卡片初始布局：播放前5个卡片以横向轮播队列形式排列，中间卡片突出显示，两侧卡片部分可见或缩小显示，整体位于页面宽度5%—95%、页面高度18%—92%范围内。", ok, ev))
    ok, ev = animation_effect_ok(slide7, cards)
    results.append(PointResult(+5, "第7页轮播动画效果：5个教师卡片播放时依次横向移动，形成类似视频中的卡片轮播效果；每次至少有1个卡片位于中间突出位置，两侧卡片向左右移动或缩小。每个教师卡片轮到中间时，尺寸或层级明显突出，宽度和高度约为两侧卡片的1.1—1.4倍，透明度为100%，位于页面中心区域。非当前展示卡片位于左右两侧时，尺寸缩小或透明度降低，且不遮挡中间卡片主要头像和信息区。卡片整体采用水平左右移动方式，不能使用随机飞入、旋转散开或与视频不相符的上下跳动效果。", ok, ev))
    ok, ev = animation_duration_ok(slide7, cards)
    results.append(PointResult(+1, "第7页动画时长：5个教师卡片完整展示一轮的总时长约8—18秒，单个卡片居中停留约1—3秒。", ok, ev))
    ok, ev = animation_trigger_ok(slide7, cards, w)
    results.append(PointResult(+3, "第7页动画触发方式：动画可设置为单击开始或自动开始；开始后卡片按顺序连续播放，不需要手动点击5次才能看完整轮。每个教师卡片的移动路径为水平或近似水平直线路径，路径不超出幻灯片左右边界。轮播动画中当前中间卡片位于最上层，两侧卡片位于其后方或边侧，不出现中间卡片被侧边卡片遮挡的情况。", ok, ev))

    # 扣分点
    merged, ev = has_merged_teacher_cards(slide7, cards, w, h)
    results.append(PointResult(-5, "扣分：第7页5个教师卡片被合并成一张不可编辑图片", merged, ev))
    pane_bad, ev = animation_pane_not_individual(slide7, cards, merged)
    results.append(PointResult(-3, "扣分：第7页动画窗格中无法单独修改每个卡片的动画顺序或持续时间。", pane_bad, ev))

    return results


def _build_dim2_items(results: List[PointResult]) -> List[dict]:
    items: List[dict] = []
    for item in results:
        # rule 使用原细则文案（扣分项保留“扣分：”前缀，便于与正向得分点区分）。
        items.append({
            "rule": item.name,
            "max_delta": item.score,
            "delta": item.score if item.hit else 0,
            "hit": item.hit,
            "detail": "",
        })
    return items


def _locate_pptx(dir_path: Path) -> Path:
    # 在脚本所在目录内定位被评估的 .pptx 文件；忽略 Office 临时文件（~$ 开头）。
    # 仅识别 .pptx（不再接受 .ppt 二进制格式）。
    if not dir_path.exists() or not dir_path.is_dir():
        raise FileNotFoundError(f"目录不存在或不是目录：{dir_path}")
    preferred_name = "公司简介_动画合并版.pptx"
    preferred = dir_path / preferred_name
    if preferred.exists():
        return preferred
    candidates: List[Path] = []
    for p in sorted(dir_path.iterdir()):
        if not p.is_file():
            continue
        if p.name.startswith("~$"):
            continue
        if p.suffix.lower() == ".pptx":
            candidates.append(p)
    if not candidates:
        raise FileNotFoundError(f"目录 {dir_path} 下未找到 .pptx 文件。")
    return candidates[0]


def evaluate(dir_path: str) -> dict:
    """评估入口：接收脚本所在目录路径，在其中定位并评估被评估的 PPT 文档。"""
    result: dict = {
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
        base_dir = Path(dir_path)
        pptx_path = _locate_pptx(base_dir)
        result["file_name"] = pptx_path.name

        if not zipfile.is_zipfile(pptx_path):
            result["dim1_pass"] = False
            result["dim1_reason"] = f"文件不是有效 PPTX/ZIP：{pptx_path.name}"
            return result

        with PptxAnalyzer(pptx_path) as analyzer:
            d1_ok, d1_reasons = dimension1(analyzer)
            result["dim1_pass"] = d1_ok
            result["dim1_reason"] = "" if d1_ok else "；".join(d1_reasons)
            if not d1_ok:
                # 维度一不通过：不进入维度二评分，total_score 记 0。
                return result

            dim2_results = evaluate_dimension2(analyzer)
            result["dim2_items"] = _build_dim2_items(dim2_results)
            # 得分 = 命中的加分项 + 命中的减分项之和（即所有 hit 项的 delta 之和）。
            result["total_score"] = sum(it["delta"] for it in result["dim2_items"])
            # 满分（“总分”）= 所有加分项 max_delta 之和；减分项不计入满分。
            result["max_score"] = sum(it["max_delta"] for it in result["dim2_items"] if it["max_delta"] > 0)
            return result
    except zipfile.BadZipFile as exc:
        result["status"] = "error"
        result["error"] = f"文件无法作为 PPTX 打开：{exc}"
        return result
    except ET.ParseError as exc:
        result["status"] = "error"
        result["error"] = f"PPTX XML 损坏或无法解析：{exc}"
        return result
    except FileNotFoundError as exc:
        result["status"] = "error"
        result["error"] = str(exc)
        return result
    except Exception as exc:  # 兜底：任何未预期异常都归为 error，便于批量运行区分“评估 0 分”与“脚本崩溃”。
        result["status"] = "error"
        result["error"] = f"{type(exc).__name__}: {exc}"
        return result


if __name__ == "__main__":
    import json

    _target_dir = sys.argv[1] if len(sys.argv) > 1 else str(Path(__file__).resolve().parent)
    print(json.dumps(evaluate(_target_dir), ensure_ascii=False, indent=2))
