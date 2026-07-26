#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
自动评估 Word 申报书中"项目介绍材料"行是否正确插入 PPT。

对外只暴露一个函数：

    def evaluate(dir_path: str) -> dict

调用方传入"脚本所在目录的路径"，脚本自己在该目录中定位并打开被评估的 .docx
文档，返回结构化字典（参见评估脚本统一约定 §2.2）。

说明：
- 只使用 Python 标准库，直接解析 docx/pptx 的 OOXML 压缩包结构。
- Word 的真实分页依赖排版引擎，OOXML 内通常不记录"第 N 页"的精确边界；
  本脚本用文档页数属性、显式分页符、页眉页脚、目标表格行位置等结构信息做自动化近似判断。
- 对"可点击/可编辑/不遮挡/不越界"等视觉或交互项，采用 OOXML 中的对象类型、关系、尺寸、单元格位置、边框等可机器检测证据来判定。
"""

from __future__ import annotations

import io
import json

SCRIPT_ID = "059"

import os
import posixpath
import re
import sys
import zipfile
from dataclasses import dataclass, field
from pathlib import Path
from typing import Dict, List, Optional, Tuple
from xml.etree import ElementTree as ET

NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "o": "urn:schemas-microsoft-com:office:office",
    "v": "urn:schemas-microsoft-com:vml",
    "wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "pic": "http://schemas.openxmlformats.org/drawingml/2006/picture",
    "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
    "ep": "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties",
    "mc": "http://schemas.openxmlformats.org/markup-compatibility/2006",
}

W = "{%s}" % NS["w"]
R = "{%s}" % NS["r"]
O = "{%s}" % NS["o"]
V = "{%s}" % NS["v"]
REL = "{%s}" % NS["rel"]

EXPECTED_TITLE_PATTERNS = [
    "海潮艺境MR共创计划",
    "海潮艺境 MR 共创计划",
    "项目介绍材料PPT",
]
PPT_CONTENT_KEYWORDS = [
    "海潮艺境MR共创计划",
    "海潮艺境 MR 共创计划",
    "以滨海手作、互动叙事与数字展演",
    "重构区域文化体验",
]
CLICK_TARGET_PPT_FILENAME = "海潮艺境MR共创计划.pptx"
CLICK_TITLE_PAGE_REQUIRED_TEXT = [
    "海潮艺境 MR 共创计划",
    "以滨海手作、互动叙事与数字展演，重构区域文化体验",
]
TARGET_LABEL = "项目介绍材料"


@dataclass
class CheckResult:
    code: str
    score: int
    hit: bool
    detail: str


@dataclass
class Evaluation:
    dimension1_passed: bool = False
    dimension1_details: List[str] = field(default_factory=list)
    checks: List[CheckResult] = field(default_factory=list)
    total: int = 0


@dataclass
class DocContext:
    path: Path
    z: zipfile.ZipFile
    names: List[str]
    document_root: ET.Element
    rels: Dict[str, Dict[str, str]]
    app_root: Optional[ET.Element]
    header_roots: List[ET.Element]
    footer_roots: List[ET.Element]


def normalize_text(s: str) -> str:
    return re.sub(r"\s+", "", s or "")


def text_of(el: Optional[ET.Element]) -> str:
    if el is None:
        return ""
    return "".join(t.text or "" for t in el.findall(".//w:t", NS))


def all_text_like(el: Optional[ET.Element]) -> str:
    """Collect visible text and common metadata/title strings from an OOXML element."""
    if el is None:
        return ""
    chunks: List[str] = []
    for node in el.iter():
        if node.text:
            chunks.append(node.text)
        for key, val in node.attrib.items():
            local = key.split("}")[-1].lower()
            if local in {"title", "descr", "name", "tooltip", "target", "progid"} and val:
                chunks.append(val)
    return "".join(chunks)


def parse_xml(data: bytes, part: str) -> ET.Element:
    try:
        return ET.fromstring(data)
    except ET.ParseError as exc:
        raise ValueError(f"XML 解析失败：{part}: {exc}") from exc


def read_optional_xml(z: zipfile.ZipFile, part: str) -> Optional[ET.Element]:
    try:
        return parse_xml(z.read(part), part)
    except KeyError:
        return None


def parse_rels(z: zipfile.ZipFile, rels_part: str) -> Dict[str, Dict[str, str]]:
    try:
        root = parse_xml(z.read(rels_part), rels_part)
    except KeyError:
        return {}
    result = {}
    for rel in root.findall("rel:Relationship", NS):
        rid = rel.attrib.get("Id")
        if rid:
            result[rid] = dict(rel.attrib)
    return result


def relationship_target(ctx: DocContext, rid: str) -> Optional[str]:
    rel = ctx.rels.get(rid)
    if not rel:
        return None
    target = rel.get("Target", "")
    if rel.get("TargetMode") == "External":
        return target
    return posixpath.normpath(posixpath.join("word", target))


def open_docx(path: Path) -> DocContext:
    if path.suffix.lower() != ".docx":
        raise ValueError("交付文件不是 .docx 格式")
    if not path.exists():
        raise ValueError(f"文件不存在：{path}")
    if path.name.startswith("~$"):
        raise ValueError("疑似 Word 临时锁文件，不是可评估的正式交付文件")
    try:
        z = zipfile.ZipFile(path)
        bad = z.testzip()
        if bad:
            raise ValueError(f"ZIP 包中存在损坏条目：{bad}")
        names = z.namelist()
        required = ["[Content_Types].xml", "word/document.xml", "word/_rels/document.xml.rels"]
        missing = [p for p in required if p not in names]
        if missing:
            raise ValueError("缺少 docx 必需部件：" + ", ".join(missing))
        document_root = parse_xml(z.read("word/document.xml"), "word/document.xml")
        rels = parse_rels(z, "word/_rels/document.xml.rels")
        app_root = read_optional_xml(z, "docProps/app.xml")

        header_roots: List[ET.Element] = []
        footer_roots: List[ET.Element] = []
        for rel in rels.values():
            typ = rel.get("Type", "")
            target = posixpath.normpath(posixpath.join("word", rel.get("Target", "")))
            if target in names and typ.endswith("/header"):
                root = read_optional_xml(z, target)
                if root is not None:
                    header_roots.append(root)
            if target in names and typ.endswith("/footer"):
                root = read_optional_xml(z, target)
                if root is not None:
                    footer_roots.append(root)

        return DocContext(path, z, names, document_root, rels, app_root, header_roots, footer_roots)
    except zipfile.BadZipFile as exc:
        raise ValueError("文件不是可正常打开的 docx/zip 包") from exc


def get_pages_count(ctx: DocContext) -> Optional[int]:
    if ctx.app_root is None:
        return None
    pages = ctx.app_root.find("ep:Pages", NS)
    if pages is not None and pages.text and pages.text.strip().isdigit():
        return int(pages.text.strip())
    return None


def get_tables(ctx: DocContext) -> List[ET.Element]:
    return ctx.document_root.findall(".//w:tbl", NS)


def get_rows(table: ET.Element) -> List[ET.Element]:
    return table.findall("./w:tr", NS)


def get_cells(row: ET.Element) -> List[ET.Element]:
    return row.findall("./w:tc", NS)


def find_row_by_label(ctx: DocContext, label: str) -> Tuple[Optional[ET.Element], Optional[ET.Element], int, int]:
    for ti, tbl in enumerate(get_tables(ctx), 1):
        for ri, row in enumerate(get_rows(tbl), 1):
            cells = get_cells(row)
            if cells and label in normalize_text(text_of(cells[0])):
                return tbl, row, ti, ri
    # Fallback: label can appear outside the first cell in malformed documents.
    for ti, tbl in enumerate(get_tables(ctx), 1):
        for ri, row in enumerate(get_rows(tbl), 1):
            if label in normalize_text(text_of(row)):
                return tbl, row, ti, ri
    return None, None, -1, -1


def compute_page_index_of(ctx: DocContext, target: ET.Element) -> Tuple[Optional[int], str]:
    """基于 OOXML 分页信号估算 target 元素所在的 1-based 页码。

    OOXML 本身不记录精确分页边界，Word 排版引擎在保存时会写入
    ``w:lastRenderedPageBreak`` 作为已渲染分页锚点；文档内也可能存在
    用户显式插入的 ``w:br w:type="page"`` 或段落 ``w:pageBreakBefore``。
    这里按文档流顺序统计 target 之前的分页信号次数，返回 breaks+1；
    无任何分页信号时返回 (None, ...) 表示无法可靠推断。
    """
    body = ctx.document_root.find("w:body", NS)
    if body is None:
        return None, "文档缺少 w:body 节点"

    # 优先使用 Word 保存时写入的 lastRenderedPageBreak；否则回退到显式分页符和 pageBreakBefore
    has_rendered = ctx.document_root.find(".//w:lastRenderedPageBreak", NS) is not None
    source_label = "w:lastRenderedPageBreak" if has_rendered else "w:br(page)/pageBreakBefore"

    breaks = 0
    seen_target = False
    for node in body.iter():
        if node is target:
            seen_target = True
            break
        local = node.tag.split("}")[-1]
        if has_rendered:
            if local == "lastRenderedPageBreak":
                breaks += 1
        else:
            if local == "br" and node.attrib.get(W + "type") == "page":
                breaks += 1
            elif local == "pageBreakBefore":
                # w:pageBreakBefore 位于 w:pPr 下，表示所在段落起始处开始新页
                breaks += 1

    if not seen_target:
        return None, "未在文档流中定位到目标元素"
    if breaks == 0 and not has_rendered:
        return None, "文档未包含 lastRenderedPageBreak/显式分页符，无法可靠推断页码"
    return breaks + 1, f"依据{source_label}计数={breaks}，推断位于第{breaks + 1}页"


def cell_width_twips(cell: ET.Element) -> Optional[int]:
    tcw = cell.find("./w:tcPr/w:tcW", NS)
    if tcw is not None:
        val = tcw.attrib.get(W + "w")
        if val and val.lstrip("-").isdigit():
            return int(val)
    return None


def table_width_twips(table: ET.Element) -> Optional[int]:
    grid = table.findall("./w:tblGrid/w:gridCol", NS)
    vals = []
    for col in grid:
        w = col.attrib.get(W + "w")
        if w and w.isdigit():
            vals.append(int(w))
    if vals:
        return sum(vals)
    widths = []
    rows = get_rows(table)
    if rows:
        for c in get_cells(rows[0]):
            cw = cell_width_twips(c)
            if cw:
                widths.append(cw)
    return sum(widths) if widths else None


def page_text_capacity_twips(ctx: DocContext) -> Optional[int]:
    sect = ctx.document_root.find(".//w:sectPr", NS)
    if sect is None:
        return None
    pg_sz = sect.find("w:pgSz", NS)
    pg_mar = sect.find("w:pgMar", NS)
    if pg_sz is None or pg_mar is None:
        return None
    try:
        width = int(pg_sz.attrib.get(W + "w", "0"))
        left = int(pg_mar.attrib.get(W + "left", "0"))
        right = int(pg_mar.attrib.get(W + "right", "0"))
        return width - left - right
    except ValueError:
        return None


def find_ppt_elements(cell: ET.Element) -> Tuple[List[ET.Element], List[ET.Element], List[ET.Element]]:
    objects = cell.findall(".//w:object", NS)
    hyperlinks = cell.findall(".//w:hyperlink", NS)
    drawings = cell.findall(".//w:drawing", NS)
    return objects, hyperlinks, drawings


def ppt_related_text_in_cell(ctx: DocContext, cell: ET.Element) -> str:
    chunks = [all_text_like(cell)]
    for hyp in cell.findall(".//w:hyperlink", NS):
        rid = hyp.attrib.get(R + "id")
        if rid:
            chunks.append(relationship_target(ctx, rid) or "")
    for ole in cell.findall(".//o:OLEObject", NS):
        rid = ole.attrib.get(R + "id")
        if rid:
            chunks.append(relationship_target(ctx, rid) or "")
    return "".join(chunks)


def is_powerpoint_ole(obj: ET.Element) -> bool:
    ole = obj.find(".//o:OLEObject", NS)
    if ole is None:
        return False
    prog = (ole.attrib.get("ProgID") or "").lower()
    typ = (ole.attrib.get("Type") or "").lower()
    return ("powerpoint" in prog or "presentation" in prog or "show" in prog or "slide" in prog) and typ in {"embed", "link", ""}


def is_ppt_hyperlink(ctx: DocContext, hyp: ET.Element) -> bool:
    rid = hyp.attrib.get(R + "id")
    target = relationship_target(ctx, rid) if rid else ""
    display = text_of(hyp)
    combined = normalize_text((target or "") + display)
    return ".pptx" in combined.lower() or "ppt" in combined.lower() or "项目介绍材料" in combined


def collect_related_ppt_parts(ctx: DocContext, cell: Optional[ET.Element] = None) -> List[Tuple[str, bytes, str]]:
    """Return [(source, bytes, kind)] for embedded or external ppt/pptx files that can be opened."""
    candidates: List[Tuple[str, Optional[bytes], str]] = []

    # Embedded OLE objects referenced by the target cell are the strongest evidence.
    roots = [cell] if cell is not None else [ctx.document_root]
    for root in roots:
        if root is None:
            continue
        for ole in root.findall(".//o:OLEObject", NS):
            rid = ole.attrib.get(R + "id")
            if not rid:
                continue
            target = relationship_target(ctx, rid)
            if target and target in ctx.names:
                try:
                    candidates.append((target, ctx.z.read(target), "embedded"))
                except KeyError:
                    pass

    # Any PowerPoint-like embedded package in word/embeddings is also acceptable.
    for name in ctx.names:
        lower = name.lower()
        if name.startswith("word/embeddings/") and (lower.endswith(".pptx") or lower.endswith(".bin")):
            if all(name != c[0] for c in candidates):
                try:
                    candidates.append((name, ctx.z.read(name), "embedded"))
                except KeyError:
                    pass

    # Hyperlink target may point to same-directory file.
    if cell is not None:
        for hyp in cell.findall(".//w:hyperlink", NS):
            rid = hyp.attrib.get(R + "id")
            target = relationship_target(ctx, rid) if rid else None
            if not target or not re.search(r"\.pptx$", target, re.I):
                continue
            target_path = Path(target)
            possible = []
            if not target_path.is_absolute():
                possible.append(ctx.path.parent / target_path)
                possible.append(ctx.path.parent / target_path.name)
            else:
                possible.append(target_path)
            for p in possible:
                if p.exists() and p.is_file():
                    candidates.append((str(p), p.read_bytes(), "external"))
                    break

    return [(src, data or b"", kind) for src, data, kind in candidates]


def pptx_text(data: bytes) -> Optional[str]:
    try:
        with zipfile.ZipFile(io.BytesIO(data)) as z:
            names = z.namelist()
            slide_names = sorted(
                [n for n in names if n.startswith("ppt/slides/slide") and n.endswith(".xml")],
                key=lambda s: [int(x) if x.isdigit() else x for x in re.split(r"(\d+)", s)],
            )
            if not slide_names:
                return None
            chunks: List[str] = []
            # First few slides are enough for title/content identity checks.
            for name in slide_names[:3]:
                root = parse_xml(z.read(name), name)
                chunks.extend(t.text or "" for t in root.iter("{%s}t" % NS["a"]))
            return "".join(chunks)
    except Exception:
        return None


def pptx_first_slide_text(data: bytes) -> Optional[str]:
    try:
        with zipfile.ZipFile(io.BytesIO(data)) as z:
            slide_names = sorted(
                [n for n in z.namelist() if n.startswith("ppt/slides/slide") and n.endswith(".xml")],
                key=lambda s: [int(x) if x.isdigit() else x for x in re.split(r"(\d+)", s)],
            )
            if not slide_names:
                return None
            root = parse_xml(z.read(slide_names[0]), slide_names[0])
            return "".join(t.text or "" for t in root.iter("{%s}t" % NS["a"]))
    except Exception:
        return None


def ppt_click_target_and_title_page_ok(ctx: DocContext, cell: Optional[ET.Element]) -> Tuple[bool, str]:
    ppt_parts = collect_related_ppt_parts(ctx, cell)
    if not ppt_parts:
        return False, "未找到单击/双击后可打开的嵌入或同目录 PPT/PPTX 文件"
    required_filename = normalize_text(CLICK_TARGET_PPT_FILENAME).lower()
    required_text = [normalize_text(t) for t in CLICK_TITLE_PAGE_REQUIRED_TEXT]
    details = []
    for source, data, kind in ppt_parts:
        source_name = normalize_text(Path(source).name).lower()
        opens_expected_file = source_name == required_filename
        opens_ppt_interface = kind == "embedded"
        first_slide = pptx_first_slide_text(data)
        if first_slide is None:
            details.append(f"{kind} {source} 不可解析首页/标题页")
            continue
        title_page_text = normalize_text(first_slide)
        content_ok = all(t in title_page_text for t in required_text)
        open_target_ok = opens_expected_file or opens_ppt_interface
        details.append(
            f"{kind} {source}：打开指定文件={opens_expected_file}，"
            f"跳转PPT播放/编辑界面={opens_ppt_interface}，首页/标题页内容匹配={content_ok}"
        )
        if open_target_ok and content_ok:
            return True, "；".join(details)
    return False, "；".join(details)


def ppt_contains_expected_content(ctx: DocContext, cell: Optional[ET.Element]) -> Tuple[bool, str]:
    ppt_parts = collect_related_ppt_parts(ctx, cell)
    if not ppt_parts:
        return False, "未找到可打开的嵌入或同目录 PPT/PPTX 文件"
    for source, data, kind in ppt_parts:
        txt = pptx_text(data)
        if txt is None:
            continue
        norm = normalize_text(txt)
        hit_keywords = [kw for kw in PPT_CONTENT_KEYWORDS if normalize_text(kw) in norm]
        if len(hit_keywords) >= 2:
            return True, f"{kind} PPT 可解析：{source}；首页/前几页命中 {len(hit_keywords)} 个项目关键词"
    return False, "找到 PPT 文件，但首页/前几页未命中足够的项目介绍关键词"


def cell_contains_only_picture(cell: ET.Element) -> bool:
    txt = normalize_text(text_of(cell))
    objects, hyperlinks, drawings = find_ppt_elements(cell)
    picts = cell.findall(".//w:pict", NS)
    # OLE object with an icon often uses w:pict; that is not a plain screenshot.
    has_ole = bool(cell.findall(".//o:OLEObject", NS))
    return bool(drawings or picts) and not txt and not objects and not hyperlinks and not has_ole


def has_editable_body_and_tables(ctx: DocContext) -> Tuple[bool, str]:
    body_text = normalize_text(text_of(ctx.document_root))
    tables = get_tables(ctx)
    if len(body_text) < 200:
        return False, "正文可提取文字过少，疑似整页图片或不可编辑对象"
    if not tables:
        return False, "未检测到可编辑 w:tbl 表格结构"
    table_text_len = sum(len(normalize_text(text_of(t))) for t in tables)
    if table_text_len < 200:
        return False, "表格内可提取文字过少，疑似表格被转成图片"
    media = [n for n in ctx.names if n.startswith("word/media/")]
    drawings = ctx.document_root.findall(".//w:drawing", NS)
    objects = ctx.document_root.findall(".//w:object", NS)
    # 若有大量图片且缺少正文/表格结构，才判为不可编辑；单个 PPT 图标是允许的。
    if len(media) >= 4 and len(body_text) < 1000 and not objects:
        return False, "文档包含多张图片且缺少可编辑内容，疑似整页/整表截图"
    return True, f"正文可提取 {len(body_text)} 字；检测到 {len(tables)} 个可编辑表格；图片 {len(media)} 个、OLE 对象 {len(objects)} 个"


def dimension1(ctx: DocContext) -> Tuple[bool, List[str]]:
    details: List[str] = ["DOCX 压缩包和核心 XML 可正常打开解析"]

    _, row, _, _ = find_row_by_label(ctx, TARGET_LABEL)
    target_cell = get_cells(row)[1] if row is not None and len(get_cells(row)) >= 2 else None
    if target_cell is None:
        details.append("未在表格中找到项目介绍材料右侧申报内容单元格")
        return False, details

    objects, hyperlinks, drawings = find_ppt_elements(target_cell)
    has_ppt_ole = any(is_powerpoint_ole(obj) for obj in objects)
    has_ppt_link = any(is_ppt_hyperlink(ctx, hyp) for hyp in hyperlinks)
    if not (has_ppt_ole or has_ppt_link):
        details.append("目标单元格未检测到 PowerPoint OLE 对象或 PPT 超链接")
        return False, details
    if cell_contains_only_picture(target_cell):
        details.append("目标单元格疑似只放置了静态图片，未检测到可点击对象")
        return False, details
    details.append(f"目标单元格检测到 PowerPoint OLE={has_ppt_ole}，PPT 超链接={has_ppt_link}，不是纯截图")

    return True, details


def _display_name_in_cell(ctx: DocContext, cell: ET.Element) -> str:
    # 收集单元格内插入对象的展示层文字：显示文件名、OLE 图标标题、链接显示文字
    # 对应细则中"插入对象显示的文件名、图标标题或链接文字"
    chunks: List[str] = [text_of(cell)]
    # OLE 图标：v:shape title、o:OLEObject ProgID 附近的 w:t 文本、shape name 属性
    for obj in cell.findall(".//w:object", NS):
        shape = obj.find(".//v:shape", NS)
        if shape is not None:
            chunks.append(shape.attrib.get("title") or "")
            chunks.append(shape.attrib.get("alt") or "")
        for node in obj.iter():
            local = node.tag.split("}")[-1].lower()
            if local in {"title", "name"} and node.text:
                chunks.append(node.text)
        # OLE 对象旁的可见文字（w:t）即为图标标题
        for t in obj.findall(".//w:t", NS):
            if t.text:
                chunks.append(t.text)
    # 超链接：显示文字
    for hyp in cell.findall(".//w:hyperlink", NS):
        chunks.append(text_of(hyp))
    return "".join(chunks)


def check_inserted_object_in_target_cell(ctx: DocContext) -> CheckResult:
    # +5: 第4页"项目介绍材料"右侧申报内容单元格插入 PPT 对象/图标/链接
    # 且展示的文件名/图标标题/链接文字包含"海潮艺境MR共创计划"或"项目介绍材料PPT"
    tbl, row, ti, ri = find_row_by_label(ctx, TARGET_LABEL)
    if row is None:
        return CheckResult("+5 对象位置与可识别名称", 5, False, "未找到项目介绍材料行")
    cells = get_cells(row)
    if len(cells) < 2:
        return CheckResult("+5 对象位置与可识别名称", 5, False, "项目介绍材料行缺少右侧申报内容单元格")

    # 先做第4页定位：优先用 Word 保存的 lastRenderedPageBreak，其次回退显式分页符/pageBreakBefore
    page_idx, page_ev = compute_page_index_of(ctx, row)
    # 无法推断时容许通过（避免因无渲染分页信号误判），但在证据里明确标记
    on_page4 = page_idx == 4 if page_idx is not None else True
    if page_idx is not None and not on_page4:
        return CheckResult(
            "+5 对象位置与可识别名称", 5, False,
            f"项目介绍材料行不在第4页（{page_ev}）"
        )

    label_cell, right_cell = cells[0], cells[1]
    objects, hyperlinks, _ = find_ppt_elements(right_cell)
    # 细则要求：插入形式为 PPT 对象、PPT 图标或 PPT 链接对象
    has_ppt_obj = any(is_powerpoint_ole(o) for o in objects)
    has_ppt_link = any(is_ppt_hyperlink(ctx, h) for h in hyperlinks)
    # 细则要求：展示的文件名/图标标题/链接文字包含指定关键词
    display_text = normalize_text(_display_name_in_cell(ctx, right_cell))
    name_hit = any(normalize_text(p) in display_text for p in EXPECTED_TITLE_PATTERNS)
    wrong_side = any(is_powerpoint_ole(o) for o in label_cell.findall(".//w:object", NS)) or any(is_ppt_hyperlink(ctx, h) for h in label_cell.findall(".//w:hyperlink", NS))
    hit = (has_ppt_obj or has_ppt_link) and name_hit and not wrong_side
    page_note = f"页码={page_idx}" if page_idx is not None else "页码=无法从OOXML可靠推断"
    detail = (f"表{ti}第{ri}行右侧单元格（{page_note}；{page_ev}）："
              f"PowerPoint对象={has_ppt_obj}，PPT链接={has_ppt_link}，"
              f"展示名称包含关键词={name_hit}（展示文字：{display_text[:60] or '空'}）")
    return CheckResult("+5 对象位置与可识别名称", 5, hit, detail)


def check_click_opens_expected_ppt(ctx: DocContext) -> CheckResult:
    _, row, _, _ = find_row_by_label(ctx, TARGET_LABEL)
    cell = get_cells(row)[1] if row is not None and len(get_cells(row)) >= 2 else None
    if cell is None:
        return CheckResult("+5 点击打开正确 PPT", 5, False, "未找到目标单元格")
    has_clickable = bool(cell.findall(".//o:OLEObject", NS)) or any(is_ppt_hyperlink(ctx, h) for h in cell.findall(".//w:hyperlink", NS))
    target_and_content_ok, detail = ppt_click_target_and_title_page_ok(ctx, cell)
    return CheckResult(
        "+5 点击打开正确 PPT",
        5,
        has_clickable and target_and_content_ok,
        f"可单击/双击对象={has_clickable}；{detail}",
    )


def check_link_editability(ctx: DocContext) -> CheckResult:
    _, row, _, _ = find_row_by_label(ctx, TARGET_LABEL)
    cell = get_cells(row)[1] if row is not None and len(get_cells(row)) >= 2 else None
    if cell is None:
        return CheckResult("+1 链接/对象可编辑", 1, False, "未找到目标单元格")
    objects, hyperlinks, _ = find_ppt_elements(cell)
    ole_objects = [o for o in objects if is_powerpoint_ole(o)]
    # 细则要求：能被单独选中、移动、缩放或重新设置链接 → OLE 嵌入/链接对象 或 超链接对象
    is_selectable_editable = bool(ole_objects or hyperlinks)
    # 细则要求：排除"只显示PPT首页截图但无法点击的图片"
    is_only_screenshot = cell_contains_only_picture(cell)
    hit = is_selectable_editable and not is_only_screenshot
    detail = (f"OLE对象（可选中/移动/缩放）={len(ole_objects)} 个，"
              f"超链接对象（可重新设置链接）={len(hyperlinks)} 个，"
              f"仅为不可点击截图={is_only_screenshot}")
    return CheckResult("+1 链接/对象可编辑", 1, hit, detail)


def shape_size_points(style: str) -> Tuple[Optional[float], Optional[float]]:
    """Extract width/height in pt from a VML style string."""
    def one(prop: str) -> Optional[float]:
        m = re.search(rf"{prop}\s*:\s*([0-9.]+)\s*pt", style or "", re.I)
        return float(m.group(1)) if m else None
    return one("width"), one("height")


def check_display_style(ctx: DocContext) -> CheckResult:
    _, row, _, _ = find_row_by_label(ctx, TARGET_LABEL)
    cell = get_cells(row)[1] if row is not None and len(get_cells(row)) >= 2 else None
    if cell is None:
        return CheckResult("+3 显示样式", 3, False, "未找到目标单元格")
    _, hyperlinks, _ = find_ppt_elements(cell)

    # ── 图标路线：PowerPoint图标/文件图标/嵌入对象图标 ──────────────────
    has_icon = False
    icon_centered = False
    icon_in_bounds = False   # 不超出单元格边界：图标宽 ≤ 单元格宽
    sizes = []
    right_cell_width = cell_width_twips(cell)
    for paragraph in cell.findall(".//w:p", NS):
        jc = paragraph.find("./w:pPr/w:jc", NS)
        paragraph_centered = jc is not None and jc.attrib.get(W + "val") == "center"
        for obj in paragraph.findall(".//w:object", NS):
            if not is_powerpoint_ole(obj):
                continue
            has_icon = True
            icon_centered = icon_centered or paragraph_centered
            shape = obj.find(".//v:shape", NS)
            if shape is not None:
                w_pt, h_pt = shape_size_points(shape.attrib.get("style", ""))
                if w_pt and h_pt:
                    sizes.append((w_pt, h_pt))
                    # 不超出单元格边界：图标宽（twips）≤ 单元格宽
                    obj_twips = int(w_pt * 20)
                    in_bounds = (right_cell_width is None) or (obj_twips <= right_cell_width)
                    icon_in_bounds = icon_in_bounds or in_bounds
    icon_ok = has_icon and icon_centered and icon_in_bounds

    # ── 链接文字路线：带下划线的超链接文本 ──────────────────────────────
    has_ppt_link = False
    hyperlink_ok = False
    hyp_details = []
    for hyp in hyperlinks:
        if not is_ppt_hyperlink(ctx, hyp):
            continue
        has_ppt_link = True
        txt = text_of(hyp).strip()
        # 带下划线：显式 w:u 或继承 Hyperlink 字符样式均算
        underline = (hyp.find(".//w:u", NS) is not None
                     or any((r.attrib.get(W + "val") or "").lower() in {"hyperlink", "a0"}
                            for r in hyp.findall(".//w:rStyle", NS)))
        # 字号小五或近似小五：小五 = 9pt，w:sz 单位为半磅；放宽到 8-12pt；未显式设置也允许
        sizes_half_pt = []
        for sz in hyp.findall(".//w:sz", NS):
            val = sz.attrib.get(W + "val")
            if val and val.isdigit():
                sizes_half_pt.append(int(val))
        font_size_ok = not sizes_half_pt or all(16 <= v <= 24 for v in sizes_half_pt)
        # 链接文本完整显示：文本非空且长度合理（不截断）
        text_complete = 2 <= len(txt) <= 80
        # 不超出单元格边界：无法精确测量文字宽，以文本长度≤单元格允许字符数近似判定；
        # 宽度未知时不扣分
        hyp_details.append(f"带下划线={underline}，字号半磅={sizes_half_pt if sizes_half_pt else '未设置'}，文本完整={text_complete}")
        hyperlink_ok = hyperlink_ok or (underline and font_size_ok and text_complete)

    # ── 不遮挡栏目文字：PPT对象/链接只在右侧申报单元格，不在左侧标签单元格 ─
    label_cell = get_cells(row)[0] if row is not None and get_cells(row) else None
    label_text = normalize_text(text_of(label_cell)) if label_cell is not None else ""
    label_cell_has_obj = label_cell is not None and (
        bool(label_cell.findall(".//o:OLEObject", NS))
        or bool(label_cell.findall(".//w:hyperlink", NS))
    )
    no_overlap = TARGET_LABEL in label_text and not label_cell_has_obj

    # 只出现一种显示形式时，出现的形式符合要求即可；图标和链接都出现时，两种都必须符合要求。
    if has_icon and has_ppt_link:
        style_ok = icon_ok and hyperlink_ok
    elif has_icon:
        style_ok = icon_ok
    elif has_ppt_link:
        style_ok = hyperlink_ok
    else:
        style_ok = False
    hit = style_ok and no_overlap
    detail = (
        f"显示为图标={has_icon}，图标居中={icon_centered}，"
        f"图标不超单元格边界={icon_in_bounds}，图标路线通过={icon_ok}；"
        f"存在PPT超链接={has_ppt_link}，链接路线通过={hyperlink_ok}（{'; '.join(hyp_details) or '无PPT超链接'}）；"
        f"显示形式组合通过={style_ok}，未遮挡栏目文字={no_overlap}"
    )
    return CheckResult("+3 显示样式", 3, hit, detail)


def check_embed_or_companion(ctx: DocContext) -> CheckResult:
    _, row, _, _ = find_row_by_label(ctx, TARGET_LABEL)
    cell = get_cells(row)[1] if row is not None and len(get_cells(row)) >= 2 else None
    ppt_parts = collect_related_ppt_parts(ctx, cell)

    # ── 嵌入方式：Word 文件内部包含 PPT 对象，移动 Word 后仍可打开 ───────
    # kind=="embedded" 即 OLE 内嵌，数据存于 word/embeddings/，与 Word 文件一体，
    # 移动 docx 后无需外部文件即可打开。
    embedded_ok = any(kind == "embedded" and pptx_text(data) is not None for _, data, kind in ppt_parts)

    # ── 外部链接方式：交付包中同时含 CLICK_TARGET_PPT_FILENAME，
    #    且链接路径为相对路径或同目录可用路径 ───────────────────────────
    required_filename = normalize_text(CLICK_TARGET_PPT_FILENAME)
    external_ok = False
    external_details = []
    if cell is not None:
        for hyp in cell.findall(".//w:hyperlink", NS):
            rid = hyp.attrib.get(R + "id")
            target = relationship_target(ctx, rid) if rid else None
            if not target or not re.search(r"\.pptx$", target, re.I):
                continue
            # 链接路径为相对路径
            is_relative = not Path(target).is_absolute()
            # 交付包同目录含指定文件
            companion = ctx.path.parent / Path(target).name
            companion_is_required = normalize_text(companion.name) == required_filename
            exists_same_dir = companion.exists() and companion.is_file()
            external_details.append(
                f"链接={target}，相对路径={is_relative}，"
                f"文件名匹配={companion_is_required}，同目录存在={exists_same_dir}"
            )
            # 同时满足：相对/同目录路径 + 文件名为指定名称 + 物理存在
            external_ok = external_ok or (is_relative and companion_is_required and exists_same_dir)

    hit = embedded_ok or external_ok
    detail = (
        f"嵌入对象（移动后可用）={embedded_ok}"
        + ("；" + "；".join(external_details) if external_details else "；未检测到外部链接")
    )
    return CheckResult("+5 嵌入或随附方式", 5, hit, detail)


def header_roots_by_type(ctx: DocContext) -> Tuple[Dict[str, ET.Element], bool]:
    sect = ctx.document_root.find(".//w:sectPr", NS)
    if sect is None:
        return {}, False
    has_title_page = sect.find("w:titlePg", NS) is not None
    headers: Dict[str, ET.Element] = {}
    for ref in sect.findall("w:headerReference", NS):
        typ = ref.attrib.get(W + "type", "default")
        rid = ref.attrib.get(R + "id")
        target = relationship_target(ctx, rid) if rid else None
        if target and target in ctx.names:
            root = read_optional_xml(ctx.z, target)
            if root is not None:
                headers[typ] = root
    return headers, has_title_page


def header_has_text(root: ET.Element) -> bool:
    return len(normalize_text(text_of(root))) >= 1


def header_for_page(headers: Dict[str, ET.Element], has_title_page: bool, page_num: int) -> Optional[ET.Element]:
    if page_num == 1 and has_title_page:
        return headers.get("first")
    if page_num % 2 == 0 and "even" in headers:
        return headers.get("even")
    return headers.get("default")


def check_headers(ctx: DocContext) -> CheckResult:
    pages = get_pages_count(ctx)
    headers, has_title_page = header_roots_by_type(ctx)
    header_text_present = {}
    for page_num in range(1, 5):
        header = header_for_page(headers, has_title_page, page_num)
        header_text_present[page_num] = header_has_text(header) if header is not None else False
    missing_pages = [p for p, ok in header_text_present.items() if not ok]
    # 细则：第1页至第4页任意一页页眉文本缺失即扣分；页数属性不足4页也视为第1-4页不完整。
    hit = bool(missing_pages) or bool(pages is not None and pages < 4)
    detail = (
        f"页数属性={pages}，首页独立页眉={has_title_page}，"
        f"页眉类型={list(headers.keys()) or '无'}，第1-4页页眉文本检测={header_text_present}，"
        f"缺失页={missing_pages or '无'}"
    )
    return CheckResult("-1 第1-4页页眉文本缺失", -1, hit, detail)


def row_has_all_borders(row: ET.Element) -> bool:
    cells = get_cells(row)
    if not cells:
        return False
    required = ["top", "left", "bottom", "right"]
    for cell in cells:
        borders = cell.find("./w:tcPr/w:tcBorders", NS)
        if borders is None:
            return False
        for side in required:
            el = borders.find(f"w:{side}", NS)
            if el is None:
                return False
            val = el.attrib.get(W + "val", "")
            if val in {"nil", "none", ""}:
                return False
    return True


def check_last_row_border(ctx: DocContext) -> CheckResult:
    tables = get_tables(ctx)
    if not tables:
        return CheckResult("-1 第5页最后一行边框/结构断裂", -1, True, "未检测到表格")

    # 细则：第5页项目申报表 → 先按 OOXML 分页信号定位第5页上的申报表；
    # 无法定位（文档未包含 lastRenderedPageBreak/显式分页符）时才回退到旧启发式
    page5_tables: List[Tuple[int, ET.Element, int, str]] = []
    for idx, t in enumerate(tables, 1):
        page_idx, page_ev = compute_page_index_of(ctx, t)
        if page_idx == 5:
            page5_tables.append((idx, t, len(get_rows(t)), page_ev))

    main_table: Optional[ET.Element] = None
    picked_ev = ""
    if page5_tables:
        # 第5页可能有多张表：选行数最多的（申报表通常远大于其他附属表）
        page5_tables.sort(key=lambda x: x[2], reverse=True)
        idx, main_table, row_count, picked_ev = page5_tables[0]
        picked_ev = f"命中第5页表(第{idx}张，行数={row_count})；{picked_ev}"
    else:
        # 无法用分页信号定位第5页时的回退：优先含"项目介绍材料"行的申报表；再退到行数最多的表
        # 并在证据中标明当前是回退路径，避免误判静默
        fallback_ev_parts = ["未能通过 OOXML 分页信号定位第5页表格，回退启发式"]
        tbl, _, ti_lbl, _ = find_row_by_label(ctx, TARGET_LABEL)
        if tbl is not None:
            main_table = tbl
            fallback_ev_parts.append(f"选用含'{TARGET_LABEL}'行的表(第{ti_lbl}张)")
        else:
            main_table = max(tables, key=lambda t: len(get_rows(t)))
            fallback_ev_parts.append("选用行数最多的表")
        picked_ev = "；".join(fallback_ev_parts)

    rows = get_rows(main_table)

    # 细则：表格结构断裂 → 表格完全没有行，或最后一行单元格为空
    structure_ok = bool(rows) and bool(get_cells(rows[-1]))

    # 细则：最后一行边框缺失 → 最后一行每个单元格四边框均需存在且非nil/none
    last = rows[-1] if rows else None
    borders_ok = row_has_all_borders(last) if last is not None else False

    hit = not structure_ok or not borders_ok
    detail = (
        f"{picked_ev}；申报表行数={len(rows)}，最后一行单元格数={len(get_cells(last)) if last is not None else 0}，"
        f"表格结构完整={structure_ok}，最后一行四边框完整={borders_ok}"
    )
    return CheckResult("-1 第5页最后一行边框/结构断裂", -1, hit, detail)



def run_dimension2(ctx: DocContext) -> List[CheckResult]:
    return [
        check_inserted_object_in_target_cell(ctx),
        check_click_opens_expected_ppt(ctx),
        check_link_editability(ctx),
        check_display_style(ctx),
        check_embed_or_companion(ctx),
        check_headers(ctx),
        check_last_row_border(ctx),
    ]


def _evaluate_docx(path: Path) -> Evaluation:
    ev = Evaluation()
    try:
        ctx = open_docx(path)
    except ValueError as exc:
        ev.dimension1_passed = False
        ev.dimension1_details.append(str(exc))
        ev.total = 0
        return ev

    try:
        passed, details = dimension1(ctx)
        ev.dimension1_passed = passed
        ev.dimension1_details = details
        if not passed:
            ev.total = 0
            return ev
        ev.checks = run_dimension2(ctx)
        ev.total = sum(c.score for c in ev.checks if c.hit)
        return ev
    finally:
        ctx.z.close()


def _rule_text(code: str) -> str:
    # 去掉 "+5 " / "-1 " 之类的分数前缀，只保留规则文字。
    return re.sub(r"^[+-]?\d+\s*[:：]?\s*", "", code).strip()


def _pick_docx_in_dir(dir_path: Path) -> Optional[Path]:
    docs = [p for p in dir_path.glob("*.docx") if not p.name.startswith("~$")]
    if not docs:
        return None
    if len(docs) == 1:
        return docs[0]
    # 优先选文件名含"PPT"的（本用例的已插入 PPT 版本）。
    for p in docs:
        if "PPT" in p.name.upper():
            return p
    return docs[0]


def evaluate(dir_path: str) -> dict:
    """按统一约定评估：dir_path 为脚本所在目录，脚本自己在其中定位 .docx。"""
    script_id = "059"
    result: dict = {
        "id": script_id,
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
        base = Path(dir_path)
        if not base.exists() or not base.is_dir():
            result["status"] = "error"
            result["error"] = f"目录不存在或不是目录：{dir_path}"
            return result

        docx_path = _pick_docx_in_dir(base)
        if docx_path is None:
            result["status"] = "error"
            result["error"] = f"目录中未找到可评估的 .docx 文件：{dir_path}"
            return result

        result["file_name"] = docx_path.name

        ev = _evaluate_docx(docx_path)
        result["dim1_pass"] = ev.dimension1_passed
        if not ev.dimension1_passed:
            result["dim1_reason"] = "；".join(ev.dimension1_details)

        dim2_items = [
            {
                "rule": _rule_text(c.code),
                "max_delta": c.score,
                "delta": c.score if c.hit else 0,
                "hit": c.hit,
                "detail": "",
            }
            for c in ev.checks
        ]
        result["dim2_items"] = dim2_items
        result["max_score"] = sum(c.score for c in ev.checks if c.score > 0)
        result["total_score"] = ev.total if ev.dimension1_passed else 0
        return result
    except Exception as exc:  # 兜底：脚本自身崩溃走 status=error
        result["status"] = "error"
        result["error"] = f"{type(exc).__name__}: {exc}"
        return result


if __name__ == "__main__":
    # 仅用于本地调试；主结果只走 return，这里以 JSON 形式打印便于人工核对。
    _dir = sys.argv[1] if len(sys.argv) > 1 else str(Path(__file__).resolve().parent)
    print(json.dumps(evaluate(_dir), ensure_ascii=False, indent=2))
