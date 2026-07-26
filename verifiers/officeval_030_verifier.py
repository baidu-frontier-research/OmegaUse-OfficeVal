#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""自动评估毕业设计成果说明书 docx 格式。

评估逻辑：
1. 先检查维度1“可用与可修改性”；不满足则最终得分为 0，并跳过维度2。
2. 维度1通过后，逐条检查维度2得分/扣分点；命中即累计该点分数。

本脚本只使用 Python 标准库，直接解析 .docx 内部的 WordprocessingML。

说明：按接口统一约定，本模块不修改全局 sys.stdout，也不 print 主结果。
主结果统一由 evaluate(dir_path) -> dict 返回；本文件仅在
__main__ 下作为本地调试入口，才把 dict 序列化为 JSON 打印。
"""

from __future__ import annotations

import json
import re
import sys
import zipfile
from dataclasses import dataclass
from pathlib import Path
from typing import Callable, Optional
from xml.etree import ElementTree as ET

SCRIPT_ID = "030"
# 维度二所有评分项 max_delta 之和；用于维度一未通过 / 脚本异常时兜底
MAX_SCORE_TOTAL = 51

NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
}

W = "{%s}" % NS["w"]
R = "{%s}" % NS["r"]

# Word 常用单位：twips、EMU、半磅。
TWIPS_PER_CM = 1440 / 2.54
EMU_PER_CM = 360000

FONT_PT = {
    "小一": 24.0,
    "二号": 22.0,
    "小二": 18.0,
    "三号": 16.0,
    "小三": 15.0,
    "四号": 14.0,
    "小四": 12.0,
    "五号": 10.5,
    "小五": 9.0,
}

CHAPTER_MAP = {
    "设计思路": "一、设计思路",
    "设计内容": "二、设计内容",
    "设计过程": "三、设计过程",
    "作品及特点": "四、作品及特点",
    "致谢": "五、致谢",
    "参考资料": "六、参考资料",
}

REQUIRED_TOC_ITEMS = ["设计思路", "设计内容", "设计过程", "作品及特点", "致谢", "参考资料"]
REQUIRED_CAPTIONS = [
    "图1 折叠式仓储搬运车外观结构",
    "图2 模块化底盘与副车架布局",
    "图3 机械转向机构",
    "图4 辅助动力与传动布置",
    "图5 集中式控制面板",
    "图6 悬架轮端结构",
    "图7 整机模块关系",
]


@dataclass
class ParagraphInfo:
    index: int
    element: ET.Element
    text: str
    style_id: str
    in_table: bool = False


@dataclass
class TableInfo:
    index: int
    element: ET.Element
    rows: list[list[str]]
    cell_paragraphs: list[ParagraphInfo]


@dataclass
class BodyItem:
    kind: str
    index: int
    text: str = ""


@dataclass
class CheckResult:
    passed: bool
    evidence: str


@dataclass
class ScoreItem:
    id: str
    score: int
    name: str
    check: Callable[["DocxReader"], CheckResult]


def qn(local: str) -> str:
    prefix, name = local.split(":", 1)
    return "{%s}%s" % (NS[prefix], name)


def attr(el: Optional[ET.Element], name: str, default: Optional[str] = None) -> Optional[str]:
    if el is None:
        return default
    if ":" in name:
        return el.get(qn(name), default)
    return el.get(name, default)


def to_int(value: Optional[str], default: Optional[int] = None) -> Optional[int]:
    try:
        return int(value) if value is not None else default
    except (TypeError, ValueError):
        return default


def twips_to_cm(value: Optional[str | int]) -> Optional[float]:
    n = to_int(str(value)) if value is not None else None
    if n is None:
        return None
    return n / TWIPS_PER_CM


def emu_to_cm(value: Optional[str | int]) -> Optional[float]:
    n = to_int(str(value)) if value is not None else None
    if n is None:
        return None
    return n / EMU_PER_CM


def half_points_to_pt(value: Optional[str | int]) -> Optional[float]:
    n = to_int(str(value)) if value is not None else None
    if n is None:
        return None
    return n / 2.0


def pt_to_twips(pt: float) -> float:
    return pt * 20


def cm_to_twips(cm: float) -> float:
    return cm * TWIPS_PER_CM


def normalize_text(text: str) -> str:
    """去掉所有空白，统一中文标点，便于内容比对。"""
    text = text.replace("：", ":").replace("（", "(").replace("）", ")")
    return re.sub(r"\s+", "", text)


def normalize_caption(text: str) -> str:
    return normalize_text(text).replace("图", "图").replace("表", "表")


def approx(value: Optional[float], expected: float, tolerance: float) -> bool:
    return value is not None and abs(value - expected) <= tolerance


def in_range(value: Optional[float], low: float, high: float) -> bool:
    return value is not None and low <= value <= high


def bool_element_is_true(el: Optional[ET.Element]) -> bool:
    if el is None:
        return False
    value = attr(el, "w:val")
    return value not in {"0", "false", "False", "off", "none"}


def paragraph_text(p: ET.Element) -> str:
    parts: list[str] = []
    for node in p.iter():
        if node.tag == qn("w:t"):
            parts.append(node.text or "")
        elif node.tag == qn("w:tab"):
            parts.append("\t")
        elif node.tag == qn("w:br"):
            parts.append("\n")
    return "".join(parts)


class DocxReader:
    def __init__(self, path: Path):
        self.path = Path(path)
        self.zip: Optional[zipfile.ZipFile] = None
        self.names: set[str] = set()
        self.document: Optional[ET.Element] = None
        self.styles_root: Optional[ET.Element] = None
        self.settings_root: Optional[ET.Element] = None
        self.numbering_root: Optional[ET.Element] = None
        self.rels_root: Optional[ET.Element] = None
        self.relationships: dict[str, dict[str, str]] = {}
        self.styles: dict[str, ET.Element] = {}
        self.body: Optional[ET.Element] = None
        self.paragraphs: list[ParagraphInfo] = []
        self.all_paragraphs: list[ParagraphInfo] = []
        self.tables: list[TableInfo] = []
        self.body_items: list[BodyItem] = []
        self.errors: list[str] = []

    def load(self) -> None:
        self.zip = zipfile.ZipFile(self.path)
        self.names = set(self.zip.namelist())
        self.document = self.read_xml("word/document.xml", required=True)
        self.styles_root = self.read_xml("word/styles.xml", required=False)
        self.settings_root = self.read_xml("word/settings.xml", required=False)
        self.numbering_root = self.read_xml("word/numbering.xml", required=False)
        self.rels_root = self.read_xml("word/_rels/document.xml.rels", required=False)
        self._parse_relationships()
        self._parse_styles()
        self._parse_body()

    def has(self, name: str) -> bool:
        return name in self.names

    def read_xml(self, name: str, required: bool = False) -> Optional[ET.Element]:
        if not self.zip or name not in self.names:
            if required:
                raise FileNotFoundError(f"缺少 {name}")
            return None
        try:
            return ET.fromstring(self.zip.read(name))
        except ET.ParseError as exc:
            if required:
                raise
            self.errors.append(f"{name} XML 解析失败：{exc}")
            return None

    def _parse_relationships(self) -> None:
        if self.rels_root is None:
            return
        for rel in list(self.rels_root):
            rid = rel.get("Id")
            if not rid:
                continue
            self.relationships[rid] = {
                "type": rel.get("Type", ""),
                "target": rel.get("Target", ""),
                "mode": rel.get("TargetMode", ""),
            }

    def _parse_styles(self) -> None:
        if self.styles_root is None:
            return
        for style in self.styles_root.findall("w:style", NS):
            style_id = attr(style, "w:styleId")
            if style_id:
                self.styles[style_id] = style

    def _parse_body(self) -> None:
        if self.document is None:
            return
        self.body = self.document.find("w:body", NS)
        if self.body is None:
            return
        p_index = 0
        t_index = 0
        all_index = 0
        for child in list(self.body):
            if child.tag == qn("w:p"):
                text = paragraph_text(child)
                info = ParagraphInfo(p_index, child, text, self.paragraph_style_id(child), False)
                self.paragraphs.append(info)
                self.all_paragraphs.append(info)
                self.body_items.append(BodyItem("p", p_index, text))
                p_index += 1
                all_index += 1
            elif child.tag == qn("w:tbl"):
                rows: list[list[str]] = []
                cell_paras: list[ParagraphInfo] = []
                for tr in child.findall("w:tr", NS):
                    row: list[str] = []
                    for tc in tr.findall("w:tc", NS):
                        cell_texts: list[str] = []
                        for p in tc.findall("w:p", NS):
                            p_text = paragraph_text(p)
                            cell_texts.append(p_text)
                            info = ParagraphInfo(all_index, p, p_text, self.paragraph_style_id(p), True)
                            self.all_paragraphs.append(info)
                            cell_paras.append(info)
                            all_index += 1
                        row.append("".join(cell_texts))
                    rows.append(row)
                self.tables.append(TableInfo(t_index, child, rows, cell_paras))
                self.body_items.append(BodyItem("tbl", t_index, "\n".join("\t".join(r) for r in rows)))
                t_index += 1

    def full_text(self) -> str:
        chunks = [p.text for p in self.paragraphs]
        for table in self.tables:
            for row in table.rows:
                chunks.extend(row)
        return "\n".join(chunks)

    def media_files(self) -> list[str]:
        return sorted(n for n in self.names if n.startswith("word/media/"))

    def paragraph_style_id(self, p: ET.Element) -> str:
        p_style = p.find("w:pPr/w:pStyle", NS)
        return attr(p_style, "w:val", "") or ""

    def style(self, style_id: str) -> Optional[ET.Element]:
        return self.styles.get(style_id)

    def style_ppr(self, style_id: str) -> Optional[ET.Element]:
        st = self.style(style_id)
        return st.find("w:pPr", NS) if st is not None else None

    def style_rpr(self, style_id: str) -> Optional[ET.Element]:
        st = self.style(style_id)
        return st.find("w:rPr", NS) if st is not None else None

    def ppr(self, p: ET.Element) -> Optional[ET.Element]:
        return p.find("w:pPr", NS)

    def rpr_candidates(self, p: ET.Element) -> list[ET.Element]:
        candidates: list[ET.Element] = []
        # 段落内所有 run（包括嵌套在 w:hyperlink / w:smartTag / w:fldSimple 等容器里的）
        # 都视为格式来源；否则像目录条目这种把 run 全放进 w:hyperlink 的段落，
        # 字体/加粗/字号都会被漏读。
        for r in p.iter(qn("w:r")):
            if paragraph_text(r).strip() or r.find("w:drawing", NS) is not None:
                rpr = r.find("w:rPr", NS)
                if rpr is not None:
                    candidates.append(rpr)
        style_rpr = self.style_rpr(self.paragraph_style_id(p))
        if style_rpr is not None:
            candidates.append(style_rpr)
        return candidates

    def paragraph_jc(self, p: ET.Element) -> str:
        direct = p.find("w:pPr/w:jc", NS)
        if direct is not None:
            return attr(direct, "w:val", "") or ""
        style_jc = self._style_child(p, "w:pPr/w:jc")
        return attr(style_jc, "w:val", "") or ""

    def paragraph_spacing(self, p: ET.Element) -> dict[str, Optional[str]]:
        merged: dict[str, Optional[str]] = {}
        style_spacing = self._style_child(p, "w:pPr/w:spacing")
        direct_spacing = p.find("w:pPr/w:spacing", NS)
        for spacing in (style_spacing, direct_spacing):
            if spacing is None:
                continue
            for key in ["before", "after", "line", "lineRule"]:
                value = attr(spacing, f"w:{key}")
                if value is not None:
                    merged[key] = value
        return merged

    def paragraph_ind(self, p: ET.Element) -> dict[str, Optional[str]]:
        """合并段落最终生效的缩进（含样式继承）。

        优先级（后者覆盖前者）：
        1. 文档默认 `w:docDefaults/w:pPrDefault/w:pPr/w:ind`
        2. 段落样式链（沿 basedOn 自顶向下）
        3. 段落直接属性 `w:pPr/w:ind`
        办公软件实际渲染就是按这个优先级合并的；只看 1 和 3 会漏掉样式默认。
        """
        merged: dict[str, Optional[str]] = {}

        def _absorb(ind: "ET.Element | None") -> None:
            if ind is None:
                return
            for key in ["left", "right", "firstLine", "hanging",
                        "leftChars", "rightChars", "firstLineChars"]:
                value = attr(ind, f"w:{key}")
                if value is not None:
                    merged[key] = value

        # 1. 文档默认
        if self.styles_root is not None:
            doc_default = self.styles_root.find(
                "w:docDefaults/w:pPrDefault/w:pPr/w:ind", NS
            )
            _absorb(doc_default)

        # 2. 样式链（顶到底）
        # style_id 为空时，Word/WPS 实际套用"默认段落样式"（w:default="1" 的 paragraph 样式，
        # 一般是 Normal）。忽略这一步会漏掉默认样式里的缩进/字体等设置。
        chain: list[ET.Element] = []
        seen: set[str] = set()
        sid = self.paragraph_style_id(p)
        if not sid and self.styles_root is not None:
            for st in self.styles_root.findall("w:style", NS):
                if attr(st, "w:type") == "paragraph" and attr(st, "w:default") == "1":
                    sid = attr(st, "w:styleId", "") or ""
                    break
        while sid and sid not in seen:
            seen.add(sid)
            st = self.style(sid)
            if st is None:
                break
            chain.append(st)
            based = st.find("w:basedOn", NS)
            sid = attr(based, "w:val", "") or ""
        for st in reversed(chain):
            _absorb(st.find("w:pPr/w:ind", NS))

        # 3. 段落直接
        _absorb(p.find("w:pPr/w:ind", NS))
        return merged

    def _style_child(self, p: ET.Element, path: str) -> Optional[ET.Element]:
        style_id = self.paragraph_style_id(p)
        st = self.style(style_id)
        return st.find(path, NS) if st is not None else None

    def run_size_pt(self, p: ET.Element) -> Optional[float]:
        for rpr in self.rpr_candidates(p):
            sz = rpr.find("w:sz", NS)
            if sz is not None and attr(sz, "w:val"):
                return half_points_to_pt(attr(sz, "w:val"))
        return None

    def run_fonts(self, p: ET.Element) -> dict[str, str]:
        for rpr in self.rpr_candidates(p):
            fonts = rpr.find("w:rFonts", NS)
            if fonts is not None:
                return {
                    "eastAsia": attr(fonts, "w:eastAsia", "") or attr(fonts, "w:eastAsiaTheme", "") or "",
                    "ascii": attr(fonts, "w:ascii", "") or attr(fonts, "w:asciiTheme", "") or "",
                    "hAnsi": attr(fonts, "w:hAnsi", "") or attr(fonts, "w:hAnsiTheme", "") or "",
                }
        return {"eastAsia": "", "ascii": "", "hAnsi": ""}

    def run_bold(self, p: ET.Element) -> bool:
        for rpr in self.rpr_candidates(p):
            b = rpr.find("w:b", NS)
            if b is not None:
                return bool_element_is_true(b)
        return False

    def run_color(self, p: ET.Element) -> str:
        for rpr in self.rpr_candidates(p):
            color = rpr.find("w:color", NS)
            if color is not None and attr(color, "w:val"):
                return attr(color, "w:val", "") or ""
        return ""

    def outline_level(self, p: ET.Element) -> Optional[int]:
        direct = p.find("w:pPr/w:outlineLvl", NS)
        if direct is not None:
            return to_int(attr(direct, "w:val"))
        style_outline = self._style_child(p, "w:pPr/w:outlineLvl")
        return to_int(attr(style_outline, "w:val")) if style_outline is not None else None

    def has_page_break(self, p: ET.Element) -> bool:
        return any(attr(br, "w:type") == "page" for br in p.findall(".//w:br", NS))

    _real_page_map: "dict[str, int] | None" = None
    _real_page_by_pos: "dict[int, int] | None" = None
    _image_pages: "list[int] | None" = None

    def _build_real_page_map(self) -> dict[str, int]:
        """通过 WPS/Word COM 实拿每段所在真实页码（与办公软件打开后显示一致）。

        - 对每个段落取 `Range.Information(3)` 拿"段落首字光标"所在页；
        - 对每个 InlineShape 取 `Range.Information(3)` 拿"图片自身"所在页
          （图片所在段落首字可能在上一页底部，但图片视觉上已翻到下一页，
           只看段落页码会把图片误判到上一页）；
        - 三份索引（首字文本键、段落 1 基序号、图片 1 基序号）一起回填。
        失败时返回空字典，由调用方退化到 XML 估算。
        """
        # PowerShell 用 -Command 调用时不能用 param() 接收参数；
        # 直接把绝对路径内联进脚本，单引号包裹并把内部单引号转义。
        abs_path = str(self.path.resolve())
        ps_path = abs_path.replace("'", "''")
        ps = (
            '$ErrorActionPreference = "Stop"\n'
            '[Console]::OutputEncoding = [System.Text.UTF8Encoding]::new()\n'
            '$OutputEncoding = [System.Text.UTF8Encoding]::new()\n'
            '$app = $null\n'
            'try { $app = New-Object -ComObject KWPS.Application } catch {}\n'
            'if (-not $app) { try { $app = New-Object -ComObject Word.Application } catch {} }\n'
            'if (-not $app) { Write-Output "NO_COM"; exit 1 }\n'
            '$app.Visible = $false\n'
            f"$doc = $app.Documents.Open('{ps_path}', $false, $true)\n"
            'try {\n'
            '  $i = 0\n'
            '  foreach ($p in $doc.Paragraphs) {\n'
            '    $i++\n'
            '    $page = $p.Range.Information(3)\n'
            '    $t = $p.Range.Text\n'
            '    if ($t.Length -gt 120) { $t = $t.Substring(0,120) }\n'
            '    $t = $t -replace "[\\r\\n\\t \\x07\\x0b\\x0c]", ""\n'
            '    Write-Output ("PG|" + $i + "|" + $page + "|" + $t)\n'
            '  }\n'
            '  $j = 0\n'
            '  foreach ($s in $doc.InlineShapes) {\n'
            '    $j++\n'
            '    $page = $s.Range.Information(3)\n'
            '    Write-Output ("IM|" + $j + "|" + $page)\n'
            '  }\n'
            '} finally {\n'
            '  $doc.Close($false)\n'
            '  $app.Quit()\n'
            '}\n'
        )
        try:
            import subprocess as _sp
            proc = _sp.run(
                ["powershell", "-NoProfile", "-Command", ps],
                capture_output=True, text=True, timeout=180, encoding="utf-8", errors="replace",
            )
        except Exception as exc:
            self.errors.append(f"调用 COM 取真实页码失败：{exc}")
            return {}
        if proc.returncode != 0:
            self.errors.append(f"COM 取页码非0退出：{proc.stderr.strip()[:200]}")
            return {}

        mapping: dict[str, int] = {}
        pos_map: dict[int, int] = {}
        image_pages: list[int] = []
        for line in proc.stdout.splitlines():
            if line.startswith("PG|"):
                parts = line.split("|", 3)
                if len(parts) != 4:
                    continue
                try:
                    pos = int(parts[1])
                    page = int(parts[2])
                except ValueError:
                    continue
                key = parts[3]
                pos_map[pos] = page
                if key:
                    mapping.setdefault(key[:60], page)
            elif line.startswith("IM|"):
                parts = line.split("|", 2)
                if len(parts) != 3:
                    continue
                try:
                    page = int(parts[2])
                except ValueError:
                    continue
                image_pages.append(page)
        DocxReader._real_page_by_pos = pos_map
        DocxReader._image_pages = image_pages
        return mapping

    def real_image_pages(self) -> list[int]:
        """按 InlineShapes 顺序返回每张图片的真实页码（COM 实拿）。"""
        if DocxReader._image_pages is None:
            # 触发 COM 加载
            if DocxReader._real_page_map is None:
                DocxReader._real_page_map = self._build_real_page_map()
        return list(DocxReader._image_pages or [])

    def real_page_of(self, target: ET.Element) -> "int | None":
        """返回该段落在办公软件里打开后实际显示的页码。

        优先用段落文本归一化前缀匹配 COM 输出（与办公软件显示完全一致，
        且不受表内/正文段落混合顺序影响）。空段落或匹配失败时再按段落
        1 基序号兜底；COM 不可用或全失败返回 None。
        """
        if DocxReader._real_page_map is None:
            DocxReader._real_page_map = self._build_real_page_map()
        if not DocxReader._real_page_map and not (DocxReader._real_page_by_pos or {}):
            return None

        text = paragraph_text(target)
        norm = re.sub(r"[\r\n\t \x07\x0b\x0c]", "", text)
        if norm and DocxReader._real_page_map:
            # 段落键缓存最多 60 字符；先尝试 60，再退 40，再退前 20 子串前缀匹配。
            key = norm[:60]
            if key in DocxReader._real_page_map:
                return DocxReader._real_page_map[key]
            key = norm[:40]
            for k, v in DocxReader._real_page_map.items():
                if k.startswith(key) or key.startswith(k):
                    return v
            key = norm[:20]
            for k, v in DocxReader._real_page_map.items():
                if k.startswith(key) or key.startswith(k):
                    return v

        # 空段落或匹配失败：按段落 1 基序号兜底
        pos_map = DocxReader._real_page_by_pos or {}
        if pos_map:
            for i, p in enumerate(self.all_paragraphs, start=1):
                if p.element is target:
                    return pos_map.get(i)
        return None

    def paragraph_page_index(self, target: ET.Element) -> int:
        """估算段落所在页码（1 基）。

        识别三类分页信号，依正文段落顺序累加：
        1. w:br[@w:type="page"] —— 显式插入的分页符
        2. w:lastRenderedPageBreak —— Word 上次渲染时记录的自然分页
        3. w:pPr/w:pageBreakBefore —— 段前分页属性
        """
        page = 1
        for p in self.paragraphs:
            if p.element is target:
                # pageBreakBefore 在该段落之前生效
                if p.element.find("w:pPr/w:pageBreakBefore", NS) is not None:
                    page += 1
                return page
            pbb = p.element.find("w:pPr/w:pageBreakBefore", NS)
            if pbb is not None and bool_element_is_true(pbb):
                page += 1
            # 段内出现的分页符 / 上次渲染分页都计入"该段之后翻页"
            for br in p.element.findall(".//w:br", NS):
                if attr(br, "w:type") == "page":
                    page += 1
            for _ in p.element.findall(".//w:lastRenderedPageBreak", NS):
                page += 1
        return page

    def sections(self) -> list[ET.Element]:
        if self.document is None:
            return []
        return self.document.findall(".//w:sectPr", NS)

    def section_page_settings(self) -> list[dict[str, object]]:
        result: list[dict[str, object]] = []
        for sect in self.sections():
            pg_sz = sect.find("w:pgSz", NS)
            pg_mar = sect.find("w:pgMar", NS)
            result.append(
                {
                    "element": sect,
                    "width_cm": twips_to_cm(attr(pg_sz, "w:w")),
                    "height_cm": twips_to_cm(attr(pg_sz, "w:h")),
                    "orient": attr(pg_sz, "w:orient", "portrait"),
                    "top_cm": twips_to_cm(attr(pg_mar, "w:top")),
                    "bottom_cm": twips_to_cm(attr(pg_mar, "w:bottom")),
                    "left_cm": twips_to_cm(attr(pg_mar, "w:left")),
                    "right_cm": twips_to_cm(attr(pg_mar, "w:right")),
                    "footer_cm": twips_to_cm(attr(pg_mar, "w:footer")),
                    "page_start": attr(sect.find("w:pgNumType", NS), "w:start"),
                    "type": attr(sect.find("w:type", NS), "w:val"),
                    "footer_refs": [f for f in sect.findall("w:footerReference", NS)],
                }
            )
        return result

    def image_references(self) -> list[dict[str, object]]:
        images: list[dict[str, object]] = []
        for pinfo in self.paragraphs:
            for blip in pinfo.element.findall(".//a:blip", NS):
                rid = attr(blip, "r:embed") or attr(blip, "r:link") or ""
                rel = self.relationships.get(rid, {})
                target = rel.get("target", "")
                extent = None
                # 找到离 blip 最近的 drawing 宽高。
                for e in pinfo.element.findall(".//wp:extent", NS):
                    extent = e
                    break
                images.append(
                    {
                        "paragraph_index": pinfo.index,
                        "rid": rid,
                        "target": target,
                        "path": "word/" + target if target and not target.startswith("word/") else target,
                        "width_cm": emu_to_cm(attr(extent, "cx")) if extent is not None else None,
                        "height_cm": emu_to_cm(attr(extent, "cy")) if extent is not None else None,
                        "centered": self.paragraph_jc(pinfo.element) == "center",
                    }
                )
        return images

    def footer_xml_by_rid(self, rid: str) -> Optional[ET.Element]:
        target = self.relationships.get(rid, {}).get("target", "")
        if not target:
            return None
        path = "word/" + target if not target.startswith("word/") else target
        return self.read_xml(path, required=False)

    def footer_infos(self) -> list[dict[str, object]]:
        infos: list[dict[str, object]] = []
        seen: set[str] = set()
        for sect in self.sections():
            for ref in sect.findall("w:footerReference", NS):
                rid = attr(ref, "r:id", "") or ""
                if rid in seen:
                    continue
                seen.add(rid)
                root = self.footer_xml_by_rid(rid)
                target = self.relationships.get(rid, {}).get("target", "")
                text = "" if root is None else "\n".join(paragraph_text(p) for p in root.findall(".//w:p", NS))
                instrs = [] if root is None else [i.text or "" for i in root.findall(".//w:instrText", NS)]
                paras = [] if root is None else root.findall(".//w:p", NS)
                infos.append({"rid": rid, "target": target, "root": root, "text": text, "instrs": instrs, "paragraphs": paras})
        return infos

    def document_instrs(self) -> list[str]:
        if self.document is None:
            return []
        return [i.text or "" for i in self.document.findall(".//w:instrText", NS)]

    def has_toc_field(self) -> bool:
        return any("TOC" in instr.upper() for instr in self.document_instrs())

    def has_hyperlinks(self) -> bool:
        if self.document is None:
            return False
        return bool(self.document.findall(".//w:hyperlink", NS))

    def bookmarks(self) -> set[str]:
        if self.document is None:
            return set()
        return {attr(b, "w:name", "") or "" for b in self.document.findall(".//w:bookmarkStart", NS)}


# ------------------------- 通用格式判断 -------------------------


def find_body_paragraph(reader: DocxReader, keyword: str, normalized: bool = True) -> Optional[ParagraphInfo]:
    needle = normalize_text(keyword) if normalized else keyword
    for p in reader.paragraphs:
        hay = normalize_text(p.text) if normalized else p.text
        if needle in hay:
            return p
    return None


def find_any_paragraph(reader: DocxReader, keyword: str, normalized: bool = True) -> Optional[ParagraphInfo]:
    needle = normalize_text(keyword) if normalized else keyword
    for p in reader.all_paragraphs:
        hay = normalize_text(p.text) if normalized else p.text
        if needle in hay:
            return p
    return None


def paragraph_has_font(reader: DocxReader, p: ParagraphInfo, east_asia: Optional[str] = None, ascii_font: Optional[str] = None) -> bool:
    fonts = reader.run_fonts(p.element)
    ok = True
    if east_asia:
        ok = ok and (east_asia in fonts.get("eastAsia", ""))
    if ascii_font:
        ok = ok and (ascii_font in fonts.get("ascii", "") or ascii_font in fonts.get("hAnsi", ""))
    return ok


def paragraph_size_is(reader: DocxReader, p: ParagraphInfo, pt: float, tol: float = 0.35) -> bool:
    return approx(reader.run_size_pt(p.element), pt, tol)


def paragraph_line_pt(reader: DocxReader, p: ParagraphInfo) -> Optional[float]:
    spacing = reader.paragraph_spacing(p.element)
    line = to_int(spacing.get("line"))
    if line is None:
        return None
    # lineRule=auto 时 240 表示单倍；非 exact 时不是固定磅值。
    if spacing.get("lineRule") == "auto":
        return line / 240.0
    return line / 20.0


def paragraph_before_after_pt(reader: DocxReader, p: ParagraphInfo) -> tuple[Optional[float], Optional[float]]:
    spacing = reader.paragraph_spacing(p.element)
    before = to_int(spacing.get("before"))
    after = to_int(spacing.get("after"))
    return (before / 20.0 if before is not None else None, after / 20.0 if after is not None else None)


def paragraph_first_line_cm(reader: DocxReader, p: ParagraphInfo) -> Optional[float]:
    ind = reader.paragraph_ind(p.element)
    return twips_to_cm(ind.get("firstLine"))


def paragraph_left_cm(reader: DocxReader, p: ParagraphInfo) -> Optional[float]:
    ind = reader.paragraph_ind(p.element)
    return twips_to_cm(ind.get("left"))


def is_center(reader: DocxReader, p: ParagraphInfo) -> bool:
    return reader.paragraph_jc(p.element) == "center"


def is_left_or_default(reader: DocxReader, p: ParagraphInfo) -> bool:
    return reader.paragraph_jc(p.element) in {"", "left"}


def is_justified(reader: DocxReader, p: ParagraphInfo) -> bool:
    return reader.paragraph_jc(p.element) in {"both", "distribute", ""}


def table_has_borders(tbl: ET.Element) -> bool:
    """直接边框检查：本表 `w:tblBorders` 含非 nil 边框，或任一 `w:tcBorders` 存在。

    样式继承的边框由 `module_table_borders_complete` 单独判定。
    """
    borders = tbl.find("w:tblPr/w:tblBorders", NS)
    if borders is not None:
        for child in list(borders):
            if attr(child, "w:val") not in {None, "nil", "none"}:
                return True
    return tbl.find(".//w:tcBorders", NS) is not None


def _collect_tbl_borders_from_style(reader: "DocxReader", style_id: str) -> dict[str, str]:
    """沿表样式 basedOn 链合并 `w:tblBorders` 各方向，返回 {top/left/.../insideV: val}。"""
    merged: dict[str, str] = {}
    seen: set[str] = set()
    sid = style_id
    chain: list[ET.Element] = []
    while sid and sid not in seen:
        seen.add(sid)
        st = reader.style(sid)
        if st is None:
            break
        chain.append(st)
        based = st.find("w:basedOn", NS)
        sid = attr(based, "w:val", "") or ""
    for st in reversed(chain):
        borders = st.find("w:tblPr/w:tblBorders", NS)
        if borders is None:
            continue
        for child in list(borders):
            tag = child.tag.split("}", 1)[-1]
            val = attr(child, "w:val")
            if val is not None:
                merged[tag] = val
    return merged


def module_table_borders_complete(reader: "DocxReader", tbl: ET.Element) -> bool:
    """外框 + 内部横竖线连续完整：
    six 个方向 top/left/bottom/right/insideH/insideV 都必须存在且非 `nil/none`。
    依次查：表直接 `w:tblBorders` → 表样式链 `w:tblBorders`。
    """
    directions = ["top", "left", "bottom", "right", "insideH", "insideV"]
    # 直接 tblBorders
    direct = tbl.find("w:tblPr/w:tblBorders", NS)
    have: dict[str, str] = {}
    if direct is not None:
        for child in list(direct):
            tag = child.tag.split("}", 1)[-1]
            val = attr(child, "w:val")
            if val is not None:
                have[tag] = val

    # 样式链 tblBorders（仅补齐未在 direct 出现的方向）
    style_ref = tbl.find("w:tblPr/w:tblStyle", NS)
    style_id = attr(style_ref, "w:val", "") or ""
    if style_id:
        style_borders = _collect_tbl_borders_from_style(reader, style_id)
        for k, v in style_borders.items():
            if k not in have:
                have[k] = v

    for d in directions:
        v = have.get(d)
        if v is None or v in {"nil", "none"}:
            return False
    return True


def cell_vertical_center(tc: ET.Element) -> bool:
    v_align = tc.find("w:tcPr/w:vAlign", NS)
    return attr(v_align, "w:val") == "center"


# ------------------------- 维度1 -------------------------


def check_dimension1(reader: DocxReader) -> tuple[bool, list[tuple[bool, str]]]:
    details: list[tuple[bool, str]] = []
    path = reader.path

    exists = path.exists()
    details.append((exists, f"文件存在：{path}" if exists else f"文件不存在：{path}"))
    if not exists:
        return False, details

    ext_ok = path.suffix.lower() == ".docx"
    details.append((ext_ok, "交付文件扩展名为 .docx" if ext_ok else f"扩展名不是 .docx：{path.suffix}"))

    size = path.stat().st_size
    details.append((size > 0, f"文件大小 {size} 字节" if size > 0 else "文件大小为 0"))

    zip_ok = zipfile.is_zipfile(path)
    details.append((zip_ok, "文件是合法 zip/docx 包" if zip_ok else "文件不是合法 zip/docx 包"))
    if not (ext_ok and size > 0 and zip_ok):
        return False, details

    try:
        reader.load()
        details.append((True, "docx 包可打开，核心 XML 已读取"))
    except Exception as exc:
        details.append((False, f"docx 打开或核心 XML 解析失败：{exc}"))
        return False, details

    required = ["[Content_Types].xml", "_rels/.rels", "word/document.xml"]
    for name in required:
        details.append((reader.has(name), f"存在必需文件 {name}" if reader.has(name) else f"缺少必需文件 {name}"))

    body_ok = reader.body is not None
    details.append((body_ok, "word/document.xml 中存在 w:body" if body_ok else "word/document.xml 中缺少 w:body"))

    return all(ok for ok, _ in details), details


# ------------------------- 维度2检查 -------------------------


def check_d2_01_page_setup(reader: DocxReader) -> CheckResult:
    sections = reader.section_page_settings()
    if not sections:
        return CheckResult(False, "未检测到页面 section 设置")
    # 细则点 1-7：A4 纵向；宽 21cm、高 29.7cm；上 3.6 / 下 3.4 / 左 2.7 / 右 2.5 cm
    page_ok = all(
        approx(s["width_cm"], 21.0, 0.08)
        and approx(s["height_cm"], 29.7, 0.08)
        and s["orient"] in {None, "", "portrait"}
        and approx(s["top_cm"], 3.6, 0.08)
        and approx(s["bottom_cm"], 3.4, 0.08)
        and approx(s["left_cm"], 2.7, 0.08)
        and approx(s["right_cm"], 2.5, 0.08)
        for s in sections
    )

    # 细则点 8：第1页至第10页正文、图片和表格均位于页面左右边距内，
    # 不出现文字、图片或表格超出页面边界。
    observed = sections[0]
    width_cm = float(observed["width_cm"]) if isinstance(observed["width_cm"], (int, float)) else 21.0
    left_cm = float(observed["left_cm"]) if isinstance(observed["left_cm"], (int, float)) else 2.7
    right_cm = float(observed["right_cm"]) if isinstance(observed["right_cm"], (int, float)) else 2.5
    usable_width_cm = width_cm - left_cm - right_cm
    tol_cm = 0.2

    # 图片越界（宽度超出左右边距内可用宽度）
    image_over = [
        img for img in reader.image_references()
        if isinstance(img["width_cm"], (int, float)) and float(img["width_cm"]) > usable_width_cm + tol_cm
    ]

    # 表格越界：tblW 以 dxa(twips) 给出绝对宽度时，与可用宽度比较
    table_over = 0
    for tbl in reader.tables:
        tbl_w = tbl.element.find("w:tblPr/w:tblW", NS)
        if tbl_w is None:
            continue
        if attr(tbl_w, "w:type") == "dxa":
            w_cm = twips_to_cm(attr(tbl_w, "w:w"))
            if w_cm is not None and w_cm > usable_width_cm + tol_cm:
                table_over += 1

    # 文字越界：段落的左/右缩进若为负，或左缩进+右缩进 >= 可用宽度，
    # 视为文字溢出左右边距。
    text_over = 0
    for p in reader.paragraphs:
        ind = reader.paragraph_ind(p.element)
        left = twips_to_cm(ind.get("left")) or 0.0
        right = twips_to_cm(ind.get("right")) or 0.0
        if left < -tol_cm or right < -tol_cm:
            text_over += 1
            continue
        if left + right >= usable_width_cm:
            text_over += 1

    object_ok = not image_over and table_over == 0 and text_over == 0

    evidence = (
        f"section {len(sections)} 个；首个页面 {width_cm:.2f}×"
        f"{float(observed['height_cm']) if isinstance(observed['height_cm'], (int, float)) else 0.0:.2f}cm，"
        f"边距 上{float(observed['top_cm']) if isinstance(observed['top_cm'], (int, float)) else 0.0:.2f}/"
        f"下{float(observed['bottom_cm']) if isinstance(observed['bottom_cm'], (int, float)) else 0.0:.2f}/"
        f"左{left_cm:.2f}/右{right_cm:.2f}cm；"
        f"可用宽度{usable_width_cm:.2f}cm；图片越界 {len(image_over)}，表格越界 {table_over}，文字越界 {text_over}"
    )
    return CheckResult(page_ok and object_ok, evidence)


def _adjacent_field_content(reader: DocxReader, p: ParagraphInfo, label: str) -> tuple[bool, str]:
    """六字段"对应内容"限定在：同段（标签及冒号之外仍有文本）
    或 同一表格行的紧邻单元格（p 在表格内时）
    或 紧邻的下一个非空段落（p 在正文时，index+1，仅限相邻）。
    返回 (是否存在对应内容, 证据串)。
    """
    # 同段：去掉标签与冒号后仍存在实际内容
    stripped = p.text.replace(label, "").replace("：", "").replace(":", "").strip()
    if stripped:
        return True, f"同段:{stripped[:20]}"

    # 若段落位于表格单元格内，查找紧邻的下一个单元格
    if p.in_table and reader.body is not None:
        for tbl in reader.body.findall(".//w:tbl", NS):
            for tr in tbl.findall("w:tr", NS):
                tcs = tr.findall("w:tc", NS)
                for i, tc in enumerate(tcs):
                    if any(tp is p.element for tp in tc.findall("w:p", NS)):
                        # 找到当前单元格：允许"同单元格其它段落"或"紧邻下一单元格"含内容
                        same_cell_rest = "".join(
                            paragraph_text(tp) for tp in tc.findall("w:p", NS) if tp is not p.element
                        ).strip()
                        if same_cell_rest and label not in same_cell_rest:
                            return True, f"同单元格:{same_cell_rest[:20]}"
                        if i + 1 < len(tcs):
                            nxt_text = "".join(
                                paragraph_text(tp) for tp in tcs[i + 1].findall("w:p", NS)
                            ).strip()
                            if nxt_text and label not in nxt_text:
                                return True, f"邻单元格:{nxt_text[:20]}"
                        return False, "邻近单元格无对应内容"

    # 非表格：只看 all_paragraphs 中紧邻的下一段（index+1），且必须非空且不含标签自身
    nxt = next((q for q in reader.all_paragraphs if q.index == p.index + 1), None)
    if nxt is not None:
        txt = nxt.text.strip()
        if txt and label not in txt:
            return True, f"下一段:{txt[:20]}"
    return False, "紧邻段落无对应内容"


def check_d2_02_cover_content_font(reader: DocxReader) -> CheckResult:
    checks: list[tuple[str, bool, str]] = []

    # 点1：江海应用技术学院 完整保留 + 宋体二号加粗，且位于第1页
    school = find_body_paragraph(reader, "江海应用技术学院")
    if school and normalize_text("江海应用技术学院") in normalize_text(school.text):
        page_no = _real_page_or_xml(reader, school.element)
        on_p1 = page_no == 1
        ok = (
            on_p1
            and paragraph_has_font(reader, school, "宋体")
            and paragraph_size_is(reader, school, FONT_PT["二号"], 0.35)
            and reader.run_bold(school.element)
        )
        checks.append((
            "江海应用技术学院 宋体二号加粗(第1页)",
            ok,
            f"页码{page_no} 字体{reader.run_fonts(school.element)} 字号{reader.run_size_pt(school.element)} 加粗{reader.run_bold(school.element)}",
        ))
    else:
        checks.append(("江海应用技术学院 完整保留", False, "未找到"))

    # 点2：毕业设计成果说明书 完整保留 + 黑体小一加粗，且位于第1页
    main_title = find_body_paragraph(reader, "毕业设计成果说明书")
    if main_title and normalize_text("毕业设计成果说明书") in normalize_text(main_title.text):
        page_no = _real_page_or_xml(reader, main_title.element)
        on_p1 = page_no == 1
        ok = (
            on_p1
            and paragraph_has_font(reader, main_title, "黑体")
            and paragraph_size_is(reader, main_title, FONT_PT["小一"], 0.35)
            and reader.run_bold(main_title.element)
        )
        checks.append((
            "毕业设计成果说明书 黑体小一加粗(第1页)",
            ok,
            f"页码{page_no} 字体{reader.run_fonts(main_title.element)} 字号{reader.run_size_pt(main_title.element)} 加粗{reader.run_bold(main_title.element)}",
        ))
    else:
        checks.append(("毕业设计成果说明书 完整保留", False, "未找到"))

    # 点3：（全日制专科生）完整保留 + 宋体小三，且位于第1页
    sub = find_body_paragraph(reader, "全日制专科生")
    if sub and normalize_text("(全日制专科生)") in normalize_text(sub.text):
        page_no = _real_page_or_xml(reader, sub.element)
        on_p1 = page_no == 1
        ok = (
            on_p1
            and paragraph_has_font(reader, sub, "宋体")
            and paragraph_size_is(reader, sub, FONT_PT["小三"], 0.35)
        )
        checks.append((
            "（全日制专科生）宋体小三(第1页)",
            ok,
            f"页码{page_no} 字体{reader.run_fonts(sub.element)} 字号{reader.run_size_pt(sub.element)}",
        ))
    else:
        checks.append(("（全日制专科生）完整保留", False, "未找到"))

    # 点4-9：六个字段标签 + 对应内容 完整保留（要求：标签在第1页，且对应内容位于同段/同表格紧邻单元格/紧邻下一段）
    # 细则只要求"对应内容"存在，不约束字段标签/内容的字体字号。
    field_labels = ["毕业设计题目", "二级学院", "专业班级", "学生学号", "学生姓名", "指导老师"]
    field_missing: list[str] = []
    field_not_p1: list[str] = []
    field_no_content: list[str] = []
    field_details: list[str] = []
    for label in field_labels:
        p = find_any_paragraph(reader, label)
        if p is None:
            field_missing.append(label)
            continue
        page_no = _real_page_or_xml(reader, p.element)
        if page_no != 1:
            field_not_p1.append(f"{label}(第{page_no}页)")
        has_content, evi = _adjacent_field_content(reader, p, label)
        if not has_content:
            field_no_content.append(label)
        field_details.append(f"{label}:页{page_no} {evi}")
    checks.append((
        "六字段标签及对应内容完整保留(第1页,同段/邻单元格/邻段)",
        not field_missing and not field_not_p1 and not field_no_content,
        f"缺失标签 {field_missing or '无'}；非第1页 {field_not_p1 or '无'}；缺失对应内容 {field_no_content or '无'}；明细 {field_details}",
    ))

    # 点10：结构方案、控制布置与使用性能说明 完整保留 + 宋体四号，且位于第1页
    plan = find_body_paragraph(reader, "结构方案、控制布置与使用性能说明")
    if plan and normalize_text("结构方案、控制布置与使用性能说明") in normalize_text(plan.text):
        page_no = _real_page_or_xml(reader, plan.element)
        on_p1 = page_no == 1
        ok = (
            on_p1
            and paragraph_has_font(reader, plan, "宋体")
            and paragraph_size_is(reader, plan, FONT_PT["四号"], 0.35)
        )
        checks.append((
            "结构方案、控制布置与使用性能说明 宋体四号(第1页)",
            ok,
            f"页码{page_no} 字体{reader.run_fonts(plan.element)} 字号{reader.run_size_pt(plan.element)}",
        ))
    else:
        checks.append(("结构方案、控制布置与使用性能说明 完整保留", False, "未找到"))

    fmt_bad = [f"{name}({ev})" for name, ok, ev in checks if not ok]
    passed = not fmt_bad
    evidence = "格式不符 " + (str(fmt_bad) if fmt_bad else "无")
    return CheckResult(passed, evidence)


def check_d2_03_cover_field_paragraphs(reader: DocxReader) -> CheckResult:
    labels = ["毕业设计题目", "二级学院", "专业班级", "学生学号", "学生姓名", "指导老师"]
    found: list[str] = []
    ok_count = 0
    details: list[str] = []
    for label in labels:
        p = find_any_paragraph(reader, label)
        if not p:
            details.append(f"{label}: 未找到")
            continue
        found.append(label)
        # 细则点 1-2：段前段后都为 5 磅
        before, after = paragraph_before_after_pt(reader, p)
        before_ok = approx(before, 5, 0.1)
        after_ok = approx(after, 5, 0.1)
        # 细则点 3：文本前缩进 2.47cm
        left = paragraph_left_cm(reader, p)
        left_ok = approx(left, 2.47, 0.03)
        # 细则点 4：行距为固定值 25 磅
        spacing = reader.paragraph_spacing(p.element)
        line_rule = spacing.get("lineRule")
        line_twips = to_int(spacing.get("line"))
        line_pt = line_twips / 20.0 if line_twips is not None else None
        line_ok = line_rule == "exact" and approx(line_pt, 25, 0.1)
        # 细则点 5：左对齐（明确为 left，不接受未设置/两端对齐等）
        align = reader.paragraph_jc(p.element)
        align_ok = align == "left"

        ok = before_ok and after_ok and left_ok and line_ok and align_ok
        ok_count += int(ok)
        details.append(
            f"{label}: 段前{before}磅 段后{after}磅 左缩进{left}cm 行距{line_pt}磅({line_rule}) 对齐{align!r}"
        )
    return CheckResult(
        len(found) == len(labels) and ok_count == len(labels),
        f"找到 {len(found)}/6；格式满足 {ok_count}/6；" + "；".join(details),
    )


def check_d2_04_originality_title(reader: DocxReader) -> CheckResult:
    # 细则点 1：文本"毕业设计独创性声明"（精确文本，去空白后等值）
    p = next(
        (
            q for q in reader.paragraphs
            if normalize_text(q.text) == normalize_text("毕业设计独创性声明")
        ),
        None,
    )
    if not p:
        return CheckResult(False, "未找到标题文本“毕业设计独创性声明”")

    # 约束：必须出现在第2页（以办公软件实际打开后的页码为准）
    page_no = reader.real_page_of(p.element)
    if page_no is None:
        # COM 不可用时退化到 XML 估算
        page_no = reader.paragraph_page_index(p.element)
    if page_no != 2:
        return CheckResult(False, f"标题“毕业设计独创性声明”未出现在第2页（实际第{page_no}页）")

    # 细则点 2：字体为黑体
    font_ok = paragraph_has_font(reader, p, "黑体")
    # 细则点 3：字号三号(16pt)
    size_ok = paragraph_size_is(reader, p, FONT_PT["三号"], 0.1)
    # 细则点 4：加粗
    bold_ok = reader.run_bold(p.element)
    # 细则点 5：水平居中
    center_ok = is_center(reader, p)
    # 细则点 6：20 磅行距（细则未指明固定/最小值，按磅值数核对即可）
    line_pt = paragraph_line_pt(reader, p)
    line_ok = approx(line_pt, 20, 0.1)

    ok = font_ok and size_ok and bold_ok and center_ok and line_ok
    return CheckResult(
        ok,
        (
            f"字体{reader.run_fonts(p.element)} 字号{reader.run_size_pt(p.element)} "
            + f"加粗{bold_ok} 居中{center_ok} 行距{line_pt}磅"
        ),
    )


def check_d2_05_originality_body(reader: DocxReader) -> CheckResult:
    # 定位"毕业设计独创性声明"标题段，之后到"学生签字"之前的非空段落即为声明正文。
    start = next(
        (p.index for p in reader.paragraphs if "毕业设计独创性声明" in p.text),
        -1,
    )
    if start < 0:
        return CheckResult(False, "未找到“毕业设计独创性声明”标题，无法定位正文段落")

    candidates: list[ParagraphInfo] = []
    for p in reader.paragraphs:
        if p.index <= start:
            continue
        if "学生签字" in p.text:
            break
        if p.text.strip():
            candidates.append(p)
    if not candidates:
        return CheckResult(False, "未找到独创性声明正文段落")

    # 必须出现在第2页（以办公软件实际页码为准）
    off_page: list[tuple[str, "int | None"]] = []
    for p in candidates:
        page_no = reader.real_page_of(p.element)
        if page_no is None:
            page_no = reader.paragraph_page_index(p.element)
        if page_no != 2:
            off_page.append((p.text[:20], page_no))
    if off_page:
        return CheckResult(False, f"独创性声明正文未全部出现在第2页：{off_page}")

    # 细则点 1：完整保留 本人声明 / 成果来源 / 责任承担 三类内容
    text = normalize_text("".join(p.text for p in candidates))
    statement_ok = "本人" in text and "声明" in text
    source_ok = "成果" in text
    responsibility_ok = "承担" in text or "责任" in text
    content_ok = statement_ok and source_ok and responsibility_ok

    # 细则点 2-4：中文宋体小四号 / 首行缩进 0.85cm / 固定值 20 磅行距
    fmt_ok = 0
    for p in candidates:
        font_ok = paragraph_has_font(reader, p, "宋体")
        size_ok = paragraph_size_is(reader, p, FONT_PT["小四"], 0.1)
        indent_ok = approx(paragraph_first_line_cm(reader, p), 0.85, 0.03)
        spacing = reader.paragraph_spacing(p.element)
        line_twips = to_int(spacing.get("line"))
        line_pt = line_twips / 20.0 if line_twips is not None else None
        line_ok = spacing.get("lineRule") == "exact" and approx(line_pt, 20, 0.1)
        if font_ok and size_ok and indent_ok and line_ok:
            fmt_ok += 1

    passed = content_ok and fmt_ok == len(candidates)
    return CheckResult(
        passed,
        (
            f"声明正文段落 {len(candidates)}，格式满足 {fmt_ok}/{len(candidates)}；"
            + f"本人声明 {statement_ok}，成果来源 {source_ok}，责任承担 {responsibility_ok}"
        ),
    )


def _real_page_or_xml(reader: DocxReader, p_element: ET.Element) -> int:
    """优先取办公软件实拿页码，COM 失败时退化到 XML 估算。"""
    page_no = reader.real_page_of(p_element)
    if page_no is None:
        page_no = reader.paragraph_page_index(p_element)
    return page_no


def check_d2_06_signature_area(reader: DocxReader) -> CheckResult:
    # 细则点 1：完整保留"学生签字："和"日期：　年　月　日"
    p = find_body_paragraph(reader, "学生签字")
    if p is None:
        return CheckResult(False, "未找到“学生签字”")
    # 必须出现在第2页
    page_no = _real_page_or_xml(reader, p.element)
    if page_no != 2:
        return CheckResult(False, f"签字区域未出现在第2页（实际第{page_no}页）")

    text = p.text
    has_sign = "学生签字" in text and ("：" in text or ":" in text)
    has_date = "日期" in text and "年" in text and "月" in text and "日" in text
    # 细则点 2：文字与填写空白位于同一行或相邻区域
    # 同一行 = 同一段落同时包含"学生签字"和"日期"；相邻区域 = 紧邻下一段含"日期"。
    same_line = has_sign and has_date
    adjacent = False
    if has_sign and not has_date:
        nxt = next((q for q in reader.paragraphs if q.index == p.index + 1), None)
        adjacent = nxt is not None and "日期" in nxt.text and "年" in nxt.text and "月" in nxt.text and "日" in nxt.text
    layout_ok = same_line or adjacent
    content_ok = has_sign and (has_date or adjacent)

    # 细则点 3：段前 24 磅
    before, _ = paragraph_before_after_pt(reader, p)
    before_ok = approx(before, 24, 0.1)
    # 细则点 4：行距固定值 22 磅
    spacing = reader.paragraph_spacing(p.element)
    line_twips = to_int(spacing.get("line"))
    line_pt = line_twips / 20.0 if line_twips is not None else None
    line_ok = spacing.get("lineRule") == "exact" and approx(line_pt, 22, 0.1)

    ok = content_ok and layout_ok and before_ok and line_ok
    return CheckResult(
        ok,
        (
            f"学生签字 {has_sign}；日期同段{same_line}或相邻{adjacent}；"
            + f"段前{before}磅；行距{line_pt}磅({spacing.get('lineRule')})"
        ),
    )


def check_d2_07_toc_title(reader: DocxReader) -> CheckResult:
    # 定位目录标题段（按"目"+"录"宽松定位，再精确校对字符序列）
    p = next(
        (
            q for q in reader.paragraphs
            if "目" in q.text and "录" in q.text and len(q.text.strip()) <= 8
        ),
        None,
    )
    if p is None:
        return CheckResult(False, "未找到含“目”“录”的标题段落")

    # 细则点 1：必须在第3页（COM 实拿页码）
    page_no = _real_page_or_xml(reader, p.element)
    if page_no != 3:
        return CheckResult(False, f"目录标题未出现在第3页（实际第{page_no}页）")

    # 细则点 2：文本为"目　　录"——两字之间**两个全角字符**（U+3000）
    text_ok = p.text.strip() == "目　　录"

    # 细则点 3：字体黑体
    font_ok = paragraph_has_font(reader, p, "黑体")
    # 细则点 4：字号小三(15pt)
    size_ok = paragraph_size_is(reader, p, FONT_PT["小三"], 0.1)
    # 细则点 5：加粗
    bold_ok = reader.run_bold(p.element)
    # 细则点 6：水平居中
    center_ok = is_center(reader, p)
    # 细则点 7：行距 1.5 倍
    sp = reader.paragraph_spacing(p.element)
    line_twips = to_int(sp.get("line"))
    line_rule = sp.get("lineRule")
    # 1.5 倍行距：lineRule=auto 且 line=360（一倍=240，1.5 倍=360）；
    # 或 lineRule=multiple 且 line=360。
    line_ok = line_rule in {"auto", "multiple", None} and line_twips == 360

    ok = text_ok and font_ok and size_ok and bold_ok and center_ok and line_ok
    return CheckResult(
        ok,
        (
            f"文本{p.text!r}（两全角空格={text_ok}） 字体{reader.run_fonts(p.element)} "
            + f"字号{reader.run_size_pt(p.element)} 加粗{bold_ok} 居中{center_ok} "
            + f"行距 line={line_twips} rule={line_rule}（1.5倍={line_ok}）"
        ),
    )


def toc_paragraphs(reader: DocxReader) -> list[ParagraphInfo]:
    start = next((p.index for p in reader.paragraphs if "录" in p.text and "目" in p.text), -1)
    if start < 0:
        return []
    items: list[ParagraphInfo] = []
    for p in reader.paragraphs:
        if p.index <= start:
            continue
        if re.match(r"\s*[1-6]\s+", p.text):
            items.append(p)
        elif items:
            break
    return items


def check_d2_08_auto_toc_object(reader: DocxReader) -> CheckResult:
    items = toc_paragraphs(reader)
    if not items:
        return CheckResult(False, "未找到目录条目")

    # 细则点 1：必须在第3页
    off_page = [
        (p.text[:15], _real_page_or_xml(reader, p.element))
        for p in items
        if _real_page_or_xml(reader, p.element) != 3
    ]
    if off_page:
        return CheckResult(False, f"目录条目未全部在第3页：{off_page}")

    # 细则点 2：目录为 Word 目录域或等效自动目录对象
    # —— "Word 目录域" = 含 TOC 字段；
    # —— "等效自动目录对象" = 每个目录条目都是 Word 字段（HYPERLINK \l "书签"），
    #    可 Ctrl+单击跳转、右键可"更新域"，且锚点书签在文档内存在；
    # —— 排除"只是普通文本和手工圆点"：要求 6 个条目都不是纯文本。
    bookmarks = reader.bookmarks()
    has_toc_field = reader.has_toc_field()
    field_anchored = 0
    for p in items:
        # 段内是否存在 HYPERLINK 字段（含 instrText 以 HYPERLINK 开头），且 anchor 指向有效书签
        instr_texts = [i.text or "" for i in p.element.findall(".//w:instrText", NS)]
        is_hyperlink_field = any(it.strip().upper().startswith("HYPERLINK") for it in instr_texts)
        hyperlinks = p.element.findall(".//w:hyperlink", NS)
        anchors = [attr(h, "w:anchor", "") or "" for h in hyperlinks]
        anchor_valid = any(a and a in bookmarks for a in anchors)
        if (is_hyperlink_field or hyperlinks) and anchor_valid:
            field_anchored += 1
    equivalent_auto = field_anchored == len(items)
    object_ok = has_toc_field or equivalent_auto

    # 细则点 3：依次显示"1　　设计思路""2　　设计内容""3　　设计过程""4　　作品及特点""5　　致谢""6　　参考资料"
    expected = [
        ("1", "设计思路"),
        ("2", "设计内容"),
        ("3", "设计过程"),
        ("4", "作品及特点"),
        ("5", "致谢"),
        ("6", "参考资料"),
    ]
    sequence_detail: list[str] = []
    order_ok = True
    if len(items) != 6:
        order_ok = False
        sequence_detail.append(f"条目数 {len(items)}≠6")
    for idx, (num, name) in enumerate(expected):
        if idx >= len(items):
            break
        t = items[idx].text
        # 必须匹配 "N　　名称" —— 序号与名称之间两个全角空格（U+3000）
        # 制表符/普通空格不接受（细则原文为两个全角字符）
        if not re.match(rf"^[\s\t]*{re.escape(num)}　　{re.escape(name)}", t):
            order_ok = False
            sequence_detail.append(f"第{idx+1}项不匹配“{num}　　{name}”（实际{t!r}）")

    ok = object_ok and order_ok
    return CheckResult(
        ok,
        (
            f"TOC域 {has_toc_field}；字段化条目 {field_anchored}/{len(items)}；"
            + f"等效自动目录 {equivalent_auto}；顺序与文本 {order_ok}；"
            + (f"{sequence_detail}" if sequence_detail else "无异常")
        ),
    )


def check_d2_09_toc_hyperlinks(reader: DocxReader) -> CheckResult:
    items = toc_paragraphs(reader)
    if not items:
        return CheckResult(False, "未找到目录条目")
    # 必须出现在第3页
    off_page = [(p.text[:15], _real_page_or_xml(reader, p.element)) for p in items if _real_page_or_xml(reader, p.element) != 3]
    if off_page:
        return CheckResult(False, f"目录条目未全部在第3页：{off_page}")

    # 6 个目录条目都包含 hyperlink，且 anchor 指向文档内书签
    bookmarks = reader.bookmarks()
    required_names = ["设计思路", "设计内容", "设计过程", "作品及特点", "致谢", "参考资料"]
    linked = 0
    detail: list[str] = []
    for name in required_names:
        item = next((p for p in items if name in p.text), None)
        if item is None:
            detail.append(f"{name}: 条目缺失")
            continue
        hyperlinks = item.element.findall(".//w:hyperlink", NS)
        anchors = [attr(h, "w:anchor", "") or "" for h in hyperlinks]
        # anchor 必须指向文档内有效书签
        anchor_ok = any(a and a in bookmarks for a in anchors)
        if hyperlinks and anchor_ok:
            linked += 1
        else:
            detail.append(f"{name}: 超链接{len(hyperlinks)} 锚点{anchors}")

    ok = linked == 6
    return CheckResult(ok, f"6条目超链接命中 {linked}/6；{'；'.join(detail) or '全部有效'}")


def check_d2_10_toc_entry_format(reader: DocxReader) -> CheckResult:
    items = toc_paragraphs(reader)
    if not items:
        return CheckResult(False, "未找到目录条目")
    # 必须出现在第3页
    off_page = [(p.text[:15], _real_page_or_xml(reader, p.element)) for p in items if _real_page_or_xml(reader, p.element) != 3]
    if off_page:
        return CheckResult(False, f"目录条目未全部在第3页：{off_page}")

    primary_names = {"设计思路", "设计内容", "设计过程", "作品及特点", "致谢", "参考资料"}
    fail_detail: list[str] = []
    for p in items:
        # 细则：中文宋体小四号，英文数字 Times New Roman 小四号
        fonts = reader.run_fonts(p.element)
        font_ok = "宋体" in fonts.get("eastAsia", "") and (
            "Times New Roman" in fonts.get("ascii", "")
            or "Times New Roman" in fonts.get("hAnsi", "")
        )
        size_ok = paragraph_size_is(reader, p, FONT_PT["小四"], 0.1)

        # 段后 5 磅，行距固定值 19 磅
        before, after = paragraph_before_after_pt(reader, p)
        after_ok = approx(after, 5, 0.1)
        sp = reader.paragraph_spacing(p.element)
        line_twips = to_int(sp.get("line"))
        line_pt = line_twips / 20.0 if line_twips is not None else None
        line_ok = sp.get("lineRule") == "exact" and approx(line_pt, 19, 0.1)

        # 点状引导符
        tab_leader = p.element.find("w:pPr/w:tabs/w:tab", NS)
        leader_ok = (tab_leader is not None and attr(tab_leader, "w:leader") in {"dot", "middleDot"}) or "……" in p.text or "..." in p.text

        # 一级条目：序号与名称之间空两个字符 + 加粗
        # "两个字符" 的允许形式：两个全角空格(U+3000 × 2) 或 两个 ASCII 空格；
        # 不接受 3 个及以上、不接受制表符/换行、不接受半/全角混排。
        is_primary = any(name in p.text for name in primary_names) and bool(re.match(r"^\s*[1-6][　 ]", p.text))
        if is_primary:
            two_space_ok = bool(re.match(r"^\s*[1-6](?:　　|  )\S", p.text))
            bold_ok = reader.run_bold(p.element)
            indent_ok = True
            secondary_unbold_ok = True
        else:
            # 细则点：二级条目首行缩进约1个字符且不加粗
            # 小四(12pt)中文字符宽约0.42cm；接受 0.3-0.6cm
            first_line = paragraph_first_line_cm(reader, p) or 0.0
            indent_ok = 0.3 <= first_line <= 0.6
            bold_ok = True
            two_space_ok = True
            secondary_unbold_ok = not reader.run_bold(p.element)

        all_ok = font_ok and size_ok and after_ok and line_ok and leader_ok and two_space_ok and bold_ok and indent_ok and secondary_unbold_ok
        if not all_ok:
            fail_detail.append(
                f"{p.text[:15]!r}: 字体{fonts} 字号{reader.run_size_pt(p.element)} "
                + f"段前后{before}/{after} 行距{line_pt}({sp.get('lineRule')}) "
                + f"引导符{leader_ok} 一级{is_primary} 两字符{two_space_ok} "
                + f"加粗{reader.run_bold(p.element)} 缩进{paragraph_first_line_cm(reader, p)}"
            )

    ok = not fail_detail
    return CheckResult(ok, f"目录条目 {len(items)}；不合格 {len(fail_detail)}；{'；'.join(fail_detail[:2]) or '全部合格'}")


def check_d2_11_toc_page_column(reader: DocxReader) -> CheckResult:
    items = toc_paragraphs(reader)
    if not items:
        return CheckResult(False, "未找到目录条目")
    # 必须出现在第3页
    off_page = [(p.text[:15], _real_page_or_xml(reader, p.element)) for p in items if _real_page_or_xml(reader, p.element) != 3]
    if off_page:
        return CheckResult(False, f"目录条目未全部在第3页：{off_page}")

    # 细则点 1：页码右对齐（每条目都需有 w:val="right" 的制表位）
    tab_right_count = 0
    for p in items:
        for tab in p.element.findall("w:pPr/w:tabs/w:tab", NS):
            if attr(tab, "w:val") == "right":
                tab_right_count += 1
                break
    right_ok = tab_right_count == len(items)

    # 细则点 2：显示正文实际起始页码——每条目末尾必须以数字结尾
    ends_with_page = all(re.search(r"\d\s*$", p.text.strip()) for p in items)

    # 细则点 3：不使用"01""02"等带前导零的文本格式
    no_leading_zero = all(not re.search(r"\b0\d+\b\s*$", p.text.strip()) for p in items)

    ok = right_ok and ends_with_page and no_leading_zero
    return CheckResult(
        ok,
        f"右对齐制表位 {tab_right_count}/{len(items)}；末尾页码 {ends_with_page}；无前导零 {no_leading_zero}",
    )


def check_d2_13_section_break_after_toc(reader: DocxReader) -> CheckResult:
    toc_items = toc_paragraphs(reader)
    if not toc_items:
        return CheckResult(False, "未找到目录条目，无法定位目录页末尾")
    last_toc = toc_items[-1].index
    # 细则点 1：目录页末尾设置"下一页"分节符
    found = False
    type_value = ""
    for p in reader.paragraphs:
        if p.index < last_toc:
            continue
        sect = p.element.find("w:pPr/w:sectPr", NS)
        if sect is not None:
            found = True
            type_value = attr(sect.find("w:type", NS), "w:val", "nextPage") or "nextPage"
            break
        if p.index > last_toc + 5:
            break
    type_ok = type_value in {"nextPage", ""}

    # 细则点 2：分节符不能显示为普通正文文字
    text_fake = any("分节符" in p.text for p in reader.paragraphs)

    ok = found and type_ok and not text_fake
    return CheckResult(ok, f"目录后sectPr存在 {found}，类型 {type_value or '未声明'}；文本'分节符' {text_fake}")


def check_d2_14_page_number_position(reader: DocxReader) -> CheckResult:
    # 细则要求：
    #   点1：第1页封面、第2页独创性声明、第3页目录 均不显示页码
    #   点2：第4页正文开始显示阿拉伯数字 "1"，后续正文页码连续递增
    #   点3：页码位于页脚水平居中
    #
    # 判定方法（按 section 与页码范围映射）：
    #   (a) 建立 段落→section 的映射：按 body 顺序遍历，
    #       段落 pPr/sectPr 关闭当前 section；无 pPr/sectPr 的段落归属当前 section；
    #       最后一批段落属于 body 直接子级 sectPr 所在的默认 section。
    #   (b) 建立 页码→section 的映射：对正文段落用 _real_page_or_xml 取其真实页码，
    #       每一页取其首个正文段落所在 section。
    #   (c) 前3页所在 section 的默认页脚必须"无 PAGE 字段 且 文本为空"（或未声明页脚引用）。
    #   (d) 第4页所在 section = 正文 section；其 pgNumType.start == "1"，且其页脚含 PAGE 字段。
    #   (e) 正文 section 之后的 section 不得再次 pgNumType 重置起始页码。
    #   (f) 第5/6页（若存在）所在 section 索引 >= 正文 section 索引。
    #   (g) 正文页脚含 PAGE 字段的段落 jc=center。
    sections = reader.section_page_settings()
    footers = reader.footer_infos()
    if not sections:
        return CheckResult(False, "未找到 sectPr")

    # (a) 段落 → section 索引
    section_of_para: dict[int, int] = {}
    last_sid = max(0, len(sections) - 1)
    cur = 0
    if reader.body is not None:
        for child in list(reader.body):
            if child.tag == qn("w:p"):
                section_of_para[id(child)] = cur
                if child.find("w:pPr/w:sectPr", NS) is not None:
                    cur = min(cur + 1, last_sid)
            elif child.tag == qn("w:tbl"):
                for tp in child.findall(".//w:p", NS):
                    section_of_para[id(tp)] = cur

    # (b) 页码 → section
    page_section: dict[int, int] = {}
    for pinfo in reader.paragraphs:
        pg = _real_page_or_xml(reader, pinfo.element)
        sid = section_of_para.get(id(pinfo.element))
        if sid is None:
            continue
        if pg not in page_section:
            page_section[pg] = sid

    # 工具：某 section 的默认页脚引用
    by_rid = {str(info.get("rid", "")): info for info in footers}

    def _section_footer_infos(sect_idx: int) -> list[dict[str, object]]:
        s = sections[sect_idx]
        refs_raw = s.get("footer_refs", [])
        refs: list[ET.Element] = [r for r in refs_raw if isinstance(r, ET.Element)] if isinstance(refs_raw, list) else []
        out: list[dict[str, object]] = []
        for ref in refs:
            # 只挑选 default 类型（未声明 w:type 视为 default）
            ftype = attr(ref, "w:type", "default") or "default"
            if ftype != "default":
                continue
            rid = attr(ref, "r:id", "") or ""
            info = by_rid.get(rid)
            if info is not None:
                out.append(info)
        return out

    def _footer_has_page(info: dict[str, object]) -> bool:
        instrs_raw = info.get("instrs", [])
        instrs: list[str] = [str(i) for i in instrs_raw] if isinstance(instrs_raw, list) else []
        return any("PAGE" in i.upper() for i in instrs)

    def _footer_is_empty(info: dict[str, object]) -> bool:
        text = normalize_text(str(info.get("text", "") or ""))
        return not text and not _footer_has_page(info)

    # (c) 前3页页脚干净
    front_ok = True
    front_detail: list[str] = []
    for pg in (1, 2, 3):
        sid = page_section.get(pg)
        if sid is None:
            front_ok = False
            front_detail.append(f"第{pg}页未定位到section")
            continue
        infos = _section_footer_infos(sid)
        if not infos:
            # 未声明默认页脚 = 该页无页脚显示，视为满足"不显示页码"
            front_detail.append(f"第{pg}页section{sid}未声明默认页脚")
            continue
        clean = all(_footer_is_empty(info) for info in infos)
        if not clean:
            front_ok = False
            front_detail.append(f"第{pg}页section{sid}页脚含PAGE或有文本")

    # (d)(e)(f) 正文 section
    body_sid = page_section.get(4)
    if body_sid is None:
        return CheckResult(False, f"未定位到第4页；页-section映射 {page_section}")

    body_section = sections[body_sid]
    body_footer_infos = _section_footer_infos(body_sid)
    body_has_page = any(_footer_has_page(i) for i in body_footer_infos)
    start1 = body_section.get("page_start") == "1"

    later_restarts: list[int] = []
    for i in range(body_sid + 1, len(sections)):
        if sections[i].get("page_start") is not None:
            later_restarts.append(i)

    seq_ok = True
    seq_detail: list[str] = []
    for pg in (5, 6):
        sid = page_section.get(pg)
        if sid is None:
            continue
        if sid < body_sid:
            seq_ok = False
            seq_detail.append(f"第{pg}页section{sid}早于正文section{body_sid}")

    # (g) 页脚居中：正文 section 的默认页脚中，含 PAGE 字段的段落 jc=center
    center = False
    for info in body_footer_infos:
        paras_raw = info.get("paragraphs", [])
        paras = list(paras_raw) if isinstance(paras_raw, list) else []
        for p in paras:
            if any("PAGE" in (i.text or "").upper() for i in p.findall(".//w:instrText", NS)):
                center = attr(p.find("w:pPr/w:jc", NS), "w:val") == "center"
                break
        if center:
            break

    ok = (
        front_ok
        and body_has_page
        and start1
        and not later_restarts
        and seq_ok
        and center
    )
    return CheckResult(
        ok,
        (
            f"页→section {page_section}；前3页页脚干净 {front_ok}({front_detail or '无'})；"
            f"正文section {body_sid} PAGE字段 {body_has_page} 起始1 {start1}；"
            f"后续section重置 {later_restarts or '无'}；4/5/6连续 {seq_ok}({seq_detail or '无'})；"
            f"居中 {center}"
        ),
    )


def check_d2_15_footer_page_numbers(reader: DocxReader) -> CheckResult:
    footers = reader.footer_infos()
    sections = reader.section_page_settings()
    page_para: Optional[ET.Element] = None
    for info in footers:
        paragraphs_raw = info.get("paragraphs", [])
        para_list = list(paragraphs_raw) if isinstance(paragraphs_raw, list) else []
        for p in para_list:
            if any("PAGE" in (i.text or "").upper() for i in p.findall(".//w:instrText", NS)):
                page_para = p
                break
    if page_para is None:
        return CheckResult(False, "页脚未检测到 PAGE 自动页码字段")

    pinfo = ParagraphInfo(-1, page_para, paragraph_text(page_para), "", False)
    fonts = reader.run_fonts(page_para)
    size = reader.run_size_pt(page_para)

    # 细则点 1：页脚中间显示阿拉伯数字从"1"开始连续递增
    start1 = any(s["page_start"] == "1" for s in sections)
    # 细则点 2：水平居中
    center = attr(page_para.find("w:pPr/w:jc", NS), "w:val") == "center"
    # 细则点 3：Times New Roman
    tnr_ok = "Times New Roman" in (fonts.get("ascii", "") + fonts.get("hAnsi", ""))
    # 细则点 4：小五号(9pt)
    size_ok = paragraph_size_is(reader, pinfo, FONT_PT["小五"], 0.1)
    # 细则点 5：距离页面底边约 1.3-2.0cm
    footer_dist_ok = any(in_range(s["footer_cm"], 1.3, 2.0) for s in sections)
    # 细则点 6：Word 自动页码字段（已通过 PAGE 字段检测确认）
    auto_field = True

    ok = start1 and center and tnr_ok and size_ok and footer_dist_ok and auto_field
    return CheckResult(
        ok,
        f"起始1 {start1}；居中 {center}；Times {tnr_ok}；字号{size}pt({size_ok})；距底边 {footer_dist_ok}；自动字段 {auto_field}",
    )


def check_d2_16_primary_headings(reader: DocxReader) -> CheckResult:
    required = [
        ("一、设计思路", "设计思路"),
        ("二、设计内容", "设计内容"),
        ("三、设计过程", "设计过程"),
        ("四、作品及特点", "作品及特点"),
        ("五、致谢", "致谢"),
    ]
    full = normalize_text(reader.full_text())
    missing = [zh for zh, _ in required if normalize_text(zh) not in full]
    if missing:
        return CheckResult(False, f"中文编号标题缺失：{missing}")

    actual: list[ParagraphInfo] = []
    for _zh, key in required:
        for p in reader.paragraphs:
            if (p.style_id == "Heading1" or reader.outline_level(p.element) == 0) and key in p.text:
                actual.append(p)
                break

    if len(actual) != 5:
        return CheckResult(False, f"5个一级标题对应 Heading1/大纲0 段落仅找到 {len(actual)} 个")

    fail: list[str] = []
    for p in actual:
        # 第4-10页约束
        page_no = _real_page_or_xml(reader, p.element)
        page_ok = 4 <= page_no <= 10

        # 字体黑体、四号(14pt)、加粗
        font_ok = paragraph_has_font(reader, p, "黑体")
        size_ok = paragraph_size_is(reader, p, FONT_PT["四号"], 0.1)
        bold_ok = reader.run_bold(p.element)
        # 顶格：左缩进与首行缩进都为 0
        left = paragraph_left_cm(reader, p) or 0.0
        first = paragraph_first_line_cm(reader, p) or 0.0
        top_align_ok = abs(left) < 0.03 and abs(first) < 0.03

        # 上方和下方各保留约一行空白：用段前/段后磅值代理（半行高及以上即接受）
        before, after = paragraph_before_after_pt(reader, p)
        gap_ok = (before or 0) >= 10 and (after or 0) >= 10

        # 标题不单独位于页面底部：widowControl/keepNext 任一启用
        ppr = reader.ppr(p.element)
        keep_next = ppr is not None and ppr.find("w:keepNext", NS) is not None
        # 段后必须存在非空段落（标题后至少保留一段正文）
        next_text = next((q for q in reader.paragraphs if q.index > p.index and q.text.strip()), None)
        has_body = next_text is not None and next_text.text.strip() not in {zh for zh, _ in required}
        not_orphan_ok = keep_next or has_body

        # 应用 Heading1 样式或大纲级别1
        style_ok = p.style_id == "Heading1" or reader.outline_level(p.element) == 0

        if not (page_ok and font_ok and size_ok and bold_ok and top_align_ok and gap_ok and not_orphan_ok and style_ok):
            fail.append(
                f"{p.text[:10]!r}: 页{page_no} 字体{reader.run_fonts(p.element)} 字号{reader.run_size_pt(p.element)} "
                f"加粗{bold_ok} 顶格{top_align_ok} 段前后{before}/{after} 后续{has_body} 样式{p.style_id}"
            )

    ok = not fail
    return CheckResult(ok, f"5个一级标题；不合格 {len(fail)}；{'；'.join(fail) or '全部合格'}")


def check_d2_17_reference_heading(reader: DocxReader) -> CheckResult:
    p = find_body_paragraph(reader, "六、参考资料")
    if not p:
        return CheckResult(False, "未找到精确标题“六、参考资料”")
    # 细则点 1：黑体
    font_ok = paragraph_has_font(reader, p, "黑体")
    # 细则点 2：小四号(12pt)
    size_ok = paragraph_size_is(reader, p, FONT_PT["小四"], 0.1)
    # 细则点 3：加粗
    bold_ok = reader.run_bold(p.element)
    # 细则点 4：居中
    center_ok = is_center(reader, p)
    ok = font_ok and size_ok and bold_ok and center_ok
    return CheckResult(
        ok,
        f"字体{reader.run_fonts(p.element)} 字号{reader.run_size_pt(p.element)} 加粗{bold_ok} 居中{center_ok}",
    )


def check_d2_19_body_paragraph_format(reader: DocxReader) -> CheckResult:
    paras = ordinary_body_paragraphs(reader)
    if not paras:
        return CheckResult(False, "未识别到正文普通段落")

    fail: list[str] = []
    for p in paras:
        # 细则点 1：中文宋体，英文/数字 Times New Roman
        fonts = reader.run_fonts(p.element)
        font_ok = "宋体" in fonts.get("eastAsia", "") and (
            "Times New Roman" in fonts.get("ascii", "")
            or "Times New Roman" in fonts.get("hAnsi", "")
        )
        # 细则点 2：小四号(12pt)
        size_ok = paragraph_size_is(reader, p, FONT_PT["小四"], 0.1)
        # 细则点 3：颜色为黑色
        color = reader.run_color(p.element)
        color_ok = color in {"", "000000", "auto"}
        # 细则点 4：首行缩进 2 个字符（小四 12pt 字符宽约 0.42cm，2 字符约 0.85cm）
        first_line = paragraph_first_line_cm(reader, p)
        # docx 里"首行缩进2字符"按字符存储时 firstLineChars=200；按 twips 时常为 0.74cm (21pt)
        first_line_chars = None
        ind = p.element.find("w:pPr/w:ind", NS)
        if ind is not None:
            fl_chars = attr(ind, "w:firstLineChars")
            if fl_chars:
                first_line_chars = to_int(fl_chars)
        indent_ok = first_line_chars == 200 or approx(first_line, 0.85, 0.03)
        # 细则点 5：固定值 20 磅行距
        sp = reader.paragraph_spacing(p.element)
        line_twips = to_int(sp.get("line"))
        line_pt = line_twips / 20.0 if line_twips is not None else None
        line_ok = sp.get("lineRule") == "exact" and approx(line_pt, 20, 0.1)
        # 细则点 6-7：段前 0、段后 0
        before, after = paragraph_before_after_pt(reader, p)
        before_ok = approx(before, 0, 0.1)
        after_ok = approx(after, 0, 0.1)
        # 细则点 8：两端对齐
        align = reader.paragraph_jc(p.element)
        align_ok = align in {"both", "distribute"}

        if not (font_ok and size_ok and color_ok and indent_ok and line_ok and before_ok and after_ok and align_ok):
            fail.append(
                f"{p.text[:15]!r}: 字体{fonts} 字号{reader.run_size_pt(p.element)} 色{color or 'auto'} "
                f"缩进{first_line}cm(chars={first_line_chars}) 行距{line_pt}({sp.get('lineRule')}) "
                f"段前后{before}/{after} 对齐{align!r}"
            )

    ok = not fail
    return CheckResult(ok, f"正文段落 {len(paras)}；不合格 {len(fail)}；{'；'.join(fail[:2]) or '全部合格'}")


def check_d2_20_images(reader: DocxReader) -> CheckResult:
    images = reader.image_references()
    # 细则点 1：图1至图7 —— 图片数量为 7
    if len(images) != 7:
        return CheckResult(False, f"图片数量 {len(images)}≠7")

    # 细则点 2：分布在第4-9页（"第4页至第9页图片"）
    # 用 COM InlineShapes 拿"图片自身"的视觉页码，而不是图片所在段落首字的页码——
    # 段落首字可能在上一页底，图片视觉上已翻到下一页。
    image_pages = reader.real_image_pages()
    off_page: list[tuple[int, "int | None"]] = []
    if len(image_pages) == 7:
        for idx, pg in enumerate(image_pages, 1):
            if not (4 <= pg <= 9):
                off_page.append((idx, pg))
    else:
        # COM 未取到 InlineShapes，则退到段落页码
        for idx, img in enumerate(images, 1):
            pinfo = next((p for p in reader.paragraphs if p.index == img["paragraph_index"]), None)
            if pinfo is None:
                off_page.append((idx, None))
                continue
            page_no = _real_page_or_xml(reader, pinfo.element)
            if not (4 <= page_no <= 9):
                off_page.append((idx, page_no))
    if off_page:
        return CheckResult(False, f"图片未全部在第4-9页：{off_page}")

    # 细则点 3：水平居中
    centered = [img for img in images if img["centered"]]

    def _as_float(v: object) -> "float | None":
        return float(v) if isinstance(v, (int, float)) else None

    # 细则点 4：宽度约 10—13 厘米
    width_ok = [img for img in images if in_range(_as_float(img["width_cm"]), 10, 13)]
    # 细则点 5：高度约 5—10 厘米
    height_ok = [img for img in images if in_range(_as_float(img["height_cm"]), 5, 10)]

    ok = len(centered) == 7 and len(width_ok) == 7 and len(height_ok) == 7
    dims: list[str] = []
    for img in images:
        w = _as_float(img["width_cm"])
        h = _as_float(img["height_cm"])
        if w is not None and h is not None:
            dims.append(f"{w:.1f}×{h:.1f}")
    return CheckResult(
        ok,
        f"图片 {len(images)}；视觉页{image_pages}；居中 {len(centered)}/7；宽合格 {len(width_ok)}/7；高合格 {len(height_ok)}/7；尺寸 {dims}",
    )


def check_d2_21_image_captions(reader: DocxReader) -> CheckResult:
    # 图题：段落文本以 "图 N " 起头
    captions = [
        p for p in reader.paragraphs
        if p.text.strip().startswith("图") and re.match(r"^图\s*[1-9]", p.text.strip())
    ]
    cap_norms = [normalize_caption(p.text) for p in captions]
    required_norms = [normalize_caption(c) for c in REQUIRED_CAPTIONS]
    # 细则点 1：依次显示图1至图7（顺序与内容一致）
    content_ok = all(c in cap_norms for c in required_norms)
    order_ok = cap_norms[:7] == required_norms if len(cap_norms) >= 7 else False

    # 细则点 2：黑体五号水平居中
    fmt_ok = 0
    for p in captions[:7]:
        if paragraph_has_font(reader, p, "黑体") and paragraph_size_is(reader, p, FONT_PT["五号"], 0.1) and is_center(reader, p):
            fmt_ok += 1

    # 细则点 3：图题位于"对应图片"下方（正文段落顺序上紧随图片）且视觉上同页
    #
    # 严格配对方法：
    #   - 用 image_references() 拿到每张图片所在的正文段落 index（按图片出现顺序）；
    #   - 对第 i 张图片（1..7）：从其所在段落之后的下一个"含图片段落"之前，
    #     期望首个符合 r'^图\s*i\b' 的段落作为其图题；期间不得出现其他图片；
    #   - 图题页码 == 图片视觉页码（COM 优先，退到 XML）。
    images = reader.image_references()
    # image_paragraphs[i] = 第 i+1 张图片所在的段落 index（body-level）
    image_paragraphs: list[int] = []
    for img in images:
        raw = img.get("paragraph_index")
        if isinstance(raw, int):
            image_paragraphs.append(raw)
    image_pages = reader.real_image_pages()

    para_by_index: dict[int, ParagraphInfo] = {p.index: p for p in reader.paragraphs}

    pair_ok = 0
    same_page = 0
    detail: list[str] = []
    for i in range(min(7, len(image_paragraphs))):
        img_para_idx = image_paragraphs[i]
        next_img_idx = image_paragraphs[i + 1] if i + 1 < len(image_paragraphs) else None

        # 在 (img_para_idx, next_img_idx) 范围内寻找匹配 "图{i+1}" 的首个段落
        target_prefix = re.compile(rf"^图\s*{i + 1}(?![0-9])")
        found_cap: "ParagraphInfo | None" = None
        j = img_para_idx + 1
        while j in para_by_index and (next_img_idx is None or j < next_img_idx):
            q = para_by_index[j]
            if target_prefix.match(q.text.strip()):
                found_cap = q
                break
            j += 1

        if found_cap is None:
            detail.append(f"图{i+1}:未在其后紧随找到图题")
            continue

        pair_ok += 1
        cap_page = _real_page_or_xml(reader, found_cap.element)
        img_page = image_pages[i] if i < len(image_pages) else None
        if img_page is not None and img_page == cap_page:
            same_page += 1
        else:
            detail.append(f"图{i+1}:图页{img_page}/题页{cap_page}")

    ok = (
        content_ok
        and order_ok
        and fmt_ok == 7
        and pair_ok == 7
        and same_page == 7
        and len(image_paragraphs) == 7
    )
    return CheckResult(
        ok,
        (
            f"图题 {len(captions)}；内容完整 {content_ok}；顺序 {order_ok}；"
            f"格式 {fmt_ok}/7；紧随配对 {pair_ok}/7；同页 {same_page}/7；"
            f"图片段索引 {image_paragraphs}；{detail or '全部同页且紧随'}"
        ),
    )


def check_d2_22_table_title(reader: DocxReader) -> CheckResult:
    p = find_body_paragraph(reader, "表1 主要模块及设计要点")
    if p is None:
        # 允许"表1"和题名之间用全角空格
        p = next(
            (
                q for q in reader.paragraphs
                if normalize_text("表1主要模块及设计要点") in normalize_text(q.text)
            ),
            None,
        )
    if p is None:
        actual = [q.text for q in reader.paragraphs if q.text.strip().startswith("表")]
        return CheckResult(False, f"未找到“表1 主要模块及设计要点”；当前表题 {actual}")

    # 必须在第5页
    page_no = _real_page_or_xml(reader, p.element)
    if page_no != 5:
        return CheckResult(False, f"表题未出现在第5页（实际第{page_no}页）")

    # 表题必须位于表格上方：紧随其后的下一个表格存在
    p_pos = next((i for i, b in enumerate(reader.body_items) if b.kind == "p" and b.index == p.index), -1)
    next_table = any(b.kind == "tbl" for b in reader.body_items[p_pos + 1: p_pos + 4])

    # 字体黑体五号水平居中
    font_ok = paragraph_has_font(reader, p, "黑体")
    size_ok = paragraph_size_is(reader, p, FONT_PT["五号"], 0.1)
    center_ok = is_center(reader, p)

    ok = next_table and font_ok and size_ok and center_ok
    return CheckResult(
        ok,
        f"表上方 {next_table}；字体{reader.run_fonts(p.element)} 字号{reader.run_size_pt(p.element)} 居中{center_ok}",
    )


def check_d2_23_module_table(reader: DocxReader) -> CheckResult:
    # 细则点 1：3 列 5 行可编辑 Word 表格
    target: Optional[TableInfo] = None
    for tbl in reader.tables:
        if len(tbl.rows) == 5 and all(len(r) == 3 for r in tbl.rows):
            flat = normalize_text("".join("".join(r) for r in tbl.rows))
            if all(k in flat for k in ["模块", "主要功能", "设计要点", "承载箱体", "副车架", "悬挂轮组", "控制面板"]):
                target = tbl
                break
    if target is None:
        return CheckResult(
            False,
            f"未找到 3列5行且包含指定记录的模块表；当前表格行数 {[len(t.rows) for t in reader.tables]}",
        )

    # 细则点 2：在第5页
    first_p = target.cell_paragraphs[0] if target.cell_paragraphs else None
    if first_p is None:
        return CheckResult(False, "表格无段落，无法定位页码")
    page_no = _real_page_or_xml(reader, first_p.element)
    if page_no != 5:
        return CheckResult(False, f"模块表未出现在第5页（实际第{page_no}页）")

    # 细则点 3：表头依次为"模块""主要功能""设计要点"
    header_ok = target.rows[0] == ["模块", "主要功能", "设计要点"]

    # 细则点 4：下方包含承载箱体、副车架、悬挂轮组、控制面板 4 条记录（首列依次匹配）
    expected_first_col = ["承载箱体", "副车架", "悬挂轮组", "控制面板"]
    actual_first_col = [row[0] for row in target.rows[1:]]
    records_ok = actual_first_col == expected_first_col

    # 细则点 5：外框和内部横竖线连续完整（top/left/bottom/right/insideH/insideV 全部非 nil）
    border_ok = module_table_borders_complete(reader, target.element)

    # 细则点 6+7：表格内部宋体五号；英文/数字 Times New Roman 五号
    # 细则点 8：表头加粗
    font_fail: list[str] = []
    for p in target.cell_paragraphs:
        if not p.text.strip():
            continue
        fonts = reader.run_fonts(p.element)
        cn_ok = "宋体" in fonts.get("eastAsia", "")
        has_en = bool(re.search(r"[A-Za-z0-9]", p.text))
        en_ok = (
            "Times New Roman" in fonts.get("ascii", "")
            or "Times New Roman" in fonts.get("hAnsi", "")
        )
        font_ok = cn_ok and (en_ok or not has_en)
        size_ok = paragraph_size_is(reader, p, FONT_PT["五号"], 0.1)
        if not (font_ok and size_ok):
            font_fail.append(f"{p.text[:10]!r}:字体{fonts} 字号{reader.run_size_pt(p.element)}")

    header_bold = all(reader.run_bold(p.element) for p in target.cell_paragraphs[:3])

    # 细则点 9：文字垂直居中
    vcenter_cells = 0
    total_cells = 0
    for tc in target.element.findall(".//w:tc", NS):
        total_cells += 1
        vcenter_cells += int(cell_vertical_center(tc))
    vcenter_ok = vcenter_cells == total_cells

    # 细则点 10：长文字自动换行，不超出单元格——单元格未禁止换行（w:noWrap 未设）
    no_wrap_cells = 0
    for tc in target.element.findall(".//w:tc", NS):
        if tc.find("w:tcPr/w:noWrap", NS) is not None:
            no_wrap_cells += 1
    wrap_ok = no_wrap_cells == 0

    ok = (
        header_ok
        and records_ok
        and border_ok
        and not font_fail
        and header_bold
        and vcenter_ok
        and wrap_ok
    )
    return CheckResult(
        ok,
        (
            f"表头 {header_ok}；4条记录 {records_ok}（{actual_first_col}）；"
            + f"边框连续 {border_ok}；表头加粗 {header_bold}；"
            + f"字体不合格 {len(font_fail)}；垂直居中 {vcenter_cells}/{total_cells}；"
            + f"禁换行单元格 {no_wrap_cells}"
        ),
    )


def reference_paragraphs(reader: DocxReader) -> list[ParagraphInfo]:
    start = next((p.index for p in reader.paragraphs if "参考资料" in p.text and (p.style_id == "Heading1" or re.match(r"^\d+\s+参考资料", p.text.strip()) or "六、" in p.text)), -1)
    if start < 0:
        return []
    return [p for p in reader.paragraphs if p.index > start and p.text.strip().startswith("[")]


def check_d2_24_references_body(reader: DocxReader) -> CheckResult:
    refs = reference_paragraphs(reader)
    if len(refs) != 4:
        return CheckResult(False, f"参考资料条目 {len(refs)}≠4")

    # 必须在第10页
    off_page = [(p.text[:15], _real_page_or_xml(reader, p.element)) for p in refs if _real_page_or_xml(reader, p.element) != 10]
    if off_page:
        return CheckResult(False, f"参考资料未全部在第10页：{off_page}")

    fail: list[str] = []
    for p in refs:
        # 细则点 1：中文宋体五号，英文/数字 Times New Roman 五号
        fonts = reader.run_fonts(p.element)
        font_ok = "宋体" in fonts.get("eastAsia", "") and (
            "Times New Roman" in fonts.get("ascii", "")
            or "Times New Roman" in fonts.get("hAnsi", "")
        )
        size_ok = paragraph_size_is(reader, p, FONT_PT["五号"], 0.1)
        # 细则点 2：顶格排列（左缩进与首行缩进都为 0）
        left = paragraph_left_cm(reader, p) or 0.0
        first = paragraph_first_line_cm(reader, p) or 0.0
        top_ok = abs(left) < 0.03 and abs(first) < 0.03

        if not (font_ok and size_ok and top_ok):
            fail.append(
                f"{p.text[:15]!r}: 字体{fonts} 字号{reader.run_size_pt(p.element)} 顶格{top_ok}"
            )

    ok = not fail
    return CheckResult(ok, f"参考资料 4 条；不合格 {len(fail)}；{'；'.join(fail) or '全部合格'}")


def check_d2_25_reference_numbers(reader: DocxReader) -> CheckResult:
    refs = reference_paragraphs(reader)
    if len(refs) != 4:
        return CheckResult(False, f"参考资料条目 {len(refs)}≠4")

    # 必须在第10页
    off_page = [(p.text[:15], _real_page_or_xml(reader, p.element)) for p in refs if _real_page_or_xml(reader, p.element) != 10]
    if off_page:
        return CheckResult(False, f"参考资料未全部在第10页：{off_page}")

    # 细则点 1：依次使用 [1] [2] [3] [4]
    starts = [p.text.strip()[:3] for p in refs]
    expected = ["[1]", "[2]", "[3]", "[4]"]
    seq_ok = starts == expected
    # 细则点 2：不使用项目符号代替（无 w:numPr）
    bullet_used = any(p.element.find("w:pPr/w:numPr", NS) is not None for p in refs)
    ok = seq_ok and not bullet_used
    return CheckResult(ok, f"编号 {starts}；连续 {seq_ok}；项目符号 {bullet_used}")


def ordinary_body_paragraphs(reader: DocxReader) -> list[ParagraphInfo]:
    result: list[ParagraphInfo] = []
    in_body = False
    for p in reader.paragraphs:
        text = p.text.strip()
        if not text:
            continue
        if "设计思路" in text and (p.style_id == "Heading1" or reader.outline_level(p.element) == 0):
            in_body = True
            continue
        if "参考资料" in text and (p.style_id == "Heading1" or reader.outline_level(p.element) == 0):
            break
        if not in_body:
            continue
        if p.style_id in {"Heading1", "CaptionCN", "TocEntryCN", "TocTitleCN"}:
            continue
        if text.startswith("图") or text.startswith("表"):
            continue
        result.append(p)
    return result


RUBRIC_DIM2: list[ScoreItem] = [
    ScoreItem("D2-01", 5, "整篇文档页面设置与对象边界", check_d2_01_page_setup),
    ScoreItem("D2-02", 3, "第1页封面内容与字体", check_d2_02_cover_content_font),
    ScoreItem("D2-03", 1, "封面字段段落格式", check_d2_03_cover_field_paragraphs),
    ScoreItem("D2-04", 1, "第2页独创性声明标题", check_d2_04_originality_title),
    ScoreItem("D2-05", 1, "第2页独创性声明正文", check_d2_05_originality_body),
    ScoreItem("D2-06", 1, "第2页签字区域", check_d2_06_signature_area),
    ScoreItem("D2-07", 1, "第3页目录标题", check_d2_07_toc_title),
    ScoreItem("D2-08", 1, "第3页自动目录对象", check_d2_08_auto_toc_object),
    ScoreItem("D2-09", 5, "第3页目录跳转功能", check_d2_09_toc_hyperlinks),
    ScoreItem("D2-10", 1, "第3页目录正文字体与引导符", check_d2_10_toc_entry_format),
    ScoreItem("D2-11", 1, "第3页目录页码列", check_d2_11_toc_page_column),
    ScoreItem("D2-13", 1, "第3页后下一页分节符", check_d2_13_section_break_after_toc),
    ScoreItem("D2-14", 3, "页码对应位置", check_d2_14_page_number_position),
    ScoreItem("D2-15", 1, "第4页至第10页自动页码", check_d2_15_footer_page_numbers),
    ScoreItem("D2-16", 1, "正文一级标题", check_d2_16_primary_headings),
    ScoreItem("D2-17", 1, "“六、参考资料”标题", check_d2_17_reference_heading),
    ScoreItem("D2-19", 3, "正文段落格式", check_d2_19_body_paragraph_format),
    ScoreItem("D2-20", 5, "第4页至第9页图片", check_d2_20_images),
    ScoreItem("D2-21", 5, "图1至图7图题", check_d2_21_image_captions),
    ScoreItem("D2-22", 1, "第5页模块表表题", check_d2_22_table_title),
    ScoreItem("D2-23", 5, "第5页模块表", check_d2_23_module_table),
    ScoreItem("D2-24", 1, "第10页参考资料正文", check_d2_24_references_body),
    ScoreItem("D2-25", 1, "第10页参考资料编号", check_d2_25_reference_numbers),
]


def evaluate_dimension2(reader: DocxReader) -> tuple[int, list[tuple[ScoreItem, CheckResult]], list[tuple[ScoreItem, CheckResult]]]:
    total = 0
    hits: list[tuple[ScoreItem, CheckResult]] = []
    misses: list[tuple[ScoreItem, CheckResult]] = []
    for item in RUBRIC_DIM2:
        try:
            result = item.check(reader)
        except Exception as exc:
            result = CheckResult(False, f"检查函数异常：{exc.__class__.__name__}: {exc}")
        if result.passed:
            total += item.score
            hits.append((item, result))
        else:
            misses.append((item, result))
    return total, hits, misses


def _find_docx_in_dir(dir_path: str) -> Path:
    """在给定目录内查找待评估的 .docx 文件。

    目录内只放一份被评估文档，因此直接取第一个非临时文件（跳过
    Word 打开时产生的 ~$ 锁文件）；找不到则抛错，由 evaluate 捕获为
    status="error"。
    """
    base = Path(dir_path)
    candidates = sorted(
        p for p in base.glob("*.docx") if not p.name.startswith("~$")
    )
    if not candidates:
        raise FileNotFoundError(f"目录中未找到 .docx 文件：{base}")
    return candidates[0]


def evaluate(dir_path: str) -> dict:
    """按 §2.2 约定，返回结构化 dict；不 print、不改 sys.stdout、不 sys.exit。

    dir_path 为脚本所在目录的路径；脚本自己负责在该目录里定位并打开
    被评估的文档。脚本自身崩溃（含文件不存在等）由本函数捕获并返回
    status="error"。
    """
    result: dict = {
        "id": SCRIPT_ID,
        "file_name": "",
        "status": "ok",
        "error": None,
        "dim1_pass": True,
        "dim1_reason": "",
        "dim2_items": [],
        "total_score": 0,
        "max_score": MAX_SCORE_TOTAL,
    }
    try:
        docx_path = _find_docx_in_dir(dir_path)
        result["file_name"] = docx_path.name
        return _evaluate_impl(docx_path, result)
    except Exception as exc:
        result["status"] = "error"
        result["error"] = f"{type(exc).__name__}: {exc}"
        result["dim2_items"] = []
        result["total_score"] = 0
        return result


def _evaluate_impl(docx_path: Path, result: dict) -> dict:
    reader = DocxReader(docx_path)

    d1_passed, d1_details = check_dimension1(reader)
    result["dim1_pass"] = d1_passed
    if not d1_passed:
        fail_reasons = [message for ok, message in d1_details if not ok]
        result["dim1_reason"] = "；".join(fail_reasons)
        result["dim2_items"] = []
        result["total_score"] = 0
        return result

    total, hits, misses = evaluate_dimension2(reader)
    dim2_items = []
    for item, _check_result in hits:
        dim2_items.append({
            "rule": item.name,
            "max_delta": item.score,
            "delta": item.score,
            "hit": True,
            "detail": "",
        })
    for item, _check_result in misses:
        dim2_items.append({
            "rule": item.name,
            "max_delta": item.score,
            "delta": 0,
            "hit": False,
            "detail": "",
        })
    result["dim2_items"] = dim2_items
    result["total_score"] = total
    return result


if __name__ == "__main__":
    _dir_path = sys.argv[1] if len(sys.argv) > 1 else "."
    print(json.dumps(evaluate(_dir_path), ensure_ascii=False, indent=2))
