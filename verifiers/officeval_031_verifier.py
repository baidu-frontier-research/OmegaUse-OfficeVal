#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""按给定细则自动评估 Word 试卷 DOCX 文件。

统一入口：
    def evaluate(dir_path: str) -> dict
调用方传入"脚本所在目录的路径"，脚本自己在该目录下定位并打开被评估的 docx。
返回结构见《脚本接口差异与统一建议.md》§2.2。

说明：
- 只依赖 Python 标准库。
- 直接解析 DOCX 内部 OOXML，尽量自动判断版式、字体、表格、边框、页眉页脚等。
- 只识别 .docx；页数按 OOXML 的显式分页符与元数据静态推断，不使用 Word COM/LibreOffice。
"""

from __future__ import annotations

import json
import re
import sys
import zipfile
from dataclasses import dataclass, field
from pathlib import Path
from typing import Iterable, NamedTuple, Optional
import xml.etree.ElementTree as ET

SCRIPT_ID = "031"

NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "pic": "http://schemas.openxmlformats.org/drawingml/2006/picture",
    "wps": "http://schemas.microsoft.com/office/word/2010/wordprocessingShape",
    "v": "urn:schemas-microsoft-com:vml",
    "vt": "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes",
}

W = NS["w"]
R = NS["r"]
WP = NS["wp"]
A = NS["a"]

TWIPS_PER_CM = 1440 / 2.54
EMU_PER_CM = 914400 / 2.54

ANSWER_RE = re.compile(r"参考答案|答案|评分标准|解析")
PINYIN_RE = re.compile(r"[āáǎàēéěèīíǐìōóǒòūúǔùǖǘǚǜü]|\b[a-zü]{2,}(?:\s+[a-zü]{2,}){2,}", re.I)
CHINESE_RE = re.compile(r"[一-鿿]")
MAJOR_TITLE_RE = re.compile(r"[一二三四五六七八九]、")

EXPECTED_MARGINS_CM = {"top": 1.0, "bottom": 0.8, "left": 1.3, "right": 1.3}
EXPECTED_MODULES = ["一、", "二、", "三、", "四、", "五、", "六、", "七、", "八、", "九、"]


def qn(ns: str, tag: str) -> str:
    return f"{{{NS[ns]}}}{tag}"


def wattr(el: Optional[ET.Element], name: str, default: Optional[str] = None) -> Optional[str]:
    if el is None:
        return default
    return el.get(qn("w", name), default)


def local_name(tag: str) -> str:
    return tag.rsplit("}", 1)[-1]


def to_int(value: Optional[str], default: int = 0) -> int:
    try:
        return int(value) if value is not None else default
    except ValueError:
        return default


def twips_to_cm(value: Optional[str]) -> Optional[float]:
    if value is None:
        return None
    return to_int(value) / TWIPS_PER_CM


def emu_to_cm(value: Optional[str]) -> Optional[float]:
    if value is None:
        return None
    return to_int(value) / EMU_PER_CM


def element_text(el: Optional[ET.Element]) -> str:
    if el is None:
        return ""
    return "".join(t.text or "" for t in el.findall(f".//{{{W}}}t"))


def has_page_break(el: ET.Element) -> bool:
    for br in el.findall(f".//{{{W}}}br"):
        if br.get(qn("w", "type")) == "page":
            return True
    return False


def rgb(hex_color: Optional[str]) -> Optional[tuple[int, int, int]]:
    if not hex_color:
        return None
    s = hex_color.strip().replace("#", "").upper()
    if s in {"AUTO", "NONE"} or len(s) != 6:
        return None
    try:
        return int(s[0:2], 16), int(s[2:4], 16), int(s[4:6], 16)
    except ValueError:
        return None


def is_near(value: Optional[float], expected: float, tolerance: float) -> bool:
    return value is not None and abs(value - expected) <= tolerance


def is_light_blue(hex_color: Optional[str]) -> bool:
    c = rgb(hex_color)
    if not c:
        return False
    r, g, b = c
    # 极浅蓝常见为 EAF4FC / DDEAF7；允许“蓝通道最高且整体很亮”。
    return b >= g >= r and b >= 210 and g >= 200


def is_light_green(hex_color: Optional[str]) -> bool:
    c = rgb(hex_color)
    if not c:
        return False
    r, g, b = c
    # 极浅绿如 F1F8E9，绿色只比红色略高，因此采用亮度+相对优势判断。
    return g >= r and g >= b and g >= 220 and (g - b) >= 5


def is_light_orange(hex_color: Optional[str]) -> bool:
    c = rgb(hex_color)
    if not c:
        return False
    r, g, b = c
    # 极浅橙如 FFF7EF：R最高、G次之、B最低，整体很亮。
    return r >= 240 and g >= 210 and b <= 245 and r >= g >= b and (g - b) >= 5


def is_light_yellow(hex_color: Optional[str]) -> bool:
    c = rgb(hex_color)
    if not c:
        return False
    r, g, b = c
    # 浅黄如 FFFDEB：R/G都很高，B略低。
    return r >= 240 and g >= 235 and b <= 245 and abs(r - g) <= 18


def is_blue(hex_color: Optional[str]) -> bool:
    c = rgb(hex_color)
    if not c:
        return False
    r, g, b = c
    # 标准蓝：蓝通道明显高于红绿；不至于过亮（与浅蓝区分）。
    return b > r and b > g and b >= 130 and (b - max(r, g)) >= 30 and max(r, g) <= 200


def is_green(hex_color: Optional[str]) -> bool:
    c = rgb(hex_color)
    if not c:
        return False
    r, g, b = c
    # 标准绿：绿通道明显高于红蓝；不至于过亮。
    return g > r and g > b and g >= 110 and (g - max(r, b)) >= 25 and max(r, b) <= 200


def is_orange(hex_color: Optional[str]) -> bool:
    c = rgb(hex_color)
    if not c:
        return False
    r, g, b = c
    # 标准橙：红 > 绿 > 蓝，红色突出，蓝色明显偏低。
    return r > g > b and r >= 180 and g >= 80 and b <= 120 and (r - b) >= 80


def is_dark_orange(hex_color: Optional[str]) -> bool:
    c = rgb(hex_color)
    if not c:
        return False
    r, g, b = c
    # 深橙：红通道最高，绿其次，蓝最低，且红色不至过亮（区别于浅橙）。
    return r > g > b and 150 <= r <= 230 and 60 <= g <= 170 and b <= 120


def is_dark_blue(hex_color: Optional[str]) -> bool:
    c = rgb(hex_color)
    if not c:
        return False
    r, g, b = c
    # 深蓝：蓝通道明显高于红绿且整体偏暗。
    return b > r and b > g and r <= 120 and g <= 140 and b >= 80


def is_gray(hex_color: Optional[str]) -> bool:
    c = rgb(hex_color)
    if not c:
        return False
    r, g, b = c
    # 灰色：三通道接近，且不属于纯黑。
    return max(r, g, b) - min(r, g, b) <= 20 and 60 <= max(r, g, b) <= 200


def is_black(hex_color: Optional[str]) -> bool:
    c = rgb(hex_color)
    if not c:
        return False
    r, g, b = c
    return r <= 40 and g <= 40 and b <= 40


def is_gray_blue(hex_color: Optional[str]) -> bool:
    c = rgb(hex_color)
    if not c:
        return False
    r, g, b = c
    return 80 <= r <= 160 and 80 <= g <= 170 and b >= r


def color_name(hex_color: Optional[str]) -> str:
    if is_light_blue(hex_color):
        return "blue"
    if is_light_green(hex_color):
        return "green"
    if is_light_orange(hex_color):
        return "orange"
    if is_light_yellow(hex_color):
        return "yellow"
    return "other"


@dataclass
class CheckResult:
    id: str
    name: str
    points: int
    score: int
    passed: bool
    evidence: list[str] = field(default_factory=list)

    @property
    def symbol(self) -> str:
        return "✓" if self.passed else "✗"


@dataclass
class TableInfo:
    index: int
    el: ET.Element
    text: str
    rows: int
    cells_per_row: list[int]
    all_cells: int
    fills: list[str]
    border_colors: list[str]

    @property
    def first_fill(self) -> Optional[str]:
        return self.fills[0] if self.fills else None


class DocxInspector:
    def __init__(self, path: Path):
        self.path = path
        self.zf: Optional[zipfile.ZipFile] = None
        self.names: list[str] = []
        self.document: Optional[ET.Element] = None
        self.styles: Optional[ET.Element] = None
        self.app: Optional[ET.Element] = None
        self.headers: list[ET.Element] = []
        self.footers: list[ET.Element] = []
        self.open_error: Optional[str] = None

    def open(self) -> bool:
        try:
            self.zf = zipfile.ZipFile(self.path)
            self.names = self.zf.namelist()
            required = ["[Content_Types].xml", "word/document.xml"]
            missing = [name for name in required if name not in self.names]
            if missing:
                self.open_error = "缺少DOCX必要组件：" + ", ".join(missing)
                return False
            self.document = ET.fromstring(self.zf.read("word/document.xml"))
            if "word/styles.xml" in self.names:
                self.styles = ET.fromstring(self.zf.read("word/styles.xml"))
            if "docProps/app.xml" in self.names:
                self.app = ET.fromstring(self.zf.read("docProps/app.xml"))
            for name in sorted(n for n in self.names if n.startswith("word/header") and n.endswith(".xml")):
                self.headers.append(ET.fromstring(self.zf.read(name)))
            for name in sorted(n for n in self.names if n.startswith("word/footer") and n.endswith(".xml")):
                self.footers.append(ET.fromstring(self.zf.read(name)))
            return True
        except Exception as exc:  # noqa: BLE001 - 作为评分工具，需把异常转成人可读失败原因
            self.open_error = f"无法打开或解析DOCX：{exc}"
            return False

    def text(self) -> str:
        return element_text(self.document)

    def header_text(self) -> str:
        return "\n".join(element_text(h) for h in self.headers)

    def footer_text(self) -> str:
        return "\n".join(element_text(f) for f in self.footers)

    def footer_instr_text(self) -> str:
        chunks: list[str] = []
        for footer in self.footers:
            chunks.extend(e.text or "" for e in footer.findall(f".//{{{W}}}instrText"))
        return " ".join(chunks)

    def media_names(self) -> list[str]:
        return [n for n in self.names if n.startswith("word/media/")]

    def paragraphs(self) -> list[ET.Element]:
        return self.document.findall(f".//{{{W}}}p") if self.document is not None else []

    def tables(self) -> list[TableInfo]:
        if self.document is None:
            return []
        result: list[TableInfo] = []
        for idx, tbl in enumerate(self.document.findall(f".//{{{W}}}tbl")):
            rows = tbl.findall(f"{{{W}}}tr")
            cells_per_row = [len(row.findall(f"{{{W}}}tc")) for row in rows]
            fills = []
            for shd in tbl.findall(f".//{{{W}}}shd"):
                fill = wattr(shd, "fill")
                if fill and fill.upper() not in {"AUTO", "FFFFFF"}:
                    fills.append(fill.upper())
            border_colors = []
            for bdr in tbl.findall(f".//{{{W}}}tblBorders/*") + tbl.findall(f".//{{{W}}}tcBorders/*"):
                val = wattr(bdr, "val")
                color = wattr(bdr, "color")
                if val and val != "nil" and color:
                    border_colors.append(color.upper())
            result.append(
                TableInfo(
                    index=idx,
                    el=tbl,
                    text=element_text(tbl),
                    rows=len(rows),
                    cells_per_row=cells_per_row,
                    all_cells=sum(cells_per_row),
                    fills=fills,
                    border_colors=border_colors,
                )
            )
        return result

    def section_properties(self) -> list[ET.Element]:
        if self.document is None:
            return []
        return self.document.findall(f".//{{{W}}}sectPr")

    def explicit_page_break_count(self) -> int:
        if self.document is None:
            return 0
        return sum(1 for br in self.document.findall(f".//{{{W}}}br") if br.get(qn("w", "type")) == "page")

    def app_page_count(self) -> Optional[int]:
        if self.app is None:
            return None
        for child in self.app:
            if local_name(child.tag) == "Pages" and child.text:
                return to_int(child.text, 0) or None
        return None

    def page_segments_by_breaks(self) -> list[str]:
        """用显式分页符粗略切分正文，供自动评分兜底使用。"""
        if self.document is None:
            return []
        body = self.document.find(f"{{{W}}}body")
        if body is None:
            return [self.text()]
        pages: list[str] = []
        buf: list[str] = []
        for child in list(body):
            if local_name(child.tag) == "sectPr":
                continue
            buf.append(element_text(child))
            if has_page_break(child):
                pages.append("".join(buf))
                buf = []
        pages.append("".join(buf))
        return pages

    def drawing_extents_cm(self) -> list[tuple[float, float]]:
        if self.document is None:
            return []
        extents = []
        for ext in self.document.findall(f".//{{{WP}}}extent"):
            cx = emu_to_cm(ext.get("cx"))
            cy = emu_to_cm(ext.get("cy"))
            if cx and cy:
                extents.append((cx, cy))
        return extents

    def run_font_values(self, root: Optional[ET.Element] = None) -> list[str]:
        root = root if root is not None else self.document
        if root is None:
            return []
        fonts = []
        for rfonts in root.findall(f".//{{{W}}}rFonts"):
            for key in ("ascii", "hAnsi", "eastAsia", "cs"):
                val = wattr(rfonts, key)
                if val:
                    fonts.append(val)
        return fonts

    def run_sizes_half_points(self, root: Optional[ET.Element] = None) -> list[int]:
        root = root if root is not None else self.document
        if root is None:
            return []
        sizes = []
        for tag in ("sz", "szCs"):
            for el in root.findall(f".//{{{W}}}{tag}"):
                val = to_int(wattr(el, "val"), 0)
                if val:
                    sizes.append(val)
        return sizes

    def run_colors(self, root: Optional[ET.Element] = None) -> list[str]:
        root = root if root is not None else self.document
        if root is None:
            return []
        colors = []
        for color in root.findall(f".//{{{W}}}color"):
            val = wattr(color, "val")
            if val:
                colors.append(val.upper())
        return colors

    def centered_paragraph_count(self, root: Optional[ET.Element] = None) -> int:
        root = root if root is not None else self.document
        if root is None:
            return 0
        count = 0
        for jc in root.findall(f".//{{{W}}}jc"):
            if wattr(jc, "val") in {"center", "both"}:
                count += 1
        return count


def get_page_count(inspector: DocxInspector) -> tuple[Optional[int], str, list[str]]:
    """静态推断页数：仅依据 OOXML 的显式分页符与 docProps 元数据，不使用 Word COM。"""
    evidence: list[str] = []

    app_pages = inspector.app_page_count()
    explicit = inspector.explicit_page_break_count()
    segments = inspector.page_segments_by_breaks()
    text = inspector.text()
    ans_match = ANSWER_RE.search(text)

    if explicit == 4:
        return 5, "显式分页符", ["检测到4个显式分页符，推断5页"]
    if explicit == 3 and ans_match and ans_match.start() > int(len(text) * 0.75):
        return 5, "显式分页符+答案页兜底", ["检测到3个显式分页符且答案段位于文末，按4页试题+1页答案推断"]
    if explicit:
        ev = [f"检测到{explicit}个显式分页符，推断{explicit + 1}页"]
        if app_pages is not None:
            ev.append(f"docProps/app.xml Pages={app_pages}（该元数据可能过时，未优先采用）")
        return explicit + 1, "显式分页符", ev
    if len(segments) > 1:
        ev = [f"正文被切为{len(segments)}段"]
        if app_pages is not None:
            ev.append(f"docProps/app.xml Pages={app_pages}（该元数据可能过时，未优先采用）")
        return len(segments), "显式分页符分段", ev
    if app_pages is not None:
        evidence.append(f"docProps/app.xml Pages={app_pages}（该元数据可能过时）")
        return app_pages, "docProps/app.xml", evidence
    return None, "未知", ["无可用分页证据"]


class RubricEvaluator:
    def __init__(self, path: Path):
        self.path = path
        self.inspector = DocxInspector(path)
        self.dimension1: list[CheckResult] = []
        self.dimension2: list[CheckResult] = []
        self.d1_passed = False

    def add_d1(self, cid: str, name: str, passed: bool, evidence: Iterable[str]) -> None:
        self.dimension1.append(CheckResult(cid, name, 0, 0, passed, list(evidence)))

    def add_d2(self, cid: str, name: str, points: int, passed: bool, evidence: Iterable[str]) -> None:
        self.dimension2.append(CheckResult(cid, name, points, points if passed else 0, passed, list(evidence)))

    def evaluate(self) -> None:
        self.check_dimension1()
        if self.d1_passed:
            self.check_dimension2()

    def check_dimension1(self) -> None:
        self.add_d1("D1.1", ".docx格式", self.path.suffix.lower() == ".docx", [f"文件扩展名：{self.path.suffix}"])

        opened = self.inspector.open()
        self.add_d1(
            "D1.2",
            "文件可正常打开",
            opened,
            ["DOCX ZIP包与word/document.xml解析成功"] if opened else [self.inspector.open_error or "打开失败"],
        )
        if not opened:
            self.d1_passed = False
            return

        text = self.inspector.text()
        media = self.inspector.media_names()
        chars = len(text)
        chinese_chars = len(CHINESE_RE.findall(text))
        table_count = len(self.inspector.tables())
        large_images = [xy for xy in self.inspector.drawing_extents_cm() if xy[0] >= 15 and xy[1] >= 20]
        editable_ok = (not media and chars >= 300 and table_count >= 3) or (chars >= 1000 and chinese_chars >= 300 and table_count >= 3)
        screenshot_like = bool(large_images) and chars < 300
        self.add_d1(
            "D1.5",
            "可编辑Word对象，非PDF整页截图",
            editable_ok and not screenshot_like,
            [f"可编辑文本字符={chars}，中文字符={chinese_chars}，表格={table_count}，媒体图片={len(media)}，大图={len(large_images)}"],
        )

        self.d1_passed = all(r.passed for r in self.dimension1)

    def check_dimension2(self) -> None:
        self.check_page_setup()
        self.check_header()
        self.check_title_table()
        self.check_prompt_box()
        self.check_score_table()
        self.check_module_frames()
        self.check_module_colors()
        self.check_pinyin_titles()
        self.check_body_font()
        self.check_footer_pages()
        self.check_page_distribution()

    def check_page_setup(self) -> None:
        sects = self.inspector.section_properties()
        evidence: list[str] = []
        if not sects:
            self.add_d2(
                "D2.1",
                "页面设置：A4纵向21×29.7cm、上1/下0.8/左1.3/右1.3cm、四周浅蓝1磅细边框位于页边内侧且正文页码在框内",
                1,
                False,
                ["未发现sectPr"],
            )
            return
        ok_all = True
        # 页脚（页码）是否位于页面底边距以内，用于判断“页码位于边框内部”。
        footer_inside = True
        for i, sect in enumerate(sects, 1):
            pg_sz = sect.find(f"{{{W}}}pgSz")
            pg_mar = sect.find(f"{{{W}}}pgMar")
            pg_borders = sect.find(f"{{{W}}}pgBorders")

            # 纸张：A4 纵向，宽 21cm=11906 twips，高 29.7cm=16838 twips。
            w = to_int(wattr(pg_sz, "w"), 0)
            h = to_int(wattr(pg_sz, "h"), 0)
            orient = wattr(pg_sz, "orient", "portrait")
            size_ok = abs(w - 11906) <= 20 and abs(h - 16838) <= 20
            orient_ok = orient == "portrait" and h >= w
            a4_ok = size_ok and orient_ok

            # 页边距：上1 / 下0.8 / 左1.3 / 右1.3 cm。
            margin_ok = True
            margin_desc = []
            margin_values: dict[str, Optional[float]] = {}
            if pg_mar is None:
                margin_ok = False
            else:
                for key, expected in EXPECTED_MARGINS_CM.items():
                    actual = twips_to_cm(wattr(pg_mar, key))
                    margin_values[key] = actual
                    margin_desc.append(f"{key}={actual:.2f}cm" if actual is not None else f"{key}=缺失")
                    # 1mm 容差，处理 twips↔cm 量化误差。
                    if not is_near(actual, expected, 0.05):
                        margin_ok = False

            # 页面边框：四周齐全、浅蓝、线宽=1磅(sz=8)，且位于页面边缘内侧。
            border_ok = False
            border_desc = []
            if pg_borders is not None:
                sides = {local_name(ch.tag): ch for ch in pg_borders if local_name(ch.tag) in {"top", "bottom", "left", "right"}}
                border_ok = all(side in sides for side in ["top", "bottom", "left", "right"])
                # offsetFrom="page" 表示边框相对页面边缘度量；细则要求“位于页面边缘内侧”，
                # 该属性需为 page，且每边 offset 应为正数（非 0 即代表在页边以内）。
                offset_from = wattr(pg_borders, "offsetFrom")
                if offset_from != "page":
                    border_ok = False
                for side, ch in sides.items():
                    color = wattr(ch, "color")
                    sz = to_int(wattr(ch, "sz"), 0)
                    val = wattr(ch, "val", "")
                    space = to_int(wattr(ch, "space"), 0)
                    border_desc.append(f"{side}:{color}/sz={sz}/val={val}/space={space}")
                    # 浅蓝 + 1磅(sz=8) + 实线（非 nil/none）。
                    if not is_light_blue(color) or sz != 8 or val in {"", "nil", "none"}:
                        border_ok = False
                    # 在页面边缘内侧：偏移量必须为正。
                    if space <= 0:
                        border_ok = False

            # 页码位于边框内部：页脚距底边距离 < 下边距，即页码落在下边框以上区域。
            section_footer_ok = True
            if pg_mar is not None:
                footer_dist = twips_to_cm(wattr(pg_mar, "footer"))
                bottom_cm = margin_values.get("bottom")
                if footer_dist is not None and bottom_cm is not None and footer_dist >= bottom_cm:
                    section_footer_ok = False
                    footer_inside = False

            ok_all = ok_all and a4_ok and margin_ok and border_ok and section_footer_ok
            evidence.append(
                f"section{i}: pgSz={w}x{h}/orient={orient} A4={a4_ok}; "
                f"{'/'.join(margin_desc)}; border={' '.join(border_desc)}; footer_inside={section_footer_ok}"
            )
        self.add_d2(
            "D2.1",
            "页面设置：A4纵向21×29.7cm、上1/下0.8/左1.3/右1.3cm、四周浅蓝1磅细边框位于页边内侧且正文页码在框内",
            1,
            ok_all,
            evidence,
        )

    def check_header(self) -> None:
        # 细则：第1页至第5页顶部页眉中的【拼音装饰短句】和【中文短句】，
        # 拼音短句字体 Times New Roman、中文短句字体 Noto Sans CJK SC，
        # 二者字号均为 8.5 磅(=17 半磅)，颜色为灰蓝色，水平居中。
        #
        # 关键判据（相对旧实现的修正）：
        #   · 不再"页眉内任意位置出现 Times/Noto/17/灰蓝/居中"即通过——那样
        #     Times 出现在 A run、17 出现在 B run 也会误判为满足；
        #   · 改为把字体/字号/颜色逐 run 绑定到其所属短句：
        #       拼音 run  -> 西文字体(ascii/hAnsi)=Times New Roman，17 半磅，灰蓝；
        #       中文 run  -> 中文字体(eastAsia)=Noto Sans CJK SC，17 半磅，灰蓝；
        #     且承载该短句的段落必须水平居中；两句都必须存在。
        #   · 页码范围：无渲染无法把 header part 精确映射到"第几页"，
        #     采用 OOXML 既有事实——header 通过 sectPr 的 headerReference 作用于
        #     其所在 section 覆盖的所有页。因此要求：文档页数覆盖到第5页，且
        #     所有【被引用且非空】的页眉 part 都逐句满足上述格式（无论落在哪一页
        #     都合规）。若确需精确到"每一页"须渲染，会另行告知。
        headers = self.inspector.headers
        rule_name = ("第1页至第5页顶部页眉：拼音短句(Times New Roman)+中文短句(Noto Sans CJK SC)，"
                     "均8.5磅、灰蓝色、水平居中")
        evidence: list[str] = [f"页眉part数={len(headers)}"]

        if not headers:
            self.add_d2("D2.2", rule_name, 1, False, ["未发现页眉part"])
            return

        page_count, _, _ = get_page_count(self.inspector)
        pages_ok = page_count is not None and page_count >= 5
        evidence.append(f"页数={page_count} 覆盖第5页={pages_ok}")

        latin_re = re.compile(r"[A-Za-zāáǎàēéěèīíǐìōóǒòūúǔùǖǘǚǜü]")

        class HRun(NamedTuple):
            text: str
            fonts: dict[str, str]
            size: int
            color: str
            jc: Optional[str]

        def para_run_infos(header: ET.Element) -> list[HRun]:
            """返回该 header 内每个 run 的 (text,fonts,size,color,jc)（跳过空 run）。
            size/color 若 run 级缺省则回退到段落标记 rPr(pPr/rPr)。"""
            infos: list[HRun] = []
            for p in header.findall(f".//{{{W}}}p"):
                pPr = p.find(f"{{{W}}}pPr")
                jc = wattr(pPr.find(f"{{{W}}}jc"), "val") if pPr is not None else None
                pmark_size, pmark_color = 0, ""
                pmark_rpr = pPr.find(f"{{{W}}}rPr") if pPr is not None else None
                if pmark_rpr is not None:
                    pmark_size = to_int(wattr(pmark_rpr.find(f"{{{W}}}sz"), "val"), 0)
                    pmark_color = (wattr(pmark_rpr.find(f"{{{W}}}color"), "val") or "").upper()
                for r in p.findall(f"{{{W}}}r"):
                    text = "".join(x.text or "" for x in r.findall(f".//{{{W}}}t"))
                    if not text.strip():
                        continue
                    rPr = r.find(f"{{{W}}}rPr")
                    fonts: dict[str, str] = {}
                    rfonts = rPr.find(f"{{{W}}}rFonts") if rPr is not None else None
                    if rfonts is not None:
                        for k in ("ascii", "hAnsi", "eastAsia", "cs"):
                            v = wattr(rfonts, k)
                            if v:
                                fonts[k] = v
                    size = to_int(wattr(rPr.find(f"{{{W}}}sz"), "val"), 0) if rPr is not None else 0
                    color = ((wattr(rPr.find(f"{{{W}}}color"), "val") or "").upper()
                             if rPr is not None else "")
                    infos.append(HRun(text, fonts, size or pmark_size, color or pmark_color, jc))
            return infos

        any_nonempty = False
        all_headers_ok = True

        for idx, header in enumerate(headers, 1):
            runs = para_run_infos(header)
            if not runs:
                evidence.append(f"header{idx}: 空")
                continue
            any_nonempty = True

            pinyin_runs = [r for r in runs
                           if latin_re.search(r.text) and not CHINESE_RE.search(r.text)]
            cn_runs = [r for r in runs if CHINESE_RE.search(r.text)]

            has_pinyin = len(pinyin_runs) > 0
            has_cn = len(cn_runs) > 0

            def sentence_ok(sruns: list[HRun], font_name: str, font_keys: tuple[str, ...]) -> bool:
                if not sruns:
                    return False
                for r in sruns:
                    if font_name not in [r.fonts.get(k) for k in font_keys]:
                        return False
                    if r.size != 17:
                        return False
                    if not is_gray_blue(r.color):
                        return False
                    if r.jc not in {"center", "both"}:
                        return False
                return True

            pinyin_ok = sentence_ok(pinyin_runs, "Times New Roman", ("ascii", "hAnsi"))
            cn_ok = sentence_ok(cn_runs, "Noto Sans CJK SC", ("eastAsia", "ascii", "hAnsi"))

            header_ok = has_pinyin and has_cn and pinyin_ok and cn_ok
            all_headers_ok = all_headers_ok and header_ok

            evidence.append(
                f"header{idx}: 拼音句={has_pinyin}/格式ok={pinyin_ok} "
                f"中文句={has_cn}/格式ok={cn_ok} "
                f"pinyin_fonts={sorted({v for r in pinyin_runs for v in r.fonts.values()})} "
                f"cn_fonts={sorted({v for r in cn_runs for v in r.fonts.values()})}"
            )

        ok = pages_ok and any_nonempty and all_headers_ok
        self.add_d2("D2.2", rule_name, 1, ok, evidence)

    def check_title_table(self) -> None:
        # 细则：
        # 1) 顶部使用 1 行 2 列可编辑 Word 表格；
        # 2) 左侧单元格填充浅蓝色，右侧单元格填充浅橙色；
        # 3) 两列边框为浅蓝色细实线；
        # 4) 左侧标题：依次显示
        #    a. 「三年级下册语文达标测试卷」上方拼音，Times New Roman 小五倾斜；
        #    b. 「三年级下册语文达标测试卷」Noto Sans CJK SC 小一加粗深蓝色；
        #    c. 「第四单元测试卷」「时间：60分钟　满分：100分」Noto Sans CJK SC 五号灰色；
        # 5) 右侧学生信息：分行显示「学校：…」「班级：…」「姓名：…」「得分：…」，
        #    Noto Sans CJK SC 9.5 磅，黑色，水平居中。
        tables = self.inspector.tables()
        if not tables:
            self.add_d2(
                "D2.3",
                "第1页顶部标题区域：1行2列表格，左浅蓝右浅橙，浅蓝细实线边框；左侧标题三段、右侧学生信息四行",
                5,
                False,
                ["未发现表格"],
            )
            return
        t = tables[0]

        # 表格形状：1 行 2 列。
        shape_ok = t.rows == 1 and t.cells_per_row == [2]

        # 左右单元格填充色。
        cells = t.el.findall(f".//{{{W}}}tc")
        left_fill = cell_fill(cells[0]) if len(cells) >= 1 else None
        right_fill = cell_fill(cells[1]) if len(cells) >= 2 else None
        fill_ok = is_light_blue(left_fill) and is_light_orange(right_fill)

        # 边框：浅蓝色细实线（sz<=8 ≈ 1磅）。
        border_color_ok = bool(t.border_colors) and all(is_light_blue(c) for c in t.border_colors)
        border_thin_solid = True
        for bdr in t.el.findall(f".//{{{W}}}tblBorders/*") + t.el.findall(f".//{{{W}}}tcBorders/*"):
            val = wattr(bdr, "val", "")
            sz = to_int(wattr(bdr, "sz"), 0)
            if val in {"nil", "none"}:
                continue
            if val != "single" or sz == 0 or sz > 8:
                border_thin_solid = False
        border_ok = border_color_ok and border_thin_solid

        # 提取左右单元格的段落+run详细信息。
        def cell_paragraphs(tc: ET.Element) -> list[dict]:
            paras: list[dict] = []
            for p in tc.findall(f"{{{W}}}p"):
                jc = p.find(f"{{{W}}}pPr/{{{W}}}jc")
                jc_val = wattr(jc, "val")
                text = element_text(p)
                runs = []
                for r in p.findall(f"{{{W}}}r"):
                    rfonts = r.find(f"{{{W}}}rPr/{{{W}}}rFonts")
                    sz = r.find(f"{{{W}}}rPr/{{{W}}}sz")
                    color = r.find(f"{{{W}}}rPr/{{{W}}}color")
                    bold = r.find(f"{{{W}}}rPr/{{{W}}}b") is not None
                    italic = r.find(f"{{{W}}}rPr/{{{W}}}i") is not None
                    rt = "".join(x.text or "" for x in r.findall(f".//{{{W}}}t"))
                    fonts = {}
                    if rfonts is not None:
                        for k in ("ascii", "hAnsi", "eastAsia"):
                            v = wattr(rfonts, k)
                            if v:
                                fonts[k] = v
                    runs.append({
                        "text": rt,
                        "fonts": fonts,
                        "size": to_int(wattr(sz, "val"), 0),
                        "color": (wattr(color, "val") or "").upper(),
                        "bold": bold,
                        "italic": italic,
                    })
                paras.append({"text": text, "jc": jc_val, "runs": runs})
            return paras

        left_paras = cell_paragraphs(cells[0]) if len(cells) >= 1 else []
        right_paras = cell_paragraphs(cells[1]) if len(cells) >= 2 else []

        # 左侧：找到三类段落。
        pinyin_para = next((p for p in left_paras if PINYIN_RE.search(p["text"]) and not CHINESE_RE.search(p["text"])), None)
        main_title_para = next((p for p in left_paras if "三年级下册语文达标测试卷" in p["text"]), None)
        subtitle_para = next((p for p in left_paras if "第四单元测试卷" in p["text"] and "时间" in p["text"] and "满分" in p["text"]), None)

        def run_attr_ok(para: Optional[dict], font_name: str, size_hp: int, *, bold: Optional[bool] = None, italic: Optional[bool] = None, color_pred=None) -> bool:
            if not para or not para["runs"]:
                return False
            for run in para["runs"]:
                if font_name not in run["fonts"].values():
                    return False
                if run["size"] != size_hp:
                    return False
                if bold is not None and run["bold"] != bold:
                    return False
                if italic is not None and run["italic"] != italic:
                    return False
                if color_pred is not None and not color_pred(run["color"]):
                    return False
            return True

        # 拼音：Times New Roman / 小五(9磅=18半磅) / 倾斜。
        pinyin_ok = run_attr_ok(pinyin_para, "Times New Roman", 18, italic=True)

        # 主标题：Noto Sans CJK SC / 小一(24磅=48半磅) / 加粗 / 深蓝色。
        main_title_ok = run_attr_ok(main_title_para, "Noto Sans CJK SC", 48, bold=True, color_pred=is_dark_blue)

        # 副标题：Noto Sans CJK SC / 五号(10.5磅=21半磅) / 灰色。
        subtitle_ok = run_attr_ok(subtitle_para, "Noto Sans CJK SC", 21, color_pred=is_gray)

        left_ok = pinyin_ok and main_title_ok and subtitle_ok

        # 右侧：4 行，分别含学校/班级/姓名/得分；Noto Sans CJK SC，9.5磅(=19半磅)，黑色，水平居中。
        required_labels = ["学校", "班级", "姓名", "得分"]
        right_lines = {}
        for label in required_labels:
            right_lines[label] = next((p for p in right_paras if p["text"].startswith(label) or label + "：" in p["text"] or label + ":" in p["text"]), None)
        right_ok = True
        for label in required_labels:
            para = right_lines[label]
            if not para:
                right_ok = False
                continue
            if para["jc"] != "center":
                right_ok = False
            if not para["runs"]:
                right_ok = False
                continue
            for run in para["runs"]:
                if "Noto Sans CJK SC" not in run["fonts"].values():
                    right_ok = False
                if run["size"] != 19:
                    right_ok = False
                if not is_black(run["color"]):
                    right_ok = False

        ok = shape_ok and fill_ok and border_ok and left_ok and right_ok

        self.add_d2(
            "D2.3",
            "第1页顶部标题区域：1行2列表格，左浅蓝右浅橙，浅蓝细实线边框；左侧标题三段、右侧学生信息四行",
            5,
            ok,
            [
                f"shape rows={t.rows} cells_per_row={t.cells_per_row} ok={shape_ok}",
                f"fills left={left_fill} right={right_fill} ok={fill_ok}",
                f"border colors={sorted(set(t.border_colors))} thin_solid={border_thin_solid} ok={border_ok}",
                f"pinyin_ok={pinyin_ok} para={(pinyin_para or {}).get('text','')[:30]}",
                f"main_title_ok={main_title_ok} para={(main_title_para or {}).get('text','')[:30]}",
                f"subtitle_ok={subtitle_ok} para={(subtitle_para or {}).get('text','')[:40]}",
                f"right_ok={right_ok} labels_found={[k for k,v in right_lines.items() if v]}",
            ],
        )

    def check_prompt_box(self) -> None:
        # 细则：
        # 1) 标题区域下方设置浅黄色单元格；
        # 2) 显示文本「闯关提示：请认真拼读每一道题，书写要端正，遇到轻声和整体认读音节要多读几遍。」；
        # 3) 「闯关提示：」字体为 Noto Sans CJK SC 五号(21 半磅)，加粗，且使用深橙色；
        # 4) 「请认真拼读每一道题，书写要端正，遇到轻声和整体认读音节要多读几遍。」
        #    字体为 Noto Serif CJK SC 五号(21 半磅)。
        EXPECTED_TEXT = "闯关提示：请认真拼读每一道题，书写要端正，遇到轻声和整体认读音节要多读几遍。"
        LABEL_TEXT = "闯关提示："

        tables = self.inspector.tables()
        candidates = [t for t in tables if LABEL_TEXT.rstrip("：:") in t.text]
        if not candidates:
            self.add_d2(
                "D2.4",
                "第1页考试信息提示框：浅黄色单元格；「闯关提示：」NotoSansCJKSC五号加粗深橙色；正文 NotoSerifCJKSC 五号",
                1,
                False,
                ["未发现含「闯关提示」的表格"],
            )
            return
        t = candidates[0]

        # 1) 浅黄色单元格填充。
        cells = t.el.findall(f".//{{{W}}}tc")
        cell = cells[0] if cells else None
        fill = cell_fill(cell)
        fill_ok = is_light_yellow(fill)

        # 2) 完整文本匹配（去除可能的空白差异）。
        normalized = re.sub(r"\s+", "", t.text)
        text_ok = re.sub(r"\s+", "", EXPECTED_TEXT) in normalized

        # 3) 收集表格内所有非空 run；按 run 文本归属到 label 段 / 正文段。
        # 真实模板里「闯关提示：」常单独一个 run，正文文字另一个 run；以子串归类，
        # 跨 run 拆分时按落入哪个区间归到对应集合。
        @dataclass
        class RunInfo:
            text: str
            fonts: list[str]
            size: int
            color: str
            bold: bool

        label_runs: list[RunInfo] = []
        body_runs: list[RunInfo] = []
        for r in t.el.findall(f".//{{{W}}}r"):
            rt = "".join(x.text or "" for x in r.findall(f".//{{{W}}}t"))
            if not rt.strip():
                continue
            rfonts = r.find(f"{{{W}}}rPr/{{{W}}}rFonts")
            sz = r.find(f"{{{W}}}rPr/{{{W}}}sz")
            color = r.find(f"{{{W}}}rPr/{{{W}}}color")
            bold = r.find(f"{{{W}}}rPr/{{{W}}}b") is not None
            font_list: list[str] = []
            if rfonts is not None:
                for k in ("ascii", "hAnsi", "eastAsia"):
                    v = wattr(rfonts, k)
                    if v:
                        font_list.append(v)
            info = RunInfo(
                text=rt,
                fonts=font_list,
                size=to_int(wattr(sz, "val"), 0),
                color=(wattr(color, "val") or "").upper(),
                bold=bold,
            )
            # 「闯关提示」（含/不含冒号）算 label run；否则算正文 run。
            if "闯关提示" in rt:
                label_runs.append(info)
            else:
                body_runs.append(info)

        # 3) 「闯关提示：」：Noto Sans CJK SC / 五号(21) / 加粗 / 深橙色。
        label_font_ok = bool(label_runs) and all("Noto Sans CJK SC" in r.fonts for r in label_runs)
        label_size_ok = bool(label_runs) and all(r.size == 21 for r in label_runs)
        label_bold_ok = bool(label_runs) and all(r.bold for r in label_runs)
        label_color_ok = bool(label_runs) and all(is_dark_orange(r.color) for r in label_runs)

        # 4) 「请认真拼读每一道题，书写要端正，遇到轻声和整体认读音节要多读几遍。」：
        #    Noto Serif CJK SC / 五号(21)。
        body_font_ok = bool(body_runs) and all("Noto Serif CJK SC" in r.fonts for r in body_runs)
        body_size_ok = bool(body_runs) and all(r.size == 21 for r in body_runs)

        ok = (
            fill_ok
            and text_ok
            and label_font_ok and label_size_ok and label_bold_ok and label_color_ok
            and body_font_ok and body_size_ok
        )

        self.add_d2(
            "D2.4",
            "第1页考试信息提示框：浅黄色单元格；「闯关提示：」NotoSansCJKSC五号加粗深橙色；正文 NotoSerifCJKSC 五号",
            1,
            ok,
            [
                f"候选表格#{t.index} text={t.text[:80]}",
                f"fill={fill} 浅黄={fill_ok}",
                f"text_ok={text_ok}",
                f"label_runs={[(r.text, r.bold, r.color, r.size, r.fonts) for r in label_runs]}",
                f"label font_ok={label_font_ok} size_ok={label_size_ok} bold_ok={label_bold_ok} color_ok={label_color_ok}",
                f"body_runs={[(r.text, r.size, r.fonts) for r in body_runs]}",
                f"body font_ok={body_font_ok} size_ok={body_size_ok}",
            ],
        )

    def check_score_table(self) -> None:
        # 细则：
        # 1) 使用 2 行 11 列可编辑 Word 表格；
        # 2) 第一行依次显示「题号、一、二、三、四、五、六、七、八、九、总分」；
        # 3) 第二行显示「得分」及 10 个空白计分单元格；
        # 4) 表格边框完整；
        # 5) 字体 Noto Sans CJK SC 五号(21 半磅)；（不再要求加粗）
        # 6) 文字水平居中、垂直居中。
        ROW0_HEADERS = ["题号", "一", "二", "三", "四", "五", "六", "七", "八", "九", "总分"]

        tables = self.inspector.tables()
        candidates = [t for t in tables if all(k in t.text for k in ["题号", "得分", "总分"])]
        if not candidates:
            self.add_d2(
                "D2.5",
                "第1页成绩统计表：2行11列、表头与得分行、边框完整、Noto Sans CJK SC 五号、文字水平居中且垂直居中",
                1,
                False,
                ["未发现含题号/得分/总分的表格"],
            )
            return
        t = candidates[0]

        # 1) 2 行 11 列。
        shape_ok = t.rows == 2 and t.cells_per_row == [11, 11]

        rows = t.el.findall(f"{{{W}}}tr")

        def cell_text(tc: ET.Element) -> str:
            return "".join(x.text or "" for x in tc.findall(f".//{{{W}}}t"))

        # 2) 第一行表头文本依次匹配。
        row0_ok = False
        if rows and len(rows[0].findall(f"{{{W}}}tc")) == 11:
            actual0 = [cell_text(tc).strip() for tc in rows[0].findall(f"{{{W}}}tc")]
            row0_ok = actual0 == ROW0_HEADERS
        else:
            actual0 = []

        # 3) 第二行：c0="得分"，c1..c10 均为空白。
        row1_ok = False
        if len(rows) >= 2 and len(rows[1].findall(f"{{{W}}}tc")) == 11:
            row1_cells = rows[1].findall(f"{{{W}}}tc")
            label_ok = cell_text(row1_cells[0]).strip() == "得分"
            blanks_ok = all(cell_text(tc).strip() == "" for tc in row1_cells[1:])
            row1_ok = label_ok and blanks_ok
        else:
            label_ok = False
            blanks_ok = False

        # 4) 边框完整：tblBorders 包含 top/bottom/left/right/insideH/insideV，且均非 nil。
        border_ok = False
        tblBorders = t.el.find(f".//{{{W}}}tblBorders")
        if tblBorders is not None:
            sides = {local_name(ch.tag): ch for ch in tblBorders}
            required = {"top", "bottom", "left", "right", "insideH", "insideV"}
            border_ok = required.issubset(sides.keys())
            for side in required:
                if side in sides:
                    val = wattr(sides[side], "val", "")
                    if val in {"", "nil", "none"}:
                        border_ok = False

        # 5)+6) 字体/字号 + 水平居中/垂直居中。所有有文本的 run 与所有单元格统一校验。（不再校验加粗）
        font_ok = True
        size_ok = True
        hcenter_ok = True
        vcenter_ok = True
        any_run = False

        for row in rows:
            for tc in row.findall(f"{{{W}}}tc"):
                # 垂直居中：tcPr/vAlign=center。
                vAlign = tc.find(f"{{{W}}}tcPr/{{{W}}}vAlign")
                if wattr(vAlign, "val") != "center":
                    vcenter_ok = False
                for p in tc.findall(f"{{{W}}}p"):
                    # 水平居中：pPr/jc=center。
                    jc = p.find(f"{{{W}}}pPr/{{{W}}}jc")
                    if wattr(jc, "val") != "center":
                        hcenter_ok = False
                    for r in p.findall(f"{{{W}}}r"):
                        rt = "".join(x.text or "" for x in r.findall(f".//{{{W}}}t"))
                        if not rt:
                            continue
                        any_run = True
                        rfonts = r.find(f"{{{W}}}rPr/{{{W}}}rFonts")
                        sz = r.find(f"{{{W}}}rPr/{{{W}}}sz")
                        fonts: list[str] = []
                        if rfonts is not None:
                            for k in ("ascii", "hAnsi", "eastAsia"):
                                v = wattr(rfonts, k)
                                if v:
                                    fonts.append(v)
                        if "Noto Sans CJK SC" not in fonts:
                            font_ok = False
                        if to_int(wattr(sz, "val"), 0) != 21:
                            size_ok = False

        runs_present = any_run

        ok = (
            shape_ok
            and row0_ok
            and row1_ok
            and border_ok
            and runs_present
            and font_ok
            and size_ok
            and hcenter_ok
            and vcenter_ok
        )

        self.add_d2(
            "D2.5",
            "第1页成绩统计表：2行11列、表头与得分行、边框完整、Noto Sans CJK SC 五号、文字水平居中且垂直居中",
            1,
            ok,
            [
                f"候选表格#{t.index}: rows={t.rows}, cells_per_row={t.cells_per_row}, shape_ok={shape_ok}",
                f"row0 actual={actual0} ok={row0_ok}",
                f"row1 得分={'ok' if row1_ok else 'no'} (label_ok={label_ok if rows else False}, blanks_ok={blanks_ok if rows else False})",
                f"border_ok={border_ok}",
                f"font_ok={font_ok} size_ok={size_ok}",
                f"hcenter_ok={hcenter_ok} vcenter_ok={vcenter_ok}",
            ],
        )

    def module_tables(self) -> list[TableInfo]:
        modules: list[TableInfo] = []
        seen_titles = set()
        for t in self.inspector.tables():
            found = [m for m in EXPECTED_MODULES if m in t.text]
            if found:
                # 一个大题只取首次出现的外层/标题表，避免作文格子等内层表重复干扰。
                title = found[0]
                if title not in seen_titles:
                    seen_titles.add(title)
                    modules.append(t)
        return modules

    def check_module_frames(self) -> None:
        # 细则：
        # 1) 第1页至第4页每道大题（一~九共9个模块）使用单列可编辑Word表格；
        # 2) 表格宽度为 18.4 厘米；
        # 3) 边框为浅蓝色 0.5 磅实线。
        EXPECTED_WIDTH_TWIPS = round(18.4 * TWIPS_PER_CM)  # 18.4cm ≈ 10433 twips
        modules = self.module_tables()

        details: list[str] = []
        all_ok = True

        if len(modules) < 9:
            all_ok = False

        for t in modules[:9]:
            tag = next((m for m in EXPECTED_MODULES if m in t.text), "?")

            # 1) 单列表格：cells_per_row 每行均为 1。
            single_col = bool(t.cells_per_row) and all(c == 1 for c in t.cells_per_row)

            # 2) 模块宽度=18.4cm：必须显式 tblW，type=dxa，w≈10433 twips。
            tblW = t.el.find(f".//{{{W}}}tblPr/{{{W}}}tblW")
            w_type = wattr(tblW, "type")
            w_val = to_int(wattr(tblW, "w"), 0)
            w_cm = w_val / TWIPS_PER_CM if w_val else 0.0
            # 1mm 容差。
            width_ok = w_type == "dxa" and abs(w_val - EXPECTED_WIDTH_TWIPS) <= round(0.05 * TWIPS_PER_CM)

            # 3) 边框：四周齐全；val=single；sz=4 (0.5磅)；颜色为浅蓝。
            tblBorders = t.el.find(f".//{{{W}}}tblPr/{{{W}}}tblBorders")
            border_sides_ok = False
            border_style_ok = True
            border_color_ok = True
            border_desc: list[str] = []
            if tblBorders is not None:
                sides = {local_name(ch.tag): ch for ch in tblBorders if local_name(ch.tag) in {"top", "bottom", "left", "right"}}
                border_sides_ok = all(s in sides for s in ("top", "bottom", "left", "right"))
                for side in ("top", "bottom", "left", "right"):
                    if side not in sides:
                        border_style_ok = False
                        border_color_ok = False
                        continue
                    ch = sides[side]
                    val = wattr(ch, "val", "")
                    sz = to_int(wattr(ch, "sz"), 0)
                    color = wattr(ch, "color")
                    border_desc.append(f"{side}:{val}/sz={sz}/{color}")
                    if val != "single" or sz != 4:
                        border_style_ok = False
                    if not is_light_blue(color):
                        border_color_ok = False
            else:
                border_style_ok = False
                border_color_ok = False

            module_ok = single_col and width_ok and border_sides_ok and border_style_ok and border_color_ok
            all_ok = all_ok and module_ok
            details.append(
                f"#{t.index}:{tag} single_col={single_col} "
                + f"tblW={w_val}({w_cm:.3f}cm if dxa)/type={w_type} width_ok={width_ok} "
                + f"sides_ok={border_sides_ok} style_ok={border_style_ok} color_ok={border_color_ok} "
                + f"borders=[{', '.join(border_desc)}]"
            )

        self.add_d2(
            "D2.6",
            "第1页至第4页题目模块外框：每道大题使用单列可编辑Word表格，宽度18.4cm，浅蓝色0.5磅实线边框",
            3,
            all_ok,
            [f"检测到大题模块={len(modules)}/9"] + details,
        )

    def check_module_colors(self) -> None:
        # 细则：
        # 1) 第一、第四、第七大题 → 背景极浅蓝色，中文标题 Noto Sans CJK SC 13.5 磅 加粗 蓝色；
        # 2) 第二、第五、第八大题 → 背景极浅绿色，中文标题 Noto Sans CJK SC 13.5 磅 加粗 绿色;
        # 3) 第三、第六、第九大题 → 背景极浅橙色，中文标题 Noto Sans CJK SC 13.5 磅 加粗 橙色。
        EXPECTED_BG = ["blue", "green", "orange", "blue", "green", "orange", "blue", "green", "orange"]
        EXPECTED_TITLE_LABEL = ["蓝", "绿", "橙", "蓝", "绿", "橙", "蓝", "绿", "橙"]
        title_color_predicates = {"蓝": is_blue, "绿": is_green, "橙": is_orange}

        modules = self.module_tables()
        details: list[str] = []
        all_ok = len(modules) >= 9

        for idx, t in enumerate(modules[:9]):
            tag = next((m for m in EXPECTED_MODULES if m in t.text), "?")
            expected_bg = EXPECTED_BG[idx]
            expected_title = EXPECTED_TITLE_LABEL[idx]
            title_pred = title_color_predicates[expected_title]

            # 背景：第一个单元格 shd@fill。
            bg_fill = cell_fill(t.el.find(f".//{{{W}}}tc"))
            bg_ok = color_name(bg_fill) == expected_bg

            # 标题段：包含「X、」的段落（第一段，跳过拼音）。
            title_para = None
            for p in t.el.findall(f".//{{{W}}}tc")[0].findall(f"{{{W}}}p") if t.el.find(f".//{{{W}}}tc") is not None else []:
                ptext = element_text(p)
                if tag in ptext:
                    title_para = p
                    break

            font_ok = False
            size_ok = False
            bold_ok = False
            color_ok = False
            title_text = ""
            if title_para is not None:
                title_text = element_text(title_para)
                # 中文标题的 run：含中文字符。
                cn_runs: list[ET.Element] = []
                for r in title_para.findall(f"{{{W}}}r"):
                    rt = "".join(x.text or "" for x in r.findall(f".//{{{W}}}t"))
                    if CHINESE_RE.search(rt):
                        cn_runs.append(r)
                if cn_runs:
                    font_ok = True
                    size_ok = True
                    bold_ok = True
                    color_ok = True
                    for r in cn_runs:
                        rfonts = r.find(f"{{{W}}}rPr/{{{W}}}rFonts")
                        sz = r.find(f"{{{W}}}rPr/{{{W}}}sz")
                        color = r.find(f"{{{W}}}rPr/{{{W}}}color")
                        bold = r.find(f"{{{W}}}rPr/{{{W}}}b") is not None
                        fonts_v: list[str] = []
                        if rfonts is not None:
                            for k in ("ascii", "hAnsi", "eastAsia"):
                                v = wattr(rfonts, k)
                                if v:
                                    fonts_v.append(v)
                        if "Noto Sans CJK SC" not in fonts_v:
                            font_ok = False
                        # 13.5 磅 = 27 半磅。
                        if to_int(wattr(sz, "val"), 0) != 27:
                            size_ok = False
                        if not bold:
                            bold_ok = False
                        if not title_pred(wattr(color, "val")):
                            color_ok = False

            module_ok = bg_ok and font_ok and size_ok and bold_ok and color_ok
            all_ok = all_ok and module_ok
            details.append(
                f"#{t.index}:{tag} bg_fill={bg_fill}({color_name(bg_fill)}/期望{expected_bg}) bg_ok={bg_ok} "
                + f"title='{title_text[:24]}' font_ok={font_ok} size_ok={size_ok} bold_ok={bold_ok} "
                + f"color_ok={color_ok}(期望{expected_title})"
            )

        self.add_d2(
            "D2.7",
            "第1~9大题：背景极浅蓝/绿/橙循环，中文标题 Noto Sans CJK SC 13.5磅加粗，颜色对应蓝/绿/橙",
            3,
            all_ok,
            [f"检测到模块={len(modules)}/9"] + details,
        )

    def check_pinyin_titles(self) -> None:
        # 细则：第1页至第4页每道大题中文标题上方显示对应完整拼音；
        # 字体 Times New Roman 8.5 磅（17 半磅），倾斜，灰色；
        # 拼音与中文标题位于同一模块内。
        #
        # 关键判据（相对旧实现的修正）：
        #   · 不再只判断"存在拼音形态文本"——旧实现会因 tag 段落里含拉丁字符
        #     也算命中；
        #   · "对应完整拼音"通过【拼音音节数 == 中文标题汉字数】强校验：
        #       去声调 → 按空白/连字符切分成音节 token → 每个 token 必须是纯拉丁小写
        #       字母（≥1）；中文标题剔除 "X、" 前缀标记与所有非汉字字符后统计汉字数。
        #     音节数与汉字数一致才算"完整拼音一一对应"。
        #   · "拼音位于中文标题上方"通过【同一 tc 内段落索引，pinyin_idx < title_idx】强校验。
        #   · 若期望硬编码 9 条标题拼音，本 verifier 未内建（避免装 pypinyin/xpinyin
        #     等第三方库；且未持有源题官方 Chinese 标题清单）。归一化音节一一对应
        #     即可满足 rubric "对应完整拼音" 的实质要求；如需精确到字面拼音串，
        #     请提供 9 道大题的中文标题，我再补入 MODULE_EXPECTED_PINYIN 字典。
        modules = self.module_tables()
        details: list[str] = []
        all_ok = len(modules) >= 9

        # 声调 → 无调映射
        tone_map = str.maketrans(
            "āáǎàēéěèīíǐìōóǒòūúǔùǖǘǚǜü",
            "aaaaeeeeiiiioooouuuuuuuuu",
        )

        def pinyin_syllables(text: str) -> list[str]:
            """把拼音文本切分为音节 token；仅返回纯拉丁字母 token。"""
            s = text.translate(tone_map).lower()
            # 常见拼音间以空格/连字符/顿号/括号分隔
            tokens = re.split(r"[\s\-·\.,;:\(\)（）、，。！？!\?]+", s)
            return [t for t in tokens if t and t.isascii() and t.isalpha()]

        def han_count(text: str, tag: str) -> int:
            """去掉"X、"前缀标记和所有非汉字字符，返回汉字数。"""
            s = text.replace(tag, "", 1) if tag in text else text
            return len(CHINESE_RE.findall(s))

        for t in modules[:9]:
            tag = next((m for m in EXPECTED_MODULES if m in t.text), "?")
            first_tc = t.el.find(f".//{{{W}}}tc")

            pinyin_para = None
            pinyin_idx: Optional[int] = None
            chinese_title_para = None
            title_idx: Optional[int] = None
            if first_tc is not None:
                for i, p in enumerate(first_tc.findall(f"{{{W}}}p")):
                    ptext = element_text(p)
                    if tag in ptext and chinese_title_para is None:
                        chinese_title_para = p
                        title_idx = i
                        continue
                    # 拼音段：不含"X、"且能切出至少 1 个纯字母音节
                    if pinyin_para is None and tag not in ptext and pinyin_syllables(ptext):
                        pinyin_para = p
                        pinyin_idx = i

            same_module = pinyin_para is not None and chinese_title_para is not None
            order_ok = (pinyin_idx is not None and title_idx is not None
                        and pinyin_idx < title_idx)

            syllables: list[str] = []
            han_n = 0
            full_ok = False
            if pinyin_para is not None and chinese_title_para is not None:
                syllables = pinyin_syllables(element_text(pinyin_para))
                han_n = han_count(element_text(chinese_title_para), tag)
                full_ok = han_n > 0 and len(syllables) == han_n

            font_ok = False
            size_ok = False
            italic_ok = False
            color_ok = False
            if pinyin_para is not None:
                runs = pinyin_para.findall(f"{{{W}}}r")
                # 只对含拉丁字母（去调后 ASCII 字母）的 run 校验字体格式
                latin_runs = []
                for r in runs:
                    rt = "".join(x.text or "" for x in r.findall(f".//{{{W}}}t"))
                    if pinyin_syllables(rt):
                        latin_runs.append(r)
                if latin_runs:
                    font_ok = True
                    size_ok = True
                    italic_ok = True
                    color_ok = True
                    for r in latin_runs:
                        rfonts = r.find(f"{{{W}}}rPr/{{{W}}}rFonts")
                        sz = r.find(f"{{{W}}}rPr/{{{W}}}sz")
                        color = r.find(f"{{{W}}}rPr/{{{W}}}color")
                        italic = r.find(f"{{{W}}}rPr/{{{W}}}i") is not None
                        fonts_v: list[str] = []
                        if rfonts is not None:
                            for k in ("ascii", "hAnsi", "eastAsia"):
                                v = wattr(rfonts, k)
                                if v:
                                    fonts_v.append(v)
                        if "Times New Roman" not in fonts_v:
                            font_ok = False
                        if to_int(wattr(sz, "val"), 0) != 17:
                            size_ok = False
                        if not italic:
                            italic_ok = False
                        if not is_gray(wattr(color, "val")):
                            color_ok = False

            module_ok = (same_module and order_ok and full_ok
                         and font_ok and size_ok and italic_ok and color_ok)
            all_ok = all_ok and module_ok
            details.append(
                f"#{t.index}:{tag} same_module={same_module} order={pinyin_idx}<{title_idx}={order_ok} "
                f"音节={len(syllables)}/汉字={han_n} full_ok={full_ok} "
                f"font_ok={font_ok} size_ok={size_ok} italic_ok={italic_ok} color_ok={color_ok}"
            )

        self.add_d2(
            "D2.8",
            "第1页至第4页大题拼音标题：完整拼音(音节数=标题汉字数)，位于中文标题上方，Times New Roman 8.5磅倾斜灰色，与中文标题同模块",
            1,
            all_ok,
            [f"模块={len(modules)}/9"] + details,
        )

    def check_body_font(self) -> None:
        # 细则：第1页至第4页题目正文中文使用 Noto Serif CJK SC 五号(21 半磅)，
        # 黑色，单倍行距，段前 0 磅、段后 0 磅。
        #
        # 关键判据（相对旧实现的修正）：
        #   · 旧实现遍历 body 内所有含中文段落，会把顶部标题表（三年级下册语文
        #     达标测试卷/学生信息 9.5 磅）、提示框（闯关提示 五号加粗深橙 +
        #     Noto Serif 正文）、成绩统计表（Noto Sans 五号加粗 = 21 半磅，
        #     但字体为 Sans 而非 Serif）、每个大题的模块标题（Noto Sans 13.5
        #     磅加粗蓝/绿/橙）、大题拼音标题（Times New Roman 8.5磅 斜体 灰色）
        #     一并卷入，导致合规文档中其他区域的合法不同字体被误判为失败；
        #     也未按第 1–4 页范围裁剪，会把答案页(第5页)正文误计入。
        #   · 新范围：仅遍历 9 个大题的【模块表】(module_tables 的前 9 张，
        #     内容即第1至第4页题目正文)；对每个模块内部：
        #       · 跳过含"X、"标记的中文标题段；
        #       · 跳过拼音段（含拉丁/带调元音音节但不含中文）；
        #       · 其余含中文的段落 = 题目正文段。
        #     这样自然排除顶部标题表、提示框、成绩表、页眉页脚（后者本就在
        #     header/footer part，不在 body）、答案页(不在 modules[:9])。
        rule_name = "第1页至第4页题目正文：Noto Serif CJK SC 五号 黑色 单倍行距 段前0段后0"
        if self.inspector.document is None:
            self.add_d2("D2.9", rule_name, 1, False, ["文档未解析"])
            return

        modules = self.module_tables()
        if len(modules) < 9:
            self.add_d2("D2.9", rule_name, 1, False,
                        [f"仅识别到 {len(modules)}/9 个大题模块，无法框定题目正文范围"])
            return

        latin_re = re.compile(r"[A-Za-zāáǎàēéěèīíǐìōóǒòūúǔùǖǘǚǜü]")
        tags = tuple(EXPECTED_MODULES)

        def is_title_para(text: str) -> bool:
            """含'X、'即视为大题中文标题段。"""
            return any(tag in text for tag in tags)

        def is_pinyin_para(text: str) -> bool:
            """段落含拉丁/带调元音音节，且不含中文 -> 拼音装饰段。"""
            return bool(latin_re.search(text)) and not CHINESE_RE.search(text)

        font_bad: list[str] = []
        size_bad: list[str] = []
        color_bad: list[str] = []
        spacing_bad: list[str] = []
        total_cn_runs = 0
        total_cn_paras = 0

        for t in modules[:9]:
            for p in t.el.findall(f".//{{{W}}}p"):
                ptext = element_text(p)
                if not CHINESE_RE.search(ptext):
                    continue
                if is_title_para(ptext) or is_pinyin_para(ptext):
                    continue
                total_cn_paras += 1

                # 段前/段后/行距。
                sp = p.find(f"{{{W}}}pPr/{{{W}}}spacing")
                before = to_int(wattr(sp, "before"), 0) if sp is not None else 0
                after = to_int(wattr(sp, "after"), 0) if sp is not None else 0
                line = to_int(wattr(sp, "line"), 0) if sp is not None else 0
                rule = wattr(sp, "lineRule", "") if sp is not None else ""
                # 段前=0、段后=0；单倍行距：line==240 与 rule∈{"auto","atLeast",""} 视为单倍。
                single_line = (line in (0, 240)) and rule in ("", "auto")
                if before != 0 or after != 0 or not single_line:
                    spacing_bad.append(f"before={before}/after={after}/line={line}/rule={rule} text='{ptext[:20]}'")

                for r in p.findall(f".//{{{W}}}r"):
                    rt = "".join(x.text or "" for x in r.findall(f".//{{{W}}}t"))
                    if not CHINESE_RE.search(rt):
                        continue
                    total_cn_runs += 1
                    rfonts = r.find(f"{{{W}}}rPr/{{{W}}}rFonts")
                    sz = r.find(f"{{{W}}}rPr/{{{W}}}sz")
                    color = r.find(f"{{{W}}}rPr/{{{W}}}color")
                    fonts_v: list[str] = []
                    if rfonts is not None:
                        for k in ("ascii", "hAnsi", "eastAsia"):
                            v = wattr(rfonts, k)
                            if v:
                                fonts_v.append(v)
                    if "Noto Serif CJK SC" not in fonts_v:
                        font_bad.append(f"fonts={fonts_v} text='{rt[:16]}'")
                    if to_int(wattr(sz, "val"), 0) != 21:
                        size_bad.append(f"sz={wattr(sz, 'val')} text='{rt[:16]}'")
                    if not is_black(wattr(color, "val")):
                        color_bad.append(f"color={wattr(color, 'val')} text='{rt[:16]}'")

        ok = (
            total_cn_runs > 0
            and not font_bad
            and not size_bad
            and not color_bad
            and not spacing_bad
        )
        self.add_d2(
            "D2.9",
            rule_name,
            1,
            ok,
            [
                f"范围=9个大题模块内除标题/拼音段外的中文段/run",
                f"中文段={total_cn_paras} 中文run={total_cn_runs}",
                f"font_bad={len(font_bad)} sample={font_bad[:3]}",
                f"size_bad={len(size_bad)} sample={size_bad[:3]}",
                f"color_bad={len(color_bad)} sample={color_bad[:3]}",
                f"spacing_bad={len(spacing_bad)} sample={spacing_bad[:3]}",
            ],
        )

    def check_footer_pages(self) -> None:
        # 细则：第1页至第4页页脚居中显示「第1页/共4页」…「第4页/共4页」；
        # 字体为宋体或 Noto Serif CJK SC，8.5 磅（17 半磅），灰色。
        #
        # 关键判据（相对旧实现的修正）：
        #   · 旧实现正则 `第\d+页/共\d+页` 放任任意总页数 Y，且"任一 footer part
        #     通过即可"，无法证明第1-4页都正确，也不稳健于"域无缓存文本"的情形。
        #   · 新实现：
        #       (1) 解析 PAGE / NUMPAGES / SECTIONPAGES 域（w:instrText 与
        #           w:fldSimple@w:instr）与缓存文本，双通道识别页码结构；
        #       (2) 明确要求显示总页数 = 4（"共4页"）：优先取缓存文本中"共N页"
        #           的 N；若域无缓存数字，则要求总数来源为 SECTIONPAGES 域
        #           （分节页数，可为 4）而非 NUMPAGES（全文档页数，本卷含答案页
        #           通常为 5，渲染即非 4），否则判为无法确认 4 而失败；
        #       (3) 所有【被引用且非空】的 footer part 都必须合规（而非任一通过），
        #           以覆盖第1-4页。
        rule_name = "第1页至第4页页脚页码：居中「第X页/共4页」(共4页)，宋体或 Noto Serif CJK SC，8.5磅，灰色"
        footers = self.inspector.footers
        if not footers:
            self.add_d2("D2.10", rule_name, 1, False, ["未发现页脚"])
            return

        def field_instr(footer: ET.Element) -> str:
            parts = [e.text or "" for e in footer.findall(f".//{{{W}}}instrText")]
            parts += [wattr(f, "instr") or "" for f in footer.findall(f".//{{{W}}}fldSimple")]
            return " ".join(parts).upper()

        nonempty_seen = 0
        all_ok = True
        evidence: list[str] = []

        for idx, footer in enumerate(footers, 1):
            cached = re.sub(r"\s+", "", element_text(footer))
            instr = field_instr(footer)
            if not cached and not instr:
                evidence.append(f"footer{idx}: 空(未引用)")
                continue
            nonempty_seen += 1

            # 结构：第[页码]页/共[总数]页（页码/总数可能是域，缓存中或有或无数字）
            struct_ok = re.search(r"第\d*页/共\d*页", cached) is not None

            # 页码来源：缓存中"第N页"的 N，或 PAGE 域（\bPAGE\b 避免匹配 NUMPAGES）。
            has_page_field = re.search(r"\bPAGE\b", instr) is not None
            m_page = re.search(r"第(\d+)页", cached)
            page_src_ok = (m_page is not None) or has_page_field

            # 总数来源与取值：优先缓存"共N页"；否则依赖分节页数域。
            has_numpages = re.search(r"\bNUMPAGES\b", instr) is not None
            has_sectionpages = re.search(r"\bSECTIONPAGES\b", instr) is not None
            m_total = re.search(r"共(\d+)页", cached)
            if m_total is not None:
                total_ok = int(m_total.group(1)) == 4
                total_src = f"缓存共{m_total.group(1)}页"
            elif has_sectionpages:
                total_ok = True  # 分节页数域，可渲染为4；静态无法读数，按结构接受
                total_src = "SECTIONPAGES域(分节页数,按4接受)"
            elif has_numpages:
                total_ok = False  # 全文档域含答案页通常=5，非4
                total_src = "NUMPAGES域无缓存,无法确认共4页(全档含答案页多为5)"
            else:
                total_ok = False
                total_src = "无共N页文本也无页数域"

            # 居中：所有含文本段落 pPr/jc == center。
            paras = footer.findall(f".//{{{W}}}p")
            paras_with_text = [p for p in paras if element_text(p).strip()]
            centered_ok = bool(paras_with_text) and all(
                wattr(p.find(f"{{{W}}}pPr/{{{W}}}jc"), "val") == "center" for p in paras_with_text
            )

            # 字体/字号/颜色：所有含文本 run 必须满足。
            font_ok = True
            size_ok = True
            color_ok = True
            runs_seen = 0
            for r in footer.findall(f".//{{{W}}}r"):
                rt = "".join(x.text or "" for x in r.findall(f".//{{{W}}}t"))
                if not rt.strip():
                    continue
                runs_seen += 1
                rfonts = r.find(f"{{{W}}}rPr/{{{W}}}rFonts")
                sz = r.find(f"{{{W}}}rPr/{{{W}}}sz")
                color = r.find(f"{{{W}}}rPr/{{{W}}}color")
                fonts_v: list[str] = []
                if rfonts is not None:
                    for k in ("ascii", "hAnsi", "eastAsia"):
                        v = wattr(rfonts, k)
                        if v:
                            fonts_v.append(v)
                if not any(f in ("SimSun", "宋体", "Noto Serif CJK SC") for f in fonts_v):
                    font_ok = False
                if to_int(wattr(sz, "val"), 0) != 17:
                    size_ok = False
                if not is_gray(wattr(color, "val")):
                    color_ok = False

            footer_ok = (struct_ok and page_src_ok and total_ok and centered_ok
                         and runs_seen > 0 and font_ok and size_ok and color_ok)
            all_ok = all_ok and footer_ok
            evidence.append(
                f"footer{idx}: text='{cached}' struct={struct_ok} page_src={page_src_ok} "
                + f"total_ok={total_ok}({total_src}) centered={centered_ok} "
                + f"font_ok={font_ok} size_ok={size_ok} color_ok={color_ok}"
            )

        ok = nonempty_seen > 0 and all_ok
        self.add_d2("D2.10", rule_name, 1, ok, [f"非空页脚part={nonempty_seen}"] + evidence)

    def check_page_distribution(self) -> None:
        # 细则：
        # 第1页：第一至第三大题；
        # 第2页：第四至第六大题；
        # 第3页：第七大题 + 阅读文章前半部分（即第八大题的前半部分）；
        # 第4页：阅读后半部分（第八大题剩余）+ 第九大题。
        # 「页数按办公软件实际页数」——静态依据显式分页符切分推断。
        evidence: list[str] = []

        # 仅看 1-4 页分布，不限制总页数。
        segments = self.inspector.page_segments_by_breaks()
        evidence.append(f"显式分页段数={len(segments)}")

        expectations = [
            (0, ["一、", "二、", "三、"], "第1页：一、二、三 大题"),
            (1, ["四、", "五、", "六、"], "第2页：四、五、六 大题"),
            (2, ["七、", "八、"], "第3页：七 大题 + 八(阅读前半)"),
            (3, ["八、", "九、"], "第4页：八(阅读后半) + 九 大题"),
        ]
        ok_parts: list[bool] = []
        for idx, keys, desc in expectations:
            if idx < len(segments):
                seg = segments[idx]
                found = [k for k in keys if k in seg]
                ok_parts.append(len(found) == len(keys))
                evidence.append(f"{desc}: 命中{found}")
            else:
                ok_parts.append(False)
                evidence.append(f"{desc}: 缺失")

        ok = len(segments) >= 4 and all(ok_parts)
        self.add_d2(
            "D2.11",
            "第1页至第4页分页顺序：1页(一二三) / 2页(四五六) / 3页(七+阅读前半) / 4页(阅读后半+九)",
            3,
            ok,
            evidence,
        )

    @property
    def final_score(self) -> int:
        if not self.d1_passed:
            return 0
        return sum(r.score for r in self.dimension2)

    @property
    def max_score(self) -> int:
        return 21

    def dim1_reason(self) -> str:
        """维度一未通过时的原因说明；通过时为空字符串。"""
        if self.d1_passed:
            return ""
        return "；".join(f"{r.name}：{'；'.join(r.evidence)}" for r in self.dimension1 if not r.passed)

    def dim2_items(self) -> list[dict[str, object]]:
        """维度二逐项结果，命中与未命中均登记，供 Excel 矩阵对比。"""
        return [
            {
                "rule": r.name,
                "max_delta": r.points,
                "delta": r.score,
                "hit": r.passed,
                "detail": "",
            }
            for r in self.dimension2
        ]

    def as_dict(self) -> dict[str, object]:
        return {
            "file": str(self.path),
            "dimension1_passed": self.d1_passed,
            "dimension1": [r.__dict__ for r in self.dimension1],
            "dimension2_skipped": not self.d1_passed,
            "dimension2": [r.__dict__ for r in self.dimension2],
            "final_score": self.final_score,
            "max_score": self.max_score,
        }

    def print_report(self) -> None:
        if not self.d1_passed:
            print("维度一：不通过")
            print("维度二：跳过")
            print(f"最终得分：0 / {self.max_score}")
            return

        print("维度一：通过")
        print("维度二：评分结果")
        for r in self.dimension2:
            if r.passed:
                print(f"+{r.points}：{r.name}")
        print(f"最终得分：{self.final_score} / {self.max_score}")


def cell_fill(cell: Optional[ET.Element]) -> Optional[str]:
    if cell is None:
        return None
    shd = cell.find(f".//{{{W}}}shd")
    fill = wattr(shd, "fill")
    return fill.upper() if fill else None


def _locate_docx(dir_path: str) -> Optional[Path]:
    """在给定目录中定位待评估的 .docx 文件。

    约定：批量 runner 传入"脚本所在目录路径"，脚本自己在该目录下找唯一的 docx。
    过滤 Office 打开时留下的临时锁文件（以 ~$ 开头）。目录内若存在多个 .docx，
    优先取不含"备份/副本/backup/copy"关键字的一个。
    """
    directory = Path(dir_path)
    if not directory.is_dir():
        return None
    candidates = [p for p in directory.iterdir() if p.suffix.lower() == ".docx" and not p.name.startswith("~$")]
    if not candidates:
        return None
    if len(candidates) == 1:
        return candidates[0]

    def prefer(p: Path) -> int:
        lower = p.name.lower()
        return 0 if any(k in lower or k in p.name for k in ("备份", "副本", "backup", "copy")) else 1

    candidates.sort(key=lambda p: (-prefer(p), len(p.name)))
    return candidates[0]


def evaluate(dir_path: str) -> dict[str, object]:
    """统一评估入口：传入脚本所在目录路径，脚本自行在该目录中定位并评估 docx。

    返回结构见《脚本接口差异与统一建议.md》§2.2。
    发生异常时返回 status="error"，不抛出，不 print 主结果。
    """
    max_score = 21
    docx_path = _locate_docx(dir_path)
    if docx_path is None:
        return {
            "id": SCRIPT_ID,
            "file_name": "",
            "status": "error",
            "error": f"未在目录中找到 .docx 文件: {dir_path!r}",
            "dim1_pass": False,
            "dim1_reason": "",
            "dim2_items": [],
            "total_score": 0,
            "max_score": max_score,
        }

    try:
        evaluator = RubricEvaluator(docx_path)
        evaluator.evaluate()
        return {
            "id": SCRIPT_ID,
            "file_name": docx_path.name,
            "status": "ok",
            "error": None,
            "dim1_pass": evaluator.d1_passed,
            "dim1_reason": evaluator.dim1_reason(),
            "dim2_items": evaluator.dim2_items(),
            "total_score": evaluator.final_score,
            "max_score": evaluator.max_score,
        }
    except Exception as exc:  # noqa: BLE001 - 顶层兜底，脚本崩溃需转成 status="error"
        return {
            "id": SCRIPT_ID,
            "file_name": docx_path.name,
            "status": "error",
            "error": f"{type(exc).__name__}: {exc}",
            "dim1_pass": False,
            "dim1_reason": "",
            "dim2_items": [],
            "total_score": 0,
            "max_score": max_score,
        }


if __name__ == "__main__":
    # 仅用于本地调试：从命令行读取目录路径并打印 JSON 结果；
    # 未传参时默认使用脚本自身所在目录。
    _dir = sys.argv[1] if len(sys.argv) > 1 else str(Path(__file__).resolve().parent)
    _payload = json.dumps(evaluate(_dir), ensure_ascii=False, indent=2)
    _ = sys.stdout.buffer.write((_payload + "\n").encode("utf-8"))
