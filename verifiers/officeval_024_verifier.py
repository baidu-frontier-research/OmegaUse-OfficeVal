# -*- coding: utf-8 -*-
import json
import math
import os
import re
import struct
import sys
import zipfile
from collections import Counter, defaultdict
from dataclasses import dataclass, field
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Tuple
from xml.etree import ElementTree as ET

# 说明：按接口统一约定，本模块不修改全局 sys.stdout，也不 print 主结果。
# 主结果统一由 evaluate(dir_path) -> dict 返回；本文件仅在
# __main__ 下作为本地调试入口，才把 dict 序列化为 JSON 打印。

SCRIPT_ID = "024"
CM_TO_TWIPS = 567
EMU_PER_CM = 360000
EMU_PER_INCH = 914400
TWIPS_PER_INCH = 1440

NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "pic": "http://schemas.openxmlformats.org/drawingml/2006/picture",
    "v": "urn:schemas-microsoft-com:vml",
    "o": "urn:schemas-microsoft-com:office:office",
}

CHINESE_RE = re.compile(r"[\u4e00-\u9fff]")
ENGLISH_RE = re.compile(r"[A-Za-z]")
CITATION_RE = re.compile(r"[\[【](\d+(?:\s*[,，、-]\s*\d+)*)[\]】]")
BAD_CITATION_RE = re.compile(r"(?<![\[【])(?<!\d)(\d{1,2})(?:\s*[,，、-]\s*\d{1,2})?(?![\]】])(?!\d)")
REF_ITEM_RE = re.compile(r"^\s*(\d+)\s*[.\．\s]\s*\S+")

TABLE_TITLES = [
    "研究变量与资料来源", "问卷样本人口学特征", "社区绿地可达性等级划分", "问卷信度与效度检验", "步行行为描述性统计",
    "社区建成环境指标统计", "健康感知回归模型结果", "路径模型标准化系数", "稳健性检验结果", "社区更新策略矩阵",
]
FIGURE_TITLES = [
    "研究技术路线图", "多源数据处理流程质量控制图", "样区绿地可达性格网分布图", "居民休闲步行频率分布图",
    "绿地可达性与步行时长散点图", "步行行为影响因素模型系数图", "绿地使用体验评价雷达图", "健康感知路径模型图",
]
SECTION_REF_REQUIREMENTS = [
    # (组名, 组内子章节标题列表, 必须包含的引用编号集合;
    #  若最后一项是字符串 "and_after"，表示"及后续参考文献交叉引用"——所有 ≥ min(required) 的编号)
    ("前言", ["前言"], {1, 2}, False),
    ("研究方法章", ["研究设计与资料来源", "指标体系与变量设定", "空间数据整理",
                    "问卷与访谈资料", "统计分析方法"], {3, 4, 5, 6}, False),
    ("分析与讨论章", ["居民步行行为分析", "健康感知与路径模型", "稳健性与分层分析", "讨论"],
                    {7, 8}, True),  # True 表示"及后续": 需覆盖 {7,8,...,最大编号}
]

@dataclass
class Hit:
    score: int
    rule: str
    passed: bool
    detail: str

@dataclass
class WordInfo:
    path: Path
    is_docx: bool = False
    zip_ok: bool = False
    package_files: List[str] = field(default_factory=list)
    xml: Dict[str, ET.Element] = field(default_factory=dict)
    rels: Dict[str, str] = field(default_factory=dict)
    paragraphs: List[ET.Element] = field(default_factory=list)
    tables: List[ET.Element] = field(default_factory=list)
    body_blocks: List[Tuple[str, ET.Element]] = field(default_factory=list)
    texts: List[str] = field(default_factory=list)
    page_width_twips: int = 11906
    page_height_twips: int = 16838
    margin_left_twips: int = 1800
    margin_right_twips: int = 1800
    margin_top_twips: int = 1440
    margin_bottom_twips: int = 1440
    com: Dict[str, object] = field(default_factory=dict)
    image_sizes: Dict[str, Tuple[int, int]] = field(default_factory=dict)

    @property
    def text_width_twips(self) -> int:
        return max(1, self.page_width_twips - self.margin_left_twips - self.margin_right_twips)

    @property
    def printable_width_emu(self) -> int:
        return int(self.text_width_twips / TWIPS_PER_INCH * EMU_PER_INCH)

    @property
    def text_height_twips(self) -> int:
        return max(1, self.page_height_twips - self.margin_top_twips - self.margin_bottom_twips)


def qn(name: str) -> str:
    prefix, local = name.split(":", 1)
    return "{%s}%s" % (NS[prefix], local)


def attr(el: Optional[ET.Element], name: str, default=None):
    if el is None:
        return default
    return el.attrib.get(qn(name), default)


def all_text(el: Optional[ET.Element]) -> str:
    if el is None:
        return ""
    return "".join(t.text or "" for t in el.findall(".//w:t", NS))


def child(el: Optional[ET.Element], path: str) -> Optional[ET.Element]:
    return None if el is None else el.find(path, NS)


def children(el: Optional[ET.Element], path: str) -> List[ET.Element]:
    return [] if el is None else el.findall(path, NS)


def int_attr(el: Optional[ET.Element], name: str, default: int = 0) -> int:
    try:
        return int(attr(el, name, default))
    except Exception:
        return default


def load_word(path: Path) -> WordInfo:
    # 本 verifier 仅识别 .docx；.doc 由目录级 _locate_docx 过滤，不再进入解析。
    info = WordInfo(path=path, is_docx=path.suffix.lower() == ".docx")
    if not path.exists():
        return info
    if not info.is_docx:
        return info
    try:
        with zipfile.ZipFile(path) as zf:
            info.zip_ok = True
            info.package_files = zf.namelist()
            for name in info.package_files:
                if name.endswith(".xml") and (name.startswith("word/") or name == "[Content_Types].xml"):
                    try:
                        info.xml[name] = ET.fromstring(zf.read(name))
                    except Exception:
                        pass
            rel_path = "word/_rels/document.xml.rels"
            if rel_path in info.package_files:
                rel_root = ET.fromstring(zf.read(rel_path))
                for rel in rel_root:
                    rid = rel.attrib.get("Id")
                    target = rel.attrib.get("Target", "")
                    if rid:
                        info.rels[rid] = target
            for name in info.package_files:
                if name.startswith("word/media/"):
                    data = zf.read(name)
                    size = image_pixel_size(data)
                    if size:
                        info.image_sizes[name] = size
    except Exception:
        return info
    doc = info.xml.get("word/document.xml")
    body = child(doc, "w:body")
    if body is not None:
        for item in list(body):
            if item.tag == qn("w:p"):
                info.paragraphs.append(item)
                info.body_blocks.append(("p", item))
                info.texts.append(all_text(item))
            elif item.tag == qn("w:tbl"):
                info.tables.append(item)
                info.body_blocks.append(("tbl", item))
                info.texts.append(all_text(item))
        sect = body.find(".//w:sectPr", NS)
        pg_sz = child(sect, "w:pgSz")
        pg_mar = child(sect, "w:pgMar")
        info.page_width_twips = int_attr(pg_sz, "w:w", info.page_width_twips)
        info.page_height_twips = int_attr(pg_sz, "w:h", info.page_height_twips)
        info.margin_left_twips = int_attr(pg_mar, "w:left", info.margin_left_twips)
        info.margin_right_twips = int_attr(pg_mar, "w:right", info.margin_right_twips)
        info.margin_top_twips = int_attr(pg_mar, "w:top", info.margin_top_twips)
        info.margin_bottom_twips = int_attr(pg_mar, "w:bottom", info.margin_bottom_twips)
    info.com = inspect_with_word_com(path)
    return info


def inspect_with_word_com(path: Path) -> Dict[str, object]:
    """本 verifier 固定使用 OOXML 普通解析，不启动 Word COM。

    历史上此函数曾经用 win32com.client / pythoncom 打开 Word 采集页数、页码、
    表格行高、段落垂直坐标等信息；后按"非必要不用 COM"约束整体禁用，仅保留
    return {"available": False} 的空实现。之前 return 后残留的 ~300 行 COM 死代
    码已删除，`info.com` 恒为 {"available": False}，下游 info.com.get(...) 全部
    回退到 OOXML 近似路径（已在 image_pixel_size / _para_default_size_hp / 表格
    行高 / 段落定位等分支里实现）。
    """
    del path  # 保留形参以维持调用点签名不变
    return {"available": False}


def image_pixel_size(data: bytes) -> Optional[Tuple[int, int]]:
    if data.startswith(b"\x89PNG\r\n\x1a\n") and len(data) >= 24:
        return struct.unpack(">II", data[16:24])
    if data.startswith(b"\xff\xd8"):
        i = 2
        while i + 9 < len(data):
            if data[i] != 0xFF:
                i += 1
                continue
            marker = data[i + 1]
            length = int.from_bytes(data[i + 2:i + 4], "big")
            if marker in {0xC0, 0xC1, 0xC2, 0xC3, 0xC5, 0xC6, 0xC7, 0xC9, 0xCA, 0xCB, 0xCD, 0xCE, 0xCF}:
                return int.from_bytes(data[i + 7:i + 9], "big"), int.from_bytes(data[i + 5:i + 7], "big")
            i += 2 + length
    if data.startswith(b"GIF") and len(data) >= 10:
        return struct.unpack("<HH", data[6:10])
    return None


def run_props(run: ET.Element) -> ET.Element:
    props = child(run, "w:rPr")
    return props if props is not None else ET.Element("empty")


def run_font_names(run: ET.Element) -> Dict[str, str]:
    fonts = child(run_props(run), "w:rFonts")
    return {k: attr(fonts, "w:" + k, "") for k in ["ascii", "hAnsi", "eastAsia", "cs"]}


def run_size_half_points(run: ET.Element) -> Optional[int]:
    size = child(run_props(run), "w:sz")
    val = attr(size, "w:val")
    try:
        return int(val) if val is not None else None
    except Exception:
        return None


def run_is_bold(run: ET.Element) -> bool:
    return child(run_props(run), "w:b") is not None


def run_is_superscript(run: ET.Element) -> bool:
    return attr(child(run_props(run), "w:vertAlign"), "w:val") == "superscript"


def para_jc(paragraph: ET.Element) -> str:
    return attr(child(child(paragraph, "w:pPr"), "w:jc"), "w:val", "")


def para_spacing(paragraph: ET.Element) -> Dict[str, str]:
    sp = child(child(paragraph, "w:pPr"), "w:spacing")
    return {
        "before": attr(sp, "w:before", ""),
        "after": attr(sp, "w:after", ""),
        "beforeLines": attr(sp, "w:beforeLines", ""),
        "afterLines": attr(sp, "w:afterLines", ""),
        "line": attr(sp, "w:line", ""),
        "lineRule": attr(sp, "w:lineRule", ""),
    }


def paragraph_runs(paragraph: ET.Element) -> List[ET.Element]:
    return paragraph.findall(".//w:r", NS)


def table_rows(table: ET.Element) -> List[ET.Element]:
    return children(table, "w:tr")


def table_cells(table: ET.Element) -> List[ET.Element]:
    return table.findall(".//w:tc", NS)


def table_width_twips(table: ET.Element) -> Optional[int]:
    w_el = child(child(table, "w:tblPr"), "w:tblW")
    val = attr(w_el, "w:w")
    typ = attr(w_el, "w:type")
    try:
        value = int(val)
    except Exception:
        return None
    if typ == "pct":
        return None
    return value


def table_align(table: ET.Element) -> str:
    return attr(child(child(table, "w:tblPr"), "w:jc"), "w:val", "")


def table_indent_twips(table: ET.Element) -> int:
    """读取 w:tblInd（表左缩进）。type=dxa 时按 twips；其它类型（pct/nil）视为 0。"""
    ind_el = child(child(table, "w:tblPr"), "w:tblInd")
    if ind_el is None:
        return 0
    typ = attr(ind_el, "w:type", "dxa")
    if typ != "dxa":
        return 0
    try:
        return int(attr(ind_el, "w:w", 0))
    except Exception:
        return 0


def table_effective_width_twips(table: ET.Element, text_width_twips: int) -> Optional[int]:
    """解析 w:tblW：dxa 直接返回；pct 换算为 twips（5000=100%）；auto/nil 用 tblGrid 合计。
    与 Word 渲染宽度一致，用于办公软件下的实际宽度校验。"""
    w_el = child(child(table, "w:tblPr"), "w:tblW")
    val = attr(w_el, "w:w")
    typ = attr(w_el, "w:type", "dxa")
    try:
        value = int(val) if val is not None else None
    except Exception:
        value = None
    if typ == "pct" and value is not None:
        return int(text_width_twips * value / 5000)
    if typ == "dxa" and value is not None:
        return value
    # auto / nil / 缺失：以 tblGrid 列宽合计推断
    grid_cols = table.findall("./w:tblGrid/w:gridCol", NS)
    total = sum(int_attr(gc, "w:w", 0) for gc in grid_cols)
    return total if total > 0 else (value if typ == "dxa" else None)


def find_numbered_tables(info: "WordInfo", table_captions: Dict[int, int]) -> Dict[int, ET.Element]:
    """将"表 N ..."题注段落绑定到其后紧邻的第一个 w:tbl 块（跳过空段/图题）。"""
    result: Dict[int, ET.Element] = {}
    for num, cap_idx in table_captions.items():
        for i in range(cap_idx + 1, min(cap_idx + 6, len(info.body_blocks))):
            if info.body_blocks[i][0] == "tbl":
                result[num] = info.body_blocks[i][1]
                break
    return result


def caption_titles_in_order(info: "WordInfo", table_captions: Dict[int, int], titles: List[str]) -> bool:
    """依次校验：表1..表10 的题注段落文本按顺序包含 TABLE_TITLES 中的对应标题。"""
    if not all(n in table_captions for n in range(1, 11)):
        return False
    prev_idx = -1
    for n in range(1, 11):
        idx = table_captions[n]
        if idx <= prev_idx:
            return False
        cap_text = all_text(info.body_blocks[idx][1])
        if titles[n - 1] not in cap_text:
            return False
        prev_idx = idx
    return True


def border_spec(border: Optional[ET.Element]) -> Tuple[str, Optional[int]]:
    return attr(border, "w:val", ""), int_attr(border, "w:sz", -1)


def table_border_summary(table: ET.Element) -> Dict[str, Tuple[str, Optional[int]]]:
    borders = child(child(table, "w:tblPr"), "w:tblBorders")
    return {name: border_spec(child(borders, "w:" + name)) for name in ["top", "bottom", "left", "right", "insideH", "insideV"]}


def is_none_border(spec: Tuple[str, Optional[int]]) -> bool:
    return spec[0] in {"", "nil", "none"} or spec[1] == 0


def line_width_ok(spec: Tuple[str, Optional[int]], points: float, tol: int = 2) -> bool:
    expected = int(round(points * 8))
    return spec[0] not in {"", "nil", "none"} and spec[1] is not None and abs(spec[1] - expected) <= tol


def _cell_border(cell: ET.Element, name: str) -> Tuple[str, Optional[int]]:
    """读取单元格 w:tcBorders/w:<name>；缺失返回空。"""
    b = child(child(cell, "w:tcPr"), "w:tcBorders")
    return border_spec(child(b, "w:" + name)) if b is not None else ("", None)


def _cell_border_active(cell: ET.Element, name: str) -> bool:
    """单元格该方向是否画了线（非 nil/none/0，且样式非空）。"""
    val, sz = _cell_border(cell, name)
    return val not in {"", "nil", "none"} and (sz or 0) > 0


def header_bottom_line_ok(table: ET.Element) -> bool:
    """表头下方横线线宽=0.75磅（sz=6，1/8pt）。优先取首行单元格 w:bottom，回退到 tblBorders/insideH。"""
    rows = table_rows(table)
    if not rows:
        return False
    for cell in rows[0].findall(".//w:tc", NS):
        val, sz = _cell_border(cell, "bottom")
        if val not in {"", "nil", "none"} and sz is not None and abs(sz - 6) <= 1:
            return True
    borders = table_border_summary(table)
    val, sz = borders.get("insideH", ("", None))
    return val not in {"", "nil", "none"} and sz is not None and abs(sz - 6) <= 1


def is_cover_table(table: ET.Element) -> bool:
    text = all_text(table).replace(" ", "").replace("　", "")
    return "申请人" in text and "学科" in text and "指导教师" in text


def is_three_line_table(table: ET.Element) -> bool:
    """严格按细则校验三线表——办公软件（Word/WPS）实际渲染的效果：
    仅保留三条主横线：表格上边线、表头下方横线、表格下边线；
    上下边线线宽 1.5 磅（sz=12），表头下方横线 0.75 磅（sz=6）；
    无左右边线，无表内竖线（insideV），无除表头外的其他主横线（insideH）。"""
    rows = table_rows(table)
    if len(rows) < 2:
        return False
    borders = table_border_summary(table)

    # 1) 上边线：办公软件渲染优先级 = cell 级 tcBorders/top（若显式声明，含 nil）
    #    覆盖 tblBorders/top。任意首行单元格显式把 top 声明为 nil/none 时，
    #    办公软件渲染中该处上边线被抹掉——即便 tblBorders/top=single sz=12。
    def _top_1p5(spec: Tuple[str, Optional[int]]) -> bool:
        return spec[0] == "single" and spec[1] is not None and abs(spec[1] - 12) <= 1

    first_cells = rows[0].findall(".//w:tc", NS)
    first_cells_top = [_cell_border(c, "top") for c in first_cells]
    any_cell_top_declared = any(spec[0] not in {"", None} for spec in first_cells_top)
    if any_cell_top_declared:
        # 任一首行单元格声明了 top → 完全按 cell 级判定（全部 1.5pt 才算保留上边线）
        top_ok = bool(first_cells) and all(_top_1p5(spec) for spec in first_cells_top)
    else:
        top_ok = _top_1p5(borders.get("top", ("", None)))

    # 2) 下边线：末行 cell 级 tcBorders/bottom 优先于 tblBorders/bottom。
    def _bot_1p5(spec: Tuple[str, Optional[int]]) -> bool:
        return spec[0] == "single" and spec[1] is not None and abs(spec[1] - 12) <= 1

    last_cells = rows[-1].findall(".//w:tc", NS)
    last_cells_bot = [_cell_border(c, "bottom") for c in last_cells]
    any_cell_bot_declared = any(spec[0] not in {"", None} for spec in last_cells_bot)
    if any_cell_bot_declared:
        bottom_ok = bool(last_cells) and all(_bot_1p5(spec) for spec in last_cells_bot)
    else:
        bottom_ok = _bot_1p5(borders.get("bottom", ("", None)))

    # 3) 表头下方横线：首行所有单元格 bottom=0.75pt（sz=6，允许±1 抖动）
    header_cells = rows[0].findall(".//w:tc", NS)
    def _h_075(spec: Tuple[str, Optional[int]]) -> bool:
        return spec[0] == "single" and spec[1] is not None and abs(spec[1] - 6) <= 1
    header_bottom_ok = bool(header_cells) and all(
        _h_075(_cell_border(c, "bottom")) for c in header_cells
    )

    # 4) 无左右边线（tblBorders 与 cell 均不得实心）
    left_none = is_none_border(borders.get("left", ("", None)))
    right_none = is_none_border(borders.get("right", ("", None)))
    if left_none:
        for r in rows:
            cells = r.findall(".//w:tc", NS)
            if cells and _cell_border_active(cells[0], "left"):
                left_none = False
                break
    if right_none:
        for r in rows:
            cells = r.findall(".//w:tc", NS)
            if cells and _cell_border_active(cells[-1], "right"):
                right_none = False
                break

    # 5) 表内无竖线（insideV）——tblBorders 层与任意单元格 left/right 均不得画竖线
    inside_v_none = is_none_border(borders.get("insideV", ("", None)))
    if inside_v_none:
        for r in rows:
            cells = r.findall(".//w:tc", NS)
            for ci, c in enumerate(cells):
                if 0 < ci and _cell_border_active(c, "left"):
                    inside_v_none = False
                    break
                if ci < len(cells) - 1 and _cell_border_active(c, "right"):
                    inside_v_none = False
                    break
            if not inside_v_none:
                break

    # 6) 除表头下方外，表内不得出现其他主横线：
    #    tblBorders/insideH 不能实心；除首行 bottom(=0.75pt) 外，任何单元格
    #    top/bottom 均不得画线（末行 bottom 承担表格下边线，允许）。
    inside_h_none = is_none_border(borders.get("insideH", ("", None)))
    extra_h_ok = True
    if inside_h_none:
        for ri, r in enumerate(rows):
            cells = r.findall(".//w:tc", NS)
            for c in cells:
                # 首行 bottom 已由 header_bottom_ok 校验，跳过
                if ri == 0:
                    if _cell_border_active(c, "top"):
                        extra_h_ok = False
                        break
                    continue
                # 末行 bottom 承担表下边线
                if ri == len(rows) - 1:
                    if _cell_border_active(c, "top"):
                        extra_h_ok = False
                        break
                    continue
                # 中间行 top/bottom 都不得画线
                if _cell_border_active(c, "top") or _cell_border_active(c, "bottom"):
                    extra_h_ok = False
                    break
            if not extra_h_ok:
                break
    else:
        extra_h_ok = False

    return (
        top_ok and bottom_ok and header_bottom_ok
        and left_none and right_none
        and inside_v_none and extra_h_ok
    )


def table_row_heights_ok(table: ET.Element) -> bool:
    """静态 XML 层的行高判定——只能识别显式行高。当 trHeight 未设置时，
    Word/WPS 按内容自适应绘制，此处按"未强制约束高度即视为满足最小高度"处理，
    真正的裁切风险在 COM 通道结合实测行高判定。"""
    rows = table_rows(table)
    if not rows:
        return False
    min_twips = int(0.6 * CM_TO_TWIPS)  # 0.6cm = 340 twips
    for row in rows:
        h = child(child(row, "w:trPr"), "w:trHeight")
        val = int_attr(h, "w:val", 0)
        rule = attr(h, "w:hRule", "")
        if val == 0:
            # 未设置显式行高：Word 会按内容自适应，通常 >= 0.6cm；此处放行
            continue
        if rule == "exact" and val < min_twips:
            return False
        if rule in {"atLeast"} and val < min_twips:
            return False
    return True


def _row_cant_split(row: ET.Element) -> bool:
    """w:trPr/w:cantSplit 存在即允许行跨页拆分被禁止（不改变高度语义，仅记录）。"""
    return child(child(row, "w:trPr"), "w:cantSplit") is not None


def table_rows_heights_meet_min_cm(table: ET.Element, info: "WordInfo", table_index_in_doc: int) -> Tuple[bool, str]:
    """按细则校验：全文所有表格行高不低于 0.6cm。
    Word/WPS 视角优先使用 COM 采集的实际渲染行高（点，1cm=28.3465pt）；
    COM 不可用时回退到 XML 显式行高 + "未显式则视为自适应通过"。
    返回 (是否通过, 说明)。"""
    min_cm = 0.6
    min_pts = min_cm * 72.0 / 2.54  # ≈ 17.0079 pt
    tables_rh = info.com.get("tables_row_heights") if isinstance(info.com, dict) else None
    if isinstance(tables_rh, list) and 0 <= table_index_in_doc < len(tables_rh):
        per_rows = tables_rh[table_index_in_doc]
        if per_rows:
            for i, (h_pts, rule) in enumerate(per_rows):
                if h_pts is None or h_pts < 0:
                    continue  # 采集失败，忽略该行
                # COM 的 Height 报的是实际渲染点数，允许 0.5pt 舍入抖动
                if h_pts + 0.5 < min_pts:
                    return False, f"行{i+1}高度{h_pts:.2f}pt<{min_pts:.2f}pt(0.6cm), rule={rule}"
            return True, "COM渲染高度合格"
    # 回退：XML 显式行高
    return table_row_heights_ok(table), "XML显式行高判定"


def table_no_text_clipping(table: ET.Element, info: "WordInfo", table_index_in_doc: int) -> Tuple[bool, str]:
    """按细则校验：表格文字没有被上下边框裁切。
    Word/WPS 触发裁切的两个主要成因：
      1) w:trHeight 使用 hRule='exact' 固定行高，且该固定值小于文字所需高度；
      2) 单元格上下内边距（w:tcMar/top、bottom）为负；或行高被强制并覆盖内容。
    静态层判据：
      - 任一行 hRule='exact' 且高度小于 0.6cm 视为潜在裁切；
      - 任一行/表 tcMar 或 tblCellMar 的 top/bottom 为负值视为潜在裁切。
    动态层判据（COM 可用时）：
      - 若 COM 报告的实测 Height 小于文字五号最小行高（约 12.6pt = 0.44cm），
        且行 HeightRule=2(Exactly)，判定为裁切；rule=1/0 时不判定。"""
    # 静态层
    for row in table_rows(table):
        h = child(child(row, "w:trPr"), "w:trHeight")
        val = int_attr(h, "w:val", 0)
        rule = attr(h, "w:hRule", "")
        if rule == "exact" and val > 0 and val < int(0.6 * CM_TO_TWIPS):
            return False, f"存在 hRule=exact 且行高{val}<0.6cm"

    def _neg_margin(margins_el: Optional[ET.Element]) -> Optional[str]:
        if margins_el is None:
            return None
        for name in ("top", "bottom"):
            m = child(margins_el, "w:" + name)
            if m is None:
                continue
            try:
                v = int(attr(m, "w:w", 0))
            except Exception:
                v = 0
            if v < 0:
                return f"{name}边距={v}"
        return None

    tbl_cell_mar = child(child(table, "w:tblPr"), "w:tblCellMar")
    neg = _neg_margin(tbl_cell_mar)
    if neg:
        return False, f"tblCellMar {neg}"
    for tc in table.findall(".//w:tc", NS):
        tc_mar = child(child(tc, "w:tcPr"), "w:tcMar")
        neg = _neg_margin(tc_mar)
        if neg:
            return False, f"单元格 tcMar {neg}"

    # 动态层：结合 COM 实测行高
    tables_rh = info.com.get("tables_row_heights") if isinstance(info.com, dict) else None
    if isinstance(tables_rh, list) and 0 <= table_index_in_doc < len(tables_rh):
        per_rows = tables_rh[table_index_in_doc]
        # 五号(10.5pt)单倍行距最小需要 ≈12.6pt；再加最小上下边距余量，用 0.6cm(≈17pt)
        min_pts = 0.6 * 72.0 / 2.54
        for i, (h_pts, rule) in enumerate(per_rows or []):
            if h_pts is None or h_pts < 0:
                continue
            # 仅当 HeightRule=Exactly(2) 且实测低于 0.6cm 时判定为裁切
            if rule == 2 and h_pts + 0.5 < min_pts:
                return False, f"行{i+1} Exactly高度{h_pts:.2f}pt<0.6cm"
    return True, "OK"


def font_check_runs(runs: List[ET.Element], require_bold: Optional[bool] = None, size_hp: int = 21) -> Tuple[int, int]:
    checked = 0
    good = 0
    for run in runs:
        text = all_text(run)
        if not text.strip():
            continue
        checked += 1
        fonts = run_font_names(run)
        size = run_size_half_points(run)
        cn_ok = not CHINESE_RE.search(text) or fonts.get("eastAsia") in {"宋体", "SimSun", "simsun", ""}
        en_ok = not re.search(r"[A-Za-z0-9]", text) or fonts.get("ascii") in {"Times New Roman", ""} or fonts.get("hAnsi") in {"Times New Roman", ""}
        size_ok = size in {None, size_hp}
        bold_ok = True if require_bold is None else run_is_bold(run) == require_bold
        if cn_ok and en_ok and size_ok and bold_ok:
            good += 1
    return good, checked


def table_text_font_ok(table: ET.Element, header: bool) -> bool:
    """(legacy) 用于非表头正文字体校验；表头校验请改用 is_table_header_row_font_ok。"""
    rows = table_rows(table)
    selected = rows[:1] if header else rows[1:] or rows[:]
    good = checked = 0
    for row in selected:
        for run in row.findall(".//w:r", NS):
            g, c = font_check_runs([run], True if header else None, 21)
            good += g
            checked += c
    return checked > 0 and good / checked >= 0.85


def _run_effective_bold(run: ET.Element) -> bool:
    """Word/WPS 渲染是否加粗：单元格中 rPr/w:b 存在且未显式 val='0'/'false'。"""
    b = child(run_props(run), "w:b")
    if b is None:
        return False
    val = attr(b, "w:val")
    return val not in {"0", "false"}


def _run_east_asia_font(run: ET.Element) -> str:
    fonts = run_font_names(run)
    ea = fonts.get("eastAsia", "") or ""
    return ea.strip()


def _run_ascii_font(run: ET.Element) -> str:
    fonts = run_font_names(run)
    # Word 中英数字符使用 w:ascii（0x20-0x7F）与 w:hAnsi（拉丁扩展），只要有一个满足即可
    for key in ("ascii", "hAnsi"):
        name = (fonts.get(key) or "").strip()
        if name:
            return name
    return ""


def _para_default_size_hp(paragraph: ET.Element) -> Optional[int]:
    """段落标记(paragraph mark)的 rPr/w:sz —— Word/WPS 里未显式设字号的 run
    继承该段落默认字号。取 w:pPr/w:rPr/w:sz@val（半磅）；无则 None。"""
    pPr = child(paragraph, "w:pPr")
    if pPr is None:
        return None
    rPr = child(pPr, "w:rPr")
    if rPr is None:
        return None
    val = attr(child(rPr, "w:sz"), "w:val")
    try:
        return int(val) if val is not None else None
    except Exception:
        return None


def _run_effective_size_hp(run: ET.Element, paragraph: ET.Element) -> Optional[int]:
    """run 在办公软件里的生效字号(半磅): run 直设 w:sz → 段落默认字号 → None.
    (样式链更深层未解析——办公软件里 run 无 sz 时先落到段落默认, 已覆盖绝大多数文档.)"""
    s = run_size_half_points(run)
    if s is not None:
        return s
    return _para_default_size_hp(paragraph)


def _para_chinese_size_hp(paragraph: ET.Element) -> Optional[int]:
    """段落里中文字符的生效字号(半磅)——取首个含中文且能定出字号的 run;
    无中文 run 时退到段落默认字号; 再无则 None (该段无中文参照)."""
    for run in paragraph_runs(paragraph):
        if CHINESE_RE.search(all_text(run)):
            s = _run_effective_size_hp(run, paragraph)
            if s is not None:
                return s
    return _para_default_size_hp(paragraph)


def _is_songti(name: str) -> bool:
    return name in {"宋体", "SimSun", "simsun", "NSimSun", "PMingLiU", "MingLiU"}


def is_table_header_row_font_ok(table: ET.Element) -> Tuple[bool, str]:
    """按细则校验表格表头（首行）字体——办公软件（Word/WPS）实际渲染依据：
       1) 表头文字加粗（rPr/w:b 存在且未 val='0'）
       2) 中文字符：eastAsia = 宋体，字号 = 五号（w:sz=21 半磅）
       3) 英文字母与阿拉伯数字：ascii/hAnsi = Times New Roman，字号 = 五号（21）
    未含中文的 run 不做中文字体判断，未含英文/数字的 run 不做英数字体判断；
    空白 run 忽略。任一非空 run 违反其一即判不通过。"""
    rows = table_rows(table)
    if not rows:
        return False, "无行"
    header_runs = rows[0].findall(".//w:r", NS)
    non_empty = 0
    for run in header_runs:
        text = all_text(run)
        if not text.strip():
            continue
        non_empty += 1

        # 1) 加粗
        if not _run_effective_bold(run):
            return False, f"表头未加粗: {text!r}"

        # 五号 = 10.5pt = 21 半磅
        size = run_size_half_points(run)
        has_cn = bool(CHINESE_RE.search(text))
        has_en = bool(re.search(r"[A-Za-z0-9]", text))

        # 2) 中文：宋体 + 五号
        if has_cn:
            ea = _run_east_asia_font(run)
            if not _is_songti(ea):
                return False, f"表头中文非宋体: {text!r} eastAsia={ea!r}"
            if size != 21:
                return False, f"表头中文非五号: {text!r} sz={size}"

        # 3) 英文/数字：Times New Roman + 五号
        if has_en:
            asc = _run_ascii_font(run)
            if asc != "Times New Roman":
                return False, f"表头英/数非Times New Roman: {text!r} ascii={asc!r}"
            if size != 21:
                return False, f"表头英/数非五号: {text!r} sz={size}"

    if non_empty == 0:
        return False, "首行无文字"
    return True, "OK"


def is_table_body_font_ok(table: ET.Element) -> Tuple[bool, str]:
    """按细则校验表格正文（除首行外的所有数据行）字体——办公软件（Word/WPS）
    实际渲染依据：
       1) 中文字符：eastAsia = 宋体，字号 = 五号（w:sz=21 半磅）
       2) 英文字母与阿拉伯数字：ascii/hAnsi = Times New Roman，字号 = 五号（21）
    未含中文的 run 不做中文字体判断，未含英文/数字的 run 不做英数字体判断；
    空白 run 忽略。任一非空 run 违反其一即判不通过。细则未提到加粗，此处不约束。
    """
    rows = table_rows(table)
    if not rows:
        return False, "无行"
    body_rows = rows[1:] if len(rows) > 1 else []
    if not body_rows:
        return False, "无正文行"
    non_empty = 0
    for row in body_rows:
        for run in row.findall(".//w:r", NS):
            text = all_text(run)
            if not text.strip():
                continue
            non_empty += 1

            size = run_size_half_points(run)
            has_cn = bool(CHINESE_RE.search(text))
            has_en = bool(re.search(r"[A-Za-z0-9]", text))

            # 1) 中文：宋体 + 五号
            if has_cn:
                ea = _run_east_asia_font(run)
                if not _is_songti(ea):
                    return False, f"正文中文非宋体: {text!r} eastAsia={ea!r}"
                if size != 21:
                    return False, f"正文中文非五号: {text!r} sz={size}"

            # 2) 英文/数字：Times New Roman + 五号
            if has_en:
                asc = _run_ascii_font(run)
                if asc != "Times New Roman":
                    return False, f"正文英/数非Times New Roman: {text!r} ascii={asc!r}"
                if size != 21:
                    return False, f"正文英/数非五号: {text!r} sz={size}"

    if non_empty == 0:
        return False, "无正文文字"
    return True, "OK"


def find_caption_indices(info: WordInfo, prefix: str) -> Dict[int, int]:
    """定位每个"表 N ..."/"图 N ..."题注段落的索引；返回 {编号: 段落索引}。
    出现重复编号时取"位于其后紧邻表格/图片"的那一处（真正的题注，而非目录项）。"""
    captions: Dict[int, int] = {}
    # 允许中文全角空格作为分隔；数字后不得紧跟非空白/非文字（避免匹配 表9-1）
    pat = re.compile(rf"^\s*{prefix}\s*(\d+)(?=[\s　]|$)")
    candidates: Dict[int, List[int]] = {}
    for idx, (kind, block) in enumerate(info.body_blocks):
        if kind != "p":
            continue
        text = all_text(block).strip()
        match = pat.match(text)
        if not match:
            continue
        # 排除目录项：题注文本内含省略点或以页码结尾（如"..... 10"）
        if "..." in text or "……" in text or re.search(r"\.{3,}\s*\d+\s*$", text):
            continue
        candidates.setdefault(int(match.group(1)), []).append(idx)
    for num, idxs in candidates.items():
        # 优先取"其后紧邻块为对应表格/图片"的段落；没有则取最后一个（多为正文题注）
        chosen = None
        for i in idxs:
            for j in range(i + 1, min(i + 6, len(info.body_blocks))):
                kind, blk = info.body_blocks[j]
                if kind == "tbl" and prefix == "表":
                    chosen = i
                    break
                if kind == "p" and not all_text(blk).strip():
                    continue
                if kind == "p" and prefix == "图" and blk.find(".//w:drawing", NS) is not None:
                    chosen = i
                    break
                break
            if chosen is not None:
                break
        captions[num] = chosen if chosen is not None else idxs[-1]
    return captions


def block_texts(info: WordInfo) -> List[str]:
    return [all_text(block) for _, block in info.body_blocks]


def drawing_info(info: WordInfo) -> List[Dict[str, object]]:
    items = []
    for idx, (kind, block) in enumerate(info.body_blocks):
        for drawing in block.findall(".//w:drawing", NS):
            extent = drawing.find(".//wp:extent", NS)
            cx = int(extent.attrib.get("cx", "0")) if extent is not None else 0
            cy = int(extent.attrib.get("cy", "0")) if extent is not None else 0
            rid_el = drawing.find(".//a:blip", NS)
            rid = rid_el.attrib.get(qn("r:embed"), "") if rid_el is not None else ""
            target = info.rels.get(rid, "")
            media = "word/" + target.lstrip("/") if target and not target.startswith("word/") else target
            anchor = drawing.find(".//wp:anchor", NS)
            inline = drawing.find(".//wp:inline", NS)
            crop = drawing.find(".//a:srcRect", NS)
            wrap_none = drawing.find(".//wp:wrapNone", NS) is not None
            behind_doc = anchor.attrib.get("behindDoc", "0") if anchor is not None else "0"
            items.append({"block_index": idx, "paragraph": block if kind == "p" else None, "cx": cx, "cy": cy, "rid": rid, "media": media, "anchor": anchor is not None, "anchor_el": anchor, "inline": inline is not None, "crop": crop is not None, "wrap_none": wrap_none, "behind_doc": behind_doc, "element": drawing})
    return items


def _figure_center_deviation_twips(d: Dict[str, object], info: "WordInfo") -> Optional[int]:
    """图片左右边界相对正文版心居中的偏差（twips），取 |左距 - 右距|。
    针对办公软件 Word/WPS 的真实渲染路径：
      - 嵌入式图片 (wp:inline)：水平位置 = 段落 w:jc 与 w:ind 的联合结果。
          jc=center：偏差 = |左缩进 - 右缩进|；
          jc=right ：图片贴右边缘；
          其它（left/start/未设）：图片贴左边缘 + 段落左缩进。
      - 浮动图片 (wp:anchor)：水平位置由 wp:positionH 决定。
          wp:align=center 且 relativeFrom ∈ {margin, column}：偏差 = 0；
          wp:posOffset  ：按 offset EMU 值换算至 twips 后计算。
    无法从 XML 直接判定时返回 None。"""
    text_w = info.text_width_twips
    cx_emu = d.get("cx") or 0
    if not cx_emu:
        return None
    cx_twips = int(cx_emu * TWIPS_PER_INCH / EMU_PER_INCH)

    # 嵌入式
    if d.get("inline"):
        para = d.get("paragraph")
        if para is None:
            return None
        jc = para_jc(para)
        ppr = child(para, "w:pPr")
        ind_el = child(ppr, "w:ind")
        left_ind = int_attr(ind_el, "w:left", 0) + int_attr(ind_el, "w:start", 0)
        right_ind = int_attr(ind_el, "w:right", 0) + int_attr(ind_el, "w:end", 0)
        if jc == "center":
            # 段落居中 → 图片中心与"版心内可用宽度"的中心重合
            # 左距 = left_ind + (avail - cx)/2, 右距 = right_ind + (avail - cx)/2
            # |左距 - 右距| = |left_ind - right_ind|
            return abs(left_ind - right_ind)
        if jc == "right":
            left_dist = text_w - right_ind - cx_twips
            right_dist = right_ind
            return abs(left_dist - right_dist)
        # 默认左对齐
        left_dist = left_ind
        right_dist = text_w - left_ind - cx_twips
        return abs(left_dist - right_dist)

    # 浮动
    anchor_el = d.get("anchor_el")
    if anchor_el is None:
        return None
    positionH = anchor_el.find("wp:positionH", NS)
    if positionH is None:
        return None
    rel_from = positionH.attrib.get("relativeFrom", "")
    align_el = positionH.find("wp:align", NS)
    posOffset_el = positionH.find("wp:posOffset", NS)
    if align_el is not None and (align_el.text or "").strip() == "center" and rel_from in {"margin", "column"}:
        return 0
    if posOffset_el is not None:
        try:
            offset_emu = int((posOffset_el.text or "0").strip())
        except Exception:
            return None
        offset_twips = int(offset_emu * TWIPS_PER_INCH / EMU_PER_INCH)
        if rel_from in {"margin", "column"}:
            left_dist = offset_twips
            right_dist = text_w - offset_twips - cx_twips
            return abs(left_dist - right_dist)
        if rel_from == "page":
            left_dist = offset_twips - info.margin_left_twips
            right_dist = info.margin_right_twips - (info.page_width_twips - offset_twips - cx_twips)
            return abs(abs(left_dist) - abs(right_dist))
    return None


def page_count(info: WordInfo) -> int:
    if info.com.get("available") and info.com.get("pages"):
        return int(info.com["pages"])
    chars = sum(len(t) for t in info.texts)
    table_weight = len(info.tables) * 700
    figure_weight = len(drawing_info(info)) * 500
    explicit_breaks = sum(1 for p in info.paragraphs for br in p.findall(".//w:br", NS) if attr(br, "w:type") == "page")
    return max(1, math.ceil((chars + table_weight + figure_weight) / 1700) + explicit_breaks)


def dimension_one(info: WordInfo) -> Tuple[bool, List[str]]:
    issues = []
    if info.path.suffix.lower() != ".docx":
        issues.append("文件不是 .docx")
    elif not info.zip_ok:
        issues.append("docx 包无法正常打开")
    # 说明：按用户要求，删除以下维度一检查
    #   - "无连续 2 页以上空白页"
    #   - "无超过 1/3 页面面积乱码"
    #   - "无超过 1/3 页面面积文字重叠"
    #   - "正文/参考文献引用/表格/图片/图题/表题均可编辑或可正常选中，不能整篇导出为PDF或图片"
    return len(issues) == 0, issues


def add_hit(hits: List[Hit], score: int, rule: str, passed: bool, detail: str):
    # 说明：按接口统一约定，命中与未命中的规则都要记录，供 evaluate(dir_path)
    # 组装 dim2_items 时区分 hit=True/False；不改变各调用点原有的判断逻辑。
    hits.append(Hit(score, rule, passed, detail))


def section_text(info: WordInfo, start_keywords: List[str], end_keywords: Optional[List[str]] = None) -> str:
    texts = block_texts(info)
    start = None
    for i, text in enumerate(texts):
        stripped = text.replace(" ", "").replace("　", "")
        if any(k in text or k in stripped for k in start_keywords):
            # Skip TOC entries (keyword followed only by page number)
            clean = stripped.rstrip("0123456789")
            if clean != stripped and re.match(r"^[一-鿿\w]+\d+$", stripped):
                continue
            if start is None:
                start = i
    if start is None:
        return ""
    end = len(texts)
    if end_keywords:
        for i in range(start + 1, len(texts)):
            if any(k in texts[i] for k in end_keywords):
                end = i
                break
    return "\n".join(texts[start:end])


def _collect_ref_anchors(info: "WordInfo") -> set:
    """收集"参考文献条目"的书签锚点集合——与 evaluate_reference_citations 中的规则一致。"""
    doc = info.xml.get("word/document.xml")
    if doc is None:
        return set()
    ref_anchors: set = set()
    paragraphs = doc.findall(".//w:p", NS)
    in_ref = False
    for p in paragraphs:
        txt = all_text(p).strip()
        if not in_ref:
            if re.match(r"^参\s*考\s*文\s*献\s*$", txt) or txt == "参考文献":
                in_ref = True
            continue
        if txt.startswith("致谢") or txt.startswith("附录"):
            break
        for bk in p.findall(".//w:bookmarkStart", NS):
            name = bk.attrib.get(qn("w:name")) or ""
            if name and not name.startswith("_Toc") and not name.startswith("_GoBack"):
                ref_anchors.add(name)
    for bk in doc.findall(".//w:bookmarkStart", NS):
        name = bk.attrib.get(qn("w:name")) or ""
        if re.match(r"^[Rr]ef[_\-]?\d+$", name) or re.match(r"^ref\d+$", name, re.IGNORECASE):
            ref_anchors.add(name)
    return ref_anchors


def _paragraph_regularized_ref_nums(paragraph: ET.Element, ref_anchors: set) -> set:
    """返回该段落中"规范引用"覆盖到的参考文献编号集合。
    规范引用 = 引用位置([N])的数字部分位于 REF/NOTEREF/PAGEREF 域或指向
    参考文献书签的 HYPERLINK/w:hyperlink 内。"""
    result: set = set()
    ptxt = all_text(paragraph)
    matches = list(CITATION_RE.finditer(ptxt))
    if not matches:
        return result
    pos_map = _paragraph_run_positions(paragraph)
    spans: List[Tuple[int, int]] = []
    for f in paragraph.findall(".//w:fldSimple", NS):
        instr = (f.attrib.get(qn("w:instr")) or "").strip()
        anchor = _instr_bookmark(instr)
        kind = _instr_kind(instr)
        if kind in {"REF", "NOTEREF", "PAGEREF"} and anchor and anchor in ref_anchors:
            s, e = _element_char_span(f, pos_map)
            if s >= 0:
                spans.append((s, e))
    for r in _find_complex_field_ranges(paragraph, ref_anchors):
        spans.append(r["span"])
    for h in paragraph.findall(".//w:hyperlink", NS):
        anch = h.attrib.get(qn("w:anchor")) or ""
        if anch and anch in ref_anchors:
            s, e = _element_char_span(h, pos_map)
            if s >= 0:
                spans.append((s, e))
    for m in matches:
        g1s, g1e = m.start(1), m.end(1)
        if any(s <= g1s and g1e <= e for s, e in spans):
            # 解析 group(1) 内出现的所有数字（支持 [7,8]/[7-9]/[7、8]）
            for tok in re.split(r"[,，、\-]", m.group(1)):
                tok = tok.strip()
                if tok.isdigit():
                    result.add(int(tok))
    return result


def _collect_section_regularized_refs(info: "WordInfo") -> Dict[str, set]:
    """按 body_blocks 的段落位置把文档切成"章节区间"，对每个已知的小标题
    (纯文本恰为该标题、或以" 该标题"结尾且开头为编号)，聚合区间内所有规范
    引用的参考文献编号集合并返回。使用 body_blocks 索引避免正文中同名字符串
    误命中；同时忽略目录项(含"..."/"……"或以页码结尾)。"""
    ref_anchors = _collect_ref_anchors(info)
    known_titles = [
        "前言",
        "研究设计与资料来源", "指标体系与变量设定", "空间数据整理",
        "问卷与访谈资料", "统计分析方法",
        "居民步行行为分析", "健康感知与路径模型", "稳健性与分层分析", "讨论",
    ]

    # 找到每个小标题的段落索引
    heading_positions: List[Tuple[int, str]] = []
    for i, (kind, blk) in enumerate(info.body_blocks):
        if kind != "p":
            continue
        txt = all_text(blk).strip()
        if not txt:
            continue
        # 排除目录项
        if "..." in txt or "……" in txt or re.search(r"\.{3,}\s*\d+\s*$", txt):
            continue
        # 提取"最后的中文短语"部分：允许前缀为"1"/"1.1"/"1.2.3"/"1.2.3 "/"3  "等
        cleaned = re.sub(r"^\s*[0-9]+(?:\.[0-9]+)*\s+", "", txt)
        cleaned = cleaned.strip()
        if cleaned in known_titles:
            heading_positions.append((i, cleaned))
        elif txt.strip() in known_titles:
            heading_positions.append((i, txt.strip()))

    # 按段落索引升序切段：每段自身作为"章节起点"，直到下一个 heading 之前
    heading_positions.sort()
    result: Dict[str, set] = {t: set() for t in known_titles}
    if not heading_positions:
        return result
    for idx, (start_idx, title) in enumerate(heading_positions):
        end_idx = heading_positions[idx + 1][0] if idx + 1 < len(heading_positions) else len(info.body_blocks)
        # 收集该区间内(不含标题段自身)所有 w:p 里的规范引用
        for j in range(start_idx + 1, end_idx):
            kind, blk = info.body_blocks[j]
            if kind != "p":
                continue
            result[title] |= _paragraph_regularized_ref_nums(blk, ref_anchors)
    return result


def _max_regularized_ref_num(info: "WordInfo") -> int:
    """全文正文中"规范引用"覆盖到的最大参考文献编号，用于判定"及后续"边界。"""
    ref_anchors = _collect_ref_anchors(info)
    doc = info.xml.get("word/document.xml")
    if doc is None:
        return 0
    max_n = 0
    body = doc.find("w:body", NS)
    if body is None:
        return 0
    for p in body.findall(".//w:p", NS):
        for n in _paragraph_regularized_ref_nums(p, ref_anchors):
            if n > max_n:
                max_n = n
    return max_n


def _numbering_decimal_dot_num_ids(info: "WordInfo") -> set:
    """扫描 numbering.xml，返回所有"lvl=0 且 numFmt=decimal 且 lvlText 为 '%1.'"
    的 w:numId 集合。命中即办公软件会把段落编号渲染成 "N." 形式。"""
    result: set = set()
    numbering = info.xml.get("word/numbering.xml")
    if numbering is None:
        return result
    # abstractNumId → 是否为 decimal.dot (lvl0)
    abstract_ok: Dict[str, bool] = {}
    for an in numbering.findall("w:abstractNum", NS):
        aid = an.attrib.get(qn("w:abstractNumId"))
        if aid is None:
            continue
        # 检查 lvl0
        ok = False
        for lvl in an.findall("w:lvl", NS):
            if lvl.attrib.get(qn("w:ilvl")) != "0":
                continue
            fmt_el = lvl.find("w:numFmt", NS)
            txt_el = lvl.find("w:lvlText", NS)
            fmt_v = fmt_el.attrib.get(qn("w:val")) if fmt_el is not None else None
            txt_v = txt_el.attrib.get(qn("w:val")) if txt_el is not None else None
            if fmt_v == "decimal" and txt_v == "%1.":
                ok = True
                break
        abstract_ok[aid] = ok
    # numId → abstractNumId
    for num in numbering.findall("w:num", NS):
        nid = num.attrib.get(qn("w:numId"))
        aid_el = num.find("w:abstractNumId", NS)
        aid = aid_el.attrib.get(qn("w:val")) if aid_el is not None else None
        if nid and aid and abstract_ok.get(aid):
            result.add(nid)
    return result


def _numbering_decimal_bare_space_num_ids(info: "WordInfo") -> set:
    """扫描 numbering.xml，返回所有 lvl=0 满足如下条件的 w:numId 集合：
       - numFmt = decimal（阿拉伯数字）
       - lvlText = "%1"（序号后无任何标点符号）
       - w:suff  = "space"（序号后以一个空格分隔正文，办公软件即渲染为 "N ..."）
    命中即办公软件把段落自动渲染成"阿拉伯数字 + 空格 + 正文"，与本细则字面一致。"""
    result: set = set()
    numbering = info.xml.get("word/numbering.xml")
    if numbering is None:
        return result
    abstract_ok: Dict[str, bool] = {}
    for an in numbering.findall("w:abstractNum", NS):
        aid = an.attrib.get(qn("w:abstractNumId"))
        if aid is None:
            continue
        ok = False
        for lvl in an.findall("w:lvl", NS):
            if lvl.attrib.get(qn("w:ilvl")) != "0":
                continue
            fmt_el = lvl.find("w:numFmt", NS)
            txt_el = lvl.find("w:lvlText", NS)
            suff_el = lvl.find("w:suff", NS)
            fmt_v = fmt_el.attrib.get(qn("w:val")) if fmt_el is not None else None
            txt_v = txt_el.attrib.get(qn("w:val")) if txt_el is not None else None
            suff_v = suff_el.attrib.get(qn("w:val")) if suff_el is not None else None
            if fmt_v == "decimal" and txt_v == "%1" and suff_v == "space":
                ok = True
                break
        abstract_ok[aid] = ok
    for num in numbering.findall("w:num", NS):
        nid = num.attrib.get(qn("w:numId"))
        aid_el = num.find("w:abstractNumId", NS)
        aid = aid_el.attrib.get(qn("w:val")) if aid_el is not None else None
        if nid and aid and abstract_ok.get(aid):
            result.add(nid)
    return result


def _evaluate_reference_list_numbering(info: "WordInfo") -> Tuple[bool, Dict[str, object]]:
    """按细则判定参考文献列表编号——细则原文：
       "参考文献内容中每条文献编号阿拉伯数字编号，序号从1开始连续递增，
        序号后无标点符号，序号后空一格"

    拆点（针对 Word/WPS 办公软件真实渲染）：
       点1）编号为**阿拉伯数字**（0-9 十进制数字，非罗马/中文/带圈序号）；
       点2）序号从 **1** 开始；
       点3）序号**连续递增**（[1..N]，无跳号、无重号）；
       点4）序号后**无标点符号**（不得出现 "."、"、"、"．"、"："、")" 等）；
       点5）序号后**空一格**（紧跟 1 个 ASCII 空格，然后接正文）。
    办公软件下满足以上五点的渲染形态有两种等价途径：
       (a) 段落文本直接以 "N␣正文" 开头（手动键入的"阿拉伯数字+空格"）；
       (b) 段落 numPr 指向自动列表 numFmt=decimal、lvlText="%1"、w:suff="space"，
           Word/WPS 将该段落渲染为 "N␣正文"。
    细则未认可全角句点/顿号/半角句点等标点前缀，因此本函数一律拒绝。"""
    # 1) 定位参考文献章节的段落区间
    start_idx: Optional[int] = None
    end_idx = len(info.body_blocks)
    for i, (kind, blk) in enumerate(info.body_blocks):
        if kind != "p":
            continue
        txt = all_text(blk).strip()
        stripped = txt.replace("　", "").replace(" ", "")
        if re.match(r"^参考文献$", stripped):
            start_idx = i
            break
    if start_idx is None:
        return False, {"numbers": [], "fmt_source": "未找到参考文献章节"}
    for j in range(start_idx + 1, len(info.body_blocks)):
        kind, blk = info.body_blocks[j]
        if kind != "p":
            continue
        t = all_text(blk).strip()
        if t.startswith("致谢") or t.startswith("附录"):
            end_idx = j
            break

    # 2) 识别自动列表中命中 decimal + "%1" + suff=space 的 numId 集合
    bare_space_ids = _numbering_decimal_bare_space_num_ids(info)

    # 3) 遍历章节内每个非空段落，抽取编号
    numbers: List[int] = []
    fmt_source: List[str] = []
    for j in range(start_idx + 1, end_idx):
        kind, blk = info.body_blocks[j]
        if kind != "p":
            continue
        txt = all_text(blk)
        if not txt.strip():
            continue

        # (a) 手写 "N␣正文"——严格匹配细则的 4 个字面要求：
        #     开头是十进制数字(点1)、数字后紧跟恰好 1 个 ASCII 空格(点5)、
        #     该空格后紧跟任意非空白正文字符(排除"仅数字+空格"这种空条目)、
        #     且中间不含任何标点(点4)。
        m = re.match(r"^\s*(\d+) (\S)", txt)
        if m:
            # 点4 兜底：确保紧随空格之后的首字符不是标点符号本身
            #          （例如 "1 、张三" 这种"空格+顿号"仍视为带标点，判违规）
            first = m.group(2)
            if first in ".,;:、。．，；：)）]】":
                numbers.append(-1)
                fmt_source.append("bad_punct")
            else:
                numbers.append(int(m.group(1)))
                fmt_source.append("manual")
            continue

        # (b) 自动列表：段落 numPr 指向 decimal+"%1"+suff=space 的定义；
        #     办公软件 Word/WPS 会将该段落渲染为 "N␣正文" —— 与细则五点一致。
        numPr = blk.find(".//w:numPr", NS)
        numId_el = numPr.find(".//w:numId", NS) if numPr is not None else None
        numId_v = numId_el.attrib.get(qn("w:val")) if numId_el is not None else None
        if numId_v and numId_v in bare_space_ids:
            # 段落显示编号由 Word 自动递增；这里以自动列表出现次序 +1 作为其编号。
            numbers.append(len([s for s in fmt_source if s == "auto"]) + 1)
            fmt_source.append("auto")
            continue

        # (c) 段落以"数字"开头但不符合"数字+空格+正文"——细则不接受，判为不规范。
        #     覆盖：数字后带任何标点(1. / 1、 / 1) / 1．)、数字后为全角空格、
        #     数字与正文之间超过 1 个空格、数字后紧接非空白字符(1张三)等。
        m2 = re.match(r"^\s*(\d+)", txt)
        if m2:
            numbers.append(-1)
            fmt_source.append("bad_fmt")
            continue

        # 其它文本行（跨行续写等）忽略

    # 4) 判定（严格对齐细则五点）：
    #    - 至少存在一条条目；
    #    - 每条编号格式合规（无 -1，即无"带标点/无空格/格式错"）；
    #    - 编号序列 == [1, 2, ..., N]（首为 1、连续递增，且隐含无重号/跳号）。
    if not numbers:
        return False, {"numbers": [], "fmt_source": "章节内未找到条目"}
    if any(n == -1 for n in numbers):
        return False, {"numbers": numbers, "fmt_source": "存在带标点或格式错误的编号"}
    ok = numbers == list(range(1, len(numbers) + 1))
    return ok, {
        "numbers": numbers,
        "fmt_source": "/".join(fmt_source) if fmt_source else "",
    }


def evaluate_reference_citations(info: WordInfo) -> Tuple[bool, Dict[str, object]]:
    """按细则判定：正文中每一处 [N]/【N】 的引用位置，是否落在
    Word 交叉引用域 (REF/NOTEREF)、书签引用域 (PAGEREF)、
    或指向参考文献书签的可跳转对象 (HYPERLINK\\l / w:hyperlink @anchor) 内。

    办公软件 Word/WPS 中，只有以上三类结构点击时才能跳转到参考文献条目；
    直接键入的方括号数字不构成"引用对象"，不计入通过。"""
    doc = info.xml.get("word/document.xml")
    if doc is None:
        return False, {"citation_count": 0, "covered": 0, "ref_field_hits": 0,
                       "hyperlink_hits": 0, "uncovered_sample": []}

    # 1) 收集参考文献锚点集合：
    #    (a) 参考文献章节内的所有书签名(通常以 Ref_ 或数字开头)；
    #    (b) 目录锚点 _Toc* 一律排除。
    ref_anchors: set = set()
    ref_sec_ids: set = set()
    # 找出正文中"参考文献"标题所在段落之后的所有书签
    paragraphs = doc.findall(".//w:p", NS)
    in_ref_section = False
    for p in paragraphs:
        txt = all_text(p).strip()
        if not in_ref_section:
            if re.match(r"^参\s*考\s*文\s*献\s*$", txt) or txt == "参考文献":
                in_ref_section = True
            continue
        # 命中"致谢"/"附录"视为出参考文献段
        if txt.startswith("致谢") or txt.startswith("附录"):
            break
        for bk in p.findall(".//w:bookmarkStart", NS):
            name = bk.attrib.get(qn("w:name")) or ""
            if name and not name.startswith("_Toc") and not name.startswith("_GoBack"):
                ref_anchors.add(name)

    # 兜底：文档中所有以 Ref/ref 开头的书签也纳入
    for bk in doc.findall(".//w:bookmarkStart", NS):
        name = bk.attrib.get(qn("w:name")) or ""
        if re.match(r"^[Rr]ef[_\-]?\d+$", name) or re.match(r"^ref\d+$", name, re.IGNORECASE):
            ref_anchors.add(name)

    if not ref_anchors:
        return False, {"citation_count": 0, "covered": 0, "ref_field_hits": 0,
                       "hyperlink_hits": 0, "uncovered_sample": [],
                       "reason": "未找到参考文献书签"}

    # 2) 遍历正文段落，逐个 [N]/【N】 判定其运行区域是否被
    #    交叉引用域 / 书签引用域 / 指向 ref_anchors 的超链接 包裹。
    ref_field_hits = 0
    hyperlink_hits = 0
    citation_count = 0
    covered = 0
    uncovered_sample: List[str] = []

    body = doc.find("w:body", NS)
    body_paragraphs = body.findall(".//w:p", NS) if body is not None else []

    for p in body_paragraphs:
        # 段落文本用于计数引用出现次数
        ptxt = all_text(p)
        matches = list(CITATION_RE.finditer(ptxt))
        if not matches:
            continue

        # (a) 段内 fldSimple: 若 instr 命中 REF/NOTEREF/PAGEREF 且锚点在 ref_anchors，
        #     视为一个"合规引用位置"
        fldsimple_ok: List[Tuple[int, int]] = []  # 记录被域包裹的字符区间
        pos_map = _paragraph_run_positions(p)  # {run_element: (start_char, end_char)}

        for f in p.findall(".//w:fldSimple", NS):
            instr = (f.attrib.get(qn("w:instr")) or "").strip()
            anchor = _instr_bookmark(instr)
            kind = _instr_kind(instr)
            if kind in {"REF", "NOTEREF", "PAGEREF"} and anchor and anchor in ref_anchors:
                # fldSimple 的内容 run 都在其子树里
                start, end = _element_char_span(f, pos_map)
                if start >= 0:
                    fldsimple_ok.append((start, end))
                    ref_field_hits += 1

        # (b) 复杂域（fldChar begin/separate/end）：使用 run 序列扫描
        complex_ranges = _find_complex_field_ranges(p, ref_anchors)
        for r in complex_ranges:
            fldsimple_ok.append(r["span"])
            if r["kind"] == "HYPERLINK":
                hyperlink_hits += 1
            else:
                ref_field_hits += 1

        # (c) w:hyperlink @w:anchor 指向 ref_anchors
        for h in p.findall(".//w:hyperlink", NS):
            anch = h.attrib.get(qn("w:anchor")) or ""
            if anch and anch in ref_anchors:
                start, end = _element_char_span(h, pos_map)
                if start >= 0:
                    fldsimple_ok.append((start, end))
                    hyperlink_hits += 1

        # 判定每处引用 match 是否落在任一合规区间内。
        # Word 的标准做法：方括号是用户手动键入的字符，而 REF/HYPERLINK 域返回的是
        # 参考文献书签的编号；因此域实际只包裹方括号内的数字部分。此处按"数字部分
        # 位于域范围内"作为该引用位置具备交叉引用/书签引用/可跳转对象的证据。
        for m in matches:
            citation_count += 1
            # 数字子串位置 = 方括号内的第一个 group 的 span
            g1_start, g1_end = m.start(1), m.end(1)
            hit = any(s <= g1_start and g1_end <= e for s, e in fldsimple_ok)
            if hit:
                covered += 1
            else:
                if len(uncovered_sample) < 10:
                    ctx = ptxt[max(0, m.start() - 8):m.end() + 8]
                    uncovered_sample.append(ctx)

    # 细则要求"所有引用位置"，因此必须 covered == citation_count 且 citation_count > 0
    ok = citation_count > 0 and covered == citation_count

    return ok, {
        "citation_count": citation_count,
        "covered": covered,
        "ref_field_hits": ref_field_hits,
        "hyperlink_hits": hyperlink_hits,
        "uncovered_sample": uncovered_sample,
        "ref_anchor_count": len(ref_anchors),
    }


def _instr_kind(instr: str) -> str:
    u = instr.upper().strip()
    if u.startswith("REF ") or u == "REF" or " REF " in " " + u + " ":
        # 但要排除 PAGEREF/NOTEREF 已单独处理
        if u.startswith("PAGEREF") or u.startswith("NOTEREF"):
            pass
        else:
            return "REF"
    if u.startswith("NOTEREF"):
        return "NOTEREF"
    if u.startswith("PAGEREF"):
        return "PAGEREF"
    if u.startswith("HYPERLINK"):
        return "HYPERLINK"
    return ""


def _instr_bookmark(instr: str) -> Optional[str]:
    """从域代码字符串中解析出锚点/书签名。"""
    s = instr.strip()
    u = s.upper()
    if u.startswith("REF ") or u.startswith("NOTEREF ") or u.startswith("PAGEREF "):
        # 形如: REF Ref_1 \h  → 第 2 个 token
        parts = s.split()
        if len(parts) >= 2:
            return parts[1]
    if u.startswith("HYPERLINK"):
        # 形如: HYPERLINK \l "_Ref_1" 或 HYPERLINK \l Ref_1
        m = re.search(r'\\l\s+"?([^"\s]+)"?', s)
        if m:
            return m.group(1)
    return None


def _paragraph_run_positions(paragraph: ET.Element) -> Dict[ET.Element, Tuple[int, int]]:
    """按段落文本流位置为每个 w:r/w:hyperlink/w:fldSimple 建立字符区间索引。
    与 all_text() 一致的字符累积顺序。"""
    positions: Dict[ET.Element, Tuple[int, int]] = {}
    pos = [0]

    def visit(el: ET.Element, path: List[ET.Element]):
        start = pos[0]
        for c in list(el):
            tag = c.tag.split("}", 1)[1] if "}" in c.tag else c.tag
            if tag == "t":
                pos[0] += len(c.text or "")
            elif tag == "tab":
                pos[0] += 1
            elif tag == "br":
                pass
            else:
                visit(c, path + [c])
        end = pos[0]
        tag = el.tag.split("}", 1)[1] if "}" in el.tag else el.tag
        if tag in {"r", "hyperlink", "fldSimple"}:
            positions[el] = (start, end)

    visit(paragraph, [])
    return positions


def _element_char_span(el: ET.Element, pos_map: Dict[ET.Element, Tuple[int, int]]) -> Tuple[int, int]:
    if el in pos_map:
        return pos_map[el]
    # 兜底：用其子孙 run 的极值
    starts = []
    ends = []
    for r in el.findall(".//w:r", NS):
        if r in pos_map:
            s, e = pos_map[r]
            starts.append(s)
            ends.append(e)
    if starts:
        return min(starts), max(ends)
    return -1, -1


def _find_complex_field_ranges(paragraph: ET.Element, ref_anchors: set) -> List[Dict[str, object]]:
    """扫描段落中的复杂域(fldChar begin/separate/end)，返回落在
    ref_anchors 范围内的每个域区间(字符起止 + 类型)。"""
    pos_map = _paragraph_run_positions(paragraph)
    result: List[Dict[str, object]] = []
    # 收集段内所有 w:r 及其 fldChar/instrText
    runs = paragraph.findall(".//w:r", NS)
    # 简单状态机
    active_stack: List[Dict[str, object]] = []
    for r in runs:
        for fc in r.findall("w:fldChar", NS):
            ft = fc.attrib.get(qn("w:fldCharType"))
            if ft == "begin":
                active_stack.append({"instr": "", "start_run": r})
            elif ft == "end" and active_stack:
                cur = active_stack.pop()
                instr = cur.get("instr", "")
                anchor = _instr_bookmark(instr)
                kind = _instr_kind(instr)
                if kind in {"REF", "NOTEREF", "PAGEREF", "HYPERLINK"} and anchor and anchor in ref_anchors:
                    s0, _ = _element_char_span(cur["start_run"], pos_map)
                    _, e1 = _element_char_span(r, pos_map)
                    if s0 >= 0 and e1 >= 0:
                        result.append({"kind": kind, "span": (s0, e1)})
        for it in r.findall("w:instrText", NS):
            if active_stack:
                active_stack[-1]["instr"] = active_stack[-1].get("instr", "") + (it.text or "")
    return result


def _evaluate_hits(info: WordInfo) -> List[Hit]:
    """维度二评分内核：逐条细则判断，返回命中与未命中的 Hit 列表（不对外暴露）。
    对外统一入口见 evaluate(dir_path) -> dict。"""
    hits: List[Hit] = []
    all_text_joined = "\n".join(info.texts)
    table_captions = find_caption_indices(info, "表")
    figure_captions = find_caption_indices(info, "图")
    drawings = drawing_info(info)
    texts = block_texts(info)

    # ============ 细则 +5 表格保留/居中/宽度/未越界 ============
    # 逐点对齐细则，任一点不满足则不得分；仅约束细则明确要求的点。
    tolerance_twips = int(round(0.5 * CM_TO_TWIPS))   # 0.5 cm
    width_diff_max = CM_TO_TWIPS                       # 1 cm
    text_w = info.text_width_twips

    numbered_tables = find_numbered_tables(info, table_captions)

    # 点1：全文保留表1至表10共10个表格
    ten_tables_present = all(n in numbered_tables for n in range(1, 11))
    # 点2：表题依次包含 TABLE_TITLES 中的10个标题
    titles_in_order_ok = caption_titles_in_order(info, table_captions, TABLE_TITLES)

    ordered_tables = [numbered_tables[n] for n in range(1, 11)] if ten_tables_present else []

    # 点3：所有表格设置为居中对齐（w:jc=center；空值不视为居中）
    all_center_ok = bool(ordered_tables) and all(table_align(t) == "center" for t in ordered_tables)

    # 有效宽度（twips），用于点5/6/7/8
    eff_widths = [table_effective_width_twips(t, text_w) for t in ordered_tables]
    widths_known = ordered_tables and all(w is not None and w > 0 for w in eff_widths)

    # 点4：表格左右边界相对正文版心居中，偏差≤0.5cm
    #      左缩进 tblInd 视为距正文左边距的偏移，右侧剩余=text_w-ind-width
    #      要求 |ind - (text_w-ind-width)| = |2*ind + width - text_w| ≤ 0.5cm
    center_geom_ok = True
    if widths_known:
        for t, w in zip(ordered_tables, eff_widths):
            ind = table_indent_twips(t)
            if abs(2 * ind + w - text_w) > tolerance_twips:
                center_geom_ok = False
                break
    else:
        center_geom_ok = False

    # 点5：宽度统一为版心宽度的 85%–100%
    width_range_ok = widths_known and all(0.85 * text_w <= w <= text_w for w in eff_widths)

    # 点6：任意两个表格宽度差 ≤ 1cm
    width_uniform_ok = widths_known and (max(eff_widths) - min(eff_widths) <= width_diff_max)

    # 点7：全文所有表格宽度未超出页面可打印区域（=正文版心宽度）
    all_tables_eff = [table_effective_width_twips(t, text_w) for t in info.tables]
    printable_ok = all((w is None) or (w <= text_w) for w in all_tables_eff)

    # 点8：表格右边界没有越过正文右边距，即 ind + width ≤ text_w
    right_within_ok = True
    for t in info.tables:
        w = table_effective_width_twips(t, text_w)
        if w is None:
            continue
        ind = table_indent_twips(t)
        if ind + w > text_w + 2:  # 2 twips 容差，避免舍入抖动
            right_within_ok = False
            break

    plus5_passed = (
        ten_tables_present and titles_in_order_ok and all_center_ok
        and center_geom_ok and width_range_ok and width_uniform_ok
        and printable_ok and right_within_ok
    )
    add_hit(
        hits, 5,
        "全文保留表1-表10、题目依序、居中、宽度85%-100%版心、两两差≤1cm、未越界",
        plus5_passed,
        (
            f"表1-10齐备={ten_tables_present}; 题序符合={titles_in_order_ok}; "
            f"居中={all_center_ok}; 几何居中(≤0.5cm)={center_geom_ok}; "
            f"85-100%={width_range_ok}; 宽差≤1cm={width_uniform_ok}; "
            f"未超版心={printable_ok}; 右边界未越={right_within_ok}; "
            f"版心twips={text_w}; 有效宽度={eff_widths}"
        ),
    )

    # ============ 细则 +5 三线表样式与线宽 ============
    # 细则要求：全文所有表格均为三线表样式，仅保留三条主横线——表格上边线、
    # 表头下方横线、表格下边线；上/下边线 1.5 磅，表头下方线 0.75 磅。
    # 说明：封面等布局用表已被 is_cover_table 单独识别；此处按细则"全文所有表格"
    # 对全部 w:tbl 元素逐一校验（含封面表——原实现口径一致）。
    three_line_details = []
    three_line_count = 0
    for i, t in enumerate(info.tables):
        ok = is_three_line_table(t)
        if ok:
            three_line_count += 1
        else:
            three_line_details.append(i + 1)
    add_hit(
        hits, 5,
        "全文所有表格均为三线表(仅上/表头下/下三条主横线；1.5/0.75/1.5磅)",
        len(info.tables) > 0 and three_line_count == len(info.tables),
        f"三线表 {three_line_count}/{len(info.tables)}；不合格表序号 {three_line_details[:10]}"
    )
    # ============ 细则 +3 表头字体：加粗；中文宋体五号；英/数 Times New Roman 五号 ============
    # 说明：细则口径为"全文所有表格"。封面/学位表（is_cover_table）是版式化的
    # 学位信息表，其首行不构成研究表格的"表头"；这里按含义排除，仅对内容表校验。
    content_tables = [t for t in info.tables if not is_cover_table(t)]
    header_font_pass = 0
    header_font_fail = []
    for i, t in enumerate(content_tables):
        ok, reason = is_table_header_row_font_ok(t)
        if ok:
            header_font_pass += 1
        else:
            header_font_fail.append((i + 1, reason))
    add_hit(
        hits, 3,
        "全文所有表格表头文字加粗，中文宋体五号，英/数Times New Roman五号",
        len(content_tables) > 0 and header_font_pass == len(content_tables),
        f"表头合格 {header_font_pass}/{len(content_tables)}；不合格示例 {header_font_fail[:3]}"
    )
    # ============ 细则 +3 表格正文字体：中文宋体五号；英/数 Times New Roman 五号 ============
    # 说明：细则口径为"全文所有表格"。封面/学位表（is_cover_table）为版式表，
    # 无常规"表格正文"含义，按此排除；对内容表逐个校验除首行外的所有数据行。
    body_font_pass = 0
    body_font_fail = []
    for i, t in enumerate(content_tables):
        ok, reason = is_table_body_font_ok(t)
        if ok:
            body_font_pass += 1
        else:
            body_font_fail.append((i + 1, reason))
    add_hit(
        hits, 3,
        "全文所有表格正文中文宋体五号，英/数Times New Roman五号",
        len(content_tables) > 0 and body_font_pass == len(content_tables),
        f"正文合格 {body_font_pass}/{len(content_tables)}；不合格示例 {body_font_fail[:3]}"
    )
    # ============ 细则 +3 全文所有表格行高≥0.6cm 且文字未被上下边框裁切 ============
    # 说明：细则口径为"全文所有表格"，包含封面/学位版式表；两条独立子项均需通过。
    rh_pass = 0
    rh_fail = []
    clip_pass = 0
    clip_fail = []
    for i, t in enumerate(info.tables):
        ok_rh, why_rh = table_rows_heights_meet_min_cm(t, info, i)
        if ok_rh:
            rh_pass += 1
        else:
            rh_fail.append((i + 1, why_rh))
        ok_clip, why_clip = table_no_text_clipping(t, info, i)
        if ok_clip:
            clip_pass += 1
        else:
            clip_fail.append((i + 1, why_clip))
    add_hit(
        hits, 3,
        "全文所有表格行高≥0.6cm 且表格文字未被上下边框裁切",
        len(info.tables) > 0 and rh_pass == len(info.tables) and clip_pass == len(info.tables),
        f"行高合格 {rh_pass}/{len(info.tables)} 示例{rh_fail[:2]}；未裁切 {clip_pass}/{len(info.tables)} 示例{clip_fail[:2]}"
    )

    # ============ 细则 +3 表题格式：位于表上方；表序阿拉伯数字连续；
    #            "表 数字  表名"，表号与表名间 2 个空格；段前段后各 0.5 行；
    #            居中对齐；行距 12 磅（固定值） ============
    caption_details = []
    caption_all_ok = True
    # 收集正文中 tbl 元素对应的段落索引，用于校验表题在其上方
    tbl_positions: List[Tuple[int, ET.Element]] = [
        (i, b) for i, b in enumerate(info.body_blocks) if b[0] == "tbl"
    ]
    # 表序需为阿拉伯数字连续编号（1..N）
    expected_numbers = list(range(1, len(info.tables) + 1))
    caption_numbers_seq = sorted(k for k in table_captions.keys() if 1 <= k <= len(info.tables))
    numbers_sequential = caption_numbers_seq == expected_numbers
    if not numbers_sequential:
        caption_all_ok = False
        caption_details.append(f"编号非阿拉伯数字连续: {caption_numbers_seq} vs {expected_numbers}")

    for number in expected_numbers:
        cap_idx = table_captions.get(number)
        # 目标表格在 body_blocks 中的位置：正文里的第 number 张 w:tbl
        tbl_idx = tbl_positions[number - 1][0] if number - 1 < len(tbl_positions) else None
        if cap_idx is None:
            caption_all_ok = False
            caption_details.append(f"表{number} 未找到题注")
            continue
        if tbl_idx is None or cap_idx >= tbl_idx:
            caption_all_ok = False
            caption_details.append(f"表{number} 题注未位于表上方 (cap={cap_idx}, tbl={tbl_idx})")
            continue
        # 题注与表之间不能夹入正文内容段落（允许空段）
        between_has_content = False
        for i in range(cap_idx + 1, tbl_idx):
            k, blk = info.body_blocks[i]
            if k == "p" and all_text(blk).strip():
                between_has_content = True
                break
        if between_has_content:
            caption_all_ok = False
            caption_details.append(f"表{number} 题注与表格之间夹入正文段落")
            continue

        cap_block = info.body_blocks[cap_idx][1]
        cap_text = all_text(cap_block)

        # 表号"表 <N>" + 2 个空格 + 表名（表名非空）
        # Word 中文文档常见半角空格 " " 与全角空格 "　" 混用；按细则严格要求 2 个空格
        pat_strict = re.compile(rf"^表\s{number}[ 　]{{2}}\S.*$")
        pat_loose = re.compile(rf"^表\s?{number}[ 　]{{2}}\S.*$")
        # 严格接受"表 N␣␣<name>"，宽松再放行"表N␣␣<name>"（"表 数字 表名"允许 表-N 之间可选空格）
        if not (pat_strict.match(cap_text) or pat_loose.match(cap_text)):
            caption_all_ok = False
            caption_details.append(f"表{number} 题注格式不符: {cap_text!r}")
            continue

        # 居中对齐（w:jc=center；空值不视为居中，办公软件实际渲染以显式为准）
        if para_jc(cap_block) != "center":
            caption_all_ok = False
            caption_details.append(f"表{number} 未居中: jc={para_jc(cap_block)!r}")
            continue

        sp = para_spacing(cap_block)
        # 段前段后 0.5 行：w:beforeLines / w:afterLines 单位为 1/100 行，值应为 "50"
        if sp.get("beforeLines") != "50" or sp.get("afterLines") != "50":
            caption_all_ok = False
            caption_details.append(
                f"表{number} 段前后非0.5行: beforeLines={sp.get('beforeLines')!r}, afterLines={sp.get('afterLines')!r}"
            )
            continue

        # 行距 12 磅（固定值）：Word 中"固定值 12 磅" -> w:line="240" w:lineRule="exact"
        # (1 磅 = 20 twips → 12 磅 = 240 twips)
        line_val = sp.get("line", "")
        line_rule = sp.get("lineRule", "")
        if line_val != "240" or line_rule != "exact":
            caption_all_ok = False
            caption_details.append(
                f"表{number} 行距非12磅固定值: line={line_val!r}, lineRule={line_rule!r}"
            )
            continue

    add_hit(
        hits, 3,
        "表题位于表上方；表序阿拉伯数字连续；\"表 N  名\"；段前后0.5行；居中；行距12磅固定值",
        caption_all_ok and bool(table_captions),
        f"表题索引 {sorted(table_captions.keys())}；问题 {caption_details[:3]}"
    )

    # ============ 细则 +3 英文表题：位于中文表题下方；Times New Roman 五号；
    #            且与对应中文表题和表格相邻（三者顺序：中文表题 → 英文表题 → 表格） ============
    # 说明：细则口径是"全文所有英文表题"——只对存在中文题注的表格作双语校验；
    # 封面/缩略语/补充表等无中文题注的表格不属于本项范围。
    en_cap_ok = True
    en_cap_details: List[str] = []
    numbered_caps = sorted(table_captions.keys())
    for number in numbered_caps:
        cn_idx = table_captions.get(number)
        if cn_idx is None:
            en_cap_ok = False
            en_cap_details.append(f"表{number} 未找到中文表题")
            continue
        # 目标表格 = 中文表题往下最近的一张 tbl；跳过夹入的空段和候选英文题注
        tbl_idx: Optional[int] = None
        for j in range(cn_idx + 1, len(info.body_blocks)):
            if info.body_blocks[j][0] == "tbl":
                tbl_idx = j
                break
        if tbl_idx is None:
            en_cap_ok = False
            en_cap_details.append(f"表{number} 未找到对应表格")
            continue

        # 1) 英文表题必须紧跟在中文表题的下一个非空段落（"位于中文表题下方"）
        en_idx: Optional[int] = None
        for j in range(cn_idx + 1, tbl_idx):
            k, blk = info.body_blocks[j]
            if k != "p":
                break
            if all_text(blk).strip():
                en_idx = j
                break
        if en_idx is None:
            en_cap_ok = False
            en_cap_details.append(f"表{number} 中文表题下方缺少英文表题")
            continue

        en_block = info.body_blocks[en_idx][1]
        en_text = all_text(en_block)
        # 英文表题必须以 Table N 起首（Word/WPS 双语题注标准写法）
        if not re.match(rf"^\s*Table\s*{number}\b", en_text, re.IGNORECASE):
            en_cap_ok = False
            en_cap_details.append(f"表{number} 英文表题格式不符: {en_text[:40]!r}")
            continue

        # 2) 英文表题必须与"对应表格"相邻——en_idx 之后紧接的下一个非空块必须是该表格
        next_non_empty: Optional[int] = None
        for j in range(en_idx + 1, len(info.body_blocks)):
            k, blk = info.body_blocks[j]
            if k == "p" and not all_text(blk).strip():
                continue
            next_non_empty = j
            break
        if next_non_empty != tbl_idx:
            en_cap_ok = False
            en_cap_details.append(f"表{number} 英文表题与表格不相邻 (en={en_idx}, next={next_non_empty}, tbl={tbl_idx})")
            continue

        # 3) 英文表题字体：Times New Roman 五号（sz=21）——办公软件按 ascii/hAnsi 渲染英文
        bad = None
        checked_any = False
        for run in paragraph_runs(en_block):
            rtext = all_text(run)
            if not rtext.strip():
                continue
            if not re.search(r"[A-Za-z0-9]", rtext):
                continue
            checked_any = True
            asc = _run_ascii_font(run)
            size = run_size_half_points(run)
            if asc != "Times New Roman":
                bad = f"英/数非Times New Roman: {rtext!r} ascii={asc!r}"
                break
            if size != 21:
                bad = f"英/数非五号: {rtext!r} sz={size}"
                break
        if not checked_any:
            en_cap_ok = False
            en_cap_details.append(f"表{number} 英文表题未包含英文字符")
            continue
        if bad is not None:
            en_cap_ok = False
            en_cap_details.append(f"表{number} {bad}")
            continue

    add_hit(
        hits, 3,
        "英文表题位于中文表题下方；Times New Roman五号；与对应中文表题和表格相邻",
        en_cap_ok and bool(table_captions),
        f"问题 {en_cap_details[:3]}"
    )

    # ============ 细则 +1 表名上方与表注下方各空一行；若位于页首，上面不空行 ============
    # 说明：
    #   - "表名"=中文表题段落；"表注"=表格下方紧邻的注释段落(以"注:"/"注："/"Note"起首)。
    #     若不存在表注，则以表格自身作为"表注区末端"。
    #   - "空一行"= 版面上一个空段落(全空白 w:p)，办公软件即按段落序渲染。
    #   - "位于页首"= 该表题段落的顶端落在页面上边距处；Word COM 给出的
    #     Range.Information(10) 是该段起点距页面顶端的距离(点)。当值 ≤ 上边距+2pt
    #     视为页首。COM 不可用时，回退按 body_blocks 里 cap_idx < 3 的启发式豁免。
    def _is_empty_para(idx: int) -> bool:
        if idx < 0 or idx >= len(info.body_blocks):
            return False
        k, blk = info.body_blocks[idx]
        return k == "p" and not all_text(blk).strip()

    cap_page_info = info.com.get("caption_page_info") if isinstance(info.com, dict) else None
    top_margin_pt = float(info.com.get("page_top_margin_pt") or 0.0) if isinstance(info.com, dict) else 0.0
    page_top_by_num: Dict[int, bool] = {}
    if isinstance(cap_page_info, list) and cap_page_info:
        # 按编号收集"是否位于页首"标记；同编号出现多次时任一为页首即视为豁免
        for rec in cap_page_info:
            try:
                num = int(rec["num"])
                v_pos = float(rec["v_pos_pt"])
            except Exception:
                continue
            is_top = top_margin_pt > 0 and v_pos <= top_margin_pt + 2.0
            page_top_by_num[num] = page_top_by_num.get(num, False) or is_top

    blank_ok = True
    blank_details: List[str] = []
    for number, cap_idx in table_captions.items():
        # 只对目标 10 张表校验(细则针对"表 1..10"研究表格)
        if not (1 <= number <= 10):
            continue

        # 1) 表名上方：上一块必须是空段落；除非位于页首
        at_page_top = page_top_by_num.get(number, False) if page_top_by_num else (cap_idx < 3)
        if not at_page_top:
            if not _is_empty_para(cap_idx - 1):
                blank_ok = False
                prev_text = all_text(info.body_blocks[cap_idx - 1][1]) if cap_idx - 1 >= 0 else ""
                blank_details.append(f"表{number} 表名上方非空段: {prev_text[:20]!r}")
                continue

        # 2) 找到该编号对应的表格
        tbl_idx: Optional[int] = None
        for j in range(cap_idx + 1, len(info.body_blocks)):
            if info.body_blocks[j][0] == "tbl":
                tbl_idx = j
                break
        if tbl_idx is None:
            blank_ok = False
            blank_details.append(f"表{number} 未找到对应表格")
            continue

        # 3) 表注：表格之后紧接的、以"注:"/"注："/"Note"起首的段落即视为表注；否则末端=表格自身
        note_end_idx = tbl_idx
        if tbl_idx + 1 < len(info.body_blocks):
            nk, nblk = info.body_blocks[tbl_idx + 1]
            if nk == "p":
                ntext = all_text(nblk).lstrip()
                if ntext.startswith("注:") or ntext.startswith("注：") or re.match(r"^Note\s*[:：]", ntext, re.IGNORECASE):
                    note_end_idx = tbl_idx + 1

        # 4) 表注下方：紧接的下一块必须是空段落
        if not _is_empty_para(note_end_idx + 1):
            blank_ok = False
            after_text = all_text(info.body_blocks[note_end_idx + 1][1]) if note_end_idx + 1 < len(info.body_blocks) else ""
            blank_details.append(f"表{number} 表注下方非空段: {after_text[:20]!r}")
            continue

    add_hit(
        hits, 1,
        "表名上方与表注下方各空一行；若位于页首则上面不空行",
        blank_ok and bool(table_captions),
        f"页首豁免依据 COM 垂直位置；问题 {blank_details[:3]}"
    )

    # ============ 细则 +5 正文中所有引用参考文献的位置均设置为
    #            Word 交叉引用域、书签引用域，或可点击跳转到参考文献条目的引用对象 ============
    # 判定口径（针对办公软件 Word/WPS 的真实点击行为）：
    #   - 交叉引用域：`REF <bookmark> ...` 或 `NOTEREF <bookmark> ...`；
    #   - 书签引用域：`PAGEREF <bookmark> ...`（书签引用的一种表现形态）；
    #   - 可跳转对象：`HYPERLINK \l <bookmark>` 内部锚点跳转，或 `w:hyperlink @w:anchor`；
    #   - 目标锚点必须是"参考文献条目对应的书签"，排除目录锚点（_Toc*）。
    # 逐个引用位置校验：正文中出现 [N] / 【N】 的每一处，其运行(run)所在的
    #   段落必须被上述任一类域/超链接包裹，且目标锚点在参考文献书签集合内。
    ref_ok, ref_detail = evaluate_reference_citations(info)
    # 保留以下变量供后续依赖此项的评分点/扣分项使用
    citation_count = ref_detail["citation_count"]
    ref_field_count = ref_detail["ref_field_hits"]
    hyperlink_count = ref_detail["hyperlink_hits"]
    add_hit(
        hits, 5,
        "正文所有参考文献引用位置为Word交叉引用域/书签引用域/可跳转引用对象",
        ref_ok,
        f"引用总数 {citation_count}；已包裹 {ref_detail['covered']}；域命中 {ref_field_count}；超链接命中 {hyperlink_count}；未包裹样例 {ref_detail['uncovered_sample'][:3]}"
    )

    # ============ 细则 +5 正文中引用编号：方括号格式（[N]/[N,N]）；
    #            均为上标；字体 Times New Roman；字号四号（sz=28） ============
    # 判定口径（针对办公软件 Word/WPS 的真实渲染，从严按 rubric 示例执行）：
    #   - 方括号：只承认半角 "[" / "]"（"例如[1]、[2]、[7,8]"，细则不允许【】等中文括号）；
    #   - 内部字符：仅允许数字与"半角逗号(,)"作为编号分隔——
    #     rubric 示例仅给出半角逗号形态([7,8])，不出现中文逗号"，"、顿号"、"、连字符"-"，
    #     故从严按示例判定，其他分隔符视为不合规；
    #   - 上标：run 的 w:rPr/w:vertAlign@w:val="superscript"；
    #   - 字体：run 的 w:rFonts@w:ascii 或 @w:hAnsi 为 "Times New Roman"
    #     （未显式时按段落级 rPr 继承——办公软件实际就是按此渲染）；
    #   - 字号：w:sz@w:val = 28（半磅），即四号=14pt；
    #     Word/WPS 应用上标时，视觉上会自动缩小到 ~65%，但 XML 中 sz 值仍
    #     以"原始字号"记录，等于四号即 28（细则要求"字号为四号"就是这个字段值）。
    def _run_size_effective(run: ET.Element, para: ET.Element) -> Optional[int]:
        s = run_size_half_points(run)
        if s is not None:
            return s
        # 继承段落 rPr（Word/WPS 中段落 pPr/rPr 会作用到未显式的 run）
        ppr = child(para, "w:pPr")
        pr_of_p = child(ppr, "w:rPr") if ppr is not None else None
        sz_p = child(pr_of_p, "w:sz") if pr_of_p is not None else None
        val = attr(sz_p, "w:val")
        try:
            return int(val) if val is not None else None
        except Exception:
            return None

    def _run_font_effective(run: ET.Element, para: ET.Element) -> str:
        f = _run_ascii_font(run)
        if f:
            return f
        ppr = child(para, "w:pPr")
        pr_of_p = child(ppr, "w:rPr") if ppr is not None else None
        fonts = child(pr_of_p, "w:rFonts") if pr_of_p is not None else None
        for key in ("ascii", "hAnsi"):
            name = (attr(fonts, "w:" + key, "") or "").strip()
            if name:
                return name
        return ""

    cite_total = 0        # 引用位置数
    cite_ok = 0           # 完整满足四条子项的位置数
    cite_bad_samples: List[str] = []

    for p in info.paragraphs:
        ptxt = all_text(p)
        matches = list(CITATION_RE.finditer(ptxt))
        if not matches:
            continue
        pos_map = _paragraph_run_positions(p)  # {element: (start_char, end_char)}
        # 建立"字符位置 → 覆盖它的 run"映射
        char_to_run: Dict[int, ET.Element] = {}
        for r in p.findall(".//w:r", NS):
            span = pos_map.get(r)
            if span is None:
                continue
            s, e = span
            for i in range(s, e):
                char_to_run[i] = r

        for m in matches:
            cite_total += 1
            full = m.group(0)          # 原始子串，可能含 [ / 【
            inner = m.group(1)         # 数字与分隔符
            # 1) 方括号格式：只承认半角 [ ]
            if not (full.startswith("[") and full.endswith("]")):
                cite_bad_samples.append(f"非半角方括号: {full!r}")
                continue
            # 2) 内容合规：从严按 rubric 示例([1]/[2]/[7,8])——
            #    仅允许"数字"与"半角逗号,"作为编号内部分隔，
            #    中文逗号"，"、顿号"、"、连字符"-"均视为不合规。
            if not re.match(r"^\d+(?:,\d+)*$", inner):
                cite_bad_samples.append(f"引用内部字符不合规(仅允许数字与半角逗号): {full!r}")
                continue

            # 采样"整个 [N...] 子串"覆盖到的每一个 run
            runs_in: List[ET.Element] = []
            seen = set()
            for pos in range(m.start(), m.end()):
                r = char_to_run.get(pos)
                if r is None:
                    continue
                if id(r) in seen:
                    continue
                seen.add(id(r))
                runs_in.append(r)
            if not runs_in:
                cite_bad_samples.append(f"引用未定位到 run: {full!r}")
                continue

            # 3) 上标格式：整个引用位置(包括方括号与数字)所属的所有 run 均为 superscript
            all_sup = all(run_is_superscript(r) for r in runs_in)
            if not all_sup:
                cite_bad_samples.append(f"引用非上标: {full!r}")
                continue

            # 4) 字体 Times New Roman & 字号四号(sz=28)
            #    仅对含 ASCII 字符([]/数字/分隔符)的 run 校验(细则针对英/数字符)
            bad = None
            for r in runs_in:
                rtext = all_text(r)
                if not rtext.strip():
                    continue
                if not re.search(r"[A-Za-z0-9\[\],\-]", rtext):
                    continue
                fname = _run_font_effective(r, p)
                if fname != "Times New Roman":
                    bad = f"字体非TNR: run={rtext!r} font={fname!r}"
                    break
                sz = _run_size_effective(r, p)
                if sz != 28:
                    bad = f"字号非四号(sz=28): run={rtext!r} sz={sz}"
                    break
            if bad is not None:
                cite_bad_samples.append(f"{full!r} → {bad}")
                continue

            cite_ok += 1

    add_hit(
        hits, 5,
        "正文引用编号：方括号([1]/[7,8])；上标；Times New Roman；四号(sz=28)",
        cite_total > 0 and cite_ok == cite_total,
        f"引用总数 {cite_total}；合格 {cite_ok}；不合格样例 {cite_bad_samples[:3]}"
    )

    # ============ 细则 +5 各指定章节至少保留并"规范"对应参考文献交叉引用 ============
    # 判定口径（与办公软件真实点击行为一致）：
    #   1) 章节归属按标题段落的**位置**划分（body_blocks 索引），不是纯文本包含；
    #      避免"目录/清单里出现同名字样"被误当作章节内容。
    #   2) "规范"= 该引用位置本身必须是 Word 交叉引用域 (REF/NOTEREF)、
    #      书签引用域 (PAGEREF) 或指向参考文献书签的可跳转对象 (HYPERLINK\\l /
    #      w:hyperlink @anchor)——即前面 evaluate_reference_citations 的合规判据。
    #   3) 各章节的必需编号：
    #       前言 → {1, 2}
    #       研究方法章(5 个子节) → {3, 4, 5, 6}
    #       分析与讨论章(4 个子节) → {7, 8, ...}"及后续参考文献交叉引用"，
    #         即 7 起到全文最大编号的连续编号均需出现。
    section_map = _collect_section_regularized_refs(info)
    max_ref_num = _max_regularized_ref_num(info)
    ref_section_ok = True
    ref_section_details: List[str] = []
    for group_name, headings, required, and_after in SECTION_REF_REQUIREMENTS:
        got = set()
        for h in headings:
            got |= section_map.get(h, set())
        need = set(required)
        if and_after and max_ref_num >= max(required):
            # "[7]、[8]及后续": 需要覆盖 {7,8,...,max_ref_num}
            need = set(range(min(required), max_ref_num + 1))
        missing = sorted(need - got)
        if missing:
            ref_section_ok = False
            ref_section_details.append(f"{group_name} 缺少规范引用 {missing}")
    add_hit(
        hits, 5,
        "各指定章节至少保留并规范对应参考文献交叉引用(前言{1,2}/方法章{3-6}/分析讨论章{7,8,及后续})",
        ref_section_ok,
        f"最大编号={max_ref_num}；问题 {ref_section_details[:3]}"
    )

    # ============ 细则 +1 参考文献内容中每条文献编号阿拉伯数字编号，
    #            序号从1开始连续递增，序号后无标点符号，序号后空一格 ============
    # 细则拆点（一 一对应细则字面每一个点；均基于 Word/WPS 办公软件真实渲染）：
    #   点1）编号为**阿拉伯数字**（0-9 十进制数字，非罗马/中文/带圈序号）；
    #   点2）序号从 **1** 开始；
    #   点3）序号**连续递增**（[1..N]，无跳号、无重号）；
    #   点4）序号后**无标点符号**（不接受 "."/"、"/"．"/"："/")" 等任何标点）；
    #   点5）序号后**空一格**（紧跟 1 个 ASCII 空格，然后接正文）。
    # 办公软件下满足以上五点的等价渲染路径：
    #   (a) 段落文本以 "N␣正文" 开头（手动键入）；
    #   (b) 段落 numPr 指向自动列表 numFmt=decimal、lvlText="%1"、w:suff="space"，
    #       Word/WPS 将该段落渲染为 "N␣正文"。
    ok_ref_num, ref_num_detail = _evaluate_reference_list_numbering(info)
    ref_nums = ref_num_detail["numbers"]
    add_hit(
        hits, 1,
        "参考文献下方内容中每条文献编号阿拉伯数字编号，序号从1开始连续递增，序号后无标点符号，序号后空一格",
        ok_ref_num,
        f"编号 {ref_nums[:20]}；格式来源 {ref_num_detail['fmt_source']}"
    )

    # ============ 细则 +1 全文英文和阿拉伯数字字体为 Times New Roman,
    #            字号与所在段落中文字号一致 ============
    # 细则两个点(均基于 Word/WPS 办公软件真实渲染, 逐段逐 run 校验):
    #   点1) 含英文字母/阿拉伯数字(A-Za-z0-9)的 run, 其英数字体 = Times New Roman
    #        —— 取 ascii 或 hAnsi 生效字体名 (_run_ascii_font: run 直设 → 空视为继承默认);
    #   点2) 该英数 run 的生效字号 == 同段落中文字号
    #        —— 同段中文字号由 _para_chinese_size_hp 取 (首个含中文 run 的生效 sz, 半磅);
    #        —— 英数 run 生效字号由 _run_effective_size_hp 取 (run sz → 段落默认 sz).
    # 判定口径 (按建议"要求全部合格"):
    #   - 逐段: 仅当段落里存在中文参照 (cn_hp 可定) 时才比较字号; 无中文参照的段落
    #     (纯英数/无字号信息) 只校验字体, 不做字号比较 (无可比对象, 不算违规);
    #   - 只要有任一英数 run 违反字体或字号 → 整条 fail (不再用 90% 通过率).
    doc_font_bad_samples: List[str] = []
    doc_font_checked = 0
    for p in info.paragraphs:
        cn_hp = _para_chinese_size_hp(p)  # 同段中文字号(半磅), None 表示该段无中文参照
        for run in paragraph_runs(p):
            text = all_text(run)
            if not re.search(r"[A-Za-z0-9]", text):
                continue
            doc_font_checked += 1
            # 点1 字体: 生效 ascii/hAnsi 名 (空字符串视为未直设 → 继承, 不判违规)
            asc = _run_ascii_font(run)
            if asc and asc != "Times New Roman":
                doc_font_bad_samples.append(f"字体={asc!r}:{text.strip()[:8]!r}")
                continue
            # 点2 字号: 与同段中文字号一致(仅当该段有中文参照时比较)
            if cn_hp is not None:
                en_hp = _run_effective_size_hp(run, p)
                if en_hp is not None and en_hp != cn_hp:
                    doc_font_bad_samples.append(
                        f"字号={en_hp/2.0}pt≠中文{cn_hp/2.0}pt:{text.strip()[:8]!r}")
    doc_font_ok = doc_font_checked > 0 and not doc_font_bad_samples
    add_hit(hits, 1, "全文英文和阿拉伯数字字体为Times New Roman且字号与中文一致", doc_font_ok,
            f"英数 run 共 {doc_font_checked} 个；不合格样本 {doc_font_bad_samples[:5]}")

    # ============ 细则 +3 全文保留图1至图8共8张图片；图题依次为
    #            "研究技术路线图" ... "健康感知路径模型图"；
    #            所有图片均居中放置，图片左边界与右边界相对正文版心居中，
    #            左右偏差不超过 0.5 厘米；图片宽度设置为正文版心宽度的 55%-90%，
    #            且未超出页面可打印区域 ============
    # 细则精确拆分为 4 个点，逐点判定（针对办公软件 Word/WPS 渲染路径）：
    #   点1：全文"保留"图1至图8共 8 张图片（drawing 对象计数 ≥ 8）。
    #   点2：图题依次包含指定 8 个名称——按图序 1..8 出现的图题文本
    #        必须依序**包含**对应 FIGURE_TITLES 里的名称。
    #   点3：所有图片均居中放置（细则原文"均"，故对全部图片而非仅前 8 张评估）；
    #        且左右边界相对正文版心居中，偏差 ≤ 0.5 厘米（0.5cm ≈ 283 twips）。
    #        - 嵌入式：由段落 w:jc / w:ind 决定；
    #        - 浮动：由 wp:positionH 的 wp:align=center 或 wp:posOffset 决定。
    #   点4：每张图片宽度 ∈ [55%, 90%] × 正文版心宽度，且未超出可打印区域
    #        （即 ≤ 版心宽度）。以图片 wp:extent@cx (EMU) 计。
    fig_center_tol_twips = int(round(0.5 * CM_TO_TWIPS))  # 0.5 cm

    # 点1：数量
    count_ok = len(drawings) >= 8

    # 点2：图题依序包含指定 8 个名称
    figure_titles_seq_ok = True
    figure_titles_seq_detail: List[str] = []
    for i in range(1, 9):
        cap_idx = figure_captions.get(i)
        cap_text = texts[cap_idx] if cap_idx is not None else ""
        expected_name = FIGURE_TITLES[i - 1]
        if expected_name not in cap_text:
            figure_titles_seq_ok = False
            figure_titles_seq_detail.append(f"图{i} 缺'{expected_name}'")

    # 点3：所有图片居中（段落 jc 或浮动 positionH），且偏差 ≤ 0.5cm
    fig_center_ok = True
    fig_center_detail: List[str] = []
    for k, d in enumerate(drawings, 1):
        # 3a) 居中判定
        centered = False
        if d.get("inline") and d.get("paragraph") is not None:
            centered = para_jc(d["paragraph"]) == "center"
        elif d.get("anchor_el") is not None:
            positionH = d["anchor_el"].find("wp:positionH", NS)
            align_el = positionH.find("wp:align", NS) if positionH is not None else None
            rel_from = positionH.attrib.get("relativeFrom", "") if positionH is not None else ""
            centered = (
                align_el is not None
                and (align_el.text or "").strip() == "center"
                and rel_from in {"margin", "column"}
            )
        if not centered:
            fig_center_ok = False
            fig_center_detail.append(f"第{k}张 未居中")
            continue
        # 3b) 左右边界偏差 ≤ 0.5cm
        dev = _figure_center_deviation_twips(d, info)
        if dev is None or dev > fig_center_tol_twips:
            fig_center_ok = False
            fig_center_detail.append(f"第{k}张 偏差{dev}twips>{fig_center_tol_twips}")

    # 点4：宽度 ∈ [55%, 90%] × 版心；且未超出可打印区域
    fig_widths = [d["cx"] for d in drawings if d["cx"]]
    fig_width_ok = bool(fig_widths) and all(
        0.55 * info.printable_width_emu <= w <= 0.90 * info.printable_width_emu
        for w in fig_widths
    ) and all(w <= info.printable_width_emu for w in fig_widths)

    plus3_fig_passed = count_ok and figure_titles_seq_ok and fig_center_ok and fig_width_ok
    add_hit(
        hits, 3,
        "全文保留图1-图8、图题依序、均居中(左右偏差≤0.5cm)、宽度55%-90%版心、未越界",
        plus3_fig_passed,
        f"图片 {len(drawings)}；标题问题 {figure_titles_seq_detail[:3]}；"
        f"居中问题 {fig_center_detail[:3]}；宽度EMU {fig_widths[:8]}"
    )

    figure_caption_ok = True
    for number in range(1, 9):
        cap_idx = figure_captions.get(number)
        prev_has_drawing = cap_idx is not None and any(d["block_index"] < cap_idx and d["block_index"] >= cap_idx - 3 for d in drawings)
        text = texts[cap_idx] if cap_idx is not None else ""
        runs = paragraph_runs(info.body_blocks[cap_idx][1]) if cap_idx is not None and info.body_blocks[cap_idx][0] == "p" else []
        good, checked = font_check_runs(runs, None, 21)
        figure_caption_ok = figure_caption_ok and cap_idx is not None and prev_has_drawing and re.match(rf"^\s*图\s*{number}\s{{2,}}", text) and para_jc(info.body_blocks[cap_idx][1]) in {"center", ""} and checked > 0 and good / checked >= 0.8
    add_hit(hits, 3, "图1至图8位于图题上方且中文图题格式、字体、居中规范", figure_caption_ok, f"图题索引 {sorted(figure_captions.keys())}")

    # ============ 细则 +1 图的上方及图注下方各空一行；
    #            若图位于页首，上面不空行 ============
    # 逐点对齐细则（针对办公软件 Word/WPS 真实渲染）：
    #   点1：图的**上方**空一行 —— 图片所在段落的紧邻上一个段落必须为空段落
    #        （w:p 中无文字），Word/WPS 渲染时该空段落即"一行空白"。
    #   点2：图注的**下方**空一行 —— 图题段落的紧邻下一个段落必须为空段落。
    #   点3：若图**位于页首**，上面不空行（豁免点1）。
    #        办公软件下"位于页首"由以下任一 XML 结构决定：
    #          (a) 图片段落是正文首段（block_index==0）；
    #          (b) 上一段落含 <w:br w:type="page"/>（硬分页符）；
    #          (c) 图片段落自身 <w:pPr><w:pageBreakBefore/></w:pPr>（段前分页）；
    #          (d) 上一段落含 <w:pPr><w:sectPr/></w:pPr> 且该 sectPr 的
    #              <w:type w:val="..."/> ∈ {"", "nextPage", "oddPage", "evenPage"}
    #              （新一页的分节符；默认省略即 nextPage）。
    def _fig_at_page_top(fig_block_idx: int) -> bool:
        if fig_block_idx <= 0:
            return True
        cur_kind, cur_blk = info.body_blocks[fig_block_idx]
        # (c) 段前分页
        if cur_kind == "p":
            pbb = cur_blk.find(".//w:pPr/w:pageBreakBefore", NS)
            if pbb is not None and attr(pbb, "w:val", "1") not in {"0", "false"}:
                return True
        prev_kind, prev_blk = info.body_blocks[fig_block_idx - 1]
        if prev_kind == "p":
            # (b) 硬分页符
            for br in prev_blk.findall(".//w:br", NS):
                if attr(br, "w:type", "") == "page":
                    return True
            # (d) 分节符起新页
            sectPr = prev_blk.find(".//w:pPr/w:sectPr", NS)
            if sectPr is not None:
                sec_type_el = sectPr.find("w:type", NS)
                sec_type = attr(sec_type_el, "w:val", "") if sec_type_el is not None else ""
                if sec_type in {"", "nextPage", "oddPage", "evenPage"}:
                    return True
        return False

    def _is_empty_para(block_idx: int) -> bool:
        if block_idx < 0 or block_idx >= len(info.body_blocks):
            return False
        kind, blk = info.body_blocks[block_idx]
        if kind != "p":
            return False
        # 段落无任何文字（保留纯空白也视为空段落，办公软件渲染即空一行）
        return all_text(blk).strip() == ""

    blank_rule_ok = True
    blank_rule_detail: List[str] = []
    for number in range(1, 9):
        cap_idx = figure_captions.get(number)
        if cap_idx is None:
            blank_rule_ok = False
            blank_rule_detail.append(f"图{number} 无题注")
            continue
        # 图片段落：图题段落之前紧邻的、含 w:drawing 的段落
        fig_block_idx: Optional[int] = None
        for d in drawings:
            if d["block_index"] < cap_idx and d["block_index"] >= cap_idx - 4:
                fig_block_idx = d["block_index"] if fig_block_idx is None else max(fig_block_idx, d["block_index"])
        if fig_block_idx is None:
            blank_rule_ok = False
            blank_rule_detail.append(f"图{number} 未定位图片段落")
            continue
        # 点1 + 点3：图上方空一行；页首豁免
        if not _fig_at_page_top(fig_block_idx):
            if not _is_empty_para(fig_block_idx - 1):
                blank_rule_ok = False
                blank_rule_detail.append(f"图{number} 上方缺空行")
        # 点2：图注下方空一行
        if not _is_empty_para(cap_idx + 1):
            blank_rule_ok = False
            blank_rule_detail.append(f"图{number} 图注下方缺空行")

    add_hit(
        hits, 1,
        "图上方空一行、图注下方空一行；位于页首时上方免空行",
        blank_rule_ok and bool(figure_captions),
        f"问题 {blank_rule_detail[:5]}"
    )

    # ============ 细则 +3 全文所有英文图题位于中文图题下方；Times New Roman 五号；居中显示 ============
    # 说明：rubric 只要求英文图题在中文图题"下方"，不限制承载形式。办公软件里
    # 存在两种等价的"下方"写法，均需支持：
    #   形式A（同段软换行）：中文图题与英文图题写在同一个段落，中间用 w:br 软换行分行——
    #     Word/WPS 渲染为紧邻下一行；此时取该段落最后一个 w:br/w:cr 之后的 run 作为英文行。
    #   形式B（独立下一段）：英文图题独立成为下一非空段落；仅当下一非空段落形如
    #     "Figure N ..." 时视为"下方"的英文图题。
    # 判定顺序：优先按形式A；A 未命中时回退到形式B（下一非空段落）。
    # 每种形式下：
    #   - "位于下方"：形式A由段内软换行保证，形式B由段序保证；
    #   - "字体 Times New Roman 五号"：英文行 run 的 w:rFonts@w:ascii 为 "Times New Roman"
    #     且 w:sz@w:val=21（半磅），未显式字体时按段落 rPr 继承的默认 ascii 处理；
    #   - "居中显示"：形式A取中文图题段的 w:jc(同段落无法分行单独对齐)；
    #     形式B取英文图题独立段的 w:jc。
    en_fig_ok = True
    en_fig_details: List[str] = []
    for number in sorted(figure_captions.keys()):
        cn_idx = figure_captions.get(number)
        if cn_idx is None:
            en_fig_ok = False
            en_fig_details.append(f"图{number} 未找到中文图题")
            continue

        if info.body_blocks[cn_idx][0] != "p":
            en_fig_ok = False
            en_fig_details.append(f"图{number} 题注定位异常")
            continue
        cap_block = info.body_blocks[cn_idx][1]

        # 形式A：同段最后一个 w:br 之后的 run
        runs_all = paragraph_runs(cap_block)
        last_break_pos = -1
        for ridx, run in enumerate(runs_all):
            if run.find(".//w:br", NS) is not None or run.find(".//w:cr", NS) is not None:
                last_break_pos = ridx
        en_runs: List[ET.Element] = []
        en_container: Optional[ET.Element] = None
        en_source = ""
        if last_break_pos >= 0:
            candidate_runs = runs_all[last_break_pos + 1:]
            candidate_text = "".join(all_text(r) for r in candidate_runs).strip()
            if candidate_text:
                en_runs = candidate_runs
                en_container = cap_block
                en_source = "same-para-softbreak"

        # 形式B：下一非空段落
        if not en_runs:
            j = cn_idx + 1
            while j < len(info.body_blocks):
                kind_j, block_j = info.body_blocks[j]
                if kind_j == "p":
                    txt_j = all_text(block_j).strip()
                    if txt_j:
                        en_runs = paragraph_runs(block_j)
                        en_container = block_j
                        en_source = "next-para"
                        break
                    j += 1
                    continue
                # 非段落(表格等):不再向后视为"下方"英文图题
                break

        if not en_runs or en_container is None:
            en_fig_ok = False
            en_fig_details.append(f"图{number} 中文图题下方缺少英文图题")
            continue

        en_text = "".join(all_text(r) for r in en_runs).strip()
        if not re.match(rf"^\s*Figure\s*{number}\b", en_text, re.IGNORECASE):
            en_fig_ok = False
            en_fig_details.append(f"图{number} 英文图题格式不符({en_source}): {en_text[:40]!r}")
            continue

        # 居中：形式A使用中文题注段 w:jc；形式B使用英文题注段自身的 w:jc
        if para_jc(en_container) != "center":
            en_fig_ok = False
            en_fig_details.append(f"图{number} 英文图题未居中({en_source})")
            continue

        checked_runs = 0
        good_runs = 0
        for run in en_runs:
            if not all_text(run).strip():
                continue
            checked_runs += 1
            ascii_font = _run_ascii_font(run)
            size_hp = run_size_half_points(run)
            if ascii_font in {"Times New Roman", ""} and size_hp == 21:
                good_runs += 1
        if checked_runs == 0 or good_runs != checked_runs:
            en_fig_ok = False
            en_fig_details.append(
                f"图{number} 英文图题字体/字号不符({en_source}): 合格 {good_runs}/{checked_runs}"
            )
            continue

    add_hit(
        hits, 3,
        "全文所有英文图题位于中文图题下方，字体为Times New Roman五号，居中显示",
        en_fig_ok and bool(figure_captions),
        f"图题索引 {sorted(figure_captions.keys())}；问题 {en_fig_details[:5]}"
    )

    # ============ 细则 +5 全文中文正文段落中的中文标点使用全角标点，
    #            包括 "，" "。" "；" "：" "（" "）" ============
    # 逐点对齐细则（针对办公软件 Word/WPS 真实字符渲染）：
    #   点1：判定范围 = "中文正文段落" —— 严格按"正文段落"限定：
    #        (a) 只统计 body_blocks 里 kind=="p" 的段落，排除表格单元格
    #            （表格里的表头/表体不属于"正文段落"）；
    #        (b) 排除"参考文献"章节及其后（致谢/附录）的所有段落——
    #            参考文献条目是引用格式，不是中文正文；
    #        (c) 剩余段落里，凡文本含中文字符者视为中文正文段落。
    #   点2..7：细则明确列出的 6 种半角标点必须全部替换为对应全角：
    #        "," -> "，"；"." -> "。"；";" -> "；"；":" -> "："；"(" -> "（"；")" -> "）"
    #        每种独立计数并输出，任何一种在中文段落中出现即视为违规。
    #          - "," / ";" / ":" / "(" / ")" 只要出现在中文段落内即算违规。
    #          - "." 需排除数字小数点(如 1.5)与英文缩写内的点(如 e.g.)，
    #            即只算"中文字符紧邻半角句点"或"句点紧邻中文字符"的情形。
    #   点8：判定阈值 = 0 —— 细则原文用"使用"(未给容忍数)，严格判定为"完全
    #        不允许出现"。既有的 -1 扣分项"超过 10 处"仍复用 bad_cn_punc，不受影响。
    cn_punc_counts: Dict[str, int] = {c: 0 for c in [",", ".", ";", ":", "(", ")"]}
    cn_punc_examples: List[str] = []
    # 定位"参考文献"章节的起始索引——从该段（含）起后续所有段落全部排除。
    ref_body_start_idx: Optional[int] = None
    for _bi, (_kind, _blk) in enumerate(info.body_blocks):
        if _kind != "p":
            continue
        _t = all_text(_blk).replace(" ", "").replace("　", "")
        if _t == "参考文献":
            ref_body_start_idx = _bi
            break
    for _bi, (_kind, _blk) in enumerate(info.body_blocks):
        if _kind != "p":
            continue  # 点1(a)：排除表格单元格
        if ref_body_start_idx is not None and _bi >= ref_body_start_idx:
            break  # 点1(b)：排除参考文献章节及其后
        para_text = all_text(_blk)
        if not CHINESE_RE.search(para_text):
            continue  # 点1(c)：仅统计中文正文段落
        for ch in [",", ";", ":", "(", ")"]:
            cn_punc_counts[ch] += para_text.count(ch)
        for m in re.finditer(r"\.", para_text):
            i = m.start()
            left = para_text[i - 1] if i > 0 else ""
            right = para_text[i + 1] if i + 1 < len(para_text) else ""
            if (left and CHINESE_RE.search(left)) or (right and CHINESE_RE.search(right)):
                cn_punc_counts["."] += 1
        if len(cn_punc_examples) < 3 and re.search(r"[一-鿿][,\.:;()]|[,\.:;()][一-鿿]", para_text):
            cn_punc_examples.append(para_text[:60])
    strict_cn_punc_bad = sum(cn_punc_counts.values())
    # 复用给后续 -1 扣分项("超过 10 处半角"): 与旧变量名保持兼容
    bad_cn_punc = strict_cn_punc_bad
    add_hit(
        hits, 5,
        "全文中文正文段落中的中文标点使用全角(，。；：（）)",
        strict_cn_punc_bad == 0,
        f"半角逗{cn_punc_counts[',']}/句{cn_punc_counts['.']}/分{cn_punc_counts[';']}/"
        f"冒{cn_punc_counts[':']}/左括{cn_punc_counts['(']}/右括{cn_punc_counts[')']}；"
        f"样例 {cn_punc_examples}"
    )

    # ============ 细则 +5 全文英文摘要、英文题名、英文图题、英文表题、英文参考文献
    #            中的英文标点使用半角标点，包括 "," "." ";" ":" "(" ")" ============
    # 逐点对齐细则（针对办公软件 Word/WPS 真实字符渲染，判定基于段落文本字符）：
    #   点1..5 判定范围（五个"英文*"子范围逐一定位并合并）：
    #     (a) 英文题名 —— 文档前部第一段以 ASCII 字母开头、长度较大、位于中文题名
    #         之后的段落。以段落形态匹配。
    #     (b) 英文摘要 —— "Abstract" 段落到下一个非英文段落之前的所有段落
    #         (含 Objective/Methods/Results/Conclusions 各分节)。
    #     (c) 英文关键词 —— "Key Words:" 起始段落，作为英文摘要延伸算入 (b)。
    #     (d) 英文图题 —— 每个中文图题段之后紧邻的英文图题段（本项已由 +3 项定位，
    #         此处沿用 figure_captions 的下一非空段）。
    #     (e) 英文表题 —— 同 (d)，用 table_captions。
    #     (f) 英文参考文献 —— 参考文献章节内**含英文字符**的条目段（作者/期刊为英文
    #         的条目）；纯中文条目不入本范围。
    #   点6..11 半角标点判定：细则明确的 6 种"应为半角"标点，其对应**全角形态**
    #         出现即视为违规：
    #             "，" / "。" / "；" / "：" / "（" / "）"
    #         每种独立计数，任一大于 0 即整项不通过。
    #   点12 判定阈值 = 0（细则原文用"使用"，未给容忍数）。既有 -1 扣分项
    #        "英文摘要中出现超过 5 处..." 仍复用 bad_en_punc（保持向下兼容）。
    def _para_is_english_dominant(text: str) -> bool:
        # 段落英文字符数 > 中文字符数，且有英文
        en = len(ENGLISH_RE.findall(text))
        cn = len(CHINESE_RE.findall(text))
        return en > 0 and en > cn

    en_scope_texts: List[str] = []
    texts_arr = info.texts

    # (a) 英文题名：正文首 30 段内第一段"以字母开头且英文占优"的长段
    for i, t in enumerate(texts_arr[:30]):
        s = t.strip()
        if len(s) >= 20 and re.match(r"^[A-Za-z]", s) and _para_is_english_dominant(s):
            en_scope_texts.append(s)
            break

    # (b)(c) 英文摘要 + 英文关键词：定位 "Abstract" 段落（非目录项）
    abstract_idx: Optional[int] = None
    for i, t in enumerate(texts_arr):
        s = t.strip()
        # 目录项形如 "Abstract3"，排除
        if s.lower() == "abstract" or (s.startswith("Abstract") and not re.search(r"\d\s*$", s)):
            abstract_idx = i
            break
    if abstract_idx is not None:
        j = abstract_idx + 1
        while j < len(texts_arr):
            s = texts_arr[j].strip()
            if not s:
                j += 1
                continue
            # 首个"非英文占优且非空"段视为摘要区结束
            if not _para_is_english_dominant(s):
                break
            en_scope_texts.append(s)
            j += 1
        # 再向下延伸抓 Key Words 段（可能被空段隔开）
        for k in range(j, min(j + 4, len(texts_arr))):
            s = texts_arr[k].strip()
            if re.match(r"^Key\s*Words?\s*[:：]", s, re.IGNORECASE):
                en_scope_texts.append(s)
                break

    # (d)(e) 英文图题/表题：范围识别与"+3 英文图题/表题"格式项保持一致——
    #   同时支持两种"下方"写法：
    #     形式A：中文题注段内最后一个 w:br 之后的 run 段（同段软换行）
    #     形式B：中文题注段之后紧邻的下一个非空段落
    # 遇到 A 命中则以 A 的文本入范围；A 未命中时再回退到 B。
    def _collect_en_caption_text(cn_idx: Optional[int]) -> Optional[str]:
        if cn_idx is None or cn_idx < 0 or cn_idx >= len(info.body_blocks):
            return None
        kind_cn, blk_cn = info.body_blocks[cn_idx]
        if kind_cn != "p":
            return None
        # 形式A：同段 w:br 之后
        runs_all = paragraph_runs(blk_cn)
        last_break_pos = -1
        for ridx, r in enumerate(runs_all):
            if r.find(".//w:br", NS) is not None or r.find(".//w:cr", NS) is not None:
                last_break_pos = ridx
        if last_break_pos >= 0:
            after_text = "".join(all_text(r) for r in runs_all[last_break_pos + 1:]).strip()
            if after_text and ENGLISH_RE.search(after_text):
                return after_text
        # 形式B：下一非空段落
        j = cn_idx + 1
        while j < len(info.body_blocks):
            kind, blk = info.body_blocks[j]
            if kind == "p":
                st = all_text(blk).strip()
                if st:
                    return st if ENGLISH_RE.search(st) else None
                j += 1
                continue
            return None
        return None

    # (d) 英文图题
    for _, cap_idx in figure_captions.items():
        s = _collect_en_caption_text(cap_idx)
        if s:
            en_scope_texts.append(s)

    # (e) 英文表题
    for _, cap_idx in table_captions.items():
        s = _collect_en_caption_text(cap_idx)
        if s:
            en_scope_texts.append(s)

    # (f) 英文参考文献：参考文献章节内含英文字符的条目段
    ref_start_idx: Optional[int] = None
    for i, (kind, blk) in enumerate(info.body_blocks):
        if kind != "p":
            continue
        s = all_text(blk).strip().replace(" ", "").replace("　", "")
        if s == "参考文献":
            ref_start_idx = i
            break
    if ref_start_idx is not None:
        for j in range(ref_start_idx + 1, len(info.body_blocks)):
            kind, blk = info.body_blocks[j]
            if kind != "p":
                continue
            s = all_text(blk).strip()
            if s.startswith("致谢") or s.startswith("附录"):
                break
            if s and ENGLISH_RE.search(s) and _para_is_english_dominant(s):
                en_scope_texts.append(s)

    # 汇总五个子范围内 6 种全角标点的出现次数（每种独立计数）
    en_punc_map = {"，": ",", "。": ".", "；": ";", "：": ":", "（": "(", "）": ")"}
    en_punc_counts: Dict[str, int] = {full: 0 for full in en_punc_map}
    en_punc_examples: List[str] = []
    for para in en_scope_texts:
        hit_here = False
        for full in en_punc_map:
            n = para.count(full)
            if n:
                en_punc_counts[full] += n
                hit_here = True
        if hit_here and len(en_punc_examples) < 3:
            en_punc_examples.append(para[:80])

    strict_en_punc_bad = sum(en_punc_counts.values())
    # 兼容既有 -1 扣分项 bad_en_punc
    bad_en_punc = strict_en_punc_bad
    add_hit(
        hits, 5,
        "英文摘要/题名/图题/表题/参考文献的标点使用半角(,.;:())",
        strict_en_punc_bad == 0,
        f"全角逗{en_punc_counts['，']}/句{en_punc_counts['。']}/分{en_punc_counts['；']}/"
        f"冒{en_punc_counts['：']}/左括{en_punc_counts['（']}/右括{en_punc_counts['）']}；"
        f"扫描段落 {len(en_scope_texts)}；样例 {en_punc_examples}"
    )

    # ============ 细则 +5 全文阿拉伯数字、英文字母、统计符号和公式中的标点
    #            使用半角符号；示例 300 m、78.3%、Cronbach α、P<0.01、R²、
    #            500 m、800 m、1000 m ============
    # 细则的核心约束是"以下三类语境中的**标点**必须为半角"。逐点对齐（针对
    # Word/WPS 办公软件真实字符渲染，判定基于字符 code point 本身）：
    #
    #   语境1：阿拉伯数字上下文
    #     - 数字与单位之间的空格/标点（示例 "300 m"、"1000 m"）；
    #     - 数字与百分号（示例 "78.3%"）；
    #     - 数字内小数点（示例 "78.3"）。
    #     违规形态（全角）：
    #       - 全角小数点：数字之间 "．" 或 "。"（如 "78．3" / "78。3"）；
    #       - 全角百分号 "％" 紧邻数字（如 "78.3％"）；
    #       - 全角括号/逗号/分号/冒号 紧邻数字（如 "（3）" "3，" "3；" "3："）；
    #       - 全角减号/斜杠等虽细则未列 6 种以外的标点，但"数字上下文标点"这一
    #         语境明确要求半角——保守起见仅检测细则示例语境相关的 6 种全角标点
    #         紧邻数字的情形。
    #
    #   语境2：英文字母上下文
    #     - 字母与紧邻标点（示例 "Cronbach α" 空格；"P<0.01" 尖括号视为运算符
    #       半角，此处只管标点：字母紧邻全角标点即违规）。
    #     违规形态：全角逗号/句号/分号/冒号/左右括号 紧邻 ASCII 字母。
    #
    #   语境3：统计符号与公式中的标点
    #     - 示例 "P<0.01"、"R²"，其中"<"和"²"属数学符号(非标点)，本项只判
    #       "标点"——即公式/统计式中若使用了全角比较号(＜/＞/＝)或全角减号(－)等，
    #       视为违规。示例中已给出：P 后接比较号必须为半角。
    #     违规形态：
    #       - 大/小写 P/R 后紧邻全角 ＜/＞/＝；
    #       - 公式中"数字 全角运算符 数字"(如 3＋4)。
    #
    # 判定阈值 = 0（细则原文用"使用"），任一违规即整项不通过。
    stat_full_punc = "，。；：（）"  # 6 种全角标点（与本项相关）
    bad_stat = 0
    bad_stat_examples: List[str] = []

    def _add_stat_hit(sample: str):
        nonlocal bad_stat
        bad_stat += 1
        if len(bad_stat_examples) < 5 and sample not in bad_stat_examples:
            bad_stat_examples.append(sample)

    # 中文正文自然叙述句中的"中文逗号/句号紧邻数字"排除（用户要求）——
    # 形如  [中文字符][，或。][数字]  或  [数字][，或。][中文字符]
    # 视为中文正文的自然叙述标点，属于"中文正文段落"范畴，本项不计入违规。
    # 其余情形（例如 [数字][，][数字]、[数字][；][数字] 等）仍严格判违规。
    narrative_punc = set("，。")
    _CHINESE_CHAR_RE = re.compile(r"[一-鿿]")
    for m in re.finditer(rf"\d[{stat_full_punc}]|[{stat_full_punc}]\d", all_text_joined):
        matched = m.group()
        # 形态: [全角标点][数字]——检查左邻字符
        if matched[0] in narrative_punc:
            i = m.start() - 1
            if i >= 0 and _CHINESE_CHAR_RE.match(all_text_joined[i]):
                continue
        # 形态: [数字][全角标点]——检查右邻字符
        elif matched[1] in narrative_punc:
            j = m.end()
            if j < len(all_text_joined) and _CHINESE_CHAR_RE.match(all_text_joined[j]):
                continue
        _add_stat_hit(all_text_joined[max(0, m.start() - 5): m.end() + 5])
    # 数字之间的全角小数点/句号
    for m in re.finditer(r"\d[．。]\d", all_text_joined):
        _add_stat_hit(all_text_joined[max(0, m.start() - 5): m.end() + 5])
    # 全角百分号紧邻数字
    for m in re.finditer(r"\d\s*％", all_text_joined):
        _add_stat_hit(all_text_joined[max(0, m.start() - 5): m.end() + 5])
    # 字母紧邻上述 6 种全角标点
    for m in re.finditer(rf"[A-Za-z][{stat_full_punc}]|[{stat_full_punc}][A-Za-z]", all_text_joined):
        _add_stat_hit(all_text_joined[max(0, m.start() - 5): m.end() + 5])
    # 统计比较号/等号全角形态（P<、R=、数字之间±/－ 等）
    for m in re.finditer(r"[A-Za-z0-9]\s*[＜＞＝]|[＜＞＝]\s*[A-Za-z0-9]", all_text_joined):
        _add_stat_hit(all_text_joined[max(0, m.start() - 5): m.end() + 5])
    # 公式中数字-全角运算符-数字（±／＋－×÷ 的全角形态）
    for m in re.finditer(r"\d\s*[＋－×÷／]\s*\d", all_text_joined):
        _add_stat_hit(all_text_joined[max(0, m.start() - 5): m.end() + 5])

    add_hit(
        hits, 5,
        "数字/英文字母/统计符号/公式中的标点使用半角",
        bad_stat == 0,
        f"违规 {bad_stat} 处；样例 {bad_stat_examples}"
    )

    # ============ 细则 +1 插图和附表清单标题字体为小三号黑体，居中显示 ============
    # 细则拆点（针对 Word/WPS 真实渲染）：
    #   点1：字体 = 黑体（eastAsia 字体名为 "黑体"/"SimHei"，办公软件下即以黑体渲染）
    #   点2：字号 = 小三号（Word 小三 = 15pt = 30 半磅点，sz=30）
    #   点3：居中显示（段落 w:jc = "center"；不接受未设置回退，
    #                Word 默认对齐为左对齐，办公软件下未设置≠居中）
    # 只要"插图和附表清单"这一标题段落同时满足以上 3 点即通过。
    list_title_found = False
    list_title_center = False
    list_title_font_ok = False
    list_title_size_ok = False
    for p in info.paragraphs:
        if "插图和附表清单" not in all_text(p):
            continue
        list_title_found = True
        title_runs = [r for r in paragraph_runs(p) if all_text(r).strip()]
        if not title_runs:
            continue
        center_ok = para_jc(p) == "center"
        font_ok = all(
            run_font_names(r).get("eastAsia") in {"SimHei", "黑体", "simhei"}
            for r in title_runs
        )
        size_ok = all(run_size_half_points(r) == 30 for r in title_runs)
        if center_ok:
            list_title_center = True
        if font_ok:
            list_title_font_ok = True
        if size_ok:
            list_title_size_ok = True
        if center_ok and font_ok and size_ok:
            break
    list_title_ok = list_title_found and list_title_center and list_title_font_ok and list_title_size_ok
    add_hit(
        hits, 1,
        "插图和附表清单标题字体为小三号黑体，居中显示",
        list_title_ok,
        f"标题存在={list_title_found}; 黑体={list_title_font_ok}; 小三号(30半磅点)={list_title_size_ok}; 居中={list_title_center}"
    )
    # ============ 细则 +1 插图和附表清单内容字体为四号宋体，英文和数字为
    #                       Times New Roman 四号 ============
    # 细则拆点（针对 Word/WPS 真实渲染）：
    #   点1：字号 = 四号（Word 四号 = 14pt = 28 半磅点，sz=28）
    #        对内容段落中每个非空 run 生效。
    #   点2：中文字符 = 宋体（eastAsia 字体名 ∈ {"SimSun","宋体"}）
    #        仅当 run 文本包含中文字符时校验 eastAsia 字体。
    #   点3：英文和数字字符 = Times New Roman
    #        （ascii/hAnsi 字体名 ∈ {"Times New Roman","TimesNewRoman"}）
    #        仅当 run 文本包含 A-Za-z0-9 时校验 ascii/hAnsi 字体。
    # 内容范围：从"插图和附表清单"标题段之后到下一章节标题之前的所有非空段落。
    list_content = section_text(info, ["插图和附表清单"], ["摘要", "前言", "目录"])
    list_content_items_ok = all(t in list_content for t in FIGURE_TITLES[:8]) and all(t in list_content for t in TABLE_TITLES[:10])

    list_size_ok = True
    list_cn_font_ok = True
    list_en_font_ok = True
    size_bad_ex: List[str] = []
    cn_font_bad_ex: List[str] = []
    en_font_bad_ex: List[str] = []

    list_start_idx = None
    for i, text in enumerate(texts):
        stripped = text.replace(" ", "").replace("　", "")
        if "插图和附表清单" in text or "插图和附表清单" in stripped:
            clean = stripped.rstrip("0123456789")
            if clean != stripped and re.match(r"^[一-鿿\w]+\d+$", stripped):
                continue
            list_start_idx = i
            break

    if list_start_idx is not None:
        cn_ok_fonts = {"SimSun", "宋体", "simsun"}
        en_ok_fonts = {"Times New Roman", "TimesNewRoman", "times new roman"}
        for idx in range(list_start_idx + 1, len(info.body_blocks)):
            block = info.body_blocks[idx]
            if block[0] != "p":
                continue
            para_text = texts[idx] if idx < len(texts) else ""
            if not para_text.strip():
                continue
            # 遇到下一章节标题即停止（宽松地按目录常见截断词）
            if any(k in para_text for k in ["摘要", "前言", "目录", "缩略语表", "学位论文原创性"]):
                break
            for run in paragraph_runs(block[1]):
                rt = all_text(run)
                if not rt.strip():
                    continue
                # 点1 字号
                size = run_size_half_points(run)
                if size is not None and size != 28:
                    list_size_ok = False
                    if len(size_bad_ex) < 3:
                        size_bad_ex.append(f"{rt[:15]}(sz={size})")
                fonts = run_font_names(run)
                # 点2 中文 → 宋体
                if CHINESE_RE.search(rt):
                    ea = fonts.get("eastAsia", "")
                    if ea and ea not in cn_ok_fonts:
                        list_cn_font_ok = False
                        if len(cn_font_bad_ex) < 3:
                            cn_font_bad_ex.append(f"{rt[:15]}(eastAsia={ea})")
                    elif not ea:
                        list_cn_font_ok = False
                        if len(cn_font_bad_ex) < 3:
                            cn_font_bad_ex.append(f"{rt[:15]}(eastAsia=未设置)")
                # 点3 英文和数字 → Times New Roman
                if re.search(r"[A-Za-z0-9]", rt):
                    ascii_f = fonts.get("ascii", "")
                    hansi_f = fonts.get("hAnsi", "")
                    if ascii_f and ascii_f not in en_ok_fonts:
                        list_en_font_ok = False
                        if len(en_font_bad_ex) < 3:
                            en_font_bad_ex.append(f"{rt[:15]}(ascii={ascii_f})")
                    elif hansi_f and hansi_f not in en_ok_fonts:
                        list_en_font_ok = False
                        if len(en_font_bad_ex) < 3:
                            en_font_bad_ex.append(f"{rt[:15]}(hAnsi={hansi_f})")
                    elif not ascii_f and not hansi_f:
                        list_en_font_ok = False
                        if len(en_font_bad_ex) < 3:
                            en_font_bad_ex.append(f"{rt[:15]}(ascii/hAnsi=未设置)")
    else:
        list_size_ok = False
        list_cn_font_ok = False
        list_en_font_ok = False

    list_content_font_ok = list_size_ok and list_cn_font_ok and list_en_font_ok
    add_hit(
        hits, 1,
        "插图和附表清单内容字体为四号宋体，英文和数字为Times New Roman四号",
        list_content_items_ok and list_content_font_ok,
        f"目录项完整={list_content_items_ok}; 四号(28)={list_size_ok}({size_bad_ex}); 中文宋体={list_cn_font_ok}({cn_font_bad_ex}); 英数TNR={list_en_font_ok}({en_font_bad_ex})"
    )

    negative_checks(info, hits, all_text_joined, drawings, figure_captions, table_captions, ref_nums, citation_count, ref_field_count, hyperlink_count, bad_cn_punc, bad_en_punc, bad_stat, list_content)
    return hits


def negative_checks(info: WordInfo, hits: List[Hit], text: str, drawings: List[Dict[str, object]], figure_captions: Dict[int, int], table_captions: Dict[int, int], ref_nums: List[int], citation_count: int, ref_field_count: int, hyperlink_count: int, bad_cn_punc: int, bad_en_punc: int, bad_stat: int, list_content: str):
    phrase_pages = info.com.get("phrase_pages") or {}
    # ============ 细则 -3 "学位论文原创性承诺"标题及内容未放置在论文第三页顶部 ============
    # 细则拆点（Word/WPS 真实分页与页面顶部渲染）：
    #   点1：标题"学位论文原创性承诺"必须出现在第三页；
    #   点2：内容（标题下方的承诺书正文）必须与标题同处第三页；
    #   点3：标题及内容位于第三页"顶部"——即标题为第三页的第一处非空白内容，
    #        前面不得先出现其他章节文本（如封面残余、目录、摘要等）。
    # 判定基于 Word 自身分页结果 doc.ComputeStatistics/GoTo Page（info.com.page_texts），
    # 与办公软件所见即所得一致。
    pledge_phrase = "学位论文原创性承诺"
    pledge_page = phrase_pages.get(pledge_phrase)  # None 表示未找到
    page_texts_all = info.com.get("page_texts") or []
    pledge_title_ok = pledge_page == 3
    pledge_content_ok = False
    pledge_top_ok = False
    if pledge_title_ok and len(page_texts_all) >= 3:
        p3 = page_texts_all[2] or ""
        # 点3 顶部：以第三页文本剔除起始空白/换行/制表/软回车后，
        #          必须以"学位论文原创性承诺"开头。
        p3_lstrip = re.sub(r"^[\s　\r\n\t\x07]+", "", p3)
        if p3_lstrip.startswith(pledge_phrase):
            pledge_top_ok = True
        # 点2 内容：标题后同页需存在承诺书正文（非仅一个孤立标题）。
        after_title = p3.split(pledge_phrase, 1)[1] if pledge_phrase in p3 else ""
        if len(re.sub(r"\s+", "", after_title)) >= 20:
            pledge_content_ok = True
    if not (pledge_title_ok and pledge_content_ok and pledge_top_ok):
        add_hit(
            hits, -3,
            "学位论文原创性承诺标题及内容未放置在论文第三页顶部",
            True,
            f"标题所在页={pledge_page}; 位于第三页={pledge_title_ok}; 位于页顶={pledge_top_ok}; 同页有正文={pledge_content_ok}"
        )
    if "插图和附表清单" in text:
        pass  # 该 -3 扣分项已按用户要求删除
    # ============ 细则 -3 缩略语表表格不满足三线表 ============
    # 细则拆点（针对 Word/WPS 表格边框真实渲染）：
    #   点1：仅保留三类主横线——表格上边线、表头下方横线、表格下边线；
    #   点2：上边线线宽 = 1.5 磅（Word sz=12，1/8pt 单位）；
    #   点3：下边线线宽 = 1.5 磅（Word sz=12）；
    #   点4：表头下方横线线宽 = 0.75 磅（Word sz=6）；
    #        隐含：无左右边线、无表内竖线、除表头外无其他主横线。
    # 判定使用已封装的 is_three_line_table()——该函数严格按上述四类边线检查
    # tblBorders + 单元格 tcBorders 两级，办公软件所见即所得。
    abbrev_idx = next((i for i, t in enumerate(block_texts(info)) if "缩略语" in t.replace(" ", "").replace("　", "") and not re.match(r"^.{0,20}\d+$", t.replace(" ", "").replace("　", ""))), None)
    if abbrev_idx is not None:
        next_table = next((b for i, b in enumerate(info.body_blocks[abbrev_idx + 1:], abbrev_idx + 1) if b[0] == "tbl"), None)
        if next_table:
            atbl = next_table[1]
            ab = table_border_summary(atbl)
            arows = table_rows(atbl)
            # 分项状态供 detail 定位——按办公软件渲染优先级：
            # 单元格 tcBorders 若显式声明（哪怕是 nil），会覆盖 tblBorders 同名边。
            first_cells = arows[0].findall(".//w:tc", NS) if arows else []
            first_cells_top = [_cell_border(c, "top") for c in first_cells]
            any_ftop_declared = any(spec[0] not in {"", None} for spec in first_cells_top)
            if any_ftop_declared:
                top_ok = bool(first_cells) and all(
                    spec[0] == "single" and (spec[1] or 0) in range(11, 14)
                    for spec in first_cells_top
                )
            else:
                top_ok = ab.get("top", ("", None))[0] == "single" and (ab.get("top", ("", None))[1] or 0) in range(11, 14)
            last_cells = arows[-1].findall(".//w:tc", NS) if arows else []
            last_cells_bot = [_cell_border(c, "bottom") for c in last_cells]
            any_lbot_declared = any(spec[0] not in {"", None} for spec in last_cells_bot)
            if any_lbot_declared:
                bottom_ok = bool(last_cells) and all(
                    spec[0] == "single" and (spec[1] or 0) in range(11, 14)
                    for spec in last_cells_bot
                )
            else:
                bottom_ok = ab.get("bottom", ("", None))[0] == "single" and (ab.get("bottom", ("", None))[1] or 0) in range(11, 14)
            header_line_ok = False
            if arows:
                header_cells = arows[0].findall(".//w:tc", NS)
                header_line_ok = bool(header_cells) and all(
                    _cell_border(c, "bottom")[0] == "single" and (_cell_border(c, "bottom")[1] or 0) in range(5, 8)
                    for c in header_cells
                )
            only_three_lines_ok = is_three_line_table(atbl)
            abbrev_ok = top_ok and bottom_ok and header_line_ok and only_three_lines_ok
            if not abbrev_ok:
                add_hit(
                    hits, -3,
                    "缩略语表表格不满足三线表",
                    True,
                    f"仅三线={only_three_lines_ok}; 上1.5磅={top_ok}; 下1.5磅={bottom_ok}; 表头下0.75磅={header_line_ok}"
                )
    if TABLE_TITLES[0] not in text:
        pass  # -1 "表1 研究变量与资料来源"未出现 已按用户要求删除
    grid_tables = sum(1 for t in info.tables if not is_none_border(table_border_summary(t).get("insideV", ("", None))) and not is_none_border(table_border_summary(t).get("insideH", ("", None))))
    if text.count(TABLE_TITLES[0]) > 1:
        pass  # -1 出现两张相同的"表1"任意1个表格仍保留完整网格线 已按用户要求删除
    if grid_tables > 0:
        add_hit(hits, -1, "任意1个表格仍保留完整网格线", True, f"完整网格线表格 {grid_tables} 个")
    if any(not table_row_heights_ok(t) for t in info.tables):
        pass  # -1 任意1个表格文字被边框裁切 已按用户要求删除
    if citation_count > 0 and ref_field_count + hyperlink_count == 0:
        pass  # -5 正文参考文献引用全部不是交叉引用 已按用户要求删除
    bad_citation_count = 0
    if bad_citation_count > 5:
        pass  # -3 正文中超过5处引用编号不是方括号格式 已按用户要求删除
    non_sup = 0
    for p in info.paragraphs:
        ptext = all_text(p)
        if CITATION_RE.search(ptext):
            runs = paragraph_runs(p)
            # Check if any run near brackets is superscript
            has_sup_citation = False
            for i, r in enumerate(runs):
                rtext = all_text(r)
                if CITATION_RE.search(rtext) or (rtext.strip().isdigit() and i > 0 and all_text(runs[i-1]).strip() == "["):
                    if run_is_superscript(r):
                        has_sup_citation = True
                        break
            if not has_sup_citation:
                non_sup += 1
    if non_sup > 5:
        pass  # -3 正文中超过5处引用编号不是上标格式 已按用户要求删除
    if citation_count > 5 and ref_field_count + hyperlink_count < citation_count - 5:
        pass  # -3 正文中超过5处引用编号点击后不能跳转到参考文献对应条目 已按用户要求删除
    if ref_nums and ref_nums[0] != 1:
        pass  # -1 参考文献列表编号不是从1开始 已按用户要求删除
    if ref_nums and sorted(ref_nums) != list(range(min(ref_nums), max(ref_nums) + 1)):
        pass  # -1 参考文献列表编号出现跳号 已按用户要求删除
    duplicates = [n for n, c in Counter(ref_nums).items() if c > 1]
    if duplicates:
        pass  # -1 参考文献列表编号出现重号 已按用户要求删除
    if bad_cn_punc > 10:
        pass  # -1 中文正文中出现超过10处半角逗号、半角句号、半角冒号 已按用户要求删除
    if bad_en_punc > 5:
        pass  # -1 英文摘要中出现超过5处中文全角逗号、全角句号 已按用户要求删除
    if bad_stat > 0:
        pass  # -1 统计量中出现P＜0．01这类全角符号 已按用户要求删除
    bad_unit = len(re.findall(r"\d+[ｍＭ]|\d+[．。]\d+[％]", text))
    if bad_unit > 0:
        pass  # -1 数字单位中出现500ｍ或78．3％这类全角数字符号 已按用户要求删除
    if FIGURE_TITLES[0] not in text:
        pass  # -1 图1研究技术路线图不可被删除 已按用户要求删除
    foreword_idx = next((i for i, t in enumerate(block_texts(info)) if t.replace(" ", "").replace("　", "") == "前言"), 0)
    text_after_foreword = "\n".join(block_texts(info)[foreword_idx:])
    if text_after_foreword.count(FIGURE_TITLES[0]) > 1:
        pass  # -1 出现两张相同的图1研究技术路线图 已按用户要求删除
    foreword_idx = next((i for i, t in enumerate(block_texts(info)) if t.replace(" ", "").replace("　", "") == "前言"), 0)
    content_drawings = [d for d in drawings if d["block_index"] >= foreword_idx]
    not_centered = sum(1 for d in content_drawings if d["paragraph"] is not None and para_jc(d["paragraph"]) not in {"center", ""})
    if not_centered:
        pass  # -1 任意1张图片未居中放置 已按用户要求删除
    over_width = sum(1 for d in content_drawings if d["cx"] and d["cx"] > info.printable_width_emu)
    if over_width:
        pass  # -1 任意1张图片宽度超出页面可打印区域 已按用户要求删除
    stretched = 0
    cropped = 0
    for d in content_drawings:
        media = d.get("media")
        intrinsic = info.image_sizes.get(str(media))
        if intrinsic and d["cx"] and d["cy"]:
            shown_ratio = d["cx"] / d["cy"]
            original_ratio = intrinsic[0] / intrinsic[1]
            if abs(shown_ratio / original_ratio - 1) > 0.10:
                stretched += 1
        if d.get("crop"):
            cropped += 1
    if stretched:
        pass  # -1 任意1张图片被明显拉伸变形 已按用户要求删除
    if cropped:
        pass  # -1 任意1张图片被裁切到缺少主要图形内容 已按用户要求删除
        add_hit(hits, -3, "全文所有图片不可出现被裁切到缺少坐标轴、图例、标题、主要图形内容或关键文字", True, f"检测到裁切属性 {cropped} 张")
    overlay = sum(1 for d in content_drawings if d["anchor"] and d["wrap_none"] and d["behind_doc"] != "1")
    if overlay:
        pass  # -1 任意1张图片遮挡正文文字 已按用户要求删除
        pass  # -1 任意1张图片遮挡图题 已按用户要求删除
    # -3 任意一页图片、表格与正文之间出现超过页面高度1/3的无内容空白区域 已按用户要求删除
    # 保留下方 -1 大片空白规则所需的 COM 采集数据。
    page_height_pt = float(info.com.get("page_height_pt") or 0.0)
    para_positions = info.com.get("all_paragraph_positions") or []
    table_blocks_com = info.com.get("table_blocks") or []
    inline_img_blocks = info.com.get("inline_image_blocks") or []
    float_img_blocks = info.com.get("floating_image_blocks") or []
    # ============ 细则 -1 任意1页出现超过页面高度1/3的大片空白 ============
    # 细则拆点（一 一对应细则字面，均基于 Word/WPS 办公软件真实分页/排版）：
    #   点1：判定单位为"任意 1 页"——任意 1 页命中即扣 1 分；
    #   点2：命中对象为"大片空白"——页面正文可用区域内一段连续无内容的垂直区间；
    #        不要求该空白位于"图/表↔正文"之间（与 -3 那条严格版区别）；
    #   点3：阈值为"超过页面高度 1/3"——严格 ">"，等于不算；
    #   点4：判定范围 = 页面正文可用区域（页顶边距 → 页高−页底边距），
    #        细则里说"页面高度 1/3"，页边距外的天然留白不计入"空白"。
    # 数据来源同 -3 空白规则：段落位置 + 表格位置 + 图片位置，均来自 Word COM
    # 采集（doc.Paragraphs / doc.Tables / doc.InlineShapes / doc.Shapes），
    # 与办公软件所见即所得。空白高度阈值仍按细则字面用整页高度 1/3。
    top_margin_pt = float(info.com.get("page_top_margin_pt") or 0.0)
    bot_margin_pt = float(info.com.get("page_bottom_margin_pt") or 0.0)
    huge_blank_pages: List[int] = []
    if page_height_pt > 0 and para_positions:
        threshold_pt = page_height_pt / 3.0  # 点3
        # 版心上下界（点4）：Word COM 段落/表格/图片的 v_pt 均以"页面顶端"为原点。
        area_top = top_margin_pt
        area_bot = page_height_pt - bot_margin_pt
        if area_bot <= area_top:
            area_top, area_bot = 0.0, page_height_pt

        # 汇总每一页的"有内容"垂直区间（正文段/表格/内嵌图片/浮动图形）。
        content_by_page: Dict[int, List[Tuple[float, float]]] = defaultdict(list)
        for pp in para_positions:
            sp, sv = pp.get("page"), pp.get("v_pos_pt")
            ep, ev = pp.get("end_page"), pp.get("end_v_pos_pt")
            if not sp or sv is None or ev is None:
                continue
            if ep == sp:
                content_by_page[int(sp)].append((float(sv), float(ev)))
            else:
                content_by_page[int(sp)].append((float(sv), page_height_pt))
                content_by_page[int(ep)].append((0.0, float(ev)))
        for b in list(table_blocks_com) + list(inline_img_blocks) + list(float_img_blocks):
            tp, tv = b.get("top_page"), b.get("top_v_pt")
            bp, bv = b.get("bot_page"), b.get("bot_v_pt")
            if not tp or tv is None or bv is None:
                continue
            if bp == tp:
                content_by_page[int(tp)].append((float(tv), float(bv)))
            else:
                content_by_page[int(tp)].append((float(tv), page_height_pt))
                content_by_page[int(bp)].append((0.0, float(bv)))

        for page, ivs in content_by_page.items():
            # 裁到版心内 + 合并重叠区间
            clipped: List[Tuple[float, float]] = []
            for a, b in ivs:
                lo = max(min(a, b), area_top)
                hi = min(max(a, b), area_bot)
                if hi > lo:
                    clipped.append((lo, hi))
            if not clipped:
                continue  # 点1：整页无内容不属于"1页出现空白"，跳过
            clipped.sort()
            merged: List[Tuple[float, float]] = [clipped[0]]
            for lo, hi in clipped[1:]:
                mlo, mhi = merged[-1]
                if lo <= mhi:
                    merged[-1] = (mlo, max(mhi, hi))
                else:
                    merged.append((lo, hi))
            # 版心内的空白段：版心顶↔首块、块间、末块↔版心底
            max_gap = 0.0
            cursor = area_top
            for lo, hi in merged:
                max_gap = max(max_gap, lo - cursor)
                cursor = hi
            max_gap = max(max_gap, area_bot - cursor)
            if max_gap > threshold_pt:  # 点2 + 点3
                huge_blank_pages.append(page)

    if huge_blank_pages:
        huge_blank_pages = sorted(set(huge_blank_pages))
        add_hit(
            hits, -1,
            "任意1页出现超过页面高度1/3的大片空白",
            True,
            f"疑似页 {huge_blank_pages[:10]} (共 {len(huge_blank_pages)} 页)"
        )
    if "插图和附表清单" not in text:
        pass  # -3 修改过程中不可出现删除插图和附表清单页 已按用户要求删除
    missing_fig_items = [t for t in FIGURE_TITLES if t not in list_content]
    missing_tbl_items = [t for t in TABLE_TITLES if t not in list_content]
    if missing_fig_items:
        pass  # -1 插图和附表清单中缺少任意1个图目录项 已按用户要求删除
    if missing_tbl_items:
        pass  # -1 插图和附表清单中缺少任意1个表目录项 已按用户要求删除
    ref_section_found = any("参考文献" in t.replace(" ", "").replace("　", "") for t in block_texts(info))
    if not ref_section_found:
        pass  # -3 文中没有出现参考文献章节 已按用户要求删除
    doc_xml = ET.tostring(info.xml.get("word/document.xml", ET.Element("x")), encoding="unicode")
    if any(name in info.package_files for name in ["word/comments.xml", "word/people.xml"]) or "w:ins" in doc_xml or "w:del" in doc_xml or int(info.com.get("comments") or 0) or int(info.com.get("revisions") or 0):
        pass  # -3 不可出现新增封面页、水印、残留批注或修订痕迹 已按用户要求删除
    first_page_text = (info.com.get("page_texts") or [text[:2000]])[0]
    required_cover = ["青岚大学城市发展学院", "硕士研究生学位论文", "基于多源数据的城市社区绿地可达性与居民步行行为研究"]
    missing_cover = [item for item in required_cover if item not in first_page_text]
    if missing_cover:
        pass  # -1 封面页不可缺少基本信息任意一项 已按用户要求删除
    if len(drawings) == 0:
        pass  # -1 封面页顶部的校徽图片不可被删除或被移动到页面下方 已按用户要求删除
    # ============ 细则 -3 目录页前一页的"学位论文原创性承诺"部分文本被移到了封面页后的空白页上 ============
    # 细则拆点（针对 Word/WPS 办公软件真实分页渲染，一 一对应细则字面每一个点）：
    #   点1：文档中存在"学位论文原创性承诺"文本
    #        —— 用 doc.Content.Find + Information(3) 定位其所在页（办公软件真实分页）。
    #   点2：文档中存在"目录"页
    #        —— 作为"目录页前一页"这一原始位置的参照，同样以 Find + Information(3) 定位。
    #   点3：封面页 = 第 1 页；"封面页后的空白页" = 第 2 页
    #        —— 论文首页固定为封面，其后紧邻的第 2 页即细则所指"封面页后的空白页"。
    #   点4：承诺文本"被移到了封面页后的空白页上"
    #        —— 判定 pledge_page == 2（承诺文本出现在封面页之后的第 2 页）。
    #   点5：承诺原本应位于"目录页前一页"
    #        —— 判定 (toc_page - 1) != pledge_page，即承诺不在其原始位置，
    #           从而认定发生了"被移到"的动作；若目录页前一页恰为第 2 页，
    #           则承诺仍在原位、并未被移动，不视为违规。
    # 判定完全基于 Word/WPS 自身分页结果，办公软件所见即所得。
    pledge_page = phrase_pages.get("学位论文原创性承诺")
    toc_page = phrase_pages.get("目录")
    cover_next_blank_page = 2  # 封面(第1页) 之后紧邻的那一页
    pledge_moved_to_blank = (
        pledge_page is not None            # 点1
        and toc_page is not None           # 点2
        and pledge_page == cover_next_blank_page  # 点3 + 点4
        and (toc_page - 1) != pledge_page  # 点5
    )
    if pledge_moved_to_blank:
        add_hit(
            hits, -3,
            "目录页前一页的学位论文原创性承诺文本被移到封面页后的空白页上",
            True,
            f"承诺页={pledge_page}; 目录页={toc_page}; 封面页后空白页={cover_next_blank_page}; 目录前一页={toc_page - 1}"
        )
    if "插图和附表清单" in text and list_content and not all(t in list_content for t in FIGURE_TITLES + TABLE_TITLES):
        pass  # -3 插图和附表清单页下方图目录和表目录字体、行距、对齐方式被改坏 已按用户要求删除
    # -3 全文所有表格单元格文字水平居中或左对齐规则一致，同一张表中不出现随机居中、随机右对齐混用 已按用户要求删除
    pass
    if citation_count and ref_field_count + hyperlink_count and ref_nums and max(ref_nums) < citation_count / 2:
        pass  # -3 不可出现引用编号点击跳转到参考文献中不对应编号条目 已按用户要求删除
    # ============ 细则 -1 从前言页出现任意一页页码缺失、重复或跳号 ============
    # 细则拆点（一 一对应细则字面每一个点；均基于 Word/WPS 办公软件真实分页 + 页脚 PAGE 域渲染）：
    #   点1：判定范围 = "从前言页开始"——
    #        用 doc.Content.Find("前言") + Information(3) 定位"前言"所在页；
    #        自该页起(含)向后所有页参与判定，前言之前页(封面/承诺/目录/摘要 等)不参与。
    #   点2：判定粒度 = "任意一页"——
    #        逐页扫描，只要出现 1 页命中即扣分。
    #   点3：违规形态 A = "页码缺失"——
    #        某页在办公软件中未渲染出页码（该页所属节的页脚既无 PAGE 域(Type=33)
    #        也无数字文本），即 has_pagenum == False。
    #   点4：违规形态 B = "重复"——
    #        某页的 Information(1)=wdActiveEndAdjustedPageNumber(即页脚实际显示值)
    #        与"前言页起"某另一页的显示值相同。
    #   点5：违规形态 C = "跳号"——
    #        自前言页(含)向后，按物理页顺序，相邻两页的显示页码值应满足
    #        next == prev + 1；否则视为跳号（含倒退、缺号）。
    per_page = info.com.get("per_page_pagenum") or []
    foreword_page = phrase_pages.get("前言")  # 点1: 前言所在页(Word 分页真实值)
    pagenum_bad_pages: List[int] = []
    pagenum_reason: List[str] = []
    if per_page and foreword_page:
        scoped = [pp for pp in per_page if pp.get("page") and pp["page"] >= foreword_page]  # 点1
        # 点3: 缺失
        for pp in scoped:
            if not pp.get("has_pagenum") or pp.get("displayed") in (None, 0):
                pagenum_bad_pages.append(int(pp["page"]))
        # 点4: 重复
        seen: Dict[int, int] = {}
        for pp in scoped:
            d = pp.get("displayed")
            if d is None:
                continue
            if d in seen:
                pagenum_bad_pages.append(int(pp["page"]))
                pagenum_reason.append(f"重复:第{pp['page']}页显示={d}(同={seen[d]})")
            else:
                seen[d] = int(pp["page"])
        # 点5: 跳号
        prev = None
        for pp in scoped:
            d = pp.get("displayed")
            if d is None:
                prev = None
                continue
            if prev is not None and d != prev + 1:
                pagenum_bad_pages.append(int(pp["page"]))
                pagenum_reason.append(f"跳号:第{pp['page']}页显示={d}(前一页={prev})")
            prev = d
    if pagenum_bad_pages:
        uniq = sorted(set(pagenum_bad_pages))
        add_hit(
            hits, -1,
            "从前言页出现任意一页页码缺失、重复或跳号",
            True,
            f"前言页={foreword_page}; 命中页={uniq[:10]}(共{len(uniq)}); {'; '.join(pagenum_reason[:3])}"
        )
    if int(info.com.get("header_missing_count") or 0) > 0:
        pass  # -1 从前言页之后出现任意一页页眉缺失 已按用户要求删除
    if int(info.com.get("footer_high_count") or 0) > 0:
        pass  # -1 任意一页的页脚出现在页面三分之一位置 已按用户要求删除


def _locate_docx(dir_path: str) -> Optional[Path]:
    """在给定目录中定位待评估的 .docx 文件。

    约定：批量 runner 传入"脚本所在目录路径"，脚本自己在该目录下找唯一的文档。
    过滤 Office 打开时留下的临时锁文件（以 ~$ 开头）。目录内若存在多个候选，
    优先取文件名含 "024" 的；否则按文件名排序取第一个，避免依赖硬编码文件名。
    """
    base = Path(dir_path)
    if not base.is_dir():
        return None
    candidates = sorted(
        p for p in base.iterdir()
        if p.is_file() and p.suffix.lower() == ".docx" and not p.name.startswith("~$")
    )
    if not candidates:
        return None
    if len(candidates) == 1:
        return candidates[0]
    preferred = [p for p in candidates if "024" in p.stem]
    return preferred[0] if preferred else candidates[0]


def _build_dim2_items(hits: List[Hit]) -> List[dict]:
    """把内部 Hit 列表转换为统一约定的 dim2_items：命中与未命中都保留。
    正向细则和扣分细则均使用原始 score 作为 max_delta；扣分项为负数，
    delta 命中时取原始 score（可能为负），未命中时为 0。"""
    items = []
    for h in hits:
        hit = bool(h.passed)
        items.append({
            "rule": h.rule,
            "max_delta": h.score,
            "delta": h.score if hit else 0,
            "hit": hit,
            "detail": "",
        })
    return items


def evaluate(dir_path: str) -> dict:
    """统一评估入口：传入脚本所在目录路径，脚本自行在该目录中定位并评估 .docx。

    返回结构见《脚本接口差异与统一建议.md》§2.2。
    发生异常时返回 status="error"，不抛出，不 print 主结果。
    """
    result = {
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
        file_path = _locate_docx(dir_path)
        if file_path is None:
            result["status"] = "error"
            result["error"] = f"未在目录中找到 .docx 文件: {dir_path!r}"
            return result
        result["file_name"] = file_path.name

        info = load_word(file_path)
        dim1_ok, dim1_issues = dimension_one(info)
        result["dim1_pass"] = dim1_ok
        result["dim1_reason"] = "；".join(dim1_issues)
        if not dim1_ok:
            return result

        hits = _evaluate_hits(info)
        dim2_items = _build_dim2_items(hits)
        result["dim2_items"] = dim2_items
        result["total_score"] = sum(item["delta"] for item in dim2_items)
        result["max_score"] = sum(
            item["max_delta"] for item in dim2_items if item["max_delta"] > 0
        )
        return result
    except Exception as exc:
        result["status"] = "error"
        result["error"] = f"{type(exc).__name__}: {exc}"
        result["dim2_items"] = []
        result["total_score"] = 0
        return result


if __name__ == "__main__":
    # 仅用于本地调试：从命令行读取目录路径并打印 JSON 结果；
    # 未传参时默认使用脚本自身所在目录。evaluate() 本身不做路径假设。
    _dir = sys.argv[1] if len(sys.argv) > 1 else os.path.dirname(os.path.abspath(__file__))
    _output = json.dumps(evaluate(_dir), ensure_ascii=False, indent=2) + "\n"
    if hasattr(sys.stdout, "buffer"):
        sys.stdout.buffer.write(_output.encode("utf-8", errors="replace"))
    else:
        print(_output)
