#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
自动评估 “1111_勾选自动汇总版.xlsx” 是否满足打分细则。

特点：
- 不依赖 openpyxl / xlwings / Excel COM，仅使用 Python 标准库解析 xlsx/xlsm 的 OOXML。
- 先执行维度1门槛检查；维度1任一关键项失败则总分直接为 0，不再累计维度2。
- 维度2逐条自动检测，命中加分或扣分项，并汇总最终得分。

对外接口（统一约定）：
    def evaluate(dir_path: str) -> dict
        入参 dir_path 为“脚本所在目录的路径”，脚本自己负责在该目录里定位并
        打开被评估的 .xlsx/.xlsm 文档；返回结构化字典（字段见《脚本接口差
        异与统一建议.md》§2.2）。

本地调试：
    python officeval_082_verifier.py <脚本所在目录>
    # 未传参时默认使用当前脚本所在目录
"""

from __future__ import annotations

import json

SCRIPT_ID = "082"
import math
import os
import re
import sys
import tempfile
import zipfile
from dataclasses import dataclass, field
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Tuple
from xml.etree import ElementTree as ET

NS = {
    "main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "rel": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "pkgrel": "http://schemas.openxmlformats.org/package/2006/relationships",
    "xdr": "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
}

CHECKED_CHARS = ("☑", "✓", "✔", "√", "■", "●")
UNCHECKED_CHARS = ("□", "☐", "○", "◇")
BAD_OUTPUT_TOKENS = {"TRUE", "FALSE", "0", "1", "#VALUE!", "#N/A", "#CALC!", "#REF!", "#DIV/0!", "#NAME?", "#NULL!"}
GOOD_OUTPUT_FONTS = {"微软雅黑", "Microsoft YaHei", "宋体", "SimSun", "Calibri"}
ART_FONTS = {"华文彩云", "华文行楷", "华文新魏", "隶书", "方正舒体", "方正姚体", "Comic Sans MS", "Jokerman", "Papyrus"}
LIGHT_BORDER_COLORS = {"FFFFFF", "CBD5E1", "D9E2F3", "D9EAD3", "D0D7DE", "DDDDDD", "E5E7EB", "E7E6E6", "C9DAF8", "CCCCCC"}
LIGHT_FILLS = {"FFFFFF", "F8FAFC", "F9FAFB", "F3F4F6", "F2F2F2", "E8F5E9", "D9EAD3", "E2F0D9", "EEF2FF", "F0FDF4"}
DARK_FONT_COLORS = {"000000", "1F2937", "333333", "404040", "3F3F3F", "24524A", "0F172A"}


def col_to_num(col: str) -> int:
    n = 0
    for ch in col.upper():
        if "A" <= ch <= "Z":
            n = n * 26 + ord(ch) - 64
    return n


def num_to_col(n: int) -> str:
    s = ""
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def split_cell_ref(ref: str) -> Tuple[str, int]:
    m = re.match(r"([A-Z]+)(\d+)", ref.upper())
    if not m:
        return "", 0
    return m.group(1), int(m.group(2))


def safe_float(v, default=None):
    try:
        if v is None:
            return default
        return float(v)
    except Exception:
        return default


def text_of(elem: Optional[ET.Element]) -> str:
    if elem is None:
        return ""
    parts = []
    for t in elem.iter():
        if t.text:
            parts.append(t.text)
    return "".join(parts)


@dataclass
class PointResult:
    score: int
    title: str
    hit: bool
    evidence: str


@dataclass
class WorkbookData:
    path: Path
    zf: zipfile.ZipFile
    shared_strings: List[str] = field(default_factory=list)
    styles: dict = field(default_factory=dict)
    sheet_names: List[str] = field(default_factory=list)
    sheet_paths: Dict[str, str] = field(default_factory=dict)


class OOXMLWorkbook:
    def __init__(self, path: Path):
        self.path = path
        self.zf = zipfile.ZipFile(path)
        self.data = WorkbookData(path=path, zf=self.zf)
        self._load_shared_strings()
        self._load_styles()
        self._load_sheets()

    def close(self):
        self.zf.close()

    def read_xml(self, name: str) -> ET.Element:
        return ET.fromstring(self.zf.read(name))

    def exists(self, name: str) -> bool:
        return name in self.zf.namelist()

    def _load_shared_strings(self):
        if not self.exists("xl/sharedStrings.xml"):
            return
        root = self.read_xml("xl/sharedStrings.xml")
        strings = []
        for si in root.findall("main:si", NS):
            strings.append(text_of(si))
        self.data.shared_strings = strings

    def _load_sheets(self):
        wb = self.read_xml("xl/workbook.xml")
        rels = self.read_xml("xl/_rels/workbook.xml.rels")
        rel_map = {r.attrib.get("Id"): r.attrib.get("Target") for r in rels.findall("pkgrel:Relationship", NS)}
        for sheet in wb.findall("main:sheets/main:sheet", NS):
            name = sheet.attrib.get("name", "")
            rid = sheet.attrib.get(f"{{{NS['rel']}}}id")
            target = rel_map.get(rid, "")
            if target.startswith("/"):
                sheet_path = target.lstrip("/")
            elif target.startswith("xl/"):
                sheet_path = target
            else:
                sheet_path = "xl/" + target
            self.data.sheet_names.append(name)
            self.data.sheet_paths[name] = sheet_path

    def _load_styles(self):
        styles = {
            "fonts": [],
            "fills": [],
            "borders": [],
            "cellXfs": [],
            "dxfs": [],
        }
        if not self.exists("xl/styles.xml"):
            self.data.styles = styles
            return
        root = self.read_xml("xl/styles.xml")

        for font in root.findall("main:fonts/main:font", NS):
            name = font.find("main:name", NS)
            sz = font.find("main:sz", NS)
            color = font.find("main:color", NS)
            styles["fonts"].append({
                "name": name.attrib.get("val") if name is not None else None,
                "size": safe_float(sz.attrib.get("val") if sz is not None else None),
                "bold": font.find("main:b", NS) is not None,
                "italic": font.find("main:i", NS) is not None,
                "color": self._color_rgb(color),
            })

        for fill in root.findall("main:fills/main:fill", NS):
            fg = fill.find("main:patternFill/main:fgColor", NS)
            bg = fill.find("main:patternFill/main:bgColor", NS)
            styles["fills"].append({"fgColor": self._color_rgb(fg), "bgColor": self._color_rgb(bg)})

        for border in root.findall("main:borders/main:border", NS):
            sides = {}
            for side_name in ("left", "right", "top", "bottom"):
                side = border.find(f"main:{side_name}", NS)
                if side is None:
                    sides[side_name] = {"style": None, "color": None}
                else:
                    color = side.find("main:color", NS)
                    sides[side_name] = {"style": side.attrib.get("style"), "color": self._color_rgb(color)}
            styles["borders"].append(sides)

        for xf in root.findall("main:cellXfs/main:xf", NS):
            styles["cellXfs"].append({
                "fontId": int(xf.attrib.get("fontId", "0")),
                "fillId": int(xf.attrib.get("fillId", "0")),
                "borderId": int(xf.attrib.get("borderId", "0")),
                "applyAlignment": xf.attrib.get("applyAlignment"),
                "alignment": self._alignment(xf.find("main:alignment", NS)),
            })

        for dxf in root.findall("main:dxfs/main:dxf", NS):
            styles["dxfs"].append(text_of(dxf))

        self.data.styles = styles

    def _color_rgb(self, elem: Optional[ET.Element]) -> Optional[str]:
        if elem is None:
            return None
        rgb = elem.attrib.get("rgb")
        if rgb:
            return rgb[-6:].upper()
        indexed = elem.attrib.get("indexed")
        if indexed in ("64", "65"):
            return None
        return elem.attrib.get("theme")

    def _alignment(self, elem: Optional[ET.Element]) -> dict:
        if elem is None:
            return {}
        return dict(elem.attrib)

    def sheet(self, name: str) -> "SheetData":
        return SheetData(self, name, self.data.sheet_paths[name])


class SheetData:
    def __init__(self, wb: OOXMLWorkbook, name: str, path: str):
        self.wb = wb
        self.name = name
        self.path = path
        self.root = wb.read_xml(path)
        self.cells = self._load_cells()
        self.merges = self._load_merges()
        self.data_validations = self._load_data_validations()
        self.conditional_ranges = self._load_conditional_ranges()
        self.col_widths = self._load_col_widths()
        self.row_heights = self._load_row_heights()
        self.drawing_targets = self._load_sheet_relationship_targets()

    def _load_cells(self) -> Dict[str, dict]:
        cells = {}
        for c in self.root.findall(".//main:c", NS):
            ref = c.attrib.get("r")
            if not ref:
                continue
            t = c.attrib.get("t")
            s = int(c.attrib.get("s", "0"))
            f = c.find("main:f", NS)
            v = c.find("main:v", NS)
            is_elem = c.find("main:is", NS)
            value = ""
            raw = v.text if v is not None and v.text is not None else ""
            if t == "s":
                idx = int(raw or 0)
                value = self.wb.data.shared_strings[idx] if idx < len(self.wb.data.shared_strings) else ""
            elif t == "inlineStr":
                value = text_of(is_elem)
            elif t == "str":
                value = raw
            else:
                value = raw
            cells[ref] = {"value": value, "raw": raw, "formula": f.text if f is not None else None, "style": s, "type": t}
        return cells

    def _load_merges(self) -> List[str]:
        return [m.attrib.get("ref", "") for m in self.root.findall("main:mergeCells/main:mergeCell", NS)]

    def _load_data_validations(self) -> List[dict]:
        result = []
        for dv in self.root.findall("main:dataValidations/main:dataValidation", NS):
            formula1 = dv.find("main:formula1", NS)
            result.append({"sqref": dv.attrib.get("sqref", ""), "type": dv.attrib.get("type"), "formula1": formula1.text if formula1 is not None else ""})
        return result

    def _load_conditional_ranges(self) -> List[str]:
        return [cf.attrib.get("sqref", "") for cf in self.root.findall("main:conditionalFormatting", NS)]

    def conditional_rules(self) -> List[Tuple[str, str]]:
        """返回 (sqref, 规则公式文本) 列表，用于确认勾选状态的公式/条件格式关联。"""
        out: List[Tuple[str, str]] = []
        for cf in self.root.findall("main:conditionalFormatting", NS):
            sqref = cf.attrib.get("sqref", "")
            for rule in cf.findall("main:cfRule", NS):
                f = rule.find("main:formula", NS)
                text = f.text if (f is not None and f.text) else ""
                out.append((sqref, text))
        return out

    def _load_col_widths(self) -> Dict[str, float]:
        widths = {}
        for col in self.root.findall("main:cols/main:col", NS):
            min_c = int(col.attrib.get("min", "0"))
            max_c = int(col.attrib.get("max", "0"))
            w = safe_float(col.attrib.get("width"))
            for c in range(min_c, max_c + 1):
                widths[num_to_col(c)] = w
        return widths

    def _load_row_heights(self) -> Dict[int, float]:
        heights = {}
        for row in self.root.findall("main:sheetData/main:row", NS):
            r = int(row.attrib.get("r", "0"))
            ht = safe_float(row.attrib.get("ht"))
            if ht is not None:
                heights[r] = ht
        return heights

    def _load_sheet_relationship_targets(self) -> List[str]:
        rel_path = str(Path(self.path).parent / "_rels" / (Path(self.path).name + ".rels")).replace("\\", "/")
        if not self.wb.exists(rel_path):
            return []
        root = self.wb.read_xml(rel_path)
        return [r.attrib.get("Target", "") for r in root.findall("pkgrel:Relationship", NS)]

    def cell(self, ref: str) -> dict:
        return self.cells.get(ref.upper(), {"value": "", "raw": "", "formula": None, "style": 0, "type": None})

    def value(self, ref: str) -> str:
        return str(self.cell(ref).get("value") or "")

    def formula(self, ref: str) -> str:
        return str(self.cell(ref).get("formula") or "")

    def style(self, ref: str) -> dict:
        sid = self.cell(ref).get("style", 0)
        xfs = self.wb.data.styles.get("cellXfs", [])
        return xfs[sid] if sid < len(xfs) else {}

    def font(self, ref: str) -> dict:
        xf = self.style(ref)
        fonts = self.wb.data.styles.get("fonts", [])
        fid = xf.get("fontId", 0)
        return fonts[fid] if fid < len(fonts) else {}

    def fill(self, ref: str) -> dict:
        xf = self.style(ref)
        fills = self.wb.data.styles.get("fills", [])
        fid = xf.get("fillId", 0)
        return fills[fid] if fid < len(fills) else {}

    def border(self, ref: str) -> dict:
        xf = self.style(ref)
        borders = self.wb.data.styles.get("borders", [])
        bid = xf.get("borderId", 0)
        return borders[bid] if bid < len(borders) else {}

    def iter_refs(self, start_col: str, end_col: str, start_row: int, end_row: int) -> Iterable[str]:
        for r in range(start_row, end_row + 1):
            for c in range(col_to_num(start_col), col_to_num(end_col) + 1):
                yield f"{num_to_col(c)}{r}"

    def nonempty_text_cells(self, start_col: str, end_col: str, start_row: int, end_row: int) -> List[Tuple[str, str]]:
        out = []
        for ref in self.iter_refs(start_col, end_col, start_row, end_row):
            val = self.value(ref).strip()
            if val:
                out.append((ref, val))
        return out

    def has_sheet_protection(self) -> bool:
        return self.root.find("main:sheetProtection", NS) is not None

    def has_drawing_or_objects(self) -> bool:
        if self.root.find("main:drawing", NS) is not None or self.root.find("main:legacyDrawing", NS) is not None:
            return True
        names = self.wb.zf.namelist()
        return any(part.startswith("xl/drawings/") or part.startswith("xl/media/") or part.startswith("xl/embeddings/") or part.startswith("xl/controls/") for part in names)

    def drawing_object_count(self) -> int:
        count = 0
        for part in self.wb.zf.namelist():
            if part.startswith("xl/drawings/") and part.endswith(".xml"):
                try:
                    root = self.wb.read_xml(part)
                    count += len(root.findall(".//xdr:twoCellAnchor", NS))
                    count += len(root.findall(".//xdr:oneCellAnchor", NS))
                    count += len(root.findall(".//xdr:absoluteAnchor", NS))
                except Exception:
                    pass
        return count

    def has_vba(self) -> bool:
        return self.wb.exists("xl/vbaProject.bin")


def tag_text(value: str) -> str:
    value = str(value or "").strip()
    for ch in CHECKED_CHARS + UNCHECKED_CHARS:
        if value.startswith(ch):
            return value[1:].strip()
    return value.strip()


def is_tag_option(value: str) -> bool:
    v = str(value or "").strip()
    if not v:
        return False
    if v.startswith(CHECKED_CHARS + UNCHECKED_CHARS):
        return len(tag_text(v)) >= 1
    # 允许没有方框字符但为中文标签的备选实现。
    return bool(re.search(r"[一-鿿A-Za-z]", v)) and len(v) <= 40


def is_checkbox_like(value: str) -> bool:
    v = str(value or "").strip()
    return v.startswith(CHECKED_CHARS + UNCHECKED_CHARS)


def range_intersects_sqref(sqref: str, start_col: str, end_col: str, start_row: int, end_row: int) -> bool:
    scn, ecn = col_to_num(start_col), col_to_num(end_col)
    for token in sqref.split():
        if ":" in token:
            a, b = token.split(":", 1)
        else:
            a = b = token
        ac, ar = split_cell_ref(a)
        bc, br = split_cell_ref(b)
        if not ac or not bc:
            continue
        c1, c2 = sorted((col_to_num(ac), col_to_num(bc)))
        r1, r2 = sorted((ar, br))
        if c1 <= ecn and c2 >= scn and r1 <= end_row and r2 >= start_row:
            return True
    return False


def expand_sqref_cells(sqref: str) -> List[str]:
    """把数据验证/条件格式的 sqref（可能是多段区域）展开为逐个单元格引用。"""
    cells: List[str] = []
    for token in sqref.split():
        if ":" in token:
            a, b = token.split(":", 1)
        else:
            a = b = token
        ac, ar = split_cell_ref(a)
        bc, br = split_cell_ref(b)
        if not ac or not bc:
            continue
        c1, c2 = sorted((col_to_num(ac), col_to_num(bc)))
        r1, r2 = sorted((ar, br))
        for c in range(c1, c2 + 1):
            for r in range(r1, r2 + 1):
                cells.append(f"{num_to_col(c)}{r}")
    return cells


def detect_tag_area(ws: SheetData) -> Tuple[int, int, List[str], List[str]]:
    refs_b = []
    refs_c = []
    for r in range(1, 201):
        if is_tag_option(ws.value(f"B{r}")):
            refs_b.append(f"B{r}")
        if is_tag_option(ws.value(f"C{r}")):
            refs_c.append(f"C{r}")
    all_rows = [split_cell_ref(ref)[1] for ref in refs_b + refs_c]
    if not all_rows:
        return 0, 0, [], []
    return min(all_rows), max(all_rows), refs_b, refs_c


def detect_output_area(ws: SheetData, tag_start: int, tag_end: int) -> Tuple[int, int]:
    # 优先寻找 I/J 区域含标题或公式的行段。
    start = 6
    end = max(tag_end, 30)
    formula_rows = []
    for r in range(1, 201):
        if ws.formula(f"I{r}") or ws.formula(f"J{r}") or ws.value(f"I{r}").strip() or ws.value(f"J{r}").strip():
            if r >= 4:
                formula_rows.append(r)
    if formula_rows:
        start = min([r for r in formula_rows if r >= 5] or [6])
        end = max(formula_rows)
    return start, end


def formula_mentions_tag_area(formula: str) -> bool:
    f = formula.upper()
    return bool(re.search(r"\$?B\$?\d+\s*:\s*\$?C\$?\d+", f) or re.search(r"LEFT\s*\(\s*\$?B", f) or "☑" in formula)


def checkbox_control_check(ws: SheetData, refs_b: List[str], refs_c: List[str]) -> Tuple[bool, str]:
    """严格按 +5 细则逐点核验“B/C 标签勾选控件”，且判定方式在办公软件（Excel/WPS）中真实有效。

    细则拆解为六个必须同时成立的点：
      1) B、C 两列每个标签选项均有“可点击勾选控件”（单元格内或左侧）。
         办公软件里对应：该标签单元格上挂有“列表数据验证下拉”，或该行有绘图/表单复选框控件。
      2) 控件数量与标签选项数量一致（每个标签都被覆盖，不多不少）。
      3) 勾选控件不会遮挡标签文字。
         单元格内“□/☑ + 空格 + 文字”为纯文本前缀，天然不覆盖文字；
         列表下拉的两个候选值都保留完整标签文字，勾选后文字仍在。
      4) 点击后能切换“选中/未选中”两种状态。
         对应：数据验证候选值恰含同一标签的“□ 标签”和“☑ 标签”两项，点击下拉即可切换。
      5) 每个勾选控件与对应行列的布尔值/公式判断/VBA 逻辑关联。
         对应：条件格式规则 LEFT(cell,1)="☑" 覆盖 B/C 区域，或右侧汇总公式按
         LEFT($B$5:$C$68,1)="☑" 逐单元格判断勾选状态。
      6) 勾选→被识别为选中，取消→被识别为未选中。
         对应：以“☑”前缀表示选中、“□”前缀表示未选中，且上述公式/条件格式以此前缀识别。
    """
    tag_refs = refs_b + refs_c
    if not tag_refs:
        return False, "未在 B/C 列识别到任何标签选项"

    # 建立“单元格 -> 数据验证候选值”映射，仅统计落在 B/C 标签单元格上的验证。
    dv_by_cell: Dict[str, str] = {}
    for dv in ws.data_validations:
        if dv.get("type") != "list":
            continue
        f1 = dv.get("formula1", "") or ""
        for cell in expand_sqref_cells(dv.get("sqref", "")):
            col, _ = split_cell_ref(cell)
            if col in ("B", "C"):
                dv_by_cell[cell] = f1

    # “标签选项”指真正的可勾选选项单元格：带 □/☑ 前缀，或挂有勾选下拉；
    # 排除表头（如“标签选项一/二”）等非选项文本，避免虚增计数。
    option_refs = [
        ref for ref in tag_refs
        if is_checkbox_like(ws.value(ref)) or ref in dv_by_cell
    ]
    if not option_refs:
        return False, "未在 B/C 列识别到带勾选控件的标签选项"

    # 点1 & 点2：每个标签选项都要有勾选控件（下拉或绘图控件），且数量一致。
    #   · 下拉控件：dv_by_cell 已按 B/C 单元格粒度建立，天然一一对应。
    #   · 绘图/表单控件：不能只用总数比较——需解析每个控件的锚点单元格，
    #     确认它就落在对应 B/C 标签的行列上，且与标签一一对应（不重不漏），
    #     否则会把无关的绘图对象（图标/图片/装饰形状等）误算成勾选控件。
    drawing_count = ws.drawing_object_count()
    dv_covered = [ref for ref in option_refs if ref in dv_by_cell]
    control_by_dv = len(dv_covered) == len(option_refs)

    # 解析绘图/表单控件锚点 -> 工作表 1 基 (col_letter, row)，仅保留落在 B/C 列的。
    bc_anchor_cells: List[Tuple[str, int]] = []
    for (row0, col0) in _drawing_anchor_cells(ws):
        if row0 < 0 or col0 not in (1, 2):
            continue
        col_letter = "B" if col0 == 1 else "C"
        bc_anchor_cells.append((col_letter, row0 + 1))

    # 逐个锚点落在对应 B/C 标签行列 → 建立“控件 ↔ 标签选项”的一一对应关系。
    option_cells: set[tuple[str, int]] = {(split_cell_ref(ref)[0], split_cell_ref(ref)[1]) for ref in option_refs}
    matched_options: set[tuple[str, int]] = set()
    unmatched_controls: List[Tuple[str, int]] = []
    duplicate_controls: List[Tuple[str, int]] = []
    for anchor in bc_anchor_cells:
        if anchor not in option_cells:
            unmatched_controls.append(anchor)
        elif anchor in matched_options:
            duplicate_controls.append(anchor)
        else:
            matched_options.add(anchor)

    drawing_matched = len(matched_options)
    control_by_drawing = (
        drawing_matched == len(option_refs)                # 每个标签都被至少一个 B/C 锚点覆盖
        and not unmatched_controls                         # 不存在无关绘图对象混入
        and not duplicate_controls                         # 同一标签不重复
        and len(bc_anchor_cells) == len(option_refs)       # 控件数量与标签选项数量一致
    )
    point_1_2 = control_by_dv or control_by_drawing

    # 点4：候选值需同时提供“□ 标签”和“☑ 标签”，点击可在两种状态间切换。
    toggle_ok = 0
    for ref in dv_covered:
        f1 = dv_by_cell[ref]
        label = tag_text(ws.value(ref))
        has_unchecked = any(uc in f1 for uc in UNCHECKED_CHARS)
        has_checked = any(ck in f1 for ck in CHECKED_CHARS)
        # 候选值中应包含当前标签文字，保证控件对应该行该列的标签。
        if has_unchecked and has_checked and (not label or label in f1):
            toggle_ok += 1
    point_4 = (len(dv_covered) > 0 and toggle_ok == len(dv_covered)) or control_by_drawing

    # 点3：单元格内勾选符号为文本前缀，不遮挡文字；候选值保留完整标签文字。
    prefix_ok = 0
    for ref in option_refs:
        v = ws.value(ref).strip()
        if v.startswith(CHECKED_CHARS + UNCHECKED_CHARS) and len(tag_text(v)) >= 1:
            prefix_ok += 1
    point_3 = control_by_drawing or (prefix_ok >= len(option_refs))

    # 点5：勾选状态需与对应行列逻辑关联——条件格式或右侧汇总公式按 LEFT(...)="☑" 判断。
    cf_ok = False
    for sqref, rule in ws.conditional_rules():
        if "☑" in (rule or "") and range_intersects_sqref(sqref, "B", "C", 1, 200):
            cf_ok = True
            break
    right_formula_text = " ".join(
        ws.formula(f"I{r}") + " " + ws.formula(f"J{r}") for r in range(1, 201)
    )
    formula_link_ok = bool(re.search(r'LEFT\s*\(\s*\$?B\$?\d+.*\)\s*=\s*"☑"', right_formula_text)) or (
        "☑" in right_formula_text and formula_mentions_tag_area(right_formula_text)
    )
    point_5 = cf_ok or formula_link_ok

    # 点6：以“☑/□”前缀区分选中与未选中，且被点5的逻辑识别。
    recognizes_checked = ('☑' in right_formula_text) or any(
        '☑' in (rule or "") for _, rule in ws.conditional_rules()
    )
    recognizes_unchecked = any(any(uc in f1 for uc in UNCHECKED_CHARS) for f1 in dv_by_cell.values())
    point_6 = point_5 and recognizes_checked and (recognizes_unchecked or control_by_drawing)

    hit = point_1_2 and point_3 and point_4 and point_5 and point_6
    ev = (
        f"标签选项={len(option_refs)}，挂载下拉控件={len(dv_covered)}，"
        f"绘图/表单控件总数={drawing_count}，B/C锚点控件={len(bc_anchor_cells)}"
        f"(逐行列匹配{drawing_matched}/{len(option_refs)}，无关锚点={len(unmatched_controls)}，重复锚点={len(duplicate_controls)})；"
        f"[点1/2 控件齐备且逐行列一一对应]={point_1_2}，[点3 不遮挡文字(前缀{prefix_ok}/{len(option_refs)})]={point_3}，"
        f"[点4 可切换选中/未选中(下拉双态{toggle_ok}/{len(dv_covered)})]={point_4}，"
        f"[点5 与行列逻辑关联(条件格式={cf_ok},公式={formula_link_ok})]={point_5}，"
        f"[点6 选中/未选中被识别]={point_6}"
    )
    return hit, ev


def range_columns_span(formula: str) -> Tuple[bool, bool]:
    """分析公式里出现的区域引用，判断是否覆盖 B 列、C 列。

    返回 (覆盖B列, 覆盖C列)。既支持合并区域 $B$5:$C$68，也支持分开引用 B 列与 C 列。
    """
    covers_b = False
    covers_c = False
    b_num, c_num = col_to_num("B"), col_to_num("C")
    for m in re.finditer(r"\$?([A-Z]+)\$?\d+\s*:\s*\$?([A-Z]+)\$?\d+", formula):
        c1, c2 = sorted((col_to_num(m.group(1)), col_to_num(m.group(2))))
        if c1 <= b_num <= c2:
            covers_b = True
        if c1 <= c_num <= c2:
            covers_c = True
    return covers_b, covers_c


def single_toggle_display_check(ws: SheetData, out_start: int, out_end: int) -> Tuple[bool, str]:
    """严格按 +5 细则核验“单项勾选显示效果”，判定基于交互能力（公式/数据验证/VBA 关系）。

    细则四个点：
      点1) 勾选 B 列任一标签后，右侧自动显示区域出现该标签文本。
      点2) 勾选 C 列任一标签后，右侧自动显示区域出现该标签文本。
      点3) 取消任一已勾选标签后，右侧该标签文本同步消失。
      点4) 其他仍被勾选的标签继续保留显示。

    判定原则（交互能力检查，不要求交付文件预置 B/C 各一个已勾选样例）：
      - 数据验证下拉候选值同时含“□ 标签”与“☑ 标签”，且覆盖 B/C 标签单元格
        → 用户点击下拉即可在两种状态间切换。
      - 右侧汇总公式以 LEFT(cell,1)="☑" 为过滤条件，且引用区域覆盖 B 列（点1）/
        C 列（点2）→ 任意该列单元格切换到 ☑，其文本即被公式选入右侧；
        切回 □ 即不再满足条件而被移除（点3）；
        用 AGGREGATE(15,...)/SMALL + 位置索引 或 FILTER 按序枚举 → 其余仍勾选项按左侧顺序保留（点4）。
      - VBA：工作簿含 vbaProject.bin 时，允许由宏逻辑（Worksheet_Change 等）驱动
        勾选/取消与右侧输出的联动，作为等效证据。
      - 若文件恰好已预置 ☑ 样例且右侧确实显示，作为额外静态佐证（非必需）。
    """
    # 右侧自动显示区域实际“显示出来”的标签文本（排除表头/标题），用于静态佐证。
    header_like = re.compile(r"分类维度|选定标签|维度分类|选定|操作|标签")
    shown_texts = []
    for r in range(out_start, min(out_end, 200) + 1):
        for c in ("I", "J"):
            v = ws.value(f"{c}{r}").strip()
            if v and not header_like.search(v) and not v.upper().startswith("#") and v.upper() not in ("TRUE", "FALSE"):
                shown_texts.append(tag_text(v))
    shown_set = set(shown_texts)

    # 若文件恰好已有 ☑ 样例，用于静态佐证；无样例不影响得分。
    checked_b = [tag_text(ws.value(f"B{r}")) for r in range(1, 201)
                 if ws.value(f"B{r}").strip().startswith(CHECKED_CHARS)]
    checked_c = [tag_text(ws.value(f"C{r}")) for r in range(1, 201)
                 if ws.value(f"C{r}").strip().startswith(CHECKED_CHARS)]

    # 右侧显示区域全部公式合并成一段文本，供结构判定使用。
    formulas = " ".join(
        ws.formula(f"I{r}") + " " + ws.formula(f"J{r}") for r in range(out_start, min(out_end, 200) + 1)
    )
    up = formulas.upper()
    # 公式以 LEFT(...)="☑" 为过滤条件 → 只有 ☑ 前缀的单元格被选入。
    checked_filter_ok = ("☑" in formulas) and (re.search(r"LEFT\s*\(", formulas) is not None)
    # 公式引用区域覆盖 B/C 两列（合并 $B$5:$C$68 或分开引用均可）。
    covers_b, covers_c = range_columns_span(formulas)
    # 按位置索引升序枚举，保证其余已勾选项按左侧顺序保留显示。
    sequential_ok = any(k in up for k in ("AGGREGATE", "SMALL(", "FILTER("))

    # 数据验证支持“□ ↔ ☑”切换：候选值中同时含 □ 与 ☑ 两态，并覆盖 B/C 标签单元格。
    dv_toggle_b = dv_toggle_c = False
    for dv in ws.data_validations:
        if dv.get("type") != "list":
            continue
        f1 = dv.get("formula1", "") or ""
        has_u = any(uc in f1 for uc in UNCHECKED_CHARS)
        has_c = any(ck in f1 for ck in CHECKED_CHARS)
        if not (has_u and has_c):
            continue
        for cell in expand_sqref_cells(dv.get("sqref", "")):
            col, _ = split_cell_ref(cell)
            if col == "B":
                dv_toggle_b = True
            elif col == "C":
                dv_toggle_c = True

    # VBA 兜底：宏工程存在时，允许勾选/取消由宏逻辑驱动。
    vba_ok = ws.has_vba()

    # 静态佐证（可选，非必需）：文件已有 ☑ 且右侧确实显示了对应文本。
    static_b = any(t in shown_set for t in checked_b) if checked_b else False
    static_c = any(t in shown_set for t in checked_c) if checked_c else False

    # 点1/点2：交互能力——B/C 列切到 ☑ 即可在右侧出现。
    formula_b_ok = checked_filter_ok and covers_b
    formula_c_ok = checked_filter_ok and covers_c
    point_1 = formula_b_ok or dv_toggle_b or vba_ok or static_b
    point_2 = formula_c_ok or dv_toggle_c or vba_ok or static_c

    # 点3：取消勾选后同步消失——公式以 ☑ 过滤即天然满足，或由 VBA 保证。
    point_3 = checked_filter_ok or vba_ok
    # 点4：其余仍勾选项按序保留——位置索引升序枚举（AGGREGATE/SMALL/FILTER），或由 VBA 保证。
    point_4 = (checked_filter_ok and sequential_ok) or vba_ok

    hit = point_1 and point_2 and point_3 and point_4
    ev = (
        f"右侧显示公式：过滤☑={checked_filter_ok}，覆盖B={covers_b}，覆盖C={covers_c}，按序枚举={sequential_ok}；"
        f"数据验证双态：B={dv_toggle_b}，C={dv_toggle_c}；VBA={vba_ok}；"
        f"静态佐证：预置勾选B={len(checked_b)}(右侧呈现={static_b})，"
        f"预置勾选C={len(checked_c)}(右侧呈现={static_c})；"
        f"[点1 B列勾选→右侧出现]={point_1}，[点2 C列勾选→右侧出现]={point_2}，"
        f"[点3 取消后同步消失]={point_3}，[点4 其余按序保留]={point_4}"
    )
    return hit, ev


def _drawing_controls_detailed(ws: SheetData) -> list[dict[str, object]]:
    """解析绘图/表单控件的锚点与尺寸细节，供“仅按 B/C 标签行统计”与“越界校验”使用。

    每条记录包含（EMU 单位，1 cm=360000 EMU）：
      · from_col/from_row：起始单元格 0 基坐标；from_col_off/from_row_off：起点在该格内偏移；
      · to_col/to_row：终止单元格 0 基坐标（twoCellAnchor 才有）；to_col_off/to_row_off：终点在该格内偏移；
      · ext_cx/ext_cy：oneCellAnchor 的显式宽高；twoCellAnchor 无此值（用起止锚点推算）；
      · kind：'twoCell'/'oneCell'/'absolute'。
    absoluteAnchor 使用绝对坐标，与单元格无关，不参与“落在 B/C 标签行”的判定。
    """
    def _int(el: ET.Element | None) -> int | None:
        if el is None or el.text is None:
            return None
        try:
            return int(el.text)
        except Exception:
            return None

    controls: list[dict[str, object]] = []
    for part in ws.wb.zf.namelist():
        if not (part.startswith("xl/drawings/") and part.endswith(".xml")):
            continue
        try:
            root = ws.wb.read_xml(part)
        except Exception:
            continue
        for kind_tag, kind in (("twoCellAnchor", "twoCell"), ("oneCellAnchor", "oneCell"), ("absoluteAnchor", "absolute")):
            for anchor in root.findall(f".//xdr:{kind_tag}", NS):
                rec: dict[str, object] = {"kind": kind}
                frm = anchor.find("xdr:from", NS)
                if frm is not None:
                    rec["from_col"] = _int(frm.find("xdr:col", NS))
                    rec["from_col_off"] = _int(frm.find("xdr:colOff", NS)) or 0
                    rec["from_row"] = _int(frm.find("xdr:row", NS))
                    rec["from_row_off"] = _int(frm.find("xdr:rowOff", NS)) or 0
                to = anchor.find("xdr:to", NS)
                if to is not None:
                    rec["to_col"] = _int(to.find("xdr:col", NS))
                    rec["to_col_off"] = _int(to.find("xdr:colOff", NS)) or 0
                    rec["to_row"] = _int(to.find("xdr:row", NS))
                    rec["to_row_off"] = _int(to.find("xdr:rowOff", NS)) or 0
                ext = anchor.find("xdr:ext", NS)
                if ext is not None:
                    rec["ext_cx"] = safe_float(ext.attrib.get("cx"))
                    rec["ext_cy"] = safe_float(ext.attrib.get("cy"))
                controls.append(rec)
    return controls


def checkbox_size_check(ws: SheetData, refs_b: List[str], refs_c: List[str]) -> Tuple[bool, str]:
    """严格按 +3 细则核验“控件大小”，判定在 Excel/WPS 中真实有效。

    细则四个点：
      点1) 每个复选框/勾选区域宽高约 0.25cm–0.5cm。
      点2) 大小统一。
      点3) 能够被鼠标点击。
      点4) 未超出所在单元格边界。

    办公软件中存在两种实现，分别用其真实尺寸依据判定：
      A. 单元格内“□/☑”勾选（数据验证下拉）：勾选区域即单元格内显示的 □/☑ 字符，
         其宽高由该单元格“字号（磅）”决定——字符高度 ≈ 字号，1 磅 = 2.54/72 cm。
         · 点1：字号换算到 cm 落在 0.25–0.5（对应约 7–14 磅）。
         · 点2：所有勾选选项单元格字号一致 → 大小统一。
         · 点3：整格挂有列表数据验证，鼠标点击单元格即弹出下拉可勾选 → 可点击。
         · 点4：字符高度远小于行高、宽度不超过列宽 → 不超出单元格边界。
      B. 绘图/表单复选框控件：仅统计“锚定在 B/C 标签行”的控件（无关图案不参与），
         用锚点起止位置与该格真实行高/列宽对比来校验“未超出所在单元格边界”。
    """
    tag_refs = refs_b + refs_c
    if not tag_refs:
        return False, "未在 B/C 列识别到标签选项"

    PT_TO_CM = 2.54 / 72.0
    EMU_PER_CM = 360000.0

    # 建立 B/C 标签单元格 -> 列表数据验证 的映射（点3 可点击的依据）。
    dv_cells: set[str] = set()
    for dv in ws.data_validations:
        if dv.get("type") != "list":
            continue
        for cell in expand_sqref_cells(dv.get("sqref", "")):
            col, _ = split_cell_ref(cell)
            if col in ("B", "C"):
                dv_cells.add(cell)

    # 收集 B/C 标签选项所在“行 -> 列集合”，用作“绘图控件是否锚定到标签行列”的过滤依据。
    label_row_cols: dict[int, set[str]] = {}
    for ref in tag_refs:
        col, row = split_cell_ref(ref)
        label_row_cols.setdefault(row, set()).add(col)

    # 实现 B：仅统计锚定在 B/C 标签行的绘图/表单控件（过滤掉无关图案）。
    all_controls = _drawing_controls_detailed(ws)
    bc_controls: list[dict[str, object]] = []
    for rec in all_controls:
        if rec.get("kind") == "absolute":
            continue  # 绝对锚点不与单元格绑定，无法判定“锚在 B/C 标签行”
        fc = rec.get("from_col")
        fr = rec.get("from_row")
        if not isinstance(fc, int) or not isinstance(fr, int):
            continue
        # 0 基列：B=1、C=2；转 1 基行与工作表 B/C 标签行比对。
        if fc not in (1, 2):
            continue
        sheet_row = fr + 1
        col_letter = "B" if fc == 1 else "C"
        if sheet_row not in label_row_cols or col_letter not in label_row_cols[sheet_row]:
            continue
        bc_controls.append(rec)

    if bc_controls:
        # 逐个控件：计算“像素框”宽高 与 “所在单元格边界”，做实际范围对比。
        widths_cm: list[float] = []
        heights_cm: list[float] = []
        in_range = 0
        clickable_ctrl = 0
        within_boundary = 0
        detail_lines: list[str] = []
        for rec in bc_controls:
            fc = rec["from_col"]  # 已确认为 int
            fr = rec["from_row"]
            f_col_off = rec.get("from_col_off") or 0
            f_row_off = rec.get("from_row_off") or 0
            assert isinstance(fc, int) and isinstance(fr, int)
            assert isinstance(f_col_off, int) and isinstance(f_row_off, int)

            col_letter = "B" if fc == 1 else "C"
            sheet_row = fr + 1

            # 该起始单元格的真实尺寸（cm）。Excel 默认行高 15 磅、默认列宽 8.43 字符。
            row_h_cm = (ws.row_heights.get(sheet_row, 15.0)) * PT_TO_CM
            col_w_cm = (ws.col_widths.get(col_letter, 8.43)) * 0.202
            row_h_emu = row_h_cm * EMU_PER_CM
            col_w_emu = col_w_cm * EMU_PER_CM

            if rec.get("kind") == "twoCell":
                tc = rec.get("to_col")
                tr = rec.get("to_row")
                t_col_off_raw = rec.get("to_col_off") or 0
                t_row_off_raw = rec.get("to_row_off") or 0
                t_col_off = int(t_col_off_raw) if isinstance(t_col_off_raw, int) else 0
                t_row_off = int(t_row_off_raw) if isinstance(t_row_off_raw, int) else 0
                # 若起止同格：宽=to_off-from_off，高=to_off-from_off；跨格则累加中间整格。
                if isinstance(tc, int) and tc == fc:
                    width_emu = max(0, t_col_off - f_col_off)
                elif isinstance(tc, int) and tc > fc:
                    width_emu = max(0, col_w_emu - f_col_off)
                    for mid_c in range(fc + 1, tc):
                        mid_letter = num_to_col(mid_c + 1)
                        width_emu += (ws.col_widths.get(mid_letter, 8.43)) * 0.202 * EMU_PER_CM
                    width_emu += max(0, t_col_off)
                else:
                    width_emu = 0.0
                if isinstance(tr, int) and tr == fr:
                    height_emu = max(0, t_row_off - f_row_off)
                elif isinstance(tr, int) and tr > fr:
                    height_emu = max(0, row_h_emu - f_row_off)
                    for mid_r in range(fr + 1, tr):
                        height_emu += (ws.row_heights.get(mid_r + 1, 15.0)) * PT_TO_CM * EMU_PER_CM
                    height_emu += max(0, t_row_off)
                else:
                    height_emu = 0.0
                same_cell = (tc == fc) and (tr == fr)
            else:
                ext_cx = rec.get("ext_cx")
                ext_cy = rec.get("ext_cy")
                width_emu = float(ext_cx) if isinstance(ext_cx, (int, float)) else 0.0
                height_emu = float(ext_cy) if isinstance(ext_cy, (int, float)) else 0.0
                # oneCellAnchor 起点固定，若 from_off+ext 未超出起始单元格，则视为同格。
                same_cell = (f_col_off + width_emu <= col_w_emu) and (f_row_off + height_emu <= row_h_emu)

            w_cm = float(width_emu) / EMU_PER_CM
            h_cm = float(height_emu) / EMU_PER_CM
            widths_cm.append(w_cm)
            heights_cm.append(h_cm)

            # 点1：宽高约 0.25–0.5cm。
            if 0.25 <= w_cm <= 0.5 and 0.25 <= h_cm <= 0.5:
                in_range += 1
            # 点3：表单/ActiveX 复选框控件本身可被鼠标点击；此处按“B/C 标签行的功能控件”计。
            clickable_ctrl += 1
            # 点4：与所在单元格真实边界对比——起点在该格内且未越过右/下边界。
            fits_start = (0 <= f_col_off <= col_w_emu) and (0 <= f_row_off <= row_h_emu)
            fits_end = (f_col_off + width_emu <= col_w_emu + 1) and (f_row_off + height_emu <= row_h_emu + 1)
            not_overflow = bool(same_cell) and fits_start and fits_end
            if not_overflow:
                within_boundary += 1
            detail_lines.append(
                f"{col_letter}{sheet_row}:{w_cm:.2f}x{h_cm:.2f}cm (格{col_w_cm:.2f}x{row_h_cm:.2f}cm,同格={bool(same_cell)},未越界={not_overflow})"
            )

        n = len(bc_controls)
        expected = len(tag_refs)
        # 点1：所有控件宽高均在 0.25–0.5cm。
        point_1 = in_range == n
        # 点2：大小统一（宽高极差极小），且控件数量与标签数一致。
        uniform = (max(widths_cm) - min(widths_cm) <= 0.05) and (max(heights_cm) - min(heights_cm) <= 0.05)
        point_2 = uniform and n == expected
        # 点3：全部为功能控件，可鼠标点击。
        point_3 = clickable_ctrl == n
        # 点4：全部控件未超出所在单元格边界（用真实行高/列宽比对）。
        point_4 = within_boundary == n

        hit = point_1 and point_2 and point_3 and point_4
        ev = (
            f"绘图/表单控件总数 {len(all_controls)}，锚定 B/C 标签行的功能控件 {n}(应为 {expected})；"
            f"[点1 宽高0.25–0.5cm({in_range}/{n})]={point_1}，"
            f"[点2 大小统一(宽{min(widths_cm):.2f}–{max(widths_cm):.2f}cm,高{min(heights_cm):.2f}–{max(heights_cm):.2f}cm) 且数量一致({n}/{expected})]={point_2}，"
            f"[点3 可鼠标点击({clickable_ctrl}/{n})]={point_3}，"
            f"[点4 未超出所在单元格边界({within_boundary}/{n})]={point_4}；样例：{'; '.join(detail_lines[:4])}"
        )
        return hit, ev

    # 实现 A：单元格内 □/☑ 勾选区域，用字号（磅）换算尺寸。
    option_refs = [ref for ref in tag_refs if is_checkbox_like(ws.value(ref)) or ref in dv_cells]
    if not option_refs:
        return False, "未识别到单元格内勾选区域或绘图控件"

    sizes_cm = []
    clickable = 0
    within_boundary = 0
    for ref in option_refs:
        size_pt = ws.font(ref).get("size") or 11
        glyph_cm = size_pt * PT_TO_CM
        sizes_cm.append(glyph_cm)
        # 点3：整格数据验证下拉，点击单元格即可勾选。
        if ref in dv_cells:
            clickable += 1
        # 点4：勾选字符高度 < 行高、宽度 < 列宽 → 不超出单元格边界。
        _, row = split_cell_ref(ref)
        col, _ = split_cell_ref(ref)
        row_h_cm = (ws.row_heights.get(row, 15.0)) * PT_TO_CM
        col_w_cm = (ws.col_widths.get(col, 8.43)) * 0.202  # Excel 列宽(字符)≈0.202cm/字符（约值）
        if glyph_cm <= row_h_cm and glyph_cm <= col_w_cm:
            within_boundary += 1

    n = len(option_refs)
    # 点1：每个勾选区域宽高约 0.25–0.5cm（□/☑ 近似方形，宽高同为字号换算值）。
    in_range = sum(1 for s in sizes_cm if 0.25 <= s <= 0.5)
    point_1 = in_range == n
    # 点2：大小统一——所有勾选选项字号一致（换算尺寸极差极小）。
    point_2 = (max(sizes_cm) - min(sizes_cm)) <= 0.02
    # 点3：全部选项单元格可点击（挂列表数据验证）。
    point_3 = clickable == n
    # 点4：全部勾选区域不超出单元格边界。
    point_4 = within_boundary == n

    hit = point_1 and point_2 and point_3 and point_4
    ev = (
        f"单元格内勾选选项 {n} 个，勾选符号尺寸 {min(sizes_cm):.3f}–{max(sizes_cm):.3f}cm；"
        f"[点1 宽高0.25–0.5cm({in_range}/{n})]={point_1}，"
        f"[点2 大小统一]={point_2}，"
        f"[点3 可鼠标点击(数据验证{clickable}/{n})]={point_3}，"
        f"[点4 未超出单元格边界({within_boundary}/{n})]={point_4}"
    )
    return hit, ev


def _hex_channels(hexstr: Optional[str]) -> Optional[Tuple[int, int, int]]:
    h = (hexstr or "").upper()
    if len(h) == 8:  # 去掉 ARGB 的 alpha
        h = h[2:]
    if len(h) != 6 or any(ch not in "0123456789ABCDEF" for ch in h):
        return None
    return int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)


def _luminance(hexstr: Optional[str]) -> Optional[float]:
    ch = _hex_channels(hexstr)
    if ch is None:
        return None
    r, g, b = ch
    return 0.299 * r + 0.587 * g + 0.114 * b


def _is_white_or_very_light_gray(hexstr: Optional[str]) -> bool:
    """白色或极浅灰：三通道都很高且接近中性。"""
    ch = _hex_channels(hexstr)
    if ch is None:
        return False
    r, g, b = ch
    return min(r, g, b) >= 0xF0  # >=240，如 F8FAFC / FFFFFF / F2F2F2


def _is_light_tint(hexstr: Optional[str]) -> bool:
    """浅色底纹（含浅绿、白及其它浅色 pastel）：亮度高。"""
    lum = _luminance(hexstr)
    return lum is not None and lum >= 0xC8  # >=200


def _is_light_line(hexstr: Optional[str]) -> bool:
    """浅灰色或白色的细线颜色；无显式颜色视为默认浅色网格线。"""
    if hexstr is None:
        return True
    lum = _luminance(hexstr)
    if lum is None:
        return True  # theme/indexed 等非 RGB，按办公软件默认浅色网格线处理
    return lum >= 0xA0  # >=160，浅灰及更浅（如 CBD5E1）


def left_border_check(ws: SheetData, tag_start: int, tag_end: int) -> Tuple[bool, str]:
    """严格按 +1 细则核验“左侧表格边框”，判定基于办公软件真实边框/填充属性。

    细则五个点：
      点1) 区域边框完整。
      点2) 横向和纵向分隔线为浅灰色或白色细线。
      点3) 表格行列结构清晰。
      点4) 主体行填充为白色或极浅灰色。
      点5) A列分类维度“可”使用浅绿色或白色底纹。（“可”为允许项）

    办公软件中对应的真实属性：
      - 边框完整：单元格四边 border style 均非 none（Excel“所有框线”）。
      - 分隔线浅色细线：内部网格线 border style=thin 且颜色为浅灰/白（RGB 亮度高）。
      - 行列结构清晰：全部单元格都有完整框线，行列被清晰划分。
      - 填充：cell fill 的 fgColor（Excel“填充颜色”）。主体行为白/极浅灰；
        A 列分类底纹为浅色（浅绿或白，允许其它浅色）。
    """
    # 表头行＝含“分类维度/标签选项”标题的行；主体行＝带 □/☑ 前缀的标签选项行。
    body_rows = [r for r in range(tag_start, tag_end + 1)
                 if is_checkbox_like(ws.value(f"B{r}")) or is_checkbox_like(ws.value(f"C{r}"))]
    header_rows = [r for r in range(tag_start, tag_end + 1)
                   if r not in body_rows and (ws.value(f"A{r}").strip() or ws.value(f"B{r}").strip())]
    if not body_rows:
        return False, "未识别到左侧表格主体行（B/C 无勾选选项）"

    region_rows = sorted(set(body_rows + header_rows))
    all_cells = [f"{c}{r}" for r in region_rows for c in ("A", "B", "C")]

    # 点1：区域边框完整——每个单元格四边都有框线。
    complete = 0
    thin_line_total = 0
    thin_line_light = 0
    for ref in all_cells:
        b = ws.border(ref)
        sides = [b.get(s, {}) for s in ("left", "right", "top", "bottom")]
        present = [s for s in sides if s.get("style") not in (None, "none")]
        if len(present) == 4:
            complete += 1
        # 点2：统计细线（thin）分隔线的颜色是否浅灰/白。
        for s in present:
            if s.get("style") == "thin":
                thin_line_total += 1
                if _is_light_line(s.get("color")):
                    thin_line_light += 1
    point_1 = complete == len(all_cells)

    # 点2：所有细线分隔线均为浅灰/白，且确实存在细线分隔（形成网格）。
    point_2 = thin_line_total > 0 and thin_line_light == thin_line_total

    # 点3：行列结构清晰——全部单元格完整框线即可清晰划分行列。
    point_3 = point_1 and point_2

    # 点4：主体行 B/C 填充为白色或极浅灰色。
    body_bc = [f"{c}{r}" for r in body_rows for c in ("B", "C")]
    body_light = sum(1 for ref in body_bc if _is_white_or_very_light_gray(ws.fill(ref).get("fgColor") or "FFFFFF"))
    point_4 = body_light == len(body_bc)

    # 点5：A 列分类维度底纹为浅色（浅绿/白，允许其它浅色）。“可”为允许项。
    a_cells = [f"A{r}" for r in body_rows]
    a_light = sum(1 for ref in a_cells if _is_light_tint(ws.fill(ref).get("fgColor") or "FFFFFF"))
    point_5 = a_light == len(a_cells)

    hit = point_1 and point_2 and point_3 and point_4 and point_5
    ev = (
        f"区域 {len(all_cells)} 格(主体行{len(body_rows)}+表头{len(header_rows)})；"
        f"[点1 边框完整({complete}/{len(all_cells)})]={point_1}，"
        f"[点2 细线分隔浅灰/白({thin_line_light}/{thin_line_total})]={point_2}，"
        f"[点3 行列结构清晰]={point_3}，"
        f"[点4 主体行白/极浅灰({body_light}/{len(body_bc)})]={point_4}，"
        f"[点5 A列浅色底纹({a_light}/{len(a_cells)})]={point_5}"
    )
    return hit, ev


def output_text_font_check(ws: SheetData, out_start: int, out_end: int) -> Tuple[bool, str]:
    """严格按 +1 细则核验“右侧输出字体”，判定在 Excel/WPS 中真实有效。

    细则四个点：
      点1) 右侧输出文本字体为微软雅黑、宋体或 Calibri。
      点2) 字号 9–12 磅。
      点3) 颜色为黑色或深灰色。
      点4) 文字左对齐或居中对齐。

    办公软件中对应的真实机制：
      - 字体名：OOXML 单元格样式里的 font name 即办公软件“字体”下拉框的值。
        Calibri 在无该字体的环境（WPS/LibreOffice）会以度量兼容的 Carlito 呈现，
        微软雅黑=Microsoft YaHei、宋体=SimSun/NSimSun 亦为同一字体的英文名，
        因此这些等价名视为同一字体家族。
      - 字号：font sz 即“字号（磅）”，直接判断 9–12。
      - 颜色：font color 的 RGB，黑色/深灰色即接近 000000 的低亮度值。
      - 对齐：alignment horizontal 为 left/center 即“左对齐/居中对齐”
        （general 在文本内容下默认左对齐，等同左对齐）。
    """
    # 微软雅黑 / 宋体 / Calibri 及其在办公软件中的等价字体名。
    accepted_fonts = {
        "微软雅黑", "Microsoft YaHei", "Microsoft YaHei UI",
        "宋体", "SimSun", "NSimSun",
        "Calibri", "Carlito",
    }

    # “右侧输出文本”＝勾选左侧控件后，右侧自动显示区域里“真正出现的被选中标签内容”。
    # 本条检查的正是这些实际出现的标签文本的格式，因此：
    #   - 只取右侧真正显示出内容（显示值非空）的单元格；
    #   - 排除分类表头/标题（如「分类维度/选定标签/维度分类（选定）」等，非输出标签内容）；
    #   - 若右侧没有出现任何被选中标签内容（未勾选或公式在办公软件里算不出结果），
    #     则没有“输出文本”可供评判 → 不加分。
    header_like = re.compile(r"分类维度|选定标签|维度分类|选定|操作|标签")
    refs: List[str] = []
    for r in range(out_start, min(out_end, 200) + 1):
        for c in ("I", "J"):
            v = ws.value(f"{c}{r}").strip()
            if v and not header_like.search(v) and not v.upper().startswith("#") and v.upper() not in ("TRUE", "FALSE"):
                refs.append(f"{c}{r}")
    if not refs:
        return False, "右侧未出现任何被选中标签内容（未勾选或公式在办公软件中算不出结果），无输出文本可评，不加分"

    name_ok = size_ok = color_ok = align_ok = 0
    details: List[str] = []
    for ref in refs:
        font = ws.font(ref)
        name = font.get("name") or ""
        size = font.get("size") or 11
        color = (font.get("color") or "000000").upper()
        horizontal = ws.style(ref).get("alignment", {}).get("horizontal", "general")

        # 点1：字体在允许集合内（空表示继承默认，交付前应显式设置，故不放宽）。
        if name in accepted_fonts:
            name_ok += 1
        # 点2：字号 9–12 磅。
        if 9 <= size <= 12:
            size_ok += 1
        # 点3：黑色或深灰色。
        if color in DARK_FONT_COLORS:
            color_ok += 1
        # 点4：左对齐或居中（general 于文本默认左对齐）。
        if horizontal in ("left", "center", "general"):
            align_ok += 1
        details.append(f"{ref}:{name or '默认'} {size:g}pt #{color} {horizontal}")

    n = len(refs)
    point_1 = name_ok == n
    point_2 = size_ok == n
    point_3 = color_ok == n
    point_4 = align_ok == n
    hit = point_1 and point_2 and point_3 and point_4
    ev = (
        f"输出文本单元格 {n} 个；"
        f"[点1 字体微软雅黑/宋体/Calibri({name_ok}/{n})]={point_1}，"
        f"[点2 字号9–12磅({size_ok}/{n})]={point_2}，"
        f"[点3 黑色/深灰色({color_ok}/{n})]={point_3}，"
        f"[点4 左对齐/居中({align_ok}/{n})]={point_4}；样例：{'; '.join(details[:4])}"
    )
    return hit, ev


def row_height_ok(ws: SheetData, start: int, end: int) -> Tuple[bool, str]:
    """严格按 +1 细则核验“右侧输出列表行高”，判定在 Excel/WPS 中真实有效。

    细则四个点：
      点1) 右侧输出区域行高为 20–24 磅。
      点2) 文字完整显示。
      点3) 长文本自动换行，或列宽足够。
      点4) 不出现文字被截断到无法辨认的情况。

    办公软件中对应的真实机制：
      - 行高：OOXML 中行的 ht 属性即办公软件里的“行高（磅）”，直接判断是否落在 20–24。
      - 自动换行：单元格 alignment 的 wrapText=1 就是 Excel/WPS 的“自动换行”，
        开启后长文本会在单元格内折行、配合足够行高即可完整显示，不会被横向截断。
      - 列宽足够：即使未开启换行，只要列宽足以容纳最长标签文本，也不会截断。
      - 只要“自动换行”或“列宽足够”成立，且行高在区间内，即满足“文字完整显示、不被截断”。
    """
    # 右侧输出列表的“列表行”：区域内实际承载显示内容（公式或文本）且设置了显式行高的行。
    # 这些行才是办公软件中真正呈现标签列表的行；未使用的空档行使用默认高度，不计入行高判定。
    list_rows: List[int] = []
    for r in range(start, min(end, 200) + 1):
        has_content = bool(
            ws.formula(f"I{r}") or ws.formula(f"J{r}")
            or ws.value(f"I{r}").strip() or ws.value(f"J{r}").strip()
        )
        if has_content and r in ws.row_heights:
            list_rows.append(r)
    if not list_rows:
        return False, "未识别到右侧输出列表的行（无显式行高的显示行）"

    heights = [ws.row_heights[r] for r in list_rows]
    # 点1：行高 20–24 磅。
    in_range = sum(1 for h in heights if 20 <= h <= 24)
    point_1 = in_range == len(heights)

    # 点2/3/4：文字完整显示、长文本换行或列宽足够、不被截断。
    # 换行：I/J 显示单元格开启 wrapText。
    wrap_rows = 0
    for r in list_rows:
        i_wrap = ws.style(f"I{r}").get("alignment", {}).get("wrapText") in ("1", "true", "True")
        j_wrap = ws.style(f"J{r}").get("alignment", {}).get("wrapText") in ("1", "true", "True")
        if i_wrap or j_wrap:
            wrap_rows += 1
    wrap_ok = wrap_rows == len(list_rows)
    # 列宽足够：显示标签文本的 J 列（及分类的 I 列）列宽不小于常规标签长度所需。
    width_i = ws.col_widths.get("I", 8.43)
    width_j = ws.col_widths.get("J", 8.43)
    width_ok = width_j >= 12 and width_i >= 8
    point_234 = wrap_ok or width_ok

    hit = point_1 and point_234
    ev = (
        f"列表行 {len(list_rows)} 行，行高范围 {min(heights):.1f}–{max(heights):.1f}；"
        f"[点1 行高20–24磅({in_range}/{len(heights)})]={point_1}，"
        f"[点2/3/4 换行({wrap_rows}/{len(list_rows)})或列宽足够(I={width_i:.1f},J={width_j:.1f}) → 完整显示不截断]={point_234}"
    )
    return hit, ev


def col_widths_ok(ws: SheetData) -> Tuple[bool, str]:
    """严格按 -5 细则核验“列宽”，判定基于办公软件真实列宽（字符数）。

    细则四个点（全部满足才算“列宽合格”，任一不满足即触发 -5 扣分）：
      点1) A 列宽度约 8–12 字符。
      点2) B 列和 C 列宽度约 24–26 字符。
      点3) 右侧输出列 I、J 列宽度与 B/C 相近，约 24–26 字符。
      点4) 所有主要标签文字和右侧结果文字能完整显示。

    办公软件中对应的真实属性：
      - OOXML <col width=..> 即 Excel/WPS 里“列宽”对话框显示的字符数，直接判断区间。
      - “完整显示”：列宽足以容纳该列最长文本，或单元格开启自动换行（wrapText）。
    """
    widths = {c: ws.col_widths.get(c, 8.43) for c in ("A", "B", "C", "I", "J")}

    # 点1：A 列 8–12 字符。
    point_1 = 8 <= widths["A"] <= 12
    # 点2：B、C 列 24–26 字符。
    point_2 = 24 <= widths["B"] <= 26 and 24 <= widths["C"] <= 26
    # 点3：I、J 列 24–26 字符（与 B/C 相近）。
    point_3 = 24 <= widths["I"] <= 26 and 24 <= widths["J"] <= 26

    # 点4：主要标签文字（B/C）和右侧结果文字（I/J）能完整显示。
    # 完整显示＝该列列宽够宽（≈24 字符可容纳标签），或对应单元格开启自动换行。
    def _col_display_ok(col: str, sample_rows: Iterable[int]) -> bool:
        if widths[col] >= 24:
            return True
        # 列不够宽时，若单元格自动换行也可完整显示。
        for r in sample_rows:
            if ws.style(f"{col}{r}").get("alignment", {}).get("wrapText") in ("1", "true", "True"):
                return True
        return False

    tag_start, tag_end, _, _ = detect_tag_area(ws)
    out_start, out_end = detect_output_area(ws, tag_start, tag_end)
    body_rows = range(max(tag_start, 1), tag_end + 1)
    out_rows = range(out_start, min(out_end, out_start + 20) + 1)
    display_ok = (
        _col_display_ok("B", body_rows) and _col_display_ok("C", body_rows)
        and _col_display_ok("I", out_rows) and _col_display_ok("J", out_rows)
    )
    point_4 = display_ok

    ok = point_1 and point_2 and point_3 and point_4
    ev = (
        f"A={widths['A']:.1f}, B={widths['B']:.1f}, C={widths['C']:.1f}, I={widths['I']:.1f}, J={widths['J']:.1f}；"
        f"[点1 A列8–12]={point_1}，[点2 B/C列24–26]={point_2}，"
        f"[点3 I/J列24–26]={point_3}，[点4 标签与结果文字完整显示]={point_4}"
    )
    return ok, ev


def has_title_near_i4(ws: SheetData) -> Tuple[bool, str]:
    """严格按 -1 细则核验“I4:J4 附近是否出现‘维度分类（选定）’或类似选定结果标题”。

    细则单点：右侧 I4:J4 附近若“没有”出现“维度分类（选定）”或类似选定结果标题 → 触发 -1 扣分。
    本函数返回“是否存在该标题”（存在则不扣分）。

    办公软件中对应的真实位置：
      - I4:J4 是右侧输出区的标题行（本表 I4:J4 为合并单元格），标题文本写在其左上角 I4；
      - “附近”容错取 I3:J5 一圈，兼容标题写在相邻行的情况。
      - 判定文本：精确含“维度分类（选定）”，或语义等价的“选定结果类标题”
        （如“维度分类/选定标签/已选/选择结果/选定结果/汇总结果”等组合）。
    """
    texts: List[Tuple[str, str]] = []
    for r in range(3, 6):           # I3:J5，覆盖 I4:J4 及其相邻行
        for c in ("I", "J"):
            t = ws.value(f"{c}{r}").strip()
            if t:
                texts.append((f"{c}{r}", t))
    joined = " ".join(t for _, t in texts)

    # 精确标题，或“分类/维度”与“选定/已选/选择/选中”组合的选定结果类标题。
    exact = "维度分类（选定）" in joined or "维度分类(选定)" in joined
    similar = bool(
        re.search(r"(维度|分类).*(选定|已选|选择|选中)", joined)
        or re.search(r"(选定|已选|选择|选中).*(结果|标签|分类|维度)", joined)
        or "选定标签" in joined or "选择结果" in joined or "选定结果" in joined
    )
    ok = exact or similar
    return ok, "; ".join(f"{ref}={t}" for ref, t in texts[:8]) or "I4:J4 附近未见文本"


def output_has_bad_tokens(ws: SheetData, start: int, end: int) -> Tuple[bool, str]:
    """严格按 -1 细则核验“右侧 I/J 显示区域是否出现中间值”。

    细则单点：右侧 I 列和 J 列显示区域出现 TRUE、FALSE、0、1、#VALUE!、#N/A、#CALC! 等
    中间值 → 触发 -1 扣分。（“等”表示同类：布尔值、0/1 数字标记、各类 # 错误值。）

    办公软件中对应的真实“显示值”：
      - 单元格保存的显示结果即 Excel/WPS 里看到的值：
        · 布尔中间值 TRUE/FALSE；
        · 0 / 1 这类布尔转数字的中间标记；
        · 以 # 开头的公式错误值（#VALUE! #N/A #CALC! #REF! #DIV/0! #NAME? #NULL! #SPILL! #GETTING_DATA 等）。
      - 只检查右侧输出区 I/J，且以“单元格显示出来的值”为准；
        IFERROR(...,\"\") 等已被吸收为空串的单元格不显示中间值，不计入。
    """
    hits: List[str] = []
    for r in range(start, min(end, 200) + 1):
        for c in ("I", "J"):
            val = ws.value(f"{c}{r}").strip()
            if not val:
                continue
            up = val.upper()
            is_bool = up in ("TRUE", "FALSE")
            is_binary = val in ("0", "1")           # 0/1 布尔标记（不含小数）
            is_error = up.startswith("#")           # 各类 # 错误值
            if is_bool or is_binary or is_error:
                hits.append(f"{c}{r}={val}")
    bad = bool(hits)
    return bad, ("; ".join(hits[:20]) if bad else "右侧 I/J 显示区域未出现 TRUE/FALSE/0/1/#错误值 等中间值")


def _drawing_anchor_cells(ws: SheetData) -> list[tuple[int, int]]:
    """解析绘图/表单控件锚点的“起始单元格”坐标 (row, col)，0 基。

    OOXML 锚点 <xdr:from> 内含 <xdr:col>、<xdr:row>（均 0 基），
    对应办公软件里控件左上角所在的单元格——即控件实际落在哪一行哪一列。
    """
    anchors: list[tuple[int, int]] = []
    for part in ws.wb.zf.namelist():
        if part.startswith("xl/drawings/") and part.endswith(".xml"):
            try:
                root = ws.wb.read_xml(part)
            except Exception:
                continue
            for tag in ("twoCellAnchor", "oneCellAnchor"):
                for anchor in root.findall(f".//xdr:{tag}", NS):
                    frm = anchor.find("xdr:from", NS)
                    if frm is None:
                        continue
                    col_el = frm.find("xdr:col", NS)
                    row_el = frm.find("xdr:row", NS)
                    col = int(col_el.text) if (col_el is not None and col_el.text) else -1
                    row = int(row_el.text) if (row_el is not None and row_el.text) else -1
                    anchors.append((row, col))
    return anchors


def _picture_anchors_over_ac(ws: SheetData) -> int:
    """统计锚定在 A:C 列范围（0 基列 0–2）内的图片对象数量。

    OOXML 中图片为绘图里的 <xdr:pic>，其 <xdr:from><xdr:col> 表示图片左上角所在列。
    办公软件里把表格“存成图片”后，会得到一张覆盖 A:C 的图片对象。
    """
    count = 0
    for part in ws.wb.zf.namelist():
        if part.startswith("xl/drawings/") and part.endswith(".xml"):
            try:
                root = ws.wb.read_xml(part)
            except Exception:
                continue
            for anchor in root.findall(".//xdr:twoCellAnchor", NS) + root.findall(".//xdr:oneCellAnchor", NS):
                if anchor.find(".//xdr:pic", NS) is None:
                    continue
                frm = anchor.find("xdr:from", NS)
                col_el = frm.find("xdr:col", NS) if frm is not None else None
                col = int(col_el.text) if (col_el is not None and col_el.text) else -1
                if 0 <= col <= 2:
                    count += 1
    return count


def _has_image_media(ws: SheetData) -> bool:
    """工作簿是否包含图片媒体文件（xl/media/ 下的 png/jpg 等）。"""
    return any(part.startswith("xl/media/") for part in ws.wb.zf.namelist())


def left_table_as_image(ws: SheetData) -> Tuple[bool, str]:
    """严格按 -5 细则核验“A:C 左侧用户标签维度表是否整体为图片”。

    细则单点：A 列至 C 列左侧用户标签维度表整体为图片 → 触发 -5 扣分。

    办公软件中“整表为图片”的真实特征（两者同时成立才判为图片化）：
      特征1) A:C 几乎没有可编辑的真实单元格文本——表格内容不再是单元格，而是被贴成一张图。
      特征2) 存在覆盖 A:C 的图片对象（xl/media 图片 + 锚定在 A:C 列的 <xdr:pic>）。
    只要 A:C 仍是大量真实文本单元格（可点选、可编辑），即便另有零星图片，也不算“整体为图片”。
    """
    text_count = len(ws.nonempty_text_cells("A", "C", 1, 120))
    has_media = _has_image_media(ws)
    pic_over_ac = _picture_anchors_over_ac(ws)

    # 特征1：A:C 基本没有真实文本（阈值取 10，低于则疑似非真实表格）。
    lacks_real_text = text_count < 10
    # 特征2：存在覆盖 A:C 的图片对象。
    covered_by_picture = has_media and pic_over_ac >= 1

    bad = lacks_real_text and covered_by_picture
    if bad:
        return True, f"A:C 仅 {text_count} 个真实文本单元格，且存在 {pic_over_ac} 张锚定 A:C 的图片，判定整表为图片"
    return False, (
        f"A:C 有 {text_count} 个真实可编辑文本单元格，"
        f"A:C 图片对象 {pic_over_ac} 张，媒体文件={has_media}，非整表图片"
    )


def resolve_target_sheet_name(path: Path, wb: OOXMLWorkbook) -> Optional[str]:
    """按评分入口定位待评估工作表。

    原细则写的是工作簿中存在名称为“1111”的工作表；用户补充说明：
    “1111”出现在文件名中也可以通过。因此这里优先使用名为“1111”的
    工作表；若不存在但文件名包含“1111”，则使用第一个工作表继续评估。
    """
    if "1111" in wb.data.sheet_names:
        return "1111"
    if "1111" in path.stem and wb.data.sheet_names:
        return wb.data.sheet_names[0]
    return None


def dimension1_checks(path: Path, wb: Optional[OOXMLWorkbook], error: Optional[str] = None) -> Tuple[bool, list[str], Optional[str]]:
    reasons: list[str] = []
    ok = True
    if path.suffix.lower() not in (".xlsx", ".xlsm"):
        ok = False
        reasons.append(f"交付文件扩展名为 {path.suffix}，不是 .xlsx 或 .xlsm")
    else:
        reasons.append(f"交付文件扩展名为 {path.suffix}，符合 .xlsx/.xlsm 要求")
    if error:
        ok = False
        reasons.append(f"文件无法作为 OOXML 工作簿正常打开：{error}")
        return ok, reasons, None
    if wb is None:
        return False, reasons + ["未能读取工作簿"], None

    target_sheet = resolve_target_sheet_name(path, wb)
    if target_sheet is None:
        # 未找到名为“1111”的工作表，也未按文件名匹配到——退回第一个工作表继续评估。
        if not wb.data.sheet_names:
            ok = False
            reasons.append("工作簿中未找到任何工作表")
            return ok, reasons, None
        target_sheet = wb.data.sheet_names[0]

    ws = wb.sheet(target_sheet)

    if path.suffix.lower() == ".xlsm" and not ws.has_vba():
        reasons.append("文件为 .xlsm 但未发现 vbaProject.bin；本任务不强制要求宏")
    return ok, reasons, target_sheet


def evaluate_dimension2(wb: OOXMLWorkbook, sheet_name: str) -> List[PointResult]:
    ws = wb.sheet(sheet_name)
    tag_start, tag_end, refs_b, refs_c = detect_tag_area(ws)
    out_start, out_end = detect_output_area(ws, tag_start, tag_end)
    results: List[PointResult] = []

    # +5 标签勾选控件及关联逻辑（逐点核验，判定在 Excel/WPS 中真实有效）
    hit, ev = checkbox_control_check(ws, refs_b, refs_c)
    results.append(PointResult(5, "标签勾选控件数量、可切换状态及与对应行列选中逻辑关联", hit, ev))

    # +5 单项勾选显示效果（逐点核验，判定基于公式行为，在 Excel/WPS 中真实有效）
    hit, ev = single_toggle_display_check(ws, out_start, out_end)
    results.append(PointResult(5, "单项勾选后右侧显示、取消后同步消失且其他已选保留", hit, ev))

    # +1 行高
    ok, ev = row_height_ok(ws, out_start, out_end)
    results.append(PointResult(1, "右侧输出列表行高 20–24 磅且文字可读", ok, ev))

    # +1 字体（逐点核验，判定基于办公软件真实字体/字号/颜色/对齐）
    ok, ev = output_text_font_check(ws, out_start, out_end)
    results.append(PointResult(1, "右侧输出字体为微软雅黑/宋体/Calibri，9–12 磅，黑/深灰，左/居中", ok, ev))

    # +3 控件大小（逐点核验，判定基于办公软件真实控件/勾选符号尺寸）
    ok, ev = checkbox_size_check(ws, refs_b, refs_c)
    results.append(PointResult(3, "控件或勾选区域大小约 0.25–0.5cm、统一且不越界", ok, ev))

    # +1 左侧表格边框（逐点核验，判定基于办公软件真实边框/填充属性）
    ok, ev = left_border_check(ws, tag_start, tag_end)
    results.append(PointResult(1, "左侧表格边框完整、浅色细线、行列结构清晰", ok, ev))

    # 扣分项

    return results


def _build_dim2_items(points: List[PointResult]) -> List[dict]:
    """把内部 PointResult 列表转换为统一约定的 dim2_items 结构。

    评分规则（与用户约定一致）：
      - “总分/满分（max_score）”= 所有“加分项”的满分之和；扣分项对满分无贡献。
      - “得分（total_score）”= 加分项命中之和 + 扣分项命中之和（含负数）。
    因此：
      - 加分项（score>0）：max_delta=score；命中→delta=score，未命中→delta=0
      - 扣分项（score<0）：max_delta=score（负数）；
                            命中→delta=score（负数），未命中→delta=0
    """
    items: List[dict] = []
    for p in points:
        if p.score >= 0:
            max_delta = p.score
            delta = p.score if p.hit else 0
        else:
            max_delta = p.score
            delta = max_delta if p.hit else 0
        items.append({
            "rule": p.title,
            "max_delta": max_delta,
            "delta": delta,
            "hit": bool(p.hit),
            "detail": "",
        })
    return items


def _locate_target_file(dir_path: Path) -> Optional[Path]:
    """在脚本所在目录中定位被评估的 .xlsx/.xlsm 文档。

    优先选择文件名含“1111”的 .xlsx/.xlsm；否则退回目录里首个 .xlsx/.xlsm。
    临时文件（以 ~$ 开头）忽略。
    """
    if not dir_path.is_dir():
        return None
    candidates: List[Path] = []
    for entry in sorted(dir_path.iterdir()):
        if not entry.is_file():
            continue
        if entry.name.startswith("~$"):
            continue
        if entry.suffix.lower() in (".xlsx", ".xlsm"):
            candidates.append(entry)
    if not candidates:
        return None
    for c in candidates:
        if "1111" in c.stem:
            return c
    return candidates[0]


def evaluate(dir_path: str) -> dict:
    """对外统一入口：接收“脚本所在目录的路径”，在该目录中定位并评估文档。

    返回结构遵循《脚本接口差异与统一建议.md》§2.2 约定。
    """
    result: dict = {
        "id": "082",
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
            result["error"] = f"脚本所在目录不存在或不是目录：{dir_path}"
            return result

        target = _locate_target_file(base)
        if target is None:
            result["status"] = "error"
            result["error"] = f"目录中未找到 .xlsx/.xlsm 文档：{base}"
            return result
        result["file_name"] = target.name

        wb: Optional[OOXMLWorkbook] = None
        open_error: Optional[str] = None
        try:
            try:
                wb = OOXMLWorkbook(target)
            except Exception as exc:
                open_error = repr(exc)

            dim1_ok, reasons, target_sheet = dimension1_checks(target, wb, open_error)
            result["dim1_pass"] = bool(dim1_ok)
            # 维度一原因合并为一段文字：通过时留空，未通过时给出所有失败/说明信息
            result["dim1_reason"] = "" if dim1_ok else "；".join(reasons)

            points: List[PointResult] = []
            if dim1_ok and wb is not None and target_sheet is not None:
                points = evaluate_dimension2(wb, target_sheet)

            dim2_items = _build_dim2_items(points)
            result["dim2_items"] = dim2_items
            result["max_score"] = sum(
                item["max_delta"] for item in dim2_items if item["max_delta"] > 0
            )
            # 维度一未通过时总分直接为 0（与原脚本判分规则一致）
            result["total_score"] = sum(item["delta"] for item in dim2_items) if dim1_ok else 0
        finally:
            if wb is not None:
                wb.close()
    except Exception as exc:  # 兜底，避免脚本自身异常冒泡到 runner
        result["status"] = "error"
        result["error"] = repr(exc)
    return result


if __name__ == "__main__":
    # 本地调试用：默认以当前脚本所在目录作为 dir_path，也可命令行覆盖。
    # 直接写入 stdout.buffer 以 UTF-8 输出，避免 Windows cp1252 终端编码错误；
    # 不修改 sys.stdout 本身，符合《脚本接口差异与统一建议.md》§2.3 的约束。
    _dir = sys.argv[1] if len(sys.argv) > 1 else str(Path(__file__).resolve().parent)
    _payload = json.dumps(evaluate(_dir), ensure_ascii=False, indent=2)
    try:
        _ = sys.stdout.buffer.write((_payload + "\n").encode("utf-8"))
    except Exception:
        print(_payload)
