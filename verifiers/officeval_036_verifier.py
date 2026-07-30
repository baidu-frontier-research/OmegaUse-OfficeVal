#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
自动评估 public_ready_statement_收支记录整理.docx。

评估策略：
1. 先执行“维度1：可用与可修改性”的门槛检查；任一门槛不满足，直接输出 0 分并跳过维度2。
2. 维度1通过后，按维度2的得分点/扣分点累计分数，并打印命中的点、未命中的点和最终得分。

说明：docx 本身不保存 Word 渲染后的真实分页结果。对“第一页/第2-3页数据分布”、
“无严重重叠/裁切”等只能通过 OOXML 中的结构、页面宽度、表格宽度、行属性、重复表头等信息做自动化近似检测。
"""

from __future__ import annotations

import json
import re
import sys
import zipfile
import xml.etree.ElementTree as ET
from dataclasses import dataclass
from decimal import Decimal, InvalidOperation
from pathlib import Path

NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "pic": "http://schemas.openxmlformats.org/drawingml/2006/picture",
    "w14": "http://schemas.microsoft.com/office/word/2010/wordml",
    "v": "urn:schemas-microsoft-com:vml",
    "o": "urn:schemas-microsoft-com:office:office",
    "wps": "http://schemas.microsoft.com/office/word/2010/wordprocessingShape",
    "wpg": "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup",
    "wpi": "http://schemas.microsoft.com/office/word/2010/wordprocessingInk",
}

HEADER = ["Date", "Description", "Withdrawals", "Deposits", "Balance"]
TWIPS_PER_CM = 1440 / 2.54
EMU_PER_CM = 360000

# 作为 PDF 标准答案的自动化判定基准由维度二各评分点单独定义，此处不再维护 48 行完整记录列表。


@dataclass
class CheckResult:
    name: str
    passed: bool
    detail: str = ""


@dataclass
class ScorePoint:
    name: str
    points: int
    hit: bool
    detail: str = ""


@dataclass
class DocxContext:
    path: Path
    zip_file: zipfile.ZipFile
    document: ET.Element
    rels: ET.Element | None
    namelist: list[str]
    styles: ET.Element | None = None


def qn(tag: str) -> str:
    prefix, local = tag.split(":", 1)
    return f"{{{NS[prefix]}}}{local}"


def attr(el: ET.Element | None, name: str, default: str | None = None) -> str | None:
    if el is None:
        return default
    return el.get(qn(name), default)


def children(el: ET.Element | None, path: str) -> list[ET.Element]:
    if el is None:
        return []
    return el.findall(path, NS)


def first(el: ET.Element | None, path: str) -> ET.Element | None:
    if el is None:
        return None
    return el.find(path, NS)


def all_text(el: ET.Element | None) -> str:
    if el is None:
        return ""
    return "".join(t.text or "" for t in el.findall(".//w:t", NS)).strip()


def cell_text(tc: ET.Element) -> str:
    parts: list[str] = []
    for p in tc.findall("./w:p", NS):
        text = "".join(t.text or "" for t in p.findall(".//w:t", NS))
        if text:
            parts.append(text)
    return "\n".join(parts).strip()


def row_texts(tr: ET.Element) -> list[str]:
    return [cell_text(tc) for tc in tr.findall("./w:tc", NS)]


def rows(tbl: ET.Element) -> list[ET.Element]:
    return tbl.findall("./w:tr", NS)


def tables(root: ET.Element) -> list[ET.Element]:
    return root.findall(".//w:tbl", NS)


def normalize_space(text: str) -> str:
    return re.sub(r"\s+", " ", text).strip()


def is_true_word_value(el: ET.Element | None) -> bool:
    if el is None:
        return False
    value = attr(el, "w:val")
    return value is None or value.lower() not in {"0", "false", "off", "no"}


def is_bold(rpr: ET.Element | None) -> bool:
    b = first(rpr, "./w:b") if rpr is not None else None
    return is_true_word_value(b)


def is_regular(rpr: ET.Element | None) -> bool:
    b = first(rpr, "./w:b") if rpr is not None else None
    i = first(rpr, "./w:i") if rpr is not None else None
    return not is_true_word_value(b) and not is_true_word_value(i)


def run_rprs(tc: ET.Element) -> list[ET.Element | None]:
    return [first(r, "./w:rPr") for r in tc.findall(".//w:r", NS) if all_text(r) or r.find(".//w:t", NS) is not None]


def text_run_rprs(tc: ET.Element) -> list[ET.Element | None]:
    return [first(r, "./w:rPr") for r in tc.findall(".//w:r", NS) if "".join(t.text or "" for t in r.findall("./w:t", NS)).strip()]


def font_name_ok(rpr: ET.Element | None, expected: str = "Arial") -> bool:
    fonts = first(rpr, "./w:rFonts") if rpr is not None else None
    if fonts is None:
        return False
    values = [attr(fonts, "w:ascii"), attr(fonts, "w:hAnsi"), attr(fonts, "w:eastAsia")]
    return any(v == expected for v in values) and all(v in {None, expected} for v in values)


def font_size_ok(rpr: ET.Element | None, expected_pt: float, tolerance_pt: float = 0.05) -> bool:
    sz = first(rpr, "./w:sz") if rpr is not None else None
    value = attr(sz, "w:val")
    if value is None:
        return False
    try:
        return abs(int(value) / 2 - expected_pt) <= tolerance_pt
    except ValueError:
        return False


def color_black_ok(rpr: ET.Element | None) -> bool:
    color = first(rpr, "./w:color") if rpr is not None else None
    value = attr(color, "w:val")
    # 未显式设置颜色时，Word 默认通常为黑色；为减少误判，这里接受缺省/auto/黑色。
    return value is None or value.lower() in {"000000", "auto"}


def paragraph_alignment(tc: ET.Element) -> str | None:
    p = first(tc, "./w:p")
    jc = first(first(p, "./w:pPr"), "./w:jc") if p is not None else None
    return attr(jc, "w:val")


def paragraph_format_ok(tc: ET.Element, before: int, after: int, line: int, line_rule: str = "auto") -> bool:
    ps = tc.findall("./w:p", NS)
    # 细则：单元格内不出现额外空段落——每个数据单元格恰好一个段落。
    if len(ps) != 1:
        return False
    # 细则：单元格内不出现多余换行——不含 w:br。
    if tc.findall(".//w:br", NS):
        return False
    spacing = first(first(ps[0], "./w:pPr"), "./w:spacing")
    # 细则：段前0磅、段后10磅、1.15倍行距；
    # OOXML 中段前/段后单位为 twips，10磅=200twips；1.15倍行距=276，lineRule=auto。
    return (
        attr(spacing, "w:before") == str(before)
        and attr(spacing, "w:after") == str(after)
        and attr(spacing, "w:line") == str(line)
        and attr(spacing, "w:lineRule") == line_rule
    )


def row_height(tr: ET.Element) -> tuple[str | None, str | None]:
    h = first(first(tr, "./w:trPr"), "./w:trHeight")
    return attr(h, "w:val"), attr(h, "w:hRule")


def cell_width_twips(tc: ET.Element) -> int | None:
    tcw = first(first(tc, "./w:tcPr"), "./w:tcW")
    value = attr(tcw, "w:w")
    try:
        return int(value) if value is not None else None
    except ValueError:
        return None


def table_grid_widths(tbl: ET.Element) -> list[int]:
    widths: list[int] = []
    for col in children(first(tbl, "./w:tblGrid"), "./w:gridCol"):
        value = attr(col, "w:w")
        try:
            widths.append(int(value))
        except (TypeError, ValueError):
            pass
    return widths


def cm_to_twips(cm: float) -> int:
    return round(cm * TWIPS_PER_CM)


def twips_to_cm(twips: int) -> float:
    return twips / TWIPS_PER_CM


def emu_to_cm(emu: int) -> float:
    return emu / EMU_PER_CM


def money_to_decimal(text: str) -> Decimal | None:
    if not text:
        return None
    if not re.fullmatch(r"\$\d{1,3}(?:,\d{3})*\.\d{2}", text):
        return None
    try:
        return Decimal(text.replace("$", "").replace(",", ""))
    except InvalidOperation:
        return None


def canonical_money(text: str) -> bool:
    value = money_to_decimal(text)
    if value is None:
        return False
    expected = f"${value:,.2f}"
    return text == expected


def load_docx(path: Path) -> tuple[DocxContext | None, str]:
    if path.suffix.lower() != ".docx":
        return None, "交付文件不是 .docx 格式"
    try:
        zf = zipfile.ZipFile(path)
        namelist = zf.namelist()
        document = ET.fromstring(zf.read("word/document.xml"))
        rels = None
        if "word/_rels/document.xml.rels" in namelist:
            rels = ET.fromstring(zf.read("word/_rels/document.xml.rels"))
        styles = None
        if "word/styles.xml" in namelist:
            try:
                styles = ET.fromstring(zf.read("word/styles.xml"))
            except ET.ParseError:
                styles = None
        return DocxContext(path, zf, document, rels, namelist, styles), ""
    except Exception as exc:  # noqa: BLE001 - 这里需要捕获任意坏包/坏 XML。
        return None, f"文件无法作为 Word docx 正常打开/解析：{exc}"


def find_record_table(ctx: DocxContext) -> ET.Element | None:
    for tbl in tables(ctx.document):
        trs = rows(tbl)
        if not trs:
            continue
        if row_texts(trs[0]) == HEADER:
            return tbl
    return None


def data_records(tbl: ET.Element | None) -> list[tuple[str, str, str, str, str]]:
    if tbl is None:
        return []
    out: list[tuple[str, str, str, str, str]] = []
    for tr in rows(tbl)[1:]:
        cells = row_texts(tr)
        if len(cells) == 5:
            out.append(tuple(cells))  # type: ignore[arg-type]
    return out


def dimension1(ctx: DocxContext | None, load_error: str) -> tuple[bool, list[CheckResult], ET.Element | None]:
    checks: list[CheckResult] = []
    if ctx is None:
        checks.append(CheckResult("交付文件为 .docx 格式，文件可正常打开", False, load_error))
        return False, checks, None

    checks.append(CheckResult("交付文件为 .docx 格式，文件可正常打开", True, "docx zip 包与 word/document.xml 可解析"))
    tbl = find_record_table(ctx)
    return all(c.passed for c in checks), checks, tbl


def check_table_format(tbl: ET.Element) -> ScorePoint:
    trs = rows(tbl)
    header = trs[0]
    header_cells = header.findall("./w:tc", NS)
    row_cell_counts = [len(tr.findall("./w:tc", NS)) for tr in trs]
    five_column_table_ok = bool(trs) and all(count == 5 for count in row_cell_counts)
    has_noneditable_objects = bool(tbl.findall(".//w:object", NS) or tbl.findall(".//o:OLEObject", NS) or tbl.findall(".//w:pict", NS) or tbl.findall(".//w:drawing", NS))
    editable_word_table_ok = five_column_table_ok and not has_noneditable_objects and any((t.text or "").strip() for t in tbl.findall(".//w:t", NS))
    text_ok = row_texts(header) == HEADER
    expected_align = ["left", "left", "right", "right", "right"]
    actual_align = [paragraph_alignment(tc) for tc in header_cells]
    align_ok = len(actual_align) == 5 and all(
        actual in {None, "left"} if expected == "left" else actual == "right"
        for actual, expected in zip(actual_align, expected_align)
    )
    font_ok = True
    for tc in header_cells:
        rprs = text_run_rprs(tc)
        font_ok = font_ok and bool(rprs) and all(font_name_ok(rpr) and font_size_ok(rpr, 11.5) and is_bold(rpr) for rpr in rprs)
    hit = editable_word_table_ok and text_ok and align_ok and font_ok
    return ScorePoint(
        "收支记录表格：使用5列可编辑Word表格，列顺序/表头、表头 Arial 11.5磅加粗、对齐",
        3,
        hit,
        f"5列表格={five_column_table_ok}，可编辑Word表格={editable_word_table_ok}，表头={row_texts(header)}，字体={font_ok}，对齐={actual_align}",
    )


def check_column_widths(tbl: ET.Element) -> ScorePoint:
    # 细则：Date/Description/Withdrawals/Deposits/Balance 五列列宽统一 3.68 厘米，各页续表列宽保持一致。
    target = cm_to_twips(3.68)
    col_names = HEADER  # ["Date", "Description", "Withdrawals", "Deposits", "Balance"]

    # 1. tblGrid 五列均为 3.68 cm（定义各列基准宽度，续表列也依此对齐）
    widths = table_grid_widths(tbl)
    grid_ok = len(widths) == 5 and all(abs(w - target) <= 10 for w in widths)

    # 2. 逐行（含表头及所有数据行/续表行）检查每个单元格的实际宽度均为 3.68 cm，
    #    以此确保各页续表列宽保持一致。
    cell_ok = True
    bad_examples: list[str] = []
    for r_idx, tr in enumerate(rows(tbl), start=1):
        cell_widths = [cell_width_twips(tc) for tc in tr.findall("./w:tc", NS)]
        if len(cell_widths) != 5:
            if len(bad_examples) < 3:
                bad_examples.append(f"第{r_idx}行列数={len(cell_widths)}（期望5列）")
            cell_ok = False
            continue
        for col_idx, (w, col_name) in enumerate(zip(cell_widths, col_names), start=1):
            if w is None or abs(w - target) > 10:
                if len(bad_examples) < 3:
                    bad_examples.append(f"第{r_idx}行{col_name}列宽={w}twips（期望约{target}twips）")
                cell_ok = False

    hit = grid_ok and cell_ok
    return ScorePoint(
        "收支记录表格列宽：Date/Description/Withdrawals/Deposits/Balance 五列均为统一3.68厘米，各页续表列宽保持一致",
        3,
        hit,
        f"目标={target}twips({twips_to_cm(target):.2f}cm)；tblGrid={widths}；逐行单元格检查={'通过' if cell_ok else '失败'}；异常示例={'; '.join(bad_examples) if bad_examples else '无'}",
    )


def check_data_font(tbl: ET.Element) -> ScorePoint:
    # 细则：所有数据行 Arial 五号（10.5磅）黑色；
    #       列0(Date)和列1(Description)常规字形；列4(Balance)加粗；
    #       列2(Withdrawals)和列3(Deposits)不约束字形。
    bad: list[str] = []
    col_names = HEADER  # Date/Description/Withdrawals/Deposits/Balance
    for r_idx, tr in enumerate(rows(tbl)[1:], start=1):
        cells = tr.findall("./w:tc", NS)
        for c_idx, tc in enumerate(cells):
            text = cell_text(tc)
            # 空白金额单元格没有可见文字，不作为失败原因。
            if not text:
                continue
            rprs = text_run_rprs(tc)
            # 所有数据行：Arial 五号（10.5磅）、黑色
            common_ok = bool(rprs) and all(
                font_name_ok(rpr) and font_size_ok(rpr, 10.5) and color_black_ok(rpr)
                for rpr in rprs
            )
            # 日期(列0)和交易说明(列1)：常规字形
            if c_idx in (0, 1):
                style_ok = all(is_regular(rpr) for rpr in rprs)
            # Balance列(列4)：加粗
            elif c_idx == 4:
                style_ok = all(is_bold(rpr) for rpr in rprs)
            # Withdrawals(列2)和Deposits(列3)：细则未约束字形
            else:
                style_ok = True
            if not (common_ok and style_ok) and len(bad) < 5:
                col_label = col_names[c_idx] if c_idx < len(col_names) else str(c_idx + 1)
                bad.append(f"数据第{r_idx}行{col_label}列='{text}'")
    return ScorePoint(
        "收支记录数据字体：所有数据行Arial五号黑色；日期/交易说明常规字形；Balance列加粗",
        1,
        not bad,
        "异常示例=" + ("；".join(bad) if bad else "无"),
    )


def check_paragraph_format(tbl: ET.Element) -> ScorePoint:
    # 细则：数据单元格段前0磅、段后10磅、1.15倍行距；单元格内不出现额外空段落或多余换行。
    # OOXML 对应：before=0twips，after=200twips（10磅），line=276，lineRule=auto（1.15倍行距）。
    bad: list[str] = []
    for r_idx, tr in enumerate(rows(tbl)[1:], start=1):
        for c_idx, tc in enumerate(tr.findall("./w:tc", NS), start=1):
            if not paragraph_format_ok(tc, before=0, after=200, line=276, line_rule="auto") and len(bad) < 5:
                p = first(tc, "./w:p")
                spacing = first(first(p, "./w:pPr"), "./w:spacing") if p is not None else None
                ps_count = len(tc.findall("./w:p", NS))
                br_count = len(tc.findall(".//w:br", NS))
                bad.append(
                    f"数据第{r_idx}行第{c_idx}列 段落数={ps_count} 换行符={br_count} "
                    f"spacing(before={attr(spacing, 'w:before')}, after={attr(spacing, 'w:after')}, "
                    f"line={attr(spacing, 'w:line')}, lineRule={attr(spacing, 'w:lineRule')})"
                )
    return ScorePoint(
        "收支记录段落格式：数据单元格段前0磅、段后10磅、1.15倍行距，单元格内不出现额外空段落或多余换行",
        1,
        not bad,
        "异常示例=" + ("；".join(bad) if bad else "无"),
    )


def border_edge(tc: ET.Element, edge: str) -> ET.Element | None:
    return first(first(first(tc, "./w:tcPr"), "./w:tcBorders"), f"./w:{edge}")


def is_light_gray(color: str | None) -> bool:
    if color is None:
        return False
    c = color.upper()
    if not re.fullmatch(r"[0-9A-F]{6}", c):
        return False
    r, g, b = int(c[0:2], 16), int(c[2:4], 16), int(c[4:6], 16)
    return abs(r - g) <= 8 and abs(g - b) <= 8 and 150 <= r <= 230


def check_borders(tbl: ET.Element) -> ScorePoint:
    # 细则：每一条记录下方使用浅灰色细实线，线宽0.75磅。
    bad_bottom: list[str] = []
    data_rows = rows(tbl)[1:]
    for r_idx, tr in enumerate(data_rows, start=1):
        cells = tr.findall("./w:tc", NS)
        for c_idx, tc in enumerate(cells, start=1):
            bottom = border_edge(tc, "bottom")
            try:
                bottom_sz = int(attr(bottom, "w:sz") or "0")
            except ValueError:
                bottom_sz = 0
            bottom_ok = (
                attr(bottom, "w:val") == "single"
                # OOXML 边框 sz 单位为 1/8 磅，6 = 0.75 磅。
                and bottom_sz == 6
                and is_light_gray(attr(bottom, "w:color"))
            )
            if not bottom_ok and len(bad_bottom) < 5:
                bad_bottom.append(
                    f"数据第{r_idx}行第{c_idx}列 bottom(val={attr(bottom, 'w:val')}, "
                    f"sz={attr(bottom, 'w:sz')}，期望6(0.75磅), color={attr(bottom, 'w:color')})"
                )
    return ScorePoint(
        "收支记录横向边框：每一条记录下方使用浅灰色细实线，线宽0.75磅",
        3,
        not bad_bottom,
        f"底边异常={'; '.join(bad_bottom) if bad_bottom else '无'}",
    )


def check_money_format(tbl: ET.Element) -> ScorePoint:
    # 细则：Withdrawals(列2)、Deposits(列3)、Balance(列4)列金额统一显示为美元格式，
    #       使用千位分隔符并保留两位小数，例如 "$6,870.20"、"$43,712.48"。
    col_names = {2: "Withdrawals", 3: "Deposits", 4: "Balance"}
    bad: list[str] = []
    for r_idx, tr in enumerate(rows(tbl)[1:], start=1):
        cells = row_texts(tr)
        for c_idx, col_name in col_names.items():
            text = cells[c_idx] if c_idx < len(cells) else ""
            # 空白单元格不作为失败原因（收入/支出二选一为空属正常）。
            if not text:
                continue
            if not canonical_money(text) and len(bad) < 5:
                bad.append(f"数据第{r_idx}行{col_name}列='{text}'")
    return ScorePoint(
        "收支记录金额格式：Withdrawals/Deposits/Balance列金额为美元格式，千位分隔符，两位小数",
        1,
        not bad,
        "异常示例=" + ("；".join(bad) if bad else "无"),
    )


def check_record_count_completion(tbl: ET.Element) -> ScorePoint:
    # 细则：除表头外共50行数据。
    recs = data_records(tbl)
    total_ok = len(recs) == 50
    return ScorePoint(
        "记录完整数量：除表头外共50行数据",
        5,
        total_ok,
        f"实际数据行={len(recs)}（期望50）",
    )


def check_repeating_header(tbl: ET.Element) -> ScorePoint:
    # 细则：记录表格跨页时，每一页顶部重复显示 Date/Description/Withdrawals/Deposits/Balance 表头；
    #       列宽、字体和对齐方式与前一页一致。
    # 本项只检查“续表表头之间是否一致”，不重复约束 3.68cm、Arial 11.5磅、左右对齐等另一评分项的具体格式。
    def cell_font_signature(tc: ET.Element) -> tuple[tuple[str | None, str | None, str | None, str | None, bool, bool, str | None], ...]:
        signature: list[tuple[str | None, str | None, str | None, str | None, bool, bool, str | None]] = []
        for rpr in text_run_rprs(tc):
            fonts = first(rpr, "./w:rFonts") if rpr is not None else None
            size = first(rpr, "./w:sz") if rpr is not None else None
            color = first(rpr, "./w:color") if rpr is not None else None
            signature.append((
                attr(fonts, "w:ascii"),
                attr(fonts, "w:hAnsi"),
                attr(fonts, "w:eastAsia"),
                attr(size, "w:val"),
                is_bold(rpr),
                is_true_word_value(first(rpr, "./w:i") if rpr is not None else None),
                attr(color, "w:val"),
            ))
        return tuple(signature)

    def header_signature(tr: ET.Element) -> tuple[tuple[int | None, ...], tuple[tuple[tuple[str | None, str | None, str | None, str | None, bool, bool, str | None], ...], ...], tuple[str | None, ...]]:
        cells = tr.findall("./w:tc", NS)
        widths = tuple(cell_width_twips(tc) for tc in cells)
        fonts = tuple(cell_font_signature(tc) for tc in cells)
        aligns = tuple(paragraph_alignment(tc) for tc in cells)
        return widths, fonts, aligns

    all_rows = rows(tbl)
    header = all_rows[0]
    header_cells = header.findall("./w:tc", NS)

    tbl_header = first(first(header, "./w:trPr"), "./w:tblHeader")
    repeat_ok = is_true_word_value(tbl_header)
    text_ok = row_texts(header) == HEADER and len(header_cells) == 5

    repeated_headers = [tr for tr in all_rows if row_texts(tr) == HEADER]
    signatures = [header_signature(tr) for tr in repeated_headers]
    width_consistency_ok = bool(signatures) and all(sig[0] == signatures[idx - 1][0] for idx, sig in enumerate(signatures[1:], start=1))
    font_consistency_ok = bool(signatures) and all(sig[1] == signatures[idx - 1][1] for idx, sig in enumerate(signatures[1:], start=1))
    align_consistency_ok = bool(signatures) and all(sig[2] == signatures[idx - 1][2] for idx, sig in enumerate(signatures[1:], start=1))

    hit = repeat_ok and text_ok and width_consistency_ok and font_consistency_ok and align_consistency_ok
    detail = (
        f"w:tblHeader={attr(tbl_header, 'w:val') if tbl_header is not None else None}，"
        + f"表头={row_texts(header)}，重复表头来源数={len(repeated_headers)}，"
        + f"列宽一致={width_consistency_ok}，字体一致={font_consistency_ok}，对齐一致={align_consistency_ok}"
    )
    return ScorePoint(
        "续表表头：记录表格跨页时，每一页顶部重复显示Date/Description/Withdrawals/Deposits/Balance表头，列宽、字体和对齐方式与前一页一致",
        1,
        hit,
        detail,
    )


def cover_picture_size_penalty(ctx: DocxContext) -> ScorePoint:
    # 只检查真正的图片 pic:pic，不把封面装饰形状算作封面图片；取面积最大的图片作为封面图片。
    pics: list[tuple[float, float]] = []
    for pic in ctx.document.findall(".//pic:pic", NS):
        parent = None
        # ElementTree 无父指针；向上查找包含该 pic 的 drawing。
        for drawing in ctx.document.findall(".//w:drawing", NS):
            if pic in list(drawing.iter()):
                parent = drawing
                break
        extent = first(parent, ".//wp:extent") if parent is not None else None
        try:
            cx = int(extent.get("cx")) if extent is not None and extent.get("cx") else 0
            cy = int(extent.get("cy")) if extent is not None and extent.get("cy") else 0
        except ValueError:
            continue
        if cx and cy:
            pics.append((emu_to_cm(cx), emu_to_cm(cy)))
    if not pics:
        return ScorePoint("扣分：封面图片大小不满足宽15-17cm高9-11cm", -1, True, "未找到可测量图片，按不满足扣分")
    width, height = max(pics, key=lambda wh: wh[0] * wh[1])
    bad = not (15 <= width <= 17 and 9 <= height <= 11)
    return ScorePoint(
        "扣分：封面图片大小不满足宽15-17cm高9-11cm",
        -1,
        bad,
        f"最大图片尺寸={width:.2f}cm × {height:.2f}cm；{'不满足' if bad else '满足'}",
    )


def title_is_wordart_penalty(ctx: DocxContext) -> ScorePoint:
    title = "Activity details — 出支记录"
    bad = False
    details: list[str] = []

    # WordArt/艺术字在 OOXML 中并没有单一的"是否为艺术字"标志，Word 各版本会用不同的载体：
    #   - Word 2010+ 走 DrawingML：<w:drawing>/<wp:*>/<a:*> 中的 <wps:txbx>/<w:txbxContent> 文本形状。
    #   - Word 97-2007 走 VML：<w:pict>/<v:shape>，其中 <v:shapetype spt=136> 或 t202xxx 系列为经典 WordArt。
    #   - 文本本身通常带有 <w14:textFill/textOutline/textEffect/glow/reflection/shadow/scene3d/props3d> 之一。
    #   - 通过样式继承（<w:pStyle>/<w:rStyle>）间接引用一个带上述文本效果的段落/字符样式。
    # 判定标题是否为艺术体：以下任一条件命中即视为艺术字。

    # 收集所有含标题的段落，供后续多路径判定。
    title_paragraphs: list[ET.Element] = [p for p in ctx.document.findall(".//w:p", NS) if title in all_text(p)]

    def _iter_local_names(el: ET.Element) -> set[str]:
        return {node.tag.split("}", 1)[-1] for node in el.iter()}

    # 常见 WordArt/文本效果的 tag 本地名集合（Drawing/VML/Word2010 扩展）。
    effect_locals = {
        # w14 文本效果（Word 2010+ 常用于标题艺术字）
        "textFill", "textOutline", "textEffect", "glow", "reflection", "shadow", "scene3d", "props3d",
        # DrawingML 3D/艺术效果
        "sp3d", "bevelT", "bevelB", "camera", "lightRig", "gradFill",
        # DrawingML/VML 变形（WordArt 的核心变形，如 textArch/textCanUp 等 preset shape）
        "prstTxWarp",
    }
    # VML WordArt 的经典载体：v:shape/v:shapetype/v:textpath、o:OLEObject（ProgID 含 WordArt），
    # 以及 shapetype spt=136（Word 中所有 WordArt shape 的类型编号）。下方按本地名直接比较。

    for p in title_paragraphs:
        # 1) DrawingML 文本形状（wps:txbx/w:txbxContent 或 wp:*）承载标题文字。
        for drawing in p.findall(".//w:drawing", NS):
            if title in all_text(drawing):
                bad = True
                details.append("标题位于 DrawingML 文本形状/图形对象内")
                # 检查其中的 DrawingML 文本效果节点。
                locals_in_drawing = _iter_local_names(drawing)
                hit_effects = sorted(locals_in_drawing & effect_locals)
                if hit_effects:
                    details.append(f"DrawingML 含艺术字效果：{hit_effects}")

        # 2) VML 对象承载标题：<w:pict>/<v:shape>/<v:shapetype>/<v:textpath>。
        for pict in p.findall(".//w:pict", NS):
            if title in all_text(pict):
                bad = True
                details.append("标题位于 VML(pict) 对象内")
                # v:textpath 是 WordArt 文本沿路径变形的核心元素；spt=136 亦为 WordArt shape。
                for node in pict.iter():
                    local = node.tag.split("}", 1)[-1]
                    if local == "textpath":
                        details.append("VML 含 <v:textpath>（艺术字文本路径）")
                    if local == "shapetype":
                        spt = node.get(qn("o:spt")) or node.get("spt")
                        if spt == "136":
                            details.append("VML shapetype spt=136（经典 WordArt 类型）")
                    if local == "OLEObject":
                        prog = node.get(qn("r:ProgID")) or node.get("ProgID") or ""
                        if "WordArt" in prog or "MSDraw" in prog:
                            details.append(f"OLE 对象为艺术字：{prog}")

        # 3) 段落/run 上直接挂 WordArt 文本效果节点。
        for rpr in [first(r, "./w:rPr") for r in p.findall(".//w:r", NS)]:
            if rpr is None:
                continue
            for node in rpr.iter():
                local = node.tag.split("}", 1)[-1]
                if local in effect_locals:
                    bad = True
                    details.append(f"标题 run 含文本效果 <{local}>")

        # 4) 段落/run 通过样式继承引用了含艺术字效果的字符/段落样式。
        style_ids: set[str] = set()
        p_pr = first(p, "./w:pPr")
        p_style = first(p_pr, "./w:pStyle") if p_pr is not None else None
        if p_style is not None:
            sid = p_style.get(qn("w:val"))
            if sid:
                style_ids.add(sid)
        for r in p.findall(".//w:r", NS):
            rpr = first(r, "./w:rPr")
            if rpr is None:
                continue
            r_style = first(rpr, "./w:rStyle")
            if r_style is not None:
                sid = r_style.get(qn("w:val"))
                if sid:
                    style_ids.add(sid)
        for sid in style_ids:
            if _style_has_wordart_effect(ctx, sid):
                bad = True
                details.append(f"标题通过样式 {sid!r} 继承艺术字效果")

    # 去重，保留出现顺序。
    seen: set[str] = set()
    dedup: list[str] = []
    for d in details:
        if d not in seen:
            seen.add(d)
            dedup.append(d)
    return ScorePoint(
        "扣分：封面“Activity details — 出支记录”为艺术体",
        -1,
        bad,
        "；".join(dedup) if dedup else "未检测到艺术字/文本效果",
    )


def _style_has_wordart_effect(ctx: DocxContext, style_id: str, _seen: set[str] | None = None) -> bool:
    """在 styles.xml 中递归查找 styleId 及其 basedOn 祖先是否挂载 WordArt/文本效果节点。"""
    if ctx.styles is None or not style_id:
        return False
    if _seen is None:
        _seen = set()
    if style_id in _seen:
        return False
    _seen.add(style_id)
    effect_locals = {
        "textFill", "textOutline", "textEffect", "glow", "reflection", "shadow", "scene3d", "props3d",
        "sp3d", "bevelT", "bevelB", "camera", "lightRig", "gradFill", "prstTxWarp",
    }
    for style in ctx.styles.findall(".//w:style", NS):
        if style.get(qn("w:styleId")) != style_id:
            continue
        for node in style.iter():
            local = node.tag.split("}", 1)[-1]
            if local in effect_locals:
                return True
        based_on = first(style, "./w:basedOn")
        if based_on is not None:
            parent_id = based_on.get(qn("w:val"))
            if parent_id and _style_has_wordart_effect(ctx, parent_id, _seen):
                return True
    return False


def blank_amount_cells_penalty(tbl: ET.Element) -> ScorePoint:
    # 细则：属于收入记录时 Withdrawals 单元格保持空白；属于支出记录时 Deposits 单元格保持空白；
    #       不能使用 0、横杠或错误金额填充空白列。
    actual = data_records(tbl)
    bad: list[str] = []
    pseudo_blank_values = {"0", "0.00", "$0", "$0.00", "-", "—", "–"}

    for idx, row in enumerate(actual, start=1):
        desc = row[1]
        withdrawal, deposit = row[2], row[3]

        # 本条细则只约束收入/支出记录，不额外约束 Balance forward 或 Closing totals。
        if desc in {"Balance forward", "Closing totals"}:
            continue

        withdrawal_is_blank = not withdrawal.strip()
        deposit_is_blank = not deposit.strip()
        withdrawal_is_pseudo_blank = withdrawal.strip() in pseudo_blank_values
        deposit_is_pseudo_blank = deposit.strip() in pseudo_blank_values

        if withdrawal_is_pseudo_blank:
            bad.append(f"第{idx}行 Withdrawals 使用0/横杠填充空白列：{withdrawal!r}")
        if deposit_is_pseudo_blank:
            bad.append(f"第{idx}行 Deposits 使用0/横杠填充空白列：{deposit!r}")

        if withdrawal and deposit:
            bad.append(f"第{idx}行空白列被错误金额填充：Withdrawals={withdrawal!r}, Deposits={deposit!r}")
        elif withdrawal_is_blank and deposit_is_blank:
            bad.append(f"第{idx}行无法判定收入/支出记录：Withdrawals和Deposits均为空")
        elif withdrawal and deposit_is_blank:
            # 支出记录：Deposits 单元格保持空白。
            pass
        elif deposit and withdrawal_is_blank:
            # 收入记录：Withdrawals 单元格保持空白。
            pass

    return ScorePoint(
        "扣分：收支记录空白金额单元格不满足：收入记录Withdrawals空白、支出记录Deposits空白，不能用0、横杠或错误金额填充空白列",
        -3,
        bool(bad),
        "异常示例=" + ("；".join(bad[:5]) if bad else "无"),
    )


def dimension2(ctx: DocxContext, tbl: ET.Element) -> list[ScorePoint]:
    return [
        check_table_format(tbl),
        check_column_widths(tbl),
        check_data_font(tbl),
        check_paragraph_format(tbl),
        check_borders(tbl),
        check_money_format(tbl),
        check_record_count_completion(tbl),
        check_repeating_header(tbl),
        cover_picture_size_penalty(ctx),
        title_is_wordart_penalty(ctx),
        blank_amount_cells_penalty(tbl),
    ]


def _locate_docx(directory: Path) -> Path | None:
    # 在脚本所在目录内定位待评估的 .docx 文件（忽略 Word 临时文件 ~$xxx.docx）。
    candidates = [p for p in sorted(directory.iterdir()) if p.is_file() and p.suffix.lower() == ".docx" and not p.name.startswith("~$")]
    return candidates[0] if candidates else None


def evaluate(dir_path: str) -> dict[str, object]:
    # 统一入口：接收脚本所在目录的路径，脚本自己在目录内定位并打开被评估文档，返回结构化评估结果。
    result: dict[str, object] = {
        "id": "036",
        "file_name": "",
        "status": "ok",
        "error": None,
        "dim1_pass": False,
        "dim1_reason": "",
        "dim2_items": [],
        "total_score": 0,
        "max_score": 0,
    }
    ctx: DocxContext | None = None
    try:
        directory = Path(dir_path)
        if not directory.exists() or not directory.is_dir():
            result["status"] = "error"
            result["error"] = f"目录不存在或不是目录：{dir_path}"
            return result

        docx_path = _locate_docx(directory)
        if docx_path is None:
            result["status"] = "error"
            result["error"] = f"目录内未找到 .docx 文件：{dir_path}"
            return result
        result["file_name"] = docx_path.name

        ctx, load_error = load_docx(docx_path)
        dim1_ok, dim1_checks, tbl = dimension1(ctx, load_error)
        result["dim1_pass"] = dim1_ok
        if not dim1_ok:
            failed = [f"{c.name}（{c.detail}）" if c.detail else c.name for c in dim1_checks if not c.passed]
            result["dim1_reason"] = "；".join(failed)

        if dim1_ok and ctx is not None and tbl is not None:
            points = dimension2(ctx, tbl)
            items: list[dict[str, object]] = []
            total = 0
            max_score = 0
            for p in points:
                delta = p.points if p.hit else 0
                rule = p.name.removeprefix("扣分：") if p.name.startswith("扣分：") else p.name
                items.append({
                    "rule": rule,
                    "max_delta": p.points,
                    "delta": delta,
                    "hit": p.hit,
                    "detail": "",
                })
                total += delta
                if p.points > 0:
                    max_score += p.points
            result["dim2_items"] = items
            result["total_score"] = total
            result["max_score"] = max_score
        else:
            # 维度一未通过：维度二不计分，dim2_items 保持为空，满分按维度二正向评分点静态合计。
            result["total_score"] = 0
            result["max_score"] = 18
    except Exception as exc:  # noqa: BLE001 - 顶层兜底：任何异常都归到 status=error。
        result["status"] = "error"
        result["error"] = f"{type(exc).__name__}: {exc}"
    finally:
        if ctx is not None:
            try:
                ctx.zip_file.close()
            except Exception:  # noqa: BLE001
                pass
    return result


if __name__ == "__main__":
    # 本地调试入口：默认对脚本所在目录发起评估，也支持从命令行传入目录路径。
    target_dir = sys.argv[1] if len(sys.argv) > 1 else str(Path(__file__).resolve().parent)
    print(json.dumps(evaluate(target_dir), ensure_ascii=False, indent=2))
