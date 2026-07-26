#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
自动评估《家居零售价目册_零售价1.1倍替换完成.docx》。

对外统一接口：
    from officeval_033_verifier import evaluate
    result = evaluate(dir_path)   # dir_path 为脚本所在目录路径
    # result 为结构化 dict，字段见 evaluate() 文档字符串

实现原则：
1. 先检查“维度1：可用与可修改性”。不通过则维度2不再评分。
2. 维度2按评分细则逐条自动检查：命中得分点累计正分，命中扣分点累计负分。
3. 只使用 Python 标准库解析 docx（docx 本质是 zip + XML），不依赖人工判断或第三方库。
"""

from __future__ import annotations

import json
import re
import sys
import zipfile
from dataclasses import dataclass, field
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Sequence, Tuple
import xml.etree.ElementTree as ET

NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "v": "urn:schemas-microsoft-com:vml",
}
W = NS["w"]
SCRIPT_ID = "033"
NOTE_TEXT = "注：图形为原创示意图形，色彩仅用于表达材质层次；实际交付以订购确认单为准。"


def wtag(name: str) -> str:
    return f"{{{W}}}{name}"


def attr(ns: str, name: str) -> str:
    return f"{{{NS[ns]}}}{name}"


def cell_text(el: ET.Element, include_nested: bool = True) -> str:
    """提取元素文字。include_nested=False 时只取单元格直接段落，避免外层版式表吞掉内层表格内容。"""
    if include_nested:
        return "".join(t.text or "" for t in el.findall(".//w:t", NS))
    parts: List[str] = []
    for p in el.findall("w:p", NS):
        parts.append("".join(t.text or "" for t in p.findall(".//w:t", NS)))
    return "".join(parts)


def norm_text(s: str) -> str:
    return re.sub(r"\s+", "", s or "").replace("“", "").replace("”", "").replace('"', "")


def norm_price(s: str) -> str:
    """用于正向得分项比对：允许空格、逗号和￥符号，不允许把多个数字拼成一个价格。"""
    s = (s or "").strip()
    s = s.replace(",", "").replace("，", "").replace("￥", "").replace("¥", "").replace("元", "")
    s = re.sub(r"\s+", "", s)
    return s if re.fullmatch(r"\d+", s) else ""


def values_match_in_order(actual: Sequence[str], expected: Sequence[int]) -> bool:
    """判断 expected 是否按从上到下顺序出现在 actual 中，允许 actual 有额外行。"""
    wanted = [str(x) for x in expected]
    i = 0
    for value in actual:
        if i < len(wanted) and norm_price(value) == wanted[i]:
            i += 1
    return i == len(wanted)


def printable_values(values: Sequence[str]) -> str:
    return "[" + ", ".join(v if v else "空" for v in values) + "]"


@dataclass
class TableInfo:
    page: int
    depth: int
    index_on_page: int
    order: int
    rows: List[List[str]]
    kind: str = "other"

    @property
    def header(self) -> List[str]:
        return self.rows[0] if self.rows else []


@dataclass
class PageStats:
    text_chars: int = 0
    table_count: int = 0
    image_count: int = 0
    max_image_area_ratio: float = 0.0


@dataclass
class DocxInfo:
    path: Path
    zip_names: List[str]
    document_xml: str
    root: ET.Element
    tables: List[TableInfo]
    pages: Dict[int, PageStats]
    page_count: int
    comments_count: int
    red_text_chars: int
    red_run_count: int
    has_embedded_pdf: bool
    header_footer_image_only_parts: List[str]
    note_style_failures: List[str]
    note_count: int


@dataclass
class CheckResult:
    title: str
    score: int
    hit: bool
    details: List[str] = field(default_factory=list)


class DocxParseError(Exception):
    pass


def count_page_breaks(el: ET.Element) -> int:
    return sum(1 for br in el.findall(".//w:br", NS) if br.get(wtag("type")) == "page")


def drawing_area_ratios(el: ET.Element, page_area_emu2: int) -> List[float]:
    ratios: List[float] = []
    if page_area_emu2 <= 0:
        return ratios
    for extent in el.findall(".//wp:extent", NS):
        try:
            cx = int(extent.get("cx") or "0")
            cy = int(extent.get("cy") or "0")
        except ValueError:
            continue
        if cx > 0 and cy > 0:
            ratios.append((cx * cy) / page_area_emu2)
    return ratios


def get_page_area_emu2(root: ET.Element) -> int:
    # w:pgSz 单位是 twentieths of a point；1 twip = 635 EMU。
    pg_sz = root.find(".//w:sectPr/w:pgSz", NS)
    if pg_sz is not None:
        try:
            w = int(pg_sz.get(wtag("w")) or "0") * 635
            h = int(pg_sz.get(wtag("h")) or "0") * 635
            if w > 0 and h > 0:
                return w * h
        except ValueError:
            pass
    # A4 兜底。
    return 7772400 * 10058400


def table_rows(tbl: ET.Element) -> List[List[str]]:
    rows: List[List[str]] = []
    for tr in tbl.findall("w:tr", NS):
        row: List[str] = []
        for tc in tr.findall("w:tc", NS):
            row.append(cell_text(tc, include_nested=False).strip())
        rows.append(row)
    return rows


def classify_table(rows: List[List[str]]) -> str:
    if not rows:
        return "other"
    header = [norm_text(c) for c in rows[0]]
    header_joined = "|".join(header)

    if all(h in header_joined for h in ["定制料号", "物料描述", "尺寸", "A1", "A2", "A3", "A4", "A5", "A6", "标准零售价"]):
        return "custom"
    if all(h in header_joined for h in ["ERP料号", "规格说明", "物料描述", "零售价（元）", "备注"]):
        return "combo"
    if all(h in header_joined for h in ["料号", "物料描述", "零售价（元）", "备注"]) and "ERP料号" not in header_joined:
        return "recommended"
    if all(h in header_joined for h in ["产品组", "基础零售价（元）", "材质调价", "尺寸调价", "核算后零售价（元）"]):
        return "conversion"
    if all(h in header_joined for h in ["序号", "产品编码", "产品名称", "零售价（元）"]):
        return "order"
    return "other"


def update_page_stats(stats: PageStats, el: ET.Element, page_area_emu2: int, table_delta: int = 0) -> None:
    text = cell_text(el, include_nested=True)
    stats.text_chars += len(re.sub(r"\s+", "", text))
    stats.table_count += table_delta
    stats.image_count += len(el.findall(".//w:drawing", NS)) + len(el.findall(".//w:pict", NS))
    ratios = drawing_area_ratios(el, page_area_emu2)
    if ratios:
        stats.max_image_area_ratio = max(stats.max_image_area_ratio, max(ratios))


def collect_nested_tables(el: ET.Element, page: int, depth: int, order_ref: List[int], out: List[TableInfo], stats: PageStats) -> None:
    per_depth_count = 0
    for child in list(el):
        if child.tag == wtag("tbl"):
            per_depth_count += 1
            order_ref[0] += 1
            rows = table_rows(child)
            info = TableInfo(page=page, depth=depth, index_on_page=per_depth_count, order=order_ref[0], rows=rows)
            info.kind = classify_table(rows)
            out.append(info)
            stats.table_count += 1
            collect_nested_tables(child, page, depth + 1, order_ref, out, stats)
        else:
            collect_nested_tables(child, page, depth, order_ref, out, stats)


def parse_docx(path: Path) -> DocxInfo:
    if path.suffix.lower() != ".docx":
        raise DocxParseError("交付文件不是 .docx 格式")
    try:
        zf = zipfile.ZipFile(path)
    except zipfile.BadZipFile as exc:
        raise DocxParseError(f"文件不是可正常打开的 docx/zip：{exc}") from exc

    with zf:
        names = zf.namelist()
        if "word/document.xml" not in names:
            raise DocxParseError("docx 中缺少 word/document.xml，无法正常解析正文")
        try:
            document_xml = zf.read("word/document.xml").decode("utf-8")
            root = ET.fromstring(document_xml)
        except Exception as exc:  # noqa: BLE001 - 输出给评估报告
            raise DocxParseError(f"document.xml 解析失败：{exc}") from exc

        content_types = zf.read("[Content_Types].xml").decode("utf-8", errors="ignore") if "[Content_Types].xml" in names else ""
        has_embedded_pdf = any(n.lower().endswith(".pdf") for n in names) or "application/pdf" in content_types.lower()

        page_area = get_page_area_emu2(root)
        body = root.find("w:body", NS)
        if body is None:
            raise DocxParseError("docx 正文为空，缺少 w:body")

        tables: List[TableInfo] = []
        pages: Dict[int, PageStats] = {1: PageStats()}
        page = 1
        order_ref = [0]
        top_table_count_by_page: Dict[int, int] = {}

        for child in list(body):
            pages.setdefault(page, PageStats())
            if child.tag == wtag("p"):
                update_page_stats(pages[page], child, page_area)
                page += count_page_breaks(child)
                pages.setdefault(page, PageStats())
            elif child.tag == wtag("tbl"):
                top_table_count_by_page[page] = top_table_count_by_page.get(page, 0) + 1
                order_ref[0] += 1
                rows = table_rows(child)
                info = TableInfo(page=page, depth=0, index_on_page=top_table_count_by_page[page], order=order_ref[0], rows=rows)
                info.kind = classify_table(rows)
                tables.append(info)
                update_page_stats(pages[page], child, page_area, table_delta=1)
                collect_nested_tables(child, page, 1, order_ref, tables, pages[page])
                page += count_page_breaks(child)
                pages.setdefault(page, PageStats())

        comments_count = 0
        for n in names:
            if n.startswith("word/comments") and n.endswith(".xml"):
                try:
                    croot = ET.fromstring(zf.read(n))
                    comments_count += len(croot.findall(".//w:comment", NS))
                except ET.ParseError:
                    comments_count += 1

        red_text_chars, red_run_count = count_red_text(root)
        header_footer_image_only_parts = inspect_header_footer_parts(zf, names)
        note_count, note_style_failures = inspect_note_styles(root)

    return DocxInfo(
        path=path,
        zip_names=names,
        document_xml=document_xml,
        root=root,
        tables=tables,
        pages=pages,
        page_count=max(pages.keys()) if pages else 0,
        comments_count=comments_count,
        red_text_chars=red_text_chars,
        red_run_count=red_run_count,
        has_embedded_pdf=has_embedded_pdf,
        header_footer_image_only_parts=header_footer_image_only_parts,
        note_style_failures=note_style_failures,
        note_count=note_count,
    )


def count_red_text(root: ET.Element) -> Tuple[int, int]:
    red_chars = 0
    red_runs = 0
    red_values = {"FF0000", "C00000", "E60000", "FF3333", "DC143C"}
    for r in root.findall(".//w:r", NS):
        text = "".join(t.text or "" for t in r.findall(".//w:t", NS))
        if not text.strip():
            continue
        rpr = r.find("w:rPr", NS)
        if rpr is None:
            continue
        color = rpr.find("w:color", NS)
        highlight = rpr.find("w:highlight", NS)
        color_val = (color.get(wtag("val")) if color is not None else "") or ""
        highlight_val = (highlight.get(wtag("val")) if highlight is not None else "") or ""
        if color_val.upper() in red_values or highlight_val.lower() == "red":
            red_chars += len(text.strip())
            red_runs += 1
    return red_chars, red_runs


def inspect_header_footer_parts(zf: zipfile.ZipFile, names: Sequence[str]) -> List[str]:
    image_only: List[str] = []
    for n in names:
        low = n.lower()
        if not (low.startswith("word/header") or low.startswith("word/footer")) or not low.endswith(".xml"):
            continue
        try:
            root = ET.fromstring(zf.read(n))
        except ET.ParseError:
            image_only.append(f"{n} 无法解析")
            continue
        txt_len = len(re.sub(r"\s+", "", cell_text(root, include_nested=True)))
        img_count = len(root.findall(".//w:drawing", NS)) + len(root.findall(".//w:pict", NS))
        if img_count > 0 and txt_len < 2:
            image_only.append(n)
    return image_only


def inspect_note_styles(root: ET.Element) -> Tuple[int, List[str]]:
    failures: List[str] = []
    count = 0
    for p in root.findall(".//w:p", NS):
        paragraph_text = cell_text(p, include_nested=True)
        if NOTE_TEXT not in paragraph_text:
            continue
        count += 1
        for r in p.findall("w:r", NS):
            run_text = "".join(t.text or "" for t in r.findall(".//w:t", NS))
            if not run_text.strip():
                continue
            rpr = r.find("w:rPr", NS)
            if rpr is None:
                failures.append(f"第{count}处说明文字 run 缺少字体属性")
                continue
            rfonts = rpr.find("w:rFonts", NS)
            font_values = []
            if rfonts is not None:
                for key in ["ascii", "hAnsi", "eastAsia", "cs"]:
                    val = rfonts.get(wtag(key))
                    if val:
                        font_values.append(val)
            sz = rpr.find("w:sz", NS)
            size_val = sz.get(wtag("val")) if sz is not None else None
            if "Noto Sans CJK SC" not in font_values or size_val != "13":
                failures.append(
                    f"第{count}处说明文字样式异常：字体={font_values or '未声明'}，字号w:sz={size_val or '未声明'}"
                )
    return count, failures


def whole_page_image_pages(info: DocxInfo) -> List[int]:
    bad: List[int] = []
    for page, st in sorted(info.pages.items()):
        # 细则只扣“整页图片化，导致该页价格和文字不可编辑”：必须有图片，且几乎没有可编辑文字/真实表格。
        image_only_page = st.image_count > 0 and st.text_chars < 5 and st.table_count <= 1
        if image_only_page and (st.max_image_area_ratio >= 0.65 or st.max_image_area_ratio == 0):
            bad.append(page)
    return bad


def find_header_index(row: Sequence[str], header: str) -> Optional[int]:
    target = norm_text(header)
    for idx, cell in enumerate(row):
        if norm_text(cell) == target:
            return idx
    for idx, cell in enumerate(row):
        if target in norm_text(cell):
            return idx
    return None


def table_on_page(info: DocxInfo, page: int, kind: str, ordinal: int = 1, product_price_order: bool = False) -> Optional[TableInfo]:
    if product_price_order:
        matches = [t for t in info.tables if t.page == page and t.kind in {"combo", "recommended", "custom"}]
        matches.sort(key=lambda t: t.order)
        if len(matches) < ordinal:
            return None
        table = matches[ordinal - 1]
        return table if table.kind == kind else None
    matches = [t for t in info.tables if t.page == page and t.kind == kind]
    matches.sort(key=lambda t: t.order)
    return matches[ordinal - 1] if len(matches) >= ordinal else None


def column_values(table: TableInfo, column_name: str) -> List[str]:
    if not table.rows:
        return []
    col = find_header_index(table.rows[0], column_name)
    if col is None:
        return []
    vals: List[str] = []
    for row in table.rows[1:]:
        vals.append(row[col].strip() if col < len(row) else "")
    return vals


def check_column_sequence(
    info: DocxInfo,
    page: int,
    kind: str,
    column_name: str,
    expected: Sequence[int],
    ordinal: int = 1,
    product_price_order: bool = False,
) -> Tuple[bool, str]:
    table = table_on_page(info, page, kind, ordinal=ordinal, product_price_order=product_price_order)
    if table is None:
        scope = "产品页价格表" if product_price_order else f"{kind} 类型表格"
        return False, f"第{page}页未找到第{ordinal}个 {scope}（期望为 {kind} 表）"
    vals = column_values(table, column_name)
    ok = values_match_in_order(vals, expected)
    status = "✓" if ok else "✗"
    ordinal_label = f"第{ordinal}个" if ordinal != 1 or product_price_order else ""
    scope_label = "产品页价格表" if product_price_order else f"{kind} 表"
    return ok, f"{status} 第{page}页{ordinal_label}{scope_label} `{column_name}` 实际={printable_values(vals)}，期望按序出现={list(expected)}"


def evaluate_positive_points(info: DocxInfo) -> List[CheckResult]:
    results: List[CheckResult] = []

    groups = [
        (
            "+5：产品页组合规格零售价表",
            5,
            [
                (4, "combo", "零售价（元）", [2317, 3814, 4340, 18059], 1),
                (5, "combo", "零售价（元）", [28397, 22470, 42585], 2),
                (63, "combo", "零售价（元）", [29140, 46305, 17328, 38779], 2),
                (62, "combo", "零售价（元）", [18154, 47152, 14699], 2),
                (35, "combo", "零售价（元）", [9390, 45214, 27634], 2),
            ],
        ),
        (
            "+5：产品页推荐客餐厅配套品表",
            5,
            [
                (4, "recommended", "零售价（元）", [13841, 4241, 8125], 2),
                (5, "recommended", "零售价（元）", [9478, 13182, 12638], 3),
                (63, "recommended", "零售价（元）", [5457, 13956, 3090], 1),
                (62, "recommended", "零售价（元）", [12344, 10546, 8342], 3),
            ],
        ),
    ]

    for title, score, checks in groups:
        details: List[str] = []
        ok_all = True
        for check in checks:
            page, kind, col, expected = check[:4]
            ordinal = check[4] if len(check) > 4 else 1
            ok, detail = check_column_sequence(
                info, page, kind, col, expected, ordinal=ordinal, product_price_order=True
            )
            ok_all = ok_all and ok
            details.append(detail)
        results.append(CheckResult(title=title, score=score, hit=ok_all, details=details))

    custom_checks = [
        (4, 3, {
            "A1": [14942, 16097, 17882], "A2": [15564, 16765, 18564], "A3": [16060, 17268, 19223],
            "A4": [16696, 18046, 19969], "A5": [17597, 18944, 21025], "A6": [18736, 20136, 22408],
            "标准零售价": [15564, 16097, 21025],
        }),
        (5, 1, {
            "A1": [9365, 10998, 11673], "A2": [9765, 11452, 12120], "A3": [10040, 11835, 12558],
            "A4": [10482, 12357, 13100], "A5": [11053, 12969, 13743], "A6": [11778, 13795, 14620],
            "标准零售价": [9765, 11835, 12558],
        }),
        (62, 1, {
            "A1": [7497, 8713], "A2": [7809, 9086], "A3": [8102, 9384], "A4": [8449, 9756],
            "A5": [8842, 10302], "A6": [9466, 10891], "标准零售价": [8102, 10891],
        }),
        (63, 3, {
            "A1": [13343, 14331], "A2": [13883, 14895], "A3": [14306, 15371], "A4": [14971, 16017],
            "A5": [15741, 16848], "A6": [16704, 17939], "标准零售价": [13343, 14331],
        }),
        (35, 1, {
            "A1": [6811, 8267], "A2": [7107, 8624], "A3": [7346, 8858], "A4": [7618, 9309],
            "A5": [8070, 9722], "A6": [8601, 10393], "标准零售价": [7346, 8267],
        }),
    ]
    details = []
    ok_all = True
    for page, ordinal, columns in custom_checks:
        for col, expected in columns.items():
            ok, detail = check_column_sequence(
                info, page, "custom", col, expected, ordinal=ordinal, product_price_order=True
            )
            ok_all = ok_all and ok
            details.append(detail)
    results.append(CheckResult(title="+5：产品页定制价格表", score=5, hit=ok_all, details=details))

    p65_details = []
    p65_ok1, d1 = check_column_sequence(info, 65, "conversion", "基础零售价（元）", [28249, 10441, 15744, 12546])
    p65_ok2, d2 = check_column_sequence(info, 65, "conversion", "核算后零售价（元）", [34458, 13518, 19049, 14631])
    p65_details.extend([d1, d2])
    results.append(CheckResult(title="+1：订购与交付说明页价格换算提醒表", score=1, hit=p65_ok1 and p65_ok2, details=p65_details))

    p66_ok, p66_detail = check_column_sequence(
        info,
        66,
        "order",
        "零售价（元）",
        [40943, 32309, 42543, 41592, 32511, 17915, 1636, 8999, 5831, 1131, 9920],
    )
    results.append(CheckResult(title="+1：门店订购记录页零售价列", score=1, hit=p66_ok, details=[p66_detail]))

    return results


def price_columns_for_table(table: TableInfo) -> List[int]:
    if not table.rows:
        return []
    headers = table.rows[0]
    wanted = ["零售价（元）", "基础零售价（元）", "核算后零售价（元）", "A1", "A2", "A3", "A4", "A5", "A6", "标准零售价"]
    cols = []
    for h in wanted:
        idx = find_header_index(headers, h)
        if idx is not None and idx not in cols:
            cols.append(idx)
    return cols


def strict_valid_price_cell(s: str) -> bool:
    raw = s or ""
    # 超过页面50%的空白：单元格内容去除空白后几乎为空，视为大面积空白异常。
    if len(raw.strip()) == 0:
        return False
    compact = raw.strip().replace(",", "").replace("，", "").replace("￥", "").replace("¥", "").replace("元", "")
    # 乱码、非数字文本或错误符号：去除合法分隔符后不是纯数字，则为异常。
    return bool(re.fullmatch(r"\d+", compact))


def detect_invalid_price_cells(info: DocxInfo) -> List[str]:
    problems: List[str] = []
    for table in info.tables:
        if table.kind not in {"combo", "recommended", "custom", "conversion", "order"}:
            continue
        price_cols = price_columns_for_table(table)
        for r_idx, row in enumerate(table.rows[1:], start=2):
            for col in price_cols:
                value = row[col] if col < len(row) else ""
                if not strict_valid_price_cell(value):
                    header = table.rows[0][col] if col < len(table.rows[0]) else f"第{col + 1}列"
                    problems.append(f"第{table.page}页 {table.kind} 表 第{r_idx}行 `{header}` 单元格异常：{value!r}")
    return problems


def evaluate_negative_points(info: DocxInfo) -> List[CheckResult]:
    results: List[CheckResult] = []

    image_pages = whole_page_image_pages(info)
    results.append(CheckResult(
        title="-5：文档任意一页被转换为整页图片，导致该页价格和文字不可编辑",
        score=-5,
        hit=bool(image_pages),
        details=[f"疑似整页图片页：{image_pages}"] if image_pages else ["未发现疑似整页图片页"],
    ))

    invalid_price = detect_invalid_price_cells(info)
    results.append(CheckResult(
        title="-1：任意价格单元格出现超过页面50%的空白、乱码、非数字文本或错误符号",
        score=-1,
        hit=bool(invalid_price),
        details=invalid_price[:20] if invalid_price else ["所有识别到的价格单元格均为纯数字且非空"],
    ))

    note_bad = info.note_count == 0 or bool(info.note_style_failures)
    note_details = []
    if info.note_count == 0:
        note_details.append("未找到要求保持的价格说明文字")
    elif info.note_style_failures:
        note_details.extend(info.note_style_failures[:20])
    else:
        note_details.append(f"共 {info.note_count} 处价格说明文字，均为 Noto Sans CJK SC 且 w:sz=13（小六）")
    results.append(CheckResult(
        title='-3：价格说明文字保持情况不满足："注：图形为原创示意图形，色彩仅用于表达材质层次；实际交付以订购确认单为准。"等说明文字保持Noto Sans CJK SC小六不变',
        score=-3,
        hit=note_bad,
        details=note_details,
    ))

    return results


def evaluate_dimension1(path: Path) -> Tuple[bool, List[str], Optional[DocxInfo]]:
    details: List[str] = []
    if not path.exists():
        return False, [f"文件不存在：{path}"], None
    try:
        info = parse_docx(path)
    except DocxParseError as exc:
        return False, [str(exc)], None

    details.append("✓ 文件扩展名为 .docx，且可作为 Office Open XML 文档正常解析")
    details.append(f"✓ 正文 XML 可读取；按显式分页符识别到约 {info.page_count} 页、{len(info.tables)} 个表格（含嵌套表格）")

    return True, details, info


def print_result(result: CheckResult) -> None:
    marker = "命中" if result.hit else "未命中"
    earned = result.score if result.hit else 0
    print(f"{marker} {result.title} => {earned:+d} 分")
    for d in result.details:
        print(f"  {d}")


def _find_docx_in_dir(dir_path: Path) -> Optional[Path]:
    """在给定目录中寻找待评估的 .docx 文件，忽略以 ~$ 开头的 Office 临时文件。"""
    if not dir_path.is_dir():
        return None
    candidates = [
        p for p in dir_path.iterdir()
        if p.is_file() and p.suffix.lower() == ".docx" and not p.name.startswith("~$")
    ]
    if not candidates:
        return None
    # 若有多个 docx，优先选择名称包含“价目册”的目标文件；否则按名称排序取第一个。
    preferred = [p for p in candidates if "价目册" in p.name]
    if preferred:
        return sorted(preferred)[0]
    return sorted(candidates)[0]


def evaluate(dir_path: str) -> dict:
    """统一入口：接收脚本所在目录路径，自行在该目录内定位并评估 docx 文件。

    返回结构：
        {
            "id": "033",
            "file_name": "xxx.docx",
            "status": "ok" | "error",
            "error": None | str,
            "dim1_pass": bool,
            "dim1_reason": str,
            "dim2_items": [
                {"rule": str, "max_delta": int, "delta": int, "hit": bool, "detail": str},
                ...
            ],
            "total_score": int,
            "max_score": int,
        }
    """
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
        base = Path(dir_path)
        if not base.exists():
            result["status"] = "error"
            result["error"] = f"目录不存在：{dir_path}"
            return result
        if not base.is_dir():
            result["status"] = "error"
            result["error"] = f"路径不是目录：{dir_path}"
            return result

        docx_path = _find_docx_in_dir(base)
        if docx_path is None:
            result["status"] = "error"
            result["error"] = f"目录中未找到 .docx 文件：{dir_path}"
            return result
        result["file_name"] = docx_path.name

        dim1_ok, dim1_details, info = evaluate_dimension1(docx_path)
        result["dim1_pass"] = bool(dim1_ok)
        if not dim1_ok or info is None:
            # 维度一是一票否决门槛：未通过时短路返回，
            # total_score=0、dim2_items=[]，不再生成任何维度二占位项。
            fails = [d for d in dim1_details if d.startswith("✗")] or dim1_details
            result["dim1_reason"] = "；".join(fails)
            result["total_score"] = 0
            result["dim2_items"] = []
            return result

        checks = evaluate_positive_points(info) + evaluate_negative_points(info)
        total = 0
        max_score = 0
        for chk in checks:
            delta = chk.score if chk.hit else 0
            total += delta
            if chk.score > 0:
                max_score += chk.score
            result["dim2_items"].append({
                "rule": chk.title,
                "max_delta": chk.score,
                "delta": delta,
                "hit": bool(chk.hit),
                "detail": "",
            })
        result["total_score"] = total
        result["max_score"] = max_score
        return result
    except Exception as exc:  # noqa: BLE001 - 顶层兜底，统一转成 status=error
        result["status"] = "error"
        result["error"] = f"{type(exc).__name__}: {exc}"
        return result


if __name__ == "__main__":
    _arg = sys.argv[1] if len(sys.argv) > 1 else str(Path(__file__).resolve().parent)
    print(json.dumps(evaluate(_arg), ensure_ascii=False, indent=2))
