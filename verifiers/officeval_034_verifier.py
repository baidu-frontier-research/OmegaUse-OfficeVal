#!/usr/bin/env python3
"""
PDF 自动评分脚本：智能办公设备租赁及支持服务合同

对外只暴露一个入口函数 ``evaluate(dir_path: str) -> dict``：接收脚本所在目录的路径，
脚本在该目录内定位并打开待评估的 PDF（默认文件名为
``managed_workspace_equipment_leasing_support_contract_formatted.pdf``；若不存在则回退
到目录内首个 PDF 文件），返回统一的结构化结果字典（见项目《脚本接口差异与统一建议》§2.2）。
依赖：pdfplumber（经 pdf_backend 适配层）。如环境缺失，可安装：python -m pip install pdfplumber

评分逻辑：
1. 先检查“维度1：可用与可修改性”。若不满足，直接返回 0 分并跳过维度2。
2. 维度1通过后，按“维度2：完成度评分细则”逐项加分/扣分。
3. 命中项与未命中项都在 ``dim2_items`` 中返回，便于批量汇总。

说明：PDF 的版面语义无法像源 DOCX 一样完全确定，本脚本对“可自动检测但边界模糊”的要求
采用可解释的启发式规则，例如用文字块/绘图线条识别表格、用下划线文本宽度估计填写线长度、
用文本框重叠率估计大面积重叠等，以尽量贴近评分意图且不依赖人工判断。
"""

from __future__ import annotations

import json

SCRIPT_ID = "034"
import math
import os
import re
import statistics
import sys
from dataclasses import dataclass, field
from typing import Iterable, Optional

try:
    try:
        import pdf_backend
    except ImportError:
        from verifiers import pdf_backend
except ImportError as exc:  # pragma: no cover
    raise ImportError("缺少依赖 pdfplumber。请先执行：python -m pip install pdfplumber") from exc

PdfRect = pdf_backend.PdfRect


class _Page:
    """页面句柄：提供 rect 与矢量绘图访问（替代 fitz.Page）。"""

    def __init__(self, doc: "pdf_backend.PdfDocument", index: int):
        self._doc = doc
        self.index = index
        w, h = doc.page_size(index)
        self.rect = PdfRect(0.0, 0.0, w, h)

    def raw_lines(self):
        return self._doc.extract_raw_lines(self.index, line_gap=6.0)

    def drawings(self):
        return self._doc.extract_drawings(self.index)

PDF_NAME = "managed_workspace_equipment_leasing_support_contract_formatted.pdf"
PT_PER_CM = 72 / 2.54
MIN_FILL_LINE_PT = 4 * PT_PER_CM
MAX_FILL_LINE_PT = 7 * PT_PER_CM
# 允许 PDF 字体度量、下划线字符宽度和排版压缩带来的偏差；PDF抽取无法直接给出
# 下划线子串的真实宽度，按字符比例估算会略偏大，因此保留较宽容差。
FILL_LINE_TOLERANCE = 0.40
FILL_LINE_MIN_TOL = MIN_FILL_LINE_PT * (1 - FILL_LINE_TOLERANCE)
FILL_LINE_MAX_TOL = MAX_FILL_LINE_PT * (1 + FILL_LINE_TOLERANCE)
# Signature-page fields have an explicit 4cm-7cm requirement. Keep only a small
# PDF extraction tolerance here; the generic 40% tolerance would let 8-9cm lines pass.
SIGNATURE_FILL_LINE_TOLERANCE = 0.08
SIGNATURE_FILL_LINE_MIN_TOL = MIN_FILL_LINE_PT * (1 - SIGNATURE_FILL_LINE_TOLERANCE)
SIGNATURE_FILL_LINE_MAX_TOL = MAX_FILL_LINE_PT * (1 + SIGNATURE_FILL_LINE_TOLERANCE)

CN_TITLES = {
    1: "第一条 合同标的及设备清单",
    2: "第二条 服务期限与交付",
    3: "第三条 费用及付款安排",
    4: "第四条 所有权及使用限制",
    5: "第五条 安装、网络环境及访问",
    6: "第六条 维护及服务水平",
    7: "第七条 数据处理与保密",
    8: "第八条 安全措施",
    9: "第九条 变更请求及额外工作",
    10: "第十条 遗失、损坏及保险",
    11: "第十一条 审核、记录及盘点",
    12: "第十二条 暂停与终止",
    13: "第十三条 返还及最终结算",
    14: "第十四条 责任限制",
    15: "第十五条 争议解决及适用法律",
    16: "第十六条 其他",
}

EN_TITLES = {
    1: "Clause 1: Contract Subject and Equipment List",
    2: "Clause 2: Service Period and Delivery",
    3: "Clause 3: Fees and Payment Arrangement",
    4: "Clause 4: Ownership and Use Restrictions",
    5: "Clause 5: Installation, Network Environment and Access",
    6: "Clause 6: Maintenance and Service Level",
    7: "Clause 7: Data Handling and Confidentiality",
    8: "Clause 8: Security Measures",
    9: "Clause 9: Change Requests and Additional Work",
    10: "Clause 10: Loss, Damage and Insurance",
    11: "Clause 11: Audit, Records and Inventory Check",
    12: "Clause 12: Suspension and Termination",
    13: "Clause 13: Return and Final Settlement",
    14: "Clause 14: Liability Limitation",
    15: "Clause 15: Dispute Resolution and Governing Law",
    16: "Clause 16: Miscellaneous",
}


@dataclass
class TextLine:
    page_index: int
    text: str
    bbox: PdfRect
    spans: list[dict]

    @property
    def x0(self) -> float:
        return self.bbox.x0

    @property
    def y0(self) -> float:
        return self.bbox.y0

    @property
    def x1(self) -> float:
        return self.bbox.x1

    @property
    def y1(self) -> float:
        return self.bbox.y1

    @property
    def width(self) -> float:
        return self.bbox.width

    @property
    def height(self) -> float:
        return self.bbox.height

    def main_font(self) -> str:
        if not self.spans:
            return ""
        return self.spans[0].get("font", "")

    def main_size(self) -> float:
        if not self.spans:
            return 0.0
        return float(self.spans[0].get("size", 0.0))

    def is_blackish(self) -> bool:
        # 适配层 span color 为整数 0xRRGGBB，0 是黑色。
        if not self.spans:
            return True
        colors = [int(s.get("color", 0)) for s in self.spans]
        return all(c <= 0x303030 for c in colors)


@dataclass
class TableBox:
    page_index: int
    rect: PdfRect
    divider_x: Optional[float] = None
    header_bottom_y: Optional[float] = None

    @property
    def width(self) -> float:
        return self.rect.width

    def left_width(self) -> Optional[float]:
        return None if self.divider_x is None else self.divider_x - self.rect.x0

    def right_width(self) -> Optional[float]:
        return None if self.divider_x is None else self.rect.x1 - self.divider_x


@dataclass
class CheckResult:
    name: str
    score: int
    passed: bool
    detail: str


@dataclass
class PdfInfo:
    path: str
    doc: "pdf_backend.PdfDocument"
    pages: list[_Page] = field(default_factory=list)
    lines_by_page: list[list[TextLine]] = field(default_factory=list)
    text_by_page: list[str] = field(default_factory=list)
    tables_by_page: list[list[TableBox]] = field(default_factory=list)

    @property
    def all_text(self) -> str:
        return "\n".join(self.text_by_page)


def normalize_text(text: str) -> str:
    return re.sub(r"\s+", " ", text).strip()


def rect_area(rect: PdfRect) -> float:
    return max(0.0, rect.width) * max(0.0, rect.height)


def intersection_area(a: PdfRect, b: PdfRect) -> float:
    x0 = max(a.x0, b.x0)
    y0 = max(a.y0, b.y0)
    x1 = min(a.x1, b.x1)
    y1 = min(a.y1, b.y1)
    if x1 <= x0 or y1 <= y0:
        return 0.0
    return (x1 - x0) * (y1 - y0)


def get_text_lines(page: _Page, page_index: int) -> list[TextLine]:
    result: list[TextLine] = []
    for line in page.raw_lines():
        if not line.text.strip():
            continue
        # span 转为与历史 rawdict 兼容的字典结构（text/font/size/color/bbox/chars）
        spans = [
            {
                "text": sp.text,
                "font": sp.font,
                "size": sp.size,
                "color": sp.color,
                "bbox": sp.bbox,
                "chars": [
                    {"c": ch.c, "bbox": (ch.bbox.x0, ch.bbox.y0, ch.bbox.x1, ch.bbox.y1)}
                    for ch in sp.chars
                ],
            }
            for sp in line.spans
        ]
        result.append(TextLine(page_index, line.text.strip(), line.bbox, spans))
    result.sort(key=lambda ln: (ln.y0, ln.x0))
    return result


def line_contains(line: TextLine, needle: str) -> bool:
    return needle in normalize_text(line.text)


def find_lines(lines: Iterable[TextLine], needle: str) -> list[TextLine]:
    return [line for line in lines if line_contains(line, needle)]


def find_first_line(lines: Iterable[TextLine], needle: str) -> Optional[TextLine]:
    matches = find_lines(lines, needle)
    return matches[0] if matches else None


def approx_underline_widths(line: TextLine) -> list[float]:
    """Estimate consecutive underscore/tab underline widths in PDF points."""
    widths: list[float] = []
    for span in line.spans:
        text = span.get("text", "")
        if not text:
            continue
        bbox = span.get("bbox", line.bbox)
        span_width = max(0.0, bbox.width)
        # Treat tab-underlines and long underscore runs as fill lines.
        for match in re.finditer(r"[_\t]{6,}", text):
            # Character-count proportional estimate is robust when a whole line is one span.
            widths.append(span_width * (match.end() - match.start()) / max(1, len(text)))
    if widths:
        return widths
    # Fallback if extractor split is odd.
    for match in re.finditer(r"[_\t]{6,}", line.text):
        widths.append(line.width * (match.end() - match.start()) / max(1, len(line.text)))
    return widths


def measured_underline_widths(line: TextLine) -> list[float]:
    """Measure underscore runs from extracted character boxes when available."""
    widths: list[float] = []
    for span in line.spans:
        text = span.get("text", "")
        chars = span.get("chars") or []
        if not text or len(chars) != len(text):
            continue
        for match in re.finditer(r"[_\t]{6,}", text):
            run_chars = chars[match.start() : match.end()]
            if not run_chars:
                continue
            x0 = min(float(ch.get("bbox", [0, 0, 0, 0])[0]) for ch in run_chars)
            x1 = max(float(ch.get("bbox", [0, 0, 0, 0])[2]) for ch in run_chars)
            if x1 > x0:
                widths.append(x1 - x0)
    return widths or approx_underline_widths(line)


def has_fill_line(line: TextLine, relaxed_short: bool = False) -> bool:
    widths = approx_underline_widths(line)
    if not widths:
        return False
    lower = FILL_LINE_MIN_TOL if not relaxed_short else 2.3 * PT_PER_CM
    return any(lower <= w <= FILL_LINE_MAX_TOL for w in widths)


def has_signature_fill_line(line: TextLine) -> bool:
    widths = measured_underline_widths(line)
    return any(SIGNATURE_FILL_LINE_MIN_TOL <= w <= SIGNATURE_FILL_LINE_MAX_TOL for w in widths)


def underline_width_between(line: TextLine, before: str, after: str) -> Optional[float]:
    """测量夹在 before 与 after 两段文字之间那段连续下划线的宽度（PDF点）。

    仅当下划线（下划线字符或制表符下划线）确实出现在 before 之后、after 之前时返回其宽度，
    否则返回 None。优先用抽取到的字符盒精确测量，无字符盒时回退到按字符比例估算。
    """
    for span in line.spans:
        text = span.get("text", "")
        if not text or before not in text or after not in text:
            continue
        b_end = text.index(before) + len(before)
        a_start = text.index(after, b_end)
        if a_start < b_end:
            continue
        middle = text[b_end:a_start]
        m = re.search(r"[_\t]{6,}", middle)
        if not m:
            continue
        run_start = b_end + m.start()
        run_end = b_end + m.end()
        chars = span.get("chars") or []
        if len(chars) == len(text):
            run_chars = chars[run_start:run_end]
            x0 = min(float(ch.get("bbox", [0, 0, 0, 0])[0]) for ch in run_chars)
            x1 = max(float(ch.get("bbox", [0, 0, 0, 0])[2]) for ch in run_chars)
            if x1 > x0:
                return x1 - x0
        bbox = span.get("bbox", line.bbox)
        return max(0.0, bbox.width) * (run_end - run_start) / max(1, len(text))
    return None


def signing_fill_line_ok(line: TextLine, before: str, after: str) -> bool:
    """签署说明填写线：位于 before 与 after 之间，长度约 4cm-7cm。"""
    width = underline_width_between(line, before, after)
    if width is None:
        return False
    return SIGNATURE_FILL_LINE_MIN_TOL <= width <= SIGNATURE_FILL_LINE_MAX_TOL


def signature_underline_consistent(lines: Iterable[TextLine]) -> bool:
    widths: list[float] = []
    for line in lines:
        valid = [w for w in measured_underline_widths(line) if SIGNATURE_FILL_LINE_MIN_TOL <= w <= SIGNATURE_FILL_LINE_MAX_TOL]
        if not valid:
            return False
        widths.append(valid[0])
    if len(widths) < 8:
        return False
    avg = sum(widths) / len(widths)
    if avg <= 0:
        return False
    return max(abs(w - avg) / avg for w in widths) <= 0.20


def underline_consistent(lines: Iterable[TextLine]) -> bool:
    # Collect widths from all lines. For address fields that overflow to a second line,
    # `__...` appears as a standalone line; include widths down to 1.8cm so we don't
    # drop legitimate continuation overflow underlines.
    widths = [w for line in lines for w in approx_underline_widths(line) if 1.8 * PT_PER_CM <= w <= FILL_LINE_MAX_TOL]
    if len(widths) < 3:
        return False
    # Only require the *main* fill-lines (4-7cm range) to be consistent.
    long_widths = [w for w in widths if w >= FILL_LINE_MIN_TOL]
    basis = long_widths if len(long_widths) >= 3 else widths
    avg = sum(basis) / len(basis)
    if avg <= 0:
        return False
    return max(abs(w - avg) / avg for w in basis) <= 0.40


def horizontal_segments(page: _Page) -> list[tuple[float, float, float, float, Optional[tuple]]]:
    segments: list[tuple[float, float, float, float, Optional[tuple]]] = []
    for d in page.drawings():
        r = d.rect
        if r.width > 1 and r.height < 1.5:
            segments.append((r.x0, r.x1, (r.y0 + r.y1) / 2, d.line_width, d.stroke_color))
    return segments


def vertical_segments(page: _Page) -> list[tuple[float, float, float, float]]:
    segments: list[tuple[float, float, float, float]] = []
    for d in page.drawings():
        r = d.rect
        if r.height > 1 and r.width < 1.5:
            segments.append(((r.x0 + r.x1) / 2, r.y0, r.y1, d.line_width))
    return segments


def detect_tables(page: _Page, page_index: int) -> list[TableBox]:
    """Detect rectangular two-column tables from vector horizontal/vertical lines."""
    hs = horizontal_segments(page)
    vs = vertical_segments(page)
    page_width = page.rect.width
    candidates: list[TableBox] = []

    # Large horizontal rules likely representing table top/bottom/rows.
    long_h = sorted(
        [(x0, x1, y, w) for x0, x1, y, w, _ in hs if (x1 - x0) >= page_width * 0.55],
        key=lambda t: (t[2], t[0]),
    )
    used: set[tuple[int, int]] = set()
    for i, top in enumerate(long_h):
        for j in range(i + 1, len(long_h)):
            bottom = long_h[j]
            if (i, j) in used:
                continue
            top_x0, top_x1, top_y, _ = top
            bot_x0, bot_x1, bot_y, _ = bottom
            if bot_y - top_y < 45:
                continue
            # Top and bottom should have almost same left/right boundary.
            if abs(top_x0 - bot_x0) > 4 or abs(top_x1 - bot_x1) > 4:
                continue
            x0, x1 = (top_x0 + bot_x0) / 2, (top_x1 + bot_x1) / 2
            # Need left and right vertical boundaries spanning most table height.
            boundary_vs = [v for v in vs if v[1] <= top_y + 3 and v[2] >= bot_y - 3]
            has_left = any(abs(v[0] - x0) <= 4 for v in boundary_vs)
            has_right = any(abs(v[0] - x1) <= 4 for v in boundary_vs)
            dividers = [v for v in boundary_vs if x0 + (x1 - x0) * 0.35 <= v[0] <= x0 + (x1 - x0) * 0.65]
            if not (has_left and has_right and dividers):
                continue
            divider_x = sorted(dividers, key=lambda v: abs(v[0] - (x0 + x1) / 2))[0][0]
            inner_rows = [h for h in long_h if abs(h[0] - x0) <= 5 and abs(h[1] - x1) <= 5 and top_y < h[2] < bot_y]
            header_bottom = min((h[2] for h in inner_rows), default=None)
            rect = PdfRect(x0, top_y, x1, bot_y)
            # Avoid duplicates from double-stroked borders.
            if any(abs(rect.x0 - t.rect.x0) < 3 and abs(rect.y0 - t.rect.y0) < 3 and abs(rect.y1 - t.rect.y1) < 3 for t in candidates):
                continue
            candidates.append(TableBox(page_index=page_index, rect=rect, divider_x=divider_x, header_bottom_y=header_bottom))
            used.add((i, j))
            break
    candidates.sort(key=lambda t: (t.rect.y0, t.rect.x0))
    return candidates


def load_pdf(path: str) -> PdfInfo:
    doc = pdf_backend.open_pdf(path)
    info = PdfInfo(path=path, doc=doc)
    info.pages = [_Page(doc, i) for i in range(doc.page_count)]
    for idx, page in enumerate(info.pages):
        lines = get_text_lines(page, idx)
        info.lines_by_page.append(lines)
        info.text_by_page.append("\n".join(ln.text for ln in lines))
        info.tables_by_page.append(detect_tables(page, idx))
    return info


def max_vertical_blank_ratio(page: _Page, lines: list[TextLine]) -> float:
    intervals = []
    for line in lines:
        # Ignore footer page number for blank-body detection.
        if re.fullmatch(r"\d+", line.text.strip()) and line.y0 > page.rect.height * 0.9:
            continue
        intervals.append((max(0.0, line.y0), min(page.rect.height, line.y1)))
    for d in page.drawings():
        if rect_area(d.rect) > 5:
            intervals.append((max(0.0, d.rect.y0), min(page.rect.height, d.rect.y1)))
    if not intervals:
        return 1.0
    intervals.sort()
    cursor = 0.0
    max_gap = 0.0
    for y0, y1 in intervals:
        if y0 > cursor:
            max_gap = max(max_gap, y0 - cursor)
        cursor = max(cursor, y1)
    max_gap = max(max_gap, page.rect.height - cursor)
    return max_gap / page.rect.height


def text_overlap_problem(lines: list[TextLine]) -> bool:
    # Only compare text lines on similar vertical rows; normal table cells in separate columns should not overlap.
    checked = 0
    bad = 0
    for i, a in enumerate(lines):
        if a.text.strip().isdigit():
            continue
        for b in lines[i + 1 :]:
            if b.y0 > a.y1 + 3:
                break
            if b.text.strip().isdigit():
                continue
            inter = intersection_area(a.bbox, b.bbox)
            if inter <= 0:
                continue
            denom = min(rect_area(a.bbox), rect_area(b.bbox))
            if denom > 0:
                checked += 1
                if inter / denom > 0.25:
                    bad += 1
    return bad >= 3 or (checked >= 5 and bad / checked > 0.2)


def dimension1(info: PdfInfo) -> tuple[bool, list[str], dict[str, bool]]:
    failures: list[str] = []
    flags = {
        "openable_pdf": True,
    }

    if info.doc.page_count == 0:
        failures.append("PDF 可打开但没有页面")
        flags["openable_pdf"] = False
        return False, failures, flags

    return not failures, failures, flags


def title_horizontal_aligned(line: TextLine, page: _Page) -> bool:
    """左对齐或居中（不含垂直位置约束）。"""
    left_ok = line.x0 <= page.rect.width * 0.18
    center_ok = abs((line.x0 + line.x1) / 2 - page.rect.width / 2) <= page.rect.width * 0.08
    return left_ok or center_ok


def title_alignment_ok(line: TextLine, page: _Page) -> bool:
    top_ok = line.y0 <= page.rect.height * 0.14
    return top_ok and title_horizontal_aligned(line, page)


def check_home_titles(info: PdfInfo) -> tuple[CheckResult, CheckResult, CheckResult]:
    page = info.pages[0]
    lines = info.lines_by_page[0]
    cn = find_first_line(lines, "智能办公设备租赁及支持服务合同")
    en = find_first_line(lines, "Managed Workspace Equipment Leasing and Support Contract")
    cn_ok = bool(cn and title_alignment_ok(cn, page))
    en_ok = bool(en and cn and title_horizontal_aligned(en, page) and en.y0 > cn.y0)
    # -3 细则：首页（第1页）没有出现中文标题或英文标题即扣分。此项只判定“是否出现”，
    # 不约束位置/对齐/上下顺序（那些属于 +1 首页标题项）。用首页整页文本、去除全部空白后
    # 做子串匹配，兼容办公软件导出 PDF 时标题因排版换行被拆成多行的情况——只要标题文字
    # 出现在首页即视为存在。
    home_text = re.sub(r"\s+", "", info.text_by_page[0])
    cn_present = re.sub(r"\s+", "", "智能办公设备租赁及支持服务合同") in home_text
    en_present = re.sub(r"\s+", "", "Managed Workspace Equipment Leasing and Support Contract") in home_text
    both_ok = cn_present and en_present
    return (
        CheckResult("+1 首页中文标题", 1, cn_ok, "第1页顶部检测到中文标题且位置为左对齐/居中" if cn_ok else "未在第1页顶部正文区域检测到合规中文标题"),
        CheckResult("+1 首页英文标题", 1, en_ok, "中文标题下方检测到英文标题且位置为左对齐/居中" if en_ok else "未在中文标题下方检测到合规英文标题"),
        CheckResult("-3 首页标题缺失", -3, not both_ok, "首页未出现中文标题或英文标题" if not both_ok else "首页中英文标题均出现，不扣分"),
    )


def check_signing_intro(info: PdfInfo) -> CheckResult:
    lines = info.lines_by_page[0]
    en_title = find_first_line(lines, "Managed Workspace Equipment Leasing and Support Contract")
    cn_intro = next((ln for ln in lines if "本合同由以下双方于" in ln.text and "签署" in ln.text), None)
    en_intro = next((ln for ln in lines if "This Contract is entered into on" in ln.text and "by and between" in ln.text), None)
    ok = bool(
        cn_intro
        and en_intro
        and (en_title is None or cn_intro.y0 > en_title.y0)
        and en_intro.y0 > cn_intro.y0
        and en_intro.y0 - cn_intro.y0 <= 28
        and signing_fill_line_ok(cn_intro, "双方于", "签署")
        and signing_fill_line_ok(en_intro, "entered into on", "by and between")
    )
    return CheckResult(
        "+1 首页合同签署说明区",
        1,
        ok,
        "检测到中英文签署说明、中文行在上英文行紧随、下划线位于“双方于/签署”与“on/by”之间且约4-7cm" if ok else "未检测到合规的中英文签署说明区/填写线",
    )


def adjacent_pair(lines: list[TextLine], cn_kw: str, en_kw: str, start_y: float, end_y: float) -> bool:
    cn_lines = [ln for ln in lines if start_y <= ln.y0 <= end_y and cn_kw in ln.text]
    en_lines = [ln for ln in lines if start_y <= ln.y0 <= end_y and en_kw in ln.text]
    for cn in cn_lines:
        for en in en_lines:
            if 0 < en.y0 - cn.y0 <= 35 and abs(en.x0 - cn.x0) <= 12:
                return True
    return False


def info_section_lines(info: PdfInfo, party: str) -> list[TextLine]:
    lines = info.lines_by_page[0]
    if party == "A":
        start = next((ln.y0 for ln in lines if "甲方 / 服务提供方" in ln.text), 0)
        end = next((ln.y0 for ln in lines if "乙方 / 客户方" in ln.text), 10**9)
    else:
        start = next((ln.y0 for ln in lines if "乙方 / 客户方" in ln.text), 0)
        end = next((ln.y0 for ln in lines if "甲方与乙方以下合称" in ln.text), 10**9)
    return [ln for ln in lines if start <= ln.y0 < end]


def field_fill_widths(section: list[TextLine], field_labels: list[str]) -> Optional[list[float]]:
    """按“字段标签在前、其后紧跟一段下划线”的方式，测量每个字段对应填写线的宽度（PDF点）。

    支持同一行内含多个字段（各字段各自取其标签后的那段下划线），也支持填写线折行到
    紧邻下一行（该下一行整行是一段独立下划线）。任一字段找不到填写线则返回 None。
    """
    # 按 y 排序，便于处理折行到下一行的填写线。
    ordered = sorted(section, key=lambda ln: (ln.y0, ln.x0))
    widths: list[float] = []
    for label in field_labels:
        width: Optional[float] = None
        for idx, ln in enumerate(ordered):
            if label not in ln.text:
                continue
            # 情形1：同一行内，标签后紧跟下划线。
            w = underline_width_after(ln, label)
            if w is not None:
                width = w
                break
            # 情形2：标签在行尾，填写线折到下一行（下一行以下划线开头）。
            if idx + 1 < len(ordered):
                nxt = ordered[idx + 1]
                if 0 < nxt.y0 - ln.y0 <= 22 and re.match(r"\s*[_\t]{6,}", nxt.text):
                    w2 = leading_underline_width(nxt)
                    if w2 is not None:
                        width = w2
                        break
        if width is None:
            return None
        widths.append(width)
    return widths


def underline_width_after(line: TextLine, before: str) -> Optional[float]:
    """测量 before 之后紧邻那段连续下划线的宽度（PDF点）；无则返回 None。"""
    for span in line.spans:
        text = span.get("text", "")
        if not text or before not in text:
            continue
        b_end = text.index(before) + len(before)
        m = re.match(r"[ \t：:]*([_\t]{6,})", text[b_end:])
        if not m:
            continue
        run_start = b_end + m.start(1)
        run_end = b_end + m.end(1)
        chars = span.get("chars") or []
        if len(chars) == len(text):
            run_chars = chars[run_start:run_end]
            x0 = min(float(ch.get("bbox", [0, 0, 0, 0])[0]) for ch in run_chars)
            x1 = max(float(ch.get("bbox", [0, 0, 0, 0])[2]) for ch in run_chars)
            if x1 > x0:
                return x1 - x0
        bbox = span.get("bbox", line.bbox)
        return max(0.0, bbox.width) * (run_end - run_start) / max(1, len(text))
    return None


def leading_underline_width(line: TextLine) -> Optional[float]:
    """测量整行开头那段连续下划线的宽度（PDF点）；无则返回 None。"""
    for span in line.spans:
        text = span.get("text", "")
        m = re.match(r"\s*([_\t]{6,})", text)
        if not m:
            continue
        run_start, run_end = m.start(1), m.end(1)
        chars = span.get("chars") or []
        if len(chars) == len(text):
            run_chars = chars[run_start:run_end]
            x0 = min(float(ch.get("bbox", [0, 0, 0, 0])[0]) for ch in run_chars)
            x1 = max(float(ch.get("bbox", [0, 0, 0, 0])[2]) for ch in run_chars)
            if x1 > x0:
                return x1 - x0
        bbox = span.get("bbox", line.bbox)
        return max(0.0, bbox.width) * (run_end - run_start) / max(1, len(text))
    return None


def check_party_section(info: PdfInfo, party: str) -> CheckResult:
    lines = info.lines_by_page[0]
    section = info_section_lines(info, party)
    if party == "A":
        title_cn, title_en = "甲方 / 服务提供方", "Party A / Service Provider"
        name = "+3 首页甲方信息区"
    else:
        title_cn, title_en = "乙方 / 客户方", "Party B / Client"
        name = "+3 首页乙方信息区"
    text = "\n".join(ln.text for ln in section)
    fields_present = all(
        kw in text
        for kw in [title_cn, title_en, "统一识别号", "UEN", "授权代表", "Authorized Representative", "地址", "Address", "行政联系", "Administrative Contact"]
    )
    start = min((ln.y0 for ln in section), default=0)
    end = max((ln.y1 for ln in section), default=0)
    # 每个中文字段行与其对应英文行上下相邻排列：逐项检查五组字段对。
    pair_specs = [
        (title_cn, title_en),
        ("统一识别号", "UEN"),
        ("授权代表", "Authorized Representative"),
        ("地址", "Address"),
        ("行政联系", "Administrative Contact"),
    ]
    pair_results = {cn: adjacent_pair(lines, cn, en, start, end) for cn, en in pair_specs}
    pairs_ok = all(pair_results.values())
    missing_pairs = [cn for cn, ok in pair_results.items() if not ok]
    field_labels = [title_cn, title_en, "统一识别号", "UEN", "授权代表", "Authorized Representative", "地址", "Address", "行政联系", "Administrative Contact"]
    widths = field_fill_widths(section, field_labels)

    if party == "A":
        # 甲方细则：每个字段后都带填写线、下划线/制表符下划线、长度一致、4cm-7cm。
        underline_ok = widths is not None
        length_ok = False
        if widths:
            in_range = all(SIGNATURE_FILL_LINE_MIN_TOL <= w <= SIGNATURE_FILL_LINE_MAX_TOL for w in widths)
            avg = sum(widths) / len(widths)
            consistent = avg > 0 and max(abs(w - avg) / avg for w in widths) <= 0.12
            length_ok = in_range and consistent
        ok = fields_present and pairs_ok and underline_ok and length_ok
        if ok:
            detail = "字段齐全，中英文上下相邻，每个字段后均有4-7cm且长度一致的填写线"
        else:
            detail = "字段、上下相邻排版或填写线长度(4-7cm/一致性)未全部满足"
            if missing_pairs:
                detail += f"；未成对相邻的字段：{missing_pairs}"
    else:
        # 乙方细则：字段齐全、每个字段后都存在填写横线、中英文上下相邻、填写横线长度一致
        # （不含4-7cm/类型的额外约束）。
        underline_ok = widths is not None and len(widths) == len(field_labels)
        length_ok = False
        if widths is not None and underline_ok:
            avg = sum(widths) / len(widths)
            length_ok = avg > 0 and max(abs(w - avg) / avg for w in widths) <= 0.12
        ok = fields_present and pairs_ok and underline_ok and length_ok
        if ok:
            detail = "字段齐全，中英文上下相邻，每个字段后均有长度一致的填写横线"
        else:
            reasons: list[str] = []
            if not fields_present:
                reasons.append("字段不齐全")
            if not pairs_ok:
                reasons.append(f"未成对相邻的字段：{missing_pairs}")
            if not underline_ok:
                reasons.append("存在字段后缺少填写横线")
            elif not length_ok:
                reasons.append("填写横线长度不一致")
            detail = "；".join(reasons) if reasons else "字段、上下相邻排版或填写横线长度一致性未全部满足"
    return CheckResult(name, 3, ok, detail)


def check_collective_text(info: PdfInfo) -> CheckResult:
    lines = info.lines_by_page[0]
    cn = find_first_line(lines, "甲方与乙方以下合称为“双方”。")
    en = find_first_line(lines, 'Party A and Party B are hereinafter collectively referred to as the "Parties".')
    # 位于甲乙方信息区下方：取乙方信息区最后一行（Administrative Contact）作为信息区下边界。
    party_info_bottom = max(
        (ln.y1 for ln in lines if "Administrative Contact" in ln.text or "行政联系" in ln.text),
        default=None,
    )
    below_info = bool(cn and party_info_bottom is not None and cn.y0 >= party_info_bottom)
    # 中文行在英文行上方 或 中文行紧邻英文行。
    order_ok = bool(cn and en and (en.y0 > cn.y0 or abs(en.y0 - cn.y0) <= 30))
    ok = bool(cn and en and below_info and order_ok)
    return CheckResult(
        "+1 首页双方合称说明文本",
        1,
        ok,
        "信息区下方检测到中英文合称说明，中文行在英文行上方或紧邻" if ok else "未检测到合规的双方合称说明中英文相邻文本（需位于甲乙方信息区下方）",
    )


def check_home_divider(info: PdfInfo) -> CheckResult:
    page = info.pages[0]
    lines = info.lines_by_page[0]
    party_end = next((ln.y1 for ln in lines if "Parties" in ln.text), page.rect.height * 0.45)
    clause_start = next((ln.y0 for ln in lines if "第一条" in ln.text or "Clause 1:" in ln.text), page.rect.height)
    ok = False
    for x0, x1, y, width, color in horizontal_segments(page):
        if not (party_end < y < clause_start):
            continue
        # 位于页面宽度 10% 至 90% 范围内：线的左端不晚于10%、右端不早于90%（±1%抽取容差）。
        if x0 <= page.rect.width * 0.11 and x1 >= page.rect.width * 0.89:
            # 浅灰色：RGB 各通道偏高且接近（灰）。color 为 None 时信息不足，按可接受处理。
            light_grey = color is None or (len(color) >= 3 and min(color[:3]) >= 0.70 and max(color[:3]) - min(color[:3]) <= 0.08)
            if light_grey:
                ok = True
                break
    return CheckResult(
        "+1 首页分隔线",
        1,
        ok,
        "双方信息区与第一条之间检测到位于页宽10%-90%的浅灰水平分隔线" if ok else "未在双方信息区与第一条之间检测到合规浅灰水平分隔线(需覆盖页宽10%-90%)",
    )


def locate_clause_titles(info: PdfInfo) -> dict[int, tuple[Optional[TextLine], Optional[TextLine]]]:
    locations: dict[int, tuple[Optional[TextLine], Optional[TextLine]]] = {}
    for n in range(1, 17):
        cn_line = None
        en_line = None
        for page_lines in info.lines_by_page:
            if cn_line is None:
                cn_line = find_first_line(page_lines, CN_TITLES[n])
            if en_line is None:
                en_line = find_first_line(page_lines, EN_TITLES[n])
        locations[n] = (cn_line, en_line)
    return locations


def table_after_title(info: PdfInfo, title_line: TextLine) -> Optional[TableBox]:
    tables = info.tables_by_page[title_line.page_index]
    below = [t for t in tables if t.rect.y0 >= title_line.y1 - 2]
    if not below:
        return None
    return sorted(below, key=lambda t: t.rect.y0 - title_line.y1)[0]


def table_text(info: PdfInfo, table: TableBox) -> str:
    lines = info.lines_by_page[table.page_index]
    selected = [ln.text for ln in lines if table.rect.y0 - 3 <= ln.y0 <= table.rect.y1 + 3 and table.rect.x0 - 3 <= ln.x0 <= table.rect.x1 + 3]
    return "\n".join(selected)


def clause_table_for(info: PdfInfo, n: int) -> Optional[TableBox]:
    _, en = locate_clause_titles(info)[n]
    if not en:
        return None
    table = table_after_title(info, en)
    if not table:
        return None
    text = table_text(info, table)
    return table if re.search(rf"\b{n}\.\d+\b", text) else None


def check_clause_titles(info: PdfInfo) -> CheckResult:
    locations = locate_clause_titles(info)
    bad: list[str] = []
    for n, (cn, en) in locations.items():
        # ① 中英文标题均保留
        if not cn or not en:
            bad.append(f"Clause {n} 标题缺失")
            continue
        # ② 中文标题位于英文标题上方（同页且中文在上）
        if cn.page_index != en.page_index or not (en.y0 > cn.y0):
            bad.append(f"Clause {n} 中文标题未位于英文标题上方")
            continue
        # ③ 标题与其下方条款表格保持在同一页
        table = table_after_title(info, en)
        if not table or table.page_index != en.page_index:
            bad.append(f"Clause {n} 标题与下方条款表格不在同一页")
            continue
        # ④ 同一条款的条款内容不可跨页：
        #    在 PDF 中，Word/WPS 排版把同一张表拆到相邻两页时，会生成两个分处不同页
        #    的 TableBox；单页 rect 高度永远落在页面范围内，因此不能靠 rect vs page_h
        #    判断跨页。改为在“该条款标题页之后、下一条款标题之前”的页面中，检查是否
        #    仍出现该条款的分项编号（n.x）或延续的表格。
        title_page = en.page_index
        next_title_page: int | None = None
        if n < 16:
            nxt_cn, nxt_en = locations[n + 1]
            candidates = [x.page_index for x in (nxt_cn, nxt_en) if x is not None]
            if candidates:
                next_title_page = min(candidates)
        pattern = re.compile(rf"(?<!\d){n}\.\d+\b")
        overflow_page: int | None = None
        for p_idx in range(title_page + 1, len(info.lines_by_page)):
            if next_title_page is not None and p_idx >= next_title_page:
                break
            page_text = "\n".join(ln.text for ln in info.lines_by_page[p_idx])
            if pattern.search(page_text):
                overflow_page = p_idx
                break
            for t in info.tables_by_page[p_idx]:
                if pattern.search(table_text(info, t)):
                    overflow_page = p_idx
                    break
            if overflow_page is not None:
                break
        if overflow_page is not None:
            bad.append(f"Clause {n} 条款内容跨页（延续至第{overflow_page + 1}页）")
    ok = not bad
    return CheckResult(
        "+5 第1条至第16条条款标题",
        5,
        ok,
        "Clause 1-16 中英文标题均保留，中文在英文上方，标题与其下方表格同页且条款内容未跨页" if ok else "; ".join(bad[:5]) + (" ..." if len(bad) > 5 else ""),
    )


def clause_tables(info: PdfInfo) -> dict[int, Optional[TableBox]]:
    return {n: clause_table_for(info, n) for n in range(1, 17)}


def check_table_structure(info: PdfInfo) -> CheckResult:
    tables = clause_tables(info)
    bad: list[str] = []
    for n, table in tables.items():
        # ① 使用表格排版；② 2列结构（存在中间竖分隔线）。
        #   clause_table_for 只在检测到矢量线框表格时返回 TableBox，文本框堆叠/图片替代
        #   均无法形成该表格结构，因而会在此处判为缺失，对应“不能用文本框堆叠或图片替代表格”。
        if table is None or table.divider_x is None:
            bad.append(f"Clause {n} 未检测到2列表格结构")
            continue
        mid = table.divider_x
        page_lines = info.lines_by_page[table.page_index]
        in_table = [ln for ln in page_lines if table.rect.y0 <= ln.y0 <= table.rect.y1]

        # ③ 左列表头“中文”，右列表头“English”。
        cn_header = next((ln for ln in in_table if ln.text.strip() == "中文" and ln.x0 < mid), None)
        en_header = next((ln for ln in in_table if ln.text.strip() == "English" and ln.x0 >= mid), None)
        if not (cn_header and en_header):
            bad.append(f"Clause {n} 左列表头非“中文”或右列表头非“English”")
            continue

        # ④ 表头下方为对应条款正文（表头之下才是正文）。
        header_bottom = max(cn_header.y1, en_header.y1)
        left_body = [ln for ln in in_table if ln.x0 < mid and ln.y0 >= header_bottom - 1]
        right_body = [ln for ln in in_table if ln.x0 >= mid and ln.y0 >= header_bottom - 1]

        left_text = "\n".join(ln.text for ln in left_body)
        right_text = "\n".join(ln.text for ln in right_body)

        # ⑤ 左列必须是该条款的中文正文，⑥ 右列必须是该条款的英文正文，且两列一一对应。
        # 逐条款编号（n.x）核对：左右两列都应出现该条款的分项编号，且集合一致。
        sub_pat = re.compile(rf"(?<!\d)({n}\.\d+)\b")
        left_subs = set(sub_pat.findall(left_text))
        right_subs = set(sub_pat.findall(right_text))
        # 语言主体判定：以中文字符数与英文单词数的相对占比衡量，避免个别专有名词/编号造成误判。
        left_cn = len(re.findall(r"[一-鿿]", left_text))
        left_en = len(re.findall(r"[A-Za-z]{3,}", left_text))
        right_cn = len(re.findall(r"[一-鿿]", right_text))
        right_en = len(re.findall(r"[A-Za-z]{3,}", right_text))

        reasons: list[str] = []
        # 左列存在中文条款正文，右列存在英文条款正文。
        if not (left_subs and left_cn > 0):
            reasons.append("左列缺少对应中文条款正文")
        if not (right_subs and right_en > 0):
            reasons.append("右列缺少对应英文条款正文")
        # 左右列条款编号必须对应一致。
        if left_subs and right_subs and left_subs != right_subs:
            reasons.append(f"左右列条款编号不对应(左{sorted(left_subs)}/右{sorted(right_subs)})")
        # 禁止左列英文主体：左列英文单词数明显多于中文字符即视为英文主体混入。
        if left_en > 0 and left_en >= left_cn:
            reasons.append("左列混入英文主体")
        # 禁止右列中文主体：右列中文字符数明显多于英文单词即视为中文主体混入。
        if right_cn > 0 and right_cn >= right_en:
            reasons.append("右列混入中文主体")
        if reasons:
            bad.append(f"Clause {n} " + "、".join(reasons))
    ok = not bad
    return CheckResult(
        "+5 条款内容双列表格结构",
        5,
        ok,
        "Clause 1-16 均为2列表格，左列表头“中文”/右列表头“English”，表头下方左中文右英文条款" if ok else "; ".join(bad[:5]) + (" ..." if len(bad) > 5 else ""),
    )


def check_signature_page(info: PdfInfo) -> CheckResult:
    last_page_index = info.doc.page_count - 1
    # 文档末尾：优先末页，允许倒数两页承载签署区。
    candidate_lines = [ln for i in range(max(0, last_page_index - 1), last_page_index + 1) for ln in info.lines_by_page[i]]
    text = "\n".join(ln.text for ln in candidate_lines)
    # ① 出现“签署页”和“Signature Page”；② 保留 Party A/B、授权代表签字、姓名、日期及对应中文字段。
    title_present = "签署页" in text and "Signature Page" in text
    field_present = all(
        item in text
        for item in [
            "甲方 / 服务提供方",
            "Party A / Service Provider",
            "乙方 / 客户方",
            "Party B / Client",
            "授权代表签字",
            "Authorized Representative Signature",
            "姓名",
            "Name",
            "日期",
            "Date",
        ]
    )
    fields_present = title_present and field_present

    # ③ 中文字段与英文字段按“一行中文一行英文”上下对应排列。
    pairs = [
        ("甲方 / 服务提供方", "Party A / Service Provider"),
        ("授权代表签字", "Authorized Representative Signature"),
        ("姓名", "Name"),
        ("日期", "Date"),
        ("乙方 / 客户方", "Party B / Client"),
    ]
    pair_ok = True
    for cn_kw, en_kw in pairs:
        cn_lines = [ln for ln in candidate_lines if cn_kw in ln.text]
        en_lines = [ln for ln in candidate_lines if en_kw in ln.text]
        if not any(cn.page_index == en.page_index and 0 < en.y0 - cn.y0 <= 28 for cn in cn_lines for en in en_lines):
            pair_ok = False
            break

    # 签署字段行（每个字段中/英各一行，共 8 对 = 16 行）。
    field_labels = [
        "甲方 / 服务提供方",
        "Party A / Service Provider",
        "乙方 / 客户方",
        "Party B / Client",
        "授权代表签字",
        "Authorized Representative Signature",
        "姓名",
        "Name",
        "日期",
        "Date",
    ]
    expected_signature_lines = [ln for ln in candidate_lines if any(item in ln.text for item in field_labels)]
    signature_field_count_ok = len(expected_signature_lines) >= 16

    # ④ 文字后面都带填写横线，⑤ 为下划线或制表符下划线，⑥ 长度4cm-7cm。
    field_widths: list[float] = []
    underline_ok = bool(expected_signature_lines)
    for ln in expected_signature_lines:
        widths = measured_underline_widths(ln)  # 基于 [_\t]{6,} 的下划线/制表符下划线
        if not widths:
            underline_ok = False
            break
        field_widths.append(max(widths))
    length_range_ok = underline_ok and all(SIGNATURE_FILL_LINE_MIN_TOL <= w <= SIGNATURE_FILL_LINE_MAX_TOL for w in field_widths)
    # ⑦ 填写横线长度一致。
    consistency_ok = False
    if field_widths:
        avg = sum(field_widths) / len(field_widths)
        consistency_ok = avg > 0 and max(abs(w - avg) / avg for w in field_widths) <= 0.12

    ok = fields_present and pair_ok and signature_field_count_ok and underline_ok and length_range_ok and consistency_ok
    return CheckResult(
        "+5 Signature Page签署页",
        5,
        ok,
        "文档末尾检测到签署页、Party A/B等字段、中英文上下对应，字段后填写线为下划线且长度一致(4-7cm)" if ok else "签署页标题、字段、中英文一行中一行英对应或填写线(类型/4-7cm/一致性)未全部满足",
    )


def check_footer_page_numbers(info: PdfInfo) -> CheckResult:
    bad: list[str] = []
    # 细则明确将页码字体限定为 Calibri 或微软雅黑，不把 Helvetica/Arial 等通用基础字体
    # 当作别名，否则用错字体的文档会被误判合格。
    acceptable_fonts = ["calibri", "microsoft yahei", "微软雅黑", "msyh"]
    seen_numbers: dict[str, int] = {}
    for idx, page in enumerate(info.pages):
        expected = str(idx + 1)
        # 全页所有等于本页序号的文本（用于判定唯一性与位置）。
        candidates = [ln for ln in info.lines_by_page[idx] if ln.text.strip() == expected]
        # 本页出现的任意“纯数字”文本，用于检测跳号/错号/多个页码。
        numeric_lines = [ln for ln in info.lines_by_page[idx] if re.fullmatch(r"\d+", ln.text.strip())]
        ok_line = None
        for ln in candidates:
            # 页面底部居中（页脚），排除页眉、正文中间、左下/右下角。
            center_ok = abs((ln.x0 + ln.x1) / 2 - page.rect.width / 2) <= page.rect.width * 0.04
            bottom_ok = ln.y0 >= page.rect.height * 0.90
            black_ok = ln.is_blackish()
            font = ln.main_font().lower()
            font_ok = any(name in font for name in acceptable_fonts)
            if center_ok and bottom_ok and black_ok and font_ok:
                ok_line = ln
                break
        if ok_line is None:
            detail = "未检测到底部居中且字体为Calibri/微软雅黑的黑色页码（含页眉/正文中间/左下角/右下角均不合规）"
            if candidates:
                fonts = ", ".join(sorted({ln.main_font() for ln in candidates}))
                detail += f"（候选字体：{fonts}）"
            bad.append(f"第{idx + 1}页{detail}")
            continue
        # 连续编号、无重复号：该页序号数字不应重复出现，且不应出现其它页序号数字（跳号/错号）。
        wrong_numbers = [ln.text.strip() for ln in numeric_lines if ln.text.strip() != expected]
        if wrong_numbers:
            bad.append(f"第{idx + 1}页出现非本页序号的页码数字：{sorted(set(wrong_numbers))}")
            continue
        if len([ln for ln in numeric_lines if ln.text.strip() == expected]) != 1:
            bad.append(f"第{idx + 1}页页码“{expected}”重复出现")
            continue
        seen_numbers[expected] = seen_numbers.get(expected, 0) + 1

    # 从第1页开始连续编号：1..N 每个都恰好命中一次。
    if not bad:
        for n in range(1, info.doc.page_count + 1):
            if seen_numbers.get(str(n), 0) != 1:
                bad.append(f"页码编号不连续或缺失：缺少第{n}页页码")
                break

    ok = not bad
    return CheckResult(
        "+1 页面页脚页码",
        1,
        ok,
        "每页页脚底部水平居中检测到从1开始连续编号的黑色页码，字体为Calibri/微软雅黑，无跳号/重复号/漏号" if ok else "; ".join(bad[:3]) + (" ..." if len(bad) > 3 else ""),
    )


def check_table_width(info: PdfInfo) -> CheckResult:
    tables = clause_tables(info)
    valid_tables = [table for table in tables.values() if table is not None]
    bad: list[str] = []
    if not valid_tables:
        return CheckResult("+3 表格宽度", 3, False, "未检测到可测量的条款表格")

    # 页面正文宽度（文本区宽度）由稳定的“页面边距”估计，而非全文对象的左右极值——
    # 后者会被过宽表格/图片/水印等对象“撑大”，导致本应超出正文边距的表格反被判为合格。
    # 稳定估计做法（不引入 COM）：
    #   ① 只用页面 body 的正文文本行（排除数字页码），每页统计 x0/x1；
    #   ② 用众数式的稳健估计（每页的最小 x0 与最大 x1 的“中位数”）作为左右页面边距；
    #      中位数天然抗离群，避免个别过宽/过窄行影响。
    # 该估计等价于对全文各页的“正文文本区”取中位边界，能正确处理不对称边距。
    per_page_left: list[float] = []
    per_page_right: list[float] = []
    for idx in range(info.doc.page_count):
        page_lines = [ln for ln in info.lines_by_page[idx] if not re.fullmatch(r"\d+", ln.text.strip())]
        if not page_lines:
            continue
        per_page_left.append(min(ln.x0 for ln in page_lines))
        per_page_right.append(max(ln.x1 for ln in page_lines))
    if not per_page_left or not per_page_right:
        return CheckResult("+3 表格宽度", 3, False, "未检测到可测量的正文文本区")
    left_margin = statistics.median(per_page_left)
    content_right = statistics.median(per_page_right)
    body_width = content_right - left_margin
    loc = locate_clause_titles(info)

    for n, table in tables.items():
        cn, en = loc[n]
        if not table or not cn:
            bad.append(f"Clause {n} 缺少可测量表格")
            continue
        page = info.pages[table.page_index]
        # ① 表格宽度为正文宽度的 95%-100%（rubric上限即 100%，仅给极小抽取容差 0.5%）。
        width_ratio = table.width / body_width if body_width > 0 else 0
        width_ok = 0.95 <= width_ratio <= 1.005
        # ② 表格左边界与标题左边界对齐。
        left_align = abs(table.rect.x0 - cn.x0) <= 3
        # ③ 右边界不超出页面边距：不越过正文文本区右边界，且不越出页面。
        right_inside = table.rect.x1 <= content_right + 1 and table.rect.x1 <= page.rect.width
        if not (width_ok and left_align and right_inside):
            reasons: list[str] = []
            if not width_ok:
                reasons.append(f"宽度占比{width_ratio:.1%}(要求95%-100%)")
            if not left_align:
                reasons.append("左边界未与标题对齐")
            if not right_inside:
                reasons.append("右边界超出正文边距")
            bad.append(f"Clause {n} 表格{('、'.join(reasons))}")
    ok = not bad
    return CheckResult(
        "+3 表格宽度",
        3,
        ok,
        "所有条款表格宽度为正文宽度95%-100%，左边界与标题对齐且右边界未超出页面边距" if ok else "; ".join(bad[:5]) + (" ..." if len(bad) > 5 else ""),
    )


def signature_missing_party_penalty(info: PdfInfo) -> CheckResult:
    # -3 细则：签署页缺失 Party A 或 Party B 任意一方的签署字段即扣分。
    # 逐方在签署页上定位其区块，核对该方的签署字段是否齐全：
    #   方别标识（甲方 / 服务提供方、Party A / Service Provider；乙方 / 客户方、Party B / Client）、
    #   授权代表签字 / Authorized Representative Signature、姓名 / Name、日期 / Date。
    # 任一方缺少上述任一签署字段即触发扣分。不额外约束填写线长度、对齐等（属于其它项）。

    # 定位签署页：取含“签署页”和“Signature Page”的页（办公软件导出通常在末页）；找不到则用末页。
    sig_idx = None
    for idx in range(info.doc.page_count):
        if "签署页" in info.text_by_page[idx] and "Signature Page" in info.text_by_page[idx]:
            sig_idx = idx
    if sig_idx is None:
        sig_idx = info.doc.page_count - 1

    lines = sorted(info.lines_by_page[sig_idx], key=lambda ln: (ln.y0, ln.x0))
    a_start = next((ln.y0 for ln in lines if "甲方 / 服务提供方" in ln.text), None)
    b_start = next((ln.y0 for ln in lines if "乙方 / 客户方" in ln.text), None)

    def block_has_fields(start_y: Optional[float], end_y: float, id_cn: str, id_en: str) -> bool:
        if start_y is None:
            return False
        block_text = "\n".join(ln.text for ln in lines if start_y <= ln.y0 < end_y)
        required = [
            id_cn,
            id_en,
            "授权代表签字",
            "Authorized Representative Signature",
            "姓名",
            "Name",
            "日期",
            "Date",
        ]
        return all(field in block_text for field in required)

    b_end = 10**9
    a_end = b_start if b_start is not None else b_end
    party_a = block_has_fields(a_start, a_end, "甲方 / 服务提供方", "Party A / Service Provider")
    party_b = block_has_fields(b_start, b_end, "乙方 / 客户方", "Party B / Client")

    fail = not (party_a and party_b)
    return CheckResult(
        "-3 Signature Page签署页缺失Party A或Party B任意一方签署字段",
        -3,
        fail,
        "签署页缺失Party A或Party B的签署字段" if fail else "签署页Party A与Party B签署字段均齐全，不扣分",
    )


def table_header_or_width_penalty(info: PdfInfo) -> CheckResult:
    # -3 细则：条款表格没有出现“中文”和“English”两个表头，
    #          或 左右两列宽度明显不一致（判定标准：任意一列宽度小于页面正文宽度的35%），
    #          任一情况即扣分。
    # 说明：
    #   ① 表头判定要求“中文”在左列 header 区、“English”在右列 header 区，
    #      仅在整表文本中同时出现“中文/English”不足以判为合格（可能出现在正文里）。
    #   ② 正文宽度改为使用真实左右正文边界，避免左右边距不对称时按“页宽-2*左边距”估算偏差。
    #      与 check_table_width 一致：按各页非页码正文行的最小 x0/最大 x1 的中位数估计左右边界。
    per_page_left: list[float] = []
    per_page_right: list[float] = []
    for idx in range(info.doc.page_count):
        page_lines = [ln for ln in info.lines_by_page[idx] if not re.fullmatch(r"\d+", ln.text.strip())]
        if not page_lines:
            continue
        per_page_left.append(min(ln.x0 for ln in page_lines))
        per_page_right.append(max(ln.x1 for ln in page_lines))
    if per_page_left and per_page_right:
        body_width_global = statistics.median(per_page_right) - statistics.median(per_page_left)
    else:
        body_width_global = 0.0

    bad: list[str] = []
    for n, table in clause_tables(info).items():
        if not table:
            bad.append(f"Clause {n} 表格缺失")
            continue
        # ① 表头：中文/English 必须位于表头行，且中文在左列、English在右列。
        header_ok = False
        if table.divider_x is not None:
            mid = table.divider_x
            page_lines = info.lines_by_page[table.page_index]
            in_table = [ln for ln in page_lines if table.rect.y0 <= ln.y0 <= table.rect.y1]
            # 表头行：位于表格顶部若干行内（取最靠上的一行 y0 起 20pt 内为 header 区），
            # 该 header 区中，左列必须是“中文”，右列必须是“English”。
            if in_table:
                top_y = min(ln.y0 for ln in in_table)
                header_zone = [ln for ln in in_table if ln.y0 - top_y <= 20]
                left_header = any(ln.text.strip() == "中文" and ln.x0 < mid and ln.x1 <= mid + 2 for ln in header_zone)
                right_header = any(ln.text.strip() == "English" and ln.x0 >= mid - 2 for ln in header_zone)
                header_ok = left_header and right_header

        # ② 列宽：左右两列宽度均不得小于页面正文宽度的35%。
        threshold = body_width_global * 0.35
        left_w, right_w = table.left_width(), table.right_width()
        width_ok = bool(body_width_global > 0 and left_w and right_w and left_w >= threshold and right_w >= threshold)
        if not header_ok:
            bad.append(f"Clause {n} 缺少中文(左)/English(右)表头")
        elif not width_ok:
            bad.append(f"Clause {n} 左右列宽有一列小于正文宽度35%")
    fail = bool(bad)
    return CheckResult(
        "-3 条款表格没有出现“中文”和“English”两个表头或左右两列宽度明显不一致",
        -3,
        fail,
        "; ".join(bad[:5]) + (" ..." if len(bad) > 5 else "") if fail else "所有条款表格均含“中文”/“English”表头且两列宽度均不小于正文宽度35%，不扣分",
    )


def numbering_mismatch_penalty(info: PdfInfo) -> CheckResult:
    # -1 细则：任意三个以上条款编号出现中英文不对应即扣分。
    # 逐条款比较：中文一侧出现的分项编号集合（如 3.1、3.2）与英文一侧出现的分项编号集合
    # 是否一致；不一致即记为该条款“中英文编号不对应”。统计不对应的条款数，达到 3 个（“三个
    # 以上”含本数）即扣分。
    #
    # 针对办公软件导出 PDF：优先用检测到的双列表格中缝竖线划分左右两列；若某条款未识别出
    # 表格竖线，则回退到用该条款标题下方区域按页面水平中点划分左右列，仍照常比较编号，
    # 而不是把“未识别到表格”本身当作编号不对应。
    loc = locate_clause_titles(info)
    mismatches: list[int] = []
    for n in range(1, 17):
        table = clause_table_for(info, n)
        cn_title, en_title = loc[n]
        # 确定该条款正文所在页与左右分界 x。
        if table is not None and table.divider_x is not None:
            page_index = table.page_index
            top, bottom = table.rect.y0, table.rect.y1
            divider_x = table.divider_x
        elif en_title is not None:
            page_index = en_title.page_index
            top = min(cn_title.y1 if cn_title else en_title.y1, en_title.y1)
            # 下界：下一条款标题或页面底部。
            next_titles = [
                t.y0
                for m in range(n + 1, 17)
                for t in loc[m]
                if t is not None and t.page_index == page_index and t.y0 > top
            ]
            bottom = min(next_titles) if next_titles else info.pages[page_index].rect.height
            divider_x = info.pages[page_index].rect.width / 2
        else:
            mismatches.append(n)
            continue

        left_nums: set[str] = set()
        right_nums: set[str] = set()
        for ln in info.lines_by_page[page_index]:
            if not (top <= ln.y0 <= bottom):
                continue
            match = re.match(rf"\s*({n}\.\d+)\b", ln.text)
            if not match:
                continue
            if ln.x0 < divider_x:
                left_nums.add(match.group(1))
            else:
                right_nums.add(match.group(1))
        if left_nums != right_nums:
            mismatches.append(n)
    fail = len(mismatches) >= 3
    return CheckResult(
        "-1 任意三个以上条款编号出现中英文不对应",
        -1,
        fail,
        f"编号中英文不对应的条款达到3个及以上：{mismatches}" if fail else f"编号不对应条款数为{len(mismatches)}，不扣分",
    )


def clipped_text_penalty(info: PdfInfo) -> CheckResult:
    # -3 细则：任意条款表格中文字超出单元格边界，或被截断到无法阅读，即扣分。
    # 两个点：
    #   ① 文字超出单元格边界——每个单元格的横向范围为 [左边界, 中缝] 或 [中缝, 右边界]，
    #      单元格内任一文字行的左右端超出其所属单元格边界即判超界；
    #   ② 被截断到无法阅读——出现乱码/缺字替代字符（� / □）视为不可阅读的截断。
    # 检查对象为条款表格内的所有正文行（不仅是带编号的首行），因为折行续行同样可能溢出。
    # 针对办公软件导出 PDF 的抽取误差保留 2pt 容差；不额外附加细则未要求的约束。
    TOL = 2.0
    bad: list[str] = []
    for n, table in clause_tables(info).items():
        if not table or table.divider_x is None:
            continue
        left_x0, right_x1, mid = table.rect.x0, table.rect.x1, table.divider_x
        overflow = False
        for ln in info.lines_by_page[table.page_index]:
            if not (table.rect.y0 <= ln.y0 <= table.rect.y1):
                continue
            # 按行的中心归入所属单元格列。
            center = (ln.x0 + ln.x1) / 2
            if center < mid:
                # 左单元格：不得越过左边界，也不得越过中缝进入右列。
                if ln.x0 < left_x0 - TOL or ln.x1 > mid + TOL:
                    overflow = True
                    bad.append(f"Clause {n} 左列文字超出单元格边界")
                    break
            else:
                # 右单元格：不得越过中缝，也不得越过表格右边界。
                if ln.x0 < mid - TOL or ln.x1 > right_x1 + TOL:
                    overflow = True
                    bad.append(f"Clause {n} 右列文字超出单元格边界")
                    break
        if overflow:
            continue
    # 乱码/缺字替代字符视为“被截断到无法阅读”。
    if "�" in info.all_text or "□" in info.all_text:
        bad.append("表格文字出现乱码/缺字替代字符（无法阅读）")
    fail = bool(bad)
    return CheckResult(
        "-3 任意条款表格中文字超出单元格边界或被截断到无法阅读",
        -3,
        fail,
        "; ".join(bad[:5]) if fail else "未发现条款表格文字超出单元格边界或被截断到无法阅读，不扣分",
    )


def blank_or_overlap_penalty(info: PdfInfo) -> CheckResult:
    # -3 细则：任意页面出现「超过页面高度50%的无内容空白区域」或「文本重叠情况」即扣分。
    # 两个点：
    #   ① 无内容空白区域 > 页面高度的50%——在页面纵向上，任意相邻内容之间（或首/末内容与
    #      页边之间）的最大连续空白高度占页高比例是否超过50%。内容包含正文行与表格线等绘图。
    #   ② 文本重叠——同一页上任意两条文本行的包围盒发生实质性重叠（交叠面积超过较小盒的50%），
    #      即文字压字、无法正常阅读。左右分栏（不同列）在水平方向分开，不会计为重叠。
    # 针对办公软件导出 PDF 的抽取误差：相邻行包围盒常有 1-2pt 轻微相接，故要求“实质性”重叠
    # （>50%）而非任何接触，避免误报；不附加细则未要求的其它约束。
    bad: list[str] = []
    for idx, page in enumerate(info.pages):
        page_h = page.rect.height
        lines = info.lines_by_page[idx]

        # ① 最大连续空白高度占比。
        blank_ratio = max_vertical_blank_ratio(page, lines)
        blank_over = blank_ratio > 0.50

        # ② 实质性文本重叠：任意两行交叠面积 > 较小盒面积的 50%。
        overlap = False
        content = [ln for ln in lines if not ln.text.strip().isdigit()]
        content.sort(key=lambda ln: (ln.y0, ln.x0))
        for i, a in enumerate(content):
            for b in content[i + 1 :]:
                if b.y0 > a.y1:  # 已无纵向交叠可能（按 y 排序）
                    break
                inter = intersection_area(a.bbox, b.bbox)
                if inter <= 0:
                    continue
                denom = min(rect_area(a.bbox), rect_area(b.bbox))
                if denom > 0 and inter / denom > 0.50:
                    overlap = True
                    break
            if overlap:
                break

        if blank_over or overlap:
            msg = f"第{idx + 1}页"
            if blank_over:
                msg += f"存在超过页高50%的连续空白({blank_ratio:.0%})"
            if overlap:
                msg += "存在文本重叠"
            bad.append(msg)
    fail = bool(bad)
    return CheckResult(
        "-3 任意页面出现超过页面高度50%的无内容空白区域或文本重叠情况",
        -3,
        fail,
        "; ".join(bad) if fail else "未发现超过页高50%的连续空白区域或文本重叠，不扣分",
    )


def evaluate_dimension2(info: PdfInfo) -> list[CheckResult]:
    cn_title, en_title, title_penalty = check_home_titles(info)
    positive_checks = [
        cn_title,
        en_title,
        check_signing_intro(info),
        check_party_section(info, "A"),
        check_party_section(info, "B"),
        check_collective_text(info),
        check_home_divider(info),
        check_clause_titles(info),
        check_table_structure(info),
        check_signature_page(info),
        check_footer_page_numbers(info),
        check_table_width(info),
    ]
    penalty_checks = [
        title_penalty,
        signature_missing_party_penalty(info),
        table_header_or_width_penalty(info),
        numbering_mismatch_penalty(info),
        clipped_text_penalty(info),
        blank_or_overlap_penalty(info),
    ]
    return positive_checks + penalty_checks


def _build_dim2_items(checks: list[CheckResult]) -> list[dict]:
    items: list[dict] = []
    for check in checks:
        # Strip the leading score prefix from name (e.g. "+1 首页中文标题" -> "首页中文标题").
        rule = re.sub(r"^[+\-]\d+\s+", "", check.name)
        items.append(
            {
                "rule": rule,
                "max_delta": check.score,
                "delta": check.score if check.passed else 0,
                "hit": check.passed,
                "detail": "",
            }
        )
    return items


def _locate_pdf(dir_path: str) -> Optional[str]:
    """在给定目录内定位待评估的PDF文件；优先使用默认文件名，其次扫描目录内任意PDF。"""
    default_path = os.path.join(dir_path, PDF_NAME)
    if os.path.isfile(default_path):
        return default_path
    if os.path.isdir(dir_path):
        for name in sorted(os.listdir(dir_path)):
            if name.lower().endswith(".pdf"):
                return os.path.join(dir_path, name)
    return None


def evaluate(dir_path: str) -> dict:
    """统一入口：接收脚本所在目录路径，在该目录内定位并评估PDF，返回结构化结果字典。"""
    result: dict = {
        "id": "034",
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
        pdf_path = _locate_pdf(dir_path)
        if not pdf_path or not os.path.isfile(pdf_path):
            result["status"] = "error"
            result["error"] = f"未在目录中找到待评估的PDF文件：{dir_path}"
            result["file_name"] = PDF_NAME
            return result

        result["file_name"] = os.path.basename(pdf_path)

        try:
            info = load_pdf(pdf_path)
        except Exception as exc:
            result["status"] = "error"
            result["error"] = f"PDF无法正常打开或解析：{exc}"
            return result

        dim1_ok, dim1_failures, _ = dimension1(info)
        result["dim1_pass"] = dim1_ok
        result["dim1_reason"] = "" if dim1_ok else "；".join(dim1_failures)

        if dim1_ok:
            checks = evaluate_dimension2(info)
            items = _build_dim2_items(checks)
            result["dim2_items"] = items
            result["total_score"] = sum(item["delta"] for item in items)
            result["max_score"] = sum(item["max_delta"] for item in items if item["max_delta"] > 0)
        else:
            # 维度一未通过：按规范返回空的 dim2_items 与 0 总分/满分（见文档 §2.2）。
            result["dim2_items"] = []
            result["total_score"] = 0
            result["max_score"] = 0
    except Exception as exc:  # pragma: no cover - 兜底保护
        result["status"] = "error"
        result["error"] = f"脚本内部异常：{exc}"

    return result


if __name__ == "__main__":
    _dir = sys.argv[1] if len(sys.argv) > 1 else os.path.dirname(os.path.abspath(__file__))
    print(json.dumps(evaluate(_dir), ensure_ascii=False, indent=2))
