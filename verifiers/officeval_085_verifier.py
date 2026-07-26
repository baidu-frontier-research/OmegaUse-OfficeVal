# -*- coding: utf-8 -*-
"""自动评估 S11 可编辑图表工作簿。

对外统一入口 ``evaluate(dir_path)``：接收脚本所在目录的路径，脚本自己在该目录里
定位并打开以下三个目标文档：
  1_美化版_S11可编辑图表.xlsx
  2_美化版_S11可编辑图表.xlsx
  3_美化版_S11可编辑图表.xlsx

实现原则：只使用 Python 标准库解析 .xlsx/.xlsm 的 OOXML，不依赖 Excel、openpyxl 或人工判断。
"""
from __future__ import annotations

import json
import math
import posixpath
import re
import sys
import zipfile
from dataclasses import dataclass, field, asdict
from pathlib import Path
from typing import Any, cast
import xml.etree.ElementTree as ET

NS = {
    "x": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "c": "http://schemas.openxmlformats.org/drawingml/2006/chart",
    "xdr": "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
}

TARGET_FILES = [
    "1_美化版_S11可编辑图表.xlsx",
    "2_美化版_S11可编辑图表.xlsx",
    "3_美化版_S11可编辑图表.xlsx",
]
EMU_PER_CM = 360000.0
EMU_PER_PT = 12700.0


def q(ns: str, tag: str) -> str:
    return f"{{{NS[ns]}}}{tag}"


@dataclass
class RuleResult:
    rule_id: str
    name: str
    points: int
    matched: bool
    reason: str
    evidence: dict[str, Any] = field(default_factory=dict)
    file: str | None = None
    kind: str = "positive"


@dataclass
class DataRegion:
    header_row: int = 1
    freq_col: str | None = None
    s11_col: str | None = None
    ref_col: str | None = None
    start_row: int = 2
    end_row: int = 1
    freq_values: list[float] = field(default_factory=list)
    s11_values: list[float] = field(default_factory=list)
    ref_values: list[float] = field(default_factory=list)
    headers: dict[str, str] = field(default_factory=dict)
    non_numeric_count: int = 0
    blank_count: int = 0

    @property
    def valid_rows(self) -> int:
        return min(len(self.freq_values), len(self.s11_values))

    @property
    def max_col_index(self) -> int:
        cols = [col_to_num(c) for c in (self.freq_col, self.s11_col, self.ref_col) if c]
        return max(cols) if cols else 0


@dataclass
class Anchor:
    kind: str
    col0: int | None = None
    row0: int | None = None
    col1: int | None = None
    row1: int | None = None
    cx_emu: int = 0
    cy_emu: int = 0

    @property
    def width_cm(self) -> float:
        return self.cx_emu / EMU_PER_CM if self.cx_emu else 0.0

    @property
    def height_cm(self) -> float:
        return self.cy_emu / EMU_PER_CM if self.cy_emu else 0.0


@dataclass
class SeriesModel:
    name: str | None = None
    x_values: list[float] = field(default_factory=list)
    y_values: list[float] = field(default_factory=list)
    x_ref: str | None = None
    y_ref: str | None = None
    line: dict[str, Any] = field(default_factory=dict)
    marker: dict[str, Any] = field(default_factory=dict)
    smooth: bool = False
    # 数值缓存（c:val/c:yVal）中实际写出的数据点 idx 序列与声明的总点数 ptCount，
    # 用于判断折线是否因空值/缺失点而断裂（idx 不连续或缺点即断裂）。
    y_indices: list[int] = field(default_factory=list)
    y_pt_count: int | None = None
    # 单个数据点（c:dPt）对标记的覆盖：每项 {idx, has_marker, symbol}。
    # 用于判断是否存在按点隐藏/改形标记（如某点 marker symbol=none 或改成非圆形）。
    dpt_markers: list[dict[str, Any]] = field(default_factory=list)


@dataclass
class AxisModel:
    tag: str
    pos: str | None = None
    title: str | None = None
    orientation: str | None = None
    min_val: float | None = None
    max_val: float | None = None
    major_unit: float | None = None
    num_fmt: str | None = None
    line: dict[str, Any] = field(default_factory=dict)
    font_sizes: list[float] = field(default_factory=list)
    font_names: list[str] = field(default_factory=list)
    font_colors: list[str] = field(default_factory=list)
    title_rotation: int | None = None
    # 轴标题 a:bodyPr/@vert 竖排属性值（如 "vert"/"vert270"/"eaVert"/"wordArtVert" 等表示竖向）；
    # 未写出时为 None，此时是否竖向需要结合 rot 判断。
    title_vert: str | None = None
    # 轴标题（c:title/c:layout/c:manualLayout）显式布局，用于判断"底部居中"等位置证据。
    # 空字典表示 OOXML 未写出手动布局（Excel/WPS 默认自动布局），交由具体检查决定如何处置。
    title_layout: dict[str, Any] = field(default_factory=dict)


@dataclass
class ChartModel:
    path: str
    chart_type: str | None = None
    title: str | None = None
    series: list[SeriesModel] = field(default_factory=list)
    axes: list[AxisModel] = field(default_factory=list)
    has_legend: bool = False
    anchor: Anchor | None = None
    chart_line: dict[str, Any] = field(default_factory=dict)
    plot_line: dict[str, Any] = field(default_factory=dict)
    plot_fill: str | None = None
    plot_layout: dict[str, Any] = field(default_factory=dict)
    font_sizes: list[float] = field(default_factory=list)
    font_names: list[str] = field(default_factory=list)
    font_colors: list[str] = field(default_factory=list)
    # 仅取自图表标题 c:title 节点的字体/字号/颜色，避免混入坐标轴/图例等其他文本，
    # 供 P12 判断标题本身的字体格式（与上面覆盖整个 chart XML 的三个字段区分开）。
    title_font_sizes: list[float] = field(default_factory=list)
    title_font_names: list[str] = field(default_factory=list)
    title_font_colors: list[str] = field(default_factory=list)
    wordart_like: bool = False
    scatter_style: str | None = None


@dataclass
class SheetModel:
    path: str
    cells: dict[str, Any] = field(default_factory=dict)
    row_heights: dict[int, float] = field(default_factory=dict)
    col_widths: dict[int, float] = field(default_factory=dict)
    data: DataRegion = field(default_factory=DataRegion)


@dataclass
class WorkbookModel:
    path: Path
    ok_zip: bool = False
    suffix_ok: bool = False
    package: dict[str, bytes] = field(default_factory=dict, repr=False)
    shared_strings: list[str] = field(default_factory=list)
    sheet: SheetModel | None = None
    charts: list[ChartModel] = field(default_factory=list)
    media_files: list[str] = field(default_factory=list)
    image_anchors: list[Anchor] = field(default_factory=list)
    parse_errors: list[str] = field(default_factory=list)


@dataclass
class FileEvaluation:
    file: str
    dimension1_passed: bool
    gate_results: list[RuleResult]
    positive_results: list[RuleResult] = field(default_factory=list)
    positive_score: int = 0
    note: str = ""


# ───────────────────────────── 基础 XML / OOXML 解析 ─────────────────────────────

def parse_xml(data: bytes) -> ET.Element:
    text = data.decode("utf-8-sig", errors="replace")
    return ET.fromstring(text)


def open_package(path: Path) -> tuple[bool, dict[str, bytes], list[str]]:
    errors: list[str] = []
    try:
        with zipfile.ZipFile(path) as zf:
            return True, {name: zf.read(name) for name in zf.namelist()}, errors
    except Exception as exc:  # noqa: BLE001 - report exact open failure
        errors.append(str(exc))
        return False, {}, errors


def resolve_target(base_part: str, target: str) -> str:
    target = target.replace("\\", "/")
    if target.startswith("/"):
        return target.lstrip("/")
    base_dir = posixpath.dirname(base_part)
    return posixpath.normpath(posixpath.join(base_dir, target)).lstrip("/")


def read_rels(pkg: dict[str, bytes], part: str) -> dict[str, str]:
    rels_path = posixpath.join(posixpath.dirname(part), "_rels", posixpath.basename(part) + ".rels")
    rels_path = rels_path.lstrip("/")
    if rels_path not in pkg:
        return {}
    root = parse_xml(pkg[rels_path])
    rels: dict[str, str] = {}
    for rel in root.findall("rel:Relationship", NS):
        rid = rel.get("Id")
        target = rel.get("Target")
        if rid and target:
            rels[rid] = resolve_target(part, target)
    return rels


def extract_text(el: ET.Element | None) -> str:
    if el is None:
        return ""
    texts: list[str] = []
    for t in el.findall(".//a:t", NS):
        if t.text:
            texts.append(t.text)
    # chart series title may use c:v directly
    if not texts:
        for v in el.findall(".//c:v", NS):
            if v.text:
                texts.append(v.text)
    return "".join(texts)


def parse_shared_strings(pkg: dict[str, bytes]) -> list[str]:
    if "xl/sharedStrings.xml" not in pkg:
        return []
    root = parse_xml(pkg["xl/sharedStrings.xml"])
    values: list[str] = []
    for si in root.findall("x:si", NS):
        txt = "".join(t.text or "" for t in si.findall(".//x:t", NS))
        values.append(txt)
    return values


def parse_cell_value(c: ET.Element, shared_strings: list[str]) -> Any:
    t = c.get("t", "n")
    if t == "inlineStr":
        return "".join(x.text or "" for x in c.findall(".//x:t", NS))
    v = c.find("x:v", NS)
    if v is None or v.text is None:
        return None
    raw = v.text
    if t == "s":
        try:
            idx = int(raw)
            return shared_strings[idx] if 0 <= idx < len(shared_strings) else raw
        except ValueError:
            return raw
    if t in {"str", "e"}:
        return raw
    try:
        return float(raw)
    except ValueError:
        return raw


def col_to_num(col: str | None) -> int:
    if not col:
        return 0
    n = 0
    for ch in col.upper():
        if "A" <= ch <= "Z":
            n = n * 26 + ord(ch) - 64
    return n


def num_to_col(n: int) -> str:
    out = ""
    while n:
        n, r = divmod(n - 1, 26)
        out = chr(65 + r) + out
    return out


def split_ref(ref: str) -> tuple[str, int]:
    m = re.match(r"([A-Za-z]+)(\d+)$", ref)
    if not m:
        return "", 0
    return m.group(1).upper(), int(m.group(2))


def numeric(v: Any) -> float | None:
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        if math.isfinite(float(v)):
            return float(v)
        return None
    if isinstance(v, str):
        try:
            x = float(v.strip())
            return x if math.isfinite(x) else None
        except ValueError:
            return None
    return None


def normalize_text(s: Any) -> str:
    s = "" if s is None else str(s)
    repl = {"（": "(", "）": ")", "₁": "1", "₁₁": "11", "－": "-", " ": ""}
    for k, v in repl.items():
        s = s.replace(k, v)
    return s.strip().lower()


def infer_data_region(cells: dict[str, Any]) -> DataRegion:
    headers: dict[str, str] = {}
    for ref, value in cells.items():
        col, row = split_ref(ref)
        if row == 1:
            headers[col] = "" if value is None else str(value)

    freq_col = None
    s11_col = None
    ref_col = None
    for col, text in headers.items():
        nt = normalize_text(text)
        if "频率" in nt or "freq" in nt or "frequency" in nt:
            freq_col = col
        if "s11" in nt or "s₁₁" in str(text).lower():
            s11_col = col
        if "-10" in nt or "参考" in nt:
            ref_col = col

    # 宽松 fallback：按前三个常见列推断。
    freq_col = freq_col or "A"
    s11_col = s11_col or "B"
    ref_col = ref_col or "C"

    max_row = max((split_ref(ref)[1] for ref in cells), default=1)
    freq_values: list[float] = []
    s11_values: list[float] = []
    ref_values: list[float] = []
    blank_count = 0
    non_numeric_count = 0
    end_row = 1
    for row in range(2, max_row + 1):
        fv = cells.get(f"{freq_col}{row}")
        sv = cells.get(f"{s11_col}{row}")
        rv = cells.get(f"{ref_col}{row}")
        # 一行两列都空，认为超出数据区而不是数据空白。
        if fv is None and sv is None:
            continue
        end_row = row
        fn = numeric(fv)
        sn = numeric(sv)
        rn = numeric(rv)
        if fv is None or sv is None:
            blank_count += 1
        if fn is None or sn is None:
            non_numeric_count += 1
        else:
            freq_values.append(fn)
            s11_values.append(sn)
        if rn is not None:
            ref_values.append(rn)

    return DataRegion(
        header_row=1,
        freq_col=freq_col,
        s11_col=s11_col,
        ref_col=ref_col,
        start_row=2,
        end_row=end_row,
        freq_values=freq_values,
        s11_values=s11_values,
        ref_values=ref_values,
        headers=headers,
        non_numeric_count=non_numeric_count,
        blank_count=blank_count,
    )


def parse_sheet(pkg: dict[str, bytes], part: str, shared_strings: list[str]) -> SheetModel | None:
    if part not in pkg:
        return None
    root = parse_xml(pkg[part])
    cells: dict[str, Any] = {}
    row_heights: dict[int, float] = {}
    col_widths: dict[int, float] = {}

    for col in root.findall(".//x:col", NS):
        try:
            width = float(col.get("width", ""))
            for i in range(int(col.get("min", "0")), int(col.get("max", "0")) + 1):
                col_widths[i] = width
        except ValueError:
            pass

    for row in root.findall(".//x:row", NS):
        try:
            row_num = int(row.get("r", "0"))
        except ValueError:
            row_num = 0
        if row.get("ht"):
            try:
                row_heights[row_num] = float(row.get("ht", ""))
            except ValueError:
                pass
        for c in row.findall("x:c", NS):
            ref = c.get("r")
            if ref:
                cells[ref] = parse_cell_value(c, shared_strings)

    sheet = SheetModel(part, cells, row_heights, col_widths)
    sheet.data = infer_data_region(cells)
    return sheet


def extract_num_values(parent: ET.Element | None) -> tuple[list[float], str | None]:
    if parent is None:
        return [], None
    ref = None
    f = parent.find(".//c:f", NS)
    if f is not None and f.text:
        ref = f.text
    pts = []
    for pt in parent.findall(".//c:pt", NS):
        try:
            idx = int(pt.get("idx", "0"))
        except ValueError:
            idx = 0
        v = pt.find("c:v", NS)
        if v is not None:
            n = numeric(v.text)
            if n is not None:
                pts.append((idx, n))
    pts.sort(key=lambda x: x[0])
    return [v for _, v in pts], ref


def extract_pt_index_info(parent: ET.Element | None) -> tuple[list[int], int | None]:
    """解析数值缓存的数据点 idx 序列与声明总点数 ptCount。

    OOXML 中 c:numCache/c:strCache 用 c:ptCount@val 声明数据点总数，
    每个实际有值的点用 c:pt@idx 标记其序号；被跳过（源单元格为空）的点不会写出 c:pt。
    因此：idx 序列不等于 0..ptCount-1 连续序列，即说明折线中间有空值断点。
    返回 (已写出且有数值的 idx 升序列表, ptCount)。
    """
    if parent is None:
        return [], None
    pt_count: int | None = None
    pc = parent.find(".//c:ptCount", NS)
    if pc is not None and pc.get("val") and pc.get("val", "").isdigit():
        pt_count = int(pc.get("val", "0"))
    indices: list[int] = []
    for pt in parent.findall(".//c:pt", NS):
        try:
            idx = int(pt.get("idx", "0"))
        except ValueError:
            continue
        v = pt.find("c:v", NS)
        # 仅统计写出且为有效数值的点（空 c:v 或非数值同样视为断点）。
        if v is not None and numeric(v.text) is not None:
            indices.append(idx)
    indices.sort()
    return indices, pt_count


def line_props(sppr: ET.Element | None) -> dict[str, Any]:
    if sppr is None:
        return {}
    ln = sppr.find(".//a:ln", NS)
    if ln is None:
        return {}
    out: dict[str, Any] = {}
    if ln.get("w"):
        try:
            w = int(ln.get("w", "0"))
            out["width_emu"] = w
            out["width_pt"] = w / EMU_PER_PT
        except ValueError:
            pass
    rgb = ln.find(".//a:srgbClr", NS)
    if rgb is not None and rgb.get("val"):
        out["color"] = rgb.get("val", "").upper()
        out["color_source"] = "srgbClr"
    scheme = ln.find(".//a:schemeClr", NS)
    if scheme is not None and scheme.get("val"):
        out["scheme_color"] = scheme.get("val")
        if scheme.get("lastClr") and "color" not in out:
            out["color"] = scheme.get("lastClr", "").upper()
            out["color_source"] = "schemeClr.lastClr"
    sysclr = ln.find(".//a:sysClr", NS)
    if sysclr is not None:
        if sysclr.get("val"):
            out["sys_color"] = sysclr.get("val")
        if sysclr.get("lastClr") and "color" not in out:
            out["color"] = sysclr.get("lastClr", "").upper()
            out["color_source"] = "sysClr.lastClr"
    dash = ln.find("a:prstDash", NS)
    if dash is not None and dash.get("val"):
        out["dash"] = dash.get("val")
    return out


def fill_color(sppr: ET.Element | None) -> str | None:
    if sppr is None:
        return None
    nofill = sppr.find(".//a:noFill", NS)
    if nofill is not None:
        return "none"
    rgb = sppr.find(".//a:solidFill/a:srgbClr", NS)
    if rgb is not None and rgb.get("val"):
        return rgb.get("val", "").upper()
    scheme = sppr.find(".//a:solidFill/a:schemeClr", NS)
    if scheme is not None and scheme.get("lastClr"):
        return scheme.get("lastClr", "").upper()
    sysclr = sppr.find(".//a:solidFill/a:sysClr", NS)
    if sysclr is not None and sysclr.get("lastClr"):
        return sysclr.get("lastClr", "").upper()
    return None


def parse_manual_layout(layout_el: ET.Element | None) -> dict[str, Any]:
    """解析 c:layout/c:manualLayout，返回 plotArea 相对 chartSpace 的归一化布局（0–1）。

    c:x/c:y/c:w/c:h 的取值是相对 chartSpace 宽高的比例（0–1），
    xMode/yMode 为 "edge" 时 x/y 是左上角坐标，为 "factor" 时是相对默认位置的偏移量
    （Excel/WPS 生成的 manualLayout 一般用 edge，这里按 edge 语义处理，
    factor 场景没有足够信息换算，不作为居中判断依据，返回空）。
    """
    if layout_el is None:
        return {}
    manual = layout_el.find("c:manualLayout", NS)
    if manual is None:
        return {}

    def val(tag: str) -> float | None:
        e = manual.find(f"c:{tag}", NS)
        if e is None or e.get("val") is None:
            return None
        try:
            return float(e.get("val", ""))
        except ValueError:
            return None

    def mode(tag: str) -> str:
        e = manual.find(f"c:{tag}", NS)
        val_attr = e.get("val") if e is not None else None
        return val_attr if val_attr else "factor"

    x_mode = mode("xMode")
    y_mode = mode("yMode")
    x, y, w, h = val("x"), val("y"), val("w"), val("h")
    if x_mode != "edge" or y_mode != "edge" or None in (x, y, w, h):
        # factor 模式或缺失坐标：无法直接换算为绝对布局，不参与居中判断。
        return {}
    return {"x": x, "y": y, "w": w, "h": h, "x_mode": x_mode, "y_mode": y_mode}


def collect_fonts(root: ET.Element) -> tuple[list[float], list[str], list[str], bool]:
    sizes: list[float] = []
    names: list[str] = []
    colors: list[str] = []
    wordart = False
    for el in root.iter():
        tag = el.tag.split("}")[-1]
        if tag in {"effectDag", "scene3d", "sp3d", "prstGeom"}:
            wordart = True
        if tag in {"rPr", "defRPr"}:
            if el.get("sz"):
                try:
                    sizes.append(int(el.get("sz", "0")) / 100.0)
                except ValueError:
                    pass
            latin = el.find("a:latin", NS)
            if latin is not None and latin.get("typeface"):
                names.append(latin.get("typeface", ""))
            rgb = el.find(".//a:srgbClr", NS)
            if rgb is not None and rgb.get("val"):
                colors.append(rgb.get("val", "").upper())
    return sizes, names, colors, wordart


def parse_axis(ax: ET.Element, tag: str) -> AxisModel:
    def child_val(name: str) -> str | None:
        e = ax.find(f"c:{name}", NS)
        return e.get("val") if e is not None else None

    scaling = ax.find("c:scaling", NS)
    min_el = scaling.find("c:min", NS) if scaling is not None else None
    max_el = scaling.find("c:max", NS) if scaling is not None else None
    major = ax.find("c:majorUnit", NS)
    title_el = ax.find("c:title", NS)
    title = extract_text(title_el) if title_el is not None else None
    body_pr = title_el.find(".//a:bodyPr", NS) if title_el is not None else None
    rot = None
    vert = None
    if body_pr is not None:
        if body_pr.get("rot"):
            try:
                rot = int(body_pr.get("rot", "0"))
            except ValueError:
                pass
        # a:bodyPr/@vert 描述文本竖向排布方式，非"horz"（水平）即视为竖排；
        # 常见竖排值：vert（自上而下）、vert270、eaVert（东亚竖排）、wordArtVert(RTL)。
        vert = body_pr.get("vert")
    # 解析 c:title/c:layout/c:manualLayout（如存在），用于底部居中位置判定。
    title_layout: dict[str, Any] = {}
    if title_el is not None:
        title_layout = parse_manual_layout(title_el.find("c:layout", NS))
    sizes, names, colors, _ = collect_fonts(ax)
    return AxisModel(
        tag=tag,
        pos=child_val("axPos"),
        title=title,
        orientation=child_val("orientation"),
        min_val=float(min_el.get("val")) if min_el is not None and min_el.get("val") else None,
        max_val=float(max_el.get("val")) if max_el is not None and max_el.get("val") else None,
        major_unit=float(major.get("val")) if major is not None and major.get("val") else None,
        num_fmt=(ax.find("c:numFmt", NS).get("formatCode") if ax.find("c:numFmt", NS) is not None else None),
        line=line_props(ax.find("c:spPr", NS)),
        font_sizes=sizes,
        font_names=names,
        font_colors=colors,
        title_rotation=rot,
        title_vert=vert,
        title_layout=title_layout,
    )


def parse_chart(pkg: dict[str, bytes], part: str, anchor: Anchor | None) -> ChartModel | None:
    if part not in pkg:
        return None
    root = parse_xml(pkg[part])
    chart = ChartModel(path=part, anchor=anchor)
    for typ in ["scatterChart", "lineChart", "barChart", "pieChart", "areaChart"]:
        if root.find(f".//c:{typ}", NS) is not None:
            chart.chart_type = typ
            break
    title_el = root.find(".//c:title", NS)
    chart.title = extract_text(title_el) if title_el is not None else None
    # 仅从 c:title 子树提取标题字体/字号/颜色，避免混入坐标轴、图例等其他文本。
    if title_el is not None:
        chart.title_font_sizes, chart.title_font_names, chart.title_font_colors, _ = collect_fonts(title_el)
    chart.has_legend = root.find(".//c:legend", NS) is not None
    chart.chart_line = line_props(root.find("c:spPr", NS))
    plot_area = root.find(".//c:plotArea", NS)
    if plot_area is not None:
        chart.plot_line = line_props(plot_area.find("c:spPr", NS))
        chart.plot_fill = fill_color(plot_area.find("c:spPr", NS))
        chart.plot_layout = parse_manual_layout(plot_area.find("c:layout", NS))
    chart.font_sizes, chart.font_names, chart.font_colors, chart.wordart_like = collect_fonts(root)

    _sc = root.find(".//c:scatterChart", NS)
    if _sc is not None:
        scatter_style = _sc.find("c:scatterStyle", NS)
        chart.scatter_style = scatter_style.get("val") if scatter_style is not None else None
    container = _sc if _sc is not None else root.find(".//c:lineChart", NS)
    if container is not None:
        for ser in container.findall("c:ser", NS):
            model = SeriesModel()
            tx = ser.find("c:tx", NS)
            model.name = extract_text(tx) if tx is not None else None
            _xv = ser.find("c:xVal", NS)
            model.x_values, model.x_ref = extract_num_values(_xv if _xv is not None else ser.find("c:cat", NS))
            _yv = ser.find("c:yVal", NS)
            model.y_values, model.y_ref = extract_num_values(_yv if _yv is not None else ser.find("c:val", NS))
            # 记录 y 轴数据点 idx 与 ptCount，供断裂检测使用（idx 不连续 => 存在空值断点）。
            model.y_indices, model.y_pt_count = extract_pt_index_info(_yv if _yv is not None else ser.find("c:val", NS))
            model.line = line_props(ser.find("c:spPr", NS))
            smooth = ser.find("c:smooth", NS)
            model.smooth = smooth is not None and smooth.get("val", "1") not in {"0", "false", "False"}
            marker = ser.find("c:marker", NS)
            if marker is not None:
                sym = marker.find("c:symbol", NS)
                size = marker.find("c:size", NS)
                sppr = marker.find("c:spPr", NS)
                model.marker = {
                    "symbol": sym.get("val") if sym is not None else None,
                    "size": int(size.get("val", "0")) if size is not None and size.get("val", "0").isdigit() else None,
                    "line": line_props(sppr),
                    "fill": fill_color(sppr),
                }
            # 解析按点覆盖 c:dPt：某个数据点可单独设置 marker（隐藏或改形/改色）。
            for dpt in ser.findall("c:dPt", NS):
                idx_el = dpt.find("c:idx", NS)
                try:
                    dpt_idx = int(idx_el.get("val", "")) if idx_el is not None else -1
                except ValueError:
                    dpt_idx = -1
                dmarker = dpt.find("c:marker", NS)
                d_symbol: str | None = None
                has_marker = True
                if dmarker is not None:
                    dsym = dmarker.find("c:symbol", NS)
                    d_symbol = dsym.get("val") if dsym is not None else None
                    # 该点显式将标记设为 none，即隐藏此点标记。
                    if d_symbol == "none":
                        has_marker = False
                model.dpt_markers.append({"idx": dpt_idx, "has_marker": has_marker, "symbol": d_symbol})
            chart.series.append(model)

    for tag in ["valAx", "catAx", "dateAx", "serAx"]:
        for ax in root.findall(f".//c:{tag}", NS):
            chart.axes.append(parse_axis(ax, tag))
    return chart


def parse_anchor(anchor_el: ET.Element) -> Anchor:
    kind = anchor_el.tag.split("}")[-1]
    out = Anchor(kind=kind)
    frm = anchor_el.find("xdr:from", NS)
    if frm is not None:
        col = frm.find("xdr:col", NS)
        row = frm.find("xdr:row", NS)
        out.col0 = int(col.text) if col is not None and col.text else None
        out.row0 = int(row.text) if row is not None and row.text else None
    to = anchor_el.find("xdr:to", NS)
    if to is not None:
        col = to.find("xdr:col", NS)
        row = to.find("xdr:row", NS)
        out.col1 = int(col.text) if col is not None and col.text else None
        out.row1 = int(row.text) if row is not None and row.text else None
    ext = anchor_el.find("xdr:ext", NS)
    if ext is not None:
        out.cx_emu = int(ext.get("cx", "0"))
        out.cy_emu = int(ext.get("cy", "0"))
    return out


def parse_drawings_and_charts(model: WorkbookModel, sheet_part: str) -> None:
    pkg = model.package
    sheet_rels = read_rels(pkg, sheet_part)
    drawing_parts = [t for t in sheet_rels.values() if "drawing" in t.lower() and t in pkg]
    for drawing_part in drawing_parts:
        root = parse_xml(pkg[drawing_part])
        drawing_rels = read_rels(pkg, drawing_part)
        for anchor_el in root.findall(".//xdr:oneCellAnchor", NS) + root.findall(".//xdr:twoCellAnchor", NS):
            anchor = parse_anchor(anchor_el)
            if anchor_el.find(".//xdr:pic", NS) is not None:
                model.image_anchors.append(anchor)
            chart_ref = anchor_el.find(".//c:chart", NS)
            if chart_ref is not None:
                rid = chart_ref.get(q("r", "id"))
                chart_part = drawing_rels.get(rid or "")
                if chart_part:
                    chart = parse_chart(pkg, chart_part, anchor)
                    if chart:
                        model.charts.append(chart)


def parse_workbook(path: Path) -> WorkbookModel:
    model = WorkbookModel(path=path, suffix_ok=path.suffix.lower() in {".xlsx", ".xlsm"})
    model.ok_zip, model.package, errors = open_package(path)
    model.parse_errors.extend(errors)
    if not model.ok_zip:
        return model
    model.media_files = [p for p in model.package if p.startswith("xl/media/")]
    try:
        model.shared_strings = parse_shared_strings(model.package)
        # 本任务文件都是 sheet1；若 workbook.xml 可解析，也优先找第一个 worksheet part。
        sheet_part = "xl/worksheets/sheet1.xml"
        if sheet_part not in model.package:
            sheet_parts = sorted(p for p in model.package if p.startswith("xl/worksheets/sheet") and p.endswith(".xml"))
            sheet_part = sheet_parts[0] if sheet_parts else sheet_part
        model.sheet = parse_sheet(model.package, sheet_part, model.shared_strings)
        if model.sheet:
            parse_drawings_and_charts(model, sheet_part)
    except Exception as exc:  # noqa: BLE001
        model.parse_errors.append(f"解析失败：{exc}")
    return model


# ───────────────────────────── 评分辅助 ─────────────────────────────

def close_list(a: list[float], b: list[float], tol: float = 1e-6) -> bool:
    return len(a) == len(b) and all(abs(x - y) <= tol for x, y in zip(a, b))


def is_increasing(values: list[float]) -> bool:
    return len(values) >= 2 and all(values[i] <= values[i + 1] for i in range(len(values) - 1))


def is_strictly_decreasing(values: list[float]) -> bool:
    return len(values) >= 2 and all(values[i] > values[i + 1] for i in range(len(values) - 1))


def parse_ref_range(formula: str | None) -> dict[str, Any] | None:
    if not formula:
        return None
    expr = formula.split(",")[0].strip().lstrip("=")
    if "!" in expr:
        expr = expr.rsplit("!", 1)[1]
    expr = expr.replace("'", "").replace("$", "").strip()
    m = re.fullmatch(r"([A-Za-z]+)(\d+)(?::([A-Za-z]+)(\d+))?", expr)
    if not m:
        return None
    col1 = m.group(1).upper()
    row1 = int(m.group(2))
    col2 = (m.group(3) or m.group(1)).upper()
    row2 = int(m.group(4) or m.group(2))
    c1, c2 = col_to_num(col1), col_to_num(col2)
    row_start, row_end = min(row1, row2), max(row1, row2)
    col_start, col_end = min(c1, c2), max(c1, c2)
    return {
        "raw": formula,
        "col_start": col_start,
        "col_end": col_end,
        "row_start": row_start,
        "row_end": row_end,
        "row_count": row_end - row_start + 1,
        "single_column": col_start == col_end,
    }


def referenced_values(model: WorkbookModel, formula: str | None) -> list[float]:
    if not formula or not model.sheet:
        return []
    ref_range = parse_ref_range(formula)
    if not ref_range:
        return []
    values: list[float] = []
    for row in range(ref_range["row_start"], ref_range["row_end"] + 1):
        for col in range(ref_range["col_start"], ref_range["col_end"] + 1):
            n = numeric(model.sheet.cells.get(f"{num_to_col(col)}{row}"))
            if n is not None:
                values.append(n)
    return values


def series_x_values(model: WorkbookModel, series: SeriesModel | None, fallback_source: bool = False) -> list[float]:
    if not series:
        return []
    if series.x_values:
        return series.x_values
    vals = referenced_values(model, series.x_ref)
    if vals:
        return vals
    return model.sheet.data.freq_values if fallback_source and model.sheet else []


def series_y_values(model: WorkbookModel, series: SeriesModel | None, fallback_source: bool = False) -> list[float]:
    if not series:
        return []
    if series.y_values:
        return series.y_values
    vals = referenced_values(model, series.y_ref)
    if vals:
        return vals
    return model.sheet.data.s11_values if fallback_source and model.sheet else []


def is_reference_series(series: SeriesModel, model: WorkbookModel | None = None) -> bool:
    vals = series_y_values(model, series) if model else series.y_values
    name_hit = "-10" in normalize_text(series.name) or "参考" in normalize_text(series.name)
    if len(vals) < 2:
        return name_hit
    near = sum(1 for v in vals if abs(v + 10) <= 0.25)
    return near / len(vals) >= 0.8 or name_hit


def main_s11_series(chart: ChartModel | None, model: WorkbookModel | None = None) -> SeriesModel | None:
    if not chart:
        return None
    non_ref = [s for s in chart.series if not is_reference_series(s, model)]
    named = [s for s in non_ref if "s11" in normalize_text(s.name)]
    if named:
        return named[0]
    if model and model.sheet:
        d = model.sheet.data
        for s in non_ref:
            if close_list(series_y_values(model, s), d.s11_values) or (d.s11_col and d.s11_col in (s.y_ref or "")):
                return s
    return non_ref[0] if non_ref else (chart.series[0] if chart.series else None)


def reference_series(chart: ChartModel | None, model: WorkbookModel | None = None) -> SeriesModel | None:
    if not chart:
        return None
    for s in chart.series:
        if is_reference_series(s, model):
            return s
    return None


def first_chart(model: WorkbookModel) -> ChartModel | None:
    return model.charts[0] if model.charts else None


def chart_identity(model: WorkbookModel, chart: ChartModel | None) -> dict[str, object]:
    return {"chart_path": chart.path if chart else None, "chart_count": len(model.charts)}


def select_s11_chart(model: WorkbookModel) -> ChartModel | None:
    if not model.charts:
        return None
    if len(model.charts) == 1:
        return model.charts[0]
    d = model.sheet.data if model.sheet else DataRegion()
    best: tuple[int, int, ChartModel] | None = None
    for idx, chart in enumerate(model.charts):
        score = 0
        if chart.chart_type in {"scatterChart", "lineChart"}:
            score += 4
        if "s11" in normalize_text(chart.title):
            score += 3
        main = main_s11_series(chart, model)
        ref = reference_series(chart, model)
        if main:
            score += 2
            if "s11" in normalize_text(main.name):
                score += 3
            xvals = series_x_values(model, main)
            yvals = series_y_values(model, main)
            if close_list(xvals, d.freq_values) or (d.freq_col and d.freq_col in (main.x_ref or "")):
                score += 3
            if close_list(yvals, d.s11_values) or (d.s11_col and d.s11_col in (main.y_ref or "")):
                score += 4
            if is_increasing(xvals):
                score += 1
        if ref:
            score += 2
        if axis_by_pos(chart, "b"):
            score += 1
        if axis_by_pos(chart, "l"):
            score += 1
        if chart.anchor and chart.anchor.width_cm > 0 and chart.anchor.height_cm > 0:
            score += 1
        item = (score, -idx, chart)
        if best is None or item > best:
            best = item
    return best[2] if best and best[0] > 0 else first_chart(model)


def normalize_rgb(color: str | None) -> str | None:
    if not color:
        return None
    value = color.strip().upper().lstrip("#")
    if len(value) == 8:
        value = value[-6:]
    if len(value) != 6 or not re.fullmatch(r"[0-9A-F]{6}", value):
        return None
    return value


def color_is_blue(color: str | None) -> bool:
    rgb = normalize_rgb(color)
    if not rgb:
        return False
    r, g, b = int(rgb[0:2], 16), int(rgb[2:4], 16), int(rgb[4:6], 16)
    return b >= 150 and b > r * 1.3 and b > g * 1.2


def color_is_gray(color: str | None) -> bool:
    rgb = normalize_rgb(color)
    if not rgb:
        return False
    r, g, b = int(rgb[0:2], 16), int(rgb[2:4], 16), int(rgb[4:6], 16)
    return max(r, g, b) - min(r, g, b) <= 35


def color_is_dark(color: str | None) -> bool:
    rgb = normalize_rgb(color)
    if not rgb:
        return False
    r, g, b = int(rgb[0:2], 16), int(rgb[2:4], 16), int(rgb[4:6], 16)
    return max(r, g, b) < 120


def color_is_blackish(color: str | None) -> bool:
    if color is None:
        return True  # 未显式设置时按 Excel 默认黑色处理
    return normalize_rgb(color) in {"000000"} or color_is_dark(color)


def line_is_blue(line: dict[str, object]) -> bool:
    color = line.get("color")
    return color_is_blue(color if isinstance(color, str) else None) or line.get("scheme_color") in {"accent1", "accent5", "hlink"}


def line_is_gray(line: dict[str, object], allow_theme: bool = False) -> bool:
    color = line.get("color")
    return color_is_gray(color if isinstance(color, str) else None) or (allow_theme and line.get("scheme_color") in {"tx1", "tx2", "bg1", "bg2"})


def effective_line_width(line: dict[str, object], default: float | None = None) -> tuple[float | None, bool]:
    width = line.get("width_pt")
    if isinstance(width, (int, float)):
        return float(width), False
    return default, default is not None


def axis_by_pos(chart: ChartModel | None, pos: str) -> AxisModel | None:
    if not chart:
        return None
    for ax in chart.axes:
        if ax.pos == pos:
            return ax
    if pos == "b":
        return next((ax for ax in chart.axes if ax.tag in {"catAx", "dateAx"}), None)
    if pos == "l":
        vals = [ax for ax in chart.axes if ax.tag == "valAx"]
        return vals[-1] if vals else None
    return None


def pass_result(rule_id: str, name: str, points: int, reason: str, evidence: dict[str, Any] | None = None, file: str | None = None) -> RuleResult:
    return RuleResult(rule_id, name, points, True, reason, evidence or {}, file=file)


def fail_result(rule_id: str, name: str, points: int, reason: str, evidence: dict[str, Any] | None = None, file: str | None = None, kind: str = "positive") -> RuleResult:
    return RuleResult(rule_id, name, points, False, reason, evidence or {}, file=file, kind=kind)


# ───────────────────────────── 维度 1 ─────────────────────────────

def gate_file_open(model: WorkbookModel) -> RuleResult:
    ok = model.suffix_ok and model.ok_zip and "[Content_Types].xml" in model.package and "xl/workbook.xml" in model.package
    return RuleResult(
        "D1-G1", "文件为 .xlsx/.xlsm 且工作簿可正常打开", 0, ok,
        "可打开 OOXML 工作簿" if ok else "文件格式不正确或无法作为 OOXML 工作簿打开",
        {"suffix_ok": model.suffix_ok, "ok_zip": model.ok_zip, "errors": model.parse_errors}, model.path.name, "gate",
    )


# ───────────────────────────── 维度 2 正向规则 ─────────────────────────────

def check_p01(model: WorkbookModel) -> RuleResult:
    # 细则：图表左上角位于数据区域右侧或下方空白区域，距离数据区域边界至少1列或2行；
    # 图表整体宽度18cm–21cm，高度10cm–13cm；图表为Excel可编辑折线图对象，不是图片。
    name = "S11图表对象位于数据工作表内数据区域右侧（≥1列）或下方（≥2行）空白区域，宽度18–21cm，高度10–13cm，且为可编辑折线图对象（非图片）"
    chart = select_s11_chart(model)
    d = model.sheet.data if model.sheet else DataRegion()
    if not chart or not chart.anchor:
        return fail_result("P01", name, 1, "未发现图表 anchor", file=model.path.name)
    a = chart.anchor

    # ① 图表左上角位于数据区域右侧空白区域，距离至少1列：
    #    xdr:col 为0-indexed，d.max_col_index 为1-indexed列号（A=1）；
    #    "右侧至少1列"要求图表起始列编号 > 数据最大列，即 (col0+1) > max_col_index，即 col0 >= max_col_index。
    right_of_data = (a.col0 is not None and d.max_col_index > 0 and a.col0 >= d.max_col_index)

    # ② 图表左上角位于数据区域下方空白区域，距离至少2行：
    #    xdr:row 为0-indexed，d.end_row 为1-indexed行号；
    #    "下方至少2行"要求图表起始行编号 >= 数据最后行+2，即 (row0+1) >= end_row+2，即 row0 >= end_row+1。
    below_data = (a.row0 is not None and d.end_row > 0 and a.row0 >= d.end_row + 1)

    # 满足"右侧或下方"其中之一即可
    position_ok = right_of_data or below_data

    # ③ 图表整体宽度18cm–21cm，高度10cm–13cm（通过anchor的EMU尺寸判断）
    size_ok = 18 <= a.width_cm <= 21 and 10 <= a.height_cm <= 13

    # ④ 图表为Excel可编辑折线图对象，不是图片：
    #    细则明确说"折线图"，只接受 lineChart；
    #    同时确认该对象不是图片（image_anchors中无重叠的图片锚点）。
    type_ok = chart.chart_type == "lineChart"
    not_image = not any(
        img.col0 == a.col0 and img.row0 == a.row0
        for img in model.image_anchors
    )

    ok = position_ok and size_ok and type_ok and not_image

    reason_parts = [
        f"左上角 col={a.col0}(0-idx), row={a.row0}(0-idx)",
        f"数据区最大列={d.max_col_index}(1-idx), 数据末行={d.end_row}(1-idx)",
        f"右侧≥1列={'是' if right_of_data else '否'}, 下方≥2行={'是' if below_data else '否'}",
        f"尺寸 {a.width_cm:.2f}cm × {a.height_cm:.2f}cm (要求18–21cm × 10–13cm)",
        f"图表类型={chart.chart_type} (要求lineChart)",
        f"非图片对象={'是' if not_image else '否'}",
    ]
    evidence = {
        "anchor": asdict(a),
        "data_max_col_1idx": d.max_col_index,
        "data_end_row_1idx": d.end_row,
        "right_of_data": right_of_data,
        "below_data": below_data,
        "width_cm": a.width_cm,
        "height_cm": a.height_cm,
        "chart_type": chart.chart_type,
        "not_image": not_image,
        **chart_identity(model, chart),
    }
    return RuleResult("P01", name, 1, ok, "；".join(reason_parts), evidence, model.path.name)


def check_p02(model: WorkbookModel) -> RuleResult:
    # 细则：图表横轴数据来源于数据区域中的频率列，纵轴数据来源于同一行对应的S11数值列；
    # 图表数据点数量与源数据有效行数一致，不能手动编造或遗漏数据点。
    name = "S11图表横轴引用频率列、纵轴引用同一行S11数值列，数据点数量与源数据有效行数一致且无手动编造/遗漏"
    chart = select_s11_chart(model)
    main = main_s11_series(chart, model)
    d = model.sheet.data if model.sheet else DataRegion()
    if not main:
        return fail_result("P02", name, 1, "未发现主 S11 系列", chart_identity(model, chart), model.path.name)

    x_range = parse_ref_range(main.x_ref)
    y_range = parse_ref_range(main.y_ref)
    xv = series_x_values(model, main)
    yv = series_y_values(model, main)

    freq_col_index = col_to_num(d.freq_col)
    s11_col_index = col_to_num(d.s11_col)

    # ① 横轴数据来源于数据区域中的频率列：必须是办公软件图表公式中的单列引用，且列号匹配频率列。
    x_from_freq_col = bool(
        x_range
        and x_range["single_column"]
        and x_range["col_start"] == freq_col_index
        and x_range["col_end"] == freq_col_index
    )

    # ② 纵轴数据来源于同一行对应的S11数值列：必须是单列引用，列号匹配S11列，且行区间与横轴完全一致。
    y_from_s11_col = bool(
        y_range
        and y_range["single_column"]
        and y_range["col_start"] == s11_col_index
        and y_range["col_end"] == s11_col_index
    )
    same_rows = bool(
        x_range
        and y_range
        and x_range["row_start"] == y_range["row_start"]
        and x_range["row_end"] == y_range["row_end"]
    )

    # ③ 图表数据点数量与源数据有效行数一致：引用行数、横轴缓存点数、纵轴缓存点数均需一致。
    point_count_ok = bool(
        x_range
        and y_range
        and x_range["row_count"] == d.valid_rows
        and y_range["row_count"] == d.valid_rows
        and len(xv) == d.valid_rows
        and len(yv) == d.valid_rows
    )

    # ④ 不能手动编造或遗漏数据点：图表缓存值必须与所引用的源单元格数值一致。
    values_match_source = close_list(xv, d.freq_values) and close_list(yv, d.s11_values)

    ok = x_from_freq_col and y_from_s11_col and same_rows and point_count_ok and values_match_source
    evidence = {
        "x_ref": main.x_ref,
        "y_ref": main.y_ref,
        "x_range": x_range,
        "y_range": y_range,
        "freq_col": d.freq_col,
        "s11_col": d.s11_col,
        "source_valid_rows": d.valid_rows,
        "chart_x_count": len(xv),
        "chart_y_count": len(yv),
        "x_from_freq_col": x_from_freq_col,
        "y_from_s11_col": y_from_s11_col,
        "same_rows": same_rows,
        "point_count_ok": point_count_ok,
        "values_match_source": values_match_source,
        **chart_identity(model, chart),
    }
    reason = (
        f"横轴引用={main.x_ref}，纵轴引用={main.y_ref}；"
        f"横轴来自频率列={'是' if x_from_freq_col else '否'}，"
        f"纵轴来自S11列={'是' if y_from_s11_col else '否'}，"
        f"同行对应={'是' if same_rows else '否'}，"
        f"点数={len(xv)}/{len(yv)}，源有效行数={d.valid_rows}，"
        f"缓存值匹配源数据={'是' if values_match_source else '否'}"
    )
    return RuleResult("P02", name, 1, ok, reason, evidence, model.path.name)



def color_is_visible_border(color: str | None) -> bool:
    """判断边框颜色在白色/浅色背景上是否肉眼可见。

    细则只要求"外侧出现矩形"，未限定边框必须是黑色或中深灰；
    只要不是几乎融入白底的极浅色（如 RGB 三分量均 >=245，接近纯白），
    包括浅灰（如 D9D9D9）在内都应视为可见矩形。
    """
    rgb = normalize_rgb(color)
    if not rgb:
        return False
    r, g, b = int(rgb[0:2], 16), int(rgb[2:4], 16), int(rgb[4:6], 16)
    # 仅排除几乎与白色背景无法区分的极浅色；给予合理容差，不误判浅灰矩形为不可见。
    near_white = r >= 245 and g >= 245 and b >= 245
    return not near_white


def resolve_plot_rect(chart: ChartModel) -> dict[str, float] | None:
    """计算 plotArea 在 chartSpace 内的归一化矩形（0–1坐标，原点左上角）。

    若图表写出了显式 c:layout/c:manualLayout（edge 模式）则直接使用；
    否则无法从 OOXML 得到精确矩形，返回 None（不能凭空假造坐标）。
    """
    layout = chart.plot_layout
    if not layout:
        return None
    x, y, w, h = layout.get("x"), layout.get("y"), layout.get("w"), layout.get("h")
    if None in (x, y, w, h):
        return None
    return {"x0": x, "y0": y, "x1": x + w, "y1": y + h}


def series_centered_in_plot_rect(chart: ChartModel, model: WorkbookModel, tol: float = 0.12) -> tuple[bool, dict[str, Any]]:
    """基于系列数值范围在坐标轴范围中的相对位置，结合 plotArea 矩形判断折线是否居中。

    折线本身贯穿整个坐标轴数据范围（频率从最小到最大、S11从数据最小到最大），
    在没有留白挤到一侧的 manualLayout 时，折线在矩形内的居中程度取决于：
    ① plotArea 矩形本身在 chartSpace 内是否大致居中（四周留白不悬殊）；
    ② 若无显式 manualLayout（Excel/WPS 自动布局），默认即按坐标轴+四周留白居中排布，
       只要矩形有效占据了 chartSpace 主体部分（不是被推到一角），就认为满足"居中"。
    """
    detail: dict[str, Any] = {}
    rect = resolve_plot_rect(chart)
    detail["plot_rect"] = rect

    if rect is None:
        # 无显式 manualLayout：Excel/WPS 默认布局本身即让 plotArea 居中于 chartSpace，
        # 只要图表有可绘制的坐标轴和系列即可视为满足，不能因为缺少显式坐标而误判不合格。
        detail["layout_mode"] = "auto"
        return True, detail

    detail["layout_mode"] = "manual"
    cx = (rect["x0"] + rect["x1"]) / 2
    cy = (rect["y0"] + rect["y1"]) / 2
    # chartSpace 归一化中心为 (0.5, 0.5)；允许 tol 容差（默认 ±0.12）。
    centered = abs(cx - 0.5) <= tol and abs(cy - 0.5) <= tol
    detail["center_x"] = cx
    detail["center_y"] = cy
    detail["tol"] = tol
    detail["centered"] = centered
    return centered, detail


def check_p03(model: WorkbookModel) -> RuleResult:
    # 细则：图表外侧出现矩形，矩形底边是横轴，矩形左侧边为纵轴，图表折线图像放置于矩形内居中。
    # 在办公软件（Excel/WPS）生成的 OOXML 中：
    # "图表外侧矩形"对应 chartSpace（图表整体外框）的显式边框（chartSpace/c:spPr/a:ln），
    #   不是 plotArea 内框；plotArea 是绘图区，位于外框内部。
    # 横轴底边：存在 axPos=b 的坐标轴（catAx 或 valAx 均可）。
    # 纵轴左边：存在 axPos=l 的坐标轴。
    # 折线放置于矩形内居中：基于 plotArea 的 manualLayout（若有显式坐标）判断绘图区
    #   是否大致居中于 chartSpace；若无显式坐标则按 Excel/WPS 默认自动居中布局处理。
    name = "图表外侧出现矩形（chartSpace外框），矩形底边是横轴（axPos=b），矩形左侧边为纵轴（axPos=l），折线放置于矩形内居中"
    chart = select_s11_chart(model)
    if not chart:
        return fail_result("P03", name, 3, "未发现图表", file=model.path.name)

    # ① 图表外侧出现矩形：chartSpace 有显式外框线（宽度写出），且颜色在白底上肉眼可见即可；
    #    细则只要求"出现矩形"，不要求边框必须是黑色/中深灰，浅灰矩形同样可见，应合理判定为合格。
    outer_rect_ok = bool(
        chart.chart_line
        and chart.chart_line.get("width_emu") is not None
        and color_is_visible_border(chart.chart_line.get("color"))
    )

    # ② 矩形底边是横轴：存在 axPos=b 的坐标轴。
    x_axis = axis_by_pos(chart, "b")
    x_axis_bottom = x_axis is not None

    # ③ 矩形左侧边为纵轴：存在 axPos=l 的坐标轴。
    y_axis = axis_by_pos(chart, "l")
    y_axis_left = y_axis is not None

    # ④ 折线放置于矩形内居中：结合 plotArea 相对 chartSpace 的布局（manualLayout，若存在）
    #    判断绘图区矩形中心是否落在 chartSpace 中心附近；无显式布局时按自动居中处理。
    series_centered, center_detail = series_centered_in_plot_rect(chart, model)
    has_series = bool(chart.series and any(len(s.x_values) > 0 or len(s.y_values) > 0 for s in chart.series))
    anchor_valid = bool(chart.anchor and chart.anchor.width_cm > 0 and chart.anchor.height_cm > 0)
    series_in_rect = has_series and anchor_valid and series_centered

    ok = outer_rect_ok and x_axis_bottom and y_axis_left and series_in_rect

    evidence = {
        "outer_rect_ok": outer_rect_ok,
        "chart_line": chart.chart_line,
        "x_axis_bottom": x_axis_bottom,
        "y_axis_left": y_axis_left,
        "axis_positions": [a.pos for a in chart.axes],
        "has_series": has_series,
        "anchor_valid": anchor_valid,
        "series_centered": series_centered,
        "center_detail": center_detail,
        **chart_identity(model, chart),
    }
    reason = (
        f"图表外侧矩形={'是' if outer_rect_ok else '否'}（chartSpace外框，颜色={chart.chart_line.get('color')}）；"
        f"横轴底边（axPos=b）={'是' if x_axis_bottom else '否'}；"
        f"纵轴左边（axPos=l）={'是' if y_axis_left else '否'}；"
        f"折线在矩形内居中={'是' if series_in_rect else '否'}"
        f"（布局模式={center_detail.get('layout_mode')}，"
        f"plotArea矩形={center_detail.get('plot_rect')}）"
    )
    return RuleResult("P03", name, 3, ok, reason, evidence, model.path.name)


def check_p04(model: WorkbookModel) -> RuleResult:
    # 细则：每个图表中出现1条蓝色折线，线条为单实线，线宽1.5磅–2.5磅，
    # 数据点之间按频率顺序从左到右连接；折线没有被平滑成曲线，也没有断裂为多条无关线段。
    name = "图表中恰好1条蓝色单实线S11折线，线宽1.5-2.5磅，按频率顺序从左到右连接，未平滑，未断裂为多段"
    chart = select_s11_chart(model)
    main = main_s11_series(chart, model)
    if not chart or not main:
        return fail_result("P04", name, 3, "未发现图表或主S11系列", file=model.path.name)

    # 细则没有要求 lineChart，只要求"折线"外观。
    # Excel/WPS 的散点折线图（scatterChart + spline=0）与折线图（lineChart）视觉上等价，均满足。
    # 只计非参考系列，确认图表中恰好1条蓝色折线。
    non_ref_series = [s for s in chart.series if not is_reference_series(s, model)]

    # ① 恰好1条蓝色折线：非参考系列数量为1。
    one_line = len(non_ref_series) == 1

    # ② 线条为蓝色：series spPr/ln 颜色为蓝色（RGB 或 scheme_color）。
    line = main.line
    is_blue = line_is_blue(line)

    # ③ 单实线：Excel/WPS 对默认实线通常省略 a:prstDash，仅在虚线时才写出 val="dash"等。
    #    因此缺省（dash 未写出，即 None）或显式 "solid" 都按实线处理；
    #    只有显式写出下列虚线/点线/短划线值时才判为非实线。
    dash_val = line.get("dash")
    non_solid_dashes = {"dash", "dashDot", "sysDash", "lgDash", "sysDot", "dot", "lgDashDot", "lgDashDotDot"}
    is_solid = dash_val not in non_solid_dashes

    # ④ 线宽1.5–2.5磅：宽度必须在 OOXML 中显式写出，不接受推断值。
    width, width_inferred = effective_line_width(line)
    width_ok = (
        width is not None
        and not width_inferred
        and 1.5 <= width <= 2.5
    )

    # ⑤ 数据点按频率顺序从左到右连接：x 值（频率）严格单调不降。
    xv = series_x_values(model, main)
    x_ordered = is_increasing(xv)

    # ⑥ 未平滑成曲线：
    #    lineChart：检查 series/c:smooth，不存在或 val=0 才算未平滑；
    #    scatterChart：必须显式 scatterStyle="line" 才算直线连接。
    #    WPS/Excel 中 scatterStyle 缺失或 smooth/smoothMarker 往往显示为平滑曲线，不能按未平滑加分。
    if chart.chart_type == "scatterChart":
        not_smooth = chart.scatter_style == "line"
    else:
        not_smooth = not main.smooth

    # ⑦ 未断裂为多条无关线段：系列数据点必须连续，无空值断开。
    #    Excel/WPS 中若数据区间含空单元格且以"空距"方式处理，折线会在缺口处断裂成多段；
    #    OOXML 中这些空点不会写出 c:pt，导致 pt idx 序列相对 ptCount 出现缺口（不连续）。
    #    判定：以 ptCount 声明的总点数为基准，检查实际写出的 y idx 是否构成 0..ptCount-1 的完整连续序列。
    yv = main.y_values  # OOXML 缓存中实际有值的数据点
    y_idx = main.y_indices
    y_pt_count = main.y_pt_count
    if y_pt_count is not None:
        expected = y_pt_count
    else:
        # 无 ptCount 声明时退化为用实际写出点数作为基准。
        expected = (max(y_idx) + 1) if y_idx else len(yv)
    # idx 去重后应恰为 0,1,...,expected-1，且点数 >=2 才构成一条折线。
    idx_set = set(y_idx)
    idx_contiguous = (
        expected >= 2
        and len(idx_set) == expected
        and (min(idx_set) == 0 if idx_set else False)
        and (max(idx_set) == expected - 1 if idx_set else False)
    )
    # 同时要求 x/y 点数一致（成对），且缓存点数与 idx 数量吻合（无空 c:v 被跳过）。
    pairs_ok = len(xv) == len(yv) and len(yv) == len(idx_set)
    not_broken = idx_contiguous and pairs_ok and len(yv) >= 2

    ok = one_line and is_blue and is_solid and width_ok and x_ordered and not_smooth and not_broken

    evidence = {
        "chart_type": chart.chart_type,
        "scatter_style": chart.scatter_style,
        "non_ref_series_count": len(non_ref_series),
        "one_line": one_line,
        "line": line,
        "is_blue": is_blue,
        "is_solid": is_solid,
        "dash_val": dash_val,
        "width_pt": width,
        "width_inferred": width_inferred,
        "width_ok": width_ok,
        "x_point_count": len(xv),
        "y_point_count": len(yv),
        "y_pt_count": y_pt_count,
        "y_indices": y_idx,
        "idx_contiguous": idx_contiguous,
        "pairs_ok": pairs_ok,
        "x_ordered": x_ordered,
        "not_smooth": not_smooth,
        "not_broken": not_broken,
        **chart_identity(model, chart),
    }
    reason = (
        f"恰好1条蓝色折线={'是' if one_line else '否'}（非参考系列数={len(non_ref_series)}）；"
        f"蓝色={'是' if is_blue else '否'}（color={line.get('color') or line.get('scheme_color')}）；"
        f"单实线={'是' if is_solid else '否'}（dash={dash_val if dash_val is not None else '缺省(按实线)'}）；"
        f"线宽{'=%.2f磅' % width if width is not None else '未写出'}（要求1.5–2.5磅，ok={'是' if width_ok else '否'}）；"
        f"频率顺序从左到右={'是' if x_ordered else '否'}；"
        f"未平滑={'是' if not_smooth else '否'}（chart_type={chart.chart_type}, scatterStyle={chart.scatter_style}）；"
        f"未断裂={'是' if not_broken else '否'}（ptCount={y_pt_count}，写出{len(idx_set)}点，"
        f"idx连续={'是' if idx_contiguous else '否'}，x/y成对={'是' if pairs_ok else '否'}）"
    )
    return RuleResult("P04", name, 3, ok, reason, evidence, model.path.name)


def check_p05(model: WorkbookModel) -> RuleResult:
    # 细则：每个数据点位置出现蓝色圆形标记，标记直径约0.18cm–0.35cm，
    # 填充为蓝色，边线为蓝色或深蓝色；标记点数量与折线数据点数量一致。
    name = "S11折线每个数据点有蓝色圆形标记，直径0.18-0.35cm，填充蓝色，边线蓝色或深蓝色，标记点数量与数据点一致"
    chart = select_s11_chart(model)
    main = main_s11_series(chart, model)
    if not chart or not main:
        return fail_result("P05", name, 3, "未发现图表或主S11系列", file=model.path.name)

    marker = main.marker if main else {}
    xv = series_x_values(model, main)
    yv = series_y_values(model, main)
    n_points = len(yv) if yv else len(xv)

    # ① 每个数据点显示圆形标记：
    #    c:marker/c:symbol 语义——显式 "circle" 才是圆形；显式 "none" 表示不显示标记（不合格）；
    #    其他形状（square/diamond/triangle 等）非圆形。
    #    symbol 缺省（未写出）在散点/折线图中语义不确定：Excel 自动配色方案下可能显示
    #    自动形状（不保证圆形），也可能不显示。为避免"把缺省直接当圆形"的误判，
    #    要求系列级 marker 必须显式写出 symbol="circle" 才认定为圆形标记。
    symbol = marker.get("symbol") if marker else None
    series_marker_visible = symbol not in {None, "none"}
    series_symbol_circle = symbol == "circle"

    # ② 排除按点 dPt 覆盖导致部分点无标记或非圆形：
    #    若某数据点 c:dPt 将 marker 设为 none（隐藏），或改成非 circle 形状，则并非"每个点都是圆形标记"。
    dpt_overrides = main.dpt_markers if main else []
    dpt_hidden = [d for d in dpt_overrides if not d.get("has_marker")]
    dpt_non_circle = [
        d for d in dpt_overrides
        if d.get("has_marker") and d.get("symbol") is not None and d.get("symbol") != "circle"
    ]
    no_bad_dpt = not dpt_hidden and not dpt_non_circle

    # 综合：所有点都显示圆形标记 = 系列级为显式圆形 且 无按点隐藏/改形覆盖。
    is_circle = series_symbol_circle and no_bad_dpt and series_marker_visible

    # ③ 标记直径约0.18cm–0.35cm：
    #    Excel/WPS 中 c:marker/c:size 的单位是"磅"（point，1pt≈0.03528cm）；
    #    0.18cm≈5.1pt，0.35cm≈9.9pt，故有效范围 5–10（含）。
    size = marker.get("size") if marker else None
    size_ok = size is not None and 5 <= int(size) <= 10

    # ④ 填充为蓝色：c:marker/c:spPr 的 solidFill 为蓝色。
    fill_color_val = marker.get("fill") if marker else None
    fill_blue = color_is_blue(fill_color_val)

    # ⑤ 边线为蓝色或深蓝色：c:marker/c:spPr/a:ln 的颜色为蓝色或深蓝色。
    #    深蓝色：B 通道高、且比 R/G 明显高即可（与 color_is_blue 使用相同判定）。
    marker_line = marker.get("line", {}) if marker else {}
    border_color_val = marker_line.get("color") if marker_line else None
    border_blue_or_dark_blue = (
        color_is_blue(border_color_val)
        or color_is_blue(fill_color_val)   # 边线未显式写时默认跟填充色，视为满足
    )

    # ⑥ 标记点数量与折线数据点数量一致：
    #    系列级 marker 对所有数据点生效；再要求没有 dPt 隐藏任何点（隐藏会使显示的标记点少于数据点）。
    #    x/y 点数须相等且 >=1，且无按点隐藏。
    count_ok = len(xv) == len(yv) and n_points >= 1 and not dpt_hidden

    ok = (
        bool(marker)
        and is_circle
        and size_ok
        and fill_blue
        and border_blue_or_dark_blue
        and count_ok
    )

    evidence = {
        "marker": marker,
        "symbol": symbol,
        "series_marker_visible": series_marker_visible,
        "series_symbol_circle": series_symbol_circle,
        "dpt_override_count": len(dpt_overrides),
        "dpt_hidden": dpt_hidden,
        "dpt_non_circle": dpt_non_circle,
        "no_bad_dpt": no_bad_dpt,
        "is_circle": is_circle,
        "size_pt": size,
        "size_ok": size_ok,
        "fill_color": fill_color_val,
        "fill_blue": fill_blue,
        "border_color": border_color_val,
        "border_blue_or_dark_blue": border_blue_or_dark_blue,
        "x_count": len(xv),
        "y_count": len(yv),
        "count_ok": count_ok,
        **chart_identity(model, chart),
    }
    reason = (
        f"c:marker存在={'是' if bool(marker) else '否'}；"
        f"每点圆形标记={'是' if is_circle else '否'}"
        f"（系列symbol={symbol}，"
        f"dPt隐藏{len(dpt_hidden)}点，dPt非圆{len(dpt_non_circle)}点）；"
        f"尺寸={size}pt（要求5-10pt，即约0.18-0.35cm），ok={'是' if size_ok else '否'}；"
        f"填充蓝色（fill={fill_color_val}）={'是' if fill_blue else '否'}；"
        f"边线蓝/深蓝（border={border_color_val}）={'是' if border_blue_or_dark_blue else '否'}；"
        f"点数一致={len(xv)}x/{len(yv)}y（无隐藏点={'是' if not dpt_hidden else '否'}），ok={'是' if count_ok else '否'}"
    )
    return RuleResult("P05", name, 3, ok, reason, evidence, model.path.name)


def check_p06(model: WorkbookModel) -> RuleResult:
    # 细则：图表底部横轴标题文本为"频率（GHz）"或"频率 (GHz)"，
    # 位于图表底部居中区域，字体为微软雅黑、宋体或Calibri，字号10磅–14磅，颜色为黑色。
    name = '横轴标题文本为"频率（GHz）"或"频率 (GHz)"，位于图表底部居中，字体微软雅黑/宋体/Calibri，字号10-14磅，颜色黑色'
    chart = select_s11_chart(model)

    # ① 位于图表底部横轴：取 axPos=b 的坐标轴。
    ax = axis_by_pos(chart, "b")
    if not ax:
        return fail_result("P06", name, 1, "未找到底部横轴（axPos=b）", chart_identity(model, chart), model.path.name)

    title = ax.title  # 由 parse_axis 从 c:title/c:tx/a:rich 提取

    # ② 标题文本严格为"频率（GHz）"或"频率 (GHz)"（细则仅允许这两种中文写法）：
    #    normalize_text 将全角括号转半角、去空格、转小写后，二者都归一为 "频率(ghz)"。
    #    细则未允许英文 "Frequency(GHz)"，故不再接受该写法。
    nt = normalize_text(title)
    title_ok = nt == "频率(ghz)"

    # ③ 字体为微软雅黑、宋体或Calibri：
    #    Excel/WPS 将字体名写入 c:title 内的 a:rPr/a:latin typeface；
    #    若未显式写出字体名，按 Excel 默认字体（Calibri）处理，视为满足。
    allowed_fonts = {"微软雅黑", "宋体", "Calibri", "+mn-lt", "+mj-lt"}  # +mn-lt/+mj-lt 是 Office 主题默认
    if not ax.font_names:
        # 未写出字体名 → 跟随主题默认（Calibri），视为满足
        font_name_ok = True
    else:
        font_name_ok = any(n in allowed_fonts for n in ax.font_names)

    # ④ 字号10磅–14磅：
    #    Excel/WPS 将字号写入 a:rPr sz（单位：百分之一磅）；
    #    若未显式写出字号，按 Excel 默认字号（10pt 或 11pt）处理，视为满足。
    if not ax.font_sizes:
        font_size_ok = True
    else:
        font_size_ok = all(10 <= s <= 14 for s in ax.font_sizes)

    # ⑤ 颜色为黑色：
    #    Excel/WPS 将颜色写入 a:rPr 的 a:solidFill/a:srgbClr；
    #    若未显式写出颜色，按 Excel 默认颜色（黑色）处理，视为满足。
    if not ax.font_colors:
        font_color_ok = True
    else:
        font_color_ok = all(color_is_blackish(c) for c in ax.font_colors)

    # ⑥ 位于图表底部居中区域：axPos=b 保证在底部；居中需进一步验证标题布局，
    #    不能仅凭 axPos=b 推断。解析 c:title/c:layout/c:manualLayout：
    #    - 有显式手动布局：以标题框水平中心是否接近 chartSpace 中心(x≈0.5)判断居中；
    #      若手动布局把标题明显推离水平中心，则判为不居中（不给通过）。
    #    - 无显式手动布局：Excel/WPS 默认将底部横轴标题水平居中，按默认合规处理。
    title_layout = ax.title_layout
    if title_layout:
        cx = title_layout["x"] + title_layout["w"] / 2
        at_bottom_center = abs(cx - 0.5) <= 0.15
        center_mode = "manual"
    else:
        at_bottom_center = True
        center_mode = "auto(默认居中)"

    ok = bool(title_ok and font_name_ok and font_size_ok and font_color_ok and at_bottom_center)

    evidence = {
        "title": title,
        "normalized": nt,
        "title_ok": title_ok,
        "font_names": ax.font_names,
        "font_name_ok": font_name_ok,
        "font_sizes": ax.font_sizes,
        "font_size_ok": font_size_ok,
        "font_colors": ax.font_colors,
        "font_color_ok": font_color_ok,
        "title_layout": title_layout,
        "center_mode": center_mode,
        "at_bottom_center": at_bottom_center,
        **chart_identity(model, chart),
    }
    reason = (
        f"横轴标题文本={repr(title)}，"
        f"文本匹配={'是' if title_ok else '否'}（仅允许“频率（GHz）”/“频率 (GHz)”）；"
        f"字体={ax.font_names or '默认(Calibri)'}，ok={'是' if font_name_ok else '否'}；"
        f"字号={ax.font_sizes or '默认'}pt，ok={'是' if font_size_ok else '否'}；"
        f"颜色={ax.font_colors or '默认(黑色)'}，ok={'是' if font_color_ok else '否'}；"
        f"底部居中={'是' if at_bottom_center else '否'}（{center_mode}）"
    )
    return RuleResult("P06", name, 1, ok, reason, evidence, model.path.name)


def check_p07(model: WorkbookModel) -> RuleResult:
    # 细则：图表左侧纵轴标题文本为"S11"，位于图表左侧中部，
    # 文字竖向显示或旋转90度，字体为微软雅黑、宋体或Calibri，字号10磅–14磅，颜色为黑色。
    name = '纵轴标题文本为"S11"，位于图表左侧中部，竖向显示或旋转90度，字体微软雅黑/宋体/Calibri，字号10-14磅，颜色黑色'
    chart = select_s11_chart(model)

    # ① 位于图表左侧纵轴：取 axPos=l 的坐标轴。
    ax = axis_by_pos(chart, "l")
    if not ax:
        return fail_result("P07", name, 1, "未找到左侧纵轴（axPos=l）", chart_identity(model, chart), model.path.name)

    title = ax.title

    # ② 标题文本为"S11"。
    nt = normalize_text(title)
    title_ok = nt == "s11"

    # ③ 位于左侧中部：axPos=l 已确认左侧；进一步用 c:title/c:layout/c:manualLayout
    #    判断标题在垂直方向是否位于中部（y 中心接近 0.5）：
    #    - 有显式手动布局：以标题框垂直中心 y+h/2 与 chartSpace 中心 0.5 比较，容差 ±0.20；
    #      若手动布局把标题明显推离垂直中部，则判为不在中部。
    #    - 无显式手动布局：Excel/WPS 默认将左侧纵轴标题垂直居中于坐标轴中部，按默认合规处理。
    title_layout = ax.title_layout
    if title_layout:
        cy = title_layout["y"] + title_layout["h"] / 2
        at_left_middle = abs(cy - 0.5) <= 0.20
        middle_mode = "manual"
    else:
        at_left_middle = True
        middle_mode = "auto(默认居中)"

    # ④ 文字竖向显示或旋转90度：
    #    OOXML 中两种表达方式任一满足即可：
    #    (a) a:bodyPr/@rot 单位 1/60000 度；±5400000 表示 ±90°；
    #    (b) a:bodyPr/@vert 表示竖排文本（非 "horz"、非缺省），如 vert/vert270/eaVert/wordArtVert；
    #    仅缺少 rot 不等于不合格——只要 vert 属性表明竖排也算满足。
    rotation = ax.title_rotation
    vert = ax.title_vert
    rot_90 = rotation is not None and abs(abs(rotation) - 5400000) <= 100000
    vert_ok = vert is not None and vert != "" and vert != "horz"
    vertical_or_90 = rot_90 or vert_ok

    # ⑤ 字体为微软雅黑、宋体或Calibri；未显式写出字体名时按 Office 默认 Calibri 处理。
    allowed_fonts = {"微软雅黑", "宋体", "Calibri", "+mn-lt", "+mj-lt"}
    if not ax.font_names:
        font_name_ok = True
    else:
        font_name_ok = any(n in allowed_fonts for n in ax.font_names)

    # ⑥ 字号10磅–14磅；未显式写出字号时按 Office 默认字号处理。
    if not ax.font_sizes:
        font_size_ok = True
    else:
        font_size_ok = all(10 <= s <= 14 for s in ax.font_sizes)

    # ⑦ 颜色为黑色；未显式写出颜色时按 Office 默认黑色处理。
    if not ax.font_colors:
        font_color_ok = True
    else:
        font_color_ok = all(color_is_blackish(c) for c in ax.font_colors)

    ok = bool(title_ok and at_left_middle and vertical_or_90 and font_name_ok and font_size_ok and font_color_ok)

    evidence = {
        "title": title,
        "normalized": nt,
        "title_ok": title_ok,
        "title_layout": title_layout,
        "middle_mode": middle_mode,
        "at_left_middle": at_left_middle,
        "rotation": rotation,
        "vert": vert,
        "rot_90": rot_90,
        "vert_ok": vert_ok,
        "vertical_or_90": vertical_or_90,
        "font_names": ax.font_names,
        "font_name_ok": font_name_ok,
        "font_sizes": ax.font_sizes,
        "font_size_ok": font_size_ok,
        "font_colors": ax.font_colors,
        "font_color_ok": font_color_ok,
        **chart_identity(model, chart),
    }
    reason = (
        f"纵轴标题文本={repr(title)}，文本匹配={'是' if title_ok else '否'}；"
        f"左侧中部={'是' if at_left_middle else '否'}（{middle_mode}）；"
        f"竖向或90°={'是' if vertical_or_90 else '否'}"
        f"（rot={rotation} 90°判定={'是' if rot_90 else '否'}，vert={vert} 竖排属性={'是' if vert_ok else '否'}）；"
        f"字体={ax.font_names or '默认(Calibri)'}，ok={'是' if font_name_ok else '否'}；"
        f"字号={ax.font_sizes or '默认'}pt，ok={'是' if font_size_ok else '否'}；"
        f"颜色={ax.font_colors or '默认(黑色)'}，ok={'是' if font_color_ok else '否'}"
    )
    return RuleResult("P07", name, 1, ok, reason, evidence, model.path.name)


def font_rule_ok(ax: AxisModel | None, min_pt: float, max_pt: float, allowed_names: set[str]) -> bool:
    if ax is None:
        return False
    sizes_ok = True if not ax.font_sizes else all(min_pt <= s <= max_pt for s in ax.font_sizes)
    names_ok = True if not ax.font_names else any(n in allowed_names for n in ax.font_names)
    colors_ok = True if not ax.font_colors else all(color_is_blackish(c) for c in ax.font_colors)
    return sizes_ok and names_ok and colors_ok


def arithmetic_ticks(start: float, step: float, stop: float, max_count: int = 500) -> list[float]:
    """按 start、step 生成到 stop（含）为止的等差刻度序列。"""
    if step <= 0 or stop < start:
        return []
    ticks: list[float] = []
    v = start
    n = 0
    while v <= stop + 1e-9 and n < max_count:
        ticks.append(round(v, 6))
        v += step
        n += 1
    return ticks


def sequences_match(seq: list[float], target: list[float], tol: float = 0.01) -> bool:
    """两组数值序列是否逐项相等（浮点容差）。"""
    if len(seq) != len(target):
        return False
    return all(abs(a - b) <= tol for a, b in zip(seq, target))


def infer_axis_tick_labels(ax: AxisModel | None, data_x: list[float], target: list[float]) -> tuple[bool, str]:
    """结合轴类型、源数据、min/max 与 majorUnit 推断横轴实际显示的刻度标签是否等于 target。

    细则允许"刻度间隔均匀"或"按源数据频率间隔显示"两种方式，因此不能只认显式 majorUnit=4：
    ① 显式 majorUnit：以轴 min/max（缺省回退到数据 min/max）按 majorUnit 生成等差刻度，与 target 比对；
    ② 分类轴（catAx）：刻度标签即各分类值本身，用去重升序的源 x 值与 target 比对；
    ③ 按源数据频率间隔：源 x 值去重升序后本身等于 target（无论 majorUnit 是否写出，
       值轴自动刻度或分类轴逐类显示都会呈现这些标签），视为满足。
    返回 (是否匹配, 命中方式说明)。
    """
    uniq = sorted(set(data_x))
    axis_min = ax.min_val if (ax and ax.min_val is not None) else (uniq[0] if uniq else None)
    axis_max = ax.max_val if (ax and ax.max_val is not None) else (uniq[-1] if uniq else None)

    # ① 显式 majorUnit 生成等差刻度。
    if ax is not None and ax.major_unit is not None and axis_min is not None and axis_max is not None:
        ticks = arithmetic_ticks(axis_min, ax.major_unit, axis_max)
        if sequences_match(ticks, target):
            return True, f"explicit-majorUnit={ax.major_unit}"

    # ②③ 分类轴 / 按源数据频率间隔：源 x 去重升序等于目标标签。
    if sequences_match(uniq, target):
        if ax is not None and ax.tag == "catAx":
            return True, "category-source"
        return True, "source-frequency"

    return False, "no-match"


def check_p08(model: WorkbookModel) -> RuleResult:
    # 细则：横轴刻度从最小频率值到最大频率值按递增顺序排列，刻度间隔均匀或按源数据频率间隔显示，
    # 刻度标签为"0、4、8、12、16、20、24"，刻度标签为黑色，字号8磅–12磅，
    # 图表左侧为低频、右侧为高频。
    name = "横轴刻度0/4/8/12/16/20/24按递增顺序排列，间隔均匀或按源频率显示，黑色8-12磅，左低右高"
    chart = select_s11_chart(model)
    main = main_s11_series(chart, model)
    ax = axis_by_pos(chart, "b")
    if not main:
        return fail_result("P08", name, 3, "未发现主系列", file=model.path.name)

    data_x = series_x_values(model, main, True)
    if not data_x:
        return fail_result("P08", name, 3, "无法解析横轴数据", chart_identity(model, chart), model.path.name)

    # ① 横轴刻度从最小频率值到最大频率值按递增顺序排列；
    # ② 图表左侧为低频、右侧为高频：轴方向为 minMax 或默认，且频率 x 值递增。
    x_increasing = is_increasing(data_x)
    orient_ok = ax is None or ax.orientation in {None, "minMax"}
    left_low_right_high = x_increasing and orient_ok

    # ③ 刻度标签为 0、4、8、12、16、20、24：细则允许"间隔均匀"或"按源数据频率间隔显示"，
    #    因此不再强制显式 majorUnit=4；改为结合轴类型/源 x 值/min-max/majorUnit 推断实际标签。
    #    缺省 majorUnit（自动轴或分类轴）只要源频率恰为 0/4/.../24，即按"源数据频率间隔"判定合格。
    target_labels = [0.0, 4.0, 8.0, 12.0, 16.0, 20.0, 24.0]
    tick_labels_ok, tick_mode = infer_axis_tick_labels(ax, data_x, target_labels)

    # ④ 刻度标签为黑色；未显式写出颜色时按 Excel/WPS 默认黑色处理。
    if ax is not None and ax.font_colors:
        font_color_ok = all(color_is_blackish(c) for c in ax.font_colors)
    else:
        font_color_ok = True

    # ⑤ 字号8磅–12磅；未显式写出字号时按 Excel/WPS 默认字号处理。
    if ax is not None and ax.font_sizes:
        font_size_ok = all(8 <= s <= 12 for s in ax.font_sizes)
    else:
        font_size_ok = True

    ok = bool(left_low_right_high and tick_labels_ok and font_color_ok and font_size_ok)
    evidence = {
        "x_min": min(data_x),
        "x_max": max(data_x),
        "x_count": len(data_x),
        "x_unique": sorted(set(data_x)),
        "x_increasing": x_increasing,
        "axis_tag": ax.tag if ax else None,
        "axis_min": ax.min_val if ax else None,
        "axis_max": ax.max_val if ax else None,
        "orientation": ax.orientation if ax else None,
        "left_low_right_high": left_low_right_high,
        "majorUnit": ax.major_unit if ax else None,
        "tick_mode": tick_mode,
        "tick_labels_ok": tick_labels_ok,
        "font_sizes": ax.font_sizes if ax else [],
        "font_size_ok": font_size_ok,
        "font_colors": ax.font_colors if ax else [],
        "font_color_ok": font_color_ok,
        **chart_identity(model, chart),
    }
    reason = (
        f"x范围={min(data_x)}–{max(data_x)}，递增={'是' if x_increasing else '否'}，"
        f"轴方向={ax.orientation if ax else '默认minMax'}，左低右高={'是' if left_low_right_high else '否'}；"
        f"刻度模式={tick_mode}（轴类型={ax.tag if ax else None}，majorUnit={ax.major_unit if ax else None}），"
        f"标签0/4/8/12/16/20/24={'是' if tick_labels_ok else '否'}；"
        f"字号={ax.font_sizes if ax and ax.font_sizes else '默认'}，ok={'是' if font_size_ok else '否'}；"
        f"颜色={ax.font_colors if ax and ax.font_colors else '默认黑色'}，ok={'是' if font_color_ok else '否'}"
    )
    return RuleResult("P08", name, 3, ok, reason, evidence, model.path.name)


def check_p09(model: WorkbookModel) -> RuleResult:
    # 细则：纵轴刻度覆盖源数据S11最小值和最大值，显示范围包含约-24到-8或与数据接近的负值区间，
    # 刻度标签为"-8、-12、-16、-20、-24"，标签为黑色，字号8磅–12磅，
    # 纵轴上方数值较大、下方数值较小。
    name = "纵轴刻度覆盖S11源数据最小/最大值，范围约-24到-8，刻度标签-8/-12/-16/-20/-24，黑色8-12磅，上大下小"
    chart = select_s11_chart(model)
    main = main_s11_series(chart, model)
    ax = axis_by_pos(chart, "l")
    if not main:
        return fail_result("P09", name, 3, "未发现主系列", file=model.path.name)
    yvals = series_y_values(model, main, True)
    if not yvals:
        return fail_result("P09", name, 3, "无法解析纵轴数据", chart_identity(model, chart), model.path.name)

    y_data_min = min(yvals)
    y_data_max = max(yvals)
    target_labels = [-8.0, -12.0, -16.0, -20.0, -24.0]  # 细则给定的负值刻度标签
    target_min = min(target_labels)  # -24
    target_max = max(target_labels)  # -8

    # ① 纵轴刻度覆盖源数据S11最小值和最大值，且显示范围包含"约-24到-8或与数据接近的负值区间"：
    #    细则给出了两种合格情形，判定不再依赖硬编码数据阈值（如原 y_min<=-16、y_max>=-10）：
    #    (a) OOXML 显式写出轴 min/max：轴范围须包含数据 [y_data_min, y_data_max]，
    #        且轴范围与目标 [-24, -8] 具有较高重叠度（认为"与数据接近的负值区间"）；
    #    (b) 未显式写出（自动范围）：Excel/WPS 会用数据范围自动扩展一小段外边距，
    #        视为覆盖数据；只要数据整体落在负值区间且与 [-24, -8] 有实质性重叠即算合格，
    #        避免对具体上下界做硬编码。
    axis_min = ax.min_val if ax is not None else None
    axis_max = ax.max_val if ax is not None else None

    def _overlap_ratio(lo: float, hi: float, tlo: float, thi: float) -> float:
        """两个区间的重叠长度占目标区间长度的比例（用于判断"接近目标区间"）。"""
        inter = min(hi, thi) - max(lo, tlo)
        span = thi - tlo
        if span <= 0:
            return 0.0
        return max(0.0, inter) / span

    if axis_min is not None and axis_max is not None:
        covers_data = axis_min <= y_data_min and axis_max >= y_data_max
        overlap = _overlap_ratio(axis_min, axis_max, target_min, target_max)
        range_ok = covers_data and overlap >= 0.6
        range_mode = "explicit"
    else:
        # 自动范围：默认覆盖数据（Excel/WPS 会自适应），只需判断数据本身是否位于合理负值区间，
        # 且与目标 [-24, -8] 有实质性重叠（重叠占目标长度 >=60%）——不再硬编码 -16/-10。
        covers_data = True
        in_negative = y_data_max <= 0 + 1e-6
        overlap = _overlap_ratio(y_data_min, y_data_max, target_min, target_max)
        range_ok = in_negative and overlap >= 0.6
        range_mode = "auto"

    # ② 刻度标签为 -8、-12、-16、-20、-24：细则允许"接近的负值区间"，
    #    因此不再强制显式 majorUnit=4；改为结合轴 min/max、majorUnit、数据自身推断实际标签，
    #    与横轴 P08 的做法一致（复用 infer_axis_tick_labels）。
    tick_ok, tick_mode = infer_axis_tick_labels(ax, yvals, target_labels)

    # ③ 纵轴上方数值较大、下方数值较小（minMax方向）：
    #    orientation 为 minMax 或未写出（默认即 minMax）。
    orient_ok = ax is None or ax.orientation in {None, "minMax"}

    # ④ 标签颜色为黑色；未显式写出时按 Excel/WPS 默认黑色处理。
    if ax is not None and ax.font_colors:
        font_color_ok = all(color_is_blackish(c) for c in ax.font_colors)
    else:
        font_color_ok = True

    # ⑤ 字号8磅–12磅；未显式写出时按 Excel/WPS 默认字号处理。
    if ax is not None and ax.font_sizes:
        font_size_ok = all(8 <= s <= 12 for s in ax.font_sizes)
    else:
        font_size_ok = True

    ok = bool(range_ok and tick_ok and orient_ok and font_color_ok and font_size_ok)
    evidence = {
        "axis_min": axis_min,
        "axis_max": axis_max,
        "y_data_min": y_data_min,
        "y_data_max": y_data_max,
        "target_range": [target_min, target_max],
        "overlap_ratio": overlap,
        "covers_data": covers_data,
        "range_mode": range_mode,
        "range_ok": range_ok,
        "axis_tag": ax.tag if ax else None,
        "majorUnit": ax.major_unit if ax else None,
        "tick_mode": tick_mode,
        "tick_ok": tick_ok,
        "orientation": ax.orientation if ax else None,
        "orient_ok": orient_ok,
        "font_sizes": ax.font_sizes if ax else [],
        "font_size_ok": font_size_ok,
        "font_colors": ax.font_colors if ax else [],
        "font_color_ok": font_color_ok,
        **chart_identity(model, chart),
    }
    reason = (
        f"数据范围={y_data_min}–{y_data_max}；"
        f"轴范围={axis_min}–{axis_max}（None=自动），"
        f"范围模式={range_mode}，与目标[-24,-8]重叠={overlap:.2f}，覆盖ok={'是' if range_ok else '否'}；"
        f"刻度模式={tick_mode}（轴类型={ax.tag if ax else None}，majorUnit={ax.major_unit if ax else None}），"
        f"刻度-8/-12/-16/-20/-24 ok={'是' if tick_ok else '否'}；"
        f"方向={ax.orientation if ax else '默认minMax'}，上大下小={'是' if orient_ok else '否'}；"
        f"字号={ax.font_sizes if ax and ax.font_sizes else '默认'}，ok={'是' if font_size_ok else '否'}；"
        f"颜色={ax.font_colors if ax and ax.font_colors else '默认黑色'}，ok={'是' if font_color_ok else '否'}"
    )
    return RuleResult("P09", name, 3, ok, reason, evidence, model.path.name)


def check_p10(model: WorkbookModel) -> RuleResult:
    # 细则：图表绘图区上方区域出现一条灰色水平虚线参考线，位置约对应S11=-10附近或任务参考图中的阈值位置，
    # 线宽1磅–1.5磅，线型为虚线或短划线，贯穿主要绘图区宽度横轴"0"-"24"。
    name = "灰色水平虚线参考线位于S11=-10附近，线宽约1-1.5磅，虚线/短划线，贯穿横轴0-24"
    chart = select_s11_chart(model)
    ref = reference_series(chart, model)
    if not ref:
        return fail_result("P10", name, 1, "未发现 -10 附近参考线系列", chart_identity(model, chart), model.path.name)

    line = ref.line
    xvals = series_x_values(model, ref)
    yvals = series_y_values(model, ref)
    width, width_inferred = effective_line_width(line)

    # ① 灰色水平虚线参考线：必须有参考线系列，且 y 值基本恒定。
    horizontal_ok = bool(yvals) and (max(yvals) - min(yvals) <= 0.2)

    # ② 位置约对应 S11=-10 附近或阈值位置：允许 ±0.25dB 的办公软件显示/录入误差。
    y_ok = bool(yvals) and sum(1 for v in yvals if abs(v + 10) <= 0.25) / max(1, len(yvals)) >= 0.8

    # ③ 贯穿主要绘图区宽度横轴 0–24：
    #    rubric 允许"横轴显式为 0–24"或"自动轴实际显示 0–24"两种合格情形。
    #    仅当横轴显式写出且末端明显超出 24（自动扩展到 25/30 等）才判为不贯穿；
    #    未显式写出（自动轴）时，Excel/WPS 会根据主系列 x 值自适应范围，
    #    只要主系列 x 覆盖 0–24 附近且参考线 x 与主系列同宽/更宽，即视为贯穿主要绘图区。
    x_axis = axis_by_pos(chart, "b")
    main = main_s11_series(chart, model)
    main_x = series_x_values(model, main, True) if main else []
    ref_covers_0_24 = bool(xvals and min(xvals) <= 0.01 and max(xvals) >= 23.99)

    if x_axis is not None and x_axis.min_val is not None and x_axis.max_val is not None:
        # 显式范围：允许 min≈0（≤0.01），max 在 [23.99, 24.5] 内视为贯穿；
        # 若显式 max>24.5，说明轴被拉长到 25/30 等，参考线不会贯穿整个绘图区。
        axis_range_ok = x_axis.min_val <= 0.01 and 23.99 <= x_axis.max_val <= 24.5
        axis_mode = "explicit"
    else:
        # 自动范围：结合主系列 x 范围推断实际绘图区宽度——
        # 主系列 x 覆盖 0–24（min≈0，max≈24），且参考线 x 范围与之匹配（min≤主系列min，max≥主系列max-1e-2）
        # 即认为参考线贯穿主要绘图区。若无主系列可参照，则退化到参考线自身是否覆盖 0–24。
        if main_x:
            main_min, main_max = min(main_x), max(main_x)
            axis_range_ok = (
                main_min <= 0.01
                and 23.99 <= main_max <= 24.5
                and bool(xvals)
                and min(xvals) <= main_min + 1e-6
                and max(xvals) + 1e-2 >= main_max
            )
            axis_mode = "auto-by-main-series"
        else:
            axis_range_ok = ref_covers_0_24
            axis_mode = "auto-fallback-ref"

    x_ok = ref_covers_0_24 and axis_range_ok

    # ④ 灰色：线条颜色为灰色。
    gray_ok = line_is_gray(line)

    # ⑤ 线型为虚线或短划线。
    dash_ok = line.get("dash") in {"dash", "dashDot", "sysDash", "lgDash", "sysDot", "dot"}

    # ⑥ 线宽1磅–1.5磅：Excel/WPS 内部用 EMU 存储，常见1pt会导出为约0.94pt，
    #    因此按办公软件有效性给 1pt 下限留 0.1pt 容差；上限仍按1.5pt控制。
    width_ok = width is not None and not width_inferred and 0.9 <= width <= 1.5

    ok = bool(horizontal_ok and y_ok and x_ok and gray_ok and dash_ok and width_ok)
    evidence = {
        "line": line,
        "width_pt": width,
        "width_inferred": width_inferred,
        "width_ok": width_ok,
        "dash_ok": dash_ok,
        "gray_ok": gray_ok,
        "horizontal_ok": horizontal_ok,
        "y_min": min(yvals) if yvals else None,
        "y_max": max(yvals) if yvals else None,
        "y_ok": y_ok,
        "x_min": min(xvals) if xvals else None,
        "x_max": max(xvals) if xvals else None,
        "x_axis_min": x_axis.min_val if x_axis else None,
        "x_axis_max": x_axis.max_val if x_axis else None,
        "main_x_min": min(main_x) if main_x else None,
        "main_x_max": max(main_x) if main_x else None,
        "axis_mode": axis_mode,
        "axis_range_ok": axis_range_ok,
        "ref_covers_0_24": ref_covers_0_24,
        "x_ok": x_ok,
        **chart_identity(model, chart),
    }
    reason = (
        f"水平={'是' if horizontal_ok else '否'}，y范围={min(yvals) if yvals else None}–{max(yvals) if yvals else None}，"
        f"-10附近={'是' if y_ok else '否'}；"
        f"灰色={'是' if gray_ok else '否'}（color={line.get('color') or line.get('scheme_color')}）；"
        f"虚线/短划线={'是' if dash_ok else '否'}（dash={line.get('dash')}）；"
        f"线宽={width:.2f}pt，ok={'是' if width_ok else '否'}；" if width is not None else "线宽未写出；"
    ) + (
        f"贯穿0-24={'是' if x_ok else '否'}（"
        f"参考线x={min(xvals) if xvals else None}–{max(xvals) if xvals else None}，"
        f"轴显式范围={x_axis.min_val if x_axis else None}–{x_axis.max_val if x_axis else None}，"
        f"主系列x={min(main_x) if main_x else None}–{max(main_x) if main_x else None}，"
        f"模式={axis_mode}）"
    )
    return RuleResult("P10", name, 1, ok, reason, evidence, model.path.name)


def check_p11(model: WorkbookModel) -> RuleResult:
    # 细则：绘图区背景为白色或无填充，边框为黑色或深灰色单实线，线宽0.75磅–1.25磅。
    name = "绘图区背景白色/无填充，边框黑色或深灰单实线0.75-1.25磅"
    chart = select_s11_chart(model)
    if not chart:
        return fail_result("P11", name, 1, "未发现图表", file=model.path.name)

    # ① 绘图区背景为白色或无填充：
    #    检查 plotArea/c:spPr 的填充；None 或 "none" 表示无填充，"FFFFFF" 为白色。
    #    Excel/WPS 未写 plotArea spPr 时，绘图区默认为无填充（透明），视为满足。
    fill_ok = chart.plot_fill in {None, "none", "FFFFFF"}

    # ② 边框为黑色或深灰色单实线，线宽0.75磅–1.25磅：
    #    Excel/WPS 的绘图区边框写在 plotArea/c:spPr/a:ln；
    #    若 plotArea 无显式边框（plot_line 为空），则检查 chartSpace 外框（chart_line），
    #    因为 P03 已确认 chartSpace 外框是"图表外侧矩形"，即视觉上的边框。
    #    边框颜色允许黑色或深灰色（D9D9D9 属于浅灰，即 color_is_gray 范围内，
    #    细则"深灰"对应 hex 约 555555–888888 范围，用 color_is_gray 宽松判定）。
    border_line = chart.plot_line if chart.plot_line else chart.chart_line
    if not border_line:
        # 无任何显式边框：Excel/WPS 默认绘图区无边框，仍按细则判断为无边框，不满足"有边框"。
        border_ok = False
        border_solid = False
        border_color_ok = False
        border_width_ok = False
    else:
        border_width = border_line.get("width_pt")
        border_color_ok = (
            color_is_blackish(border_line.get("color"))
            or color_is_gray(border_line.get("color"))
        )
        border_solid = border_line.get("dash") == "solid"
        border_width_ok = border_width is not None and 0.75 <= border_width <= 1.25
        border_ok = border_color_ok and border_solid and border_width_ok

    ok = bool(fill_ok and border_ok)
    evidence = {
        "plot_fill": chart.plot_fill,
        "fill_ok": fill_ok,
        "plot_line": chart.plot_line,
        "chart_line": chart.chart_line,
        "border_line_used": "plot_line" if chart.plot_line else ("chart_line" if chart.chart_line else "none"),
        "border_color_ok": border_color_ok,
        "border_solid": border_solid,
        "border_width_pt": border_line.get("width_pt") if border_line else None,
        "border_width_ok": border_width_ok,
        "border_ok": border_ok,
        **chart_identity(model, chart),
    }
    reason = (
        f"绘图区填充={chart.plot_fill or '默认无填充'}，白色/无填充={'是' if fill_ok else '否'}；"
        f"边框颜色黑/深灰={'是' if border_color_ok else '否'}，"
        f"单实线={'是' if border_solid else '否'}，"
        f"线宽={border_line.get('width_pt') if border_line else None}pt，"
        f"0.75-1.25磅={'是' if border_width_ok else '否'}"
    )
    return RuleResult("P11", name, 1, ok, reason, evidence, model.path.name)


def check_p12(model: WorkbookModel) -> RuleResult:
    # 细则：图表顶部没有出现与任务无关的默认标题"图表标题"；
    # 若设置标题，标题文本包含"S11参数曲线"或"S11曲线"，
    # 字体为微软雅黑或Calibri，字号16磅–20磅，颜色为黑色。
    name = '图表顶部无默认"图表标题"；若有标题则包含"S11参数曲线"或"S11曲线"，字体微软雅黑/Calibri，字号16-20磅，黑色'
    chart = select_s11_chart(model)
    if not chart:
        return fail_result("P12", name, 1, "未发现图表", file=model.path.name)

    title = chart.title or ""
    nt = normalize_text(title)

    # ① 没有出现默认标题"图表标题"。
    no_default = nt not in {"图表标题", "charttitle", "chart title"}

    if not title:
        # 无标题：不违反"无默认标题"要求，细则"若设置标题"条件不触发，视为满足。
        content_ok = True
        font_name_ok = True
        font_size_ok = True
        font_color_ok = True
    else:
        # ② 若设置了标题，标题文本包含"S11参数曲线"或"S11曲线"。
        content_ok = "s11参数曲线" in nt or "s11曲线" in nt

        # 以下字体/字号/颜色只取自 c:title 节点（chart.title_font_*），
        # 不再使用覆盖整个 chart XML 的 chart.font_*，避免混入坐标轴/图例的字体颜色导致误判。
        title_names = chart.title_font_names
        title_sizes_all = chart.title_font_sizes
        title_colors = chart.title_font_colors

        # ③ 字体为微软雅黑或Calibri：
        #    Excel/WPS 将图表标题字体写入 c:title 内的 a:rPr/a:latin typeface；
        #    若未写出字体，按 Office 默认（Calibri）处理，视为满足。
        allowed_title_fonts = {"微软雅黑", "Calibri", "+mn-lt", "+mj-lt"}
        if not title_names:
            font_name_ok = True
        else:
            font_name_ok = all(n in allowed_title_fonts for n in title_names)

        # ④ 字号16磅–20磅：
        #    Excel/WPS 将字号写入 a:rPr/@sz（单位：百分之一磅）；
        #    只取标题节点内的字号，若未写出按 Office 默认图表标题字号（通常18pt）处理，视为满足。
        if not title_sizes_all:
            font_size_ok = True
        else:
            font_size_ok = all(16 <= s <= 20 for s in title_sizes_all)

        # ⑤ 颜色为黑色；只看标题节点颜色，未显式写出时按 Office 默认黑色处理。
        if not title_colors:
            font_color_ok = True
        else:
            font_color_ok = all(color_is_blackish(c) for c in title_colors)

    ok = bool(no_default and content_ok and font_name_ok and font_size_ok and font_color_ok)
    evidence = {
        "title": title,
        "normalized": nt,
        "no_default": no_default,
        "content_ok": content_ok,
        "title_font_names": chart.title_font_names,
        "font_name_ok": font_name_ok,
        "title_font_sizes": chart.title_font_sizes,
        "font_size_ok": font_size_ok,
        "title_font_colors": chart.title_font_colors,
        "font_color_ok": font_color_ok,
        **chart_identity(model, chart),
    }
    reason = (
        f"标题={repr(title)}，无默认标题='{'是' if no_default else '否'}'；"
        f"含S11参数曲线/S11曲线='{'是' if content_ok else '否'}'；"
        f"标题字体={chart.title_font_names or '默认(Calibri)'}，ok='{'是' if font_name_ok else '否'}'；"
        f"标题字号={chart.title_font_sizes or '默认'}pt，ok='{'是' if font_size_ok else '否'}'；"
        f"标题颜色={chart.title_font_colors or '默认黑色'}，ok='{'是' if font_color_ok else '否'}'"
    )
    return RuleResult("P12", name, 1, ok, reason, evidence, model.path.name)


POSITIVE_CHECKS = [
    check_p01, check_p02, check_p03, check_p04, check_p05, check_p06,
    check_p07, check_p08, check_p09, check_p10, check_p11, check_p12,
]


# ───────────────────────────── 扣分规则 ─────────────────────────────
# 两条扣分规则(N06 表头缺失、N07 第3个工作簿数据行少于28行)已按需求删除。


# ───────────────────────────── 评估与报告 ─────────────────────────────

def evaluate_file(model: WorkbookModel) -> FileEvaluation:
    gates = [gate_file_open(model)]
    passed = all(r.matched for r in gates)
    if not passed:
        return FileEvaluation(model.path.name, False, gates, [], 0, "维度1未通过，按规则直接0分，不检查维度2")
    positives = [check(model) for check in POSITIVE_CHECKS]
    score = sum(r.points for r in positives if r.matched)
    return FileEvaluation(model.path.name, True, gates, positives, score)


def evaluate_all(files: list[Path]) -> dict[str, Any]:
    models = [parse_workbook(p) for p in files]
    file_evals = [evaluate_file(m) for m in models]

    deductions: list[RuleResult] = []

    positive_total = sum(ev.positive_score for ev in file_evals if ev.dimension1_passed)
    deduction_total = sum(r.points for r in deductions if r.matched)
    total_score = max(0, positive_total + deduction_total)
    return {
        "files": file_evals,
        "deductions": deductions,
        "positive_total": positive_total,
        "deduction_total": deduction_total,
        "total_score": total_score,
        "max_positive_score": 66,
    }


def result_to_jsonable(obj: Any) -> Any:
    if hasattr(obj, "__dataclass_fields__"):
        return asdict(obj)
    if isinstance(obj, Path):
        return str(obj)
    raise TypeError(type(obj).__name__)


SCRIPT_ID = "085"


def evaluate(dir_path: str) -> dict[str, Any]:
    """统一入口：接收脚本所在目录路径，脚本自己在该目录里定位三个目标 xlsx 并评估。

    返回结构遵循 `脚本接口差异与统一建议.md` §2.2 约定的字段。
    """
    file_name = "、".join(TARGET_FILES)
    try:
        base_dir = Path(dir_path)
        if not base_dir.is_dir():
            raise FileNotFoundError(f"目录不存在: {dir_path}")
        files = [base_dir / name for name in TARGET_FILES]
        missing = [path.name for path in files if not path.is_file()]
        if missing:
            raise FileNotFoundError("目录内缺少必需文档: " + ", ".join(missing))

        result = evaluate_all(files)
        file_evals = cast(list[FileEvaluation], result["files"])
        deductions = cast(list[RuleResult], result["deductions"])

        dim1_pass = bool(all(ev.dimension1_passed for ev in file_evals))
        if not dim1_pass:
            reasons: list[str] = []
            for ev in file_evals:
                for r in ev.gate_results:
                    if not r.matched:
                        reasons.append(f"{ev.file} {r.rule_id} {r.name}：{r.reason}")
            dim1_reason = "；".join(reasons)
        else:
            dim1_reason = ""

        dim2_items: list[dict[str, Any]] = []
        if dim1_pass:
            # 正向评分点：每个文件每条规则都作为一个 item 列出（命中和未命中均列出）。
            for ev in file_evals:
                for r in ev.positive_results:
                    dim2_items.append({
                        "rule": f"[{ev.file}] {r.rule_id} {r.name}",
                        "max_delta": r.points,
                        "delta": r.points if r.matched else 0,
                        "hit": bool(r.matched),
                        "detail": "",
                    })
            # 扣分点使用负数 max_delta；命中表示触发扣分条件。
            for r in deductions:
                dim2_items.append({
                    "rule": f"{r.rule_id} {r.name}",
                    "max_delta": r.points,
                    "delta": r.points if r.matched else 0,
                    "hit": bool(r.matched),
                    "detail": "",
                })

        return {
            "id": SCRIPT_ID,
            "file_name": file_name,
            "status": "ok",
            "error": None,
            "dim1_pass": dim1_pass,
            "dim1_reason": dim1_reason,
            "dim2_items": dim2_items,
            "total_score": int(result["total_score"]),
            "max_score": int(result["max_positive_score"]),
        }
    except Exception as exc:  # noqa: BLE001 - 顶层兜底，转成 status="error"
        return {
            "id": SCRIPT_ID,
            "file_name": file_name,
            "status": "error",
            "error": f"{type(exc).__name__}: {exc}",
            "dim1_pass": False,
            "dim1_reason": "",
            "dim2_items": [],
            "total_score": 0,
            "max_score": 0,
        }


if __name__ == "__main__":
    # 仅用于本地调试：默认使用脚本所在目录，也可通过 sys.argv[1] 指定其它目录。
    target = sys.argv[1] if len(sys.argv) > 1 else str(Path(__file__).parent)
    print(json.dumps(evaluate(target), ensure_ascii=False, indent=2, default=result_to_jsonable))
