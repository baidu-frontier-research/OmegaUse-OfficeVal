  # -*- coding: utf-8 -*-
"""
自动评估器 —— 评估"堆叠柱状图对比"作答 xlsx/xlsm 文件。

评分模型（与"打分细则"一致）：
  维度1（可用与可修改性，门槛项）：不满足 -> 直接 0 分，不再检查维度2。
  维度2（完成度）：得分点（满足全部子点才加分）+ 扣分点（满足任一子点即扣分）累加。

实现说明：
  - 仅依赖 Python 标准库（zipfile + xml）解析 xlsx 内部 XML，无需 openpyxl。
  - 对"不可编辑/截图代替"等通过结构判断：存在原生 chart part 即视为可编辑图表对象。
  - 颜色判断：取系列填充色，按 RGB 距离归类到 蓝/绿/橙。
  - 标签、坐标轴刻度等通过 chartXML 中对应节点判定。
  - 对难以严格量化的几何项（柱间距 0-0.1cm）用堆叠/重叠/间隙参数做近似推断，并在结果中说明。

对外接口：
  evaluate(dir_path: str) -> dict
    传入"脚本所在目录的路径"，函数自己在该目录内定位并打开被评估的 xlsx/xlsm 文档，
    返回结构化字典（含维度一通过与否、维度二逐项得分、总分等），
    不 print 主结果、不改 sys.stdout、不 sys.exit。总分为所有命中的加分项之和。

  本文件 __main__ 仅用于本地自测：
    python officeval_100_verifier.py <脚本所在目录>
"""

import sys
import os
import re
import json
import math
import zipfile
import xml.etree.ElementTree as ET

# ---------- XML 命名空间 ----------
NS = {
    "c": "http://schemas.openxmlformats.org/drawingml/2006/chart",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "ct": "http://schemas.openxmlformats.org/package/2006/content-types",
    "pr": "http://schemas.openxmlformats.org/package/2006/relationships",
    "s": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "xdr": "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
}

# 期望的资源类别与数据（来自原表与细则）
CATEGORIES = ["便携学习终端", "智能显示设备", "网络接入设备", "多功能文印设备", "移动采集设备"]
EXP_ZG = [6, 22, 18, 12, 5]        # 综管科 登记
EXP_FW = [9, 8, 6, 3, 2]           # 服务科 登记
EXP_TOTAL = [15, 30, 24, 15, 7]    # 合计登记
EXP_VERIFY = [13, 27, 22, 14, 6]   # 核验数量


# ============================================================
#  基础工具
# ============================================================
def q(tag):
    """把 'c:ser' 这种带前缀的标签转成 ElementTree 的 {ns}local 形式。"""
    pre, local = tag.split(":")
    return "{%s}%s" % (NS[pre], local)


def findall(el, path):
    return el.findall(path, NS)


def find(el, path):
    return el.find(path, NS)


class ZipBundle:
    """加载 xlsx 内部各 XML，提供按需访问。"""

    def __init__(self, path):
        self.path = path
        self.ok = False
        self.error = None
        self.names = []
        self._cache = {}
        try:
            self.zf = zipfile.ZipFile(path, "r")
            self.names = self.zf.namelist()
            self.ok = True
        except Exception as e:  # 损坏 / 非 zip / 打不开
            self.error = str(e)

    def read(self, name):
        if name in self._cache:
            return self._cache[name]
        try:
            data = self.zf.read(name)
        except KeyError:
            data = None
        self._cache[name] = data
        return data

    def read_text(self, name):
        data = self.read(name)
        if data is None:
            return None
        return data.decode("utf-8", errors="replace")

    def xml(self, name):
        data = self.read(name)
        if data is None:
            return None
        try:
            return ET.fromstring(data)
        except ET.ParseError:
            return None


def hex_to_rgb(h):
    h = h.strip().lstrip("#")
    if len(h) != 6:
        return None
    try:
        return (int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))
    except ValueError:
        return None


def classify_color(rgb):
    """把一个 RGB 颜色归类为 blue/green/orange/other（用主导通道+简单规则）。"""
    if rgb is None:
        return "unknown"
    r, g, b = rgb
    # 橙色：红高、绿中、蓝低
    if r >= 180 and 90 <= g <= 200 and b <= 110 and r > b:
        return "orange"
    # 蓝色：蓝为主导
    if b >= 130 and b >= r + 25 and b >= g:
        return "blue"
    # 绿色：绿为主导
    if g >= 120 and g >= r and g >= b - 10 and not (r >= 180 and b <= 110):
        return "green"
    # 兜底：按最大通道
    mx = max(r, g, b)
    if mx == b and b - max(r, g) > 15:
        return "blue"
    if mx == g and g - max(r, b) > 15:
        return "green"
    if mx == r and r - b > 60:
        return "orange"
    return "other"


def rgb_to_hex(rgb):
    if rgb is None:
        return None
    return "%02X%02X%02X" % tuple(max(0, min(255, int(round(c)))) for c in rgb)


DEFAULT_THEME_COLORS = {
    # Office 默认主题常见色；实际文件存在 xl/theme/theme1.xml 时会被真实主题覆盖。
    "dk1": (0, 0, 0),
    "lt1": (255, 255, 255),
    "dk2": (31, 78, 121),
    "lt2": (238, 236, 225),
    "accent1": (68, 114, 196),   # 蓝
    "accent2": (237, 125, 49),   # 橙
    "accent3": (112, 173, 71),   # 绿（兼容旧模板里把accent3用作绿色）
    "accent4": (255, 192, 0),
    "accent5": (91, 155, 213),
    "accent6": (112, 173, 71),   # 绿
    "hlink": (5, 99, 193),
    "folHlink": (149, 79, 114),
}

THEME_ALIASES = {
    "bg1": "lt1",
    "tx1": "dk1",
    "bg2": "lt2",
    "tx2": "dk2",
}


def _local_name(tag):
    return tag.rsplit("}", 1)[-1] if "}" in tag else tag


def _apply_color_transform(rgb, kind, val):
    try:
        factor = int(val) / 100000.0
    except (TypeError, ValueError):
        return rgb
    r, g, b = rgb
    if kind in ("lumMod", "shade"):
        return (r * factor, g * factor, b * factor)
    if kind in ("lumOff", "tint"):
        return (r + (255 - r) * factor,
                g + (255 - g) * factor,
                b + (255 - b) * factor)
    return rgb


def _apply_color_transforms(rgb, color_el):
    if rgb is None or color_el is None:
        return rgb
    out = rgb
    for child in list(color_el):
        kind = _local_name(child.tag)
        if kind in ("lumMod", "lumOff", "tint", "shade"):
            out = _apply_color_transform(out, kind, child.get("val"))
    return tuple(max(0, min(255, int(round(c)))) for c in out)


def load_theme_colors(bundle):
    """读取 workbook 主题色，返回 {schemeName: (r,g,b)}；取不到则用常见 Office 主题兜底。"""
    colors = dict(DEFAULT_THEME_COLORS)
    theme_paths = [n for n in bundle.names if re.search(r"^xl/theme/theme\d+\.xml$", n)]
    root = bundle.xml(theme_paths[0]) if theme_paths else None
    clr_scheme = root.find(".//a:clrScheme", NS) if root is not None else None
    if clr_scheme is not None:
        for node in list(clr_scheme):
            key = _local_name(node.tag)
            srgb = find(node, "a:srgbClr")
            sysclr = find(node, "a:sysClr")
            raw = None
            if srgb is not None and srgb.get("val"):
                raw = srgb.get("val")
            elif sysclr is not None and sysclr.get("lastClr"):
                raw = sysclr.get("lastClr")
            rgb = hex_to_rgb(raw) if raw else None
            if rgb is not None:
                colors[key] = rgb
    for alias, target in THEME_ALIASES.items():
        if target in colors:
            colors[alias] = colors[target]
    return colors


# ============================================================
#  读取工作簿结构（sheet 名、Sheet1 单元格、共享字符串）
# ============================================================
def load_shared_strings(bundle):
    root = bundle.xml("xl/sharedStrings.xml")
    out = []
    if root is None:
        return out
    for si in findall(root, "s:si"):
        # 拼接 si 下所有 <t> 文本（含富文本 run）
        texts = [t.text or "" for t in si.iter(q("s:t"))]
        out.append("".join(texts))
    return out


def get_sheet_targets(bundle):
    """返回 {sheet显示名: worksheet xml 路径}。"""
    wb = bundle.xml("xl/workbook.xml")
    rels = bundle.xml("xl/_rels/workbook.xml.rels")
    if wb is None or rels is None:
        return {}
    rid_to_target = {}
    for rel in findall(rels, "pr:Relationship"):
        rid_to_target[rel.get("Id")] = rel.get("Target")
    result = {}
    for sh in findall(wb, "s:sheets/s:sheet"):
        name = sh.get("name")
        rid = sh.get("{%s}id" % NS["r"])
        tgt = rid_to_target.get(rid, "")
        if tgt:
            # Target 可能是 '/xl/worksheets/sheet1.xml'（绝对，含 xl）
            # 或 'worksheets/sheet1.xml'（相对于 xl/）
            t = tgt.lstrip("/")
            if t.startswith("xl/"):
                tgt = t
            else:
                tgt = "xl/" + t
        result[name] = tgt
    return result


def col_letter_to_idx(letters):
    n = 0
    for ch in letters:
        n = n * 26 + (ord(ch.upper()) - ord("A") + 1)
    return n


def parse_cell_ref(ref):
    m = re.match(r"([A-Za-z]+)(\d+)", ref)
    if not m:
        return None, None
    return col_letter_to_idx(m.group(1)), int(m.group(2))


def read_sheet_cells(bundle, sheet_path, shared):
    """读取一个 worksheet 的单元格 {(col,row): value_str} 与合并单元格列表。"""
    root = bundle.xml(sheet_path)
    cells = {}
    merges = []
    if root is None:
        return cells, merges
    for c in root.iter(q("s:c")):
        ref = c.get("r")
        if not ref:
            continue
        col, row = parse_cell_ref(ref)
        t = c.get("t")
        v_el = find(c, "s:v")
        is_el = find(c, "s:is")
        val = None
        if t == "s" and v_el is not None and v_el.text is not None:
            idx = int(v_el.text)
            val = shared[idx] if 0 <= idx < len(shared) else ""
        elif t == "inlineStr" and is_el is not None:
            val = "".join(x.text or "" for x in is_el.iter(q("s:t")))
        elif v_el is not None:
            val = v_el.text
        cells[(col, row)] = "" if val is None else str(val)
    mc = find(root, "s:mergeCells")
    if mc is not None:
        for m in findall(mc, "s:mergeCell"):
            merges.append(m.get("ref"))
    return cells, merges


# ============================================================
#  解析图表（chart part）
# ============================================================
def find_chart_paths(bundle):
    """从 [Content_Types].xml 找出所有 chartN.xml。"""
    paths = []
    ct = bundle.xml("[Content_Types].xml")
    if ct is not None:
        for ov in findall(ct, "ct:Override"):
            pn = ov.get("PartName", "")
            if "/charts/chart" in pn and pn.endswith(".xml"):
                paths.append(pn.lstrip("/"))
    if not paths:
        # 兜底：直接按命名扫描
        paths = [n for n in bundle.names
                 if re.search(r"charts/chart\d+\.xml$", n)]
    return sorted(set(paths))


def _norm_part(target, base_dir):
    """把 rels 里的 Target 规整成包内绝对路径（去掉前导/，相对路径接 base_dir）。"""
    if not target:
        return ""
    t = target.replace("\\", "/")
    if t.startswith("/"):
        return t.lstrip("/")
    # 相对路径：相对 base_dir 解析（处理 ../）
    parts = (base_dir.split("/") + t.split("/")) if base_dir else t.split("/")
    stack = []
    for p in parts:
        if p in ("", "."):
            continue
        if p == "..":
            if stack:
                stack.pop()
        else:
            stack.append(p)
    return "/".join(stack)


def find_chart_host_sheets(bundle):
    """
    反查每个 worksheet 通过 drawing 关联了哪些 chart。
    返回 [{sheet_path, drawing_path, charts:[chartPath...]}], 供遮挡检测定位
    "图表所在的那个 sheet"。
    """
    result = []
    for name in bundle.names:
        m = re.match(r"xl/worksheets/(sheet\d+\.xml)$", name)
        if not m:
            continue
        sheet_path = name
        rels_path = "xl/worksheets/_rels/%s.rels" % m.group(1)
        rels = bundle.xml(rels_path)
        if rels is None:
            continue
        drawing_targets = []
        for rel in findall(rels, "pr:Relationship"):
            if rel.get("Type", "").endswith("/drawing"):
                drawing_targets.append(
                    _norm_part(rel.get("Target", ""), "xl/worksheets"))
        for dpath in drawing_targets:
            ddir = dpath.rsplit("/", 1)[0] if "/" in dpath else ""
            drels = bundle.xml("%s/_rels/%s.rels" % (
                ddir, dpath.rsplit("/", 1)[-1]))
            chart_parts = []
            if drels is not None:
                for rel in findall(drels, "pr:Relationship"):
                    if rel.get("Type", "").endswith("/chart"):
                        chart_parts.append(
                            _norm_part(rel.get("Target", ""), ddir))
            if chart_parts:
                result.append({
                    "sheet_path": sheet_path,
                    "drawing_path": dpath,
                    "charts": chart_parts,
                })
    return result


EMU_PER_CM = 360000.0
EMU_PER_COL = 640080   # 默认列宽 ~8.43 字符 ≈ 0.7in 的近似 EMU
EMU_PER_ROW = 190500   # 默认行高 15pt = 0.2in 的 EMU


def _xml_child_int(parent, path, default=0):
    el = find(parent, path) if parent is not None else None
    if el is None or el.text is None:
        return default
    try:
        return int(el.text)
    except (TypeError, ValueError):
        return default


def _anchor_extent_emu(anchor):
    """从 drawing anchor 取图表尺寸（EMU）。优先读 xdr:ext；twoCellAnchor 则按默认行列尺寸估算。"""
    ext = find(anchor, "xdr:ext")
    if ext is not None and ext.get("cx") and ext.get("cy"):
        try:
            return int(ext.get("cx")), int(ext.get("cy")), "xdr:ext"
        except (TypeError, ValueError):
            pass

    from_el = find(anchor, "xdr:from")
    to_el = find(anchor, "xdr:to")
    if from_el is None or to_el is None:
        return None, None, "无xdr:ext/to"
    c1 = _xml_child_int(from_el, "xdr:col")
    c1o = _xml_child_int(from_el, "xdr:colOff")
    r1 = _xml_child_int(from_el, "xdr:row")
    r1o = _xml_child_int(from_el, "xdr:rowOff")
    c2 = _xml_child_int(to_el, "xdr:col")
    c2o = _xml_child_int(to_el, "xdr:colOff")
    r2 = _xml_child_int(to_el, "xdr:row")
    r2o = _xml_child_int(to_el, "xdr:rowOff")
    cx = (c2 - c1) * EMU_PER_COL + (c2o - c1o)
    cy = (r2 - r1) * EMU_PER_ROW + (r2o - r1o)
    if cx <= 0 or cy <= 0:
        return None, None, "twoCellAnchor尺寸无效"
    return cx, cy, "twoCellAnchor默认行列估算"


def chart_anchor_extent(bundle, chart_path):
    """按 drawing rels 将 chart part 定位到对应锚点，返回尺寸与来源说明。"""
    target_chart = (chart_path or "").replace("\\", "/").lstrip("/")
    candidates = []
    for name in bundle.names:
        if not re.search(r"drawings/drawing\d+\.xml$", name):
            continue
        ddir = name.rsplit("/", 1)[0] if "/" in name else ""
        rels = bundle.xml("%s/_rels/%s.rels" % (ddir, name.rsplit("/", 1)[-1]))
        if rels is None:
            continue
        rid_to_chart = {}
        for rel in findall(rels, "pr:Relationship"):
            if rel.get("Type", "").endswith("/chart"):
                rid_to_chart[rel.get("Id")] = _norm_part(rel.get("Target", ""), ddir)
        if not rid_to_chart:
            continue
        root = bundle.xml(name)
        if root is None:
            continue
        anchors = []
        for tag in ("xdr:twoCellAnchor", "xdr:oneCellAnchor", "xdr:absoluteAnchor"):
            anchors.extend(findall(root, tag))
        for anchor in anchors:
            for chart_ref in anchor.findall(".//c:chart", NS):
                rid = chart_ref.get(q("r:id"))
                cpath = rid_to_chart.get(rid)
                if not cpath:
                    continue
                cx, cy, source = _anchor_extent_emu(anchor)
                item = {
                    "chart_path": cpath,
                    "drawing_path": name,
                    "cx": cx,
                    "cy": cy,
                    "source": source,
                }
                if target_chart and cpath == target_chart:
                    return item
                candidates.append(item)
    if not target_chart and len(candidates) == 1:
        return candidates[0]
    if target_chart:
        for item in candidates:
            if item["chart_path"] == target_chart:
                return item
    if len(candidates) == 1:
        return candidates[0]
    return None


def sheet_used_range(bundle, sheet_path, shared):
    """
    读取一个 worksheet 的"已用区域"：返回 (max_col, max_row)（1-based），
    依据是有值或有边框样式（s 属性，近似当作有内容/边框）的单元格。
    无内容则返回 (0, 0)。
    """
    cells, _merges = read_sheet_cells(bundle, sheet_path, shared)
    max_col = max_row = 0
    for (c, r), v in cells.items():
        if v != "":
            if c > max_col:
                max_col = c
            if r > max_row:
                max_row = r
    return max_col, max_row



def series_color_hex(ser, theme_colors=None):
    """取系列填充色，解析 srgbClr / schemeClr(+tint/shade/lum*) 为实际 RGB hex。"""
    sppr = find(ser, "c:spPr")
    if sppr is None:
        return None
    fill = find(sppr, "a:solidFill")
    if fill is None:
        return None
    clr = find(fill, "a:srgbClr")
    if clr is not None and clr.get("val"):
        rgb = hex_to_rgb(clr.get("val"))
        return rgb_to_hex(_apply_color_transforms(rgb, clr)) if rgb else None
    schemeclr = find(fill, "a:schemeClr")
    if schemeclr is not None and schemeclr.get("val"):
        theme_colors = theme_colors or DEFAULT_THEME_COLORS
        key = schemeclr.get("val")
        key = THEME_ALIASES.get(key, key)
        rgb = theme_colors.get(key)
        return rgb_to_hex(_apply_color_transforms(rgb, schemeclr)) if rgb else None
    return None


def num_vals(parent_val_el):
    """从 c:val 节点提取数值列表（numRef/numLit 均支持），按 idx 排序补 0。"""
    if parent_val_el is None:
        return []
    container = find(parent_val_el, "c:numRef/c:numCache")
    if container is None:
        container = find(parent_val_el, "c:numLit")
    if container is None:
        return []
    cnt_el = find(container, "c:ptCount")
    cnt = int(cnt_el.get("val")) if cnt_el is not None else 0
    arr = [0.0] * cnt
    for pt in findall(container, "c:pt"):
        idx = int(pt.get("idx"))
        v = find(pt, "c:v")
        if v is not None and v.text not in (None, ""):
            try:
                if 0 <= idx < cnt:
                    arr[idx] = float(v.text)
            except ValueError:
                pass
    return arr


def ref_formula(parent_el):
    """从 c:val/c:cat 节点下取 numRef 或 strRef 里的 c:f 公式引用文本
    （形如 'Sheet1!$D$3:$D$17'），取不到则返回 None。"""
    if parent_el is None:
        return None
    for path in ("c:numRef/c:f", "c:strRef/c:f"):
        f_el = find(parent_el, path)
        if f_el is not None and f_el.text:
            return f_el.text.strip()
    return None


def parse_sheet_ref(formula):
    """解析形如 'Sheet1!$D$3:$D$17' 或 "'Sheet 1'!D3:D17" 的公式引用，
    返回 (sheet_name, col_start, row_start, col_end, row_end)；解析失败返回 None。
    列/行以数字（1-based）表示；单值引用（无冒号）col_end/row_end 与起点相同。"""
    if not formula:
        return None
    m = re.match(r"^(?:'([^']+)'|([^'!]+))!(.+)$", formula.strip())
    if not m:
        return None
    sheet_name = m.group(1) if m.group(1) is not None else m.group(2)
    rng = m.group(3).replace("$", "")
    parts = rng.split(":")
    cell_re = re.compile(r"^([A-Za-z]+)(\d+)$")
    m1 = cell_re.match(parts[0])
    if not m1:
        return None
    col_start = col_letter_to_idx(m1.group(1))
    row_start = int(m1.group(2))
    if len(parts) > 1:
        m2 = cell_re.match(parts[1])
        if not m2:
            return None
        col_end = col_letter_to_idx(m2.group(1))
        row_end = int(m2.group(2))
    else:
        col_end, row_end = col_start, row_start
    return sheet_name, col_start, row_start, col_end, row_end


INNER_DLBL_POS = ("ctr", "inBase", "inEnd")
OUTER_DLBL_POS = ("outEnd",)


def _dlbl_text(dlbl_el):
    """取单个 c:dLbl 的自定义文本（c:tx/c:rich 下所有 a:t 拼接），
    没有自定义文本（走默认显示值）则返回 None。"""
    tx = find(dlbl_el, "c:tx")
    if tx is None:
        return None
    texts = [t.text or "" for t in tx.iter(q("a:t"))]
    if not texts:
        return None
    return "".join(texts)


def series_dlbl_info(ser, vals):
    """解析一个系列的数据标签信息：
    - pos: 标签位置（c:dLbls/c:dLblPos 的 val，取不到则 None）；
    - point_pos: 每个数据点各自的位置覆盖（c:dLbl/c:dLblPos），按 idx -> val；
    - point_text: 每个数据点实际显示的文本；若该点有自定义 c:tx 文本则用它，
      否则若整体/该点 showVal 开启则用该点数值（保留原始数值精度的字符串），
      标签被单独删除（c:delete）或未开启显示则为 None。
    """
    dlbls = find(ser, "c:dLbls")
    pos = None
    show_val = False
    point_pos = {}
    point_deleted = set()
    point_text = {}
    if dlbls is not None:
        dp = find(dlbls, "c:dLblPos")
        pos = dp.get("val") if dp is not None else None
        sv = find(dlbls, "c:showVal")
        show_val = sv is not None and sv.get("val") in ("1", "true")
        for dlbl in findall(dlbls, "c:dLbl"):
            idx_el = find(dlbl, "c:idx")
            if idx_el is None or idx_el.get("val") is None:
                continue
            try:
                idx = int(idx_el.get("val"))
            except ValueError:
                continue
            del_el = find(dlbl, "c:delete")
            if del_el is not None and del_el.get("val") in ("1", "true"):
                point_deleted.add(idx)
                continue
            dp_pt = find(dlbl, "c:dLblPos")
            if dp_pt is not None and dp_pt.get("val"):
                point_pos[idx] = dp_pt.get("val")
            custom_text = _dlbl_text(dlbl)
            if custom_text is not None:
                point_text[idx] = custom_text
            else:
                pt_show_val = find(dlbl, "c:showVal")
                if (pt_show_val is not None and pt_show_val.get("val") in ("1", "true")
                        and idx < len(vals)):
                    point_text[idx] = "%g" % vals[idx]
    # 每个数据点最终显示的文本：自定义文本/单点 showVal 优先；否则若未被单独删除且
    # （该系列整体 showVal 开启），用该点的实际数值（去掉多余的浮点小数位）。
    resolved_text = {}
    for idx, v in enumerate(vals):
        if idx in point_deleted:
            continue
        if idx in point_text:
            resolved_text[idx] = point_text[idx]
        elif show_val:
            resolved_text[idx] = ("%g" % v) if v is not None else None
    return {
        "pos": pos,
        "point_pos": point_pos,
        "point_deleted": point_deleted,
        "text": resolved_text,
    }


def str_vals(cat_el):
    """从 c:cat 节点提取分类文本列表。"""
    if cat_el is None:
        return []
    container = find(cat_el, "c:strRef/c:strCache")
    if container is None:
        container = find(cat_el, "c:strLit")
    if container is None:
        container = find(cat_el, "c:numRef/c:numCache")
    if container is None:
        container = find(cat_el, "c:numLit")
    if container is None:
        return []
    cnt_el = find(container, "c:ptCount")
    cnt = int(cnt_el.get("val")) if cnt_el is not None else 0
    arr = [""] * cnt
    for pt in findall(container, "c:pt"):
        idx = int(pt.get("idx"))
        v = find(pt, "c:v")
        if v is not None and 0 <= idx < cnt:
            arr[idx] = v.text or ""
    return arr


def parse_chart(root, theme_colors=None):
    """把一个 chart xml 解析成结构化 dict。"""
    info = {
        "raw": ET.tostring(root, encoding="unicode"),
        "path": None,
        "title": "",
        "bar_dir": None,
        "grouping": None,
        "gap_width": None,
        "overlap": None,
        "plot_manual_w": None,
        "plot_layout_target": None,
        "series": [],          # [{name, color_hex, color_class, vals, cats, formulas, has_dlbl}]
        "legend_pos": None,
        "cat_axis_present": False,
        "val_axis_present": False,
        "cat_axis_labels": [],
        "cat_axis_rot": None,
        "cat_axis_pos": None,
        "cat_axis_lbl_pos": None,
        "val_axis_title": "",
        "val_axis_min": None,
        "val_axis_max": None,
        "val_axis_unit": None,
        "val_axis_num_cache": [],
    }
    # 标题
    title_el = find(root, "c:chart/c:title")
    if title_el is not None:
        info["title"] = "".join(t.text or "" for t in title_el.iter(q("a:t")))

    bar = find(root, "c:chart/c:plotArea/c:barChart")
    if bar is not None:
        bd = find(bar, "c:barDir")
        info["bar_dir"] = bd.get("val") if bd is not None else None
        gr = find(bar, "c:grouping")
        info["grouping"] = gr.get("val") if gr is not None else None
        gw = find(bar, "c:gapWidth")
        info["gap_width"] = int(gw.get("val")) if gw is not None and gw.get("val") else None
        ov = find(bar, "c:overlap")
        info["overlap"] = int(ov.get("val")) if ov is not None and ov.get("val") else None
        for ser in findall(bar, "c:ser"):
            tx = find(ser, "c:tx")
            name = ""
            if tx is not None:
                name = "".join(t.text or "" for t in tx.iter(q("c:v"))) or \
                       "".join(t.text or "" for t in tx.iter(q("a:t")))
            chx = series_color_hex(ser, theme_colors)
            rgb = hex_to_rgb(chx) if chx and re.fullmatch(r"[0-9A-Fa-f]{6}", chx) else None
            dlbls = find(ser, "c:dLbls")
            has_dlbl = False
            if dlbls is not None:
                sv = find(dlbls, "c:showVal")
                has_dlbl = sv is not None and sv.get("val") in ("1", "true")
            val_el = find(ser, "c:val")
            cat_el = find(ser, "c:cat")
            vals = num_vals(val_el)
            dlbl_info = series_dlbl_info(ser, vals)
            has_dlbl = has_dlbl or bool(dlbl_info["text"])
            info["series"].append({
                "name": name,
                "color_hex": chx,
                "color_class": classify_color(rgb),
                "vals": vals,
                "cats": str_vals(cat_el),
                "val_formula": ref_formula(val_el),
                "cat_formula": ref_formula(cat_el),
                "has_dlbl": has_dlbl,
                "dlbl_pos": dlbl_info["pos"],
                "dlbl_point_pos": dlbl_info["point_pos"],
                "dlbl_text": dlbl_info["text"],
            })
        # 图表级 dLbls（整体显示数值；位置若系列级未单独指定，沿用图表级 dLblPos）
        chart_dlbls = find(bar, "c:dLbls")
        if chart_dlbls is not None:
            chart_pos_el = find(chart_dlbls, "c:dLblPos")
            chart_pos = chart_pos_el.get("val") if chart_pos_el is not None else None
            sv = find(chart_dlbls, "c:showVal")
            if sv is not None and sv.get("val") in ("1", "true"):
                series_list = info["series"] if isinstance(info["series"], list) else []
                for s in series_list:
                    s["has_dlbl"] = True
                    if s["dlbl_pos"] is None:
                        s["dlbl_pos"] = chart_pos
                    vals = s["vals"]
                    text_map = s["dlbl_text"]
                    for idx, v in enumerate(vals):
                        if idx not in text_map:
                            text_map[idx] = "%g" % v

    # 图例
    legend = find(root, "c:chart/c:plotArea/c:legend") or find(root, "c:chart/c:legend")
    if legend is not None:
        lp = find(legend, "c:legendPos")
        info["legend_pos"] = lp.get("val") if lp is not None else "present"
        # 图例条目删除项：c:legendEntry/c:idx 对应系列顺序索引（0-based），
        # c:delete val="1" 表示该系列条目被从图例中隐藏（不显示）。
        deleted_idx = set()
        for entry in findall(legend, "c:legendEntry"):
            idx_el = find(entry, "c:idx")
            del_el = find(entry, "c:delete")
            if idx_el is None or idx_el.get("val") is None:
                continue
            if del_el is not None and del_el.get("val") in ("1", "true"):
                try:
                    deleted_idx.add(int(idx_el.get("val")))
                except ValueError:
                    pass
        info["legend_deleted_idx"] = deleted_idx
    else:
        info["legend_deleted_idx"] = set()

    # plotArea 手动布局（c:manualLayout）：若存在，可精确得到绘图区占图表区的
    # 宽度比例（w，fraction），用于把柱间距从"占柱宽百分比"换算到 EMU/厘米。
    # 无手动布局时 Excel 走自动布局算法，XML 里没有该信息，只能退回近似。
    plot_area = find(root, "c:chart/c:plotArea")
    if plot_area is not None:
        layout = find(plot_area, "c:layout")
        manual = find(layout, "c:manualLayout") if layout is not None else None
        if manual is not None:
            w_el = find(manual, "c:w")
            target_el = find(manual, "c:layoutTarget")
            if w_el is not None and w_el.get("val"):
                try:
                    info["plot_manual_w"] = float(w_el.get("val"))
                except ValueError:
                    pass
            info["plot_layout_target"] = (
                target_el.get("val") if target_el is not None else None)

    # 坐标轴
    cat_ax = find(root, "c:chart/c:plotArea/c:catAx")
    if cat_ax is not None:
        info["cat_axis_present"] = True
        info["cat_axis_labels"] = str_vals(find(cat_ax, "c:cat"))
        # 轴位置（b=bottom/底部，须为底部才谈得上"标签位于柱形组合下方"）
        ax_pos = find(cat_ax, "c:axPos")
        info["cat_axis_pos"] = ax_pos.get("val") if ax_pos is not None else None
        # 标签相对轴线的位置（low/high/nextTo；nextTo=紧贴轴线，即柱形组合下方；
        # low/high 会把标签挪到绘图区边缘，不是"柱形组合下方"）
        tick_lbl_pos = find(cat_ax, "c:tickLblPos")
        info["cat_axis_lbl_pos"] = (
            tick_lbl_pos.get("val") if tick_lbl_pos is not None else None)
        # 轴文字旋转
        txpr = find(cat_ax, "c:txPr")
        if txpr is not None:
            bodypr = find(txpr, "a:bodyPr")
            if bodypr is not None and bodypr.get("rot"):
                info["cat_axis_rot"] = int(bodypr.get("rot"))
    val_ax = find(root, "c:chart/c:plotArea/c:valAx")
    if val_ax is not None:
        info["val_axis_present"] = True
        tt = find(val_ax, "c:title")
        if tt is not None:
            info["val_axis_title"] = "".join(t.text or "" for t in tt.iter(q("a:t")))
        # 数值轴刻度范围（scaling/min、scaling/max）
        scaling = find(val_ax, "c:scaling")
        if scaling is not None:
            mn = find(scaling, "c:min")
            if mn is not None and mn.get("val"):
                try:
                    info["val_axis_min"] = float(mn.get("val"))
                except ValueError:
                    pass
            mx = find(scaling, "c:max")
            if mx is not None and mx.get("val"):
                try:
                    info["val_axis_max"] = float(mx.get("val"))
                except ValueError:
                    pass
        # 主刻度间隔（majorUnit）
        mu = find(val_ax, "c:majorUnit")
        if mu is not None and mu.get("val"):
            try:
                info["val_axis_unit"] = float(mu.get("val"))
            except ValueError:
                pass
        # 数值轴系列缓存值（c:numCache，Excel 打开并保存过后会写入实际渲染用到
        # 的数据点数值——不是刻度值，但用于在 c:max 缺失（自动最大值）时，
        # 结合系列真实最大值倒推 Excel 大概率会自动取整到的坐标上限）。
        num_cache_vals = []
        for ser_val in root.iter(q("c:val")):
            num_cache_vals.extend(num_vals(ser_val))
        info["val_axis_num_cache"] = num_cache_vals

    return info


PLACEHOLDER = "<<APPEND>>"


# ============================================================
#  辅助匹配：把 chart 中 5 类资源的 综管科/服务科/核验 值对齐到顺序
# ============================================================
def map_series_values_by_category(chart):
    """
    cat 形如 '便携学习终端 登记' / '便携学习终端 核验'（10项）；
    series 名含 综管科/服务科/核验数量。
    返回 {'zg':[5],'fw':[5],'verify':[5]}（按 CATEGORIES 顺序），缺则 None。
    """
    out = {"zg": [None] * 5, "fw": [None] * 5, "verify": [None] * 5}

    def cat_index(t):
        for i, name in enumerate(CATEGORIES):
            if name in t:
                return i
        return None

    def role_of(n):
        if "综管" in n:
            return "zg"
        if "服务" in n:
            return "fw"
        if "核验" in n:
            return "verify"
        return None

    for s in chart["series"]:
        role = role_of(s["name"])
        if role is None:
            continue
        for j, cv in enumerate(s["cats"]):
            ci = cat_index(cv)
            if ci is None or j >= len(s["vals"]):
                continue
            v = s["vals"][j]
            if v and v != 0:
                out[role][ci] = v
    return out


def count_match(got, expected, tol=0.5):
    return sum(1 for g, e in zip(got, expected)
               if g is not None and abs(g - e) <= tol)


# ============================================================
#  维度2：完成度（得分点 + 扣分点）
# ============================================================
def eval_dimension2(bundle, sheet_targets, sheet1_cells, sheet1_raw_text,
                    charts, drawing_text, sheet1_merges, shared):
    """返回 (total, hits)；hits=[(score, hit_bool, label, detail)]。"""
    hits = []

    # 选主图：第一个二维堆叠柱图
    main = None
    for ch in charts:
        if ch["bar_dir"] == "col" and ch["grouping"] == "stacked":
            main = ch
            break
    if main is None and charts:
        main = charts[0]

    mapped = map_series_values_by_category(main) if main else \
        {"zg": [None] * 5, "fw": [None] * 5, "verify": [None] * 5}

    # ---------- +3 图表对象 ----------
    # 细则：新增1个可编辑图表；类型为二维堆叠柱状图/柱形图；
    #       用于比较"综管科+服务科+合计的登记数量"与"实际核验数量"；
    #       图表覆盖5类资源（便携学习终端/智能显示设备/网络接入设备/
    #       多功能文印设备/移动采集设备）；数据来自原表3-17行的登记数量与核验数量。
    def chk_chart_object():
        # 新增1个可编辑图表（存在原生 chart part 即视为可编辑），且数量恰好为1个：
        # 排除 percentStacked（百分比堆叠不是"二维堆叠柱状图/柱形图"），
        # 只统计 barDir=col 且 grouping=stacked 的图表数量。
        qualifying = [ch for ch in charts
                      if ch["bar_dir"] == "col" and ch["grouping"] == "stacked"]
        count_ok = len(qualifying) == 1
        if main is None or not count_ok:
            return False, "无可编辑图表对象" if main is None else \
                "二维堆叠柱状图数量=%d（须恰好为1，percentStacked不计入）" % len(qualifying)
        chart = main
        # 二维堆叠柱状图/柱形图：barDir=col + grouping=stacked（percentStacked 不认可）
        is_2d_stacked = chart["bar_dir"] == "col" and chart["grouping"] == "stacked"
        # 用于比较"综管科+服务科+合计的登记数量"与"实际核验数量"：
        # 图表系列须含登记侧（综管科/服务科，二者堆叠即为合计）与核验侧（核验数量）
        names = [s["name"] for s in chart["series"]]
        has_zg = any("综管" in n for n in names)
        has_fw = any("服务" in n for n in names)
        has_verify = any("核验" in n for n in names)
        compare_ok = has_zg and has_fw and has_verify
        # 覆盖5类资源（且类别正是细则指定的5个名称）
        cats_all = set()
        for s in chart["series"]:
            for cv in s["cats"]:
                for name in CATEGORIES:
                    if name in cv:
                        cats_all.add(name)
        cover = len(cats_all)
        cover_ok = cover == 5
        # 数据来自原表3-17行：登记数量取自 Sheet1 D 列(D3:D17)，核验数量取自
        # Sheet1 E 列(E3:E17)。校验综管科/服务科(登记)系列与核验数量系列的
        # c:val 公式引用（c:f）落在 Sheet1!D3:D17 / Sheet1!E3:E17 范围内。
        def series_ref_in_range(role_names, col_letter):
            col_idx = col_letter_to_idx(col_letter)
            matched = 0
            for s in chart["series"]:
                if not any(k in s["name"] for k in role_names):
                    continue
                matched += 1
                parsed = parse_sheet_ref(s.get("val_formula"))
                if parsed is None:
                    return False
                sheet_name, c1, r1, c2, r2 = parsed
                if sheet_name != "Sheet1":
                    return False
                if not (c1 == col_idx and c2 == col_idx):
                    return False
                if not (r1 >= 3 and r2 <= 17):
                    return False
            return matched > 0
        source_ok = (series_ref_in_range(("综管",), "D")
                     and series_ref_in_range(("服务",), "D")
                     and series_ref_in_range(("核验",), "E"))
        ok = is_2d_stacked and count_ok and compare_ok and cover_ok and source_ok
        return ok, ("二维堆叠柱图=%s(数量=%d), 比较登记(综管=%s/服务=%s)与核验(=%s), "
                    "覆盖资源类别=%d/5 %s, 系列公式引用Sheet1原表3-17行(D/E列)=%s") % (
            is_2d_stacked, len(qualifying), has_zg, has_fw, has_verify,
            cover, sorted(cats_all), source_ok)
    ok, det = chk_chart_object()
    hits.append((3, ok, "图表对象", det))


    # ---------- +5 服务科+综管科堆叠柱 与 核验柱相连(距离0-0.1cm) ----------
    # 细则：服务科和综管科的堆叠柱形，与核验数量的柱形相连在一起，
    #       两个柱形之间距离为 0-0.1 厘米。
    # "两个柱形之间的距离"指 登记堆叠柱 与 核验柱 之间的横向间距，到底由
    # gapWidth 还是 overlap 控制，取决于图表的数据结构：
    #
    #   结构A（同分类点、不同系列）：每类资源是 1 个分类点，登记(综管+服务堆叠)
    #     与核验是该分类点下的两个并列系列。此时两柱间距由 overlap 控制，
    #     overlap=100 表示完全重叠贴合（间距≈0）。
    #
    #   结构B（相邻分类点）：把每类资源拆成"X 登记""X 核验"两个相邻分类点，
    #     登记柱与核验柱分属相邻的两个分类点。此时两柱间距由 gapWidth 控制
    #     （类别间间隙，占柱宽的百分比），overlap 只影响同一分类点内系列的重叠、
    #     对登记↔核验间距无作用。要"相连/间距 0-0.1cm"需 gapWidth 接近 0。
    #
    # 通过分类标签是否成对出现"登记/核验"来区分结构 B；否则按结构 A 处理。
    def chk_adjacent():
        if main is None:
            return False, "无图表"
        ov = main["overlap"]
        gw = main["gap_width"]
        # 判断是否为"结构B"：分类标签里同时出现"登记"和"核验"后缀
        cat_labels = []
        for s in main["series"]:
            if s["cats"]:
                cat_labels = s["cats"]
                break
        has_reg_cat = any("登记" in c for c in cat_labels)
        has_ver_cat = any("核验" in c for c in cat_labels)
        structure_b = has_reg_cat and has_ver_cat

        # 分类点个数：取任一系列的类别数；取不到则退回 5（CATEGORIES 数量）。
        n_cats = len(cat_labels) if cat_labels else 5
        # 每个分类点内并列的系列数（结构A: 综管科+服务科堆叠 与 核验 两组并列；
        # 结构B: 每个分类点只有1组柱，登记/核验各占一个分类点）。
        series_per_cat = 1 if structure_b else 2

        # ---- 取图表实际渲染尺寸（EMU）：按 drawing rels 精确关联到本 chart part ----
        anchor = chart_anchor_extent(bundle, main.get("path"))
        chart_cx = anchor["cx"] if anchor else None
        anchor_src = anchor["source"] if anchor else "未找到锚点"

        if chart_cx is None or n_cats <= 0:
            # 没有可用的图表尺寸，无法做几何换算，退回不达标（无法验证≠达标）。
            return False, ("无法定位图表锚点尺寸，无法计算实际柱间距厘米"
                           "（结构%s，gapWidth=%s, overlap=%s）"
                           % ("B" if structure_b else "A", gw, ov))

        # ---- 绘图区宽度（EMU）：优先用 c:manualLayout 的精确比例，
        #      否则用"图表区宽度 - 估算的坐标轴/图例/标题边距"做近似。----
        manual_w = main.get("plot_manual_w")
        if manual_w is not None:
            plot_w_emu = chart_cx * manual_w
            plot_src = "manualLayout精确宽度比例=%.4f" % manual_w
        else:
            # 无手动布局：Excel 走自动布局，XML 未暴露具体算法，只能近似。
            # 经验估算：图例/数值轴占用图表宽度的一部分，扣除后取 0.86 作图表区
            # 转绘图区宽度比例的近似系数（存在图例时更保守地扣 0.78）。
            approx_ratio = 0.78 if main.get("legend_pos") else 0.86
            plot_w_emu = chart_cx * approx_ratio
            plot_src = "无manualLayout，按图表宽度近似系数%.2f估算" % approx_ratio

        # ---- 分类槏宽（每个分类点占用的绘图区宽度，EMU）----
        cat_slot_emu = plot_w_emu / n_cats

        # ---- 由 gapWidth/overlap 换算柱宽与柱间距 ----
        # OOXML: gapWidth=类别簇之间的间隙，占"柱宽"的百分比；
        #        overlap=同一分类点内相邻系列的重叠百分比（100=完全重叠贴合，
        #        0=并列不重叠，负值=系列间有间隙）。
        # 单个分类槏 = 簇宽(cluster_w) + 间隙(gap)；gap = cluster_w * gw/100
        #   => cat_slot = cluster_w * (1 + gw/100)  => cluster_w = cat_slot/(1+gw/100)
        gw_frac = (gw or 0) / 100.0
        cluster_w_emu = cat_slot_emu / (1 + gw_frac) if (1 + gw_frac) > 0 else cat_slot_emu

        if structure_b:
            # 结构B：登记柱与核验柱分属相邻两个分类点，每个分类点内只有1根柱
            # （簇宽即柱宽），两柱间距 = 分类间隙 = cluster_w * gw/100。
            bar_w_emu = cluster_w_emu
            gap_emu = cluster_w_emu * gw_frac
            structure_desc = "结构B(登记/核验为相邻分类点)"
        else:
            # 结构A：同一分类点内综管+服务堆叠柱 与 核验柱 并列，重叠由 overlap 控制。
            # overlap 为相邻系列宽度的重叠百分比：两柱中心距 = bar_w*(1-overlap/100)，
            # 单柱宽 bar_w = cluster_w / (series - (series-1)*overlap/100)（series=2时）。
            ov_frac = (ov or 0) / 100.0
            denom = series_per_cat - (series_per_cat - 1) * ov_frac
            bar_w_emu = cluster_w_emu / denom if denom > 0 else cluster_w_emu
            # 两柱边缘间距 = 中心距 - 柱宽 = bar_w*(1-ov_frac) - bar_w = -bar_w*ov_frac
            # overlap<100 时两柱有缝隙；>=100 时完全贴合/重叠，间距记为0（不为负）。
            gap_emu = max(0.0, bar_w_emu * (1 - ov_frac))
            structure_desc = "结构A(登记/核验为同分类点系列)"

        gap_cm = gap_emu / EMU_PER_CM
        ok = 0 <= gap_cm <= 0.1
        return ok, ("%s：图表尺寸cx=%s(%s)，绘图区宽度≈%.0fEMU(%s)，"
                    "分类槏宽≈%.0fEMU(共%d类)，柱宽≈%.0fEMU，"
                    "gapWidth=%s/overlap=%s，实际柱间距≈%.4f厘米（要求0-0.1厘米）"
                    % (structure_desc, chart_cx, anchor_src, plot_w_emu, plot_src,
                       cat_slot_emu, n_cats, bar_w_emu, gw, ov, gap_cm))
    ok, det = chk_adjacent()
    hits.append((5, ok, "柱形相邻", det))


    # ---------- +3 图表分类轴/数值轴 ----------
    # 细则：
    #  · 横轴依次显示"便携学习终端""智能显示设备""网络接入设备"
    #    "多功能文印设备""移动采集设备"5个资源类别；
    #  · 文本水平显示于柱形组合下方，不可出现文字倾斜显示；
    #  · 纵轴分类依次显示 0、5、10、15、20、25、30、35、40 数字
    #    （即范围 0-40、主刻度间隔 5）；
    #  · 数字左侧有纵轴坐标单位"数量（台）"。
    def chk_axes():
        if main is None:
            return False, "无图表"
        # —— 横轴：依次覆盖细则指定的 5 个资源类别 ——
        axis_labels = list(main["cat_axis_labels"])
        # 若 catAx 自身无 cat（图表用系列 cat），退而取系列 cat 作为横轴文本
        if not any(any(name in lab for name in CATEGORIES) for lab in axis_labels):
            for s in main["series"]:
                if s["cats"]:
                    axis_labels = list(s["cats"])
                    break
        # "依次"显示：按出现顺序提取类别，需与 CATEGORIES 顺序一致且齐 5 类
        seq = []
        for lab in axis_labels:
            for name in CATEGORIES:
                if name in lab and name not in seq:
                    seq.append(name)
        cats_ok = seq == CATEGORIES

        # —— 横轴文本水平、不倾斜：rot 为 None 或 0 视为水平 ——
        rot = main["cat_axis_rot"]
        horiz_ok = rot in (None, 0)

        # —— 横轴标签须位于"柱形组合下方"：轴位置为底部（b），
        #    且标签位置为 nextTo（紧贴轴线/柱形下方）；low/high 会把标签
        #    挪到绘图区上下边缘，不满足"柱形组合下方"。两者缺省值均为
        #    Excel 默认（b / nextTo），XML 中省略该节点时按默认值处理。
        ax_pos = main["cat_axis_pos"]
        lbl_pos = main["cat_axis_lbl_pos"]
        pos_ok = ax_pos in (None, "b") and lbl_pos in (None, "nextTo")

        # —— 纵轴刻度 0,5,10,...,40：范围 0-40、主刻度间隔 5 ——
        # c:max 缺失 = Excel 用自动最大值：不能直接判 None 为不合格，也不能
        # 直接判 None 为合格——需结合系列真实数据的最大值来推断"自动算出的
        # 上限是否等于 40"。Excel 的自动轴算法会把上限取整到"比数据最大值大
        # 且是主刻度间隔整数倍"的最小值；已知细则要求主刻度间隔为 5，故当
        # c:max 缺失时，用 数据最大值 与 majorUnit（若也缺失则按细则要求的 5
        # 兜底）推算：ceil(数据最大值 / unit) * unit 是否等于 40。
        vmin = main["val_axis_min"]
        vmax = main["val_axis_max"]
        vunit = main["val_axis_unit"]
        data_vals = [v for v in main["val_axis_num_cache"] if v is not None]
        unit_for_calc = vunit if vunit else 5
        if vmax is not None:
            # 写死了 c:max：必须恰好是 40
            vmax_ok = vmax == 40
            vmax_src = "显式c:max=%s" % vmax
        elif data_vals:
            # 未写 c:max（自动最大值）：按数据最大值向上取整到 unit 的整数倍推算
            data_max = max(data_vals)
            auto_vmax = math.ceil(data_max / unit_for_calc) * unit_for_calc
            vmax_ok = auto_vmax == 40
            vmax_src = "自动最大值(数据最大值=%s,按步长%s取整推算)=%s" % (
                data_max, unit_for_calc, auto_vmax)
        else:
            # 既无 c:max 也无系列数据缓存可用：无法验证，判不合格（不可"无法验证即达标"）
            vmax_ok = False
            vmax_src = "无c:max且无数据缓存，无法推算自动最大值"
        vmin_ok = vmin in (0, 0.0, None)
        vunit_ok = vunit == 5
        scale_ok = vmin_ok and vmax_ok and vunit_ok

        # —— 纵轴坐标单位"数量（台）"（数字左侧的轴标题）——
        vtitle = main["val_axis_title"]
        unit_ok = ("数量" in vtitle) and ("台" in vtitle)

        ok = cats_ok and horiz_ok and pos_ok and scale_ok and unit_ok
        return ok, ("横轴依次5类=%s(%s), 文本水平=%s(rot=%s), "
                    "轴位置=%s/标签位置=%s(柱形组合下方=%s), "
                    "纵轴范围=%s~%s(%s) 间隔=%s(0-40步长5=%s), "
                    "纵轴单位'数量(台)'=%s") % (
            cats_ok, seq, horiz_ok, rot, ax_pos, lbl_pos, pos_ok,
            vmin, vmax, vmax_src, vunit, scale_ok, unit_ok)
    ok, det = chk_axes()
    hits.append((3, ok, "分类轴/数值轴", det))

    # ---------- +1 图表标题 ----------
    # 细则：图表标题包含"固定资产登记数量（综管科+服务科）与核验数量对比"等关键信息，
    #       能够明确表达登记数量与核验数量的对比关系。
    # 判定：标题须体现"登记数量"与"核验数量"两侧，并表达出二者的"对比"关系。
    def chk_title():
        if main is None:
            return False, "无图表"
        t = main["title"]
        has_reg = "登记" in t            # 登记数量
        has_verify = "核验" in t         # 核验数量
        has_compare = "对比" in t        # 明确的对比关系
        ok = has_reg and has_verify and has_compare
        return ok, "标题=「%s」(含登记=%s/核验=%s/对比=%s)" % (
            t, has_reg, has_verify, has_compare)
    ok, det = chk_title()
    hits.append((1, ok, "图表标题", det))

    # ---------- +1 图表图例 ----------
    # 细则：图例包含"综管科""服务科""核验数量"3个系列名称；
    #       图例内容及颜色与图表系列一一对应。
    # 问题：仅检查 legend_pos 非空和系列名包含综管/服务/核验，未解析
    #       c:legendEntry/c:delete（图例条目可被单独隐藏，即使系列本身存在，
    #       图例里也可能看不到该条目），也未核对图例实际显示文本与颜色是否
    #       与对应系列一一对应。
    # 措施：解析 c:legendEntry 的删除标记，排除被隐藏的图例条目后，
    #       确认"综管科""服务科""核验数量"三项均实际显示在图例中，
    #       且各自图例条目（按系列索引对应）的颜色与该系列本身颜色一致
    #       （OOXML 图例条目默认继承对应系列颜色，一一对应由结构保证，
    #       这里显式核对 color_hex 避免系列被单独改色导致图例与图形不一致）。
    def chk_legend():
        if main is None:
            return False, "无图表"
        legend_on = main["legend_pos"] is not None
        if not legend_on:
            return False, "图例不存在"
        deleted_idx = main.get("legend_deleted_idx") or set()
        # 实际显示的图例条目：系列列表中排除被 c:legendEntry/c:delete 隐藏的项
        visible = [(i, s) for i, s in enumerate(main["series"]) if i not in deleted_idx]
        need = [("综管", "综管科"), ("服务", "服务科"), ("核验", "核验数量")]
        found = {}
        for key, _ in need:
            for i, s in visible:
                if key in s["name"]:
                    found[key] = (i, s)
                    break
        present = all(key in found for key, _ in need)
        # 颜色一一对应：图例条目颜色即所关联系列自身的 color_hex（结构上保证），
        # 这里核对该系列确有可辨识的填充色（否则图例色块与图形对不上/无法比对）。
        color_ok = present and all(found[key][1]["color_hex"] for key, _ in need)
        ok = present and color_ok
        return ok, ("图例存在=%s, 隐藏条目idx=%s, 实际显示系列=%s, "
                    "综管科/服务科/核验数量三项均显示=%s, "
                    "各自图例颜色与系列颜色一致(存在有效填充色)=%s") % (
            legend_on, sorted(deleted_idx),
            [s["name"] for _, s in visible], present, color_ok)
    ok, det = chk_legend()
    hits.append((1, ok, "图表图例", det))

    # ---------- +3 图表颜色 ----------
    # 细则：综管科系列使用蓝色或近似蓝色，服务科系列使用绿色或近似绿色，
    #       核验数量系列使用橙色或近似橙色；三类系列颜色彼此易于区分。
    def chk_colors():
        if main is None:
            return False, "无图表"
        role_color = {}
        role_hex = {}
        for s in main["series"]:
            n = s["name"]
            if "综管" in n:
                role_color["zg"] = s["color_class"]
                role_hex["zg"] = s["color_hex"]
            elif "服务" in n:
                role_color["fw"] = s["color_class"]
                role_hex["fw"] = s["color_hex"]
            elif "核验" in n:
                role_color["verify"] = s["color_class"]
                role_hex["verify"] = s["color_hex"]
        zg = role_color.get("zg")
        fw = role_color.get("fw")
        verify = role_color.get("verify")
        # 综管=蓝(近似)、服务=绿(近似)、核验=橙(近似)
        color_ok = (zg == "blue" and fw == "green" and verify == "orange")
        # 三类系列颜色彼此易于区分：三者已识别且互不相同
        got = [c for c in (zg, fw, verify) if c is not None]
        distinct_ok = len(got) == 3 and len(set(got)) == 3
        ok = color_ok and distinct_ok
        return ok, ("综管=%s(蓝,hex=%s), 服务=%s(绿,hex=%s), 核验=%s(橙,hex=%s), "
                    "三色互异=%s") % (
            zg, role_hex.get("zg"), fw, role_hex.get("fw"),
            verify, role_hex.get("verify"), distinct_ok)
    ok, det = chk_colors()
    hits.append((3, ok, "图表颜色", det))

    # ---------- +5 图表堆叠内部标签 ----------
    # 细则：综管科和服务科两个堆叠系列的柱段内部显示各自数值标签；
    #       5类资源中能识别出"综管科"数据 6/22/18/12/5 和
    #       "服务科"数据 9/8/6/3/2。
    # 问题：原代码只看 showVal 是否开启和系列值是否匹配，不检查标签位置是否为
    #       "内部"（ctr/inBase/inEnd），也不验证实际标签文本是否等于这些数值。
    # 措施：解析 dLblPos（系列级/单点级）是否为内部位置，并结合
    #       series_dlbl_info 推断出的每个数据点实际显示文本，逐类核对文本值。
    def chk_inner_labels():
        if main is None:
            return False, "无图表"

        def role_series(role_key):
            for s in main["series"]:
                if role_key in s["name"]:
                    return s
            return None

        zg_ser = role_series("综管")
        fw_ser = role_series("服务")

        def point_pos(s, idx):
            """取该系列第idx个数据点实际生效的标签位置：单点覆盖优先，否则用系列级 dlbl_pos。"""
            if s is None:
                return None
            return s["dlbl_point_pos"].get(idx, s["dlbl_pos"])

        def inner_ok_for(s):
            """该系列是否开启显示且每个仍显示的数据点标签位置都在柱段内部。"""
            if s is None or not s["has_dlbl"]:
                return False
            text_map = s["dlbl_text"] or {}
            if not text_map:
                return False
            return all(point_pos(s, idx) in INNER_DLBL_POS for idx in text_map)

        zg_inner_ok = inner_ok_for(zg_ser)
        fw_inner_ok = inner_ok_for(fw_ser)
        dlbl_on = zg_inner_ok and fw_inner_ok

        def label_text_match(s, cat_names, expected):
            """按 CATEGORIES 顺序，核对该系列每个分类点实际显示的标签文本
            （dlbl_text）是否等于期望数值（允许 6 / 6.0 / "6" 等常见数值文本形式）。"""
            if s is None:
                return 0
            matched = 0
            for j, cv in enumerate(cat_names):
                ci = None
                for i, name in enumerate(CATEGORIES):
                    if name in cv:
                        ci = i
                        break
                if ci is None:
                    continue
                txt = s["dlbl_text"].get(j)
                if txt is None:
                    continue
                try:
                    if abs(float(txt) - expected[ci]) <= 0.5:
                        matched += 1
                except ValueError:
                    continue
            return matched

        zg_text_match = label_text_match(zg_ser, zg_ser["cats"] if zg_ser else [], EXP_ZG)
        fw_text_match = label_text_match(fw_ser, fw_ser["cats"] if fw_ser else [], EXP_FW)

        # 5类资源能识别出综管科 6/22/18/12/5、服务科 9/8/6/3/2
        # ——同时要求：系列数值本身对应（mapped，兼容无自定义文本、走缓存数值的常见情况）
        #             以及 标签实际显示文本 也核对一致（避免自定义文本被改写成与数值不符的内容）。
        zg_match = count_match(mapped["zg"], EXP_ZG)
        fw_match = count_match(mapped["fw"], EXP_FW)
        ok = (dlbl_on and zg_match == 5 and fw_match == 5
              and zg_text_match == 5 and fw_text_match == 5)
        return ok, ("综管标签内部位置=%s/服务标签内部位置=%s, "
                    "综管数值识别 %d/5(%s) 标签文本核对 %d/5, "
                    "服务数值识别 %d/5(%s) 标签文本核对 %d/5") % (
            zg_inner_ok, fw_inner_ok, zg_match, mapped["zg"], zg_text_match,
            fw_match, mapped["fw"], fw_text_match)
    ok, det = chk_inner_labels()
    hits.append((5, ok, "堆叠内部标签", det))

    # ---------- +5 图表柱顶合计标签 ----------
    # 细则：每个堆叠柱顶部显示合计登记数量标签 15、30、24、15、7；
    #       核验数量柱顶部显示 13、27、22、14、6。
    # 问题：原代码用综管+服务系列值计算合计，并把任一堆叠系列 showVal 当作
    #       堆叠柱顶标签；这不能证明图表顶部实际显示合计15/30/24/15/7，普通
    #       堆叠内部标签（+5 那一条细则的场景）也会被误判为柱顶合计标签。
    # 措施：检查是否存在合计系列/自定义标签，或数据标签位置确实在柱顶
    #       （堆叠图中柱顶标签只能来自堆叠顶层系列且位置为 outEnd，或来自
    #       独立的"合计"系列/图表级顶部标签），并核对实际显示的标签文本
    #       是否等于期望的合计/核验柱顶数值。
    def chk_top_labels():
        if main is None:
            return False, "无图表"
        series_list = main["series"]

        def role_series(role_key):
            for s in series_list:
                if role_key in s["name"]:
                    return s
            return None

        # 堆叠顶层系列：图表系列列表中最后一个"综管"或"服务"系列（堆叠图里
        # 排在后面的系列画在上层，其 outEnd 标签位置才对应堆叠柱的顶部）。
        stack_roles = [s for s in series_list
                       if "综管" in s["name"] or "服务" in s["name"]]
        top_stack_ser = stack_roles[-1] if stack_roles else None
        total_ser = role_series("合计")
        verify_ser = role_series("核验")

        def point_pos(s, idx):
            if s is None:
                return None
            return s["dlbl_point_pos"].get(idx, s["dlbl_pos"])

        def top_ok_for(s, require_outer):
            """该系列是否开启显示标签，且（若要求外部位置）每个仍显示的
            数据点标签位置都在柱顶外部（outEnd）。"""
            if s is None or not s["has_dlbl"]:
                return False
            text_map = s["dlbl_text"] or {}
            if not text_map:
                return False
            if not require_outer:
                return True
            return all(point_pos(s, idx) in OUTER_DLBL_POS for idx in text_map)

        def label_text_match(s, expected):
            """核对该系列每个分类点实际显示的标签文本是否等于期望数值。"""
            if s is None:
                return 0
            matched = 0
            for j, cv in enumerate(s["cats"]):
                ci = None
                for i, name in enumerate(CATEGORIES):
                    if name in cv:
                        ci = i
                        break
                if ci is None:
                    continue
                txt = s["dlbl_text"].get(j)
                if txt is None:
                    continue
                try:
                    if abs(float(txt) - expected[ci]) <= 0.5:
                        matched += 1
                except ValueError:
                    continue
            return matched

        # 合计标签的来源：独立"合计"系列（自定义文本/showVal 均可，不要求
        # outEnd 位置，因为该系列本身即代表合计）；否则要求堆叠顶层系列
        # 在 outEnd 位置显示（堆叠柱顶=该点标签，实际数值等于累计高度）。
        if total_ser is not None:
            total_dlbl_ok = top_ok_for(total_ser, require_outer=False)
            total_text_match = label_text_match(total_ser, EXP_TOTAL)
        else:
            total_dlbl_ok = top_ok_for(top_stack_ser, require_outer=True)
            total_text_match = label_text_match(top_stack_ser, EXP_TOTAL)
        # 核验柱顶部显示标签：核验系列开启显示，且标签位置在柱外顶部。
        verify_dlbl_ok = top_ok_for(verify_ser, require_outer=True)
        verify_text_match = label_text_match(verify_ser, EXP_VERIFY)

        ok = (total_dlbl_ok and verify_dlbl_ok
              and total_text_match == 5 and verify_text_match == 5)
        return ok, ("合计柱顶标签=%s(有合计系列=%s) 文本核对 %d/5, "
                    "核验柱顶标签=%s 文本核对 %d/5") % (
            total_dlbl_ok, total_ser is not None, total_text_match,
            verify_dlbl_ok, verify_text_match)
    ok, det = chk_top_labels()
    hits.append((5, ok, "柱顶合计标签", det))

    total = sum(score for score, hit, _, _ in hits if hit)
    return total, hits


# 评分细则文本（按 label 对应到完整细则内容，用于结果打印）
RUBRIC_TEXT = {
    "图表对象": "图表对象：新增1个可编辑图表，图表类型为二维堆叠柱状图或二维堆叠柱形图，用于比较“综管科+服务科+合计的登记数量”与“实际核验数量”。图表覆盖5类资源，类别分别为“便携学习终端”“智能显示设备”“网络接入设备”“多功能文印设备”“移动采集设备”，对应数据来自原表3-17行中的登记数量与核验数量。",
    "柱形相邻": "服务科和综管科的堆叠柱形与核验数量的柱形相连在一起，两个柱形之间距离为0-0.1厘米。",
    "分类轴/数值轴": "图表分类轴：横轴依次显示“便携学习终端”“智能显示设备”“网络接入设备”“多功能文印设备”“移动采集设备”5个资源类别，文本水平显示于柱形组合下方，不可出现文字倾斜显示；纵轴分类依次显示0、5、10、15、20、25、30、35、40数字，数字左侧有纵轴坐标单位“数量（台）”。",
    "图表标题": "图表标题：图表标题包含“固定资产登记数量（综管科+服务科）与核验数量对比”等关键信息，能够明确表达登记数量与核验数量对比关系。",
    "图表图例": "图表图例：图例包含“综管科”“服务科”“核验数量”3个系列名称，图例内容及颜色与图表系列一一对应。",
    "图表颜色": "图表颜色：综管科系列使用蓝色或近似蓝色，服务科系列使用绿色或近似绿色，核验数量系列使用橙色或近似橙色，三类系列颜色彼此易于区分。",
    "堆叠内部标签": "图表堆叠内部标签：综管科和服务科两个堆叠系列的柱段内部显示各自数值标签，5类资源中能识别出”综管科“数据6/22/18/12/5和”服务科“数据9/8/6/3/2。",
    "柱顶合计标签": "图表柱顶合计标签：每个堆叠柱顶部显示合计登记数量标签15、30、24、15、7；核验数量柱顶部显示13、27、22、14、6。",
}


# ============================================================
#  对外接口：evaluate(dir_path)
# ============================================================
SCRIPT_ID = "100"


def _find_target_file(dir_path):
    """在指定目录内定位被评估的 xlsx/xlsm 文档。
    优先匹配默认名"工作簿1_堆叠对比图_可编辑.xlsx"，否则取目录下第一个 .xlsx/.xlsm。
    """
    if not os.path.isdir(dir_path):
        return None
    preferred = os.path.join(dir_path, "工作簿1_堆叠对比图_可编辑.xlsx")
    if os.path.isfile(preferred):
        return preferred
    candidates = []
    for name in os.listdir(dir_path):
        low = name.lower()
        if low.endswith((".xlsx", ".xlsm")) and not name.startswith("~$"):
            candidates.append(os.path.join(dir_path, name))
    candidates.sort()
    return candidates[0] if candidates else None


# ============================================================
#  维度1：可用与可修改性（门槛）
# ============================================================
def eval_dimension1(path, bundle):
    """细则：交付文件为xlsx或.xlsm格式，文件可正常打开。"""
    d = []
    ext_ok = path.lower().endswith((".xlsx", ".xlsm"))
    c1 = ext_ok and bundle.ok
    d.append((c1, "交付文件为xlsx或.xlsm格式，文件可正常打开" if c1
              else "格式不符或无法打开（扩展名ok=%s, 打开ok=%s）" % (ext_ok, bundle.ok)))
    return all(ok for ok, _ in d), d


def _build_error_result(file_name, error_msg, rubric_items):
    """脚本自身错误 / 文件不存在等：status=error。
    契约：维度一未通过（或未评估）时 dim2_items 恒为 []，不返回任何维度二
    条目（即使全部标记为未命中）。"""
    max_score = sum(md for _, md in rubric_items if md > 0)
    return {
        "id": SCRIPT_ID,
        "file_name": file_name or "",
        "status": "error",
        "error": error_msg,
        "dim1_pass": False,
        "dim1_reason": error_msg,
        "dim2_items": [],
        "total_score": 0,
        "max_score": max_score,
    }


# 维度二评分项定义（顺序与 eval_dimension2 内一致，含扣分项）
_RUBRIC_ITEMS = [
    ("图表对象", 3),
    ("柱形相邻", 5),
    ("分类轴/数值轴", 3),
    ("图表标题", 1),
    ("图表图例", 1),
    ("图表颜色", 3),
    ("堆叠内部标签", 5),
    ("柱顶合计标签", 5),
]


def evaluate(dir_path: str) -> dict:
    """评估入口：接收脚本所在目录路径，返回结构化评分字典。

    总分 = dim2_items 所有条目 delta 之和。
    """
    file_path = _find_target_file(dir_path)
    if file_path is None:
        return _build_error_result(
            "", "目录不存在或未找到 .xlsx/.xlsm 文件：%s" % dir_path, _RUBRIC_ITEMS)

    file_name = os.path.basename(file_path)

    try:
        bundle = ZipBundle(file_path)
        if not bundle.ok:
            return {
                "id": SCRIPT_ID,
                "file_name": file_name,
                "status": "ok",
                "error": None,
                "dim1_pass": False,
                "dim1_reason": "文件无法作为 xlsx 打开：%s" % (bundle.error or ""),
                "dim2_items": [],
                "total_score": 0,
                "max_score": sum(md for _, md in _RUBRIC_ITEMS if md > 0),
            }

        shared = load_shared_strings(bundle)
        sheet_targets = get_sheet_targets(bundle)
        sheet1_path = sheet_targets.get("Sheet1")
        if sheet1_path:
            sheet1_cells, sheet1_merges = read_sheet_cells(
                bundle, sheet1_path, shared)
        else:
            sheet1_cells = {}
            sheet1_merges = []
        sheet1_raw_text = "".join(shared)
        sheet1_raw_text += "".join(sheet1_cells.values())

        theme_colors = load_theme_colors(bundle)
        chart_paths = find_chart_paths(bundle)
        charts = []
        for cp in chart_paths:
            root = bundle.xml(cp)
            if root is not None:
                ch = parse_chart(root, theme_colors)
                ch["path"] = cp
                charts.append(ch)

        drawing_text = ""
        for n in bundle.names:
            if re.search(r"drawings/drawing\d+\.xml$", n):
                drawing_text += bundle.read_text(n) or ""

        # 维度一
        passed, d1 = eval_dimension1(file_path, bundle)
        dim1_reason = "" if passed else "; ".join(
            msg for ok, msg in d1 if not ok)

        if not passed:
            dim2_items = [
                {
                    "rule": RUBRIC_TEXT.get(label, label),
                    "max_delta": md,
                    "delta": 0,
                    "hit": False,
                    "detail": "",
                }
                for label, md in _RUBRIC_ITEMS
            ]
            return {
                "id": SCRIPT_ID,
                "file_name": file_name,
                "status": "ok",
                "error": None,
                "dim1_pass": False,
                "dim1_reason": dim1_reason,
                "dim2_items": dim2_items,
                "total_score": 0,
                "max_score": sum(md for _, md in _RUBRIC_ITEMS if md > 0),
            }

        # 维度二
        _total, hits = eval_dimension2(
            bundle, sheet_targets, sheet1_cells, sheet1_raw_text,
            charts, drawing_text, sheet1_merges, shared)

        dim2_items = []
        for score, hit, label, det in hits:
            dim2_items.append({
                "rule": RUBRIC_TEXT.get(label, label),
                "max_delta": score,
                "delta": score if hit else 0,
                "hit": bool(hit),
                "detail": "",
            })

        # 脚本契约：total_score = sum(item.delta)，扣分项命中时 delta 为负数
        # 并计入总分。
        total_score = sum(it["delta"] for it in dim2_items)
        max_score = sum(it["max_delta"] for it in dim2_items if it["max_delta"] > 0)

        return {
            "id": SCRIPT_ID,
            "file_name": file_name,
            "status": "ok",
            "error": None,
            "dim1_pass": True,
            "dim1_reason": "",
            "dim2_items": dim2_items,
            "total_score": total_score,
            "max_score": max_score,
        }

    except Exception as e:
        return _build_error_result(file_name, "脚本执行异常：%s" % e, _RUBRIC_ITEMS)


# ============================================================
#  本地调试入口（不作为主输出通道）
# ============================================================
if __name__ == "__main__":
    if len(sys.argv) >= 2:
        _dir = sys.argv[1]
    else:
        _dir = os.path.dirname(os.path.abspath(__file__))
    print(json.dumps(evaluate(_dir), ensure_ascii=False, indent=2))
