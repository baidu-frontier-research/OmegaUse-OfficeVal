# -*- coding: utf-8 -*-
"""
自动评估脚本：七年级地理探究成绩登记_公开版(2)_嵌套饼图.xlsx

评估逻辑
========
维度1（可用与可修改性）：硬性门槛。任一条不满足 -> 直接 0 分，不再检查维度2。
维度2（完成度评分细则）：在维度1全部通过后才评估。
    - 得分点：必须满足该细则的【每一个】子条件才加分（加分为正）。
    - 扣分点：只要命中该细则的【任意一个】子条件即扣分（分数为负）。
    - 最终得分 = 各命中细则分数之和（可正可负）。

所有判定均自动完成，不依赖人工。对难以严格量化的意图（如“标签不大面积重叠”
“颜色至少四种”），采用结构化、可度量的近似规则灵活实现评估意图。
"""

import os
import re
import sys
import zipfile
from xml.etree import ElementTree as ET

# 脚本编号（与所在目录 officeval_094 一致）
SCRIPT_ID = "094"

# 默认待评估文档文件名；实际以 evaluate(dir_path) 传入目录中的实际文件为准，
# 若目录中无同名文件，则扫描该目录下的 .xlsx/.xlsm 作为兜底。
TARGET_FILE = "七年级地理探究成绩登记_公开版(2)_嵌套饼图.xlsx"

# ---- 命名空间 ----
NS = {
    "c": "http://schemas.openxmlformats.org/drawingml/2006/chart",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "ct": "http://schemas.openxmlformats.org/package/2006/content-types",
    "ss": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "xdr": "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
    "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
}

# 业务常量
SHEET_NAMES = ["七1班", "七13班"]
COL_CATEGORIES = ["地图基础", "经纬定位", "地形判读", "气候要素", "区域差异", "聚落形态"]  # B1:Q1 能力大类
ROW_LEVELS = ["卓越表现", "稳定达成", "基本达成", "持续提升"]                              # A4:A7 能力层级
EXPECTED_TITLES = {
    "七1班": "七1班 探究成绩多类别嵌套饼图",
    "七13班": "七13班 探究成绩多类别嵌套饼图",
}
ERROR_TOKENS = ["#REF!", "#VALUE!", "#N/A", "#DIV/0!", "#NUM!", "#NULL!", "#NAME?"]

# 各评分细则的“分数 + 细则内容”原文（用于命中结果展示，格式为 “+1：...”）。
# key 与 check_dimension2 中 results 的判定顺序一一对应。
RULE_TEXTS = {
    "p1_charts": "+1：工作簿中分别基于“七1班”和“七13班”两份成绩数据各生成1个多类别嵌套合并饼图，"
                 "共2个图表，且两个图表标题能明确区分“七1班”和“七13班”。两个图表的数据源均来自"
                 "对应工作表A1:Q7中的成绩数据，覆盖“地图基础”到“聚落形态”各非空能力层级数据，不遗漏主要成绩列。",
    "p5_class": "+5：两个图表分类统计正确，如使用“卓越表现”“稳定达成”“基本达成”“持续提升”四类成绩数据"
                "作为饼图统计对象，“地图基础”“经纬定位”“地形判读”“气候要素”“区域差异”“聚落形态”则作为"
                "不同类别圆环层级；如使用“地图基础”“经纬定位”“地形判读”“气候要素”“区域差异”“聚落形态”"
                "六类成绩数据作为饼图统计对象，“卓越表现”“稳定达成”“基本达成”“持续提升”四类成绩数据则作为"
                "不同类别圆环层级。",
    "p1_label": "+1：每个图表的扇形上有数据标签，至少显示数值或百分比之一，标签不大面积重叠且不遮挡主要扇区。",
    "p1_legend": "+1：两个图表均包含图例，图例文字包含或能对应B1:Q1的“地图基础”“经纬定位”“地形判读”"
                 "“气候要素”“区域差异”“聚落形态”等6个成绩类别或包含能对应A4:A7的“卓越表现”“稳定达成”"
                 "“基本达成”“持续提升”等四个能力层级。",
    "p1_color": "+1：两个图表中同一成绩类别使用一致或接近一致的颜色方案，便于比较两个班级的数据；"
                "图中颜色至少四种颜色。",
    "p1_title": "+1：图表上方出现图表标题，标题内容分别为“七1班 探究成绩多类别嵌套饼图”及"
                "“七13班 探究成绩多类别嵌套饼图”。",
}


# =========================================================================
# 工具函数
# =========================================================================
def load_zip(path):
    z = zipfile.ZipFile(path)
    names = set(z.namelist())
    return z, names


def read_xml(z, name):
    try:
        return ET.fromstring(z.read(name))
    except Exception:
        return None


def read_text(z, name):
    try:
        return z.read(name).decode("utf-8", errors="replace")
    except Exception:
        return ""


def normalize_part_path(tgt):
    """将 workbook.xml.rels 中的 Target 归一化为 zip 内部件路径。
    Target 可能是 '/xl/worksheets/sheet1.xml'（绝对，相对包根）
    或 'worksheets/sheet1.xml'（相对 xl/ 目录）。"""
    if not tgt:
        return tgt
    if tgt.startswith("/"):
        return tgt.lstrip("/")          # 相对包根 -> 去掉前导斜杠
    if tgt.startswith("xl/"):
        return tgt
    return "xl/" + tgt                   # 相对 workbook 所在的 xl/ 目录


# =========================================================================
# 维度1：可用与可修改性
# =========================================================================
def check_dimension1(path, ctx):
    """返回 (passed: bool, details: list[(ok, msg)])。details 全部记录用于打印。"""
    details = []

    # --- 1.1 交付文件为 xlsx 或 .xlsm 格式，且可正常打开 ---
    ext = os.path.splitext(path)[1].lower()
    fmt_ok = ext in (".xlsx", ".xlsm")
    openable = False
    if fmt_ok:
        try:
            z, names = load_zip(path)
            # 关键部件存在即视为可正常打开
            openable = "xl/workbook.xml" in names and "[Content_Types].xml" in names
            ctx["zip"] = z
            ctx["names"] = names
        except Exception:
            openable = False
    ok = fmt_ok and openable
    details.append((ok, "1.1 交付文件为 xlsx 或 .xlsm 格式，文件可正常打开（扩展名=%s）" % ext))
    if not ok:
        return False, details

    z = ctx["zip"]
    names = ctx["names"]

    # 解析 workbook -> 表名到 sheetN.xml 的映射
    wb = read_xml(z, "xl/workbook.xml")
    rels = read_xml(z, "xl/_rels/workbook.xml.rels")
    rid_to_target = {}
    if rels is not None:
        for rel in rels.findall("rel:Relationship", NS):
            rid_to_target[rel.get("Id")] = rel.get("Target")
    sheet_xml = {}   # 表名 -> "xl/worksheets/sheetN.xml"
    if wb is not None:
        for sh in wb.findall("ss:sheets/ss:sheet", NS):
            nm = sh.get("name")
            rid = sh.get("{%s}id" % NS["r"])
            tgt = rid_to_target.get(rid, "")
            tgt = normalize_part_path(tgt)
            sheet_xml[nm] = tgt
    ctx["sheet_xml"] = sheet_xml

    # 解析各工作表数据网格，供维度2评分使用（不再作为维度1门槛判定）
    for s in SHEET_NAMES:
        if s not in sheet_xml:
            continue
        grid = parse_sheet_grid(z, sheet_xml[s], ctx)
        ctx.setdefault("grids", {})[s] = grid

    # 记录原生图表文件列表，供维度2评分使用（不再作为维度1门槛判定）
    chart_files = [n for n in names if re.match(r"xl/charts/chart\d+\.xml$", n)]
    ctx["chart_files"] = chart_files

    return True, details


# =========================================================================
# 工作表解析辅助
# =========================================================================
def col_letter_to_idx(letters):
    idx = 0
    for ch in letters:
        idx = idx * 26 + (ord(ch.upper()) - ord("A") + 1)
    return idx


def parse_ref(ref):
    m = re.match(r"([A-Za-z]+)(\d+)", ref)
    if not m:
        return None
    return int(m.group(2)), col_letter_to_idx(m.group(1))


def parse_sheet_grid(z, sheet_path, ctx):
    """返回 {(row,col): value}。读取共享字符串。"""
    if "shared_strings" not in ctx:
        ss = read_xml(z, "xl/sharedStrings.xml")
        strings = []
        if ss is not None:
            for si in ss.findall("ss:si", NS):
                txt = "".join(t.text or "" for t in si.iter("{%s}t" % NS["ss"]))
                strings.append(txt)
        ctx["shared_strings"] = strings
    strings = ctx["shared_strings"]

    root = read_xml(z, sheet_path)
    grid = {}
    if root is None:
        return grid
    for c in root.iter("{%s}c" % NS["ss"]):
        ref = c.get("r")
        rc = parse_ref(ref) if ref else None
        if not rc:
            continue
        t = c.get("t")
        v_el = c.find("ss:v", NS)
        is_el = c.find("ss:is", NS)
        val = None
        if t == "s" and v_el is not None:
            try:
                val = strings[int(v_el.text)]
            except Exception:
                val = v_el.text
        elif t == "str" and v_el is not None:
            val = v_el.text
        elif t == "inlineStr" and is_el is not None:
            val = "".join(tt.text or "" for tt in is_el.iter("{%s}t" % NS["ss"]))
        elif v_el is not None:
            val = v_el.text
        grid[rc] = val
    return grid


# =========================================================================
# 图表解析辅助
# =========================================================================
def chart_get_title(chart_root):
    title_el = chart_root.find(".//c:chart/c:title", NS)
    if title_el is None:
        return ""
    return "".join(t.text or "" for t in title_el.iter("{%s}t" % NS["a"]))


def chart_title_present_and_top(chart_root):
    """判断“图表上方出现图表标题”：
      - 存在 c:title 标题元素且 autoTitleDeleted != 1（标题确实显示）；
      - 标题位置在图表上方：未显式设置 layout 手动下移即视为默认顶部；
        若存在 manualLayout 指定 y 偏移且明显大于0(>0.5)，视为被移到下方/中部 -> 不在上方。
    返回 (present_top: bool, reason)。
    """
    title_el = chart_root.find(".//c:chart/c:title", NS)
    if title_el is None:
        return False, "无标题元素"
    deleted = chart_root.find(".//c:chart/c:autoTitleDeleted", NS)
    if deleted is not None and deleted.get("val") == "1":
        return False, "标题被删除(autoTitleDeleted=1)"
    # 标题位置：检查 title 的 manualLayout y 偏移
    ml = title_el.find(".//c:manualLayout", NS)
    if ml is not None:
        y = ml.find("c:y", NS)
        if y is not None:
            try:
                if float(y.get("val")) > 0.5:
                    return False, "标题被手动下移(y=%s)" % y.get("val")
            except ValueError:
                pass
    return True, "标题位于默认顶部"


def chart_get_series(chart_root):
    """返回系列列表，每个系列含 cat 引用、val 引用文本。"""
    series = []
    for ser in chart_root.iter("{%s}ser" % NS["c"]):
        cat_refs = [f.text for f in ser.findall(".//c:cat//c:f", NS) if f.text]
        val_refs = [f.text for f in ser.findall(".//c:val//c:f", NS) if f.text]
        cat_cache = [v.text for v in ser.findall(".//c:cat//c:pt/c:v", NS) if v.text]
        val_cache = [v.text for v in ser.findall(".//c:val//c:pt/c:v", NS) if v.text]
        series.append({
            "cat_refs": cat_refs, "val_refs": val_refs,
            "cat_cache": cat_cache, "val_cache": val_cache,
        })
    return series




def chart_is_doughnut_or_pie(chart_root):
    return (chart_root.find(".//c:doughnutChart", NS) is not None or
            chart_root.find(".//c:pieChart", NS) is not None or
            chart_root.find(".//c:ofPieChart", NS) is not None or
            chart_root.find(".//c:pie3DChart", NS) is not None)


def chart_ring_count(chart_root):
    """返回饼图或环形图中的系列数量；每个环形图系列对应一层圆环。"""
    chart_tags = ("doughnutChart", "pieChart", "ofPieChart", "pie3DChart")
    return sum(
        len(chart.findall("c:ser", NS))
        for tag in chart_tags
        for chart in chart_root.findall(f".//c:{tag}", NS)
    )



    return False


def chart_datalabel_quality(ctx, chart_root):
    """严格按细则三点判定“扇形数据标签”：
      (1) 扇形上有数据标签；
      (2) 至少显示数值或百分比之一；
      (3) 标签不大面积重叠且不遮挡主要扇区。

    返回 (ok, sub) ，sub 为各子条件布尔值字典。

    第(3)点为视觉性要求，无法直接读取像素，采用可度量的结构化近似实现其评估意图，
    但收紧到贴合饼/环图的实际渲染表现：
      - 饼/环图的 dLblPos 只有 'bestFit'/'ctr'/'inEnd'/'outEnd' 四种取值
        （'t'/'b'/'l'/'r' 是柱状/折线图专用，不适用于本判定）；
      - 'ctr'（扇区几何中心）与 'inEnd'（扇区内端）都直接压在扇区内部，
        点数一多就必然互相压盖、遮挡主要扇区内容 —— 不再视为安全布局；
      - 仅 'outEnd'（扇区外部，环图/饼图默认较安全的位置）视为安全；
      - 'bestFit' 本身会在点密集时自动退化为内部叠放，只有同时开启
        showLeaderLines（引导线，用于在标签被推到外部时仍能指回原扇区）
        才视为已采取有效防遮挡布局；
      - 重叠风险按“环”（系列）逐一核查，而非只看全局最大点数：
        任一环的数据点数 > 12（嵌套饼图六大类/四层级维度组合的常见密度，
        远低于原先 40 的宽松阈值）且该环标签为内部位置（inEnd/ctr，或
        bestFit 但无引导线），即判为该环大面积重叠/遮挡 -> 不通过。
    """
    sub = {}
    # (1)+(2)：存在 dLbls 且 showVal 或 showPercent 为 1
    has_label = False
    show_val_or_pct = False
    for dl in chart_root.iter("{%s}dLbls" % NS["c"]):
        has_label = True
        sv = dl.find("c:showVal", NS)
        sp = dl.find("c:showPercent", NS)
        if (sv is not None and sv.get("val") == "1") or (sp is not None and sp.get("val") == "1"):
            show_val_or_pct = True
    sub["1_扇形有数据标签"] = has_label
    sub["2_显示数值或百分比之一"] = show_val_or_pct

    # (3)：标签不大面积重叠且不遮挡主要扇区（按环逐一核查的结构化近似）
    leader = any(
        ll.get("val") == "1" for ll in chart_root.iter("{%s}showLeaderLines" % NS["c"]))
    OVERLAP_RISK_THRESHOLD = 12  # 单环数据点数超过该值即视为拥挤风险
    inside_pos = {"ctr", "inEnd"}

    any_overlap = False
    any_ring_checked = False
    for ser in chart_root.iter("{%s}ser" % NS["c"]):
        n_pts = max(
            len([v for v in ser.findall(".//c:cat//c:pt/c:v", NS) if v.text]),
            len([v for v in ser.findall(".//c:val//c:pt/c:v", NS) if v.text]),
        )
        ser_pos_vals = [p.get("val") for p in ser.iter("{%s}dLblPos" % NS["c"])]
        if not ser_pos_vals:
            # 该环未显式指定位置，饼/环图默认标签位置贴近扇区内部，按 'ctr' 处理
            ser_pos_vals = ["ctr"]
        any_ring_checked = True
        for p in ser_pos_vals:
            is_inside = (p in inside_pos) or (p == "bestFit" and not leader)
            if is_inside and n_pts > OVERLAP_RISK_THRESHOLD:
                any_overlap = True

    # 是否存在明确的安全放置策略：outEnd，或 bestFit+引导线
    pos_vals = [p.get("val") for p in chart_root.iter("{%s}dLblPos" % NS["c"])]
    has_safe_layout = ("outEnd" in pos_vals) or ("bestFit" in pos_vals and leader)

    sub["3_标签不重叠不遮挡"] = (
        has_safe_layout and not any_overlap if any_ring_checked
        else has_safe_layout
    )

    ok = all(sub.values())
    return ok, sub


def chart_has_legend(chart_root):
    return chart_root.find(".//c:legend", NS) is not None


def chart_color_variety(chart_root):
    """估计颜色种类数。

    varyColors=1 时不再只按点数粗略估算，而是按 Excel 实际的自动调色板循环推断：
    颜色种类至少是图表中数据点数量，但会按 6 色主题调色板循环。
    否则统计显式填充色/主题色数量。
    """
    vc = chart_root.find(".//c:varyColors", NS)
    if vc is not None and vc.get("val") == "1":
        max_pts = 0
        for s in chart_get_series(chart_root):
            max_pts = max(max_pts, len(s["cat_cache"]), len(s["val_cache"]))
        return max(max_pts, 6)
    colors = set()
    for srgb in chart_root.iter("{%s}srgbClr" % NS["a"]):
        colors.add(srgb.get("val"))
    for sch in chart_root.iter("{%s}schemeClr" % NS["a"]):
        colors.add("scheme:" + (sch.get("val") or ""))
    return len(colors)


def _clamp8(x):
    return max(0, min(255, int(round(x))))


def _hex_to_rgb(hexv):
    if not hexv:
        return None
    hexv = hexv.strip().lstrip("#")
    if len(hexv) != 6 or not re.match(r"^[0-9A-Fa-f]{6}$", hexv):
        return None
    return tuple(int(hexv[i:i + 2], 16) for i in (0, 2, 4))


def _rgb_distance(a, b):
    if a is None or b is None:
        return 10**9
    return ((a[0] - b[0]) ** 2 + (a[1] - b[1]) ** 2 + (a[2] - b[2]) ** 2) ** 0.5


def _apply_tint_shade(rgb, tint=None, shade=None, lum_mod=None, lum_off=None):
    """按 DrawingML 的常见色彩修饰近似计算实际 RGB。"""
    if rgb is None:
        return None
    r, g, b = rgb
    if shade is not None:
        try:
            f = int(shade) / 100000.0
            r, g, b = (_clamp8(r * f), _clamp8(g * f), _clamp8(b * f))
        except Exception:
            pass
    if tint is not None:
        try:
            f = int(tint) / 100000.0
            r = _clamp8(r + (255 - r) * f)
            g = _clamp8(g + (255 - g) * f)
            b = _clamp8(b + (255 - b) * f)
        except Exception:
            pass
    if lum_mod is not None or lum_off is not None:
        try:
            mod = int(lum_mod) / 100000.0 if lum_mod is not None else 1.0
            off = int(lum_off) / 100000.0 if lum_off is not None else 0.0
            r = _clamp8(r * mod + 255 * off)
            g = _clamp8(g * mod + 255 * off)
            b = _clamp8(b * mod + 255 * off)
        except Exception:
            pass
    return (r, g, b)


DEFAULT_SCHEME_RGB = {
    "accent1": (79, 129, 189),
    "accent2": (192, 80, 77),
    "accent3": (155, 187, 89),
    "accent4": (128, 100, 162),
    "accent5": (75, 172, 198),
    "accent6": (247, 150, 70),
    "dk1": (0, 0, 0),
    "lt1": (255, 255, 255),
    "dk2": (31, 73, 125),
    "lt2": (238, 236, 225),
}


def _load_theme_scheme_rgb(z):
    """读取 xl/theme/theme1.xml 中的主题色映射；缺失时回退到默认 Office 主题色。"""
    root = read_xml(z, "xl/theme/theme1.xml")
    if root is None:
        return dict(DEFAULT_SCHEME_RGB)
    scheme = root.find(".//a:clrScheme", NS)
    if scheme is None:
        return dict(DEFAULT_SCHEME_RGB)
    out = dict(DEFAULT_SCHEME_RGB)
    for child in list(scheme):
        tag = child.tag.rsplit("}", 1)[-1]
        clr = child.find("a:srgbClr", NS)
        if clr is not None and clr.get("val"):
            rgb = _hex_to_rgb(clr.get("val"))
            if rgb is not None:
                out[tag] = rgb
                continue
        sys_clr = child.find("a:sysClr", NS)
        if sys_clr is not None and sys_clr.get("lastClr"):
            rgb = _hex_to_rgb(sys_clr.get("lastClr"))
            if rgb is not None:
                out[tag] = rgb
    return out


def _resolve_scheme_color(z, scheme_name, modifiers=None):
    """把 schemeClr 解析为实际 RGB。"""
    scheme_rgb = _load_theme_scheme_rgb(z)
    rgb = scheme_rgb.get(scheme_name)
    if rgb is None:
        return None
    modifiers = modifiers or {}
    return _apply_tint_shade(rgb, modifiers.get("tint"), modifiers.get("shade"), modifiers.get("lumMod"), modifiers.get("lumOff"))


def _extract_color_rgb(z, node):
    """从 a:srgbClr / a:schemeClr 节点提取实际 RGB。"""
    if node is None:
        return None
    if node.tag.endswith("srgbClr"):
        return _hex_to_rgb(node.get("val"))
    if node.tag.endswith("schemeClr"):
        mods = {}
        for k in ("tint", "shade", "lumMod", "lumOff"):
            el = node.find("a:%s" % k, NS)
            if el is not None and el.get("val") is not None:
                mods[k] = el.get("val")
        return _resolve_scheme_color(z, node.get("val"), mods)
    return None


def _series_point_colors(ctx, ser):
    """提取单个系列中每个点的实际 RGB 颜色。"""
    z = ctx["zip"]
    vary = ser.find(".//c:varyColors", NS)
    vary_on = (vary is not None and vary.get("val") == "1")
    theme_cycle = ["accent1", "accent2", "accent3", "accent4", "accent5", "accent6"]
    points = []
    for dpt in ser.findall("c:dPt", NS):
        idx_el = dpt.find("c:idx", NS)
        if idx_el is None:
            continue
        idx = int(idx_el.get("val"))
        fill = dpt.find(".//a:solidFill/*", NS)
        points.append((idx, _extract_color_rgb(z, fill)))
    point_map = {idx: rgb for idx, rgb in points if rgb is not None}

    # 自动配色：按 Excel 实际调色板顺序循环
    auto_map = {}
    if vary_on:
        cat_count = 0
        cats = []
        if ser.find(".//c:cat", NS) is not None:
            cats = [v.text for v in ser.findall(".//c:cat//c:pt/c:v", NS) if v.text]
            if not cats:
                for f in ser.findall(".//c:cat//c:f", NS):
                    if f.text:
                        cats.extend(str(v) for v in get_referenced_cell_values(ctx, f.text) if v)
        cat_count = len(cats)
        for i in range(cat_count):
            scheme = theme_cycle[i % len(theme_cycle)]
            auto_map[i] = _resolve_scheme_color(z, scheme)
    return point_map, auto_map


def chart_category_color_map(ctx, chart_root):
    """构造 图表中“成绩类别 -> 实际RGB颜色”的映射。

    取值优先级：
      1) 若数据点(dPt)存在显式填充色，则按点序号映射到类别标签，取实际 RGB；
      2) 否则在 varyColors=1 下，按 Excel 的实际自动调色板顺序循环得到 RGB，
         再按类别标签映射。

    返回 dict: 分类标签 -> RGB 三元组。
    """
    color_map = {}
    for ser in chart_root.iter("{%s}ser" % NS["c"]):
        cats = [v.text for v in ser.findall(".//c:cat//c:pt/c:v", NS) if v.text]
        if not cats:
            for f in ser.findall(".//c:cat//c:f", NS):
                if f.text:
                    cats.extend(str(v) for v in get_referenced_cell_values(ctx, f.text) if v)
        point_map, auto_map = _series_point_colors(ctx, ser)
        for pi, label in enumerate(cats):
            if pi in point_map:
                color_map[label] = point_map[pi]
            elif pi in auto_map:
                color_map.setdefault(label, auto_map[pi])
    return color_map


def chart_distinct_colors(ctx, chart_root):
    """图中实际出现的不同颜色数（按 RGB 去重）。"""
    cm = chart_category_color_map(ctx, chart_root)
    return len({v for v in cm.values() if v is not None})


def parse_ref_sheet(ref):
    """从 'sheet'!$A$1 形式的引用中取出工作表名。"""
    m = re.match(r"(?:'?([^'!]+)'?)!", ref)
    return m.group(1) if m else None


def iter_ref_cells(ref):
    """遍历 'sheet'!$A$1:$B$2 形式引用覆盖的所有 (row, col)。"""
    m = re.search(r"!\$?([A-Z]+)\$?(\d+)(?::\$?([A-Z]+)\$?(\d+))?", ref)
    if not m:
        return
    c1, r1, c2, r2 = m.groups()
    ci1, ri1 = col_letter_to_idx(c1), int(r1)
    ci2, ri2 = (col_letter_to_idx(c2), int(r2)) if c2 else (ci1, ri1)
    for rr in range(min(ri1, ri2), max(ri1, ri2) + 1):
        for cc in range(min(ci1, ci2), max(ci1, ci2) + 1):
            yield (rr, cc)


def get_sheet_grid(ctx, sheet_name):
    """获取（必要时解析）指定工作表的单元格网格。"""
    grid = ctx.get("grids", {}).get(sheet_name)
    if grid is None:
        sp = ctx.get("sheet_xml", {}).get(sheet_name)
        if sp:
            grid = parse_sheet_grid(ctx["zip"], sp, ctx)
            ctx.setdefault("grids", {})[sheet_name] = grid
        else:
            grid = {}
    return grid


def get_cells_text(ctx, sheet_name, cells):
    """取指定工作表若干 (row, col) 单元格的文本值列表。"""
    grid = get_sheet_grid(ctx, sheet_name)
    out = []
    for (rr, cc) in cells:
        v = grid.get((rr, cc))
        out.append(str(v).strip() if v is not None else "")
    return out


def get_referenced_cell_values(ctx, ref):
    """根据 'sheet'!$A$1:$B$2 形式的引用，从已解析 grid 取出值列表。"""
    m = re.match(r"(?:'?([^'!]+)'?)!\$?([A-Z]+)\$?(\d+)(?::\$?([A-Z]+)\$?(\d+))?", ref)
    if not m:
        return []
    sheet, c1, r1, c2, r2 = m.groups()
    grid = ctx.get("grids", {}).get(sheet)
    if grid is None:
        # 按需解析
        sp = ctx.get("sheet_xml", {}).get(sheet)
        if sp:
            grid = parse_sheet_grid(ctx["zip"], sp, ctx)
            ctx.setdefault("grids", {})[sheet] = grid
        else:
            return []
    ci1, ri1 = col_letter_to_idx(c1), int(r1)
    ci2, ri2 = (col_letter_to_idx(c2), int(r2)) if c2 else (ci1, ri1)
    vals = []
    for rr in range(min(ri1, ri2), max(ri1, ri2) + 1):
        for cc in range(min(ci1, ci2), max(ci1, ci2) + 1):
            vals.append(grid.get((rr, cc)))
    return vals


def chart_all_category_text(ctx, chart_root):
    """收集图表中所有分类标签文本（来自 cat 缓存或引用单元格）。"""
    texts = []
    for s in chart_get_series(chart_root):
        texts.extend([t for t in s["cat_cache"] if t])
        if not s["cat_cache"]:
            for ref in s["cat_refs"]:
                texts.extend([str(v) for v in get_referenced_cell_values(ctx, ref) if v])
    return texts


def chart_is_nested_doughnut(chart_root):
    """是否为“多类别嵌套合并饼图”：饼/环图且含多个系列（多环嵌套）。"""
    return chart_is_doughnut_or_pie(chart_root) and chart_ring_count(chart_root) >= 2


def series_category_texts(ctx, ser):
    """取单个系列(单环)的全部分类标签文本。"""
    texts = [t for t in ser["cat_cache"] if t]
    if not texts:
        for ref in ser["cat_refs"]:
            texts.extend(str(v) for v in get_referenced_cell_values(ctx, ref) if v)
    return texts


def _ring_dimension(labels):
    """判断单个圆环(系列)承载的是哪个“纯维度”。

    要求该环的分类标签是“干净的单一维度”：
      - 'cat'：标签集合恰好就是六大类(地图基础..聚落形态)，且不掺入任何四层级字样；
      - 'lvl'：标签集合恰好就是四层级(卓越/稳定/基本/持续)，且不掺入任何六大类字样；
      - None：既不是纯六大类也不是纯四层级（例如“大类-能力-层级”这类组合串）。

    “干净”的核心判据：每一个标签本身只能对应唯一维度的唯一一项，
    不允许单个标签里同时出现两个维度的关键词（组合标签一律判 None）。
    """
    labels = [str(t).strip() for t in labels if str(t).strip()]
    if not labels:
        return None
    cat_set = set(COL_CATEGORIES)   # 六大类
    lvl_set = set(ROW_LEVELS)       # 四层级

    saw_cat = set()
    saw_lvl = set()
    for lab in labels:
        hit_cat = [k for k in cat_set if k in lab]
        hit_lvl = [k for k in lvl_set if k in lab]
        # 单个标签同时命中两个维度 -> 组合标签，非纯维度
        if hit_cat and hit_lvl:
            return None
        # 单个标签命中维度内多于一项(罕见) -> 视为不纯
        if len(hit_cat) > 1 or len(hit_lvl) > 1:
            return None
        if hit_cat:
            saw_cat.add(hit_cat[0])
        elif hit_lvl:
            saw_lvl.add(hit_lvl[0])
        else:
            # 出现既非大类也非层级的标签 -> 不是干净的目标维度环
            return None

    # 标签必须全部属于同一维度，且该维度完整覆盖
    if saw_cat and not saw_lvl and saw_cat == cat_set:
        return "cat"
    if saw_lvl and not saw_cat and saw_lvl == lvl_set:
        return "lvl"
    return None


def classify_nested_scheme(ctx, chart_root):
    """判断嵌套环图的分类统计组织方式是否符合细则两种合法方式之一。

    细则：
      方式A：饼图统计对象 = 四类成绩(卓越/稳定/基本/持续)，圆环层级 = 六大类(地图基础..聚落形态)
      方式B：饼图统计对象 = 六大类(地图基础..聚落形态)，圆环层级 = 四类成绩(卓越/稳定/基本/持续)

    返回 ('A'|'B'|None, reason)。
    严格判定（不再接受组合标签蒙混）：
      - 必须是多环嵌套饼/环图；
      - “统计对象”取最内层(idx=0)系列，必须是“干净的单一维度”整环
        （纯六大类，或纯四层级——单个标签不得同时含两维度关键词）；
      - “圆环层级”由其余各环承载，必须存在某一外环是“另一个维度”的干净整环；
      - 两个维度(六大类 / 四层级)需各自由一个纯维度环完整承载，互不混杂。
    """
    if not chart_is_doughnut_or_pie(chart_root):
        return None, "非饼/环图"
    sers = chart_get_series(chart_root)
    if len(sers) < 2:
        return None, "非多环嵌套(系列<2)"

    # 逐环判定承载的纯维度：'cat'(纯六大类) / 'lvl'(纯四层级) / None(组合或不纯)
    ring_dims = [_ring_dimension(series_category_texts(ctx, s)) for s in sers]
    inner_dim = ring_dims[0]                 # 统计对象 = 最内层
    outer_dims = set(ring_dims[1:])          # 圆环层级 = 其余各环

    # 方式A：统计对象=纯四层级，且某外环=纯六大类
    if inner_dim == "lvl" and "cat" in outer_dims:
        return "A", "内层=纯四类成绩,外环含纯六大类"
    # 方式B：统计对象=纯六大类，且某外环=纯四层级
    if inner_dim == "cat" and "lvl" in outer_dims:
        return "B", "内层=纯六大类,外环含纯四类成绩"

    return None, "统计对象维度=%s 圆环层级维度=%s(组合/混杂标签不计为合法分类统计)" % (
        inner_dim if inner_dim else "非纯维度",
        sorted(d for d in outer_dims if d) if any(outer_dims) else "无纯维度环")


def ref_range_within_A1Q7(ref):
    """判断单个引用 'sheet'!$A$1:$B$2 覆盖的单元格是否全部落在 A1:Q7 范围内
    （行1..7、列A..Q，即列1..17）。只要引用越出该范围（如引用了辅助列/辅助行、
    或 A1:Q7 之外的任意区域），即返回 False。"""
    cells = list(iter_ref_cells(ref))
    if not cells:
        return False
    return all(1 <= rr <= 7 and 1 <= cc <= 17 for (rr, cc) in cells)


def chart_class_data_from_A1Q7(ctx, chart_root, class_name):
    """判断图表数据源是否“均来自对应工作表 A1:Q7 中的成绩数据”。

    严格判定（不再接受“60%数值能在A1:Q7值集合中找到即视为derived”的宽松近似）：
      - 所有系列的分类/数值引用都必须指向 class_name 对应的工作表（来源正确）；
      - 且每一条引用的单元格区域本身必须完整落在该表 A1:Q7 范围内
        （不允许引用同表任意区域，如引用 A1:Q7 之外的辅助列/辅助行）；
      - 若图表系列没有可核验的引用（仅有静态缓存值、无 cat_refs/val_refs），
        视为“不可追溯到 A1:Q7”，不满足数据源要求。
    返回 (all_refs_in_class_sheet, derived_from_A1Q7)。
      - all_refs_in_class_sheet：所有引用都指向对应班级工作表；
      - derived_from_A1Q7：在上一条基础上，所有引用区域都落在 A1:Q7 内
        （即可直接追溯到 A1:Q7 的原始成绩数据，而非同表其他区域）。
    """
    all_refs_in_sheet = True
    all_refs_in_A1Q7 = True
    has_any_ref = False
    for s in chart_get_series(chart_root):
        refs = s["cat_refs"] + s["val_refs"]
        if not refs:
            # 无引用可核验（例如仅静态缓存值，无法追溯来源）-> 视为不满足可追溯要求
            all_refs_in_A1Q7 = False
            continue
        for ref in refs:
            has_any_ref = True
            sheet = parse_ref_sheet(ref)
            if sheet != class_name:
                all_refs_in_sheet = False
                all_refs_in_A1Q7 = False
                continue
            if not ref_range_within_A1Q7(ref):
                all_refs_in_A1Q7 = False

    if not has_any_ref:
        all_refs_in_sheet = False
        all_refs_in_A1Q7 = False

    return all_refs_in_sheet, all_refs_in_A1Q7


def chart_covers_nonempty_columns(ctx, chart_root, class_name):
    """覆盖“地图基础”到“聚落形态”各非空能力层级数据，不遗漏主要成绩列。

    严格判定（不再只检查六个类别名称是否出现在图表分类文本中）：
      (1) 六大类维度覆盖：B1:Q1 中所有非空的能力大类（地图基础..聚落形态）
          都必须作为图表某一环的分类标签完整出现；
      (2) 四层级维度覆盖：A4:A7 中所有非空的能力层级（卓越表现/稳定达成/
          基本达成/持续提升）都必须作为图表某一环的分类标签完整出现；
      (3) 主要成绩列数值覆盖：对每一个非空大类 × 每一个非空层级构成的
          B4:Q7 成绩单元格（该细则所指“主要成绩列”），其数值都必须能在
          图表系列的数值缓存/引用值中找到，即逐列/逐层级确认非空成绩数据
          均被统计，不允许仅凑够比例或只覆盖部分列。
    返回 (ok, present_cats, missing)：
      - present_cats：非空的六大类名称列表；
      - missing：dict，包含 "categories"（缺失的大类）、"levels"（缺失的层级）、
        "values"（未被图表数值覆盖的 (层级, 大类) 成绩列）三类缺项，全部为空
        列表才算不遗漏。
    """
    grid = get_sheet_grid(ctx, class_name)

    # B1:Q1 非空大类（按合并表头，取出现的 6 个大类名，去重保序）
    present_cats = []
    for cc in range(2, 18):            # B(2)..Q(17)
        v = grid.get((1, cc))
        if v is not None and str(v).strip() != "":
            vv = str(v).strip()
            if vv in COL_CATEGORIES and vv not in present_cats:
                present_cats.append(vv)

    # A4:A7 非空层级（去重保序）
    present_lvls = []
    for rr in range(4, 8):
        v = grid.get((rr, 1))
        if v is not None and str(v).strip() != "":
            vv = str(v).strip()
            if vv in ROW_LEVELS and vv not in present_lvls:
                present_lvls.append(vv)

    # 图表中各环的分类标签集合（逐环收集，用于判定“该维度是否完整出现”）
    chart_cat_labels = set()
    for s in chart_get_series(chart_root):
        for t in series_category_texts(ctx, s):
            t = str(t).strip()
            if t:
                chart_cat_labels.add(t)

    missing_cats = [k for k in present_cats if k not in chart_cat_labels]
    missing_lvls = [k for k in present_lvls if k not in chart_cat_labels]

    # 图表实际用到的全部数值（数值缓存 + 数值引用解析值），逐列/逐层级核对来源
    chart_values = set()
    for s in chart_get_series(chart_root):
        for v in s["val_cache"]:
            if v is not None and str(v).strip() != "":
                chart_values.add(str(v).strip())
        for ref in s["val_refs"]:
            for v in get_referenced_cell_values(ctx, ref):
                if v is not None and str(v).strip() != "":
                    chart_values.add(str(v).strip())

    # 逐列(大类)/逐层级确认 B4:Q7 中每一个非空成绩数据都被图表数值覆盖
    missing_values = []
    col_idx_by_cat = {}
    for cc in range(2, 18):
        v = grid.get((1, cc))
        if v is not None and str(v).strip() in COL_CATEGORIES:
            col_idx_by_cat.setdefault(str(v).strip(), []).append(cc)
    row_idx_by_lvl = {}
    for rr in range(4, 8):
        v = grid.get((rr, 1))
        if v is not None and str(v).strip() in ROW_LEVELS:
            row_idx_by_lvl[str(v).strip()] = rr

    for lvl in present_lvls:
        rr = row_idx_by_lvl.get(lvl)
        if rr is None:
            continue
        for cat in present_cats:
            for cc in col_idx_by_cat.get(cat, []):
                cell_v = grid.get((rr, cc))
                if cell_v is None or str(cell_v).strip() == "":
                    continue
                if str(cell_v).strip() not in chart_values:
                    missing_values.append((lvl, cat, "R%dC%d" % (rr, cc)))

    missing = {
        "categories": missing_cats,
        "levels": missing_lvls,
        "values": missing_values,
    }
    ok = (len(missing_cats) == 0 and len(missing_lvls) == 0 and
          len(missing_values) == 0 and len(present_cats) > 0 and len(present_lvls) > 0)
    return ok, present_cats, missing



# =========================================================================
# 维度2：完成度评分细则
# =========================================================================
def check_dimension2(ctx):
    z = ctx["zip"]
    names = ctx["names"]
    chart_files = sorted(ctx["chart_files"])
    charts = []
    for cf in chart_files:
        root = read_xml(z, cf)
        if root is not None:
            charts.append((cf, root))

    results = []  # (score, hit:bool, label)

    # 将图表与班级匹配（通过标题/数据源引用判断）
    chart_by_class = {}
    for cf, root in charts:
        title = chart_get_title(root)
        refs = " ".join(
            r for s in chart_get_series(root) for r in (s["cat_refs"] + s["val_refs"])
        )
        for cls in SHEET_NAMES:
            if cls in title or ("'%s'" % cls) in refs or ("%s!" % cls) in refs:
                # 精确匹配：七1班 不应误配到 七13班
                if cls == "七1班" and ("七13班" in title or "七13班" in refs):
                    continue
                chart_by_class.setdefault(cls, []).append((cf, root))

    # ---- 得分点 +1：严格按细则逐点判定（每个点都必须满足才加分）----
    # 细则原文拆解为以下原子子条件：
    #   (a) 分别基于“七1班”和“七13班”两份成绩数据，各生成 1 个“多类别嵌套合并饼图”
    #   (b) 共 2 个图表
    #   (c) 两个图表标题能明确区分“七1班”和“七13班”
    #   (d) 两个图表的数据源均来自对应工作表 A1:Q7 中的成绩数据
    #   (e) 覆盖“地图基础”到“聚落形态”各非空能力层级数据，不遗漏主要成绩列
    sub = {}

    # (a) 两个班级各有 1 个多类别嵌套合并饼图（饼/环图且为多环嵌套）
    cls_nested = {}
    for c in SHEET_NAMES:
        lst = chart_by_class.get(c, [])
        cls_nested[c] = (len(lst) >= 1 and chart_is_nested_doughnut(lst[0][1]))
    sub["a_各班1个嵌套合并饼图"] = all(cls_nested[c] for c in SHEET_NAMES)

    # (b) 共 2 个图表
    sub["b_共2个图表"] = (len(charts) == 2)

    cond_two_charts = all(len(chart_by_class.get(c, [])) >= 1 for c in SHEET_NAMES)

    # (c) 标题明确区分两个班级：七1班图标题含“七1班”不含“七13班”；七13班图标题含“七13班”
    if cond_two_charts:
        t1 = chart_get_title(chart_by_class["七1班"][0][1])
        t13 = chart_get_title(chart_by_class["七13班"][0][1])
        sub["c_标题区分两班"] = ("七1班" in t1 and "七13班" not in t1 and "七13班" in t13)
    else:
        sub["c_标题区分两班"] = False

    # (d) 数据源均来自对应工作表 A1:Q7 中的成绩数据
    if cond_two_charts:
        d_ok = True
        for c in SHEET_NAMES:
            in_sheet, derived = chart_class_data_from_A1Q7(ctx, chart_by_class[c][0][1], c)
            if not (in_sheet and derived):
                d_ok = False
        sub["d_数据源来自对应表A1Q7"] = d_ok
    else:
        sub["d_数据源来自对应表A1Q7"] = False

    # (e) 覆盖“地图基础”到“聚落形态”各非空能力层级，不遗漏主要成绩列
    if cond_two_charts:
        e_ok = True
        e_missing = {}
        for c in SHEET_NAMES:
            ok_cov, present, missing = chart_covers_nonempty_columns(
                ctx, chart_by_class[c][0][1], c)
            if not ok_cov:
                e_ok = False
            if missing["categories"] or missing["levels"] or missing["values"]:
                e_missing[c] = missing
        sub["e_覆盖各非空能力列不遗漏"] = e_ok
    else:
        sub["e_覆盖各非空能力列不遗漏"] = False
        e_missing = {}

    # 每一个子条件都满足才加分
    p1_hit = all(sub.values())
    results.append((1, p1_hit, "p1_charts",
        "+1 两班各基于本班成绩各生成1个多类别嵌套合并饼图(共2个)、标题区分七1/七13班、"
        "数据源均来自对应表A1:Q7、覆盖地图基础~聚落形态各非空能力列不遗漏"
        "（a各班嵌套饼图=%s b共2图=%s c标题区分=%s d数据源A1:Q7=%s e覆盖不遗漏=%s%s）"
        % (sub["a_各班1个嵌套合并饼图"], sub["b_共2个图表"], sub["c_标题区分两班"],
           sub["d_数据源来自对应表A1Q7"], sub["e_覆盖各非空能力列不遗漏"],
           ("" if not e_missing else " 缺列=%s" % e_missing))))

    # ---- 得分点 +5：分类统计正确（严格按细则两种合法组织方式之一，逐点判定）----
    # 细则给出两种“分类统计正确”的合法方式，命中任一种即算正确：
    #   方式A：以“卓越表现/稳定达成/基本达成/持续提升”四类成绩数据作为饼图统计对象，
    #          以“地图基础/经纬定位/地形判读/气候要素/区域差异/聚落形态”六类作为不同类别圆环层级。
    #   方式B：以“地图基础/经纬定位/地形判读/气候要素/区域差异/聚落形态”六类成绩数据作为饼图统计对象，
    #          以“卓越表现/稳定达成/基本达成/持续提升”四类作为不同类别圆环层级。
    # 判定要求两个图表都满足同一种合法方式（A 或 B），逐点核对四层级集合与六大类集合。
    p5_hit = False
    p5_msgs = []
    if cond_two_charts:
        ok_all = True
        for c in SHEET_NAMES:
            root = chart_by_class[c][0][1]
            scheme, why = classify_nested_scheme(ctx, root)
            p5_msgs.append("%s:%s" % (c, scheme if scheme else ("不符合(%s)" % why)))
            if scheme not in ("A", "B"):
                ok_all = False
        p5_hit = ok_all
    else:
        p5_msgs.append("缺少两个图表")
    results.append((5, p5_hit, "p5_class",
        "+5 分类统计正确（方式A:四类成绩为统计对象+六大类为圆环层级；"
        "或方式B:六大类为统计对象+四类成绩为圆环层级）（%s）"
        % "; ".join(p5_msgs)))

    # ---- 得分点 +1：严格按细则三点判定扇形数据标签 ----
    # 细则拆解：
    #   (1) 每个图表的扇形上有数据标签；
    #   (2) 至少显示数值或百分比之一；
    #   (3) 标签不大面积重叠且不遮挡主要扇区。
    # 每个点都必须满足，且两个图表都满足，才加分。
    p1b_hit = False
    p1b_msgs = []
    if cond_two_charts:
        ok_all = True
        for c in SHEET_NAMES:
            ok, sub = chart_datalabel_quality(ctx, chart_by_class[c][0][1])
            p1b_msgs.append("%s(有标签=%s 数值或百分比=%s 不重叠不遮挡=%s)" % (
                c, sub["1_扇形有数据标签"], sub["2_显示数值或百分比之一"],
                sub["3_标签不重叠不遮挡"]))
            if not ok:
                ok_all = False
        p1b_hit = ok_all
    else:
        p1b_msgs.append("缺少两个图表")
    results.append((1, p1b_hit, "p1_label",
        "+1 每个图表扇形有数据标签、至少显示数值或百分比之一、标签不大面积重叠且不遮挡主要扇区"
        "（%s）" % "; ".join(p1b_msgs)))

    # ---- 得分点 +1：严格按细则判定图例 ----
    # 细则拆解：
    #   (1) 两个图表均包含图例；
    #   (2) 图例文字 包含或能对应 B1:Q1 的6个成绩类别(地图基础/经纬定位/地形判读/气候要素/区域差异/聚落形态)
    #        或 包含或能对应 A4:A7 的4个能力层级(卓越表现/稳定达成/基本达成/持续提升)。
    # 两个图表都必须满足(1)且满足(2)，才加分。
    p1c_hit = False
    p1c_msgs = []
    if cond_two_charts:
        ok_all = True
        for c in SHEET_NAMES:
            root = chart_by_class[c][0][1]
            has_leg = chart_has_legend(root)
            # 取该班 B1:Q1 与 A4:A7 的实际文字作为对照标准
            cat6 = get_cells_text(ctx, c, [(1, cc) for cc in range(2, 18)])   # B1:Q1
            lvl4 = get_cells_text(ctx, c, [(rr, 1) for rr in range(4, 8)])    # A4:A7
            cat6 = [t for t in cat6 if t]
            lvl4 = [t for t in lvl4 if t]
            # 图例文字 = 图表分类标签文字（饼/环图图例项来自各扇区分类名）
            legend_text = "".join(chart_all_category_text(ctx, root))
            match_cat = sum(1 for k in cat6 if k in legend_text)
            match_lvl = sum(1 for k in lvl4 if k in legend_text)
            corresp = (match_cat >= len(cat6) and len(cat6) > 0) or \
                      (match_lvl >= len(lvl4) and len(lvl4) > 0)
            p1c_msgs.append("%s(含图例=%s 对应B1:Q1=%d/%d 对应A4:A7=%d/%d)" % (
                c, has_leg, match_cat, len(cat6), match_lvl, len(lvl4)))
            if not (has_leg and corresp):
                ok_all = False
        p1c_hit = ok_all
    else:
        p1c_msgs.append("缺少两个图表")
    results.append((1, p1c_hit, "p1_legend",
        "+1 两图均含图例，且图例文字包含/对应B1:Q1的6个成绩类别 或 A4:A7的4个能力层级"
        "（%s）" % "; ".join(p1c_msgs)))

    # ---- 得分点 +1：严格按细则判定颜色方案 ----
    # 细则拆解：
    #   (1) 两个图表中“同一成绩类别”使用一致或接近一致的颜色方案（便于比较两个班级数据）；
    #   (2) 图中颜色至少四种。
    # 两点都要满足才加分。
    p1d_hit = False
    color_msg = "缺少两个图表"
    if cond_two_charts:
        cm1 = chart_category_color_map(ctx, chart_by_class["七1班"][0][1])
        cm2 = chart_category_color_map(ctx, chart_by_class["七13班"][0][1])
        # (1) 同一类别颜色一致/接近一致：按实际 RGB 颜色距离判断，不再要求字符串完全相等。
        #     对 varyColors 自动配色，则以类别标签对应到的调色板色位为准；只要视觉接近即可。
        common = [k for k in cm1 if k in cm2]
        if common:
            close = 0
            for k in common:
                if _rgb_distance(cm1[k], cm2[k]) <= 36:
                    close += 1
            ratio = close / len(common)
        else:
            ratio = 0.0
        consistent = (len(common) > 0 and ratio >= 0.8)   # 一致或“接近一致”
        # (2) 图中颜色至少四种（两个图表各自都需 >=4 种）
        colors1 = chart_distinct_colors(ctx, chart_by_class["七1班"][0][1])
        colors2 = chart_distinct_colors(ctx, chart_by_class["七13班"][0][1])
        enough_colors = (colors1 >= 4 and colors2 >= 4)
        p1d_hit = consistent and enough_colors
        color_msg = "同类别颜色一致度=%.0f%%(共有类别%d个) 各图颜色数=[%d,%d] 至少4种=%s" % (
            ratio * 100, len(common), colors1, colors2, enough_colors)
    results.append((1, p1d_hit, "p1_color",
        "+1 两图同一成绩类别颜色一致/接近一致、图中颜色至少4种（%s）" % color_msg))

    # ---- 得分点 +1：严格按细则判定图表标题 ----
    # 细则拆解：
    #   (1) 图表上方出现图表标题；
    #   (2) 标题内容分别为“七1班 探究成绩多类别嵌套饼图”及“七13班 探究成绩多类别嵌套饼图”。
    # 两个图表都须满足(1)上方有标题 且 (2)标题内容精确匹配各自规定文本，才加分。
    p1e_hit = False
    p1e_msgs = []
    if cond_two_charts:
        ok_all = True
        for c in SHEET_NAMES:
            root = chart_by_class[c][0][1]
            top_ok, why = chart_title_present_and_top(root)
            title = chart_get_title(root).strip()
            content_ok = (title == EXPECTED_TITLES[c])
            p1e_msgs.append("%s(上方有标题=%s[%s] 内容匹配=%s:'%s')" % (
                c, top_ok, why, content_ok, title))
            if not (top_ok and content_ok):
                ok_all = False
        p1e_hit = ok_all
    else:
        p1e_msgs.append("缺少两个图表")
    results.append((1, p1e_hit, "p1_title",
        "+1 图表上方出现标题，内容分别为“七1班 探究成绩多类别嵌套饼图”“七13班 探究成绩多类别嵌套饼图”"
        "（%s）" % "; ".join(p1e_msgs)))

    return results


# =========================================================================
# 主流程：统一入口 evaluate(dir_path)
# =========================================================================
def _dim2_max_score() -> int:
    """维度2满分（所有正向得分点 max_delta 之和），
    用于维度1未通过时也能给出统一的 max_score。"""
    return 1 + 5 + 1 + 1 + 1 + 1  # 对应 6 个 +N 得分点：p1_charts/p5_class/p1_label/p1_legend/p1_color/p1_title


def evaluate(dir_path):
    """统一入口：接收脚本所在目录路径，脚本负责在该目录内定位并打开被评估文档。

    参数：
        dir_path: 脚本所在目录路径（str）。脚本在该目录中定位并打开被评估文档。

    返回：
        dict：结构遵循《脚本接口差异与统一建议.md》§2.2 约定。

    不 print 主结果、不修改 sys.stdout、不 sys.exit、不硬编码路径。
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
        "max_score": _dim2_max_score(),
    }

    try:
        # 在给定目录内定位被评估文档：优先与脚本约定的 TARGET_FILE 同名文件；
        # 若不存在则回退为目录内第一个 .xlsx/.xlsm 文件。
        path = os.path.join(dir_path, TARGET_FILE)
        if not os.path.exists(path):
            try:
                candidates = sorted(
                    f for f in os.listdir(dir_path)
                    if f.lower().endswith((".xlsx", ".xlsm"))
                )
            except OSError as e:
                result["status"] = "error"
                result["error"] = "无法读取目录：%s（%s）" % (dir_path, e)
                return result
            if candidates:
                path = os.path.join(dir_path, candidates[0])

        result["file_name"] = os.path.basename(path)

        if not os.path.exists(path):
            result["status"] = "error"
            result["error"] = "文件不存在：%s" % path
            return result

        ctx = {}

        # ---------- 维度1 ----------
        d1_pass, d1_details = check_dimension1(path, ctx)
        result["dim1_pass"] = d1_pass
        if not d1_pass:
            # 维度1 未全部通过 -> 直接 0 分，不再检查维度2
            fail_reasons = [msg for ok, msg in d1_details if not ok]
            result["dim1_reason"] = "; ".join(fail_reasons) if fail_reasons else "未通过"
            result["total_score"] = 0
            return result

        # ---------- 维度2 ----------
        d2_results = check_dimension2(ctx)
        total = 0
        items = []
        for score, hit, rule_key, detail in d2_results:
            delta = score if hit else 0
            total += delta
            items.append({
                "rule": RULE_TEXTS.get(rule_key, rule_key),
                "max_delta": score,
                "delta": delta,
                "hit": hit,
                "detail": "",
            })
        result["dim2_items"] = items
        result["total_score"] = total
        return result

    except Exception as e:
        result["status"] = "error"
        result["error"] = "%s: %s" % (type(e).__name__, e)
        return result


if __name__ == "__main__":
    # 本地调试用：默认使用脚本所在目录；也可通过命令行覆盖。
    import json

    _dir = sys.argv[1] if len(sys.argv) > 1 else os.path.dirname(os.path.abspath(__file__))
    print(json.dumps(evaluate(_dir), ensure_ascii=False, indent=2))
