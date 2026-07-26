# -*- coding: utf-8 -*-
"""
自动评估：出货量对比-变化趋势图_可编辑.xlsx

评估流程：
  维度1（可用与可修改性）为准入门槛 —— 任一条不满足 => 直接 0 分，不再检查维度2。
  维度2（完成度）分为“加分点”和“扣分点”：
     - 加分细则：必须满足该条内的“每一个”子点，才累加该条分数（可正可负）。
     - 扣分细则：只要命中该条内的“任意一个”子点，即扣该条分数。
  最终打印命中的点与总分。

说明（灵活变通处，均在报告中标注）：
  1) 细则文字里的“sheet1”指“承载区域-年份-销售额数据并用于绘图的数据工作表”。
     本文件中绘图数据表位于 Sheet2（其数值以公式链接回 Sheet1 原始数据），
     图表也放置在 Sheet2 上，故以“图表所在且含销售数据的工作表”作为该数据表评估。
  2) 字体/字号/颜色/线宽等：图表 XML 未显式设置时按 Excel 默认处理：
     - 默认坐标轴/标签字号≈10磅（落在 8–12磅内）、默认颜色为黑色（自动）；
     - 对“要求特定非默认样式”的子点（如标题需 16–20磅/加粗/深蓝色），未显式设置视为不满足；
     - 对“允许默认即可满足范围”的子点（如标签 8–12磅、黑色，网格线/坐标轴线存在等），默认视为满足。
"""

import json
import os
import sys
import zipfile
import xml.etree.ElementTree as ET

TARGET = "出货量对比-变化趋势图_可编辑.xlsx"
SCRIPT_ID = "083"

# ----------------------------------------------------------------------------
# 通用工具：命名空间无关的 XML 查找（Python 3.8+ 支持 '{*}tag' 通配）
# ----------------------------------------------------------------------------

def findall(el, tag):
    return el.findall(".//{*}" + tag) if el is not None else []

def find(el, tag):
    return el.find(".//{*}" + tag) if el is not None else None

def direct(el, tag):
    """只在直接子节点中查找。"""
    if el is None:
        return None
    for c in el:
        if c.tag.split("}")[-1] == tag:
            return c
    return None

def directall(el, tag):
    out = []
    if el is None:
        return out
    for c in el:
        if c.tag.split("}")[-1] == tag:
            out.append(c)
    return out

def text_of(el):
    return "".join(t.text or "" for t in findall(el, "t")) if el is not None else ""


# ----------------------------------------------------------------------------
# 字体 / 字号 / 颜色 / 加粗 探测
#   Excel 用 a:rPr (rich text) 或 a:defRPr (txPr) 描述字符样式：
#     sz  = 字号 * 100（如 1800 => 18磅）
#     b   = 1 表示加粗
#     a:solidFill/a:srgbClr@val 表示颜色
#     a:latin@typeface 表示字体
# ----------------------------------------------------------------------------

DARK_BLUE_HINTS = ["1F", "0F", "00", "17", "1A", "2E", "31", "33", "44", "1C", "20"]


def hexcolor(el):
    """从一个包含 solidFill 的属性节点里取 srgbClr 值（大写十六进制）。"""
    if el is None:
        return None
    for sf in findall(el, "solidFill"):
        c = find(sf, "srgbClr")
        if c is not None and c.get("val"):
            return c.get("val").upper()
    return None


def is_dark_blue(hexval):
    if not hexval or len(hexval) != 6:
        return False
    r = int(hexval[0:2], 16); g = int(hexval[2:4], 16); b = int(hexval[4:6], 16)
    # 深蓝：蓝分量偏高、整体偏暗、红分量较低
    return b >= 80 and b > r + 20 and r < 140 and g < 160


def is_dark_color(hexval):
    if not hexval or len(hexval) != 6:
        return False
    r = int(hexval[0:2], 16); g = int(hexval[2:4], 16); b = int(hexval[4:6], 16)
    return (r + g + b) / 3 < 128


def get_run_props(container):
    """
    从标题 / 坐标轴等容器中提取首个可用的字符属性。
    返回 dict: {size(pt or None), bold(bool), color(hex or None), font(str or None)}
    优先 a:r/a:rPr，其次 a:defRPr。
    """
    props = {"size": None, "bold": False, "color": None, "font": None, "explicit": False}
    if container is None:
        return props
    rpr = find(container, "rPr")
    if rpr is None:
        rpr = find(container, "defRPr")
    if rpr is None:
        return props
    props["explicit"] = len(list(rpr)) > 0 or bool(rpr.attrib)
    if rpr.get("sz"):
        try:
            props["size"] = int(rpr.get("sz")) / 100.0
        except ValueError:
            pass
    if rpr.get("b") in ("1", "true"):
        props["bold"] = True
    props["color"] = hexcolor(rpr)
    latin = find(rpr, "latin")
    if latin is not None:
        props["font"] = latin.get("typeface")
    return props


# ----------------------------------------------------------------------------
# 载入工作簿：解析出评估所需的全部结构
# ----------------------------------------------------------------------------

class Workbook:
    def __init__(self, path):
        self.path = path
        self.ok_open = False
        self.parts = {}          # 内部路径 -> bytes
        self.sheets = []         # [(name, target_path)]
        self.chart_xml = None    # 解析后的图表 chartSpace 元素
        self.chart_raw = ""      # 图表原始字符串
        self.chart_sheet = None  # 图表所在工作表名
        self.data_grid = {}      # sheetname -> {(row,col): value}
        self.theme_major_font = None   # 主题主要(标题)拉丁字体
        self.theme_minor_font = None   # 主题次要(正文)拉丁字体
        self._load()

    def _load(self):
        if not os.path.isfile(self.path):
            return
        try:
            z = zipfile.ZipFile(self.path)
        except zipfile.BadZipFile:
            return
        try:
            names = z.namelist()
            for n in names:
                self.parts[n] = z.read(n)
            # 校验每个 xml 可解析
            for n in names:
                if n.endswith(".xml") or n.endswith(".rels"):
                    ET.fromstring(self.parts[n])
            self.ok_open = True
        except Exception:
            self.ok_open = False
            return
        finally:
            z.close()

        # 工作表名与关系
        self._parse_sheets()
        # 单元格网格
        for name, tgt in self.sheets:
            self.data_grid[name] = self._parse_grid(tgt)
        # 图表
        self._parse_chart()
        # 主题字体
        self._parse_theme_fonts()

    def _parse_theme_fonts(self):
        """解析主题的主要/次要拉丁字体（图表未显式设字体时会套用主题字体）。"""
        theme = self.parts.get("xl/theme/theme1.xml")
        if not theme:
            return
        try:
            root = ET.fromstring(theme)
        except ET.ParseError:
            return
        for kind, attr in (("majorFont", "theme_major_font"), ("minorFont", "theme_minor_font")):
            node = find(root, kind)
            if node is not None:
                latin = direct(node, "latin")
                if latin is not None and latin.get("typeface"):
                    setattr(self, attr, latin.get("typeface"))

    def _rel_map(self, rels_path):
        m = {}
        if rels_path in self.parts:
            root = ET.fromstring(self.parts[rels_path])
            for r in root:
                m[r.get("Id")] = r.get("Target")
        return m

    def _parse_sheets(self):
        wb = self.parts.get("xl/workbook.xml")
        if not wb:
            return
        root = ET.fromstring(wb)
        rels = self._rel_map("xl/_rels/workbook.xml.rels")
        for s in findall(root, "sheet"):
            name = s.get("name")
            rid = None
            for k, v in s.attrib.items():
                if k.endswith("}id") or k == "id":
                    rid = v
            tgt = rels.get(rid, "")
            if tgt.startswith("/"):
                tgt = tgt[1:]
            elif tgt and not tgt.startswith("xl/"):
                tgt = "xl/" + tgt
            self.sheets.append((name, tgt))

    def _parse_grid(self, tgt):
        grid = {}
        if tgt not in self.parts:
            return grid
        root = ET.fromstring(self.parts[tgt])
        for row in findall(root, "row"):
            for c in directall(row, "c") or findall(row, "c"):
                ref = c.get("r")
                if not ref:
                    continue
                col = "".join(ch for ch in ref if ch.isalpha())
                rownum = "".join(ch for ch in ref if ch.isdigit())
                t = c.get("t")
                val = None
                if t == "inlineStr":
                    val = text_of(find(c, "is"))
                else:
                    v = direct(c, "v")
                    f = direct(c, "f")
                    if v is not None and v.text not in (None, ""):
                        val = v.text
                    elif f is not None:
                        val = "=" + (f.text or "")   # 记录公式引用
                if val is not None and val != "":
                    grid[(int(rownum), _col_to_num(col))] = val
        return grid

    def _parse_chart(self):
        # 找到图表 part 与其所在 sheet
        for name, tgt in self.sheets:
            rels_path = os.path.dirname(tgt) + "/_rels/" + os.path.basename(tgt) + ".rels"
            rels_path = rels_path.replace("\\", "/")
            if rels_path not in self.parts:
                continue
            rmap = self._rel_map(rels_path)
            # sheet -> drawing
            for rid, dtgt in rmap.items():
                if "drawing" in dtgt:
                    dpath = _norm(dtgt)
                    if dpath in self.parts and "drawings/drawing" in dpath:
                        draw = self.parts[dpath]
                        drels = self._rel_map(os.path.dirname(dpath) + "/_rels/" + os.path.basename(dpath) + ".rels")
                        for _, ct in drels.items():
                            if "chart" in ct:
                                cpath = _norm(ct)
                                if cpath in self.parts:
                                    self.chart_raw = self.parts[cpath].decode("utf-8", "ignore")
                                    self.chart_xml = ET.fromstring(self.parts[cpath])
                                    self.chart_sheet = name
                                    self.chart_drawing = draw
                                    return


def _col_to_num(col):
    n = 0
    for ch in col:
        n = n * 26 + (ord(ch.upper()) - ord("A") + 1)
    return n


def _norm(p):
    if p.startswith("/"):
        return p[1:]
    if p.startswith("../"):
        return "xl/" + p[3:]
    if not p.startswith("xl/"):
        return "xl/" + p
    return p


# ----------------------------------------------------------------------------
# 维度1：可用与可修改性（准入门槛）
# ----------------------------------------------------------------------------

def check_dimension1(wb):
    """返回 (passed:bool, checks:list[(desc, ok, note)])。"""
    checks = []

    # 1.1 格式为 .xlsx/.xlsm 且可正常打开
    ext_ok = wb.path.lower().endswith((".xlsx", ".xlsm"))
    open_ok = ext_ok and wb.ok_open
    checks.append(("交付文件为 .xlsx/.xlsm 且可正常打开（zip 结构完整、各 XML 可解析）",
                   open_ok, "扩展名=%s, 打开=%s" % (os.path.splitext(wb.path)[1], wb.ok_open)))
    if not open_ok:
        return False, checks

    passed = all(ok for _, ok, _ in checks)
    return passed, checks


def _is_number(v):
    if isinstance(v, (int, float)):
        return True
    if isinstance(v, str):
        s = v.strip()
        if s.startswith("="):
            return False
        try:
            float(s)
            return True
        except ValueError:
            return False
    return False


# ----------------------------------------------------------------------------
# 图表事实提取：把图表 XML 归纳成便于逐条评估的结构
# ----------------------------------------------------------------------------

class ChartFacts:
    def __init__(self, wb):
        self.wb = wb
        cs = wb.chart_xml
        self.title_text = ""
        self.title_props = {}
        self.chart_type = None       # bar3DChart / barChart / lineChart ...
        self.bar_dir = None
        self.grouping = None
        self.series = []             # [{name, name_ref, cat_ref, val_ref, color}]
        self.cat_ax = None
        self.val_ax = None
        self.ser_ax = None
        self.axis_titles = {}        # pos -> (text, props)
        self.legend_pos = None
        self.has_val_gridlines = False
        self.view3d = False
        self.has_dlbls = False
        self._extract(cs)

    def _extract(self, cs):
        if cs is None:
            return
        chart = direct(cs, "chart")
        # 主标题
        title = direct(chart, "title")
        if title is not None:
            self.title_text = text_of(title)
            self.title_props = get_run_props(find(title, "rich"))
        # 3D
        self.view3d = direct(chart, "view3D") is not None
        plot = find(chart, "plotArea")
        # 图表类型
        for tp in ["bar3DChart", "barChart", "line3DChart", "lineChart",
                   "pie3DChart", "pieChart", "area3DChart"]:
            node = find(plot, tp)
            if node is not None:
                self.chart_type = tp
                bd = find(node, "barDir")
                self.bar_dir = bd.get("val") if bd is not None else None
                gr = find(node, "grouping")
                self.grouping = gr.get("val") if gr is not None else None
                self.has_dlbls = find(node, "dLbls") is not None and \
                    (find(find(node, "dLbls"), "showVal") is not None)
                # 系列
                for ser in findall(node, "ser"):
                    txref = find(find(ser, "tx"), "f")
                    catref = find(find(ser, "cat"), "f")
                    valref = find(find(ser, "val"), "f")
                    self.series.append({
                        "name_ref": txref.text if txref is not None else None,
                        "cat_ref": catref.text if catref is not None else None,
                        "val_ref": valref.text if valref is not None else None,
                        "color": hexcolor(direct(ser, "spPr")),
                        "cat_pts": self._ref_count(catref.text if catref is not None else None),
                        "el": ser,   # 保留原 ser 元素以便解析系列级 dLbls/dLbl
                    })
                break
        # 坐标轴
        self.cat_ax = find(plot, "catAx")
        self.val_ax = find(plot, "valAx")
        self.ser_ax = find(plot, "serAx")
        for ax, key in [(self.cat_ax, "cat"), (self.val_ax, "val"), (self.ser_ax, "ser")]:
            if ax is not None:
                t = direct(ax, "title")
                if t is not None:
                    self.axis_titles[key] = (text_of(t), get_run_props(find(t, "rich")))
        # 数值轴网格线
        if self.val_ax is not None:
            self.has_val_gridlines = direct(self.val_ax, "majorGridlines") is not None
        # 图例
        legend = find(chart, "legend")
        if legend is not None:
            lp = find(legend, "legendPos")
            self.legend_pos = lp.get("val") if lp is not None else "r"

    def _ref_count(self, ref):
        """粗略计算引用区域覆盖的单元格数量（用于数据点数量校验）。"""
        if not ref or ":" not in ref:
            return None
        try:
            rng = ref.split("!")[-1].replace("$", "")
            a, b = rng.split(":")
            r1 = int("".join(c for c in a if c.isdigit()))
            r2 = int("".join(c for c in b if c.isdigit()))
            c1 = _col_to_num("".join(c for c in a if c.isalpha()))
            c2 = _col_to_num("".join(c for c in b if c.isalpha()))
            return (abs(r2 - r1) + 1) * (abs(c2 - c1) + 1)
        except Exception:
            return None

    def resolve_ref_values(self, ref):
        """把 'Sheet2'!$A$3:$A$6 解析为实际单元格值列表（顺着公式链回溯到数值）。"""
        if not ref:
            return []
        parts = ref.split("!")
        if len(parts) != 2:
            return []
        sheet = parts[0].replace("'", "").strip()
        rng = parts[1].replace("$", "")
        cells = []
        if ":" in rng:
            a, b = rng.split(":")
        else:
            a = b = rng
        r1 = int("".join(c for c in a if c.isdigit()))
        r2 = int("".join(c for c in b if c.isdigit()))
        c1 = _col_to_num("".join(c for c in a if c.isalpha()))
        c2 = _col_to_num("".join(c for c in b if c.isalpha()))
        grid = self.wb.data_grid.get(sheet, {})
        for r in range(min(r1, r2), max(r1, r2) + 1):
            for c in range(min(c1, c2), max(c1, c2) + 1):
                cells.append(self._deref(grid.get((r, c))))
        return cells

    def _deref(self, val):
        """若单元格是公式（=Sheet1!B3），回溯取被引用单元格的实际值。"""
        if isinstance(val, str) and val.startswith("="):
            expr = val[1:]
            if "!" in expr and ":" not in expr:
                sh, cell = expr.split("!")
                sh = sh.replace("'", "").strip()
                col = _col_to_num("".join(c for c in cell if c.isalpha()))
                row = int("".join(c for c in cell if c.isdigit()))
                target = self.wb.data_grid.get(sh, {}).get((row, col))
                return self._deref(target)
            return None  # 复杂公式，无缓存值
        return val


# ----------------------------------------------------------------------------
# 维度2：完成度评分细则
#   每条“加分细则” -> 需全部子点为真才计分；“扣分细则” -> 命中任一子点即扣分。
# ----------------------------------------------------------------------------

GOOD_FONTS = ["微软雅黑", "黑体", "等线", "宋体", "Calibri"]
ALLOWED_YEARS = ["2020", "2021", "2022"]
ALLOWED_REGIONS = ["华北", "华南", "中西部", "西部"]


def num(v):
    try:
        return float(v)
    except (TypeError, ValueError):
        return None


def get_source_sales(wb):
    """从 Sheet1 原始表提取 {year: {region: value}}。"""
    g = wb.data_grid.get("Sheet1", {})
    # Sheet1: 行3-5 年份 2020/2021/2022；列B-E 华南/华北/中西部/西部
    regions = {2: g.get((2, 2)), 3: g.get((2, 3)), 4: g.get((2, 4)), 5: g.get((2, 5))}
    data = {}
    for r in (3, 4, 5):
        year = g.get((r, 1))
        row = {}
        for c in (2, 3, 4, 5):
            row[regions.get(c)] = num(g.get((r, c)))
        if year is not None:
            data[str(int(num(year)))] = row
    return data, [regions[c] for c in (2, 3, 4, 5)]


def eval_dimension2(wb):
    f = ChartFacts(wb)
    src, src_regions = get_source_sales(wb)
    src_max = max((v for row in src.values() for v in row.values() if v is not None), default=0)
    rules = []

    def add_rule(points, name, subpoints):
        awarded = all(ok for _, ok, _ in subpoints)
        rules.append({"kind": "add", "points": points, "name": name,
                      "subpoints": subpoints, "hit": awarded})

    # ---- 图表放置位置（锚点：drawing 中 <from> = 图表左上角所在单元格，0 基）----
    anchor_row = anchor_col = None
    if getattr(wb, "chart_drawing", None) is not None:
        dr = ET.fromstring(wb.chart_drawing)
        frm = find(dr, "from")
        if frm is not None:
            rr = find(frm, "row"); cc = find(frm, "col")
            anchor_row = int(rr.text) if rr is not None else None
            anchor_col = int(cc.text) if cc is not None else None

    # 图表所在工作表
    chart_sheet = f.wb.chart_sheet
    chart_grid = wb.data_grid.get(chart_sheet, {}) if chart_sheet else {}

    # 细则要求：图表必须位于名为 "Sheet1" 的工作表内。
    on_sheet1 = chart_sheet == "Sheet1"

    # 定位该表内“区域产品销售额数据区域”的边界（仅统计销售数据表格本身，
    # 排除无关说明文字等孤立单元格；办公软件中数据表由 区域名/年份/表头/销售额 组成）。
    data_cells = []
    has_region_label = False   # 是否含销售区域名称（华北/华南/中西部/西部）
    has_sales_value = False     # 是否含销售额数值或其链接公式
    for (r, c), v in chart_grid.items():
        sv = str(v).strip()
        is_region = sv in ALLOWED_REGIONS
        is_year = sv in ALLOWED_YEARS
        is_header = sv in ("销售区域", "年份")
        is_value = _is_number(v) or (isinstance(v, str) and v.startswith("="))
        if is_region:
            has_region_label = True
        if is_value:
            has_sales_value = True
        if is_region or is_year or is_header or is_value:
            data_cells.append((r, c))
    data_max_row = max((r for (r, c) in data_cells), default=0)
    data_max_col = max((c for (r, c) in data_cells), default=0)
    # 子点①：图表位于 Sheet1 工作表内，且该表包含区域产品销售额数据
    contains_sales_data = on_sheet1 and has_region_label and has_sales_value

    is_3d_bar = f.chart_type == "bar3DChart"

    # 图表是否为“Excel 可编辑三维柱形图对象、非图片”：
    #   三维柱形图 = bar3DChart 且 barDir=col（柱形/纵向），且为原生图表对象(graphicFrame)非图片(pic)
    not_picture = wb.chart_xml is not None
    if getattr(wb, "chart_drawing", None) is not None:
        draw_root = ET.fromstring(wb.chart_drawing)
        has_frame = len(findall(draw_root, "graphicFrame")) > 0
        has_pic = len(findall(draw_root, "pic")) > 0
        not_picture = has_frame and not (has_pic and not has_frame)
    is_3d_col_chart = is_3d_bar and f.bar_dir == "col" and not_picture

    # +1 数据工作表变化趋势图对象（严格对齐细则三点）
    #   ① 位于包含区域产品销售额数据的工作表内
    #   ② 图表左上角位于数据区域右侧或下方空白区域，距数据区域边界至少1列或2行
    #   ③ 图表为 Excel 可编辑三维柱形图对象，不是图片对象
    below_or_right = False
    gap_note = ""
    if anchor_row is not None and anchor_col is not None and data_cells:
        anchor_r1 = anchor_row + 1   # 转 Excel 行号（1 基）
        anchor_c1 = anchor_col + 1   # 转 Excel 列号（1 基）
        right = (anchor_c1 - data_max_col) >= 1   # 位于数据区右侧且距边界≥1列
        below = (anchor_r1 - data_max_row) >= 2   # 位于数据区下方且距边界≥2行
        below_or_right = right or below
        gap_note = "锚点(行%d,列%d) 数据区右下界(行%d,列%d) 右侧距≥1列=%s 下方距≥2行=%s" % (
            anchor_r1, anchor_c1, data_max_row, data_max_col, right, below)
    add_rule(1, "数据工作表变化趋势图对象", [
        ("位于包含区域产品销售额数据的sheet1工作表内", contains_sales_data,
         "图表所在表=%s(须为Sheet1) 含区域标签=%s 含销售额数据=%s" % (chart_sheet, has_region_label, has_sales_value)),
        ("图表左上角位于数据区域右侧或下方空白区域，距数据区域边界至少1列或2行", below_or_right, gap_note),
        ("图表为 Excel 可编辑三维柱形图对象，不是图片对象", is_3d_col_chart,
         "图表类型=%s barDir=%s 非图片=%s" % (f.chart_type, f.bar_dir, not_picture)),
    ])

    # +1 变化趋势图数据源（严格对齐细则每一点）
    #   ① 数据来源于工作簿区域产品销售额表格的单元格区域
    #   ② 横向分类使用销售区域数据
    #   ③ 系列使用年份数据
    #   ④ 数值使用对应销售额数据
    #   ⑤ 系列数量与 sheet1 年份列数量一致
    #   ⑥ 数据点数量与 sheet1 销售区域数量一致
    ser_count = len(f.series)
    cat_pts = f.series[0]["cat_pts"] if f.series else None
    src_years = list(src.keys())

    # 图表实际引用解析：横向分类值、各系列名（年份）
    cat_vals = [str(v).strip() for v in f.resolve_ref_values(f.series[0]["cat_ref"])
                if v not in (None, "")] if f.series else []
    ser_year_names = []
    for s in f.series:
        nm = _single_ref_value(f, s["name_ref"])
        ser_year_names.append(str(nm).strip() if nm is not None else "")

    # ① 各系列的分类(cat)、系列名(tx)、数值(val)均为单元格区域引用（非手填常量/图片）
    refs_ok = bool(f.series) and all(
        s["cat_ref"] and s["val_ref"] and s["name_ref"] for s in f.series)

    # ② 横向分类使用销售区域数据：分类标签落在销售区域集合内，且覆盖源表所有区域
    cat_is_regions = bool(cat_vals) and all(cv in ALLOWED_REGIONS for cv in cat_vals) and \
        all(any(cv == r for cv in cat_vals) for r in src_regions if r)

    # ③ 系列使用年份数据：系列名均为年份，且覆盖源表所有年份
    ser_is_years = bool(ser_year_names) and all(sn in ALLOWED_YEARS for sn in ser_year_names) and \
        all(any(sn == y for sn in ser_year_names) for y in src_years)

    # ④ 数值使用对应销售额数据：逐系列(年份)逐分类(区域)与 sheet1 销售额比对一致
    values_match = bool(f.series)
    for i, s in enumerate(f.series):
        year = ser_year_names[i] if i < len(ser_year_names) else None
        vals = [num(v) for v in f.resolve_ref_values(s["val_ref"])]
        expect = [src.get(year, {}).get(cv) for cv in cat_vals]
        if not (vals and expect and all(v is not None for v in vals) and vals == expect):
            values_match = False

    add_rule(1, "变化趋势图数据源", [
        ("图表数据来源于工作簿区域产品销售额表格的单元格区域", refs_ok,
         "各系列均含 分类/系列名/数值 单元格引用=%s" % refs_ok),
        ("横向分类使用销售区域数据", cat_is_regions, "横向分类值=%s" % cat_vals),
        ("系列使用年份数据", ser_is_years, "系列名(年份)=%s" % ser_year_names),
        ("数值使用对应销售额数据", values_match, "各系列数值与 sheet1 对应销售额逐一比对"),
        ("图表数据系列数量与sheet1表格年份列数量一致（=%d）" % len(src_years),
         ser_count == len(src_years), "系列数=%d 年份数=%d" % (ser_count, len(src_years))),
        ("数据点数量与sheet1销售区域数量一致（=%d）" % len(src_regions),
         cat_pts == len(src_regions), "每系列数据点=%s 销售区域数=%d" % (cat_pts, len(src_regions))),
    ])

    # +1 变化趋势图类型（严格对齐细则每一点）
    #   ① 图表类型为三维簇状柱形图或三维柱形图
    #   ② 柱形具有立体厚度和透视角度
    #   ③ 不同年份的柱形按销售区域分组排列，同一区域内多个年份柱形并列显示
    #   ④ 不能做成普通二维柱形图、折线图、饼图或截图
    #
    #   办公软件依据（OOXML）：
    #     三维柱形图 = <c:bar3DChart>，纵向柱形 = <c:barDir val="col"/>
    #     簇状（并列）= <c:grouping val="clustered"/>（三维柱形亦可为 standard，
    #       standard 时各年份沿深度轴排布，仍属三维柱形图但非簇状；
    #       “同一区域内多个年份柱形并列显示”对应 clustered）
    #     立体厚度+透视角度 = <c:view3D>，其 rotX/rotY 提供透视旋转角度
    is_3d_col = is_3d_bar and f.bar_dir == "col"
    is_clustered = f.grouping == "clustered"

    # 透视角度：view3D 存在且含旋转角度 rotX/rotY
    has_perspective = False
    if f.wb.chart_xml is not None:
        chart_el = direct(f.wb.chart_xml, "chart")
        v3d = direct(chart_el, "view3D") if chart_el is not None else None
        if v3d is not None:
            rotx = direct(v3d, "rotX")
            roty = direct(v3d, "rotY")
            has_perspective = (rotx is not None) or (roty is not None)
    # 立体厚度由三维柱形图本身提供（bar3DChart 渲染带深度的立方体）
    has_depth = is_3d_bar

    # 三维图三轴位置校验：真正的三维柱形图应有 X(分类,底部b)、Y(数值,左l)、
    #   Z(系列/深度轴,右r 或 后)三根位置合理的轴。若三根轴的 axPos 全部相同
    #   （如本文件全为 l），说明轴位置定义损坏，不是规范可用的三维坐标系。
    def _axpos(ax):
        el = direct(ax, "axPos") if ax is not None else None
        return el.get("val") if el is not None else None
    cat_pos = _axpos(f.cat_ax)
    val_pos = _axpos(f.val_ax)
    ser_pos = _axpos(f.ser_ax)
    has_three_axes = (f.cat_ax is not None) and (f.val_ax is not None) and (f.ser_ax is not None)
    axis_positions = [cat_pos, val_pos, ser_pos]
    # 三轴位置合理：三根轴都存在，且三者位置不完全相同（存在区分的 X/Y/Z 布局）
    axes_pos_ok = has_three_axes and len(set(p for p in axis_positions if p)) >= 2
    is_valid_3d = is_3d_col and axes_pos_ok

    not_2d_line_pie_pic = is_3d_bar and not_picture

    add_rule(1, "变化趋势图类型", [
        ("图表类型为三维簇状柱形图或三维柱形图", is_valid_3d,
         "类型=%s barDir=%s grouping=%s 三轴位置(cat/val/ser)=%s 位置合理=%s" % (
             f.chart_type, f.bar_dir, f.grouping, axis_positions, axes_pos_ok)),
        ("柱形具有立体厚度和透视角度", has_depth and has_perspective,
         "三维柱形立体厚度=%s view3D透视角度=%s" % (has_depth, has_perspective)),
        ("不同年份柱形按销售区域分组排列，同一区域内多个年份柱形并列显示", is_3d_col and is_clustered,
         "簇状并列(grouping=clustered)=%s" % is_clustered),
        ("非普通二维柱形图/折线图/饼图/截图", not_2d_line_pie_pic,
         "类型=%s 非图片=%s" % (f.chart_type, not_picture)),
    ])

    # +1 变化趋势图白色图表区（严格对齐细则每一点）
    #   ① 图表区背景为白色或无填充
    #   ② 图表外框为无边框或浅灰色细边框
    #   ③ 图表区没有深色底纹、图片背景、水印或无关大色块
    #
    #   办公软件依据（OOXML）：图表区样式在 <c:chartSpace>/<c:spPr> 上：
    #     背景填充 = spPr/a:solidFill 或 a:noFill；外框 = spPr/a:ln（含其填充/线宽）
    #     图片背景/水印 = a:blipFill；渐变大色块 = a:gradFill/a:pattFill
    space_spPr = direct(wb.chart_xml, "spPr")

    # ① 背景：无 spPr(默认白) / noFill / 白色 solidFill 均视为满足；深色 solidFill 不满足
    space_fill = hexcolor(space_spPr)
    has_no_fill = space_spPr is not None and direct(space_spPr, "noFill") is not None
    bg_white_or_none = (space_spPr is None) or has_no_fill or (space_fill is None) or \
        (not is_dark_color(space_fill))
    bg_dark = space_fill is not None and is_dark_color(space_fill)

    # ② 外框：无 a:ln / a:ln 内 noFill(无边框) / 边框为浅灰色 均满足；深色边框不满足
    ln = direct(space_spPr, "ln") if space_spPr is not None else None
    border_ok = True
    if ln is not None:
        if direct(ln, "noFill") is not None:
            border_ok = True                       # 无边框
        else:
            ln_color = hexcolor(ln)
            if ln_color is None:
                border_ok = True                   # 未指定颜色，采用默认浅色
            else:
                border_ok = not is_dark_color(ln_color)   # 浅灰细边框可，深色不可

    # ③ 图片背景/水印/大色块：blipFill / gradFill / pattFill 视为不满足
    has_img_bg = False
    has_block_fill = False
    if space_spPr is not None:
        has_img_bg = find(space_spPr, "blipFill") is not None
        has_block_fill = (find(space_spPr, "gradFill") is not None) or \
            (find(space_spPr, "pattFill") is not None)
    # 兜底：整个图表 XML 若含图片填充引用（水印/背景图）
    if "blipFill" in wb.chart_raw:
        has_img_bg = True
    no_bad_decor = (not bg_dark) and (not has_img_bg) and (not has_block_fill)

    add_rule(1, "变化趋势图白色图表区", [
        ("图表区背景为白色或无填充", bg_white_or_none,
         "图表区填充=%s" % (space_fill or ("无填充" if has_no_fill else "无(默认白)"))),
        ("图表外框为无边框或浅灰色细边框", border_ok,
         "外框=%s" % ("无(默认)" if ln is None else (hexcolor(ln) or "无填充/浅色"))),
        ("图表区无深色底纹/图片背景/水印/无关大色块", no_bad_decor,
         "深色底纹=%s 图片背景=%s 大色块填充=%s" % (bg_dark, has_img_bg, has_block_fill)),
    ])

    # +1 变化趋势图绘图区（严格对齐细则每一点）
    #   ① 绘图区位于图表中部（居中）
    #   ② 宽度占图表宽度 75%–90%
    #   ③ 高度占图表高度 65%–82%
    #   ④ 绘图区背景为白色或极浅灰色
    #   ⑤ 内部出现浅灰色三维网格线和后侧透视网格面
    #
    #   办公软件依据（OOXML）：
    #     绘图区布局 = plotArea/c:layout/c:manualLayout（c:x/c:y/c:w/c:h 为占图表比例 0–1）
    #       未设 manualLayout 时 Excel 自动布局：绘图区默认居中、宽高落在常见比例范围内
    #     绘图区背景 = plotArea/c:spPr（a:solidFill / a:noFill）
    #     三维主网格线 = valAx/c:majorGridlines；后侧透视网格面 = chart/c:backWall
    chart_el2 = direct(wb.chart_xml, "chart")
    plot = find(chart_el2, "plotArea")
    layout = direct(plot, "layout") if plot is not None else None
    manual = direct(layout, "manualLayout") if layout is not None else None

    def _lay_val(tag):
        el = direct(manual, tag) if manual is not None else None
        try:
            return float(el.get("val")) if el is not None and el.get("val") else \
                (float(el.text) if el is not None and el.text else None)
        except (TypeError, ValueError):
            return None

    lw = _lay_val("w"); lh = _lay_val("h")
    lx = _lay_val("x"); ly = _lay_val("y")

    if manual is None:
        # 自动布局：Excel 默认绘图区居中、宽高在要求范围内
        centered = True
        width_ok = True
        height_ok = True
        lay_note = "manualLayout 未设置，采用 Excel 默认自动布局（居中、宽高落在要求范围内）"
    else:
        width_ok = lw is not None and 0.75 <= lw <= 0.90
        height_ok = lh is not None and 0.65 <= lh <= 0.82
        # 居中：左边距与右侧留白大致对称（|x - (1-x-w)| 较小）
        if lx is not None and lw is not None:
            right_gap = 1 - lx - lw
            centered = abs(lx - right_gap) <= 0.12
        else:
            centered = True
        lay_note = "manualLayout x=%s y=%s w=%s h=%s" % (lx, ly, lw, lh)

    # ④ 绘图区背景：白色或极浅灰色（无 spPr/noFill=默认白 视为满足；深色不满足）
    plot_spPr = direct(plot, "spPr") if plot is not None else None
    plot_fill = hexcolor(plot_spPr)
    plot_nofill = plot_spPr is not None and direct(plot_spPr, "noFill") is not None
    plot_bg_ok = (plot_spPr is None) or plot_nofill or (plot_fill is None) or \
        (not is_dark_color(plot_fill))

    # ⑤ 三维主网格线 + 后侧透视网格面（背景墙）
    #    "三维网格线/后侧透视网格面"须建立在规范可用的三维坐标系上：
    #    若三轴位置损坏（is_valid_3d=False），则不认定为有效的三维网格/透视面。
    has_gridlines = f.has_val_gridlines
    back_wall = direct(chart_el2, "backWall") is not None
    valid_3d_grid = has_gridlines and back_wall and is_valid_3d

    add_rule(1, "变化趋势图绘图区", [
        ("绘图区位于图表中部（居中）", centered, lay_note),
        ("宽度占图表宽度75%–90%", width_ok,
         "绘图区宽度比例=%s" % (lw if lw is not None else "自动(默认满足)")),
        ("高度占图表高度65%–82%", height_ok,
         "绘图区高度比例=%s" % (lh if lh is not None else "自动(默认满足)")),
        ("绘图区背景为白色或极浅灰色", plot_bg_ok,
         "绘图区填充=%s" % (plot_fill or ("无填充" if plot_nofill else "无(默认白)"))),
        ("内部出现浅灰色三维网格线和后侧透视网格面", valid_3d_grid,
         "valAx主网格线=%s 后侧背景墙(backWall)=%s 三维坐标系有效=%s" % (
             has_gridlines, back_wall, is_valid_3d)),
    ])

    # +1 变化趋势图顶部标题文本（严格对齐细则每一点）
    #   ① 位于图表顶部居中位置
    #   ② 文本为"区域出货量变化趋势图"
    #   ③ 字体为微软雅黑、黑体或等线
    #   ④ 字号 16–20 磅
    #   ⑤ 加粗
    #   ⑥ 颜色为深蓝色
    #   ⑦ 标题不与绘图区重叠
    #
    #   办公软件依据（OOXML）：主标题 = chart/c:title
    #     位置 = title/c:layout（无 manualLayout 时 Excel 默认顶部居中）
    #     是否叠加在绘图区上 = title/c:overlay（val=0 或缺省 => 不与绘图区重叠）
    #     文本样式 = title/c:tx/c:rich 内 a:rPr/a:defRPr（sz/b/latin/solidFill）
    tp = f.title_props
    title_el = direct(chart_el2, "title") if chart_el2 is not None else None

    # ① 顶部居中：标题存在且未用 manualLayout 强制移位（默认顶部居中）
    title_layout = direct(title_el, "layout") if title_el is not None else None
    title_manual = direct(title_layout, "manualLayout") if title_layout is not None else None
    top_center = title_el is not None and title_manual is None

    # ⑦ 不与绘图区重叠：overlay 缺省或 val=0 视为不重叠；val=1 视为重叠
    overlay_el = direct(title_el, "overlay") if title_el is not None else None
    overlay_on = overlay_el is not None and overlay_el.get("val") in ("1", "true")
    not_overlap = not overlay_on

    add_rule(1, "变化趋势图顶部标题文本", [
        ("位于图表顶部居中位置", top_center,
         "标题存在=%s 顶部居中(无手动移位)=%s" % (title_el is not None, title_manual is None)),
        ("文本为“区域出货量变化趋势图”", f.title_text.strip() == "区域出货量变化趋势图",
         "标题=%r" % f.title_text),
        ("字体为微软雅黑/黑体/等线", tp.get("font") in ("微软雅黑", "黑体", "等线"),
         "字体=%s" % tp.get("font")),
        ("字号16磅–20磅", tp.get("size") is not None and 16 <= tp["size"] <= 20,
         "字号=%s" % tp.get("size")),
        ("加粗", tp.get("bold", False), "bold=%s" % tp.get("bold")),
        ("颜色为深蓝色", is_dark_blue(tp.get("color")), "颜色=%s" % tp.get("color")),
        ("标题不与绘图区重叠", not_overlap, "overlay=%s" % (overlay_el.get("val") if overlay_el is not None else "缺省(不重叠)")),
    ])

    # +1 变化趋势图左侧纵轴标题文本（严格对齐细则每一点）
    #   ① 位于图表左侧中部
    #   ② 文本为"出货量（万元）"
    #   ③ 文字竖向排列或旋转90度
    #   ④ 字体为微软雅黑、黑体或 Calibri
    #   ⑤ 字号 8磅–12磅
    #   ⑥ 颜色为深蓝色或黑色
    #
    #   办公软件依据（OOXML）：数值轴标题 = valAx/c:title
    #     左侧位置 = valAx/c:axPos val="l"
    #     竖排/旋转 = title/c:tx/c:rich/a:bodyPr 的 rot（60000=1度，-5400000=-90度）
    #       或 vert 属性（"vert"/"vert270"/"wordArtVert" 等竖排）；
    #       Excel 数值轴标题默认即旋转 -90 度竖排（bodyPr 缺省时按默认竖排处理）
    #     文本样式 = title/c:tx/c:rich 内 a:rPr/a:defRPr
    vt_text, vt_props = f.axis_titles.get("val", ("", {}))
    val_title_el = direct(f.val_ax, "title") if f.val_ax is not None else None

    # ① 左侧：valAx 的 axPos = "l"
    val_axpos_el = direct(f.val_ax, "axPos") if f.val_ax is not None else None
    val_left = val_axpos_el is not None and val_axpos_el.get("val") == "l"

    # ③ 竖排或旋转90度：读 title 内 a:bodyPr 的 rot / vert；缺省按 Excel 默认竖排
    vt_body = find(val_title_el, "bodyPr") if val_title_el is not None else None
    vert_or_rot = True
    rot_note = "bodyPr 缺省，采用 Excel 纵轴标题默认竖排(-90°)"
    if vt_body is not None:
        rot = vt_body.get("rot")
        vert = vt_body.get("vert")
        if rot is not None:
            try:
                deg = int(rot) / 60000.0
                vert_or_rot = abs(abs(deg) - 90) <= 5   # 接近 ±90 度
                rot_note = "旋转角度=%.0f°" % deg
            except ValueError:
                vert_or_rot = True
        elif vert in ("vert", "vert270", "wordArtVert", "wordArtVertRtl", "eaVert", "mongolianVert"):
            vert_or_rot = True
            rot_note = "竖向排列 vert=%s" % vert
        else:
            # bodyPr 存在但既未旋转也未竖排（横排）=> 不满足
            vert_or_rot = False
            rot_note = "横向排列(rot/vert 均未设竖排)"

    add_rule(1, "变化趋势图左侧纵轴标题文本", [
        ("位于图表左侧中部", val_left,
         "valAx axPos=%s" % (val_axpos_el.get("val") if val_axpos_el is not None else "无")),
        ("文本为“出货量（万元）”", vt_text.strip() in ("出货量（万元）", "出货量(万元)"),
         "纵轴标题=%r" % vt_text),
        ("文字竖向排列或旋转90度", vert_or_rot, rot_note),
        ("字体为微软雅黑/黑体/Calibri", vt_props.get("font") in ("微软雅黑", "黑体", "Calibri"),
         "字体=%s" % vt_props.get("font")),
        ("字号8磅–12磅", _size_ok(vt_props, 8, 12),
         "字号=%s" % vt_props.get("size")),
        ("颜色为深蓝色或黑色", _color_dark_or_default(vt_props),
         "颜色=%s" % vt_props.get("color")),
    ])

    # +1 变化趋势图底部横轴标题文本（严格对齐细则每一点）
    #   ① 位于图表底部居中位置
    #   ② 文本为"销售区域"
    #   ③ 字体为微软雅黑、黑体或 Calibri
    #   ④ 字号 8磅–12磅
    #   ⑤ 加粗或半加粗
    #   ⑥ 颜色为深蓝色或黑色
    #
    #   办公软件依据（OOXML）：分类轴标题 = catAx/c:title
    #     底部位置 = catAx/c:axPos val="b"
    #     文本样式 = title/c:tx/c:rich 内 a:rPr/a:defRPr（sz/b/latin/solidFill）
    ct_text, ct_props = f.axis_titles.get("cat", ("", {}))

    # ① 位于图表底部居中位置：
    #    柱形图(barDir=col)的分类轴(catAx)即底部水平轴，办公软件默认将其标题
    #    渲染在图表底部居中。因此只要该横轴标题存在于分类轴上即视为"底部居中"。
    cat_title_el = direct(f.cat_ax, "title") if f.cat_ax is not None else None
    cat_bottom = cat_title_el is not None and f.bar_dir == "col"

    add_rule(1, "变化趋势图底部横轴标题文本", [
        ("位于图表底部居中位置", cat_bottom,
         "分类轴(底部水平轴)标题存在=%s 列图表=%s" % (cat_title_el is not None, f.bar_dir == "col")),
        ("文本为“销售区域”", ct_text.strip() == "销售区域", "横轴标题=%r" % ct_text),
        ("字体为微软雅黑/黑体/Calibri", _font_ok(ct_props, wb, ("微软雅黑", "黑体", "Calibri")),
         _font_note(ct_props, wb)),
        ("字号8磅–12磅", _size_ok(ct_props, 8, 12), "字号=%s" % ct_props.get("size")),
        ("加粗或半加粗", _bold_ok(ct_props, wb), _bold_note(ct_props, wb)),
        ("颜色为深蓝色或黑色", _color_dark_or_default(ct_props), "颜色=%s" % ct_props.get("color")),
    ])

    # +1 变化趋势图右侧深度轴标题文本（严格对齐细则每一点）
    #   ① 位于图表右侧中下部
    #   ② 文本为"年份"
    #   ③ 字体为微软雅黑、黑体或 Calibri
    #   ④ 字号 8磅–12磅
    #   ⑤ 加粗或半加粗
    #   ⑥ 颜色为深蓝色或黑色
    #
    #   办公软件依据（OOXML）：深度轴（系列轴）标题 = serAx/c:title
    #     右侧位置 = serAx/c:axPos val="r"（三维图表深度轴默认位于右侧）
    #     文本样式 = title/c:tx/c:rich 内 a:rPr/a:defRPr（sz/b/latin/solidFill）
    st_text, st_props = f.axis_titles.get("ser", ("", {}))

    # ① 右侧：serAx 的 axPos = "r"
    ser_axpos_el = direct(f.ser_ax, "axPos") if f.ser_ax is not None else None
    ser_right = ser_axpos_el is not None and ser_axpos_el.get("val") == "r"

    add_rule(1, "变化趋势图右侧深度轴标题文本", [
        ("位于图表右侧中下部", ser_right,
         "serAx axPos=%s" % (ser_axpos_el.get("val") if ser_axpos_el is not None else "无")),
        ("文本为“年份”", st_text.strip() == "年份", "深度轴标题=%r" % (st_text or "无")),
        ("字体为微软雅黑/黑体/Calibri", st_props.get("font") in ("微软雅黑", "黑体", "Calibri"),
         "字体=%s" % st_props.get("font")),
        ("字号8磅–12磅", _size_ok(st_props, 8, 12), "字号=%s" % st_props.get("size")),
        ("加粗或半加粗", st_props.get("bold", False), "bold=%s" % st_props.get("bold")),
        ("颜色为深蓝色或黑色", _color_dark_or_default(st_props), "颜色=%s" % st_props.get("color")),
    ])

    # +1 变化趋势图销售区域分类标签（严格对齐细则每一点）
    #   ① 横轴分类标签从左到右显示 sheet1 表中的销售区域名称
    #   ② 应包含"华北""华南""中西部""西部"
    #   ③ 标签字体为微软雅黑、宋体或 Calibri
    #   ④ 字号 8磅–12磅
    #   ⑤ 颜色为深蓝色或黑色
    #   ⑥ 标签不相互重叠
    #
    #   办公软件依据（OOXML）：分类轴刻度标签样式 = catAx/c:txPr（a:defRPr）
    #     是否显示/是否重叠 = catAx/c:tickLblPos（"none"=不显示；非none=显示，
    #       Excel 自动排布刻度标签，默认不重叠）
    cat_vals = [str(v).strip() for v in f.resolve_ref_values(f.series[0]["cat_ref"])
                if v not in (None, "")] if f.series else []
    # ② 包含四个销售区域
    regions_included = all(any(r == cv for cv in cat_vals) for r in ALLOWED_REGIONS)
    # ① 从左到右为 sheet1 销售区域名称（顺序即引用区域读出的顺序，且均为合法区域名）
    labels_are_regions = bool(cat_vals) and all(cv in ALLOWED_REGIONS for cv in cat_vals)

    # ③④⑤ 分类轴刻度标签文字样式（catAx/txPr）
    cat_lbl_props = get_run_props(find(f.cat_ax, "txPr")) if f.cat_ax is not None else {}

    # ⑥ 不相互重叠：tickLblPos != "none"（显示标签）时，Excel 自动排布不重叠
    cat_ticklbl = direct(f.cat_ax, "tickLblPos") if f.cat_ax is not None else None
    labels_shown = cat_ticklbl is None or cat_ticklbl.get("val") != "none"

    add_rule(1, "变化趋势图销售区域分类标签", [
        ("横轴分类标签从左到右显示sheet1表中的销售区域名称", labels_are_regions,
         "分类标签(左→右)=%s" % cat_vals),
        ("包含华北/华南/中西部/西部", regions_included, "分类标签=%s" % cat_vals),
        ("标签字体为微软雅黑/宋体/Calibri",
         _font_ok(cat_lbl_props, wb, ("微软雅黑", "宋体", "Calibri")),
         _font_note(cat_lbl_props, wb)),
        ("字号8磅–12磅", _size_ok(cat_lbl_props, 8, 12), "字号=%s" % cat_lbl_props.get("size")),
        ("颜色为深蓝色或黑色", _color_dark_or_default(cat_lbl_props), "颜色=%s" % cat_lbl_props.get("color")),
        ("标签不相互重叠", labels_shown,
         "tickLblPos=%s（非none=显示，自动排布不重叠）" % (cat_ticklbl.get("val") if cat_ticklbl is not None else "缺省(nextTo)")),
    ])

    # +1 变化趋势图年份系列标签（严格对齐细则每一点）
    #   ① 深度轴或图例中显示 sheet1 表年份系列，应包含"2020""2021""2022"
    #   ② 字号 8磅–12磅
    #   ③ 年份顺序从前到后或从左到右与 sheet1 表顺序一致
    #
    #   办公软件依据（OOXML）：年份系列名 = 各 ser/c:tx（引用 Sheet2 年份表头，
    #     其顺序即系列 idx/order 顺序）；显示于图例(c:legend)或深度轴(serAx)。
    #     系列标签文字样式默认取自图例 txPr（未显式设置时按 Excel 默认，约10磅）。
    ser_names = []
    for s in f.series:
        # 系列名多为单格引用（如 'Sheet2'!B2）
        nm = _single_ref_value(f, s["name_ref"])
        ser_names.append(str(nm).strip() if nm is not None else "")

    # ① 包含 2020/2021/2022，且显示于图例或深度轴
    years_included = all(any(y == sn for sn in ser_names) for y in ALLOWED_YEARS)
    shown_in_legend_or_serax = (f.legend_pos is not None) or (f.ser_ax is not None)
    years_ok = years_included and shown_in_legend_or_serax

    # ③ 顺序与 sheet1 年份顺序一致（sheet1 行3→5 = 2020,2021,2022）
    src_year_order = list(src.keys())   # 按 get_source_sales 读入顺序（行3,4,5）
    order_ok = ser_names == src_year_order

    # ② 字号：图例文字样式（未显式设置时 Excel 默认约10磅，落在 8–12）
    legend_el = find(chart_el2, "legend") if chart_el2 is not None else None
    legend_props = get_run_props(find(legend_el, "txPr")) if legend_el is not None else {}

    add_rule(1, "变化趋势图年份系列标签", [
        ("深度轴或图例显示年份系列，包含2020/2021/2022", years_ok,
         "系列名=%s 图例位置=%s 深度轴=%s" % (ser_names, f.legend_pos, f.ser_ax is not None)),
        ("字号8磅–12磅", _size_ok(legend_props, 8, 12), "字号=%s" % legend_props.get("size")),
        ("年份顺序从前到后/从左到右与sheet1表顺序一致", order_ok,
         "系列顺序=%s sheet1年份顺序=%s" % (ser_names, src_year_order)),
    ])

    # +1 变化趋势图纵轴刻度标签（严格对齐细则每一点）
    #   ① 纵轴刻度从0开始
    #   ② 最大刻度覆盖 sheet1 表格中最高销售额数据
    #   ③ 纵轴最大刻度应为10000左右
    #   ④ 刻度标签"0、1000…10000"为黑色或深蓝色
    #   ⑤ 字号 8磅–12磅
    #
    #   办公软件依据（OOXML）：数值轴 = valAx
    #     刻度范围 = valAx/c:scaling/c:min、c:max（未设=Excel 自动：min 自动为0、
    #       max 自动向上取整到覆盖最高值的整刻度）
    #     刻度步长 = valAx/c:majorUnit（决定相邻刻度标签的间隔）。细则明确要求
    #       刻度标签为“0、1000、2000…10000”，即 min=0、max=10000、步长=1000。
    #       Excel 在该量级(0–10000)的自动步长通常为 2000（而非 1000），
    #       故不能仅凭“max 缺省”就假定步长=1000；必须解析 majorUnit（或据
    #       min/max/majorUnit 推算实际刻度标签序列）来确认是否按 1000 间隔显示。
    #     刻度标签文字样式 = valAx/c:txPr（a:defRPr）
    vmin_el = vmax_el = None
    if f.val_ax is not None:
        sc = direct(f.val_ax, "scaling")
        if sc is not None:
            vmin_el = direct(sc, "min")
            vmax_el = direct(sc, "max")

    # ① 从0开始：未设 min（自动=0）或显式 min=0
    try:
        min_val = float(vmin_el.get("val")) if vmin_el is not None and vmin_el.get("val") else None
    except (TypeError, ValueError):
        min_val = None
    starts_zero = (vmin_el is None) or (min_val == 0)

    # ③ 最大刻度≈10000：未设 max（自动上限=10000）或显式 max 在 9500–11000
    try:
        max_val = float(vmax_el.get("val")) if vmax_el is not None and vmax_el.get("val") else None
    except (TypeError, ValueError):
        max_val = None
    auto_or_10000 = (vmax_el is None) or (max_val is not None and 9500 <= max_val <= 11000)
    effective_min = min_val if min_val is not None else 0     # 自动下限=0
    effective_max = max_val if max_val is not None else 10000  # 自动上限=10000

    # ② 覆盖最高销售额：有效最大刻度 >= sheet1 最高销售额
    covers_max = effective_max >= src_max

    # ④ 刻度标签“0、1000、2000…10000”：解析 majorUnit 并据 min/max/majorUnit
    #    推算实际渲染的刻度标签序列，确认从0到10000按1000间隔显示。
    mu_el = direct(f.val_ax, "majorUnit") if f.val_ax is not None else None
    try:
        major_unit = float(mu_el.get("val")) if mu_el is not None and mu_el.get("val") else None
    except (TypeError, ValueError):
        major_unit = None

    def _tick_labels(lo, hi, step):
        """按 [lo, hi] 与步长 step 推算刻度标签序列（含首末刻度）。"""
        if not step or step <= 0:
            return None
        labels = []
        x = lo
        n = 0
        while x <= hi + 1e-6 and n <= 1000:   # 防浮点误差与死循环
            labels.append(int(round(x)))
            x += step
            n += 1
        return labels

    expected_labels = list(range(0, 10001, 1000))   # 0,1000,2000,…,10000
    # 仅当步长被显式设置（majorUnit）时才能确认间隔；未显式设置时不假定=1000
    tick_labels = _tick_labels(effective_min, effective_max, major_unit) if major_unit else None
    labels_are_expected = tick_labels == expected_labels
    if major_unit is None:
        mu_note = "未显式设置 majorUnit，无法确认刻度标签按1000间隔显示（Excel 自动步长在该量级通常为2000，非1000）"
    else:
        mu_note = "majorUnit=%s min=%s max=%s -> 刻度标签=%s（期望=%s）" % (
            int(major_unit), effective_min, effective_max, tick_labels, expected_labels)

    # ⑤ 刻度标签文字样式（颜色/字号）
    val_lbl_props = get_run_props(find(f.val_ax, "txPr")) if f.val_ax is not None else {}

    add_rule(1, "变化趋势图纵轴刻度标签", [
        ("纵轴刻度从0开始", starts_zero,
         "min=%s" % (min_val if vmin_el is not None else "自动(0)")),
        ("最大刻度覆盖sheet1表格中最高销售额数据（最高=%s）" % src_max, covers_max,
         "有效最大刻度=%s >= 最高销售额=%s" % (effective_max, src_max)),
        ("纵轴最大刻度为10000左右", auto_or_10000,
         "max=%s" % (max_val if vmax_el is not None else "自动(上限10000)")),
        ("刻度标签为0、1000、2000…10000（从0到10000按1000间隔显示）", labels_are_expected,
         mu_note),
        ("刻度标签为黑色或深蓝色", _color_dark_or_default(val_lbl_props),
         "颜色=%s（默认黑色）" % val_lbl_props.get("color")),
        ("字号8磅–12磅", _size_ok(val_lbl_props, 8, 12), "字号=%s" % val_lbl_props.get("size")),
    ])

    # +1 变化趋势图纵轴主网格线（严格对齐细则每一点）
    #   ① 绘图区内出现浅灰色水平网格线或三维透视网格线
    #   ② 线条为浅灰色单实线
    #   ③ 线宽 0.5磅–1磅
    #   ④ 网格线不遮挡柱形和数据标签
    #
    #   办公软件依据（OOXML）：数值轴主网格线 = valAx/c:majorGridlines
    #     线样式 = majorGridlines/c:spPr/a:ln：颜色 a:solidFill/a:srgbClr，
    #       线宽 a:ln@w（EMU，1磅=12700EMU），线型 a:prstDash（solid=单实线）
    #     未显式设置时 Excel 默认网格线为浅灰色(约 D9D9D9)、单实线、约0.75磅、
    #       且默认渲染在柱形/数据标签之后（不遮挡）。
    grid_el = direct(f.val_ax, "majorGridlines") if f.val_ax is not None else None
    has_gridline = grid_el is not None
    grid_ln = direct(direct(grid_el, "spPr"), "ln") if grid_el is not None else None

    # ② 浅灰色单实线：颜色浅灰(非深色) + 线型 solid(或默认)
    grid_color = hexcolor(grid_ln) if grid_ln is not None else None
    color_light = grid_color is None or (not is_dark_color(grid_color))
    dash_el = direct(grid_ln, "prstDash") if grid_ln is not None else None
    solid_line = dash_el is None or dash_el.get("val") == "solid"
    light_solid = has_gridline and color_light and solid_line

    # ③ 线宽 0.5–1 磅：读 a:ln@w（EMU）；未设=Excel 默认约0.75磅，落在范围内
    grid_w_emu = None
    if grid_ln is not None and grid_ln.get("w"):
        try:
            grid_w_emu = int(grid_ln.get("w"))
        except ValueError:
            grid_w_emu = None
    if grid_w_emu is None:
        width_ok = has_gridline    # 默认约0.75磅
        width_note = "线宽未显式设置，采用 Excel 默认约0.75磅"
    else:
        pt = grid_w_emu / 12700.0
        width_ok = 0.5 <= pt <= 1.0
        width_note = "线宽=%.2f磅" % pt

    # ④ 不遮挡柱形和数据标签：Excel 默认网格线渲染在数据系列之后（存在即满足）
    not_block = has_gridline

    add_rule(1, "变化趋势图纵轴主网格线", [
        ("绘图区内出现浅灰色水平网格线或三维透视网格线", has_gridline,
         "valAx majorGridlines=%s" % has_gridline),
        ("线条为浅灰色单实线", light_solid,
         "颜色=%s 线型=%s" % (grid_color or "默认浅灰", (dash_el.get("val") if dash_el is not None else "默认solid"))),
        ("线宽0.5磅–1磅", width_ok, width_note),
        ("网格线不遮挡柱形和数据标签", not_block, "网格线默认渲染于数据系列之后，不遮挡"),
    ])

    # +5 x3 年份柱形系列
    # +5 x3 年份柱形系列（严格对齐细则每一点）
    #   ① 对应 sheet1 表对应年销售额数据
    #   ② 柱形具有三维立体效果
    #   ③ 同一销售区域内该年柱形位于同组靠前或靠后固定位置（簇状排列，位置固定）
    #   ④ 数据点数量与销售区域数量一致
    #
    #   办公软件依据（OOXML）：
    #     ① ser/c:val 引用值逐一等于 sheet1 对应年份·各区域销售额
    #     ② bar3DChart 渲染带深度的立方体 => 三维立体效果
    #     ③ grouping=clustered + ser 的 c:order 固定 => 同一区域内各年份柱形
    #        按系列顺序并列于固定位置（同组靠前/靠后固定）
    #     ④ ser/c:cat 引用覆盖的分类数 = 销售区域数
    for yi, year in enumerate(ALLOWED_YEARS):
        ser = f.series[yi] if yi < len(f.series) else None
        vals = [num(v) for v in f.resolve_ref_values(ser["val_ref"])] if ser else []
        src_row = src.get(year, {})
        expect = [src_row.get(r) for r in src_regions]
        match = ser is not None and vals == expect and all(v is not None for v in vals)
        # ③ 固定位置：须为“规范可用的三维柱形图（三轴位置合理）”+ 簇状分组，
        #    且该系列有确定的 order（idx/order）。三轴位置损坏的伪三维图不算固定立体位置。
        fixed_pos = is_valid_3d and is_clustered
        add_rule(5, "%s年份柱形系列" % year, [
            ("对应sheet1表%s年销售额数据" % year, match,
             "系列值=%s 期望=%s" % (vals, expect)),
            ("柱形具有三维立体效果", is_valid_3d,
             "三维柱形图=%s 三轴位置合理=%s (type=%s, 三轴位置=%s)" % (
                 is_3d_bar, axes_pos_ok, f.chart_type, axis_positions)),
            ("同一销售区域内%s柱形位于同组靠前或靠后固定位置" % year, fixed_pos,
             "三维簇状(clustered)=%s 三轴位置合理=%s 系列固定顺序=%s" % (
                 is_clustered, axes_pos_ok, ser is not None)),
            ("数据点数量与销售区域数量一致（=%d）" % len(src_regions),
             ser is not None and (ser["cat_pts"] == len(src_regions)),
             "数据点=%s 销售区域数=%d" % ((ser["cat_pts"] if ser else None), len(src_regions))),
        ])

    # +3 2020、2021、2022柱形图使用三个不同颜色填充（严格对齐细则）
    #   细则唯一要求：2020、2021、2022 三个年份柱形使用三个不同颜色填充。
    #   办公软件依据（OOXML）：各 ser/c:spPr/a:solidFill/a:srgbClr 为该系列柱形填充色。
    colors = [s["color"] for s in f.series]
    valid_colors = [c for c in colors if c]
    # 三个年份系列均设置了填充色，且三色互不相同
    distinct = len(valid_colors) == 3 and len(set(valid_colors)) == 3
    add_rule(3, "2020、2021、2022柱形图使用三个不同颜色填充", [
        ("2020、2021、2022柱形使用三个不同颜色填充", distinct,
         "三系列填充色=%s" % colors),
    ])

    # +3 变化趋势图数据标签（严格对齐细则每一点）
    #   ① 每个柱形顶部或上方出现对应销售额数值标签
    #   ② 标签数值与 sheet1 表对应单元格一致
    #   ③ 标签字体为微软雅黑、Calibri 或宋体
    #   ④ 字号 8磅–12磅
    #   ⑤ 颜色与对应柱形系列相近或为深色
    #   ⑥ 标签不超出图表区
    #   ⑦ 标签颜色与下方对应柱形颜色一致
    #
    #   办公软件依据（OOXML）：数据标签 = bar3DChart/c:dLbls（图表级），
    #     系列级 = ser/c:dLbls，单点级 = ser/c:dLbls/c:dLbl(idx)。
    #     显示数值 = dLbls/c:showVal val="1"（每个数据点显示其销售额）
    #     标签文字颜色 = dLbls|dLbl/c:txPr/a:p/a:pPr/a:defRPr/a:solidFill/a:srgbClr。
    #     未显式设置时 Excel 默认标签为深色(黑)、约9–10磅、不超出图表区。
    #     细则最后一句“标签颜色与下方对应柱形颜色一致”要求逐（系列/数据点）标签
    #     颜色 == 对应柱形系列填充色，而非“深色即可”。
    bar_node = find(plot, f.chart_type) if plot is not None and f.chart_type else None
    dlbls_el = direct(bar_node, "dLbls") if bar_node is not None else None
    showval_el = direct(dlbls_el, "showVal") if dlbls_el is not None else None
    show_val = showval_el is not None and showval_el.get("val") in ("1", "true")

    # ② 标签数值与 sheet1 一致：数据标签直接取自各系列数值。此处只需校验
    #    "每条系列的 val 引用值 = sheet1 对应销售额"（即柱形系列子点①），
    #    与"三维立体效果/固定位置"等呈现类子点无关。
    def _series_values_all_match():
        if not f.series:
            return False
        for yi, year in enumerate(ALLOWED_YEARS):
            if yi >= len(f.series):
                return False
            ser = f.series[yi]
            vals = [num(v) for v in f.resolve_ref_values(ser["val_ref"])]
            expect = [src.get(year, {}).get(r) for r in src_regions]
            if not (vals and expect and all(v is not None for v in vals) and vals == expect):
                return False
        return True
    series_values_match = _series_values_all_match()
    labels_match = show_val and series_values_match

    # ③④⑤ 标签文字样式（字体/字号取图表级 dLbls/txPr 或系列级，供字体、字号判定）
    dlbl_props = get_run_props(find(dlbls_el, "txPr")) if dlbls_el is not None else {}
    if not dlbl_props.get("font") and not dlbl_props.get("size"):
        # 图表级未设时，退回首个系列级 dLbls/txPr
        for s in f.series:
            s_dlbls = direct(s.get("el"), "dLbls")
            sp = get_run_props(find(s_dlbls, "txPr")) if s_dlbls is not None else {}
            if sp.get("font") or sp.get("size"):
                dlbl_props = sp
                break
    dlbl_color = dlbl_props.get("color")
    # ⑤ 颜色为深色 或 与某个柱形系列填充色一致 => 满足"深色/与柱形相近/与柱形一致"
    color_ok = show_val and (
        dlbl_color is None or                          # 默认深色(黑)
        is_dark_color(dlbl_color) or                   # 深色
        dlbl_color in [c for c in colors if c]          # 与对应柱形颜色一致
    )

    # ⑦ 标签颜色与下方对应柱形颜色“逐一”一致：
    #   对每个系列，取其标签有效颜色（系列级 dLbls/txPr 优先，缺省回退图表级 dLbls/txPr），
    #   若该系列有单点级 dLbl(idx) 覆盖颜色则以单点颜色为准；再与该系列填充色比较。
    #   要求“每个”柱形标签颜色都等于其所在系列填充色，才判该子点命中。
    def _label_color_of(container):
        """从 dLbls/dLbl 容器的 txPr 里取标签文字颜色（大写hex或None）。"""
        if container is None:
            return None
        txpr = find(container, "txPr")
        return get_run_props(txpr).get("color") if txpr is not None else None

    chart_level_lbl_color = dlbl_props.get("color")
    per_series_color_notes = []
    labels_color_match_bars = show_val and bool(f.series)
    for s in f.series:
        fill = s.get("color")
        ser_el = s.get("el")
        ser_dlbls = direct(ser_el, "dLbls") if ser_el is not None else None
        ser_lbl_color = _label_color_of(ser_dlbls)
        # 单点级 dLbl 颜色（若存在任一点自定义颜色，逐点参与比较）
        point_colors = []
        if ser_dlbls is not None:
            for dlbl in directall(ser_dlbls, "dLbl"):
                point_colors.append(_label_color_of(dlbl))
        # 该系列所有“待比较标签颜色”集合：优先单点色，其余点用系列/图表级色
        effective = ser_lbl_color if ser_lbl_color is not None else chart_level_lbl_color
        compare_list = point_colors if point_colors else [effective]
        for lc in compare_list:
            # 标签颜色必须显式等于该系列柱形填充色，才算“与下方对应柱形颜色一致”。
            # 未设填充色，或标签颜色缺省/未显式设为柱色，均不满足“一致”。
            if not (fill and lc and lc == fill):
                labels_color_match_bars = False
        per_series_color_notes.append(
            "系列填充=%s 标签色=%s" % (fill, point_colors if point_colors else effective))

    add_rule(3, "变化趋势图数据标签", [
        ("每个柱形顶部或上方出现对应销售额数值标签", show_val,
         "showVal=%s" % show_val),
        ("标签数值与sheet1表对应单元格一致", labels_match,
         "标签取自系列数值，已与sheet1校验一致=%s" % series_values_match),
        ("标签字体为微软雅黑/Calibri/宋体",
         _font_ok(dlbl_props, wb, ("微软雅黑", "Calibri", "宋体")),
         _font_note(dlbl_props, wb)),
        ("字号8磅–12磅", show_val and _size_ok(dlbl_props, 8, 12),
         "字号=%s" % dlbl_props.get("size")),
        ("标签颜色与对应柱形系列相近或为深色", color_ok,
         "标签颜色=%s（默认深色）" % (dlbl_color or "默认(黑)")),
        ("标签不超出图表区", show_val, "采用 Excel 默认标签位置，不超出图表区"),
        ("标签颜色与下方对应柱形颜色一致", labels_color_match_bars,
         "; ".join(per_series_color_notes) if per_series_color_notes else "无数据标签/系列"),
    ])

    # +1 变化趋势图底部图例（严格对齐细则每一点）
    #   ① 位于图表底部或右下方空白区域
    #   ② 图例包含2020、2021、2022三个系列名称
    #   ③ 图例色块分别为三个不同颜色
    #   ④ 颜色与2020、2021、2022柱形图所用颜色对应
    #   ⑤ 字体为微软雅黑、宋体或 Calibri
    #   ⑥ 字号 8磅–12磅
    #   ⑦ 图例不遮挡横轴标题和销售区域标签
    #   ⑧ 图例从左到右或从上到下按2020、2021、2022顺序排列，与柱形系列顺序一致
    #
    #   办公软件依据（OOXML）：图例 = chart/c:legend
    #     位置 = legend/c:legendPos（b=底部，r=右侧，tr/br 等）
    #     文字样式 = legend/c:txPr（a:defRPr）
    #     是否叠加绘图区 = legend/c:overlay（val=0/缺省 => 不遮挡坐标轴标题与刻度标签）
    #     图例条目名与顺序默认取自各 ser/c:tx 与其 order（Excel 自动按系列顺序排列，
    #     色块自动使用对应系列填充色）。
    legend_pos = f.legend_pos
    # ① 底部或右下方
    legend_bottom = legend_pos in ("b", "r", "br", "tr")

    # ⑦ 不遮挡：legend/c:overlay 缺省或 val=0
    legend_overlay_el = direct(legend_el, "overlay") if legend_el is not None else None
    legend_no_overlap = not (legend_overlay_el is not None and
                             legend_overlay_el.get("val") in ("1", "true"))

    # ⑤⑥ 图例文字样式（legend/txPr）
    add_rule(1, "变化趋势图底部图例", [
        ("位于图表底部或右下方空白区域", legend_bottom, "legendPos=%s" % legend_pos),
        ("图例包含2020、2021、2022三个系列名称", years_included,
         "系列名=%s" % ser_names),
        ("图例色块分别为三个不同颜色", distinct, "三系列填充色=%s" % colors),
        ("颜色与2020、2021、2022柱形图所用颜色对应", distinct,
         "图例色块自动取自各系列填充色=%s" % colors),
        ("字体为微软雅黑/宋体/Calibri",
         _font_ok(legend_props, wb, ("微软雅黑", "宋体", "Calibri")),
         _font_note(legend_props, wb)),
        ("字号8磅–12磅", _size_ok(legend_props, 8, 12), "字号=%s" % legend_props.get("size")),
        ("图例不遮挡横轴标题和销售区域标签", legend_no_overlap,
         "overlay=%s" % (legend_overlay_el.get("val") if legend_overlay_el is not None else "缺省(不遮挡)")),
        ("图例按2020、2021、2022顺序排列，与柱形系列顺序一致", order_ok,
         "图例顺序=%s 柱形系列顺序=%s" % (ser_names, src_year_order)),
    ])

    # +1 变化趋势图坐标轴线（严格对齐细则每一点）
    #   ① 横轴、纵轴和深度轴均为深蓝色、黑色或深灰色单实线
    #   ② 线宽 0.75磅–1.25磅
    #   ③ 坐标轴线不缺失
    #   ④ 轴线与刻度标签位置对应
    #
    #   办公软件依据（OOXML）：坐标轴线样式 = 各轴(catAx/valAx/serAx)/c:spPr/a:ln
    #     颜色 = a:solidFill/a:srgbClr；线宽 = a:ln@w（EMU，1磅=12700）；
    #     线型 = a:prstDash（solid=单实线）；未显式设置时 Excel 默认坐标轴线为
    #       深灰/黑色、单实线、约0.75–1磅。轴线与刻度标签对应 = 该轴含 c:tickLblPos
    #       且非 "none"（Excel 默认 nextTo，刻度标签紧贴轴线）。
    axes = [("横轴", f.cat_ax), ("纵轴", f.val_ax), ("深度轴", f.ser_ax)]
    # 三根轴节点存在 且 三轴位置合理（axPos 不全相同）——三轴 axPos 全为同一值时，
    # 深度轴与横/纵轴挤在同一位置，办公软件中实际渲染不出可辨识的深度轴。
    axes_present = all(ax is not None for _, ax in axes) and axes_pos_ok

    def _axis_line_ok(ax):
        """返回 (颜色深色单实线 ok, 线宽 ok, 颜色说明, 线宽说明)。"""
        if ax is None:
            return False, False, "缺失", "缺失"
        ln = direct(direct(ax, "spPr"), "ln")
        # 颜色：未设=默认深色可；设了则须为深蓝/黑/深灰（深色）
        color = hexcolor(ln) if ln is not None else None
        color_ok = color is None or is_dark_color(color) or is_dark_blue(color)
        # 线型：未设=默认 solid；设了须为 solid
        dash = direct(ln, "prstDash") if ln is not None else None
        solid = dash is None or dash.get("val") == "solid"
        col_solid_ok = color_ok and solid
        # 线宽：未设=默认约0.75–1磅（可接受）；设了须落在 0.75–1.25
        w_emu = None
        if ln is not None and ln.get("w"):
            try:
                w_emu = int(ln.get("w"))
            except ValueError:
                w_emu = None
        if w_emu is None:
            width_ok = True
            w_note = "默认约0.75磅"
        else:
            pt = w_emu / 12700.0
            width_ok = 0.75 <= pt <= 1.25
            w_note = "%.2f磅" % pt
        return col_solid_ok, width_ok, (color or "默认深色单实线"), w_note

    lines_dark_solid = True
    widths_ok = True
    color_notes = []
    width_notes = []
    for nm, ax in axes:
        cs_ok, w_ok, cnote, wnote = _axis_line_ok(ax)
        lines_dark_solid = lines_dark_solid and cs_ok
        widths_ok = widths_ok and w_ok
        color_notes.append("%s=%s" % (nm, cnote))
        width_notes.append("%s=%s" % (nm, wnote))

    # ④ 轴线与刻度标签位置对应：各轴 tickLblPos 非 "none"
    def _ticklbl_shown(ax):
        if ax is None:
            return False
        el = direct(ax, "tickLblPos")
        return el is None or el.get("val") != "none"
    ticklbl_ok = all(_ticklbl_shown(ax) for _, ax in axes)

    add_rule(1, "变化趋势图坐标轴线", [
        ("横轴、纵轴和深度轴均为深蓝色/黑色/深灰色单实线", axes_present and lines_dark_solid,
         "; ".join(color_notes)),
        ("线宽0.75磅–1.25磅", axes_present and widths_ok, "; ".join(width_notes)),
        ("坐标轴线不缺失", axes_present,
         "catAx=%s valAx=%s serAx=%s 三轴位置合理=%s" % (
             f.cat_ax is not None, f.val_ax is not None, f.ser_ax is not None, axes_pos_ok)),
        ("轴线与刻度标签位置对应", ticklbl_ok,
         "各轴刻度标签显示(tickLblPos≠none)=%s" % ticklbl_ok),
    ])

    # ===================== 扣分项 =====================

    return f, rules


def _size_ok(props, lo, hi):
    sz = props.get("size")
    if sz is None:
        # 未显式设置 -> Excel 默认约10磅，落在常见 8–12 范围内
        return lo <= 10 <= hi
    return lo <= sz <= hi


# 会让坐标轴/图表标题默认以粗体外观显示的内置图表样式（<c:style val="N"/>）。
# 这些样式由主题赋予标题粗体字重，Excel 不会额外写入 b="1"，但视觉上即为加粗。
BOLD_TITLE_CHART_STYLES = {2, 4, 6, 8, 10, 12, 14, 16, 34, 36, 38, 40, 42, 44, 46, 48}


def _chart_style_num(wb):
    if wb.chart_xml is None:
        return None
    st = direct(wb.chart_xml, "style")
    if st is None:
        st = find(wb.chart_xml, "style")
    if st is not None and st.get("val"):
        try:
            return int(st.get("val"))
        except ValueError:
            return None
    return None


def _bold_ok(props, wb):
    """加粗判定：显式 b="1" 加粗，或图表样式使标题默认呈粗体外观。"""
    if props.get("bold", False):
        return True
    return _chart_style_num(wb) in BOLD_TITLE_CHART_STYLES


def _bold_note(props, wb):
    if props.get("bold", False):
        return "bold=True（显式加粗）"
    sn = _chart_style_num(wb)
    if sn in BOLD_TITLE_CHART_STYLES:
        return "bold=样式加粗（图表样式%s 使标题默认呈粗体外观）" % sn
    return "bold=False（图表样式%s 未使标题加粗）" % sn


def _effective_font(props, wb):
    """有效字体：优先标题显式字体；未显式设置时套用主题字体（图表默认字体）。"""
    font = props.get("font")
    if font:
        return font
    return getattr(wb, "theme_major_font", None) or getattr(wb, "theme_minor_font", None)


def _font_ok(props, wb, allowed):
    return _effective_font(props, wb) in allowed


def _font_note(props, wb):
    if props.get("font"):
        return "字体=%s（显式设置）" % props.get("font")
    return "字体=%s（未显式设置，套用主题字体）" % _effective_font(props, wb)


def _color_dark_or_default(props):
    col = props.get("color")
    if col is None:
        return True  # 默认自动色=黑色
    return is_dark_color(col) or is_dark_blue(col)


def _single_ref_value(f, ref):
    if not ref:
        return None
    if ":" in ref:
        vals = f.resolve_ref_values(ref)
        return vals[0] if vals else None
    # 单格引用，如 'Sheet2'!B2
    return f.resolve_ref_values(ref + ":" + ref.split("!")[-1])[0] if "!" in ref else None


# ----------------------------------------------------------------------------
# 统一对外入口
# ----------------------------------------------------------------------------

def _locate_target(dir_path):
    """在给定目录中定位被评估文档：优先精确文件名 TARGET，其次目录内首个 .xlsx/.xlsm。"""
    if not dir_path or not os.path.isdir(dir_path):
        return None
    exact = os.path.join(dir_path, TARGET)
    if os.path.isfile(exact):
        return exact
    for name in os.listdir(dir_path):
        low = name.lower()
        if low.endswith((".xlsx", ".xlsm")) and not name.startswith("~$"):
            return os.path.join(dir_path, name)
    return None


def _build_dim2_items(rules):
    """将内部 rules 结构转成 §2.2 要求的 dim2_items 列表。"""
    items = []
    for r in rules:
        pts = r["points"]
        hit = r["hit"]
        if r["kind"] == "add":
            max_delta = pts
            delta = pts if hit else 0
        else:  # sub：命中即扣分，最好情况为 0
            max_delta = pts
            delta = max_delta if hit else 0
        detail_lines = []
        for desc, ok, note in r["subpoints"]:
            detail_lines.append("[%s] %s | %s" % ("命中" if ok else "未命中", desc, note))
        items.append({
            "rule": r["name"],
            "max_delta": max_delta,
            "delta": delta,
            "hit": hit,
            "detail": "",
        })
    return items


def evaluate(dir_path: str) -> dict:
    """统一入口：接收脚本所在目录的路径，脚本自行在该目录定位并评估被评估文档。"""
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
        path = _locate_target(dir_path)
        if not path:
            result["status"] = "error"
            result["error"] = "在目录 %r 中未找到被评估文档（期望 %s 或任一 .xlsx/.xlsm）" % (
                dir_path, TARGET)
            return result
        result["file_name"] = os.path.basename(path)

        wb = Workbook(path)

        # ---------- 维度1（准入门槛）----------
        passed1, checks1 = check_dimension1(wb)
        result["dim1_pass"] = passed1
        if not passed1:
            failed = [desc for desc, ok, _ in checks1 if not ok]
            result["dim1_reason"] = "；".join(failed) if failed else "维度一未通过"
            result["total_score"] = 0
            result["max_score"] = 0
            return result

        # ---------- 维度2 ----------
        _, rules = eval_dimension2(wb)
        items = _build_dim2_items(rules)
        result["dim2_items"] = items
        result["total_score"] = sum(it["delta"] for it in items)
        result["max_score"] = sum(
            it["max_delta"] for it in items if it["max_delta"] > 0
        )
        return result
    except Exception as e:
        result["status"] = "error"
        result["error"] = "%s: %s" % (type(e).__name__, e)
        return result


if __name__ == "__main__":
    # 本地调试入口：默认使用脚本所在目录；亦可通过 argv[1] 指定目录。
    _dir = sys.argv[1] if len(sys.argv) > 1 else os.path.dirname(os.path.abspath(__file__))
    print(json.dumps(evaluate(_dir), ensure_ascii=False, indent=2))
