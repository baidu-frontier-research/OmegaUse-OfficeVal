# -*- coding: utf-8 -*-
"""
11～20岁身高统计图模板 —— 自动评估脚本（officeval_096）

评估逻辑：
  维度1（可用与可修改性）：硬性门槛。任意一条不满足 => 直接 0 分，不再检查维度2。
  维度2（完成度评分细则）：在维度1通过后逐条评估。
      - 加分细则：必须满足该细则中的【每一个】子点，才累加该细则的分数。
      - 扣分细则：满足该细则中【任意一个】子点，即扣除该细则的分数。
      （本模板细则全部为加分项，未给出扣分项；代码已预留扣分项处理框架。）

所有细则均由代码自动判定，不依赖人工。对于不易精确判定的点（如"图表能随数据自动更新"），
采用结构性判据灵活变通：图表数据源为指向工作表单元格的引用(numRef/c:f)而非写死的常量，
即认定为"数据驱动、可自动更新"。

对外接口（统一约定）：
  evaluate(dir_path: str) -> dict
    - 入参：脚本所在目录的路径；脚本自己在该目录里定位并打开被评估的 .xlsx/.xlsm 文档。
    - 返回：结构化字典（含 id/file_name/status/dim1_pass/dim2_items/total_score/max_score）。
"""

import os
import re
import sys
import json
import zipfile
import xml.etree.ElementTree as ET

# ---------------------------------------------------------------------------
# 命名空间
# ---------------------------------------------------------------------------
NS = {
    "x":  "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "c":  "http://schemas.openxmlformats.org/drawingml/2006/chart",
    "a":  "http://schemas.openxmlformats.org/drawingml/2006/main",
    "r":  "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "ct": "http://schemas.openxmlformats.org/package/2006/content-types",
    "rel":"http://schemas.openxmlformats.org/package/2006/relationships",
}

REQUIRED_AGES = [11, 12, 13, 14, 15, 16, 17, 18, 19, 20]


# ---------------------------------------------------------------------------
# 工具：列号 <-> 数字
# ---------------------------------------------------------------------------
def col_to_num(col):
    """'A'->1, 'B'->2 ... 'AA'->27"""
    n = 0
    for ch in col:
        n = n * 26 + (ord(ch.upper()) - ord('A') + 1)
    return n


def split_ref(ref):
    """'B7' -> ('B', 7)"""
    m = re.match(r"([A-Za-z]+)(\d+)", ref)
    return m.group(1), int(m.group(2))


# ---------------------------------------------------------------------------
# 读取 xlsx/xlsm 内部结构
# ---------------------------------------------------------------------------
class Workbook:
    """直接解析 OOXML 包，避免依赖会重算/丢弃图表的库。"""

    def __init__(self, path):
        self.path = path
        self.openable = False
        self.open_error = None
        self.zip = None
        self.names = []
        self.shared_strings = []
        self.sheets = []          # [{name, part, xmlroot}]
        self.charts = []          # [ET.Element(chartSpace)]
        self.content_types = None
        self._load()

    # -- 底层读取 ----------------------------------------------------------
    def _read(self, name):
        try:
            with self.zip.open(name) as f:
                return f.read()
        except KeyError:
            return None

    def _xml(self, name):
        data = self._read(name)
        if data is None:
            return None
        try:
            return ET.fromstring(data)
        except ET.ParseError:
            return None

    def _load(self):
        try:
            self.zip = zipfile.ZipFile(self.path, "r")
            self.names = self.zip.namelist()
        except (zipfile.BadZipFile, FileNotFoundError, OSError) as e:
            self.open_error = str(e)
            return

        # [Content_Types].xml —— 用于判断 xlsm/损坏
        self.content_types = self._read("[Content_Types].xml")

        # sharedStrings
        sst = self._xml("xl/sharedStrings.xml")
        if sst is not None:
            for si in sst.findall("x:si", NS):
                self.shared_strings.append(self._si_text(si))

        # workbook -> sheet 列表 + 关系
        wb = self._xml("xl/workbook.xml")
        rels = self._parse_rels("xl/_rels/workbook.xml.rels")
        if wb is not None:
            for sh in wb.findall(".//x:sheet", NS):
                name = sh.get("name", "")
                rid = sh.get("{%s}id" % NS["r"])
                target = rels.get(rid)
                part = self._norm_part("xl", target) if target else None
                root = self._xml(part) if part else None
                self.sheets.append({"name": name, "part": part, "xmlroot": root})

        # 收集所有图表 part
        for n in self.names:
            if re.match(r"xl/(drawings/)?charts/chart\d+\.xml$", n) or \
               re.match(r"xl/charts/chart\d+\.xml$", n):
                root = self._xml(n)
                if root is not None:
                    self.charts.append(root)

        self.openable = True

    @staticmethod
    def _si_text(si):
        parts = []
        for t in si.iter("{%s}t" % NS["x"]):
            parts.append(t.text or "")
        return "".join(parts)

    def _parse_rels(self, relpath):
        out = {}
        root = self._xml(relpath)
        if root is None:
            return out
        for rel in root.findall("rel:Relationship", NS):
            out[rel.get("Id")] = rel.get("Target")
        return out

    @staticmethod
    def _norm_part(base, target):
        # target 形如 "worksheets/sheet1.xml" 或 "/xl/worksheets/sheet1.xml"
        if target.startswith("/"):
            return target.lstrip("/")
        # 处理 ../
        path = base + "/" + target
        segs = []
        for s in path.split("/"):
            if s == "..":
                if segs:
                    segs.pop()
            elif s in ("", "."):
                continue
            else:
                segs.append(s)
        return "/".join(segs)

    # -- 便捷访问 ----------------------------------------------------------
    def cell_text(self, sheet, ref):
        """返回单元格显示文本（解析共享字符串 / inline / 数值）。"""
        root = sheet["xmlroot"]
        if root is None:
            return None
        for c in root.iter("{%s}c" % NS["x"]):
            if c.get("r") == ref:
                return self._cell_value(c)
        return None

    def _cell_value(self, c):
        t = c.get("t")
        v = c.find("x:v", NS)
        if t == "s":  # shared string
            if v is not None and v.text is not None:
                idx = int(v.text)
                if 0 <= idx < len(self.shared_strings):
                    return self.shared_strings[idx]
            return None
        if t == "inlineStr":
            is_el = c.find("x:is", NS)
            return self._si_text(is_el) if is_el is not None else None
        if v is not None:
            return v.text
        return None

    def iter_cells(self, sheet):
        """yield (ref, cell_element, text)"""
        root = sheet["xmlroot"]
        if root is None:
            return
        for c in root.iter("{%s}c" % NS["x"]):
            ref = c.get("r")
            if ref:
                yield ref, c, self._cell_value(c)

    def all_text(self):
        """工作簿中所有可见文本（标题/说明/共享串），用于关键字匹配。"""
        texts = list(self.shared_strings)
        for sh in self.sheets:
            for _ref, _c, txt in self.iter_cells(sh):
                if txt:
                    texts.append(txt)
        return texts


# ---------------------------------------------------------------------------
# 图表辅助
# ---------------------------------------------------------------------------
def chart_series(chart):
    """返回图表中所有 <c:ser>。"""
    return chart.findall(".//c:ser", NS)


def chart_type(chart):
    """返回主要图表类型标签集合，如 {'scatterChart'}, {'lineChart'} 等。"""
    types = set()
    plot = chart.find(".//c:plotArea", NS)
    if plot is None:
        return types
    for child in plot:
        tag = child.tag.split("}")[-1]
        if tag.endswith("Chart"):
            types.add(tag)
    return types


def numref_points(numref):
    """从 numRef/numCache 取出数值点列表。"""
    pts = []
    cache = numref.find("c:numCache", NS)
    if cache is None:
        return pts
    for pt in cache.findall("c:pt", NS):
        v = pt.find("c:v", NS)
        if v is not None and v.text is not None:
            try:
                pts.append(float(v.text))
            except ValueError:
                pass
    return pts


def series_refs(chart):
    """返回 [(xval_f, xpts, yval_f, ypts)]，f 为公式引用字符串。"""
    out = []
    for ser in chart_series(chart):
        xref = ser.find(".//c:xVal//c:numRef", NS)
        if xref is None:
            xref = ser.find(".//c:xVal//c:strRef", NS)
        yref = ser.find(".//c:yVal//c:numRef", NS)
        # 折线/柱状图用 cat/val 而非 xVal/yVal
        if yref is None:
            yref = ser.find(".//c:val//c:numRef", NS)
        if xref is None:
            xref = ser.find(".//c:cat//c:numRef", NS)
        if xref is None:
            xref = ser.find(".//c:cat//c:strRef", NS)
        xf = xref.find("c:f", NS).text if (xref is not None and xref.find("c:f", NS) is not None) else None
        yf = yref.find("c:f", NS).text if (yref is not None and yref.find("c:f", NS) is not None) else None
        xpts = numref_points(xref) if xref is not None else []
        ypts = numref_points(yref) if yref is not None else []
        out.append((xf, xpts, yf, ypts))
    return out


# ---------------------------------------------------------------------------
# 维度1：可用与可修改性（硬门槛）
# ---------------------------------------------------------------------------
def evaluate_dim1(wb):
    """返回 (passed: bool, details: list[(ok, desc)])"""
    details = []

    # 1.1 交付文件为 xlsx 或 .xlsm 格式，文件可正常打开
    ext = os.path.splitext(wb.path)[1].lower()
    ok_fmt = ext in (".xlsx", ".xlsm") and wb.openable
    desc = "交付文件为 xlsx 或 .xlsm 格式，文件可正常打开"
    if not wb.openable and wb.open_error:
        desc += "（打开失败：%s）" % wb.open_error
    details.append((ok_fmt, desc))

    passed = all(ok for ok, _ in details)
    return passed, details


# ---------------------------------------------------------------------------
# 维度2：完成度评分细则
# ---------------------------------------------------------------------------
def _find_age_height_table(wb):
    """
    在所有工作表中寻找"年龄行/列 + 身高行/列"结构。
    返回 dict: {
        'found': bool,                    # 是否定位到年龄数据表
        'ages': set(已覆盖年龄),
        'age_full': bool,                 # 年龄是否覆盖 11~20 全部 10 个
        'height_cells_editable': bool,    # 10 个年龄是否都有可编辑身高单元格
        'has_height_header': bool,        # 表内是否有"身高/cm"等身高数据行/列标识
        'sheet': sheet, 'orient': 'row'|'col',
        'age_refs': {age: ref}, 'age_row': int|None, 'age_col': str|None,
        'height_refs': [...], 'height_area': set(全部身高单元格ref),
        'table_top_row': int|None,        # 数据表所在的最上一行（用于"上方/下方"判定）
    }
    """
    best = {"found": False, "ages": set(), "age_full": False,
            "height_cells_editable": False, "has_height_header": False,
            "sheet": None, "orient": None, "age_refs": {}, "age_row": None,
            "age_col": None, "height_refs": [], "height_area": set(),
            "table_top_row": None}

    for s in wb.sheets:
        if s["xmlroot"] is None:
            continue
        # 建立 ref -> (cellobj, text) 索引
        cells = {}
        for ref, c, txt in wb.iter_cells(s):
            cells[ref] = (c, txt)

        # ---- 横向布局：某行是"年龄"，相邻行是"身高" ----
        # 找包含年龄数字 11..20 的行
        rows = {}
        for ref, (c, txt) in cells.items():
            col, row = split_ref(ref)
            rows.setdefault(row, []).append((col, c, txt))

        for row, items in rows.items():
            ages_here = {}      # age -> col
            for col, c, txt in items:
                val = _as_int(txt)
                if val in REQUIRED_AGES:
                    ages_here[val] = col
            age_set = set(ages_here.keys()) & set(REQUIRED_AGES)
            if len(age_set) >= 8:
                # 该行是年龄行；身高数据通常在下一行（同列）
                hrow = row + 1
                editable = 0
                hrefs = []
                harea = set()
                for age in REQUIRED_AGES:
                    col = ages_here.get(age)
                    if not col:
                        continue
                    href = "%s%d" % (col, hrow)
                    harea.add(href)
                    hcell = cells.get(href)
                    if hcell is not None and _is_editable_numeric(hcell[0]):
                        editable += 1
                        hrefs.append(href)
                # "身高/cm"标识：年龄行或身高行的行首标签含"身高"
                header_txt = _row_label(cells, row) + _row_label(cells, hrow)
                has_hdr = ("身高" in header_txt) or ("cm" in header_txt.lower())
                cand = {
                    "found": True, "ages": age_set,
                    "age_full": set(REQUIRED_AGES).issubset(age_set),
                    "height_cells_editable": editable >= 10,
                    "has_height_header": has_hdr,
                    "sheet": s, "orient": "row",
                    "age_refs": {a: "%s%d" % (ages_here[a], row) for a in age_set},
                    "age_row": row, "age_col": None,
                    "height_refs": hrefs, "height_area": harea,
                    "table_top_row": row,
                }
                if _better(cand, best):
                    best = cand

        # ---- 纵向布局：某列是"年龄"，相邻列是"身高" ----
        cols = {}
        for ref, (c, txt) in cells.items():
            col, row = split_ref(ref)
            cols.setdefault(col, []).append((row, c, txt))

        for col, items in cols.items():
            ages_here = {}      # age -> row
            for row, c, txt in items:
                val = _as_int(txt)
                if val in REQUIRED_AGES:
                    ages_here[val] = row
            age_set = set(ages_here.keys()) & set(REQUIRED_AGES)
            if len(age_set) >= 8:
                next_col = _col_shift(col, 1)
                editable = 0
                hrefs = []
                harea = set()
                min_row = min(ages_here.values()) if ages_here else None
                for age in REQUIRED_AGES:
                    row = ages_here.get(age)
                    if not row:
                        continue
                    href = "%s%d" % (next_col, row)
                    harea.add(href)
                    hcell = cells.get(href)
                    if hcell is not None and _is_editable_numeric(hcell[0]):
                        editable += 1
                        hrefs.append(href)
                # "身高/cm"标识：年龄列或身高列的列首表头含"身高"
                header_txt = _col_label(cells, col) + _col_label(cells, next_col)
                has_hdr = ("身高" in header_txt) or ("cm" in header_txt.lower())
                cand = {
                    "found": True, "ages": age_set,
                    "age_full": set(REQUIRED_AGES).issubset(age_set),
                    "height_cells_editable": editable >= 10,
                    "has_height_header": has_hdr,
                    "sheet": s, "orient": "col",
                    "age_refs": {a: "%s%d" % (col, ages_here[a]) for a in age_set},
                    "age_row": None, "age_col": col,
                    "height_refs": hrefs, "height_area": harea,
                    "table_top_row": min_row,
                }
                if _better(cand, best):
                    best = cand

    return best


def _row_label(cells, row):
    """取某一行所有单元格文本拼接（用于识别行表头，如'身高/cm'）。"""
    parts = []
    for ref, (_c, txt) in cells.items():
        _col, r = split_ref(ref)
        if r == row and txt:
            parts.append(str(txt))
    return " ".join(parts)


def _col_label(cells, col):
    """取某一列所有单元格文本拼接（用于识别列表头）。"""
    parts = []
    for ref, (_c, txt) in cells.items():
        c, _r = split_ref(ref)
        if c == col and txt:
            parts.append(str(txt))
    return " ".join(parts)


def _better(cand, best):
    score = (cand["age_full"], cand["has_height_header"],
             cand["height_cells_editable"], len(cand["ages"]))
    bscore = (best["age_full"], best["has_height_header"],
              best["height_cells_editable"], len(best["ages"]))
    return score > bscore


def _as_int(txt):
    try:
        f = float(txt)
        if f == int(f):
            return int(f)
    except (TypeError, ValueError):
        pass
    return None


def _is_editable_numeric(cell):
    """单元格可作为身高输入：是数值或空值，且不是被锁定的纯文本标题。
    公式单元格(<f>)视为派生显示，不算可填写输入。"""
    if cell.find("x:f", NS) is not None:
        return False
    t = cell.get("t")
    if t in (None, "n"):   # 数值或空
        return True
    # 文本但内容是纯数字也接受
    v = cell.find("x:v", NS)
    if v is not None and _as_int(v.text) is not None:
        return True
    # 空文本单元格（可填写）
    if v is None:
        return True
    return False


def _col_shift(col, delta):
    return _num_to_col(col_to_num(col) + delta)


def _num_to_col(n):
    s = ""
    while n > 0:
        n, rem = divmod(n - 1, 26)
        s = chr(ord('A') + rem) + s
    return s


def _chart_bound_to_table(wb, table):
    """
    图表是否"基于身高统计表"，且"修改身高数据后对应数据点能自动更新"。
    判据（结构性，等价于自动更新）：图表某系列的数值数据源 y 引用，
    经公式链最终落在身高统计表的可编辑身高单元格区域上。
    返回 (based_on_table: bool, auto_update: bool)
    """
    height_area = table.get("height_area", set())
    height_sheet = table["sheet"]["name"] if table.get("sheet") else None
    if not wb.charts:
        return False, False

    based = False
    auto = False
    for ch in wb.charts:
        for (_xf, _xp, yf, _yp) in series_refs(ch):
            if not yf:
                continue
            # y 数据源最终引用到的单元格集合（含经一层公式中转）
            targets = _resolve_ref_to_cells(wb, yf)
            if not targets:
                continue
            based = True
            # 是否命中身高统计表的可编辑身高单元格
            for (sh_name, ref) in targets:
                if (height_sheet is None or sh_name in (None, height_sheet)) \
                        and ref in height_area:
                    auto = True
                    break
        if auto:
            break
    return based, auto


def _resolve_ref_to_cells(wb, formula):
    """
    把图表数据源公式（如 '身高记录'!$B$13:$B$22）展开为单元格集合 [(sheet, ref)]。
    若这些单元格本身是公式（如 =IF(B7="",NA(),B7)），再向上追溯其引用的源单元格，
    从而判断图表是否最终绑定到可填写的身高输入单元格（=> 数据改动会自动传导更新）。
    """
    direct = _expand_ref(wb, formula)
    out = set(direct)
    for (sh_name, ref) in direct:
        sheet = _sheet_by_name(wb, sh_name)
        if sheet is None:
            continue
        cell = _get_cell(wb, sheet, ref)
        if cell is None:
            continue
        f = cell.find("x:f", NS)
        if f is not None and f.text:
            # 该图表点单元格是公式 —— 取公式中引用的源单元格
            for r in _refs_in_formula(f.text):
                out.add((sh_name, r))
    return out


def _expand_ref(wb, formula):
    """解析 \"'sheet'!$A$13:$B$22\" / 'Sheet1!A1' 为 [(sheet, 'A13'), ...]。"""
    cells = set()
    if not formula:
        return cells
    # 拆多区域（逗号分隔）
    for part in formula.split(","):
        part = part.strip()
        sh = None
        if "!" in part:
            sh, rng = part.split("!", 1)
            sh = sh.strip().strip("'")
        else:
            rng = part
        rng = rng.replace("$", "")
        if ":" in rng:
            a, b = rng.split(":", 1)
            for r in _cells_in_range(a, b):
                cells.add((sh, r))
        else:
            m = re.match(r"^[A-Za-z]+\d+$", rng)
            if m:
                cells.add((sh, rng))
    return cells


def _cells_in_range(a, b):
    ca, ra = split_ref(a)
    cb, rb = split_ref(b)
    n1, n2 = col_to_num(ca), col_to_num(cb)
    out = []
    for col in range(min(n1, n2), max(n1, n2) + 1):
        for row in range(min(ra, rb), max(ra, rb) + 1):
            out.append("%s%d" % (_num_to_col(col), row))
    return out


def _refs_in_formula(text):
    """从公式文本中抽取单元格引用（如 IF(B7="",NA(),B7) -> ['B7']）。"""
    refs = set()
    for m in re.finditer(r"\$?([A-Za-z]{1,3})\$?(\d+)", text):
        refs.add("%s%s" % (m.group(1), m.group(2)))
    return refs


def _sheet_by_name(wb, name):
    for s in wb.sheets:
        if s["name"] == name:
            return s
    # 引用未带表名时，默认取首个工作表
    if name is None and wb.sheets:
        return wb.sheets[0]
    return None


def _get_cell(wb, sheet, ref):
    root = sheet["xmlroot"]
    if root is None:
        return None
    for c in root.iter("{%s}c" % NS["x"]):
        if c.get("r") == ref:
            return c
    return None


# ---- 加分细则判定函数：每个返回 (passed, subpoints[(ok,desc)]) ----------
def rule_plus3(wb, table):
    """+3：逐点对应细则 ——
    工作表上方制作了"11～20岁身高统计表"或同义标题的数据表；
    表中年龄覆盖 11、12、…、20 岁；
    身高统计表包含可直接填写的身高/cm 数据行或数据列；
    10 个年龄均有对应的可编辑身高输入单元格；
    下方制作了基于身高统计表的可编辑 Excel 图表；
    修改任意年龄对应身高数据后，图表中对应数据点能自动更新。
    """
    subs = []

    # 点1：工作表"上方"制作了"11～20岁身高统计表"或同义标题的数据表
    #   注意：标题必须是"数据表"的标题，而非图表标题；
    #   因此不能在全工作簿文本里随意查找关键词（那样会把图表标题里的
    #   "身高统计图"/"成长记录"之类同义词也当成表标题命中，造成冒充）。
    #   这里改为：只在数据表附近区域（数据表上方若干行、同列范围内）查找标题文本。
    table_kw = ["身高统计表", "身高情况统计表", "身高记录表",
                "11～20岁身高", "11~20岁身高"]
    has_title = _table_has_nearby_title(wb, table, table_kw)
    # "上方"：数据表位于图表之上（数据表起始行号小于图表锚点行）
    above = _table_above_chart(wb, table)
    subs.append((has_title and above,
                 "工作表上方制作了“11～20岁身高统计表”或同义标题的数据表"))

    # 点2：表中年龄覆盖 11、12、13、14、15、16、17、18、19、20 岁
    ages_ok = table["found"] and table["age_full"]
    subs.append((ages_ok, "表中年龄覆盖 11、12、…、20 共 10 个（实测覆盖 %d 个）"
                 % (len(table["ages"]) if table["found"] else 0)))

    # 点3：身高统计表包含可直接填写的"身高/cm"数据行或数据列
    header_ok = table["found"] and table["has_height_header"]
    subs.append((header_ok,
                 "身高统计表含可直接填写的“身高/cm”数据行或数据列"))

    # 点4：10 个年龄均有对应的可编辑身高输入单元格
    cells_ok = table["found"] and table["height_cells_editable"]
    subs.append((cells_ok, "10 个年龄均有对应的可编辑身高输入单元格"))

    # 点5 + 点6：下方制作了基于身高统计表的可编辑 Excel 图表，
    #            且修改身高数据后对应数据点能自动更新
    based, auto = _chart_bound_to_table(wb, table)
    below = _chart_below_table(wb, table)
    subs.append((bool(wb.charts) and based and below,
                 "下方制作了基于身高统计表的可编辑 Excel 图表"))
    subs.append((auto,
                 "修改任意年龄对应身高数据后，图表中对应数据点能自动更新"))

    passed = all(ok for ok, _ in subs)   # 加分项：每一个点都要踩到才得分
    return passed, subs


def _table_above_chart(wb, table):
    """数据表是否在图表上方：数据表起始行 < 图表锚点起始行。无图表锚点信息时宽松返回 True。"""
    top = table.get("table_top_row")
    anchor = _chart_anchor_row(wb)
    if top is None or anchor is None:
        return True
    return top < anchor


def _chart_below_table(wb, table):
    """图表是否在数据表下方：图表锚点行 > 数据表（含身高区）最大行。"""
    anchor = _chart_anchor_row(wb)
    if anchor is None:
        return True
    rows = []
    if table.get("age_row"):
        rows.append(table["age_row"])
    for ref in table.get("height_area", set()):
        rows.append(split_ref(ref)[1])
    table_bottom = max(rows) if rows else None
    if table_bottom is None:
        return True
    return anchor >= table_bottom


def _chart_anchor_row(wb):
    """从 drawing XML 读取图表锚点的起始行（0-based -> 转 1-based）。取最小者。"""
    best = None
    for n in wb.names:
        if re.match(r"xl/drawings/drawing\d+\.xml$", n):
            root = wb._xml(n)
            if root is None:
                continue
            xdr = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
            for frm in root.iter("{%s}from" % xdr):
                rowel = frm.find("{%s}row" % xdr)
                if rowel is not None and rowel.text is not None:
                    r = int(rowel.text) + 1
                    best = r if best is None else min(best, r)
    return best


def rule_plus5_axes(wb, table):
    """+5：逐点对应细则 ——
    横轴：箭头指向为年龄；刻度范围覆盖 11~20 岁；
          横轴标签显示 11、12、…、20 或等效连续年龄刻度。
    纵轴：箭头指向为身高/cm；最小值为 0 或接近 0；最大值为 200cm 或接近 200cm；
          主刻度单位为 50cm；能显示 0、50、100、150、200 或接近该间隔的刻度效果。
    """
    subs = []
    if not wb.charts:
        # 细则的每个点都依赖图表，无图表则全部不满足
        subs.append((False, "横轴箭头指向为年龄"))
        subs.append((False, "横轴刻度范围覆盖 11~20 岁"))
        subs.append((False, "横轴标签显示 11、12、…、20 或等效连续年龄刻度"))
        subs.append((False, "纵轴箭头指向为身高/cm"))
        subs.append((False, "纵轴最小值为 0 或接近 0"))
        subs.append((False, "纵轴最大值为 200cm 或接近 200cm"))
        subs.append((False, "纵轴主刻度单位为 50cm"))
        subs.append((False, "纵轴能显示 0、50、100、150、200 或接近该间隔的刻度效果"))
        return False, subs

    # 取第一个含数据的图表
    chart = None
    for ch in wb.charts:
        if series_refs(ch):
            chart = ch
            break
    if chart is None:
        chart = wb.charts[0]

    # --- 解析坐标轴 ---
    val_axes = chart.findall(".//c:valAx", NS)
    cat_axes = chart.findall(".//c:catAx", NS)
    # 横轴：年龄。散点图横轴是 valAx(axPos=b)，折线/柱状是 catAx。
    x_axis = _axis_by_pos(val_axes + cat_axes, "b")
    y_axis = _axis_by_pos(val_axes + cat_axes, "l")

    # ===== 横轴 =====
    # 点1：横轴箭头指向为年龄
    #   "箭头指向为年龄"——横轴用于表达年龄维度。判据：横轴绑定的是年龄数据
    #   （x 系列值落在 11~20 年龄集合）或横轴标题/标签语义为年龄。
    x_is_age = _axis_is_age(x_axis, wb, chart, table)
    subs.append((x_is_age, "横轴箭头指向为年龄"))

    # 点2：横轴刻度范围覆盖 11~20 岁
    #   类别轴（catAx，折线图/柱状图常见）通常不写 c:scaling/min/max，
    #   用 _axis_num 取值会恒为 None 而误判失败；这类轴的"范围"应由实际
    #   类别标签集合是否覆盖 11~20 来判断。数值轴（valAx，散点图常见）才用
    #   min/max 判断范围覆盖。
    x_min = _axis_num(x_axis, "min")
    x_max = _axis_num(x_axis, "max")
    if x_axis is not None and x_axis.tag.endswith("}catAx"):
        cat_ages = _axis_category_values(chart)
        range_cover = set(REQUIRED_AGES).issubset(cat_ages)
        range_desc = "类别标签=%s" % (",".join(str(a) for a in sorted(cat_ages)) or "无")
    else:
        range_cover = (x_min is not None and x_max is not None
                       and x_min <= 11 and x_max >= 20)
        range_desc = "min=%s, max=%s" % (_fmt(x_min), _fmt(x_max))
    subs.append((range_cover,
                 "横轴刻度范围覆盖 11~20 岁（实测 %s）" % range_desc))

    # 点3：横轴标签显示 11、12、…、20 或等效连续年龄刻度
    #   判据：横轴标签集合（散点图 = x 数据点值；类别轴 = 类别文本）含连续 11..20，
    #   或主刻度单位为 1 且范围覆盖 11~20（=> 必然逐格显示 11,12,...,20）。
    labels_ok = _axis_labels_show_ages(x_axis, wb, chart)
    subs.append((labels_ok,
                 "横轴标签显示 11、12、…、20 或等效连续年龄刻度"))

    # ===== 纵轴 =====
    # 点4：纵轴箭头指向为身高/cm
    #   判据：纵轴绑定身高数据系列，或纵轴标题/系列名含"身高"或"cm"。
    y_is_height = _axis_is_height(y_axis, wb, chart, table)
    subs.append((y_is_height, "纵轴箭头指向为身高/cm"))

    y_min = _axis_num(y_axis, "min")
    y_max = _axis_num(y_axis, "max")
    y_unit = _axis_num(y_axis, "majorUnit")

    # 点5/6/7：纵轴参数允许"接近"，但容差要收紧并统一。
    # 这里统一按 ±2 处理：足以覆盖少量四舍五入/模板微调，
    # 也避免把明显偏离的刻度（例如 0±10、200±20、50±10）误算进去。
    tol = 2

    # 点5：纵轴最小值为 0 或接近 0
    min_ok = (y_min is not None and abs(y_min - 0) <= tol)
    subs.append((min_ok, "纵轴最小值为 0 或接近 0（实测 %s，容差±%s）" % (_fmt(y_min), tol)))

    # 点6：纵轴最大值为 200cm 或接近 200cm
    max_ok = (y_max is not None and abs(y_max - 200) <= tol)
    subs.append((max_ok, "纵轴最大值为 200cm 或接近 200cm（实测 %s，容差±%s）" % (_fmt(y_max), tol)))

    # 点7：纵轴主刻度单位为 50cm
    unit_ok = (y_unit is not None and abs(y_unit - 50) <= tol)
    subs.append((unit_ok, "纵轴主刻度单位为 50cm（实测 %s，容差±%s）" % (_fmt(y_unit), tol)))

    # 点8：能显示 0、50、100、150、200 或接近该间隔的刻度效果
    #   由 min≈0、max≈200、unit≈50 共同决定的刻度序列；
    #   据此生成实际主刻度序列，逐一比对是否接近 {0,50,100,150,200}。
    ticks_ok, tick_seq = _ticks_match_target(y_min, y_max, y_unit,
                                             target=[0, 50, 100, 150, 200])
    subs.append((ticks_ok,
                 "纵轴能显示 0、50、100、150、200 或接近该间隔的刻度效果（实测刻度：%s）"
                 % (",".join(_fmt(t) for t in tick_seq) if tick_seq else "无")))

    passed = all(ok for ok, _ in subs)   # 加分项：每一个点都要踩到才得分
    return passed, subs


def _axis_is_age(axis, wb, chart, table):
    """横轴是否表达年龄维度。"""
    # 1) 横轴绑定的 x 系列值集合是否落在年龄域 11..20
    xs = set()
    for (xf, xp, _yf, _yp) in series_refs(chart):
        for v in xp:
            if v == int(v):
                xs.add(int(v))
    if xs and (set(xs) & set(REQUIRED_AGES)) and max(xs) <= 30:
        return True
    # 2) 横轴标题文本含"年龄"
    if axis is not None and _axis_title_text(axis) and "年龄" in _axis_title_text(axis):
        return True
    # 3) x 数据源引用年龄行/列单元格
    if table.get("found"):
        age_refs = set(table.get("age_refs", {}).values())
        for (xf, _xp, _yf, _yp) in series_refs(chart):
            if not xf:
                continue
            cells = {ref for (_sh, ref) in _expand_ref(wb, xf)}
            if cells & age_refs:
                return True
    return False


def _axis_is_height(axis, wb, chart, table):
    """纵轴是否表达身高/cm 维度。"""
    # 1) 纵轴标题含"身高"或"cm"
    t = _axis_title_text(axis)
    if t and ("身高" in t or "cm" in t.lower()):
        return True
    # 2) 系列名（y 系列）含"身高"或"cm"
    for ser in chart_series(chart):
        tx = ser.find(".//c:tx//c:v", NS)
        name = tx.text if tx is not None and tx.text else ""
        if "身高" in name or "cm" in name.lower():
            return True
    # 3) y 数据源引用身高数据区域
    if table.get("found"):
        harea = table.get("height_area", set())
        for (_xf, _xp, yf, _yp) in series_refs(chart):
            if not yf:
                continue
            targets = _resolve_ref_to_cells(wb, yf)
            if {ref for (_sh, ref) in targets} & harea:
                return True
    return False


def _title_text_from_chart(wb, chart):
    """提取图表标题文本。

    优先读取 c:title 内的富文本 a:t；若标题通过 c:tx/c:strRef/c:f 绑定到单元格，
    则按公式引用解析到工作表单元格并拼接单元格文本。只处理传入的这个 chart，
    避免把工作簿里其它图表的标题混进来。
    """
    title = chart.find("c:title", NS)
    if title is None:
        return ""

    # 1) 富文本标题：<c:title><c:tx><c:rich>...<a:t>
    txts = [t.text for t in title.iter("{%s}t" % NS["a"]) if t.text]
    rich_text = "".join(txts).strip()
    if rich_text:
        return rich_text

    # 2) 单元格引用标题：<c:title><c:tx><c:strRef><c:f>Sheet!$A$1</c:f>
    str_ref = title.find(".//c:tx/c:strRef", NS)
    if str_ref is None:
        return ""
    f = str_ref.find("c:f", NS)
    if f is None or not f.text:
        return ""

    parts = []
    for sh_name, ref in _expand_ref(wb, f.text):
        sheet = _sheet_by_name(wb, sh_name)
        if sheet is None:
            continue
        txt = wb.cell_text(sheet, ref)
        if txt:
            parts.append(str(txt).strip())
    return " ".join(p for p in parts if p)


def _axis_title_text(axis):
    if axis is None:
        return ""
    title = axis.find("c:title", NS)
    if title is None:
        return ""
    return "".join(t.text or "" for t in title.iter("{%s}t" % NS["a"]))


def _axis_category_values(chart):
    """取类别轴（catAx）的类别标签值集合（作为整数年龄）。
    与 _axis_labels_show_ages 中 B) 分支使用同一数据源：系列的 c:cat。
    用于类别轴场景下判断"刻度范围覆盖 11~20"——类别轴没有 min/max，
    其"范围"就是类别标签本身覆盖的取值集合。
    """
    cats = set()
    for ser in chart_series(chart):
        cat = ser.find(".//c:cat", NS)
        if cat is not None:
            for v in cat.iter("{%s}v" % NS["c"]):
                iv = _as_int(v.text)
                if iv is not None:
                    cats.add(iv)
    return cats


def _axis_labels_show_ages(axis, wb, chart):
    """横轴标签是否显示 11..20 或等效连续年龄刻度。"""
    # A) 数值轴（散点图常见）：标签由 min/max/majorUnit 决定。
    #    要保证 11~20 十个整数刻度都能显示，主刻度单位必须恰为 1——
    #    单位为 2 时刻度落在 11,13,15,...或 12,14,16,...，无法保证十个年龄
    #    全部显示，因此不再把 unit=2 当作"等效连续刻度"接受。
    if axis is not None:
        amin = _axis_num(axis, "min")
        amax = _axis_num(axis, "max")
        unit = _axis_num(axis, "majorUnit")
        if amin is not None and amax is not None and unit is not None:
            if amin <= 11 and amax >= 20 and abs(unit - 1) < 1e-6:
                return True
    # B) 类别轴：直接取类别文本，判断是否含连续 11..20
    cats = _axis_category_values(chart)
    if set(REQUIRED_AGES).issubset(cats):
        return True
    # C) 散点 x 数据点值含完整 11..20（等效连续年龄刻度标签）
    xs = set()
    for (_xf, xp, _yf, _yp) in series_refs(chart):
        for v in xp:
            iv = _as_int(v)
            if iv is not None:
                xs.add(iv)
    return set(REQUIRED_AGES).issubset(xs)


def _ticks_match_target(amin, amax, unit, target):
    """由 min/max/unit 生成主刻度序列，并判断是否覆盖/接近 target 刻度。"""
    if amin is None or amax is None or unit is None or unit <= 0:
        return False, []
    ticks = []
    v = amin
    # 防御：限制循环次数
    guard = 0
    while v <= amax + 1e-6 and guard < 1000:
        ticks.append(round(v, 6))
        v += unit
        guard += 1
    if not ticks:
        return False, ticks
    # target 中的每个目标刻度，在实测刻度里都能找到一个接近值（容差 ±5cm）
    tol = 5
    all_hit = all(any(abs(t - tg) <= tol for t in ticks) for tg in target)
    return all_hit, ticks


def rule_plus5_markers_only(wb, table):
    """+5：逐点对应细则 ——
    点1：图表显示数据点（marker）。
    点2：不显示连接各数据点的折线。
    两点都要踩到才计分。
    """
    subs = []
    if not wb.charts:
        subs.append((False, "图表显示数据点"))
        subs.append((False, "不显示连接各数据点的折线"))
        return False, subs

    # 取含数据系列的图表（与其它细则一致的选取口径）
    chart = None
    for ch in wb.charts:
        if chart_series(ch):
            chart = ch
            break
    if chart is None:
        chart = wb.charts[0]

    types = chart_type(chart)
    shows_markers, line_shown, detail = _marker_line_state(chart, types)

    # 点1：显示数据点
    subs.append((shows_markers, "图表显示数据点（%s）" % detail))
    # 点2：不显示连接各数据点的折线
    subs.append((not line_shown, "不显示连接各数据点的折线（%s）" % detail))

    passed = all(ok for ok, _ in subs)
    return passed, subs


def _marker_line_state(chart, types):
    """
    判定图表的"是否显示数据点"与"是否显示连线"。
    返回 (shows_markers: bool, line_shown: bool, detail: str)
    """
    # ---- 散点图 ----
    if "scatterChart" in types:
        style = chart.find(".//c:scatterStyle", NS)
        sval = style.get("val") if style is not None else None
        # scatterStyle 的 OOXML 语义（ECMA-376 §21.2.2.140）：
        #   marker       —— 仅标记，不连线
        #   lineMarker   —— 标记 + 直线连接
        #   line         —— 仅直线，不显示标记
        #   smooth       —— 仅平滑线
        #   smoothMarker —— 标记 + 平滑线
        #   none         —— 未设置，图表应用程序按默认（通常为 lineMarker）渲染
        # scatterStyle=marker 是"不显示连接线"的显式声明，可作为强证据直接采信；
        # 无需再要求每个系列都额外写 a:ln/a:noFill——那是过严的要求，
        # 会把 Office 中标准的"仅标记散点图"误判为有连线。
        # 系列级设置仅在可能覆盖/矛盾全局样式时才参与判断：
        #   - 若某系列显式将连线设为 noFill，则该系列一定无连接线；
        #   - 若某系列显式设置了非 none 的连线（有宽度/颜色的 a:ln），
        #     则该系列存在连接线，即使全局 scatterStyle=marker 也不能认为无线。
        no_line_styles = ("marker",)
        line_styles = ("line", "lineMarker", "smooth", "smoothMarker")
        if sval in no_line_styles:
            line_shown = _any_series_line_explicit_on(chart)
        elif sval in line_styles:
            line_shown = not _all_series_line_off(chart)
        else:
            # scatterStyle 缺失/none：无法确定默认渲染，退回系列级判据——
            # 只有当所有系列都显式关闭连线时才认为无连接线。
            line_shown = not _all_series_line_off(chart)
        # 显示数据点：style 含 marker，或系列显式设置了 marker 符号(非 none)
        shows_markers = (sval in ("marker", "lineMarker", "smoothMarker")) \
            or _any_series_marker_on(chart)
        # 散点图默认即显示点；style=none 时若无明确关闭 marker 也按显示点处理
        if sval in (None, "none") and not _any_series_marker_off(chart):
            shows_markers = True
        detail = "散点图 scatterStyle=%s" % sval
        return shows_markers, line_shown, detail

    # ---- 折线图 ----
    if "lineChart" in types:
        # 折线图默认显示连线；仅当每个系列连线都 noFill 才算"不显示折线"
        line_shown = not _all_series_line_off(chart)
        shows_markers = _any_series_marker_on(chart)
        detail = "折线图"
        return shows_markers, line_shown, detail

    # ---- 其它类型 ----
    # 细则只关心"数据点"与"连接折线"。非散点/折线类（柱状/饼图等）没有
    # 连接各数据点的折线，但也并非以数据点形式呈现 —— 按不满足细则处理。
    detail = "图表类型 %s（非散点/折线，不符合“只显示数据点”形态）" % (",".join(types) or "未知")
    return False, False, detail


def _all_series_line_off(chart):
    """所有数据系列的连线是否都被显式隐藏（a:ln/a:noFill）。无系列则 False。"""
    sers = chart_series(chart)
    if not sers:
        return False
    for ser in sers:
        nofill = ser.find(".//c:spPr/a:ln/a:noFill", NS)
        if nofill is None:
            return False
    return True


def _any_series_line_explicit_on(chart):
    """是否存在系列显式启用了连接线（a:ln 设置了非 noFill 的填充，
    如 a:solidFill/a:pattFill/a:gradFill，或指定了线宽 w）。
    用于在 scatterStyle=marker（全局声明"仅标记、不连线"）时，
    识别系列级样式覆盖全局设置、仍画出连接线的情形。"""
    for ser in chart_series(chart):
        ln = ser.find(".//c:spPr/a:ln", NS)
        if ln is None:
            continue
        if ln.find("a:noFill", NS) is not None:
            continue
        has_fill = any(ln.find("a:%s" % tag, NS) is not None
                        for tag in ("solidFill", "gradFill", "pattFill"))
        if has_fill or ln.get("w") is not None:
            return True
    return False


def _any_series_marker_on(chart):
    """是否有系列显式启用了数据点标记（c:marker/c:symbol != none）。"""
    for ser in chart_series(chart):
        sym = ser.find(".//c:marker/c:symbol", NS)
        if sym is not None and sym.get("val") not in (None, "none"):
            return True
    return False


def _any_series_marker_off(chart):
    """是否存在系列显式关闭了数据点标记（c:marker/c:symbol = none）。"""
    for ser in chart_series(chart):
        sym = ser.find(".//c:marker/c:symbol", NS)
        if sym is not None and sym.get("val") == "none":
            return True
    return False


def rule_plus1_title(wb, table):
    """+1：逐点对应细则 ——
    点1：图表具有标题。
    点2：标题包含“身高”“统计图”或“成长记录”等能说明图表用途的文字。
    两点都要踩到才计分。
    """
    subs = []

    # 点1：图表具有标题
    # 这里按“每个图表自己的标题”判断，不从工作簿全局取文本。
    # 既支持富文本标题，也支持通过 c:tx/c:strRef/c:f 绑定到单元格的标题。
    title_text = ""
    has_title = False
    for ch in wb.charts:
        deleted = ch.find(".//c:autoTitleDeleted", NS)
        if deleted is not None and deleted.get("val") == "1":
            continue
        title_text = _title_text_from_chart(wb, ch)
        if title_text.strip():
            has_title = True
            break
    subs.append((has_title, "图表具有标题（标题：%s）" % (title_text or "无")))

    # 点2：标题包含"身高""统计图"或"成长记录"等能说明图表用途的文字
    kw = ["身高", "统计图", "成长记录"]
    kw_ok = has_title and any(k in title_text for k in kw)
    subs.append((kw_ok,
                 "标题包含“身高”“统计图”或“成长记录”等能说明图表用途的文字"))

    passed = all(ok for ok, _ in subs)
    return passed, subs


# ---- 坐标轴解析辅助 ----------------------------------------------------
def _axis_by_pos(axes, pos):
    for ax in axes:
        p = ax.find("c:axPos", NS)
        if p is not None and p.get("val") == pos:
            return ax
    return axes[0] if axes else None


def _axis_num(axis, kind):
    if axis is None:
        return None
    if kind in ("min", "max"):
        node = axis.find("c:scaling/c:%s" % kind, NS)
    else:  # majorUnit
        node = axis.find("c:%s" % kind, NS)
    if node is not None and node.get("val") is not None:
        try:
            return float(node.get("val"))
        except ValueError:
            return None
    return None


def _table_has_nearby_title(wb, table, keywords):
    """判断"数据表标题"，而非在全工作簿文本中盲目搜关键词。

    问题背景：原实现在全工作簿文本（wb.all_text()）中查找标题关键词，
    未绑定到具体的数据表；同时把"身高统计图""成长记录"等图表标题的
    同义词也纳入表标题关键词，导致图表标题被误判为数据表标题（冒充）。

    修正：只在数据表所在工作表、且位于数据表上方邻近区域（数据表起始行
    及其上方最多 3 行、覆盖数据表所跨列范围）内查找标题文本；不再匹配
    "身高统计图""成长记录"等明显属于图表标题的措辞，避免与
    rule_plus1_title 的图表标题关键词混淆。
    """
    if not table.get("found") or table.get("sheet") is None:
        return False

    sheet = table["sheet"]
    top_row = table.get("table_top_row")
    if top_row is None:
        return False

    # 数据表跨列范围：由年龄引用和身高区域的列共同确定
    cols = set()
    for ref in table.get("age_refs", {}).values():
        c, _r = split_ref(ref)
        cols.add(col_to_num(c))
    for ref in table.get("height_area", set()):
        c, _r = split_ref(ref)
        cols.add(col_to_num(c))
    if table.get("age_col"):
        cols.add(col_to_num(table["age_col"]))
    if not cols:
        return False
    min_col, max_col = min(cols), max(cols)

    # 标题通常在数据表正上方 1~3 行，允许略微超出数据表列范围（左右各放宽 1 列）
    search_rows = range(max(1, top_row - 3), top_row)
    search_cols = range(max(1, min_col - 1), max_col + 2)

    for ref, _c, txt in wb.iter_cells(sheet):
        if not txt:
            continue
        col, row = split_ref(ref)
        if row not in search_rows:
            continue
        if col_to_num(col) not in search_cols:
            continue
        if any(k in txt for k in keywords):
            return True
    return False


def _fmt(v):
    if v is None:
        return "无"
    if v == int(v):
        return str(int(v))
    return str(v)


# ---------------------------------------------------------------------------
# 维度2 评估编排
# ---------------------------------------------------------------------------
# 评分细则表：(分值, 类型, 细则内容原文, 判定函数)
#   type = "plus" 加分项（须满足全部子点）/ "minus" 扣分项（满足任一子点即扣）
SCORING_RULES = [
    (+3, "plus",
     "工作表上方制作了“11～20岁身高统计表”或同义标题的数据表，表中年龄覆盖11、12、13、14、15、16、17、18、19、20岁。"
     "身高统计表包含可直接填写的身高/cm数据行或数据列，10个年龄均有对应的可编辑身高输入单元格。"
     "下方制作了基于身高统计表的可编辑Excel图表，修改任意年龄对应身高数据后，图表中的对应数据点能自动更新。",
     rule_plus3),
    (+5, "plus",
     "图表横轴箭头指向为年龄，刻度范围覆盖11～20岁，横轴标签显示11、12、13、14、15、16、17、18、19、20或等效连续年龄刻度。"
     "纵轴箭头指向为身高/cm，最小值为0或接近0，最大值为200cm或接近200cm；"
     "纵轴主刻度单位为50cm，能显示0、50、100、150、200或接近该间隔的刻度效果。",
     rule_plus5_axes),
    (+5, "plus",
     "图表只显示数据点，不显示连接各数据点的折线。",
     rule_plus5_markers_only),
    (+1, "plus",
     "图表具有标题，标题包含“身高”“统计图”或“成长记录”等能说明图表用途的文字。",
     rule_plus1_title),
    # 扣分项（本模板细则未给出，留作扩展示例）：
    # (-2, "minus", "示例扣分项内容", rule_minus_example),
]


def evaluate_dim2(wb):
    table = _find_age_height_table(wb)
    results = []   # [(score, rtype, content, hit, applied, subpoints)]
    total = 0
    for score, rtype, content, func in SCORING_RULES:
        passed, subs = func(wb, table)
        if rtype == "plus":
            hit = passed                      # 加分：全部子点满足
        else:  # minus
            hit = any(ok for ok, _ in subs)   # 扣分：任一子点满足即触发
        applied = score if hit else 0
        total += applied
        results.append((score, rtype, content, hit, applied, subs))
    return total, results, table


# ---------------------------------------------------------------------------
# 报告打印
# ---------------------------------------------------------------------------
def _mark(ok):
    return "[√]" if ok else "[×]"


# ---------------------------------------------------------------------------
# 统一对外接口
# ---------------------------------------------------------------------------
SCRIPT_ID = "096"


def _locate_target_file(dir_path):
    """在给定目录里定位被评估的 .xlsx/.xlsm 文档。

    选取规则：
      1) 过滤掉临时文件（以 ~$ 开头）、以 _ 开头的辅助文件（如 _copy.xlsx）、
         以及非 xlsx/xlsm 后缀。
      2) 若目录中恰好有一个候选文件，直接返回；
      3) 若存在多个候选，则优先选择文件名包含"身高"关键词的那个；
      4) 否则按文件名排序返回第一个。
    """
    if not os.path.isdir(dir_path):
        return None
    candidates = []
    for name in os.listdir(dir_path):
        low = name.lower()
        if name.startswith("~$") or name.startswith("_"):
            continue
        if not (low.endswith(".xlsx") or low.endswith(".xlsm")):
            continue
        candidates.append(name)
    if not candidates:
        return None
    if len(candidates) == 1:
        return os.path.join(dir_path, candidates[0])
    # 优先包含"身高"关键词的候选
    preferred = [n for n in candidates if "身高" in n]
    pick = preferred[0] if preferred else sorted(candidates)[0]
    return os.path.join(dir_path, pick)


def _build_dim2_items(results, max_total):
    """将 evaluate_dim2 的原始结果转换为统一的 dim2_items 列表。"""
    items = []
    for score, rtype, content, hit, applied, subs in results:
        # detail：逐子点 [√]/[×] 说明，便于排查（此处按需求置空，保留计算以维持原逻辑）
        detail = "; ".join(("%s %s" % (_mark(ok), desc)) for ok, desc in subs)
        _ = detail  # 保留原始 detail 计算以不破坏逻辑，输出中置空
        items.append({
            "rule": content,
            "max_delta": score,
            "delta": applied,
            "hit": bool(hit),
            "detail": "",
        })
    return items


def evaluate(dir_path: str) -> dict:
    """统一入口：接收脚本所在目录的路径，脚本自身在该目录内定位并评估文档。"""
    result = {
        "id": SCRIPT_ID,
        "file_name": None,
        "status": "ok",
        "error": None,
        "dim1_pass": False,
        "dim1_reason": "",
        "dim2_items": [],
        "total_score": 0,
        "max_score": sum(score for score, rtype, *_ in SCORING_RULES
                         if rtype == "plus"),
    }

    try:
        path = _locate_target_file(dir_path)
        if not path or not os.path.exists(path):
            result["status"] = "error"
            result["error"] = "在目录中未找到被评估的 .xlsx/.xlsm 文件：%s" % dir_path
            return result
        result["file_name"] = os.path.basename(path)

        wb = Workbook(path)

        # ---- 维度1 ----
        d1_pass, d1_details = evaluate_dim1(wb)
        result["dim1_pass"] = bool(d1_pass)
        if not d1_pass:
            fails = [desc for ok, desc in d1_details if not ok]
            result["dim1_reason"] = "; ".join(fails)
            # 维度一未通过：dim2_items 保持为空，总分为 0
            return result

        # ---- 维度2 ----
        total, results_raw, _table = evaluate_dim2(wb)
        result["dim2_items"] = _build_dim2_items(results_raw, result["max_score"])
        result["total_score"] = int(total)
        return result
    except Exception as e:  # 兜底：任何未预期异常都归为 status=error
        result["status"] = "error"
        result["error"] = "%s: %s" % (type(e).__name__, e)
        return result


if __name__ == "__main__":
    # 本地调试：默认以脚本所在目录为参数；也允许通过命令行显式指定目录。
    # 使用 ensure_ascii=True 避免在 GBK 等非 UTF-8 控制台下崩溃
    # （遵循约定：脚本本身不修改 sys.stdout）。
    if len(sys.argv) >= 2:
        _dir = sys.argv[1]
    else:
        _dir = os.path.dirname(os.path.abspath(__file__))
    print(json.dumps(evaluate(_dir), ensure_ascii=True, indent=2))
