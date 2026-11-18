#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# pyright: reportUnknownMemberType=false, reportUnknownVariableType=false
# pyright: reportUnknownArgumentType=false, reportMissingTypeStubs=false
# pyright: reportOptionalMemberAccess=false, reportGeneralTypeIssues=false
#
# 说明：按接口统一约定，本模块不修改全局 sys.stdout，也不 print 主结果。
# 主结果统一由 evaluate(dir_path) -> dict 返回；入参为“脚本所在目录的路径”，
# 由脚本自己在该目录内定位并打开被评估的 .xlsx/.xlsm。本文件仅在 __main__ 下
# 作为本地调试入口，才把 dict 序列化为 JSON 打印。
#
# 为什么不用 openpyxl：
#   ① 数据条条件格式的 x14 扩展（minLength/maxLength/axisPosition）会被
#      openpyxl 静默丢弃，而“数据条最短/最长长度”是本细则的判定项；
#   ② 交付物可能是网页图表组件（WPS web-chart），其定义在
#      xl/webExtensions/*.xml，openpyxl 完全看不到；
#   ③ 三维视图参数（俯视角/旋转角/直角坐标轴）与系列缓存需要按原样读取。
# 因此本模块直接以 zipfile + xml.etree.ElementTree 读取 OOXML 包。

import os
import re
import sys
import json
import zipfile
import posixpath
import traceback
import xml.etree.ElementTree as ET

SCRIPT_ID = "083"

# 维度2 加分项合计 72 分（8×5 + 10×3 + 2×1），减分项合计 48 分（13×3 + 9×1）。
MAX_SCORE_TOTAL = 72
MAX_PENALTY_TOTAL = 48

# 每个 results.append 对应一项，顺序与细则文本逐条对齐。
# 加分项取正值，减分项取负值；负值项“命中”表示该处破坏性修改确实发生、要扣分。
_MAX_DELTAS_ORDER = [
    # ---- 加分项 1~8（各 5 分）----
    5, 5, 5, 5, 5, 5, 5, 5,
    # ---- 加分项 9~18（各 3 分）----
    3, 3, 3, 3, 3, 3, 3, 3, 3, 3,
    # ---- 加分项 19~20（各 1 分）----
    1, 1,
    # ---- 减分项 1~8（各 -3 分）----
    -3, -3, -3, -3, -3, -3, -3, -3,
    # ---- 减分项 9（-1 分）----
    -1,
    # ---- 减分项 10~13（各 -3 分）----
    -3, -3, -3, -3,
    # ---- 减分项 14~17（各 -1 分）----
    -1, -1, -1, -1,
    # ---- 减分项 18（-3 分）----
    -3,
    # ---- 减分项 19~22（各 -1 分）----
    -1, -1, -1, -1,
]
assert sum(d for d in _MAX_DELTAS_ORDER if d > 0) == MAX_SCORE_TOTAL
assert -sum(d for d in _MAX_DELTAS_ORDER if d < 0) == MAX_PENALTY_TOTAL

# ===== 细则中出现的固定业务常量 =====
TITLE_TEXT = "区域出货量对比分析看板"
HEADER_TEXTS = ["年份", "华南", "华北", "中西部", "西部"]
REGION_TEXTS = ["华南", "华北", "中西部", "西部"]
YEAR_NUMS = [2020, 2021, 2022]
YEAR_TEXTS = ["2020年", "2021年", "2022年"]
DATA_ROWS = [
    [2000, 4500, 5000, 9500],
    [6000, 1000, 4900, 3900],
    [9000, 3000, 8000, 4200],
]
ALL_DATA_NUMS = [v for row in DATA_ROWS for v in row]
BAR_COLOR = "38BDF8"          # 浅蓝色数据条
HEADER_FILL = "0F172A"        # 深蓝黑标题行
WHITE = "FFFFFF"              # 白色
STRIPE_FILL = "EEF2FF"        # 浅蓝紫间隔行
YAHEI_NAMES = ("microsoft yahei", "微软雅黑", "microsoft yahei ui")
NOTE_KEYWORDS = ["说明：", "不是图片", "可编辑图表", "修改数据后",
                 "图表可更新", "制图说明", "生成说明"]

# ===== 命名空间 =====
NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_PKGREL = "http://schemas.openxmlformats.org/package/2006/relationships"
NS_C = "http://schemas.openxmlformats.org/drawingml/2006/chart"
NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main"
NS_XDR = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
NS_X14 = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"
NS_XM = "http://schemas.microsoft.com/office/excel/2006/main"
NS_WPSWE = "http://www.wps.cn/officeDocument/2018/webExtension"

EMU_PER_CM = 360000.0


def _q(ns, tag):
    """拼装带命名空间的元素名。"""
    return "{%s}%s" % (ns, tag)


def _findall_local(root, local):
    """按本地名（忽略命名空间）递归查找全部元素。

    交付物可能来自不同办公软件，同一语义元素的命名空间前缀写法不一，
    因此凡是不涉及歧义的查找统一按本地名匹配，避免因前缀差异漏判。
    """
    out = []
    if root is None:
        return out
    for el in root.iter():
        if el.tag.rpartition("}")[2] == local:
            out.append(el)
    return out


def _find_local(root, local):
    got = _findall_local(root, local)
    return got[0] if got else None


def _children_local(el, local):
    out = []
    if el is None:
        return out
    for ch in list(el):
        if ch.tag.rpartition("}")[2] == local:
            out.append(ch)
    return out


def _child_local(el, local):
    got = _children_local(el, local)
    return got[0] if got else None


def _attr(el, name, ns=None):
    if el is None:
        return None
    if ns:
        v = el.get(_q(ns, name))
        if v is not None:
            return v
    return el.get(name)


def _val_of(el):
    """读取 <xxx val="..."/> 形式的取值。"""
    if el is None:
        return None
    return el.get("val")


def _bool_val(el, default=False):
    """OOXML 布尔属性：元素存在且 val 省略视为 True。"""
    if el is None:
        return default
    v = el.get("val")
    if v is None:
        return True
    return str(v).strip().lower() in ("1", "true", "on")


def _bool_val_attr(el, name, default=False):
    """读取某个布尔属性；元素存在而属性省略时取 OOXML 默认值。"""
    if el is None:
        return default
    v = el.get(name)
    if v is None:
        return default
    return str(v).strip().lower() in ("1", "true", "on")


def _num(txt):
    if txt is None:
        return None
    s = str(txt).strip()
    if s == "":
        return None
    try:
        return float(s)
    except ValueError:
        return None


def _col_letters_to_idx(letters):
    """列字母转 1 起始的列号：A->1、F->6。"""
    n = 0
    for ch in letters.upper():
        if "A" <= ch <= "Z":
            n = n * 26 + (ord(ch) - 64)
    return n


def _split_ref(ref):
    """单元格地址拆分为 (列号, 行号)，均为 1 起始。"""
    m = re.match(r"^\$?([A-Za-z]+)\$?(\d+)$", str(ref).strip())
    if not m:
        return None, None
    return _col_letters_to_idx(m.group(1)), int(m.group(2))


def _parse_area(area):
    """区域地址（A1 或 A1:E5）转 (c1, r1, c2, r2)，均为 1 起始闭区间。"""
    s = str(area).replace("$", "").strip()
    if ":" in s:
        a, _, b = s.partition(":")
        c1, r1 = _split_ref(a)
        c2, r2 = _split_ref(b)
    else:
        c1, r1 = _split_ref(s)
        c2, r2 = c1, r1
    if None in (c1, r1, c2, r2):
        return None
    return (min(c1, c2), min(r1, r2), max(c1, c2), max(r1, r2))


def _area_cells(area):
    """区域地址展开为单元格地址列表。"""
    box = _parse_area(area)
    if not box:
        return []
    c1, r1, c2, r2 = box
    out = []
    for r in range(r1, r2 + 1):
        for c in range(c1, c2 + 1):
            out.append(_addr(c, r))
    return out


def _addr(col, row):
    """列号行号转 A1 形式地址。"""
    s = ""
    n = col
    while n > 0:
        n, rem = divmod(n - 1, 26)
        s = chr(65 + rem) + s
    return "%s%d" % (s, row)


def _split_sheet_ref(f):
    """图表引用式拆分为 (工作表名, 区域地址)。

    形如 'Sheet1'!$A$2:$E$5 / Sheet2!B2 / Sheet1!$A$3:$A$6 均可解析。
    """
    if not f:
        return None, None
    s = str(f).strip()
    if "!" not in s:
        return None, s.replace("$", "")
    name, _, area = s.rpartition("!")
    name = name.strip()
    if len(name) >= 2 and name[0] == "'" and name[-1] == "'":
        name = name[1:-1].replace("''", "'")
    return name, area.replace("$", "")


# Excel 内置 indexed 调色板中本细则会用到的少量取值（其余按未知处理）。
_INDEXED_COLORS = {
    0: "000000", 1: "FFFFFF", 2: "FF0000", 3: "00FF00", 4: "0000FF",
    5: "FFFF00", 6: "FF00FF", 7: "00FFFF", 8: "000000", 9: "FFFFFF",
    64: None, 65: None,   # 64/65 表示“自动”，无固定色值
}

# 主题色下标到 clrScheme 子元素名的映射。
# 注意：电子表格里 theme="0..3" 与 clrScheme 中 dk1/lt1/dk2/lt2 的书写顺序
# 前四项是交换过的（0=lt1、1=dk1、2=lt2、3=dk2），之后才依次 accent1..6、hlink。
_THEME_ORDER = ["lt1", "dk1", "lt2", "dk2", "accent1", "accent2", "accent3",
                "accent4", "accent5", "accent6", "hlink", "folHlink"]


def _apply_tint(rgb, tint):
    """按 OOXML 规则对色值施加 tint（近似即可，仅用于“是否浅色/白色”判断）。"""
    if not rgb or tint is None:
        return rgb
    try:
        t = float(tint)
    except (TypeError, ValueError):
        return rgb
    out = ""
    for i in range(0, 6, 2):
        v = int(rgb[i:i + 2], 16)
        if t < 0:
            v = int(round(v * (1.0 + t)))
        elif t > 0:
            v = int(round(v * (1.0 - t) + 255.0 * t))
        out += "%02X" % max(0, min(255, v))
    return out


def _color_rgb(el, theme_colors):
    """把 fgColor/color/srgbClr 等颜色元素统一解析为 6 位 RRGGBB（大写）。"""
    if el is None:
        return None
    rgb = el.get("rgb") or el.get("val")
    if rgb:
        s = str(rgb).strip().upper()
        if len(s) == 8:
            s = s[2:]
        if re.match(r"^[0-9A-F]{6}$", s):
            return _apply_tint(s, el.get("tint"))
    th = el.get("theme")
    if th is not None:
        try:
            idx = int(th)
        except ValueError:
            idx = -1
        if 0 <= idx < len(_THEME_ORDER):
            base = (theme_colors or {}).get(_THEME_ORDER[idx])
            if base:
                return _apply_tint(base, el.get("tint"))
    ix = el.get("indexed")
    if ix is not None:
        try:
            return _INDEXED_COLORS.get(int(ix))
        except ValueError:
            return None
    return None


class Package(object):
    """OOXML 包的最小读取器：按部件名取字节 / 取 XML 根 / 解关系。"""

    def __init__(self, path):
        self.path = path
        self.zf = zipfile.ZipFile(path)
        self.names = set(self.zf.namelist())
        self._xml_cache = {}

    def close(self):
        try:
            self.zf.close()
        except Exception:
            pass

    def has(self, name):
        return name in self.names

    def read(self, name):
        if name not in self.names:
            return None
        return self.zf.read(name)

    def xml(self, name):
        if name in self._xml_cache:
            return self._xml_cache[name]
        data = self.read(name)
        root = None
        if data is not None:
            try:
                root = ET.fromstring(data)
            except ET.ParseError:
                root = None
        self._xml_cache[name] = root
        return root

    def rels(self, part_name):
        """取某部件的关系表：{Id: {"type":..., "target":规范化部件名, "mode":...}}"""
        d = posixpath.dirname(part_name)
        base = posixpath.basename(part_name)
        rel_name = posixpath.join(d, "_rels", base + ".rels") if d else "_rels/%s.rels" % base
        root = self.xml(rel_name)
        out = {}
        if root is None:
            return out
        for rel in _findall_local(root, "Relationship"):
            rid = rel.get("Id")
            tgt = rel.get("Target") or ""
            mode = rel.get("TargetMode") or "Internal"
            typ = rel.get("Type") or ""
            if mode == "External":
                norm = tgt
            elif tgt.startswith("/"):
                norm = tgt[1:]
            else:
                norm = posixpath.normpath(posixpath.join(d, tgt)) if d else posixpath.normpath(tgt)
            out[rid] = {"type": typ, "target": norm, "mode": mode, "raw": tgt}
        return out


# 内置数字格式中与本细则判定相关的若干项。
_BUILTIN_NUMFMT = {
    0: "General", 1: "0", 2: "0.00", 3: "#,##0", 4: "#,##0.00",
    9: "0%", 10: "0.00%", 11: "0.00E+00", 37: "#,##0 ;(#,##0)",
    38: "#,##0 ;[Red](#,##0)", 39: "#,##0.00;(#,##0.00)",
    40: "#,##0.00;[Red](#,##0.00)", 49: "@",
}


class Styles(object):
    """xl/styles.xml 解析结果：把 cellXfs 下标还原为可判定的字体/填充/边框/对齐。"""

    def __init__(self, pkg, theme_colors):
        self.theme_colors = theme_colors
        self.numfmts = {}
        self.fonts = []
        self.fills = []
        self.borders = []
        self.cell_style_xfs = []
        self.cell_xfs = []
        root = pkg.xml("xl/styles.xml")
        if root is None:
            return
        for nf in _findall_local(_find_local(root, "numFmts"), "numFmt"):
            try:
                self.numfmts[int(nf.get("numFmtId"))] = nf.get("formatCode") or ""
            except (TypeError, ValueError):
                pass
        for fo in _children_local(_find_local(root, "fonts"), "font"):
            self.fonts.append(self._font(fo))
        for fi in _children_local(_find_local(root, "fills"), "fill"):
            self.fills.append(self._fill(fi))
        for bo in _children_local(_find_local(root, "borders"), "border"):
            self.borders.append(self._border(bo))
        for xf in _children_local(_find_local(root, "cellStyleXfs"), "xf"):
            self.cell_style_xfs.append(self._xf(xf))
        for xf in _children_local(_find_local(root, "cellXfs"), "xf"):
            self.cell_xfs.append(self._xf(xf))

    def _font(self, fo):
        name_el = _child_local(fo, "name")
        if name_el is None:
            name_el = _child_local(fo, "rFont")
        sz = _num(_val_of(_child_local(fo, "sz")))
        return {
            "name": (_val_of(name_el) or "").strip(),
            "sz": sz,
            "bold": _bool_val(_child_local(fo, "b"), False),
            "italic": _bool_val(_child_local(fo, "i"), False),
            "color": _color_rgb(_child_local(fo, "color"), self.theme_colors),
        }

    def _fill(self, fi):
        pf = _child_local(fi, "patternFill")
        if pf is None:
            return {"pattern": None, "fg": None}
        return {
            "pattern": pf.get("patternType"),
            "fg": _color_rgb(_child_local(pf, "fgColor"), self.theme_colors),
        }


    def _border(self, bo):
        out = {}
        for side in ("left", "right", "top", "bottom"):
            el = _child_local(bo, side)
            out[side] = (el.get("style") if el is not None else None)
        return out

    def _xf(self, xf):
        al = _child_local(xf, "alignment")
        return {
            "numFmtId": xf.get("numFmtId"),
            "fontId": xf.get("fontId"),
            "fillId": xf.get("fillId"),
            "borderId": xf.get("borderId"),
            "xfId": xf.get("xfId"),
            "applyAlignment": xf.get("applyAlignment"),
            "h": (al.get("horizontal") if al is not None else None),
            "v": (al.get("vertical") if al is not None else None),
            "has_align": al is not None,
        }

    # ---- 由单元格 s 下标取生效属性 ----
    def xf(self, s):
        try:
            i = int(s) if s is not None else 0
        except (TypeError, ValueError):
            i = 0
        if 0 <= i < len(self.cell_xfs):
            return self.cell_xfs[i]
        return None

    def font_of(self, s):
        x = self.xf(s)
        if not x:
            return None
        try:
            i = int(x.get("fontId") or 0)
        except ValueError:
            return None
        return self.fonts[i] if 0 <= i < len(self.fonts) else None

    def fill_of(self, s):
        x = self.xf(s)
        if not x:
            return None
        try:
            i = int(x.get("fillId") or 0)
        except ValueError:
            return None
        return self.fills[i] if 0 <= i < len(self.fills) else None

    def border_of(self, s):
        x = self.xf(s)
        if not x:
            return None
        try:
            i = int(x.get("borderId") or 0)
        except ValueError:
            return None
        return self.borders[i] if 0 <= i < len(self.borders) else None


    def align_of(self, s):
        """对齐方式：cellXfs 自带的优先，缺省时回落到其引用的 cellStyleXfs。

        与办公软件一致：单元格样式（cellStyleXfs）里的 <alignment> 会被继承，
        本工作簿的“常规”样式即带 vertical="center"，不回落会造成整片误判。
        """
        x = self.xf(s)
        if not x:
            return (None, None)
        h, v = x.get("h"), x.get("v")
        if h is None or v is None:
            try:
                i = int(x.get("xfId") or 0)
            except ValueError:
                i = -1
            if 0 <= i < len(self.cell_style_xfs):
                base = self.cell_style_xfs[i]
                if h is None:
                    h = base.get("h")
                if v is None:
                    v = base.get("v")
        return (h, v)

    def numfmt_of(self, s):
        """数字格式代码：优先自定义 numFmts，其次内置表。"""
        x = self.xf(s)
        if not x:
            return None
        try:
            fid = int(x.get("numFmtId") or 0)
        except ValueError:
            return None
        if fid in self.numfmts:
            return self.numfmts[fid]
        return _BUILTIN_NUMFMT.get(fid)


def _is_int_thousands_fmt(code):
    """判定“三位一分隔、无小数位的整数格式”。"""
    if code is None:
        return False
    c = str(code).strip()
    if c == "" or c == "@" or c.lower() == "general":
        return False
    if "%" in c or "E+" in c or "E-" in c:
        return False
    if "#,##0" not in c and "#,###" not in c:
        return False
    if re.search(r"[#0]\.[#0]", c):     # 出现小数位
        return False
    return True


class Sheet(object):
    """单个工作表：单元格取值/样式下标、行高列宽、合并区、条件格式、视图、页边距。"""

    def __init__(self, pkg, name, part, shared_strings):
        self.pkg = pkg
        self.name = name
        self.part = part
        self.root = pkg.xml(part)
        self.rels = pkg.rels(part)
        self.ss = shared_strings or []
        self.cells = {}          # 地址 -> {"text":..., "num":..., "s":..., "formula":...}
        self.styles_idx = {}     # 地址 -> s 下标（含仅有样式而无内容的单元格）
        self.row_heights = {}
        self.col_ranges = []
        self.default_row_height = None
        self.default_col_width = None
        self.merges = []
        self.freeze = False
        self.has_autofilter = False
        self.page_margins = {}
        self.databars = []       # 基础数据条规则
        self.other_cf = []       # 图标集/色阶等其他条件格式
        self.databar_ext = {}    # x14 扩展：id -> {minLength, maxLength, axisPosition}
        self.tables = []         # 表格对象
        if self.root is not None:
            self._parse()
            self._parse_tables()

    def _parse(self):
        root = self.root
        fmt = _find_local(root, "sheetFormatPr")
        if fmt is not None:
            self.default_row_height = _num(fmt.get("defaultRowHeight"))
            self.default_col_width = _num(fmt.get("defaultColWidth"))
        for col in _findall_local(_find_local(root, "cols"), "col"):
            try:
                cmin = int(col.get("min"))
                cmax = int(col.get("max"))
            except (TypeError, ValueError):
                continue
            self.col_ranges.append((cmin, cmax, _num(col.get("width"))))
        for row in _findall_local(_find_local(root, "sheetData"), "row"):
            try:
                r = int(row.get("r"))
            except (TypeError, ValueError):
                continue
            ht = _num(row.get("ht"))
            if ht is not None:
                self.row_heights[r] = ht
            for c in _children_local(row, "c"):
                self._parse_cell(c)
        for mc in _findall_local(_find_local(root, "mergeCells"), "mergeCell"):
            if mc.get("ref"):
                self.merges.append(mc.get("ref"))
        self._parse_view(root)
        self._parse_cf(root)
        pm = _find_local(root, "pageMargins")
        if pm is not None:
            for k in ("left", "right", "top", "bottom", "header", "footer"):
                self.page_margins[k] = _num(pm.get(k))


    def _parse_cell(self, c):
        """解析单个 <c>。

        与办公软件对齐的两处要点：
        ① 只有样式没有内容的单元格（形如 <c r="F1" s="18"/>、
           <c r="F1" s="18" t="n"></c>）不算“写入了内容”；
        ② 公式单元格若缓存值为空（<v></v>），文件内并未保存结果值，
           同样不算“存有数值”。
        """
        ref = c.get("r")
        if not ref:
            return
        s = c.get("s")
        if s is not None:
            self.styles_idx[ref] = s
        t = c.get("t")
        f_el = _child_local(c, "f")
        formula = (f_el.text or "") if f_el is not None else None
        v_el = _child_local(c, "v")
        v_txt = (v_el.text if v_el is not None else None)
        text = None
        num = None
        if t == "s":
            i = _num(v_txt)
            if i is not None and 0 <= int(i) < len(self.ss):
                text = self.ss[int(i)]
        elif t == "inlineStr":
            is_el = _child_local(c, "is")
            parts = [(x.text or "") for x in _findall_local(is_el, "t")]
            joined = "".join(parts)
            text = joined if joined != "" else None
        elif t == "str":
            text = v_txt if (v_txt or "") != "" else None
        elif t in ("b", "e"):
            text = v_txt if (v_txt or "") != "" else None
        else:
            num = _num(v_txt)
        if text is None and num is None and formula is None:
            return
        self.cells[ref] = {"text": text, "num": num, "s": s, "formula": formula}

    def cell(self, ref):
        return self.cells.get(ref)

    def text_at(self, ref):
        c = self.cells.get(ref)
        return None if not c else c.get("text")

    def num_at(self, ref):
        c = self.cells.get(ref)
        return None if not c else c.get("num")

    def style_at(self, ref):
        c = self.cells.get(ref)
        if c and c.get("s") is not None:
            return c.get("s")
        return self.styles_idx.get(ref)


    def _parse_view(self, root):
        for sv in _findall_local(_find_local(root, "sheetViews"), "sheetView"):
            for pane in _children_local(sv, "pane"):
                state = (pane.get("state") or "").lower()
                xs = _num(pane.get("xSplit")) or 0.0
                ys = _num(pane.get("ySplit")) or 0.0
                if state in ("frozen", "frozensplit") and (xs > 0 or ys > 0):
                    self.freeze = True
        # 工作表级自动筛选区域
        for af in _findall_local(root, "autoFilter"):
            self.has_autofilter = True
            break

    def _parse_cf(self, root):
        """条件格式：基础 dataBar 规则 + x14 扩展中的长度/坐标轴设定。"""
        for cf in _findall_local(root, "conditionalFormatting"):
            # x14 扩展里的 conditionalFormatting 用 xm:sqref 子元素，单独处理
            if cf.tag.startswith("{%s}" % NS_X14):
                continue
            sqref = cf.get("sqref") or ""
            for rule in _children_local(cf, "cfRule"):
                rtype = (rule.get("type") or "").strip()
                if rtype == "dataBar":
                    db = _child_local(rule, "dataBar")
                    cfvos = [x.get("type") for x in _children_local(db, "cfvo")]
                    ext_id = None
                    for idel in _findall_local(rule, "id"):
                        ext_id = (idel.text or "").strip().lower()
                    self.databars.append({
                        "sqref": sqref,
                        "gradient": _bool_val_attr(db, "gradient", True),
                        "color": _color_rgb(_child_local(db, "color"), None),
                        "cfvos": cfvos,
                        "ext_id": ext_id,
                    })
                elif rtype in ("iconSet", "colorScale"):
                    self.other_cf.append({"sqref": sqref, "type": rtype})
        for x14cf in _findall_local(root, "conditionalFormatting"):
            if not x14cf.tag.startswith("{%s}" % NS_X14):
                continue
            for rule in _children_local(x14cf, "cfRule"):
                if (rule.get("type") or "") != "dataBar":
                    continue
                rid = (rule.get("id") or "").strip().lower()
                db = _child_local(rule, "dataBar")
                if db is None:
                    continue
                self.databar_ext[rid] = {
                    "minLength": _num(db.get("minLength")),
                    "maxLength": _num(db.get("maxLength")),
                    "axisPosition": (db.get("axisPosition") or "").strip().lower(),
                }


    def _parse_tables(self):
        """带筛选按钮的表格对象：由工作表关系指向 xl/tables/*.xml。"""
        for rid, rel in self.rels.items():
            if not rel["type"].endswith("/table"):
                continue
            troot = self.pkg.xml(rel["target"])
            if troot is None:
                continue
            self.tables.append({
                "name": troot.get("name") or troot.get("displayName") or "",
                "ref": troot.get("ref") or "",
                "autofilter": _find_local(troot, "autoFilter") is not None,
                "header": troot.get("headerRowCount"),
            })

    def col_width(self, col):
        """列的实际生效宽度：由连续列区间统一设定的宽度同样计入。"""
        w = None
        for cmin, cmax, width in self.col_ranges:
            if cmin <= col <= cmax and width is not None:
                w = width
        if w is None:
            w = self.default_col_width
        return w

    def row_height(self, row):
        if row in self.row_heights:
            return self.row_heights[row]
        return self.default_row_height

    def merge_boxes(self):
        out = []
        for ref in self.merges:
            box = _parse_area(ref)
            if box:
                out.append(box)
        return out

    def drawing_parts(self):
        out = []
        for rid, rel in self.rels.items():
            if rel["type"].endswith("/drawing"):
                out.append(rel["target"])
        return out


class Workbook(object):
    """工作簿：工作表清单、共享字符串、样式、主题色、批注文本。"""

    def __init__(self, pkg):
        self.pkg = pkg
        self.theme_colors = self._theme()
        self.shared_strings = self._shared_strings()
        self.styles = Styles(pkg, self.theme_colors)
        self.sheets = []
        self._load_sheets()
        self.comment_texts = self._comment_texts()

    def _theme(self):
        out = {}
        for name in pkg_theme_candidates(self.pkg):
            root = self.pkg.xml(name)
            scheme = _find_local(root, "clrScheme")
            if scheme is None:
                continue
            for ch in list(scheme):
                key = ch.tag.rpartition("}")[2]
                srgb = _child_local(ch, "srgbClr")
                sysc = _child_local(ch, "sysClr")
                if srgb is not None and srgb.get("val"):
                    out[key] = srgb.get("val").upper()
                elif sysc is not None and sysc.get("lastClr"):
                    out[key] = sysc.get("lastClr").upper()
            if out:
                break
        return out

    def _shared_strings(self):
        root = self.pkg.xml("xl/sharedStrings.xml")
        out = []
        if root is None:
            return out
        for si in _children_local(root, "si"):
            parts = [(t.text or "") for t in _findall_local(si, "t")]
            out.append("".join(parts))
        return out

    def _load_sheets(self):
        root = self.pkg.xml("xl/workbook.xml")
        rels = self.pkg.rels("xl/workbook.xml")
        if root is None:
            return
        for sh in _findall_local(_find_local(root, "sheets"), "sheet"):
            name = sh.get("name") or ""
            rid = _attr(sh, "id", NS_REL)
            rel = rels.get(rid)
            part = rel["target"] if rel else None
            if not part or not self.pkg.has(part):
                continue
            self.sheets.append(Sheet(self.pkg, name, part, self.shared_strings))

    def sheet_by_name(self, name):
        for s in self.sheets:
            if s.name == name:
                return s
        return None

    def _comment_texts(self):
        out = []
        for n in sorted(self.pkg.names):
            if re.match(r"^xl/comments\d*\.xml$", n) or "threadedComment" in n:
                root = self.pkg.xml(n)
                for t in _findall_local(root, "t"):
                    if (t.text or "").strip():
                        out.append(t.text)
        return out


def pkg_theme_candidates(pkg):
    return [n for n in sorted(pkg.names) if re.match(r"^xl/theme/theme\d*\.xml$", n)]


def _anchor_pos(anchor):
    """取锚点起止位置（列/行均转为 1 起始）。"""
    def pos(tag):
        el = _child_local(anchor, tag)
        if el is None:
            return (None, None, None, None)

        def sub(name):
            # 说明：ElementTree 的元素在无子节点时布尔值为假，此处必须显式判 None
            e = _child_local(el, name)
            return _num(e.text) if e is not None else None

        c = sub("col")
        co = sub("colOff") or 0.0
        r = sub("row")
        ro = sub("rowOff") or 0.0
        return (None if c is None else int(c) + 1, co,
                None if r is None else int(r) + 1, ro)
    f_col, f_coff, f_row, f_roff = pos("from")
    t_col, t_coff, t_row, t_roff = pos("to")
    return {"from_col": f_col, "from_col_off": f_coff,
            "from_row": f_row, "from_row_off": f_roff,
            "to_col": t_col, "to_col_off": t_coff,
            "to_row": t_row, "to_row_off": t_roff}


def _anchor_size_cm(anchor, sheet, pos):
    """取显示宽高（厘米）。

    优先使用随文件保存的显式尺寸（oneCellAnchor 的 ext 或图形自身的 a:ext）；
    若两者都没有，则按锚定跨越的列宽行高折算，作为近似值。
    """
    ext = None
    for e in _findall_local(anchor, "ext"):
        if e.get("cx") and e.get("cy"):
            ext = e
            break
    if ext is not None:
        cx = _num(ext.get("cx"))
        cy = _num(ext.get("cy"))
        if cx and cy:
            return (cx / EMU_PER_CM, cy / EMU_PER_CM)
    if not sheet or pos["from_col"] is None or pos["to_col"] is None:
        return (None, None)
    w = 0.0
    for c in range(pos["from_col"], (pos["to_col"] or pos["from_col"]) + 1):
        cw = sheet.col_width(c) or 8.43
        w += cw * 0.1804          # 1 字符宽 ≈ 7 像素 ≈ 0.1804cm
    h = 0.0
    for r in range(pos["from_row"], (pos["to_row"] or pos["from_row"]) + 1):
        rh = sheet.row_height(r) or 14.25
        h += rh * 2.54 / 72.0     # 磅转厘米
    off_w = ((pos["to_col_off"] or 0.0) - (pos["from_col_off"] or 0.0)) / EMU_PER_CM
    off_h = ((pos["to_row_off"] or 0.0) - (pos["from_row_off"] or 0.0)) / EMU_PER_CM
    return (max(0.0, w + off_w - (sheet.col_width(pos["to_col"]) or 8.43) * 0.1804),
            max(0.0, h + off_h - (sheet.row_height(pos["to_row"]) or 14.25) * 2.54 / 72.0))


def _shape_texts(sp):
    """取形状内的段落文本。

    返回 (顺序文本, 逆序文本, 是否竖排)。竖排文本框允许逐字分段，
    其字符顺序需按竖排阅读方向判定，故同时给出正序与逆序两种拼接结果。
    """
    body = _find_local(sp, "txBody")
    if body is None:
        return ("", "", False)
    body_pr = _child_local(body, "bodyPr")
    vert = (_attr(body_pr, "vert") or "").strip().lower()
    xfrm = _find_local(sp, "xfrm")
    rot = _num(_attr(xfrm, "rot")) or 0.0
    is_vertical = ("vert" in vert and vert != "horz") or abs(abs(rot) - 16200000) < 60000 \
        or abs(abs(rot) - 5400000) < 60000
    paras = []
    for p in _children_local(body, "p"):
        paras.append("".join((t.text or "") for t in _findall_local(p, "t")))
    fwd = "".join(paras)
    rev = "".join(reversed(paras))
    return (fwd, rev, is_vertical)


def _shape_font_pts(sp):
    """形状内出现过的字号（磅）。OOXML 中 sz 以百分之一磅计。"""
    out = []
    for rpr in _findall_local(sp, "rPr"):
        v = _num(rpr.get("sz"))
        if v:
            out.append(v / 100.0)
    for rpr in _findall_local(sp, "defRPr"):
        v = _num(rpr.get("sz"))
        if v:
            out.append(v / 100.0)
    return out


def _line_width_pt(ln_el):
    """线宽（磅）。w 属性以 EMU 计，1 磅 = 12700 EMU；缺省视为默认细线。"""
    if ln_el is None:
        return 0.0
    w = _num(ln_el.get("w"))
    if w is None:
        return 0.75
    return w / 12700.0


def collect_drawing_objects(pkg, sheet):
    """遍历工作表挂载的绘图部件，归集图表对象、网页图表组件与文本框。"""
    charts = []      # 原生图表：{"sheet","pos","w_cm","h_cm","part"}
    webcharts = []   # 网页图表组件：{"sheet","pos","w_cm","h_cm","part","ln_pt","img"}
    textboxes = []   # 文本框：{"sheet","pos","fwd","rev","vertical","sizes"}
    pics = []        # 其余图片对象
    for dpart in sheet.drawing_parts():
        droot = pkg.xml(dpart)
        if droot is None:
            continue
        drels = pkg.rels(dpart)
        for anchor in list(droot):
            local = anchor.tag.rpartition("}")[2]
            if local not in ("twoCellAnchor", "oneCellAnchor", "absoluteAnchor"):
                continue
            pos = _anchor_pos(anchor)
            w_cm, h_cm = _anchor_size_cm(anchor, sheet, pos)
            gf = _child_local(anchor, "graphicFrame")
            if gf is not None:
                cel = _find_local(gf, "chart")
                if cel is not None:
                    rid = _attr(cel, "id", NS_REL)
                    rel = drels.get(rid)
                    charts.append({"sheet": sheet.name, "pos": pos,
                                   "w_cm": w_cm, "h_cm": h_cm,
                                   "part": rel["target"] if rel else None})
                    continue
            pic = _child_local(anchor, "pic")
            if pic is not None:
                sppr = _child_local(pic, "spPr")
                ln_pt = _line_width_pt(_child_local(sppr, "ln"))
                ref = None
                for wer in _findall_local(pic, "webExtensionRef"):
                    ref = _attr(wer, "id", NS_REL)
                    break
                blip = _find_local(pic, "blip")
                img_rid = _attr(blip, "embed", NS_REL) if blip is not None else None
                item = {"sheet": sheet.name, "pos": pos, "w_cm": w_cm, "h_cm": h_cm,
                        "ln_pt": ln_pt,
                        "img": (drels.get(img_rid) or {}).get("target") if img_rid else None}
                if ref is not None:
                    rel = drels.get(ref)
                    item["part"] = rel["target"] if rel else None
                    webcharts.append(item)
                else:
                    pics.append(item)
                continue
            sp = _child_local(anchor, "sp")
            if sp is not None:
                fwd, rev, vertical = _shape_texts(sp)
                if fwd.strip() != "":
                    textboxes.append({"sheet": sheet.name, "pos": pos,
                                      "fwd": fwd, "rev": rev,
                                      "vertical": vertical,
                                      "sizes": _shape_font_pts(sp)})
    return charts, webcharts, textboxes, pics


_SCHEME_ALIAS = {"bg1": "lt1", "tx1": "dk1", "bg2": "lt2", "tx2": "dk2"}


def _ref_of(holder):
    """从 cat/val/tx 容器中取出引用信息。

    返回 {"kind": "strRef"|"numRef"|"multiLvlStrRef"|"lit"|None,
          "f": 引用式, "cache": [已随文件保存的取值], "has_cache": bool}
    """
    if holder is None:
        return {"kind": None, "f": None, "cache": [], "has_cache": False}
    for kind in ("strRef", "numRef", "multiLvlStrRef"):
        ref = _child_local(holder, kind)
        if ref is None:
            continue
        f_el = _child_local(ref, "f")
        f = (f_el.text or "").strip() if f_el is not None else None
        cache = []
        has_cache = False
        for cn in ("strCache", "numCache", "multiLvlStrCache"):
            cel = _child_local(ref, cn)
            if cel is None:
                continue
            has_cache = True
            for pt in _findall_local(cel, "pt"):
                v = _child_local(pt, "v")
                cache.append((v.text or "") if v is not None else "")
        return {"kind": kind, "f": f, "cache": cache, "has_cache": has_cache}
    for kind in ("strLit", "numLit"):
        lit = _child_local(holder, kind)
        if lit is None:
            continue
        cache = []
        for pt in _findall_local(lit, "pt"):
            v = _child_local(pt, "v")
            cache.append((v.text or "") if v is not None else "")
        return {"kind": "lit", "f": None, "cache": cache, "has_cache": True}
    v = _child_local(holder, "v")
    if v is not None:
        return {"kind": "lit", "f": None, "cache": [(v.text or "")], "has_cache": True}
    return {"kind": None, "f": None, "cache": [], "has_cache": False}


def _axis_title_text(ax):
    t = _child_local(ax, "title")
    if t is None:
        return None
    txt = "".join((x.text or "") for x in _findall_local(t, "t"))
    return txt.strip() or None


def _shape_fill_rgb(sppr, theme_colors):
    sf = _child_local(sppr, "solidFill")
    if sf is None:
        return None
    for tag in ("srgbClr", "schemeClr", "sysClr"):
        el = _child_local(sf, tag)
        if el is not None:
            if tag == "schemeClr":
                key = _SCHEME_ALIAS.get(el.get("val"), el.get("val"))
                base = (theme_colors or {}).get(key)
                return _apply_tint(base, None) if base else None
            return _color_rgb(el, theme_colors)
    return None


def _blank_desc():
    """归一化图表描述子：原生图表与网页图表组件都折算到这一组字段上，
    使 20 条加分项只写一套判定逻辑。"""
    return {
        "kind": None, "sheet": None, "pos": None, "w_cm": None, "h_cm": None,
        "is_bar3d": False, "bar_vertical": False, "clustered": False,
        "has_cat_ax": False, "has_val_ax": False, "has_ser_ax": False,
        "rot_x": None, "rot_y": None, "r_ang_ax": False,
        "series": [],                 # [{"name","fill","vals","vals_cached"}]
        "cat_labels": [], "cat_from_str_ref": False, "cat_cached": False,
        "ref_sheets": set(), "ref_areas": [], "all_by_ref": False,
        "has_data_cache": False, "has_render_cache": False,
        "title_text": None, "title_top": False,
        "val_min": None, "val_max": None, "val_unit": None,
        "val_explicit": False, "val_gridlines": False,
        "val_ax_title": None, "cat_ax_title": None, "ser_ax_title": None,
        "dlbl_show_val": False, "dlbl_font_pt": None, "dlbl_integer": True,
        "legend_bottom": False, "legend_entries": [],
        "area_fill_ok": False, "area_border_pt": 0.0,
        "texts": [],
    }


def build_native_desc(pkg, wb, chart_obj):
    """把 OOXML 原生图表折算为归一化描述子。"""
    d = _blank_desc()
    d["kind"] = "native"
    d["sheet"] = chart_obj["sheet"]
    d["pos"] = chart_obj["pos"]
    d["w_cm"] = chart_obj["w_cm"]
    d["h_cm"] = chart_obj["h_cm"]
    root = pkg.xml(chart_obj["part"]) if chart_obj.get("part") else None
    if root is None:
        return d
    theme = wb.theme_colors
    # ---- 图表类型与三维视图 ----
    bar3d = _find_local(root, "bar3DChart")
    d["is_bar3d"] = bar3d is not None
    if bar3d is not None:
        d["bar_vertical"] = (_val_of(_child_local(bar3d, "barDir")) or "col") == "col"
        d["clustered"] = (_val_of(_child_local(bar3d, "grouping")) or "") == "clustered"
    v3 = _find_local(root, "view3D")
    if v3 is not None:
        d["rot_x"] = _num(_val_of(_child_local(v3, "rotX")))
        d["rot_y"] = _num(_val_of(_child_local(v3, "rotY")))
        d["r_ang_ax"] = _bool_val(_child_local(v3, "rAngAx"), False)


    # ---- 坐标轴 ----
    cat_ax = _find_local(root, "catAx")
    val_ax = _find_local(root, "valAx")
    ser_ax = _find_local(root, "serAx")
    d["has_cat_ax"] = cat_ax is not None
    d["has_val_ax"] = val_ax is not None
    d["has_ser_ax"] = ser_ax is not None
    d["cat_ax_title"] = _axis_title_text(cat_ax)
    d["val_ax_title"] = _axis_title_text(val_ax)
    d["ser_ax_title"] = _axis_title_text(ser_ax)
    if val_ax is not None:
        sc = _child_local(val_ax, "scaling")
        d["val_min"] = _num(_val_of(_child_local(sc, "min")))
        d["val_max"] = _num(_val_of(_child_local(sc, "max")))
        d["val_unit"] = _num(_val_of(_child_local(val_ax, "majorUnit")))
        d["val_explicit"] = (d["val_min"] is not None and d["val_max"] is not None
                             and d["val_unit"] is not None)
        d["val_gridlines"] = _child_local(val_ax, "majorGridlines") is not None
    # ---- 标题 ----
    ttl = _find_local(root, "title")
    if ttl is not None and not _bool_val(_find_local(root, "autoTitleDeleted"), False):
        d["title_text"] = ("".join((t.text or "") for t in _findall_local(ttl, "t"))).strip() or None
        overlay = _bool_val(_child_local(ttl, "overlay"), False)
        lay = _child_local(ttl, "layout")
        y = None
        if lay is not None:
            for yel in _findall_local(lay, "y"):
                y = _num(yel.get("val"))
        # 未显式给出手动布局时，办公软件默认把标题排在图表顶部。
        d["title_top"] = (not overlay) and (y is None or y <= 0.25)
    # ---- 数据标签 ----
    dl = _find_local(root, "dLbls")
    if dl is not None:
        d["dlbl_show_val"] = _bool_val(_child_local(dl, "showVal"), False)
        for rpr in _findall_local(dl, "defRPr"):
            sz = _num(rpr.get("sz"))
            if sz:
                d["dlbl_font_pt"] = sz / 100.0
        nf = _child_local(dl, "numFmt")
        code = nf.get("formatCode") if nf is not None else None
        if code and (re.search(r"[#0]\.[#0]", code) or "E+" in code):
            d["dlbl_integer"] = False
    if d["dlbl_font_pt"] is None:
        d["dlbl_font_pt"] = 10.0      # 未显式设定时取办公软件默认字号
    # ---- 图例 ----
    lg = _find_local(root, "legend")
    if lg is not None:
        d["legend_bottom"] = (_val_of(_child_local(lg, "legendPos")) or "") == "b"
    # ---- 系列、分类标签、引用信息、外观 ----
    _fill_native_series(pkg, wb, root, d)
    build_native_desc_finish(pkg, wb, root, d)
    return d


def resolve_ref_cells(wb, f):
    """把图表引用式解析到工作表，取出文件内已保存的取值。

    仅读取随文件保存下来的内容：文本单元格取文本，数值单元格取数值；
    公式单元格若没有缓存结果值，则该位置视为“文件内无取值”。
    """
    name, area = _split_sheet_ref(f)
    if not area:
        return []
    sheet = wb.sheet_by_name(name) if name else (wb.sheets[0] if wb.sheets else None)
    if sheet is None:
        return []
    out = []
    for ref in _area_cells(area):
        c = sheet.cell(ref)
        if not c:
            out.append(None)
        elif c.get("text") is not None:
            out.append(c["text"])
        elif c.get("num") is not None:
            out.append(c["num"])
        else:
            out.append(None)
    return out


def _as_text(v):
    if v is None:
        return ""
    if isinstance(v, float):
        return str(int(v)) if abs(v - int(v)) < 1e-9 else str(v)
    return str(v)


def _cache_nums(cache):
    out = []
    for x in cache:
        n = _num(x)
        if n is not None:
            out.append(n)
    return out


def _fill_native_series(pkg, wb, root, d):
    """归集原生图表的系列、分类标签与引用信息。"""
    theme = wb.theme_colors
    sers = _findall_local(root, "ser")
    cat_ref_seen = None
    by_ref = True
    for ser in sers:
        tx = _ref_of(_child_local(ser, "tx"))
        cat = _ref_of(_child_local(ser, "cat"))
        val = _ref_of(_child_local(ser, "val"))
        for r in (tx, cat, val):
            if r["kind"] in (None, "lit"):
                by_ref = False
            if r["f"]:
                sname, sarea = _split_sheet_ref(r["f"])
                if sname:
                    d["ref_sheets"].add(sname)
                if sarea:
                    d["ref_areas"].append((sname, sarea))
        name = ""
        if tx["cache"]:
            name = _as_text(tx["cache"][0])
        if not name and tx["f"]:
            got = [x for x in resolve_ref_cells(wb, tx["f"]) if x is not None]
            name = _as_text(got[0]) if got else ""
        vals_cached = _cache_nums(val["cache"]) if val["has_cache"] else []
        d["series"].append({
            "name": name,
            "fill": _shape_fill_rgb(_child_local(ser, "spPr"), theme),
            "vals_cached": vals_cached,
            "val_ref": val["f"],
        })
        if cat_ref_seen is None and (cat["kind"] or cat["cache"]):
            cat_ref_seen = cat
    d["all_by_ref"] = by_ref and len(sers) > 0


    # ---- 分类标签：优先图表缓存，其次引用所指单元格已保存的取值 ----
    if cat_ref_seen is not None:
        d["cat_from_str_ref"] = cat_ref_seen["kind"] in ("strRef", "multiLvlStrRef")
        if cat_ref_seen["has_cache"] and cat_ref_seen["cache"]:
            d["cat_labels"] = [_as_text(x) for x in cat_ref_seen["cache"]]
            d["cat_cached"] = True
        elif cat_ref_seen["f"]:
            got = resolve_ref_cells(wb, cat_ref_seen["f"])
            d["cat_labels"] = [_as_text(x) for x in got if x is not None]
    d["legend_entries"] = [s["name"] for s in d["series"]]
    # 完整数据缓存：3 个系列各 4 个数值 + 4 个分类名称都已随文件存下
    d["has_data_cache"] = (
        len(d["series"]) == 3
        and all(len(s["vals_cached"]) == 4 for s in d["series"])
        and d["cat_cached"] and len(d["cat_labels"]) == 4
    )
    return d


def build_native_desc_finish(pkg, wb, root, d):
    """图表区外观与文本汇总。"""
    sppr = None
    for ch in list(root):
        if ch.tag.rpartition("}")[2] == "spPr":
            sppr = ch
            break
    fill_ok = True
    if sppr is not None:
        if _child_local(sppr, "gradFill") is not None:
            fill_ok = False
        else:
            rgb = _shape_fill_rgb(sppr, wb.theme_colors)
            if _child_local(sppr, "noFill") is not None:
                fill_ok = True
            elif rgb is None:
                fill_ok = True
            else:
                fill_ok = (rgb.upper() == WHITE)
        d["area_border_pt"] = _line_width_pt(_child_local(sppr, "ln"))
    else:
        # 未写图表区外观时按办公软件默认处理：白色底、细边框。
        d["area_border_pt"] = 0.75
    d["area_fill_ok"] = fill_ok
    texts = []
    for k in ("title_text", "val_ax_title", "cat_ax_title", "ser_ax_title"):
        if d.get(k):
            texts.append(d[k])
    texts.extend([x for x in d["cat_labels"] if x])
    texts.extend([s["name"] for s in d["series"] if s["name"]])
    d["texts"] = texts
    return d


def _jget(obj, *path, **kw):
    """按路径安全取 JSON 字段。"""
    default = kw.get("default")
    cur = obj
    for k in path:
        if not isinstance(cur, dict) or k not in cur:
            return default
        cur = cur[k]
    return cur if cur is not None else default


def _webchart_grid(dr):
    """把网页图表组件随文件保存的单元格绑定信息还原成二维取值表。

    <dataRange> 下的 context 给出绑定的单元格区域，其后按顺序保存
    该区域内每一行（或每一列）的取值。用区域自身的行列数判定分组方向。
    """
    ctx_el = _child_local(dr, "context")
    context = (ctx_el.text or "").strip() if ctx_el is not None else ""
    sheet_name, area = _split_sheet_ref(context)
    box = _parse_area(area) if area else None
    groups = []
    for cells in _children_local(dr, "cells"):
        vals = []
        for cell in _children_local(cells, "cell"):
            v = _child_local(cell, "value")
            vals.append((v.text or "") if v is not None else "")
        groups.append(vals)
    if not groups or not box:
        return sheet_name, area, box, []
    nrow = box[3] - box[1] + 1
    ncol = box[2] - box[0] + 1
    if len(groups) == nrow and (len(groups[0]) == ncol or nrow != ncol):
        grid = groups
    elif len(groups) == ncol:
        grid = [list(col) for col in zip(*groups)]
    else:
        grid = groups
    return sheet_name, area, box, grid


# 网页图表组件的轴语义映射：三维柱形图下 x 承载分类、y 承载系列（深度）、z 承载数值。
_WEB_AXIS_ROLE = {"category": "xAxis", "series": "yAxis", "value": "zAxis"}


def build_webchart_desc(pkg, wb, obj):
    """把网页图表组件折算为同一套归一化描述子。

    组件内的 demoData、extStyle 以及样式里 xAxis/yAxis 的 data 数组都是
    模板占位内容（系列1/系列2/类别1……），不代表实际承载的数据，一律忽略；
    实际数据只认随文件保存的单元格区域绑定信息与其取值缓存。
    """
    d = _blank_desc()
    d["kind"] = "webchart"
    d["sheet"] = obj["sheet"]
    d["pos"] = obj["pos"]
    d["w_cm"] = obj["w_cm"]
    d["h_cm"] = obj["h_cm"]
    d["area_border_pt"] = obj.get("ln_pt") or 0.0
    root = pkg.xml(obj["part"]) if obj.get("part") else None
    if root is None:
        return d
    src = _find_local(root, "extSource")
    if src is None or (src.get("id") or "").lower() != "webchart":
        return d
    props = {}
    for p in _findall_local(root, "property"):
        k = p.get("key") or p.get("name")
        if k:
            props[k] = p.get("value") or ""
    try:
        style = json.loads(props.get("style") or "{}")
    except ValueError:
        style = {}
    if not isinstance(style, dict):
        style = {}
    ctype = (props.get("type") or "").lower()
    # ---- 图表类型与三维视图 ----
    d["is_bar3d"] = ("3d" in ctype) and ("column" in ctype or "bar" in ctype or "col" in ctype)
    d["bar_vertical"] = ("column" in ctype) or ("col" in ctype and "bar" not in ctype)
    d["clustered"] = not any(x in ctype for x in ("stack", "percent"))
    vc = _jget(style, "grid3D", "viewControl", default={})
    if isinstance(vc, dict):
        d["rot_x"] = _num(vc.get("alpha"))
        d["rot_y"] = _num(vc.get("beta"))
        d["r_ang_ax"] = bool(vc.get("rAngAx"))
    # ---- 三条坐标轴：三维柱形图下 x 承载分类、y 承载系列、z 承载数值 ----
    roles = dict(_WEB_AXIS_ROLE)
    if "zAxis" not in style:
        roles = {"category": "xAxis", "series": None, "value": "yAxis"}
    _webchart_axes(style, roles, d)
    _webchart_look(style, props, root, d)
    _webchart_data(root, style, d)
    return d


def _web_axis_title(ax):
    """网页图表组件的轴标题：仅在显示开关为真且文本非空时才算存在。"""
    if not isinstance(ax, dict):
        return None
    nm = ax.get("name")
    if not isinstance(nm, dict):
        return None
    if nm.get("show") in (False, 0, "0", "false"):
        return None
    txt = (nm.get("text") or "").strip()
    return txt or None


def _webchart_axes(style, roles, d):
    """坐标轴的存在性、轴标题、数值轴刻度与网格线。"""
    cat_ax = style.get(roles["category"]) if roles["category"] else None
    ser_ax = style.get(roles["series"]) if roles["series"] else None
    val_ax = style.get(roles["value"]) if roles["value"] else None
    d["has_cat_ax"] = isinstance(cat_ax, dict)
    d["has_ser_ax"] = isinstance(ser_ax, dict)
    d["has_val_ax"] = isinstance(val_ax, dict)
    d["cat_ax_title"] = _web_axis_title(cat_ax)
    d["ser_ax_title"] = _web_axis_title(ser_ax)
    d["val_ax_title"] = _web_axis_title(val_ax)
    # 数值轴的上下限与主要刻度写在样式顶层
    d["val_min"] = _num(style.get("min"))
    d["val_max"] = _num(style.get("max"))
    d["val_unit"] = _num(style.get("interval"))
    d["val_explicit"] = (d["val_min"] is not None and d["val_max"] is not None
                         and d["val_unit"] is not None)
    show = _jget(val_ax if isinstance(val_ax, dict) else {}, "splitLine", "show")
    d["val_gridlines"] = show not in (False, 0, "0", "false", None)


def _webchart_look(style, props, root, d):
    """标题、数据标签、图例、系列填充色与图表区外观。"""
    ttl = style.get("title")
    if isinstance(ttl, dict) and ttl.get("show") not in (False, 0, "0", "false"):
        d["title_text"] = (ttl.get("text") or "").strip() or None
        d["title_top"] = str(ttl.get("position") or "").lower().startswith("top")
    lbl = style.get("label")
    if isinstance(lbl, dict):
        d["dlbl_show_val"] = lbl.get("show") not in (False, 0, "0", "false", None)
        d["dlbl_font_pt"] = _num(_jget(lbl, "textStyle", "fontSize"))
        fmt = str(lbl.get("formatter") or "")
        if re.search(r"[#0]\.[#0]", fmt) or "e+" in fmt.lower():
            d["dlbl_integer"] = False
    lg = style.get("legend")
    if isinstance(lg, dict) and lg.get("show") not in (False, 0, "0", "false"):
        d["legend_bottom"] = str(lg.get("position") or "").lower().startswith("bottom")
    bg = style.get("backgroundColor")
    rgb = (bg.get("rgb") if isinstance(bg, dict) else bg) or ""
    rgb = str(rgb).lstrip("#").upper()
    d["area_fill_ok"] = (rgb == "" or rgb == WHITE or rgb.startswith("RGBA(255,255,255"))
    if _find_local(root, "snapshot") is not None:
        d["has_render_cache"] = True


def _webchart_data(root, style, d):
    """按随文件保存的单元格区域绑定信息还原系列、分类名称与取值。"""
    dr = _find_local(root, "dataRange")
    if dr is None:
        return d
    sheet_name, area, box, grid = _webchart_grid(dr)
    if sheet_name:
        d["ref_sheets"].add(sheet_name)
    if area:
        d["ref_areas"].append((sheet_name, area))
    # 取值缓存中出现小数写法时，数据标签不能算整数
    for cells in _children_local(dr, "cells"):
        codes = [(_child_local(cells, "formatCode").text or "")
                 if _child_local(cells, "formatCode") is not None else ""]
        for cell in _children_local(cells, "cell"):
            codes.append(_attr(cell, "formatCode", NS_WPSWE) or "")
        for code in codes:
            if re.search(r"[#0]\.[#0]", code) or "E+" in code.upper():
                d["dlbl_integer"] = False
    if not grid or len(grid) < 2:
        return d
    series_dir = str(style.get("seriesDir") or "row").lower()
    if series_dir.startswith("col"):
        grid = [list(col) for col in zip(*grid)]
    cats = [_as_text(x).strip() for x in grid[0][1:]]
    colors = style.get("seriesThemeColor")
    colors = colors if isinstance(colors, list) else []
    for i, row in enumerate(grid[1:]):
        name = _as_text(row[0]).strip() if row else ""
        vals = []
        for x in row[1:]:
            n = _num(x)
            if n is not None:
                vals.append(n)
                if abs(n - int(n)) > 1e-9:
                    d["dlbl_integer"] = False
        fill = str(colors[i]).lstrip("#").upper() if i < len(colors) else None
        d["series"].append({"name": name, "fill": fill,
                            "vals_cached": vals, "val_ref": area})
    d["cat_labels"] = [c for c in cats if c != ""]
    d["cat_cached"] = len(d["cat_labels"]) > 0
    # 分类名称取自单元格且保存为文本，非序号占位
    d["cat_from_str_ref"] = (len(d["cat_labels"]) > 0
                             and all(_num(c) is None for c in d["cat_labels"]))
    d["legend_entries"] = [s["name"] for s in d["series"]]
    d["all_by_ref"] = bool(area) and len(d["series"]) > 0
    d["has_data_cache"] = (len(d["series"]) == 3
                           and all(len(s["vals_cached"]) == 4 for s in d["series"])
                           and len(d["cat_labels"]) == 4)
    texts = []
    for k in ("title_text", "val_ax_title", "cat_ax_title", "ser_ax_title"):
        if d.get(k):
            texts.append(d[k])
    texts.extend([x for x in d["cat_labels"] if x])
    texts.extend([s["name"] for s in d["series"] if s["name"]])
    d["texts"] = texts
    return d


_GARBLED_CHARS = ("�", "?", "？", "□", "▯", "■", "锟", "�")


def _has_garbled(texts):
    """文本中出现问号、方框、替换字符或成串私用区字符即视为不能正常显示。"""
    for t in texts:
        if not t:
            continue
        for ch in _GARBLED_CHARS:
            if ch in t:
                return True
        if re.search(r"[-]{2,}", t):
            return True
    return False


def _has_cjk(texts):
    for t in texts:
        if t and re.search(r"[一-鿿]", t):
            return True
    return False


def _inside(inner, outer):
    """闭区间矩形包含判定。"""
    if not inner or not outer:
        return False
    return (inner[0] >= outer[0] and inner[1] >= outer[1]
            and inner[2] <= outer[2] and inner[3] <= outer[3])


def _boxes_hit(box, outer):
    """两个闭区间矩形是否相交。"""
    if not box or not outer:
        return False
    return not (box[2] < outer[0] or box[0] > outer[2]
                or box[3] < outer[1] or box[1] > outer[3])


def _norm_year(s):
    """年份文本归一化：允许写作“2020”或“2020年”。"""
    return re.sub(r"\s|年", "", _as_text(s))


def _is_pale(rgb):
    """红、绿、蓝三通道同时大于 220（白或浅灰）。"""
    if not rgb or len(rgb) < 6:
        return False
    try:
        r = int(rgb[-6:-4], 16)
        g = int(rgb[-4:-2], 16)
        b = int(rgb[-2:], 16)
    except ValueError:
        return False
    return r > 220 and g > 220 and b > 220


def _find_data_sheet(wb):
    """数据表所在工作表：以 A1 写着看板标题的那张为准。"""
    for sh in wb.sheets:
        if (sh.text_at("A1") or "").strip() == TITLE_TEXT:
            return sh
    return wb.sheets[0] if wb.sheets else None


def _pick_desc(descs):
    """选定被评判的趋势图：优先三维柱形图，其次系列数最多的那个。"""
    if not descs:
        return None
    ranked = sorted(descs, key=lambda x: (1 if x["is_bar3d"] else 0,
                                          len(x["series"]),
                                          1 if x["has_data_cache"] or x["has_render_cache"] else 0),
                    reverse=True)
    return ranked[0]


def evaluate(dir_path):
    """按接口统一约定，返回结构化 dict；不 print、不改 sys.stdout、不 sys.exit。
    入参 dir_path 为“脚本所在目录的路径”，脚本自己在该目录里定位并打开被评估的
    .xlsx/.xlsm（忽略以 ~$ 开头的电子表格软件临时锁文件）。

    返回信息约定：
      status="error"（目录不存在、目录中没有交付文件、脚本自身异常）表示这份作业
      根本没评上，此时 dim1_pass 为空串（未判定）、dim1_reason 为空串、
      dim2_items 为 []、total_score 与 max_score 均为 0；
      维度一未通过时 status 仍为 "ok"，dim1_pass 为 False、dim1_reason 写错误摘要、
      dim2_items 为 []（未真正逐条评过）、total_score 为 0、max_score 为满分；
      维度一通过时 dim1_pass 为 True、dim1_reason 保持空串。
      dim2_items 每条只含 rule / max_delta / delta / hit 四个键。"""
    result = {
        "id": SCRIPT_ID,
        "file_name": "",
        "status": "ok",
        "error": None,
        "dim1_pass": "",
        "dim1_reason": "",
        "dim2_items": [],
        "total_score": 0,
        "max_score": 0,
    }
    try:
        if not dir_path or not os.path.isdir(dir_path):
            result["status"] = "error"
            result["error"] = "目录不存在：%s" % dir_path
            return result
        cands = sorted(
            f for f in os.listdir(dir_path)
            if f.lower().endswith((".xlsx", ".xlsm")) and not f.startswith("~$")
        )
        if not cands:
            # 目录里有交付物但格式不对时，交给维度一记 0，而不是当脚本错误
            others = sorted(f for f in os.listdir(dir_path)
                            if os.path.isfile(os.path.join(dir_path, f))
                            and not f.startswith("~$")
                            and not f.lower().endswith(".py"))
            if others:
                result["file_name"] = others[0]
                return _dim1_fail(result, "交付文件不是 .xlsx/.xlsm 工作簿")
            result["status"] = "error"
            result["error"] = "目录中未找到 .xlsx/.xlsm 文件：%s" % dir_path
            return result
        file_path = os.path.join(dir_path, cands[0])
        result["file_name"] = os.path.basename(file_path)
        return _evaluate_impl(file_path, result)
    except Exception as exc:
        # 脚本异常同样视为“没评上”：维度一未判定，两项分数归零。
        result["status"] = "error"
        result["error"] = "%s: %s" % (type(exc).__name__, exc)
        result["dim1_pass"] = ""
        result["dim1_reason"] = ""
        result["dim2_items"] = []
        result["total_score"] = 0
        result["max_score"] = 0
        return result


# 细则条目文本：顺序与 _MAX_DELTAS_ORDER 严格一一对应，也是 dim2_items 的 rule 字段
_RULE_TEXTS = [
    # ---- 加分项 1~8（各 5 分）----
    "趋势图为图表对象而非位图，其3个系列的数值、系列名称与分类名称均以指向工作表单元格区域的引用给出，"
    "图表内部无写死的常量数值数组；电子表格软件原生图表，或随文件保存了单元格区域绑定信息的网页图表组件，两者均计分",
    "趋势图的系列数值、系列名称与分类名称引用均指向含“区域出货量对比分析看板”标题的工作表，且引用范围落在该表A2:E5之内",
    "趋势图随文件保存了完整数据缓存或渲染结果：3个系列的12个数值与4个分类名称已随文件存下，"
    "或随文件存有图表渲染结果，文件打开即显示全部12根柱体，无需触发重算、无需重选数据源",
    "趋势图类型为三维柱形图：柱体方向垂直向上、分组方式为簇状，且同时具备分类轴、数值轴、深度轴（系列轴）三条坐标轴",
    "趋势图三维立体效果实际生效：绕水平轴俯视旋转5°至30°、绕垂直轴旋转10°至40°，启用直角坐标轴，"
    "3排柱体沿深度方向前后错开而非叠为一排",
    "趋势图深度轴（系列轴）方向承载3个年份：该轴刻度标签显示“2020”“2021”“2022”（允许写作“2020年”“2021年”“2022年”）"
    "三个年份文本，非1、2、3序号或“系列1”“系列2”之类占位文本，且该轴轴标题文本为“年份”",
    "趋势图分类轴方向承载4个销售区域：刻度标签依次显示“华南”“华北”“中西部”“西部”，"
    "且以文本型引用取自单元格，非1、2、3、4序号或“类别1”“类别2”之类占位文本",
    "趋势图数据点数为12，随文件保存的数值为：“2020”系列在四个区域上依次2000、4500、5000、9500；"
    "“2021”系列依次6000、1000、4900、3900；“2022”系列依次9000、3000、8000、4200；文件内未保存数值的不计此分",
    # ---- 加分项 9~18（各 3 分）----
    "趋势图标题位置在图表顶部，标题文本含“出货量”及“趋势”或“变化”字样（如“区域出货量变化趋势图”）",
    "趋势图数值轴刻度下限为0、上限在9500至20000之间、主要刻度为1000、2000或2500的固定间隔，"
    "三项均为显式设定而非交由软件自动缩放，且该轴显示主要网格线",
    "趋势图数值轴轴标题文本含“销售额”或“出货量”字样（如“销售额（万元）”）；"
    "该轴标题也可由紧贴数值轴左侧、文字竖排的文本框承担（竖排文本框允许逐字分段，字符顺序按竖排阅读方向判定）",
    "趋势图分类轴轴标题文本为“销售区域”；该轴标题也可由位于分类轴下方或右下方的横排文本框承担",
    "趋势图数据标签显示内容含数值，标签数字为整数（可带千位分隔符）、无小数位与科学计数法写法，标签字号不大于14磅",
    "趋势图图例位置为底部，图例项数为3且图例项文本均为年份文本，无“系列1”“系列2”“类别1”之类占位名",
    "趋势图3个系列各设独立填充色，3个颜色值互不相同，且红、绿、蓝三通道不同时大于220（非白、非浅灰）",
    "趋势图与A1:E5数据表位于同一工作表，且锚定范围与A1:E5不相交（左边界所在列不早于F列，或上边界所在行不早于第7行）",
    "趋势图显示宽度不小于15.00cm、显示高度不小于8.00cm",
    "趋势图标题、轴标题、刻度标签、图例项、数据标签内的中文均为可正常显示的汉字，无问号、方框及连续乱码字符",
    # ---- 加分项 19~20（各 1 分）----
    "趋势图图表区填充为白色或无填充，无深色底纹、无彩色渐变底纹，边框线宽不超过2.00磅",
    "工作簿内图表对象数为1，无内容重复的第二个图表",
    # ---- 减分项 1~8（各 3 分）----
    "工作簿工作表数不为1，或工作表名不为“Sheet1”",
    "工作簿内A1:E5以外的任一单元格出现“华南”“华北”“中西部”“西部”中任一文本，"
    "或出现2000、4500、5000、9500、6000、1000、4900、3900、9000、3000、8000、4200中任一数值",
    "A1:E1不为合并单元格，或数据表所在工作表中存在A1:E1以外的合并单元格",
    "A1单元格文本不为“区域出货量对比分析看板”",
    "A2、B2、C2、D2、E2单元格文本不依次为“年份”“华南”“华北”“中西部”“西部”",
    "A3、A4、A5单元格内容既不依次为数值2020、2021、2022，也不依次为文本“2020年”“2021年”“2022年”",
    "B3、C3、D3、E3不依次为2000、4500、5000、9500，或B4、C4、D4、E4不依次为6000、1000、4900、3900，"
    "或B5、C5、D5、E5不依次为9000、3000、8000、4200",
    "B3:E5未设渐变数据条条件格式，或数据条颜色不为浅蓝色（#38BDF8），或该区域被设为图标集、色阶等其他条件格式",
    # ---- 减分项 9（1 分）----
    "B3:E5数据条最短长度不为10%、最长长度不为90%（未显式指定长度时按10%与90%计），或数据条坐标轴位置被设为中点",
    # ---- 减分项 10~13（各 3 分）----
    "A1及B3:E5单元格字体不为微软雅黑（Microsoft YaHei）、10磅、常规，"
    "或A2:E2及A3:A5单元格字体不为微软雅黑（Microsoft YaHei）、10磅、加粗",
    "A2:E2填充色不为深蓝黑（#0F172A），或该行字体颜色不为白色（#FFFFFF）",
    "A3:E3、A5:E5填充色不为白色（#FFFFFF），或A4:E4填充色不为浅蓝紫（#EEF2FF）",
    "B3:E5数字格式不为三位一分隔、无小数位的整数格式，或其中任一单元格为文本型、百分比格式",
    # ---- 减分项 14~17（各 1 分）----
    "A1单元格及A2:E5区域内任一单元格水平对齐不为居中，或垂直对齐不为居中（A1:E1合并区内被覆盖的B1至E1不计）",
    "A2、A3、A4、A5左侧不为细实线边框，或A5、B5、C5、D5、E5下侧不为细实线边框",
    "A列列宽不为10.00字符、B至E列列宽不为12.00字符、F列列宽不为13.00字符"
    "（各列允许±1.50字符；列宽以该列实际生效值计，含由连续列区间统一设定的宽度）",
    "第1、2行行高不为32.0磅（0.85cm），第3行行高不为22.0磅（0.58cm），第4、5行行高不为24.0磅（0.63cm）（各行允许±1.5磅）",
    # ---- 减分项 18（3 分）----
    "工作簿内任一工作表存在带筛选按钮的表格对象或自动筛选区域",
    # ---- 减分项 19~22（各 1 分）----
    "工作簿内任一工作表设有冻结窗格",
    "工作簿内任一单元格文本、文本框文本或批注中出现“说明：”、“不是图片”、“可编辑图表”、“修改数据后”、"
    "“图表可更新”、“制图说明”、“生成说明”中任一字样",
    "数据表所在工作表第6行及以下行、F列及以后列的单元格写入了除区域名与出货量数值以外的文本或数值内容"
    "（仅设有行高、列宽或单元格样式而无内容的不计）",
    "数据表所在工作表页边距左值或右值不为0.7英寸（1.78cm），或上值或下值不为0.75英寸（1.91cm），"
    "或页眉值或页脚值不为0.3英寸（0.76cm）",
]
assert len(_RULE_TEXTS) == len(_MAX_DELTAS_ORDER)


def _dim1_fail(result, reason):
    """维度一未通过：完成度总分记 0，不再逐条罗列（未真正逐条评过）。"""
    result["dim1_pass"] = False
    result["dim1_reason"] = reason or "维度一未通过（未给出具体原因）"
    result["dim2_items"] = []
    result["total_score"] = 0
    result["max_score"] = MAX_SCORE_TOTAL
    return result


def _dim1_check(file_path):
    """维度1：可用与可修改性。只查交付格式与能否正常打开，不掺入完成度内容。

    ① 扩展名必须是 .xlsx / .xlsm；
    ② 压缩包结构完好（CRC 校验通过），否则打开时会弹出修复提示；
    ③ 工作簿必备部件齐全且 XML 可解析。
    """
    ext = os.path.splitext(file_path)[1].lower()
    if ext not in (".xlsx", ".xlsm"):
        return False, "交付文件扩展名为 %s，不是 .xlsx/.xlsm" % (ext or "（无）")
    if not zipfile.is_zipfile(file_path):
        return False, "文件不是有效的工作簿压缩包结构，无法正常打开"
    try:
        z = zipfile.ZipFile(file_path)
    except Exception as exc:
        return False, "工作簿无法打开：%s" % exc
    try:
        bad = z.testzip()
        if bad is not None:
            return False, "工作簿内部部件损坏：%s" % bad
        names = set(z.namelist())
        for need in ("[Content_Types].xml", "xl/workbook.xml"):
            if need not in names:
                return False, "工作簿缺少必备部件：%s" % need
        for n in names:
            if not (n.endswith(".xml") or n.endswith(".rels")):
                continue
            try:
                ET.fromstring(z.read(n))
            except ET.ParseError as exc:
                return False, "部件 %s 内容不合法，打开时会触发修复：%s" % (n, exc)
    finally:
        z.close()
    return True, ""


def _evaluate_impl(file_path, result):
    """维度一 + 维度二的评估主体；对内部构造 results 列表沿用统一格式，
    末尾按 _MAX_DELTAS_ORDER 反解成结构化 dim2_items 填入 result。"""
    results = []
    details = []

    # ========== 维度1：可用与可修改性 ==========
    ok, reason = _dim1_check(file_path)
    if not ok:
        return _dim1_fail(result, reason)
    result["dim1_pass"] = True
    # 约定：维度一通过时 dim1_reason 保持空串，只有未通过才写错误摘要。
    result["dim1_reason"] = ""

    pkg = Package(file_path)
    try:
        _dim2(pkg, results, details)
    finally:
        pkg.close()
    return _finish(result, results, details)


def _dim2(pkg, results, details):
    """维度2：逐条按细则顺序判定，命中即按格式追加一条结果字符串。"""

    def emit(hit, detail=""):
        """按 “+N：规则说明” / “-N：规则说明” 追加；未命中记 +0 / -0。"""
        i = len(results)
        md = _MAX_DELTAS_ORDER[i]
        head = ("%+d" % md) if hit else ("+0" if md > 0 else "-0")
        results.append("%s：%s" % (head, _RULE_TEXTS[i]))
        details.append(detail)

    wb = Workbook(pkg)
    st = wb.styles
    ds = _find_data_sheet(wb)                    # 数据表所在工作表
    ds_titled = bool(ds) and (ds.text_at("A1") or "").strip() == TITLE_TEXT
    # ---- 归集全部工作表上的图表对象、网页图表组件与文本框 ----
    charts, webcharts, textboxes = [], [], []
    for sh in wb.sheets:
        c, w, tb, _p = collect_drawing_objects(pkg, sh)
        charts.extend(c)
        webcharts.extend(w)
        textboxes.extend(tb)
    descs = [build_native_desc(pkg, wb, c) for c in charts]
    descs += [build_webchart_desc(pkg, wb, w) for w in webcharts]
    d = _pick_desc(descs)
    chart_count = len(charts) + len(webcharts)
    box_a1e5 = (1, 1, 5, 5)
    box_a2e5 = (1, 2, 5, 5)

    # ================= 加分项 =================
    # +5：趋势图为图表对象而非位图，3 个系列的数值/系列名称/分类名称均以单元格区域引用给出
    # ① 位图交付：既没有原生图表也没有网页图表组件，取不到描述子，本条不得分；
    # ② 原生图表：tx/cat/val 三者都必须是引用，出现常量数组（numLit/strLit）即失分；
    # ③ 网页图表组件：随文件保存了单元格区域绑定信息即计分，组件内的模板占位数组不计。
    r1 = bool(d) and bool(d["all_by_ref"])
    emit(r1, "" if r1 else ("未取到图表对象" if not d else "存在非引用给出的系列数据"))

    # +5：三处引用均指向含看板标题的工作表，且范围落在该表 A2:E5 之内
    r2 = False
    if d and ds_titled and d["ref_areas"]:
        r2 = True
        for sname, sarea in d["ref_areas"]:
            if sname and sname != ds.name:
                r2 = False
                break
            if not _inside(_parse_area(sarea), box_a2e5):
                r2 = False
                break
        if d["ref_sheets"] and not d["ref_sheets"].issubset(set([ds.name])):
            r2 = False
    emit(r2, "" if r2 else "引用未全部落在看板工作表的 A2:E5 内")

    # +5：随文件保存了完整数据缓存（12 个数值 + 4 个分类名称）或图表渲染结果
    r3 = bool(d) and (d["has_data_cache"] or d["has_render_cache"])
    emit(r3, "" if r3 else "文件内既无完整数据缓存也无渲染结果")

    # +5：类型为三维柱形图（柱体垂直向上、簇状分组、三条坐标轴齐备）
    r4 = bool(d) and d["is_bar3d"] and d["bar_vertical"] and d["clustered"] \
        and d["has_cat_ax"] and d["has_val_ax"] and d["has_ser_ax"]
    emit(r4, "" if r4 else "非簇状三维柱形图或缺少三条坐标轴")

    # +5：三维立体效果实际生效
    # ① 绕水平轴俯视旋转 5°~30°；② 绕垂直轴旋转 10°~40°；③ 启用直角坐标轴；
    # ④ 3 排柱体沿深度方向前后错开：簇状分组 + 深度轴承载 3 个系列。
    r5 = False
    if d and d["r_ang_ax"] and d["has_ser_ax"] and d["clustered"] and len(d["series"]) == 3:
        rx, ry = d["rot_x"], d["rot_y"]
        r5 = (rx is not None and 5 <= rx <= 30) and (ry is not None and 10 <= ry <= 40)
    emit(r5, "" if r5 else "三维旋转角度或直角坐标轴未落在细则范围内")

    # +5：深度轴（系列轴）承载 3 个年份，且该轴轴标题为“年份”
    ser_names = [_norm_year(s["name"]) for s in d["series"]] if d else []
    year_ok = (len(ser_names) == 3
               and sorted(ser_names) == [str(y) for y in YEAR_NUMS])
    r6 = bool(d) and year_ok and (d["ser_ax_title"] or "").strip() == "年份"
    emit(r6, "" if r6 else "深度轴年份文本或轴标题“年份”缺失")

    # +7：分类轴承载 4 个销售区域，且以文本型引用取自单元格
    r7 = bool(d) and d["cat_from_str_ref"] \
        and [x.strip() for x in d["cat_labels"]] == REGION_TEXTS
    emit(r7, "" if r7 else "分类标签非文本型引用取自单元格的四个区域名")

    # +5：随文件保存的 12 个数值与细则给定的三行数据完全一致
    r8 = False
    if d and len(d["series"]) == 3:
        got = {}
        for s in d["series"]:
            got[_norm_year(s["name"])] = s["vals_cached"]
        r8 = all(got.get(str(YEAR_NUMS[i])) == DATA_ROWS[i] for i in range(3))
    emit(r8, "" if r8 else "文件内未保存 12 个数值或数值与细则不一致")

    # +3：标题位置在图表顶部，文本含“出货量”及“趋势”或“变化”
    ttl = (d["title_text"] or "") if d else ""
    r9 = bool(d) and d["title_top"] and ("出货量" in ttl) \
        and ("趋势" in ttl or "变化" in ttl)
    emit(r9, "" if r9 else "标题不在顶部或文本不含约定字样")

    # +3：数值轴下限 0、上限 9500~20000、主要刻度 1000/2000/2500，且显示主要网格线
    r10 = False
    if d and d["val_explicit"] and d["val_gridlines"]:
        r10 = (abs((d["val_min"] or 0) - 0) < 1e-9
               and 9500 <= (d["val_max"] or 0) <= 20000
               and any(abs((d["val_unit"] or 0) - u) < 1e-9 for u in (1000, 2000, 2500)))
    emit(r10, "" if r10 else "数值轴刻度未显式设定或不在细则范围内")

    # ---- 轴标题也可由文本框承担，先备好位置判定 ----
    def _tb_left(p, cp):
        """紧贴数值轴左侧：不超出图表左边界一列，且纵向与图表有交叠。"""
        if cp["from_col"] is None or p["from_col"] is None:
            return False
        if p["from_col"] > cp["from_col"] + 1:
            return False
        top = cp["from_row"] or 1
        bot = cp["to_row"] or top
        p_top = p["from_row"] or 1
        p_bot = p["to_row"] or p_top
        return p_top <= bot and p_bot >= top

    def _tb_below(p, cp):
        """位于分类轴下方或右下方：落在图表纵向的下半部分，且不在图表左侧之外。"""
        top = cp["from_row"] or 1
        bot = cp["to_row"] or top
        if (p["from_row"] or 1) < top + (bot - top) * 0.5:
            return False
        return (p["from_col"] or 1) >= (cp["from_col"] or 1)

    def _tb_match(vertical, keywords, where):
        for tb in textboxes:
            if not d or tb["sheet"] != d["sheet"] or bool(tb["vertical"]) != vertical:
                continue
            hit_txt = any(any(k in t for k in keywords) for t in (tb["fwd"], tb["rev"]))
            if hit_txt and where(tb["pos"], d["pos"]):
                return True
        return False

    # +3：数值轴轴标题含“销售额”或“出货量”；也可由紧贴数值轴左侧的竖排文本框承担
    vt = (d["val_ax_title"] or "") if d else ""
    r11 = bool(d) and (("销售额" in vt) or ("出货量" in vt)
                       or _tb_match(True, ("销售额", "出货量"), _tb_left))
    emit(r11, "" if r11 else "数值轴轴标题缺失且无承担该标题的竖排文本框")

    # +3：分类轴轴标题为“销售区域”；也可由分类轴下方/右下方的横排文本框承担
    ct = (d["cat_ax_title"] or "").strip() if d else ""
    r12 = bool(d) and (ct == "销售区域" or _tb_match(False, ("销售区域",), _tb_below))
    emit(r12, "" if r12 else "分类轴轴标题缺失且无承担该标题的横排文本框")

    # +3：数据标签显示数值，为整数写法，字号不大于 14 磅
    r13 = bool(d) and d["dlbl_show_val"] and d["dlbl_integer"] \
        and (d["dlbl_font_pt"] is None or d["dlbl_font_pt"] <= 14.0)
    emit(r13, "" if r13 else "数据标签未显示数值、含小数位或字号超过 14 磅")

    # +3：图例位置为底部，图例项数为 3 且均为年份文本
    r14 = bool(d) and d["legend_bottom"] and len(d["legend_entries"]) == 3 \
        and all(_norm_year(x) in [str(y) for y in YEAR_NUMS] for x in d["legend_entries"])
    emit(r14, "" if r14 else "图例不在底部或图例项非 3 个年份文本")

    # +3：3 个系列各设独立填充色，颜色互不相同且非白、非浅灰
    fills = [(s["fill"] or "").upper()[-6:] for s in d["series"]] if d else []
    r15 = (len(fills) == 3 and all(f != "" for f in fills)
           and len(set(fills)) == 3 and not any(_is_pale(f) for f in fills))
    emit(r15, "" if r15 else "系列填充色缺失、重复或为白/浅灰")

    # +3：与 A1:E5 数据表位于同一工作表，且锚定范围与 A1:E5 不相交
    r16 = False
    if d and ds_titled and d["sheet"] == ds.name and d["pos"]:
        fc = d["pos"]["from_col"] or 1
        fr = d["pos"]["from_row"] or 1
        r16 = (fc >= 6) or (fr >= 7)
    emit(r16, "" if r16 else "图表未与数据表同表，或锚定范围压住 A1:E5")

    # +3：显示宽度不小于 15.00cm、显示高度不小于 8.00cm
    r17 = bool(d) and (d["w_cm"] or 0) >= 15.0 - 1e-6 and (d["h_cm"] or 0) >= 8.0 - 1e-6
    emit(r17, "" if r17 else "图表显示尺寸小于 15.00cm × 8.00cm")

    # +3：图表内中文均为可正常显示的汉字
    # 判定范围：标题、三条轴标题、分类刻度标签、图例项，以及承担轴标题的文本框文本。
    chart_texts = list(d["texts"]) if d else []
    if d:
        for tb in textboxes:
            if tb["sheet"] == d["sheet"]:
                chart_texts.append(tb["fwd"])
    r18 = bool(d) and _has_cjk(chart_texts) and not _has_garbled(chart_texts)
    emit(r18, "" if r18 else "图表文本缺少可显示的汉字或含乱码字符")

    # +1：图表区填充为白色或无填充，边框线宽不超过 2.00 磅
    r19 = bool(d) and d["area_fill_ok"] and (d["area_border_pt"] or 0) <= 2.0 + 1e-9
    emit(r19, "" if r19 else "图表区底纹非白或边框线宽超过 2.00 磅")

    # +1：工作簿内图表对象数为 1
    r20 = (chart_count == 1)
    emit(r20, "" if r20 else "图表对象数为 %d" % chart_count)

    _dim2_penalty(wb, st, ds, ds_titled, textboxes, emit)
    return results


def _dim2_penalty(wb, st, ds, ds_titled, textboxes, emit):
    """减分项：指令未涉及、却把交付件既有格式改坏的多余修改项。"""
    box_a1e5 = (1, 1, 5, 5)
    box_a2e5 = (1, 2, 5, 5)

    def sfont(ref):
        return st.font_of(ds.style_at(ref)) if ds else None

    def sfill(ref):
        """单元格填充色；无填充返回 None。"""
        fi = st.fill_of(ds.style_at(ref)) if ds else None
        if not fi or not fi.get("pattern") or fi.get("pattern") == "none":
            return None
        fg = (fi.get("fg") or "").upper()
        return fg[-6:] if fg else None

    def sborder(ref):
        return (st.border_of(ds.style_at(ref)) if ds else None) or {}

    # -3：工作表数不为 1，或工作表名不为“Sheet1”
    p1 = not (len(wb.sheets) == 1 and wb.sheets[0].name == "Sheet1")
    emit(p1, "工作表：%s" % "、".join(s.name for s in wb.sheets))

    # -3：A1:E5 以外的单元格出现区域名文本或表内数值
    # ① 数据表所在工作表只放行 A1:E5；② 其余工作表上任一位置出现即命中；
    # ③ 公式单元格没有缓存结果值时，文件内并未保存数值，不算“出现该数值”。
    p2 = False
    p2_where = ""
    for sh in wb.sheets:
        for ref, c in sh.cells.items():
            col, row = _split_ref(ref)
            if ds is not None and sh is ds and col and _inside((col, row, col, row), box_a1e5):
                continue
            txt = c.get("text") or ""
            num = c.get("num")
            bad = any(rg in txt for rg in REGION_TEXTS)
            if not bad and num is not None:
                bad = any(abs(num - v) < 1e-9 for v in ALL_DATA_NUMS)
            if bad:
                p2 = True
                p2_where = "%s!%s" % (sh.name, ref)
                break
        if p2:
            break
    emit(p2, p2_where)

    # -3：A1:E1 不为合并单元格，或数据表所在工作表存在 A1:E1 以外的合并单元格
    boxes = ds.merge_boxes() if ds else []
    p3 = not (len(boxes) == 1 and boxes[0] == (1, 1, 5, 1))
    emit(p3, "合并区数量 %d" % len(boxes))

    # -3：A1 单元格文本不为看板标题
    p4 = not ds_titled
    emit(p4, "A1=%s" % ((ds.text_at("A1") if ds else None) or ""))

    # -3：A2、B2、C2、D2、E2 文本不依次为“年份”“华南”“华北”“中西部”“西部”
    got2 = [(ds.text_at(_addr(c, 2)) or "").strip() if ds else "" for c in range(1, 6)]
    p5 = (got2 != HEADER_TEXTS)
    emit(p5, "、".join(got2))

    # -3：A3、A4、A5 既不依次为数值 2020/2021/2022，也不依次为文本“2020年”…
    nums3 = [ds.num_at("A%d" % r) if ds else None for r in (3, 4, 5)]
    txts3 = [(ds.text_at("A%d" % r) or "").strip() if ds else "" for r in (3, 4, 5)]
    as_num = all(nums3[i] is not None and abs(nums3[i] - YEAR_NUMS[i]) < 1e-9 for i in range(3))
    as_txt = (txts3 == YEAR_TEXTS)
    p6 = not (as_num or as_txt)
    emit(p6, "、".join(txts3) or str(nums3))

    # -3：三行出货量数值不依次匹配
    p7 = False
    bad_v = []
    for i, row in enumerate(DATA_ROWS):
        r = 3 + i
        for j, want in enumerate(row):
            ref = _addr(2 + j, r)
            got = ds.num_at(ref) if ds else None
            if got is None or abs(got - want) > 1e-9:
                p7 = True
                bad_v.append("%s=%s" % (ref, got))
    emit(p7, "、".join(bad_v))

    # -3：B3:E5 未设渐变数据条条件格式，或颜色不为浅蓝色，或被设为图标集、色阶等
    box_b3e5 = (2, 3, 5, 5)
    db = None
    for item in (ds.databars if ds else []):
        for sq in (item["sqref"] or "").split():
            if _inside(box_b3e5, _parse_area(sq)):
                db = item
                break
        if db:
            break
    p8 = True
    p8_detail = "未找到覆盖 B3:E5 的数据条"
    if db is not None:
        col_ok = (db["color"] or "").upper()[-6:] == BAR_COLOR
        p8 = (not db["gradient"]) or (not col_ok)
        p8_detail = "数据条色 #%s，渐变=%s" % (db["color"] or "", db["gradient"])
    for oc in (ds.other_cf if ds else []):
        for sq in (oc["sqref"] or "").split():
            if _boxes_hit(_parse_area(sq), box_b3e5):
                p8 = True
                p8_detail = "存在其他条件格式：%s" % oc["type"]
    emit(p8, p8_detail)

    # -1：数据条最短/最长长度不为 10%/90%，或坐标轴位置为中点
    # 未显式指定长度时按 OOXML 默认 10 与 90 计。
    p9 = True
    p9_detail = "无数据条"
    if db is not None:
        ext = (ds.databar_ext or {}).get(db["ext_id"] or "", {})
        mn = ext.get("minLength")
        mx = ext.get("maxLength")
        mn = 10.0 if mn is None else mn
        mx = 90.0 if mx is None else mx
        axp = (ext.get("axisPosition") or "automatic").lower()
        p9 = (abs(mn - 10.0) > 0.01) or (abs(mx - 90.0) > 0.01) or (axp == "midpoint")
        p9_detail = "minLength=%s maxLength=%s axisPosition=%s" % (mn, mx, axp)
    emit(p9, p9_detail)

    # -3：字体不为微软雅黑 10 磅（A1、B3:E5 常规；A2:E2、A3:A5 加粗）
    def font_ok(ref, bold):
        f = sfont(ref) or {}
        nm = (f.get("name") or "").strip().lower()
        if nm not in YAHEI_NAMES:
            return False
        sz = f.get("sz")
        if sz is None or abs(sz - 10.0) > 0.01:
            return False
        if bool(f.get("bold")) != bold:
            return False
        return not bool(f.get("italic"))

    p10 = False
    bad_f = []
    checks10 = [("A1", False)]
    checks10 += [(ref, False) for ref in _area_cells(box_b3e5)]
    checks10 += [(_addr(c, 2), True) for c in range(1, 6)]
    checks10 += [("A%d" % r, True) for r in (3, 4, 5)]
    for ref, bold in checks10:
        if not font_ok(ref, bold):
            p10 = True
            bad_f.append(ref)
    emit(p10, "、".join(bad_f))

    # -3：A2:E2 填充色不为深蓝黑，或该行字体颜色不为白色
    p11 = False
    bad_h2 = []
    for c in range(1, 6):
        ref = _addr(c, 2)
        fg = sfill(ref)
        fc = ((sfont(ref) or {}).get("color") or "").upper()[-6:]
        if fg != HEADER_FILL or fc != WHITE:
            p11 = True
            bad_h2.append("%s(填充%s/字色%s)" % (ref, fg, fc))
    emit(p11, "、".join(bad_h2))

    # -3：A3:E3、A5:E5 填充色不为白色，或 A4:E4 填充色不为浅蓝紫
    # 未设填充与纯白显示效果一致，按白色计。
    p12 = False
    bad_fill = []
    for r in (3, 5):
        for c in range(1, 6):
            fg = sfill(_addr(c, r))
            if fg is not None and fg != WHITE:
                p12 = True
                bad_fill.append("%s=%s" % (_addr(c, r), fg))
    for c in range(1, 6):
        fg = sfill(_addr(c, 4))
        if fg != STRIPE_FILL:
            p12 = True
            bad_fill.append("%s=%s" % (_addr(c, 4), fg))
    emit(p12, "、".join(bad_fill))

    # -3：B3:E5 数字格式不为三位一分隔、无小数位的整数格式，或为文本型、百分比
    p13 = False
    bad_nf = []
    for ref in _area_cells(box_b3e5):
        code = st.numfmt_of(ds.style_at(ref)) if ds else None
        if not _is_int_thousands_fmt(code):
            p13 = True
            bad_nf.append("%s=%s" % (ref, code))
    emit(p13, "、".join(bad_nf))

    # -1：A1 及 A2:E5 内任一单元格水平或垂直对齐不为居中
    # A1:E1 合并区内被覆盖的 B1:E1 不参与判定。
    p14 = False
    bad_al = []
    for ref in ["A1"] + _area_cells(box_a2e5):
        h, v = st.align_of(ds.style_at(ref)) if ds else (None, None)
        if h != "center" or v != "center":
            p14 = True
            bad_al.append("%s(%s/%s)" % (ref, h, v))
    emit(p14, "、".join(bad_al))

    # -1：A2:A5 左侧不为细实线边框，或 A5:E5 下侧不为细实线边框
    p15 = False
    bad_bd = []
    for r in (2, 3, 4, 5):
        if sborder("A%d" % r).get("left") != "thin":
            p15 = True
            bad_bd.append("A%d 左" % r)
    for c in range(1, 6):
        if sborder(_addr(c, 5)).get("bottom") != "thin":
            p15 = True
            bad_bd.append("%s 下" % _addr(c, 5))
    emit(p15, "、".join(bad_bd))

    # -1：列宽不为 A=10.00、B~E=12.00、F=13.00 字符（各列允许 ±1.50）
    want_w = {1: 10.0, 2: 12.0, 3: 12.0, 4: 12.0, 5: 12.0, 6: 13.0}
    p16 = False
    bad_w = []
    for ci, want in sorted(want_w.items()):
        got = ds.col_width(ci) if ds else None
        if got is None or abs(got - want) > 1.5:
            p16 = True
            bad_w.append("%s列=%s" % (_addr(ci, 1)[0], got))
    emit(p16, "、".join(bad_w))

    # -1：行高不为第 1、2 行 32.0 磅，第 3 行 22.0 磅，第 4、5 行 24.0 磅（±1.5）
    want_h = {1: 32.0, 2: 32.0, 3: 22.0, 4: 24.0, 5: 24.0}
    p17 = False
    bad_h = []
    for ri, want in sorted(want_h.items()):
        got = ds.row_height(ri) if ds else None
        if got is None or abs(got - want) > 1.5:
            p17 = True
            bad_h.append("第%d行=%s" % (ri, got))
    emit(p17, "、".join(bad_h))

    # -3：任一工作表存在带筛选按钮的表格对象或自动筛选区域
    p18 = False
    p18_detail = ""
    for sh in wb.sheets:
        if sh.has_autofilter:
            p18 = True
            p18_detail = "%s 自动筛选" % sh.name
        for tb in sh.tables:
            hdr = tb.get("header")
            # 表格对象带标题行时默认显示筛选按钮；headerRowCount="0" 才没有
            if tb.get("autofilter") or hdr is None or str(hdr).strip() != "0":
                p18 = True
                p18_detail = "%s 表格对象 %s" % (sh.name, tb.get("name") or "")
    emit(p18, p18_detail)

    # -1：任一工作表设有冻结窗格
    p19 = False
    p19_detail = ""
    for sh in wb.sheets:
        if sh.freeze:
            p19 = True
            p19_detail = "%s 冻结窗格" % sh.name
    emit(p19, p19_detail)

    # -1：任一单元格文本、文本框文本或批注中出现制图说明类字样
    p20 = False
    p20_detail = ""
    hay = []
    for sh in wb.sheets:
        for c in sh.cells.values():
            if c.get("text"):
                hay.append(c["text"])
    for tb in (textboxes or []):
        if tb.get("fwd"):
            hay.append(tb["fwd"])
    hay.extend(wb.comment_texts or [])
    for t in hay:
        for kw in NOTE_KEYWORDS:
            if kw in t:
                p20 = True
                p20_detail = kw
                break
        if p20:
            break
    emit(p20, p20_detail)

    # -1：第 6 行及以下、F 列及以后写入了除区域名与出货量数值以外的内容
    # 仅设有行高、列宽或单元格样式而无内容的不计。
    p21 = False
    p21_detail = ""
    for ref, c in (ds.cells.items() if ds else []):
        col, row = _split_ref(ref)
        if not col or (row < 6 and col < 6):
            continue
        txt = (c.get("text") or "").strip()
        num = c.get("num")
        if txt:
            if txt not in REGION_TEXTS:
                p21 = True
                p21_detail = "%s=%s" % (ref, txt[:20])
                break
            continue
        if num is not None:
            if not any(abs(num - v) < 1e-9 for v in ALL_DATA_NUMS):
                p21 = True
                p21_detail = "%s=%s" % (ref, num)
                break
            continue
        if c.get("formula"):
            p21 = True
            p21_detail = "%s 含公式" % ref
            break
    emit(p21, p21_detail)

    # -1：页边距左右不为 0.7 英寸、上下不为 0.75 英寸、页眉页脚不为 0.3 英寸
    pm = (ds.page_margins if ds else None) or {}
    want_m = [("left", 0.7), ("right", 0.7), ("top", 0.75), ("bottom", 0.75),
              ("header", 0.3), ("footer", 0.3)]
    p22 = False
    bad_m = []
    for key, want in want_m:
        got = pm.get(key)
        if got is None or abs(got - want) > 0.005:
            p22 = True
            bad_m.append("%s=%s" % (key, got))
    emit(p22, "、".join(bad_m))


def _finish(result, results, details):
    """把 “+N/-N：规则说明” 字符串反解为 dim2_items，并汇总总分。

    details 是判定时收集的实测值，只用于本地排查，不写入返回信息：命中与未命中
    登记的规则原文完全一致，返回的 rule 不随判定结果变化。"""
    if len(results) != len(_MAX_DELTAS_ORDER):
        raise RuntimeError("规则条数不匹配：%d != %d" % (len(results), len(_MAX_DELTAS_ORDER)))
    dim2_items = []
    total_score = 0
    for i, raw in enumerate(results):
        max_delta = _MAX_DELTAS_ORDER[i]
        head, _, rule = raw.partition("：")
        try:
            delta = int(head.replace("+", "").strip())
        except ValueError:
            delta = 0
        hit = (delta != 0)
        applied = max_delta if hit else 0
        dim2_items.append({
            "rule": rule,
            "max_delta": max_delta,
            "delta": applied,
            "hit": hit,
        })
        total_score += applied
    result["dim2_items"] = dim2_items
    result["total_score"] = max(0, total_score)
    result["max_score"] = MAX_SCORE_TOTAL
    return result


if __name__ == "__main__":
    if len(sys.argv) >= 2:
        _dir = sys.argv[1]
    else:
        _dir = os.path.dirname(os.path.abspath(__file__))
    _out = evaluate(_dir)
    print(json.dumps(_out, ensure_ascii=False, indent=2))
