#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# pyright: reportUnknownMemberType=false, reportUnknownVariableType=false, reportUnknownArgumentType=false, reportMissingTypeStubs=false, reportUnknownParameterType=false, reportMissingParameterType=false, reportOptionalMemberAccess=false, reportOptionalSubscript=false, reportGeneralTypeIssues=false

import os
import re
import sys
import json
import zipfile
import xml.etree.ElementTree as ET

# 说明：按接口统一约定，本模块不修改全局 sys.stdout，也不 print 主结果。
# 主结果统一由 evaluate(dir_path) -> dict 返回；入参为“脚本所在目录的路径”，
# 由脚本自己在该目录内定位并打开被评估的 .xlsx / .xlsm。本文件仅在 __main__ 下
# 作为本地调试入口，才把 dict 序列化为 JSON 打印。
#
# 与 008（Word）脚本的差异：本任务的细则要点全部落在 OOXML 原始部件上
# （ofPieChart 的 splitType/splitPos、strRef vs strLit、serLines、
# theme 字体表、pageMargins、mergeCell 等），
# 这些是 openpyxl 读不全或会在读写往返中丢失的信息，因此本脚本直接用
# zipfile + ElementTree 解析包内 XML，不依赖第三方库。
#
# 评分符号约定（沿用 008 的 "±N：说明" 扁平字符串，再反解为 dim2_items）：
#   加分项命中 -> "+N：…"，未命中 -> "+0：…"
#   减分项命中 -> "-N：…"（N 为该项扣分绝对值），未命中 -> "-0：…"
# 因此 _MAX_DELTAS_ORDER 中加分项为正、减分项为负；MAX_SCORE_TOTAL 只统计加分项。

SCRIPT_ID = "086"
MAX_SCORE_TOTAL = 42

_MAX_DELTAS_ORDER = [
    # ---- 一、加分项（均来源于指令要求的制图任务）----
    5,    # 01 Sheet2 内为 1 个原生复合饼图（ofPieChart），主子饼同属一个图表对象
    3,    # 02 主饼三块：一级/二级/三级产品，二级 34%、三级 11%
    3,    # 03 子饼三块：优等品/合格品/残次品，且仅含这三项
    5,    # 04 子饼数值已折算为对全体占比（8.25%/35.75%/11%，合计 55%，全图 100%）
    3,    # 05 被展开的那一块显示为“一级产品”，而非“其他/Other”
    3,    # 06 子饼拆分方式为按位置取末 3 项（splitType=pos、splitPos=3）
    1,    # 07 图表落位在名为 Sheet2 的工作表内
    5,    # 08 数据系列通过单元格区域引用取数（strRef/numRef），而非写死字面量
    3,    # 09 子饼三项折算值由公式生成
    1,    # 10 图表引用的类别名与数值在表内以单元格形式完整列出
    3,    # 11 数据标签同时显示“类别名称”与“百分比”
    3,    # 12 被展开扇区与子饼之间绘有连接线（系列线）
    1,    # 13 未使用图例，类别名直接标注在扇区上
    1,    # 14 百分比标签为百分比格式
    1,    # 15 图表带有标题
    1,    # 16 主饼左、子饼右且明显更小，图表宽度不小于约 12 厘米
    # ---- 二、减分项（指令未要求、却被改动或凭空添加）----
    -3,   # 17 标题/类别名/取数区域出现与给定数据无关的自造名称
    -1,   # 18 Sheet2 添加未被要求的装饰性内容（跨列合并大标题、“说明：…”文字块）
    -1,   # 19 Sheet2 添加数据中不存在的统计项（合计/求和行）
    -1,   # 20 除 Sheet2 外新增工作表，或已有工作表被改名
    -3,   # 21 六项给定名称或数值被删除、改写或被覆盖
    -1,   # 22 明细表左侧纵向合并的“一级产品”单元格被取消合并
    -1,   # 23 两张数据表被搬移到其他工作表，或被图表压盖
    -1,   # 24 比例数值被改为文本型，或 0.55 被改写为 55
    -1,   # 25 数据区单元格格式（0% / 深色填充+白色加粗表头 / 边框）被破坏
    -1,   # 26 数据所在工作表页边距被改动
    -1,   # 27 工作簿主题字体方案被简化
]
assert sum(d for d in _MAX_DELTAS_ORDER if d > 0) == MAX_SCORE_TOTAL

# ---- 命名空间 ----
NS_MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
NS_C = 'http://schemas.openxmlformats.org/drawingml/2006/chart'
NS_A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
NS_XDR = 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing'

# ---- 给定数据（细则“说明”里写明的两张小表）----
SHEET2_NAME = 'Sheet2'
DATA_SHEET_NAME = 'Sheet1'
MAIN_ITEMS = (('一级产品', 0.55), ('二级产品', 0.34), ('三级产品', 0.11))
SUB_ITEMS = (('优等品', 0.15), ('合格品', 0.65), ('残次品', 0.20))
# 折算到全体后的子饼三项：0.55×0.15、0.55×0.65、0.55×0.20
SUB_FOLDED = (('优等品', 0.0825), ('合格品', 0.3575), ('残次品', 0.1100))
GIVEN_NAMES = tuple(n for n, _ in MAIN_ITEMS) + tuple(n for n, _ in SUB_ITEMS)
EXPLODED_NAME = '一级产品'
# 默认聚合名（细则第 5 项要求被展开那块不能显示成这些名字）
DEFAULT_AGG_NAMES = ('其他', '其它', 'Other', 'others', '未分类')
# 细则明示的“自造名称”样例（用于图表标题这类自由文本的判定）
FABRICATED_HINTS = ('人员薪酬', '薪酬', '物料', '交通', '住宿', '餐饮',
                    '总费用', '费用构成', '其他费用', '其它费用', '费用明细', '其他明细')
# 细则明示的“凭空添加的统计项”特征词
TOTAL_ROW_HINTS = ('合计', '总计', '小计', '求和', '总和')

EMU_PER_CM = 360000.0
MIN_CHART_WIDTH_CM = 12.0          # 细则：图表宽度不小于约 12 厘米
CHART_WIDTH_TOLERANCE_CM = 0.2     # “约 12 厘米”留 2 毫米容差
SECOND_PIE_MAX = 90                # 子饼“明显小于”主饼：secondPieSize ≤ 90%（Excel 默认 75）
EXPECTED_MARGINS = {'left': 0.75, 'right': 0.75, 'top': 1.0,
                    'bottom': 1.0, 'header': 0.5, 'footer': 0.5}
# 内置数字格式：9 = 0%，10 = 0.00%
BUILTIN_NUMFMT = {'9': '0%', '10': '0.00%'}
VALUE_EPS = 5e-5


def _ln(tag):
    """去掉命名空间前缀，返回本地名。"""
    return tag.split('}')[-1]


def _join_part(base_dir, target):
    """把关系里的 Target 解析为包内绝对部件名（兼容 '/xl/...' 与 '../' 写法）。"""
    if target.startswith('/'):
        return target.lstrip('/')
    parts = [p for p in base_dir.split('/') if p]
    for seg in target.split('/'):
        if seg in ('', '.'):
            continue
        if seg == '..':
            if parts:
                parts.pop()
        else:
            parts.append(seg)
    return '/'.join(parts)


def _col_to_idx(col):
    """列字母转 1 起序号：A->1、B->2、AA->27。"""
    n = 0
    for ch in col.upper():
        if not ('A' <= ch <= 'Z'):
            return 0
        n = n * 26 + (ord(ch) - 64)
    return n


def _idx_to_col(idx):
    """1 起序号转列字母。"""
    s = ''
    while idx > 0:
        idx, rem = divmod(idx - 1, 26)
        s = chr(65 + rem) + s
    return s


_CELL_RE = re.compile(r'^\$?([A-Za-z]{1,3})\$?(\d+)$')
_REF_RE = re.compile(r"^(?:'((?:[^']|'')*)'|([^!]+))!(.+)$")


def _split_ref(ref):
    """把 "Sheet1!$B$13:$B$17" 拆成 (工作表名, [A1 式单元格列表])；解析失败返回 (None, [])。"""
    if not ref:
        return (None, [])
    m = _REF_RE.match(ref.strip())
    if not m:
        return (None, [])
    sheet = (m.group(1) or m.group(2) or '').replace("''", "'").strip()
    body = m.group(3).split(',')[0].strip()   # 多区域引用只取第 1 段
    if ':' in body:
        a, b = body.split(':', 1)
    else:
        a = b = body
    ma, mb = _CELL_RE.match(a.strip()), _CELL_RE.match(b.strip())
    if not ma or not mb:
        return (sheet, [])
    c1, r1 = _col_to_idx(ma.group(1)), int(ma.group(2))
    c2, r2 = _col_to_idx(mb.group(1)), int(mb.group(2))
    if not c1 or not c2:
        return (sheet, [])
    cells = []
    for r in range(min(r1, r2), max(r1, r2) + 1):
        for c in range(min(c1, c2), max(c1, c2) + 1):
            cells.append('%s%d' % (_idx_to_col(c), r))
    return (sheet, cells)


class _Pkg(object):
    """.xlsx（OPC 包）只读访问器：按部件名取 XML，并解析部件的关系表。"""

    def __init__(self, path):
        self.zf = zipfile.ZipFile(path)
        self.names = set(self.zf.namelist())
        self._xml_cache = {}

    def has(self, part):
        return part in self.names

    def raw(self, part):
        return self.zf.read(part)

    def text(self, part):
        return self.zf.read(part).decode('utf-8', 'replace')

    def xml(self, part):
        if part not in self._xml_cache:
            self._xml_cache[part] = ET.fromstring(self.zf.read(part))
        return self._xml_cache[part]

    def rels(self, part):
        """返回 {rId: 部件名}；外部链接的值为 None。"""
        d, base = os.path.split(part)
        rp = ('%s/' % d if d else '') + '_rels/' + base + '.rels'
        out = {}
        if rp not in self.names:
            return out
        for rel in self.xml(rp):
            rid = rel.get('Id')
            tgt = rel.get('Target') or ''
            if (rel.get('TargetMode') or '') == 'External' or tgt.startswith('http'):
                out[rid] = None
                continue
            out[rid] = _join_part(d, tgt)
        return out

    def close(self):
        try:
            self.zf.close()
        except Exception:
            pass


def _norm_rgb(color_el):
    """把 <color>/<fgColor> 归一为 6 位 RGB；主题色/索引色返回 'theme:N'/'indexed:N'。"""
    if color_el is None:
        return None
    rgb = color_el.get('rgb')
    if rgb:
        rgb = rgb.strip()
        return rgb[-6:].upper() if len(rgb) >= 6 else rgb.upper()
    if color_el.get('theme') is not None:
        return 'theme:%s' % color_el.get('theme')
    if color_el.get('indexed') is not None:
        return 'indexed:%s' % color_el.get('indexed')
    return None


class _Styles(object):
    """styles.xml：把 cellXfs 索引解析为“数字格式 / 字体 / 填充 / 边框”的实际取值。"""

    def __init__(self, pkg):
        self.numfmts = dict(BUILTIN_NUMFMT)
        self.fonts = []
        self.fills = []
        self.borders = []
        self.cellxfs = []
        if not pkg.has('xl/styles.xml'):
            return
        root = pkg.xml('xl/styles.xml')
        for el in root.iter():
            if _ln(el.tag) == 'numFmt' and el.get('numFmtId'):
                self.numfmts[el.get('numFmtId')] = el.get('formatCode') or ''
        for name, sink in (('fonts', self.fonts), ('fills', self.fills),
                           ('borders', self.borders), ('cellXfs', self.cellxfs)):
            for el in root:
                if _ln(el.tag) == name:
                    sink.extend(list(el))
                    break

    def _xf(self, s):
        try:
            i = int(s) if s is not None else 0
        except (TypeError, ValueError):
            i = 0
        return self.cellxfs[i] if 0 <= i < len(self.cellxfs) else None

    def numfmt_of(self, s):
        """单元格实际生效的数字格式代码（办公软件“设置单元格格式-数字”里看到的那个）。"""
        xf = self._xf(s)
        nid = xf.get('numFmtId') if xf is not None else '0'
        return self.numfmts.get(nid or '0', 'General')

    def font_of(self, s):
        """返回 (bold, 颜色, 字体名)。"""
        xf = self._xf(s)
        if xf is None:
            return (False, None, None)
        try:
            fi = int(xf.get('fontId') or 0)
        except ValueError:
            fi = 0
        if not (0 <= fi < len(self.fonts)):
            return (False, None, None)
        f = self.fonts[fi]
        bold = False
        color = None
        name = None
        for ch in f:
            n = _ln(ch.tag)
            if n == 'b':
                bold = (ch.get('val') in (None, '1', 'true'))
            elif n == 'color':
                color = _norm_rgb(ch)
            elif n == 'name':
                name = ch.get('val')
        return (bold, color, name)

    def fill_of(self, s):
        """返回 (patternType, 前景色)。"""
        xf = self._xf(s)
        if xf is None:
            return (None, None)
        try:
            fi = int(xf.get('fillId') or 0)
        except ValueError:
            fi = 0
        if not (0 <= fi < len(self.fills)):
            return (None, None)
        for pf in self.fills[fi]:
            if _ln(pf.tag) != 'patternFill':
                continue
            fg = None
            for ch in pf:
                if _ln(ch.tag) == 'fgColor':
                    fg = _norm_rgb(ch)
            return (pf.get('patternType'), fg)
        return (None, None)

    def has_border(self, s):
        """单元格四边是否至少有一条实际边线（办公软件“边框”按钮的效果）。"""
        xf = self._xf(s)
        if xf is None:
            return False
        try:
            bi = int(xf.get('borderId') or 0)
        except ValueError:
            bi = 0
        if not (0 <= bi < len(self.borders)):
            return False
        for side in self.borders[bi]:
            if _ln(side.tag) in ('left', 'right', 'top', 'bottom') and side.get('style'):
                return True
        return False


def _is_white(color):
    """颜色是否为白色：兼容 8 位 ARGB(FFFFFFFF)、6 位(FFFFFF) 与主题背景色。"""
    return color in ('FFFFFF', 'theme:0', 'indexed:65')


class _Sheet(object):
    """一张工作表：单元格、合并区、页边距、以及挂在其上的 drawing 部件。"""

    def __init__(self, pkg, name, sheet_id, part, sst):
        self.name = name
        self.sheet_id = sheet_id
        self.part = part
        self.cells = {}          # 'A1' -> {'t','s','v','f','text','num'}
        self.merges = []         # ['A5:A7', ...]
        self.margins = {}
        self.drawing_parts = []
        self.exists = bool(part and pkg.has(part))
        if not self.exists:
            return
        root = pkg.xml(part)
        rels = pkg.rels(part)
        for el in root.iter():
            n = _ln(el.tag)
            if n == 'c':
                self._add_cell(el, sst)
            elif n == 'mergeCell' and el.get('ref'):
                self.merges.append(el.get('ref'))
            elif n == 'pageMargins':
                self.margins = dict(el.attrib)
            elif n == 'drawing':
                tgt = rels.get(el.get('{%s}id' % NS_R))
                if tgt:
                    self.drawing_parts.append(tgt)

    def _add_cell(self, el, sst):
        ref = el.get('r')
        if not ref:
            return
        t = el.get('t')
        v = None
        f = None
        inline = None
        for ch in el:
            n = _ln(ch.tag)
            if n == 'v':
                v = ch.text
            elif n == 'f':
                f = ch.text or ''
            elif n == 'is':
                inline = ''.join(x.text or '' for x in ch.iter() if _ln(x.tag) == 't')
        text = None
        num = None
        if t == 's' and v is not None and v.strip().isdigit():
            i = int(v.strip())
            text = sst[i] if 0 <= i < len(sst) else None
        elif t in ('str', 'inlineStr'):
            text = inline if inline is not None else v
        elif t == 'e':
            text = v
        elif v is not None:
            try:
                num = float(v)
            except ValueError:
                text = v
        self.cells[ref.upper()] = {'t': t, 's': el.get('s'), 'v': v, 'f': f,
                                   'text': (text or '').strip() or None, 'num': num}

    def cell(self, ref):
        return self.cells.get(ref.upper())

    def find_text(self, needle):
        """返回文本恰好等于 needle 的单元格引用列表（按行列排序）。"""
        out = []
        for ref, c in self.cells.items():
            if c['text'] == needle:
                out.append(ref)
        return sorted(out, key=_ref_sort_key)

    def right_neighbors(self, ref, span=3):
        """从 ref 向右最多 span 格，返回非空单元格引用列表。"""
        m = _CELL_RE.match(ref)
        if not m:
            return []
        col, row = _col_to_idx(m.group(1)), int(m.group(2))
        out = []
        for d in range(1, span + 1):
            r = '%s%d' % (_idx_to_col(col + d), row)
            c = self.cells.get(r)
            if c is not None and (c['text'] is not None or c['num'] is not None):
                out.append(r)
        return out


def _ref_sort_key(ref):
    m = _CELL_RE.match(ref)
    if not m:
        return (0, 0)
    return (int(m.group(2)), _col_to_idx(m.group(1)))


def _ref_rc(ref):
    """'B13' -> (row, col)，均为 1 起。"""
    m = _CELL_RE.match(ref)
    if not m:
        return (0, 0)
    return (int(m.group(2)), _col_to_idx(m.group(1)))


def _range_rc(ref_range):
    """'A5:A7' -> (r1, c1, r2, c2)。"""
    body = ref_range.replace('$', '')
    if ':' in body:
        a, b = body.split(':', 1)
    else:
        a = b = body
    r1, c1 = _ref_rc(a)
    r2, c2 = _ref_rc(b)
    return (min(r1, r2), min(c1, c2), max(r1, r2), max(c1, c2))


def _load_sst(pkg):
    """共享字符串表。"""
    part = 'xl/sharedStrings.xml'
    if not pkg.has(part):
        return []
    out = []
    for si in pkg.xml(part):
        if _ln(si.tag) != 'si':
            continue
        out.append(''.join(t.text or '' for t in si.iter() if _ln(t.tag) == 't'))
    return out


def _load_sheets(pkg, sst):
    """按 workbook.xml 顺序返回 [_Sheet]。"""
    wb = pkg.xml('xl/workbook.xml')
    rels = pkg.rels('xl/workbook.xml')
    out = []
    for el in wb.iter():
        if _ln(el.tag) != 'sheet':
            continue
        rid = el.get('{%s}id' % NS_R)
        out.append(_Sheet(pkg, (el.get('name') or '').strip(),
                          el.get('sheetId'), rels.get(rid), sst))
    return out


# 估算值：默认列宽 8.89 字符 ≈ 1.87cm，默认行高 13.5 磅 ≈ 0.48cm。
# 仅当 anchor 里没有显式 ext（cx/cy 为 0）时用于粗估图表宽度与压盖范围。
EST_COL_CM = 1.87
EST_ROW_CM = 0.48


def _anchor_geometry(anchor):
    """解析 anchor 的位置与尺寸：返回 (from_r, from_c, to_r, to_c, w_cm, h_cm)。"""
    def _i(el, name):
        for ch in el:
            if _ln(ch.tag) == name:
                try:
                    return int((ch.text or '0').strip())
                except ValueError:
                    return 0
        return 0

    from_r = from_c = to_r = to_c = 0
    w_emu = h_emu = 0
    for ch in anchor:
        n = _ln(ch.tag)
        if n == 'from':
            from_c, from_r = _i(ch, 'col'), _i(ch, 'row')
        elif n == 'to':
            to_c, to_r = _i(ch, 'col'), _i(ch, 'row')
        elif n == 'ext':                      # oneCellAnchor 的 xdr:ext
            w_emu = max(w_emu, int(ch.get('cx') or 0))
            h_emu = max(h_emu, int(ch.get('cy') or 0))
    for ext in anchor.iter('{%s}ext' % NS_A):  # graphicFrame/xfrm 的 a:ext
        w_emu = max(w_emu, int(ext.get('cx') or 0))
        h_emu = max(h_emu, int(ext.get('cy') or 0))
    w_cm = w_emu / EMU_PER_CM
    h_cm = h_emu / EMU_PER_CM
    if _ln(anchor.tag) == 'twoCellAnchor':
        if w_cm <= 0:
            w_cm = max(0, to_c - from_c) * EST_COL_CM
        if h_cm <= 0:
            h_cm = max(0, to_r - from_r) * EST_ROW_CM
    else:
        to_c = from_c + max(1, int(round(w_cm / EST_COL_CM)))
        to_r = from_r + max(1, int(round(h_cm / EST_ROW_CM)))
    # anchor 里的行列是 0 起，转成与单元格一致的 1 起
    return (from_r + 1, from_c + 1, to_r + 1, to_c + 1, w_cm, h_cm)


def _collect_graphics(pkg, sheet):
    """遍历工作表挂载的 drawing，收集其中的原生图表与图片对象。

    返回 (charts, pics)：
      charts -> [{'part','root','sheet','rect','w_cm','h_cm'}]，rect=(r1,c1,r2,c2)
      pics   -> [{'sheet','rect'}]（xdr:pic / OLE 对象，用于维度一判定“图表被贴成图片”）
    """
    charts = []
    pics = []
    for dpart in sheet.drawing_parts:
        if not pkg.has(dpart):
            continue
        droot = pkg.xml(dpart)
        drels = pkg.rels(dpart)
        for anchor in droot:
            if _ln(anchor.tag) not in ('twoCellAnchor', 'oneCellAnchor', 'absoluteAnchor'):
                continue
            r1, c1, r2, c2, w_cm, h_cm = _anchor_geometry(anchor)
            found_chart = False
            for gd in anchor.iter('{%s}graphicData' % NS_A):
                for ch in gd:
                    if _ln(ch.tag) != 'chart':
                        continue
                    cpart = drels.get(ch.get('{%s}id' % NS_R))
                    if not cpart or not pkg.has(cpart):
                        continue
                    charts.append({'part': cpart, 'root': pkg.xml(cpart),
                                   'sheet': sheet.name, 'rect': (r1, c1, r2, c2),
                                   'w_cm': w_cm, 'h_cm': h_cm})
                    found_chart = True
            if found_chart:
                continue
            for el in anchor.iter():
                if _ln(el.tag) in ('pic', 'oleObject'):
                    pics.append({'sheet': sheet.name, 'rect': (r1, c1, r2, c2)})
                    break
    return (charts, pics)


def _rects_overlap(a, b):
    return not (a[2] < b[0] or b[2] < a[0] or a[3] < b[1] or b[3] < a[1])


def _c(el, name):
    return el.find('{%s}%s' % (NS_C, name)) if el is not None else None


def _cs(el, name):
    return el.findall('{%s}%s' % (NS_C, name)) if el is not None else []


def _cval(el, name, default=None):
    x = _c(el, name)
    if x is None:
        return default
    return x.get('val')


def _cint(el, name, default=None):
    v = _cval(el, name)
    try:
        return int(v)
    except (TypeError, ValueError):
        return default


def _tx_text(tx):
    """<c:tx>/<c:title> 里的可见文本（富文本取所有 a:t，退化取 c:v）。"""
    if tx is None:
        return ''
    ts = [t.text or '' for t in tx.iter('{%s}t' % NS_A)]
    if ts:
        return ''.join(ts)
    for v in tx.iter('{%s}v' % NS_C):
        return v.text or ''
    return ''


def _cache_points(holder):
    """从 c:strRef/c:numRef/c:strLit/c:numLit 里取 {idx: 文本或数值} 与 ptCount、formatCode。

    返回 (ref_formula_or_None, is_literal, {idx: value}, ptCount, formatCode)
    """
    if holder is None:
        return (None, False, {}, 0, None)
    ref = None
    literal = False
    src = None
    for ch in holder:
        n = _ln(ch.tag)
        if n in ('strRef', 'numRef', 'multiLvlStrRef'):
            fe = _c(ch, 'f')
            ref = (fe.text or '').strip() if fe is not None else None
            src = ch
        elif n in ('strLit', 'numLit'):
            literal = True
            src = ch
    pts = {}
    ptcount = 0
    fmt = None
    if src is not None:
        for cache in list(src) + [src]:
            if _ln(cache.tag) not in ('strCache', 'numCache', 'strLit', 'numLit',
                                      'multiLvlStrCache'):
                continue
            fe = _c(cache, 'formatCode')
            if fe is not None and fe.text:
                fmt = fe.text.strip()
            pc = _cint(cache, 'ptCount')
            if pc is not None:
                ptcount = max(ptcount, pc)
            for pt in _cs(cache, 'pt'):
                try:
                    i = int(pt.get('idx') or -1)
                except ValueError:
                    continue
                ve = _c(pt, 'v')
                pts[i] = (ve.text or '').strip() if ve is not None else None
    return (ref, literal, pts, ptcount, fmt)


def _chart_facts(root):
    """从 chartN.xml 抽出评分需要的全部事实（不做判分，只做事实提取）。"""
    f = {
        'group_kinds': [], 'is_ofpie': False, 'ofpie_type': None,
        'split_type': None, 'split_pos': None, 'split_val': None,
        'cust_split': [], 'second_pie_size': None,
        'ser_lines': False, 'ser_lines_dash': None,
        'has_legend': False, 'title_text': '', 'auto_title_deleted': False,
        'cat_ref': None, 'cat_lit': False, 'cat_pts': {}, 'cat_count': 0,
        'val_ref': None, 'val_lit': False, 'val_pts': {}, 'val_count': 0,
        'val_fmt': None, 'lbl_texts': {}, 'show': {}, 'lbl_num_fmts': [],
    }
    chart = _c(root, 'chart')
    title = _c(chart, 'title')
    f['title_text'] = _tx_text(title).strip()
    f['auto_title_deleted'] = (_cval(chart, 'autoTitleDeleted') in ('1', 'true'))
    f['has_legend'] = (_c(chart, 'legend') is not None)
    plot = _c(chart, 'plotArea')
    if plot is None:
        return f
    groups = [el for el in plot if _ln(el.tag).endswith('Chart')]
    f['group_kinds'] = [_ln(el.tag) for el in groups]
    ofpie = None
    for el in groups:
        if _ln(el.tag) == 'ofPieChart':
            ofpie = el
            break
    f['is_ofpie'] = ofpie is not None
    grp = ofpie if ofpie is not None else (groups[0] if groups else None)
    if grp is None:
        return f
    f['ofpie_type'] = _cval(grp, 'ofPieType', 'pie' if ofpie is not None else None)
    f['split_type'] = _cval(grp, 'splitType')
    f['split_pos'] = _cint(grp, 'splitPos')
    f['split_val'] = _cval(grp, 'splitVal')
    f['second_pie_size'] = _cint(grp, 'secondPieSize')
    for cs in _cs(grp, 'custSplit'):
        for pt in _cs(cs, 'secondPiePt'):
            try:
                f['cust_split'].append(int(pt.get('val')))
            except (TypeError, ValueError):
                pass
    sl = _c(grp, 'serLines')
    if sl is not None:
        f['ser_lines'] = True
        for d in sl.iter('{%s}prstDash' % NS_A):
            f['ser_lines_dash'] = d.get('val')
            break
    return _chart_series_facts(grp, f)


def _chart_series_facts(grp, f):
    """补齐系列级事实：类别/数值来源、数据标签内容与开关。"""
    ser = _c(grp, 'ser')
    if ser is None:
        return f
    ref, lit, pts, cnt, _fmt = _cache_points(_c(ser, 'cat'))
    f['cat_ref'], f['cat_lit'], f['cat_pts'], f['cat_count'] = ref, lit, pts, cnt
    ref, lit, pts, cnt, fmt = _cache_points(_c(ser, 'val'))
    f['val_ref'], f['val_lit'], f['val_pts'], f['val_count'] = ref, lit, pts, cnt
    f['val_fmt'] = fmt
    # 数据标签：系列级 dLbls（含逐点 dLbl 覆盖） + 图表组级 dLbls
    ser_dlbls = _c(ser, 'dLbls')
    grp_dlbls = _c(grp, 'dLbls')
    for dl in _cs(ser_dlbls, 'dLbl'):
        i = _cint(dl, 'idx')
        txt = _tx_text(_c(dl, 'tx')).strip()
        if i is not None and txt:
            f['lbl_texts'][i] = txt
        nf = _c(dl, 'numFmt')
        if nf is not None and nf.get('formatCode'):
            f['lbl_num_fmts'].append(nf.get('formatCode'))
    for holder in (ser_dlbls, grp_dlbls):
        if holder is None:
            continue
        nf = _c(holder, 'numFmt')
        if nf is not None and nf.get('formatCode'):
            f['lbl_num_fmts'].append(nf.get('formatCode'))
        for key in ('showVal', 'showCatName', 'showSerName', 'showPercent',
                    'showLegendKey', 'showBubbleSize'):
            v = _cval(holder, key)
            if v is None:
                continue
            # 系列级优先：先写入的（ser_dlbls）不被组级覆盖
            f['show'].setdefault(key, v in ('1', 'true'))
        pos = _cval(holder, 'dLblPos')
        if pos:
            f['show'].setdefault('dLblPos', pos)
    return f


def _shows(f, key):
    return bool(f['show'].get(key))


def _to_float(x):
    try:
        return float(x)
    except (TypeError, ValueError):
        return None


def _series_points(f, sheets_by_name):
    """把系列还原成数据点列表。

    优先按 c:f 指向的单元格取值（这才是“可编辑”的真实取数），
    单元格取不到时退回图表内缓存/字面量。
    每点：{'idx','name','value','name_cell','val_cell','has_formula'}
    """
    n = max(f['cat_count'], f['val_count'],
            (max(f['cat_pts']) + 1) if f['cat_pts'] else 0,
            (max(f['val_pts']) + 1) if f['val_pts'] else 0)
    cat_sheet, cat_cells = _split_ref(f['cat_ref'])
    val_sheet, val_cells = _split_ref(f['val_ref'])
    cs = sheets_by_name.get((cat_sheet or '').lower())
    vs = sheets_by_name.get((val_sheet or '').lower())
    pts = []
    for i in range(n):
        name = f['cat_pts'].get(i)
        value = _to_float(f['val_pts'].get(i))
        name_cell = val_cell = None
        has_formula = False
        if cs is not None and i < len(cat_cells):
            name_cell = '%s!%s' % (cs.name, cat_cells[i])
            c = cs.cell(cat_cells[i])
            if c is not None and c['text']:
                name = c['text']
        if vs is not None and i < len(val_cells):
            val_cell = '%s!%s' % (vs.name, val_cells[i])
            c = vs.cell(val_cells[i])
            if c is not None:
                if c['num'] is not None:
                    value = c['num']
                has_formula = bool(c['f'])
        pts.append({'idx': i, 'name': (name or '').strip(), 'value': value,
                    'name_cell': name_cell, 'val_cell': val_cell,
                    'has_formula': has_formula})
    return pts


def _split_sets(f, pts):
    """按 splitType 计算“留在主饼”与“进入子饼”的点序号集合。

    与办公软件对齐：右键系列-“设置数据系列格式”-“系列分割依据”
      位置(pos)   -> 末 splitPos 项进子饼
      自定义(cust) -> custSplit 里列出的点进子饼
      值(val)/百分比(percent) -> 小于阈值的点进子饼
      auto/缺省    -> 无法从文件唯一确定，按“末 3 项”粗略推断（仅用于兜底展示）
    """
    n = len(pts)
    st = f['split_type']
    sub = []
    if st == 'pos' and f['split_pos']:
        k = max(0, min(n, f['split_pos']))
        sub = list(range(n - k, n))
    elif st == 'cust':
        sub = sorted(i for i in f['cust_split'] if 0 <= i < n)
    elif st in ('val', 'percent'):
        thr = _to_float(f['split_val'])
        total = sum(p['value'] or 0.0 for p in pts) or 1.0
        for p in pts:
            v = p['value'] or 0.0
            cmpv = v if st == 'val' else v / total * 100.0
            if thr is not None and cmpv < thr:
                sub.append(p['idx'])
    else:
        sub = list(range(max(0, n - 3), n))
    main = [i for i in range(n) if i not in sub]
    return (main, sub)


def _near(a, b, eps=VALUE_EPS):
    return a is not None and b is not None and abs(a - b) <= eps


def evaluate(dir_path):
    """按 §2.2 约定，返回结构化 dict；不 print、不改 sys.stdout、不 sys.exit。

    入参：dir_path —— 被评估文件所在目录（由外部显式传入）。
    返回：{id, file_name, status, error, dim1_pass, dim1_reason, dim2_items,
           total_score, max_score}
    其中 dim2_items 每项为 {rule, max_delta, delta, hit}；
    加分项 max_delta > 0，减分项 max_delta < 0（hit=True 表示“踩到了该扣分项”）。
    目录不存在或目录内没有工作簿时，返回“未找到文件”的空结果（各评分字段留空/归零）；
    其余异常（包结构损坏等）由通用兜底捕获并返回 status='error'。
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
        "max_score": MAX_SCORE_TOTAL,
    }
    try:
        if not dir_path or not os.path.isdir(dir_path):
            raise FileNotFoundError("目录不存在或不是目录：%s" % dir_path)
        # 在 dir_path 目录内定位被评估的工作簿（忽略 ~$ 打头的 Excel 临时锁文件）
        cands = sorted(f for f in os.listdir(dir_path)
                       if f.lower().endswith(('.xlsx', '.xlsm'))
                       and not f.startswith('~$')
                       and os.path.isfile(os.path.join(dir_path, f)))
        if not cands:
            raise FileNotFoundError("目录中未找到 .xlsx 文件：%s" % dir_path)
        file_path = os.path.join(dir_path, cands[0])
        result["file_name"] = os.path.basename(file_path)
        return _evaluate_impl(file_path, result)
    except FileNotFoundError as exc:
        # 未找到待评文件：不进入维度判定，dim1_pass 留空，分值一并归零
        result["file_name"] = ""
        result["status"] = "error"
        result["error"] = str(exc)
        result["dim1_pass"] = ""
        result["dim1_reason"] = ""
        result["dim2_items"] = []
        result["total_score"] = 0
        result["max_score"] = 0
        return result
    except Exception as exc:
        result["status"] = "error"
        result["error"] = "脚本运行异常：%s: %s" % (type(exc).__name__, exc)
        result["dim1_pass"] = False
        result["dim1_reason"] = ("维度一未通过：文件无法打开或解析（%s: %s），完成度总分记 0"
                                % (type(exc).__name__, exc))
        result["dim2_items"] = []
        result["total_score"] = 0
        return result


def _fail_dim1(result, reason):
    """维度一门禁未通过：完成度总分直接记 0。

    dim1_reason 统一带上「维度一未通过：…」前缀与计分后果，便于直接作为错误摘要展示。
    """
    result["dim1_pass"] = False
    result["dim1_reason"] = "维度一未通过：%s，完成度总分记 0" % reason
    result["dim2_items"] = []
    result["total_score"] = 0
    result["max_score"] = MAX_SCORE_TOTAL
    return result


def _evaluate_impl(file_path, result):
    results = []

    # ========== 维度1：可用与可修改性 ==========
    # 细则门禁三条（任一不满足则完成度总分记 0）：
    #   ① 交付物为 1 个可正常打开的 Excel 工作簿（.xlsx，接受 .xlsm）
    #   ② 其中包含名为 Sheet2 的工作表
    #   ③ Sheet2 中的图表为可选中、可右键“选择数据/更改图表类型/设置数据系列格式”
    #      的原生图表对象（不是截图/位图/OLE 图片）
    #
    # 与办公软件对齐：
    #   ① 能否打开 -> 包是否为合法 zip、[Content_Types].xml 与 xl/workbook.xml 是否
    #      存在且 XML 可解析（打开即弹“修复”提示的文件通常在这一层就已破损）
    #   ② Sheet2 -> xl/workbook.xml 的 <sheet name="Sheet2">
    #   ③ 原生图表 -> Sheet2 的 <drawing> 指向的 drawing 部件里存在
    #      <xdr:graphicFrame>/a:graphicData/c:chart（指向 xl/charts/chartN.xml）；
    #      若只有 <xdr:pic>/OLE 对象，则是被贴成图片，不可编辑
    if not zipfile.is_zipfile(file_path):
        return _fail_dim1(result, "文件不是合法的 OOXML 工作簿包（无法打开或已损坏）：%s"
                          % os.path.basename(file_path))
    pkg = _Pkg(file_path)
    try:
        for must in ('[Content_Types].xml', 'xl/workbook.xml'):
            if not pkg.has(must):
                return _fail_dim1(result, "工作簿缺少必需部件 %s，无法正常打开" % must)
        try:
            sst = _load_sst(pkg)
            sheets = _load_sheets(pkg, sst)
        except ET.ParseError as exc:
            return _fail_dim1(result, "工作簿 XML 解析失败（打开时会提示修复）：%s" % exc)
        if not sheets:
            return _fail_dim1(result, "工作簿内没有任何工作表")

        sheets_by_name = {}
        for sh in sheets:
            sheets_by_name[sh.name.lower()] = sh
        sheet2 = sheets_by_name.get(SHEET2_NAME.lower())
        if sheet2 is None or not sheet2.exists:
            return _fail_dim1(result, "工作簿内没有名为 %s 的工作表（现有工作表：%s）"
                              % (SHEET2_NAME, '、'.join(s.name for s in sheets)))

        # 收集所有工作表上的图表与图片对象
        charts_all = []
        pics_all = []
        for sh in sheets:
            cs, ps = _collect_graphics(pkg, sh)
            charts_all.extend(cs)
            pics_all.extend(ps)
        charts_s2 = [c for c in charts_all if c['sheet'] == sheet2.name]
        pics_s2 = [p for p in pics_all if p['sheet'] == sheet2.name]
        if not charts_s2:
            if pics_s2:
                return _fail_dim1(result, "%s 中没有原生图表对象，只有图片/OLE 对象"
                                          "（图表被贴成了截图，无法右键编辑）" % SHEET2_NAME)
            return _fail_dim1(result, "%s 中没有任何图表对象" % SHEET2_NAME)

        # 数据是否整体丢失（“版面严重错乱-数据丢失”门禁；单项名值改动由减分项处理）
        surviving = [nm for nm in GIVEN_NAMES
                     if any(sh.find_text(nm) for sh in sheets)]
        if not surviving:
            return _fail_dim1(result, "给定数据的六项名称在工作簿中全部丢失（数据丢失）")

        result["dim1_pass"] = True
        result["dim1_reason"] = ("可正常解析的 .xlsx；存在 %s；%s 内含 %d 个原生图表对象"
                                % (SHEET2_NAME, SHEET2_NAME, len(charts_s2)))

        # ========== 维度2：完成度评分 ==========
        styles = _Styles(pkg)
        # 目标图表：Sheet2 内的复合饼图；没有复合饼图时取第 1 个图表，
        # 以便“落位/取数/标签/标题”等与图表类型无关的细则项仍能逐条判定。
        of_pies = [c for c in charts_s2 if _chart_facts(c['root'])['is_ofpie']]
        target = of_pies[0] if of_pies else charts_s2[0]
        facts = _chart_facts(target['root'])
        pts = _series_points(facts, sheets_by_name)
        main_idx, sub_idx = _split_sets(facts, pts)
        by_name = {}
        for p in pts:
            by_name.setdefault(p['name'], p)
        kinds = '/'.join(facts['group_kinds']) or '无'

        # +5：Sheet2 中为 1 个原生“复合饼图/子母饼图”（Pie of Pie，图表类型 ofPieChart），
        #     主饼与子饼属于同一个图表对象，而非并排摆放的两张彼此独立的饼图
        # 严格对齐细则，仅踩细则明确写出的点，不额外约束：
        #   ① Sheet2 内图表对象数量为 1（并排两张彼此独立的饼图即不给分）
        #   ② 该图表为原生复合饼图：图表组元素是 <c:ofPieChart>
        #   ③ 是“子母饼图”而非“复合条饼图”：<c:ofPieType val="pie"/>（缺省即 pie）
        #
        # 与办公软件对齐：插入-图表-饼图-“子母饼图”所生成的图表，plotArea 下的图表组
        #   为 <c:ofPieChart>，主饼与子饼共用同一个 <c:ser>（同一图表对象）；
        #   “复合条饼图”是 <c:ofPieType val="bar"/>，细则要求的是 Pie of Pie。
        if len(charts_s2) == 1 and facts['is_ofpie'] and (facts['ofpie_type'] or 'pie') == 'pie':
            results.append("+5：Sheet2 中为 1 个原生复合饼图（ofPieChart，主子饼同属一个图表对象）")
        else:
            results.append("+0：Sheet2 中为 1 个原生复合饼图（ofPieChart，主子饼同属一个图表对象）"
                           "（Sheet2 图表数=%d，图表类型=%s）" % (len(charts_s2), kinds))

        # +3：主饼（左侧）由“一级产品”“二级产品”“三级产品”三块构成，
        #     其中“二级产品”为 34%、“三级产品”为 11%
        # 严格对齐细则，仅踩细则明确写出的点，不额外约束：
        #   ① 主饼可见扇区为 3 块：留在主饼的点数 + 1 块（被展开、进入子饼的聚合扇区）
        #   ② 留在主饼的点里有“二级产品”，其值为 34%
        #   ③ 留在主饼的点里有“三级产品”，其值为 11%
        #
        # 与办公软件对齐：复合饼图的一条系列按“系列分割依据”把末几项挪到子饼，
        #   主饼上显示的是“未被挪走的点 + 1 个聚合扇区”，即右键“选择数据”里
        #   看到的前 2 项加上那块被展开的合计扇区。
        main_names = [pts[i]['name'] for i in main_idx]
        main_slices = len(main_idx) + (1 if sub_idx else 0)
        m2 = next((pts[i] for i in main_idx if pts[i]['name'] == '二级产品'), None)
        m3 = next((pts[i] for i in main_idx if pts[i]['name'] == '三级产品'), None)
        main_ok = (main_slices == 3
                   and m2 is not None and _near(m2['value'], 0.34)
                   and m3 is not None and _near(m3['value'], 0.11))
        _detail = "主饼可见扇区=%d，留在主饼的项=%s" % (main_slices, '/'.join(main_names) or '无')
        if main_ok:
            results.append("+3：主饼由一级/二级/三级产品三块构成，二级产品 34%%、三级产品 11%%（%s）"
                           % _detail)
        else:
            results.append("+0：主饼由一级/二级/三级产品三块构成，二级产品 34%%、三级产品 11%%（%s）"
                           % _detail)

        # +3：子饼（右侧）由“优等品”“合格品”“残次品”三块构成，且仅含这三项
        # 严格对齐细则，仅踩细则明确写出的点，不额外约束：
        #   ① 进入子饼的点恰为 3 个
        #   ② 这 3 个点的类别名集合恰为 {优等品, 合格品, 残次品}（不多不少）
        #
        # 与办公软件对齐：子饼里的扇区就是被“系列分割依据”挪到第二绘图区的那些点，
        #   在“选择数据”对话框里属于同一条系列的末尾几项。
        sub_names = [pts[i]['name'] for i in sub_idx]
        sub_ok = (len(sub_idx) == 3
                  and set(sub_names) == set(n for n, _ in SUB_ITEMS))
        if sub_ok:
            results.append("+3：子饼由优等品/合格品/残次品三块构成且仅含这三项（%s）"
                           % '/'.join(sub_names))
        else:
            results.append("+0：子饼由优等品/合格品/残次品三块构成且仅含这三项（实际子饼项=%s）"
                           % ('/'.join(sub_names) or '无'))

        # +5：子饼三项数值已按“一级产品占比 × 内部比例”折算为对全体的占比
        #     （优等品 8.25%、合格品 35.75%、残次品 11%），三项之和等于主饼中
        #     “一级产品”的 55%，全图合计 100%
        # 严格对齐细则，仅踩细则明确写出的点，不额外约束：
        #   ① 三项取值分别为 0.0825 / 0.3575 / 0.11（按类别名对应，容差 5e-5）
        #   ② 三项之和为 0.55
        #   ③ 全部数据点之和为 1.00
        #
        # 与办公软件对齐：复合饼图的子饼是把同一条系列的末几项单独画出来，
        #   其百分比标签是“占全部数据点之和”的比例；因此只有把 15%/65%/20%
        #   折算成 8.25%/35.75%/11% 后，子饼标签才与“一级产品 55%”自洽。
        folded_ok = True
        for nm, expect in SUB_FOLDED:
            p = next((pts[i] for i in sub_idx if pts[i]['name'] == nm), None)
            if p is None or not _near(p['value'], expect):
                folded_ok = False
                break
        sub_sum = sum((pts[i]['value'] or 0.0) for i in sub_idx)
        all_sum = sum((p['value'] or 0.0) for p in pts)
        folded_ok = folded_ok and _near(sub_sum, 0.55) and _near(all_sum, 1.0)
        _detail = "子饼三项之和=%.4f，全图合计=%.4f" % (sub_sum, all_sum)
        if folded_ok:
            results.append("+5：子饼三项已折算为对全体的占比（8.25%%/35.75%%/11%%，合计 55%%，"
                           "全图 100%%）（%s）" % _detail)
        else:
            results.append("+0：子饼三项已折算为对全体的占比（8.25%%/35.75%%/11%%，合计 55%%，"
                           "全图 100%%）（%s）" % _detail)

        # +3：主饼中被展开的那一块显示为“一级产品”，而非“其他”“Other”等默认聚合名称
        # 严格对齐细则，仅踩细则明确写出的点，不额外约束：
        #   ① 被展开（聚合）的那块扇区上显示的文字里含“一级产品”
        #      —— 或者该块本身就是一个名为“一级产品”的主饼数据点
        #   ② 显示文字不是“其他/其它/Other”等默认聚合名称
        #
        # 与办公软件对齐：复合饼图里被展开的那块是软件自动生成的聚合扇区，
        #   其类别名不来自单元格，默认显示“其他”；要显示成“一级产品”，
        #   需要单击该扇区的数据标签直接改写文字，在 XML 中即该点的
        #   <c:dLbl><c:idx val="N"/><c:tx><c:rich> 富文本覆盖（N 等于数据点总数）。
        agg_text = facts['lbl_texts'].get(len(pts), '')
        main_has_name = any(pts[i]['name'] == EXPLODED_NAME for i in main_idx)
        is_default_agg = any(d.lower() in agg_text.lower()
                             for d in DEFAULT_AGG_NAMES) if agg_text else False
        agg_ok = ((EXPLODED_NAME in agg_text) or main_has_name) and not is_default_agg
        _shown = agg_text.replace('\r', ' ').replace('\n', ' ') or '（无标签覆盖，按默认聚合名显示）'
        if agg_ok:
            results.append("+3：被展开的扇区显示为“一级产品”而非“其他/Other”（展开块标签：%s）"
                           % _shown)
        else:
            results.append("+0：被展开的扇区显示为“一级产品”而非“其他/Other”（展开块标签：%s）"
                           % _shown)

        # +3：子饼拆分方式设为“按位置”取数据序列末 3 项（splitType=pos、splitPos=3），
        #     前 2 项留在主饼、后 3 项进入子饼
        # 严格对齐细则，仅踩细则明确写出的点，不额外约束：
        #   ① <c:splitType val="pos"/>
        #   ② <c:splitPos val="3"/>
        #
        # 与办公软件对齐：右键系列-“设置数据系列格式”-“系列分割依据=位置”，
        #   “第二绘图区中的值”= 3。
        split_ok = (facts['split_type'] == 'pos' and facts['split_pos'] == 3)
        _detail = "splitType=%s、splitPos=%s" % (facts['split_type'] or '缺省',
                                                facts['split_pos'])
        if split_ok:
            results.append("+3：子饼拆分方式为按位置取末 3 项（%s）" % _detail)
        else:
            results.append("+0：子饼拆分方式为按位置取末 3 项（%s）" % _detail)

        # +1：图表落位在名为 Sheet2 的工作表内，而非放在数据所在工作表或另建的新工作表
        # 严格对齐细则，仅踩细则明确写出的点，不额外约束：
        #   ① Sheet2 内存在图表对象
        #   ② 其他工作表上没有图表对象（没有放到数据表或另建的新表上）
        #
        # 与办公软件对齐：图表落在哪张表，取决于该表的 <drawing r:id> 指向的
        #   drawing 部件里是否挂着 graphicFrame/chart；图表工作表（chartsheet）
        #   属于“另建的新工作表”，同样不算落位在 Sheet2。
        other_chart_sheets = sorted(set(c['sheet'] for c in charts_all
                                        if c['sheet'] != sheet2.name))
        if charts_s2 and not other_chart_sheets:
            results.append("+1：图表落位在名为 Sheet2 的工作表内")
        else:
            results.append("+0：图表落位在名为 Sheet2 的工作表内（其他表上的图表：%s）"
                           % ('、'.join(other_chart_sheets) or '无'))

        # +5：图表数据系列通过单元格区域引用取数（如 类别 Sheet1!$B$13:$B$17、
        #     数值 Sheet1!$C$13:$C$17），修改单元格数值后图表随之更新；
        #     而非把 55%/34%/11% 等数值以常量方式写死在图表定义中（strLit/numLit 字面量）
        # 严格对齐细则，仅踩细则明确写出的点，不额外约束：
        #   ① 类别来自 <c:cat><c:strRef><c:f>区域</c:f>
        #   ② 数值来自 <c:val><c:numRef><c:f>区域</c:f>
        #   ③ 不使用 strLit/numLit 字面量
        #
        # 与办公软件对齐：选中图表-右键“选择数据”，若“轴标签/系列值”框里是
        #   =Sheet1!$B$13:$B$17 这样的区域地址，就是引用取数；若是 ={0.55,0.34,…}
        #   这种大括号常量数组，则改单元格图表不会变。
        ref_ok = (bool(facts['cat_ref']) and bool(facts['val_ref'])
                  and not facts['cat_lit'] and not facts['val_lit'])
        _detail = "类别=%s、数值=%s" % (facts['cat_ref'] or ('字面量' if facts['cat_lit'] else '无'),
                                     facts['val_ref'] or ('字面量' if facts['val_lit'] else '无'))
        if ref_ok:
            results.append("+5：数据系列通过单元格区域引用取数（%s）" % _detail)
        else:
            results.append("+0：数据系列通过单元格区域引用取数（%s）" % _detail)

        # +3：子饼三项的折算值由公式生成（如 =C2*C5、=C2*C6、=C2*C7），
        #     而非手工敲入的静态数字
        # 严格对齐细则，仅踩细则明确写出的点，不额外约束：
        #   ① 子饼三项取数所在的单元格都带公式（<f> 节点）
        #
        # 与办公软件对齐：单击该单元格，编辑栏里显示的是 =C2*C5 这样的公式而不是
        #   0.0825 这样的常量；在 XML 中即 <c><f>C2*C5</f><v>0.0825</v></c>。
        sub_cells = [pts[i] for i in sub_idx]
        formula_cells = [p for p in sub_cells if p['has_formula']]
        formula_ok = bool(sub_cells) and len(formula_cells) == len(sub_cells)
        _detail = "带公式取数单元格 %d/%d" % (len(formula_cells), len(sub_cells))
        if formula_ok:
            results.append("+3：子饼三项折算值由公式生成（%s）" % _detail)
        else:
            results.append("+0：子饼三项折算值由公式生成（%s）" % _detail)

        # +1：图表所引用的类别名与数值在工作表中以单元格形式完整列出，
        #     可直接在表内改名、改值
        # 严格对齐细则，仅踩细则明确写出的点，不额外约束：
        #   ① 图表用到的每一个类别名都能在表内找到承载它的单元格
        #   ② 图表用到的每一个数值都能在表内找到承载它的单元格
        #   （引用取数时即检查引用区域内单元格是否逐格非空；写死字面量时，
        #     退化为“这些名与值是否仍以单元格形式列在表内”——细则这一项只要求
        #     “在工作表中以单元格形式完整列出”，未要求必须被图表引用）
        #
        # 与办公软件对齐：在表内直接改这些单元格即可改名改值，对应
        #   “选择数据”里能看到的取数区域，或工作表上肉眼可见的那两列。
        listed_ok = True
        if facts['cat_ref'] and facts['val_ref']:
            for p in pts:
                if not p['name'] or p['value'] is None:
                    listed_ok = False
                    break
                if not p['name_cell'] or not p['val_cell']:
                    listed_ok = False
                    break
        else:
            for p in pts:
                name_hit = any(sh.find_text(p['name']) for sh in sheets) if p['name'] else False
                val_hit = any(any(_near(c['num'], p['value']) for c in sh.cells.values())
                              for sh in sheets) if p['value'] is not None else False
                if not (name_hit and val_hit):
                    listed_ok = False
                    break
        if listed_ok and pts:
            results.append("+1：图表引用的类别名与数值在表内以单元格形式完整列出（共 %d 项）"
                           % len(pts))
        else:
            results.append("+0：图表引用的类别名与数值在表内以单元格形式完整列出（共 %d 项）"
                           % len(pts))

        # +3：数据标签同时显示“类别名称”与“百分比”（如“二级产品 34%”“合格品 36%”），
        #     而非只显示数值或只显示百分比
        # 严格对齐细则，仅踩细则明确写出的点，不额外约束：
        #   ① 标签开关同时开启“类别名称”与“百分比”
        #      （<c:showCatName val="1"/> + <c:showPercent val="1"/>）
        #   ② 或者：逐点富文本标签里同时出现类别名与百分号（手改标签的等效效果）
        #
        # 与办公软件对齐：右键数据标签-“设置数据标签格式”-“标签包括”里同时勾选
        #   “类别名称”与“百分比”；系列级 dLbls 的开关优先于图表组级 dLbls。
        rich_ok = bool(pts) and all(
            p['name'] and p['name'] in facts['lbl_texts'].get(p['idx'], '')
            and '%' in facts['lbl_texts'].get(p['idx'], '')
            for p in pts)
        lbl_ok = (_shows(facts, 'showCatName') and _shows(facts, 'showPercent')) or rich_ok
        _detail = "showCatName=%s、showPercent=%s、showVal=%s、逐点富文本标签 %d 个" % (
            facts['show'].get('showCatName'), facts['show'].get('showPercent'),
            facts['show'].get('showVal'), len(facts['lbl_texts']))
        if lbl_ok:
            results.append("+3：数据标签同时显示类别名称与百分比（%s）" % _detail)
        else:
            results.append("+0：数据标签同时显示类别名称与百分比（%s）" % _detail)

        # +3：主饼被展开的扇区与子饼之间绘有连接线（系列线/引导线，虚线或点线样式）
        # 严格对齐细则，仅踩细则明确写出的点，不额外约束：
        #   ① 图表组下存在 <c:serLines>（复合饼图的“系列线”）
        #
        # 与办公软件对齐：右键系列-“设置数据系列格式”里勾选“系列线”后，
        #   主饼展开块与子饼之间出现两条连接线；线型（虚线/点线）由
        #   <c:serLines><c:spPr><a:ln><a:prstDash val="dash"/> 决定，此处一并报出。
        _detail = "serLines=%s、线型=%s" % (facts['ser_lines'],
                                          facts['ser_lines_dash'] or '默认实线')
        if facts['ser_lines']:
            results.append("+3：主饼展开块与子饼之间绘有连接线（%s）" % _detail)
        else:
            results.append("+0：主饼展开块与子饼之间绘有连接线（%s）" % _detail)

        # +1：未使用图例，类别名称直接标注在各扇区之上或紧邻扇区
        # 严格对齐细则，仅踩细则明确写出的点，不额外约束：
        #   ① 图表内没有 <c:legend>
        #   ② 类别名称落在数据标签上（showCatName 开启，或逐点富文本标签里带类别名）
        #
        # 与办公软件对齐：图表右上“+”号-取消勾选“图例”、勾选“数据标签”并在标签里
        #   包含“类别名称”，即示意图那种“名字写在扇区旁边”的样式。
        cat_on_label = _shows(facts, 'showCatName') or rich_ok
        legend_ok = (not facts['has_legend']) and cat_on_label
        _detail = "图例=%s、类别名标注在标签上=%s" % ('有' if facts['has_legend'] else '无',
                                                cat_on_label)
        if legend_ok:
            results.append("+1：未使用图例，类别名称直接标注在扇区上（%s）" % _detail)
        else:
            results.append("+0：未使用图例，类别名称直接标注在扇区上（%s）" % _detail)

        # +1：百分比标签为百分比格式（如 55%、34%、11%），不是 0.55、0.34 这类小数
        # 严格对齐细则，仅踩细则明确写出的点，不额外约束：
        #   ① 标签按“百分比”呈现：<c:showPercent val="1"/>（由软件按占比自动渲染成 55%）
        #   ② 或者显示数值、且标签自带百分比数字格式：
        #      <c:dLbls><c:numFmt formatCode="0%"/>（显式设定，不含从数值源继承的格式）
        #   ③ 或者逐点富文本标签文字本身就写着带 % 的百分数
        #
        # 与办公软件对齐：勾“百分比”时软件直接输出 55%；勾“值”时输出的是单元格的值，
        #   需在“设置数据标签格式-数字”里显式选百分比才算。判定口径取“显式的百分比
        #   机制”：数值源自带的 formatCode（numRef/numLit 里的 c:formatCode）虽然会被
        #   标签继承并显示成 55%，但不作为本项的得分依据。
        fmt_pool = [x for x in facts['lbl_num_fmts'] if x]
        fmt_pct = any('%' in x for x in fmt_pool)
        text_pct = bool(facts['lbl_texts']) and all('%' in t for t in facts['lbl_texts'].values())
        pct_ok = (_shows(facts, 'showPercent')
                  or (_shows(facts, 'showVal') and fmt_pct)
                  or text_pct)
        _detail = "showPercent=%s、标签自带数字格式=%s、数值源格式=%s" % (
            facts['show'].get('showPercent'), '/'.join(fmt_pool) or '无',
            facts['val_fmt'] or '无')
        if pct_ok:
            results.append("+1：百分比标签为百分比格式而非 0.55 这类小数（%s）" % _detail)
        else:
            results.append("+0：百分比标签为百分比格式而非 0.55 这类小数（%s）" % _detail)

        # +1：图表带有标题
        # 严格对齐细则，仅踩细则明确写出的点，不额外约束：
        #   ① 存在 <c:title> 且其中有非空可见文字
        #   ② 未被 <c:autoTitleDeleted val="1"/> 关掉
        #
        # 与办公软件对齐：图表顶部那行标题文字；勾掉“图表标题”后写入
        #   autoTitleDeleted=1，此时即使残留 title 节点也不显示。
        title_ok = bool(facts['title_text']) and not facts['auto_title_deleted']
        _shown = facts['title_text'].replace('\r', ' ').replace('\n', ' ') or '无'
        if title_ok:
            results.append("+1：图表带有标题（%s）" % _shown)
        else:
            results.append("+0：图表带有标题（标题文字=%s，autoTitleDeleted=%s）"
                           % (_shown, facts['auto_title_deleted']))

        # +1：主饼在左、子饼在右横向并排，子饼明显小于主饼，图表宽度不小于约 12 厘米，
        #     标签文字不互相重叠、不被图表边框截断
        # 严格对齐细则，仅踩细则明确写出的点，不额外约束：
        #   ① 主饼在左、子饼在右横向并排：复合饼图（ofPieChart）的固有版式即左主右子
        #   ② 子饼明显小于主饼：<c:secondPieSize> 不大于 90（%，相对主饼直径）
        #   ③ 图表宽度不小于约 12 厘米（EMU/360000，留 0.2 厘米容差）
        #   （“标签不重叠、不被边框截断”属于渲染期结果，文件中不存在可判定的落点信息，
        #     故不作为判分条件，只在此说明）
        #
        # 与办公软件对齐：secondPieSize 就是“设置数据系列格式”里的“第二绘图区大小”，
        #   默认 75%；图表宽度即拖动图表边框得到的 <xdr:ext cx> / <a:ext cx>。
        w_cm = target['w_cm']
        sps = facts['second_pie_size']
        sps_ok = (sps is None) or (sps <= SECOND_PIE_MAX)
        layout_ok = (facts['is_ofpie']
                     and w_cm >= MIN_CHART_WIDTH_CM - CHART_WIDTH_TOLERANCE_CM
                     and sps_ok)
        _detail = "图表宽度≈%.2f 厘米、secondPieSize=%s" % (w_cm, sps if sps is not None else '默认 75')
        if layout_ok:
            results.append("+1：左主右子横向并排、子饼明显更小、图表宽度不小于 12 厘米（%s）" % _detail)
        else:
            results.append("+0：左主右子横向并排、子饼明显更小、图表宽度不小于 12 厘米（%s）" % _detail)

        # ========== 减分项（命中即扣，可累计）==========
        data_sheet = sheets_by_name.get(DATA_SHEET_NAME.lower()) or sheets[0]
        allowed_names = set(GIVEN_NAMES)

        # -3：图表标题、类别名称或取数区域中出现与给定数据无关的自造名称
        #     （如“人员薪酬”“物料”“交通”“住宿”“餐饮”“总费用构成”“其他费用明细”等）
        # 严格对齐细则，仅踩细则明确写出的点，不额外约束：
        #   ① 图表标题里出现细则列举的自造名称特征词
        #   ② 图表类别名（含手改的数据标签文字）不在给定六项之内
        #   ③ 图表取数区域内的文本单元格不在给定六项之内
        #   （本项只看“图表标题 / 类别名称 / 取数区域”这三处——细则原文的范围；
        #     工作表上未被图表引用的其他残留文字不属于本项，由其他减分项覆盖）
        #
        # 与办公软件对齐：即图表标题栏文字、“选择数据”里的轴标签，以及轴标签/系列值
        #   所指向的那片单元格区域。
        bad_names = []
        for h in FABRICATED_HINTS:
            if h in facts['title_text']:
                bad_names.append('标题:%s' % h)
        for p in pts:
            if p['name'] and p['name'] not in allowed_names:
                bad_names.append('类别:%s' % p['name'])
        for txt in facts['lbl_texts'].values():
            for h in FABRICATED_HINTS:
                if h in txt:
                    bad_names.append('标签:%s' % h)
        for ref in (facts['cat_ref'], facts['val_ref']):
            shn, cells = _split_ref(ref)
            sh = sheets_by_name.get((shn or '').lower())
            if sh is None:
                continue
            for a1 in cells:
                c = sh.cell(a1)
                if c is not None and c['text'] and c['text'] not in allowed_names:
                    bad_names.append('取数区域 %s!%s:%s' % (sh.name, a1, c['text']))
        bad_names = sorted(set(bad_names))
        _detail = '、'.join(bad_names[:6]) or '无'
        if bad_names:
            results.append("-3：图表标题/类别名称/取数区域出现与给定数据无关的自造名称（%s）" % _detail)
        else:
            results.append("-0：图表标题/类别名称/取数区域出现与给定数据无关的自造名称（%s）" % _detail)

        # -1：Sheet2 中添加了未被要求的装饰性内容，如跨列合并的大标题文字块、
        #     “说明：……”提示文字块
        # 严格对齐细则，仅踩细则明确写出的点，不额外约束：
        #   ① Sheet2 上存在跨 3 列及以上的合并单元格且其中写着文字（“大标题文字块”）
        #   ② Sheet2 上存在“说明：/注：/提示：”这类提示文字
        #   （指令只要求在 Sheet2 放一个复合饼图，Sheet2 上的文字块都是多出来的）
        #
        # 与办公软件对齐：即打开 Sheet2 时除图表之外还能看到的那些文字。
        deco = []
        for mg in sheet2.merges:
            r1, c1, r2, c2 = _range_rc(mg)
            if (c2 - c1 + 1) >= 3:
                c = sheet2.cell('%s%d' % (_idx_to_col(c1), r1))
                if c is not None and c['text']:
                    deco.append('跨列合并标题 %s:%s' % (mg, c['text']))
        for ref, c in sheet2.cells.items():
            t = c['text'] or ''
            if any(k in t for k in ('说明：', '说明:', '注：', '注:', '提示：', '提示:')):
                deco.append('提示文字 %s:%s' % (ref, t[:20]))
        deco = sorted(set(deco))
        _detail = '、'.join(deco[:4]) or '无'
        if deco:
            results.append("-1：%s 中添加了未被要求的装饰性内容（%s）" % (SHEET2_NAME, _detail))
        else:
            results.append("-0：%s 中添加了未被要求的装饰性内容（%s）" % (SHEET2_NAME, _detail))

        # -1：Sheet2 中添加了数据中不存在的统计项，如“总费用合计”“其他明细合计”等求和行
        # 严格对齐细则，仅踩细则明确写出的点，不额外约束：
        #   ① Sheet2 上出现“合计/总计/小计/求和/总和”等统计项文字
        #   ② 或 Sheet2 上出现 SUM 求和公式（即凭空加出来的求和行）
        #
        # 与办公软件对齐：给定数据里两张小表本身就各自合计 100%，无需再补一行合计；
        #   Sheet2 上任何求和行都是多出来的。
        totals = []
        for ref, c in sheet2.cells.items():
            t = c['text'] or ''
            if any(k in t for k in TOTAL_ROW_HINTS):
                totals.append('%s:%s' % (ref, t[:20]))
            if c['f'] and 'SUM(' in c['f'].upper():
                totals.append('%s:=%s' % (ref, c['f'][:24]))
        totals = sorted(set(totals))
        _detail = '、'.join(totals[:4]) or '无'
        if totals:
            results.append("-1：%s 中添加了数据中不存在的统计项/求和行（%s）" % (SHEET2_NAME, _detail))
        else:
            results.append("-0：%s 中添加了数据中不存在的统计项/求和行（%s）" % (SHEET2_NAME, _detail))

        # -1：除 Sheet2 外新增了其他工作表，或已有工作表被改名为 Sheet1 之外的名称
        # 严格对齐细则，仅踩细则明确写出的点，不额外约束：
        #   ① 工作表名集合恰为 {Sheet1, Sheet2}（原有 Sheet1 + 指令要求新增的 Sheet2）
        #
        # 与办公软件对齐：底部工作表标签栏应只有 Sheet1、Sheet2 两个标签，
        #   多出的“数据源”“图表”“Sheet3”等标签，或把 Sheet1 改名，都算多余修改。
        names_now = [s.name for s in sheets]
        expected_names = {DATA_SHEET_NAME, SHEET2_NAME}
        extra_sheets = sorted(set(names_now) - expected_names)
        missing_sheets = sorted(expected_names - set(names_now))
        sheet_bad = bool(extra_sheets or missing_sheets)
        _detail = "现有工作表=%s" % '、'.join(names_now)
        if sheet_bad:
            results.append("-1：新增了其他工作表或已有工作表被改名（%s）" % _detail)
        else:
            results.append("-0：新增了其他工作表或已有工作表被改名（%s）" % _detail)

        # -3：“一级产品 55%”“二级产品 34%”“三级产品 11%”“优等品 15%”“合格品 65%”
        #     “残次品 20%”中任一名称或数值被删除、改写或被其他内容覆盖
        # 严格对齐细则，仅踩细则明确写出的点，不额外约束：
        #   ① 六项名称仍以单元格文本存在
        #   ② 每个名称右侧仍有与原值一致的数值单元格（名值成对存活）
        #   （新增的“折算区”不影响本项——本项只看原始两张小表的名与值是否还在）
        #
        # 与办公软件对齐：肉眼核对两张小表的每一行“名称 + 百分比”是否与原文件一致。
        given_pairs = {}
        for nm, expect in MAIN_ITEMS + SUB_ITEMS:
            found = None
            for sh in sheets:
                for ref in sh.find_text(nm):
                    for nb in sh.right_neighbors(ref, 3):
                        c = sh.cell(nb)
                        if c is not None and _near(c['num'], expect):
                            found = (sh, ref, nb)
                            break
                    if found:
                        break
                if found:
                    break
            given_pairs[nm] = found
        lost = [nm for nm, v in given_pairs.items() if v is None]
        _detail = "存活 %d/6，丢失或被改写：%s" % (6 - len(lost), '、'.join(lost) or '无')
        if lost:
            results.append("-3：给定六项名称或数值被删除/改写/覆盖（%s）" % _detail)
        else:
            results.append("-0：给定六项名称或数值被删除/改写/覆盖（%s）" % _detail)

        # -1：明细表左侧纵向合并的“一级产品”单元格（三行合并为一格）被取消合并或被拆分
        # 严格对齐细则，仅踩细则明确写出的点，不额外约束：
        #   ① 仍存在一处“单列、跨 3 行”的合并区
        #   ② 该合并区左上角单元格的文字为“一级产品”
        #
        # 与办公软件对齐：明细表左侧那一格纵向跨“优等品/合格品/残次品”三行的
        #   合并单元格，在 XML 中即 <mergeCell ref="A5:A7"/>。
        merge_ok = False
        merge_hits = []
        for sh in sheets:
            for mg in sh.merges:
                r1, c1, r2, c2 = _range_rc(mg)
                if c1 != c2 or (r2 - r1 + 1) != 3:
                    continue
                c = sh.cell('%s%d' % (_idx_to_col(c1), r1))
                if c is not None and c['text'] == EXPLODED_NAME:
                    merge_ok = True
                    merge_hits.append('%s!%s' % (sh.name, mg))
        _detail = "三行纵向合并区=%s" % ('、'.join(merge_hits) or '无')
        if merge_ok:
            results.append("-0：明细表左侧“一级产品”三行合并单元格被取消合并（%s）" % _detail)
        else:
            results.append("-1：明细表左侧“一级产品”三行合并单元格被取消合并（%s）" % _detail)

        # -1：两张数据表被搬移到其他工作表，或被图表压盖而无法直接查看
        # 严格对齐细则，仅踩细则明确写出的点，不额外约束：
        #   ① 六项名称仍都落在原数据工作表 Sheet1 上（没有被搬走）
        #   ② 该表上的数据区矩形没有被任何图表对象的锚定矩形压盖
        #
        # 与办公软件对齐：图表的 twoCellAnchor 给出它盖住的行列范围；若与数据表
        #   所在的行列范围相交，打开工作表时表格就被图表挡住了。
        on_data_sheet = []
        moved = []
        for nm in GIVEN_NAMES:
            refs = data_sheet.find_text(nm)
            if refs:
                on_data_sheet.extend(refs)
            else:
                moved.append(nm)
        covered = []
        if on_data_sheet:
            rows = [_ref_rc(r)[0] for r in on_data_sheet]
            cols = [_ref_rc(r)[1] for r in on_data_sheet]
            data_rect = (min(rows), min(cols), max(rows), max(cols) + 1)
            for ch in charts_all:
                if ch['sheet'] == data_sheet.name and _rects_overlap(ch['rect'], data_rect):
                    covered.append('图表锚定 %s' % str(ch['rect']))
        _detail = "不在 %s 上的名称=%s；压盖情况=%s" % (
            data_sheet.name, '、'.join(moved) or '无', '、'.join(covered) or '无')
        if moved or covered:
            results.append("-1：两张数据表被搬移到其他工作表或被图表压盖（%s）" % _detail)
        else:
            results.append("-0：两张数据表被搬移到其他工作表或被图表压盖（%s）" % _detail)

        # -1：比例数值被改为文本型，或由 0.55、0.15 这类小数改写为 55、15 这类数字，
        #     致数值含义改变
        # 严格对齐细则，仅踩细则明确写出的点，不额外约束：
        #   ① 名称右侧的比例单元格不是文本型（t="s"/"str"/"inlineStr" 承载的“55%”字样）
        #   ② 也不是被写成 55、15 这类放大 100 倍的数字
        #
        # 与办公软件对齐：单元格里存的应是 0.55 并靠“0%”格式显示成 55%；若直接存 55
        #   或存文本“55%”，参与计算时含义就变了（55 会被当成 5500%）。
        type_bad = []
        for nm, expect in MAIN_ITEMS + SUB_ITEMS:
            for sh in sheets:
                for ref in sh.find_text(nm):
                    for nb in sh.right_neighbors(ref, 3):
                        c = sh.cell(nb)
                        if c is None:
                            continue
                        if _near(c['num'], expect * 100.0):
                            type_bad.append('%s!%s=%g' % (sh.name, nb, c['num']))
                        elif c['num'] is None and c['text'] \
                                and re.match(r'^-?\d+(\.\d+)?%?$', c['text']):
                            type_bad.append('%s!%s 文本型“%s”' % (sh.name, nb, c['text']))
        type_bad = sorted(set(type_bad))
        _detail = '、'.join(type_bad[:6]) or '无'
        if type_bad:
            results.append("-1：比例数值被改为文本型或被放大 100 倍写成 55、15（%s）" % _detail)
        else:
            results.append("-0：比例数值被改为文本型或被放大 100 倍写成 55、15（%s）" % _detail)

        # -1：数据区单元格格式不满足——数值单元格为“0%”百分比格式、
        #     表头为深色填充配白色加粗字体、各单元格带边框
        # 严格对齐细则，仅踩细则明确写出的点，不额外约束：
        #   ① 六项比例单元格的数字格式为百分比（formatCode 含 %，内置 9 即 0%）
        #   ② 数据表内存在“深色填充 + 白色加粗字体”的表头单元格
        #   ③ 六项的名称与数值单元格都带边框
        #
        # 与办公软件对齐：选中比例列-“设置单元格格式-百分比”，表头那行深底白字加粗，
        #   整片区域套了“所有框线”。
        fmt_bad = []
        for nm, _expect in MAIN_ITEMS + SUB_ITEMS:
            found = given_pairs.get(nm)
            if found is None:
                fmt_bad.append('%s 名值已缺失' % nm)
                continue
            sh, nref, vref = found
            code = styles.numfmt_of(sh.cell(vref)['s'])
            if '%' not in code:
                fmt_bad.append('%s 数值格式=%s' % (nm, code))
            for r in (nref, vref):
                if not styles.has_border(sh.cell(r)['s']):
                    fmt_bad.append('%s!%s 无边框' % (sh.name, r))
        header_ok = False
        for _ref, c in data_sheet.cells.items():
            pat, fg = styles.fill_of(c['s'])
            bold, color, _fn = styles.font_of(c['s'])
            if (pat and pat != 'none' and fg and not _is_white(fg)
                    and bold and _is_white(color)):
                header_ok = True
                break
        if not header_ok:
            fmt_bad.append('未见深色填充+白色加粗字体的表头')
        _detail = '、'.join(sorted(set(fmt_bad))[:6]) or '百分比格式/表头样式/边框均满足'
        if fmt_bad:
            results.append("-1：数据区单元格格式不满足（%s）" % _detail)
        else:
            results.append("-0：数据区单元格格式不满足（%s）" % _detail)

        # -1：数据所在工作表的页边距不满足——左右各 0.75 英寸、上下各 1 英寸、
        #     页眉与页脚各 0.5 英寸
        # 严格对齐细则，仅踩细则明确写出的点，不额外约束：
        #   ① <pageMargins> 的 left/right/top/bottom/header/footer 六个值与期望一致
        #
        # 与办公软件对齐：页面布局-页边距对话框里的六个数字；该节点被整体丢弃时，
        #   打开文件会回落到软件默认页边距，同样算不满足。
        mg_bad = []
        for key, expect in sorted(EXPECTED_MARGINS.items()):
            got = _to_float(data_sheet.margins.get(key))
            if got is None or abs(got - expect) > 0.005:
                mg_bad.append('%s=%s(应%s)' % (key, data_sheet.margins.get(key) or '缺失', expect))
        _detail = '、'.join(mg_bad) or '六项页边距均一致'
        if mg_bad:
            results.append("-1：%s 的页边距被改动（%s）" % (data_sheet.name, _detail))
        else:
            results.append("-0：%s 的页边距被改动（%s）" % (data_sheet.name, _detail))

        # -1：工作簿主题字体方案不满足——标题字体 Calibri Light、正文字体 Calibri，
        #     且中文（Hans）映射为宋体、日文映射为 ＭＳ Ｐゴシック 等多语言字体表完整
        #     （被整体替换为 latin/ea/cs 均为 Calibri 的简化字体表即计）
        # 严格对齐细则，仅踩细则明确写出的点，不额外约束：
        #   ① majorFont 的 latin typeface 为 Calibri Light、minorFont 为 Calibri
        #   ② 主题里保留 <a:font script="Hans" typeface="宋体"/>
        #   ③ 主题里保留 <a:font script="Jpan" typeface="ＭＳ Ｐゴシック"/>
        #
        # 与办公软件对齐：页面布局-字体所选主题字体方案；多语言字体表决定中日文
        #   文本回退用哪种字体，被简化成 latin/ea/cs 全 Calibri 后中文显示会变。
        theme_part = 'xl/theme/theme1.xml'
        if not pkg.has(theme_part):
            theme_part = next((n for n in sorted(pkg.names)
                              if n.startswith('xl/theme/') and n.endswith('.xml')), None)
        theme_bad = []
        if not theme_part:
            theme_bad.append('主题部件缺失')
        else:
            traw = pkg.text(theme_part)
            m = re.search(r'<a:fontScheme.*?</a:fontScheme>', traw, re.S)
            fs = m.group(0) if m else ''
            major = re.search(r'<a:majorFont>\s*<a:latin[^>]*typeface="([^"]*)"', fs)
            minor = re.search(r'<a:minorFont>\s*<a:latin[^>]*typeface="([^"]*)"', fs)
            if not major or major.group(1) != 'Calibri Light':
                theme_bad.append('标题字体=%s' % (major.group(1) if major else '缺失'))
            if not minor or minor.group(1) != 'Calibri':
                theme_bad.append('正文字体=%s' % (minor.group(1) if minor else '缺失'))
            if not re.search(r'script="Hans"\s+typeface="宋体"', fs):
                theme_bad.append('Hans→宋体 映射丢失')
            if not re.search(r'script="Jpan"\s+typeface="ＭＳ Ｐゴシック"', fs):
                theme_bad.append('Jpan→ＭＳ Ｐゴシック 映射丢失')
        _detail = '、'.join(theme_bad) or '主题字体方案与多语言字体表完整'
        if theme_bad:
            results.append("-1：工作簿主题字体方案被简化或替换（%s）" % _detail)
        else:
            results.append("-0：工作簿主题字体方案被简化或替换（%s）" % _detail)
    finally:
        pkg.close()

    # ========== 将 results 反解为结构化 dim2_items ==========
    # results 与 _MAX_DELTAS_ORDER 严格一一对应：加分项以 "+N：" 开头，
    # 减分项以 "-N：" 开头；未命中时分别写作 "+0：" / "-0："。
    if len(results) != len(_MAX_DELTAS_ORDER):
        raise RuntimeError(f"dim2 项数不一致：results={len(results)} vs order={len(_MAX_DELTAS_ORDER)}")
    dim2_items = []
    total_score = 0
    for raw, max_delta in zip(results, _MAX_DELTAS_ORDER):
        head, _, rule = raw.partition("：")
        try:
            delta = int(head.replace("+", "").strip())
        except ValueError:
            delta = 0
        hit = delta != 0
        dim2_items.append({"rule": rule, "max_delta": max_delta,
                           "delta": max_delta if hit else 0, "hit": hit})
        total_score += (max_delta if hit else 0)
    result["dim2_items"] = dim2_items
    result["total_score"] = total_score
    result["max_score"] = MAX_SCORE_TOTAL
    return result


if __name__ == "__main__":
    if len(sys.argv) >= 2:
        _dir = sys.argv[1]
    else:
        _dir = os.path.dirname(os.path.abspath(__file__))
    _out = evaluate(_dir)
    print(json.dumps(_out, ensure_ascii=False, indent=2))
