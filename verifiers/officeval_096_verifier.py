#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# pyright: reportUnknownMemberType=false, reportUnknownVariableType=false, reportUnknownArgumentType=false, reportUnknownParameterType=false, reportMissingParameterType=false, reportPrivateUsage=false, reportUnknownLambdaType=false, reportGeneralTypeIssues=false

import os
import re
import sys
import json
import zipfile
import xml.etree.ElementTree as ET

# 说明：按接口统一约定，本模块不修改全局 sys.stdout，也不 print 主结果。
# 主结果统一由 evaluate(dir_path) -> dict 返回；入参为“脚本所在目录的路径”，
# 由脚本自己在该目录内定位并打开被评估的 .xlsx。本文件仅在 __main__ 下作为
# 本地调试入口，才把 dict 序列化为 JSON 打印。
#
# 任务：按输入图片新建 Excel——上方“11～20 岁身高统计表（可直接填写数据）”，
# 下方仅有数据点、无折线的统计图；横轴 11～20 岁，纵轴 0-200cm、每 50 一格；
# 表中一输入数据，下方图表自动显示对应数据点。
# 本任务为“从零生成”类：输入仅一张参考图片，无可被破坏的原始文件，故 dim2
# 只设加分项——命中记 max_delta，未命中记 0，total_score ∈ [0, +49]；
# “惹麻烦”（预填占位数据、残留辅助区、限制与 0-200cm 矛盾等）以对应加分项
# 不得分的方式体现，不另设负分。
#
# 实现口径：不依赖 openpyxl（其对 chart 的读回有损），直接解包 xlsx 读 XML，
# 判定的是“用户打开办公软件真正看到/用到的效果”。

SCRIPT_ID = "096"

# 维度二加分项按“结果追加顺序”的 max_delta；末端把 results 里的 "+N：..."
# 字符串反解为结构化 dim2_items 时用。顺序须与 _evaluate_impl 中
# results.append(...) 的顺序严格一致；共 23 项，累加 = MAX_SCORE_TOTAL。
_ADD_DELTAS_ORDER = [
    3,  # A01 版面：表在上/图在下，图表尺寸适中且未压盖统计表
    3,  # A02 交付整洁：无可见的中间辅助数据区
    1,  # A03 统计表标题（11-20 岁身高统计表）
    3,  # A04 年龄行 11、12…20 共 10 个，完整有序
    3,  # A05 身高/cm 行：10 个与年龄一一对应的可填写单元格
    3,  # A06 身高单元格留空交付，无需先清除示例数据
    1,  # A07 表格边框/内容居中/列宽足够显示
    1,  # A08 无与 0-200cm 矛盾的限制或提示
    3,  # A09 XY 散点图且数据点间无连线
    1,  # A10 数据点标记清晰可见
    1,  # A11 数据点显示数值标签
    1,  # A12 图表标题
    3,  # A13 横轴：最小 11、最大 20、单位 1
    3,  # A14 纵轴：最小 0、最大 200
    1,  # A15 纵轴主要刻度单位 50
    1,  # A16 横轴标题“年龄”
    1,  # A17 纵轴标题“身高/cm”
    1,  # A18 绘图区主要网格线
    5,  # A19 Y 值系列联动统计表身高单元格
    3,  # A20 X 值系列引用统计表年龄单元格（非常量数组）
    3,  # A21 空值不产生 0 值假数据点
    3,  # A22 10 个年龄全部纳入联动区域
    1,  # A23 工作簿自动重算
]
MAX_SCORE_TOTAL = sum(_ADD_DELTAS_ORDER)   # +49：完成度上限
assert MAX_SCORE_TOTAL == 49
_MAX_DELTAS_ORDER = _ADD_DELTAS_ORDER

# ---------- 命名空间 ----------
NS_MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
NS_C = 'http://schemas.openxmlformats.org/drawingml/2006/chart'
NS_A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
NS_XDR = 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing'
NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
NS_PR = 'http://schemas.openxmlformats.org/package/2006/relationships'

# 指令要求的年龄序列与纵轴口径
_AGES = list(range(11, 21))
_Y_MAX = 200.0
_Y_UNIT = 50.0
_X_MIN, _X_MAX, _X_UNIT = 11.0, 20.0, 1.0
# 图表尺寸“适中”区间（EMU；1cm = 360000 EMU）：约 5cm~40cm 宽、3cm~30cm 高
_SIZE_MIN_CX, _SIZE_MAX_CX = 1800000, 14400000
_SIZE_MIN_CY, _SIZE_MAX_CY = 1080000, 10800000

_ERR_VALUES = ('#REF!', '#VALUE!', '#NAME?', '#DIV/0!', '#NULL!', '#NUM!')
_CELL_RE = re.compile(r'\$?([A-Z]{1,3})\$?(\d{1,7})')


def q(ns, tag):
    return '{%s}%s' % (ns, tag)


def _col_to_num(col):
    """列字母 -> 1 起的列号。"""
    n = 0
    for ch in col:
        n = n * 26 + (ord(ch) - 64)
    return n


def _num_to_col(n):
    """1 起的列号 -> 列字母。"""
    s = ''
    while n > 0:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def _split_ref(ref):
    """'B12' -> ('B', 12)；解析失败返回 (None, None)。"""
    m = re.match(r'^\$?([A-Z]{1,3})\$?(\d{1,7})$', (ref or '').upper())
    if not m:
        return None, None
    return m.group(1), int(m.group(2))


def _norm_ref(ref):
    """去掉 $ 并大写：'$B$12' -> 'B12'。"""
    c, r = _split_ref(ref)
    return None if c is None else '%s%d' % (c, r)


def _expand_area(area):
    """'$B$3:$K$3' / 'B3' -> ['B3', 'C3', ...]（单区域，上限 4096 格）。"""
    area = (area or '').replace('$', '').upper()
    if not area:
        return []
    if ':' not in area:
        one = _norm_ref(area)
        return [one] if one else []
    a, _, b = area.partition(':')
    c1, r1 = _split_ref(a)
    c2, r2 = _split_ref(b)
    if c1 is None or c2 is None:
        return []
    n1, n2 = _col_to_num(c1), _col_to_num(c2)
    lo_c, hi_c = min(n1, n2), max(n1, n2)
    lo_r, hi_r = min(r1, r2), max(r1, r2)
    if (hi_c - lo_c + 1) * (hi_r - lo_r + 1) > 4096:
        return []
    out = []
    for r in range(lo_r, hi_r + 1):
        for c in range(lo_c, hi_c + 1):
            out.append('%s%d' % (_num_to_col(c), r))
    return out


def _parse_formula_ref(f):
    """解析图表数据源公式 -> (sheet_name 或 None, [单元格引用...])。
    形如 Sheet1!$B$3:$K$3 / '身高记录'!$B$13:$B$22 / (A1:A5,B1:B5)。"""
    if not f:
        return None, []
    sheet = None
    cells = []
    for part in str(f).replace('(', '').replace(')', '').split(','):
        part = part.strip()
        if not part:
            continue
        if '!' in part:
            sname, _, area = part.rpartition('!')
            sname = sname.strip()
            if sname.startswith("'") and sname.endswith("'"):
                sname = sname[1:-1].replace("''", "'")
            if sheet is None:
                sheet = sname
        else:
            area = part
        cells.extend(_expand_area(area))
    return sheet, cells


class _Sheet(object):
    """一张工作表的“用户可见事实”：单元格值/公式/样式、合并区、数据有效性、
    行列隐藏与列宽、以及挂在该表上的图表。"""

    def __init__(self, name):
        self.name = name
        self.cells = {}          # 'A1' -> {'text','num','f','s'}
        self.merged = []         # ['A1:K1', ...]
        self.dv = []             # [{'sqref','type','f1','f2'}]
        self.col_width = {}      # 列号 -> 宽度
        self.col_hidden = set()  # 列号
        self.row_hidden = set()  # 行号
        self.default_col_width = 8.43
        self.charts = []         # [chart dict]
        self.anchors = []        # [{'from_row','to_row','cx','cy','has_chart'}]
        self.pictures = 0

    def text(self, ref):
        c = self.cells.get(_norm_ref(ref) or '')
        return (c or {}).get('text', '') or ''

    def num(self, ref):
        c = self.cells.get(_norm_ref(ref) or '')
        return (c or {}).get('num')

    def formula(self, ref):
        c = self.cells.get(_norm_ref(ref) or '')
        return (c or {}).get('f')

    def width_of(self, col_letter):
        n = _col_to_num(col_letter)
        return self.col_width.get(n, self.default_col_width)


class _Xlsx(object):
    """解包并解析 xlsx：工作簿 / 共享字符串 / 样式 / 工作表 / 绘图 / 图表。"""

    def __init__(self, path):
        self.path = path
        self.zip = zipfile.ZipFile(path)
        self.names = set(self.zip.namelist())
        self.shared = []
        self.cell_xfs = []       # [{'borderId','halign','fillId'}]
        self.border_has_edge = []
        self.fill_is_yellow = []
        self.sheets = []         # [_Sheet]
        self.calc_mode = 'auto'
        self.chart_parts = 0
        self._load()

    # ---------- 基础 ----------
    def _xml(self, name):
        if name not in self.names:
            return None
        try:
            return ET.fromstring(self.zip.read(name))
        except ET.ParseError:
            return None

    def _rels_of(self, part):
        """返回 part 的 rId -> 目标绝对路径 映射。"""
        d = os.path.dirname(part)
        rels = '%s/_rels/%s.rels' % (d, os.path.basename(part))
        root = self._xml(rels)
        out = {}
        if root is None:
            return out
        for rel in root.findall(q(NS_PR, 'Relationship')):
            tgt = rel.get('Target') or ''
            rid = rel.get('Id')
            if tgt.startswith('/'):
                full = tgt.lstrip('/')
            else:
                full = os.path.normpath(os.path.join(d, tgt)).replace('\\', '/')
            out[rid] = full
        return out

    # ---------- 装载 ----------
    def _load(self):
        self.chart_parts = len([n for n in self.names
                                if re.match(r'xl/(charts|drawings/charts)/chart\d*\.xml$', n)])
        self._load_shared()
        self._load_styles()
        wb = self._xml('xl/workbook.xml')
        if wb is None:
            raise ValueError('缺少 xl/workbook.xml，非有效 xlsx 工作簿')
        calc = wb.find(q(NS_MAIN, 'calcPr'))
        if calc is not None and (calc.get('calcMode') or 'auto') == 'manual':
            self.calc_mode = 'manual'
        rels = self._rels_of('xl/workbook.xml')
        holder = wb.find(q(NS_MAIN, 'sheets'))
        for sh in (holder if holder is not None else []):
            rid = sh.get(q(NS_R, 'id'))
            target = rels.get(rid)
            sheet = _Sheet(sh.get('name') or '')
            if target and target in self.names:
                self._load_sheet(sheet, target)
            self.sheets.append(sheet)

    def _load_shared(self):
        root = self._xml('xl/sharedStrings.xml')
        if root is None:
            return
        for si in root.findall(q(NS_MAIN, 'si')):
            self.shared.append(''.join(t.text or '' for t in si.iter(q(NS_MAIN, 't'))))

    def _load_styles(self):
        root = self._xml('xl/styles.xml')
        if root is None:
            return
        borders = root.find(q(NS_MAIN, 'borders'))
        for b in (borders if borders is not None else []):
            has = False
            for edge in ('left', 'right', 'top', 'bottom'):
                e = b.find(q(NS_MAIN, edge))
                if e is not None and e.get('style'):
                    has = True
            self.border_has_edge.append(has)
        fills = root.find(q(NS_MAIN, 'fills'))
        for fl in (fills if fills is not None else []):
            rgb = ''
            pf = fl.find(q(NS_MAIN, 'patternFill'))
            if pf is not None:
                fg = pf.find(q(NS_MAIN, 'fgColor'))
                if fg is not None:
                    rgb = (fg.get('rgb') or '').upper()
            # 判定“黄色底”：R、G 高且 B 低
            yellow = False
            if len(rgb) == 8:
                r, g, b2 = int(rgb[2:4], 16), int(rgb[4:6], 16), int(rgb[6:8], 16)
                yellow = r > 200 and g > 180 and b2 < 150
            self.fill_is_yellow.append(yellow)
        xfs = root.find(q(NS_MAIN, 'cellXfs'))
        for xf in (xfs if xfs is not None else []):
            al = xf.find(q(NS_MAIN, 'alignment'))
            self.cell_xfs.append({
                'borderId': int(xf.get('borderId') or 0),
                'fillId': int(xf.get('fillId') or 0),
                'halign': (al.get('horizontal') if al is not None else None) or '',
            })

    def style_of(self, s_idx):
        try:
            return self.cell_xfs[int(s_idx)]
        except (TypeError, ValueError, IndexError):
            return {'borderId': 0, 'fillId': 0, 'halign': ''}

    def has_border(self, s_idx):
        bid = self.style_of(s_idx)['borderId']
        return bool(bid) and bid < len(self.border_has_edge) and self.border_has_edge[bid]

    def has_yellow_fill(self, s_idx):
        fid = self.style_of(s_idx)['fillId']
        return bool(fid) and fid < len(self.fill_is_yellow) and self.fill_is_yellow[fid]

    def _load_sheet(self, sheet, part):
        root = self._xml(part)
        if root is None:
            return
        fmt = root.find(q(NS_MAIN, 'sheetFormatPr'))
        if fmt is not None:
            for k in ('defaultColWidth', 'baseColWidth'):
                if fmt.get(k):
                    try:
                        sheet.default_col_width = float(fmt.get(k))
                        break
                    except ValueError:
                        pass
        cols = root.find(q(NS_MAIN, 'cols'))
        for col in (cols if cols is not None else []):
            try:
                lo, hi = int(col.get('min')), int(col.get('max'))
            except (TypeError, ValueError):
                continue
            for n in range(lo, min(hi, lo + 256) + 1):
                if col.get('width'):
                    try:
                        sheet.col_width[n] = float(col.get('width'))
                    except ValueError:
                        pass
                if (col.get('hidden') or '0') in ('1', 'true'):
                    sheet.col_hidden.add(n)
        data = root.find(q(NS_MAIN, 'sheetData'))
        for row in (data if data is not None else []):
            try:
                r_idx = int(row.get('r'))
            except (TypeError, ValueError):
                r_idx = None
            if r_idx is not None and (row.get('hidden') or '0') in ('1', 'true'):
                sheet.row_hidden.add(r_idx)
            for c in row.findall(q(NS_MAIN, 'c')):
                ref = _norm_ref(c.get('r') or '')
                if not ref:
                    continue
                sheet.cells[ref] = self._cell_of(c)
        mc = root.find(q(NS_MAIN, 'mergeCells'))
        for m in (mc if mc is not None else []):
            if m.get('ref'):
                sheet.merged.append(m.get('ref'))
        dvs = root.find(q(NS_MAIN, 'dataValidations'))
        for dv in (dvs if dvs is not None else []):
            f1 = dv.find(q(NS_MAIN, 'formula1'))
            f2 = dv.find(q(NS_MAIN, 'formula2'))
            sheet.dv.append({
                'sqref': dv.get('sqref') or '',
                'type': dv.get('type') or '',
                'operator': dv.get('operator') or '',
                'f1': (f1.text if f1 is not None else None),
                'f2': (f2.text if f2 is not None else None),
            })
        dr = root.find(q(NS_MAIN, 'drawing'))
        if dr is not None:
            target = self._rels_of(part).get(dr.get(q(NS_R, 'id')))
            if target:
                self._load_drawing(sheet, target)

    def _cell_of(self, c):
        t = c.get('t') or 'n'
        v = c.find(q(NS_MAIN, 'v'))
        f = c.find(q(NS_MAIN, 'f'))
        raw = v.text if v is not None else None
        text, num = '', None
        if t == 's':
            try:
                text = self.shared[int(raw)]
            except (TypeError, ValueError, IndexError):
                text = ''
        elif t == 'inlineStr':
            is_el = c.find(q(NS_MAIN, 'is'))
            text = '' if is_el is None else ''.join(x.text or '' for x in is_el.iter(q(NS_MAIN, 't')))
        elif raw is not None:
            text = raw
            try:
                num = float(raw)
            except ValueError:
                num = None
        if t in ('str', 'e', 'b'):
            num = None
        f_txt = None if f is None else (f.text or '')
        return {'text': text or '', 'num': num, 'f': f_txt, 't': t, 's': c.get('s')}

    def _load_drawing(self, sheet, part):
        root = self._xml(part)
        if root is None:
            return
        rels = self._rels_of(part)
        for tag in ('twoCellAnchor', 'oneCellAnchor', 'absoluteAnchor'):
            for a in root.findall(q(NS_XDR, tag)):
                item = {'from_row': None, 'to_row': None, 'cx': None, 'cy': None,
                        'has_chart': False}
                frm = a.find(q(NS_XDR, 'from'))
                if frm is not None:
                    rw = frm.find(q(NS_XDR, 'row'))
                    if rw is not None and (rw.text or '').strip().isdigit():
                        item['from_row'] = int(rw.text)
                to = a.find(q(NS_XDR, 'to'))
                if to is not None:
                    rw = to.find(q(NS_XDR, 'row'))
                    if rw is not None and (rw.text or '').strip().isdigit():
                        item['to_row'] = int(rw.text)
                ext = a.find(q(NS_XDR, 'ext'))
                if ext is None:
                    ext = a.find('.//' + q(NS_A, 'ext'))
                if ext is not None:
                    try:
                        item['cx'] = int(ext.get('cx'))
                        item['cy'] = int(ext.get('cy'))
                    except (TypeError, ValueError):
                        pass
                ch = a.find('.//' + q(NS_C, 'chart'))
                if ch is not None:
                    item['has_chart'] = True
                    target = rels.get(ch.get(q(NS_R, 'id')))
                    parsed = self._load_chart(target) if target else None
                    if parsed is not None:
                        parsed['anchor'] = item
                        sheet.charts.append(parsed)
                if a.find('.//' + q(NS_XDR, 'pic')) is not None:
                    sheet.pictures += 1
                sheet.anchors.append(item)

    # ---------- 图表 ----------
    def _load_chart(self, part):
        root = self._xml(part)
        if root is None:
            return None
        chart = root.find(q(NS_C, 'chart'))
        if chart is None:
            return None
        plot = chart.find(q(NS_C, 'plotArea'))
        if plot is None:
            return None
        out = {
            'part': part,
            'title': _rich_text(chart.find(q(NS_C, 'title'))),
            'title_deleted': _val(chart.find(q(NS_C, 'autoTitleDeleted'))) == '1',
            'disp_blanks': _val(chart.find(q(NS_C, 'dispBlanksAs'))) or '',
            'groups': [],
            'series': [],
            'axes': [],
        }
        for g in list(plot):
            local = g.tag.split('}')[-1]
            if not local.endswith('Chart'):
                continue
            out['groups'].append(local)
            grp_showval = _val(g.find(q(NS_C, 'dLbls') + '/' + q(NS_C, 'showVal')))
            scatter_style = _val(g.find(q(NS_C, 'scatterStyle')))
            for ser in g.findall(q(NS_C, 'ser')):
                out['series'].append(self._ser_of(ser, local, scatter_style, grp_showval))
        for local in ('valAx', 'catAx', 'dateAx', 'serAx'):
            for ax in plot.findall(q(NS_C, local)):
                sc = ax.find(q(NS_C, 'scaling'))
                out['axes'].append({
                    'kind': local,
                    'pos': _val(ax.find(q(NS_C, 'axPos'))) or '',
                    'deleted': _val(ax.find(q(NS_C, 'delete'))) == '1',
                    'min': _fval(sc.find(q(NS_C, 'min'))) if sc is not None else None,
                    'max': _fval(sc.find(q(NS_C, 'max'))) if sc is not None else None,
                    'major_unit': _fval(ax.find(q(NS_C, 'majorUnit'))),
                    'title': _rich_text(ax.find(q(NS_C, 'title'))),
                    'gridlines': ax.find(q(NS_C, 'majorGridlines')) is not None,
                })
        return out

    def _ser_of(self, ser, group, scatter_style, grp_showval):
        s = {'group': group, 'scatter_style': scatter_style or '',
             'no_line': False, 'marker': None, 'show_val': grp_showval or '0'}
        sp = ser.find(q(NS_C, 'spPr'))
        if sp is not None:
            ln = sp.find(q(NS_A, 'ln'))
            if ln is not None:
                if ln.find(q(NS_A, 'noFill')) is not None:
                    s['no_line'] = True
                else:
                    try:
                        if int(ln.get('w') or -1) == 0:
                            s['no_line'] = True
                    except ValueError:
                        pass
        mk = ser.find(q(NS_C, 'marker'))
        if mk is not None:
            s['marker'] = _val(mk.find(q(NS_C, 'symbol'))) or 'auto'
        own = _val(ser.find(q(NS_C, 'dLbls') + '/' + q(NS_C, 'showVal')))
        if own is not None:
            s['show_val'] = own
        for key, tag in (('x', 'xVal'), ('y', 'yVal'), ('cat', 'cat'), ('val', 'val')):
            node = ser.find(q(NS_C, tag))
            if node is None:
                continue
            ref = node.find(q(NS_C, 'numRef'))
            lit = node.find(q(NS_C, 'numLit'))
            if ref is None and lit is None:
                ref = node.find(q(NS_C, 'strRef'))
                lit = node.find(q(NS_C, 'strLit'))
            f = None if ref is None else _text(ref.find(q(NS_C, 'f')))
            cnt = None
            holder = ref if ref is not None else lit
            if holder is not None:
                pc = holder.find('.//' + q(NS_C, 'ptCount'))
                cnt = _fval(pc)
            s[key + '_f'] = f
            s[key + '_is_lit'] = (ref is None and lit is not None)
            s[key + '_count'] = int(cnt) if cnt is not None else 0
        return s


def _val(el):
    return None if el is None else (el.get('val') or None)


def _fval(el):
    v = _val(el)
    if v is None:
        return None
    try:
        return float(v)
    except ValueError:
        return None


def _text(el):
    return None if el is None else (el.text or '')


def _rich_text(el):
    """取 c:title / c:tx 里 a:t 的拼接文本。"""
    if el is None:
        return ''
    return ''.join(t.text or '' for t in el.iter(q(NS_A, 't'))).strip()


def evaluate(dir_path):
    """按接口约定返回结构化 dict；不 print、不改 sys.stdout、不 sys.exit。
    入参 dir_path 为“脚本所在目录的路径”，脚本自己在该目录里定位并打开被评估的
    .xlsx（忽略以 ~$ 打头的 Excel 临时锁文件）。

    未找到待评文件（目录不存在或目录内没有工作簿）时，返回 status='error'、
    file_name 与 dim1_pass/dim1_reason 均留空、dim2_items 为空、分值全 0；
    其他运行异常则给出“脚本运行异常：…”摘要，并按维度一未通过计 0 分。"""
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
            raise FileNotFoundError(f"目录不存在或不是目录：{dir_path}")
        cands = sorted(
            f for f in os.listdir(dir_path)
            if f.lower().endswith(('.xlsx', '.xlsm', '.xls')) and not f.startswith('~$')
            and os.path.isfile(os.path.join(dir_path, f))
        )
        if not cands:
            raise FileNotFoundError(f"目录中未找到 .xlsx 文件：{dir_path}")
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
        result["error"] = f"脚本运行异常：{type(exc).__name__}: {exc}"
        result["dim1_pass"] = False
        result["dim1_reason"] = (f"维度一未通过：文件无法打开或解析"
                                f"（{type(exc).__name__}: {exc}），完成度总分记 0")
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
    return result


def _evaluate_impl(file_path, result):
    """维度一门禁 + 维度二加分项评估主体；对内构造 "+N：..." 结果串，
    末尾按 _MAX_DELTAS_ORDER 反解成结构化 dim2_items 填入 result。"""
    if not zipfile.is_zipfile(file_path):
        return _fail_dim1(result, "文件非 OOXML 工作簿（无法解包），Excel 不可正常打开")
    wb = _Xlsx(file_path)
    results = []

    # ========== 维度1：可用与可修改性 ==========
    if not wb.sheets:
        return _fail_dim1(result, "工作簿内无工作表，文件不可用")
    filled = sum(len(s.cells) for s in wb.sheets)
    if filled == 0:
        return _fail_dim1(result, "工作簿内无任何单元格内容，交付物为空")
    err_cells = 0
    for s in wb.sheets:
        for c in s.cells.values():
            if c['t'] == 'e' or (c['text'] or '').strip() in _ERR_VALUES:
                err_cells += 1
    if err_cells > 3:
        return _fail_dim1(result, f"公式大面积报错（{err_cells} 个错误值单元格），交付物不可用")
    if wb.chart_parts == 0:
        pics = sum(s.pictures for s in wb.sheets)
        if pics > 0:
            return _fail_dim1(result, "统计图为图片/截图对象而非可编辑的原生图表")
        return _fail_dim1(result, "工作簿内不存在任何原生图表对象，无统计图可评")
    result["dim1_pass"] = True
    result["dim1_reason"] = "文件可正常打开，统计表为可编辑单元格、统计图为原生图表对象，未见公式报错"

    # ========== 公共定位：统计表所在工作表、年龄行、身高行 ==========
    def _locate_table():
        """返回 (sheet, age_row, age_cols, height_row, best_score)。
        年龄行 = 含“年龄”标签且横向连续出现 11…20 的行；身高行 = 含“身高”标签、
        且与年龄列一一对应的行。找不到完整结构时返回尽可能多的定位信息。"""
        best = (None, None, [], None, -1)
        for sh in wb.sheets:
            rows = {}
            for ref, c in sh.cells.items():
                col, r = _split_ref(ref)
                if col is None:
                    continue
                rows.setdefault(r, []).append((_col_to_num(col), col, c))
            for r, items in sorted(rows.items()):
                items.sort()
                has_label = any('年龄' in (c['text'] or '') for (_, _, c) in items)
                # 该行按列序出现的、落在 11~20 的整数
                got = [(int(c['num']), col) for (_, col, c) in items
                       if c['num'] is not None and float(c['num']).is_integer()
                       and 11 <= c['num'] <= 20]
                if not has_label and not got:
                    continue
                ages = [v for v, _ in got]
                score = (10 if ages == _AGES else len(set(ages) & set(_AGES)))
                score += 2 if has_label else 0
                if score <= best[4]:
                    continue
                age_cols = [col for _, col in got]
                h_row = None
                for r2, items2 in sorted(rows.items()):
                    if r2 <= r:
                        continue
                    if any('身高' in (c['text'] or '') for (_, _, c) in items2):
                        h_row = r2
                        break
                best = (sh, r, age_cols, h_row, score)
        return best

    tsheet, age_row, age_cols, height_row, _ = _locate_table()
    age_cells = ['%s%d' % (c, age_row) for c in age_cols] if (tsheet and age_row) else []
    height_cells = (['%s%d' % (c, height_row) for c in age_cols]
                    if (tsheet and height_row and age_cols) else [])
    ages_got = [int(tsheet.num(ref)) for ref in age_cells] if tsheet else []

    # 取被评图表：优先统计表所在工作表上的第一张图表
    chart = None
    chart_sheet = None
    for sh in ([tsheet] if tsheet else []) + wb.sheets:
        if sh is not None and sh.charts:
            chart, chart_sheet = sh.charts[0], sh
            break
    series = (chart or {}).get('series') or []
    axes = [a for a in ((chart or {}).get('axes') or []) if not a['deleted']]
    x_axes = [a for a in axes if a['pos'] in ('b', 't')]
    y_axes = [a for a in axes if a['pos'] in ('l', 'r')]
    x_ax = x_axes[0] if x_axes else None
    y_ax = y_axes[0] if y_axes else None

    def _sheet_by_name(name):
        if not name:
            return chart_sheet or tsheet
        for sh in wb.sheets:
            if sh.name == name:
                return sh
        return None

    def _reach(f, targets, target_sheet):
        """判断图表数据源公式 f 能“活引用”到 targets（同表单元格列表）的哪些格：
        ① 直接落在 targets 上；② 引用的是同/异表的中间格，而该格公式又引用了
        targets（放行一跳辅助列，用户在表里输入后图表仍会自动刷新）。
        返回可达的 targets 集合。"""
        if not f or not targets or target_sheet is None:
            return set()
        sname, cells = _parse_formula_ref(f)
        src = _sheet_by_name(sname)
        if src is None or not cells:
            return set()
        hit = set()
        tset = set(targets)
        for ref in cells:
            if src is target_sheet and ref in tset:
                hit.add(ref)
                continue
            fml = src.formula(ref)
            if not fml:
                continue
            for m in _CELL_RE.finditer(fml.upper()):
                cand = '%s%d' % (m.group(1), int(m.group(2)))
                if cand in tset:
                    hit.add(cand)
        return hit

    # ========== 维度2：完成度加分项 ==========

    # +3：A01 版面为“统计表在上、统计图在下”，图表尺寸适中且未压盖统计表
    #     注：辅助数据被图表压盖的情形由 A02 判定，此处不重复扣。
    table_last_row = 0
    if tsheet and age_row:
        table_last_row = max(age_row, height_row or 0)
    anchor = (chart or {}).get('anchor') or {}
    layout_ok = False
    if chart is not None and chart_sheet is tsheet and table_last_row and anchor:
        from_row_1b = None if anchor['from_row'] is None else anchor['from_row'] + 1
        below_ok = from_row_1b is not None and from_row_1b > table_last_row
        cx, cy = anchor.get('cx'), anchor.get('cy')
        if cx and cy:
            size_ok = (_SIZE_MIN_CX <= cx <= _SIZE_MAX_CX and _SIZE_MIN_CY <= cy <= _SIZE_MAX_CY)
        elif anchor.get('to_row') is not None and anchor['from_row'] is not None:
            size_ok = (anchor['to_row'] - anchor['from_row']) >= 8
        else:
            size_ok = False
        layout_ok = bool(below_ok and size_ok)
    if layout_ok:
        results.append("+3：版面为统计表在上、统计图在下，图表尺寸适中且未压盖统计表")
    else:
        results.append("+0：版面为统计表在上、统计图在下，图表尺寸适中且未压盖统计表")

    # +3：A02 交付整洁——统计表所在工作表上无可见的中间辅助数据区
    #     口径：除统计表区域外，凡带公式或数值的单元格即视为辅助数据；
    #     置于独立工作表或真正隐藏的行/列可放行，白字、被图表压盖不算隐藏。
    table_region = set(age_cells) | set(height_cells)
    if tsheet and age_row:
        for c in age_cols:
            table_region.add('%s%d' % (c, age_row))
        for ref, cell in tsheet.cells.items():
            col, r = _split_ref(ref)
            if r in (age_row, height_row) and ('年龄' in cell['text'] or '身高' in cell['text']):
                table_region.add(ref)
    helper_cells = []
    if tsheet:
        for ref, cell in tsheet.cells.items():
            if ref in table_region:
                continue
            if cell['f'] is None and cell['num'] is None:
                continue          # 纯文字（标题/说明）不算辅助数据
            col, r = _split_ref(ref)
            if r in tsheet.row_hidden or _col_to_num(col) in tsheet.col_hidden:
                continue          # 真正隐藏的行列放行
            helper_cells.append(ref)
    if tsheet is not None and not helper_cells:
        results.append("+3：交付整洁，统计表所在工作表无需用户维护的可见辅助数据区")
    else:
        results.append("+0：交付整洁，统计表所在工作表无需用户维护的可见辅助数据区")

    # +1：A03 统计表标题（体现“11-20 岁身高统计表”）
    title_ok = False
    if tsheet and age_row:
        for ref, cell in tsheet.cells.items():
            _, r = _split_ref(ref)
            if r is None or r >= age_row:
                continue
            t = (cell['text'] or '').replace(' ', '')
            if '统计表' in t or ('身高' in t and '表' in t):
                if re.search(r'11\s*[-~～－—至到]\s*20', t):
                    title_ok = True
                    break
    if title_ok:
        results.append("+1：统计表设有标题，体现“11-20 岁身高统计表”")
    else:
        results.append("+0：统计表设有标题，体现“11-20 岁身高统计表”")

    # +3：A04 年龄行完整有序填入 11、12…20 共 10 个
    if ages_got == _AGES:
        results.append("+3：年龄行完整有序填入 11、12…20 共 10 个年龄")
    else:
        results.append("+0：年龄行完整有序填入 11、12…20 共 10 个年龄")

    # +3：A05 身高/cm 行：与年龄一一对应的 10 个可直接录入的单元格（非公式）
    h_struct_ok = False
    if tsheet and height_cells and len(height_cells) == 10:
        label_ok = any('身高' in (c['text'] or '')
                       for ref, c in tsheet.cells.items()
                       if _split_ref(ref)[1] == height_row)
        no_formula = all(tsheet.formula(ref) is None for ref in height_cells)
        no_text = all(tsheet.num(ref) is not None or not (tsheet.text(ref) or '').strip()
                      for ref in height_cells)
        h_struct_ok = bool(label_ok and no_formula and no_text)
    if h_struct_ok:
        results.append("+3：设有“身高/cm”行，10 个与年龄一一对应、可直接录入的单元格")
    else:
        results.append("+0：设有“身高/cm”行，10 个与年龄一一对应、可直接录入的单元格")

    # +3：A06 身高单元格留空交付，用户可直接填写、无需先清除示例数据
    blank_ok = bool(height_cells) and all(
        (tsheet.num(ref) is None and not (tsheet.text(ref) or '').strip())
        for ref in height_cells)
    if blank_ok:
        results.append("+3：身高单元格留空交付，用户可直接填写、无需先清除示例数据")
    else:
        results.append("+0：身高单元格留空交付，用户可直接填写、无需先清除示例数据")

    # +1：A07 表格具备边框、内容居中，列宽足以完整显示“身高/cm”等文字
    style_ok = False
    if tsheet and age_cells and height_cells:
        checked = list(age_cells) + list(height_cells)
        label_refs = [ref for ref, c in tsheet.cells.items()
                      if _split_ref(ref)[1] in (age_row, height_row)
                      and ('年龄' in c['text'] or '身高' in c['text'])]
        checked += label_refs
        all_styled = True
        for ref in checked:
            cell = tsheet.cells.get(ref)
            if cell is None:
                all_styled = False
                break
            if not wb.has_border(cell['s']) or wb.style_of(cell['s'])['halign'] != 'center':
                all_styled = False
                break
        label_cols = [_split_ref(r)[0] for r in label_refs] or ['A']
        width_ok = all(tsheet.width_of(c) >= 6.0 for c in label_cols)
        style_ok = bool(all_styled and width_ok)
    if style_ok:
        results.append("+1：统计表加边框、内容居中，列宽足以完整显示“年龄”“身高/cm”")
    else:
        results.append("+0：统计表加边框、内容居中，列宽足以完整显示“年龄”“身高/cm”")

    # +1：A08 未设置与“0-200cm”矛盾的限制或提示（数据有效性上限、提示文案口径）
    conflict = []
    if tsheet and height_cells:
        hset = set(height_cells)
        for dv in tsheet.dv:
            covered = set()
            for area in (dv['sqref'] or '').split():
                covered |= set(_expand_area(area))
            if not (covered & hset):
                continue
            try:
                lo = float(dv['f1']) if dv['f1'] is not None else None
                hi = float(dv['f2']) if dv['f2'] is not None else None
            except ValueError:
                lo = hi = None
            if hi is None or hi < _Y_MAX or (lo is not None and lo > 0):
                conflict.append('数据有效性范围 %s~%s 与 0-200cm 不一致' % (dv['f1'], dv['f2']))
        for ref, cell in tsheet.cells.items():
            t = (cell['text'] or '')
            for m in re.finditer(r'0\s*[-~～－—到至]\s*(\d{2,3})', t):
                v = float(m.group(1))
                if 100 <= v <= 300 and abs(v - _Y_MAX) > 0.01:
                    conflict.append('提示文案写作 0-%s，与 0-200cm 不一致' % m.group(1))
            if '黄色' in t or '黄底' in t:
                if not any(wb.has_yellow_fill((tsheet.cells.get(h) or {}).get('s'))
                           for h in height_cells):
                    conflict.append('提示“黄色单元格”但身高输入区并无黄色填充')
    if tsheet is not None and height_cells and not conflict:
        results.append("+1：未设置与 0-200cm 矛盾的输入限制或提示文案")
    else:
        results.append("+0：未设置与 0-200cm 矛盾的输入限制或提示文案")

    # +3：A09 图表为 XY 散点图，且数据点之间不绘制折线
    #     判定依据只认“系列线条被显式设为无线条”（ser/spPr/a:ln/a:noFill，或 ln@w=0）：
    #     c:scatterStyle 只是子类型提示，Excel/WPS 对散点系列一律按 spPr 里的线条设定
    #     渲染——只写 scatterStyle="marker" 而不关掉系列线条时，打开文件仍会画出折线。
    groups = set((chart or {}).get('groups') or [])
    scatter_ok = False
    if chart is not None and series:
        is_scatter = 'scatterChart' in groups
        no_other = not (groups - {'scatterChart'})
        markers_only = all(s['no_line'] for s in series)
        scatter_ok = bool(is_scatter and no_other and markers_only)
    if scatter_ok:
        results.append("+3：图表为 XY 散点图且仅显示数据点，数据点之间无连线")
    else:
        results.append("+0：图表为 XY 散点图且仅显示数据点，数据点之间无连线")

    # +1：A10 数据点标记清晰可见（未把标记设为“无”）
    marker_ok = bool(series) and all((s['marker'] or 'auto') != 'none' for s in series)
    if marker_ok:
        results.append("+1：数据点使用可见标记（未设为“无”）")
    else:
        results.append("+0：数据点使用可见标记（未设为“无”）")

    # +1：A11 数据点旁显示身高数值标签（系列级设定优先于图表组级）
    label_val_ok = bool(series) and all(s['show_val'] == '1' for s in series)
    if label_val_ok:
        results.append("+1：数据点旁显示身高数值标签")
    else:
        results.append("+0：数据点旁显示身高数值标签")

    # +1：A12 图表标题体现“11～20 岁身高统计图”
    ct = ((chart or {}).get('title') or '').replace(' ', '')
    ctitle_ok = bool(ct) and not (chart or {}).get('title_deleted') and '身高' in ct \
        and ('统计' in ct or '图' in ct)
    if ctitle_ok:
        results.append("+1：图表设有标题，体现“11～20 岁身高统计图”")
    else:
        results.append("+0：图表设有标题，体现“11～20 岁身高统计图”")

    # +3：A13 横轴（年龄轴）最小值 11、最大值 20、主要刻度单位 1
    def _eq(a, b):
        return a is not None and abs(a - b) < 0.01

    x_ok = bool(x_ax) and _eq(x_ax['min'], _X_MIN) and _eq(x_ax['max'], _X_MAX) \
        and _eq(x_ax['major_unit'], _X_UNIT)
    if x_ok:
        results.append("+3：横轴最小值 11、最大值 20、主要刻度单位 1（10 个年龄刻度齐全）")
    else:
        results.append("+0：横轴最小值 11、最大值 20、主要刻度单位 1（10 个年龄刻度齐全）")

    # +3：A14 纵轴（身高轴）最小值 0、最大值 200（min 缺省=自动从 0 起，视为达标）
    y_ok = bool(y_ax) and _eq(y_ax['max'], _Y_MAX) and (y_ax['min'] is None or _eq(y_ax['min'], 0.0))
    if y_ok:
        results.append("+3：纵轴最小值 0、最大值 200")
    else:
        results.append("+0：纵轴最小值 0、最大值 200")

    # +1：A15 纵轴主要刻度单位 50（每 50 一格）
    y_unit_ok = bool(y_ax) and _eq(y_ax['major_unit'], _Y_UNIT)
    if y_unit_ok:
        results.append("+1：纵轴主要刻度单位 50（0/50/100/150/200 每 50 一格）")
    else:
        results.append("+0：纵轴主要刻度单位 50（0/50/100/150/200 每 50 一格）")

    # +1：A16 横轴标题“年龄”
    xt_ok = bool(x_ax) and '年龄' in (x_ax['title'] or '')
    if xt_ok:
        results.append("+1：横轴设有轴标题“年龄”")
    else:
        results.append("+0：横轴设有轴标题“年龄”")

    # +1：A17 纵轴标题“身高/cm”（含单位 cm）
    yt = (y_ax['title'] if y_ax else '') or ''
    yt_ok = '身高' in yt and re.search(r'cm', yt, re.I) is not None
    if yt_ok:
        results.append("+1：纵轴设有轴标题“身高/cm”，体现单位 cm")
    else:
        results.append("+0：纵轴设有轴标题“身高/cm”，体现单位 cm")

    # +1：A18 绘图区有主要网格线（至少纵轴方向，便于按格读数）
    grid_ok = bool(y_ax) and y_ax['gridlines']
    if grid_ok:
        results.append("+1：绘图区设有主要网格线，便于按格读取数据点")
    else:
        results.append("+0：绘图区设有主要网格线，便于按格读取数据点")

    # ---- 表图联动：先把各系列的 X/Y 数据源解析成“可达的表内单元格” ----
    y_reach, x_reach = set(), set()
    y_is_lit = x_is_lit = False
    y_counts = []
    helper_formulas = []
    for s in series:
        if s.get('y_is_lit') or s.get('val_is_lit'):
            y_is_lit = True
        if s.get('x_is_lit') or s.get('cat_is_lit'):
            x_is_lit = True
        yf = s.get('y_f') or s.get('val_f')
        xf = s.get('x_f') or s.get('cat_f')
        y_reach |= _reach(yf, height_cells, tsheet)
        x_reach |= _reach(xf, age_cells, tsheet)
        y_counts.append(s.get('y_count') or s.get('val_count') or 0)
        for f in (yf, xf):
            sname, cells = _parse_formula_ref(f)
            src = _sheet_by_name(sname)
            if src is None:
                continue
            for ref in cells:
                fml = src.formula(ref)
                if fml:
                    helper_formulas.append(fml.upper())

    # +5：A19 Y 值系列联动统计表身高单元格，输入即自动出现数据点
    y_link_ok = bool(height_cells) and not y_is_lit and len(y_reach) >= 1
    if y_link_ok:
        results.append("+5：Y 值系列联动统计表身高单元格，输入数据后图表自动显示数据点")
    else:
        results.append("+0：Y 值系列联动统计表身高单元格，输入数据后图表自动显示数据点")

    # +3：A20 X 值系列引用统计表年龄单元格（非写死的常量数组）
    x_link_ok = bool(age_cells) and not x_is_lit and len(x_reach) == len(age_cells)
    if x_link_ok:
        results.append("+3：X 值系列引用统计表年龄单元格，非写死的常量数组")
    else:
        results.append("+0：X 值系列引用统计表年龄单元格，非写死的常量数组")

    # +3：A21 空值不产生 0 值假数据点（显式空距，或用 NA()/IF 屏蔽空单元格）
    blank_guard_ok = ((chart or {}).get('disp_blanks') == 'gap'
                      or any('NA()' in f for f in helper_formulas))
    if blank_guard_ok:
        results.append("+3：未填写的身高不产生 0 值假数据点（空距显示或 NA() 屏蔽）")
    else:
        results.append("+0：未填写的身高不产生 0 值假数据点（空距显示或 NA() 屏蔽）")

    # +3：A22 10 个年龄全部纳入联动区域，任意年龄单独填写均能正确出点
    #     前置 y_link_ok：避免出现“联动未打通却判全区域纳入”的自相矛盾组合
    #     （某系列 Y 值写死常量时 y_is_lit 置真，而 y_reach 是跨系列并集）。
    full_ok = (y_link_ok and len(height_cells) == 10 and len(y_reach) == 10
               and any(c == 10 for c in y_counts))
    if full_ok:
        results.append("+3：10 个年龄全部纳入联动区域，任意年龄单独填写均能正确出点")
    else:
        results.append("+0：10 个年龄全部纳入联动区域，任意年龄单独填写均能正确出点")

    # +1：A23 工作簿处于自动重算，输入后图表即时刷新
    if wb.calc_mode != 'manual':
        results.append("+1：工作簿处于自动重算，输入数据后图表即时刷新")
    else:
        results.append("+0：工作簿处于自动重算，输入数据后图表即时刷新")

    # ========== 将 results 反解为结构化 dim2_items ==========
    # 逐项按追加顺序对齐 _MAX_DELTAS_ORDER：解析每条 "+N：说明" 的 N 与 rule。
    if len(results) != len(_MAX_DELTAS_ORDER):
        raise RuntimeError(
            f"dim2 项数不一致：results={len(results)} vs order={len(_MAX_DELTAS_ORDER)}"
        )
    dim2_items = []
    total_score = 0
    for raw, max_delta in zip(results, _MAX_DELTAS_ORDER):
        head, _, rule = raw.partition("：")
        try:
            delta = int(head.replace("+", "").strip())
        except ValueError:
            delta = 0
        hit = delta > 0
        dim2_items.append({
            "rule": rule,
            "max_delta": max_delta,
            "delta": max_delta if hit else 0,
            "hit": hit,
        })
        total_score += (max_delta if hit else 0)

    result["dim2_items"] = dim2_items
    result["total_score"] = total_score
    result["max_score"] = MAX_SCORE_TOTAL
    return result


if __name__ == "__main__":
    # 本地调试入口：evaluate 只接收“脚本所在目录的路径”，由脚本自己在该目录里
    # 定位并打开被评估的 .xlsx。命令行第 1 个参数可覆盖目录路径；未指定时使用
    # 脚本自身所在目录。
    if len(sys.argv) >= 2:
        _dir = sys.argv[1]
    else:
        _dir = os.path.dirname(os.path.abspath(__file__))
    _out = evaluate(_dir)
    print(json.dumps(_out, ensure_ascii=False, indent=2))
