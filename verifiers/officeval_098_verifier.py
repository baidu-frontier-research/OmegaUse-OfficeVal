# -*- coding: utf-8 -*-
"""
自动评估脚本：对 "满意度调查-指标分析_重要性满意度象限图.xlsx" 按打分细则进行自动评分。

评分逻辑（与题目要求一致）：
  维度1（可用与可修改性）—— 闸门。任意一条不满足 -> 直接 0 分，不再检查维度2。
  维度2（完成度评分细则）—— 在通过维度1后逐条检查：
      * 加分细则：必须满足该细则中【每一个点】才计该分；
      * 扣分细则：只要满足该细则中【任意一点】即扣分（本细则集中无扣分项，但框架已预留）。
  最终打印命中的点及总分。

实现说明：
  本脚本不依赖 Excel/COM，直接解析 .xlsx/.xlsm（本质是 zip + xml）。
  - 通过解压读取 worksheets / charts / drawings 等 OOXML 部件做结构化判断。
  - 对"图表为原生可编辑对象""散点图""参考线""字体"等，均通过 XML 结构精确判定。
  - 对不易精确判定的意图（如"基本一致""未严重遮挡"），采用数值容差/几何近似的方式灵活满足评估意图。
"""

import os
import sys
import re
import json
import zipfile
import xml.etree.ElementTree as ET

# ----------------------------------------------------------------------------
# 命名空间
# ----------------------------------------------------------------------------
NS = {
    'main': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
    'r':    'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    'c':    'http://schemas.openxmlformats.org/drawingml/2006/chart',
    'a':    'http://schemas.openxmlformats.org/drawingml/2006/main',
    'xdr':  'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing',
    'ct':   'http://schemas.openxmlformats.org/package/2006/content-types',
    'pr':   'http://schemas.openxmlformats.org/package/2006/relationships',
}

# 评估意图涉及的关键常量
VERTICAL_REF_X = 4.11    # 纵轴参考线（X=体验总均值）
HORIZONTAL_REF_Y = 3.63  # 横轴参考线（Y=关注度总均值）
EXPECTED_POINTS = 20     # 散点数量
TOL = 0.02               # 数值容差（用于"基本一致"等近似判断）


class XlsxPackage:
    """以 zip 方式读取 xlsx，提供按部件名读取 / 解析 XML 的能力。"""

    def __init__(self, path):
        self.path = path
        self.ok_open = False
        self.names = []
        self._cache = {}
        try:
            self.zf = zipfile.ZipFile(path, 'r')
            self.names = self.zf.namelist()
            self.ok_open = True
        except Exception as e:                       # noqa
            self.zf = None
            self.open_error = str(e)

    def read(self, name):
        if name in self._cache:
            return self._cache[name]
        try:
            data = self.zf.read(name)
        except KeyError:
            data = None
        self._cache[name] = data
        return data

    def read_xml(self, name):
        data = self.read(name)
        if data is None:
            return None
        try:
            return ET.fromstring(data)
        except Exception:                            # noqa
            return None

    def exists(self, name):
        return name in self.names

    def close(self):
        if self.zf:
            self.zf.close()


# ----------------------------------------------------------------------------
# 通用工具
# ----------------------------------------------------------------------------
def col_letter_to_idx(col):
    """A->1, B->2 ..."""
    col = col.upper()
    n = 0
    for ch in col:
        n = n * 26 + (ord(ch) - ord('A') + 1)
    return n


def split_cell_ref(ref):
    """'C21' -> ('C', 21)"""
    m = re.match(r'^([A-Za-z]+)(\d+)$', ref)
    if not m:
        return None, None
    return m.group(1).upper(), int(m.group(2))


def ref_range_matches(formula, sheet_name, col, row_start, row_end):
    """校验图表系列公式（如 c:xVal/c:numRef/c:f 的内容）是否引用了
    指定工作表某一列的连续区间，例如 formula='Sheet1!$C$2:$C$21'，
    col='C', row_start=2, row_end=21 -> True。
    兼容：工作表名是否加引号（Sheet1! / 'Sheet1'!）、列/行是否有 $ 绝对引用符。
    """
    if not formula or not isinstance(formula, str):
        return False
    f = formula.strip()
    # 拆出工作表名与区间部分：'Sheet1'!$C$2:$C$21  或  Sheet1!$C$2:$C$21
    m = re.match(r"^(?:'([^']+)'|([A-Za-z0-9_]+))!(.+)$", f)
    if not m:
        return False
    sheet = m.group(1) if m.group(1) is not None else m.group(2)
    rng = m.group(3)
    if sheet != sheet_name:
        return False
    m2 = re.match(r'^\$?([A-Za-z]+)\$?(\d+):\$?([A-Za-z]+)\$?(\d+)$', rng)
    if not m2:
        return False
    c1, r1, c2, r2 = m2.group(1).upper(), int(m2.group(2)), m2.group(3).upper(), int(m2.group(4))
    return (c1 == col.upper() and c2 == col.upper()
            and r1 == row_start and r2 == row_end)


def load_shared_strings(pkg):
    root = pkg.read_xml('xl/sharedStrings.xml')
    out = []
    if root is None:
        return out
    for si in root.findall('main:si', NS):
        # 拼接 si 下所有 t 文本
        texts = [t.text or '' for t in si.iter('{%s}t' % NS['main'])]
        out.append(''.join(texts))
    return out


def get_sheet_map(pkg):
    """返回 {sheet_name: worksheet_part_path}"""
    wb = pkg.read_xml('xl/workbook.xml')
    rels = pkg.read_xml('xl/_rels/workbook.xml.rels')
    if wb is None or rels is None:
        return {}
    rid_to_target = {}
    for rel in rels.findall('pr:Relationship', NS):
        rid_to_target[rel.get('Id')] = rel.get('Target')
    out = {}
    for sh in wb.findall('main:sheets/main:sheet', NS):
        name = sh.get('name')
        rid = sh.get('{%s}id' % NS['r'])
        target = rid_to_target.get(rid)
        if target:
            if not target.startswith('xl/'):
                target = 'xl/' + target.lstrip('/')
            out[name] = target
    return out


def read_sheet_cells(pkg, sheet_path, shared):
    """读取一个 worksheet 的单元格 {cellref: value(str/float)}，并返回 dimension。"""
    root = pkg.read_xml(sheet_path)
    cells = {}
    dimension = None
    if root is None:
        return cells, dimension
    dim_el = root.find('main:dimension', NS)
    if dim_el is not None:
        dimension = dim_el.get('ref')
    for c in root.iter('{%s}c' % NS['main']):
        ref = c.get('r')
        t = c.get('t')
        v_el = c.find('main:v', NS)
        is_el = c.find('main:is', NS)
        val = None
        if t == 's' and v_el is not None:            # shared string
            try:
                val = shared[int(v_el.text)]
            except Exception:                        # noqa
                val = None
        elif t == 'inlineStr' and is_el is not None:
            val = ''.join(x.text or '' for x in is_el.iter('{%s}t' % NS['main']))
        elif v_el is not None:
            txt = v_el.text
            try:
                val = float(txt)
            except (TypeError, ValueError):
                val = txt
        cells[ref] = val
    return cells, dimension


# ----------------------------------------------------------------------------
# 评分结果容器
# ----------------------------------------------------------------------------
class Report:
    def __init__(self):
        self.dim1_checks = []   # [(desc, passed, note)]
        self.dim2_items = []    # [(score, desc, hit, note)]
        self.dim1_failed = False
        self.total = 0.0

    def add_dim1(self, desc, passed, note=''):
        self.dim1_checks.append((desc, passed, note))
        if not passed:
            self.dim1_failed = True

    def add_dim2(self, score, desc, hit, note=''):
        self.dim2_items.append((score, desc, hit, note))
        if hit:
            self.total += score

    def render(self):
        lines = []

        # 维度1未全部通过 -> 直接 0 分
        if self.dim1_failed:
            lines.append('维度一：不通过')
            for desc, passed, note in self.dim1_checks:
                if not passed:
                    lines.append('  ✘ %s' % desc)
                    if note:
                        lines.append('     · %s' % note)
            lines.append('')
            lines.append('总分：0 分')
            return '\n'.join(lines)

        # 维度1全部通过
        lines.append('维度一：通过')
        lines.append('')
        lines.append('维度二：评分结果')
        # 只显示命中项，命中项只显示对应的分数和评分细则内容，分数按"+N："格式
        for score, desc, hit, _note in self.dim2_items:
            if hit:
                lines.append('+%g：%s' % (score, desc))
        lines.append('')
        # 总分写在评分结果最后面
        lines.append('总分：%g 分' % self.total)
        return '\n'.join(lines)


# ----------------------------------------------------------------------------
# 图表解析：把 chart*.xml 解析为结构化信息
# ----------------------------------------------------------------------------
def find_chart_parts(pkg):
    """返回所有 xl/charts/chartN.xml 部件名。"""
    return [n for n in pkg.names
            if re.match(r'xl/charts/chart\d+\.xml$', n)]


def text_of(el):
    """取一个元素下所有 a:t 文本拼接。"""
    if el is None:
        return ''
    return ''.join(t.text or '' for t in el.iter('{%s}t' % NS['a']))


def parse_chart(pkg, chart_path):
    """解析散点图，返回结构化字典。"""
    root = pkg.read_xml(chart_path)
    info = {
        'path': chart_path,
        'is_scatter': False,
        'has_bar': False, 'has_pie': False, 'has_line': False, 'has_radar': False,
        'title': '',
        'series': [],        # 每个 ser: {'name','xf','yf','xvals','yvals','dlbls':[...]}
        'val_axes': [],      # 数值轴标题文本列表
        'cat_axes': [],
        'axis_titles': [],   # 所有轴标题文本
        'fonts': set(),      # 出现过的字体 typeface
        # 坐标轴范围（用于判断参考线是否贯穿全图）：
        #   axPos='b'(底部)→X轴；axPos='l'(左侧)→Y轴。无显式min/max则为None（自动）。
        'x_axis_min': None, 'x_axis_max': None,
        'y_axis_min': None, 'y_axis_max': None,
        'raw': root,
    }
    if root is None:
        return info

    # 图表类型
    plot = root.find('.//c:plotArea', NS)
    if plot is not None:
        if plot.find('c:scatterChart', NS) is not None:
            info['is_scatter'] = True
        if plot.find('c:barChart', NS) is not None:
            info['has_bar'] = True
        if plot.find('c:pieChart', NS) is not None or plot.find('c:pie3DChart', NS) is not None:
            info['has_pie'] = True
        if plot.find('c:lineChart', NS) is not None:
            info['has_line'] = True
        if plot.find('c:radarChart', NS) is not None:
            info['has_radar'] = True

    # 标题
    title_el = root.find('.//c:chart/c:title', NS)
    info['title'] = text_of(title_el)

    # 字体（latin typeface）
    for latin in root.iter('{%s}latin' % NS['a']):
        tf = latin.get('typeface')
        if tf:
            info['fonts'].add(tf)

    # 解析每个散点 series
    scatter = plot.find('c:scatterChart', NS) if plot is not None else None
    if scatter is not None:
        for ser in scatter.findall('c:ser', NS):
            s = {'name': '', 'xf': '', 'yf': '', 'catf': '', 'xvals': [], 'yvals': [], 'dlbls': [],
                 # 线条样式（用于判断参考线是否为"可视化实线"）：
                 #   line_visible：是否画了可见线条（有 ln 且非 noFill）；
                 #   line_solid：线型是否为实线（无 prstDash 或 prstDash=solid）；
                 #   line_w：线宽 EMU（0 表示未显式设置）。
                 'line_visible': False, 'line_solid': False, 'line_w': 0}
            # 线条样式：spPr/a:ln
            sppr = ser.find('c:spPr', NS)
            if sppr is not None:
                ln = sppr.find('a:ln', NS)
                if ln is not None:
                    no_fill = ln.find('a:noFill', NS) is not None
                    has_solid_fill = ln.find('a:solidFill', NS) is not None
                    s['line_visible'] = (not no_fill) and has_solid_fill
                    dash = ln.find('a:prstDash', NS)
                    dash_val = dash.get('val') if dash is not None else None
                    # 无 prstDash 或 prstDash=solid 均视为实线
                    s['line_solid'] = (dash_val is None or dash_val == 'solid')
                    try:
                        s['line_w'] = int(ln.get('w')) if ln.get('w') else 0
                    except Exception:                # noqa
                        s['line_w'] = 0
            # series name
            tx = ser.find('c:tx', NS)
            if tx is not None:
                v = tx.find('c:v', NS)
                if v is not None and v.text:
                    s['name'] = v.text
                else:
                    s['name'] = text_of(tx)
            # x / y refs + caches
            xval = ser.find('c:xVal/c:numRef', NS)
            if xval is not None:
                f = xval.find('c:f', NS)
                s['xf'] = f.text if f is not None else ''
                for pt in xval.findall('c:numCache/c:pt', NS):
                    vv = pt.find('c:v', NS)
                    try:
                        s['xvals'].append(float(vv.text))
                    except Exception:                # noqa
                        pass
            yval = ser.find('c:yVal/c:numRef', NS)
            if yval is not None:
                f = yval.find('c:f', NS)
                s['yf'] = f.text if f is not None else ''
                for pt in yval.findall('c:numCache/c:pt', NS):
                    vv = pt.find('c:v', NS)
                    try:
                        s['yvals'].append(float(vv.text))
                    except Exception:                # noqa
                        pass
            # 类别/标签引用（c:cat，可能是 numRef 或 strRef），用于校验标签是否来自 A 列
            cat_ref = ser.find('c:cat/c:strRef', NS)
            if cat_ref is None:
                cat_ref = ser.find('c:cat/c:numRef', NS)
            if cat_ref is not None:
                f = cat_ref.find('c:f', NS)
                s['catf'] = f.text if f is not None else ''
            # data labels（取每个 dLbl 的富文本内容）
            for dlbl in ser.findall('c:dLbls/c:dLbl', NS):
                txt = text_of(dlbl.find('c:tx', NS))
                idx_el = dlbl.find('c:idx', NS)
                idx = int(idx_el.get('val')) if idx_el is not None else -1
                if txt:
                    s['dlbls'].append((idx, txt))
            info['series'].append(s)

    # 轴标题
    for vax in root.findall('.//c:valAx', NS):
        t = text_of(vax.find('c:title', NS))
        info['val_axes'].append(t)
        if t:
            info['axis_titles'].append(t)
    for cax in root.findall('.//c:catAx', NS):
        t = text_of(cax.find('c:title', NS))
        info['cat_axes'].append(t)
        if t:
            info['axis_titles'].append(t)

    # 坐标轴范围（用于判断参考线是否贯穿全图）：
    #   axPos='b'(底部)→X轴；axPos='l'(左侧)→Y轴。读 scaling 下的 min/max。
    # 同时读取 crosses/crossesAt（该轴与其交叉轴相交的位置），
    #   用于判断横纵轴是否位于图表中心交叉（而非仅凭参考线代替）：
    #   crosses='autoZero' 通常表示在 0 处交叉；crossesAt=数值 表示在该数值处交叉。
    for vax in root.findall('.//c:valAx', NS):
        pos_el = vax.find('c:axPos', NS)
        pos = pos_el.get('val') if pos_el is not None else ''
        scaling = vax.find('c:scaling', NS)
        amin = amax = None
        if scaling is not None:
            mn = scaling.find('c:min', NS)
            mx = scaling.find('c:max', NS)
            try:
                amin = float(mn.get('val')) if mn is not None else None
            except Exception:                # noqa
                amin = None
            try:
                amax = float(mx.get('val')) if mx is not None else None
            except Exception:                # noqa
                amax = None
        crosses_el = vax.find('c:crosses', NS)
        crosses_val = crosses_el.get('val') if crosses_el is not None else None
        crosses_at_el = vax.find('c:crossesAt', NS)
        try:
            crosses_at = float(crosses_at_el.get('val')) if crosses_at_el is not None else None
        except Exception:                # noqa
            crosses_at = None
        if pos == 'b':            # 底部轴 = X 轴
            info['x_axis_min'], info['x_axis_max'] = amin, amax
            info['x_crosses'], info['x_crosses_at'] = crosses_val, crosses_at
        elif pos == 'l':          # 左侧轴 = Y 轴
            info['y_axis_min'], info['y_axis_max'] = amin, amax
            info['y_crosses'], info['y_crosses_at'] = crosses_val, crosses_at

    return info


def classify_series(chart):
    """从所有散点 series 中区分：数据点系列 / 参考线系列 / 其它。
    数据点系列：点数接近 20 的那个（取点数最多且 >=20-容差的）。"""
    data_ser = None
    ref_series = []
    for s in chart['series']:
        n = max(len(s['xvals']), len(s['yvals']))
        if n >= EXPECTED_POINTS - 2 and (data_ser is None or n > max(len(data_ser['xvals']), len(data_ser['yvals']))):
            data_ser = s
    for s in chart['series']:
        if s is data_ser:
            continue
        ref_series.append(s)
    return data_ser, ref_series


# ----------------------------------------------------------------------------
# 维度1：可用与可修改性（闸门）
# ----------------------------------------------------------------------------
def check_dimension1(pkg, path, sheets, sheet1_cells, sheet1_dim, charts, report):
    # 1.1 格式为 xlsx 或 .xlsm，且文件可正常打开
    ext_ok = path.lower().endswith(('.xlsx', '.xlsm'))
    open_ok = pkg.ok_open
    report.add_dim1(
        '交付文件为xlsx或.xlsm格式，文件可正常打开',
        ext_ok and open_ok,
        '扩展名=%s；可正常打开=%s' % (os.path.splitext(path)[1], open_ok))


# ----------------------------------------------------------------------------
# 维度2：完成度评分细则
# ----------------------------------------------------------------------------
def check_dimension2(pkg, sheets, sheet1_cells, charts, report):
    # 选用主图表（含散点数据系列的那个）
    main_chart = None
    main_data = None
    for ch in charts:
        ds, _ = classify_series(ch)
        if ds is not None:
            main_chart = ch
            main_data = ds
            break
    if main_chart is None and charts:
        main_chart = charts[0]

    # Sheet1 原始数据（用于"基本一致"比对）：B=关注度均值, C=体验均值
    src_B = [sheet1_cells.get('B%d' % r) for r in range(2, 22)]   # 关注度
    src_C = [sheet1_cells.get('C%d' % r) for r in range(2, 22)]   # 体验

    # ---- +1 图表数据源：图表使用A2:C21的20项指标数据作为绘图数据源，
    #            A列"评估维度"为指标名称，B列为关注度均值，C列为体验均值 ----
    # 严格按细则逐点拆解，每一个点都必须踩到（全部满足才加分），
    # 不附加细则未要求的额外约束：
    #   点① 图表使用"A2:C21 的20项指标数据"作为绘图数据源——数据系列点数=20，
    #        且系列公式 xf/yf 实际引用 Sheet1!C2:C21 / Sheet1!B2:B21（而非仅缓存值凑巧相同）；
    #   点② A列"评估维度"数据为指标名称——A1 表头为"评估维度"，A2:A21 为指标名称文本，
    #        且图表类别/标签引用（若存在）实际指向 Sheet1!A2:A21；
    #   点③ B列数据为关注度均值——B1 表头为"关注度均值"，B2:B21 为数值；
    #   点④ C列数据为体验均值——C1 表头为"体验均值"，C2:C21 为数值。
    src_ok = False
    note = ''
    if main_data:
        xs = main_data['xvals']
        ys = main_data['yvals']

        # 点① 图表使用 A2:C21 的20项指标数据作为绘图数据源
        #     -> 数据系列恰好 20 个点，且 xVal/yVal 公式确实引用 C2:C21 / B2:B21
        p1_count = (len(xs) == EXPECTED_POINTS and len(ys) == EXPECTED_POINTS)
        p1_xf_ref_ok = ref_range_matches(main_data.get('xf', ''), 'Sheet1', 'C', 2, 21)
        p1_yf_ref_ok = ref_range_matches(main_data.get('yf', ''), 'Sheet1', 'B', 2, 21)
        p1_count = p1_count and p1_xf_ref_ok and p1_yf_ref_ok

        # 点② A列"评估维度"数据为指标名称；若图表有类别/标签引用，须指向 A2:A21
        a_header = sheet1_cells.get('A1')
        a_header_ok = isinstance(a_header, str) and ('评估维度' in a_header)
        a_names = [sheet1_cells.get('A%d' % r) for r in range(2, 22)]   # A2:A21
        a_names_ok = (len(a_names) == EXPECTED_POINTS
                      and all(isinstance(v, str) and v.strip() != '' for v in a_names))
        catf = main_data.get('catf', '')
        cat_ref_ok = (not catf) or ref_range_matches(catf, 'Sheet1', 'A', 2, 21)
        p2_a_name = a_header_ok and a_names_ok and cat_ref_ok

        # 点③ B列数据为关注度均值
        b_header = sheet1_cells.get('B1')
        b_header_ok = isinstance(b_header, str) and ('关注度均值' in b_header)
        b_vals = [sheet1_cells.get('B%d' % r) for r in range(2, 22)]    # B2:B21
        b_vals_ok = (len(b_vals) == EXPECTED_POINTS
                     and all(isinstance(v, (int, float)) for v in b_vals))
        p3_b_attention = b_header_ok and b_vals_ok

        # 点④ C列数据为体验均值
        c_header = sheet1_cells.get('C1')
        c_header_ok = isinstance(c_header, str) and ('体验均值' in c_header)
        c_vals = [sheet1_cells.get('C%d' % r) for r in range(2, 22)]    # C2:C21
        c_vals_ok = (len(c_vals) == EXPECTED_POINTS
                     and all(isinstance(v, (int, float)) for v in c_vals))
        p4_c_experience = c_header_ok and c_vals_ok

        src_ok = p1_count and p2_a_name and p3_b_attention and p4_c_experience
        note = ('①图表使用A2:C21的20项数据作绘图源(点数=%d,xf引用C2:C21=%s,yf引用B2:B21=%s)=%s；'
                 '②A列"评估维度"为指标名称(类别引用A2:A21=%s)=%s；③B列为关注度均值=%s；④C列为体验均值=%s'
                 % (len(xs), p1_xf_ref_ok, p1_yf_ref_ok, p1_count,
                    cat_ref_ok, p2_a_name, p3_b_attention, p4_c_experience))
    else:
        note = '未找到散点数据系列'
    report.add_dim2(1, '图表数据源：使用A2:C21的20项指标数据作为绘图数据源，'
                       'A列"评估维度"为指标名称，B列为关注度均值，C列为体验均值', src_ok, note)

    # ---- +5 图表类型：新增图表为重要性-满意度象限图，采用XY散点图或等效的
    #            二维坐标散点图绘制，横轴和纵轴位于图表中心交叉作为四个象限的分界线，
    #            不是饼图、柱形图、折线图或雷达图 ----
    # 严格按细则逐点拆解，每一个点都必须踩到（全部满足才加分），
    # 不附加细则未要求的额外约束（如标题文字——那是"图表标题"另一条细则的要求）：
    #   点① 新增图表为"重要性-满意度象限图"——该图为重要性-满意度象限图，
    #        即以 XY 散点图/二维坐标方式呈现象限分析（由散点图类型承载该性质）；
    #   点② 采用XY散点图或等效的二维坐标散点图绘制——图表类型为 scatterChart；
    #   点③ 横轴和纵轴位于图表中心交叉作为四个象限的分界线——
    #        图中同时存在一条纵向（X恒定）与一条横向（Y恒定）的可视化实线分界线
    #        （可见且为实线）并相交成4象限；与"图表参考线"细则口径保持一致；
    #   点④ 不是饼图；
    #   点⑤ 不是柱形图；
    #   点⑥ 不是折线图；
    #   点⑦ 不是雷达图。
    type_ok = False
    note = ''
    if main_chart:
        # 点① 新增图表为"重要性-满意度象限图"（二维坐标散点形式的象限图）
        p1_quadrant = main_chart['is_scatter']
        # 点② 采用XY散点图（或等效二维坐标散点图）绘制
        p2_scatter = main_chart['is_scatter']
        # 点③ 横轴和纵轴位于图表中心交叉作为四个象限的分界线——
        #      核心判定：坐标轴本身的交叉位置（valAx 的 crosses/crossesAt），
        #      而不是用参考线替代坐标轴中心交叉的判断：
        #        · X轴（axPos='b'）的 crossesAt 应落在数据 x 范围中部
        #          （或 crosses='autoZero' 且 0 落在范围中部）；
        #        · Y轴（axPos='l'）同理，crossesAt 落在数据 y 范围中部。
        #      若坐标轴未显式声明交叉位置（Excel 默认 crosses=autoZero，即0），
        #      则退化为按"存在贯穿全图的纵/横可视化实线参考线"作为等效判据
        #      （与"图表参考线"细则口径一致，仅在缺失显式交叉声明时兜底使用）。
        data_ser_for_type, _type_ref_series = classify_series(main_chart)
        dxs_t = data_ser_for_type['xvals'] if data_ser_for_type else []
        dys_t = data_ser_for_type['yvals'] if data_ser_for_type else []

        def _crosses_at_center(crosses, crosses_at, lo, hi):
            """判断轴的交叉位置是否落在数据范围的中部（非贴边）。"""
            if not lo < hi:
                return False
            if crosses_at is not None:
                pos = crosses_at
            elif crosses == 'autoZero':
                pos = 0.0
            else:
                return None   # 未显式声明交叉位置，交由调用方走兜底逻辑
            margin = (hi - lo) * 0.1   # 交叉点须离两端至少10%，排除贴边
            return (lo + margin) <= pos <= (hi - margin)

        x_cross_center = _crosses_at_center(
            main_chart.get('x_crosses'), main_chart.get('x_crosses_at'),
            min(dxs_t) if dxs_t else 0.0, max(dxs_t) if dxs_t else 0.0)
        y_cross_center = _crosses_at_center(
            main_chart.get('y_crosses'), main_chart.get('y_crosses_at'),
            min(dys_t) if dys_t else 0.0, max(dys_t) if dys_t else 0.0)

        if x_cross_center is not None and y_cross_center is not None:
            # 坐标轴显式声明了交叉位置：以此为准
            p3_axes_cross = bool(x_cross_center and y_cross_center)
        else:
            # 未显式声明交叉位置，兜底：纵/横可视化实线参考线贯穿全图并相交
            _has_vertical = any(len(s['xvals']) >= 2 and _all_close(s['xvals'], s['xvals'][0])
                                and s['line_visible'] and s['line_solid']
                                for s in _type_ref_series)
            _has_horizontal = any(len(s['yvals']) >= 2 and _all_close(s['yvals'], s['yvals'][0])
                                  and s['line_visible'] and s['line_solid']
                                  for s in _type_ref_series)
            p3_axes_cross = _has_vertical and _has_horizontal
        # 点④ 不是饼图
        p4_not_pie = not main_chart['has_pie']
        # 点⑤ 不是柱形图
        p5_not_bar = not main_chart['has_bar']
        # 点⑥ 不是折线图
        p6_not_line = not main_chart['has_line']
        # 点⑦ 不是雷达图
        p7_not_radar = not main_chart['has_radar']

        type_ok = (p1_quadrant and p2_scatter and p3_axes_cross and p4_not_pie
                   and p5_not_bar and p6_not_line and p7_not_radar)
        note = ('①为重要性-满意度象限图(二维坐标散点形式)=%s；②采用XY散点图绘制=%s；'
                '③横纵轴在图表中部交叉作为四象限分界线=%s；'
                '④非饼图=%s；⑤非柱形图=%s；⑥非折线图=%s；⑦非雷达图=%s'
                % (p1_quadrant, p2_scatter, p3_axes_cross,
                   p4_not_pie, p5_not_bar, p6_not_line, p7_not_radar))
    report.add_dim2(5, '图表类型：新增图表为重要性-满意度象限图，采用XY散点图或等效的'
                       '二维坐标散点图绘制，横轴和纵轴位于图表中心交叉作为四个象限的分界线，'
                       '不是饼图、柱形图、折线图或雷达图', type_ok, note)

    # ---- +5 图表参考线：图中包含1条纵向参考线和1条横向参考线；
    #            纵轴参考线对应体验总均值4.11，横轴参考线对应关注度总均值3.63，
    #            两条参考线贯穿全图并将图表划分为4个象限 ----
    # 严格按细则逐点拆解，每一个点都必须踩到（全部满足才加分），
    # 不附加细则未要求的额外约束：
    #   点① 图中包含1条纵向参考线——存在恰好1条纵向（X恒定）参考线系列；
    #   点② 图中包含1条横向参考线——存在恰好1条横向（Y恒定）参考线系列；
    #   点③ 纵轴参考线对应体验总均值4.11——纵向线 X≈4.11；
    #   点④ 横轴参考线对应关注度总均值3.63——横向线 Y≈3.63；
    #   点⑤ 两条参考线均贯穿全图——纵向线纵向跨度、横向线横向跨度
    #        既贴合坐标轴范围、又覆盖全部数据点范围；
    #   点⑥ 两条线将图表划分为4个象限——一纵一横同时存在并相交，构成4象限。
    # 注：rubric 未要求参考线必须为"可视化实线"（不要求 solidFill/非虚线），
    #     故不再校验 line_visible/line_solid——默认可见线条或虚线样式若满足
    #     方向、数量、均值坐标、跨度与四象限划分，同样应视为合格。
    ref_ok = False
    note = ''
    if main_chart:
        data_ser, ref_series = classify_series(main_chart)
        vertical_lines = []     # 纵向线：x 恒定（≥2点且x全相等）
        horizontal_lines = []   # 横向线：y 恒定（≥2点且y全相等）
        for s in ref_series:
            xs, ys = s['xvals'], s['yvals']
            if len(xs) >= 2 and _all_close(xs, xs[0]):
                vertical_lines.append(s)
            elif len(ys) >= 2 and _all_close(ys, ys[0]):
                horizontal_lines.append(s)

        # 数据点范围（用于判断参考线是否覆盖全部数据点）
        dxs = data_ser['xvals'] if data_ser else []
        dys = data_ser['yvals'] if data_ser else []
        x_lo, x_hi = (min(dxs), max(dxs)) if dxs else (0.0, 0.0)
        y_lo, y_hi = (min(dys), max(dys)) if dys else (0.0, 0.0)

        # 点① 包含1条纵向参考线
        p1_one_vertical = (len(vertical_lines) == 1)
        # 点② 包含1条横向参考线
        p2_one_horizontal = (len(horizontal_lines) == 1)
        # 点③ 纵向参考线对应体验总均值4.11
        p3_vertical_411 = (p1_one_vertical
                           and _all_close(vertical_lines[0]['xvals'], VERTICAL_REF_X))
        # 点④ 横向参考线对应关注度总均值3.63
        p4_horizontal_363 = (p2_one_horizontal
                             and _all_close(horizontal_lines[0]['yvals'], HORIZONTAL_REF_Y))
        # 点⑤ 两条参考线均贯穿全图（端点贴合轴范围 + 覆盖全部数据点）
        v_full = (p1_one_vertical and _line_spans_full(
            vertical_lines[0], main_chart['y_axis_min'], main_chart['y_axis_max'],
            y_lo, y_hi, along='y'))
        h_full = (p2_one_horizontal and _line_spans_full(
            horizontal_lines[0], main_chart['x_axis_min'], main_chart['x_axis_max'],
            x_lo, x_hi, along='x'))
        p5_full_span = v_full and h_full
        # 点⑥ 两条线将图表划分为4个象限（一纵一横同时存在即相交成4象限）
        p6_four_quadrants = p1_one_vertical and p2_one_horizontal

        ref_ok = (p1_one_vertical and p2_one_horizontal and p3_vertical_411
                  and p4_horizontal_363 and p5_full_span and p6_four_quadrants)
        note = ('①含1条纵向参考线(实测%d条)=%s；②含1条横向参考线(实测%d条)=%s；'
                '③纵向线对应体验总均值4.11=%s；④横向线对应关注度总均值3.63=%s；'
                '⑤两线均贯穿全图(纵线=%s,横线=%s)=%s；⑥划分为4个象限=%s'
                % (len(vertical_lines), p1_one_vertical,
                   len(horizontal_lines), p2_one_horizontal,
                   p3_vertical_411, p4_horizontal_363,
                   v_full, h_full, p5_full_span, p6_four_quadrants))
    report.add_dim2(5, '图表参考线：含1条纵向参考线(对应体验总均值4.11)与1条横向参考线'
                       '(对应关注度总均值3.63)，两条线贯穿全图并将图表'
                       '划分为4个象限', ref_ok, note)

    # ---- +1 图表坐标轴：横轴表示"体验"或"体验均值"，纵轴表示"关注度"或
    #            "关注度均值"，横纵轴方向正确，未出现坐标轴含义互换 ----
    # 严格按细则逐点拆解，每一个点都必须踩到（全部满足才加分），
    # 不附加细则未要求的额外约束：
    #   点① 横轴表示"体验"或"体验均值"——横轴(axPos=b)标题为"体验"或"体验均值"；
    #   点② 纵轴表示"关注度"或"关注度均值"——纵轴(axPos=l)标题为"关注度"或"关注度均值"；
    #   点③ 横纵轴方向正确，未出现坐标轴含义互换——横轴不被"关注度"占用、
    #        纵轴不被"体验"占用，即两轴含义未互换。
    axis_ok = False
    note = ''
    if main_chart:
        x_title, y_title = _axis_titles_by_pos(pkg, main_chart)

        # 点① 横轴表示"体验"或"体验均值"
        p1_x_experience = ('体验' in x_title)
        # 点② 纵轴表示"关注度"或"关注度均值"
        p2_y_attention = ('关注度' in y_title)
        # 点③ 横纵轴方向正确，未出现含义互换
        #     互换的表现：横轴写成了"关注度"、或纵轴写成了"体验"
        p3_not_swapped = ('关注度' not in x_title) and ('体验' not in y_title)

        axis_ok = p1_x_experience and p2_y_attention and p3_not_swapped
        note = ('①横轴表示体验/体验均值=%s(横轴标题="%s")；'
                '②纵轴表示关注度/关注度均值=%s(纵轴标题="%s")；'
                '③横纵轴方向正确未互换=%s'
                % (p1_x_experience, x_title, p2_y_attention, y_title, p3_not_swapped))
    report.add_dim2(1, '图表坐标轴：横轴表示"体验"或"体验均值"，纵轴表示"关注度"或'
                       '"关注度均值"，横纵轴方向正确，未出现坐标轴含义互换', axis_ok, note)

    # ---- +1 图表散点数量：图中共显示20个散点，与A2:A21的20个评估维度的
    #            编号一一对应 ----
    # 严格按细则逐点拆解，每一个点都必须踩到（全部满足才加分），
    # 不附加细则未要求的额外约束：
    #   点① 图中共显示20个散点——数据系列散点数=20；
    #   点② 与A2:A21的20个评估维度的编号一一对应——A2:A21 恰有20个非空维度，
    #        且散点数与维度数一致（一一对应）。
    count_ok = False
    note = ''
    if main_data:
        n = max(len(main_data['xvals']), len(main_data['yvals']))
        # 点① 图中共显示20个散点
        p1_count20 = (n == EXPECTED_POINTS)

        # 点② 与A2:A21的20个评估维度一一对应
        a_dims = [sheet1_cells.get('A%d' % r) for r in range(2, 22)]   # A2:A21
        dims_n = sum(1 for v in a_dims if isinstance(v, str) and v.strip() != '')
        p2_one_to_one = (dims_n == EXPECTED_POINTS and n == dims_n)

        count_ok = p1_count20 and p2_one_to_one
        note = ('①图中共显示20个散点(散点数=%d)=%s；'
                '②与A2:A21的20个维度一一对应(维度数=%d)=%s'
                % (n, p1_count20, dims_n, p2_one_to_one))
    report.add_dim2(1, '图表散点数量：图中共显示20个散点，与A2:A21的20个评估维度的'
                       '编号一一对应', count_ok, note)

    # ---- +1 图表数据位置：各散点的横纵坐标分别对应C2:C21与B2:B21的数值，
    #            点位与原始数据基本一致 ----
    # 严格按细则逐点拆解，每一个点都必须踩到（全部满足才加分），
    # 不附加细则未要求的额外约束：
    #   点① 各散点的横坐标分别对应 C2:C21 的数值——散点 x 值逐一对应 C2:C21；
    #   点② 各散点的纵坐标分别对应 B2:B21 的数值——散点 y 值逐一对应 B2:B21；
    #   点③ 点位与原始数据基本一致——允许极小误差（容差内即视为基本一致）。
    pos_ok = False
    note = ''
    if main_data:
        x_clean = [v for v in src_C if isinstance(v, (int, float))]   # C2:C21
        y_clean = [v for v in src_B if isinstance(v, (int, float))]   # B2:B21

        # 点① 横坐标对应 C2:C21（长度一致且逐点在容差内）
        x_len_ok = (len(main_data['xvals']) == len(x_clean))
        x_mismatch = 0
        if x_len_ok:
            for i in range(len(x_clean)):
                if abs(main_data['xvals'][i] - x_clean[i]) > TOL:
                    x_mismatch += 1
        p1_x_corr = x_len_ok and x_mismatch == 0

        # 点② 纵坐标对应 B2:B21（长度一致且逐点在容差内）
        y_len_ok = (len(main_data['yvals']) == len(y_clean))
        y_mismatch = 0
        if y_len_ok:
            for i in range(len(y_clean)):
                if abs(main_data['yvals'][i] - y_clean[i]) > TOL:
                    y_mismatch += 1
        p2_y_corr = y_len_ok and y_mismatch == 0

        # 点③ 点位与原始数据基本一致（横纵均在容差内即基本一致）
        p3_basically_consistent = p1_x_corr and p2_y_corr

        pos_ok = p1_x_corr and p2_y_corr and p3_basically_consistent
        note = ('①横坐标对应C2:C21(偏差点=%d)=%s；②纵坐标对应B2:B21(偏差点=%d)=%s；'
                '③点位与原始数据基本一致=%s'
                % (x_mismatch, p1_x_corr, y_mismatch, p2_y_corr, p3_basically_consistent))
    report.add_dim2(1, '图表数据位置：各散点的横纵坐标分别对应C2:C21与B2:B21的数值，'
                       '点位与原始数据基本一致', pos_ok, note)

    # ---- +1 图表标签：20个散点均带有可识别的数据标签，标签内容为Q6-Q25题号，
    #            能够区分各评估维度 ----
    # 严格按细则逐点拆解，每一个点都必须踩到（全部满足才加分），
    # 不附加细则未要求的额外约束：
    #   点① 20个散点均带有可识别的数据标签——带标签的散点数=20 且标签文本非空（可识别）；
    #   点② 标签内容为Q6-Q25题号——标签题号集合恰为 {Q6..Q25}；
    #   点③ 能够区分各评估维度——20个标签题号互不相同（一一区分）。
    label_ok = False
    note = ''
    if main_data:
        labels = [t for (_i, t) in main_data['dlbls']]

        # 点① 20个散点均带有可识别的数据标签
        recognizable = [t for t in labels if t and t.strip() != '']
        p1_all_labeled = (len(recognizable) == EXPECTED_POINTS)

        # 解析每个标签中的题号
        expected_q = set('Q%d' % q for q in range(6, 26))   # Q6..Q25
        q_list = []
        for t in labels:
            m = re.search(r'Q\d+', t)
            if m:
                q_list.append(m.group(0))
        found_q = set(q_list)

        # 点② 标签内容为 Q6-Q25 题号（题号集合恰为 Q6..Q25）
        p2_q6_q25 = (found_q == expected_q)

        # 点③ 能够区分各评估维度（题号互不重复，共20个唯一题号）
        p3_distinct = (len(q_list) == EXPECTED_POINTS and len(found_q) == EXPECTED_POINTS)

        label_ok = p1_all_labeled and p2_q6_q25 and p3_distinct
        note = ('①20个散点均带可识别标签(带标签数=%d)=%s；②标签内容为Q6-Q25题号(缺失:%s,多余:%s)=%s；'
                '③能区分各维度(唯一题号数=%d)=%s'
                % (len(recognizable), p1_all_labeled,
                   ','.join(sorted(expected_q - found_q)) or '无',
                   ','.join(sorted(found_q - expected_q)) or '无', p2_q6_q25,
                   len(found_q), p3_distinct))
    report.add_dim2(1, '图表标签：20个散点均带有可识别的数据标签，标签内容为Q6-Q25题号，'
                       '能够区分各评估维度', label_ok, note)

    # ---- +1 图表标题：图表包含标题，标题含有"重要性-满意度象限图"，
    #            能明确说明图表用途 ----
    # 严格按细则逐点拆解，每一个点都必须踩到（全部满足才加分），
    # 不附加细则未要求的额外约束：
    #   点① 图表包含标题——存在非空的图表标题；
    #   点② 标题含有"重要性-满意度象限图"——标题文本包含该字样；
    #   点③ 能明确说明图表用途——标题含"重要性-满意度象限图"即明确说明了用途。
    title_ok = False
    note = ''
    if main_chart:
        title = main_chart['title']
        norm = title.replace(' ', '').replace('—', '-').replace('－', '-')

        # 点① 图表包含标题
        p1_has_title = (title.strip() != '')
        # 点② 标题含有"重要性-满意度象限图"
        p2_has_keyword = ('重要性-满意度象限图' in norm)
        # 点③ 能明确说明图表用途（含该关键字即明确说明用途）
        p3_clarifies = p2_has_keyword

        title_ok = p1_has_title and p2_has_keyword and p3_clarifies
        note = ('①图表包含标题=%s；②标题含"重要性-满意度象限图"=%s；③能明确说明图表用途=%s（标题="%s"）'
                % (p1_has_title, p2_has_keyword, p3_clarifies, title))
    report.add_dim2(1, '图表标题：图表包含标题，标题含有"重要性-满意度象限图"，'
                       '能明确说明图表用途', title_ok, note)

    # ---- +1 图内文字字体：图表标题、坐标轴标题、中文数据标签、中文图例或
    #            中文说明文字采用宋体小五号或9磅 ----
    # 严格按细则逐点拆解，每一个点都必须踩到（全部满足才加分），
    # 不附加细则未要求的额外约束：
    #   点① 图表标题——采用宋体，字号为小五号(9磅)；
    #   点② 坐标轴标题——采用宋体，字号为小五号(9磅)；
    #   点③ 中文数据标签——采用宋体，字号为小五号(9磅)；
    #   点④ 中文图例或中文说明文字——采用宋体，字号为小五号(9磅)。
    #   说明：小五号 = 9磅；OOXML 中 sz=900（百分之一磅）。
    cn_font_ok = False
    note = ''
    if main_chart:
        cn_font_ok, note = _check_chinese_font(main_chart['raw'])
    report.add_dim2(1, '图内文字字体：图表标题、坐标轴标题、中文数据标签、'
                       '中文图例或中文说明文字采用宋体小五号或9磅', cn_font_ok, note)

    # ---- +1 图内英文与数字字体：图表中的英文、阿拉伯数字、坐标刻度数字、
    #            均值数值、Q题号等采用Times New Roman小五号或9磅 ----
    # 严格按细则逐点拆解，每一个点都必须踩到（全部满足才加分），
    # 不附加细则未要求的额外约束：
    #   点① 英文——采用 Times New Roman，字号小五号(9磅)；
    #   点② 阿拉伯数字——采用 Times New Roman，字号小五号(9磅)；
    #   点③ 坐标刻度数字——坐标轴刻度文本采用 Times New Roman 9磅；
    #   点④ 均值数值——数据点/数值相关文本采用 Times New Roman 9磅；
    #   点⑤ Q题号——数据标签中的 Q 题号采用 Times New Roman 9磅。
    #   说明：小五号 = 9磅；OOXML 中 sz=900（百分之一磅）。
    en_font_ok = False
    note = ''
    if main_chart:
        en_font_ok, note = _check_western_font(main_chart['raw'])
    report.add_dim2(1, '图内英文与数字字体：图表中的英文、阿拉伯数字、坐标刻度数字、'
                       '均值数值、Q题号等采用Times New Roman小五号或9磅', en_font_ok, note)

    # ---- +1 图表坐标轴单位：横轴标题和纵轴标题均注明单位，
    #            如"体验均值（分）""关注度均值（分）"或等效表述 ----
    # 严格按细则逐点拆解，每一个点都必须踩到（全部满足才加分），
    # 不附加细则未要求的额外约束：
    #   点① 横轴标题注明单位——横轴标题中带有明确的单位标注
    #        （如"（分）""(分)""单位：分""/分"等模式，而非仅包含"分"字）；
    #   点② 纵轴标题注明单位——纵轴标题同上。
    #   说明：细则给出示例为"（分）"，并允许"等效表述"，但必须是明确的
    #         单位标注模式，不能仅因标题中出现"分"字（如"体验分布"）就通过。
    unit_ok = False
    note = ''
    if main_chart:
        x_title, y_title = _axis_titles_by_pos(pkg, main_chart)

        # 明确的单位标注模式：
        #   （分）/(分) —— 括号单位，允许括号内有其他字符（如"（满分5分）"）；
        #   单位：分/单位:分/单位为分 —— "单位"+分隔符+分；
        #   /分 —— 斜杠单位（如"分/题"这类度量单位写法，此处特指"XX/分"）。
        _UNIT_PATTERNS = (
            re.compile(r'[（(][^（）()]*分[^（）()]*[）)]'),   # （...分...）或(...分...)
            re.compile(r'单位[：:]?\s*[为是]?\s*分'),           # 单位：分 / 单位:分 / 单位为分
            re.compile(r'/\s*分\b'),                            # /分
        )

        def _has_unit(title):
            if not title:
                return False
            return any(p.search(title) for p in _UNIT_PATTERNS)

        # 点① 横轴标题注明单位
        p1_x_unit = _has_unit(x_title)
        # 点② 纵轴标题注明单位
        p2_y_unit = _has_unit(y_title)

        unit_ok = p1_x_unit and p2_y_unit
        note = ('①横轴标题注明单位=%s(横轴标题="%s")；②纵轴标题注明单位=%s(纵轴标题="%s")'
                % (p1_x_unit, x_title, p2_y_unit, y_title))
    report.add_dim2(1, '图表坐标轴单位：横轴标题和纵轴标题均注明单位，'
                       '如"体验均值（分）""关注度均值（分）"或等效表述', unit_ok, note)

    # ---- +1 图表布局：图表放置在原始数据表右侧、下方或其他空白区域，
    #            未覆盖A1:C22主要数据区域超过30% ----
    # 严格按细则逐点拆解，每一个点都必须踩到（全部满足才加分），
    # 不附加细则未要求的额外约束：
    #   点① 图表放置在原始数据表右侧、下方或其他空白区域——图表锚点位于
    #        A1:C22 的右侧（起始列在C列之后）或下方（起始行在第22行之后）
    #        或其他空白区域（与A1:C22无显著重叠）；
    #   点② 未覆盖A1:C22主要数据区域超过30%——图表与A1:C22重叠比例≤30%。
    layout_ok = False
    note = ''
    layout_ok, note = _check_layout(pkg)
    report.add_dim2(1, '图表布局：图表放置在原始数据表右侧、下方或其他空白区域，'
                       '未覆盖A1:C22主要数据区域超过30%', layout_ok, note)


# ----------------------------------------------------------------------------
# 维度2 辅助判定函数
# ----------------------------------------------------------------------------
def _all_close(values, target, tol=TOL):
    return all(abs(v - target) <= tol for v in values)


def _line_spans_full(line, axis_min, axis_max, data_lo, data_hi, along='y'):
    """判断一条参考线是否"贯穿全图"。
    要求（两者都满足，取用户选择的严格口径）：
      A. 线两端 ≈ 坐标轴 min/max（若图表设了固定轴范围）；
      B. 线的跨度同时覆盖全部数据点的范围 [data_lo, data_hi]。
    along='y'：纵向线，看其 yvals 的跨度；along='x'：横向线，看其 xvals 的跨度。
    任一参照缺失（如未设固定轴范围）时，该项视为不要求、自动通过，
    以免把"其实贯穿了、只是没设固定轴"的情况误判为未达标。"""
    span = line['yvals'] if along == 'y' else line['xvals']
    if not span:
        return False
    lo, hi = min(span), max(span)
    # A. 端点贴合坐标轴范围
    a_ok = True
    if axis_min is not None:
        a_ok = a_ok and (lo <= axis_min + TOL)
    if axis_max is not None:
        a_ok = a_ok and (hi >= axis_max - TOL)
    # B. 覆盖全部数据点范围
    b_ok = (lo <= data_lo + TOL) and (hi >= data_hi - TOL)
    return a_ok and b_ok


def _multiset_close(a, b, tol=TOL):
    """两组数值作为多重集合是否近似相等（长度相同且排序后逐一接近）。"""
    if len(a) != len(b):
        return False
    sa, sb = sorted(a), sorted(b)
    return all(abs(x - y) <= tol for x, y in zip(sa, sb))


def _axis_titles_by_pos(pkg, chart):
    """按 axPos 取横轴(b)与纵轴(l)标题文本。"""
    root = chart['raw']
    x_title, y_title = '', ''
    if root is None:
        return x_title, y_title
    for vax in root.findall('.//c:valAx', NS):
        pos_el = vax.find('c:axPos', NS)
        pos = pos_el.get('val') if pos_el is not None else ''
        t = text_of(vax.find('c:title', NS))
        if pos == 'b':
            x_title = t
        elif pos == 'l':
            y_title = t
    return x_title, y_title


def _font_sizes_ok(sz_str):
    """sz 以百分之一磅为单位，900 = 9磅 = 小五号。"""
    try:
        return int(sz_str) == 900
    except Exception:                                # noqa
        return False


def _check_chinese_font(root):
    """中文文字字体细则逐点检查：图表标题、坐标轴标题、中文数据标签、
    中文图例或中文说明文字，均应采用宋体小五号(9磅，sz=900)。
    rubric 要求列出的四类中，只要该类实际存在文本，就必须整体满足
    "宋体+9磅"；缺失该类样式（如未单独设字体/字号，走全局默认）或
    非宋体/非9磅，均判该类不通过——不能因为"整体至少一类通过"就整体合格。
    """
    if root is None:
        return False, '无图表XML'

    def _rpr_iter(scope):
        """遍历某作用域下所有 defRPr/rPr 文本样式。"""
        if scope is None:
            return
        for rpr in scope.iter():
            tag = rpr.tag.split('}')[-1]
            if tag in ('defRPr', 'rPr'):
                yield rpr

    def _has_text(scope):
        """该作用域下是否存在实际文本内容（a:t 非空）。"""
        if scope is None:
            return False
        return any((t.text or '').strip() != '' for t in scope.iter('{%s}t' % NS['a']))

    def _has_chinese_text(scope):
        """该作用域下是否存在中文文本内容（用于数据标签/图例类目——
        rubric 仅要求"中文"数据标签、"中文"图例采用宋体，纯英文/数字标签
        不适用本条中文字体规则，避免误伤未使用中文标签的图表）。"""
        if scope is None:
            return False
        for t in scope.iter('{%s}t' % NS['a']):
            txt = t.text or ''
            if any('一' <= ch <= '鿿' for ch in txt):
                return True
        return False

    def _check_scope(scope, require_chinese=False):
        """返回 (该作用域是否存在需要判定的文本, 是否整体满足宋体9磅,
        宋体样式数, 字号不符数, 非宋体样式数)。
        判定口径：只要该作用域存在（中）文文本，就要求其上出现的
        每一个显式字体/字号样式（defRPr/rPr）都是宋体且9磅；
        存在非宋体样式、或宋体但字号非9磅，均视为不通过。"""
        text_present = _has_chinese_text(scope) if require_chinese else _has_text(scope)
        if not text_present:
            return False, False, 0, 0, 0
        total = 0           # 设了字体的样式数
        songti = 0          # 宋体样式数
        bad = 0             # 宋体但字号非9磅数
        non_songti = 0      # 非宋体样式数
        for rpr in _rpr_iter(scope):
            latin = rpr.find('a:latin', NS)
            ea = rpr.find('a:ea', NS)
            tf_latin = latin.get('typeface') if latin is not None else None
            tf_ea = ea.get('typeface') if ea is not None else None
            sz = rpr.get('sz')
            has_font = (tf_latin is not None or tf_ea is not None)
            is_songti = (tf_latin == '宋体' or tf_ea == '宋体')
            if has_font:
                total += 1
                if is_songti:
                    songti += 1
                    if not _font_sizes_ok(sz):
                        bad += 1
                else:
                    non_songti += 1
        # 存在文本但未显式设置任何字体样式（走全局默认），视为不满足宋体9磅要求
        ok = (total > 0) and (non_songti == 0) and (bad == 0)
        return True, ok, songti, bad, non_songti

    # 点① 图表标题
    title_scope = root.find('.//c:chart/c:title', NS)
    # 点② 坐标轴标题
    axis_title_scopes = []
    for ax_tag in ('c:valAx', 'c:catAx'):
        for ax in root.findall('.//' + ax_tag, NS):
            t = ax.find('c:title', NS)
            if t is not None:
                axis_title_scopes.append(t)
    # 点③ 中文数据标签
    dlbls_scopes = root.findall('.//c:ser/c:dLbls', NS) + root.findall('.//c:plotArea/c:dLbls', NS)
    # 点④ 中文图例或中文说明文字
    legend_scope = root.find('.//c:legend', NS)

    results = []   # (名称, present, ok, songti, bad, non_songti)

    p, ok, s, b, ns = _check_scope(title_scope)
    results.append(('图表标题', p, ok, s, b, ns))

    # 坐标轴标题：合并所有轴标题作用域
    if axis_title_scopes:
        tot_s = tot_b = tot_ns = 0
        any_present = False
        all_ok = True
        for sc in axis_title_scopes:
            pp, oo, ss, bb, nn = _check_scope(sc)
            if pp:
                any_present = True
                all_ok = all_ok and oo
            tot_s += ss
            tot_b += bb
            tot_ns += nn
        results.append(('坐标轴标题', any_present, any_present and all_ok, tot_s, tot_b, tot_ns))
    else:
        results.append(('坐标轴标题', False, False, 0, 0, 0))

    # 中文数据标签：合并所有数据标签作用域，仅当标签含中文文本时才纳入判定
    if dlbls_scopes:
        tot_s = tot_b = tot_ns = 0
        any_present = False
        all_ok = True
        for sc in dlbls_scopes:
            pp, oo, ss, bb, nn = _check_scope(sc, require_chinese=True)
            if pp:
                any_present = True
                all_ok = all_ok and oo
            tot_s += ss
            tot_b += bb
            tot_ns += nn
        results.append(('中文数据标签', any_present, any_present and all_ok, tot_s, tot_b, tot_ns))
    else:
        results.append(('中文数据标签', False, False, 0, 0, 0))

    # 中文图例/说明文字：仅当图例含中文文本时才纳入判定
    p, ok, s, b, ns = _check_scope(legend_scope, require_chinese=True)
    results.append(('中文图例/说明文字', p, ok, s, b, ns))

    # rubric 列出的四类中，实际存在文本的每一类都必须整体满足宋体9磅；
    # 缺失该类文本（如坐标轴本身没有标题）视为该类不适用、不参与判定；
    # 但只要存在中文/文本却未设宋体9磅（含未显式设字体、或字体非宋体、或宋体但字号不对），
    # 该类即不通过，整体也不通过——不再是"至少一类满足即通过"。
    present_items = [r for r in results if r[1]]
    all_present_ok = all(r[2] for r in present_items)
    overall = len(present_items) > 0 and all_present_ok

    def _reason(present, ok, bad, non_songti):
        if not present:
            return '(无中文/无文本，不适用)'
        if ok:
            return '(宋体9磅✓)'
        if non_songti > 0:
            return '(存在非宋体样式:%d)' % non_songti
        if bad > 0:
            return '(宋体但字号非9磅:%d)' % bad
        return '(未显式设置字体，视为不合规)'

    note = '；'.join(
        '%s%s' % (name, _reason(present, ok, bad, non_songti))
        for (name, present, ok, _songti, bad, non_songti) in results)
    return overall, note


def _check_western_font(root):
    """英文与数字字体细则逐点检查：图表中的英文、阿拉伯数字、坐标刻度数字、
    均值数值、Q题号等，均应采用 Times New Roman 小五号(9磅，sz=900)。
    rubric 要求列出的类别中，只要该类实际存在英文/数字文本，就必须整体满足
    "TNR+9磅"；缺失该类显式字体样式（走全局默认）或字体非TNR（如 Arial），
    均判该类不通过——不能只统计已是TNR的样式数而忽略同一作用域内的非TNR样式。
      · 坐标刻度数字 —— 坐标轴(valAx/catAx)上 txPr 中的刻度文本样式；
      · Q题号       —— 数据标签(dLbls)中的文本样式；
      · 英文/阿拉伯数字/均值数值 —— 其余 Times New Roman 文本样式（标题、图例等）。
    """
    if root is None:
        return False, '无图表XML'

    TNR = 'Times New Roman'

    def _has_western_text(scope):
        """该作用域下是否存在英文字母或阿拉伯数字文本内容（a:t 非空）。"""
        if scope is None:
            return False
        for t in scope.iter('{%s}t' % NS['a']):
            txt = t.text or ''
            if any(ch.isascii() and (ch.isalpha() or ch.isdigit()) for ch in txt):
                return True
        return False

    def _scan(scope):
        """统计某作用域内：是否存在英文/数字文本、TNR样式数、
        非TNR样式数、TNR但字号非9磅数。"""
        text_present = _has_western_text(scope)
        if not text_present:
            return False, 0, 0, 0
        tnr = 0
        non_tnr = 0
        bad = 0
        for rpr in scope.iter():
            tag = rpr.tag.split('}')[-1]
            if tag not in ('defRPr', 'rPr'):
                continue
            latin = rpr.find('a:latin', NS)
            tf = latin.get('typeface') if latin is not None else None
            sz = rpr.get('sz')
            if tf is None:
                continue
            if tf == TNR:
                tnr += 1
                if not _font_sizes_ok(sz):
                    bad += 1
            else:
                non_tnr += 1
        return True, tnr, non_tnr, bad

    # 点③ 坐标刻度数字：坐标轴的刻度文本样式（txPr），排除轴标题(title)
    tick_present = False
    tick_tnr = tick_non_tnr = tick_bad = 0
    for ax_tag in ('c:valAx', 'c:catAx'):
        for ax in root.findall('.//' + ax_tag, NS):
            txpr = ax.find('c:txPr', NS)
            p, t, nt, b = _scan(txpr)
            tick_present = tick_present or p
            tick_tnr += t
            tick_non_tnr += nt
            tick_bad += b

    # 点⑤ Q题号：数据标签中的文本样式
    qno_present = False
    qno_tnr = qno_non_tnr = qno_bad = 0
    for dlbls in (root.findall('.//c:ser/c:dLbls', NS)
                  + root.findall('.//c:plotArea/c:dLbls', NS)):
        p, t, nt, b = _scan(dlbls)
        qno_present = qno_present or p
        qno_tnr += t
        qno_non_tnr += nt
        qno_bad += b

    # 点①②④ 英文/阿拉伯数字/均值数值：全图范围（覆盖标题、图例及其它文本）
    all_present, all_tnr, all_non_tnr, all_bad = _scan(root)

    items = [
        ('英文/阿拉伯数字/均值数值', all_present, all_tnr, all_non_tnr, all_bad),
        ('坐标刻度数字', tick_present, tick_tnr, tick_non_tnr, tick_bad),
        ('Q题号', qno_present, qno_tnr, qno_non_tnr, qno_bad),
    ]
    present_items = [it for it in items if it[1]]
    # 存在英文/数字文本的类别，须字体全部为TNR(non_tnr==0)且字号全部为9磅(bad==0)
    all_present_ok = all((non_tnr == 0 and bad == 0) for (_n, _p, _t, non_tnr, bad) in present_items)
    overall = len(present_items) > 0 and all_present_ok

    def _reason(present, tnr, non_tnr, bad):
        if not present:
            return '(无英文/数字文本，不适用)'
        if non_tnr == 0 and bad == 0:
            return '(TNR9磅✓,样式%d)' % tnr
        if non_tnr > 0:
            return '(存在非TNR样式:%d)' % non_tnr
        return '(TNR但字号非9磅:%d)' % bad

    note = '；'.join(
        '%s%s' % (name, _reason(present, tnr, non_tnr, bad))
        for (name, present, tnr, non_tnr, bad) in items)
    return overall, note


def _check_layout(pkg):
    """图表布局细则逐点检查：
      点① 图表放置在原始数据表右侧、下方或其他空白区域——
           图表锚点起始列在C列之后(右侧)、或起始行在第22行之后(下方)、
           或与A1:C22无显著重叠(其他空白区域)；
      点② 未覆盖A1:C22主要数据区域超过30%——与A1:C22重叠比例≤30%。
    A1:C22 = 列 0..2（A-C），行 0..21（1-22，0基）。
    覆盖比例 = 图表与该区域的重叠面积 / 该区域面积。"""
    target_c0, target_c1 = 0, 2     # A..C (0-based, inclusive)
    target_r0, target_r1 = 0, 21    # row1..row22 (0-based)
    target_area = (target_c1 - target_c0 + 1) * (target_r1 - target_r0 + 1)

    for n in pkg.names:
        if not re.match(r'xl/drawings/drawing\d+\.xml$', n):
            continue
        root = pkg.read_xml(n)
        if root is None:
            continue
        for anchor in root.iter('{%s}twoCellAnchor' % NS['xdr']):
            # 仅考虑含图表的锚点
            gf = anchor.find('xdr:graphicFrame', NS)
            if gf is None or anchor.find('.//c:chart', NS) is None:
                continue
            f = anchor.find('xdr:from', NS)
            t = anchor.find('xdr:to', NS)
            if f is None or t is None:
                continue
            fc = int(f.find('xdr:col', NS).text)
            fr = int(f.find('xdr:row', NS).text)
            tc = int(t.find('xdr:col', NS).text)
            tr = int(t.find('xdr:row', NS).text)
            # 重叠区域
            oc0, oc1 = max(fc, target_c0), min(tc, target_c1)
            or0, or1 = max(fr, target_r0), min(tr, target_r1)
            overlap = 0
            if oc0 <= oc1 and or0 <= or1:
                overlap = (oc1 - oc0 + 1) * (or1 - or0 + 1)
            ratio = overlap / target_area if target_area else 0

            # 点① 放置在右侧/下方/其他空白区域
            placed_right = fc > target_c1          # 起始列在C列之后 -> 右侧
            placed_below = fr > target_r1          # 起始行在第22行之后 -> 下方
            placed_blank = overlap == 0            # 与A1:C22无重叠 -> 其他空白区域
            p1_placement = placed_right or placed_below or placed_blank
            # 点② 未覆盖A1:C22超过30%
            p2_overlap_ok = ratio <= 0.30

            ok = p1_placement and p2_overlap_ok
            where = ('右侧' if placed_right else '') + ('下方' if placed_below else '') \
                    + ('空白区' if placed_blank else '')
            note = ('①放置在右侧/下方/空白区=%s(位置:%s,锚点列[%d-%d]行[%d-%d])；'
                    '②未覆盖A1:C22超过30%%=%s(重叠比例=%.0f%%)'
                    % (p1_placement, where or '与数据区有重叠', fc, tc, fr, tr,
                       p2_overlap_ok, ratio * 100))
            return ok, note
    return False, '未找到图表锚点'


# ----------------------------------------------------------------------------
# 主流程
# ----------------------------------------------------------------------------
# 脚本编号（对应目录 officeval_098；返回给汇总侧的 id 使用 3 位补零形式）
SCRIPT_ID = '098'
# 该脚本负责评估的目标文档（若目录内不存在同名文件，则退化为扫描目录里的 .xlsx/.xlsm）
PREFERRED_FILENAME = '满意度调查-指标分析_重要性满意度象限图.xlsx'


def _locate_target(dir_path: str) -> str:
    """在给定目录中定位待评估的 .xlsx/.xlsm 文档。
    优先使用与题目一致的文件名；若不存在则退化为该目录下首个 .xlsx/.xlsm 文件。"""
    if not os.path.isdir(dir_path):
        raise FileNotFoundError('目录不存在：%s' % dir_path)
    preferred = os.path.join(dir_path, PREFERRED_FILENAME)
    if os.path.isfile(preferred):
        return preferred
    for name in sorted(os.listdir(dir_path)):
        low = name.lower()
        if low.endswith(('.xlsx', '.xlsm')) and not name.startswith('~$'):
            return os.path.join(dir_path, name)
    raise FileNotFoundError('目录中未找到 .xlsx/.xlsm 文档：%s' % dir_path)


def evaluate(dir_path: str) -> dict:  # type: ignore[type-arg]
    """统一入口：接收"脚本所在目录的路径"，由脚本自身在该目录里定位并打开被评估的文档。
    返回结构化 dict（字段含义见《脚本接口差异与统一建议.md》§2.2）。"""
    result = {
        'id': SCRIPT_ID,
        'file_name': None,
        'status': 'ok',
        'error': None,
        'dim1_pass': True,
        'dim1_reason': '',
        'dim2_items': [],
        'total_score': 0,
        'max_score': 0,
    }

    target = None
    eval_path = None
    try:
        target = _locate_target(dir_path)
        result['file_name'] = os.path.basename(target)

        report = Report()
        eval_path = target
        pkg = XlsxPackage(eval_path)

        if not pkg.ok_open:
            # 连包都打不开：维度1第一条即失败
            report.add_dim1(
                '交付文件为xlsx或.xlsm格式，文件可正常打开', False,
                '无法正常打开：%s' % getattr(pkg, 'open_error', '未知'))
        else:
            shared = load_shared_strings(pkg)
            sheets = get_sheet_map(pkg)

            sheet1_cells, sheet1_dim = ({}, None)
            if 'Sheet1' in sheets:
                sheet1_cells, sheet1_dim = read_sheet_cells(pkg, sheets['Sheet1'], shared)

            chart_parts = find_chart_parts(pkg)
            charts = [parse_chart(pkg, cp) for cp in chart_parts]

            # 维度1
            check_dimension1(pkg, target, sheets, sheet1_cells, sheet1_dim, charts, report)

            # 维度1未全部通过 -> 直接 0 分，不查维度2
            if not report.dim1_failed:
                check_dimension2(pkg, sheets, sheet1_cells, charts, report)

            pkg.close()

        # ---- 组装返回结构 ----
        result['dim1_pass'] = not report.dim1_failed
        if report.dim1_failed:
            reasons = []
            for desc, passed, note in report.dim1_checks:
                if not passed:
                    reasons.append(desc + (('（' + note + '）') if note else ''))
            result['dim1_reason'] = '；'.join(reasons)

        dim2_items = []
        for score, desc, hit, note in report.dim2_items:
            dim2_items.append({
                'rule': desc,
                'max_delta': score,
                'delta': score if hit else 0,
                'hit': hit,
                'detail': '',
            })
        result['dim2_items'] = dim2_items

        # 总分 = 所有加分项 delta 之和；未通过维度一则强制为 0；不保留小数
        total = 0 if report.dim1_failed else sum(int(it['delta']) for it in dim2_items)
        result['total_score'] = int(total)
        result['max_score'] = int(sum(
            int(it['max_delta']) for it in dim2_items if int(it['max_delta']) > 0
        ))
    except Exception as e:                               # noqa
        # 脚本自身异常（含"目录/文件不存在"）—— 归入 status=error，与"通过维度一但0分"区分。
        # file_name 定位失败时保持契约要求的 None（不回退为空字符串）；
        # 同时把 dim1_pass 显式置为 False 并写明原因，避免 error 状态下
        # dim1_pass 仍停留在初始化的 True、与 status=error 的语义相互矛盾。
        result['status'] = 'error'
        result['error'] = str(e)
        result['dim1_pass'] = False
        result['dim1_reason'] = result['dim1_reason'] or str(e)
        result['total_score'] = 0

    return result


if __name__ == '__main__':
    # 本地调试入口：默认以脚本所在目录作为 dir_path；也支持通过 argv[1] 指定其它目录。
    # 主结果只走 return + 此处的 json 打印，不在 evaluate 内部 print、不改 sys.stdout、不 sys.exit。
    default_dir = os.path.dirname(os.path.abspath(__file__))
    target_dir = sys.argv[1] if len(sys.argv) > 1 else default_dir
    print(json.dumps(evaluate(target_dir), ensure_ascii=False, indent=2))
