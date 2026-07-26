# -*- coding: utf-8 -*-
"""
自动评估脚本（officeval_095）：对“乡村公共服务调查汇总_已添加环状图”xlsx 按打分细则评估。

对外接口：
    evaluate(dir_path: str) -> dict
调用方只传入“脚本所在目录的路径”，脚本自己在该目录内定位并打开被评估的 .xlsx/.xlsm 文件，
返回结构化字典（不 print 主结果、不改 sys.stdout、不 sys.exit）。

评估逻辑：
  维度1（可用与可修改性）：任何一条不满足 -> 直接 0 分，不再检查维度2。
  维度2（完成度）：得分点（每条全部子条件满足才加分）+ 扣分点（任意一条命中即扣分），累加得最终分。

实现方式：直接解析 xlsx（本质是 zip 包），读取 sheet/chart/drawing 的 XML，
不依赖人工判断；对难以严格判定的点采用合理变通（见各检查项注释）。
"""

import sys
import os
import re
import json
import zipfile
import xml.etree.ElementTree as ET

# 脚本编号（与目录 officeval_095 对应）
SCRIPT_ID = "095"

# ---------------------------------------------------------------------------
# 命名空间
# ---------------------------------------------------------------------------
NS = {
    'x':   'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
    'c':   'http://schemas.openxmlformats.org/drawingml/2006/chart',
    'a':   'http://schemas.openxmlformats.org/drawingml/2006/main',
    'xdr': 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing',
    'r':   'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    'rel': 'http://schemas.openxmlformats.org/package/2006/relationships',
    'ct':  'http://schemas.openxmlformats.org/package/2006/content-types',
}

# 第一个问题的 5 个回答选项与期望比例（来自 D3:D7，四舍五入百分比）
EXPECTED_OPTIONS = ['十分清楚', '比较清楚', '有所了解', '了解较少', '从未关注']
EXPECTED_PCTS = [0.11888111888111888, 0.23426573426573427, 0.3181818181818182,
                 0.24475524475524477, 0.08391608391608392]


class WorkbookData:
    """加载并缓存 xlsx 内部各 XML 部件。"""

    def __init__(self, path):
        self.path = path
        self.ok_open = False
        self.parts = {}          # part name -> raw bytes
        self.sheet_root = None   # sheet1.xml root
        self.chart_roots = []    # list of chart XML roots
        self.drawing_roots = []  # list of drawing XML roots
        self.rels_roots = {}     # rels part name -> relationships XML root
        self.theme_root = None   # theme1.xml root（用于解析 schemeClr）
        self.workbook_root = None
        self._load()

    def _load(self):
        if not zipfile.is_zipfile(self.path):
            return
        try:
            with zipfile.ZipFile(self.path) as z:
                bad = z.testzip()
                if bad is not None:
                    return
                for name in z.namelist():
                    try:
                        self.parts[name] = z.read(name)
                    except Exception:
                        self.parts[name] = b''
            self.ok_open = True
        except Exception:
            return

        # 解析常用部件
        if 'xl/workbook.xml' in self.parts:
            self.workbook_root = self._parse('xl/workbook.xml')
        if 'xl/theme/theme1.xml' in self.parts:
            self.theme_root = self._parse('xl/theme/theme1.xml')
        # 第一个工作表（按 r:id 找较繁琐，这里 sheet1.xml 即为 Sheet1）
        sheet_part = self._find_first_sheet_part()
        if sheet_part:
            self.sheet_root = self._parse(sheet_part)
        for name in self.parts:
            if re.match(r'xl/(drawings/)?charts/chart\d+\.xml$', name):
                root = self._parse(name)
                if root is not None:
                    self.chart_roots.append((name, root))
            if re.match(r'xl/drawings/drawing\d+\.xml$', name):
                root = self._parse(name)
                if root is not None:
                    self.drawing_roots.append((name, root))
            if re.match(r'.*/_rels/[^/]+\.rels$', name):
                root = self._parse(name)
                if root is not None:
                    self.rels_roots[name] = root

    def resolve_rel(self, part_name, r_id):
        """将某个 part（如 xl/drawings/drawing1.xml）里的 r:id 解析为其目标 part 的规范路径。"""
        if not r_id:
            return None
        d = part_name.rsplit('/', 1)[0] if '/' in part_name else ''
        base = part_name.rsplit('/', 1)[1] if '/' in part_name else part_name
        rels_name = (d + '/_rels/' + base + '.rels') if d else ('_rels/' + base + '.rels')
        rels_root = self.rels_roots.get(rels_name)
        if rels_root is None:
            return None
        for rel in rels_root.findall('rel:Relationship', NS):
            if rel.get('Id') == r_id:
                target = rel.get('Target')
                if not target:
                    return None
                if target.startswith('/'):
                    return target.lstrip('/')
                # Target 相对于 part 所在目录解析（可能含 ../）
                return self._normalize_path(d, target)
        return None

    @staticmethod
    def _normalize_path(base_dir, rel_target):
        parts = base_dir.split('/') if base_dir else []
        for seg in rel_target.split('/'):
            if seg == '..':
                if parts:
                    parts.pop()
            elif seg == '.' or seg == '':
                continue
            else:
                parts.append(seg)
        return '/'.join(parts)

    def _find_first_sheet_part(self):
        # 优先 sheet1.xml
        for cand in ['xl/worksheets/sheet1.xml']:
            if cand in self.parts:
                return cand
        for name in sorted(self.parts):
            if re.match(r'xl/worksheets/sheet\d+\.xml$', name):
                return name
        return None

    def _parse(self, name):
        try:
            return ET.fromstring(self.parts[name])
        except Exception:
            return None


# ---------------------------------------------------------------------------
# 工作表单元格读取工具
# ---------------------------------------------------------------------------
def split_ref(ref):
    m = re.match(r'([A-Z]+)(\d+)', ref)
    return m.group(1), int(m.group(2))


def build_cell_map(sheet_root):
    """返回 {cellRef: {'v':值文本, 't':类型, 'f':公式}} 。"""
    cells = {}
    if sheet_root is None:
        return cells
    data = sheet_root.find('x:sheetData', NS)
    if data is None:
        return cells
    for row in data.findall('x:row', NS):
        for c in row.findall('x:c', NS):
            ref = c.get('r')
            t = c.get('t')
            v_el = c.find('x:v', NS)
            f_el = c.find('x:f', NS)
            cells[ref] = {
                'v': v_el.text if v_el is not None else None,
                't': t,
                'f': f_el.text if f_el is not None else None,
            }
    return cells


# ---------------------------------------------------------------------------
# 维度1：可用与可修改性
# ---------------------------------------------------------------------------
def check_dimension1(wb):
    """返回 (passed: bool, details: list[(ok, msg)])。任一 False 即维度1不通过。"""
    d: "list[tuple[bool, str]]" = []

    # 1.1 交付文件为 xlsx/xlsm 格式，且可正常打开
    ext = os.path.splitext(wb.path)[1].lower()
    ext_ok = ext in ('.xlsx', '.xlsm')
    d.append((bool(ext_ok and wb.ok_open),
              "交付文件为 %s 格式且可正常打开（zip 结构完整、XML 可解析）" % ext))

    passed = all(ok for ok, _ in d)
    return passed, d


# ---------------------------------------------------------------------------
# 图表分析工具
# ---------------------------------------------------------------------------
def title_box_geometry(chart):
    """计算图表标题（中心文本）文本框相对绘图区的实际几何。

    返回 dict：{
        'has_layout': bool,          # 是否存在 manualLayout
        'xmode','ymode': str,        # 'edge' / 'factor'
        'x','y','w','h': float|None, # manualLayout 原始值
        'cx','cy': float|None,       # 文本框实际中心（相对绘图区，0~1）
        'detail': str,
    }

    坐标语义（OOXML c:manualLayout）：
      - edge   : x/y 为相对绘图区左上角的绝对位置 → 中心 = (x+w/2, y+h/2)
      - factor : x/y 为相对“默认位置”的偏移量。overlay=1 的标题默认水平/垂直居中，
                 默认中心约为绘图区 (0.5, 0.5) → 实际中心 = (0.5+x, 0.5+y)。
                 （这正是 x=0.34,y=0.40 会把文本推到偏右下、离开中空圆的原因。）

    无 manualLayout 时：Excel 对 overlay=1 的标题采用默认居中定位，此时不能因为
    x/y 缺失就判定“未居中”——退化为默认中心 (0.5, 0.5)，宽高未知记为 0（不做
    尺寸溢出判断），交由调用方结合 overlay 值判断是否真正居中叠加。
    """
    res = {'has_layout': False, 'xmode': None, 'ymode': None,
           'x': None, 'y': None, 'w': 0.0, 'h': 0.0,
           'cx': None, 'cy': None, 'detail': '无 manualLayout'}
    if chart is None:
        res['detail'] = '无图表'
        return res
    ml = chart.find('.//c:title//c:manualLayout', NS)
    if ml is None:
        overlay = chart.find('.//c:title/c:overlay', NS)
        if overlay is not None and overlay.get('val') == '1':
            # 无手动布局但 overlay=1：Excel 默认将标题居中叠加到绘图区
            res['cx'] = 0.5
            res['cy'] = 0.5
            res['detail'] = '无 manualLayout，overlay=1 → 按默认居中(0.5,0.5)推断'
        return res

    def _val(tag):
        el = ml.find('c:%s' % tag, NS)
        return el.get('val') if el is not None else None

    def _flt(tag, default=None):
        raw = _val(tag)
        if raw is None:
            return default
        try:
            return float(raw)
        except (TypeError, ValueError):
            return default

    res['has_layout'] = True
    res['xmode'] = (_val('xMode') or 'factor').lower()   # 缺省即 factor
    res['ymode'] = (_val('yMode') or 'factor').lower()
    res['x'] = _flt('x')
    res['y'] = _flt('y')
    res['w'] = _flt('w', 0.0)
    res['h'] = _flt('h', 0.0)

    if res['x'] is not None and res['y'] is not None:
        if res['xmode'] == 'edge':
            res['cx'] = res['x'] + res['w'] / 2.0
        else:   # factor：相对默认居中位置(0.5)的偏移
            res['cx'] = 0.5 + res['x']
        if res['ymode'] == 'edge':
            res['cy'] = res['y'] + res['h'] / 2.0
        else:
            res['cy'] = 0.5 + res['y']
        res['detail'] = ('xMode=%s,yMode=%s,x=%s,y=%s → 实际中心(cx=%.3f,cy=%.3f)'
                         % (res['xmode'], res['ymode'], res['x'], res['y'],
                            res['cx'], res['cy']))
    else:
        res['detail'] = ('xMode=%s,yMode=%s,x/y 缺失，无法定位'
                         % (res['xmode'], res['ymode']))
    return res


def get_doughnut_chart(wb):
    for name, root in wb.chart_roots:
        if root.find('.//c:doughnutChart', NS) is not None:
            return name, root
    # 没有 doughnut，则返回第一个图表（用于判断是否为饼图等）
    if wb.chart_roots:
        return wb.chart_roots[0]
    return None, None


def chart_refs(root):
    """提取系列的 cat / val 引用字符串。"""
    cat = val = None
    cat_el = root.find('.//c:ser/c:cat//c:f', NS)
    val_el = root.find('.//c:ser/c:val//c:f', NS)
    if cat_el is not None:
        cat = cat_el.text
    if val_el is not None:
        val = val_el.text
    return cat, val


def resolve_ref_to_options(cells, ref):
    """将形如 'Sheet1'!$M$3:$M$7 的引用解析为对应取值列表（跟随单元格内公式/值）。"""
    if not ref:
        return []
    m = re.search(r'\$?([A-Z]+)\$?(\d+):\$?([A-Z]+)\$?(\d+)', ref)
    if not m:
        return []
    c1, r1, c2, r2 = m.group(1), int(m.group(2)), m.group(3), int(m.group(4))
    out = []
    for rr in range(r1, r2 + 1):
        info = cells.get('%s%d' % (c1, rr), {})
        out.append(info.get('v'))
    return out


def ref_resolves_to(cells, ref, target_col, target_r1, target_r2):
    """判断图表数据源引用是否等价于 target_col 的 target_r1:target_r2 区域。

    满足任一即视为来自该区域：
      1) 引用本身就是 target_col$target_r1:target_col$target_r2（直接引用 B3:B7 / D3:D7）；
      2) 引用指向某辅助列，但该辅助列各单元格通过公式逐行回溯到 target_col 的对应行
         （如 M3=B3, M4=B4, ... 或 N3=D3, ...）。
    """
    if not ref:
        return False
    m = re.search(r'\$?([A-Z]+)\$?(\d+):\$?([A-Z]+)\$?(\d+)', ref)
    if not m:
        return False
    col1, r1, col2, r2 = m.group(1), int(m.group(2)), m.group(3), int(m.group(4))
    if col1 != col2:
        return False
    # 区域行数需与目标一致
    if (r2 - r1) != (target_r2 - target_r1):
        return False

    # 情况1：直接引用目标列目标行
    if col1 == target_col and r1 == target_r1 and r2 == target_r2:
        return True

    # 情况2：辅助列逐行公式回溯到目标列对应行
    for offset in range(0, r2 - r1 + 1):
        cell_ref = '%s%d' % (col1, r1 + offset)
        f = (cells.get(cell_ref, {}).get('f') or '').strip()
        expected = '%s%d' % (target_col, target_r1 + offset)
        # 公式应等价于直接引用目标单元格（允许 = 前缀与 $ 绝对引用符）
        f_norm = f.lstrip('=').replace('$', '').upper()
        if f_norm != expected:
            return False
    return True


# ---------------------------------------------------------------------------
# 维度2：完成度评分
# ---------------------------------------------------------------------------
def check_dimension2(wb, cells):
    """返回 (total, hits: list[(item_id, score, ok, msg)])。

    item_id 对应 _DIM2_RULES 中的稳定标识，用于在 evaluate() 中精确对齐每个评分项，
    避免同分值多个扣分项（如两个 -3）靠 score+文本前缀匹配时相互冲突。
    """
    hits: list[tuple[str, int, bool, str]] = []
    chart = get_doughnut_chart(wb)[1]

    is_doughnut = chart is not None and chart.find('.//c:doughnutChart', NS) is not None
    hole = None
    if is_doughnut:
        he = chart.find('.//c:doughnutChart/c:holeSize', NS)
        if he is not None:
            hole = int(he.get('val', '0'))

    cat_ref, val_ref = chart_refs(chart) if chart is not None else (None, None)
    cat_vals = resolve_ref_to_options(cells, cat_ref)
    # 解析 val 引用对应比例
    val_pcts_cells = []
    if val_ref:
        m = re.search(r'\$?([A-Z]+)\$?(\d+):\$?([A-Z]+)\$?(\d+)', val_ref)
        if m:
            c1, r1, r2 = m.group(1), int(m.group(2)), int(m.group(4))
            for rr in range(r1, r2 + 1):
                info = cells.get('%s%d' % (c1, rr), {})
                try:
                    val_pcts_cells.append(float(info.get('v')))
                except (TypeError, ValueError):
                    val_pcts_cells.append(None)

    # ---- +3：逐条对应细则的每一个点 ----
    # 细则点①：Sheet1 第一个问题右侧空白区域新增“1个”环状图
    # 细则点②：图表数据源来自 B3:B7 回答选项 和 D3:D7 比例数据
    # 细则点③：环状图包含第一个问题的 5 个回答选项（十分清楚/比较清楚/有所了解/了解较少/从未关注）
    # 细则点④：各扇区占比与 D3:D7 基本一致
    # 细则点⑤：环状图为圆环样式，具有明显中空区域
    # 细则点⑥：不是普通饼图、柱状图、折线图或散点图

    # 点①：右侧空白区域 + 恰好新增 1 个环状图
    right_side = chart_on_right(wb)                      # drawing 锚点起始列在第一个问题右侧
    doughnut_count = 0
    for _, croot in wb.chart_roots:
        if croot.find('.//c:doughnutChart', NS) is not None:
            doughnut_count += 1
    p1_one_ring_right = right_side and doughnut_count == 1

    # 点②：数据源来自 B3:B7 与 D3:D7（直接引用，或经辅助列公式等价回溯到 B3:B7/D3:D7）
    src_from_B3B7 = ref_resolves_to(cells, cat_ref, 'B', 3, 7)
    src_from_D3D7 = ref_resolves_to(cells, val_ref, 'D', 3, 7)
    p2_source = src_from_B3B7 and src_from_D3D7

    # 点③：包含 5 个指定回答选项
    p3_five_options = ([v for v in cat_vals] == EXPECTED_OPTIONS)

    # 点④：各扇区占比与 D3:D7 基本一致
    p4_pcts = False
    if len(val_pcts_cells) == 5 and all(v is not None for v in val_pcts_cells):
        p4_pcts = all(abs(val_pcts_cells[i] - EXPECTED_PCTS[i]) < 0.01 for i in range(5))

    # 点⑤：圆环样式，具有明显中空区域（doughnutChart 且 holeSize > 0）
    p5_ring = is_doughnut and (hole is not None and hole > 0)

    # 点⑥：不是普通饼图/柱状图/折线图/散点图
    has_other_chart = chart is not None and (
        chart.find('.//c:pieChart', NS) is not None
        or chart.find('.//c:bar3DChart', NS) is not None
        or chart.find('.//c:barChart', NS) is not None
        or chart.find('.//c:lineChart', NS) is not None
        or chart.find('.//c:line3DChart', NS) is not None
        or chart.find('.//c:scatterChart', NS) is not None)
    p6_not_other = is_doughnut and not has_other_chart

    plus3_ok = (p1_one_ring_right and p2_source and p3_five_options
                and p4_pcts and p5_ring and p6_not_other)
    hits.append(("plus3_ring", 3, plus3_ok,
                 "Sheet1第一个问题右侧空白区新增1个环状图，数据源来自B3:B7/D3:D7，含5个回答选项，"
                 "各扇区占比与D3:D7基本一致，圆环样式有明显中空，非饼/柱/折线/散点图"
                 " [①右侧空白1个环图=%s, ②数据源B3:B7&D3:D7=%s, ③含5个回答选项=%s, "
                 "④占比与D3:D7一致=%s, ⑤圆环中空(hole=%s)=%s, ⑥非饼/柱/折线/散点=%s]"
                 % (p1_one_ring_right, p2_source, p3_five_options,
                    p4_pcts, hole, p5_ring, p6_not_other)))

    # ---- +1：逐条对应细则的每一个点 ----
    # 细则点①：环状图中心位置放置“您对村内便民服务事项的知晓程度”文本
    # 细则点②：文本位于圆环中间
    # 细则点③：文本完整可读
    full_text = '您对村内便民服务事项的知晓程度'
    p1_has_text = False     # ① 图表中放置了该文本
    p2_in_middle = False    # ② 文本位于圆环中间（覆盖在绘图区、且手动定位到中心区域）
    p3_complete = False     # ③ 文本完整可读（完整出现 + 未被设为隐藏/空）
    center_reason = '无图表'
    if chart is not None:
        title = chart.find('.//c:title', NS)
        title_text = ''
        if title is not None:
            for t in title.iter('{%s}t' % NS['a']):
                title_text += (t.text or '')
        title_text_norm = title_text.replace('\n', '').replace(' ', '')

        # ① 中心位置放置该文本：图表标题文本中包含目标文本
        p1_has_text = full_text in title_text_norm

        # ② 位于圆环中间：标题 overlay=1（叠加到绘图区而非占据顶部），
        #    且文本框**实际中心**落在绘图区中部（统一用 title_box_geometry 换算
        #    edge/factor 两种模式，factor 偏移量也能还原出真实中心）。
        overlay = chart.find('.//c:title/c:overlay', NS)
        overlay_on = overlay is not None and overlay.get('val') == '1'

        geo = title_box_geometry(chart)
        centered = False
        center_detail = geo['detail']
        if geo['cx'] is not None and geo['cy'] is not None:
            centered = (0.35 <= geo['cx'] <= 0.65) and (0.35 <= geo['cy'] <= 0.65)
        p2_in_middle = overlay_on and centered

        # ③ 完整可读：目标文本完整出现，标题未被删除/置空(autoTitleDeleted!=1 且有可见文字)，
        #    且若能取到文本框尺寸(w/h)，其不应超出绘图区（>1.0 即明显裁剪）；
        #    无 manualLayout（尺寸未知）时不能仅因缺省布局就判失败，视为"未裁剪"。
        auto_del = chart.find('.//c:autoTitleDeleted', NS)
        title_deleted = auto_del is not None and auto_del.get('val') == '1'
        try:
            w = float(geo['w'] or 0.0)
            h = float(geo['h'] or 0.0)
        except (TypeError, ValueError):
            w = h = 0.0
        not_clipped = not ((w > 1.0) or (h > 1.0))
        p3_complete = p1_has_text and (not title_deleted) and len(title_text_norm) > 0 and not_clipped

        center_reason = ("①含目标文本=%s, ②overlay居中(overlay=%s,%s,中心定位=%s)=%s, ③完整可读(未裁剪=%s)=%s"
                         % (p1_has_text, overlay_on, geo['detail'], centered, p2_in_middle,
                            not_clipped, p3_complete))

    center_text_ok = p1_has_text and p2_in_middle and p3_complete
    hits.append(("plus1_center", 1, center_text_ok,
                 "环状图中心位置放置“您对村内便民服务事项的知晓程度”文本，位于圆环中间且完整可读"
                 " [%s]" % center_reason))

    # ---- +5：环状图整体效果使用图案填充，前景为白色，后景为五种不同的彩色配色 ----
    # 逐条对应细则：①整体使用图案填充 ②前景白色 ③后景五种不同的彩色配色
    pattern_ok, pattern_reason = check_pattern_fill(chart)
    hits.append(("plus5_pattern", 5, pattern_ok,
                 "环状图整体效果使用图案填充，前景为白色，后景为五种不同的彩色配色 [%s]" % pattern_reason))

    # ---- +1：环状图显示百分比数据标签 ----
    # 细则仅要求"显示百分比数据标签"：数据标签整体未被删除，且 showPercent=1 或
    # 数字格式为百分比，真实作用于该系列（即生效，而非被 delete 覆盖）。
    # 不与被评估文件无关的固定 target 百分数比较——那是常量本身的自我比较，恒真，无意义。
    p1_show_pct = False
    label_reason = '无图表'
    if chart is not None:
        dlbls = chart.find('.//c:dLbls', NS)
        dlbls_deleted = False
        if dlbls is not None:
            d = dlbls.find('c:delete', NS)
            dlbls_deleted = d is not None and d.get('val') == '1'

        show_pct = chart.find('.//c:dLbls/c:showPercent', NS)
        show_pct_on = show_pct is not None and show_pct.get('val') == '1'
        numfmt = chart.find('.//c:dLbls/c:numFmt', NS)
        is_pct_fmt = numfmt is not None and '%' in (numfmt.get('formatCode') or '')

        p1_show_pct = (not dlbls_deleted) and (show_pct_on or is_pct_fmt)

        label_reason = ("dLbls未删除=%s, showPercent=%s, 百分比格式=%s → 显示百分比标签=%s"
                        % (not dlbls_deleted, show_pct_on, is_pct_fmt, p1_show_pct))

    pct_label_ok = p1_show_pct
    hits.append(("plus1_pctlabel", 1, pct_label_ok,
                 "环状图显示百分比数据标签"
                 " [%s]" % label_reason))

    # ---- +5：环状图上的百分比数据标签颜色与下方对应的扇形区填充的后景颜色一致 ----
    # 逐条对应细则：①存在百分比数据标签 ②扇形区填充有后景颜色 ③标签颜色与对应扇区后景色一致
    label_color_ok, lc_reason = check_label_color_matches_sector(chart, wb)
    hits.append(("plus5_labelclr", 5, label_color_ok,
                 "环状图上的百分比数据标签颜色与下方对应的扇形区填充的后景颜色一致 [%s]" % lc_reason))

    # ---- +5：环状图中其中四部分连接，另一部分扇形独立分离出环形外 ----
    # 逐条对应细则：①四部分相连 ②另一部分扇形独立分离出环形外
    sep_ok, sep_reason = check_one_slice_separated(chart)
    hits.append(("plus5_separate", 5, sep_ok,
                 "环状图中其中四部分连接，另一部分扇形独立分离出环形外 [%s]" % sep_reason))

    total = sum(score for _id, score, ok, _ in hits if ok)
    return total, hits


def chart_on_right(wb):
    """目标环状图的 drawing anchor 位于 E 列以后，且行范围贴近第一个问题区域。"""
    target_chart_part, _ = get_doughnut_chart(wb)
    if not target_chart_part:
        return False

    # 第一题主体为 B3:B7/D3:D7，对应 0-based 行 2..6。图表应放在该题右侧附近，
    # 允许标题或上边距使锚点略高，但不能用其他远处对象误判为合格。
    first_q_min_row = 0
    first_q_max_row = 6
    sheet_drawing_part = None
    if wb.sheet_root is not None:
        drawing_el = wb.sheet_root.find('.//x:drawing', NS)
        if drawing_el is not None:
            sheet_drawing_part = wb.resolve_rel('xl/worksheets/sheet1.xml', drawing_el.get('{%s}id' % NS['r']))

    for drawing_part, droot in wb.drawing_roots:
        if sheet_drawing_part and drawing_part != sheet_drawing_part:
            continue
        anchors = list(droot.findall('.//xdr:twoCellAnchor', NS)) + list(droot.findall('.//xdr:oneCellAnchor', NS))
        for anchor in anchors:
            chart_el = anchor.find('.//c:chart', NS)
            if chart_el is None:
                continue
            r_id = chart_el.get('{%s}id' % NS['r'])
            if wb.resolve_rel(drawing_part, r_id) != target_chart_part:
                continue

            from_el = anchor.find('xdr:from', NS)
            if from_el is None:
                continue
            col_el = from_el.find('xdr:col', NS)
            row_el = from_el.find('xdr:row', NS)
            try:
                fcol = int(col_el.text) if col_el is not None else -1
                frow = int(row_el.text) if row_el is not None else -1
            except (TypeError, ValueError):
                continue

            if fcol >= 4 and first_q_min_row <= frow <= first_q_max_row:
                return True
    return False


def _theme_color_map(wb):
    """解析 theme1.xml 的 clrScheme，返回 scheme 名称 -> RRGGBB。"""
    root = getattr(wb, 'theme_root', None)
    if root is None:
        return {}
    out = {}
    clr_scheme = root.find('.//a:clrScheme', NS)
    if clr_scheme is None:
        return out
    for item in list(clr_scheme):
        name = item.tag.rsplit('}', 1)[-1]
        srgb = item.find('a:srgbClr', NS)
        sysclr = item.find('a:sysClr', NS)
        val = None
        if srgb is not None:
            val = srgb.get('val')
        elif sysclr is not None:
            val = sysclr.get('lastClr') or sysclr.get('val')
        if val:
            out[name] = val.upper()
    return out


def _resolve_drawing_color(container, theme_colors=None):
    """从 DrawingML 颜色容器中解析 srgbClr/schemeClr/sysClr 为 RRGGBB。"""
    if container is None:
        return None
    srgb = container.find('.//a:srgbClr', NS)
    if srgb is not None and srgb.get('val'):
        return srgb.get('val').upper()
    sysclr = container.find('.//a:sysClr', NS)
    if sysclr is not None:
        val = sysclr.get('lastClr') or sysclr.get('val')
        if val:
            return val.upper()
    scheme = container.find('.//a:schemeClr', NS)
    if scheme is not None and theme_colors is not None:
        val = scheme.get('val')
        if val in theme_colors:
            return theme_colors[val]
    return None


def _fg_bg_of_pattfill(elem, theme_colors=None):
    """从含 pattFill 的元素提取 (fgClr_hex, bgClr_hex)，支持 srgbClr/schemeClr。"""
    patt = elem.find('.//a:pattFill', NS)
    if patt is None:
        return None
    fg = patt.find('a:fgClr', NS)
    bg = patt.find('a:bgClr', NS)
    return (_resolve_drawing_color(fg, theme_colors), _resolve_drawing_color(bg, theme_colors))


def _is_colorful(hexv):
    """判断一个 RRGGBB 颜色是否为“彩色”（非白、非黑、非灰）。"""
    if not hexv or len(hexv) != 6:
        return False
    try:
        r = int(hexv[0:2], 16)
        g = int(hexv[2:4], 16)
        b = int(hexv[4:6], 16)
    except ValueError:
        return False
    # 灰度（含黑白）：R≈G≈B。彩色要求三通道有明显差异。
    return (max(r, g, b) - min(r, g, b)) > 20


def check_pattern_fill(chart):
    """逐条对应细则的每一个点：
       ① 环状图整体效果使用图案填充（5 个扇区均采用 pattFill）
       ② 前景为白色（每个图案填充 fgClr = 白色 FFFFFF）
       ③ 后景为五种不同的彩色配色（5 个扇区 bgClr 互不相同且均为彩色）
    """
    if chart is None:
        return False, '无图表'
    dpts = chart.findall('.//c:ser/c:dPt', NS)

    # ① 整体图案填充：存在 5 个数据点且每个均使用 pattFill
    fgs, bgs = [], []
    all_pattern = len(dpts) >= 5
    if all_pattern:
        for dp in dpts[:5]:
            spPr = dp.find('c:spPr', NS)
            fb = _fg_bg_of_pattfill(spPr) if spPr is not None else None
            if fb is None:
                all_pattern = False
                break
            fgs.append(fb[0])
            bgs.append(fb[1])
    p1_pattern = all_pattern

    # ② 前景为白色：所有图案填充前景色均为白
    p2_fg_white = p1_pattern and all(c == 'FFFFFF' for c in fgs)

    # ③ 后景为五种不同的彩色配色：5 个后景互不相同，且均为彩色（非白/黑/灰）
    p3_bg_five_colors = (
        p1_pattern
        and len([b for b in bgs if b]) == 5
        and len(set(bgs)) == 5
        and all(_is_colorful(b) for b in bgs)
    )

    ok = p1_pattern and p2_fg_white and p3_bg_five_colors
    reason = ("①整体图案填充(5扇区pattFill)=%s, ②前景白色=%s, ③后景五种不同彩色=%s (后景=%s)"
              % (p1_pattern, p2_fg_white, p3_bg_five_colors, bgs if bgs else '无'))
    return ok, reason


def check_label_color_matches_sector(chart, wb):
    """逐条对应细则的每一个点：
       ① 环状图上存在“百分比数据标签”（逐扇区的标签，可读取其颜色）
       ② 每个扇形区填充具有“后景颜色”（pattFill 的 bgClr）
       ③ 每个百分比标签的颜色与“对应”扇区的后景颜色一致
    """
    if chart is None:
        return False, '无图表'
    theme_colors = _theme_color_map(wb)
    ser = chart.find('.//c:ser', NS)
    dpts = chart.findall('.//c:ser/c:dPt', NS)

    if ser is None:
        return False, '无系列'

    # ① 先确认这些数据标签确实是百分比标签，且 dLbls 没有被删除。
    dLbls = ser.find('c:dLbls', NS)
    if dLbls is None:
        dLbls = chart.find('.//c:dLbls', NS)
    deleted = False
    if dLbls is not None:
        del_el = dLbls.find('c:delete', NS)
        deleted = del_el is not None and del_el.get('val') == '1'
    show_pct = dLbls.find('c:showPercent', NS) if dLbls is not None else None
    show_pct_on = show_pct is not None and show_pct.get('val') == '1'
    numfmt = dLbls.find('c:numFmt', NS) if dLbls is not None else None
    is_pct_fmt = numfmt is not None and '%' in (numfmt.get('formatCode') or '')
    p1_has_pct_labels = (dLbls is not None) and (not deleted) and (show_pct_on or is_pct_fmt)

    # ② 标签文字颜色：优先逐点 dLbl/txPr，其次继承系列级/全局 dLbls/txPr。
    inherited_txpr = dLbls.find('c:txPr', NS) if dLbls is not None else None
    label_colors = {}
    point_dlbls = {}
    if dLbls is not None:
        for dl in dLbls.findall('c:dLbl', NS):
            idx_el = dl.find('c:idx', NS)
            if idx_el is not None:
                try:
                    point_dlbls[int(idx_el.get('val'))] = dl
                except (TypeError, ValueError):
                    pass

    for idx in range(5):
        txpr = None
        dl = point_dlbls.get(idx)
        if dl is not None:
            txpr = dl.find('c:txPr', NS)
        if txpr is None:
            txpr = inherited_txpr
        label_colors[idx] = _resolve_drawing_color(txpr, theme_colors) if txpr is not None else None
    p2_has_label_colors = len([c for c in label_colors.values() if c]) >= 5

    # ③ 各扇区填充的“后景颜色”：取每个 dPt 图案填充的 bgClr，支持 srgbClr/schemeClr。
    sector_bg = {}
    for dp in dpts:
        idx_el = dp.find('c:idx', NS)
        spPr = dp.find('c:spPr', NS)
        if idx_el is None or spPr is None:
            continue
        try:
            idx = int(idx_el.get('val'))
        except (TypeError, ValueError):
            continue
        fb = _fg_bg_of_pattfill(spPr, theme_colors)
        sector_bg[idx] = fb[1] if fb else None     # 细则明确指“后景颜色”=bgClr
    p3_has_bg = (len([b for b in sector_bg.values() if b]) >= 5)

    matched = 0
    for idx, bg in sector_bg.items():
        if bg is not None and label_colors.get(idx) == bg:
            matched += 1
    p4_all_match = p1_has_pct_labels and p2_has_label_colors and p3_has_bg and (matched >= 5)

    ok = p1_has_pct_labels and p2_has_label_colors and p3_has_bg and p4_all_match
    reason = ("①百分比标签(showPercent=%s,百分比格式=%s,未删除=%s)=%s, "
              "②标签色数=%d/5=%s, ③扇区后景色数=%d/5=%s, "
              "④标签色继承后与对应扇区后景色一致数=%d/5=%s"
              % (show_pct_on, is_pct_fmt, not deleted, p1_has_pct_labels,
                 len([c for c in label_colors.values() if c]), p2_has_label_colors,
                 len([b for b in sector_bg.values() if b]), p3_has_bg,
                 matched, p4_all_match))
    return ok, reason


def check_one_slice_separated(chart):
    """逐条对应细则的每一个点：
       ① 环状图共有 5 部分（5 个扇区），其中四部分相连（即 4 个扇区未分离、保持相连）
       ② 另一部分扇形独立分离出环形外（恰好 1 个扇区设置了 explosion，明显移出环外）
    """
    if chart is None:
        return False, '无图表'
    # 逐扇区 explosion 值（数据点级）
    dpt_expl = []
    for dp in chart.findall('.//c:ser/c:dPt', NS):
        e = dp.find('c:explosion', NS)
        dpt_expl.append(int(e.get('val')) if e is not None else 0)

    # 共有 5 个扇区
    five_sectors = len(dpt_expl) == 5

    # 分离的扇区：explosion 明显（>=10）即视为移出环形外
    separated = [v for v in dpt_expl if v >= 10]
    connected = [v for v in dpt_expl if v < 10]   # 未分离（相连）的扇区

    # ① 四部分相连：恰好 4 个扇区未分离
    p1_four_connected = five_sectors and len(connected) == 4
    # ② 另一部分独立分离出环外：恰好 1 个扇区分离
    p2_one_separated = five_sectors and len(separated) == 1

    ok = p1_four_connected and p2_one_separated
    if ok:
        reason = ("①四部分相连(未分离扇区数=4)=%s, ②另一扇区独立分离出环外(explosion=%d)=%s"
                  % (p1_four_connected, separated[0], p2_one_separated))
        return True, reason

    ser_e = chart.find('.//c:ser/c:explosion', NS)
    ser_v = int(ser_e.get('val')) if ser_e is not None else 0
    reason = ("①四部分相连(未分离扇区数=%d,需=4)=%s, ②恰一扇区分离出环外(分离数=%d,需=1)=%s "
              "[逐点explosion=%s, 系列级explosion=%d]"
              % (len(connected) if five_sectors else 0, p1_four_connected,
                 len(separated) if five_sectors else 0, p2_one_separated,
                 dpt_expl, ser_v))
    return False, reason


# ---------------------------------------------------------------------------
# 统一入口
# ---------------------------------------------------------------------------
# 维度二评分项满分表：id 为稳定标识，与 check_dimension2 内部按同一 id 写入的结果一一对应，
# 不依赖顺序或文本前缀匹配（同分值的多个扣分项也能正确区分，如两个 -3 项）。
_DIM2_RULES = [
    ("plus3_ring",     3,  ("Sheet1第一个问题右侧空白区新增1个环状图，数据源来自B3:B7/D3:D7，含5个回答选项，"
                            "各扇区占比与D3:D7基本一致，圆环样式有明显中空，非饼/柱/折线/散点图")),
    ("plus1_center",   1,  "环状图中心位置放置“您对村内便民服务事项的知晓程度”文本，位于圆环中间且完整可读"),
    ("plus5_pattern",  5,  "环状图整体效果使用图案填充，前景为白色，后景为五种不同的彩色配色"),
    ("plus1_pctlabel", 1,  "环状图显示百分比数据标签"),
    ("plus5_labelclr", 5,  "环状图上的百分比数据标签颜色与下方对应的扇形区填充的后景颜色一致"),
    ("plus5_separate", 5,  "环状图中其中四部分连接，另一部分扇形独立分离出环形外"),
]


def _locate_target_file(dir_path: str):
    """在脚本所在目录里定位被评估的 xlsx/xlsm 文件。

    优先匹配含关键字“乡村公共服务调查汇总”的文件；否则退化为目录内首个 .xlsx/.xlsm。
    忽略 Excel 临时锁文件（以 ~$ 开头）。
    """
    if not os.path.isdir(dir_path):
        return None
    candidates = []
    for name in sorted(os.listdir(dir_path)):
        if name.startswith('~$'):
            continue
        low = name.lower()
        if low.endswith('.xlsx') or low.endswith('.xlsm'):
            candidates.append(name)
    if not candidates:
        return None
    for name in candidates:
        if '乡村公共服务调查汇总' in name:
            return os.path.join(dir_path, name)
    return os.path.join(dir_path, candidates[0])


def evaluate(dir_path: str) -> dict:
    """统一入口：接收脚本所在目录的路径，脚本自行定位并评估该目录内的 xlsx 文档。

    返回结构见 §2.2 约定。
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
        "max_score": sum(md for _, md, _ in _DIM2_RULES if md > 0),
    }

    try:
        path = _locate_target_file(dir_path)
        if not path or not os.path.exists(path):
            result["status"] = "error"
            result["error"] = "目录内未找到可评估的 .xlsx/.xlsm 文件：%s" % dir_path
            return result

        result["file_name"] = os.path.basename(path)

        wb = WorkbookData(path)
        cells = build_cell_map(wb.sheet_root)

        # ---------- 维度1 ----------
        d1_pass, d1_details = check_dimension1(wb)
        result["dim1_pass"] = bool(d1_pass)
        if not d1_pass:
            fails = [msg for ok, msg in d1_details if not ok]
            result["dim1_reason"] = "；".join(fails)
            result["dim2_items"] = []
            result["total_score"] = 0
            return result

        # ---------- 维度2 ----------
        total, hits = check_dimension2(wb, cells)

        # 将 hits 按稳定规则 ID 对齐；未返回的评分项视为未命中。
        hit_by_id = {
            item_id: (score, ok, msg)
            for item_id, score, ok, msg in hits
        }
        items = []
        for rule_id, max_delta, rule in _DIM2_RULES:
            hit = hit_by_id.get(rule_id)
            if hit is not None:
                score, ok, msg = hit
                items.append({
                    "rule": msg.split(' [')[0],
                    "max_delta": max_delta,
                    "delta": score if ok else 0,
                    "hit": bool(ok),
                    "detail": "",
                })
            else:
                items.append({
                    "rule": rule,
                    "max_delta": max_delta,
                    "delta": 0,
                    "hit": False,
                    "detail": "",
                })

        result["dim2_items"] = items
        result["total_score"] = int(total)
        return result

    except Exception as exc:
        result["status"] = "error"
        result["error"] = "%s: %s" % (type(exc).__name__, exc)
        if not result["dim2_items"]:
            result["dim2_items"] = []
        return result


if __name__ == '__main__':
    # 本地调试用：默认以脚本所在目录作为入参
    _dir = sys.argv[1] if len(sys.argv) >= 2 else os.path.dirname(os.path.abspath(__file__))
    print(json.dumps(evaluate(_dir), ensure_ascii=False, indent=2))
