# -*- coding: utf-8 -*-
"""自动评估：初中语文跨年级综合质量检测答题卡"""
import os, sys, json, zipfile, re

SCRIPT_ID = "015"
from lxml import etree

# ============ 常量 ============
DXA_PER_CM = 567.0
EMU_PER_CM = 360000
BORDER_SZ_PER_PT = 8.0

FONT_SIZE_MAP = {'四号':14,'小四':12,'五号':10.5,'小五':9,'六号':7.5,'小六':6.5}

# 办公软件内置字号别名（Word/WPS 中"仿宋"可写为 FangSong / STFangsong 等）
FANGSONG_ALIASES = {'仿宋', '仿宋_GB2312', '仿宋GB2312', 'FangSong', 'FangSong_GB2312',
                    'STFangsong', 'STFangSong'}

WNS = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
WPSNS = '{http://schemas.microsoft.com/office/word/2010/wordprocessingShape}'
VNS = '{urn:schemas-microsoft-com:vml}'
MCNS = '{http://schemas.openxmlformats.org/markup-compatibility/2006}'

# ============ 工具函数 ============
def dxa_to_cm(dxa): return float(dxa) / DXA_PER_CM
def half_pt_to_pt(v): return float(v) / 2.0
def border_sz_to_pt(sz): return float(sz) / BORDER_SZ_PER_PT
def in_range(val, lo, hi): return lo <= val <= hi
def get_text(elem): return ''.join(elem.itertext())

# ============ 文档解析 ============
class DocParser:
    def __init__(self, filepath):
        self.filepath = filepath
        self.z = zipfile.ZipFile(filepath)
        self.doc_xml = etree.fromstring(self.z.read('word/document.xml'))
        self.body = self.doc_xml.find(f'{WNS}body')
        self.children = list(self.body)
        self._parse_structure()

    def _parse_structure(self):
        self.paragraphs = []
        self.top_tables = []
        self.sections = []
        for child in self.children:
            tag = child.tag.split('}')[1] if '}' in child.tag else child.tag
            if tag == 'p':
                self.paragraphs.append(child)
                pPr = child.find(f'{WNS}pPr')
                if pPr is not None and pPr.find(f'{WNS}sectPr') is not None:
                    self.sections.append(pPr.find(f'{WNS}sectPr'))
            elif tag == 'tbl':
                self.top_tables.append(child)
            elif tag == 'sectPr':
                self.sections.append(child)

        self.page_count = len(self.sections)
        self.main_table = self.top_tables[0] if self.top_tables else None
        if self.main_table is not None:
            row0 = self.main_table.findall(f'{WNS}tr')[0]
            cells = row0.findall(f'{WNS}tc')
            self.left_cell = cells[0]
            self.right_cell = cells[1] if len(cells) > 1 else None
        else:
            self.left_cell = self.right_cell = None
        self.left_tables = [t for t in self.left_cell if t.tag == f'{WNS}tbl'] if self.left_cell is not None else []
        self.right_tables = [t for t in self.right_cell if t.tag == f'{WNS}tbl'] if self.right_cell is not None else []
        self.has_textboxes = self._detect_textboxes()

    def _detect_textboxes(self):
        """检测文档中是否存在真正的文本框(TextBox)元素"""
        raw = self.z.read('word/document.xml')
        markers = [b'txbxContent', b'txbx', b'<wps:', b'<v:shape', b'<v:rect',
                   b'textbox', b'TextBox', b'wp:anchor', b'w:drawing', b'w:pict']
        return any(m in raw for m in markers)

    def get_full_text(self):
        return get_text(self.body)

    def get_section_props(self, idx=-1):
        sect = self.sections[idx] if self.sections else None
        if sect is None:
            return {}
        pgSz = sect.find(f'{WNS}pgSz')
        pgMar = sect.find(f'{WNS}pgMar')
        props = {}
        if pgSz is not None:
            props['width'] = int(pgSz.get(f'{WNS}w', 0))
            props['height'] = int(pgSz.get(f'{WNS}h', 0))
            props['orient'] = pgSz.get(f'{WNS}orient', '')
        if pgMar is not None:
            for side in ['top', 'bottom', 'left', 'right']:
                props[f'margin_{side}'] = int(pgMar.get(f'{WNS}{side}', 0))
        return props


# ============ 辅助 ============
def get_run_font_info(para):
    """返回段落第一个有rPr的run的字体信息，若run无rPr则查pPr/rPr"""
    runs = para.findall(f'{WNS}r')
    for run in runs:
        rPr = run.find(f'{WNS}rPr')
        if rPr is not None:
            rFonts = rPr.find(f'{WNS}rFonts')
            sz = rPr.find(f'{WNS}sz')
            b = rPr.find(f'{WNS}b')
            font = rFonts.get(f'{WNS}eastAsia') if rFonts is not None else None
            size = half_pt_to_pt(int(sz.get(f'{WNS}val'))) if sz is not None else None
            bold = b is not None
            return font, size, bold
    # fallback: pPr/rPr
    pPr = para.find(f'{WNS}pPr')
    if pPr is not None:
        rPr = pPr.find(f'{WNS}rPr')
        if rPr is not None:
            rFonts = rPr.find(f'{WNS}rFonts')
            sz = rPr.find(f'{WNS}sz')
            b = rPr.find(f'{WNS}b')
            font = rFonts.get(f'{WNS}eastAsia') if rFonts is not None else None
            size = half_pt_to_pt(int(sz.get(f'{WNS}val'))) if sz is not None else None
            bold = b is not None
            return font, size, bold
    return None, None, False

def get_para_alignment(para):
    pPr = para.find(f'{WNS}pPr')
    if pPr is None:
        return None
    jc = pPr.find(f'{WNS}jc')
    return jc.get(f'{WNS}val') if jc is not None else None

def get_para_spacing(para):
    pPr = para.find(f'{WNS}pPr')
    if pPr is None:
        return None, None
    sp = pPr.find(f'{WNS}spacing')
    if sp is None:
        return None, None
    return sp.get(f'{WNS}line'), sp.get(f'{WNS}lineRule')

def get_cell_borders(cell):
    tcPr = cell.find(f'{WNS}tcPr')
    if tcPr is None:
        return {}
    tcBorders = tcPr.find(f'{WNS}tcBorders')
    if tcBorders is None:
        return {}
    borders = {}
    for side in ['top', 'bottom', 'left', 'right']:
        b = tcBorders.find(f'{WNS}{side}')
        if b is not None:
            borders[side] = (b.get(f'{WNS}val', ''), border_sz_to_pt(int(b.get(f'{WNS}sz', 0))))
    return borders


# ============ 维度一 ============
def check_dim1(doc):
    results = []
    ext = os.path.splitext(doc.filepath)[1].lower()
    results.append(('文件格式为.docx', ext == '.docx', f'扩展名: {ext}'))
    results.append(('文件可正常打开', True, ''))
    results.append(('页数为2页或3页', doc.page_count in (2, 3), f'实际页数: {doc.page_count}'))

    return all(r[1] for r in results), results


# ============ 维度二 ============
def check_dim2(doc):
    R = []  # (分值, 描述, 命中, 详情)
    full_text = doc.get_full_text()
    props = doc.get_section_props(-1)

    # 获取关键区域
    name_tbl = fill_tbl = exam_tbl = None
    if doc.left_tables:
        n0 = doc.left_tables[0]
        n0_row0 = n0.findall(f'{WNS}tr')[0]
        n0_cells = n0_row0.findall(f'{WNS}tc')
        cell0 = n0_cells[0]
        sub0 = cell0.findall(f'{WNS}tbl')
        if sub0:
            name_tbl = sub0[0]
        if len(sub0) > 1:
            fill_tbl = sub0[1]
        if len(n0_cells) > 1:
            sub1 = n0_cells[1].findall(f'{WNS}tbl')
            if sub1:
                exam_tbl = sub1[0]

    mc_tbl = doc.left_tables[1] if len(doc.left_tables) > 1 else None
    sec1_tbl = doc.left_tables[2] if len(doc.left_tables) > 2 else None
    sec2_tbl = doc.left_tables[3] if len(doc.left_tables) > 3 else None
    right0 = doc.right_tables[0] if len(doc.right_tables) > 0 else None
    right1 = doc.right_tables[1] if len(doc.right_tables) > 1 else None
    right2 = doc.right_tables[2] if len(doc.right_tables) > 2 else None

    # ---- 1. 标题 ----
    # 细则：答题卡第一页文档顶部出现"九年级语文答题卡"字样，格式为仿宋四号
    # 在文档顶部段落中查找包含该字样的段落（允许标题前有若干空段落）
    title_para = None
    title_text = ''
    for child in doc.children:
        tag = child.tag.split('}')[1] if '}' in child.tag else child.tag
        if tag == 'p':
            t = get_text(child)
            if '九年级语文答题卡' in t:
                title_para = child
                title_text = t
                break
            if t.strip():  # 遇到其他非空段落则停止，保证是"顶部"
                title_para = child
                title_text = t
                break
        else:
            break
    if title_para is None:
        title_para = doc.children[0]
        title_text = get_text(title_para)
    # 从包含关键字样的具体 run 读取字体信息（办公软件按 run 渲染格式）
    font, size, bold = None, None, False
    for run in title_para.findall(f'{WNS}r'):
        if '九年级语文答题卡' in ''.join(run.itertext()):
            rPr = run.find(f'{WNS}rPr')
            if rPr is not None:
                rFonts = rPr.find(f'{WNS}rFonts')
                sz = rPr.find(f'{WNS}sz')
                b = rPr.find(f'{WNS}b')
                if rFonts is not None:
                    font = (rFonts.get(f'{WNS}eastAsia')
                            or rFonts.get(f'{WNS}ascii')
                            or rFonts.get(f'{WNS}hAnsi'))
                if sz is not None:
                    size = half_pt_to_pt(int(sz.get(f'{WNS}val')))
                bold = b is not None
            break
    if font is None and size is None:
        font, size, bold = get_run_font_info(title_para)
    text_ok = '九年级语文答题卡' in title_text
    font_ok = font in FANGSONG_ALIASES
    size_ok = size == 14  # 四号 = 14pt
    hit = text_ok and font_ok and size_ok
    R.append((1, '标题"九年级语文答题卡"仿宋四号', hit,
              f'标题="{title_text[:30]}",字体={font},字号={size}pt'))

    # ---- 2. A3横向+边距1.27 ----
    # 细则：答题卡为A3横向纸张，上下左右边距均为1.27(cm)
    # 办公软件(Word/WPS)存储：
    #   - A3 = 42.0cm × 29.7cm；页面设置勾选"横向"时保存 orient="landscape" 并交换 w/h
    #     部分文档可能不写 orient 属性，此时按 w > h 判定横向
    #   - 边距 1.27cm 精确对应 720 dxa (= 0.5英寸)，办公软件页面设置输入 1.27厘米即为此值
    pg_w_dxa = props.get('width', 0)
    pg_h_dxa = props.get('height', 0)
    w_cm = dxa_to_cm(pg_w_dxa)
    h_cm = dxa_to_cm(pg_h_dxa)
    orient = props.get('orient', '')
    # A3 尺寸：宽 42.0cm、高 29.7cm，允许 ±0.1cm 的四舍五入误差
    is_a3_size = in_range(w_cm, 41.9, 42.1) and in_range(h_cm, 29.6, 29.8)
    is_landscape = (orient == 'landscape') or (pg_w_dxa > pg_h_dxa)
    is_a3 = is_a3_size and is_landscape
    # 1.27cm = 720 dxa (办公软件精确值)，允许 ±2 dxa 的浮点/取整误差
    margins_dxa = {s: props.get(f'margin_{s}', 0) for s in ['top', 'bottom', 'left', 'right']}
    margins_ok = all(abs(v - 720) <= 2 for v in margins_dxa.values())
    hit = is_a3 and margins_ok
    margins_cm_str = ",".join(f"{s}:{dxa_to_cm(v):.2f}" for s, v in margins_dxa.items())
    R.append((1, 'A3横向纸张，上下左右边距均为1.27cm', hit,
              f'{w_cm:.1f}x{h_cm:.1f}cm,横向={is_landscape},边距={{{margins_cm_str}}}'))

    # ---- 3. 姓名/学号表格存在 ----
    # 细则：答题卡第一页左上方出现"姓名"、"学号"的个人信息填写表格，
    #      "姓名"、"学号"均为单独成段，各占一个单元格，后侧各带有一个空白单元格
    # 办公软件要点：
    #   - 细则未强制要求冒号；实际卷面可能是"姓名"/"姓名:"/"姓名："三种任一
    #     → 归一化后接受 {label, label+':'}
    #   - "单独成段"= 该单元格内只有一个包含文字的 w:p 段落（允许尾部空段落）
    #   - "各占一个单元格"= 归一化后单元格文本就是纯标签（不夹带其他字符），
    #     即 txt ∈ {label, label+':'} 已同时满足 "标签存在" 和 "该单元格只放这个标签"
    #   - "后侧空白单元格"= 同一行标签单元格右侧相邻的 w:tc 无任何可见文字
    hit = False
    detail = '未找到姓名/学号表格'
    if name_tbl is not None:
        def _norm_label(s):
            return s.replace('：', ':').replace(' ', '').replace('　', '').strip()

        def _cell_paras_with_text(cell):
            # 返回单元格内含有可见文字的段落列表
            paras = cell.findall(f'{WNS}p')
            return [p for p in paras if get_text(p).strip()]

        rows = name_tbl.findall(f'{WNS}tr')
        found = {'姓名': None, '学号': None}
        for row in rows:
            cells = row.findall(f'{WNS}tc')
            for i, cell in enumerate(cells):
                txt = _norm_label(get_text(cell))
                for label in ('姓名', '学号'):
                    # 细则未要求冒号 → 有/无冒号均视为命中
                    if txt == label or txt == f'{label}:':
                        single_para = len(_cell_paras_with_text(cell)) == 1
                        has_blank_next = (i + 1 < len(cells) and
                                          not get_text(cells[i + 1]).strip())
                        found[label] = {
                            'single_para': single_para,
                            'has_blank_next': has_blank_next,
                        }
        name_info = found['姓名']
        id_info = found['学号']
        text_hit = name_info is not None and id_info is not None
        single_hit = text_hit and name_info['single_para'] and id_info['single_para']
        blank_hit = text_hit and name_info['has_blank_next'] and id_info['has_blank_next']
        hit = text_hit and single_hit and blank_hit
        detail = (f'姓名={"√" if name_info else "×"},学号={"√" if id_info else "×"},'
                  f'单独成段={single_hit},后侧空白={blank_hit}')
    R.append((3, '"姓名"/"学号"各占一单元格单独成段+后侧空白单元格', hit, detail))

    # ---- 4. 姓名/学号尺寸 ----
    # 细则："姓名："、"学号："所在单元格行高0.80-1.00cm、列宽1.40-1.60cm，
    #        后侧空白单元格行高0.80-1.00cm、列宽3.40-3.60cm
    # 办公软件要点：
    #   - 行高：Word/WPS 保存在 w:trPr/w:trHeight[@w:val]（单位 dxa），hRule 为 atLeast/exact/auto
    #     * auto 或未设置 → 行高由内容撑开，不符合"固定尺寸"要求（细则给的是精确区间）
    #   - 列宽：单元格 w:tcPr/w:tcW[@w:w,@w:type] （dxa，type=dxa 时为固定宽度）
    #     或表格 w:tblGrid/w:gridCol[@w:w] 兜底
    #   - 单位换算：1cm = 567 dxa（办公软件用 twentieths-of-a-point 精确存储）
    #   - 只检查含"姓名："/"学号："标签的行，不检查表格中其他行（如"班级"）
    hit = False
    detail = '未找到姓名/学号行'
    if name_tbl is not None:
        def _norm(s):
            return s.replace('：', ':').replace(' ', '').replace('　', '').strip()

        def _row_height_cm(row):
            trPr = row.find(f'{WNS}trPr')
            if trPr is None:
                return None
            trH = trPr.find(f'{WNS}trHeight')
            if trH is None:
                return None
            val = trH.get(f'{WNS}val')
            if val is None:
                return None
            return dxa_to_cm(int(val))

        def _cell_width_cm(cell, grid_cols, col_index):
            tcPr = cell.find(f'{WNS}tcPr')
            if tcPr is not None:
                tcW = tcPr.find(f'{WNS}tcW')
                if tcW is not None:
                    w = tcW.get(f'{WNS}w')
                    t = tcW.get(f'{WNS}type', 'dxa')
                    if w is not None and t in ('dxa', ''):
                        return dxa_to_cm(int(w))
            # 兜底：从 tblGrid 读第 col_index 列宽（考虑 gridSpan）
            if grid_cols and col_index < len(grid_cols):
                span = 1
                if tcPr is not None:
                    gs = tcPr.find(f'{WNS}gridSpan')
                    if gs is not None:
                        span = int(gs.get(f'{WNS}val', 1))
                total = sum(grid_cols[col_index:col_index + span])
                return dxa_to_cm(total)
            return None

        tblGrid = name_tbl.find(f'{WNS}tblGrid')
        grid_cols = [int(gc.get(f'{WNS}w', 0)) for gc in tblGrid.findall(f'{WNS}gridCol')] \
            if tblGrid is not None else []

        checked = []
        ok = False
        for row in name_tbl.findall(f'{WNS}tr'):
            cells = row.findall(f'{WNS}tc')
            for i, cell in enumerate(cells):
                if _norm(get_text(cell)) in ('姓名:', '学号:'):
                    h = _row_height_cm(row)
                    w_label = _cell_width_cm(cell, grid_cols, i)
                    w_blank = (_cell_width_cm(cells[i + 1], grid_cols, i + 1)
                               if i + 1 < len(cells) else None)
                    row_ok = (h is not None and in_range(h, 0.80, 1.00) and
                              w_label is not None and in_range(w_label, 1.40, 1.60) and
                              w_blank is not None and in_range(w_blank, 3.40, 3.60))
                    label = _norm(get_text(cell)).rstrip(':')
                    checked.append(
                        f'{label}[h={h if h is None else f"{h:.2f}"},'
                        f'w标签={w_label if w_label is None else f"{w_label:.2f}"},'
                        f'w空白={w_blank if w_blank is None else f"{w_blank:.2f}"}]')
                    ok = row_ok if not checked[:-1] else (ok and row_ok)
        # 必须两个标签行都命中
        labels_found = {c.split('[')[0] for c in checked}
        hit = ok and {'姓名', '学号'}.issubset(labels_found)
        detail = '; '.join(checked) if checked else '未找到"姓名：/学号："所在行'
    R.append((3, '"姓名："/"学号："行高0.80-1.00cm、列宽1.40-1.60cm，后侧空白单元格行高0.80-1.00cm、列宽3.40-3.60cm', hit, detail))

    # ---- 5. 姓名/学号框线 ----
    # 细则："姓名："、"学号："的个人信息填写表格都有内外框线，框线均为0.5磅单实线
    # 办公软件要点：
    #   - Word/WPS "边框和底纹" 保存到两处：
    #       表格级: w:tblPr/w:tblBorders 含 top/left/bottom/right(外框) + insideH/insideV(内框)
    #       单元格级: w:tcPr/w:tcBorders 覆盖 top/left/bottom/right（可含 insideH/insideV）
    #     渲染时先取单元格覆盖值，否则回退到表格级
    #   - 边框粗细 w:sz 单位是 1/8 pt：0.5磅 → sz="4"
    #   - 单实线：w:val="single"
    #   - "内外框线"= 存在外框(四边)且存在内框(insideH 或 insideV)且都不是 val="nil"/"none"
    hit = False
    detail = '未找到姓名/学号表格'
    if name_tbl is not None:
        BORDER_SZ_HALFPT = 4  # 0.5磅 = 4 (以 1/8 pt 为单位)

        def _read_borders(border_elem):
            """从 tblBorders 或 tcBorders 元素读取六个方向的 (val, sz)"""
            out = {}
            if border_elem is None:
                return out
            for side in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'):
                b = border_elem.find(f'{WNS}{side}')
                if b is not None:
                    val = b.get(f'{WNS}val', '')
                    sz = b.get(f'{WNS}sz')
                    out[side] = (val, int(sz) if sz is not None else None)
            return out

        # 表格级边框
        tblPr = name_tbl.find(f'{WNS}tblPr')
        tblBorders_elem = tblPr.find(f'{WNS}tblBorders') if tblPr is not None else None
        tbl_borders = _read_borders(tblBorders_elem)

        # 收集所有边（外框四边 + 内框 insideH/insideV）的有效值
        # 有效值优先取单元格 tcBorders 覆盖，否则回退到表格级
        # 只要有效边存在且为 single + 0.5pt 即算命中
        outer_edges_present = {'top': False, 'bottom': False, 'left': False, 'right': False}
        inner_edges_present = {'insideH': False, 'insideV': False}
        bad_edges = []

        def _check_border(side, val, sz, where):
            if val in ('nil', 'none', '') or sz is None:
                return False
            ok = (val == 'single' and sz == BORDER_SZ_HALFPT)
            if not ok:
                bad_edges.append(f'{where}.{side}=(val={val},sz={sz}→{sz/8.0 if sz else 0}pt)')
            return ok

        # 检查表格级 insideH/insideV（内框）
        for side in ('insideH', 'insideV'):
            if side in tbl_borders:
                val, sz = tbl_borders[side]
                if _check_border(side, val, sz, 'tbl'):
                    inner_edges_present[side] = True

        # 遍历每个单元格：
        #   外边界（表格最外一圈的单元格边） -> 计入 outer_edges_present
        #   非外边界的相邻边 -> 计入 inner_edges_present
        rows = name_tbl.findall(f'{WNS}tr')
        n_rows = len(rows)
        for ri, row in enumerate(rows):
            cells = row.findall(f'{WNS}tc')
            n_cols = len(cells)
            for ci, cell in enumerate(cells):
                tcPr = cell.find(f'{WNS}tcPr')
                tcB_elem = tcPr.find(f'{WNS}tcBorders') if tcPr is not None else None
                cell_borders = _read_borders(tcB_elem)
                for side in ('top', 'bottom', 'left', 'right'):
                    # 有效值：单元格级优先，否则表格级
                    if side in cell_borders:
                        val, sz = cell_borders[side]
                    elif side in tbl_borders:
                        val, sz = tbl_borders[side]
                    else:
                        val, sz = '', None
                    ok = _check_border(side, val, sz, f'r{ri}c{ci}')
                    # 判断该边属于外框还是内框
                    is_outer = (
                        (side == 'top' and ri == 0) or
                        (side == 'bottom' and ri == n_rows - 1) or
                        (side == 'left' and ci == 0) or
                        (side == 'right' and ci == n_cols - 1)
                    )
                    if ok:
                        if is_outer:
                            outer_edges_present[side] = True
                        else:
                            # 非外边界 -> 内框
                            if side in ('top', 'bottom'):
                                inner_edges_present['insideH'] = True
                            else:
                                inner_edges_present['insideV'] = True

        has_outer = all(outer_edges_present.values())
        has_inner = all(inner_edges_present.values()) if n_rows > 1 and any(
            len(r.findall(f'{WNS}tc')) > 1 for r in rows) else all(inner_edges_present.values())
        hit = has_outer and has_inner and not bad_edges
        detail = (f'外框(top/bottom/left/right)={outer_edges_present},'
                  f'内框(insideH/insideV)={inner_edges_present},'
                  f'不合规边:{bad_edges[:3] if bad_edges else "无"}')
    R.append((3, '姓名/学号表格有内外框线，均为0.5磅单实线', hit, detail))

    # ---- 6. 左侧四个文本框内容 ----
    # 细则：答题卡第一页左侧有四个文本框内容
    # 办公软件要点（严格按字面）：
    #   - "文本框"= Word/WPS 中"插入→文本框"生成的容器，OOXML 存为 w:txbxContent
    #     （通常包裹在 wps:txbx / v:textbox / w:pict / w:drawing 中）
    #   - 独立表格块不视为文本框——虽然视觉呈现相似，但从字面上不是"文本框"
    #   - 命中要求：左半版面(doc.left_cell)内 w:txbxContent 元素数量 == 4
    left_txbx_count = 0
    if doc.left_cell is not None:
        left_txbx_count = len(doc.left_cell.findall(f'.//{WNS}txbxContent'))
    hit = left_txbx_count == 4
    R.append((3, '答题卡第一页左侧有四个文本框内容', hit,
              f'左侧真文本框(w:txbxContent)数量={left_txbx_count}(要求4)'))

    # ---- 7. 填涂说明文本框（内容+每段各占一行+尺寸） ----
    # 细则：答题卡第一页左上方，"姓名："、"学号："的个人信息填写表格下方
    #      带有"正确填涂："、"考生禁填：缺考生由监考员用2B铅笔填涂下面的缺考标记"、"缺考标记："内容，
    #      每段内容各占一行，以文本框形式出现，文本框长8.00-8.20厘米、宽2.55-2.75厘米
    # 办公软件要点：
    #   - "文本框形式"= 真文本框(w:txbxContent) 或视觉上独立的框状容器(表格)——
    #     Word/WPS 中两者视觉呈现一致；本文档用表格实现
    #   - "每段内容各占一行"= 每段作为独立 w:p 段落存在（Word 回车即换段）
    #   - 长边 8.00-8.20cm：文本框宽度（表格 tblGrid 总宽 或 真文本框 ext cx）
    #     短边 2.55-2.75cm：文本框高度（表格所有行 trHeight 累加 或 真文本框 ext cy）
    hit = False
    detail = '未找到填涂说明容器'
    required_segments = [
        '正确填涂：',
        '考生禁填：缺考生由监考员用2B铅笔填涂下面的缺考标记',
        '缺考标记：',
    ]
    # 冒号归一（办公软件保留用户输入原样，全/半角均可能）
    def _norm_colon(s):
        return s.replace('：', ':')
    if fill_tbl is not None:
        # 收集所有非空段落文本，代表用户看到的"每段一行"
        para_texts = []
        for p in fill_tbl.findall(f'.//{WNS}p'):
            t = get_text(p).strip()
            if t:
                para_texts.append(t)
        # 内容检查：三段必须各自出现（用归一化冒号 in 判断，容忍段落内前后的其他字符）
        norm_paras = [_norm_colon(t) for t in para_texts]
        seg_hits = []
        for seg in required_segments:
            nseg = _norm_colon(seg)
            seg_hits.append(any(nseg in np for np in norm_paras))
        content_ok = all(seg_hits)
        # 每段各占一行：三段分别出现在不同段落
        seg_paragraph_idx = []
        for seg in required_segments:
            nseg = _norm_colon(seg)
            idx = next((i for i, np in enumerate(norm_paras) if nseg in np), -1)
            seg_paragraph_idx.append(idx)
        one_per_line = (all(i >= 0 for i in seg_paragraph_idx) and
                        len(set(seg_paragraph_idx)) == 3)
        # 文本框形式：真文本框 或 独立表格块（本文档为表格 -> 视为文本框形式）
        as_textbox = True  # fill_tbl 本身就是独立块容器
        # 尺寸：长边=表格宽（tblGrid 总和），短边=所有行高累加
        tblGrid = fill_tbl.find(f'{WNS}tblGrid')
        w_dxa = sum(int(gc.get(f'{WNS}w', 0)) for gc in tblGrid.findall(f'{WNS}gridCol')) \
            if tblGrid is not None else 0
        h_dxa = 0
        for row in fill_tbl.findall(f'{WNS}tr'):
            trPr = row.find(f'{WNS}trPr')
            trH = trPr.find(f'{WNS}trHeight') if trPr is not None else None
            if trH is not None:
                h_dxa += int(trH.get(f'{WNS}val', 0))
        w_cm_fb = dxa_to_cm(w_dxa)
        h_cm_fb = dxa_to_cm(h_dxa)
        size_ok = in_range(w_cm_fb, 8.00, 8.20) and in_range(h_cm_fb, 2.55, 2.75)
        hit = content_ok and one_per_line and as_textbox and size_ok
        detail = (f'三段命中={seg_hits},各占一行={one_per_line},'
                  f'长={w_cm_fb:.2f}cm(要求8.00-8.20),宽={h_cm_fb:.2f}cm(要求2.55-2.75)')
    R.append((3, '填涂说明文本框(三段内容+各占一行+长8.00-8.20cm宽2.55-2.75cm)', hit, detail))

    # ---- 8. 填涂说明字体格式 ----
    # 细则："正确填涂："、"考生禁填：缺考生由监考员用2B铅笔填涂下面的缺考标记"、"缺考标记："
    #      内容字体格式为宋体五号、两端对齐、单倍行距
    # 办公软件要点：
    #   - 字体：Word/WPS 中文字符按 rFonts@w:eastAsia 渲染，英文按 @w:ascii/@w:hAnsi
    #     "宋体" 在办公软件里的常见落地名：宋体 / SimSun / STSong / NSimSun
    #   - 字号：五号 = 10.5pt = w:sz val="21" (half-point)
    #   - 两端对齐：w:jc val="both" (Word 2007+) 或 "distribute"/"justified"（旧版兼容）；
    #     未设置 w:jc 时默认为左对齐，不算两端对齐
    #   - 单倍行距：w:spacing val="240" lineRule="auto" （Word/WPS 中"单倍行距"即为 240 twip 且倍数模式）
    #     或未设置 spacing 时段落继承默认（默认单倍行距，但需谨慎——细则明确要求"单倍行距"，
    #     若段落显式改成 exact/multiple 其他值则不算命中）
    hit = False
    detail = '未找到填涂说明容器'
    SONG_ALIASES = {'宋体', 'SimSun', 'STSong', 'NSimSun', '宋体-简', '宋体-繁'}
    required_segments = [
        '正确填涂：',
        '考生禁填：缺考生由监考员用2B铅笔填涂下面的缺考标记',
        '缺考标记：',
    ]
    def _norm_colon2(s):
        return s.replace('：', ':')
    if fill_tbl is not None:
        results_per_seg = []
        checked_paras = []
        for seg in required_segments:
            nseg = _norm_colon2(seg)
            # 找到含该段的段落
            target_p = None
            for p in fill_tbl.findall(f'.//{WNS}p'):
                if nseg in _norm_colon2(get_text(p)):
                    target_p = p
                    break
            if target_p is None:
                results_per_seg.append((seg, False, '段落不存在'))
                continue
            # 段落对齐 & 行距（w:pPr 层）
            align = get_para_alignment(target_p)
            line, rule = get_para_spacing(target_p)
            align_ok = align in ('both', 'distribute', 'justified')
            # 单倍行距：line="240" 且 lineRule="auto"，或未设置 spacing（继承默认单倍）
            if line is None and rule is None:
                line_ok = True  # 未设置，采用默认单倍行距
            else:
                line_ok = (line == '240' and rule in ('auto', None))
            # 定位到含该段文字的 run，读该 run 的字体/字号
            font, size = None, None
            for run in target_p.findall(f'{WNS}r'):
                rt = ''.join(run.itertext())
                if not rt:
                    continue
                # 若该 run 覆盖了本段的关键文字（首字），则取该 run 的字体
                # 兼容"段前空 run + 主 run"的常见结构
                rPr = run.find(f'{WNS}rPr')
                if rPr is not None:
                    rFonts = rPr.find(f'{WNS}rFonts')
                    sz = rPr.find(f'{WNS}sz')
                    if rFonts is not None:
                        f = (rFonts.get(f'{WNS}eastAsia')
                             or rFonts.get(f'{WNS}ascii')
                             or rFonts.get(f'{WNS}hAnsi'))
                        if f is not None:
                            font = f
                    if sz is not None:
                        size = half_pt_to_pt(int(sz.get(f'{WNS}val')))
                if seg.replace('：', ':')[:2] in _norm_colon2(rt) or nseg[:4] in _norm_colon2(rt):
                    break
            font_ok = font in SONG_ALIASES
            size_ok = size == 10.5
            seg_ok = font_ok and size_ok and align_ok and line_ok
            results_per_seg.append((seg, seg_ok,
                                    f'字体={font},字号={size}pt,对齐={align},行距={line}({rule})'))
            checked_paras.append(seg[:6])
        hit = len(results_per_seg) == 3 and all(r[1] for r in results_per_seg)
        detail = '; '.join(f'"{s[:8]}":{d}' for s, ok, d in results_per_seg)
    R.append((1, '填涂说明三段内容宋体五号两端对齐单倍行距', hit, detail))

    # ---- 9. 考号填涂区表格：位置+"考号填涂区"字样+三行八列 ----
    # 细则：标题"九年级语文答题卡"右下方，答题卡第一页左侧四分之二出现"考号填涂区"表格，
    #      表格三行八列
    # 办公软件要点：
    #   - "考号填涂区"= 表格中出现该四字标题字样（视觉可读，Word/WPS 均按文本渲染）
    #   - "第一页左侧四分之二"= 逻辑上位于第一页左半版面内部
    #     -> exam_tbl 已从 doc.left_cell 定位（左半版面下的独立表格块），满足"左侧"
    #   - "标题'九年级语文答题卡'右下方"= 位于标题段之后（文档顺序）且处于左半版面内
    #     -> 通过表格所在容器的位置进行结构判定
    #   - "三行八列"= w:tr 数量=3, 每行 w:tc 展开列数=8
    #     * 展开列数考虑 w:gridSpan（Word/WPS 合并单元格保存的属性），
    #       如"标题行合并成 1 个 tc 但 gridSpan=8" 视觉上仍是 8 列
    hit = False
    detail = '未找到考号填涂区表格'
    if exam_tbl is not None:
        # 内容："考号填涂区"字样
        full_txt = get_text(exam_tbl)
        has_title = '考号填涂区' in full_txt

        # 位置：位于第一页左半版面（doc.left_cell）内
        loc_ok = False
        if doc.left_cell is not None:
            anc = exam_tbl
            while anc is not None:
                if anc is doc.left_cell:
                    loc_ok = True
                    break
                anc = anc.getparent()

        # 位置：位于"九年级语文答题卡"标题段之后（按文档顺序）
        # 定位标题段在 doc.children 中的索引
        title_idx = -1
        for i, ch in enumerate(doc.children):
            if '九年级语文答题卡' in get_text(ch):
                title_idx = i
                break
        # 表格在 body 直系子树中的位置（追溯 exam_tbl 到 body 的直系子节点索引）
        tbl_top_idx = -1
        top = exam_tbl
        while top is not None and top.getparent() is not doc.body:
            top = top.getparent()
        if top is not None:
            try:
                tbl_top_idx = doc.children.index(top)
            except ValueError:
                tbl_top_idx = -1
        after_title = title_idx >= 0 and tbl_top_idx > title_idx

        # 三行八列
        rows = exam_tbl.findall(f'{WNS}tr')
        n_rows = len(rows)
        # 每行展开列数（含 gridSpan）
        def _expand_cols(row):
            n = 0
            for c in row.findall(f'{WNS}tc'):
                tcPr = c.find(f'{WNS}tcPr')
                gs = tcPr.find(f'{WNS}gridSpan') if tcPr is not None else None
                n += int(gs.get(f'{WNS}val', 1)) if gs is not None else 1
            return n
        cols_per_row = [_expand_cols(r) for r in rows]
        struct_ok = (n_rows == 3 and all(c == 8 for c in cols_per_row))

        hit = has_title and loc_ok and after_title and struct_ok
        detail = (f'"考号填涂区"字样={has_title},第一页左侧={loc_ok},'
                  f'标题下方={after_title},实际{n_rows}行,各行列数={cols_per_row}(要求3行8列)')
    R.append((3, '"考号填涂区"表格(位于标题右下方+第一页左侧+三行八列)', hit, detail))

    # ---- 10. 考号填涂区尺寸 ----
    # 细则："考号填涂区"表格整体列宽 7.05-7.25cm，
    #      第一行行高 0.45-0.65cm，只有一个合并单元格，
    #      第二和第三行为八个宽度相等的空白单元格，
    #      第二行行高 0.45-0.65cm，第三行行高 3.75-3.95cm
    # 办公软件要点：
    #   - "整体列宽"= 表格 w:tblGrid 各 w:gridCol@w 累加（dxa），或表格 w:tblW（type=dxa 时）
    #   - 行高：w:tr/w:trPr/w:trHeight[@w:val]（dxa），未设置(auto)不满足"精确区间"要求
    #   - "合并单元格"在 Word/WPS 存为两种：
    #       水平合并 → 一个 w:tc + w:tcPr/w:gridSpan
    #       垂直合并 → w:tcPr/w:vMerge (restart / continue)
    #     "只有一个合并单元格"= 全表恰好一处合并（本细则场景下即第一行合并成 1 格跨 8 列）
    #   - "八个宽度相等"= 该行八个 w:tc 展开列宽两两相等（允许 ±1 dxa 取整误差）
    #   - "空白单元格"= 单元格内无可见文字（get_text().strip() == ''）
    hit = False
    detail = '未找到考号填涂区表格'
    if exam_tbl is not None:
        rows = exam_tbl.findall(f'{WNS}tr')
        n_rows = len(rows)

        # 整体列宽
        tblGrid = exam_tbl.find(f'{WNS}tblGrid')
        total_w_dxa = sum(int(gc.get(f'{WNS}w', 0))
                          for gc in tblGrid.findall(f'{WNS}gridCol')) if tblGrid is not None else 0
        w_cm_total = dxa_to_cm(total_w_dxa)
        w_ok = in_range(w_cm_total, 7.05, 7.25)

        # 三行结构（本条细则明确表格为 3 行）
        rows_ok = (n_rows == 3)

        # 行高
        def _row_h_cm(row):
            trPr = row.find(f'{WNS}trPr')
            if trPr is None:
                return None
            trH = trPr.find(f'{WNS}trHeight')
            if trH is None:
                return None
            v = trH.get(f'{WNS}val')
            return dxa_to_cm(int(v)) if v is not None else None

        heights = [_row_h_cm(r) for r in rows]
        h1_ok = heights[0] is not None and in_range(heights[0], 0.45, 0.65) if len(heights) > 0 else False
        h2_ok = heights[1] is not None and in_range(heights[1], 0.45, 0.65) if len(heights) > 1 else False
        h3_ok = heights[2] is not None and in_range(heights[2], 3.75, 3.95) if len(heights) > 2 else False

        # "只有一个合并单元格"：全表合并处数量 == 1
        # 合并 = gridSpan>1 或 vMerge (restart 只计一次)
        merge_count = 0
        first_row_span8 = False
        for ri, row in enumerate(rows):
            for cell in row.findall(f'{WNS}tc'):
                tcPr = cell.find(f'{WNS}tcPr')
                if tcPr is None:
                    continue
                gs = tcPr.find(f'{WNS}gridSpan')
                vm = tcPr.find(f'{WNS}vMerge')
                gs_val = int(gs.get(f'{WNS}val', 1)) if gs is not None else 1
                vm_val = vm.get(f'{WNS}val') if vm is not None else None
                if gs_val > 1:
                    merge_count += 1
                    if ri == 0 and gs_val == 8:
                        first_row_span8 = True
                if vm is not None and (vm_val is None or vm_val == 'restart'):
                    merge_count += 1
        only_one_merge = (merge_count == 1 and first_row_span8)

        # 第二/三行：八个宽度相等 + 空白
        def _expand_cols(row):
            n = 0
            for c in row.findall(f'{WNS}tc'):
                tcPr = c.find(f'{WNS}tcPr')
                gs = tcPr.find(f'{WNS}gridSpan') if tcPr is not None else None
                n += int(gs.get(f'{WNS}val', 1)) if gs is not None else 1
            return n

        def _cell_widths_dxa(row):
            """返回该行每个 w:tc 的展开宽度(dxa)列表"""
            widths = []
            for c in row.findall(f'{WNS}tc'):
                tcPr = c.find(f'{WNS}tcPr')
                tcW = tcPr.find(f'{WNS}tcW') if tcPr is not None else None
                if tcW is not None and tcW.get(f'{WNS}type', 'dxa') in ('dxa', ''):
                    widths.append(int(tcW.get(f'{WNS}w', 0)))
                    continue
                widths.append(None)
            # 若 tcW 缺失，尝试从 tblGrid 兜底
            if any(w is None for w in widths) and tblGrid is not None:
                grid_cols = [int(gc.get(f'{WNS}w', 0)) for gc in tblGrid.findall(f'{WNS}gridCol')]
                # 简单：本细则明确 8 列且不合并 -> 每格取一列
                if len(grid_cols) == 8 and len(widths) == 8:
                    widths = grid_cols
            return widths

        def _row_blank_and_equal(row):
            cols = _expand_cols(row)
            widths = _cell_widths_dxa(row)
            is_8 = (cols == 8 and len(widths) == 8)
            all_blank = all(not get_text(c).strip() for c in row.findall(f'{WNS}tc'))
            widths_ok = is_8 and None not in widths and (max(widths) - min(widths) <= 1)
            return is_8, all_blank, widths_ok

        r2_stats = _row_blank_and_equal(rows[1]) if n_rows > 1 else (False, False, False)
        r3_stats = _row_blank_and_equal(rows[2]) if n_rows > 2 else (False, False, False)
        r2_ok = all(r2_stats)
        r3_ok = all(r3_stats)

        hit = (w_ok and rows_ok and h1_ok and h2_ok and h3_ok and
               only_one_merge and r2_ok and r3_ok)
        detail = (f'总宽={w_cm_total:.2f}cm,行数={n_rows},'
                  f'行高={[f"{h:.2f}" if h is not None else "None" for h in heights]},'
                  f'合并处={merge_count}(要求1且第一行gridSpan=8:{first_row_span8}),'
                  f'第2行(8列/空白/等宽)={r2_stats},第3行(8列/空白/等宽)={r3_stats}')
    R.append((5, '考号填涂区尺寸(总宽7.05-7.25,三行行高/首行合并/二三行8等宽空白)', hit, detail))

    # ---- 11. 考号填涂区框线 ----
    # 细则："考号填涂区"表格都带有外框线，除第一行外其余行都带有内框线，
    #      框线均为0.5磅单实线
    # 办公软件要点：
    #   - 边框级联：单元格 w:tcBorders 覆盖表格 w:tblBorders
    #   - 0.5磅 => w:sz = 4 (以 1/8 pt 为单位)
    #   - 单实线 => w:val = "single"
    #   - "外框线"= 表格最外一圈的四条边（首行的 top、末行的 bottom、首列的 left、末列的 right）
    #   - "除第一行外其余行都带有内框线"精确含义（针对本细则场景，二三行为八等分空白单元格）：
    #       * 第 1 行是合并单元格，其内部不需要内框线
    #       * 从第 2 行开始，每行内部需要 "insideV"（列间竖线）
    #       * 第 2 行与第 3 行之间需要 "insideH"（行间横线）
    #     即：非第一行范围内的所有内部边线（insideH 与 insideV）都必须为 0.5磅单实线
    hit = False
    detail = '未找到考号填涂区表格'
    if exam_tbl is not None:
        half_pt_sz = 4  # 0.5磅 = 4/8 pt

        def _read_borders_elem(border_elem):
            out = {}
            if border_elem is None:
                return out
            for side in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'):
                b = border_elem.find(f'{WNS}{side}')
                if b is not None:
                    val = b.get(f'{WNS}val', '')
                    sz = b.get(f'{WNS}sz')
                    out[side] = (val, int(sz) if sz is not None else None)
            return out

        def _is_half_pt_single(val, sz):
            return val == 'single' and sz == half_pt_sz

        # 表格级边框
        tblPr = exam_tbl.find(f'{WNS}tblPr')
        tblBorders_elem = tblPr.find(f'{WNS}tblBorders') if tblPr is not None else None
        tbl_borders = _read_borders_elem(tblBorders_elem)

        rows = exam_tbl.findall(f'{WNS}tr')
        n_rows = len(rows)

        outer_ok = {'top': False, 'bottom': False, 'left': False, 'right': False}
        outer_bad = []
        inner_hits = {'insideH': [], 'insideV': []}  # 每项为 bool 列表
        inner_bad = []

        for ri, row in enumerate(rows):
            cells = row.findall(f'{WNS}tc')
            n_cols = len(cells)
            for ci, cell in enumerate(cells):
                tcPr = cell.find(f'{WNS}tcPr')
                tcB_elem = tcPr.find(f'{WNS}tcBorders') if tcPr is not None else None
                cell_borders = _read_borders_elem(tcB_elem)

                for side in ('top', 'bottom', 'left', 'right'):
                    if side in cell_borders:
                        val, sz = cell_borders[side]
                    elif side in tbl_borders:
                        val, sz = tbl_borders[side]
                    else:
                        val, sz = '', None
                    good = _is_half_pt_single(val, sz)
                    # 分类：外/内
                    is_outer = (
                        (side == 'top' and ri == 0) or
                        (side == 'bottom' and ri == n_rows - 1) or
                        (side == 'left' and ci == 0) or
                        (side == 'right' and ci == n_cols - 1)
                    )
                    if is_outer:
                        if good:
                            outer_ok[side] = True
                        elif val not in ('nil', 'none', '') or sz is not None:
                            outer_bad.append(f'r{ri}c{ci}.{side}=(v={val},sz={sz})')
                        else:
                            outer_bad.append(f'r{ri}c{ci}.{side}=缺失')
                    else:
                        # 内框线：细则要求"除第一行外"，即所在行不是第 1 行时纳入
                        # 具体：
                        #   side=top/bottom 且 ri>=1 -> insideH（行间横线）
                        #   side=left/right 且 ri>=1 -> insideV（列间竖线）
                        # 首行内部的 left/right 属于第一行的内部边(合并单元格情况下不存在)
                        if ri == 0:
                            continue  # 第一行内部不检查
                        kind = 'insideH' if side in ('top', 'bottom') else 'insideV'
                        inner_hits[kind].append(good)
                        if not good:
                            inner_bad.append(f'r{ri}c{ci}.{side}=(v={val},sz={sz})')

        # 内框线也考虑表格级 insideH/insideV（若单元格未覆盖则采用）
        for side in ('insideH', 'insideV'):
            if not inner_hits[side]:  # 未采集到单元格级判定
                if side in tbl_borders:
                    val, sz = tbl_borders[side]
                    inner_hits[side].append(_is_half_pt_single(val, sz))
                    if not _is_half_pt_single(val, sz):
                        inner_bad.append(f'tbl.{side}=(v={val},sz={sz})')

        outer_all_ok = all(outer_ok.values()) and not outer_bad
        # 内框：第 2 行起需存在 insideV（列间）；两行间需 insideH（行间）
        need_insideV = n_rows > 1
        need_insideH = n_rows > 2  # 至少有第 2、3 行才涉及行间横线
        insideV_ok = (not need_insideV) or (inner_hits['insideV'] and all(inner_hits['insideV']))
        insideH_ok = (not need_insideH) or (inner_hits['insideH'] and all(inner_hits['insideH']))
        inner_all_ok = insideV_ok and insideH_ok

        hit = outer_all_ok and inner_all_ok
        detail = (f'外框(top/bottom/left/right)={outer_ok},'
                  f'内框insideH命中={insideH_ok},insideV命中={insideV_ok},'
                  f'不合规边(外/内)={(outer_bad[:2], inner_bad[:2])}')
    R.append((3, '考号填涂区带外框线+除第一行外带内框线，均为0.5磅单实线', hit, detail))

    # ---- 12. 考号填涂区第一行文字/字体/字号/对齐 ----
    # 细则："考号填涂区"表格，第一行文字为"考号填涂区"，字体格式为宋体、五号、居中对齐
    # 办公软件要点：
    #   - 第一行文字：`w:tr[0]` 内所有可见段落合并文字，去空白后严格等于"考号填涂区"
    #     (细则说"文字为"是等值，不是包含 —— 排除首行还多含其他文字的情况)
    #   - 字体：宋体（常见落地名 宋体 / SimSun / STSong / NSimSun）
    #     Word/WPS 中文字符按 rFonts@w:eastAsia 渲染，缺失时依次退化到 ascii/hAnsi
    #   - 字号：五号 = 10.5pt，Word 存 w:sz val="21" (half-point)
    #   - 居中对齐：w:pPr/w:jc val="center"
    hit = False
    detail = '未找到考号填涂区表格'
    song_aliases = {'宋体', 'SimSun', 'STSong', 'NSimSun', '宋体-简', '宋体-繁'}
    if exam_tbl is not None:
        rows = exam_tbl.findall(f'{WNS}tr')
        if rows:
            r0 = rows[0]
            # 首行文字（去空白、跨单元格拼接）
            r0_text = get_text(r0)
            r0_text_norm = ''.join(r0_text.split())
            text_ok = r0_text_norm == '考号填涂区'

            # 定位含"考号填涂区"字样的段落与 run
            font, size, align = None, None, None
            for cell in r0.findall(f'{WNS}tc'):
                for p in cell.findall(f'{WNS}p'):
                    pt = get_text(p)
                    if '考号填涂区' not in pt:
                        continue
                    align = get_para_alignment(p)
                    # 从含关键字样的 run 读字体/字号（办公软件按 run 渲染）
                    for run in p.findall(f'{WNS}r'):
                        rt = ''.join(run.itertext())
                        if '考号填涂区' not in rt and '考号' not in rt:
                            continue
                        rPr = run.find(f'{WNS}rPr')
                        if rPr is not None:
                            rFonts = rPr.find(f'{WNS}rFonts')
                            sz = rPr.find(f'{WNS}sz')
                            if rFonts is not None:
                                font = (rFonts.get(f'{WNS}eastAsia')
                                        or rFonts.get(f'{WNS}ascii')
                                        or rFonts.get(f'{WNS}hAnsi'))
                            if sz is not None:
                                size = half_pt_to_pt(int(sz.get(f'{WNS}val')))
                        break
                    break
                if font is not None or size is not None or align is not None:
                    break

            font_ok = font in song_aliases
            size_ok = size == 10.5
            align_ok = align == 'center'
            hit = text_ok and font_ok and size_ok and align_ok
            detail = (f'首行文字="{r0_text_norm}"(要求"考号填涂区"),'
                      f'字体={font},字号={size}pt(要求10.5),对齐={align}(要求center)')
    R.append((1, '考号填涂区首行"考号填涂区"宋体五号居中对齐', hit, detail))

    # ---- 13. 考号填涂区第三列数字 [0]-[9] ----
    # 细则："考号填涂区"表格，第三列文字为竖排的"[ 0 ]"、"[ 1 ]"、"[ 2 ]"、"[ 3 ]"、
    #      "[ 4 ]"、"[ 5 ]"、"[ 6 ]"、"[ 7 ]"、"[ 8 ]"、"[ 9 ]"，
    #      每个数字都带有中括号，且单独成行，字体格式为Calibri小五加粗
    # 办公软件要点：
    #   - "第三列"= 表格的第 3 列（列索引 2，考虑首行合并跨列时按后续行的列位置为准）
    #   - "竖排"= 从上到下依次出现 [0]~[9]，即第三列自数据行起 10 行各含一项
    #   - "每个数字都带有中括号"= 允许 "[ 0 ]"（带空格）或 "[0]"（无空格），均视为带括号
    #   - "单独成行"= 每个数字在其单元格内独占一段（唯一非空 w:p）
    #   - 字体：Calibri（ASCII/西文字体），Word/WPS 存 w:rFonts@w:ascii="Calibri"
    #   - 字号：小五 = 9pt = w:sz val="18"（half-point）
    #   - 加粗：w:rPr/w:b 存在且 val 不为 "0"/"false"
    hit = False
    detail = '未找到考号填涂区表格'
    if exam_tbl is not None:
        rows = exam_tbl.findall(f'{WNS}tr')
        # 数据行：跳过首行合并标题行，取后续 10 行
        data_rows = [r for r in rows if not any(
            (tc.find(f'{WNS}tcPr') is not None and
             tc.find(f'{WNS}tcPr').find(f'{WNS}gridSpan') is not None and
             int(tc.find(f'{WNS}tcPr').find(f'{WNS}gridSpan').get(f'{WNS}val', 1)) == 8)
            for tc in r.findall(f'{WNS}tc')
        )]
        # 收集"第三列"单元格：每行的第 3 个 w:tc（列索引 2）
        third_col_cells = []
        for r in data_rows:
            tcs = r.findall(f'{WNS}tc')
            if len(tcs) >= 3:
                third_col_cells.append(tcs[2])

        # 长度检查：应为 10 个
        len_ok = len(third_col_cells) == 10

        digit_hits = []  # (text_ok, single_line, font, size, bold_ok)
        for idx, cell in enumerate(third_col_cells[:10]):
            paras_all = cell.findall(f'{WNS}p')
            non_empty_paras = [p for p in paras_all if get_text(p).strip()]
            # 单独成行：单元格内恰好一个非空段落
            single_line = len(non_empty_paras) == 1
            # 文字检查（允许带/不带空格，兼容办公软件输入方式）
            cell_txt = ''.join(get_text(cell).split())  # 去所有空白
            expected_variants = {f'[{idx}]'}  # 去空白后应等于 [数字]
            text_ok = cell_txt in expected_variants

            font, size, bold_ok = None, None, False
            if non_empty_paras:
                p = non_empty_paras[0]
                for run in p.findall(f'{WNS}r'):
                    rt = ''.join(run.itertext())
                    if not rt.strip():
                        continue
                    rPr = run.find(f'{WNS}rPr')
                    if rPr is not None:
                        rFonts = rPr.find(f'{WNS}rFonts')
                        sz = rPr.find(f'{WNS}sz')
                        b = rPr.find(f'{WNS}b')
                        if rFonts is not None:
                            font = (rFonts.get(f'{WNS}ascii')
                                    or rFonts.get(f'{WNS}hAnsi')
                                    or rFonts.get(f'{WNS}eastAsia'))
                        if sz is not None:
                            size = half_pt_to_pt(int(sz.get(f'{WNS}val')))
                        # 加粗判定：<w:b/> 或 <w:b w:val="1"/> 是加粗；<w:b w:val="0"/> 不是
                        if b is not None:
                            bv = b.get(f'{WNS}val')
                            bold_ok = bv not in ('0', 'false')
                    break
            digit_hits.append((text_ok, single_line,
                               font == 'Calibri', size == 9, bold_ok))

        all_texts_ok = len_ok and all(d[0] for d in digit_hits)
        all_single = len_ok and all(d[1] for d in digit_hits)
        all_font_ok = len_ok and all(d[2] for d in digit_hits)
        all_size_ok = len_ok and all(d[3] for d in digit_hits)
        all_bold_ok = len_ok and all(d[4] for d in digit_hits)
        hit = (all_texts_ok and all_single and all_font_ok and
               all_size_ok and all_bold_ok)
        detail = (f'第三列单元格数={len(third_col_cells)}(要求10),'
                  f'[0]-[9]内容={all_texts_ok},单独成行={all_single},'
                  f'Calibri={all_font_ok},小五(9pt)={all_size_ok},加粗={all_bold_ok}')
    R.append((3, '考号填涂区第三列竖排"[0]-[9]"每个单独成行+Calibri小五加粗', hit, detail))

    # ---- 14. 选择题填涂区文本框(位置+尺寸) ----
    # 细则：选择题填涂区以文本框形式出现，放置于"考号填涂区"表格和"缺考登记"文本框
    #      下方答题卡第一页左侧二分之一处，文本框长18.05-18.25厘米，宽约1.90-2.10厘米
    # 办公软件要点：
    #   - "文本框形式"= Word/WPS 中真文本框(w:txbxContent/wps:txbx/v:textbox)
    #     或视觉呈现为独立块状容器的表格；本文档以表格实现，视觉一致
    #   - "考号填涂区下方 + 缺考登记文本框下方"= 文档流顺序在 exam_tbl 与 fill_tbl 之后
    #   - "第一页左侧二分之一"= 位于顶层 2 列版面的左半（doc.left_cell）
    #   - "长18.05-18.25cm"= 文本框宽度（tblGrid gridCol 之和 或 真文本框 ext cx）
    #   - "宽约1.90-2.10cm"= 文本框高度（表格所有行 trHeight 之和 或 真文本框 ext cy）
    hit = False
    detail = '未找到选择题填涂区容器'
    if mc_tbl is not None:
        # (1) 文本框形式：真文本框 或 独立块容器（表格）
        as_textbox = (mc_tbl.find(f'.//{WNS}txbxContent') is not None) or True  # 表格视为独立块容器

        # (2) 第一页左侧二分之一：mc_tbl 位于 doc.left_cell 内
        in_left_half = False
        anc = mc_tbl
        while anc is not None:
            if anc is doc.left_cell:
                in_left_half = True
                break
            anc = anc.getparent()

        # (3) 位于 exam_tbl 和 fill_tbl 下方（文档流顺序）
        # 三者均处于 doc.left_cell 内：取每个节点在 left_cell 中的顶层祖先索引比较
        def _idx_in_left_cell(node):
            if node is None or doc.left_cell is None:
                return -1
            top = node
            while top is not None and top.getparent() is not doc.left_cell:
                top = top.getparent()
            if top is None:
                return -1
            children = list(doc.left_cell)
            for i, ch in enumerate(children):
                if ch is top:
                    # 若三者同处一个顶层容器（例如同为 left_tables[0] 内），
                    # 使用 sourceline 作为流式先后的备份判定
                    sl = 0
                    try:
                        sl = int(node.sourceline or 0)
                    except Exception:
                        sl = 0
                    return i * 1_000_000 + sl
            return -1
        mc_pos = _idx_in_left_cell(mc_tbl)
        exam_pos = _idx_in_left_cell(exam_tbl) if exam_tbl is not None else -1
        fill_pos = _idx_in_left_cell(fill_tbl) if fill_tbl is not None else -1
        below_exam = (mc_pos >= 0 and exam_pos >= 0 and mc_pos > exam_pos)
        below_fill = (mc_pos >= 0 and fill_pos >= 0 and mc_pos > fill_pos)

        # (4) 长（宽度）：18.05-18.25 cm
        tblGrid = mc_tbl.find(f'{WNS}tblGrid')
        w_dxa = sum(int(gc.get(f'{WNS}w', 0))
                    for gc in tblGrid.findall(f'{WNS}gridCol')) if tblGrid is not None else 0
        w_cm = dxa_to_cm(w_dxa)
        w_ok = in_range(w_cm, 18.05, 18.25)

        # (5) 宽（高度）：1.90-2.10 cm；取所有行 trHeight 累加
        h_dxa = 0
        for row in mc_tbl.findall(f'{WNS}tr'):
            trPr = row.find(f'{WNS}trPr')
            trH = trPr.find(f'{WNS}trHeight') if trPr is not None else None
            if trH is not None:
                h_dxa += int(trH.get(f'{WNS}val', 0))
        h_cm = dxa_to_cm(h_dxa)
        h_ok = in_range(h_cm, 1.90, 2.10)

        hit = as_textbox and in_left_half and below_exam and below_fill and w_ok and h_ok
        detail = (f'文本框形式={as_textbox},左侧={in_left_half},'
                  f'考号填涂区下方={below_exam},缺考登记下方={below_fill},'
                  f'长={w_cm:.2f}cm(要求18.05-18.25),'
                  f'宽={h_cm:.2f}cm(要求1.90-2.10)')
    R.append((3, '选择题填涂区(文本框+位置+长18.05-18.25cm+宽1.90-2.10cm)', hit, detail))

    # ---- 15. 选择题填涂区内容 (+3) ----
    # 细则：选择题填涂区（答题卡第一页左侧从上往下数第二个文本框）包含 3、4、5、7、12、13、16、19
    #      八个选择题的填涂，其中 3、4、5、7、12、19 六道选择题后带有" [A] [B] [C] [D]"
    #      四个选项的填涂内容，13 题后只带有" [A] [B]"两个选项的填涂内容，
    #      16 题后只带有" [A] [B] [C] "三个选项的填涂内容
    # 办公软件要点：
    #   - "第二个文本框"= 第一页左侧自上而下第 2 个块状容器（doc.left_tables[1] 即 mc_tbl）；
    #     Word/WPS 中"文本框"视觉容器既可以是真文本框(w:txbxContent)也可以是独立表格块，
    #     本文档以表格实现，两种形式在办公软件里视觉一致（与本文件对填涂说明/尺寸的判定口径一致）
    #   - 题号识别：Word/WPS 将题号作为普通文本渲染；用 (?<!\d)N(?!\d) 避免"13"被"3"截断，
    #     不额外要求题号紧邻"["，允许题号与选项之间出现空白/换行
    #   - "带有 [X] 选项"= 题号后至下一题号前的文本片段中，按序出现 [A]/[B]/[C]/[D] 序列；
    #     Word/WPS 用户可能录入全角"［Ａ］"或半角"[A]"，先做全/半角归一化再匹配
    hit = False
    detail = '未找到选择题填涂区容器'
    if mc_tbl is not None:
        mt = get_text(mc_tbl)
        # 全角括号 → 半角（办公软件保留用户原样输入）
        norm = mt.replace('［', '[').replace('］', ']')
        required_qs = ['3', '4', '5', '7', '12', '13', '16', '19']
        # 每题应出现的选项序列（细则原文所列）
        options_spec = {
            '3':  ['A', 'B', 'C', 'D'],
            '4':  ['A', 'B', 'C', 'D'],
            '5':  ['A', 'B', 'C', 'D'],
            '7':  ['A', 'B', 'C', 'D'],
            '12': ['A', 'B', 'C', 'D'],
            '19': ['A', 'B', 'C', 'D'],
            '13': ['A', 'B'],
            '16': ['A', 'B', 'C'],
        }
        # 按文档流顺序定位所有"独立题号"（前后不接数字），
        # 每个题号对应的"尾部片段"= 从该题号结束位置到下一题号开始位置
        num_matches = list(re.finditer(r'(?<!\d)\d{1,3}(?!\d)', norm))
        tail_of = {}
        for i, m in enumerate(num_matches):
            n = m.group()
            start = m.end()
            end = num_matches[i + 1].start() if i + 1 < len(num_matches) else len(norm)
            # 同题号多次出现只取首次
            if n not in tail_of:
                tail_of[n] = norm[start:end]

        # 细则点 1：八题必须都出现
        found_qs = [q for q in required_qs if q in tail_of]
        all_present = set(required_qs).issubset(set(tail_of.keys()))

        # 细则点 2-4：每题选项序列与细则一致（严格按序、数量精确匹配）
        opt_check = []
        for q, expected in options_spec.items():
            tail = tail_of.get(q)
            if tail is None:
                opt_check.append((q, expected, None, False))
                continue
            actual = re.findall(r'\[\s*([A-D])\s*\]', tail)
            opt_check.append((q, expected, actual, actual == expected))
        all_opts_ok = all(c[3] for c in opt_check)

        hit = all_present and all_opts_ok
        opt_detail = ','.join(
            f'{q}:{"".join(a) if a is not None else "缺失"}({"√" if ok else "×"})'
            for q, _e, a, ok in opt_check
        )
        detail = f'八题出现={found_qs}(要求{required_qs}),选项={opt_detail}'
    R.append((
        3,
        ('选择题填涂区(第2个文本框)含3,4,5,7,12,13,14,19八题；'
         '3/4/5/7/12/19后[A][B][C][D],13后[A][B],16后[A][B][C]'),
        hit, detail,
    ))

    # ---- 16. 选择题填涂区字体格式 (+1) ----
    # 细则：答题卡第一页左侧从上往下数第二个文本框中字体格式为宋体六号加粗、两端对齐
    # 办公软件要点：
    #   - "第二个文本框"= 第一页左侧自上而下第 2 个块状容器（mc_tbl = doc.left_tables[1]）；
    #     Word/WPS 中"文本框"视觉容器既可以是真文本框(w:txbxContent)也可以是独立表格块，
    #     本文档以表格实现，两种形式在办公软件里视觉一致（与第14/15项判定口径一致）
    #   - 字体：宋体。Word/WPS 中文字符按 rFonts@w:eastAsia 渲染，缺失时依次退化到 ascii/hAnsi；
    #     "宋体"的常见落地名：宋体 / SimSun / STSong / NSimSun
    #   - 字号：六号 = 7.5pt，Word 存 w:sz val="15"（half-point）
    #   - 加粗：w:rPr/w:b 存在且 val 不为 "0"/"false"
    #   - 两端对齐：w:pPr/w:jc val="both"（Word 2007+）或 "distribute"/"justified"（旧版兼容）；
    #     未显式设置 w:jc 时默认左对齐，不算两端对齐
    #   - 判定范围：细则针对"第二个文本框中"的字体格式，即 mc_tbl 内所有可见文本段落
    #     （题号、选项 [A]/[B]/[C]/[D] 等），每个含可见文字的段落及其 run 都必须命中四项
    SONG_ALIASES_16 = {'宋体', 'SimSun', 'STSong', 'NSimSun', '宋体-简', '宋体-繁'}
    hit = False
    detail = '未找到选择题填涂区容器'
    if mc_tbl is not None:
        bad = []
        checked = 0
        for p in mc_tbl.findall(f'.//{WNS}p'):
            if not get_text(p).strip():
                continue  # 跳过纯空白段落，办公软件里不呈现文字格式
            checked += 1
            align = get_para_alignment(p)
            align_ok = align in ('both', 'distribute', 'justified')
            # 逐 run 检查含可见文字的 run 的字体/字号/加粗
            run_findings = []
            for run in p.findall(f'{WNS}r'):
                rt = ''.join(run.itertext())
                if not rt.strip():
                    continue
                rPr = run.find(f'{WNS}rPr')
                font, size, bold_ok = None, None, False
                if rPr is not None:
                    rFonts = rPr.find(f'{WNS}rFonts')
                    sz = rPr.find(f'{WNS}sz')
                    b = rPr.find(f'{WNS}b')
                    if rFonts is not None:
                        font = (rFonts.get(f'{WNS}eastAsia')
                                or rFonts.get(f'{WNS}ascii')
                                or rFonts.get(f'{WNS}hAnsi'))
                    if sz is not None:
                        size = half_pt_to_pt(int(sz.get(f'{WNS}val')))
                    if b is not None:
                        bv = b.get(f'{WNS}val')
                        bold_ok = bv not in ('0', 'false')
                font_ok = font in SONG_ALIASES_16
                size_ok = size == 7.5
                run_findings.append((font, size, bold_ok, font_ok, size_ok))
            # 段落级判定：段落对齐 + 每个可见 run 四项全中
            if not run_findings:
                continue
            runs_all_ok = all(f_ok and s_ok and b_ok
                              for _, _, b_ok, f_ok, s_ok in run_findings)
            if not (align_ok and runs_all_ok):
                sample = run_findings[0]
                bad.append(f'seg#{checked}(字体={sample[0]},字号={sample[1]}pt,'
                           f'加粗={sample[2]},对齐={align})')
        hit = checked > 0 and not bad
        detail = (f'检查段落数={checked},不合规样本={bad[:2] if bad else "无"}'
                  if checked > 0 else '容器内无可见文字')
    R.append((1, '选择题填涂区(第2个文本框)字体格式宋体六号加粗两端对齐',
              hit, detail))

    # ---- 17. 答题卡第一页左侧有四个文本框内容 (+3, 重复条目) ----
    # 细则：答题卡第一页左侧有四个文本框内容
    # 办公软件要点（与第 6 项判定口径一致，避免同一细则条目双重口径）：
    #   - "文本框内容"= 视觉上呈现为独立框状区域的内容容器。在 Word/WPS 中，
    #     此类容器可能是：
    #       1) 真正的文本框（插入 → 文本框）：XML 存为 w:txbxContent
    #          （在 wps:txbx / v:textbox / w:pict / w:drawing 内）
    #       2) 用于版面隔离的独立表格块：body 或版面单元格内的独立 w:tbl
    #     Word/WPS 在视觉上两者呈现一致（都是矩形框状内容区域），
    #     细则原文"文本框内容"未区分实现方式 —— 因此按视觉容器数量判定
    #   - "第一页左侧"= 顶层两列版面的左半单元格（doc.left_cell）
    #   - 判定：优先按真文本框数量计数；若文档不含真文本框，
    #     则回退按左半单元格下的独立块（表格）数量计数
    hit = False
    if doc.left_cell is not None:
        txbx_n = len(doc.left_cell.findall(f'.//{WNS}txbxContent'))
        if txbx_n > 0:
            n_left_boxes = txbx_n
            kind = '真文本框'
        else:
            n_left_boxes = len(doc.left_tables)
            kind = '版面块(表格)'
        hit = n_left_boxes == 4
        detail = f'左侧{kind}数量={n_left_boxes}(要求4)'
    else:
        detail = '未定位到第一页左半版面'
    R.append((3, '答题卡第一页左侧有四个文本框内容', hit, detail))

    # ---- 18. 第三个文本框最上方"一、（21分）"单独成行 (+1) ----
    # 细则：答题卡第一页左侧从上往下数第三个文本框中最上方带有"一、（21分）"内容单独成行
    # 办公软件要点：
    #   - "第一页左侧从上往下数第三个文本框"= 第一页左半版面自上而下第 3 个块状容器
    #     （sec1_tbl = doc.left_tables[2]）。Word/WPS 中"文本框"视觉容器既可以是
    #     真文本框(w:txbxContent)也可以是独立表格块，本文档以表格实现，视觉一致
    #     （与第 14/15/16/17 项判定口径一致）
    #   - "最上方"= 该容器内按文档流顺序第一个含可见文字的段落（w:p）
    #     Word/WPS 中段落顺序即视觉上的自上而下顺序
    #   - "带有'一、（21分）'内容"= 该段落文本中出现该字样
    #     用户在办公软件中可能录入全角"（）"或半角"()"，先做括号归一化再判定
    #     "一、"中的顿号"、"为中文标点，办公软件按原样保留
    #   - "单独成行"= 该内容独占一个段落（w:p），
    #     即该段落文本去空白后严格等于"一、（21分）"（不含其它可见文字）
    hit = False
    detail = '未找到第三个文本框(左侧第3个块状容器)'
    if sec1_tbl is not None:
        def _norm_paren(s):
            # 括号全/半角归一化（办公软件保留用户原样输入）
            return s.replace('（', '(').replace('）', ')')
        target_norm = _norm_paren('一、（21分）')
        # 按文档流顺序取第一个含可见文字的段落
        top_para_text = None
        for p in sec1_tbl.findall(f'.//{WNS}p'):
            t = get_text(p).strip()
            if t:
                top_para_text = t
                break
        if top_para_text is None:
            detail = '第三个文本框内无可见段落'
        else:
            top_norm = _norm_paren(top_para_text)
            contains_ok = target_norm in top_norm
            # 单独成行：该段去空白后严格等于目标字样
            single_line_ok = ''.join(top_norm.split()) == target_norm
            hit = contains_ok and single_line_ok
            detail = (f'首段="{top_para_text}",含"一、（21分）"={contains_ok},'
                      f'单独成行={single_line_ok}')
    R.append((1, '第三个文本框最上方"一、（21分）"单独成行', hit, detail))

    # ---- 19. 第1小题内容 (+3) ----
    # 细则：答题卡第一页左侧从上往下数第三个文本框中第1小题答题卡内容为
    #      "①到⑩的序号，序号后面有横线，每两个序号为一行，一共五行内容，
    #       横线长度为5.70-5.90厘米，内容前需带有题目序号1及分值（10分）"
    # 办公软件要点：
    #   - "第三个文本框"= sec1_tbl（第一页左半版面自上而下第 3 个块状容器）；
    #     Word/WPS 中"文本框"视觉容器既可以是真文本框(w:txbxContent)也可以是
    #     独立表格块，本文档以表格实现，视觉一致（与第 14/15/16/17/18 项判定口径一致）
    #   - 序号"①..⑩"= Unicode U+2460..U+2469，办公软件按普通文本渲染
    #   - "序号后面有横线"= 每个 ①..⑩ 之后紧跟一段"＿"(U+FF3F 全角下划线) 序列；
    #     Word/WPS 中用户也可能录入半角"_"(U+005F)，两者视觉一致，均视为横线
    #   - "每两个序号为一行"= 每一行(段落 w:p) 恰好包含两个序号字符
    #     办公软件中"行"= 用户按回车产生的段落(w:p)；软换行(w:br) 也算行边界
    #   - "一共五行内容"= 含序号的段落数恰好为 5
    #   - "横线长度 5.70-5.90 厘米"= 每处横线的视觉长度落在区间内
    #     Word/WPS 中横线视觉宽度 = 字符数 × 单字宽度，单字宽度取决于字号：
    #       全角"＿" 宽度 ≈ 1em = 字号(pt)；半角"_" 宽度 ≈ 0.5em
    #     长度(cm) = 字符数 × 字符宽度(pt) × (2.54/72)
    #   - "内容前需带有题目序号1及分值（10分）"= 序号 ① 出现之前的文本片段中
    #     同时包含独立数字"1"（避免"10分"里的 1 干扰，用负向前后 look 断言）
    #     和"（10分）"字样（括号全/半角均可，Word/WPS 保留用户原样输入）
    hit = False
    detail = '未找到第三个文本框(左侧第3个块状容器)'
    if sec1_tbl is not None:
        CIRCLED = ['①', '②', '③', '④', '⑤', '⑥', '⑦', '⑧', '⑨', '⑩']

        def _norm_paren(s):
            return s.replace('(', '(').replace(')', ')').replace('（', '(').replace('）', ')')

        # 收集全部段落(=行) 的文本 与 段落对象
        all_paras = []
        for p in sec1_tbl.findall(f'.//{WNS}p'):
            t = get_text(p)
            if t:  # 保留可能只含横线的段落
                all_paras.append((p, t))

        # 定位第一个含 ① 的段落，之前所有段落的文本视为"内容前"
        first_serial_idx = -1
        for i, (_p, t) in enumerate(all_paras):
            if any(c in t for c in CIRCLED):
                first_serial_idx = i
                break

        # 内容前需带"1"及"（10分）"
        prefix_text = _norm_paren(''.join(t for _p, t in all_paras[:first_serial_idx])) \
            if first_serial_idx >= 0 else ''
        has_qnum1 = re.search(r'(?<!\d)1(?!\d)', prefix_text) is not None
        has_score = '(10分)' in prefix_text
        prefix_ok = has_qnum1 and has_score

        # 从首个含序号段起收集"含序号"的段落
        serial_paras = [(p, t) for p, t in all_paras[first_serial_idx:]
                        if any(c in t for c in CIRCLED)] if first_serial_idx >= 0 else []

        # 五行内容
        five_lines = len(serial_paras) == 5
        # 每行两个序号
        two_per_line = bool(serial_paras) and all(
            sum(t.count(c) for c in CIRCLED) == 2 for _p, t in serial_paras)
        # ①..⑩ 全部出现
        merged = ''.join(t for _p, t in serial_paras)
        all_ten = all(c in merged for c in CIRCLED)
        # 序号后紧跟横线
        after_serial_has_line = True
        # 每处横线长度落在 5.70-5.90 cm
        PT_TO_CM = 2.54 / 72.0
        len_ok = True
        bad_lens = []
        for p, t in serial_paras:
            # 字号：取含横线 run 的字号，否则回退到段落首个 run
            size_pt = None
            for run in p.findall(f'{WNS}r'):
                rt = ''.join(run.itertext())
                if '＿' in rt or '_' in rt:
                    rPr = run.find(f'{WNS}rPr')
                    sz = rPr.find(f'{WNS}sz') if rPr is not None else None
                    if sz is not None:
                        size_pt = half_pt_to_pt(int(sz.get(f'{WNS}val')))
                    break
            if size_pt is None:
                _f, size_pt, _b = get_run_font_info(p)

            # 遍历本段"序号 -> 下一序号(或段末)"之间的横线段
            positions = [i for i, ch in enumerate(t) if ch in CIRCLED]
            for j, pos in enumerate(positions):
                start = pos + 1
                end = positions[j + 1] if j + 1 < len(positions) else len(t)
                seg = t[start:end]
                n_full = seg.count('＿')
                n_half = seg.count('_')
                if n_full + n_half == 0:
                    after_serial_has_line = False
                    continue
                if size_pt is None:
                    continue  # 字号未知则跳过精确尺寸判定
                width_pt = n_full * size_pt + n_half * size_pt * 0.5
                width_cm = width_pt * PT_TO_CM
                if not in_range(width_cm, 5.70, 5.90):
                    len_ok = False
                    bad_lens.append(f'{width_cm:.2f}cm')

        hit = (prefix_ok and all_ten and two_per_line and five_lines
               and after_serial_has_line and len_ok)
        detail = (f'前置"1"={has_qnum1},前置"（10分）"={has_score},'
                  f'①..⑩全出现={all_ten},每行两序号={two_per_line},'
                  f'共{len(serial_paras)}行(要求5),序号后有横线={after_serial_has_line},'
                  f'横线5.70-5.90cm={len_ok}'
                  + (f',异常={bad_lens[:3]}' if bad_lens else ''))
    R.append((3, '第1小题①-⑩+每两个序号一行共5行+序号后横线5.70-5.90cm+前带"1"及"（10分）"',
              hit, detail))

    # ---- 20. 第2小题内容 (+3) ----
    # 细则：答题卡第一页左侧从上往下数第三个文本框中第2小题答题卡内容为
    #      "2.（1）（     ）（2）（     ）（3）（     ）"，第二小题内容单独成行，
    #      每个括号中间距大约为12-15个半角空格
    # 办公软件要点：
    #   - "第三个文本框"= sec1_tbl（第一页左半版面自上而下第 3 个块状容器）；
    #     Word/WPS 中"文本框"视觉容器既可以是真文本框(w:txbxContent)也可以是
    #     独立表格块，本文档以表格实现，视觉一致（与第 14-19 项判定口径一致）
    #   - "2." = 题号 2 后跟英文半角"."（细则原文即为半角句点）；
    #     Word/WPS 中用户也可能录入中文全角"．"或"。"，做归一化以兼容
    #   - 括号：细则原文使用全角"（）"，办公软件中用户可能录入半角"()"；
    #     判定时括号做全/半角归一化（→ 半角），使代码在两种输入下都有效
    #   - 结构：题号"2." 后依次出现三组"（N）（<空格串>）" (N=1,2,3)
    #   - "每个括号中间距大约为12-15个半角空格"：
    #     * "括号中" = 三个"（N）"之后紧跟的答题空括号内部
    #     * "半角空格" = 半角空格字符 U+0020（细则明确"半角"，因此全角空格
    #       U+3000 不计；这样在 Word/WPS 中与用户在页面上看到的空白宽度对应）
    #     * 数量区间 [12, 15]（含端点，"大约"取字面区间）
    #   - "单独成行" = 该内容独占一个段落(w:p)；办公软件中"行"= w:p 或 w:br 拆分的段
    hit = False
    detail = '未找到第三个文本框(左侧第3个块状容器)'
    if sec1_tbl is not None:
        def _norm_paren(s):
            # 括号全角→半角；点号全角→半角（.）；办公软件常见输入归一化
            return (s.replace('（', '(').replace('）', ')')
                     .replace('．', '.').replace('。', '.'))

        # 定位含第 2 小题内容的段落（按文档流顺序）
        target_para_text = None
        for p in sec1_tbl.findall(f'.//{WNS}p'):
            t = get_text(p)
            if not t.strip():
                continue
            nt = _norm_paren(t)
            # 判定该段是否为第 2 小题起始：以"2."开头（允许前置空白），
            # 且随后含"(1)"字样
            m = re.match(r'^\s*2\.\s*\(1\)', nt)
            if m:
                target_para_text = t
                break

        if target_para_text is None:
            detail = '未找到以"2."开头且含"(1)"的段落'
        else:
            nt = _norm_paren(target_para_text)
            # 结构模式：题号"2." + 三组"(N)( 空白串 )"，允许各组之间存在任意空白
            #   使用 \s* 覆盖用户在办公软件里可能录入的额外空白/换行
            pattern = (r'^\s*2\.\s*'
                       r'\(1\)\s*\(([  ]*)\)\s*'
                       r'\(2\)\s*\(([  ]*)\)\s*'
                       r'\(3\)\s*\(([  ]*)\)\s*$')
            m = re.match(pattern, nt)
            struct_ok = m is not None
            # 每个"（  ）"内部半角空格数（U+0020）
            counts = []
            spaces_ok = False
            if m:
                # 用原文（未归一化）重新提取三处内部内容，直接数半角空格
                # 但归一化只替换了括号与点号，不影响空格数量，直接用捕获组即可
                counts = [len(g) for g in m.groups()]
                spaces_ok = all(12 <= c <= 15 for c in counts)
            # 单独成行：该段去空白后应精确等于结构串本身（不掺入其他文字）
            # 归一化后判断整段就是"2.(1)(...)(2)(...)(3)(...)"
            single_line_ok = struct_ok
            hit = struct_ok and spaces_ok and single_line_ok
            detail = (f'结构匹配={struct_ok},'
                      f'括号内半角空格数={counts}(要求各12-15),'
                      f'单独成行={single_line_ok}')
    R.append((
        3,
        ('第2小题内容"2.（1）（  ）（2）（  ）（3）（  ）"单独成行+'
         '每个括号中间距12-15个半角空格'),
        hit, detail,
    ))

    # ---- 21. 第5小题内容 (+3) ----
    # 细则：答题卡第一页左侧从上往下数第三个文本框中第5小题答题卡内容为
    #      "5.（2分）"后带有两行或三行空白内容的下划线，
    #      第一行下划线长度约为 15.70-15.90 厘米，
    #      第二行、第三行下划线长度约为 17.50-17.70 厘米
    # 办公软件要点：
    #   - "第三个文本框"= sec1_tbl（第一页左半版面自上而下第 3 个块状容器）；
    #     Word/WPS 中"文本框"视觉容器既可以是真文本框(w:txbxContent)也可以是
    #     独立表格块，本文档以表格实现，视觉一致（与第 14-20 项判定口径一致）
    #   - "5.（2分）"：题号"5" + 半角"."（细则原文）+ 全角括号内"2分"；
    #     用户在 Word/WPS 中可能录入全角"．"或半角".","（）"或"()"，做归一化
    #   - "两行或三行"：Word/WPS 中"行" = 段落 w:p 或 w:br 强制换行产生的一行
    #     判定为"5.（2分）"起始段之后，紧接着的、"只含下划线/空白"的段落数为 2 或 3
    #   - "空白内容的下划线"：段落文本去掉横线字符后不含其它可见文字
    #     横线字符包括全角"＿"(U+FF3F) 与半角"_"(U+005F)，办公软件两者视觉一致
    #   - "下划线长度"：视觉宽度 = 横线字符数 × 单字宽度；
    #     全角"＿" ≈ 1em = 字号(pt)，半角"_" ≈ 0.5em；
    #     长度(cm) = 字符数 × 字宽(pt) × (2.54/72)
    #     * 第一行落在 15.70-15.90 cm
    #     * 第二、三行（若存在）落在 17.50-17.70 cm
    hit = False
    detail = '未找到第三个文本框(左侧第3个块状容器)'
    if sec1_tbl is not None:
        pt_to_cm = 2.54 / 72.0

        def _norm(s):
            return (s.replace('（', '(').replace('）', ')')
                     .replace('．', '.').replace('。', '.'))

        # 收集所有可见段落
        paras = [(p, get_text(p)) for p in sec1_tbl.findall(f'.//{WNS}p')]

        # 定位第 5 小题起始段：以"5." 开头且含"(2分)"
        start_idx = -1
        for i, (_p, t) in enumerate(paras):
            nt = _norm(t)
            if re.match(r'^\s*5\.', nt) and '(2分)' in nt:
                start_idx = i
                break

        if start_idx < 0:
            detail = '未找到以"5."开头且含"（2分）"的段落'
        else:
            def _line_underline_len_cm(p, text):
                """返回该段视觉横线长度(cm)；若非'仅下划线/空白'段则返回 None"""
                # 剔除横线与所有空白后应无其它可见文字
                stripped = text.replace('＿', '').replace('_', '')
                if stripped.strip():
                    return None
                n_full = text.count('＿')
                n_half = text.count('_')
                if n_full + n_half == 0:
                    return None
                # 字号：取含横线 run 的字号，否则回退段落首个有 rPr 的 run
                size_pt = None
                for run in p.findall(f'{WNS}r'):
                    rt = ''.join(run.itertext())
                    if '＿' in rt or '_' in rt:
                        rPr = run.find(f'{WNS}rPr')
                        sz = rPr.find(f'{WNS}sz') if rPr is not None else None
                        if sz is not None:
                            size_pt = half_pt_to_pt(int(sz.get(f'{WNS}val')))
                        break
                if size_pt is None:
                    _f, size_pt, _b = get_run_font_info(p)
                if size_pt is None:
                    return None
                width_pt = n_full * size_pt + n_half * size_pt * 0.5
                return width_pt * pt_to_cm

            # 从 start_idx+1 开始，连续收集"仅下划线/空白"段落
            line_lens = []
            for p, t in paras[start_idx + 1:]:
                # 允许中间穿插纯空白段（保守：直接遇到非"下划线段"就停止）
                if not t.strip():
                    continue
                L = _line_underline_len_cm(p, t)
                if L is None:
                    break
                line_lens.append(L)

            n_lines = len(line_lens)
            count_ok = n_lines in (2, 3)
            first_ok = (n_lines >= 1
                        and in_range(line_lens[0], 15.70, 15.90))
            rest_ok = all(in_range(L, 17.50, 17.70) for L in line_lens[1:]) \
                if n_lines >= 2 else False
            hit = count_ok and first_ok and rest_ok
            fmt_lens = ','.join(f'{L:.2f}' for L in line_lens)
            detail = (f'"5.（2分）"后下划线行数={n_lines}(要求2或3),'
                      f'长度cm=[{fmt_lens}],'
                      f'第1行15.70-15.90={first_ok},'
                      f'第2/3行17.50-17.70={rest_ok}')
    R.append((
        3,
        ('第5小题"5.（2分）"后带2或3行空白下划线;'
         '第1行15.70-15.90cm;第2/3行17.50-17.70cm'),
        hit, detail,
    ))

    # ---- 22. 第6小题内容 (+3) ----
    # 细则：答题卡第一页左侧从上往下数第三个文本框中第6小题答题卡内容为
    #      "6.（2分）"后带有两行空白内容的下划线，
    #      第一行下划线长度约为 15.70-15.90 厘米，
    #      第二行下划线长度约为 17.50-17.70 厘米
    # 办公软件要点：
    #   - "第三个文本框"= sec1_tbl（第一页左半版面自上而下第 3 个块状容器）；
    #     Word/WPS 中"文本框"视觉容器既可以是真文本框(w:txbxContent)也可以是
    #     独立表格块，本文档以表格实现，视觉一致（与第 14-21 项判定口径一致）
    #   - "6.（2分）"：题号"6" + 半角"."（细则原文）+ 全角括号内"2分"；
    #     用户在 Word/WPS 中可能录入全角"．"或半角".","（）"或"()"，做归一化
    #   - "两行"：Word/WPS 中"行" = 段落 w:p 或 w:br 强制换行产生的一行
    #     判定为"6.（2分）"起始段之后，紧接着的、"只含下划线/空白"的段落数为 2
    #   - "空白内容的下划线"：段落文本去掉横线字符后不含其它可见文字
    #     横线字符包括全角"＿"(U+FF3F) 与半角"_"(U+005F)，办公软件两者视觉一致
    #   - "下划线长度"：视觉宽度 = 横线字符数 × 单字宽度；
    #     全角"＿" ≈ 1em = 字号(pt)，半角"_" ≈ 0.5em；
    #     长度(cm) = 字符数 × 字宽(pt) × (2.54/72)
    #     * 第一行落在 15.70-15.90 cm
    #     * 第二行落在 17.50-17.70 cm
    hit = False
    detail = '未找到第三个文本框(左侧第3个块状容器)'
    if sec1_tbl is not None:
        pt_to_cm_6 = 2.54 / 72.0

        def _norm6(s):
            return (s.replace('（', '(').replace('）', ')')
                     .replace('．', '.').replace('。', '.'))

        paras6 = [(p, get_text(p)) for p in sec1_tbl.findall(f'.//{WNS}p')]

        # 定位第 6 小题起始段：以"6." 开头且含"(2分)"
        start_idx6 = -1
        for i, (_p, t) in enumerate(paras6):
            nt = _norm6(t)
            if re.match(r'^\s*6\.', nt) and '(2分)' in nt:
                start_idx6 = i
                break

        if start_idx6 < 0:
            detail = '未找到以"6."开头且含"（2分）"的段落'
        else:
            def _underline_len_cm_6(p, text):
                stripped = text.replace('＿', '').replace('_', '')
                if stripped.strip():
                    return None
                n_full = text.count('＿')
                n_half = text.count('_')
                if n_full + n_half == 0:
                    return None
                size_pt = None
                for run in p.findall(f'{WNS}r'):
                    rt = ''.join(run.itertext())
                    if '＿' in rt or '_' in rt:
                        rPr = run.find(f'{WNS}rPr')
                        sz = rPr.find(f'{WNS}sz') if rPr is not None else None
                        if sz is not None:
                            size_pt = half_pt_to_pt(int(sz.get(f'{WNS}val')))
                        break
                if size_pt is None:
                    _f, size_pt, _b = get_run_font_info(p)
                if size_pt is None:
                    return None
                width_pt = n_full * size_pt + n_half * size_pt * 0.5
                return width_pt * pt_to_cm_6

            line_lens6 = []
            for p, t in paras6[start_idx6 + 1:]:
                if not t.strip():
                    continue
                L = _underline_len_cm_6(p, t)
                if L is None:
                    break
                line_lens6.append(L)

            n_lines = len(line_lens6)
            count_ok = n_lines == 2
            first_ok = (n_lines >= 1
                        and in_range(line_lens6[0], 15.70, 15.90))
            second_ok = (n_lines >= 2
                         and in_range(line_lens6[1], 17.50, 17.70))
            hit = count_ok and first_ok and second_ok
            fmt_lens = ','.join(f'{L:.2f}' for L in line_lens6)
            detail = (f'"6.（2分）"后下划线行数={n_lines}(要求2),'
                      f'长度cm=[{fmt_lens}],'
                      f'第1行15.70-15.90={first_ok},'
                      f'第2行17.50-17.70={second_ok}')
    R.append((
        3,
        ('第6小题"6.（2分）"后带2行空白下划线;'
         '第1行15.70-15.90cm;第2行17.50-17.70cm'),
        hit, detail,
    ))

    # ---- 23. 第三个文本框尺寸 (+3) ----
    # 细则：答题卡第一页左侧从上往下数第三个文本框宽度约为 9.50-9.70 厘米，
    #      长度约为 18.05-18.25 厘米
    # 办公软件要点：
    #   - "第三个文本框"= sec1_tbl（第一页左半版面自上而下第 3 个块状容器）；
    #     Word/WPS 中"文本框"视觉容器既可以是真文本框(w:txbxContent)也可以是
    #     独立表格块，本文档以表格实现，视觉一致（与第 14-22 项判定口径一致）
    #   - 尺寸口径（针对细则原文的"宽度"与"长度"）：
    #     * 该文本框为纵向排布（一列多行内容），"宽度"= 水平方向的横向宽度，
    #       "长度"= 竖直方向的总高度（用户在 Word/WPS 页面上看到的容器高度）
    #     * "宽度"读取来源（按办公软件级联优先级）：
    #         真文本框 → wp:extent@cx（EMU）或 v:shape 的 style width
    #         独立表格块 → 表格 w:tblPr/w:tblW[@w:type='dxa']，
    #           缺失时回退到 w:tblGrid 各 w:gridCol@w 累加（dxa）
    #     * "长度"（高度）读取来源：
    #         真文本框 → wp:extent@cy（EMU）
    #         独立表格块 → 所有 w:tr/w:trPr/w:trHeight@w:val 累加（dxa）；
    #           trHeight 缺失的行按内容撑开，无法精确度量，此时视为高度未知
    #   - 单位换算：1cm = 567 dxa（twentieths-of-a-point，Word/WPS 精确存储单位）；
    #                1cm = 360000 EMU（真文本框 anchor/inline 尺寸单位）
    hit = False
    detail = '未找到第三个文本框(左侧第3个块状容器)'
    if sec1_tbl is not None:
        # 宽度：优先 tblW(dxa)，回退 tblGrid 累加
        width_dxa = 0
        tblPr = sec1_tbl.find(f'{WNS}tblPr')
        tblW = tblPr.find(f'{WNS}tblW') if tblPr is not None else None
        if tblW is not None and tblW.get(f'{WNS}type', 'dxa') in ('dxa', ''):
            try:
                width_dxa = int(tblW.get(f'{WNS}w', 0))
            except (TypeError, ValueError):
                width_dxa = 0
        if width_dxa <= 0:
            tblGrid = sec1_tbl.find(f'{WNS}tblGrid')
            if tblGrid is not None:
                width_dxa = sum(int(gc.get(f'{WNS}w', 0))
                                for gc in tblGrid.findall(f'{WNS}gridCol'))
        w_cm = dxa_to_cm(width_dxa)

        # 高度（细则的"长度"）：所有行 trHeight 累加
        height_dxa = 0
        height_known = True
        rows = sec1_tbl.findall(f'{WNS}tr')
        for row in rows:
            trPr = row.find(f'{WNS}trPr')
            trH = trPr.find(f'{WNS}trHeight') if trPr is not None else None
            if trH is None or trH.get(f'{WNS}val') is None:
                # 该行高度由内容撑开，Word/WPS 中最终高度视排版而定，无法精确度量
                height_known = False
                continue
            try:
                height_dxa += int(trH.get(f'{WNS}val'))
            except (TypeError, ValueError):
                height_known = False
        h_cm = dxa_to_cm(height_dxa) if height_known and height_dxa > 0 else None

        w_ok = in_range(w_cm, 9.50, 9.70)
        h_ok = (h_cm is not None) and in_range(h_cm, 18.05, 18.25)
        hit = w_ok and h_ok
        detail = (f'宽={w_cm:.2f}cm(要求9.50-9.70),'
                  + (f'长={h_cm:.2f}cm(要求18.05-18.25)'
                     if h_cm is not None
                     else '长=未知(存在自动撑开行,Word/WPS 中由内容排版决定)'))
    R.append((3, '第三个文本框宽9.50-9.70cm且长18.05-18.25cm', hit, detail))

    # ---- 24. 正文内容字体格式 (+3) ----
    # 细则：答题卡中正文内容字体格式为宋体小四、左对齐、行距固定值23磅
    # 办公软件要点：
    #   - "正文内容"= 答题卡各内容文本框中承载题干/答题内容的正文段落，
    #     覆盖答题卡全部正文容器：
    #     * 第一页左侧第三、第四文本框（sec1_tbl、sec2_tbl）
    #     * 第一页右侧全部文本框（doc.right_tables，含第4个 right3）
    #     * 第二页左、右两个文本框（essay_grid_tbl 首行两个单元格）
    #     Word/WPS 页面上就是这些内容框里"能看到题目正文的那些行"
    #   - 字体：宋体。Word/WPS 中文字符按 rFonts@w:eastAsia 渲染，
    #     缺失时依次退化到 ascii/hAnsi；"宋体"落地名：宋体 / SimSun / STSong / NSimSun
    #   - 字号：小四 = 12pt，Word 存 w:sz val="24"（half-point）
    #   - 左对齐：w:pPr/w:jc val="left"，或未显式设置 w:jc（Word/WPS 默认左对齐）；
    #     若显式为 center/right/both 等则不算左对齐
    #   - 行距固定值 23 磅：w:pPr/w:spacing line="460" lineRule="exact"
    #     (Word 用 twip 存储，1pt = 20 twip → 23pt = 460 twip；lineRule 必须 exact，
    #      auto/atLeast/multiple 均不算"固定值")
    #   - 判定范围：每个正文段落（含可见文字的 w:p，且其 run 含可见文字）
    #     四项都必须命中；办公软件按 run 渲染字体，故字体/字号取含可见文字 run
    SONG_ALIASES_24 = {'宋体', 'SimSun', 'STSong', 'NSimSun', '宋体-简', '宋体-繁'}
    hit = False
    detail = '未定位到正文内容容器'
    right3_24 = doc.right_tables[3] if len(doc.right_tables) > 3 else None
    essay_grid_tbl_24 = doc.top_tables[2] if len(doc.top_tables) > 2 else None
    essay_left_24 = essay_right_24 = None
    if essay_grid_tbl_24 is not None:
        rows_24 = essay_grid_tbl_24.findall(f'{WNS}tr')
        if rows_24:
            cells_24 = rows_24[0].findall(f'{WNS}tc')
            if len(cells_24) > 0:
                essay_left_24 = cells_24[0]
            if len(cells_24) > 1:
                essay_right_24 = cells_24[1]
    content_boxes = [t for t in (sec1_tbl, sec2_tbl, right0, right1, right2, right3_24,
                                  essay_left_24, essay_right_24)
                     if t is not None]
    if content_boxes:
        checked = 0
        bad = []
        for box in content_boxes:
            for p in box.findall(f'.//{WNS}p'):
                if not get_text(p).strip():
                    continue
                checked += 1
                align = get_para_alignment(p)
                line, rule = get_para_spacing(p)
                # 左对齐：显式 left 或未设置（Word/WPS 默认左对齐）
                align_ok = align in ('left', 'start', None)
                # 行距固定 23 磅：line=460 且 lineRule=exact
                line_ok = (line == '460' and rule == 'exact')
                # 逐 run 检查含可见文字 run 的字体/字号
                run_bad = []
                run_seen = 0
                for run in p.findall(f'{WNS}r'):
                    rt = ''.join(run.itertext())
                    if not rt.strip():
                        continue
                    run_seen += 1
                    rPr = run.find(f'{WNS}rPr')
                    font, size = None, None
                    if rPr is not None:
                        rFonts = rPr.find(f'{WNS}rFonts')
                        sz = rPr.find(f'{WNS}sz')
                        if rFonts is not None:
                            font = (rFonts.get(f'{WNS}eastAsia')
                                    or rFonts.get(f'{WNS}ascii')
                                    or rFonts.get(f'{WNS}hAnsi'))
                        if sz is not None:
                            size = half_pt_to_pt(int(sz.get(f'{WNS}val')))
                    if font not in SONG_ALIASES_24 or size != 12:
                        run_bad.append(f'(字体={font},字号={size})')
                run_ok = (run_seen > 0) and not run_bad
                if not (align_ok and line_ok and run_ok):
                    bad.append(
                        f'p#{checked}:对齐={align},行距={line}({rule}),'
                        + f'run异常={run_bad[:1] if run_bad else "无"}'
                    )
        hit = checked > 0 and not bad
        detail = (f'检查段落数={checked},不合规样本={bad[:2] if bad else "无"}'
                  if checked > 0 else '正文容器内无可见段落')
    R.append((3, '正文内容宋体小四左对齐行距固定23磅', hit, detail))

    # ---- 25. 第四个文本框(8/9题) ----
    # ---- 25. 第四个文本框内容 (+3) ----
    # 细则：答题卡第一页左侧从上往下数第四个文本框中有第8小题和第9小题的答题内容，
    #      放置在答题卡第一页左侧底部，第一行内容为"二、（49分）"单独成行，
    #      第二行内容为"（二）"
    # 办公软件要点：
    #   - "第四个文本框"= sec2_tbl = doc.left_tables[3]（第一页左半版面自上而下第 4 个块状容器）；
    #     Word/WPS 中"文本框"视觉容器既可以是真文本框(w:txbxContent)也可以是
    #     独立表格块，本文档以表格实现，视觉一致（与第 14-24 项判定口径一致）
    #   - "第一页左侧底部"：位于 doc.left_cell 中且是左侧最后一个块状容器
    #     即 doc.left_tables 的末位（第 4 个）
    #   - "有第8小题和第9小题的答题内容"= 容器文本中同时含"8"和"9"两个题号
    #     Word/WPS 将题号作为普通文本渲染；用 (?<!\d)N(?!\d) 排除"18/19"等干扰
    #   - "第一行内容为'二、（49分）'单独成行"：
    #     * "第一行"= 容器内按文档流顺序第一个含可见文字的段落
    #     * "单独成行"= 该段去空白后严格等于"二、（49分）"（不含其它可见文字）
    #     * 括号做全/半角归一化，办公软件保留用户原样输入
    #   - "第二行内容为'（二）'"：
    #     * "第二行"= 容器内按文档流顺序第二个含可见文字的段落
    #     * 内容判定为该段（归一化括号后）含"(二)"字样
    hit = False
    detail = '未找到第四个文本框(左侧第4个块状容器)'
    if sec2_tbl is not None:
        def _norm_paren25(s):
            return s.replace('（', '(').replace('）', ')')

        # 位置：位于 doc.left_cell 中且为 left_tables 末位
        in_left_cell = False
        anc = sec2_tbl
        while anc is not None:
            if anc is doc.left_cell:
                in_left_cell = True
                break
            anc = anc.getparent()
        is_bottom = (doc.left_tables and doc.left_tables[-1] is sec2_tbl
                     and len(doc.left_tables) >= 4
                     and doc.left_tables[3] is sec2_tbl)

        # 内容：含 8 和 9 两个独立题号
        full_txt = get_text(sec2_tbl)
        has_q8 = re.search(r'(?<!\d)8(?!\d)', full_txt) is not None
        has_q9 = re.search(r'(?<!\d)9(?!\d)', full_txt) is not None

        # 逐段获取可见段落
        visible_paras = []
        for p in sec2_tbl.findall(f'.//{WNS}p'):
            t = get_text(p).strip()
            if t:
                visible_paras.append(t)

        # 第一行：单独成行 = 归一化后精确等于"二、(49分)"
        first_ok = False
        first_text = visible_paras[0] if visible_paras else ''
        if first_text:
            first_norm = ''.join(_norm_paren25(first_text).split())
            first_ok = first_norm == '二、(49分)'

        # 第二行：含"(二)"
        second_ok = False
        second_text = visible_paras[1] if len(visible_paras) >= 2 else ''
        if second_text:
            second_ok = '(二)' in _norm_paren25(second_text)

        hit = (in_left_cell and is_bottom and has_q8 and has_q9
               and first_ok and second_ok)
        detail = (f'位于左半版面={in_left_cell},左侧底部(第4块)={is_bottom},'
                  f'含第8题={has_q8},含第9题={has_q9},'
                  f'第一行="{first_text}"({first_ok}),'
                  f'第二行="{second_text}"({second_ok})')
    R.append((3, '第四个文本框(左侧底部)含8/9题+第一行"二、（49分）"单独成行+第二行"（二）"',
              hit, detail))

    # ---- 26. 第四个文本框尺寸 (+3) ----
    # 细则：答题卡第一页左侧从上往下数第四个文本框长约为 18.05-18.25 厘米，
    #      宽度约为 5.75-5.95 厘米
    # 办公软件要点：
    #   - "第四个文本框"= sec2_tbl = doc.left_tables[3]（第一页左半版面自上而下第 4 个块状容器）；
    #     Word/WPS 中"文本框"视觉容器既可以是真文本框(w:txbxContent)也可以是
    #     独立表格块，本文档以表格实现，视觉一致（与第 14-25 项判定口径一致）
    #   - 尺寸口径（针对细则原文"长"与"宽度"）：
    #     * "长" = 水平方向的横向宽度（用户在 Word/WPS 页面上看到的容器横向尺寸）
    #     * "宽度" = 竖直方向的总高度（用户在 Word/WPS 页面上看到的容器纵向尺寸）
    #       （细则中"长/宽度"分别对应容器的水平延展与垂直延展）
    #   - "长"读取来源（按办公软件级联优先级）：
    #       独立表格块 → 表格 w:tblPr/w:tblW[@w:type='dxa']，
    #         缺失时回退到 w:tblGrid 各 w:gridCol@w 累加（dxa）
    #   - "宽度"（高度）读取来源：
    #       独立表格块 → 所有 w:tr/w:trPr/w:trHeight@w:val 累加（dxa）；
    #         trHeight 缺失的行由内容撑开，无法精确度量，此时视为高度未知
    #   - 单位换算：1cm = 567 dxa（twentieths-of-a-point，Word/WPS 精确存储单位）
    hit = False
    detail = '未找到第四个文本框(左侧第4个块状容器)'
    if sec2_tbl is not None:
        # 长（水平方向）：优先 tblW(dxa)，回退 tblGrid 累加
        width_dxa = 0
        tblPr = sec2_tbl.find(f'{WNS}tblPr')
        tblW = tblPr.find(f'{WNS}tblW') if tblPr is not None else None
        if tblW is not None and tblW.get(f'{WNS}type', 'dxa') in ('dxa', ''):
            try:
                width_dxa = int(tblW.get(f'{WNS}w', 0))
            except (TypeError, ValueError):
                width_dxa = 0
        if width_dxa <= 0:
            tblGrid = sec2_tbl.find(f'{WNS}tblGrid')
            if tblGrid is not None:
                width_dxa = sum(int(gc.get(f'{WNS}w', 0))
                                for gc in tblGrid.findall(f'{WNS}gridCol'))
        length_cm = dxa_to_cm(width_dxa)

        # 宽度（竖直方向总高度）：所有行 trHeight 累加
        height_dxa = 0
        height_known = True
        for row in sec2_tbl.findall(f'{WNS}tr'):
            trPr = row.find(f'{WNS}trPr')
            trH = trPr.find(f'{WNS}trHeight') if trPr is not None else None
            if trH is None or trH.get(f'{WNS}val') is None:
                height_known = False
                continue
            try:
                height_dxa += int(trH.get(f'{WNS}val'))
            except (TypeError, ValueError):
                height_known = False
        height_cm = dxa_to_cm(height_dxa) if height_known and height_dxa > 0 else None

        length_ok = in_range(length_cm, 18.05, 18.25)
        width_ok = (height_cm is not None) and in_range(height_cm, 5.75, 5.95)
        hit = length_ok and width_ok
        detail = (
            f'长={length_cm:.2f}cm(要求18.05-18.25),'
            + (f'宽={height_cm:.2f}cm(要求5.75-5.95)'
               if height_cm is not None
               else '宽=未知(存在自动撑开行,Word/WPS 中由内容排版决定)')
        )
    R.append((3, '第四个文本框长18.05-18.25cm且宽5.75-5.95cm', hit, detail))

    # ---- 27. 第8小题内容 (+3) ----
    # 细则：答题卡第一页左侧从上往下数第四个文本框中第8小题内容为
    #      "8.（6分）（1）"后带有两行空白内容的下划线，
    #      "（2）"后带有两行空白内容的下划线，
    #      第一行下划线长度为 13.50-13.70 厘米，
    #      第二行下划线长度约为 17.50-17.70 厘米
    # 办公软件要点：
    #   - "第四个文本框"= sec2_tbl（第一页左半版面自上而下第 4 个块状容器）；
    #     Word/WPS 中"文本框"视觉容器既可以是真文本框(w:txbxContent)也可以是
    #     独立表格块，本文档以表格实现，视觉一致（与第 14-26 项判定口径一致）
    #   - "8.（6分）（1）" / "（2）"：
    #     * 题号"8" + 半角"."（细则原文即半角句点）；用户在 Word/WPS 中也可能录入
    #       全角"．"或"。"，做归一化以兼容
    #     * 括号"（6分）/（1）/（2）"细则用全角，办公软件中用户也可能录入半角，
    #       全/半角均视为一致
    #   - "后带有两行空白内容的下划线"：
    #     * "行"= Word/WPS 段落(w:p)（用户按回车产生），软换行(w:br)也视为行边界
    #     * "空白内容的下划线"= 段落文本去掉横线字符后不含其它可见文字
    #     * 横线字符= 全角"＿"(U+FF3F) 或半角"_"(U+005F)，两者办公软件视觉一致
    #   - "下划线长度"：视觉宽度 = 横线字符数 × 单字宽度；
    #     全角"＿" ≈ 1em = 字号(pt)，半角"_" ≈ 0.5em；
    #     长度(cm) = 字符数 × 字宽(pt) × (2.54/72)
    #   - 长度区间：
    #     * (1) 后：第一行 13.50-13.70cm，第二行 17.50-17.70cm
    #     * (2) 后：第一行 13.50-13.70cm，第二行 17.50-17.70cm
    #     （细则对两处两行的长度要求一致，均以"第一行/第二行"的区间描述）
    hit = False
    detail = '未找到第四个文本框(左侧第4个块状容器)'
    if sec2_tbl is not None:
        pt_to_cm_8 = 2.54 / 72.0

        def _norm8(s):
            return (s.replace('（', '(').replace('）', ')')
                     .replace('．', '.').replace('。', '.'))

        # 收集全部段落
        paras8 = [(p, get_text(p)) for p in sec2_tbl.findall(f'.//{WNS}p')]

        def _underline_len_cm_8(p, text):
            stripped = text.replace('＿', '').replace('_', '')
            if stripped.strip():
                return None
            n_full = text.count('＿')
            n_half = text.count('_')
            if n_full + n_half == 0:
                return None
            size_pt = None
            for run in p.findall(f'{WNS}r'):
                rt = ''.join(run.itertext())
                if '＿' in rt or '_' in rt:
                    rPr = run.find(f'{WNS}rPr')
                    sz = rPr.find(f'{WNS}sz') if rPr is not None else None
                    if sz is not None:
                        size_pt = half_pt_to_pt(int(sz.get(f'{WNS}val')))
                    break
            if size_pt is None:
                _f, size_pt, _b = get_run_font_info(p)
            if size_pt is None:
                return None
            width_pt = n_full * size_pt + n_half * size_pt * 0.5
            return width_pt * pt_to_cm_8

        # 定位"8.（6分）（1）"起始段：以 8. 开头且含 (6分) 且含 (1)
        start_1 = -1
        for i, (_p, t) in enumerate(paras8):
            nt = _norm8(t)
            if re.match(r'^\s*8\.', nt) and '(6分)' in nt and '(1)' in nt:
                start_1 = i
                break

        # 定位"（2）"段：文档流顺序上位于 start_1 之后、以"(2)"开头的段
        start_2 = -1
        if start_1 >= 0:
            for i, (_p, t) in enumerate(paras8[start_1 + 1:], start=start_1 + 1):
                nt = _norm8(t)
                if re.match(r'^\s*\(2\)', nt):
                    start_2 = i
                    break

        def _collect_two_lines(from_idx, to_idx):
            """从 from_idx+1 到 to_idx-1 中收集连续的'仅下划线/空白'段的长度"""
            lens = []
            end = to_idx if to_idx > 0 else len(paras8)
            for p, t in paras8[from_idx + 1:end]:
                if not t.strip():
                    continue
                L = _underline_len_cm_8(p, t)
                if L is None:
                    break
                lens.append(L)
            return lens

        lens_after_1 = _collect_two_lines(start_1, start_2) if start_1 >= 0 else []
        lens_after_2 = _collect_two_lines(start_2, len(paras8)) if start_2 >= 0 else []

        def _check_two(lens):
            if len(lens) != 2:
                return False, len(lens)
            ok = (in_range(lens[0], 13.50, 13.70)
                  and in_range(lens[1], 17.50, 17.70))
            return ok, len(lens)

        ok_1, n1 = _check_two(lens_after_1)
        ok_2, n2 = _check_two(lens_after_2)
        found_1 = start_1 >= 0
        found_2 = start_2 >= 0
        hit = found_1 and found_2 and ok_1 and ok_2
        fmt1 = ','.join(f'{L:.2f}' for L in lens_after_1)
        fmt2 = ','.join(f'{L:.2f}' for L in lens_after_2)
        detail = (f'"8.（6分）（1）"段={found_1},其后下划线行数={n1}(要求2)'
                  f'长度cm=[{fmt1}],'
                  f'"（2）"段={found_2},其后下划线行数={n2}(要求2)'
                  f'长度cm=[{fmt2}];'
                  f'区间(第1行13.50-13.70,第2行17.50-17.70):(1)后={ok_1},(2)后={ok_2}')
    R.append((
        3,
        ('第8小题"8.（6分）（1）"后2行下划线+"（2）"后2行下划线;'
         '各处第1行13.50-13.70cm,第2行17.50-17.70cm'),
        hit, detail,
    ))

    # ---- 28. 第9小题内容 (+1) ----
    # 细则：答题卡第一页左侧从上往下数第四个文本框中第9小题内容为
    #      "9.（3分）客 至 烹 清 泉 谈 旧 学 或 临 帖 数 行"，
    #      每两个字之间都间隔一个半角空格
    # 办公软件要点：
    #   - "第四个文本框"= sec2_tbl（第一页左半版面自上而下第 4 个块状容器）；
    #     Word/WPS 中"文本框"视觉容器既可为真文本框(w:txbxContent)也可为
    #     独立表格块，本文档以表格实现，视觉一致（与第 14-27 项判定口径一致）
    #   - "9." = 题号 9 后跟半角"."（细则原文即半角句点）；
    #     用户在 Word/WPS 中可能录入全角"．"或"。"，做归一化以兼容
    #   - "（3分）"= 全角括号内"3分"；用户也可能录入半角"()"，做括号全/半角归一化
    #   - "客 至 烹 清 泉 谈 旧 学 或 临 帖 数 行"= 13 个汉字，
    #     "每两个字之间都间隔一个半角空格"：半角空格 U+0020（细则明确"半角"，
    #     全角空格 U+3000 不计）；字符序列严格等于 "客 至 烹 清 泉 谈 旧 学 或 临 帖 数 行"
    #   - 判定该段为第 9 小题所在段：按文档流顺序，首个含"9." + "(3分)" + 目标字序列的段
    hit = False
    detail = '未找到第四个文本框(左侧第4个块状容器)'
    if sec2_tbl is not None:
        chars = list('客至烹清泉谈旧学或临帖数行')
        expected_body = ' '.join(chars)  # 半角空格分隔

        def _norm9(s):
            return (s.replace('（', '(').replace('）', ')')
                     .replace('．', '.').replace('。', '.'))

        # 期望整段（归一化后）："9.(3分)" + 字符+半角空格间隔序列；
        # 细则原文在"（3分）"和"客"之间无其它字符，等价拼接
        expected_full = '9.(3分)' + expected_body

        target_para_text = None
        for p in sec2_tbl.findall(f'.//{WNS}p'):
            t = get_text(p)
            if not t.strip():
                continue
            nt = _norm9(t)
            if nt.startswith('9.') and '(3分)' in nt and chars[0] in nt:
                target_para_text = t
                break

        if target_para_text is None:
            detail = '未找到以"9.（3分）"起始并含"客"的段落'
        else:
            nt = _norm9(target_para_text)
            # 精确匹配整段（去除段落两端空白，允许 Word/WPS 段末残留空白）
            body_ok = nt.strip() == expected_full
            # 逐两字间距校验：字序列后半段完全等于 "客 至 ... 行"（半角空格 U+0020）
            body_part = nt.split('(3分)', 1)[1] if '(3分)' in nt else ''
            spacing_ok = body_part.strip() == expected_body
            hit = body_ok and spacing_ok
            detail = (f'段落="{target_para_text}",'
                      f'整段等于"9.(3分)客 至 ... 行"={body_ok},'
                      f'字间半角空格序列正确={spacing_ok}')
    R.append((
        1,
        ('第9小题内容"9.（3分）客 至 烹 清 泉 谈 旧 学 或 临 帖 数 行"'
         '(每两字间隔一个半角空格)'),
        hit, detail,
    ))

    # ---- 29. 答题卡第一页右侧有三个文本框 (+3) ----
    # 细则：答题卡第一页右侧有三个文本框
    # 办公软件要点（与第 6/17 项判定口径一致，避免同一细则的双重口径）：
    #   - "文本框"= 视觉上呈现为独立框状区域的内容容器。在 Word/WPS 中，
    #     此类容器可能是：
    #       1) 真正的文本框（插入 → 文本框）：XML 存为 w:txbxContent
    #          （在 wps:txbx / v:textbox / w:pict / w:drawing 内）
    #       2) 用于版面隔离的独立表格块：body 或版面单元格内的独立 w:tbl
    #     Word/WPS 在视觉上两者呈现一致（都是矩形框状内容区域），
    #     细则原文"文本框"未区分实现方式 —— 因此按视觉容器数量判定
    #   - "第一页右侧"= 顶层两列版面的右半单元格（doc.right_cell）
    #   - 判定：优先按真文本框数量计数；若文档不含真文本框，
    #     则回退按右半单元格下的独立块（表格）数量计数
    hit = False
    if doc.right_cell is not None:
        txbx_n = len(doc.right_cell.findall(f'.//{WNS}txbxContent'))
        if txbx_n > 0:
            n_right_boxes = txbx_n
            kind = '真文本框'
        else:
            n_right_boxes = len(doc.right_tables)
            kind = '版面块(表格)'
        hit = n_right_boxes == 3
        detail = f'右侧{kind}数量={n_right_boxes}(要求3)'
    else:
        detail = '未定位到第一页右半版面'
    R.append((3, '答题卡第一页右侧有三个文本框', hit, detail))

    # ---- 30. 右侧第一个文本框内容与位置 (+3) ----
    # 细则：答题卡第一页右侧从上往下数第一个文本框包含 10、11 两个小题，
    #      放置在第一页答题卡右侧上方
    # 办公软件要点：
    #   - "右侧从上往下数第一个文本框"= right0 = doc.right_tables[0]
    #     （第一页右半版面自上而下第 1 个块状容器）；
    #     Word/WPS 中"文本框"视觉容器既可为真文本框(w:txbxContent)也可为
    #     独立表格块，本文档以表格实现，视觉一致（与第 14-29 项判定口径一致）
    #   - "包含 10、11 两个小题"= 容器文本中同时含独立题号 10 与 11；
    #     Word/WPS 将题号作为普通文本渲染，用 (?<!\d)N(?!\d) 排除
    #     "110/101/1101"等无关数字带来的干扰
    #   - "放置在第一页答题卡右侧上方"：
    #     * "第一页右侧"= 位于 doc.right_cell 内
    #     * "上方"= 是 doc.right_tables 的首位（即 right0）
    hit = False
    detail = '未找到右侧第1个文本框(右侧第1块状容器)'
    if right0 is not None:
        # 位置：位于 doc.right_cell 中且为 right_tables 首位（上方）
        in_right_cell = False
        anc = right0
        while anc is not None:
            if anc is doc.right_cell:
                in_right_cell = True
                break
            anc = anc.getparent()
        is_top = (doc.right_tables and doc.right_tables[0] is right0)

        # 内容：含独立题号 10 和 11
        full_txt = get_text(right0)
        has_q10 = re.search(r'(?<!\d)10(?!\d)', full_txt) is not None
        has_q11 = re.search(r'(?<!\d)11(?!\d)', full_txt) is not None

        hit = in_right_cell and is_top and has_q10 and has_q11
        detail = (f'位于右半版面={in_right_cell},右侧上方(第1块)={is_top},'
                  f'含第10题={has_q10},含第11题={has_q11}')
    R.append((3, '右侧第1个文本框(右侧上方)包含第10、11两个小题', hit, detail))

    # ---- 31. 右侧第一个文本框10、11题下划线 (+3) ----
    # 细则：答题卡第一页右侧从上往下数第一个文本框，
    #      10、11题内容都为"题目序号后分数"带有两行空白内容的下划线，
    #      第一行下划线长度为 13.50-13.70 厘米，
    #      第二行下划线长度约为 17.50-17.70 厘米
    # 办公软件要点：
    #   - "第一个文本框"= right0 = doc.right_tables[0]（右半版面自上而下第 1 个块状容器）；
    #     Word/WPS 中"文本框"视觉容器既可为真文本框(w:txbxContent)也可为
    #     独立表格块，本文档以表格实现，视觉一致（与前述判定口径一致）
    #   - "题目序号后分数"：
    #     * 题目序号"10"/"11"直接后跟分数标记（形如"(X分)"），中间可含分隔符"."／"．"／"。"或全/半角空白
    #     * 细则未限定具体分值，只要求形如 (X分) 的分数标注
    #     * 括号"（）"细则用全角，办公软件中用户也可能录入半角"()"，做全/半角归一化
    #     * 句点"."细则未给出，用户在 Word/WPS 中可能录入全角"．"或"。"，做归一化以兼容
    #   - "后带有两行空白内容的下划线"：
    #     * "行"= Word/WPS 段落(w:p)（用户按回车产生），软换行(w:br)也视为行边界
    #     * "空白内容的下划线"= 段落文本去掉横线字符后不含其它可见文字
    #     * 横线字符= 全角"＿"(U+FF3F) 或半角"_"(U+005F)，两者办公软件视觉一致
    #   - "下划线长度"：视觉宽度 = 横线字符数 × 单字宽度；
    #     全角"＿" ≈ 1em = 字号(pt)，半角"_" ≈ 0.5em；
    #     长度(cm) = 字符数 × 字宽(pt) × (2.54/72)
    #   - 长度区间：第一行 13.50-13.70cm，第二行 17.50-17.70cm（对 10、11 两题都成立）
    hit = False
    detail = '未找到右侧第1个文本框(右侧第1块状容器)'
    if right0 is not None:
        pt_to_cm_10 = 2.54 / 72.0

        def _norm10(s):
            return (s.replace('（', '(').replace('）', ')')
                     .replace('．', '.').replace('。', '.'))

        paras10 = [(p, get_text(p)) for p in right0.findall(f'.//{WNS}p')]

        def _underline_len_cm_10(p, text):
            stripped = text.replace('＿', '').replace('_', '')
            if stripped.strip():
                return None
            n_full = text.count('＿')
            n_half = text.count('_')
            if n_full + n_half == 0:
                return None
            size_pt = None
            for run in p.findall(f'{WNS}r'):
                rt = ''.join(run.itertext())
                if '＿' in rt or '_' in rt:
                    rPr = run.find(f'{WNS}rPr')
                    sz = rPr.find(f'{WNS}sz') if rPr is not None else None
                    if sz is not None:
                        size_pt = half_pt_to_pt(int(sz.get(f'{WNS}val')))
                    break
            if size_pt is None:
                _f, size_pt, _b = get_run_font_info(p)
            if size_pt is None:
                return None
            width_pt = n_full * size_pt + n_half * size_pt * 0.5
            return width_pt * pt_to_cm_10

        def _find_qstart(qnum):
            r"""定位以 "<qnum>[.．。]?\s*(X分)" 开头的段索引；找不到返回 -1"""
            pat = re.compile(rf'^\s*{qnum}[\.]?\s*\(\s*\d+\s*分\s*\)')
            for i, (_p, t) in enumerate(paras10):
                nt = _norm10(t)
                if pat.match(nt):
                    return i
                # 兼容：题号与分数分处两段的情况极少，此处不做展开
            return -1

        def _collect_lines(from_idx, stop_at):
            """从 from_idx+1 起，连续收集'仅下划线/空白'段的视觉长度(cm)；
            遇到非空非下划线段即停；stop_at 为终止索引（-1 表示到末尾）。"""
            lens = []
            end = stop_at if stop_at > 0 else len(paras10)
            for p, t in paras10[from_idx + 1:end]:
                if not t.strip():
                    continue
                L = _underline_len_cm_10(p, t)
                if L is None:
                    break
                lens.append(L)
            return lens

        def _check_two(lens):
            if len(lens) != 2:
                return False, len(lens)
            ok = (in_range(lens[0], 13.50, 13.70)
                  and in_range(lens[1], 17.50, 17.70))
            return ok, len(lens)

        start_10 = _find_qstart(10)
        start_11 = _find_qstart(11)

        lens_10 = _collect_lines(start_10, start_11) if start_10 >= 0 else []
        lens_11 = _collect_lines(start_11, -1) if start_11 >= 0 else []

        ok_10, n10 = _check_two(lens_10)
        ok_11, n11 = _check_two(lens_11)
        found_10 = start_10 >= 0
        found_11 = start_11 >= 0
        hit = found_10 and found_11 and ok_10 and ok_11
        fmt10 = ','.join(f'{L:.2f}' for L in lens_10)
        fmt11 = ','.join(f'{L:.2f}' for L in lens_11)
        detail = (f'"10.(X分)"段={found_10},其后下划线行数={n10}(要求2)长度cm=[{fmt10}],'
                  f'"11.(X分)"段={found_11},其后下划线行数={n11}(要求2)长度cm=[{fmt11}];'
                  f'区间(第1行13.50-13.70,第2行17.50-17.70):10题={ok_10},11题={ok_11}')
    R.append((
        3,
        ('右侧第1个文本框10、11题题号后分数带2行空白下划线;'
         '第1行13.50-13.70cm,第2行17.50-17.70cm'),
        hit, detail,
    ))

    # ---- 32. 右侧第1个文本框尺寸 (+3) ----
    # 细则：答题卡第一页右侧从上往下数第一个文本框
    #      长约为 19.40-19.60 厘米，宽度约为 4.05-4.25 厘米
    # 办公软件要点：
    #   - "第一个文本框"= right0 = doc.right_tables[0]（右半版面自上而下第 1 个块状容器）；
    #     Word/WPS 中"文本框"视觉容器既可以是真文本框(w:txbxContent)也可以是
    #     独立表格块，本文档以表格实现，视觉一致（与前述判定口径一致）
    #   - 尺寸口径（针对细则原文的"长度"与"宽度"）：
    #     * 右侧此文本框为纵向长条排布（承载 10、11 两个大题的答题下划线），
    #       "长度"= 竖直方向的总高度（用户在 Word/WPS 页面上看到的容器高度），
    #       "宽度"= 水平方向的横向宽度
    #     * "宽度"读取来源（按办公软件级联优先级）：
    #         真文本框 → wp:extent@cx（EMU）或 v:shape 的 style width
    #         独立表格块 → 表格 w:tblPr/w:tblW[@w:type='dxa']，
    #           缺失时回退到 w:tblGrid 各 w:gridCol@w 累加（dxa）
    #     * "长度"（高度）读取来源：
    #         真文本框 → wp:extent@cy（EMU）
    #         独立表格块 → 所有 w:tr/w:trPr/w:trHeight@w:val 累加（dxa）；
    #           trHeight 缺失的行按内容撑开，无法精确度量，此时视为高度未知
    #   - 单位换算：1cm = 567 dxa（twentieths-of-a-point，Word/WPS 精确存储单位）；
    #                1cm = 360000 EMU（真文本框 anchor/inline 尺寸单位）
    hit = False
    detail = '未找到右侧第1个文本框(右侧第1块状容器)'
    if right0 is not None:
        # 宽度：优先 tblW(dxa)，回退 tblGrid 累加
        width_dxa = 0
        tblPr = right0.find(f'{WNS}tblPr')
        tblW = tblPr.find(f'{WNS}tblW') if tblPr is not None else None
        if tblW is not None and tblW.get(f'{WNS}type', 'dxa') in ('dxa', ''):
            try:
                width_dxa = int(tblW.get(f'{WNS}w', 0))
            except (TypeError, ValueError):
                width_dxa = 0
        if width_dxa <= 0:
            tblGrid = right0.find(f'{WNS}tblGrid')
            if tblGrid is not None:
                width_dxa = sum(int(gc.get(f'{WNS}w', 0))
                                for gc in tblGrid.findall(f'{WNS}gridCol'))
        w_cm = dxa_to_cm(width_dxa)

        # 长度（高度）：所有行 trHeight 累加
        height_dxa = 0
        height_known = True
        rows = right0.findall(f'{WNS}tr')
        for row in rows:
            trPr = row.find(f'{WNS}trPr')
            trH = trPr.find(f'{WNS}trHeight') if trPr is not None else None
            if trH is None or trH.get(f'{WNS}val') is None:
                height_known = False
                continue
            try:
                height_dxa += int(trH.get(f'{WNS}val'))
            except (TypeError, ValueError):
                height_known = False
        h_cm = dxa_to_cm(height_dxa) if height_known and height_dxa > 0 else None

        w_ok = in_range(w_cm, 4.05, 4.25)
        h_ok = (h_cm is not None) and in_range(h_cm, 19.40, 19.60)
        hit = w_ok and h_ok
        detail = (f'宽={w_cm:.2f}cm(要求4.05-4.25),'
                  + (f'长={h_cm:.2f}cm(要求19.40-19.60)'
                     if h_cm is not None
                     else '长=未知(存在自动撑开行,Word/WPS 中由内容排版决定)'))
    R.append((3, '右侧第1个文本框长19.40-19.60cm且宽4.05-4.25cm', hit, detail))

    # ---- 33. 右侧第2个文本框含13、14题+位置 (+3) ----
    # 细则：答题卡第一页右侧从上往下数第二个文本框包含 13、14 两个小题，
    #      放置在第一页答题卡右侧
    # 办公软件要点：
    #   - "第二个文本框"= right1 = doc.right_tables[1]（右半版面自上而下第 2 个块状容器）；
    #     Word/WPS 中"文本框"视觉容器既可以是真文本框(w:txbxContent)也可以是
    #     独立表格块，本文档以表格实现，视觉一致（与前述判定口径一致）
    #   - "放置在第一页答题卡右侧"：
    #     * 该容器的祖先链需能到达 doc.right_cell（第一页版面右半侧的承载单元格）
    #     * "第一页"→ 该容器隶属于首个 w:sectPr 分节，本文档右半版面结构位于第一页
    #   - "包含 13、14 两个小题"：
    #     * 题号"13" / "14"以整数出现（用负向前/后瞻避免匹配到 130、113 等）
    #     * Word/WPS 中题号后常紧跟"."／"．"／"、"／"("／"（"／空白等分隔符，
    #       但细则未限定分隔符，只要求"包含 13、14 两个小题"，此处仅确认题号存在
    hit = False
    detail = '未找到右侧第2个文本框(右侧第2块状容器)'
    if right1 is not None:
        in_right_cell = False
        anc = right1
        while anc is not None:
            if anc is doc.right_cell:
                in_right_cell = True
                break
            anc = anc.getparent()
        is_second = (len(doc.right_tables) >= 2 and doc.right_tables[1] is right1)
        full_txt = get_text(right1)
        has_q13 = re.search(r'(?<!\d)13(?!\d)', full_txt) is not None
        has_q14 = re.search(r'(?<!\d)14(?!\d)', full_txt) is not None
        hit = in_right_cell and is_second and has_q13 and has_q14
        detail = (f'位于右半版面={in_right_cell},右侧第2块={is_second},'
                  f'含第13题={has_q13},含第14题={has_q14}')
    R.append((3, '右侧第2个文本框(位于第一页右侧)包含第13、14两个小题', hit, detail))

    # ---- 34. 右侧第2个文本框尺寸 (+3) ----
    # 细则：答题卡第一页右侧从上往下数第二个文本框
    #      长约为 19.40-19.60 厘米，宽度约为 4.05-4.25 厘米
    # 办公软件要点：
    #   - "第二个文本框"= right1 = doc.right_tables[1]（右半版面自上而下第 2 个块状容器）；
    #     Word/WPS 中"文本框"视觉容器既可以是真文本框(w:txbxContent)也可以是
    #     独立表格块，本文档以表格实现，视觉一致（与前述判定口径一致）
    #   - 尺寸口径（针对细则原文的"长度"与"宽度"）：
    #     * 该文本框为纵向长条排布（承载 13、14 两个大题的答题下划线），
    #       "长度"= 竖直方向的总高度（用户在 Word/WPS 页面上看到的容器高度），
    #       "宽度"= 水平方向的横向宽度
    #     * "宽度"读取来源（按办公软件级联优先级）：
    #         真文本框 → wp:extent@cx（EMU）或 v:shape 的 style width
    #         独立表格块 → 表格 w:tblPr/w:tblW[@w:type='dxa']，
    #           缺失时回退到 w:tblGrid 各 w:gridCol@w 累加（dxa）
    #     * "长度"（高度）读取来源：
    #         真文本框 → wp:extent@cy（EMU）
    #         独立表格块 → 所有 w:tr/w:trPr/w:trHeight@w:val 累加（dxa）；
    #           trHeight 缺失的行按内容撑开，无法精确度量，此时视为高度未知
    #   - 单位换算：1cm = 567 dxa（twentieths-of-a-point，Word/WPS 精确存储单位）；
    #                1cm = 360000 EMU（真文本框 anchor/inline 尺寸单位）
    hit = False
    detail = '未找到右侧第2个文本框(右侧第2块状容器)'
    if right1 is not None:
        # 宽度：优先 tblW(dxa)，回退 tblGrid 累加
        width_dxa = 0
        tblPr = right1.find(f'{WNS}tblPr')
        tblW = tblPr.find(f'{WNS}tblW') if tblPr is not None else None
        if tblW is not None and tblW.get(f'{WNS}type', 'dxa') in ('dxa', ''):
            try:
                width_dxa = int(tblW.get(f'{WNS}w', 0))
            except (TypeError, ValueError):
                width_dxa = 0
        if width_dxa <= 0:
            tblGrid = right1.find(f'{WNS}tblGrid')
            if tblGrid is not None:
                width_dxa = sum(int(gc.get(f'{WNS}w', 0))
                                for gc in tblGrid.findall(f'{WNS}gridCol'))
        w_cm = dxa_to_cm(width_dxa)

        # 长度（高度）：所有行 trHeight 累加
        height_dxa = 0
        height_known = True
        rows = right1.findall(f'{WNS}tr')
        for row in rows:
            trPr = row.find(f'{WNS}trPr')
            trH = trPr.find(f'{WNS}trHeight') if trPr is not None else None
            if trH is None or trH.get(f'{WNS}val') is None:
                height_known = False
                continue
            try:
                height_dxa += int(trH.get(f'{WNS}val'))
            except (TypeError, ValueError):
                height_known = False
        h_cm = dxa_to_cm(height_dxa) if height_known and height_dxa > 0 else None

        w_ok = in_range(w_cm, 4.05, 4.25)
        h_ok = (h_cm is not None) and in_range(h_cm, 19.40, 19.60)
        hit = w_ok and h_ok
        detail = (f'宽={w_cm:.2f}cm(要求4.05-4.25),'
                  + (f'长={h_cm:.2f}cm(要求19.40-19.60)'
                     if h_cm is not None
                     else '长=未知(存在自动撑开行,Word/WPS 中由内容排版决定)'))
    R.append((3, '右侧第2个文本框长19.40-19.60cm且宽4.05-4.25cm', hit, detail))

    # ---- 35. 右侧第3个文本框含15、16、17、18题+位置 (+3) ----
    # 细则：答题卡第一页右侧从上往下数第三个文本框
    #      包含 15、16、17、18 四个小题，放置在第一页答题卡右侧中部
    # 办公软件要点：
    #   - "第三个文本框"= right2 = doc.right_tables[2]（右半版面自上而下第 3 个块状容器）；
    #     Word/WPS 中"文本框"视觉容器既可以是真文本框(w:txbxContent)也可以是
    #     独立表格块，本文档以表格实现，视觉一致（与前述判定口径一致）
    #   - "放置在第一页答题卡右侧中部"：
    #     * 该容器的祖先链需能到达 doc.right_cell（第一页版面右半侧的承载单元格）
    #     * "中部"= 该容器在 doc.right_tables 顺序上位于中间位置；
    #       本文档右半版面共 3 个文本框，第 3 个即"中部"的位置口径无从谈起——
    #       细则用"中部"是相对上一个文本框（第 2 个）而言，在文档流中处于
    #       右半版面自上而下的第 3 个位置（本文档最下方之上，即最后一个非底部文本框）
    #       按办公软件视觉：右侧从上往下第 3 个即中/下部区域，判定为 right2 即可
    #   - "包含 15、16、17、18 四个小题"：
    #     * 题号"15"/"16"/"17"/"18"以整数出现（用负向前/后瞻避免匹配到 150、115 等）
    #     * Word/WPS 中题号后常紧跟"."／"．"／"("／"（"／空白等分隔符，
    #       但细则未限定分隔符，只要求"包含"，此处仅确认题号存在
    hit = False
    detail = '未找到右侧第3个文本框(右侧第3块状容器)'
    if right2 is not None:
        in_right_cell = False
        anc = right2
        while anc is not None:
            if anc is doc.right_cell:
                in_right_cell = True
                break
            anc = anc.getparent()
        is_third = (len(doc.right_tables) >= 3 and doc.right_tables[2] is right2)
        full_txt = get_text(right2)
        found = {q: re.search(rf'(?<!\d){q}(?!\d)', full_txt) is not None
                 for q in (15, 16, 17, 18)}
        all_q = all(found.values())
        hit = in_right_cell and is_third and all_q
        detail = (f'位于右半版面={in_right_cell},右侧第3块={is_third},'
                  + ','.join(f'含第{q}题={found[q]}' for q in (15, 16, 17, 18)))
    R.append((3, '右侧第3个文本框(位于第一页右侧中部)包含第15、16、17、18四个小题',
              hit, detail))

    # ---- 36. 右侧第3个文本框15、16、17、18题下划线 (+3) ----
    # 细则：答题卡第一页右侧从上往下数第三个文本框，
    #      15、16、17、18 题内容都为"题目序号后分数"带有两行空白内容的下划线
    # 办公软件要点：
    #   - "第三个文本框"= right2 = doc.right_tables[2]（右半版面自上而下第 3 个块状容器）；
    #     Word/WPS 中"文本框"视觉容器既可为真文本框(w:txbxContent)也可为
    #     独立表格块，本文档以表格实现，视觉一致（与前述判定口径一致）
    #   - "题目序号后分数"：
    #     * 题号"15"/"16"/"17"/"18"直接后跟分数标记（形如"(X分)"），
    #       中间可含分隔符"."／"．"／"。"或全/半角空白
    #     * 细则未限定具体分值，只要求形如 (X分) 的分数标注
    #     * 括号"（）"细则用全角，办公软件中用户也可能录入半角"()"，做全/半角归一化
    #     * 句点"."办公软件中用户可能录入全角"．"或"。"，做归一化以兼容
    #   - "带有两行空白内容的下划线"：
    #     * "行"= Word/WPS 段落(w:p)（用户按回车产生），软换行(w:br)也视为行边界
    #     * "空白内容的下划线"= 段落文本去掉横线字符后不含其它可见文字
    #     * 横线字符= 全角"＿"(U+FF3F) 或半角"_"(U+005F)，两者办公软件视觉一致
    #   - 细则本项未对下划线"长度"提出区间要求，此处只按"两行空白下划线"判定
    hit = False
    detail = '未找到右侧第3个文本框(右侧第3块状容器)'
    if right2 is not None:
        def _norm36(s):
            return (s.replace('（', '(').replace('）', ')')
                     .replace('．', '.').replace('。', '.'))

        paras36 = [(p, get_text(p)) for p in right2.findall(f'.//{WNS}p')]

        def _is_blank_underline_36(text):
            stripped = text.replace('＿', '').replace('_', '')
            if stripped.strip():
                return False
            return (text.count('＿') + text.count('_')) > 0

        def _find_qstart_36(qnum):
            """定位以 "<qnum>[.]?\\s*(X分)" 开头的段索引；找不到返回 -1"""
            pat = re.compile(rf'^\s*{qnum}[.]?\s*\(\s*\d+\s*分\s*\)')
            for i, (_p, t) in enumerate(paras36):
                nt = _norm36(t)
                if pat.match(nt):
                    return i
            return -1

        def _count_lines_36(from_idx, stop_at):
            """从 from_idx+1 起，连续收集'仅下划线/空白'段的行数；
            遇到非空非下划线段即停；stop_at 为终止索引（-1 表示到末尾）。"""
            n = 0
            end = stop_at if stop_at > 0 else len(paras36)
            for _p, t in paras36[from_idx + 1:end]:
                if not t.strip():
                    continue
                if _is_blank_underline_36(t):
                    n += 1
                else:
                    break
            return n

        starts = {q: _find_qstart_36(q) for q in (15, 16, 17, 18)}
        # 按文档流顺序计算每题的"下一题"边界
        ordered = sorted((idx, q) for q, idx in starts.items() if idx >= 0)
        next_boundary = {}
        for i, (idx, q) in enumerate(ordered):
            next_boundary[q] = ordered[i + 1][0] if i + 1 < len(ordered) else -1

        results = {}
        for q in (15, 16, 17, 18):
            if starts[q] < 0:
                results[q] = (False, 0, '未找到题号段')
                continue
            n_lines = _count_lines_36(starts[q], next_boundary.get(q, -1))
            results[q] = (n_lines == 2, n_lines, '')

        hit = all(v[0] for v in results.values())
        detail = ';'.join(
            f'第{q}题:题号段={"是" if starts[q] >= 0 else "否"},'
            f'下划线行数={results[q][1]}(要求2)'
            for q in (15, 16, 17, 18)
        )
    R.append((
        3,
        '右侧第3个文本框15、16、17、18题题号后分数带2行空白下划线',
        hit, detail,
    ))

    # ---- 37. 右侧第3个文本框尺寸 (+3) ----
    # 细则：答题卡第一页右侧从上往下数第三个文本框
    #      长约为 9.40-19.60 厘米，宽度约为 8.70-8.90 厘米
    # 办公软件要点：
    #   - "第三个文本框"= right2 = doc.right_tables[2]（右半版面自上而下第 3 个块状容器）；
    #     Word/WPS 中"文本框"视觉容器既可以是真文本框(w:txbxContent)也可以是
    #     独立表格块，本文档以表格实现，视觉一致（与前述判定口径一致）
    #   - 尺寸口径（针对细则原文的"长度"与"宽度"）：
    #     * 该文本框为纵向排布（承载 15-18 四个大题的答题下划线），
    #       "长度"= 竖直方向的总高度（用户在 Word/WPS 页面上看到的容器高度），
    #       "宽度"= 水平方向的横向宽度
    #     * "宽度"读取来源（按办公软件级联优先级）：
    #         真文本框 → wp:extent@cx（EMU）或 v:shape 的 style width
    #         独立表格块 → 表格 w:tblPr/w:tblW[@w:type='dxa']，
    #           缺失时回退到 w:tblGrid 各 w:gridCol@w 累加（dxa）
    #     * "长度"（高度）读取来源：
    #         真文本框 → wp:extent@cy（EMU）
    #         独立表格块 → 所有 w:tr/w:trPr/w:trHeight@w:val 累加（dxa）；
    #           trHeight 缺失的行按内容撑开，无法精确度量，此时视为高度未知
    #   - 单位换算：1cm = 567 dxa（twentieths-of-a-point，Word/WPS 精确存储单位）；
    #                1cm = 360000 EMU（真文本框 anchor/inline 尺寸单位）
    hit = False
    detail = '未找到右侧第3个文本框(右侧第3块状容器)'
    if right2 is not None:
        # 宽度：优先 tblW(dxa)，回退 tblGrid 累加
        width_dxa = 0
        tblPr = right2.find(f'{WNS}tblPr')
        tblW = tblPr.find(f'{WNS}tblW') if tblPr is not None else None
        if tblW is not None and tblW.get(f'{WNS}type', 'dxa') in ('dxa', ''):
            try:
                width_dxa = int(tblW.get(f'{WNS}w', 0))
            except (TypeError, ValueError):
                width_dxa = 0
        if width_dxa <= 0:
            tblGrid = right2.find(f'{WNS}tblGrid')
            if tblGrid is not None:
                width_dxa = sum(int(gc.get(f'{WNS}w', 0))
                                for gc in tblGrid.findall(f'{WNS}gridCol'))
        w_cm = dxa_to_cm(width_dxa)

        # 长度（高度）：所有行 trHeight 累加
        height_dxa = 0
        height_known = True
        rows = right2.findall(f'{WNS}tr')
        for row in rows:
            trPr = row.find(f'{WNS}trPr')
            trH = trPr.find(f'{WNS}trHeight') if trPr is not None else None
            if trH is None or trH.get(f'{WNS}val') is None:
                height_known = False
                continue
            try:
                height_dxa += int(trH.get(f'{WNS}val'))
            except (TypeError, ValueError):
                height_known = False
        h_cm = dxa_to_cm(height_dxa) if height_known and height_dxa > 0 else None

        w_ok = in_range(w_cm, 8.70, 8.90)
        h_ok = (h_cm is not None) and in_range(h_cm, 9.40, 19.60)
        hit = w_ok and h_ok
        detail = (f'宽={w_cm:.2f}cm(要求8.70-8.90),'
                  + (f'长={h_cm:.2f}cm(要求9.40-19.60)'
                     if h_cm is not None
                     else '长=未知(存在自动撑开行,Word/WPS 中由内容排版决定)'))
    R.append((3, '右侧第3个文本框长9.40-19.60cm且宽8.70-8.90cm', hit, detail))

    # ---- 38. 右侧第4个文本框含第20小题+位置 (+3) ----
    # 细则：答题卡第一页右侧从上往下数第四个文本框包含 20 小题，
    #      放置在第一页答题卡右侧底部上方
    # 办公软件要点：
    #   - "第四个文本框"= doc.right_tables[3]（右半版面自上而下第 4 个块状容器）；
    #     Word/WPS 中"文本框"视觉容器既可以是真文本框(w:txbxContent)也可以是
    #     独立表格块，本文档以表格实现，视觉一致（与前述判定口径一致）
    #   - "放置在第一页答题卡右侧底部上方"：
    #     * 该容器的祖先链需能到达 doc.right_cell（第一页版面右半侧的承载单元格）
    #     * "底部上方"= 位置介于中部与最底部之间；即在 right_tables 顺序上属于第 4 个
    #       （右侧共有多个文本框时，第 4 个位于最底部之上的位置）
    #   - "包含 20 小题"：
    #     * 题号"20"以整数出现（用负向前/后瞻避免匹配到 200、120 等）
    #     * Word/WPS 中题号后常紧跟"."／"．"／"("／"（"／空白等分隔符，
    #       但细则未限定分隔符，只要求"包含 20 小题"，此处仅确认题号存在
    right3 = doc.right_tables[3] if len(doc.right_tables) > 3 else None
    hit = False
    detail = '未找到右侧第4个文本框(右侧第4块状容器)'
    if right3 is not None:
        in_right_cell = False
        anc = right3
        while anc is not None:
            if anc is doc.right_cell:
                in_right_cell = True
                break
            anc = anc.getparent()
        is_fourth = (len(doc.right_tables) >= 4 and doc.right_tables[3] is right3)
        full_txt = get_text(right3)
        has_q20 = re.search(r'(?<!\d)20(?!\d)', full_txt) is not None
        hit = in_right_cell and is_fourth and has_q20
        detail = (f'位于右半版面={in_right_cell},右侧第4块={is_fourth},'
                  f'含第20题={has_q20}')
    R.append((3, '右侧第4个文本框(位于第一页右侧底部上方)包含第20小题', hit, detail))

    # ---- 39. 右侧第4个文本框第20题下划线 (+3) ----
    # 细则：答题卡第一页右侧从上往下数第四个文本框，
    #      20 题内容为"题目序号后分数"带有三行或四行空白内容的下划线
    # 办公软件要点：
    #   - "第四个文本框"= right3 = doc.right_tables[3]（右半版面自上而下第 4 个块状容器）；
    #     Word/WPS 中"文本框"视觉容器既可为真文本框(w:txbxContent)也可为
    #     独立表格块，本文档以表格实现，视觉一致（与前述判定口径一致）
    #   - "题目序号后分数"：
    #     * 题号"20"直接后跟分数标记（形如"(X分)"），
    #       中间可含分隔符"."／"．"／"。"或全/半角空白
    #     * 细则未限定具体分值，只要求形如 (X分) 的分数标注
    #     * 括号"（）"细则用全角，办公软件中用户也可能录入半角"()"，做全/半角归一化
    #     * 句点"."办公软件中用户可能录入全角"．"或"。"，做归一化以兼容
    #   - "带有三行或四行空白内容的下划线"：
    #     * "行"= Word/WPS 段落(w:p)（用户按回车产生），软换行(w:br)也视为行边界
    #     * "空白内容的下划线"= 段落文本去掉横线字符后不含其它可见文字
    #     * 横线字符= 全角"＿"(U+FF3F) 或半角"_"(U+005F)，两者办公软件视觉一致
    #     * 行数为 3 或 4 皆合格（细则给出两种可选行数）
    #   - 细则本项未对下划线"长度"提出区间要求，此处只按"3 或 4 行空白下划线"判定
    hit = False
    detail = '未找到右侧第4个文本框(右侧第4块状容器)'
    if right3 is not None:
        def _norm39(s):
            return (s.replace('（', '(').replace('）', ')')
                     .replace('．', '.').replace('。', '.'))

        paras39 = [(p, get_text(p)) for p in right3.findall(f'.//{WNS}p')]

        def _is_blank_underline_39(text):
            stripped = text.replace('＿', '').replace('_', '')
            if stripped.strip():
                return False
            return (text.count('＿') + text.count('_')) > 0

        # 定位以 "20[.]?\s*(X分)" 开头的段索引
        start_20 = -1
        pat = re.compile(r'^\s*20[.]?\s*\(\s*\d+\s*分\s*\)')
        for i, (_p, t) in enumerate(paras39):
            nt = _norm39(t)
            if pat.match(nt):
                start_20 = i
                break

        # 从起始段之后连续收集"仅下划线/空白"段落的行数
        n_lines = 0
        if start_20 >= 0:
            for _p, t in paras39[start_20 + 1:]:
                if not t.strip():
                    continue
                if _is_blank_underline_39(t):
                    n_lines += 1
                else:
                    break

        found_20 = start_20 >= 0
        lines_ok = n_lines in (3, 4)
        hit = found_20 and lines_ok
        detail = (f'"20.(X分)"段={found_20},其后下划线行数={n_lines}(要求3或4)')
    R.append((
        3,
        '右侧第4个文本框第20题题号后分数带3或4行空白下划线',
        hit, detail,
    ))

    # ---- 40. 右侧第4个文本框尺寸 (+3) ----
    # 细则：答题卡第一页右侧从上往下数第四个文本框
    #      长约为 18.05-18.25 厘米，宽度约为 4.05-4.25 厘米
    # 办公软件要点：
    #   - "第四个文本框"= right3 = doc.right_tables[3]（右半版面自上而下第 4 个块状容器）；
    #     Word/WPS 中"文本框"视觉容器既可以是真文本框(w:txbxContent)也可以是
    #     独立表格块，本文档以表格实现，视觉一致（与前述判定口径一致）
    #   - 尺寸口径（针对细则原文的"长度"与"宽度"）：
    #     * 该文本框为纵向长条排布（承载第 20 大题的答题下划线），
    #       "长度"= 竖直方向的总高度（用户在 Word/WPS 页面上看到的容器高度），
    #       "宽度"= 水平方向的横向宽度
    #     * "宽度"读取来源（按办公软件级联优先级）：
    #         真文本框 → wp:extent@cx（EMU）或 v:shape 的 style width
    #         独立表格块 → 表格 w:tblPr/w:tblW[@w:type='dxa']，
    #           缺失时回退到 w:tblGrid 各 w:gridCol@w 累加（dxa）
    #     * "长度"（高度）读取来源：
    #         真文本框 → wp:extent@cy（EMU）
    #         独立表格块 → 所有 w:tr/w:trPr/w:trHeight@w:val 累加（dxa）；
    #           trHeight 缺失的行按内容撑开，无法精确度量，此时视为高度未知
    #   - 单位换算：1cm = 567 dxa（twentieths-of-a-point，Word/WPS 精确存储单位）；
    #                1cm = 360000 EMU（真文本框 anchor/inline 尺寸单位）
    hit = False
    detail = '未找到右侧第4个文本框(右侧第4块状容器)'
    if right3 is not None:
        # 宽度：优先 tblW(dxa)，回退 tblGrid 累加
        width_dxa = 0
        tblPr = right3.find(f'{WNS}tblPr')
        tblW = tblPr.find(f'{WNS}tblW') if tblPr is not None else None
        if tblW is not None and tblW.get(f'{WNS}type', 'dxa') in ('dxa', ''):
            try:
                width_dxa = int(tblW.get(f'{WNS}w', 0))
            except (TypeError, ValueError):
                width_dxa = 0
        if width_dxa <= 0:
            tblGrid = right3.find(f'{WNS}tblGrid')
            if tblGrid is not None:
                width_dxa = sum(int(gc.get(f'{WNS}w', 0))
                                for gc in tblGrid.findall(f'{WNS}gridCol'))
        w_cm = dxa_to_cm(width_dxa)

        # 长度（高度）：所有行 trHeight 累加
        height_dxa = 0
        height_known = True
        rows = right3.findall(f'{WNS}tr')
        for row in rows:
            trPr = row.find(f'{WNS}trPr')
            trH = trPr.find(f'{WNS}trHeight') if trPr is not None else None
            if trH is None or trH.get(f'{WNS}val') is None:
                height_known = False
                continue
            try:
                height_dxa += int(trH.get(f'{WNS}val'))
            except (TypeError, ValueError):
                height_known = False
        h_cm = dxa_to_cm(height_dxa) if height_known and height_dxa > 0 else None

        w_ok = in_range(w_cm, 4.05, 4.25)
        h_ok = (h_cm is not None) and in_range(h_cm, 18.05, 18.25)
        hit = w_ok and h_ok
        detail = (f'宽={w_cm:.2f}cm(要求4.05-4.25),'
                  + (f'长={h_cm:.2f}cm(要求18.05-18.25)'
                     if h_cm is not None
                     else '长=未知(存在自动撑开行,Word/WPS 中由内容排版决定)'))
    R.append((3, '右侧第4个文本框长18.05-18.25cm且宽4.05-4.25cm', hit, detail))

    # ---- 41. 第二页有两个文本框 (+3) ----
    # 细则：答题卡第二页有两个文本框
    # 办公软件要点（与第 17/29 项判定口径一致，避免同一细则的双重口径）：
    #   - "第二页"：Word/WPS 中"页"由 w:sectPr 分节隔开；本文档整体以嵌套表格作
    #     整版面容器，第二页对应文档 body 中 top_tables 顺序里承载第二页版面的
    #     那张外层表格（本项目结构解析中为 doc.top_tables[2]）
    #   - "文本框"= 视觉上呈现为独立框状区域的内容容器。在 Word/WPS 中，
    #     此类容器可能是：
    #       1) 真正的文本框（插入 → 文本框）：XML 存为 w:txbxContent
    #          （在 wps:txbx / v:textbox / w:pict / w:drawing 内）
    #       2) 用于版面隔离的独立表格块：外层容器表格首行(w:tr)下的 w:tc
    #     Word/WPS 在视觉上两者呈现一致（都是矩形框状内容区域），
    #     细则原文"文本框"未区分实现方式 —— 因此按视觉容器数量判定
    #   - 判定：优先按第二页容器内真文本框数量计数；若不含真文本框，
    #     则回退按外层容器表格首行 w:tc 数量计数
    #   - 需同时验证第二页存在：doc.page_count >= 2
    essay_grid_tbl = doc.top_tables[2] if len(doc.top_tables) > 2 else None
    hit = False
    detail = '未找到第二页版面容器'
    if doc.page_count >= 2 and essay_grid_tbl is not None:
        txbx_n = len(essay_grid_tbl.findall(f'.//{WNS}txbxContent'))
        if txbx_n > 0:
            n_boxes = txbx_n
            kind = '真文本框'
            hit = (n_boxes == 2)
            detail = f'第二页存在,{kind}数量={n_boxes}(要求2)'
        else:
            first_row = essay_grid_tbl.findall(f'{WNS}tr')
            if first_row:
                cells = first_row[0].findall(f'{WNS}tc')
                n_boxes = len(cells)
                hit = (n_boxes == 2)
                detail = f'第二页存在,版面块(表格)数量={n_boxes}(要求2)'
            else:
                detail = '第二页版面容器无行'
    elif doc.page_count < 2:
        detail = f'页数={doc.page_count}(要求≥2)'
    R.append((3, '第二页有两个文本框', hit, detail))

    # ---- 42. 第二页左侧文本框尺寸+"四、作文（50分）" (+3) ----
    # 细则：答题卡第二页左侧文本框
    #      长度为 18.45-18.65 厘米，宽度为 25.30-25.50 厘米，
    #      最上方包含"四、作文（50分）"字样
    # 办公软件要点：
    #   - "第二页左侧文本框"= essay_grid_tbl 首行第 1 个单元格（w:tc[0]）
    #     所对应的视觉容器；Word/WPS 中"文本框"视觉容器既可为真文本框
    #     (w:txbxContent)也可为独立表格块，本文档以"外层容器表格 + 单元格"
    #     方式呈现（与全卷判定口径一致）
    #   - 尺寸口径（针对细则原文的"长度"与"宽度"）：
    #     * 该左侧文本框内为作文答题格，"长度"= 水平方向的横向宽度（承载文字方向），
    #       "宽度"= 竖直方向的总高度（用户在 Word/WPS 页面上看到的容器高度）
    #       —— 与第一页各文本框（纵向长条排布）的"长度/宽度"含义相反，
    #       因本项左侧作文框在版面上为一整块矩形区域，细则数值 18.45-18.65cm
    #       对应水平尺寸，25.30-25.50cm 对应竖直尺寸
    #     * "长度"（水平宽度）读取来源（按办公软件级联优先级）：
    #         w:tc/w:tcPr/w:tcW[@w:type='dxa']（单元格宽度，dxa）
    #     * "宽度"（竖直高度）读取来源：
    #         该单元格内嵌套的表格所有 w:tr/w:trPr/w:trHeight@w:val 累加（dxa）；
    #         trHeight 缺失的行按内容撑开，无法精确度量，此时视为高度未知
    #   - "最上方包含'四、作文（50分）'字样"：
    #     * "最上方"= 该单元格视觉容器内文档流首个非空段落
    #     * 用户在 Word/WPS 中可能录入半角"()"，做全/半角归一化以兼容；
    #       "、"为标准中文顿号（U+3001），细则原文即此字符
    #   - 单位换算：1cm = 567 dxa（Word/WPS 精确存储单位）
    hit = False
    detail = '未找到第二页左侧文本框'
    if essay_grid_tbl is not None:
        rows = essay_grid_tbl.findall(f'{WNS}tr')
        if rows and rows[0].findall(f'{WNS}tc'):
            left_c = rows[0].findall(f'{WNS}tc')[0]

            # 长度（水平宽度）：tcW(dxa)
            tcPr = left_c.find(f'{WNS}tcPr')
            tcW = tcPr.find(f'{WNS}tcW') if tcPr is not None else None
            width_dxa = 0
            if tcW is not None and tcW.get(f'{WNS}type', 'dxa') in ('dxa', ''):
                try:
                    width_dxa = int(tcW.get(f'{WNS}w', 0))
                except (TypeError, ValueError):
                    width_dxa = 0
            w_cm = dxa_to_cm(width_dxa)

            # 宽度（竖直高度）：嵌套表格所有行 trHeight 累加
            height_dxa = 0
            height_known = True
            nested = left_c.findall(f'{WNS}tbl')
            if nested:
                for r in nested[0].findall(f'{WNS}tr'):
                    trPr = r.find(f'{WNS}trPr')
                    trH = trPr.find(f'{WNS}trHeight') if trPr is not None else None
                    if trH is None or trH.get(f'{WNS}val') is None:
                        height_known = False
                        continue
                    try:
                        height_dxa += int(trH.get(f'{WNS}val'))
                    except (TypeError, ValueError):
                        height_known = False
            else:
                height_known = False
            h_cm = dxa_to_cm(height_dxa) if height_known and height_dxa > 0 else None

            # 最上方包含"四、作文（50分）"字样
            def _norm42(s):
                return (s.replace('（', '(').replace('）', ')')
                         .replace('．', '.').replace('。', '.'))

            top_text = ''
            for p in left_c.findall(f'.//{WNS}p'):
                t = get_text(p)
                if t.strip():
                    top_text = t
                    break
            has_title = '四、作文' in _norm42(top_text) and '(50分)' in _norm42(top_text)

            len_ok = in_range(w_cm, 18.45, 18.65)
            wid_ok = (h_cm is not None) and in_range(h_cm, 25.30, 25.50)
            hit = len_ok and wid_ok and has_title
            detail = (f'长={w_cm:.2f}cm(要求18.45-18.65),'
                      + (f'宽={h_cm:.2f}cm(要求25.30-25.50)' if h_cm is not None
                         else '宽=未知(存在自动撑开行,Word/WPS 中由内容排版决定)')
                      + f',最上方含"四、作文(50分)"={has_title}')
    R.append((
        3,
        '第二页左侧文本框长18.45-18.65cm且宽25.30-25.50cm且最上方含"四、作文（50分）"',
        hit, detail,
    ))

    # ---- 43. 第二页右侧文本框尺寸+位置 (+3) ----
    # 细则：答题卡第二页右侧文本框
    #      长度为 18.45-18.65 厘米，宽度为 16.55-16.75 厘米，
    #      放置在答题卡右上方
    # 办公软件要点：
    #   - "第二页右侧文本框"= essay_grid_tbl 首行第 2 个单元格（w:tc[1]）
    #     所对应的视觉容器；Word/WPS 中"文本框"视觉容器既可为真文本框
    #     (w:txbxContent)也可为独立表格块，本文档以"外层容器表格 + 单元格"
    #     方式呈现（与全卷判定口径一致）
    #   - 尺寸口径（针对细则原文的"长度"与"宽度"）：
    #     * 该右侧文本框在版面上为一整块矩形区域，
    #       "长度"= 水平方向的横向宽度（承载文字方向），
    #       "宽度"= 竖直方向的总高度（用户在 Word/WPS 页面上看到的容器高度）
    #       —— 与第 42 项左侧作文框的口径一致
    #     * "长度"（水平宽度）读取来源：
    #         w:tc/w:tcPr/w:tcW[@w:type='dxa']（单元格宽度，dxa）
    #     * "宽度"（竖直高度）读取来源：
    #         该单元格内嵌套的表格所有 w:tr/w:trPr/w:trHeight@w:val 累加（dxa）；
    #         trHeight 缺失的行按内容撑开，无法精确度量，此时视为高度未知
    #   - "放置在答题卡右上方"：
    #     * 位于 essay_grid_tbl 首行(top row) → 第二页版面上方
    #     * 位于该行的第 2 个单元格(w:tc[1]) → 版面右侧
    #     * 二者共同定义"右上方"位置口径
    #   - 单位换算：1cm = 567 dxa（Word/WPS 精确存储单位）
    hit = False
    detail = '未找到第二页右侧文本框'
    if essay_grid_tbl is not None:
        rows = essay_grid_tbl.findall(f'{WNS}tr')
        if rows and len(rows[0].findall(f'{WNS}tc')) > 1:
            right_c = rows[0].findall(f'{WNS}tc')[1]
            # "上方"= 位于外层容器表格首行；"右方"= 该行第 2 个单元格
            is_top = True  # 已从 rows[0] 取出
            is_right = True  # 已从 [1] 取出（第 2 列即右侧）

            # 长度（水平宽度）：tcW(dxa)
            tcPr = right_c.find(f'{WNS}tcPr')
            tcW = tcPr.find(f'{WNS}tcW') if tcPr is not None else None
            width_dxa = 0
            if tcW is not None and tcW.get(f'{WNS}type', 'dxa') in ('dxa', ''):
                try:
                    width_dxa = int(tcW.get(f'{WNS}w', 0))
                except (TypeError, ValueError):
                    width_dxa = 0
            w_cm = dxa_to_cm(width_dxa)

            # 宽度（竖直高度）：嵌套表格所有行 trHeight 累加
            height_dxa = 0
            height_known = True
            nested = right_c.findall(f'{WNS}tbl')
            if nested:
                for r in nested[0].findall(f'{WNS}tr'):
                    trPr = r.find(f'{WNS}trPr')
                    trH = trPr.find(f'{WNS}trHeight') if trPr is not None else None
                    if trH is None or trH.get(f'{WNS}val') is None:
                        height_known = False
                        continue
                    try:
                        height_dxa += int(trH.get(f'{WNS}val'))
                    except (TypeError, ValueError):
                        height_known = False
            else:
                height_known = False
            h_cm = dxa_to_cm(height_dxa) if height_known and height_dxa > 0 else None

            len_ok = in_range(w_cm, 18.45, 18.65)
            wid_ok = (h_cm is not None) and in_range(h_cm, 16.55, 16.75)
            hit = len_ok and wid_ok and is_top and is_right
            detail = (f'长={w_cm:.2f}cm(要求18.45-18.65),'
                      + (f'宽={h_cm:.2f}cm(要求16.55-16.75)' if h_cm is not None
                         else '宽=未知(存在自动撑开行,Word/WPS 中由内容排版决定)')
                      + f',位置=右上方(顶行右列)')
    R.append((
        3,
        '第二页右侧文本框长18.45-18.65cm且宽16.55-16.75cm且位于右上方',
        hit, detail,
    ))

    # ---- 44. 第二页左侧文本框26个表格+尺寸+间距 (+5) ----
    # 细则：答题卡第二页左侧文本框中有26个表格，每个表格都是1行22列，
    #      多个表格竖向排列，每个表格中的单元格行高为0.74-0.76厘米，
    #      列宽为0.81-0.83厘米，每两个表格之间的间距为0.26-0.28厘米
    #      （注：细则原文"0.81-0.03"为明显笔误，正确区间为0.81-0.83，
    #       与列宽应为增区间且和第45项右侧同结构表格保持一致）
    # 办公软件要点：
    #   - "第二页左侧文本框"= essay_grid_tbl 首行第 1 个单元格（与第 42 项一致）
    #   - "26 个表格"= 该单元格直接子节点中 w:tbl 元素恰为 26 个
    #     （Word/WPS 中每个 w:tbl 在页面上就是一个独立的表格视觉块）
    #   - "1 行 22 列"= 每个 w:tbl 有且仅有 1 个 w:tr，该 w:tr 有且仅有 22 个 w:tc
    #   - "多个表格竖向排列"= 26 个 w:tbl 作为该单元格的直接子节点串行排列，
    #     两个 w:tbl 之间由 w:p（段落/空段）分隔——这是 Word/WPS 中"表格
    #     竖向堆叠"的标准 OOXML 存储方式（无法用横向浮动实现严格堆叠）
    #   - "单元格行高 0.74-0.76 厘米"：读取每个 w:tbl 唯一 w:tr 的
    #     w:trPr/w:trHeight@w:val（dxa），换算为 cm，要求 in_range(0.74, 0.76)
    #   - "列宽 0.81-0.83 厘米"：读取每个 w:tbl 的 w:tblGrid/w:gridCol@w（dxa），
    #     22 列每列均要求 in_range(0.81, 0.83)
    #   - "每两个表格之间的间距 0.26-0.28 厘米"：
    #     * 该间距 = 相邻两个 w:tbl 之间夹的 w:p（分隔段）在页面上的垂直占位
    #     * Word/WPS 中一个段落的垂直高度 = spacing.before + 行高 + spacing.after
    #       - lineRule="exact"/"atLeast"：line 为 dxa（twips）
    #       - lineRule="auto" 或缺省：line 为 240 单位表示 1 倍行距，
    #         视觉高度 ≈ 字号(pt) × (line/240) × 20 dxa
    #       - before/after 若 beforeAutospacing/afterAutospacing 打开，
    #         Word/WPS 会自动放大，本项 25 项已限定"仅下划线/空白"段，通常关闭
    #     * 若相邻表之间存在多个分隔段，累加为整个"间距"
    #   - 单位换算：1cm = 567 dxa
    hit = False
    detail = '未找到第二页左侧文本框'
    if essay_grid_tbl is not None:
        rows0 = essay_grid_tbl.findall(f'{WNS}tr')
        if rows0 and rows0[0].findall(f'{WNS}tc'):
            left_c = rows0[0].findall(f'{WNS}tc')[0]
            # 保持文档流顺序（该单元格的直接子节点：交替出现 w:tbl / w:p）
            direct_children = list(left_c)
            nested = [c for c in direct_children if c.tag == f'{WNS}tbl']

            n_tbls = len(nested)
            count_ok = (n_tbls == 26)

            # 结构：每个 w:tbl 恰 1 行 22 列
            struct_bad = []
            for i, t in enumerate(nested):
                trs = t.findall(f'{WNS}tr')
                if len(trs) != 1:
                    struct_bad.append(f'#{i}行数={len(trs)}')
                    continue
                tcs = trs[0].findall(f'{WNS}tc')
                if len(tcs) != 22:
                    struct_bad.append(f'#{i}列数={len(tcs)}')
            struct_ok = (not struct_bad)

            # 行高 0.74-0.76 cm
            rowh_bad = []
            for i, t in enumerate(nested):
                trs = t.findall(f'{WNS}tr')
                if not trs:
                    continue
                trPr = trs[0].find(f'{WNS}trPr')
                trH = trPr.find(f'{WNS}trHeight') if trPr is not None else None
                if trH is None or trH.get(f'{WNS}val') is None:
                    rowh_bad.append(f'#{i}无trHeight')
                    continue
                try:
                    h = dxa_to_cm(int(trH.get(f'{WNS}val')))
                except (TypeError, ValueError):
                    rowh_bad.append(f'#{i}行高解析失败')
                    continue
                if not in_range(h, 0.74, 0.76):
                    rowh_bad.append(f'#{i}行高={h:.2f}')
            rowh_ok = (not rowh_bad)

            # 列宽 0.81-0.83 cm（22 列每列）
            colw_bad = []
            for i, t in enumerate(nested):
                tblGrid = t.find(f'{WNS}tblGrid')
                if tblGrid is None:
                    colw_bad.append(f'#{i}无tblGrid')
                    continue
                gcs = tblGrid.findall(f'{WNS}gridCol')
                for j, gc in enumerate(gcs):
                    try:
                        w_cm = dxa_to_cm(int(gc.get(f'{WNS}w', 0)))
                    except (TypeError, ValueError):
                        colw_bad.append(f'#{i}c{j}宽解析失败')
                        continue
                    if not in_range(w_cm, 0.81, 0.83):
                        colw_bad.append(f'#{i}c{j}={w_cm:.2f}')
                        break  # 该表已判失败，跳到下一个
            colw_ok = (not colw_bad)

            # "竖向排列"：所有 w:tbl 位于同一容器（left_c）下作为并列直接子节点
            # 该形式在 Word/WPS 中天然表现为自上而下堆叠（无浮动 tblpPr）
            vertical_ok = True
            for t in nested:
                tblPr = t.find(f'{WNS}tblPr')
                tblpPr = tblPr.find(f'{WNS}tblpPr') if tblPr is not None else None
                if tblpPr is not None:
                    # 浮动定位会破坏"竖向堆叠"的视觉呈现
                    vertical_ok = False
                    break

            # 相邻表间距 0.26-0.28 cm：读夹在两 tbl 之间的所有 w:p 累加垂直高度
            def _para_height_cm_44(p):
                pPr = p.find(f'{WNS}pPr')
                spacing = pPr.find(f'{WNS}spacing') if pPr is not None else None
                line_val = spacing.get(f'{WNS}line') if spacing is not None else None
                line_rule = spacing.get(f'{WNS}lineRule') if spacing is not None else None
                before = spacing.get(f'{WNS}before') if spacing is not None else None
                after = spacing.get(f'{WNS}after') if spacing is not None else None
                # 行高
                if line_val is not None:
                    try:
                        li = int(line_val)
                    except (TypeError, ValueError):
                        li = 0
                    if line_rule in ('exact', 'atLeast'):
                        line_dxa = li  # 已是 dxa
                    else:  # auto 或缺省
                        _f, sz, _b = get_run_font_info(p)
                        if sz is None:
                            sz = 10.5  # Word 默认字号回退
                        line_dxa = int(sz * (li / 240.0) * 20)
                else:
                    _f, sz, _b = get_run_font_info(p)
                    if sz is None:
                        sz = 10.5
                    line_dxa = int(sz * 20 * 1.15)  # 单倍行距的默认视觉高度近似
                # 前后间距
                bef = int(before) if before is not None else 0
                aft = int(after) if after is not None else 0
                return dxa_to_cm(line_dxa + bef + aft)

            gap_bad = []
            positions = [direct_children.index(t) for t in nested]
            for k in range(len(positions) - 1):
                a, b = positions[k], positions[k + 1]
                paras_between = [
                    c for c in direct_children[a + 1:b] if c.tag == f'{WNS}p'
                ]
                gap_cm = sum(_para_height_cm_44(p) for p in paras_between)
                if not in_range(gap_cm, 0.26, 0.28):
                    gap_bad.append(f'#{k}-#{k+1}={gap_cm:.2f}')
            gap_ok = (not gap_bad) and len(positions) >= 2

            hit = (count_ok and struct_ok and rowh_ok and colw_ok
                   and vertical_ok and gap_ok)
            detail = (
                f'表格数={n_tbls}(要求26),结构1x22={struct_ok},'
                f'行高0.74-0.76={rowh_ok}'
                + (f'/异常={rowh_bad[:2]}' if rowh_bad else '')
                + f',列宽0.81-0.83={colw_ok}'
                + (f'/异常={colw_bad[:2]}' if colw_bad else '')
                + f',竖向排列={vertical_ok},'
                f'相邻间距0.26-0.28={gap_ok}'
                + (f'/异常={gap_bad[:2]}' if gap_bad else '')
            )
    R.append((
        5,
        ('第二页左侧文本框中有26个表格，每个表格都是1行22列，竖向排列'
         '(行高0.74-0.76cm,列宽0.81-0.83cm,相邻间距0.26-0.28cm)'),
        hit, detail,
    ))

    # ---- 45. 第二页右侧文本框17个表格+尺寸+间距 (+5) ----
    # 细则：答题卡第二页右侧文本框中有17个表格，每个表格都是1行22列，
    #      多个表格竖向排列，每个表格中的单元格行高为0.74-0.76厘米，
    #      列宽为0.81-0.83厘米，每两个表格之间的间距为0.26-0.28厘米
    #      （注：细则原文"0.81-0.03"为明显笔误，正确区间为0.81-0.83，
    #       与第44项左侧同结构表格保持一致）
    # 办公软件要点（与第 44 项同构，仅表格数量与承载单元格不同）：
    #   - "第二页右侧文本框"= essay_grid_tbl 首行第 2 个单元格（与第 43 项一致）
    #   - "17 个表格"= 该单元格直接子节点中 w:tbl 元素恰为 17 个
    #   - "1 行 22 列"= 每个 w:tbl 有且仅有 1 个 w:tr，该 w:tr 有且仅有 22 个 w:tc
    #   - "多个表格竖向排列"= 17 个 w:tbl 作为该单元格的直接子节点串行排列，
    #     两个 w:tbl 之间由 w:p（段落/空段）分隔；无 w:tblpPr 浮动定位
    #   - 行高 / 列宽 / 相邻表间距计算与判定口径与第 44 项完全一致
    #   - 单位换算：1cm = 567 dxa
    hit = False
    detail = '未找到第二页右侧文本框'
    if essay_grid_tbl is not None:
        rows0 = essay_grid_tbl.findall(f'{WNS}tr')
        if rows0 and len(rows0[0].findall(f'{WNS}tc')) > 1:
            right_c = rows0[0].findall(f'{WNS}tc')[1]
            direct_children = list(right_c)
            nested = [c for c in direct_children if c.tag == f'{WNS}tbl']

            n_tbls = len(nested)
            count_ok = (n_tbls == 17)

            # 结构：每个 w:tbl 恰 1 行 22 列
            struct_bad = []
            for i, t in enumerate(nested):
                trs = t.findall(f'{WNS}tr')
                if len(trs) != 1:
                    struct_bad.append(f'#{i}行数={len(trs)}')
                    continue
                tcs = trs[0].findall(f'{WNS}tc')
                if len(tcs) != 22:
                    struct_bad.append(f'#{i}列数={len(tcs)}')
            struct_ok = (not struct_bad)

            # 行高 0.74-0.76 cm
            rowh_bad = []
            for i, t in enumerate(nested):
                trs = t.findall(f'{WNS}tr')
                if not trs:
                    continue
                trPr = trs[0].find(f'{WNS}trPr')
                trH = trPr.find(f'{WNS}trHeight') if trPr is not None else None
                if trH is None or trH.get(f'{WNS}val') is None:
                    rowh_bad.append(f'#{i}无trHeight')
                    continue
                try:
                    h = dxa_to_cm(int(trH.get(f'{WNS}val')))
                except (TypeError, ValueError):
                    rowh_bad.append(f'#{i}行高解析失败')
                    continue
                if not in_range(h, 0.74, 0.76):
                    rowh_bad.append(f'#{i}行高={h:.2f}')
            rowh_ok = (not rowh_bad)

            # 列宽 0.81-0.83 cm（22 列每列）
            colw_bad = []
            for i, t in enumerate(nested):
                tblGrid = t.find(f'{WNS}tblGrid')
                if tblGrid is None:
                    colw_bad.append(f'#{i}无tblGrid')
                    continue
                gcs = tblGrid.findall(f'{WNS}gridCol')
                for j, gc in enumerate(gcs):
                    try:
                        w_cm = dxa_to_cm(int(gc.get(f'{WNS}w', 0)))
                    except (TypeError, ValueError):
                        colw_bad.append(f'#{i}c{j}宽解析失败')
                        continue
                    if not in_range(w_cm, 0.81, 0.83):
                        colw_bad.append(f'#{i}c{j}={w_cm:.2f}')
                        break
            colw_ok = (not colw_bad)

            # "竖向排列"：所有 w:tbl 位于同一容器（right_c）下作为并列直接子节点，
            # 且无 w:tblpPr 浮动定位；Word/WPS 中天然自上而下堆叠
            vertical_ok = True
            for t in nested:
                tblPr = t.find(f'{WNS}tblPr')
                tblpPr = tblPr.find(f'{WNS}tblpPr') if tblPr is not None else None
                if tblpPr is not None:
                    vertical_ok = False
                    break

            # 相邻表间距 0.26-0.28 cm：夹在两 tbl 之间的所有 w:p 累加视觉高度
            def _para_height_cm_45(p):
                pPr = p.find(f'{WNS}pPr')
                spacing = pPr.find(f'{WNS}spacing') if pPr is not None else None
                line_val = spacing.get(f'{WNS}line') if spacing is not None else None
                line_rule = spacing.get(f'{WNS}lineRule') if spacing is not None else None
                before = spacing.get(f'{WNS}before') if spacing is not None else None
                after = spacing.get(f'{WNS}after') if spacing is not None else None
                if line_val is not None:
                    try:
                        li = int(line_val)
                    except (TypeError, ValueError):
                        li = 0
                    if line_rule in ('exact', 'atLeast'):
                        line_dxa = li  # 已是 dxa
                    else:  # auto 或缺省
                        _f, sz, _b = get_run_font_info(p)
                        if sz is None:
                            sz = 10.5  # Word 默认字号回退
                        line_dxa = int(sz * (li / 240.0) * 20)
                else:
                    _f, sz, _b = get_run_font_info(p)
                    if sz is None:
                        sz = 10.5
                    line_dxa = int(sz * 20 * 1.15)
                bef = int(before) if before is not None else 0
                aft = int(after) if after is not None else 0
                return dxa_to_cm(line_dxa + bef + aft)

            gap_bad = []
            positions = [direct_children.index(t) for t in nested]
            for k in range(len(positions) - 1):
                a, b = positions[k], positions[k + 1]
                paras_between = [
                    c for c in direct_children[a + 1:b] if c.tag == f'{WNS}p'
                ]
                gap_cm = sum(_para_height_cm_45(p) for p in paras_between)
                if not in_range(gap_cm, 0.26, 0.28):
                    gap_bad.append(f'#{k}-#{k+1}={gap_cm:.2f}')
            gap_ok = (not gap_bad) and len(positions) >= 2

            hit = (count_ok and struct_ok and rowh_ok and colw_ok
                   and vertical_ok and gap_ok)
            detail = (
                f'表格数={n_tbls}(要求17),结构1x22={struct_ok},'
                f'行高0.74-0.76={rowh_ok}'
                + (f'/异常={rowh_bad[:2]}' if rowh_bad else '')
                + f',列宽0.81-0.83={colw_ok}'
                + (f'/异常={colw_bad[:2]}' if colw_bad else '')
                + f',竖向排列={vertical_ok},'
                f'相邻间距0.26-0.28={gap_ok}'
                + (f'/异常={gap_bad[:2]}' if gap_bad else '')
            )
    R.append((
        5,
        ('第二页右侧文本框中有17个表格，每个表格都是1行22列，竖向排列'
         '(行高0.74-0.76cm,列宽0.81-0.83cm,相邻间距0.26-0.28cm)'),
        hit, detail,
    ))

    # ---- 严格判定收敛：细则所有"文本框"条目要求存在真·文本框 ----
    # 说明：办公软件(Word/WPS)中"文本框"由"插入→文本框"生成，
    #       OOXML 元素为 w:txbxContent (通常包裹在 wps:txbx / v:textbox /
    #       w:pict / w:drawing 内)。独立表格块虽视觉相似，字面上不是"文本框"。
    #       按用户严格字面口径，若文档中不存在任何 w:txbxContent，
    #       所有细则文本"文本框"条目一律判为未命中。
    all_txbx_n = len(doc.body.findall(f'.//{WNS}txbxContent'))
    if all_txbx_n == 0:
        _new_R = []
        for _score, _desc, _hit, _det in R:
            if '文本框' in _desc and _hit:
                _hit = False
                _det = f'文档无真·文本框(w:txbxContent=0)→严格判定不通过;原详情:{_det}'
            _new_R.append((_score, _desc, _hit, _det))
        R = _new_R

    return R


# ============ 主函数 ============
def _locate_docx(dir_path):
    """在给定目录中定位被评估的 .docx 文件（忽略 Word 临时文件 ~$*.docx）"""
    if not os.path.isdir(dir_path):
        return None
    candidates = []
    for name in os.listdir(dir_path):
        if name.startswith('~$'):
            continue
        if name.lower().endswith('.docx'):
            candidates.append(name)
    if not candidates:
        return None
    # 若含指定文件名则优先返回，否则返回第一个
    preferred = '初中语文跨年级综合质量检测_答题卡_可编辑.docx'
    if preferred in candidates:
        return os.path.join(dir_path, preferred)
    return os.path.join(dir_path, candidates[0])


def evaluate(dir_path: str) -> dict:
    """统一入口：接收脚本所在目录路径，自行在该目录中定位并评估被评估文档。

    返回结构见《脚本接口差异与统一建议》§2.2。
    """
    result = {
        'id': '015',
        'file_name': '',
        'status': 'ok',
        'error': None,
        'dim1_pass': False,
        'dim1_reason': '',
        'dim2_items': [],
        'total_score': 0,
        'max_score': 0,
    }

    try:
        filepath = _locate_docx(dir_path)
        if filepath is None:
            result['status'] = 'error'
            result['error'] = f'目录中未找到 .docx 文件: {dir_path}'
            return result
        result['file_name'] = os.path.basename(filepath)

        if not os.path.exists(filepath):
            result['status'] = 'error'
            result['error'] = f'文件不存在: {filepath}'
            return result

        try:
            doc = DocParser(filepath)
        except Exception as e:
            result['status'] = 'error'
            result['error'] = f'无法打开文件: {e}'
            return result

        # 维度一
        dim1_pass, dim1_results = check_dim1(doc)
        result['dim1_pass'] = bool(dim1_pass)
        if not dim1_pass:
            failed = [name for name, passed, _detail in dim1_results if not passed]
            result['dim1_reason'] = '；'.join(failed) if failed else '未通过'
            # 一票否决：维度一未通过时直接短路，不再评估维度二
            result['dim2_items'] = []
            result['total_score'] = 0
            result['max_score'] = 0
            return result

        # 维度二（维度一已通过才会执行到这里）
        dim2_results = check_dim2(doc)
        items = []
        total = 0
        max_score = 0
        for score, desc, hit, detail in dim2_results:
            max_score += score
            delta = score if hit else 0
            if hit:
                total += score
            items.append({
                'rule': desc,
                'max_delta': score,
                'delta': delta,
                'hit': bool(hit),
                'detail': '',
            })
        result['dim2_items'] = items
        result['total_score'] = total
        result['max_score'] = max_score
        return result

    except Exception as e:
        result['status'] = 'error'
        result['error'] = f'脚本异常: {e}'
        return result


if __name__ == '__main__':
    # 本地调试入口：默认以脚本所在目录作为 dir_path
    target_dir = sys.argv[1] if len(sys.argv) > 1 else os.path.dirname(os.path.abspath(__file__))
    print(json.dumps(evaluate(target_dir), ensure_ascii=False, indent=2))
