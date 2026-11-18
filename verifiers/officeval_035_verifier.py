#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# pyright: reportUnknownMemberType=false, reportUnknownVariableType=false, reportUnknownArgumentType=false, reportMissingImports=false, reportAttributeAccessIssue=false, reportOptionalMemberAccess=false, reportOptionalSubscript=false, reportOptionalIterable=false, reportUnknownParameterType=false, reportMissingParameterType=false, reportIndexIssue=false, reportArgumentType=false, reportCallIssue=false, reportGeneralTypeIssues=false

"""外文文献翻译任务（002 毕业设计外文翻译封面及格式要求·普通班）自动评分脚本。

接口约定：
- 本模块不修改全局 sys.stdout，也不打印主结果；
- 主结果只由 evaluate(dir_path) -> dict 返回；
- 入参是"脚本所在目录的路径"，脚本自行在该目录下定位并打开被评测的 .docx；
- __main__ 只是本地调试入口，把 evaluate 的返回值序列化成 JSON 打印。

计分方式：33 条加分项不分组、不加权，命中即得该条满分，得分直接累加；
    完成度 = total_score ÷ max_score × 100（max_score = 95）。
    维度1（文档可正常打开）不给分，只作为闸门：不通过时 dim1_pass=False、
    dim1_reason 回填失败原因、dim2_items 为空、total_score 记 0；通过时
    dim1_reason 为空字符串。脚本未能完成判定（含目录内找不到 .docx 等取件
    失败）时 status='error'，dim1_pass 与 dim1_reason 均留空。
"""

import os
import re
import sys
import json
import math

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH

from docx.oxml.ns import qn


SCRIPT_ID = "035"
MAX_SCORE_TOTAL = 95

# 满分序列：顺序必须与 _evaluate_impl 中 results 的追加顺序严格一致
# 33 条加分项统一放在一起，不分组、不加权
_MAX_DELTAS_ORDER = [
    3,   # 01 13个一级章节标题全部存在并保留原编号
    5,   # 02 24个二级小节标题全部存在并保留原编号
    5,   # 03 全部37个章节下均有中文正文，无空章节、无整段英文未译
    3,   # 04 摘要与5个关键词完整存在
    3,   # 05 后置4项（作者贡献声明、利益冲突声明、数据可用性、致谢）齐全并沿用无编号形态
    1,   # 06 参考文献[1]–[15]齐全并保留英文原文
    5,   # 07 5张表格为真实表格对象且行列数与原文一致
    3,   # 08 表格单元格无空缺、无合并丢失、无行列错位
    5,   # 09 图1–图7全部插入
    3,   # 10 7张插图均位于原文对应章节位置
    3,   # 11 全部12个题注存在，位置正确、编号正确、未被拆成两段
    1,   # 12 正文段落划分与原文一致，无拆段断词
    1,   # 13 无原文页眉、期刊名、DOI、页码等非正文信息被搬入
    1,   # 14 002模板占位文字与格式说明文字全部清除
    3,   # 15 7张插图比例未失真、不超版心、清晰、边界干净
    1,   # 16 客户原始输入文件保持原样未被改动
    3,   # 17 页面为A4且四边页边距均为2.5cm
    1,   # 18 页眉距边界1.5cm、页脚距边界1.5cm，且启用文档网格
    5,   # 19 封面固定元素完整复刻
    1,   # 20 封面题目栏填入翻译后的中文题名
    3,   # 21 封面独立成节，无页眉、无页脚、无页码
    5,   # 22 奇偶页页眉分设，均居中且带下框线
    3,   # 23 页脚为居中PAGE页码域，且正文首页页码为1
    3,   # 24 篇名二号黑体居中，段前0.5行、段后0.5行
    1,   # 25 摘要为五号宋体+Times New Roman
    1,   # 26 「关键词」标签加粗，词条本身不加粗
    1,   # 27 摘要末段后、关键词前、关键词后共三处空行齐全
    5,   # 28 全部13个一级标题为三号黑体居左、段前后各1行、单倍行距
    5,   # 29 全部24个二级标题为小三号黑体+TNR、两端对齐、段前后各0.5行、节号空1格
    3,   # 30 全部37个标题设有正确大纲级别（一级=级别1，二级=级别2）
    3,   # 31 正文为小四号，中文宋体、西文Times New Roman
    3,   # 32 正文为单倍行距且已勾选"对齐到文档网格"
    3,   # 33 正文首行缩进0.85cm
]
assert sum(_MAX_DELTAS_ORDER) == MAX_SCORE_TOTAL

# ======================= 原文（neighborhood_reading_garden_article.pdf）踩点常量 =======================

EXPECTED_L1_NUMBERS = tuple(range(1, 14))                                   # 13 个一级章节：1–13
EXPECTED_L2_KEYS = tuple((s, n) for s in range(1, 13) for n in (1, 2))      # 24 个二级小节：1.1–12.2
EXPECTED_SECTION_TOTAL = 37                                                 # 13 + 24
EXPECTED_KEYWORD_COUNT = 5
EXPECTED_REF_NUMBERS = tuple(range(1, 16))                                  # 参考文献 [1]–[15]
EXPECTED_TABLE_SHAPES = ((7, 3), (13, 7), (7, 4), (6, 3), (7, 4))           # 表1–表5 行×列
EXPECTED_FIGURE_COUNT = 7
EXPECTED_CAPTION_COUNT = 12                                                 # 图1–图7 + 表1–表5
FIGURE_HOST_SECTION = {1: 1, 2: 2, 3: 3, 4: 4, 5: 6, 6: 7, 7: 9}            # 插图所属一级章节
TABLE_HOST_SECTION = {1: 2, 2: 4, 3: 6, 4: 8, 5: 12}                        # 表格所属一级章节

# ======================= 002 模板踩点常量 =======================

COVER_TITLE_MAIN = '毕业设计（论文）'
COVER_TITLE_SUB = '外文翻译'
COVER_INFO_SHAPE = (7, 2)
COVER_YEAR_SHAPE = (1, 2)
COVER_TITLE_PLACEHOLDER = '文献翻译后的中文篇名'
ODD_HEADER_TEXT = '外文翻译'
EVEN_HEADER_TEXT = '2022届材料科学与工程专业毕业设计（论文）'

A4_WIDTH_CM = 21.0
A4_HEIGHT_CM = 29.7
MARGIN_CM = 2.5
HEADER_FOOTER_DIST_CM = 1.5
FIRST_LINE_TWIPS = 482          # 0.85cm ≈ 482 twips；字符写法须折算成实际长度后再比对

PT_TITLE = 22.0                 # 二号（篇名）
PT_COVER_MAIN = 42.0            # 微软雅黑 42pt
PT_COVER_SUB = 22.0             # 等线 22pt 加粗
PT_ABSTRACT = 10.5              # 五号
PT_H1 = 16.0                    # 三号
PT_H2 = 15.0                    # 小三号
PT_BODY = 12.0                  # 小四号

STRICT_SONGTI = ('宋体', 'SimSun', 'NSimSun', 'STSong', 'Songti SC')
STRICT_HEITI = ('黑体', 'SimHei', 'STHeiti', 'Heiti SC')
STRICT_TNR = ('Times New Roman', 'TimesNewRoman', 'Times New Roman PS')
STRICT_YAHEI = ('微软雅黑', 'Microsoft YaHei', 'MicrosoftYaHei')
STRICT_DENGXIAN = ('等线', 'DengXian', '等线 Light', 'DengXian Light')

# ======================= 识别用正则 =======================

W_P = qn('w:p')
W_TBL = qn('w:tbl')
MC_FALLBACK = '{http://schemas.openxmlformats.org/markup-compatibility/2006}Fallback'

_RE_CJK = re.compile(r'[一-鿿㐀-䶿]')
_RE_L1 = re.compile(r'^(\d{1,2})[.、]?[\s　\xa0]+\S')
_RE_L2 = re.compile(r'^(\d{1,2})\.(\d{1,2})([\s　\xa0]*)(\S.*)$')
# 题注：编号后必须有分隔符（标点或空格），据此与正文里的"图4展示了…"这类引用区分开
_RE_CAPTION = re.compile(r'^(图|表)[\s　]*(\d{1,2})[\s　]*(?:[.．、：:，,]|[\s　])[\s　]*\S')
_RE_REF_ENTRY = re.compile(r'^\[(\d{1,2})\]')
_RE_REFS_HEAD = re.compile(r'^参考文献')
_RE_ABSTRACT = re.compile(r'^[【\[]?\s*摘\s*要')
_RE_KEYWORD = re.compile(r'^[【\[]?\s*关\s*键\s*词')
_RE_YEAR_CELL = re.compile(r'二[〇零○0]')

# 后置 4 项：只按关键词判存在，不校验译法（翻译存在不确定性）
_BACK_PATTERNS = (
    ('contrib', re.compile(r'^(作者)?贡献')),
    ('coi', re.compile(r'^(利益冲突|竞争性利益|竞争利益|竞争性权益)')),
    ('data', re.compile(r'^数据(可用性|可获得性|可得性|可用)')),
    ('ack', re.compile(r'^致\s*谢')),
)

# 原文页眉／期刊信息／DOI／页码等非正文信息
_FOREIGN_ARTIFACTS = (
    re.compile(r'Contents\s+lists\s+available', re.I),
    re.compile(r'ScienceDirect', re.I),
    re.compile(r'journal\s+homepage', re.I),
    re.compile(r'\bdoi\b\s*[:：]', re.I),
    re.compile(r'https?://doi\.org', re.I),
    re.compile(r'\b10\.\d{4,9}/\S+'),
    re.compile(r'\bISSN\b', re.I),
    re.compile(r'Received\s+\d{1,2}\s+\w+\s+\d{4}', re.I),
    re.compile(r'All\s+rights\s+reserved', re.I),
    re.compile(r'Elsevier', re.I),
    re.compile(r'©\s*\d{4}'),
)

# 002 模板占位文字与格式说明文字
_TEMPLATE_RESIDUE = (
    '省略不翻', '文献翻译后的中文篇名', '文献翻译后的中文题目', '关键词1',
    '正文内容用小四号字', '对齐到文档网格', '二号黑体居中', '三号黑体居左',
    '小三号黑体', '此处填写', '（空一行）', '(空一行)', '空一行',
)

_TERMINAL_CHARS = '。！？；：!?;:.）)】」』”"…—'

# ======================= 通用工具 =======================

_STYLE_CACHE = {}


def _norm(text):
    """去掉全部空白，便于与模板固定文字做"与办公软件所见一致"的比对。"""
    return re.sub(r'[\s　\xa0]+', '', text or '')


def _el_text(el):
    if el is None:
        return ''
    return ''.join(node.text or '' for node in el.iter(qn('w:t')))


def _has_cjk(text):
    return bool(_RE_CJK.search(text or ''))


def _font_hit(name, allow):
    if not name:
        return False
    key = str(name).strip().lower()
    return any(key == str(a).strip().lower() for a in allow)


def _pt_from_sz(val):
    try:
        return float(val) / 2.0
    except (TypeError, ValueError):
        return None


def _bool_of(el):
    if el is None:
        return None
    val = el.get(qn('w:val'))
    if val is None:
        return True
    return str(val).strip().lower() not in ('0', 'false', 'off')


def _int_of(val, default=None):
    try:
        return int(round(float(val)))
    except (TypeError, ValueError):
        return default


def _emu_to_cm(val):
    num = _int_of(val)
    return None if num is None else num / 360000.0


def _close(a, b, tol):
    return a is not None and abs(float(a) - float(b)) <= tol

# ======================= 样式继承链解析 =======================

def _style_index(doc):
    key = id(doc)
    cached = _STYLE_CACHE.get(key)
    if cached is None:
        cached = {}
        for st in doc.styles.element.findall(qn('w:style')):
            sid = st.get(qn('w:styleId'))
            if sid:
                cached[sid] = st
        _STYLE_CACHE[key] = cached
    return cached


def _default_para_style_id(doc):
    for st in doc.styles.element.findall(qn('w:style')):
        if st.get(qn('w:type')) == 'paragraph' and st.get(qn('w:default')) == '1':
            return st.get(qn('w:styleId'))
    return 'Normal'


def _para_style_id(doc, p_el):
    ppr = p_el.find(qn('w:pPr'))
    if ppr is not None:
        pstyle = ppr.find(qn('w:pStyle'))
        if pstyle is not None and pstyle.get(qn('w:val')):
            return pstyle.get(qn('w:val'))
    return _default_para_style_id(doc)


def _iter_style_chain(doc, style_id):
    """沿 basedOn 逐级向上遍历样式链。"""
    index = _style_index(doc)
    seen = set()
    cur = style_id
    while cur and cur in index and cur not in seen:
        seen.add(cur)
        style = index[cur]
        yield style
        based = style.find(qn('w:basedOn'))
        cur = based.get(qn('w:val')) if based is not None else None


def _doc_defaults_rpr(doc):
    dd = doc.styles.element.find(qn('w:docDefaults'))
    if dd is None:
        return None
    rpr_default = dd.find(qn('w:rPrDefault'))
    return None if rpr_default is None else rpr_default.find(qn('w:rPr'))

def _run_element(run):
    return run._element if hasattr(run, '_element') else run


def _iter_rpr_chain(doc, run):
    """run 直接属性 -> 所在段落的样式链 -> docDefaults，依次产出 rPr 元素。"""
    el = _run_element(run)
    rpr = el.find(qn('w:rPr'))
    if rpr is not None:
        yield rpr
    parent = el.getparent()
    if parent is not None and parent.tag == W_P:
        for style in _iter_style_chain(doc, _para_style_id(doc, parent)):
            srpr = style.find(qn('w:rPr'))
            if srpr is not None:
                yield srpr
    default_rpr = _doc_defaults_rpr(doc)
    if default_rpr is not None:
        yield default_rpr


def _resolve_effective_eastasia(doc, run):
    """解析 run 实际生效的中文（eastAsia）字体名，取不到返回 None。"""
    for rpr in _iter_rpr_chain(doc, run):
        rfonts = rpr.find(qn('w:rFonts'))
        if rfonts is not None:
            val = rfonts.get(qn('w:eastAsia'))
            if val:
                return val
    return None


def _resolve_effective_ascii(doc, run):
    """解析 run 实际生效的西文（ascii）字体名，取不到返回 None。"""
    for rpr in _iter_rpr_chain(doc, run):
        rfonts = rpr.find(qn('w:rFonts'))
        if rfonts is not None:
            val = rfonts.get(qn('w:ascii'))
            if val:
                return val
    return None


def _resolve_effective_size_pt(doc, run):
    for rpr in _iter_rpr_chain(doc, run):
        sz = rpr.find(qn('w:sz'))
        if sz is not None:
            return _pt_from_sz(sz.get(qn('w:val')))
    return None


def _resolve_effective_bold(doc, run):
    for rpr in _iter_rpr_chain(doc, run):
        node = rpr.find(qn('w:b'))
        if node is not None:
            return _bool_of(node)
    return False

def _resolve_ppr_child(doc, p_el, tag):
    """段落属性同样按 直接 pPr -> 样式链 -> docDefaults 解析。"""
    ppr = p_el.find(qn('w:pPr'))
    if ppr is not None:
        node = ppr.find(qn(tag))
        if node is not None:
            return node
    for style in _iter_style_chain(doc, _para_style_id(doc, p_el)):
        sppr = style.find(qn('w:pPr'))
        if sppr is not None:
            node = sppr.find(qn(tag))
            if node is not None:
                return node
    dd = doc.styles.element.find(qn('w:docDefaults'))
    if dd is not None:
        ppr_default = dd.find(qn('w:pPrDefault'))
        if ppr_default is not None:
            sppr = ppr_default.find(qn('w:pPr'))
            if sppr is not None:
                node = sppr.find(qn(tag))
                if node is not None:
                    return node
    return None


def _resolve_ind_attr(doc, p_el, attr):
    """缩进按属性逐项继承：直接 pPr/w:ind 未写该属性时，继续沿样式链找同名属性。"""
    ppr = p_el.find(qn('w:pPr'))
    if ppr is not None:
        node = ppr.find(qn('w:ind'))
        if node is not None and node.get(qn(attr)) is not None:
            return _int_of(node.get(qn(attr)))
    for style in _iter_style_chain(doc, _para_style_id(doc, p_el)):
        sppr = style.find(qn('w:pPr'))
        if sppr is not None:
            node = sppr.find(qn('w:ind'))
            if node is not None and node.get(qn(attr)) is not None:
                return _int_of(node.get(qn(attr)))
    return None


_JC_ALIGN = {
    'center': WD_ALIGN_PARAGRAPH.CENTER,
    'both': WD_ALIGN_PARAGRAPH.JUSTIFY,
    'distribute': WD_ALIGN_PARAGRAPH.JUSTIFY,
    'left': WD_ALIGN_PARAGRAPH.LEFT,
    'start': WD_ALIGN_PARAGRAPH.LEFT,
    'right': WD_ALIGN_PARAGRAPH.RIGHT,
    'end': WD_ALIGN_PARAGRAPH.RIGHT,
}


def _resolve_align(doc, p_el):
    node = _resolve_ppr_child(doc, p_el, 'w:jc')
    if node is None:
        return None
    return _JC_ALIGN.get((node.get(qn('w:val')) or '').strip().lower())


def _resolve_spacing(doc, p_el):
    """返回 (beforeLines, afterLines, before, after, line, lineRule)。"""
    node = _resolve_ppr_child(doc, p_el, 'w:spacing')
    if node is None:
        return (None, None, None, None, None, None)
    return (
        _int_of(node.get(qn('w:beforeLines'))), _int_of(node.get(qn('w:afterLines'))),
        _int_of(node.get(qn('w:before'))), _int_of(node.get(qn('w:after'))),
        _int_of(node.get(qn('w:line'))), (node.get(qn('w:lineRule')) or '').strip().lower() or None,
    )

def _half_line_ok(before_lines):
    """段前/段后 0.5 行：必须以「行」为单位设置，即 beforeLines≈50。

    细则写的是「行」，磅值写法（只有 w:before 而无 w:beforeLines）不按等价折算处理：
    行距随字号变化，同一磅值在不同字号下并不等于半行。
    """
    return before_lines is not None and 40 <= before_lines <= 60


def _one_line_ok(before_lines):
    """段前/段后 1 行：必须以「行」为单位设置，即 beforeLines≈100。理由同 _half_line_ok。"""
    return before_lines is not None and 90 <= before_lines <= 110


def _single_line_ok(line, line_rule):
    """单倍行距：未设 line，或 line=240 且 lineRule 为 auto。"""
    if line is None:
        return line_rule in (None, 'auto')
    if line_rule in (None, 'auto'):
        return 230 <= line <= 250
    return False


def _has_bottom_border(doc, p_el):
    pbdr = _resolve_ppr_child(doc, p_el, 'w:pBdr')
    if pbdr is None:
        return False
    bottom = pbdr.find(qn('w:bottom'))
    if bottom is None:
        return False
    val = (bottom.get(qn('w:val')) or '').strip().lower()
    return bool(val) and val not in ('nil', 'none')


def _iter_picture_nodes(el):
    """产出真正承载位图的 w:drawing 节点，跳过 mc:Fallback 里的兼容副本。"""
    for node in el.iter(qn('w:drawing')):
        anc = node.getparent()
        skip = False
        while anc is not None:
            if anc.tag == MC_FALLBACK:
                skip = True
                break
            anc = anc.getparent()
        if not skip and node.find('.//' + qn('a:blip')) is not None:
            yield node


def _picture_geometry(doc, drawing_el):
    """返回 (显示宽EMU, 显示高EMU, 像素宽, 像素高)。"""
    cx = cy = px_w = px_h = None
    extent = drawing_el.find('.//' + qn('wp:extent'))
    if extent is not None:
        cx, cy = _int_of(extent.get('cx')), _int_of(extent.get('cy'))
    blip = drawing_el.find('.//' + qn('a:blip'))
    rid = blip.get(qn('r:embed')) if blip is not None else None
    if rid:
        part = doc.part.related_parts.get(rid)
        image = getattr(part, 'image', None)
        if image is not None:
            px_w, px_h = image.px_width, image.px_height
    return cx, cy, px_w, px_h

def _table_shape(tbl_el):
    """行数取 w:tr 个数；列数优先取 tblGrid/gridCol 个数，缺失时退化为首行单元格数。"""
    rows = tbl_el.findall(qn('w:tr'))
    grid = tbl_el.find(qn('w:tblGrid'))
    cols = len(grid.findall(qn('w:gridCol'))) if grid is not None else 0
    if not cols and rows:
        cols = len(rows[0].findall(qn('w:tc')))
    return len(rows), cols


def _is_floating_table(tbl_el):
    """浮动表（tblpPr 锚定）时 XML 兄弟次序不代表渲染次序。"""
    tbl_pr = tbl_el.find(qn('w:tblPr'))
    return tbl_pr is not None and tbl_pr.find(qn('w:tblpPr')) is not None


def _part_paragraphs(part):
    if part is None:
        return []
    try:
        return list(part.paragraphs)
    except Exception:
        return []


def _part_text(part):
    return _norm(''.join(p.text for p in _part_paragraphs(part)))


def _part_has_page_field(part):
    if part is None:
        return False
    try:
        xml = part._element.xml
    except Exception:
        return False
    if re.search(r'<w:instrText[^>]*>\s*PAGE', xml):
        return True
    return bool(re.search(r'w:instr="[^"]*PAGE', xml))


def _open_word_document(file_path):
    """打开被评测文档；打不开/需修复时由调用方捕获异常。"""
    return Document(file_path)

# ======================= 文档结构一次性解析 =======================

def _parse_structure(doc):
    items = [ch for ch in doc.element.body if ch.tag in (W_P, W_TBL)]
    kinds = ['tbl' if el.tag == W_TBL else 'p' for el in items]
    texts = [_el_text(el).strip() for el in items]

    # 封面区：以第一处写在段落 pPr 里的 sectPr 为界（封面独立成节时必然存在）
    cover_end = -1
    for i, el in enumerate(items):
        if el.tag == W_P:
            ppr = el.find(qn('w:pPr'))
            if ppr is not None and ppr.find(qn('w:sectPr')) is not None:
                cover_end = i
                break
    if cover_end < 0:  # 未分节：退化为最后一处年份元素
        for i, text in enumerate(texts):
            if _RE_YEAR_CELL.search(text) and '年' in text:
                cover_end = i

    # 一级 / 二级标题：只按编号识别，不校验译名（翻译具有不确定性）
    l1, l2 = {}, {}
    for i in range(cover_end + 1, len(items)):
        text = texts[i]
        if kinds[i] != 'p' or not text or len(text) > 40 or text.endswith('。'):
            continue
        m2 = _RE_L2.match(text)
        if m2:
            l2.setdefault((int(m2.group(1)), int(m2.group(2))), i)
            continue
        m1 = _RE_L1.match(text)
        if m1:
            num = int(m1.group(1))
            if 1 <= num <= 20:
                l1.setdefault(num, i)

    heading_idx = set(l1.values()) | set(l2.values())
    first_l1 = min(l1.values()) if l1 else cover_end + 1

    # 后置 4 项 + 参考文献标题
    back = {}
    refs_idx = None
    for i in range(cover_end + 1, len(items)):
        text = texts[i]
        if kinds[i] != 'p' or not text or len(text) > 20 or i in heading_idx:
            continue
        if refs_idx is None and _RE_REFS_HEAD.match(text):
            refs_idx = i
        for name, pattern in _BACK_PATTERNS:
            if name not in back and pattern.match(text):
                back[name] = i
    tails = [v for v in back.values()] + ([refs_idx] if refs_idx is not None else [])
    back_start = min(tails) if tails else len(items)

    # 题注、插图、表格（封面区的校徽与两张信息表不计入正文统计）
    captions = []
    for i in range(cover_end + 1, len(items)):
        if kinds[i] != 'p':
            continue
        m = _RE_CAPTION.match(texts[i])
        if m:
            captions.append({'index': i, 'kind': m.group(1), 'num': int(m.group(2)), 'text': texts[i]})
    caption_idx = set(c['index'] for c in captions)

    cover_pics, content_pics = [], []
    for i, el in enumerate(items):
        for node in _iter_picture_nodes(el):
            (cover_pics if i <= cover_end else content_pics).append({'index': i, 'node': node})

    cover_tables, content_tables = [], []
    for i, el in enumerate(items):
        if kinds[i] != 'tbl':
            continue
        rows, cols = _table_shape(el)
        rec = {'index': i, 'el': el, 'rows': rows, 'cols': cols, 'float': _is_floating_table(el)}
        (cover_tables if i <= cover_end else content_tables).append(rec)

    # 摘要 / 关键词（位于封面之后、第一个一级标题之前的前置区）
    abstract_idx = keyword_idx = None
    for i in range(cover_end + 1, min(first_l1, len(items))):
        if kinds[i] != 'p' or not texts[i]:
            continue
        if abstract_idx is None and _RE_ABSTRACT.match(texts[i]):
            abstract_idx = i
        if keyword_idx is None and _RE_KEYWORD.match(texts[i]):
            keyword_idx = i

    # 篇名：封面之后第一个非空且含中文的段落
    title_idx = None
    for i in range(cover_end + 1, min(first_l1, len(items))):
        if kinds[i] == 'p' and texts[i] and _has_cjk(texts[i]) and i not in (abstract_idx, keyword_idx):
            title_idx = i
            break

    # 正文段落：一级标题起、参考文献表止（后置 4 项下的中文正文同属正文），
    # 排除标题段、后置项无编号标题、题注段与纯图段
    body_stop = refs_idx if refs_idx is not None else len(items)
    back_heads = set(back.values())
    body_idx = [i for i in range(first_l1, body_stop)
                if kinds[i] == 'p' and texts[i] and i not in heading_idx
                and i not in caption_idx and i not in back_heads]

    return {
        'items': items, 'kinds': kinds, 'texts': texts, 'cover_end': cover_end,
        'l1': l1, 'l2': l2, 'heading_idx': heading_idx, 'first_l1': first_l1,
        'back': back, 'refs_idx': refs_idx, 'back_start': back_start,
        'captions': captions, 'caption_idx': caption_idx,
        'cover_pics': cover_pics, 'content_pics': content_pics,
        'cover_tables': cover_tables, 'content_tables': content_tables,
        'abstract_idx': abstract_idx, 'keyword_idx': keyword_idx,
        'title_idx': title_idx, 'body_idx': body_idx,
    }

# ======================= 对外接口 =======================

def evaluate(dir_path):
    result = {
        'id': SCRIPT_ID,
        'file_name': '',
        'status': 'ok',
        'error': None,
        'dim1_pass': True,
        'dim1_reason': '',
        'dim2_items': [],
        'total_score': 0,
        'max_score': MAX_SCORE_TOTAL,
    }
    try:
        if not dir_path or not os.path.isdir(dir_path):
            raise FileNotFoundError('目录不存在：%s' % dir_path)
        cands = sorted(
            f for f in os.listdir(dir_path)
            if f.lower().endswith('.docx') and not f.startswith('~$')
        )
        if not cands:
            raise FileNotFoundError('目录内未找到 .docx 文件：%s' % dir_path)
        file_name = cands[0]
        result['file_name'] = file_name
        _evaluate_impl(os.path.join(dir_path, file_name), result)
    except Exception as exc:
        # status 为 error 时未能完成判定，dim1_pass 留空，dim1_reason 也为空
        result['status'] = 'error'
        result['error'] = '%s: %s' % (type(exc).__name__, exc)
        result['dim1_pass'] = ''
        result['dim1_reason'] = ''
        result['dim2_items'] = []
        result['total_score'] = 0
    return result

def _bracket_unbalanced(text):
    """题注被拆成两段时，括号往往在断点处失配。"""
    for left, right in (('（', '）'), ('(', ')'), ('【', '】'), ('[', ']'), ('《', '》')):
        if text.count(left) != text.count(right):
            return True
    return False


def _evaluate_impl(file_path, result):
    _STYLE_CACHE.clear()
    # 维度1：可用与可修改性。打不开/需修复即整份完成度记 0，且不再进入维度2 的逐条判定。
    try:
        doc = _open_word_document(file_path)
    except Exception as exc:
        result['dim1_pass'] = False
        result['dim1_reason'] = '交付的 .docx 无法正常打开（%s: %s），完成度总分记0' % (
            type(exc).__name__, exc)
        result['dim2_items'] = []
        result['total_score'] = 0
        return
    # 维度1 通过时 dim1_reason 保持为空字符串，只在不通过时回填失败原因。
    info = _parse_structure(doc)
    items, kinds, texts = info['items'], info['kinds'], info['texts']
    sections = list(doc.sections)
    results = []

    def add(hit, delta, rule):
        results.append('+%d：%s' % (delta if hit else 0, rule))

    def runs_of(index_or_el):
        el = items[index_or_el] if isinstance(index_or_el, int) else index_or_el
        return [r for r in el.findall(qn('w:r')) if _el_text(r).strip()]

    def host_section(index):
        """给定元素下标，返回它所处的一级章节号。"""
        cur = None
        for num in sorted(info['l1']):
            if info['l1'][num] <= index:
                cur = num
        return cur

    def check_runs(el, pt, ea_allow=None, ascii_allow=None, need_bold=False, tol=0.6):
        """字号对全部 run 生效；中文字体只校验含汉字的 run，西文字体只校验含拉丁字母的 run。"""
        found = runs_of(el)
        if not found:
            return False
        for run in found:
            text = _el_text(run)
            if pt is not None and not _close(_resolve_effective_size_pt(doc, run), pt, tol):
                return False
            if ea_allow and _has_cjk(text):
                if not _font_hit(_resolve_effective_eastasia(doc, run), ea_allow):
                    return False
            if ascii_allow and re.search(r'[A-Za-z]', text):
                if not _font_hit(_resolve_effective_ascii(doc, run), ascii_allow):
                    return False
            if need_bold and not _resolve_effective_bold(doc, run):
                return False
        return True

    # ==================== 加分项（33 条 / 95 分）====================

    # 01 13 个一级章节标题：只认编号，不校验译名
    add(all(n in info['l1'] for n in EXPECTED_L1_NUMBERS), 3,
        '13个一级章节标题全部存在并保留原编号')

    # 02 24 个二级小节标题：1.1–12.2 全覆盖
    add(all(k in info['l2'] for k in EXPECTED_L2_KEYS), 5,
        '24个二级小节标题全部存在并保留原编号')

    # 03 37 个章节下均有中文正文；且正文里没有整段未译的英文
    heads = sorted([(i, 1) for i in info['l1'].values()] + [(i, 2) for i in info['l2'].values()])
    sec_ok = len(heads) >= EXPECTED_SECTION_TOTAL
    sec_seen = False
    for pos, (idx, _lvl) in enumerate(heads):
        stop = heads[pos + 1][0] if pos + 1 < len(heads) else info['back_start']
        sec_seen = True
        prose = False
        for j in range(idx + 1, stop):
            if kinds[j] == 'p' and texts[j] and j not in info['caption_idx'] and _has_cjk(texts[j]):
                prose = True
                break
        if not prose:
            sec_ok = False
            break
    if sec_ok:
        for j in info['body_idx']:
            if len(texts[j]) >= 40 and not _has_cjk(texts[j]):
                sec_ok = False
                break
    add(sec_ok and sec_seen, 5, '全部37个章节下均有中文正文，无空章节、无整段英文未译')

    # 04 摘要 + 5 个关键词
    abs_idx, kw_idx = info['abstract_idx'], info['keyword_idx']
    kw_count = 0
    if kw_idx is not None:
        tail = re.sub(r'^[【\[]?\s*关\s*键\s*词\s*[】\]]?\s*[:：]?\s*', '', texts[kw_idx])
        kw_count = len([w for w in re.split(r'[，,；;、/]+', tail) if w.strip()])
    add(abs_idx is not None and len(texts[abs_idx]) > 20 and _has_cjk(texts[abs_idx])
        and kw_count == EXPECTED_KEYWORD_COUNT, 3, '摘要与5个关键词完整存在')

    # 05 后置 4 项齐全，且沿用无编号形态（被续编成"14 致谢"时匹配不上，自然不得分）
    back_ok = len(info['back']) == 4
    if back_ok:
        for idx in info['back'].values():
            if re.match(r'^\d', texts[idx]):
                back_ok = False
                break
    add(back_ok, 3, '后置4项（作者贡献声明、利益冲突声明、数据可用性、致谢）齐全并沿用无编号形态')

    # 06 参考文献 [1]–[15] 齐全，且保留英文原文（模板要求"省略不翻"即不翻译）
    ref_nums = set()
    ref_en_ok = True
    if info['refs_idx'] is not None:
        for j in range(info['refs_idx'] + 1, len(items)):
            if kinds[j] != 'p':
                continue
            m = _RE_REF_ENTRY.match(texts[j])
            if m:
                ref_nums.add(int(m.group(1)))
                if not re.search(r'[A-Za-z]', texts[j]):
                    ref_en_ok = False
    add(all(n in ref_nums for n in EXPECTED_REF_NUMBERS) and ref_en_ok, 1,
        '参考文献[1]–[15]齐全并保留英文原文')

    # 07 5 张表格为真实表格对象且行列数与原文一致
    shapes = tuple((t['rows'], t['cols']) for t in info['content_tables'])
    add(shapes == EXPECTED_TABLE_SHAPES, 5,
        '5张表格为真实表格对象且行列数与原文一致（表1 7×3、表2 13×7、表3 7×4、表4 6×3、表5 7×4）')

    # 08 单元格无空缺、无合并丢失、无行列错位（逐行核对 gridSpan 之和是否等于列数）
    cell_ok = len(info['content_tables']) == len(EXPECTED_TABLE_SHAPES)
    for rec in info['content_tables']:
        if not cell_ok:
            break
        for tr in rec['el'].findall(qn('w:tr')):
            tcs = tr.findall(qn('w:tc'))
            span = 0
            for tc in tcs:
                tc_pr = tc.find(qn('w:tcPr'))
                gs = tc_pr.find(qn('w:gridSpan')) if tc_pr is not None else None
                span += _int_of(gs.get(qn('w:val')), 1) if gs is not None else 1
            if span != rec['cols']:
                cell_ok = False
                break
            for tc in tcs:
                if not _el_text(tc).strip():
                    cell_ok = False
                    break
            if not cell_ok:
                break
    add(cell_ok, 3, '表格单元格无空缺、无合并丢失、无行列错位')

    # 09 图 1–图 7 全部插入（封面校徽不计入）
    add(len(info['content_pics']) == EXPECTED_FIGURE_COUNT, 5, '图1–图7全部插入')

    # 10 插图落在原文对应的一级章节内
    pic_pos_ok = len(info['content_pics']) == EXPECTED_FIGURE_COUNT
    if pic_pos_ok:
        for seq, pic in enumerate(info['content_pics'], start=1):
            if host_section(pic['index']) != FIGURE_HOST_SECTION.get(seq):
                pic_pos_ok = False
                break
    add(pic_pos_ok, 3, '7张插图均位于原文对应章节位置')

    # 11 12 个题注：存在、编号连续、图题在下表题在上、未被拆成两段
    caps = info['captions']
    fig_caps = [c for c in caps if c['kind'] == '图']
    tbl_caps = [c for c in caps if c['kind'] == '表']
    cap_ok = (len(caps) == EXPECTED_CAPTION_COUNT
              and [c['num'] for c in fig_caps] == list(range(1, EXPECTED_FIGURE_COUNT + 1))
              and [c['num'] for c in tbl_caps] == list(range(1, len(EXPECTED_TABLE_SHAPES) + 1)))
    if cap_ok:
        pic_at = dict((seq, p['index']) for seq, p in enumerate(info['content_pics'], start=1))
        for cap in fig_caps:                       # 图题在图下方
            host = pic_at.get(cap['num'])
            if host is None or not 0 < cap['index'] - host <= 2:
                cap_ok = False
                break
    if cap_ok:
        tbl_at = dict((seq, t) for seq, t in enumerate(info['content_tables'], start=1))
        for cap in tbl_caps:                       # 表题在表上方
            rec = tbl_at.get(cap['num'])
            if rec is None:
                cap_ok = False
                break
            gap = rec['index'] - cap['index']
            # 浮动表锚定后 XML 兄弟次序不代表渲染次序，紧邻任一侧都按"表题在上"认可
            if rec['float']:
                if abs(gap) > 2:
                    cap_ok = False
                    break
            elif not 0 < gap <= 2:
                cap_ok = False
                break
    if cap_ok:
        for cap in caps:
            if _bracket_unbalanced(cap['text']):
                cap_ok = False
                break
            if cap['text'][-1] in _TERMINAL_CHARS:
                continue
            nxt = cap['index'] + 1                 # 断点后残留的短碎片
            if nxt < len(items) and kinds[nxt] == 'p' and texts[nxt] and len(texts[nxt]) <= 20 \
                    and nxt not in info['heading_idx'] and nxt not in info['caption_idx']:
                cap_ok = False
                break
    add(cap_ok, 3, '全部12个题注存在，位置正确（图题在下、表题在上）、编号正确、未被拆成两段')

    # 12 正文段落划分与原文一致：段末无终止标点却紧接另一个正文段，即视为拆段断词
    body_set = set(info['body_idx'])
    split_ok = bool(info['body_idx'])
    for idx in info['body_idx']:
        if texts[idx][-1] in _TERMINAL_CHARS:
            continue
        if idx + 1 in body_set:
            split_ok = False
            break
    add(split_ok, 1, '正文段落划分与原文一致，无拆段断词')

    # 13 原文页眉/期刊名/DOI/页码等非正文信息未被搬入（参考文献条目保留英文，不参与该项）
    artifact_ok = True
    for i, kind in enumerate(kinds):
        if kind != 'p':
            continue
        text = texts[i]
        if not text or _RE_REF_ENTRY.match(text):
            continue
        if re.match(r'^\d{1,4}$', text) and i not in info['caption_idx']:
            artifact_ok = False
            break
        if any(rx.search(text) for rx in _FOREIGN_ARTIFACTS):
            artifact_ok = False
            break
    add(artifact_ok, 1, '无原文页眉、期刊名、DOI、页码等非正文信息被搬入')

    # 14 002 模板占位文字与格式说明文字全部清除（含表格单元格）
    residue_ok = True
    for text in texts:
        if text and any(token in text for token in _TEMPLATE_RESIDUE):
            residue_ok = False
            break
    add(residue_ok, 1, '002模板占位文字与格式说明文字全部清除')

    # 15 插图比例未失真、不超版心、清晰
    usable_emu = None
    if sections:
        last = sections[-1]
        if None not in (last.page_width, last.left_margin, last.right_margin):
            usable_emu = int(last.page_width) - int(last.left_margin) - int(last.right_margin)
    pic_q_ok = len(info['content_pics']) == EXPECTED_FIGURE_COUNT
    for pic in info['content_pics']:
        if not pic_q_ok:
            break
        cx, cy, px_w, px_h = _picture_geometry(doc, pic['node'])
        if not cx or not cy or not px_w or not px_h:
            pic_q_ok = False
            break
        if not math.isclose(cx / float(cy), px_w / float(px_h), rel_tol=0.02):
            pic_q_ok = False          # 比例失真（横竖向缩放不一致）
            break
        if usable_emu and cx > usable_emu * 1.02:
            pic_q_ok = False          # 超出版心宽度
            break
        if px_w / (cx / 914400.0) < 120.0:
            pic_q_ok = False          # 有效分辨率过低，打印发虚
            break
    add(pic_q_ok, 3, '7张插图比例未失真、不超版心、清晰、边界干净')

    # 16 客户原始输入文件保持原样：向上最多三级查找"客户"目录并逐个校验可打开性
    src_dir = None
    probe = os.path.dirname(os.path.abspath(file_path))
    for _ in range(3):
        cand = os.path.join(probe, '客户')
        if os.path.isdir(cand):
            src_dir = cand
            break
        parent = os.path.dirname(probe)
        if parent == probe:
            break
        probe = parent
    input_ok = True
    if src_dir is not None:
        for name in sorted(os.listdir(src_dir)):
            path = os.path.join(src_dir, name)
            if name.startswith('~$') or not os.path.isfile(path):
                continue
            low = name.lower()
            try:
                if os.path.getsize(path) <= 0:
                    input_ok = False
                elif low.endswith('.docx'):
                    Document(path)
                elif low.endswith('.pdf'):
                    with open(path, 'rb') as handle:
                        if handle.read(4) != b'%PDF':
                            input_ok = False
            except Exception:
                input_ok = False
            if not input_ok:
                break
    add(input_ok, 1, '客户原始输入文件保持原样未被改动')

    # 17 A4 + 四边 2.5cm
    page_ok = bool(sections)
    for sect in sections:
        if not (_close(_emu_to_cm(sect.page_width), A4_WIDTH_CM, 0.2)
                and _close(_emu_to_cm(sect.page_height), A4_HEIGHT_CM, 0.2)):
            page_ok = False
            break
        for margin in (sect.top_margin, sect.right_margin, sect.bottom_margin, sect.left_margin):
            if not _close(_emu_to_cm(margin), MARGIN_CM, 0.06):
                page_ok = False
                break
        if not page_ok:
            break
    add(page_ok, 3, '页面为A4且四边页边距均为2.5cm')

    # 18 页眉页脚距边界 1.5cm，且启用文档网格
    def grid_of(sect):
        node = sect._sectPr.find(qn('w:docGrid'))
        if node is None:
            return (None, None)
        return ((node.get(qn('w:type')) or '').strip().lower(), _int_of(node.get(qn('w:linePitch'))))

    dist_ok = bool(sections)
    for sect in sections:
        if not (_close(_emu_to_cm(sect.header_distance), HEADER_FOOTER_DIST_CM, 0.06)
                and _close(_emu_to_cm(sect.footer_distance), HEADER_FOOTER_DIST_CM, 0.06)):
            dist_ok = False
            break
        gtype, pitch = grid_of(sect)
        if gtype not in ('lines', 'linesandchars', 'snaptochars') or not pitch:
            dist_ok = False
            break
    add(dist_ok, 1, '页眉距边界1.5cm、页脚距边界1.5cm，且启用文档网格')

    # 19 封面固定元素完整复刻：两行标题 + 7行2列信息表 + 表下方校徽 + 1行2列年份表
    cover_end = info['cover_end']

    def find_cover_title(target, fonts, pt, need_bold):
        for i in range(0, cover_end + 1):
            if kinds[i] != 'p' or _norm(texts[i]) != _norm(target):
                continue
            if _resolve_align(doc, items[i]) != WD_ALIGN_PARAGRAPH.CENTER:
                continue
            ok = bool(runs_of(i))
            for run in runs_of(i):
                ea = _resolve_effective_eastasia(doc, run)
                asc = _resolve_effective_ascii(doc, run)
                if not (_font_hit(ea, fonts) or _font_hit(asc, fonts)):
                    ok = False
                    break
                if not _close(_resolve_effective_size_pt(doc, run), pt, 0.6):
                    ok = False
                    break
                if need_bold and not _resolve_effective_bold(doc, run):
                    ok = False
                    break
            if ok:
                return i
        return None

    main_i = find_cover_title(COVER_TITLE_MAIN, STRICT_YAHEI, PT_COVER_MAIN, False)
    sub_i = find_cover_title(COVER_TITLE_SUB, STRICT_DENGXIAN, PT_COVER_SUB, True)
    info_tbl = next((t for t in info['cover_tables']
                     if (t['rows'], t['cols']) == COVER_INFO_SHAPE), None)
    year_tbl = next((t for t in info['cover_tables']
                     if (t['rows'], t['cols']) == COVER_YEAR_SHAPE
                     and _RE_YEAR_CELL.search(texts[t['index']]) and '年' in texts[t['index']]), None)
    logo = None
    if info_tbl is not None:
        logo = next((p for p in info['cover_pics'] if p['index'] > info_tbl['index']), None)
    add(None not in (main_i, sub_i, info_tbl, year_tbl, logo)
        and main_i < sub_i < info_tbl['index'] < logo['index'] < year_tbl['index'], 5,
        '封面固定元素完整复刻（「毕业设计（论文）」微软雅黑42pt居中、「外文翻译」等线22pt加粗居中、'
        '7行2列信息表、校徽图位于信息表下方、「二〇二六 | 年」1行2列表）')

    # 20 封面题目栏填入翻译后的中文题名（不限信息表行数，只看"题目"这一行的取值）
    title_cell_ok = False
    for rec in info['cover_tables']:
        for tr in rec['el'].findall(qn('w:tr')):
            tcs = tr.findall(qn('w:tc'))
            if len(tcs) < 2 or _norm(_el_text(tcs[0])) != '题目':
                continue
            value = _el_text(tcs[1]).strip()
            title_cell_ok = bool(value) and _has_cjk(value) and COVER_TITLE_PLACEHOLDER not in value
            break
        if title_cell_ok:
            break
    add(title_cell_ok, 1, '封面题目栏填入翻译后的中文题名')

    # 21 封面独立成节，且该节页眉页脚（含首页版本）均为空、无 PAGE 页码域
    cover_clean = len(sections) >= 2
    if cover_clean:
        first = sections[0]
        for part in (first.header, first.even_page_header, first.first_page_header,
                     first.footer, first.even_page_footer, first.first_page_footer):
            if _part_text(part):
                cover_clean = False
                break
    if cover_clean:
        first = sections[0]
        for part in (first.footer, first.even_page_footer, first.first_page_footer):
            if _part_has_page_field(part):
                cover_clean = False
                break
    add(cover_clean, 3, '封面独立成节，无页眉、无页脚、无页码')

    # 22 奇偶页页眉分设，均居中且带下框线
    even_odd = _bool_of(doc.settings.element.find(qn('w:evenAndOddHeaders'))) is True
    body_sect = sections[-1] if sections else None
    hdr_ok = even_odd and body_sect is not None
    if hdr_ok:
        hdr_ok = (_part_text(body_sect.header) == _norm(ODD_HEADER_TEXT)
                  and _part_text(body_sect.even_page_header) == _norm(EVEN_HEADER_TEXT))
    if hdr_ok:
        for part in (body_sect.header, body_sect.even_page_header):
            paras = [p for p in _part_paragraphs(part) if p.text.strip()]
            if not paras:
                hdr_ok = False
                break
            for para in paras:
                if _resolve_align(doc, para._p) != WD_ALIGN_PARAGRAPH.CENTER \
                        or not _has_bottom_border(doc, para._p):
                    hdr_ok = False
                    break
            if not hdr_ok:
                break
    add(hdr_ok, 5, '奇偶页页眉分设——奇数页「外文翻译」、偶数页「2022届材料科学与工程专业毕业设计（论文）」，'
                   '均居中且带下框线')

    # 23 页脚为居中 PAGE 页码域，且正文首页页码重编为 1
    def footer_page_centered(part):
        if not _part_has_page_field(part):
            return False
        for para in _part_paragraphs(part):
            if 'PAGE' in para._p.xml and _resolve_align(doc, para._p) == WD_ALIGN_PARAGRAPH.CENTER:
                return True
        # 页码放在浮动文本框里时，段落本身可能是左对齐，要看文本框自身是否水平居中
        for anchor in part._element.iter(qn('wp:anchor')):
            instr = ''.join(node.text or '' for node in anchor.iter(qn('w:instrText')))
            if 'PAGE' not in instr:
                continue
            pos_h = anchor.find(qn('wp:positionH'))
            align = pos_h.find(qn('wp:align')) if pos_h is not None else None
            if align is not None and (align.text or '').strip().lower() == 'center':
                return True
            for inner in anchor.iter(W_P):
                if 'PAGE' in inner.xml and _resolve_align(doc, inner) == WD_ALIGN_PARAGRAPH.CENTER:
                    return True
        return False

    ftr_ok = body_sect is not None and footer_page_centered(body_sect.footer)
    if ftr_ok and even_odd and _part_has_page_field(body_sect.even_page_footer):
        # 启用奇偶页且偶数页页脚另有定义时，偶数页页码也必须居中
        ftr_ok = footer_page_centered(body_sect.even_page_footer)
    start_ok = False
    for sect in sections:
        pg_num = sect._sectPr.find(qn('w:pgNumType'))
        if pg_num is not None and _int_of(pg_num.get(qn('w:start'))) == 1:
            start_ok = True
            break
    add(ftr_ok and start_ok, 3, '页脚为居中PAGE页码域，且正文首页页码为1')

    # 24 篇名二号黑体居中，段前 0.5 行、段后 0.5 行
    t_idx = info['title_idx']
    title_ok = t_idx is not None
    if title_ok:
        before_l, after_l, _before_t, _after_t, _line, _rule = _resolve_spacing(doc, items[t_idx])
        title_ok = (_resolve_align(doc, items[t_idx]) == WD_ALIGN_PARAGRAPH.CENTER
                    and _half_line_ok(before_l) and _half_line_ok(after_l)
                    and check_runs(t_idx, PT_TITLE, ea_allow=STRICT_HEITI))
    add(title_ok, 3, '篇名二号黑体居中，段前0.5行、段后0.5行')

    # 25 摘要为五号宋体 + Times New Roman
    add(abs_idx is not None
        and check_runs(abs_idx, PT_ABSTRACT, ea_allow=STRICT_SONGTI, ascii_allow=STRICT_TNR, tol=0.3),
        1, '摘要为五号宋体+Times New Roman')

    # 26 「关键词」标签加粗，词条本身不加粗
    kw_bold_ok = kw_idx is not None
    if kw_bold_ok:
        full = _el_text(items[kw_idx])
        match = re.search(r'[:：]', full)
        colon = match.start() if match else len('关键词')
        pos, label_seen, term_seen = 0, False, False
        for run in items[kw_idx].findall(qn('w:r')):
            text = _el_text(run)
            if not text:
                continue
            start, end = pos, pos + len(text)
            pos = end
            bold = _resolve_effective_bold(doc, run)
            if end <= colon:                                  # 整段落在标签内
                label_seen = True
                if not bold:
                    kw_bold_ok = False
                    break
            elif start >= colon + 1:                          # 整段落在词条内
                term_seen = True
                if bold:
                    kw_bold_ok = False
                    break
            elif not (start == colon and end == colon + 1):    # 标签与词条共用一个 run，无法分设加粗
                kw_bold_ok = False
                break
        kw_bold_ok = kw_bold_ok and label_seen and term_seen
    add(kw_bold_ok, 1, '「关键词」标签加粗，词条本身不加粗')

    # 27 摘要末段后（＝关键词前）空一行、关键词后空一行
    def is_blank(i):
        return 0 <= i < len(items) and kinds[i] == 'p' and not texts[i]

    blank_ok = abs_idx is not None and kw_idx is not None and kw_idx > abs_idx
    if blank_ok:
        blank_ok = any(is_blank(i) for i in range(abs_idx + 1, kw_idx)) and is_blank(kw_idx + 1)
    add(blank_ok, 1, '摘要末段后、关键词前、关键词后共三处空行齐全')

    # 28 13 个一级标题：三号黑体、居左、段前后各 1 行、单倍行距
    h1_ok = len(info['l1']) == len(EXPECTED_L1_NUMBERS)
    for num in sorted(info['l1']):
        if not h1_ok:
            break
        idx = info['l1'][num]
        before_l, after_l, _before_t, _after_t, line, rule = _resolve_spacing(doc, items[idx])
        if _resolve_align(doc, items[idx]) != WD_ALIGN_PARAGRAPH.LEFT:
            h1_ok = False
            break
        if not (_one_line_ok(before_l) and _one_line_ok(after_l)
                and _single_line_ok(line, rule)):
            h1_ok = False
            break
        if not check_runs(idx, PT_H1, ea_allow=STRICT_HEITI):
            h1_ok = False
            break
    add(h1_ok, 5, '全部13个一级标题为三号黑体居左、段前后各1行、单倍行距')

    # 29 24 个二级标题：小三号黑体 + TNR、两端对齐、段前后各 0.5 行、节号与标题间空 1 格
    h2_ok = len(info['l2']) == len(EXPECTED_L2_KEYS)
    for key in sorted(info['l2']):
        if not h2_ok:
            break
        idx = info['l2'][key]
        match = _RE_L2.match(texts[idx])
        if match is None or match.group(3) not in (' ', '　', '\xa0'):
            h2_ok = False
            break
        if _resolve_align(doc, items[idx]) != WD_ALIGN_PARAGRAPH.JUSTIFY:
            h2_ok = False
            break
        before_l, after_l, _before_t, _after_t, _line, _rule = _resolve_spacing(doc, items[idx])
        if not (_half_line_ok(before_l) and _half_line_ok(after_l)):
            h2_ok = False
            break
        if not check_runs(idx, PT_H2, ea_allow=STRICT_HEITI, ascii_allow=STRICT_TNR):
            h2_ok = False
            break
    add(h2_ok, 5, '全部24个二级标题为小三号黑体+Times New Roman、两端对齐、段前后各0.5行、节号与标题间空1格')

    # 30 37 个标题的大纲级别（一级=级别1 即 outlineLvl 0，二级=级别2 即 outlineLvl 1）
    outline_ok = (len(info['l1']) == len(EXPECTED_L1_NUMBERS)
                  and len(info['l2']) == len(EXPECTED_L2_KEYS))
    for level, mapping in ((0, info['l1']), (1, info['l2'])):
        if not outline_ok:
            break
        for idx in mapping.values():
            node = _resolve_ppr_child(doc, items[idx], 'w:outlineLvl')
            if node is None or _int_of(node.get(qn('w:val'))) != level:
                outline_ok = False
                break
    add(outline_ok, 3, '全部37个标题设有正确大纲级别（一级=级别1，二级=级别2）')

    # 31 正文小四号，中文宋体、西文 Times New Roman（正文＝一级标题起至后置项前，排除标题与题注）
    body_font_ok = bool(info['body_idx'])
    for idx in info['body_idx']:
        if not check_runs(idx, PT_BODY, ea_allow=STRICT_SONGTI, ascii_allow=STRICT_TNR, tol=0.3):
            body_font_ok = False
            break
    add(body_font_ok, 3, '正文为小四号，中文宋体、西文Times New Roman')

    # 32 正文单倍行距且勾选"对齐到文档网格"（无 docGrid 时该复选框无从生效，直接不得分）
    grid_type, grid_pitch = grid_of(body_sect) if body_sect is not None else (None, None)
    snap_ok = (bool(info['body_idx']) and grid_type in ('lines', 'linesandchars', 'snaptochars')
               and bool(grid_pitch))
    for idx in info['body_idx']:
        if not snap_ok:
            break
        node = _resolve_ppr_child(doc, items[idx], 'w:snapToGrid')
        if node is not None and _bool_of(node) is False:
            snap_ok = False
            break
        _bl, _al, _bt, _at, line, rule = _resolve_spacing(doc, items[idx])
        if not _single_line_ok(line, rule):
            snap_ok = False
            break
    add(snap_ok, 3, '正文为单倍行距且已勾选「对齐到文档网格」')

    # 33 正文首行缩进 0.85cm
    # 字符写法（firstLineChars）只有折算成实际长度后落在 0.85cm 才算，避免字号不同时
    # "2 字符"被无条件当成 0.85cm：五号下 2 字符只有 0.74cm。
    indent_ok = bool(info['body_idx'])
    for idx in info['body_idx']:
        if not indent_ok:
            break
        if _resolve_ind_attr(doc, items[idx], 'w:hanging') or \
                _resolve_ind_attr(doc, items[idx], 'w:hangingChars'):
            indent_ok = False
            break
        chars = _resolve_ind_attr(doc, items[idx], 'w:firstLineChars')
        twips = _resolve_ind_attr(doc, items[idx], 'w:firstLine')
        if chars is not None:
            runs = runs_of(idx)
            size_pt = _resolve_effective_size_pt(doc, runs[0]) if runs else None
            if size_pt and abs(chars / 100.0 * size_pt * 20.0 - FIRST_LINE_TWIPS) <= 30:
                continue
        if twips is not None and abs(twips - FIRST_LINE_TWIPS) <= 30:
            continue
        indent_ok = False
        break
    add(indent_ok, 3, '正文首行缩进0.85cm')

    # ==================== 汇总 ====================

    if len(results) != len(_MAX_DELTAS_ORDER):
        raise RuntimeError('检查项数量与 _MAX_DELTAS_ORDER 不一致: %d vs %d'
                           % (len(results), len(_MAX_DELTAS_ORDER)))

    dim2_items = []
    total_score = 0
    for raw, max_delta in zip(results, _MAX_DELTAS_ORDER):
        head, _, rule = raw.partition('：')
        try:
            delta = int(head.replace('+', '').strip())
        except ValueError:
            delta = 0
        hit = delta > 0
        dim2_items.append({
            'rule': rule,
            'max_delta': max_delta,
            'delta': max_delta if hit else 0,
            'hit': hit,
        })
        total_score += max_delta if hit else 0

    result['dim2_items'] = dim2_items
    result['total_score'] = total_score
    result['max_score'] = MAX_SCORE_TOTAL
    return result


if __name__ == '__main__':
    _dir = sys.argv[1] if len(sys.argv) >= 2 else os.path.dirname(os.path.abspath(__file__))
    print(json.dumps(evaluate(_dir), ensure_ascii=False, indent=2))
