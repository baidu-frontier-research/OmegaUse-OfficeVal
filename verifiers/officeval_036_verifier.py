#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# pyright: reportUnknownMemberType=false, reportUnknownVariableType=false, reportUnknownArgumentType=false, reportUnknownParameterType=false, reportMissingParameterType=false, reportPrivateUsage=false, reportUnknownLambdaType=false, reportGeneralTypeIssues=false

import os
import re
import sys
import json
import zipfile
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn

# 说明：按接口统一约定，本模块不修改全局 sys.stdout，也不 print 主结果。
# 主结果统一由 evaluate(dir_path) -> dict 返回；入参为“脚本所在目录的路径”，
# 由脚本自己在该目录内定位并打开被评估的 .docx（接受 .doc 后缀但仍须为 OOXML）。
# 本文件仅在 __main__ 下作为本地调试入口，才把 dict 序列化为 JSON 打印。
#
# 与 008 号脚本的差异：本任务（把 PDF 对账单收支记录整理进 Word 明细表并统一记录
# 行高）的细则同时含“加分项”（指令要求的记录搬运 + 行高统一 + 与既有版式对齐）与
# “减分项”（指令未涉及却被多余改动的第 1 页明细表 / 第 2 页表头 / 页面设置 / 标题条 /
# 说明框 / 插图 / 页眉页脚等）。故 dim2 项的 max_delta 允许为负：正值=加分项（命中记
# 满分，未命中记 0）；负值=减分项（触发“破坏”记该负分，未触发记 0）。total_score 为
# 全部 delta 之和，取值区间 [-10, +41]；max_score 仍为所有加分项满分之和（+41），
# 作为完成度上限。
#
# 判定口径（细则原文）：
#   ① 列宽以 Word 实际呈现（w:tblGrid/w:gridCol）为准，不以 w:tcW 单元格首选宽度为准
#      ——交付前第 2 页表头写有 1080/4896/1800/1584/1440 缇的 tcW，但排版时并不生效。
#   ② “记录行高一致”以行高的规则与数值设置是否统一为准，不要求各行在页面上的实际显示
#      高度相同（摘要长短不同导致的折行属正常排版）。
#   ③ 凡以缇 / 磅 / EMU 给出的数值允许 ±2 缇（及等效换算）的存盘舍入偏差。
#   ④ 与既有框线同色同粗、叠加后看不出变化的重复框线定义，以及 w:val="none" 的空边框
#      定义，均不计扣分。

SCRIPT_ID = "036"

# —— 加分项按“结果追加顺序”的 max_delta（正值）——
_ADD_DELTAS_ORDER = [
    5,  # A01 第 2 页明细表填入 48 条记录行，与 PDF 逐条对应，无遗漏无重复
    3,  # A02 记录按 01 Dec.→31 Dec. 排列，Date 列 48 行取值与 PDF 逐行一致
    5,  # A03 Description 列 48 条摘要与 PDF 逐条一致
    5,  # A04 支出只进 Withdrawals、收入只进 Deposits，两列不串位
    3,  # A05 无对应金额的单元格保持空白，不填 0 / - / N/A 占位
    5,  # A06 Balance 列 48 行余额与 PDF 逐行一致
    5,  # A07 48 条记录行全部显式设置行高，且行高数值完全相同
    3,  # A08 记录行 Arial 10.5 磅非加粗，Balance 列加粗
    1,  # A09 Date/Description 左对齐，Withdrawals/Deposits/Balance 右对齐
    3,  # A10 金额带 $ 前缀、两位小数、千位分隔符
    3,  # A11 记录行之间以细横线分隔，保持无左右外框线的通栏样式
]

# —— 减分项按“结果追加顺序”的 max_delta（负值）——
_DED_DELTAS_ORDER = [
    -5,  # D01 第 1 页五列明细表整表或其表头/记录行缺失（命中则 D02~D07 不再叠加）
    -1,  # D02 第 1 页表头 Arial 12.5 加粗 / 底纹 F2F7F7 被改
    -1,  # D03 第 1 页表头上边框 sz12 7A7A7A / 下边框 sz8 DDDDDD 被改
    -1,  # D04 第 1 页记录行下边框 sz6 D8D8D8 被改，或两记录行间新增可见横线
    -1,  # D05 第 1 页表同一行内混用不同垂直对齐，造成文字上下错位
    -1,  # D06 第 1 页两条记录的金额或正负号被改动
    -1,  # D07 第 1 页表不再呈现五列等宽 2086 缇，或整表宽度超出正文区 10430 缇
    -1,  # D08 第 2 页表头文字不满足 Arial 11.5 磅 加粗
    -1,  # D09 第 2 页表头底纹被改为深色 / 高饱和色
    -1,  # D10 第 2 页表头行上边框不满足 sz12 黑色单实线
    -1,  # D11 第 2 页表头行下框线被加为黑色粗线 / 双线 / 彩色线
    -1,  # D12 第 2 页表头单元格内边距不满足 上下 90 缇、左右 80 缇
    -1,  # D13 第 2 页五列列宽失衡，或整表宽度超出正文区 10430 缇
    -1,  # D14 第 2 页表头同一行内混用不同垂直对齐
    -1,  # D15 记录单元格内被插入多余的空文本片段或空段落
    -1,  # D16 记录行同一行内混用不同垂直对齐
    -1,  # D17 表头行行高数值与记录行行高数值不一致
    -3,  # D18 页面尺寸不满足 A4（11906 x 16838 缇）
    -3,  # D19 页边距不满足 上 1814 / 下 1304 / 左 737 / 右 737 缇
    -1,  # D20 页眉页脚距边界被改动，或新增页面上可见的页面边框
    -3,  # D21 顶部「Activity details — 出支记录」标题条缺失或文字被改写
    -1,  # D22 标题条 Arial 16 加粗 / 底纹 E8F2F1 / 无框线 / 内边距被改
    -3,  # D23 圆角矩形说明框缺失，或框内文字被改写
    -1,  # D24 说明框虚线边框 / 1A5D60 / 尺寸 / 宋体 12 磅被改
    -3,  # D25 第 2 页插图缺失、被替换或被压缩变形失真
    -1,  # D26 插图尺寸 / 四周型环绕 / 垂直偏移 -3336290 EMU 被改
    -1,  # D27 「Activity details」小标题不满足 Arial 13.5 磅 加粗
    -1,  # D28 分页符被删除，或明细表前用于占位的手动换行符被删除
    -3,  # D29 页眉「Income and Expense Record」文字缺失或被改写
    -3,  # D30 页脚「Page X of Y」页码域缺失或被改为静态文字
    -3,  # D31 新增 PDF 中不存在的记录行 / 小计行 / 说明行
]

MAX_SCORE_TOTAL = sum(_ADD_DELTAS_ORDER)               # +41：完成度上限
_MIN_TOTAL = MAX_SCORE_TOTAL + sum(_DED_DELTAS_ORDER)  # -10：全部命中且全部踩坑
assert MAX_SCORE_TOTAL == 41
assert sum(_DED_DELTAS_ORDER) == -51

# 结果追加顺序 = 先加分项、后减分项
_MAX_DELTAS_ORDER = _ADD_DELTAS_ORDER + _DED_DELTAS_ORDER
_ADD_COUNT = len(_ADD_DELTAS_ORDER)

# —— PDF 对账单 12 月收支记录基准（48 条，逐条硬编码）——
# 元组含义：(Date, Description, Withdrawals, Deposits, Balance)；金额已去掉 $ 前缀。
PDF_ROWS = (
    ('01 Dec.', 'Balance forward', '', '', '43,712.48'),
    ('01 Dec.', 'Insurance premium - studio toolkits', '218.40', '', '43,494.08'),
    ('01 Dec.', 'Supplier payout - compact label printers', '184.75', '', '43,309.33'),
    ('02 Dec.', 'Merchant batch - bamboo organizer trays', '', '6,870.20', '50,179.53'),
    ('02 Dec.', 'Merchant batch - LED task lamps', '', '7,940.00', '58,119.53'),
    ('02 Dec.', 'Merchant batch - ceramic storage jars', '', '2,215.35', '60,334.88'),
    ('03 Dec.', 'Client transfer - folding display stands', '', '1,135.70', '61,470.58'),
    ('05 Dec.', 'Marketplace settlement - wool desk mats', '', '695.42', '62,166.00'),
    ('09 Dec.', 'Platform fee - travel cable pouches', '63.25', '', '62,102.75'),
    ('10 Dec.', 'Client transfer - glass water carafes', '', '3,420.00', '65,522.75'),
    ('11 Dec.', 'Utilities payment - workshop power', '562.90', '', '64,959.85'),
    ('12 Dec.', 'Payment gateway batch - recycled notebooks', '', '11,780.50', '76,740.35'),
    ('15 Dec.', 'Client transfer - countertop spice racks', '', '2,640.00', '79,380.35'),
    ('15 Dec.', 'Cheque cleared - portable shelving units', '3,496.80', '', '75,883.55'),
    ('15 Dec.', 'ATM cash deposit - weekend sales', '', '7,530.00', '83,413.55'),
    ('15 Dec.', 'Remote cheque deposit - acoustic wall panels', '', '615.90', '84,029.45'),
    ('15 Dec.', 'Cheque cleared - magnetic whiteboards', '2,745.60', '', '81,283.85'),
    ('15 Dec.', 'Bill payment - warehouse heat', '684.30', '', '80,599.55'),
    ('15 Dec.', 'Cheque cleared - washable floor runners', '2,430.00', '', '78,169.55'),
    ('15 Dec.', 'Cheque cleared - adjustable coat racks', '3,215.45', '', '74,954.10'),
    ('15 Dec.', 'Cheque cleared - kraft mailer cartons', '1,955.20', '', '72,998.90'),
    ('15 Dec.', 'Cheque cleared - countertop scales', '3,118.70', '', '69,880.20'),
    ('17 Dec.', 'Supplier transfer - insulated lunch tins', '1,022.50', '', '68,857.70'),
    ('17 Dec.', 'Supplier transfer - silicone cable clips', '742.80', '', '68,114.90'),
    ('17 Dec.', 'Supplier transfer - walnut monitor risers', '1,604.40', '', '66,510.50'),
    ('18 Dec.', 'Supplier transfer - linen storage bins', '1,885.00', '', '64,625.50'),
    ('18 Dec.', 'Cheque cleared - reusable shipping crates', '2,875.35', '', '61,750.15'),
    ('19 Dec.', 'Remote cheque deposit - desk organizers', '', '2,940.75', '64,690.90'),
    ('23 Dec.', 'Payment gateway batch - clip-on reading lights', '', '8,210.60', '72,901.50'),
    ('24 Dec.', 'Wholesale payment - cotton tote sets', '', '5,476.25', '78,377.75'),
    ('24 Dec.', 'ATM deposit - cash sales', '', '6,800.00', '85,177.75'),
    ('29 Dec.', 'Client transfer - enamel plant tags', '', '2,125.00', '87,302.75'),
    ('29 Dec.', 'Client transfer - cork message boards', '', '3,480.00', '90,782.75'),
    ('29 Dec.', 'Marketplace settlement - soft-close storage boxes', '', '2,870.65', '93,653.40'),
    ('29 Dec.', 'Marketplace settlement - acrylic recipe holders', '', '1,120.30', '94,773.70'),
    ('30 Dec.', 'Payment gateway batch - glass spice mills', '', '4,435.40', '99,209.10'),
    ('30 Dec.', 'Cheque cleared - collapsible laundry carts', '9,750.00', '', '89,459.10'),
    ('30 Dec.', 'Platform fee - leather key trays', '483.60', '', '88,975.50'),
    ('30 Dec.', 'Platform fee - brushed metal hooks', '37.25', '', '88,938.25'),
    ('30 Dec.', 'Platform fee - canvas tool rolls', '52.18', '', '88,886.07'),
    ('31 Dec.', 'Cheque cleared - cedar drawer inserts', '3,120.70', '', '85,765.37'),
    ('31 Dec.', 'Cheque cleared - pegboard panels', '2,948.30', '', '82,817.07'),
    ('31 Dec.', 'Cheque cleared - wire baskets', '2,104.90', '', '80,712.17'),
    ('31 Dec.', 'Cheque cleared - rolling trolleys', '3,302.75', '', '77,409.42'),
    ('31 Dec.', 'Cheque cleared - cable trays', '3,005.45', '', '74,403.97'),
    ('31 Dec.', 'Cheque cleared - bamboo file holders', '3,260.20', '', '71,143.77'),
    ('31 Dec.', 'Monthly service fee - operating account', '21.00', '', '71,122.77'),
    ('31 Dec.', 'Closing totals', '54,891.73', '82,302.02', '71,122.77'),
)

PDF_KEYS = tuple((r[0], r[1]) for r in PDF_ROWS)
PDF_DESCS = frozenset(r[1] for r in PDF_ROWS)
N_PDF = len(PDF_ROWS)  # 48

# —— 交付前文档（输入文件）的版式基准 ——
HDR_LABELS = ('date', 'description', 'withdrawals', 'deposits', 'balance')

# 顶部标题条
BANNER_TEXT = 'Activity details — 出支记录'
BANNER_FILL = 'E8F2F1'
BANNER_FONT = 'Arial'
BANNER_SIZE = 16.0
BANNER_MAR = {'top': 480, 'bottom': 420, 'left': 420, 'right': 360}

# 第 1 页五列明细表
T1_HDR_SIZE = 12.5
T1_HDR_FILL = 'F2F7F7'
T1_HDR_TOP = ('single', 12, '7A7A7A')
T1_HDR_BOTTOM = ('single', 8, 'DDDDDD')
T1_REC_BOTTOM = ('single', 6, 'D8D8D8')
T1_AMOUNTS = ('$8,732.40', '$9,016.55', '-$6,918.75', '$2,097.80')
T1_DATE_KEYS = ('Mar 22', 'Mar 25')

# 第 2 页五列明细表（表头）
T2_HDR_SIZE = 11.5
T2_HDR_TOP_SZ = 12
T2_HDR_MAR = {'top': 90, 'bottom': 90, 'left': 80, 'right': 80}
GRID_EVEN = 2086          # 五列等宽（Word 实际呈现值）
GRID_DESIGN = (1080, 4896, 1800, 1584, 1440)   # 为让摘要单行容纳而加宽的比例
TEXT_WIDTH_TW = 10430     # 正文区宽度上限

# 记录行
REC_FONT = 'Arial'
REC_SIZE = 10.5

# 页面设置
PG_SZ = (11906, 16838)
PG_MAR = {'top': 1814, 'bottom': 1304, 'left': 737, 'right': 737}
PG_HDR_DIST = 0
PG_FTR_DIST = 624

# 圆角矩形说明框
NOTE_TEXT = 'Keep daily financial transactions in order.'
NOTE_LINE_CLR = '1A5D60'
NOTE_EXTENT = (1079500, 3397250)
NOTE_EA_FONT = ('宋体', 'SimSun')
NOTE_SIZE = 12.0

# 第 2 页插图
PIC_EXTENT = (5848350, 3665220)
PIC_POS_V = -3336290
PIC_RATIO = PIC_EXTENT[0] / float(PIC_EXTENT[1])

# 小标题 / 页眉 / 页脚
SUB_TITLE = 'Activity details'
SUB_TITLE_SIZE = 13.5
HEADER_TEXT = 'Income and Expense Record'

TW_TOL = 2            # ±2 缇
EMU_PER_TW = 635      # 1 缇 = 635 EMU
EMU_TOL = TW_TOL * EMU_PER_TW


# ==================== 通用工具 ====================

def _int(val):
    """把 XML 属性值转成 int；无法解析时返回 None。"""
    if val is None:
        return None
    try:
        return int(round(float(str(val).strip())))
    except (TypeError, ValueError):
        return None


def _near(got, want, tol=TW_TOL):
    """数值在容差内视为相等（None 一律视为不相等）。"""
    if got is None or want is None:
        return False
    return abs(got - want) <= tol


def _norm_ws(text):
    """压缩空白：把连续空白（含不换行空格、制表符、换行）折叠为单个空格。"""
    if not text:
        return ''
    text = text.replace(' ', ' ').replace('　', ' ')
    return re.sub(r'\s+', ' ', text).strip()


def _norm_dash(text):
    """把各种连字符/短横线统一为 ASCII '-'，避免因 ‐ ‑ – — 的差异误判摘要。"""
    return re.sub(r'[‐‑‒–—−]', '-', text)


def _norm_money(text):
    """金额归一：去掉 $ 、千位分隔符、正负号与空白，仅保留数字与小数点。"""
    s = _norm_ws(text).replace('$', '').replace(',', '').replace('＄', '')
    s = s.replace('-', '').replace('−', '').replace('+', '').strip()
    return s


_PLACEHOLDERS = ('0', '0.00', '$0', '$0.00', '-', '--', '—', '/', 'n/a',
                 'na', 'nil', 'none', '无', '不适用')


def _is_placeholder(text):
    """是否填了 0 / - / N/A 一类占位符。"""
    return _norm_ws(text).lower() in _PLACEHOLDERS


# ==================== OOXML 取值 ====================

def _tcpr(cell):
    return cell._tc.find(qn('w:tcPr'))


def _cell_fill(cell):
    """单元格底纹填充色（大写）；无底纹或 auto 返回 None。"""
    pr = _tcpr(cell)
    if pr is None:
        return None
    shd = pr.find(qn('w:shd'))
    if shd is None:
        return None
    fill = shd.get(qn('w:fill'))
    if fill is None or fill.lower() in ('auto',):
        return None
    return fill.upper()


def _cell_valign(cell):
    """单元格垂直对齐；未设置按 OOXML 默认的顶端对齐（top）处理，
    以免“显式写 top”与“沿用默认”这种页面上看不出差别的写法被判为混用。"""
    pr = _tcpr(cell)
    if pr is None:
        return 'top'
    va = pr.find(qn('w:vAlign'))
    if va is None:
        return 'top'
    return (va.get(qn('w:val')) or 'top')


def _cell_mar(cell):
    """单元格内边距（缇）；未单独设置的边返回 None。
    兼容以 w:start / w:end 代替 w:left / w:right 的写法。"""
    out = {'top': None, 'bottom': None, 'left': None, 'right': None}
    pr = _tcpr(cell)
    if pr is None:
        return out
    mar = pr.find(qn('w:tcMar'))
    if mar is None:
        return out
    for side in out:
        node = mar.find(qn('w:' + side))
        if node is not None:
            out[side] = _int(node.get(qn('w:w')))
    for side, alias in (('left', 'start'), ('right', 'end')):
        if out[side] is None:
            node = mar.find(qn('w:' + alias))
            if node is not None:
                out[side] = _int(node.get(qn('w:w')))
    return out


def _border_tuple(node):
    """把边框节点读成 (val, sz, color)；color 归一为大写，auto 视为 None。"""
    if node is None:
        return None
    val = node.get(qn('w:val'))
    sz = _int(node.get(qn('w:sz')))
    clr = node.get(qn('w:color'))
    if clr is not None:
        clr = None if clr.lower() == 'auto' else clr.upper()
    return (val, sz, clr)
def _cell_border(cell, side):
    """单元格自身 tcBorders 上某一边的定义。"""
    pr = _tcpr(cell)
    if pr is None:
        return None
    bd = pr.find(qn('w:tcBorders'))
    if bd is None:
        return None
    return _border_tuple(bd.find(qn('w:' + side)))


def _tbl_border(tbl, side):
    """表级 tblBorders 上某一边的定义（含 insideH / insideV）。"""
    pr = tbl._tbl.find(qn('w:tblPr'))
    if pr is None:
        return None
    bd = pr.find(qn('w:tblBorders'))
    if bd is None:
        return None
    return _border_tuple(bd.find(qn('w:' + side)))


def _visible(bd):
    """框线是否在页面上可见：val 非 nil/none 且线宽 > 0。"""
    if bd is None:
        return False
    val, sz, _clr = bd
    if val is None or val.lower() in ('nil', 'none'):
        return False
    return (sz or 0) > 0


def _grid_cols(tbl):
    """w:tblGrid/w:gridCol 列宽（缇）——Word 实际呈现的列宽即以此为准。"""
    grid = tbl._tbl.find(qn('w:tblGrid'))
    if grid is None:
        return []
    return [_int(gc.get(qn('w:w'))) or 0 for gc in grid.findall(qn('w:gridCol'))]


def _row_height(row):
    """行高设置：返回 (数值缇, 规则)；未显式设置返回 (None, None)。
    OOXML 中 w:trHeight 省略 w:hRule 时默认按“最小值”(atLeast) 处理。"""
    tr_pr = row._tr.find(qn('w:trPr'))
    if tr_pr is None:
        return (None, None)
    th = tr_pr.find(qn('w:trHeight'))
    if th is None:
        return (None, None)
    rule = th.get(qn('w:hRule')) or 'atLeast'
    return (_int(th.get(qn('w:val'))), rule)


# ==================== 字体属性的有效值解析 ====================

def _style_chain(style, limit=12):
    """样式继承链：自身 -> basedOn -> ...（防御性限制层数，避免循环引用死循环）。"""
    out = []
    cur = style
    while cur is not None and len(out) < limit:
        out.append(cur)
        try:
            cur = cur.base_style
        except (AttributeError, KeyError):
            break
    return out


def _rpr_nodes(doc, para, run):
    """按“直接格式 -> 字符样式 -> 段落样式继承链 -> 文档默认”的顺序给出 rPr 节点。"""
    nodes = []
    r_pr = run._r.find(qn('w:rPr'))
    if r_pr is not None:
        nodes.append(r_pr)
    try:
        cs = run.style
    except (AttributeError, KeyError):
        cs = None
    for st in _style_chain(cs):
        node = st.element.find(qn('w:rPr'))
        if node is not None:
            nodes.append(node)
    try:
        ps = para.style
    except (AttributeError, KeyError):
        ps = None
    for st in _style_chain(ps):
        node = st.element.find(qn('w:rPr'))
        if node is not None:
            nodes.append(node)
    dd = doc.styles.element.find(qn('w:docDefaults'))
    if dd is not None:
        rd = dd.find(qn('w:rPrDefault'))
        if rd is not None:
            node = rd.find(qn('w:rPr'))
            if node is not None:
                nodes.append(node)
    return nodes
def _eff_size_pt(doc, para, run):
    """有效字号（磅）；解析不到返回 None。"""
    for node in _rpr_nodes(doc, para, run):
        sz = node.find(qn('w:sz'))
        if sz is not None:
            half = _int(sz.get(qn('w:val')))
            if half:
                return half / 2.0
    return None


def _eff_font(doc, para, run, attr='ascii'):
    """有效西文/中文字体名；解析不到返回 None。"""
    for node in _rpr_nodes(doc, para, run):
        rf = node.find(qn('w:rFonts'))
        if rf is not None:
            name = rf.get(qn('w:' + attr))
            if name:
                return name
    return None


def _eff_bold(doc, para, run):
    """有效加粗；解析不到按未加粗处理。"""
    for node in _rpr_nodes(doc, para, run):
        b = node.find(qn('w:b'))
        if b is not None:
            val = b.get(qn('w:val'))
            return val is None or str(val).lower() not in ('0', 'false', 'off')
    return False


def _text_runs(cell):
    """单元格内所有含可见文字的 (paragraph, run)。"""
    out = []
    for para in cell.paragraphs:
        for run in para.runs:
            if _norm_ws(run.text) != '':
                out.append((para, run))
    return out


def _cell_text(cell):
    """单元格文字（多段落以空格相连，空白折叠）。"""
    return _norm_ws('\n'.join(p.text for p in cell.paragraphs))
def _eff_jc(doc, para):
    """有效段落对齐；返回 left / center / right / both 等字符串，未设置返回 None。"""
    p_pr = para._p.find(qn('w:pPr'))
    if p_pr is not None:
        jc = p_pr.find(qn('w:jc'))
        if jc is not None and jc.get(qn('w:val')):
            return jc.get(qn('w:val')).lower()
    try:
        ps = para.style
    except (AttributeError, KeyError):
        ps = None
    for st in _style_chain(ps):
        node = st.element.find(qn('w:pPr'))
        if node is not None:
            jc = node.find(qn('w:jc'))
            if jc is not None and jc.get(qn('w:val')):
                return jc.get(qn('w:val')).lower()
    return None


def _cell_jc(doc, cell):
    """单元格内首个含文字段落的对齐（无文字时取首段）。"""
    for para in cell.paragraphs:
        if _norm_ws(para.text) != '':
            return _eff_jc(doc, para)
    if cell.paragraphs:
        return _eff_jc(doc, cell.paragraphs[0])
    return None


def _est_lines(text, col_tw, mar_l, mar_r, size_pt):
    """按列宽与字号粗估折行数（仅用于维度 1 判断“固定值行高是否把文字压没”）。"""
    text = _norm_ws(text)
    if text == '':
        return 1
    usable_pt = (float(col_tw or 0) - (mar_l or 0) - (mar_r or 0)) / 20.0
    if usable_pt <= 0 or not size_pt:
        return 1
    char_pt = 0.52 * size_pt          # Arial 小写字母平均字宽约为字号的 0.52 倍
    per_line = max(1, int(usable_pt / char_pt))
    return max(1, -(-len(text) // per_line))
# ==================== 表格定位 ====================

def _row_cells(row):
    try:
        return list(row.cells)
    except (IndexError, ValueError):
        return []


def _is_detail_table(tbl):
    """是否为 Date/Description/Withdrawals/Deposits/Balance 五列明细表。"""
    if not tbl.rows:
        return False
    cells = _row_cells(tbl.rows[0])
    if len(cells) != 5:
        return False
    texts = [_norm_ws(c.text).lower() for c in cells]
    return all(HDR_LABELS[i] in texts[i] for i in range(5))


def _col_dates(tbl):
    """明细表除表头外各行的 Date 列文字。"""
    out = []
    for row in tbl.rows[1:]:
        cells = _row_cells(row)
        out.append(_norm_ws(cells[0].text) if cells else '')
    return out


def _locate_tables(doc):
    """定位标题条表、第 1 页五列明细表、第 2 页（本次要填写的）五列明细表。"""
    details = [t for t in doc.tables if _is_detail_table(t)]
    banner = None
    for tbl in doc.tables:
        if tbl in details or not tbl.rows:
            continue
        txt = _norm_ws(' '.join(c.text for c in _row_cells(tbl.rows[0])))
        if 'Activity details' in txt and '出支记录' in txt:
            banner = tbl
            break
    if banner is None:
        for tbl in doc.tables:
            if tbl in details or not tbl.rows:
                continue
            cells = _row_cells(tbl.rows[0])
            if cells and _cell_fill(cells[0]) == BANNER_FILL:
                banner = tbl
                break

    t1 = None
    for tbl in details:
        if any('Mar' in d for d in _col_dates(tbl)):
            t1 = tbl
            break
    rest = [t for t in details if t is not t1]
    if rest:
        t2 = max(rest, key=lambda t: len(t.rows))
    elif t1 is not None and any('Dec' in d for d in _col_dates(t1)):
        t2 = t1                      # 12 月记录被并入了第 1 页那张表
    elif details:
        t2 = max(details, key=lambda t: len(t.rows))
    else:
        t2 = None
    return banner, t1, t2
# ==================== 图形 / 页眉页脚 / 节属性 ====================

def _shape_segments(xml):
    """正文中的所有图形片段（DrawingML 的 w:drawing 与 VML 的 w:pict）。"""
    segs = re.findall(r'<w:drawing[ >].*?</w:drawing>', xml, re.S)
    segs += re.findall(r'<w:pict[ >].*?</w:pict>', xml, re.S)
    return segs


def _seg_texts(seg):
    """图形片段内的文字（按出现顺序）。"""
    return re.findall(r'<w:t(?:\s[^>]*)?>(.*?)</w:t>', seg, re.S)


def _seg_extent(seg):
    """图形片段的显示尺寸（EMU）；取不到返回 (None, None)。"""
    m = re.search(r'<wp:extent\s+cx="(-?\d+)"\s+cy="(-?\d+)"', seg)
    if m:
        return (int(m.group(1)), int(m.group(2)))
    m = re.search(r'<a:ext\s+cx="(-?\d+)"\s+cy="(-?\d+)"', seg)
    if m:
        return (int(m.group(1)), int(m.group(2)))
    return (None, None)


def _seg_pos(seg, axis='V'):
    """图形片段的锚点偏移（EMU）；取不到返回 None。"""
    m = re.search(r'<wp:position%s\b[^>]*>\s*<wp:posOffset>(-?\d+)' % axis, seg, re.S)
    return int(m.group(1)) if m else None


def _seg_wrap(seg):
    """环绕方式：返回 wp:wrap* 的类型名集合，如 {'Square'} / {'None'}。"""
    return set(re.findall(r'<wp:wrap([A-Za-z]+)', seg))


def _find_note_seg(segs):
    for seg in segs:
        if 'roundRect' in seg or 'roundrect' in seg:
            return seg
    return None


def _find_pic_seg(segs):
    for seg in segs:
        if '<a:blip' in seg or 'imagedata' in seg:
            return seg
    return None
def _hf_texts(doc, kind):
    """页眉('header')或页脚('footer')部件内的文字与域代码。"""
    texts, instrs = [], []
    try:
        parts = list(doc.part.package.iter_parts())
    except AttributeError:
        parts = []
    for part in parts:
        name = str(getattr(part, 'partname', '')).lower()
        if kind not in name or not name.endswith('.xml'):
            continue
        try:
            blob = part.blob.decode('utf-8', 'ignore')
        except Exception:
            continue
        texts += [_norm_ws(t) for t in re.findall(r'<w:t(?:\s[^>]*)?>(.*?)</w:t>', blob, re.S)]
        instrs += [_norm_ws(t) for t in
                   re.findall(r'<w:instrText(?:\s[^>]*)?>(.*?)</w:instrText>', blob, re.S)]
    return texts, instrs


def _emu_to_tw(val):
    return None if val is None else int(round(float(val) / EMU_PER_TW))


def _sect_geometry(doc):
    """首节的页面尺寸 / 页边距 / 页眉页脚距边界（均折算为缇）。"""
    sect = doc.sections[0]
    geo = {
        'w': _emu_to_tw(sect.page_width),
        'h': _emu_to_tw(sect.page_height),
        'top': _emu_to_tw(sect.top_margin),
        'bottom': _emu_to_tw(sect.bottom_margin),
        'left': _emu_to_tw(sect.left_margin),
        'right': _emu_to_tw(sect.right_margin),
        'hdr': _emu_to_tw(sect.header_distance),
        'ftr': _emu_to_tw(sect.footer_distance),
    }
    return geo


def _has_visible_pg_borders(doc):
    """是否新增了在页面上可见的页面边框（w:val 为 none/nil 的空定义不算）。"""
    for sect in doc.sections:
        pg = sect._sectPr.find(qn('w:pgBorders'))
        if pg is None:
            continue
        for child in pg:
            if _visible(_border_tuple(child)):
                return True
    return False
def _rgb(clr):
    if not clr or not re.fullmatch(r'[0-9A-Fa-f]{6}', clr):
        return None
    return (int(clr[0:2], 16), int(clr[2:4], 16), int(clr[4:6], 16))


def _is_light_fill(fill):
    """底纹是否仍为白色或与 F2F7F7 同族的极浅色（无底纹/auto 视为白色）。"""
    if fill is None or fill.upper() == 'FFFFFF':
        return True
    rgb = _rgb(fill)
    if rgb is None:
        return False
    return min(rgb) >= 0xDC and (max(rgb) - min(rgb)) <= 0x28


def _is_light_gray(clr):
    """框线颜色是否为浅灰（近中性色且明度足够高）。"""
    if clr is None:
        return False          # auto 即黑色
    rgb = _rgb(clr)
    if rgb is None:
        return False
    return (max(rgb) - min(rgb)) <= 0x14 and min(rgb) >= 0xB0


def _is_black(clr):
    """框线颜色是否为黑色（auto 亦按黑色处理）。"""
    if clr is None:
        return True
    rgb = _rgb(clr)
    return rgb is not None and max(rgb) <= 0x20


def _mar_ok(mar, want):
    """内边距四边是否都落在基准值 ±2 缇内。"""
    return all(_near(mar.get(k), want[k]) for k in want)


def _fail_dim1(result, reason):
    """维度 1 未通过：完成度不再计分，total_score 归 0。

    dim1_reason 统一带上「维度一未通过：…」前缀与计分后果，便于直接作为错误摘要展示。
    """
    result["dim1_pass"] = False
    result["dim1_reason"] = "维度一未通过：%s，完成度总分记 0" % reason
    result["dim2_items"] = []
    result["total_score"] = 0
    return result


# 交付物只接受 .docx：细则要求交付「1 个可正常打开的 Word 文档（.docx）」，其余
# 格式（.doc / .docm / .rtf / .odt / .pdf 等）一律在维度一判不通过。
_ACCEPTED_EXT = ".docx"
# 目录里没有 .docx 时，这些扩展名仍收作被评对象，交给维度一判不通过（总分记 0），
# 而不是当成「没有交付物」直接跳过评测。
_WRONG_FMT_EXTS = (".doc", ".docm", ".dotx", ".dotm", ".dot", ".rtf", ".odt",
                   ".wps", ".pages", ".pdf")


def _fmt_reject(file_path):
    """维度 1 的格式门禁：交付物必须是可正常打开的 .docx（OOXML 文档包）。

    不通过返回原因文本，通过返回 None。判定与 Word 的「能不能双击打开」对齐：
    扩展名须为 .docx，文件须是 zip 容器，且包内含 word/document.xml。
    """
    ext = os.path.splitext(file_path)[1].lower()
    if ext != _ACCEPTED_EXT:
        return "交付物为 %s 格式，只接受 .docx 文档" % (ext or "无扩展名")
    if not zipfile.is_zipfile(file_path):
        return "交付物不是有效的 .docx 包（无法作为 zip 容器打开）"
    try:
        with zipfile.ZipFile(file_path) as zf:
            names = set(zf.namelist())
    except Exception as exc:                                  # noqa: BLE001
        return "交付物无法作为 .docx 打开：%s: %s" % (type(exc).__name__, exc)
    if "word/document.xml" not in names:
        return "交付物包内缺少 word/document.xml，不是 Word 文档"
    return None
# ==================== 对外入口 ====================

def evaluate(dir_path):
    """入参为目录路径；在该目录内定位被评估的 Word 文档并返回评分结果 dict。

    交付物只接受 .docx：目录内优先取 .docx；若只交了 .doc / .rtf / .odt / .pdf 等
    其他格式，仍把它收作被评对象，由维度一判不通过（total_score 记 0），而不是当成
    「没有交付物」跳过评测。

    返回固定 9 个字段：id / file_name / status / error / dim1_pass / dim1_reason /
    dim2_items / total_score / max_score，其中 dim2_items 每项为
    {rule, max_delta, delta, hit}。

    两类异常分开回传：
      · 找不到目录 / 目录内没有被评估文件——维度一根本没进入判定，error 只带纯文本
        原因（不带异常类名），file_name、dim1_pass、dim1_reason 一律留空（dim1_pass
        为空字符串而非 False），dim2 不出项，总分与满分一并归零；
      · 其他运行期异常——以「脚本运行异常：」打头，dim1_pass 记 False 并在
        dim1_reason 中给出错误摘要，max_score 仍保留真实满分。
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
        files = sorted(f for f in os.listdir(dir_path)
                       if not f.startswith('~$') and not f.startswith('.')
                       and os.path.isfile(os.path.join(dir_path, f)))
        cands = [f for f in files if f.lower().endswith(_ACCEPTED_EXT)]
        if not cands:
            # 没有 .docx，但交了 .doc/.rtf/.odt/.pdf 等其他格式：仍收作被评对象，
            # 由维度一判不通过，而不是当成「没有交付物」跳过评测。
            cands = [f for f in files if f.lower().endswith(_WRONG_FMT_EXTS)]
        if not cands:
            raise FileNotFoundError("目录中未找到 .docx 文件：%s" % dir_path)
        file_name = cands[0]
        result["file_name"] = file_name
        _evaluate_impl(os.path.join(dir_path, file_name), result)
    except FileNotFoundError as exc:
        result["status"] = "error"
        result["error"] = str(exc)
        result["file_name"] = ""
        result["dim1_pass"] = ""
        result["dim1_reason"] = ""
        result["dim2_items"] = []
        result["total_score"] = 0
        result["max_score"] = 0
    except Exception as exc:                                  # noqa: BLE001
        result["status"] = "error"
        result["error"] = "脚本运行异常：%s: %s" % (type(exc).__name__, exc)
        result["dim1_pass"] = False
        result["dim1_reason"] = ("文档无法打开或解析：%s: %s"
                                 % (type(exc).__name__, exc))
        result["dim2_items"] = []
        result["total_score"] = 0
    return result
def _norm_date(text):
    """日期归一：'01 Dec.' / '1 Dec' / '01 dec.' 统一为 (1, 'dec')。"""
    s = _norm_ws(text)
    m = re.match(r'^(\d{1,2})\s*([A-Za-z]{3})', s)
    if m:
        return (int(m.group(1)), m.group(2).lower())
    return (0, s.lower())


def _norm_desc(text):
    """摘要归一：统一短横线、折叠空白、忽略大小写。"""
    return _norm_dash(_norm_ws(text)).lower()


def _eff_tbl_border(doc, tbl, side):
    """表级框线的有效值：tblPr/tblBorders 优先，其次表格样式（含继承链）。"""
    bd = _tbl_border(tbl, side)
    if bd is not None:
        return bd
    pr = tbl._tbl.find(qn('w:tblPr'))
    style_id = None
    if pr is not None:
        node = pr.find(qn('w:tblStyle'))
        if node is not None:
            style_id = node.get(qn('w:val'))
    if not style_id:
        return None
    for st_el in doc.styles.element.findall(qn('w:style')):
        if st_el.get(qn('w:styleId')) != style_id:
            continue
        t_pr = st_el.find(qn('w:tblPr'))
        if t_pr is None:
            return None
        bs = t_pr.find(qn('w:tblBorders'))
        if bs is None:
            return None
        return _border_tuple(bs.find(qn('w:' + side)))
    return None


def _eff_cell_border(doc, tbl, cell, side):
    """单元格某一边的有效框线：单元格自身定义优先于表级 / 表格样式定义。"""
    bd = _cell_border(cell, side)
    if bd is not None:
        return bd
    return _eff_tbl_border(doc, tbl, side)


def _sep_line(doc, tbl, row, next_row, col=1):
    """记录行之间的分隔横线：本行下框线 / 下一行上框线 / 表级 insideH 任一生效即可。"""
    cands = []
    cells = _row_cells(row)
    if len(cells) > col:
        cands.append(_cell_border(cells[col], 'bottom'))
    if next_row is not None:
        ncells = _row_cells(next_row)
        if len(ncells) > col:
            cands.append(_cell_border(ncells[col], 'top'))
    cands.append(_eff_tbl_border(doc, tbl, 'insideH'))
    for bd in cands:
        if _visible(bd):
            return bd
    return None
# ==================== 维度 1 / 维度 2 主体 ====================

def _evaluate_impl(file_path, result):
    # —— 维度 1 的第一道门禁：交付物格式只接受 .docx，不合格即判 0，不再解析内容 ——
    reject = _fmt_reject(file_path)
    if reject:
        return _fail_dim1(result, reject)

    doc = Document(file_path)
    xml = doc.element.xml
    banner, t1, t2 = _locate_tables(doc)

    # —— 维度 1：可用与可修改性 ——
    if t2 is None:
        return _fail_dim1(result, "文档内找不到 Date/Description/Withdrawals/Deposits/Balance "
                                 "五列明细表：收支记录未以可编辑表格承载（表格缺失或被整页图片、"
                                 "截图、文本框拼贴替代）")
    rec_rows = list(t2.rows[1:])
    if len(rec_rows) == 0:
        return _fail_dim1(result, "第 2 页明细表只有表头、没有任何记录行，指令要求的收支记录未落入文档")

    grid2 = _grid_cols(t2)
    clipped = []
    for idx, row in enumerate(rec_rows):
        h_val, h_rule = _row_height(row)
        if h_val is None or h_rule != 'exact':
            continue
        cells = _row_cells(row)
        for ci, cell in enumerate(cells[:5]):
            txt = _cell_text(cell)
            if txt == '':
                continue
            runs = _text_runs(cell)
            size = _eff_size_pt(doc, runs[0][0], runs[0][1]) if runs else REC_SIZE
            mar = _cell_mar(cell)
            col_tw = grid2[ci] if ci < len(grid2) else 0
            lines = _est_lines(txt, col_tw, mar['left'] or 0, mar['right'] or 0, size)
            need = lines * (size or REC_SIZE) * 1.15 * 20
            avail = h_val - (mar['top'] or 0) - (mar['bottom'] or 0)
            if need > avail * 1.15:
                clipped.append(idx)
                break
    if len(clipped) >= max(3, int(0.2 * len(rec_rows))):
        return _fail_dim1(result, "记录行采用「固定值」行高且高度不足，共 %d/%d 行的摘要文字被行高"
                                 "裁切压没、在页面上读不全" % (len(clipped), len(rec_rows)))

    results = []
    result["dim1_pass"] = True          # 三项维度一硬门槛全部通过，才进入完成度计分
    _score_add(doc, t2, rec_rows, grid2, results)
    _score_ded(doc, xml, banner, t1, t2, rec_rows, grid2, results)

    if len(results) != len(_MAX_DELTAS_ORDER):
        raise RuntimeError("dim2 项数不一致：results=%d vs order=%d"
                           % (len(results), len(_MAX_DELTAS_ORDER)))
    dim2_items = []
    total = 0
    for raw, max_delta in zip(results, _MAX_DELTAS_ORDER):
        head, _, rule = raw.partition("：")
        got = int(head.replace("+", "").replace("-", "").strip())
        hit = got > 0
        delta = max_delta if hit else 0
        dim2_items.append({"rule": rule, "max_delta": max_delta,
                           "delta": delta, "hit": hit})
        total += delta
    result["dim2_items"] = dim2_items
    result["total_score"] = total
    return result
# ==================== 加分项 A01 ~ A14 ====================

def _score_add(doc, t2, rec_rows, grid2, results):
    cells_of = [_row_cells(r) for r in rec_rows]
    texts = [[_norm_ws(c.text) for c in cs[:5]] + [''] * max(0, 5 - len(cs))
             for cs in cells_of]
    n_rec = len(rec_rows)

    keys = [(_norm_date(t[0]), _norm_desc(t[1])) for t in texts]
    pdf_keys = [(_norm_date(p[0]), _norm_desc(p[1])) for p in PDF_ROWS]

    # A01 记录条数与内容逐条对应，无遗漏无重复
    a01 = (n_rec == N_PDF and sorted(keys) == sorted(pdf_keys))
    results.append(("+5" if a01 else "+0") +
                   "：第 2 页明细表填入 48 条收支记录行，与 PDF 对账单逐条对应，"
                   "既无遗漏也无重复（实测 %d 行）" % n_rec)

    # 逐行配对：行数一致时按序号对齐，否则按 (日期, 摘要) 回填配对
    if n_rec == N_PDF:
        pairs = list(zip(texts, PDF_ROWS))
    else:
        pairs, used = [], set()
        for t, k in zip(texts, keys):
            for j, pk in enumerate(pdf_keys):
                if j not in used and pk == k:
                    pairs.append((t, PDF_ROWS[j]))
                    used.add(j)
                    break
    full = (len(pairs) == N_PDF)

    # A02 排列顺序与 Date 列
    a02 = (n_rec == N_PDF and
           all(_norm_date(t[0]) == _norm_date(p[0]) for t, p in zip(texts, PDF_ROWS)))
    results.append(("+3" if a02 else "+0") +
                   "：记录按 01 Dec. 到 31 Dec. 的先后顺序排列，Date 列 48 行取值与 PDF 逐行一致")

    # A03 Description 列
    a03 = full and all(_norm_desc(t[1]) == _norm_desc(p[1]) for t, p in pairs)
    results.append(("+5" if a03 else "+0") +
                   "：Description 列 48 条摘要与 PDF 逐条一致，无错字、无截断、无自行改写")
    # A04 支出/收入不串位
    a04 = full
    for t, p in pairs:
        if p[2] and _norm_money(t[2]) != _norm_money(p[2]):
            a04 = False
        if p[3] and _norm_money(t[3]) != _norm_money(p[3]):
            a04 = False
    results.append(("+5" if a04 else "+0") +
                   "：支出金额只填在 Withdrawals 列、收入金额只填在 Deposits 列，两列不串位")

    # A05 空单元格保持空白
    a05 = bool(pairs)
    for t, p in pairs:
        for ci in (2, 3):
            if not p[ci] and (t[ci] != '' or _is_placeholder(t[ci])):
                a05 = False
    results.append(("+3" if a05 else "+0") +
                   "：该行没有对应金额的单元格保持空白，不填 0、-、N/A 等占位内容")

    # A06 Balance 列
    a06 = full and all(_norm_money(t[4]) == _norm_money(p[4]) for t, p in pairs)
    results.append(("+5" if a06 else "+0") +
                   "：Balance 列 48 行余额与 PDF 逐行一致")

    # A07 行高全部显式设置且数值一致
    heights = [_row_height(r) for r in rec_rows]
    h_vals = [h[0] for h in heights]
    a07 = (n_rec == N_PDF and all(v is not None for v in h_vals)
           and (max(h_vals) - min(h_vals)) <= TW_TOL)
    results.append(("+5" if a07 else "+0") +
                   "：48 条记录行全部显式设置了行高，且 48 行的行高数值完全相同"
                   "（如统一 0.89 厘米 / 505 缇），不出现部分行有设置、部分行沿用默认，"
                   "或数值一行一变（实测 %s）"
                   % ('/'.join(str(v) for v in sorted(set(h_vals), key=lambda x: (x is None, x))[:4])
                      + ' 缇' if h_vals else '无'))

    # A08 记录行字体字号与 Balance 列加粗
    a08 = bool(rec_rows)
    for cs in cells_of:
        for ci, cell in enumerate(cs[:5]):
            for para, run in _text_runs(cell):
                size = _eff_size_pt(doc, para, run)
                font = _eff_font(doc, para, run, 'ascii')
                bold = _eff_bold(doc, para, run)
                if font != REC_FONT or size is None or abs(size - REC_SIZE) > 0.25:
                    a08 = False
                if ci == 4 and not bold:
                    a08 = False
                if ci != 4 and bold:
                    a08 = False
    results.append(("+3" if a08 else "+0") +
                   "：记录行文字统一为 Arial 10.5 磅非加粗，仅 Balance 列加粗，"
                   "与第 1 页明细表记录行的字体字号一致")
    # A09 左右对齐分工
    a09 = bool(rec_rows)
    for cs in cells_of:
        for ci, cell in enumerate(cs[:5]):
            if _cell_text(cell) == '':
                continue
            jc = _cell_jc(doc, cell)
            if ci in (0, 1) and jc not in (None, 'left', 'start'):
                a09 = False
            if ci in (2, 3, 4) and jc not in ('right', 'end'):
                a09 = False
    results.append(("+1" if a09 else "+0") +
                   "：Date、Description 两列左对齐，Withdrawals、Deposits、Balance 三列右对齐，"
                   "与第 1 页明细表的对齐方式一致")

    # A10 金额格式
    a10 = bool(rec_rows)
    money_re = re.compile(r'^-?\$\d{1,3}(?:,\d{3})*\.\d{2}$')
    for cs in cells_of:
        for ci, cell in enumerate(cs[:5]):
            if ci < 2:
                continue
            txt = _cell_text(cell)
            if txt == '':
                continue
            if not money_re.match(_norm_dash(txt).replace(' ', '')):
                a10 = False
    results.append(("+3" if a10 else "+0") +
                   "：金额沿用 $ 前缀 + 两位小数 + 千位分隔符的写法（如 $6,870.20），"
                   "与第 1 页明细表的金额写法一致")

    # A11 细横线分隔 + 无左右外框线的通栏样式
    a11 = bool(rec_rows)
    for i, row in enumerate(rec_rows):
        nxt = rec_rows[i + 1] if i + 1 < len(rec_rows) else None
        if nxt is not None:
            bd = _sep_line(doc, t2, row, nxt)
            if bd is None or (bd[1] or 0) > 8:
                a11 = False
    for cs in cells_of:
        if len(cs) >= 5:
            if _visible(_eff_cell_border(doc, t2, cs[0], 'left')) or \
                    _visible(_eff_cell_border(doc, t2, cs[4], 'right')):
                a11 = False
    if _visible(_eff_tbl_border(doc, t2, 'insideV')):
        a11 = False
    results.append(("+3" if a11 else "+0") +
                   "：记录行之间以细横线分隔，且左右两侧不加竖线与外框线，"
                   "保持与第 1 页明细表相同的通栏表格样式")
EQ_TOL = 20           # 判断“等宽”时的视觉容差（20 缇约合 0.035 厘米，页面上看不出差别）


def _is_even_grid(grid):
    """五列是否仍呈现等宽 2086 缇的版式。"""
    if len(grid) != 5:
        return False
    return (max(grid) - min(grid)) <= EQ_TOL and all(abs(g - GRID_EVEN) <= EQ_TOL for g in grid)


def _is_design_grid(grid):
    """五列是否按 1080/4896/1800/1584/1440 的比例加宽了 Description 列。"""
    if len(grid) != 5:
        return False
    total = float(sum(grid))
    base = float(sum(GRID_DESIGN))
    if total <= 0:
        return False
    scale = total / base
    return all(abs(g - GRID_DESIGN[i] * scale) <= max(60.0, 0.08 * GRID_DESIGN[i] * scale)
               for i, g in enumerate(grid))


def _mixed_valign(rows):
    """是否存在同一行内混用不同垂直对齐的行。"""
    for row in rows:
        cells = _row_cells(row)
        if len(cells) < 2:
            continue
        if len(set(_cell_valign(c) for c in cells)) > 1:
            return True
    return False


# ==================== 减分项 D01 ~ D31 ====================

def _score_ded(doc, xml, banner, t1, t2, rec_rows, grid2, results):
    _ded_t1(doc, t1, results)
    _ded_t2_header(doc, t2, grid2, results)
    _ded_records(doc, t2, rec_rows, results)
    _ded_page(doc, results)
    _ded_banner_note(doc, xml, banner, results)
    _ded_pic_misc(doc, xml, results)
    _ded_hf_extra(doc, t2, rec_rows, results)
def _ded_t1(doc, t1, results):
    """D01~D07：第 1 页那张五列明细表（指令未涉及，应原样保留）。"""
    rec1 = []
    if t1 is not None:
        for row in t1.rows[1:]:
            cs = _row_cells(row)
            if cs and any(k in _norm_ws(cs[0].text) for k in T1_DATE_KEYS):
                rec1.append(row)
    d01 = (t1 is None or len(t1.rows) < 3 or len(rec1) < 2)
    results.append(("-5" if d01 else "-0") +
                   "：第 1 页原有的 Date/Description/Withdrawals/Deposits/Balance 五列明细表"
                   "整表被删除，或其表头行、Mar 22 与 Mar 25 两条记录行缺失")

    hdr = _row_cells(t1.rows[0]) if not d01 else []

    # D02 表头文字与底纹
    d02 = False
    if not d01:
        for cell in hdr:
            if _cell_fill(cell) != T1_HDR_FILL:
                d02 = True
            for para, run in _text_runs(cell):
                size = _eff_size_pt(doc, para, run)
                if (_eff_font(doc, para, run, 'ascii') != 'Arial'
                        or size is None or abs(size - T1_HDR_SIZE) > 0.25
                        or not _eff_bold(doc, para, run)):
                    d02 = True
    results.append(("-1" if d02 else "-0") +
                   "：第 1 页明细表表头文字不再是 Arial 12.5 磅加粗，或表头底纹不再是 F2F7F7")

    # D03 表头上下框线
    d03 = False
    if not d01:
        for cell in hdr:
            if _cell_border(cell, 'top') != T1_HDR_TOP or \
                    _cell_border(cell, 'bottom') != T1_HDR_BOTTOM:
                d03 = True
    results.append(("-1" if d03 else "-0") +
                   "：第 1 页明细表表头行的上边框不再是 1.5 磅 7A7A7A 灰线（sz12），"
                   "或下边框不再是 1 磅 DDDDDD 浅灰线（sz8）")
    # D04 记录行下框线与行间新增横线
    #     与既有框线同色同粗（表头下 sz8 DDDDDD、记录行下 sz6 D8D8D8）的重复定义，
    #     叠加后在页面上看不出变化，按细则不计扣分。
    dup_ok = (None, T1_HDR_BOTTOM, T1_REC_BOTTOM, ('nil', None, None), ('none', None, None))
    d04 = False
    if not d01:
        for row in rec1:
            for cell in _row_cells(row):
                if _cell_border(cell, 'bottom') != T1_REC_BOTTOM:
                    d04 = True
                top = _cell_border(cell, 'top')
                if top not in dup_ok and _visible(top):
                    d04 = True
    results.append(("-1" if d04 else "-0") +
                   "：第 1 页明细表记录行的下边框不再是 0.75 磅 D8D8D8 细线（sz6），"
                   "或两条记录行之间被新增了与原样式不同的可见横线")

    # D05 同一行内混用垂直对齐
    d05 = (not d01) and _mixed_valign(t1.rows)
    results.append(("-1" if d05 else "-0") +
                   "：第 1 页明细表出现同一行内各单元格垂直对齐方式不一致的情况，造成文字上下错位")

    # D06 两条记录的金额
    d06 = False
    if not d01:
        got = []
        for row in rec1:
            for cell in _row_cells(row)[2:5]:
                txt = _norm_dash(_cell_text(cell)).replace(' ', '')
                if txt:
                    got.append(txt)
        d06 = (tuple(got) != T1_AMOUNTS)
    results.append(("-1" if d06 else "-0") +
                   "：第 1 页两条记录的金额或正负号被改动（应为 $8,732.40、$9,016.55、"
                   "-$6,918.75、$2,097.80）")

    # D07 列宽与整表宽度
    d07 = False
    if not d01:
        grid1 = _grid_cols(t1)
        d07 = (not _is_even_grid(grid1)) or (sum(grid1) > TEXT_WIDTH_TW + TW_TOL)
    results.append(("-1" if d07 else "-0") +
                   "：第 1 页明细表不再呈现五列等宽 2086 缇的版式，或整表宽度超出正文区 10430 缇")
def _ded_t2_header(doc, t2, grid2, results):
    """D08~D14：第 2 页明细表表头（交付前已排好，指令只要求填记录）。"""
    hdr = _row_cells(t2.rows[0])

    # D08 表头文字
    d08 = False
    for cell in hdr:
        runs = _text_runs(cell)
        if not runs:
            d08 = True
        for para, run in runs:
            size = _eff_size_pt(doc, para, run)
            if (_eff_font(doc, para, run, 'ascii') != 'Arial'
                    or size is None or abs(size - T2_HDR_SIZE) > 0.25
                    or not _eff_bold(doc, para, run)):
                d08 = True
    results.append(("-1" if d08 else "-0") +
                   "：第 2 页明细表表头文字不再满足 Arial 11.5 磅加粗")

    # D09 表头底纹
    d09 = any(not _is_light_fill(_cell_fill(c)) for c in hdr)
    results.append(("-1" if d09 else "-0") +
                   "：第 2 页明细表表头底纹被改成深色或高饱和色，不再是白色或与 F2F7F7 "
                   "同族的极浅色")

    # D10 表头上边框
    d10 = False
    for cell in hdr:
        bd = _cell_border(cell, 'top')
        if bd is None or bd[0] != 'single' or not _near(bd[1], T2_HDR_TOP_SZ) \
                or not _is_black(bd[2]):
            d10 = True
    results.append(("-1" if d10 else "-0") +
                   "：第 2 页明细表表头行的上边框不再是 1.5 磅黑色单实线（sz12）")

    # D11 表头下框线
    d11 = False
    for cell in hdr:
        bd = _cell_border(cell, 'bottom')
        if not _visible(bd):
            continue
        val, sz, clr = bd
        if val != 'single' or (sz or 0) > 8 or not _is_light_gray(clr):
            d11 = True
    results.append(("-1" if d11 else "-0") +
                   "：第 2 页明细表表头行的下框线被加成黑色粗线、双线或彩色线（原为无下框线，"
                   "补一条 1 磅以内的浅灰细线不计扣分）")
    # D12 表头单元格内边距
    d12 = any(not _mar_ok(_cell_mar(c), T2_HDR_MAR) for c in hdr)
    results.append(("-1" if d12 else "-0") +
                   "：第 2 页明细表表头单元格内边距不再是上下 90 缇、左右 80 缇，表头行高矮变形")

    # D13 列宽与整表宽度
    d13 = (not (_is_even_grid(grid2) or _is_design_grid(grid2))) \
        or (sum(grid2) > TEXT_WIDTH_TW + TW_TOL)
    results.append(("-1" if d13 else "-0") +
                   "：第 2 页明细表五列列宽既不是等宽 2086 缇、也不是按 1080/4896/1800/1584/1440 "
                   "的比例加宽 Description 列，而是被改成任意失衡的列宽，或整表宽度超出正文区 "
                   "10430 缇（实测各列 %s，合计 %d 缇）"
                   % ('/'.join(str(g) for g in grid2), sum(grid2)))

    # D14 表头同一行内混用垂直对齐
    d14 = _mixed_valign(t2.rows[:1])
    results.append(("-1" if d14 else "-0") +
                   "：第 2 页明细表表头同一行内各单元格垂直对齐方式不一致")


def _ded_records(doc, t2, rec_rows, results):
    """D15~D17：为塞进记录而对记录行本身做的过度改动。"""
    cells_of = [_row_cells(r) for r in rec_rows]

    # D15 多余的空文本片段或空段落
    d15 = False
    for cs in cells_of:
        for cell in cs[:5]:
            has_text = _cell_text(cell) != ''
            blanks = sum(1 for p in cell.paragraphs if _norm_ws(p.text) == '')
            if blanks > (0 if has_text else 1):
                d15 = True
            for para in cell.paragraphs:
                for run in para.runs:
                    if run.text == '' and run._r.find(qn('w:br')) is None:
                        d15 = True
    results.append(("-1" if d15 else "-0") +
                   "：记录单元格内被插入多余的空文本片段或空段落，虚增行高、给后续编辑留下"
                   "看不见的冗余内容")

    # D16 记录行同一行内混用垂直对齐
    d16 = _mixed_valign(rec_rows)
    results.append(("-1" if d16 else "-0") +
                   "：记录行出现同一行内各单元格垂直对齐方式不一致的情况")

    # D17 表头行行高与记录行行高不一致
    h_hdr = _row_height(t2.rows[0])[0]
    h_recs = [_row_height(r)[0] for r in rec_rows]
    h_set = set(v for v in h_recs if v is not None)
    d17 = False
    if h_hdr is not None and h_set:
        d17 = any(not _near(h_hdr, v) for v in h_set)
    results.append(("-1" if d17 else "-0") +
                   "：表头行的行高数值被设成与记录行不一致（表头 %s / 记录行 %s 缇），"
                   "表头与正文之间出现突兀的高度差"
                   % (h_hdr, '/'.join(str(v) for v in sorted(h_set)) if h_set else '未设置'))
def _ded_page(doc, results):
    """D18~D20：页面设置（指令未涉及）。"""
    geo = _sect_geometry(doc)

    d18 = not (_near(geo['w'], PG_SZ[0]) and _near(geo['h'], PG_SZ[1]))
    results.append(("-3" if d18 else "-0") +
                   "：页面尺寸不再是 A4 的 11906 × 16838 缇（实测 %s × %s 缇），"
                   "打印时纸张对不上" % (geo['w'], geo['h']))

    d19 = not all(_near(geo[k], PG_MAR[k]) for k in PG_MAR)
    results.append(("-3" if d19 else "-0") +
                   "：页边距不再是上 1814、下 1304、左 737、右 737 缇（实测上 %s / 下 %s / "
                   "左 %s / 右 %s 缇），正文区宽度随之改变、既有版式跟着走形"
                   % (geo['top'], geo['bottom'], geo['left'], geo['right']))

    d20 = (not _near(geo['hdr'], PG_HDR_DIST)) or (not _near(geo['ftr'], PG_FTR_DIST)) \
        or _has_visible_pg_borders(doc)
    results.append(("-1" if d20 else "-0") +
                   "：页眉距边界 0 缇、页脚距边界 624 缇被改动（实测 %s / %s 缇），"
                   "或页面属性里被新增了在页面上可见的页面边框" % (geo['hdr'], geo['ftr']))


def _ded_banner_note(doc, xml, banner, results):
    """D21~D24：顶部标题条与圆角矩形说明框（指令未涉及）。"""
    b_cell = None
    if banner is not None:
        for cell in _row_cells(banner.rows[0]):
            if _norm_ws(cell.text) != '':
                b_cell = cell
                break

    d21 = (b_cell is None or _norm_ws(b_cell.text) != BANNER_TEXT)
    results.append(("-3" if d21 else "-0") +
                   "：顶部「Activity details — 出支记录」标题条被删除，或标题文字被改写")

    d22 = True
    if b_cell is not None:
        d22 = False
        if _cell_fill(b_cell) != BANNER_FILL:
            d22 = True
        if not _mar_ok(_cell_mar(b_cell), BANNER_MAR):
            d22 = True
        for side in ('top', 'bottom', 'left', 'right'):
            if _visible(_cell_border(b_cell, side)):
                d22 = True
        for para, run in _text_runs(b_cell):
            size = _eff_size_pt(doc, para, run)
            if (_eff_font(doc, para, run, 'ascii') != BANNER_FONT or size is None
                    or abs(size - BANNER_SIZE) > 0.25 or not _eff_bold(doc, para, run)):
                d22 = True
    results.append(("-1" if d22 else "-0") +
                   "：标题条不再是 Arial 16 磅加粗 + E8F2F1 底纹 + 四周无框线 + 上 480、"
                   "下 420、左 420、右 360 缇内边距的样式")
    segs = _shape_segments(xml)
    note = _find_note_seg(segs)

    note_text = _norm_ws(_norm_dash(''.join(_seg_texts(note)))) if note is not None else ''
    d23 = (note is None or NOTE_TEXT.lower() not in note_text.lower())
    results.append(("-3" if d23 else "-0") +
                   "：右上角圆角矩形说明框被删除，或框内「Keep daily financial transactions "
                   "in order.」文字被改写")

    d24 = True
    if note is not None:
        d24 = False
        cx, cy = _seg_extent(note)
        if not (_near(cx, NOTE_EXTENT[0], EMU_TOL) and _near(cy, NOTE_EXTENT[1], EMU_TOL)):
            d24 = True
        if not re.search(r'<a:prstDash\s+val="(dash|sysDash|lgDash|dashDot|sysDashDot)"', note):
            d24 = True
        if NOTE_LINE_CLR not in note.upper():
            d24 = True
        fonts = re.findall(r'<w:rFonts[^>]*w:eastAsia="([^"]+)"', note)
        if fonts and not any(f in NOTE_EA_FONT for f in fonts):
            d24 = True
        szs = [_int(s) for s in re.findall(r'<w:sz\s+w:val="(\d+)"', note)]
        if szs and not any(_near(s, int(NOTE_SIZE * 2)) for s in szs):
            d24 = True
    results.append(("-1" if d24 else "-0") +
                   "：说明框的虚线边框、1A5D60 线条颜色、1079500 × 3397250 EMU 的尺寸或框内"
                   "宋体 12 磅的字体字号被改动")


def _ded_pic_misc(doc, xml, results):
    """D25~D28：第 2 页插图、小标题与分页/占位换行符。"""
    segs = _shape_segments(xml)
    pic = _find_pic_seg(segs)
    cx, cy = _seg_extent(pic) if pic is not None else (None, None)

    d25 = (pic is None)
    if not d25 and cx and cy:
        ratio = cx / float(cy)
        if abs(ratio - PIC_RATIO) / PIC_RATIO > 0.02:
            d25 = True
    results.append(("-3" if d25 else "-0") +
                   "：第 2 页那张插图被删除、被替换成别的图，或被拉伸压缩成与原比例不符的"
                   "变形失真状态")

    d26 = True
    if pic is not None:
        d26 = not (_near(cx, PIC_EXTENT[0], EMU_TOL) and _near(cy, PIC_EXTENT[1], EMU_TOL)
                   and 'Square' in _seg_wrap(pic)
                   and _near(_seg_pos(pic, 'V'), PIC_POS_V, EMU_TOL))
    results.append(("-1" if d26 else "-0") +
                   "：插图的 5848350 × 3665220 EMU 尺寸、四周型文字环绕或 -3336290 EMU 的"
                   "垂直偏移被改动，图文位置关系跟着变化")
    # D27 「Activity details」小标题
    d27 = True
    for para in doc.paragraphs:
        if _norm_ws(para.text) != SUB_TITLE:
            continue
        d27 = False
        runs = [(para, r) for r in para.runs if _norm_ws(r.text) != '']
        if not runs:
            d27 = True
        for p, run in runs:
            size = _eff_size_pt(doc, p, run)
            if (_eff_font(doc, p, run, 'ascii') != 'Arial' or size is None
                    or abs(size - SUB_TITLE_SIZE) > 0.25 or not _eff_bold(doc, p, run)):
                d27 = True
        break
    results.append(("-1" if d27 else "-0") +
                   "：正文中的「Activity details」小标题缺失，或不再是 Arial 13.5 磅加粗")

    # D28 分页符与明细表前的占位手动换行符
    brs = re.findall(r'<w:br\b[^>]*/?>', xml)
    page_br = [b for b in brs if 'w:type="page"' in b]
    text_br = [b for b in brs if 'w:type="page"' not in b and 'w:type="column"' not in b]
    has_page = bool(page_br) or ('<w:pageBreakBefore' in xml) or len(doc.sections) > 1
    d28 = (not has_page) or (not text_br)
    results.append(("-1" if d28 else "-0") +
                   "：分页符被删除导致原本分处两页的内容挤到一起，或明细表前用于占位的"
                   "手动换行符被删除")


def _ded_hf_extra(doc, t2, rec_rows, results):
    """D29~D31：页眉、页脚与凭空编造的记录。"""
    h_texts, _h_instr = _hf_texts(doc, 'header')
    f_texts, f_instr = _hf_texts(doc, 'footer')

    d29 = HEADER_TEXT.lower() not in ' '.join(h_texts).lower()
    results.append(("-3" if d29 else "-0") +
                   "：页眉里的「Income and Expense Record」文字缺失或被改写")

    joined = ' '.join(f_instr).upper()
    d30 = not (re.search(r'(?<!NUM)PAGE\b', joined) and 'NUMPAGES' in joined)
    results.append(("-3" if d30 else "-0") +
                   "：页脚「Page X of Y」的 PAGE / NUMPAGES 页码域缺失，或被改成写死的静态文字")

    extra = []
    for row in rec_rows:
        cs = _row_cells(row)
        if len(cs) < 2:
            continue
        desc = _norm_desc(cs[1].text)
        if desc == '' and _norm_ws(cs[0].text) == '':
            continue
        if desc not in set(_norm_desc(d) for d in PDF_DESCS):
            extra.append(_norm_ws(cs[1].text) or _norm_ws(cs[0].text))
    d31 = bool(extra)
    results.append(("-3" if d31 else "-0") +
                   "：明细表里新增了 PDF 对账单中不存在的记录行、小计行或说明行（实测 %d 条：%s）"
                   % (len(extra), '、'.join(extra[:3]) if extra else '无'))


if __name__ == "__main__":
    _dir = sys.argv[1] if len(sys.argv) >= 2 else os.path.dirname(os.path.abspath(__file__))
    print(json.dumps(evaluate(_dir), ensure_ascii=False, indent=2))

