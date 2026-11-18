#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# pyright: reportUnknownMemberType=false, reportUnknownVariableType=false, reportUnknownArgumentType=false, reportUnknownParameterType=false, reportMissingParameterType=false, reportPrivateUsage=false, reportUnknownLambdaType=false, reportGeneralTypeIssues=false

import os
import re
import sys
import json
import math
import zipfile
import xml.etree.ElementTree as ET
from openpyxl import load_workbook

# 说明：按接口统一约定，本模块不修改全局 sys.stdout，也不 print 主结果。
# 主结果统一由 evaluate(dir_path) -> dict 返回；入参为“脚本所在目录的路径”，
# 由脚本自己在该目录内定位并打开被评估的 .xlsx。本文件仅在 __main__ 下作为
# 本地调试入口，才把 dict 序列化为 JSON 打印。
#
# 与 008 号脚本的差异：
#   ① 本任务交付物是 3 个工作簿（三张 S₁₁ 数据表各一张图），不是单文档。脚本在
#      目录内取全部 .xlsx（忽略 ~$ 临时锁文件），按“数据指纹”（数据行数 22/25/28
#      与频率-S₁₁ 数值）把每个文件归位到表1/表2/表3，不依赖文件名。
#   ② 细则同时含“加分项”（指令要求的绘图任务）与“减分项”（指令未涉及却改坏数据表
#      或在图上凭空添加的多余内容）。故 dim2 项的 max_delta 允许为负：正值=加分项
#      （命中记满分，未命中记 0）；负值=减分项（触发“破坏”记该负分，未触发记 0）。
#      total_score 为全部 delta 之和，取值区间 [-22, +48]；max_score 仍为加分项
#      满分之和（+48），作为完成度上限。
#   ③ 加分项按“三个文件全部满足才得分”判定；减分项按“任一文件触发即扣分”判定，
#      同一问题在多个文件重复出现只计一次（与细则计分口径一致）。
#   ④ 图表相关判定不经 openpyxl 的图表对象模型（会丢 userShapes/主题色等信息），
#      而是直接解 xl/charts/*.xml、xl/drawings/*.xml 原始 OOXML。

SCRIPT_ID = "085"

# —— 加分项按“结果追加顺序”的 max_delta（正值）——
_ADD_DELTAS_ORDER = [
    3,  # A01 交付3个文件分别承载表1/表2/表3数据，一表一图
    3,  # A02 图与数据同处一个工作簿的同一工作表
    5,  # A03 图为原生可编辑图表对象（非位图）
    5,  # A04 数据系列以单元格区域引用绑定
    3,  # A05 图表类型为带直线和数据标记的散点图
    1,  # A06 数据点全部绘入（22/25/28）
    3,  # A07 数据点带圆形标记（约5磅）
    3,  # A08 蓝色实线、线宽1.5~2磅、不平滑
    3,  # A09 横轴 0~24、主刻度单位 4
    3,  # A10 纵轴 -24~-8、主刻度单位 4
    3,  # A11 横轴标题“频率（GHz）”
    3,  # A12 纵轴标题“S11”
    5,  # A13 -10 dB 灰色虚线水平参考线
    1,  # A14 不显示图例
    1,  # A15 不显示数据标签
    1,  # A16 图表区白色/浅色底、无深色填充与粗边框
    1,  # A17 图表尺寸 宽14~22cm、高8~13cm
    1,  # A18 图表锚定在空白区，未覆盖 A/B 两列数据
]
# —— 减分项按“结果追加顺序”的 max_delta（负值）——
_DED_DELTAS_ORDER = [
    -3,  # D01 数值被修改/四舍五入/行序被调整
    -3,  # D02 数据行被删减或表头文字被改
    -1,  # D03 数据区旁新增额外数据列
    -1,  # D04 单元格中新增说明性文字
    -1,  # D05 新增合并单元格区域
    -1,  # D06 新增无关工作表或工作表名不为 Sheet1
    -1,  # D07 A/B 两列列宽不为 16 字符
    -1,  # D08 第1行行高不为26磅或数据行行高不为21磅
    -1,  # D09 表头字体/字色/居中被改
    -1,  # D10 表头底纹色被改或被统一成同一颜色
    -1,  # D11 数据单元格字体/字色/居中被改
    -1,  # D12 数据行隔行底纹被改
    -1,  # D13 数据单元格数字格式不为一位小数
    -1,  # D14 表格框线粗细被改
    -1,  # D15 新增内容字体与表内不一致
    -1,  # D16 图表中加入多余数据系列
    -1,  # D17 图表中加入水印/说明文本框/深色底纹/粗边框等装饰
    -1,  # D18 文件名无法与表1/表2/表3一一对应
]

MAX_SCORE_TOTAL = sum(_ADD_DELTAS_ORDER)               # +48：完成度上限
_MIN_TOTAL = MAX_SCORE_TOTAL + sum(_DED_DELTAS_ORDER)  # +26：全部命中且全部踩坑
assert MAX_SCORE_TOTAL == 48
assert sum(_DED_DELTAS_ORDER) == -22

# 结果追加顺序 = 先加分项、后减分项
_MAX_DELTAS_ORDER = _ADD_DELTAS_ORDER + _DED_DELTAS_ORDER
_ADD_COUNT = len(_ADD_DELTAS_ORDER)

# OOXML 命名空间
NS = {
    'x':   'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
    'a':   'http://schemas.openxmlformats.org/drawingml/2006/main',
    'c':   'http://schemas.openxmlformats.org/drawingml/2006/chart',
    'xdr': 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing',
    'cdr': 'http://schemas.openxmlformats.org/drawingml/2006/chartDrawing',
    'r':   'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    'rel': 'http://schemas.openxmlformats.org/package/2006/relationships',
    'ct':  'http://schemas.openxmlformats.org/package/2006/content-types',
}
EMU_PER_CM = 360000.0
EMU_PER_PT = 12700.0
# —— 三张数据表的基准指纹与格式基准（用于归位与减分项判定）——
# 频率列/S₁₁ 列数值、表头底纹色、数据行隔行底纹色；数据行数 22/25/28 互不相同，
# 可单独作为归位主键，数值再做二次校验。
_EXPECTED = {
    1: {
        'freq': [0, 1.5, 3, 4, 5.5, 6.5, 7.5, 8.5, 9.5, 10.5, 11.5, 12.5, 13.5,
                 14.5, 15.5, 16.5, 17.5, 18.5, 19.5, 21, 22.5, 24],
        's11': [-10.5, -15, -22, -21.5, -16.5, -14.5, -14, -15.5, -17.5, -18.5,
                -19.5, -18.8, -18, -19, -22.2, -18, -13, -10.5, -11.5, -10.8, -9, -8.5],
        'header_fill': '1F4E78',
        'band_fill': 'EAF3F8',
    },
    2: {
        'freq': [0, 1.2, 2.4, 3.6, 4.8, 6, 7.2, 8.4, 9.6, 10.8, 12, 13.2, 14.4,
                 15.6, 16.8, 18, 19, 20, 21, 21.8, 22.5, 23, 23.4, 23.7, 24],
        's11': [-10.8, -15.2, -22.3, -21.7, -16.8, -14.8, -14.2, -14.1, -15.8, -17.8,
                -18.8, -19.8, -19, -18.2, -19.2, -22.5, -18.3, -13.3, -10.7, -11.7,
                -11, -10, -9.2, -8.8, -8.3],
        'header_fill': '5B3A70',
        'band_fill': 'F1EAF6',
    },
    3: {
        'freq': [0, 0.9, 1.8, 2.7, 3.6, 4.5, 5.4, 6.3, 7.2, 8.1, 9, 9.9, 10.8, 11.7,
                 12.6, 13.5, 14.4, 15.3, 16.2, 17.1, 18, 19, 20, 21, 22, 22.8, 23.5, 24],
        's11': [-10.6, -14.9, -22.1, -21.6, -16.6, -14.7, -14.3, -14.1, -14, -15.7,
                -17.7, -18.7, -19.7, -19.1, -18.3, -18.1, -19.1, -22.4, -18.2, -13.2,
                -10.6, -11.6, -11.1, -10.1, -9.3, -8.9, -8.6, -8.2],
        'header_fill': '2F6B4F',
        'band_fill': 'EAF4EE',
    },
}
_HEADER_TEXTS = ('频率 (GHz)', 'S₁₁(dB)')
_TABLE_FONT = '微软雅黑'
_HEADER_SIZE = 11.0
_DATA_SIZE = 10.0
_DATA_FONT_COLOR = '1F2937'
_COL_WIDTH = 16.0
_ROW1_HEIGHT = 26.0
_DATA_ROW_HEIGHT = 21.0
_DATA_NUMFMT = '0.0'
# 表头行/数据区的框线基准：(左, 右, 上, 下)
_BD_HDR_A = ('medium', 'thin', 'medium', 'medium')
_BD_HDR_B = ('thin', 'medium', 'medium', 'medium')
_BD_DATA = ('thin', 'thin', 'thin', 'thin')
# 图面判定阈值
_REF_LINE_Y = -10.0          # 参考线所在 S₁₁ 值
_AXIS_TOL = 0.001            # 坐标轴刻度容差
_LINE_W_MIN, _LINE_W_MAX = 19050, 25400   # 1.5~2 磅（EMU）
_MARKER_SIZE_MIN, _MARKER_SIZE_MAX = 4, 7  # 圆形标记直径（磅）
_CHART_W_MIN, _CHART_W_MAX = 14.0, 22.0    # 图表宽（cm）
_CHART_H_MIN, _CHART_H_MAX = 8.0, 13.0     # 图表高（cm）
_DASH_STYLES = ('dash', 'sysDash', 'lgDash', 'sysDot', 'dot', 'dashDot', 'lgDashDot')
def _q(ns, tag):
    """返回带命名空间的标签名，便于 ElementTree 检索。"""
    return '{%s}%s' % (NS[ns], tag)


def _zread(z, name):
    """读取包内部件文本；不存在返回 None。"""
    try:
        return z.read(name).decode('utf-8', 'replace')
    except KeyError:
        return None


def _zxml(z, name):
    """读取包内部件并解析为 XML 根节点；不存在或解析失败返回 None。"""
    txt = _zread(z, name)
    if txt is None:
        return None
    try:
        return ET.fromstring(txt.encode('utf-8'))
    except ET.ParseError:
        return None


def _close(a, b, tol=1e-6):
    """浮点近似相等（None 不参与比较）。"""
    if a is None or b is None:
        return False
    return abs(float(a) - float(b)) <= tol


def _hexrgb(v):
    """把 ARGB/RRGGBB 统一成 6 位大写 RRGGBB；无效返回 None。"""
    if not v or not isinstance(v, str):
        return None
    s = v.strip().upper()
    if len(s) == 8:
        s = s[2:]
    return s if re.fullmatch(r'[0-9A-F]{6}', s) else None


def _rgb_tuple(hexv):
    h = _hexrgb(hexv)
    if not h:
        return None
    return (int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))


def _is_gray(hexv, tol=28):
    """灰（含浅灰/深灰）：三通道极差不超过 tol。"""
    t = _rgb_tuple(hexv)
    return bool(t) and (max(t) - min(t)) <= tol


def _is_blue(hexv):
    """蓝：蓝通道明显高于红/绿通道，且蓝通道足够深（排除浅灰蓝）。"""
    t = _rgb_tuple(hexv)
    if not t:
        return False
    r, g, b = t
    return b >= 0x80 and b - r >= 0x40 and b - g >= 0x40


def _is_light(hexv):
    """浅色（白/近白/浅灰）：相对亮度不低于 0.80。"""
    t = _rgb_tuple(hexv)
    if not t:
        return False
    r, g, b = t
    return (0.299 * r + 0.587 * g + 0.114 * b) / 255.0 >= 0.80
def _theme_scheme(z):
    """读取 xl/theme/theme1.xml 的配色方案，返回 {'dk1':'000000', 'accent1':'4874CB', ...}。
       同时补上 DrawingML 里 tx1/bg1/tx2/bg2 到 dk1/lt1/dk2/lt2 的别名。"""
    scheme = {}
    root = _zxml(z, 'xl/theme/theme1.xml')
    if root is None:
        return scheme
    for el in root.iter(_q('a', 'clrScheme')):
        for child in el:
            name = child.tag.split('}')[-1]
            srgb = child.find(_q('a', 'srgbClr'))
            sysc = child.find(_q('a', 'sysClr'))
            val = None
            if srgb is not None:
                val = _hexrgb(srgb.get('val'))
            elif sysc is not None:
                val = _hexrgb(sysc.get('lastClr'))
            if val:
                scheme[name] = val
        break
    for alias, real in (('tx1', 'dk1'), ('bg1', 'lt1'), ('tx2', 'dk2'), ('bg2', 'lt2')):
        if real in scheme:
            scheme[alias] = scheme[real]
    return scheme


def _apply_lum(hexv, lum_mod, lum_off):
    """把 a:lumMod / a:lumOff 施加到基色上（按 HSL 的 L 通道换算），用于把
       schemeClr+lumMod 这类“主题色的深浅变体”还原成看到的颜色。"""
    t = _rgb_tuple(hexv)
    if not t:
        return hexv
    r, g, b = (v / 255.0 for v in t)
    mx, mn = max(r, g, b), min(r, g, b)
    lum = (mx + mn) / 2.0
    if mx == mn:
        h = s = 0.0
    else:
        d = mx - mn
        s = d / (2.0 - mx - mn) if lum > 0.5 else d / (mx + mn)
        if mx == r:
            h = ((g - b) / d) % 6.0
        elif mx == g:
            h = (b - r) / d + 2.0
        else:
            h = (r - g) / d + 4.0
        h /= 6.0
    if lum_mod is not None:
        lum *= lum_mod
    if lum_off is not None:
        lum += lum_off
    lum = min(1.0, max(0.0, lum))
    if s == 0:
        rr = gg = bb = lum
    else:
        q = lum * (1 + s) if lum < 0.5 else lum + s - lum * s
        p = 2 * lum - q

        def _hue(tc):
            tc = tc % 1.0
            if tc < 1 / 6.0:
                return p + (q - p) * 6 * tc
            if tc < 0.5:
                return q
            if tc < 2 / 3.0:
                return p + (q - p) * (2 / 3.0 - tc) * 6
            return p
        rr, gg, bb = _hue(h + 1 / 3.0), _hue(h), _hue(h - 1 / 3.0)
    return '%02X%02X%02X' % (round(rr * 255), round(gg * 255), round(bb * 255))
def _fill_color(el, scheme):
    """解析 <a:solidFill> 容器（可为 a:ln / c:spPr 等）内的填充色 → RRGGBB。
       支持 a:srgbClr 与 a:schemeClr（含 lumMod/lumOff 变体）。无实色返回 None。"""
    if el is None:
        return None
    fill = el.find(_q('a', 'solidFill'))
    if fill is None:
        return None
    srgb = fill.find(_q('a', 'srgbClr'))
    if srgb is not None:
        return _hexrgb(srgb.get('val'))
    sch = fill.find(_q('a', 'schemeClr'))
    if sch is not None:
        base = scheme.get(sch.get('val') or '')
        if not base:
            return None
        lm = sch.find(_q('a', 'lumMod'))
        lo = sch.find(_q('a', 'lumOff'))
        lm_v = int(lm.get('val')) / 100000.0 if (lm is not None and lm.get('val')) else None
        lo_v = int(lo.get('val')) / 100000.0 if (lo is not None and lo.get('val')) else None
        if lm_v is None and lo_v is None:
            return base
        return _apply_lum(base, lm_v, lo_v)
    return None


def _has_nofill(el):
    return el is not None and el.find(_q('a', 'noFill')) is not None


def _all_text(el):
    """收集元素子树内所有 <a:t> 文本，拼成一个字符串。"""
    if el is None:
        return ''
    return ''.join(t.text or '' for t in el.iter(_q('a', 't')))


def _norm_text(s):
    """归一化文本：去空白、全角括号转半角、下标数字转常规数字，便于比对轴标题。"""
    if not s:
        return ''
    s = re.sub(r'\s+', '', s)
    table = {'（': '(', '）': ')', '₀': '0', '₁': '1', '₂': '2', '₃': '3', '₄': '4',
             '₅': '5', '₆': '6', '₇': '7', '₈': '8', '₉': '9'}
    return ''.join(table.get(ch, ch) for ch in s)


def _rels(z, part_name):
    """读取某部件的 .rels，返回 {rId: 绝对部件名}。"""
    d = os.path.dirname(part_name).replace('\\', '/')
    rels_name = (d + '/_rels/' + os.path.basename(part_name) + '.rels').lstrip('/')
    root = _zxml(z, rels_name)
    out = {}
    if root is None:
        return out
    for rel in root.findall(_q('rel', 'Relationship')):
        rid = rel.get('Id')
        tgt = rel.get('Target') or ''
        if rel.get('TargetMode') == 'External' or not rid:
            continue
        if tgt.startswith('/'):
            abs_name = tgt.lstrip('/')
        else:
            abs_name = os.path.normpath(os.path.join(d, tgt)).replace('\\', '/')
        out[rid] = abs_name
    return out
def _num_vals(holder):
    """解析 c:xVal / c:yVal：返回 (是否单元格引用, 引用式, 点数, 数值列表)。
       numRef=绑定单元格区域（缓存值可读）；numLit=数值被硬编码进图表内部。"""
    if holder is None:
        return (False, None, 0, [])
    ref = holder.find(_q('c', 'numRef'))
    lit = holder.find(_q('c', 'numLit'))
    src = ref if ref is not None else lit
    if src is None:
        return (False, None, 0, [])
    f_el = src.find(_q('c', 'f'))
    formula = (f_el.text or '').strip() if f_el is not None else None
    cache = src.find(_q('c', 'numCache')) if ref is not None else src
    vals = []
    n = 0
    if cache is not None:
        pc = cache.find(_q('c', 'ptCount'))
        if pc is not None and pc.get('val'):
            try:
                n = int(pc.get('val'))
            except ValueError:
                n = 0
        for pt in cache.findall(_q('c', 'pt')):
            v = pt.find(_q('c', 'v'))
            try:
                vals.append(float((v.text or '').strip()))
            except (AttributeError, ValueError):
                pass
    if not n:
        n = len(vals)
    return (ref is not None, formula, n, vals)


def _parse_axis(ax):
    """解析一个 c:valAx / c:catAx：位置、显示范围、主刻度单位、轴标题文字。"""
    pos_el = ax.find(_q('c', 'axPos'))
    scaling = ax.find(_q('c', 'scaling'))
    mn = mx = None
    if scaling is not None:
        for key, tag in (('min', 'min'), ('max', 'max')):
            el = scaling.find(_q('c', tag))
            if el is not None and el.get('val') is not None:
                try:
                    v = float(el.get('val'))
                except ValueError:
                    v = None
                if key == 'min':
                    mn = v
                else:
                    mx = v
    mu_el = ax.find(_q('c', 'majorUnit'))
    try:
        mu = float(mu_el.get('val')) if (mu_el is not None and mu_el.get('val')) else None
    except ValueError:
        mu = None
    del_el = ax.find(_q('c', 'delete'))
    return {
        'pos': (pos_el.get('val') if pos_el is not None else None),
        'min': mn,
        'max': mx,
        'major_unit': mu,
        'title': _all_text(ax.find(_q('c', 'title'))),
        'deleted': (del_el is not None and del_el.get('val') in ('1', 'true')),
        'is_cat': ax.tag.endswith('catAx'),
    }
def _tx_text(el):
    """取 c:tx 一类容器内的文本（优先 strRef 缓存值，其次富文本）。"""
    if el is None:
        return ''
    parts = [t.text or '' for t in el.iter(_q('c', 'v'))]
    if any(p.strip() for p in parts):
        return ''.join(parts)
    return _all_text(el)


def _line_props(sp_pr, scheme):
    """解析 spPr 中的 a:ln：线宽(EMU)、线色、虚线样式、是否显式无线。"""
    out = {'w': None, 'color': None, 'dash': None, 'nofill': False, 'has_ln': False}
    if sp_pr is None:
        return out
    ln = sp_pr.find(_q('a', 'ln'))
    if ln is None:
        return out
    out['has_ln'] = True
    if ln.get('w'):
        try:
            out['w'] = int(ln.get('w'))
        except ValueError:
            pass
    out['color'] = _fill_color(ln, scheme)
    out['nofill'] = _has_nofill(ln)
    dash = ln.find(_q('a', 'prstDash'))
    if dash is not None:
        out['dash'] = dash.get('val')
    return out


def _show_val(dlbls):
    """判断一个 c:dLbls 是否真的显示数值标签（delete=1 视为不显示）。"""
    if dlbls is None:
        return False
    d = dlbls.find(_q('c', 'delete'))
    if d is not None and (d.get('val') is None or d.get('val') in ('1', 'true')):
        return False
    sv = dlbls.find(_q('c', 'showVal'))
    return sv is not None and sv.get('val') in ('1', 'true')
def _parse_user_shapes(z, chart_part, scheme):
    """解析图表内嵌 userShapes（c:userShapes → chartDrawing 部件）。
       返回 (线条列表, 文本列表)：线条含是否水平、纵向位置、颜色、虚线样式，
       用于识别“-10 dB 灰色水平虚线”这类画在图上的参考线。"""
    lines, texts = [], []
    for name in _rels(z, chart_part).values():
        root = _zxml(z, name)
        if root is None or not root.tag.endswith('userShapes'):
            continue
        for anchor in list(root):
            ys = []
            for tag in ('from', 'to'):
                node = anchor.find(_q('cdr', tag))
                if node is None:
                    continue
                y = node.find(_q('cdr', 'y'))
                try:
                    ys.append(float((y.text or '').strip()))
                except (AttributeError, ValueError):
                    pass
            for sp in list(anchor):
                if sp.tag.split('}')[-1] in ('from', 'to', 'ext'):
                    continue
                txt = _all_text(sp)
                if txt.strip():
                    texts.append(txt)
                sp_pr = sp.find(_q('cdr', 'spPr'))
                lp = _line_props(sp_pr, scheme)
                if not lp['has_ln'] or lp['nofill']:
                    continue
                lines.append({
                    'horiz': (len(ys) == 2 and abs(ys[0] - ys[1]) <= 0.02),
                    'y': (ys[0] if ys else None),
                    'color': lp['color'],
                    'dash': lp['dash'],
                    'w': lp['w'],
                })
    return lines, texts
def _str_ref(holder):
    """c:cat 之类若为文本引用（strRef），返回 (引用式, 点数)；否则 (None, 0)。"""
    if holder is None:
        return (None, 0)
    sr = holder.find(_q('c', 'strRef'))
    if sr is None:
        return (None, 0)
    f_el = sr.find(_q('c', 'f'))
    cache = sr.find(_q('c', 'strCache'))
    n = 0
    if cache is not None:
        pc = cache.find(_q('c', 'ptCount'))
        if pc is not None and pc.get('val'):
            try:
                n = int(pc.get('val'))
            except ValueError:
                n = 0
    return ((f_el.text or '').strip() if f_el is not None else None, n)


def _parse_series(ser, kind, scheme):
    """解析一个 c:ser：名称、X/Y 数据来源与点数、标记、线型、平滑、数据标签。"""
    s = {'kind': kind, 'name': _tx_text(ser.find(_q('c', 'tx'))),
         'marker': None, 'marker_size': None, 'marker_color': None,
         'line_w': None, 'line_color': None, 'line_dash': None,
         'line_nofill': False, 'smooth': None, 'show_val': False}
    xv = ser.find(_q('c', 'xVal'))
    s['x_tag'] = 'xVal' if xv is not None else None
    if xv is None:
        xv = ser.find(_q('c', 'cat'))
        s['x_tag'] = 'cat' if xv is not None else None
    yv = ser.find(_q('c', 'yVal'))
    if yv is None:
        yv = ser.find(_q('c', 'val'))
    s['x_ref'], s['x_formula'], s['x_n'], s['x_vals'] = _num_vals(xv)
    s['y_ref'], s['y_formula'], s['y_n'], s['y_vals'] = _num_vals(yv)
    if not s['x_ref']:
        sf, sn = _str_ref(xv)
        if sf is not None:
            s['x_ref'], s['x_formula'] = True, sf
            s['x_n'] = s['x_n'] or sn
    mk = ser.find(_q('c', 'marker'))
    if mk is not None:
        sym = mk.find(_q('c', 'symbol'))
        s['marker'] = sym.get('val') if sym is not None else None
        sz = mk.find(_q('c', 'size'))
        if sz is not None and sz.get('val'):
            try:
                s['marker_size'] = float(sz.get('val'))
            except ValueError:
                pass
        s['marker_color'] = _fill_color(mk.find(_q('c', 'spPr')), scheme)
    lp = _line_props(ser.find(_q('c', 'spPr')), scheme)
    s['line_w'], s['line_color'] = lp['w'], lp['color']
    s['line_dash'], s['line_nofill'] = lp['dash'], lp['nofill']
    sm = ser.find(_q('c', 'smooth'))
    if sm is not None:
        s['smooth'] = sm.get('val') not in ('0', 'false')
    s['show_val'] = _show_val(ser.find(_q('c', 'dLbls')))
    return s
def _parse_chart(z, part, scheme):
    """直接解析 xl/charts/*.xml（不经 openpyxl 图表模型），抽取判定所需图面事实。"""
    f = {'part': part, 'kinds': [], 'scatter_style': None, 'has_scatter_style': False,
         'series': [], 'axes': [], 'has_legend': False, 'title': '',
         'space_fill': None, 'space_nofill': False, 'space_line_w': None,
         'space_line_color': None, 'plot_fill': None, 'shape_lines': [],
         'shape_texts': [], 'show_val': False}
    root = _zxml(z, part)
    if root is None:
        return f
    chart = root.find(_q('c', 'chart'))
    if chart is None:
        return f
    f['has_legend'] = chart.find(_q('c', 'legend')) is not None
    f['title'] = _all_text(chart.find(_q('c', 'title')))
    sp = root.find(_q('c', 'spPr'))
    f['space_fill'] = _fill_color(sp, scheme)
    f['space_nofill'] = _has_nofill(sp)
    lp = _line_props(sp, scheme)
    f['space_line_w'], f['space_line_color'] = lp['w'], lp['color']
    plot = chart.find(_q('c', 'plotArea'))
    if plot is None:
        return f
    f['plot_fill'] = _fill_color(plot.find(_q('c', 'spPr')), scheme)
    for grp in list(plot):
        tag = grp.tag.split('}')[-1]
        if tag in ('valAx', 'catAx', 'dateAx'):
            f['axes'].append(_parse_axis(grp))
            continue
        if not tag.endswith('Chart'):
            continue
        f['kinds'].append(tag)
        if tag == 'scatterChart':
            st = grp.find(_q('c', 'scatterStyle'))
            if st is not None:
                f['has_scatter_style'] = True
                f['scatter_style'] = st.get('val')
        if _show_val(grp.find(_q('c', 'dLbls'))):
            f['show_val'] = True
        for ser in grp.findall(_q('c', 'ser')):
            s = _parse_series(ser, tag, scheme)
            if s['show_val']:
                f['show_val'] = True
            f['series'].append(s)
    f['shape_lines'], f['shape_texts'] = _parse_user_shapes(z, part, scheme)
    return f
def _sheet_parts(z):
    """返回 [(工作表名, 工作表部件名), ...]，顺序与工作簿内一致。"""
    out = []
    root = _zxml(z, 'xl/workbook.xml')
    if root is None:
        return out
    rels = _rels(z, 'xl/workbook.xml')
    for sh in root.iter(_q('x', 'sheet')):
        rid = sh.get(_q('r', 'id'))
        part = rels.get(rid or '')
        if part:
            out.append((sh.get('name') or '', part))
    return out


def _parse_sheet_drawing(z, sheet_part, scheme):
    """解析工作表 drawing：图表锚点与显示尺寸、贴入的位图、以及独立文本形状。"""
    out = {'charts': [], 'pics': [], 'texts': []}
    for name in _rels(z, sheet_part).values():
        if '/drawings/' not in name or not name.endswith('.xml'):
            continue
        root = _zxml(z, name)
        if root is None or not root.tag.endswith('wsDr'):
            continue
        drels = _rels(z, name)
        for anchor in list(root):
            frm = anchor.find(_q('xdr', 'from'))
            to = anchor.find(_q('xdr', 'to'))
            pos = {}
            for key, node in (('from', frm), ('to', to)):
                if node is None:
                    continue
                c_el, r_el = node.find(_q('xdr', 'col')), node.find(_q('xdr', 'row'))
                try:
                    pos[key] = (int(c_el.text), int(r_el.text))
                except (AttributeError, TypeError, ValueError):
                    pass
            cx = cy = None
            ext = anchor.find(_q('xdr', 'ext'))
            if ext is not None:
                try:
                    cx, cy = int(ext.get('cx') or 0), int(ext.get('cy') or 0)
                except ValueError:
                    cx = cy = None
            for sp in list(anchor):
                tag = sp.tag.split('}')[-1]
                if tag == 'graphicFrame':
                    _collect_frame(sp, drels, pos, cx, cy, out)
                elif tag == 'pic':
                    out['pics'].append(pos.get('from'))
                elif tag in ('sp', 'grpSp'):
                    txt = _all_text(sp)
                    if txt.strip():
                        out['texts'].append(txt)
    return out
def _collect_frame(frame, drels, pos, cx, cy, out):
    """把一个 xdr:graphicFrame 里引用的图表部件与其锚点/尺寸登记到结果中。
       oneCellAnchor 用 xdr:ext；twoCellAnchor 退回 graphicFrame 自身的 a:ext。"""
    gd = frame.find(_q('a', 'graphic'))
    gd = gd.find(_q('a', 'graphicData')) if gd is not None else None
    ref = gd.find(_q('c', 'chart')) if gd is not None else None
    if ref is None:
        return
    part = drels.get(ref.get(_q('r', 'id')) or '')
    if cx is None or cy is None:
        xfrm = frame.find(_q('xdr', 'xfrm'))
        e = xfrm.find(_q('a', 'ext')) if xfrm is not None else None
        if e is not None:
            try:
                cx, cy = int(e.get('cx') or 0), int(e.get('cy') or 0)
            except ValueError:
                cx = cy = None
    out['charts'].append({'part': part, 'from': pos.get('from'),
                          'to': pos.get('to'), 'cx': cx, 'cy': cy})


def _media_names(z):
    """包内所有位图媒体文件名，用于识别“图以图片形式贴入”。"""
    return [n for n in z.namelist() if n.startswith('xl/media/')]


def _chart_parts(z):
    """依据 [Content_Types].xml 的 Override 列出全部图表部件（不限目录名）。"""
    parts = []
    root = _zxml(z, '[Content_Types].xml')
    if root is None:
        return parts
    want = 'application/vnd.openxmlformats-officedocument.drawingml.chart+xml'
    for ov in root.findall(_q('ct', 'Override')):
        if ov.get('ContentType') == want:
            parts.append((ov.get('PartName') or '').lstrip('/'))
    return parts
def _cell_fill(cell):
    """单元格实色底纹 → RRGGBB；无填充按白色，主题色返回 'THEME' 以便识别改动。"""
    f = cell.fill
    if f is None or f.patternType != 'solid':
        return 'FFFFFF'
    fg = f.fgColor
    if fg is not None and getattr(fg, 'type', None) == 'theme':
        return 'THEME'
    return _hexrgb(getattr(fg, 'rgb', None)) or 'FFFFFF'


def _font_color(cell):
    """单元格字色 → RRGGBB；主题色/索引色返回 'THEME'。"""
    c = cell.font.color
    if c is None:
        return None
    if getattr(c, 'type', None) in ('theme', 'indexed'):
        return 'THEME'
    return _hexrgb(getattr(c, 'rgb', None))


def _borders(cell):
    """返回 (左, 右, 上, 下) 四边框线样式名，未设置为 None。"""
    b = cell.border
    return tuple(getattr(getattr(b, side, None), 'style', None)
                 for side in ('left', 'right', 'top', 'bottom'))


def _cell_style(cell):
    """汇总单元格判定所需的字体/对齐/底纹/数字格式/框线。"""
    al = cell.alignment
    return {
        'font': cell.font.name,
        'size': float(cell.font.size) if cell.font.size is not None else None,
        'bold': bool(cell.font.bold),
        'color': _font_color(cell),
        'halign': al.horizontal if al is not None else None,
        'valign': al.vertical if al is not None else None,
        'fill': _cell_fill(cell),
        'numfmt': cell.number_format,
        'border': _borders(cell),
    }
def _probe_cells(path):
    """用 openpyxl 读取单元格层面的事实：数据、表头、行列尺寸、样式、多余内容。"""
    info = {'ok': False, 'err': None, 'sheets': [], 'ws_name': None, 'n': 0,
            'headers': ['', ''], 'freq': [], 's11': [], 'merged': [],
            'extra': [], 'col_w': {}, 'row1_h': None, 'data_h': [],
            'hdr': [], 'data': [], 'bands': []}
    try:
        wb = load_workbook(path)
    except Exception as exc:                      # noqa: BLE001 - 打不开即判不可用
        info['err'] = '%s: %s' % (type(exc).__name__, exc)
        return info
    try:
        info['sheets'] = list(wb.sheetnames)
        ws = wb['Sheet1'] if 'Sheet1' in wb.sheetnames else wb.worksheets[0]
        info['ws_name'] = ws.title
        info['headers'] = [str(ws['A1'].value or ''), str(ws['B1'].value or '')]
        r = 2
        while ws.cell(r, 1).value is not None or ws.cell(r, 2).value is not None:
            a, b = ws.cell(r, 1).value, ws.cell(r, 2).value
            try:
                info['freq'].append(float(a))
            except (TypeError, ValueError):
                info['freq'].append(None)
            try:
                info['s11'].append(float(b))
            except (TypeError, ValueError):
                info['s11'].append(None)
            info['data'].append(_cell_style(ws.cell(r, 1)))
            info['data'].append(_cell_style(ws.cell(r, 2)))
            info['bands'].append((r, _cell_fill(ws.cell(r, 1)), _cell_fill(ws.cell(r, 2))))
            dim = ws.row_dimensions.get(r)
            info['data_h'].append(getattr(dim, 'height', None) if dim is not None else None)
            r += 1
        info['n'] = r - 2
        info['hdr'] = [_cell_style(ws['A1']), _cell_style(ws['B1'])]
        d1 = ws.row_dimensions.get(1)
        info['row1_h'] = getattr(d1, 'height', None) if d1 is not None else None
        # <col min="1" max="2" width="16"/> 这类区间写法下 openpyxl 只按 min 建键，
        # 故按 min~max 覆盖范围回填 A/B 两列列宽，避免误判为“未设置”。
        widths = {1: None, 2: None}
        for dim in ws.column_dimensions.values():
            lo, hi = getattr(dim, 'min', None), getattr(dim, 'max', None)
            if dim.width is None or lo is None or hi is None:
                continue
            for ci in (1, 2):
                if lo <= ci <= hi:
                    widths[ci] = dim.width
        info['col_w'] = {'A': widths[1], 'B': widths[2]}
        info['merged'] = [str(rg) for rg in ws.merged_cells.ranges]
        for row in ws.iter_rows(min_row=1, max_row=max(ws.max_row, 1),
                                max_col=max(ws.max_column, 2)):
            for cell in row:
                if cell.value is None:
                    continue
                if cell.column <= 2 and cell.row <= info['n'] + 1:
                    continue
                info['extra'].append({'coord': cell.coordinate, 'col': cell.column,
                                      'row': cell.row, 'is_str': isinstance(cell.value, str),
                                      'text': str(cell.value)[:60],
                                      'font': cell.font.name})
        info['ok'] = True
    except Exception as exc:                      # noqa: BLE001
        info['err'] = '%s: %s' % (type(exc).__name__, exc)
    finally:
        wb.close()
    return info
def _probe_zip(path):
    """用 zipfile 直接读原始 OOXML：图表事实、锚点尺寸、位图、结构完整性。"""
    info = {'ok': False, 'err': None, 'charts': [], 'anchors': [], 'pics': [],
            'sheet_texts': [], 'media': [], 'declared_charts': [],
            'sheets': []}
    try:
        with zipfile.ZipFile(path) as z:
            bad = z.testzip()
            if bad:
                info['err'] = '压缩包损坏：%s' % bad
                return info
            scheme = _theme_scheme(z)
            info['media'] = _media_names(z)
            info['declared_charts'] = _chart_parts(z)
            sheets = _sheet_parts(z)
            info['sheets'] = sheets
            first = sheets[0][1] if sheets else None
            for _name, part in sheets:
                dr = _parse_sheet_drawing(z, part, scheme)
                for ch in dr['charts']:
                    ch['sheet'] = part
                    ch['same_sheet'] = (part == first)
                    info['anchors'].append(ch)
                for p in dr['pics']:
                    info['pics'].append({'sheet': part, 'from': p})
                info['sheet_texts'].extend(dr['texts'])
            seen = set()
            for ch in info['anchors']:
                if ch['part'] and ch['part'] not in seen:
                    seen.add(ch['part'])
                    info['charts'].append(_parse_chart(z, ch['part'], scheme))
            for part in info['declared_charts']:
                if part not in seen:
                    seen.add(part)
                    # 有图表部件但没有任何工作表锚点引用，同样纳入图面事实。
                    info['charts'].append(_parse_chart(z, part, scheme))
            info['ok'] = True
    except (zipfile.BadZipFile, OSError) as exc:
        info['err'] = '%s: %s' % (type(exc).__name__, exc)
    return info


def _probe(path):
    """单个交付文件的完整探测结果。"""
    p = {'path': path, 'name': os.path.basename(path)}
    p['cells'] = _probe_cells(path)
    p['zip'] = _probe_zip(path)
    p['openable'] = bool(p['cells']['ok'] and p['zip']['ok'])
    return p
def _list_books(dir_path):
    """列出目录内待评的工作簿（忽略 ~$ 锁文件与本脚本自身产生的临时文件）。"""
    out = []
    try:
        names = sorted(os.listdir(dir_path))
    except OSError:
        return out
    for nm in names:
        if nm.startswith('~$') or nm.startswith('.'):
            continue
        if os.path.splitext(nm)[1].lower() not in ('.xlsx', '.xlsm'):
            continue
        full = os.path.join(dir_path, nm)
        if os.path.isfile(full):
            out.append(full)
    return out


def _match_index(cells):
    """按数据指纹归位到表1/表2/表3：数据行数 22/25/28 互不相同，可作主键；
       行数不匹配时退回“频率-S₁₁ 数值命中最多”的表号。"""
    n = cells['n']
    same = [i for i in (1, 2, 3) if len(_EXPECTED[i]['freq']) == n]
    if len(same) == 1:
        return same[0]
    best, best_hit = None, 0
    for i in (1, 2, 3):
        exp = _EXPECTED[i]
        hit = sum(1 for a, b in zip(cells['freq'], exp['freq']) if _close(a, b, 1e-6))
        hit += sum(1 for a, b in zip(cells['s11'], exp['s11']) if _close(a, b, 1e-6))
        if hit > best_hit:
            best, best_hit = i, hit
    return best


def _fname_index(name):
    """从文件名推断其对应的表号（供“文件名可一一对应”判定）。"""
    stem = os.path.splitext(os.path.basename(name))[0]
    m = re.search(r'表\s*([123])', stem)
    if m:
        return int(m.group(1))
    m = re.search(r'(?<!\d)([123])(?!\d)', stem)
    return int(m.group(1)) if m else None
def _main_chart(p):
    """取该文件的主图表事实与其工作表锚点（优先被工作表引用的第一个图表）。"""
    z = p['zip']
    charts = {c['part']: c for c in z['charts']}
    for a in z['anchors']:
        if a.get('part') and a['part'] in charts:
            return charts[a['part']], a
    return ((z['charts'][0], None) if z['charts'] else (None, None))


def _is_flat_ref(s):
    """该系列是否为“整条 -10”的参考线系列。"""
    ys = [v for v in s['y_vals'] if v is not None]
    return bool(ys) and all(_close(v, _REF_LINE_Y, 1e-6) for v in ys)


def _main_series(ch):
    """取承载 S₁₁ 曲线的主系列：排除整条 -10 的参考线系列，取点数最多者。"""
    if ch is None or not ch['series']:
        return None
    ranked = sorted(ch['series'], key=lambda s: (0 if _is_flat_ref(s) else 1, s['y_n']),
                    reverse=True)
    return ranked[0]


def _ref_series(ch):
    """取被做成额外数据系列的 -10 dB 参考线（若存在）。"""
    if ch is None:
        return None
    main = _main_series(ch)
    for s in ch['series']:
        if s is not main and _is_flat_ref(s):
            return s
    return None


def _axis_by_pos(ch, poses):
    """按 axPos 取轴（横轴 'b'/'t'，纵轴 'l'/'r'）；取不到返回 None。"""
    if ch is None:
        return None
    for ax in ch['axes']:
        if ax['pos'] in poses:
            return ax
    return None


def _pt(emu):
    """EMU → 磅，保留两位，便于在取证说明里展示。"""
    return None if emu is None else round(emu / EMU_PER_PT, 2)
def _check_add(p, idx):
    """逐文件判定 18 项加分点；返回 {键: (是否满足, 说明)}。"""
    c, z = p['cells'], p['zip']
    ch, anch = _main_chart(p)
    ser = _main_series(ch)
    n_exp = len(_EXPECTED[idx]['freq']) if idx else 0
    r = {}
    r['A01'] = (bool(idx) and c['n'] > 0 and ch is not None,
                '归位=表%s，数据%d行，图表部件%d个' % (idx, c['n'], len(z['charts'])))
    data_part = None
    for nm, part in z['sheets']:
        if nm == c['ws_name']:
            data_part = part
    r['A02'] = (anch is not None and anch.get('sheet') == data_part,
                '图表锚定于 %s，数据所在工作表 %s' % (
                    (anch.get('sheet') if anch else None), data_part))
    r['A03'] = (ch is not None and bool(ch['kinds']),
                '图表部件=%s，图表组=%s，包内位图%d张' % (
                    (ch['part'] if ch else None),
                    (','.join(ch['kinds']) if ch else '无'), len(z['media'])))
    xf = (ser['x_formula'] if ser else None) or ''
    yf = (ser['y_formula'] if ser else None) or ''
    r['A04'] = (bool(ser) and ser['x_ref'] and ser['y_ref'] and '!' in xf and '!' in yf,
                'X 引用=%s；Y 引用=%s' % (xf or '无（数值被硬编码）', yf or '无（数值被硬编码）'))
    r['A05'] = (bool(ch) and 'scatterChart' in ch['kinds'] and bool(ser)
                and ser['x_tag'] == 'xVal' and not ser['line_nofill'],
                '类型=%s，scatterStyle=%s，X 来源=%s' % (
                    (','.join(ch['kinds']) if ch else '无'),
                    (ch['scatter_style'] if ch else None),
                    (ser['x_tag'] if ser else None)))
    r['A06'] = (bool(ser) and ser['x_n'] == n_exp and ser['y_n'] == n_exp,
                '绘入点数 X=%s / Y=%s，应为 %d' % (
                    (ser['x_n'] if ser else 0), (ser['y_n'] if ser else 0), n_exp))
    mk_ok = bool(ser) and ser['marker'] == 'circle' and (
        ser['marker_size'] is None
        or _MARKER_SIZE_MIN <= ser['marker_size'] <= _MARKER_SIZE_MAX)
    mk_same = bool(ser) and (ser['marker_color'] is None
                             or ser['marker_color'] == ser['line_color']
                             or _is_blue(ser['marker_color']))
    r['A07'] = (bool(mk_ok and mk_same), '标记=%s，直径=%s磅，标记色=%s，线色=%s' % (
        (ser['marker'] if ser else None), (ser['marker_size'] if ser else None),
        (ser['marker_color'] if ser else None), (ser['line_color'] if ser else None)))
    if ser is None:
        r['A08'] = (False, '未找到数据系列')
    else:
        ok = (_is_blue(ser['line_color']) and ser['line_dash'] in (None, 'solid')
              and ser['line_w'] is not None
              and _LINE_W_MIN <= ser['line_w'] <= _LINE_W_MAX
              and ser['smooth'] is not True and not ser['line_nofill'])
        r['A08'] = (ok, '线色=%s，线宽=%s磅，线型=%s，平滑=%s' % (
            ser['line_color'], _pt(ser['line_w']), ser['line_dash'] or 'solid',
            ser['smooth']))
    xax, yax = _axis_by_pos(ch, ('b', 't')), _axis_by_pos(ch, ('l', 'r'))
    xs = [v for v in (ser['x_vals'] if ser else []) if v is not None]
    ok = bool(xax) and not xax['is_cat'] and not xax['deleted'] \
        and (_close(xax['min'], 0.0, _AXIS_TOL)
             or (xax['min'] is None and bool(xs) and _close(min(xs), 0.0, _AXIS_TOL))) \
        and _close(xax['max'], 24.0, _AXIS_TOL) \
        and _close(xax['major_unit'], 4.0, _AXIS_TOL)
    r['A09'] = (ok, '横轴 min=%s max=%s 主单位=%s 类别轴=%s' % (
        (xax['min'] if xax else None), (xax['max'] if xax else None),
        (xax['major_unit'] if xax else None), (xax['is_cat'] if xax else None)))
    ok = bool(yax) and not yax['deleted'] and _close(yax['min'], -24.0, _AXIS_TOL) \
        and _close(yax['max'], -8.0, _AXIS_TOL) \
        and _close(yax['major_unit'], 4.0, _AXIS_TOL)
    r['A10'] = (ok, '纵轴 min=%s max=%s 主单位=%s' % (
        (yax['min'] if yax else None), (yax['max'] if yax else None),
        (yax['major_unit'] if yax else None)))
    x_title = _norm_text(xax['title'] if xax else '')
    extra_txt = [_norm_text(t) for t in (z['sheet_texts'] + (ch['shape_texts'] if ch else []))]
    ok = x_title == '频率(GHz)' or any(t == '频率(GHz)' for t in extra_txt)
    r['A11'] = (ok, '横轴标题=%s，图面文本=%s' % (
        (xax['title'] if xax else None) or '无', extra_txt or '无'))
    y_title = _norm_text(yax['title'] if yax else '')
    ok = y_title.startswith('S11') or any(t.startswith('S11') for t in extra_txt)
    r['A12'] = (ok, '纵轴标题=%s' % ((yax['title'] if yax else None) or '无'))
    gl = [l for l in (ch['shape_lines'] if ch else [])
          if l['horiz'] and _is_gray(l['color']) and l['dash'] in _DASH_STYLES]
    rs = _ref_series(ch)
    rs_ok = rs is not None and (rs['line_dash'] in _DASH_STYLES or _is_gray(rs['line_color']))
    r['A13'] = (bool(gl) or rs_ok, '图上水平虚线%d条（%s）；-10 参考线系列=%s' % (
        len(gl), (gl[0] if gl else '无'),
        ('有（%s/%s）' % (rs['line_color'], rs['line_dash']) if rs else '无')))
    r['A14'] = (ch is not None and not ch['has_legend'],
                '图例=%s' % ('显示' if (ch and ch['has_legend']) else '未显示'))
    r['A15'] = (ch is not None and not ch['show_val'],
                '数据标签=%s' % ('显示' if (ch and ch['show_val']) else '未显示'))
    fill = ch['space_fill'] if ch else None
    ok = bool(ch) and (ch['space_nofill'] or fill is None or _is_light(fill)) \
        and (ch['space_line_w'] is None or ch['space_line_w'] <= 25400)
    r['A16'] = (ok, '图表区填充=%s，外框线宽=%s磅' % (
        fill or ('无填充' if (ch and ch['space_nofill']) else '默认'),
        _pt(ch['space_line_w']) if ch else None))
    cx = anch.get('cx') if anch else None
    cy = anch.get('cy') if anch else None
    w_cm = round(cx / EMU_PER_CM, 2) if cx else None
    h_cm = round(cy / EMU_PER_CM, 2) if cy else None
    ok = (w_cm is not None and h_cm is not None
          and _CHART_W_MIN <= w_cm <= _CHART_W_MAX and _CHART_H_MIN <= h_cm <= _CHART_H_MAX)
    r['A17'] = (ok, '图表尺寸 宽%scm × 高%scm' % (w_cm, h_cm))
    frm = anch.get('from') if anch else None
    ok = bool(frm) and (frm[0] >= 2 or frm[1] >= c['n'] + 1)
    r['A18'] = (ok, '锚点起始单元格 列索引=%s 行索引=%s（数据占 A/B 列前 %d 行）' % (
        (frm[0] if frm else None), (frm[1] if frm else None), c['n'] + 1))
    return r
def _check_ded(p, idx):
    """逐文件判定 20 项减分点；返回 {键: (是否触发, 说明)}。触发=惹了麻烦。"""
    c, z = p['cells'], p['zip']
    exp = _EXPECTED[idx] if idx else None
    n_exp = len(exp['freq']) if exp else 0
    d = {}
    diff = []
    if exp:
        for i, (a, b) in enumerate(zip(c['freq'], exp['freq'])):
            if not _close(a, b, 1e-6):
                diff.append('第%d行频率 %s→%s' % (i + 2, b, a))
        for i, (a, b) in enumerate(zip(c['s11'], exp['s11'])):
            if not _close(a, b, 1e-6):
                diff.append('第%d行S₁₁ %s→%s' % (i + 2, b, a))
    d['D01'] = (bool(diff), '；'.join(diff[:4]) if diff else '频率与 S₁₁ 数值及行序与输入一致')
    hdr_ok = tuple(t.strip() for t in c['headers']) == _HEADER_TEXTS
    d['D02'] = ((c['n'] != n_exp) or not hdr_ok,
                '数据行=%d（应%d），表头=%s' % (c['n'], n_exp, c['headers']))
    ext_in = [e for e in c['extra'] if e['col'] >= 3 and e['row'] <= c['n'] + 1]
    inset = {e['coord'] for e in ext_in}
    d['D03'] = (bool(ext_in), '数据区右侧新增单元格 %s' % (
        [e['coord'] for e in ext_in[:4]] or '无'))
    notes = [e for e in c['extra'] if e['is_str'] and e['coord'] not in inset]
    d['D04'] = (bool(notes), '新增说明性文字 %s' % (
        '；'.join('%s=%s' % (e['coord'], e['text']) for e in notes[:2]) or '无'))
    d['D05'] = (bool(c['merged']), '合并单元格区域=%s' % (c['merged'] or '无'))
    d['D06'] = (len(c['sheets']) != 1 or c['sheets'][0] != 'Sheet1',
                '工作表列表=%s' % c['sheets'])
    d['D07'] = (not (_close(c['col_w'].get('A'), _COL_WIDTH, 0.5)
                     and _close(c['col_w'].get('B'), _COL_WIDTH, 0.5)),
                '列宽 A=%s B=%s（应16字符）' % (c['col_w'].get('A'), c['col_w'].get('B')))
    hbad = [h for h in c['data_h'] if not _close(h, _DATA_ROW_HEIGHT, 0.5)]
    d['D08'] = ((not _close(c['row1_h'], _ROW1_HEIGHT, 0.5)) or bool(hbad),
                '第1行行高=%s（应26磅），数据行行高不合规%d处（应21磅）' % (c['row1_h'], len(hbad)))
    hb = [st for st in c['hdr']
          if st['font'] != _TABLE_FONT or not _close(st['size'], _HEADER_SIZE, 0.01)
          or not st['bold'] or st['color'] != 'FFFFFF'
          or st['halign'] != 'center' or st['valign'] != 'center']
    d['D09'] = (bool(hb), '表头字体/字色/居中=%s' % [
        (st['font'], st['size'], st['bold'], st['color'], st['halign'], st['valign'])
        for st in c['hdr']])
    fills = [st['fill'] for st in c['hdr']]
    d['D10'] = (bool(exp) and any(f != exp['header_fill'] for f in fills),
                '表头底纹=%s（表%s应为 %s）' % (
                    fills, idx, exp['header_fill'] if exp else '?'))
    db = [st for st in c['data']
          if st['font'] != _TABLE_FONT or not _close(st['size'], _DATA_SIZE, 0.01)
          or st['color'] != _DATA_FONT_COLOR
          or st['halign'] != 'center' or st['valign'] != 'center']
    d['D11'] = (bool(db), '数据单元格样式不合规%d处（样例=%s）' % (
        len(db), (db[0]['font'], db[0]['size'], db[0]['color']) if db else None))
    band_bad = []
    if exp:
        for row, fa, fb in c['bands']:
            want = exp['band_fill'] if row % 2 == 0 else 'FFFFFF'
            if fa != want or fb != want:
                band_bad.append(row)
    d['D12'] = (bool(band_bad), '隔行底纹异常行=%s（应 %s 与白色交替）' % (
        band_bad[:5], exp['band_fill'] if exp else '?'))
    fbad = [st['numfmt'] for st in c['data'] if st['numfmt'] != _DATA_NUMFMT]
    d['D13'] = (bool(fbad), '数字格式不为 0.0 的单元格%d个（如 %s）' % (
        len(fbad), fbad[0] if fbad else '—'))
    bd = []
    if len(c['hdr']) == 2:
        if tuple(c['hdr'][0]['border']) != _BD_HDR_A:
            bd.append('A1 框线=%s' % (c['hdr'][0]['border'],))
        if tuple(c['hdr'][1]['border']) != _BD_HDR_B:
            bd.append('B1 框线=%s' % (c['hdr'][1]['border'],))
    bd += ['数据区框线=%s' % (st['border'],)
           for st in c['data'] if tuple(st['border']) != _BD_DATA][:1]
    d['D14'] = (bool(bd), '；'.join(bd) if bd else '表头与数据区框线与输入一致')
    ffx = [e for e in c['extra'] if e['font'] and e['font'] != _TABLE_FONT]
    d['D15'] = (bool(ffx), '新增单元格字体=%s' % (
        [(e['coord'], e['font']) for e in ffx[:3]] or '无新增内容'))
    ch, _anch = _main_chart(p)
    sers = ch['series'] if ch else []
    d['D16'] = (len(sers) > 1, '数据系列%d个（%s）' % (
        len(sers), [s['name'] or '未命名' for s in sers]))
    deco = []
    # 轴标题与图/曲线名（S11 一类）属指令要求的图面元素，不计入“多余装饰”；
    # 其余文本（说明语、水印等）以及深色底纹、粗边框才计。
    axis_txt = {'频率(GHz)', 'S11', 'S11(dB)'}
    title_txt = _norm_text(ch['title'] if ch else '')
    for t in (ch['shape_texts'] if ch else []) + z['sheet_texts']:
        nt = _norm_text(t)
        if nt in axis_txt or nt.startswith('S11') or (title_txt and nt == title_txt):
            continue
        deco.append(t.strip()[:24])
    if ch and ch['space_fill'] and not _is_light(ch['space_fill']):
        deco.append('图表区深色填充 %s' % ch['space_fill'])
    if ch and ch['space_line_w'] and ch['space_line_w'] > 25400:
        deco.append('图表区粗边框 %s磅' % _pt(ch['space_line_w']))
    d['D17'] = (bool(deco), '图面多余装饰=%s' % (deco[:3] or '无'))
    d['D18'] = (False, '文件名归属在汇总阶段判定')
    return d
_ADD_RULES = (
    '交付3个 Excel 文件，分别承载表1、表2、表3的频率-S₁₁ 数据，一表一图、三份齐全',
    '每张图与其数据处于同一工作簿的同一工作表中',
    '图为 Excel 原生图表对象（可选中后改类型、改样式、改数据范围），而非位图或形状拼图',
    '图表数据系列以单元格区域引用绑定（形如 Sheet1!$A$2:$A$23 与 Sheet1!$B$2:$B$23）',
    '图表类型为“带直线和数据标记的散点图”，X 值取频率列、Y 值取 S₁₁ 列',
    '数据点全部绘入，表1为22个点、表2为25个点、表3为28个点',
    '每个数据点带圆形标记（circle，直径约5磅），标记与线同色',
    '曲线为蓝色实线、线宽约1.5~2磅，点与点之间为直线段连接（不做曲线平滑）',
    '横轴显示范围 0~24，主刻度单位为4',
    '纵轴显示范围 -24~-8，主刻度单位为4',
    '横轴带轴标题“频率（GHz）”',
    '纵轴带轴标题“S11”',
    '图中含 -10 dB 处的水平参考线，样式为灰色虚线，横贯绘图区',
    '图表不显示图例（图面无图例区块）',
    '数据点不显示数据标签（图面无逐点数值文字）',
    '图表区为白色或浅色底、无深色填充、无粗重外框',
    '图表尺寸适中（宽约14~22cm、高约8~13cm），曲线与两轴刻度完整可见',
    '图表锚定在数据区右侧或下方的空白单元格区域，未覆盖 A、B 两列的表头与数据',
)
_DED_RULES = (
    '频率或 S₁₁ 数值被修改、被四舍五入到整数、或行的先后顺序被调整',
    '数据行被删减（表1应22行、表2应25行、表3应28行），或表头文字不再为“频率 (GHz)”与“S₁₁(dB)”',
    '在数据区旁新增额外数据列，使两列数据表被扩成三列',
    '在工作表单元格中新增说明性文字',
    '新增合并单元格区域',
    '新增与绘图无关的工作表，或工作表名不再为 Sheet1',
    'A、B 两列列宽不满足16字符',
    '第1行行高不满足26磅，或第2行起各数据行行高不满足21磅',
    '表头单元格“频率 (GHz)”与“S₁₁(dB)”不满足微软雅黑11磅加粗、白色字、水平垂直居中',
    '表头底纹色被改动或被统一成同一颜色（表1为 1F4E78、表2为 5B3A70、表3为 2F6B4F）',
    '数据单元格文字不满足微软雅黑10磅、字色 1F2937、水平垂直居中',
    '数据行隔行底纹被改动（表1为 EAF3F8、表2为 F1EAF6、表3为 EAF4EE 与白色交替）',
    '数据单元格数字格式不满足保留一位小数（0.0）',
    '表格框线粗细被改动（表头上下与表格外框为中粗线、数据区内部为细线）',
    '新增列/新增单元格所用字体与表内不一致',
    '图表中加入指令未要求的额外数据系列（如把 -10 dB 参考线做成第二个数据系列）',
    '图表中加入指令未要求的水印、说明文本框、深色底纹、粗边框等装饰元素',
    '交付文件名无法与表1、表2、表3建立一一对应',
)
_RULES = _ADD_RULES + _DED_RULES
assert len(_RULES) == len(_MAX_DELTAS_ORDER)
def _aggregate(probes):
    """把各交付文件的逐项判定汇总：加分项须三份齐满足，减分项任一份触发即计一次。"""
    idx_map = {}
    for p in probes:
        p['idx'] = _match_index(p['cells']) if p['cells']['ok'] else None
        if p['idx'] and p['idx'] not in idx_map:
            idx_map[p['idx']] = p
        p['add'] = _check_add(p, p['idx']) if p['cells']['ok'] else {}
        p['ded'] = _check_ded(p, p['idx']) if p['cells']['ok'] else {}
    full = all(i in idx_map for i in (1, 2, 3))
    adds, deds = {}, {}
    for i in range(1, _ADD_COUNT + 1):
        key = 'A%02d' % i
        oks, dts = [], []
        for t in (1, 2, 3):
            p = idx_map.get(t)
            if p is None:
                oks.append(False)
                dts.append('表%d：未找到对应交付文件' % t)
                continue
            ok, dt = p['add'].get(key, (False, '未判定'))
            oks.append(bool(ok))
            dts.append('表%d：%s' % (t, dt))
        adds[key] = (full and all(oks), '｜'.join(dts))
    for i in range(1, len(_DED_DELTAS_ORDER) + 1):
        key = 'D%02d' % i
        fired, dts = [], []
        for p in probes:
            hit, dt = p['ded'].get(key, (False, '未判定'))
            if hit:
                fired.append(p['name'])
            dts.append('%s：%s' % (p['name'], dt))
        deds[key] = (bool(fired), ('触发文件=%s；' % '、'.join(fired) if fired else '')
                     + '｜'.join(dts))
    named, unnamed = {}, []
    for p in probes:
        fi = _fname_index(p['name'])
        if fi is None or fi in named or (p['idx'] and fi != p['idx']):
            unnamed.append(p['name'])
        else:
            named[fi] = p['name']
    deds['D18'] = (bool(unnamed) or len(named) < 3,
                   '文件名→表号=%s；无法一一对应=%s' % (
                       {k: named[k] for k in sorted(named)}, unnamed or '无'))
    return idx_map, adds, deds
def _check_dim1(probes, idx_map):
    """维度1 门禁：文件可正常打开、数据区可点选编辑、图为可选中的图表对象、
       版面未错乱（图表未压盖 A/B 两列数据）。任一条不满足则完成度总分记 0。"""
    bad = []
    for p in probes:
        c, z = p['cells'], p['zip']
        if not c['ok']:
            bad.append('%s 无法打开或读取失败（%s）' % (p['name'], c['err']))
            continue
        if not z['ok']:
            bad.append('%s 包结构损坏（%s）' % (p['name'], z['err']))
            continue
        if c['n'] <= 0:
            bad.append('%s 未读到可编辑的数据行（疑似数据表被压成图片）' % p['name'])
        if not z['charts'] and z['media']:
            # 只有位图、没有图表部件 → 图以截图形式贴入，不可编辑，属维度1 门禁项；
            # 若既无图表也无位图，只是没做绘图任务，由加分项记 0 体现，不触发门禁。
            bad.append('%s 无图表部件、仅有位图%d张（图以图片形式贴入，不可编辑）'
                       % (p['name'], len(z['media'])))
        for a in z['anchors']:
            frm = a.get('from')
            if frm and frm[0] <= 1 and frm[1] <= c['n']:
                bad.append('%s 图表锚定于 A/B 列数据区内，压盖数据导致不可读' % p['name'])
                break
    if not idx_map:
        bad.append('未能按数据指纹把任何交付文件归位到表1/表2/表3')
    if bad:
        return False, '；'.join(bad)
    return True, ('交付%d个 .xlsx 均可正常打开，数据区单元格可点选编辑，图为可选中的'
                  '原生图表对象，版面无错乱' % len(probes))


def _cut(text, limit=420):
    """文本过长时截断，避免结果体积失控。"""
    text = text or ''
    return text if len(text) <= limit else text[:limit] + '…'
def evaluate(dir_path):
    """按接口约定返回结构化 dict；不 print、不改 sys.stdout、不 sys.exit。
    入参 dir_path 为“脚本所在目录的路径”，脚本自己在该目录里定位并打开被评估的
    3 个 .xlsx（忽略以 ~$ 开头的 Excel 临时锁文件），并按数据指纹归位到表1/表2/表3。
    目录不存在或目录内没有 .xlsx 时，返回“未找到文件”的空结果（各评分字段留空/归零）；
    其余脚本自身崩溃由通用兜底捕获并返回 status='error'。"""
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
            raise FileNotFoundError('目录不存在或不是目录：%s' % dir_path)
        files = _list_books(dir_path)
        if not files:
            raise FileNotFoundError('目录中未找到 .xlsx 文件：%s' % dir_path)
        result["file_name"] = '、'.join(os.path.basename(f) for f in files)
        return _evaluate_impl(files, result)
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
    except Exception as exc:                      # noqa: BLE001 - 崩溃兜底
        result["status"] = "error"
        result["error"] = '脚本运行异常：%s: %s' % (type(exc).__name__, exc)
        result["dim1_pass"] = False
        result["dim1_reason"] = ('维度一未通过：文件无法打开或解析（%s: %s），完成度总分记 0'
                                % (type(exc).__name__, exc))
        result["dim2_items"] = []
        result["total_score"] = 0
        return result
def _evaluate_impl(files, result):
    """维度一 + 维度二的评估主体：先逐文件探测，再汇总为 results 列表，
    末尾按 _MAX_DELTAS_ORDER 反解成结构化 dim2_items 填入 result。"""
    probes = [_probe(f) for f in files]
    idx_map, adds, deds = _aggregate(probes)
    ok1, why1 = _check_dim1(probes, idx_map)
    result["dim1_pass"] = ok1
    if ok1:
        result["dim1_reason"] = _cut(why1, 600)
    else:
        # dim1_reason 统一带上「维度一未通过：…」前缀与计分后果，便于直接作为错误摘要展示
        result["dim1_reason"] = '维度一未通过：%s，完成度总分记 0' % _cut(why1, 600)

    # ========== 维度2：完成度评分 ==========
    # results 逐条形如 "+N：规则说明"；前 18 条为加分项，后 18 条为减分项。
    # 各判定内部产出的取证说明只用于命中判断，不随结果返回。
    results = []
    for i in range(1, _ADD_COUNT + 1):
        hit, _dt = adds['A%02d' % i]
        delta = _ADD_DELTAS_ORDER[i - 1] if hit else 0
        results.append('%+d：%s' % (delta, _ADD_RULES[i - 1]))
    for i in range(1, len(_DED_DELTAS_ORDER) + 1):
        hit, _dt = deds['D%02d' % i]
        delta = _DED_DELTAS_ORDER[i - 1] if hit else 0
        results.append('%+d：%s' % (delta, _DED_RULES[i - 1]))

    # ========== 把 results 反解为结构化 dim2_items ==========
    # 逐项按追加顺序对齐 _MAX_DELTAS_ORDER：解析每条 "+N：规则说明" 的 N 与 rule。
    if len(results) != len(_MAX_DELTAS_ORDER):
        raise RuntimeError('dim2 项数不一致：results=%d vs order=%d'
                           % (len(results), len(_MAX_DELTAS_ORDER)))
    dim2_items = []
    total_score = 0
    for raw, max_delta in zip(results, _MAX_DELTAS_ORDER):
        head, _, rule = raw.partition('：')
        try:
            delta = int(head)
        except ValueError:
            delta = 0
        hit = delta != 0
        # 加分项：命中记 max_delta（正），未命中记 0；
        # 减分项：触发记 max_delta（负），未触发记 0；维度1 不过则一律记 0。
        eff = (max_delta if hit else 0) if ok1 else 0
        dim2_items.append({
            "rule": rule,
            "max_delta": max_delta,
            "delta": eff,
            "hit": hit,
        })
        total_score += eff

    result["dim2_items"] = dim2_items
    result["total_score"] = total_score
    result["max_score"] = MAX_SCORE_TOTAL
    return result


if __name__ == "__main__":
    # 本地调试入口：evaluate 只接收“脚本所在目录的路径”，由脚本自己在该目录里定位
    # 并打开被评估的 .xlsx。命令行第 1 个参数可覆盖目录路径；未指定时使用脚本自身
    # 所在目录。注：evaluate() 的接口约定就是“入参为目录路径”——路径始终由外部显式
    #     传入，这里的默认值仅为本地调试便利，不影响 evaluate 的接口约定。
    if len(sys.argv) >= 2:
        _dir = sys.argv[1]
    else:
        _dir = os.path.dirname(os.path.abspath(__file__))
    _out = evaluate(_dir)
    print(json.dumps(_out, ensure_ascii=False, indent=2))
