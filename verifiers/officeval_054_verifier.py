#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# pyright: reportUnknownMemberType=false, reportUnknownVariableType=false
# pyright: reportUnknownArgumentType=false, reportMissingImports=false
# pyright: reportOptionalMemberAccess=false, reportGeneralTypeIssues=false

import os
import re
import sys
import json
import math
import zipfile
import colorsys

from pptx import Presentation

# 说明：按接口统一约定，本模块不修改全局 sys.stdout，也不 print 主结果。
# 主结果统一由 evaluate(dir_path) -> dict 返回；入参为“脚本所在目录的路径”，
# 由脚本自己在该目录内定位并打开被评估的 .pptx。本文件仅在 __main__ 下作为
# 本地调试入口，才把 dict 序列化为 JSON 打印。
#
# 被评估对象：按流程图内容复刻的单页可编辑 PPT（工业园中水回用处理流程与水质保障路径）。
# 全部判定都读 OOXML（PresentationML/DrawingML）原始属性，与 PowerPoint/WPS
# 打开后所见效果对齐：形状树 p:sp/p:pic/p:cxnSp/p:grpSp、位置尺寸 a:xfrm/a:off/a:ext、
# 组合缩放 a:chOff/a:chExt、几何 a:prstGeom@prst、描边 a:ln@w 与 a:prstDash、
# 竖排 a:bodyPr@vert、主题东亚字体 a:fontScheme/a:ea、填充 a:solidFill/a:gradFill。

SCRIPT_ID = "054"
# 维度二所有评分项 max_delta 之和；用于维度一未通过 / 脚本异常时兜底
MAX_SCORE_TOTAL = 97

# 顺序表：每一项对应 _evaluate_impl 内一次 results.append(...)，与 results.append(...)
# 的顺序严格一致；共 27 项，累加 = MAX_SCORE_TOTAL。
_MAX_DELTAS_ORDER = [
    3,  # 01 幻灯片页面采用标准演示尺寸
    1,  # 02 仅含1页幻灯片
    5,  # 03 整幅流程图整体高度7cm且比例未变形
    5,  # 04 图形水平垂直居中且四周留白
    3,  # 05 所有描边线条线宽统一0.75磅
    5,  # 06 中文统一宋体（东亚字体）
    3,  # 07 全图组合为一个组合对象
    3,  # 08 顶部深青绿标题条与居中白色加粗标题
    3,  # 09 标题字号≥流程框/底部说明文字的1.5倍
    1,  # 10 标题条青绿到深绿横向渐变
    3,  # 11 双层圆角框结构
    5,  # 12 10个流程框文字齐全准确
    3,  # 13 10个流程框等宽等高、顶端对齐、间距均匀
    5,  # 14 流程框内为真正的竖排文字
    5,  # 15 10个圆形序号标记1-10
    3,  # 16 9个深蓝色向右箭头串联
    5,  # 17 4条引出线连接4个浅绿底标签
    3,  # 18 4条引出线为虚线
    3,  # 19 回用清水池下向下箭头接绿色标签
    5,  # 20 完整“异常回流”回路且各段可正常渲染
    3,  # 21 “异常回流”标签压在回流横线中部
    5,  # 22 底部4组说明项文字齐全准确
    3,  # 23 4组说明项由一条长条承载并以3条竖虚线分隔
    1,  # 24 说明项主标题加粗且字号大于副标题
    5,  # 25 全图9个可编辑小图标齐备
    5,  # 26 全部文字完整显示、无溢出无压盖
    3,  # 27 青绿加深蓝双色系配色统一
]
assert sum(_MAX_DELTAS_ORDER) == MAX_SCORE_TOTAL

# ==================== 常量 ====================

_NS = {
    'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
    'p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
}

EMU_CM = 360000.0          # 1cm = 360000 EMU
EMU_PT = 12700.0           # 1pt  = 12700 EMU
LINE_W_075PT = 9525        # 0.75磅 = 9525 EMU

# 10个主流程框自左至右的文字
BOX_LABELS = [
    '厂区综合废水收集', '粗格栅与调节井', 'pH均衡池', '混凝沉淀池', '多介质过滤器',
    '活性炭吸附塔', '超滤膜组件', '消毒接触池', '回用清水池', '冷却补水系统',
]
# 4个辅助标签：(所属流程框文字, 标签文字)
TAG_SPECS = [
    ('pH均衡池', '酸碱调节'),
    ('混凝沉淀池', 'PAC/PAM投加'),
    ('超滤膜组件', '核心净化单元'),
    ('消毒接触池', '次氯酸钠投加'),
]
TITLE_TEXT = '工业园中水回用处理流程与水质保障路径'
GREEN_TAG_TEXT = '绿化与冲洗回用'
BACKFLOW_TEXT = '异常回流'
BACKFLOW_TO_BOX = '粗格栅与调节井'
REUSE_BOX = '回用清水池'
BOTTOM_TITLES = ['全过程在线监测', '水质达标保障', '设备定期维护', '数据记录追溯']
BOTTOM_SUBS = ['pH/浊度/余氯/电导率', '满足回用水质标准', '保障系统稳定运行', '管理可视化、可追溯']

# 常规演示页面尺寸（cm）：4:3、16:9（宽屏/全屏两种）、A4 横纵
STANDARD_SLIDE_SIZES = [
    (25.4, 19.05), (33.867, 19.05), (25.4, 14.288), (30.48, 22.86),
    (29.7, 21.0), (21.0, 29.7), (33.867, 25.4), (19.05, 25.4),
]
# 真正的竖排文字方向（a:bodyPr@vert）
VERT_MODES = ('eaVert', 'vert', 'vert270', 'wordArtVertRtl', 'mongolianVert')
# 结构性几何：不计入“小图标”候选
STRUCT_GEOMS = {
    'rect', 'roundRect', 'round1Rect', 'round2SameRect', 'round2DiagRect',
    'snip1Rect', 'line', 'straightConnector1', 'bentConnector2', 'bentConnector3',
    'bentConnector4', 'bentConnector5', 'curvedConnector2', 'curvedConnector3',
    'rightArrow', 'leftArrow', 'upArrow', 'downArrow', 'leftRightArrow',
    'bentArrow', 'bentUpArrow', 'uturnArrow',
}
ARROW_GEOM_DIR = {'rightArrow': 'r', 'leftArrow': 'l', 'upArrow': 'u', 'downArrow': 'd'}
ARROW_END_TYPES = ('triangle', 'arrow', 'stealth', 'diamond', 'oval', 'open')

# ==================== 通用小工具 ====================


def _q(tag):
    """'a:xfrm' -> '{ns}xfrm'，用于 lxml 查找。"""
    prefix, local = tag.split(':')
    return '{%s}%s' % (_NS[prefix], local)


def _norm_text(s):
    """文字比对口径：去除所有空白（含全角空格）、统一小写。
    仅用于“文字是否齐全准确”的比对，不改变字形判定。"""
    if not s:
        return ''
    return re.sub(r'[\s　 ]+', '', s).lower()


def _cjk(s):
    """文本中是否含中日韩统一表意文字（中文字符）。"""
    return bool(re.search(r'[一-鿿㐀-䶿]', s or ''))


def _visual_units(text):
    """文本视觉宽度（以“1个全角字宽”为单位）：ASCII 半角≈0.5，其余≈1.0。
    与办公软件排版口径一致（中文全角占一个字宽，西文字母占半个字宽）。"""
    w = 0.0
    for ch in text or '':
        o = ord(ch)
        if 0x20 <= o <= 0x2ff:
            w += 0.5
        else:
            w += 1.0
    return w


def _seq_value(text):
    """把序号文本解析为 1-10 的整数：支持阿拉伯数字、全角数字与各种带圈数字符号。
    非序号返回 None。"""
    t = re.sub(r'[\s　.、,，)）(（]+', '', text or '')
    if not t:
        return None
    # 全角数字 -> 半角
    t = ''.join(chr(ord(c) - 0xfee0) if 0xff10 <= ord(c) <= 0xff19 else c for c in t)
    if re.fullmatch(r'[0-9]+', t):
        v = int(t)
        return v if 1 <= v <= 10 else None
    if len(t) != 1:
        return None
    o = ord(t)
    for base in (0x2460, 0x2474, 0x2488, 0x2776, 0x2780, 0x278a, 0x24f5):
        if base <= o <= base + 9:
            return o - base + 1
    if o in (0x24ea, 0x24ff):
        return None
    return None


def _is_circled_glyph(text):
    """单字符带圈数字符号（①❶➊➀ 等）——本身即“圆形序号标记”的字形。"""
    t = (text or '').strip()
    if len(t) != 1:
        return False
    o = ord(t)
    return any(base <= o <= base + 9 for base in (0x2460, 0x2776, 0x2780, 0x278a, 0x24f5))

# ==================== 颜色解析与色系判定 ====================


def _hex_to_hsv(hexv):
    """'02807E' -> (h, s, v)，h 归一化到 [0,1)。"""
    r = int(hexv[0:2], 16) / 255.0
    g = int(hexv[2:4], 16) / 255.0
    b = int(hexv[4:6], 16) / 255.0
    return colorsys.rgb_to_hsv(r, g, b)


def _rgb(hexv):
    return (int(hexv[0:2], 16), int(hexv[2:4], 16), int(hexv[4:6], 16))


def _is_teal(hexv):
    """青绿/青蓝绿系（含深青绿）：色相 144°-200°，且有足够饱和度。"""
    if not hexv:
        return False
    h, s, v = _hex_to_hsv(hexv)
    return 0.40 <= h <= 0.56 and s >= 0.20 and v >= 0.12


def _is_dark_teal(hexv):
    if not _is_teal(hexv):
        return False
    return _hex_to_hsv(hexv)[2] <= 0.75


def _is_green(hexv):
    """绿/青绿系（比 _is_teal 放宽到偏黄绿，用于“绿色标签”判定）。"""
    if not hexv:
        return False
    h, s, v = _hex_to_hsv(hexv)
    if s < 0.05:
        r, g, b = _rgb(hexv)
        return g >= r and g >= b and (g - min(r, b)) >= 3
    return 0.28 <= h <= 0.56


def _is_dark_blue(hexv):
    """深蓝系：色相 198°-260°，明度偏低。"""
    if not hexv:
        return False
    h, s, v = _hex_to_hsv(hexv)
    return 0.55 <= h <= 0.72 and s >= 0.25 and v <= 0.75


def _is_light_green(hexv):
    """浅绿底：明度很高且绿通道为最大通道。"""
    if not hexv:
        return False
    r, g, b = _rgb(hexv)
    _, _, v = _hex_to_hsv(hexv)
    return v >= 0.85 and g >= r and g >= b and (g - min(r, b)) >= 2


def _is_light_blue(hexv):
    """浅蓝底：明度很高且蓝通道为最大通道。"""
    if not hexv:
        return False
    r, g, b = _rgb(hexv)
    _, _, v = _hex_to_hsv(hexv)
    return v >= 0.85 and b >= r and b >= g and (b - min(r, g)) >= 2


def _is_whiteish(hexv):
    """白底/近白底。"""
    if not hexv:
        return False
    r, g, b = _rgb(hexv)
    return min(r, g, b) >= 235


def _is_darkish(hexv):
    """深色文字（深蓝或近黑）：明度低。None 视为深色（继承主题正文色，默认近黑）。"""
    if hexv is None:
        return True
    return _hex_to_hsv(hexv)[2] <= 0.45

def _apply_lum(hexv, lum_mod, lum_off):
    """按 a:lumMod / a:lumOff 调整亮度，与办公软件对主题色的加深/变浅一致。"""
    r, g, b = [c / 255.0 for c in _rgb(hexv)]
    h, l, s = colorsys.rgb_to_hls(r, g, b)
    if lum_mod is not None:
        l *= int(lum_mod) / 100000.0
    if lum_off is not None:
        l += int(lum_off) / 100000.0
    l = max(0.0, min(1.0, l))
    r2, g2, b2 = colorsys.hls_to_rgb(h, l, s)
    return '%02X%02X%02X' % (int(round(r2 * 255)), int(round(g2 * 255)), int(round(b2 * 255)))


def _color_of(container, theme_colors):
    """从任意“可含颜色”的元素（a:solidFill / a:rPr 等）里取第一个颜色，
    返回 (hex, alpha)；解析 a:srgbClr / a:schemeClr（含 lumMod/lumOff）/ a:sysClr。"""
    if container is None:
        return (None, 1.0)
    for el in container.iter():
        tag = el.tag.split('}')[-1]
        if tag == 'srgbClr':
            alpha = el.find(_q('a:alpha'))
            a = int(alpha.get('val')) / 100000.0 if alpha is not None else 1.0
            return ((el.get('val') or '').upper(), a)
        if tag == 'schemeClr':
            hexv = theme_colors.get(el.get('val'))
            alpha = el.find(_q('a:alpha'))
            a = int(alpha.get('val')) / 100000.0 if alpha is not None else 1.0
            lm = el.find(_q('a:lumMod'))
            lo = el.find(_q('a:lumOff'))
            if hexv and (lm is not None or lo is not None):
                hexv = _apply_lum(hexv,
                                  lm.get('val') if lm is not None else None,
                                  lo.get('val') if lo is not None else None)
            return (hexv, a)
        if tag == 'sysClr':
            return ((el.get('lastClr') or 'FFFFFF').upper(), 1.0)
    return (None, 1.0)


def _read_theme(pptx_path):
    """读取主题：返回 (theme_colors, major_ea, minor_ea)。
    theme_colors 已按母版 p:clrMap 完成 bg1/tx1/... 到 lt1/dk1/... 的映射，
    即打开 PowerPoint 时“主题颜色”实际取到的值。"""
    theme_colors, major_ea, minor_ea = {}, '', ''
    try:
        with zipfile.ZipFile(pptx_path) as z:
            names = z.namelist()
            tname = next((n for n in names if re.match(r'ppt/theme/theme\d+\.xml$', n)), None)
            if tname:
                xml = z.read(tname).decode('utf-8', 'ignore')
                scheme = {}
                m = re.search(r'<a:clrScheme.*?</a:clrScheme>', xml, re.S)
                if m:
                    for name, body in re.findall(
                            r'<a:(dk1|lt1|dk2|lt2|accent[1-6]|hlink|folHlink)>(.*?)</a:\1>',
                            m.group(0), re.S):
                        c = re.search(r'val="([0-9A-Fa-f]{6})"', body)
                        if c:
                            scheme[name] = c.group(1).upper()
                        elif 'window"' in body or 'windowText' in body:
                            lc = re.search(r'lastClr="([0-9A-Fa-f]{6})"', body)
                            scheme[name] = (lc.group(1).upper() if lc else 'FFFFFF')
                theme_colors.update(scheme)
                mj = re.search(r'<a:majorFont>.*?</a:majorFont>', xml, re.S)
                mn = re.search(r'<a:minorFont>.*?</a:minorFont>', xml, re.S)
                if mj:
                    e = re.search(r'<a:ea typeface="([^"]*)"', mj.group(0))
                    major_ea = e.group(1) if e else ''
                if mn:
                    e = re.search(r'<a:ea typeface="([^"]*)"', mn.group(0))
                    minor_ea = e.group(1) if e else ''
            mst = next((n for n in names if re.match(r'ppt/slideMasters/slideMaster\d+\.xml$', n)), None)
            if mst:
                mxml = z.read(mst).decode('utf-8', 'ignore')
                cm = re.search(r'<p:clrMap([^>]*)/>', mxml)
                if cm:
                    for k, v in re.findall(r'(\w+)="(\w+)"', cm.group(1)):
                        if v in theme_colors:
                            theme_colors[k] = theme_colors[v]
    except Exception:
        pass
    return theme_colors, major_ea, minor_ea

# ==================== 形状树解析 ====================


def _grp_map(el):
    """组合的坐标映射：(off_x, off_y, ext_cx, ext_cy, chOff_x, chOff_y, chExt_cx, chExt_cy)。"""
    x = el.find(_q('p:grpSpPr') + '/' + _q('a:xfrm'))
    if x is None:
        return None
    off, ext = x.find(_q('a:off')), x.find(_q('a:ext'))
    cho, che = x.find(_q('a:chOff')), x.find(_q('a:chExt'))
    if None in (off, ext, cho, che):
        return None
    return (int(off.get('x')), int(off.get('y')), int(ext.get('cx')), int(ext.get('cy')),
            int(cho.get('x')), int(cho.get('y')), int(che.get('cx')), int(che.get('cy')))


def _raw_xfrm(el):
    """取形状自身的 a:xfrm（p:spPr/p:grpSpPr/p:xfrm 三种承载位置）。"""
    for path in (_q('p:spPr') + '/' + _q('a:xfrm'),
                 _q('p:grpSpPr') + '/' + _q('a:xfrm'),
                 _q('p:xfrm')):
        x = el.find(path)
        if x is not None:
            return x
    return el.find('.//' + _q('a:xfrm'))


def _geom_of(el):
    spPr = el.find(_q('p:spPr'))
    if spPr is None:
        return ''
    g = spPr.find(_q('a:prstGeom'))
    if g is not None:
        return g.get('prst') or ''
    if spPr.find(_q('a:custGeom')) is not None:
        return 'custGeom'
    return ''


def _fill_of(el, theme_colors):
    """形状填充：{'kind': 'solid'|'grad'|'none'|'blip'|'inherit', 'color', 'alpha', 'colors'}。"""
    spPr = el.find(_q('p:spPr'))
    out = {'kind': 'inherit', 'color': None, 'alpha': 1.0, 'colors': []}
    if spPr is None:
        return out
    if spPr.find(_q('a:noFill')) is not None:
        out['kind'] = 'none'
        return out
    sf = spPr.find(_q('a:solidFill'))
    if sf is not None:
        c, a = _color_of(sf, theme_colors)
        out.update(kind='solid', color=c, alpha=a)
        return out
    gf = spPr.find(_q('a:gradFill'))
    if gf is not None:
        cols = []
        for gs in gf.findall('.//' + _q('a:gs')):
            c, _a = _color_of(gs, theme_colors)
            if c:
                cols.append(c)
        out.update(kind='grad', color=(cols[0] if cols else None), colors=cols)
        return out
    if spPr.find(_q('a:blipFill')) is not None:
        out['kind'] = 'blip'
    return out


def _line_of(el, theme_colors):
    """形状描边：{'has_el','visible','w','color','dash','head','tail'}。
    visible 表示“打开办公软件能看见这条描边”（存在 a:ln 且不是 a:noFill）。"""
    out = {'has_el': False, 'visible': False, 'w': None, 'color': None,
           'dash': None, 'head': None, 'tail': None}
    spPr = el.find(_q('p:spPr'))
    if spPr is None:
        return out
    ln = spPr.find(_q('a:ln'))
    if ln is None:
        return out
    out['has_el'] = True
    out['w'] = int(ln.get('w')) if (ln.get('w') or '').lstrip('-').isdigit() else None
    d = ln.find(_q('a:prstDash'))
    out['dash'] = d.get('val') if d is not None else None
    he, te = ln.find(_q('a:headEnd')), ln.find(_q('a:tailEnd'))
    out['head'] = he.get('type') if he is not None else None
    out['tail'] = te.get('type') if te is not None else None
    if ln.find(_q('a:noFill')) is None:
        c, a = _color_of(ln, theme_colors)
        out['color'] = c
        out['visible'] = (a > 0.05)
    return out

def _text_of(el, theme_colors):
    """文本正文：返回 (paras, vert, wrap, algn, ins)。
    paras 为段落列表，每段是 run 列表：{'t','sz','b','ea','latin','color'}；
    ins 为 (lIns,tIns,rIns,bIns) 的 cm 值（缺省取 PowerPoint 默认内边距）。"""
    tx = el.find(_q('p:txBody'))
    if tx is None:
        tx = el.find('.//' + _q('p:txBody'))
    if tx is None:
        return ([], None, None, None, (0.254, 0.127, 0.254, 0.127))
    bp = tx.find(_q('a:bodyPr'))
    vert = wrap = None
    ins = [91440, 45720, 91440, 45720]
    if bp is not None:
        vert, wrap = bp.get('vert'), bp.get('wrap')
        for i, k in enumerate(('lIns', 'tIns', 'rIns', 'bIns')):
            v = bp.get(k)
            if v is not None and v.lstrip('-').isdigit():
                ins[i] = int(v)
    algn = None
    paras = []
    for p in tx.findall(_q('a:p')):
        pPr = p.find(_q('a:pPr'))
        if algn is None and pPr is not None and pPr.get('algn'):
            algn = pPr.get('algn')
        runs = []
        for r in p.findall(_q('a:r')):
            t = r.find(_q('a:t'))
            rPr = r.find(_q('a:rPr'))
            info = {'t': (t.text or '') if t is not None else '', 'sz': None, 'b': False,
                    'ea': None, 'latin': None, 'color': None}
            if rPr is not None:
                if (rPr.get('sz') or '').isdigit():
                    info['sz'] = int(rPr.get('sz')) / 100.0
                info['b'] = (rPr.get('b') == '1')
                ea, la = rPr.find(_q('a:ea')), rPr.find(_q('a:latin'))
                info['ea'] = ea.get('typeface') if ea is not None else None
                info['latin'] = la.get('typeface') if la is not None else None
                sf = rPr.find(_q('a:solidFill'))
                if sf is not None:
                    info['color'] = _color_of(sf, theme_colors)[0]
            runs.append(info)
        paras.append(runs)
    return (paras, vert, wrap, algn,
            (ins[0] / EMU_CM, ins[1] / EMU_CM, ins[2] / EMU_CM, ins[3] / EMU_CM))


def _collect(prs, theme_colors):
    """把第1页形状树拉平成记录列表；坐标按组合的 a:off/a:ext 与 a:chOff/a:chExt
    换算成幻灯片绝对坐标（cm），与办公软件中所见位置一致。
    另返回 (顶层形状数, 顶层组合数, 组合内后代数, 是否存在负值 ext)。"""
    recs = []
    tops = {'n': 0, 'groups': 0, 'in_group': 0}
    neg = []

    def walk(parent, chain, depth):
        for child in parent:
            kind = child.tag.split('}')[-1]
            if kind not in ('sp', 'pic', 'graphicFrame', 'cxnSp', 'grpSp'):
                continue
            if depth == 0:
                tops['n'] += 1
                if kind == 'grpSp':
                    tops['groups'] += 1
            else:
                tops['in_group'] += 1
            rec = _make_rec(child, kind, chain, depth, len(recs), theme_colors)
            if rec is not None:
                recs.append(rec)
                if rec['raw_cx'] < 0 or rec['raw_cy'] < 0:
                    neg.append(rec)
            if kind == 'grpSp':
                walk(child, chain + [_grp_map(child)], depth + 1)

    walk(prs.slides[0].shapes._spTree, [], 0)
    return recs, tops, neg

def _make_rec(el, kind, chain, depth, z, theme_colors):
    """构造单个形状记录（坐标已换算到幻灯片绝对坐标，单位 cm）。"""
    xf = _raw_xfrm(el)
    if xf is None:
        return None
    off, ext = xf.find(_q('a:off')), xf.find(_q('a:ext'))
    if off is None or ext is None:
        return None
    x, y = float(off.get('x')), float(off.get('y'))
    cx, cy = float(ext.get('cx')), float(ext.get('cy'))
    raw_cx, raw_cy = cx, cy
    # 逐层套用组合映射（由内层到外层）
    for gm in reversed(chain):
        if gm is None:
            continue
        gx, gy, gcx, gcy, chx, chy, chcx, chcy = gm
        sx = (gcx / chcx) if chcx else 1.0
        sy = (gcy / chcy) if chcy else 1.0
        x = gx + (x - chx) * sx
        y = gy + (y - chy) * sy
        cx, cy = cx * sx, cy * sy
    x1, x2 = sorted((x / EMU_CM, (x + cx) / EMU_CM))
    y1, y2 = sorted((y / EMU_CM, (y + cy) / EMU_CM))
    nv = el.find('.//' + _q('p:cNvPr'))
    paras, vert, wrap, algn, ins = _text_of(el, theme_colors)
    text = ''.join(r['t'] for p in paras for r in p)
    rec = {
        'el': el, 'kind': kind, 'depth': depth, 'z': z,
        'id': (nv.get('id') if nv is not None else ''),
        'name': (nv.get('name') if nv is not None else ''),
        'geom': _geom_of(el),
        'x1': x1, 'y1': y1, 'x2': x2, 'y2': y2,
        'w': x2 - x1, 'h': y2 - y1, 'cxc': (x1 + x2) / 2.0, 'cyc': (y1 + y2) / 2.0,
        'raw_cx': raw_cx, 'raw_cy': raw_cy,
        'rot': xf.get('rot'), 'flipH': xf.get('flipH') == '1', 'flipV': xf.get('flipV') == '1',
        'fill': _fill_of(el, theme_colors), 'line': _line_of(el, theme_colors),
        'paras': paras, 'text': text, 'vert': vert, 'wrap': wrap, 'algn': algn, 'ins': ins,
    }
    return rec


def _runs(rec):
    return [r for p in rec['paras'] for r in p]


def _max_sz(rec, default=18.0):
    szs = [r['sz'] for r in _runs(rec) if r['sz']]
    return max(szs) if szs else default


def _inside(rec, rect, pad=0.0):
    """rec 的中心是否落在 rect=(x1,y1,x2,y2) 内（可加 pad 放宽）。"""
    return (rect[0] - pad <= rec['cxc'] <= rect[2] + pad
            and rect[1] - pad <= rec['cyc'] <= rect[3] + pad)


def _rect_of(rec):
    return (rec['x1'], rec['y1'], rec['x2'], rec['y2'])


def _union(*rects):
    return (min(r[0] for r in rects), min(r[1] for r in rects),
            max(r[2] for r in rects), max(r[3] for r in rects))


def _x_overlap(a, b):
    """两个矩形在水平方向的重叠长度（cm）。"""
    return max(0.0, min(a[2], b[2]) - max(a[0], b[0]))


def _inter_area(a, b):
    return (max(0.0, min(a[2], b[2]) - max(a[0], b[0]))
            * max(0.0, min(a[3], b[3]) - max(a[1], b[1])))


def _arrow_dirs(rec):
    """形状表达的箭头指向集合：'r'/'l'/'u'/'d'。
    支持两种落地方式：① 箭头类自动图形（rightArrow/downArrow/upArrow）；
    ② 直线/连接符 + a:headEnd/a:tailEnd 箭头端点（含 flipH/flipV 翻转）。"""
    g = rec['geom']
    if g in ARROW_GEOM_DIR:
        d = ARROW_GEOM_DIR[g]
        if rec['flipH'] and d in 'rl':
            d = 'l' if d == 'r' else 'r'
        if rec['flipV'] and d in 'ud':
            d = 'd' if d == 'u' else 'u'
        return {d}
    dirs = set()
    ln = rec['line']
    horiz = rec['h'] <= 0.12 and rec['w'] > rec['h']
    verti = rec['w'] <= 0.12 and rec['h'] > rec['w']
    if not (horiz or verti):
        return dirs
    for end, at_start in ((ln['tail'], False), (ln['head'], True)):
        if not end or end == 'none' or end not in ARROW_END_TYPES:
            continue
        if horiz:
            d = 'r' if (at_start == rec['flipH']) else 'l'
        else:
            d = 'd' if (at_start == rec['flipV']) else 'u'
        dirs.add(d)
    return dirs

def _text_overflow(rec, tol=1.15):
    """估算“文字是否溢出所属形状/文本框”，与办公软件排版口径对齐：
      ① 可用区域 = 形状尺寸 - a:bodyPr 的四边内边距（lIns/tIns/rIns/bIns）；
      ② 1个全角字宽/字高 ≈ 字号磅值，行高 ≈ 1.2 倍字号；
      ③ 横排：按可用宽度折行，累计行数 × 行高 与可用高度比较（单行不判，
         单行文本框的垂直溢出在办公软件里不产生断词/裁切）；
      ④ 竖排（a:bodyPr@vert）：文字沿高度方向排列，用每列字数×字高与可用高度比较，
         列数×行宽与可用宽度比较。
    留 tol 倍宽容量，避免把“刚好排满”判成溢出。"""
    paras = [p for p in rec['paras']]
    if not paras:
        return False
    if len(re.sub(r'\s', '', rec['text'])) <= 1:
        return False           # 单字符（序号标记等）不判溢出
    availw = rec['w'] - rec['ins'][0] - rec['ins'][2]
    availh = rec['h'] - rec['ins'][1] - rec['ins'][3]
    if availw <= 0 or availh <= 0:
        return True
    if rec['vert'] in VERT_MODES:
        need_len, cols = 0.0, 0
        for p in paras:
            sz = max([r['sz'] for r in p if r['sz']] or [_max_sz(rec)]) / 72.0 * 2.54
            units = _visual_units(''.join(r['t'] for r in p))
            need_len = max(need_len, units * sz)
            cols += 1
            colw = sz * 1.2
        return need_len > availh * tol or cols * colw > availw * tol
    lines, line_h = 0, 0.0
    for p in paras:
        sz = max([r['sz'] for r in p if r['sz']] or [_max_sz(rec)]) / 72.0 * 2.54
        need_w = _visual_units(''.join(r['t'] for r in p)) * sz
        if rec['wrap'] == 'none':
            n = 1
        else:
            n = max(1, int(math.ceil(need_w / (availw * tol))))
        lines += n
        line_h = max(line_h, sz * 1.2)
    if lines < 2:
        return False
    return lines * line_h > availh * tol


# ==================== 对外接口 ====================


def evaluate(dir_path):
    """按接口约定返回结构化 dict；不 print、不改 sys.stdout、不 sys.exit。
    入参 dir_path 为“脚本所在目录的路径”，脚本自己在该目录里定位并打开被评估的
    .pptx（忽略以 ~$ 开头的临时锁文件）。脚本自身崩溃（含目录不存在、未发现
    文档等）由本函数捕获并返回 status='error'；此时未进入评估流程，dim1_pass
    置空字符串（既非 True 也非 False），max_score 一并记 0。"""
    result = {
        "id": SCRIPT_ID,
        "file_name": "",
        "status": "ok",
        "error": None,
        "dim1_pass": True,
        "dim1_reason": "",
        "dim2_items": [],
        "total_score": 0,
        "max_score": MAX_SCORE_TOTAL,
    }

    def _fail(msg):
        """未取到被评估文件 / 脚本异常：按约定格式返回。"""
        result["status"] = "error"
        result["error"] = msg
        result["dim1_pass"] = ""
        result["dim1_reason"] = ""
        result["dim2_items"] = []
        result["total_score"] = 0
        result["max_score"] = 0
        return result

    try:
        if not dir_path or not os.path.isdir(dir_path):
            return _fail(f"目录不存在：{dir_path}")
        cands = sorted(
            f for f in os.listdir(dir_path)
            if f.lower().endswith('.pptx') and not f.startswith('~$')
        )
        if not cands:
            return _fail(f"目录中未找到 .pptx 文件：{dir_path}")
        file_path = os.path.join(dir_path, cands[0])
        result["file_name"] = os.path.basename(file_path)
        return _evaluate_impl(file_path, result)
    except Exception as exc:
        return _fail(f"{type(exc).__name__}: {exc}")

def _evaluate_impl(file_path, result):
    """维度一 + 维度二的评估主体；内部构造 results 列表（每条形如 '+N：说明'），
    末尾按 _MAX_DELTAS_ORDER 反解成结构化 dim2_items 填入 result。"""
    theme_colors, major_ea, minor_ea = _read_theme(file_path)
    results = []

    # ========== 维度1：可用与可修改性 ==========
    # 仅保留一条硬性条件：交付文件是1个能正常打开的 .pptx（打不开/需修复/结构损坏 →
    # 完成度总分直接记0）。其余可用性表现全部交由维度2逐项体现，不再另设记0闸门。
    try:
        prs = Presentation(file_path)
        n_slides = len(prs.slides)
        recs, tops, negs = _collect(prs, theme_colors)
    except Exception as exc:
        result['dim1_pass'] = False
        result['dim1_reason'] = f"PPT 文件无法正常打开或结构损坏：{type(exc).__name__}: {exc}"
        result['dim2_items'] = []
        result['total_score'] = 0
        return result

    slide_w = prs.slide_width / EMU_CM
    slide_h = prs.slide_height / EMU_CM
    text_recs = [r for r in recs if r['text'].strip()]

    # ========== 共用识别：图形整体 / 流程框 / 标签 / 底部说明区 ==========
    # 图形整体范围：优先取“唯一顶层组合”的外框，否则取全部图元的并集外框
    grp_tops = [r for r in recs if r['kind'] == 'grpSp' and r['depth'] == 0]
    if len(grp_tops) == 1 and tops['n'] == 1:
        fig = _rect_of(grp_tops[0])
    else:
        fig = _union(*[_rect_of(r) for r in recs]) if recs else (0.0, 0.0, 0.0, 0.0)
    fig_w, fig_h = fig[2] - fig[0], fig[3] - fig[1]
    fig_area = max(fig_w * fig_h, 1e-6)
    all_text_norm = _norm_text(''.join(r['text'] for r in recs))

    def _texts_in(rect, pad=0.15):
        return [t for t in text_recs if _inside(t, rect, pad)]

    def _find_text_rec(target):
        """整形状文字与 target 相同的记录（文字承载体）。"""
        n = _norm_text(target)
        return next((t for t in text_recs if _norm_text(t['text']) == n), None)

    def _find_para(target):
        """文字与 target 相同的“段落”所在记录及其 run 列表（支持一个形状里多行）。"""
        n = _norm_text(target)
        for r in text_recs:
            for p in r['paras']:
                if _norm_text(''.join(x['t'] for x in p)) == n:
                    return (r, p)
        return (None, [])

    def _label_block(trec):
        """标签整体范围 = 文字承载体与其重叠的实底小形状（标签底板）的并集。"""
        if trec is None:
            return (None, [])
        fills = []
        ta = max(trec['w'] * trec['h'], 1e-6)
        for r in recs:
            if r['fill']['kind'] != 'solid' or r['fill']['alpha'] <= 0.5:
                continue
            if r['w'] > 4.5 * max(trec['w'], 0.1) or r['h'] > 4.0 * max(trec['h'], 0.1):
                continue
            inter = _inter_area(_rect_of(r), _rect_of(trec))
            if inter > 0 and inter >= 0.30 * min(r['w'] * r['h'], ta):
                fills.append(r)
        rect = _rect_of(trec)
        if fills:
            rect = _union(rect, *[_rect_of(f) for f in fills])
        return (rect, fills)

    # 主流程框：白底 + 青绿描边 + 圆角矩形 + 竖长条（相对整幅图的尺寸占比受限）
    boxes = [r for r in recs
             if r['geom'].startswith('round')
             and r['fill']['kind'] == 'solid' and r['fill']['alpha'] > 0.5
             and _is_whiteish(r['fill']['color'])
             and r['line']['visible'] and _is_teal(r['line']['color'])
             and r['h'] > r['w']
             and 0.02 <= (r['w'] / fig_w if fig_w else 0) <= 0.20
             and 0.15 <= (r['h'] / fig_h if fig_h else 0) <= 0.65]
    boxes = sorted(boxes, key=lambda r: r['cxc'])
    box_of = {}          # 流程框文字 -> 框记录
    box_text_of = {}     # 流程框文字 -> 该框文字的承载体记录
    for b in boxes:
        for t in ([b] if b['text'].strip() else []) + _texts_in(_rect_of(b), 0.05):
            n = _norm_text(t['text'])
            for lab in BOX_LABELS:
                if n == _norm_text(lab) and lab not in box_of:
                    box_of[lab], box_text_of[lab] = b, t
    boxes_ordered_ok = (len(boxes) == 10 and len(box_of) == 10
                        and all(box_of.get(lab) is boxes[i] for i, lab in enumerate(BOX_LABELS)))

    # 顶部标题条：贴图形上沿、横贯图宽的实底/渐变横条
    band_cands = [r for r in recs
                  if r['fill']['kind'] in ('solid', 'grad') and r['fill']['alpha'] > 0.5
                  and fig_w and r['w'] >= 0.70 * fig_w
                  and fig_h and r['h'] <= 0.30 * fig_h
                  and r['y1'] <= fig[1] + 0.12 * fig_h]
    band = min(band_cands, key=lambda r: (r['y1'], -r['w'])) if band_cands else None
    title_rec = None
    if band is not None:
        if _norm_text(band['text']) == _norm_text(TITLE_TEXT):
            title_rec = band
        else:
            title_rec = next((t for t in _texts_in(_rect_of(band), 0.35)
                              if _norm_text(t['text']) == _norm_text(TITLE_TEXT)), None)
    if title_rec is None:
        title_rec = _find_text_rec(TITLE_TEXT)

    # 底部说明区：4组主标题/副标题的文字承载体
    bottom_titles = [_find_para(t) for t in BOTTOM_TITLES]
    bottom_subs = [_find_para(t) for t in BOTTOM_SUBS]
    bottom_y0 = fig[1] + 0.60 * fig_h

    # 4个辅助标签 + 绿色回用标签
    tag_blocks = {}      # 标签文字 -> (rect, 底板形状列表, 文字承载体)
    for _bx, _tag in TAG_SPECS + [(REUSE_BOX, GREEN_TAG_TEXT)]:
        _t = _find_text_rec(_tag)
        _rect, _fills = _label_block(_t)
        tag_blocks[_tag] = (_rect, _fills, _t)

    def _tag_fill_colors(tag):
        """标签底色候选（底板形状的实底色 + 文字承载体自身的实底色）。"""
        _r, _fs, _tr = tag_blocks.get(tag, (None, [], None))
        cols = [f['fill']['color'] for f in _fs if f['fill']['kind'] == 'solid']
        if _tr is not None and _tr['fill']['kind'] == 'solid':
            cols.append(_tr['fill']['color'])
        return [c for c in cols if c]

    def _mark_color(rec):
        """图元“可见主色”：有实底取填充色，否则取描边色（箭头/线条按描边取色）。"""
        if rec['fill']['kind'] == 'solid' and rec['fill']['color'] and rec['fill']['alpha'] > 0.5:
            return rec['fill']['color']
        return rec['line']['color']

    # ========== 维度2：完成度逐项判定 ==========
    # ---------- 01 幻灯片页面采用标准演示尺寸 ----------
    # 严格对齐细则，仅踩细则明确写出的点：
    #   ① 页面尺寸须与常规演示尺寸之一吻合（4:3、16:9 宽屏/全屏、A4 等，容差1cm）；
    #   ② 由①自然排除“把页面缩成与图形等大的异形版面”。
    # 与办公软件对齐：读 p:presentation/p:sldSz 的 @cx/@cy（即“设计→幻灯片大小”）。
    ok01 = any(abs(slide_w - a) <= 1.0 and abs(slide_h - b) <= 1.0
               for a, b in STANDARD_SLIDE_SIZES)
    results.append(("+3：" if ok01 else "+0：")
                   + "幻灯片页面采用标准演示尺寸（4:3 约25.4×19.05cm、16:9 约33.87×19.05cm 等常规比例），"
                     "未将页面尺寸缩成与图形等大的异形版面")

    # ---------- 02 仅含1页幻灯片 ----------
    # 严格对齐细则：① 幻灯片总页数为1；② 无多余空白页。
    # 与办公软件对齐：读 ppt/presentation.xml 的 p:sldIdLst 条目数（大纲/幻灯片窗格页数）。
    ok02 = (n_slides == 1)
    results.append(("+1：" if ok02 else "+0：") + "文件仅含1页幻灯片，无多余空白页")

    # ---------- 03 整幅流程图整体高度7cm且未变形 ----------
    # 严格对齐细则，仅踩细则明确写出的点：
    #   ① 整体高度实测落在 6.8-7.2cm；② 整体宽度落在 11-14cm（横纵比未被拉伸）。
    # 与办公软件对齐：整体外框取“唯一顶层组合”的 a:ext（选中后“大小”面板的高/宽），
    # 未成组时退化为全部图元外接矩形的并集。
    ok03 = (6.8 <= fig_h <= 7.2 and 11.0 <= fig_w <= 14.0)
    results.append(("+5：" if ok03 else "+0：")
                   + "整幅流程图为一个高度7cm的整体对象（实测整体高度在6.8-7.2cm之间），"
                     "且横纵比未被拉伸变形（宽度约11-14cm）")

    # ---------- 04 图形居中且四周留白 ----------
    # 严格对齐细则，仅踩细则明确写出的点：
    #   ① 图形中心与页面中心的横、纵向偏差均 ≤1.5cm；
    #   ② 上下左右四侧都留有可见空白页边（各侧留白 ≥0.5cm）。
    # 与办公软件对齐：由 a:off/a:ext 换算出的绝对位置与 p:sldSz 比较（“对齐→水平/垂直居中”后的结果）。
    dx = abs((fig[0] + fig[2]) / 2.0 - slide_w / 2.0)
    dy = abs((fig[1] + fig[3]) / 2.0 - slide_h / 2.0)
    margins = [fig[0], slide_w - fig[2], fig[1], slide_h - fig[3]]
    ok04 = (dx <= 1.5 and dy <= 1.5 and min(margins) >= 0.5)
    results.append(("+5：" if ok04 else "+0：")
                   + "图形整体在页面中水平且垂直居中（图形中心与页面中心的横、纵向偏差均≤1.5cm），"
                     "且四周留有可见空白页边（上下左右任一侧留白≥0.5cm）")

    # ---------- 05 所有描边线宽统一0.75磅 ----------
    # 严格对齐细则，仅踩细则明确写出的点：
    #   ① 凡“看得见描边”的图元（框线/箭头/引出线/分隔线/外框）线宽都等于0.75磅=9525EMU；
    #   ② 描边图元数量足够（≥10），避免只画了个别线条就算“统一”。
    # 与办公软件对齐：读 a:ln@w（“形状轮廓→粗细”），a:noFill 的轮廓视为无描边不参与判定。
    strokes = [r for r in recs if r['line']['has_el'] and r['line']['visible']]
    ok05 = (len(strokes) >= 10 and all(s['line']['w'] == LINE_W_075PT for s in strokes))
    results.append(("+3：" if ok05 else "+0：")
                   + "所有描边线条（流程框边框、连接箭头、引出线、分隔线、外层边框）线宽统一为0.75磅")

    # ---------- 06 中文统一宋体 ----------
    # 严格对齐细则，仅踩细则明确写出的点：
    #   ① 逐个 run 解析“含中文的文字”的有效东亚字体；
    #   ② 有效东亚字体可以来自 run 自身，也可以由主题东亚字体统一指定（+mj-ea/+mn-ea 或缺省继承）；
    #   ③ 有效东亚字体须为宋体（SimSun/NSimSun 同族），指定为 Times New Roman/Arial/微软雅黑等不得分；
    #   ④ 不含中文的符号/数字 run（带圈序号、PAC/PAM 等）不参与判定。
    # 与办公软件对齐：读 a:rPr/a:ea@typeface 与主题 a:fontScheme 的 majorFont/minorFont 的 a:ea。
    def _eff_ea(run):
        ea = run['ea']
        if not ea:
            return minor_ea
        if ea.startswith('+'):
            return major_ea if 'mj' in ea else minor_ea
        return ea
    cjk_runs = [x for r in text_recs for x in _runs(r) if _cjk(x['t'])]
    ok06 = bool(cjk_runs) and all(
        re.search(r'宋体|SimSun|NSimSun', _eff_ea(x) or '', re.I) for x in cjk_runs)
    results.append(("+5：" if ok06 else "+0：")
                   + "中文文字统一使用宋体——各文本的东亚字体设为宋体（可由主题东亚字体统一指定为宋体），"
                     "标题、流程框文字、辅助标签、底部说明文字均呈宋体字形；若中文被指定为 Times New Roman、"
                     "Arial、微软雅黑等非宋体字体则不得分（带圈序号等符号字形不计入）")

    # ---------- 07 全图组合为一个组合对象 ----------
    # 严格对齐细则，仅踩细则明确写出的点：
    #   ① 页面上的顶层图元只有1个；② 该顶层图元是组合（p:grpSp）；
    #   ③ 组合内含足够多的成员（≥20），即“全图各元素”都在这一个组合里。
    # 与办公软件对齐：p:spTree 的顶层子节点与 p:grpSp（“选中即可整体移动/按高度缩放”）。
    ok07 = (tops['n'] == 1 and tops['groups'] == 1 and tops['in_group'] >= 20)
    results.append(("+3：" if ok07 else "+0：")
                   + "全图各元素组合为一个可整体移动/缩放的组合对象，便于按7cm高度整体调整而不散乱")

    # ---------- 08 顶部深青绿标题条 + 居中白色加粗标题 ----------
    # 严格对齐细则，仅踩细则明确写出的点：
    #   ① 图形顶部有横贯图宽的色条（宽度 ≥0.75 图宽）且底色为深青绿（渐变时任一色标为深青绿）；
    #   ② 条内标题文字与指定标题完全一致（无错漏）；③ 文字加粗；④ 文字白色；⑤ 居中。
    # 与办公软件对齐：a:solidFill/a:gradFill 取色、a:rPr@b、a:rPr 的 a:solidFill、a:pPr@algn。
    ok08 = False
    if band is not None and title_rec is not None:
        _bf = band['fill']
        _bcols = _bf['colors'] if _bf['kind'] == 'grad' else ([_bf['color']] if _bf['color'] else [])
        _band_ok = bool(_bcols) and any(_is_dark_teal(c) for c in _bcols) and band['w'] >= 0.75 * fig_w
        _truns = [x for x in _runs(title_rec) if x['t'].strip()]
        _title_ok = (_norm_text(title_rec['text']) == _norm_text(TITLE_TEXT)
                     and bool(_truns)
                     and all(x['b'] for x in _truns)
                     and all(_is_whiteish(x['color']) for x in _truns)
                     and (title_rec['algn'] == 'ctr'
                          or abs(title_rec['cxc'] - band['cxc']) <= 0.05 * max(fig_w, 1e-6)))
        ok08 = bool(_band_ok and _title_ok)
    results.append(("+3：" if ok08 else "+0：")
                   + "顶部有横贯图宽的深青绿色标题条，内含居中、白色、加粗的标题"
                     "\"工业园中水回用处理流程与水质保障路径\"，文字无错漏")

    # ---------- 09 标题字号形成主次层级 ----------
    # 严格对齐细则，仅踩细则明确写出的点：
    #   ① 参照字号取“流程框文字 + 底部说明文字”的最大字号；
    #   ② 标题字号 ≥ 参照字号的1.5倍。
    # 与办公软件对齐：读 a:rPr@sz（百分之一磅，除以100即“字号”框中的磅值）。
    _ref_szs = [x['sz'] for t in box_text_of.values() for x in _runs(t) if x['sz']]
    for _rc, _rp in bottom_titles + bottom_subs:
        if _rc is not None:
            _ref_szs += [x['sz'] for x in _rp if x['sz']]
    _tsz = max([x['sz'] for x in _runs(title_rec) if x['sz']] or [0]) if title_rec else 0
    ok09 = bool(_tsz and _ref_szs) and _tsz >= 1.5 * max(_ref_szs)
    results.append(("+3：" if ok09 else "+0：")
                   + "标题字号明显大于流程框与底部说明文字（≥其1.5倍），在整幅图中形成明确的主标题层级")

    # ---------- 10 标题条横向渐变 ----------
    # 严格对齐细则，仅踩细则明确写出的点：
    #   ① 标题条填充方式为渐变；② 至少2个不同的青绿/绿色色标（非单一纯色平铺）。
    # 与办公软件对齐：读 a:gradFill/a:gsLst 的各 a:gs 颜色（“形状填充→渐变填充”）。
    ok10 = False
    if band is not None and band['fill']['kind'] == 'grad':
        _gcols = {c for c in band['fill']['colors'] if _is_teal(c) or _is_green(c)}
        ok10 = len(_gcols) >= 2
    results.append(("+1：" if ok10 else "+0：")
                   + "标题条底色为青绿到深绿的横向渐变，非单一纯色平铺")

    # ---------- 11 双层圆角框结构 ----------
    # 严格对齐细则，仅踩细则明确写出的点：
    #   ① 存在覆盖整幅图（宽高均 ≥95%）的圆角矩形外框；
    #   ② 其内部另有一个圆角矩形面板（浅色底或无填充）承载流程区，面积占比 ≥5% 且未与外框等大。
    # 与办公软件对齐：读 a:prstGeom@prst=roundRect 系列与 a:off/a:ext 的包含关系。
    _outer = [r for r in recs if r['geom'].startswith('round')
              and r['w'] >= 0.95 * fig_w and r['h'] >= 0.95 * fig_h]
    _inner = []
    if _outer:
        _o0 = _outer[0]
        for r in recs:
            if r is _o0 or not r['geom'].startswith('round') or not _inside(r, _rect_of(_o0)):
                continue
            _ar = r['w'] * r['h']
            if not (0.05 * fig_area <= _ar <= 0.99 * fig_area) or r['w'] < 0.50 * fig_w:
                continue
            _f = r['fill']
            if (_f['kind'] == 'none' or _f['alpha'] <= 0.1
                    or (_f['kind'] in ('solid', 'grad') and _f['color']
                        and _hex_to_hsv(_f['color'])[2] >= 0.85)):
                _inner.append(r)
    ok11 = bool(_outer) and bool(_inner)
    results.append(("+3：" if ok11 else "+0：")
                   + "图形外层为一个圆角矩形边框，其内部另有一个浅色底圆角面板承载流程区，形成双层圆角框结构")

    # ---------- 12 10个流程框文字齐全准确 ----------
    # 严格对齐细则，仅踩细则明确写出的点：
    #   ① 主流程框恰好10个（白底 + 青绿描边 + 竖长条形，按占图比例筛选，排除内外层面板）；
    #   ② 自左至右的文字与指定顺序逐一对应（无错字、无漏项、无多余项）；
    #   ③ 文字可写在框自身，也可由居中放置于框内的独立文本框承载。
    # 与办公软件对齐：a:prstGeom、a:solidFill、a:ln 取色 + a:off/a:ext 排序 + a:t 文本比对。
    ok12 = boxes_ordered_ok
    results.append(("+5：" if ok12 else "+0：")
                   + "主流程为10个横向排列的白底圆角矩形框，自左至右文字依次为\"厂区综合废水收集\""
                     "\"粗格栅与调节井\"\"pH均衡池\"\"混凝沉淀池\"\"多介质过滤器\"\"活性炭吸附塔\""
                     "\"超滤膜组件\"\"消毒接触池\"\"回用清水池\"\"冷却补水系统\"，无错字、无漏项、无多余项")

    # ---------- 13 10个流程框等宽等高、顶端对齐、间距均匀 ----------
    # 严格对齐细则，仅踩细则明确写出的点：
    #   ① 宽度、高度极差 ≤0.05cm（等宽等高）；② 顶边坐标极差 ≤0.05cm（顶端对齐）；
    #   ③ 相邻框水平间距互差 ≤0.1cm（间距均匀）。
    # 与办公软件对齐：读各框 a:off@x/@y 与 a:ext@cx/@cy（“大小和位置”面板数值）。
    ok13 = False
    if len(boxes) == 10:
        _ws = [b['w'] for b in boxes]
        _hs = [b['h'] for b in boxes]
        _ys = [b['y1'] for b in boxes]
        _gaps = [boxes[i + 1]['x1'] - boxes[i]['x2'] for i in range(9)]
        ok13 = (max(_ws) - min(_ws) <= 0.05 and max(_hs) - min(_hs) <= 0.05
                and max(_ys) - min(_ys) <= 0.05 and max(_gaps) - min(_gaps) <= 0.1)
    results.append(("+3：" if ok13 else "+0：")
                   + "10个流程框等宽等高、顶端对齐、水平间距均匀（相邻框间距互差≤0.1cm）")

    # ---------- 14 流程框内为真正的竖排文字 ----------
    # 严格对齐细则，仅踩细则明确写出的点：
    #   ① 10个流程框的文字承载体都设置了竖排文字方向；
    #   ② “每字单独一个段落/手动换行”硬拆模拟竖排（无竖排方向设置）不得分。
    # 与办公软件对齐：读 a:bodyPr@vert（“文字方向→竖排/所有文字旋转”）。
    ok14 = (len(box_text_of) == 10
            and all(t['vert'] in VERT_MODES for t in box_text_of.values()))
    results.append(("+5：" if ok14 else "+0：")
                   + "流程框内文字为真正的竖排文字（通过文本框竖排方向设置实现，自上而下排列，"
                     "改写文字后仍能自动竖排）；以每字单独一个段落或手动换行硬拆的方式模拟竖排不得分")

    # ---------- 15 10个圆形序号标记1-10 ----------
    # 严格对齐细则，仅踩细则明确写出的点：
    #   ① 每个流程框左上部区域内有且仅有一个序号标记（阿拉伯数字或带圈数字符号）；
    #   ② 序号值与该框的左右次序一一对应（不错位、不重号，每个标记只认领一次）；
    #   ③ 呈“圆形”：带圈数字字形本身，或青绿实底的正圆/椭圆衬底；
    #   ④ 配色为青绿圆底白字，或青绿色数字。
    # 与办公软件对齐：a:t 文本 + a:prstGeom@prst=ellipse + a:solidFill + a:rPr 的 a:solidFill。
    ok15 = False
    if len(boxes) == 10 and len(box_of) == 10:
        _hit15, _used15 = 0, set()
        for _i, _lab in enumerate(BOX_LABELS):
            _b = box_of[_lab]
            _zone = (_b['x1'] - 0.10, _b['y1'] - 0.15, _b['x2'] + 0.10, _b['y1'] + 0.45 * _b['h'])
            _cand = []
            for t in text_recs:
                if id(t) in _used15 or _seq_value(t['text']) != _i + 1 or not _inside(t, _zone):
                    continue
                _rs = [x for x in _runs(t) if x['t'].strip()]
                _col = _rs[0]['color'] if _rs else None
                _circ = _is_circled_glyph(t['text'].strip())
                _disk = t if (t['geom'] in ('ellipse', 'circle')
                              and t['fill']['kind'] == 'solid'
                              and _is_teal(t['fill']['color'])) else None
                if _disk is None:
                    _disk = next((o for o in recs
                                  if o['geom'] in ('ellipse', 'circle')
                                  and o['fill']['kind'] == 'solid' and _is_teal(o['fill']['color'])
                                  and _inter_area(_rect_of(o), _rect_of(t)) > 0), None)
                if ((_circ and _is_teal(_col))
                        or (_disk is not None and (_is_whiteish(_col) or _is_teal(_col)))):
                    _cand.append(t)
            if len(_cand) == 1:
                _used15.add(id(_cand[0]))
                _hit15 += 1
        ok15 = (_hit15 == 10)
    results.append(("+5：" if ok15 else "+0：")
                   + "10个流程框左上部各有一个圆形序号标记，依次为1至10（阿拉伯数字或带圈数字符号），"
                     "青绿色圆底配白色数字（或青绿色数字），与所属框一一对应且不错位、不重号")

    # ---------- 16 9个深蓝色向右箭头串联 ----------
    # 严格对齐细则，仅踩细则明确写出的点：
    #   ① 箭头为深蓝色（箭头形状取填充色、线条取描边色）；
    #   ② 指向一致向右（同时含左向端点的双向线条不计）；
    #   ③ 位于框体中部高度；④ 9个相邻框间隙各有一个，串联全部10个框。
    # 与办公软件对齐：a:prstGeom@prst=rightArrow 或 a:ln 的 a:headEnd/a:tailEnd 箭头端点（含 flipH）。
    arrows = []
    if len(boxes) == 10:
        _mlo = min(b['y1'] for b in boxes) + 0.10 * max(b['h'] for b in boxes)
        _mhi = max(b['y2'] for b in boxes) - 0.10 * max(b['h'] for b in boxes)
        for r in recs:
            _d = _arrow_dirs(r)
            if 'r' in _d and 'l' not in _d and _mlo <= r['cyc'] <= _mhi \
                    and _is_dark_blue(_mark_color(r)):
                arrows.append(r)
    _gap_hit = sum(1 for i in range(9)
                   if any(boxes[i]['x2'] - 0.25 <= a['cxc'] <= boxes[i + 1]['x1'] + 0.25
                          for a in arrows)) if len(boxes) == 10 else 0
    ok16 = (len(arrows) >= 9 and _gap_hit == 9)
    results.append(("+3：" if ok16 else "+0：")
                   + "相邻流程框之间有9个深蓝色向右箭头，依次串联全部10个框，位于框体中部高度、方向一致")

    # ---------- 17 4条引出线连接4个浅绿底标签 ----------
    # 严格对齐细则，仅踩细则明确写出的点：
    #   ① 4个标签文字分别正确；② 标签为浅绿底圆角形状；
    #   ③ 标签横向位置与所属流程框对应（水平投影有交集）；
    #   ④ 该框下方有一条引出线，上端接框底、下端接标签顶、横向位于框的左右范围内。
    # 与办公软件对齐：a:t 文本 + a:prstGeom + a:solidFill 取色 + 线条 a:off/a:ext 的首尾坐标。
    leaders = {}
    for _bx, _tag in TAG_SPECS:
        _b = box_of.get(_bx)
        _rect, _fills, _trec = tag_blocks.get(_tag, (None, [], None))
        if _b is None or _rect is None or _trec is None:
            continue
        if not any(_is_light_green(c) for c in _tag_fill_colors(_tag)):
            continue
        if not (any(f['geom'].startswith('round') for f in _fills)
                or _trec['geom'].startswith('round')):
            continue
        if _x_overlap(_rect, _rect_of(_b)) <= 0:
            continue
        _ln = next((r for r in recs
                    if r['w'] <= 0.15 and r['h'] >= 0.05
                    and (r['line']['visible'] or r['geom'] in ('line', 'straightConnector1'))
                    and _b['x1'] - 0.10 <= r['cxc'] <= _b['x2'] + 0.10
                    and abs(r['y1'] - _b['y2']) <= 0.60
                    and abs(r['y2'] - _rect[1]) <= 0.60), None)
        if _ln is not None:
            leaders[_tag] = _ln
    ok17 = (len(leaders) == 4)
    results.append(("+5：" if ok17 else "+0：")
                   + "\"pH均衡池\"\"混凝沉淀池\"\"超滤膜组件\"\"消毒接触池\"四个框下方各有一条引出线连接到"
                     "一个浅绿底圆角标签，四个标签文字分别为\"酸碱调节\"\"PAC/PAM投加\"\"核心净化单元\""
                     "\"次氯酸钠投加\"，标签横向位置与所属流程框对应")

    # ---------- 18 4条引出线为虚线 ----------
    # 严格对齐细则，仅踩细则明确写出的点：
    #   ① 上述4条引出线都设置了虚线线型（短划线/点划线等）；② 实线或未设置线型不得分。
    # 与办公软件对齐：读 a:ln/a:prstDash@val（“形状轮廓→虚线”）。
    ok18 = (len(leaders) == 4
            and all(l['line']['dash'] not in (None, 'solid') for l in leaders.values()))
    results.append(("+3：" if ok18 else "+0：")
                   + "上述4条引出线为虚线样式（短划线/点划线），非实线")

    # ---------- 19 回用清水池下向下箭头接绿色标签 ----------
    # 严格对齐细则，仅踩细则明确写出的点：
    #   ① 存在绿色圆角标签\"绿化与冲洗回用\"；
    #   ② 该框下方有一条向下的箭头，横向位于框的左右范围内、上端起于框底、下端接到标签顶。
    # 与办公软件对齐：a:prstGeom@prst=downArrow 或 a:tailEnd 箭头端点 + a:off/a:ext 首尾坐标。
    ok19 = False
    _b9 = box_of.get(REUSE_BOX)
    _grect, _gfills, _gtrec = tag_blocks.get(GREEN_TAG_TEXT, (None, [], None))
    if _b9 is not None and _grect is not None and _gtrec is not None:
        _green_ok = any(_is_green(c) for c in _tag_fill_colors(GREEN_TAG_TEXT))
        _da = next((r for r in recs if 'd' in _arrow_dirs(r)
                    and _b9['x1'] - 0.30 <= r['cxc'] <= _b9['x2'] + 0.30
                    and r['y1'] >= _b9['y2'] - 0.60
                    and abs(r['y2'] - _grect[1]) <= 0.80), None)
        ok19 = bool(_green_ok and _da is not None)
    results.append(("+3：" if ok19 else "+0：")
                   + "\"回用清水池\"框下方有一条向下的箭头连接到右下角的绿色圆角标签\"绿化与冲洗回用\"")

    # ---------- 20 完整“异常回流”回路且各段可正常渲染 ----------
    # 严格对齐细则，仅踩细则明确写出的点：
    #   ① 任一图元出现宽/高为负值（无法渲染的线段）即不得分；
    #   ② 存在横贯图面的深蓝色回流横线（宽 ≥50% 图宽），位于流程框下方；
    #   ③ 横线右端伸到右下角回用处；④ 横线左端有一段向上的箭头段，
    #      其下端与横线相接、上端指回\"粗格栅与调节井\"框底（首尾相接、走向正确）。
    # 与办公软件对齐：a:ext@cx/@cy（ST_PositiveCoordinate 不允许负值）+ 各段 a:off/a:ext 端点坐标。
    ok20 = False
    if not negs and len(boxes) == 10:
        _bbot = max(b['y2'] for b in boxes)
        _b2 = box_of.get(BACKFLOW_TO_BOX)
        _hcands = [r for r in recs
                   if r['w'] >= 0.50 * fig_w and r['h'] <= 0.10
                   and r['cyc'] >= _bbot - 0.20 and _is_dark_blue(_mark_color(r))]
        for _hl in _hcands:
            if _b9 is not None and _hl['x2'] < _b9['x1'] - 1.0:
                continue
            _up = next((r for r in recs if 'u' in _arrow_dirs(r) and r['w'] <= 0.30
                        and abs(r['cxc'] - _hl['x1']) <= 0.50
                        and abs(r['y2'] - _hl['cyc']) <= 0.50
                        and _b2 is not None and abs(r['y1'] - _b2['y2']) <= 1.00), None)
            if _up is not None:
                ok20 = True
                break
    results.append(("+5：" if ok20 else "+0：")
                   + "设有完整的\"异常回流\"回路：自右下角回用处向下、向左横贯图面、再向上并以箭头指回"
                     "\"粗格栅与调节井\"框，折线各段首尾相接、走向与箭头方向正确，各段均能正常显示"
                     "（不得出现宽/高为负值等无法渲染的线段，不得缺段或断开）")

    # ---------- 21 “异常回流”标签压在回流横线中部 ----------
    # 严格对齐细则，仅踩细则明确写出的点：
    #   ① 存在可正常渲染的回流横线（有负值线段时无从谈“压在线上”）；
    #   ② 标签为深蓝色实底小标签；③ 文字为白色；
    #   ④ 标签中心位于横线中部（与横线中点的横向偏差 ≤横线长度的10%）；
    #   ⑤ 标签垂直方向覆盖该横线（压在线上）。
    # 与办公软件对齐：a:solidFill 取色、a:rPr 的 a:solidFill、a:off/a:ext 的相对位置。
    ok21 = False
    if not negs and len(boxes) == 10:
        _hl2 = next((r for r in recs
                     if r['w'] >= 0.50 * fig_w and r['h'] <= 0.10
                     and r['cyc'] >= max(b['y2'] for b in boxes) - 0.20
                     and _is_dark_blue(_mark_color(r))), None)
        _brec = _find_text_rec(BACKFLOW_TEXT)
        if _hl2 is not None and _brec is not None:
            _brect, _bfills = _label_block(_brec)
            _pads = [f for f in _bfills
                     if f['fill']['kind'] == 'solid' and _is_dark_blue(f['fill']['color'])]
            _brs = [x for x in _runs(_brec) if x['t'].strip()]
            ok21 = bool(_pads and _brs
                        and all(_is_whiteish(x['color']) for x in _brs)
                        and abs((_brect[0] + _brect[2]) / 2.0 - _hl2['cxc']) <= 0.10 * _hl2['w']
                        and _brect[1] - 0.10 <= _hl2['cyc'] <= _brect[3] + 0.10)
    results.append(("+3：" if ok21 else "+0：")
                   + "\"异常回流\"文字位于回流横线中部，为深蓝色实底小标签配白色文字，压在线上且不被线条穿断")

    # ---------- 22 底部4组说明项文字齐全准确 ----------
    # 严格对齐细则，仅踩细则明确写出的点：
    #   ① 4条主标题文字齐全准确；② 4条副标题文字齐全准确（比对时忽略空白与大小写差异）。
    # 与办公软件对齐：读全部 a:t 文本（即幻灯片上可见的文字内容）。
    ok22 = all(_norm_text(t) in all_text_norm for t in BOTTOM_TITLES + BOTTOM_SUBS)
    results.append(("+5：" if ok22 else "+0：")
                   + "底部设有4组说明项，文字齐全准确：主标题依次为\"全过程在线监测\"\"水质达标保障\""
                     "\"设备定期维护\"\"数据记录追溯\"，对应副标题依次为\"pH/浊度/余氯/电导率\""
                     "\"满足回用水质标准\"\"保障系统稳定运行\"\"管理可视化、可追溯\"")

    # ---------- 23 一条长条承载 + 3条竖虚线分隔 ----------
    # 严格对齐细则，仅踩细则明确写出的点：
    #   ① 底部区域有一条横贯图宽（≥75% 图宽）的浅色长条统一承载；
    #   ② 长条范围内有3条竖向虚线分隔；
    #   ③ 不是拆成4个各自带边框的独立小块（≥4个带框小块承载说明项则不得分）。
    # 与办公软件对齐：a:off/a:ext + a:solidFill + a:ln/a:prstDash（分隔线的“虚线”线型）。
    _bar = next((r for r in recs
                 if r['w'] >= 0.75 * fig_w and r['h'] <= 0.25 * fig_h
                 and r['cyc'] >= bottom_y0
                 and r['fill']['kind'] in ('solid', 'grad') and r['fill']['alpha'] > 0.5
                 and r['fill']['color'] and _hex_to_hsv(r['fill']['color'])[2] >= 0.85), None)
    _divs = []
    if _bar is not None:
        _divs = [r for r in recs
                 if r is not _bar and r['w'] <= 0.15 and r['h'] >= 0.10
                 and _inside(r, (_bar['x1'], _bar['y1'] - 0.20, _bar['x2'], _bar['y2'] + 0.20))
                 and r['line']['dash'] not in (None, 'solid')]
    _split = [r for r in recs
              if r['w'] <= 0.35 * fig_w and r['cyc'] >= bottom_y0
              and r['line']['visible'] and r['fill']['kind'] == 'solid' and r['fill']['alpha'] > 0.5
              and any(_norm_text(t) in _norm_text(r['text']) for t in BOTTOM_TITLES)]
    ok23 = (_bar is not None and len(_divs) >= 3 and len(_split) < 4)
    results.append(("+3：" if ok23 else "+0：")
                   + "4组说明项由一条横贯图宽的浅色圆角长条统一承载、等距分布，项间以3条竖向虚线分隔，"
                     "而非拆成4个各自带框的独立小块")

    # ---------- 24 说明项主标题加粗且字号大于副标题 ----------
    # 严格对齐细则，仅踩细则明确写出的点：
    #   ① 4组的主标题文字都加粗；② 每组主标题字号 > 其副标题字号。
    # 与办公软件对齐：读 a:rPr@b 与 a:rPr@sz（“加粗”按钮与“字号”框）。
    ok24 = False
    if all(_r is not None for _r, _p in bottom_titles) and all(_r is not None for _r, _p in bottom_subs):
        _n24 = 0
        for (_tr, _tp), (_sr, _sp) in zip(bottom_titles, bottom_subs):
            _ts = max([x['sz'] for x in _tp if x['sz']] or [0])
            _ss = max([x['sz'] for x in _sp if x['sz']] or [0])
            if _tp and all(x['b'] for x in _tp if x['t'].strip()) and _ts > _ss > 0:
                _n24 += 1
        ok24 = (_n24 == 4)
    results.append(("+1：" if ok24 else "+0：")
                   + "每组说明项主标题加粗且字号大于其下方副标题，形成主次层级")

    # ---------- 25 全图9个可编辑小图标齐备 ----------
    # 严格对齐细则，仅踩细则明确写出的点：
    #   ① 小图标候选 = 无文字的图片，或无文字的非结构性小自动图形（尺寸 ≤1cm）；
    #   ② 4个辅助标签各1个、\"绿化与冲洗回用\"标签1个、底部4组说明项各1个，合计9个到位。
    # 与办公软件对齐：p:pic（插入的图标/图片）与 a:prstGeom 非矩形/线条类的小形状 + 位置包含关系。
    icon_cands = [r for r in recs
                  if not r['text'].strip()
                  and (r['kind'] == 'pic'
                       or (r['kind'] == 'sp' and r['geom'] and r['geom'] not in STRUCT_GEOMS))
                  and 0.05 <= r['w'] <= 1.0 and 0.05 <= r['h'] <= 1.0]
    _placed = 0
    for _bx, _tag in TAG_SPECS + [(REUSE_BOX, GREEN_TAG_TEXT)]:
        _rect = tag_blocks.get(_tag, (None, [], None))[0]
        if _rect is not None and any(_inside(ic, _rect, 0.30) for ic in icon_cands):
            _placed += 1
    for _tr, _tp in bottom_titles:
        if _tr is None:
            continue
        if any(_tr['x1'] - 1.5 <= ic['cxc'] <= _tr['x1'] + 0.5
               and _tr['y1'] - 0.5 <= ic['cyc'] <= _tr['y1'] + 1.0 for ic in icon_cands):
            _placed += 1
    ok25 = (len(icon_cands) >= 9 and _placed == 9)
    results.append(("+5：" if ok25 else "+0：")
                   + "全图小图标齐备共9个——4个辅助标签、\"绿化与冲洗回用\"标签、底部4组说明项各带一个"
                     "可编辑小图标（试剂瓶、水滴、加号、叶片、盾牌、清单、齿轮、文档等类）")

    # ---------- 26 文字完整显示、无溢出无压盖 ----------
    # 严格对齐细则，仅踩细则明确写出的点：
    #   ① 逐个文本按字号、字数、内边距估算所需排版尺寸，超出可用区域即判溢出/断词；
    #   ② 相邻文字图元互相压盖（重叠面积 > 较小者面积的50%）即判遮挡；
    #   ③ 序号标记压在圆底上属正常设计，不计入压盖。
    # 与办公软件对齐：a:ext 的框体尺寸、a:bodyPr 的 lIns/tIns/rIns/bIns 内边距、a:rPr@sz 字号、
    # a:bodyPr@vert 竖排方向与 a:p 段落数（即办公软件的自动换行与溢出表现）。
    _over = [r for r in text_recs if _text_overflow(r)]
    _multi = [r for r in text_recs
              if len(re.sub(r'\s', '', r['text'])) >= 2
              and not (_seq_value(r['text']) is not None
                       and len(re.sub(r'\s', '', r['text'])) <= 2)]
    _clash = []
    for _i in range(len(_multi)):
        for _j in range(_i + 1, len(_multi)):
            _a, _b_ = _multi[_i], _multi[_j]
            _ia = _inter_area(_rect_of(_a), _rect_of(_b_))
            _sm = min(_a['w'] * _a['h'], _b_['w'] * _b_['h'])
            if _sm > 0 and _ia > 0.5 * _sm:
                _clash.append((_a['id'], _b_['id']))
    ok26 = (not _over) and (not _clash)
    results.append(("+5：" if ok26 else "+0：")
                   + "全部文字完整显示在所属形状/文本框内——无文字溢出框外、无被裁切、无因框体过小导致的"
                     "非预期换行断词，相邻图元之间无互相压盖遮挡")

    # ---------- 27 青绿加深蓝双色系配色统一 ----------
    # 严格对齐细则，仅踩细则明确写出的点：
    #   ① 流程框白底青绿描边；② 串联箭头深蓝；③ 回流横线深蓝；
    #   ④ 4个辅助标签浅绿底；⑤ 底部说明条浅蓝底；⑥ 流程框内文字加粗且为深色（深蓝或近黑）。
    # 与办公软件对齐：a:solidFill/a:ln 取色（含 a:schemeClr 按主题与母版 p:clrMap 解析）、a:rPr@b。
    _c1 = (len(boxes) == 10
           and all(_is_whiteish(b['fill']['color']) and _is_teal(b['line']['color']) for b in boxes))
    _c2 = (len(arrows) >= 9)
    _c3 = any(r['w'] >= 0.50 * fig_w and r['h'] <= 0.10 and _is_dark_blue(_mark_color(r))
              for r in recs)
    _c4 = all(any(_is_light_green(c) for c in _tag_fill_colors(_tag)) for _bx, _tag in TAG_SPECS)
    _c5 = any(r['cyc'] >= bottom_y0 and r['fill']['kind'] == 'solid' and r['fill']['alpha'] > 0.5
              and _is_light_blue(r['fill']['color']) for r in recs)
    _c6 = (len(box_text_of) == 10
           and all(any(x['b'] for x in _runs(t) if x['t'].strip())
                   and all(_is_darkish(x['color']) for x in _runs(t) if x['t'].strip())
                   for t in box_text_of.values()))
    ok27 = bool(_c1 and _c2 and _c3 and _c4 and _c5 and _c6)
    results.append(("+3：" if ok27 else "+0：")
                   + "整体配色统一为青绿加深蓝双色系：流程框白底青绿描边、串联箭头与回流线深蓝、"
                     "辅助标签浅绿底、底部说明条浅蓝底，流程框内文字加粗且为深色（深蓝或近黑）")

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
    # 定位并打开被评估的 .pptx。命令行第 1 个参数可覆盖目录路径；未指定时使用
    # 脚本自身所在目录。
    if len(sys.argv) >= 2:
        _dir = sys.argv[1]
    else:
        _dir = os.path.dirname(os.path.abspath(__file__))
    _out = evaluate(_dir)
    print(json.dumps(_out, ensure_ascii=False, indent=2))
