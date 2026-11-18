#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# pyright: reportUnknownMemberType=false, reportUnknownVariableType=false, reportUnknownArgumentType=false, reportUnknownParameterType=false, reportMissingParameterType=false, reportPrivateUsage=false, reportUnknownLambdaType=false, reportGeneralTypeIssues=false
import os
import re
import sys
import json
import math
import zipfile
import colorsys
import traceback

from pptx import Presentation

# 说明：按接口统一约定，本模块不修改全局 sys.stdout，也不 print 主结果。
# 主结果统一由 evaluate(dir_path) -> dict 返回；入参为“脚本所在目录的路径”，
# 由脚本自己在该目录内定位并打开被评估的 .pptx。
# 本文件仅在 __main__ 下作为本地调试入口，才把 dict 序列化为 JSON 打印。
# 返回字段：id / file_name / status / error / dim1_pass / dim1_reason /
#           dim2_items（每项只含 rule / max_delta / delta / hit）/
#           total_score / max_score
# status='error' 表示这份提交没被真正评过（目录不存在、目录内无 .pptx、脚本自身
# 异常），此时 dim1_pass 为空串（维度一未判定）、dim2_items 为空表、total_score
# 与 max_score 一律为 0；dim1_pass 为 False 时 dim1_reason 必须给出错误摘要，
# 通过时为空串。

SCRIPT_ID = "050"
MAX_SCORE_TOTAL = 127

_MAX_DELTAS_ORDER = [
    3,  # 01 第1页 浅蓝色圆角矩形 + “Macro Environment”
    3,  # 02 第1页 六段宏观环境标签文本
    5,  # 03 第1页 四条横贯平滑曲线
    1,  # 04 第1页 两个蓝色竖直向下箭头
    3,  # 05 第1页 浅绿色椭圆 + “Local Health Innovation Context”
    3,  # 06 第1页 上排三个小椭圆
    3,  # 07 第1页 下排三个小椭圆
    1,  # 08 第1页 两个绿色竖直向下箭头
    3,  # 09 第1页 圆角矩形 + 深蓝标题条“Innovation Dynamics Arena”
    5,  # 10 第1页 两个虚线边框矩形 + 竖直双向箭头
    3,  # 11 第1页 柱状图与人物两个拼合图形
    1,  # 12 第1页 一条竖直虚线
    3,  # 13 第1页 四角分布的四个椭圆
    5,  # 14 第1页 四个拼合小图标
    5,  # 15 第1页 四条实线双向箭头
    5,  # 16 第1页 斜体协商文本 + 四条点线箭头
    3,  # 17 第1页 “Digital health initiatives (T1)” + 水平箭头
    3,  # 18 第1页 “Digital health outcomes (T2)” + 水平箭头
    3,  # 19 第1页 竖直坐标轴箭头 + “System Conditions”
    1,  # 20 第1页 水平坐标轴箭头 + “Time (T)”
    3,  # 21 第2页 竖直坐标轴箭头 + 三段轴标签
    1,  # 22 第2页 水平坐标轴箭头 + “Time”
    5,  # 23 第2页 平滑上升曲线 + 三个圆环标记
    5,  # 24 第2页 三个横向色条 + 白色圆环序号
    1,  # 25 第2页 三段白色阶段标题
    1,  # 26 第2页 三段年份区间
    3,  # 27 第2页 三段“Representative milestones”+ 下横线
    5,  # 28 第2页 左列四行年份事项
    5,  # 29 第2页 中列五行年份事项
    5,  # 30 第2页 右列五行年份事项
    3,  # 31 第2页 三个卡片状圆角矩形
    5,  # 32 第2页 三个浅色圆角矩形 + 圆形底盘图标
    1,  # 33 第2页 三段“Stage profile”
    1,  # 34 第2页 三段阶段描述文本
    3,  # 35 第2页 两个圆形徽标
    3,  # 36 第2页 两条竖直虚线
    1,  # 37 幻灯片页面尺寸与备注页尺寸
    1,  # 38 无空白占位符框与空白文本框
    3,  # 39 字号均为 0.5 磅的整数倍
    5,  # 40 无负值宽高、不提示修复、无反向绘制
    5,  # 41 未启用缩放文字以适应形状且文字不溢出不出界
]
assert sum(_MAX_DELTAS_ORDER) == MAX_SCORE_TOTAL

# ========== 基础常量 ==========
_A = '{http://schemas.openxmlformats.org/drawingml/2006/main}'
_P = '{http://schemas.openxmlformats.org/presentationml/2006/main}'
_CM = 360000.0          # 1cm = 360000 EMU
_PAGE_W_CM = 33.87      # 期望页宽
_PAGE_H_CM = 19.05      # 期望页高
_NOTES_W_CM = 19.05     # 期望备注页宽
_NOTES_H_CM = 25.40     # 期望备注页高
# 符号字符集：箭头/数学/几何/杂项符号/装饰符号/私用区(Wingdings 类)/表情符号。
# 用于识别“以文字字符冒充图形”的做法（如用 ✚ 代替十字形状、用 ✓ 代替对勾形状）。
_GLYPH_RE = re.compile(
    '[←-⏿■-➿⬀-⯿-\U0001f300-\U0001faff]')


# ========== 颜色：主题解析与色系判定 ==========
def _theme_scheme(prs):
    """读取 ppt/theme/theme1.xml 的 <a:clrScheme>，并按 slideMaster 的 <p:clrMap> 建立别名。"""
    scheme = {}
    try:
        pkg = prs.part.package
        blob = None
        for part in pkg.iter_parts():
            if part.partname.endswith('theme1.xml'):
                blob = part.blob
                break
        if blob is None:
            return scheme
        txt = blob.decode('utf8', 'ignore')
        seg = re.search(r'<a:clrScheme.*?</a:clrScheme>', txt, re.S)
        if not seg:
            return scheme
        for m in re.finditer(r'<a:(\w+)>\s*<a:(srgbClr val="([0-9A-Fa-f]{6})"|sysClr[^>]*?lastClr="([0-9A-Fa-f]{6})")', seg.group(0)):
            name = m.group(1)
            val = m.group(3) or m.group(4)
            if val:
                scheme[name] = val.upper()
    except Exception:
        return scheme
    # clrMap 别名：bg1/tx1/bg2/tx2 → lt1/dk1/lt2/dk2
    for alias, target in (('bg1', 'lt1'), ('tx1', 'dk1'), ('bg2', 'lt2'), ('tx2', 'dk2')):
        if alias not in scheme and target in scheme:
            scheme[alias] = scheme[target]
    for alias, target in (('lt1', 'bg1'), ('dk1', 'tx1'), ('lt2', 'bg2'), ('dk2', 'tx2')):
        if alias not in scheme and target in scheme:
            scheme[alias] = scheme[target]
    return scheme


def _hex_to_rgb(val):
    try:
        v = str(val).strip().lstrip('#').upper()
        if len(v) != 6:
            return None
        return (int(v[0:2], 16), int(v[2:4], 16), int(v[4:6], 16))
    except Exception:
        return None


def _apply_transforms(rgb, el):
    """套用 <a:lumMod>/<a:lumOff>/<a:shade>/<a:tint>，与办公软件的“变亮/变暗”一致。"""
    if rgb is None or el is None:
        return rgb
    r, g, b = [c / 255.0 for c in rgb]
    def _val(tag):
        node = el.find(_A + tag)
        if node is None or node.get('val') is None:
            return None
        try:
            return float(node.get('val')) / 100000.0
        except Exception:
            return None
    sh, ti = _val('shade'), _val('tint')
    lm, lo = _val('lumMod'), _val('lumOff')
    if sh is not None:
        r, g, b = r * sh, g * sh, b * sh
    if ti is not None:
        r, g, b = r * ti + (1 - ti), g * ti + (1 - ti), b * ti + (1 - ti)
    if lm is not None or lo is not None:
        h, l, s = colorsys.rgb_to_hls(r, g, b)
        if lm is not None:
            l *= lm
        if lo is not None:
            l += lo
        l = max(0.0, min(1.0, l))
        r, g, b = colorsys.hls_to_rgb(h, l, s)
    return (int(round(max(0.0, min(1.0, r)) * 255)),
            int(round(max(0.0, min(1.0, g)) * 255)),
            int(round(max(0.0, min(1.0, b)) * 255)))


def _color_of(node, scheme):
    """从任意可含颜色的容器节点（solidFill/lnRef/fillRef 等）解析出 RGB。"""
    if node is None:
        return None
    srgb = node.find(_A + 'srgbClr')
    if srgb is not None:
        return _apply_transforms(_hex_to_rgb(srgb.get('val')), srgb)
    sysc = node.find(_A + 'sysClr')
    if sysc is not None:
        return _apply_transforms(_hex_to_rgb(sysc.get('lastClr') or 'FFFFFF'), sysc)
    sch = node.find(_A + 'schemeClr')
    if sch is not None:
        base = scheme.get(sch.get('val'))
        return _apply_transforms(_hex_to_rgb(base), sch) if base else None
    return None


def _lum(rgb):
    return 0.299 * rgb[0] + 0.587 * rgb[1] + 0.114 * rgb[2]


def _is_white(rgb):
    return rgb is not None and min(rgb) >= 235


def _is_black(rgb):
    return rgb is not None and max(rgb) <= 95 and (max(rgb) - min(rgb)) <= 30


def _is_blue(rgb):
    if rgb is None:
        return False
    r, g, b = rgb
    return b >= r + 20 and b >= g + 8


def _is_dark_blue(rgb):
    return _is_blue(rgb) and _lum(rgb) <= 170


def _is_light_blue(rgb):
    if rgb is None:
        return False
    r, g, b = rgb
    return b >= r and b >= g and min(rgb) >= 195 and (b - min(r, g)) >= 3


def _is_cyan(rgb):
    if rgb is None:
        return False
    r, g, b = rgb
    return g >= r + 25 and b >= r + 25 and abs(g - b) <= 95


def _is_orange(rgb):
    if rgb is None:
        return False
    r, g, b = rgb
    return r >= 175 and r >= g + 45 and g >= b + 25


def _is_green(rgb):
    if rgb is None:
        return False
    r, g, b = rgb
    return g >= r + 10 and g >= b + 10


def _is_dark_green(rgb):
    return _is_green(rgb) and _lum(rgb) <= 170


def _is_light_green(rgb):
    if rgb is None:
        return False
    r, g, b = rgb
    return g >= r and g >= b and min(rgb) >= 195 and (g - min(r, b)) >= 3


def _is_verylight(rgb):
    """白色或极浅色填充。"""
    return rgb is not None and min(rgb) >= 225


def _is_colored(rgb):
    """彩色（非黑白灰）。"""
    return rgb is not None and (max(rgb) - min(rgb)) >= 25


# ========== 形状：几何、填充、线条、文本 ==========
def _visual_box(l, t, w, h, rot):
    """绕形状中心旋转后的可视外接框；同时把负值宽高归一化为正值（与办公软件呈现一致）。"""
    cx, cy = l + w / 2.0, t + h / 2.0
    a = math.radians(rot or 0.0)
    nw = abs(w * math.cos(a)) + abs(h * math.sin(a))
    nh = abs(w * math.sin(a)) + abs(h * math.cos(a))
    return cx - nw / 2.0, cy - nh / 2.0, nw, nh


def _sp_pr(el):
    for tag in ('spPr', 'grpSpPr'):
        node = el.find(_P + tag)
        if node is not None:
            return node
    node = el.find('.//' + _P + 'spPr')
    if node is None:
        node = el.find('.//' + _A + 'spPr')
    return node


def _geom_of(el):
    node = el.find('.//' + _A + 'prstGeom')
    if node is not None:
        return node.get('prst') or '-'
    if el.find('.//' + _A + 'custGeom') is not None:
        return 'custGeom'
    return '-'


def _fill_of(el, scheme):
    """先读 spPr 的显式填充；无显式填充时回落到 <p:style><a:fillRef>（办公软件的形状样式）。"""
    pr = _sp_pr(el)
    if pr is not None:
        if pr.find(_A + 'noFill') is not None:
            return None
        sf = pr.find(_A + 'solidFill')
        if sf is not None:
            return _color_of(sf, scheme)
        for tag in ('gradFill', 'blipFill', 'pattFill'):
            if pr.find(_A + tag) is not None:
                return None
    st = el.find(_P + 'style')
    if st is not None:
        ref = st.find(_A + 'fillRef')
        if ref is not None and (ref.get('idx') or '0') != '0':
            return _color_of(ref, scheme)
    return None


def _ln_of(el):
    pr = _sp_pr(el)
    return pr.find(_A + 'ln') if pr is not None else None


def _line_color_of(el, scheme):
    ln = _ln_of(el)
    if ln is not None:
        if ln.find(_A + 'noFill') is not None:
            return None
        sf = ln.find(_A + 'solidFill')
        if sf is not None:
            return _color_of(sf, scheme)
    st = el.find(_P + 'style')
    if st is not None:
        ref = st.find(_A + 'lnRef')
        if ref is not None and (ref.get('idx') or '0') != '0':
            return _color_of(ref, scheme)
    return None


def _dash_class(el):
    """线型归类：solid 实线 / dash 虚线 / dashdot 点划线 / dot 点线。无 <a:ln> 视为实线。"""
    ln = _ln_of(el)
    if ln is None:
        return 'solid'
    node = ln.find(_A + 'prstDash')
    val = (node.get('val') or '') if node is not None else ''
    d = val.lower()
    if not d:
        return 'solid'
    if 'dashdot' in d:
        return 'dashdot'
    if 'dot' in d:
        return 'dot'
    if 'dash' in d:
        return 'dash'
    return 'solid'


def _ends_of(el):
    """(headEnd, tailEnd) 箭头类型；none/缺失记 None。"""
    ln = _ln_of(el)
    out = []
    for tag in ('headEnd', 'tailEnd'):
        node = ln.find(_A + tag) if ln is not None else None
        t = (node.get('type') if node is not None else None) or 'none'
        out.append(None if t == 'none' else t)
    return out[0], out[1]


def _default_sz_pt(prs):
    """<p:defaultTextStyle> lvl1 的默认字号，供未显式设置 sz 的 run 继承。"""
    try:
        txt = prs.part.blob.decode('utf8', 'ignore')
        m = re.search(r'<p:defaultTextStyle>.*?<a:lvl1pPr[^>]*>.*?<a:defRPr[^>]*?sz="(\d+)"', txt, re.S)
        if m:
            return int(m.group(1)) / 100.0
    except Exception:
        pass
    return 18.0


def _norm(s):
    """空白折叠 + 连接号统一，用于文本比对。"""
    t = re.sub(r'\s+', ' ', (s or '').replace(' ', ' ')).strip()
    return t.replace('–', '-').replace('—', '-').replace('−', '-')


def _font_ok(name):
    """Arial 或主题默认西文字体（未显式指定即为主题默认）。"""
    if name is None or str(name).strip() == '':
        return True
    n = str(name).strip().lower()
    return 'arial' in n or n.startswith('+mn') or n.startswith('+mj')


def _runs_of(sh, default_pt, scheme):
    out = []
    if not sh.has_text_frame:
        return out
    for para in sh.text_frame.paragraphs:
        for run in para.runs:
            rpr = run._r.find(_A + 'rPr')
            sz = None
            if rpr is not None and rpr.get('sz'):
                sz = int(rpr.get('sz')) / 100.0
            if sz is None:
                sz = default_pt
            col = None
            if rpr is not None:
                sf = rpr.find(_A + 'solidFill')
                if sf is not None:
                    col = _color_of(sf, scheme)
            out.append({
                'text': run.text or '',
                'sz': sz,
                'explicit_sz': (rpr is not None and rpr.get('sz') is not None),
                'b': (rpr is not None and rpr.get('b') == '1'),
                'i': (rpr is not None and rpr.get('i') == '1'),
                'color': col,
                'name': run.font.name,
            })
    return out


def _collect(slide, pw_cm, ph_cm, default_pt, scheme):
    """把一页上的顶层形状统一整理为可度量的记录：厘米尺寸 + 页面百分比位置 + 填充/线条/文本属性。"""
    recs = []
    for i, sh in enumerate(slide.shapes):
        el = sh._element
        try:
            l, t = float(sh.left), float(sh.top)
            w, h = float(sh.width), float(sh.height)
        except Exception:
            continue
        try:
            rot = float(sh.rotation or 0.0)
        except Exception:
            rot = 0.0
        vl, vt, vw, vh = _visual_box(l, t, w, h, rot)
        txt = ''
        if sh.has_text_frame:
            txt = _norm(' '.join(p.text for p in sh.text_frame.paragraphs))
        ln = _ln_of(el)
        lw = None
        if ln is not None and ln.get('w'):
            try:
                lw = int(ln.get('w')) / 12700.0
            except Exception:
                lw = None
        he, te = _ends_of(el)
        rec = {
            'i': i, 'sp': sh, 'el': el,
            'geom': _geom_of(el),
            'kind': str(getattr(sh, 'shape_type', '') or ''),
            'l': vl / _CM, 't': vt / _CM, 'w': vw / _CM, 'h': vh / _CM,
            'cx': (vl + vw / 2.0) / _CM, 'cy': (vt + vh / 2.0) / _CM,
            'xl': 100.0 * vl / (pw_cm * _CM), 'xr': 100.0 * (vl + vw) / (pw_cm * _CM),
            'yt': 100.0 * vt / (ph_cm * _CM), 'yb': 100.0 * (vt + vh) / (ph_cm * _CM),
            'rot': rot,
            'fill': _fill_of(el, scheme),
            'line': _line_color_of(el, scheme),
            'dash': _dash_class(el),
            'lw': lw, 'he': he, 'te': te,
            'text': txt, 'has_text': bool(txt),
            'runs': _runs_of(sh, default_pt, scheme),
            'is_pic': el.tag.endswith('}pic'),
        }
        rec['pcx'] = (rec['xl'] + rec['xr']) / 2.0
        rec['pcy'] = (rec['yt'] + rec['yb']) / 2.0
        recs.append(rec)
    return recs


def _in_win(r, x0, x1, y0=None, y1=None):
    """位置判定：以元素可视中心是否落在“距左侧/顶部页边线百分比”窗口内为准。"""
    if not (x0 - 0.01 <= r['pcx'] <= x1 + 0.01):
        return False
    if y0 is None:
        return True
    return y0 - 0.01 <= r['pcy'] <= y1 + 0.01


def _inside(r, x0, x1, y0, y1, tol=0.05):
    return (r['xl'] >= x0 - tol and r['xr'] <= x1 + tol
            and r['yt'] >= y0 - tol and r['yb'] <= y1 + tol)


def _btw(v, lo, hi):
    return v is not None and lo - 1e-9 <= v <= hi + 1e-9


def _text_recs(recs, target, x0=None, x1=None, y0=None, y1=None):
    """按内容（空白折叠、连接号统一、忽略大小写、忽略空格差异）找文本块，可附加位置窗口过滤。"""
    key = _norm(target).lower()
    tight = key.replace(' ', '')
    out = []
    for r in recs:
        if not r['has_text']:
            continue
        low = r['text'].lower()
        if key not in low and tight not in low.replace(' ', ''):
            continue
        if x0 is not None and not _in_win(r, x0, x1, y0, y1):
            continue
        out.append(r)
    return out


def _main_runs(r):
    """忽略纯空白 run，取承载文字的 run。"""
    return [x for x in r['runs'] if (x['text'] or '').strip()]


def _run_stat(r):
    """返回 (字号集合, 是否全部加粗, 是否全部斜体, 颜色列表, 字体是否合规)。"""
    mr = _main_runs(r)
    if not mr:
        return set(), False, False, [], True
    szs = set(round(x['sz'], 2) for x in mr)
    bold = all(x['b'] for x in mr)
    ital = all(x['i'] for x in mr)
    cols = [x['color'] for x in mr if x['color'] is not None]
    fonts = all(_font_ok(x['name']) for x in mr)
    return szs, bold, ital, cols, fonts


def _sz_in(r, lo, hi):
    szs, _, _, _, _ = _run_stat(r)
    return bool(szs) and all(_btw(s, lo, hi) for s in szs)


def _sz_main(r):
    szs, _, _, _, _ = _run_stat(r)
    return max(szs) if szs else None


def _any_col(r, pred):
    _, _, _, cols, _ = _run_stat(r)
    return any(pred(c) for c in cols)


def _cluster(recs, x0, x1, y0, y1, min_n=2, exclude=()):
    """“由多个形状拼合的图形”：取完全落在窗口内的无文字形状，返回 (成员数, 宽cm, 高cm, 并集框)。
    连接线（<p:cxnSp> 的 *Connector*）只是牵引线，不计入拼合图形的组成部分。"""
    mem = [r for r in recs
           if not r['has_text'] and r['i'] not in exclude
           and 'onnector' not in (r['geom'] or '')
           and _inside(r, x0, x1, y0, y1)]
    if len(mem) < min_n:
        return None
    xl = min(r['xl'] for r in mem)
    xr = max(r['xr'] for r in mem)
    yt = min(r['yt'] for r in mem)
    yb = max(r['yb'] for r in mem)
    return (len(mem), (xr - xl) / 100.0 * _PAGE_W_CM, (yb - yt) / 100.0 * _PAGE_H_CM,
            (xl, xr, yt, yb))


def _add(results, score, rule, ok, detail=''):
    """登记一条 "+N：说明" 评分项；命中记 +score，未命中记 +0。

    detail 是判定时收集的实测值，只用于本地调试，不写入返回信息：命中与未命中
    两个分支登记的规则原文完全一致，返回的 rule 不随判定结果变化。
    """
    results.append('+%d：%s' % (score if ok else 0, rule))


# ========== 维度1：可用与可修改性（门禁） ==========
def _check_dim1(prs, slides, pages, pw_cm, ph_cm):
    """①交付文件为 .pptx 且可正常打开（能被解析即视为可打开）；②只包含 2 页幻灯片；
    ③没有整页不可编辑的情况（某页没有任何可编辑文字、内容为整页图片或空白，即判整页不可编辑）；
    ④形状与文字均可单独选中、移动、修改、删除（<a:spLocks>/<a:cxnSpLocks> 未设 noSelect/noMove/noResize/noTextEdit）。
    与办公软件对齐：读取 ppt/slides/slideN.xml 的 <p:sp>/<p:pic> 结构与 <a:*Locks> 属性。"""
    if len(slides) != 2:
        return False, '交付PPT包含 %d 页幻灯片，不满足“只包含2页幻灯片”' % len(slides)
    for si, recs in enumerate(pages, 1):
        if not recs:
            return False, '第%d页幻灯片没有任何可编辑形状' % si
        page_area = pw_cm * ph_cm
        texts = [r for r in recs if r['has_text']]
        pics = [r for r in recs if r['is_pic']]
        vecs = [r for r in recs
                if not r['is_pic'] and (r['fill'] is not None or _ln_of(r['el']) is not None)]
        if not texts:
            return False, '第%d页幻灯片没有任何可编辑文字，内容为整页图片或空白，属整页不可编辑' % si
        big_pic = [r for r in pics if r['w'] * r['h'] >= 0.60 * page_area]
        if big_pic and len(texts) <= 1 and not vecs:
            return False, '第%d页幻灯片主体为大幅图片且几乎没有可编辑内容，属整页不可编辑' % si
    for si, slide in enumerate(slides, 1):
        xml = slide._element.xml
        for attr in ('noSelect', 'noMove', 'noResize', 'noTextEdit', 'noEditPoints', 'noDrilldown'):
            if re.search(r'Locks[^>]*\s%s="(1|true)"' % attr, xml):
                return False, '第%d页幻灯片存在被锁定（%s）的形状，无法单独选中或修改' % (si, attr)
    return True, ''


# ========== 对外接口 ==========
def evaluate(dir_path):
    """按 §2.2 约定，返回结构化 dict；不 print、不改 sys.stdout、不 sys.exit。

    status='error' 只在“这份提交没被真正评过”时出现——目录不存在、目录内没有
    .pptx、或脚本自身抛异常；此时 dim1_pass 为空串（未判定）、dim1_reason 为
    空串、dim2_items 为空表、total_score 与 max_score 一律为 0。
    文件能定位到但门禁不达标，仍按 status='ok' + 维度一未通过处理，此时
    dim1_pass=False 且 dim1_reason 带有未通过的错误摘要。"""
    result = {
        "id": SCRIPT_ID,
        "file_name": "",
        "status": "ok",
        "error": None,
        # 空串表示“维度一尚未判定”；status='error' 的三条出口都保持空串，
        # 只有真正跑过 _check_dim1 才会写入 True / False。
        "dim1_pass": "",
        "dim1_reason": "",
        "dim2_items": [],
        "total_score": 0,
        "max_score": 0,
    }
    try:
        if not dir_path or not os.path.isdir(dir_path):
            result["status"] = "error"
            result["error"] = f"目录不存在：{dir_path}"
            return result
        cands = sorted(
            f for f in os.listdir(dir_path)
            if f.lower().endswith('.pptx') and not f.startswith('~$')
        )
        if not cands:
            result["status"] = "error"
            result["error"] = f"目录中未找到 .pptx 文件：{dir_path}"
            return result
        file_path = os.path.join(dir_path, cands[0])
        result["file_name"] = os.path.basename(file_path)
        return _evaluate_impl(file_path, result)
    except Exception as exc:
        # _evaluate_impl 是就地写 result 的，异常可能发生在它写了一半之后，
        # 因此这里把所有承载评分结论的字段显式复位，保证 error 的返回值是干净的全 0；
        # dim1_pass 复位为空串而不是 False——异常意味着维度一的结论作废，不是判定为不通过。
        result["status"] = "error"
        result["error"] = f"{type(exc).__name__}: {exc}"
        result["dim1_pass"] = ""
        result["dim1_reason"] = ""
        result["dim2_items"] = []
        result["total_score"] = 0
        result["max_score"] = 0
        return result


def _notes_size_cm(file_path):
    try:
        with zipfile.ZipFile(file_path) as z:
            txt = z.read('ppt/presentation.xml').decode('utf8', 'ignore')
        m = re.search(r'<p:notesSz[^>]*cx="(\d+)"[^>]*cy="(\d+)"', txt)
        if m:
            return int(m.group(1)) / _CM, int(m.group(2)) / _CM
    except Exception:
        pass
    return None, None


def _slides_xml(file_path):
    out = []
    try:
        with zipfile.ZipFile(file_path) as z:
            for name in ('ppt/slides/slide1.xml', 'ppt/slides/slide2.xml'):
                if name in z.namelist():
                    out.append(z.read(name).decode('utf8', 'ignore'))
    except Exception:
        pass
    return out


def _evaluate_impl(file_path, result):
    prs = Presentation(file_path)
    slides = list(prs.slides)
    pw_cm = float(prs.slide_width) / _CM
    ph_cm = float(prs.slide_height) / _CM
    scheme = _theme_scheme(prs)
    default_pt = _default_sz_pt(prs)
    pages = [_collect(sl, pw_cm, ph_cm, default_pt, scheme) for sl in slides]
    s1 = pages[0] if len(pages) > 0 else []
    s2 = pages[1] if len(pages) > 1 else []
    xmls = _slides_xml(file_path)

    dim1_ok, dim1_reason = _check_dim1(prs, slides, pages, pw_cm, ph_cm)
    result["dim1_pass"] = bool(dim1_ok)
    # 约定：dim1_pass 为 False 时 dim1_reason 必须带错误摘要，通过时必须为空串。
    if dim1_ok:
        result["dim1_reason"] = ""
    else:
        result["dim1_reason"] = dim1_reason or "维度一未通过（未给出具体原因）"
        # 维度一为门禁：未通过时不再逐条评维度二（未真正逐条评过），dim2_items 留空、
        # total_score 记 0，max_score 仍为满分常量。
        result["dim2_items"] = []
        result["total_score"] = 0
        result["max_score"] = MAX_SCORE_TOTAL
        return result

    results = []

    # ---------------------------------------------------------------- 01
    # 规则：第1页幻灯片距左侧页边线18%–88%、距顶部边线0%–25%位置出现浅蓝色圆角矩形形状：
    #       整体宽20-23cm，高3.5-4.5cm；其内距左侧页边线38%–64%、距顶部边线1%–10%位置出现
    #       “Macro Environment”文本，字体为Arial或主题默认西文字体，加粗，14-20磅，颜色为深蓝色。
    # ①圆角矩形位置窗口 ②整体宽高区间 ③浅蓝色填充 ④文本内容与位置 ⑤字体/加粗/字号/深蓝色
    # 与办公软件对齐：读 <a:prstGeom prst="roundRect">、<a:xfrm>/<a:off>/<a:ext>、<a:solidFill>
    #                （无显式填充时回落 <p:style><a:fillRef>）、<a:rPr b/sz/latin> 与 <a:solidFill>。
    rule = '第1页幻灯片距左侧页边线18%–88%、距顶部边线0%–25%位置出现浅蓝色圆角矩形形状：整体宽20-23cm，高3.5-4.5cm；其内距左侧页边线38%–64%、距顶部边线1%–10%位置出现“Macro Environment”文本，字体为Arial或主题默认西文字体，加粗，14-20磅，颜色为深蓝色'
    box = [r for r in s1 if r['geom'] == 'roundRect' and _in_win(r, 18, 88, 0, 25)
           and _btw(r['w'], 20, 23) and _btw(r['h'], 3.5, 4.5) and _is_light_blue(r['fill'])]
    ok, detail = False, ''
    for r in _text_recs(s1, 'Macro Environment', 38, 64, 1, 10):
        szs, bold, _, _, fonts = _run_stat(r)
        if box and bold and fonts and szs and all(_btw(s, 14, 20) for s in szs) and _any_col(r, _is_dark_blue):
            ok = True
            detail = '圆角矩形宽%.2fcm高%.2fcm位于x%.1f%%y%.1f%%，文本字号%.1f磅加粗' % (
                box[0]['w'], box[0]['h'], box[0]['pcx'], box[0]['pcy'], max(szs))
            break
    _add(results, 3, rule, ok, detail)


    # ---------------------------------------------------------------- 02
    # 规则：第1页幻灯片距左侧页边线22%–83%、距顶部边线4%–16%位置自左至右依次出现
    #       “Technological advances”…“Socioeconomic trends”六段文本：每段单独成块、互不重叠；
    #       字体为Arial或主题默认西文字体，加粗，8-11磅。
    # ①六段文本齐备 ②各自位置在窗口内 ③自左至右顺序 ④分属六个独立形状且互不重叠 ⑤字体/加粗/8-11磅
    # 与办公软件对齐：读 <p:sp> 的 <a:xfrm> 与 <a:t> 文本、<a:rPr b="1" sz>、<a:latin typeface>。
    rule = '第1页幻灯片距左侧页边线22%–83%、距顶部边线4%–16%位置自左至右依次出现“Technological advances”、“Policy & reforms”、“Funding landscape”、“Demographic shifts”、“Global health priorities”、“Socioeconomic trends”六段文本：每段单独成块、互不重叠；字体为Arial或主题默认西文字体，加粗，8-11磅'
    labels = ['Technological advances', 'Policy & reforms', 'Funding landscape',
              'Demographic shifts', 'Global health priorities', 'Socioeconomic trends']
    got = []
    for name in labels:
        hit = [r for r in _text_recs(s1, name, 22, 83, 4, 16)
               if _sz_in(r, 8, 11) and _run_stat(r)[1] and _run_stat(r)[4]]
        got.append(hit[0] if hit else None)
    ok = all(g is not None for g in got) and len(set(g['i'] for g in got if g)) == 6
    if ok:
        ok = all(got[k]['pcx'] < got[k + 1]['pcx'] for k in range(5))
    if ok:
        for a in range(6):
            for b in range(a + 1, 6):
                ra, rb = got[a], got[b]
                ox = max(0.0, min(ra['xr'], rb['xr']) - max(ra['xl'], rb['xl']))
                oy = max(0.0, min(ra['yb'], rb['yb']) - max(ra['yt'], rb['yt']))
                small = min((ra['xr'] - ra['xl']) * (ra['yb'] - ra['yt']),
                            (rb['xr'] - rb['xl']) * (rb['yb'] - rb['yt']))
                if small > 0 and ox * oy > 0.2 * small:
                    ok = False
    _add(results, 3, rule, ok,
         '六段文本均为独立形状、加粗、字号%s磅' % (
             '/'.join('%.1f' % _sz_main(g) for g in got) if ok else ''))

    # ---------------------------------------------------------------- 03
    # 规则：第1页幻灯片距左侧页边线19%–84%、距顶部边线10%–24%位置出现四条自左向右横贯的平滑曲线：
    #       四条曲线横向合计跨度19-22cm；线型至少出现实线、虚线、点划线三种；右端出现箭头。
    # ①平滑曲线（自由曲线 custGeom，非折线段拼接）至少四条 ②横向合计跨度19-22cm
    # ③线型至少覆盖实线/虚线/点划线三类 ④右端存在箭头
    # 与办公软件对齐：读 <a:custGeom>（平滑曲线）、<a:prstDash val>、<a:tailEnd/headEnd type>。
    rule = '第1页幻灯片距左侧页边线19%–84%、距顶部边线10%–24%位置出现四条自左向右横贯的平滑曲线：四条曲线横向合计跨度19-22cm；线型至少出现实线、虚线、点划线三种；右端出现箭头'
    curves = [r for r in s1 if r['geom'] == 'custGeom' and _in_win(r, 19, 84, 10, 24)
              and r['w'] >= 1.0]
    span = 0.0
    if curves:
        span = (max(c['xr'] for c in curves) - min(c['xl'] for c in curves)) / 100.0 * pw_cm
    classes = set(c['dash'] for c in curves)
    arrow = any(c['he'] or c['te'] for c in curves)
    ok = (len(curves) >= 4 and _btw(span, 19, 22)
          and {'solid', 'dash', 'dashdot'} <= classes and arrow)
    _add(results, 5, rule, ok, '%d条平滑曲线，横向跨度%.2fcm，线型%s，右端带箭头' % (
        len(curves), span, '/'.join(sorted(classes))))


    # ---------------------------------------------------------------- 04
    # 规则：第1页幻灯片距顶部边线20%–30%位置出现两个竖直向下的箭头形状：分别位于距左侧页边线
    #       32%–42%与62%–70%范围内；填充为蓝色，每个宽0.5-1.2cm，高0.8-1.6cm。
    # ①两个下箭头 ②左右位置窗口 ③纵向位置窗口 ④蓝色填充 ⑤各自宽高区间
    # 与办公软件对齐：读 <a:prstGeom prst="downArrow">、<a:ext cx/cy>、<a:solidFill>
    #                （无显式填充时回落 <p:style><a:fillRef>）。
    rule = '第1页幻灯片距顶部边线20%–30%位置出现两个竖直向下的箭头形状：分别位于距左侧页边线32%–42%与62%–70%范围内；填充为蓝色，每个宽0.5-1.2cm，高0.8-1.6cm'
    def _down_arrows(recs, x0, x1, x2, x3, y0, y1, pred):
        hits = []
        for lo, hi in ((x0, x1), (x2, x3)):
            cand = [r for r in recs if r['geom'] in ('downArrow', 'downArrowCallout')
                    and _in_win(r, lo, hi, y0, y1)
                    and _btw(r['w'], 0.5, 1.2) and _btw(r['h'], 0.8, 1.6) and pred(r['fill'])]
            hits.append(cand[0] if cand else None)
        return hits
    a1 = _down_arrows(s1, 32, 42, 62, 70, 20, 30, _is_blue)
    ok = all(x is not None for x in a1)
    _add(results, 1, rule, ok, '两个蓝色下箭头，宽%s高%s' % (
        '/'.join('%.2f' % x['w'] for x in a1 if x), '/'.join('%.2f' % x['h'] for x in a1 if x)))

    # ---------------------------------------------------------------- 05
    # 规则：第1页幻灯片距左侧页边线26%–78%、距顶部边线23%–55%位置出现浅绿色椭圆形状：
    #       整体宽14-16cm，高4.5-5.8cm；其内距左侧页边线35%–69%、距顶部边线26%–35%位置出现
    #       “Local Health Innovation Context”文本，字体为Arial或主题默认西文字体，加粗，13-17磅，颜色为深绿色。
    # ①椭圆位置窗口 ②整体宽高区间 ③浅绿色填充 ④文本内容与位置 ⑤字体/加粗/字号/深绿色
    # 与办公软件对齐：读 <a:prstGeom prst="ellipse">、<a:ext>、<a:solidFill>、<a:rPr b/sz> 与其 <a:solidFill>。
    rule = '第1页幻灯片距左侧页边线26%–78%、距顶部边线23%–55%位置出现浅绿色椭圆形状：整体宽14-16cm，高4.5-5.8cm；其内距左侧页边线35%–69%、距顶部边线26%–35%位置出现“Local Health Innovation Context”文本，字体为Arial或主题默认西文字体，加粗，13-17磅，颜色为深绿色'
    ell = [r for r in s1 if r['geom'] == 'ellipse' and _in_win(r, 26, 78, 23, 55)
           and _btw(r['w'], 14, 16) and _btw(r['h'], 4.5, 5.8) and _is_light_green(r['fill'])]
    ok, detail = False, ''
    for r in _text_recs(s1, 'Local Health Innovation Context', 35, 69, 26, 35):
        szs, bold, _, _, fonts = _run_stat(r)
        if ell and bold and fonts and szs and all(_btw(s, 13, 17) for s in szs) and _any_col(r, _is_dark_green):
            ok = True
            detail = '椭圆宽%.2fcm高%.2fcm，文本字号%.1f磅加粗深绿色' % (ell[0]['w'], ell[0]['h'], max(szs))
            break
    _add(results, 3, rule, ok, detail)


    # ---------------------------------------------------------------- 06 / 07 公共判定
    def _trio_ellipse(recs, x0, x1, y0, y1, labels):
        """三个小椭圆 + 其内三段不加粗 8-11 磅文本（文本常为独立文本框，按中心落入椭圆判定归属）。"""
        els = sorted([r for r in recs if r['geom'] == 'ellipse' and _in_win(r, x0, x1, y0, y1)
                      and _btw(r['w'], 2.4, 3.2) and _btw(r['h'], 1.0, 1.6)],
                     key=lambda r: r['pcx'])
        if len(els) < 3:
            return False, ''
        els = els[:3]
        texts = []
        for name in labels:
            hit = None
            for r in _text_recs(recs, name):
                if not _sz_in(r, 8, 11):
                    continue
                if any(x['b'] for x in _main_runs(r)):
                    continue
                if not _run_stat(r)[4]:
                    continue
                for e in els:
                    if (e['xl'] - 1.0 <= r['pcx'] <= e['xr'] + 1.0
                            and e['yt'] - 2.0 <= r['pcy'] <= e['yb'] + 2.0):
                        hit = (r, e)
                        break
                if hit:
                    break
            texts.append(hit)
        if any(t is None for t in texts):
            return False, ''
        if [t[1]['i'] for t in texts] != [e['i'] for e in els]:
            return False, ''
        return True, '三个椭圆宽%s高%s，文本不加粗、字号%s磅' % (
            '/'.join('%.2f' % e['w'] for e in els), '/'.join('%.2f' % e['h'] for e in els),
            '/'.join('%.1f' % _sz_main(t[0]) for t in texts))

    # ---------------------------------------------------------------- 06
    # 规则：第1页幻灯片距左侧页边线30%–72%、距顶部边线30%–42%位置自左至右出现三个小椭圆形状：
    #       每个宽2.4-3.2cm，高1.0-1.6cm；其中文本依次为“clinical infrastructure”、“data resources”、
    #       “regulatory support”，字体为Arial或主题默认西文字体，不加粗，8-11磅。
    # ①三个椭圆 ②每个宽高区间 ③位置窗口 ④文本内容与自左至右次序 ⑤字体/不加粗/8-11磅
    # 与办公软件对齐：读 <a:prstGeom prst="ellipse">、<a:ext>、<a:t>、<a:rPr b/sz/latin>。
    rule = '第1页幻灯片距左侧页边线30%–72%、距顶部边线30%–42%位置自左至右出现三个小椭圆形状：每个宽2.4-3.2cm，高1.0-1.6cm；其中文本依次为“clinical infrastructure”、“data resources”、“regulatory support”，字体为Arial或主题默认西文字体，不加粗，8-11磅'
    ok, detail = _trio_ellipse(s1, 30, 72, 30, 42,
                              ['clinical infrastructure', 'data resources', 'regulatory support'])
    _add(results, 3, rule, ok, detail)

    # ---------------------------------------------------------------- 07
    # 规则：第1页幻灯片距左侧页边线33%–69%、距顶部边线39%–52%位置自左至右出现三个小椭圆形状：
    #       每个宽2.4-3.2cm，高1.0-1.6cm；其中文本依次为“talent base”、“public trust”、
    #       “research networks”，字体为Arial或主题默认西文字体，不加粗，8-11磅。
    # ①三个椭圆 ②每个宽高区间 ③位置窗口 ④文本内容与自左至右次序 ⑤字体/不加粗/8-11磅
    # 与办公软件对齐：同上，读 <a:prstGeom>/<a:ext>/<a:t>/<a:rPr>。
    rule = '第1页幻灯片距左侧页边线33%–69%、距顶部边线39%–52%位置自左至右出现三个小椭圆形状：每个宽2.4-3.2cm，高1.0-1.6cm；其中文本依次为“talent base”、“public trust”、“research networks”，字体为Arial或主题默认西文字体，不加粗，8-11磅'
    ok, detail = _trio_ellipse(s1, 33, 69, 39, 52,
                              ['talent base', 'public trust', 'research networks'])
    _add(results, 3, rule, ok, detail)

    # ---------------------------------------------------------------- 08
    # 规则：第1页幻灯片距顶部边线47%–57%位置出现两个竖直向下的箭头形状：分别位于距左侧页边线
    #       32%–42%与62%–70%范围内；填充为绿色，每个宽0.5-1.2cm，高0.8-1.6cm。
    # ①两个下箭头 ②左右位置窗口 ③纵向位置窗口 ④绿色填充 ⑤各自宽高区间
    # 与办公软件对齐：读 <a:prstGeom prst="downArrow">、<a:ext>、<a:solidFill> 或 <p:style><a:fillRef>。
    rule = '第1页幻灯片距顶部边线47%–57%位置出现两个竖直向下的箭头形状：分别位于距左侧页边线32%–42%与62%–70%范围内；填充为绿色，每个宽0.5-1.2cm，高0.8-1.6cm'
    a2 = _down_arrows(s1, 32, 42, 62, 70, 47, 57, _is_green)
    ok = all(x is not None for x in a2)
    _add(results, 1, rule, ok, '两个绿色下箭头，宽%s高%s' % (
        '/'.join('%.2f' % x['w'] for x in a2 if x), '/'.join('%.2f' % x['h'] for x in a2 if x)))


    # ---------------------------------------------------------------- 09
    # 规则：第1页幻灯片距左侧页边线24%–79%、距顶部边线52%–94%位置出现圆角矩形形状：整体宽15-18cm，
    #       高6.2-7.6cm；其顶部为深蓝色横向标题条，条内距左侧页边线38%–62%、距顶部边线52%–61%位置
    #       出现白色文本“Innovation Dynamics Arena”，字体为Arial或主题默认西文字体，加粗，10-14磅。
    # ①圆角矩形位置与整体宽高 ②顶部存在深蓝色填充的横向标题条 ③白色标题文本内容与位置 ④字体/加粗/10-14磅
    # 与办公软件对齐：读 <a:prstGeom prst="roundRect">、<a:ext>、<a:solidFill>（含 schemeClr 经 clrMap/lumMod 解析）、<a:rPr b/sz>。
    rule = '第1页幻灯片距左侧页边线24%–79%、距顶部边线52%–94%位置出现圆角矩形形状：整体宽15-18cm，高6.2-7.6cm；其顶部为深蓝色横向标题条，条内距左侧页边线38%–62%、距顶部边线52%–61%位置出现白色文本“Innovation Dynamics Arena”，字体为Arial或主题默认西文字体，加粗，10-14磅'
    card = [r for r in s1 if r['geom'] == 'roundRect' and _in_win(r, 24, 79, 52, 94)
            and _btw(r['w'], 15, 18) and _btw(r['h'], 6.2, 7.6)]
    bar = [r for r in s1 if _is_dark_blue(r['fill']) and _in_win(r, 38, 62, 52, 61)
           and r['w'] > r['h']]
    ok, detail = False, ''
    for r in _text_recs(s1, 'Innovation Dynamics Arena', 38, 62, 52, 61):
        szs, bold, _, _, fonts = _run_stat(r)
        if card and bar and bold and fonts and szs and all(_btw(s, 10, 14) for s in szs) and _any_col(r, _is_white):
            ok = True
            detail = '圆角矩形宽%.2fcm高%.2fcm，深蓝标题条宽%.2fcm，白色文本%.1f磅加粗' % (
                card[0]['w'], card[0]['h'], bar[0]['w'], max(szs))
            break
    _add(results, 3, rule, ok, detail)

    # ---------------------------------------------------------------- 10
    # 规则：第1页幻灯片距左侧页边线26%–43%、距顶部边线58%–89%位置上下出现两个虚线边框矩形：
    #       每个宽3.5-4.5cm，高1.4-2.0cm；上框文本为“Service demand”、下框文本为“Equity needs”，
    #       字体为Arial或主题默认西文字体，加粗，10-13磅；两框之间出现一条长1.0-2.0cm的竖直双向箭头相连。
    # ①两个矩形位置与宽高 ②边框为虚线 ③上下框文本内容与字体/加粗/10-13磅 ④两框之间竖直双向箭头且长度1.0-2.0cm
    # 与办公软件对齐：读 <a:prstDash val="sysDash/dash">、<a:ext>、同一 <a:ln> 上的 <a:headEnd>+<a:tailEnd>。
    rule = '第1页幻灯片距左侧页边线26%–43%、距顶部边线58%–89%位置上下出现两个虚线边框矩形：每个宽3.5-4.5cm，高1.4-2.0cm；上框文本为“Service demand”、下框文本为“Equity needs”，字体为Arial或主题默认西文字体，加粗，10-13磅；两框之间出现一条长1.0-2.0cm的竖直双向箭头相连'
    boxes = sorted([r for r in s1 if r['geom'] in ('rect', 'roundRect') and _in_win(r, 26, 43, 58, 89)
                    and _btw(r['w'], 3.5, 4.5) and _btw(r['h'], 1.4, 2.0) and r['dash'] != 'solid'],
                   key=lambda r: r['pcy'])
    t_up = [r for r in _text_recs(s1, 'Service demand', 26, 43, 58, 89)
            if _sz_in(r, 10, 13) and _run_stat(r)[1] and _run_stat(r)[4]]
    t_dn = [r for r in _text_recs(s1, 'Equity needs', 26, 43, 58, 89)
            if _sz_in(r, 10, 13) and _run_stat(r)[1] and _run_stat(r)[4]]
    dbl = []
    if len(boxes) >= 2:
        lo, hi = boxes[0]['yb'], boxes[-1]['yt']
        dbl = [r for r in s1 if r['he'] and r['te'] and r['h'] > r['w']
               and _btw(r['h'], 1.0, 2.0) and _in_win(r, 26, 43, lo - 4, hi + 4)]
    ok = (len(boxes) >= 2 and t_up and t_dn and dbl
          and t_up[0]['pcy'] < t_dn[0]['pcy'])
    _add(results, 5, rule, bool(ok), '两个%s边框矩形宽%.2fcm高%.2fcm，双向箭头长%.2fcm' % (
        boxes[0]['dash'] if boxes else '', boxes[0]['w'] if boxes else 0,
        boxes[0]['h'] if boxes else 0, dbl[0]['h'] if dbl else 0))


    # ---------------------------------------------------------------- 11
    # 规则：第1页幻灯片距左侧页边线27%–35%位置出现两个由多个形状拼合的图形或图片：一个位于距顶部边线
    #       61%–70%范围内，为柱状图图案；另一个位于距顶部边线78%–87%范围内，为多人并排的人物图案；
    #       每个整体宽0.3-1.4cm，高0.3-1.4cm。
    # ①两处图案均由多个形状拼合（成员≥3）或为图片 ②各自纵向位置窗口 ③横向位置窗口 ④每个整体宽高0.3-1.4cm
    # 与办公软件对齐：统计窗口内无文字 <p:sp>/<p:pic> 的数量与 <a:off>/<a:ext> 并集外接框。
    rule = '第1页幻灯片距左侧页边线27%–35%位置出现两个由多个形状拼合的图形或图片：一个位于距顶部边线61%–70%范围内，为柱状图图案；另一个位于距顶部边线78%–87%范围内，为多人并排的人物图案；每个整体宽0.3-1.4cm，高0.3-1.4cm'
    c_bar = _cluster(s1, 27, 35, 61, 70, min_n=3)
    c_ppl = _cluster(s1, 27, 35, 78, 87, min_n=3)
    ok = bool(c_bar and c_ppl
              and _btw(c_bar[1], 0.3, 1.4) and _btw(c_bar[2], 0.3, 1.4)
              and _btw(c_ppl[1], 0.3, 1.4) and _btw(c_ppl[2], 0.3, 1.4))
    _add(results, 3, rule, ok, '柱状图由%d个形状拼合(%.2f×%.2fcm)，人物图案由%d个形状拼合(%.2f×%.2fcm)' % (
        (c_bar[0], c_bar[1], c_bar[2], c_ppl[0], c_ppl[1], c_ppl[2]) if ok else (0, 0, 0, 0, 0, 0)))

    # ---------------------------------------------------------------- 12
    # 规则：第1页幻灯片距左侧页边线40%–46%、距顶部边线57%–91%位置出现一条竖直虚线：长4.8-6.2cm，
    #       线型为虚线或点线。
    # ①竖直方向（高>宽） ②位置窗口 ③长度4.8-6.2cm ④线型为虚线或点线
    # 与办公软件对齐：读 <a:prstDash val>（dash/sysDash/lgDash 或 dot/sysDot）与 <a:ext>。
    rule = '第1页幻灯片距左侧页边线40%–46%、距顶部边线57%–91%位置出现一条竖直虚线：长4.8-6.2cm，线型为虚线或点线'
    vlines = [r for r in s1 if r['h'] > r['w'] and _in_win(r, 40, 46, 57, 91)
              and _btw(r['h'], 4.8, 6.2) and r['dash'] in ('dash', 'dot', 'dashdot')]
    ok = bool(vlines)
    _add(results, 1, rule, ok, '竖直%s线长%.2fcm' % (
        vlines[0]['dash'] if ok else '', vlines[0]['h'] if ok else 0))

    # ---------------------------------------------------------------- 13
    # 规则：第1页幻灯片距左侧页边线42%–74%、距顶部边线57%–90%位置出现四个椭圆形状并按左上、右上、
    #       左下、右下四角分布：每个宽2.8-3.6cm，高1.4-2.0cm；文本依次为“Hospitals”、“Start-ups”、
    #       “Government”、“Citizens”，字体为Arial或主题默认西文字体，不加粗，10-13磅，四者字号与加粗设置一致。
    # ①四个椭圆及宽高 ②四角分布 ③四段文本内容与归属 ④字体/不加粗/10-13磅 ⑤四者字号与加粗一致
    # 与办公软件对齐：读 <a:prstGeom prst="ellipse">、<a:ext>、<a:t>、<a:rPr b/sz>。
    rule = '第1页幻灯片距左侧页边线42%–74%、距顶部边线57%–90%位置出现四个椭圆形状并按左上、右上、左下、右下四角分布：每个宽2.8-3.6cm，高1.4-2.0cm；文本依次为“Hospitals”、“Start-ups”、“Government”、“Citizens”，字体为Arial或主题默认西文字体，不加粗，10-13磅，四者字号与加粗设置一致'
    quad = [r for r in s1 if r['geom'] == 'ellipse' and _in_win(r, 42, 74, 57, 90)
            and _btw(r['w'], 2.8, 3.6) and _btw(r['h'], 1.4, 2.0)]
    labs = ['Hospitals', 'Start-ups', 'Government', 'Citizens']
    found = []
    for name in labs:
        hit = [r for r in _text_recs(s1, name, 42, 74, 57, 90) if _main_runs(r)]
        found.append(hit[0] if hit else None)
    ok, detail = False, ''
    if len(quad) >= 4 and all(f is not None for f in found):
        szset = set()
        bset = set()
        for f in found:
            szs, _, _, _, fonts = _run_stat(f)
            szset |= szs
            bset.add(any(x['b'] for x in _main_runs(f)))
            if not fonts:
                bset.add('bad')
        cs = [(round(q['pcx'], 1), round(q['pcy'], 1)) for q in quad[:4]]
        mx = sum(c[0] for c in cs) / 4.0
        my = sum(c[1] for c in cs) / 4.0
        corners = set((c[0] < mx, c[1] < my) for c in cs)
        ok = (len(szset) == 1 and all(_btw(s, 10, 13) for s in szset)
              and bset == {False} and len(corners) == 4)
        detail = '四个椭圆四角分布，文本字号统一%.1f磅且均不加粗' % max(szset)
    _add(results, 3, rule, ok, detail if ok else '')


    # ---------------------------------------------------------------- 14
    # 规则：第1页幻灯片距左侧页边线46%–70%、距顶部边线62%–89%位置出现四个由多个形状拼合的图形或图片：
    #       每个整体宽0.2-1.2cm，高0.2-1.2cm；图案分别为医院十字、火箭、银行、人群。
    # ①四处图案分居左上/右上/左下/右下 ②每处均由多个形状拼合（成员≥2），且不以符号字符充当图案
    # ③每处整体宽0.2-1.2cm ④每处整体高0.2-1.2cm
    # 与办公软件对齐：统计各象限内无文字 <p:sp>/<p:pic> 的 <a:off>/<a:ext> 并集外接框（文字字符形状不计）；
    #               象限内若出现 ✚/✓/★/⚡ 一类符号字符的文本，说明该图案是打字打出来的而非画出来的。
    rule = '第1页幻灯片距左侧页边线46%–70%、距顶部边线62%–89%位置出现四个由多个形状拼合的图形或图片：每个整体宽0.2-1.2cm，高0.2-1.2cm；图案分别为医院十字、火箭、银行、人群'
    quads = [(46, 58, 62, 75), (58, 70, 62, 75), (46, 58, 75, 89), (58, 70, 75, 89)]
    cl4 = [_cluster(s1, *q, min_n=2) for q in quads]
    glyphs = [r['i'] for r in s1
              if _GLYPH_RE.search(r['text']) and any(_in_win(r, *q) for q in quads)]
    ok = (not glyphs
          and all(c is not None and _btw(c[1], 0.2, 1.2) and _btw(c[2], 0.2, 1.2) for c in cl4))
    _add(results, 5, rule, ok, '四处拼合图案尺寸%s' % (
        '、'.join('%.2f×%.2fcm' % (c[1], c[2]) for c in cl4) if ok else ''))

    # ---------------------------------------------------------------- 15
    # 规则：第1页幻灯片出现四条实线双向箭头：两条水平箭头分别位于距左侧页边线52%–64%、距顶部边线
    #       62%–66%与81%–85%范围内，每条长2.2-3.2cm；两条竖直箭头分别位于距左侧页边线47%–51%与
    #       65%–69%、距顶部边线67%–81%范围内，每条长1.5-2.2cm。
    # ①同一形状上同时具备 headEnd 与 tailEnd（真双向箭头，而非两条单向箭头叠放） ②线型为实线
    # ③两条水平箭头位置与长度 ④两条竖直箭头位置与长度
    # 与办公软件对齐：读同一 <a:ln> 内的 <a:headEnd type>/<a:tailEnd type> 与 <a:prstDash>、<a:ext>。
    rule = '第1页幻灯片出现四条实线双向箭头：两条水平箭头分别位于距左侧页边线52%–64%、距顶部边线62%–66%与81%–85%范围内，每条长2.2-3.2cm；两条竖直箭头分别位于距左侧页边线47%–51%与65%–69%、距顶部边线67%–81%范围内，每条长1.5-2.2cm'
    dblarr = [r for r in s1 if r['he'] and r['te'] and r['dash'] == 'solid']
    h_top = [r for r in dblarr if r['w'] > r['h'] and _in_win(r, 52, 64, 62, 66) and _btw(r['w'], 2.2, 3.2)]
    h_bot = [r for r in dblarr if r['w'] > r['h'] and _in_win(r, 52, 64, 81, 85) and _btw(r['w'], 2.2, 3.2)]
    v_l = [r for r in dblarr if r['h'] > r['w'] and _in_win(r, 47, 51, 67, 81) and _btw(r['h'], 1.5, 2.2)]
    v_r = [r for r in dblarr if r['h'] > r['w'] and _in_win(r, 65, 69, 67, 81) and _btw(r['h'], 1.5, 2.2)]
    ok = bool(h_top and h_bot and v_l and v_r)
    _add(results, 5, rule, ok, '水平双向箭头长%.2f/%.2fcm，竖直双向箭头长%.2f/%.2fcm' % (
        (h_top[0]['w'], h_bot[0]['w'], v_l[0]['h'], v_r[0]['h']) if ok else (0, 0, 0, 0)))

    # ---------------------------------------------------------------- 16
    # 规则：第1页幻灯片距左侧页边线49%–68%、距顶部边线68%–82%位置出现“alignment, negotiation, and
    #       shared experimentation”文本：字体为Arial或主题默认西文字体，斜体，10-13磅；该文本左上、右上、
    #       左下、右下四个方向各出现一条自该文本向外指出的点线箭头，每条长0.65-2.6cm。
    # ①文本内容与位置窗口 ②字体/斜体/10-13磅 ③四个斜向方向各一条点线箭头 ④每条长0.65-2.6cm
    # 与办公软件对齐：读 <a:rPr i="1" sz>、<a:prstDash val="sysDot/dot">、<a:tailEnd type>、<a:ext>。
    rule = '第1页幻灯片距左侧页边线49%–68%、距顶部边线68%–82%位置出现“alignment, negotiation, and shared experimentation”文本：字体为Arial或主题默认西文字体，斜体，10-13磅；该文本左上、右上、左下、右下四个方向各出现一条自该文本向外指出的点线箭头，每条长0.65-2.6cm'
    ok, detail = False, ''
    for r in _text_recs(s1, 'alignment, negotiation, and shared experimentation', 49, 68, 68, 82):
        szs, _, ital, _, fonts = _run_stat(r)
        if not (ital and fonts and szs and all(_btw(s, 10, 13) for s in szs)):
            continue
        dirs = {}
        for a in s1:
            if a['dash'] != 'dot' or not (a['he'] or a['te']):
                continue
            length = math.hypot(a['w'], a['h'])
            if not _btw(length, 0.65, 2.6):
                continue
            if abs(a['pcx'] - r['pcx']) > 12 or abs(a['pcy'] - r['pcy']) > 18:
                continue
            dirs[(a['pcx'] < r['pcx'], a['pcy'] < r['pcy'])] = length
        if len(dirs) == 4:
            ok = True
            detail = '斜体%.1f磅，四向点线箭头长%s' % (
                max(szs), '/'.join('%.2f' % v for v in dirs.values()))
            break
    _add(results, 5, rule, ok, detail)


    # ---------------------------------------------------------------- 17 / 18 公共判定
    def _h_arrow(recs, x0, x1, ycen, lo, hi):
        """水平箭头：rightArrow 预设形状，或带 <a:headEnd>/<a:tailEnd> 的横向线条。"""
        out = []
        for r in recs:
            if r['w'] <= r['h']:
                continue
            if not (r['geom'] in ('rightArrow', 'leftRightArrow', 'stripedRightArrow')
                    or r['he'] or r['te']):
                continue
            if not (_btw(r['w'], lo, hi) and _in_win(r, x0, x1)):
                continue
            if abs(r['pcy'] - ycen) > 9:
                continue
            out.append(r)
        return out

    # ---------------------------------------------------------------- 17
    # 规则：第1页幻灯片距左侧页边线14%–27%、距顶部边线67%–82%位置出现“Digital health initiatives (T1)”
    #       文本：字体为Arial或主题默认西文字体，加粗，10-13磅，颜色为青色或深蓝色；其右侧距左侧页边线
    #       23%–29%位置出现一条长0.8-1.5cm的水平箭头，箭头指向右方。
    # ①文本内容与位置窗口 ②字体/加粗/10-13磅 ③文字颜色为青色或深蓝色 ④右侧水平箭头位置与长度0.8-1.5cm
    # 与办公软件对齐：读 <a:t>、<a:rPr b/sz> 与其 <a:solidFill>、<a:prstGeom prst="rightArrow"> 或 <a:tailEnd>。
    rule = '第1页幻灯片距左侧页边线14%–27%、距顶部边线67%–82%位置出现“Digital health initiatives (T1)”文本：字体为Arial或主题默认西文字体，加粗，10-13磅，颜色为青色或深蓝色；其右侧距左侧页边线23%–29%位置出现一条长0.8-1.5cm的水平箭头，箭头指向右方'
    ok, detail = False, ''
    for r in _text_recs(s1, 'Digital health initiatives (T1)', 14, 27, 67, 82):
        szs, bold, _, _, fonts = _run_stat(r)
        if not (bold and fonts and szs and all(_btw(s, 10, 13) for s in szs)):
            continue
        if not (_any_col(r, _is_cyan) or _any_col(r, _is_dark_blue)):
            continue
        arr = _h_arrow(s1, 23, 29, r['pcy'], 0.8, 1.5)
        if arr:
            ok = True
            detail = '文本%.1f磅加粗，右侧水平箭头长%.2fcm' % (max(szs), arr[0]['w'])
            break
    _add(results, 3, rule, ok, detail)

    # ---------------------------------------------------------------- 18
    # 规则：第1页幻灯片距左侧页边线78%–91%、距顶部边线67%–83%位置出现“Digital health outcomes (T2)”
    #       文本：字体为Arial或主题默认西文字体，加粗，10-13磅，颜色为青色或深蓝色；其左侧距左侧页边线
    #       74%–81%位置出现一条长0.8-1.5cm的水平箭头，箭头指向右方。
    # ①文本内容与位置窗口 ②字体/加粗/10-13磅 ③文字颜色为青色或深蓝色 ④左侧水平箭头位置与长度0.8-1.5cm
    # 与办公软件对齐：同 17，读 <a:t>/<a:rPr>/<a:solidFill> 与箭头形状的 <a:ext>、箭头端属性。
    rule = '第1页幻灯片距左侧页边线78%–91%、距顶部边线67%–83%位置出现“Digital health outcomes (T2)”文本：字体为Arial或主题默认西文字体，加粗，10-13磅，颜色为青色或深蓝色；其左侧距左侧页边线74%–81%位置出现一条长0.8-1.5cm的水平箭头，箭头指向右方'
    ok, detail = False, ''
    for r in _text_recs(s1, 'Digital health outcomes (T2)', 78, 91, 67, 83):
        szs, bold, _, _, fonts = _run_stat(r)
        if not (bold and fonts and szs and all(_btw(s, 10, 13) for s in szs)):
            continue
        if not (_any_col(r, _is_cyan) or _any_col(r, _is_dark_blue)):
            continue
        arr = _h_arrow(s1, 74, 81, r['pcy'], 0.8, 1.5)
        if arr:
            ok = True
            detail = '文本%.1f磅加粗，左侧水平箭头长%.2fcm' % (max(szs), arr[0]['w'])
            break
    _add(results, 3, rule, ok, detail)


    # ---------------------------------------------------------------- 19
    # 规则：第1页幻灯片距左侧页边线13%–17%位置出现一条竖直向上的坐标轴箭头：长15-18cm；其左侧距左侧
    #       页边线10%–17%、距顶部边线32%–63%位置出现旋转竖排的“System Conditions”文本，字体为Arial
    #       或主题默认西文字体，加粗，13-17磅。
    # ①竖直线条位置窗口 ②长度15-18cm ③线端带箭头 ④文本内容与位置窗口 ⑤文本经旋转竖排 ⑥字体/加粗/13-17磅
    # 与办公软件对齐：读 <a:ext cy>、<a:headEnd/tailEnd type>、<a:xfrm rot>、<a:rPr b/sz>。
    rule = '第1页幻灯片距左侧页边线13%–17%位置出现一条竖直向上的坐标轴箭头：长15-18cm；其左侧距左侧页边线10%–17%、距顶部边线32%–63%位置出现旋转竖排的“System Conditions”文本，字体为Arial或主题默认西文字体，加粗，13-17磅'
    axis_v1 = [r for r in s1 if r['h'] > r['w'] and _in_win(r, 13, 17)
               and _btw(r['h'], 15, 18) and (r['he'] or r['te'])]
    ok, detail = False, ''
    for r in _text_recs(s1, 'System Conditions', 10, 17, 32, 63):
        szs, bold, _, _, fonts = _run_stat(r)
        rotated = (round(r['rot']) % 360) in (90, 180, 270) or r['h'] > r['w']
        if axis_v1 and bold and fonts and rotated and szs and all(_btw(s, 13, 17) for s in szs):
            ok = True
            detail = '坐标轴箭头长%.2fcm，竖排文本旋转%.0f度、%.1f磅加粗' % (
                axis_v1[0]['h'], r['rot'], max(szs))
            break
    _add(results, 3, rule, ok, detail)

    # ---------------------------------------------------------------- 20
    # 规则：第1页幻灯片距顶部边线92%–96%位置出现一条水平向右的坐标轴箭头：长21-25cm；其下方距左侧
    #       页边线43%–58%、距顶部边线94%–100%位置出现“Time (T)”文本，字体为Arial或主题默认西文字体，
    #       加粗，13-17磅。
    # ①水平线条纵向位置窗口 ②长度21-25cm ③线端带箭头 ④文本内容与位置窗口 ⑤字体/加粗/13-17磅
    # 与办公软件对齐：读 <a:ext cx>、<a:tailEnd type>、<a:t>、<a:rPr b/sz>。
    rule = '第1页幻灯片距顶部边线92%–96%位置出现一条水平向右的坐标轴箭头：长21-25cm；其下方距左侧页边线43%–58%、距顶部边线94%–100%位置出现“Time (T)”文本，字体为Arial或主题默认西文字体，加粗，13-17磅'
    axis_h1 = [r for r in s1 if r['w'] > r['h'] and _in_win(r, 0, 100, 92, 96)
               and _btw(r['w'], 21, 25) and (r['he'] or r['te'])]
    ok, detail = False, ''
    for r in _text_recs(s1, 'Time (T)', 43, 58, 94, 100):
        szs, bold, _, _, fonts = _run_stat(r)
        if axis_h1 and bold and fonts and szs and all(_btw(s, 13, 17) for s in szs):
            ok = True
            detail = '坐标轴箭头长%.2fcm，文本%.1f磅加粗' % (axis_h1[0]['w'], max(szs))
            break
    _add(results, 1, rule, ok, detail)

    # ---------------------------------------------------------------- 21
    # 规则：第2页幻灯片距左侧页边线14%–18%位置出现一条竖直向上的坐标轴箭头：长14-17cm；其左侧距左侧
    #       页边线10%–25%位置出现三段文本，距顶部边线1%–13%处为黑色加粗11-15磅的“Industry Maturity”、
    #       11%–20%处为蓝色11-15磅的“High”、85%–94%处为蓝色11-15磅的“Low”；字体均为Arial或主题默认西文字体。
    # ①竖直线条位置与长度14-17cm且带箭头 ②“Industry Maturity”黑色加粗11-15磅及位置
    # ③“High”蓝色11-15磅及位置 ④“Low”蓝色11-15磅及位置 ⑤三段字体均合规
    # 与办公软件对齐：读 <a:ext cy>、<a:tailEnd>、<a:t>、<a:rPr b/sz> 与其 <a:solidFill>（schemeClr 经主题解析）。
    rule = '第2页幻灯片距左侧页边线14%–18%位置出现一条竖直向上的坐标轴箭头：长14-17cm；其左侧距左侧页边线10%–25%位置出现三段文本，距顶部边线1%–13%处为黑色加粗11-15磅的“Industry Maturity”、11%–20%处为蓝色11-15磅的“High”、85%–94%处为蓝色11-15磅的“Low”；字体均为Arial或主题默认西文字体'
    axis_v2 = [r for r in s2 if r['h'] > r['w'] and _in_win(r, 14, 18)
               and _btw(r['h'], 14, 17) and (r['he'] or r['te'])]
    t_im = [r for r in _text_recs(s2, 'Industry Maturity', 10, 25, 1, 13)
            if _sz_in(r, 11, 15) and _run_stat(r)[1] and _run_stat(r)[4] and _any_col(r, _is_black)]
    t_hi = [r for r in _text_recs(s2, 'High', 10, 25, 11, 20)
            if _sz_in(r, 11, 15) and _run_stat(r)[4] and _any_col(r, _is_blue)]
    t_lo = [r for r in _text_recs(s2, 'Low', 10, 25, 85, 94)
            if _sz_in(r, 11, 15) and _run_stat(r)[4] and _any_col(r, _is_blue)]
    ok = bool(axis_v2 and t_im and t_hi and t_lo)
    _add(results, 3, rule, ok, '坐标轴箭头长%.2fcm，三段轴标签字号%.1f/%.1f/%.1f磅' % (
        (axis_v2[0]['h'], _sz_main(t_im[0]), _sz_main(t_hi[0]), _sz_main(t_lo[0])) if ok else (0, 0, 0, 0)))

    # ---------------------------------------------------------------- 22
    # 规则：第2页幻灯片距顶部边线90%–95%位置出现一条水平向右的坐标轴箭头：长22-26cm；其下方距左侧
    #       页边线45%–58%、距顶部边线92%–100%位置出现“Time”文本，字体为Arial或主题默认西文字体，加粗，13-17磅。
    # ①水平线条纵向位置窗口 ②长度22-26cm ③线端带箭头 ④文本内容与位置窗口 ⑤字体/加粗/13-17磅
    # 与办公软件对齐：读 <a:ext cx>、<a:tailEnd type>、<a:t>、<a:rPr b/sz>。
    rule = '第2页幻灯片距顶部边线90%–95%位置出现一条水平向右的坐标轴箭头：长22-26cm；其下方距左侧页边线45%–58%、距顶部边线92%–100%位置出现“Time”文本，字体为Arial或主题默认西文字体，加粗，13-17磅'
    axis_h2 = [r for r in s2 if r['w'] > r['h'] and _in_win(r, 0, 100, 90, 95)
               and _btw(r['w'], 22, 26) and (r['he'] or r['te'])]
    ok, detail = False, ''
    for r in _text_recs(s2, 'Time', 45, 58, 92, 100):
        szs, bold, _, _, fonts = _run_stat(r)
        if axis_h2 and bold and fonts and szs and all(_btw(s, 13, 17) for s in szs):
            ok = True
            detail = '坐标轴箭头长%.2fcm，文本%.1f磅加粗' % (axis_h2[0]['w'], max(szs))
            break
    _add(results, 1, rule, ok, detail)


    # ---------------------------------------------------------------- 23
    # 规则：第2页幻灯片距左侧页边线15%–88%、距顶部边线3%–27%位置出现一条自左下向右上的平滑上升曲线：
    #       整体宽21-24cm，高2.6-3.8cm；线条颜色沿走向由蓝色经青色变为橙色；右端带箭头；曲线上出现三个
    #       圆环形标记点，直径0.4-0.8cm，自左向右依次抬高。
    # ①曲线为平滑曲线（含 <a:custGeom> 自由曲线段，而非纯直线段折线） ②整体宽21-24cm、高2.6-3.8cm
    # ③自左下向右上上升 ④颜色沿走向蓝→青→橙 ⑤右端带箭头 ⑥三个圆环标记直径0.4-0.8cm且自左向右抬高
    # 与办公软件对齐：读 <a:custGeom>/<a:prstGeom prst="line">、<a:ln><a:solidFill>、<a:tailEnd>、
    #                <a:prstGeom prst="ellipse/donut"> + <a:noFill>、<a:off>/<a:ext>。
    rule = '第2页幻灯片距左侧页边线15%–88%、距顶部边线3%–27%位置出现一条自左下向右上的平滑上升曲线：整体宽21-24cm，高2.6-3.8cm；线条颜色沿走向由蓝色经青色变为橙色；右端带箭头；曲线上出现三个圆环形标记点，直径0.4-0.8cm，自左向右依次抬高'
    segs = [r for r in s2 if r['geom'] in ('custGeom', 'line', 'straightConnector1')
            and _in_win(r, 15, 88, 3, 27) and r['h'] <= 1.2 and r['w'] >= 0.3]
    rings = sorted([r for r in s2 if r['geom'] in ('ellipse', 'donut') and _in_win(r, 15, 88, 3, 27)
                    and _btw(r['w'], 0.4, 0.8) and _btw(r['h'], 0.4, 0.8) and abs(r['w'] - r['h']) <= 0.15],
                   key=lambda r: r['pcx'])
    smooth = [r for r in segs if r['geom'] == 'custGeom']
    cw = ch = 0.0
    if segs:
        cw = (max(r['xr'] for r in segs) - min(r['xl'] for r in segs)) / 100.0 * pw_cm
        ch = (max(r['yb'] for r in segs) - min(r['yt'] for r in segs)) / 100.0 * ph_cm
    marks = [(r['pcx'], r['line'] or r['fill']) for r in segs] + [(r['pcx'], r['line'] or r['fill']) for r in rings]
    c_blue = any(x < 42 and _is_blue(c) for x, c in marks)
    c_cyan = any(33 <= x <= 72 and _is_cyan(c) for x, c in marks)
    c_orng = any(x > 58 and _is_orange(c) for x, c in marks)
    rising = False
    if len(segs) >= 2:
        left = min(segs, key=lambda r: r['pcx'])
        right = max(segs, key=lambda r: r['pcx'])
        rising = right['pcy'] < left['pcy'] - 2
    tip = any(r['he'] or r['te'] for r in segs)
    if not tip and segs:
        xr_max = max(r['xr'] for r in segs)
        tip = any(r['fill'] is not None and r['w'] <= 0.9 and r['h'] <= 0.9
                  and r['xr'] >= xr_max - 5 for r in segs + [x for x in s2 if _in_win(x, 15, 88, 3, 27)])
    ring_ok = (len(rings) >= 3
               and all(rings[k + 1]['pcy'] < rings[k]['pcy'] for k in range(len(rings) - 1)))
    ok = bool(len(smooth) >= 2 and _btw(cw, 21, 24) and _btw(ch, 2.6, 3.8) and rising
              and c_blue and c_cyan and c_orng and tip and ring_ok)
    _add(results, 5, rule, ok, '平滑曲线段%d条、整体%.2f×%.2fcm，蓝→青→橙，%d个圆环直径%.2fcm依次抬高' % (
        len(smooth), cw, ch, len(rings), rings[0]['w'] if rings else 0))

    # ---------------------------------------------------------------- 24
    # 规则：第2页幻灯片距顶部边线20%–34%位置出现三个并排的横向色条形状：每条宽5-6.5cm，高1.2-1.9cm，
    #       分别位于距左侧页边线18%–40%、41%–64%、64%–87%范围内；填充自左至右依次为蓝色、青色、橙色；
    #       每条左端各出现一个白色圆环，直径1.2-1.9cm，环内文本依次为“01”、“02”、“03”，字体为Arial或
    #       主题默认西文字体，加粗，18-22磅。
    # ①三条色条位置与宽高 ②填充依次蓝/青/橙 ③每条左端白色圆环且直径1.2-1.9cm ④环内序号文本 ⑤字体/加粗/18-22磅
    # 与办公软件对齐：读 <a:ext cx/cy>、<a:solidFill>（含 schemeClr）、<a:prstGeom prst="donut/ellipse">、<a:rPr b/sz>。
    rule = '第2页幻灯片距顶部边线20%–34%位置出现三个并排的横向色条形状：每条宽5-6.5cm，高1.2-1.9cm，分别位于距左侧页边线18%–40%、41%–64%、64%–87%范围内；填充自左至右依次为蓝色、青色、橙色；每条左端各出现一个白色圆环，直径1.2-1.9cm，环内文本依次为“01”、“02”、“03”，字体为Arial或主题默认西文字体，加粗，18-22磅'
    bar_wins = ((18, 40, _is_blue), (41, 64, _is_cyan), (64, 87, _is_orange))
    nums = ('01', '02', '03')
    bars, ringz, numz = [], [], []
    for k, (bx0, bx1, pred) in enumerate(bar_wins):
        cand = [r for r in s2 if _in_win(r, bx0, bx1, 20, 34) and _btw(r['w'], 5, 6.5)
                and _btw(r['h'], 1.2, 1.9) and pred(r['fill'])]
        bars.append(cand[0] if cand else None)
        rg = [r for r in s2 if r['geom'] in ('donut', 'ellipse') and _is_white(r['fill'])
              and _btw(r['w'], 1.2, 1.9) and _btw(r['h'], 1.2, 1.9) and _in_win(r, bx0 - 3, bx1, 20, 34)
              and (cand and abs(r['pcx'] - cand[0]['xl']) <= 4.5)]
        ringz.append(rg[0] if rg else None)
        nm = [r for r in s2 if r['text'].strip() == nums[k] and _in_win(r, bx0 - 3, bx1, 20, 34)
              and _sz_in(r, 18, 22) and _run_stat(r)[1] and _run_stat(r)[4]]
        numz.append(nm[0] if nm else None)
    ok = all(bars) and all(ringz) and all(numz)
    _add(results, 5, rule, ok, '三条色条宽%s高%s，白色圆环直径%s，序号%s磅加粗' % (
        '/'.join('%.2f' % b['w'] for b in bars if b), '/'.join('%.2f' % b['h'] for b in bars if b),
        '/'.join('%.2f' % r['w'] for r in ringz if r), '/'.join('%.1f' % _sz_main(n) for n in numz if n)))


    # ---------------------------------------------------------------- 25~27/33/34 公共判定
    def _triple(recs, items, wins, y0, y1, szlo, szhi, bold=None, preds=None, exact=False):
        """三段并排文本的统一判定：内容、所在横向窗口、纵向窗口、字号区间、加粗与否、颜色。"""
        got = []
        for k, name in enumerate(items):
            x0, x1 = wins[k]
            hit = None
            key = _norm(name).lower()
            for r in recs:
                if exact:
                    if r['text'].strip().lower() != key:
                        continue
                elif key not in r['text'].lower():
                    continue
                if not (_in_win(r, x0, x1, y0, y1) and _sz_in(r, szlo, szhi) and _run_stat(r)[4]):
                    continue
                b = any(x['b'] for x in _main_runs(r))
                if bold is True and not b:
                    continue
                if bold is False and b:
                    continue
                if preds is not None and not _any_col(r, preds[k]):
                    continue
                hit = r
                break
            got.append(hit)
        return got

    # ---------------------------------------------------------------- 25
    # 规则：第2页幻灯片距顶部边线20%–31%位置出现三段白色标题文本，分别位于距左侧页边线20%–38%、
    #       42%–63%、64%–89%范围内，自左至右依次为“Pilot Exploration Stage”、“Expansion &
    #       Standardization Stage”、“Integrated Optimization Stage”：字体为Arial或主题默认西文字体，加粗，8-13磅。
    # ①三段内容与自左至右分布 ②各自横向窗口 ③纵向窗口 ④白色文字 ⑤字体/加粗/8-13磅
    # 与办公软件对齐：读 <a:t>、<a:rPr b/sz>、<a:solidFill><a:schemeClr val="bg1">（经 clrMap 解析为白色）。
    rule = '第2页幻灯片距顶部边线20%–31%位置出现三段白色标题文本，分别位于距左侧页边线20%–38%、42%–63%、64%–89%范围内，自左至右依次为“Pilot Exploration Stage”、“Expansion & Standardization Stage”、“Integrated Optimization Stage”：字体为Arial或主题默认西文字体，加粗，8-13磅'
    got = _triple(s2, ['Pilot Exploration Stage', 'Expansion & Standardization Stage',
                       'Integrated Optimization Stage'],
                  [(20, 38), (42, 63), (64, 89)], 20, 31, 8, 13, bold=True,
                  preds=[_is_white, _is_white, _is_white])
    ok = all(got)
    _add(results, 1, rule, ok, '三段白色标题字号%s磅加粗' % (
        '/'.join('%.1f' % _sz_main(g) for g in got if g)))

    # ---------------------------------------------------------------- 26
    # 规则：第2页幻灯片距顶部边线24%–34%位置出现三段年份区间文本，分别位于距左侧页边线22%–36%、
    #       45%–66%、68%–85%范围内，自左至右依次为“(1998–2008)”、“(2009–2018)”、“(2019–Present)”：
    #       字体为Arial或主题默认西文字体，加粗，10-14磅；连接符可为连接号“–”或半角连字符“-”。
    # ①三段内容（连接号与半角连字符均可） ②各自横向窗口 ③纵向窗口 ④字体/加粗/10-14磅
    # 与办公软件对齐：读 <a:t>（比对前统一 – — − 为 -）、<a:rPr b/sz>。
    rule = '第2页幻灯片距顶部边线24%–34%位置出现三段年份区间文本，分别位于距左侧页边线22%–36%、45%–66%、68%–85%范围内，自左至右依次为“(1998–2008)”、“(2009–2018)”、“(2019–Present)”：字体为Arial或主题默认西文字体，加粗，10-14磅；连接符可为连接号“–”或半角连字符“-”'
    got = _triple(s2, ['(1998–2008)', '(2009–2018)', '(2019–Present)'],
                  [(22, 36), (45, 66), (68, 85)], 24, 34, 10, 14, bold=True)
    ok = all(got)
    _add(results, 1, rule, ok, '三段年份区间字号%s磅加粗' % (
        '/'.join('%.1f' % _sz_main(g) for g in got if g)))

    # ---------------------------------------------------------------- 27
    # 规则：第2页幻灯片距顶部边线31%–41%位置出现三段居中的“Representative milestones”文本，分别位于
    #       距左侧页边线18%–39%、41%–62%、64%–86%范围内：字体为Arial或主题默认西文字体，9-13磅，
    #       三段颜色自左至右依次为深蓝色、青色、橙色；每段下方各出现一条长5-6.5cm的水平实线。
    # ①三段内容与横向/纵向窗口 ②字体与9-13磅 ③颜色依次深蓝/青/橙 ④每段居中于其下横线
    # ⑤每段下方一条长5-6.5cm的水平实线
    # 与办公软件对齐：读 <a:t>、<a:rPr sz>、<a:solidFill>、<a:prstDash>（缺失即实线）、<a:ext cx>。
    rule = '第2页幻灯片距顶部边线31%–41%位置出现三段居中的“Representative milestones”文本，分别位于距左侧页边线18%–39%、41%–62%、64%–86%范围内：字体为Arial或主题默认西文字体，9-13磅，三段颜色自左至右依次为深蓝色、青色、橙色；每段下方各出现一条长5-6.5cm的水平实线'
    wins27 = [(18, 39), (41, 62), (64, 86)]
    got = _triple(s2, ['Representative milestones'] * 3, wins27, 31, 41, 9, 13,
                  preds=[_is_dark_blue, _is_cyan, _is_orange])
    ok = all(got) and len(set(g['i'] for g in got if g)) == 3
    lines27 = []
    if ok:
        for k, g in enumerate(got):
            cand = [r for r in s2 if r['w'] > r['h'] and _btw(r['w'], 5, 6.5) and r['dash'] == 'solid'
                    and _in_win(r, wins27[k][0] - 2, wins27[k][1] + 2)
                    and g['pcy'] < r['pcy'] <= g['pcy'] + 9
                    and abs(r['pcx'] - g['pcx']) <= 2.0]
            lines27.append(cand[0] if cand else None)
        ok = all(lines27)
    _add(results, 3, rule, ok, '三段文本字号%s磅、颜色深蓝/青/橙，下横线长%s' % (
        '/'.join('%.1f' % _sz_main(g) for g in got if g),
        '/'.join('%.2fcm' % l['w'] for l in lines27 if l)))


    # ---------------------------------------------------------------- 28~30 公共判定
    def _milestones(recs, x0, x1, y0, y1, rows):
        """“年份+事项”行：年份加粗10-14磅、事项不加粗9-13磅，行首彩色圆点直径0.2-0.6cm，且自上而下排列。"""
        got = []
        for year, item in rows:
            yr = None
            for r in recs:
                if r['text'].strip() != year:
                    continue
                if not (_in_win(r, x0, x1, y0, y1) and _sz_in(r, 10, 14) and _run_stat(r)[4]):
                    continue
                if not all(x['b'] for x in _main_runs(r)):
                    continue
                yr = r
                break
            if yr is None:
                return None, ''
            it = None
            for r in _text_recs(recs, item, x0, x1, y0, y1):
                if not (_sz_in(r, 9, 13) and _run_stat(r)[4]):
                    continue
                if any(x['b'] for x in _main_runs(r)):
                    continue
                if abs(r['pcy'] - yr['pcy']) > 6:
                    continue
                it = r
                break
            if it is None:
                return None, ''
            dot = None
            for r in recs:
                if r['geom'] not in ('ellipse', 'donut'):
                    continue
                if not (_btw(r['w'], 0.2, 0.6) and _btw(r['h'], 0.2, 0.6) and _is_colored(r['fill'])):
                    continue
                if r['pcx'] > yr['pcx'] or abs(r['pcy'] - yr['pcy']) > 4:
                    continue
                if not _in_win(r, x0 - 4, x1, y0 - 3, y1 + 3):
                    continue
                dot = r
                break
            if dot is None:
                return None, ''
            got.append((yr, it, dot))
        for k in range(len(got) - 1):
            if got[k][0]['pcy'] >= got[k + 1][0]['pcy']:
                return None, ''
        return got, '%d行年份+事项：年份%s磅加粗、事项%s磅不加粗、行首圆点直径%.2fcm' % (
            len(got), '/'.join('%.1f' % _sz_main(g[0]) for g in got),
            '/'.join('%.1f' % _sz_main(g[1]) for g in got), got[0][2]['w'])

    # ---------------------------------------------------------------- 28
    # 规则：第2页幻灯片距左侧页边线16%–42%、距顶部边线37%–67%位置自上而下出现四行“年份+事项”文本，
    #       依次为“1999 Municipal clean transit study launched”…“2007 Local procurement guideline for
    #       low-emission fleets”：年份为加粗10-14磅、事项为不加粗9-13磅，字体均为Arial或主题默认西文字体；
    #       每行行首各出现一个直径0.2-0.6cm的彩色圆点。
    # ①四行内容齐备且自上而下 ②位置窗口 ③年份加粗10-14磅 ④事项不加粗9-13磅 ⑤行首彩色圆点直径0.2-0.6cm
    # 与办公软件对齐：读 <a:t>、<a:rPr b/sz/latin>、<a:prstGeom prst="ellipse">+<a:ext>+填充色。
    rule = '第2页幻灯片距左侧页边线16%–42%、距顶部边线37%–67%位置自上而下出现四行“年份+事项”文本，依次为“1999 Municipal clean transit study launched”、“2002 First battery-bus field trials”、“2005 Depot charging demonstration corridor”、“2007 Local procurement guideline for low-emission fleets”：年份为加粗10-14磅、事项为不加粗9-13磅，字体均为Arial或主题默认西文字体；每行行首各出现一个直径0.2-0.6cm的彩色圆点'
    got, detail = _milestones(s2, 16, 42, 37, 67, [
        ('1999', 'Municipal clean transit study launched'),
        ('2002', 'First battery-bus field trials'),
        ('2005', 'Depot charging demonstration corridor'),
        ('2007', 'Local procurement guideline for low-emission fleets')])
    _add(results, 5, rule, got is not None, detail)

    # ---------------------------------------------------------------- 29
    # 规则：第2页幻灯片距左侧页边线38%–65%、距顶部边线37%–72%位置自上而下出现五行“年份+事项”文本，
    #       依次为“2009 Fleet subsidy mechanism introduced”…“2018 Operations monitoring platform connected
    #       citywide”：年份为加粗10-14磅、事项为不加粗9-13磅，字体均为Arial或主题默认西文字体；每行行首
    #       各出现一个直径0.2-0.6cm的彩色圆点。
    # ①五行内容齐备且自上而下 ②位置窗口 ③年份加粗10-14磅 ④事项不加粗9-13磅 ⑤行首彩色圆点直径0.2-0.6cm
    # 与办公软件对齐：同 28，读 <a:t>/<a:rPr>/<a:ext>/填充色。
    rule = '第2页幻灯片距左侧页边线38%–65%、距顶部边线37%–72%位置自上而下出现五行“年份+事项”文本，依次为“2009 Fleet subsidy mechanism introduced”、“2011 Standard charging interface adopted”、“2013 Public-private depot investment model”、“2016 Route electrification targets released”、“2018 Operations monitoring platform connected citywide”：年份为加粗10-14磅、事项为不加粗9-13磅，字体均为Arial或主题默认西文字体；每行行首各出现一个直径0.2-0.6cm的彩色圆点'
    got, detail = _milestones(s2, 38, 65, 37, 72, [
        ('2009', 'Fleet subsidy mechanism introduced'),
        ('2011', 'Standard charging interface adopted'),
        ('2013', 'Public-private depot investment model'),
        ('2016', 'Route electrification targets released'),
        ('2018', 'Operations monitoring platform connected citywide')])
    _add(results, 5, rule, got is not None, detail)

    # ---------------------------------------------------------------- 30
    # 规则：第2页幻灯片距左侧页边线62%–89%、距顶部边线37%–72%位置自上而下出现五行“年份+事项”文本，
    #       依次为“2019 Battery lifecycle management rules”…“2025 Smart maintenance and second-life storage
    #       programs”：年份为加粗10-14磅、事项为不加粗9-13磅，字体均为Arial或主题默认西文字体；每行行首
    #       各出现一个直径0.2-0.6cm的彩色圆点。
    # ①五行内容齐备且自上而下 ②位置窗口 ③年份加粗10-14磅 ④事项不加粗9-13磅 ⑤行首彩色圆点直径0.2-0.6cm
    # 与办公软件对齐：同 28，读 <a:t>/<a:rPr>/<a:ext>/填充色。
    rule = '第2页幻灯片距左侧页边线62%–89%、距顶部边线37%–72%位置自上而下出现五行“年份+事项”文本，依次为“2019 Battery lifecycle management rules”、“2021 Renewable-powered charging pilots”、“2023 Data-driven dispatch optimization”、“2024 Resilience plan for winter operations”、“2025 Smart maintenance and second-life storage programs”：年份为加粗10-14磅、事项为不加粗9-13磅，字体均为Arial或主题默认西文字体；每行行首各出现一个直径0.2-0.6cm的彩色圆点'
    got, detail = _milestones(s2, 62, 89, 37, 72, [
        ('2019', 'Battery lifecycle management rules'),
        ('2021', 'Renewable-powered charging pilots'),
        ('2023', 'Data-driven dispatch optimization'),
        ('2024', 'Resilience plan for winter operations'),
        ('2025', 'Smart maintenance and second-life storage programs')])
    _add(results, 5, rule, got is not None, detail)


    # ---------------------------------------------------------------- 31
    # 规则：第2页幻灯片距顶部边线30%–92%位置出现三个并排的卡片状圆角矩形形状：每个宽6.2-7.6cm，
    #       高10-11.5cm，分别位于距左侧页边线16%–39%、38%–63%、62%–87%范围内；边线为浅灰色或浅色实线，
    #       内部为白色或极浅色填充。
    # ①三个圆角矩形并排 ②每个宽6.2-7.6cm、高10-11.5cm（按旋转后可视外接框度量） ③各自横向窗口与纵向窗口
    # ④边线为浅色实线 ⑤填充为白色或极浅色
    # 与办公软件对齐：读 <a:prstGeom prst="roundRect">、<a:xfrm rot>+<a:ext>、<a:ln><a:solidFill>+<a:prstDash>、<a:solidFill>。
    rule = '第2页幻灯片距顶部边线30%–92%位置出现三个并排的卡片状圆角矩形形状：每个宽6.2-7.6cm，高10-11.5cm，分别位于距左侧页边线16%–39%、38%–63%、62%–87%范围内；边线为浅灰色或浅色实线，内部为白色或极浅色填充'
    wins31 = [(16, 39), (38, 63), (62, 87)]
    cards = []
    for x0, x1 in wins31:
        cand = [r for r in s2 if r['geom'] == 'roundRect' and _in_win(r, x0, x1, 30, 92)
                and _btw(r['w'], 6.2, 7.6) and _btw(r['h'], 10, 11.5)
                and _is_verylight(r['fill']) and r['dash'] == 'solid'
                and r['line'] is not None and _lum(r['line']) >= 60]
        cards.append(cand[0] if cand else None)
    ok = all(cards)
    _add(results, 3, rule, ok, '三个卡片宽%s高%s，浅色实线边框、极浅色填充' % (
        '/'.join('%.2f' % c['w'] for c in cards if c), '/'.join('%.2f' % c['h'] for c in cards if c)))

    # ---------------------------------------------------------------- 32
    # 规则：第2页幻灯片距顶部边线68%–90%位置出现三个并排的浅色圆角矩形形状：每个宽5.5-7cm，高2.6-3.5cm，
    #       分别位于距左侧页边线16%–39%、39%–62%、62%–87%范围内；每个框左部各出现一个圆形底盘，
    #       直径1.8-2.6cm，盘内由多个形状拼合的图形或图片自左至右依次为文件板加齿轮、充电桩、带图表的显示器。
    # ①三个圆角矩形位置与宽高 ②每框左部一个圆形底盘 ③底盘直径1.8-2.6cm ④盘内图案由多个形状拼合或为图片
    # 与办公软件对齐：读 <a:prstGeom prst="roundRect"/"ellipse">、<a:ext>、以及落在底盘外接框内的无文字形状数量。
    rule = '第2页幻灯片距顶部边线68%–90%位置出现三个并排的浅色圆角矩形形状：每个宽5.5-7cm，高2.6-3.5cm，分别位于距左侧页边线16%–39%、39%–62%、62%–87%范围内；每个框左部各出现一个圆形底盘，直径1.8-2.6cm，盘内由多个形状拼合的图形或图片自左至右依次为文件板加齿轮、充电桩、带图表的显示器'
    wins32 = [(16, 39), (39, 62), (62, 87)]
    panels, plates, icons = [], [], []
    for x0, x1 in wins32:
        cand = [r for r in s2 if r['geom'] == 'roundRect' and _in_win(r, x0, x1, 68, 90)
                and _btw(r['w'], 5.5, 7) and _btw(r['h'], 2.6, 3.5)
                and (r['fill'] is None or _is_verylight(r['fill']))]
        p = cand[0] if cand else None
        panels.append(p)
        pl = None
        if p is not None:
            for r in s2:
                if r['geom'] not in ('ellipse', 'donut'):
                    continue
                if not (_btw(r['w'], 1.8, 2.6) and _btw(r['h'], 1.8, 2.6) and abs(r['w'] - r['h']) <= 0.5):
                    continue
                if not (p['xl'] - 1 <= r['pcx'] <= p['pcx'] and p['yt'] - 2 <= r['pcy'] <= p['yb'] + 2):
                    continue
                pl = r
                break
        plates.append(pl)
        icons.append(_cluster(s2, pl['xl'] - 0.6, pl['xr'] + 0.6, pl['yt'] - 0.6, pl['yb'] + 0.6,
                              min_n=2, exclude=(pl['i'],)) if pl is not None else None)
    ok = all(panels) and all(plates) and all(icons)
    _add(results, 5, rule, ok, '三个浅色圆角矩形宽%s，圆形底盘直径%s，盘内拼合形状%s个' % (
        '/'.join('%.2fcm' % p['w'] for p in panels if p),
        '/'.join('%.2fcm' % p['w'] for p in plates if p),
        '/'.join('%d' % c[0] for c in icons if c)))

    # ---------------------------------------------------------------- 33
    # 规则：第2页幻灯片距顶部边线70%–79%位置出现三段“Stage profile”文本，分别位于距左侧页边线
    #       23%–37%、45%–61%、69%–85%范围内：字体为Arial或主题默认西文字体，加粗，10-14磅，
    #       颜色自左至右依次为深蓝色、青色、橙色。
    # ①三段内容与横向窗口 ②纵向窗口 ③字体/加粗/10-14磅 ④颜色依次深蓝/青/橙
    # 与办公软件对齐：读 <a:t>、<a:rPr b/sz>、<a:solidFill>（含 schemeClr 经主题与 lumMod 解析）。
    rule = '第2页幻灯片距顶部边线70%–79%位置出现三段“Stage profile”文本，分别位于距左侧页边线23%–37%、45%–61%、69%–85%范围内：字体为Arial或主题默认西文字体，加粗，10-14磅，颜色自左至右依次为深蓝色、青色、橙色'
    got = _triple(s2, ['Stage profile'] * 3, [(23, 37), (45, 61), (69, 85)], 70, 79, 10, 14,
                  bold=True, preds=[_is_dark_blue, _is_cyan, _is_orange])
    ok = all(got) and len(set(g['i'] for g in got if g)) == 3
    _add(results, 1, rule, ok, '三段“Stage profile”字号%s磅加粗、颜色深蓝/青/橙' % (
        '/'.join('%.1f' % _sz_main(g) for g in got if g)))

    # ---------------------------------------------------------------- 34
    # 规则：第2页幻灯片距顶部边线74%–87%位置出现三段描述文本，分别位于距左侧页边线23%–39%、45%–62%、
    #       69%–89%范围内，自左至右依次为“technology testing, limited routes, institutional learning”、
    #       “policy support, scale-up, infrastructure coordination”、“market maturity, efficiency optimization,
    #       system integration”：字体为Arial或主题默认西文字体，不加粗，8-12磅。
    # ①三段内容与横向窗口 ②纵向窗口 ③字体 ④不加粗 ⑤8-12磅
    # 与办公软件对齐：读 <a:t>（空白折叠后比对）、<a:rPr b/sz/latin>。
    rule = '第2页幻灯片距顶部边线74%–87%位置出现三段描述文本，分别位于距左侧页边线23%–39%、45%–62%、69%–89%范围内，自左至右依次为“technology testing, limited routes, institutional learning”、“policy support, scale-up, infrastructure coordination”、“market maturity, efficiency optimization, system integration”：字体为Arial或主题默认西文字体，不加粗，8-12磅'
    got = _triple(s2, ['technology testing, limited routes, institutional learning',
                       'policy support, scale-up, infrastructure coordination',
                       'market maturity, efficiency optimization, system integration'],
                  [(23, 39), (45, 62), (69, 89)], 74, 87, 8, 12, bold=False)
    ok = all(got)
    _add(results, 1, rule, ok, '三段描述文本字号%s磅、不加粗' % (
        '/'.join('%.1f' % _sz_main(g) for g in got if g)))


    # ---------------------------------------------------------------- 35
    # 规则：第2页幻灯片距顶部边线51%–61%位置出现两个圆形徽标形状：分别位于距左侧页边线35%–43%与
    #       58%–67%范围内；直径0.7-1.2cm；徽标内出现白色的向右尖括号图形。
    # ①两个圆形位置窗口 ②直径0.7-1.2cm ③徽标内的尖括号为图形（由线条/形状拼合），而非文字字符
    # ④尖括号为白色
    # 与办公软件对齐：读 <a:prstGeom prst="ellipse">+<a:ext>，并统计落在其外接框内的无文字线条形状及其 <a:ln><a:solidFill>。
    rule = '第2页幻灯片距顶部边线51%–61%位置出现两个圆形徽标形状：分别位于距左侧页边线35%–43%与58%–67%范围内；直径0.7-1.2cm；徽标内出现白色的向右尖括号图形'
    badges, chevrons = [], []
    for x0, x1 in ((35, 43), (58, 67)):
        cand = [r for r in s2 if r['geom'] == 'ellipse' and _in_win(r, x0, x1, 51, 61)
                and _btw(r['w'], 0.7, 1.2) and _btw(r['h'], 0.7, 1.2)]
        b = cand[0] if cand else None
        badges.append(b)
        if b is None:
            chevrons.append(None)
            continue
        mem = [r for r in s2 if r['i'] != b['i'] and not r['has_text']
               and _inside(r, b['xl'] - 0.6, b['xr'] + 0.6, b['yt'] - 0.6, b['yb'] + 0.6)
               and (_is_white(r['line']) or _is_white(r['fill']))]
        chevrons.append(mem if len(mem) >= 2 else None)
    ok = all(badges) and all(chevrons)
    _add(results, 3, rule, ok, '两个圆形徽标直径%s，内含白色尖括号图形（%s个形状拼合）' % (
        '/'.join('%.2fcm' % b['w'] for b in badges if b),
        '/'.join('%d' % len(c) for c in chevrons if c)))

    # ---------------------------------------------------------------- 36
    # 规则：第2页幻灯片距左侧页边线36%–42%与60%–65%位置各出现一条竖直虚线：每条长11-14cm。
    # ①两条竖直线条分别位于两个横向窗口 ②线型为虚线（含点线/点划线等非实线虚化线型） ③每条长11-14cm
    # 与办公软件对齐：读 <a:prstDash val>（非 solid）与 <a:ext cy>。
    rule = '第2页幻灯片距左侧页边线36%–42%与60%–65%位置各出现一条竖直虚线：每条长11-14cm'
    vd = []
    for x0, x1 in ((36, 42), (60, 65)):
        cand = [r for r in s2 if r['h'] > r['w'] and _in_win(r, x0, x1)
                and _btw(r['h'], 11, 14) and r['dash'] != 'solid']
        vd.append(cand[0] if cand else None)
    ok = all(vd)
    _add(results, 3, rule, ok, '两条竖直%s线长%s' % (
        vd[0]['dash'] if ok else '', '/'.join('%.2fcm' % r['h'] for r in vd if r)))

    # ---------------------------------------------------------------- 37
    # 规则：幻灯片页面尺寸为宽33.87cm×高19.05cm，备注页尺寸为宽19.05cm×高25.4cm。
    # ①页面宽33.87cm ②页面高19.05cm ③备注页宽19.05cm ④备注页高25.4cm（容差0.05cm）
    # 与办公软件对齐：读 ppt/presentation.xml 的 <p:sldSz cx/cy> 与 <p:notesSz cx/cy>（EMU→cm）。
    rule = '幻灯片页面尺寸为宽33.87cm×高19.05cm，备注页尺寸为宽19.05cm×高25.4cm'
    nw, nh = _notes_size_cm(file_path)
    ok = (abs(pw_cm - _PAGE_W_CM) <= 0.05 and abs(ph_cm - _PAGE_H_CM) <= 0.05
          and nw is not None and abs(nw - _NOTES_W_CM) <= 0.05 and abs(nh - _NOTES_H_CM) <= 0.05)
    _add(results, 1, rule, ok, '页面%.2f×%.2fcm，备注页%.2f×%.2fcm' % (
        pw_cm, ph_cm, nw or 0, nh or 0))

    # ---------------------------------------------------------------- 38
    # 规则：两页幻灯片上不出现空白占位符框，也不出现既无文字又无填充与边线的空白文本框。
    # ①无内容为空的占位符（<p:ph>） ②无“无文字且无填充且无边线”的文本框（<p:cNvSpPr txBox="1">）
    # 与办公软件对齐：读 <p:nvSpPr><p:nvPr><p:ph>、<p:cNvSpPr txBox>、<a:solidFill>/<a:noFill>、<a:ln>。
    rule = '两页幻灯片上不出现空白占位符框，也不出现既无文字又无填充与边线的空白文本框'
    bad_ph, bad_tb = 0, 0
    for recs in pages:
        for r in recs:
            if r['el'].find('.//' + _P + 'ph') is not None and not r['has_text']:
                bad_ph += 1
                continue
            cnv = r['el'].find('.//' + _P + 'cNvSpPr')
            is_txbox = cnv is not None and cnv.get('txBox') == '1'
            if is_txbox and not r['has_text'] and r['fill'] is None and _ln_of(r['el']) is None:
                bad_tb += 1
    ok = (bad_ph == 0 and bad_tb == 0)
    _add(results, 1, rule, ok, '无空白占位符框，无空白文本框')


    # ---------------------------------------------------------------- 39
    # 规则：两页幻灯片的文本字号均为0.5磅的整数倍，不出现9.6、10.4、10.8、12.3磅之类的小数磅值。
    # ①遍历两页所有 run 的字号 ②每个字号乘以2后为整数（即 0.5 磅的整数倍）
    # 与办公软件对齐：读 <a:rPr sz>（单位 1/100 磅），sz 必须是 50 的整数倍。
    rule = '两页幻灯片的文本字号均为0.5磅的整数倍，不出现9.6、10.4、10.8、12.3磅之类的小数磅值'
    odd_sz = []
    for recs in pages:
        for r in recs:
            for run in r['runs']:
                if not run['explicit_sz']:
                    continue
                if abs(run['sz'] * 2 - round(run['sz'] * 2)) > 1e-6:
                    odd_sz.append(round(run['sz'], 2))
    ok = not odd_sz
    _add(results, 3, rule, ok, '全部字号均为0.5磅的整数倍')

    # ---------------------------------------------------------------- 40
    # 规则：两页幻灯片均不出现宽度或高度为负值的形状，文件打开时不提示修复，形状无反向绘制。
    # ①<a:ext> 的 cx/cy 均为非负 ②XML 可被正常解析（可正常打开、不提示修复） ③无负向延伸导致的反向绘制
    # 与办公软件对齐：扫描 ppt/slides/slideN.xml 全部 <a:ext cx cy>，出现负值即判不通过。
    rule = '两页幻灯片均不出现宽度或高度为负值的形状，文件打开时不提示修复，形状无反向绘制'
    neg = 0
    for xml in xmls:
        neg += len(re.findall(r'<a:ext[^>]*?c[xy]="-\d+"', xml))
    ok = (neg == 0 and len(xmls) == 2)
    _add(results, 5, rule, ok, '无负值宽高，形状无反向绘制')

    # ---------------------------------------------------------------- 41
    # 规则：两页幻灯片的文本框均未启用“缩放文字以适应形状”的自动调整，文字不超出所属形状可视区域、
    #       不被相邻形状遮挡或裁切，形状与文字不超出页面边界0.5cm以上。
    # ①所有 <a:bodyPr> 下均无 <a:normAutofit>（即未启用“缩放文字以适应形状”）
    # ②形状与文字的可视外接框未超出页面边界0.5cm以上
    # 与办公软件对齐：扫描 <a:bodyPr><a:normAutofit fontScale/lnSpcReduction>，并按 <a:off>/<a:ext>+rot 计算可视框与页面边界比较。
    rule = '两页幻灯片的文本框均未启用“缩放文字以适应形状”的自动调整，文字不超出所属形状可视区域、不被相邻形状遮挡或裁切，形状与文字不超出页面边界0.5cm以上'
    naf = sum(xml.count('normAutofit') for xml in xmls)
    out_of_page = []
    for recs in pages:
        for r in recs:
            if (r['l'] < -0.5 or r['t'] < -0.5
                    or r['l'] + r['w'] > pw_cm + 0.5 or r['t'] + r['h'] > ph_cm + 0.5):
                out_of_page.append(r['i'])
    ok = (naf == 0 and not out_of_page)
    _add(results, 5, rule, ok, '未启用缩放文字以适应形状，形状与文字未超出页面边界0.5cm')

    # ========== 将 results 反解为结构化 dim2_items ==========
    # 逐项按追加顺序对齐 _MAX_DELTAS_ORDER：解析每条 "+N：说明" 的 N 与 rule。
    if len(results) != len(_MAX_DELTAS_ORDER):
        # 顺序表与实际条目不一致：视为脚本内部错误，交给外层兜底
        raise RuntimeError(
            f"dim2 项数不一致：results={len(results)} vs order={len(_MAX_DELTAS_ORDER)}"
        )
    dim2_items = []
    total_score = 0
    for raw, max_delta in zip(results, _MAX_DELTAS_ORDER):
        # 格式约定：'+N：' 前缀 + 规则说明；命中 N=max_delta，未命中 N=0
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
    # 注：evaluate() 的接口约定就是“入参为目录路径”——路径始终由外部显式传入；
    #     这里的默认值仅为本地调试便利，不影响 evaluate 的接口约定。
    if len(sys.argv) >= 2:
        _dir = sys.argv[1]
    else:
        _dir = os.path.dirname(os.path.abspath(__file__))
    _out = evaluate(_dir)
    print(json.dumps(_out, ensure_ascii=False, indent=2))
