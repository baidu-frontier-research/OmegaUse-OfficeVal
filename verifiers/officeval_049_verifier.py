#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# pyright: reportUnknownMemberType=false, reportUnknownVariableType=false, reportUnknownArgumentType=false, reportUnknownParameterType=false, reportMissingParameterType=false, reportPrivateUsage=false, reportUnknownLambdaType=false, reportGeneralTypeIssues=false
"""
PPT 自动评估脚本（officeval_049）

对外接口：
    evaluate(dir_path) -> dict
        入参为“脚本所在目录的路径”，脚本自己在该目录内定位并打开被评估的
        .pptx，返回结构化评估结果字典（维度一是否通过、维度二逐项得分、总分）。

说明：按接口统一约定，本模块不修改全局 sys.stdout，也不 print 主结果。
主结果统一由 evaluate(dir_path) -> dict 返回；本文件仅在 __main__ 下作为
本地调试入口，才把 dict 序列化为 JSON 打印。

返回字段：id / file_name / status / error / dim1_pass / dim1_reason /
          dim2_items（每项只含 rule / max_delta / delta / hit）/
          total_score / max_score
status='error' 表示这份提交没被真正评过（目录不存在、目录内无演示文稿、脚本
自身异常），此时 dim1_pass 为空串（维度一未判定）、dim2_items 为空表、
total_score 与 max_score 一律为 0；dim1_pass 为 False 时 dim1_reason 必须给出
错误摘要，通过时为空串。

判定口径与办公软件一致：所有坐标、尺寸均换算为幻灯片页面上的绝对位置
（组合内的元素按组合变换递归换算为绝对坐标），单位 cm；线宽单位磅；
颜色只认显式的 RGB 值，未显式指定颜色（沿用主题色/主题字体）的一律按
“沿用主题”放过，不倒扣。
"""
import os
import re
import sys
import json
import math
import zipfile
from xml.etree import ElementTree as ET

# ─── OOXML 命名空间 ─────────────────────────────────────────────────────────
_NS = {
    'p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
    'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
}


def _q(name):
    """'p:spPr' → '{namespace}spPr'，用于 ElementTree 的限定名查找。"""
    prefix, local = name.split(':')
    return '{%s}%s' % (_NS[prefix], local)


# ─── 单位换算 ───────────────────────────────────────────────────────────────
EMU_PER_CM = 360000.0     # 1cm = 360000 EMU
EMU_PER_PT = 12700.0      # 1磅 = 12700 EMU（线宽 a:ln/@w 的单位）

SCRIPT_ID = "049"
# 维度二所有评分项 max_delta 之和；用于维度一未通过 / 脚本异常时兜底
MAX_SCORE_TOTAL = 98

# 维度二所有评分项按“结果追加顺序”的 max_delta 列表；末端把 results 里的
# "+N：..." 字符串反解为结构化 dim2_items 时用。顺序须与 _evaluate_impl 中
# results.append(...) 的顺序严格一致；共 34 项，累加 = MAX_SCORE_TOTAL。
_MAX_DELTAS_ORDER = [
    3,  # 01 页面尺寸与形状总数
    3,  # 02 主标题 RIVERDALE TUNNEL – SECTION 7 MONITORING LAYOUT
    1,  # 03 副标题 Displacement Sensors (mm)
    3,  # 04 纵轴标题 Vertical Position (m) 竖向排列
    1,  # 05 横轴标题 Lateral Position (m)
    3,  # 06 竖直实轴线
    3,  # 07 水平实轴线
    5,  # 08 纵轴 7 个刻度值
    1,  # 09 纵轴刻度正负号齐全
    3,  # 10 横轴 5 个刻度值
    1,  # 11 坐标轴刻度短线
    3,  # 12 水平虚线（中线）
    3,  # 13 竖直虚线（中线）
    5,  # 14 绿色网线数量与整体范围
    5,  # 15 绿色网线为自由曲线
    3,  # 16 绿色网线线宽与颜色统一
    3,  # 17 绿色网线倾角分布
    3,  # 18 绿色网线长弧线连续未拆段
    3,  # 19 绿色网线首尾相连成网
    3,  # 20 中部 11 个蓝色圆点
    3,  # 21 11 个蓝点编号 B114–B125
    3,  # 22 上排 8 个黄色圆点
    3,  # 23 下排 11 个黄色圆点
    3,  # 24 上排 8 个黄点编号 K908–K679
    3,  # 25 下排 11 个黄点编号 A201–A314
    3,  # 26 左下比例尺圆角矩形框
    3,  # 27 左下比例尺一横一竖双箭头
    1,  # 28 比例尺 Y (Up) / X (Right) 标注
    1,  # 29 比例尺 Displacement Scale: 1 mm 说明
    3,  # 30 右下图例圆角矩形框
    3,  # 31 右下图例蓝、黄两个圆点
    3,  # 32 右下图例两行说明文字
    5,  # 33 可编辑性：无图片/艺术字/嵌入对象，全部由形状与文本框构成
    3,  # 34 交付规范：正尺寸、整数字号、无翻转、白底、无效果
]
assert sum(_MAX_DELTAS_ORDER) == MAX_SCORE_TOTAL

# ════════════════════════════════════════════════════════════════════════════
# PPTX 解析：把幻灯片上的每个形状/文本框摊平成绝对坐标的属性字典
# ════════════════════════════════════════════════════════════════════════════
def _color_of(parent):
    """解析 a:ln 或 a:spPr / a:rPr 下的颜色，返回 (种类, (r,g,b))：
       ('rgb',(r,g,b))  显式 srgbClr；('theme',None) 主题色/其它色彩空间；
       ('noFill',None)  显式无填充/无线条；('none',None) 未设定（继承）。"""
    if parent is None:
        return ('none', None)
    if parent.find(_q('a:noFill')) is not None:
        return ('noFill', None)
    sf = parent.find(_q('a:solidFill'))
    if sf is None:
        return ('none', None)
    srgb = sf.find(_q('a:srgbClr'))
    if srgb is not None:
        val = srgb.get('val') or ''
        try:
            return ('rgb', (int(val[0:2], 16), int(val[2:4], 16), int(val[4:6], 16)))
        except ValueError:
            return ('theme', None)
    return ('theme', None)


def _record_shape(ch, tf, kind, out):
    """把一个 p:sp / p:cxnSp / p:pic 元素解析为属性字典追加到 out。
    tf = (sx, sy, dx, dy) 为祖先组合累积的仿射变换（缩放 + 平移，EMU 口径）。"""
    spPr = ch.find('./' + _q('p:spPr'))
    if spPr is None:
        return
    xfrm = spPr.find(_q('a:xfrm'))
    if xfrm is None:
        return
    off = xfrm.find(_q('a:off'))
    ext = xfrm.find(_q('a:ext'))
    # 注意：a:extLst 下也有名为 ext 的元素，但它只带 uri，没有 cx/cy，须排除
    if off is None or ext is None or ext.get('cx') is None or ext.get('cy') is None:
        return
    X = (tf[2] + int(off.get('x')) * tf[0]) / EMU_PER_CM
    Y = (tf[3] + int(off.get('y')) * tf[1]) / EMU_PER_CM
    W = int(ext.get('cx')) * tf[0] / EMU_PER_CM
    H = int(ext.get('cy')) * tf[1] / EMU_PER_CM
    rot = (int(xfrm.get('rot') or 0) / 60000.0) % 360
    flip = (xfrm.get('flipH') == '1' or xfrm.get('flipV') == '1')

    prst = spPr.find(_q('a:prstGeom'))
    cust = spPr.find(_q('a:custGeom'))
    geo = prst.get('prst') if prst is not None else ('cust' if cust is not None else None)
    # 含曲线控制点（三次/二次贝塞尔）即视为自由曲线形状
    curved = cust is not None and (cust.find('.//' + _q('a:cubicBezTo')) is not None
                                   or cust.find('.//' + _q('a:quadBezTo')) is not None)

    ln = spPr.find(_q('a:ln'))
    line_w = (int(ln.get('w')) / EMU_PER_PT) if (ln is not None and ln.get('w')) else None
    line_kind, line_rgb = _color_of(ln)
    dash = head_end = tail_end = None
    if ln is not None:
        d = ln.find(_q('a:prstDash'))
        if d is not None:
            dash = d.get('val')
        h = ln.find(_q('a:headEnd'))
        if h is not None:
            head_end = h.get('type')
        t = ln.find(_q('a:tailEnd'))
        if t is not None:
            tail_end = t.get('type')
    fill_kind, fill_rgb = _color_of(spPr)

    # 形状在页面上的两个端点（用于“绿色网线首尾相连”的连通性判定）：
    #   直线/直线连接符取外框对角两端（按 flipH/flipV 绕中心镜像）；
    #   自由曲线取其路径的起点与终点，闭合路径视为起点与终点重合。
    ends = None
    if geo in ('line', 'straightConnector1'):
        a = [X, Y]
        b = [X + W, Y + H]
        mx = (a[0] + b[0]) / 2.0
        my = (a[1] + b[1]) / 2.0
        if xfrm.get('flipH') == '1':
            a[0] = 2 * mx - a[0]
            b[0] = 2 * mx - b[0]
        if xfrm.get('flipV') == '1':
            a[1] = 2 * my - a[1]
            b[1] = 2 * my - b[1]
        ends = (tuple(a), tuple(b))
    elif cust is not None:
        path = cust.find(_q('a:pathLst') + '/' + _q('a:path'))
        if path is not None:
            pw_ = int(path.get('w') or 0) or 1
            ph_ = int(path.get('h') or 0) or 1
            bx0, bx1 = min(X, X + W), max(X, X + W)
            by0, by1 = min(Y, Y + H), max(Y, Y + H)
            sx = (bx1 - bx0) / pw_
            sy = (by1 - by0) / ph_
            fh = xfrm.get('flipH') == '1'
            fv = xfrm.get('flipV') == '1'

            def _to_page(px, py):
                u = px * sx
                v = py * sy
                if fh:
                    u = (bx1 - bx0) - u
                if fv:
                    v = (by1 - by0) - v
                return (bx0 + u, by0 + v)

            pts = [_to_page(int(p.get('x')), int(p.get('y')))
                   for p in path.findall('.//' + _q('a:pt'))]
            if pts:
                closed = path.find(_q('a:close')) is not None
                ends = (pts[0], pts[0] if closed else pts[-1])

    effect_lst = spPr.find(_q('a:effectLst'))
    has_effect = effect_lst is not None and len(effect_lst) > 0

    txBody = ch.find('./' + _q('p:txBody'))
    text = ''
    runs = []
    autofit = None
    vert = None
    if txBody is not None:
        text = ''.join(t.text or '' for t in txBody.iter(_q('a:t')))
        bodyPr = txBody.find(_q('a:bodyPr'))
        if bodyPr is not None:
            vert = bodyPr.get('vert')          # 非空即文字竖向排列
            autofit = ('norm' if bodyPr.find(_q('a:normAutofit')) is not None else
                       'sp' if bodyPr.find(_q('a:spAutoFit')) is not None else 'none')
        for rPr in txBody.iter(_q('a:rPr')):
            ck, cv = _color_of(rPr)
            latin = rPr.find(_q('a:latin'))
            runs.append(dict(
                sz=(int(rPr.get('sz')) / 100.0 if rPr.get('sz') else None),
                b=(rPr.get('b') == '1'),
                font=(latin.get('typeface') if latin is not None else None),
                ck=ck, cv=cv))

    out.append(dict(
        kind=kind,
        x=min(X, X + W), y=min(Y, Y + H),   # 外框左上角（已归一，负尺寸也能定位）
        w=abs(W), h=abs(H),                 # 外框宽高（绝对值）
        sw=W, sh=H,                         # 带符号的外框宽高（负值即畸形尺寸）
        geo=geo, rot=rot, flip=flip, curved=curved,
        lw=line_w, lk=line_kind, lc=line_rgb, dash=dash,
        fk=fill_kind, fc=fill_rgb, eff=has_effect,
        text=text, runs=runs, af=autofit, vert=vert,
        ends=ends, hd=head_end, tl=tail_end))


def _walk(el, tf, out):
    """递归遍历形状树；遇到组合就把组合的 off/ext 与 chOff/chExt 换算成
    子元素的缩放平移，使组合内的元素也得到页面绝对坐标。"""
    for ch in el:
        name = ch.tag.split('}')[-1]
        if name == 'grpSp':
            xfrm = ch.find('./' + _q('p:grpSpPr') + '/' + _q('a:xfrm'))
            if xfrm is None:
                _walk(ch, tf, out)
                continue
            off = xfrm.find(_q('a:off'))
            ext = xfrm.find(_q('a:ext'))
            ch_off = xfrm.find(_q('a:chOff'))
            ch_ext = xfrm.find(_q('a:chExt'))
            if off is None or ext is None or ch_off is None or ch_ext is None:
                _walk(ch, tf, out)
                continue
            gcx, gcy = int(ext.get('cx')), int(ext.get('cy'))
            ccx, ccy = int(ch_ext.get('cx')), int(ch_ext.get('cy'))
            sx = tf[0] * (gcx / ccx if ccx else 1)
            sy = tf[1] * (gcy / ccy if ccy else 1)
            aX = tf[2] + int(off.get('x')) * tf[0]
            aY = tf[3] + int(off.get('y')) * tf[1]
            _walk(ch, (sx, sy, aX - int(ch_off.get('x')) * sx,
                       aY - int(ch_off.get('y')) * sy), out)
        elif name in ('sp', 'cxnSp', 'pic'):
            _record_shape(ch, tf, name, out)

def _load_pptx(file_path):
    """读取 PPTX 包：页面尺寸(cm)、幻灯片页数、首页所有形状、图片/嵌入对象
    个数、页面背景是否白色或无填充。"""
    with zipfile.ZipFile(file_path) as z:
        pres = ET.fromstring(z.read('ppt/presentation.xml'))
        sld_sz = pres.find(_q('p:sldSz'))
        page_w = int(sld_sz.get('cx')) / EMU_PER_CM
        page_h = int(sld_sz.get('cy')) / EMU_PER_CM
        slides = sorted(
            (n for n in z.namelist() if re.match(r'ppt/slides/slide\d+\.xml$', n)),
            key=lambda n: int(re.search(r'(\d+)\.xml$', n).group(1)))
        if not slides:
            raise RuntimeError('PPTX 包内未发现任何幻灯片')
        root = ET.fromstring(z.read(slides[0]))

    cSld = root.find(_q('p:cSld'))
    sp_tree = cSld.find(_q('p:spTree'))
    shapes = []
    _walk(sp_tree, (1, 1, 0, 0), shapes)

    # 图片、艺术字、图表/表格等嵌入对象：p:pic 与 p:graphicFrame
    n_embed = (len(root.findall('.//' + _q('p:pic')))
               + len(root.findall('.//' + _q('p:graphicFrame'))))

    bg = cSld.find(_q('p:bg'))
    bg_ok = True
    if bg is not None:
        bgPr = bg.find(_q('p:bgPr'))
        if bgPr is not None:
            k, c = _color_of(bgPr)
            bg_ok = (k in ('noFill', 'theme', 'none')) or (k == 'rgb' and min(c) >= 250)

    return dict(page_w=page_w, page_h=page_h, slide_count=len(slides),
                shapes=shapes, n_embed=n_embed, bg_ok=bg_ok)


# ════════════════════════════════════════════════════════════════════════════
# 文本定位：同一段文字允许被拆到多个文本框，也允许多段文字合并在一个文本框
# ════════════════════════════════════════════════════════════════════════════
def _norm(s):
    """归一化文本：连字符/减号/负号视为等同字符，忽略所有空白。"""
    for a, b in (('−', '-'), ('–', '-'), ('—', '-'), ('‑', '-')):
        s = s.replace(a, b)
    return re.sub(r'[\s ]+', '', s)


def _rows(txt_shapes, tol=0.35):
    """按绝对坐标先按行由上至下分行（行内竖直容差 tol cm），行内由左至右排序。"""
    out = []
    for s in sorted(txt_shapes, key=lambda s: (s['y'], s['x'])):
        if out and abs(s['y'] - out[-1][0]) <= tol:
            out[-1][1].append(s)
        else:
            out.append([s['y'], [s]])
    return [(y, sorted(g, key=lambda s: s['x'])) for y, g in out]

def _row_str(group):
    """把一行内的所有文本框拼接为字符串，并记录每个字符来自哪个文本框。"""
    seq = []
    for s in group:
        for c in _norm(s['text']):
            seq.append((c, s))
    return ''.join(c for c, _ in seq), seq


def _uniq(seq_slice):
    """从字符序列切片中按出现顺序取出去重后的承载文本框。"""
    res = []
    seen = set()
    for _, s in seq_slice:
        if id(s) not in seen:
            seen.add(id(s))
            res.append(s)
    return res


def _find_all(row_list, target):
    """在每一行的拼接串里把目标文字当作连续子串查找，返回所有命中的承载文本框组。"""
    t = _norm(target)
    res = []
    for _, group in row_list:
        rs, seq = _row_str(group)
        start = 0
        while True:
            i = rs.find(t, start)
            if i < 0:
                break
            res.append(_uniq(seq[i:i + len(t)]))
            start = i + 1
    return res


def _bbox(shapes):
    """一组形状的整体外接框 (x0, y0, x1, y1)。"""
    return (min(s['x'] for s in shapes), min(s['y'] for s in shapes),
            max(s['x'] + s['w'] for s in shapes), max(s['y'] + s['h'] for s in shapes))


def _all_runs(shapes):
    res = []
    for s in shapes:
        res += s['runs']
    return res


def _size_in(shapes, lo, hi):
    """承载这段文字的所有文字段字号都落在 [lo, hi] 磅内；未显式设字号则放过。"""
    vals = [r['sz'] for r in _all_runs(shapes) if r['sz']]
    return all(lo <= v <= hi for v in vals) if vals else True

# ════════════════════════════════════════════════════════════════════════════
# 颜色判定：只认显式 RGB；未显式指定颜色（沿用主题）一律放过
# ════════════════════════════════════════════════════════════════════════════
def _is_green_line(s):
    """线条颜色的绿分量不低于100，且比红分量、蓝分量各高60以上。"""
    return (s['lk'] == 'rgb' and s['lc'][1] >= 100
            and s['lc'][1] - s['lc'][0] > 60 and s['lc'][1] - s['lc'][2] > 60)


def _is_green_any(s):
    """线条色或填充色任一为绿色（箭头既可能靠填充上色，也可能靠线条上色）。"""
    for k, c in ((s['lk'], s['lc']), (s['fk'], s['fc'])):
        if k == 'rgb' and c[1] >= 100 and c[1] - c[0] > 60 and c[1] - c[2] > 60:
            return True
    return False


def _is_blue_fill(s):
    """填充色的蓝分量不低于170，且比红分量、绿分量各高60以上。"""
    return (s['fk'] == 'rgb' and s['fc'][2] >= 170
            and s['fc'][2] - s['fc'][0] > 60 and s['fc'][2] - s['fc'][1] > 60)


def _is_yellow_fill(s):
    """填充色的红分量与绿分量均不低于200，且蓝分量不高于120。"""
    return (s['fk'] == 'rgb' and s['fc'][0] >= 200
            and s['fc'][1] >= 200 and s['fc'][2] <= 120)


def _is_gray_like(kind, rgb, max_channel=215):
    """黑色/灰色或沿用主题色：近中性且不过亮；未显式指定颜色一律放过。"""
    if kind != 'rgb':
        return True
    return max(rgb) - min(rgb) <= 25 and max(rgb) <= max_channel


def _line_width_in(s, lo=0.5, hi=1.5):
    """线宽落在 [lo, hi] 磅内；未显式设线宽（沿用默认）放过。"""
    return s['lw'] is None or lo <= s['lw'] <= hi


# ════════════════════════════════════════════════════════════════════════════
# 维度一：可用与可修改性
# ════════════════════════════════════════════════════════════════════════════
def _locate_pptx(dir_path):
    """在给定目录内定位待评估的演示文稿。维度一只接受 .pptx，故优先取 .pptx；
    目录内没有 .pptx 但有其它演示文稿格式（.ppt/.pptm/.pps/.odp 等）时，仍把它
    取回来，交给维度一按“格式不符”判不通过，而不是当成脚本异常。多个候选取
    修改时间最新的，忽略 ~$ 打头的临时锁文件。找不到返回 None。"""
    if not dir_path or not os.path.isdir(dir_path):
        return None
    other_exts = ('.ppt', '.pptm', '.ppsx', '.pps', '.ppsm', '.odp', '.key')
    pptx, other = [], []
    for f in os.listdir(dir_path):
        full = os.path.join(dir_path, f)
        if not os.path.isfile(full) or f.startswith('~$'):
            continue
        ext = os.path.splitext(f)[1].lower()
        if ext == '.pptx':
            pptx.append(full)
        elif ext in other_exts:
            other.append(full)
    pool = pptx if pptx else other
    if not pool:
        return None
    return max(pool, key=lambda p: os.path.getmtime(p))

def _check_dim1(file_path):
    """维度一门禁：交付文件为 .pptx 且可正常打开、只包含 1 页幻灯片。
    通过返回 (解析结果 dict, '')；不通过返回 (None, 不通过原因)。"""
    ext = os.path.splitext(file_path)[1].lower()
    if ext != '.pptx':
        return None, '交付文件格式不符：后缀为 %s，只接受 .pptx' % (ext or '（无后缀）')
    if not zipfile.is_zipfile(file_path):
        return None, '文件无法作为 PPTX 包打开（文件已损坏、需修复，或只是把其它格式改了后缀名）'
    try:
        doc = _load_pptx(file_path)
    except Exception as exc:
        return None, '打开/解析演示文稿失败（文件损坏或需修复）：%s: %s' % (type(exc).__name__, exc)
    if doc['slide_count'] != 1:
        return None, '幻灯片页数为 %d 页，不等于 1 页' % doc['slide_count']
    return doc, ''


# ════════════════════════════════════════════════════════════════════════════
# 维度二：完成度评分（全部为加分项，命中记 max_delta，未命中记 0）
# ════════════════════════════════════════════════════════════════════════════
def _evaluate_impl(doc, result):
    """维度二评估主体；对内构造 results（"+N：说明" 字符串）与等长的 details
    （实测证据，仅供调试，不写入返回信息），末尾按 _MAX_DELTAS_ORDER 反解成结构化
    dim2_items 填入 result，每项只含 rule / max_delta / delta / hit 四个键。"""
    page_w = doc['page_w']
    page_h = doc['page_h']
    sh = doc['shapes']
    results = []
    details = []

    def emit(max_delta, hit, rule, detail=''):
        results.append('+%d：%s' % (max_delta if hit else 0, rule))
        details.append(detail)

    # ─── 公共量 ───────────────────────────────────────────────────────────
    def px(s):
        """形状外框左上角距页面左边线的百分比。"""
        return s['x'] / page_w * 100.0

    def py(s):
        """形状外框左上角距页面顶部边线的百分比。"""
        return s['y'] / page_h * 100.0

    def center(s):
        return (s['x'] + s['w'] / 2.0, s['y'] + s['h'] / 2.0)

    txts = [s for s in sh if _norm(s['text'])]
    # 拼接行时不计入旋转角度不为 0° 或文字竖向排列的文本框
    plain = [s for s in txts if s['rot'] < 1 and not s['vert']]
    row_list = _rows(plain)
    n_shape_ok = len(sh) >= 100

    # 圆形：预设椭圆，或宽高之差不超过 0.05cm 的自定义形状
    circles = [s for s in sh if s['geo'] == 'ellipse'
               or (s['geo'] == 'cust' and abs(s['w'] - s['h']) <= 0.05)]
    # 水平/竖直线条：一个方向的外框尺寸不超过 0.1cm
    h_lines = [s for s in sh if s['h'] <= 0.1 and s['w'] > 0.1]
    v_lines = [s for s in sh if s['w'] <= 0.1 and s['h'] > 0.1]

    def dark_line(s):
        """线条颜色为黑色、深灰色或沿用主题色。"""
        return s['lk'] != 'rgb' or max(s['lc']) <= 110

    def find_text(target, x_range, y_range, size_range=None, need_bold=False):
        """在指定的页面百分比范围内查找目标文字，返回承载它的文本框组或 None。"""
        for group in _find_all(row_list, target):
            x0, y0, _, _ = _bbox(group)
            if not (x_range[0] <= x0 / page_w * 100 <= x_range[1]
                    and y_range[0] <= y0 / page_h * 100 <= y_range[1]):
                continue
            if size_range and not _size_in(group, *size_range):
                continue
            if need_bold and not any(r['b'] for r in _all_runs(group)):
                continue
            return group
        return None

    # ── 01 +3：形状总数不少于100个，页面 35.5–38.0 × 26.5–28.5cm，宽高比 1.28–1.39
    ratio = page_w / page_h if page_h else 0
    hit = (n_shape_ok and 35.5 <= page_w <= 38.0 and 26.5 <= page_h <= 28.5
           and 1.28 <= ratio <= 1.39)
    emit(3, hit, '页面宽度35.5–38.0cm、高度26.5–28.5cm、宽高比1.28–1.39，且形状与文本框总数不少于100个',
         '实测页面 %.2f×%.2fcm，宽高比 %.3f，形状与文本框 %d 个' % (page_w, page_h, ratio, len(sh)))

    # ── 02 +3：主标题，Arial/微软雅黑/沿用主题字体，14–18磅，加粗
    grp = find_text('RIVERDALE TUNNEL - SECTION 7 MONITORING LAYOUT',
                    (5, 30), (0, 8), (14, 18), True)
    fonts = sorted({r['font'] for r in _all_runs(grp) if r['font']}) if grp else []
    emit(3, grp is not None,
         '距左边线5%–30%、顶部边线0%–8%内出现主标题“RIVERDALE TUNNEL – SECTION 7 MONITORING LAYOUT”，字号14–18磅且加粗',
         ('命中，字体 %s' % (','.join(fonts) or '沿用主题')) if grp else '未在指定范围内找到符合字号/加粗要求的主标题')

    # ── 03 +1：副标题，12–17磅，非加粗
    grp = find_text('Displacement Sensors (mm)', (5, 30), (0, 14), (12, 17))
    emit(1, grp is not None,
         '距左边线5%–30%、顶部边线0%–14%内出现副标题“Displacement Sensors (mm)”，字号12–17磅',
         '命中' if grp else '未在指定范围内找到符合字号要求的副标题')

    # ── 04 +3：纵轴标题竖排（旋转90°/270° 或文字竖向排列），外框中心在
    #          距左边线0%–8%、从上至下25%–65%，字号14–18磅
    #          竖排文本框不参与行拼接，故这里单独在“含竖排”的全量文本里找
    vp = None
    for group in _find_all(_rows(txts), 'Vertical Position (m)'):
        x0, y0, x1, y1 = _bbox(group)
        rotated = any(abs(s['rot'] - a) < 1 for a in (90, 270) for s in group)
        vertical = any(s['vert'] for s in group)
        if not (rotated or vertical):
            continue
        if not (0 <= (x0 + x1) / 2 / page_w * 100 <= 8
                and 25 <= (y0 + y1) / 2 / page_h * 100 <= 65):
            continue
        if _size_in(group, 14, 18):
            vp = group
    emit(3, vp is not None,
         '出现文字“Vertical Position (m)”且旋转90°/270°或竖向排列，外框中心距左边线0%–8%、从上至下25%–65%，字号14–18磅',
         '命中，外框中心距左 %.1f%%、距顶 %.1f%%' % (
             (_bbox(vp)[0] + _bbox(vp)[2]) / 2 / page_w * 100,
             (_bbox(vp)[1] + _bbox(vp)[3]) / 2 / page_h * 100) if vp else '未找到竖向排列且位置/字号合规的纵轴标题')

    # ── 05 +1：横轴标题，旋转角度0°，字号14–18磅
    lp = find_text('Lateral Position (m)', (30, 70), (92, 100), (14, 18))
    hit = lp is not None and all(s['rot'] < 1 for s in lp)
    emit(1, hit,
         '距左边线30%–70%、顶部边线92%–100%内出现文字“Lateral Position (m)”，旋转角度0°，字号14–18磅',
         '命中' if hit else '未在指定范围内找到符合旋转角度/字号要求的横轴标题')

    # ── 06 +3：竖直实轴线，距左5%–9%，外框宽≤0.1cm、高23–26cm，线宽0.5–1.5磅，黑/深灰
    cand = [s for s in v_lines if 5 <= px(s) <= 9 and 23 <= s['h'] <= 26
            and s['dash'] in (None, 'solid') and _line_width_in(s) and dark_line(s)]
    emit(3, bool(cand),
         '距左边线5%–9%内出现一条竖直单实线，外框宽度不超过0.1cm、高度23–26cm，线宽0.5–1.5磅，颜色为黑色/深灰色/沿用主题色',
         '命中 %d 条，高 %.2fcm' % (len(cand), cand[0]['h']) if cand else '未找到符合位置/长度/线型的竖直实轴线')

    # ── 07 +3：水平实轴线，距顶91%–96%、距左5%–9%，外框高≤0.1cm、宽32–35cm
    cand = [s for s in h_lines if 91 <= py(s) <= 96 and 5 <= px(s) <= 9
            and 32 <= s['w'] <= 35 and s['dash'] in (None, 'solid')
            and _line_width_in(s) and dark_line(s)]
    emit(3, bool(cand),
         '距顶部边线91%–96%、左边线5%–9%内出现一条水平单实线，外框高度不超过0.1cm、宽度32–35cm，线宽0.5–1.5磅，颜色为黑色/深灰色/沿用主题色',
         '命中 %d 条，宽 %.2fcm' % (len(cand), cand[0]['w']) if cand else '未找到符合位置/长度/线型的水平实轴线')

    # ── 08 +5：距左边线0%–7%内自上而下7段刻度值 0.025/0.015/0.005/0/0.005/0.015/0.025，
    #          竖直位置分别落在 6–11 / 20–26 / 35–41 / 46–53 / 58–65 / 73–80 / 89–95%，字号10–15磅
    band_rows = _rows([s for s in plain if s['x'] / page_w <= 0.07])
    y_ticks = [('0.025', 6, 11), ('0.015', 20, 26), ('0.005', 35, 41), ('0', 46, 53),
               ('0.005', 58, 65), ('0.015', 73, 80), ('0.025', 89, 95)]

    def match_ticks(expected):
        """在左侧竖条带里按自上而下的顺序依次匹配各段刻度（去除正负号后比较）。"""
        i = 0
        got = []
        for want, lo, hi in expected:
            while i < len(band_rows):
                rs, _ = _row_str(band_rows[i][1])
                if rs.lstrip('+-') == want:
                    got.append((band_rows[i], lo, hi))
                    i += 1
                    break
                i += 1
            else:
                return None
        return got

    y_got = match_ticks(y_ticks)
    hit = bool(y_got) and all(lo <= _bbox(r[1])[1] / page_h * 100 <= hi and _size_in(r[1], 10, 15)
                              for r, lo, hi in y_got)
    emit(5, hit,
         '距左边线0%–7%内自上而下出现7段刻度文字，去除正负号后依次为0.025/0.015/0.005/0/0.005/0.015/0.025，竖直位置分别落在指定区间内，字号10–15磅',
         '按序命中 7 段，竖直位置 %s' % (
             '、'.join('%.1f%%' % (_bbox(r[1])[1] / page_h * 100) for r, _, _ in y_got))
         if hit else '仅按序匹配到 %d/7 段或位置、字号越界' % (len(y_got) if y_got else 0))

    # ── 09 +1：7段刻度正负号齐全（+/+/+/无/-/-/-）；正负号可以是文字字符，
    #          也可以是紧邻数字左侧的短线段或小形状（外框≤0.5cm、横向距离≤0.5cm、
    #          竖直位置落在该段数字外框上下范围内）
    def sign_ok(row, sign):
        rs, _ = _row_str(row[1])
        if sign in rs:
            return True
        x0 = min(s['x'] for s in row[1])
        y0 = min(s['y'] for s in row[1])
        y1 = max(s['y'] + s['h'] for s in row[1])
        for s in sh:
            if _norm(s['text']) or s['w'] > 0.5 or s['h'] > 0.5:
                continue
            if abs(s['x'] + s['w'] / 2 - x0) <= 0.5 and y0 <= s['y'] + s['h'] / 2 <= y1:
                return True
        return False

    signs = ['+', '+', '+', '', '-', '-', '-']
    sign_hits = ([sign_ok(r, sg) for (r, _, _), sg in zip(y_got, signs) if sg]
                 if y_got else [])
    hit = bool(y_got) and all(sign_hits)
    emit(1, hit,
         '距左边线0%–7%内自上而下的7段刻度正负号齐全，依次呈现为+0.025/+0.015/+0.005/0/-0.005/-0.015/-0.025，正负号可由文字字符或紧邻数字左侧的短线段、小形状表示',
         '6 个正负号命中 %d 个' % sum(sign_hits) if y_got else '未先匹配到7段刻度')

    # ── 10 +3：距顶93%–99%内自左至右5段刻度 -0.020/-0.010/0/+0.010/+0.020，
    #          横向位置分别落在 5–12 / 24–32 / 50–56 / 71–79 / 92–100%，字号10–15磅
    x_ticks = [('-0.020', 5, 12), ('-0.010', 24, 32), ('0', 50, 56),
               ('+0.010', 71, 79), ('+0.020', 92, 100)]
    x_ok = False
    x_detail = '未在距顶93%–99%的行内按序匹配到5段横轴刻度'
    for _, group in _rows([s for s in plain if 93 <= py(s) <= 99]):
        rs, seq = _row_str(group)
        pos = 0
        matched = []
        for want, lo, hi in x_ticks:
            i = rs.find(_norm(want), pos)
            if i < 0:
                matched = None
                break
            grp = _uniq(seq[i:i + len(_norm(want))])
            pos = i + len(_norm(want))
            if not (lo <= _bbox(grp)[0] / page_w * 100 <= hi):
                matched = None
                break
            matched.append(grp)
        if matched and all(_size_in(m, 10, 15) for m in matched):
            x_ok = True
            x_detail = '按序命中 5 段，横向位置 %s' % '、'.join(
                '%.1f%%' % (_bbox(m)[0] / page_w * 100) for m in matched)
    emit(3, x_ok,
         '距顶部边线93%–99%内自左至右出现5段刻度文字“-0.020”“-0.010”“0”“+0.010”“+0.020”，横向位置分别落在指定区间内，字号10–15磅',
         x_detail)

    # ── 11 +1：距左5%–8%内不少于6条水平短线段、距顶92%–96%内不少于4条竖直短线段，
    #          每条长度0.1–0.6cm，颜色为黑色/深灰色/沿用主题色
    h_tick = [s for s in h_lines if 5 <= px(s) <= 8 and 0.1 <= s['w'] <= 0.6 and dark_line(s)]
    v_tick = [s for s in v_lines if 92 <= py(s) <= 96 and 0.1 <= s['h'] <= 0.6 and dark_line(s)]
    hit = len(h_tick) >= 6 and len(v_tick) >= 4
    emit(1, hit,
         '距左边线5%–8%内出现不少于6条水平短线段、距顶部边线92%–96%内出现不少于4条竖直短线段，每条长度0.1–0.6cm，颜色为黑色/深灰色/沿用主题色',
         '水平短线 %d 条、竖直短线 %d 条' % (len(h_tick), len(v_tick)))

    # ── 12 +3：距顶47%–52%内一条水平虚线（短划线类），外框高≤0.1cm、宽32–35cm，
    #          线宽0.5–1.5磅，中等灰色或沿用主题色
    cand = [s for s in h_lines if 47 <= py(s) <= 52 and 32 <= s['w'] <= 35
            and s['dash'] and s['dash'] != 'solid' and _line_width_in(s)
            and _is_gray_like(s['lk'], s['lc'])
            and (s['lk'] != 'rgb' or 100 <= max(s['lc']) <= 215)]
    emit(3, bool(cand),
         '距顶部边线47%–52%内出现一条短划线类水平虚线，外框高度不超过0.1cm、宽度32–35cm，线宽0.5–1.5磅，颜色为中等灰色或沿用主题色',
         '命中，线型 %s、宽 %.2fcm' % (cand[0]['dash'], cand[0]['w']) if cand
         else '未找到符合位置/长度的水平虚线（%d 条水平线中虚线 %d 条）' % (
             len(h_lines), len([s for s in h_lines if s['dash'] and s['dash'] != 'solid'])))

    # ── 13 +3：距左51%–56%内一条竖直虚线（短划线类），外框宽≤0.1cm、高22–25cm
    cand = [s for s in v_lines if 51 <= px(s) <= 56 and 22 <= s['h'] <= 25
            and s['dash'] and s['dash'] != 'solid' and _line_width_in(s)
            and _is_gray_like(s['lk'], s['lc'])
            and (s['lk'] != 'rgb' or 100 <= max(s['lc']) <= 215)]
    emit(3, bool(cand),
         '距左边线51%–56%内出现一条短划线类竖直虚线，外框宽度不超过0.1cm、高度22–25cm，线宽0.5–1.5磅，颜色为中等灰色或沿用主题色',
         '命中，线型 %s、高 %.2fcm' % (cand[0]['dash'], cand[0]['h']) if cand
         else '未找到符合位置/长度的竖直虚线（%d 条竖直线中虚线 %d 条）' % (
             len(v_lines), len([s for s in v_lines if s['dash'] and s['dash'] != 'solid'])))

    # ─── 绿色网线（等值线网格）───────────────────────────────────────────
    green = [s for s in sh if _is_green_line(s)]
    if green:
        gx0 = min(s['x'] for s in green)
        gy0 = min(s['y'] for s in green)
        gx1 = max(s['x'] + s['w'] for s in green)
        gy1 = max(s['y'] + s['h'] for s in green)
    else:
        gx0 = gy0 = gx1 = gy1 = 0.0

    # ── 14 +5：100–300个绿色线条形状，线宽均0.5–1.5磅，整体外接范围宽≥28cm、高≥13cm，
    #          且整体落在距左6%–100%、距顶6%–92%的区域内
    hit = (100 <= len(green) <= 300 and all(_line_width_in(s) for s in green)
           and bool(green) and gx1 - gx0 >= 28 and gy1 - gy0 >= 13
           and 6 <= gx0 / page_w * 100 and gx1 / page_w * 100 <= 100
           and 6 <= gy0 / page_h * 100 and gy1 / page_h * 100 <= 92)
    emit(5, hit,
         '页面内出现100–300个绿色线条形状（绿分量不低于100且比红、蓝分量各高60以上），线宽均为0.5–1.5磅，整体外接范围宽度不小于28cm、高度不小于13cm，且整体落在距左边线6%–100%、顶部边线6%–92%的区域内',
         '绿色线条 %d 个，外接范围 %.2f×%.2fcm，位于距左 %.1f%%–%.1f%%、距顶 %.1f%%–%.1f%%' % (
             len(green), gx1 - gx0, gy1 - gy0, gx0 / page_w * 100, gx1 / page_w * 100,
             gy0 / page_h * 100, gy1 / page_h * 100) if green else '未找到任何绿色线条形状')

    # ── 15 +5：不少于80个绿色线条为含曲线控制点的自由曲线形状
    n_curved = len([s for s in green if s['curved']])
    emit(5, n_curved >= 80,
         '绿色线条形状中不少于80个为含曲线控制点的自由曲线形状，即绿色线条并非全部由直线段构成',
         '自由曲线 %d 个 / 绿色线条 %d 个' % (n_curved, len(green)))

    # ── 16 +3：绿色线条的线宽取值唯一、颜色取值唯一
    widths = {s['lw'] for s in green}
    colors = {s['lc'] for s in green}
    hit = bool(green) and len(widths) == 1 and len(colors) == 1
    emit(3, hit,
         '绿色线条形状的线宽取值唯一、颜色取值唯一，不存在粗细不一或深浅不一的绿色线条',
         '线宽取值 %d 种、颜色取值 %d 种' % (len(widths), len(colors)) if green else '未找到任何绿色线条形状')

    # ── 17 +3：外框宽高均大于0.05cm的倾斜形状占绿色线条总数80%以上；按外框宽高之比
    #          换算的倾角落在0–15 / 15–45 / 45–75 / 75–90° 四个区间的各不少于5个
    slant = [s for s in green if s['w'] > 0.05 and s['h'] > 0.05]
    bins = {'0-15': 0, '15-45': 0, '45-75': 0, '75-90': 0}
    for s in green:
        ang = math.degrees(math.atan2(s['h'], max(s['w'], 1e-9)))
        key = ('0-15' if ang < 15 else '15-45' if ang < 45
               else '45-75' if ang < 75 else '75-90')
        bins[key] += 1
    hit = bool(green) and len(slant) >= 0.8 * len(green) and all(v >= 5 for v in bins.values())
    emit(3, hit,
         '绿色线条形状中外框宽度与高度均大于0.05cm的倾斜形状占其总数的80%以上，且按外框宽高之比换算的倾角落在0°–15°、15°–45°、45°–75°、75°–90°四个区间内的各不少于5个',
         '倾斜 %d/%d 个（%.1f%%），四区间分布 %s' % (
             len(slant), len(green), (len(slant) / len(green) * 100) if green else 0,
             '/'.join('%s:%d' % (k, bins[k]) for k in ('0-15', '15-45', '45-75', '75-90'))))

    # ── 18 +3：外框对角长度不小于12cm的不少于2个、不小于8cm的不少于5个
    diags = sorted(math.hypot(s['w'], s['h']) for s in green)
    n12 = sum(1 for v in diags if v >= 12)
    n8 = sum(1 for v in diags if v >= 8)
    emit(3, n12 >= 2 and n8 >= 5,
         '绿色线条形状中外框对角长度不小于12cm的不少于2个、不小于8cm的不少于5个，即勾勒整体轮廓的长弧线为连续的整条形状，未被拆成多段短线拼接',
         '对角长 ≥12cm 的 %d 个、≥8cm 的 %d 个，最长 %.2fcm' % (
             n12, n8, diags[-1] if diags else 0))

    # ── 19 +3：绿色线条彼此首尾相连构成连通网格：端点间直线距离不超过0.05cm 视为相接；
    #          全部端点中相接比例不低于50%，两端都不相接的孤立形状不超过总数的30%
    with_ends = [s for s in green if s['ends']]
    eps = [(e, i) for i, s in enumerate(with_ends) for e in s['ends']]
    TOL = 0.05
    linked = [any(j != i and math.hypot(e[0] - f[0], e[1] - f[1]) <= TOL for f, j in eps)
              for e, i in eps]
    per_shape = {}
    for k, (_, i) in enumerate(eps):
        per_shape.setdefault(i, []).append(linked[k])
    isolated = sum(1 for v in per_shape.values() if not any(v))
    hit = (bool(with_ends) and len(eps) >= 100 and sum(linked) >= 0.5 * len(eps)
           and isolated <= 0.3 * len(green))
    emit(3, hit,
         '绿色线条形状彼此首尾相连、构成连通的网格：端点之间直线距离不超过0.05cm 视为相接，全部端点中相接的比例不低于50%，两个端点都不与任何其他绿色线条相接的孤立形状不超过绿色线条总数的30%',
         '端点 %d 个，相接率 %.1f%%，孤立形状 %d 个（%.1f%%）' % (
             len(eps), (sum(linked) / len(eps) * 100) if eps else 0, isolated,
             (isolated / len(green) * 100) if green else 0))

    # ── 20 +3：距顶43%–55%内11个蓝色圆形，直径0.25–0.40cm 且相互差值≤0.03cm，
    #          圆心横向位置全部落在距左28%–74%
    blue = [s for s in circles if _is_blue_fill(s) and 43 <= center(s)[1] / page_h * 100 <= 55]
    hit = (len(blue) == 11 and all(0.25 <= s['w'] <= 0.40 for s in blue)
           and max(s['w'] for s in blue) - min(s['w'] for s in blue) <= 0.03
           and all(28 <= center(s)[0] / page_w * 100 <= 74 for s in blue))
    emit(3, hit,
         '距顶部边线43%–55%内出现11个蓝色圆形（填充蓝分量不低于170且比红、绿分量各高60以上），直径均为0.25–0.40cm且相互差值不超过0.03cm，圆心横向位置全部落在距左边线28%–74%内',
         '蓝色圆形 %d 个，直径 %.2f–%.2fcm' % (
             len(blue), min(s['w'] for s in blue), max(s['w'] for s in blue))
         if blue else '未在指定竖直范围内找到蓝色圆形')

    # ── 21 +3：11段蓝点编号 B114–B125，每段文字外框中心与最近蓝色圆形圆心的
    #          横向、竖直距离均不超过1.2cm，一一对应，字号8–13磅
    blue_all = [s for s in circles if _is_blue_fill(s)]
    labels_b = ['B114', 'B116', 'B117', 'B118', 'B119', 'B120',
                'B121', 'B122', 'B123', 'B124', 'B125']
    matched_b = []
    for lab in labels_b:
        for group in _find_all(row_list, lab):
            x0, y0, x1, y1 = _bbox(group)
            cx, cy = (x0 + x1) / 2, (y0 + y1) / 2
            if any(abs(cx - center(c)[0]) <= 1.2 and abs(cy - center(c)[1]) <= 1.2
                   for c in blue_all):
                matched_b.append(group)
                break
    hit = len(matched_b) == 11 and all(_size_in(m, 8, 13) for m in matched_b)
    emit(3, hit,
         '页面内出现文字“B114”“B116”“B117”“B118”“B119”“B120”“B121”“B122”“B123”“B124”“B125”共11段，每段文字外框中心与最近的蓝色圆形圆心横向、竖直距离均不超过1.2cm且一一对应，字号8–13磅',
         '与蓝点配对成功 %d/11 段' % len(matched_b))

    # ── 22 +3：距顶13%–36%内8个黄色圆形，直径0.25–0.40cm，按圆心横坐标由小到大
    #          分别落在 10–15 / 19–24 / 32–38 / 63–68 / 74–79 / 84–89 / 90–94 / 94–100%
    y_up = sorted([s for s in circles if _is_yellow_fill(s)
                   and 13 <= center(s)[1] / page_h * 100 <= 36],
                  key=lambda s: center(s)[0])
    x_up = [(10, 15), (19, 24), (32, 38), (63, 68), (74, 79), (84, 89), (90, 94), (94, 100)]
    hit = (len(y_up) == 8 and all(0.25 <= s['w'] <= 0.40 for s in y_up)
           and all(a <= center(s)[0] / page_w * 100 <= b for s, (a, b) in zip(y_up, x_up)))
    emit(3, hit,
         '距顶部边线13%–36%内出现8个黄色圆形（填充红、绿分量均不低于200且蓝分量不高于120），直径均为0.25–0.40cm，按圆心横坐标由小到大分别落在距左边线10%–15%、19%–24%、32%–38%、63%–68%、74%–79%、84%–89%、90%–94%、94%–100%内',
         '黄色圆形 %d 个，横向位置 %s' % (len(y_up), '、'.join(
             '%.1f%%' % (center(s)[0] / page_w * 100) for s in y_up)))

    # ── 23 +3：距顶56%–73%内11个黄色圆形，直径0.25–0.40cm，按圆心横坐标由小到大
    #          分别落在 7–11 / 13–17 / 22–26 / 31–35 / 39–43 / 49–53 / 63–67 / 71–75
    #          / 78–82 / 85–89 / 92–96%
    y_dn = sorted([s for s in circles if _is_yellow_fill(s)
                   and 56 <= center(s)[1] / page_h * 100 <= 73],
                  key=lambda s: center(s)[0])
    x_dn = [(7, 11), (13, 17), (22, 26), (31, 35), (39, 43), (49, 53),
            (63, 67), (71, 75), (78, 82), (85, 89), (92, 96)]
    hit = (len(y_dn) == 11 and all(0.25 <= s['w'] <= 0.40 for s in y_dn)
           and all(a <= center(s)[0] / page_w * 100 <= b for s, (a, b) in zip(y_dn, x_dn)))
    emit(3, hit,
         '距顶部边线56%–73%内出现11个黄色圆形（填充红、绿分量均不低于200且蓝分量不高于120），直径均为0.25–0.40cm，按圆心横坐标由小到大分别落在距左边线7%–11%、13%–17%、22%–26%、31%–35%、39%–43%、49%–53%、63%–67%、71%–75%、78%–82%、85%–89%、92%–96%内',
         '黄色圆形 %d 个，横向位置 %s' % (len(y_dn), '、'.join(
             '%.1f%%' % (center(s)[0] / page_w * 100) for s in y_dn)))

    def match_yellow_labels(names, y_lo, y_hi):
        """在指定竖直范围内，把每个编号文字与最近的黄色圆形配对。"""
        ys = [c for c in circles if _is_yellow_fill(c)
              and y_lo <= center(c)[1] / page_h * 100 <= y_hi]
        found = []
        for lab in names:
            for group in _find_all(row_list, lab):
                x0, y0, x1, y1 = _bbox(group)
                cx, cy = (x0 + x1) / 2, (y0 + y1) / 2
                if not (y_lo <= y0 / page_h * 100 <= y_hi):
                    continue
                if any(abs(cx - center(c)[0]) <= 1.2 and abs(cy - center(c)[1]) <= 1.2
                       for c in ys):
                    found.append((cx, lab, group))
                    break
        return found

    # ── 24 +3：距顶13%–36%内按横坐标由小到大出现 K908/K903/K899/K651/K657/K663/K670/K679
    labels_k = ['K908', 'K903', 'K899', 'K651', 'K657', 'K663', 'K670', 'K679']
    found_k = match_yellow_labels(labels_k, 13, 36)
    order_k = [lab for _, lab, _ in sorted(found_k, key=lambda t: t[0])]
    hit = (len(found_k) == 8 and order_k == labels_k
           and all(_size_in(g, 8, 13) for _, _, g in found_k))
    emit(3, hit,
         '距顶部边线13%–36%内按横坐标由小到大出现文字“K908”“K903”“K899”“K651”“K657”“K663”“K670”“K679”共8段，每段文字外框中心与最近的黄色圆形圆心横向、竖直距离均不超过1.2cm，字号8–13磅',
         '与黄点配对成功 %d/8 段，实际左右顺序 %s' % (len(found_k), '/'.join(order_k)))

    # ── 25 +3：距顶56%–73%内按横坐标由小到大出现 A201/A203/A205/A207/A209/A211/
    #          A301/A303/A301/A308/A314 共11段（其中 A301 重复出现两次，按不同文字段计）
    labels_a = ['A201', 'A203', 'A205', 'A207', 'A209', 'A211',
                'A301', 'A303', 'A301', 'A308', 'A314']
    uniq_a = list(dict.fromkeys(labels_a))
    found_a = match_yellow_labels(uniq_a, 56, 73)
    hit = len(found_a) == len(uniq_a) and all(_size_in(g, 8, 13) for _, _, g in found_a)
    emit(3, hit,
         '距顶部边线56%–73%内按横坐标由小到大出现文字“A201”“A203”“A205”“A207”“A209”“A211”“A301”“A303”“A301”“A308”“A314”共11段，每段文字外框中心与最近的黄色圆形圆心横向、竖直距离均不超过1.2cm，字号8–13磅',
         '与黄点配对成功 %d/%d 段（A301 重复出现，按去重后的 %d 种编号计）' % (
             len(found_a), len(uniq_a), len(uniq_a)))

    # ─── 左下比例尺 / 右下图例 ────────────────────────────────────────────
    round_rects = [s for s in sh if s['geo'] == 'roundRect']

    def box_ok(s, x_range, y_range, w_range, h_range):
        """圆角矩形：位置、宽高在范围内，单实线、线宽0.5–1.2磅、黑/灰/主题色，
        填充为白色或无填充。"""
        return (x_range[0] <= px(s) <= x_range[1] and y_range[0] <= py(s) <= y_range[1]
                and w_range[0] <= s['w'] <= w_range[1]
                and h_range[0] <= s['h'] <= h_range[1]
                and s['dash'] in (None, 'solid')
                and (s['lw'] is None or 0.5 <= s['lw'] <= 1.2)
                and _is_gray_like(s['lk'], s['lc'])
                and (s['fk'] in ('noFill', 'none', 'theme')
                     or (s['fk'] == 'rgb' and min(s['fc']) >= 250)))

    # ── 26 +3：距左7%–11%、距顶79%–85%内一个圆角矩形，宽4.2–4.8cm，高2.3–2.8cm
    cand = [s for s in round_rects if box_ok(s, (7, 11), (79, 85), (4.2, 4.8), (2.3, 2.8))]
    emit(3, bool(cand),
         '距左边线7%–11%、顶部边线79%–85%内出现一个圆角矩形，宽4.2–4.8cm、高2.3–2.8cm，边线为单实线、线宽0.5–1.2磅、颜色为黑色/灰色/沿用主题色，填充为白色或无填充',
         '命中，%.2f×%.2fcm' % (cand[0]['w'], cand[0]['h']) if cand
         else '未找到符合位置/尺寸/边线/填充要求的圆角矩形（页面共 %d 个圆角矩形）' % len(round_rects))

    # ── 27 +3：距左9%–15%、距顶83%–90%内两个绿色箭头形状，一横一竖、均带明确箭头端、
    #          外框宽高均为正值、长边0.5–1.0cm、两者外框左边线横向间距不超过较宽者宽度+0.3cm
    def effective_box(s):
        """按形状在页面上的实际显示方向计算外框：旋转90°/270°时宽高互换。"""
        cx, cy = center(s)
        ew, eh = ((s['h'], s['w']) if (85 <= s['rot'] <= 95 or 265 <= s['rot'] <= 275)
                  else (s['w'], s['h']))
        return (cx - ew / 2, cy - eh / 2, ew, eh)

    def is_arrow(s):
        """带明确箭头端：箭头类预设形状、设置了箭头端点样式的线条、或带填充的自定义箭头形状。"""
        return ('arrow' in (s['geo'] or '').lower()
                or s['hd'] not in (None, 'none')
                or s['tl'] not in (None, 'none')
                or (s['geo'] == 'cust' and s['fk'] == 'rgb'))

    arrows = []
    for s in sh:
        if not _is_green_any(s):
            continue
        ex, ey, ew, eh = effective_box(s)
        if 9 <= ex / page_w * 100 <= 15 and 83 <= ey / page_h * 100 <= 90:
            arrows.append((ex, ey, ew, eh, s))
    hit = (len(arrows) == 2
           and all(is_arrow(s) for *_, s in arrows)
           and all(s['sw'] >= 0 and s['sh'] >= 0 for *_, s in arrows)
           and any(w > h for _, _, w, h, _ in arrows)
           and any(h > w for _, _, w, h, _ in arrows)
           and all(0.5 <= max(w, h) <= 1.0 for _, _, w, h, _ in arrows)
           and abs(arrows[0][0] - arrows[1][0]) <= max(arrows[0][2], arrows[1][2]) + 0.3)
    emit(3, hit,
         '距左边线9%–15%、顶部边线83%–90%内出现两个绿色箭头形状，均带有明确的箭头端，一个横向且箭头端朝右、一个竖向且箭头端朝上，外框宽度与高度均为正值（不依靠负尺寸或翻转表达方向），长边均为0.5–1.0cm，两者外框左边线横向间距不超过较宽者宽度加0.3cm',
         '命中 %d 个绿色形状：%s' % (len(arrows), '；'.join(
             '%s %.2f×%.2fcm rot=%.0f° 箭头端=%s 带符号尺寸(%.2f,%.2f)' % (
                 s['geo'], w, h, s['rot'], '有' if is_arrow(s) else '无', s['sw'], s['sh'])
             for _, _, w, h, s in arrows)) if arrows else '未在指定范围内找到绿色箭头形状')

    # ── 28 +1：比例尺内 Y (Up) 与 X (Right) 两段标注，字号8–13磅
    g1 = find_text('Y (Up)', (8, 13), (80, 85), (8, 13))
    g2 = find_text('X (Right)', (10, 16), (83, 88), (8, 13))
    emit(1, g1 is not None and g2 is not None,
         '距左边线8%–13%、顶部边线80%–85%内出现文字“Y (Up)”，距左边线10%–16%、顶部边线83%–88%内出现文字“X (Right)”，字号8–13磅',
         'Y (Up) %s，X (Right) %s' % ('命中' if g1 else '未命中', '命中' if g2 else '未命中'))

    # ── 29 +1：比例尺说明文字 Displacement Scale: 1 mm，字号6–12磅
    g1 = find_text('Displacement Scale: 1 mm', (7, 12), (86, 91), (6, 12))
    emit(1, g1 is not None,
         '距左边线7%–12%、顶部边线86%–91%内出现文字“Displacement Scale: 1 mm”，字号6–12磅',
         '命中' if g1 else '未在指定范围内找到符合字号要求的比例尺说明文字')

    # ── 30 +3：距左77%–82%、距顶79%–85%内一个圆角矩形，宽6.5–7.2cm，高1.7–2.2cm
    cand = [s for s in round_rects if box_ok(s, (77, 82), (79, 85), (6.5, 7.2), (1.7, 2.2))]
    emit(3, bool(cand),
         '距左边线77%–82%、顶部边线79%–85%内出现一个圆角矩形，宽6.5–7.2cm、高1.7–2.2cm，边线为单实线、线宽0.5–1.2磅、颜色为黑色/灰色/沿用主题色，填充为白色或无填充',
         '命中，%.2f×%.2fcm' % (cand[0]['w'], cand[0]['h']) if cand
         else '未找到符合位置/尺寸/边线/填充要求的圆角矩形（页面共 %d 个圆角矩形）' % len(round_rects))

    # ── 31 +3：图例内一个蓝色圆形（距左79%–83%、距顶82%–86%）与一个黄色圆形
    #          （距左79%–83%、距顶85%–89%），直径均0.25–0.40cm，圆心横坐标相差≤0.1cm
    lb = [s for s in circles if _is_blue_fill(s)
          and 79 <= center(s)[0] / page_w * 100 <= 83
          and 82 <= center(s)[1] / page_h * 100 <= 86]
    ly = [s for s in circles if _is_yellow_fill(s)
          and 79 <= center(s)[0] / page_w * 100 <= 83
          and 85 <= center(s)[1] / page_h * 100 <= 89]
    hit = (len(lb) == 1 and len(ly) == 1
           and all(0.25 <= s['w'] <= 0.40 for s in lb + ly)
           and abs(center(lb[0])[0] - center(ly[0])[0]) <= 0.1)
    emit(3, hit,
         '距左边线79%–83%、顶部边线82%–86%内出现一个蓝色圆形，距左边线79%–83%、顶部边线85%–89%内出现一个黄色圆形，两个圆形直径均为0.25–0.40cm，圆心横坐标相差不超过0.1cm',
         '蓝色圆形 %d 个、黄色圆形 %d 个' % (len(lb), len(ly)))

    # ── 32 +3：图例两行说明文字，字号7–12磅
    g1 = find_text('Crown/Centerline Sensors', (80, 86), (81, 85), (7, 12))
    g2 = find_text('Upper/Lower Boundary Sensors', (80, 86), (84, 88), (7, 12))
    emit(3, g1 is not None and g2 is not None,
         '距左边线80%–86%、顶部边线81%–85%内出现文字“Crown/Centerline Sensors”，距左边线80%–86%、顶部边线84%–88%内出现文字“Upper/Lower Boundary Sensors”，字号7–12磅',
         'Crown/Centerline Sensors %s，Upper/Lower Boundary Sensors %s' % (
             '命中' if g1 else '未命中', '命中' if g2 else '未命中'))

    # ── 33 +5：形状与文本框总数不少于100个，且不存在图片、艺术字与嵌入对象，
    #          全部内容由形状、连接符与文本框构成（每个元素均可单独选中并修改）
    hit = n_shape_ok and doc['n_embed'] == 0
    emit(5, hit,
         '形状与文本框总数不少于100个，且不存在图片、艺术字与嵌入对象，全部内容由形状、连接符与文本框构成，每条网线、每个圆形、每条坐标轴与虚线、每个圆角矩形与箭头、每个文本框均可单独选中并修改',
         '形状与文本框 %d 个，图片/嵌入对象 %d 个' % (len(sh), doc['n_embed']))

    # ── 34 +3：交付规范——正尺寸、未启用缩放文字以适应形状、整数字号、
    #          除竖排纵轴标题外文本框不旋转不翻转且完整落在页面内、白底、无大色块、无效果
    neg = [s for s in sh if s['sw'] < 0 or s['sh'] < 0]
    zero = [s for s in sh if s['geo'] not in ('line', 'straightConnector1', None)
            and (s['w'] <= 0 or s['h'] <= 0)]
    sizes = [r['sz'] for s in txts for r in s['runs'] if r['sz']]
    non_int = sorted({v for v in sizes if abs(v - round(v)) > 1e-6})
    autofit_bad = [s for s in txts if s['af'] == 'norm']
    in_bounds = all(-0.1 <= s['x'] and -0.1 <= s['y']
                    and s['x'] + s['w'] <= page_w + 0.1
                    and s['y'] + s['h'] <= page_h + 0.1
                    for s in txts if s['rot'] < 1)
    rot_ok = all(s['rot'] < 1 and not s['flip'] for s in txts
                 if 'VerticalPosition' not in _norm(s['text']))
    big_fill = [s for s in sh if s['w'] > 30 and s['h'] > 20 and s['fk'] == 'rgb']
    effects = [s for s in sh if s['eff']]
    hit = (n_shape_ok and not neg and not zero and not autofit_bad and not non_int
           and rot_ok and in_bounds and doc['bg_ok'] and not big_fill and not effects)
    emit(3, hit,
         '形状与文本框总数不少于100个，所有形状宽高不为负值且除线条外均大于0，文本框未启用“缩放文字以适应形状”，所有字号为整数磅值，除竖向排列的“Vertical Position (m)”外文本框旋转角度均为0°且未翻转、外框完整落在页面内（各边溢出不超过0.1cm），页面背景为白色或无填充，不存在宽度大于30cm且高度大于20cm的填充形状，所有形状均未设置阴影/发光/映像/柔化边缘/三维等效果',
         '负尺寸形状 %d 个，零尺寸形状 %d 个，非整数字号 %s，自动缩放文字 %d 个，'
         '文本框旋转/翻转%s，文本框越界%s，背景%s，超大填充形状 %d 个，带效果形状 %d 个' % (
             len(neg), len(zero),
             ('/'.join('%.1f磅' % v for v in non_int) if non_int else '无'),
             len(autofit_bad), '异常' if not rot_ok else '正常',
             '异常' if not in_bounds else '正常',
             '合规' if doc['bg_ok'] else '非白底', len(big_fill), len(effects)))

    # ========== 将 results 反解为结构化 dim2_items ==========
    # 逐项按追加顺序对齐 _MAX_DELTAS_ORDER：解析每条 "+N：说明" 的 N 与 rule。
    if len(results) != len(_MAX_DELTAS_ORDER) or len(details) != len(results):
        # 顺序表与实际条目不一致：视为脚本内部错误，交给外层兜底
        raise RuntimeError(
            'dim2 项数不一致：results=%d, details=%d, order=%d'
            % (len(results), len(details), len(_MAX_DELTAS_ORDER)))
    dim2_items = []
    total_score = 0
    for raw, max_delta in zip(results, _MAX_DELTAS_ORDER):
        # 格式约定：'+N：' 前缀 + 规则说明；命中 N=max_delta，未命中 N=0
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
        total_score += (max_delta if hit else 0)

    result['dim2_items'] = dim2_items
    result['total_score'] = total_score
    result['max_score'] = MAX_SCORE_TOTAL
    return result


def evaluate(dir_path):
    """按接口统一约定返回结构化 dict；不 print、不改 sys.stdout、不 sys.exit。
    入参 dir_path 为“脚本所在目录的路径”，脚本自己在该目录里定位并打开被评估的
    .pptx（忽略以 ~$ 开头的临时锁文件）。

    status='error' 只在“这份提交没被真正评过”时出现——目录不存在、目录内没有
    演示文稿、或脚本自身抛异常；此时 dim1_pass 为空串（未判定）、dim1_reason 为
    空串、dim2_items 为空表、total_score 与 max_score 一律为 0。
    文件能定位到但格式不符 / 损坏 / 页数不为 1，仍按 status='ok' + 维度一未通过
    处理，此时 dim1_pass=False 且 dim1_reason 带有未通过的错误摘要。"""
    result = {
        'id': SCRIPT_ID,
        'file_name': '',
        'status': 'ok',
        'error': None,
        # 空串表示“维度一尚未判定”；status='error' 的三条出口都保持空串，
        # 只有真正跑过 _check_dim1 才会写入 True / False。
        'dim1_pass': '',
        'dim1_reason': '',
        'dim2_items': [],
        'total_score': 0,
        'max_score': 0,
    }
    try:
        if not dir_path or not os.path.isdir(dir_path):
            result['status'] = 'error'
            result['error'] = '目录不存在：%s' % dir_path
            return result
        file_path = _locate_pptx(dir_path)
        if file_path is None:
            result['status'] = 'error'
            result['error'] = '目录中未找到 .pptx 文件：%s' % dir_path
            return result
        result['file_name'] = os.path.basename(file_path)

        # ─── 维度1：可用与可修改性（不通过则完成度总分记 0）───
        doc, reason = _check_dim1(file_path)
        if doc is None:
            # 约定：dim1_pass 为 False 时 dim1_reason 必须带错误摘要。
            result['dim1_pass'] = False
            result['dim1_reason'] = reason or '维度一未通过（未给出具体原因）'
            result['dim2_items'] = []
            result['total_score'] = 0
            result['max_score'] = MAX_SCORE_TOTAL
            return result

        # 维度一通过：按统一约定 dim1_reason 保持空串（只有未通过才写原因）。
        result['dim1_pass'] = True
        result['dim1_reason'] = ''

        # ─── 维度2：完成度评分 ───
        return _evaluate_impl(doc, result)
    except Exception as exc:
        # _evaluate_impl 是就地写 result 的，异常可能发生在它写了一半之后，
        # 因此这里把所有承载评分结论的字段显式复位，保证 error 的返回值是干净的全 0；
        # dim1_pass 复位为空串而不是 False——异常意味着维度一的结论作废，不是判定为不通过。
        result['status'] = 'error'
        result['error'] = '%s: %s' % (type(exc).__name__, exc)
        result['dim1_pass'] = ''
        result['dim1_reason'] = ''
        result['dim2_items'] = []
        result['total_score'] = 0
        result['max_score'] = 0
        return result


if __name__ == '__main__':
    # 本地调试入口：evaluate 只接收“脚本所在目录的路径”，由脚本自己在该目录里
    # 定位并打开被评估的 .pptx。命令行第 1 个参数可覆盖目录路径；未指定时使用
    # 脚本自身所在目录。主结果只走 evaluate 的 return，这里仅作可视化输出。
    if len(sys.argv) >= 2:
        _dir = sys.argv[1]
    else:
        _dir = os.path.dirname(os.path.abspath(__file__))
    _out = evaluate(_dir)
    # 直接写 UTF-8 字节到 stdout.buffer，避免 Windows 控制台默认 GBK 编码打印
    # 中文报 UnicodeEncodeError，同时不修改 sys.stdout 本身
    _payload = json.dumps(_out, ensure_ascii=False, indent=2) + '\n'
    try:
        _ = sys.stdout.buffer.write(_payload.encode('utf-8'))
    except AttributeError:
        print(json.dumps(_out, ensure_ascii=True, indent=2))

