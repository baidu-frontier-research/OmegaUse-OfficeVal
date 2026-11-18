#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# pyright: reportUnknownMemberType=false, reportUnknownVariableType=false, reportUnknownArgumentType=false, reportUnknownParameterType=false, reportMissingParameterType=false, reportPrivateUsage=false, reportUnknownLambdaType=false, reportGeneralTypeIssues=false

import os
import re
import sys
import json
import math
import colorsys
from lxml import etree
from pptx import Presentation

# 说明：按接口统一约定，本模块不修改全局 sys.stdout，也不 print 主结果。
# 主结果统一由 evaluate(dir_path) -> dict 返回；入参为“脚本所在目录的路径”，
# 由脚本自己在该目录内定位并打开被评估的 .pptx。本文件仅在 __main__ 下作为
# 本地调试入口，才把 dict 序列化为 JSON 打印。

SCRIPT_ID = "046"
# 维度二所有评分项 max_delta 之和；用于维度一未通过 / 脚本异常时兜底
MAX_SCORE_TOTAL = 128

# 维度二所有评分项按“结果追加顺序”的 max_delta 列表；末端把 results 里的
# "+N：..." 字符串反解为结构化 dim2_items 时用。顺序须与 _evaluate_impl 中
# results.append(...) 的顺序严格一致；共 32 项，累加 = MAX_SCORE_TOTAL。
_MAX_DELTAS_ORDER = [
    3,  # 01 共2页、一页一图、无多余空白页
    5,  # 02 页面尺寸 30cm × 20cm
    5,  # 03 等比置入：主体宽高比 1.33-1.37、两页缩放一致
    3,  # 04 主体水平且垂直居中（左右/上下留白差≤0.5cm）
    3,  # 05 元素不越页被裁切、无散落线条与页角斜线
    3,  # 06 底部单行居中图题：黑色加粗、字号大于卡片正文
    5,  # 07 开放式立体箱体三面（正面/侧面/顶面）闭合衔接
    3,  # 08 箱体顶面 4 个同向倾斜平行四边形标签
    5,  # 09 左侧竖栏 2 虚线分 3 格 + 图标在上文字在下 + 2 向下箭头
    3,  # 10 上部 4 个等大圆角框 + 3 个实线右箭头
    3,  # 11 4 框虚线折线汇聚到浅蓝框（膜污染演化…）
    3,  # 12 橙色虚线浅橙框：上行常规 / 下行橙色加粗更大
    3,  # 13 左右带箭头折线指向 2 个并列面板（标题蓝色加粗居顶居中）
    5,  # 14 两面板内各 2×2 四卡 + 每卡左侧配图标左右并排
    3,  # 15 面板下汇合折线指向“通量衰减减缓·清洗周期延长·分离效率提升”
    5,  # 16 箱体内立式膜组件压力容器装置（高度占比 0.55-0.85）
    5,  # 17 箱体右外红色加粗竖排 12 字
    5,  # 18 顶部 4 个右指箭头色带（青绿/蓝/紫/橙 + 圆形图标 + 衔接小箭头）
    5,  # 19 左右两栏各 5 模块、对称等宽对齐 + 四个标题条各带 2 个子项框
    3,  # 20 左右各 1 个橙黄虚线浅米黄说明框（两行不溢出不折第三行）
    3,  # 21 说明框下各 1 个粗下箭头 + 倒梯形块（健康数据 / 随访评估）
    5,  # 22 底部 3 个圆角色块等高底边对齐 + 中块左右粗箭头 + 各配图标
    3,  # 23 蓝色虚线大圆角区域四边闭合 + 底部居中蓝字
    5,  # 24 中央深蓝实心圆 + 白色加粗“风险等级” + 完整蓝色虚线圆环
    5,  # 25 圆外 4 个彩色几何块 + 各配矢量图标、文字单行
    5,  # 26 4 个白底小标签 + 8 条放射状双向箭头（不穿入中央圆）
    5,  # 27 左侧条形图面板（7 标签 + 7 递减蓝条 + 横轴刻度/标题）
    5,  # 28 右侧折线图面板（3 条连续折线抬升不互穿 + 双轴刻度/标题）
    3,  # 29 纵轴竖排标题方向正立 + 右上 3 行图例单行不压盖
    5,  # 30 两页文字与结构元素均为可单独编辑的对象、非整片贴图
    3,  # 31 图标逐个矢量/独立图片放置，不用符号字符代替
    3,  # 32 观感规整：同类卡片统一、虚实线用法一致、无压盖漏字
]
assert sum(_MAX_DELTAS_ORDER) == MAX_SCORE_TOTAL

_NS_A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
_NS_P = 'http://schemas.openxmlformats.org/presentationml/2006/main'
EMU_CM = 360000.0          # 1cm = 360000 EMU
_PT_CM = 2.54 / 72.0       # 1pt = 0.0353cm
_DEF_FONT_PT = 18.0        # PowerPoint 文本默认字号，run/段落/形状均未显式设定时按此回落


def _qa(tag):
    return '{%s}%s' % (_NS_A, tag)


# ============================================================================
# 一、几何：分组坐标变换 + 形状“视觉墨迹外框”
# ============================================================================

def _group_tf(grp, tf):
    """把分组的子坐标系变换叠加到外层变换上（与办公软件的渲染口径一致）：
       分组 p:grpSpPr/a:xfrm 提供 off/ext/chOff/chExt，子元素坐标 x 映射为
       X = (ext/chExt) * x + (off - chOff * ext/chExt)；多层分组按外∘内复合。
       变换用四元组 (ax, bx, ay, by) 表示：X = ax*x + bx，Y = ay*y + by。"""
    xf = grp._element.grpSpPr.find(_qa('xfrm'))
    if xf is None:
        return tf
    off, ext = xf.find(_qa('off')), xf.find(_qa('ext'))
    cho, che = xf.find(_qa('chOff')), xf.find(_qa('chExt'))
    if None in (off, ext, cho, che):
        return tf
    try:
        ox, oy = int(off.get('x')), int(off.get('y'))
        ew, eh = int(ext.get('cx')), int(ext.get('cy'))
        cx, cy = int(cho.get('x')), int(cho.get('y'))
        cw, ch = int(che.get('cx')), int(che.get('cy'))
    except (TypeError, ValueError):
        return tf
    ax = (ew / cw) if cw else 1.0
    ay = (eh / ch) if ch else 1.0
    tax, tbx, tay, tby = tf
    return (tax * ax, tax * (ox - cx * ax) + tbx,
            tay * ay, tay * (oy - cy * ay) + tby)


def _iter_leaf(shapes, tf=(1.0, 0.0, 1.0, 0.0)):
    """递归展开分组，只产出叶子形状及其累计坐标变换。"""
    for sh in shapes:
        try:
            is_grp = (sh.shape_type == 6)
        except Exception:
            is_grp = sh.element.tag.endswith('grpSp')
        if is_grp:
            for item in _iter_leaf(sh.shapes, _group_tf(sh, tf)):
                yield item
        else:
            yield sh, tf


def _frame_cm(sh, tf):
    """形状未旋转时的框（cm）：(x0, y0, w, h)；取不到位置的形状返回 None。"""
    try:
        l, t, w, h = sh.left, sh.top, sh.width, sh.height
    except Exception:
        return None
    if None in (l, t, w, h):
        return None
    ax, bx, ay, by = tf
    return ((ax * l + bx) / EMU_CM, (ay * t + by) / EMU_CM,
            abs(ax * w) / EMU_CM, abs(ay * h) / EMU_CM)

def _rot_deg(sh):
    try:
        return float(sh.rotation or 0.0)
    except Exception:
        return 0.0


def _flips(sh):
    xf = sh._element.find('.//' + _qa('xfrm'))
    if xf is None:
        return False, False
    return xf.get('flipH') == '1', xf.get('flipV') == '1'


def _custgeom_pts(sh):
    """custGeom 路径点，归一化到 [0,1]×[0,1] 的“路径空间”。
       a:path/@w,@h 声明路径空间尺寸并整体映射到形状 ext；个别生成器会写出
       略微超出该空间的点，这里统一夹到 [0,1]，即以形状自身框为墨迹上界。
       只取显式 a:pt（moveTo/lnTo/cubicBezTo 控制点），arcTo 等按外框近似。"""
    cg = sh._element.find('.//' + _qa('custGeom'))
    if cg is None:
        return None
    pts = []
    for p in cg.iter(_qa('path')):
        try:
            pw = float(p.get('w') or 0)
            ph = float(p.get('h') or 0)
        except (TypeError, ValueError):
            continue
        if pw <= 0 or ph <= 0:
            return None
        for t in p.iter(_qa('pt')):
            try:
                u = float(t.get('x')) / pw
                v = float(t.get('y')) / ph
            except (TypeError, ValueError):
                continue
            pts.append((min(max(u, 0.0), 1.0), min(max(v, 0.0), 1.0)))
    return pts or None


def _map_pt_list(pts, x0, y0, w, h, rot, fh, fv):
    """把路径空间的点逐个映射到页面坐标：先按 flipH/flipV 翻转，
       再按形状自身旋转（绕框几何中心、顺时针为正）旋转。"""
    cx, cy = x0 + w / 2.0, y0 + h / 2.0
    th = math.radians(rot % 360.0)
    co, si = math.cos(th), math.sin(th)
    out = []
    for u, v in pts:
        if fh:
            u = 1.0 - u
        if fv:
            v = 1.0 - v
        dx, dy = x0 + u * w - cx, y0 + v * h - cy
        out.append((cx + dx * co - dy * si, cy + dx * si + dy * co))
    return out


def _map_pts(pts, x0, y0, w, h, rot, fh, fv):
    """同 _map_pts_list，但只返回映射后点集的 AABB。"""
    q = _map_pt_list(pts, x0, y0, w, h, rot, fh, fv)
    xs = [p[0] for p in q]
    ys = [p[1] for p in q]
    return (min(xs), min(ys), max(xs), max(ys))

# ============================================================================
# 二、样式：填充色 / 描边色 / 虚线 / 箭头端点 / 颜色族
# ============================================================================

_THEME = {}
_SCHEME_ALIAS = {'bg1': 'lt1', 'bg2': 'lt2', 'tx1': 'dk1', 'tx2': 'dk2'}
_SCHEME_DEFAULT = {'lt1': (255, 255, 255), 'lt2': (238, 236, 225),
                   'dk1': (0, 0, 0), 'dk2': (31, 73, 125)}


def _load_theme(prs):
    """解析 /ppt/theme/themeN.xml 的 a:clrScheme，建立 schemeClr 名 → RGB 映射。
       主题色解析失败时退回 lt1/dk1 等常规默认值；用于识别以主题色（而非 srgbClr）
       着色的文字与填充，避免把“白色加粗文字”误判为无颜色。"""
    out = dict(_SCHEME_DEFAULT)
    try:
        for part in prs.part.package.iter_parts():
            nm = str(part.partname)
            if '/theme/theme' not in nm:
                continue
            root = etree.fromstring(part.blob)
            cs = root.find('.//' + _qa('clrScheme'))
            if cs is None:
                continue
            for child in cs:
                key = etree.QName(child).localname
                srgb = child.find(_qa('srgbClr'))
                sysc = child.find(_qa('sysClr'))
                v = None
                if srgb is not None:
                    v = (srgb.get('val') or '').strip()
                elif sysc is not None:
                    v = (sysc.get('lastClr') or '').strip()
                if v and len(v) == 6:
                    try:
                        out[key] = (int(v[0:2], 16), int(v[2:4], 16), int(v[4:6], 16))
                    except ValueError:
                        pass
            break
    except Exception:
        pass
    return out


def _srgb_of(el):
    """从一个填充/描边容器里取颜色：优先 srgbClr，其次按主题色表解析 schemeClr。"""
    if el is None:
        return None
    c = el.find('.//' + _qa('srgbClr'))
    if c is not None:
        v = (c.get('val') or '').strip()
        if len(v) == 6:
            try:
                return (int(v[0:2], 16), int(v[2:4], 16), int(v[4:6], 16))
            except ValueError:
                return None
        return None
    c = el.find('.//' + _qa('schemeClr'))
    if c is None:
        return None
    nm = (c.get('val') or '').strip()
    nm = _SCHEME_ALIAS.get(nm, nm)
    return _THEME.get(nm)


def _sp_pr(sh):
    """形状属性容器：幻灯片里是 p:spPr（p:sp / p:pic / p:cxnSp）或 p:grpSpPr。
       注意其命名空间是 presentationml 而非 drawingml，直接找 a:spPr 会取不到填充色。"""
    el = sh._element
    pr = getattr(el, 'spPr', None)
    if pr is not None:
        return pr
    for ns in (_NS_P, _NS_A):
        p = el.find('.//{%s}spPr' % ns)
        if p is not None:
            return p
    return el.find('.//{%s}grpSpPr' % _NS_P)


def _fill_rgb(sh):
    """形状填充色（仅认 solidFill/gradFill 的第一个 srgbClr）；无填充或主题色返回 None。"""
    pr = _sp_pr(sh)
    if pr is None:
        return None
    for tag in ('solidFill', 'gradFill'):
        f = pr.find(_qa(tag))
        if f is not None:
            return _srgb_of(f)
    return None


def _has_grad_fill(sh):
    pr = _sp_pr(sh)
    return pr is not None and pr.find(_qa('gradFill')) is not None


def _no_fill(sh):
    pr = _sp_pr(sh)
    return pr is not None and pr.find(_qa('noFill')) is not None


def _line_el(sh):
    pr = _sp_pr(sh)
    return None if pr is None else pr.find(_qa('ln'))


def _line_rgb(sh):
    return _srgb_of(_line_el(sh))


_DASH_VALS = ('dash', 'sysDash', 'lgDash', 'dashDot', 'lgDashDot',
              'lgDashDotDot', 'sysDashDot', 'sysDashDotDot', 'dot', 'sysDot')


def _is_dashed(sh):
    """描边为虚线族（a:prstDash 或 a:custDash）。"""
    ln = _line_el(sh)
    if ln is None:
        return False
    if ln.find(_qa('custDash')) is not None:
        return True
    d = ln.find(_qa('prstDash'))
    return d is not None and (d.get('val') or '') in _DASH_VALS

def _arrow_ends(sh):
    """描边两端的箭头类型：返回 (头端有箭头, 尾端有箭头)。"""
    ln = _line_el(sh)
    if ln is None:
        return False, False
    def _on(tag):
        e = ln.find(_qa(tag))
        return e is not None and (e.get('type') or 'none') != 'none'
    return _on('headEnd'), _on('tailEnd')


def _hue_family(rgb):
    """把 RGB 归到粗颜色族，用于判定“绿/蓝/紫/橙/青/红/黄”等指定配色。
       口径：先按明度/饱和度挑出白、黑、灰，再按色相扇区归族。"""
    if rgb is None:
        return None
    r, g, b = [v / 255.0 for v in rgb]
    h, s, v = colorsys.rgb_to_hsv(r, g, b)
    deg = h * 360.0
    if v >= 0.93 and s <= 0.10:
        return 'white'
    if v <= 0.22:
        return 'black'
    if s <= 0.10:
        return 'gray'
    if deg < 16 or deg >= 345:
        return 'red'
    if deg < 45:
        return 'orange'
    if deg < 70:
        return 'yellow'
    if deg < 160:
        return 'green'
    if deg < 200:
        return 'cyan'
    if deg < 255:
        return 'blue'
    if deg < 290:
        return 'purple'
    return 'magenta'


def _is_lightish(rgb, thr=0.80):
    """浅色填充（明度高）：用于“浅蓝/浅橙/浅米黄填充”一类判定。"""
    if rgb is None:
        return False
    v = max(rgb) / 255.0
    s = 0.0 if max(rgb) == 0 else (max(rgb) - min(rgb)) / float(max(rgb))
    return v >= thr and s <= 0.45


def _is_darkish(rgb, thr=0.72):
    if rgb is None:
        return False
    return (max(rgb) / 255.0) <= thr


def _is_whiteish(rgb):
    return rgb is not None and min(rgb) >= 235

# ============================================================================
# 三、文本：有效字号/加粗/颜色、文字视觉宽度估算、文本单元与“视觉行”重建
# ============================================================================

def _defrpr_of(sh):
    """形状层 a:lstStyle/a:lvl1pPr/a:defRPr（未在 run/段落显式设定时的落地层）。"""
    el = sh._element.find('.//' + _qa('lstStyle'))
    if el is None:
        return None
    lvl = el.find(_qa('lvl1pPr'))
    return None if lvl is None else lvl.find(_qa('defRPr'))


def _eff_size(run, para, sh, defrpr):
    """run 中文字实际生效的字号（pt）：run → 段落 → 形状 defRPr → 默认 18pt。
       与办公软件的继承链一致；末端回落 18pt 是 PowerPoint 的默认正文字号。"""
    try:
        if run is not None and run.font.size:
            return float(run.font.size.pt)
    except Exception:
        pass
    try:
        if para is not None and para.font.size:
            return float(para.font.size.pt)
    except Exception:
        pass
    for el in (defrpr,):
        if el is not None and el.get('sz'):
            try:
                return float(el.get('sz')) / 100.0
            except ValueError:
                pass
    return _DEF_FONT_PT


def _eff_bold(run, para, defrpr):
    for f in (getattr(run, 'font', None), getattr(para, 'font', None)):
        try:
            if f is not None and f.bold is not None:
                return bool(f.bold)
        except Exception:
            pass
    if defrpr is not None and defrpr.get('b') is not None:
        return defrpr.get('b') in ('1', 'true')
    return False


def _eff_color(run, para, defrpr):
    for f in (getattr(run, 'font', None), getattr(para, 'font', None)):
        try:
            if f is not None and f.color is not None and f.color.rgb is not None:
                v = str(f.color.rgb)
                return (int(v[0:2], 16), int(v[2:4], 16), int(v[4:6], 16))
        except Exception:
            pass
    if defrpr is not None:
        return _srgb_of(defrpr)
    return None

def _char_w(ch, pt):
    """单字符视觉宽度（cm）：中日韩与全角标点按 1 字宽，半角字母数字≈0.52 字宽。"""
    if ch in ('\t',):
        return 2.0 * pt * _PT_CM
    o = ord(ch)
    if o == 32:
        return 0.28 * pt * _PT_CM
    if o < 128:
        return 0.52 * pt * _PT_CM
    if 0x3000 <= o <= 0x303F or 0xFF00 <= o <= 0xFF60:
        return 1.0 * pt * _PT_CM
    if o >= 0x1100:
        return 1.0 * pt * _PT_CM
    return 0.55 * pt * _PT_CM


def _para_text(para):
    return ''.join(r.text or '' for r in para.runs) or (para.text or '')


def _para_width_cm(para, sh, defrpr):
    """段落文字排成一行时的视觉宽度（cm）。"""
    w = 0.0
    runs = list(para.runs)
    if not runs:
        pt = _eff_size(None, para, sh, defrpr)
        return sum(_char_w(c, pt) for c in (para.text or ''))
    for r in runs:
        pt = _eff_size(r, para, sh, defrpr)
        for c in (r.text or ''):
            w += _char_w(c, pt)
    return w


def _inner_w(sh, fw):
    """文本框可排版内宽（cm）：扣掉左右内边距，取不到按 PowerPoint 默认 0.25cm。"""
    try:
        li = sh.text_frame.margin_left
        ri = sh.text_frame.margin_right
        li = 0.25 if li is None else li / EMU_CM
        ri = 0.25 if ri is None else ri / EMU_CM
    except Exception:
        li = ri = 0.25
    return max(fw - li - ri, 0.05)


def _wrap_lines(para, sh, fw, fh, defrpr, pt):
    """该段落在框内实际占的视觉行数（>1 即发生自动换行/折行）。
       宽度估算天生偏保守（字宽按等宽近似），因此再用框高做一次上限约束：
       框高只放得下一行时按一行计，避免把“文字略微顶满窄框”误判成折行。"""
    tw = _para_width_cm(para, sh, defrpr)
    iw = _inner_w(sh, fw)
    try:
        wrap = sh.text_frame.word_wrap
    except Exception:
        wrap = None
    if wrap is False:
        return 1
    if tw <= iw * 1.15 or iw <= 0.05:
        return 1
    est = int(math.ceil(tw / iw - 0.15))
    lh = max(1.2 * pt * _PT_CM, 0.05)
    cap = max(1, int(fh / lh + 0.35))
    return max(1, min(est, cap))

def _text_units(sh, tf):
    """把一个形状里的文字拆成“视觉文本单元”（每段一个），并估算其墨迹范围。
       ① 水平位置：按段落对齐（左/居中/右）在框内定位，宽度取 min(内宽, 文字估算宽)，
          避免个别大文本框把整页外框撑大；
       ② 垂直位置：按 anchor（顶/中/底）与行高 1.2×字号 逐段堆叠；
       ③ 形状自身旋转≠0 时，直接用旋转后的框 AABB 作为该单元范围。"""
    if not getattr(sh, 'has_text_frame', False):
        return []
    tfm = sh.text_frame
    if not (tfm.text or '').strip():
        return []
    fr = _frame_cm(sh, tf)
    if fr is None:
        return []
    x0, y0, fw, fh = fr
    defrpr = _defrpr_of(sh)
    rot = _rot_deg(sh)
    try:
        li = 0.25 if tfm.margin_left is None else tfm.margin_left / EMU_CM
        ti = 0.13 if tfm.margin_top is None else tfm.margin_top / EMU_CM
        ri = 0.25 if tfm.margin_right is None else tfm.margin_right / EMU_CM
    except Exception:
        li, ti, ri = 0.25, 0.13, 0.25
    anchor = str(getattr(tfm, 'vertical_anchor', None) or '')
    paras = [p for p in tfm.paragraphs if (_para_text(p) or '').strip()]
    if not paras:
        return []
    metas = []
    total_h = 0.0
    for p in paras:
        pts = [_eff_size(r, p, sh, defrpr) for r in p.runs] or [_eff_size(None, p, sh, defrpr)]
        pt = max(pts)
        nl = _wrap_lines(p, sh, fw, fh, defrpr, pt)
        metas.append((p, pt, nl, _para_width_cm(p, sh, defrpr)))
        total_h += nl * 1.2 * pt * _PT_CM
    if 'MIDDLE' in anchor:
        cur = y0 + max((fh - total_h) / 2.0, 0.0)
    elif 'BOTTOM' in anchor:
        cur = y0 + max(fh - total_h - 0.05, 0.0)
    else:
        cur = y0 + min(ti, max(fh - total_h, 0.0) if fh > total_h else ti)
    out = []
    for p, pt, nl, tw in metas:
        hh = nl * 1.2 * pt * _PT_CM
        vis_w = min(max(_inner_w(sh, fw), 0.05), tw if nl == 1 else _inner_w(sh, fw))
        algn = str(p.alignment or '')
        if 'CENTER' in algn:
            ux0 = x0 + (fw - vis_w) / 2.0
        elif 'RIGHT' in algn:
            ux0 = x0 + fw - ri - vis_w
        else:
            ux0 = x0 + li
        ux1, uy0, uy1 = ux0 + vis_w, cur, cur + hh
        if abs(rot % 360.0) > 0.5:
            bx = _map_pts([(0, 0), (1, 0), (1, 1), (0, 1)], x0, y0, fw, fh, rot, False, False)
            ux0, uy0, ux1, uy1 = bx
        out.append({
            'sh': sh, 'text': _para_text(p), 'pt': pt, 'lines': nl, 'rot': rot,
            'bold': all(_eff_bold(r, p, defrpr) for r in p.runs) if p.runs else False,
            'color': next((c for c in (_eff_color(r, p, defrpr) for r in p.runs) if c), None),
            'x0': ux0, 'y0': uy0, 'x1': ux1, 'y1': uy1,
            'cx': (ux0 + ux1) / 2.0, 'cy': (uy0 + uy1) / 2.0,
        })
        cur += hh
    return out

_WS_RE = re.compile(r'[\s　 ]+')


def _norm(s):
    """归一化文本用于比对：去空白、统一全角/半角括号与间隔符。"""
    s = _WS_RE.sub('', s or '')
    for a, b in (('（', '('), ('）', ')'), ('，', ','), ('：', ':'),
                 ('•', '·'), ('・', '·'), ('．', '.')):
        s = s.replace(a, b)
    return s


def _counter(s):
    d = {}
    for ch in s:
        d[ch] = d.get(ch, 0) + 1
    return d


def _cover(cand, target):
    """cand 对 target 的字符覆盖率（按重数计），用于容忍个别缺字/错序。"""
    if not target:
        return 0.0
    cc, tc = _counter(cand), _counter(target)
    hit = sum(min(n, cc.get(ch, 0)) for ch, n in tc.items())
    return hit / float(len(target))


def _rows_of(units):
    """把文本单元按“视觉行”聚类：y 区间重叠或中心线接近的归为一行。
       复刻类文件常把整句拆成逐字文本框，按行聚类后才能还原成可比对的行文本。"""
    rows = []
    for u in sorted(units, key=lambda z: (z['cy'], z['x0'])):
        placed = False
        for r in rows:
            tol = max(0.20, 0.45 * min(u['y1'] - u['y0'], r['h']))
            if abs(u['cy'] - r['cy']) <= tol:
                r['units'].append(u)
                n = len(r['units'])
                r['cy'] = sum(z['cy'] for z in r['units']) / n
                r['h'] = sum(z['y1'] - z['y0'] for z in r['units']) / n
                r['y0'] = min(r['y0'], u['y0'])
                r['y1'] = max(r['y1'], u['y1'])
                placed = True
                break
        if not placed:
            rows.append({'units': [u], 'cy': u['cy'], 'h': u['y1'] - u['y0'],
                         'y0': u['y0'], 'y1': u['y1']})
    for r in rows:
        r['units'].sort(key=lambda z: z['x0'])
        segs = []
        for u in r['units']:
            if segs and u['x0'] - segs[-1]['x1'] <= 0.42:
                segs[-1]['units'].append(u)
                segs[-1]['x1'] = max(segs[-1]['x1'], u['x1'])
            else:
                segs.append({'units': [u], 'x0': u['x0'], 'x1': u['x1']})
        for s in segs:
            s['text'] = _norm(''.join(z['text'] for z in s['units']))
        r['segs'] = segs
        r['text'] = _norm(''.join(s['text'] for s in segs))
        r['x0'] = min(u['x0'] for u in r['units'])
        r['x1'] = max(u['x1'] for u in r['units'])
    rows.sort(key=lambda r: r['cy'])
    return rows

def _mk_hit(row, segs, target):
    us = [u for s in segs for u in s['units']]
    txt = _norm(''.join(s['text'] for s in segs))
    return {
        'row': row, 'units': us, 'text': txt, 'cover': _cover(txt, target),
        'x0': min(s['x0'] for s in segs), 'x1': max(s['x1'] for s in segs),
        'y0': min(u['y0'] for u in us), 'y1': max(u['y1'] for u in us),
        'pt': max(u['pt'] for u in us),
        'bold': all(u['bold'] for u in us),
        'colors': [u['color'] for u in us if u['color']],
        'wrapped': any(u['lines'] > 1 for u in us),
        'nsegs': len(segs),
    }


def _find_all(rows, target, region=None, min_cover=None, slack=2):
    """在“视觉行”里找目标短语：同一行内取若干相邻文本段拼接后按字符覆盖率比对。
       ① 逐字文本框、行内错序都能命中（按重数覆盖，不看顺序）；
       ② 长度上限 len(target)+slack，避免把整行别的文字也吞进来；
       ③ 命中结果自带 wrapped 标记，供“单行不换行”类细则判定。"""
    tgt = _norm(target)
    if min_cover is None:
        min_cover = 1.0 if len(tgt) <= 3 else 0.8
    hits = []
    for r in rows:
        segs = r['segs']
        for i in range(len(segs)):
            acc, ln = [], 0
            for j in range(i, len(segs)):
                acc.append(segs[j])
                ln += len(segs[j]['text'])
                if ln > len(tgt) + slack:
                    break
                if ln < max(len(tgt) - slack, 1):
                    continue
                h = _mk_hit(r, acc, tgt)
                if h['cover'] >= min_cover:
                    hits.append(h)
                    break
    if region is not None:
        hits = [h for h in hits if _pt_in(h['x0'] * .5 + h['x1'] * .5,
                                          h['y0'] * .5 + h['y1'] * .5, region)]
    hits.sort(key=lambda h: (-h['cover'], abs(len(h['text']) - len(tgt))))
    out, seen = [], set()
    for h in hits:
        key = (round(h['x0'], 1), round(h['y0'], 1))
        if key in seen:
            continue
        seen.add(key)
        out.append(h)
    return out


def _find(rows, target, region=None, min_cover=None, slack=2):
    hits = _find_all(rows, target, region, min_cover, slack)
    return hits[0] if hits else None


def _pt_in(x, y, box, tol=0.0):
    return (box[0] - tol) <= x <= (box[2] + tol) and (box[1] - tol) <= y <= (box[3] + tol)


def _box_in(b, box, tol=0.0):
    return (b[0] >= box[0] - tol and b[1] >= box[1] - tol
            and b[2] <= box[2] + tol and b[3] <= box[3] + tol)


def _inter(a, b):
    w = min(a[2], b[2]) - max(a[0], b[0])
    h = min(a[3], b[3]) - max(a[1], b[1])
    return max(w, 0.0) * max(h, 0.0)

# ============================================================================
# 四、页面模型：把一页拍平成带“视觉墨迹外框”的形状表 + 视觉行表
# ============================================================================

def _prst_of(sh):
    g = sh._element.find('.//' + _qa('prstGeom'))
    if g is not None:
        return g.get('prst') or ''
    return 'custom' if sh._element.find('.//' + _qa('custGeom')) is not None else ''


def _shape_kind(sh, prst):
    try:
        st = int(sh.shape_type) if sh.shape_type is not None else -1
    except Exception:
        st = -1
    if st == 13:
        return 'pic'
    if st == 9 or prst in ('line', 'straightConnector1'):
        return 'line'
    if prst == 'custom':
        return 'free'
    return 'auto'


def _visible(sh, kind, has_text):
    if kind == 'pic' or has_text:
        return True
    if _fill_rgb(sh) is not None or _has_grad_fill(sh):
        return True
    ln = _line_el(sh)
    if ln is not None and ln.find(_qa('noFill')) is None:
        return True
    return not _no_fill(sh)


def _ink_of(sh, tf, kind, units):
    fr = _frame_cm(sh, tf)
    if fr is None:
        return None
    x0, y0, w, h = fr
    rot = _rot_deg(sh)
    fh, fv = _flips(sh)
    if kind == 'line':
        pts = [(1.0, 0.0), (0.0, 1.0)] if fh != fv else [(0.0, 0.0), (1.0, 1.0)]
        return _map_pts(pts, x0, y0, w, h, rot, False, False)
    if kind == 'free':
        pts = _custgeom_pts(sh)
        if pts:
            return _map_pts(pts, x0, y0, w, h, rot, fh, fv)
    if units and kind == 'auto' and _fill_rgb(sh) is None and not _has_grad_fill(sh) \
            and _line_el(sh) is None:
        # 纯文本框：以估算出的文字墨迹为准，避免超宽空文本框撑大整页外框
        return (min(u['x0'] for u in units), min(u['y0'] for u in units),
                max(u['x1'] for u in units), max(u['y1'] for u in units))
    return _map_pts([(0, 0), (1, 0), (1, 1), (0, 1)], x0, y0, w, h, rot, False, False)

def _sample_pts(sh, tf, kind, ink):
    """形状的“折线取样点”（页面坐标）：line 取两端点，free 取 custGeom 路径点，
       其余（如小圆点数据标记）退化为墨迹中心一个点。用于折线走向/次序判定。"""
    fr = _frame_cm(sh, tf)
    if fr is None or ink is None:
        return [((ink[0] + ink[2]) / 2.0, (ink[1] + ink[3]) / 2.0)] if ink else []
    x0, y0, w, h = fr
    rot = _rot_deg(sh)
    fh, fv = _flips(sh)
    if kind == 'line':
        pts = [(1.0, 0.0), (0.0, 1.0)] if fh != fv else [(0.0, 0.0), (1.0, 1.0)]
        return _map_pt_list(pts, x0, y0, w, h, rot, False, False)
    if kind == 'free':
        pts = _custgeom_pts(sh)
        if pts:
            return _map_pt_list(pts, x0, y0, w, h, rot, fh, fv)
    return [((ink[0] + ink[2]) / 2.0, (ink[1] + ink[3]) / 2.0)]


_ARROW_PRST = ('rightArrow', 'leftArrow', 'upArrow', 'downArrow', 'bentArrow',
               'curvedDownArrow', 'curvedUpArrow', 'curvedLeftArrow', 'stripedRightArrow',
               'curvedRightArrow', 'notchedRightArrow', 'leftRightArrow', 'upDownArrow',
               'bentUpArrow', 'triangle', 'chevron', 'homePlate', 'pentagon')


def _build_slide(slide, pw, ph):
    """把一页拍平成模型：shapes（带墨迹外框与样式）、units/rows（文本）、bbox（画面主体）。"""
    shapes, units = [], []
    for sh, tf in _iter_leaf(slide.shapes):
        prst = _prst_of(sh)
        kind = _shape_kind(sh, prst)
        try:
            has_tf = bool(getattr(sh, 'has_text_frame', False)
                          and (sh.text_frame.text or '').strip())
        except Exception:
            has_tf = False
        us = _text_units(sh, tf) if has_tf else []
        ink = _ink_of(sh, tf, kind, us)
        if ink is None:
            continue
        w, h = ink[2] - ink[0], ink[3] - ink[1]
        fill = _fill_rgb(sh)
        lnc = _line_rgb(sh)
        he, te = _arrow_ends(sh)
        rec = {
            'sh': sh, 'sid': getattr(sh, 'shape_id', 0), 'prst': prst, 'kind': kind,
            'ink': ink, 'x0': ink[0], 'y0': ink[1], 'x1': ink[2], 'y1': ink[3],
            'w': w, 'h': h, 'cx': (ink[0] + ink[2]) / 2.0, 'cy': (ink[1] + ink[3]) / 2.0,
            'area': w * h, 'fill': fill, 'line': lnc, 'dashed': _is_dashed(sh),
            'fillfam': _hue_family(fill), 'linefam': _hue_family(lnc),
            'text': _norm(sh.text_frame.text) if has_tf else '',
            'has_text': has_tf, 'rot': _rot_deg(sh),
            'seglen': math.hypot(w, h) if kind in ('line', 'free') else None,
            'arrowish': (prst in _ARROW_PRST) or he or te,
            'biarrow': he and te,
            'visible': _visible(sh, kind, has_tf),
            'grad': _has_grad_fill(sh),
            'pts': _sample_pts(sh, tf, kind, ink),
        }
        shapes.append(rec)
        units.extend(us)
    vis = [s for s in shapes if s['visible'] and s['w'] * s['h'] >= 0.0 and (s['w'] + s['h']) > 0.02]
    if vis:
        bbox = (min(s['x0'] for s in vis), min(s['y0'] for s in vis),
                max(s['x1'] for s in vis), max(s['y1'] for s in vis))
    else:
        bbox = (0.0, 0.0, pw, ph)
    return {'shapes': shapes, 'vis': vis, 'units': units, 'rows': _rows_of(units),
            'bbox': bbox, 'pw': pw, 'ph': ph,
            'bw': bbox[2] - bbox[0], 'bh': bbox[3] - bbox[1]}

# ============================================================================
# 五、区域查询：形状筛选 / 虚线证据 / 图标候选 / 箭头候选
# ============================================================================

def _in_reg(model, region, pred=None, key='center'):
    out = []
    for s in model['vis']:
        if key == 'center':
            ok = _pt_in(s['cx'], s['cy'], region)
        else:
            ok = _box_in(s['ink'], region, 0.25)
        if ok and (pred is None or pred(s)):
            out.append(s)
    return out


def _short_segs(model, region, maxlen=0.62):
    """区域内的“短线段”数量。复刻类文件常把虚线画成成百上千段实线短段，
       只看 a:prstDash 会漏判，故把短线段数量也作为虚线证据之一。"""
    n = 0
    for s in model['vis']:
        if s['kind'] not in ('line', 'free') or s['has_text']:
            continue
        if s['seglen'] is None or s['seglen'] > maxlen:
            continue
        if _pt_in(s['cx'], s['cy'], region):
            n += 1
    return n


def _dashed_in(model, region, need_dashed=1, need_segs=6):
    """区域内是否存在虚线：① 有描边为虚线族的形状；或 ② 短线段数量达阈值。"""
    nd = len([s for s in model['vis'] if s['dashed'] and _pt_in(s['cx'], s['cy'], region)])
    return nd >= need_dashed or _short_segs(model, region) >= need_segs


def _perim_dashed(model, box, need_segs=8):
    """某个框的“边框虚线”证据：框外沿一圈（外扩0.25cm、内缩0.3cm 的环带）内
       是否有虚线描边形状或足够多的短线段；框本身描边为虚线族时直接成立。"""
    outer = (box[0] - 0.25, box[1] - 0.25, box[2] + 0.25, box[3] + 0.25)
    inner = (box[0] + 0.32, box[1] + 0.32, box[2] - 0.32, box[3] - 0.32)
    n = 0
    for s in model['vis']:
        if s['has_text'] or s['kind'] not in ('line', 'free'):
            continue
        if s['seglen'] is None or s['seglen'] > 0.62:
            continue
        if _pt_in(s['cx'], s['cy'], outer) and not _pt_in(s['cx'], s['cy'], inner):
            n += 1
    return n >= need_segs


def _icons(model, region):
    """图标候选：区域内不带文字的小尺寸矢量形状或独立图片（排除长线段与大面板）。"""
    out = []
    for s in model['vis']:
        if s['has_text'] or s['kind'] == 'line':
            continue
        if not (0.10 <= s['w'] <= 2.40 and 0.10 <= s['h'] <= 2.40):
            continue
        if s['area'] < 0.02:
            continue
        if _pt_in(s['cx'], s['cy'], region):
            out.append(s)
    return out


def _arrows(model, region, minlen=0.25):
    out = []
    for s in model['vis']:
        if s['has_text']:
            continue
        if not s['arrowish'] and not (s['kind'] == 'line' and (s['w'] + s['h']) >= minlen):
            continue
        if _pt_in(s['cx'], s['cy'], region):
            out.append(s)
    return out

def _fig_score(model, keys):
    txt = ''.join(u['text'] for u in model['units'])
    return sum(txt.count(k) for k in keys)


def _region_units(model, region):
    """区域内的全部文本单元（先按行、再按列排序）。"""
    us = [u for u in model['units']
          if _pt_in((u['x0'] + u['x1']) / 2.0, u['cy'], region)]
    us.sort(key=lambda z: (z['cy'], z['x0']))
    return us


def _ubox(us):
    """一组文本单元的外框。"""
    if not us:
        return None
    return (min(u['x0'] for u in us), min(u['y0'] for u in us),
            max(u['x1'] for u in us), max(u['y1'] for u in us))


def _cover_in(model, region, target):
    """区域内文字对目标短语的字符覆盖率。卡片/色块内的短语常被排成两行、
       又被拆成多个逐字文本框，按“视觉行”找不全，故整区域拼接后按重数比对。"""
    us = _region_units(model, region)
    return _cover(_norm(''.join(u['text'] for u in us)), _norm(target))


def _stack_box(cands, core):
    """以 core 为基准，向外并入“中心落在 core 纵向范围附近”的同列构件，
       返回整组构件的并集外框与构件个数。用于把一台由多个零件拼出的
       竖直装置还原成一个整体外框。"""
    lo, hi = core['y0'] - 0.80, core['y1'] + 0.80
    grp = [s for s in cands if lo <= s['cy'] <= hi]
    if core not in grp:
        grp.append(core)
    return ((min(s['x0'] for s in grp), min(s['y0'] for s in grp),
             max(s['x1'] for s in grp), max(s['y1'] for s in grp)), len(grp))



class _TargetMissing(Exception):
    """目录不存在或目录中没有 .pptx：连待评测对象都没有，不进入维度判定。"""


def evaluate(dir_path):
    """按接口约定返回结构化 dict；不 print、不改 sys.stdout、不 sys.exit。
    入参 dir_path 为“脚本所在目录的路径”，脚本自己在该目录里定位并打开被评估的
    .pptx（忽略以 ~$ 开头的 PowerPoint 临时锁文件）。目录不存在 / 未发现文档时
    返回 status='error' 且维度一未参与判定（dim1_pass 留空）；文件存在但打不开
    或脚本崩溃时返回 status='error' 且 dim1_pass=False 并给出错误摘要。"""
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
    try:
        if not dir_path or not os.path.isdir(dir_path):
            raise _TargetMissing(f"目录不存在：{dir_path}")
        cands = sorted(
            f for f in os.listdir(dir_path)
            if f.lower().endswith('.pptx') and not f.startswith('~$')
        )
        if not cands:
            raise _TargetMissing(f"目录中未找到 .pptx 文件：{dir_path}")
        file_path = os.path.join(dir_path, cands[0])
        result["file_name"] = os.path.basename(file_path)
        # 维度一唯一门禁：交付文件必须是能正常打开的 .pptx；打不开即完成度总分记 0。
        try:
            prs = Presentation(file_path)
        except Exception as exc:
            result["status"] = "error"
            result["error"] = f"{type(exc).__name__}: {exc}"
            result["dim1_pass"] = False
            result["dim1_reason"] = f"文件损坏或格式错误，无法作为 .pptx 正常打开（{type(exc).__name__}: {exc}）"
            result["dim2_items"] = []
            result["total_score"] = 0
            return result
        return _evaluate_impl(prs, result)
    except _TargetMissing as exc:
        # 找不到待评测文件：维度1 未参与判定，dim1_pass / dim1_reason 一律留空，满分也置 0
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
        # 文件存在但解析失败：维度1 判不通过，dim1_reason 给出错误摘要
        result["status"] = "error"
        result["error"] = f"{type(exc).__name__}: {exc}"
        result["dim1_pass"] = False
        result["dim1_reason"] = f"文件无法解析或无法打开：{type(exc).__name__}: {exc}"
        result["dim2_items"] = []
        result["total_score"] = 0
        return result


def _evaluate_impl(prs, result):
    """维度一门禁 + 维度二 32 项评分；末尾按 _MAX_DELTAS_ORDER 把 results 反解成
    结构化 dim2_items 填入 result。"""
    global _THEME
    _THEME = _load_theme(prs)
    PW, PH = prs.slide_width / EMU_CM, prs.slide_height / EMU_CM
    slides = list(prs.slides)
    results = []

    # ========== 维度1：可用与可修改性 ==========
    # 门禁只保留“交付文件为1个可正常打开的 .pptx”这一项，且已在 evaluate() 中完成：
    # 打不开即 dim1_pass=False、total_score=0。页数、空白页、整页位图、文字是否固化
    # 在图片内、版面越界等均不再作为门禁，交由维度二对应评分项逐条扣分。
    models = [_build_slide(s, PW, PH) for s in slides]
    if len(models) < 2:
        raise ValueError(f"幻灯片不足2页，无法定位两张框架图（实际{len(models)}页）")

    # ========== 维度2：完成度评分 ==========
    # 页面归属：按页内关键词把两页分别认定为“图2 膜分离页”与“图1 慢病风险页”，
    # 不依赖页序，避免作答把两页顺序调换导致误判。
    _K2 = ('膜', '箱', '污染', '通量', '分离', '压力')
    _K1 = ('风险', '慢病', 'SHAP', 'AUC', '随访', '预测', '生理')
    s2 = [(_fig_score(m, _K2), i) for i, m in enumerate(models)]
    s1 = [(_fig_score(m, _K1), i) for i, m in enumerate(models)]
    i2 = max(s2)[1]
    i1 = max((v, i) for v, i in s1 if i != i2)[1]
    M2, M1 = models[i2], models[i1]
    B2, B1 = M2['bbox'], M1['bbox']

    # +3：文档共2页，两页分别承载“膜分离…”立体箱体框架图与“多源生理信号…”三栏框架图，
    #     一页一图、无多余空白页
    # 严格对齐细则，仅踩细则明确写出的点：
    #   ① 页数恰为 2（多出任何一页即视为多余页，无论其是否空白）；
    #   ② 两页各自承载一张图：按关键词命中度分别认定为“膜分离页”“慢病风险页”，
    #      两页归属必须互不相同且各自关键词命中数≥5；
    #   ③ 每页均有实际内容（可见形状≥20、文本行≥5），排除半空白页。
    _n_slide = len(slides)
    _ok01 = (_n_slide == 2 and i1 != i2
             and _fig_score(M2, _K2) >= 5 and _fig_score(M1, _K1) >= 5
             and min(len(M2['vis']), len(M1['vis'])) >= 20
             and min(len(M2['rows']), len(M1['rows'])) >= 5)
    if _ok01:
        results.append(f"+3：文档共2页、一页一图承载两张框架图、无多余空白页（{_n_slide}页）")
    else:
        results.append(f"+0：文档共2页、一页一图承载两张框架图、无多余空白页（{_n_slide}页）")

    # +5：两页幻灯片页面尺寸均为宽30cm×高20cm（29.9-30.1cm×19.9-20.1cm）
    # 幻灯片尺寸是文档级属性，两页共用；按细则给定的容差区间判定。
    _ok02 = (29.9 <= PW <= 30.1) and (19.9 <= PH <= 20.1)
    if _ok02:
        results.append(f"+5：页面尺寸为30cm×20cm（实测{PW:.2f}×{PH:.2f}cm）")
    else:
        results.append(f"+0：页面尺寸为30cm×20cm（实测{PW:.2f}×{PH:.2f}cm）")

    # +5：两页画面主体均按等比缩放整体置入，画面主体宽高比为1.33-1.37，无横向拉伸或纵向压扁
    #     变形（两页缩放比例一致）
    # 严格对齐细则，仅踩细则明确写出的点：
    #   ① “画面主体”= 该页所有可见元素（含底部图题）的视觉墨迹外框；
    #   ② 宽高比落在 1.33-1.37，考虑版式估算误差放宽 ±0.02（即 1.31-1.39）；
    #   ③ “两页缩放比例一致”：两页主体宽高比之差 ≤0.06。
    _r2 = M2['bw'] / max(M2['bh'], 0.01)
    _r1 = M1['bw'] / max(M1['bh'], 0.01)
    _ok03 = (1.31 <= _r2 <= 1.39) and (1.31 <= _r1 <= 1.39) and abs(_r2 - _r1) <= 0.06
    if _ok03:
        results.append(f"+5：两页主体等比置入、宽高比1.33-1.37且两页一致（实测{_r2:.3f}／{_r1:.3f}）")
    else:
        results.append(f"+0：两页主体等比置入、宽高比1.33-1.37且两页一致（实测{_r2:.3f}／{_r1:.3f}）")

    # +3：两页画面主体在页面内水平且垂直居中，左右留白之差≤0.5cm、上下留白之差≤0.5cm
    def _center_gap(m):
        b = m['bbox']
        return abs(b[0] - (m['pw'] - b[2])), abs(b[1] - (m['ph'] - b[3]))
    _c2 = _center_gap(M2)
    _c1 = _center_gap(M1)
    _ok04 = max(_c2) <= 0.5 and max(_c1) <= 0.5
    if _ok04:
        results.append(f"+3：两页主体水平且垂直居中（左右差{_c2[0]:.2f}/{_c1[0]:.2f}cm，上下差{_c2[1]:.2f}/{_c1[1]:.2f}cm）")
    else:
        results.append(f"+0：两页主体水平且垂直居中（左右差{_c2[0]:.2f}/{_c1[0]:.2f}cm，上下差{_c2[1]:.2f}/{_c1[1]:.2f}cm）")

    # +3：所有元素完整落在页面可视范围内，无元素越出页边被裁切，无与画面无关的散落线条、
    #     色块或延伸到页角的斜线
    # 严格对齐细则，仅踩细则明确写出的点：
    #   ① 越页：任一可见元素的视觉墨迹外框超出页面 >0.3cm（0.3cm 为版式估算容差）；
    #   ② 页角斜线：横纵跨度均≥1.5cm 的斜向线条，其端点落在页面四角 1.2cm 内。
    def _out_cnt(m):
        n = 0
        for s in m['vis']:
            if (s['x0'] < -0.30 or s['y0'] < -0.30
                    or s['x1'] > m['pw'] + 0.30 or s['y1'] > m['ph'] + 0.30):
                n += 1
        return n

    def _corner_line(m):
        cs = ((0.0, 0.0), (m['pw'], 0.0), (0.0, m['ph']), (m['pw'], m['ph']))
        for s in m['vis']:
            if s['kind'] != 'line' or s['has_text'] or s['w'] < 1.5 or s['h'] < 1.5:
                continue
            for px, py in ((s['x0'], s['y0']), (s['x1'], s['y1']),
                           (s['x0'], s['y1']), (s['x1'], s['y0'])):
                if any(math.hypot(px - cx, py - cy) <= 1.2 for cx, cy in cs):
                    return True
        return False
    _o2, _o1 = _out_cnt(M2), _out_cnt(M1)
    _ok05 = (_o2 == 0 and _o1 == 0 and not _corner_line(M2) and not _corner_line(M1))
    if _ok05:
        results.append("+3：所有元素完整落在页面内、无散落线条与页角斜线（越页元素0个）")
    else:
        results.append(f"+0：所有元素完整落在页面内、无散落线条与页角斜线（越页元素{_o2 + _o1}个）")

    # +3：两页底部各设1行居中图题，文字分别为“图2  膜分离…”与“图1  多源生理信号…”，
    #     黑色加粗、单行居中、字号大于卡片正文
    # 严格对齐细则，仅踩细则明确写出的点：
    #   ① 位置：位于页面底部区域（页高 82% 以下）；
    #   ② 文字：按字符覆盖率比对给定图题全文，且必须落在同一视觉行内、未折行；
    #   ③ 黑色加粗：图题各段均加粗，字色为黑/深色（未显式设色按默认黑处理）；
    #   ④ 单行居中：图题行中心与页面水平中线偏差 ≤1.2cm；
    #   ⑤ 字号大于卡片正文：图题字号 > 该页正文字号中位数。
    def _cap_ok(m, cap_text):
        reg = (-1.0, m['ph'] * 0.82, m['pw'] + 1.0, m['ph'] + 1.0)
        h = _find(m['rows'], cap_text, region=reg, min_cover=0.85, slack=4)
        if h is None or h['wrapped']:
            return False, 'none'
        if not h['bold']:
            return False, 'bold'
        cols = h['colors']
        if cols and not all(_is_darkish(c, 0.45) for c in cols):
            return False, 'color'
        if abs((h['x0'] + h['x1']) / 2.0 - m['pw'] / 2.0) > 1.2:
            return False, 'center'
        capset = set(id(u) for u in h['units'])
        body = sorted(u['pt'] for u in m['units'] if id(u) not in capset)
        med = body[len(body) // 2] if body else 0.0
        if h['pt'] <= med:
            return False, 'size'
        return True, 'ok'
    _cap2 = _cap_ok(M2, '图2膜分离水处理系统污染控制与通量提升研究')
    _cap1 = _cap_ok(M1, '图1多源生理信号慢病风险预测与个体化干预研究框架')
    if _cap2[0] and _cap1[0]:
        results.append("+3：两页底部图题文字正确、黑色加粗单行居中且字号大于正文")
    else:
        results.append(f"+0：两页底部图题文字正确、黑色加粗单行居中且字号大于正文（图2页{_cap2[1]}／图1页{_cap1[1]}）")

    # ---------- 图2页：膜分离水处理系统框架 ----------
    _b2w, _b2h = M2['bw'], M2['bh']

    def _rel2(fx0, fy0, fx1, fy1):
        """按图2页主体外框的相对比例取区域，避免绑死绝对坐标。"""
        return (B2[0] + _b2w * fx0, B2[1] + _b2h * fy0,
                B2[0] + _b2w * fx1, B2[1] + _b2h * fy1)

    # +5：以一个开放式立体箱体作为整页承载框架，含浅蓝色正面、右侧斜切侧面、顶部斜切顶面
    #     三个面，三面转折处线条闭合衔接、侧面带浅灰蓝渐变，透视方向统一
    # 严格对齐细则，仅踩细则明确写出的点：
    #   ① 承载框架：大型结构元素（面积≥20cm²）的并集需覆盖主体外框 ≥55% 面积；
    #   ② 三个面：除正面外还需 ≥2 处“斜切面”证据——平行四边形/梯形类形状（面积≥6cm²）
    #      或横纵跨度均≥0.8cm 的斜向线条/折线（两者合计计数）；
    #   ③ 浅蓝：大型结构元素中至少 1 个为浅蓝/浅灰蓝填充。
    _big2 = [s for s in M2['vis'] if s['area'] >= 20.0 and not s['has_text']]
    if _big2:
        _ux0 = min(s['x0'] for s in _big2)
        _uy0 = min(s['y0'] for s in _big2)
        _ux1 = max(s['x1'] for s in _big2)
        _uy1 = max(s['y1'] for s in _big2)
        _cov2 = ((_ux1 - _ux0) * (_uy1 - _uy0)) / max(_b2w * _b2h, 0.01)
    else:
        _cov2 = 0.0
    _slant = [s for s in M2['vis'] if not s['has_text']
              and ((s['prst'] in ('parallelogram', 'trapezoid') and s['area'] >= 6.0)
                   or (s['kind'] in ('line', 'free') and s['w'] >= 0.8 and s['h'] >= 0.8
                       and math.hypot(s['w'], s['h']) >= 1.5))]
    _lightblue = [s for s in _big2
                  if s['fill'] and _is_lightish(s['fill'], 0.80)
                  and s['fill'][2] >= s['fill'][0]]
    _ok07 = (_cov2 >= 0.55 and len(_slant) >= 2 and len(_lightblue) >= 1)
    _d07 = f"承载框架覆盖{_cov2 * 100:.0f}%、斜切面证据{len(_slant)}处、浅蓝面{len(_lightblue)}个"
    results.append(("+5：" if _ok07 else "+0：")
                   + f"开放式立体箱体含浅蓝正面与右侧、顶部斜切面且转折衔接（{_d07}）")

    # +3：箱体顶面自左至右排布4个同向倾斜的平行四边形标签，文字依次为“运行系统”“膜材料”
    #     “污染层生长”“处理系统”，标签大小相仿、间距均匀，文字随标签排布且不越框
    # 严格对齐细则，仅踩细则明确写出的点：
    #   ① 4 段文字均命中且自左至右次序正确（位于主体外框上部 20% 带内）；
    #   ② 每段文字有一个承载它的倾斜标签形状（非矩形：平行四边形/梯形/自由形/切角形），
    #      文字外框须落在标签内（0.25cm 容差）——即“文字不越框”；
    #   ③ 大小相仿：4 个标签宽度 max/min ≤1.8；
    #   ④ 间距均匀：相邻标签中心距与均值偏差 ≤25%。
    _TOP2 = _rel2(-0.05, -0.03, 1.05, 0.22)
    _lb08 = ('运行系统', '膜材料', '污染层生长', '处理系统')
    _h08 = [_find(M2['rows'], t, region=_TOP2, min_cover=1.0, slack=2) for t in _lb08]
    _ok08 = all(h is not None for h in _h08)
    _hosts08 = []
    if _ok08:
        _ok08 = all(_h08[i]['x0'] <= _h08[i + 1]['x0'] for i in range(3))
    if _ok08:
        for h in _h08:
            hb = (h['x0'], h['y0'], h['x1'], h['y1'])
            cand = [s for s in M2['vis'] if not s['has_text'] and s['prst'] != 'rect'
                    and s['area'] >= 1.5 and _box_in(hb, s['ink'], 0.25)]
            _hosts08.append(min(cand, key=lambda z: z['area']) if cand else None)
        _ok08 = all(z is not None for z in _hosts08)
    if _ok08:
        _ws = [z['w'] for z in _hosts08]
        _cs = sorted(z['cx'] for z in _hosts08)
        _gaps = [_cs[i + 1] - _cs[i] for i in range(3)]
        _mg = sum(_gaps) / 3.0
        _wr = max(_ws) / max(min(_ws), 0.01)
        _dev = max(abs(g - _mg) for g in _gaps) / max(_mg, 0.01)
        _ok08 = (_wr <= 1.8 and _dev <= 0.25)
        _d08 = f"宽度比{_wr:.2f}、间距偏差{_dev * 100:.0f}%"
    else:
        _d08 = f"命中标签{sum(1 for h in _h08 if h)}/4、倾斜承载框{sum(1 for z in _hosts08 if z)}个"
    results.append(("+3：" if _ok08 else "+0：")
                   + f"箱体顶面4个同向倾斜标签文字正确、大小相仿间距均匀且文字不越框（{_d08}）")

    # +5：箱体左侧设一条竖向窄栏，栏内用2条水平虚线均匀分隔为上中下3格，3格自上而下依次为
    #     “压力表”图标+“跨膜压差”、“水滴分子”图标+“污染迁移速率”、“多孔膜块”图标+“传质阻力”，
    #     每格图标在上、文字在下并居中且单行不换行，3格之间另设2个竖直向下的箭头
    # 严格对齐细则，仅踩细则明确写出的点：
    #   ① 3 段文字均命中且自上而下次序正确、均未折行（单行）；
    #   ② 图标在上：每段文字上方 2.6cm 内有一个与之横向重叠的非文本图形；
    #   ③ 居中：3 段文字中心横坐标彼此偏差 ≤0.8cm（同一竖栏居中排布）；
    #   ④ 2 条水平虚线：栏内可聚成 ≥2 个横向虚线层（显式虚线，或同一高度上 ≥4 段短线且
    #      横向跨度≥1.0cm 的“碎段虚线”画法）；
    #   ⑤ 2 个竖直向下箭头：栏内 ≥2 个竖长（高/宽≥1.4、高≥0.6cm）的非文本箭头类图形。
    _LB2 = (B2[0] - 0.30, B2[1] - 0.30, B2[0] + 0.16 * _b2w, B2[3] + 0.30)
    _lb09 = ('跨膜压差', '污染迁移速率', '传质阻力')
    _h09 = [_find(M2['rows'], t, region=_LB2, min_cover=1.0, slack=2) for t in _lb09]
    _n09 = sum(1 for h in _h09 if h)
    _ic09 = 0
    _bar09 = _arw09 = 0
    _ok09 = (_n09 == 3)
    if _ok09:
        _ok09 = all(_h09[i]['y0'] <= _h09[i + 1]['y0'] for i in range(2)) \
            and not any(h['wrapped'] for h in _h09)
    if _ok09:
        _mcx = sum((h['x0'] + h['x1']) / 2.0 for h in _h09) / 3.0
        _ok09 = all(abs((h['x0'] + h['x1']) / 2.0 - _mcx) <= 0.8 for h in _h09)
    if _ok09:
        for h in _h09:
            for s in M2['vis']:
                if s['has_text'] or not _pt_in(s['cx'], s['cy'], _LB2):
                    continue
                if s['area'] < 0.08 or s['w'] > 3.0 or s['h'] > 3.0:
                    continue
                if 0.02 <= h['y0'] - s['cy'] <= 2.60 \
                        and min(s['x1'], h['x1']) - max(s['x0'], h['x0']) >= 0.20:
                    _ic09 += 1
                    break
        _ok09 = (_ic09 == 3)
    if _ok09:
        _segs = [s for s in M2['vis'] if not s['has_text'] and s['kind'] in ('line', 'free')
                 and _pt_in(s['cx'], s['cy'], _LB2) and (s['seglen'] or 0) <= 0.62
                 and s['h'] <= 0.30]
        _lv = {}
        for s in _segs:
            _lv.setdefault(round(s['cy'] / 0.25), []).append(s)
        _bar09 = len([g for g in _lv.values()
                      if len(g) >= 4 and (max(x['x1'] for x in g) - min(x['x0'] for x in g)) >= 1.0])
        _bar09 += len([s for s in M2['vis'] if s['dashed'] and not s['has_text']
                       and _pt_in(s['cx'], s['cy'], _LB2) and s['w'] >= 1.0 and s['h'] <= 0.30])
        _arw09 = len([s for s in M2['vis'] if not s['has_text']
                      and _pt_in(s['cx'], s['cy'], _LB2) and s['h'] >= 0.60 and s['w'] <= 1.20
                      and s['h'] >= 1.40 * max(s['w'], 0.01)
                      and (s['arrowish'] or (s['kind'] in ('free', 'line') and s['w'] >= 0.08))])
        _ok09 = (_bar09 >= 2 and _arw09 >= 2)
    _d09 = f"命中文字{_n09}/3、图标{_ic09}/3、横向虚线层{_bar09}、竖向箭头{_arw09}"
    results.append(("+5：" if _ok09 else "+0：")
                   + f"左侧竖栏2条水平虚线分3格、每格图标在上文字在下居中且格间2个向下箭头（{_d09}）")

    # +3：画面上部横排4个等大圆角矩形框，文字依次为“膜组件”“污染沉积层”“表面相互作用”
    #     “溶液性质”，4框顶边对齐、框内文字单行不换行，相邻框之间用3个实线右向箭头连接
    # 严格对齐细则，仅踩细则明确写出的点：
    #   ① 4 段文字均命中、自左至右次序正确、均未折行（框内文字单行不换行）；
    #   ② 每段文字有承载它的圆角矩形（文字外框整体落在其中，取面积最小者为宿主）；
    #   ③ 顶边对齐：4 框上沿最大差 ≤0.25cm；等大：高度 max/min ≤1.35、下沿差 ≤0.25cm，
    #      宽度 max/min ≤1.55（标签框宽度常随文字长度微调，故宽度容差放宽）；
    #   ④ 3 个右向箭头：相邻两框的水平间隙内各有 ≥1 个横向元素（宽≥0.40cm 且宽≥1.2×高，
    #      箭头常被画成实心三角自由形，故不强求 a:tailEnd/headEnd 标记）。
    _lb10 = ('膜组件', '污染沉积层', '表面相互作用', '溶液性质')
    _h10 = [_find(M2['rows'], t, min_cover=1.0, slack=2) for t in _lb10]
    _n10 = sum(1 for h in _h10 if h)
    _hs10, _con10 = [], 0
    if _n10 == 4:
        for h in _h10:
            hb = (h['x0'], h['y0'], h['x1'], h['y1'])
            cand = [s for s in M2['vis'] if not s['has_text'] and s['area'] >= 1.5
                    and ('round' in s['prst'] or s['prst'] == 'rect')
                    and _box_in(hb, s['ink'], 0.30)]
            if cand:
                _hs10.append(min(cand, key=lambda z: z['area']))
        _hs10.sort(key=lambda z: z['x0'])
    _ok10 = (_n10 == 4) and not any(h['wrapped'] for h in _h10) \
        and all(_h10[i]['x0'] <= _h10[i + 1]['x0'] for i in range(3)) and len(_hs10) == 4
    if _ok10:
        _ws10 = [z['w'] for z in _hs10]
        _hh10 = [z['h'] for z in _hs10]
        _top10 = max(z['y0'] for z in _hs10) - min(z['y0'] for z in _hs10)
        _bot10 = max(z['y1'] for z in _hs10) - min(z['y1'] for z in _hs10)
        _wr10 = max(_ws10) / max(min(_ws10), 0.01)
        _hr10 = max(_hh10) / max(min(_hh10), 0.01)
        for i in range(3):
            _g = (_hs10[i]['x1'], min(z['y0'] for z in _hs10) - 0.35,
                  _hs10[i + 1]['x0'], max(z['y1'] for z in _hs10) + 0.35)
            if [s for s in M2['vis'] if not s['has_text'] and _pt_in(s['cx'], s['cy'], _g)
                    and s['w'] >= 0.40 and s['w'] >= 1.2 * s['h']]:
                _con10 += 1
        _ok10 = (_top10 <= 0.25 and _bot10 <= 0.25 and _wr10 <= 1.55
                 and _hr10 <= 1.35 and _con10 == 3)
        _d10 = f"顶边差{_top10:.2f}cm、宽度比{_wr10:.2f}、高度比{_hr10:.2f}、连接箭头{_con10}/3"
    else:
        _d10 = f"命中文字{_n10}/4、圆角矩形宿主{len(_hs10)}个"
    results.append(("+3：" if _ok10 else "+0：")
                   + f"上部横排4个等大圆角矩形框文字正确单行、顶边对齐且框间3个右向箭头（{_d10}）")

    # +3：上述4个框各自向下引出虚线折线并汇聚指向1个浅蓝填充圆角矩形，其内文字为
    #     “膜污染演化是影响通量衰减的重要因素之一”，单行居中显示
    # 严格对齐细则，仅踩细则明确写出的点：
    #   ① 该长句命中且未折行（单行显示）；
    #   ② 承载它的形状为浅蓝填充（蓝分量≥红分量且整体偏亮）圆角矩形，
    #      文字与框水平居中（中心横坐标差 ≤0.7cm）；
    #   ③ 汇聚虚线：4 框底沿与浅蓝框顶沿之间的横带内存在虚线证据
    #      （显式虚线形状 ≥1，或“碎段虚线”短线段 ≥8）。
    _t11 = '膜污染演化是影响通量衰减的重要因素之一'
    _h11 = _find(M2['rows'], _t11, min_cover=0.85, slack=6)
    _bx11 = None
    _ok11 = (_h11 is not None) and not _h11['wrapped']
    if _ok11:
        _hb11 = (_h11['x0'], _h11['y0'], _h11['x1'], _h11['y1'])
        _c11 = [s for s in M2['vis'] if not s['has_text'] and s['area'] >= 1.5
                and 'round' in s['prst'] and _box_in(_hb11, s['ink'], 0.30)
                and s['fill'] and _is_lightish(s['fill'], 0.80) and s['fill'][2] >= s['fill'][0]]
        _bx11 = min(_c11, key=lambda z: z['area']) if _c11 else None
        _ok11 = _bx11 is not None
    if _ok11:
        _ok11 = abs((_h11['x0'] + _h11['x1']) / 2.0 - _bx11['cx']) <= 0.70
    _dash11 = False
    if _ok11 and len(_hs10) == 4:
        _bd11 = (min(z['x0'] for z in _hs10) - 0.60, max(z['y1'] for z in _hs10) + 0.02,
                 max(z['x1'] for z in _hs10) + 0.60, _bx11['y0'] - 0.02)
        _dash11 = _dashed_in(M2, _bd11, need_dashed=1, need_segs=8)
        _ok11 = _dash11
    elif _ok11:
        _ok11 = False
    _d11 = ("浅蓝圆角矩形%s、单行%s、居中%s、汇聚虚线%s"
            % ('有' if _bx11 else '无', '是' if (_h11 and not _h11['wrapped']) else '否',
               '是' if (_bx11 and _h11 and abs((_h11['x0'] + _h11['x1']) / 2.0 - _bx11['cx']) <= 0.70) else '否',
               '有' if _dash11 else '无'))
    results.append(("+3：" if _ok11 else "+0：")
                   + f"4框向下引出虚线折线汇聚指向浅蓝填充圆角矩形且其内文字单行居中（{_d11}）")

    # +3：该浅蓝框下方用竖直实线箭头连接1个橙色虚线边框、浅橙填充的圆角矩形；框内为两行文字，
    #     上行“精细调控膜表面结构与运行参数”为深色常规，下行“优化污染控制状态”为橙色加粗
    #     且字号明显大于上行
    # 严格对齐细则，仅踩细则明确写出的点：
    #   ① 两行文字均命中，承载框为浅橙填充（r>g>b 且 r-b≥6）圆角矩形；
    #   ② 竖直实线箭头：浅蓝框底沿与该框顶沿之间有竖向元素（高≥0.35cm 且高≥1.2×宽）；
    #   ③ 边框虚线：该框描边为虚线族，或框边一圈存在“碎段虚线”短线段 ≥8；
    #   ④ 下行橙色加粗且字号明显大于上行：下行字号中位数 ≥1.15×上行、加粗、
    #      文字颜色满足 r>g>b 且 r-b≥40。
    _u12 = _find(M2['rows'], '精细调控膜表面结构与运行参数', min_cover=0.85, slack=6)
    _l12 = _find(M2['rows'], '优化污染控制状态', min_cover=0.85, slack=6)
    _bx12 = None
    _ok12 = (_u12 is not None and _l12 is not None)
    if _ok12:
        _hb12 = (min(_u12['x0'], _l12['x0']), min(_u12['y0'], _l12['y0']),
                 max(_u12['x1'], _l12['x1']), max(_u12['y1'], _l12['y1']))
        _c12 = [s for s in M2['vis'] if not s['has_text'] and s['area'] >= 1.5
                and 'round' in s['prst'] and _box_in(_hb12, s['ink'], 0.35)
                and s['fill'] and s['fill'][0] > s['fill'][1] > s['fill'][2]
                and s['fill'][0] - s['fill'][2] >= 6]
        _bx12 = min(_c12, key=lambda z: z['area']) if _c12 else None
        _ok12 = _bx12 is not None
    _arw12 = _dsh12 = _fnt12 = False
    if _ok12:
        _ab12 = (_bx12['x0'] - 0.60, _bx12['y0'] - 1.70, _bx12['x1'] + 0.60, _bx12['y0'] - 0.02)
        _arw12 = bool([s for s in M2['vis'] if not s['has_text'] and _pt_in(s['cx'], s['cy'], _ab12)
                       and s['h'] >= 0.35 and s['h'] >= 1.2 * s['w']])
        _dsh12 = _bx12['dashed'] or _perim_dashed(M2, _bx12['ink'], need_segs=8)
        _ui12 = _region_units(M2, (_bx12['x0'], _bx12['y0'], _bx12['x1'], _bx12['cy']))
        _li12 = _region_units(M2, (_bx12['x0'], _bx12['cy'], _bx12['x1'], _bx12['y1']))
        if _ui12 and _li12:
            _pu = sorted(u['pt'] for u in _ui12)[len(_ui12) // 2]
            _pl = sorted(u['pt'] for u in _li12)[len(_li12) // 2]
            _cl = [u['color'] for u in _li12 if u['color']]
            _fnt12 = (_pl >= 1.15 * _pu and all(u['bold'] for u in _li12)
                      and bool(_cl) and all(c[0] > c[1] > c[2] and c[0] - c[2] >= 40 for c in _cl))
            _d12 = f"上行{_pu:.0f}pt／下行{_pl:.0f}pt"
        else:
            _d12 = "两行文字未分离"
        _ok12 = (_arw12 and _dsh12 and _fnt12)
        _d12 += f"、竖直箭头{'有' if _arw12 else '无'}、边框虚线{'是' if _dsh12 else '否'}"
    else:
        _d12 = f"命中两行文字{sum(1 for h in (_u12, _l12) if h)}/2、浅橙圆角矩形{'有' if _bx12 else '无'}"
    results.append(("+3：" if _ok12 else "+0：")
                   + f"浅蓝框下经竖直箭头接橙色虚线边浅橙框且下行橙色加粗字号更大（{_d12}）")

    # +3：橙色虚线框底部分出左右两条带箭头折线，分别指向2个并列的浅蓝边框圆角面板，
    #     面板标题分别为“界面相容优化”（左）与“运行抑污控制”（右），标题蓝色加粗、居面板顶部居中
    # 严格对齐细则，仅踩细则明确写出的点：
    #   ① 两个标题均命中且左右次序正确，均为加粗、蓝色系（b>r 且 b-r≥25）；
    #   ② 每个标题有承载它的圆角面板（宽≥4cm、面积≥15cm²），左右并列不同框；
    #   ③ 居面板顶部居中：标题中心与面板中心横坐标差 ≤0.9cm，
    #      标题中心距面板上沿 ≤面板高度的 35%；
    #   ④ 左右两条带箭头折线：橙色框底沿与面板上沿之间的横带内 ≥2 个线条/折线元素。
    _t13 = ('界面相容优化', '运行抑污控制')
    _h13 = [_find(M2['rows'], t, min_cover=1.0, slack=2) for t in _t13]
    _pn13 = []
    _ok13 = all(h is not None for h in _h13)
    if _ok13:
        _ok13 = _h13[0]['x0'] < _h13[1]['x0'] and all(h['bold'] for h in _h13) \
            and all(h['colors'] and all(c[2] > c[0] and c[2] - c[0] >= 25 for c in h['colors'])
                    for h in _h13)
    if _ok13:
        for h in _h13:
            hb = (h['x0'], h['y0'], h['x1'], h['y1'])
            cand = [s for s in M2['vis'] if not s['has_text'] and s['area'] >= 15.0
                    and s['w'] >= 4.0 and 'round' in s['prst'] and _box_in(hb, s['ink'], 0.45)]
            _pn13.append(min(cand, key=lambda z: z['area']) if cand else None)
        _ok13 = (all(z is not None for z in _pn13)
                 and _pn13[0]['sid'] != _pn13[1]['sid'] and _pn13[0]['x0'] < _pn13[1]['x0'])
    _br13 = 0
    if _ok13:
        for h, p in zip(_h13, _pn13):
            if abs((h['x0'] + h['x1']) / 2.0 - p['cx']) > 0.90:
                _ok13 = False
            if (h['y0'] + h['y1']) / 2.0 - p['y0'] > 0.35 * p['h']:
                _ok13 = False
    if _ok13:
        _bb13 = (_pn13[0]['x0'], (_bx12['y1'] if _bx12 else _pn13[0]['y0'] - 1.5) + 0.02,
                 _pn13[1]['x1'], min(p['y0'] for p in _pn13) - 0.02)
        _br13 = len([s for s in M2['vis'] if not s['has_text']
                     and s['kind'] in ('line', 'free') and _pt_in(s['cx'], s['cy'], _bb13)])
        _ok13 = (_br13 >= 2)
    _d13 = f"命中标题{sum(1 for h in _h13 if h)}/2、面板{sum(1 for z in _pn13 if z)}个、折线元素{_br13}"
    results.append(("+3：" if _ok13 else "+0：")
                   + f"橙色框底分出左右带箭头折线指向2个并列面板且标题蓝色加粗顶部居中（{_d13}）")

    # +5：左右两个面板内各为2行2列共4张小卡片，文字依次为指定4项，每张卡片左侧配1个
    #     对应图标，图标与文字左右并排且互不压盖（左右面板同时达标才给分）
    # 严格对齐细则，仅踩细则明确写出的点：
    #   ① 4 项文字均在面板内命中（卡片文字常被排成两行又拆成逐字文本框，
    #      故按“整面板区域拼接后按字符重数比对”，不按单个视觉行比对）；
    #   ② 2行2列：面板内取面积最大的 4 张卡片，其中心须聚成 2 个横坐标档位与
    #      2 个纵坐标档位（同档差 ≤0.60cm、跨档差 ≥1.5cm／≥0.8cm）；
    #   ③ 每张卡片内 ≥1 个图标类元素（非文本、非长线条、面积 0.02-2.0cm²），
    #      且位于该卡文字左侧（图标中心横坐标 < 文字中心）、与文字横向压盖 ≤图标宽度的 60%。
    def _panel_cards(pan, labels):
        if pan is None:
            return False, "面板未识别"
        inner = (pan['x0'] - 0.10, pan['y0'] + 0.35, pan['x1'] + 0.10, pan['y1'] + 0.10)
        _nt = sum(1 for t in labels if _cover_in(M2, inner, t) >= 0.99)
        cards = [s for s in M2['vis'] if not s['has_text'] and 1.5 <= s['area'] <= 14.0
                 and _pt_in(s['cx'], s['cy'], inner) and s['w'] >= 1.5 and s['h'] >= 0.8]
        cards.sort(key=lambda z: -z['area'])
        cards = cards[:4]
        if _nt < 4 or len(cards) < 4:
            return False, f"命中文字{_nt}/4、卡片{len(cards)}张"
        _xs = sorted(z['cx'] for z in cards)
        _ys = sorted(z['cy'] for z in cards)
        _grid = (abs(_xs[0] - _xs[1]) <= 0.60 and abs(_xs[2] - _xs[3]) <= 0.60
                 and _xs[2] - _xs[1] >= 1.5 and abs(_ys[0] - _ys[1]) <= 0.60
                 and abs(_ys[2] - _ys[3]) <= 0.60 and _ys[2] - _ys[1] >= 0.80)
        _nic = 0
        for c in cards:
            cb = (c['x0'] - 0.05, c['y0'] - 0.05, c['x1'] + 0.05, c['y1'] + 0.05)
            tb = _ubox(_region_units(M2, cb))
            if tb is None:
                continue
            for s in M2['vis']:
                if s['has_text'] or s['kind'] == 'line' or not (0.02 <= s['area'] <= 2.0):
                    continue
                if not _pt_in(s['cx'], s['cy'], cb):
                    continue
                _ov = max(0.0, min(s['x1'], tb[2]) - max(s['x0'], tb[0]))
                if s['cx'] < (tb[0] + tb[2]) / 2.0 and _ov <= 0.60 * max(s['w'], 0.01):
                    _nic += 1
                    break
        return (_grid and _nic == 4), f"命中文字{_nt}/4、2行2列{'是' if _grid else '否'}、配图标{_nic}/4"

    _p14 = _pn13[0] if len(_pn13) == 2 and _pn13[0] is not None else None
    _p15 = _pn13[1] if len(_pn13) == 2 and _pn13[1] is not None else None
    _okL14, _dL14 = _panel_cards(_p14, ('增强表面亲水性', '改进孔结构', '膜基材', '表面涂层'))
    _okR14, _dR14 = _panel_cards(_p15, ('抑制浓差极化', '降低不可逆污染', '气水反冲', '运行参数'))
    _ok14 = _okL14 and _okR14
    results.append(("+5：" if _ok14 else "+0：")
                   + "左右两个面板内各2行2列4张小卡片文字正确且每卡左侧配图标不压盖文字"
                   + f"（左：{_dL14}；右：{_dR14}）")

    # +3：两个面板底部各向下引出竖线并汇合成折线，指向1个圆角矩形，框内单行文字为
    #     “通量衰减减缓 · 清洗周期延长 · 分离效率提升”，三段文字以间隔点分隔并水平均匀分布
    # 严格对齐细则，仅踩细则明确写出的点：
    #   ① 面板下方存在一个圆角矩形，框内文字同时覆盖三段短语（整框拼接后按字符重数比对，
    #      因为三段可能写在同一个文本对象里，按视觉行分段会切不开）；
    #   ② 单行：框内文字单元的纵向中心差 ≤0.35cm，且均未折行；
    #   ③ 间隔点分隔：框内文字含 ≥2 个间隔点，或两段间空隙内各有 1 个小圆点图形；
    #   ④ 水平均匀分布：三段各为独立文本段时，两处段间空隙之差 ≤较大者的 45%；
    #      三段写在同一个文本对象内时，字符流已决定其等距关系，只需整行在框内居中
    #      （中心横坐标差 ≤0.6cm）即视为均匀；
    #   ⑤ 汇合折线：面板下沿与该框上沿之间 ≥2 个线条/折线元素。
    _t16 = ('通量衰减减缓', '清洗周期延长', '分离效率提升')
    _y16 = max(p['y1'] for p in _pn13 if p is not None) if any(_pn13) else B2[1] + 0.7 * _b2h
    _bx16, _dot16, _cn16, _uni16 = None, False, 0, False
    _c16 = [s for s in M2['vis'] if not s['has_text'] and s['area'] >= 4.0
            and 'round' in s['prst'] and s['cy'] >= _y16 - 0.30
            and all(_cover_in(M2, s['ink'], t) >= 0.99 for t in _t16)]
    _bx16 = min(_c16, key=lambda z: z['area']) if _c16 else None
    _ok16 = _bx16 is not None
    if _ok16:
        _us16 = _region_units(M2, _bx16['ink'])
        _cy16 = [u['cy'] for u in _us16]
        _ok16 = (max(_cy16) - min(_cy16) <= 0.35) and not any(u['lines'] > 1 for u in _us16)
    if _ok16:
        _tb16 = _ubox(_us16)
        _sg16 = []
        for u in sorted(_us16, key=lambda z: z['x0']):
            if _sg16 and u['x0'] - _sg16[-1][1] <= 0.42:
                _sg16[-1][1] = max(_sg16[-1][1], u['x1'])
            else:
                _sg16.append([u['x0'], u['x1']])
        _dot16 = _norm(''.join(u['text'] for u in _us16)).count('·') >= 2
        if not _dot16 and len(_sg16) >= 3:
            _dot16 = all([s for s in M2['vis'] if not s['has_text'] and s['area'] <= 0.30
                          and _pt_in(s['cx'], s['cy'], (_sg16[i][1], _bx16['y0'],
                                                        _sg16[i + 1][0], _bx16['y1']))]
                         for i in (0, 1))
        if len(_sg16) >= 3:
            _g16 = [_sg16[1][0] - _sg16[0][1], _sg16[2][0] - _sg16[1][1]]
            _uni16 = abs(_g16[0] - _g16[1]) / max(max(_g16), 0.01) <= 0.45
            _du16 = f"段间隙偏差{abs(_g16[0] - _g16[1]) / max(max(_g16), 0.01) * 100:.0f}%"
        else:
            _uni16 = abs((_tb16[0] + _tb16[2]) / 2.0 - _bx16['cx']) <= 0.60
            _du16 = f"整行居中差{abs((_tb16[0] + _tb16[2]) / 2.0 - _bx16['cx']):.2f}cm"
        _cn16 = len([s for s in M2['vis'] if not s['has_text'] and s['kind'] in ('line', 'free')
                     and _pt_in(s['cx'], s['cy'], (_bx16['x0'] - 0.5, _y16 - 0.05,
                                                   _bx16['x1'] + 0.5, _bx16['y0'] + 0.05))])
        _ok16 = (_dot16 and _uni16 and _cn16 >= 2)
        _d16 = f"间隔点{'有' if _dot16 else '无'}、{_du16}、汇合折线{_cn16}"
    else:
        _d16 = "未找到含三段文字的圆角矩形"
    results.append(("+3：" if _ok16 else "+0：")
                   + f"两面板底部竖线汇合折线指向圆角矩形且框内三段文字以间隔点均匀分隔（{_d16}）")

    # +5：箱体右侧内部放置1个立式膜组件压力容器装置图形，含顶部压力表、上下法兰端盖、
    #     筒体内竖向纤维束纹理、侧向进出水接管与底部支脚，装置高度占箱体正面高度的
    #     0.55-0.85且完整落在箱体范围内
    # 严格对齐细则，仅踩细则明确写出的点：
    #   ① 箱体正面高度取整页最高的大型结构元素（面积≥20cm²）高度；
    #   ② 装置：箱体右侧 30% 区域内，以“最高的窄长构件”为核心，并入中心落在其纵向范围
    #      附近（上下各 0.8cm 内）的同列构件，得到装置整体外框；
    #   ③ 立式且多构件：整体外框高/宽 ≥2.0，构件数 ≥3（压力表、端盖、筒体纹理、接管、
    #      支脚等分件绘制）；
    #   ④ 高度占比 0.50-0.90（细则给 0.55-0.85，因“箱体正面高度”只能由最高结构件近似
    #      量得，故两端各留 0.05 容差）；⑤ 整体外框完整落在主体外框内。
    _fh17 = max([s['h'] for s in _big2] or [_b2h])
    _cd17 = [s for s in M2['vis'] if not s['has_text'] and s['kind'] != 'line'
             and s['area'] >= 0.05 and s['w'] <= 3.60 and s['h'] <= 0.90 * _fh17
             and s['cx'] >= B2[0] + 0.70 * _b2w]
    _core17 = None
    for s in _cd17:
        if s['h'] >= 1.50 * max(s['w'], 0.01) and (_core17 is None or s['h'] > _core17['h']):
            _core17 = s
    _bb17, _nm17, _hr17 = None, 0, 0.0
    _ok17 = _core17 is not None
    if _ok17:
        _bb17, _nm17 = _stack_box(_cd17, _core17)
        _hh17 = _bb17[3] - _bb17[1]
        _ww17 = max(_bb17[2] - _bb17[0], 0.01)
        _hr17 = _hh17 / max(_fh17, 0.01)
        _ok17 = (_hh17 / _ww17 >= 2.0 and _nm17 >= 3 and 0.50 <= _hr17 <= 0.90
                 and _box_in(_bb17, B2, 0.30))
        _d17 = f"装置高宽比{_hh17 / _ww17:.2f}、构件{_nm17}个、高度占比{_hr17:.2f}"
    else:
        _d17 = "未在箱体右侧找到竖直装置构件"
    results.append(("+5：" if _ok17 else "+0：")
                   + f"箱体右侧内部立式膜组件压力容器装置分件绘制且高度占正面0.55-0.85（{_d17}）")

    # +5：箱体右侧外部设1列红色加粗竖排大字，自上而下逐字为“降低膜污染并提升分离效率”共12字，
    #     字号明显大于画面内正文，整列左右居中对齐、字间距均匀
    # 严格对齐细则，仅踩细则明确写出的点：
    #   ① 竖排逐字：主体外框右侧 18% 带内的文本单元自上而下拼接后须等于该 12 字
    #      （按字符重数覆盖 =1.0，且单元数 ≥12——每字一个文本对象即为“逐字”竖排）；
    #   ② 红色加粗：全部单元加粗，且有颜色的单元均为红色系（r>g 且 r>b 且 r-g≥60）；
    #   ③ 整列左右居中对齐：各字中心横坐标极差 ≤0.35cm；
    #   ④ 字间距均匀：相邻字中心纵距与均值偏差 ≤20%；
    #   ⑤ 字号明显大于正文：该列字号中位数 ≥1.25×画面其余文字字号中位数。
    _xr18 = B2[2] - 0.18 * _b2w
    _us18 = [u for u in M2['units'] if (u['x0'] + u['x1']) / 2.0 >= _xr18]
    _us18.sort(key=lambda z: z['cy'])
    _bd18 = sorted(u['pt'] for u in M2['units'] if (u['x0'] + u['x1']) / 2.0 < _xr18)
    _t18 = '降低膜污染并提升分离效率'
    _ok18 = (len(_us18) >= 12 and bool(_bd18)
             and _cover(_norm(''.join(u['text'] for u in _us18)), _t18) >= 1.0)
    _cx18 = _gp18 = _rt18 = 0.0
    if _ok18:
        _cs18 = [(u['x0'] + u['x1']) / 2.0 for u in _us18]
        _cx18 = max(_cs18) - min(_cs18)
        _gs18 = [_us18[i + 1]['cy'] - _us18[i]['cy'] for i in range(len(_us18) - 1)]
        _mg18 = sum(_gs18) / max(len(_gs18), 1)
        _gp18 = max(abs(g - _mg18) for g in _gs18) / max(_mg18, 0.01)
        _pc18 = sorted(u['pt'] for u in _us18)[len(_us18) // 2]
        _rt18 = _pc18 / max(_bd18[len(_bd18) // 2], 0.01)
        _cl18 = [u['color'] for u in _us18 if u['color']]
        _ok18 = (all(u['bold'] for u in _us18)
                 and bool(_cl18)
                 and all(c[0] > c[1] and c[0] > c[2] and c[0] - c[1] >= 60 for c in _cl18)
                 and _cx18 <= 0.35 and _gp18 <= 0.20 and _rt18 >= 1.25)
        _d18 = f"共{len(_us18)}字、居中极差{_cx18:.2f}cm、字距偏差{_gp18 * 100:.0f}%、字号为正文{_rt18:.2f}倍"
    else:
        _d18 = f"右侧竖排文字单元{len(_us18)}个、文字未完整命中"
    results.append(("+5：" if _ok18 else "+0：")
                   + f"箱体右侧外部红色加粗竖排12字、字号明显大于正文且居中等距（{_d18}）")

    # ---------- 图1页公共判定件：文字宿主图形、按原始RGB的色系判定 ----------
    # 说明：浅色/低饱和填充用 HSV 归类会大量落到 white/gray（如浅绿 (235,244,232)），
    #       故色系判定统一改为通道关系判定；色系名按细则中出现的说法命名。
    def _host1(h, minarea=1.0, maxarea=None, pred=None):
        """承载某段文字的最小图形（该段文字墨迹被其覆盖）。"""
        if h is None:
            return None
        hb = (h['x0'], h['y0'], h['x1'], h['y1'])
        c = [s for s in M1['vis'] if not s['has_text'] and s['area'] >= minarea
             and (maxarea is None or s['area'] <= maxarea)
             and _box_in(hb, s['ink'], 0.35) and (pred is None or pred(s))]
        c.sort(key=lambda z: z['area'])
        return c[0] if c else None

    def _fam(c, name):
        if c is None:
            return False
        r, g, b = c
        if name == 'cyan':
            return g > r + 25 and b > r + 15 and abs(g - b) <= 30
        if name == 'green':
            return g > r + 20 and g >= b - 15
        if name == 'blue':
            return b > r + 40 and b > g + 25
        if name == 'purple':
            return b > g + 30 and r > g + 10
        if name == 'orange':
            return r > g + 35 and g > b + 35
        if name == 'yellow':
            return r > b + 40 and g > b + 30 and abs(r - g) <= 45
        if name == 'red':
            return r > g + 40 and r > b + 40 and not (g > b + 35)
        return False

    # +5：页面顶部横排4个同向右指的箭头形色带，填充色自左至右依次为青绿、蓝、紫、橙，带内
    #     文字依次为“数据获取与治理”“特征提取与建模”“模型评估与解释”“干预建议生成”，白色
    #     加粗单行显示不越框；每条色带左端配1个圆形底矢量图标，相邻色带之间用同色系小箭头衔接
    # 严格对齐细则，仅踩细则明确写出的点：
    #   ① 4 段文字须全部命中，各自的承载图形（面积≥3cm²）互不相同且自左至右排列；
    #   ② 4 条色带填充依次为青绿/蓝/紫/橙（按通道关系判定）；
    #   ③ 白色加粗单行不越框：加粗、未折行、文字墨迹落在色带内；字色若能取到须为白色系
    #      （复刻件常把白字写成主题色 bg1，取不到具体色值时不据此扣分）；
    #   ④ 左端圆形底矢量图标：色带左端方形区域内的圆形里，须另有非文字矢量构件（用符号
    #      字符充当图标时圆内没有任何构件，本条不成立）；
    #   ⑤ 相邻色带之间的 3 个间隙各需 1 个带箭头构件，且其填充或描边与相邻两条色带之一同色系。
    _t19 = ('数据获取与治理', '特征提取与建模', '模型评估与解释', '干预建议生成')
    _f19 = ('cyan', 'blue', 'purple', 'orange')
    _h19 = [_find(M1['rows'], t, min_cover=1.0, slack=2) for t in _t19]
    _bd19 = [_host1(h, 3.0) for h in _h19]
    _ic19 = _cn19 = 0
    _ok19 = (all(h is not None for h in _h19) and all(z is not None for z in _bd19)
             and len(set(z['sid'] for z in _bd19)) == 4)
    if _ok19:
        _ok19 = (all(_bd19[i]['x0'] < _bd19[i + 1]['x0'] for i in range(3))
                 and all(_fam(_bd19[i]['fill'], _f19[i]) for i in range(4))
                 and all(h['bold'] and not h['wrapped'] for h in _h19)
                 and all(all(_is_whiteish(c) for c in h['colors']) for h in _h19)
                 and all(_box_in((h['x0'], h['y0'], h['x1'], h['y1']), b['ink'], 0.30)
                         for h, b in zip(_h19, _bd19)))
    if _ok19:
        for b in _bd19:
            _rg = (b['x0'] - 0.20, b['y0'] - 0.10,
                   b['x0'] + max(1.70, 0.95 * b['h']), b['y1'] + 0.10)
            _cr = [s for s in M1['vis'] if not s['has_text'] and s['w'] >= 0.45
                   and s['prst'] in ('ellipse', 'circle') and _pt_in(s['cx'], s['cy'], _rg)]
            for _cc in sorted(_cr, key=lambda z: -z['area']):
                _in = (_cc['x0'] + 0.02, _cc['y0'] + 0.02, _cc['x1'] - 0.02, _cc['y1'] - 0.02)
                if [s for s in M1['vis'] if not s['has_text'] and s['sid'] != _cc['sid']
                        and _pt_in(s['cx'], s['cy'], _in)]:
                    _ic19 += 1
                    break
        _yb19 = (min(b['y0'] for b in _bd19) - 0.25, max(b['y1'] for b in _bd19) + 0.25)
        for i in range(3):
            _g19 = (_bd19[i]['x1'] - 0.05, _yb19[0], _bd19[i + 1]['x0'] + 0.05, _yb19[1])
            for s in M1['vis']:
                if s['has_text'] or not s['arrowish'] or not _pt_in(s['cx'], s['cy'], _g19):
                    continue
                if any(_fam(s['fill'], _f19[j]) or _fam(s['line'], _f19[j]) for j in (i, i + 1)):
                    _cn19 += 1
                    break
        _ok19 = (_ic19 == 4 and _cn19 == 3)
        _d19 = f"色带4条、圆形底矢量图标{_ic19}/4、同色系衔接箭头{_cn19}/3"
    else:
        _d19 = f"命中带内文字{sum(1 for h in _h19 if h)}/4、箭头形色带{sum(1 for z in _bd19 if z)}个"
    results.append(("+5：" if _ok19 else "+0：")
                   + f"顶部4条右指箭头色带按青绿/蓝/紫/橙排布、带内白色加粗单行文字、左端圆形底矢量图标、相邻同色系箭头衔接（{_d19}）")
    # ---------- 图1页左右两栏 5 个模块的共用定位（第20/21/22/23/24条复用） ----------
    def _lrmod(title, subs):
        """“标题条+2个浅色子项框”卡片：返回模块并集外框、标题条与子项框。
           模块外框取三者并集而不取外层大卡片，因作答件可能不画外层卡片。"""
        ht = _find(M1['rows'], title, min_cover=1.0, slack=2)
        bar = _host1(ht, 1.0, pred=lambda z: z['fill'] is not None)
        bx = [_host1(_find(M1['rows'], t, min_cover=1.0, slack=2), 0.8,
                     pred=lambda z: z['fill'] is not None) for t in subs]
        if bar is None or any(z is None for z in bx):
            return None
        gp = [bar] + bx
        return {'box': (min(z['x0'] for z in gp), min(z['y0'] for z in gp),
                        max(z['x1'] for z in gp), max(z['y1'] for z in gp)),
                'bar': bar, 'subs': bx, 'title': ht}

    _nbL = _host1(_find(M1['rows'], '利用时间依赖建模捕捉', min_cover=0.95, slack=4), 4.0)
    _nbR = _host1(_find(M1['rows'], '采用ROC曲线与校准曲线', min_cover=0.95, slack=4), 4.0)
    _tzL = _host1(_find(M1['rows'], '健康数据', min_cover=1.0, slack=2), 2.0)
    _tzR = _host1(_find(M1['rows'], '随访评估', min_cover=1.0, slack=2), 2.0)
    _bkL = _host1(_find(M1['rows'], '构建预测模型', min_cover=1.0, slack=2), 4.0)
    _bkM = _host1(_find(M1['rows'], '模型验证与整合', min_cover=1.0, slack=2), 4.0)
    _bkR = _host1(_find(M1['rows'], '可视化趋势曲线', min_cover=1.0, slack=2), 4.0)
    _mLa = _lrmod('传统机器学习', ('随机森林', '支持向量机'))
    _mLb = _lrmod('时序深度学习', ('生理信号', '生活行为序列'))
    _mRa = _lrmod('训练策略', ('交叉验证', '外部测试'))
    _mRb = _lrmod('模型泛化能力', ('稳定可靠', '参考基准'))

    def _col5(ma, mb, nb, tz, bk, fa, fb):
        """一栏 5 模块判定：模块齐备、自上而下次序正确、左右边界对齐且宽度一致、
           两张卡片标题条色系正确且各带 2 个浅色子项框。
           边界对齐只比“上卡片/下卡片/说明框/底部色块”4 个矩形模块：倒梯形块按细则
           本就上宽下窄，宽度天然小于其余模块，故对它只校验是否居栏中。"""
        if ma is None or mb is None or nb is None or tz is None or bk is None:
            return False, "5个模块未齐备"
        bs = [ma['box'], mb['box'], nb['ink'], tz['ink'], bk['ink']]
        cys = [(b[1] + b[3]) / 2.0 for b in bs]
        if any(cys[i] >= cys[i + 1] for i in range(4)):
            return False, "5模块自上而下次序不符"
        al = [bs[0], bs[1], bs[2], bs[4]]
        ls, rs = [b[0] for b in al], [b[2] for b in al]
        ws = [b[2] - b[0] for b in al]
        dl, dr = max(ls) - min(ls), max(rs) - min(rs)
        wr = max(ws) / max(min(ws), 0.01)
        d = f"左边界差{dl:.2f}cm、右边界差{dr:.2f}cm、宽度比{wr:.2f}"
        if dl > 0.90 or dr > 0.90 or wr > 1.40:
            return False, d
        if abs((tz['x0'] + tz['x1']) / 2.0 - (min(ls) + max(rs)) / 2.0) > 0.80:
            return False, d + "、倒梯形块未居栏中"
        for m, fm in ((ma, fa), (mb, fb)):
            if not _fam(m['bar']['fill'], fm):
                return False, d + "、标题条填充色系不符"
            if not all(_is_lightish(z['fill']) for z in m['subs']):
                return False, d + "、子项框非浅色填充"
        return True, d

    # +5：页面左栏自上而下依次为“传统机器学习”卡片、“时序深度学习”卡片、虚线说明框、倒梯形块、
    #     底部圆角色块共5个模块，右栏与左栏左右对称等宽、同为5个模块；模块左右边界对齐、宽度
    #     一致；四张卡片的标题条依次为绿色、深蓝、紫色、蓝紫填充，每张各带2个浅色子项框
    # 严格对齐细则，仅踩细则明确写出的点：
    #   ① 两栏各 5 个模块须全部定位到（卡片按“标题条+2子项框”定位，其余三个按框内文字定位）；
    #   ② 自上而下次序：每栏 5 个模块中心纵坐标严格递增；
    #   ③ 左右边界对齐、宽度一致：每栏 4 个矩形模块左边界极差与右边界极差均 ≤0.90cm、最大/最小
    #      宽度比 ≤1.40（卡片宽度按标题条与子项框并集量得，与说明框/底部色块之间存在正常的
    #      内外边距差，故宽度比的上限与左右边界极差 0.90cm 相匹配）；
    #   ④ 标题条色系：左栏上绿下深蓝、右栏上紫下蓝紫；四张卡片的 2 个子项框均为浅色填充；
    #   ⑤ 左右对称：左栏左边距（相对画面主体左边）与右栏右边距之差 ≤0.90cm；
    #   ⑥ 等宽：左右两栏宽度比 ≤1.20。
    _okL20, _dL20 = _col5(_mLa, _mLb, _nbL, _tzL, _bkL, 'green', 'blue')
    _okR20, _dR20 = _col5(_mRa, _mRb, _nbR, _tzR, _bkR, 'purple', 'blue')
    _ok20 = _okL20 and _okR20
    _sm20 = ''
    if _ok20:
        _cl20 = [_mLa['box'], _mLb['box'], _nbL['ink'], _bkL['ink']]
        _cr20 = [_mRa['box'], _mRb['box'], _nbR['ink'], _bkR['ink']]
        _wl20 = max(b[2] for b in _cl20) - min(b[0] for b in _cl20)
        _wr20 = max(b[2] for b in _cr20) - min(b[0] for b in _cr20)
        _dl20 = min(b[0] for b in _cl20) - B1[0]
        _dr20 = B1[2] - max(b[2] for b in _cr20)
        _sy20 = abs(_dl20 - _dr20)
        _rw20 = max(_wl20, _wr20) / max(min(_wl20, _wr20), 0.01)
        _ok20 = (_sy20 <= 0.90 and _rw20 <= 1.20)
        _sm20 = f"；左右边距差{_sy20:.2f}cm、两栏宽度比{_rw20:.2f}"
    results.append(("+5：" if _ok20 else "+0：")
                   + "左右两栏各5个模块齐备、左右对称等宽且边界对齐，四张卡片标题条按"
                   + "绿/深蓝/紫/蓝紫配色并各带2个浅色子项框"
                   + f"（左：{_dL20}；右：{_dR20}{_sm20}）")

    def _lnw(s):
        """形状描边宽度（pt）：细则里的“粗箭头”既可能画成有填充的箭头形状，也可能画成
           一条加粗的带箭头直线，后者只能从描边宽度上看出“粗”。"""
        el = _line_el(s['sh'])
        if el is None:
            return 0.0
        try:
            return float(el.get('w') or 0) / 12700.0
        except Exception:
            return 0.0

    # +3：左右两栏各设1个橙黄虚线边、浅米黄填充的说明框，左框内两行居中文字为“利用时间依赖建模
    #     捕捉”“风险状态的动态变化”，右框内两行居中文字为“采用ROC曲线与校准曲线”“综合评估模型
    #     判别与泛化性能”，四行文字均完整居于框内、无溢出框外或折出第三行
    # 严格对齐细则，仅踩细则明确写出的点：
    #   ① 浅米黄填充：填充为高明度低饱和且 R≥G>B（米黄向暖色偏移）；
    #   ② 橙黄虚线边：框自身描边为虚线族，或框四周环带内有足够多的短线段（复刻件常把虚线
    #      拆成成百上千段实线短段绘制，只看 a:prstDash 会漏判）；
    #   ③ 两行文字：两段文字均命中、均未折行（折行即等于折出第三行）；
    #   ④ 完整居于框内：文字墨迹落在框内（容差 0.15cm）；⑤ 居中：行中心与框中心横向偏差 ≤0.60cm；
    #   ⑥ 两行上下相邻：上行底边不低于下行顶边。
    def _note22(nb, t1, t2):
        if nb is None:
            return False, "说明框未识别"
        _fl = nb['fill']
        if not (_is_lightish(_fl) and _fl[0] >= _fl[1] >= _fl[2] and _fl[0] - _fl[2] >= 6):
            return False, "说明框非浅米黄填充"
        _ds = nb['dashed'] or _perim_dashed(M1, nb['ink'], need_segs=8)
        _rg = (nb['x0'] - 0.35, nb['y0'] - 0.35, nb['x1'] + 0.35, nb['y1'] + 0.35)
        _hh = [_find(M1['rows'], t, region=_rg, min_cover=0.95, slack=4) for t in (t1, t2)]
        _sd = "虚线边框" if _ds else "边框非虚线"
        if any(h is None for h in _hh):
            return False, _sd + "、框内两行文字未命中"
        _cx = (nb['x0'] + nb['x1']) / 2.0
        _tx = (all(not h['wrapped'] for h in _hh)
               and all(_box_in((h['x0'], h['y0'], h['x1'], h['y1']), nb['ink'], 0.15) for h in _hh)
               and all(abs((h['x0'] + h['x1']) / 2.0 - _cx) <= 0.60 for h in _hh)
               and _hh[0]['y1'] <= _hh[1]['y0'] + 0.10)
        return (_ds and _tx), _sd + ("、两行文字居中于框内" if _tx else "、文字未居中或溢出框外")
    _o22a, _d22a = _note22(_nbL, '利用时间依赖建模捕捉', '风险状态的动态变化')
    _o22b, _d22b = _note22(_nbR, '采用ROC曲线与校准曲线', '综合评估模型判别与泛化性能')
    _ok22 = _o22a and _o22b
    results.append(("+3：" if _ok22 else "+0：")
                   + f"左右两栏各1个橙黄虚线边浅米黄说明框、框内两行文字居中且不溢出不折第三行（左：{_d22a}；右：{_d22b}）")

    # +3：左右两栏虚线说明框下方各设1个向下粗箭头，箭头下方各设1个上宽下窄的倒梯形块，块内文字
    #     分别为“健康数据”（左）与“随访评估”（右）
    # 严格对齐细则，仅踩细则明确写出的点：
    #   ① 位置：箭头须落在“说明框底边—倒梯形块顶边”之间的带内，且与倒梯形块横向对中
    #      （中心横向偏差 ≤0.90cm），从而排除说明框自身的虚线短段；
    #   ② 粗：或为有实体的箭头形状（宽高均 ≥0.25cm），或为描边 ≥3pt 的带箭头直线；
    #   ③ 上宽下窄：预设几何为梯形/三角族即认定，自定义几何按路径取样点比较顶部与底部
    #      横向跨度（顶宽 ≥1.15×底宽）；④ 块内文字命中“健康数据”／“随访评估”且落在块内。
    def _down23(nb, tz, tag):
        if nb is None or tz is None:
            return False, tag + "说明框或倒梯形块未识别"
        _rg = (min(nb['x0'], tz['x0']) - 0.30, nb['y1'] - 0.05,
               max(nb['x1'], tz['x1']) + 0.30, tz['y0'] + 0.05)
        _tc = (tz['x0'] + tz['x1']) / 2.0
        _ar = None
        for s in M1['vis']:
            if s['has_text'] or not _pt_in(s['cx'], s['cy'], _rg):
                continue
            if abs(s['cx'] - _tc) > 0.90 or max(s['w'], s['h']) < 0.30:
                continue
            if (s['w'] >= 0.25 and s['h'] >= 0.25) or _lnw(s) >= 3.0:
                _ar = s
                break
        _inv = tz['prst'] in ('trapezoid', 'triangle', 'flowChartManualOperation')
        if not _inv and len(tz['pts']) >= 4:
            _ys = [p[1] for p in tz['pts']]
            _lo, _hi = min(_ys), max(_ys)
            _tp = [p[0] for p in tz['pts'] if p[1] <= _lo + 0.15 * (_hi - _lo)]
            _bt = [p[0] for p in tz['pts'] if p[1] >= _hi - 0.15 * (_hi - _lo)]
            if _tp and _bt:
                _inv = (max(_tp) - min(_tp)) >= 1.15 * max(max(_bt) - min(_bt), 0.01)
        _d = tag + ("有粗箭头" if _ar is not None else "无粗箭头") + ("、倒梯形块上宽下窄" if _inv else "、块形非上宽下窄")
        return (_ar is not None and _inv), _d
    _o23a, _d23a = _down23(_nbL, _tzL, '左')
    _o23b, _d23b = _down23(_nbR, _tzR, '右')
    _ok23 = _o23a and _o23b
    results.append(("+3：" if _ok23 else "+0：")
                   + f"左右两栏说明框下方各1个向下粗箭头、箭头下方各1个上宽下窄倒梯形块并含“健康数据”“随访评估”（{_d23a}；{_d23b}）")

    # +5：页面底部横排3个圆角矩形色块，文字依次为“构建预测模型”（左，绿色系）、“模型验证与整合”
    #     （中，蓝色系）、“可视化趋势曲线”（右，紫色系），三块高度一致、底边对齐于同一横带；中间
    #     色块分别向左、向右引出2个粗箭头指向左右两块，3个色块内各配1个图标（数据库、盾牌对勾、
    #     折线图表）且图标与文字左右并排不压盖
    # 严格对齐细则，仅踩细则明确写出的点：
    #   ① 3 个色块自左至右排列且填充依次为绿/蓝/紫色系；
    #   ② 高度一致：最大/最小高度比 ≤1.25；③ 底边对齐：3 块底边极差 ≤0.35cm；
    #   ④ 中间块左右各 1 个粗箭头：两个间隙内各有 1 个横向跨度 ≥0.8cm 的构件，且或为有实体的
    #      箭头形状（宽高均 ≥0.25cm）、或为描边 ≥3pt 的带箭头直线；
    #   ⑤ 块内图标且不压盖：块内除色块本身外另有非文字构件完整位于该块文字左侧（图标右边界
    #      不越过文字左边界 0.20cm），用符号字符充当图标时块内没有这样的构件，本条不成立。
    _t24 = ('构建预测模型', '模型验证与整合', '可视化趋势曲线')
    _f24 = ('green', 'blue', 'purple')
    _bk24 = [_bkL, _bkM, _bkR]
    _ic24 = _ar24 = 0
    _hr24 = _bd24 = 0.0
    _ok24 = (all(z is not None for z in _bk24) and len(set(z['sid'] for z in _bk24)) == 3)
    if _ok24:
        _ok24 = (_bk24[0]['x0'] < _bk24[1]['x0'] < _bk24[2]['x0']
                 and all(_fam(_bk24[i]['fill'], _f24[i]) for i in range(3)))
    if _ok24:
        _hs24 = [z['h'] for z in _bk24]
        _bs24 = [z['y1'] for z in _bk24]
        _hr24 = max(_hs24) / max(min(_hs24), 0.01)
        _bd24 = max(_bs24) - min(_bs24)
        for i in (0, 1):
            _g24 = (_bk24[i]['x1'] - 0.05, min(z['y0'] for z in _bk24) - 0.30,
                    _bk24[i + 1]['x0'] + 0.05, max(_bs24) + 0.30)
            for s in M1['vis']:
                if s['has_text'] or not _pt_in(s['cx'], s['cy'], _g24) or s['w'] < 0.80:
                    continue
                if (s['w'] >= 0.25 and s['h'] >= 0.25) or _lnw(s) >= 3.0:
                    _ar24 += 1
                    break
        for z, t in zip(_bk24, _t24):
            _h = _find(M1['rows'], t, min_cover=1.0, slack=2)
            if _h is None:
                continue
            if [s for s in M1['vis'] if not s['has_text'] and s['sid'] != z['sid']
                    and 0.01 <= s['area'] <= 1.80 and s['x1'] <= _h['x0'] + 0.20
                    and _pt_in(s['cx'], s['cy'], (z['x0'], z['y0'], z['x1'], z['y1']))]:
                _ic24 += 1
        _ok24 = (_hr24 <= 1.25 and _bd24 <= 0.35 and _ar24 == 2 and _ic24 == 3)
        _d24 = f"高度比{_hr24:.2f}、底边极差{_bd24:.2f}cm、粗箭头{_ar24}/2、块内图标{_ic24}/3"
    else:
        _d24 = f"底部色块{sum(1 for z in _bk24 if z)}个或填充色系不符"
    results.append(("+5：" if _ok24 else "+0：")
                   + f"底部3个圆角矩形色块按绿/蓝/紫排布、高度一致底边对齐、中间块左右各1个粗箭头、3块内各1个图标与文字并排不压盖（{_d24}）")

    # ---------- 图1页中部构件的共用定位（第25/26/27/28/29/30/31条复用） ----------
    def _famlt(c, name):
        """浅色几何块的色系判定：浅色填充通道差很小（如浅绿 (235,244,232)），
           只判主色方向，阈值必须比深色块小，否则一律落到白色。"""
        if c is None:
            return False
        r, g, b = c
        if name == 'green':
            return g > r + 6 and g > b + 6
        if name == 'purple':
            return b > r + 6 and b > g + 6
        if name == 'yellow':
            return r > b + 15 and g > b + 12
        if name == 'cyan':
            return g > r + 8 and b > r + 8 and abs(g - b) <= 28
        return False

    def _panel1(h):
        """图表面板：把标题整体包住、面积≥12cm² 的最小图形。"""
        if h is None:
            return None
        c = [s for s in M1['vis'] if not s['has_text'] and s['area'] >= 12.0
             and s['x0'] <= h['x0'] + 0.50 and s['x1'] >= h['x1'] - 0.50
             and s['y0'] <= h['y0'] + 0.60 and s['y1'] >= h['y1']]
        c.sort(key=lambda z: z['area'])
        return c[0] if c else None

    _hb29 = _find(M1['rows'], '特征重要性（示例）', min_cover=0.85, slack=5)
    _hl30 = _find(M1['rows'], '风险随时间变化（示例）', min_cover=0.85, slack=5)
    _pnL, _pnR = _panel1(_hb29), _panel1(_hl30)
    _h26 = _find(M1['rows'], '风险等级', min_cover=1.0, slack=2)
    _cc26 = None
    if _h26 is not None:
        _c26 = [s for s in M1['vis'] if not s['has_text'] and s['prst'] in ('ellipse', 'circle')
                and s['fill'] is not None and s['fill'][2] > s['fill'][0] + 20
                and _box_in((_h26['x0'], _h26['y0'], _h26['x1'], _h26['y1']), s['ink'], 0.30)]
        _c26.sort(key=lambda z: z['area'])
        _cc26 = _c26[0] if _c26 else None
    _t27 = ('连续指标', '症状量表', '生活方式因素', '异常事件识别')
    _f27 = ('green', 'purple', 'yellow', 'cyan')
    _h27 = [_find(M1['rows'], t, min_cover=1.0, slack=2) for t in _t27]
    _bl27 = [_host1(h, 0.80) for h in _h27]
    _sg1 = [s for s in M1['vis'] if not s['has_text'] and s['kind'] in ('line', 'free')
            and s['seglen'] is not None and 0.05 <= s['seglen'] <= 0.62]

    # 23) +3：页面中部蓝色虚线边框大圆角矩形区域，四边完整闭合并框住中央关系图与两个
    #     图表面板；区域底部居中蓝色文字"临床表型解析与功能关联"，位于两面板下方且不与
    #     虚线边框及底部色块重叠。
    #     严格对齐细则，仅踩细则明确写出的点：①存在把"中央圆+4个几何块+2个图表面板"整体
    #     框住的边框，且四边都能查到虚线证据（虚线既可能是一个 prstDash 边框，也可能由密集
    #     短线段拼出，两种画法都认；描边色由主题继承、读不到显式色值时不强求蓝色）；
    #     ②图题文字单行、位于两面板下方；③文字整体在边框以内、且不压到底部3个色块。
    _ok25 = (_cc26 is not None and _pnL is not None and _pnR is not None
             and all(z is not None for z in _bl27))
    _ed25, _R25, _fr25 = 0, None, None
    if _ok25:
        _bs25 = [_cc26['ink'], _pnL['ink'], _pnR['ink']] + [z['ink'] for z in _bl27]
        _R25 = (min(b[0] for b in _bs25), min(b[1] for b in _bs25),
                max(b[2] for b in _bs25), max(b[3] for b in _bs25))
        _bg25 = [s for s in M1['vis'] if not s['has_text'] and s['area'] >= 60.0
                 and _box_in(_R25, s['ink'], 0.60)]
        _fr25 = min(_bg25, key=lambda z: z['area']) if _bg25 else None
        _dl25 = _fr25 is not None and (_fr25['dashed'] or _perim_dashed(M1, _fr25['ink'], 8))
        for _st25, _ax25 in (((_R25[0] - 1.2, _R25[1] - 1.2, _R25[2] + 1.2, _R25[1] + 0.10), 0),
                             ((_R25[0] - 1.2, _R25[3] - 0.10, _R25[2] + 1.2, _R25[3] + 1.2), 0),
                             ((_R25[0] - 1.2, _R25[1] - 1.2, _R25[0] + 0.10, _R25[3] + 1.2), 1),
                             ((_R25[2] - 0.10, _R25[1] - 1.2, _R25[2] + 1.2, _R25[3] + 1.2), 1)):
            _q25 = [s for s in _sg1 if _pt_in(s['cx'], s['cy'], _st25)]
            _vs25 = [(s['cx'] if _ax25 == 0 else s['cy']) for s in _q25]
            _sp25 = (max(_vs25) - min(_vs25)) if _vs25 else 0.0
            _ln25 = (_R25[2] - _R25[0]) if _ax25 == 0 else (_R25[3] - _R25[1])
            if (len(_q25) >= 8 and _sp25 >= 0.60 * _ln25) or _dl25:
                _ed25 += 1
    _hc25 = _find(M1['rows'], '临床表型解析与功能关联', min_cover=0.95, slack=4)
    _in25 = _bw25 = False
    if _ok25 and _hc25 is not None:
        _bd25 = [z for z in _bk24 if z is not None]
        _in25 = ((_fr25 is None or _box_in((_hc25['x0'], _hc25['y0'], _hc25['x1'], _hc25['y1']),
                                           _fr25['ink'], 0.10))
                 and _hc25['y0'] >= max(_pnL['y1'], _pnR['y1']) - 0.15
                 and (not _bd25 or _hc25['y1'] <= min(z['y0'] for z in _bd25) - 0.02))
        _bw25 = (not _hc25['wrapped']
                 and (not _hc25['colors'] or all(_fam(c, 'blue') or _fam(c, 'purple')
                                                 for c in _hc25['colors']))
                 and abs((_hc25['x0'] + _hc25['x1']) / 2.0 - (_R25[0] + _R25[2]) / 2.0) <= 1.30)
    if _ed25 == 4 and _in25 and _bw25:
        results.append('+3：中部虚线边框区域四边闭合并框住中央关系图与两个图表面板，'
                       '底部图题"临床表型解析与功能关联"单行居中、位于面板下方且不与边框、底部色块重叠')
    else:
        results.append('+0：中部虚线框区域不合要求（%s；图题位置%s）'
                       % ('虚线边四数%d/4' % _ed25 if _ok25 else '区域内构件未定位',
                          '正常' if (_in25 and _bw25) else '异常'))

    # 24) +5：中央1个深蓝实心圆，圆内白色加粗文字"风险等级"，圆外环绕1圈完整蓝色虚线圆环。
    #     严格对齐细则，仅踩细则明确写出的点：①正圆（长短轴比≤1.20）、有填充且为深蓝色系；
    #     ②"风险等级"落在圆内、加粗、白色（字色继承主题读不到时不强求）；③圆外一圈虚线圆环
    #     ——虚线圆环既可能是一个带虚线描边的同心圆，也可能由密集短弧段拼出，两种画法都认，
    #     但都要求"完整环绕"：短弧段须落在 1.02r~1.90r 的圆环带内并覆盖8个方位中的≥5个。
    _rd26 = _rn26 = 0.0
    _sec26, _cn26 = 0, 0
    _ok26 = _cc26 is not None and max(_cc26['w'], _cc26['h']) <= 1.20 * max(min(_cc26['w'], _cc26['h']), 0.01)
    if _ok26:
        _ok26 = _fam(_cc26['fill'], 'blue') and max(_cc26['fill']) <= 225
    _tw26 = (_h26 is not None and _cc26 is not None and _h26['bold'] and not _h26['wrapped']
             and _box_in((_h26['x0'], _h26['y0'], _h26['x1'], _h26['y1']), _cc26['ink'], 0.10)
             and (not _h26['colors'] or all(_is_whiteish(c) for c in _h26['colors'])))
    if _cc26 is not None:
        _cx26, _cy26 = _cc26['cx'], _cc26['cy']
        _rd26 = max(_cc26['w'], _cc26['h']) / 2.0
        _sg26 = []
        for s in _sg1:
            _d26 = math.hypot(s['cx'] - _cx26, s['cy'] - _cy26)
            if 1.02 * _rd26 <= _d26 <= 1.90 * _rd26:
                _sg26.append(s)
        _cn26 = len(_sg26)
        _sec26 = len(set(int((math.degrees(math.atan2(s['cy'] - _cy26, s['cx'] - _cx26)) + 360) % 360) // 45
                         for s in _sg26))
        _rg26 = [s for s in M1['vis'] if not s['has_text'] and s['prst'] in ('ellipse', 'circle')
                 and math.hypot(s['cx'] - _cx26, s['cy'] - _cy26) <= 0.60 * _rd26
                 and 1.05 * _rd26 <= max(s['w'], s['h']) / 2.0 <= 2.40 * _rd26
                 and (s['dashed'] or _perim_dashed(M1, s['ink'], 8))]
        _rn26 = len(_rg26)
    _rk26 = (_cn26 >= 8 and _sec26 >= 5) or _rn26 >= 1
    if _ok26 and _tw26 and _rk26:
        results.append('+5：中央深蓝实心圆内含白色加粗"风险等级"，圆外环绕1圈完整虚线圆环'
                       '（半径%.2fcm、环带虚线段%d处覆盖%d/8方位、虚线同心圆%d个）'
                       % (_rd26, _cn26, _sec26, _rn26))
    else:
        results.append('+0：中央深蓝实心圆或外围虚线圆环不合要求（深蓝实心圆%s、白色加粗文字%s、'
                       '虚线圆环%s：环带虚线段%d处覆盖%d/8方位、虚线同心圆%d个）'
                       % ('有' if _ok26 else '无', '有' if _tw26 else '无',
                          '有' if _rk26 else '无', _cn26, _sec26, _rn26))

    # 25) +5：中央圆外围均匀分布4个彩色几何块（左上绿"连续指标"、右上紫"症状量表"、
    #     左下黄"生活方式因素"、右下青"异常事件识别"），每块内配1个矢量图标且文字单行不换行。
    #     严格对齐细则，仅踩细则明确写出的点：①4段文字各自有承载图形，且相对中央圆分居
    #     左上/右上/左下/右下四个方位、到圆心距离相近（最大/最小≤1.60，即"均匀分布"）；
    #     ②填充色依次为绿/紫/黄/青色系（浅色填充按通道方向判色）；③每块内另有≥1个非文字
    #     小图形作矢量图标；④文字单行不换行。几何块用圆形还是六边形属画法差异，此处只认
    #     位置、颜色、图标与文字，不强求 prstGeom 名称（矢量图标常以自由曲线拼出）。
    _ic27, _ds27, _po27 = 0, [], 0
    _ok27 = _cc26 is not None and all(z is not None for z in _bl27) and all(h is not None for h in _h27)
    if _ok27:
        for _i27, _b27 in enumerate(_bl27):
            _ds27.append(math.hypot(_b27['cx'] - _cc26['cx'], _b27['cy'] - _cc26['cy']))
            _sx = -1 if _i27 in (0, 2) else 1
            _sy = -1 if _i27 in (0, 1) else 1
            if (_b27['cx'] - _cc26['cx']) * _sx > 0 and (_b27['cy'] - _cc26['cy']) * _sy > 0:
                _po27 += 1
            _ii27 = [s for s in M1['vis'] if not s['has_text'] and s['sid'] != _b27['sid']
                     and 0.01 <= s['area'] <= 2.00 and _pt_in(s['cx'], s['cy'], _b27['ink'])
                     and not _pt_in(s['cx'], s['cy'], (_h27[_i27]['x0'] - 0.05, _h27[_i27]['y0'] - 0.05,
                                                       _h27[_i27]['x1'] + 0.05, _h27[_i27]['y1'] + 0.05))]
            if _ii27:
                _ic27 += 1
    _un27 = (max(_ds27) <= 1.60 * max(min(_ds27), 0.01)) if _ds27 else False
    _cl27 = _ok27 and all(_fam(_bl27[i]['fill'], _f27[i]) or _famlt(_bl27[i]['fill'], _f27[i])
                          for i in range(4))
    _tx27 = _ok27 and all(not h['wrapped'] for h in _h27)
    if _ok27 and _po27 == 4 and _un27 and _cl27 and _tx27 and _ic27 == 4:
        results.append('+5：中央圆外围4个彩色几何块按左上绿/右上紫/左下黄/右下青均匀分布，'
                       '每块内配1个矢量图标且文字单行不换行（半径比%.2f、块内图标4/4）'
                       % (max(_ds27) / max(min(_ds27), 0.01)))
    else:
        results.append('+0：外围4个彩色几何块不合要求（承载图形%d/4、方位%d/4、均匀分布%s、'
                       '色系%s、单行%s、块内矢量图标%d/4）'
                       % (len([z for z in _bl27 if z is not None]), _po27,
                          '是' if _un27 else '否', '是' if _cl27 else '否',
                          '是' if _tx27 else '否', _ic27))

    # 26) +5：中央圆正上/正左/正右/正下各1个白底细框小标签（AUC、F1、SHAP、随访评分）；
    #     中央圆与4个几何块及4个小标签之间共8条放射状均匀分布的双向箭头，箭头两端止于图形
    #     边缘，不穿入中央圆内部、不彼此交叉压盖。
    #     严格对齐细则，仅踩细则明确写出的点：①4段标签文字各有承载图形，填充为白/近白、
    #     且有描边（细框），并分居中央圆的正上/正左/正右/正下（同轴偏移≤1.8cm）；
    #     ②同色的放射连线（含箭头笔画）在 0.9r~3.0r 环带内≥8条并覆盖8个方位中的≥5个，
    #     即"8条放射状均匀分布"；③每条连线的取样端点都不落进 0.75r 以内，即"不穿入中央圆"。
    #     箭头头部可能是与杆线分开的独立笔画，故此处不要求连线自身带 arrow 标记；"不彼此
    #     交叉压盖"由"分居≥5个方位"的均匀性一并覆盖。
    _t28 = ('AUC', 'F1', 'SHAP', '随访评分')
    _h28 = [_find(M1['rows'], t, min_cover=1.0, slack=2) for t in _t28]
    _b28 = [_host1(h, 0.05, 4.00) for h in _h28]
    _lb28 = 0
    if _cc26 is not None:
        for _i28, _z28 in enumerate(_b28):
            if _z28 is None or _z28['line'] is None or _z28['fill'] is None \
                    or not _is_lightish(_z28['fill'], 0.88):
                continue
            _dx28, _dy28 = _z28['cx'] - _cc26['cx'], _z28['cy'] - _cc26['cy']
            if ((_i28 == 0 and _dy28 < 0 and abs(_dx28) <= 1.80)
                    or (_i28 == 1 and _dx28 < 0 and abs(_dy28) <= 1.80)
                    or (_i28 == 2 and _dx28 > 0 and abs(_dy28) <= 1.80)
                    or (_i28 == 3 and _dy28 > 0 and abs(_dx28) <= 1.80)):
                _lb28 += 1
    _cn28, _sc28, _pi28, _bs28 = 0, 0, 0.0, 0
    if _cc26 is not None and _rd26 > 0:
        _gp28 = {}
        for s in M1['vis']:
            if s['has_text'] or s['kind'] not in ('line', 'free') or s['line'] is None:
                continue
            _d28 = math.hypot(s['cx'] - _cc26['cx'], s['cy'] - _cc26['cy'])
            if 0.90 * _rd26 <= _d28 <= 3.00 * _rd26:
                _gp28.setdefault(s['line'], []).append(s)
        _bg28 = []
        for _v28 in _gp28.values():
            if len(_v28) < 8:
                continue
            _s28 = len(set(int((math.degrees(math.atan2(s['cy'] - _cc26['cy'], s['cx'] - _cc26['cx']))
                                + 360) % 360) // 45 for s in _v28))
            if not _bg28 or _s28 > _bs28:
                _bg28, _bs28 = _v28, _s28
        _cn28 = len(_bg28)
        _sc28 = _bs28 if _bg28 else 0
        _pi28 = min([math.hypot(p[0] - _cc26['cx'], p[1] - _cc26['cy'])
                     for s in _bg28 for p in (s['pts'] or [(s['cx'], s['cy'])])] or [0.0])
    _ra28 = _cn28 >= 8 and _sc28 >= 5 and _pi28 >= 0.75 * _rd26
    if _lb28 == 4 and _ra28:
        results.append('+5：中央圆正上/正左/正右/正下各1个白底细框小标签，四周共%d条放射状'
                       '双向箭头覆盖%d/8方位且两端止于图形边缘未穿入圆内（最近端点距圆心%.2fcm、'
                       '圆半径%.2fcm）' % (_cn28, _sc28, _pi28, _rd26))
    else:
        results.append('+0：4个小标签或8条放射状双向箭头不合要求（白底细框标签%d/4、放射连线%d条'
                       '覆盖%d/8方位、最近端点距圆心%.2fcm（须≥%.2fcm））'
                       % (_lb28, _cn28, _sc28, _pi28, 0.75 * _rd26))

    # 27) +5：虚线区域内左侧条形图面板：标题"特征重要性（示例）"，纵向7个类别标签及省略号，
    #     对应7根自上而下依次变短的蓝色横条，横轴刻度0.0/0.1/0.2/0.3/0.4，横轴标题
    #     "平均绝对SHAP值"，标签与条形一一对齐不重叠。
    #     严格对齐细则，仅踩细则明确写出的点：①面板与标题存在；②7个类别标签齐备且自上而下
    #     顺序正确，标签下方另有省略号；③每个标签都能在同一水平线上（中心差≤0.25cm）配到一根
    #     蓝色横条，且横条起点在标签右侧（不重叠），7根条宽严格递减；④横轴5个刻度值与横轴标题
    #     齐备。刻度可能写成5个独立文本，也可能写成一行用空格排开的文本，故按面板内文本
    #     去空格后含这5个刻度串判定。
    _t29 = ('心率变异性', '睡眠时长', '日均步数', '体重指数', '血压波动', '吸烟状态', '饮酒频率')
    _lb29 = [_find(M1['rows'], t, min_cover=1.0, slack=2) for t in _t29]
    _ha29 = _find(M1['rows'], '平均绝对SHAP值', min_cover=0.90, slack=4)
    _ok29 = (_pnL is not None and _hb29 is not None and _ha29 is not None
             and all(h is not None for h in _lb29))
    _od29 = _dc29 = _tk29 = _el29 = False
    _ws29, _mt29 = [], []
    if _ok29:
        _cy29 = [(h['y0'] + h['y1']) / 2.0 for h in _lb29]
        _od29 = all(_cy29[i] < _cy29[i + 1] for i in range(6))
        _bs29 = [s for s in M1['vis'] if not s['has_text'] and _pt_in(s['cx'], s['cy'], _pnL['ink'])
                 and s['fill'] is not None and _fam(s['fill'], 'blue')
                 and s['w'] >= 3.0 * max(s['h'], 0.01) and s['h'] <= 0.60 and s['w'] >= 0.40]
        for _i29, _h29 in enumerate(_lb29):
            _c29 = [b for b in _bs29 if abs(b['cy'] - _cy29[_i29]) <= 0.25
                    and b['x0'] >= _h29['x1'] - 0.10 and b['sid'] not in _mt29]
            if _c29:
                _c29.sort(key=lambda z: abs(z['cy'] - _cy29[_i29]))
                _mt29.append(_c29[0]['sid'])
                _ws29.append(_c29[0]['w'])
        _dc29 = len(_ws29) == 7 and all(_ws29[i] > _ws29[i + 1] + 0.01 for i in range(6))
        _tj29 = _norm(''.join(u['text'] for u in _region_units(M1, _pnL['ink'])))
        _tk29 = all(t in _tj29 for t in ('0.0', '0.1', '0.2', '0.3', '0.4'))
        _el29 = any(t in _tj29 for t in ('…', '...', '⋯'))
    if _ok29 and _od29 and _dc29 and _tk29 and _el29:
        results.append('+5：左侧条形图面板标题、7个类别标签与省略号、7根依次变短的蓝色横条'
                       '（%s cm）、横轴刻度与标题齐备，标签与横条一一对齐不重叠'
                       % '/'.join('%.2f' % w for w in _ws29))
    else:
        results.append('+0：左侧条形图面板不合要求（面板与标题%s、类别标签%d/7且顺序%s、'
                       '省略号%s、对齐横条%d/7、条宽递减%s、横轴刻度%s）'
                       % ('有' if (_pnL is not None and _hb29 is not None) else '无',
                          len([h for h in _lb29 if h is not None]), '正确' if _od29 else '异常',
                          '有' if _el29 else '无', len(_ws29), '是' if _dc29 else '否',
                          '齐备' if _tk29 else '缺失'))

    # 28) +5：虚线区域内右侧折线图面板：标题"风险随时间变化（示例）"，图内3条连续折线
    #     （绿、橙、红）自左下向右上抬升、上下顺序不互穿，纵轴刻度0.0/0.5/1.0、横轴刻度
    #     0/3/6/9/12，横轴标题"随访时间（月）"。
    #     严格对齐细则，仅踩细则明确写出的点：①面板与标题存在；②按描边色分组，绿/橙/红三色
    #     各有一条横跨绘图区（≥面板宽40%）的折线——折线既可能是一条多点自由曲线，也可能由
    #     多段短线拼接，两种画法都认；③每条折线左三分之一的平均纵坐标高于右三分之一（即
    #     自左下向右上抬升）；④三条折线在左右两端的上下次序一致（不互穿）；⑤两轴刻度值与
    #     横轴标题齐备（刻度可能是独立文本或一行空格排开的文本，按去空格后含刻度串判定）。
    _f30 = ('green', 'orange', 'red')
    _hx30 = _find(M1['rows'], '随访时间（月）', min_cover=0.90, slack=4)
    _ok30 = _pnR is not None and _hl30 is not None and _hx30 is not None
    _cv30, _up30, _nx30, _tk30 = [], 0, False, False
    if _ok30:
        _pw30 = _pnR['x1'] - _pnR['x0']
        _gp30 = {}
        for s in M1['vis']:
            if s['has_text'] or s['kind'] not in ('line', 'free') or s['line'] is None:
                continue
            if _pt_in(s['cx'], s['cy'], _pnR['ink']):
                _gp30.setdefault(s['line'], []).append(s)
        for _c30, _v30 in _gp30.items():
            _fm30 = [f for f in _f30 if _fam(_c30, f)]
            if not _fm30:
                continue
            _mj30 = [s for s in _v30 if s['w'] >= 0.40 * _pw30]
            _us30 = _mj30 if _mj30 else _v30
            _px30 = [p for s in _us30 for p in (s['pts'] or [(s['cx'], s['cy'])])]
            _xx30 = [p[0] for p in _px30]
            if len(_px30) < 2 or max(_xx30) - min(_xx30) < 0.40 * _pw30:
                continue
            _q30 = (max(_xx30) - min(_xx30)) / 3.0
            _ly30 = [p[1] for p in _px30 if p[0] <= min(_xx30) + _q30]
            _ry30 = [p[1] for p in _px30 if p[0] >= max(_xx30) - _q30]
            if not _ly30 or not _ry30:
                continue
            _la30 = sum(_ly30) / len(_ly30)
            _ra30 = sum(_ry30) / len(_ry30)
            _cv30.append((_fm30[0], _la30, _ra30))
            if _la30 - _ra30 >= 0.10:
                _up30 += 1
        _fs30 = set(z[0] for z in _cv30)
        _nx30 = (len(_cv30) == 3 and len(_fs30) == 3
                 and [z[0] for z in sorted(_cv30, key=lambda z: z[1])]
                 == [z[0] for z in sorted(_cv30, key=lambda z: z[2])])
        _tj30 = _norm(''.join(u['text'] for u in _region_units(M1, _pnR['ink'])))
        _tk30 = (all(t in _tj30 for t in ('0.0', '0.5', '1.0'))
                 and all(t in _tj30 for t in ('3', '6', '9', '12')))
    if _ok30 and len(_cv30) == 3 and _up30 == 3 and _nx30 and _tk30:
        results.append('+5：右侧折线图面板标题、绿/橙/红3条折线自左下向右上抬升且上下顺序不互穿、'
                       '两轴刻度与横轴标题齐备（3条折线左右端纵坐标 %s）'
                       % '，'.join('%s %.2f→%.2f' % z for z in _cv30))
    else:
        results.append('+0：右侧折线图面板不合要求（面板与标题%s、绿橙红折线%d/3、抬升%d/3、'
                       '互不穿越%s、两轴刻度%s、横轴标题%s）'
                       % ('有' if (_pnR is not None and _hl30 is not None) else '无',
                          len(_cv30), _up30, '是' if _nx30 else '否',
                          '齐备' if _tk30 else '缺失', '有' if _hx30 is not None else '无'))

    # 29) +3：折线图纵轴左侧设竖排（旋转90°）轴标题文字，字序自上而下、方向正立可读；
    #     图内右上方3行图例（低风险/中风险/高风险），每行由1段对应颜色短线与文字组成且
    #     单行不换行、不压盖折线绘图区。
    #     严格对齐细则，仅踩细则明确写出的点：①存在旋转约90°/270°的竖排文字，位置贴着折线图
    #     面板左边（面板左边界±1.4cm 内）——字面朝向属渲染效果，非渲染无法判读，故只认"竖排
    #     且贴纵轴"；②三行图例文字齐备、自上而下顺序为低/中/高且各自单行；③三行都位于面板
    #     右上区域（不压盖绘图区主体）；④每行文字左侧1.4cm 内有1段短线作色标（宽≥0.20cm、
    #     高≤0.25cm，以区别于折线本身与圆点标记），三段色标颜色互不相同。
    _t31 = ('低风险', '中风险', '高风险')
    _h31 = [_find(M1['rows'], t, min_cover=1.0, slack=2) for t in _t31]
    _vt31, _ky31, _po31, _od31 = False, [], 0, False
    if _pnR is not None:
        _vt31 = any(s['has_text'] and s['text'].strip()
                    and (60.0 <= (abs(s['rot']) % 180.0) <= 120.0)
                    and _pnR['x0'] - 1.40 <= s['cx'] <= _pnR['x0'] + 1.40
                    for s in M1['shapes'])
        if all(h is not None for h in _h31):
            _cy31 = [(h['y0'] + h['y1']) / 2.0 for h in _h31]
            _od31 = (_cy31[0] < _cy31[1] < _cy31[2]) and all(not h['wrapped'] for h in _h31)
            for _i31, _h in enumerate(_h31):
                if (_h['x0'] + _h['x1']) / 2.0 >= (_pnR['x0'] + _pnR['x1']) / 2.0 \
                        and _cy31[_i31] <= (_pnR['y0'] + _pnR['y1']) / 2.0:
                    _po31 += 1
                _rg31 = (_h['x0'] - 1.40, _cy31[_i31] - 0.15, _h['x0'] - 0.01, _cy31[_i31] + 0.15)
                _cd31 = [s for s in M1['vis'] if not s['has_text'] and s['line'] is not None
                         and s['w'] >= 0.20 and s['h'] <= 0.25
                         and _pt_in(s['cx'], s['cy'], _rg31)]
                if _cd31:
                    _ky31.append(_cd31[0]['line'])
    if _vt31 and _od31 and _po31 == 3 and len(_ky31) == 3 and len(set(_ky31)) == 3:
        results.append('+3：折线图纵轴左侧有竖排轴标题，图内右上方3行图例按低/中/高风险排列、'
                       '各由1段对应颜色短线与单行文字组成（色标颜色%s）'
                       % '/'.join(str(c) for c in _ky31))
    else:
        results.append('+0：纵轴竖排标题或3行图例不合要求（竖排轴标题%s、图例文字顺序与单行%s、'
                       '位于面板右上%d/3、颜色短线色标%d/3且异色%d）'
                       % ('有' if _vt31 else '无', '正确' if _od31 else '异常', _po31,
                          len(_ky31), len(set(_ky31))))

    # 30) +5：两页全部文字均为可直接改字的文本对象，且方框、面板、箭头、连接线、几何块、
    #     立体箱体各面等结构元素均为可单独编辑的矢量形状，不存在整片贴图或被固化在图片内的文字。
    #     严格对齐细则，仅踩细则明确写出的点：①两页各自的文本对象数量足以承载版面上的全部
    #     文字（≥20 个），说明标题/标签/说明文字都是真文本；②两页的非文字矢量形状（自动图形/
    #     线条/自由曲线）数量都足以覆盖上述结构元素（≥30 个）；③两页都没有充当"整片贴图"或
    #     吞掉文字的大图片（单张图片面积不超过画面主体的 25%）；④细则点到的各处文字与形状
    #     此前已逐条按文本对象/形状命中，此处只做"是否被贴成图片"的整体判定，字体/字号/颜色
    #     与填充/描边/尺寸的可改性随文本对象与矢量形状成立。
    _tu30, _vs30, _pr30 = [], [], []
    for _m30 in (M2, M1):
        _b30 = _m30['bbox']
        _ab30 = max((_b30[2] - _b30[0]) * (_b30[3] - _b30[1]), 1.0)
        _tu30.append(len(_m30['units']))
        _vs30.append(len([s for s in _m30['vis'] if not s['has_text']
                          and s['kind'] in ('auto', 'line', 'free')]))
        _pr30.append(max([s['area'] / _ab30 for s in _m30['vis'] if s['kind'] == 'pic'] or [0.0]))
    _ok30 = min(_tu30) >= 20 and min(_vs30) >= 30 and max(_pr30) <= 0.25
    if _ok30:
        results.append('+5：两页文字均为可直接改字的文本对象，方框/面板/箭头/连接线/几何块与'
                       '箱体各面均为可单独编辑的矢量形状（文本对象%d/%d个、非文字矢量形状'
                       '%d/%d个、最大图片仅占画面%.0f%%/%.0f%%）'
                       % (_tu30[0], _tu30[1], _vs30[0], _vs30[1],
                          _pr30[0] * 100, _pr30[1] * 100))
    else:
        results.append('+0：存在文字固化在图片内、结构元素非可编辑矢量形状或整片贴图'
                       '（文本对象%d/%d个、非文字矢量形状%d/%d个、最大图片占画面%.0f%%/%.0f%%）'
                       % (_tu30[0], _tu30[1], _vs30[0], _vs30[1],
                          _pr30[0] * 100, _pr30[1] * 100))

    # 31) +3：装饰性小图标以矢量形状或独立图片对象逐个放置、可单独选中移动缩放，不使用
    #     符号字符（如⚙✣⌕⊞▧）临时代替图标。
    #     严格对齐细则，仅踩细则明确写出的点：①两页文本中不得出现"拿符号字符当图标"的字符
    #     （箭头符、几何图形符、杂项符号与图形符、装饰符、私用区图标字体码位）；②图标须是
    #     形状或图片对象，故两页各须有一批小尺寸非文字图形（≥6 个，面积 0.01~2.0cm²）。
    #     省略号"…"与间隔点"·"属正文标点、不是图标替身，明确放行。
    _BAD34 = ((0x2190, 0x21FF), (0x2295, 0x22FF), (0x2300, 0x23FF), (0x2500, 0x27BF),
              (0x2B00, 0x2BFF), (0x1F000, 0x1FAFF), (0xE000, 0xF8FF))
    _OK34 = ('…', '·', '—', '–')
    _gl34, _ic34 = [], []
    for _m34 in (M2, M1):
        for _u34 in _m34['units']:
            for _ch34 in (_u34['text'] or ''):
                if _ch34 in _OK34:
                    continue
                _cp34 = ord(_ch34)
                if any(a <= _cp34 <= b for a, b in _BAD34) and _ch34 not in _gl34:
                    _gl34.append(_ch34)
        _ic34.append(len([s for s in _m34['vis'] if not s['has_text']
                          and s['kind'] in ('auto', 'line', 'free')
                          and 0.01 <= s['area'] <= 2.00]))
    _ok34 = (not _gl34) and min(_ic34) >= 6
    if _ok34:
        results.append('+3：装饰性小图标以独立矢量形状逐个放置、未用符号字符替代'
                       '（两页小图形%d/%d个、未见符号字符图标）' % (_ic34[0], _ic34[1]))
    else:
        results.append('+0：图标以符号字符替代或未逐个独立放置（符号字符%s、两页小图形%d/%d个）'
                       % (''.join(_gl34) if _gl34 else '无', _ic34[0], _ic34[1]))

    # 32) +3：两页观感规整专业：同类卡片尺寸统一、行列对齐，图标风格一致，实线/虚线用法与
    #     指定处一致，文字与图形无相互压盖、无空框漏字。
    #     严格对齐细则，仅踩细则明确写出的点：①"实线/虚线用法与指定处一致"——细则共指定了
    #     6 处虚线（竖栏2条水平虚线、4框向下汇聚折线、橙色说明框边、左右两栏说明框边、中部
    #     大圆角区域边、中央圆外圆环），要求其中≥5 处确为虚线画法，其余连接处保持实线；
    #     ②"无空框漏字"——两页中有填充的成型方框（面积 2~20cm²，即卡片/子项框一类）都须
    #     内含文字或内含更小的子图形，允许≤1 处纯装饰底板；③"图标风格一致"——两页都须有
    #     成批的独立小图形承担图标（≥6 个）；④"同类卡片尺寸统一、行列对齐"已在第08/10/14/
    #     15/19/20/21/24 条按具体位置逐条判定，此处不重复计分；文字逐字拆框摆放属复刻常见
    #     画法，其字框相互交叠不算压盖，故不据此扣分。
    _dd35 = [_bar09 >= 2, _dash11, _dsh12,
             (_nbL is not None and (_nbL['dashed'] or _perim_dashed(M1, _nbL['ink'], 8))
              and _nbR is not None and (_nbR['dashed'] or _perim_dashed(M1, _nbR['ink'], 8))),
             _ed25 == 4, _rk26]
    _nd35 = len([z for z in _dd35 if z])
    _eb35, _ic35 = 0, []
    for _m35 in (M2, M1):
        _ic35.append(len([s for s in _m35['vis'] if not s['has_text']
                          and s['kind'] in ('auto', 'line', 'free')
                          and 0.01 <= s['area'] <= 2.00]))
        for s in _m35['vis']:
            if s['has_text'] or s['kind'] != 'auto' or s['fill'] is None:
                continue
            if not (2.00 <= s['area'] <= 20.00):
                continue
            if any(_box_in((u['x0'], u['y0'], u['x1'], u['y1']), s['ink'], 0.10)
                   for u in _m35['units'] if (u['text'] or '').strip()):
                continue
            if any(z['sid'] != s['sid'] and z['area'] <= 0.60 * s['area']
                   and _pt_in(z['cx'], z['cy'], s['ink']) for z in _m35['vis']):
                continue
            _eb35 += 1
    _ok35 = _nd35 >= 5 and _eb35 <= 1 and min(_ic35) >= 6
    if _ok35:
        results.append('+3：两页观感规整——指定处虚线用法%d/6处到位、无空框漏字、'
                       '图标均为独立小图形（两页%d/%d个）' % (_nd35, _ic35[0], _ic35[1]))
    else:
        results.append('+0：整体观感存在问题（指定处虚线%d/6、空框%d个、独立小图形%d/%d个）'
                       % (_nd35, _eb35, _ic35[0], _ic35[1]))

    # ========== 汇总：把 results 反解为结构化 dim2_items ==========
    if len(results) != len(_MAX_DELTAS_ORDER):
        raise RuntimeError(f"dim2 项数不一致：results={len(results)} vs order={len(_MAX_DELTAS_ORDER)}")
    dim2_items = []
    total_score = 0
    for raw, max_delta in zip(results, _MAX_DELTAS_ORDER):
        head, _, rule = raw.partition("：")
        try:
            delta = int(head.replace("+", "").strip())
        except ValueError:
            delta = 0
        hit = delta > 0
        dim2_items.append({"rule": rule, "max_delta": max_delta,
                           "delta": max_delta if hit else 0, "hit": hit})
        total_score += (max_delta if hit else 0)
    result["dim2_items"] = dim2_items
    result["total_score"] = total_score
    result["max_score"] = MAX_SCORE_TOTAL
    return result


if __name__ == "__main__":
    if len(sys.argv) >= 2:
        _dir = sys.argv[1]
    else:
        _dir = os.path.dirname(os.path.abspath(__file__))
    _out = evaluate(_dir)
    print(json.dumps(_out, ensure_ascii=False, indent=2))

