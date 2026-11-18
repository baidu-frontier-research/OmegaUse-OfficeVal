#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# pyright: reportUnknownMemberType=false, reportUnknownVariableType=false, reportUnknownArgumentType=false, reportUnknownParameterType=false, reportMissingParameterType=false, reportPrivateUsage=false, reportUnknownLambdaType=false, reportGeneralTypeIssues=false

import os
import sys
import json
from pptx import Presentation

# 说明：按接口统一约定，本模块不修改全局 sys.stdout，也不 print 主结果。
# 主结果统一由 evaluate(dir_path) -> dict 返回；入参为"脚本所在目录的路径"，
# 由脚本自己在该目录内定位并打开被评估的 .pptx。本文件仅在 __main__ 下作为
# 本地调试入口，才把 dict 序列化为 JSON 打印。

SCRIPT_ID = "071"
# 维度二所有加分项 max_delta 之和；用于维度一未通过 / 脚本异常时兜底
MAX_SCORE_TOTAL = 44

# 维度二所有评分项按"结果追加顺序"的 max_delta 列表；末端把 results 里的
# "+N：..." / "-N：..." 字符串反解为结构化 dim2_items 时用。顺序须与
# _evaluate_impl 中 results.append(...) 的顺序严格一致；共 24 项，
# 其中正数为加分项（累加 = MAX_SCORE_TOTAL），负数为减分项。
_MAX_DELTAS_ORDER = [
    3,   # 01 首页主标题蓝色字芯
    5,   # 02 首页主标题白色描边
    3,   # 03 描边宽度 0.5-2.0 磅
    1,   # 04 主标题居中/落位/自动换行
    5,   # 05 第二页左右两栏并排
    3,   # 06 两板块半宽且齐平
    5,   # 07 两组元素各在本侧不越中缝
    5,   # 08 两张照片各在本板块右上角
    3,   # 09 左右对称（照片/标题/正文首行）
    5,   # 10 第二页文案逐字一致
    3,   # 11 正文块收窄至栏内且左对齐
    3,   # 12 文字与板块底色对比充分
    -3,  # 13 主标题字体 Kaiti SC 60 磅不加粗
    -3,  # 14 首页重复叠放的同文标题框
    -1,  # 15 首页副标题格式
    -1,  # 16 首页背景图与空白占位形状
    -3,  # 17 第二页页面标题格式与定位
    -3,  # 18 第二页正文字体字号字色
    -3,  # 19 两侧板块标题字色与字体
    -3,  # 20 护鸟带板块底衬样式
    -3,  # 21 第二页新增装饰元素
    -3,  # 22 两张照片外形与描边
    -1,  # 23 护鸟带要点被合并为一个文本框
    -1,  # 24 第三、四页整页背景图
]
assert sum(d for d in _MAX_DELTAS_ORDER if d > 0) == MAX_SCORE_TOTAL

# ===================== 基础常量与几何/颜色工具 =====================

A_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main'
P_NS = 'http://schemas.openxmlformats.org/presentationml/2006/main'
EMU_PER_CM = 360000.0
TOL = 0.10          # 通用长度容差（cm）
PAGE_W_CM = 33.87   # 页面宽（cm）
PAGE_H_CM = 19.05   # 页面高（cm）

# 第二页各元素的文本锚点：全部按文案定位，不依赖形状名称与 z 序
KEY_PAGE_TITLE = '田间防护'
KEY_A_TITLE = '护鸟带「'
KEY_B_TITLE = '驱鸟器「'
KEYS_A_BODY = ('材质：', '原理：', '设计：', '该装置适合', '布设灵活', '维护简单')
KEYS_B_BODY = ('内置分时播放模块',)
KEY_MAIN_TITLE = '田间共生观察'
KEY_SUBTITLE = '生态友好型农事管理片段'

# 第二页应当逐字保留的 10 条文案（判定"无删改、无改写、无新增"）
EXPECTED_SLIDE2_TEXTS = [
    '田间防护：护鸟带与驱鸟器应用说明',
    '护鸟带「护鸟不伤鸟」',
    '材质：柔性反光复合膜，边缘圆润，适合短期田间布设。',
    '原理：借助风动飘摆与光线闪烁，形成温和视觉干扰，引导鸟群绕行。',
    '设计：带状结构不设网孔，便于回收和重复安装。',
    '该装置适合在果蔬、稻田等开阔区域配合田埂或支杆使用，重点用于降低啄食和踩踏造成的局部损耗。',
    '•布设灵活：可按作物高度调整悬挂位置，风力变化时保持自然摆动。',
    '•维护简单：定期检查固定点和清洁表面即可，避免长期遗留在田间。',
    '驱鸟器「驱鸟还不吵」',
    '内置分时播放模块，可在清晨与傍晚等重点时段发出短促提示音。音量可调，适合与反光带、巡田管理组合使用。',
]

# schemeClr 到主题色槽的映射（含 clrMap 的常规约定：tx1→dk1、bg1→lt1）
_SCHEME_SLOT = {
    'tx1': 'dk1', 'bg1': 'lt1', 'tx2': 'dk2', 'bg2': 'lt2',
    'dk1': 'dk1', 'lt1': 'lt1', 'dk2': 'dk2', 'lt2': 'lt2',
}


def _q(tag):
    """DrawingML 限定名."""
    return '{%s}%s' % (A_NS, tag)


def _qp(tag):
    """PresentationML 限定名."""
    return '{%s}%s' % (P_NS, tag)


def _cm(emu):
    """EMU → cm；None 安全."""
    return None if emu is None else float(emu) / EMU_PER_CM

def _box(shape):
    """形状外框 (left, top, right, bottom)，单位 cm；缺少定位信息返回 None."""
    if shape.left is None or shape.top is None or shape.width is None or shape.height is None:
        return None
    l = _cm(shape.left)
    t = _cm(shape.top)
    return (l, t, l + _cm(shape.width), t + _cm(shape.height))


def _inter_area(b1, b2):
    """两外框相交面积（cm²）."""
    if not b1 or not b2:
        return 0.0
    w = min(b1[2], b2[2]) - max(b1[0], b2[0])
    h = min(b1[3], b2[3]) - max(b1[1], b2[1])
    return w * h if (w > 0 and h > 0) else 0.0


def _area(b):
    return 0.0 if not b else max(0.0, b[2] - b[0]) * max(0.0, b[3] - b[1])


def _union_box(boxes):
    boxes = [b for b in boxes if b]
    if not boxes:
        return None
    return (min(b[0] for b in boxes), min(b[1] for b in boxes),
            max(b[2] for b in boxes), max(b[3] for b in boxes))


def _v_overlap_ratio(b1, b2):
    """两外框纵向区间的重合度（以较矮者高度为分母）."""
    if not b1 or not b2:
        return 0.0
    ov = min(b1[3], b2[3]) - max(b1[1], b2[1])
    base = min(b1[3] - b1[1], b2[3] - b2[1])
    return 0.0 if base <= 0 else max(0.0, ov) / base


def _center(b):
    return ((b[0] + b[2]) / 2.0, (b[1] + b[3]) / 2.0)


def _in_range(v, lo, hi):
    return v is not None and lo <= v <= hi

def _load_theme_colors(prs):
    """读取 theme1.xml 的 clrScheme，得到 {'dk1': '000000', 'lt1': 'FFFFFF', ...}."""
    slots = {}
    try:
        from lxml import etree
        for part in prs.part.package.iter_parts():
            if str(part.partname).endswith('theme1.xml'):
                root = etree.fromstring(part.blob)
                scheme = root.find('.//' + _q('clrScheme'))
                if scheme is None:
                    break
                for child in scheme:
                    slot = etree.QName(child).localname
                    srgb = child.find(_q('srgbClr'))
                    sysc = child.find(_q('sysClr'))
                    if srgb is not None and srgb.get('val'):
                        slots[slot] = srgb.get('val').upper()
                    elif sysc is not None:
                        slots[slot] = (sysc.get('lastClr') or '000000').upper()
                break
    except Exception:
        pass
    # 兜底：主题缺失时给出常规明暗值，避免对比度判定误伤
    slots.setdefault('dk1', '000000')
    slots.setdefault('lt1', 'FFFFFF')
    return slots


def _color_hex(color_el, theme):
    """把 <a:srgbClr>/<a:schemeClr> 元素解析为 6 位十六进制大写串."""
    if color_el is None:
        return None
    tag = color_el.tag
    if tag == _q('srgbClr'):
        val = color_el.get('val')
        return val.upper() if val else None
    if tag == _q('schemeClr'):
        val = (color_el.get('val') or '').lower()
        return theme.get(_SCHEME_SLOT.get(val, val))
    if tag == _q('sysClr'):
        val = color_el.get('lastClr')
        return val.upper() if val else None
    return None


def _solid_fill_hex(holder, theme):
    """holder（rPr / spPr / a:ln 等）直接子级 <a:solidFill> 的颜色."""
    if holder is None:
        return None
    fill = holder.find(_q('solidFill'))
    if fill is None:
        return None
    for child in fill:
        hexv = _color_hex(child, theme)
        if hexv:
            return hexv
    return None

def _rgb_tuple(hexv):
    if not hexv or len(hexv) != 6:
        return None
    try:
        return (int(hexv[0:2], 16), int(hexv[2:4], 16), int(hexv[4:6], 16))
    except ValueError:
        return None


def _rel_luminance(hexv):
    """WCAG 相对亮度，取值 0-1."""
    rgb = _rgb_tuple(hexv)
    if rgb is None:
        return None
    out = []
    for c in rgb:
        c = c / 255.0
        out.append(c / 12.92 if c <= 0.03928 else ((c + 0.055) / 1.055) ** 2.4)
    return 0.2126 * out[0] + 0.7152 * out[1] + 0.0722 * out[2]


def _contrast_ratio(hex1, hex2):
    """两色对比度（1-21）；任一色未知返回 None."""
    l1 = _rel_luminance(hex1)
    l2 = _rel_luminance(hex2)
    if l1 is None or l2 is None:
        return None
    hi, lo = max(l1, l2), min(l1, l2)
    return (hi + 0.05) / (lo + 0.05)


def _hsv(hexv):
    """RGB → (色相 0-360, 饱和度 0-1, 明度 0-1)."""
    rgb = _rgb_tuple(hexv)
    if rgb is None:
        return None
    r, g, b = [c / 255.0 for c in rgb]
    mx, mn = max(r, g, b), min(r, g, b)
    d = mx - mn
    if d == 0:
        h = 0.0
    elif mx == r:
        h = 60 * (((g - b) / d) % 6)
    elif mx == g:
        h = 60 * (((b - r) / d) + 2)
    else:
        h = 60 * (((r - g) / d) + 4)
    return (h, 0.0 if mx == 0 else d / mx, mx)


def _is_blue(hexv):
    """蓝色系纯色：色相 195°-235°、饱和度≥0.30、明度≥0.20."""
    hsv = _hsv(hexv)
    if hsv is None:
        return False
    h, s, v = hsv
    return 195.0 <= h <= 235.0 and s >= 0.30 and v >= 0.20


def _is_white(hexv):
    """白色（含近白）：三通道均≥245."""
    rgb = _rgb_tuple(hexv)
    return bool(rgb) and all(c >= 245 for c in rgb)

# ===================== 形状 / 文字属性读取 =====================

def _norm(text):
    """文本归一化：去首尾空白、去全角空格、统一西文引号，便于逐字比对."""
    if text is None:
        return ''
    s = text.replace('　', ' ').replace('\xa0', ' ').strip()
    return ' '.join(s.split())


def _shape_text(shape):
    """形状内全部文字（段落间用 \\n 连接）；无文本框返回空串."""
    try:
        if not shape.has_text_frame:
            return ''
        return shape.text_frame.text or ''
    except Exception:
        return ''


def _paragraph_texts(shape):
    """形状内各段落的归一化文本（跳过空段）."""
    out = []
    try:
        if not shape.has_text_frame:
            return out
        for p in shape.text_frame.paragraphs:
            t = _norm(''.join(r.text for r in p.runs))
            if t:
                out.append(t)
    except Exception:
        pass
    return out


def _is_picture(shape):
    try:
        return shape.shape_type is not None and int(shape.shape_type) == 13
    except Exception:
        return shape.__class__.__name__ == 'Picture'


def _sp_pr(shape):
    """形状的 spPr 元素（图片为 p:pic/p:spPr）."""
    el = shape._element
    for tag in (_qp('spPr'), _q('spPr')):
        found = el.find(tag)
        if found is not None:
            return found
    return el.find('.//' + _qp('spPr'))


def _geom_prst(shape):
    """预设几何名（rect / roundRect / ...）与圆角 adj 值（千分比，None 表示未显式给出）."""
    spPr = _sp_pr(shape)
    if spPr is None:
        return (None, None)
    geom = spPr.find(_q('prstGeom'))
    if geom is None:
        return (None, None)
    adj = None
    gd = geom.find('.//' + _q('gd'))
    if gd is not None and (gd.get('name') or '') == 'adj':
        fmla = gd.get('fmla') or ''
        if fmla.startswith('val '):
            try:
                adj = int(fmla[4:])
            except ValueError:
                adj = None
    return (geom.get('prst'), adj)

def _shape_fill_hex(shape, theme):
    """形状实心填充色；非实心填充（noFill/渐变/图片）返回 None."""
    return _solid_fill_hex(_sp_pr(shape), theme)


def _shape_line(shape, theme):
    """形状边框：(是否为无填充边框, 边框颜色, 线宽磅值)."""
    spPr = _sp_pr(shape)
    if spPr is None:
        return (True, None, None)
    ln = spPr.find(_q('ln'))
    if ln is None:
        return (True, None, None)
    no_fill = ln.find(_q('noFill')) is not None
    hexv = _solid_fill_hex(ln, theme)
    w = ln.get('w')
    pt = None
    if w is not None:
        try:
            pt = int(w) / 12700.0
        except ValueError:
            pt = None
    return (no_fill or hexv is None, hexv, pt)


def _has_outer_shadow(shape):
    spPr = _sp_pr(shape)
    return spPr is not None and spPr.find('.//' + _q('outerShdw')) is not None


def _body_wrap(shape):
    """文本框 bodyPr 的 wrap 值；未显式给出时按 OOXML 默认 'square' 处理."""
    try:
        if not shape.has_text_frame:
            return None
        body = shape.text_frame._txBody.find(_q('bodyPr'))
        if body is None:
            return 'square'
        return body.get('wrap') or 'square'
    except Exception:
        return None


def _runs(shape):
    """形状内全部非空 run."""
    out = []
    try:
        if not shape.has_text_frame:
            return out
        for p in shape.text_frame.paragraphs:
            for r in p.runs:
                if r.text:
                    out.append(r)
    except Exception:
        pass
    return out

def _inherited_fill_hex(run, theme):
    """run 未显式声明字色时的生效字色：沿 段落 a:pPr/a:defRPr → 文本体
    a:lstStyle 各级 a:defRPr 逐级回落，仍未命中则落到主题 tx1（dk1）——
    非占位符形状的默认文字色即母版 otherStyle 的 tx1。"""
    r = run._r
    para = r.getparent()
    if para is not None:
        pPr = para.find(_q('pPr'))
        if pPr is not None:
            hexv = _solid_fill_hex(pPr.find(_q('defRPr')), theme)
            if hexv:
                return hexv
        body = para.getparent()
        if body is not None:
            lst = body.find(_q('lstStyle'))
            if lst is not None:
                for lvl in lst:
                    hexv = _solid_fill_hex(lvl.find(_q('defRPr')), theme)
                    if hexv:
                        return hexv
    return theme.get('dk1') or '000000'


def _run_props(run, theme):
    """单个 run 的关键属性：字号(磅)、加粗、中西文字体名、字色、文字描边色与线宽。
    fill 为"生效字色"：run 级未显式声明 solidFill 时按继承链回落，不视为无色。"""
    rPr = run._r.find(_q('rPr'))
    info = {
        'text': run.text, 'sz': None, 'bold': False,
        'latin': None, 'ea': None, 'fill': _inherited_fill_hex(run, theme),
        'ln_hex': None, 'ln_pt': None, 'has_ln': False,
    }
    if rPr is None:
        return info
    sz = rPr.get('sz')
    if sz is not None:
        try:
            info['sz'] = int(sz) / 100.0
        except ValueError:
            info['sz'] = None
    info['bold'] = (rPr.get('b') or '0') in ('1', 'true')
    latin = rPr.find(_q('latin'))
    ea = rPr.find(_q('ea'))
    info['latin'] = latin.get('typeface') if latin is not None else None
    info['ea'] = ea.get('typeface') if ea is not None else None
    info['fill'] = _solid_fill_hex(rPr, theme) or _inherited_fill_hex(run, theme)
    ln = rPr.find(_q('ln'))
    if ln is not None:
        info['has_ln'] = ln.find(_q('noFill')) is None
        info['ln_hex'] = _solid_fill_hex(ln, theme)
        w = ln.get('w')
        if w is not None:
            try:
                info['ln_pt'] = int(w) / 12700.0
            except ValueError:
                info['ln_pt'] = None
        else:
            info['ln_pt'] = 0.75   # 未标注线宽时按 OOXML 细线默认值处理
    return info


def _font_of(props):
    """中文实际生效字体：优先 ea，回落 latin."""
    return props.get('ea') or props.get('latin')


def _slide_bg_hex(slide, theme):
    """幻灯片背景实心填充色；缺页或无显式背景返回 None."""
    if slide is None:
        return None
    bg = slide._element.find('.//' + _qp('bgPr'))
    return _solid_fill_hex(bg, theme) if bg is not None else None

# ===================== 对外接口 =====================

class _TargetMissing(Exception):
    """目录不存在或目录中没有 .pptx：连待评测对象都没有，不进入维度判定。"""


def evaluate(dir_path):
    """按接口约定返回结构化 dict；不 print、不改 sys.stdout、不 sys.exit。
    入参 dir_path 为"脚本所在目录的路径"，脚本自己在该目录里定位并打开被评估的
    .pptx（忽略以 ~$ 开头的临时锁文件）。目录不存在 / 未发现文档时返回
    status='error' 且维度一未参与判定（dim1_pass 留空）；文件存在但打不开或
    脚本崩溃时返回 status='error' 且 dim1_pass=False 并给出错误摘要。"""
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
        return _evaluate_impl(file_path, result)
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
        # 文件存在但打不开 / 解析失败：属于交付文件问题，记为维度一不通过、正常 0 分
        result["status"] = "ok"
        result["error"] = None
        result["dim1_pass"] = False
        result["dim1_reason"] = f"文件无法解析或无法打开：{type(exc).__name__}: {exc}"
        result["dim2_items"] = []
        result["total_score"] = 0
        return result


def _fail_dim1(result, reason):
    """维度一未通过：完成度总分直接记 0（属于正常评测结果，不是脚本错误）。"""
    result["status"] = "ok"
    result["error"] = None
    result["dim1_pass"] = False
    result["dim1_reason"] = reason
    result["dim2_items"] = []
    result["total_score"] = 0
    result["max_score"] = MAX_SCORE_TOTAL
    return result

def _evaluate_impl(file_path, result):
    """维度一门禁 + 维度二评分主体；内部按追加顺序构造 results 列表，
    末尾按 _MAX_DELTAS_ORDER 反解成结构化 dim2_items 填入 result。"""
    # ========== 维度1：可用与可修改性 ==========
    # 仅门禁"文件能否作为 .pptx 正常打开"，打不开则完成度总分直接记 0；
    # 页数、页面尺寸、版面重叠越界等一律不设门禁，交由维度二各评分项据实计分。
    try:
        prs = Presentation(file_path)
        slides = list(prs.slides)
    except Exception as exc:
        return _fail_dim1(result, f"文件损坏或非有效 .pptx，无法打开（{type(exc).__name__}: {exc}）")

    theme = _load_theme_colors(prs)
    results = []

    page_w = _cm(prs.slide_width)
    page_h = _cm(prs.slide_height)

    def _slide_at(idx):
        return slides[idx] if idx < len(slides) else None

    def _shapes_at(idx):
        s = _slide_at(idx)
        return list(s.shapes) if s is not None else []

    s2 = _slide_at(1)
    s1_shapes = _shapes_at(0)
    s2_shapes = _shapes_at(1)

    def _text_shapes(shapes):
        return [sh for sh in shapes if not _is_picture(sh) and _norm(_shape_text(sh))]

    s1_texts = _text_shapes(s1_shapes)
    s2_texts = _text_shapes(s2_shapes)

    # ========== 维度2：完成度评分 ==========

    # --- 首页元素定位（按文案，不依赖形状名）---
    s1_pics = [sh for sh in s1_shapes if _is_picture(sh)]
    title_shapes = [sh for sh in s1_texts if KEY_MAIN_TITLE in _norm(_shape_text(sh))]
    sub_shapes = [sh for sh in s1_texts if KEY_SUBTITLE in _norm(_shape_text(sh))]
    main_title = title_shapes[-1] if title_shapes else None   # z 序最上层者为可见标题
    subtitle = sub_shapes[-1] if sub_shapes else None
    title_props = [_run_props(r, theme) for r in _runs(main_title)] if main_title else []
    sub_props = [_run_props(r, theme) for r in _runs(subtitle)] if subtitle else []

    # --- 第二页元素定位 ---
    page_area = page_w * page_h
    s2_pics = [sh for sh in s2_shapes if _is_picture(sh)]
    content_pics = [sh for sh in s2_pics if _area(_box(sh)) < 0.9 * page_area]

    def _hit_keys(shape, keys):
        t = _norm(_shape_text(shape))
        return any(k in t for k in keys)

    page_title_sh = next((sh for sh in s2_texts if KEY_PAGE_TITLE in _norm(_shape_text(sh))), None)
    a_title_sh = next((sh for sh in s2_texts if KEY_A_TITLE in _norm(_shape_text(sh))), None)
    b_title_sh = next((sh for sh in s2_texts if KEY_B_TITLE in _norm(_shape_text(sh))), None)
    a_body_shs = [sh for sh in s2_texts if _hit_keys(sh, KEYS_A_BODY)]
    b_body_shs = [sh for sh in s2_texts if _hit_keys(sh, KEYS_B_BODY)]
    group_a = [sh for sh in ([a_title_sh] if a_title_sh else []) + a_body_shs]
    group_b = [sh for sh in ([b_title_sh] if b_title_sh else []) + b_body_shs]
    box_a = _union_box([_box(sh) for sh in group_a])
    box_b = _union_box([_box(sh) for sh in group_b])

    # 板块底衬：无文字、非图片、实心填充且面积≥40cm² 的形状
    panels = [
        sh for sh in s2_shapes
        if not _is_picture(sh) and not _norm(_shape_text(sh))
        and _shape_fill_hex(sh, theme) and _area(_box(sh)) >= 40.0
    ]
    panels.sort(key=lambda sh: _box(sh)[0])

    def _panel_of(box):
        """外框中心所落入的板块；无包含关系时取中心距离最近者."""
        if box is None or not panels:
            return None
        cx, cy = _center(box)
        for sh in panels:
            b = _box(sh)
            if b[0] - TOL <= cx <= b[2] + TOL and b[1] - TOL <= cy <= b[3] + TOL:
                return sh
        return min(panels, key=lambda sh: abs(_center(_box(sh))[0] - cx) + abs(_center(_box(sh))[1] - cy))

    def _panel_containing(box):
        """外框中心确实压在某块底衬上时返回该底衬，否则返回 None。
        用于取"文字实际叠在什么颜色上"，不做就近兜底——文字落在底衬之外时，
        它的背景是幻灯片底色而非任何板块。"""
        if box is None:
            return None
        cx, cy = _center(box)
        for sh in panels:
            b = _box(sh)
            if b[0] - TOL <= cx <= b[2] + TOL and b[1] - TOL <= cy <= b[3] + TOL:
                return sh
        return None

    panel_a = _panel_of(box_a)
    panel_b = _panel_of(box_b)
    two_panels = (panel_a is not None and panel_b is not None and panel_a is not panel_b)
    pa, pb = (_box(panel_a) if panel_a else None), (_box(panel_b) if panel_b else None)
    panels_side_by_side = bool(
        two_panels and (pa[2] <= pb[0] + TOL or pb[2] <= pa[0] + TOL)
        and _v_overlap_ratio(pa, pb) >= 0.9)
    left_panel, right_panel = (pa, pb) if (pa and pb and pa[0] <= pb[0]) else (pb, pa)
    mid_x = ((left_panel[2] + right_panel[0]) / 2.0
             if panels_side_by_side else page_w / 2.0)

    def _align_name(paragraph):
        """段落对齐方式名（LEFT/CENTER/...）；None 表示继承默认（按左对齐处理）."""
        al = paragraph.alignment
        return None if al is None else str(al).split()[0]

    # 01 +3：首页主标题蓝色字芯
    ok = bool(title_props) and all(_is_blue(p['fill']) for p in title_props)
    results.append(
        ('+3' if ok else '+0')
        + '：首页主标题"田间共生观察"的文字填充为蓝色系纯色（色相200°-230°区间），6个字全部为蓝色')

    # 02 +5：首页主标题白色描边
    ok = bool(title_props) and all(p['has_ln'] and _is_white(p['ln_hex']) for p in title_props)
    results.append(
        ('+5' if ok else '+0')
        + '：首页主标题"田间共生观察"设有白色文字描边（run 级 a:ln 白色实线），6个字全部带描边，与蓝色字芯构成艺术字')

    # 03 +3：描边宽度 0.5-2.0 磅
    ok = bool(title_props) and all(_in_range(p['ln_pt'], 0.5, 2.0) for p in title_props)
    results.append(
        ('+3' if ok else '+0')
        + '：主标题描边宽度0.5-2.0磅，白边不侵蚀笔画，未出现≥3磅粗边把文字压成一团白边')

    # 04 +1：主标题居中、落位与自动换行
    # 前提：主标题确已做过艺术字改动（run 级描边或蓝色字芯）。未动过标题的文件
    # 天然满足"落位没偏"，不应凭"什么都没做"拿分。
    title_reworked = bool(title_props) and any(
        p['has_ln'] or _is_blue(p['fill']) for p in title_props)
    ok = False
    if title_reworked and main_title is not None:
        b = _box(main_title)
        paras = [p for p in main_title.text_frame.paragraphs if _norm(''.join(r.text for r in p.runs))]
        ok = bool(
            b and len(paras) == 1
            and _norm(_shape_text(main_title)) == KEY_MAIN_TITLE
            and all(_align_name(p) == 'CENTER' for p in paras)
            and _in_range(b[0], 2.6, 3.0) and _in_range(b[1], 6.2, 6.5)
            and _in_range(b[2] - b[0], 28.0, 28.5) and _in_range(b[3] - b[1], 2.6, 3.0)
            and _body_wrap(main_title) == 'square')
    results.append(
        ('+1' if ok else '+0')
        + '：在主标题已完成艺术字改动的前提下，主标题水平居中并落位在标题带内（左边界2.6-3.0cm、上边界6.2-6.5cm、尺寸28.0-28.5cm×2.6-3.0cm），文本框保持自动换行且未多出换行')

    # 05 +5：第二页左右两栏并排
    groups_side = bool(
        box_a and box_b
        and (box_a[2] <= box_b[0] + TOL or box_b[2] <= box_a[0] + TOL)
        and _v_overlap_ratio(box_a, box_b) >= 0.9)
    ok = bool(groups_side and panels_side_by_side)
    results.append(
        ('+5' if ok else '+0')
        + '：第二页内容区为左右两栏并排的两个独立板块（左板块右边界≤右板块左边界、纵向重合度≥90%），护鸟带与驱鸟器各占一块，不是上下堆叠的通栏色块')

    # 第二页是否确已做过分栏改版：出现两块相互独立的板块底衬，或两组元素已左右并排。
    # 未改版的文件天然"文案逐字一致"，不应凭"什么都没做"拿第 10 条的分。
    s2_reworked = bool(two_panels or groups_side)

    # 06 +3：两板块各占半宽且相互齐平
    ok = False
    if panels_side_by_side:
        lw, rw = left_panel[2] - left_panel[0], right_panel[2] - right_panel[0]
        lh, rh = left_panel[3] - left_panel[1], right_panel[3] - right_panel[1]
        gap = right_panel[0] - left_panel[2]
        ok = bool(
            _in_range(lw, 15.0, 16.5) and _in_range(rw, 15.0, 16.5) and abs(lw - rw) <= 1.0
            and _in_range(left_panel[0], 0.7, 1.2) and _in_range(right_panel[2], 32.5, 33.2)
            and abs(left_panel[1] - right_panel[1]) <= 0.2 and abs(lh - rh) <= 0.2
            and _in_range(lh, 13.0, 13.6) and _in_range(rh, 13.0, 13.6)
            and _in_range(gap, 0.3, 1.5))
    results.append(
        ('+3' if ok else '+0')
        + '：两板块各占页宽约一半且相互齐平（单块宽15.0-16.5cm、宽度差≤1.0cm，左板块左边界0.7-1.2cm、右板块右边界32.5-33.2cm，上边界差≤0.2cm、高度差≤0.2cm、高13.0-13.6cm，中缝0.3-1.5cm不重叠）')

    # 07 +5：两组元素各在本侧、不越中缝
    def _inside_panel(shapes, panel_box):
        if not shapes or panel_box is None:
            return False
        for sh in shapes:
            b = _box(sh)
            if b is None or b[0] < panel_box[0] - TOL or b[2] > panel_box[2] + TOL:
                return False
        return True

    a_left = bool(group_a) and all(_box(sh)[2] <= mid_x + TOL for sh in group_a)
    a_right = bool(group_a) and all(_box(sh)[0] >= mid_x - TOL for sh in group_a)
    b_left = bool(group_b) and all(_box(sh)[2] <= mid_x + TOL for sh in group_b)
    b_right = bool(group_b) and all(_box(sh)[0] >= mid_x - TOL for sh in group_b)
    ok = bool(
        panels_side_by_side and group_a and group_b
        and _inside_panel(group_a, pa) and _inside_panel(group_b, pb)
        and ((a_left and b_right) or (a_right and b_left)))
    results.append(
        ('+5' if ok else '+0')
        + '：护鸟带的板块标题与其全部正文完全落在左板块内，驱鸟器的板块标题与其正文完全落在右板块内，无任何元素跨越中缝或伸入对侧')

    # 08 +5：两张照片各在本板块右上角空白区
    ok = False
    if len(content_pics) == 2 and two_panels:
        pic_panel = {}
        for pic in content_pics:
            pic_panel[id(pic)] = _panel_of(_box(pic))
        mapped = {id(panel_a): None, id(panel_b): None}
        for pic in content_pics:
            pn = pic_panel[id(pic)]
            if pn is not None and id(pn) in mapped and mapped[id(pn)] is None:
                mapped[id(pn)] = pic
        if all(v is not None for v in mapped.values()):
            ok = True
            for panel, grp in ((panel_a, group_a), (panel_b, group_b)):
                pic = mapped[id(panel)]
                pb_box, pic_box = _box(panel), _box(pic)
                cx = (pb_box[0] + pb_box[2]) / 2.0
                if not (pic_box[0] >= cx - TOL and pic_box[2] <= pb_box[2] + TOL
                        and (pic_box[1] - pb_box[1]) <= 1.5 + TOL):
                    ok = False
                if any(_inter_area(pic_box, _box(sh)) > 0.05 for sh in grp):
                    ok = False
    results.append(
        ('+5' if ok else '+0')
        + '：两张照片分别置于各自板块的右上角空白区（左边界≥所属板块横向中点、右边界不超出该板块、上边界距板块顶端≤1.5cm），与本块标题、正文均无重叠压字')

    # 09 +3：左右对称（照片规格 / 板块标题 / 正文首行）
    ok = False
    if len(content_pics) == 2 and a_title_sh is not None and b_title_sh is not None:
        p1, p2 = _box(content_pics[0]), _box(content_pics[1])
        w1, h1 = p1[2] - p1[0], p1[3] - p1[1]
        w2, h2 = p2[2] - p2[0], p2[3] - p2[1]
        at, bt = _box(a_title_sh), _box(b_title_sh)
        a_top = min([_box(sh)[1] for sh in a_body_shs] or [None]) if a_body_shs else None
        b_top = min([_box(sh)[1] for sh in b_body_shs] or [None]) if b_body_shs else None
        ta = [_run_props(r, theme) for r in _runs(a_title_sh)]
        tb = [_run_props(r, theme) for r in _runs(b_title_sh)]
        ok = bool(
            _in_range(w1, 4.5, 6.0) and _in_range(h1, 4.5, 6.0)
            and _in_range(w2, 4.5, 6.0) and _in_range(h2, 4.5, 6.0)
            and abs(w1 - w2) <= 0.2 and abs(h1 - h2) <= 0.2 and abs(p1[1] - p2[1]) <= 0.3
            and ta and tb
            and all(p['sz'] == 24.0 and p['bold'] for p in ta)
            and all(p['sz'] == 24.0 and p['bold'] for p in tb)
            and abs(at[1] - bt[1]) <= 0.5
            and a_top is not None and b_top is not None and abs(a_top - b_top) <= 0.5)
    results.append(
        ('+3' if ok else '+0')
        + '：左右两栏呈对称版式（两张照片边长4.5-6.0cm、宽高差≤0.2cm、上边界差≤0.3cm；两侧板块标题同为24磅加粗且上边界差≤0.5cm；两侧正文首行上边界差≤0.5cm）')

    # 10 +5：第二页文案逐字一致（前提：第二页确已做过分栏改版）
    collected = []
    for sh in s2_texts:
        collected.extend(_paragraph_texts(sh))
    ok = bool(s2_reworked) and sorted(collected) == sorted(_norm(t) for t in EXPECTED_SLIDE2_TEXTS)
    results.append(
        ('+5' if ok else '+0')
        + f'：在第二页已完成分栏改版的前提下，第二页文字与指定的10条文案逐字一致，无删减、改写、缩写或新增（实际条目{len(collected)}条）')

    # 11 +3：正文块收窄至栏内且左对齐
    body_shs = a_body_shs + b_body_shs
    ok = bool(body_shs)
    for sh in body_shs:
        b = _box(sh)
        if b is None or (b[2] - b[0]) > 14.5 + TOL:
            ok = False
            break
        if any(_align_name(p) not in (None, 'LEFT')
               for p in sh.text_frame.paragraphs if _norm(''.join(r.text for r in p.runs))):
            ok = False
            break
        own = pa if sh in a_body_shs else pb
        if own is not None and (b[0] < own[0] - TOL or b[2] > own[2] + TOL):
            ok = False
            break
        if panels_side_by_side:
            if sh in a_body_shs and not (b[2] <= mid_x + TOL or b[0] >= mid_x - TOL):
                ok = False
                break
    results.append(
        ('+3' if ok else '+0')
        + '：两侧正文块宽度随栏宽收窄至≤14.5cm并保持左对齐，文字在本栏内自然换行，未横向溢出所属板块、未压到中缝')

    # 12 +3：文字与板块底色对比充分（前提：确实存在左右两块底衬）
    s2_bg = _slide_bg_hex(s2, theme)
    ok = bool(two_panels and group_a and group_b)
    for sh in group_a + group_b:
        own_panel = _panel_containing(_box(sh))
        bg_hex = _shape_fill_hex(own_panel, theme) if own_panel is not None else s2_bg
        if not bg_hex:
            continue
        for p in [_run_props(r, theme) for r in _runs(sh)]:
            ratio = _contrast_ratio(p['fill'], bg_hex)
            if ratio is not None and ratio < 3.0:
                ok = False
                break
        if not ok:
            break
    results.append(
        ('+3' if ok else '+0')
        + '：左右两板块内文字与板块底色对比充分（深底配浅色字、浅底配深色字，对比度≥3:1），标题与正文均清晰可读')

    # 13 -3：主标题字体 Kaiti SC、60磅、不加粗
    bad = (not title_props) or any(
        _font_of(p) != 'Kaiti SC' or p['sz'] != 60.0 or p['bold'] for p in title_props)
    results.append(('-3' if bad else '-0')
                   + '：首页主标题"田间共生观察"不满足字体Kaiti SC、字号60磅、不加粗')

    # 14 -3：首页存在重复叠放的同文标题框
    bad = False
    for i in range(len(s1_texts)):
        for j in range(i + 1, len(s1_texts)):
            ti, tj = _norm(_shape_text(s1_texts[i])), _norm(_shape_text(s1_texts[j]))
            if not ti or ti != tj:
                continue
            bi, bj = _box(s1_texts[i]), _box(s1_texts[j])
            base = min(_area(bi), _area(bj))
            if base > 0 and _inter_area(bi, bj) / base >= 0.8:
                bad = True
    results.append(('-3' if bad else '-0')
                   + '：首页存在两个文案同为"田间共生观察"、位置与尺寸完全重叠的文本框（以叠放两层文字的方式凑出白边效果，产生重复元素）')

    # 15 -1：首页副标题格式
    bad = (subtitle is None) or (not sub_props) or any(
        _font_of(p) != 'Noto Sans SC' or p['sz'] != 30.0 or p['fill'] != '000000'
        for p in sub_props)
    results.append(('-1' if bad else '-0')
                   + '：首页副标题"生态友好型农事管理片段"不满足字体Noto Sans SC、字号30磅、字色黑色（000000）')

    # 16 -1：首页背景图与页面下部空白占位形状
    has_bg = any(
        (lambda b: b is not None and abs(b[0]) <= 0.05 and abs(b[1]) <= 0.05
         and abs(b[2] - PAGE_W_CM) <= 0.05 and abs(b[3] - PAGE_H_CM) <= 0.05)(_box(sh))
        for sh in s1_pics)
    has_placeholder = False
    for sh in s1_shapes:
        b = _box(sh)
        if _is_picture(sh) or b is None or _norm(_shape_text(sh)):
            continue
        if (abs(b[0] - 8.71) <= 0.2 and abs(b[1] - 19.76) <= 0.2
                and abs((b[2] - b[0]) - 3.99) <= 0.2 and abs((b[3] - b[1]) - 0.85) <= 0.2):
            has_placeholder = True
    bad = not (has_bg and has_placeholder)
    results.append(('-1' if bad else '-0')
                   + '：首页背景图片不满足左上角(0cm,0cm)定位、尺寸33.87cm×19.05cm铺满整页，或页面下部不满足保留1个左边界约8.71cm、上边界约19.76cm、尺寸约3.99cm×0.85cm的空白占位形状')

    # 17 -3：第二页页面标题格式与定位
    bad = True
    if page_title_sh is not None:
        b = _box(page_title_sh)
        props = [_run_props(r, theme) for r in _runs(page_title_sh)]
        bad = not (
            props and b is not None
            and all(_font_of(p) == 'Noto Sans SC' and p['sz'] == 20.0 and p['bold']
                    and p['fill'] == '135E35' for p in props)
            and abs(b[0] - 1.06) <= TOL and abs(b[1] - 0.35) <= TOL
            and abs((b[2] - b[0]) - 31.75) <= TOL)
    results.append(('-3' if bad else '-0')
                   + '：第二页页面标题"田间防护：护鸟带与驱鸟器应用说明"不满足字体Noto Sans SC、字号20磅、加粗、字色135E35，或不满足左边界1.06cm、上边界0.35cm、宽度31.75cm的顶部通栏单行')

    # 18 -3：第二页正文字体、字号与字色
    bad = not body_shs
    for sh in body_shs:
        own_panel = _panel_containing(_box(sh))
        bg_hex = _shape_fill_hex(own_panel, theme) if own_panel is not None else s2_bg
        bg_lum = _rel_luminance(bg_hex)
        for p in [_run_props(r, theme) for r in _runs(sh)]:
            if _font_of(p) != 'Noto Sans SC' or not _in_range(p['sz'], 13.0, 14.0):
                bad = True
                break
            lum = _rel_luminance(p['fill'])
            if lum is not None and bg_lum is not None:
                if bg_lum < 0.5 and lum < 0.5:      # 深底必须配浅色字
                    bad = True
                    break
                if bg_lum >= 0.5 and lum >= 0.5:    # 浅底必须配深色字
                    bad = True
                    break
        if bad:
            break
    results.append(('-3' if bad else '-0')
                   + '：第二页各段正文不满足字体Noto Sans SC、字号13-14磅，或字色不满足深底板块内为浅色、浅底板块内为深色文字')

    # 19 -3：两侧板块标题字色与字体
    a_t = [_run_props(r, theme) for r in _runs(a_title_sh)] if a_title_sh else []
    b_t = [_run_props(r, theme) for r in _runs(b_title_sh)] if b_title_sh else []
    bad = not (
        a_t and b_t
        and all(p['fill'] == 'FFFFFF' and _font_of(p) == 'Noto Sans SC' and p['bold'] for p in a_t)
        and all(p['fill'] == '135E35' and _font_of(p) == 'Noto Sans SC' and p['bold'] for p in b_t))
    results.append(('-3' if bad else '-0')
                   + '：第二页两侧板块标题字色不满足"护鸟带「护鸟不伤鸟」"为白色（FFFFFF）、"驱鸟器「驱鸟还不吵」"为深绿（135E35），或字体不满足Noto Sans SC加粗')

    # 20 -3：护鸟带板块底衬样式（深绿实心 + 圆角矩形 + 无边框 + 柔和外阴影）
    bad = True
    if panel_a is not None:
        prst, adj = _geom_prst(panel_a)
        no_line, _line_hex, _line_pt = _shape_line(panel_a, theme)
        fill = _shape_fill_hex(panel_a, theme)
        bad = not (
            fill == '166534' and prst == 'roundRect'
            and (adj is None or _in_range(adj / 1000.0, 1.6, 3.6))
            and no_line and _has_outer_shadow(panel_a))
    results.append(('-3' if bad else '-0')
                   + '：护鸟带所在板块底衬不满足深绿色（166534）实心填充、圆角矩形（圆角率约2.6%）、无边框线并带柔和外阴影')

    # 21 -3：第二页新增装饰元素（除两块底衬外的无文字非图片形状）
    extras = [
        sh for sh in s2_shapes
        if not _is_picture(sh) and not _norm(_shape_text(sh))
        and sh is not panel_a and sh is not panel_b
    ]
    bad = bool(extras)
    results.append(('-3' if bad else '-0')
                   + f'：第二页出现指令未要求的新增装饰元素（板块内侧竖向色条、板块标题下方细横分隔线、照片外圈方框、页面标题下方短色块等，实测多余形状{len(extras)}个）')

    # 22 -3：两张照片外形与描边
    bad = len(content_pics) != 2
    for pic in content_pics:
        prst, _adj = _geom_prst(pic)
        no_line, line_hex, _pt = _shape_line(pic, theme)
        if prst != 'roundRect' or not no_line or line_hex is not None:
            bad = True
    results.append(('-3' if bad else '-0')
                   + '：第二页两张照片不满足圆角矩形外形且无边框描边（被改为直角矩形并加白色/彩色外边框）')

    # 23 -1：护鸟带三条要点与三段说明分置于两个独立文本块
    key_pt = ('材质：', '原理：', '设计：')
    key_note = ('该装置适合', '布设灵活', '维护简单')
    pt_shs = {id(sh) for sh in a_body_shs if _hit_keys(sh, key_pt)}
    note_shs = {id(sh) for sh in a_body_shs if _hit_keys(sh, key_note)}
    bad = not (pt_shs and note_shs and not (pt_shs & note_shs))
    results.append(('-1' if bad else '-0')
                   + '：护鸟带的三条要点（材质/原理/设计）与"该装置适合…"等三段说明不满足分置于两个可各自独立选中、独立移动的文本块（被合并进同一个文本框）')

    # 24 -1：第三、四页各仅含 1 张铺满整页的图片（缺页同样计为不满足）
    bad = False
    for idx in (2, 3):
        shs = _shapes_at(idx)
        pics = [sh for sh in shs if _is_picture(sh)]
        if _slide_at(idx) is None or len(shs) != 1 or len(pics) != 1:
            bad = True
            continue
        b = _box(pics[0])
        if b is None or not (
                abs(b[0]) <= TOL and abs(b[1]) <= TOL
                and abs((b[2] - b[0]) - PAGE_W_CM) <= TOL
                and abs((b[3] - b[1]) - PAGE_H_CM) <= TOL):
            bad = True
    results.append(('-1' if bad else '-0')
                   + '：第三页或第四页不满足各仅含1张左上角(0cm,0cm)定位、33.87cm×19.05cm铺满整页的图片且无其他元素')

    # ========== results 反解为结构化 dim2_items ==========
    if len(results) != len(_MAX_DELTAS_ORDER):
        raise RuntimeError(
            f"评分项数量不匹配：results={len(results)}，_MAX_DELTAS_ORDER={len(_MAX_DELTAS_ORDER)}")

    items = []
    total = 0
    for text, max_delta in zip(results, _MAX_DELTAS_ORDER):
        head, rule = text.split('：', 1)
        delta = int(head)
        total += delta
        items.append({
            "rule": rule,
            "max_delta": max_delta,
            "delta": delta,
            "hit": delta != 0,
        })

    result["dim1_pass"] = True
    result["dim1_reason"] = f"交付文件可作为 .pptx 正常打开（共 {len(slides)} 页）"
    result["dim2_items"] = items
    # 减分项据实累加，不做归零截断：总分可为负数
    result["total_score"] = total
    result["max_score"] = MAX_SCORE_TOTAL
    return result


if __name__ == '__main__':
    # 本地调试入口：python officeval_071_verifier.py [被评估文档所在目录]
    _dir = sys.argv[1] if len(sys.argv) > 1 else os.path.dirname(os.path.abspath(__file__))
    print(json.dumps(evaluate(_dir), ensure_ascii=False, indent=2))


