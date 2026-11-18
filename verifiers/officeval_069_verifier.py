#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# pyright: reportUnknownMemberType=false, reportUnknownVariableType=false, reportUnknownArgumentType=false, reportUnknownParameterType=false, reportMissingParameterType=false, reportPrivateUsage=false, reportUnknownLambdaType=false, reportGeneralTypeIssues=false

import os
import re
import sys
import json
import zipfile
from lxml import etree
try:
    from pptx import Presentation  # 仅用于"能否作为演示文稿正常打开"的门槛校验
except ImportError:
    Presentation = None

# 说明：按接口统一约定，本模块不修改全局 sys.stdout，也不 print 主结果。
# 主结果统一由 evaluate(dir_path) -> dict 返回，入参为"脚本所在目录的路径"，
# 由脚本自己在该目录内定位并打开被评估的 .pptx。本文件仅在 __main__ 下作为
# 本地调试入口，把该 dict 序列化为 JSON 打印。

SCRIPT_ID = "069"
# 维度二加分项满分，也是最终完成度的分母；维度一未通过 / 脚本异常时回落
MAX_SCORE_TOTAL = 53

# 维度二加分项的 max_delta，顺序与 _evaluate_impl 中 results.append 的顺序严格一致。
# 分档口径：按"需要动手改的工作量"定 1 / 3 / 5——
#   宽度变化 ≥2 厘米（文字必然重新折行、框体必须重新安置）→ 5
#   一组颜色（填充 + 边框色 + 边框粗细）或 0.2~2 厘米的尺寸调整 → 3
#   单个字号 / 加粗 / 对齐开关，或 <0.2 厘米的微调 → 1
_PLUS_DELTAS_ORDER = [
    3,   # 01 2023年框 淡蓝填充 深蓝边框1.3磅 且 5.08x2.11厘米 13.2磅加粗 左22.68上5.94
    3,   # 02 2025年框 淡蓝填充 深蓝边框1.3磅 且 5.08x2.11厘米 13.2磅加粗 左39.32上5.94
    3,   # 03 林远航65%框 4.45x2.08厘米圆角矩形 淡蓝填充深蓝边框1磅 面板内不重叠
    1,   # 04 林远航65%框 12.3磅加粗
    3,   # 05 周予安35%框 4.45x2.08厘米圆角矩形 面板内不重叠
    1,   # 06 周予安35%框 12.3磅加粗
    3,   # 07 周予安35%框 淡蓝填充 深蓝边框1磅
    5,   # 08 董事会决策层框 4.45x2.08厘米圆角矩形 面板内不重叠（宽度缩 2.51 厘米）
    1,   # 09 董事会决策层框 12.3磅加粗
    3,   # 10 董事会决策层框 淡蓝填充 深蓝边框1磅
    3,   # 11 周予安30%框 淡蓝填充深蓝边框1磅 且 4.45x2.08厘米 12.3磅加粗 左27上11.51
    5,   # 12 仓配试点小组框 4.45x2.08厘米圆角矩形 面板内不重叠（宽度缩 5.2 厘米）
    1,   # 13 仓配试点小组框 12磅加粗
    3,   # 14 仓配试点小组框 淡粉填充 深玫红边框1磅
    5,   # 15 日常运营链路框 4.45x2.08厘米圆角矩形 面板内不重叠（宽度缩 2.54 厘米）
    1,   # 16 日常运营链路框 12磅加粗
    3,   # 17 日常运营链路框 淡粉填充 深玫红边框1磅
    1,   # 18 经营执行线框 4.45x2.08厘米圆角矩形 12磅加粗 面板内不重叠
    3,   # 19 经营执行线框 淡粉填充 深玫红边框1磅
    1,   # 20 监督复核线框 4.45x2.08厘米圆角矩形 12磅加粗 淡粉填充深玫红边框1磅
    1,   # 21 日常运营链路框 文字水平居中且垂直居中
]

# 维度二扣分项的 max_delta（均为负值），顺序与 _evaluate_impl 中追加顺序严格一致。
# 分档口径：只看"被改动的属性数量与还原成本"，不判断改动方向好不好看——
#   整框尺寸/字号/颜色被改，或全页级（字体、文字内容、形状增删、母版主色）损伤 → -5
#   多处轻量属性（对齐、文字颜色、装饰件位置）→ -3
#   开关型设置（动画/切换/超链接/备注）→ -1
_MINUS_DELTAS_ORDER = [
    -5,  # 22 红色箭头指向的2021年基准框 未被改动
    -5,  # 23 绿色箭头指向的林远航55%基准框 未被改动
    -5,  # 24 粉色箭头指向的许知澜15%基准框 未被改动
    -5,  # 25 初始职责说明框 未被改动
    -5,  # 26 调整后职责说明框 未被改动
    -5,  # 27 闭环机制说明框 未被改动
    -5,  # 28 三个澜江智仓主体框 未被改动
    -5,  # 29 三块面板底板与顶部色条 未被改动
    -5,  # 30 三组阶段标题与副标题 未被改动
    -3,  # 31 主标题/副标题/右上标识 未被改动
    -3,  # 32 底部分隔线/页脚/页码 未被改动
    -3,  # 33 时间轴小方块与横线 未被改动
    -3,  # 34 面板内连接线 未被改动
    -5,  # 35 字体名称 未被改动
    -3,  # 36 文字颜色 未被改动
    -3,  # 37 标注框以外的对齐方式 未被改动
    -5,  # 38 文字内容 未被改动
    -5,  # 39 形状总数为64 无增删
    -1,  # 40 无动画/切换/超链接/备注
    -5,  # 41 版式母版主题主色 未被改动
    -5,  # 42 数据权限岗框 未被改动
]

_MAX_DELTAS_ORDER = _PLUS_DELTAS_ORDER + _MINUS_DELTAS_ORDER
assert sum(_PLUS_DELTAS_ORDER) == MAX_SCORE_TOTAL

# ---------------------------------------------------------------------------
# OOXML 命名空间与单位换算
# ---------------------------------------------------------------------------
_A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
_P = 'http://schemas.openxmlformats.org/presentationml/2006/main'
_NS = {'a': _A, 'p': _P}

# 与办公软件对齐：形状的位置/尺寸在"大小和位置"面板里以厘米显示，文件里存 EMU，
# 1 厘米 = 360000 EMU；字号在"字体"面板里以磅显示，文件里存百分之一磅；
# 边框粗细在"线条-宽度"里以磅显示，文件里存 EMU，1 磅 = 12700 EMU。
_EMU_PER_CM = 360000.0
_EMU_PER_PT = 12700.0

# 判定容差：长度 ±0.05 厘米、位置 ±0.15 厘米、字号 ±0.1 磅、边框粗细 ±0.1 磅；
# 颜色按十六进制精确比对（办公软件的取色器给出的就是同一组值）。
_TOL_LEN = 0.05
_TOL_POS = 0.15
_TOL_SZ = 0.1
_TOL_LW = 0.1

# 颜色（与"填充-其他颜色-自定义"里的十六进制一致）
_BLUE_L, _BLUE_D = 'DDE9FF', '3768B3'      # 淡蓝填充 / 深蓝边框
_PINK_L, _PINK_D = 'F8E3E8', 'AD4E66'      # 淡粉填充 / 深玫红边框
_GRAY_L, _GRAY_M = 'F5F7FA', 'C2CAD6'      # 说明框浅灰填充 / 中灰边框
_DARK = '1D2733'                           # 正文深灰蓝
_WHITE = 'FFFFFF'
_PANEL_LINE = 'B8C2D0'                     # 面板底板边框
_BAR_COLORS = [_BLUE_D, '168B8B', 'B87814']  # 三块面板顶部色条
_PURPLE_L, _PURPLE_D = 'E9E4FF', '6B57B7'  # 数据权限岗框

# 三块阶段面板的左右边界与上下边界（厘米），标注框须整体落在面板内并留 0.3 厘米空白
_PANELS = [(1.78, 16.26), (18.16, 32.64), (34.54, 49.02)]
_PANEL_Y = (8.51, 26.04)
_CLEAR = 0.3

# 页面尺寸门槛（厘米）
_PAGE_W, _PAGE_H = 50.8, 28.575

# ---------------------------------------------------------------------------
# 基础工具
# ---------------------------------------------------------------------------
def _cm(value):
    """EMU 字符串 -> 厘米（保留 4 位，避免浮点噪声影响比对）。"""
    return round(int(value) / _EMU_PER_CM, 4)


def _near(a, b, tol=_TOL_LEN):
    """数值在容差内相等；a 为 None（属性缺失）时一律判否。"""
    return a is not None and abs(a - b) <= tol


def _pos(a, b):
    """位置比对，使用较宽的位置容差。"""
    return _near(a, b, _TOL_POS)


class _ShapeInfo(object):
    """把一个顶层 p:sp / p:cxnSp 摊平成扁平属性，便于逐条规则直读。"""

    def __init__(self, el):
        nv = el.find('.//p:cNvPr', _NS)
        self.id = int(nv.get('id')) if (nv is not None and nv.get('id')) else -1
        self.name = nv.get('name') if nv is not None else ''
        # ① 位置与尺寸：a:xfrm/a:off 是左上角，a:xfrm/a:ext 是宽高
        self.x = self.y = self.w = self.h = None
        xf = el.find('.//a:xfrm', _NS)
        if xf is not None:
            off = xf.find('a:off', _NS)
            ext = xf.find('a:ext', _NS)
            if off is not None:
                self.x, self.y = _cm(off.get('x')), _cm(off.get('y'))
            if ext is not None:
                self.w, self.h = _cm(ext.get('cx')), _cm(ext.get('cy'))
        # ② 文字：拼接形状内全部 a:t，供按内容定位形状使用
        self.text = ''.join(t.text or '' for t in el.iter('{%s}t' % _A))
        sp = el.find('p:spPr', _NS)
        # ③ 填充色：仅取纯色填充的十六进制值
        fill = sp.find('a:solidFill/a:srgbClr', _NS) if sp is not None else None
        self.fill = fill.get('val') if fill is not None else None
        # ④ 边框色与边框粗细
        ln = sp.find('a:ln', _NS) if sp is not None else None
        lnc = ln.find('a:solidFill/a:srgbClr', _NS) if ln is not None else None
        self.line = lnc.get('val') if lnc is not None else None
        self.lw = (round(int(ln.get('w')) / _EMU_PER_PT, 2)
                   if (ln is not None and ln.get('w')) else None)
        # ⑤ 形状几何：roundRect 即办公软件里的"圆角矩形"
        geom = sp.find('a:prstGeom', _NS) if sp is not None else None
        self.geom = geom.get('prst') if geom is not None else None
        self._collect_text_format(el)

    def _collect_text_format(self, el):
        """汇总形状内所有文字段的字号 / 加粗 / 字体 / 文字色 / 对齐 / 垂直对齐。

        一个形状可能有多段多行文字，这里做成去重后的有序集合：
        集合只有一个元素才说明"整框统一"，这正是本任务要求的判定粒度。
        """
        runs = [r for r in el.iter('{%s}rPr' % _A)]
        # 字号：sz 为百分之一磅，除 100 后即办公软件"字号"框里显示的数字
        self.sz = sorted({round(int(r.get('sz')) / 100.0, 2)
                          for r in runs if r.get('sz')})
        # 加粗：b="1" 即"加粗"按钮按下；缺省视为未加粗
        self.bold = sorted({r.get('b') for r in runs}, key=str)
        # 字体：a:latin@typeface 即"字体"下拉框里的名称
        self.fonts = sorted({(r.find('a:latin', _NS).get('typeface')
                              if r.find('a:latin', _NS) is not None else None)
                             for r in runs}, key=str)
        # 文字颜色
        self.tcol = sorted({(r.find('a:solidFill/a:srgbClr', _NS).get('val')
                             if r.find('a:solidFill/a:srgbClr', _NS) is not None else None)
                            for r in runs}, key=str)
        # 段落水平对齐：l/ctr/r 对应"左对齐/居中/右对齐"
        self.algn = sorted({p.get('algn') for p in el.iter('{%s}pPr' % _A)}, key=str)
        # 文本框垂直对齐：a:bodyPr@anchor="ctr" 即"垂直居中"
        bp = el.find('.//a:bodyPr', _NS)
        self.anchor = bp.get('anchor') if bp is not None else None

    def one_size(self, pt):
        """整框只有一个字号且等于 pt。"""
        return len(self.sz) == 1 and _near(self.sz[0], pt, _TOL_SZ)

    def is_bold(self):
        """整框文字全部加粗。"""
        return self.bold == ['1']

    def not_bold(self):
        """整框文字均未加粗。"""
        return '1' not in self.bold


def _open_presentation(file_path):
    """维度一门槛之一：能否被演示文稿程序正常打开。"""
    if Presentation is None:
        raise RuntimeError('缺少 python-pptx 依赖，无法完成可打开性校验')
    return Presentation(file_path)

def _load_package(file_path):
    """直读 .pptx 内部 XML，取回评分需要的全部原始信息。

    python-pptx 不便直接读到主题配色、切换动画与备注原文，因此这里用 zipfile
    + lxml 直读；可打开性由 _open_presentation 单独把关，两者互不替代。
    """
    with zipfile.ZipFile(file_path) as zf:
        names = zf.namelist()
        if 'ppt/slides/slide1.xml' not in names:
            raise RuntimeError('演示文稿内未找到第 1 页幻灯片')
        raw = zf.read('ppt/slides/slide1.xml').decode('utf-8')
        root = etree.fromstring(raw.encode('utf-8'))
        tree = root.find('.//p:cSld/p:spTree', _NS)
        if tree is None:
            raise RuntimeError('第 1 页幻灯片缺少形状树')
        # 只统计顶层形状与连接符，不下钻组合内部
        shapes = [_ShapeInfo(c) for c in tree
                  if etree.QName(c).localname in ('sp', 'cxnSp')]
        pres = etree.fromstring(zf.read('ppt/presentation.xml'))
        sld_sz = pres.find('p:sldSz', _NS)
        page_w = _cm(sld_sz.get('cx')) if sld_sz is not None else None
        page_h = _cm(sld_sz.get('cy')) if sld_sz is not None else None
        # 主题主色：设计-配色方案里的"着色 1"
        accent1 = None
        if 'ppt/theme/theme1.xml' in names:
            th = etree.fromstring(zf.read('ppt/theme/theme1.xml'))
            ac = th.find('.//a:clrScheme/a:accent1/a:srgbClr', _NS)
            accent1 = ac.get('val') if ac is not None else None
        # 备注：只关心是否写入了可见文字
        notes = ''
        for n in names:
            if n.startswith('ppt/notesSlides/notesSlide'):
                nr = etree.fromstring(zf.read(n))
                notes += ''.join(t.text or '' for t in nr.iter('{%s}t' % _A))
        n_slides = len([n for n in names
                        if re.match(r'ppt/slides/slide\d+\.xml$', n)])
    return {
        'shapes': shapes, 'raw': raw, 'page_w': page_w, 'page_h': page_h,
        'accent1': accent1, 'notes': notes, 'n_slides': n_slides,
    }


def _find_shape(shapes, *keys):
    """按文字关键字唯一定位形状；命中 0 个或多个都返回 None（判否而非猜测）。"""
    hits = [s for s in shapes if all(k in s.text.replace(' ', '') for k in keys)]
    return hits[0] if len(hits) == 1 else None

# ---------------------------------------------------------------------------
# 统一入口
# ---------------------------------------------------------------------------
def evaluate(dir_path):
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
            raise FileNotFoundError(f"目录不存在：{dir_path}")
        cands = sorted(f for f in os.listdir(dir_path)
                       if f.lower().endswith('.pptx') and not f.startswith('~$'))
        if not cands:
            raise FileNotFoundError(f"目录内未找到 .pptx 文件：{dir_path}")
        file_path = os.path.join(dir_path, cands[0])
        result["file_name"] = os.path.basename(file_path)
        return _evaluate_impl(file_path, result)
    except Exception as exc:
        # status 为 error 时未能完成判定，dim1_pass 留空，dim1_reason 也为空
        result["status"] = "error"
        result["error"] = f"{type(exc).__name__}: {exc}"
        result["dim1_pass"] = ""
        result["dim1_reason"] = ""
        result["dim2_items"] = []
        result["total_score"] = 0
        return result

def _evaluate_impl(file_path, result):
    # ========== 维度1：可用性与可修改性 ==========
    # 三条门槛：能被演示文稿程序正常打开、仅 1 页、页面尺寸仍是 50.8 x 28.575 厘米。
    # 任一不满足即维度一不通过，维度二不再计分（完成度按 0 记）。
    _open_presentation(file_path)
    pkg = _load_package(file_path)
    gate_msgs = []
    if pkg['n_slides'] != 1:
        gate_msgs.append('页数为 %d，应为 1' % pkg['n_slides'])
    if not _near(pkg['page_w'], _PAGE_W, 0.1) or not _near(pkg['page_h'], _PAGE_H, 0.1):
        gate_msgs.append('页面尺寸为 %s x %s 厘米，应为 50.8 x 28.575 厘米'
                         % (pkg['page_w'], pkg['page_h']))
    if gate_msgs:
        result["dim1_pass"] = False
        result["dim1_reason"] = '；'.join(gate_msgs)
        result["dim2_items"] = []
        result["total_score"] = 0
        result["max_score"] = MAX_SCORE_TOTAL
        return result
    # dim1_reason 只记录"维度一不通过的原因"，通过时保持为空字符串。
    result["dim1_pass"] = True
    result["dim1_reason"] = ""

    # ========== 维度2：完成度评分 ==========
    shapes = pkg['shapes']
    results = []
    details = []          # 判定过程中的诊断信息，仅用于调试，不进入返回结果
    item_details = []     # 与 results 一一对应
    _cursor = [0]

    def g(*keys):
        return _find_shape(shapes, *keys)

    # 三个标注框带内涉及的框体（按图上矩形边界判定范围）：
    # 红色带 3 个年份框、绿色带 5 个人名/机构框、粉色带 5 个框。
    # 数据权限岗框主体落在粉色矩形之外，不属于标注范围，改到扣分项把关。
    y21, y23, y25 = g('2021年'), g('2023年'), g('2025年')
    l65, z35 = g('持股65%'), g('持股35%')
    l55, z30 = g('持股55%'), g('持股30%')
    dsh = g('董事会决策层')
    cang, rich = g('仓配试点小组'), g('日常运营链路')
    jing, jian = g('经营执行线'), g('监督复核线')
    shuj, x15 = g('数据权限岗'), g('持股15%')
    band = [s for s in (y21, y23, y25, l65, z35, l55, z30, dsh,
                        cang, rich, jing, jian, shuj, x15)
            if isinstance(s, _ShapeInfo)]

    def ok(s):
        return isinstance(s, _ShapeInfo) and s.x is not None and s.w is not None

    def inside(s):
        """框体整体落在所属阶段面板内，四边各留至少 0.3 厘米空白。"""
        if not ok(s):
            return False
        cx = s.x + s.w / 2
        hit = [p for p in _PANELS if p[0] <= cx <= p[1]]
        if not hit:
            details.append('%s：不在任何阶段面板内' % s.text[:8])
            return False
        left, right = hit[0]
        bad = []
        if s.x < left + _CLEAR:
            bad.append('左 %.2f 越界 %.2f' % (s.x, left + _CLEAR))
        if s.x + s.w > right - _CLEAR:
            bad.append('右 %.2f 越界 %.2f' % (s.x + s.w, right - _CLEAR))
        if s.y < _PANEL_Y[0] + _CLEAR:
            bad.append('上 %.2f 越界 %.2f' % (s.y, _PANEL_Y[0] + _CLEAR))
        if s.y + s.h > _PANEL_Y[1] - _CLEAR:
            bad.append('下 %.2f 越界 %.2f' % (s.y + s.h, _PANEL_Y[1] - _CLEAR))
        if bad:
            details.append('%s：%s' % (s.text[:8], '/'.join(bad)))
            return False
        return True

    def free(s):
        """框体与其余人名/机构框互不重叠。"""
        if not ok(s):
            return False
        for o in band:
            if o is s or not ok(o):
                continue
            if (s.x < o.x + o.w and o.x < s.x + s.w
                    and s.y < o.y + o.h and o.y < s.y + s.h):
                details.append('%s：与「%s」重叠' % (s.text[:8], o.text[:6]))
                return False
        return True

    def geo(s, w, h):
        """尺寸为 w x h 厘米且形状为圆角矩形。"""
        return ok(s) and _near(s.w, w) and _near(s.h, h) and s.geom == 'roundRect'

    def col(s, fill, line, lw):
        """填充色、边框色、边框粗细三项同时符合。"""
        return (ok(s) and s.fill == fill and s.line == line
                and _near(s.lw, lw, _TOL_LW))

    def fnt(s, pt):
        """整框统一字号 pt 且全部加粗。"""
        return ok(s) and s.one_size(pt) and s.is_bold()

    def at(s, x, y):
        """左上角位置符合。"""
        return ok(s) and _pos(s.x, x) and _pos(s.y, y)

    def box(s, x, y, w, h):
        """位置与尺寸四项同时符合（不含颜色字号）。"""
        return (ok(s) and _pos(s.x, x) and _pos(s.y, y)
                and _near(s.w, w) and _near(s.h, h))

    def spec(s, x, y, w, h, pt, fill, line, lw, bold=False):
        """整框"原样未动"的完整指纹：位置 + 尺寸 + 字号 + 加粗 + 填充 + 边框 + 粗细。"""
        return (box(s, x, y, w, h) and s.one_size(pt)
                and (s.is_bold() if bold else s.not_bold())
                and s.fill == fill and s.line == line and _near(s.lw, lw, _TOL_LW))

    def add(delta, hit, rule):
        """按模板约定，把每条规则的结果压成 "±N：描述" 字符串。

        同时把本条规则判定过程中产生的诊断信息（几何越界、重叠、缺失文字等）
        切给该条目，仅留作调试用，不写进返回结果。
        """
        sign = '+' if delta >= 0 else '-'
        results.append('%s%d：%s' % (sign, abs(delta) if hit else 0, rule))
        item_details.append(' | '.join(details[_cursor[0]:]))
        _cursor[0] = len(details)

    # ---- 加分项：客户明确要求改动的三组标注框 ----
    # 01 2023年框：与2021年基准框统一为淡蓝填充 + 深蓝边框 1.3 磅，
    #    同时保持 5.08x2.11 厘米、13.2 磅加粗、左 22.68 上 5.94 厘米不被顺手挪动。
    add(3, (col(y23, _BLUE_L, _BLUE_D, 1.3) and geo(y23, 5.08, 2.11)
            and fnt(y23, 13.2) and at(y23, 22.68, 5.94)),
        '2023年框为淡蓝填充深蓝边框1.3磅，且5.08x2.11厘米、13.2磅加粗、左22.68上5.94厘米')
    # 02 2025年框：同上，位置为左 39.32 上 5.94 厘米。
    add(3, (col(y25, _BLUE_L, _BLUE_D, 1.3) and geo(y25, 5.08, 2.11)
            and fnt(y25, 13.2) and at(y25, 39.32, 5.94)),
        '2025年框为淡蓝填充深蓝边框1.3磅，且5.08x2.11厘米、13.2磅加粗、左39.32上5.94厘米')
    # 03 林远航65%框：缩到 4.45x2.08 厘米圆角矩形、淡蓝填充深蓝边框 1 磅，
    #    且缩放后仍整体在阶段面板内、与相邻框不重叠。
    add(3, (geo(l65, 4.45, 2.08) and col(l65, _BLUE_L, _BLUE_D, 1)
            and inside(l65) and free(l65)),
        '林远航65%框为4.45x2.08厘米圆角矩形、淡蓝填充深蓝边框1磅，且在面板内不与其他框重叠')
    # 04 林远航65%框字号：整框 12.3 磅加粗。
    add(1, fnt(l65, 12.3), '林远航65%框整框12.3磅加粗')
    # 05 周予安35%框：4.45x2.08 厘米圆角矩形，面板内不重叠。
    add(3, (geo(z35, 4.45, 2.08) and inside(z35) and free(z35)),
        '周予安35%框为4.45x2.08厘米圆角矩形，且在面板内不与其他框重叠')
    # 06 周予安35%框字号：整框 12.3 磅加粗。
    add(1, fnt(z35, 12.3), '周予安35%框整框12.3磅加粗')
    # 07 周予安35%框配色：淡蓝填充 + 深蓝边框 1 磅。
    add(3, col(z35, _BLUE_L, _BLUE_D, 1), '周予安35%框为淡蓝填充、深蓝边框1磅')

    # 08 董事会决策层框：原宽 6.96 厘米，缩到 4.45x2.08 厘米圆角矩形，
    #    宽度变化 2.51 厘米，文字必然重新折行且框体要重新安置，故计 5 分。
    add(5, (geo(dsh, 4.45, 2.08) and inside(dsh) and free(dsh)),
        '董事会决策层框为4.45x2.08厘米圆角矩形，且在面板内不与其他框重叠')
    # 09 董事会决策层框字号：整框 12.3 磅加粗。
    add(1, fnt(dsh, 12.3), '董事会决策层框整框12.3磅加粗')
    # 10 董事会决策层框配色：淡蓝填充 + 深蓝边框 1 磅。
    add(3, col(dsh, _BLUE_L, _BLUE_D, 1), '董事会决策层框为淡蓝填充、深蓝边框1磅')
    # 11 周予安30%框：只需统一配色，尺寸字号位置本来就与基准一致，
    #    这里把"改到位"和"没被顺手挪动"合为一条，避免与扣分项重叠。
    add(3, (col(z30, _BLUE_L, _BLUE_D, 1) and geo(z30, 4.45, 2.08)
            and fnt(z30, 12.3) and at(z30, 27.00, 11.51)),
        '周予安30%框为淡蓝填充深蓝边框1磅，且4.45x2.08厘米、12.3磅加粗、左27上11.51厘米')
    # 12 仓配试点小组框：原宽 9.65 厘米，缩到 4.45x2.08 厘米，宽度变化 5.2 厘米。
    add(5, (geo(cang, 4.45, 2.08) and inside(cang) and free(cang)),
        '仓配试点小组框为4.45x2.08厘米圆角矩形，且在面板内不与其他框重叠')
    # 13 仓配试点小组框字号：整框 12 磅加粗。
    add(1, fnt(cang, 12), '仓配试点小组框整框12磅加粗')
    # 14 仓配试点小组框配色：淡粉填充 + 深玫红边框 1 磅。
    add(3, col(cang, _PINK_L, _PINK_D, 1), '仓配试点小组框为淡粉填充、深玫红边框1磅')
    # 15 日常运营链路框：原宽 6.99 厘米，缩到 4.45x2.08 厘米，宽度变化 2.54 厘米。
    add(5, (geo(rich, 4.45, 2.08) and inside(rich) and free(rich)),
        '日常运营链路框为4.45x2.08厘米圆角矩形，且在面板内不与其他框重叠')
    # 16 日常运营链路框字号：整框 12 磅加粗。
    add(1, fnt(rich, 12), '日常运营链路框整框12磅加粗')
    # 17 日常运营链路框配色：淡粉填充 + 深玫红边框 1 磅。
    add(3, col(rich, _PINK_L, _PINK_D, 1), '日常运营链路框为淡粉填充、深玫红边框1磅')
    # 18 经营执行线框：宽度只差 0.12 厘米，尺寸与字号合为一条按 1 分计。
    add(1, (geo(jing, 4.45, 2.08) and fnt(jing, 12)
            and inside(jing) and free(jing)),
        '经营执行线框为4.45x2.08厘米圆角矩形、整框12磅加粗，且在面板内不与其他框重叠')
    # 19 经营执行线框配色：淡粉填充 + 深玫红边框 1 磅。
    add(3, col(jing, _PINK_L, _PINK_D, 1), '经营执行线框为淡粉填充、深玫红边框1磅')
    # 20 监督复核线框：尺寸微调 + 字号 + 配色（配色本已一致）合为一条按 1 分计。
    add(1, (geo(jian, 4.45, 2.08) and fnt(jian, 12)
            and col(jian, _PINK_L, _PINK_D, 1) and inside(jian) and free(jian)),
        '监督复核线框为4.45x2.08厘米圆角矩形、整框12磅加粗、淡粉填充深玫红边框1磅，且在面板内不与其他框重叠')
    # 21 日常运营链路框对齐：文字水平居中且垂直居中（原为左对齐上对齐）。
    add(1, (ok(rich) and rich.algn == ['ctr'] and rich.anchor == 'ctr'),
        '日常运营链路框文字水平居中且垂直居中')

    # ---- 扣分项：客户没要求改、但被顺手改坏的地方 ----
    # 22~24 三个基准框本身就是"参照物"，只要指纹有任何一处变化即扣分。
    add(-5, not spec(y21, 6.72, 5.94, 5.08, 2.11, 13.2, _BLUE_L, _BLUE_D, 1.3, True),
        '红色箭头指向的2021年基准框被改动（位置/尺寸/字号/加粗/填充/边框任一变化）')
    add(-5, not spec(l55, 19.30, 11.51, 4.45, 2.08, 12.3, _BLUE_L, _BLUE_D, 1, True),
        '绿色箭头指向的林远航55%基准框被改动')
    add(-5, not spec(x15, 19.30, 20.27, 4.45, 2.08, 12, _PINK_L, _PINK_D, 1, True),
        '粉色箭头指向的许知澜15%基准框被改动')
    # 25~27 三个说明框：浅灰填充 + 中灰边框 1 磅、非加粗，字号各不相同。
    chu, tiao, bi = g('初始职责'), g('调整后职责'), g('闭环机制')
    add(-5, not spec(chu, 3.61, 23.19, 11.33, 1.98, 11.2, _GRAY_L, _GRAY_M, 1),
        '初始职责说明框被改动')
    add(-5, not spec(tiao, 19.84, 22.86, 11.30, 2.11, 10.9, _GRAY_L, _GRAY_M, 1),
        '调整后职责说明框被改动')
    add(-5, not spec(bi, 36.20, 24.54, 11.18, 1.07, 10.7, _GRAY_L, _GRAY_M, 1),
        '闭环机制说明框被改动')
    # 28 三个"澜江智仓科技有限公司"主体框：白底 + 深灰蓝边框 1.4 磅、加粗，按左到右比对。
    lan = sorted([s for s in shapes if '澜江智仓科技有限公司' in s.text.replace(' ', '')],
                 key=lambda s: (s.x if s.x is not None else 0))
    lan_spec = [(5.77, 15.57, 6.99, 2.08, 12.8),
                (22.12, 15.42, 6.73, 1.91, 12.6),
                (38.28, 15.32, 7.01, 2.08, 12.5)]
    add(-5, not (len(lan) == 3
                 and all(spec(s, x, y, w, h, pt, _WHITE, _DARK, 1.4, True)
                         for s, (x, y, w, h, pt) in zip(lan, lan_spec))),
        '三个澜江智仓科技有限公司主体框被改动')
    # 29 三块阶段面板底板（14.48x17.53 厘米白底灰蓝边框 1.1 磅）与顶部色条
    #    （14.48x0.3 厘米，自左至右深蓝/深青/深金）。
    panels = sorted([s for s in shapes
                     if not s.text and _near(s.w, 14.48) and _near(s.h, 17.53)],
                    key=lambda s: s.x)
    bars = sorted([s for s in shapes
                   if not s.text and _near(s.w, 14.48) and _near(s.h, 0.30)],
                  key=lambda s: s.x)
    add(-5, not (len(panels) == 3 and len(bars) == 3
                 and all(s.fill == _WHITE and s.line == _PANEL_LINE
                         and _near(s.lw, 1.1, _TOL_LW)
                         and _pos(s.y, 8.51) and _pos(s.x, X)
                         for s, X in zip(panels, [1.78, 18.16, 34.54]))
                 and [s.fill for s in bars] == _BAR_COLORS),
        '三块阶段面板底板或顶部色条被改动')

    # 30 三组阶段标题（19 磅加粗深灰蓝）与阶段副标题（11.5 磅非加粗灰蓝）。
    stage_t = [g('阶段一'), g('阶段二'), g('阶段三')]
    stage_s = [g('股权出资'), g('新增财务风控'), g('经营层')]
    add(-5, not (all(ok(s) and s.one_size(19) and s.is_bold() and s.tcol == [_DARK]
                     and box(s, X, 9.02, 13.21, 0.97)
                     for s, X in zip(stage_t, [2.41, 18.80, 35.18]))
                 and all(ok(s) and s.one_size(11.5) and s.not_bold()
                         and s.tcol == ['778391'] and box(s, X, 10.08, 13.21, 0.66)
                         for s, X in zip(stage_s, [2.41, 18.80, 35.18]))),
        '三组阶段标题或阶段副标题被改动')
    # 31 主标题（28 磅加粗）、副标题（13.5 磅）、右上标识（10.5 磅）。
    t1 = g('智慧仓储运营主体治理演进')
    t2 = g('从共同设立到')
    t3 = g('CONFIDENTIAL')
    add(-3, not (ok(t1) and t1.one_size(28) and t1.is_bold() and t1.tcol == [_DARK]
                 and box(t1, 1.78, 1.07, 22.35, 1.40)
                 and ok(t2) and t2.one_size(13.5) and t2.tcol == ['5E6B7A']
                 and box(t2, 1.83, 2.67, 23.37, 0.81)
                 and ok(t3) and t3.one_size(10.5) and t3.tcol == ['A1ACB8']
                 and box(t3, 38.35, 1.32, 10.67, 0.71)),
        '主标题、副标题或右上标识被改动')
    # 32 底部分隔线（宽 47.24 厘米的水平线）、页脚文字（8.5 磅）、页码"01"（10.5 磅）。
    hline = [s for s in shapes
             if not s.text and _near(s.h, 0, 0.02) and _near(s.w, 47.24, _TOL_POS)]
    footer = g('资料用途')
    pageno = [s for s in shapes if s.text.strip() == '01']
    add(-3, not (len(hline) == 1 and _pos(hline[0].y, 27.30) and _pos(hline[0].x, 1.78)
                 and hline[0].line == 'D5DCE7'
                 and ok(footer) and footer.one_size(8.5) and footer.tcol == ['9AA6B5']
                 and box(footer, 1.91, 27.51, 29.21, 0.56)
                 and len(pageno) == 1 and pageno[0].one_size(10.5)
                 and pageno[0].tcol == ['9AA6B5']),
        '底部分隔线、页脚或页码被改动')
    # 33 时间轴：3 个 1.27x0.92 厘米小方块 + 2 条连接横线。
    squares = sorted([s for s in shapes
                      if not s.text and _near(s.w, 1.27) and _near(s.h, 0.92)],
                     key=lambda s: s.x)
    tlines = sorted([s for s in shapes if not s.text and _near(s.h, 0, 0.02)
                     and _pos(s.y, 4.83)], key=lambda s: s.x)
    add(-3, not (len(squares) == 3
                 and all(_pos(s.y, 4.37) and _pos(s.x, X)
                         for s, X in zip(squares, [8.76, 24.59, 41.15]))
                 and len(tlines) == 2
                 and _pos(tlines[0].x, 10.03) and _near(tlines[0].w, 14.55, _TOL_POS)
                 and _pos(tlines[1].x, 25.86) and _near(tlines[1].w, 15.29, _TOL_POS)),
        '时间轴小方块或横线被改动')

    # 34 面板内连接线：抽检 9 条竖线的横坐标与 3 条横线的横坐标 + 长度。
    #    连接线不随框体自动跟随，被误拖后图会散，属"惹麻烦"。
    v_x = [5.41, 12.52, 9.27, 21.54, 29.24, 25.48, 41.78, 37.97, 45.59]
    v_lines = [s for s in shapes if not s.text and _near(s.w, 0, 0.02)
               and s.y is not None and s.y > _PANEL_Y[0]]
    h_spec = [(5.41, 7.11), (21.54, 7.70), (37.97, 7.62)]
    h_lines = [s for s in shapes if not s.text and _near(s.h, 0, 0.02)
               and s.y is not None and s.y > _PANEL_Y[0]]
    add(-3, not (all(any(_pos(s.x, X) for s in v_lines) for X in v_x)
                 and all(any(_pos(s.x, X) and _near(s.w, W, _TOL_POS) for s in h_lines)
                         for X, W in h_spec)),
        '面板内连接线被改动')
    # 35 字体：全页文字统一使用 Noto Sans CJK SC，换字体会让全图字形错位。
    add(-5, any(f != 'Noto Sans CJK SC'
                for s in shapes for f in s.fonts if f is not None),
        '字体名称被改动（全页文字应统一为 Noto Sans CJK SC）')
    # 36 文字颜色：标注框带、主体框、说明框内文字应统一为深灰蓝。
    add(-3, any(ok(s) and s.tcol != [_DARK] for s in band + lan + [chu, tiao, bi]),
        '框体内文字颜色被改动')
    # 37 标注框以外的对齐方式：说明框与阶段标题、主副标题左对齐；
    #    闭环机制与三个主体框居中；右上标识与页码右对齐。
    lefts = [chu, tiao] + stage_t + [t1, t2]
    ctrs = [bi] + lan
    rights = [t3] + pageno
    add(-3, not (all(ok(s) and s.algn == ['l'] for s in lefts)
                 and all(ok(s) and s.algn == ['ctr'] for s in ctrs)
                 and all(ok(s) and s.algn == ['r'] for s in rights)),
        '标注框以外的对齐方式被改动')
    # 38 文字内容：抽检 24 个关键词，缺任何一个说明文案被删改。
    need = ['林远航', '周予安', '许知澜', '陈屿舟', '65%', '35%', '55%', '30%', '15%',
            '发起投资', '产业资源', '董事长', '总经理', '财务风控', '监事', '执行董事',
            '渠道协同', '2021', '2023', '2025', '澜江智仓科技有限公司',
            '智慧仓储运营主体治理演进', 'CONFIDENTIAL STRUCTURE MAP', '资料用途']
    all_text = ''.join(s.text for s in shapes).replace(' ', '')
    missing = [n for n in need if n.replace(' ', '') not in all_text]
    if missing:
        details.append('缺失文字：%s' % '、'.join(missing))
    add(-5, bool(missing), '文字内容被改动或删除')
    # 39 形状总数：原图 64 个顶层形状，多删少增都说明动了结构。
    if len(shapes) != 64:
        details.append('形状总数为 %d，应为 64' % len(shapes))
    add(-5, len(shapes) != 64, '形状总数不再是64（有增删）')

    # 40 开关型设置：不应新增切换效果、动画、超链接或备注文字。
    add(-1, (('<p:transition' in pkg['raw']) or ('hlinkClick' in pkg['raw'])
             or ('<a:anim' in pkg['raw']) or bool(pkg['notes'].strip())),
        '新增了动画、切换、超链接或备注')
    # 41 主题主色：设计-配色方案里的"着色 1"应仍为标准蓝 4472C4。
    add(-5, pkg['accent1'] != '4472C4', '版式母版的主题主色被改动')
    # 42 数据权限岗框：主体落在粉色标注矩形之外，客户没要求改，
    #    淡紫填充 + 深紫边框 1 磅、6.6x2.08 厘米、11.7 磅加粗应保持原样。
    add(-5, not spec(shuj, 38.48, 22.10, 6.60, 2.08, 11.7, _PURPLE_L, _PURPLE_D, 1, True),
        '数据权限岗框被改动')

    # ---- 汇总 ----
    if len(results) != len(_MAX_DELTAS_ORDER):
        raise RuntimeError(f"dim2 项数不一致：results={len(results)} vs order={len(_MAX_DELTAS_ORDER)}")
    dim2_items = []
    total_score = 0
    for raw, max_delta in zip(results, _MAX_DELTAS_ORDER):
        head, _, rule = raw.partition("：")
        m = re.match(r'^([+-])(\d+)$', head.strip())
        try:
            delta = int(m.group(2)) * (-1 if m.group(1) == '-' else 1)
        except (AttributeError, ValueError):
            delta = 0
        hit = delta != 0
        dim2_items.append({
            "rule": rule,
            "max_delta": max_delta,
            "delta": max_delta if hit else 0,
            "hit": hit,
        })
        total_score += (max_delta if hit else 0)
    # 加分与扣分共同决定最终得分，实事求是输出：合计为负时保留负值，不做截断
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

