#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# pyright: reportUnknownMemberType=false, reportUnknownVariableType=false, reportUnknownArgumentType=false, reportMissingImports=false, reportPrivateUsage=false, reportOptionalMemberAccess=false, reportGeneralTypeIssues=false

import os, re, sys, json, math, traceback
from pptx import Presentation
from pptx.util import Emu

# 说明：按接口统一约定，本模块不修改全局 sys.stdout，也不 print 主结果。
# 主结果统一由 evaluate(dir_path) -> dict 返回；入参为“脚本所在目录的路径”，
# 由脚本自己在该目录内定位并打开被评估的 .pptx。本文件仅在 __main__ 下作为
# 本地调试入口，才把 dict 序列化为 JSON 打印。

SCRIPT_ID = "079"
# 维度二所有加分项 max_delta 之和；用于维度一未通过 / 脚本异常时兜底
MAX_SCORE_TOTAL = 65
# 维度二所有减分项 max_delta 之和（负值），仅用于自检
MIN_SCORE_TOTAL = -44

# 维度二所有评分项按“结果追加顺序”的 max_delta 列表；末端把 results 里的
# "+N：..." / "-N：..." 字符串反解为结构化 dim2_items 时用。顺序须与
# _evaluate_impl 中 results.append(...) 的顺序严格一致；共 33 项，
# 其中 17 项加分累加 = MAX_SCORE_TOTAL，16 项减分累加 = MIN_SCORE_TOTAL。
_MAX_DELTAS_ORDER = [
    3,   # 01 13页幻灯片与页码02—13一一对应
    5,   # 02 13页主标题全部译为蒙古文
    3,   # 03 13页章节标签译出且序号01—12保留
    3,   # 04 13页副标题说明行译为蒙古文
    3,   # 05 13页底部两行页脚译为蒙古文
    5,   # 06 第5—12页并列说明卡片约30个文本块译出
    3,   # 07 第1页五个圆角标签译出且互不重复
    3,   # 08 三处箭头流程行译出并保留九个“→”
    5,   # 09 第3页碳数—名称对照表表头与十个烷烃名
    3,   # 10 第3页三个小标题及说明段、“十二烷”
    5,   # 11 第4页三个碳架卡片与三段说明、三个戊烷命名
    3,   # 12 第2页四个烷基名称
    5,   # 13 全篇化学式与元素符号原样保留
    5,   # 14 第12页五个同分异构体名称与位号
    3,   # 15 位号与编号数字原样保留
    3,   # 16 第13页题目A、题目B全文译出且数量关系保留
    5,   # 17 主标题与副标题译文完整可见、未折行溢出、未压盖
    -5,  # 18 副标题说明行字体格式 微软雅黑 12.5磅 #6B7C84
    -3,  # 19 第3、6、11页主标题 微软雅黑 26磅 加粗 #0B5965
    -3,  # 20 页脚 微软雅黑 9.5磅 #6B7C84
    -5,  # 21 各页卡片正文字号分级（≥10处或缩放磅值扣5，个别微调扣1）
    -3,  # 22 第4页三个碳架说明块末尾不追加空段
    -3,  # 23 十处正文末尾不追加空段
    -3,  # 24 第4页三段碳架说明 微软雅黑 #22313F
    -3,  # 25 第5页强调行整行 微软雅黑 14磅 加粗 #0B5965
    -3,  # 26 中文标点残留（句号／全角空格／顿号）
    -3,  # 27 13页主标题文本框 左1.4 顶2.08 高1.22
    -3,  # 28 13页副标题说明行 左1.47 顶3.45 宽24.89 高0.71
    -3,  # 29 各页并列说明卡片定位与块高
    -1,  # 30 各页正文分隔线坐标
    -1,  # 31 第1页五个圆角标签坐标与文字格式
    -1,  # 32 第3页碳数对照表十数字块与十名称块字号坐标
    -1,  # 33 四张示意图片与页面装饰元素坐标
]
assert sum(d for d in _MAX_DELTAS_ORDER if d > 0) == MAX_SCORE_TOTAL
assert sum(d for d in _MAX_DELTAS_ORDER if d < 0) == MIN_SCORE_TOTAL

# ========== 目标版式常量（全部为本套课件应有的目标值，不依赖任何外部文件） ==========
_A = '{http://schemas.openxmlformats.org/drawingml/2006/main}'
_PT_PER_CM = 28.3465          # 1cm = 28.3465pt
_GEO_TOL = 0.06               # 坐标/尺寸判定容差（cm），约等于 1/2 磅
_OVERLAP_TOL = 0.25           # 压盖判定容差（cm），覆盖字宽估算误差
_SLIDE_W, _SLIDE_H = 33.86, 19.05

_MSYH = ('微软雅黑', 'Microsoft YaHei')
_TEAL, _GRAY, _ORANGE, _INK, _WHITE = '0B5965', '6B7C84', 'E69A3A', '22313F', 'FFFFFF'

# 主标题 / 副标题说明行 / 章节标签 / 页码标签：每页形状名一致
_TITLE, _SUB, _KICKER, _PAGENO = 'Text 4', 'Text 5', 'Text 3', 'Text 6'
# 13页底部两行页脚（第1页仅一行）
_FOOTERS = {
    1: ['Text 18'], 2: ['Text 26', 'Text 27'], 3: ['Text 61', 'Text 62'],
    4: ['Text 44', 'Text 45'], 5: ['Text 18', 'Text 19'], 6: ['Text 34', 'Text 35'],
    7: ['Text 15', 'Text 16'], 8: ['Text 17', 'Text 18'], 9: ['Text 15', 'Text 16'],
    10: ['Text 23', 'Text 24'], 11: ['Text 17', 'Text 18'], 12: ['Text 14', 'Text 15'],
    13: ['Text 18', 'Text 19'],
}
# 第1页五个圆角标签（自左至右）
_CHIPS = ['Text 7', 'Text 9', 'Text 11', 'Text 13', 'Text 15']
_CHIP_LEFTS = [1.70, 4.78, 7.85, 10.92, 13.59]
# 第5—12页并列说明卡片（小标题＋正文段落）
_CARD_BLOCKS = {
    5: ['Text 8', 'Text 9', 'Text 11', 'Text 12', 'Text 14', 'Text 15'],
    6: ['Text 8', 'Text 9'],
    7: ['Text 8', 'Text 9', 'Text 11', 'Text 12'],
    8: ['Text 8', 'Text 9', 'Text 11', 'Text 12', 'Text 14', 'Text 15'],
    9: ['Text 8', 'Text 9', 'Text 11', 'Text 12'],
    10: ['Text 8', 'Text 9'],
    11: ['Text 8', 'Text 9', 'Text 11', 'Text 12', 'Text 14', 'Text 15'],
    12: ['Text 8', 'Text 9', 'Text 11', 'Text 12'],
}
# 各页正文/说明行应有字号（磅）——细则“卡片正文与说明行字号分级”条目的目标值
_BODY_PT = {
    (1, 'Text 5'): 16.0, (1, 'Text 16'): 14.0,
    (2, 'Text 9'): 14.0, (2, 'Text 12'): 14.0,
    (3, 'Text 53'): 14.0, (3, 'Text 56'): 14.0, (3, 'Text 59'): 13.5,
    (4, 'Text 9'): 13.5, (4, 'Text 21'): 13.5, (4, 'Text 33'): 13.5,
    (5, 'Text 9'): 17.0, (5, 'Text 12'): 17.0, (5, 'Text 15'): 16.0,
    (6, 'Text 9'): 16.0,
    (7, 'Text 9'): 17.0, (7, 'Text 12'): 16.0,
    (8, 'Text 9'): 19.0, (8, 'Text 12'): 19.0, (8, 'Text 15'): 20.0,
    (9, 'Text 12'): 18.0,
    (10, 'Text 9'): 18.0,
    (11, 'Text 9'): 16.0, (11, 'Text 12'): 16.0, (11, 'Text 15'): 15.5,
    (12, 'Text 9'): 21.0, (12, 'Text 12'): 21.0,
    (13, 'Text 9'): 15.5, (13, 'Text 12'): 15.5, (13, 'Text 15'): 17.0,
    (13, 'Text 16'): 16.0,
}
# 按 0.86 / 0.92 倍等比缩放后出现的非整数磅值（成批改动的指纹）
_SCALED_PT = (12.04, 12.42, 13.33, 13.76, 14.62, 14.72, 15.48, 15.64,
              17.48, 18.06, 18.40, 19.32)

# 第3页碳数—名称对照表：十个数字块与十个名称块
_TAB_NUM = ['Text 10', 'Text 12', 'Text 14', 'Text 16', 'Text 18',
            'Text 20', 'Text 22', 'Text 24', 'Text 26', 'Text 28']
_TAB_NAME = ['Text 32', 'Text 34', 'Text 36', 'Text 38', 'Text 40',
             'Text 42', 'Text 44', 'Text 46', 'Text 48', 'Text 50']
_TAB_DIGITS = ['1', '2', '3', '4', '5', '6', '7', '8', '9', '10']
_ALKANES = ['метан', 'этан', 'пропан', 'бутан', 'пентан',
            'гексан', 'гептан', 'октан', 'нонан', 'декан']
_TAB_NUM_LEFT = [4.63, 7.35, 10.08, 12.80, 15.53, 18.25, 20.98, 23.70, 26.43, 29.15]
# 第6页主链编号七个数字标签
_CHAIN_NUM = ['Text 11', 'Text 14', 'Text 17', 'Text 20', 'Text 23', 'Text 26', 'Text 29']
# 全篇化学式与元素符号：形状 -> 应原样保留的文字
_FORMULA_EXACT = {
    (2, 'Text 15'): '—CH3', (2, 'Text 18'): '—CH2CH3',
    (2, 'Text 21'): '—CH2CH2CH3', (2, 'Text 24'): '—CH(CH3)2',
    (4, 'Text 10'): 'CH3', (4, 'Text 12'): 'CH2', (4, 'Text 14'): 'CH2',
    (4, 'Text 16'): 'CH2', (4, 'Text 18'): 'CH3', (4, 'Text 22'): 'CH3',
    (4, 'Text 24'): 'CH', (4, 'Text 26'): 'CH2', (4, 'Text 28'): 'CH3',
    (4, 'Text 29'): 'CH3', (4, 'Text 34'): 'C', (4, 'Text 39'): 'CH3',
    (4, 'Text 40'): 'CH3', (4, 'Text 41'): 'CH3', (4, 'Text 42'): 'CH3',
    (6, 'Text 31'): 'CH3',
    (10, 'Text 10'): 'C', (10, 'Text 12'): 'C', (10, 'Text 14'): 'C',
    (10, 'Text 16'): 'C', (10, 'Text 18'): 'C', (10, 'Text 20'): 'C',
}
_FORMULA_SUBSTR = {
    (2, 'Text 9'): ['—R'], (2, 'Text 12'): ['—CnH2n+1'],
    (4, 'Text 5'): ['C5H12'], (10, 'Text 8'): ['C6H14'], (12, 'Text 11'): ['C6H14'],
}
# 第12页五个同分异构体的蒙古文命名（含位号与逗号）
_ISOMERS = ['n-гексан', '2-метилпентан', '3-метилпентан',
            '2,2-диметилбутан', '2,3-диметилбутан']
# 三处箭头流程行及各自应保留的“→”数量（合计九个）
_ARROW_ROWS = [(1, 'Text 16', 3), (8, 'Text 15', 2), (13, 'Text 16', 4)]
# 正文末尾不得追加空段的十处文本块
_NO_TAIL_EMPTY = [(1, 'Text 4'), (3, 'Text 53'), (4, 'Text 44'), (6, 'Text 9'),
                  (7, 'Text 9'), (7, 'Text 12'), (8, 'Text 15'), (9, 'Text 12'),
                  (10, 'Text 9'), (12, 'Text 12')]
# 各页并列说明卡片应有的定位与块高
_CARD_GEO = [
    (6, 'Text 9', 'left', 2.36), (6, 'Text 9', 'top', 5.77),
    (8, 'Text 9', 'top', 5.84), (8, 'Text 12', 'top', 5.84),
    (9, 'Text 12', 'height', 9.78), (4, 'Text 33', 'top', 5.84),
]
# 各页正文分隔线
_DIVIDERS = [(6, 'Shape 33', 1.40, 17.91, 30.99), (9, 'Shape 13', 17.40, 8.00, 11.05)]
# 四张示意图片
_PICTURES = {1: (17.40, 0.00, 16.46, 19.05), 2: (27.05, 12.57, 4.57, 3.81),
             10: (21.34, 4.45, 10.54, 5.21), 13: (19.18, 4.19, 12.57, 4.45)}
# 每页装饰元素：左上角色块 / 右下角色块 / 左侧竖条
_DECOR = {'Shape 0': (0.46, 0.46, 4.32, 2.92), 'Shape 1': (28.83, 15.37, 3.94, 2.67),
          'Shape 2': (0.00, 0.00, 0.36, 19.05)}
# 各页“并列说明卡片与图片区”的顶边，副标题说明行不得压到该线以下
_CONTENT_TOP = {1: 11.02, 2: 4.45, 3: 4.45, 4: 4.45, 5: 4.45, 6: 4.37, 7: 4.45,
                8: 4.45, 9: 4.45, 10: 4.32, 11: 4.45, 12: 4.45, 13: 4.19}


# ========== 通用小工具 ==========
def _cm(v):
    """EMU -> cm（保留两位，与办公软件“形状格式”面板显示口径一致）。"""
    if v is None:
        return None
    return round(Emu(int(v)).cm, 2)


def _near(a, b, tol=_GEO_TOL):
    return a is not None and b is not None and abs(a - b) <= tol


def _open_presentation(file_path):
    """打开 .pptx；打不开（损坏、非法包）时抛出异常，由外层兜底为 status='error'。"""
    return Presentation(file_path)


def _shape_map(prs):
    """(页序号从1开始, 形状名) -> shape；形状名在本套课件中稳定且唯一。"""
    out = {}
    for i, slide in enumerate(prs.slides, 1):
        for sh in slide.shapes:
            out[(i, sh.name)] = sh
    return out


def _text(sh):
    """文本框全文；软回车 \\v 统一按换行处理。"""
    if sh is None or not sh.has_text_frame:
        return ''
    return sh.text_frame.text.replace('\v', '\n')


def _paras(sh):
    if sh is None or not sh.has_text_frame:
        return []
    return list(sh.text_frame.paragraphs)


def _runs(sh):
    """有文字的 run 列表：(文本, 字号pt, 加粗, 颜色RGB大写, 西文字体, 中文字体)。"""
    out = []
    for p in _paras(sh):
        for r in p.runs:
            if r.text == '':
                continue
            sz = round(r.font.size.pt, 2) if r.font.size is not None else None
            try:
                clr = r.font.color.rgb
                clr = str(clr).upper() if clr is not None else None
            except Exception:
                clr = None
            rpr = r._r.find(_A + 'rPr')
            la = ea = None
            if rpr is not None:
                nl, ne = rpr.find(_A + 'latin'), rpr.find(_A + 'ea')
                la = nl.get('typeface') if nl is not None else None
                ea = ne.get('typeface') if ne is not None else None
            out.append((r.text, sz, bool(r.font.bold), clr, la, ea))
    return out


def _has_cjk(t):
    """是否残留汉字；化学元素符号、化学式、阿拉伯数字不属于汉字区间。"""
    return re.search(r'[一-鿿]', t or '') is not None


def _has_cyr(t):
    """是否含西里尔字母（蒙古文正文的判定依据）。"""
    return re.search(r'[Ѐ-ӿ]', t or '') is not None


def _translated(sh):
    """该文本框已译为蒙古文：有内容、无汉字残留、且出现西里尔字母。"""
    t = _text(sh)
    return bool(t.strip()) and not _has_cjk(t) and _has_cyr(t)


def _font_ok(sh, size=None, bold=None, color=None, latin=True):
    """文本框内每个有字的 run 是否都满足给定字号/加粗/颜色/字体。"""
    rs = _runs(sh)
    if not rs:
        return False
    for _t, sz, bd, clr, la, ea in rs:
        if size is not None and not _near(sz, size, 0.05):
            return False
        if bold is not None and bd != bold:
            return False
        if color is not None and (clr or '').upper() != color:
            return False
        if latin and (la not in _MSYH or ea not in _MSYH):
            return False
    return True


def _block_pt(sh):
    """文本块字号：取第一个有字 run 的字号；未指定字号时返回 None。"""
    for _t, sz, _b, _c, _la, _ea in _runs(sh):
        return sz
    return None


def _tail_empty(sh):
    """正文末尾追加的空段数量（最后一个有字段落之后的空段个数）。"""
    ps = [p.text.replace('\v', '\n') for p in _paras(sh)]
    if not any(p.strip() for p in ps):
        return 0
    n = 0
    for p in reversed(ps):
        if p.strip() == '':
            n += 1
        else:
            break
    return n


def _box(sh):
    if sh is None:
        return (None, None, None, None)
    return (_cm(sh.left), _cm(sh.top), _cm(sh.width), _cm(sh.height))


# ========== 版面占位模型（用于“完整可见、未折行溢出、未压盖”一条） ==========
# 与办公软件对齐：微软雅黑下按平均字宽估算——全角字符 1.0em、西里尔/拉丁字母
# 0.52em、数字 0.55em、空格 0.28em、标点 0.32em；行高取 1.2 倍字号；文本框
# 左右内边距取 bodyPr 的 lIns/rIns（缺省 0.254cm）；wrap="square" 时按可用宽度
# 折行；anchor="ctr"/"b" 时整块文字在框内居中/贴底展开，超高部分对称溢出。
def _wunits(t):
    w = 0.0
    for ch in t:
        if re.match(r'[一-鿿　-〿＀-￯]', ch):
            w += 1.0
        elif ch == ' ':
            w += 0.28
        elif ch.isdigit():
            w += 0.55
        elif ch in '.,:;!?\'"“”()-—·、。':
            w += 0.32
        else:
            w += 0.52
    return w


def _ins(bp, key, default):
    if bp is None or bp.get(key) is None:
        return default
    try:
        return int(bp.get(key)) / 360000.0
    except (TypeError, ValueError):
        return default


def _ink_span(sh):
    """返回 (字迹顶边cm, 字迹底边cm, 渲染总行数)；空段不计字迹。无字返回 None。"""
    if sh is None or not sh.has_text_frame:
        return None
    tf = sh.text_frame
    bp = tf._txBody.find(_A + 'bodyPr')
    anchor = bp.get('anchor', 't') if bp is not None else 't'
    avail = (_cm(sh.width) - _ins(bp, 'lIns', 0.254) - _ins(bp, 'rIns', 0.254)) * _PT_PER_CM
    if avail <= 0:
        return None
    rows, maxpt = [], 0.0
    for p in _paras(sh):
        pt = None
        for r in p.runs:
            if r.font.size is not None:
                pt = r.font.size.pt
                break
        pt = pt or 18.0
        maxpt = max(maxpt, pt)
        for seg in p.text.replace('\v', '\n').split('\n'):
            if seg.strip() == '':
                rows.append(False)
            else:
                rows.extend([True] * max(1, int(math.ceil(_wunits(seg) * pt / avail))))
    if not any(rows):
        return None
    lh = maxpt * 1.2 / _PT_PER_CM
    total, top, hgt = len(rows) * lh, _cm(sh.top), _cm(sh.height)
    if anchor == 'ctr':
        y0 = top + hgt / 2.0 - total / 2.0
    elif anchor == 'b':
        y0 = top + hgt - total
    else:
        y0 = top
    first = rows.index(True)
    last = len(rows) - 1 - rows[::-1].index(True)
    return (y0 + first * lh, y0 + (last + 1) * lh, len(rows))


def evaluate(dir_path):
    """按 §2.2 约定，返回结构化 dict；不 print、不改 sys.stdout、不 sys.exit。
    入参 dir_path 为“脚本所在目录的路径”，脚本自己在该目录里定位并打开被评估的
    .pptx（忽略以 ~$ 开头的 PowerPoint 临时锁文件）。脚本自身崩溃（含目录不存在、
    未发现文档等）由本函数捕获并返回 status='error'。"""
    result = {
        "id": SCRIPT_ID, "file_name": "", "status": "ok", "error": None,
        "dim1_pass": True, "dim1_reason": "", "dim2_items": [],
        "total_score": 0, "max_score": MAX_SCORE_TOTAL,
    }
    try:
        if not dir_path or not os.path.isdir(dir_path):
            raise FileNotFoundError(f"目录不存在：{dir_path}")
        cands = sorted(f for f in os.listdir(dir_path)
                       if f.lower().endswith('.pptx') and not f.startswith('~$'))
        if not cands:
            raise FileNotFoundError(f"目录中未找到 .pptx 文件：{dir_path}")
        file_path = os.path.join(dir_path, cands[0])
        result["file_name"] = os.path.basename(file_path)
        return _evaluate_impl(file_path, result)
    except Exception as exc:
        # 脚本未能进入评估（目录不存在／未找到文档／文档打不开）：维度一未曾判定，
        # dim1_pass 置空字符串以区别于“判定为不通过”的 False，分值全部归零。
        result["status"] = "error"
        result["error"] = (str(exc) if isinstance(exc, FileNotFoundError)
                           else f"{type(exc).__name__}: {exc}")
        result["dim1_pass"] = ""
        result["dim1_reason"] = ""
        result["dim2_items"] = []
        result["total_score"] = 0
        result["max_score"] = 0
        return result


def _dim1_check(prs, sm):
    """维度一：可用与可修改性。仅拦“打不开／整页是位图不可编辑／版面严重错乱”，
    不把任何完成度问题当门禁。返回 (pass, reason)。"""
    n = len(prs.slides)
    if n == 0:
        return False, "演示文稿内没有任何幻灯片，文件不可用"
    # ① 逐页必须有可单独点选、单独编辑的文本框文字；整页粘贴为一张位图时命中
    for i, slide in enumerate(prs.slides, 1):
        live = [sh for sh in slide.shapes
                if sh.has_text_frame and sh.text_frame.text.strip()]
        if len(live) < 3:
            return False, f"第{i}页可单独编辑的文本框不足3个（{len(live)}个），疑为整页图片粘贴或文字被转为轮廓"
    # ② 版面严重错乱：主标题与副标题说明行的字迹大面积重叠（>1.2cm）且波及≥3页
    heavy = 0
    for i in range(1, n + 1):
        ts, ss = _ink_span(sm.get((i, _TITLE))), _ink_span(sm.get((i, _SUB)))
        if ts and ss and ts[1] - ss[0] > 1.2:
            heavy += 1
    if heavy >= 3:
        return False, f"{heavy}页主标题与副标题说明行文字大面积重叠，内容无法辨认"
    return True, ""


def _evaluate_impl(file_path, result):
    """维度一 + 维度二的评估主体；对内部构造 results 列表沿用原逻辑，
    末尾按 _MAX_DELTAS_ORDER 反解成结构化 dim2_items 填入 result。"""
    prs = _open_presentation(file_path)
    sm = _shape_map(prs)
    npg = len(prs.slides)
    pages = list(range(1, npg + 1))
    results = []

    # ========== 维度1：可用与可修改性 ==========
    ok1, why1 = _dim1_check(prs, sm)
    result["dim1_pass"] = ok1
    result["dim1_reason"] = why1
    if not ok1:
        result["dim2_items"] = []
        result["total_score"] = 0
        result["max_score"] = MAX_SCORE_TOTAL
        return result

    # ========== 通用格式门槛（不计分，仅作为加分项的计分前提） ==========
    # 严格对齐细则：页面尺寸 33.86cm×19.05cm 且13页一致；蒙古文正文沿用微软雅黑；
    # 译文须可正常显示（非乱码、非方框）。门槛不满足时加分项不予计分。
    gate = (_near(_cm(prs.slide_width), _SLIDE_W, 0.05)
            and _near(_cm(prs.slide_height), _SLIDE_H, 0.05))
    if gate:
        # 蒙古文正文沿用微软雅黑：抽查13页主标题与副标题说明行的字体
        bad_font = 0
        for i in pages:
            for nm in (_TITLE, _SUB):
                rs = _runs(sm.get((i, nm)))
                if rs and any(r[4] not in _MSYH and r[5] not in _MSYH for r in rs):
                    bad_font += 1
        gate = bad_font == 0
    if gate:
        # 非乱码：正文不得出现替换字符或成片方框
        allt = ''.join(_text(sh) for sh in sm.values())
        gate = ('�' not in allt) and (allt.count('□') + allt.count('☐') < 3)

    # ---------- 01 交付为13页幻灯片，页序与右上角页码标签02—13一一对应 ----------
    # 严格对齐细则，仅踩细则明确写出的点，不额外约束：
    #   ① 幻灯片总数为13页，未增页、删页、并页；
    #   ② 第2—13页右上角页码标签文字依次为 02—13，与页序一一对应；
    #   ③ 页码为阿拉伯数字原样保留，未改为蒙古文数词、未重新编号。
    # 与办公软件对齐：页码标签是每页独立的文本框（形状名 Text 6），按页序读取。
    hit01 = gate and npg == 13
    if hit01:
        for i in range(2, 14):
            if _text(sm.get((i, _PAGENO))).strip() != '%02d' % i:
                hit01 = False
                break
    results.append(("+3：" if hit01 else "+0：")
                   + "交付为13页幻灯片，页序与右上角页码标签02—13一一对应，未增页、删页、并页，且页码数字原样保留未被改为蒙古文数词或重新编号")

    # ---------- 02 13页主标题全部译为蒙古文 ----------
    # 严格对齐细则，仅踩细则明确写出的点，不额外约束：
    #   ① 13页页面上部的单行大字主标题（形状名 Text 4）逐页检查；
    #   ② 每页均无汉字残留且出现西里尔字母，即视为已译出；
    #   ③ 13处无一处漏译（空文本框同样计为漏译）。
    # 与办公软件对齐：主标题为独立文本框，位于章节标签下方。
    hit02 = gate and all(_translated(sm.get((i, _TITLE))) for i in pages)
    results.append(("+5：" if hit02 else "+0：")
                   + "13页主标题（页面上部单行大字、位于章节标签下方）全部由汉语译为蒙古文，13处无一处漏译或保留汉字")


    # ---------- 03 13页章节标签译出且序号01—12原样保留 ----------
    # 严格对齐细则，仅踩细则明确写出的点，不额外约束：
    #   ① 13页主标题上方的橙色小字（形状名 Text 3）全部译为蒙古文；
    #   ② 第2—13页标签内两位序号 01—12 原样保留；
    #   ③ 第1页“第二章 · 第一节”无阿拉伯序号，仅判定译文。
    # 与办公软件对齐：章节标签为独立文本框，颜色 #E69A3A、加粗 13磅（第1页15磅）。
    hit03 = gate and all(_translated(sm.get((i, _KICKER))) for i in pages)
    if hit03:
        for i in range(2, 14):
            if '%02d' % (i - 1) not in _text(sm.get((i, _KICKER))):
                hit03 = False
                break
    results.append(("+3：" if hit03 else "+0：")
                   + "13页章节标签（主标题上方的橙色小字，如\"第二章 · 第一节\"\"概念 01\"\"规则 02\"\"方法 09\"\"整理 11\"\"练习 12\"）全部译为蒙古文，且序号01—12原样保留")

    # ---------- 04 13页副标题说明行全部译为蒙古文 ----------
    # 严格对齐细则，仅踩细则明确写出的点，不额外约束：
    #   ① 第2—13页主标题正下方的浅灰单行长句（形状名 Text 5）；
    #   ② 第1页同名文本框为封面概览说明行，一并计入13处；
    #   ③ 逐处要求无汉字残留且出现西里尔字母。
    # 与办公软件对齐：副标题说明行为独立文本框，颜色 #6B7C84。
    hit04 = gate and all(_translated(sm.get((i, _SUB))) for i in pages)
    results.append(("+3：" if hit04 else "+0：")
                   + "13页副标题说明行（主标题正下方的浅灰单行长句）全部译为蒙古文")

    # ---------- 05 13页底部两行页脚全部译为蒙古文 ----------
    # 严格对齐细则，仅踩细则明确写出的点，不额外约束：
    #   ① 逐页取底部“承接：…”“下一步：…”两行页脚（第1页为合并的一行）；
    #   ② 每行均无汉字残留且出现西里尔字母；
    #   ③ 共25处，缺一不可。
    # 与办公软件对齐：页脚为顶端 18.19cm（第1页 17.42cm）的独立文本框。
    hit05 = gate
    for i, names in _FOOTERS.items():
        for nm in names:
            if not _translated(sm.get((i, nm))):
                hit05 = False
    results.append(("+3：" if hit05 else "+0：")
                   + "13页底部两行\"承接：…\"\"下一步：…\"页脚全部译为蒙古文")

    # ---------- 06 第5—12页并列说明卡片约30个文本块译出 ----------
    # 严格对齐细则，仅踩细则明确写出的点，不额外约束：
    #   ① 覆盖第5—9页与第10—12页共八页的并列说明卡片；
    #   ② 每页 2—3 个小标题＋对应正文段落，逐块判定；
    #   ③ 全部文本块无中文残留。
    # 与办公软件对齐：卡片小标题与正文分属独立文本框，叠放在圆角色块之上。
    hit06 = gate
    for i, names in _CARD_BLOCKS.items():
        for nm in names:
            if not _translated(sm.get((i, nm))):
                hit06 = False
    results.append(("+5：" if hit06 else "+0：")
                   + "第5—9页与第10—12页共八页的并列说明卡片（每页2—3个小标题＋正文段落）标题与正文全部译为蒙古文，约30个文本块无中文残留")


    # ---------- 07 第1页五个圆角标签译出且互不重复 ----------
    # 严格对齐细则，仅踩细则明确写出的点，不额外约束：
    #   ① 自左至右五个圆角标签（形状名 Text 7/9/11/13/15）全部译为蒙古文；
    #   ② 五个词互不重复；
    #   ③ 第5个须为“异构（изомер）”义，不得与第4个“命名（нэршил）”重复。
    # 与办公软件对齐：五个标签为叠在圆角色块上的独立文本框，居中、白色、10.5磅。
    chips = [_text(sm.get((1, nm))).strip() for nm in _CHIPS]
    hit07 = gate and all(_translated(sm.get((1, nm))) for nm in _CHIPS)
    if hit07:
        low = [c.lower() for c in chips]
        hit07 = len(set(low)) == 5 and 'изомер' in low[4]
    results.append(("+3：" if hit07 else "+0：")
                   + "第1页五个圆角标签位（自左至右第1—5个，对应\"烃基\"\"主链\"\"编号\"\"命名\"\"异构\"）全部译为蒙古文，且五个词互不重复——第5个标签须为\"异构（изомер）\"义，不得与第4个\"命名（нэршил）\"重复")

    # ---------- 08 三处箭头流程行译出并保留九个“→” ----------
    # 严格对齐细则，仅踩细则明确写出的点，不额外约束：
    #   ① 第1页学习路线行、第8页记忆口令行、第13页小结行三处；
    #   ② 三处均译为蒙古文；
    #   ③ 三处“→”数量依次为 3/2/4，合计九个，全部保留。
    # 与办公软件对齐：箭头为普通文本字符，随正文一起排版。
    hit08 = gate
    for i, nm, cnt in _ARROW_ROWS:
        sh = sm.get((i, nm))
        if not _translated(sh) or _text(sh).count('→') != cnt:
            hit08 = False
    results.append(("+3：" if hit08 else "+0：")
                   + "三处箭头流程行（第1页\"学习路线：概念打底 → 规则拆解 → 书写训练 → 判断巩固\"、第8页\"近 → 简 → 小\"记忆口令、第13页\"小结：识别烷基 → 选择主链 → 规范编号 → 写出名称 → 检查异构\"）全部译为蒙古文，并保留其中共九个\"→\"箭头")

    # ---------- 09 第3页碳数—名称对照表表头与十个烷烃名 ----------
    # 严格对齐细则，仅踩细则明确写出的点，不额外约束：
    #   ① 两列表头（“碳数”“名称”位，形状名 Text 8 / Text 30）译为蒙古文；
    #   ② 十个名称块依次为 Метан—Декан（大小写不限）；
    #   ③ 左列十个数字 1—10 原样保留、顺序不变。
    # 与办公软件对齐：表格由 20 个圆角色块＋20 个独立文本框拼成，按左端坐标排序。
    hit09 = gate and _translated(sm.get((3, 'Text 8'))) and _translated(sm.get((3, 'Text 30')))
    if hit09:
        for nm, want in zip(_TAB_NUM, _TAB_DIGITS):
            if _text(sm.get((3, nm))).strip() != want:
                hit09 = False
        for nm, want in zip(_TAB_NAME, _ALKANES):
            if _text(sm.get((3, nm))).strip().lower() != want:
                hit09 = False
    results.append(("+5：" if hit09 else "+0：")
                   + "第3页碳数—名称对照表的两列表头（\"碳数\"\"名称\"位）与十个烷烃名称（\"甲烷\"—\"癸烷\"位）全部译为蒙古文（Метан／Этан／Пропан／Бутан／Пентан／Гексан／Гептан／Октан／Нонан／Декан），且左列数字1—10共十个数字原样保留、顺序不变")


    # ---------- 10 第3页三个小标题及说明段、“十二烷” ----------
    # 严格对齐细则，仅踩细则明确写出的点，不额外约束：
    #   ① “碳数不超过十”“碳数超过十”“连接下一页”三个小标题（Text 52/55/58）译出；
    #   ② 三个小标题下方说明段（Text 53/56/59）译出；
    #   ③ “十二烷”译为对应蒙古文烷烃名（додекан）。
    # 与办公软件对齐：三块说明卡片位于对照表下方，标题与正文分属独立文本框。
    hit10 = gate and all(_translated(sm.get((3, nm)))
                         for nm in ('Text 52', 'Text 55', 'Text 58',
                                    'Text 53', 'Text 56', 'Text 59'))
    if hit10:
        hit10 = 'додекан' in _text(sm.get((3, 'Text 56'))).lower()
    results.append(("+3：" if hit10 else "+0：")
                   + "第3页\"碳数不超过十\"\"碳数超过十\"\"连接下一页\"三个小标题及其下方说明段译为蒙古文，\"十二烷\"译为对应蒙古文烷烃名")

    # ---------- 11 第4页三个碳架卡片与三段说明、三个戊烷命名 ----------
    # 严格对齐细则，仅踩细则明确写出的点，不额外约束：
    #   ① 三个卡片标题（Text 8/20/32）与三段说明（Text 9/21/33）译为蒙古文；
    #   ② “正戊烷”“异戊烷”“新戊烷”依次译为 n-пентан／изопентан／неопентан；
    #   ③ 三个命名分别出现在对应卡片的说明段中。
    # 与办公软件对齐：三张卡片自左至右排布，说明段与碳架示意图分属不同形状。
    hit11 = gate and all(_translated(sm.get((4, nm)))
                         for nm in ('Text 8', 'Text 20', 'Text 32',
                                    'Text 9', 'Text 21', 'Text 33'))
    if hit11:
        for nm, want in (('Text 9', 'n-пентан'), ('Text 21', 'изопентан'),
                         ('Text 33', 'неопентан')):
            if want not in _text(sm.get((4, nm))).lower():
                hit11 = False
    results.append(("+5：" if hit11 else "+0：")
                   + "第4页三个碳架卡片标题（\"直链碳架\"\"一个支链\"\"中心分支\"位）与三段说明译为蒙古文，且\"正戊烷\"\"异戊烷\"\"新戊烷\"分别译为对应蒙古文命名（n-пентан／изопентан／неопентан）")

    # ---------- 12 第2页四个烷基名称 ----------
    # 严格对齐细则，仅踩细则明确写出的点，不额外约束：
    #   ① 四个烷基名称卡片标题（Text 14/17/20/23）译为蒙古文；
    #   ② 逐处无汉字残留；
    #   ③ 仅判定名称本身，其下方化学式另由化学式条目判定。
    # 与办公软件对齐：四个名称为右侧 2×2 卡片的标题文本框。
    hit12 = gate and all(_translated(sm.get((2, nm)))
                         for nm in ('Text 14', 'Text 17', 'Text 20', 'Text 23'))
    results.append(("+3：" if hit12 else "+0：")
                   + "第2页四个烷基名称（\"甲基\"\"乙基\"\"正丙基\"\"异丙基\"位）译为蒙古文")


    # ---------- 13 全篇化学式与元素符号原样保留 ----------
    # 严格对齐细则，仅踩细则明确写出的点，不额外约束：
    #   ① 第2页四个烷基化学式、通式—CnH2n+1、片段符号—R 原样；
    #   ② 第4页十个碳架元素标签（CH3／CH2／CH）与中心C、第6页CH3、第10页六个C 原样；
    #   ③ 分子式 C5H12、C6H14 原样，含大小写、下标数字与连接短线。
    # 与办公软件对齐：化学式为独立文本框或正文内的普通字符，字体沿用 Aptos。
    hit13 = gate
    for (i, nm), want in _FORMULA_EXACT.items():
        if _text(sm.get((i, nm))).strip() != want:
            hit13 = False
    for (i, nm), wants in _FORMULA_SUBSTR.items():
        t = _text(sm.get((i, nm)))
        for w in wants:
            if w not in t:
                hit13 = False
    results.append(("+5：" if hit13 else "+0：")
                   + "全篇化学式与元素符号原样保留、未被译成蒙古文或改写——第2页—CH3、—CH2CH3、—CH2CH2CH3、—CH(CH3)2、通式—CnH2n+1、片段符号—R；第4页十个碳架元素标签（CH3／CH2／CH）与中心C；第6页CH3；第10页六个C；分子式C5H12（第4页副标题说明行）与C6H14（第10页训练对象标题、第12页骨架清单标题），含元素符号大小写、下标数字与连接短线")

    # ---------- 14 第12页五个同分异构体名称与位号 ----------
    # 严格对齐细则，仅踩细则明确写出的点，不额外约束：
    #   ① 五个名称依次为 n-гексан／2-метилпентан／3-метилпентан／
    #      2,2-диметилбутан／2,3-диметилбутан；
    #   ② 位号 2、3、2,2、2,3 原样保留；
    #   ③ 位号中的逗号原样保留（不得改为顿号或空格）。
    # 与办公软件对齐：五个名称位于同一文本框的五个段落（骨架清单块 Text 12）。
    lines14 = [s.strip().lower() for s in _text(sm.get((12, 'Text 12'))).split('\n')
               if s.strip()]
    hit14 = gate and len(lines14) >= 5 and all(
        want in lines14[k] for k, want in enumerate(_ISOMERS))
    results.append(("+5：" if hit14 else "+0：")
                   + "第12页五个同分异构体名称译为蒙古文命名法（n-гексан／2-метилпентан／3-метилпентан／2,2-диметилбутан／2,3-диметилбутан），且位号2、3、2,2、2,3与其中的逗号原样保留")

    # ---------- 15 位号与编号数字原样保留 ----------
    # 严格对齐细则，仅踩细则明确写出的点，不额外约束：
    #   ① 第6页主链编号七个数字标签依次为 1—7；
    #   ② 第8页方向A位号组合为“2、3、5”、方向B为“2、4、5”，数值与排列顺序不变；
    #   ③ 两页其余说明文字已译为蒙古文。
    # 与办公软件对齐：编号标签为叠在小圆块上的独立文本框，位号组合写在正文内。
    hit15 = gate
    for k, nm in enumerate(_CHAIN_NUM, 1):
        if _text(sm.get((6, nm))).strip() != str(k):
            hit15 = False
    for nm, pat in (('Text 9', r'2\D{0,3}3\D{0,3}5'), ('Text 12', r'2\D{0,3}4\D{0,3}5')):
        sh = sm.get((8, nm))
        if not _translated(sh) or re.search(pat, _text(sh)) is None:
            hit15 = False
    results.append(("+3：" if hit15 else "+0：")
                   + "位号与编号数字原样保留（第6页主链编号七个数字标签1—7、第8页方向A与方向B的位号组合\"2、3、5\"与\"2、4、5\"），数值与排列顺序不变，其余说明文字译为蒙古文")


    # ---------- 16 第13页题目A、题目B全文译出且数量关系保留 ----------
    # 严格对齐细则，仅踩细则明确写出的点，不额外约束：
    #   ① 题目A、题目B两块标题与正文（Text 8/9、Text 11/12）全部译为蒙古文；
    #   ② 相对分子质量数值 100 原样保留；
    #   ③ “主链含四个碳”“总碳数为六”“三个甲基”三处数量关系未被改写或漏译
    #      （数词写作蒙古文词或阿拉伯数字均可）。
    # 与办公软件对齐：两道练习题各为一张卡片，标题与正文分属独立文本框。
    hit16 = gate and all(_translated(sm.get((13, nm)))
                         for nm in ('Text 8', 'Text 9', 'Text 11', 'Text 12'))
    if hit16:
        ta, tb = _text(sm.get((13, 'Text 9'))).lower(), _text(sm.get((13, 'Text 12'))).lower()
        hit16 = (re.search(r'дөрв|4', ta) is not None
                 and re.search(r'зургаа|6', ta) is not None
                 and '100' in tb
                 and re.search(r'гурв|3', tb) is not None
                 and 'метил' in tb)
    results.append(("+3：" if hit16 else "+0：")
                   + "第13页题目A、题目B两道练习题全文译为蒙古文，且相对分子质量数值100、\"主链含四个碳\"\"总碳数为六\"\"三个甲基\"等数量关系未被改写或漏译")

    # ---------- 17 主标题与副标题译文完整可见、未折行溢出、未压盖 ----------
    # 严格对齐细则，仅踩细则明确写出的点，不额外约束：
    #   ① 第2—13页主标题为单行大字，译文后仍须排为1行，不得折行溢出文本框；
    #   ② 主标题字迹不得与上方章节标签、下方副标题说明行的字迹相互压盖；
    #   ③ 副标题说明行字迹不得压到下方并列说明卡片与图片区的顶边。
    # 与办公软件对齐：文本框 wrap="square"、anchor="ctr"，文字在框内垂直居中后
    # 对称溢出；按微软雅黑平均字宽估算折行，压盖判定留 0.25cm 容差。
    hit17 = gate
    for i in pages:
        ts = _ink_span(sm.get((i, _TITLE)))
        ss = _ink_span(sm.get((i, _SUB)))
        ks = _ink_span(sm.get((i, _KICKER)))
        if ts is None or ss is None:
            hit17 = False
            continue
        if i >= 2 and ts[2] > 1:
            hit17 = False
        if ks is not None and ks[1] - ts[0] > _OVERLAP_TOL:
            hit17 = False
        if ts[1] - ss[0] > _OVERLAP_TOL:
            hit17 = False
        if ss[1] - _CONTENT_TOP.get(i, 4.45) > _OVERLAP_TOL:
            hit17 = False
    results.append(("+5：" if hit17 else "+0：")
                   + "13页主标题与副标题说明行的蒙古文译文完整可见，未因译文变长而折行溢出、未互相压盖，也未压到下方并列说明卡片与图片区")


    # ---------- 18 副标题说明行字体格式 微软雅黑 12.5磅 #6B7C84 ----------
    # 严格对齐细则，仅踩细则明确写出的点，不额外约束：
    #   ① 判定对象为第2—13页主标题正下方、左端1.47cm、顶端3.45cm、宽24.89cm 的
    #      浅灰单行（第1页同名文本框为封面概览说明行，字号目标16磅，另条判定）；
    #   ② 字体须为微软雅黑、字号 12.5磅、颜色 #6B7C84；
    #   ③ 出现 11.5磅、10.75磅等等比缩放后的非整数磅值即计入本条。
    # 与办公软件对齐：字号取“开始”选项卡显示值，逐 run 校验。
    bad_sub = [i for i in range(2, npg + 1)
               if not _font_ok(sm.get((i, _SUB)), size=12.5, color=_GRAY)]
    results.append(("-5：" if bad_sub else "-0：")
                   + "13页副标题说明行（主标题正下方、左端1.47cm、顶端3.45cm、宽24.89cm的浅灰单行）不满足字体格式为微软雅黑 12.5磅 颜色#6B7C84，出现11.5磅、10.75磅等等比缩放后的非整数磅值")

    # ---------- 19 第3、6、11页主标题 微软雅黑 26磅 加粗 #0B5965 ----------
    # 严格对齐细则，仅踩细则明确写出的点，不额外约束：
    #   ① 仅判定第3、6、11页三处主标题；
    #   ② 须为微软雅黑、26磅、加粗、颜色 #0B5965；
    #   ③ 出现 23.92磅等等比缩放后的非整数磅值即计入本条。
    # 与办公软件对齐：主标题为单段落文本框，整行同一格式。
    bad_ttl = [i for i in (3, 6, 11)
               if not _font_ok(sm.get((i, _TITLE)), size=26.0, bold=True, color=_TEAL)]
    results.append(("-3：" if bad_ttl else "-0：")
                   + "第3、6、11页主标题不满足字体格式为微软雅黑 26磅 加粗 颜色#0B5965，出现23.92磅等等比缩放后的非整数磅值")

    # ---------- 20 页脚 微软雅黑 9.5磅 #6B7C84 ----------
    # 严格对齐细则，仅踩细则明确写出的点，不额外约束：
    #   ① 判定13页底部“承接：…”“下一步：…”两行页脚共25处；
    #   ② 须为微软雅黑、9.5磅、颜色 #6B7C84；
    #   ③ 出现 8.74磅等等比缩放后的非整数磅值即计入本条。
    # 与办公软件对齐：页脚为顶端 18.19cm（第1页 17.42cm）的独立文本框。
    bad_ft = [(i, nm) for i, names in _FOOTERS.items() for nm in names
              if not _font_ok(sm.get((i, nm)), size=9.5, color=_GRAY)]
    results.append(("-3：" if bad_ft else "-0：")
                   + "13页底部\"承接：…\"\"下一步：…\"两行页脚（顶端18.19cm）不满足字体格式为微软雅黑 9.5磅 颜色#6B7C84，出现8.74磅等等比缩放后的非整数磅值")


    # ---------- 21 各页卡片正文与说明行字号分级 ----------
    # 严格对齐细则，仅踩细则明确写出的点，不额外约束：
    #   ① 逐块比对 30 处正文/说明行应有字号，无论调大或调小均计入，同块只算一次；
    #   ② 成批改动（偏差块数≥10，或出现 12.04／12.42／13.33／13.76／14.62／14.72／
    #      15.48／15.64／17.48／18.06／18.4／19.32磅等按0.86倍／0.92倍缩放的磅值）扣5分；
    #   ③ 仅个别文本块微调（≤3处且偏差≤1磅）扣1分；无偏差不扣。
    # 与办公软件对齐：块字号取该文本框首个有字 run 的字号；未指定字号视为偏差。
    dev, dmax, scaled = [], 0.0, False
    for (i, nm), want in sorted(_BODY_PT.items()):
        sh = sm.get((i, nm))
        if sh is None:
            continue
        got = _block_pt(sh)
        if got is None:
            # 与办公软件对齐：文本框正文未显式指定字号时按默认 18 磅 渲染，
            # 因此“未写 sz”本身不算偏差，只有渲染磅值与目标不一致才计入。
            got = 18.0
        if abs(got - want) > 0.05:
            dev.append((i, nm))
            dmax = max(dmax, abs(got - want))
        if any(abs(got - s) <= 0.01 for s in _SCALED_PT):
            scaled = True
    if not dev:
        d21 = 0
    elif len(dev) >= 10 or scaled:
        d21 = -5
    elif len(dev) <= 3 and dmax <= 1.0:
        d21 = -1
    else:
        d21 = -5
    results.append(("%+d：" % d21 if d21 else "-0：")
                   + "各页并列说明卡片正文与说明行不满足各自应有字号（第1页概览说明行16磅与\"学习路线\"行14磅、第2页两段14磅、第3页两段14磅与\"连接下一页\"段13.5磅、第4页三段13.5磅、第5页\"长\"\"多\"两段17磅与\"结果\"段16磅、第6页核心判断段16磅、第7页判断顺序四条17磅与课堂表达段16磅、第8页方向A/B两段19磅与记忆口令段20磅、第9页组合模板段18磅、第10页三步骤段18磅、第11页两段16磅与书写原则段15.5磅、第12页补氢规则段与骨架清单段21磅、第13页题目A/B段15.5磅与答题提示段17磅、小结行16磅）——无论调大或调小均计入本条，同一文本块只扣一次：成批改动（≥10处，或出现按0.86倍／0.92倍缩放后的12.04、12.42、13.33、13.76、14.62、14.72、15.48、15.64、17.48、18.06、18.4、19.32磅等非整数磅值）扣5分；仅个别文本块微调（≤3处且偏差≤1磅）扣1分")

    # ---------- 22 第4页三个碳架说明块末尾不追加空段 ----------
    # 严格对齐细则，仅踩细则明确写出的点，不额外约束：
    #   ① 判定第4页三个碳架说明块（Text 9／21／33）；
    #   ② 说明句所在段落之后不得追加空段占位；
    #   ③ 出现 3个／6个／8个空行段落使内容整体下坠即计入本条。
    # 与办公软件对齐：空段指段落无任何可见字符（含仅空白的 run）。
    bad22 = [nm for nm in ('Text 9', 'Text 21', 'Text 33')
             if _tail_empty(sm.get((4, nm))) > 0]
    results.append(("-3：" if bad22 else "-0：")
                   + "第4页三个碳架说明块不满足\"说明句所在段落之后不追加空段占位\"——不得在说明句末尾补出3个、6个、8个空行段落使文本框内容整体下坠")


    # ---------- 23 十处正文末尾不追加空段 ----------
    # 严格对齐细则，仅踩细则明确写出的点，不额外约束：
    #   ① 判定第1页主标题块、第3页“碳数不超过十”说明段、第4页页脚承接行、
    #      第6页核心判断块、第7页判断顺序块与课堂表达块、第8页记忆口令块、
    #      第9页组合模板块、第10页三步骤块、第12页骨架清单块共十处；
    #   ② 十处均不得在正文末尾追加空段占位；
    #   ③ 正文中间原有的空段用于分隔，不计入本条。
    # 与办公软件对齐：只统计最后一个有字段落之后的空段。
    bad23 = [(i, nm) for i, nm in _NO_TAIL_EMPTY if _tail_empty(sm.get((i, nm))) > 0]
    results.append(("-3：" if bad23 else "-0：")
                   + "第1页主标题块、第3页\"碳数不超过十\"说明段、第4页页脚承接行、第6页核心判断块、第7页判断顺序块与课堂表达块、第8页记忆口令块、第9页组合模板块、第10页三步骤块、第12页骨架清单块不满足\"正文末尾不追加空段占位\"")

    # ---------- 24 第4页三段碳架说明 微软雅黑 #22313F ----------
    # 严格对齐细则，仅踩细则明确写出的点，不额外约束：
    #   ① 判定第4页三段碳架说明文字（Text 9／21／33）；
    #   ② 字体须为微软雅黑、颜色须为 #22313F；
    #   ③ 不得出现未指定字体与颜色的裸文字段落。
    # 与办公软件对齐：逐 run 校验 latin/ea 字体与字符颜色是否显式指定。
    bad24 = []
    for nm in ('Text 9', 'Text 21', 'Text 33'):
        rs = _runs(sm.get((4, nm)))
        if not rs or any(r[3] != _INK or r[4] not in _MSYH or r[5] not in _MSYH
                         for r in rs):
            bad24.append(nm)
    results.append(("-3：" if bad24 else "-0：")
                   + "第4页三段碳架说明文字不满足字体为微软雅黑、颜色为#22313F——不得出现未指定字体与颜色的裸文字段落")

    # ---------- 25 第5页强调行整行 微软雅黑 14磅 加粗 #0B5965 ----------
    # 严格对齐细则，仅踩细则明确写出的点，不额外约束：
    #   ① 判定第5页“主链是系统命名的骨架中心”位强调行（Text 16）；
    #   ② 整行须为同一格式：微软雅黑、14磅、加粗、颜色 #0B5965；
    #   ③ 不得被拆成 17磅与 14磅两段。
    # 与办公软件对齐：同一行内出现两种字号即视为被拆分。
    bad25 = not _font_ok(sm.get((5, 'Text 16')), size=14.0, bold=True, color=_TEAL)
    results.append(("-3：" if bad25 else "-0：")
                   + "第5页\"主链是系统命名的骨架中心\"位强调行不满足整行同一格式为微软雅黑 14磅 加粗 颜色#0B5965，不得被拆成17磅与14磅两段")

    # ---------- 26 中文标点残留 ----------
    # 严格对齐细则，仅踩细则明确写出的点，不额外约束：
    #   ① 第4页“中心分支”卡片说明段（Text 33）句末不得残留中文句号“。”；
    #   ② 第1页页脚承接行（Text 18）不得残留全角空格“　”；
    #   ③ 第8页方向A与方向B两段（Text 9／12）不得残留顿号“、”。
    # 与办公软件对齐：按字符逐处检索，不涉及字体与字号。
    bad26 = []
    if '。' in _text(sm.get((4, 'Text 33'))):
        bad26.append('4-。')
    if '　' in _text(sm.get((1, 'Text 18'))):
        bad26.append('1-全角空格')
    for nm in ('Text 9', 'Text 12'):
        if '、' in _text(sm.get((8, nm))):
            bad26.append('8-' + nm)
    results.append(("-3：" if bad26 else "-0：")
                   + "不满足多余中文标点已随文本一并转为蒙古文行文标点——第4页\"中心分支\"卡片说明段句末残留中文句号\"。\"、第1页页脚承接行残留全角空格\"　\"、第8页方向A与方向B两段残留顿号\"、\"")


    # ---------- 27 13页主标题文本框 左1.4 顶2.08 高1.22 ----------
    # 严格对齐细则，仅踩细则明确写出的点，不额外约束：
    #   ① 第2—13页主标题文本框须为左端1.4cm、顶端2.08cm、高1.22cm；
    #   ② 第1页封面主标题按其自身目标值判定：左端1.52cm、顶端3.94cm、高4.06cm；
    #   ③ 为容纳更长译文而加宽宽度不计入本条。
    # 与办公软件对齐：坐标取“大小和属性”面板的位置与高度，容差 0.06cm。
    bad27 = []
    for i in pages:
        l, t, _w, h = _box(sm.get((i, _TITLE)))
        want = (1.52, 3.94, 4.06) if i == 1 else (1.40, 2.08, 1.22)
        if not (_near(l, want[0]) and _near(t, want[1]) and _near(h, want[2])):
            bad27.append(i)
    results.append(("-3：" if bad27 else "-0：")
                   + "13页主标题文本框不满足左端1.4cm、顶端2.08cm、高1.22cm（为容纳更长译文而加宽宽度不计入本条）")

    # ---------- 28 13页副标题说明行 左1.47 顶3.45 宽24.89 高0.71 ----------
    # 严格对齐细则，仅踩细则明确写出的点，不额外约束：
    #   ① 第2—13页副标题说明行须为左端1.47cm、顶端3.45cm、宽24.89cm、高0.71cm；
    #   ② 不得下移至3.68cm；
    #   ③ 不得改为宽32.05cm高0.74cm或宽31.56cm高0.66cm。
    # 与办公软件对齐：第1页同名文本框为封面概览说明行，不适用本条目标值。
    bad28 = []
    for i in range(2, npg + 1):
        l, t, w, h = _box(sm.get((i, _SUB)))
        if not (_near(l, 1.47) and _near(t, 3.45) and _near(w, 24.89) and _near(h, 0.71)):
            bad28.append(i)
    results.append(("-3：" if bad28 else "-0：")
                   + "13页副标题说明行不满足左端1.47cm、顶端3.45cm、宽24.89cm、高0.71cm（不得下移至3.68cm，不得改为宽32.05cm高0.74cm或宽31.56cm高0.66cm）")

    # ---------- 29 各页并列说明卡片定位与块高 ----------
    # 严格对齐细则，仅踩细则明确写出的点，不额外约束：
    #   ① 第6页核心判断块左端2.36cm、顶端5.77cm；
    #   ② 第8页说明卡片正文块顶端5.84cm、第4页“中心分支”卡片顶端5.84cm；
    #   ③ 第9页组合模板块高9.78cm。
    # 与办公软件对齐：逐项只比对细则点名的那一个属性，其余属性不参与判定。
    bad29 = []
    for i, nm, attr, want in _CARD_GEO:
        l, t, _w, h = _box(sm.get((i, nm)))
        got = {'left': l, 'top': t, 'height': h}[attr]
        if not _near(got, want):
            bad29.append((i, nm, attr))
    results.append(("-3：" if bad29 else "-0：")
                   + "各页并列说明卡片不满足各自定位与块高（第6页核心判断块左端2.36cm 顶端5.77cm、第8页课堂表达块顶端5.84cm、第9页组合模板块高9.78cm、第4页\"中心分支\"卡片顶端5.84cm）")

    # ---------- 30 各页正文分隔线坐标 ----------
    # 严格对齐细则，仅踩细则明确写出的点，不额外约束：
    #   ① 第6页分隔线顶端17.91cm、宽30.99cm；
    #   ② 第9页短分隔线左端17.4cm、顶端8.0cm、宽11.05cm；
    #   ③ 仅判定这两条，其余页分隔线不计入本条。
    # 与办公软件对齐：分隔线为高度 0 的直线形状；重画分隔线会改变形状名，
    # 故按“该页是否存在一条位于目标坐标的直线”判定，不绑定形状名。
    bad30 = []
    for i, nm, wl, wt, ww in _DIVIDERS:
        hit30 = False
        for sh in prs.slides[i - 1].shapes:
            l, t, w, h = _box(sh)
            if l is None or h is None or h > 0.05:
                continue
            if _near(l, wl) and _near(t, wt) and _near(w, ww):
                hit30 = True
                break
        if not hit30:
            bad30.append((i, nm))
    results.append(("-1：" if bad30 else "-0：")
                   + "各页正文分隔线不满足第6页顶端17.91cm 宽30.99cm、第9页短分隔线左端17.4cm 顶端8.0cm 宽11.05cm")


    # ---------- 31 第1页五个圆角标签坐标与文字格式 ----------
    # 严格对齐细则，仅踩细则明确写出的点，不额外约束：
    #   ① 五个标签顶端11.15cm、宽2.59cm、高0.56cm；
    #   ② 左端自左至右依次为 1.70／4.78／7.85／10.92／13.59cm；
    #   ③ 文字为微软雅黑 10.5磅 加粗 颜色#FFFFFF。
    # 与办公软件对齐：标签文字所在文本框独立于其下方的圆角色块。
    bad31 = []
    for nm, wl in zip(_CHIPS, _CHIP_LEFTS):
        sh = sm.get((1, nm))
        l, t, w, h = _box(sh)
        if not (_near(l, wl) and _near(t, 11.15) and _near(w, 2.59) and _near(h, 0.56)):
            bad31.append(nm)
        elif not _font_ok(sh, size=10.5, bold=True, color=_WHITE):
            bad31.append(nm)
    results.append(("-1：" if bad31 else "-0：")
                   + "第1页五个圆角标签不满足顶端11.15cm、宽2.59cm、高0.56cm、左端依次1.70／4.78／7.85／10.92／13.59cm，且文字为微软雅黑 10.5磅 加粗 颜色#FFFFFF")

    # ---------- 32 第3页碳数对照表十数字块与十名称块字号坐标 ----------
    # 严格对齐细则，仅踩细则明确写出的点，不额外约束：
    #   ① 十个数字块与十个名称块字号均为13磅；
    #   ② 数字块顶端4.88cm、名称块顶端6.45cm，宽2.72cm、高0.64cm；
    #   ③ 左端自左至右依次为 4.63／7.35／10.08／12.80／15.53／18.25／20.98／
    #      23.70／26.43／29.15cm，不得错行、错位或字号缩放。
    # 与办公软件对齐：表格由独立文本框拼成，逐块比对坐标与字号。
    bad32 = []
    for nm, wl in zip(_TAB_NUM, _TAB_NUM_LEFT):
        sh = sm.get((3, nm))
        l, t, w, h = _box(sh)
        if not (_near(l, wl) and _near(t, 4.88) and _near(w, 2.72) and _near(h, 0.64)):
            bad32.append(nm)
        elif not _font_ok(sh, size=13.0):
            bad32.append(nm)
    for nm, wl in zip(_TAB_NAME, _TAB_NUM_LEFT):
        sh = sm.get((3, nm))
        l, t, w, h = _box(sh)
        if not (_near(l, wl) and _near(t, 6.45) and _near(w, 2.72) and _near(h, 0.64)):
            bad32.append(nm)
        elif not _font_ok(sh, size=13.0):
            bad32.append(nm)
    results.append(("-1：" if bad32 else "-0：")
                   + "第3页碳数—名称对照表的十个数字块与十个名称块不满足字号13磅及各自原有坐标，出现错行、错位或字号缩放")

    # ---------- 33 四张示意图片与页面装饰元素坐标 ----------
    # 严格对齐细则，仅踩细则明确写出的点，不额外约束：
    #   ① 四张示意图片坐标尺寸（第1／2／10／13页）原样；
    #   ② 每页左上角色块、右下角色块、左侧竖条坐标尺寸原样；
    #   ③ 右上角页码标签左端30.73cm 顶端1.07cm，Aptos 17磅 加粗 颜色#1B8C8F。
    # 与办公软件对齐：图片以 PICTURE 形状读取，装饰色块按形状名定位。
    bad33 = []
    for i, want in _PICTURES.items():
        pic = None
        for sh in prs.slides[i - 1].shapes:
            if sh.shape_type is not None and 'PICTURE' in str(sh.shape_type):
                pic = sh
                break
        if pic is None:
            bad33.append((i, 'picture'))
            continue
        l, t, w, h = _box(pic)
        if not all(_near(a, b) for a, b in zip((l, t, w, h), want)):
            bad33.append((i, 'picture'))
    for i in pages:
        for nm, want in _DECOR.items():
            l, t, w, h = _box(sm.get((i, nm)))
            if not all(_near(a, b) for a, b in zip((l, t, w, h), want)):
                bad33.append((i, nm))
    for i in range(2, npg + 1):
        sh = sm.get((i, _PAGENO))
        l, t, _w, _h = _box(sh)
        rs = _runs(sh)
        if not (_near(l, 30.73) and _near(t, 1.07)):
            bad33.append((i, _PAGENO))
        elif not rs or any(not _near(r[1], 17.0, 0.05) or not r[2]
                           or (r[3] or '') != '1B8C8F'
                           or not str(r[4] or '').startswith('Aptos') for r in rs):
            bad33.append((i, _PAGENO))
    results.append(("-1：" if bad33 else "-0：")
                   + "四张示意图片与页面装饰元素不满足各自坐标与尺寸——图片（第1页左端17.4cm 顶端0cm 宽16.46cm 高19.05cm、第2页左端27.05cm 顶端12.57cm 宽4.57cm 高3.81cm、第10页左端21.34cm 顶端4.45cm 宽10.54cm 高5.21cm、第13页左端19.18cm 顶端4.19cm 宽12.57cm 高4.45cm）、每页左上角色块（左端0.46cm 顶端0.46cm 宽4.32cm 高2.92cm）、右下角色块（左端28.83cm 顶端15.37cm 宽3.94cm 高2.67cm）、左侧竖条（左端0cm 顶端0cm 宽0.36cm 高19.05cm）、右上角页码标签（左端30.73cm 顶端1.07cm、Aptos 17磅 加粗 颜色#1B8C8F）")

    # ========== 将 results 反解为结构化 dim2_items ==========
    # 逐项按追加顺序对齐 _MAX_DELTAS_ORDER：解析每条 "+N：说明" / "-N：说明" 的 N 与 rule。
    if len(results) != len(_MAX_DELTAS_ORDER):
        # 顺序表与实际条目不一致：视为脚本内部错误，交给外层兜底
        raise RuntimeError(
            f"dim2 项数不一致：results={len(results)} vs order={len(_MAX_DELTAS_ORDER)}"
        )
    dim2_items = []
    total_score = 0
    for raw, max_delta in zip(results, _MAX_DELTAS_ORDER):
        # 格式约定：加分项 '+N：' 前缀，命中 N=max_delta、未命中 N=0；
        # 减分项 '-N：' 前缀，命中 N 为实际扣分（分级条目可小于 max_delta）、未命中 N=0。
        head, _, rule = raw.partition("：")
        try:
            delta = int(head.replace("+", "").strip())
        except ValueError:
            delta = 0
        if max_delta >= 0:
            hit = delta > 0
            delta = max_delta if hit else 0
        else:
            hit = delta < 0
            delta = max(max_delta, delta) if hit else 0
        total_score += delta
        if max_delta and delta not in (0, max_delta):
            # 分级扣分项可能只触发部分额度，但统一结果契约要求每个条目的 delta
            # 只能是 0 或 max_delta，因此拆成 1 分步长的原子计分单元。
            step = 1 if max_delta > 0 else -1
            units = abs(max_delta)
            gained = min(abs(delta), units)
            for unit in range(units):
                hit_unit = unit < gained
                dim2_items.append({
                    "rule": "%s（计分单元 %d/%d）" % (rule, unit + 1, units),
                    "max_delta": step,
                    "delta": step if hit_unit else 0,
                    "hit": hit_unit,
                })
            continue
        dim2_items.append({
            "rule": rule, "max_delta": max_delta,
            "delta": delta, "hit": hit,
        })

    result["dim2_items"] = dim2_items
    result["total_score"] = max(0, total_score)
    result["max_score"] = MAX_SCORE_TOTAL
    return result


if __name__ == "__main__":
    # 本地调试入口：evaluate 只接收“脚本所在目录的路径”，默认取本文件所在目录；
    # 也可通过命令行传入待评估 .pptx 所在目录。
    if len(sys.argv) >= 2:
        _dir = sys.argv[1]
    else:
        _dir = os.path.dirname(os.path.abspath(__file__))
    _out = evaluate(_dir)
    print(json.dumps(_out, ensure_ascii=False, indent=2))
