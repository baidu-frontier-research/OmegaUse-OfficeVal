#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# pyright: reportUnknownMemberType=false, reportUnknownVariableType=false, reportUnknownArgumentType=false, reportUnknownParameterType=false, reportMissingParameterType=false, reportPrivateUsage=false, reportUnknownLambdaType=false, reportGeneralTypeIssues=false

import os
import re
import sys
import json
import math
import shutil
import zipfile
import tempfile
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.text.run import Run
from docx.oxml.ns import qn

# 说明：按接口统一约定，本模块不修改全局 sys.stdout，也不 print 主结果。
# 主结果统一由 evaluate(dir_path) -> dict 返回，入参为"脚本所在目录的路径"，
# 由脚本自己在该目录内定位并打开被评估的 .docx。本文件仅在 __main__ 下作为
# 本地调试入口，把这个 dict 序列化为 JSON 打印。

SCRIPT_ID = "037"
# 维度二所有评分项 max_delta 之和，用于维度一未通过 / 脚本异常时兜底
MAX_SCORE_TOTAL = 100

# 维度二所有评分项按"结果追加顺序"的 max_delta 列表，供末端把 results 里的
# "+N：..." 字符串拆解为结构化 dim2_items 时使用。顺序须与 _evaluate_impl 中
# results.append(...) 的顺序严格一致，共 42 项，累加 = MAX_SCORE_TOTAL。
_MAX_DELTAS_ORDER = [
    # ---- 内容 10 条 = 30 分 ----
    1,  # 01 卷首四项
    1,  # 02 三个部分标题
    1,  # 03 四个语篇标识 A-D
    3,  # 04 六个节标题括注
    5,  # 05 题号 21-35 / 41-55 共 30
    5,  # 06 四选一 ABCD 齐全
    1,  # 07 七选五 A-G 共 7 项独立成段
    5,  # 08 空白编号 21 处
    3,  # 09 读后续写材料 5 段 + 2 开头语
    5,  # 10 表格 5x5 / 20 格非空 / 表头
    # ---- 格式 1 分 18 条 = 18 分 ----
    1,  # 11 A4 纵向 单栏
    1,  # 12 页边距 1.27 厘米
    1,  # 13 页眉无文字/线条/图形
    1,  # 14 段前段后间距均为 0
    1,  # 15 语篇标识 12 磅加粗居中
    1,  # 16 五处答题指导语 12 磅加粗
    1,  # 17 七选五首行缩进 2 字符
    1,  # 18 引文/出处各自成段且水平居中
    1,  # 19 正文字号 12 磅
    1,  # 20 卷首三段 14 磅加粗居中
    1,  # 21 英文语句无连续空格
    1,  # 22 选项标号后 1 个空格
    1,  # 23 无批注/修订，页面上无文本框图片艺术字形状
    1,  # 24 写作要求条目 7 处各自成段 + 与注意段缩进 2 字符
    1,  # 25 无高亮 / 数据行 20 格无底纹
    1,  # 26 无连续空段 / 无手动分页符
    1,  # 27 表格框线单实线 0.5 磅 中灰 777777
    1,  # 28 表格无合并 / 每行 5 格
    # ---- 格式 3 分 9 条 = 27 分 ----
    3,  # 29 页脚页码段居中 + 页码/总页数为域
    3,  # 30 中文宋体 / 西文数字 Times New Roman
    3,  # 31 除卷首外行距固定 20 磅
    3,  # 32 三处下划线作答区各不少于 1 整行
    3,  # 33 应用文与续写开头语加粗
    3,  # 34 表格居中/表头加粗浅灰底纹/垂直居中/12 磅
    3,  # 35 题号非自动编号
    3,  # 36 无单词内部被空格拆开
    3,  # 37 完形每行四选项分列且有间隔
    # ---- 格式 5 分 5 条 = 25 分 ----
    5,  # 38 语篇正文与材料段首行缩进 2 字符
    5,  # 39 部分/节标题 12 磅加粗且顶格
    5,  # 40 阅读选项段首行缩进 2 字符
    5,  # 41 同题两行式 A/C、B/D 逐列对齐
    5,  # 42 完形选项段首行缩进 0.74 厘米
]
assert sum(_MAX_DELTAS_ORDER) == MAX_SCORE_TOTAL

# ======================================================================
# 附录：固定字串校验清单（全部去掉空白后比较，便于跨排版复现）
# 附录只用于"内容是否搬全"，不涉及任何格式判定。
# ======================================================================
A1_COVER = ["绝密★启用前",
            "澜桥市2026年高三年级第二学期质量检测",
            "英语试题",
            "2026年6月"]

A2_PART = ["第二部分阅读理解",
           "第三部分语言知识运用",
           "第四部分写作"]

A2_SECTION = ["第一节（共15小题；每小题2.5分，满分37.5分）",
              "第二节（共5小题；每小题2.5分，满分12.5分）",
              "第一节（共15小题；每小题1分，满分15分）",
              "第二节（共10小题；每小题1.5分，满分15分）",
              "第一节（满分15分）",
              "第二节（满分25分）"]

# 读后续写 5 个自然段的段首字串（去空白后取前 18 字符做前缀比对）
A3_STORY = ["It was the week before the final mock examination",
            "On Friday afternoon, our teacher announced",
            "Still, I took a blank card",
            "After school, I was asked to deliver the cards",
            "I thanked him and noticed the notebook in his hand"]

A4_TABLE_HEADER = ["Programme", "Date (2026)", "Time", "Fee (credits)", "Main focus"]

A5_QUOTE = "A goal without action is only a wish; action without a goal easily loses its way."
A5_SOURCE = "- M. Ellison"

# 读后续写两段开头语（加粗，各自独立成段；去空白后取前 40 字符做前缀比对）
A7_OPENING = ["Paragraph 1: I sat with him on the bench near the stairway and opened his notebook.",
              "Paragraph 2: When I returned to my own classroom on Monday, I found a small card on my desk."]

# ======================================================================
# 通用工具
# ======================================================================
def _nz(s):
    """去掉全部空白字符，用于内容字串比对（排版换行/空格差异不算内容差异）。"""
    return re.sub(r"\s+", "", s or "")


def _paragraph_alignment_name(para):
    """读取 w:jc 原始值，兼容 python-docx 未映射的 start/end 值。"""
    ppr = para._p.find(qn("w:pPr"))
    jc = ppr.find(qn("w:jc")) if ppr is not None else None
    value = jc.get(qn("w:val")) if jc is not None else None
    return {
        "left": "LEFT",
        "start": "LEFT",
        "center": "CENTER",
        "right": "RIGHT",
        "end": "RIGHT",
        "both": "JUSTIFY",
        "distribute": "DISTRIBUTE",
    }.get(value, None)


def _fnum(x):
    """把 '20.0' / None / 'None' 统一转成 float 或 None。"""
    try:
        if x is None:
            return None
        return float(x)
    except (TypeError, ValueError):
        return None


def _resolve_effective_eastasia(doc, run):
    """解析 run 实际生效的中文（eastAsia）字体名：
    run 级 w:rFonts/@w:eastAsia -> 段落样式链（含 w:basedOn 逐级上溯）
    -> w:docDefaults/w:rPrDefault。返回 None 表示各级都没写。
    注意：只看 <w:r><w:rPr>，不看 <w:pPr><w:rPr>（后者是段落标记 ¶ 自身属性）。
    """
    rpr = run._element.find(qn("w:rPr"))
    if rpr is not None:
        rf = rpr.find(qn("w:rFonts"))
        if rf is not None:
            v = rf.get(qn("w:eastAsia"))
            if v:
                return v
    style = getattr(run, "style", None)
    seen = set()
    while style is not None and style.style_id not in seen:
        seen.add(style.style_id)
        el = style.element
        rf = el.find(qn("w:rPr") + "/" + qn("w:rFonts")) if el is not None else None
        if rf is None and el is not None:
            r2 = el.find(qn("w:rPr"))
            rf = r2.find(qn("w:rFonts")) if r2 is not None else None
        if rf is not None:
            v = rf.get(qn("w:eastAsia"))
            if v:
                return v
        style = getattr(style, "base_style", None)
    v = _parastyle_font(doc, run, ("eastAsia",))
    if v:
        return v
    return _docdefault_font(doc, "eastAsia")

def _resolve_effective_ascii(doc, run):
    """解析 run 实际生效的西文（ascii）字体名，链路与 eastAsia 相同。
    ascii 缺省时回退 hAnsi。
    """
    rpr = run._element.find(qn("w:rPr"))
    if rpr is not None:
        rf = rpr.find(qn("w:rFonts"))
        if rf is not None:
            v = rf.get(qn("w:ascii")) or rf.get(qn("w:hAnsi"))
            if v:
                return v
    style = getattr(run, "style", None)
    seen = set()
    while style is not None and style.style_id not in seen:
        seen.add(style.style_id)
        el = style.element
        r2 = el.find(qn("w:rPr")) if el is not None else None
        rf = r2.find(qn("w:rFonts")) if r2 is not None else None
        if rf is not None:
            v = rf.get(qn("w:ascii")) or rf.get(qn("w:hAnsi"))
            if v:
                return v
        style = getattr(style, "base_style", None)
    v = _parastyle_font(doc, run, ("ascii", "hAnsi"))
    if v:
        return v
    return _docdefault_font(doc, "ascii") or _docdefault_font(doc, "hAnsi")


def _docdefault_font(doc, attr):
    """读取 w:docDefaults/w:rPrDefault/w:rPr/w:rFonts 上的指定字体属性。"""
    try:
        el = doc.styles.element
        rf = el.find(qn("w:docDefaults") + "/" + qn("w:rPrDefault")
                     + "/" + qn("w:rPr") + "/" + qn("w:rFonts"))
        if rf is not None:
            return rf.get(qn("w:" + attr))
    except Exception:
        pass
    return None


def _open_word_document(file_path):
    """打开被评估的 .docx。打不开即视为维度一不通过，由上层捕获。"""
    return Document(file_path)

# ======================================================================
# 版面推算（替代渲染取证）
#
# 位置类判定（同一行分列、上下两行逐列对齐、居中）改为按 OOXML 静态推算：
#   行首 x   = 左缩进 + 首行缩进 − 悬挂缩进（w:ind 的 *Chars 属性优先于缇值，
#              与"段落"对话框显示的"字符"单位一致）
#   文字推进 = 逐字符累加字体推进宽度
#   w:tab    = 跳到下一个制表位：w:pPr/w:tabs 的显式左制表位优先，用尽后按
#              w:defaultTabStop 的等距默认制表位
# 字符宽度优先用系统字体文件度量（Pillow 读 times.ttf / simsun.ttc），取不到
# 字体文件时退回按字号折算的平均宽度（全角 1 em、空格 0.25 em、其余 0.5 em）。
# ======================================================================
_HAN_RE = re.compile(r"[\u3400-\u4dbf\u4e00-\u9fff\uf900-\ufaff]")
_WIDE_RE = re.compile(r"[\u3000-\u303f\u3400-\u4dbf\u4e00-\u9fff\uf900-\ufaff"
                      r"\uff01-\uff60]")
SONGTI_NAMES = ("宋体", "SimSun", "NSimSun", "STSong")
TNR_NAMES = ("Times New Roman", "TimesNewRoman", "Times New Roman PS")

_FONT_DIR = os.path.join(os.environ.get("WINDIR", "C:\\Windows"), "Fonts")
_FONT_EM = 1000
_FONT_OBJ = {}
_CHAR_W = {}


def _sys_font(name):
    if name in _FONT_OBJ:
        return _FONT_OBJ[name]
    obj = None
    try:
        from PIL import ImageFont
        obj = ImageFont.truetype(os.path.join(_FONT_DIR, name), _FONT_EM)
    except Exception:
        obj = None
    _FONT_OBJ[name] = obj
    return obj


def _char_width_pt(ch, size_pt):
    """单个字符在给定字号下的推进宽度（磅）。"""
    key = (ch, round(size_pt, 2))
    if key in _CHAR_W:
        return _CHAR_W[key]
    wide = bool(_WIDE_RE.match(ch))
    f = _sys_font("simsun.ttc" if wide else "times.ttf")
    w = None
    if f is not None:
        try:
            w = float(f.getlength(ch)) * size_pt / _FONT_EM
        except Exception:
            w = None
    if w is None:
        if wide:
            w = size_pt
        elif ch == " ":
            w = size_pt * 0.25
        else:
            w = size_pt * 0.5
    _CHAR_W[key] = w
    return w


def _para_runs_in_order(para):
    """按出现顺序返回段落内的 run（含 w:hyperlink / w:ins / w:sdt 内的 run）。"""
    wrappers = (qn("w:hyperlink"), qn("w:ins"), qn("w:smartTag"), qn("w:sdt"))
    out = []
    for el in para._element.iterchildren():
        if el.tag == qn("w:r"):
            out.append(Run(el, para))
        elif el.tag in wrappers:
            for sub in el.iter(qn("w:r")):
                out.append(Run(sub, para))
    return out


def _iter_all_runs(doc):
    """正文段落与表格单元格内段落的全部 run。"""
    for p in doc.paragraphs:
        for r in _para_runs_in_order(p):
            yield r
    for t in doc.tables:
        for row in t.rows:
            for c in row.cells:
                for p in c.paragraphs:
                    for r in _para_runs_in_order(p):
                        yield r


def _font_in(name, allowed):
    if not name:
        return False
    n = str(name).strip()
    return any(n == a or n.lower() == a.lower() for a in allowed)


def _parastyle_font(doc, run, attrs):
    """段落样式链上的 w:rPr/w:rFonts：run 与字符样式都没写时由它决定。"""
    para = getattr(run, "_parent", None)
    st = getattr(para, "style", None) if para is not None else None
    seen = set()
    while st is not None and st.style_id not in seen:
        seen.add(st.style_id)
        rp = st.element.find(qn("w:rPr")) if st.element is not None else None
        rf = rp.find(qn("w:rFonts")) if rp is not None else None
        if rf is not None:
            for a in attrs:
                v = rf.get(qn("w:" + a))
                if v:
                    return v
        st = getattr(st, "base_style", None)
    return None


def _eff_ppr_el(doc, para, child):
    """段落 pPr 的子元素；段落没写时沿 w:pStyle 继承链上溯。"""
    pPr = para._element.find(qn("w:pPr"))
    if pPr is not None:
        el = pPr.find(qn("w:" + child))
        if el is not None:
            return el
    st = getattr(para, "style", None)
    seen = set()
    while st is not None and st.style_id not in seen:
        seen.add(st.style_id)
        spr = st.element.find(qn("w:pPr")) if st.element is not None else None
        if spr is not None:
            el = spr.find(qn("w:" + child))
            if el is not None:
                return el
        st = getattr(st, "base_style", None)
    return None


def _eff_jc(doc, para):
    """段落有效对齐方式（含样式继承链）。"""
    el = _eff_ppr_el(doc, para, "jc")
    return el.get(qn("w:val")) if el is not None else None


def _run_size_pt(doc, run):
    """run 实际生效的字号（磅）：run 级 w:sz → 段落样式链 → docDefaults → 12。"""
    rpr = run._element.find(qn("w:rPr"))
    if rpr is not None:
        sz = rpr.find(qn("w:sz"))
        v = _fnum(sz.get(qn("w:val"))) if sz is not None else None
        if v:
            return v / 2.0
    para = getattr(run, "_parent", None)
    st = getattr(para, "style", None) if para is not None else None
    seen = set()
    while st is not None and st.style_id not in seen:
        seen.add(st.style_id)
        rp = st.element.find(qn("w:rPr")) if st.element is not None else None
        sz = rp.find(qn("w:sz")) if rp is not None else None
        v = _fnum(sz.get(qn("w:val"))) if sz is not None else None
        if v:
            return v / 2.0
        st = getattr(st, "base_style", None)
    try:
        sz = doc.styles.element.find(
            qn("w:docDefaults") + "/" + qn("w:rPrDefault") + "/"
            + qn("w:rPr") + "/" + qn("w:sz"))
        v = _fnum(sz.get(qn("w:val"))) if sz is not None else None
        if v:
            return v / 2.0
    except Exception:
        pass
    return 12.0


def _run_bold(doc, run):
    """run 是否加粗（run 级 w:b → 段落样式链）。"""
    def _on(el):
        v = el.get(qn("w:val"))
        return v is None or str(v).lower() not in ("0", "false", "off")
    rpr = run._element.find(qn("w:rPr"))
    b = rpr.find(qn("w:b")) if rpr is not None else None
    if b is not None:
        return _on(b)
    para = getattr(run, "_parent", None)
    st = getattr(para, "style", None) if para is not None else None
    seen = set()
    while st is not None and st.style_id not in seen:
        seen.add(st.style_id)
        rp = st.element.find(qn("w:rPr")) if st.element is not None else None
        b = rp.find(qn("w:b")) if rp is not None else None
        if b is not None:
            return _on(b)
        st = getattr(st, "base_style", None)
    return False


def _default_tab_pt(doc):
    """w:defaultTabStop（缇）→ 磅；未声明时按中文版默认的 0.74 厘米。"""
    try:
        el = doc.settings.element.find(qn("w:defaultTabStop"))
        v = _fnum(el.get(qn("w:val"))) if el is not None else None
        if v:
            return v / 20.0
    except Exception:
        pass
    return 21.0


def _ind_metrics_pt(doc, para, size_pt):
    """返回 (左缩进, 首行缩进, 悬挂缩进, 右缩进)，单位磅。

    w:*Chars（百分之一字符）优先于缇值——中文文档对话框里显示的就是"字符"。
    """
    ind = _eff_ppr_el(doc, para, "ind")

    def pick(chars_attr, twip_attr):
        if ind is None:
            return 0.0
        v = ind.get(qn("w:" + chars_attr))
        if v is not None:
            n = _fnum(v)
            return (n / 100.0 * size_pt) if n is not None else 0.0
        v = ind.get(qn("w:" + twip_attr))
        if v is not None:
            n = _fnum(v)
            return (n / 20.0) if n is not None else 0.0
        return 0.0
    return (pick("leftChars", "left"), pick("firstLineChars", "firstLine"),
            pick("hangingChars", "hanging"), pick("rightChars", "right"))


def _ind_centered_ok(doc, para):
    """居中段落不应带首行/悬挂缩进，左右缩进也要对称，否则文字不落在页面中线上。"""
    runs = _para_runs_in_order(para)
    size = _run_size_pt(doc, runs[0]) if runs else 12.0
    left, first, hang, right = _ind_metrics_pt(doc, para, size)
    return abs(first) < 0.5 and abs(hang) < 0.5 and abs(left - right) < 1.0


def _tab_stops_pt(doc, para):
    """返回 (显式左制表位磅值升序列表, 是否存在非左对齐制表位)。"""
    tabs = _eff_ppr_el(doc, para, "tabs")
    stops = []
    nonleft = False
    if tabs is not None:
        for t in tabs.findall(qn("w:tab")):
            val = t.get(qn("w:val"))
            if val == "clear":
                continue
            if val not in (None, "left", "start", "num"):
                nonleft = True
                continue
            n = _fnum(t.get(qn("w:pos")))
            if n is not None:
                stops.append(n / 20.0)
    return (sorted(stops), nonleft)


def _next_tab_pt(x, stops, dts):
    """下一个制表位落点：显式制表位优先，用尽后落到默认制表位网格上。"""
    for p in stops:
        if p > x + 0.01:
            return p
    last = stops[-1] if stops else 0.0
    k = int(max(x, last) // dts) + 1
    p = k * dts
    while p <= x + 0.01 or p <= last + 0.01:
        k += 1
        p = k * dts
    return p


def _line_layout(doc, para, dts):
    """返回 (flat, xs, ws, end_x, nonleft)。

    flat 为段落内容（制表符记 \t、换行记 \n），xs[i] 是 flat[i] 的起点 x（磅），
    ws[i] 是它的推进宽度，end_x 是整段末端 x。
    """
    stops, nonleft = _tab_stops_pt(doc, para)
    runs = _para_runs_in_order(para)
    size0 = _run_size_pt(doc, runs[0]) if runs else 12.0
    left, first, hang, _right = _ind_metrics_pt(doc, para, size0)
    x = left + first - hang
    flat, xs, ws = [], [], []
    for r in runs:
        size = _run_size_pt(doc, r)
        for el in r._element.iter():
            if el.tag == qn("w:t"):
                for ch in (el.text or ""):
                    w = _char_width_pt(ch, size)
                    flat.append(ch)
                    xs.append(x)
                    ws.append(w)
                    x += w
            elif el.tag == qn("w:tab"):
                nx = _next_tab_pt(x, stops, dts)
                flat.append("\t")
                xs.append(x)
                ws.append(nx - x)
                x = nx
            elif el.tag == qn("w:br"):
                flat.append("\n")
                xs.append(x)
                ws.append(0.0)
                x = left
    return ("".join(flat), xs, ws, x, nonleft)


def _option_columns(doc, para, dts, labels):
    """按 labels 顺序推算各选项列的起点与该列文字末端 x。

    返回 (cols, flat, end_x, nonleft)；labels 未按序齐全时 cols 为空字典。
    """
    flat, xs, ws, end_x, nonleft = _line_layout(doc, para, dts)
    if not xs:
        return ({}, flat, end_x, nonleft)
    pos = {}
    idx = 0
    for lb in labels:
        k = flat.find(lb, idx)
        if k < 0:
            return ({}, flat, end_x, nonleft)
        pos[lb] = k
        idx = k + len(lb)
    cols = {}
    order = list(labels)
    for n, lb in enumerate(order):
        s = pos[lb]
        stop = pos[order[n + 1]] if n + 1 < len(order) else len(flat)
        seg = flat[s:stop].rstrip()
        e = s + max(len(seg) - 1, 0)
        cols[lb] = (xs[s], xs[e] + ws[e])
    return (cols, flat, end_x, nonleft)


# ---------------------------------------------------------------- 底纹与框线
def _visible_fill(v):
    """w:shd/@w:fill 是否是"看得见的底纹"。"""
    if not v:
        return None
    s = str(v).strip().upper()
    if s in ("AUTO", "FFFFFF", "NIL", "NONE", ""):
        return None
    return s


def _cell_has_shading(cell):
    for shd in cell._tc.iter(qn("w:shd")):
        if _visible_fill(shd.get(qn("w:fill"))):
            return True
    return False


def _cell_light_grey(cell):
    """单元格底纹是否为浅灰（三通道接近相等且落在 D9~FA 之间）。"""
    tcPr = cell._tc.find(qn("w:tcPr"))
    shd = tcPr.find(qn("w:shd")) if tcPr is not None else None
    v = _visible_fill(shd.get(qn("w:fill")) if shd is not None else None)
    if not v or len(v) != 6:
        return False
    try:
        r, g, b = int(v[0:2], 16), int(v[2:4], 16), int(v[4:6], 16)
    except ValueError:
        return False
    return (max(r, g, b) - min(r, g, b) <= 8
            and 0xD9 <= (r + g + b) // 3 <= 0xFA)


def _cell_all_bold(doc, cell):
    got = False
    for p in cell.paragraphs:
        for r in _para_runs_in_order(p):
            if not (r.text or "").strip():
                continue
            got = True
            if not _run_bold(doc, r):
                return False
    return got


def _cell_sizes(doc, cell):
    out = set()
    for p in cell.paragraphs:
        for r in _para_runs_in_order(p):
            if (r.text or "").strip():
                out.add(round(_run_size_pt(doc, r), 1))
    return out


def _shaded_data_cells(tbl):
    """表格除表头行以外带底纹的单元格坐标列表。"""
    bad = []
    if tbl is None:
        return bad
    for ri, row in enumerate(list(tbl.rows)[1:], start=1):
        for ci, c in enumerate(row.cells):
            if _cell_has_shading(c):
                bad.append((ri, ci))
    return bad


# ------------------------------------------------------------ 图形与图片残留
# 小于 0.1 厘米（约 3 磅）的尺寸声明属于隐形残留，不算"页面上看得见"
_SHAPE_MIN_EMU = 38100
_SHAPE_MIN_PT = 3.0


def _vml_size_pt(style):
    def g(key):
        m = re.search(r"\b%s\s*:\s*(-?[\d.]+)\s*(pt|px|in|cm|mm)?" % key, style or "")
        if not m:
            return None
        v = float(m.group(1))
        u = (m.group(2) or "pt").lower()
        return {"pt": v, "px": v * 0.75, "in": v * 72.0,
                "cm": v * 28.3465, "mm": v * 2.83465}.get(u, v)
    return (g("width"), g("height"))


def _visible_shape_count(xml):
    """XML 里"看得见的"图形对象个数（按尺寸阈值过滤隐形残留）。"""
    n = 0
    for m in re.finditer(r"<wp:extent\b[^>]*?>", xml):
        cx = re.search(r'\bcx="(\d+)"', m.group(0))
        cy = re.search(r'\bcy="(\d+)"', m.group(0))
        if (cx and cy and int(cx.group(1)) >= _SHAPE_MIN_EMU
                and int(cy.group(1)) >= _SHAPE_MIN_EMU):
            n += 1
    for m in re.finditer(r"<v:(?:shape|rect|oval|line|group|image)\b[^>]*?>", xml):
        tag = m.group(0)
        st = re.search(r'style="([^"]*)"', tag)
        w_, h_ = _vml_size_pt(st.group(1) if st else "")
        if ((w_ is not None and w_ < _SHAPE_MIN_PT)
                or (h_ is not None and h_ < _SHAPE_MIN_PT)):
            continue
        fill = (re.search(r'fillcolor="([^"]*)"', tag) or [None, ""])[1]
        filled = (re.search(r'\bfilled="([^"]*)"', tag) or [None, "t"])[1]
        stroked = (re.search(r'\bstroked="([^"]*)"', tag) or [None, "t"])[1]
        blank = (str(stroked).lower() in ("f", "false", "0")
                 and (str(filled).lower() in ("f", "false", "0")
                      or str(fill).upper().replace("#", "") == "FFFFFF"))
        if blank:
            continue  # 白色填充且无描边：页面上看不见
        n += 1
    n += len(re.findall(r"<w:object\b", xml))
    return n


def _image_rel_count(file_path):
    """word/_rels/document.xml.rels 里 image 关系的条数。

    正文里出现图片的必要条件；页眉页脚里的 0.05 磅残留图片不在本条范围内，
    是否可见由 _visible_shape_count 的尺寸阈值决定。
    """
    try:
        with zipfile.ZipFile(file_path) as z:
            txt = z.read("word/_rels/document.xml.rels").decode("utf-8", "ignore")
    except Exception:
        return 0
    return len(re.findall(r'Type="[^"]*/image"', txt))


def _read_parts(file_path):
    """读取 .docx 内部 XML 部件的原文，键为包内路径。"""
    parts = {}
    with zipfile.ZipFile(file_path) as z:
        for name in z.namelist():
            if name.startswith("word/") and name.endswith(".xml"):
                try:
                    parts[name] = z.read(name).decode("utf-8", "ignore")
                except Exception:
                    parts[name] = ""
    return parts


def _para_info(doc, para, order):
    """把一个段落摊平成判定用的字典。

    与办公软件对齐的读取路径：
      al   段落对齐        = 开始→段落→对齐方式
      flc  w:ind/@firstLineChars = 首行缩进（字符）——中文文档对话框显示的就是"字符"，
           w:firstLine 的磅值只是等价缓存，firstLineChars 才是生效值
      fli  w:ind/@firstLine (缇/20)  = 首行缩进换算成磅，用于"厘米"型缩进
      li   w:ind/@left     (缇/20)  = 左缩进磅值
      ls/lr w:spacing/@line + @lineRule = 行距（exact + 400 缇 → 固定值 20 磅）
      sb/sa w:spacing/@before + @after  = 段前 / 段后
      num  w:numPr/w:numId = 自动编号；非 0 即为自动编号
      runs 只取 <w:r><w:rPr>，不取 <w:pPr><w:rPr>（后者是 ¶ 自身属性，不管正文字符）
    """
    el = para._element
    pPr = el.find(qn("w:pPr"))
    ind = pPr.find(qn("w:ind")) if pPr is not None else None
    spc = pPr.find(qn("w:spacing")) if pPr is not None else None
    numpr = pPr.find(qn("w:numPr")) if pPr is not None else None
    pbb = pPr.find(qn("w:pageBreakBefore")) if pPr is not None else None

    def a(node, key):
        return node.get(qn("w:" + key)) if node is not None else None

    numid = None
    if numpr is not None:
        nid = numpr.find(qn("w:numId"))
        numid = a(nid, "val")
    pbb_on = False
    if pbb is not None:
        pbb_on = a(pbb, "val") not in ("0", "false", "none")

    runs = []
    for r in para.runs:
        txt = r.text
        rpr = r._element.find(qn("w:rPr"))
        szv = None
        if rpr is not None:
            s = rpr.find(qn("w:sz"))
            if s is not None and a(s, "val"):
                szv = _fnum(a(s, "val"))
                szv = szv / 2.0 if szv is not None else None
        if szv is None:
            szv = r.font.size.pt if r.font.size is not None else None
        runs.append(dict(text=txt,
                         ascii=_resolve_effective_ascii(doc, r),
                         ea=_resolve_effective_eastasia(doc, r),
                         size=(round(szv, 1) if szv is not None else None),
                         bold=bool(r.bold or (r.font.bold is True)),
                         italic=bool(r.italic),
                         underline=bool(r.underline)))

    fli = _fnum(a(ind, "firstLine"))
    li = _fnum(a(ind, "left"))
    line = _fnum(a(spc, "line"))
    return dict(n=order, ord=order, t=para.text,
                al=_paragraph_alignment_name(para),
                flc=a(ind, "firstLineChars"), lic=a(ind, "leftChars"),
                fli=(round(fli / 20.0, 2) if fli is not None else None),
                li=(round(li / 20.0, 2) if li is not None else None),
                sb=_fnum(a(spc, "before")), sa=_fnum(a(spc, "after")),
                ls=(round(line / 20.0, 1) if line is not None else None),
                lr=a(spc, "lineRule"), num=numid, pbb=pbb_on, runs=runs)


def _has_run(p, size=None, bold=None):
    """段落里是否存在满足条件的非空 run（字号 / 加粗）。"""
    for r in p["runs"]:
        if not r["text"].strip():
            continue
        if size is not None and r["size"] != size:
            continue
        if bold is not None and bool(r["bold"]) != bold:
            continue
        return True
    return False

# ======================================================================
# 对外唯一入口
# ======================================================================
def evaluate(dir_path):
    """入参为脚本所在目录，自行在该目录内定位待评估的 .docx，返回结果 dict。"""
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
            raise FileNotFoundError("目录不存在：%s" % dir_path)
        cands = sorted(f for f in os.listdir(dir_path)
                       if f.lower().endswith(".docx") and not f.startswith("~"))
        if not cands:
            raise FileNotFoundError("目录内未找到 .docx 文件：%s" % dir_path)
        file_name = cands[0]
        result["file_name"] = file_name
        file_path = os.path.join(dir_path, file_name)
        return _evaluate_impl(file_path, result)
    except Exception as exc:  # noqa: BLE001
        # status 为 error 时未能完成判定，dim1_pass 留空，dim1_reason 也为空
        result["status"] = "error"
        result["error"] = "%s: %s" % (type(exc).__name__, exc)
        result["dim1_pass"] = ""
        result["dim1_reason"] = ""
        result["dim2_items"] = []
        result["total_score"] = 0
        result["max_score"] = MAX_SCORE_TOTAL
        return result

def _evaluate_impl(file_path, result):
    tmp_dir = tempfile.mkdtemp(prefix="officeval%s_" % SCRIPT_ID)
    try:
        # ---------- 维度一：可用与可修改性（门槛项，不通过则完成度直接 0） ----------
        # 交付件损坏（打不开、包结构坏）按维度一不通过处理，不记为脚本异常。
        try:
            doc = _open_word_document(file_path)
            parts = _read_parts(file_path)
        except Exception as exc:  # noqa: BLE001
            result["dim1_pass"] = False
            result["dim1_reason"] = "文档无法打开或包结构损坏：%s: %s" % (
                type(exc).__name__, exc)
            result["total_score"] = 0
            result["max_score"] = MAX_SCORE_TOTAL
            return result
        D = parts.get("word/document.xml", "")
        if not D:
            result["dim1_pass"] = False
            result["dim1_reason"] = "word/document.xml 缺失或不可解析，文件无法正常打开编辑"
            result["total_score"] = 0
            result["max_score"] = MAX_SCORE_TOTAL
            return result
        P = [_para_info(doc, p, i) for i, p in enumerate(doc.paragraphs)]
        if not P:
            result["dim1_pass"] = False
            result["dim1_reason"] = "正文无任何段落"
            result["total_score"] = 0
            result["max_score"] = MAX_SCORE_TOTAL
            return result
        full = "\n".join(p["t"] for p in P)

        # ---------- 版面几何（不再导出 PDF，全部按 OOXML 推算） ----------
        PARAS = list(doc.paragraphs)          # 与 P 逐项对应，便于取回段落元素
        DTS = _default_tab_pt(doc)            # 默认制表位间距（磅）
        sect = doc.sections[0]
        try:
            PGW = sect.page_width.pt
        except Exception:
            PGW = 595.3                       # A4 = 595.3 x 841.9 磅
        try:
            TXTW = (sect.page_width.pt - sect.left_margin.pt - sect.right_margin.pt)
        except Exception:
            TXTW = PGW - 72.0

        # ---------- 段落分类（各评分项共用） ----------
        MARK = [p for p in P if re.fullmatch(r"[A-D]", p["t"].strip())]
        GUIDE = [p for p in P if re.match(r"^\s*(?:\d{1,2}\.\s*)?阅读[下]", p["t"].strip())]
        HD = [p for p in P if re.match(r"^第[二三四]部分", p["t"].strip())
              or re.match(r"^第[一二]节", p["t"].strip())]
        CLOZE = [p for p in P if re.match(r"^\s*(4[1-9]|5[0-5])\.\s*A\.", p["t"])]
        CLO = set(p["ord"] for p in CLOZE)
        s7 = next((p["ord"] for p in P if _nz(p["t"]).startswith("第二节（共5小题")), 0)
        e7 = next((p["ord"] for p in P if re.match(r"^第三部分", p["t"].strip())), len(P))
        SEV = [p for p in P[s7:e7] if re.match(r"^\s*[A-G]\.\s", p["t"]) and len(p["t"]) > 25]
        SEVO = set(p["ord"] for p in SEV)
        READOPT = [p for p in P if re.match(r"^\s*[A-D]\.\s", p["t"])
                   and p["ord"] not in CLO and p["ord"] not in SEVO]
        PROSE = [p for p in P if re.match(r"^[A-Z\"“(•–-]", p["t"].strip())
                 and len(p["t"]) > 60 and not re.match(r"^\s*[A-G]\.\s", p["t"])
                 and not p["t"].strip().startswith("Paragraph")
                 and p["ord"] not in CLO and "goal without action" not in p["t"]
                 and "Learning with a clear purpose" not in p["t"]]
        TB = doc.tables[0] if doc.tables else None

        results = []

        # ==============================================================
        # 内容 10 条 = 30 分
        # ==============================================================
        # 01 卷首四项：正文最前面 4 个非空段落逐一等于附录 A.1（去空白后比较）
        head = [_nz(p["t"]) for p in P[:6] if p["t"].strip()][:4]
        ok = (head == A1_COVER)
        results.append("+%d：卷首依次出现「绝密★启用前」「澜桥市 2026 年高三年级第二学期质量检测」"
                       "「英语试题」「2026 年 6 月」四项，各自独立成段" % (1 if ok else 0))

        # 02 三个部分标题：附录 A.2 的三条部分标题各自出现在某个"第X部分"开头的段落里
        parts_t = [_nz(p["t"]) for p in P if re.match(r"^第[二三四]部分", p["t"].strip())]
        ok = all(any(a in x for x in parts_t) for a in A2_PART)
        results.append("+%d：出现「第二部分 阅读理解」「第三部分 语言知识运用」「第四部分 写作」三个部分标题"
                       % (1 if ok else 0))

        # 03 四个语篇标识：独立成段的单字母段落恰为 A、B、C、D 四个且顺序正确
        labels = [p["t"].strip() for p in P if re.fullmatch(r"[A-D]", p["t"].strip())]
        ok = (labels[:4] == ["A", "B", "C", "D"] and len(labels) == 4)
        results.append("+%d：阅读理解第一节四个语篇前分别有单独成段的语篇标识 A、B、C、D，无多余单字母段落"
                       % (1 if ok else 0))

        # 04 六个节标题括注：六条"第X节"标题的文字（含分值括注）逐一等于附录 A.2 节标题清单
        secs = [_nz(p["t"]) for p in P if re.match(r"^第[一二]节", p["t"].strip())]
        ok = (secs == A2_SECTION)
        results.append("+%d：六个节标题的分值括注完整且与附录 A.2 逐字一致（共 15 小题/共 5 小题/共 10 小题、"
                       "满分 37.5 分/12.5 分/15 分/15 分/15 分/25 分）" % (3 if ok else 0))

        # 05 题号：以"数字."开头的段落里，21-35 与 41-55 共 30 个题号齐全且不重复
        qnos = []
        for p in P:
            m = re.match(r"^\s*(\d{1,2})\.", p["t"])
            if m:
                qnos.append(int(m.group(1)))
        want = list(range(21, 36)) + list(range(41, 56))
        got = sorted(set(x for x in qnos if x in want))
        dup = [k for k in want if qnos.count(k) > 1]
        ok = (len(got) == 30 and not dup)
        results.append("+%d：题号 21-35、41-55 共 30 个齐全、不重复、不出现 36-40 与 56 以后的题号"
                       % (5 if ok else 0))

        # 06 四选一 ABCD 齐全：从每个题号段起、到下一个题号段止，聚合该题范围内的选项标号
        idx = [k for k, p in enumerate(P)
               if re.match(r"^\s*(\d{1,2})\.", p["t"])
               and int(re.match(r"^\s*(\d{1,2})\.", p["t"]).group(1)) in want]
        qbad = []
        for j, k in enumerate(idx):
            end = idx[j + 1] if j + 1 < len(idx) else min(k + 6, len(P))
            blob = "\n".join(P[x]["t"] for x in range(k, end))
            ls = re.findall(r"(?:^|[\s\t])([A-D])[\.\s]", blob)
            if sorted(set(ls)) != ["A", "B", "C", "D"]:
                qbad.append(int(re.match(r"^\s*(\d{1,2})\.", P[k]["t"]).group(1)))
        ok = (not qbad and len(idx) > 0)
        results.append("+%d：21-35 与 41-55 每题都配齐 A、B、C、D 四个选项标号，无缺项无多项"
                       % (5 if ok else 0))

        # 07 七选五：阅读理解第二节区间内，以「字母. 」开头且长度大于 25 的独立段落恰为 A-G 七条
        letters7 = [p["t"].strip()[0] for p in SEV]
        ok = ("".join(sorted(set(letters7))) == "ABCDEFG" and len(SEV) == 7)
        results.append("+%d：阅读理解第二节的备选项 A-G 共 7 条，每条各自独立成段" % (1 if ok else 0))

        # 08 空白编号：含英文的段落里出现的两位数中，36-46 与 56-65 共 21 个编号齐全
        wantb = list(range(36, 47)) + list(range(56, 66))
        bl = set()
        for p in P:
            if re.search(r"[A-Za-z]", p["t"]):
                for m in re.findall(r"(?<![\d.])(\d{2})(?![\d.])", p["t"]):
                    if int(m) in wantb:
                        bl.add(int(m))
        ok = (len(bl) == 21)
        results.append("+%d：语法填空与完形填空的空白编号 36-46、56-65 共 21 处齐全"
                       % (5 if ok else 0))

        # 09 读后续写：材料 5 个自然段（段首与附录 A.3 前缀一致）+ 两条独立成段的开头语（附录 A.7）
        k9 = next((p["ord"] for p in P if "续写两段" in p["t"]), None)
        mats = []
        if k9 is not None:
            for p in P[k9 + 1:k9 + 12]:
                if re.match(r"^\s*[A-Z]", p["t"]) and len(p["t"]) > 60 \
                        and "Paragraph" not in p["t"]:
                    mats.append(p["t"])
                if "注意" in p["t"]:
                    break
        okm = (len(mats) == 5
               and all(_nz(A3_STORY[j])[:18] in _nz(mats[j]) for j in range(min(5, len(mats)))))
        para9 = [p["t"].strip() for p in P if p["t"].strip().startswith("Paragraph")]
        ok7 = (len(para9) == 2
               and all(_nz(A7_OPENING[j])[:40] in _nz(para9[j]) for j in range(min(2, len(para9)))))
        ok = (okm and ok7)
        results.append("+%d：读后续写材料保留 5 个自然段（段首与附录 A.3 一致），并有 Paragraph 1、"
                       "Paragraph 2 两条开头语各自独立成段（与附录 A.7 一致）" % (3 if ok else 0))

        # 10 表格：5 行 5 列、表头等于附录 A.4、20 个数据格均非空
        rows10 = list(TB.rows) if TB is not None else []
        hdr10 = [c.text.strip() for c in rows10[0].cells] if rows10 else []
        cells10 = [[c.text.strip() for c in r.cells] for r in rows10[1:]]
        empt = [(ri + 1, ci) for ri, r in enumerate(cells10)
                for ci, c in enumerate(r) if not c]
        ok = (len(rows10) == 5 and hdr10 == A4_TABLE_HEADER and not empt)
        results.append("+%d：应用文表格为 5 行 5 列，表头依次为 Programme、Date (2026)、Time、"
                       "Fee (credits)、Main focus，其余 20 个数据格均非空" % (5 if ok else 0))

        # ==============================================================
        # 格式 1 分 18 条 = 18 分
        # ==============================================================
        # 11 纸张：w:pgSz w=11906 h=16838（A4 纵向）+ w:cols num=1（单栏）
        m = re.search(r"<w:pgSz ([^>]*)/>", D)
        pg = dict(re.findall(r'w:(\w+)="([^"]+)"', m.group(1))) if m else {}
        ncol = set(re.findall(r'w:num="(\d+)"', " ".join(re.findall(r"<w:cols ([^>]*)/>", D)))) or {"1"}
        ok = (pg.get("w") == "11906" and pg.get("h") == "16838"
              and pg.get("orient", "portrait") == "portrait" and ncol == {"1"})
        results.append("+%d：纸张为 A4 纵向、正文单栏排版" % (1 if ok else 0))

        # 12 页边距：w:pgMar 上下左右均为 720 缇 = 1.27 厘米（对话框显示的单位是厘米）
        m = re.search(r"<w:pgMar ([^>]*)/>", D)
        ma = dict(re.findall(r'w:(\w+)="([^"]+)"', m.group(1))) if m else {}
        ok = all(ma.get(kk) == "720" for kk in ("top", "bottom", "left", "right"))
        results.append("+%d：上下左右页边距均为 1.27 厘米" % (1 if ok else 0))

        # 13 页眉：header*.xml 内无任何文字，也无 w:pict / w:drawing / v: 图形声明
        hnames = sorted(n for n in parts if re.match(r"word/header\d*\.xml$", n))
        htxt = "".join("".join(re.findall(r"<w:t[^>]*>([^<]*)</w:t>", parts[n])) for n in hnames)
        hgfx = any(re.search(r"<w:pict|<w:drawing|<v:", parts[n]) for n in hnames)
        ok = (not htxt.strip() and not hgfx)
        results.append("+%d：页眉区域没有任何文字、横线或图形" % (1 if ok else 0))

        # 14 段前段后：全部段落 w:spacing/@before 与 @after 均为 0（或未设置）
        bad = [p["n"] for p in P if (p["sb"] or 0) > 0 or (p["sa"] or 0) > 0]
        ok = not bad
        results.append("+%d：所有段落的段前间距与段后间距均为 0" % (1 if ok else 0))

        # 15 语篇标识：A/B/C/D 四段均为水平居中，且含 12 磅加粗文字
        bad = [p["n"] for p in MARK if p["al"] != "CENTER" or not _has_run(p, 12.0, True)]
        ok = (len(MARK) == 4 and not bad)
        results.append("+%d：四个语篇标识 A、B、C、D 均为 12 磅、加粗、水平居中" % (1 if ok else 0))

        # 16 答题指导语：五处以「阅读下…」开头的段落（允许前置题号）均含 12 磅加粗文字
        bad = [p["n"] for p in GUIDE if not _has_run(p, 12.0, True)]
        ok = (len(GUIDE) == 5 and not bad)
        results.append("+%d：五处以「阅读下…」开头的答题指导语均为 12 磅加粗" % (1 if ok else 0))

        # 17 七选五：A-G 七条备选项段落首行缩进均为 2 字符（w:firstLineChars=200）
        bad = [p["n"] for p in SEV if p["flc"] != "200"]
        ok = (len(SEV) == 7 and not bad)
        results.append("+%d：阅读理解第二节 A-G 七条备选项段落首行缩进均为 2 字符" % (1 if ok else 0))

        # 18 引文与出处：各自独立成段（各只出现一段），且两段的有效 w:jc 为
        #    center；同时 w:ind 不得带首行/悬挂缩进、左右缩进要对称，否则文字
        #    实际不落在页面中线上。
        q18 = [p for p in P if _nz(A5_QUOTE)[:28] in _nz(p["t"])]
        s18 = [p for p in P if _nz(p["t"]) == _nz(A5_SOURCE) or "Ellison" in p["t"]]
        ok = (len(q18) == 1 and len(s18) == 1)
        if ok:
            for p in (q18[0], s18[0]):
                pe = PARAS[p["n"]]
                if _eff_jc(doc, pe) != "center" or not _ind_centered_ok(doc, pe):
                    ok = False
                    break
        results.append("+%d：写作第一节前的引文与出处各自独立成段，且均水平居中" % (1 if ok else 0))

        # 19 正文字号：卷首 4 段之后的全部段落，凡显式设置字号的 run 都必须是 12 磅
        bad = [p["n"] for p in P[4:]
               for r in p["runs"] if r["size"] is not None and r["size"] != 12.0]
        ok = not bad
        results.append("+%d：除卷首外，正文文字字号统一为 12 磅" % (1 if ok else 0))

        # 20 卷首三段：第 2-4 段（试卷名称等）水平居中且含 14 磅加粗文字
        bad = [p["n"] for p in P[1:4] if p["al"] != "CENTER" or not _has_run(p, 14.0, True)]
        ok = not bad
        results.append("+%d：卷首「澜桥市 2026 年高三年级第二学期质量检测」「英语试题」「2026 年 6 月」"
                       "三段为 14 磅、加粗、水平居中" % (1 if ok else 0))

        # 21 英文语句里不出现连续空格（跳过含制表符的选项行与含空白编号的行，
        #    这两类的空白是分列与填空占位，不属于语句内的连续空格）
        def _skip21(t):
            return ("\t" in t) or bool(
                re.search(r"(?<![\d.])(3[6-9]|4[0-6]|5[6-9]|6[0-5])(?![\d.])", t))
        bad = [mm for p in P if re.search(r"[A-Za-z]", p["t"]) and not _skip21(p["t"])
               for mm in re.findall(r"\w+ {2,}\w+", p["t"])]
        ok = not bad
        results.append("+%d：英文语句内部不出现两个及以上的连续空格" % (1 if ok else 0))

        # 22 选项标号后恰好 1 个空格：不允许「A.word」紧贴，也不允许「A.  word」两个空格
        bad = [x for x in re.findall(r"(?:^|[\s\t])([A-G]\.(?:[^\s]|  ))", full)
               if not re.match(r"^[A-G]\. $", x)]
        ok = not bad
        results.append("+%d：所有选项标号（A. 至 G.）后面均为 1 个空格" % (1 if ok else 0))

        # 23 无批注/修订，且页面上看不到文本框、图片、艺术字等形状。
        #    图形是否「看得见」按尺寸阈值判定：wp:extent 的 cx/cy 或 v:shape 的
        #    style 宽高小于 1 磅的声明属于网校下载件常见的隐形残留，不计入；
        #    再叠加包内 image 关系条数，页面上有图片必然先有 image 关系。
        has_comment = "word/comments.xml" in parts
        n_ins = len(re.findall(r"<w:ins[ >]", D))
        n_del = len(re.findall(r"<w:del[ >]", D))
        pnames = ["word/document.xml"] + hnames \
            + sorted(n for n in parts if re.match(r"word/footer\d*\.xml$", n))
        has_txbx = any("txbxContent" in parts.get(n, "") for n in pnames)
        n_shape = sum(_visible_shape_count(parts.get(n, "")) for n in pnames)
        n_imgrel = _image_rel_count(file_path)
        ok = (not has_comment and not n_ins and not n_del and not has_txbx
              and not n_shape and not n_imgrel)
        results.append("+%d：文档内无批注与修订痕迹，页面上也看不到文本框、图片、艺术字等形状"
                       % (1 if ok else 0))

        # 24 写作要求条目：（1）（2）（3）共 7 处各自单独成段（一段内不出现两个序号），
        #    这 7 段与两处「注意：」段的首行缩进均为 2 字符
        ITEM = [p for p in P if re.match(r"^\s*[（(]\s*[1-9]\s*[）)]", p["t"].strip())]
        NOTE = [p for p in P if re.fullmatch(r"注意[：:]?", p["t"].strip())]
        multi = [p["n"] for p in ITEM if len(re.findall(r"[（(]\s*[1-9]\s*[）)]", p["t"])) > 1]
        badind = [p["n"] for p in ITEM + NOTE if p["flc"] != "200"]
        ok = (len(ITEM) == 7 and len(NOTE) == 2 and not multi and not badind)
        results.append("+%d：两处写作要求的（1）（2）（3）条目共 7 处各自单独成段，这 7 段与两处「注意：」段"
                       "的首行缩进均为 2 字符" % (1 if ok else 0))

        # 25 无荧光笔高亮；表格 20 个数据格无底纹：逐格扫 w:tc 内的 w:shd，
        #    fill 取 auto/FFFFFF/nil 之外的值即算有底纹（表头 5 格允许有底纹）。
        hl = set(re.findall(r'<w:highlight w:val="(\w+)"', D)) - {"none"}
        shaded25 = _shaded_data_cells(TB)
        ok = (not hl and TB is not None and len(rows10) >= 2 and not shaded25)
        results.append("+%d：全文无荧光笔高亮，表格 20 个数据格均无底纹" % (1 if ok else 0))

        # 26 无连续空段、无手动分页符（w:br type=page 与 w:pageBreakBefore 都不允许）
        blanks = [p["ord"] for p in P if not p["t"].strip()]
        bset = set(blanks)
        consec = [b for b in blanks if b + 1 in bset]
        nbr = len(re.findall(r'<w:br w:type="page"', D))
        pbbs = [p["n"] for p in P if p["pbb"]]
        ok = (not consec and not nbr and not pbbs)
        results.append("+%d：不存在两个及以上相邻的空段落，也不存在手动分页符" % (1 if ok else 0))

        # 27 表格框线：只看 w:tblBorders / w:tcBorders —— 上下左右与内部横竖线
        #    六类边框必须齐全，且每一条都是 val=single、sz=4（八分之一磅制，
        #    = 0.5 磅）、color=777777；出现 nil/none 或别的线型、磅值、颜色即不得分。
        seg = "".join(re.findall(
            r"<w:tblBorders>.*?</w:tblBorders>|<w:tcBorders>.*?</w:tcBorders>", D, re.S))
        kinds = re.findall(r"<w:(top|left|bottom|right|insideH|insideV)\b([^>]*)/>", seg)
        tbb = [dict(re.findall(r'w:(\w+)="([^"]+)"', v)) for _k, v in kinds]
        vals = set(x.get("val") for x in tbb)
        szs = set(x.get("sz") for x in tbb)
        cols_ = set(x.get("color") for x in tbb)
        need27 = {"top", "left", "bottom", "right", "insideH", "insideV"}
        ok = (TB is not None and need27 <= set(k for k, _v in kinds)
              and vals == {"single"} and szs == {"4"} and cols_ == {"777777"})
        results.append("+%d：表格全部框线为单实线、0.5 磅、中灰色 777777" % (1 if ok else 0))

        # 28 表格无合并单元格（无 vMerge / hMerge / gridSpan），每行 5 格
        ncell = set(len(r.cells) for r in rows10)
        ok = ("vMerge" not in D and "hMerge" not in D and "gridSpan" not in D
              and ncell == {5})
        results.append("+%d：表格没有合并单元格，每一行都是 5 个单元格" % (1 if ok else 0))

        # ==============================================================
        # 格式 3 分 9 条 = 27 分
        # ==============================================================
        # 29 页脚：只约束「含 PAGE 域的那一段」本身——水平居中、页码与总页数都是域、
        #    段内无图形、去掉数字后文字为「第页/共页」。
        #    页脚里可能另有隐形水印段，不能把整个 footer1.xml 的文字拼起来匹配。
        fp = parts.get("word/footer1.xml", "")
        FPS = re.findall(r"<w:p\b.*?</w:p>", fp, re.S)
        pgp = [x for x in FPS
               if "PAGE" in "".join(re.findall(r"<w:instrText[^>]*>([^<]*)<", x)).upper()]
        ok = False
        if len(pgp) == 1:
            p0 = pgp[0]
            fld = "".join(re.findall(r"<w:instrText[^>]*>([^<]*)<", p0)).upper()
            jc0 = re.findall(r'<w:jc w:val="(\w+)"', p0)
            t0 = "".join(re.findall(r"<w:t[^>]*>([^<]*)</w:t>", p0))
            img0 = bool(re.search(r"<w:drawing|<w:pict|<a:blip", p0))
            t0n = re.sub(r"[\d\s ]", "", t0)
            ok = (jc0 == ["center"] and "PAGE" in fld and "NUMPAGES" in fld
                  and not img0 and t0n == "第页/共页")
        results.append("+%d：页脚只有一个页码段，水平居中，形如「第 N 页/共 M 页」，"
                       "其中当前页码与总页数均为自动更新的域" % (3 if ok else 0))

        # 30 字体：逐 run 解析实际生效的字体名（run 级 w:rFonts → 字符样式链 →
        #    段落样式链 → w:docDefaults）。含汉字的 run 其 w:eastAsia 必须是宋体；
        #    含西文字母或数字的 run 其 w:ascii/w:hAnsi 必须是 Times New Roman；
        #    既无汉字也无字母数字的纯符号 run（★ 之类）不在本条范围内。
        badfont = []
        nfont = 0
        for r30 in _iter_all_runs(doc):
            t30 = r30.text or ""
            if _HAN_RE.search(t30):
                nfont += 1
                nm = _resolve_effective_eastasia(doc, r30)
                if not _font_in(nm, SONGTI_NAMES):
                    badfont.append(("eastAsia", nm, t30[:12]))
            if re.search(r"[A-Za-z0-9]", t30):
                nfont += 1
                nm = _resolve_effective_ascii(doc, r30)
                if not _font_in(nm, TNR_NAMES):
                    badfont.append(("ascii", nm, t30[:12]))
        ok = (nfont > 0 and not badfont)
        results.append("+%d：中文字符渲染为宋体，西文字母与数字渲染为 Times New Roman"
                       % (3 if ok else 0))

        # 31 行距：卷首 4 段之后的全部段落 w:spacing lineRule=exact、line=400 缇 = 固定值 20 磅
        bad = [p["n"] for p in P[4:] if not (p["ls"] == 20.0 and p["lr"] == "exact")]
        ok = not bad
        results.append("+%d：除卷首 4 段外，全部段落行距为固定值 20 磅" % (3 if ok else 0))

        # 32 下划线作答区：连续的下划线段落合并为一处，共 3 处，每处折算长度不小于 1 整行。
        #    12 磅 Times New Roman 的下划线字符步进实测为 6.0 磅/个。
        uz = [p for p in P if p["t"].count("_") > 20]
        areas = []
        for p in uz:
            if areas and p["ord"] == areas[-1][-1]["ord"] + 1:
                areas[-1].append(p)
            else:
                areas.append([p])
        UCH = 6.0
        tot = [sum(q["t"].count("_") for q in g) * UCH for g in areas]
        ok = (len(areas) == 3 and all(t >= TXTW for t in tot))
        results.append("+%d：应用文与读后续写共 3 处下划线作答区，每处长度均不少于 1 整行"
                       % (3 if ok else 0))

        # 33 开头语加粗：应用文开头语 + Paragraph 1 / Paragraph 2 两条续写开头语均含加粗文字
        op33 = [p for p in P if p["t"].strip().startswith("Paragraph")]
        apl33 = [p for p in P if re.match(r"^(Dear|Hi |Hello)", p["t"].strip())
                 or ("Learning with a clear purpose" in p["t"])
                 or ("matters greatly" in p["t"])]
        cand33 = apl33 + op33
        bad = [p["n"] for p in cand33 if not _has_run(p, None, True)]
        ok = (len(op33) == 2 and len(apl33) >= 1 and not bad)
        results.append("+%d：应用文开头语与读后续写两条 Paragraph 开头语均加粗" % (3 if ok else 0))

        # 34 表格整体：w:tblPr/w:jc=center（表格整体水平居中）、全部单元格
        #    w:tcPr/w:vAlign=center（垂直居中）、表头 5 格 w:shd/@w:fill 为浅灰
        #    （三通道接近相等且落在 D9~FA）且格内 run 带 w:b、字号 w:sz=24（12 磅），
        #    其余数据格字号同为 12 磅。
        tjc = re.findall(r"<w:tblPr>(?:(?!</w:tblPr>).)*?<w:jc w:val=\"(\w+)\"", D, re.S)
        va = set(re.findall(r'<w:vAlign w:val="(\w+)"/>', D))
        hdr34 = False
        sz34 = False
        if TB is not None and len(rows10) == 5:
            hcells = list(rows10[0].cells)
            hdr34 = (len(hcells) == 5
                     and all(_cell_light_grey(c) for c in hcells)
                     and all(_cell_all_bold(doc, c) for c in hcells)
                     and all(_cell_sizes(doc, c) == {12.0} for c in hcells))
            sz34 = all(_cell_sizes(doc, c) == {12.0}
                       for r34 in rows10 for c in r34.cells)
        ok = (TB is not None and tjc == ["center"] and va == {"center"}
              and hdr34 and sz34)
        results.append("+%d：表格整体水平居中，表头 5 格加粗并带浅灰底纹，所有单元格文字垂直居中、字号 12 磅"
                       % (3 if ok else 0))

        # 35 题号为手工输入的普通文字，不是自动编号：无段落挂非 0 的 w:numId
        nid = [p["n"] for p in P if p["num"] not in (None, "0")]
        ok = (not nid and not re.findall(
            r"<w:numPr>(?!\s*<w:ilvl w:val=\"0\"/>\s*<w:numId w:val=\"0\"/>)", D))
        results.append("+%d：题号、选项标号与写作要求序号均为手工文字，未使用自动编号"
                       % (3 if ok else 0))

        # 36 无单词被空格拆开：不出现「三字母以上的词 + 1~2 个字母」这种断词形态
        #    （of/to/in 等常用短词排除在外）
        COMMON = {"a", "an", "of", "to", "in", "is", "it", "be", "on", "at", "as", "by",
                  "or", "if", "so", "no", "do", "we", "he", "me", "my", "us", "up",
                  "I", "A", "No", "Do", "It", "He", "We", "So", "id", "ok"}
        sus = [x for x in re.findall(r"\b([a-z]{3,}) ([a-z]{1,2})\b", full)
               if x[1] not in COMMON]
        ok = not sus
        results.append("+%d：没有任何英文单词被空格从内部拆开" % (3 if ok else 0))

        # 37 完形选项：15 个选项段各自把 A. B. C. D. 排在同一行的四列上。
        #    列起点按段落缩进 + 字符推进宽度 + w:tab 落点推算：四列起点必须依次
        #    递增，相邻两列之间至少留 2 磅空隙（上一列文字末端到下一列起点），
        #    整行末端不得超出正文宽度（否则会折行，四个选项就不在同一行）。
        ORD37 = ("A.", "B.", "C.", "D.")
        bad37 = []
        for p in CLOZE:
            pe = PARAS[p["n"]]
            cols37, flat37, endx37, nonleft37 = _option_columns(doc, pe, DTS, ORD37)
            if nonleft37 or not cols37 or "\n" in flat37:
                bad37.append(p["n"])
                continue
            st37 = [cols37[k][0] for k in ORD37]
            if any(st37[j + 1] <= st37[j] for j in range(3)):
                bad37.append(p["n"])
                continue
            gaps37 = [cols37[ORD37[j + 1]][0] - cols37[ORD37[j]][1] for j in range(3)]
            if min(gaps37) < 2.0 or endx37 > TXTW + 0.5:
                bad37.append(p["n"])
        ok = (len(CLOZE) == 15 and not bad37)
        results.append("+%d：完形填空 15 道题的四个选项在同一行内依次分列为 A. B. C. D.，"
                       "相邻两列之间留有空隙不相接" % (3 if ok else 0))

        # ==============================================================
        # 格式 5 分 5 条 = 25 分
        # ==============================================================
        # 38 语篇正文与材料段：首行缩进 2 字符（w:firstLineChars=200）
        bad = [p["n"] for p in PROSE if p["flc"] != "200"]
        ok = (not bad and len(PROSE) > 20)
        results.append("+%d：各语篇正文段落与读后续写材料段落的首行缩进均为 2 字符" % (5 if ok else 0))

        # 39 部分/节标题共 9 个：含 12 磅加粗文字，且首行缩进与左缩进均为 0（顶格）
        bad = [p["n"] for p in HD
               if (p["flc"] not in (None, "0")) or (p["fli"] not in (None, 0.0))
               or (p["li"] not in (None, 0.0)) or not _has_run(p, 12.0, True)]
        ok = (len(HD) == 9 and not bad)
        results.append("+%d：三个部分标题与六个节标题共 9 处均为 12 磅加粗，且首行缩进与左缩进为 0（顶格）"
                       % (5 if ok else 0))

        # 40 阅读选项段（不含完形选项段与七选五备选项段）：首行缩进 2 字符
        bad = [p["n"] for p in READOPT if p["flc"] != "200"]
        ok = (len(READOPT) > 20 and not bad)
        results.append("+%d：阅读理解各题的四个选项段落首行缩进均为 2 字符" % (5 if ok else 0))

        # 41 同一题内的两行式选项逐列对齐：上行以 A. 起、含 B.，下行以 C. 起、
        #    含 D.。A./C. 起点由两段的左缩进 + 首行缩进决定，B./D. 起点由制表位
        #    落点或前段文字推进宽度决定；同题 A. 与 C.、B. 与 D. 的起点差需在
        #    1 磅内。只在一道题内部比较，不要求跨题各列位置一致。
        pairs = []
        for j in range(len(P) - 1):
            t1 = (P[j]["t"] or "").strip()
            t2 = (P[j + 1]["t"] or "").strip()
            if (re.match(r"^A\.", t1) and "B." in t1
                    and re.match(r"^C\.", t2) and "D." in t2):
                pairs.append((j, j + 1))
        bad41 = []
        for a41, b41 in pairs:
            ca, _f1, _e1, nl1 = _option_columns(doc, PARAS[a41], DTS, ("A.", "B."))
            cb, _f2, _e2, nl2 = _option_columns(doc, PARAS[b41], DTS, ("C.", "D."))
            if nl1 or nl2 or not ca or not cb:
                bad41.append(a41)
                continue
            if (abs(ca["A."][0] - cb["C."][0]) >= 1.0
                    or abs(ca["B."][0] - cb["D."][0]) >= 1.0):
                bad41.append(a41)
        ok = (bool(pairs) and not bad41)
        results.append("+%d：采用上下两行排列选项的题目，同一题的 A. 与 C. 起始位置相同、"
                       "B. 与 D. 起始位置相同" % (5 if ok else 0))

        # 42 完形选项段：首行缩进 0.74 厘米（= 420 缇 = 21.0 磅），共 15 段
        bad = [p["n"] for p in CLOZE
               if p["fli"] is None or abs(p["fli"] - 20.98) >= 0.6]
        ok = (len(CLOZE) == 15 and not bad)
        results.append("+%d：完形填空 15 个选项段落的首行缩进均为 0.74 厘米" % (5 if ok else 0))

        # ==============================================================
        # 汇总：把 results 里的 "+N：规则" 拆成结构化 dim2_items
        # ==============================================================
        if len(results) != len(_MAX_DELTAS_ORDER):
            raise RuntimeError("dim2 项数不一致：results=%d vs order=%d"
                               % (len(results), len(_MAX_DELTAS_ORDER)))
        dim2_items = []
        total_score = 0
        for raw, max_delta in zip(results, _MAX_DELTAS_ORDER):
            head_s, _, rule_s = raw.partition("：")
            try:
                delta = int(head_s.replace("+", "").strip())
            except ValueError:
                delta = 0
            hit = delta > 0
            dim2_items.append({"rule": rule_s, "max_delta": max_delta,
                               "delta": max_delta if hit else 0, "hit": hit})
            total_score += (max_delta if hit else 0)
        result["dim2_items"] = dim2_items
        result["total_score"] = total_score
        result["max_score"] = MAX_SCORE_TOTAL
        return result
    finally:
        shutil.rmtree(tmp_dir, ignore_errors=True)


if __name__ == "__main__":
    _dir = sys.argv[1] if len(sys.argv) > 1 else os.path.dirname(os.path.abspath(__file__))
    try:
        _out = evaluate(_dir)
    except Exception as _exc:  # noqa: BLE001
        _out = {"id": SCRIPT_ID, "file_name": "", "status": "error",
                "error": "%s: %s" % (type(_exc).__name__, _exc),
                "dim1_pass": "", "dim1_reason": "",
                "dim2_items": [], "total_score": 0, "max_score": MAX_SCORE_TOTAL}
    print(json.dumps(_out, ensure_ascii=False, indent=2))
