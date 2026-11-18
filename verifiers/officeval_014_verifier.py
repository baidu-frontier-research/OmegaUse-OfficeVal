#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# pyright: reportUnknownMemberType=false, reportUnknownVariableType=false, reportUnknownArgumentType=false, reportUnknownParameterType=false, reportMissingParameterType=false, reportPrivateUsage=false, reportUnknownLambdaType=false, reportGeneralTypeIssues=false

import os
import re
import sys
import json
import zipfile
from docx import Document
from docx.table import Table
from docx.text.paragraph import Paragraph
from docx.text.run import Run
from docx.oxml.ns import qn
from lxml import etree

# 说明：接口统一约定，本模块不修改全局 sys.stdout，也不 print 主结果。
# 主结果统一由 evaluate(dir_path) -> dict 返回，入参为"脚本所在目录的路径"，
# 由脚本自己在该目录内定位并打开被评估的 .docx。本文件仅在 __main__ 下作为
# 本地调试入口，把该 dict 序列化为 JSON 打印。

SCRIPT_ID = "014"
# 维度二加分项满分；完成度 = max(0, 加分得分 − 扣分) ÷ MAX_SCORE_TOTAL
MAX_SCORE_TOTAL = 55
# 维度二减分项的最大扣分总额
MAX_PENALTY_TOTAL = 60

# 加分项 max_delta 的"结果追加顺序"列表，顺序须与 _evaluate_impl 中
# bonus.append(...) 的顺序严格一致，共 21 项，累加 = MAX_SCORE_TOTAL。
_MAX_DELTAS_ORDER = [
    5,  # 01 目录为可自动更新的目录域
    5,  # 02 七个一级标题样式标题1、十二个二级标题样式标题2
    3,  # 03 目录三条页码与标题实际所在页显示页码一致
    1,  # 04 目录条目四号黑体
    1,  # 05 封面中文题目与信息表两行小二楷体
    1,  # 06 封面英文题目小二 Times New Roman
    1,  # 07 关键词前空一行且"关键词"四号黑体
    1,  # 08 Subject words 前空一行且四号 TNR 加粗
    3,  # 09 三个二级标题小四黑体
    1,  # 10 图1图4图6 图题单倍行距
    3,  # 11 表1表7表13 表题单倍行距居中宋体
    5,  # 12 表1表7表13 表头行下方横线 0.75 磅
    3,  # 13 缩略词表表1表13 表头行五号加粗
    3,  # 14 表1表7表13 第2行小五
    3,  # 15 参考文献序号 [1]-[15] 连续
    3,  # 16 图2图3图5 图题单倍行距居中五号宋体
    1,  # 17 页边距 上2.5 下2.5 左3 右2.5 厘米
    5,  # 18 封面声明授权目录页脚无页码且节数不少于3
    3,  # 19 中文摘要页 Ⅰ、ABSTRACT 页 Ⅱ
    3,  # 20 缩略词表页至末页阿拉伯页码首页为1连续
    1,  # 21 摘要页正文中部末页页脚居中
]
assert sum(_MAX_DELTAS_ORDER) == MAX_SCORE_TOTAL

# 减分项 max_delta 的"结果追加顺序"列表（此处填绝对值），顺序须与
# _evaluate_impl 中 penalty.append(...) 的顺序严格一致，共 22 项，
# 累加 = MAX_PENALTY_TOTAL。命中时写入 dim2_items 的 delta 为负值。
_PENALTY_DELTAS_ORDER = [
    3,  # 01 三处正文段中文宋体小四、英文数字 TNR 小四
    3,  # 02 三处正文段 1.5 倍行距、首行缩进 2 字符
    1,  # 03 三个一级标题黑体四号
    1,  # 04 中文摘要标题与四段正文字体字号、无多余加粗与大字号
    1,  # 05 摘要分段引导词、关键词与 Subject words 内容格式
    1,  # 06 ABSTRACT 标题行及英文题目行 TNR 四号加粗
    3,  # 07 图2图5 图题、表4表11 表题的编号位置对齐
    5,  # 08 缩略词表表7表13 三线表线型
    3,  # 09 缩略词表表9表11 单元格中英文字体
    1,  # 10 参考文献第 3 8 15 条字体字号
    3,  # 11 图1图3图6 嵌入型且居中
    1,  # 12 页眉区出现文字或横线
    5,  # 13 32 个标题齐全、标题下有正文、无截断
    5,  # 14 表格数不少于 15、图片数不少于 8
    5,  # 15 15 个表格行列数、无空单元格、无错位
    3,  # 16 参考文献第 3 8 15 条著录内容
    3,  # 17 10 个"（续）"标题未被删改
    3,  # 18 封面与声明页、授权书页要件
    3,  # 19 新增标题、新增正文段落、译注或修改说明类文字
    1,  # 20 整页空白页、目录空白条目、连续 3 个以上空段落
    5,  # 21 目录变纯文本、对齐异常、标题1标题2正文样式被删
    1,  # 22 修订、批注、高亮、彩色文字
]
assert sum(_PENALTY_DELTAS_ORDER) == MAX_PENALTY_TOTAL

# ---------------------------------------------------------------------------
# 字体族与字号常量
# 字号一律换算为"字号框显示的磅值"再比较：小初 36、小二 18、四号 14、
# 小四 12、五号 10.5、小五 9。
# ---------------------------------------------------------------------------
HEITI = ("黑体", "SimHei", "Microsoft YaHei UI")
SONGTI = ("宋体", "SimSun", "NSimSun", "STSong")
KAITI = ("楷体", "KaiTi", "STKaiti", "楷体_GB2312", "KaiTi_GB2312")
TNR = ("Times New Roman", "TimesNewRoman", "Times New Roman PS")

PT_XIAOCHU = 36.0
PT_XIAOER = 18.0
PT_SIHAO = 14.0
PT_XIAOSI = 12.0
PT_WUHAO = 10.5
PT_XIAOWU = 9.0
EPS = 0.26  # 字号/磅值比较容差

CM = 360000  # 1 厘米对应的 EMU
TWIP_CM = 566.9291338582677  # 1 厘米对应的 twip

# 七个一级标题（+5 加分项 02 / -5 减分项 13 共用）
L1_TITLES = (
    "中英文缩略词表",
    "前言",
    "第一部分 多模态康复资料库构建",
    "第二部分 中医综合干预响应特征",
    "第三部分 疗效预测模型建立与解释",
    "讨论",
    "结论",
)
# 十二个二级标题
L2_TITLES = (
    "一、资料来源与研究设计",
    "二、变量清洗与质量控制",
    "三、统计方法与流程",
    "四、基线资料分析",
    "五、小结",
    "一、证候要素与干预路径",
    "二、疗效变化与阶段性表现",
    "三、亚组分析与安全性观察",
    "四、小结",
    "一、建模数据集与特征工程",
    "二、模型性能与校准评价",
    "三、关键变量解释与应用",
)

# 减分项 13 的 32 个标题清单（含重复项，按出现顺序）
ALL_TITLES = (
    "中文摘要",
    "ABSTRACT",
    "中英文缩略词表",
    "前言",
    "前言（续）",
    "第一部分 多模态康复资料库构建",
    "一、资料来源与研究设计",
    "一、资料来源与研究设计（续）",
    "二、变量清洗与质量控制",
    "三、统计方法与流程",
    "四、基线资料分析",
    "四、基线资料分析（续）",
    "四、基线资料分析（续）",
    "五、小结",
    "第二部分 中医综合干预响应特征",
    "一、证候要素与干预路径",
    "一、证候要素与干预路径（续）",
    "二、疗效变化与阶段性表现",
    "二、疗效变化与阶段性表现（续）",
    "三、亚组分析与安全性观察",
    "四、小结",
    "第三部分 疗效预测模型建立与解释",
    "一、建模数据集与特征工程",
    "二、模型性能与校准评价",
    "二、模型性能与校准评价（续）",
    "三、关键变量解释与应用",
    "三、关键变量解释与应用（续）",
    "讨论",
    "讨论（续）",
    "讨论（续）",
    "结论",
    "参考文献",
)
assert len(ALL_TITLES) == 32
_TITLE_SET = set(ALL_TITLES)

# 减分项 17 的 10 个"（续）"标题（含重复项）
CONT_TITLES = tuple(t for t in ALL_TITLES if t.endswith("（续）"))
assert len(CONT_TITLES) == 10

COVER_TITLE_CN = "基于多模态康复记录与机器学习研究中医综合干预冻结肩的疗效特征及预测因素"
COVER_TITLE_EN_HEAD = "A Multimodal Rehabilitation Record"

# 15 个表格的期望行列数，按文档中出现顺序：封面信息表、中英文缩略词表、表1-表13
TABLE_SHAPES = (
    (4, 2),   # 封面信息表
    (16, 3),  # 中英文缩略词表
    (7, 2),   # 表1
    (5, 2),   # 表2
    (6, 4),   # 表3
    (8, 3),   # 表4
    (7, 3),   # 表5
    (5, 3),   # 表6
    (5, 4),   # 表7
    (6, 5),   # 表8
    (7, 3),   # 表9
    (5, 4),   # 表10
    (6, 5),   # 表11
    (4, 3),   # 表12
    (5, 3),   # 表13
)
TABLE_NAMES = ("封面信息表", "缩略词表") + tuple("表%d" % i for i in range(1, 14))
assert len(TABLE_SHAPES) == len(TABLE_NAMES) == 15

# 减分项 16 的参考文献著录基准（比较时两侧均去除所有空白）
REF_FIELDS = {
    3: ("顾星辞,孟昭屹,樊朔,等", "温经活络方调理腰椎劳损病症临床研讨",
        "华夏中医骨伤学报", "2023", "31(06)", "28-33+39"),
    8: ("阮清禾,戚昱恒", "中西协同干预腰腿痹阻病症研究概况",
        "融合医学诊疗通报", "2025", "25(03)", "58-63"),
    15: ("CaoLN,WilsonH,ImCH", "Imagingevaluationstandardoflumbartissueinjurydegree",
         "Cross-borderMedicalImaging", "2025", "35(2)", "188-194"),
}

# 减分项 19 的修改说明类关键词
NOTE_KEYWORDS = (
    "修改说明", "修订说明", "译注", "译者注", "以下为修改", "本次修改",
    "说明：本次", "AI 生成", "AI生成", "由助手", "备注：以上", "编者注",
)

# ---------------------------------------------------------------------------
# 基础工具
# ---------------------------------------------------------------------------

def _norm(text):
    """归一化文本：去分页符、全角空格折叠为半角、压缩连续空白、去首尾空白。"""
    if not text:
        return ""
    t = text.replace("\x0c", " ").replace("　", " ").replace("\xa0", " ")
    t = re.sub(r"\s+", " ", t)
    return t.strip()


def _nospace(text):
    """去掉全部空白字符，用于著录内容这类"只比对实质字符"的场景。"""
    return re.sub(r"\s+", "", text or "")


def _para_text(para):
    return _norm(para.text)


def _is_blank(para):
    """段落是否为空段落（只含空白或分页符视为空）。"""
    return _para_text(para) == ""


def _near(a, b, eps=EPS):
    return a is not None and abs(float(a) - float(b)) <= eps


def _bool_of(el):
    """w:b / w:i 这类开关元素的三态解析：None 表示未设置。"""
    if el is None:
        return None
    v = el.get(qn("w:val"))
    if v is None:
        return True
    return str(v).lower() not in ("0", "false", "off")


def _styles_element(doc):
    try:
        return doc.styles.element
    except Exception:
        return None

def _style_by_id(doc, style_id):
    root = _styles_element(doc)
    if root is None or not style_id:
        return None
    for st in root.findall(qn("w:style")):
        if st.get(qn("w:styleId")) == style_id:
            return st
    return None


def _style_name(st):
    if st is None:
        return ""
    nm = st.find(qn("w:name"))
    return (nm.get(qn("w:val")) or "") if nm is not None else ""


def _para_style_id(para):
    pPr = para._element.find(qn("w:pPr"))
    if pPr is None:
        return None
    ps = pPr.find(qn("w:pStyle"))
    return ps.get(qn("w:val")) if ps is not None else None


def _style_chain(doc, para):
    """从段落样式起沿 w:basedOn 逐级向上，返回样式元素列表（不含 docDefaults）。"""
    chain = []
    sid = _para_style_id(para)
    seen = set()
    while sid and sid not in seen:
        seen.add(sid)
        st = _style_by_id(doc, sid)
        if st is None:
            break
        chain.append(st)
        based = st.find(qn("w:basedOn"))
        sid = based.get(qn("w:val")) if based is not None else None
    return chain


def _doc_defaults_rpr(doc):
    root = _styles_element(doc)
    if root is None:
        return None
    dd = root.find(qn("w:docDefaults"))
    if dd is None:
        return None
    rd = dd.find(qn("w:rPrDefault"))
    return rd.find(qn("w:rPr")) if rd is not None else None

def _doc_defaults_ppr(doc):
    root = _styles_element(doc)
    if root is None:
        return None
    dd = root.find(qn("w:docDefaults"))
    if dd is None:
        return None
    pd = dd.find(qn("w:pPrDefault"))
    return pd.find(qn("w:pPr")) if pd is not None else None


def _resolve_run_prop(doc, run, getter):
    """run 级 rPr → 段落样式链 rPr → docDefaults rPr，取第一个非 None。"""
    rPr = run._element.find(qn("w:rPr"))
    v = getter(rPr)
    if v is not None:
        return v
    para = getattr(run, "_parent", None)
    if para is not None and hasattr(para, "_element"):
        # 字符样式（w:rStyle）优先于段落样式
        if rPr is not None:
            rs = rPr.find(qn("w:rStyle"))
            if rs is not None:
                st = _style_by_id(doc, rs.get(qn("w:val")))
                if st is not None:
                    v = getter(st.find(qn("w:rPr")))
                    if v is not None:
                        return v
        for st in _style_chain(doc, para):
            v = getter(st.find(qn("w:rPr")))
            if v is not None:
                return v
    return getter(_doc_defaults_rpr(doc))


def _resolve_para_prop(doc, para, getter):
    """段落直接 pPr → 段落样式链 pPr → docDefaults pPr，取第一个非 None。"""
    v = getter(para._element.find(qn("w:pPr")))
    if v is not None:
        return v
    for st in _style_chain(doc, para):
        v = getter(st.find(qn("w:pPr")))
        if v is not None:
            return v
    return getter(_doc_defaults_ppr(doc))

def _get_rfonts(rPr, attr):
    if rPr is None:
        return None
    rf = rPr.find(qn("w:rFonts"))
    if rf is None:
        return None
    return rf.get(qn(attr))


def _resolve_effective_eastasia(doc, run):
    """解析 run 实际生效的中文（东亚）字体名。"""
    return _resolve_run_prop(doc, run, lambda r: _get_rfonts(r, "w:eastAsia"))


def _resolve_effective_ascii(doc, run):
    """解析 run 实际生效的西文（英文/数字）字体名。"""
    def _g(r):
        return _get_rfonts(r, "w:ascii") or _get_rfonts(r, "w:hAnsi")
    return _resolve_run_prop(doc, run, _g)


def _resolve_effective_size_pt(doc, run):
    """解析 run 实际生效的字号（磅），w:sz 为半磅值。"""
    def _g(r):
        if r is None:
            return None
        sz = r.find(qn("w:sz"))
        if sz is None:
            return None
        v = sz.get(qn("w:val"))
        try:
            return float(v) / 2.0
        except (TypeError, ValueError):
            return None
    return _resolve_run_prop(doc, run, _g)


def _resolve_effective_bold(doc, run):
    def _g(r):
        return _bool_of(r.find(qn("w:b"))) if r is not None else None
    v = _resolve_run_prop(doc, run, _g)
    return bool(v)

def _resolve_effective_color(doc, run):
    """返回 (w:color/@w:val, w:color/@w:themeColor)。"""
    def _g(r):
        if r is None:
            return None
        c = r.find(qn("w:color"))
        if c is None:
            return None
        return (c.get(qn("w:val")), c.get(qn("w:themeColor")))
    v = _resolve_run_prop(doc, run, _g)
    return v if v else (None, None)


def _run_is_superscript(doc, run):
    def _g(r):
        if r is None:
            return None
        va = r.find(qn("w:vertAlign"))
        return va.get(qn("w:val")) if va is not None else None
    return _resolve_run_prop(doc, run, _g) == "superscript"


def _run_highlight(doc, run):
    def _g(r):
        if r is None:
            return None
        h = r.find(qn("w:highlight"))
        return h.get(qn("w:val")) if h is not None else None
    v = _resolve_run_prop(doc, run, _g)
    return v if v not in (None, "none") else None


def _iter_para_runs(para):
    """按顺序产出段落内的 run，包括 w:hyperlink / w:ins / w:smartTag / w:sdt 内的
    run（目录条目的文字常被包在 w:hyperlink 内，python-docx 的 runs 取不到）。"""
    wrappers = (qn("w:hyperlink"), qn("w:ins"), qn("w:smartTag"), qn("w:sdt"))
    r_tag = qn("w:r")
    out = []
    for el in para._element.iterchildren():
        if el.tag == r_tag:
            out.append(Run(el, para))
        elif el.tag in wrappers:
            for sub in el.iter(r_tag):
                out.append(Run(sub, para))
    return out


def _text_runs(para):
    """只保留含可见文字的 run（跳过空 run 与仅含分页符的 run）。"""
    out = []
    for r in _iter_para_runs(para):
        if _norm(r.text):
            out.append(r)
    return out

_CJK_RE = re.compile(r"[㐀-䶿一-鿿豈-﫿]")
_LATIN_RE = re.compile(r"[A-Za-z0-9]")


def _font_ok(name, allowed):
    if not name:
        return False
    n = str(name).strip()
    return any(n == a or n.lower() == a.lower() for a in allowed)


def _para_font_problems(doc, para, ea_allowed, ascii_allowed, size_pt,
                        require_bold=None):
    """检查段落内所有可见 run 的中文字体、西文字体、字号（可选加粗）。

    返回问题描述列表，空列表表示全部合规。ea_allowed / ascii_allowed 传
    None 表示不检查该类字体；size_pt 传 None 表示不检查字号。
    """
    problems = []
    runs = _text_runs(para)
    if not runs:
        return ["无可见文字"]
    for r in runs:
        txt = r.text
        if ea_allowed and _CJK_RE.search(txt):
            nm = _resolve_effective_eastasia(doc, r)
            if not _font_ok(nm, ea_allowed):
                problems.append("中文字体为 %s" % (nm or "未设置"))
        if ascii_allowed and _LATIN_RE.search(txt):
            nm = _resolve_effective_ascii(doc, r)
            if not _font_ok(nm, ascii_allowed):
                problems.append("西文字体为 %s" % (nm or "未设置"))
        if size_pt is not None:
            sz = _resolve_effective_size_pt(doc, r)
            if not _near(sz, size_pt):
                problems.append("字号为 %s 磅" % (sz if sz is not None else "未设置"))
        if require_bold is not None:
            if _resolve_effective_bold(doc, r) != require_bold:
                problems.append("加粗为 %s" % _resolve_effective_bold(doc, r))
    return problems

def _line_spacing(doc, para):
    """返回 (kind, value)：kind ∈ ('multiple','exact','atLeast')。

    未显式设置时按 Word 默认的单倍行距返回 ('multiple', 1.0)。
    """
    def _g(pPr):
        if pPr is None:
            return None
        sp = pPr.find(qn("w:spacing"))
        if sp is None:
            return None
        line = sp.get(qn("w:line"))
        if line is None:
            return None
        try:
            n = float(line)
        except ValueError:
            return None
        rule = sp.get(qn("w:lineRule"))
        if rule in ("exact", "atLeast"):
            return (rule, n / 20.0)
        return ("multiple", n / 240.0)
    v = _resolve_para_prop(doc, para, _g)
    return v if v else ("multiple", 1.0)


def _is_single_spacing(doc, para):
    """段落对话框里行距是否显示为"单倍行距"。"""
    kind, val = _line_spacing(doc, para)
    return kind == "multiple" and _near(val, 1.0, 0.02)


def _is_15_spacing(doc, para):
    kind, val = _line_spacing(doc, para)
    return kind == "multiple" and _near(val, 1.5, 0.02)


def _alignment(doc, para):
    def _g(pPr):
        if pPr is None:
            return None
        jc = pPr.find(qn("w:jc"))
        return jc.get(qn("w:val")) if jc is not None else None
    return _resolve_para_prop(doc, para, _g)


def _is_centered(doc, para):
    return _alignment(doc, para) == "center"

def _ind_element(doc, para):
    def _g(pPr):
        if pPr is None:
            return None
        return pPr.find(qn("w:ind"))
    return _resolve_para_prop(doc, para, _g)


def _first_line_indent_desc(doc, para):
    """返回 (是否显示为"首行缩进 2 字符", 实际显示描述)。

    只有 w:firstLineChars=200 时对话框才会显示"2 字符"；仅给出 w:firstLine
    的绝对值时对话框显示厘米（如 0.74 厘米），按细则的单位严格性不合。
    """
    ind = _ind_element(doc, para)
    if ind is None:
        return (False, "无缩进")
    hanging = ind.get(qn("w:hanging")) or ind.get(qn("w:hangingChars"))
    if hanging and str(hanging).strip() not in ("0",):
        return (False, "特殊格式为悬挂缩进")
    chars = ind.get(qn("w:firstLineChars"))
    if chars is not None:
        try:
            n = float(chars) / 100.0
        except ValueError:
            n = None
        if n is not None and _near(n, 2.0, 0.02):
            return (True, "2 字符")
        return (False, "%s 字符" % n)
    first = ind.get(qn("w:firstLine"))
    if first is not None:
        try:
            cm = float(first) / TWIP_CM
        except ValueError:
            return (False, "首行缩进值异常")
        return (False, "%.2f 厘米" % cm)
    return (False, "无首行缩进")

def _iter_block_items(doc):
    """按正文出现顺序产出 Paragraph / Table，并下钻 w:sdt（目录常被包在其中）。"""
    body = doc.element.body

    def _walk(node):
        for child in node.iterchildren():
            tag = child.tag
            if tag == qn("w:p"):
                yield Paragraph(child, doc)
            elif tag == qn("w:tbl"):
                yield Table(child, doc)
            elif tag in (qn("w:sdt"), qn("w:sdtContent")):
                for item in _walk(child):
                    yield item
    for item in _walk(body):
        yield item


def _all_paragraphs(doc):
    """正文全部段落（含 w:sdt 内），不含表格单元格内段落。"""
    return [b for b in _iter_block_items(doc) if isinstance(b, Paragraph)]


def _all_tables(doc):
    return [b for b in _iter_block_items(doc) if isinstance(b, Table)]


def _outline_level(doc, para):
    """返回段落的有效大纲级别（0 = 一级标题，1 = 二级标题），非标题返回 None。"""
    def _g(pPr):
        if pPr is None:
            return None
        ol = pPr.find(qn("w:outlineLvl"))
        if ol is None:
            return None
        try:
            return int(ol.get(qn("w:val")))
        except (TypeError, ValueError):
            return None
    lvl = _resolve_para_prop(doc, para, _g)
    if lvl is not None:
        return lvl if lvl <= 8 else None
    for st in _style_chain(doc, para):
        m = re.match(r"^(?:heading|标题)\s*([1-9])$", _style_name(st).strip(), re.I)
        if m:
            return int(m.group(1)) - 1
    return None

def _applied_style_name(doc, para):
    """段落直接套用的段落样式在样式库里显示的名称。"""
    st = _style_by_id(doc, _para_style_id(para))
    return _style_name(st)


def _is_heading_style(doc, para, level):
    """段落样式是否为「标题 level」（level 从 1 起计）。"""
    nm = _applied_style_name(doc, para).strip()
    if re.match(r"^(?:heading|标题)\s*%d$" % level, nm, re.I):
        return True
    # 兼容以内建标题样式为基样式、且大纲级别一致的本地化命名
    if _outline_level(doc, para) == level - 1:
        for st in _style_chain(doc, para):
            if re.match(r"^(?:heading|标题)\s*%d$" % level, _style_name(st).strip(), re.I):
                return True
    return False


def _find_paras(paras, text):
    t = _norm(text)
    return [i for i, p in enumerate(paras) if _para_text(p) == t]


def _find_para(paras, text):
    hits = _find_paras(paras, text)
    return hits[0] if hits else -1


def _find_para_startswith(paras, prefix, start=0):
    pfx = _norm(prefix)
    for i in range(start, len(paras)):
        if _para_text(paras[i]).startswith(pfx):
            return i
    return -1


def _next_heading_index(doc, paras, start, allow_texts=()):
    """从 start+1 起找下一个标题段落（按标题文字或大纲级别判定，可跳过续接标题）。"""
    allow = {_norm(a) for a in allow_texts}
    for i in range(start + 1, len(paras)):
        txt = _para_text(paras[i])
        if not txt or txt in allow:
            continue
        lvl = _outline_level(doc, paras[i])
        if txt in _TITLE_SET or (lvl is not None and lvl <= 2):
            return i
    return len(paras)


def _title_level(text):
    """按 32 个标题的固有层级返回 1 或 2（原文件无标题样式时也可判定）。"""
    base = text.replace("（续）", "")
    return 2 if base in L2_TITLES else 1


def _lead_re(lead):
    """摘要分段引导词的匹配式，兼容"目的"与"[目的]"两种写法。"""
    return re.compile(r"^\s*[\[［【]?\s*%s\s*[\]］】]?" % re.escape(lead))


def _starts_with_lead(text, lead):
    return _lead_re(lead).match(text) is not None


def _lead_len(text, lead):
    """引导词在段首实际占用的字符数（含方括号）。"""
    m = _lead_re(lead).match(text)
    return len(m.group(0)) if m else len(lead)

NO_LINE = (None, "nil", "none")


def _tbl_borders_el(doc, tbl):
    tblPr = tbl._tbl.find(qn("w:tblPr"))
    if tblPr is not None:
        b = tblPr.find(qn("w:tblBorders"))
        if b is not None:
            return b
        ts = tblPr.find(qn("w:tblStyle"))
        if ts is not None:
            st = _style_by_id(doc, ts.get(qn("w:val")))
            if st is not None:
                sp = st.find(qn("w:tblPr"))
                if sp is not None:
                    return sp.find(qn("w:tblBorders"))
    return None


def _cell_borders_el(cell):
    tcPr = cell._tc.find(qn("w:tcPr"))
    return tcPr.find(qn("w:tcBorders")) if tcPr is not None else None


def _edge(borders_el, name):
    """返回 (w:val, 线宽磅值)；w:sz 以 1/8 磅为单位。"""
    if borders_el is None:
        return (None, None)
    e = borders_el.find(qn("w:" + name))
    if e is None:
        return (None, None)
    val = e.get(qn("w:val"))
    sz = e.get(qn("w:sz"))
    try:
        pt = float(sz) / 8.0 if sz is not None else None
    except ValueError:
        pt = None
    return (val, pt)


def _effective_edge(doc, tbl, cell, cell_name, tbl_name):
    val, pt = _edge(_cell_borders_el(cell), cell_name)
    if val is not None:
        return (val, pt)
    return _edge(_tbl_borders_el(doc, tbl), tbl_name)


def _has_line(val):
    return val not in NO_LINE

def _numbering_root(doc):
    try:
        return doc.part.numbering_part.element
    except Exception:
        return None


def _num_pr(para):
    pPr = para._element.find(qn("w:pPr"))
    return pPr.find(qn("w:numPr")) if pPr is not None else None


def _num_lvl_text(doc, para):
    """解析自动编号段落的编号格式串（如 "[%1]"）；无自动编号返回 None。"""
    npr = _num_pr(para)
    num_id = ilvl = None
    if npr is not None:
        ni = npr.find(qn("w:numId"))
        il = npr.find(qn("w:ilvl"))
        num_id = ni.get(qn("w:val")) if ni is not None else None
        ilvl = il.get(qn("w:val")) if il is not None else "0"
    if num_id is None:
        for st in _style_chain(doc, para):
            spr = st.find(qn("w:pPr"))
            snpr = spr.find(qn("w:numPr")) if spr is not None else None
            if snpr is not None:
                ni = snpr.find(qn("w:numId"))
                il = snpr.find(qn("w:ilvl"))
                num_id = ni.get(qn("w:val")) if ni is not None else None
                ilvl = il.get(qn("w:val")) if il is not None else "0"
                break
    if num_id is None or num_id == "0":
        return None
    root = _numbering_root(doc)
    if root is None:
        return None
    abs_id = None
    for num in root.findall(qn("w:num")):
        if num.get(qn("w:numId")) == num_id:
            a = num.find(qn("w:abstractNumId"))
            abs_id = a.get(qn("w:val")) if a is not None else None
            break
    if abs_id is None:
        return None
    for an in root.findall(qn("w:abstractNum")):
        if an.get(qn("w:abstractNumId")) != abs_id:
            continue
        for lvl in an.findall(qn("w:lvl")):
            if lvl.get(qn("w:ilvl")) == (ilvl or "0"):
                lt = lvl.find(qn("w:lvlText"))
                return lt.get(qn("w:val")) if lt is not None else None
    return None

WP_NS = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"


def _has_drawing(para):
    el = para._element
    return (el.findall(".//" + qn("w:drawing")) or
            el.findall(".//" + qn("w:pict")))


def _drawing_wrap(para):
    """返回 ('inline'|'anchor'|'vml-absolute'|None)：图片的环绕方式。"""
    el = para._element
    if el.findall(".//{%s}inline" % WP_NS):
        return "inline"
    if el.findall(".//{%s}anchor" % WP_NS):
        return "anchor"
    for sh in el.findall(".//{urn:schemas-microsoft-com:vml}shape"):
        style = (sh.get("style") or "").replace(" ", "").lower()
        if "position:absolute" in style:
            return "vml-absolute"
        return "inline"
    if el.findall(".//" + qn("w:pict")):
        return "inline"
    return None


def _doc_xml(file_path):
    """读取 word/document.xml 原文，用于 python-docx 不便覆盖的结构性判定。"""
    with zipfile.ZipFile(file_path) as zf:
        return zf.read("word/document.xml").decode("utf-8", "ignore")


def _part_xml(file_path, name):
    try:
        with zipfile.ZipFile(file_path) as zf:
            return zf.read(name).decode("utf-8", "ignore")
    except (KeyError, OSError):
        return ""


def _part_names(file_path):
    try:
        with zipfile.ZipFile(file_path) as zf:
            return list(zf.namelist())
    except OSError:
        return []

def _has_toc_field(file_path):
    """目录是否为可自动更新的目录域（TOC 域或目录内容控件）。"""
    xml = _doc_xml(file_path)
    for m in re.finditer(r"<w:instrText[^>]*>(.*?)</w:instrText>", xml, re.S):
        instr = m.group(1)
        if re.search(r"\bTOC\b", instr):
            return True
    for m in re.finditer(r'<w:fldSimple[^>]*w:instr="([^"]*)"', xml):
        if re.search(r"\bTOC\b", m.group(1)):
            return True
    if re.search(r'w:docPartGallery[^>]*w:val="Table of [Cc]ontents"', xml):
        return True
    return False


def _toc_style_paras(doc, paras):
    """按 toc N / 目录 N 样式识别目录条目段落。"""
    out = []
    for i, p in enumerate(paras):
        nm = _applied_style_name(doc, p).strip()
        if re.match(r"^(?:toc|目录)\s*[1-9]$", nm, re.I):
            out.append(i)
    return out


_TOC_LINE_RE = re.compile(r"^(.*?)[\s.…]*\t?\s*(\d+)$")


def _toc_entries(doc, paras):
    """返回目录条目 [(标题文字, 页码整数或 None)]。"""
    idxs = _toc_style_paras(doc, paras)
    if not idxs:
        start = _find_para(paras, "目录")
        if start < 0:
            return []
        stop = _next_heading_index(doc, paras, start)
        idxs = list(range(start + 1, stop))
    entries = []
    for i in idxs:
        raw = paras[i].text.replace("\x0c", "")
        txt = _norm(raw)
        if not txt:
            entries.append(("", None))
            continue
        m = _TOC_LINE_RE.match(txt)
        if m:
            entries.append((_norm(m.group(1)).rstrip(".。…·"), int(m.group(2))))
        else:
            entries.append((txt, None))
    return entries

def _open_word_document(file_path):
    return Document(file_path)


# ---------------------------------------------------------------------------
# 静态版面模型：分节 / 强制分页 / 页码域 / 页脚对齐
#
# 页码、页脚、空白页这类"排版后才成立"的细则，全部改为按 OOXML 静态推算：
#   分节        w:sectPr（w:footerReference@w:type、w:titlePg、w:pgNumType）
#   强制分页    w:br w:type="page"、w:pPr/w:pageBreakBefore、分节符 w:type
#   页码是否为域 页脚部件里 w:instrText / w:fldSimple@w:instr 是否含 PAGE
#   页脚是否居中 页脚段落有效 w:jc（含 w:pStyle 继承链），或页码被放进
#               居中锚定的文本框（wp:positionH/wp:align=center、w:framePr@w:xAlign）
# 由此得到的页序号只统计强制分页，是真实页序的下界；凡是"必须恰好在第 N 页"
# 的判据都只在该下界能唯一确定时才给分。
# ---------------------------------------------------------------------------

_ROMAN_MAP = {
    "I": 1, "II": 2, "III": 3, "IV": 4, "V": 5, "VI": 6, "VII": 7,
    "VIII": 8, "IX": 9, "X": 10, "XI": 11, "XII": 12,
    "Ⅰ": 1, "Ⅱ": 2, "Ⅲ": 3, "Ⅳ": 4, "Ⅴ": 5, "Ⅵ": 6, "Ⅶ": 7,
    "Ⅷ": 8, "Ⅸ": 9, "Ⅹ": 10, "Ⅺ": 11, "Ⅻ": 12,
}
_ARABIC_ONLY = re.compile(r"^[\-–—\s\[\(]*([0-9]{1,4})[\-–—\s\]\)]*$")
_ROMAN_ONLY = re.compile(r"^[\-–—\s\[\(]*([IVXⅠ-Ⅻ]{1,6})[\-–—\s\]\)]*$")

_R_NS = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
_WPD_NS = "{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}"


def _w(tag):
    return qn("w:" + tag)


def _el_text(el):
    """元素下所有 w:t 的文字（去空白）；w:instrText 是另一个标签，不会混进来。"""
    return _nospace("".join(t.text or "" for t in el.iter(_w("t"))))


def _switch_on(el):
    if el is None:
        return False
    v = el.get(_w("val"))
    return v is None or str(v).lower() not in ("0", "false", "off")


def _el_visible(el):
    """块级条目在页面上是否有可见内容（文字、表格、图形）。"""
    if el.tag == _w("tbl"):
        return True
    if _el_text(el):
        return True
    for tag in ("drawing", "pict", "object"):
        if el.find(".//" + _w(tag)) is not None:
            return True
    return False


def _br_page_count(el):
    """段落内手动分页符（w:br w:type="page"）个数。"""
    if el.tag != _w("p"):
        return 0
    n = 0
    for r in el.iter(_w("r")):
        for b in r.iter(_w("br")):
            if b.get(_w("type")) == "page":
                n += 1
    return n


def _page_break_before(el):
    if el.tag != _w("p"):
        return False
    pPr = el.find(_w("pPr"))
    return pPr is not None and _switch_on(pPr.find(_w("pageBreakBefore")))


def _anchored_center(para_el):
    """页码被放进居中锚定的文本框/图文框时，排版结果同样是水平居中。"""
    node = para_el
    for _ in range(12):
        if node is None:
            break
        for ph in node.iter(_WPD_NS + "positionH"):
            al = ph.find(_WPD_NS + "align")
            if al is not None and (al.text or "").strip() == "center":
                return True
        if node.tag == _w("ftr"):
            break
        node = node.getparent()
    pPr = para_el.find(_w("pPr"))
    fr = pPr.find(_w("framePr")) if pPr is not None else None
    return fr is not None and fr.get(_w("xAlign")) == "center"


def _field_result_text(el):
    """取段落内最后一个域的结果文字（fldChar separate 与 end 之间的 w:t）。"""
    res = []
    cur = None
    for node in el.iter():
        if node.tag == _w("fldChar"):
            v = node.get(_w("fldCharType"))
            if v == "separate":
                cur = []
            elif v == "end":
                if cur is not None:
                    res.append("".join(cur))
                cur = None
        elif node.tag == _w("t") and cur is not None:
            cur.append(node.text or "")
    return _norm(res[-1]) if res else ""


class _PageModel(object):
    """正文条目 → 节 / 强制分页页序 / 显示页码；以及逐节页脚取证。"""

    def __init__(self, doc, file_path, blocks):
        self.els = [b._element for b in blocks]
        self.styles = _styles_element(doc)
        self.body = doc.element.body

        # ---- 分节：段落级 w:sectPr 结束一节，body 末尾的 w:sectPr 是最后一节
        self.sectprs = []
        self.sect_of = []
        k = 0
        for el in self.els:
            self.sect_of.append(k)
            if el.tag == _w("p"):
                pPr = el.find(_w("pPr"))
                sp = pPr.find(_w("sectPr")) if pPr is not None else None
                if sp is not None:
                    self.sectprs.append(sp)
                    k += 1
        tail = self.body.find(_w("sectPr"))
        if tail is not None:
            self.sectprs.append(tail)
        while len(self.sectprs) <= k:
            self.sectprs.append(None)
        self.n_sect = len(self.sectprs)

        # ---- 强制分页
        n = len(self.els)
        self.pages = [0] * n
        self.bnd = [0] * n
        for i in range(1, n):
            br = _br_page_count(self.els[i - 1])
            cut_here = _page_break_before(self.els[i])
            if (self.sect_of[i] != self.sect_of[i - 1]
                    and self._sect_breaks_page(self.sect_of[i - 1])):
                cut_here = True
            self.bnd[i] = br + (1 if cut_here else 0)
            self.pages[i] = self.pages[i - 1] + max(br, 1 if cut_here else 0)

        # ---- 每节页区间与首页显示页码
        self.spans = {}
        for i, s in enumerate(self.sect_of):
            if s not in self.spans:
                self.spans[s] = [self.pages[i], self.pages[i]]
            else:
                self.spans[s][1] = max(self.spans[s][1], self.pages[i])
        self.dstart = {}
        prev_end = 0
        for s in range(self.n_sect):
            st = self.start(s)
            if st is None:
                st = prev_end + 1 if s else 1
            self.dstart[s] = st
            if s in self.spans:
                prev_end = st + (self.spans[s][1] - self.spans[s][0])
            else:
                prev_end = st

        # ---- 页脚部件与关系
        self.rel = {}
        rels = _part_xml(file_path, "word/_rels/document.xml.rels")
        if rels:
            try:
                for r in etree.fromstring(rels.encode("utf-8")):
                    self.rel[r.get("Id")] = r.get("Target")
            except Exception:
                pass
        self.footers = {}
        for nm in _part_names(file_path):
            if re.match(r"word/footer\d*\.xml$", nm):
                try:
                    self.footers[nm.split("/")[-1]] = etree.fromstring(
                        _part_xml(file_path, nm).encode("utf-8"))
                except Exception:
                    pass
        st_xml = _part_xml(file_path, "word/settings.xml")
        m = re.search(r"<w:evenAndOddHeaders\b([^>]*)>", st_xml)
        self.even_odd = bool(m) and not re.search(
            r'w:val="(?:0|false|off)"', m.group(1))

    # ------------------------------------------------------------- 分节属性
    def _sect(self, k):
        return self.sectprs[k] if 0 <= k < len(self.sectprs) else None

    def _sect_breaks_page(self, k):
        sp = self._sect(k)
        if sp is None:
            return True
        t = sp.find(_w("type"))
        v = t.get(_w("val")) if t is not None else None
        return v in (None, "nextPage", "oddPage", "evenPage")

    def _pgnum(self, k):
        sp = self._sect(k)
        if sp is None:
            return (None, None)
        pn = sp.find(_w("pgNumType"))
        if pn is None:
            return (None, None)
        return (pn.get(_w("fmt")), pn.get(_w("start")))

    def fmt(self, k):
        """页码格式；未写 w:pgNumType/@w:fmt 时按 OOXML 缺省的 decimal。"""
        return self._pgnum(k)[0] or "decimal"

    def start(self, k):
        try:
            return int(self._pgnum(k)[1])
        except (TypeError, ValueError):
            return None

    def expected_start(self, k):
        """按静态页数接续上一节应得的首页页码。"""
        if k <= 0 or (k - 1) not in self.spans:
            return None
        return self.dstart[k - 1] + (self.spans[k - 1][1] - self.spans[k - 1][0]) + 1

    # --------------------------------------------------------------- 页脚
    def sect_footers(self, k):
        """逐类型解析生效的页脚部件；本节未声明该类型时沿用前一节。"""
        out = {}
        for t in ("default", "first", "even"):
            part = None
            for j in range(k, -1, -1):
                sp = self._sect(j)
                if sp is None:
                    continue
                hit = None
                for fr in sp.findall(_w("footerReference")):
                    if fr.get(_w("type")) == t:
                        hit = fr
                if hit is not None:
                    tgt = self.rel.get(hit.get(_R_NS + "id")) or ""
                    part = self.footers.get(tgt.split("/")[-1])
                    break
            out[t] = part
        return out

    def applicable_footers(self, k):
        """本节各页可能用到的页脚：default，首页不同则加 first，奇偶不同则加 even。"""
        sp = self._sect(k)
        fs = self.sect_footers(k)
        out = [("default", fs.get("default"))]
        if sp is not None and _switch_on(sp.find(_w("titlePg"))):
            out.append(("first", fs.get("first")))
        if self.even_odd:
            out.append(("even", fs.get("even")))
        return out

    def _leaf_paras(self, ftr):
        if ftr is None:
            return []
        return [p for p in ftr.iter(_w("p")) if p.find(".//" + _w("p")) is None]

    def footer_num_info(self, ftr):
        """返回 (是否显示页码, 含 PAGE 域的段落, 手敲页码的段落)。"""
        if ftr is None:
            return (False, [], [])
        fld, hand = [], []
        for p in self._leaf_paras(ftr):
            instr = "".join(t.text or "" for t in p.iter(_w("instrText"))).upper()
            for fs in p.iter(_w("fldSimple")):
                instr += " " + (fs.get(_w("instr")) or "").upper()
            if re.search(r"\bPAGE\b", instr):
                fld.append(p)
                continue
            txt = _el_text(p)
            if txt and (_ARABIC_ONLY.match(txt) or _ROMAN_ONLY.match(txt)):
                hand.append(p)
        return (bool(fld or hand), fld, hand)

    def footer_no_number(self, k):
        for _t, f in self.applicable_footers(k):
            if self.footer_num_info(f)[0]:
                return False
        return True

    def footer_page_field_ok(self, k):
        """本节各页页脚都必须靠 PAGE 域出页码，手敲数字不算。"""
        got = False
        for _t, f in self.applicable_footers(k):
            if not self.footer_num_info(f)[1]:
                return False
            got = True
        return got

    def footer_centered(self, k):
        got = False
        for _t, f in self.applicable_footers(k):
            _has, fld, hand = self.footer_num_info(f)
            paras = fld or hand
            if not paras:
                return False
            for p in paras:
                if self._eff_jc(p) == "center":
                    continue
                if _anchored_center(p):
                    continue
                return False
            got = True
        return got

    def _style_by_id(self, sid):
        if self.styles is None or not sid:
            return None
        for s in self.styles.findall(_w("style")):
            if s.get(qn("w:styleId")) == sid:
                return s
        return None

    def _eff_jc(self, p):
        """页脚段落的有效对齐：段落 w:jc → w:pStyle 继承链 → docDefaults。"""
        pPr = p.find(_w("pPr"))
        if pPr is not None:
            jc = pPr.find(_w("jc"))
            if jc is not None:
                return jc.get(_w("val"))
        sid = None
        if pPr is not None:
            ps = pPr.find(_w("pStyle"))
            sid = ps.get(_w("val")) if ps is not None else None
        seen = set()
        while sid and sid not in seen:
            seen.add(sid)
            st = self._style_by_id(sid)
            if st is None:
                break
            spr = st.find(_w("pPr"))
            if spr is not None:
                jc = spr.find(_w("jc"))
                if jc is not None:
                    return jc.get(_w("val"))
            b = st.find(_w("basedOn"))
            sid = b.get(_w("val")) if b is not None else None
        if self.styles is not None:
            dd = self.styles.find(_w("docDefaults"))
            pd = dd.find(_w("pPrDefault")) if dd is not None else None
            spr = pd.find(_w("pPr")) if pd is not None else None
            jc = spr.find(_w("jc")) if spr is not None else None
            if jc is not None:
                return jc.get(_w("val"))
        return None

    # --------------------------------------------------------------- 页序
    def sect_of_page(self, page):
        best = 0
        for s in sorted(self.spans):
            if self.spans[s][0] <= page:
                best = s
        return best

    def disp_num(self, i):
        """条目 i 所在页显示的页码（按强制分页推算，真实页码的下界）。"""
        if i < 0 or i >= len(self.els):
            return None
        s = self.sect_of[i]
        if s not in self.spans:
            return None
        return self.dstart[s] + (self.pages[i] - self.spans[s][0])

    def first_visible_of_sect(self, k):
        for i, s in enumerate(self.sect_of):
            if s == k and _el_visible(self.els[i]):
                return i
        return -1


# --------------------------------------------------------------- 标题定位
def _flow_exact(doc, blocks, want):
    """正文里文字恰为 want 的块级条目序号；优先取带大纲级别/标题样式的那个。"""
    key = _nospace(want)
    hits = [i for i, b in enumerate(blocks)
            if isinstance(b, Paragraph) and _nospace(_para_text(b)) == key]
    if not hits:
        return -1
    for i in hits:
        if _outline_level(doc, blocks[i]) is not None:
            return i
    return hits[-1]


def _flow_contains(blocks, want):
    key = _nospace(want)
    for i, b in enumerate(blocks):
        if isinstance(b, Paragraph) and key in _nospace(_para_text(b)):
            return i
    return -1


def _key_locations(doc, blocks):
    return {
        "cover": _flow_contains(blocks, "硕士学位论文"),
        "orig": _flow_contains(blocks, "原创性声明"),
        "auth": _flow_contains(blocks, "版权使用授权书"),
        "toc": _flow_exact(doc, blocks, "目录"),
        "abs_cn": _flow_exact(doc, blocks, "中文摘要"),
        "abs_en": _flow_exact(doc, blocks, "ABSTRACT"),
        "abbr": _flow_exact(doc, blocks, "中英文缩略词表"),
        "ref": _flow_exact(doc, blocks, "参考文献"),
    }


# ------------------------------------------------------- 18/19/20/21 判据
def _front_no_page_number(pm, loc):
    """封面页、原创性声明页、版权使用授权书页、目录页所在各节都不出页码。"""
    idxs = [loc[k] for k in ("cover", "orig", "auth", "toc")]
    if any(i < 0 for i in idxs):
        return False
    for k in sorted({pm.sect_of[i] for i in idxs}):
        if not pm.footer_no_number(k):
            return False
    return True


def _roman_front_ok(pm, loc):
    """中文摘要页 Ⅰ、ABSTRACT 页 Ⅱ：摘要所在节 upperRoman 且 start=1，
    摘要是该节首个可见条目，页脚靠 PAGE 域出页码，ABSTRACT 落在第 2 页。"""
    i_cn, i_en = loc["abs_cn"], loc["abs_en"]
    if i_cn < 0 or i_en <= i_cn:
        return False
    k = pm.sect_of[i_cn]
    if pm.fmt(k) != "upperRoman" or pm.start(k) != 1:
        return False
    if pm.first_visible_of_sect(k) != i_cn:
        return False
    if not pm.footer_page_field_ok(k):
        return False
    k2 = pm.sect_of[i_en]
    if pm.fmt(k2) != "upperRoman" or not pm.footer_page_field_ok(k2):
        return False
    return pm.disp_num(i_en) == 2


def _arabic_body_ok(pm, loc):
    """缩略词表页起阿拉伯页码、首页为 1、其后各节不再重启编号。"""
    i, r = loc["abbr"], loc["ref"]
    if i < 0 or r < i:
        return False
    k0 = pm.sect_of[i]
    if pm.fmt(k0) != "decimal" or pm.start(k0) != 1:
        return False
    if pm.first_visible_of_sect(k0) != i:
        return False
    for k in range(k0, pm.n_sect):
        if pm.fmt(k) != "decimal":
            return False
        if not pm.footer_page_field_ok(k):
            return False
        if k > k0:
            st = pm.start(k)
            if st is not None and st != pm.expected_start(k):
                return False
    return True


def _footer_centered_ok(pm, loc):
    """中文摘要页、正文中部任一页、末页所在节的页码段落都居中。"""
    i_cn, i_ab = loc["abs_cn"], loc["abbr"]
    if i_cn < 0 or i_ab < 0:
        return False
    last_page = max(pm.pages) if pm.pages else 0
    mid_page = (pm.pages[i_ab] + last_page) // 2
    ks = {pm.sect_of[i_cn], pm.sect_of[i_ab],
          pm.sect_of_page(mid_page), pm.sect_of_page(last_page)}
    for k in sorted(ks):
        if pm.footer_centered(k) is not True:
            return False
    return True


def _has_static_blank_page(pm, loc):
    """静态判定"整页空白页"：正文区内相邻两个强制分页之间没有任何可见内容。

    失去渲染手段后不再判断"排满后是否溢出成空白页"，只认可静态可验证的形式：
    连续手动分页符、或分页符与分节符之间只剩空段落。
    """
    i0 = loc["abbr"]
    if i0 < 0:
        return False
    seen = _el_visible(pm.els[i0])
    for i in range(i0 + 1, len(pm.els)):
        b = pm.bnd[i]
        if b >= 1:
            if (not seen) or b >= 2:
                return True
            seen = False
        if _el_visible(pm.els[i]):
            seen = True
    return False


# ------------------------------------------------------------- 03 目录页码
_TOC_TARGETS = ("前言", "第二部分 中医综合干预响应特征", "参考文献")


def _bookmark_positions(pm):
    """书签名 → 所在块级条目序号（书签挂在 body 上时取其后的第一个块级条目）。"""
    idx = {}
    for i, el in enumerate(pm.els):
        idx[id(el)] = i
    pos = {}
    for bs in pm.body.iter(_w("bookmarkStart")):
        nm = bs.get(_w("name"))
        if not nm or nm in pos:
            continue
        node = bs
        hit = None
        while node is not None:
            if id(node) in idx:
                hit = idx[id(node)]
                break
            node = node.getparent()
        if hit is None:
            nxt = bs.getnext()
            while nxt is not None and id(nxt) not in idx:
                nxt = nxt.getnext()
            hit = idx.get(id(nxt), -1) if nxt is not None else -1
        pos[nm] = hit
    return pos


def _toc_entry_title(txt, res):
    t = txt
    r = _nospace(res)
    if r and t.endswith(r):
        t = t[:len(t) - len(r)]
    return t.rstrip(".。…· \t")


def _toc_pageref_entries(pm):
    """含 PAGEREF 域的段落 → (条目序号, 标题, 书签名, 是否带 \\h, 域结果)。"""
    out = []
    for i, el in enumerate(pm.els):
        if el.tag != _w("p"):
            continue
        instr = "".join(t.text or "" for t in el.iter(_w("instrText")))
        for fs in el.iter(_w("fldSimple")):
            instr += " " + (fs.get(_w("instr")) or "")
        m = re.search(r"PAGEREF\s+(\S+)([^\x00]*?)(?=PAGEREF|$)", instr)
        if m is None:
            continue
        name = m.group(1)
        hyper = re.search(r"\\h", m.group(2)) is not None
        res = _field_result_text(el)
        out.append((i, _el_text(el), name, hyper, res))
    return out


def _bookmark_hits_title(blocks, bi, key):
    """书签所在处（含其后 2 个条目）是否就是那条标题，用于验证 Ctrl 单击的落点。"""
    for j in range(bi, min(bi + 3, len(blocks))):
        b = blocks[j]
        if isinstance(b, Paragraph) and _nospace(_para_text(b)) == key:
            return True
    return False


def _toc_pagerefs_ok(doc, blocks, pm, loc):
    """目录三条：PAGEREF \\h 域 + 书签落在同名标题上 + 所标页码不小于静态下界
    且随标题在正文中的先后严格递增。"""
    ents = _toc_pageref_entries(pm)
    if not ents:
        return False
    bms = _bookmark_positions(pm)
    picked = []
    for want in _TOC_TARGETS:
        key = _nospace(want)
        hit = None
        for ent in ents:
            if _toc_entry_title(ent[1], ent[4]) == key:
                hit = ent
                break
        if hit is None:
            return False
        _i, _txt, name, hyper, res = hit
        if not hyper:
            return False
        m = _ARABIC_ONLY.match(res)
        if m is None:
            return False
        num = int(m.group(1))
        bi = bms.get(name, -1)
        if bi < 0 or not _bookmark_hits_title(blocks, bi, key):
            return False
        lb = pm.disp_num(bi)
        if lb is None or num < lb:
            return False
        picked.append((bi, num))
    picked.sort()
    for a, b in zip(picked, picked[1:]):
        if b[1] <= a[1]:
            return False
    return True


# ---------------------------------------------------------------------------
# 维度一：可用或可修改性（任一不通过则总分为 0）
# ---------------------------------------------------------------------------

def _dim1_openable(file_path):
    if not file_path.lower().endswith(".docx"):
        return "文件扩展名不是 .docx"
    try:
        with open(file_path, "rb") as fh:
            head = fh.read(8)
    except OSError as exc:
        return "文件无法读取：%s" % exc
    if head.startswith(b"\xd0\xcf\x11\xe0"):
        return "文件为加密或二进制复合格式，打开时需要密码"
    if not head.startswith(b"PK"):
        return "文件不是有效的 Office Open XML 包"
    try:
        with zipfile.ZipFile(file_path) as zf:
            names = zf.namelist()
            if "EncryptedPackage" in names:
                return "文档设置了打开密码"
    except zipfile.BadZipFile:
        return "文档包结构损坏，无法在 Word 中直接打开"
    return None


def _dim1_readable(file_path, doc):
    names = _part_names(file_path)
    for required in ("[Content_Types].xml", "word/document.xml"):
        if required not in names:
            return "缺少必需部件 %s，打开时会提示需要修复" % required
    try:
        with zipfile.ZipFile(file_path) as zf:
            if zf.testzip() is not None:
                return "文档包内有损坏的部件，打开时会提示需要修复"
    except zipfile.BadZipFile:
        return "文档包结构损坏"
    xml = _doc_xml(file_path)
    rels = _part_xml(file_path, "word/_rels/document.xml.rels")
    rel_ids = set(re.findall(r'Id="([^"]+)"', rels))
    used = set(re.findall(r'r:(?:embed|id|link)="([^"]+)"', xml))
    missing = sorted(used - rel_ids)
    if missing:
        return "存在无法读取的引用关系 %s" % ", ".join(missing[:3])
    if not _all_paragraphs(doc):
        return "正文没有任何段落，内容无法正常显示"
    return None

def _dim1_editable(file_path, doc):
    custom = _part_xml(file_path, "docProps/custom.xml")
    if "_MarkAsFinal" in custom:
        return "文档被标记为最终状态"
    settings = _part_xml(file_path, "word/settings.xml")
    m = re.search(r"<w:documentProtection\b([^>]*)/?>", settings)
    if m:
        attrs = m.group(1)
        edit = re.search(r'w:edit="([^"]+)"', attrs)
        enforce = re.search(r'w:enforcement="([^"]+)"', attrs)
        on = enforce is None or str(enforce.group(1)).lower() not in ("0", "false", "off")
        if on and edit is not None and edit.group(1) in (
                "readOnly", "forms", "comments", "trackedChanges"):
            return "文档设置了修改密码或编辑限制（%s）" % edit.group(1)
    if re.search(r"<w:writeProtection\b", settings):
        return "文档设置了修改密码"
    xml = _doc_xml(file_path)
    boxed = 0
    for m in re.finditer(r"<w:txbxContent>(.*?)</w:txbxContent>", xml, re.S):
        boxed += sum(len(_nospace(t)) for t in
                     re.findall(r"<w:t[^>]*>(.*?)</w:t>", m.group(1), re.S))
    total = 0
    for p in _all_paragraphs(doc):
        total += len(_nospace(p.text))
    for tbl in _all_tables(doc):
        for row in tbl.rows:
            for cell in row.cells:
                total += len(_nospace(cell.text))
    if total - boxed < 200:
        return "正文可选中编辑的文字过少，疑为图片或文本框拼接"
    return None


def _check_dim1(file_path, doc):
    for fn in (_dim1_openable,):
        reason = fn(file_path)
        if reason:
            return (False, reason)
    for fn in (_dim1_readable, _dim1_editable):
        reason = fn(file_path, doc)
        if reason:
            return (False, reason)
    return (True, "")

def _table_map(blocks):
    """把文档中的表格映射为 封面信息表 / 缩略词表 / 表1..表13。

    优先用表格上方最近的非空段落（表题）定位，数量恰为 15 时再按出现顺序补齐。
    """
    tables = [b for b in blocks if isinstance(b, Table)]
    mapping = {}
    for bi, b in enumerate(blocks):
        if not isinstance(b, Table):
            continue
        cap = ""
        for j in range(bi - 1, -1, -1):
            if isinstance(blocks[j], Paragraph):
                t = _para_text(blocks[j])
                if t:
                    cap = t
                    break
        m = re.match(r"^表\s*(\d{1,2})\b", cap)
        if m:
            mapping.setdefault("表%d" % int(m.group(1)), b)
        elif "缩略词表" in cap:
            mapping.setdefault("缩略词表", b)
    if len(tables) == len(TABLE_NAMES):
        for name, tbl in zip(TABLE_NAMES, tables):
            mapping.setdefault(name, tbl)
    elif tables:
        mapping.setdefault("封面信息表", tables[0])
    return mapping


def _row_cells(tbl, idx):
    try:
        return list(tbl.rows[idx].cells)
    except (IndexError, AttributeError):
        return []


def _cell_paras(cell):
    return list(cell.paragraphs)


def _cell_font_problems(doc, cell, ea_allowed, ascii_allowed, size_pt,
                        require_bold=None):
    problems = []
    for p in _cell_paras(cell):
        if not _text_runs(p):
            continue
        problems.extend(_para_font_problems(doc, p, ea_allowed, ascii_allowed,
                                            size_pt, require_bold))
    return problems

# ---------------------------------------------------------------------------
# 对外接口
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
            raise FileNotFoundError("目录不存在：%s" % dir_path)
        cands = sorted(
            f for f in os.listdir(dir_path)
            if f.lower().endswith(".docx") and not f.startswith("~$")
        )
        if not cands:
            raise FileNotFoundError("目录内未找到 .docx 文件：%s" % dir_path)
        file_path = os.path.join(dir_path, cands[0])
        result["file_name"] = os.path.basename(file_path)
        return _evaluate_impl(file_path, result)
    except Exception as exc:
        # status 为 error 时未能完成判定，dim1_pass 留空，dim1_reason 也为空
        result["status"] = "error"
        result["error"] = "%s: %s" % (type(exc).__name__, exc)
        result["dim1_pass"] = ""
        result["dim1_reason"] = ""
        result["dim2_items"] = []
        result["total_score"] = 0
        return result

def _leading_runs(para, n):
    """覆盖段落前 n 个可见字符的 run 列表。"""
    out = []
    count = 0
    for r in _text_runs(para):
        out.append(r)
        count += len(_nospace(r.text))
        if count >= n:
            break
    return out


def _trailing_runs(para, n):
    """跳过覆盖前 n 个可见字符的 run 后剩余的 run 列表。"""
    lead = _leading_runs(para, n)
    keep = {id(r._element) for r in lead}
    return [r for r in _text_runs(para) if id(r._element) not in keep]


def _caption_index(paras, prefix, maxlen=80):
    """定位以"图N""表N"开头的图题/表题段落（排除正文中的引用行）。"""
    pfx = _norm(prefix)
    for i, p in enumerate(paras):
        t = _para_text(p)
        if not t.startswith(pfx) or len(t) > maxlen:
            continue
        rest = t[len(pfx):]
        if rest and rest[0].isdigit():
            continue  # 避免"表1"误配"表13"
        return i
    return -1


def _is_chromatic(val):
    if not val or str(val).lower() == "auto":
        return False
    if not re.fullmatch(r"[0-9A-Fa-f]{6}", str(val)):
        return False
    r, g, b = int(val[0:2], 16), int(val[2:4], 16), int(val[4:6], 16)
    return (max(r, g, b) - min(r, g, b)) >= 60 and max(r, g, b) >= 90

# ---------------------------------------------------------------------------
# 维度二：完成度
# ---------------------------------------------------------------------------

def _evaluate_impl(file_path, result):
    # 交付件损坏（扩展名不对、加密、包结构坏、打不开）一律按维度一不通过返回，
    # 不让 python-docx 的异常冒泡成 status=error。
    open_reason = _dim1_openable(file_path)
    if open_reason:
        result["dim1_pass"] = False
        result["dim1_reason"] = open_reason
        result["dim2_items"] = []
        result["total_score"] = 0
        result["max_score"] = MAX_SCORE_TOTAL
        return result
    try:
        doc = _open_word_document(file_path)
    except Exception as exc:  # noqa: BLE001
        result["dim1_pass"] = False
        result["dim1_reason"] = "文档无法打开，打开时会提示需要修复：%s: %s" % (
            type(exc).__name__, exc)
        result["dim2_items"] = []
        result["total_score"] = 0
        result["max_score"] = MAX_SCORE_TOTAL
        return result

    dim1_ok, dim1_reason = _check_dim1(file_path, doc)
    result["dim1_pass"] = bool(dim1_ok)
    result["dim1_reason"] = dim1_reason
    if not dim1_ok:
        result["dim2_items"] = []
        result["total_score"] = 0
        result["max_score"] = MAX_SCORE_TOTAL
        return result

    xml = _doc_xml(file_path)
    blocks = list(_iter_block_items(doc))
    paras = [b for b in blocks if isinstance(b, Paragraph)]
    tables = [b for b in blocks if isinstance(b, Table)]
    texts = [_para_text(p) for p in paras]
    tmap = _table_map(blocks)
    pm = _PageModel(doc, file_path, blocks)
    loc = _key_locations(doc, blocks)

    # 段落元素 → 块序号，用于"图题在图下方""表题在表上方"这类位置判定
    blk_of = {}
    for k, b in enumerate(blocks):
        if isinstance(b, Paragraph):
            blk_of[id(b._element)] = k

    bonus = []
    penalty = []

    def add_bonus(rule, hit):
        d = _MAX_DELTAS_ORDER[len(bonus)]
        bonus.append("+%d：%s" % (d if hit else 0, rule))

    def add_penalty(rule, bad):
        d = _PENALTY_DELTAS_ORDER[len(penalty)]
        penalty.append("-%d：%s" % (d if bad else 0, rule))


    # 01 +5 目录为可自动更新的目录域
    add_bonus("目录页的目录为可自动更新的目录域，光标置于目录内按 F9 可弹出"
              "\"更新目录\"对话框", _has_toc_field(file_path))

    # 02 +5 七个一级标题样式为标题1、十二个二级标题样式为标题2
    style_ok = True
    for t in L1_TITLES:
        if not any(_is_heading_style(doc, paras[i], 1) for i in _find_paras(paras, t)):
            style_ok = False
            break
    if style_ok:
        for t in L2_TITLES:
            if not any(_is_heading_style(doc, paras[i], 2) for i in _find_paras(paras, t)):
                style_ok = False
                break
    add_bonus("中英文缩略词表、前言、第一部分 多模态康复资料库构建、第二部分 中医综合干预"
              "响应特征、第三部分 疗效预测模型建立与解释、讨论、结论共 7 个标题的段落样式"
              "均为「标题 1」，一、资料来源与研究设计、二、变量清洗与质量控制、三、统计方法"
              "与流程、四、基线资料分析、五、小结、一、证候要素与干预路径、二、疗效变化与"
              "阶段性表现、三、亚组分析与安全性观察、四、小结、一、建模数据集与特征工程、"
              "二、模型性能与校准评价、三、关键变量解释与应用共 12 个标题的段落样式均为"
              "「标题 2」", style_ok)

    # 03 +3 目录三条页码与标题实际所在页显示页码一致
    #        静态判据：PAGEREF _Toc… \h 域结果 + 书签落点 + 页码单调递增
    toc_page_ok = _toc_pagerefs_ok(doc, blocks, pm, loc)
    add_bonus("目录中\"前言\"\"第二部分 中医综合干预响应特征\"\"参考文献\"三条右侧均标有"
              "页码，且按住 Ctrl 单击跳转后标题所在页与所标页码一致", toc_page_ok)

    # 04 +1 目录条目四号黑体
    toc_idxs = _toc_style_paras(doc, paras)
    if not toc_idxs:
        t0 = _find_para(paras, "目录")
        if t0 >= 0:
            toc_idxs = [i for i in range(t0 + 1, _next_heading_index(doc, paras, t0))
                        if _text_runs(paras[i])]
    toc_font_ok = bool(toc_idxs)
    for i in toc_idxs:
        if not _text_runs(paras[i]):
            continue
        if _para_font_problems(doc, paras[i], HEITI, None, PT_SIHAO):
            toc_font_ok = False
            break
    add_bonus("目录条目文字字号框显示为四号（14 磅）、字体为黑体", toc_font_ok)

    # 05 +1 封面中文题目与信息表两行小二楷体
    cover_ok = False
    ti = _find_para(paras, COVER_TITLE_CN)
    info = tmap.get("封面信息表")
    if ti >= 0 and info is not None and len(info.rows) >= 1:
        cover_ok = not _para_font_problems(doc, paras[ti], KAITI, None, PT_XIAOER)
        row_ok = False
        for r in range(len(info.rows)):
            cells = _row_cells(info, r)
            if any("学院名称" in _nospace(c.text) for c in cells):
                row_ok = not any(_cell_font_problems(doc, c, KAITI, None, PT_XIAOER)
                                 for c in cells)
                break
        last_ok = not any(_cell_font_problems(doc, c, KAITI, None, PT_XIAOER)
                          for c in _row_cells(info, len(info.rows) - 1))
        cover_ok = cover_ok and row_ok and last_ok
    add_bonus("封面中文题目\"基于多模态康复记录与机器学习研究中医综合干预冻结肩的疗效特征"
              "及预测因素\"、封面信息表\"学院名称\"行、封面信息表最后一行文字字号框显示为"
              "小二（18 磅）、字体为楷体", cover_ok)

    # 06 +1 封面英文题目小二 Times New Roman
    ei = _find_para_startswith(paras, COVER_TITLE_EN_HEAD)
    add_bonus("封面英文题目\"A Multimodal Rehabilitation Record and Machine Learning...\""
              "文字字号框显示为小二（18 磅）、字体为 Times New Roman",
              ei >= 0 and not _para_font_problems(doc, paras[ei], None, TNR, PT_XIAOER))

    # 07 +1 关键词前空一行且"关键词"四号黑体
    kw_ok = False
    ki = _find_para_startswith(paras, "关键词")
    if ki >= 2 and _is_blank(paras[ki - 1]) and _starts_with_lead(_para_text(paras[ki - 2]), "结论"):
        lead = _leading_runs(paras[ki], 3)
        kw_ok = bool(lead) and all(
            _font_ok(_resolve_effective_eastasia(doc, r), HEITI)
            and _near(_resolve_effective_size_pt(doc, r), PT_SIHAO) for r in lead)
    add_bonus("中文摘要\"结论\"段与\"关键词\"段之间有 1 个空段落，且\"关键词\"三字字号框"
              "显示为四号（14 磅）、字体为黑体", kw_ok)

    # 08 +1 Subject words 前空一行且四号 TNR 加粗
    sw_ok = False
    si = _find_para_startswith(paras, "Subject words")
    if si >= 2 and _is_blank(paras[si - 1]) and _starts_with_lead(_para_text(paras[si - 2]), "Conclusion"):
        lead = _leading_runs(paras[si], 12)
        sw_ok = bool(lead) and all(
            _font_ok(_resolve_effective_ascii(doc, r), TNR)
            and _near(_resolve_effective_size_pt(doc, r), PT_SIHAO)
            and _resolve_effective_bold(doc, r) for r in lead)
    add_bonus("ABSTRACT 中\"Conclusion\"段与\"Subject words\"段之间有 1 个空段落，且"
              "\"Subject words\"字号框显示为四号（14 磅）、字体为 Times New Roman、加粗",
              sw_ok)

    # 09 +3 三个二级标题小四黑体
    l2_font_ok = True
    for t in ("一、资料来源与研究设计", "二、疗效变化与阶段性表现", "三、关键变量解释与应用"):
        idxs = _find_paras(paras, t)
        if not idxs or all(_para_font_problems(doc, paras[i], HEITI, None, PT_XIAOSI)
                           for i in idxs):
            l2_font_ok = False
            break
    add_bonus("二级标题\"一、资料来源与研究设计\"\"二、疗效变化与阶段性表现\"\"三、关键变量"
              "解释与应用\"文字字号框显示为小四（12 磅）、字体为黑体", l2_font_ok)

    # 10 +1 图1图4图6 图题单倍行距
    fig_sp_ok = True
    for name in ("图1", "图4", "图6"):
        i = _caption_index(paras, name)
        if i < 0 or not _is_single_spacing(doc, paras[i]):
            fig_sp_ok = False
            break
    add_bonus("图 1、图 4、图 6 的图题段落在段落对话框中行距显示为单倍行距", fig_sp_ok)

    # 11 +3 表1表7表13 表题单倍行距居中宋体
    cap_ok = True
    for name in ("表1", "表7", "表13"):
        i = _caption_index(paras, name)
        if i < 0:
            cap_ok = False
            break
        p = paras[i]
        if not _is_single_spacing(doc, p) or not _is_centered(doc, p):
            cap_ok = False
            break
        var_a = not _para_font_problems(doc, p, SONGTI, None, PT_XIAOSI, True)
        var_b = not _para_font_problems(doc, p, SONGTI, None, PT_WUHAO)
        if not (var_a or var_b):
            cap_ok = False
            break
    add_bonus("表 1、表 7、表 13 的表题段落在段落对话框中行距显示为单倍行距、对齐方式为居中，"
              "文字字体为宋体、字号框显示为小四（12 磅）加粗或五号（10.5 磅）", cap_ok)

    # 12 +5 表1表7表13 表头行下方横线 0.75 磅
    hline_ok = True
    for name in ("表1", "表7", "表13"):
        tbl = tmap.get(name)
        if tbl is None or len(tbl.rows) < 2:
            hline_ok = False
            break
        row0 = _row_cells(tbl, 0)
        row1 = _row_cells(tbl, 1)
        if not row0:
            hline_ok = False
            break
        for j, c in enumerate(row0):
            val, pt = _effective_edge(doc, tbl, c, "bottom", "insideH")
            if val is None and j < len(row1):
                val, pt = _effective_edge(doc, tbl, row1[j], "top", "insideH")
            if not _has_line(val) or not _near(pt, 0.75, 0.05):
                hline_ok = False
                break
        if not hline_ok:
            break
    add_bonus("表 1、表 7、表 13 的表头行下方横线在边框对话框中宽度显示为 0.75 磅", hline_ok)

    # 13 +3 缩略词表表1表13 表头行五号加粗
    head_ok = True
    for name in ("缩略词表", "表1", "表13"):
        tbl = tmap.get(name)
        if tbl is None:
            head_ok = False
            break
        cells = _row_cells(tbl, 0)
        if not cells or any(_cell_font_problems(doc, c, None, None, PT_WUHAO, True)
                            for c in cells):
            head_ok = False
            break
    add_bonus("中英文缩略词表、表 1、表 13 的表头行文字字号框显示为五号（10.5 磅）、加粗",
              head_ok)

    # 14 +3 表1表7表13 第2行小五
    row2_ok = True
    for name in ("表1", "表7", "表13"):
        tbl = tmap.get(name)
        if tbl is None or len(tbl.rows) < 2:
            row2_ok = False
            break
        cells = _row_cells(tbl, 1)
        if not cells or any(_cell_font_problems(doc, c, None, None, PT_XIAOWU)
                            for c in cells):
            row2_ok = False
            break
    add_bonus("表 1、表 7、表 13 的第 2 行文字字号框显示为小五（9 磅）", row2_ok)

    # 参考文献条目（加分项 15 与减分项 10、16 共用）
    ref_entries = []
    ri = _find_para(paras, "参考文献")
    if ri >= 0:
        for i in range(ri + 1, len(paras)):
            t = texts[i]
            if not t:
                continue
            if _outline_level(doc, paras[i]) is not None:
                break
            ref_entries.append((i, paras[i], t))

    # 15 +3 参考文献序号 [1] 至 [15] 连续
    lits = [re.match(r"^\[\s*(\d+)\s*\]", t) for (_i, _p, t) in ref_entries]
    if ref_entries and all(lits):
        num_ok = [int(m.group(1)) for m in lits] == list(range(1, 16))
    elif ref_entries:
        num_ok = len(ref_entries) == 15
        nids = []
        for (_i, p, _t) in ref_entries:
            lt = _num_lvl_text(doc, p)
            if not lt or not re.fullmatch(r"\[\s*%\d\s*\]", lt):
                num_ok = False
                break
            npr = _num_pr(p)
            ni = npr.find(qn("w:numId")) if npr is not None else None
            nids.append(ni.get(qn("w:val")) if ni is not None else None)
        if num_ok and (len(set(nids)) != 1 or None in nids):
            num_ok = False
    else:
        num_ok = False
    add_bonus("参考文献 15 条的序号为半角方括号 [1] 至 [15]，连续无缺号无重号", num_ok)

    # 16 +3 图2图3图5 图题单倍行距、五号宋体、居中
    fig2_ok = True
    for name in ("图2", "图3", "图5"):
        i = _caption_index(paras, name)
        if i < 0:
            fig2_ok = False
            break
        p = paras[i]
        if (not _is_single_spacing(doc, p) or not _is_centered(doc, p)
                or _para_font_problems(doc, p, SONGTI, None, PT_WUHAO)):
            fig2_ok = False
            break
    add_bonus("图 2、图 3、图 5 的图题段落在段落对话框中行距显示为单倍行距、对齐方式为居中，"
              "文字字体为宋体、字号框显示为五号（10.5 磅）", fig2_ok)

    # 17 +1 页边距 上2.5 下2.5 左3 右2.5 厘米
    margins = re.findall(r"<w:pgMar\b[^>]*/?>", xml)
    margin_ok = bool(margins)
    for mar in margins:
        for attr, want in (("top", 2.5), ("bottom", 2.5), ("left", 3.0), ("right", 2.5)):
            mm = re.search(r'w:%s="(-?\d+)"' % attr, mar)
            v = float(mm.group(1)) / TWIP_CM if mm else None
            if v is None or abs(v - want) > 0.03:
                margin_ok = False
                break
        if not margin_ok:
            break
    add_bonus("页面设置对话框中页边距显示为上 2.5 厘米、下 2.5 厘米、左 3 厘米、右 2.5 厘米",
              margin_ok)

    # 18 +5 封面声明授权目录页脚无页码且节数不少于 3
    # 节数按 w:sectPr 元素统计：有的交付件把命名空间前缀写成 ns0:，
    # 对原文做正则会数到 0，必须走已解析好的分节模型。
    sect_count = pm.n_sect
    front_ok = sect_count >= 3 and _front_no_page_number(pm, loc)
    add_bonus("封面页、原创性声明页、版权使用授权书页、目录页的页脚均无页码，且状态栏显示的"
              "节数不少于 3", front_ok)

    # 19 +3 中文摘要页 Ⅰ、ABSTRACT 页 Ⅱ
    roman_ok = _roman_front_ok(pm, loc)
    add_bonus("中文摘要页页脚页码显示为罗马数字 Ⅰ，ABSTRACT 页页脚页码显示为罗马数字 Ⅱ",
              roman_ok)

    # 20 +3 缩略词表页至末页阿拉伯页码首页为 1 连续
    arabic_ok = _arabic_body_ok(pm, loc)
    add_bonus("中英文缩略词表页至参考文献末页的页脚页码显示为阿拉伯数字，首页为 1，全程连续"
              "无缺号无重号", arabic_ok)

    # 21 +1 摘要页、正文中部、末页页脚居中
    center_ok = _footer_centered_ok(pm, loc)
    add_bonus("中文摘要页、正文中部任一页、末页的页脚段落对齐方式均为居中", center_ok)

    # ---- 减分项 ----
    BODY_PREFIXES = (
        "冻结肩又称粘连性肩关节囊炎",
        "随访结果显示，疼痛评分在2至4周下降较快",
        "本研究以冻结肩真实世界康复资料为基础",
    )
    body_idxs = [_find_para_startswith(paras, pfx) for pfx in BODY_PREFIXES]

    # 01 -3 三处正文段中文宋体小四、英文数字 TNR 小四
    bad = False
    for i in body_idxs:
        if i >= 0 and _para_font_problems(doc, paras[i], SONGTI, TNR, PT_XIAOSI):
            bad = True
            break
    add_penalty("前言下第 1 段（起首\"冻结肩又称粘连性肩关节囊炎\"）、第二部分\"二、疗效变化与"
                "阶段性表现\"下第 1 段（起首\"随访结果显示，疼痛评分在2至4周下降较快\"）、讨论"
                "下第 1 段（起首\"本研究以冻结肩真实世界康复资料为基础\"）的中文文字字体非宋体"
                "或字号框显示非小四（12 磅），或英文与数字字体非 Times New Roman 或字号框显示"
                "非小四（12 磅）", bad)

    # 02 -3 三处正文段 1.5 倍行距、首行缩进 2 字符
    bad = False
    for i in body_idxs:
        if i < 0:
            continue
        ok_ind, _desc = _first_line_indent_desc(doc, paras[i])
        if not _is_15_spacing(doc, paras[i]) or not ok_ind:
            bad = True
            break
    add_penalty("前言下第 1 段、第二部分\"二、疗效变化与阶段性表现\"下第 1 段、讨论下第 1 段在"
                "段落对话框中行距显示非 1.5 倍行距，或特殊格式非首行缩进，或缩进值显示非 2 字符",
                bad)

    # 03 -1 三个一级标题黑体四号
    bad = False
    for t in ("前言", "第二部分 中医综合干预响应特征", "参考文献"):
        idxs = _find_paras(paras, t)
        if idxs and all(_para_font_problems(doc, paras[i], HEITI, None, PT_SIHAO)
                        for i in idxs):
            bad = True
            break
    add_penalty("一级标题\"前言\"\"第二部分 中医综合干预响应特征\"\"参考文献\"的文字字体非黑体"
                "或字号框显示非四号（14 磅）", bad)

    cn_abs_i = _find_para(paras, "中文摘要")
    en_abs_i = _find_para(paras, "ABSTRACT")
    abbr_i = _find_para(paras, "中英文缩略词表")
    cn_end = en_abs_i if en_abs_i > cn_abs_i else len(paras)
    en_end = abbr_i if abbr_i > en_abs_i else len(paras)

    def _seek(prefix, lo, hi):
        """在区间内找以引导词开头的段落，兼容"[目的]"这类方括号写法。"""
        if lo < 0:
            return -1
        for i in range(max(0, lo), min(hi if hi > 0 else len(paras), len(paras))):
            if _starts_with_lead(texts[i], prefix):
                return i
        return -1

    # 04 -1 中文摘要标题与四段正文字体字号、无多余加粗与大字号
    bad = False
    if cn_abs_i >= 0:
        if _para_font_problems(doc, paras[cn_abs_i], HEITI, None, PT_SIHAO):
            bad = True
        for lead in ("目的", "方法", "结果", "结论"):
            j = _seek(lead, cn_abs_i + 1, cn_end)
            if j < 0:
                continue
            if _para_font_problems(doc, paras[j], SONGTI, None, PT_XIAOSI):
                bad = True
                break
            for r in _trailing_runs(paras[j], _lead_len(texts[j], lead)):
                if _resolve_effective_bold(doc, r):
                    bad = True
                    break
            for r in _text_runs(paras[j]):
                sz = _resolve_effective_size_pt(doc, r)
                if _near(sz, PT_SIHAO) or _near(sz, PT_XIAOER):
                    bad = True
                    break
            if bad:
                break
    add_penalty("中文摘要标题字体非黑体或字号框显示非四号（14 磅），或\"目的\"\"方法\"\"结果\""
                "\"结论\"四段正文字体非宋体、字号框显示非小四（12 磅），或四段中除段首引导词外"
                "出现加粗文字，或四段内出现字号框显示为四号（14 磅）或小二（18 磅）的文字", bad)

    # 05 -1 摘要分段引导词、关键词与 Subject words 内容格式
    KW_EXPECT = "冻结肩；中医康复；多模态记录；机器学习；疗效预测"
    TAIL_PUNCT = "。．.,，；;、!！?？:："
    bad = False
    if cn_abs_i >= 0:
        for lead in ("目的", "方法", "结果", "结论"):
            if _seek(lead, cn_abs_i + 1, cn_end) < 0:
                bad = True
                break
    if not bad and en_abs_i >= 0:
        for lead in ("Objective", "Methods", "Results", "Conclusion"):
            if _seek(lead, en_abs_i + 1, en_end) < 0:
                bad = True
                break
    if not bad:
        ki2 = _find_para_startswith(paras, "关键词")
        if ki2 >= 0:
            body = re.sub(r"^关键词\s*[:：]?\s*", "", texts[ki2])
            if _nospace(body) != _nospace(KW_EXPECT):
                bad = True
            elif body.rstrip()[-1:] in TAIL_PUNCT:
                bad = True
            else:
                for r in _trailing_runs(paras[ki2], 3):
                    if not _font_ok(_resolve_effective_eastasia(doc, r), SONGTI) \
                            or not _near(_resolve_effective_size_pt(doc, r), PT_XIAOSI):
                        bad = True
                        break
    if not bad:
        si2 = _find_para_startswith(paras, "Subject words")
        if si2 >= 0:
            body = re.sub(r"^Subject\s+words\s*[:：]?\s*", "", texts[si2])
            if ";" not in body and "；" not in body:
                bad = True
            elif body.rstrip()[-1:] in TAIL_PUNCT:
                bad = True
            else:
                for r in _trailing_runs(paras[si2], 13):
                    if not _font_ok(_resolve_effective_ascii(doc, r), TNR) \
                            or not _near(_resolve_effective_size_pt(doc, r), PT_XIAOSI):
                        bad = True
                        break
    add_penalty("中文摘要缺\"目的\"\"方法\"\"结果\"\"结论\"四个分段引导词，或 ABSTRACT 缺 "
                "Objective、Methods、Results、Conclusion 四个分段引导词，或\"关键词\"后内容非"
                "\"冻结肩；中医康复；多模态记录；机器学习；疗效预测\"、字体非宋体、字号框显示非"
                "小四（12 磅）、未用分号相隔、末词后有标点，或\"Subject words\"后关键词字体非 "
                "Times New Roman、字号框显示非小四（12 磅）、未用分号相隔、末词后有标点", bad)

    # 06 -1 ABSTRACT 标题行及其下英文题目行 TNR 四号加粗
    bad = False
    if en_abs_i >= 0:
        if _para_font_problems(doc, paras[en_abs_i], None, TNR, PT_SIHAO, True):
            bad = True
        else:
            j = en_abs_i + 1
            while j < len(paras) and _is_blank(paras[j]):
                j += 1
            if j < len(paras) and _para_font_problems(doc, paras[j], None, TNR,
                                                     PT_SIHAO, True):
                bad = True
    add_penalty("ABSTRACT 标题行及其下英文题目行的文字字体非 Times New Roman 或字号框显示非"
                "四号（14 磅）或未加粗", bad)

    # 07 -3 图2图5 图题、表4表11 表题的编号位置对齐
    def _prev_visible_block(k):
        for k2 in range(k - 1, -1, -1):
            b = blocks[k2]
            if isinstance(b, Table):
                return b
            if _para_text(b) or _has_drawing(b):
                return b
        return None

    def _next_visible_block(k):
        for k2 in range(k + 1, len(blocks)):
            b = blocks[k2]
            if isinstance(b, Table):
                return b
            if _para_text(b) or _has_drawing(b):
                return b
        return None

    bad = False
    for name in ("图2", "图5"):
        i = _caption_index(paras, name)
        if i < 0:
            bad = True
            break
        if not _is_centered(doc, paras[i]):
            bad = True
            break
        prev = _prev_visible_block(blk_of.get(id(paras[i]._element), 0))
        if not (isinstance(prev, Paragraph) and _has_drawing(prev)):
            bad = True
            break
    if not bad:
        for name in ("表4", "表11"):
            i = _caption_index(paras, name)
            if i < 0 or not _is_centered(doc, paras[i]):
                bad = True
                break
            nxt = _next_visible_block(blk_of.get(id(paras[i]._element), 0))
            if not isinstance(nxt, Table):
                bad = True
                break
    add_penalty("图 2、图 5 的图题编号非\"图2\"\"图5\"，或图题未位于图片下方，或图题段落对齐"
                "方式非居中，或表 4、表 11 的表题编号非\"表4\"\"表11\"，或表题未位于表格上方，"
                "或表题段落对齐方式非居中", bad)

    # 08 -5 缩略词表表7表13 三线表线型
    bad = False
    for name in ("缩略词表", "表7", "表13"):
        tbl = tmap.get(name)
        if tbl is None:
            continue
        nrows = len(tbl.rows)
        if nrows < 2:
            continue
        for c in _row_cells(tbl, 0):
            val, pt = _effective_edge(doc, tbl, c, "top", "top")
            if not _has_line(val) or not _near(pt, 1.5, 0.05):
                bad = True
        for c in _row_cells(tbl, nrows - 1):
            val, pt = _effective_edge(doc, tbl, c, "bottom", "bottom")
            if not _has_line(val) or not _near(pt, 1.5, 0.05):
                bad = True
        for r in range(nrows):
            for c in _row_cells(tbl, r):
                for en in ("left", "right"):
                    val, _pt = _effective_edge(doc, tbl, c, en, en)
                    if _has_line(val):
                        bad = True
        val, _pt = _edge(_tbl_borders_el(doc, tbl), "insideV")
        if _has_line(val):
            bad = True
        for r in range(1, nrows - 1):
            for c in _row_cells(tbl, r):
                val, _pt = _effective_edge(doc, tbl, c, "bottom", "insideH")
                if _has_line(val):
                    bad = True
        for r in range(2, nrows):
            for c in _row_cells(tbl, r):
                val, _pt = _effective_edge(doc, tbl, c, "top", "insideH")
                if _has_line(val):
                    bad = True
        if bad:
            break
    add_penalty("中英文缩略词表、表 7、表 13 出现竖线、左右外框线或除表头行下方横线以外的中间"
                "横线，或顶线、底线在边框对话框中宽度显示非 1.5 磅", bad)

    # 09 -3 缩略词表表9表11 单元格中英文字体
    bad = False
    for name in ("缩略词表", "表9", "表11"):
        tbl = tmap.get(name)
        if tbl is None:
            continue
        for row in tbl.rows:
            for cell in row.cells:
                if _cell_font_problems(doc, cell, SONGTI, TNR, None):
                    bad = True
                    break
            if bad:
                break
        if bad:
            break
    add_penalty("中英文缩略词表、表 9、表 11 单元格内中文字体非宋体，或英文与数字字体非 "
                "Times New Roman", bad)

    # 10 -1 参考文献第 3、8、15 条字体字号
    bad = False
    for k in (3, 8, 15):
        if len(ref_entries) >= k:
            if _para_font_problems(doc, ref_entries[k - 1][1], SONGTI, TNR, PT_XIAOSI):
                bad = True
                break
    add_penalty("参考文献第 3、8、15 条的中文字体非宋体或字号框显示非小四（12 磅），或英文与"
                "数字字体非 Times New Roman 或字号框显示非小四（12 磅）", bad)

    # 11 -3 图1图3图6 嵌入型且居中
    bad = False
    for name in ("图1", "图3", "图6"):
        i = _caption_index(paras, name)
        if i < 0:
            bad = True
            break
        prev = _prev_visible_block(blk_of.get(id(paras[i]._element), 0))
        if not (isinstance(prev, Paragraph) and _has_drawing(prev)):
            bad = True
            break
        if _drawing_wrap(prev) != "inline" or not _is_centered(doc, prev):
            bad = True
            break
    add_penalty("图 1、图 3、图 6 的布局选项中环绕方式显示非嵌入型，或图片所在段落对齐方式非"
                "居中", bad)

    # 12 -1 页眉区出现文字或横线
    bad = False
    for part in _part_names(file_path):
        if not re.match(r"^word/header\d*\.xml$", part):
            continue
        hx = _part_xml(file_path, part)
        if any(_nospace(t) for t in re.findall(r"<w:t[^>]*>(.*?)</w:t>", hx, re.S)):
            bad = True
            break
        for bdr in re.findall(r"<w:pBdr>(.*?)</w:pBdr>", hx, re.S):
            if re.search(r'w:val="(?!nil"|none")[a-zA-Z]+"', bdr):
                bad = True
                break
        if bad:
            break
    add_penalty("任意页面的页眉区出现文字或横线", bad)

    # 13 -5 32 个标题齐全、标题下有正文、无截断
    need = {}
    for t in ALL_TITLES:
        need[t] = need.get(t, 0) + 1
    bad = any(texts.count(t) < n for t, n in need.items())
    if not bad:
        heads = [(k, _title_level(_para_text(b))) for k, b in enumerate(blocks)
                 if isinstance(b, Paragraph) and _para_text(b) in need]
        for hi, (k, lvl) in enumerate(heads):
            end = len(blocks)
            for k2, lvl2 in heads[hi + 1:]:
                if lvl2 <= lvl:
                    end = k2
                    break
            found = False
            for k2 in range(k + 1, end):
                b2 = blocks[k2]
                if isinstance(b2, Table):
                    found = True
                    break
                t2 = _para_text(b2)
                if t2 in need:
                    continue  # 下级标题不算正文，继续向后找
                if t2 or _has_drawing(b2):
                    found = True
                    break
            if not found:
                bad = True
                break
    if not bad:
        for i, p in enumerate(paras):
            t = texts[i]
            if len(t) < 30 or t in need:
                continue
            if _outline_level(doc, p) is not None or _has_drawing(p):
                continue
            if re.match(r"^[图表]\d", t) or re.match(r"^\[\d+\]", t):
                continue
            if t[-1] in "，,、；;":
                bad = True
                break
    add_penalty("全文缺少中文摘要、ABSTRACT、中英文缩略词表、前言、前言（续）、第一部分 多模态"
                "康复资料库构建、一、资料来源与研究设计、一、资料来源与研究设计（续）、二、变量"
                "清洗与质量控制、三、统计方法与流程、四、基线资料分析、四、基线资料分析（续）、"
                "四、基线资料分析（续）、五、小结、第二部分 中医综合干预响应特征、一、证候要素"
                "与干预路径、一、证候要素与干预路径（续）、二、疗效变化与阶段性表现、二、疗效"
                "变化与阶段性表现（续）、三、亚组分析与安全性观察、四、小结、第三部分 疗效预测"
                "模型建立与解释、一、建模数据集与特征工程、二、模型性能与校准评价、二、模型性能"
                "与校准评价（续）、三、关键变量解释与应用、三、关键变量解释与应用（续）、讨论、"
                "讨论（续）、讨论（续）、结论、参考文献共 32 个标题中的任一项，或任一标题下无"
                "正文文字，或出现句子中途截断、段落只剩半句", bad)

    # 14 -5 表格数不少于 15、图片数不少于 8
    n_pic = len(re.findall(r"<w:drawing\b", xml)) + len(re.findall(r"<w:pict\b", xml))
    add_penalty("全文表格数少于 15 个，或图片数少于 8 张（封面 2 张与图 1 至图 6 共 6 张）",
                len(tables) < 15 or n_pic < 8)

    # 15 -5 15 个表格行列数、无空单元格、无错位
    bad = len(tables) != len(TABLE_SHAPES)
    if not bad:
        for (name, shape), tbl in zip(zip(TABLE_NAMES, TABLE_SHAPES), tables):
            try:
                n_col = len(tbl.columns)
            except Exception:
                n_col = 0
            if len(tbl.rows) != shape[0] or n_col != shape[1]:
                bad = True
                break
            if name == "封面信息表":
                continue  # 封面信息表为版式表格，不参与空单元格判定
            for row in tbl.rows:
                if any(not _nospace(c.text) for c in row.cells):
                    bad = True
                    break
            if bad:
                break
    add_penalty("15 个表格的行列数与封面信息表 4 行 2 列、中英文缩略词表 16 行 3 列、表 1 为 "
                "7 行 2 列、表 2 为 5 行 2 列、表 3 为 6 行 4 列、表 4 为 8 行 3 列、表 5 为 "
                "7 行 3 列、表 6 为 5 行 3 列、表 7 为 5 行 4 列、表 8 为 6 行 5 列、表 9 为 "
                "7 行 3 列、表 10 为 5 行 4 列、表 11 为 6 行 5 列、表 12 为 4 行 3 列、表 13 "
                "为 5 行 3 列不一致，或出现空单元格，或数据与表头列错位", bad)

    # 16 -3 参考文献第 3、8、15 条著录内容
    bad = False
    for k, fields in sorted(REF_FIELDS.items()):
        if len(ref_entries) < k:
            bad = True
            break
        entry = _nospace(ref_entries[k - 1][2])
        if any(_nospace(f) not in entry for f in fields):
            bad = True
            break
    add_penalty("参考文献第 3 条的著者、题名、刊名、年、卷期、起止页非\"顾星辞,孟昭屹,樊朔,等\""
                "与\"温经活络方调理腰椎劳损病症临床研讨\"与\"华夏中医骨伤学报\"与\"2023\"与"
                "\"31(06)\"与\"28-33+39\"，或第 8 条非\"阮清禾,戚昱恒\"与\"中西协同干预腰腿痹阻"
                "病症研究概况\"与\"融合医学诊疗通报\"与\"2025\"与\"25(03)\"与\"58-63\"，或第 15 "
                "条非\"Cao LN, Wilson H, Im CH\"与\"Imaging evaluation standard of lumbar "
                "tissue injury degree\"与\"Cross-border Medical Imaging\"与\"2025\"与"
                "\"35(2)\"与\"188-194\"", bad)

    # 17 -3 10 个"（续）"标题未被删改
    need_cont = {}
    for t in CONT_TITLES:
        need_cont[t] = need_cont.get(t, 0) + 1
    add_penalty("前言（续）、一、资料来源与研究设计（续）、四、基线资料分析（续）、四、基线资料"
                "分析（续）、一、证候要素与干预路径（续）、二、疗效变化与阶段性表现（续）、二、"
                "模型性能与校准评价（续）、三、关键变量解释与应用（续）、讨论（续）、讨论（续）"
                "共 10 个标题中任一项的文字被删除或改名",
                any(texts.count(t) < n for t, n in need_cont.items()))

    # 18 -3 封面与声明页、授权书页要件
    all_text = _nospace(" ".join(texts))
    bad = False
    ci = _find_para(paras, "硕士学位论文")
    if ci < 0 or _para_font_problems(doc, paras[ci], HEITI, None, PT_XIAOCHU, True):
        bad = True
    di = _find_para(paras, "星途理工大学研究生学位论文原创性声明")
    if di < 0 or _para_font_problems(doc, paras[di], HEITI, None, PT_XIAOER, True):
        bad = True
    ai2 = _find_para(paras, "学位论文版权使用授权书")
    if ai2 < 0 or _para_font_problems(doc, paras[ai2], HEITI, None, PT_XIAOER, True):
        bad = True
    for kw in ("作者签名", "签字日期", "指导教师签名"):
        if _nospace(kw) not in all_text:
            bad = True
            break
    add_penalty("封面缺\"硕士学位论文\"字样或其字体非黑体、字号框显示非小初（36 磅）、未加粗，"
                "或缺\"星途理工大学研究生学位论文原创性声明\"页或其标题字体非黑体、字号框显示非"
                "小二（18 磅）、未加粗或缺\"作者签名\"\"签字日期\"字样，或缺\"学位论文版权使用"
                "授权书\"页或其标题字体非黑体、字号框显示非小二（18 磅）、未加粗或缺\"作者签名\""
                "\"指导教师签名\"\"签字日期\"字样", bad)

    # 19 -3 新增标题、译注或修改说明类文字
    allowed_heads = set(ALL_TITLES) | {
        "目录", COVER_TITLE_CN, "硕士学位论文",
        "星途理工大学研究生学位论文原创性声明", "学位论文版权使用授权书",
    }
    bad = any(_nospace(kw) in all_text for kw in NOTE_KEYWORDS)
    if not bad:
        for i, p in enumerate(paras):
            t = texts[i]
            if not t or _outline_level(doc, p) is None:
                continue
            if t not in allowed_heads and not t.startswith(COVER_TITLE_EN_HEAD):
                bad = True
                break
    add_penalty("出现原文之外的新增标题文字、新增正文段落、译注或修改说明类文字", bad)

    # 20 -1 正文出现整页空白页、目录空白条目、连续 3 个以上空段落
    #        "整页空白页"改为静态形式判据：正文区内相邻强制分页之间无可见内容
    bad = _has_static_blank_page(pm, loc)
    if not bad:
        for title, num in _toc_entries(doc, paras):
            if not title and num is not None:
                bad = True
                break
    if not bad:
        s_i = _find_para(paras, "前言")
        e_i = _find_para(paras, "参考文献")
        if s_i >= 0 and e_i > s_i:
            run = 0
            for i in range(s_i, e_i):
                if _is_blank(paras[i]):
                    run += 1
                    if run > 3:
                        bad = True
                        break
                else:
                    run = 0
    add_penalty("中英文缩略词表页至末页之间出现整页空白页，或目录中出现只有页码没有文字的空白"
                "条目，或前言至参考文献之间出现连续 3 个以上空段落", bad)

    # 21 -5 目录变纯文本、对齐异常、标题1标题2正文样式被删
    bad = False
    if _find_para(paras, "目录") >= 0 and not _has_toc_field(file_path):
        bad = True
    root = _styles_element(doc)
    if root is not None:
        style_names = {_style_name(st).strip().lower()
                       for st in root.findall(qn("w:style"))}
        for aliases in (("heading 1", "标题 1"), ("heading 2", "标题 2"),
                        ("normal", "正文")):
            if not any(a.lower() in style_names for a in aliases):
                bad = True
                break
    else:
        bad = True
    VALID_JC = {"left", "start", "right", "end", "center", "both", "justify",
                "distribute", "mediumKashida", "lowKashida", "highKashida",
                "thaiDistribute", "numTab"}
    if not bad:
        for p in paras:
            jc = _alignment(doc, p)
            if jc is not None and jc not in VALID_JC:
                bad = True
                break
    add_penalty("目录变为纯文本无法更新，或段落对齐方式显示异常，或样式库中\"标题 1\"\"标题 2\""
                "\"正文\"样式被删除", bad)

    # 22 -1 修订、批注、高亮、彩色文字
    settings = _part_xml(file_path, "word/settings.xml")
    bad = bool(re.search(r"<w:trackChanges\b", settings))
    if not bad and re.search(r"<w:(?:ins|del)\b", xml):
        bad = True
    if not bad and re.search(r"<w:commentReference\b", xml):
        bad = True
    if not bad and any(n.startswith("word/comments") for n in _part_names(file_path)):
        bad = True
    if not bad:
        toc_set = set(_toc_style_paras(doc, paras))
        for i, p in enumerate(paras):
            if i in toc_set or _outline_level(doc, p) is not None:
                continue  # 目录条目的超链接配色不属于正文文字着色
            for r in _text_runs(p):
                if _run_highlight(doc, r):
                    bad = True
                    break
                val, _theme = _resolve_effective_color(doc, r)
                if _is_chromatic(val):
                    bad = True
                    break
            if bad:
                break
    add_penalty("修订功能处于开启状态或存在未接受的修订，或存在批注，或存在文字高亮，或正文"
                "文字显示为红、蓝、绿、黄等彩色", bad)


    if len(bonus) != len(_MAX_DELTAS_ORDER):
        raise RuntimeError("加分项数不一致：bonus=%d vs order=%d"
                           % (len(bonus), len(_MAX_DELTAS_ORDER)))
    if len(penalty) != len(_PENALTY_DELTAS_ORDER):
        raise RuntimeError("减分项数不一致：penalty=%d vs order=%d"
                           % (len(penalty), len(_PENALTY_DELTAS_ORDER)))

    dim2_items = []
    bonus_score = 0
    penalty_score = 0
    for raw, max_delta in zip(bonus, _MAX_DELTAS_ORDER):
        head, _, rule = raw.partition("：")
        try:
            delta = int(head.replace("+", "").strip())
        except ValueError:
            delta = 0
        hit = delta > 0
        dim2_items.append({"rule": rule, "max_delta": max_delta,
                           "delta": max_delta if hit else 0,
                           "hit": hit})
        bonus_score += (max_delta if hit else 0)
    for raw, max_delta in zip(penalty, _PENALTY_DELTAS_ORDER):
        head, _, rule = raw.partition("：")
        try:
            delta = int(head.replace("-", "").strip())
        except ValueError:
            delta = 0
        hit = delta > 0
        dim2_items.append({"rule": rule, "max_delta": -max_delta,
                           "delta": -max_delta if hit else 0,
                           "hit": hit})
        penalty_score += (max_delta if hit else 0)

    result["dim2_items"] = dim2_items
    result["total_score"] = max(0, bonus_score - penalty_score)
    result["max_score"] = MAX_SCORE_TOTAL
    return result

if __name__ == "__main__":
    _dir = sys.argv[1] if len(sys.argv) >= 2 else os.path.dirname(os.path.abspath(__file__))
    print(json.dumps(evaluate(_dir), ensure_ascii=False, indent=2))

