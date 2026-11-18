#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# pyright: reportUnknownMemberType=false, reportUnknownVariableType=false, reportUnknownArgumentType=false, reportUnknownParameterType=false, reportMissingParameterType=false, reportPrivateUsage=false, reportUnknownLambdaType=false, reportGeneralTypeIssues=false

import os
import re
import sys
import json
import zipfile
import unicodedata
from docx import Document
from docx.text.paragraph import Paragraph
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn

# 说明：按接口统一约定，本模块不修改全局 sys.stdout，也不 print 主结果。
# 主结果统一由 evaluate(dir_path) -> dict 返回；入参为"脚本所在目录的路径"，
# 由脚本自己在该目录内定位并打开被评估的 .docx。本文件仅在 __main__ 下作为
# 本地调试入口，才把 dict 序列化为 JSON 打印。
#
# 任务背景：把《2026年3月17日高中数学作业》(PDF) 的试题内容灌入所给 Word 试卷
# 模板，版式一律以模板为准（不是以 PDF 版式为准），且正文与公式须可编辑。
# 模板既有设定（本脚本的比对基准，硬编码于 _TPL_* 常量）：
#   A4 纵向单栏；标题行 黑体 小二号(18磅) 加粗 居中 1.1 倍行距；
#   卷首信息行 居中 加粗 五号(10.5磅) 1.1 倍行距；
#   题型小标题 五号(10.5磅) 加粗 左对齐 1.35 倍行距；
#   题位段落 中文宋体 + 西文 Times New Roman 五号(10.5磅) 不加粗 约 1.34 倍行距；
#   页脚 居中"第 X 页 共 Y 页" 小五号(9磅) 且页码为 PAGE/NUMPAGES 域；
#   全文无表格、无页眉；默认字体 西文 Times New Roman / 中文宋体。

SCRIPT_ID = "006"
# 维度二所有评分项 max_delta 之和；用于维度一未通过 / 脚本异常时兜底
MAX_SCORE_TOTAL = 60

# 维度二所有评分项按"结果追加顺序"的 max_delta 列表；末端把 results 里的
# "+N：..." 字符串反解为结构化 dim2_items 时用。顺序须与 _evaluate_impl 中
# results.append(...) 的顺序严格一致；共 30 项，累加 = MAX_SCORE_TOTAL。
_MAX_DELTAS_ORDER = [
    # ---- 题目内容搬运完整性（13 项 / 19 分）----
    3,  # 01 单选题第1-10题题干齐全
    1,  # 02 多选题第11-14题题干齐全（仅4道，与填空题3道的分值对齐）
    1,  # 03 填空题第15-17题题干齐全
    1,  # 04 解答题第18-21题题干齐全（仅4道，与多选题4道的分值对齐）
    3,  # 05 选择题 A/B/C/D 四选项齐全
    1,  # 06 解答题各小问齐全（仅4道题的小问，与解答题题干分值对齐）
    3,  # 07 题干数值与条件无抄错
    1,  # 08 题型分组顺序
    1,  # 09 卷首考生信息行已写入
    1,  # 10 填空题留有作答空位
    1,  # 11 第10题抛物线示意图
    1,  # 12 第19题等腰三角形示意图
    1,  # 13 两幅示意图位置正确
    # ---- 数学表达的可编辑性（6 项 / 16 分）----
    5,  # 14 分式为 Word 公式真分数
    3,  # 15 根式为 Word 公式根号
    3,  # 16 幂次与上下标为公式/真上下标
    3,  # 17 数学符号显示正确
    1,  # 18 角度使用度符号 °
    1,  # 19 集合与区间括号配对完整
    # ---- 与模板格式一致（11 项 / 25 分）----
    1,  # 20 标题行 黑体小二加粗居中1.1倍行距
    1,  # 21 卷首信息行 居中加粗五号1.1倍行距
    3,  # 22 题型小标题 五号加粗左对齐1.35倍行距
    3,  # 23 题干与选项 宋体+TNR 五号 不加粗
    3,  # 24 题干与选项行距约1.34倍
    5,  # 25 选项用制表位分列、全文无表格
    3,  # 26 页脚版式"第 X 页 共 Y 页"小五号居中
    1,  # 27 A4 纵向单栏
    3,  # 28 同一题内无割裂该题的手工分页符
    1,  # 29 解答题小问单独成段并缩进
    1,  # 30 正文未混入来源PDF页面标记
]
assert sum(_MAX_DELTAS_ORDER) == MAX_SCORE_TOTAL

# 各评分项的细则原文，顺序与 _MAX_DELTAS_ORDER / results.append(...) 严格一致。
# dim2_items 的 "rule" 只取本表中的细则文本; results 里的 "+N：..." 字符串仅用于
# 反解 delta, 其携带的运行期实测数据(如"（表格 14 个，制表位/空格分列 0/14 题）")
# 不写入 rule。
_RULES_ORDER = [
    # ---- 题目内容搬运完整性 ----
    "单选题第1-10题题干全部搬入，无缺题",
    "多选题第11-14题题干全部搬入，无缺题",
    "填空题第15-17题题干全部搬入，无缺题",
    "解答题第18-21题题干全部搬入，无缺题",
    "每道选择题的A/B/C/D四个选项齐全，未漏项、未串项",
    "解答题各小问齐全——18题(1)(2)、19题(1)(2)(3)、20题(1)(2)、21题(1)(2)",
    "题干中的数值与条件与来源一致，无抄错",
    "题型分组顺序为单选题→多选题→填空题→解答题，与模板题位顺序一致",
    "卷首考生信息行（学校/姓名/班级/考号）已写入，并留有可填写的空位",
    "填空题第15-17题句末留有作答横线或空位",
    "第10题的抛物线与x轴交点示意图随该题呈现，图中A、B、P、O与x、y轴标注齐全",
    "第19题的等腰三角形示意图随该题呈现，图中A、B、C、D、E、M、N标注齐全",
    "两幅示意图均置于对应题干与选项/小问之间，未错位到别的题下",
    # ---- 数学表达的可编辑性 ----
    "分式以Word公式对象的真分数形式排版，可双击进入编辑，未降级为斜杠平排写法",
    "根式以Word公式对象的根号形式排版，根号线完整覆盖被开方式",
    "幂次与上下标以公式对象或Word真上标/下标格式呈现，未用Unicode特殊字符拼凑",
    "数学符号（∩、∪、∈、≤、≠、∥、△、∞、π、α）显示正确，无乱码、方框或缺字",
    "角度单位使用度符号°（30°、45°、90°、120°），未用形近字符替代",
    "集合与区间括号（{ }、[ ]、( )、∪）配对完整，未出现断裂或缺半边",
    # ---- 与模板格式一致 ----
    "卷首标题行沿用模板标题格式——黑体、小二号（18磅）、加粗、居中、1.1倍行距",
    "卷首考生信息行沿用模板同位行格式——居中、加粗、五号（10.5磅）、1.1倍行距",
    "四个题型小标题沿用模板格式——五号（10.5磅）、加粗、左对齐、1.35倍行距，未改为居中或放大字号",
    "题干与选项段落沿用模板题位格式——中文宋体、西文Times New Roman、五号（10.5磅）、不加粗",
    "题干与选项段落行距沿用模板设定（约1.34倍），未压成单倍行距或另设其他固定值",
    "选项以段落内制表位/空格分列排版，全文不含表格对象，与模板的纯段落结构一致",
    "页脚沿用模板页码版式——居中的\"第 X 页 共 Y 页\"、小五号（9磅），未改写措辞或字号",
    "纸张为A4（210mm×297mm）纵向、单栏排版，与模板一致",
    "同一道题内未插入割裂该题的手工分页符（Ctrl+Enter 分页符、\"段前分页\"属性），"
    "整题另起一页与自然换页不计为割裂",
    "解答题小问(1)(2)(3)单独成段并统一缩进，层级清晰",
    "正文（题干、选项、小问）中未混入来源PDF的页面标记"
    "（形如\"试卷第N页，共3页\"\"答案第N页，共2页\"），"
    "模板页脚自带的\"第 X 页 共 Y 页\"不计入本项",
]
assert len(_RULES_ORDER) == len(_MAX_DELTAS_ORDER)

# ---------------- 模板基准常量 ----------------
_TPL_TITLE_SIZE = (17.5, 18.5)      # 小二号 18 磅
_TPL_TITLE_LS = (1.05, 1.15)        # 1.1 倍行距
_TPL_TITLE_EASTASIA = ('黑体', 'SimHei')
_TPL_BODY_SIZE = (10.0, 11.0)       # 五号 10.5 磅
_TPL_INFO_LS = (1.05, 1.15)         # 1.1 倍行距
_TPL_HEAD_LS = (1.30, 1.40)         # 1.35 倍行距
_TPL_QUES_LS = (1.28, 1.40)         # 约 1.34 倍行距
_TPL_FOOTER_SIZE = (8.5, 9.6)       # 小五号 9 磅（含转换取整容差）
_TPL_SONGTI = ('宋体', 'SimSun', 'NSimSun')
_TPL_ASCII = 'Times New Roman'
_TPL_PGSZ = (11906, 16838)          # A4 纵向（twips）

_NS_M = 'http://schemas.openxmlformats.org/officeDocument/2006/math'
_NS_MC = 'http://schemas.openxmlformats.org/markup-compatibility/2006'


def _m(tag):
    """OMML（数学）命名空间下的标签全名。"""
    return '{%s}%s' % (_NS_M, tag)


_MC_FALLBACK = '{%s}Fallback' % _NS_MC
_W_T = qn('w:t')
_M_T = _m('t')
_W_TAB = qn('w:tab')
_W_BR = qn('w:br')
_W_INSTR = qn('w:instrText')
_W_DEL = qn('w:delText')

# ---------------- 文本抽取与归一化 ----------------

# Unicode 上/下标字符（判"用特殊字符拼凑上下标"用，同时用于归一化比对）
_SUP_MAP = {
    '⁰': '0', '¹': '1', '²': '2', '³': '3', '⁴': '4',
    '⁵': '5', '⁶': '6', '⁷': '7', '⁸': '8', '⁹': '9',
    '⁺': '+', '⁻': '-', 'ⁿ': 'n', 'ˣ': 'x', 'ᵃ': 'a',
    'ᵇ': 'b', 'ᶜ': 'c', 'ᵏ': 'k', 'ᵐ': 'm',
}
_SUB_MAP = {
    '₀': '0', '₁': '1', '₂': '2', '₃': '3', '₄': '4',
    '₅': '5', '₆': '6', '₇': '7', '₈': '8', '₉': '9',
    '₊': '+', '₋': '-', 'ₙ': 'n', 'ₐ': 'a', 'ₑ': 'e',
    'ₖ': 'k', 'ₘ': 'm', 'ᵢ': 'i', 'ⱼ': 'j',
}
_UNI_SCRIPT_CHARS = set(_SUP_MAP) | set(_SUB_MAP)

# 各类"减号/连字符"统一成 ASCII '-'
_MINUS_CHARS = '−–—‐‑‒－˗'
# 各类括号统一成 ASCII
_BRACKET_MAP = {
    '（': '(', '）': ')', '［': '[', '］': ']',
    '｛': '{', '｝': '}', '【': '[', '】': ']',
    '｜': '|',
}


def _el_text(el):
    """按文档顺序抽取元素内的可见文字，含 OMML 公式中的 m:t。

    - 跳过 mc:Fallback 子树（AlternateContent 的回退分支会造成文字重复）；
    - 跳过域代码 w:instrText 与删除痕迹 w:delText；
    - w:tab -> '\\t'，w:br -> '\\n'。
    """
    parts = []

    def walk(node):
        for ch in node:
            tag = ch.tag
            if not isinstance(tag, str):
                continue
            if tag == _MC_FALLBACK or tag == _W_INSTR or tag == _W_DEL:
                continue
            if tag == _W_T or tag == _M_T:
                parts.append(ch.text or '')
                continue
            if tag == _W_TAB:
                parts.append('\t')
            elif tag == _W_BR:
                parts.append('\n')
            walk(ch)

    walk(el)
    return ''.join(parts)


def _norm(s):
    """数学文本归一化：Unicode 上下标还原为普通字符、统一减号/括号/全半角，
    并去掉全部空白，便于与来源条件做"是否抄错"的稳定比对。"""
    if not s:
        return ''
    out = []
    for ch in s:
        if ch in _SUP_MAP:
            out.append(_SUP_MAP[ch])
        elif ch in _SUB_MAP:
            out.append(_SUB_MAP[ch])
        elif ch in _MINUS_CHARS:
            out.append('-')
        elif ch in _BRACKET_MAP:
            out.append(_BRACKET_MAP[ch])
        elif ch.isspace():
            continue
        else:
            out.append(ch)
    txt = ''.join(out)
    # 全角字母数字 -> 半角（NFKC 只作用于这些字符，避免影响 CJK）
    txt = ''.join(unicodedata.normalize('NFKC', c) if '！' <= c <= '～' else c
                  for c in txt)
    return txt


def _in_range(val, lo_hi):
    return val is not None and lo_hi[0] <= val <= lo_hi[1]


def _has_cjk(s):
    return any('一' <= c <= '鿿' for c in (s or ''))


def _has_ascii_word(s):
    return any(('0' <= c <= '9') or ('a' <= c <= 'z') or ('A' <= c <= 'Z')
               for c in (s or ''))


# ---------------- 字体/字号继承链解析 ----------------

def _style_chain(doc, style_id, limit=8):
    """从 styles.xml 沿 w:basedOn 向上收集样式元素。"""
    chain = []
    styles_el = doc.styles.element
    cur = style_id
    seen = set()
    while cur and cur not in seen and len(chain) < limit:
        seen.add(cur)
        found = None
        for st in styles_el.findall(qn('w:style')):
            if st.get(qn('w:styleId')) == cur:
                found = st
                break
        if found is None:
            break
        chain.append(found)
        based = found.find(qn('w:basedOn'))
        cur = based.get(qn('w:val')) if based is not None else None
    return chain


def _para_style_id(para):
    p_style = para._p.find(qn('w:pPr'))
    if p_style is None:
        return None
    ps = p_style.find(qn('w:pStyle'))
    return ps.get(qn('w:val')) if ps is not None else None


def _doc_defaults_rpr(doc):
    dd = doc.styles.element.find(qn('w:docDefaults'))
    if dd is None:
        return None
    rpr_dd = dd.find(qn('w:rPrDefault'))
    if rpr_dd is None:
        return None
    return rpr_dd.find(qn('w:rPr'))


def _eastasia_of(rpr):
    if rpr is None:
        return None
    fonts = rpr.find(qn('w:rFonts'))
    if fonts is None:
        return None
    return fonts.get(qn('w:eastAsia'))


def _ascii_of(rpr):
    if rpr is None:
        return None
    fonts = rpr.find(qn('w:rFonts'))
    if fonts is None:
        return None
    return fonts.get(qn('w:ascii')) or fonts.get(qn('w:hAnsi'))


def _sz_of(rpr):
    """w:sz 为半磅值，返回磅值。"""
    if rpr is None:
        return None
    sz = rpr.find(qn('w:sz'))
    if sz is None:
        return None
    try:
        return int(sz.get(qn('w:val'))) / 2.0
    except (TypeError, ValueError):
        return None


def _resolve_effective(doc, run, para, getter):
    """run rPr -> 段落样式(含 basedOn 链) -> docDefaults 依次取值。"""
    val = getter(run._r.find(qn('w:rPr')))
    if val:
        return val
    for st in _style_chain(doc, _para_style_id(para)):
        val = getter(st.find(qn('w:rPr')))
        if val:
            return val
    return getter(_doc_defaults_rpr(doc))


def _resolve_effective_eastasia(doc, run, para):
    return _resolve_effective(doc, run, para, _eastasia_of)


def _resolve_effective_ascii(doc, run, para):
    return _resolve_effective(doc, run, para, _ascii_of)


def _resolve_effective_size(doc, run, para):
    val = _resolve_effective(doc, run, para, _sz_of)
    if val:
        return val
    # 兜底：Normal 样式字号（模板 Normal 为五号 10.5 磅）
    try:
        sz = doc.styles['Normal'].font.size
        if sz is not None:
            return sz.pt
    except Exception:
        pass
    return 10.5


def _run_bold(doc, run, para):
    """加粗的有效值：run -> 段落样式链；None 视为不加粗。"""
    rpr = run._r.find(qn('w:rPr'))
    if rpr is not None:
        b = rpr.find(qn('w:b'))
        if b is not None:
            v = b.get(qn('w:val'))
            return v not in ('0', 'false', 'off')
    for st in _style_chain(doc, _para_style_id(para)):
        rpr = st.find(qn('w:rPr'))
        if rpr is None:
            continue
        b = rpr.find(qn('w:b'))
        if b is not None:
            v = b.get(qn('w:val'))
            return v not in ('0', 'false', 'off')
    return False


# ---------------- 文档结构切分 ----------------

_QNUM_RE = re.compile(r'^(\d{1,2})\s*[．.、]')
_HEAD_KEYS = ('单选题', '多选题', '填空题', '解答题')
_OPT_RE = re.compile(r'([ABCD])\s*[．.、]')
_SUBQ_RE = re.compile(r'\(\s*([123])\s*\)')
_SRC_MARK_RE = re.compile(r'(试卷|答案)第\d+页,?共\d+页')
_FOOTER_OK_RE = re.compile(r'^(第\d*页共\d*页)+$')


def _iter_blocks(doc):
    """把正文顶层子元素展平成有序 block 列表（段落与表格），并预取文字。"""
    blocks = []
    for child in doc.element.body.iterchildren():
        tag = child.tag
        if not isinstance(tag, str):
            continue
        if tag == qn('w:p'):
            blocks.append({'kind': 'p', 'el': child, 'text': _el_text(child)})
        elif tag == qn('w:tbl'):
            blocks.append({'kind': 'tbl', 'el': child, 'text': _el_text(child)})
    return blocks


def _find_headings(blocks):
    """定位四个题型小标题 block 下标：{'单选题': idx, ...}。"""
    heads = {}
    for i, b in enumerate(blocks):
        if b['kind'] != 'p':
            continue
        t = _norm(b['text'])
        if not t or len(t) > 60:
            continue
        if _QNUM_RE.match(t):
            continue
        for key in _HEAD_KEYS:
            if key in t and key not in heads:
                heads[key] = i
    return heads


def _find_questions(blocks):
    """按 1..21 严格递增匹配题号，返回 {题号: 起始 block 下标}。"""
    starts = {}
    expected = 1
    for i, b in enumerate(blocks):
        if expected > 21:
            break
        t = _norm(b['text'])
        m = _QNUM_RE.match(t)
        if m and int(m.group(1)) == expected:
            starts[expected] = i
            expected += 1
    return starts


def _question_ranges(blocks, q_starts, heads):
    """每题的 block 区间 [start, end)：下一题起点或下一个题型小标题处截断。"""
    head_idx = sorted(heads.values())
    ranges = {}
    for n, start in q_starts.items():
        end = q_starts.get(n + 1, len(blocks))
        for hi in head_idx:
            if start < hi < end:
                end = hi
                break
        ranges[n] = (start, end)
    return ranges


def _range_text(blocks, rng):
    s, e = rng
    return '\n'.join(blocks[i]['text'] for i in range(s, e))


def _count_tags(el, tag):
    return sum(1 for _ in el.iter(tag))


def _has_figure(blocks, rng):
    """区间内是否存在图形对象（DrawingML 或 VML）。"""
    s, e = rng
    for i in range(s, e):
        el = blocks[i]['el']
        if _count_tags(el, qn('w:drawing')) or _count_tags(el, qn('w:pict')):
            return True
    return False


def _open_word_document(file_path):
    """打开 .docx；异常直接向上抛，由 evaluate 统一兜底。"""
    return Document(file_path)


def _zip_read(file_path, name_filter):
    """读取包内匹配的 XML 部件文本，返回 {部件名: 文本}。"""
    out = {}
    with zipfile.ZipFile(file_path) as zf:
        for name in zf.namelist():
            if name_filter(name):
                out[name] = zf.read(name).decode('utf-8', errors='ignore')
    return out


def evaluate(dir_path):
    """入口：在 dir_path 内定位待评估的 .docx 并返回评分字典。

    - 定位不到被评估文件时：error 只回纯文本原因（不带异常类名），
      dim1_pass 与 dim1_reason 一律留空（空字符串，表示维度一根本没进入判定），
      dim2_items=[]、total_score=max_score=0。
    - 维度一未通过时：dim1_pass=False，且 dim1_reason 带出错误摘要。
    """
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
            raise FileNotFoundError(f"目录不存在或不是目录：{dir_path}")
        cands = sorted(f for f in os.listdir(dir_path)
                       if f.lower().endswith(('.docx', '.doc'))
                       and not f.startswith('~$')
                       and os.path.isfile(os.path.join(dir_path, f)))
        if not cands:
            raise FileNotFoundError(f"目录中未找到 .docx 文件：{dir_path}")
        file_path = os.path.join(dir_path, cands[0])
        result["file_name"] = os.path.basename(file_path)
        return _evaluate_impl(file_path, result)
    except FileNotFoundError as exc:
        # 找不到被评估文件：error 只回传纯文本原因，不带异常类名；维度一根本没进入
        # 判定，dim1_pass 与 dim1_reason 一律留空（空字符串，而非 False），满分归零
        result["status"] = "error"
        result["error"] = str(exc)
        result["file_name"] = ""
        result["dim1_pass"] = ""
        result["dim1_reason"] = ""
        result["dim2_items"] = []
        result["total_score"] = 0
        result["max_score"] = 0
        return result
    except Exception as exc:
        # 文件存在但打不开/解析失败：维度一未通过，dim1_reason 带上错误摘要便于定位
        result["status"] = "error"
        result["error"] = f"脚本运行异常：{type(exc).__name__}: {exc}"
        result["dim1_pass"] = False
        result["dim1_reason"] = f"文件无法解析或打开失败：{type(exc).__name__}: {exc}"
        result["dim2_items"] = []
        result["total_score"] = 0
        return result


def _evaluate_impl(file_path, result):
    doc = _open_word_document(file_path)
    body = doc.element.body
    blocks = _iter_blocks(doc)

    full_text = '\n'.join(b['text'] for b in blocks)
    full_norm = _norm(full_text)

    heads = _find_headings(blocks)
    q_starts = _find_questions(blocks)
    q_ranges = _question_ranges(blocks, q_starts, heads)
    q_norm = {n: _norm(_range_text(blocks, r)) for n, r in q_ranges.items()}
    q_raw = {n: _range_text(blocks, r) for n, r in q_ranges.items()}

    n_tbl = _count_tags(body, qn('w:tbl'))

    # ---------- 维度1：可正常打开 ----------
    # 唯一门槛：文件能被解析打开。打不开会在 evaluate 的异常兜底里置 dim1_pass=False；
    # 能走到这里说明 Document(file_path) 已成功，即视为通过。
    result["dim1_pass"] = True
    result["dim1_reason"] = "文件可正常打开"

    results = []

    # ============ 一、题目内容搬运完整性 ============
    # 各题的特征关键词（归一化后比对），用于判断题干是否真正搬入而非仅有题号
    q_keys = {
        1: ('已知集合',), 2: ('图象经过点',), 3: ('sin',), 4: ('等差数列',),
        5: ('圆心到直线',), 6: ('红球',), 7: ('恒成立',), 8: ('正方体',),
        9: ('最大值',), 10: ('抛物线',),
        11: ('单调递增',), 12: ('等比数列',), 13: ('两个不同',), 14: ('骰子',),
        15: ('定义域',), 16: ('tan',), 17: ('渐近线',),
        18: ('已知数列',), 19: ('等腰三角形',), 20: ('椭圆',), 21: ('极小值',),
    }

    def _group_ok(nums):
        missing = []
        for n in nums:
            txt = q_norm.get(n)
            if not txt or not any(k in txt for k in q_keys[n]):
                missing.append(n)
        return missing

    # 01 单选题 1-10
    miss = _group_ok(range(1, 11))
    if not miss:
        results.append("+3：单选题第1-10题题干全部搬入，无缺题（10/10）")
    else:
        results.append(f"+0：单选题第1-10题题干全部搬入，无缺题（缺 {miss}）")

    # 02 多选题 11-14
    miss = _group_ok(range(11, 15))
    if not miss:
        results.append("+1：多选题第11-14题题干全部搬入，无缺题（4/4）")
    else:
        results.append(f"+0：多选题第11-14题题干全部搬入，无缺题（缺 {miss}）")

    # 03 填空题 15-17
    miss = _group_ok(range(15, 18))
    if not miss:
        results.append("+1：填空题第15-17题题干全部搬入，无缺题（3/3）")
    else:
        results.append(f"+0：填空题第15-17题题干全部搬入，无缺题（缺 {miss}）")

    # 04 解答题 18-21
    miss = _group_ok(range(18, 22))
    if not miss:
        results.append("+1：解答题第18-21题题干全部搬入，无缺题（4/4）")
    else:
        results.append(f"+0：解答题第18-21题题干全部搬入，无缺题（缺 {miss}）")

    # 05 选择题 A/B/C/D 四选项齐全
    bad_opt = []
    for n in range(1, 15):
        labels = set(_OPT_RE.findall(q_norm.get(n, '')))
        if labels != {'A', 'B', 'C', 'D'}:
            bad_opt.append(n)
    if not bad_opt:
        results.append("+3：每道选择题的A/B/C/D四个选项齐全，未漏项、未串项（14题）")
    else:
        results.append(f"+0：每道选择题的A/B/C/D四个选项齐全，未漏项、未串项（异常题号 {bad_opt}）")

    # 06 解答题各小问齐全
    need_sub = {18: {'1', '2'}, 19: {'1', '2', '3'}, 20: {'1', '2'}, 21: {'1', '2'}}
    bad_sub = []
    for n, need in need_sub.items():
        got = set(_SUBQ_RE.findall(q_norm.get(n, '')))
        if not need.issubset(got):
            bad_sub.append(n)
    if not bad_sub:
        results.append("+1：解答题各小问齐全——18题(1)(2)、19题(1)(2)(3)、20题(1)(2)、21题(1)(2)")
    else:
        results.append(f"+0：解答题各小问齐全（小问缺失题号 {bad_sub}）")

    # 07 题干数值与条件与来源一致
    facts = ['a1=3', 'a8=24', 'AB=AC=5', 'BC=6', 'DE=1',
             'x2+y2-4x+2y-4=0', '3个红球', '2个蓝球', '1个白球']
    lost = [f for f in facts if f not in full_norm]
    if not lost:
        results.append(f"+3：题干中的数值与条件与来源一致，无抄错（{len(facts)}项关键条件全中）")
    else:
        results.append(f"+0：题干中的数值与条件与来源一致，无抄错（未命中 {lost}）")

    # 08 题型分组顺序
    if len(heads) == 4:
        order = [heads[k] for k in _HEAD_KEYS]
        seq_ok = all(order[i] < order[i + 1] for i in range(3))
    else:
        seq_ok = False
    if seq_ok:
        results.append("+1：题型分组顺序为单选题→多选题→填空题→解答题，与模板题位顺序一致")
    else:
        results.append(f"+0：题型分组顺序为单选题→多选题→填空题→解答题（识别到 {len(heads)}/4 个题型小标题）")

    # 09 卷首考生信息行
    info_idx = None
    for i, b in enumerate(blocks[:10]):
        t = _norm(b['text'])
        if all(k in t for k in ('学校', '姓名', '班级', '考号')):
            info_idx = i
            break
    if info_idx is not None:
        results.append("+1：卷首考生信息行（学校/姓名/班级/考号）已写入，并留有可填写的空位")
    else:
        results.append("+0：卷首考生信息行（学校/姓名/班级/考号）已写入，并留有可填写的空位")

    # 10 填空题第15-17题句末留有作答横线或空位
    blank_bad = []
    for n in (15, 16, 17):
        raw = q_raw.get(n, '')
        if not (re.search(r'[_＿]{3,}', raw) or '\t' in raw
                or re.search(r'[ 　]{4,}', raw)):
            blank_bad.append(n)
    if not blank_bad:
        results.append("+1：填空题第15-17题句末留有作答横线或空位")
    else:
        results.append(f"+0：填空题第15-17题句末留有作答横线或空位（缺空位题号 {blank_bad}）")

    # 11 第10题抛物线示意图
    fig10 = 10 in q_ranges and _has_figure(blocks, q_ranges[10])
    if fig10:
        results.append("+1：第10题的抛物线与x轴交点示意图随该题呈现，图中A、B、P、O与x、y轴标注齐全")
    else:
        results.append("+0：第10题的抛物线与x轴交点示意图随该题呈现，图中A、B、P、O与x、y轴标注齐全")

    # 12 第19题等腰三角形示意图
    fig19 = 19 in q_ranges and _has_figure(blocks, q_ranges[19])
    if fig19:
        results.append("+1：第19题的等腰三角形示意图随该题呈现，图中A、B、C、D、E、M、N标注齐全")
    else:
        results.append("+0：第19题的等腰三角形示意图随该题呈现，图中A、B、C、D、E、M、N标注齐全")

    # 13 两幅示意图均置于对应题干与选项/小问之间，未错位到别的题下
    stray = [n for n, r in q_ranges.items()
             if n not in (10, 19) and _has_figure(blocks, r)]
    if fig10 and fig19 and not stray:
        results.append("+1：两幅示意图均置于对应题干与选项/小问之间，未错位到别的题下")
    else:
        results.append(f"+0：两幅示意图均置于对应题干与选项/小问之间，未错位到别的题下（错位题号 {stray}）")

    # ============ 二、数学表达的可编辑性 ============
    n_frac = _count_tags(body, _m('f'))
    n_rad = _count_tags(body, _m('rad'))
    n_ssup = _count_tags(body, _m('sSup'))
    n_ssub = _count_tags(body, _m('sSub'))
    n_vert = 0
    for va in body.iter(qn('w:vertAlign')):
        if va.get(qn('w:val')) in ('superscript', 'subscript'):
            n_vert += 1
    n_uni_script = sum(1 for c in full_text if c in _UNI_SCRIPT_CHARS)
    n_slash = len(re.findall(r'\d\s*/\s*\d|[a-zA-Z]\s*/\s*\d', full_norm))

    # 14 分式以 Word 公式对象的真分数形式排版
    if n_frac >= 8 and n_frac >= n_slash:
        results.append(f"+5：分式以Word公式对象的真分数形式排版，可双击进入编辑，未降级为斜杠平排写法"
                       f"（m:f 真分数 {n_frac} 处，斜杠平排 {n_slash} 处）")
    else:
        results.append(f"+0：分式以Word公式对象的真分数形式排版，可双击进入编辑，未降级为斜杠平排写法"
                       f"（m:f 真分数 {n_frac} 处，斜杠平排 {n_slash} 处）")

    # 15 根式以 Word 公式对象的根号形式排版
    if n_rad >= 3:
        results.append(f"+3：根式以Word公式对象的根号形式排版，根号线完整覆盖被开方式（m:rad {n_rad} 处）")
    else:
        results.append(f"+0：根式以Word公式对象的根号形式排版，根号线完整覆盖被开方式（m:rad {n_rad} 处）")

    # 16 幂次与上下标以公式对象或 Word 真上标/下标呈现
    n_script = n_ssup + n_ssub + n_vert
    if n_script >= 12 and n_uni_script == 0:
        results.append(f"+3：幂次与上下标以公式对象或Word真上标/下标格式呈现，未用Unicode特殊字符拼凑"
                       f"（公式上下标 {n_ssup + n_ssub} 处、真上下标 {n_vert} 处、Unicode拼凑 0 处）")
    else:
        results.append(f"+0：幂次与上下标以公式对象或Word真上标/下标格式呈现，未用Unicode特殊字符拼凑"
                       f"（公式上下标 {n_ssup + n_ssub} 处、真上下标 {n_vert} 处、Unicode拼凑 {n_uni_script} 处）")

    # 17 数学符号显示正确
    need_sym = ('∩', '∪', '∈', '≤', '≠', '∥', '△', '∞', 'π', 'α')
    miss_sym = [s for s in need_sym if s not in full_text]
    garbled = full_text.count('�') + full_text.count('□')
    if not miss_sym and garbled == 0:
        results.append("+3：数学符号（∩、∪、∈、≤、≠、∥、△、∞、π、α）显示正确，无乱码、方框或缺字")
    else:
        results.append(f"+0：数学符号（∩、∪、∈、≤、≠、∥、△、∞、π、α）显示正确，无乱码、方框或缺字"
                       f"（缺失 {miss_sym}，乱码/方框 {garbled} 处）")

    # 18 角度单位使用度符号 °
    n_deg = full_text.count('°')
    n_fake_deg = full_text.count('◦') + full_text.count('∘') + full_text.count('˚')
    if n_deg >= 4 and n_fake_deg == 0:
        results.append(f"+1：角度单位使用度符号°（30°、45°、90°、120°），未用形近字符替代（° {n_deg} 处）")
    else:
        results.append(f"+0：角度单位使用度符号°（30°、45°、90°、120°），未用形近字符替代"
                       f"（° {n_deg} 处，形近字符 {n_fake_deg} 处）")

    # 19 集合与区间括号配对完整
    n_open = sum(full_norm.count(c) for c in '([{')
    n_close = sum(full_norm.count(c) for c in ')]}')
    if abs(n_open - n_close) <= 1:
        results.append(f"+1：集合与区间括号（{{ }}、[ ]、( )、∪）配对完整，未出现断裂或缺半边"
                       f"（左 {n_open} / 右 {n_close}）")
    else:
        results.append(f"+0：集合与区间括号（{{ }}、[ ]、( )、∪）配对完整，未出现断裂或缺半边"
                       f"（左 {n_open} / 右 {n_close}）")

    # ============ 三、与模板格式一致 ============
    def _ls_mult(para):
        """段落行距的"倍数"值；固定值/最小值返回 None（视为未沿用模板倍数行距）。"""
        v = para.paragraph_format.line_spacing
        if v is None:
            return None
        if hasattr(v, 'pt'):
            return None
        return float(v)

    def _vis_runs(para):
        return [r for r in para.runs if (r.text or '').strip()]

    def _is_left(para):
        return para.alignment in (None, WD_ALIGN_PARAGRAPH.LEFT)

    def _para_of(el):
        return Paragraph(el, doc)

    # 20 卷首标题行——黑体、小二号(18磅)、加粗、居中、1.1 倍行距
    title_para = None
    for b in blocks:
        if b['kind'] == 'p' and _norm(b['text']):
            title_para = _para_of(b['el'])
            break
    t_align = t_ls = None
    t_bad = ['未找到标题行']
    if title_para is not None:
        t_bad = []
        t_align = title_para.alignment
        t_ls = _ls_mult(title_para)
        if t_align != WD_ALIGN_PARAGRAPH.CENTER:
            t_bad.append('未居中')
        if not _in_range(t_ls, _TPL_TITLE_LS):
            t_bad.append(f'行距 {t_ls}')
        for r in _vis_runs(title_para):
            if not _in_range(_resolve_effective_size(doc, r, title_para), _TPL_TITLE_SIZE):
                t_bad.append('字号非小二号(18磅)')
                break
        for r in _vis_runs(title_para):
            if not _run_bold(doc, r, title_para):
                t_bad.append('未加粗')
                break
        for r in _vis_runs(title_para):
            if _has_cjk(r.text) and _resolve_effective_eastasia(doc, r, title_para) \
                    not in _TPL_TITLE_EASTASIA:
                t_bad.append('中文字体非黑体')
                break
    if not t_bad:
        results.append("+1：卷首标题行沿用模板标题格式——黑体、小二号（18磅）、加粗、居中、1.1倍行距")
    else:
        results.append(f"+0：卷首标题行沿用模板标题格式——黑体、小二号（18磅）、加粗、居中、1.1倍行距"
                       f"（偏差：{'、'.join(sorted(set(t_bad)))}）")

    # 21 卷首考生信息行——居中、加粗、五号(10.5磅)、1.1 倍行距
    i_bad = ['未找到考生信息行']
    if info_idx is not None and blocks[info_idx]['kind'] == 'p':
        ip = _para_of(blocks[info_idx]['el'])
        i_bad = []
        if ip.alignment != WD_ALIGN_PARAGRAPH.CENTER:
            i_bad.append('未居中')
        if not _in_range(_ls_mult(ip), _TPL_INFO_LS):
            i_bad.append(f'行距 {_ls_mult(ip)}')
        for r in _vis_runs(ip):
            if not _run_bold(doc, r, ip):
                i_bad.append('未加粗')
                break
        for r in _vis_runs(ip):
            if not _in_range(_resolve_effective_size(doc, r, ip), _TPL_BODY_SIZE):
                i_bad.append('字号非五号(10.5磅)')
                break
    if not i_bad:
        results.append("+1：卷首考生信息行沿用模板同位行格式——居中、加粗、五号（10.5磅）、1.1倍行距")
    else:
        results.append(f"+0：卷首考生信息行沿用模板同位行格式——居中、加粗、五号（10.5磅）、1.1倍行距"
                       f"（偏差：{'、'.join(sorted(set(i_bad)))}）")

    # 22 四个题型小标题——五号(10.5磅)、加粗、左对齐、1.35 倍行距
    h_bad = []
    if len(heads) != 4:
        h_bad.append(f'仅识别到{len(heads)}个小标题')
    for key in _HEAD_KEYS:
        if key not in heads:
            continue
        hp = _para_of(blocks[heads[key]]['el'])
        if not _is_left(hp):
            h_bad.append(f'{key}未左对齐')
        if not _in_range(_ls_mult(hp), _TPL_HEAD_LS):
            h_bad.append(f'{key}行距{_ls_mult(hp)}')
        for r in _vis_runs(hp):
            if not _in_range(_resolve_effective_size(doc, r, hp), _TPL_BODY_SIZE):
                h_bad.append(f'{key}字号非五号')
                break
        for r in _vis_runs(hp):
            if not _run_bold(doc, r, hp):
                h_bad.append(f'{key}未加粗')
                break
    if not h_bad:
        results.append("+3：四个题型小标题沿用模板格式——五号（10.5磅）、加粗、左对齐、1.35倍行距，未改为居中或放大字号")
    else:
        results.append(f"+0：四个题型小标题沿用模板格式——五号（10.5磅）、加粗、左对齐、1.35倍行距，未改为居中或放大字号"
                       f"（偏差：{'、'.join(h_bad[:6])}）")

    # 题位段落集合：题干段落 + 区间内（含表格单元格内）的选项/小问段落
    ques_para_els = []
    stem_paras = []
    for n in sorted(q_ranges):
        s, e = q_ranges[n]
        if blocks[s]['kind'] == 'p':
            stem_paras.append(_para_of(blocks[s]['el']))
        for i in range(s, e):
            b = blocks[i]
            if b['kind'] == 'p':
                ques_para_els.append(b['el'])
            else:
                ques_para_els.extend(list(b['el'].iter(qn('w:p'))))

    # 23 题干与选项段落——中文宋体、西文 Times New Roman、五号(10.5磅)、不加粗
    chk = ok = 0
    f_bad = []
    for el in ques_para_els:
        p = _para_of(el)
        for r in _vis_runs(p):
            txt = r.text
            chk += 1
            bad = []
            if not _in_range(_resolve_effective_size(doc, r, p), _TPL_BODY_SIZE):
                bad.append('字号')
            if _run_bold(doc, r, p):
                bad.append('加粗')
            if _has_cjk(txt) and _resolve_effective_eastasia(doc, r, p) not in _TPL_SONGTI:
                bad.append('中文非宋体')
            if _has_ascii_word(txt) and _resolve_effective_ascii(doc, r, p) != _TPL_ASCII:
                bad.append('西文非TNR')
            if bad:
                f_bad.extend(bad)
            else:
                ok += 1
    f_ratio = (ok / chk) if chk else 0.0
    if chk >= 20 and f_ratio >= 0.90:
        results.append(f"+3：题干与选项段落沿用模板题位格式——中文宋体、西文Times New Roman、五号（10.5磅）、不加粗"
                       f"（合规 {ok}/{chk} 段落文字块）")
    else:
        results.append(f"+0：题干与选项段落沿用模板题位格式——中文宋体、西文Times New Roman、五号（10.5磅）、不加粗"
                       f"（合规 {ok}/{chk}，偏差项 {sorted(set(f_bad))}）")

    # 24 题干与选项段落行距沿用模板设定（约 1.34 倍）
    ls_vals = [_ls_mult(p) for p in stem_paras]
    ls_ok = sum(1 for v in ls_vals if _in_range(v, _TPL_QUES_LS))
    if ls_vals and ls_ok / len(ls_vals) >= 0.80:
        results.append(f"+3：题干与选项段落行距沿用模板设定（约1.34倍），未压成单倍行距或另设其他固定值"
                       f"（合规 {ls_ok}/{len(ls_vals)} 个题位段落）")
    else:
        results.append(f"+0：题干与选项段落行距沿用模板设定（约1.34倍），未压成单倍行距或另设其他固定值"
                       f"（合规 {ls_ok}/{len(ls_vals)}，实测值样例 {sorted(set(v for v in ls_vals if v is not None))[:4]}）")

    # 25 选项以段落内制表位/空格分列排版，全文不含表格对象
    tab_opt = 0
    for n in range(1, 15):
        if '\t' in q_raw.get(n, '') or re.search(r'[ 　]{2,}', q_raw.get(n, '')):
            tab_opt += 1
    if n_tbl == 0 and tab_opt >= 12:
        results.append(f"+5：选项以段落内制表位/空格分列排版，全文不含表格对象，与模板的纯段落结构一致"
                       f"（表格 0 个，制表位/空格分列 {tab_opt}/14 题）")
    else:
        results.append(f"+0：选项以段落内制表位/空格分列排版，全文不含表格对象，与模板的纯段落结构一致"
                       f"（表格 {n_tbl} 个，制表位/空格分列 {tab_opt}/14 题）")

    # ---- 页脚 / 节属性 ----
    def _strip_fallback(xml):
        return re.sub(r'<mc:Fallback\b.*?</mc:Fallback>', '', xml, flags=re.S)

    def _xml_texts(xml):
        return ''.join(mt.group(1) for mt in
                       re.finditer(r'<w:t(?:\s[^>]*)?>(.*?)</w:t>', xml, re.S))

    ftr_parts = _zip_read(file_path, lambda nm: re.match(r'word/footer\d*\.xml$', nm))
    ftr_xml = _strip_fallback('\n'.join(ftr_parts.values()))
    ftr_text = _norm(_xml_texts(ftr_xml))
    ftr_sizes = sorted({int(v) / 2.0 for v in re.findall(r'<w:sz w:val="(\d+)"', ftr_xml)})

    # 26 页脚沿用模板页码版式——居中的"第 X 页 共 Y 页"、小五号(9磅)
    ftr_bad = []
    if not ftr_parts:
        ftr_bad.append('无页脚')
    if not _FOOTER_OK_RE.match(ftr_text):
        ftr_bad.append(f'措辞"{ftr_text[:24]}"与模板不一致')
    if 'center' not in ftr_xml:
        ftr_bad.append('未居中')
    if not any(_in_range(s, _TPL_FOOTER_SIZE) for s in ftr_sizes):
        ftr_bad.append(f'字号 {ftr_sizes} 非小五号(9磅)')
    if not ftr_bad:
        results.append(f"+3：页脚沿用模板页码版式——居中的\"第 X 页 共 Y 页\"、小五号（9磅），未改写措辞或字号"
                       f"（页脚文字\"{ftr_text}\"，字号 {ftr_sizes}）")
    else:
        results.append(f"+0：页脚沿用模板页码版式——居中的\"第 X 页 共 Y 页\"、小五号（9磅），未改写措辞或字号"
                       f"（偏差：{'、'.join(ftr_bad)}）")

    # 27 纸张为 A4 纵向、单栏排版
    pg_bad = ['无节属性']
    if doc.sections:
        sec = doc.sections[0]
        pg_bad = []
        try:
            w_tw = sec.page_width.twips
            h_tw = sec.page_height.twips
        except Exception:
            w_tw = h_tw = None
        if not (w_tw and h_tw and abs(w_tw - _TPL_PGSZ[0]) <= 80
                and abs(h_tw - _TPL_PGSZ[1]) <= 80):
            pg_bad.append(f'纸张 {w_tw}x{h_tw} twips 非A4纵向')
        cols = sec._sectPr.find(qn('w:cols'))
        n_col = 1
        if cols is not None:
            try:
                n_col = int(cols.get(qn('w:num')) or 1)
            except ValueError:
                n_col = 1
        if n_col != 1:
            pg_bad.append(f'{n_col}栏')
    if not pg_bad:
        results.append("+1：纸张为A4（210mm×297mm）纵向、单栏排版，与模板一致")
    else:
        results.append(f"+0：纸张为A4（210mm×297mm）纵向、单栏排版，与模板一致"
                       f"（偏差：{'、'.join(pg_bad)}）")

    # 28 同一道题内未插入割裂该题的手工分页符
    #    识别两类手工分页：w:br[@w:type="page"]（Ctrl+Enter 分页符）与 w:pageBreakBefore（段前分页）。
    #    分页落在该题首段 = 整题另起一页，不算割裂；分页之后该题已无文字的也不算。
    #    Word 因页面排满产生的自然换页在 XML 中无标记，不在本项判定范围内。
    def _split_page_break(el):
        """返回 (本块是否含手工分页符, 分页符之后本块内剩余文字)。"""
        state = {'found': False}
        after = []

        def walk(node):
            for ch in node:
                tag = ch.tag
                if not isinstance(tag, str):
                    continue
                if tag == _MC_FALLBACK or tag == _W_INSTR or tag == _W_DEL:
                    continue
                if tag == _W_BR and ch.get(qn('w:type')) == 'page':
                    state['found'] = True
                    continue
                if tag == _W_T or tag == _M_T:
                    if state['found']:
                        after.append(ch.text or '')
                    continue
                walk(ch)

        walk(el)
        return state['found'], ''.join(after)

    split_kind = {}
    for n in sorted(q_ranges):
        s, e = q_ranges[n]
        for i in range(s, e):
            el = blocks[i]['el']
            ppr = el.find(qn('w:pPr')) if blocks[i]['kind'] == 'p' else None
            pbb = False
            if ppr is not None:
                pbb_el = ppr.find(qn('w:pageBreakBefore'))
                if pbb_el is not None:
                    # 转换器常写出显式关闭的 w:val="0"，不能当成分页符
                    pbb = pbb_el.get(qn('w:val')) not in ('0', 'false', 'off')
            has_br, after_txt = _split_page_break(el)
            if not (pbb or has_br):
                continue
            if pbb and i == s:
                continue  # 整题另起一页，未割裂本题
            tail = (_norm(blocks[i]['text']) if pbb else _norm(after_txt))
            tail += ''.join(_norm(blocks[j]['text']) for j in range(i + 1, e))
            if tail:
                split_kind[n] = '段前分页' if pbb else '分页符'
                break
    if not split_kind:
        results.append("+3：同一道题内未插入割裂该题的手工分页符（Ctrl+Enter 分页符、\"段前分页\"属性），"
                       "整题另起一页与自然换页不计为割裂")
    else:
        split_detail = '、'.join(f"第{n}题（{split_kind[n]}）" for n in sorted(split_kind))
        results.append(f"+0：同一道题内未插入割裂该题的手工分页符（Ctrl+Enter 分页符、\"段前分页\"属性），"
                       f"整题另起一页与自然换页不计为割裂（被割裂 {split_detail}）")

    # 29 解答题小问 (1)(2)(3) 单独成段并统一缩进
    sub_bad = []
    for n, need in need_sub.items():
        if n not in q_ranges:
            sub_bad.append(n)
            continue
        s, e = q_ranges[n]
        found = set()
        indent_ok = True
        for i in range(s, e):
            if blocks[i]['kind'] != 'p':
                continue
            t = _norm(blocks[i]['text'])
            mm = _SUBQ_RE.match(t)
            if not mm:
                continue
            found.add(mm.group(1))
            p = _para_of(blocks[i]['el'])
            pf = p.paragraph_format
            raw = blocks[i]['text']
            if not ((pf.left_indent and pf.left_indent > 0)
                    or (pf.first_line_indent and pf.first_line_indent > 0)
                    or raw.startswith('\t') or raw.startswith('　')):
                indent_ok = False
        if not need.issubset(found) or not indent_ok:
            sub_bad.append(n)
    if not sub_bad:
        results.append("+1：解答题小问(1)(2)(3)单独成段并统一缩进，层级清晰")
    else:
        results.append(f"+0：解答题小问(1)(2)(3)单独成段并统一缩进，层级清晰（异常题号 {sorted(sub_bad)}）")

    # 30 正文未混入来源 PDF 的页面标记（只查正文，页脚措辞由第 26 项判定）
    src_marks = [mm.group(0) for mm in _SRC_MARK_RE.finditer(full_norm)]
    if not src_marks:
        results.append("+1：正文（题干、选项、小问）中未混入来源PDF的页面标记"
                       "（形如\"试卷第N页，共3页\"\"答案第N页，共2页\"），"
                       "模板页脚自带的\"第 X 页 共 Y 页\"不计入本项")
    else:
        results.append(f"+0：正文（题干、选项、小问）中未混入来源PDF的页面标记"
                       f"（形如\"试卷第N页，共3页\"\"答案第N页，共2页\"），"
                       f"模板页脚自带的\"第 X 页 共 Y 页\"不计入本项"
                       f"（正文混入 {len(src_marks)} 处：{src_marks[:3]}）")

    # ---------- 汇总：results -> dim2_items ----------
    if len(results) != len(_MAX_DELTAS_ORDER):
        raise RuntimeError(
            f"dim2 项数不一致：results={len(results)} vs order={len(_MAX_DELTAS_ORDER)}")
    dim2_items = []
    total_score = 0
    for raw, max_delta, rule in zip(results, _MAX_DELTAS_ORDER, _RULES_ORDER):
        head = raw.partition("：")[0]
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
    if len(sys.argv) >= 2:
        _dir = sys.argv[1]
    else:
        _dir = os.path.dirname(os.path.abspath(__file__))
    _out = evaluate(_dir)
    print(json.dumps(_out, ensure_ascii=False, indent=2))
