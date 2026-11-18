#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# pyright: reportUnknownMemberType=false, reportUnknownVariableType=false, reportUnknownArgumentType=false, reportUnknownParameterType=false, reportMissingParameterType=false, reportPrivateUsage=false, reportUnknownLambdaType=false, reportGeneralTypeIssues=false
"""
毕业实习报告 / 毕业实习鉴定表 格式自动评估工具（officeval_004）

统一入口：
    evaluate(dir_path: str) -> dict
    - dir_path: 输出文件所在目录（内含待评估的两份 .docx：毕业实习报告、毕业实习鉴定表）
    - 返回字段：id / file_name / status / error / dim1_pass / dim1_reason /
                dim2_items / total_score / max_score
    - dim2_items 每项为 {rule, max_delta, delta, hit}，delta 严格二值

规范约束：
    * 只暴露 evaluate 一个函数供批量运行器调用
    * 不修改 sys.stdout、不打印主结果、不使用 sys.exit
    * 不硬编码文档路径，全部由 dir_path 参数驱动
    * status='error' 表示这份提交没被真正评过（目录不存在、目录内无可评估文档、
      脚本自身异常），此时 dim1_pass 为空串（未判定）、dim1_reason 为空串、
      dim2_items 为空表、total_score 与 max_score 一律为 0；
      文档能定位到但不达标一律 status='ok'，由 dim1_pass / dim2_items 承载
    * dim1_pass 为 False 时，dim1_reason 必定带有维度一未通过的错误摘要；
      dim1_pass 为 True 或空串时，dim1_reason 为空串
    * 本文件仅在 __main__ 下作为本地调试入口，才把 dict 序列化为 JSON 输出
"""

import os
import re
import sys
import json
from docx import Document
from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph
from docx.table import Table

SCRIPT_ID = "004"
# 维度二加分项 max_delta 之和；用于维度一未通过 / 脚本异常时兜底
MAX_SCORE_TOTAL = 59

# 维度二加分项按“结果追加顺序”的 max_delta 列表；末端把 results 里的
# “+N：...” 字符串反解为结构化 dim2_items 时用。顺序须与 _evaluate_impl 中
# results.append(...) 的顺序严格一致；共 23 项，累加 = MAX_SCORE_TOTAL。
#
# 加分项只收录“要求图片明确给出”的格式：图片红字批注的字体/字号/加粗/对齐/行距/
# 缩进/字数，以及图片版式里直接画出的结构（下划线、题序、表格行列与合并）。
# 图片未提、只能靠实测原文档才能得出的属性（如封面落款行13磅、鉴定表标题20磅、
# 落款单元格对齐），一律归减分项，按“原有效果不得被破坏”计分。
_MAX_DELTAS_ORDER = [
    3,  # 01 封面“实习报告”宋体二号加粗居中
    3,  # 02 封面“题 目”行黑体三号加粗
    3,  # 03 封面实习单位/实习日期两段中文宋体四号
    1,  # 04 封面实习单位/实习日期两段居中
    3,  # 05 封面实习单位/实习日期两段行距1.5倍
    3,  # 06 封面实习单位/实习日期两段数字TNR四号
    1,  # 07 封面实习单位/实习日期冒号后下划线空白
    5,  # 08 七个一级标题黑体小三加粗
    3,  # 09 八个二级标题黑体四号加粗
    5,  # 10 正文段落宋体四号
    3,  # 11 正文段落首行缩进2字符
    3,  # 12 正文段落行距1.1倍
    3,  # 13 正文中文字符数不少于2500
    1,  # 14 一至六题序连续无缺项
    5,  # 15 鉴定表八个信息栏标签与值宋体四号
    3,  # 16 鉴定表信息栏数字TNR四号
    1,  # 17 鉴定表信息栏文字居中
    1,  # 18 自我评价正文宋体小四
    3,  # 19 自我评价正文单倍行距 + 首行缩进2字符
    3,  # 20 两个评价栏正文宋体小四/四号、单倍行距、首行缩进2字符
    1,  # 21 两个评价栏正文居中
    1,  # 22 鉴定表8行表格、含合并单元格、栏目齐全
    1,  # 23 备注栏文字宋体四号
]
assert sum(_MAX_DELTAS_ORDER) == MAX_SCORE_TOTAL

# 维度二减分项按“结果追加顺序”的 max_delta 列表（均为负值）；共 27 项。
# 减分项只在命中“指令未涉及、却破坏了文档原有效果”的情形时扣分，不计入 max_score，
# 但会从 total_score 中扣除（扣至 0 为止）。顺序须与 penalties.append(...) 一致。
#
# 两条设计约束：
# ① 加分项与减分项不得覆盖同一个（区域, 属性）——例如标题的字体/字号/加粗归加分项，
#    标题的对齐/缩进/段后间距归减分项，互不重叠。
# ② 属性类减分项一律带前置守卫：判定对象缺失时判为未命中，缺失只由该区域唯一的
#    “存在性”减分项去扣（报告个人情况表见第 14 项、文末落款表见第 20 项、
#    鉴定表备注说明见第 24 项），保证一处破坏只扣一次分。
#
# 数值出处（重要）：减分项里的每个具体数值都取自“输入文件”里的原始模板实测
# （报告=毕业实习报告 (2).doc、鉴定表=毕业实习鉴定表 (2).docx），语义是“模板原有效果
# 不得被破坏”，因此保留模板原值必须不扣分。禁止拿 good/bad 两份示例输出反推数值——
# 那样会让细则围着某一份提交转，既循环论证，也会把“保持原样”误判为破坏。
_PENALTY_DELTAS_ORDER = [
    -3,  # 01 报告/鉴定表页边距被改动
    -3,  # 02 报告或鉴定表纸张尺寸不为 A4
    -3,  # 03 报告页眉/页脚被删除或改写
    -3,  # 04 封面与“个人情况与实习安排”之间的分页被删除
    -1,  # 05 封面区空段落少于8个
    -1,  # 06 封面校徽图片被删除，或学院落款行不为13磅/居中
    -3,  # 07 一级/二级标题对齐不为左对齐
    -1,  # 08 一级/二级标题带首行缩进
    -1,  # 09 “一、实习背景与单位概况”/“（一）实习背景”段后间距不为4磅
    -1,  # 10 报告正文段落两端对齐或带右缩进
    -1,  # 11 “实习初期…”段段前/段后间距不为0磅
    -1,  # 12 “指导教师评语：”行不为加粗
    -1,  # 13 “指导教师评语：”行出现小于12磅文字
    -3,  # 14 报告缺少“个人情况与实习安排”标题或其8行4列表格（存在性）
    -3,  # 15 个人情况表非表头单元格加粗，或“良好”未加粗
    -1,  # 16 个人情况表不为水平居中，或4列列宽不为等宽3.95cm
    -1,  # 17 个人情况表“学生姓名”“林沐言”两格字号不为10.5磅
    -3,  # 18 文末“日期：…”段不为右对齐或带首行缩进
    -1,  # 19 落款表4个标签及“良好”填写值字号不为11磅
    -1,  # 20 落款表不为4行两列结构，或标签/填写值单元格对齐被改（存在性）
    -1,  # 21 鉴定表标题不为黑体20磅加粗居中
    -3,  # 22 鉴定表签名/成绩/落款日期行出现小于11磅文字
    -3,  # 23 鉴定表“自我评价”栏正文对齐不为居中
    -1,  # 24 鉴定表“备注”栏说明被删除，或不为左对齐/单倍行距（存在性）
    -3,  # 25 鉴定表表格总宽或6列列宽被改动
    -1,  # 26 鉴定表信息行/三个评价栏行高被压缩
    -1,  # 27 鉴定表单元格内容不为垂直居中
]

# 字体名判定口径（办公软件里同一字体的不同写法都算命中）
STRICT_SONGTI = ('宋体', 'SimSun', 'NSimSun')
HEITI_KEYS = ('黑体', 'SimHei')
TNR_NAME = 'Times New Roman'

CJK_RE = re.compile(r'[一-鿿]')
DIGIT_RE = re.compile(r'[0-9]')


def _is_songti(name):
    return bool(name) and any(k in name for k in STRICT_SONGTI)


def _is_heiti(name):
    return bool(name) and any(k in name for k in HEITI_KEYS)


def _is_tnr(name):
    return bool(name) and TNR_NAME.lower() in name.lower()


def _has_cjk(text):
    return bool(CJK_RE.search(text or ''))


def _has_digit(text):
    return bool(DIGIT_RE.search(text or ''))


def _cjk_count(text):
    return len(CJK_RE.findall(text or ''))

def _style_chain(doc, run):
    """返回 run 所在段落的样式链（自身 w:pStyle → basedOn 逐级上溯）的 w:style 元素列表。"""
    styles_el = doc.styles.element
    style_id = None
    try:
        para = run._parent
        pPr = para._p.find(qn('w:pPr'))
        if pPr is not None:
            pStyle = pPr.find(qn('w:pStyle'))
            if pStyle is not None:
                style_id = pStyle.get(qn('w:val'))
    except Exception:
        style_id = None

    def _style_by_id(sid):
        for st in styles_el.findall(qn('w:style')):
            if st.get(qn('w:styleId')) == sid:
                return st
        return None

    chain = []
    seen = set()
    while style_id and style_id not in seen:
        seen.add(style_id)
        st = _style_by_id(style_id)
        if st is None:
            break
        chain.append(st)
        base = st.find(qn('w:basedOn'))
        style_id = base.get(qn('w:val')) if base is not None else None
    return chain


def _doc_defaults_rpr(doc):
    """docDefaults 里的 w:rPr（全文默认字体/字号所在处）。"""
    dd = doc.styles.element.find(qn('w:docDefaults'))
    if dd is None:
        return None
    rpd = dd.find(qn('w:rPrDefault'))
    if rpd is None:
        return None
    return rpd.find(qn('w:rPr'))

def _resolve_effective_eastasia(doc, run):
    """按 Word 的字体继承链，解析 run 中文字符实际生效的东亚字体（打开办公软件看到的）：
       run 直接设定(w:eastAsia) → 段落样式(含 basedOn 链) → docDefaults。
       任一层拿到即返回；都没有则返回 None。"""
    rPr = run._element.rPr
    if rPr is not None:
        rFonts = rPr.find(qn('w:rFonts'))
        if rFonts is not None:
            ea = rFonts.get(qn('w:eastAsia'))
            if ea:
                return ea
    for st in _style_chain(doc, run):
        rpr = st.find(qn('w:rPr'))
        if rpr is not None:
            rf = rpr.find(qn('w:rFonts'))
            if rf is not None:
                ea = rf.get(qn('w:eastAsia'))
                if ea:
                    return ea
    rpr = _doc_defaults_rpr(doc)
    if rpr is not None:
        rf = rpr.find(qn('w:rFonts'))
        if rf is not None:
            ea = rf.get(qn('w:eastAsia'))
            if ea:
                return ea
    return None


def _resolve_effective_ascii(doc, run):
    """按 Word 的字体继承链，解析 run 对西文/数字实际生效的字体（打开办公软件看到的）：
       run 直接设定(w:ascii → 回落 w:hAnsi) → 段落样式(含 basedOn 链) → docDefaults。"""
    def _ascii_of(rFonts):
        if rFonts is None:
            return None
        return rFonts.get(qn('w:ascii')) or rFonts.get(qn('w:hAnsi'))

    rPr = run._element.rPr
    if rPr is not None:
        v = _ascii_of(rPr.find(qn('w:rFonts')))
        if v:
            return v
    for st in _style_chain(doc, run):
        rpr = st.find(qn('w:rPr'))
        if rpr is not None:
            v = _ascii_of(rpr.find(qn('w:rFonts')))
            if v:
                return v
    rpr = _doc_defaults_rpr(doc)
    if rpr is not None:
        v = _ascii_of(rpr.find(qn('w:rFonts')))
        if v:
            return v
    return None

def _resolve_effective_size_pt(doc, run):
    """按 Word 的继承链解析 run 实际生效的字号（磅）：
       run 直接设定(w:sz) → 段落样式链 → docDefaults；都没有则返回 None。
       w:sz 单位为半磅，磅值 = val / 2。"""
    if run.font.size is not None:
        return run.font.size.pt
    for st in _style_chain(doc, run):
        rpr = st.find(qn('w:rPr'))
        if rpr is not None:
            sz = rpr.find(qn('w:sz'))
            if sz is not None:
                val = sz.get(qn('w:val'))
                if val and val.replace('.', '', 1).isdigit():
                    return float(val) / 2.0
    rpr = _doc_defaults_rpr(doc)
    if rpr is not None:
        sz = rpr.find(qn('w:sz'))
        if sz is not None:
            val = sz.get(qn('w:val'))
            if val and val.replace('.', '', 1).isdigit():
                return float(val) / 2.0
    return None


def _resolve_effective_bold(doc, run):
    """按 Word 的继承链解析 run 是否加粗：run 直接设定(w:b) → 段落样式链 → docDefaults。"""
    if run.font.bold is not None:
        return bool(run.font.bold)

    def _b_of(rpr):
        if rpr is None:
            return None
        b = rpr.find(qn('w:b'))
        if b is None:
            return None
        val = b.get(qn('w:val'))
        if val is None:
            return True
        return val.lower() not in ('0', 'false', 'off')

    for st in _style_chain(doc, run):
        v = _b_of(st.find(qn('w:rPr')))
        if v is not None:
            return v
    v = _b_of(_doc_defaults_rpr(doc))
    return bool(v)

def _text_runs(para):
    """段落里承载可见文字的 run（跳过纯空白/空 run：它们不承载字体字号效果）。"""
    return [r for r in para.runs if (r.text or '').strip()]


def _p_jc(para):
    """段落对齐方式（w:pPr/w:jc 的值）；未设置返回 None（Word 默认左对齐）。"""
    pPr = para._element.find(qn('w:pPr'))
    if pPr is None:
        return None
    jc = pPr.find(qn('w:jc'))
    if jc is None:
        return None
    return jc.get(qn('w:val'))


def _p_ind(para):
    """段落缩进：返回 (firstLine 磅值, firstLineChars 字符数, right 磅值)，未设置的为 None。"""
    pPr = para._element.find(qn('w:pPr'))
    ind = pPr.find(qn('w:ind')) if pPr is not None else None
    if ind is None:
        return (None, None, None)

    def _num(attr, scale):
        v = ind.get(qn(attr))
        if v is None or not v.lstrip('-').isdigit():
            return None
        return int(v) / scale

    return (_num('w:firstLine', 20.0), _num('w:firstLineChars', 100.0), _num('w:right', 20.0))


def _p_spacing(para):
    """段落行距：返回 (w:line 原值, w:lineRule 值)，未设置的为 None。"""
    pPr = para._element.find(qn('w:pPr'))
    sp = pPr.find(qn('w:spacing')) if pPr is not None else None
    if sp is None:
        return (None, None)
    line = sp.get(qn('w:line'))
    rule = sp.get(qn('w:lineRule'))
    if line is not None and line.lstrip('-').isdigit():
        line = int(line)
    else:
        line = None
    return (line, rule)

def _line_multiple(para):
    """倍数行距（Word“单倍/1.5倍/多倍”）：w:lineRule 为 auto 或缺省时，倍数 = w:line / 240。
       固定值/最小值（exact/atLeast）不是倍数行距，返回 None。"""
    line, rule = _p_spacing(para)
    if line is None:
        return None
    if rule in (None, 'auto'):
        return line / 240.0
    return None


def _is_multiple_spacing(para, target, tol=0.03):
    """段落是否为指定倍数行距（target 倍，容差 ±tol）。"""
    mult = _line_multiple(para)
    return mult is not None and abs(mult - target) <= tol


def _is_single_spacing(para):
    """段落是否为单倍行距：未设置行距即继承默认的单倍；设置了则须为 1 倍。"""
    line, rule = _p_spacing(para)
    if line is None:
        return True
    return _is_multiple_spacing(para, 1.0)


def _first_line_indent_2chars(para, base_size):
    """首行缩进是否为 2 字符（Word 口径）：
       (a) w:firstLineChars == 2 字符（Word“2 字符”的原生存储，与字号无关）；
       (b) w:firstLine 绝对磅值 ÷ 该栏字号 ≈ 2 字符（四号约 28 磅、小四约 24 磅）。"""
    first_pt, first_chars, _ = _p_ind(para)
    if first_chars is not None:
        return 1.8 <= first_chars <= 2.2
    if first_pt is not None and base_size:
        return 1.8 <= (first_pt / base_size) <= 2.2
    return False


def _para_main_size(doc, para):
    """段落主体字号（承载文字最多的 run 的字号）；取不到返回 None。"""
    main = None
    maxlen = 0
    for r in _text_runs(para):
        sz = _resolve_effective_size_pt(doc, r)
        if sz and len(r.text) > maxlen:
            maxlen = len(r.text)
            main = sz
    return main

def _para_has_image(para):
    """段落内是否含图片：新式绘图对象（a:blip）或旧式图片框（w:pict / v:imagedata）都算。"""
    el = para._element
    if el.findall('.//' + qn('a:blip')):
        return True
    if el.findall('.//' + qn('w:pict')):
        return True
    if el.findall('.//' + qn('w:drawing')):
        return True
    return False


def _iter_body_blocks(doc):
    """按文档实际顺序返回正文流中的块：[('p', 段落) | ('t', 表格), ...]。"""
    blocks = []
    for child in doc.element.body.iterchildren():
        if child.tag == qn('w:p'):
            blocks.append(('p', Paragraph(child, doc)))
        elif child.tag == qn('w:tbl'):
            blocks.append(('t', Table(child, doc)))
    return blocks


def _body_paragraphs(blocks):
    return [b for kind, b in blocks if kind == 'p']


def _body_tables(blocks):
    return [b for kind, b in blocks if kind == 't']


def _unique_cells(row):
    """行内去重后的单元格（合并单元格在 row.cells 中会重复出现，这里只取一次）。"""
    out = []
    seen = set()
    for cell in row.cells:
        key = id(cell._tc)
        if key in seen:
            continue
        seen.add(key)
        out.append(cell)
    return out


def _cell_text(cell):
    """单元格文字（多段落拼接后去掉首尾空白）。"""
    return ''.join(p.text for p in cell.paragraphs).strip()


def _cell_text_paras(cell):
    """单元格内含文字的段落。"""
    return [p for p in cell.paragraphs if p.text.strip()]

def _cjk_runs_ok(doc, paras, ea_pred, size_range):
    """段落集合中所有“含中文字符”的 run 是否都满足指定东亚字体与字号区间。
       无可判定的中文 run 时返回 False（要求的中文效果无处落地）。"""
    lo, hi = size_range
    found = False
    for para in paras:
        for r in _text_runs(para):
            if not _has_cjk(r.text):
                continue
            found = True
            if not ea_pred(_resolve_effective_eastasia(doc, r)):
                return False
            sz = _resolve_effective_size_pt(doc, r)
            if sz is None or not (lo <= sz <= hi):
                return False
    return found


def _digit_runs_ok(doc, paras, size_range):
    """段落集合中所有“含阿拉伯数字”的 run 是否都为 Times New Roman 且字号在区间内。
       无数字 run 时返回 False。"""
    lo, hi = size_range
    found = False
    for para in paras:
        for r in _text_runs(para):
            if not _has_digit(r.text):
                continue
            found = True
            if not _is_tnr(_resolve_effective_ascii(doc, r)):
                return False
            sz = _resolve_effective_size_pt(doc, r)
            if sz is None or not (lo <= sz <= hi):
                return False
    return found


def _all_runs_size_ok(doc, paras, size_range):
    """段落集合中所有承载文字的 run 字号是否都落在区间内（空白 run 不计）。"""
    lo, hi = size_range
    found = False
    for para in paras:
        for r in _text_runs(para):
            found = True
            sz = _resolve_effective_size_pt(doc, r)
            if sz is None or not (lo <= sz <= hi):
                return False
    return found


def _all_runs_bold(doc, paras):
    """段落集合中所有承载文字的 run 是否都加粗。"""
    found = False
    for para in paras:
        for r in _text_runs(para):
            found = True
            if not _resolve_effective_bold(doc, r):
                return False
    return found

def _min_run_size(doc, paras):
    """段落集合中承载文字的 run 的最小字号（磅）；无可判定字号返回 None。"""
    sizes = []
    for para in paras:
        for r in _text_runs(para):
            sz = _resolve_effective_size_pt(doc, r)
            if sz:
                sizes.append(sz)
    return min(sizes) if sizes else None


def _open_word_document(file_path):
    """打开 .docx 文档（python-docx 仅支持 OOXML/.docx）。"""
    return Document(file_path)


def _locate_files(dir_path):
    """在目录内定位两份被评估文档：文件名含“鉴定”的为毕业实习鉴定表，
       其余含“报告”的为毕业实习报告（忽略以 ~$ 开头的 Word 临时锁文件）。"""
    cands = sorted(
        f for f in os.listdir(dir_path)
        if f.lower().endswith('.docx') and not f.startswith('~$')
    )
    form_name = next((f for f in cands if '鉴定' in f), None)
    report_name = next((f for f in cands if '报告' in f and f != form_name), None)
    if report_name is None:
        report_name = next((f for f in cands if f != form_name), None)
    return (
        os.path.join(dir_path, report_name) if report_name else None,
        os.path.join(dir_path, form_name) if form_name else None,
    )


def evaluate(dir_path: str) -> dict:
    """统一入口：接收输出文件所在目录，定位并评估目录内的毕业实习报告与毕业实习鉴定表。

    评分口径：维度一为门槛项，未通过则 dim2_items 置空、total_score 记 0，
    max_score 仍为加分项满分之和（MAX_SCORE_TOTAL）；维度二逐条二值判定，
    total_score 为加分项实得之和减去命中的减分项，扣至 0 为止。

    status='error' 只在“这份提交没被真正评过”时出现——目录不存在、目录内没有
    可评估的 Word 文档、或脚本自身抛异常；此时 dim1_pass 为空串（未判定）、
    dim1_reason 为空串、dim2_items 为空表、total_score 与 max_score 一律为 0。
    两份文档只定位到一份属于交付不齐，仍按 status='ok' + 维度一未通过处理，
    此时 dim1_pass=False 且 dim1_reason 带有未通过的错误摘要。

    参数:
        dir_path: 输出文件所在目录路径。
    返回:
        dict，字段对齐统一接口约定 §2.2。
    """
    result: dict[str, object] = {
        "id": SCRIPT_ID,
        "file_name": "",
        "status": "ok",
        "error": None,
        # 空串表示“维度一尚未判定”；status='error' 的三条出口都保持空串，
        # 只有真正跑过 _dim1_check 才会写入 True / False。
        "dim1_pass": "",
        "dim1_reason": "",
        "dim2_items": [],
        "total_score": 0,
        "max_score": 0,
    }

    try:
        if not os.path.isdir(dir_path):
            result["status"] = "error"
            result["error"] = f"目录不存在：{dir_path}"
            return result

        report_path, form_path = _locate_files(dir_path)
        if report_path is None and form_path is None:
            result["status"] = "error"
            result["error"] = f"目录中未找到 .docx 文件：{dir_path}"
            return result

        result["file_name"] = ", ".join(
            os.path.basename(p) for p in (report_path, form_path) if p)
        return _evaluate_impl(report_path, form_path, result)

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

def _doc_all_text(doc):
    """文档全部可编辑文字（正文段落 + 表格单元格）。"""
    parts = []
    for kind, blk in _iter_body_blocks(doc):
        if kind == 'p':
            parts.append(blk.text)
        else:
            for row in blk.rows:
                for cell in _unique_cells(row):
                    parts.append(_cell_text(cell))
    return ''.join(parts)


def _dim1_check(report_path, form_path):
    """维度1（门槛项）：只检查交付文件的格式是否可用、内容是否为可编辑文本。
       返回 (是否通过, 不通过原因, 报告 doc, 鉴定表 doc)。"""
    if not report_path or not form_path:
        missing = []
        if not report_path:
            missing.append('毕业实习报告')
        if not form_path:
            missing.append('毕业实习鉴定表')
        return (False, '缺少 .docx 文档：' + '、'.join(missing), None, None)

    docs = {}
    for label, path in (('毕业实习报告', report_path), ('毕业实习鉴定表', form_path)):
        try:
            docs[label] = _open_word_document(path)
        except Exception as exc:
            return (False, f'{label} 无法正常打开（{type(exc).__name__}）', None, None)

    report_doc = docs['毕业实习报告']
    form_doc = docs['毕业实习鉴定表']

    # 内容须为可编辑文本，而不是图片/扫描件：正文中文字符数须成规模
    if _cjk_count(_doc_all_text(report_doc)) < 100:
        return (False, '毕业实习报告内容不是可编辑文字', None, None)
    if _cjk_count(_doc_all_text(form_doc)) < 100:
        return (False, '毕业实习鉴定表内容不是可编辑文字', None, None)
    # 鉴定表须以可编辑表格承载
    if not _body_tables(_iter_body_blocks(form_doc)):
        return (False, '毕业实习鉴定表中没有可编辑表格', None, None)
    return (True, '', report_doc, form_doc)

REPORT_TITLE_RE = re.compile(r'^题\s*目\s*[：:]')
L1_TITLES = [
    '一、实习背景与单位概况', '二、岗位任务与实践方法', '三、实习过程记录',
    '四、问题处理与能力提升', '五、实习收获与反思', '六、致谢',
    '七、指导教师评语与成绩评定',
]
L2_TITLES = [
    '（一）实习背景', '（二）实习单位概况', '（三）岗位职责说明',
    '（一）一日生活协助', '（二）教学活动支持',
    '（一）适应与观察阶段', '（二）协助与参与阶段', '（三）尝试独立组织阶段',
]
SIGN_LABELS = ('指导教师签名', '日期', '成绩评定', '二级学院')
L1_ANY_RE = re.compile(r'^[一二三四五六七八九十]\s*、')
L2_ANY_RE = re.compile(r'^[（(][一二三四五六七八九十][)）]')
SIGN_LABEL_RE = re.compile(r'^(指导教师签名|日期|成绩评定|二级学院)')


def _norm(text):
    """去掉全部空白（含全角空格）后的文字，用于标题/栏目名的稳定匹配。"""
    return re.sub(r'[\s　]', '', text or '')


def _report_layout(doc):
    """定位报告的各结构位置：封面段、一/二级标题、正文段落、
       “个人情况与实习安排”表、文末落款表等。
       正文页题目（第二处“题目：…”段）本身不再单独判定，但仍作为
       封面区右边界与正文段落的排除项参与定位。"""
    blocks = _iter_body_blocks(doc)
    paras = _body_paragraphs(blocks)
    tables = _body_tables(blocks)

    title_idx = [i for i, p in enumerate(paras) if REPORT_TITLE_RE.match(p.text.strip())]
    cover_title = title_idx[0] if title_idx else -1
    body_title = title_idx[-1] if len(title_idx) >= 2 else -1
    school = -1
    for i, p in enumerate(paras):
        if _norm(p.text) == '江淮文理学院人文与教育学院':
            school = i
    cover_end = max(cover_title, school)
    if body_title >= 0 and cover_end >= body_title:
        cover_end = body_title - 1

    cover_paras = paras[:cover_end + 1] if cover_end >= 0 else []
    cover_report_title = next((p for p in cover_paras if _norm(p.text) == '实习报告'), None)
    cover_unit = next((p for p in cover_paras if p.text.strip().startswith('实习单位')), None)
    cover_date = next((p for p in cover_paras if p.text.strip().startswith('实习日期')), None)
    school_para = paras[school] if school >= 0 else None

    # 一级/二级标题：按细则给出的题名精确匹配
    l1_paras, l2_paras = {}, {}
    for p in paras:
        t = _norm(p.text)
        for name in L1_TITLES:
            if t == _norm(name):
                l1_paras.setdefault(name, []).append(p)
        for name in L2_TITLES:
            if t == _norm(name):
                l2_paras.setdefault(name, []).append(p)

    # 正文段落：不含封面、各级标题、表格、评语标签行与文末落款行
    body_paras = []
    for i, p in enumerate(paras):
        if i <= cover_end or i == body_title:
            continue
        t = p.text.strip()
        if not t:
            continue
        if L1_ANY_RE.match(t) or L2_ANY_RE.match(t):
            continue
        if _norm(t) == '个人情况与实习安排':
            continue
        if t.startswith('指导教师评语'):
            continue
        if SIGN_LABEL_RE.match(t):
            continue
        body_paras.append(p)

    # “个人情况与实习安排”标题与其 8 行 4 列表格
    info_heading = next((p for p in paras if _norm(p.text) == '个人情况与实习安排'), None)
    info_table = next(
        (t for t in tables if len(t.rows) == 8 and len(t.columns) == 4), None)

    # 文末落款表：首列含“指导教师签名”的表格
    sign_table = None
    for t in tables:
        first_col = ''.join(_cell_text(r.cells[0]) for r in t.rows if r.cells)
        if '指导教师签名' in first_col:
            sign_table = t
            break

    # 文末“日期：…”段（供落款对齐类判定使用）
    tail_date = None
    for p in paras:
        if p.text.strip().startswith('日期'):
            tail_date = p

    return {
        'blocks': blocks,
        'paras': paras,
        'tables': tables,
        'cover_paras': cover_paras,
        'cover_title': paras[cover_title] if cover_title >= 0 else None,
        'cover_report_title': cover_report_title,
        'cover_unit': cover_unit,
        'cover_date': cover_date,
        'school_para': school_para,
        'l1_paras': l1_paras,
        'l2_paras': l2_paras,
        'body_paras': body_paras,
        'info_heading': info_heading,
        'info_table': info_table,
        'sign_table': sign_table,
        'tail_date': tail_date,
    }


def _sign_table_rows(table):
    """文末落款表按标签归类：返回 {标签: (标签单元格, 填写值单元格)}。"""
    out = {}
    if table is None:
        return out
    for row in table.rows:
        cells = _unique_cells(row)
        if not cells:
            continue
        label = _cell_text(cells[0])
        value = cells[1] if len(cells) > 1 else None
        for name in SIGN_LABELS:
            if label.startswith(name):
                out[name] = (cells[0], value)
                break
    return out

FORM_INFO_LABELS = ('姓名', '学号', '专业班级', '实习单位', '实习地点',
                    '实习名称', '指导教师', '实习时间')
FORM_ALL_LABELS = FORM_INFO_LABELS + ('实习单位评价', '教学单位评价意见', '自我评价', '备注')
SIGN_OR_DATE_RE = re.compile(r'(实习评分|最终实习成绩|实习单位负责人|指导教师签名|签章|年\s*\d*\s*月)')


def _is_sign_or_date_para(para):
    """该段是否为签名/成绩/落款日期行（与“评价正文”区分开）。"""
    t = para.text.strip()
    return bool(t) and bool(SIGN_OR_DATE_RE.search(t))


def _form_layout(doc):
    """定位鉴定表的标题、主表格，并把主表格各行按栏目归类。"""
    blocks = _iter_body_blocks(doc)
    paras = _body_paragraphs(blocks)
    tables = _body_tables(blocks)

    title_para = next(
        (p for p in paras if _norm(p.text) == '江淮文理学院毕业实习鉴定表'), None)
    if title_para is None:
        title_para = next((p for p in paras if p.text.strip()), None)

    table = None
    for t in tables:
        if len(t.rows) >= 6:
            table = t
            break
    if table is None and tables:
        table = tables[0]

    info_rows, eval_rows = [], []
    self_cells, remark_cells = None, None
    if table is not None:
        for row in table.rows:
            cells = _unique_cells(row)
            if not cells:
                continue
            label = _norm(_cell_text(cells[0]))
            if label in ('实习单位评价', '教学单位评价意见'):
                eval_rows.append((label, cells))
            elif label == '自我评价':
                self_cells = cells
            elif label == '备注':
                remark_cells = cells
            else:
                info_rows.append(cells)

    return {
        'blocks': blocks,
        'paras': paras,
        'tables': tables,
        'title_para': title_para,
        'table': table,
        'info_rows': info_rows,
        'eval_rows': eval_rows,
        'self_cells': self_cells,
        'remark_cells': remark_cells,
    }


def _info_paras(info_rows):
    """信息栏（姓名～实习时间）各单元格内含文字的段落。"""
    out = []
    for cells in info_rows:
        for cell in cells:
            out.extend(_cell_text_paras(cell))
    return out


def _eval_body_paras(cells):
    """评价栏填写值单元格中的“正文段落”（排除签名/成绩/落款日期行）。"""
    if len(cells) < 2:
        return []
    return [p for p in _cell_text_paras(cells[1]) if not _is_sign_or_date_para(p)]


def _eval_all_paras(cells):
    """评价栏填写值单元格中的全部含文字段落（含签名与落款日期行）。"""
    if len(cells) < 2:
        return []
    return _cell_text_paras(cells[1])


def _margins_cm(doc):
    """首节页边距（厘米）：(左, 右, 上, 下)。"""
    sec = doc.sections[0]
    def _cm(v):
        return None if v is None else round(v.cm, 2)
    return (_cm(sec.left_margin), _cm(sec.right_margin),
            _cm(sec.top_margin), _cm(sec.bottom_margin))


def _page_size_cm(doc):
    """首节纸张尺寸（厘米）：(宽, 高)。"""
    sec = doc.sections[0]
    def _cm(v):
        return None if v is None else round(v.cm, 2)
    return (_cm(sec.page_width), _cm(sec.page_height))


def _hf_text(doc, kind):
    """全部节的页眉或页脚文字（含首页/奇偶页版本）拼接后去空白，用于稳定匹配。"""
    names = (('header', 'first_page_header', 'even_page_header') if kind == 'header'
             else ('footer', 'first_page_footer', 'even_page_footer'))
    parts = []
    for sec in doc.sections:
        for nm in names:
            try:
                part = getattr(sec, nm)
            except Exception:
                continue
            parts.extend(p.text for p in part.paragraphs)
    return _norm(''.join(parts))


def _p_space_pt(para):
    """段前/段后间距（磅）：w:spacing/@w:before、@w:after 以二十分之一磅存储。
       注意这与行距 w:line 是不同属性——行距归加分项，段间距归减分项。"""
    pPr = para._element.find(qn('w:pPr'))
    sp = pPr.find(qn('w:spacing')) if pPr is not None else None
    if sp is None:
        return (None, None)

    def _pt(attr):
        v = sp.get(qn(attr))
        if v is None or not v.lstrip('-').isdigit():
            return None
        return int(v) / 20.0

    return (_pt('w:before'), _pt('w:after'))


def _has_page_break(para):
    """段内是否含手动分页符（w:br w:type="page"）。"""
    for br in para._element.findall('.//' + qn('w:br')):
        if br.get(qn('w:type')) == 'page':
            return True
    return False


def _page_break_before(para):
    """段落是否设置了“段前分页”（w:pPr/w:pageBreakBefore）。"""
    pPr = para._element.find(qn('w:pPr'))
    return bool(pPr is not None and pPr.find(qn('w:pageBreakBefore')) is not None)


# 缇（twip，二十分之一磅）→ 厘米
TWIP_TO_CM = 2.54 / 72 / 20.0


def _tbl_jc(table):
    """表格自身的水平对齐（w:tblPr/w:jc）；未设置返回 None。"""
    tblPr = table._tbl.find(qn('w:tblPr'))
    jc = tblPr.find(qn('w:jc')) if tblPr is not None else None
    return None if jc is None else jc.get(qn('w:val'))


def _tbl_width_cm(table):
    """表格总宽（厘米）：w:tblPr/w:tblW 以 dxa（缇）存储时换算；非 dxa 返回 None。"""
    tblPr = table._tbl.find(qn('w:tblPr'))
    tw = tblPr.find(qn('w:tblW')) if tblPr is not None else None
    if tw is None or tw.get(qn('w:type')) != 'dxa':
        return None
    v = tw.get(qn('w:w'))
    if v is None or not v.isdigit():
        return None
    return round(int(v) * TWIP_TO_CM, 2)


def _tbl_grid_cm(table):
    """表格各列列宽（厘米），取自 w:tblGrid/w:gridCol。"""
    grid = table._tbl.find(qn('w:tblGrid'))
    if grid is None:
        return []
    out = []
    for c in grid.findall(qn('w:gridCol')):
        v = c.get(qn('w:w'))
        out.append(round(int(v) * TWIP_TO_CM, 2) if v and v.isdigit() else None)
    return out


def _row_height_cm(row):
    """行高（厘米）：w:trPr/w:trHeight 以缇存储；未设置返回 None。"""
    trPr = row._tr.find(qn('w:trPr'))
    h = trPr.find(qn('w:trHeight')) if trPr is not None else None
    if h is None:
        return None
    v = h.get(qn('w:val'))
    if v is None or not v.isdigit():
        return None
    return round(int(v) * TWIP_TO_CM, 2)


def _cell_valign(cell):
    """单元格内容垂直对齐（w:tcPr/w:vAlign）；未设置返回 None。"""
    tcPr = cell._tc.find(qn('w:tcPr'))
    va = tcPr.find(qn('w:vAlign')) if tcPr is not None else None
    return None if va is None else va.get(qn('w:val'))

# 字号容差（与办公软件“磅”一致；中文字号对应磅值：二号22、三号16、小三15、四号14、小四12）
SZ_ERHAO = (20.5, 22.5)
SZ_SANHAO = (15.5, 16.5)
SZ_XIAOSAN = (14.5, 15.5)
SZ_SIHAO = (13.5, 14.5)
SZ_XIAOSI = (11.5, 12.5)
SZ_XIAOSI_OR_SIHAO = (11.5, 14.5)
SZ_TITLE20 = (19.5, 20.5)
SZ_13PT = (12.5, 13.5)
# 减分项用到的两档字号：报告“个人情况与实习安排”表为五号 10.5 磅，文末落款表为 11 磅。
# 二者仅差 0.5 磅，故取窄容差 ±0.25，避免互相误判。
SZ_WUHAO = (10.25, 10.75)
SZ_11PT = (10.75, 11.25)


def _build_dim2_items(records, deltas_order, kind):
    """按“每条规则一个 item”生成严格二值的评分明细，返回 (items, 小计)。

    records 是判定阶段按结果顺序追加的 (是否命中, 分值, 细则原文) 三元组；
    deltas_order 是同顺序的分值表，只用来交叉校验条目数与分值有没有错位——
    判定代码增删了一条却忘了同步分值表时，这里直接抛错，而不是静默算错分。"""
    if len(records) != len(deltas_order):
        raise RuntimeError('dim2 %s项数不一致：records=%d vs order=%d'
                           % (kind, len(records), len(deltas_order)))
    items = []
    subtotal = 0
    for (hit, delta, rule), max_delta in zip(records, deltas_order):
        if delta != max_delta:
            raise RuntimeError('dim2 %s分值与顺序表不一致：“%s”记为 %d，顺序表为 %d'
                               % (kind, rule, delta, max_delta))
        items.append({
            "rule": rule,
            "max_delta": max_delta,
            "delta": max_delta if hit else 0,
            "hit": hit,
        })
        subtotal += max_delta if hit else 0
    return items, subtotal


def _evaluate_impl(report_path, form_path, result):
    """维度一门槛判定 + 维度二逐项评分；内部把每条判定按结果顺序追加进
       results / penalties，末尾交给 _build_dim2_items 生成结构化 dim2_items。"""
    # ========== 维度1：可用与可修改性（门槛项） ==========
    # 交付 2 个可正常打开的 .docx 文档（毕业实习报告、毕业实习鉴定表），文字与表格
    # 均为可编辑文本；缺任一文档、非 .docx、打不开或内容不可编辑 → 完成度总分记 0。
    dim1_pass, dim1_reason, rd, fd = _dim1_check(report_path, form_path)
    result["dim1_pass"] = bool(dim1_pass)
    # 约定：dim1_pass 为 False 时 dim1_reason 必须带错误摘要，通过时必须为空串。
    if dim1_pass:
        result["dim1_reason"] = ""
    else:
        result["dim1_reason"] = dim1_reason or "维度一未通过（未给出具体原因）"
    if not dim1_pass:
        result["dim2_items"] = []
        result["total_score"] = 0
        result["max_score"] = MAX_SCORE_TOTAL
        return result

    # ========== 维度2：完成度评分 ==========
    rl = _report_layout(rd)
    fl = _form_layout(fd)
    results = []
    penalties = []

    def _emit(ok, delta, rule):
        """追加一条加分项判定：(是否命中, 满分值, 细则原文)。"""
        results.append((bool(ok), delta, rule))

    def _penalize(hit, delta, rule):
        """追加一条减分项判定：(是否命中即确有破坏性改动, 扣分值, 细则原文)。"""
        penalties.append((bool(hit), delta, rule))

    # ---------- 报告封面 ----------

    # +3：报告封面“实习报告”单独成段，宋体、二号（21-22磅）、加粗、居中
    # 判定：封面区存在文字恰为“实习报告”的独立段落；该段承载文字的 run 东亚字体为
    #   宋体/SimSun/NSimSun、字号 21–22 磅、加粗；段落对齐为居中。
    # 与办公软件对齐：字体沿 w:eastAsia 继承链解析，字号读 w:sz（半磅），加粗读 w:b，
    #   对齐读 w:pPr/w:jc，即打开 Word/WPS 看到的效果。
    p = rl['cover_report_title']
    ok = bool(p is not None
              and _cjk_runs_ok(rd, [p], _is_songti, SZ_ERHAO)
              and _all_runs_bold(rd, [p])
              and _p_jc(p) == 'center')
    _emit(ok, 3, '报告封面“实习报告”单独成段，宋体、二号（21-22磅）、加粗、居中')

    # +3：报告封面“题 目：关于在合润童悦幼儿园开展学前教育实践的报告”段为黑体、三号（16磅）、加粗
    # 判定：封面区第一处“题目：…”段（“题”“目”之间允许空格）的中文 run 东亚字体为
    #   黑体/SimHei、字号 15.5–16.5 磅，且该段承载文字的 run 全部加粗。
    p = rl['cover_title']
    ok = bool(p is not None
              and _cjk_runs_ok(rd, [p], _is_heiti, SZ_SANHAO)
              and _all_runs_bold(rd, [p]))
    _emit(ok, 3, '报告封面“题 目：关于在合润童悦幼儿园开展学前教育实践的报告”段为黑体、三号（16磅）、加粗')

    # +3：报告封面“实习单位：合润童悦幼儿园”“实习日期：2025年10月20日至2025年12月26日”两段中文为宋体、四号（14磅）
    # 判定：封面区“实习单位”“实习日期”两段均存在；两段中所有含中文字符的 run，
    #   东亚字体为 宋体/SimSun/NSimSun、字号 13.5–14.5 磅（纯空白 run 不参与判定）。
    cover_info = [x for x in (rl['cover_unit'], rl['cover_date']) if x is not None]
    ok = bool(len(cover_info) == 2 and _cjk_runs_ok(rd, cover_info, _is_songti, SZ_SIHAO))
    _emit(ok, 3, '报告封面“实习单位：合润童悦幼儿园”“实习日期：2025年10月20日至2025年12月26日”两段中文为宋体、四号（14磅）')

    # +1：报告封面“实习单位”“实习日期”两段居中对齐
    # 判定：两段的段落对齐均为居中。
    ok = bool(len(cover_info) == 2 and all(_p_jc(p) == 'center' for p in cover_info))
    _emit(ok, 1, '报告封面“实习单位”“实习日期”两段居中对齐')

    # +3：报告封面“实习单位”“实习日期”两段行距1.5倍
    # 判定：两段均为倍数行距且倍数 = 1.5（w:lineRule 为 auto/缺省时，倍数 = w:line/240，
    #   1.5 倍即 line=360），即办公软件“行距=1.5倍行距”。
    ok = bool(len(cover_info) == 2 and all(_is_multiple_spacing(p, 1.5) for p in cover_info))
    _emit(ok, 3, '报告封面“实习单位”“实习日期”两段行距1.5倍')

    # +3：报告封面“实习单位”“实习日期”两段中的数字（2025、10、20、12、26）为Times New Roman、四号（14磅），同段中文仍为宋体、四号
    # 判定：两段中所有含阿拉伯数字的 run，西文字体为 Times New Roman、字号 13.5–14.5 磅；
    #   同时两段中含中文字符的 run 仍为 宋体、13.5–14.5 磅。
    # 与办公软件对齐：数字实际生效字体沿 w:ascii→w:hAnsi 继承链解析——同一 run 内
    #   西文按 ascii 字体、汉字按 eastAsia 字体分别显示，这与 Word/WPS 的渲染一致。
    ok = bool(len(cover_info) == 2
              and _digit_runs_ok(rd, cover_info, SZ_SIHAO)
              and _cjk_runs_ok(rd, cover_info, _is_songti, SZ_SIHAO))
    _emit(ok, 3, '报告封面“实习单位”“实习日期”两段中的数字（2025、10、20、12、26）为Times New Roman、四号（14磅），同段中文仍为宋体、四号')

    # +1：报告封面“实习单位”“实习日期”冒号后有下划线空白区或连续下划线（如“实习单位：＿＿＿＿”）
    # 判定：两段各自满足其一即可——冒号之后存在带下划线格式的 run（w:u），
    #   或文字中出现连续 2 个以上的下划线字符（＿ 或 _）。
    def _has_underline_blank(doc, para):
        for r in para.runs:
            if r.font.underline:
                return True
        return bool(re.search(r'[＿_]{2,}', para.text or ''))

    ok = bool(len(cover_info) == 2 and all(_has_underline_blank(rd, p) for p in cover_info))
    _emit(ok, 1, '报告封面“实习单位”“实习日期”冒号后有下划线空白区或连续下划线（如“实习单位：＿＿＿＿”）')

    # ---------- 报告正文 ----------

    # +5：报告7个一级标题（一、实习背景与单位概况；二、岗位任务与实践方法；三、实习过程记录；四、问题处理与能力提升；五、实习收获与反思；六、致谢；七、指导教师评语与成绩评定）均为黑体、小三（15磅）、加粗
    # 判定：7 个题名全部出现；每个标题段的中文 run 为黑体/SimHei、字号 14.5–15.5 磅，
    #   且承载文字的 run 全部加粗。任一标题缺失或任一项不符即不得分。
    l1_all = [p for name in L1_TITLES for p in rl['l1_paras'].get(name, [])]
    ok = bool(len(rl['l1_paras']) == len(L1_TITLES)
              and _cjk_runs_ok(rd, l1_all, _is_heiti, SZ_XIAOSAN)
              and _all_runs_bold(rd, l1_all))
    _emit(ok, 5, '报告7个一级标题（一、实习背景与单位概况；二、岗位任务与实践方法；三、实习过程记录；四、问题处理与能力提升；五、实习收获与反思；六、致谢；七、指导教师评语与成绩评定）均为黑体、小三（15磅）、加粗')

    # +3：报告8个二级标题（（一）实习背景；（二）实习单位概况；（三）岗位职责说明；（一）一日生活协助；（二）教学活动支持；（一）适应与观察阶段；（二）协助与参与阶段；（三）尝试独立组织阶段）均为黑体、四号（14磅）、加粗
    # 判定：8 个题名全部出现；每个标题段的中文 run 为黑体/SimHei、字号 13.5–14.5 磅，
    #   且承载文字的 run 全部加粗。
    l2_all = [p for name in L2_TITLES for p in rl['l2_paras'].get(name, [])]
    ok = bool(len(rl['l2_paras']) == len(L2_TITLES)
              and _cjk_runs_ok(rd, l2_all, _is_heiti, SZ_SIHAO)
              and _all_runs_bold(rd, l2_all))
    _emit(ok, 3, '报告8个二级标题（（一）实习背景；（二）实习单位概况；（三）岗位职责说明；（一）一日生活协助；（二）教学活动支持；（一）适应与观察阶段；（二）协助与参与阶段；（三）尝试独立组织阶段）均为黑体、四号（14磅）、加粗')

    # +5：报告正文段落（不含封面、标题、表格）均为宋体、四号（14磅）
    # 判定对象（下同）：封面区之后、表格之外的非空段落，且不含各级标题、
    #   “指导教师评语：”标签行与文末落款/日期行。
    # 判定：这些段落中含中文字符的 run 东亚字体为 宋体/SimSun/NSimSun、字号 13.5–14.5 磅。
    body = rl['body_paras']
    ok = bool(body and _cjk_runs_ok(rd, body, _is_songti, SZ_SIHAO))
    _emit(ok, 5, '报告正文段落（不含封面、标题、表格）均为宋体、四号（14磅）')

    # +3：报告正文段落（不含封面、标题、表格）首行缩进2字符
    # 判定：每个正文段落满足其一——首行缩进以“字符”存储且为 2 字符，
    #   或首行缩进绝对磅值 ÷ 四号（14磅）≈ 2 字符（约 25–31 磅）。
    ok = bool(body and all(_first_line_indent_2chars(p, 14.0) for p in body))
    _emit(ok, 3, '报告正文段落（不含封面、标题、表格）首行缩进2字符')

    # +3：报告正文段落（不含封面、标题、表格）行距1.1倍
    # 判定：每个正文段落为倍数行距且倍数 = 1.1（w:line=264、w:lineRule=auto），
    #   即办公软件“行距=多倍行距 1.1”。
    ok = bool(body and all(_is_multiple_spacing(p, 1.1) for p in body))
    _emit(ok, 3, '报告正文段落（不含封面、标题、表格）行距1.1倍')

    # +3：报告正文段落（不含封面、表格）中文字符数不少于2500
    # 判定：正文段落（含各级标题文字、不含封面与表格）内中文字符累计 ≥ 2500。
    cjk_total = sum(_cjk_count(p.text) for p in body)
    cjk_total += sum(_cjk_count(p.text) for p in l1_all + l2_all)
    ok = cjk_total >= 2500
    _emit(ok, 3, '报告正文段落（不含封面、表格）中文字符数不少于2500（%d字）' % cjk_total)

    # +1：报告含“一、实习背景与单位概况”至“六、致谢”六个一级标题，题序连续无缺项
    # 判定：一～六这 6 个题名全部出现，题序连续、无缺项。
    ok = all(name in rl['l1_paras'] for name in L1_TITLES[:6])
    _emit(ok, 1, '报告含“一、实习背景与单位概况”至“六、致谢”六个一级标题，题序连续无缺项')

    # 文末落款表按标签归类；供下方“4行两列结构”“签名日期留空”与减分项复用。
    sign_rows = _sign_table_rows(rl['sign_table'])

    # ---------- 鉴定表 ----------

    # +5：鉴定表“姓名”“学号”“专业班级”“实习单位”“实习地点”“实习名称”“指导教师”“实习时间”各栏标签与填写值为宋体、四号（14磅）
    # 判定：8 个信息栏标签齐全；这些行的标签与填写值单元格中，含中文字符的 run 东亚字体为
    #   宋体/SimSun/NSimSun，且所有承载文字的 run 字号均为 13.5–14.5 磅（纯空白 run 不计）。
    info_rows = fl['info_rows']
    info_paras = _info_paras(info_rows)
    info_labels = set()
    for cells in info_rows:
        for cell in cells:
            t = _norm(_cell_text(cell))
            if t in FORM_INFO_LABELS:
                info_labels.add(t)
    labels_ok = info_labels >= set(FORM_INFO_LABELS)
    ok = bool(labels_ok
              and _cjk_runs_ok(fd, info_paras, _is_songti, SZ_SIHAO)
              and _all_runs_size_ok(fd, info_paras, SZ_SIHAO))
    _emit(ok, 5, '鉴定表“姓名”“学号”“专业班级”“实习单位”“实习地点”“实习名称”“指导教师”“实习时间”各栏标签与填写值为宋体、四号（14磅）')

    # +3：鉴定表数字（学号25B071384、班级2304、实习时间的2025、10、20、12、26）为Times New Roman、四号（14磅），同栏中文仍为宋体、四号
    # 判定：信息栏中所有含阿拉伯数字的 run，西文字体为 Times New Roman、字号 13.5–14.5 磅；
    #   同时含中文字符的 run 仍为 宋体、13.5–14.5 磅。
    ok = bool(labels_ok
              and _digit_runs_ok(fd, info_paras, SZ_SIHAO)
              and _cjk_runs_ok(fd, info_paras, _is_songti, SZ_SIHAO))
    _emit(ok, 3, '鉴定表数字（学号25B071384、班级2304、实习时间的2025、10、20、12、26）为Times New Roman、四号（14磅），同栏中文仍为宋体、四号')

    # +1：鉴定表“姓名”至“实习时间”各信息栏文字居中对齐
    # 判定：信息栏各单元格内含文字的段落，对齐方式均为居中。
    ok = bool(info_paras and all(_p_jc(p) == 'center' for p in info_paras))
    _emit(ok, 1, '鉴定表“姓名”至“实习时间”各信息栏文字居中对齐')

    # +1：鉴定表“自我评价”栏正文为宋体、小四（12磅）
    # 判定：“自我评价”栏填写值单元格内各段中，含中文字符的 run 东亚字体为
    #   宋体/SimSun/NSimSun、字号 11.5–12.5 磅。
    self_cells = fl['self_cells']
    self_paras = _cell_text_paras(self_cells[1]) if self_cells and len(self_cells) > 1 else []
    ok = bool(self_paras and _cjk_runs_ok(fd, self_paras, _is_songti, SZ_XIAOSI))
    _emit(ok, 1, '鉴定表“自我评价”栏正文为宋体、小四（12磅）')

    # +3：鉴定表“自我评价”栏正文单倍行距、各段首行缩进2字符
    # 判定：两项须同时满足——
    #   行距为单倍（未设行距即继承单倍；设了则须为 w:line=240、lineRule=auto）；
    #   首行缩进以“字符”存储且为 2 字符，或绝对磅值 ÷ 小四（12磅）≈ 2 字符
    #   （约 21.6–26.4 磅）。
    ok = bool(self_paras
              and all(_is_single_spacing(p) for p in self_paras)
              and all(_first_line_indent_2chars(p, 12.0) for p in self_paras))
    _emit(ok, 3, '鉴定表“自我评价”栏正文单倍行距、各段首行缩进2字符')

    # +3：鉴定表“实习单位评价”“教学单位评价意见”两栏正文为宋体小四（12磅）或四号（14磅）、单倍行距、首行缩进2字符
    # 判定：两栏均存在；各栏填写值单元格中的正文段落（不含签名、成绩与落款日期行）
    #   含中文的 run 为宋体、字号 11.5–14.5 磅；行距为单倍；首行缩进 2 字符
    #   （按该段主体字号折算，小四约 24 磅、四号约 28 磅）。
    eval_rows = fl['eval_rows']
    eval_body = []
    for _label, cells in eval_rows:
        eval_body.append(_eval_body_paras(cells))
    ok = bool(len(eval_rows) == 2 and all(eval_body))
    if ok:
        for paras in eval_body:
            if not _cjk_runs_ok(fd, paras, _is_songti, SZ_XIAOSI_OR_SIHAO):
                ok = False
                break
            if not all(_is_single_spacing(p) for p in paras):
                ok = False
                break
            if not all(_first_line_indent_2chars(p, _para_main_size(fd, p) or 12.0) for p in paras):
                ok = False
                break
    _emit(ok, 3, '鉴定表“实习单位评价”“教学单位评价意见”两栏正文为宋体小四（12磅）或四号（14磅）、单倍行距、首行缩进2字符')

    # +1：鉴定表“实习单位评价”“教学单位评价意见”两栏正文居中对齐
    # 判定：两栏正文段落对齐方式均为居中。
    ok = bool(len(eval_rows) == 2 and all(eval_body)
              and all(_p_jc(p) == 'center' for paras in eval_body for p in paras))
    _emit(ok, 1, '鉴定表“实习单位评价”“教学单位评价意见”两栏正文居中对齐')

    # +1：鉴定表为8行表格、含合并单元格，“姓名”“学号”“专业班级”“实习单位”“实习地点”“实习名称”“指导教师”“实习时间”“实习单位评价”“教学单位评价意见”“自我评价”“备注”栏目齐全
    # 判定：主表格为 8 行；表内存在合并单元格（横向合并 w:gridSpan 或纵向合并 w:vMerge）；
    #   12 个栏目名在表格文字中全部出现。
    tbl = fl['table']
    ok = False
    if tbl is not None:
        merged = bool(tbl._tbl.findall('.//' + qn('w:gridSpan'))
                      or tbl._tbl.findall('.//' + qn('w:vMerge')))
        table_text = _norm(''.join(
            _cell_text(c) for row in tbl.rows for c in _unique_cells(row)))
        ok = bool(len(tbl.rows) == 8 and merged
                  and all(_norm(name) in table_text for name in FORM_ALL_LABELS))
    _emit(ok, 1, '鉴定表为8行表格、含合并单元格，“姓名”“学号”“专业班级”“实习单位”“实习地点”“实习名称”“指导教师”“实习时间”“实习单位评价”“教学单位评价意见”“自我评价”“备注”栏目齐全')

    # +1：鉴定表“备注”栏文字为宋体、四号（14磅）
    # 判定：备注栏各段主体文字（承载字符最多的 run）字号为 13.5–14.5 磅、东亚字体为宋体——
    #   标点等零星短 run 不改变整段的显示效果，故以主体文字为准。
    # 归属说明：要求图片对表格“文字部分”统一批注了宋体四号，故字体字号归加分项；
    #   两条说明文字是否被删除属于“原有内容不得被破坏”，交由减分项承担。
    def _main_cjk_run(para):
        best, blen = None, 0
        for r in _text_runs(para):
            if _has_cjk(r.text) and len(r.text) > blen:
                best, blen = r, len(r.text)
        return best

    remark_cells = fl['remark_cells']
    remark_paras = (_cell_text_paras(remark_cells[1])
                    if remark_cells and len(remark_cells) > 1 else [])
    remark_text = ''.join(p.text for p in remark_paras)
    ok = bool(remark_paras)
    if ok:
        for p in remark_paras:
            sz = _para_main_size(fd, p)
            if not (sz and SZ_SIHAO[0] <= sz <= SZ_SIHAO[1]):
                ok = False
                break
            mr = _main_cjk_run(p)
            if mr is not None and not _is_songti(_resolve_effective_eastasia(fd, mr)):
                ok = False
                break
    _emit(ok, 1, '鉴定表“备注”栏文字为宋体、四号（14磅）')

    # ---------- 减分项：指令未涉及、却破坏了文档原有效果的改动 ----------
    # 与加分项的分工：同一区域按属性切开，加分项管字体/字号/加粗/行距/首行缩进/字数，
    # 减分项管页面设置、分页、对齐、右缩进、段间距、表格几何与垂直对齐，两侧不重叠。
    # 前置守卫：属性类减分项在判定对象缺失时一律判为未命中，缺失只由第 14 项
    # “报告缺少‘个人情况与实习安排’标题或其8行4列表格”去扣，保证一处破坏只扣一次分。

    # -3：报告页边距不为左2.70cm、右2.50cm、上2.50cm、下2.30cm，或鉴定表页边距不为左右3.17cm、上下2.54cm
    # 判定：读取首节 w:pgMar，与目标值比较（容差 ±0.05cm，吸收单位换算误差）。
    def _margin_off(doc, want):
        got = _margins_cm(doc)
        for g, wv in zip(got, want):
            if g is None or abs(g - wv) > 0.05:
                return True
        return False

    hit = bool(_margin_off(rd, (2.70, 2.50, 2.50, 2.30))
               or _margin_off(fd, (3.17, 3.17, 2.54, 2.54)))
    _penalize(hit, -3, '报告页边距不为左2.70cm、右2.50cm、上2.50cm、下2.30cm，或鉴定表页边距不为左右3.17cm、上下2.54cm')

    # -3：报告或鉴定表纸张尺寸不为A4（21.0cm×29.7cm）
    # 判定：读取首节 w:pgSz 的宽高，排序后与 21.0×29.7 比较（容差 ±0.05cm），
    #   这样横纵向声明不影响判定——只看纸张本身是不是 A4。
    def _not_a4(doc):
        wcm, hcm = _page_size_cm(doc)
        if wcm is None or hcm is None:
            return True
        short, long_ = sorted((wcm, hcm))
        return not (abs(short - 21.0) <= 0.05 and abs(long_ - 29.7) <= 0.05)

    hit = bool(_not_a4(rd) or _not_a4(fd))
    _penalize(hit, -3, '报告或鉴定表纸张尺寸不为A4（21.0cm×29.7cm）')

    # -3：报告页眉“江淮文理学院毕业实习报告”或页脚“毕业实习材料”被删除或改写
    # 判定：汇总报告全部节的页眉（含首页/奇偶页版本）文字，须含“江淮文理学院毕业实习报告”；
    #   页脚文字须含“毕业实习材料”。任一缺失或被改写即命中。
    hit = not ('江淮文理学院毕业实习报告' in _hf_text(rd, 'header')
               and '毕业实习材料' in _hf_text(rd, 'footer'))
    _penalize(hit, -3, '报告页眉“江淮文理学院毕业实习报告”或页脚“毕业实习材料”被删除或改写')

    # -3：报告封面与“个人情况与实习安排”之间的分页被删除，封面不再单独成页
    # 判定：封面首段至“个人情况与实习安排”标题之间存在手动分页符（w:br w:type="page"），
    #   或该标题段自带“段前分页”（w:pageBreakBefore），二者有其一即视为分页仍在。
    # 前置守卫：标题缺失时不命中，缺失交由第 14 项去扣。
    hit = False
    info_heading = rl['info_heading']
    if info_heading is not None:
        h_idx = next((i for i, p in enumerate(rl['paras']) if p is info_heading), -1)
        if h_idx >= 0:
            hit = not (any(_has_page_break(p) for p in rl['paras'][:h_idx])
                       or _page_break_before(info_heading))
    _penalize(hit, -3, '报告封面与“个人情况与实习安排”之间的分页被删除，封面不再单独成页')

    # -1：报告封面区空段落少于8个，封面留白版面被压缩
    # 判定：封面区（首段至“江淮文理学院 人文与教育学院”落款段）内不含文字的段落
    #   少于 8 个即命中——这些空段是撑起封面留白的版面结构。
    #   下限取自输入模板实测：模板封面区 13 段、其中 8 段为空。
    # 前置守卫：定位不到封面区时不命中。
    cover_all = rl['cover_paras']
    hit = bool(cover_all and sum(1 for p in cover_all if not p.text.strip()) < 8)
    _penalize(hit, -1, '报告封面区空段落少于8个，封面留白版面被压缩')

    # -1：报告封面校徽图片被删除，或“江淮文理学院 人文与教育学院”段字号不为13磅、不为居中
    # 判定：三种破坏任一出现即命中——封面区不再含图片（新式绘图对象与旧式图片框都算）；
    #   学院落款段被整段删除；该段主体字号不在 12.5–13.5 磅；对齐不为居中。
    #   行距不参与判定：未设行距即继承单倍，该判定几乎不可能失败，留着只会虚增条目。
    # 归属说明：要求图片未标注这行的字号/对齐，故它们不作为加分项，只按
    #   “原有效果不得被破坏”由本项承担；本项是该区域唯一的判定，缺失与属性同归它，
    #   不存在与其他项重复扣分。
    # 前置守卫：定位不到封面区时不命中。
    school_p = rl['school_para']
    hit = False
    if cover_all:
        if not any(_para_has_image(x) for x in cover_all) or school_p is None:
            hit = True
        else:
            sz = _para_main_size(rd, school_p)
            hit = bool(not (sz and SZ_13PT[0] <= sz <= SZ_13PT[1])
                       or _p_jc(school_p) != 'center')
    _penalize(hit, -1, '报告封面校徽图片被删除，或“江淮文理学院 人文与教育学院”段字号不为13磅、不为居中')

    # -3：报告一级标题或二级标题的对齐方式不为左对齐（被改为居中、两端对齐或右对齐）
    # 判定：各级标题段的 w:jc 须为 left/start 或未设置（未设置即继承默认左对齐）；
    #   出现 center/both/right/distribute 即命中。加分项只管标题的字体字号加粗，不涉对齐。
    # 前置守卫：标题全部缺失时 title_paras 为空，不命中。
    title_paras = l1_all + l2_all
    hit = any(_p_jc(p) not in (None, 'left', 'start') for p in title_paras)
    _penalize(hit, -3, '报告一级标题或二级标题的对齐方式不为左对齐（被改为居中、两端对齐或右对齐）')

    # -1：报告一级标题或二级标题带首行缩进
    # 判定：标题段的 w:ind/@w:firstLineChars 或 @w:firstLine 大于 0 即命中
    #   （原效果是顶格，首行缩进只属于正文段落）。
    hit = False
    for p in title_paras:
        first_pt, first_chars, _r = _p_ind(p)
        if (first_chars or 0) > 0 or (first_pt or 0) > 0:
            hit = True
            break
    _penalize(hit, -1, '报告一级标题或二级标题带首行缩进')

    # -1：报告一级标题“一、实习背景与单位概况”或二级标题“（一）实习背景”的段后间距不为4磅
    # 判定：w:spacing/@w:after ÷ 20 须为 4 磅（容差 ±0.2 磅）。段后间距与行距是不同属性，
    #   行距归加分项，段间距归减分项。
    # 前置守卫：这两个标题缺失时取不到段落，不命中。
    hit = False
    for _name, _bag in (('一、实习背景与单位概况', rl['l1_paras']),
                        ('（一）实习背景', rl['l2_paras'])):
        for p in _bag.get(_name, []):
            _bef, aft = _p_space_pt(p)
            if aft is None or abs(aft - 4.0) > 0.2:
                hit = True
    _penalize(hit, -1, '报告一级标题“一、实习背景与单位概况”或二级标题“（一）实习背景”的段后间距不为4磅')

    # -1：报告正文段落为两端对齐，或带右缩进
    # 判定：正文段落中出现 w:jc=both（两端对齐）或 w:ind/@w:right 大于 0 即命中。
    hit = False
    for p in rl['body_paras']:
        _f, _fc, right_pt = _p_ind(p)
        if _p_jc(p) == 'both' or (right_pt or 0) > 0:
            hit = True
            break
    _penalize(hit, -1, '报告正文段落为两端对齐，或带右缩进')

    # -1：报告“三、实习过程记录”下“实习初期，我首先熟悉园所作息…”段的段前或段后间距不为0磅
    # 判定：该段 w:spacing 的 @w:before、@w:after 折算后须均为 0 磅（未设置也算 0）。
    # 前置守卫：该段缺失时不命中。
    hit = False
    for p in rl['body_paras']:
        if p.text.strip().startswith('实习初期'):
            bef, aft = _p_space_pt(p)
            if (bef or 0) != 0 or (aft or 0) != 0:
                hit = True
    _penalize(hit, -1, '报告“三、实习过程记录”下“实习初期，我首先熟悉园所作息…”段的段前或段后间距不为0磅')

    # -1：报告“指导教师评语：”行不为加粗
    # 判定：该标签行中承载文字的 run 全部加粗；不满足即命中。
    #   字体字号不参与判定：要求图片对这行只标了“教师评语需要手写”，未标字体字号。
    # 前置守卫：该行缺失时不命中。
    hit = False
    comment_label = [p for p in rl['paras'] if p.text.strip().startswith('指导教师评语')]
    if comment_label:
        hit = not _all_runs_bold(rd, comment_label)
    _penalize(hit, -1, '报告“指导教师评语：”行不为加粗')

    # -1：报告“指导教师评语：”行出现小于小四（12磅）的文字
    # 判定：该标签行中承载文字的 run 的最小字号 < 12 磅即命中。
    #   下限取自输入模板实测：模板这行为 12.0 磅，故 12 磅是“原有效果”的底线；
    #   只设下限不设上限，是因为要求图片把正文放大到了四号，这行随正文一起放大
    #   属于按要求改动，不是破坏，只有把它缩小才是。
    #   字体不参与判定：模板这行的西文名为 Noto Sans CJK SC，既不是宋体也不是黑体，
    #   写死任一字体都会与模板原值冲突。
    # 前置守卫：该行缺失时不命中。
    hit = False
    if comment_label:
        min_sz = _min_run_size(rd, comment_label)
        hit = bool(min_sz is not None and min_sz < 11.95)
    _penalize(hit, -1, '报告“指导教师评语：”行出现小于小四（12磅）的文字')

    # -3：报告缺少“个人情况与实习安排”标题或其8行4列表格
    # 判定：缺少该标题段，或缺少 8 行 4 列的表格，任一缺失即命中。
    #   这是该表唯一的“存在性”减分项，第 4、15、16、17 项在它缺失时一律不命中。
    hit = bool(rl['info_heading'] is None or rl['info_table'] is None)
    _penalize(hit, -3, '报告缺少“个人情况与实习安排”标题或其8行4列表格')

    # -3：报告“个人情况与实习安排”表非表头单元格文字加粗，或成绩填写值“良好”未加粗
    # 判定：该表首行为表头，其余各行的填写值单元格（第 2、4 列）若出现加粗即命中；
    #   文末“成绩评定”对应的填写值“良好”若未加粗也命中。
    # 前置守卫：表格缺失时不检查加粗；“成绩评定”栏缺失时不检查“良好”。
    hit = False
    itbl = rl['info_table']
    if itbl is not None:
        for row in list(itbl.rows)[1:]:
            cells = _unique_cells(row)
            for idx in (1, 3):
                if idx >= len(cells):
                    continue
                for p in _cell_text_paras(cells[idx]):
                    for r in _text_runs(p):
                        if _resolve_effective_bold(rd, r):
                            hit = True
    grade_cell = sign_rows.get('成绩评定')
    grade_val = grade_cell[1] if grade_cell else None
    if grade_val is not None:
        grade_paras = [p for p in _cell_text_paras(grade_val) if '良好' in p.text]
        if grade_paras and not _all_runs_bold(rd, grade_paras):
            hit = True
    _penalize(hit, -3, '报告“个人情况与实习安排”表非表头单元格文字加粗，或成绩填写值“良好”未加粗')

    # -1：报告“个人情况与实习安排”表不为水平居中，或4列列宽不为等宽3.95cm
    # 判定：表格 w:tblPr/w:jc 须为 center；w:tblGrid 须为 4 列、每列 3.95cm（容差 ±0.05cm，
    #   缇值 2239/2239/2239/2238 换算而来）。
    # 前置守卫：表格缺失时不命中。
    hit = False
    if itbl is not None:
        icols = _tbl_grid_cm(itbl)
        hit = bool(_tbl_jc(itbl) != 'center'
                   or len(icols) != 4
                   or any(c is None or abs(c - 3.95) > 0.05 for c in icols))
    _penalize(hit, -1, '报告“个人情况与实习安排”表不为水平居中，或4列列宽不为等宽3.95cm')

    # -1：报告“个人情况与实习安排”表“学生姓名”“林沐言”两格文字字号不为五号（10.5磅）
    # 判定：这两格内承载文字的 run 字号须为 10.25–10.75 磅。
    # 前置守卫：表格或这两格缺失时不命中。
    hit = False
    if itbl is not None:
        probe_paras = []
        for row in itbl.rows:
            for c in _unique_cells(row):
                if _norm(_cell_text(c)) in ('学生姓名', '林沐言'):
                    probe_paras.extend(_cell_text_paras(c))
        if probe_paras:
            hit = not _all_runs_size_ok(rd, probe_paras, SZ_WUHAO)
    _penalize(hit, -1, '报告“个人情况与实习安排”表“学生姓名”“林沐言”两格文字字号不为五号（10.5磅）')

    # -3：报告文末“日期：2026年1月13日”段不为右对齐，或带首行缩进
    # 判定：该段（落款表之后的独立段落，不在表格内）w:jc 须为 right，且无首行缩进。
    #   加分项“文末4行两列结构”只看表内标签与填写值的对齐，不涉及本段，二者不重叠。
    # 前置守卫：该段缺失时不命中。
    hit = False
    tail_date = rl['tail_date']
    if tail_date is not None:
        first_pt, first_chars, _r = _p_ind(tail_date)
        hit = bool(_p_jc(tail_date) != 'right'
                   or (first_chars or 0) > 0 or (first_pt or 0) > 0)
    _penalize(hit, -3, '报告文末“日期：2026年1月13日”段不为右对齐，或带首行缩进')

    # -1：报告文末落款表“指导教师签名：”“日期：”“成绩评定：”“二级学院（签章）：”4个标签及“良好”填写值的字号不为11磅
    # 判定：4 个标签单元格与“成绩评定”填写值“良好”所在段落的 run 字号须为 10.75–11.25 磅。
    #   该表原本用的是 11 磅而非四号，故只判字号是否被改动，不判是否宋体四号。
    # 前置守卫：落款表缺失或取不到标签时不命中。
    hit = False
    if rl['sign_table'] is not None and sign_rows:
        tail_probe = []
        for name in SIGN_LABELS:
            pair = sign_rows.get(name)
            if pair:
                tail_probe.extend(_cell_text_paras(pair[0]))
        gp = sign_rows.get('成绩评定')
        if gp and gp[1] is not None:
            tail_probe.extend(p for p in _cell_text_paras(gp[1]) if '良好' in p.text)
        if tail_probe:
            hit = not _all_runs_size_ok(rd, tail_probe, SZ_11PT)
    _penalize(hit, -1, '报告文末落款表“指导教师签名：”“日期：”“成绩评定：”“二级学院（签章）：”4个标签及“良好”填写值的字号不为11磅')

    # -1：报告文末落款不为4行两列结构，或“指导教师签名”“日期”“成绩评定”“二级学院（签章）”标签单元格不为右对齐、填写值单元格不为左对齐
    # 判定：落款表须为 4 行 2 列且 4 个标签齐全；标签单元格内各段右对齐，填写值单元格内
    #   各段左对齐（含空段，空段也承载对齐效果）。结构或对齐任一不符即命中。
    # 归属说明：要求图片只画出“指导教师签名／日期／成绩评定”三行，既没有“二级学院（签章）”，
    #   也未标注单元格对齐，故整条按“原有结构与效果不得被破坏”计入减分项。
    # 本项是落款表的存在性判定：表被整体删除即命中，第 19 项字号项则前置守卫为不命中。
    st = rl['sign_table']
    hit = True
    if st is not None and len(st.rows) == 4 and len(st.columns) == 2 and len(sign_rows) == 4:
        hit = False
        for name in SIGN_LABELS:
            label_cell, value_cell = sign_rows[name]
            if any(_p_jc(p) != 'right' for p in label_cell.paragraphs):
                hit = True
                break
            if value_cell is None or any(_p_jc(p) != 'left' for p in value_cell.paragraphs):
                hit = True
                break
    _penalize(hit, -1, '报告文末落款不为4行两列结构，或“指导教师签名”“日期”“成绩评定”“二级学院（签章）”标签单元格不为右对齐、填写值单元格不为左对齐')

    # -1：鉴定表标题“江淮文理学院毕业实习鉴定表”不为黑体、20磅、加粗、居中
    # 判定：标题段文字须为“江淮文理学院毕业实习鉴定表”，中文 run 东亚字体为 黑体/SimHei、
    #   字号 19.5–20.5 磅、加粗，且段落居中；任一不符即命中。
    # 归属说明：要求图片对鉴定表标题没有任何红字批注，黑体/20磅/加粗只能靠实测原文档得出，
    #   故整条按“原有效果不得被破坏”计入减分项。
    # 前置守卫：定位不到标题段时不命中。
    tp = fl['title_para']
    hit = False
    if tp is not None:
        hit = not (_norm(tp.text) == '江淮文理学院毕业实习鉴定表'
                   and _cjk_runs_ok(fd, [tp], _is_heiti, SZ_TITLE20)
                   and _all_runs_bold(fd, [tp])
                   and _p_jc(tp) == 'center')
    _penalize(hit, -1, '鉴定表标题“江淮文理学院毕业实习鉴定表”不为黑体、20磅、加粗、居中')

    # -3：鉴定表“实习评分：良好”“（签章）”“最终实习成绩：良好”及两处落款日期行出现小于11磅的文字
    # 判定：两个评价栏中的签名/成绩/落款日期行（由 SIGN_OR_DATE_RE 识别，正是加分项
    #   “两栏正文”用 _eval_body_paras 排除掉的那批段落），其最小字号 < 11 磅即命中。
    #   下限取自输入模板实测：模板这 5 段均为 11.0 磅，故 11 磅是“原有效果”的底线，
    #   保留模板原值不应被扣分；低于 11 磅才是把签名区文字缩小了。
    # 前置守卫：取不到这类段落时不命中。
    sign_date_paras = []
    for _label, cells in eval_rows:
        sign_date_paras.extend(
            p for p in _eval_all_paras(cells) if _is_sign_or_date_para(p))
    hit = False
    if sign_date_paras:
        min_sz = _min_run_size(fd, sign_date_paras)
        hit = bool(min_sz is not None and min_sz < 10.95)
    _penalize(hit, -3, '鉴定表“实习评分：良好”“（签章）”“最终实习成绩：良好”及两处落款日期行出现小于11磅的文字')

    # -3：鉴定表“自我评价”栏正文对齐方式不为居中
    # 判定：“自我评价”栏填写值单元格内各段的 w:jc 须为 center；出现其他值即命中。
    #   加分项只管该栏的字体/字号/行距/首行缩进，不涉对齐。
    # 前置守卫：该栏缺失或无文字段落时不命中。
    hit = bool(self_paras and any(_p_jc(p) != 'center' for p in self_paras))
    _penalize(hit, -3, '鉴定表“自我评价”栏正文对齐方式不为居中')

    # -1：鉴定表“备注”栏两条说明文字被删除，或不为左对齐、不为单倍行距
    # 判定：备注栏须含“实习活动内容应真实”“实习鉴定的评价应客观”两条说明；各段 w:jc 须为
    #   left/start 或未设置，且行距为单倍。三者任一不符即命中。
    # 归属说明：两条说明是原文档已有内容、要求图片的备注栏是空的，故其存在性归减分项；
    #   加分项只管备注文字的字体字号，不涉内容、对齐与行距。
    # 前置守卫：整个“备注”栏（行/单元格）都不存在时不命中——那属于表格栏目缺失，
    #   由加分项“鉴定表8行表格、栏目齐全”承担；栏还在、文字被清空则由本项命中。
    hit = bool(remark_cells and len(remark_cells) > 1
               and ('实习活动内容应真实' not in remark_text
                    or '实习鉴定的评价应客观' not in remark_text
                    or any(_p_jc(p) not in (None, 'left', 'start')
                           or not _is_single_spacing(p) for p in remark_paras)))
    _penalize(hit, -1, '鉴定表“备注”栏两条说明文字被删除，或不为左对齐、不为单倍行距')

    # -3：鉴定表表格总宽不为16.06cm，或6列列宽不为2.22cm、2.86cm、1.56cm、2.99cm、2.42cm、4.01cm
    # 判定：w:tblW 须为 dxa 且折算后 16.06cm；w:tblGrid 须为 6 列且逐列匹配上述宽度
    #   （容差 ±0.05cm；缇值 9102 与 1260/1623/882/1694/1371/2271 换算而来）。
    #   加分项只看行数、是否含合并单元格与栏目是否齐全，不涉表宽列宽。
    # 前置守卫：主表格缺失时不命中。
    hit = False
    if tbl is not None:
        fw = _tbl_width_cm(tbl)
        fcols = _tbl_grid_cm(tbl)
        want_cols = (2.22, 2.86, 1.56, 2.99, 2.42, 4.01)
        hit = bool(fw is None or abs(fw - 16.06) > 0.05
                   or len(fcols) != len(want_cols)
                   or any(c is None or abs(c - t) > 0.05
                          for c, t in zip(fcols, want_cols)))
    _penalize(hit, -3, '鉴定表表格总宽不为16.06cm，或6列列宽不为2.22cm、2.86cm、1.56cm、2.99cm、2.42cm、4.01cm')

    # -1：鉴定表“姓名”“实习单位”“实习名称”“实习时间”四个信息行行高小于0.70cm，或“实习单位评价”“教学单位评价意见”“自我评价”三栏行高小于5.30cm
    # 判定：按各行首列栏目名归类，读 w:trPr/w:trHeight（缇）折算成厘米后比较下限。
    #   下限取自输入模板实测：模板信息行 0.72/0.80/0.77/0.79cm、三栏 6.02/5.37/6.00cm，
    #   故分别取 0.70cm 与 5.30cm 作为“原有版面不得被压缩”的底线。
    # 前置守卫：主表格缺失、或某行未设行高时该行不参与判定。
    hit = False
    if tbl is not None:
        for row in tbl.rows:
            cells = _unique_cells(row)
            if not cells:
                continue
            label = _norm(_cell_text(cells[0]))
            hcm = _row_height_cm(row)
            if hcm is None:
                continue
            if label in ('姓名', '实习单位', '实习名称', '实习时间') and hcm < 0.70:
                hit = True
            if label in ('实习单位评价', '教学单位评价意见', '自我评价') and hcm < 5.30:
                hit = True
    _penalize(hit, -1, '鉴定表“姓名”“实习单位”“实习名称”“实习时间”四个信息行行高小于0.70cm，或“实习单位评价”“教学单位评价意见”“自我评价”三栏行高小于5.30cm')

    # -1：鉴定表单元格内容不为垂直居中
    # 判定：主表格内各单元格 w:tcPr/w:vAlign 须为 center；出现其他值或未设置即命中。
    #   垂直对齐（w:vAlign）与加分项判的水平对齐（w:jc）是不同属性，二者不重叠。
    # 前置守卫：主表格缺失时不命中。
    hit = False
    if tbl is not None:
        for row in tbl.rows:
            for c in _unique_cells(row):
                if _cell_valign(c) != 'center':
                    hit = True
    _penalize(hit, -1, '鉴定表单元格内容不为垂直居中')

    # ========== 生成结构化 dim2_items ==========
    # 加分项在前、减分项在后，与 _MAX_DELTAS_ORDER / _PENALTY_DELTAS_ORDER 同序。
    pos_items, gained = _build_dim2_items(results, _MAX_DELTAS_ORDER, '加分')
    neg_items, lost = _build_dim2_items(penalties, _PENALTY_DELTAS_ORDER, '减分')

    result["dim2_items"] = pos_items + neg_items
    result["total_score"] = max(0, gained + lost)
    result["max_score"] = MAX_SCORE_TOTAL
    return result


if __name__ == "__main__":
    _dir = sys.argv[1] if len(sys.argv) > 1 else os.path.dirname(os.path.abspath(__file__))
    _output = json.dumps(evaluate(_dir), ensure_ascii=False, indent=2) + "\n"
    sys.stdout.buffer.write(_output.encode("utf-8", errors="replace"))
