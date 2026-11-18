#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# pyright: reportUnknownMemberType=false, reportUnknownVariableType=false, reportUnknownArgumentType=false, reportUnknownParameterType=false, reportMissingParameterType=false, reportMissingTypeArgument=false, reportPrivateUsage=false, reportUnknownLambdaType=false, reportGeneralTypeIssues=false
"""
青禾公共书房网络化运营与更新规划书 排版格式自动评估工具（officeval_018）

统一入口：
    evaluate(dir_path: str) -> dict
    - dir_path: 输出文件所在目录（内含待评估的 .docx，忽略以 ~$ 开头的 Word 临时锁文件）
    - 返回字段：id / file_name / status / error / dim1_pass / dim1_reason /
                dim2_items（命中与未命中项都返回）/ total_score / max_score
    - dim2_items 每项只含 rule / max_delta / delta / hit 四个键

规范约束：
    * 只暴露 evaluate 一个函数供批量运行器调用
    * 不修改 sys.stdout、不打印主结果、不使用 sys.exit
    * 不硬编码文档路径，全部由 dir_path 参数驱动
    * status='error' 表示这份提交没被真正评过（目录不存在、目录内无可评估文档、
      脚本自身异常），此时 dim1_pass 为空串（维度一未判定）、
      total_score 与 max_score 一律为 0；
      文档能定位到但不达标一律 status='ok'，由 dim1_pass / dim2_items 承载
    * dim1_pass 为 False 时 dim1_reason 必须给出未通过的错误摘要；通过时为空串
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

SCRIPT_ID = "018"
# 维度二加分项 max_delta 之和；用于维度一未通过 / 脚本异常时兜底
MAX_SCORE_TOTAL = 52

# 维度二加分项按“结果追加顺序”的 max_delta 列表；末端生成结构化 dim2_items 时
# 与 results 逐项对位。顺序须与 _evaluate_impl 中 _emit(...) 的调用顺序严格一致；
# 共 16 项，累加 = MAX_SCORE_TOTAL。
#
# 加分项只收录《报告排版格式要求》里明确写出的排版要求：题目与各级标题、正文、
# 图注表注、表格内文字的字体字号，正文行距与段前间距，层次编号，以及图注表注的
# 位置、连续编号与“图序与图题之间空两格”。
_MAX_DELTAS_ORDER = [
    1,  # 01 大标题单倍行距
    1,  # 02 摘要标题黑体四号
    5,  # 03 五级层次编号依次嵌套
    5,  # 04 七个一级标题黑体小三
    5,  # 05 正文宋体小四、无五号
    3,  # 06 正文行距 1.25 倍
    3,  # 07 正文段前 0.5 行
    1,  # 08 两处指定正文段落不加粗
    3,  # 09 图注宋体小四居中、位于插图下方
    3,  # 10 带图名的图注编号与图名之间空两格
    5,  # 11 15 张插图均配图注、编号连续
    3,  # 12 表注宋体小四居中、位于表格上方
    3,  # 13 带表名的表注编号与表名之间空两格
    5,  # 14 25 张表格均配表注、编号连续
    3,  # 15 表格内文字宋体小四
    3,  # 16 表格内文字单倍行距
]
assert sum(_MAX_DELTAS_ORDER) == MAX_SCORE_TOTAL

# 维度二减分项按“结果追加顺序”的 max_delta 列表（均为负值）；共 14 项。
# 减分项只在命中“指令未涉及、却破坏了文档原有效果”的情形时扣分，不计入 max_score，
# 但会从 total_score 中扣除（扣至 0 为止）。顺序须与 _penalize(...) 的调用顺序一致。
#
# 三条设计约束：
# ① 加分项与减分项不得覆盖同一个（区域, 属性）——例如大标题的行距归加分项，
#    大标题的文本内容与字体字号加粗对齐归减分项，互不重叠。
# ② 属性类减分项一律带前置守卫：判定对象缺失时判为未命中，避免一处破坏被重复扣分。
# ③ 标题层级的字体字号类减分项按“整层被改写”判定：该层级存在标题、且其中没有
#    任何一条合规时才命中。个别标题的零星差异不算破坏，交由对应加分项体现。
_PENALTY_DELTAS_ORDER = [
    -5,  # 01 大标题文本被改写，或“青禾片区公共书房建设现状”丢失
    -5,  # 02 各级标题被降为正文样式，导航目录无层级
    -5,  # 03 全文插图不足 15 张
    -3,  # 04 大标题不为宋体三号加粗居中
    -3,  # 05 二级标题整层不为黑体四号
    -3,  # 06 三级及以下标题整层不为宋体小四
    -3,  # 07 “运营团队将每日开闭馆……”正文段被排成黑体小三
    -3,  # 08 出现与图表内容不符的图题、表题
    -3,  # 09 “活动资源配置对照表”文本丢失
    -1,  # 10 摘要标题不为“//////摘要//////”
    -1,  # 11 两处含“。。”的文字丢失
    -1,  # 12 页边距被改动
    -1,  # 13 两张指定表格的行列数被改动
    -1,  # 14 25 张表格被全部改为居中对齐
]

# ---------- 判定口径常量 ----------

# 字体名判定口径（办公软件里同一字体的不同写法都算命中）
STRICT_SONGTI = ('宋体', 'SimSun', 'NSimSun')
HEITI_KEYS = ('黑体', 'SimHei')

# 字号容差（中文字号对应磅值：三号16、小三15、四号14、小四12、五号10.5）
SZ_SANHAO = (15.5, 16.5)
SZ_XIAOSAN = (14.5, 15.5)
SZ_SIHAO = (13.5, 14.5)
SZ_XIAOSI = (11.5, 12.5)

TWIP_TO_CM = 2.54 / 1440.0

# 行距/段间距容差：1.25 倍 = w:line 300（auto），单倍 = 240（auto）或未设置；
# 段前 0.5 行 = w:beforeLines 50，或 w:before 约 6 磅（120 缇）。
LINE_SINGLE = (234, 246)
LINE_125 = (294, 306)
BEFORE_LINES_HALF = (45, 55)
BEFORE_TWIP_HALF = (110, 130)

# ---------- 文本锚点（全部取自被评估文档应当保留的原文） ----------

DOC_TITLE = '青禾公共书房网络化运营与更新规划书'
ABSTRACT_TITLE = '//////摘要//////'
KEEP_STATUS_TEXT = '青禾片区公共书房建设现状'
KEEP_ACTIVITY_TABLE = '活动资源配置对照表'
# 七个一级标题的标题名（不含层次编号）；层次编号写法本身由第 03 项单独判定，
# 故这里只用标题名定位，保证编号被改写时仍能找到这七条标题。
L1_NAMES = (
    '项目定位与服务原则', '服务体系与空间更新', '数据采集原则',
    '运营模型与效益测算', '各方面影响分析', '分组检验与落地计划',
    '结语与行动清单',
)
# 不得加粗的两处正文段落（按句首片段定位）
NO_BOLD_TEXTS = (
    '数据采集坚持必要、适度、可解释的原则',
    '志愿者参与以轻任务为主，包括书架整理、活动引导、亲子阅读陪伴、旧书交换登记和社区故事收集',
)
# 不得被排成黑体小三的正文段落
NO_HEITI_TEXT = '运营团队将每日开闭馆、借阅登记、场地巡检、活动签到、物资补充和意见反馈纳入同一张工作清单'
# 与所对应图表内容不符的图题、表题：出现任意一条即命中
WRONG_CAPTION_TEXTS = (
    '空间升级——完善服务偏好分析图', '第四章 运营模型与效益测算图',
    '服务点位执行要点表', '指标统计结果表',
)
# 两处含“。。”的文字：缺任意一处即命中
DOUBLE_DOT_TEXTS = (
    '形成十五分钟文化生活圈。。为提升服务的真实可达性',
    '活动排期。。活动设计',
)
# 两张指定表格：表头关键字 → (应有行数, 应有列数)
FIXED_TABLES = (
    (('序号', '服务模块', '时段', '负责人', '资源配置', '验收口径'), 3, 6),
    (('变量', '组别', '均值差', '标准误', 'P值', '效应量', '说明'), 5, 7),
)

# ---------- 正则 ----------

L1_NUM_RE = re.compile(r'^[一二三四五六七八九十]+、')
L2_NUM_RE = re.compile(r'^[（(][一二三四五六七八九十]+[)）]')
L3_NUM_RE = re.compile(r'^\d+\.')
L4_NUM_RE = re.compile(r'^[（(]\d+[)）]')
L5_NUM_RE = re.compile(r'^\d+[)）]')
CHAPTER_RE = re.compile(r'^第[一二三四五六七八九十]+[章节]')
# 编号后紧跟顿号或多余圆点，例如“1）．”“（1）.”“（一）、”“三、、”
STRAY_MARK_RES = (
    re.compile(r'^\d+[)）]\s*[．.、]'),
    re.compile(r'^[（(]\d+[)）]\s*[．.、]'),
    re.compile(r'^[（(][一二三四五六七八九十]+[)）]\s*[、．.]'),
    re.compile(r'^[一二三四五六七八九十]+、\s*[、．.]'),
    re.compile(r'^\d+\.\s*[、．]'),
)
FIG_CAP_RE = re.compile(r'^图\s*(\d+)')
TAB_CAP_RE = re.compile(r'^表\s*(\d+)')
# 只有编号、没有图题表题文字的图注/表注（例如“图23”“表 7”）
BARE_CAP_RE = re.compile(r'^[图表]\s*\d+\s*$')
# 合规写法：“图”“表”后一个空格接编号，编号与图题/表题之间空两格
CAP_TWO_SPACE_RE = re.compile(r'^[图表] \d+ {2}\S')
CJK_RE = re.compile(r'[一-鿿]')
WS_RE = re.compile(r'[\s　]')


def _open_word_document(file_path):
    """打开 .docx 文档（python-docx 仅支持 OOXML/.docx）。"""
    return Document(file_path)


def _norm(text):
    """去掉全部空白（含全角空格）后的文字，用于标题名/表头名的稳定匹配。"""
    return WS_RE.sub('', text or '')


def _cjk_count(text):
    """中文字符个数；维度一用它判断内容是否为可编辑文字。"""
    return len(CJK_RE.findall(text or ''))


def _iter_body_blocks(doc):
    """按文档流顺序遍历正文，返回 [('p', Paragraph) | ('t', Table), ...]。
       图注在图下、表注在表上这类“位置”判定必须依赖真实文档流顺序，
       故不能用 doc.paragraphs / doc.tables 两个独立列表代替。"""
    out = []
    body = doc.element.body
    for child in body.iterchildren():
        if child.tag == qn('w:p'):
            out.append(('p', Paragraph(child, doc)))
        elif child.tag == qn('w:tbl'):
            out.append(('t', Table(child, doc)))
    return out


def _body_paragraphs(blocks):
    return [blk for kind, blk in blocks if kind == 'p']


def _body_tables(blocks):
    return [blk for kind, blk in blocks if kind == 't']


def _style_id(para):
    """段落直接引用的样式 id（w:pPr/w:pStyle）；未引用返回 None。"""
    pPr = para._p.find(qn('w:pPr'))
    if pPr is None:
        return None
    pStyle = pPr.find(qn('w:pStyle'))
    return pStyle.get(qn('w:val')) if pStyle is not None else None


def _style_chain(doc, para):
    """段落样式链（w:pStyle → styles.xml，沿 basedOn 上溯）的 w:style 元素列表。"""
    styles_el = doc.styles.element
    by_id = {}
    for st in styles_el.findall(qn('w:style')):
        sid = st.get(qn('w:styleId'))
        if sid:
            by_id[sid] = st
    out = []
    sid = _style_id(para)
    seen = set()
    while sid and sid not in seen:
        seen.add(sid)
        st = by_id.get(sid)
        if st is None:
            break
        out.append(st)
        base = st.find(qn('w:basedOn'))
        sid = base.get(qn('w:val')) if base is not None else None
    return out


def _doc_default_rpr(doc):
    """docDefaults 里的 w:rPr（字体继承链的最后一层）。"""
    dd = doc.styles.element.find(qn('w:docDefaults'))
    if dd is None:
        return None
    rpd = dd.find(qn('w:rPrDefault'))
    return rpd.find(qn('w:rPr')) if rpd is not None else None


def _resolve_run_font(doc, para, run):
    """按 Word 的继承链解析 run 中文字符实际生效的东亚字体与字号（打开办公软件
       看到的效果）：run 直接设定 → 段落样式（含 basedOn 链）→ docDefaults。
       返回 (东亚字体名或 None, 字号磅值或 None)。"""
    ea = None
    sz = None
    rPr = run._element.rPr
    if rPr is not None:
        rFonts = rPr.find(qn('w:rFonts'))
        if rFonts is not None:
            ea = rFonts.get(qn('w:eastAsia'))
        szEl = rPr.find(qn('w:sz'))
        if szEl is not None and szEl.get(qn('w:val')):
            sz = float(szEl.get(qn('w:val'))) / 2.0
    layers = [st.find(qn('w:rPr')) for st in _style_chain(doc, para)]
    layers.append(_doc_default_rpr(doc))
    for rpr in layers:
        if rpr is None:
            continue
        if ea is None:
            rf = rpr.find(qn('w:rFonts'))
            if rf is not None:
                ea = rf.get(qn('w:eastAsia'))
        if sz is None:
            szEl = rpr.find(qn('w:sz'))
            if szEl is not None and szEl.get(qn('w:val')):
                sz = float(szEl.get(qn('w:val'))) / 2.0
    return ea, sz


def _resolve_run_bold(doc, para, run):
    """按继承链解析 run 是否加粗：run 直接设定 → 段落样式链 → docDefaults。
       w:b 无 w:val 视为 True，w:val 为 0/false 视为 False。"""
    def _read(rpr):
        if rpr is None:
            return None
        b = rpr.find(qn('w:b'))
        if b is None:
            return None
        v = b.get(qn('w:val'))
        if v is None:
            return True
        return v not in ('0', 'false', 'off')

    v = _read(run._element.rPr)
    if v is not None:
        return v
    for st in _style_chain(doc, para):
        v = _read(st.find(qn('w:rPr')))
        if v is not None:
            return v
    v = _read(_doc_default_rpr(doc))
    return bool(v)


def _text_runs(para):
    """承载可见文字的 run（纯空白 run 不参与字体判定）。"""
    return [r for r in para.runs if r.text and r.text.strip()]


def _is_songti(name):
    return bool(name) and name in STRICT_SONGTI


def _is_heiti(name):
    return bool(name) and any(k in name for k in HEITI_KEYS)


def _sz_in(sz, rng):
    return sz is not None and rng[0] <= sz <= rng[1]


def _runs_font_ok(doc, paras, font_pred, sz_rng):
    """段落集合内所有承载文字的 run 是否都满足“字体谓词 + 字号区间”。
       段落集合为空时返回 False（判定对象缺失，加分项不给分）。"""
    hit_any = False
    for para in paras:
        for run in _text_runs(para):
            hit_any = True
            ea, sz = _resolve_run_font(doc, para, run)
            if not font_pred(ea) or not _sz_in(sz, sz_rng):
                return False
    return hit_any


def _para_font_conform(doc, para, font_pred, sz_rng):
    """单个段落是否“所有承载文字的 run”都满足字体与字号；无文字返回 False。"""
    return _runs_font_ok(doc, [para], font_pred, sz_rng)


def _resolve_pPr_child(doc, para, tag):
    """按继承链取段落级 pPr 子元素：段落直接设定 → 段落样式链。
       用于 w:jc（对齐）与 w:spacing（行距/段间距），与办公软件显示口径一致。"""
    pPr = para._p.find(qn('w:pPr'))
    if pPr is not None:
        el = pPr.find(qn(tag))
        if el is not None:
            return el
    for st in _style_chain(doc, para):
        spPr = st.find(qn('w:pPr'))
        if spPr is not None:
            el = spPr.find(qn(tag))
            if el is not None:
                return el
    return None


def _p_jc(doc, para):
    """段落对齐（继承后的 w:jc/@w:val）；未设置返回 None。"""
    jc = _resolve_pPr_child(doc, para, 'w:jc')
    return None if jc is None else jc.get(qn('w:val'))


def _p_spacing(doc, para):
    """段落行距与段间距的原始属性（继承后）：
       {'line': 缇, 'lineRule': auto/exact/atLeast, 'before': 缇, 'beforeLines': 百分之一行}
       注意：python-docx 的 space_before 只读 w:before、line_spacing 只读 w:line，
       都会漏掉 w:beforeLines 这类“按行数”存储的值，故这里直接读原始属性。"""
    sp = _resolve_pPr_child(doc, para, 'w:spacing')
    out = {'line': None, 'lineRule': None, 'before': None, 'beforeLines': None}
    if sp is None:
        return out
    for key, attr in (('line', 'w:line'), ('lineRule', 'w:lineRule'),
                      ('before', 'w:before'), ('beforeLines', 'w:beforeLines')):
        v = sp.get(qn(attr))
        if v is None:
            continue
        if key == 'lineRule':
            out[key] = v
        else:
            try:
                out[key] = int(v)
            except ValueError:
                out[key] = None
    return out


def _is_single_line(sp):
    """单倍行距：未设置 w:line，或 w:line 约 240 且 lineRule 为 auto。"""
    if sp['line'] is None:
        return sp['lineRule'] in (None, 'auto')
    if sp['lineRule'] not in (None, 'auto'):
        return False
    return LINE_SINGLE[0] <= sp['line'] <= LINE_SINGLE[1]


def _is_125_line(sp):
    """1.25 倍行距：w:line 约 300 且 lineRule 为 auto。"""
    if sp['line'] is None or sp['lineRule'] not in (None, 'auto'):
        return False
    return LINE_125[0] <= sp['line'] <= LINE_125[1]


def _is_half_line_before(sp):
    """段前 0.5 行：w:beforeLines 约 50，或 w:before 约 120 缇（约 6 磅）。"""
    if sp['beforeLines'] is not None and \
            BEFORE_LINES_HALF[0] <= sp['beforeLines'] <= BEFORE_LINES_HALF[1]:
        return True
    if sp['before'] is not None and \
            BEFORE_TWIP_HALF[0] <= sp['before'] <= BEFORE_TWIP_HALF[1]:
        return True
    return False


def _has_graphic(para):
    """段落内是否承载图形（插图、图表、形状、文本框都算）。"""
    xml = para._p.xml
    return ('<w:drawing' in xml) or ('<w:pict' in xml) or ('<w:object' in xml)


def _blip_count(para):
    """段落内插图张数：按图片数据引用 a:blip 计数。
       w:drawing 会把形状、文本框一并算进来，故不能用它数插图。"""
    return para._p.xml.count('<a:blip')


def _is_heading_like(text):
    """形似标题的短段落：去空白后长度不超过 40 且不以句号结尾。"""
    t = _norm(text)
    return bool(t) and len(t) <= 40 and not t.endswith('。')


def _has_heading_style(doc, para):
    """段落是否套用了“标题 N / Heading N”这类内置标题样式。
       套用标题样式的段落会出现在办公软件的导航目录里，是层级是否保留的依据。"""
    sid = _style_id(para) or ''
    if sid.lower().startswith('heading'):
        return True
    for st in _style_chain(doc, para):
        nm = st.find(qn('w:name'))
        val = nm.get(qn('w:val')) if nm is not None else ''
        if val and (val.lower().startswith('heading') or val.startswith('标题')):
            return True
    return False


def _tbl_jc(table):
    """表格整体对齐（w:tblPr/w:jc）；未设置返回 None。"""
    pr = table._tbl.find(qn('w:tblPr'))
    if pr is None:
        return None
    jc = pr.find(qn('w:jc'))
    return None if jc is None else jc.get(qn('w:val'))


def _cell_paragraphs(table):
    """表格内全部单元格段落。"""
    out = []
    for row in table.rows:
        for cell in row.cells:
            out.extend(cell.paragraphs)
    return out


def _table_size(table):
    """表格的行数与列数。"""
    try:
        return len(table.rows), len(table.columns)
    except Exception:
        return 0, 0


def _header_texts(table):
    """表格首行各单元格去空白后的文字，用于按表头关键字定位指定表格。"""
    if not table.rows:
        return []
    return [_norm(c.text) for c in table.rows[0].cells]


def _layout(doc):
    """一次性定位全文结构，后续所有判定都基于它，保证同一元素只被判定一次。

    返回字段：
      blocks      文档流顺序的 [('p'|'t', 块)] 列表
      paras       正文层段落（不含表格内段落）
      tables      正文层表格
      title       文档大标题（首个非空段落）
      abstract    摘要标题段落
      fig_caps    [(段落下标, 段落)] 图注
      tab_caps    [(段落下标, 段落)] 表注
      heads       [(段落下标, 段落)] 形似标题的段落（已剔除大标题、摘要标题、图注表注）
      l1/l2/l3p   一级 / 二级 / 三级及以下标题段落
      body        正文段落
      all_text    含表格文字的全文（已去空白）
    """
    blocks = _iter_body_blocks(doc)
    paras = _body_paragraphs(blocks)
    tables = _body_tables(blocks)
    texts = [p.text or '' for p in paras]

    title_i, abstract_i = None, None
    for i, t in enumerate(texts):
        if t.strip():
            title_i = i
            break
    for i, t in enumerate(texts):
        n = _norm(t)
        if '摘要' in n and len(n) <= 20:
            abstract_i = i
            break

    fig_caps, tab_caps = [], []
    for i, t in enumerate(texts):
        s = t.strip()
        if not s or len(s) > 60:
            continue
        if FIG_CAP_RE.match(s):
            fig_caps.append((i, paras[i]))
        elif TAB_CAP_RE.match(s):
            tab_caps.append((i, paras[i]))
    cap_idx = set(i for i, _ in fig_caps) | set(i for i, _ in tab_caps)
    special = set(cap_idx)
    for i in (title_i, abstract_i):
        if i is not None:
            special.add(i)
    heads = [(i, paras[i]) for i, t in enumerate(texts)
             if i not in special and _is_heading_like(t)]

    l1, used = [], set()
    for i, para in heads:
        n = _norm(texts[i])
        for name in L1_NAMES:
            if name not in used and n.endswith(name):
                used.add(name)
                l1.append(para)
                break
    l1_set = set(id(p) for p in l1)
    l2 = [para for i, para in heads
          if L2_NUM_RE.match(texts[i].strip()) and id(para) not in l1_set]
    l3p = [para for i, para in heads
           if id(para) not in l1_set and (L3_NUM_RE.match(texts[i].strip())
                                          or L4_NUM_RE.match(texts[i].strip())
                                          or L5_NUM_RE.match(texts[i].strip()))]

    body = []
    for i, t in enumerate(texts):
        if i in special or not t.strip():
            continue
        if _has_graphic(paras[i]) or _is_heading_like(t):
            continue
        if _has_heading_style(doc, paras[i]):
            continue
        body.append(paras[i])

    cell_text = []
    for tb in tables:
        for row in tb.rows:
            for cell in row.cells:
                cell_text.append(cell.text or '')
    all_text = _norm('\n'.join(texts + cell_text))

    return {
        'blocks': blocks, 'paras': paras, 'tables': tables, 'texts': texts,
        'title': None if title_i is None else paras[title_i],
        'abstract': None if abstract_i is None else paras[abstract_i],
        'fig_caps': fig_caps, 'tab_caps': tab_caps, 'heads': heads,
        'l1': l1, 'l2': l2, 'l3p': l3p, 'body': body, 'all_text': all_text,
    }


def _cap_num(text, regex):
    """图注/表注里的阿拉伯编号；取不到返回 None。"""
    m = regex.match((text or '').strip())
    return None if m is None else int(m.group(1))


def _nums_from_one(nums):
    """按出现顺序，编号是否从 1 开始连续、不重号不跳号。"""
    return bool(nums) and nums == list(range(1, len(nums) + 1))


def _fig_pos_ok(lay, i):
    """图注是否位于插图下方：图注段本身或其上方 5 段内承载图形
       （图形以浮动方式锚定在图注段时，显示效果同样是图在上、注在下）。"""
    paras = lay['paras']
    for j in range(max(0, i - 5), i + 1):
        if _has_graphic(paras[j]):
            return True
    return False


def _pictures_with_caption(lay):
    """返回 (插图张数, 未配图注的插图段落数)。插图段落自身或其后 5 段内
       出现图注，即视为该插图配有图注。"""
    paras = lay['paras']
    cap_idx = set(i for i, _ in lay['fig_caps'])
    total, missing = 0, 0
    for i, para in enumerate(paras):
        cnt = _blip_count(para)
        if cnt <= 0:
            continue
        total += cnt
        if not any(j in cap_idx for j in range(i, min(len(paras), i + 6))):
            missing += 1
    return total, missing


def _tables_with_caption(lay):
    """返回 (表格数, 上方紧邻非空段落不是表注的表格数)。"""
    blocks = lay['blocks']
    cap_ids = set(id(p) for _, p in lay['tab_caps'])
    total, missing = 0, 0
    for idx, (kind, _blk) in enumerate(blocks):
        if kind != 't':
            continue
        total += 1
        cap = None
        for j in range(idx - 1, -1, -1):
            if blocks[j][0] != 'p':
                break
            if not (blocks[j][1].text or '').strip():
                continue
            cap = blocks[j][1]
            break
        if cap is None or id(cap) not in cap_ids:
            missing += 1
    return total, missing


def _margins_ok(doc):
    """页边距是否为上下 2.54cm、左右 3.18cm（容差 0.05cm）。"""
    secs = list(doc.sections)
    if not secs:
        return False
    for sec in secs:
        vals = (sec.top_margin, sec.bottom_margin, sec.left_margin, sec.right_margin)
        if any(v is None for v in vals):
            return False
        top, bottom, left, right = (v.cm for v in vals)
        if abs(top - 2.54) > 0.05 or abs(bottom - 2.54) > 0.05:
            return False
        if abs(left - 3.18) > 0.05 or abs(right - 3.18) > 0.05:
            return False
    return True


def _find_paras(lay, needle):
    """按文字片段定位正文层段落（去空白后包含即命中），返回全部命中段落。
       被评估文档里同一句式在多处复现，只取首个命中会定位错对象，
       故所有按文字定位的判定都对“全部命中段落”一起判。"""
    key = _norm(needle)
    if not key:
        return []
    return [para for para, text in zip(lay['paras'], lay['texts'])
            if key in _norm(text)]



def _dim1_check(file_path):
    """维度一只看交付文件格式：是否 .docx、能否正常打开、内容是否为可编辑文字。
       返回 (是否通过, 不通过原因, 已打开的文档或 None)。"""
    if not file_path.lower().endswith('.docx'):
        return False, "交付文件不是 .docx 格式：%s" % os.path.basename(file_path), None
    try:
        doc = _open_word_document(file_path)
    except Exception as exc:
        return False, "文件损坏或无法打开：%s: %s" % (type(exc).__name__, exc), None
    try:
        blocks = _iter_body_blocks(doc)
        cjk = sum(_cjk_count(p.text) for p in _body_paragraphs(blocks))
        for tb in _body_tables(blocks):
            for row in tb.rows:
                for cell in row.cells:
                    cjk += _cjk_count(cell.text)
    except Exception as exc:
        return False, "文档结构无法解析：%s: %s" % (type(exc).__name__, exc), None
    if cjk < 200:
        return False, "文档内容不是可编辑文字（可识别中文字符仅 %d 个）" % cjk, None
    return True, "", doc


def _build_dim2_items(records, deltas_order, kind):
    """按“每条规则一个 item”生成严格二值的评分明细，返回 (items, 小计)。

    records 为 _emit / _penalize 追加的 (是否命中, 满分值, 细则原文) 三元组，
    与 deltas_order 逐项对位；两者项数或分值不一致时直接抛错，
    以免顺序表与判定点悄悄漂移后仍算出一个看似正常的分数。
    每项只输出 rule / max_delta / delta / hit 四个键，判定说明由 rule 原文承载。
    """
    if len(records) != len(deltas_order):
        raise RuntimeError('dim2 %s项数不一致：records=%d vs order=%d'
                           % (kind, len(records), len(deltas_order)))
    items = []
    subtotal = 0
    for (hit, delta, rule), max_delta in zip(records, deltas_order):
        if delta != max_delta:
            raise RuntimeError('dim2 %s分值与顺序表不一致：“%s”记为 %d，顺序表为 %d'
                               % (kind, rule, delta, max_delta))
        items.append({"rule": rule, "max_delta": max_delta,
                      "delta": max_delta if hit else 0, "hit": hit})
        subtotal += max_delta if hit else 0
    return items, subtotal


def evaluate(dir_path: str) -> dict:
    """统一入口：接收输出文件所在目录，定位并评估目录内的规划书文档。

    评分口径：维度一为门槛项，未通过则 dim2_items 置空、total_score 记 0，
    max_score 仍为加分项满分之和（MAX_SCORE_TOTAL）；维度二逐条二值判定，
    total_score 为加分项得分减去减分项扣分，扣至 0 为止。

    status='error' 只在“这份提交没被真正评过”时出现——目录不存在、目录内没有
    可评估的 Word 文档、或脚本自身抛异常；此时 dim1_pass 为空串（未判定）、
    dim1_reason 为空串、dim2_items 为空表、total_score 与 max_score 一律为 0。
    文档能定位到但打不开或不是可编辑文字，仍按 status='ok' + 维度一未通过处理，
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

        cands = sorted(f for f in os.listdir(dir_path)
                       if f.lower().endswith('.docx') and not f.startswith('~$'))
        if not cands:
            result["status"] = "error"
            result["error"] = f"目录中未找到 .docx 文件：{dir_path}"
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


def _evaluate_impl(file_path, result):
    results = []
    penalties = []

    def _emit(ok, delta, rule):
        """追加一条加分项判定：(是否命中, 满分值, 细则原文)。"""
        results.append((bool(ok), delta, rule))

    def _penalize(hit, delta, rule):
        """追加一条减分项判定：(是否命中即确有破坏性改动, 扣分值, 细则原文)。"""
        penalties.append((bool(hit), delta, rule))

    ok1, reason1, doc = _dim1_check(file_path)
    result["dim1_pass"] = bool(ok1)
    # 约定：dim1_pass 为 False 时 dim1_reason 必须带错误摘要，通过时必须为空串。
    if ok1:
        result["dim1_reason"] = ""
    else:
        result["dim1_reason"] = reason1 or "维度一未通过（未给出具体原因）"
    if not ok1 or doc is None:
        result["dim1_pass"] = False
        result["dim1_reason"] = result["dim1_reason"] or "维度一未通过（未给出具体原因）"
        result["dim2_items"] = []
        result["total_score"] = 0
        result["max_score"] = MAX_SCORE_TOTAL
        return result

    lay = _layout(doc)
    title = lay['title']
    abstract = lay['abstract']
    body = lay['body']
    head_texts = [lay['texts'][i].strip() for i, _ in lay['heads']]
    # 被单独设为标题效果的那段正文由减分项 07 判定，其字体字号不再进入
    # 加分项 05，避免同一处破坏被既扣分又扣分。
    nh_paras = _find_paras(lay, NO_HEITI_TEXT)
    nh_ids = set(id(p) for p in nh_paras)

    # 01 大标题行距
    _emit(title is not None and _is_single_line(_p_spacing(doc, title)),
          1, "文档大标题行距为单倍行距")

    # 02 摘要标题字体字号
    _emit(abstract is not None and _para_font_conform(doc, abstract, _is_heiti, SZ_SIHAO),
          1, "摘要标题为黑体、四号（14磅）")

    # 03 五级层次编号依次嵌套；不得出现“第X章”与编号后的顿号、多余圆点
    lvl_cnt = [sum(1 for t in head_texts if rex.match(t))
               for rex in (L1_NUM_RE, L2_NUM_RE, L3_NUM_RE, L4_NUM_RE, L5_NUM_RE)]
    chapter_hit = any(CHAPTER_RE.match(t) for t in head_texts)
    stray_hit = any(rex.match(t) for t in head_texts for rex in STRAY_MARK_RES)
    _emit(all(c > 0 for c in lvl_cnt) and not chapter_hit and not stray_hit, 5,
          "全文标题编号为“一、”“（一）”“1.”“（1）”“1）”五级依次嵌套，"
          "编号后无顿号和多余圆点，无“第一章”“1）．”这类写法")

    # 04 七个一级标题字体字号
    _emit(len(lay['l1']) == len(L1_NAMES)
          and _runs_font_ok(doc, lay['l1'], _is_heiti, SZ_XIAOSAN),
          5, "七个一级标题均为黑体、小三（15磅）")

    # 05 正文字体字号
    _emit(_runs_font_ok(doc, [p for p in body if id(p) not in nh_ids],
                        _is_songti, SZ_XIAOSI),
          5, "正文段落为宋体、小四（12磅），无五号（10.5磅）字号")

    # 06 正文行距
    _emit(bool(body) and all(_is_125_line(_p_spacing(doc, p)) for p in body),
          3, "正文段落行距为1.25倍")

    # 07 正文段前间距
    _emit(bool(body) and all(_is_half_line_before(_p_spacing(doc, p)) for p in body),
          3, "正文段落段前间距为0.5行（约6磅）")

    # 08 两处指定正文段落不加粗
    nb_groups = [_find_paras(lay, t) for t in NO_BOLD_TEXTS]
    nb_ok = all(bool(g) for g in nb_groups) and not any(
        _resolve_run_bold(doc, p, r) for g in nb_groups for p in g
        for r in _text_runs(p))
    _emit(nb_ok, 1, "正文段落“数据采集坚持必要、适度、可解释的原则”与“志愿者参与以轻任务为主，"
                    "包括书架整理、活动引导、亲子阅读陪伴、旧书交换登记和社区故事收集”不加粗")

    figs = lay['fig_caps']
    tabs_cap = lay['tab_caps']

    # 09 图注字体字号、居中、位于插图下方
    _emit(bool(figs) and all(_para_font_conform(doc, p, _is_songti, SZ_XIAOSI)
                             and _p_jc(doc, p) == 'center'
                             and _fig_pos_ok(lay, i) for i, p in figs),
          3, "图注为宋体、小四（12磅）、居中，位于对应插图下方")

    # 10 带图名的图注：编号与图名之间空两格
    named_figs = [p for _i, p in figs if not BARE_CAP_RE.match((p.text or '').strip())]
    _emit(bool(named_figs) and all(CAP_TWO_SPACE_RE.match((p.text or '').strip())
                                   for p in named_figs),
          3, "编号后带图名的图注，编号与图名之间空两格，写法为“图 1  公共书房功能分区关系图”")

    # 11 插图均配图注、图注编号连续
    pic_total, pic_missing = _pictures_with_caption(lay)
    fig_nums = [_cap_num(p.text, FIG_CAP_RE) for _i, p in figs]
    _emit(pic_total >= 15 and pic_missing == 0 and _nums_from_one(fig_nums),
          5, "15张插图均配有图注，图注编号从1开始连续、不重号不跳号")

    # 12 表注字体字号、居中、位于表格上方
    tab_total, tab_missing = _tables_with_caption(lay)
    _emit(bool(tabs_cap) and tab_missing == 0
          and all(_para_font_conform(doc, p, _is_songti, SZ_XIAOSI)
                  and _p_jc(doc, p) == 'center' for _i, p in tabs_cap),
          3, "表注为宋体、小四（12磅）、居中，位于对应表格上方")

    # 13 带表名的表注：编号与表名之间空两格
    named_tabs = [p for _i, p in tabs_cap if not BARE_CAP_RE.match((p.text or '').strip())]
    _emit(bool(named_tabs) and all(CAP_TWO_SPACE_RE.match((p.text or '').strip())
                                   for p in named_tabs),
          3, "编号后带表名的表注，编号与表名之间空两格，写法为“表 1  公共书房运营指标表”")

    # 14 表格均配表注、表注编号连续
    tab_nums = [_cap_num(p.text, TAB_CAP_RE) for _i, p in tabs_cap]
    _emit(tab_total >= 25 and tab_missing == 0 and _nums_from_one(tab_nums),
          5, "25张表格均配有表注，表注编号从1开始连续、不重号不跳号")

    # 15 表格内文字字体字号
    cell_paras = [p for tb in lay['tables'] for p in _cell_paragraphs(tb)]
    _emit(_runs_font_ok(doc, cell_paras, _is_songti, SZ_XIAOSI),
          3, "表格内文字为宋体、小四（12磅）")

    # 16 表格内文字行距
    cell_text_paras = [p for p in cell_paras if (p.text or '').strip()]
    _emit(bool(cell_text_paras)
          and all(_is_single_line(_p_spacing(doc, p)) for p in cell_text_paras),
          3, "表格内文字为单倍行距")

    all_text = lay['all_text']

    # 减分 01 大标题文本被改写，或关键小节名丢失
    _penalize(title is None or _norm(title.text) != _norm(DOC_TITLE)
              or _norm(KEEP_STATUS_TEXT) not in all_text,
              -5, "文档大标题不为“青禾公共书房网络化运营与更新规划书”，"
                  "或全文未出现“青禾片区公共书房建设现状”")

    # 减分 02 标题层级被抹平：形似标题的段落一条都没套用标题样式
    _penalize(bool(lay['heads'])
              and not any(_has_heading_style(doc, p) for _i, p in lay['heads']),
              -5, "各级标题均为正文样式，导航目录中不显示标题层级")

    # 减分 03 插图数量减少
    _penalize(pic_total < 15, -5, "全文插图不足15张")

    # 减分 04 大标题字体字号加粗对齐被改动
    title_fmt_ok = (title is not None
                    and _para_font_conform(doc, title, _is_songti, SZ_SANHAO)
                    and bool(_text_runs(title))
                    and all(_resolve_run_bold(doc, title, r) for r in _text_runs(title))
                    and _p_jc(doc, title) == 'center')
    _penalize(title is not None and not title_fmt_ok,
              -3, "文档大标题不为宋体、三号（16磅）、加粗、居中")

    # 减分 05 二级标题整层被改写（该层存在标题、却没有一条合规）
    _penalize(bool(lay['l2'])
              and not any(_para_font_conform(doc, p, _is_heiti, SZ_SIHAO) for p in lay['l2']),
              -3, "二级标题不为黑体、四号（14磅）")

    # 减分 06 三级及以下标题整层被改写
    _penalize(bool(lay['l3p'])
              and not any(_para_font_conform(doc, p, _is_songti, SZ_XIAOSI) for p in lay['l3p']),
              -3, "三级及以下标题不为宋体、小四（12磅）")

    # 减分 07 指定正文段落被排成标题效果
    _penalize(any(_para_font_conform(doc, p, _is_heiti, SZ_XIAOSAN) for p in nh_paras),
              -3, "正文段落“运营团队将每日开闭馆、借阅登记、场地巡检、活动签到、物资补充"
                  "和意见反馈纳入同一张工作清单”为黑体、小三（15磅）")

    # 减分 08 编造与图表内容不符的图题、表题
    _penalize(any(_norm(t) in all_text for t in WRONG_CAPTION_TEXTS),
              -3, "出现“空间升级——完善服务偏好分析图”“第四章 运营模型与效益测算图”"
                  "“服务点位执行要点表”“指标统计结果表”等与所对应图表内容不符的图题、表题")

    # 减分 09 指定表名文本丢失
    _penalize(_norm(KEEP_ACTIVITY_TABLE) not in all_text,
              -3, "没有出现“活动资源配置对照表”文本")

    # 减分 10 摘要标题文本被改写
    _penalize(abstract is None or _norm(abstract.text) != _norm(ABSTRACT_TITLE),
              -1, "摘要标题不为“//////摘要//////”")

    # 减分 11 两处含“。。”的文字被清理
    _penalize(any(_norm(t) not in all_text for t in DOUBLE_DOT_TEXTS),
              -1, "没有出现“形成十五分钟文化生活圈。。为提升服务的真实可达性”与"
                  "“活动排期。。活动设计”两处文字")

    # 减分 12 页边距被改动
    _penalize(not _margins_ok(doc), -1, "页边距不为上下2.54cm、左右3.18cm")

    # 减分 13 两张指定表格的行列数被改动
    size_bad = False
    for keys, rows, cols in FIXED_TABLES:
        found = None
        for tb in lay['tables']:
            header = _header_texts(tb)
            if all(any(k in cell for cell in header) for k in keys):
                found = tb
                break
        if found is None or _table_size(found) != (rows, cols):
            size_bad = True
            break
    _penalize(size_bad, -1,
              "“序号/服务模块/时段/负责人/资源配置/验收口径”表不为3行6列，"
              "或“变量/组别/均值差/标准误/P值/效应量/说明”表不为5行7列")

    # 减分 14 表格对齐被整体改成居中
    _penalize(bool(lay['tables'])
              and all(_tbl_jc(tb) == 'center' for tb in lay['tables']),
              -1, "25张表格全部为居中对齐")

    # ========== 生成结构化 dim2_items ==========
    # 加分项在前、减分项在后，与 _MAX_DELTAS_ORDER / _PENALTY_DELTAS_ORDER 同序。
    pos_items, gained = _build_dim2_items(results, _MAX_DELTAS_ORDER, '加分')
    neg_items, lost = _build_dim2_items(penalties, _PENALTY_DELTAS_ORDER, '减分')

    result["dim2_items"] = pos_items + neg_items
    result["total_score"] = max(0, gained + lost)
    result["max_score"] = MAX_SCORE_TOTAL
    return result


if __name__ == "__main__":
    # 本地调试入口：命令行第 1 个参数可覆盖目录路径；未指定时使用脚本所在目录。
    # 直接写 stdout 的二进制缓冲区并显式指定 utf-8，避免受控制台默认编码影响。
    _dir = sys.argv[1] if len(sys.argv) > 1 else os.path.dirname(os.path.abspath(__file__))
    _output = json.dumps(evaluate(_dir), ensure_ascii=False, indent=2) + "\n"
    sys.stdout.buffer.write(_output.encode("utf-8", errors="replace"))
