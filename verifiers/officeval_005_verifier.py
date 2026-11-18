#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# pyright: reportUnknownMemberType=false, reportUnknownVariableType=false, reportUnknownArgumentType=false, reportUnknownParameterType=false, reportMissingParameterType=false, reportPrivateUsage=false, reportUnknownLambdaType=false, reportGeneralTypeIssues=false

import os
import re
import sys
import json
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn

# 说明：按接口统一约定，本模块不修改全局 sys.stdout，也不 print 主结果。
# 主结果统一由 evaluate(dir_path) -> dict 返回；入参为“脚本所在目录的路径”，
# 由脚本自己在该目录内定位并打开被评估的 .docx。本文件仅在 __main__ 下作为
# 本地调试入口，才把 dict 序列化为 JSON 打印。
#
# 与 008 号脚本的差异：本任务（成果文本编制要求 / 打印要求）细则同时含
# “加分项”（指令要求的报告排版任务）与“减分项”（指令未涉及却破坏原文件的多余
# 修改）。故 dim2 项的 max_delta 允许为负：正值=加分项（命中记满分，未命中记0）；
# 负值=减分项（触发“破坏”记该负分，未触发记0）。total_score 为全部 delta 之和，
# 取值区间 [-25, +42]；max_score 仍为所有加分项满分之和（+42），作为完成度上限。
#
# 已移除原 A06/A07（各节二级/三级题序标题字体字号，各 +1）：本任务输入文档的标题
# 层级只到 X.Y 一级，全文无 X.Y.Z / X.Y.Z.W 形态的题序，两项恒走“无则默认达标”
# 分支而无条件送分，无区分度。故满分由 +44 降为 +42；保留其余项的 A 编号不变，
# 便于与历史跑分对照。

SCRIPT_ID = "005"

# —— 加分项按“结果追加顺序”的 max_delta（正值）——
_ADD_DELTAS_ORDER = [
    1,  # A01 纸张 A4
    3,  # A02 页边距 左30/右25/上30/下25
    3,  # A03 正文行间距固定值22磅
    3,  # A04 各章题序标题 小二号黑体
    3,  # A05 各节一级题序标题 小三号黑体
    3,  # A08 报告正文 小四号宋体
    1,  # A09 目录标题 小二号黑体
    3,  # A10 目录内容 小四号宋体
    3,  # A11 结论/参考文献/附录 标题小二黑体、正文小四宋体
    3,  # A12 页眉文字“云岭数字技术学院本科课程设计报告”
    1,  # A13 页眉宋体五号、居中
    1,  # A14 页眉下方单实线下划线
    3,  # A15 报告封面不加页眉页脚
    3,  # A16 封面外各部分页脚页码、阿拉伯连续
    5,  # A17 页码以正文为第1页（正文单独分节并重设起始页码为1）
    3,  # A18 图/表题注位置与阿拉伯编号（无图表则默认达标）
]

# —— 减分项按“结果追加顺序”的 max_delta（负值）——
_DED_DELTAS_ORDER = [
    -1,  # D01 指导书封面主标题
    -1,  # D02 指导书封面作者/单位/年月三行
    -1,  # D03 指导书“目录”标题或条目
    -1,  # D04 评分标准/备选方向 居中标题
    -1,  # D05 “三、综合设计备选方向”段落
    -1,  # D06 备选方向 5 条条目
    -1,  # D07 “实施安排：第1周…”段落
    -1,  # D08 编制要求 小标题
    -1,  # D09 编制要求 正文段落
    -1,  # D10 报告封面“云岭数字技术学院”
    -1,  # D11 报告封面“《…》/本科设计报告”
    -1,  # D12 报告封面信息行
    -1,  # D13 报告封面“二〇二六年四月”
    -1,  # D14 说明页标题
    -1,  # D15 说明页正文
    -3,  # D16 删除/改写原文文字内容
    -3,  # D17 删除/破坏表格、插图等对象
    -3,  # D18 删除原有段落或章节
    -1,  # D19 改动姓名/学号/日期等原始信息
]

MAX_SCORE_TOTAL = sum(_ADD_DELTAS_ORDER)          # +42：完成度上限
_MIN_TOTAL = MAX_SCORE_TOTAL + sum(_DED_DELTAS_ORDER)  # +17：全部命中且全部踩坑
assert MAX_SCORE_TOTAL == 42
assert sum(_DED_DELTAS_ORDER) == -25

# 结果追加顺序 = 先加分项、后减分项
_MAX_DELTAS_ORDER = _ADD_DELTAS_ORDER + _DED_DELTAS_ORDER
_ADD_COUNT = len(_ADD_DELTAS_ORDER)

# 中文字号 -> 磅值（用于阅读细则）：小二=18 小三=15 四号=14 小四=12 五号=10.5 二号=22 三号=16
STRICT_SONGTI = ('宋体', 'SimSun', 'NSimSun')
STRICT_HEITI = ('黑体', 'SimHei')


def _resolve_effective_eastasia(doc, run):
    """按 Word 的字体继承链，解析 run 中文字符实际生效的东亚字体（打开办公软件看到的）：
       run 直接设定(w:eastAsia) → 段落样式(含 basedOn 链) → docDefaults。都没有则 None。"""
    rPr = run._element.rPr
    if rPr is not None:
        rFonts = rPr.find(qn('w:rFonts'))
        if rFonts is not None:
            ea = rFonts.get(qn('w:eastAsia'))
            if ea:
                return ea
    styles_el = doc.styles.element
    para = run._parent
    style_id = None
    try:
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

    seen = set()
    while style_id and style_id not in seen:
        seen.add(style_id)
        st = _style_by_id(style_id)
        if st is None:
            break
        rpr = st.find(qn('w:rPr'))
        if rpr is not None:
            rf = rpr.find(qn('w:rFonts'))
            if rf is not None:
                ea = rf.get(qn('w:eastAsia'))
                if ea:
                    return ea
        base = st.find(qn('w:basedOn'))
        style_id = base.get(qn('w:val')) if base is not None else None

    dd = styles_el.find(qn('w:docDefaults'))
    if dd is not None:
        rpd = dd.find(qn('w:rPrDefault'))
        if rpd is not None:
            rpr = rpd.find(qn('w:rPr'))
            if rpr is not None:
                rf = rpr.find(qn('w:rFonts'))
                if rf is not None:
                    ea = rf.get(qn('w:eastAsia'))
                    if ea:
                        return ea
    return None


# ============ 段落属性小工具（均以“打开办公软件看到的效果”为准） ============

def _norm(s):
    """归一化文本：去除所有空白（含全角空格/制表符/NBSP）。"""
    return re.sub(r'[\s　\xa0]+', '', s or '')


def _first_run(para):
    for r in para.runs:
        if (r.text or '').strip():
            return r
    return para.runs[0] if para.runs else None


def _ea(doc, para):
    r = _first_run(para)
    return _resolve_effective_eastasia(doc, r) if r is not None else None


def _size_pt(para):
    r = _first_run(para)
    if r is not None and r.font.size is not None:
        return r.font.size.pt
    return None


def _align(para):
    """读取段落对齐，兼容 python-docx 未映射的 justify/start/end 值。"""
    ppr = para._p.find(qn('w:pPr'))
    jc = ppr.find(qn('w:jc')) if ppr is not None else None
    value = jc.get(qn('w:val')) if jc is not None else None
    raw_alignment = {
        'left': 'L', 'start': 'L', 'center': 'C',
        'right': 'R', 'end': 'R', 'both': 'J',
        'justify': 'J', 'distribute': 'J',
    }.get(value)
    if raw_alignment is not None:
        return raw_alignment
    try:
        a = para.paragraph_format.alignment
    except ValueError:
        return 'L'
    if a == WD_ALIGN_PARAGRAPH.CENTER:
        return 'C'
    if a == WD_ALIGN_PARAGRAPH.RIGHT:
        return 'R'
    if a == WD_ALIGN_PARAGRAPH.JUSTIFY:
        return 'J'
    return 'L'  # None / LEFT 均视为左对齐


def _line_spacing(para):
    """返回 (kind, value)：
       kind='exact' 固定值(磅) -> value=磅；kind='mult' 多倍 -> value=倍数；
       kind='single' 单倍/未设 -> value=None。"""
    pf = para.paragraph_format
    rule = pf.line_spacing_rule
    ls = pf.line_spacing
    # EXACTLY / AT_LEAST：ls 为 Length（EMU），换算磅
    from docx.enum.text import WD_LINE_SPACING
    if rule in (WD_LINE_SPACING.EXACTLY, WD_LINE_SPACING.AT_LEAST):
        try:
            return ('exact', ls.pt)
        except Exception:
            return ('exact', None)
    if rule == WD_LINE_SPACING.MULTIPLE and ls is not None:
        return ('mult', float(ls))
    if isinstance(ls, float):
        return ('mult', ls)
    return ('single', None)


def _space_after_pt(para):
    sa = para.paragraph_format.space_after
    return sa.pt if sa is not None else None


def _is_song(name):
    return bool(name) and any(k in name for k in STRICT_SONGTI)


def _is_hei(name):
    return bool(name) and any(k in name for k in STRICT_HEITI)


def _is_noto(name):
    return bool(name) and 'Noto' in name


def _section_of_paragraphs(doc):
    """返回 dict：段落元素 -> 所属节序号（0 起）。以段落 pPr/sectPr 作为分节终止符。"""
    mapping = {}
    cur = 0
    for p in doc.paragraphs:
        mapping[p._element] = cur
        pPr = p._element.find(qn('w:pPr'))
        if pPr is not None and pPr.find(qn('w:sectPr')) is not None:
            cur += 1
    return mapping


def _preserved(doc, para, font_pred, size, align, ls_expect, sa_expect=None, sa_tol=2.0):
    """判断段落是否仍保持“原格式”。任一给定项不符即返回 False（=被改动）。
       font_pred: 东亚字体判定函数（如 _is_noto / _is_song）；None 表示不校验字体
       size: 期望磅值（容差 ±1.0）；None 不校验
       align: 期望对齐 'C'/'L'…；None 不校验
       ls_expect: 'mult108'（约1.08倍）/ 'single'（单倍）/ None 不校验
       sa_expect: 期望段后磅值；None 不校验"""
    if font_pred is not None:
        name = _ea(doc, para)
        if not (name and font_pred(name)):
            return False
    if size is not None:
        s = _size_pt(para)
        if not (s is not None and abs(s - size) <= 1.0):
            return False
    if align is not None and _align(para) != align:
        return False
    if ls_expect is not None:
        k, v = _line_spacing(para)
        if ls_expect == 'mult108':
            if not (k == 'mult' and v is not None and 1.0 < v < 1.2):
                return False
        elif ls_expect == 'single':
            if k != 'single':
                return False
    if sa_expect is not None:
        sap = _space_after_pt(para)
        if not (sap is not None and abs(sap - sa_expect) <= sa_tol):
            return False
    return True


def _find(paras, pred, lo=0, hi=None):
    """返回第一个满足 pred(归一化文本) 的段落（可限定 [lo, hi) 段落区间），无则 None。"""
    hi = len(paras) if hi is None else hi
    for i in range(lo, hi):
        if pred(_norm(paras[i].text)):
            return paras[i]
    return None


def evaluate(dir_path):
    """按接口约定返回结构化 dict；不 print、不改 sys.stdout、不 sys.exit。
       入参 dir_path 为目录路径，脚本在其中定位并打开被评估的 .docx
       （忽略以 ~$ 打头的临时锁文件）。任何异常均捕获为 status='error'。

       定位不到 .docx 等前置失败时，dim1_pass=False、dim1_reason=""、
       dim2_items=[]、total_score=max_score=0，error 仅给出可读原因
       （不带异常类名前缀）。"""
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
        cands = sorted(
            f for f in os.listdir(dir_path)
            if f.lower().endswith('.docx') and not f.startswith('~$')
        )
        if not cands:
            raise FileNotFoundError(f"目录中未找到 .docx 文件：{dir_path}")
        file_path = os.path.join(dir_path, cands[0])
        result["file_name"] = os.path.basename(file_path)
        return _evaluate_impl(file_path, result)
    except FileNotFoundError as exc:
        result["status"] = "error"
        result["error"] = str(exc)
        # 维度一根本没进入判定，dim1_pass 与 dim1_reason 一律留空（空字符串，而非 False）
        result["dim1_pass"] = ""
        result["dim2_items"] = []
        result["total_score"] = 0
        result["max_score"] = 0
        return result
    except Exception as exc:
        result["status"] = "error"
        result["error"] = f"脚本运行异常：{type(exc).__name__}: {exc}"
        result["dim1_pass"] = False
        result["dim2_items"] = []
        result["total_score"] = 0
        result["max_score"] = 0
        return result


# 中文数字章序，用于识别“第X章”标题
_CHAP_RE = re.compile(r'^第[一二三四五六七八九十百]+章')
# 一级小节号 X.Y（非 X.Y.Z）
_L1_RE = re.compile(r'^\d+\.\d+(?!\.\d)')
# 二级小节号 X.Y.Z（本文档无此层级，仅用于从“正文段落”中排除小节标题）
_L2_RE = re.compile(r'^\d+\.\d+\.\d+(?!\.\d)')

_HEADER_TEXT = '云岭数字技术学院本科课程设计报告'

# —— 减分项 D16~D19 的原文基准（用于“无原文可 diff”时的哨兵判定）——
# 均取自输入文件《综合设计指导书.docx》报告部分正文，去空格后应原样保留在输出中。
_CONTENT_SENTINELS = (
    '本设计围绕城市应急物资调度与路径规划问题展开',
    '系统主要处理四类数据对象：物资记录、仓点记录、需求记录和车辆记录',
    '物资台账使用映射表保存物资编码与库存信息',
)
# 原文对象数量基准：4 张表、3 处插图/流程图（drawing）
_ORIG_MIN_TABLES = 4
_ORIG_MIN_DRAWINGS = 3
# 章节/专项标题存在性（缺失即视为删除段落或章节）
_SECTION_PATTERNS = (
    r'第一章', r'第二章', r'第三章', r'第四章',
    r'结\s*论', r'参考(文献|资料)', r'附\s*录',
)
# 原始信息数值（姓名、学号、日期），改动即扣分
_ORIG_IDENTIFIERS = (
    '26301841', '26301858', '26301873', '程奕然', '二〇二六年四月', '林沐川',
)



def _evaluate_impl(file_path, result):
    """维度一 + 维度二评估主体。加分项、减分项分别构造 "+N："/"-N：" 结果串，
       末尾按 _MAX_DELTAS_ORDER 反解成结构化 dim2_items。"""
    doc = Document(file_path)
    paras = doc.paragraphs
    secmap = _section_of_paragraphs(doc)
    results = []  # 先加分项，后减分项

    # ---------- 维度1：可用与可修改性 ----------
    text_paras = [p for p in paras if p.text.strip()]
    if len(text_paras) < 20 or len(doc.element.body.findall(qn('w:p'))) == 0:
        result["dim1_pass"] = False
        result["dim1_reason"] = "文档正文缺失或不可用（可编辑文字过少）"
        result["dim2_items"] = []
        result["total_score"] = 0
        return result
    result["dim1_pass"] = True
    result["dim1_reason"] = "文件可正常打开，正文为可编辑文字/表格，版面未见严重错乱"

    # ---------- 公共定位：报告部分起点、正文/目录/说明所在节 ----------
    def _idx(pred, lo=0, hi=None):
        hi = len(paras) if hi is None else hi
        for i in range(lo, hi):
            if pred(_norm(paras[i].text)):
                return i
        return -1

    # 报告封面唯一锚点：《工程实践与算法应用综合设计》
    report_lo = _idx(lambda t: t == '《工程实践与算法应用综合设计》')
    if report_lo < 0:
        report_lo = _idx(lambda t: t == '本科设计报告')
    if report_lo < 0:
        report_lo = 0

    # 报告正文首个章标题（居中、非目录条目）
    zhengwen_idx = -1
    for i in range(report_lo, len(paras)):
        t = _norm(paras[i].text)
        if _CHAP_RE.match(t) and _align(paras[i]) == 'C':
            zhengwen_idx = i
            break

    # 报告目录标题（报告区内，正文之前的“目录”）
    hi_toc = zhengwen_idx if zhengwen_idx > 0 else len(paras)
    rep_toc_idx = _idx(lambda t: t == '目录', report_lo, hi_toc)
    # 说明页“说明”标题（报告区内、目录之前）
    hi_expl = rep_toc_idx if rep_toc_idx > 0 else hi_toc
    expl_idx = _idx(lambda t: t == '说明', report_lo, hi_expl)

    # 正文所在节对应的 section 对象（用于页面/页边距/页码判定）
    def _section_obj_of(idx):
        if idx < 0 or idx >= len(paras):
            return doc.sections[-1]
        si = secmap.get(paras[idx]._element, len(doc.sections) - 1)
        si = min(si, len(doc.sections) - 1)
        return doc.sections[si]

    zw_section = _section_obj_of(zhengwen_idx if zhengwen_idx > 0 else report_lo)

    def _mm(twips):
        return int(twips) / 20.0 / 72.0 * 25.4 if twips else None

    # ======================= 加分项 =======================

    # A01 +1：报告纸张 A4（宽210mm 高297mm，容差 ±1mm）
    pgSz = zw_section._sectPr.find(qn('w:pgSz'))
    a4_ok = False
    if pgSz is not None:
        w = _mm(pgSz.get(qn('w:w')))
        h = _mm(pgSz.get(qn('w:h')))
        a4_ok = w is not None and h is not None and abs(w - 210) <= 1.5 and abs(h - 297) <= 1.5
    results.append(("+1：报告纸张为A4（210×297mm）") if a4_ok else "+0：报告纸张为A4（210×297mm）")

    # A02 +3：页边距 左30/右25/上30/下25 mm（四边均正确，容差 ±1mm）
    pgMar = zw_section._sectPr.find(qn('w:pgMar'))
    mar_ok = False
    if pgMar is not None:
        want = {'left': 30, 'right': 25, 'top': 30, 'bottom': 25}
        mar_ok = all(
            (_mm(pgMar.get(qn('w:' + k))) is not None and abs(_mm(pgMar.get(qn('w:' + k))) - v) <= 1.0)
            for k, v in want.items()
        )
    results.append("+3：页边距 左30/右25/上30/下25mm" if mar_ok else "+0：页边距 左30/右25/上30/下25mm")

    # A03 +3：报告正文行间距为固定值22磅
    def _report_body_paras():
        out = []
        lo = zhengwen_idx if zhengwen_idx > 0 else report_lo
        for i in range(lo, len(paras)):
            p = paras[i]
            t = p.text.strip()
            if not t or not re.search(r'[一-鿿]', t):
                continue
            nt = _norm(t)
            if _CHAP_RE.match(nt) or _L1_RE.match(nt) or _L2_RE.match(nt):
                continue
            if nt in ('结论', '参考文献', '参考资料', '附录'):
                continue
            out.append(p)
        return out

    body_ps = _report_body_paras()
    exact22 = 0
    for p in body_ps:
        k, v = _line_spacing(p)
        if k == 'exact' and v is not None and abs(v - 22) <= 1.0:
            exact22 += 1
    a03_ok = len(body_ps) >= 5 and exact22 >= 0.8 * len(body_ps)
    results.append("+3：报告正文行间距为固定值22磅" if a03_ok else "+0：报告正文行间距为固定值22磅")

    # A04 +3：各章题序及标题 小二号(18磅)黑体
    chap_titles = [
        paras[i] for i in range(report_lo, len(paras))
        if _CHAP_RE.match(_norm(paras[i].text)) and _align(paras[i]) == 'C'
    ]
    a04_ok = bool(chap_titles) and all(
        _is_hei(_ea(doc, p)) and (_size_pt(p) is not None and abs(_size_pt(p) - 18) <= 1.0)
        for p in chap_titles
    )
    results.append("+3：各章题序及标题 小二号黑体" if a04_ok else "+0：各章题序及标题 小二号黑体")

    # A05 +3：各节一级题序及标题（X.Y） 小三号(15磅)黑体
    lo_zw = zhengwen_idx if zhengwen_idx > 0 else report_lo
    l1_titles = [
        paras[i] for i in range(lo_zw, len(paras))
        if _L1_RE.match(_norm(paras[i].text)) and not _CHAP_RE.match(_norm(paras[i].text))
        and len(_norm(paras[i].text)) <= 30
    ]
    a05_ok = bool(l1_titles) and all(
        _is_hei(_ea(doc, p)) and (_size_pt(p) is not None and abs(_size_pt(p) - 15) <= 1.0)
        for p in l1_titles
    )
    results.append("+3：各节一级题序及标题 小三号黑体" if a05_ok else "+0：各节一级题序及标题 小三号黑体")

    # A08 +3：报告正文 小四号(12磅)宋体
    song12 = 0
    for p in body_ps:
        if _is_song(_ea(doc, p)) and (_size_pt(p) is not None and abs(_size_pt(p) - 12) <= 1.0):
            song12 += 1
    a08_ok = len(body_ps) >= 5 and song12 >= 0.8 * len(body_ps)
    results.append("+3：报告正文 小四号宋体" if a08_ok else "+0：报告正文 小四号宋体")

    # A09 +1：报告目录标题 小二号(18磅)黑体
    a09_ok = False
    if rep_toc_idx >= 0:
        p = paras[rep_toc_idx]
        a09_ok = _is_hei(_ea(doc, p)) and (_size_pt(p) is not None and abs(_size_pt(p) - 18) <= 1.0)
    results.append("+1：目录标题 小二号黑体" if a09_ok else "+0：目录标题 小二号黑体")

    # A10 +3：报告目录内容 小四号(12磅)宋体
    a10_ok = False
    if rep_toc_idx >= 0 and zhengwen_idx > rep_toc_idx:
        toc_entries = [
            paras[i] for i in range(rep_toc_idx + 1, zhengwen_idx)
            if paras[i].text.strip()
        ]
        if toc_entries:
            ok_cnt = sum(
                1 for p in toc_entries
                if _is_song(_ea(doc, p)) and (_size_pt(p) is not None and abs(_size_pt(p) - 12) <= 1.0)
            )
            a10_ok = ok_cnt >= 0.8 * len(toc_entries)
    results.append("+3：目录内容 小四号宋体" if a10_ok else "+0：目录内容 小四号宋体")

    # A11 +3：结论/参考文献/附录 标题小二黑体、正文小四宋体
    special_titles = ('结论', '参考文献', '参考资料', '附录')
    spec_idx = [
        i for i in range(lo_zw, len(paras))
        if _norm(paras[i].text) in special_titles and _align(paras[i]) == 'C'
    ]
    a11_title_ok = bool(spec_idx) and all(
        _is_hei(_ea(doc, paras[i])) and (_size_pt(paras[i]) is not None and abs(_size_pt(paras[i]) - 18) <= 1.0)
        for i in spec_idx
    )
    # 三部分正文：各标题后首个“实质正文段落”应为宋体小四
    # （跳过“附录 A”之类的小节标题、过短行与非中文行，避免误判子标题为正文）
    a11_body_ok = True
    for i in spec_idx:
        follow = None
        for j in range(i + 1, min(i + 8, len(paras))):
            t = paras[j].text.strip()
            if not t:
                continue
            nt = _norm(t)
            if nt in special_titles:
                continue
            if re.match(r'^附录\s*[A-Za-z]', t):   # 附录小节标题（附录A/附录B…）
                continue
            if len(nt) < 12 or not re.search(r'[一-鿿\[]', t):
                continue
            follow = paras[j]
            break
        if follow is not None:
            if not (_is_song(_ea(doc, follow)) and (_size_pt(follow) is not None and abs(_size_pt(follow) - 12) <= 1.0)):
                a11_body_ok = False
                break
    a11_ok = a11_title_ok and a11_body_ok
    results.append("+3：结论/参考文献/附录 标题小二黑体、正文小四宋体" if a11_ok
                   else "+0：结论/参考文献/附录 标题小二黑体、正文小四宋体")

    # ---- 页眉/页脚辅助 ----
    def _eff_header(sec_idx):
        i = min(sec_idx, len(doc.sections) - 1)
        while i >= 0:
            h = doc.sections[i].header
            if not h.is_linked_to_previous:
                return h
            i -= 1
        return doc.sections[0].header

    def _eff_footer(sec_idx):
        i = min(sec_idx, len(doc.sections) - 1)
        while i >= 0:
            f = doc.sections[i].footer
            if not f.is_linked_to_previous:
                return f
            i -= 1
        return doc.sections[0].footer

    def _footer_has_pagenum(sec_idx):
        f = _eff_footer(sec_idx)
        xml = f._element.xml or ''
        return ('PAGE' in xml) and ('instrText' in xml or 'fldChar' in xml or 'fldSimple' in xml)

    # 找到承载页眉文字的段落（含 _HEADER_TEXT）
    header_para = None
    for s in doc.sections:
        h = s.header
        if h.is_linked_to_previous:
            continue
        for p in h.paragraphs:
            if _HEADER_TEXT in _norm(p.text):
                header_para = p
                break
        if header_para is not None:
            break

    # A12 +3：除封面外各部分页眉文字为“云岭数字技术学院本科课程设计报告”
    results.append("+3：页眉文字为“云岭数字技术学院本科课程设计报告”" if header_para is not None
                   else "+0：页眉文字为“云岭数字技术学院本科课程设计报告”")

    # A13 +1：页眉字体 宋体五号(10.5磅)、居中
    a13_ok = False
    if header_para is not None:
        r = _first_run(header_para)
        nm = _resolve_effective_eastasia(doc, r) if r is not None else None
        sz = r.font.size.pt if (r is not None and r.font.size) else None
        a13_ok = _is_song(nm) and (sz is not None and abs(sz - 10.5) <= 0.6) and _align(header_para) == 'C'
    results.append("+1：页眉宋体五号、居中" if a13_ok else "+0：页眉宋体五号、居中")

    # A14 +1：页眉下方单实线下划线（约0.5磅的单线，容差放宽到 ≤2磅）
    a14_ok = False
    if header_para is not None:
        pPr = header_para._element.find(qn('w:pPr'))
        if pPr is not None:
            pbdr = pPr.find(qn('w:pBdr'))
            if pbdr is not None:
                bottom = pbdr.find(qn('w:bottom'))
                if bottom is not None and bottom.get(qn('w:val')) == 'single':
                    szv = bottom.get(qn('w:sz'))
                    a14_ok = (szv is None) or (szv.isdigit() and int(szv) <= 16)  # 八分之一磅：≤2磅
    results.append("+1：页眉下方单实线下划线" if a14_ok else "+0：页眉下方单实线下划线")

    # A15 +3：报告封面不加页眉页脚
    cover_sec_idx = secmap.get(paras[report_lo]._element, 0)
    cover_hdr = _norm(''.join(p.text for p in _eff_header(cover_sec_idx).paragraphs))
    a15_ok = (_HEADER_TEXT not in cover_hdr and cover_hdr == '') and not _footer_has_pagenum(cover_sec_idx)
    results.append("+3：报告封面不加页眉页脚" if a15_ok else "+0：报告封面不加页眉页脚")

    # A16 +3：封面外各部分页脚页码、阿拉伯连续
    a16_ok = False
    for si in range(len(doc.sections)):
        if si == cover_sec_idx:
            continue
        if _footer_has_pagenum(si):
            a16_ok = True
            break
    results.append("+3：封面外各部分添加页脚页码（阿拉伯、连续）" if a16_ok
                   else "+0：封面外各部分添加页脚页码（阿拉伯、连续）")

    # A17 +5：页码以正文为第1页（正文单独分节并把该节起始页码重设为1）
    a17_ok = False
    if zhengwen_idx > 0:
        zw_sec_idx = secmap.get(paras[zhengwen_idx]._element, 0)
        pgNum = _section_obj_of(zhengwen_idx)._sectPr.find(qn('w:pgNumType'))
        start_val = pgNum.get(qn('w:start')) if pgNum is not None else None
        restart_1 = (start_val == '1')
        isolated = True
        if rep_toc_idx >= 0:
            isolated = isolated and zw_sec_idx > secmap.get(paras[rep_toc_idx]._element, 0)
        if expl_idx >= 0:
            isolated = isolated and zw_sec_idx > secmap.get(paras[expl_idx]._element, 0)
        if rep_toc_idx < 0 and expl_idx < 0:
            isolated = zw_sec_idx > 0
        a17_ok = restart_1 and isolated
    results.append("+5：页码以正文为第1页（正文单独分节、起始页码重设为1）" if a17_ok
                   else "+0：页码以正文为第1页（正文单独分节、起始页码重设为1）")

    # A18 +3：图/表题注位置与阿拉伯编号（图序图名在图下且空一格；表序表名在表正上方；无图表则默认达标）
    cap_re = re.compile(r'^(图|表)\s*(\d+)')
    captions = [p for p in paras if cap_re.match(p.text.strip())]
    if not captions:
        a18_ok = True  # 报告无图表 → 默认达标
    else:
        a18_ok = True
        for p in captions:
            m = cap_re.match(p.text.strip())
            if not m or not m.group(2).isdigit():
                a18_ok = False
                break
            # 表题注应在表格正上方：其后邻近应出现表格
            if m.group(1) == '表':
                el = p._element
                nxt = el.getnext()
                hop = 0
                found_tbl = False
                while nxt is not None and hop < 3:
                    if nxt.tag == qn('w:tbl'):
                        found_tbl = True
                        break
                    hop += 1
                    nxt = nxt.getnext()
                if not found_tbl:
                    a18_ok = False
                    break
    results.append("+3：图/表题注位置与阿拉伯编号（无图表则默认达标）" if a18_ok
                   else "+0：图/表题注位置与阿拉伯编号（无图表则默认达标）")

    # ======================= 减分项 =======================
    # 约定：减分项触发（=破坏原格式）追加 "-N："，未触发追加 "-0："。
    all_text = ''.join(_norm(p.text) for p in paras)

    def _get(pred, lo=0, hi=None):
        i = _idx(pred, lo, hi)
        return paras[i] if i >= 0 else None

    # D01 -1：指导书封面主标题（二号24加粗、居中、1.08倍行距、段后26磅）
    p = _get(lambda t: t == '工程实践与算法应用综合设计指导书', 0, report_lo)
    d01 = (p is not None) and not _preserved(doc, p, _is_noto, 24, 'C', 'mult108', 26)
    results.append("-1：指导书封面主标题格式被改动" if d01 else "-0：指导书封面主标题格式被改动")

    # D02 -1：指导书封面 作者行/“云岭数字技术学院”/年月 三行（居中、1.08倍行距）
    d02_ps = [
        _get(lambda t: t.startswith('林沐川'), 0, report_lo),
        _get(lambda t: t == '云岭数字技术学院', 0, report_lo),
        _get(lambda t: t == '2026.04', 0, report_lo),
    ]
    d02 = any(
        (p is not None) and not _preserved(doc, p, _is_noto, None, 'C', 'mult108')
        for p in d02_ps
    )
    results.append("-1：指导书封面作者/单位/年月三行格式被改动" if d02 else "-0：指导书封面作者/单位/年月三行格式被改动")

    # D03 -1：指导书“目录”标题或条目（二号22加粗居中1.08、条目11磅）
    p_title = _get(lambda t: t == '目录', 0, report_lo)
    d03_title = (p_title is not None) and not _preserved(doc, p_title, _is_noto, 22, 'C', 'mult108')
    d03_entry = False
    if p_title is not None:
        ti = paras.index(p_title)
        entry = None
        for j in range(ti + 1, min(ti + 6, report_lo)):
            if paras[j].text.strip():
                entry = paras[j]
                break
        if entry is not None:
            d03_entry = not _preserved(doc, entry, _is_noto, 11, None, None)
    d03 = d03_title or d03_entry
    results.append("-1：指导书“目录”标题或条目格式被改动" if d03 else "-0：指导书“目录”标题或条目格式被改动")

    # D04 -1：评分标准/备选方向 居中标题（三号16加粗、居中、1.08倍行距、段后8磅）
    d04_anchors = ['工程实践与算法应用综合设计评分标准', '过程性评价与备选方向',
                   '1.成果文本编制要求', '成果提交与附录要求']
    d04 = False
    for a in d04_anchors:
        p = _get(lambda t, a=a: t == a, 0, report_lo)
        if p is not None and not _preserved(doc, p, _is_noto, 16, 'C', 'mult108'):
            d04 = True
            break
    results.append("-1：评分标准/备选方向居中标题格式被改动" if d04 else "-0：评分标准/备选方向居中标题格式被改动")

    # D05 -1：“三、综合设计备选方向”（13磅加粗、左对齐、1.08倍行距、段后4磅）
    p = _get(lambda t: t.startswith('三、综合设计备选方向'), 0, report_lo)
    d05 = (p is not None) and not _preserved(doc, p, _is_noto, 13, 'L', 'mult108')
    results.append("-1：“三、综合设计备选方向”段落格式被改动" if d05 else "-0：“三、综合设计备选方向”段落格式被改动")

    # D06 -1：备选方向 5 条条目（10磅、1.05倍行距、段后1.5磅）
    d06 = False
    ai = _idx(lambda t: t.startswith('三、综合设计备选方向'), 0, report_lo)
    if ai >= 0:
        for j in range(ai + 1, report_lo):
            tt = paras[j].text.strip()
            if tt.startswith('实施安排'):
                break
            if tt.startswith('•') or tt.startswith('·'):
                if not _preserved(doc, paras[j], _is_noto, 10, None, 'mult108'):
                    d06 = True
                    break
    results.append("-1：备选方向5条条目格式被改动" if d06 else "-0：备选方向5条条目格式被改动")

    # D07 -1：“实施安排：第1周…”（10.5磅、1.08倍行距、段后2磅）
    p = _get(lambda t: t.startswith('实施安排'), 0, report_lo)
    d07 = (p is not None) and not _preserved(doc, p, _is_noto, 10.5, None, 'mult108')
    results.append("-1：“实施安排…”段落格式被改动" if d07 else "-0：“实施安排…”段落格式被改动")

    # —— 编制要求正文原为：宋体、小四(12)、1.15倍行距、段后10磅（继承默认 Normal 样式）——
    #    真正的“破坏”是把它误升级为章节标题样式：改黑体、放大字号、或套用固定值行距(如22磅)。
    #    注意：1.15 倍行距是原始默认，不能当作“被加大行距”，否则会误报。
    def _spec_broken(p):
        """编制要求段落是否被破坏：改字体 / 放大字号 / 套用固定值行距。"""
        if p is None:
            return False
        name = _ea(doc, p)
        if name and not _is_song(name):     # 被改成黑体等非宋体
            return True
        s = _size_pt(p)
        if s is not None and s > 13.5:       # 被放大字号（超过小四）
            return True
        k, _v = _line_spacing(p)
        if k == 'exact':                     # 被套用固定行距（如报告正文的22磅）
            return True
        return False

    # D08 -1：编制要求 小标题（宋体、小四加粗、左对齐；不应被升级为章节标题）
    d08_anchors = ['一、报告内容', '二、打印要求', '纸张、页面、装订', '字体与字号', '页码', '页眉与页脚']
    d08 = any(_spec_broken(_get(lambda t, a=a: t == a, 0, report_lo)) for a in d08_anchors)
    results.append("-1：编制要求小标题被误改（黑体/放大/加固定行距）" if d08 else "-0：编制要求小标题被误改（黑体/放大/加固定行距）")

    # D09 -1：编制要求 正文段落（宋体、小四、1.15倍行距、段后10磅；不应被套固定行距/改字体）
    d09_anchors = ['报告包括封面、目录、正文、参考文献三部分。', '参考文献格式如下：',
                   '各章题序及标题为小二号黑体；各节的一级题序与标题为小三号黑体，各节的']
    d09 = any(_spec_broken(_get(lambda t, a=a: t.startswith(a[:12]), 0, report_lo)) for a in d09_anchors)
    results.append("-1：编制要求正文段落被强制改动" if d09 else "-0：编制要求正文段落被强制改动")

    # —— 报告封面（应保持原统一模板格式，指令未要求改封面）——
    # D10 -1：报告封面“云岭数字技术学院”（22磅加粗、居中、1.08倍行距、段后26磅）
    p = _get(lambda t: t == '云岭数字技术学院', max(0, report_lo - 4), report_lo + 1)
    d10 = (p is not None) and not _preserved(doc, p, _is_noto, 22, 'C', 'mult108')
    results.append("-1：报告封面“云岭数字技术学院”格式被改动" if d10 else "-0：报告封面“云岭数字技术学院”格式被改动")

    # D11 -1：报告封面“《工程实践…》/本科设计报告”（18/20磅加粗、居中、1.08倍行距）
    hi11 = expl_idx if expl_idx > 0 else (rep_toc_idx if rep_toc_idx > 0 else len(paras))
    p1 = _get(lambda t: t == '《工程实践与算法应用综合设计》', report_lo, hi11)
    p2 = _get(lambda t: t == '本科设计报告', report_lo, hi11)
    d11 = any(
        (p is not None) and not _preserved(doc, p, _is_noto, None, 'C', 'mult108')
        for p in (p1, p2)
    )
    results.append("-1：报告封面“《…》/本科设计报告”格式被改动" if d11 else "-0：报告封面“《…》/本科设计报告”格式被改动")

    # D12 -1：报告封面信息行（题目/组员姓名/学号/班级/指导教师）（13磅加粗、单倍行距、段后5磅）
    d12 = False
    for pre in ('题目', '组员姓名', '组员学号', '班级', '指导教师'):
        p = _get(lambda t, pre=pre: t.startswith(pre), report_lo, hi11)
        if p is not None and not _preserved(doc, p, _is_noto, None, 'L', 'single'):
            d12 = True
            break
    results.append("-1：报告封面信息行格式被改动" if d12 else "-0：报告封面信息行格式被改动")

    # D13 -1：报告封面“二〇二六年四月”（16磅加粗、居中、1.08倍行距、段后2磅）
    p = _get(lambda t: t == '二〇二六年四月', report_lo, hi11)
    d13 = (p is not None) and not _preserved(doc, p, _is_noto, 16, 'C', 'mult108')
    results.append("-1：报告封面“二〇二六年四月”格式被改动" if d13 else "-0：报告封面“二〇二六年四月”格式被改动")

    # —— 说明页（应原样保留）——
    def _expl_broken(p, need_center):
        """说明页段落是否被破坏：改字体（如误设 MS PGothic 等外文字体）/加固定行距/改对齐。"""
        if p is None:
            return False
        name = _ea(doc, p)
        if name and not _is_noto(name):   # 原为 Noto Sans CJK SC，被改字体
            return True
        k, v = _line_spacing(p)
        if k == 'exact':                  # 被套用固定行距
            return True
        if need_center and _align(p) != 'C':
            return True
        return False

    # D14 -1：说明页标题（“本科设计报告”“说明”，20/16磅加粗、居中、1.08倍行距）
    d14 = False
    if expl_idx >= 0:
        p_shuo = paras[expl_idx]
        p_bk = None
        for j in range(expl_idx - 1, max(report_lo, expl_idx - 4), -1):
            if _norm(paras[j].text) == '本科设计报告':
                p_bk = paras[j]
                break
        d14 = _expl_broken(p_shuo, True) or _expl_broken(p_bk, True)
    results.append("-1：说明页标题被改字体/行距" if d14 else "-0：说明页标题被改字体/行距")

    # D15 -1：说明页正文（“本设计围绕…”11磅、“报告题目：…”12磅，1.08倍行距、段后2磅）
    d15 = False
    if expl_idx >= 0:
        _hi15 = rep_toc_idx if rep_toc_idx > expl_idx else len(paras)
        for pre in ('本设计围绕', '报告题目'):
            p = _get(lambda t, pre=pre: t.startswith(pre), expl_idx, _hi15)
            if _expl_broken(p, False):
                d15 = True
                break
    results.append("-1：说明页正文被改动" if d15 else "-0：说明页正文被改动")

    # —— 通用破坏（针对原文内容与对象）——
    # D16 -3：删除或改写原文文字内容（增删、替换正文/标题/条目文字）
    d16 = not all(s in all_text for s in _CONTENT_SENTINELS)
    results.append("-3：删除或改写原文文字内容" if d16 else "-0：删除或改写原文文字内容")

    # D17 -3：删除或破坏原有表格、插图、流程图等对象
    n_tbl = len(doc.tables)
    n_draw = sum(1 for _ in doc.element.body.iter(qn('w:drawing')))
    d17 = (n_tbl < _ORIG_MIN_TABLES) or (n_draw < _ORIG_MIN_DRAWINGS)
    results.append("-3：删除或破坏原有表格/插图对象" if d17 else "-0：删除或破坏原有表格/插图对象")

    # D18 -3：删除原有段落或章节，导致原文内容缺失
    _sec_ok = all(any(re.search(pat, t) for t in [all_text]) for pat in _SECTION_PATTERNS)
    d18 = not _sec_ok
    results.append("-3：删除原有段落或章节" if d18 else "-0：删除原有段落或章节")

    # D19 -1：改动封面或正文中姓名、学号、日期等原始信息的数值
    d19 = not all(s in all_text for s in _ORIG_IDENTIFIERS)
    results.append("-1：改动姓名/学号/日期等原始信息" if d19 else "-0：改动姓名/学号/日期等原始信息")

    # ---------- 反解：results（"+N："/"-N：" 串）-> 结构化 dim2_items ----------
    if len(results) != len(_MAX_DELTAS_ORDER):
        raise RuntimeError(
            f"结果条目数({len(results)})与 max_delta 序列({len(_MAX_DELTAS_ORDER)})不一致"
        )
    dim2_items = []
    total = 0
    for raw, max_delta in zip(results, _MAX_DELTAS_ORDER):
        head, _, rule = raw.partition("：")
        got = int(head.replace("+", "").replace("-", "").strip())  # 该项写入的绝对分值
        hit = got > 0
        # 加分项（max_delta>=0）：命中记满分，否则 0
        # 减分项（max_delta<0）：触发“破坏”记该负分，否则 0
        delta = max_delta if hit else 0
        total += delta
        dim2_items.append({
            "rule": rule,
            "max_delta": max_delta,
            "delta": delta,
            "hit": hit,
        })

    result["dim2_items"] = dim2_items
    result["total_score"] = total
    result["max_score"] = MAX_SCORE_TOTAL
    return result


if __name__ == "__main__":
    # 本地调试入口：默认评估脚本所在目录，可用 argv[1] 覆盖目标目录。
    _dir = sys.argv[1] if len(sys.argv) > 1 else os.path.dirname(os.path.abspath(__file__))
    print(json.dumps(evaluate(_dir), ensure_ascii=False, indent=2))











