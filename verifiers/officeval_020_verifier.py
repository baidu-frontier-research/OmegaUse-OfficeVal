#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# pyright: reportUnknownMemberType=false, reportUnknownVariableType=false, reportUnknownArgumentType=false, reportUnknownParameterType=false, reportMissingParameterType=false, reportPrivateUsage=false, reportUnknownLambdaType=false, reportGeneralTypeIssues=false

import os
import re
import sys
import json
import zipfile
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn

# 说明：接口统一约定，本模块不修改全局 sys.stdout，也不 print 主结果。
# 主结果统一由 evaluate(dir_path) -> dict 返回，入参为"脚本所在目录的路径"，
# 由脚本自己在该目录内定位并打开被评估的 .docx。本文件仅在 __main__ 下作为
# 本地调试入口，把该 dict 序列化为 JSON 打印。

SCRIPT_ID = "020"
# 维度二加分项 max_delta 之和；用于维度一未通过 / 脚本异常时响应
MAX_SCORE_TOTAL = 54
# 扣分上限：扣分项累计再多也只扣到这个值
MAX_DEDUCTION_TOTAL = 18

# 维度二评分项按"结果追加顺序"的 max_delta 列表，用于末端把 results 里的
# "+N：..." / "-N：..." 字符串拆解为结构化 dim2_items 时使用。顺序须与
# _evaluate_impl 中 results.append(...) 的顺序严格一致；共 26 项：
# 前 18 项为加分项（累加 = MAX_SCORE_TOTAL），后 8 项为扣分项（max_delta 为负）。
_MAX_DELTAS_ORDER = [
    1,   # 01 页边距上下 2.6cm
    1,   # 02 页边距左右 2.4cm
    1,   # 03 课题四项（黑体/固定值38磅/居中/缩进0）
    5,   # 04 12 个板块名称顶格
    3,   # 05 12 个板块名称小四
    1,   # 06 12 个板块名称不加粗
    3,   # 07 一级标题四项（黑体/小四/不加粗/缩进2字符）
    3,   # 08 二级标题黑体
    3,   # 09 二级标题小四
    3,   # 10 三级标题宋体（教师配乐范读）
    5,   # 11 三级标题加粗（教师配乐范读）
    1,   # 12 三级标题小四（人物言行探究表）
    5,   # 13 正文宋体（教学重点）
    3,   # 14 正文小四（教师准备）
    5,   # 15 课题以外全文 1.5 倍行距
    5,   # 16 首行缩进 2 字符（教学目标第 1 条）
    3,   # 17 无负左缩进（②标注你认为最能体现）
    3,   # 18 无悬挂缩进（抓住人物言行举止）
    -5,  # 19 12 个板块名称缺失
    -5,  # 20 教学思路说明内容残缺
    -5,  # 21 板书设计图片丢失
    -3,  # 22 编号重号
    -1,  # 23 教学重难点前无空行 / 连续 3 个以上空段
    -3,  # 24 纸张非 A4 纵向
    -3,  # 25 新增页眉页脚页码分页符
    -1,  # 26 套用内置标题样式
]
assert sum(d for d in _MAX_DELTAS_ORDER if d > 0) == MAX_SCORE_TOTAL
# 12 个板块名称（按教案固定顺序）
BLOCK_NAMES = [
    "教材分析", "学情分析", "设计理念", "教学目标", "教学重难点", "教学准备",
    "课时安排", "教学思路说明", "教学过程", "板书设计", "作业设计", "教学反思",
]
# 一级标题：中文数字 + "、"；二级标题：（中文数字）
RE_L1 = re.compile(r"^[一二三四五六七八九十]+、")
RE_L2 = re.compile(r"^（[一二三四五六七八九十]+）")
# 字体别名：中文名与西文名视为同一字体
FONT_SONG = {"宋体", "SimSun", "NSimSun", "Songti SC", "STSong"}
FONT_HEI = {"黑体", "SimHei", "Heiti SC", "STHeiti", "MS Hei"}
EMU_PER_INCH = 914400.0
# 小四 = 12 磅；课题固定行距 38 磅；首行缩进 2 字符 = firstLineChars 200
SIZE_XIAOSI = 12.0
TITLE_LINE_PT = 38.0
INDENT_2CHARS = "200"
def _cm(emu):
    """EMU -> 厘米，保留 2 位小数；None 原样返回。"""
    if emu is None:
        return None
    return round(emu / EMU_PER_INCH * 2.54, 2)


def _style_elements_chain(doc, para):
    """段落样式链：pStyle -> basedOn -> ...，返回 w:style 元素列表。"""
    chain = []
    try:
        styles_el = doc.styles.element
    except Exception:
        return chain
    pPr = para._element.pPr
    style_id = None
    if pPr is not None and pPr.pStyle is not None:
        style_id = pPr.pStyle.get(qn("w:val"))
    seen = set()
    while style_id and style_id not in seen:
        seen.add(style_id)
        found = None
        for st in styles_el.findall(qn("w:style")):
            if st.get(qn("w:styleId")) == style_id:
                found = st
                break
        if found is None:
            break
        chain.append(found)
        based = found.find(qn("w:basedOn"))
        style_id = based.get(qn("w:val")) if based is not None else None
    return chain


def _doc_defaults(doc, kind):
    """kind='rPr' 取 docDefaults/rPrDefault/rPr；kind='pPr' 取 pPrDefault/pPr。"""
    try:
        styles_el = doc.styles.element
    except Exception:
        return None
    dd = styles_el.find(qn("w:docDefaults"))
    if dd is None:
        return None
    holder = dd.find(qn("w:%sDefault" % kind))
    if holder is None:
        return None
    return holder.find(qn("w:" + kind))
def _rPr_chain(doc, run):
    """字符属性继承链：run 的 rPr -> 段落样式链的 rPr -> docDefaults 的 rPr。"""
    out = []
    if run._element.rPr is not None:
        out.append(run._element.rPr)
    para = run._parent
    for st in _style_elements_chain(doc, para):
        rpr = st.find(qn("w:rPr"))
        if rpr is not None:
            out.append(rpr)
    dd = _doc_defaults(doc, "rPr")
    if dd is not None:
        out.append(dd)
    return out


def _resolve_effective_eastasia(doc, run):
    """解析 run 实际生效的中文字体名（run -> 样式链 -> docDefaults）。"""
    for rpr in _rPr_chain(doc, run):
        fonts = rpr.find(qn("w:rFonts"))
        if fonts is not None:
            val = fonts.get(qn("w:eastAsia"))
            if val:
                return val
    return None


def _resolve_effective_ascii(doc, run):
    """解析 run 实际生效的西文字体名（run -> 样式链 -> docDefaults）。"""
    for rpr in _rPr_chain(doc, run):
        fonts = rpr.find(qn("w:rFonts"))
        if fonts is not None:
            val = fonts.get(qn("w:ascii"))
            if val:
                return val
    return None


def _resolve_effective_size_pt(doc, run):
    """解析 run 实际生效的字号（磅）；w:sz 为半磅值。"""
    for rpr in _rPr_chain(doc, run):
        sz = rpr.find(qn("w:sz"))
        if sz is not None:
            val = sz.get(qn("w:val"))
            if val:
                try:
                    return float(val) / 2.0
                except ValueError:
                    return None
    return None


def _resolve_effective_bold(doc, run):
    """解析 run 实际是否加粗；w:b 无 val 属性时视为 True。"""
    for rpr in _rPr_chain(doc, run):
        b = rpr.find(qn("w:b"))
        if b is not None:
            val = b.get(qn("w:val"))
            return val not in ("0", "false", "off")
    return False
def _pPr_chain(doc, para):
    """段落属性继承链：段落 pPr -> 样式链 pPr -> docDefaults 的 pPr。"""
    out = []
    if para._element.pPr is not None:
        out.append(para._element.pPr)
    for st in _style_elements_chain(doc, para):
        ppr = st.find(qn("w:pPr"))
        if ppr is not None:
            out.append(ppr)
    dd = _doc_defaults(doc, "pPr")
    if dd is not None:
        out.append(dd)
    return out


def _eff_ind(doc, para, attr):
    """解析段落实际生效的 w:ind 某个属性（firstLine/firstLineChars/left/hanging...）。"""
    for ppr in _pPr_chain(doc, para):
        ind = ppr.find(qn("w:ind"))
        if ind is not None:
            val = ind.get(qn("w:" + attr))
            if val is not None:
                return val
    return None


def _eff_spacing(doc, para):
    """解析段落实际生效的行距，返回 (line, lineRule)，line 为字符串原值。"""
    for ppr in _pPr_chain(doc, para):
        sp = ppr.find(qn("w:spacing"))
        if sp is not None:
            line = sp.get(qn("w:line"))
            if line is not None:
                return line, sp.get(qn("w:lineRule"))
    return None, None


def _eff_align(doc, para):
    """解析段落实际生效的对齐方式（w:jc 的 val 原值）。"""
    for ppr in _pPr_chain(doc, para):
        jc = ppr.find(qn("w:jc"))
        if jc is not None:
            return jc.get(qn("w:val"))
    return None


def _text_runs(para):
    """段落中有实际文字的 run（跳过纯空白 run）。"""
    return [r for r in para.runs if r.text.strip()]


def _is_zero_ind(val):
    """缩进量视为 0：未设置或数值为 0。"""
    if val is None:
        return True
    try:
        return int(val) == 0
    except ValueError:
        return False
def _para_size_pt(doc, para):
    """段落字号：取首个有文字 run 生效的字号。"""
    rs = _text_runs(para)
    return _resolve_effective_size_pt(doc, rs[0]) if rs else None


def _para_bold(doc, para):
    """段落是否加粗：取首个有文字 run 的生效加粗值。"""
    rs = _text_runs(para)
    return _resolve_effective_bold(doc, rs[0]) if rs else False


def _para_is_hei(doc, para):
    """段落是否黑体：取首个有文字 run 的生效中文字体。"""
    rs = _text_runs(para)
    if not rs:
        return False
    return (_resolve_effective_eastasia(doc, rs[0]) or "") in FONT_HEI


def _para_is_song(doc, para):
    """段落是否整段宋体：所有有文字 run 的生效中文字体均为宋体。"""
    rs = _text_runs(para)
    if not rs:
        return False
    return all((_resolve_effective_eastasia(doc, r) or "宋体") in FONT_SONG for r in rs)


def _para_line_15(doc, para):
    """段落是否 1.5 倍行距：lineRule 为 auto/未设置且 line = 360。"""
    line, rule = _eff_spacing(doc, para)
    if line is None:
        return False
    if rule not in (None, "auto"):
        return False
    try:
        return abs(float(line) - 360.0) < 1.0
    except ValueError:
        return False


def _find_para(paras, keyword):
    """按关键字定位首个包含该文字的段落；找不到返回 None。"""
    for p in paras:
        if keyword in p.text:
            return p
    return None
def _block_indexes(paras):
    """12 个板块名称所在段落下标（按段首匹配）。"""
    out = []
    for i, p in enumerate(paras):
        t = p.text.strip()
        if any(t.startswith(b) for b in BLOCK_NAMES):
            out.append(i)
    return out


def _block_span(paras, blocks, name):
    """某板块的段落区间 [start, end)：从该板块名称段到下一个板块名称段。"""
    start = None
    for i in blocks:
        if paras[i].text.strip().startswith(name):
            start = i
            break
    if start is None:
        return None, None
    later = [j for j in blocks if j > start]
    return start, (later[0] if later else len(paras))


def _dup_numbers(paras):
    """相邻两条手写编号使用同一序号（重号）的段落下标。"""
    seq = []
    for i, p in enumerate(paras):
        m = re.match(r"^(\d+)\s*[.．]", p.text.strip())
        if m:
            seq.append((i, int(m.group(1))))
    return [seq[k][0] for k in range(1, len(seq))
            if seq[k][0] == seq[k - 1][0] + 1 and seq[k][1] == seq[k - 1][1]]


def _open_word_document(file_path):
    """打开 .docx；失败时抛出异常交由上层归入维度一未通过。"""
    return Document(file_path)


def _is_editable_docx(file_path):
    """是否为可编辑 docx：扩展名正确且包内含 word/document.xml。"""
    if not file_path.lower().endswith(".docx"):
        return False
    try:
        with zipfile.ZipFile(file_path) as zf:
            return "word/document.xml" in zf.namelist()
    except Exception:
        return False


def _is_protected(doc):
    """是否被强制保护（documentProtection 且 enforcement 打开）。"""
    m = re.search(r"<w:documentProtection[^>]*>", doc.element.xml)
    if not m:
        return False
    return bool(re.search(r'w:enforcement="(1|true|on)"', m.group(0)))
def evaluate(dir_path):
    """统一入口：入参为脚本所在目录，返回结构化评测结果 dict。"""
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
        result["file_name"] = cands[0]
        file_path = os.path.join(dir_path, cands[0])
        _evaluate_impl(file_path, result)
    except Exception as exc:
        # status 为 error 时未能完成判定，dim1_pass 留空，dim1_reason 也为空
        result["status"] = "error"
        result["error"] = "%s: %s" % (type(exc).__name__, exc)
        result["dim1_pass"] = ""
        result["dim1_reason"] = ""
        result["dim2_items"] = []
        result["total_score"] = 0
    return result
def _evaluate_impl(file_path, result):
    """维度一门槛 + 维度二 26 项细则；结果写回 result。"""
    # ---------- 维度一：可用性和可修改性（门槛，不通过则总分 0） ----------
    # 只检查整体文件级别的重大问题：能否打开、是否可编辑 docx、是否被强制保护。
    doc = None
    reasons = []
    if not (os.path.isfile(file_path) and os.path.getsize(file_path) > 0):
        reasons.append("文件不存在或为 0 字节，无法打开")
    else:
        try:
            doc = _open_word_document(file_path)
        except Exception as exc:
            reasons.append("文件无法正常打开（%s）" % type(exc).__name__)
    if not _is_editable_docx(file_path):
        reasons.append("不是可编辑的 docx 格式")
    if doc is not None and _is_protected(doc):
        reasons.append("文件被加密或开启了强制编辑保护")
    if reasons or doc is None:
        result["dim1_pass"] = False
        result["dim1_reason"] = "；".join(reasons) or "文件无法打开"
        result["dim2_items"] = []
        result["total_score"] = 0
        return

    paras = doc.paragraphs
    section = doc.sections[0]
    blocks = _block_indexes(paras)
    l1 = [i for i, p in enumerate(paras) if RE_L1.match(p.text.strip())]
    l2 = [i for i, p in enumerate(paras) if RE_L2.match(p.text.strip())]
    results = []

    # ---------- 维度二：完成度（加分项 18 项，合计 +54） ----------
    # 01 页边距上下 2.6cm
    ok = _cm(section.top_margin) == 2.6 and _cm(section.bottom_margin) == 2.6
    results.append("+1：页面上下页边距均为 2.6 厘米。" if ok
                   else "+0：页面上下页边距未改为 2.6 厘米。")

    # 02 页边距左右 2.4cm
    ok = _cm(section.left_margin) == 2.4 and _cm(section.right_margin) == 2.4
    results.append("+1：页面左右页边距均为 2.4 厘米。" if ok
                   else "+0：页面左右页边距未改为 2.4 厘米。")
    # 03 课题四项：黑体 + 固定值 38 磅 + 居中 + 左右及首行缩进 0
    title = paras[0]
    trs = _text_runs(title)
    line, rule = _eff_spacing(doc, title)
    ok_hei = bool(trs) and (_resolve_effective_eastasia(doc, trs[0]) or "") in FONT_HEI
    ok_line = False
    if rule == "exact" and line is not None:
        try:
            ok_line = abs(float(line) / 20.0 - TITLE_LINE_PT) < 0.6
        except ValueError:
            ok_line = False
    ok_pos = (_eff_align(doc, title) == "center"
              and _is_zero_ind(_eff_ind(doc, title, "left"))
              and _is_zero_ind(_eff_ind(doc, title, "right"))
              and _is_zero_ind(_eff_ind(doc, title, "firstLine")))
    results.append("+1：课题《浅赏叙事古文，读懂处世道理》为黑体、固定值 38 磅行距、居中且左右与首行缩进为 0。"
                   if (ok_hei and ok_line and ok_pos)
                   else "+0：课题《浅赏叙事古文，读懂处世道理》未同时满足黑体、固定值 38 磅行距、居中、缩进为 0。")

    # 04 12 个板块名称顶格（首行缩进与左缩进均为 0）
    bad = [i for i in blocks
           if not _is_zero_ind(_eff_ind(doc, paras[i], "firstLine"))
           or not _is_zero_ind(_eff_ind(doc, paras[i], "firstLineChars"))
           or not _is_zero_ind(_eff_ind(doc, paras[i], "left"))]
    results.append("+5：教材分析至教学反思 12 个板块名称全部顶格，无首行缩进与左缩进。" if not bad
                   else "+0：教材分析至教学反思 12 个板块名称中仍有 %d 处未顶格。" % len(bad))

    # 05 12 个板块名称小四
    bad = [i for i in blocks if _para_size_pt(doc, paras[i]) != SIZE_XIAOSI]
    results.append("+3：教材分析至教学反思 12 个板块名称字号全部为小四。" if not bad
                   else "+0：教材分析至教学反思 12 个板块名称中仍有 %d 处不是小四。" % len(bad))

    # 06 12 个板块名称不加粗
    bad = [i for i in blocks if _para_bold(doc, paras[i])]
    results.append("+1：教材分析至教学反思 12 个板块名称全部未加粗。" if not bad
                   else "+0：教材分析至教学反思 12 个板块名称中仍有 %d 处保留加粗。" % len(bad))
    # 07 一级标题（一、二、三……）四项：黑体 + 小四 + 不加粗 + 首行缩进 2 字符
    b_font = [i for i in l1 if not _para_is_hei(doc, paras[i])]
    b_size = [i for i in l1 if _para_size_pt(doc, paras[i]) != SIZE_XIAOSI]
    b_bold = [i for i in l1 if _para_bold(doc, paras[i])]
    b_ind = [i for i in l1
             if _eff_ind(doc, paras[i], "firstLineChars") != INDENT_2CHARS
             or not _is_zero_ind(_eff_ind(doc, paras[i], "left"))]
    nbad = len(b_font) + len(b_size) + len(b_bold) + len(b_ind)
    results.append("+3：教学过程中「一、」至「五、」各一级标题均为黑体、小四、不加粗且首行缩进 2 字符。"
                   if nbad == 0
                   else "+0：教学过程中一级标题仍有不合格项（非黑体 %d、非小四 %d、加粗 %d、缩进不为 2 字符 %d）。"
                        % (len(b_font), len(b_size), len(b_bold), len(b_ind)))

    # 08 二级标题（（一）（二）……）黑体
    bad = [i for i in l2 if not _para_is_hei(doc, paras[i])]
    results.append("+3：教学过程中「（一）」至「（十一）」各二级标题均为黑体。" if not bad
                   else "+0：教学过程中二级标题仍有 %d 处不是黑体。" % len(bad))

    # 09 二级标题小四
    bad = [i for i in l2 if _para_size_pt(doc, paras[i]) != SIZE_XIAOSI]
    results.append("+3：教学过程中「（一）」至「（十一）」各二级标题字号均为小四。" if not bad
                   else "+0：教学过程中二级标题仍有 %d 处不是小四。" % len(bad))

    # 10 三级标题宋体：以「教师配乐范读」所在段落为准
    p = _find_para(paras, "教师配乐范读")
    results.append("+3：「教师配乐范读」等三级标题为宋体。"
                   if p is not None and _para_is_song(doc, p)
                   else "+0：「教师配乐范读」等三级标题未改为宋体。")

    # 11 三级标题加粗：以「教师配乐范读」所在段落为准
    results.append("+5：「教师配乐范读」等三级标题已加粗。"
                   if p is not None and _para_bold(doc, p)
                   else "+0：「教师配乐范读」等三级标题未加粗。")

    # 12 三级标题小四：以「人物言行探究表」所在段落为准
    p = _find_para(paras, "人物言行探究表")
    results.append("+1：「人物言行探究表」所在三级标题字号为小四。"
                   if p is not None and _para_size_pt(doc, p) == SIZE_XIAOSI
                   else "+0：「人物言行探究表」所在三级标题字号不是小四。")
    # 13 正文宋体：以教学重难点中「教学重点」所在段落为准
    p = _find_para(paras, "教学重点")
    results.append("+5：「教学重点」所在正文段落为宋体。"
                   if p is not None and _para_is_song(doc, p)
                   else "+0：「教学重点」所在正文段落不是宋体。")

    # 14 正文小四：以教学准备中「教师准备」所在段落为准
    p = _find_para(paras, "教师准备")
    results.append("+3：「教师准备」所在正文段落字号为小四。"
                   if p is not None and _para_size_pt(doc, p) == SIZE_XIAOSI
                   else "+0：「教师准备」所在正文段落字号不是小四。")

    # 15 课题以外全文 1.5 倍行距：以教材分析首段与「教学过程」板块名为准
    p1 = _find_para(paras, "叙事古文是初中")
    p2 = _find_para(paras, "教学过程")
    ok = (p1 is not None and _para_line_15(doc, p1)
          and p2 is not None and _para_line_15(doc, p2))
    results.append("+5：除课题外，教材分析首段与「教学过程」板块名等段落行距均为 1.5 倍。" if ok
                   else "+0：除课题外的段落未统一设置为 1.5 倍行距。")

    # 16 首行缩进 2 字符：以教学目标第 1 条「感悟文本中蕴含……」为准
    p = _find_para(paras, "感悟文本中蕴含")
    results.append("+5：教学目标第 1 条「感悟文本中蕴含……」首行缩进 2 字符。"
                   if p is not None and _eff_ind(doc, p, "firstLineChars") == INDENT_2CHARS
                   else "+0：教学目标第 1 条「感悟文本中蕴含……」首行未缩进 2 字符。")

    # 17 无负左缩进：以「②标注你认为最能体现……」为准
    p = _find_para(paras, "标注你认为最能体现")
    lv = _eff_ind(doc, p, "left") if p is not None else None
    ok = p is not None and (lv is None or int(lv) >= 0)
    results.append("+3：「②标注你认为最能体现……」左缩进已改为非负值，文字未超出左边界。" if ok
                   else "+0：「②标注你认为最能体现……」仍为负左缩进，文字超出左边界。")

    # 18 无悬挂缩进：以「抓住人物言行举止……」为准
    p = _find_para(paras, "抓住人物言行举止")
    lv = _eff_ind(doc, p, "left") if p is not None else None
    ok = (p is not None
          and not _eff_ind(doc, p, "hanging")
          and not _eff_ind(doc, p, "hangingChars")
          and (lv is None or int(lv) <= 480))
    results.append("+3：「抓住人物言行举止……」已取消悬挂缩进，与相邻段落左端对齐。" if ok
                   else "+0：「抓住人物言行举止……」仍保留悬挂缩进。")
    # ---------- 维度二：完成度（扣分项 8 项，累计扣分上限 18） ----------
    # 19 12 个板块名称缺失
    miss = [b for b in BLOCK_NAMES
            if not any(p.text.strip().startswith(b) for p in paras)]
    results.append("-5：教材分析至教学反思 12 个板块名称缺失 %d 个（%s）。" % (len(miss), "、".join(miss))
                   if miss else "-0：教材分析至教学反思 12 个板块名称齐全。")

    # 20 教学思路说明内容残缺（同段余料 + 后续段落合计不足 30 字）
    s, e = _block_span(paras, blocks, "教学思路说明")
    if s is None:
        txt = ""
    else:
        head = paras[s].text.strip()[len("教学思路说明"):].strip("：: ")
        txt = head + "".join(paras[j].text.strip() for j in range(s + 1, e))
    results.append("-5：教学思路说明板块正文残缺，仅剩 %d 字。" % len(txt)
                   if (s is not None and len(txt) < 30)
                   else "-0：教学思路说明板块正文完整。")

    # 21 板书设计图片丢失
    s, e = _block_span(paras, blocks, "板书设计")
    xml = "".join(paras[j]._element.xml for j in range(s, e)) if s is not None else ""
    npic = xml.count("<w:drawing") + xml.count("<w:pict")
    results.append("-5：板书设计板块的图片丢失。"
                   if (s is not None and npic == 0)
                   else "-0：板书设计板块的图片保留完好。")

    # 22 编号重号（相邻两条编号使用同一序号）
    dup = _dup_numbers(paras)
    results.append("-3：文中编号出现重号，相邻两条编号内容使用了同一个序号（%d 处）。" % len(dup)
                   if dup else "-0：文中编号未出现重号。")

    # 23 教学重难点前无空行 / 出现连续 3 个以上空段
    hi = [i for i, p in enumerate(paras) if p.text.strip().startswith("教学重难点")]
    empt = [i for i, p in enumerate(paras) if not p.text.strip()]
    es = set(empt)
    run3 = any((i in es and i + 1 in es and i + 2 in es) for i in empt)
    nospace = bool(hi) and hi[0] > 0 and bool(paras[hi[0] - 1].text.strip())
    results.append("-1：教学重难点板块前的空行被删除或文中出现连续 3 个以上空段。"
                   if (nospace or run3)
                   else "-0：板块之间的空行保持原状，无连续 3 个以上空段。")

    # 24 纸张非 A4 纵向
    ok_page = _cm(section.page_width) == 21.0 and _cm(section.page_height) == 29.7
    results.append("-3：纸张被改为非 A4 纵向（%s×%s 厘米）。"
                   % (_cm(section.page_width), _cm(section.page_height))
                   if not ok_page else "-0：纸张仍为 A4 纵向。")

    # 25 新增页眉页脚页码或分页符
    xml_all = doc.element.xml
    hf = (any(q.text.strip() for q in section.header.paragraphs)
          or any(q.text.strip() for q in section.footer.paragraphs))
    fired = hf or xml_all.count('w:br w:type="page"') > 0 or "PAGE" in xml_all
    results.append("-3：文档新增了页眉、页脚、页码或分页符。" if fired
                   else "-0：文档未新增页眉、页脚、页码或分页符。")

    # 26 套用内置标题样式
    names = {q.style.name for q in paras if q.style is not None}
    fired = any(n.startswith("Heading") or n.startswith("标题") for n in names)
    results.append("-1：正文段落被套用了内置标题样式。" if fired
                   else "-0：正文段落未套用内置标题样式。")
    # ---------- 把 results 拆解为结构化 dim2_items 并汇总总分 ----------
    if len(results) != len(_MAX_DELTAS_ORDER):
        raise RuntimeError("dim2 项数不一致：results=%d vs order=%d"
                           % (len(results), len(_MAX_DELTAS_ORDER)))
    items = []
    plus = 0
    minus = 0
    for text, max_delta in zip(results, _MAX_DELTAS_ORDER):
        head, _, body = text.partition("：")
        try:
            got = int(head)
        except ValueError:
            got = 0
        if max_delta > 0:
            delta = max_delta if got == max_delta else 0
            hit = delta != 0
            plus += delta
        else:
            delta = max_delta if got != 0 else 0
            hit = delta != 0
            minus += -delta
        items.append({
            "rule": body,
            "max_delta": max_delta,
            "delta": delta,
            "hit": hit,
        })
    result["dim2_items"] = items
    result["total_score"] = max(0, plus - min(minus, MAX_DEDUCTION_TOTAL))


if __name__ == "__main__":
    _dir = sys.argv[1] if len(sys.argv) >= 2 else os.path.dirname(os.path.abspath(__file__))
    _out = evaluate(_dir)
    print(json.dumps(_out, ensure_ascii=False, indent=2))
