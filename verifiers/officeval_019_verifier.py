#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# pyright: reportUnknownMemberType=false, reportUnknownVariableType=false, reportUnknownArgumentType=false, reportMissingImports=false, reportAttributeAccessIssue=false, reportOptionalMemberAccess=false, reportIndexIssue=false
"""
初中化学模拟试题「加封面 / 套模板排版 / 加页码」任务的自动评测脚本（SCRIPT_ID = 019）。

本模块不修改全局 sys.stdout，也不 print 主结果。主结果统一由 evaluate(dir_path) -> dict
返回；入参为"脚本所在目录的路径"，脚本会在该目录下自动寻找待评的 .docx 文件。
本文件仅在 __main__ 下作为本地调试入口，把 evaluate() 的返回值按 JSON 打印出来。

维度2 共 35 条：加分项 17 条（指令下达的任务，满分 +51）、减分项 18 条（指令未要求
而多做的破坏性修改，下限 -36）。加分项命中记 "+N"、未命中记 "+0"；减分项命中（即
确实出现了破坏）记 "-N"、未出现记 "-0"。两类条目在 _MAX_DELTAS_ORDER 中以正负号区分。

本脚本不做任何排版渲染，所有判定均基于 OOXML 静态解析。原先依赖真实分页的条目改为
可静态验证的结构判据（页码域类型、分节与 pgNumType 起编设置）；确实只能在排版后才能
测量的量（页底连续空白高度）按"加分需证据、减分需证据"降级：加分项判为未命中，
减分项判为未出现。
"""

import os
import re
import sys
import json
import traceback

from docx import Document
from docx.oxml.ns import qn
from lxml import etree


SCRIPT_ID = "019"
MAX_SCORE_TOTAL = 51
MIN_SCORE_TOTAL = -36

_MAX_DELTAS_ORDER = [
    5,   # 01 封面单独成页、正文另起一页
    5,   # 02 封面主标题真正竖排
    1,   # 03 封面主标题加粗、≥18磅、水平居中
    1,   # 04 封面四项填写项齐全
    5,   # 05 四项填空横线为下划线格式且≥5cm
    1,   # 06 四项位于封面下半部、左端对齐
    1,   # 07 卷内主题标题 黑体三号居中
    5,   # 08 正文中文 宋体小四
    3,   # 09 正文西文数字 Times New Roman 小四
    1,   # 10 页脚"化学试题 第x页 共x页"
    3,   # 11 当前页序号用 PAGE 域
    3,   # 12 "共x页"与正文实际页数一致
    5,   # 13 封面页不显示页码
    3,   # 14 正文首页页码从"第1页"开始
    1,   # 15 页脚居中、9-12磅、宋体+Times New Roman
    5,   # 16 A4 纵向 + 上下左右 2.5cm 页边距
    3,   # 17 正文段间距0行、行距1.5倍
    -3,  # 18 第25题表格"用于食品包装袋内作保护气"被改写
    -3,  # 19 6张示意图未全部保留在对应题目下
    -3,  # 20 图2/图3/图5 非等比缩放
    -1,  # 21 图2 未独占一个居中段落
    -3,  # 22 正文页底出现 >5cm 连续空白
    -3,  # 23 第31题 6行作答横线折行
    -1,  # 24 化学式/相对原子质量内被插入多余空格
    -1,  # 25 卷内主题标题未加粗
    -3,  # 26 5处分部标题不再是 黑体13.5磅加粗
    -3,  # 27 题号分值表结构/数值被破坏
    -3,  # 28 选择题 20 小题或 ABCD 选项不全
    -1,  # 29 第21-31题题号不连续
    -1,  # 30 分值标注缺失
    -1,  # 31 第25题表格 5行2列/表头被破坏
    -3,  # 32 参考答案标题及 4 段答案未完整保留
    -1,  # 33 参考答案未与正文分页
    -1,  # 34 末尾缺少单独居中的"（试卷结束）"
    -1,  # 35 封面出现多余内容
]

assert sum(d for d in _MAX_DELTAS_ORDER if d > 0) == MAX_SCORE_TOTAL
assert sum(d for d in _MAX_DELTAS_ORDER if d < 0) == MIN_SCORE_TOTAL


# ========== 常量 ==========

_P = qn("w:p")
_TBL = qn("w:tbl")
_TR = qn("w:tr")
_TC = qn("w:tc")
_R = qn("w:r")
_T = qn("w:t")
_BR = qn("w:br")
_TAB = qn("w:tab")
_PPR = qn("w:pPr")
_RPR = qn("w:rPr")
_SECTPR = qn("w:sectPr")
_VAL = qn("w:val")
_TYPE = qn("w:type")
_EMBED = qn("r:embed")

NS_MC = "http://schemas.openxmlformats.org/markup-compatibility/2006"
NS_WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main"
_FALLBACK = "{%s}Fallback" % NS_MC
_ALT_CONTENT = "{%s}AlternateContent" % NS_MC
_DRAWING = qn("w:drawing")
_PICT = qn("w:pict")

EMU_PER_CM = 360000.0
TWIPS_PER_CM = 566.9291
PT_PER_CM = 28.34646

STRICT_SONGTI = ("宋体", "SimSun", "NSimSun", "宋体-简", "Songti SC", "STSong")
STRICT_HEITI = ("黑体", "SimHei", "黑体-简", "Heiti SC", "STHeiti")
TIMES = ("times new roman", "timesnewroman", "times")

TITLE_TEXT = "初中化学模拟试题"
ANSWER_TITLE = "参考答案与评分建议"
END_MARK = "（试卷结束）"
COVER_FIELDS = (
    ("报送单位", r"报送单位\s*[:：]"),
    ("学科", r"^\s*学科\s*[:：]"),
    ("命题人", r"^\s*命题人\s*[:：]"),
    ("命题人联系电话", r"命题人联系电话\s*[:：]"),
)
SECTION_HEADS = ("注意事项", "第一部分", "第二部分", "第三部分", "第四部分")
UNDERLINE_MIN_CM = 5.0
ANSWER_PARTS = ("一、选择题", "二、填空题", "三、实验探究题", "四、综合应用与计算题")
EXPECTED_SCORE_MARKS = 32          # 输入试题中"（x分）"类分值标注的总数
EXPECTED_FIG_ANCHORS = [5, 24, 27, 28, 29, 30]   # 6 张示意图应挂在的题号
FIG_ASPECTS = {2: 1.96, 3: 2.00, 5: 1.61}        # 图2/图3/图5 的应有宽高比（0基序号）
QNUM_RE = re.compile(r"^\s*(\d{1,2})\s*[.．、]")
SCORE_MARK_RE = re.compile(r"（[^）]{0,14}分）")
FULLWIDTH_RE = re.compile(
    "[ᄀ-ᅟ⺀-꓏ꥠ-꥿가-힣豈-﫿"
    "︐-︙︰-﹯＀-｠￠-￦　-〾]"
)
CJK_RE = re.compile("[㐀-䶿一-鿿豈-﫿　-〿！-～]")
LATIN_RE = re.compile(r"[A-Za-z0-9]")
FORMULA_SPACE_PATTERNS = (
    # 选项之间"…含O₂    B. …"式的多空格间隔属正常排版，故限定 1-2 个空格并排除选项标号
    (r"[A-Za-z][₀-₉][ ]{1,2}(?![A-D]\s*[.．、])[A-Za-z]", "化学式内部被拆开（如 2H₂ O）"),
    (r"[₀-₉][ ]{1,2}[↑↓]", "气体/沉淀符号前被插入空格（如 CO₂ ↑）"),
    (r"[A-Za-z]-\d+\.[ ]{1,2}\d", "相对原子质量被拆开（如 Cl-35. 5）"),
    (r"[A-Za-z][ ]{1,2}[₀-₉]", "元素符号与下标之间被插入空格"),
    (r"[₀-₉][ ]{1,2}[²³⁴⁻⁺]", "离子符号内部被插入空格（如 O₂ ²⁻）"),
)


# ========== 基础 XML 工具 ==========

def _walk(el):
    """深度优先遍历子孙节点，跳过 mc:Fallback（VML 兼容分支），避免同一段文字被重复统计。"""
    for child in el:
        if child.tag == _FALLBACK:
            continue
        yield child
        for sub in _walk(child):
            yield sub


def _text_of(el):
    """取节点内的可见文字；w:tab -> \\t，w:br/w:cr -> \\n（含文本框内文字）。"""
    parts = []
    for node in _walk(el):
        tag = node.tag
        if tag == _T:
            parts.append(node.text or "")
        elif tag == _TAB:
            parts.append("\t")
        elif tag == _BR or tag == qn("w:cr"):
            parts.append("\n")
    return "".join(parts)


def _norm(text):
    return re.sub(r"[ \t\u3000]+", " ", (text or "").replace("\n", " ")).strip()


def _squeeze(text):
    return re.sub(r"\s+", "", text or "")


def _is_shape_holder(run):
    """判断运行是否只是图形/文本框的宿主：这类运行自身的 w:rPr 并不作用于文本框内文字，
    其 w:t 实际属于 txbxContent 内的段落，若按普通文字运行统计会误判字体字号。"""
    for node in _walk(run):
        if node.tag in (_DRAWING, _PICT, _ALT_CONTENT):
            return True
    return False


def _runs_of(el):
    return [n for n in _walk(el) if n.tag == _R and not _is_shape_holder(n)]


def _paras_of(el):
    return [n for n in _walk(el) if n.tag == _P]


def _pPr(p):
    return p.find(_PPR)


def _rPr(r):
    return r.find(_RPR)


def _sub_val(parent, name):
    """取 parent/<name> 的 w:val。"""
    if parent is None:
        return None
    node = parent.find(qn(name))
    if node is None:
        return None
    return node.get(_VAL)


def _on_off(node):
    """w:b / w:u 这类开关型属性：元素存在且 w:val 不是 0/false/none 视为开启。"""
    if node is None:
        return None
    val = node.get(_VAL)
    if val is None:
        return True
    return str(val).lower() not in ("0", "false", "off", "none")


def _jc_of(p):
    return _sub_val(_pPr(p), "w:jc")


def _int_or_none(value):
    try:
        return int(round(float(value)))
    except (TypeError, ValueError):
        return None


def _char_width_em(ch):
    """字符相对字号的横向宽度：全角 1.0 em，半角 0.5 em。
    与 Word 实测一致：宋体下 12 磅的半角空格/下划线字符宽 6 磅。"""
    if ch in ("\t",):
        return 2.0
    if FULLWIDTH_RE.match(ch):
        return 1.0
    return 0.5


def _text_width_pt(text, size_pt):
    return sum(_char_width_em(ch) for ch in (text or "")) * float(size_pt)


# ========== 样式 / 字体 / 字号解析 ==========

def _build_ctx(doc):
    """收集全文档共用的解析上下文（样式表、docDefaults、主题字体、节属性等）。"""
    ctx = {"doc": doc, "styles": {}, "rdefault": None, "pdefault": None, "theme": {}}
    try:
        for st in doc.styles.element.findall(qn("w:style")):
            sid = st.get(qn("w:styleId"))
            if sid:
                ctx["styles"][sid] = st
        dd = doc.styles.element.find(qn("w:docDefaults"))
        if dd is not None:
            rd = dd.find(qn("w:rPrDefault"))
            pd = dd.find(qn("w:pPrDefault"))
            ctx["rdefault"] = None if rd is None else rd.find(_RPR)
            ctx["pdefault"] = None if pd is None else pd.find(_PPR)
    except Exception:
        pass
    ctx["theme"] = _load_theme_fonts(doc)
    return ctx


def _load_theme_fonts(doc):
    """读取 theme1.xml 的 majorFont/minorFont，供 asciiTheme/eastAsiaTheme 回溯。"""
    out = {}
    try:
        from lxml import etree
        for part in doc.part.package.iter_parts():
            if not str(part.partname).endswith("theme1.xml"):
                continue
            root = etree.fromstring(part.blob)
            for scheme, key in (("majorFont", "major"), ("minorFont", "minor")):
                node = root.find(".//{%s}%s" % (NS_A, scheme))
                if node is None:
                    continue
                latin = node.find("{%s}latin" % NS_A)
                if latin is not None:
                    out[key + "_latin"] = latin.get("typeface")
                for font in node.findall("{%s}font" % NS_A):
                    if font.get("script") in ("Hans", "Hant"):
                        out.setdefault(key + "_ea", font.get("typeface"))
            break
    except Exception:
        pass
    return out


def _theme_font(ctx, theme_val, is_ea):
    if not theme_val:
        return None
    which = "major" if str(theme_val).startswith("major") else "minor"
    return ctx["theme"].get(which + ("_ea" if is_ea else "_latin"))


def _style_chain(ctx, style_id, depth=0):
    """沿 w:basedOn 逐级向上返回样式元素列表。"""
    out = []
    while style_id and depth < 12:
        st = ctx["styles"].get(style_id)
        if st is None:
            break
        out.append(st)
        style_id = _sub_val(st, "w:basedOn")
        depth += 1
    return out


def _para_of_run(run):
    node = run.getparent()
    while node is not None and node.tag != _P:
        node = node.getparent()
    return node


def _pstyle_of(p):
    return _sub_val(_pPr(p), "w:pStyle") if p is not None else None


def _rpr_chain(ctx, run):
    """run 直接格式 -> w:rStyle 链 -> 段落 pStyle 链 -> docDefaults，与 Word 的继承顺序一致。"""
    chain = []
    rpr = _rPr(run)
    if rpr is not None:
        chain.append(rpr)
    for st in _style_chain(ctx, _sub_val(rpr, "w:rStyle")):
        node = st.find(_RPR)
        if node is not None:
            chain.append(node)
    for st in _style_chain(ctx, _pstyle_of(_para_of_run(run))):
        node = st.find(_RPR)
        if node is not None:
            chain.append(node)
    if ctx["rdefault"] is not None:
        chain.append(ctx["rdefault"])
    return chain


def _resolve_effective_eastasia(ctx, run):
    for rpr in _rpr_chain(ctx, run):
        fonts = rpr.find(qn("w:rFonts"))
        if fonts is None:
            continue
        val = fonts.get(qn("w:eastAsia")) or _theme_font(ctx, fonts.get(qn("w:eastAsiaTheme")), True)
        if val:
            return val
    return None


def _resolve_effective_ascii(ctx, run):
    for rpr in _rpr_chain(ctx, run):
        fonts = rpr.find(qn("w:rFonts"))
        if fonts is None:
            continue
        val = (fonts.get(qn("w:ascii"))
               or _theme_font(ctx, fonts.get(qn("w:asciiTheme")), False)
               or fonts.get(qn("w:hAnsi"))
               or _theme_font(ctx, fonts.get(qn("w:hAnsiTheme")), False))
        if val:
            return val
    return None


def _run_size_pt(ctx, run, default=10.5):
    for rpr in _rpr_chain(ctx, run):
        val = _sub_val(rpr, "w:sz")
        num = _int_or_none(val)
        if num:
            return num / 2.0
    return default


def _run_bold(ctx, run):
    for rpr in _rpr_chain(ctx, run):
        node = rpr.find(qn("w:b"))
        if node is not None:
            return _on_off(node)
    return False


def _run_underlined(ctx, run):
    for rpr in _rpr_chain(ctx, run):
        node = rpr.find(qn("w:u"))
        if node is not None:
            return _on_off(node)
    return False


def _is_songti(name):
    return bool(name) and any(name.strip().lower() == f.lower() for f in STRICT_SONGTI)


def _is_heiti(name):
    return bool(name) and any(name.strip().lower() == f.lower() for f in STRICT_HEITI)


def _is_times(name):
    return bool(name) and name.strip().lower().replace(" ", "") in [
        t.replace(" ", "") for t in TIMES]


def _spacing_of(ctx, p):
    """解析段落的 before/after/line/lineRule（段落直接属性 -> pStyle 链 -> docDefaults）。"""
    nodes = []
    pPr = _pPr(p)
    if pPr is not None:
        nodes.append(pPr)
    for st in _style_chain(ctx, _pstyle_of(p)):
        node = st.find(_PPR)
        if node is not None:
            nodes.append(node)
    if ctx["pdefault"] is not None:
        nodes.append(ctx["pdefault"])
    out = {"before": None, "after": None, "line": None, "lineRule": None}
    for node in nodes:
        sp = node.find(qn("w:spacing"))
        if sp is None:
            continue
        for key, attr in (("before", "w:before"), ("after", "w:after"),
                          ("line", "w:line"), ("lineRule", "w:lineRule")):
            if out[key] is None:
                out[key] = sp.get(qn(attr))
    return out


def _ind_of(ctx, p):
    """解析段落缩进（twips）：left / right / firstLine / hanging。"""
    nodes = []
    pPr = _pPr(p)
    if pPr is not None:
        nodes.append(pPr)
    for st in _style_chain(ctx, _pstyle_of(p)):
        node = st.find(_PPR)
        if node is not None:
            nodes.append(node)
    out = {"left": None, "right": None, "firstLine": None, "hanging": None}
    for node in nodes:
        ind = node.find(qn("w:ind"))
        if ind is None:
            continue
        for key, attr in (("left", "w:left"), ("right", "w:right"),
                          ("firstLine", "w:firstLine"), ("hanging", "w:hanging")):
            if out[key] is None:
                out[key] = _int_or_none(ind.get(qn(attr)))
    return out


# ========== 结构定位 ==========

def _split_cover_body(kids):
    """按分节符 / 分页符把 body 顶层块切成"封面区"和"试题正文区"。"""
    for i, el in enumerate(kids):
        if el.tag == _P and el.find(".//" + _SECTPR) is not None:
            sect = el.find(".//" + _SECTPR)
            stype = _sub_val(sect, "w:type")
            return kids[:i + 1], kids[i + 1:], "sectPr", stype
    for i, el in enumerate(kids):
        if el.tag == _P and any(b.get(_TYPE) == "page" for b in el.iter(_BR)):
            return kids[:i + 1], kids[i + 1:], "pageBreak", None
    return [], list(kids), "none", None


def _body_paragraphs(els, with_table=False):
    """正文顶层段落（可选包含表格内段落）。"""
    out = []
    for el in els:
        if el.tag == _P:
            out.append(el)
        elif el.tag == _TBL and with_table:
            out.extend(_paras_of(el))
    return out


def _tables_of(els):
    out = []
    for el in els:
        if el.tag == _TBL:
            out.append(el)
        else:
            out.extend(n for n in _walk(el) if n.tag == _TBL)
    return out


def _table_grid(tbl):
    rows = []
    for tr in tbl.findall(_TR):
        rows.append([_norm(_text_of(tc)) for tc in tr.findall(_TC)])
    return rows


def _is_heading_para(text):
    """主题标题 / 5 处分部标题 / 参考答案标题：这些不计入"正文文字"的字体与间距判定。"""
    stripped = _norm(text)
    if not stripped:
        return False
    if stripped == TITLE_TEXT:
        return True
    if ANSWER_TITLE in stripped:
        return True
    for head in SECTION_HEADS:
        if stripped.startswith(head):
            return True
    return False


def _drawings_of(p):
    """段落内的图片（wp:inline / wp:anchor），返回 (drawing 节点, blip 的 r:embed)。"""
    out = []
    for node in _walk(p):
        if node.tag in ("{%s}inline" % NS_WP, "{%s}anchor" % NS_WP):
            embed = None
            for blip in node.iter("{%s}blip" % NS_A):
                embed = blip.get(_EMBED)
                if embed:
                    break
            out.append((node, embed))
    return out


def _extent_of(node):
    ext = node.find("{%s}extent" % NS_WP)
    if ext is None:
        return None
    cx = _int_or_none(ext.get("cx"))
    cy = _int_or_none(ext.get("cy"))
    if not cx or not cy:
        return None
    return (cx, cy)


def _image_pixels(blob):
    """从图片字节流读取像素尺寸（PNG 的 IHDR / JPEG 的 SOFn），避免依赖 PIL。"""
    try:
        if blob[:8] == b"\x89PNG\r\n\x1a\n":
            width = int.from_bytes(blob[16:20], "big")
            height = int.from_bytes(blob[20:24], "big")
            return (width, height)
        if blob[:2] == b"\xff\xd8":
            i = 2
            while i < len(blob) - 9:
                if blob[i] != 0xFF:
                    i += 1
                    continue
                marker = blob[i + 1]
                seg = int.from_bytes(blob[i + 2:i + 4], "big")
                if 0xC0 <= marker <= 0xCF and marker not in (0xC4, 0xC8, 0xCC):
                    height = int.from_bytes(blob[i + 5:i + 7], "big")
                    width = int.from_bytes(blob[i + 7:i + 9], "big")
                    return (width, height)
                i += 2 + seg
    except Exception:
        pass
    return None


def _collect_figures(doc, body_els):
    """按文档顺序收集正文中的示意图，并记录每张图所挂的题号、所在段落及像素/显示尺寸。"""
    figs = []
    last_num = None
    for el in body_els:
        if el.tag != _P:
            continue
        text = _text_of(el)
        match = QNUM_RE.match(_norm(text))
        if match:
            last_num = int(match.group(1))
        for node, embed in _drawings_of(el):
            blob = None
            if embed:
                try:
                    blob = doc.part.related_parts[embed].blob
                except Exception:
                    blob = None
            figs.append({
                "para": el,
                "anchor": last_num,
                "extent": _extent_of(node),
                "pixels": _image_pixels(blob) if blob else None,
                "bytes": len(blob) if blob else 0,
                "cropped": node.find(".//{%s}srcRect" % NS_A) is not None,
                "text": _norm(text),
            })
    return figs


def _footer_roots(doc):
    """直接从包内读取所有 footerN.xml，避免 python-docx 对文本框页脚的解析盲区。"""
    out = []
    try:
        from lxml import etree
        for part in doc.part.package.iter_parts():
            name = str(part.partname)
            if re.search(r"/footer\d*\.xml$", name):
                out.append((name, etree.fromstring(part.blob)))
    except Exception:
        pass
    return sorted(out, key=lambda item: item[0])


def _header_roots(doc):
    out = []
    try:
        from lxml import etree
        for part in doc.part.package.iter_parts():
            name = str(part.partname)
            if re.search(r"/header\d*\.xml$", name):
                out.append((name, etree.fromstring(part.blob)))
    except Exception:
        pass
    return sorted(out, key=lambda item: item[0])


def _sect_prs(doc):
    """按文档顺序返回所有 sectPr（末节的 sectPr 位于 body 末尾）。"""
    out = []
    body = doc.element.body
    for el in body:
        if el.tag == _P:
            sect = el.find(".//" + _SECTPR)
            if sect is not None:
                out.append(sect)
    tail = body.find(_SECTPR)
    if tail is not None:
        out.append(tail)
    return out


def _field_instrs(root):
    """页眉/页脚里的域代码（w:instrText 与 w:fldSimple 的 w:instr）。"""
    out = []
    for node in _walk(root):
        if node.tag == qn("w:instrText"):
            out.append(node.text or "")
        elif node.tag == qn("w:fldSimple"):
            out.append(node.get(qn("w:instr")) or "")
    return out


# ========== 文档打开 ==========

def _open_word_document(file_path):
    return Document(file_path)


def _sect_geom(sect):
    """节的纸张与页边距（单位 cm）。"""
    out = {"w": None, "h": None, "top": None, "bottom": None, "left": None, "right": None,
           "orient": None}
    size = sect.find(qn("w:pgSz"))
    if size is not None:
        w = _int_or_none(size.get(qn("w:w")))
        h = _int_or_none(size.get(qn("w:h")))
        out["w"] = None if w is None else w / TWIPS_PER_CM
        out["h"] = None if h is None else h / TWIPS_PER_CM
        out["orient"] = size.get(qn("w:orient")) or "portrait"
    mar = sect.find(qn("w:pgMar"))
    if mar is not None:
        for key in ("top", "bottom", "left", "right"):
            val = _int_or_none(mar.get(qn("w:" + key)))
            out[key] = None if val is None else val / TWIPS_PER_CM
    return out


def _anchor_center_offset_cm(node, geom):
    """图形（wp:anchor）中心相对页面水平中线的偏移量，单位 cm；无法判定返回 None。"""
    if geom.get("w") is None:
        return None
    pos_h = node.find("{%s}positionH" % NS_WP)
    ext = _extent_of(node)
    if pos_h is None or ext is None:
        return None
    align = pos_h.find("{%s}align" % NS_WP)
    if align is not None and (align.text or "").strip() == "center":
        return 0.0
    off = pos_h.find("{%s}posOffset" % NS_WP)
    if off is None:
        return None
    x_cm = (_int_or_none(off.text) or 0) / EMU_PER_CM
    if pos_h.get("relativeFrom") not in ("page",):
        x_cm += geom.get("left") or 0.0
    center = x_cm + ext[0] / EMU_PER_CM / 2.0
    return abs(center - geom["w"] / 2.0)


def _is_fake_vertical(p):
    """逐字换行拼出的伪竖排：同段 ≥3 个换行且每行只有 1 个字符。"""
    lines = [ln for ln in _text_of(p).split("\n") if ln.strip()]
    if len(lines) < 3:
        return False
    return all(len(ln.strip()) == 1 for ln in lines)


def _cover_title(ctx, cover_els, geom):
    """定位封面主标题，并给出竖排方式、字号、加粗、水平居中偏移等信息。"""
    info = {"vertical": False, "how": "", "para": None, "runs": [], "text": "",
            "size": 0.0, "bold": False, "offset_cm": None, "fake_vertical": False,
            "drawing": None}
    vert_para = None
    for el in cover_els:
        if el.tag != _P:
            continue
        for node in _walk(el):
            tag = node.tag if isinstance(node.tag, str) else ""
            marker = ""
            if tag.endswith("}bodyPr") and node.get("vert") in ("eaVert", "vert", "mongolianVert"):
                marker = "vert=%s" % node.get("vert")
            elif tag.endswith("}textbox") and "vertical-ideographic" in (node.get("style") or ""):
                marker = "layout-flow:vertical-ideographic"
            if marker:
                text = _norm(_text_of(el))
                if len(text) >= 4:
                    vert_para = el
                    info["vertical"] = True
                    info["how"] = marker
                    break
        if vert_para is not None:
            break
    if vert_para is None:
        best = None
        for el in cover_els:
            if el.tag != _P:
                continue
            text = _norm(_text_of(el))
            if not text or any(re.search(pat, text) for _, pat in COVER_FIELDS):
                continue
            size = max([_run_size_pt(ctx, r) for r in _runs_of(el)] or [0.0])
            if best is None or size > best[0]:
                best = (size, el)
        vert_para = None if best is None else best[1]
        if vert_para is not None:
            info["fake_vertical"] = _is_fake_vertical(vert_para)
    if vert_para is None:
        return info
    info["para"] = vert_para
    runs = [r for r in _runs_of(vert_para) if (_text_of(r) or "").strip()]
    info["runs"] = runs
    info["text"] = _norm(_text_of(vert_para))
    info["size"] = max([_run_size_pt(ctx, r) for r in runs] or [0.0])
    top = [r for r in runs if abs(_run_size_pt(ctx, r) - info["size"]) < 0.01]
    info["bold"] = bool(top) and all(_run_bold(ctx, r) for r in top)
    drawings = _drawings_of(vert_para)
    if drawings:
        info["drawing"] = drawings[0][0]
        info["offset_cm"] = _anchor_center_offset_cm(drawings[0][0], geom)
    if info["offset_cm"] is None and _jc_of(vert_para) == "center":
        info["offset_cm"] = 0.0
    return info


def _cover_field_paras(cover_els):
    """按"报送单位/学科/命题人/命题人联系电话"定位封面四项填写栏所在段落。"""
    found = {}
    for el in cover_els:
        if el.tag != _P:
            continue
        text = _norm(_text_of(el))
        if not text:
            continue
        for label, pattern in COVER_FIELDS:
            if label in found:
                continue
            if re.search(pattern, text):
                found[label] = el
                break
    return found


def _underline_length_pt(ctx, p):
    """段落内"下划线格式"填空横线的长度（磅）：w:u 下划线运行的宽度 + 制表位前导线长度。"""
    total = 0.0
    for run in _runs_of(p):
        if not _run_underlined(ctx, run):
            continue
        total += _text_width_pt(_text_of(run), _run_size_pt(ctx, run))
    pPr = _pPr(p)
    if pPr is not None:
        tabs = pPr.find(qn("w:tabs"))
        if tabs is not None:
            plain = 0.0
            for run in _runs_of(p):
                if not _run_underlined(ctx, run):
                    plain += _text_width_pt(_text_of(run), _run_size_pt(ctx, run))
            for tab in tabs.findall(qn("w:tab")):
                if (tab.get(qn("w:leader")) or "none") == "none":
                    continue
                pos = _int_or_none(tab.get(qn("w:pos"))) or 0
                total += max(0.0, pos / 20.0 - plain)
    return total


def _typed_underscore_run(p):
    """段落里是否用连续键入的"_"字符充当横线。"""
    return bool(re.search(r"_{5,}", _text_of(p)))


def _effective_left_twips(ctx, p):
    ind = _ind_of(ctx, p)
    left = ind["left"] or 0
    first = ind["firstLine"] or 0
    hang = ind["hanging"] or 0
    return left + first - hang


def _text_area_cm(geom):
    if geom.get("w") is None:
        return None
    return geom["w"] - (geom.get("left") or 0.0) - (geom.get("right") or 0.0)


# ========== 对外入口 ==========

def _one_line(text, limit=200):
    """把异常信息压成单行摘要，供 dim1_reason 使用。"""
    flat = re.sub(r"\s+", " ", str(text or "")).strip()
    return flat if len(flat) <= limit else flat[:limit] + "…"


def _result_not_found(message):
    """找不到目录 / 找不到待评文件：未进入评测，dim1_pass 留空，max_score 记 0。"""
    return {
        "id": SCRIPT_ID,
        "file_name": "",
        "status": "error",
        "error": message,
        "dim1_pass": "",
        "dim1_reason": "",
        "dim2_items": [],
        "total_score": 0,
        "max_score": 0,
    }


def evaluate(dir_path):
    """入参为目录路径，自动取其中的 .docx 作为待评文件；返回结构化评测结果。"""
    if not os.path.isdir(dir_path):
        return _result_not_found("目录不存在：%s" % dir_path)
    try:
        cands = sorted(f for f in os.listdir(dir_path)
                       if f.lower().endswith(".docx") and not f.startswith("~"))
    except OSError as exc:
        return _result_not_found("目录无法读取：%s（%s）" % (dir_path, exc))
    if not cands:
        return _result_not_found("目录中未找到 .docx 文件：%s" % dir_path)

    result = {
        "id": SCRIPT_ID,
        "file_name": cands[0],
        "status": "ok",
        "error": "",
        "dim1_pass": True,
        "dim1_reason": "",
        "dim2_items": [],
        "total_score": 0,
        "max_score": MAX_SCORE_TOTAL,
    }
    try:
        return _evaluate_impl(os.path.join(dir_path, cands[0]), result)
    except Exception as exc:
        message = "%s: %s" % (type(exc).__name__, exc)
        result["status"] = "error"
        result["error"] = message
        result["traceback"] = traceback.format_exc()
        # 文件已找到但打不开/解析失败：维度1 判为不通过，dim1_reason 给出错误摘要
        result["dim1_pass"] = False
        result["dim1_reason"] = "文件无法正常打开或解析：%s" % _one_line(message)
        result["dim2_items"] = []
        result["total_score"] = 0
        return result


def _evaluate_impl(file_path, result):
    doc = _open_word_document(file_path)
    ctx = _build_ctx(doc)
    kids = list(doc.element.body)
    cover_els, body_els, split_mode, split_type = _split_cover_body(kids)

    cover_text = "\n".join(_text_of(el) for el in cover_els)
    body_text = "\n".join(_text_of(el) for el in body_els)
    body_paras = _body_paragraphs(body_els)
    body_paras_all = _body_paragraphs(body_els, with_table=True)
    tables = _tables_of(body_els)
    figs = _collect_figures(doc, body_els)
    sects = _sect_prs(doc)
    geom_list = [_sect_geom(s) for s in sects]
    geom = geom_list[-1] if geom_list else {}
    footers = _footer_roots(doc)
    headers = _header_roots(doc)
    results = []


    # ========== 维度1：可用与可修改性 ==========
    # ① 文件能被 Document() 正常打开（打不开会在 evaluate 的 except 中记为 error）。
    # ② 试题内容必须是可编辑文本：正文可见文字（含表格）不足 800 字，视为整页位图粘贴。
    # ③ 任一整页被单张图片占满（图片显示面积 ≥ 可排版面积 90%），视为不可编辑。
    editable_chars = len(_squeeze(cover_text + body_text))
    area_cm2 = None
    if geom.get("w") and geom.get("h"):
        area_cm2 = ((geom["w"] - (geom.get("left") or 0) - (geom.get("right") or 0))
                    * (geom["h"] - (geom.get("top") or 0) - (geom.get("bottom") or 0)))
    full_page_img = False
    for fig in figs:
        if not fig["extent"] or not area_cm2:
            continue
        fig_cm2 = (fig["extent"][0] / EMU_PER_CM) * (fig["extent"][1] / EMU_PER_CM)
        if fig_cm2 >= area_cm2 * 0.9:
            full_page_img = True
    if editable_chars < 800:
        result["dim1_pass"] = False
        result["dim1_reason"] = ("可编辑文字仅 %d 字，试题内容疑似以位图/扫描图整页粘贴，"
                                "无法编辑修改" % editable_chars)
    elif full_page_img:
        result["dim1_pass"] = False
        result["dim1_reason"] = "存在占满整页可排版区域的单张图片，该页内容不可单独点选编辑"
    else:
        # 维度1 通过时不输出理由（dim1_reason 仅在未通过时给出摘要）
        result["dim1_pass"] = True
        result["dim1_reason"] = ""
    if not result["dim1_pass"]:
        result["dim2_items"] = []
        result["total_score"] = 0
        return result

    # ========== 维度2·加分项（指令下达的任务） ==========

    # ---------- 01 +5 封面单独成页、试题正文另起一页 ----------
    # ① 以分节符/分页符切出封面区，节类型不能是 continuous（否则不换页）；
    # ② 封面区必须有可见内容；③ 封面区不得混排试题正文（出现注意事项/第一部分/满分说明）。
    cover_visible = len(_squeeze(cover_text)) > 0
    mixed = any(key in _squeeze(cover_text) for key in ("注意事项", "第一部分", "本试卷满分"))
    if (split_mode in ("sectPr", "pageBreak") and cover_visible and not mixed
            and (split_type or "nextPage") != "continuous"):
        results.append("+5：文档首页为独立封面页（以%s分隔），封面内容单独成页、试题正文另起一页开始"
                       % ("分节符" if split_mode == "sectPr" else "分页符"))
    else:
        results.append("+0：未形成独立封面页（分隔方式=%s，封面有可见内容=%s，与正文混排=%s）"
                       % (split_mode, cover_visible, mixed))

    # ---------- 02 +5 封面主标题真正竖排 ----------
    # ① 竖排文本框 wps/a:bodyPr 的 vert="eaVert" 或 VML layout-flow:vertical-ideographic；
    # ② 竖排容器内文字≥4字才算主标题；③ 逐字换行拼出的伪竖排不得分。
    title = _cover_title(ctx, cover_els, geom)
    if title["vertical"]:
        results.append("+5：封面主标题“%s”为真正的竖排文字（%s）" % (title["text"], title["how"]))
    else:
        results.append("+0：封面主标题不是真正的竖排文字（%s）"
                       % ("以逐字换行拼出的伪竖排" if title["fake_vertical"] else "未见竖排文字方向属性"))

    # ---------- 03 +1 封面主标题加粗、≥18磅、水平居中 ----------
    # ① 取标题中最大字号的运行判定加粗；② 居中允许偏离页面水平中线 2cm 内
    #    （文本框按 wp:align=center 或 posOffset+extent 折算中心）；③ 标题内不得有≥2连续空格。
    center_off = title["offset_cm"]
    center_ok = center_off is not None and center_off <= 2.0
    gap_ok = not re.search(r"[ 　]{2,}", title["text"])
    if title["bold"] and title["size"] >= 17.95 and center_ok and gap_ok:
        results.append("+1：封面主标题加粗、字号 %.1f 磅（≥18磅）且水平居中（偏离中线 %.2fcm），字间紧凑"
                       % (title["size"], center_off or 0.0))
    else:
        results.append("+0：封面主标题格式不满足加粗+≥18磅+水平居中（加粗=%s，字号=%.1f磅，"
                       "偏离中线=%s，字间无大段空隙=%s）"
                       % (title["bold"], title["size"],
                          "无法判定" if center_off is None else "%.2fcm" % center_off, gap_ok))


    # ---------- 04 +1 封面四项填写项齐全 ----------
    # ① 按标签正则在封面段落中定位；②“命题人”与“命题人联系电话”须为两个独立填写项；
    # ③ 缺任一项不得分。
    fields = _cover_field_paras(cover_els)
    missing_fields = [label for label, _ in COVER_FIELDS if label not in fields]
    if not missing_fields:
        results.append("+1：封面设有“报送单位”“学科”“命题人”“命题人联系电话”四项填写项")
    else:
        results.append("+0：封面四项填写项不全，缺少 %s" % "、".join(missing_fields))

    # ---------- 05 +5 四项填空横线为下划线格式且长度≥5cm ----------
    # ① 只统计带 w:u 下划线的运行宽度与制表位前导线长度（连续键入的“_”字符不计）；
    # ② 字符宽度按半角 0.5em / 全角 1.0em 折算，与 Word 实测一致（12磅半角空格=6磅）；
    # ③ 四项横线须均≥5cm。
    line_lens = {}
    for label, _ in COVER_FIELDS:
        para = fields.get(label)
        line_lens[label] = 0.0 if para is None else _underline_length_pt(ctx, para) / PT_PER_CM
    typed_only = [label for label, _ in COVER_FIELDS
                  if fields.get(label) is not None and line_lens[label] <= 0.01
                  and _typed_underscore_run(fields[label])]
    len_desc = "、".join("%s %.2fcm" % (label, line_lens[label]) for label, _ in COVER_FIELDS)
    if not missing_fields and all(line_lens[label] >= UNDERLINE_MIN_CM for label, _ in COVER_FIELDS):
        results.append("+5：封面四项填写项后均带下划线格式的填空横线且长度≥5cm（%s）" % len_desc)
    else:
        results.append("+0：封面填空横线不满足下划线格式且≥5cm（%s%s）"
                       % (len_desc,
                          "；其中 %s 为连续键入的下划线字符" % "、".join(typed_only) if typed_only else ""))

    # ---------- 06 +1 四项位于封面下半部、左端整体对齐 ----------
    # ① 左缩进（left+firstLine-hanging）须落在 1700-2600 twips（约3-4cm）且四项互差≤120 twips；
    # ② 位置按文档序判定：四项须排在封面主标题之后（不渲染，故不测绝对高度）。
    lefts = [_effective_left_twips(ctx, fields[label]) for label, _ in COVER_FIELDS
             if label in fields]
    ind_ok = (len(lefts) == 4 and all(1700 <= v <= 2600 for v in lefts)
              and (max(lefts) - min(lefts)) <= 120)
    lower_ok = None
    if title["para"] is not None:
        order = {id(el): i for i, el in enumerate(cover_els)}
        lower_ok = all(order.get(id(fields[label]), -1) > order.get(id(title["para"]), 1 << 30)
                       for label, _ in COVER_FIELDS if label in fields)
    if ind_ok and lower_ok:
        results.append("+1：封面四项填写项自上而下排在封面下半部分，左端整体对齐（左缩进 %s twips）"
                       % "/".join(str(v) for v in lefts))
    else:
        results.append("+0：封面四项填写项排布不满足下半部分+左端对齐（左缩进=%s，位于下半部=%s）"
                       % (lefts or "未定位", lower_ok))


    # ---------- 07 +1 卷内主题标题 黑体三号居中 ----------
    # ① 取正文区第一个只含“初中化学模拟试题”的段落；② eastAsia 须解析为黑体；
    # ③ 字号 16 磅（w:sz=32，三号）且段落 w:jc=center。
    body_title = None
    for para in body_paras:
        text = _norm(_text_of(para))
        if text and TITLE_TEXT in text and len(text) <= len(TITLE_TEXT) + 6:
            body_title = para
            break
    if body_title is None:
        results.append("+0：正文区未找到主题标题“%s”" % TITLE_TEXT)
    else:
        t_runs = [r for r in _runs_of(body_title) if (_text_of(r) or "").strip()]
        t_fonts = [_resolve_effective_eastasia(ctx, r) for r in t_runs]
        t_sizes = [_run_size_pt(ctx, r) for r in t_runs]
        t_jc = _jc_of(body_title)
        if (t_runs and all(_is_heiti(f) for f in t_fonts)
                and all(15.5 <= s <= 16.5 for s in t_sizes) and t_jc == "center"):
            results.append("+1：卷内主题标题“%s”为黑体三号（%.1f磅）居中" % (TITLE_TEXT, t_sizes[0]))
        else:
            results.append("+0：卷内主题标题不满足黑体三号居中（字体=%s，字号=%s磅，对齐=%s）"
                           % ("/".join(str(f) for f in set(t_fonts)),
                              "/".join("%.1f" % s for s in set(t_sizes)), t_jc))

    # ---------- 08 +5 试题正文中文 宋体小四 ----------
    # ① 判定范围＝正文区所有段落（含表格内文字），排除主题标题、5处分部标题与参考答案标题；
    # ② 只看含中日韩字符的运行，字体按 run→rStyle→pStyle→docDefaults→主题 逐级解析；
    # ③ 要求 eastAsia ∈ 宋体族 且字号 12 磅（w:sz=24，小四）。
    song_bad = []
    for para in body_paras_all:
        if _is_heading_para(_text_of(para)):
            continue
        for run in _runs_of(para):
            run_text = _text_of(run)
            if not CJK_RE.search(run_text or ""):
                continue
            name = _resolve_effective_eastasia(ctx, run)
            size = _run_size_pt(ctx, run)
            if not _is_songti(name) or not (11.5 <= size <= 12.5):
                song_bad.append("“%s”(%s/%.1f磅)" % (_norm(run_text)[:8], name, size))
    if not song_bad:
        results.append("+5：试题正文中文字体全部为宋体小四（12磅），覆盖注意事项、题干、选项、"
                       "填空简答、实验探究、计算题及表格内文字")
    else:
        results.append("+0：试题正文中文字体存在 %d 处不是宋体小四，例如 %s"
                       % (len(song_bad), "、".join(song_bad[:3])))


    # ---------- 09 +3 试题正文西文与数字 Times New Roman 小四 ----------
    # ① 判定范围同上；② 只看含英文字母/阿拉伯数字的运行；
    # ③ 要求 ascii（缺省回退 hAnsi/主题）解析为 Times New Roman 且字号 12 磅。
    tnr_bad = []
    for para in body_paras_all:
        if _is_heading_para(_text_of(para)):
            continue
        for run in _runs_of(para):
            run_text = _text_of(run)
            if not LATIN_RE.search(run_text or ""):
                continue
            name = _resolve_effective_ascii(ctx, run)
            size = _run_size_pt(ctx, run)
            if not _is_times(name) or not (11.5 <= size <= 12.5):
                tnr_bad.append("“%s”(%s/%.1f磅)" % (_norm(run_text)[:8], name, size))
    if not tnr_bad:
        results.append("+3：试题正文中的英文字母与阿拉伯数字字体为 Times New Roman 小四（12磅）")
    else:
        results.append("+0：正文西文/数字字体存在 %d 处不是 Times New Roman 小四，例如 %s"
                       % (len(tnr_bad), "、".join(tnr_bad[:3])))

    # ---------- 10 +1 页脚页码文本“化学试题 第x页 共x页” ----------
    # ① 页脚文字含文本框内文字；② 去空白后须能匹配“化学试题第…页共…页”；
    # ③ 三段（前缀、第…页、共…页）缺一不可。
    footer_texts = [_squeeze(_text_of(root)) for _, root in footers]
    footer_all = "".join(footer_texts)
    pat_full = re.compile(r"化学试题第[\d x]{0,6}页共[\d x]{0,6}页")
    if pat_full.search(footer_all):
        results.append("+1：页脚页码为“化学试题 第x页 共x页”格式，同时含“化学试题”前缀与“第…页”“共…页”两段")
    else:
        results.append("+0：页脚页码不满足“化学试题 第x页 共x页”格式（页脚实际内容：%s）"
                       % (footer_all[:40] or "空"))

    # ---------- 11 +3 当前页序号由 PAGE 域生成 ----------
    # ① 收集页脚的 w:instrText 与 w:fldSimple/@w:instr；② 用 \bPAGE\b 排除 NUMPAGES 的误命中；
    # ③ 逐页手键固定数字则不得分。
    instrs = []
    for _, root in footers:
        instrs.extend(_field_instrs(root))
    has_page_field = any(re.search(r"\bPAGE\b", s or "") for s in instrs)
    if has_page_field:
        results.append("+3：页码中的当前页序号由 PAGE 域自动生成（域代码：%s）"
                       % "、".join(sorted(set(s.strip() for s in instrs if s.strip()))[:3]))
    else:
        results.append("+0：页码中的当前页序号不是 PAGE 域自动生成（未见 PAGE 域代码）")


    # ---------- 12 +3 “共x页”由 NUMPAGES/SECTIONPAGES 域自动计算、分节起编正确 ----------
    # 真实总页数只有排版引擎才知道，OOXML 静态解析测不到；因此本项判据明确定义为
    # “页码总数由域自动计算、且分节编号设置正确”，三条须同时满足：
    # ① 页脚里必须出现 NUMPAGES 或 SECTIONPAGES 域，而不是手敲的固定数字；
    # ② 该域须与当前页序号的 PAGE 域出现在同一个页脚部件里（成对使用，构成“第x页 共y页”）；
    # ③ 封面须单独成节，且正文节写有 w:pgNumType/@w:start（页码在正文节重新起编），
    #    这样封面不参与页码计数。
    total_field = None
    total_paired = False
    total_codes = []
    for _name, root in footers:
        codes = [s or "" for s in _field_instrs(root)]
        found = next((re.search(r"\b(NUMPAGES|SECTIONPAGES)\b", s, re.I) for s in codes
                      if re.search(r"\b(NUMPAGES|SECTIONPAGES)\b", s, re.I)), None)
        if found is None:
            continue
        total_field = found.group(1).upper()
        total_codes = sorted(set(s.strip() for s in codes if s.strip()))
        total_paired = any(re.search(r"\bPAGE\b", s, re.I) for s in codes)
        break
    cover_sectioned = len(sects) >= 2
    body_num_type = sects[-1].find(qn("w:pgNumType")) if sects else None
    body_restart = (body_num_type is not None
                    and body_num_type.get(qn("w:start")) is not None)
    if total_field and total_paired and cover_sectioned and body_restart:
        results.append("+3：页码总数由 %s 域自动计算并与 PAGE 域成对写在同一页脚（域代码：%s），"
                       "封面单独成节、正文节自 w:pgNumType/@w:start=%s 重新起编，封面不计入页码"
                       % (total_field, "、".join(total_codes[:3]),
                          body_num_type.get(qn("w:start"))))
    else:
        results.append("+0：页码总数不是由 NUMPAGES/SECTIONPAGES 域自动计算，或分节起编设置不正确"
                       "（总页数域=%s，与 PAGE 域同页脚=%s，封面单独成节=%s，正文节重新起编=%s）"
                       % (total_field or "无（疑为手敲固定数字）", total_paired,
                          cover_sectioned, body_restart))

    # ---------- 13 +5 封面页不显示任何页码 ----------
    # 纯 XML 判据：封面必须单独成节，且该节不引用任何页脚（w:footerReference 缺失）——
    # 节内没有页脚部件，封面上就不可能显示页码；未与正文分节则不得分。
    cover_footer_refs = bool(sects) and bool(sects[0].findall(qn("w:footerReference")))
    cover_no_num = len(sects) >= 2 and not cover_footer_refs
    if cover_no_num:
        results.append("+5：封面页不显示任何页码（封面单独成节且该节不引用页脚）")
    else:
        results.append("+0：封面页仍显示页码或未与正文分节（封面单独成节=%s，首节引用页脚=%s）"
                       % (cover_sectioned, cover_footer_refs))

    # ---------- 14 +3 正文第一页页码从“第1页”开始 ----------
    # 纯 XML 判据：正文节（末节）须显式写 w:pgNumType/@w:start="1"，即页码在正文节
    # 重新从 1 起编；缺 pgNumType 或 start≠1（把封面计入页序）则不得分。
    body_first_ok = None
    if sects:
        num_type = sects[-1].find(qn("w:pgNumType"))
        body_first_ok = num_type is not None and (num_type.get(qn("w:start")) == "1")
    if body_first_ok:
        results.append("+3：试题正文第一页页码从“第1页”开始，未把封面计入页序")
    else:
        results.append("+0：试题正文第一页页码未从“第1页”开始（判定结果=%s）" % body_first_ok)


    # ---------- 15 +1 页脚居中、9-12磅、中文宋体+数字 Times New Roman ----------
    # ① 居中兼容三种写法：段落 w:jc=center、图形 wp:align=center、VML 水平居中；
    # ② 页脚内所有有字运行的字号须在 9-12 磅（小五至小四）；
    # ③ 中文运行 eastAsia∈宋体族、含数字运行 ascii=Times New Roman。
    footer_center = False
    footer_sizes = []
    footer_font_bad = []
    for _, root in footers:
        for para in _paras_of(root):
            if _jc_of(para) == "center":
                footer_center = True
        for node in _walk(root):
            tag = node.tag if isinstance(node.tag, str) else ""
            if tag == "{%s}align" % NS_WP and (node.text or "").strip() == "center":
                footer_center = True
            elif tag.endswith("}shape") and "mso-position-horizontal:center" in (node.get("style") or ""):
                footer_center = True
        for run in _runs_of(root):
            run_text = _text_of(run)
            if not (run_text or "").strip():
                continue
            footer_sizes.append(_run_size_pt(ctx, run))
            if CJK_RE.search(run_text) and not _is_songti(_resolve_effective_eastasia(ctx, run)):
                footer_font_bad.append("中文=%s" % _resolve_effective_eastasia(ctx, run))
            if LATIN_RE.search(run_text) and not _is_times(_resolve_effective_ascii(ctx, run)):
                footer_font_bad.append("数字=%s" % _resolve_effective_ascii(ctx, run))
    size_ok = bool(footer_sizes) and all(8.95 <= s <= 12.05 for s in footer_sizes)
    if footer_center and size_ok and not footer_font_bad:
        results.append("+1：页脚页码水平居中，字号 %.1f 磅，中文宋体、数字 Times New Roman"
                       % (footer_sizes[0] if footer_sizes else 0.0))
    else:
        results.append("+0：页脚页码不满足居中+9-12磅+宋体/Times New Roman（居中=%s，字号=%s，字体问题=%s）"
                       % (footer_center,
                          "/".join(sorted({"%.1f" % s for s in footer_sizes})) or "无",
                          "、".join(sorted(set(footer_font_bad))[:3]) or "无"))

    # ---------- 16 +5 A4 纵向 + 上下左右页边距 2.5cm ----------
    # ① 逐节校验 w:pgSz（20.9-21.1cm × 29.6-29.8cm，即 11906×16838 twips）；
    # ② 四边 w:pgMar 均为 2.5cm（1417±60 twips）；③ 任一节不符即不得分。
    geom_bad = []
    for idx, item in enumerate(geom_list):
        ok = (item.get("w") is not None and 20.9 <= item["w"] <= 21.1
              and item.get("h") is not None and 29.6 <= item["h"] <= 29.8
              and all(item.get(k) is not None and abs(item[k] - 2.5) <= 0.11
                      for k in ("top", "bottom", "left", "right")))
        if not ok:
            geom_bad.append("第%d节(%.2f×%.2fcm，边距%s)"
                            % (idx + 1, item.get("w") or 0, item.get("h") or 0,
                               "/".join("%.2f" % (item.get(k) or 0)
                                        for k in ("top", "bottom", "left", "right"))))
    if geom_list and not geom_bad:
        results.append("+5：纸张为 A4 纵向（%.2f×%.2fcm），上下左右页边距均为 2.5cm"
                       % (geom_list[0]["w"], geom_list[0]["h"]))
    else:
        results.append("+0：纸张或页边距不满足 A4 纵向+2.5cm（%s）"
                       % ("；".join(geom_bad) or "未解析到节属性"))


    # ---------- 17 +3 正文段间距 0 行、行距 1.5 倍 ----------
    # ① 判定范围＝正文区有字的顶层段落，排除主题标题与分部标题；
    # ② 段前/段后须为 0；③ 行距须为 w:line=360 且 w:lineRule=auto（1.5倍行距）。
    spacing_bad = []
    for para in body_paras:
        text = _text_of(para)
        if not _squeeze(text) or _is_heading_para(text):
            continue
        sp = _spacing_of(ctx, para)
        before = _int_or_none(sp["before"]) or 0
        after = _int_or_none(sp["after"]) or 0
        line = _int_or_none(sp["line"])
        rule = sp["lineRule"]
        if before or after or rule != "auto" or line is None or not (355 <= line <= 365):
            spacing_bad.append("“%s”(before=%s,after=%s,line=%s/%s)"
                               % (_norm(text)[:8], before, after, line, rule))
    if not spacing_bad:
        results.append("+3：试题正文段落段间距为 0 行、行间距为 1.5 倍行距（before=0，after=0，line=360 auto）")
    else:
        results.append("+0：正文有 %d 个段落不满足段间距0行+1.5倍行距，例如 %s"
                       % (len(spacing_bad), "、".join(spacing_bad[:3])))

    # ========== 维度2·减分项（指令未要求而多做的破坏性修改） ==========

    # ---------- 18 -3 第25题表格“用于食品包装袋内作保护气” ----------
    # ① 在全部表格单元格中定位含“食品包装袋”的行；② 该单元格文字须为“用于食品包装袋内作保护气”；
    # ③ 被改写为“做保护气”等即扣分。
    pack_cell = None
    for tbl in tables:
        for row in _table_grid(tbl):
            for cell in row:
                if "食品包装袋" in cell:
                    pack_cell = cell
    if pack_cell is not None and _squeeze(pack_cell) == "用于食品包装袋内作保护气":
        results.append("-0：第25题表格第二行文字保持为“用于食品包装袋内作保护气”")
    else:
        results.append("-3：第25题表格第二行文字不满足“用于食品包装袋内作保护气”（实际为“%s”）"
                       % (pack_cell if pack_cell is not None else "未找到该行"))

    # ---------- 19 -3 6张示意图全部保留且各自留在对应题目下 ----------
    # ① 按文档顺序收集正文示意图；② 每张图记录其上方最近的题号；
    # ③ 题号序列须恰为 5/24/27/28/29/30，出现丢失、错位到其他题目或重复插入即扣分。
    anchors = [fig["anchor"] for fig in figs]
    if anchors == EXPECTED_FIG_ANCHORS:
        results.append("-0：第5、24、27、28、29、30 题下的 6 张示意图全部保留，且各自留在对应题目下")
    else:
        results.append("-3：6 张示意图不满足全部保留且各自留在对应题目下（共 %d 张，所属题号 %s，"
                       "应为 %s）" % (len(figs), anchors, EXPECTED_FIG_ANCHORS))


    # ---------- 20 -3 图2/图3/图5 等比缩放 ----------
    # ① 显示宽高比取 wp:extent；② 与应有宽高比（1.96/2.00/1.61）比对，允许 5% 误差；
    # ③ 同时要求显示宽高比与图片自身像素宽高比一致，横向拉伸/纵向压扁即扣分。
    fig_names = {2: "图2 过滤装置示意图", 3: "图3 实验室制取并收集氧气装置示意图",
                 5: "图5 两种固体物质的溶解度曲线"}
    aspect_bad = []
    for idx in sorted(FIG_ASPECTS):
        target = FIG_ASPECTS[idx]
        if idx >= len(figs):
            aspect_bad.append("%s 缺失" % fig_names[idx])
            continue
        ext = figs[idx]["extent"]
        pix = figs[idx]["pixels"]
        disp = None if not ext else ext[0] / float(ext[1])
        native = None if not pix else pix[0] / float(pix[1])
        if disp is None:
            aspect_bad.append("%s 无显示尺寸" % fig_names[idx])
        elif abs(disp / target - 1.0) > 0.05 or (native and abs(disp / native - 1.0) > 0.05):
            aspect_bad.append("%s 宽高比 %.3f（应约 %.2f）" % (fig_names[idx], disp, target))
    if not aspect_bad:
        results.append("-0：图2/图3/图5 均为等比缩放，宽高比分别约 1.96:1、2.00:1、1.61:1")
    else:
        results.append("-3：图2/图3/图5 不满足等比缩放，出现横向拉伸或纵向压扁变形（%s）"
                       % "；".join(aspect_bad))

    # ---------- 21 -1 图2 独占一个居中段落 ----------
    # ① 取正文第 3 张示意图（图2 过滤装置示意图）所在段落；② 该段不得同时含题干文字；
    # ③ 段落须居中。
    fig2 = figs[2] if len(figs) > 2 else None
    if fig2 is not None and not _squeeze(fig2["text"]) and _jc_of(fig2["para"]) == "center":
        results.append("-0：第27题过滤装置示意图独占一个居中段落")
    else:
        results.append("-1：第27题过滤装置示意图不满足独占一个居中段落（%s）"
                       % ("未定位到该图" if fig2 is None
                          else "同段落含文字“%s”" % fig2["text"][:20] if _squeeze(fig2["text"])
                          else "段落对齐=%s" % _jc_of(fig2["para"])))

    # ---------- 22 -3 正文页面底部出现 >5cm 连续空白 ----------
    # 页底连续空白的高度只能在真实分页之后测量，OOXML 静态解析拿不到内容底边坐标；
    # 按“减分需证据”的口径，本项一律判为未出现（不凭猜测扣分）。
    results.append("-0：试题正文各页底部未出现高度超过 5cm 的连续空白区域（未渲染，按未出现处理）")


    # ---------- 23 -3 第31题 6 行作答横线在单行内完整排布 ----------
    # ① 取正文中整段仅由下划线字符构成的段落；② 按半角 0.5em 折算横线总宽度；
    # ③ 宽度超过该段可用行宽即会折行并在次行残留孤立下划线字符。
    avail_cm = _text_area_cm(geom) or 16.0
    wrap_bad = []
    for para in body_paras:
        text = _text_of(para)
        if not re.fullmatch(r"[_\s]{20,}", text or ""):
            continue
        width_cm = sum(_text_width_pt(_text_of(r), _run_size_pt(ctx, r))
                       for r in _runs_of(para)) / PT_PER_CM
        ind = _ind_of(ctx, para)
        used_cm = ((ind["left"] or 0) + max(ind["firstLine"] or 0, 0)
                   + (ind["right"] or 0)) / TWIPS_PER_CM
        if width_cm > (avail_cm - used_cm) + 0.02:
            wrap_bad.append("%d个下划线宽 %.2fcm＞可用 %.2fcm"
                            % (text.count("_"), width_cm, avail_cm - used_cm))
    if not wrap_bad:
        results.append("-0：第31题“解题过程”下方的 6 行作答横线均在单行内完整排布")
    else:
        results.append("-3：第31题“解题过程”下方的作答横线不满足在单行内完整排布，出现折行并在次行"
                       "残留孤立下划线字符（%d 行，如 %s）" % (len(wrap_bad), wrap_bad[0]))

    # ---------- 24 -1 化学式/相对原子质量/单位内部被插入多余空格 ----------
    # ① 在正文全文（含表格）上匹配 5 类断裂模式；② 命中即说明连写形式被拆开；
    # ③ 匹配片段一并回显，便于人工复核。
    space_hits = []
    for pattern, desc in FORMULA_SPACE_PATTERNS:
        found = re.findall(pattern, body_text)
        if found:
            space_hits.append("%s：%s" % (desc, "、".join(repr(x) for x in found[:2])))
    if not space_hits:
        results.append("-0：化学式、相对原子质量与单位保持“2H₂O”“H₂O”“CO₂↑”“Cl-35.5”“O₂²⁻”的连写形式")
    else:
        results.append("-1：化学式、相对原子质量与单位内部被插入多余空格（%s）" % "；".join(space_hits[:3]))

    # ---------- 25 -1 卷内主题标题加粗 ----------
    # ① 取第 07 条定位到的正文主题标题；② 其全部有字运行须带 w:b；③ 未加粗即扣分。
    title_bold = False
    if body_title is not None:
        t_runs = [r for r in _runs_of(body_title) if (_text_of(r) or "").strip()]
        title_bold = bool(t_runs) and all(_run_bold(ctx, r) for r in t_runs)
    if title_bold:
        results.append("-0：卷内主题标题“%s”保持加粗" % TITLE_TEXT)
    else:
        results.append("-1：卷内主题标题“%s”不满足加粗（w:b 缺失或被置为 0）" % TITLE_TEXT)


    # ---------- 26 -3 5 处分部标题 黑体 13.5 磅加粗 ----------
    # ① 定位“注意事项”“第一部分…”“第二部分…”“第三部分…”“第四部分…”5 个段落；
    # ② 每处均须 eastAsia=黑体、字号 13.5 磅（w:sz=27）、加粗；③ 任一处不符即扣分。
    head_bad = []
    for head in SECTION_HEADS:
        target = None
        for para in body_paras:
            if _norm(_text_of(para)).startswith(head):
                target = para
                break
        if target is None:
            head_bad.append("%s 缺失" % head)
            continue
        h_runs = [r for r in _runs_of(target) if (_text_of(r) or "").strip()]
        fonts = {_resolve_effective_eastasia(ctx, r) for r in h_runs}
        sizes = {_run_size_pt(ctx, r) for r in h_runs}
        bold = bool(h_runs) and all(_run_bold(ctx, r) for r in h_runs)
        if not (all(_is_heiti(f) for f in fonts) and all(13.0 <= s <= 14.0 for s in sizes) and bold):
            head_bad.append("%s(%s/%s磅/加粗=%s)"
                            % (head, "/".join(str(f) for f in fonts),
                               "/".join("%.1f" % s for s in sizes), bold))
    if not head_bad:
        results.append("-0：“注意事项”与四处分部标题均为黑体 13.5 磅加粗")
    else:
        results.append("-3：5 处分部标题不满足黑体 13.5 磅加粗（%s）" % "；".join(head_bad[:3]))

    # ---------- 27 -3 注意事项下方题号分值表 ----------
    # ① 须为 2 行 5 列；② 首行为“题号/一/二/三/四”；③ 次行为“分值/40/20/25/15”。
    grid_tbl = None
    for tbl in tables:
        grid = _table_grid(tbl)
        if len(grid) == 2 and all(len(row) == 5 for row in grid) and grid[0][0].startswith("题号"):
            grid_tbl = grid
            break
    expect_head = ["题号", "一", "二", "三", "四"]
    expect_body = ["分值", "40", "20", "25", "15"]
    if grid_tbl and [_squeeze(c) for c in grid_tbl[0]] == expect_head \
            and [_squeeze(c) for c in grid_tbl[1]] == expect_body:
        results.append("-0：注意事项下方表格保持 2 行 5 列，题号“一、二、三、四”与分值“40、20、25、15”对应")
    else:
        results.append("-3：注意事项下方表格不满足 2 行 5 列结构及题号/分值对应数值（实际 %s）"
                       % (grid_tbl if grid_tbl else "未找到该表"))


    # ---------- 28 -3 第一部分选择题 20 小题、每题 ABCD 四个选项齐全 ----------
    # ① 以“数字+.”切分题段，遇参考答案标题即停止；② 第 1-20 题必须齐全；
    # ③ 每题段内须能匹配 A./B./C./D. 四个选项。
    segs = {}
    cur_num = None
    for para in body_paras:
        text = _norm(_text_of(para))
        if ANSWER_TITLE in text:
            break
        match = QNUM_RE.match(text)
        if match:
            cur_num = int(match.group(1))
            segs.setdefault(cur_num, "")
        if cur_num is not None and text:
            segs[cur_num] = segs[cur_num] + " " + text
    choice_bad = []
    for num in range(1, 21):
        seg = segs.get(num)
        if seg is None:
            choice_bad.append("第%d题缺失" % num)
            continue
        lack = [opt for opt in "ABCD"
                if not re.search(r"(?<![A-Za-z])%s\s*[.．、]" % opt, seg)]
        if lack:
            choice_bad.append("第%d题缺选项%s" % (num, "".join(lack)))
    if not choice_bad:
        results.append("-0：第一部分选择题保持第 1 题至第 20 题共 20 道小题、每题 A/B/C/D 四个选项齐全")
    else:
        results.append("-3：第一部分选择题不满足 20 道小题、每题四个选项齐全（%s）"
                       % "；".join(choice_bad[:4]))

    # ---------- 29 -1 第 21-31 题题号连续 ----------
    # ① 复用上面的题段编号；② 21-31 必须全部出现；③ 跳号/重排/拆分/删减即扣分。
    tail_nums = sorted(n for n in segs if 21 <= n <= 31)
    if tail_nums == list(range(21, 32)):
        results.append("-0：第21题至第31题题号连续编号，无跳号或删减")
    else:
        results.append("-1：第21题至第31题的题号不满足连续编号（实际出现 %s）" % tail_nums)

    # ---------- 30 -1 各小题分值标注全部保留 ----------
    # ① 用“（…分）”正则统计正文分值标注数量；② 数量须不少于 32 处；
    # ③ 覆盖“（2分）”“（8分）”“（每空1分，共3分）”等写法。
    marks = SCORE_MARK_RE.findall(body_text)
    if len(marks) >= EXPECTED_SCORE_MARKS:
        results.append("-0：各小题的分值标注全部保留（共 %d 处）" % len(marks))
    else:
        results.append("-1：各小题的分值标注不满足全部保留（仅剩 %d 处，应有 %d 处）"
                       % (len(marks), EXPECTED_SCORE_MARKS))

    # ---------- 31 -1 第25题表格 5 行 2 列及“用途/物质”表头 ----------
    # ① 在全部表格中寻找 5 行 2 列的表；② 首行须为“用途”“物质”；③ 结构或表头被改即扣分。
    tbl25 = None
    for tbl in tables:
        grid = _table_grid(tbl)
        if len(grid) == 5 and all(len(row) == 2 for row in grid):
            if [_squeeze(c) for c in grid[0]] == ["用途", "物质"]:
                tbl25 = grid
                break
    if tbl25 is not None:
        results.append("-0：第25题表格保持 5 行 2 列结构及“用途/物质”表头")
    else:
        results.append("-1：第25题表格不满足 5 行 2 列结构及“用途/物质”表头")


    # ---------- 32 -3 文末参考答案标题及 4 段答案内容完整保留 ----------
    # ① 定位"参考答案与评分建议"标题段；② 其后须含"一、选择题…四、综合应用与计算题"4 段；
    # ③ 标题缺失或任一段答案缺失即扣分。
    ans_idx = None
    for idx, para in enumerate(body_paras):
        if ANSWER_TITLE in _norm(_text_of(para)):
            ans_idx = idx
            break
    if ans_idx is None:
        results.append("-3：文档末尾“参考答案与评分建议（教师用，可按需删除）”标题及其后 4 段答案内容不满足完整保留（未找到答案标题）")
    else:
        ans_text = _squeeze("\n".join(_text_of(p) for p in body_paras[ans_idx:]))
        miss_parts = [part for part in ANSWER_PARTS if _squeeze(part) not in ans_text]
        if not miss_parts:
            results.append("-0：文档末尾参考答案标题及“一、选择题/二、填空题/三、实验探究题/四、综合应用与计算题”4 段答案内容完整保留")
        else:
            results.append("-3：文档末尾参考答案的 4 段答案内容不满足完整保留（缺 %s）"
                           % "、".join(miss_parts))

    # ---------- 33 -1 参考答案与试题正文之间分页另起一页 ----------
    # ① 检查答案标题段自身的 pageBreakBefore；② 或答案标题段/其前一段内含 w:br type="page"；
    # ③ 或两者之间存在分节符（sectPr）；④ 均无则视为答案与试题挤在同一页。
    if ans_idx is None:
        results.append("-1：“参考答案与评分建议（教师用，可按需删除）”不满足与试题正文之间分页另起一页（未找到答案标题）")
    else:
        ans_para = body_paras[ans_idx]
        prev_para = body_paras[ans_idx - 1] if ans_idx > 0 else None
        ppr = _pPr(ans_para)
        page_new = ppr is not None and _on_off(ppr.find(qn("w:pageBreakBefore")))
        for holder in (ans_para, prev_para):
            if holder is None:
                continue
            if any(br.get(_TYPE) == "page" for br in holder.iter(_BR)):
                page_new = True
            if holder.find(".//" + _SECTPR) is not None:
                page_new = True
        if page_new:
            results.append("-0：参考答案与试题正文之间分页另起一页")
        else:
            results.append("-1：“参考答案与评分建议（教师用，可按需删除）”不满足与试题正文之间分页另起一页（答案与试题挤在同一页）")

    # ---------- 34 -1 试卷正文末尾单独一行居中的"（试卷结束）" ----------
    # ① 须存在整段文字恰为"（试卷结束）"的段落；② 该段落须水平居中。
    end_paras = [p for p in body_paras if _squeeze(_text_of(p)) == _squeeze(END_MARK)]
    if end_paras and any(_jc_of(p) == "center" for p in end_paras):
        results.append("-0：试卷正文末尾保留单独一行居中的“（试卷结束）”")
    elif end_paras:
        results.append("-1：试卷正文末尾“（试卷结束）”不满足单独一行居中（该段未水平居中）")
    else:
        results.append("-1：试卷正文末尾不满足单独一行居中的“（试卷结束）”（未找到该行）")

    # ---------- 35 -1 封面页无竖排主标题与四项填写栏之外的多余内容 ----------
    # ① 封面段落文字只允许主标题与四项填写栏；② 封面不得出现表格、段落边框（装饰线框）；
    # ③ 封面除主标题文本框/图形外不得再有图形对象；④ 封面所在节不得引用带文字的页眉。
    allow_texts = set()
    if title.get("text"):
        allow_texts.add(_squeeze(title["text"]))
    for _label, field_para in fields.items():
        allow_texts.add(_squeeze(_text_of(field_para)))
    cover_extra = []
    for el in cover_els:
        if el.tag == _TBL:
            cover_extra.append("封面存在表格")
            continue
        if el.tag != _P:
            continue
        text = _squeeze(_text_of(el))
        if text and text not in allow_texts:
            cover_extra.append("多余文字“%s”" % text[:16])
        ppr = _pPr(el)
        if ppr is not None and ppr.find(qn("w:pBdr")) is not None:
            cover_extra.append("装饰线框")
    n_shapes = sum(len(_drawings_of(el)) for el in cover_els if el.tag == _P)
    if n_shapes > 1:
        cover_extra.append("封面含 %d 个图形对象" % n_shapes)
    if sects:
        for ref in sects[0].findall(qn("w:headerReference")):
            rid = ref.get(qn("r:id"))
            try:
                part = doc.part.related_parts[rid]
                head_text = _norm(_text_of(etree.fromstring(part.blob)))
            except Exception:
                head_text = ""
            if head_text:
                cover_extra.append("页眉文字“%s”" % head_text[:16])
    if not cover_extra:
        results.append("-0：封面页仅有竖排主标题与四项填写栏，无页眉文字、装饰线框、说明文字等多余内容")
    else:
        results.append("-1：封面页出现竖排主标题与四项填写栏之外的多余内容（%s）"
                       % "；".join(dict.fromkeys(cover_extra))[:120])

    # ---------- 汇总：按 _MAX_DELTAS_ORDER 逐条回解为结构化结果 ----------
    if len(results) != len(_MAX_DELTAS_ORDER):
        raise RuntimeError("dim2 项数不一致：results=%d vs order=%d"
                           % (len(results), len(_MAX_DELTAS_ORDER)))
    dim2_items = []
    bonus_score = 0
    penalty_score = 0
    for raw, max_delta in zip(results, _MAX_DELTAS_ORDER):
        head, _, rule = raw.partition("：")
        head = head.strip()
        try:
            delta = int(head)
        except ValueError:
            delta = 0
        if delta and (delta > 0) != (max_delta > 0):
            raise RuntimeError("dim2 符号不一致：head=%s vs max_delta=%d（%s）"
                               % (head, max_delta, rule[:24]))
        hit = delta != 0
        dim2_items.append({"rule": rule, "max_delta": max_delta,
                           "delta": max_delta if hit else 0,
                           "hit": hit})
        if hit and max_delta > 0:
            bonus_score += max_delta
        elif hit:
            penalty_score += max_delta
    result["dim2_items"] = dim2_items
    result["total_score"] = bonus_score + penalty_score
    result["max_score"] = MAX_SCORE_TOTAL
    return result


if __name__ == "__main__":
    _dir = sys.argv[1] if len(sys.argv) >= 2 else os.path.dirname(os.path.abspath(__file__))
    print(json.dumps(evaluate(_dir), ensure_ascii=False, indent=2))
