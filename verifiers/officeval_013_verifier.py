#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# pyright: reportUnknownMemberType=false, reportUnknownVariableType=false, reportUnknownArgumentType=false, reportMissingImports=false, reportPrivateUsage=false, reportOptionalMemberAccess=false, reportAttributeAccessIssue=false, reportIndexIssue=false, reportArgumentType=false
"""
《基于Spark的园区能耗分析与可视化系统论文》格式核验任务 —— 完成度评测脚本。

评测原则：帮用户干了多少活（加分项 21 条，满分 57 分）、惹了多少麻烦（减分项 26 条，
全部命中最多扣 48 分）。加分项来源于《本科毕业论文内容与格式核验承诺书》17 项要求中
交付文档尚未达标、须动手修改的部分；承诺书要求中交付文档已经达标的部分（正文字体、
正文缩进与对齐、正文行距、英文题目与 Abstract 版式、表内字体、封皮页无页眉页码、标题
前后空行分隔）不计加分，只在被改坏时计入减分。减分项另含承诺书与指令均未涉及、却被
改动而破坏交付文档既有内容与版式的多余修改。两者作用于不同对象，互不重复。

接口约定：
本模块不修改全局 sys.stdout，也不 print 主结果。主结果统一由 evaluate(dir_path) -> dict
返回；入参为"被评估文档所在目录的路径"，由脚本自己在该目录内定位并打开被评估的 .docx。
本文件仅在 __main__ 下作为本地调试入口，才把 dict 序列化为 JSON 打印。

全部判定基于 OOXML（docx 包内 XML）静态解析，不依赖 Word COM 与 PDF 渲染，
可在无 Office 的环境下运行。
"""
import os
import re
import sys
import json

from docx import Document
from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph
from docx.text.run import Run

SCRIPT_ID = "013"
MAX_SCORE_TOTAL = 57       # 加分项满分
MIN_SCORE_TOTAL = -48      # 减分项全部命中时的扣分总额

# 本表须与 _evaluate_impl 中 results.append(...) 的顺序严格一致。
# 正值为加分项分值（命中即得分），负值为减分项分值（命中即扣分）。
_MAX_DELTAS_ORDER = [
    # 一、加分项（21 条 / 57 分）
    3,    # 01 一级标题宋体二号加粗居中
    3,    # 02 二级标题宋体三号加粗左对齐
    3,    # 03 三级标题宋体小四加粗左对齐
    5,    # 04 三级题序统一重排为 第一章→（一）→1、
    3,    # 05 每一新章节另起一页
    1,    # 06 摘要标题与摘要正文格式
    1,    # 07 关键词格式与分隔符
    1,    # 08 Key words 格式
    3,    # 09 一级目录条目宋体三号加粗
    3,    # 10 二级/三级目录条目宋体小四并缩进 2/4 字符
    5,    # 11 目录题序同步且页码与正文一致
    3,    # 12 自摘要页起页眉为所属章节名
    1,    # 13 页眉文字黑体五号居中
    5,    # 14 分节并采用罗马数字/正文重起 1
    1,    # 15 摘要与目录页码宋体小五居中
    1,    # 16 正文页码 TNR 小五居中且为"第 X 页"
    3,    # 17 图名在图下方、黑体五号、按章连续编号
    3,    # 18 表题在表格上方、黑体五号
    5,    # 19 全部表格改为三线表
    1,    # 20 表头文字加粗
    3,    # 21 程序代码 TNR 五号、中文注释宋体五号、1.25 倍行距
    # 二、减分项（26 条 / -48 分）
    -1,   # 22 封皮首行"澄云信息科技学院"版式被改
    -1,   # 23 封皮"毕业设计（论文）"版式被改
    -1,   # 24 封皮信息行版式/制表符/日期连排被改
    -1,   # 25 封皮末尾中英文校名字号被改
    -3,   # 26 删除封皮校名下方图形对象
    -1,   # 27 论文题目段落被并入封皮页
    -3,   # 28 封皮页被写入页眉文字或插入页码域
    -1,   # 29 页面设置（页边距、页眉页脚距边界）被改
    -3,   # 30 正文段落中文不再是宋体小四
    -3,   # 31 正文段落英文数字公式不再是 TNR 小四
    -3,   # 32 正文段落首行缩进 2 字符/两端对齐被破坏
    -3,   # 33 正文段落 1.5 倍行距、段前段后 0 行被破坏
    -1,   # 34 英文题目与 Abstract、英文摘要正文版式被破坏
    -1,   # 35 表内中文宋体五号、数字英文 TNR 五号被破坏
    -3,   # 36 参考文献与致谢正文中文被替换为第三方字体
    -3,   # 37 一级标题降级为正文样式并删除既有段落样式定义
    -3,   # 38 参考文献被添加"[n] "序号前缀
    -3,   # 39 程序代码行首缩进空白被删除
    -3,   # 40 标题之间/标题与正文之间的空行分隔被删除
    -1,   # 41 "致  谢"标题文字被改写
    -1,   # 42 目录域被改为手工纯文本
    -1,   # 43 表格单元格段落被统一改为居中对齐
    -1,   # 44 图名表题被加粗或设置段前段后 6 磅
    -1,   # 45 中文摘要正文首行缩进被清零或改为 1 字符
    -1,   # 46 "关键词：……"段落被改为左对齐、段前 12 磅
    -1,   # 47 各节实际显示的页眉残留拼音乱码或与本节不符的章节名
]

assert sum(d for d in _MAX_DELTAS_ORDER if d > 0) == MAX_SCORE_TOTAL
assert sum(d for d in _MAX_DELTAS_ORDER if d < 0) == MIN_SCORE_TOTAL

# ---------------------------------------------------------------- 常量
STRICT_SONGTI = ('宋体', 'SimSun', 'NSimSun', 'STSong')
STRICT_HEITI = ('黑体', 'SimHei', 'STHeiti')
STRICT_TNR = ('Times New Roman', 'TimesNewRoman', 'Times New Roman PS')
CJK_RE = re.compile(r'[㐀-鿿豈-﫿]')
LATIN_RE = re.compile(r'[A-Za-z0-9]')
ALPHA_RE = re.compile(r'[A-Za-z]')
CN_NUM = '一二三四五六七八九十'
# 聚合类条目（同一格式要求覆盖多个段落/单元格）的合规率门槛，容忍个别噪声
RATIO_OK = 0.85

# ---------------------------------------------------------------- 通用工具
def _norm(text):
    """去掉全部空白，用于"摘  要"/"摘要"一类标题文字的比对。"""
    return re.sub(r'\s+', '', text or '')


def _font_in(name, family):
    """字体名是否属于指定字族（兼容中英文别名与主题字体写法）。"""
    if not name:
        return False
    low = str(name).strip().lower()
    for k in family:
        if low == k.lower() or k.lower() in low:
            return True
    return False


def _near(value, target, tol=0.6):
    return value is not None and abs(float(value) - float(target)) <= tol


def _styles_map(doc):
    out = {}
    for st in doc.styles.element.findall(qn('w:style')):
        out[st.get(qn('w:styleId'))] = st
    return out


def _default_para_style_id(doc):
    """w:type="paragraph" 且 w:default="1" 的样式，即无 pStyle 时的兜底样式。"""
    for st in doc.styles.element.findall(qn('w:style')):
        if st.get(qn('w:type')) == 'paragraph' and st.get(qn('w:default')) in ('1', 'true', 'on'):
            return st.get(qn('w:styleId'))
    return None


def _para_style_id(p_el):
    pPr = p_el.find(qn('w:pPr'))
    if pPr is None:
        return None
    ps = pPr.find(qn('w:pStyle'))
    return ps.get(qn('w:val')) if ps is not None else None


def _style_chain(smap, style_id):
    """沿 w:basedOn 展开样式继承链（带环路保护）。"""
    out = []
    seen = set()
    while style_id and style_id in smap and style_id not in seen:
        seen.add(style_id)
        st = smap[style_id]
        out.append(st)
        b = st.find(qn('w:basedOn'))
        style_id = b.get(qn('w:val')) if b is not None else None
    return out

def _doc_default_rpr(doc):
    dd = doc.styles.element.find(qn('w:docDefaults'))
    if dd is None:
        return None
    rpd = dd.find(qn('w:rPrDefault'))
    return rpd.find(qn('w:rPr')) if rpd is not None else None


def _rpr_stack(doc, smap, run, p_el):
    """按 Word 的优先级返回 rPr 节点栈：直接格式 → 字符样式链 → 段落样式链 →
    默认段落样式 → docDefaults。取值时"第一个给出该属性的节点"生效。"""
    stack = []
    rPr = run._element.rPr if run is not None else None
    if rPr is not None:
        stack.append(rPr)
        rs = rPr.find(qn('w:rStyle'))
        if rs is not None:
            for st in _style_chain(smap, rs.get(qn('w:val'))):
                sub = st.find(qn('w:rPr'))
                if sub is not None:
                    stack.append(sub)
    sid = _para_style_id(p_el) if p_el is not None else None
    chain = _style_chain(smap, sid) if sid else []
    if not chain or chain[-1].get(qn('w:styleId')) != _default_para_style_id(doc):
        chain = chain + _style_chain(smap, _default_para_style_id(doc))
    for st in chain:
        sub = st.find(qn('w:rPr'))
        if sub is not None:
            stack.append(sub)
    dflt = _doc_default_rpr(doc)
    if dflt is not None:
        stack.append(dflt)
    return stack


def _resolve_effective_eastasia(doc, smap, run, p_el):
    """解析中文（东亚）字体的实际生效值。"""
    for rPr in _rpr_stack(doc, smap, run, p_el):
        rf = rPr.find(qn('w:rFonts'))
        if rf is not None:
            v = rf.get(qn('w:eastAsia')) or rf.get(qn('w:eastAsiaTheme'))
            if v:
                return v
    return None


def _resolve_effective_ascii(doc, smap, run, p_el):
    """解析西文（ascii）字体的实际生效值。"""
    for rPr in _rpr_stack(doc, smap, run, p_el):
        rf = rPr.find(qn('w:rFonts'))
        if rf is not None:
            v = rf.get(qn('w:ascii')) or rf.get(qn('w:asciiTheme'))
            if v:
                return v
    return None

def _resolve_effective_size(doc, smap, run, p_el):
    """解析字号（磅）。w:sz 以半磅为单位。"""
    for rPr in _rpr_stack(doc, smap, run, p_el):
        sz = rPr.find(qn('w:sz'))
        if sz is not None and sz.get(qn('w:val')):
            try:
                return float(sz.get(qn('w:val'))) / 2.0
            except ValueError:
                return None
    return None


def _resolve_effective_bold(doc, smap, run, p_el):
    """解析加粗。w:b 无 val 视为 True，val=0/false/off 视为 False。"""
    for rPr in _rpr_stack(doc, smap, run, p_el):
        b = rPr.find(qn('w:b'))
        if b is not None:
            return b.get(qn('w:val')) not in ('0', 'false', 'off')
    return False


def _ppr_stack(doc, smap, p_el):
    """段落属性优先级栈：段落直接 pPr → 段落样式链 pPr → 默认段落样式 pPr。"""
    stack = []
    pPr = p_el.find(qn('w:pPr'))
    if pPr is not None:
        stack.append(pPr)
    sid = _para_style_id(p_el)
    chain = _style_chain(smap, sid) if sid else []
    if not chain or chain[-1].get(qn('w:styleId')) != _default_para_style_id(doc):
        chain = chain + _style_chain(smap, _default_para_style_id(doc))
    for st in chain:
        sub = st.find(qn('w:pPr'))
        if sub is not None:
            stack.append(sub)
    return stack


def _resolve_ppr_attr(doc, smap, p_el, tag, attr):
    """取段落属性（如 w:jc/@w:val、w:spacing/@w:line、w:ind/@w:firstLineChars）。"""
    for pPr in _ppr_stack(doc, smap, p_el):
        el = pPr.find(qn(tag))
        if el is not None:
            v = el.get(qn(attr))
            if v is not None:
                return v
    return None


def _align_of(doc, smap, p_el):
    """对齐方式，返回 'left'/'center'/'right'/'both'/None（None 即未设=左对齐）。"""
    return _resolve_ppr_attr(doc, smap, p_el, 'w:jc', 'w:val')

def _line_spacing(doc, smap, p_el):
    """返回 (line, lineRule)。1.5 倍行距=（360, auto）；1.25 倍=（300, auto）。"""
    for pPr in _ppr_stack(doc, smap, p_el):
        sp = pPr.find(qn('w:spacing'))
        if sp is not None and sp.get(qn('w:line')):
            try:
                line = float(sp.get(qn('w:line')))
            except ValueError:
                return (None, None)
            return (line, sp.get(qn('w:lineRule')) or 'auto')
    return (None, None)


def _is_multiple_spacing(doc, smap, p_el, times):
    line, rule = _line_spacing(doc, smap, p_el)
    if line is None:
        return times == 1.0
    return rule in ('auto', None) and abs(line - times * 240.0) <= 6.0


def _space_pt(doc, smap, p_el, which):
    """段前/段后间距（磅）。which 取 'before' / 'after'；未设置按 0 计。"""
    for pPr in _ppr_stack(doc, smap, p_el):
        sp = pPr.find(qn('w:spacing'))
        if sp is None:
            continue
        lines = sp.get(qn('w:%sLines' % which))
        if lines:
            return float(lines) / 100.0 * 12.0
        v = sp.get(qn('w:%s' % which))
        if v:
            return float(v) / 20.0
    return 0.0


def _indent_chars(doc, smap, p_el, kind, size_pt):
    """首行/左缩进换算成"中文字符"数。kind 取 'firstLine' / 'left'。"""
    chars_attr = 'w:%sChars' % kind
    for pPr in _ppr_stack(doc, smap, p_el):
        ind = pPr.find(qn('w:ind'))
        if ind is None:
            continue
        v = ind.get(qn(chars_attr))
        if v is not None:
            return float(v) / 100.0
        v = ind.get(qn('w:%s' % kind))
        if v is not None:
            return float(v) / 20.0 / max(size_pt or 12.0, 1.0)
    return 0.0


def _indent_unit(doc, smap, p_el, kind):
    """缩进的设置单位。返回 ('chars', 字符数) 或 ('abs', 厘米)，未设置缩进返回 (None, 0.0)。
    与办公软件对齐：以"字符"为单位设置时 w:ind 带 @w:*Chars（如 firstLineChars="200"），
    以厘米/磅为单位设置时只有 @w:firstLine（缇），字号变化时不随字符宽度缩放。"""
    chars_attr = 'w:%sChars' % kind
    for pPr in _ppr_stack(doc, smap, p_el):
        ind = pPr.find(qn('w:ind'))
        if ind is None:
            continue
        v = ind.get(qn(chars_attr))
        if v is not None:
            try:
                return ('chars', float(v) / 100.0)
            except ValueError:
                return (None, 0.0)
        v = ind.get(qn('w:%s' % kind))
        if v is not None:
            try:
                return ('abs', float(v) / 1440.0 * 2.54)
            except ValueError:
                return (None, 0.0)
    return (None, 0.0)

def _outline_level(doc, smap, p_el):
    """标题级别：0=一级、1=二级、2=三级；None 表示不是标题。"""
    for pPr in _ppr_stack(doc, smap, p_el):
        ol = pPr.find(qn('w:outlineLvl'))
        if ol is not None:
            try:
                lv = int(ol.get(qn('w:val')))
            except (TypeError, ValueError):
                return None
            return lv if lv <= 8 else None
    return None


def _text_runs(para):
    """过滤掉空文本 run。"""
    return [r for r in para.runs if r.text and r.text.strip()]


def _has_drawing(p_el):
    return (p_el.find('.//' + qn('w:drawing')) is not None
            or p_el.find('.//' + qn('w:pict')) is not None)


def _is_code_para(text):
    """程序代码段落：西文为主且含代码符号/关键字，中文占比低（排除正文与参考文献）。"""
    s = (text or '').strip()
    if not s:
        return False
    if len(CJK_RE.findall(s)) / float(len(s)) > 0.3:
        return False
    if re.search(r'[=;{}\[\]]|\.\w+\(|\(\)', s):
        return True
    return bool(re.match(r'^(from|import|def|class|return|if|for|while|else|elif|try|'
                         r'except|with|SELECT|CREATE|INSERT|UPDATE|WITH)\b', s))


def _field_instr(p_el):
    """段落内的域代码文本（w:instrText 与 w:fldSimple/@w:instr 合并）。"""
    parts = []
    for it in p_el.iter(qn('w:instrText')):
        parts.append(it.text or '')
    for fs in p_el.iter(qn('w:fldSimple')):
        parts.append(fs.get(qn('w:instr')) or '')
    return ' '.join(parts)


def _open_word_document(file_path):
    return Document(file_path)

# ---------------------------------------------------------------- 表格边框
NO_LINE = ('none', 'nil')


def _tbl_borders(tbl_el):
    tblPr = tbl_el.find(qn('w:tblPr'))
    return tblPr.find(qn('w:tblBorders')) if tblPr is not None else None


def _tc_borders(tc_el):
    tcPr = tc_el.find(qn('w:tcPr'))
    return tcPr.find(qn('w:tcBorders')) if tcPr is not None else None


def _line_info(b_el):
    """返回 (has_line, width_pt)；None 表示该边未定义，需向上一级回退。"""
    if b_el is None:
        return None
    val = b_el.get(qn('w:val')) or 'single'
    if val in NO_LINE:
        return (False, 0.0)
    sz = b_el.get(qn('w:sz'))
    try:
        width = float(sz) / 8.0 if sz else 0.5
    except ValueError:
        width = 0.5
    return (True, width)


def _edge(tbl_el, tc_el, edge, first_row, last_row, first_col, last_col):
    """单元格某条边的实际生效线：单元格 tcBorders 优先，其次表级 tblBorders
    （外框用 top/bottom/left/right，内部用 insideH/insideV），都没有则视为无线。"""
    info = _line_info(_none_or_find(_tc_borders(tc_el), edge))
    if info is not None:
        return info
    tb = _tbl_borders(tbl_el)
    if edge == 'top':
        name = 'top' if first_row else 'insideH'
    elif edge == 'bottom':
        name = 'bottom' if last_row else 'insideH'
    elif edge == 'left':
        name = 'left' if first_col else 'insideV'
    else:
        name = 'right' if last_col else 'insideV'
    info = _line_info(_none_or_find(tb, name))
    return info if info is not None else (False, 0.0)


def _none_or_find(node, name):
    return None if node is None else node.find(qn('w:' + name))

def _is_damaged_package(exc):
    """判断异常是否源于交付文件自身结构损坏（缺部件、XML 不合法等），
    而非脚本逻辑缺陷。"""
    text = str(exc)
    if isinstance(exc, KeyError) and (
        "There is no item named" in text or "no relationship of type" in text
    ):
        return True
    return type(exc).__name__ in ("XMLSyntaxError", "ParseError", "BadZipFile",
                                  "InvalidXmlError")


# ---------------------------------------------------------------- 对外接口
def evaluate(dir_path):
    """入参为被评估文档所在目录；返回评测结果 dict，不 print。

    定位不到待评测文件时（目录不存在 / 目录内无 .docx），dim1_pass=False、
    dim1_reason=""、dim2_items=[]、total_score=max_score=0，error 仅给出可读
    原因（不带异常类名前缀）。
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
            raise FileNotFoundError("目录不存在或不是目录：%s" % dir_path)
        cands = sorted(f for f in os.listdir(dir_path)
                       if f.lower().endswith(('.docx', '.doc')) and not f.startswith('~$'))
        if not cands:
            raise FileNotFoundError("目录中未找到 .docx 文件：%s" % dir_path)
        file_path = os.path.join(dir_path, cands[0])
        result["file_name"] = os.path.basename(file_path)
        return _evaluate_impl(file_path, result)
    except FileNotFoundError as exc:
        result["status"] = "error"
        result["error"] = str(exc)
        # 维度一根本没进入判定，dim1_pass 与 dim1_reason 一律留空（空字符串，而非 False）
        result["dim1_pass"] = ""
        result["dim1_reason"] = ""
        result["dim2_items"] = []
        result["total_score"] = 0
        result["max_score"] = 0
        return result
    except Exception as exc:
        if _is_damaged_package(exc):
            # 交付文件结构损坏（缺部件、XML 非法、OPC 关系不完整等）属于维度一
            # 不通过，按正常 0 分处理，不记为脚本运行错误。
            result["status"] = "ok"
            result["error"] = None
            result["dim1_pass"] = False
            result["dim1_reason"] = "交付文件结构不完整，无法正常打开：%s: %s" % (
                type(exc).__name__, exc)
            result["dim2_items"] = []
            result["total_score"] = 0
            return result
        result["status"] = "error"
        result["error"] = "脚本运行异常：%s: %s" % (type(exc).__name__, exc)
        result["dim1_pass"] = False
        result["dim1_reason"] = "文档无法打开或解析：%s" % exc
        result["dim2_items"] = []
        result["total_score"] = 0
        return result


def _evaluate_impl(file_path, result):
    doc = _open_word_document(file_path)
    smap = _styles_map(doc)
    body = doc.element.body

    # ---------------- 文档结构建模 ----------------
    blocks = []            # 正文层块序列：('p'|'tbl'|'sdt', element)
    for child in body.iterchildren():
        if child.tag == qn('w:p'):
            blocks.append(('p', child))
        elif child.tag == qn('w:tbl'):
            blocks.append(('tbl', child))
        elif child.tag == qn('w:sdt'):
            blocks.append(('sdt', child))

    # 分节：段落 pPr/sectPr 标记"上一节结束"，body/sectPr 为最后一节
    sect_prs = []
    blk_sect = []
    cur = 0
    for kind, el in blocks:
        blk_sect.append(cur)
        if kind == 'p':
            pPr = el.find(qn('w:pPr'))
            sp = pPr.find(qn('w:sectPr')) if pPr is not None else None
            if sp is not None:
                sect_prs.append(sp)
                cur += 1
    tail_sect = body.find(qn('w:sectPr'))
    if tail_sect is not None:
        sect_prs.append(tail_sect)

    # 正文层段落记录
    paras = []
    for bi, (kind, el) in enumerate(blocks):
        if kind != 'p':
            continue
        para = Paragraph(el, doc)
        paras.append({
            'pi': len(paras), 'bi': bi, 'el': el, 'para': para,
            'text': para.text, 'norm': _norm(para.text),
            'ol': _outline_level(doc, smap, el),
            'sect': blk_sect[bi],
            'blank': not para.text.strip() and not _has_drawing(el),
        })

    def find_para(pred, default=None):
        for rec in paras:
            if pred(rec):
                return rec
        return default

    CHAP_RE = re.compile(r'^第\s*[0-9%s]+\s*章' % CN_NUM)
    r_abs = find_para(lambda r: r['norm'] == '摘要')
    r_eabs = find_para(lambda r: r['norm'].lower() == 'abstract')
    r_toc = find_para(lambda r: r['norm'] in ('目录', '目录contents'))
    r_ref = find_para(lambda r: r['norm'] == '参考文献')
    r_thx = find_para(lambda r: r['norm'] in ('致谢', '致  谢'.replace(' ', '')))
    i_abs = r_abs['pi'] if r_abs else 0
    i_toc = r_toc['pi'] if r_toc else i_abs
    i_ref = r_ref['pi'] if r_ref else len(paras)
    i_thx = r_thx['pi'] if r_thx else len(paras)

    # 手工目录条目：以制表符／前导点结尾接页码标签的段落，不参与标题与正文判定
    TOC_ENTRY_RE = re.compile(r'(\t|\.{3,}|…+|·{3,})\s*(\d+|[IVXLCDM]+|[ⅠⅡⅢⅣⅤⅥⅦⅧⅨⅩ]+)\s*$')

    def is_toc_entry(r):
        return TOC_ENTRY_RE.search(r['text'].rstrip()) is not None

    # 正文起点：目录之后第一个"第X章"标题段落（跳过手工目录中的同名条目）
    r_chap1 = find_para(lambda r: r['pi'] > i_toc and CHAP_RE.match(r['text'].strip())
                        and not is_toc_entry(r))
    i_body = r_chap1['pi'] if r_chap1 else i_toc + 1
    toc_range = set(range(i_toc + 1, i_body))      # 目录条目区（手工目录时为纯段落）
    cover_range = list(range(0, i_abs))            # 封皮区

    # 一级标题集合：按细则点名的部分标题 + 各章标题（排除目录条目区）
    L1_NAMES = ('摘要', 'abstract', '目录', '参考文献', '致谢')
    l1 = [r for r in paras
          if r['pi'] not in toc_range and r['text'].strip() and not is_toc_entry(r)
          and (r['norm'].lower() in L1_NAMES or CHAP_RE.match(r['text'].strip()))]
    chap_titles = [r for r in l1 if CHAP_RE.match(r['text'].strip())]
    # 二、三级标题集合：以大纲级别为准（大纲级别缺失时退化为题序模式匹配）
    l2 = [r for r in paras if r['ol'] == 1 and r['pi'] not in toc_range and r['text'].strip()
          and not is_toc_entry(r)]
    l3 = [r for r in paras if r['ol'] == 2 and r['pi'] not in toc_range and r['text'].strip()
          and not is_toc_entry(r)]
    if not l2:
        l2 = [r for r in paras if r['pi'] not in toc_range and not is_toc_entry(r)
              and re.match(r'^(（[%s]+）|\d+\.\d+(?!\.\d))' % CN_NUM, r['text'].strip())
              and len(r['text'].strip()) < 40]
    if not l3:
        l3 = [r for r in paras if r['pi'] not in toc_range and not is_toc_entry(r)
              and re.match(r'^(\d+、|\d+\.\d+\.\d+)', r['text'].strip())
              and len(r['text'].strip()) < 40]
    title_pis = set(r['pi'] for r in l1 + l2 + l3)

    CAP_FIG_RE = re.compile(r'^图\s*(\d+)\s*[.．\-]\s*(\d+)')
    CAP_TBL_RE = re.compile(r'^表\s*(\d+)\s*[.．\-]\s*(\d+)')
    figs = [r for r in paras if CAP_FIG_RE.match(r['text'].strip())]
    tbl_caps = [r for r in paras if CAP_TBL_RE.match(r['text'].strip())]
    caption_pis = set(r['pi'] for r in figs + tbl_caps)

    # 程序代码段落（正文区内）
    codes = [r for r in paras if i_body <= r['pi'] < i_ref
             and r['pi'] not in title_pis and _is_code_para(r['text'])]
    code_pis = set(r['pi'] for r in codes)

    # 正文文字段落：正文区内、非标题、非图表题注、非代码、非空、无图形
    body_paras = [r for r in paras if i_body <= r['pi'] < i_ref
                  and not r['blank'] and r['pi'] not in title_pis
                  and r['pi'] not in caption_pis and r['pi'] not in code_pis
                  and not _has_drawing(r['el'])]

    # 目录条目：优先取 w:sdt 内的目录域（python-docx 的 doc.paragraphs 不含 sdt 内容）
    toc_field = False
    toc_paras = []
    for kind, el in blocks:
        if kind != 'sdt':
            continue
        if re.search(r'\bTOC\b', _field_instr(el), re.I):
            toc_field = True
        for p_el in el.iter(qn('w:p')):
            if Paragraph(p_el, doc).text.strip():
                toc_paras.append(Paragraph(p_el, doc))
    for r in paras:
        if re.search(r'\bTOC\b', _field_instr(r['el']), re.I):
            toc_field = True
    if not toc_paras:
        toc_paras = [r['para'] for r in paras if r['pi'] in toc_range and r['text'].strip()]

    def toc_level(p):
        for st in _style_chain(smap, _para_style_id(p._p)):
            nm = st.find(qn('w:name'))
            val = (nm.get(qn('w:val')) if nm is not None else '') or ''
            m = re.match(r'^toc\s*([1-3])$', val.strip(), re.I)
            if m:
                return int(m.group(1))
        t = p.text.strip()
        if re.match(r'^第', t):
            return 1
        if re.match(r'^（', t) or re.match(r'^\d+\.\d+(?!\.\d)', t):
            return 2
        return 3

    PAGE_TAIL_RE = re.compile(r'([0-9]+|[IVXLCDM]+|[ⅠⅡⅢⅣⅤⅥⅦⅧⅨⅩ]+)\s*$')

    def toc_split(p):
        """拆出（条目文字, 页码标签）。"""
        t = re.sub(r'[\t\.·…  ]+$', '', p.text.strip())
        m = PAGE_TAIL_RE.search(t)
        label = m.group(1) if m else ''
        title = t[:m.start()] if m else t
        title = re.sub(r'[\t\.·…  ]+$', '', title)
        return (title, label)

    tables = list(doc.tables)

    # ---------------- 分节与页眉页脚 ----------------
    def sp_attr(sp, tag, attr):
        el = sp.find(qn(tag))
        return el.get(qn(attr)) if el is not None else None

    sects = []
    for sp in sect_prs:
        tp = sp.find(qn('w:titlePg'))
        rec = {
            'sp': sp,
            'titlePg': tp is not None and tp.get(qn('w:val')) not in ('0', 'false', 'off'),
            'header': {}, 'footer': {},
            'fmt': sp_attr(sp, 'w:pgNumType', 'w:fmt'),
            'start': sp_attr(sp, 'w:pgNumType', 'w:start'),
            'mar': {k: sp_attr(sp, 'w:pgMar', 'w:' + k)
                    for k in ('top', 'bottom', 'left', 'right', 'header', 'footer')},
        }
        for ref in sp.findall(qn('w:headerReference')):
            rec['header'][ref.get(qn('w:type')) or 'default'] = ref.get(qn('r:id'))
        for ref in sp.findall(qn('w:footerReference')):
            rec['footer'][ref.get(qn('w:type')) or 'default'] = ref.get(qn('r:id'))
        sects.append(rec)

    def part_of(rid):
        if not rid:
            return None
        try:
            return doc.part.related_parts[rid]
        except KeyError:
            return None

    def hf_paras(si, kind, typ):
        """取某节某类型页眉/页脚的段落；本节未定义时按 Word 规则向前继承。"""
        for k in range(si, -1, -1):
            rid = sects[k][kind].get(typ)
            part = part_of(rid)
            if part is not None:
                return [Paragraph(p, doc) for p in part.element.iter(qn('w:p'))]
        return []

    def hf_text(si, kind, typ):
        return _norm(''.join(p.text for p in hf_paras(si, kind, typ)))

    def shown_type(si):
        """该节首页实际显示的页眉/页脚类型。"""
        return 'first' if sects[si]['titlePg'] else 'default'

    # ---------------- 维度一：可用性 ----------------
    # 仅核验交付物是否为 1 个可正常打开的 Word 文档：能走到这里说明 _open_word_document
    # 已成功解析该 .docx/.doc，故判定通过；文件损坏无法打开的情形在 evaluate 中直接返回。
    result["dim1_pass"] = True
    result["dim1_reason"] = "交付物为 1 个可正常打开的 Word 文档"

    # ---------------- 维度二：逐条打分 ----------------
    results = []
    details = []

    def add(spec, rule, hit, detail=""):
        """spec 为细则分值（加分项为正、减分项为负）；hit 表示该条是否命中。"""
        results.append("%s%d：%s" % ('+' if spec > 0 else '-', abs(spec) if hit else 0, rule))
        details.append(detail if isinstance(detail, str) else str(detail))

    def ratio_ok(ok, total):
        return total > 0 and (float(ok) / float(total)) >= RATIO_OK

    # 承诺书要求中交付文档已经达标的项：在此处就近计算（依赖 body_paras/dom_size 等局部量），
    # 结果以 (是否仍然合规, 明细) 暂存，统一在减分项区按细则顺序发出。
    pend = {}

    def font_ok(p_el, para, ea=None, asc=None, size=None, bold=None, asc_alpha=False):
        """按 Word 的中西文分开设置规则校验一个段落：含中文的 run 查东亚字体，
        含西文/数字的 run 查 ascii 字体；字号与加粗对所有非空 run 生效。
        asc_alpha=True 时西文字体只核验含英文字母的 run（图表题序中的数字不计）。"""
        bad = []
        rs = _text_runs(para)
        if not rs:
            return (False, '段落无文字')
        lat = ALPHA_RE if asc_alpha else LATIN_RE
        for r in rs:
            if ea and CJK_RE.search(r.text):
                v = _resolve_effective_eastasia(doc, smap, r, p_el)
                if not _font_in(v, ea):
                    bad.append('中文字体=%s' % v)
            if asc and lat.search(r.text):
                v = _resolve_effective_ascii(doc, smap, r, p_el)
                if not _font_in(v, asc):
                    bad.append('西文字体=%s' % v)
            if size is not None:
                v = _resolve_effective_size(doc, smap, r, p_el)
                if not _near(v, size):
                    bad.append('字号=%s' % v)
            if bold is not None and _resolve_effective_bold(doc, smap, r, p_el) != bold:
                bad.append('加粗=%s' % (not bold))
        return (not bad, '；'.join(sorted(set(bad))[:3]))

    # ============ 【标题层级与编号】 ============
    # +3：各一级标题为宋体二号（22 磅）加粗、居中对齐
    # 判定：① 段落集合＝"摘 要/Abstract/目 录/参考文献/致 谢"+各章标题（排除目录条目区）；
    #       ② 含中文的 run 东亚字体为宋体；③ 全部 run 字号 22 磅；④ 加粗；⑤ w:jc=center。
    # 与办公软件对齐：字体读 w:rFonts/@w:eastAsia，字号读 w:sz（半磅），加粗读 w:b，对齐读 w:jc。
    ok, fails = 0, []
    for r in l1:
        f_ok, why = font_ok(r['el'], r['para'], ea=STRICT_SONGTI, size=22, bold=True)
        al = _align_of(doc, smap, r['el'])
        if f_ok and al == 'center':
            ok += 1
        else:
            fails.append('%s[%s%s]' % (r['norm'][:8], why, '' if al == 'center' else ' 对齐=%s' % al))
    add(3, '各一级标题（"摘  要""Abstract""目  录""第一章  绪论"等章标题、"参考文献""致  谢"）为宋体二号（22 磅）加粗、居中对齐',
        ratio_ok(ok, len(l1)), '合规 %d/%d；%s' % (ok, len(l1), '，'.join(fails[:4])))

    # +3：各二级标题为宋体三号（16 磅）加粗、左对齐
    ok, fails = 0, []
    for r in l2:
        f_ok, why = font_ok(r['el'], r['para'], ea=STRICT_SONGTI, size=16, bold=True)
        al = _align_of(doc, smap, r['el'])
        if f_ok and al in (None, 'left', 'start'):
            ok += 1
        else:
            fails.append('%s[%s%s]' % (r['text'].strip()[:10], why,
                                       '' if al in (None, 'left', 'start') else ' 对齐=%s' % al))
    add(3, '各二级标题（如"研究背景与意义""系统关键技术"）为宋体三号（16 磅）加粗、左对齐',
        ratio_ok(ok, len(l2)), '合规 %d/%d；%s' % (ok, len(l2), '，'.join(fails[:4])))

    # +3：各三级标题为宋体小四（12 磅）加粗、左对齐
    ok, fails = 0, []
    for r in l3:
        f_ok, why = font_ok(r['el'], r['para'], ea=STRICT_SONGTI, size=12, bold=True)
        al = _align_of(doc, smap, r['el'])
        if f_ok and al in (None, 'left', 'start'):
            ok += 1
        else:
            fails.append('%s[%s%s]' % (r['text'].strip()[:10], why,
                                       '' if al in (None, 'left', 'start') else ' 对齐=%s' % al))
    add(3, '各三级标题（如"Spark 分布式计算技术""可视化与系统架构技术"）为宋体小四（12 磅）加粗、左对齐',
        ratio_ok(ok, len(l3)), '合规 %d/%d；%s' % (ok, len(l3), '，'.join(fails[:4])))

    # +5：三级题序按"第一章 → （一） → 1、"体系统一重排且连续不跳号
    # 判定：① 章序为中文数字"第一章…"且从一起连续；②每章内节序为（一）（二）…连续；
    #       ③ 每节内小节序为"1、2、…"连续；任一层跳号或仍为"1.1/2.1.1"式即不得分。
    def cn_of(i):
        return CN_NUM[i] if i < len(CN_NUM) else str(i + 1)

    num_bad = []
    for ci, cr in enumerate(chap_titles):
        if not re.match(r'^第\s*%s\s*章' % cn_of(ci), cr['text'].strip()):
            num_bad.append('章序:%s' % cr['text'].strip()[:12])
    for ci, cr in enumerate(chap_titles):
        c_end = chap_titles[ci + 1]['pi'] if ci + 1 < len(chap_titles) else len(paras)
        subs = [r for r in l2 if cr['pi'] < r['pi'] < c_end]
        for k, s in enumerate(subs):
            if not re.match(r'^（\s*%s\s*）' % cn_of(k), s['text'].strip()):
                num_bad.append('节序:%s' % s['text'].strip()[:12])
            s_end = subs[k + 1]['pi'] if k + 1 < len(subs) else c_end
            for j, t in enumerate([r for r in l3 if s['pi'] < r['pi'] < s_end]):
                if not re.match(r'^%d\s*、' % (j + 1), t['text'].strip()):
                    num_bad.append('小节序:%s' % t['text'].strip()[:12])
    add(5, '三级题序按"第一章 → （一） → 1、"体系统一重排且连续不跳号（章序由"第1章"改为"第一章"，节序由"1.1、1.2"改为"（一）、（二）"，小节序由"2.1.1、2.1.2"改为"1、、2、"）',
        bool(chap_titles) and not num_bad,
        '章 %d 个、节 %d 个、小节 %d 个；不合规 %d 处：%s' % (
            len(chap_titles), len(l2), len(l3), len(num_bad), '，'.join(num_bad[:4])))

    # ============ 【正文既有字体与段落：承诺书已达标，仅在减分项区核验是否被改坏】 ============
    # 正文中文宋体小四（12 磅）
    ok, fails = 0, []
    for r in body_paras:
        f_ok, why = font_ok(r['el'], r['para'], ea=STRICT_SONGTI, size=12)
        ok += 1 if f_ok else 0
        if not f_ok:
            fails.append('P%d[%s]' % (r['pi'], why))
    pend['body_ea'] = (ratio_ok(ok, len(body_paras)),
                       '合规 %d/%d；%s' % (ok, len(body_paras), '，'.join(fails[:3])))

    # 正文中的英文、数字及公式的英文部分 Times New Roman 小四（12 磅）
    lat_paras = [r for r in body_paras if LATIN_RE.search(r['text'])]
    ok, fails = 0, []
    for r in lat_paras:
        f_ok, why = font_ok(r['el'], r['para'], asc=STRICT_TNR, size=12)
        ok += 1 if f_ok else 0
        if not f_ok:
            fails.append('P%d[%s]' % (r['pi'], why))
    pend['body_lat'] = (ratio_ok(ok, len(lat_paras)),
                        '合规 %d/%d；%s' % (ok, len(lat_paras), '，'.join(fails[:3])))

    def dom_size(rec):
        rs = _text_runs(rec['para'])
        if rs:
            v = _resolve_effective_size(doc, smap, rs[0], rec['el'])
            if v:
                return v
        return 12.0

    # 正文段落首行缩进 2 个中文字符且两端对齐
    # 与办公软件对齐：缩进须以"字符"为单位设置，即 w:ind/@w:firstLineChars=200；
    #                 只写 @w:firstLine（缇）表示按厘米/磅的绝对值设置，办公软件的段落
    #                 对话框显示为"0.85 厘米"而非"2 字符"，不满足"缩进 2 个中文字符"；
    #                 两端对齐为 w:jc=both。
    ok, fails = 0, []
    for r in body_paras:
        ch = _indent_chars(doc, smap, r['el'], 'firstLine', dom_size(r))
        unit, uval = _indent_unit(doc, smap, r['el'], 'firstLine')
        al = _align_of(doc, smap, r['el'])
        if abs(ch - 2.0) <= 0.25 and unit == 'chars' and al == 'both':
            ok += 1
        else:
            fails.append('P%d[缩进%s 对齐=%s]' % (
                r['pi'],
                '%.2f字符' % ch if unit == 'chars' else
                ('%.2f厘米（非字符单位）' % uval if unit == 'abs' else '未设置'),
                al))
    pend['body_ind'] = (ratio_ok(ok, len(body_paras)),
                        '合规 %d/%d；%s' % (ok, len(body_paras), '，'.join(fails[:3])))

    # 正文段落 1.5 倍行距、段前段后均为 0 行
    ok, fails = 0, []
    for r in body_paras:
        sb, sa = _space_pt(doc, smap, r['el'], 'before'), _space_pt(doc, smap, r['el'], 'after')
        if _is_multiple_spacing(doc, smap, r['el'], 1.5) and sb <= 0.5 and sa <= 0.5:
            ok += 1
        else:
            line, rule = _line_spacing(doc, smap, r['el'])
            fails.append('P%d[line=%s/%s 段前%.1f 段后%.1f]' % (r['pi'], line, rule, sb, sa))
    pend['body_line'] = (ratio_ok(ok, len(body_paras)),
                         '合规 %d/%d；%s' % (ok, len(body_paras), '，'.join(fails[:3])))

    # ============ 【正文分页】 ============

    # +3：每一新章节均另起一页
    # 判定：章标题位于本节起始处（其前最多允许一个题名块）且该节起始方式非"连续"，
    #       或带 w:pageBreakBefore，或紧前段落含 w:br@w:type=page。
    sect_first_blk = {}
    sect_order = {}        # 块序号 -> 该块在本节非空块中的序号（0 起）
    _cnt = {}
    for bi, (kind, el) in enumerate(blocks):
        si = blk_sect[bi]
        txt = ''.join(t.text or '' for t in el.iter(qn('w:t')))
        if not (kind == 'tbl' or txt.strip() or _has_drawing(el)):
            continue
        sect_order[bi] = _cnt.get(si, 0)
        _cnt[si] = sect_order[bi] + 1
        if si not in sect_first_blk:
            sect_first_blk[si] = bi

    def starts_new_page(rec):
        si = rec['sect']
        stype = sp_attr(sects[si]['sp'], 'w:type', 'w:val') if si < len(sects) else None
        # 分节符本身即分页：该节起始处的标题（允许其前只有一个题名/英文题目块）视为另起一页
        if stype != 'continuous' and sect_order.get(rec['bi'], 99) <= 1:
            return True
        pPr = rec['el'].find(qn('w:pPr'))
        pbb = pPr.find(qn('w:pageBreakBefore')) if pPr is not None else None
        if pbb is not None and pbb.get(qn('w:val')) not in ('0', 'false', 'off'):
            return True
        prev = paras[rec['pi'] - 1]['el'] if rec['pi'] > 0 else None
        if prev is not None:
            for br in prev.iter(qn('w:br')):
                if br.get(qn('w:type')) == 'page':
                    return True
        return False

    newpage_targets = [r for r in l1 if r['pi'] >= i_abs]
    np_ok = [r for r in newpage_targets if starts_new_page(r)]
    add(3, '每一新章节均另起一页', ratio_ok(len(np_ok), len(newpage_targets)),
        '另起一页 %d/%d；未另起：%s' % (
            len(np_ok), len(newpage_targets),
            '，'.join(r['norm'][:8] for r in newpage_targets if r not in np_ok)[:80]))

    # 标题与标题之间、标题与正文之间空一行（承诺书已达标项，仅在减分项区核验是否被删除）
    # 判定：每个标题与其下方内容（下一标题或正文）之间留出一行——相邻为空段落，
    #       或由段后/段前间距撑出不少于一行（小四正文一行约 12 磅，容差按 9 磅计）。
    LINE_PT = 9.0
    all_titles = sorted(l1 + l2 + l3, key=lambda r: r['pi'])

    def gap_ok(upper, lower):
        """upper／lower 为边界上下两个段落记录，任一为 None 视为已有分隔。"""
        if upper is None or lower is None:
            return True
        if upper['blank'] or lower['blank']:
            return True
        return (_space_pt(doc, smap, upper['el'], 'after') >= LINE_PT
                or _space_pt(doc, smap, lower['el'], 'before') >= LINE_PT)

    ok, fails = 0, []
    for r in all_titles:
        nxt = paras[r['pi'] + 1] if r['pi'] + 1 < len(paras) else None
        if gap_ok(r, nxt):
            ok += 1
        else:
            fails.append('%s[与下方未空行]' % r['text'].strip()[:10])
    pend['title_gap'] = (ratio_ok(ok, len(all_titles)),
                         '合规 %d/%d；%s' % (ok, len(all_titles), '，'.join(fails[:4])))

    # ============ 【摘要与英文摘要】 ============
    # +1：摘要标题格式 + 与下方正文空一行 + 摘要正文宋体小四、首行缩进 2 字符、1.5 倍行距
    kw = find_para(lambda r: r['text'].strip().startswith('关键词'))
    abs_body = [r for r in paras if r_abs and r['pi'] > i_abs and not r['blank']
                and (kw is None or r['pi'] < kw['pi'])]
    detail, sub = [], True
    if r_abs is None:
        sub, detail = False, ['未找到"摘  要"标题']
    else:
        t_ok, why = font_ok(r_abs['el'], r_abs['para'], ea=STRICT_SONGTI, size=22, bold=True)
        if not (t_ok and _align_of(doc, smap, r_abs['el']) == 'center'):
            sub = False
            detail.append('标题[%s 对齐=%s]' % (why, _align_of(doc, smap, r_abs['el'])))
        nxt = paras[i_abs + 1] if i_abs + 1 < len(paras) else None
        if not (nxt is not None and nxt['blank']):
            sub = False
            detail.append('标题与正文之间未空一行')
        bad_b = 0
        for r in abs_body:
            f_ok, _w = font_ok(r['el'], r['para'], ea=STRICT_SONGTI, size=12)
            if not (f_ok and abs(_indent_chars(doc, smap, r['el'], 'firstLine', dom_size(r)) - 2.0) <= 0.25
                    and _is_multiple_spacing(doc, smap, r['el'], 1.5)):
                bad_b += 1
        if not abs_body or bad_b:
            sub = False
            detail.append('摘要正文 %d/%d 段不合规' % (bad_b, len(abs_body)))
    add(1, '"摘  要"二字为宋体二号（22 磅）加粗居中并与下方正文空一行；中文摘要正文为宋体小四（12 磅）、首行缩进 2 个中文字符、1.5 倍行距',
        sub, '；'.join(detail) or '全部合规')

    # +1：关键词格式（"关键词"三字宋体小四加粗；3～6 个、中文逗号分隔、末尾无标点）
    detail, sub = [], True
    if kw is None:
        sub, detail = False, ['未找到"关键词"段落']
    else:
        heads = [r for r in _text_runs(kw['para']) if '关键词' in r.text]
        for r in heads:
            if not (_font_in(_resolve_effective_eastasia(doc, smap, r, kw['el']), STRICT_SONGTI)
                    and _near(_resolve_effective_size(doc, smap, r, kw['el']), 12)
                    and _resolve_effective_bold(doc, smap, r, kw['el'])):
                sub = False
                detail.append('"关键词"三字[字体=%s 字号=%s 加粗=%s]' % (
                    _resolve_effective_eastasia(doc, smap, r, kw['el']),
                    _resolve_effective_size(doc, smap, r, kw['el']),
                    _resolve_effective_bold(doc, smap, r, kw['el'])))
        payload = re.split(r'[：:]', kw['text'].strip(), 1)[-1].strip()
        words = [w for w in re.split(r'，', payload) if w.strip()]
        if not (3 <= len(words) <= 6):
            sub = False
            detail.append('以中文逗号分隔得 %d 个关键词' % len(words))
        if re.search(r'[；;,、]', payload):
            sub = False
            detail.append('存在非中文逗号分隔符')
        if payload and payload[-1] in '。；，、.;,':
            sub = False
            detail.append('末尾有标点"%s"' % payload[-1])
    add(1, '"关键词"三字为宋体小四（12 磅）加粗，关键词 3～6 个、以中文逗号分隔、末尾不加标点（分号等其他分隔符须改为中文逗号）',
        sub, '；'.join(detail) or '全部合规')

    # 英文题目与"Abstract"TNR 二号加粗居中、英文摘要正文 TNR 小四 1.5 倍行距
    # （承诺书已达标项，仅在减分项区核验是否被改坏）
    i_eabs = r_eabs['pi'] if r_eabs else len(paras)
    en_title = None
    for r in paras:
        if (kw['pi'] if kw else i_abs) < r['pi'] < i_eabs and not r['blank'] \
                and LATIN_RE.search(r['text']) and not CJK_RE.search(r['text']):
            en_title = r
    kwe = find_para(lambda r: r['text'].strip().lower().startswith('key word'))
    en_body = [r for r in paras if i_eabs < r['pi'] < (kwe['pi'] if kwe else len(paras))
               and not r['blank']]
    detail, sub = [], True
    for tag, rec in (('英文题目', en_title), ('Abstract', r_eabs)):
        if rec is None:
            sub = False
            detail.append('未找到%s' % tag)
            continue
        f_ok, why = font_ok(rec['el'], rec['para'], asc=STRICT_TNR, size=22, bold=True)
        if not (f_ok and _align_of(doc, smap, rec['el']) == 'center'):
            sub = False
            detail.append('%s[%s 对齐=%s]' % (tag, why, _align_of(doc, smap, rec['el'])))
    bad_b = 0
    for r in en_body:
        f_ok, _w = font_ok(r['el'], r['para'], asc=STRICT_TNR, size=12)
        if not (f_ok and _is_multiple_spacing(doc, smap, r['el'], 1.5)):
            bad_b += 1
    if not en_body or bad_b:
        sub = False
        detail.append('英文摘要正文 %d/%d 段不合规' % (bad_b, len(en_body)))
    pend['eabs'] = (sub, '；'.join(detail) or '全部合规')

    # +1："Key words"为 TNR 小四加粗，与英文摘要正文空一行，全部半角、逗号后空一格、末尾无标点
    detail, sub = [], True
    if kwe is None:
        sub, detail = False, ['未找到"Key words"段落']
    else:
        heads = [r for r in _text_runs(kwe['para']) if 'key word' in r.text.lower()]
        for r in heads:
            if not (_font_in(_resolve_effective_ascii(doc, smap, r, kwe['el']), STRICT_TNR)
                    and _near(_resolve_effective_size(doc, smap, r, kwe['el']), 12)
                    and _resolve_effective_bold(doc, smap, r, kwe['el'])):
                sub = False
                detail.append('"Key words"[字体=%s 字号=%s 加粗=%s]' % (
                    _resolve_effective_ascii(doc, smap, r, kwe['el']),
                    _resolve_effective_size(doc, smap, r, kwe['el']),
                    _resolve_effective_bold(doc, smap, r, kwe['el'])))
        prv = paras[kwe['pi'] - 1] if kwe['pi'] > 0 else None
        if not (prv is not None and prv['blank']):
            sub = False
            detail.append('与英文摘要正文之间未空一行')
        txt = kwe['text'].strip()
        if re.search(r'[，。；：、（）“”]', txt):
            sub = False
            detail.append('存在全角符号')
        if re.search(r',(?!\s)', txt):
            sub = False
            detail.append('逗号后未空一格')
        if txt and txt[-1] in '.;,':
            sub = False
            detail.append('末尾有标点"%s"' % txt[-1])
    add(1, '"Key words"为 Times New Roman 小四（12 磅）加粗，与英文摘要正文空一行，全部使用半角符号、逗号后空一格、末尾无标点',
        sub, '；'.join(detail) or '全部合规')

    # ============ 【目录】 ============
    toc_items = [(p, toc_level(p)) + toc_split(p) for p in toc_paras]
    t1 = [it for it in toc_items if it[1] == 1]
    t2 = [it for it in toc_items if it[1] == 2]
    t3 = [it for it in toc_items if it[1] == 3]

    # +3：一级目录条目为宋体三号（16 磅）加粗
    # 判定只取"条目文字"run：纯页码、制表符与前导点的 run 由域生成，不参与字体核验。
    def entry_runs(para):
        out = []
        for r in _text_runs(para):
            if re.match(r'^[\d\.\s·…\t IVXLCDMⅠⅡⅢⅣⅤⅥⅦⅧⅨⅩ]+$', r.text.strip()):
                continue
            out.append(r)
        return out

    def toc_font_ok(p, ea=None, size=None, bold=None):
        bad = []
        for r in entry_runs(p):
            if ea and CJK_RE.search(r.text):
                v = _resolve_effective_eastasia(doc, smap, r, p._p)
                if not _font_in(v, ea):
                    bad.append('中文字体=%s' % v)
            if size is not None:
                v = _resolve_effective_size(doc, smap, r, p._p)
                if not _near(v, size):
                    bad.append('字号=%s' % v)
            if bold is not None and _resolve_effective_bold(doc, smap, r, p._p) != bold:
                bad.append('加粗=%s' % (not bold))
        return (not bad, '；'.join(sorted(set(bad))[:3]))

    ok, fails = 0, []
    for p, _lv, _t, _pg in t1:
        f_ok, why = toc_font_ok(p, ea=STRICT_SONGTI, size=16, bold=True)
        ok += 1 if f_ok else 0
        if not f_ok:
            fails.append('%s[%s]' % (p.text.strip()[:10], why))
    add(3, '一级目录条目为宋体三号（16 磅）加粗', ratio_ok(ok, len(t1)),
        '合规 %d/%d；%s' % (ok, len(t1), '，'.join(fails[:3])))

    # +3：二级目录条目宋体小四并缩进 2 字符，三级目录条目宋体小四并缩进 4 字符
    ok, tot, fails = 0, len(t2) + len(t3), []
    for grp, want in ((t2, 2.0), (t3, 4.0)):
        for p, _lv, _t, _pg in grp:
            f_ok, why = toc_font_ok(p, ea=STRICT_SONGTI, size=12)
            left = _indent_chars(doc, smap, p._p, 'left', 12.0)
            ind_ok = abs(left - want) <= 0.3
            if f_ok and ind_ok:
                ok += 1
            else:
                fails.append('%s[%s%s]' % (p.text.strip()[:10], why,
                                           '' if ind_ok else ' 左缩进%.2f字符/应%d' % (left, want)))
    add(3, '二级目录条目为宋体小四（12 磅）并缩进 2 个中文字符，三级目录条目为宋体小四（12 磅）并缩进 4 个中文字符',
        ratio_ok(ok, tot), '合规 %d/%d；%s' % (ok, tot, '，'.join(fails[:3])))

    # +5：目录题序与文字随正文标题同步更新，且所标页码与正文实际页码一致
    # 判定：① 条目文字按顺序与正文标题一一对应；② 题序为"第一章／（一）／1、"体系；
    #       ③ 摘要、Abstract 页码为大写罗马数字；④ 正文条目页码为阿拉伯数字、自 1 起且不倒序。
    ROMAN_RE = re.compile(r'^([IVXLCDM]+|[ⅠⅡⅢⅣⅤⅥⅦⅧⅨⅩ]+)$')
    doc_titles = [r['norm'] for r in sorted(l1 + l2 + l3, key=lambda x: x['pi'])
                  if r_toc is None or r['pi'] != r_toc['pi']]
    detail, sub = [], True
    if not toc_items:
        sub, detail = False, ['未找到目录条目']
    else:
        pos, matched = 0, 0
        for _p, _lv, title, _pg in toc_items:
            nt = _norm(title)
            while pos < len(doc_titles) and doc_titles[pos] != nt:
                pos += 1
            if pos < len(doc_titles):
                matched += 1
                pos += 1
        if not ratio_ok(matched, len(toc_items)):
            sub = False
            detail.append('条目文字与正文标题同步 %d/%d' % (matched, len(toc_items)))
        num_bad = [title[:12] for _p, lv, title, _pg in toc_items
                   if (lv == 1 and re.match(r'^第', title.strip())
                       and not re.match(r'^第\s*[%s]+\s*章' % CN_NUM, title.strip()))
                   or (lv == 2 and not re.match(r'^（\s*[%s]+\s*）' % CN_NUM, title.strip()))
                   or (lv == 3 and not re.match(r'^\d+\s*、', title.strip()))]
        if num_bad:
            sub = False
            detail.append('题序未按"第一章／（一）／1、"重排：%s' % '，'.join(num_bad[:3]))
        front = [pg for _p, _lv, title, pg in toc_items if _norm(title).lower() in ('摘要', 'abstract')]
        if not front or not all(ROMAN_RE.match(pg or '') for pg in front):
            sub = False
            detail.append('摘要/Abstract 页码非大写罗马数字：%s' % front)
        arab = [pg for _p, lv, title, pg in toc_items
                if re.match(r'^第\s*[0-9%s]+\s*章' % CN_NUM, title.strip())]
        nums = [int(x) for x in arab if x.isdigit()]
        if len(nums) != len(arab) or not nums or nums[0] != 1 or nums != sorted(nums):
            sub = False
            detail.append('正文条目页码未自 1 起连续递增：%s' % arab)
    add(5, '目录条目的题序与文字随正文标题同步更新为"第一章／（一）／1、"体系，且所标页码与正文实际页码一致（摘要、Abstract 标 Ⅰ、Ⅱ，正文自 1 起）',
        sub, '；'.join(detail) or '全部合规')

    # ============ 【页眉页脚与页码】 ============
    def hf_runs(p):
        """页眉页脚中"承载可见字符"的 run：含 w:t 文本，或属于域代码（页码域）。"""
        out = []
        for r in p.runs:
            if (r.text and r.text.strip()) or r._element.find(qn('w:instrText')) is not None:
                out.append(r)
        return out

    def hf_font_ok(p, ea=None, asc=None, size=None):
        bad = []
        rs = hf_runs(p)
        if not rs:
            return (False, '无内容')
        for r in rs:
            txt = r.text or ''
            is_field = r._element.find(qn('w:instrText')) is not None
            if ea and CJK_RE.search(txt):
                v = _resolve_effective_eastasia(doc, smap, r, p._p)
                if not _font_in(v, ea):
                    bad.append('中文字体=%s' % v)
            if asc and (LATIN_RE.search(txt) or is_field):
                v = _resolve_effective_ascii(doc, smap, r, p._p)
                v2 = _resolve_effective_eastasia(doc, smap, r, p._p)
                if not (_font_in(v, asc) or _font_in(v2, asc)):
                    bad.append('页码字体=%s/%s' % (v, v2))
            if size is not None:
                v = _resolve_effective_size(doc, smap, r, p._p)
                if not _near(v, size):
                    bad.append('字号=%s' % v)
        return (not bad, '；'.join(sorted(set(bad))[:3]))

    def has_page_field(p):
        return bool(re.search(r'\bPAGE\b', _field_instr(p._p), re.I))

    MC_FALLBACK = ('{http://schemas.openxmlformats.org/markup-compatibility/2006}Fallback')

    def _in_fallback(el):
        node = el.getparent()
        while node is not None:
            if node.tag == MC_FALLBACK:
                return True
            node = node.getparent()
        return False

    def page_paras(si, kind='footer'):
        """真正承载页码的段落：页脚（含其中文本框）内带 PAGE 域的最内层段落，
        并剔除 mc:Fallback 分支中的重复副本。"""
        out = []
        for p in hf_paras(si, kind, shown_type(si)):
            if not has_page_field(p):
                continue
            if _in_fallback(p._p):
                continue
            inner = [q for q in p._p.iter(qn('w:p'))
                     if q is not p._p and re.search(r'\bPAGE\b', _field_instr(q), re.I)]
            if inner:
                continue                      # 外层包裹段落，实际页码在文本框内
            out.append(p)
        return out

    def page_field_runs(p):
        """页码域相关 run（域代码与域结果共享同一 rPr），页码格式以此为准。"""
        out = []
        for r_el in p._p.iter(qn('w:r')):
            if (r_el.find(qn('w:instrText')) is not None
                    or r_el.find(qn('w:fldChar')) is not None):
                out.append(Run(r_el, p))
        return out

    def page_font_ok(p, ea=None, asc=None, size=None):
        bad = []
        rs = page_field_runs(p)
        if not rs:
            return (False, '未取到页码域 run')
        for r in rs:
            if size is not None:
                v = _resolve_effective_size(doc, smap, r, p._p)
                if not _near(v, size):
                    bad.append('字号=%s' % v)
            v_ea = _resolve_effective_eastasia(doc, smap, r, p._p)
            v_as = _resolve_effective_ascii(doc, smap, r, p._p)
            if asc and not (_font_in(v_as, asc) or _font_in(v_ea, asc)):
                bad.append('页码字体=%s/%s' % (v_as, v_ea))
            if ea and not (_font_in(v_ea, ea) or _font_in(v_as, ea)):
                bad.append('页码字体=%s/%s' % (v_as, v_ea))
        return (not bad, '；'.join(sorted(set(bad))[:3]))

    def page_centered(si, p, kind='footer'):
        """页码居中：段落自身居中，或其所在文本框/外层页脚段落居中。"""
        if _align_of(doc, smap, p._p) == 'center':
            return True
        for q in hf_paras(si, kind, shown_type(si)):
            if _align_of(doc, smap, q._p) == 'center':
                return True
            for node in q._p.iter():
                tag = node.tag if isinstance(node.tag, str) else ''
                if tag.endswith('}align') and (node.text or '').strip() == 'center':
                    return True
        return False

    # 需要页眉的分节：摘要页起、且节内含一级标题
    hdr_targets = []
    for si in range(len(sects)):
        inside = [r for r in l1 if r['sect'] == si and r['pi'] >= i_abs]
        if inside:
            hdr_targets.append((si, inside[0]['norm']))

    # +3：自摘要页起各页统一添加页眉，页眉文字为该页所属章节名
    ok, fails = 0, []
    for si, name in hdr_targets:
        got = hf_text(si, 'header', shown_type(si))
        if got == name:
            ok += 1
        else:
            fails.append('第%d节 页眉"%s"应为"%s"' % (si, got[:12], name[:12]))
    add(3, '自摘要页起各页统一添加页眉，页眉文字为该页所属章节名（如"第一章  绪论""Abstract""目录""参考文献""致谢"）',
        ratio_ok(ok, len(hdr_targets)), '合规 %d/%d；%s' % (ok, len(hdr_targets), '，'.join(fails[:3])))

    # +1：页眉文字为黑体五号（10.5 磅）、居中
    ok, tot, fails = 0, 0, []
    for si, _name in hdr_targets:
        for p in hf_paras(si, 'header', shown_type(si)):
            if not (p.text or '').strip():
                continue
            tot += 1
            f_ok, why = hf_font_ok(p, ea=STRICT_HEITI, size=10.5)
            al = _align_of(doc, smap, p._p)
            if f_ok and al == 'center':
                ok += 1
            else:
                fails.append('第%d节[%s 对齐=%s]' % (si, why, al))
    add(1, '页眉文字为黑体五号（10.5 磅）、居中', ratio_ok(ok, tot),
        '合规 %d/%d；%s' % (ok, tot, '，'.join(fails[:3])))

    # 封皮页不带页眉与页码（承诺书已达标项，仅在减分项区核验是否被写入）
    # 判定：封皮所在节实际显示的页眉无文字，页脚既无文字也无 PAGE 域。
    cov_hdr = hf_text(0, 'header', shown_type(0))
    cov_ftr_paras = hf_paras(0, 'footer', shown_type(0))
    cov_ftr_txt = _norm(''.join(p.text for p in cov_ftr_paras))
    cov_has_page = any(has_page_field(p) for p in cov_ftr_paras)
    pend['cover_hf'] = (
        not cov_hdr and not cov_ftr_txt and not cov_has_page,
        '封皮页眉="%s"，页脚="%s"，页脚含PAGE域=%s' % (cov_hdr[:16], cov_ftr_txt[:16], cov_has_page))

    # +5：摘要、英文摘要、目录、正文各部分单独分节，摘要与目录页码为大写罗马数字，
    #     正文自"第一章"处重新起始编号为 1
    detail, sub = [], True
    key_sect = {}
    for tag, rec in (('摘要', r_abs), ('英文摘要', r_eabs), ('目录', r_toc), ('正文', r_chap1)):
        key_sect[tag] = rec['sect'] if rec else None
    vals = [v for v in key_sect.values() if v is not None]
    if len(set(vals)) != len(vals) or len(vals) < 4:
        sub = False
        detail.append('分节情况 %s' % key_sect)
    for tag in ('摘要', '目录'):
        si = key_sect[tag]
        fmt = sects[si]['fmt'] if si is not None and si < len(sects) else None
        if fmt != 'upperRoman':
            sub = False
            detail.append('%s节页码格式=%s' % (tag, fmt))
    si = key_sect['正文']
    if si is None or si >= len(sects):
        sub = False
        detail.append('未定位正文分节')
    else:
        if sects[si]['start'] != '1':
            sub = False
            detail.append('正文起始页码=%s' % sects[si]['start'])
        if sects[si]['fmt'] not in (None, 'decimal'):
            sub = False
            detail.append('正文页码格式=%s' % sects[si]['fmt'])
    add(5, '将摘要、英文摘要、目录、正文各部分单独分节，摘要与目录页码采用大写罗马数字（Ⅰ、Ⅱ……），正文自"第一章"处重新起始编号为 1',
        sub, '；'.join(detail) or '全部合规（共 %d 节）' % len(sects))

    # +1：摘要与目录页码为宋体小五（9 磅）、居中
    front_sects = [key_sect[t] for t in ('摘要', '英文摘要', '目录') if key_sect[t] is not None]
    ok, tot, fails = 0, 0, []
    for si in front_sects:
        for p in page_paras(si):
            tot += 1
            f_ok, why = page_font_ok(p, ea=STRICT_SONGTI, size=9)
            ctr = page_centered(si, p)
            if f_ok and ctr:
                ok += 1
            else:
                fails.append('第%d节[%s%s]' % (si, why, '' if ctr else ' 未居中'))
    add(1, '摘要与目录页码为宋体小五（9 磅）、居中', ratio_ok(ok, tot),
        '含页码的页脚 %d 个，合规 %d；%s' % (tot, ok, '，'.join(fails[:3])))

    # +1：正文页码为 Times New Roman 小五（9 磅）、居中，且书写格式为"第 X 页"
    body_sects = list(range(key_sect['正文'], len(sects))) if key_sect['正文'] is not None else []
    ok, tot, fails = 0, 0, []
    for si in body_sects:
        for p in page_paras(si):
            tot += 1
            f_ok, why = page_font_ok(p, asc=STRICT_TNR, size=9)
            ctr = page_centered(si, p)
            shape = re.search(r'第.*页', _norm(''.join(t.text or ''
                                                      for t in p._p.iter(qn('w:t'))))) is not None
            if f_ok and ctr and shape:
                ok += 1
            else:
                fails.append('第%d节[%s%s%s]' % (si, why, '' if ctr else ' 未居中',
                                                '' if shape else ' 非"第 X 页"格式'))
    add(1, '正文页码为 Times New Roman 小五（9 磅）、居中，且书写格式为"第 X 页"',
        ratio_ok(ok, tot), '含页码的页脚 %d 个，合规 %d；%s' % (tot, ok, '，'.join(fails[:3])))

    # ============ 【图表与代码】 ============
    def prev_block(bi):
        j = bi - 1
        while j >= 0:
            kind, el = blocks[j]
            if kind == 'tbl' or _has_drawing(el) or Paragraph(el, doc).text.strip():
                return (kind, el)
            j -= 1
        return (None, None)

    def next_block(bi):
        j = bi + 1
        while j < len(blocks):
            kind, el = blocks[j]
            if kind == 'tbl' or _has_drawing(el) or Paragraph(el, doc).text.strip():
                return (kind, el)
            j += 1
        return (None, None)

    def chapter_index_of(pi):
        idx = 0
        for k, cr in enumerate(chap_titles):
            if cr['pi'] <= pi:
                idx = k + 1
        return idx
    # +3：图名置于图片正下方，中文黑体五号、英文 Times New Roman 五号，图号按章连续编号
    ok, fails = 0, []
    seq = {}
    for r in figs:
        m = CAP_FIG_RE.match(r['text'].strip())
        ch, no = int(m.group(1)), int(m.group(2))
        want_ch = chapter_index_of(r['pi'])
        seq[ch] = seq.get(ch, 0) + 1
        kind, el = prev_block(r['bi'])
        pos = kind == 'p' and el is not None and _has_drawing(el)
        f_ok, why = font_ok(r['el'], r['para'], ea=STRICT_HEITI, asc=STRICT_TNR, size=10.5,
                            asc_alpha=True)
        num = (ch == want_ch and no == seq[ch])
        if pos and f_ok and num:
            ok += 1
        else:
            fails.append('%s[%s%s%s]' % (r['text'].strip()[:8], why,
                                         '' if pos else ' 上方非图片',
                                         '' if num else ' 图号应为图%d.%d' % (want_ch, seq[ch])))
    add(3, '图名置于图片正下方，中文为黑体五号（10.5 磅）、英文为 Times New Roman 五号（10.5 磅），'
           '图号按章连续编号（图 1.1、图 2.1……）',
        ratio_ok(ok, len(figs)), '合规 %d/%d；%s' % (ok, len(figs), '，'.join(fails[:4])))

    # +3：表题置于表格正上方，中文黑体五号、英文 Times New Roman 五号
    ok, fails = 0, []
    for r in tbl_caps:
        kind, el = next_block(r['bi'])
        pos = kind == 'tbl'
        f_ok, why = font_ok(r['el'], r['para'], ea=STRICT_HEITI, asc=STRICT_TNR, size=10.5,
                            asc_alpha=True)
        if pos and f_ok:
            ok += 1
        else:
            fails.append('%s[%s%s]' % (r['text'].strip()[:8], why, '' if pos else ' 下方非表格'))
    add(3, '表题置于表格正上方，中文为黑体五号（10.5 磅）、英文为 Times New Roman 五号（10.5 磅）',
        ratio_ok(ok, len(tbl_caps)), '合规 %d/%d；%s' % (ok, len(tbl_caps), '，'.join(fails[:4])))

    # +5：全部表格改为三线表——仅顶线（1.5 磅）、栏目线（0.75 磅）、底线（1.5 磅），无竖线
    # 判定：逐单元格解析生效边框（tcBorders 优先，回退 tblBorders 的外框/insideH/insideV）。
    ok, fails = 0, []
    for ti, tbl in enumerate(tables):
        tbl_el = tbl._tbl
        rows = tbl.rows
        nr = len(rows)
        bad = []
        for ri, row in enumerate(rows):
            cells = row.cells
            nc = len(cells)
            for ci, cell in enumerate(cells):
                tc = cell._tc
                fr, lr, fc, lc = ri == 0, ri == nr - 1, ci == 0, ci == nc - 1
                for edge in ('left', 'right'):
                    if _edge(tbl_el, tc, edge, fr, lr, fc, lc)[0]:
                        bad.append('第%d行存在竖线' % (ri + 1))
                        break
                top_line, top_w = _edge(tbl_el, tc, 'top', fr, lr, fc, lc)
                bot_line, bot_w = _edge(tbl_el, tc, 'bottom', fr, lr, fc, lc)
                if fr and not (top_line and _near(top_w, 1.5, 0.3)):
                    bad.append('顶线=%s/%s' % (top_line, top_w))
                if lr and not (bot_line and _near(bot_w, 1.5, 0.3)):
                    bad.append('底线=%s/%s' % (bot_line, bot_w))
                if ri == 0 and nr > 1 and not (bot_line and _near(bot_w, 0.75, 0.2)):
                    bad.append('栏目线=%s/%s' % (bot_line, bot_w))
                if 0 < ri < nr - 1 and bot_line:
                    bad.append('第%d行下多余横线' % (ri + 1))
        if bad:
            fails.append('表%d[%s]' % (ti + 1, '；'.join(sorted(set(bad))[:2])))
        else:
            ok += 1
    add(5, '全部 11 张表格改为三线表——仅保留上、中、下三条横线，全程无竖线，顶线与底线 1.5 磅、'
           '栏目线（表头下横线）0.75 磅',
        len(tables) >= 11 and ratio_ok(ok, len(tables)),
        '三线表 %d/%d；%s' % (ok, len(tables), '，'.join(fails[:4])))
    # +1：表头文字加粗
    ok, tot, fails = 0, 0, []
    for ti, tbl in enumerate(tables):
        if not tbl.rows:
            continue
        tot += 1
        bad = []
        for cell in tbl.rows[0].cells:
            for p in cell.paragraphs:
                if not p.text.strip():
                    continue
                f_ok, why = font_ok(p._p, p, bold=True)
                if not f_ok:
                    bad.append('%s[%s]' % (p.text.strip()[:6], why))
        if bad:
            fails.append('表%d %s' % (ti + 1, '，'.join(sorted(set(bad))[:2])))
        else:
            ok += 1
    add(1, '表头文字加粗', ratio_ok(ok, tot),
        '表头加粗 %d/%d；%s' % (ok, tot, '，'.join(fails[:4])))

    # 表内中文宋体五号（10.5 磅）、数字与英文 TNR 五号（承诺书已达标项，仅在减分项区核验）
    ok, tot, fails = 0, 0, []
    for ti, tbl in enumerate(tables):
        for row in tbl.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    if not p.text.strip():
                        continue
                    tot += 1
                    f_ok, why = font_ok(p._p, p, ea=STRICT_SONGTI, asc=STRICT_TNR, size=10.5)
                    if f_ok:
                        ok += 1
                    elif len(fails) < 4:
                        fails.append('表%d:%s[%s]' % (ti + 1, p.text.strip()[:6], why))
    pend['tbl_font'] = (ratio_ok(ok, tot),
                        '单元格段落合规 %d/%d；%s' % (ok, tot, '，'.join(fails[:4])))

    # +3：程序代码为 Times New Roman 五号、中文注释为宋体五号，行距 1.25 倍
    ok, fails = 0, []
    for r in codes:
        f_ok, why = font_ok(r['el'], r['para'], ea=STRICT_SONGTI, asc=STRICT_TNR, size=10.5)
        sp_ok = _is_multiple_spacing(doc, smap, r['el'], 1.25)
        if f_ok and sp_ok:
            ok += 1
        elif len(fails) < 4:
            line, rule = _line_spacing(doc, smap, r['el'])
            fails.append('%s[%s%s]' % (r['text'].strip()[:10], why,
                                       '' if sp_ok else ' 行距=%s/%s' % (line, rule)))
    add(3, '程序代码为 Times New Roman 五号（10.5 磅）、代码中的中文注释为宋体五号（10.5 磅），'
           '代码段行距为 1.25 倍',
        len(codes) > 0 and ratio_ok(ok, len(codes)),
        '代码段落合规 %d/%d；%s' % (ok, len(codes), '，'.join(fails[:4])))
    # ============ 减分项：封皮页 ============
    cov = [r for r in paras if r['pi'] in set(cover_range)]
    INFO_RE = re.compile(r'^(院\s*校|专\s*业|学\s*号|姓\s*名|指导教师|教师职称|完成日期)\s*[:：]')
    SCHOOL = '澄云信息科技学院'
    cov_school = [r for r in cov if r['norm'] == SCHOOL]
    cov_info = [r for r in cov if INFO_RE.match(r['text'].strip())]

    def cov_check(r, size, bold, after=None, before=None, line=None):
        """封皮单行版式核验，返回不合规描述（空串表示合规）。"""
        bad = []
        f_ok, why = font_ok(r['el'], r['para'], ea=STRICT_SONGTI, size=size, bold=bold)
        if not f_ok:
            bad.append(why)
        al = _align_of(doc, smap, r['el'])
        if al != 'center':
            bad.append('对齐=%s' % al)
        if line is not None and not _is_multiple_spacing(doc, smap, r['el'], line):
            bad.append('行距=%s/%s' % _line_spacing(doc, smap, r['el']))
        for which, want in (('after', after), ('before', before)):
            if want is not None:
                got = _space_pt(doc, smap, r['el'], which)
                if not _near(got, want, 0.6):
                    bad.append('段%s=%s磅' % ('后' if which == 'after' else '前', got))
        return '；'.join(bad)

    # -1：封皮首行"澄云信息科技学院"宋体 24 磅加粗、居中、1.35 倍行距、段后 6 磅
    why = cov_check(cov_school[0], 24, True, after=6, line=1.35) if cov_school else '未找到该行'
    add(-1, '封皮首行"澄云信息科技学院"不满足宋体 24 磅加粗、居中对齐、1.35 倍行距、段后 6 磅'
            '（字号被改为 22 磅、段后被改为 20 磅等均计）', bool(why), why or '该行版式保持不变')

    # -1："毕业设计（论文）"宋体 26 磅加粗、居中、段后 18 磅
    r_deg = next((r for r in cov if r['norm'] == '毕业设计（论文）'), None)
    why = cov_check(r_deg, 26, True, after=18) if r_deg else '未找到该行'
    add(-1, '封皮"毕业设计（论文）"不满足宋体 26 磅加粗、居中对齐、段后 18 磅'
            '（字号被改为 22 磅、段后被改为 60 磅等均计）', bool(why), why or '该行版式保持不变')

    # -1：七行信息行宋体三号、居中、1.35 倍行距、段前 4 磅、制表符对齐、日期连排
    bad = []
    if len(cov_info) < 7:
        bad.append('信息行仅 %d 行（应为 7 行）' % len(cov_info))
    for r in cov_info:
        w = cov_check(r, 16, None, before=4, line=1.35)
        if '\t' not in r['text']:
            w = (w + '；' if w else '') + '标签与内容之间无制表符'
        if r['text'].strip().startswith('完成日期') and not re.search(r'\d{4}年\d{1,2}月\d{1,2}日',
                                                                  r['text'].replace(' ', '')
                                                                  .replace('　', '')):
            w = (w + '；' if w else '') + '日期未连排书写'
        if w:
            bad.append('%s[%s]' % (r['text'].strip()[:6], w))
    add(-1, '封皮"论文题目／学院／专业／学号／姓名／指导教师／职称／日期"等信息行不满足宋体三号（16 磅）、'
            '居中对齐、1.35 倍行距、段前 4 磅，且不满足标签与内容之间以制表符对齐、日期连排书写为'
            '"2026年5月18日"（字号被改为 14 磅、制表符被改为空格、日期被改写为"2026 年 5 月 18 日"均计）',
        bool(bad), '，'.join(bad[:3]) or '7 行信息行版式与内容保持不变')
    # -1：封皮末尾中英文校名两行为三号（16 磅）、居中
    r_en = next((r for r in cov if 'Chengyun Institute' in r['text']), None)
    r_cn = cov_school[-1] if len(cov_school) > 1 else None
    bad = []
    for tag, r in (('中文校名', r_cn), ('英文校名', r_en)):
        if r is None:
            bad.append('%s缺失' % tag)
            continue
        f_ok, why = font_ok(r['el'], r['para'], size=16)
        al = _align_of(doc, smap, r['el'])
        if not f_ok or al != 'center':
            bad.append('%s[%s 对齐=%s]' % (tag, why, al))
    add(-1, '封皮末尾"澄云信息科技学院"与"Chengyun Institute of Information Technology"两行不满足'
            '三号（16 磅）、居中对齐（字号被改为 14 磅计）',
        bool(bad), '，'.join(bad[:2]) or '两行字号与对齐保持不变')

    # -3：删除封皮中位于校名下方的图形对象（宽 5666740 EMU、高 539750 EMU）
    EMU_W, EMU_H = 5666740.0, 539750.0
    exts = []
    for r in cov:
        for el in r['el'].iter():
            tag = el.tag if isinstance(el.tag, str) else ''
            if not (tag.endswith('}ext') or tag.endswith('}extent')):
                continue
            cx, cy = el.get('cx'), el.get('cy')
            if cx and cy:
                try:
                    exts.append((float(cx), float(cy)))
                except ValueError:
                    pass
    kept = [e for e in exts if abs(e[0] - EMU_W) <= EMU_W * 0.02 and abs(e[1] - EMU_H) <= EMU_H * 0.02]
    add(-3, '删除封皮中位于校名下方的图形对象（宽 5666740 EMU、高 539750 EMU 的图形/图片），'
            '导致该页图形缺失',
        not kept, '封皮图形尺寸清单=%s' % ([(int(a), int(b)) for a, b in exts[:4]] or '无图形对象'))

    # -1：论文题目段落单独成页、黑体二号、居中、段前段后各 12 磅
    r_title = next((r for r in cov if r['text'].strip()
                    and not INFO_RE.match(r['text'].strip())
                    and r['norm'] != SCHOOL and r['norm'] != '毕业设计（论文）'
                    and 'Chengyun Institute' not in r['text']), None)
    bad = []
    if r_title is None:
        bad.append('未找到单独成页的论文题目段落')
    else:
        if not starts_new_page(r_title):
            bad.append('未单独成页（并入封皮页）')
        f_ok, why = font_ok(r_title['el'], r_title['para'], ea=STRICT_HEITI, size=22)
        al = _align_of(doc, smap, r_title['el'])
        if not f_ok or al != 'center':
            bad.append('%s 对齐=%s' % (why, al))
        for which, cn in (('before', '前'), ('after', '后')):
            got = _space_pt(doc, smap, r_title['el'], which)
            if not _near(got, 12, 0.6):
                bad.append('段%s=%s磅' % (cn, got))
    add(-1, '将单独占一页的论文题目段落"基于Spark的园区能耗分析与可视化系统"并入封皮页，'
            '不满足其单独成页、黑体二号（22 磅）、居中对齐、段前段后各 12 磅',
        bool(bad), '；'.join(bad[:3]) or '论文题目页保持不变')

    # -3：封皮页被写入页眉文字或被插入页码域
    _ok, _dt = pend['cover_hf']
    add(-3, '封皮页不满足"不带页眉与页码"（在封皮页写入页眉文字、或为封皮页页脚插入页码域均计）',
        not _ok, _dt)

    # ============ 减分项：页面设置 ============
    # -1：页边距上下左右 25mm、页眉距边界 6mm、页脚距边界 8mm
    MM = 1440.0 / 25.4          # 1 毫米 = 56.69 缇
    WANT_MAR = (('top', 25), ('bottom', 25), ('left', 25), ('right', 25),
                ('header', 6), ('footer', 8))
    bad = []
    for si, s in enumerate(sects):
        for k, mm in WANT_MAR:
            v = s['mar'].get(k)
            try:
                got = float(v) if v is not None else None
            except ValueError:
                got = None
            if got is None or abs(got - mm * MM) > 20:
                bad.append('第%d节 %s=%s mm' % (si, k, round(got / MM, 1) if got else None))
    add(-1, '页面设置不满足上 25mm、下 25mm、左 25mm、右 25mm 页边距，页眉距边界 6mm、页脚距边界 8mm'
            '（左边距被改为 30mm、页眉距被改为 15mm、页脚距被改为 17.5mm 等均计）',
        bool(bad), '，'.join(bad[:4]) or '各节页边距与页眉页脚距保持不变')

    # ============ 减分项：正文与摘要既有版式（承诺书要求中已达标的部分，改动即为破坏） ============
    # -3：正文段落中文不再是宋体小四（12 磅）
    _ok, _dt = pend['body_ea']
    add(-3, '正文段落中文不满足宋体小四（12 磅）（被替换为"Noto Serif CJK SC"等第三方字体、'
            '或字号被改为四号均计）', not _ok, _dt)

    # -3：正文段落中的英文、数字及公式的英文部分不再是 Times New Roman 小四（12 磅）
    _ok, _dt = pend['body_lat']
    add(-3, '正文段落中的英文、数字及公式的英文部分不满足 Times New Roman 小四（12 磅）',
        not _ok, _dt)

    # -3：正文段落首行缩进 2 个中文字符、两端对齐被破坏
    _ok, _dt = pend['body_ind']
    add(-3, '正文段落不满足首行缩进 2 个中文字符、两端对齐（缩进被清零、缩进单位由"字符"改为'
            '厘米绝对值如 0.85 厘米、对齐被改为左对齐均计）', not _ok, _dt)

    # -3：正文段落 1.5 倍行距、段前段后 0 行被破坏
    _ok, _dt = pend['body_line']
    add(-3, '正文段落不满足 1.5 倍行距、段前 0 行、段后 0 行（行距被改为单倍行距、'
            '段后被改为 2 磅等均计）', not _ok, _dt)

    # -1：英文题目与"Abstract"、英文摘要正文版式被破坏
    _ok, _dt = pend['eabs']
    add(-1, '英文题目与"Abstract"两行不满足 Times New Roman 二号（22 磅）加粗、居中，'
            '英文摘要正文不满足 Times New Roman 小四（12 磅）、1.5 倍行距', not _ok, _dt)

    # -1：表内中文宋体五号、表内数字与英文 TNR 五号被破坏
    _ok, _dt = pend['tbl_font']
    add(-1, '表内中文不满足宋体五号（10.5 磅）、表内数字与英文不满足 Times New Roman 五号（10.5 磅）',
        not _ok, _dt)

    # ============ 减分项：字体与样式体系 ============
    # -3：参考文献 10 条条目与致谢正文段落的中文被替换为第三方字体
    ref_thx = []
    for anchor, stop in ((r_ref, r_thx), (r_thx, None)):
        if anchor is None:
            continue
        hi = stop['pi'] if stop is not None else len(paras)
        ref_thx += [r for r in paras if anchor['pi'] < r['pi'] < hi and not r['blank']]
    ok, fails = 0, []
    for r in ref_thx:
        f_ok, why = font_ok(r['el'], r['para'], ea=STRICT_SONGTI, size=12)
        if f_ok:
            ok += 1
        elif len(fails) < 4:
            fails.append('%s[%s]' % (r['text'].strip()[:10], why))
    add(-3, '参考文献 10 条条目与致谢正文段落的中文不满足宋体小四（12 磅）'
            '（被替换为"Noto Serif CJK SC""Noto Sans CJK SC"等第三方字体计）',
        not ratio_ok(ok, len(ref_thx)),
        '合规 %d/%d；%s' % (ok, len(ref_thx), '，'.join(fails[:4])))

    # -3：一级部分标题被降级为正文样式，且 我的正文／toc 1／toc 2／toc 3 等样式定义被删除
    ol_bad = ['%s(大纲级别=%s)' % (r['norm'][:8], r['ol'])
              for r in (r_abs, r_eabs, r_toc, r_ref, r_thx) if r is not None and r['ol'] != 0]
    style_names = set()
    for st in smap.values():
        nm = st.find(qn('w:name'))
        style_names.add(((nm.get(qn('w:val')) if nm is not None else '') or '').strip().lower())
    miss = [n for n in ('我的正文', 'toc 1', 'toc 2', 'toc 3') if n.lower() not in style_names]
    add(-3, '将"摘  要""Abstract""目  录""参考文献""致  谢"等段落由"标题 1"样式降级为正文样式，'
            '并删除文档中"我的正文""toc 1／toc 2／toc 3""正文首行缩进 2"等既有段落样式定义，'
            '导致大纲级别与导航结构丢失',
        len(ol_bad) >= 3 or bool(miss),
        '大纲级别异常：%s；缺失样式：%s' % ('，'.join(ol_bad[:4]) or '无', '、'.join(miss) or '无'))
    # ============ 减分项：文字内容 ============
    # -3：为参考文献条目添加"[1] "至"[10] "序号前缀
    refs = [r for r in paras if i_ref < r['pi'] < i_thx and not r['blank']]
    numbered = [r['text'].strip()[:10] for r in refs if re.match(r'^\[\s*\d+\s*\]', r['text'].strip())]
    add(-3, '为参考文献 10 条条目逐条添加"[1] "至"[10] "的序号前缀，属非格式性的文字内容改动',
        len(numbered) >= 2, '带"[n]"前缀的条目 %d/%d：%s' % (
            len(numbered), len(refs), '，'.join(numbered[:3]) or '无'))

    # -3：删除程序代码各行行首的缩进空白
    indented = [r for r in codes if r['text'][:1] in (' ', '\t', '　')]
    cont = [r for r in codes if r['text'].strip()[:1] == '.']
    flat_cont = [r for r in cont if r['text'][:1] not in (' ', '\t', '　')]
    add(-3, '删除程序代码各行行首的缩进空白（如"    df = df.filter("quality_flag = 1")"'
            '"        .agg(sum("delta_value")…"被改为全部顶格），破坏代码的层级结构',
        len(codes) > 0 and (not indented or (cont and len(flat_cont) == len(cont))),
        '代码段落 %d 行，行首带缩进 %d 行，续行(以"."开头) %d 行中顶格 %d 行' % (
            len(codes), len(indented), len(cont), len(flat_cont)))

    # -3：删除标题与标题之间、标题与正文之间的空行分隔
    _ok, _dt = pend['title_gap']
    blanks = [r for r in paras if i_abs <= r['pi'] < i_thx and r['blank']]
    add(-3, '删除标题与标题之间、标题与正文之间的空行分隔（空行段落被删除且未以段后间距补足）',
        not _ok, '%s；摘要至致谢区间空段落共 %d 个' % (_dt, len(blanks)))

    # -1：致谢标题"致  谢"被改写为"致 谢"
    thx_txt = r_thx['text'].strip() if r_thx else ''
    add(-1, '将致谢标题"致  谢"（两字之间两个空格）改写为"致 谢"（单空格），改动了标题文字',
        r_thx is None or not re.match(r'^致[\s　]{2}谢$', thx_txt),
        '致谢标题文字="%s"' % thx_txt)
    # ============ 减分项：其他多余改动 ============
    # -1：可自动更新的目录域被改为手工录入的纯文本段落
    add(-1, '将可自动更新的目录域改为手工录入的纯文本段落，导致目录无法随正文更新、条目超链接全部丢失',
        not toc_field, '文档中%s TOC 域' % ('存在' if toc_field else '不存在'))

    # -1：表格单元格内段落被统一改为居中对齐
    ctr, tot = 0, 0
    for tbl in tables:
        for ri, row in enumerate(tbl.rows):
            if ri == 0:
                continue                      # 表头居中属常规排版，不计入
            for cell in row.cells:
                for p in cell.paragraphs:
                    if not p.text.strip():
                        continue
                    tot += 1
                    if _align_of(doc, smap, p._p) == 'center':
                        ctr += 1
    add(-1, '表格单元格内段落不满足左对齐（被统一改为居中对齐计）',
        tot > 0 and (float(ctr) / float(tot)) > 0.5,
        '表体单元格段落居中 %d/%d' % (ctr, tot))

    # -1：图名与表题段落被加粗或被设置了段前段后间距
    caps = sorted(figs + tbl_caps, key=lambda r: r['pi'])
    ok, fails = 0, []
    for r in caps:
        bad = []
        if any(_resolve_effective_bold(doc, smap, run, r['el']) for run in _text_runs(r['para'])):
            bad.append('加粗')
        for which, cn in (('before', '前'), ('after', '后')):
            got = _space_pt(doc, smap, r['el'], which)
            if not _near(got, 0, 0.6):
                bad.append('段%s=%s磅' % (cn, got))
        if bad:
            fails.append('%s[%s]' % (r['text'].strip()[:8], '；'.join(bad)))
        else:
            ok += 1
    add(-1, '图名与表题段落不满足不加粗、段前段后 0 磅（被加粗、或被设置为段前 6 磅、段后 6 磅均计）',
        not ratio_ok(ok, len(caps)), '合规 %d/%d；%s' % (ok, len(caps), '，'.join(fails[:3])))
    # -1：中文摘要正文段落首行缩进不再是 2 个中文字符
    ok, fails = 0, []
    for r in abs_body:
        got = _indent_chars(doc, smap, r['el'], 'firstLine', dom_size(r))
        unit, uval = _indent_unit(doc, smap, r['el'], 'firstLine')
        if abs(got - 2.0) <= 0.25 and unit == 'chars':
            ok += 1
        elif len(fails) < 3:
            fails.append('%s[首行缩进=%s]' % (
                r['text'].strip()[:8],
                '%.2f字符' % got if unit == 'chars' else
                ('%.2f厘米（非字符单位）' % uval if unit == 'abs' else '未设置')))
    add(-1, '中文摘要正文段落不满足首行缩进 2 个中文字符（缩进被清零、被改为 1 个字符、'
            '或缩进单位由"字符"改为厘米绝对值如 0.85 厘米均计）',
        not ratio_ok(ok, len(abs_body)),
        '摘要正文合规 %d/%d；%s' % (ok, len(abs_body), '，'.join(fails)))

    # -1："关键词：……"段落不再是两端对齐、段前 0 磅
    bad = []
    if kw is None:
        bad.append('未找到"关键词"段落')
    else:
        al = _align_of(doc, smap, kw['el'])
        if al not in ('both', 'justify'):
            bad.append('对齐=%s' % al)
        got = _space_pt(doc, smap, kw['el'], 'before')
        if not _near(got, 0, 0.6):
            bad.append('段前=%s磅' % got)
    add(-1, '"关键词：……"段落不满足两端对齐、段前 0 磅（被改为左对齐、段前 12 磅计）',
        bool(bad), '；'.join(bad) or '关键词段落对齐与段前间距保持不变')

    # -1：各节实际显示的页眉中残留拼音乱码或与本节不符的章节名
    # 收窄判定范围：只核验各节实际渲染的页眉——未开启 w:titlePg 的节，其 first 类型页眉
    #               部件在 Word 中不显示，包内这类废弃部件不计入本条。
    PINYIN_RE = re.compile(r"[a-z]{2,}'[a-z]{2,}")
    bad = []
    for si in range(len(sects)):
        got = hf_text(si, 'header', shown_type(si))
        if got and PINYIN_RE.search(got):
            bad.append('第%d节页眉含拼音乱码"%s"' % (si, got[:16]))
    for si, nm in hdr_targets:
        got = hf_text(si, 'header', shown_type(si))
        if got and got != nm:
            bad.append('第%d节页眉"%s"应为"%s"' % (si, got[:12], nm[:12]))
    add(-1, '在各节实际显示的页眉中写入拼音乱码或与本节不符的章节名（如"第二章  xi\'tong\'sh"、'
            '致谢节页眉写成"第三章  系统设计与功能实现"），使该节页眉内容错误',
        bool(bad), '，'.join(bad[:3]) or '各节显示的页眉内容与所属章节一致')

    # ---------------- 汇总 ----------------
    if len(results) != len(_MAX_DELTAS_ORDER):
        raise RuntimeError('细则条目数不一致：实际 %d 条，预期 %d 条' % (
            len(results), len(_MAX_DELTAS_ORDER)))
    items = []
    for raw, max_delta in zip(results, _MAX_DELTAS_ORDER):
        # raw 形如 "+3：<细则原文>" / "-0：<细则原文>": 前缀仅用于反解 delta,
        # 写入 dim2_items 的 "rule" 只保留"："之后的细则原文。
        head, _, rule = raw.partition('：')
        delta = int(head)
        items.append({"rule": rule, "max_delta": max_delta, "delta": delta,
                      "hit": delta != 0})
    result["dim2_items"] = items
    result["max_score"] = MAX_SCORE_TOTAL
    result["total_score"] = sum(it["delta"] for it in items) if result["dim1_pass"] else 0
    result["status"] = "ok"
    return result


if __name__ == "__main__":
    # 本地调试入口：evaluate 只接收"脚本所在目录的路径"，由脚本自己在该目录里定位并
    # 打开被评估的 .docx。命令行第 1 个参数可覆盖目录路径；未指定时使用脚本自身所在目录。
    if len(sys.argv) >= 2:
        _dir = sys.argv[1]
    else:
        _dir = os.path.dirname(os.path.abspath(__file__))
    _out = evaluate(_dir)
    print(json.dumps(_out, ensure_ascii=False, indent=2))

