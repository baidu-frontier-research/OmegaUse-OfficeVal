#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# pyright: reportUnknownMemberType=false, reportUnknownVariableType=false, reportUnknownArgumentType=false, reportUnknownParameterType=false, reportMissingParameterType=false, reportPrivateUsage=false, reportUnknownLambdaType=false, reportGeneralTypeIssues=false

import os
import re
import sys
import json
import math
import shutil
import tempfile
from docx import Document
from docx.text.paragraph import Paragraph
from docx.enum.text import WD_ALIGN_PARAGRAPH
try:
    _VERIFIER_DIR = os.path.dirname(os.path.abspath(__file__))
    if _VERIFIER_DIR not in sys.path:
        sys.path.insert(0, _VERIFIER_DIR)
    import pymupdf_compat as fitz  # 已移除 PyMuPDF(AGPL)，改用 pdfplumber+pypdfium2 兼容层
except ImportError:
    fitz = None
from docx.oxml.ns import qn

# 说明：按接口统一约定，本模块不修改全局 sys.stdout，也不 print 主结果。
# 主结果统一由 evaluate(dir_path) -> dict 返回；入参为"脚本所在目录的路径"，
# 由脚本自己在该目录内定位并打开被评估的 .docx。本文件仅在 __main__ 下作为
# 本地调试入口，才把 dict 序列化为 JSON 打印。
#
# 评测对象：按"齐鲁启明继续教育学院论文模板"的格式要求，修改
#           "幼儿园自然探究活动课程设计"中的目录。
# 维度二含 20 项加分项（指令下达的任务）与 26 项减分项（指令未涉及、
#           一经改动即为破坏的项），减分项的 max_delta 为负值。

SCRIPT_ID = "007"
# 维度二全部"加分项"max_delta 之和；用于维度一未通过 / 脚本异常时兜底。
# 减分项不计入 max_score（未触发即为 0，触发则从 total_score 中扣减）。
MAX_SCORE_TOTAL = 56

# 维度二所有评分项按"结果追加顺序"的 max_delta 列表；末端把 results 里的
# "+N：..." / "-N：..." 字符串反解为结构化 dim2_items 时用。顺序须与
# _evaluate_impl 中 results.append(...) 的顺序严格一致；共 46 项。
_MAX_DELTAS_ORDER = [
    # ===== 加分项（20 项，累加 = MAX_SCORE_TOTAL = 56）=====
    5,   # 01 目录由目录域/目录样式承载
    5,   # 02 正文各级标题已设大纲级别
    5,   # 03 三级层级完整收录（条目 ≥45）
    5,   # 04 点线前导符 + 右对齐制表位落在版心右边界
    5,   # 05 条目文字与正文标题一致
    3,   # 06 目录标题"目  录"格式
    3,   # 07 一级条目：黑体小四/左对齐/1.5倍/不加粗
    3,   # 08 二级条目：宋体小四/左对齐/1.5倍/无缩进
    1,   # 09 三级条目：宋体小四/左对齐/1.5倍/缩进1字符
    3,   # 10 层级缩进递进且同级一致
    1,   # 11 无"参考文献""附  录"条目
    1,   # 12 收录"总结""致谢"
    3,   # 13 无手工省略号/句点串填充
    3,   # 14 每条行末有阿拉伯数字页码
    3,   # 15 页码与实际所在页一致、只增不减
    3,   # 16 条目单行显示
    1,   # 17 条目段前段后为0或全表统一
    1,   # 18 仅收录标题级文字
    1,   # 19 目录位置与独立起页
    1,   # 20 "三"章条目为完整标题
    # ===== 减分项（26 项，max_delta 为负）=====
    -5,  # 21 封面校徽图形/图片元素
    -5,  # 22 正文段落完整性（无整段删除/重复/错序）
    -5,  # 23 分节结构与分段页码编排
    -3,  # 24 A4 纸张与页边距
    -3,  # 25 前置部分大写罗马数字页码
    -3,  # 26 正文阿拉伯数字页码从 1 起编
    -3,  # 27 页眉/页脚存在与格式
    -1,  # 28 "学校代码："行
    -1,  # 29 封面"成人高等教育毕业论文（设计）"
    -1,  # 30 封面题目两行
    -1,  # 31 "启明学院毕业论文作者声明"
    -1,  # 32 作者声明两段正文
    -1,  # 33 作者签名/指导教师签名两行
    -1,  # 34 "摘  要"标题
    -1,  # 35 摘要正文
    -1,  # 36 关键词行
    -1,  # 37 "Abstract"标题
    -1,  # 38 Abstract 正文
    -1,  # 39 "Keywords:"行
    -3,  # 40 正文四个章标题
    -1,  # 41 "总结""致谢"标题
    -3,  # 42 正文二级标题
    -1,  # 43 正文叙述段落字号/行距
    -1,  # 44 "•"项目符号
    -3,  # 45 批注/修订/高亮/底纹/彩色字体
    -1,  # 46 致谢正文
]
assert sum(d for d in _MAX_DELTAS_ORDER if d > 0) == MAX_SCORE_TOTAL

# ======================= 一、格式继承链解析 =======================
# 与办公软件对齐：Word/WPS 里"看到的"格式是 直接设定 → 段落样式(沿 basedOn 上溯)
# → docDefaults 三层继承的结果。本套目录既可能把格式写在段落上（手工目录），
# 也可能写在 toc 1/toc 2/toc 3 样式上（自动目录域），因此所有判定一律走继承链，
# 不直接读 python-docx 的 run.font.* / paragraph_format.*（后者只看直接设定）。

def _style_by_id(doc, style_id):
    for st in doc.styles.element.findall(qn('w:style')):
        if st.get(qn('w:styleId')) == style_id:
            return st
    return None


def _para_style_id(para):
    if para is None:
        return None
    pPr = para._element.find(qn('w:pPr'))
    if pPr is None:
        return None
    pStyle = pPr.find(qn('w:pStyle'))
    return pStyle.get(qn('w:val')) if pStyle is not None else None


def _style_chain(doc, style_id):
    """样式自身 + basedOn 上溯链。"""
    out = []
    seen = set()
    while style_id and style_id not in seen:
        seen.add(style_id)
        st = _style_by_id(doc, style_id)
        if st is None:
            break
        out.append(st)
        base = st.find(qn('w:basedOn'))
        style_id = base.get(qn('w:val')) if base is not None else None
    return out


def _doc_defaults(doc, kind):
    """kind='rPr' 或 'pPr'，取 docDefaults 里的默认属性容器。"""
    dd = doc.styles.element.find(qn('w:docDefaults'))
    if dd is None:
        return None
    wrap = dd.find(qn('w:rPrDefault') if kind == 'rPr' else qn('w:pPrDefault'))
    if wrap is None:
        return None
    return wrap.find(qn('w:' + kind))


def _rpr_chain(doc, run):
    """run 的字符属性继承链：run 直接设定 → 段落样式链 → docDefaults。"""
    chain = []
    if run is not None and run._element.rPr is not None:
        chain.append(run._element.rPr)
    sid = _para_style_id(run._parent) if run is not None else None
    for st in _style_chain(doc, sid):
        rpr = st.find(qn('w:rPr'))
        if rpr is not None:
            chain.append(rpr)
    dd = _doc_defaults(doc, 'rPr')
    if dd is not None:
        chain.append(dd)
    return chain


def _ppr_chain(doc, para):
    """段落属性继承链：段落直接设定 → 段落样式链 → docDefaults。"""
    chain = []
    pPr = para._element.find(qn('w:pPr'))
    if pPr is not None:
        chain.append(pPr)
    for st in _style_chain(doc, _para_style_id(para)):
        p = st.find(qn('w:pPr'))
        if p is not None:
            chain.append(p)
    dd = _doc_defaults(doc, 'pPr')
    if dd is not None:
        chain.append(dd)
    return chain


def _chain_attr(chain, tag, attr):
    """沿继承链取第一个显式给出该属性的值（逐属性继承，与 Word 一致）。"""
    for el in chain:
        node = el.find(qn(tag))
        if node is not None:
            v = node.get(qn(attr))
            if v is not None:
                return v
    return None


def _chain_node(chain, tag):
    for el in chain:
        node = el.find(qn(tag))
        if node is not None:
            return node
    return None

# ---------- 字符级（run）有效属性 ----------
# 注：东亚字体若只由主题（eastAsiaTheme=minorEastAsia）给出，视为"未显式指定"，
#     按 Word 中文默认字体（宋体）处理——即 _is_songti(None) 为真、_is_heiti(None) 为假。

def _eff_eastasia(doc, run):
    return _chain_attr(_rpr_chain(doc, run), 'w:rFonts', 'w:eastAsia')


def _eff_ascii(doc, run):
    chain = _rpr_chain(doc, run)
    return (_chain_attr(chain, 'w:rFonts', 'w:ascii')
            or _chain_attr(chain, 'w:rFonts', 'w:hAnsi'))


def _eff_size_pt(doc, run):
    v = _chain_attr(_rpr_chain(doc, run), 'w:sz', 'w:val')
    try:
        return float(v) / 2.0 if v is not None else None
    except ValueError:
        return None


def _eff_bold(doc, run):
    for el in _rpr_chain(doc, run):
        b = el.find(qn('w:b'))
        if b is not None:
            return (b.get(qn('w:val')) or 'true') not in ('0', 'false', 'off')
    return False


def _eff_color(doc, run):
    """返回 (w:val, w:themeColor)。"""
    chain = _rpr_chain(doc, run)
    node = _chain_node(chain, 'w:color')
    if node is None:
        return (None, None)
    return (node.get(qn('w:val')), node.get(qn('w:themeColor')))


def _is_songti(name):
    if not name:
        return True   # 未显式指定 → Word 中文默认宋体
    return any(k in name for k in ('宋体', 'SimSun', 'NSimSun', 'Song'))


def _is_heiti(name):
    if not name:
        return False
    return any(k in name for k in ('黑体', 'SimHei', 'Hei'))


# ---------- 段落级有效属性 ----------

def _eff_jc(doc, para):
    """返回 'left'/'center'/'right'/'both'/None（None = 未指定，Word 按左/两端）。"""
    return _chain_attr(_ppr_chain(doc, para), 'w:jc', 'w:val')


def _eff_outline(doc, para):
    v = _chain_attr(_ppr_chain(doc, para), 'w:outlineLvl', 'w:val')
    try:
        return int(v) if v is not None else None
    except ValueError:
        return None


def _eff_spacing(doc, para, attr):
    """取 w:spacing 上某属性（before/after/beforeLines/afterLines/line/lineRule）。"""
    for el in _ppr_chain(doc, para):
        sp = el.find(qn('w:spacing'))
        if sp is not None:
            v = sp.get(qn('w:' + attr))
            if v is not None:
                return v
    return None


def _eff_line_multiple(doc, para):
    """行距倍数；固定值/最小值行距返回 None。"""
    line = _eff_spacing(doc, para, 'line')
    if line is None:
        return None
    rule = _eff_spacing(doc, para, 'lineRule')
    if rule not in (None, 'auto'):
        return None
    try:
        return int(line) / 240.0
    except ValueError:
        return None


def _eff_ind(doc, para, attr):
    """取 w:ind 上某属性（left/leftChars/firstLine/firstLineChars/hanging…）的整数值。"""
    for el in _ppr_chain(doc, para):
        ind = el.find(qn('w:ind'))
        if ind is not None:
            v = ind.get(qn('w:' + attr))
            if v is not None:
                try:
                    return int(v)
                except ValueError:
                    return None
    return None


def _eff_tabs(doc, para):
    """返回 [(val, pos_twips, leader), ...]；沿继承链取第一个定义 w:tabs 的层级。"""
    node = _chain_node(_ppr_chain(doc, para), 'w:tabs')
    if node is None:
        return []
    out = []
    for t in node.findall(qn('w:tab')):
        try:
            pos = int(t.get(qn('w:pos')))
        except (TypeError, ValueError):
            pos = None
        out.append((t.get(qn('w:val')), pos, t.get(qn('w:leader'))))
    return out

# ======================= 二、文本与层级判定 =======================

_WS_RE = re.compile(r'[\s　\xa0]+')
_LEADER_CHARS = '….．·‧・'          # … . ． · ‧ ・
_LVL2_RE = re.compile(r'^[（(][一二三四五六七八九十]+[)）]')
_LVL3_RE = re.compile(r'^\d+[\.．、]')
_LVL1_RE = re.compile(r'^[一二三四五六七八九十]+[\.．、]')
_TAIL_PAGE_RE = re.compile(r'[\t' + _LEADER_CHARS + r'\s]*(\d{1,3})\s*$')
_TAIL_ONLY_RE = re.compile(r'[\t' + _LEADER_CHARS + r'\s]+$')
_L1_KEYWORDS = ('总结', '致谢', '结论', '结束语', '参考文献', '附录', '摘要', '引言')


def _nows(s):
    """去掉所有空白（含全角空格/NBSP），用于文本等价比较。"""
    return _WS_RE.sub('', s or '')


def _collapse(s):
    return _WS_RE.sub(' ', (s or '').strip())


def _entry_page(text):
    """目录条目行末的阿拉伯数字页码；无则 None。"""
    m = _TAIL_PAGE_RE.search((text or '').strip())
    return int(m.group(1)) if m else None


def _entry_body(text):
    """剥掉行末页码与点线/省略号填充后的条目正文。"""
    t = (text or '').strip()
    t = _TAIL_PAGE_RE.sub('', t)
    t = _TAIL_ONLY_RE.sub('', t)
    return t.strip()


def _strip_numbering(text):
    """去掉层级序号（"一."/"（一）"/"1."），只留标题核心文字。"""
    t = _nows(text)
    for pat in (_LVL2_RE, _LVL3_RE, _LVL1_RE):
        t2 = pat.sub('', t, count=1)
        if t2 != t:
            return t2
    return t


def _entry_level(text, style_name):
    """目录条目层级：1/2/3。优先按 toc N 样式，其次按序号写法。"""
    sn = _nows(style_name).lower()
    m = re.match(r'^toc(\d+)$', sn)
    if m:
        lvl = int(m.group(1))
        return lvl if lvl in (1, 2, 3) else 3
    t = _nows(_entry_body(text))
    if _LVL2_RE.match(t):
        return 2
    if _LVL3_RE.match(t):
        return 3
    if _LVL1_RE.match(t):
        return 1
    if any(k == t or t.endswith(k) for k in _L1_KEYWORDS):
        return 1
    return 1


def _visual_units(text):
    """文本视觉宽度（以"全角字符宽"为单位）：中日韩全角=1，半角≈0.5，Tab 忽略。"""
    w = 0.0
    for ch in text or '':
        if ch == '\t':
            continue
        o = ord(ch)
        if 0x4E00 <= o <= 0x9FFF or 0x3000 <= o <= 0x303F or 0xFF00 <= o <= 0xFFEF:
            w += 1.0
        else:
            w += 0.5
    return w


def _has_drawing(element):
    if element.find('.//' + qn('w:drawing')) is not None:
        return True
    return element.find('.//' + qn('w:pict')) is not None


def _para_text_with_fields(para):
    """段落可见文本；自动目录域的域结果文本 python-docx 已能取到，此处仅做兜底拼接。"""
    txt = para.text
    if txt.strip():
        return txt
    parts = []
    for t in para._element.iter(qn('w:t')):
        parts.append(t.text or '')
    return ''.join(parts)

# ======================= 三、文档打开与目录定位 =======================

def _all_paragraphs(doc):
    """文档全部段落（含表格单元格内段落），按 XML 文档序返回。
    doc.paragraphs 只给出顶层段落，会漏掉排在表格里的标题。"""
    return [Paragraph(el, doc) for el in doc.element.body.iter(qn('w:p'))]


def _open_word_document(file_path):
    """打开 .docx 文档（python-docx 仅支持 OOXML/.docx）."""
    return Document(file_path)


def _has_toc_field(doc):
    """文档中是否存在自动目录域（w:instrText / w:fldSimple 的 TOC 指令）。"""
    for el in doc.element.body.iter(qn('w:instrText')):
        if 'TOC' in (el.text or '').upper():
            return True
    for el in doc.element.body.iter(qn('w:fldSimple')):
        if 'TOC' in (el.get(qn('w:instr')) or '').upper():
            return True
    return False


def _style_name(para):
    try:
        return para.style.name or ''
    except Exception:
        return ''


def _is_toc_entry_para(para):
    """判定某段落是否为一条目录条目（手工目录与自动目录域通用）。"""
    if _nows(_style_name(para)).lower().startswith('toc'):
        return True
    t = (para.text or '').strip()
    if not t:
        return False
    if len(t) > 80:
        return False
    if '\t' in para.text:
        return True
    if re.search(r'\d\s*$', t):
        return True
    if re.search('[' + _LEADER_CHARS + ']{3,}', t):
        return True
    return False


def _locate_toc(doc, paras):
    """返回 (目录标题段索引, [条目段索引...], 正文起始段索引)。"""
    title_idx = None
    for i, p in enumerate(paras):
        if _nows(p.text) == '目录':
            title_idx = i
            break
    if title_idx is None:
        return (None, [], 0)
    entries = []
    i = title_idx + 1
    while i < len(paras):
        p = paras[i]
        if not (p.text or '').strip():
            i += 1
            continue
        if _is_toc_entry_para(p):
            entries.append(i)
            i += 1
            continue
        break
    body_start = (entries[-1] + 1) if entries else (title_idx + 1)
    return (title_idx, entries, body_start)


def _para_section_index(doc, paras):
    """段落索引 → 所属节序号（按 body 中 sectPr 出现次序划分）。"""
    mapping = {}
    sect = 0
    idx = 0
    p_tag = qn('w:p')
    for child in doc.element.body.iterchildren():
        if child.tag == p_tag:
            mapping[idx] = sect
            idx += 1
            if child.find(qn('w:pPr')) is not None and \
                    child.find(qn('w:pPr')).find(qn('w:sectPr')) is not None:
                sect += 1
        elif child.tag == qn('w:sectPr'):
            sect += 1
    return mapping


def _render_to_pdf(file_path, tmp_dir):
    """把 .docx 渲染为 PDF，用于获取真实分页结果；仅走本机 Word COM 自动化。"""
    base_name = os.path.splitext(os.path.basename(file_path))[0]
    out_pdf = os.path.join(tmp_dir, base_name + '.pdf')
    try:
        import win32com.client  # pywin32
    except ImportError as exc:
        raise RuntimeError('未安装 pywin32, 无法调用 Word 渲染 PDF: ' + str(exc))
    word = None
    doc = None
    try:
        word = win32com.client.DispatchEx('Word.Application')
        word.Visible = False
        word.DisplayAlerts = 0
        doc = word.Documents.Open(os.path.abspath(file_path), ReadOnly=True)
        doc.ExportAsFixedFormat(
            OutputFileName=os.path.abspath(out_pdf),
            ExportFormat=17,  # wdExportFormatPDF
        )
    except Exception as exc:
        raise RuntimeError('Word PDF 渲染失败: ' + str(exc))
    finally:
        try:
            if doc is not None:
                doc.Close(False)
        except Exception:
            pass
        try:
            if word is not None:
                word.Quit()
        except Exception:
            pass
    if not os.path.isfile(out_pdf):
        raise RuntimeError('Word PDF 渲染未生成目标文件')
    return out_pdf

def _build_pdf_probe(file_path):
    """渲染为 PDF 后返回 (locate, displayed)：
       locate(snippet) -> 该文本在正文范围内首次出现的物理页序号（0-based）或 None；
       displayed[phys] -> 该物理页页脚上显示的阿拉伯页码（无法解析则 None）。
       渲染或 PyMuPDF 不可用时返回 (None, None)，调用方退回静态判定。"""
    if fitz is None:
        return (None, None)
    tmp_dir = tempfile.mkdtemp(prefix='officeval_007_pdf_')
    try:
        pdf_path = _render_to_pdf(file_path, tmp_dir)
        pdf_doc = fitz.open(pdf_path)
        page_texts = [page.get_text() for page in pdf_doc]
        pdf_doc.close()
    except Exception:
        shutil.rmtree(tmp_dir, ignore_errors=True)
        return (None, None)
    shutil.rmtree(tmp_dir, ignore_errors=True)

    # PyMuPDF 抽取文本时会在中/英/数字交界插入换行或空格，去空白后再匹配。
    norm = [_nows(t) for t in page_texts]

    # 页脚显示页码：取每页最后一个独立成行的阿拉伯数字
    displayed = []
    for txt in page_texts:
        num = None
        for line in reversed([ln.strip() for ln in (txt or '').splitlines()]):
            if re.fullmatch(r'\d{1,3}', line):
                num = int(line)
                break
        displayed.append(num)

    # 目录页自身列全了所有条目文字，若从第0页起搜索，每条都会先命中目录页而非
    # 正文标题。正文以"页脚首次出现阿拉伯页码"的物理页为界（前置部分为罗马数字
    # 或无页码），以此作搜索下界，正向游标与兜底全扫都不得越过该界回到目录页。
    body_first = next((p for p, v in enumerate(displayed) if v is not None), 0)
    state = {'cursor': body_first}

    def locate(snippet):
        ns = _nows(snippet)
        if not ns:
            return None
        for p in range(state['cursor'], len(norm)):
            if ns in norm[p]:
                state['cursor'] = p
                return p
        for p in range(body_first, len(norm)):
            if ns in norm[p]:
                return p
        return None

    return (locate, displayed)


# ======================= 四、对外接口 =======================

def evaluate(dir_path):
    """按 §2.2 约定，返回结构化 dict；不 print、不改 sys.stdout、不 sys.exit。
    入参 dir_path 为"脚本所在目录的路径"，脚本自己在该目录里定位并打开被评估的
    .docx（忽略以 ~$ 开头的 Word 临时锁文件）。脚本自身崩溃（含目录不存在、
    未找到文档等）由本函数捕获并返回 status='error'、error=异常摘要、
    dim1_pass=''、dim1_reason=''、max_score=0——此时维度一根本没跑起来，
    故不给出 True/False 结论，未通过原因由 error 字段承载。"""
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
            raise FileNotFoundError(f"目录不存在：{dir_path}")
        cands = sorted(
            f for f in os.listdir(dir_path)
            if f.lower().endswith('.docx') and not f.startswith('~$')
        )
        if not cands:
            raise FileNotFoundError(f"目录中未找到 .docx 文件：{dir_path}")
        file_path = os.path.join(dir_path, cands[0])
        result["file_name"] = os.path.basename(file_path)
        return _evaluate_impl(file_path, result)
    except Exception as exc:
        # 脚本异常（目录不存在 / 未找到 .docx / 文档打不开等）：status=error，
        # error 给出异常摘要；维度一未执行，故 dim1_pass 与 dim1_reason 均置空，
        # 满分也归零（没有实际评分过程）。目录/文件类问题只给可读消息，
        # 其余异常带上类型名便于定位。
        msg = str(exc).strip()
        if not isinstance(exc, (FileNotFoundError, NotADirectoryError)) or not msg:
            msg = f"{type(exc).__name__}: {msg}" if msg else type(exc).__name__
        result["status"] = "error"
        result["error"] = msg
        result["dim1_pass"] = ""
        result["dim1_reason"] = ""
        result["dim2_items"] = []
        result["total_score"] = 0
        result["max_score"] = 0
        return result

def _evaluate_impl(file_path, result):
    """维度一门禁 + 维度二逐项评分；results 内先按固定顺序追加
    "+N：说明" / "-N：说明"，末尾统一反解为结构化 dim2_items。"""
    doc = _open_word_document(file_path)
    results = []   # [(raw, detail), ...]

    def gain(value, rule, ok, detail=""):
        """加分项：命中记 +value，未命中记 +0。"""
        results.append(("+%d：%s" % (value if ok else 0, rule), detail))

    def penalty(value, rule, violated, detail=""):
        """减分项：触发（即出现指令未涉及的破坏性改动）记 -value，未触发记 -0。"""
        results.append(("-%d：%s" % (value if violated else 0, rule), detail))

    paras = doc.paragraphs
    all_text = '\n'.join(p.text for p in paras)
    toc_title_idx, entry_idx, body_start = _locate_toc(doc, paras)

    # 目录条目结构化：(段索引, 段落, 原文, 层级, 行末页码, 剥离页码后的正文)
    entries = []
    for i in entry_idx:
        p = paras[i]
        raw = _para_text_with_fields(p)
        entries.append({
            'idx': i,
            'para': p,
            'raw': raw,
            'level': _entry_level(raw, _style_name(p)),
            'page': _entry_page(raw),
            'body': _entry_body(raw),
            'core': _strip_numbering(_entry_body(raw)),
        })
    lv = {n: [e for e in entries if e['level'] == n] for n in (1, 2, 3)}

    # 版心尺寸（twips）：右对齐制表位应落在版心右边界
    sec0 = doc.sections[0]
    page_w = sec0.page_width.twips if sec0.page_width is not None else 11906
    lm = sec0.left_margin.twips if sec0.left_margin is not None else 1800
    rm = sec0.right_margin.twips if sec0.right_margin is not None else 1800
    text_width = page_w - lm - rm          # 本套模板 = 8306 twips = 14.65cm

    # 正文段落（目录之后）。用文档序全量段落，避免漏掉排在表格内的标题；
    # 索引 j 为"全量段落序号"，与 paras（仅顶层段落）的索引体系相互独立。
    apar = _all_paragraphs(doc)
    aels = [p._element for p in apar]
    try:
        astart = aels.index(paras[body_start]._element) if body_start < len(paras) \
            else len(apar)
    except ValueError:
        astart = 0
    body = [(j, apar[j]) for j in range(astart, len(apar))]

    # 章标题 / 节标题：优先按大纲级别取（自动目录即以此为据），
    # 大纲级别缺失时退回按序号写法识别。
    ol_group = {n: [(j, p) for (j, p) in body if _eff_outline(doc, p) == n]
                for n in (0, 1, 2)}
    chapter_heads = ol_group[0] if len(ol_group[0]) >= 4 else \
        [(j, p) for (j, p) in body
         if _LVL1_RE.match(_nows(p.text)) and 0 < len(_nows(p.text)) <= 40]
    section_heads = ol_group[1] if len(ol_group[1]) >= 10 else \
        [(j, p) for (j, p) in body
         if _LVL2_RE.match(_nows(p.text)) and 0 < len(_nows(p.text)) <= 30]
    sub_heads = ol_group[2] if len(ol_group[2]) >= 10 else \
        [(j, p) for (j, p) in body
         if _LVL3_RE.match(_nows(p.text)) and 0 < len(_nows(p.text)) <= 30]
    body_texts_nows = set(_nows(p.text) for (_j, p) in body if p.text.strip())
    body_cores = set(_strip_numbering(p.text) for (_j, p) in body if p.text.strip())

    # ========== 维度1：可用与可修改性 ==========
    dim1_fail = []
    if toc_title_idx is None:
        dim1_fail.append('未找到"目  录"标题段落')
    if len(entries) < 5:
        dim1_fail.append('目录区可编辑条目不足5条（疑为图片/不可编辑对象）')
    for key, label in (('成人高等教育毕业论文', '封面'), ('作者声明', '作者声明'),
                       ('Abstract', 'Abstract'), ('致谢', '致谢')):
        if key not in all_text:
            dim1_fail.append('缺失%s整块内容' % label)
    # 页码越出纸张可打印区域（右对齐制表位跑到纸张右边缘之外）→ 不可辨认
    over = [e for e in entries
            for (v, pos, _ld) in _eff_tabs(doc, e['para'])
            if v == 'right' and pos is not None and pos > text_width + rm]
    if over:
        dim1_fail.append('目录页码制表位越出纸张右边缘，页码被裁切')
    if dim1_fail:
        result["dim1_pass"] = False
        result["dim1_reason"] = '；'.join(dim1_fail)
        result["dim2_items"] = []
        result["total_score"] = 0
        result["max_score"] = MAX_SCORE_TOTAL
        return result

    # ---------- 01 +5 目录由TOC域或toc样式承载 ----------
    # 严格对齐细则，仅踩细则明确写出的点：TOC域 或 toc 1/2/3 专用样式
    has_field = _has_toc_field(doc)
    styled = sum(1 for e in entries if re.match(r'^toc\s*[123]$', _style_name(e['para']), re.I))
    ok01 = has_field or styled >= max(3, int(len(entries) * 0.8))
    gain(5, '目录由可自动更新的目录域（TOC域）或专用目录样式（toc 1/2/3）承载，而非逐行手工敲入的普通段落文字',
         ok01, 'TOC域=%s，toc样式条目=%d/%d' % (has_field, styled, len(entries)))

    # 第01项同时作为"前置检查项"：下列以 gain01 记分的项（06/07/12/16/17/18/19）
    # 在原稿手工目录上也会被动满足，脱离自动更新目录即失去意义，故目录未由
    # TOC域/toc样式承载时一律不给分。
    def gain01(value, rule, ok, detail=""):
        gain(value, rule, bool(ok01 and ok),
             detail + '；前置项01（自动更新目录）=%s' % ok01)

    # ---------- 02 +5 正文各级标题设置了大纲级别/标题样式 ----------
    o0, o1, o2 = (len(ol_group[n]) for n in (0, 1, 2))
    ok02 = o0 >= 4 and o1 >= 10 and o2 >= 10
    gain(5, '正文各级标题已设置对应大纲级别或标题样式（章标题=大纲1级、"（一）（二）"=大纲2级、"1./2."=大纲3级），使目录可随正文自动生成与更新',
         ok02, '大纲1级=%d、2级=%d、3级=%d' % (o0, o1, o2))

    # ---------- 03 +5 三级层级收录完整、条目≥45、无成组遗漏 ----------
    cores = [e['core'] for e in entries if e['core']]

    def covered(text):
        """该正文标题是否被某条目录条目收录（条目文字含自动编号，故按核心文字比对）。"""
        c = _strip_numbering(text)
        if not c:
            return True
        return any(c == x or c in x or x in c for x in cores)
    miss = [_nows(p.text)[:16] for (_j, p) in (section_heads + sub_heads)
            if not covered(p.text)]
    ok03 = (len(entries) >= 45 and len(lv[1]) >= 4 and len(lv[2]) >= 12
            and len(lv[3]) >= 25 and not miss)
    gain(5, '目录完整收录三级层级——4个章标题、12个"（一）（二）…"节标题及其下全部"1./2./3."小节标题，条目总数不少于45条，无成组遗漏',
         ok03, '条目=%d（一级%d/二级%d/三级%d），正文节标题%d+小节标题%d，未收录%d个%s'
         % (len(entries), len(lv[1]), len(lv[2]), len(lv[3]),
            len(section_heads), len(sub_heads), len(miss), miss[:5]))

    # ---------- 04 +5 点线前导符 + 版心右边界右对齐制表位 ----------
    tol = 113                                     # 0.2cm ≈ 113 twips
    bad04 = []
    for e in entries:
        tabs = _eff_tabs(doc, e['para'])
        right = [(pos, ld) for (v, pos, ld) in tabs if v == 'right' and pos is not None]
        hit = any(abs(pos - text_width) <= tol and (ld in ('dot', 'middleDot', 'hyphen'))
                  for (pos, ld) in right)
        if not hit:
            bad04.append('[%d]%s tabs=%s' % (e['idx'], e['body'][:12],
                                             [(p, l) for (p, l) in right]))
    gain(5, '每条目录行以点线前导符（点线或中圆点）连接条目文字与页码，并由一个右对齐制表位收尾；制表位落在版心右边界（8306 twips / 14.65cm，误差≤0.2cm），页码右端与版心右边界对齐且不越出右页边距',
         not bad04, '版心宽=%d twips，不合格%d条%s' % (text_width, len(bad04), bad04[:3]))

    # ---------- 05 +5 目录条目文字与正文标题完全一致 ----------
    bad05 = []
    for e in entries:
        t = _nows(e['body'])
        if not t:
            continue
        if t in body_texts_nows:
            continue
        core = e['core']
        if core and core in body_cores:
            if e['level'] == 1:                   # 一级条目须连序号写法一并一致
                bad05.append('[%d]%s（序号写法与正文不一致）' % (e['idx'], t[:16]))
            continue
        if not any(core and (core in b or b in core) for b in body_cores if b):
            bad05.append('[%d]%s（正文无对应标题）' % (e['idx'], t[:16]))
    gain(5, '目录条目文字与正文对应标题文字完全一致（含"一.""二.""三.""四."、"（一）""（二）"、"1.""2."等序号写法及其中的标点与空格），不出现擅自改写、缩写、漏字',
         not bad05, '不一致%d条%s' % (len(bad05), bad05[:3]))

    # ---------- 06 +3 目录标题"目  录"格式 ----------
    tp = paras[toc_title_idx]
    t06 = tp.text
    runs06 = [r for r in tp.runs if r.text.strip()]
    sb, sbl = _eff_spacing(doc, tp, 'before'), _eff_spacing(doc, tp, 'beforeLines')
    sa, sal = _eff_spacing(doc, tp, 'after'), _eff_spacing(doc, tp, 'afterLines')

    def one_line(v_twips, v_lines):
        """段前/段后"1行"：beforeLines≈100 或 before≈1×行高（312 twips，16磅行）。"""
        if v_lines is not None:
            try:
                return abs(int(v_lines) - 100) <= 20
            except ValueError:
                return False
        if v_twips is not None:
            try:
                return 260 <= int(v_twips) <= 400
            except ValueError:
                return False
        return False
    ok06 = (re.match(r'^目[\s　]{1,2}录$', (t06 or '').strip()) is not None
            and _eff_jc(doc, tp) == 'center'
            and bool(runs06)
            and all(_is_heiti(_eff_eastasia(doc, r)) for r in runs06)
            and all(_eff_size_pt(doc, r) == 16 for r in runs06)
            and not any(_eff_bold(doc, r) for r in runs06)
            and one_line(sb, sbl) and one_line(sa, sal)
            and _eff_line_multiple(doc, tp) == 2.0)
    gain01(3, '目录标题"目  录"为居中、中间空两个字符、黑体三号（16磅）、不加粗，段前段后1行、2倍行距',
         ok06, 'text=%r jc=%s 字体=%s 字号=%s 加粗=%s 段前=%s/%s 段后=%s/%s 行距=%s'
         % (t06, _eff_jc(doc, tp), [_eff_eastasia(doc, r) for r in runs06][:2],
            [_eff_size_pt(doc, r) for r in runs06][:2],
            [_eff_bold(doc, r) for r in runs06][:2],
            sb, sbl, sa, sal, _eff_line_multiple(doc, tp)))

    def check_entries(group, want_font, no_bold=False):
        """按层级校验条目字体/字号/对齐/行距（no_bold 时另校验不加粗）。"""
        bad = []
        for e in group:
            p = e['para']
            runs = [r for r in p.runs if r.text.strip()]
            f_ok = bool(runs) and all(want_font(_eff_eastasia(doc, r)) for r in runs)
            s_ok = bool(runs) and all(_eff_size_pt(doc, r) == 12 for r in runs)
            b_ok = (not any(_eff_bold(doc, r) for r in runs)) if no_bold else True
            if not (f_ok and s_ok and b_ok
                    and _eff_jc(doc, p) in ('left', 'both', None)
                    and _eff_line_multiple(doc, p) == 1.5):
                bad.append('[%d]%s 字体=%s 字号=%s 粗=%s jc=%s 行距=%s'
                           % (e['idx'], e['body'][:10],
                              [_eff_eastasia(doc, r) for r in runs][:1],
                              [_eff_size_pt(doc, r) for r in runs][:1],
                              [_eff_bold(doc, r) for r in runs][:1],
                              _eff_jc(doc, p), _eff_line_multiple(doc, p)))
        return bad

    # ---------- 07 +3 一级条目 黑体小四、左对齐、1.5倍行距、不加粗 ----------
    b07 = check_entries(lv[1], _is_heiti, no_bold=True)
    gain01(3, '一级条目（章标题行）为黑体小四（12磅）、左对齐、1.5倍行距、不加粗',
         bool(lv[1]) and not b07, '一级%d条，不合格%d条%s' % (len(lv[1]), len(b07), b07[:3]))

    # ---------- 08 +3 二级条目 宋体小四、左对齐、1.5倍行距、缩进为0 ----------
    b08 = check_entries(lv[2], _is_songti)
    for e in lv[2]:
        vals = [(k, _eff_ind(doc, e['para'], k) or 0)
                for k in ('left', 'leftChars', 'firstLine', 'firstLineChars')]
        if any(v != 0 for (_k, v) in vals):
            b08.append('[%d]%s ind=%s' % (e['idx'], e['body'][:10], vals))
    gain(3, '二级条目"（一）…"为宋体小四（12磅）、左对齐、1.5倍行距，左缩进与首行缩进均为0',
         bool(lv[2]) and not b08, '二级%d条，不合格%d条%s' % (len(lv[2]), len(b08), b08[:3]))

    # ---------- 09 +1 三级条目 宋体小四、左对齐、1.5倍行距、左缩进约1字符 ----------
    b09 = check_entries(lv[3], _is_songti)
    for e in lv[3]:
        chars = _eff_ind(doc, e['para'], 'leftChars')
        left = _eff_ind(doc, e['para'], 'left') or 0
        # 1 字符：leftChars≈100（百分之一字符为单位）或 left≈240 twips（小四12磅）
        if not ((chars is not None and 50 <= chars <= 160) or 110 <= left <= 360):
            b09.append('[%d]%s leftChars=%s left=%s' % (e['idx'], e['body'][:10], chars, left))
    gain(1, '三级条目"1. …"为宋体小四（12磅）、左对齐、1.5倍行距，左缩进约1个字符',
         bool(lv[3]) and not b09, '三级%d条，不合格%d条%s' % (len(lv[3]), len(b09), b09[:3]))

    # ---------- 10 +3 三级缩进正确递进且同级完全一致 ----------
    def ind_key(e):
        c = _eff_ind(doc, e['para'], 'leftChars')
        return c if c is not None else (_eff_ind(doc, e['para'], 'left') or 0)
    set1, set2, set3 = (set(ind_key(e) for e in lv[n]) for n in (1, 2, 3))
    ok10 = (len(set1) <= 1 and len(set2) <= 1 and len(set3) == 1
            and all(v == 0 for v in set1) and all(v == 0 for v in set2)
            and all(v > 0 for v in set3))
    gain(3, '三级层级缩进呈正确递进（一级与二级顶格、三级比二级多缩进约1字符），且同级条目缩进量完全一致',
         ok10, '一级缩进=%s 二级=%s 三级=%s' % (sorted(set1), sorted(set2), sorted(set3)))

    # ---------- 11 +3 目录中不再出现"参考文献""附  录" ----------
    ghost = [e['body'] for e in entries
             if _nows(e['body']) and re.search(r'(参考文献|附录)', _nows(e['body']))]
    gain(1, '目录中不再出现正文没有对应章节的"参考文献""附  录"两个条目',
         not ghost, '残留条目%s' % (ghost[:3],))

    # ---------- 12 +1 目录收录"总结""致谢" ----------
    has_sum = any('总结' in _nows(e['body']) for e in entries)
    has_ack = any('致谢' in _nows(e['body']) for e in entries)
    gain01(1, '目录收录正文末尾的"总结""致谢"两个条目',
         has_sum and has_ack, '总结=%s 致谢=%s' % (has_sum, has_ack))

    # ---------- 13 +3 无手工省略号/句点串填充 ----------
    filler_re = re.compile(r'…{2,}|\.{4,}|[．·‧・]{4,}')
    bad13 = [(e['idx'], e['raw'][:24]) for e in entries if filler_re.search(e['raw'] or '')]
    gain(3, '目录中不再保留手工输入的省略号/句点串（"…………""......"）作填充，各行无长度参差的省略号残留',
         not bad13, '残留填充%d条%s' % (len(bad13), bad13[:3]))

    # ---------- 14 +3 每行行末均有阿拉伯数字页码，无空缺、无占位符 ----------
    bad14 = [(e['idx'], e['raw'][-18:]) for e in entries if e['page'] is None]
    ph14 = [(e['idx'], e['raw'][-18:]) for e in entries
            if re.search(r'(页码|#|\?|？|PAGE)\s*$', (e['raw'] or '').strip(), re.I)]
    gain(3, '每条目录行末均有对应的阿拉伯数字页码，无页码空缺、无"页码""#""?"等占位符',
         not bad14 and not ph14, '缺页码%d条%s；占位符%d条%s'
         % (len(bad14), bad14[:3], len(ph14), ph14[:2]))

    # ---------- 15 +3 页码与标题实际所在页一致，且自上而下只增不减 ----------
    pages = [e['page'] for e in entries if e['page'] is not None]
    monotonic = all(b >= a for a, b in zip(pages, pages[1:]))
    locate, displayed = _build_pdf_probe(file_path)
    if locate is None:
        ok15 = bool(pages) and monotonic and len(pages) == len(entries)
        d15 = '未能渲染PDF，仅作静态判定：页码序列%s，只增不减=%s' % (pages[:12], monotonic)
    else:
        # 逐条比对：条目页码 与 该标题在重排后实际出现页的页脚页码；允许全篇一致偏移≤1页
        diffs, unfound = [], 0
        for e in entries:
            if e['page'] is None:
                continue
            phys = locate(e['body'] or e['raw'])
            if phys is None:
                unfound += 1
                continue
            real = displayed[phys] if phys < len(displayed) else None
            if real is None:
                unfound += 1
                continue
            diffs.append((e['idx'], e['page'] - real))
        offs = set(d for (_i, d) in diffs)
        ok15 = (bool(diffs) and monotonic and len(offs) == 1
                and abs(next(iter(offs))) <= 1)
        d15 = '偏移集合=%s（须唯一且|偏移|≤1），只增不减=%s，未定位%d条' \
            % (sorted(offs)[:6], monotonic, unfound)
    gain(3, '目录页码与标题实际所在页一致——逐条比对条目页码与该标题重新分页后实际出现的页码，允许全篇一致偏移1页以内，不允许单条偏差；页码自上而下只增不减，同页条目页码相同、跨页条目依次递增，不出现前大后小的逆序',
         ok15, d15)

    # ---------- 16 +3 每条条目单行显示，文字与页码同行 ----------
    # 估算：条目文字宽度（小四12磅，1全角字=240 twips）+ 页码宽度 + 左缩进 ≤ 版心宽
    bad16 = []
    for e in entries:
        left = _eff_ind(doc, e['para'], 'left')
        if left is None:
            c = _eff_ind(doc, e['para'], 'leftChars')
            left = int(c * 2.4) if c is not None else 0
        need = left + _visual_units(e['body']) * 240 \
            + len(str(e['page'] if e['page'] is not None else 0)) * 120
        if need > text_width or '\n' in (e['raw'] or '') or '\v' in (e['raw'] or ''):
            bad16.append('[%d]%s 估算宽=%d>%d' % (e['idx'], e['body'][:12], need, text_width))
    gain01(3, '每条目录条目单行显示，条目文字与页码同处一行，不出现条目折行或页码被挤到下一行',
         not bad16, '折行风险%d条%s' % (len(bad16), bad16[:3]))

    # ---------- 17 +1 条目段前段后为0（或全表统一），条目间无空行 ----------
    sp17 = set()
    for e in entries:
        sp17.add((_eff_spacing(doc, e['para'], 'before'),
                  _eff_spacing(doc, e['para'], 'beforeLines'),
                  _eff_spacing(doc, e['para'], 'after'),
                  _eff_spacing(doc, e['para'], 'afterLines')))
    blank17 = sum(1 for i in range(entries[0]['idx'], entries[-1]['idx'])
                  if not (paras[i].text or '').strip())
    zero17 = all(v in (None, '0') for tup in sp17 for v in tup)
    gain01(1, '目录各条目段前段后间距为0（或全表统一取值），条目之间无多余空行',
         (zero17 or len(sp17) == 1) and blank17 == 0,
         '段前后取值集合=%s，条目间空行=%d' % (sorted(map(str, sp17))[:3], blank17))

    # ---------- 18 +1 只收录标题级文字，不混入正文段落/图注/"（1）（2）" ----------
    bad18 = []
    for e in entries:
        t = _nows(e['body'])
        if re.match(r'^[（(]\d+[)）]', t):
            bad18.append('[%d]%s（正文小点）' % (e['idx'], t[:14]))
        elif not (_LVL1_RE.match(t) or _LVL2_RE.match(t) or _LVL3_RE.match(t)
                  or any(k in t for k in _L1_KEYWORDS)):
            bad18.append('[%d]%s（非标题级文字）' % (e['idx'], t[:14]))
        elif len(t) > 40:
            bad18.append('[%d]%s（疑为正文段落）' % (e['idx'], t[:14]))
    gain01(1, '目录只收录标题级文字，不混入正文段落内容、图注，也不收录"（1）（2）（3）"级别的正文小点',
         not bad18, '异常%d条%s' % (len(bad18), bad18[:3]))

    # ---------- 19 +1 目录位于Abstract之后、"引  言"之前，独立起页 ----------
    abs_idx = next((i for i, p in enumerate(paras)
                    if _nows(p.text).lower() == 'abstract'), None)
    intro_idx = next((i for i in range(body_start, len(paras))
                      if '引言' in _nows(paras[i].text)
                      and len(_nows(paras[i].text)) <= 12), None)
    secmap = _para_section_index(doc, paras)

    def new_page_at(i):
        """该段是否另起一页：分节、pageBreakBefore、或本段/前段含分页符。"""
        if i <= 0:
            return True
        if secmap.get(i) != secmap.get(i - 1):
            return True
        for el in _ppr_chain(doc, paras[i])[:1]:
            if el.find(qn('w:pageBreakBefore')) is not None:
                return True
        for j in (i - 1, i):
            for br in paras[j]._element.iter(qn('w:br')):
                if br.get(qn('w:type')) == 'page':
                    return True
        return False
    ok19 = (abs_idx is not None and intro_idx is not None
            and abs_idx < toc_title_idx < intro_idx
            and new_page_at(toc_title_idx) and new_page_at(intro_idx))
    gain01(1, '目录位于Abstract之后、正文"引  言"之前，独立起页，其后正文仍从新页开始',
         ok19, 'Abstract段=%s 目录段=%s 引言段=%s 目录另起页=%s 正文另起页=%s'
         % (abs_idx, toc_title_idx, intro_idx,
            new_page_at(toc_title_idx) if abs_idx is not None else None,
            new_page_at(intro_idx) if intro_idx is not None else None))

    # ---------- 20 +1 "三"章条目为完整标题，未被截断为半句 ----------
    e20 = next((e for e in lv[1] if _nows(e['body']).startswith('三')), None)
    t20 = _nows(e20['body']) if e20 else ''
    ok20 = '活动资源梳理与一日生活融合路径' in t20 and '生活环节融合切入点' in t20
    gain(1, '"三"章条目为完整标题"活动资源梳理与一日生活融合路径与生活环节融合切入点"，未被截断为半句',
         ok20, '目录条目=%r' % (e20['body'] if e20 else None,))

    # ================= 减分项：指令未要求的破坏性改动 =================

    def find_para(pred, lo=0, hi=None):
        for i in range(lo, len(paras) if hi is None else hi):
            if pred(_nows(paras[i].text)):
                return (i, paras[i])
        return (None, None)

    def fmt_of(p):
        runs = [r for r in p.runs if r.text.strip()]
        return ('jc=%s 中文字体=%s 西文=%s 字号=%s 粗=%s 行距=%s 首行=%s/%s 段前后=%s/%s'
                % (_eff_jc(doc, p), [_eff_eastasia(doc, r) for r in runs][:1],
                   [_eff_ascii(doc, r) for r in runs][:1],
                   sorted(str(_eff_size_pt(doc, r)) for r in runs),
                   [_eff_bold(doc, r) for r in runs][:1], _eff_line_multiple(doc, p),
                   _eff_ind(doc, p, 'firstLine'), _eff_ind(doc, p, 'firstLineChars'),
                   _eff_spacing(doc, p, 'before'), _eff_spacing(doc, p, 'after')))

    def sizes(p):
        return set(_eff_size_pt(doc, r) for r in p.runs if r.text.strip())

    def fonts_ea(p):
        return set(_eff_eastasia(doc, r) for r in p.runs if r.text.strip())

    def first_line_2char(p):
        fc = _eff_ind(doc, p, 'firstLineChars')
        fl = _eff_ind(doc, p, 'firstLine') or 0
        return (fc is not None and 150 <= fc <= 250) or 480 <= fl <= 760

    # ---------- 21 -5 封面校徽图形/图片缺失 ----------
    decl_idx, _decl_p = find_para(lambda t: '作者声明' in t)
    cover_hi = decl_idx if decl_idx is not None else min(len(paras), 20)
    logo = any(_has_drawing(paras[i]._element) for i in range(0, cover_hi))
    penalty(5, '封面页顶部的校徽图形/图片元素缺失或被替换', not logo,
            '封面区（段0-%d）图形对象=%s' % (cover_hi, logo))

    # ---------- 22 -5 正文段落整段删除/重复/顺序错乱 ----------
    _CN = {'一': 1, '二': 2, '三': 3, '四': 4, '五': 5, '六': 6}
    order = [_CN.get(_nows(p.text)[:1]) for (_i, p) in chapter_heads]
    order = [v for v in order if v is not None]
    asc = all(b > a for a, b in zip(order, order[1:]))
    longs = [_nows(p.text) for (_i, p) in body if len(_nows(p.text)) >= 60]
    dup = [t[:16] for t in set(longs) if longs.count(t) > 1]
    need22 = ['引言', '理论依据', '设计原则', '活动资源', '优化策略', '总结', '致谢']
    lost = [k for k in need22 if k not in _nows(all_text)]
    penalty(5, '正文四章（引  言／自然探究活动融入一日生活的理论依据与设计原则／活动资源梳理与一日生活融合路径／优化策略与研究结论）及总结、致谢的段落出现整段删除、整段重复或前后顺序错乱',
            (not asc) or bool(dup) or bool(lost) or len(order) < 4,
            '章序=%s 递增=%s 重复段=%s 缺失关键块=%s' % (order, asc, dup[:3], lost))

    # ---------- 23 -5 分节结构被破坏 ----------
    n_sect = len(doc.sections)
    s_decl = secmap.get(decl_idx) if decl_idx is not None else None
    s_toc = secmap.get(toc_title_idx)
    s_body = secmap.get(intro_idx) if intro_idx is not None else None
    penalty(5, '文档分节结构被破坏，封面、作者声明、摘要与Abstract、目录、正文不再分属独立节，导致分段页码编排失效',
            n_sect < 4 or s_toc is None or s_body is None or s_toc == s_body
            or (s_decl is not None and s_decl == s_toc),
            '节数=%d 作者声明节=%s 目录节=%s 正文节=%s' % (n_sect, s_decl, s_toc, s_body))

    # ---------- 24 -3 纸张A4与页边距 ----------
    bad24 = []
    for k, s in enumerate(doc.sections):
        def tw(v):
            return None if v is None else v.twips
        got = (tw(s.page_width), tw(s.page_height), tw(s.top_margin),
               tw(s.bottom_margin), tw(s.left_margin), tw(s.right_margin))
        want = (11906, 16838, 1440, 1440, 1800, 1800)
        if any(g is None or abs(g - w) > 30 for g, w in zip(got, want)):
            bad24.append('节%d=%s' % (k, got))
    penalty(3, '纸张不满足A4（21cm×29.7cm / 11906×16838 twips），页边距不满足上2.54cm、下2.54cm、左3.17cm、右3.17cm',
            bool(bad24), '不合格%d节%s（期望11906×16838/1440/1440/1800/1800）'
            % (len(bad24), bad24[:3]))

    def pg_num_type(sect):
        sp = sect._sectPr
        node = sp.find(qn('w:pgNumType'))
        if node is None:
            return (None, None)
        return (node.get(qn('w:fmt')), node.get(qn('w:start')))

    # ---------- 25 -3 前置部分页码为大写罗马数字 ----------
    front_sects = sorted(set(v for k, v in secmap.items()
                            if k <= (entries[-1]['idx'] if entries else toc_title_idx)))
    fmts25 = [(k, pg_num_type(doc.sections[k])[0]) for k in front_sects
              if k < len(doc.sections)]
    penalty(3, '封面、作者声明、摘要、Abstract、目录等前置部分的页码不满足大写罗马数字（Ⅰ、Ⅱ、Ⅲ…）格式',
            not fmts25 or not any(f == 'upperRoman' for (_k, f) in fmts25),
            '前置各节页码格式=%s' % (fmts25,))

    # ---------- 26 -3 正文页码为阿拉伯数字且从1重新起编 ----------
    fmt26, start26 = pg_num_type(doc.sections[s_body]) if \
        (s_body is not None and s_body < len(doc.sections)) else (None, None)
    penalty(3, '正文部分页码不满足阿拉伯数字格式且从1重新起编',
            fmt26 not in (None, 'decimal') or start26 != '1',
            '正文节=%s fmt=%s start=%s' % (s_body, fmt26, start26))

    # ---------- 27 -3 页眉/页脚被删除或格式被改动 ----------
    def part_has_content(part):
        """页眉/页脚是否有内容（含域代码，页码域在 .text 里取不到）。"""
        el = part._element
        for t in el.iter(qn('w:t')):
            if (t.text or '').strip():
                return True
        for tag in ('w:instrText', 'w:fldChar', 'w:drawing', 'w:pict'):
            if el.find('.//' + qn(tag)) is not None:
                return True
        return False

    bad27 = []
    for k in range(s_toc if s_toc is not None else 0, n_sect):
        s = doc.sections[k]
        if not part_has_content(s.header):
            bad27.append('节%d页眉被删空' % k)
        if not part_has_content(s.footer):
            bad27.append('节%d页脚被删空' % k)
    for k, s in enumerate(doc.sections):
        for p in [q for q in s.header.paragraphs if (q.text or '').strip()]:
            pbdr = _chain_node(_ppr_chain(doc, p), 'w:pBdr')
            has_bottom = pbdr is not None and pbdr.find(qn('w:bottom')) is not None
            szs = sizes(p)
            # 小五=9磅、五号=10.5磅：本文各节页眉原为该两档，超出即视为字号被改动
            if _eff_jc(doc, p) != 'center' or not has_bottom or szs - {9.0, 10.5}:
                bad27.append('节%d页眉 jc=%s 下框线=%s 字号=%s'
                             % (k, _eff_jc(doc, p), has_bottom,
                                sorted(map(str, szs))))
    penalty(3, '页眉/页脚被删除，或其居中、带下框线、字号9磅（小五）的格式被改动',
            bool(bad27), '；'.join(bad27[:4]))

    # ---------- 28 -1 "学校代码："行 右对齐、宋体四号 ----------
    i28, p28 = find_para(lambda t: t.startswith('学校代码'), 0, cover_hi)
    penalty(1, '文档首页顶部"学校代码："一行不满足右对齐、宋体四号（14磅）',
            p28 is None or _eff_jc(doc, p28) != 'right'
            or not all(_is_songti(f) for f in fonts_ea(p28)) or sizes(p28) - {14.0},
            '段%s %s' % (i28, fmt_of(p28) if p28 is not None else '未找到'))

    # ---------- 29 -1 "成人高等教育毕业论文（设计）" 居中、宋体小二、加粗 ----------
    i29, p29 = find_para(lambda t: '成人高等教育毕业论文' in t, 0, cover_hi)
    penalty(1, '封面"成人高等教育毕业论文（设计）"不满足居中、宋体小二（32磅）、加粗',
            p29 is None or _eff_jc(doc, p29) != 'center'
            or not all(_is_songti(f) for f in fonts_ea(p29)) or sizes(p29) - {32.0}
            or not all(_eff_bold(doc, r) for r in p29.runs if r.text.strip()),
            '段%s %s' % (i29, fmt_of(p29) if p29 is not None else '未找到'))

    # ---------- 30 -1 封面题目两行 各自独立成行、居中、华文新魏二号、不加粗 ----------
    t30a, t30b = '幼儿园自然探究活动融入一日生活的', '课程设计与应用'
    i30a, p30a = find_para(lambda t: t == t30a, 0, cover_hi)
    i30b, p30b = find_para(lambda t: t == t30b, 0, cover_hi)
    bad30 = p30a is None or p30b is None
    if not bad30:
        for p in (p30a, p30b):
            if (_eff_jc(doc, p) != 'center'
                    or not all('华文新魏' in (f or '') for f in fonts_ea(p))
                    or sizes(p) - {24.0}
                    or any(_eff_bold(doc, r) for r in p.runs if r.text.strip())):
                bad30 = True
    penalty(1, '封面题目"幼儿园自然探究活动融入一日生活的""课程设计与应用"两行不满足各自独立成行、居中、华文新魏二号（24磅）、不加粗',
            bad30, '段%s/%s %s | %s' % (i30a, i30b,
                                        fmt_of(p30a) if p30a is not None else '未找到',
                                        fmt_of(p30b) if p30b is not None else '未找到'))

    # ---------- 31 -1 "启明学院毕业论文作者声明" 居中、宋体三号、1.5倍行距 ----------
    i31, p31 = find_para(lambda t: '毕业论文作者声明' in t)
    penalty(1, '"启明学院毕业论文作者声明"不满足居中、宋体三号（16磅）、1.5倍行距',
            p31 is None or _eff_jc(doc, p31) != 'center'
            or not all(_is_songti(f) for f in fonts_ea(p31)) or sizes(p31) - {16.0}
            or _eff_line_multiple(doc, p31) != 1.5,
            '段%s %s' % (i31, fmt_of(p31) if p31 is not None else '未找到'))

    # ---------- 32 -1 作者声明两段正文 三号、首行缩进2字符、1.5倍行距 ----------
    decl_body = [(i, paras[i]) for i in range(i31 + 1 if i31 is not None else 0,
                                              min(len(paras), (i31 or 0) + 12))
                 if _nows(paras[i].text).startswith(('本人声明', '本人同意学校'))]
    bad32 = len(decl_body) < 2
    for _i, p in decl_body:
        if sizes(p) - {16.0} or not first_line_2char(p) \
                or _eff_line_multiple(doc, p) != 1.5:
            bad32 = True
    penalty(1, '作者声明两段正文（"本人声明：……""本人同意学校……"）不满足三号（16磅）、首行缩进2字符、1.5倍行距',
            bad32, '找到%d段：%s' % (len(decl_body),
                                    [(i, fmt_of(p)) for (i, p) in decl_body][:2]))

    # ---------- 33 -1 签名两行的文字内容/空格占位/三号字号 ----------
    i33a, p33a = find_para(lambda t: t.startswith('作者签名'))
    i33b, p33b = find_para(lambda t: t.startswith('指导教师签名'))
    bad33 = p33a is None or p33b is None
    if not bad33:
        bad33 = (_nows(p33a.text) != '作者签名：日期：2026.4.20'
                 or _nows(p33b.text) != '指导教师签名：日期：'
                 or bool(sizes(p33a) - {16.0}) or bool(sizes(p33b) - {16.0})
                 or not re.search(r'\s{4,}', p33a.text or '')
                 or not re.search(r'\s{4,}', p33b.text or ''))
    penalty(1, '"作者签名：                   日期：2026.4.20"与"指导教师签名：               日期："两行的文字内容、空格占位或三号（16磅）字号被改动',
            bad33, '段%s=%r 字号%s；段%s=%r 字号%s'
            % (i33a, (p33a.text if p33a is not None else None), sorted(map(str, sizes(p33a))) if p33a is not None else None,
               i33b, (p33b.text if p33b is not None else None), sorted(map(str, sizes(p33b))) if p33b is not None else None))

    def sp_157(p):
        """段前段后15.7磅（314 twips，容差±20）。"""
        b, a = _eff_spacing(doc, p, 'before'), _eff_spacing(doc, p, 'after')
        try:
            return abs(int(b) - 314) <= 20 and abs(int(a) - 314) <= 20
        except (TypeError, ValueError):
            return False

    # ---------- 34 -1 "摘  要" 居中、黑体三号、段前段后15.7磅、2倍行距 ----------
    i34, p34 = find_para(lambda t: t == '摘要')
    penalty(1, '"摘  要"不满足居中、黑体三号（16磅）、段前段后15.7磅、2倍行距',
            p34 is None or _eff_jc(doc, p34) != 'center'
            or not all(_is_heiti(f) for f in fonts_ea(p34)) or sizes(p34) - {16.0}
            or not sp_157(p34) or _eff_line_multiple(doc, p34) != 2.0,
            '段%s %s' % (i34, fmt_of(p34) if p34 is not None else '未找到'))

    # ---------- 35 -1 摘要正文 小四、首行缩进2字符、1.5倍行距、西文Times New Roman ----------
    i36, p36 = find_para(lambda t: t.startswith('关键词'))
    abs_body = [(i, paras[i]) for i in range((i34 or 0) + 1, i36 if i36 else (i34 or 0) + 1)
                if (paras[i].text or '').strip()]
    bad35 = not abs_body
    for _i, p in abs_body:
        if sizes(p) - {12.0} or not first_line_2char(p) \
                or _eff_line_multiple(doc, p) != 1.5 \
                or any('Times New Roman' not in (_eff_ascii(doc, r) or '')
                       for r in p.runs if r.text.strip()):
            bad35 = True
    penalty(1, '摘要正文不满足小四（12磅）、首行缩进2字符、1.5倍行距，或其西文字体不为Times New Roman',
            bad35, '%d段：%s' % (len(abs_body), [(i, fmt_of(p)) for (i, p) in abs_body][:2]))

    # ---------- 36 -1 中文关键词行 左对齐、黑体四号、条目未被增删改写 ----------
    kw = '关键词：自然探究；学前教育；一日生活；课程设计；活动应用'
    penalty(1, '"关键词：自然探究；学前教育；一日生活；课程设计；活动应用"不满足左对齐、黑体四号（14磅），或关键词条目被增删改写',
            p36 is None or _nows(p36.text) != _nows(kw)
            or _eff_jc(doc, p36) not in ('left', 'both', None)
            or not all(_is_heiti(f) for f in fonts_ea(p36)) or sizes(p36) - {14.0},
            '段%s=%r %s' % (i36, (p36.text if p36 is not None else None),
                            fmt_of(p36) if p36 is not None else ''))

    # ---------- 37 -1 "Abstract" 居中、三号、段前段后15.7磅、2倍行距 ----------
    p37 = paras[abs_idx] if abs_idx is not None else None
    penalty(1, '"Abstract"不满足居中、三号（16磅）、段前段后15.7磅、2倍行距',
            p37 is None or _eff_jc(doc, p37) != 'center' or sizes(p37) - {16.0}
            or not sp_157(p37) or _eff_line_multiple(doc, p37) != 2.0,
            '段%s %s' % (abs_idx, fmt_of(p37) if p37 is not None else '未找到'))

    # ---------- 38 -1 Abstract正文 TNR小四、两端对齐、首行缩进2字符、1.5倍行距 ----------
    i39, p39 = find_para(lambda t: t.lower().startswith('keywords'))
    en_body = [(i, paras[i]) for i in range((abs_idx or 0) + 1,
                                           i39 if i39 else (abs_idx or 0) + 1)
               if (paras[i].text or '').strip()]
    bad38 = not en_body
    for _i, p in en_body:
        if sizes(p) - {12.0} or not first_line_2char(p) \
                or _eff_line_multiple(doc, p) != 1.5 \
                or _eff_jc(doc, p) not in ('both', 'justify') \
                or any('Times New Roman' not in (_eff_ascii(doc, r) or '')
                       for r in p.runs if r.text.strip()):
            bad38 = True
    penalty(1, 'Abstract正文不满足Times New Roman小四（12磅）、两端对齐、首行缩进2字符、1.5倍行距',
            bad38, '%d段：%s' % (len(en_body), [(i, fmt_of(p)) for (i, p) in en_body][:2]))

    # ---------- 39 -1 英文关键词行 左对齐、四号、条目未被增删改写 ----------
    t39 = _nows(p39.text) if p39 is not None else ''
    penalty(1, '"Keywords:  Nature Inquiry; …"一行不满足左对齐、四号（14磅），或英文关键词被增删改写',
            p39 is None or 'NatureInquiry' not in t39 or t39.count(';') < 4
            or _eff_jc(doc, p39) not in ('left', 'both', None) or sizes(p39) - {14.0},
            '段%s=%r %s' % (i39, (p39.text if p39 is not None else None),
                            fmt_of(p39) if p39 is not None else ''))

    # ---------- 40 -3 正文章标题 居中、黑体三号、2倍行距 ----------
    # 只取"一/二/三/四"四章（"总结""致谢"另由第41项判定）
    main_chaps = [(j, p) for (j, p) in chapter_heads
                  if not re.search(r'(总结|致谢)', _nows(p.text))]
    bad40 = []
    for j, p in main_chaps:
        # 输入中"四．优化策略与研究结论"以大首行缩进呈视觉居中，等同居中处理
        centered = _eff_jc(doc, p) == 'center' or (_eff_ind(doc, p, 'firstLine') or 0) >= 1200
        if not centered or not all(_is_heiti(f) for f in fonts_ea(p)) \
                or sizes(p) - {16.0} or _eff_line_multiple(doc, p) != 2.0:
            bad40.append('[%d]%s %s' % (j, _nows(p.text)[:12], fmt_of(p)))
    penalty(3, '正文章标题"一.引  言""二.自然探究活动融入一日生活的理论依据与设计原则""三.活动资源梳理与一日生活融合路径""四．优化策略与研究结论"不满足居中、黑体三号（16磅）、2倍行距',
            len(main_chaps) < 4 or bool(bad40),
            '章标题%d个，不合格%d个%s' % (len(main_chaps), len(bad40), bad40[:2]))

    # ---------- 41 -1 "总结""致谢"标题 居中、三号、2倍行距 ----------
    bad41 = []
    for key in ('总结', '致谢'):
        i, p = find_para(lambda t, k=key: t.endswith(k) and len(t) <= 6, body_start)
        if p is None or _eff_jc(doc, p) != 'center' or sizes(p) - {16.0} \
                or _eff_line_multiple(doc, p) != 2.0:
            bad41.append('%s：%s' % (key, fmt_of(p) if p is not None else '未找到'))
    penalty(1, '"总结""致谢"两个标题不满足居中、三号（16磅）、2倍行距',
            bool(bad41), '；'.join(bad41))

    # ---------- 42 -3 正文二级标题 左对齐、黑体四号、1.5倍行距 ----------
    bad42 = []
    for j, p in section_heads:
        if _eff_jc(doc, p) not in ('left', 'both', None) \
                or not all(_is_heiti(f) for f in fonts_ea(p)) \
                or sizes(p) - {14.0} or _eff_line_multiple(doc, p) != 1.5:
            bad42.append('[%d]%s %s' % (j, _nows(p.text)[:10], fmt_of(p)))
    penalty(3, '正文二级标题"（一）研究背景""（二）研究价值""（三)相关研究述评""（一）理论依据""（二）设计原则""（一）活动资源筛选""（二）生活环节融合切入点""（三）社会领域融合切入点""（四）科学领域融合切入点""（五） 艺术领域融合切入点"不满足左对齐、黑体四号（14磅）、1.5倍行距',
            len(section_heads) < 10 or bool(bad42),
            '二级标题%d个，不合格%d个%s' % (len(section_heads), len(bad42), bad42[:2]))

    # ---------- 43 -1 正文叙述段落 小四、1.5倍行距 ----------
    head_j = set(j for (j, _p) in chapter_heads) | set(j for (j, _p) in section_heads) \
        | set(j for (j, _p) in sub_heads)
    narr, bad43 = [], []
    for j, p in body:
        t = _nows(p.text)
        if j in head_j or len(t) < 20 or t.startswith('•') \
                or _eff_outline(doc, p) is not None:
            continue
        narr.append(j)
        if sizes(p) - {12.0} or _eff_line_multiple(doc, p) != 1.5:
            bad43.append('[%d]%s' % (j, fmt_of(p)))
    # 判定为"批量改乱"而非个别段落：不合格占比超过四分之一才计入减分
    penalty(1, '正文叙述段落不满足小四（12磅）、1.5倍行距',
            bool(narr) and len(bad43) > max(10, len(narr) * 0.25),
            '叙述段%d个，不合格%d个%s' % (len(narr), len(bad43), bad43[:2]))

    # ---------- 44 -1 "•"项目符号被删除或改为其他编号 ----------
    bullets = [_nows(p.text)[:8] for (_i, p) in body if _nows(p.text).startswith('•')]
    penalty(1, '"• 自然观察""• 简单探究""• 数学经验""• 美术活动""• 音乐活动""• 律动活动"等行首的"•"项目符号被删除或改为其他编号形式',
            len(bullets) < 6, '以"•"起首的行=%d条%s' % (len(bullets), bullets[:6]))

    # ---------- 45 -3 批注/修订/高亮/底纹/彩色字体 ----------
    bad45 = []
    bodyel = doc.element.body
    if any(True for _ in bodyel.iter(qn('w:commentReference'))):
        bad45.append('存在批注')
    for tag in ('w:ins', 'w:del', 'w:moveFrom', 'w:moveTo'):
        if any(True for _ in bodyel.iter(qn(tag))):
            bad45.append('存在修订标记(%s)' % tag)
            break
    for p in apar:
        for r in p.runs:
            rpr = r._element.rPr
            if rpr is None:
                continue
            hl = rpr.find(qn('w:highlight'))
            if hl is not None and (hl.get(qn('w:val')) or 'none') != 'none':
                bad45.append('高亮:%r' % r.text[:8])
            shd = rpr.find(qn('w:shd'))
            if shd is not None and (shd.get(qn('w:fill')) or 'auto') not in \
                    ('auto', 'nil', 'FFFFFF', 'ffffff'):
                bad45.append('底纹:%r' % r.text[:8])
            col = rpr.find(qn('w:color'))
            if col is not None and (col.get(qn('w:val')) or 'auto') not in \
                    ('auto', '000000'):
                bad45.append('彩色字:%r=%s' % (r.text[:8], col.get(qn('w:val'))))
    penalty(3, '文档中被插入未经要求的批注、修订标记、文字高亮、底纹或彩色字体',
            bool(bad45), '共%d处%s' % (len(bad45), bad45[:4]))

    # ---------- 46 -1 致谢正文文字内容被改写/删减/扩写 ----------
    i46, p46 = find_para(lambda t: t.startswith('时光流转'), body_start)
    penalty(1, '致谢正文段落（"时光流转，论文写作至此……"）文字内容被改写、删减或扩写',
            p46 is None or len(_nows(p46.text)) < 60,
            '段%s 字数=%s' % (i46, len(_nows(p46.text)) if p46 is not None else None))

    # ================= 反解为结构化条目 =================
    assert len(results) == len(_MAX_DELTAS_ORDER), \
        '细则条目数与 _MAX_DELTAS_ORDER 不一致：%d vs %d' \
        % (len(results), len(_MAX_DELTAS_ORDER))
    dim2_items = []
    total = 0
    for (raw, _detail), max_delta in zip(results, _MAX_DELTAS_ORDER):
        head, _, rule = raw.partition("：")
        try:
            delta = int(head.replace("+", "").strip())
        except ValueError:
            delta = 0
        hit = delta != 0
        dim2_items.append({
            "rule": rule,
            "max_delta": max_delta,
            "delta": max_delta if hit else 0,
            "hit": hit,
        })
        total += max_delta if hit else 0
    result["dim2_items"] = dim2_items
    result["total_score"] = max(0, total)
    result["max_score"] = MAX_SCORE_TOTAL
    return result


if __name__ == "__main__":
    _out = evaluate(os.path.dirname(os.path.abspath(__file__)))
    print(json.dumps(_out, ensure_ascii=False, indent=2))















