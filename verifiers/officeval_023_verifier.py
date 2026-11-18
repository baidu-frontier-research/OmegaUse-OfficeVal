#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# pyright: reportUnknownMemberType=false, reportUnknownVariableType=false, reportUnknownArgumentType=false, reportUnknownParameterType=false, reportMissingParameterType=false, reportPrivateUsage=false, reportUnknownLambdaType=false, reportGeneralTypeIssues=false

import os
import re
import sys
import json
import shutil
import tempfile
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
# PyMuPDF 已从本项目移除（AGPL 与本项目许可证不兼容），运行环境不安装 fitz。
# 改用同目录的 pymupdf_compat（pdfplumber + pypdfium2）读取 Word 导出的 PDF，
# 用于按渲染后的真实分页结果判定页码、跨页、空白页。
try:
    _VERIFIER_DIR = os.path.dirname(os.path.abspath(__file__))
    if _VERIFIER_DIR not in sys.path:
        sys.path.insert(0, _VERIFIER_DIR)
    import pymupdf_compat as fitz
except ImportError:
    fitz = None
from docx.oxml.ns import qn

# 说明：按接口统一约定，本模块不修改全局 sys.stdout，也不 print 主结果。
# 主结果统一由 evaluate(dir_path) -> dict 返回，入参为"脚本所在目录的路径"，
# 由脚本自己在该目录内定位并打开被评估的 .docx。本文件仅在 __main__ 下作为
# 本地调试入口，把该 dict 序列化为 JSON 打印。

SCRIPT_ID = "023"
# 维度二加分项 max_delta 之和，用于维度一未通过 / 脚本异常时兜底
MAX_SCORE_TOTAL = 63
# 维度二扣分项 max_delta 之和（下限）
MIN_SCORE_TOTAL = -106

# 维度二评分项按"结果追加顺序"的 max_delta 列表，用于末端把 results 里的
# "+N：..." / "−N：..." 字符串反解为结构化 dim2_items 时取值。顺序须与
# _evaluate_impl 中 results.append(...) 的顺序严格一致：共 49 项，
# 前 17 项为加分项（累加 = MAX_SCORE_TOTAL），后 32 项为扣分项（累加 = MIN_SCORE_TOTAL）。
_MAX_DELTAS_ORDER = [
    # ---------- 加分项 17 条 / +63 ----------
    3,   # 01 章节导航中不存在任何空白条目
    3,   # 02 目錄条目与正文标题一一对应（含表 目 錄 / 圖 目 錄）
    1,   # 03 目錄中不含"致  謝"条目
    1,   # 04 目錄中不含"作者簡歷"条目
    5,   # 05 致謝末尾为"下一页"分节符，摘要是新分节首页
    3,   # 06 致謝所在页不显示页码
    5,   # 07 前置部分连续大写罗马数字，摘要 = I
    5,   # 08 目錄前置五条齐全且罗马页码全部一致
    3,   # 09 文档中不存在"作者簡歷"标题及其正文
    3,   # 10 正文结束于參考文獻最后一条，其后无空段落、无空白页
    5,   # 11 表目錄为可更新的自动域
    5,   # 12 表目錄完整收录全部 5 个表
    3,   # 13 表目錄条目文字与正文表题逐字一致
    5,   # 14 表目錄 5 条页码全部与实际一致
    3,   # 15 表目錄位于"表 目 錄"标题页，目錄之后圖目錄之前
    5,   # 16 表4-2 首行与末行跨越页数不超过 3 页
    5,   # 17 表4-2 表头位于新一页最上方
    # ---------- 扣分项 32 条 / −106 ----------
    -5,  # 18 封面信息不完整或有误
    -3,  # 19 答辯委員會審議通過表页缺失或不完整
    -3,  # 20 致謝正文不足 4 段或缺署名 / 日期 / 於南海都會大學
    -3,  # 21 摘要或 Abstract 缺正文或缺關鍵字 / Keywords 行
    -5,  # 22 章标题不足 5 个 / 节标题不足 17 个 / 大纲层级错
    -5,  # 23 正文存在段落缺损痕迹
    -3,  # 24 參考文獻少于 8 条或编号不连续
    -5,  # 25 表格数据有缺损
    -3,  # 26 表格行列数不符
    -3,  # 27 5 个表格的格式不统一
    -3,  # 28 图片不足 3 张或图题缺失 / 未居中
    -5,  # 29 目錄不是由 TOC 域自动生成
    -5,  # 30 目錄中正文条目的阿拉伯页码不符
    -3,  # 31 正文格式不符
    -3,  # 32 页面设置不符
    -3,  # 33 正文页眉缺失或文字有误
    -3,  # 34 页脚页码不是小五号 / 不居中
    -3,  # 35 正文页脚页码不是自动 PAGE 域或编号格式错
    -3,  # 36 封面页或審議通過表页显示了页码
    -5,  # 37 存在整页空白页
    -3,  # 38 表格所在页出现大片留白
    -3,  # 39 表2-1 或表4-3 的首行与末行不在同一页
    -3,  # 40 表3-1 的首行与末行不在同一页
    -3,  # 41 表4-1 的首行与末行不在同一页
    -3,  # 42 表4-2 未开启"跨页时重复标题行"
    -3,  # 43 表题与表格不同页，或表4-2 的"註："与表格分离
    -3,  # 44 任一表格列宽合计超出版心 15.28cm
    -3,  # 45 目錄二级条目未设置缩进 2 字符
    -3,  # 46 圖目錄页码不符 / 缺标题 / 条目不足 3 条 / 未独立成页
    -3,  # 47 出现指令未要求的新增内容
    -1,  # 48 出现简体字或中英文标点混用
    -1,  # 49 文档总页数不在 26—29 页区间
]
assert sum(d for d in _MAX_DELTAS_ORDER if d > 0) == MAX_SCORE_TOTAL
assert sum(d for d in _MAX_DELTAS_ORDER if d < 0) == MIN_SCORE_TOTAL
assert len(_MAX_DELTAS_ORDER) == 49

# ==========================================================================
# 字体解析：run 的 rPr/rFonts → 段落 pStyle 链（沿 w:basedOn 上溯）
# → docDefaults/rPrDefault。返回第一个命中的字体名，全部缺省时返回 None。
# ==========================================================================

def _style_element_by_id(doc, style_id):
    if not style_id:
        return None
    styles_root = doc.styles.element
    for st in styles_root.findall(qn('w:style')):
        if st.get(qn('w:styleId')) == style_id:
            return st
    return None


def _para_style_id(para):
    pPr = para._element.find(qn('w:pPr'))
    if pPr is None:
        return None
    pStyle = pPr.find(qn('w:pStyle'))
    if pStyle is None:
        return None
    return pStyle.get(qn('w:val'))


def _walk_style_chain(doc, style_id, picker, depth=0):
    """沿 w:basedOn 上溯样式链，返回 picker(style_element) 的第一个非 None 结果。"""
    if depth > 12:
        return None
    st = _style_element_by_id(doc, style_id)
    if st is None:
        return None
    got = picker(st)
    if got is not None:
        return got
    basedOn = st.find(qn('w:basedOn'))
    if basedOn is None:
        return None
    return _walk_style_chain(doc, basedOn.get(qn('w:val')), picker, depth + 1)


def _default_para_style_id(doc):
    """文档默认段落样式（w:type="paragraph" 且 w:default="1"）的 styleId。
    正文段落往往不写 w:pStyle，字号 / 行距 / 缩进全部由它提供，必须纳入解析链。"""
    for st in doc.styles.element.findall(qn('w:style')):
        if st.get(qn('w:type')) == 'paragraph' \
                and st.get(qn('w:default')) in ('1', 'true', 'on'):
            return st.get(qn('w:styleId'))
    return None


def _doc_defaults_rpr(doc):
    styles_root = doc.styles.element
    dd = styles_root.find(qn('w:docDefaults'))
    if dd is None:
        return None
    rd = dd.find(qn('w:rPrDefault'))
    if rd is None:
        return None
    return rd.find(qn('w:rPr'))


def _rfonts_attr(rpr, attr):
    if rpr is None:
        return None
    rf = rpr.find(qn('w:rFonts'))
    if rf is None:
        return None
    return rf.get(qn(attr))


def _resolve_effective_eastasia(doc, run):
    """解析 run 生效的中文（eastAsia）字体名。"""
    got = _rfonts_attr(run._element.find(qn('w:rPr')), 'w:eastAsia')
    if got:
        return got
    sid = _para_style_id(run._parent) if hasattr(run, '_parent') else None
    sid = sid or _default_para_style_id(doc)
    if sid:
        got = _walk_style_chain(
            doc, sid,
            lambda st: _rfonts_attr(st.find(qn('w:rPr')), 'w:eastAsia'))
        if got:
            return got
    return _rfonts_attr(_doc_defaults_rpr(doc), 'w:eastAsia')


def _resolve_effective_ascii(doc, run):
    """解析 run 生效的西文（ascii）字体名。"""
    got = _rfonts_attr(run._element.find(qn('w:rPr')), 'w:ascii')
    if got:
        return got
    sid = _para_style_id(run._parent) if hasattr(run, '_parent') else None
    sid = sid or _default_para_style_id(doc)
    if sid:
        got = _walk_style_chain(
            doc, sid,
            lambda st: _rfonts_attr(st.find(qn('w:rPr')), 'w:ascii'))
        if got:
            return got
    return _rfonts_attr(_doc_defaults_rpr(doc), 'w:ascii')


def _resolve_effective_size(doc, para, run):
    """解析 run 生效的字号（磅）。w:sz 为半磅值。"""
    for rpr in (run._element.find(qn('w:rPr')),):
        if rpr is not None:
            sz = rpr.find(qn('w:sz'))
            if sz is not None and sz.get(qn('w:val')):
                return float(sz.get(qn('w:val'))) / 2.0
    sid = _para_style_id(para) or _default_para_style_id(doc)
    if sid:
        got = _walk_style_chain(
            doc, sid,
            lambda st: (st.find(qn('w:rPr')).find(qn('w:sz'))
                        if st.find(qn('w:rPr')) is not None
                        and st.find(qn('w:rPr')).find(qn('w:sz')) is not None
                        else None))
        if got is not None and got.get(qn('w:val')):
            return float(got.get(qn('w:val'))) / 2.0
    dd = _doc_defaults_rpr(doc)
    if dd is not None:
        sz = dd.find(qn('w:sz'))
        if sz is not None and sz.get(qn('w:val')):
            return float(sz.get(qn('w:val'))) / 2.0
    return None


# ==========================================================================
# 文档打开与 PDF 渲染
# ==========================================================================

def _open_word_document(file_path):
    return Document(file_path)


def _render_to_pdf(file_path, tmp_dir):
    """用 Word / WPS 的 COM 接口把 .docx 导出为 PDF（只读打开，绝不回写）。
    返回生成的 PDF 路径；任何一步失败都抛 RuntimeError。"""
    errors = []
    out_pdf = os.path.join(tmp_dir, 'render.pdf')
    try:
        import win32com.client  # type: ignore
    except ImportError as exc:
        raise RuntimeError('缺少 pywin32，无法渲染 PDF：%s' % exc)
    word = None
    doc = None
    try:
        word = win32com.client.DispatchEx('Word.Application')
        word.Visible = False
        word.DisplayAlerts = 0
        doc = word.Documents.Open(os.path.abspath(file_path), ReadOnly=True,
                                  AddToRecentFiles=False)
        doc.ExportAsFixedFormat(OutputFileName=out_pdf, ExportFormat=17)
    except Exception as exc:
        errors.append('%s: %s' % (type(exc).__name__, exc))
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
        raise RuntimeError('; '.join(errors) or '无法渲染 PDF')
    return out_pdf


def _norm(s):
    """归一化：去掉所有空白（含全角空格）与制表符，便于逐字比对标题文字。"""
    return re.sub(r'[\s　 \x0b\x0c]+', '', s or '')


def _build_page_model(file_path):
    """渲染 PDF 后逐页抽取：正文文字、页脚显示的页码串、正文底部纵坐标、
    是否含图片。任何环节不可用时返回 None（页码类判定随之全部跳过，不误判）。"""
    if fitz is None:
        return None
    tmp_dir = tempfile.mkdtemp(prefix='oev023_')
    try:
        pdf_path = _render_to_pdf(file_path, tmp_dir)
        pages = []
        with fitz.open(pdf_path) as pdf:
            for page in pdf:
                ph = float(page.rect.height)
                # 版心上边界 2.54cm≈72pt、下边界≈769.9pt；页眉在 1.4cm≈40pt，
                # 页脚在距底 1.27cm≈36pt。阈值取 0.085 / 0.925 以便把页眉页脚
                # 剥离干净，同时不误伤版心内的首行与末行。
                head_y = ph * 0.085
                foot_y = ph * 0.925
                blocks = page.get_text('blocks') or []
                body_txt = []
                foot_txt = []
                bottom = 0.0
                top = None
                first_block = ''
                for b in blocks:
                    x0, y0, x1, y1, txt = b[0], b[1], b[2], b[3], (b[4] or '')
                    if not txt.strip():
                        continue
                    if y0 >= foot_y:
                        foot_txt.append(txt.strip())
                        continue
                    if y1 <= head_y:
                        continue
                    body_txt.append(txt)
                    bottom = max(bottom, float(y1))
                    if top is None or float(y0) < top:
                        top = float(y0)
                        first_block = txt.strip()
                try:
                    n_img = len(page.get_images(full=True) or [])
                except Exception:
                    n_img = 0
                pages.append({
                    'text': ''.join(body_txt),
                    'norm': _norm(''.join(body_txt)),
                    'footer': _norm(' '.join(foot_txt)),
                    'bottom': bottom,
                    'top': top if top is not None else 0.0,
                    'first': _norm(first_block),
                    'images': n_img,
                    'height': ph,
                })
        return pages or None
    except Exception:
        return None
    finally:
        shutil.rmtree(tmp_dir, ignore_errors=True)


def _make_locator(pages):
    """返回 _locate(snippet, from_page=0) -> 1 基页号；找不到返回 0。
    带单调游标，避免同一文字在目錄与正文各出现一次时定位到前面的目錄条目。"""
    state = {'cursor': 0}

    def _locate(snippet, from_page=None, advance=True):
        if not pages:
            return 0
        key = _norm(snippet)
        if not key:
            return 0
        start = state['cursor'] if from_page is None else max(0, from_page)
        for i in range(start, len(pages)):
            if key in pages[i]['norm']:
                if advance:
                    state['cursor'] = i
                return i + 1
        for i in range(0, len(pages)):
            if key in pages[i]['norm']:
                return i + 1
        return 0

    def _reset(page_index=0):
        state['cursor'] = max(0, page_index)

    return _locate, _reset


def _locate_nth(pages, key, ordinal):
    """按"出现次序"定位：返回归一化串 key 在渲染结果中第 ordinal 次出现所在的
    1 基页号。PDF 里的文字出现顺序与 body 段落顺序一致，因此同一文字在目錄条目
    与正文标题各出现一次时也能分辨。找不到返回 0。"""
    if not pages or not key or ordinal <= 0:
        return 0
    seen = 0
    for i, pg in enumerate(pages):
        c = pg['norm'].count(key)
        if c and seen + c >= ordinal:
            return i + 1
        seen += c
    return 0


def _build_para_page_map(pages, paras, key_of):
    """为每个段落算出它渲染后所在的页号，基于 _locate_nth 的出现次序法。"""
    if not pages:
        return {}
    ordinals = {}
    out = {}
    for p in paras:
        key = key_of(p)
        if not key:
            continue
        ordinals[key] = ordinals.get(key, 0) + 1
        out[id(p._element)] = _locate_nth(pages, key, ordinals[key])
    return out


_ROMAN = ['I', 'II', 'III', 'IV', 'V', 'VI', 'VII', 'VIII', 'IX', 'X',
          'XI', 'XII', 'XIII', 'XIV', 'XV', 'XVI', 'XVII', 'XVIII', 'XIX', 'XX']





def _para_texts(doc):
    return [p.text for p in doc.paragraphs]


def _is_page_break_only(para):
    """段落只含分页符（w:br type=page）而无任何文字：Word / WPS 的导航窗格
    不会把它列为条目，判定空白条目时必须先剔除。"""
    el = para._element
    has_text = any((t.text or '').strip() for t in el.iter(qn('w:t')))
    if has_text:
        return False
    for br in el.iter(qn('w:br')):
        if br.get(qn('w:type')) in ('page', 'column'):
            return True
    return False


_HEADING_STYLE_RE = re.compile(r'^(?:heading|标题|標題)\s*([1-9])$', re.I)
_TOC_STYLE_RE = re.compile(r'^(?:toc|目录|目錄)\s*([1-9])$', re.I)
# 表目錄 / 圖目錄的条目用的是"图表目录"样式（英文名 table of figures，
# styleId 常见 TableofFigures / 29 之类的数字 id），它同样是目錄条目，
# 只有一级，必须一并识别，否则条目会被当成正文里的表题 / 图题。
_TOF_STYLE_RE = re.compile(
    r'^(?:tableoffigures|图表目录|圖表目錄|表格目录|表格目錄|图表目次|圖表目次)$',
    re.I)


def _outline_level(doc, para):
    """段落大纲级别：1–9 表示会出现在导航窗格；None 表示正文级。
    读取顺序：段落直接 pPr/outlineLvl → 样式链 outlineLvl → 样式名 标题N。"""
    pPr = para._element.find(qn('w:pPr'))
    if pPr is not None:
        ol = pPr.find(qn('w:outlineLvl'))
        if ol is not None and ol.get(qn('w:val')) is not None:
            v = int(ol.get(qn('w:val')))
            return None if v >= 9 else v + 1
    sid = _para_style_id(para)
    if sid:
        got = _walk_style_chain(
            doc, sid,
            lambda st: (st.find(qn('w:pPr')).find(qn('w:outlineLvl'))
                        if st.find(qn('w:pPr')) is not None
                        and st.find(qn('w:pPr')).find(qn('w:outlineLvl')) is not None
                        else None))
        if got is not None and got.get(qn('w:val')) is not None:
            v = int(got.get(qn('w:val')))
            return None if v >= 9 else v + 1
    try:
        name = para.style.name or ''
    except Exception:
        name = ''
    m = _HEADING_STYLE_RE.match(name.strip())
    if m:
        return int(m.group(1))
    return None


def _toc_level(para):
    """目錄条目级别：目录 1 / 目录 2 / TOC 1 / 图表目录 ... 返回 1–9；否则 None。"""
    try:
        name = (para.style.name or '').strip()
    except Exception:
        name = ''
    m = _TOC_STYLE_RE.match(name)
    if m:
        return int(m.group(1))
    if _TOF_STYLE_RE.match(name.replace(' ', '')):
        return 1
    sid = _para_style_id(para) or ''
    m = re.match(r'^TOC(\d)$', sid, re.I)
    if m:
        return int(m.group(1))
    if _TOF_STYLE_RE.match(sid.replace(' ', '')):
        return 1
    return None


def _split_toc_entry(text):
    """把目錄条目拆成（条目文字, 页码串）。制表符 / 连续点前导符都当分隔。"""
    t = (text or '').replace('', ' ')
    t = re.sub(r'\.{3,}', '\t', t)
    if '\t' in t:
        label, _, page = t.rpartition('\t')
    else:
        m = re.match(r'^(.*?)[\s　]*([IVXLCDMivxlcdm]+|\d+)[\s　]*$', t)
        if not m:
            return t.strip(), ''
        label, page = m.group(1), m.group(2)
    return label.strip(), page.strip()


def _para_section_index(doc):
    """返回 {段落 XML 元素 id: 分节序号(0 基)}，分节序号与 doc.sections 对齐。

    正常 OOXML 要求最后一节的 w:sectPr 直接挂在 body 下（不在任何段落里）；
    若交付文件缺失这个收尾 sectPr（结构性损坏），body 内出现的段落级 sectPr
    数量会等于 doc.sections 的总数，导致最后一段之后的段落被顺势记成一个
    不存在的"下一节"。这里钳制到最后一节，避免调用方按此下标访问
    doc.sections 时越界——损坏文件应在维度一判负，不应造成脚本异常。"""
    mapping = {}
    idx = 0
    last = max(0, len(doc.sections) - 1)
    for child in doc.element.body.iterchildren():
        if child.tag == qn('w:p'):
            mapping[id(child)] = min(idx, last)
            pPr = child.find(qn('w:pPr'))
            if pPr is not None and pPr.find(qn('w:sectPr')) is not None:
                idx += 1
        elif child.tag == qn('w:tbl'):
            for p in child.iter(qn('w:p')):
                mapping.setdefault(id(p), min(idx, last))
    return mapping


def _sect_pr(section):
    return section._sectPr


def _pg_num_type(section):
    """返回 (fmt, start)，取自 sectPr/pgNumType。"""
    pn = _sect_pr(section).find(qn('w:pgNumType'))
    if pn is None:
        return None, None
    fmt = pn.get(qn('w:fmt'))
    start = pn.get(qn('w:start'))
    return fmt, (int(start) if start is not None else None)


def _section_break_type(section):
    """本节结尾的分节符类型：nextPage / continuous / evenPage / oddPage。
    w:type 缺省即 nextPage。"""
    t = _sect_pr(section).find(qn('w:type'))
    if t is None:
        return 'nextPage'
    return t.get(qn('w:val')) or 'nextPage'


def _effective_part(doc, sec_idx, kind):
    """解析第 sec_idx 节生效的页眉 / 页脚对象（沿"链接到前一节"上溯）。
    kind = 'header' | 'footer'。返回 (part, 实际生效的节序号) 或 (None, -1)。

    sec_idx 由段落里数 w:sectPr 个数推算而来；若交付文件段落级 sectPr 数量
    多于 python-docx 实际解析出的节数（分节结构异常），须钳制到最后一节，
    否则 doc.sections[i] 会越界抛 IndexError，把"文件本身分节混乱"误判成
    脚本运行错误。"""
    n = len(doc.sections)
    if n == 0:
        return None, -1
    i = min(sec_idx, n - 1)
    while i >= 0:
        sec = doc.sections[i]
        part = sec.header if kind == 'header' else sec.footer
        try:
            linked = bool(part.is_linked_to_previous)
        except Exception:
            linked = False
        if not linked:
            return part, i
        i -= 1
    return None, -1


def _part_text(part):
    if part is None:
        return ''
    try:
        return ''.join(p.text for p in part.paragraphs)
    except Exception:
        return ''


def _part_field_codes(part):
    """页眉 / 页脚里的域代码文本（w:instrText + w:fldSimple/@w:instr）。"""
    if part is None:
        return []
    codes = []
    el = part._element
    for it in el.iter(qn('w:instrText')):
        codes.append((it.text or '').strip())
    for fs in el.iter(qn('w:fldSimple')):
        codes.append((fs.get(qn('w:instr')) or '').strip())
    return [c for c in codes if c]


def _doc_field_codes(doc):
    codes = []
    for it in doc.element.body.iter(qn('w:instrText')):
        codes.append((it.text or '').strip())
    for fs in doc.element.body.iter(qn('w:fldSimple')):
        codes.append((fs.get(qn('w:instr')) or '').strip())
    return [c for c in codes if c]


# ==========================================================================
# 维度一：可用或可修改性（门槛项，不满足即判定为未完成）
# ==========================================================================

_BAD_CHARS_RE = re.compile(r'[�￾￿□■]|[-]')


def _check_dim1(doc, file_path):
    """返回 (pass, reason)。两条门槛：
    ① 结果文件为 .docx，打开时不提示"内容有问题需要修复"，正文无乱码 / 缺字 / 方框
       —— 用 python-docx 能完整解析 body 视作可正常打开；再扫全文替换字符
       U+FFFD、方框 □/■、私用区码位。
    ② 文件未加密、未设为只读，文档中无残留的修订标记、批注、高亮
       —— 加密文件 python-docx 直接打不开；只读读 settings.xml 的
       w:documentProtection / w:writeProtection；修订读 w:ins/w:del/
       w:rPrChange/w:pPrChange/w:tblPrChange；批注读 w:commentReference；
       高亮读 w:highlight（非 none）。"""
    reasons = []
    if not file_path.lower().endswith('.docx'):
        reasons.append('文件不是 .docx 格式')

    full = []
    for p in doc.paragraphs:
        full.append(p.text)
    for t in doc.tables:
        for row in t.rows:
            for c in row.cells:
                full.append(c.text)
    body_text = '\n'.join(full)
    if _BAD_CHARS_RE.search(body_text):
        reasons.append('正文存在乱码 / 缺字 / 方框字符')

    body = doc.element.body
    rev_tags = ('w:ins', 'w:del', 'w:rPrChange', 'w:pPrChange',
                'w:tblPrChange', 'w:tcPrChange', 'w:sectPrChange')
    n_rev = 0
    for tag in rev_tags:
        n_rev += sum(1 for _ in body.iter(qn(tag)))
    if n_rev:
        reasons.append('文档中残留修订标记 %d 处' % n_rev)

    n_cmt = sum(1 for _ in body.iter(qn('w:commentReference')))
    if n_cmt:
        reasons.append('文档中残留批注 %d 处' % n_cmt)

    n_hl = 0
    for hl in body.iter(qn('w:highlight')):
        if (hl.get(qn('w:val')) or 'none') != 'none':
            n_hl += 1
    if n_hl:
        reasons.append('文档中残留高亮 %d 处' % n_hl)

    try:
        settings = doc.settings.element
        dp = settings.find(qn('w:documentProtection'))
        if dp is not None and (dp.get(qn('w:enforcement')) in ('1', 'true', 'on')):
            reasons.append('文档设置了强制保护 / 只读')
        wp = settings.find(qn('w:writeProtection'))
        if wp is not None and (wp.get(qn('w:recommended')) in ('1', 'true', 'on')
                               or wp.get(qn('w:hash')) is not None):
            reasons.append('文档设置了写保护')
    except Exception:
        pass

    return (not reasons), '；'.join(reasons)


# ==========================================================================
# 对外唯一入口
# ==========================================================================

def _is_damaged_package(exc):
    """判断异常是否源于交付文件自身结构损坏（缺部件、XML 不合法等）。"""
    text = str(exc)
    if isinstance(exc, KeyError) and (
        "There is no item named" in text or "no relationship of type" in text
    ):
        return True
    return type(exc).__name__ in ("XMLSyntaxError", "ParseError", "BadZipFile",
                                  "InvalidXmlError")


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
            raise FileNotFoundError('目录不存在：%s' % dir_path)
        cands = sorted(f for f in os.listdir(dir_path)
                       if f.lower().endswith('.docx') and not f.startswith('~$'))
        if not cands:
            raise FileNotFoundError('目录内未找到 .docx 文件：%s' % dir_path)
        file_path = os.path.join(dir_path, cands[0])
        result["file_name"] = os.path.basename(file_path)
        return _evaluate_impl(file_path, result)
    except Exception as exc:
        if _is_damaged_package(exc):
            result["status"] = "ok"
            result["error"] = None
            result["dim1_pass"] = False
            result["dim1_reason"] = "交付文件结构不完整，无法正常打开：%s: %s" % (
                type(exc).__name__, exc)
            result["dim2_items"] = []
            result["total_score"] = 0
            return result
        # status 为 error 时未能完成判定，dim1_pass 留空，dim1_reason 也为空
        result["status"] = "error"
        result["error"] = "%s: %s" % (type(exc).__name__, exc)
        result["dim1_pass"] = ""
        result["dim1_reason"] = ""
        result["dim2_items"] = []
        result["total_score"] = 0
        result["max_score"] = MAX_SCORE_TOTAL
        return result


_TBL_CAP_RE = re.compile(r'^表\s*(\d+)\s*-\s*(\d+)')
_FIG_CAP_RE = re.compile(r'^圖\s*(\d+)\s*-\s*(\d+)')
_CHAP_RE = re.compile(r'^第[一二三四五六七八九十]+章')
_SEC_RE = re.compile(r'^\d+\.\d+(?:[\s　]|$)')
_FRONT_TITLES = ('摘要', 'Abstract', '目錄', '表目錄', '圖目錄',
                 '致謝', '作者簡歷', '參考文獻')

# 只收录"简化字形与繁体字形确实不同"的常用字，避免把两体通用字误判成简体。
_SIMPLIFIED = set(
    '这说会学国现实认识经济响绩华门题变数据处业转为与关键词图参献论讨严谨设计资'
    '检场养层结构应该时间问类别样确义试验证显单双众决术传统组织员团队责权财务产'
    '电报导领势战规划创发环节约减价营销费户际张长开联释选择边归纳综议讲让观觉视'
    '语读训记达过还进远连运违适递迁银钱钟针铁误谈诉调请课谁谓谋谱贵贷贸赛赞质'
)


def _body_order(doc):
    """按 body 顺序返回 [('p', 段落对象), ('tbl', 表格对象), ...]。"""
    p_by_el = {id(p._element): p for p in doc.paragraphs}
    t_by_el = {id(t._element): t for t in doc.tables}
    seq = []
    for child in doc.element.body.iterchildren():
        if child.tag == qn('w:p') and id(child) in p_by_el:
            seq.append(('p', p_by_el[id(child)]))
        elif child.tag == qn('w:tbl') and id(child) in t_by_el:
            seq.append(('tbl', t_by_el[id(child)]))
    return seq


def _table_caption_map(seq):
    """每个表格向上找最近的一个"表N-N ..."段落作为表题。
    返回 [(表格对象, 表题段落或 None, 表号串或 '')]，按文档顺序。"""
    out = []
    for i, (kind, obj) in enumerate(seq):
        if kind != 'tbl':
            continue
        cap = None
        for j in range(i - 1, max(-1, i - 8), -1):
            if seq[j][0] != 'p':
                continue
            txt = _norm(seq[j][1].text)
            if _TBL_CAP_RE.match(txt):
                cap = seq[j][1]
                break
            if txt:
                # 遇到非空的非表题段落即停止，避免跨表错配
                if not _TBL_CAP_RE.match(txt):
                    continue
        num = ''
        if cap is not None:
            m = _TBL_CAP_RE.match(_norm(cap.text))
            if m:
                num = '表%s-%s' % (m.group(1), m.group(2))
        out.append((obj, cap, num))
    return out


def _resolve_ppr_child(doc, para, tag):
    """段落属性：直接格式优先，其次沿样式链上溯。返回 w:pPr 下 tag 的元素。"""
    pPr = para._element.find(qn('w:pPr'))
    if pPr is not None:
        el = pPr.find(qn(tag))
        if el is not None:
            return el
    sid = _para_style_id(para) or _default_para_style_id(doc)
    if not sid:
        return None
    return _walk_style_chain(
        doc, sid,
        lambda st: (st.find(qn('w:pPr')).find(qn(tag))
                    if st.find(qn('w:pPr')) is not None
                    and st.find(qn('w:pPr')).find(qn(tag)) is not None
                    else None))


def _para_jc(doc, para):
    el = _resolve_ppr_child(doc, para, 'w:jc')
    return (el.get(qn('w:val')) if el is not None else None)


def _twips(v):
    try:
        return float(v)
    except (TypeError, ValueError):
        return None


def _para_ind(doc, para):
    """返回 {left, leftChars, firstLine, firstLineChars, hanging, hangingChars}。"""
    el = _resolve_ppr_child(doc, para, 'w:ind')
    out = {}
    if el is None:
        return out
    for k in ('left', 'leftChars', 'firstLine', 'firstLineChars',
              'hanging', 'hangingChars', 'start', 'startChars'):
        v = el.get(qn('w:' + k))
        if v is not None:
            out[k] = _twips(v)
    return out


def _para_line_spacing(doc, para):
    """返回 (line, lineRule)。line 为二十分之一磅或倍数×240。"""
    el = _resolve_ppr_child(doc, para, 'w:spacing')
    if el is None:
        return None, None
    return _twips(el.get(qn('w:line'))), el.get(qn('w:lineRule'))


# ==========================================================================
# 维度二：完成度
# ==========================================================================

def _evaluate_impl(file_path, result):
    doc = _open_word_document(file_path)

    dim1_ok, dim1_reason = _check_dim1(doc, file_path)
    result["dim1_pass"] = dim1_ok
    result["dim1_reason"] = dim1_reason
    if not dim1_ok:
        result["dim2_items"] = []
        result["total_score"] = 0
        result["max_score"] = MAX_SCORE_TOTAL
        return result

    results = []

    # rule 只放"规则表述"本身：实测值 / 抽检结果不写进返回结果。
    def _plus(n, ok, text):
        results.append(('+%d：' % n if ok else '+0：') + text)

    def _minus(n, bad, text):
        results.append(('−%d：' % n if bad else '−0：') + text)

    # ---------- 共用数据 ----------
    seq = _body_order(doc)
    paras = [obj for kind, obj in seq if kind == 'p']
    sec_map = _para_section_index(doc)
    in_table = set()
    for t in doc.tables:
        for p in t._element.iter(qn('w:p')):
            in_table.add(id(p))

    def is_in_table(para):
        return id(para._element) in in_table

    toc_lv = {}
    for p in paras:
        lv = _toc_level(p)
        if lv:
            toc_lv[id(p._element)] = lv

    def is_toc_entry(para):
        return id(para._element) in toc_lv

    # 大纲标题（会出现在导航窗格 / 被 TOC 域收录的段落）
    headings = []          # [(body 序号, para, level, 归一化文字)]
    for i, p in enumerate(paras):
        if is_in_table(p) or is_toc_entry(p):
            continue
        lv = _outline_level(doc, p)
        txt = _norm(p.text)
        if lv and txt:
            headings.append((i, p, lv, txt))

    def find_heading(name):
        for h in headings:
            if h[3] == name:
                return h
        # 与办公软件对齐：把某行标题从目錄里去掉的常规做法是把它的大纲级别改成
        # "正文"（w:outlineLvl=9）。此时它不再是大纲标题，但页面上那行标题仍在，
        # 定位它必须回退到按文字匹配，否则依赖该标题的检查会全部落空。
        for i, p in enumerate(paras):
            if is_in_table(p) or is_toc_entry(p):
                continue
            if _norm(p.text) == name:
                return (i, p, 0, name)
        return None

    # 三个"目錄"标题的 body 位置
    pos_toc = pos_tbl_toc = pos_fig_toc = -1
    for i, p in enumerate(paras):
        if is_toc_entry(p) or is_in_table(p):
            continue
        t = _norm(p.text)
        if t == '目錄' and pos_toc < 0:
            pos_toc = i
        elif t == '表目錄' and pos_tbl_toc < 0:
            pos_tbl_toc = i
        elif t == '圖目錄' and pos_fig_toc < 0:
            pos_fig_toc = i

    # 目錄条目按所属"目錄标题"分组
    grp_toc, grp_tbl, grp_fig = [], [], []
    for i, p in enumerate(paras):
        if not is_toc_entry(p):
            continue
        txt = (p.text or '').strip()
        if not txt:
            continue
        item = (i, p, _toc_level(p)) + _split_toc_entry(p.text)
        if pos_fig_toc >= 0 and i > pos_fig_toc:
            grp_fig.append(item)
        elif pos_tbl_toc >= 0 and i > pos_tbl_toc:
            grp_tbl.append(item)
        else:
            grp_toc.append(item)

    # 表格 / 表题 / 图题
    tbl_caps = _table_caption_map(seq)
    captions = [(i, p, _norm(p.text)) for i, p in enumerate(paras)
                if not is_in_table(p) and not is_toc_entry(p)
                and _TBL_CAP_RE.match(_norm(p.text))]
    figcaps = [(i, p, _norm(p.text)) for i, p in enumerate(paras)
               if not is_in_table(p) and not is_toc_entry(p)
               and _FIG_CAP_RE.match(_norm(p.text))]

    field_codes = _doc_field_codes(doc)
    body_all_text = '\n'.join(p.text for p in paras)

    # 渲染分页模型：页码 / 跨页 / 空白页 / 留白 类判定全部基于它；
    # 渲染不可用时这些项一律按"无法测量"处理（加分项不给分、扣分项不扣分）。
    pages = _build_page_model(file_path)
    n_pages = len(pages) if pages else 0
    locate, reset_cursor = _make_locator(pages)

    def _key_of(p):
        if is_toc_entry(p):
            s = _norm(_split_toc_entry(p.text)[0])
        else:
            s = _norm(p.text)
        return s[:24]

    ppage = _build_para_page_map(pages, paras, _key_of)

    def footer_label(pg):
        if not pages or pg <= 0 or pg > len(pages):
            return ''
        return pages[pg - 1]['footer']

    def page_of_para(para, hint=None):
        pg = ppage.get(id(para._element), 0)
        if pg:
            return pg
        txt = (para.text or '').strip()
        if not txt:
            return 0
        return locate(txt[:40], from_page=hint, advance=False)

    def page_of_caption(para):
        """表题 / 图题渲染后所在页：按去掉编号的题目正文定位，取最后一次出现。

        Word 重新计算 SEQ 域后，正文渲染出的表号可能与 OOXML 缓存文字（也与
        表目錄条目）不一致。此时按"表4-2……"整串去找必然落空，再退回可回绕的
        全文搜索就会命中前面的表目錄条目，把前置页的罗马页码当成正文实际页码。
        改用稳定的题目正文定位，并取最后一次出现（表目錄条目在正文题注之前）。
        """
        if not pages or para is None:
            return 0
        text = _norm(para.text)
        tail = re.sub(r'^[表圖]\d+[-－]\d+', '', text)
        if tail:
            hits = [i + 1 for i, pg in enumerate(pages) if tail in pg['norm']]
            if hits:
                return hits[-1]
        return page_of_para(para)

    def page_of_heading(name):
        h = find_heading(name)
        if h is None:
            return 0
        return page_of_para(h[1])

    def row_keys(t, r):
        """按长度返回该行可用于渲染定位的单元格文本。

        表头常含"模型設定"等会在正文中重复出现的短词，不能把单个短词拿到全文
        做可回绕搜索；调用方必须从表题页开始、只向后定位。
        """
        keys = []
        try:
            cells = t.rows[r].cells
        except Exception:
            return keys
        for c in cells:
            s = _norm(c.text)
            if s and s not in keys:
                keys.append(s)
        return sorted(keys, key=len, reverse=True)

    def locate_forward(keys, start_page, min_matches=1):
        """从指定物理页起只向后查找，不回绕到目錄或正文中的早期同名文本。"""
        if not pages:
            return 0
        start = max(0, (start_page or 1) - 1)
        for i in range(start, len(pages)):
            matched = sum(1 for key in keys if key and key in pages[i]['norm'])
            if matched >= min(min_matches, len(keys)):
                return i + 1
        return 0

    def last_row_page(head_keys, tail_keys, p1):
        """末行所在页：先用"只出现在一页"的末行文本定位，再退回连续页扩展。

        末行里像"行業/年度"这类短文本会在表格的每一页重复出现，取首个命中页会
        把跨页表格误判成只占一页。开启"跨页重复标题行"时表头也会逐页重复，因此
        退化路径按"仍能命中表头或末行文本"的连续页向后扩展，遇到第一个不命中的
        页即停。
        """
        if not pages or not p1:
            return 0
        unique = [k for k in tail_keys
                  if sum(1 for pg in pages if k in pg['norm']) == 1]
        if unique:
            hits = [i + 1 for i in range(p1 - 1, len(pages))
                    if any(k in pages[i]['norm'] for k in unique)]
            if hits:
                return max(hits)
        last = p1
        for i in range(p1 - 1, len(pages)):
            norm = pages[i]['norm']
            head_hit = sum(1 for k in head_keys if k in norm) >= 2
            tail_hit = any(k in norm for k in tail_keys)
            if head_hit or tail_hit:
                last = i + 1
            else:
                break
        return last

    def table_span(t, hint=None):
        """返回 (首行页, 末行页)；无法定位时返回 (0, 0)。"""
        try:
            nrow = len(t.rows)
        except Exception:
            return 0, 0
        if nrow == 0:
            return 0, 0
        k1 = row_keys(t, 0)
        k2 = row_keys(t, nrow - 1)
        # 表头中的单个短词可能出现在表题页的正文里；至少同时命中三个不同的
        # 表头单元格，才把该页认作表格首行所在页。
        p1 = locate_forward(k1, hint or 1, min_matches=3)
        p2 = last_row_page(k1, k2, p1) if p1 else 0
        return p1, p2

    # 表号 → (表格, 表题段落, 表题页, 首行页, 末行页)
    tbl_info = {}
    for t, cap, num in tbl_caps:
        cap_pg = page_of_caption(cap) if cap is not None else 0
        hint = cap_pg or 1
        p1, p2 = table_span(t, hint)
        key = num or ('表#%d' % (len(tbl_info) + 1))
        tbl_info[key] = {'t': t, 'cap': cap, 'cap_pg': cap_pg,
                         'first': p1, 'last': p2}

    def ti(num):
        return tbl_info.get(num)

    # ======================= 加分项（17 条 / +63） =======================

    # 01 +3 章节导航中不存在任何空白条目
    # 严格对齐细则：①凡大纲级别 1—9 的段落都会出现在导航窗格；②文字为空即空白
    # 条目；③只含分页符（w:br type="page"）而无文字的段落不算条目，须先剔除。
    # 与办公软件对齐：读 w:pPr/w:outlineLvl（缺省时沿样式链上溯至 标题N），
    # 文字读 w:t，分页符读 w:br/@w:type。另外，带 w:sectPr 的空段落是"分节符"
    # 那一行本身，Word / WPS 的导航窗格不把它当条目列出（它常从被拆分的标题继承
    # outlineLvl），必须同分页符段落一并剔除。
    blank_nav = 0
    for p in paras:
        if is_in_table(p) or is_toc_entry(p):
            continue
        if not _outline_level(doc, p):
            continue
        if _norm(p.text):
            continue
        if _is_page_break_only(p):
            continue
        pPr_ = p._element.find(qn('w:pPr'))
        if pPr_ is not None and pPr_.find(qn('w:sectPr')) is not None:
            continue
        blank_nav += 1
    _plus(3, blank_nav == 0, '章节导航中不存在任何空白条目')

    # 02 +3 目錄条目与正文标题一一对应，且含"表 目 錄""圖 目 錄"两条，文字逐字一致
    # 严格对齐细则：①按 TOC 域的 \o "a-b" 开关确定应收录的大纲级别；②该级别内
    # 的全部标题（"致  謝""作者簡歷"除外，由 03/04/09 单独计）须与目錄条目顺序
    # 一致、逐字一致；③不得有重复条目；④必须含"表 目 錄""圖 目 錄"。
    # 与办公软件对齐：条目读 目录N / TOCN 样式段落，标题读 w:outlineLvl。
    toc_code = ''
    for c in field_codes:
        if c.upper().lstrip().startswith('TOC') and '\\c' not in c.lower():
            toc_code = c
            break
    lo, hi = 1, 9
    m = re.search(r'\\o\s*"?\s*(\d)\s*-\s*(\d)', toc_code)
    if m:
        lo, hi = int(m.group(1)), int(m.group(2))
    exp_titles = [h[3] for h in headings
                  if lo <= h[2] <= hi and h[3] not in ('致謝', '作者簡歷')]
    got_titles = [_norm(it[3]) for it in grp_toc]
    if toc_code:
        ok02 = (got_titles == exp_titles
                and len(set(got_titles)) == len(got_titles))
    else:
        # 无 TOC 域时退化为"条目均为真实标题、顺序一致、章与两个目錄齐全"
        it2 = iter(exp_titles)
        ok02 = all(g in it2 for g in got_titles) and bool(got_titles)
        for h in headings:
            if _CHAP_RE.match(h[3]) and h[3] not in got_titles:
                ok02 = False
    ok02 = ok02 and '表目錄' in got_titles and '圖目錄' in got_titles
    _plus(3, ok02,
          '目錄条目与正文标题一一对应且逐字一致，含"表 目 錄""圖 目 錄"两条')

    # 03 +1 目錄中不含"致  謝"条目
    # 与办公软件对齐：比对目錄条目文字（去空白）是否等于"致謝"。
    _plus(1, '致謝' not in got_titles, '目錄中不含"致  謝"条目')

    # 04 +1 目錄中不含"作者簡歷"条目
    _plus(1, '作者簡歷' not in got_titles, '目錄中不含"作者簡歷"条目')

    h_ack = find_heading('致謝')
    h_abs = find_heading('摘要')
    h_eng = find_heading('Abstract')
    h_ref = find_heading('參考文獻')
    s_ack = sec_map.get(id(h_ack[1]._element), -1) if h_ack else -1
    s_abs = sec_map.get(id(h_abs[1]._element), -1) if h_abs else -1

    # 05 +5 "致  謝"末尾为"下一页"分节符，"摘  要"是新分节的首页第一段
    # 严格对齐细则：①致謝与摘要不在同一节，且摘要所在节紧接致謝所在节；
    # ②该分节符类型为"下一页"（w:type 缺省或 nextPage）；③摘要标题是新节的
    # 第一个非空段落。与办公软件对齐：读 w:sectPr/w:type/@w:val。
    ok05 = False
    if s_ack >= 0 and s_abs == s_ack + 1 and s_ack < len(doc.sections):
        if _section_break_type(doc.sections[s_ack]) == 'nextPage':
            firsts = [p for p in paras
                      if sec_map.get(id(p._element)) == s_abs and _norm(p.text)]
            ok05 = bool(firsts) and firsts[0] is h_abs[1]
    _plus(5, ok05, '"致  謝"末尾为"下一页"分节符，"摘  要"为新分节的首页第一段')

    # 06 +3 "致  謝"所在页不显示页码
    # 严格对齐细则：以渲染结果为准——该页页脚区域没有任何数字 / 罗马字符；
    # 渲染不可用时退化读该节生效页脚（沿 w:footerReference 与"链接到前一节"
    # 上溯）里既无 PAGE 域也无数字文字。
    ok06 = False
    f_ack, _ = _effective_part(doc, s_ack, 'footer') if s_ack >= 0 else (None, -1)
    ack_codes = _part_field_codes(f_ack)
    ack_ftxt = _norm(_part_text(f_ack))
    xml_ok06 = (not any('PAGE' in c.upper() for c in ack_codes)
                and not re.search(r'[0-9IVXLCDM]', ack_ftxt))
    if pages and h_ack:
        pg_ack = page_of_para(h_ack[1])
        ok06 = bool(pg_ack) and footer_label(pg_ack) == ''
    elif h_ack:
        ok06 = xml_ok06
    _plus(3, ok06, '"致  謝"所在页不显示页码')

    # 07 +5 摘要至圖目錄的前置部分页码为连续的大写罗马数字，且"摘  要"页为 I
    # 严格对齐细则：①摘要所在页页脚显示 I；②自摘要页起至圖目錄页逐页为
    # I、II、III… 连续无跳号无重复；③编号格式在节属性中登记为 upperRoman，
    # 不是手工敲上去的静态文字。与办公软件对齐：读 w:pgNumType/@w:fmt 与 @w:start。
    p_abs = page_of_heading('摘要')
    p_fig_t = page_of_heading('圖目錄')
    ok07 = False
    if pages and p_abs and p_fig_t and p_abs <= p_fig_t <= len(pages):
        labels = [footer_label(pg) for pg in range(p_abs, p_fig_t + 1)]
        ok07 = (len(labels) <= len(_ROMAN)
                and all(labels[k] == _ROMAN[k] for k in range(len(labels))))
    s_fig = -1
    if pos_fig_toc >= 0:
        s_fig = sec_map.get(id(paras[pos_fig_toc]._element), -1)
    if s_abs >= 0 and s_fig >= s_abs:
        for si in range(s_abs, s_fig + 1):
            if _pg_num_type(doc.sections[si])[0] != 'upperRoman':
                ok07 = False
    else:
        ok07 = False
    _plus(5, ok07, '摘要至圖目錄的前置部分为连续大写罗马数字页码，"摘  要"页为 I')

    # 08 +5 目錄前置五条齐全（摘 要 / Abstract / 目 錄 / 表 目 錄 / 圖 目 錄），
    #        且五条罗马页码全部与实际所在页一致（全部检查）
    # 严格对齐细则：①五条缺一不可；②每条声明的页码须与该标题渲染后所在页的
    # 页脚显示值逐条相等；③大小写与字形按罗马数字大写比对。
    # 与办公软件对齐：条目页码读目錄条目文字尾部（PAGEREF 域的缓存结果），
    # 实际页码读渲染 PDF 该页页脚区域文字。
    FRONT5 = ['摘要', 'Abstract', '目錄', '表目錄', '圖目錄']
    front_all = all(f in got_titles for f in FRONT5)
    mism08 = []
    for it in grp_toc:
        lbl = _norm(it[3])
        if lbl not in FRONT5:
            continue
        declared = _norm(it[4]).upper()
        h = find_heading(lbl)
        actual = page_of_para(h[1]) if h else 0
        shown = footer_label(actual).upper()
        if not declared or not actual or declared != shown:
            mism08.append('%s(%s≠%s)' % (lbl, declared or '空', shown or '空'))
    ok08 = bool(pages) and front_all and not mism08
    _plus(5, ok08, '目錄前置五条齐全且五条罗马页码全部与实际一致')

    # 09 +3 文档中不存在"作者簡歷"标题及其正文
    # 严格对齐细则：全文（含表格）不出现"作者簡歷"字样，也不存在该大纲条目。
    ok09 = ('作者簡歷' not in _norm(body_all_text)) and (find_heading('作者簡歷') is None)
    _plus(3, ok09, '文档中不存在"作者簡歷"标题及其正文')

    # 10 +3 正文结束于參考文獻最后一条，其后无占位空段落、无空白页
    # 严格对齐细则：①最后一个非空段落属于參考文獻部分（在"參考文獻"标题之后）；
    # ②该段落之后不存在任何空段落；③渲染后最后一页有正文内容。
    last_idx = -1
    for i, p in enumerate(paras):
        if _norm(p.text):
            last_idx = i
    trailing_blank = 0
    if last_idx >= 0:
        for p in paras[last_idx + 1:]:
            if not _norm(p.text) and not _is_page_break_only(p):
                trailing_blank += 1
    ref_pos = -1
    if h_ref is not None:
        ref_pos = h_ref[0]
    tail_page_ok = True
    if pages:
        tail_page_ok = bool(pages[-1]['norm'])
    # "參考文獻"必须是文档最后一个大纲标题（其后不得再有致謝 / 作者簡歷之类的章节），
    # 且最后一个非空段落必须是形如 [n] 的文献条目本身。
    ref_is_last = bool(headings) and ref_pos >= 0 and headings[-1][0] == ref_pos
    tail_is_ref_item = False
    if last_idx >= 0:
        tail_is_ref_item = bool(re.match(r'^\s*[\[［]\s*\d+\s*[\]］]',
                                         paras[last_idx].text or ''))
    ok10 = (ref_pos >= 0 and last_idx > ref_pos and ref_is_last
            and tail_is_ref_item and trailing_blank == 0 and tail_page_ok)
    _plus(3, ok10, '正文结束于參考文獻最后一条，其后无占位空段落、无空白页')

    def region_codes(start, end):
        """取 body 区间 (start, end) 内的域代码，用于判定某个目錄是否为自动域。"""
        lo_ = start + 1 if start >= 0 else 0
        hi_ = end if end > lo_ else len(paras)
        out = []
        for p in paras[lo_:hi_]:
            buf = []
            for it in p._element.iter(qn('w:instrText')):
                buf.append(it.text or '')
            for fs in p._element.iter(qn('w:fldSimple')):
                buf.append(fs.get(qn('w:instr')) or '')
            if buf:
                out.append(''.join(buf))
        return ' '.join(out)

    pos_ch1 = -1
    for h in headings:
        if _CHAP_RE.match(h[3]):
            pos_ch1 = h[0]
            break
    code_toc = region_codes(pos_toc, pos_tbl_toc)
    code_tbl = region_codes(pos_tbl_toc, pos_fig_toc)
    code_fig = region_codes(pos_fig_toc, pos_ch1)

    def tnum(s):
        m = _TBL_CAP_RE.match(s)
        return '表%s-%s' % (m.group(1), m.group(2)) if m else ''

    # 11 +5 表目錄为可更新的自动域（选中有域底纹、F9 可整体更新）
    # 严格对齐细则：表目錄区域内存在 TOC 域，且带图表目录开关 \c 与题注标签"表"；
    # 纯静态文字或纯手打条目不给分。
    # 与办公软件对齐：读该区间 w:instrText / w:fldSimple@w:instr 的域代码。
    ok11 = ('TOC' in code_tbl.upper() and '\\c' in code_tbl.lower()
            and '表' in code_tbl and len(grp_tbl) > 0)
    _plus(5, ok11, '表目錄为可更新的自动域（TOC \\c "表"），可整体更新')

    # 12 +5 表目錄完整收录全部 5 个表，无缺漏、无重复、顺序与正文一致
    exp_tbl = [tnum(txt) for (_i, _p, txt) in captions]
    got_tbl = [tnum(_norm(it[3])) for it in grp_tbl]
    ok12 = (len(exp_tbl) == 5 and got_tbl == exp_tbl
            and len(set(got_tbl)) == len(got_tbl))
    _plus(5, ok12, '表目錄完整收录全部 5 个表且顺序与正文一致')

    # 13 +3 表目錄条目文字与正文表题逐字一致
    cap_texts = [txt for (_i, _p, txt) in captions]
    lbl_texts = [_norm(it[3]) for it in grp_tbl]
    ok13 = bool(lbl_texts) and lbl_texts == cap_texts
    _plus(3, ok13, '表目錄条目文字与正文表题逐字一致')

    # 14 +5 表目錄 5 条页码全部与表格实际所在页一致（全部检查）
    # 严格对齐细则：以表题渲染后所在页的页脚显示值为实际页码，与条目声明值逐条相等。
    cap_by_num = {}
    for (_i, p, txt) in captions:
        k = tnum(txt)
        if k:
            cap_by_num.setdefault(k, p)
    mism14 = []
    for it in grp_tbl:
        k = tnum(_norm(it[3]))
        declared = _norm(it[4])
        cp = cap_by_num.get(k)
        actual = page_of_caption(cp) if cp is not None else 0
        shown = footer_label(actual)
        if not k or not declared or not actual or declared != shown:
            mism14.append('%s(%s≠%s)' % (k or '?', declared or '空', shown or '空'))
    ok14 = bool(pages) and len(grp_tbl) == 5 and not mism14
    _plus(5, ok14, '表目錄 5 条页码全部与实际一致')

    # 15 +3 表目錄位于"表 目 錄"标题所在页，且在目錄之后、圖目錄之前
    # 严格对齐细则：①body 顺序为 目錄 → 表目錄 → 圖目錄；②全部表目錄条目与
    # "表 目 錄"标题渲染在同一页。
    pg_ttoc = page_of_para(paras[pos_tbl_toc]) if pos_tbl_toc >= 0 else 0
    ok15 = (pos_toc >= 0 and pos_tbl_toc > pos_toc
            and pos_fig_toc > pos_tbl_toc and bool(grp_tbl))
    if ok15 and pages:
        ok15 = bool(pg_ttoc) and all(page_of_para(it[1]) == pg_ttoc
                                     for it in grp_tbl)
    _plus(3, ok15, '表目錄位于"表 目 錄"标题所在页，且在目錄之后、圖目錄之前')

    # 16 +5 表4-2 的首行与末行跨越页数不超过 3 页
    # 严格对齐细则：以渲染结果为准，末行所在页 − 首行所在页 + 1 ≤ 3。
    i42 = ti('表4-2')
    span42 = 0
    if i42 and i42['first'] and i42['last']:
        span42 = i42['last'] - i42['first'] + 1
    ok16 = bool(pages) and span42 > 0 and span42 <= 3
    _plus(5, ok16, '表4-2 的首行与末行跨越页数不超过 3 页')

    # 17 +5 表4-2 的表头位于新一页的最上方
    # 严格对齐细则：①表头（首行）所在页的最上方一块内容是表4-2 的表题或表头本身；
    # ②该页在表4-2 之前没有其他正文内容；③顶页手段不限（回车 / 空段落 /
    # 分页符 / 段前分页任一均可）。与办公软件对齐：渲染后取该页 y0 最小的文本块。
    ok17 = False
    if pages and i42 and i42['first']:
        pg = i42['first']
        first_blk = pages[pg - 1]['first']
        cap42 = _norm(i42['cap'].text) if i42['cap'] is not None else ''
        cap_tail42 = re.sub(r'^表\d+[-－]\d+', '', cap42)
        head_keys42 = row_keys(i42['t'], 0)
        ok17 = bool(first_blk) and (
            (cap42 and first_blk.startswith(cap42[:12]))
            # SEQ 域更新后表号可能变化，只要稳定的表题正文位于页首，同样满足。
            or (cap_tail42 and cap_tail42 in first_blk[:max(16, len(cap_tail42) + 8)])
            # PDF 文本层可能把同一表头的相邻单元格交错拼接；首个单元格的前三个
            # 字符足以确认页首从表头开始，同时避免依赖整行拼接顺序。
            or any(len(key) >= 3 and first_blk.startswith(key[:3])
                   for key in head_keys42))
    _plus(5, ok17, '表4-2 的表头位于新一页的最上方')

    # ======================= 扣分项（32 条 / −106） =======================

    head_positions = sorted(h[0] for h in headings)

    def block_after(pos, keep_empty=False):
        """取某标题之后、下一个标题之前的段落（默认只取非空段）。"""
        if pos < 0:
            return []
        nxt = len(paras)
        for hp in head_positions:
            if hp > pos:
                nxt = hp
                break
        out = paras[pos + 1:nxt]
        return out if keep_empty else [p for p in out if _norm(p.text)]

    def paras_on_page(pg):
        if not pages or not pg:
            return []
        return [p for p in paras
                if _norm(p.text) and ppage.get(id(p._element)) == pg]

    # 18 −5 封面信息不完整或有误
    # 严格对齐细则（绝对标准，逐项核对）：姓名 林澄韵；學號 M260742018555；
    # 院系 國際商務與數位經濟學院；課程 企業管理；指導教師 許承嶧；電郵一栏留空；
    # 中文题名"董事會數位素養對企業低碳轉型績效的影響"与对应英文题名并存；
    # 文种标识"碩士研究生學位論文"；日期 2026 年 5 月 18 日。任一项缺失或有误即扣。
    cover_paras = [p for p in paras if sec_map.get(id(p._element), 0) == 0]
    cover_raw = ''.join(p.text for p in cover_paras)
    cover_txt = _norm(cover_raw)
    cover_need = ['林澄韵', 'M260742018555', '國際商務與數位經濟學院', '企業管理',
                  '許承嶧', '董事會數位素養對企業低碳轉型績效的影響',
                  '碩士研究生學位論文', '2026年5月18日']
    missing18 = [s for s in cover_need if s not in cover_txt]
    if not re.search(r'[A-Za-z][A-Za-z ,\-]{19,}', cover_raw):
        missing18.append('英文题名')
    if '@' in cover_txt:
        missing18.append('電郵一栏未留空')
    _minus(5, bool(missing18), '封面信息不完整或有误')

    # 19 −3 答辯委員會審議通過表页缺失或不完整
    # 严格对齐细则（绝对标准）：①文档中存在"答辯委員會審議通過表"标题；②它排在
    # 封面之后的独立一页上，该页除这行标题外没有其他正文内容。
    # 与办公软件对齐：审议通过表是答辯后手写签署的表格页，正文层面只有一行标题，
    # 页内其余是留白。判"完整"只能判"标题在、且独占一页"，页内文字段数不作要求。
    has19 = '答辯委員會審議通過表' in _norm(body_all_text)
    own19 = True
    page19 = -1
    if pages:
        for i_, pg_ in enumerate(pages):
            if '答辯委員會審議通過表' in pg_['norm']:
                page19 = i_
                break
        if page19 < 0:
            own19 = False
        else:
            rest19 = pages[page19]['norm'].replace('答辯委員會審議通過表', '')
            own19 = (page19 >= 1) and (not rest19)
    bad19 = (not has19) or (not own19)
    _minus(3, bad19, '答辯委員會審議通過表页缺失或不完整')

    # 20 −3 致謝正文不足 4 段，或缺署名 / 日期 / "於南海都會大學"
    # 严格对齐细则：①"致  謝"标题之后至下一标题之前，除署名行与日期行外的非空
    # 段落不少于 4 段；②署名为"林澄韵"；③含"於南海都會大學"；④含"YYYY 年 M 月"。
    ack_block = block_after(h_ack[0]) if h_ack else []
    ack_join = _norm(''.join(p.text for p in ack_block))
    ack_body = [p for p in ack_block
                if not re.search(r'林澄韵|於南海都會大學|\d{4}\s*年', p.text)]
    bad20 = (not h_ack) or (len(ack_body) < 4) or ('林澄韵' not in ack_join) \
        or ('於南海都會大學' not in ack_join) \
        or (not re.search(r'\d{4}年\d{1,2}月', ack_join))
    _minus(3, bad20, '致謝正文不足 4 段，或缺署名 / 日期 / "於南海都會大學"')

    # 21 −3 摘要或 Abstract 缺正文，或缺"關鍵字"/"Keywords"行
    # 严格对齐细则：摘要块与 Abstract 块各自的非空段落不少于 2 段，且分别含
    # "關鍵字"与"Keywords"（大小写不敏感）。
    abs_block = block_after(h_abs[0]) if h_abs else []
    eng_block = block_after(h_eng[0]) if h_eng else []
    abs_join = ''.join(p.text for p in abs_block)
    eng_join = ''.join(p.text for p in eng_block)
    bad21 = (len(abs_block) < 2) or (len(eng_block) < 2) \
        or ('關鍵字' not in _norm(abs_join)) \
        or ('keywords' not in eng_join.lower())
    _minus(3, bad21, '摘要或 Abstract 缺正文，或缺"關鍵字"/"Keywords"行')

    # 22 −5 章标题不足 5 个 / 节标题不足 17 个 / 大纲层级不正确
    # 严格对齐细则：①"第X章"标题恰为 5 个且大纲级别为 1；②"N.N"节标题不少于
    # 17 个且大纲级别为 2。与办公软件对齐：读 w:outlineLvl。
    chaps = [h for h in headings if _CHAP_RE.match(h[3])]
    secs = [h for h in headings if re.match(r'^\d+\.\d+', h[3])]
    bad22 = (len(chaps) < 5) or (len(secs) < 17) \
        or any(h[2] != 1 for h in chaps) or any(h[2] != 2 for h in secs)
    _minus(5, bad22, '章标题不足 5 个 / 节标题不足 17 个 / 大纲层级不正确')

    # 23 −5 正文存在段落缺损痕迹（抽检 1.1 / 2.2 / 4.3 / 5.1 各自首段）
    # 严格对齐细则：抽检段落须成句——字数不少于 15，且结尾为句末标点
    # （。！？」』）），以逗号 / 顿号 / 分号 / 冒号收尾或字数过少视为缺损。
    def first_body_para(prefix):
        for h in headings:
            if h[3].startswith(prefix):
                for p in block_after(h[0]):
                    t = _norm(p.text)
                    if t and not _TBL_CAP_RE.match(t) and not _FIG_CAP_RE.match(t):
                        return p
        return None

    broken23 = []
    for pref in ('1.1', '2.2', '4.3', '5.1'):
        p = first_body_para(pref)
        t = _norm(p.text) if p is not None else ''
        if (not t) or len(t) < 15 or t[-1] in '，、；：,;:':
            broken23.append(pref)
    _minus(5, bool(broken23),
           '正文存在段落缺损痕迹（抽检 1.1 / 2.2 / 4.3 / 5.1 首段）')

    # 24 −3 參考文獻少于 8 条，或编号不连续
    # 严格对齐细则：條目以"[N]"起头，条数不少于 8，且 N 从 1 起逐一递增无跳号。
    ref_block = block_after(h_ref[0]) if h_ref else []
    ref_nums = []
    for p in ref_block:
        m = re.match(r'^\s*\[(\d+)\]', p.text)
        if m:
            ref_nums.append(int(m.group(1)))
    bad24 = (len(ref_nums) < 8) or (ref_nums != list(range(1, len(ref_nums) + 1)))
    _minus(3, bad24, '參考文獻少于 8 条，或编号不连续')

    # 25 −5 表格数据有缺损
    # 严格对齐细则：抽检表3-1 全部 12 行、表4-2 第 1 / 32 / 64 行、表4-1 表头行。
    # 数据行出现空单元格即判缺损；表头行允许 1 个约定留空的角单元格，空 2 个即判缺损。
    def row_defect(t, r):
        try:
            cells = t.rows[r].cells
        except Exception:
            return True
        empties = sum(1 for c in cells if not _norm(c.text))
        return (empties >= 2) if r == 0 else (empties >= 1)

    bad25 = False
    for num, rows_idx in (('表3-1', list(range(12))),
                          ('表4-2', [0, 31, 63]),
                          ('表4-1', [0])):
        info = ti(num)
        if info is None:
            bad25 = True
            continue
        for r in rows_idx:
            if row_defect(info['t'], r):
                bad25 = True
    _minus(5, bad25,
           '表格数据有缺损（抽检表3-1 全 12 行、表4-2 第 1/32/64 行、'
           '表4-1 表头行）')

    # 26 −3 表格行列数不符
    # 严格对齐细则（绝对标准）：表2-1 为 5 行 4 列、表3-1 为 12 行 4 列、
    # 表4-1 为 11 行 6 列、表4-2 为 64 行 7 列、表4-3 为 5 行 5 列。
    EXPECT_RC = {'表2-1': (5, 4), '表3-1': (12, 4), '表4-1': (11, 6),
                 '表4-2': (64, 7), '表4-3': (5, 5)}

    def rc_of(t):
        try:
            nr = len(t.rows)
        except Exception:
            nr = 0
        grid = t._tbl.find(qn('w:tblGrid'))
        nc = len(grid.findall(qn('w:gridCol'))) if grid is not None else 0
        return nr, nc

    wrong26 = []
    for num, exp in EXPECT_RC.items():
        info = ti(num)
        if info is None:
            wrong26.append('%s缺失' % num)
            continue
        got = rc_of(info['t'])
        if got != exp:
            wrong26.append('%s(%d×%d)' % (num, got[0], got[1]))
    _minus(3, bool(wrong26), '表格行列数不符')

    # 27 −3 5 个表格的格式不统一
    # 严格对齐细则：五个表格的"表格样式 + 表格对齐 + 表格边框组合 + 表内主体字号"
    # 四项特征须完全一致，任一项出现分化即判不统一。
    def tbl_sig(t):
        el = t._tbl
        tblPr = el.find(qn('w:tblPr'))
        sty = jc = bsig = ''
        if tblPr is not None:
            s = tblPr.find(qn('w:tblStyle'))
            sty = s.get(qn('w:val')) if s is not None else ''
            j = tblPr.find(qn('w:jc'))
            jc = j.get(qn('w:val')) if j is not None else ''
            bd = tblPr.find(qn('w:tblBorders'))
            if bd is not None:
                bsig = ','.join(sorted(
                    '%s:%s' % (re.sub(r'^\{[^}]*\}', '', c.tag),
                               c.get(qn('w:val')) or '')
                    for c in bd))
        sizes = {}
        try:
            rows = t.rows
        except Exception:
            rows = []
        for row in rows:
            for c in row.cells:
                for p in c.paragraphs:
                    for r in p.runs:
                        sz = _resolve_effective_size(doc, p, r)
                        if sz:
                            sizes[sz] = sizes.get(sz, 0) + 1
        dom = max(sizes, key=lambda k: sizes[k]) if sizes else None
        return (sty, jc, bsig, dom)

    sigs27 = [tbl_sig(ti(n)['t']) for n in EXPECT_RC if ti(n) is not None]
    bad27 = (len(sigs27) != 5) or (len(set(sigs27)) > 1)
    _minus(3, bad27, '5 个表格的格式不统一')

    # 28 −3 图片不足 3 张，或图题缺失 / 未居中
    # 严格对齐细则：①正文图片不少于 3 张（内嵌 wp:inline 与浮动 wp:anchor 均计）；
    # ②"圖N-N ..."图题不少于 3 条；③每条图题段落对齐方式为居中。
    n_inline = len(doc.element.body.findall('.//' + qn('wp:inline')))
    n_anchor = len(doc.element.body.findall('.//' + qn('wp:anchor')))
    n_pic = n_inline + n_anchor
    notcenter28 = [t for (_i, p, t) in figcaps if _para_jc(doc, p) != 'center']
    bad28 = (n_pic < 3) or (len(figcaps) < 3) or bool(notcenter28)
    _minus(3, bad28, '图片不足 3 张，或图题缺失 / 未居中')

    # 29 −5 目錄不是由 TOC 域自动生成（选中无域底纹、F9 无反应）
    # 严格对齐细则：目錄区域内必须存在 TOC 域，整体为静态文字即扣。
    # 与办公软件对齐：只判"目錄"本体。圖目錄用不用域属于另一条评分项的范围，
    # 混在这里会让同一处缺陷被扣两次。
    bad29 = 'TOC' not in code_toc.upper()
    _minus(5, bad29, '目錄不是由 TOC 域自动生成')

    # 30 −5 目錄中正文条目的阿拉伯页码与实际不符（抽检 第四章 / 4.3 / 參考文獻）
    def toc_entry_of(prefix):
        for it in grp_toc:
            if _norm(it[3]).startswith(prefix):
                return it
        return None

    def heading_of(prefix):
        for h in headings:
            if h[3].startswith(prefix):
                return h
        return None

    spot30 = []
    for pref in ('第四章', '4.3', '參考文獻'):
        ent = toc_entry_of(pref)
        h = heading_of(pref)
        declared = _norm(ent[4]) if ent else ''
        actual = page_of_para(h[1]) if h else 0
        shown = footer_label(actual)
        if (ent is None) or (h is None) or (not declared) or (not actual) \
                or declared != shown:
            spot30.append('%s(%s≠%s)' % (pref, declared or '空', shown or '空'))
    _minus(5, bool(pages) and bool(spot30),
           '目錄中正文条目的阿拉伯页码与实际不符（抽检 第四章 / 4.3 / 參考文獻）')

    # 31 −3 正文格式不符（抽检 1.1 首段与 4.4 末段）
    # 严格对齐细则（绝对标准）：字号四号（14 磅）、行距 1.5 倍（w:line=360 且
    # lineRule=auto）、首行缩进 2 字符。
    # 与办公软件对齐：Word / WPS 里"缩进 2 字符"写成 w:firstLineChars="200"，
    # 此时同时存在的 w:firstLine 只是给不支持字符单位的阅读器的备份值（本文档为
    # 420 = 2×10.5 磅），不能拿它跟正文字号硬算。只有在没有 firstLineChars 时
    # 才按"2 × 实际字号"折算成缇来比。
    def last_body_para(prefix):
        for h in headings:
            if h[3].startswith(prefix):
                blk = [p for p in block_after(h[0])
                       if not _TBL_CAP_RE.match(_norm(p.text))
                       and not _FIG_CAP_RE.match(_norm(p.text))]
                return blk[-1] if blk else None
        return None

    def body_fmt_ok(p):
        if p is None:
            return False
        szs = [s for s in (_resolve_effective_size(doc, p, r) for r in p.runs) if s]
        if not szs or any(abs(s - 14.0) > 0.01 for s in szs):
            return False
        line, lrule = _para_line_spacing(doc, p)
        if not (line and abs(line - 360.0) < 1.0 and lrule in (None, 'auto')):
            return False
        ind = _para_ind(doc, p)
        flc = ind.get('firstLineChars')
        if flc is not None:
            return abs(flc - 200.0) < 0.01
        fl = ind.get('firstLine')
        return bool(fl and abs(fl - 2.0 * szs[0] * 20.0) <= 20.0)

    smp31 = [('1.1 首段', first_body_para('1.1')),
             ('4.4 末段', last_body_para('4.4'))]
    bad31_list = [name for name, p in smp31 if not body_fmt_ok(p)]
    _minus(3, bool(bad31_list),
           '正文格式不符（四号 / 1.5 倍行距 / 首行缩进 2 字符）')

    # 32 −3 页面设置不符
    # 严格对齐细则（绝对标准，全部分节均须满足）：A4 纵向（21.0×29.7cm）、
    # 上边距 2.54cm、下边距 2.54cm、右边距 2.54cm、左边距 3.17cm、
    # 页眉距边界 1.4cm、页脚距边界 1.27cm。容差 0.05cm。
    def near(v, tgt, tol=0.05):
        try:
            return abs(float(v.cm) - tgt) <= tol
        except Exception:
            return False

    bad32 = False
    for sec in doc.sections:
        if not (near(sec.page_width, 21.0, 0.1)
                and near(sec.page_height, 29.7, 0.1)):
            bad32 = True
        for val, tgt in ((sec.top_margin, 2.54), (sec.bottom_margin, 2.54),
                         (sec.right_margin, 2.54), (sec.left_margin, 3.17),
                         (sec.header_distance, 1.4),
                         (sec.footer_distance, 1.27)):
            if not near(val, tgt):
                bad32 = True
    _minus(3, bad32, '页面设置不符（A4 纵向 / 上下右 2.54cm / 左 3.17cm / '
                     '页眉 1.4cm / 页脚 1.27cm）')

    # 33 −3 正文页眉缺失，或页眉文字不是"南海都會大學研究生學位論文"与论文题名的组合
    # 与办公软件对齐：读该节生效页眉（沿"链接到前一节"上溯）的文字。
    s_body = sec_map.get(id(chaps[0][1]._element), -1) if chaps else -1
    bad33 = (s_body < 0)
    for si in range(max(0, s_body), len(doc.sections)):
        if s_body < 0:
            break
        part, _ = _effective_part(doc, si, 'header')
        htxt = _norm(_part_text(part))
        if ('南海都會大學研究生學位論文' not in htxt) or ('董事會數位素養' not in htxt):
            bad33 = True
    _minus(3, bad33, '正文页眉缺失，或页眉文字不是"南海都會大學研究生學位論文"'
                     '与论文题名的组合')

    # 34 −3 页脚页码不是小五号（9 磅），或未居中
    bad34 = False
    checked34 = 0
    for si in range(len(doc.sections)):
        part, _ = _effective_part(doc, si, 'footer')
        if part is None:
            continue
        if not any('PAGE' in c.upper() for c in _part_field_codes(part)):
            continue
        checked34 += 1
        for p in part.paragraphs:
            has_fld = any(True for _ in p._element.iter(qn('w:instrText')))
            if not (p.text.strip() or has_fld):
                continue
            if _para_jc(doc, p) != 'center':
                bad34 = True
            for r in p.runs:
                sz = _resolve_effective_size(doc, p, r)
                if sz and abs(sz - 9.0) > 0.01:
                    bad34 = True
    if checked34 == 0:
        bad34 = True
    _minus(3, bad34, '页脚页码不是小五号（9 磅），或未居中')

    # 35 −3 正文页脚页码不是自动 PAGE 域，或编号格式不是从 1 起的阿拉伯数字
    # 与办公软件对齐：页脚域代码读 w:instrText；编号格式读 w:pgNumType/@w:fmt
    # 与 @w:start（decimal 且 start=1）。
    bad35 = (s_body < 0)
    if s_body >= 0:
        part, _ = _effective_part(doc, s_body, 'footer')
        codes = _part_field_codes(part)
        if not any(re.search(r'\bPAGE\b', c.upper()) for c in codes):
            bad35 = True
        fmt35, start35 = _pg_num_type(doc.sections[s_body])
        if (fmt35 not in (None, 'decimal')) or (start35 != 1):
            bad35 = True
    _minus(3, bad35, '正文页脚页码不是自动 PAGE 域，或编号格式不是从 1 起的阿拉伯数字')

    # 36 −3 封面页或答辯委員會審議通過表页显示了页码
    # 与办公软件对齐：读渲染结果中这两页页脚区域是否有任何文字。
    bad36 = False
    if pages:
        if footer_label(1):
            bad36 = True
        if page19 >= 0 and footer_label(page19 + 1):
            bad36 = True
    _minus(3, bad36, '封面页或答辯委員會審議通過表页显示了页码')

    # 37 −5 存在整页空白页
    # 与办公软件对齐：渲染后某页版心内既无文字块也无图片。
    blank_pages = []
    if pages:
        for i, pg in enumerate(pages):
            if not pg['norm'] and not pg['images']:
                blank_pages.append(i + 1)
    _minus(5, bool(blank_pages), '存在整页空白页')

    # 38 −3 表格所在页出现大片留白
    # 严格对齐细则：表格所在页版心内容底部不足 385 磅（版心底约 769.9 磅的一半），
    # 且该页除表题外没有任何正文段落。
    sparse38 = []
    if pages:
        tbl_pages = set()
        for info in tbl_info.values():
            a, b = info['first'], info['last']
            if a and b:
                tbl_pages.update(range(a, b + 1))
            elif info['cap_pg']:
                tbl_pages.add(info['cap_pg'])
        for pg in sorted(tbl_pages):
            if pages[pg - 1]['bottom'] >= 385.0:
                continue
            others = [p for p in paras_on_page(pg)
                      if not _TBL_CAP_RE.match(_norm(p.text))
                      and not _FIG_CAP_RE.match(_norm(p.text))]
            if not others:
                sparse38.append(pg)
    _minus(3, bool(sparse38), '表格所在页出现大片留白')

    def crosses(num):
        info = ti(num)
        if info is None or not info['first'] or not info['last']:
            return False
        return info['first'] != info['last']

    # 39 −3 表2-1 或表4-3 的首行与末行不在同一页
    _minus(3, bool(pages) and (crosses('表2-1') or crosses('表4-3')),
           '表2-1 或表4-3 的首行与末行不在同一页')

    # 40 −3 表3-1 的首行与末行不在同一页
    _minus(3, bool(pages) and crosses('表3-1'),
           '表3-1 的首行与末行不在同一页')

    # 41 −3 表4-1 的首行与末行不在同一页
    _minus(3, bool(pages) and crosses('表4-1'),
           '表4-1 的首行与末行不在同一页')

    # 42 −3 表4-2 未开启"跨页时重复标题行"
    # 与办公软件对齐：读首行 w:trPr/w:tblHeader。
    i42b = ti('表4-2')
    ok42 = False
    if i42b is not None:
        try:
            trPr = i42b['t'].rows[0]._tr.find(qn('w:trPr'))
        except Exception:
            trPr = None
        if trPr is not None:
            th = trPr.find(qn('w:tblHeader'))
            if th is not None:
                ok42 = (th.get(qn('w:val')) in (None, '1', 'true', 'on'))
    _minus(3, not ok42, '表4-2 未开启"跨页时重复标题行"')

    # 43 −3 表題与表格不在同一页（抽检表3-1、表4-1、表4-2、表4-3），
    #        或表4-2 的"註："与表格分离
    sep43 = []
    for num in ('表3-1', '表4-1', '表4-2', '表4-3'):
        info = ti(num)
        if info is None:
            continue
        if info['cap_pg'] and info['first'] and info['cap_pg'] != info['first']:
            sep43.append(num)
    note_para = None
    if i42b is not None:
        for k, (kind, obj) in enumerate(seq):
            if kind == 'tbl' and obj is i42b['t']:
                for j in range(k + 1, min(len(seq), k + 6)):
                    if seq[j][0] == 'p' and _norm(seq[j][1].text).startswith('註'):
                        note_para = seq[j][1]
                        break
                break
    if note_para is not None and i42b is not None and i42b['last']:
        npg = page_of_para(note_para)
        if npg and npg != i42b['last']:
            sep43.append('表4-2 的"註："')
    _minus(3, bool(pages) and bool(sep43),
           '表題与表格不在同一页，或表4-2 的"註："与表格分离')

    # 44 −3 任一表格的列宽合计超出版心宽度 15.28cm
    # 与办公软件对齐：读 w:tblGrid/w:gridCol/@w:w（缇），换算为厘米后求和。
    def grid_cm(t):
        grid = t._tbl.find(qn('w:tblGrid'))
        if grid is None:
            return None
        tot = 0.0
        for gc in grid.findall(qn('w:gridCol')):
            w = gc.get(qn('w:w'))
            if w is None:
                return None
            tot += float(w)
        return tot / 1440.0 * 2.54

    over44 = []
    for num in EXPECT_RC:
        info = ti(num)
        if info is None:
            continue
        cmv = grid_cm(info['t'])
        if cmv is not None and cmv > 15.28 + 0.05:
            over44.append('%s(%.2fcm)' % (num, cmv))
    _minus(3, bool(over44), '任一表格的列宽合计超出版心宽度 15.28cm')

    # 45 −3 目錄二级条目未设置"文本之前缩进 2 字符"/"首行缩进 2 字符"
    # 与办公软件对齐：读 w:ind/@w:leftChars（=200）与 @w:firstLineChars（=200），
    # 缺字符单位时按"2 × 该条目实际字号"折算成缇。细则用"/"表示二者取其一，
    # Word 的自动目錄默认写 leftChars，手工排版的往往只写 firstLine，两种都算合规。
    lv2 = [it for it in grp_toc
           if it[2] == 2 or re.match(r'^\d+\.\d+', _norm(it[3]))]
    bad45_n = 0
    for it in lv2:
        p45 = it[1]
        szs45 = [s for s in (_resolve_effective_size(doc, p45, r) for r in p45.runs) if s]
        exp45 = 2.0 * (szs45[0] if szs45 else 10.5) * 20.0
        ind = _para_ind(doc, p45)
        left_ok = (ind.get('leftChars') == 200.0
                   or ind.get('startChars') == 200.0
                   or (ind.get('left') and abs(ind['left'] - exp45) <= 20.0)
                   or (ind.get('start') and abs(ind['start'] - exp45) <= 20.0))
        fl_ok = (ind.get('firstLineChars') == 200.0
                 or (ind.get('firstLine') and abs(ind['firstLine'] - exp45) <= 20.0))
        if not (left_ok or fl_ok):
            bad45_n += 1
    _minus(3, bad45_n > 0,
           '目錄二级条目未设置"文本之前缩进 2 字符"/"首行缩进 2 字符"')

    # 46 −3 圖目錄页码与实际不符，或缺标题 / 条目不足 3 条 / 未独立成页
    # 严格对齐细则：①存在"圖 目 錄"标题；②条目不少于 3 条；③每条声明页码与
    # 对应图题渲染后所在页的页脚显示值相等；④该页除标题与条目外无其他正文段落。
    fig_by_num = {}
    for (_i, p, txt) in figcaps:
        m = _FIG_CAP_RE.match(txt)
        if m:
            fig_by_num.setdefault('圖%s-%s' % (m.group(1), m.group(2)), p)
    mism46 = []
    for it in grp_fig:
        m = _FIG_CAP_RE.match(_norm(it[3]))
        k = '圖%s-%s' % (m.group(1), m.group(2)) if m else ''
        fp = fig_by_num.get(k)
        actual = page_of_caption(fp) if fp is not None else 0
        declared = _norm(it[4])
        if not k or not declared or not actual or declared != footer_label(actual):
            mism46.append('%s(%s≠%s)' % (k or '?', declared or '空',
                                         footer_label(actual) or '空'))
    pg_ftoc = page_of_para(paras[pos_fig_toc]) if pos_fig_toc >= 0 else 0
    alone46 = True
    if pages and pg_ftoc:
        extra = [p for p in paras_on_page(pg_ftoc)
                 if not is_toc_entry(p) and p is not paras[pos_fig_toc]]
        alone46 = not extra
    bad46 = (pos_fig_toc < 0) or (len(grp_fig) < 3) or (not alone46) \
        or (bool(pages) and bool(mism46))
    _minus(3, bad46,
           '圖目錄页码与实际不符，或缺标题 / 条目不足 3 条 / 未独立成页')

    # 47 −3 出现指令未要求的新增内容
    # 严格对齐细则：①大纲标题只允许"摘 要 / Abstract / 目 錄 / 表 目 錄 /
    # 圖 目 錄 / 致 謝 / 參考文獻 / 答辯委員會審議通過表 / 第X章 / N.N"这些形态；
    # ②目錄条目不得添加"一、二、"之类指令未要求的序号前缀。
    extra47 = []
    for h in headings:
        t = h[3]
        if t in _FRONT_TITLES or _CHAP_RE.match(t) or re.match(r'^\d+(\.\d+)+', t):
            continue
        if '答辯委員會' in t:
            continue
        extra47.append(t)
    for it in grp_toc:
        if re.match(r'^[一二三四五六七八九十]+、', _norm(it[3])):
            extra47.append('目錄条目附加序号前缀')
            break
    _minus(3, bool(extra47), '出现指令未要求的新增内容')

    # 48 −1 出现简体字，或中英文标点混用
    # 严格对齐细则：①正文（含表格）不得出现简化字；②中文字符后不得紧跟半角
    # 逗号 / 分号 / 问号 / 叹号。
    tbl_text = []
    for t in doc.tables:
        for row in t.rows:
            for c in row.cells:
                tbl_text.append(c.text)
    all_text = body_all_text + '\n' + '\n'.join(tbl_text)
    simp_hit = sorted(set(ch for ch in all_text if ch in _SIMPLIFIED))
    mix_hit = re.findall(r'[一-鿿][,;?!]', all_text)
    bad48 = bool(simp_hit) or bool(mix_hit)
    _minus(1, bad48, '出现简体字，或中英文标点混用')

    # 49 −1 文档总页数不在 26—29 页区间
    _minus(1, bool(pages) and not (26 <= n_pages <= 29),
           '文档总页数不在 26—29 页区间')

    # ---------- 反解为结构化 dim2_items ----------
    if len(results) != len(_MAX_DELTAS_ORDER):
        raise RuntimeError('dim2 项数不一致：results=%d vs order=%d'
                           % (len(results), len(_MAX_DELTAS_ORDER)))
    dim2_items = []
    total_score = 0
    for raw, max_delta in zip(results, _MAX_DELTAS_ORDER):
        head, _, rule = raw.partition('：')
        h = head.strip().replace('＋', '+').replace('−', '-') \
                        .replace('–', '-').replace('—', '-')
        try:
            delta = int(h)
        except ValueError:
            delta = 0
        hit = delta != 0
        dim2_items.append({"rule": rule, "max_delta": max_delta,
                           "delta": max_delta if hit else 0,
                           "hit": hit})
        total_score += (max_delta if hit else 0)
    result["dim2_items"] = dim2_items
    result["total_score"] = max(0, total_score)
    result["max_score"] = MAX_SCORE_TOTAL
    return result


if __name__ == "__main__":
    _dir = sys.argv[1] if len(sys.argv) >= 2 else os.path.dirname(
        os.path.abspath(__file__))
    _out = evaluate(_dir)
    print(json.dumps(_out, ensure_ascii=False, indent=2))

