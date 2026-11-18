#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# pyright: reportMissingImports=false, reportMissingModuleSource=false
"""021 「开题报告按模板改版并删除两处技术路线图」任务 自动评测脚本。

维度1：可用与可修改性（不通过则维度2 记 0 分）
维度2：完成度评分细则（21 项加分 = +69，9 项减分 = -17）

入口：evaluate(dir_path) -> dict
"""
import os
import re
import sys
import json
import shutil
import tempfile

from docx import Document
from docx.oxml.ns import qn
from docx.table import Table
from docx.text.paragraph import Paragraph

try:
    _VERIFIER_DIR = os.path.dirname(os.path.abspath(__file__))
    if _VERIFIER_DIR not in sys.path:
        sys.path.insert(0, _VERIFIER_DIR)
    import pymupdf_compat as fitz  # 已移除 PyMuPDF(AGPL)，改用 pdfplumber+pypdfium2 兼容层
except ImportError:
    fitz = None

SCRIPT_ID = "021"
MAX_SCORE_TOTAL = 69
MIN_SCORE_TOTAL = -17

# 与 _evaluate_impl 中 results.append() 的顺序严格一一对应
_MAX_DELTAS_ORDER = [
    5,    # 01 +5 页面设置为 A4 纵向、四边页边距 2.2/2.0/2.5/2.3cm
    3,    # 02 +3 封面顶部居中校徽图片，尺寸达标且等比未变形
    3,    # 03 +3 封面校名 楷体_GB2312 二号、主标题 黑体 一号，各占一行居中
    3,    # 04 +3 封面信息栏标签加粗三号、填写内容加下划线三号、左端对齐、行距段前
    3,    # 05 +3 封面下部单独一行居中「年    月    日」宋体小四
    5,    # 06 +5 正文装入 4 个表格框，外框上下双细线、左右单线
    3,    # 07 +3 共 3 处居中报告名称行，小二加粗
    5,    # 08 +5 正文汉字宋体、西文 Times New Roman，字号统一小四
    5,    # 09 +5 正文段落行距为固定值 22 磅
    3,    # 10 +3 正文段落首行缩进 2 字符
    3,    # 11 +3 五处章标题加粗且汉字为宋体
    1,    # 12 +1 五处章标题序号为手动文本，无残留自动编号
    3,    # 13 +3 工作进度表 3 列结构、表头与 4 行阶段数据一一对应
    1,    # 14 +1 工作进度表单元格小四居中且无制表符
    1,    # 15 +1 参考文献 4 条顶格、无首行缩进与自动编号
    3,    # 16 +3 文末 3 行意见签署表，含签名横线与日期行
    3,    # 17 +3 三个意见栏除栏目名、签名横线、日期行外保持空白
    5,    # 18 +5 全文无竖向排列的技术路线流程图
    5,    # 19 +5 全文无横向排列的技术路线流程图
    3,    # 20 +3 两处流程图引导文字同步删除，无孤立标签
    3,    # 21 +3 保留「图1 播撒机构示意结构」及图题，均居中且图题在下
    -1,   # 22 -1 封面主标题或四、五章标题措辞被改写
    -3,   # 23 -3 意见签署栏出现非遗、新媒体等与本题无关的评语
    -3,   # 24 -3 「图1」图片被压缩/拉伸，宽高比偏离约 1.06:1
    -1,   # 25 -1 图题「图1 播撒机构示意结构」不是黑体五号
    -3,   # 26 -3 出现仅含报告名称行、其余全空的整页
    -1,   # 27 -1 首个表格框被改为浮动/绝对定位表格或竖向框线被取消
    -1,   # 28 -1 删除流程图后遗留的空白段落占一页版面 50% 以上
    -1,   # 29 -1 参考文献条目标点被改为全角逗号、冒号
    -3,   # 30 -3 工作进度表表头行与 4 行阶段数据被分隔到相邻两页
]
assert sum(d for d in _MAX_DELTAS_ORDER if d > 0) == MAX_SCORE_TOTAL
assert sum(d for d in _MAX_DELTAS_ORDER if d < 0) == MIN_SCORE_TOTAL

EMU_PER_CM = 360000.0
TWIPS_PER_CM = 566.9291338582677
PT_PER_CM = 28.346456692913385

# 页面与页边距目标值（厘米），误差 <= 0.1cm
PAGE_W_CM = (20.9, 21.1)
PAGE_H_CM = (29.6, 29.8)
MARGIN_TARGET_CM = {"top": 2.2, "bottom": 2.0, "left": 2.5, "right": 2.3}
MARGIN_TOL_CM = 0.1

# 校徽与「图1」的版面特征
BADGE_W_CM = (14.5, 15.1)
BADGE_H_CM = (4.0, 4.4)
FIG1_RATIO = 1.063
FIG1_RATIO_TOL = 0.05
FLOW_V_RATIO_MAX = 0.60          # 竖向流程图：宽高比 <= 0.60
FLOW_H_RATIO_MIN = 4.50          # 横向流程图：宽高比 >= 4.50（校徽约 3.54，不会误判）

CHAPTER_TITLES = ("研究目的及意义", "行业研究现状", "研究内容与技术方案")
CHAPTER_RE = re.compile(r"^[一二三四五][．.、,，]")
NAME_ROW_RE = re.compile(r"^毕业设计[（(]?论?文?[)）]?开题报告$")
REF_HEAD_RE = re.compile(r"^\[\s*([1-4])\s*\]")
FIG1_CAPTION_KEY = "播撒机构示意结构"
SCHEDULE_HEADER = ("阶段", "主要任务", "时间安排")
SCHEDULE_ROWS = (
    ("前期调研", "2026.01", "2026.02"),
    ("结构设计", "2026.03", "2026.04"),
    ("仿真验证", "2026.04", "2026.05"),
    ("成果整理", "2026.05", "2026.06"),
)
OPINION_TITLES = ("指导教师意见", "教研室审查意见", "所在二级学院审查意见")
IRRELEVANT_WORDS = ("非遗", "非物质文化遗产", "新媒体", "文化传承")
FLOW_LEAD_WORDS = ("技术路线图", "研究技术路线如下")
FLOW_V_NODE_WORDS = ("现状调研剖析", "明确优化设计目标", "形成完整设计方案")


def _squeeze(text):
    """去掉全部空白字符，便于稳定比对。"""
    return re.sub(r"\s+", "", text or "")


def _has_cjk(text):
    for ch in text or "":
        if "一" <= ch <= "鿿":
            return True
    return False


def _has_latin(text):
    return re.search(r"[A-Za-z0-9]", text or "") is not None


def _norm_font(name):
    return (name or "").strip().lower()


def _is_songti(name):
    n = _norm_font(name)
    return ("宋体" in n) or n.startswith("simsun") or n.startswith("nsimsun") or ("songti" in n)


def _is_heiti(name):
    n = _norm_font(name)
    return ("黑体" in n) or n.startswith("simhei") or ("heiti" in n)


def _is_kaiti_gb(name):
    n = _norm_font(name)
    return ("楷体" in n) or ("kaiti" in n)


def _is_tnr(name):
    n = _norm_font(name)
    return n.replace(" ", "") == "timesnewroman"


def _pPr(para):
    return para._p.find(qn("w:pPr"))


def _para_mark_rPr(para):
    pPr = _pPr(para)
    return None if pPr is None else pPr.find(qn("w:rPr"))


def _para_style_id(para):
    pPr = _pPr(para)
    if pPr is None:
        return None
    el = pPr.find(qn("w:pStyle"))
    return None if el is None else el.get(qn("w:val"))


def _run_style_id(run):
    rPr = run._r.find(qn("w:rPr"))
    if rPr is None:
        return None
    el = rPr.find(qn("w:rStyle"))
    return None if el is None else el.get(qn("w:val"))


def _style_by_id(doc, style_id):
    if not style_id:
        return None
    for st in doc.styles.element.findall(qn("w:style")):
        if st.get(qn("w:styleId")) == style_id:
            return st
    return None


def _style_chain(doc, style_id, limit=12):
    """按 basedOn 依次返回样式元素，避免环引用。"""
    out = []
    seen = set()
    sid = style_id
    while sid and sid not in seen and len(out) < limit:
        seen.add(sid)
        st = _style_by_id(doc, sid)
        if st is None:
            break
        out.append(st)
        bo = st.find(qn("w:basedOn"))
        sid = None if bo is None else bo.get(qn("w:val"))
    return out


def _child_attr(parent, child_tag, attr):
    if parent is None:
        return None
    el = parent.find(qn(child_tag))
    if el is None:
        return None
    return el.get(qn(attr))


def _eff_pPr_attr(doc, para, child_tag, attr):
    """段落属性取值：段落直接设置 -> 段落样式链。"""
    val = _child_attr(_pPr(para), child_tag, attr)
    if val is not None:
        return val
    for st in _style_chain(doc, _para_style_id(para)):
        val = _child_attr(st.find(qn("w:pPr")), child_tag, attr)
        if val is not None:
            return val
    return None


def _eff_jc(doc, para):
    """未显式设置时 docDefaults 的 pPrDefault 为空，实际左对齐。"""
    return _eff_pPr_attr(doc, para, "w:jc", "w:val") or "left"


def _int_or_none(val):
    try:
        return int(round(float(val)))
    except (TypeError, ValueError):
        return None


def _para_indent(doc, para):
    """返回 (firstLine, firstLineChars, left, hanging)，单位为缇/百分字符。"""
    return (
        _int_or_none(_eff_pPr_attr(doc, para, "w:ind", "w:firstLine")),
        _int_or_none(_eff_pPr_attr(doc, para, "w:ind", "w:firstLineChars")),
        _int_or_none(_eff_pPr_attr(doc, para, "w:ind", "w:left")),
        _int_or_none(_eff_pPr_attr(doc, para, "w:ind", "w:hanging")),
    )


def _para_spacing(doc, para):
    """返回 (line, lineRule, before, after)。"""
    return (
        _int_or_none(_eff_pPr_attr(doc, para, "w:spacing", "w:line")),
        _eff_pPr_attr(doc, para, "w:spacing", "w:lineRule"),
        _int_or_none(_eff_pPr_attr(doc, para, "w:spacing", "w:before")),
        _int_or_none(_eff_pPr_attr(doc, para, "w:spacing", "w:after")),
    )


def _para_numbering(doc, para):
    """返回 numId（0 表示编号已被取消），无编号返回 None。"""
    pPr = _pPr(para)
    numPr = None if pPr is None else pPr.find(qn("w:numPr"))
    if numPr is None:
        for st in _style_chain(doc, _para_style_id(para)):
            spPr = st.find(qn("w:pPr"))
            numPr = None if spPr is None else spPr.find(qn("w:numPr"))
            if numPr is not None:
                break
    if numPr is None:
        return None
    return _int_or_none(_child_attr(numPr, "w:numId", "w:val"))


def _para_page_break_before(doc, para):
    val = _eff_pPr_attr(doc, para, "w:pageBreakBefore", "w:val")
    return val not in ("0", "false", "off")  if val is not None else False


def _docdefaults_rPr(doc):
    dd = doc.styles.element.find(qn("w:docDefaults"))
    if dd is None:
        return None
    rd = dd.find(qn("w:rPrDefault"))
    return None if rd is None else rd.find(qn("w:rPr"))


def _rfonts_attr(rPr, attr):
    return _child_attr(rPr, "w:rFonts", attr)


def _resolve_effective_font(doc, para, run, attr):
    """字体解析：run -> 段落标记 -> 字符样式链 -> 段落样式链 -> docDefaults。"""
    rPr = run._r.find(qn("w:rPr")) if run is not None else None
    for src in (rPr, _para_mark_rPr(para)):
        val = _rfonts_attr(src, attr)
        if val:
            return val
    chains = []
    if run is not None:
        chains.append(_style_chain(doc, _run_style_id(run)))
    chains.append(_style_chain(doc, _para_style_id(para)))
    for chain in chains:
        for st in chain:
            val = _rfonts_attr(st.find(qn("w:rPr")), attr)
            if val:
                return val
    return _rfonts_attr(_docdefaults_rPr(doc), attr)


def _resolve_effective_eastasia(doc, para, run):
    return _resolve_effective_font(doc, para, run, "w:eastAsia")


def _resolve_effective_ascii(doc, para, run):
    return _resolve_effective_font(doc, para, run, "w:ascii")


def _run_effective_size_halfpt(doc, para, run):
    """字号（半磅）解析，顺序与字体一致。"""
    rPr = run._r.find(qn("w:rPr")) if run is not None else None
    for src in (rPr, _para_mark_rPr(para)):
        val = _child_attr(src, "w:sz", "w:val")
        if val is not None:
            return _int_or_none(val)
    chains = []
    if run is not None:
        chains.append(_style_chain(doc, _run_style_id(run)))
    chains.append(_style_chain(doc, _para_style_id(para)))
    for chain in chains:
        for st in chain:
            val = _child_attr(st.find(qn("w:rPr")), "w:sz", "w:val")
            if val is not None:
                return _int_or_none(val)
    return _int_or_none(_child_attr(_docdefaults_rPr(doc), "w:sz", "w:val"))


def _toggle_on(rPr, tag):
    if rPr is None:
        return None
    el = rPr.find(qn(tag))
    if el is None:
        return None
    val = el.get(qn("w:val"))
    return val not in ("0", "false", "off")


def _run_bold(doc, para, run):
    rPr = run._r.find(qn("w:rPr")) if run is not None else None
    for src in (rPr, _para_mark_rPr(para)):
        val = _toggle_on(src, "w:b")
        if val is not None:
            return val
    chains = []
    if run is not None:
        chains.append(_style_chain(doc, _run_style_id(run)))
    chains.append(_style_chain(doc, _para_style_id(para)))
    for chain in chains:
        for st in chain:
            val = _toggle_on(st.find(qn("w:rPr")), "w:b")
            if val is not None:
                return val
    return False


def _run_underlined(run):
    rPr = run._r.find(qn("w:rPr"))
    val = _child_attr(rPr, "w:u", "w:val")
    return val is not None and val != "none"


def _run_has_tab(run):
    return run._r.find(qn("w:tab")) is not None or "\t" in (run.text or "")


def _real_runs(para):
    """只保留带可见文字的 run，忽略换行、书签等空 run。"""
    return [r for r in para.runs if (r.text or "").strip() != ""]


def _tbl_borders(table):
    tblPr = table._tbl.find(qn("w:tblPr"))
    return None if tblPr is None else tblPr.find(qn("w:tblBorders"))


def _get_border(borders, side):
    """返回 (val, sz)，缺失返回 (None, None)。"""
    if borders is None:
        return (None, None)
    el = borders.find(qn("w:" + side))
    if el is None:
        return (None, None)
    return (el.get(qn("w:val")), _int_or_none(el.get(qn("w:sz"))))


def _has_line(borders, side):
    val, _sz = _get_border(borders, side)
    return val not in (None, "none", "nil")


def _is_double(borders, side):
    val, _sz = _get_border(borders, side)
    return val == "double"


def _is_single(borders, side):
    val, _sz = _get_border(borders, side)
    return val in ("single", "thinThickSmallGap", "thickThinSmallGap")


def _tbl_float_pr(table):
    tblPr = table._tbl.find(qn("w:tblPr"))
    return None if tblPr is None else tblPr.find(qn("w:tblpPr"))


def _tbl_jc(table):
    tblPr = table._tbl.find(qn("w:tblPr"))
    return _child_attr(tblPr, "w:jc", "w:val")


def _tbl_grid_cols(table):
    grid = table._tbl.find(qn("w:tblGrid"))
    if grid is None:
        return 0
    return len(grid.findall(qn("w:gridCol")))


def _cell_paragraphs(cell):
    """单元格内按文档顺序返回段落（含嵌套表格中的段落）。"""
    out = []
    for child in cell._tc.iterchildren():
        if child.tag == qn("w:p"):
            out.append(Paragraph(child, cell))
        elif child.tag == qn("w:tbl"):
            for row in Table(child, cell).rows:
                for sub in row.cells:
                    out.extend(_cell_paragraphs(sub))
    return out


def _table_paragraphs(table):
    out = []
    seen = set()
    for row in table.rows:
        for cell in row.cells:
            if id(cell._tc) in seen:
                continue
            seen.add(id(cell._tc))
            out.extend(_cell_paragraphs(cell))
    return out


def _table_text(table):
    return _squeeze("".join(p.text for p in _table_paragraphs(table)))


def _row_cells_unique(row):
    out = []
    seen = set()
    for cell in row.cells:
        if id(cell._tc) in seen:
            continue
        seen.add(id(cell._tc))
        out.append(cell)
    return out


def _iter_body_blocks(doc):
    """按文档顺序返回正文一级块 ("p"/"tbl", 对象)。"""
    out = []
    for child in doc.element.body.iterchildren():
        if child.tag == qn("w:p"):
            out.append(("p", Paragraph(child, doc)))
        elif child.tag == qn("w:tbl"):
            out.append(("tbl", Table(child, doc)))
    return out


def _all_paragraphs(doc):
    """正文全部段落（含表格内），按文档顺序。"""
    out = []
    for kind, obj in _iter_body_blocks(doc):
        if kind == "p":
            out.append(obj)
        else:
            out.extend(_table_paragraphs(obj))
    return out


def _para_images(para):
    """返回该段落内的图片列表 [(宽cm, 高cm, 宽高比, 是否浮动)]。"""
    out = []
    for tag, floating in (("wp:inline", False), ("wp:anchor", True)):
        for el in para._p.iter(qn(tag)):
            ext = el.find(qn("wp:extent"))
            if ext is None:
                continue
            cx = _int_or_none(ext.get("cx"))
            cy = _int_or_none(ext.get("cy"))
            if not cx or not cy:
                continue
            out.append((cx / EMU_PER_CM, cy / EMU_PER_CM, cx / float(cy), floating))
    return out


def _para_pict_count(para):
    """兼容 VML（w:pict）图形，仅计数，无法取得可靠尺寸。"""
    return len(list(para._p.iter(qn("w:pict"))))


def _collect_images(doc):
    """返回 [(段落, 宽cm, 高cm, 宽高比, 是否浮动)]。"""
    out = []
    for para in _all_paragraphs(doc):
        for w_cm, h_cm, ratio, floating in _para_images(para):
            out.append((para, w_cm, h_cm, ratio, floating))
    return out


def _is_badge(w_cm, h_cm):
    return (BADGE_W_CM[0] <= w_cm <= BADGE_W_CM[1]
            and BADGE_H_CM[0] <= h_cm <= BADGE_H_CM[1])


def _is_fig1(w_cm, h_cm, ratio):
    return 0.70 <= ratio <= 1.45 and 3.0 <= w_cm <= 12.0 and 3.0 <= h_cm <= 12.0


def _sect_pr(doc):
    return doc.element.body.find(qn("w:sectPr"))


def _sect_page(doc):
    """返回 (宽cm, 高cm, {边距: cm}, orient)。"""
    sect = _sect_pr(doc)
    if sect is None:
        return (None, None, {}, None)
    pg_sz = sect.find(qn("w:pgSz"))
    pg_mar = sect.find(qn("w:pgMar"))
    w = _int_or_none(pg_sz.get(qn("w:w"))) if pg_sz is not None else None
    h = _int_or_none(pg_sz.get(qn("w:h"))) if pg_sz is not None else None
    orient = pg_sz.get(qn("w:orient")) if pg_sz is not None else None
    mar = {}
    for side in ("top", "bottom", "left", "right"):
        val = _int_or_none(pg_mar.get(qn("w:" + side))) if pg_mar is not None else None
        mar[side] = None if val is None else val / TWIPS_PER_CM
    return (None if w is None else w / TWIPS_PER_CM,
            None if h is None else h / TWIPS_PER_CM,
            mar, orient)


def _open_word_document(path):
    return Document(path)


def _render_to_pdf(docx_path):
    """优先 Word COM，其次 WPS；不可用时返回 (None, 原因)。"""
    if os.name != "nt":
        return (None, "非 Windows 平台，无法调用 Office COM")
    try:
        import win32com.client as win32
    except ImportError:
        return (None, "未安装 pywin32，无法调用 Office COM")
    tmp_dir = tempfile.mkdtemp(prefix="oe021_")
    out_pdf = os.path.join(tmp_dir, "render.pdf")
    src = os.path.abspath(docx_path)
    errors = []
    for prog_id in ("Word.Application", "KWps.Application"):
        app = None
        try:
            app = win32.DispatchEx(prog_id)
            try:
                app.Visible = False
            except Exception:
                pass
            try:
                app.DisplayAlerts = 0
            except Exception:
                pass
            doc = app.Documents.Open(src, ReadOnly=True, AddToRecentFiles=False)
            try:
                doc.ExportAsFixedFormat(OutputFileName=out_pdf, ExportFormat=17)
            except Exception:
                doc.SaveAs2(out_pdf, 17)
            doc.Close(0)
            app.Quit()
            if os.path.isfile(out_pdf) and os.path.getsize(out_pdf) > 0:
                return (out_pdf, "")
            errors.append("%s 未生成 PDF" % prog_id)
        except Exception as exc:
            errors.append("%s: %s" % (prog_id, exc))
            try:
                if app is not None:
                    app.Quit()
            except Exception:
                pass
    shutil.rmtree(tmp_dir, ignore_errors=True)
    return (None, "；".join(errors) or "渲染失败")


def _build_page_locator(pdf_path):
    """返回每页压缩后的文本列表；渲染或解析失败返回 None。"""
    if fitz is None or not pdf_path:
        return None
    try:
        doc = fitz.open(pdf_path)
    except Exception:
        return None
    try:
        return [_squeeze(page.get_text()) for page in doc]
    except Exception:
        return None
    finally:
        try:
            doc.close()
        except Exception:
            pass


def _find_page(pages, needle):
    """返回首个包含 needle 的页码（1 起），未找到返回 0。"""
    if not pages or not needle:
        return 0
    for idx, text in enumerate(pages):
        if needle in text:
            return idx + 1
    return 0


def _cell_paras_containing(table, key):
    """返回表格中首个包含 key 的单元格的段落列表。"""
    if table is None:
        return []
    for row in table.rows:
        for cell in _row_cells_unique(row):
            paras = _cell_paragraphs(cell)
            if key in _squeeze("".join(p.text for p in paras)):
                return paras
    return []


def _find_table(tables, keys):
    for table in tables:
        text = _table_text(table)
        if all(k in text for k in keys):
            return table
    return None


def _build_ctx(doc):
    """一次性提取后续所有判定需要的结构信息。"""
    ctx = {"doc": doc}
    blocks = _iter_body_blocks(doc)
    ctx["blocks"] = blocks
    tables = [obj for kind, obj in blocks if kind == "tbl"]
    ctx["tables"] = tables

    ctx["t_main1"] = _find_table(tables, ("研究目的及意义",))
    ctx["t_main2"] = _find_table(tables, ("研究内容与技术方案",))
    ctx["t_ref"] = _find_table(tables, ("[1]", "[4]"))
    ctx["t_opinion"] = _find_table(tables, ("指导教师意见",))
    ctx["t_schedule"] = _find_table(tables, SCHEDULE_HEADER)

    # 注意：ctx 中持有段落对象本身，避免 lxml 代理对象被回收后 id() 复用导致误判
    opinion_paras = _table_paragraphs(ctx["t_opinion"]) if ctx["t_opinion"] is not None else []
    ctx["opinion_paras"] = opinion_paras
    opinion_ids = set(id(p._p) for p in opinion_paras)
    ctx["opinion_ids"] = opinion_ids

    # 封面范围：正文首个「报告名称行 / 填写要求标题 / 表格框」之前的一级段落
    cover_end = len(blocks)
    for idx, (kind, obj) in enumerate(blocks):
        if kind == "tbl":
            cover_end = idx
            break
        sq = _squeeze(obj.text)
        if sq == "开题报告填写要求" or NAME_ROW_RE.match(sq) and idx > 3:
            cover_end = idx
            break
    ctx["cover_end"] = cover_end
    cover = [obj for kind, obj in blocks[:cover_end] if kind == "p"]
    ctx["cover_paras"] = cover
    ctx["info_rows"] = [p for p in cover
                        if ("：" in p.text or ":" in p.text)
                        and any(_run_underlined(r) for r in p.runs)]
    ctx["cover_title"] = None
    for para in cover:
        sq = _squeeze(para.text)
        if sq.startswith("毕业设计") and sq.endswith("开题报告"):
            ctx["cover_title"] = para
            break

    # 章标题（排除意见签署表内的「一、优点」等条目）
    chapters = []
    for para in _all_paragraphs(doc):
        if id(para._p) in opinion_ids:
            continue
        sq = _squeeze(para.text)
        if len(sq) >= 6 and CHAPTER_RE.match(sq):
            chapters.append(para)
    ctx["chapters"] = chapters

    # 正文段落范围
    main1 = _cell_paras_containing(ctx["t_main1"], "研究目的及意义")
    main2 = _cell_paras_containing(ctx["t_main2"], "研究内容与技术方案")
    refs = _table_paragraphs(ctx["t_ref"]) if ctx["t_ref"] is not None else []
    ctx["main_cell_paras"] = main1 + main2
    ctx["ref_paras"] = [p for p in refs if REF_HEAD_RE.match(_squeeze(p.text))]
    ctx["prose_paras"] = [p for p in (main1 + main2 + refs)
                          if _squeeze(p.text) and FIG1_CAPTION_KEY not in p.text]
    ctx["long_paras"] = [p for p in (main1 + main2) if len(_squeeze(p.text)) >= 40]

    # 名称行（封面主标题不计入）
    ctx["name_rows"] = [obj for kind, obj in blocks[ctx["cover_end"]:]
                        if kind == "p" and NAME_ROW_RE.match(_squeeze(obj.text))]

    # 图片与「图1」图题
    ctx["images"] = _collect_images(doc)
    ctx["caption_para"] = None
    for para in _all_paragraphs(doc):
        if FIG1_CAPTION_KEY in para.text:
            ctx["caption_para"] = para
            break
    ctx["body_text"] = _squeeze("".join(p.text for p in _all_paragraphs(doc)))
    ctx["all_paras"] = _all_paragraphs(doc)
    return ctx


def _badge_image(ctx):
    """封面校徽图片：封面段落中的首张图片。"""
    cover_ids = set(id(p._p) for p in ctx["cover_paras"])
    for para, w_cm, h_cm, ratio, floating in ctx["images"]:
        if id(para._p) in cover_ids:
            return (para, w_cm, h_cm, ratio, floating)
    return None


def _fig1_image(ctx):
    """「图1」示意图：图题所在段落内的图片，或其紧邻的前一张图片。"""
    caption = ctx["caption_para"]
    if caption is None:
        return None
    for para, w_cm, h_cm, ratio, floating in ctx["images"]:
        if id(para._p) == id(caption._p):
            return (para, w_cm, h_cm, ratio, floating)
    paras = _all_paragraphs(ctx["doc"])
    try:
        pos = [id(p._p) for p in paras].index(id(caption._p))
    except ValueError:
        return None
    for para in reversed(paras[:pos]):
        imgs = _para_images(para)
        if imgs:
            w_cm, h_cm, ratio, floating = imgs[0]
            if _is_fig1(w_cm, h_cm, ratio):
                return (para, w_cm, h_cm, ratio, floating)
            return None
    return None


def _blank_para(para):
    """无文字且无图片的空白段落。"""
    return (not _squeeze(para.text)) and (not _para_images(para)) and (not _para_pict_count(para))


def _para_height_pt(doc, para):
    """段落占用的竖向高度（磅）：行距按固定值/倍数折算，并计入段前段后间距。"""
    line, rule, before, after = _para_spacing(doc, para)
    size_hp = _run_effective_size_halfpt(doc, para, None) or 24
    size_pt = size_hp / 2.0
    if line and rule in ("exact", "atLeast"):
        height = line / 20.0
    elif line:
        height = size_pt * 1.3 * (line / 240.0)
    else:
        height = size_pt * 1.3
    return height + (before or 0) / 20.0 + (after or 0) / 20.0


def _max_blank_band(doc, paras):
    """最长连续空白段落的 (段数, 累计高度磅)，以累计高度最大者为准。"""
    best = (0, 0.0)
    run = 0
    height = 0.0
    for para in paras:
        if _blank_para(para):
            run += 1
            height += _para_height_pt(doc, para)
            if height > best[1]:
                best = (run, height)
        else:
            run = 0
            height = 0.0
    return best


def _usable_height_pt(doc):
    """可用版面高度（磅）= 页高 - 上下页边距；取不到页面尺寸时返回 None。"""
    _w_cm, h_cm, mar, _orient = _sect_page(doc)
    if not h_cm or mar.get("top") is None or mar.get("bottom") is None:
        return None
    return (h_cm - mar["top"] - mar["bottom"]) * PT_PER_CM


def _chapter_body(para):
    """去掉章序号后的标题正文，例如「四．工作进度安排」-> 「工作进度安排」。"""
    sq = _squeeze(para.text)
    return CHAPTER_RE.sub("", sq)


def _dim1_reason(doc, ctx):
    """维度1 可用与可修改性：不通过时返回原因，通过返回空串。"""
    text_len = len(ctx["body_text"])
    if not ctx["tables"]:
        return "文档内无任何 Word 表格对象，正文表格框缺失，无法直接修改"
    if text_len < 500:
        return "可提取文字仅 %d 字，正文疑似以整页截图替代文字，无法编辑修改" % text_len
    if len(ctx["images"]) >= 6 and text_len < 1200:
        return "图片 %d 张而可提取文字仅 %d 字，正文疑似以图片替代文字" % (
            len(ctx["images"]), text_len)
    editable = sum(1 for para in _all_paragraphs(doc) if _squeeze(para.text))
    if editable < 20:
        return "可编辑文字段落仅 %d 段，版面内容缺失严重，无法直接使用" % editable
    return ""


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
    """入参为目录路径，自动取其中的 .docx 作为待评文件；返回结构化评测结果。"""
    result = {
        "id": SCRIPT_ID,
        "file_name": "",
        "status": "ok",
        "error": "",
        "dim1_pass": True,
        "dim1_reason": "",
        "dim2_items": [],
        "total_score": 0,
        "max_score": MAX_SCORE_TOTAL,
    }
    try:
        if not os.path.isdir(dir_path):
            raise FileNotFoundError("目录不存在：%s" % dir_path)
        cands = sorted(f for f in os.listdir(dir_path)
                       if f.lower().endswith(".docx") and not f.startswith("~"))
        if not cands:
            raise FileNotFoundError("目录中未找到 .docx 文件：%s" % dir_path)
        result["file_name"] = cands[0]
        return _evaluate_impl(os.path.join(dir_path, cands[0]), result)
    except Exception as exc:
        if _is_damaged_package(exc):
            result["status"] = "ok"
            result["error"] = ""
            result["dim1_pass"] = False
            result["dim1_reason"] = "交付文件结构不完整，无法正常打开：%s: %s" % (
                type(exc).__name__, exc)
            result["dim2_items"] = []
            result["total_score"] = 0
            return result
        # status=error 时：error 承载失败摘要；维度1 未实际判定，
        # dim1_pass / dim1_reason 一律留空，分值与满分均记 0。
        result["status"] = "error"
        result["error"] = str(exc).strip() or type(exc).__name__
        result["dim1_pass"] = ""
        result["dim1_reason"] = ""
        result["dim2_items"] = []
        result["total_score"] = 0
        result["max_score"] = 0
        return result


def _evaluate_impl(docx_path, result):
    doc = _open_word_document(docx_path)
    ctx = _build_ctx(doc)

    # ---------- 维度1 可用与可修改性 ----------
    reason = _dim1_reason(doc, ctx)
    if reason:
        result["dim1_pass"] = False
        result["dim1_reason"] = reason
        result["dim2_items"] = []
        result["total_score"] = 0
        return result

    # 分页相关判定需要渲染证据：加分项无证据时按 XML 结构降级判定，
    # 减分项无证据时一律不扣分（减分需证据）。
    pdf_path, render_err = _render_to_pdf(docx_path)
    pages = _build_page_locator(pdf_path)
    render_note = "" if pages else (render_err or "PDF 解析失败")

    results = []

    # ---------- 01 +5 页面设置为 A4 纵向且四边页边距达标 ----------
    # ① 纸张宽 20.9~21.1cm、高 29.6~29.8cm 且为纵向
    # ② 上 2.2cm、下 2.0cm、左 2.5cm、右 2.3cm，逐边误差 <= 0.1cm
    pg_w, pg_h, pg_mar, pg_orient = _sect_page(doc)
    size_ok = (pg_w is not None and pg_h is not None
               and PAGE_W_CM[0] <= pg_w <= PAGE_W_CM[1]
               and PAGE_H_CM[0] <= pg_h <= PAGE_H_CM[1]
               and (pg_orient or "portrait") == "portrait")
    bad_edges = []
    for _side, _label in (("top", "上"), ("bottom", "下"), ("left", "左"), ("right", "右")):
        _cur = pg_mar.get(_side)
        _target = MARGIN_TARGET_CM[_side]
        if _cur is None or abs(_cur - _target) > MARGIN_TOL_CM + 1e-9:
            bad_edges.append("%s%s（应%.1fcm）" % (
                _label, "缺失" if _cur is None else "%.3fcm" % _cur, _target))
    if size_ok and not bad_edges:
        results.append("+5：页面为 A4 纵向 %.2f×%.2fcm，页边距上下左右 %.2f/%.2f/%.2f/%.2fcm 均在允差内"
                       % (pg_w, pg_h, pg_mar["top"], pg_mar["bottom"],
                          pg_mar["left"], pg_mar["right"]))
    else:
        results.append("+0：页面设置未达标（%s%s）" % (
            "" if size_ok else "纸张 %s×%s；" % (
                "?" if pg_w is None else "%.2fcm" % pg_w,
                "?" if pg_h is None else "%.2fcm" % pg_h),
            "、".join(bad_edges) or "页边距达标"))

    # ---------- 02 +3 封面顶部居中校徽图片 ----------
    # ① 封面段落中存在图片且所在段落居中
    # ② 显示宽 14.5~15.1cm、高 4.0~4.4cm
    # ③ 显示宽高比接近 3.54:1，未被拉伸压扁
    badge = _badge_image(ctx)
    if badge is None:
        results.append("+0：封面未发现校徽图片")
    else:
        _bp, _bw, _bh, _br, _bfloat = badge
        _center = _eff_jc(doc, _bp) == "center" or _bfloat
        if _is_badge(_bw, _bh) and _center and abs(_br - 3.54) <= 0.20:
            results.append("+3：封面顶部居中校徽图片 %.2f×%.2fcm（宽高比 %.3f），等比未变形"
                           % (_bw, _bh, _br))
        else:
            results.append("+0：封面校徽图片不达标（%.2f×%.2fcm 宽高比 %.3f，%s）"
                           % (_bw, _bh, _br, "居中" if _center else "未居中"))

    # ---------- 03 +3 封面校名与主标题字体字号 ----------
    # ① 「澜州应用技术学院」单独一行、居中、楷体_GB2312 二号（26 磅 = 52 半磅）
    # ② 主标题单独一行、居中、黑体 一号（36 磅 = 72 半磅）
    school_para = None
    for _p in ctx["cover_paras"]:
        if _squeeze(_p.text) == "澜州应用技术学院":
            school_para = _p
            break
    title_para = ctx["cover_title"]
    _miss3 = []
    if school_para is None:
        _miss3.append("缺少单独一行的「澜州应用技术学院」")
    else:
        _runs = _real_runs(school_para)
        if _eff_jc(doc, school_para) != "center":
            _miss3.append("校名行未居中")
        if not all(_is_kaiti_gb(_resolve_effective_eastasia(doc, school_para, r)) for r in _runs):
            _miss3.append("校名行非楷体_GB2312")
        if not all(_run_effective_size_halfpt(doc, school_para, r) == 52 for r in _runs):
            _miss3.append("校名行非二号（26 磅）")
    if title_para is None:
        _miss3.append("缺少封面主标题行")
    else:
        _runs = _real_runs(title_para)
        if _eff_jc(doc, title_para) != "center":
            _miss3.append("主标题未居中")
        if not all(_is_heiti(_resolve_effective_eastasia(doc, title_para, r)) for r in _runs):
            _miss3.append("主标题非黑体")
        if not all(_run_effective_size_halfpt(doc, title_para, r) == 72 for r in _runs):
            _miss3.append("主标题非一号（36 磅）")
    if not _miss3:
        results.append("+3：封面校名为楷体_GB2312 二号、主标题为黑体 一号，均单独一行居中")
    else:
        results.append("+0：封面校名或主标题字体字号不达标（%s）" % "、".join(_miss3))

    # ---------- 04 +3 封面信息栏「标签+填写内容」结构 ----------
    # ① 每行标签为三号（16 磅 = 32 半磅）加粗
    # ② 填写内容加下划线且字号统一为三号
    # ③ 各行标签左端在同一竖直位置（对齐方式与左/首行缩进一致）
    # ④ 行距 1.35 倍（line=324 自动）、段前 4 磅（before=80）
    info_rows = ctx["info_rows"]
    _miss4 = []
    if len(info_rows) < 5:
        _miss4.append("信息栏行数仅 %d 行" % len(info_rows))
    _anchors = set()
    for _p in info_rows:
        _fl, _flc, _left, _hang = _para_indent(doc, _p)
        _anchors.add((_eff_jc(doc, _p), _fl or 0, _left or 0))
        _labels = [r for r in _real_runs(_p) if "：" in (r.text or "") or ":" in (r.text or "")]
        if not _labels:
            _miss4.append("「%s」缺少标签文字" % _squeeze(_p.text)[:8])
        for _r in _labels:
            if not _run_bold(doc, _p, _r) or _run_effective_size_halfpt(doc, _p, _r) != 32:
                _miss4.append("「%s」标签非三号加粗" % _squeeze(_r.text)[:8])
        _fills = [r for r in _real_runs(_p) if _run_underlined(r)]
        if not _fills:
            _miss4.append("「%s」填写内容无下划线" % _squeeze(_p.text)[:8])
        for _r in _fills:
            if _run_effective_size_halfpt(doc, _p, _r) != 32:
                _miss4.append("「%s」填写内容非三号" % _squeeze(_r.text)[:8])
        _line, _rule, _before, _after = _para_spacing(doc, _p)
        if _line is None or abs(_line - 324) > 10 or (_rule or "auto") != "auto":
            _miss4.append("「%s」行距非 1.35 倍" % _squeeze(_p.text)[:8])
        if _before is None or abs(_before - 80) > 4:
            _miss4.append("「%s」段前非 4 磅" % _squeeze(_p.text)[:8])
    if info_rows and len(_anchors) > 1:
        _miss4.append("各行标签左端未对齐（%d 种缩进）" % len(_anchors))
    if info_rows and not _miss4:
        results.append("+3：封面信息栏 %d 行均为三号加粗标签+三号下划线填写内容，左端对齐、行距 1.35 倍、段前 4 磅"
                       % len(info_rows))
    else:
        results.append("+0：封面信息栏不达标（%s）" % ("；".join(dict.fromkeys(_miss4))[:200]
                                                 or "未识别到信息栏"))

    # ---------- 05 +3 封面下部单独一行居中「年    月    日」 ----------
    # ① 封面存在压缩后等于「年月日」的独立段落且居中
    # ② 宋体 小四（12 磅 = 24 半磅）
    date_para = None
    for _p in ctx["cover_paras"]:
        if _squeeze(_p.text) == "年月日":
            date_para = _p
            break
    if date_para is None:
        results.append("+0：封面下部缺少单独一行的「年    月    日」")
    else:
        _runs = _real_runs(date_para)
        _ok = (_eff_jc(doc, date_para) == "center"
               and all(_is_songti(_resolve_effective_eastasia(doc, date_para, r)) for r in _runs)
               and all(_run_effective_size_halfpt(doc, date_para, r) == 24 for r in _runs))
        if _ok:
            results.append("+3：封面下部保留单独一行居中的「年    月    日」，宋体小四")
        else:
            results.append("+0：封面「年    月    日」行的居中或宋体小四不达标")

    # ---------- 06 +5 正文装入 4 个表格框 ----------
    # ① 正文一级块中恰有 4 个表格
    # ② 依次承载「研究目的及意义/行业研究现状」「研究内容与技术方案/工作进度」
    #    「主要参考文献」「意见签署」四部分
    # ③ 每个表格外框上下为双细线、左右为单线
    frames = [("研究目的及意义/行业研究现状", ctx["t_main1"]),
              ("研究内容与技术方案/工作进度", ctx["t_main2"]),
              ("主要参考文献", ctx["t_ref"]),
              ("意见签署", ctx["t_opinion"])]
    _miss6 = [name for name, tbl in frames if tbl is None]
    _ids7 = set(id(tbl._tbl) for _n, tbl in frames if tbl is not None)
    if len(ctx["tables"]) != 4:
        _miss6.append("表格框数量为 %d 个" % len(ctx["tables"]))
    if len(_ids7) != 4:
        _miss6.append("四部分未分置于 4 个独立表格框")
    for name, tbl in frames:
        if tbl is None:
            continue
        _bd = _tbl_borders(tbl)
        if not (_is_double(_bd, "top") and _is_double(_bd, "bottom")):
            _miss6.append("「%s」框上下边框非双细线" % name)
        if not (_is_single(_bd, "left") and _is_single(_bd, "right")):
            _miss6.append("「%s」框左右边框非单线" % name)
    if not _miss6:
        results.append("+5：正文内容装入 4 个表格框，依次承载四部分，外框上下双细线、左右单线")
    else:
        results.append("+0：正文表格框结构或边框不达标（%s）"
                       % "；".join(dict.fromkeys(_miss6))[:180])

    # ---------- 07 +3 共 3 处居中报告名称行 ----------
    # ① 封面主标题之外，正文与意见签署部分共 3 处报告名称行
    # ② 每处均居中、小二（18 磅 = 36 半磅）、加粗
    name_rows = ctx["name_rows"]
    _miss7 = []
    if len(name_rows) != 3:
        _miss7.append("名称行共 %d 处" % len(name_rows))
    for _p in name_rows:
        _runs = _real_runs(_p)
        if _eff_jc(doc, _p) != "center":
            _miss7.append("「%s」未居中" % _squeeze(_p.text)[:12])
        if not all(_run_effective_size_halfpt(doc, _p, r) == 36 for r in _runs):
            _miss7.append("「%s」非小二（18 磅）" % _squeeze(_p.text)[:12])
        if not all(_run_bold(doc, _p, r) for r in _runs):
            _miss7.append("「%s」未加粗" % _squeeze(_p.text)[:12])
    if not _miss7:
        results.append("+3：正文与意见签署部分共 3 处居中报告名称行，小二加粗")
    else:
        results.append("+0：报告名称行不达标（%s）" % "；".join(dict.fromkeys(_miss7))[:160])

    # ---------- 08 +5 正文汉字宋体、西文 Times New Roman、字号统一小四 ----------
    # ① 覆盖各章标题、（一）（二）分级标题、（1）～（5）条目、正文段落与参考文献
    # ② 含汉字的 run 东亚字体须为宋体，含西文字母数字的 run 西文字体须为 Times New Roman
    # ③ 全部 run 字号为小四（12 磅 = 24 半磅）
    _bad_font = []
    _bad_size = []
    for _p in ctx["prose_paras"]:
        for _r in _real_runs(_p):
            _txt = _r.text or ""
            if _has_cjk(_txt) and not _is_songti(_resolve_effective_eastasia(doc, _p, _r)):
                _bad_font.append("「%s」汉字为 %s" % (
                    _squeeze(_txt)[:8], _resolve_effective_eastasia(doc, _p, _r)))
            if _has_latin(_txt) and not _is_tnr(_resolve_effective_ascii(doc, _p, _r)):
                _bad_font.append("「%s」西文为 %s" % (
                    _squeeze(_txt)[:8], _resolve_effective_ascii(doc, _p, _r)))
            _sz = _run_effective_size_halfpt(doc, _p, _r)
            if _sz != 24:
                _bad_size.append("「%s」%s 磅" % (
                    _squeeze(_txt)[:8], "缺失" if _sz is None else "%.1f" % (_sz / 2.0)))
    if ctx["prose_paras"] and not _bad_font and not _bad_size:
        results.append("+5：正文 %d 段（含章标题、分级标题、条目与参考文献）汉字为宋体、西文为 Times New Roman，字号统一小四"
                       % len(ctx["prose_paras"]))
    elif _bad_font or _bad_size:
        results.append("+0：正文字体或字号不统一（%s）"
                       % "；".join(dict.fromkeys(_bad_font[:3] + _bad_size[:3])))
    else:
        results.append("+0：未识别到正文段落")

    # ---------- 09 +5 正文段落行距为固定值 22 磅 ----------
    # ① 取两处正文表格框内 40 字以上的正文段落
    # ② 行距须为 line=440、lineRule=exact（固定值 22 磅）
    _bad_line = []
    for _p in ctx["long_paras"]:
        _line, _rule, _before, _after = _para_spacing(doc, _p)
        if _line != 440 or _rule != "exact":
            _bad_line.append("「%s」line=%s/%s" % (_squeeze(_p.text)[:8], _line, _rule))
    if ctx["long_paras"] and not _bad_line:
        results.append("+5：正文 %d 个段落行距均为固定值 22 磅" % len(ctx["long_paras"]))
    else:
        results.append("+0：正文段落行距非固定值 22 磅（%s）"
                       % ("；".join(_bad_line[:3]) or "未识别到正文段落"))

    # ---------- 10 +3 正文段落首行缩进 2 字符 ----------
    # ① firstLine 在 440~520 缇（约 0.78~0.92cm），或 firstLineChars 约 200
    _bad_ind = []
    for _p in ctx["long_paras"]:
        _fl, _flc, _left, _hang = _para_indent(doc, _p)
        if not ((_fl is not None and 440 <= _fl <= 520) or (_flc is not None and 190 <= _flc <= 210)):
            _bad_ind.append("「%s」firstLine=%s/chars=%s" % (_squeeze(_p.text)[:8], _fl, _flc))
    if ctx["long_paras"] and not _bad_ind:
        results.append("+3：正文段落首行缩进均为 2 字符（约 0.85cm）")
    else:
        results.append("+0：正文段落首行缩进非 2 字符（%s）"
                       % ("；".join(_bad_ind[:3]) or "未识别到正文段落"))

    # ---------- 11 +3 五个章标题加粗且为宋体 ----------
    # ① 一．研究目的及意义 / 二．行业研究现状 / 三．研究内容与技术方案 /
    #    四．工作进度安排 / 五．主要参考文献 共五处章标题
    # ② 每处章标题的全部 run 均加粗，东亚字体为宋体
    _chs = ctx["chapters"]
    _bad_ch = []
    for _p in _chs:
        for _r in _real_runs(_p):
            if not _run_bold(doc, _p, _r):
                _bad_ch.append("「%s」未加粗" % _squeeze(_p.text)[:10])
                break
            if not _is_songti(_resolve_effective_eastasia(doc, _p, _r)):
                _bad_ch.append("「%s」为 %s" % (_squeeze(_p.text)[:10],
                                              _resolve_effective_eastasia(doc, _p, _r)))
                break
    if len(_chs) == 5 and not _bad_ch:
        results.append("+3：五个章标题均为宋体加粗")
    elif len(_chs) != 5:
        results.append("+0：未识别到五个章标题（实际 %d 处）" % len(_chs))
    else:
        results.append("+0：章标题字体或加粗不达标（%s）" % "；".join(dict.fromkeys(_bad_ch))[:160])

    # ---------- 12 +1 章序号为手动录入的中文序号，未使用自动编号 ----------
    # ① 五个章标题依次以 一 二 三 四 五 开头
    # ② 段落无自动编号（无 numPr 或 numId=0）
    _bad_num = []
    _order = "一二三四五"
    for _i, _p in enumerate(_chs):
        _sq = _squeeze(_p.text)
        if _i < len(_order) and not _sq.startswith(_order[_i]):
            _bad_num.append("第 %d 处以「%s」开头" % (_i + 1, _sq[:4]))
        _nid = _para_numbering(doc, _p)
        if _nid not in (None, 0):
            _bad_num.append("「%s」带自动编号 numId=%s" % (_sq[:8], _nid))
    if len(_chs) == 5 and not _bad_num:
        results.append("+1：章序号为手动录入的中文序号，未使用自动编号")
    else:
        results.append("+0：章序号不达标（%s）"
                       % ("；".join(dict.fromkeys(_bad_num))[:160] or "未识别到五个章标题"))

    # ---------- 13 +3 工作进度安排表为三列结构且四个阶段行完整 ----------
    # ① 表格为 3 列，含表头行「阶段 / 主要任务 / 时间安排」
    # ② 四个阶段行（前期调研、结构设计、仿真验证、成果整理）及其起止月份齐全
    _tsch = ctx["t_schedule"]
    _miss13 = []
    if _tsch is None:
        _miss13.append("未找到工作进度安排表")
    else:
        if _tbl_grid_cols(_tsch) != 3:
            _miss13.append("表格为 %d 列" % _tbl_grid_cols(_tsch))
        _rows_txt = [_squeeze("".join(_c.text for _c in _row_cells_unique(_r)))
                     for _r in _tsch.rows]
        if not any(all(_k in _t for _k in SCHEDULE_HEADER) for _t in _rows_txt):
            _miss13.append("缺表头行「阶段/主要任务/时间安排」")
        for _stage in SCHEDULE_ROWS:
            if not any(all(_k in _t for _k in _stage) for _t in _rows_txt):
                _miss13.append("缺「%s %s-%s」行" % _stage)
    if not _miss13:
        results.append("+3：工作进度安排表为三列结构，表头与四个阶段行完整")
    else:
        results.append("+0：工作进度安排表不完整（%s）" % "；".join(_miss13)[:160])

    # ---------- 14 +1 工作进度表单元格小四居中且无制表符 ----------
    # ① 表头行与 4 行阶段数据的单元格段落均居中
    # ② 单元格文字为小四（24 半磅），且不使用制表符排版
    _miss14 = []
    if _tsch is None:
        _miss14.append("未找到工作进度安排表")
    else:
        _grid_rows = [_r for _r in _tsch.rows if len(_row_cells_unique(_r)) == 3]
        if len(_grid_rows) < 5:
            _miss14.append("三列数据行仅 %d 行" % len(_grid_rows))
        for _r in _grid_rows:
            for _c in _row_cells_unique(_r):
                for _p in _c.paragraphs:
                    if not _squeeze(_p.text):
                        continue
                    if _eff_jc(doc, _p) != "center":
                        _miss14.append("「%s」未居中" % _squeeze(_p.text)[:8])
                    for _run in _real_runs(_p):
                        if _run_effective_size_halfpt(doc, _p, _run) != 24:
                            _miss14.append("「%s」非小四" % _squeeze(_p.text)[:8])
                            break
                    if any(_run_has_tab(_run) for _run in _p.runs):
                        _miss14.append("「%s」含制表符" % _squeeze(_p.text)[:8])
    if not _miss14:
        results.append("+1：工作进度表单元格均为小四居中，未使用制表符排版")
    else:
        results.append("+0：工作进度表单元格排版不达标（%s）"
                       % "；".join(dict.fromkeys(_miss14))[:160])

    # ---------- 15 +1 参考文献 4 条顶格、无首行缩进与自动编号 ----------
    # ① 以 [1]~[4] 开头的条目共 4 条
    # ② 首行缩进与左缩进均为 0，且无自动编号
    _refs = ctx["ref_paras"]
    _miss15 = []
    if len(_refs) != 4:
        _miss15.append("条目为 %d 条" % len(_refs))
    for _p in _refs:
        _fl, _flc, _left, _hang = _para_indent(doc, _p)
        if (_fl or 0) != 0 or (_flc or 0) != 0 or (_left or 0) != 0:
            _miss15.append("「%s」缩进 firstLine=%s/left=%s" % (_squeeze(_p.text)[:6], _fl, _left))
        if _para_numbering(doc, _p) not in (None, 0):
            _miss15.append("「%s」带自动编号" % _squeeze(_p.text)[:6])
    if not _miss15:
        results.append("+1：参考文献 4 条均顶格排列，无首行缩进与自动编号")
    else:
        results.append("+0：参考文献条目排版不达标（%s）" % "；".join(dict.fromkeys(_miss15))[:160])

    # ---------- 16 +3 文末 3 行意见签署表，含签名横线与日期行 ----------
    # ① 依次为「指导教师意见」「教研室审查意见」「所在二级学院审查意见」三行
    # ② 每行含对应签名栏（指导教师/专业负责人/二级学院负责人）与「年 月 日」
    _topi = ctx["t_opinion"]
    _sign_labels = ("指导教师", "专业负责人", "二级学院负责人")
    _miss16 = []
    _opi_rows = []
    if _topi is None:
        _miss16.append("未找到意见签署表")
    else:
        for _r in _topi.rows:
            _opi_rows.append(_squeeze("".join(_c.text for _c in _row_cells_unique(_r))))
        if len(_opi_rows) != 3:
            _miss16.append("意见栏为 %d 行" % len(_opi_rows))
        for _i, _title in enumerate(OPINION_TITLES):
            _txt = _opi_rows[_i] if _i < len(_opi_rows) else ""
            if _title not in _txt:
                _miss16.append("缺「%s」" % _title)
                continue
            if not any(re.search(_lab + "[：:]", _txt) for _lab in _sign_labels):
                _miss16.append("「%s」缺签名栏" % _title)
            if not re.search(r"年.{0,8}月.{0,8}日", _txt):
                _miss16.append("「%s」缺日期行" % _title)
    if not _miss16:
        results.append("+3：文末意见签署表为 3 行，签名栏与「年 月 日」齐全")
    else:
        results.append("+0：意见签署表不完整（%s）" % "；".join(dict.fromkeys(_miss16))[:160])

    # ---------- 17 +3 三个意见栏除栏目名、签名横线、日期行外保持空白 ----------
    # ① 逐行剔除栏目名、签名栏字样、日期行、下划线与空白后应无剩余内容
    _miss17 = []
    if _topi is None:
        _miss17.append("未找到意见签署表")
    for _i, _title in enumerate(OPINION_TITLES):
        if _i >= len(_opi_rows):
            continue
        _res = _opi_rows[_i].replace(_title, "")
        for _lab in _sign_labels:
            _res = _res.replace(_lab + "：", "").replace(_lab + ":", "")
        _res = re.sub(r"年.{0,8}月.{0,8}日", "", _res)
        _res = re.sub(r"[_\-—－–．.、,，。…：:\s]", "", _res)
        if len(_res) > 10:
            _miss17.append("「%s」残留 %d 字（%s…）" % (_title, len(_res), _res[:10]))
    if not _miss17:
        results.append("+3：三个意见栏均仅保留栏目名、签名栏与日期行，正文留空待签")
    else:
        results.append("+0：意见栏被填入内容（%s）" % "；".join(_miss17)[:160])

    # ---------- 18 +5 全文无竖向排列的技术路线流程图 ----------
    # ① 不存在宽高比 ≤0.60 的竖长条图片（原竖向流程图约 5.40×19.24cm，比值 0.28）
    # ② 全文不出现「现状调研剖析 / 明确优化设计目标 / 形成完整设计方案」等流程节点文字
    _v_imgs = [(w, h, ra) for _p, w, h, ra, _fl in ctx["images"] if ra <= FLOW_V_RATIO_MAX]
    _v_words = [w for w in FLOW_V_NODE_WORDS if w in ctx["body_text"]]
    if not _v_imgs and not _v_words:
        results.append("+5：全文已无竖向排列的技术路线流程图")
    else:
        _d = []
        if _v_imgs:
            _d.append("残留竖向图片 %.2f×%.2fcm（比值 %.2f）" % _v_imgs[0])
        if _v_words:
            _d.append("残留流程节点文字「%s」" % "、".join(_v_words[:3]))
        results.append("+0：竖向技术路线流程图未删除（%s）" % "；".join(_d))

    # ---------- 19 +5 全文无横向排列的技术路线流程图 ----------
    # ① 不存在宽高比 ≥4.50 的横长条图片（原横向流程图约 16.58×2.67cm，比值 6.21）
    _h_imgs = [(w, h, ra) for _p, w, h, ra, _fl in ctx["images"] if ra >= FLOW_H_RATIO_MIN]
    if not _h_imgs:
        results.append("+5：全文已无横向排列的技术路线流程图")
    else:
        results.append("+0：横向技术路线流程图未删除（残留 %.2f×%.2fcm，比值 %.2f）" % _h_imgs[0])

    # ---------- 20 +3 两处流程图引导文字同步删除，无孤立标签 ----------
    # ① 全文不再出现「技术路线图」标签与「研究技术路线如下」引导句
    _leads = [w for w in FLOW_LEAD_WORDS if w in ctx["body_text"]]
    if not _leads:
        results.append("+3：两处流程图的引导文字已同步删除，无孤立标签残留")
    else:
        results.append("+0：残留流程图引导文字「%s」" % "、".join(_leads))

    # ---------- 21 +3 保留「图1 播撒机构示意结构」及图题，均居中且图题在下 ----------
    # ① 示意图图片仍在文中，图题文字「图1 播撒机构示意结构」保留
    # ② 图片段落与图题段落均居中，且图题位于图片下方（同段换行或紧随其后的段落）
    _fig1 = _fig1_image(ctx)
    _cap = ctx["caption_para"]
    _miss21 = []
    if _fig1 is None:
        _miss21.append("未找到「图1」示意图图片")
    if _cap is None:
        _miss21.append("未找到图题「图1 播撒机构示意结构」")
    if _fig1 is not None and _cap is not None:
        _img_para = _fig1[0]
        if _eff_jc(doc, _img_para) != "center":
            _miss21.append("图片段落未居中（%s）" % _eff_jc(doc, _img_para))
        if _eff_jc(doc, _cap) != "center":
            _miss21.append("图题段落未居中（%s）" % _eff_jc(doc, _cap))
        _paras = ctx["all_paras"]
        _ids = [id(p._p) for p in _paras]
        if id(_img_para._p) == id(_cap._p):
            _pos_img = _pos_cap = None
            for _i, _run in enumerate(_cap.runs):
                if _pos_img is None and _run._r.findall(".//" + qn("w:drawing")):
                    _pos_img = _i
                if _pos_cap is None and FIG1_CAPTION_KEY in (_run.text or ""):
                    _pos_cap = _i
            if _pos_img is None or _pos_cap is None or _pos_img > _pos_cap:
                _miss21.append("图题排在图片上方")
        else:
            try:
                if _ids.index(id(_cap._p)) < _ids.index(id(_img_para._p)):
                    _miss21.append("图题排在图片上方")
            except ValueError:
                _miss21.append("无法定位图片与图题的先后顺序")
    if not _miss21:
        results.append("+3：「图1 播撒机构示意结构」图片与图题均保留且居中，图题位于图片下方")
    else:
        results.append("+0：图1 与图题不达标（%s）" % "；".join(dict.fromkeys(_miss21))[:160])

    # ---------- 22 -1 封面主标题或四、五章标题措辞被改写 ----------
    # ① 封面主标题须为「毕业设计开题报告」
    # ② 第四、五章标题须为「关键参数与实施计划」「主要参考资料」
    _bad22 = []
    _cov_title = ctx["cover_title"]
    if _cov_title is not None and _squeeze(_cov_title.text) != "毕业设计开题报告":
        _bad22.append("主标题为「%s」" % _squeeze(_cov_title.text)[:20])
    _want_ch = {3: "关键参数与实施计划", 4: "主要参考资料"}
    for _i, _want in _want_ch.items():
        if _i < len(_chs):
            _body = _chapter_body(_chs[_i])
            if _body != _want:
                _bad22.append("第 %d 章标题为「%s」" % (_i + 1, _body[:16]))
    if not _bad22:
        results.append("-0：封面主标题与四、五章标题措辞未被改写")
    else:
        results.append("-1：标题措辞被改写（%s）" % "；".join(_bad22)[:160])

    # ---------- 23 -3 意见签署栏出现与本课题无关的评语 ----------
    # ① 意见签署表中不得出现非遗、非物质文化遗产、新媒体、文化传承等无关内容
    _opi_text = _squeeze(_table_text(_topi)) if _topi is not None else ""
    _hit23 = [w for w in IRRELEVANT_WORDS if w in _opi_text]
    if not _hit23:
        results.append("-0：意见签署栏无与本课题无关的评语文字")
    else:
        results.append("-3：意见签署栏出现无关评语（含「%s」等非本课题内容）"
                       % "、".join(_hit23[:3]))

    # ---------- 24 -3 「图1」图片被压缩/拉伸，宽高比偏离约 1.06:1 ----------
    if _fig1 is None:
        results.append("-0：未找到「图1」示意图，不判定压缩拉伸")
    elif abs(_fig1[3] - FIG1_RATIO) <= FIG1_RATIO_TOL:
        results.append("-0：「图1」宽高比 %.3f，保持约 1.06:1 等比显示" % _fig1[3])
    else:
        results.append("-3：「图1」被压缩或拉伸（显示 %.2f×%.2fcm，宽高比 %.3f，应约 1.06:1）"
                       % (_fig1[1], _fig1[2], _fig1[3]))

    # ---------- 25 -1 图题「图1 播撒机构示意结构」不是黑体五号 ----------
    _bad25 = []
    if _cap is None:
        results.append("-0：未找到图题，不判定图题字体")
    else:
        for _r in _real_runs(_cap):
            _fnt = _resolve_effective_eastasia(doc, _cap, _r)
            if _has_cjk(_r.text or "") and not _is_heiti(_fnt):
                _bad25.append("字体为 %s" % _fnt)
            _sz = _run_effective_size_halfpt(doc, _cap, _r)
            if _sz != 21:
                _bad25.append("字号为 %s" % ("缺失" if _sz is None else "%.1f 磅" % (_sz / 2.0)))
        if not _bad25:
            results.append("-0：图题为黑体五号")
        else:
            results.append("-1：图题不是黑体五号（%s）" % "；".join(dict.fromkeys(_bad25))[:120])

    # ---------- 26 -3 出现仅含报告名称行、其余版面全部空白的整页 ----------
    # ① 渲染后存在某一整页，压缩文字不超过 16 字且整页仅为报告名称行
    if not pages:
        results.append("-0：渲染不可用，未取得空白整页证据，不扣分（%s）" % render_note)
    else:
        _blank_pages = [i + 1 for i, t in enumerate(pages)
                        if len(t) <= 16 and NAME_ROW_RE.match(t)]
        if not _blank_pages:
            results.append("-0：无仅含报告名称行的空白整页")
        else:
            results.append("-3：第 %s 页仅含报告名称行「%s」，其余版面全部空白"
                           % ("、".join(str(n) for n in _blank_pages),
                              pages[_blank_pages[0] - 1]))

    # ---------- 27 -1 首个表格框被改为浮动/绝对定位表格或竖向框线被取消 ----------
    # ① 承载「研究目的及意义」「行业研究现状」的表格框须为水平居中的普通表格
    # ② 不得设置文字环绕/绝对定位（tblpPr）；已浮动时 jc 失效，只报浮动这一根因
    # ③ 竖向内框线仅在该表为多列时才有意义，单列表格不判 insideV
    _t1 = ctx["t_main1"]
    _bad27 = []
    if _t1 is None:
        results.append("-0：未找到承载「研究目的及意义」的表格框，不判定表格定位")
    else:
        _fpr = _tbl_float_pr(_t1)
        if _fpr is not None:
            _bad27.append("被设为文字环绕/绝对定位表格（tblpX=%s，tblpY=%s）"
                          % (_fpr.get(qn("w:tblpX")), _fpr.get(qn("w:tblpY"))))
        elif _tbl_jc(_t1) != "center":
            _bad27.append("表格未水平居中（%s）" % (_tbl_jc(_t1) or "缺失"))
        _iv_val, _iv_sz = _get_border(_tbl_borders(_t1), "insideV")
        _iv_other = any(_has_line(_tbl_borders(_t), "insideV")
                        for _t in ctx["tables"]
                        if _t is not _t1 and _tbl_grid_cols(_t) >= 2)
        if (_tbl_grid_cols(_t1) >= 2 and _iv_other
                and _iv_val is not None and _iv_val in ("none", "nil")):
            _bad27.append("内部竖向框线被取消（insideV=%s）" % _iv_val)
        if not _bad27:
            results.append("-0：首个表格框为水平居中的普通表格，框线设置与其余表格框一致")
        else:
            results.append("-1：首个表格框被破坏性修改（%s）" % "；".join(_bad27)[:160])

    # ---------- 28 -1 删除流程图后遗留的空白段落占一页版面 50% 以上 ----------
    # ① 正文表格框内最长的一段连续空白段落，其累计高度不得超过可用版面高度的 50%
    # ② 取不到页面尺寸时不予扣分（减分需证据）
    _usable_pt = _usable_height_pt(doc)
    _blank_n, _blank_pt = _max_blank_band(doc, ctx["main_cell_paras"])
    if not _usable_pt:
        results.append("-0：未取得页面尺寸，不判定空白段落占比")
    elif _blank_pt <= _usable_pt * 0.5:
        results.append("-0：正文最长连续空白 %d 段（%.0f 磅，占一页版面 %.0f%%），未超过半页"
                       % (_blank_n, _blank_pt, 100.0 * _blank_pt / _usable_pt))
    else:
        results.append("-1：删除流程图后遗留 %d 个连续空白段落（%.0f 磅，占一页版面 %.0f%%）"
                       % (_blank_n, _blank_pt, 100.0 * _blank_pt / _usable_pt))

    # ---------- 29 -1 参考文献条目标点被改为全角逗号、冒号 ----------
    # ① 4 条参考文献中不得出现全角「，」「：」，须为半角「, 」「: 」
    _full_hits = []
    for _p in _refs:
        for _ch in ("：", "，"):
            if _ch in _p.text:
                _full_hits.append("「%s」含全角「%s」" % (_squeeze(_p.text)[:6], _ch))
    if not _full_hits:
        results.append("-0：参考文献条目均使用半角逗号与半角冒号")
    else:
        results.append("-1：参考文献条目标点被改为全角（%s）" % "；".join(dict.fromkeys(_full_hits))[:160])

    # ---------- 30 -3 工作进度表表头行与 4 行阶段数据被分隔到相邻两页 ----------
    # ① 渲染后表头行「阶段 主要任务 时间安排」与 4 个阶段行须落在同一页
    if not pages:
        results.append("-0：渲染不可用，未取得进度表跨页证据，不扣分（%s）" % render_note)
    else:
        _hp = _find_page(pages, "".join(SCHEDULE_HEADER))
        _split = []
        for _stage in SCHEDULE_ROWS:
            _sp = _find_page(pages, "%s-%s" % (_stage[1], _stage[2]))
            if _hp and _sp and _sp != _hp:
                _split.append("「%s」行在第 %d 页" % (_stage[0], _sp))
        if not _hp:
            results.append("-0：渲染文本中未定位到进度表表头行，不扣分")
        elif not _split:
            results.append("-0：进度表表头行与 4 行阶段数据同处第 %d 页" % _hp)
        else:
            results.append("-3：进度表表头行在第 %d 页，%s，表头与阶段数据被分隔到相邻两页"
                           % (_hp, "；".join(_split)))

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

