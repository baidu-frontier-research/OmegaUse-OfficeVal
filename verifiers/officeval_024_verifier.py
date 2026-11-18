#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# pyright: reportUnknownMemberType=false, reportUnknownVariableType=false, reportUnknownArgumentType=false, reportUnknownParameterType=false, reportMissingParameterType=false, reportPrivateUsage=false, reportUnknownLambdaType=false, reportGeneralTypeIssues=false

import os
import re
import sys
import json
import math
import shutil
import zipfile
import tempfile
import statistics
import traceback
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
try:
    _VERIFIER_DIR = os.path.dirname(os.path.abspath(__file__))
    if _VERIFIER_DIR not in sys.path:
        sys.path.insert(0, _VERIFIER_DIR)
    import pymupdf_compat as fitz
except ImportError:
    fitz = None
from docx.oxml.ns import qn

# 说明：与既有接口统一约定，本模块不修改全局 sys.stdout，也不 print 主结果。
# 主结果统一由 evaluate(dir_path) -> dict 返回，入参为"脚本所在目录的路径"，
# 由脚本自己在该目录内定位并打开被评估的 .docx。本文件仅在 __main__ 下作为
# 本地调试入口，把该 dict 序列化为 JSON 打印。

SCRIPT_ID = "024"
# 维度二所有评分项 max_delta 中加分侧之和，用于维度一未通过 / 脚本异常时兜底
MAX_SCORE_TOTAL = 24

# 维度二所有评分项按"结果追加顺序"的 max_delta 列表，供末端把 results 里的
# "+N：..." / "-N：..." 字符串拆解为结构化 dim2_items 时使用。顺序须与
# _evaluate_impl 中 results.append(...) 的顺序严格一致，共 34 项：
# 前 6 项为加分项（正值，累加 = MAX_SCORE_TOTAL），后 28 项为减分项（负值）。
_MAX_DELTAS_ORDER = [
    # ---- 加分项 6 条，满分 24 ----
    5,   # 01 17 张内容表格彼此同宽
    3,   # 02 表格宽度不超过版心 8879 twips
    5,   # 03 表格内容字号：16 张五号、缩略语表小四
    3,   # 04 三线表 1.5 / 0.75 / 1.5 磅且无竖线
    5,   # 05 图1-图8 上下方空白不超过 10 行
    3,   # 06 图1-图8 保持长宽比且至少一张被缩放
    # ---- 减分项 28 条 ----
    -5,  # 07 图片文件 / 内嵌图 / 图题序号缺失
    -5,  # 08 18 张表格行列数不符
    -5,  # 09 三线表上框线或下框线缺失
    -5,  # 10 原创性承诺勾选块缺段
    -5,  # 11 授权声明标题 / 正文 / 签署行缺失
    -5,  # 12 正文、摘要、声明页正文字号非小四 12pt
    -5,  # 13 分节数 / 页码重编异常
    -3,  # 14 标点全角半角
    -3,  # 15 封面信息表被改动
    -3,  # 16 内容表格未整表居中
    -3,  # 17 三线表表头行未整行加粗
    -3,  # 18 出现自撰说明性文字
    -3,  # 19 关键词行 / Key Words 行 / 缩略语表
    -3,  # 20 正文区非空段落总数非 241
    -3,  # 21 参考文献条目数 / 自动编号 / 引文上标
    -3,  # 22 单元格合并或同一表题下两张表
    -3,  # 23 各级标题字号
    -3,  # 24 左对齐段落多于 8 段
    -3,  # 25 行距设置
    -3,  # 26 目录页码序列回退
    -3,  # 27 插图和附表清单制表位 / 句号
    -3,  # 28 页眉页脚内容
    -3,  # 29 空表格或空的标题级段落
    -1,  # 30 字体中英文异名并存
    -1,  # 31 首行缩进段落数
    -1,  # 32 TOC / PAGEREF 域被转为静态文字
    -1,  # 33 标题内空格数
    -1,  # 34 附录 A-E 标题缺失
]
assert sum(d for d in _MAX_DELTAS_ORDER if d > 0) == MAX_SCORE_TOTAL
assert len(_MAX_DELTAS_ORDER) == 34

# ============================================================================
# 细则内固化的绝对基准值：全部为写死的目标状态，判定只依赖"完成文件"本身，
# 不需要原文件、指令文档或论文模版参与比对。
# ============================================================================

TEXT_WIDTH_TWIPS = 8879          # 版心宽度 15.66 cm
TBL_WIDTH_TOL = 20               # 表格同宽允许的取整误差（twips）

# 18 张表格按文档顺序的 行数 x 列数
TABLE_SHAPES = [
    (5, 2), (16, 3), (6, 4), (9, 4), (5, 5), (6, 5), (6, 5), (6, 5), (7, 5),
    (7, 5), (6, 4), (5, 5), (6, 4), (6, 4), (16, 4), (18, 4), (13, 4), (11, 4),
]
ABBR_TABLE_INDEX = 1             # 缩略语表在 doc.tables 中的下标
SZ_WUHAO = 21                    # 五号 10.5pt（w:sz 半磅值）
SZ_XIAOSI = 24                   # 小四 12pt
SZ_SIHAO = 28                    # 四号 14pt

BORDER_SZ_THICK = 12             # 1.5 磅（八分之一磅单位）
BORDER_SZ_THIN = 6               # 0.75 磅
NO_LINE = {None, 'nil', 'none'}

# 图1-图8 的显示长宽比（宽/高）
FIG_RATIOS = [2.294, 2.368, 1.140, 1.647, 1.648, 1.807, 1.042, 1.809]
# 图1-图8 未经缩放时的显示尺寸（英寸，宽 x 高）
FIG_BASE_INCHES = [
    (5.315, 2.317), (5.197, 2.195), (4.133, 3.627), (4.848, 2.943),
    (4.803, 2.915), (5.118, 2.832), (4.753, 4.562), (5.039, 2.785),
]
FIG_BLANK_MAX_LINES = 10.0       # 图片上方/下方连续空白上限（行）

MEDIA_MIN = 9                    # word/media 下图片文件数下限
FIG_COUNT = 8                    # 正文内嵌图张数
BODY_PARA_TOTAL = 241            # 表格外非空段落总数
REF_COUNT = 14                   # 参考文献条目数
PAGEREF_MIN = 24                 # PAGEREF 域数量下限
TOC_FIELD_COUNT = 1              # TOC 域数量
SECT_COUNT = 3                   # sectPr 节数
PGNUM_COUNT = 2                  # pgNumType 元素数
PGNUM_START1_COUNT = 2           # pgNumType start="1" 的处数
SZ24_PARA_MIN = 166              # 可见字号为小四的非空段落数下限
SZ28_PARA_MAX = 2                # 可见字号为四号的非空段落数上限
JC_LEFT_MAX = 8                  # 显式 jc=left 段落数上限
LINE480_MAX = 4                  # line=480/auto（2.0 倍行距）段落数上限
LINE_EXACT_MIN = 24              # lineRule=exact（固定行距）段落数下限
FIRSTLINE_MIN = 104              # 2 字符首行缩进段落数下限
EMPTY_HEADING_MAX = 8            # 空的标题级段落数上限
LIST_TAB_POS = '8400'            # 插图和附表清单制表位位置
LIST_ENTRY_COUNT = 18            # 插图和附表清单条目数

HEADER_ODD = '青岚大学城市发展学院硕士学位论文'
HEADER_EVEN = '基于多源数据的城市社区绿地可达性与居民步行行为研究'
KEYWORDS_LINE = '关键词：城市绿地；可达性；步行行为；健康感知；多源数据'
KEYWORDS_EN_HEAD = 'Key Words:'

# 标题原文与其应有的可见字号（w:sz 半磅值）
TITLE_SIZES = [
    ('基于多源数据的城市社区绿地可达性与居民步行行为研究', 30),
    ('学位论文原创性承诺', 40),
    ('学位论文成果使用授权声明', 40),
    ('目录', 32),
    ('插图和附表清单', 32),
    ('摘要', 32),
    ('缩略语表', 32),
    ('致谢', 32),
    ('作者简介', 32),
    ('1研究内容及方法', 24),
    ('2结果', 24),
    ('3讨论', 24),
    ('4结论', 32),
]
# 标题内空格数须与原样一致的段落（键为去空格后的文本）
TITLE_SPACING = {
    '目录': '目    录',
    '摘要': '摘    要',
    '致谢': '致    谢',
    '缩略语表': '缩  略  语  表',
    '附录': '附录',
}
APPENDIX_TITLES = [
    '附录A问卷核心题项节选',
    '附录B社区步行环境现场观察表',
    '附录C居民问卷量表核心题项',
    '附录D指标计算说明',
    '附录E访谈提纲与编码结果',
]
COVER_LABELS = ['申请人：', '学科门类：', '学科专业：', '研究方向：', '指导教师：']
# 出现即判为自撰说明性文字
SELF_WRITTEN_MARKS = [
    '按学校相关管理规定执行',
    '本页为学位论文提交所需授权文本',
    '本表格由作者自行整理',
    '格式已按模板统一调整',
]
HALF_WIDTH_PUNCT = set(',.;:?!()')
FULL_WIDTH_PUNCT = set('，。；：？！（）、《》【】—…')
# 纯英文段落中允许出现的全角字符（弯撇号、弯引号属排印惯例，不判违规）
FULL_WIDTH_ALLOWED = set('‘’“”')

def _style_chain(doc, style_id, limit=12):
    """从 styles.xml 沿 w:basedOn 逐级向上返回样式元素列表。"""
    chain = []
    try:
        styles_el = doc.styles.element
    except Exception:
        return chain
    cur = style_id
    seen = set()
    while cur and cur not in seen and len(chain) < limit:
        seen.add(cur)
        node = None
        for st in styles_el.iter(qn('w:style')):
            if st.get(qn('w:styleId')) == cur:
                node = st
                break
        if node is None:
            break
        chain.append(node)
        based = node.find(qn('w:basedOn'))
        cur = based.get(qn('w:val')) if based is not None else None
    return chain


def _para_style_id(para):
    pPr = para._p.find(qn('w:pPr'))
    if pPr is None:
        return None
    ps = pPr.find(qn('w:pStyle'))
    return ps.get(qn('w:val')) if ps is not None else None


def _default_para_style_id(doc):
    try:
        styles_el = doc.styles.element
    except Exception:
        return None
    for st in styles_el.iter(qn('w:style')):
        if st.get(qn('w:type')) == 'paragraph' and st.get(qn('w:default')) == '1':
            return st.get(qn('w:styleId'))
    return None


def _doc_defaults_rpr(doc):
    try:
        styles_el = doc.styles.element
    except Exception:
        return None
    dd = styles_el.find(qn('w:docDefaults'))
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
    return rf.get(qn('w:' + attr))


def _resolve_effective_eastasia(doc, run):
    """解析 run 实际生效的中文字体名：run 直接设置 -> 段落样式链 -> docDefaults。"""
    val = _rfonts_attr(run._r.find(qn('w:rPr')), 'eastAsia')
    if val:
        return val
    para = None
    try:
        para = run._parent
    except Exception:
        para = None
    if para is not None and hasattr(para, '_p'):
        pPr = para._p.find(qn('w:pPr'))
        if pPr is not None:
            val = _rfonts_attr(pPr.find(qn('w:rPr')), 'eastAsia')
            if val:
                return val
        sid = _para_style_id(para) or _default_para_style_id(doc)
        for st in _style_chain(doc, sid):
            val = _rfonts_attr(st.find(qn('w:rPr')), 'eastAsia')
            if val:
                return val
    return _rfonts_attr(_doc_defaults_rpr(doc), 'eastAsia')


def _resolve_effective_ascii(doc, run):
    """解析 run 实际生效的西文字体名：run 直接设置 -> 段落样式链 -> docDefaults。"""
    val = _rfonts_attr(run._r.find(qn('w:rPr')), 'ascii')
    if val:
        return val
    para = None
    try:
        para = run._parent
    except Exception:
        para = None
    if para is not None and hasattr(para, '_p'):
        pPr = para._p.find(qn('w:pPr'))
        if pPr is not None:
            val = _rfonts_attr(pPr.find(qn('w:rPr')), 'ascii')
            if val:
                return val
        sid = _para_style_id(para) or _default_para_style_id(doc)
        for st in _style_chain(doc, sid):
            val = _rfonts_attr(st.find(qn('w:rPr')), 'ascii')
            if val:
                return val
    return _rfonts_attr(_doc_defaults_rpr(doc), 'ascii')

def _sz_of(rpr):
    if rpr is None:
        return None
    e = rpr.find(qn('w:sz'))
    if e is None:
        return None
    v = e.get(qn('w:val'))
    try:
        return int(float(v))
    except (TypeError, ValueError):
        return None


def _resolve_effective_size(doc, para, run=None):
    """解析实际生效的字号（w:sz 半磅值）。

    run 为 None 时返回"段落可见字号"：取第一个文本非空 run 的生效字号。
    空白 run 常残留旧的 w:sz，必须跳过，否则会读到失效值。
    """
    runs = [run] if run is not None else [
        r for r in para.runs if (r.text or '').strip()
    ]
    for r in runs:
        v = _sz_of(r._r.find(qn('w:rPr')))
        if v:
            return v
    pPr = para._p.find(qn('w:pPr'))
    if pPr is not None:
        # 注意：pPr/rPr 只作用于段落标记，但在无 run 级设置时它是最近的可继承来源
        v = _sz_of(pPr.find(qn('w:rPr')))
        if v:
            return v
    sid = _para_style_id(para) or _default_para_style_id(doc)
    for st in _style_chain(doc, sid):
        v = _sz_of(st.find(qn('w:rPr')))
        if v:
            return v
    return _sz_of(_doc_defaults_rpr(doc))


def _open_word_document(file_path):
    return Document(file_path)

def _render_to_pdf(file_path, tmp_dir):
    """用本机排版引擎把 docx 渲染为 PDF，返回 PDF 路径；失败抛 RuntimeError。

    优先 Microsoft Word（Word.Application），本机若只装了 WPS Office 则退回
    KWPS.Application（WPS Writer 暴露的 Word 兼容 COM 接口）。图片上下方空白
    必须按真实分页结果测量，空段落计数抓不到分页挤压产生的空白。
    """
    try:
        import win32com.client  # type: ignore
    except ImportError as exc:
        raise RuntimeError('无法导入 win32com.client，无法渲染 PDF：%s' % exc)

    src = os.path.abspath(file_path)
    dst = os.path.join(os.path.abspath(tmp_dir), 'render.pdf')
    last_err = None
    for prog_id in ('Word.Application', 'KWPS.Application'):
        app = None
        try:
            app = win32com.client.DispatchEx(prog_id)
            try:
                app.Visible = False
            except Exception:
                pass
            try:
                app.DisplayAlerts = 0
            except Exception:
                pass
            wdoc = app.Documents.Open(src, ReadOnly=True, AddToRecentFiles=False)
            wdoc.ExportAsFixedFormat(OutputFileName=dst, ExportFormat=17)
            wdoc.Close(0)
            if os.path.isfile(dst) and os.path.getsize(dst) > 0:
                return dst
            last_err = '%s 导出后未生成有效 PDF' % prog_id
        except Exception as exc:
            last_err = '%s: %s' % (prog_id, exc)
        finally:
            if app is not None:
                try:
                    app.Quit()
                except Exception:
                    pass
    raise RuntimeError('渲染 PDF 失败：%s' % last_err)

def _measure_figure_blank_lines(file_path, doc):
    """按真实分页结果测量图1-图8 上方/下方的连续空白行数。

    返回 (max_lines, detail)；无法测量时返回 (None, 原因)。
    上方空白 = 图片位于页首时前一页页尾剩余空白 + 本页页首空白，否则为与上一个
    版面元素的垂直间距；下方同理。点数除以该文件正文行距中位数折算为行。
    """
    if fitz is None:
        return None, 'PDF 兼容后端未安装'
    tmp_dir = tempfile.mkdtemp(prefix='officeval_%s_' % SCRIPT_ID)
    try:
        pdf_path = _render_to_pdf(file_path, tmp_dir)
        sec = doc.sections[0]
        band_top = sec.top_margin.pt
        band_bottom = sec.page_height.pt - sec.bottom_margin.pt
        pdf = fitz.open(pdf_path)
        pages = []
        line_gaps = []
        for page in pdf:
            items = []
            for blk in page.get_text('dict').get('blocks', []):
                x0, y0, x1, y1 = blk.get('bbox', (0, 0, 0, 0))
                if y1 < band_top - 1 or y0 > band_bottom + 1:
                    continue  # 落在页眉页脚带内，不参与版面空白计算
                kind = 'image' if blk.get('type') == 1 else 'text'
                text = ''
                if kind == 'text':
                    lines = blk.get('lines', [])
                    text = ' '.join(
                        ''.join(sp.get('text', '') for sp in ln.get('spans', []))
                        for ln in lines
                    )
                    tops = [ln.get('bbox', (0, 0, 0, 0))[1] for ln in lines]
                    for a, b in zip(tops, tops[1:]):
                        if 5.0 < b - a < 60.0:
                            line_gaps.append(b - a)
                items.append({'kind': kind, 'y0': y0, 'y1': y1, 'text': text})
            items.sort(key=lambda it: it['y0'])
            pages.append(items)
        pdf.close()
        if not line_gaps:
            return None, 'PDF 中未取到可用行距'
        line_h = statistics.median(line_gaps)
        # 展平为全局顺序，便于按"图 N"题注回溯到它对应的那张图
        seq = []
        for pi, items in enumerate(pages):
            for j, it in enumerate(items):
                rec = dict(it)
                rec['page'] = pi
                rec['pos'] = j
                rec['last'] = (j == len(items) - 1)
                rec['first'] = (j == 0)
                seq.append(rec)

        def _prev_page_tail(pi):
            for k in range(pi - 1, -1, -1):
                if pages[k]:
                    return band_bottom - pages[k][-1]['y1']
            return 0.0

        def _next_page_head(pi):
            for k in range(pi + 1, len(pages)):
                if pages[k]:
                    return pages[k][0]['y0'] - band_top
            return 0.0

        def _gaps(gi):
            rec = seq[gi]
            pi = rec['page']
            items = pages[pi]
            if rec['first']:
                above = (rec['y0'] - band_top) + _prev_page_tail(pi)
            else:
                above = rec['y0'] - items[rec['pos'] - 1]['y1']
            if rec['last']:
                below = (band_bottom - rec['y1']) + _next_page_head(pi)
            else:
                below = items[rec['pos'] + 1]['y0'] - rec['y1']
            return max(above, 0.0), max(below, 0.0)

        worst = 0.0
        found = 0
        for n in range(1, FIG_COUNT + 1):
            cap = None
            pat = re.compile(r'^\s*图\s*%d(?![0-9])' % n)
            for gi, rec in enumerate(seq):
                if rec['kind'] == 'text' and pat.match(rec['text']):
                    cap = gi  # 取最后一处：靠前的同名文字属于"插图和附表清单"
            if cap is None:
                continue
            img = None
            for gi in range(cap - 1, -1, -1):
                if seq[gi]['kind'] == 'image':
                    img = gi
                    break
            if img is None:
                continue
            found += 1
            above, below = _gaps(img)
            worst = max(worst, above / line_h, below / line_h)
        if found < FIG_COUNT:
            return None, '仅定位到 %d/%d 张图' % (found, FIG_COUNT)
        return worst, '最大空白 %.1f 行（行距中位数 %.2fpt，共 %d 页）' % (
            worst, line_h, len(pages))
    except Exception as exc:
        return None, '%s: %s' % (type(exc).__name__, exc)
    finally:
        shutil.rmtree(tmp_dir, ignore_errors=True)

def _check_dim1(file_path):
    """维度一：可用或可修改性。只校验文件本身是否正常，不检查文件内部排版。

    返回 (pass, reason)；reason 只在不通过时给出，通过时为空串。两条门槛：
    1 结果文件为 .docx，且实际格式与扩展名一致；
    2 文件完整可用——能正常打开且无修复提示，包内含 word/document.xml，内容仍
      可编辑，未加密码或只读保护、未转为图片、未拆分。
    """
    if not file_path.lower().endswith('.docx'):
        return False, '门槛1不满足：结果文件扩展名不是 .docx'
    if not os.path.isfile(file_path) or os.path.getsize(file_path) == 0:
        return False, '门槛1不满足：文件不存在或为 0 字节'
    try:
        with open(file_path, 'rb') as fh:
            if fh.read(2) != b'PK':
                return False, '门槛1不满足：实际格式不是 OOXML 包（非 zip），与 .docx 扩展名不符'
    except Exception as exc:
        return False, '门槛1不满足：文件无法读取（%s）' % exc
    try:
        with zipfile.ZipFile(file_path) as zf:
            if zf.testzip() is not None:
                return False, '门槛2不满足：包内存在损坏项，解压会报错'
            names = zf.namelist()
            if 'word/document.xml' not in names:
                return False, '门槛2不满足：包内缺少 word/document.xml'
            if '[Content_Types].xml' not in names:
                return False, '门槛1不满足：包内缺少 [Content_Types].xml，不是有效的 docx'
            ct = zf.read('[Content_Types].xml').decode('utf-8', 'ignore')
            if 'wordprocessingml.document.main+xml' not in ct:
                return False, '门槛1不满足：包内主文档部件类型不是 WordprocessingML，与 .docx 扩展名不符'
            settings = ''
            if 'word/settings.xml' in names:
                settings = zf.read('word/settings.xml').decode('utf-8', 'ignore')
    except zipfile.BadZipFile:
        return False, '门槛1不满足：文件不是有效的 zip 包，无法作为 docx 打开'
    except Exception as exc:
        return False, '门槛2不满足：文件包无法正常解压（%s）' % exc
    if 'w:writeProtection' in settings:
        return False, '门槛2不满足：文档设置了写保护，内容不可编辑'
    m = re.search(r'<w:documentProtection[^>]*>', settings)
    if m and re.search(r'w:edit="(readOnly|forms|comments|trackedChanges)"', m.group(0)):
        return False, '门槛2不满足：文档设置了编辑保护，内容不可编辑'
    try:
        doc = Document(file_path)
    except Exception as exc:
        return False, '门槛2不满足：文档无法正常打开（%s: %s）' % (type(exc).__name__, exc)
    body_text = sum(1 for p in doc.paragraphs if (p.text or '').strip())
    if body_text == 0 and not doc.tables:
        return False, '门槛2不满足：文档内无任何可编辑文字，疑似内容已转为图片'
    return True, ''

def evaluate(dir_path):
    """主入口：入参为脚本所在目录，自行在该目录内定位被评估的 .docx。"""
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
        cands = sorted(
            f for f in os.listdir(dir_path)
            if f.lower().endswith('.docx') and not f.startswith('~')
        )
        if not cands:
            raise FileNotFoundError('目录内未找到 .docx 文件：%s' % dir_path)
        file_name = cands[0]
        result["file_name"] = file_name
        return _evaluate_impl(os.path.join(dir_path, file_name), result)
    except Exception as exc:
        # status 为 error 时未能完成判定，dim1_pass 留空，dim1_reason 也为空
        result["status"] = "error"
        result["error"] = '%s: %s' % (type(exc).__name__, exc)
        result["dim1_pass"] = ""
        result["dim1_reason"] = ""
        result["dim2_items"] = []
        result["total_score"] = 0
        result["max_score"] = MAX_SCORE_TOTAL
        return result

def _edge(el, name):
    """从 tblBorders / tcBorders 元素上取某条边的 (val, sz)。"""
    if el is None:
        return None, None
    e = el.find(qn('w:' + name))
    if e is None:
        return None, None
    sz = e.get(qn('w:sz'))
    try:
        sz = int(float(sz))
    except (TypeError, ValueError):
        sz = None
    return e.get(qn('w:val')), sz


def _tbl_borders(tbl):
    tblPr = tbl._tbl.find(qn('w:tblPr'))
    return tblPr.find(qn('w:tblBorders')) if tblPr is not None else None


def _tc_borders(cell):
    tcPr = cell._tc.find(qn('w:tcPr'))
    return tcPr.find(qn('w:tcBorders')) if tcPr is not None else None


def _cell_edge(tbl, cell, name, fallback):
    """单元格某条边的生效值：tcBorders 覆盖 tblBorders。"""
    val, sz = _edge(_tc_borders(cell), name)
    if val is None:
        val, sz = _edge(_tbl_borders(tbl), fallback)
    return val, sz


def _visible(val, sz):
    return val not in NO_LINE and (sz is None or sz > 0)


def _is_single(val, sz, lo, hi):
    return val == 'single' and sz is not None and lo <= sz <= hi


def _tbl_width_twips(tbl):
    """表格声明宽度，换算为 twips；pct 按版心折算；auto/nil 返回 None。"""
    tblPr = tbl._tbl.find(qn('w:tblPr'))
    if tblPr is None:
        return None
    tblW = tblPr.find(qn('w:tblW'))
    if tblW is None:
        return None
    raw = tblW.get(qn('w:w'))
    typ = tblW.get(qn('w:type'))
    try:
        val = float(raw)
    except (TypeError, ValueError):
        return None
    if typ == 'dxa':
        return val
    if typ == 'pct':
        return val / 5000.0 * TEXT_WIDTH_TWIPS
    return None

def _three_line_report(tbl):
    """返回三线表结构检查结果 dict：上线 / 表头下线 / 下线 / 竖线。"""
    rows = tbl.rows
    n = len(rows)
    rep = {'top_ok': False, 'mid_ok': False, 'bottom_ok': False,
           'no_vertical': True, 'no_extra_h': True,
           'top_visible': False, 'bottom_visible': False}
    if n < 2:
        return rep
    top = [_cell_edge(tbl, c, 'top', 'top') for c in rows[0].cells]
    rep['top_visible'] = any(_visible(v, s) for v, s in top)
    rep['top_ok'] = bool(top) and all(
        _is_single(v, s, BORDER_SZ_THICK - 1, BORDER_SZ_THICK + 1) for v, s in top)
    bot = [_cell_edge(tbl, c, 'bottom', 'bottom') for c in rows[n - 1].cells]
    rep['bottom_visible'] = any(_visible(v, s) for v, s in bot)
    rep['bottom_ok'] = bool(bot) and all(
        _is_single(v, s, BORDER_SZ_THICK - 1, BORDER_SZ_THICK + 1) for v, s in bot)

    def boundary(ri):
        pairs = []
        upper = rows[ri].cells
        lower = rows[ri + 1].cells
        width = max(len(upper), len(lower))
        for ci in range(width):
            vis, best = False, None
            if ci < len(upper):
                v, s = _cell_edge(tbl, upper[ci], 'bottom',
                                  'bottom' if ri == n - 1 else 'insideH')
                if _visible(v, s):
                    vis, best = True, s
            if ci < len(lower):
                v, s = _cell_edge(tbl, lower[ci], 'top',
                                  'top' if ri + 1 == 0 else 'insideH')
                if _visible(v, s):
                    vis = True
                    best = s if best is None else max(best, s or 0)
            pairs.append((vis, best))
        return pairs

    mid = boundary(0)
    rep['mid_ok'] = bool(mid) and all(
        vis and s is not None and BORDER_SZ_THIN - 1 <= s <= BORDER_SZ_THIN + 1
        for vis, s in mid)
    for ri in range(1, n - 1):
        if any(vis for vis, _ in boundary(ri)):
            rep['no_extra_h'] = False
            break
    for row in rows:
        cells = row.cells
        for ci, cell in enumerate(cells):
            lv, ls = _cell_edge(tbl, cell, 'left', 'left' if ci == 0 else 'insideV')
            rv, rs = _cell_edge(tbl, cell, 'right',
                                'right' if ci == len(cells) - 1 else 'insideV')
            if _visible(lv, ls) or _visible(rv, rs):
                rep['no_vertical'] = False
                break
        if not rep['no_vertical']:
            break
    return rep

def _evaluate_impl(file_path, result):
    # ===== 维度一：只校验文件是否正常，不检查文件内部 =====
    dim1_ok, dim1_reason = _check_dim1(file_path)
    result["dim1_pass"] = dim1_ok
    result["dim1_reason"] = dim1_reason
    if not dim1_ok:
        result["dim2_items"] = []
        result["total_score"] = 0
        result["max_score"] = MAX_SCORE_TOTAL
        return result

    doc = _open_word_document(file_path)
    with zipfile.ZipFile(file_path) as zf:
        names = zf.namelist()
        dx = zf.read('word/document.xml').decode('utf-8', 'ignore')
        st = (zf.read('word/settings.xml').decode('utf-8', 'ignore')
              if 'word/settings.xml' in names else '')
        media = [n for n in names
                 if n.startswith('word/media/') and n.count('/') == 2]
        hdr_txt, ftr_xml = [], []
        for n in names:
            if re.match(r'word/header\d*\.xml$', n):
                raw = zf.read(n).decode('utf-8', 'ignore')
                hdr_txt.append(re.sub(r'<[^>]+>', '', raw).strip())
            elif re.match(r'word/footer\d*\.xml$', n):
                ftr_xml.append(zf.read(n).decode('utf-8', 'ignore'))

    paras = doc.paragraphs
    tables = doc.tables
    content_tables = tables[1:] if len(tables) > 1 else []
    nonempty = [p for p in paras if (p.text or '').strip()]

    def norm(s):
        return re.sub(r'\s+', '', s or '')

    def has_field(p):
        return (next(p._p.iter(qn('w:instrText')), None) is not None
                or next(p._p.iter(qn('w:fldChar')), None) is not None)

    plain = [p for p in nonempty if not has_field(p)]
    plain_norm = {}
    for p in plain:
        plain_norm.setdefault(norm(p.text), p)
    all_norm_text = set(plain_norm.keys())
    full_text = '\n'.join(p.text for p in paras)

    def pPr(p):
        return p._p.find(qn('w:pPr'))

    def jc_of(p):
        e = pPr(p)
        if e is None:
            return None
        j = e.find(qn('w:jc'))
        return j.get(qn('w:val')) if j is not None else None

    results = []

    # ========================= 加分项 6 条，满分 24 =========================

    # 01 +5 除封面信息表外的 17 张内容表格宽度彼此相同（允许 ±20 twips 取整误差）
    rule = '除封面信息表外的 17 张内容表格宽度彼此相同'
    widths = [_tbl_width_twips(t) for t in content_tables]
    same_width = (len(content_tables) == 17 and all(w is not None for w in widths)
                  and max(widths) - min(widths) <= TBL_WIDTH_TOL)
    results.append(('+5：' if same_width else '+0：') + rule)

    # 02 +3 统一后的表格宽度不超过版心 8879 twips
    rule = '17 张内容表格的宽度均不超过版心宽度 15.66 cm'
    in_width = (len(content_tables) == 17 and all(w is not None for w in widths)
                and max(widths) <= TEXT_WIDTH_TWIPS + TBL_WIDTH_TOL)
    results.append(('+3：' if in_width else '+0：') + rule)

    # 03 +5 表格内容字号：16 张五号，缩略语表小四
    rule = '缩略语表单元格文字字号为小四 12pt，其余 16 张内容表格单元格文字字号为五号 10.5pt'
    size_ok = len(tables) == 18
    if size_ok:
        for ti, tbl in enumerate(tables):
            if ti == 0:
                continue
            want = SZ_XIAOSI if ti == ABBR_TABLE_INDEX else SZ_WUHAO
            for row in tbl.rows:
                for cell in row.cells:
                    for p in cell.paragraphs:
                        if not (p.text or '').strip():
                            continue
                        if _resolve_effective_size(doc, p) != want:
                            size_ok = False
                            break
                    if not size_ok:
                        break
                if not size_ok:
                    break
            if not size_ok:
                break
    results.append(('+5：' if size_ok else '+0：') + rule)

    # 04 +3 三线表线宽：上线 1.5 磅、表头下线 0.75 磅、下线 1.5 磅，无竖线无其他横线
    rule = '17 张内容表格均为标准三线表：上框线 1.5 磅、表头行下框线 0.75 磅、末行下框线 1.5 磅，且无竖线、无其他横线'
    reports = [_three_line_report(t) for t in content_tables]
    three_ok = (len(content_tables) == 17 and all(
        r['top_ok'] and r['mid_ok'] and r['bottom_ok']
        and r['no_vertical'] and r['no_extra_h'] for r in reports))
    results.append(('+3：' if three_ok else '+0：') + rule)

    # 05 +5 图1-图8 上方与下方的连续空白均不超过 10 行
    rule = '图1-图8 每张图的上方与下方的连续空白均不超过 10 行'
    blank_worst, blank_detail = _measure_figure_blank_lines(file_path, doc)
    blank_ok = (blank_worst is not None and blank_worst <= FIG_BLANK_MAX_LINES)
    results.append(('+5：' if blank_ok else '+0：') + rule)

    # 06 +3 图1-图8 保持长宽比，且至少一张确实被缩放过
    rule = '图1-图8 均保持原图长宽比等比缩放，且至少一张图片的显示尺寸相对原始尺寸有过缩放'
    # 用"图 N"题注回溯定位其上方最近的一张内嵌图，避免把封面图标计入
    draw_paras = []
    for pi, p in enumerate(paras):
        for ext in p._p.iter(qn('wp:extent')):
            try:
                cx, cy = int(ext.get('cx')), int(ext.get('cy'))
            except (TypeError, ValueError):
                continue
            if cx > 0 and cy > 0:
                draw_paras.append((pi, cx / 914400.0, cy / 914400.0))
    cap_idx = {}
    for pi, p in enumerate(paras):
        m = re.match(r'^\s*图\s*(\d+)(?![0-9])', p.text or '')
        if m:
            n = int(m.group(1))
            # 取最后一处匹配：靠前的同名段落属于"插图和附表清单"，不是正文图题
            if 1 <= n <= FIG_COUNT and not has_field(p):
                cap_idx[n] = pi
    figs = {}
    for n, ci in cap_idx.items():
        cand = [d for d in draw_paras if d[0] < ci]
        if cand:
            figs[n] = cand[-1]
    ratio_ok = len(figs) == FIG_COUNT
    scaled = False
    if ratio_ok:
        for n in range(1, FIG_COUNT + 1):
            _, wi, hi = figs[n]
            want = FIG_RATIOS[n - 1]
            if hi <= 0 or abs(wi / hi - want) / want > 0.01:
                ratio_ok = False
                break
            bw, bh = FIG_BASE_INCHES[n - 1]
            if abs(wi - bw) > 0.01 or abs(hi - bh) > 0.01:
                scaled = True
    results.append(('+3：' if (ratio_ok and scaled) else '+0：') + rule)

    # ========================= 减分项 28 条 =========================

    # 07 -5 图片文件 / 内嵌图 / 图题序号缺失
    rule = '文档内图片不足 8 张，或图题「图1」至「图8」有缺号、重号，或与图片先后顺序错乱'
    seq_ok = True
    order = [cap_idx[n] for n in sorted(cap_idx)] if cap_idx else []
    if order != sorted(order):
        seq_ok = False
    bad = (len(media) < MEDIA_MIN or len(draw_paras) < FIG_COUNT
           or len(cap_idx) != FIG_COUNT or not seq_ok)
    results.append(('-5：' if bad else '+0：') + rule)

    # 08 -5 18 张表格行列数不符
    rule = '文档内表格不足 18 张，或按顺序的行×列不为 5×2、16×3、6×4、9×4、5×5、6×5、6×5、6×5、7×5、7×5、6×4、5×5、6×4、6×4、16×4、18×4、13×4、11×4'
    shapes = [(len(t.rows), len(t.columns)) for t in tables]
    results.append(('-5：' if shapes != TABLE_SHAPES else '+0：') + rule)

    # 09 -5 三线表上框线或下框线缺失
    rule = '17 张内容表格中有任一张的上框线或末行下框线缺失，成为缺线的两线表'
    miss_edge = (len(content_tables) != 17 or any(
        (not r['top_visible']) or (not r['bottom_visible']) for r in reports))
    results.append(('-5：' if miss_edge else '+0：') + rule)

    # 10 -5 原创性承诺勾选块缺段
    rule = '缺少「保密□，在年解密后适用本授权书。」「本学位论文属于不保密□。」「（请在以上方框内打"√"）」三段中的任意一段'
    ck1 = any('保密□' in (p.text or '') for p in paras)
    ck2 = any('不保密□' in (p.text or '') for p in paras)
    ck3 = any('请在以上方框内打' in (p.text or '') and '√' in (p.text or '')
              for p in paras)
    results.append(('-5：' if not (ck1 and ck2 and ck3) else '+0：') + rule)

    # 11 -5 授权声明标题 / 正文 / 签署行缺失
    rule = '缺少「学位论文原创性承诺」或「学位论文成果使用授权声明」标题段，或缺少其正文段，或缺少「导师签名：」「日期：」签署行'
    d1 = '学位论文原创性承诺' in all_norm_text
    d2 = '学位论文成果使用授权声明' in all_norm_text
    d3 = '导师签名' in norm(full_text)
    d4 = '日期' in norm(full_text)
    d5 = '本人郑重声明' in norm(full_text) or '本人完全了解' in norm(full_text)
    results.append(('-5：' if not (d1 and d2 and d3 and d4 and d5) else '+0：') + rule)

    # 12 -5 正文、摘要、声明页正文字号非小四 12pt
    rule = '正文、摘要与声明页的正文字号不是小四 12pt'
    cnt24 = sum(1 for p in nonempty if _resolve_effective_size(doc, p) == SZ_XIAOSI)
    cnt28 = sum(1 for p in nonempty if _resolve_effective_size(doc, p) == SZ_SIHAO)
    results.append((
        '-5：' if (cnt24 < SZ24_PARA_MIN or cnt28 > SZ28_PARA_MAX) else '+0：') + rule)

    # 13 -5 分节数 / 页码重编异常
    rule = '分节数或页码编号设置被改动，正文区被二次分节重编页码，导致同一逻辑页码重复出现'
    n_sect = dx.count('<w:sectPr')
    n_pgnum = dx.count('<w:pgNumType')
    n_start1 = len(re.findall(r'<w:pgNumType[^>]*w:start="1"', dx))
    results.append((
        '-5：' if (n_sect != SECT_COUNT or n_pgnum != PGNUM_COUNT
                   or n_start1 != PGNUM_START1_COUNT) else '+0：') + rule)

    # 14 -3 标点全角半角
    rule = '中文段落内出现半角标点，或英文段落内出现全角标点，标点全角半角混用'
    ref_pat = re.compile(r'\[(J|M|D|C|R|N|P|S|EB/OL|J/OL|M/OL)\]')
    hw_pat = re.compile(r'[一-鿿][,.;:?!()]|[,.;:?!()][一-鿿]')
    punct_bad = False
    for p in paras:
        t = p.text or ''
        if not t.strip():
            continue
        if ref_pat.search(t):
            continue
        has_cjk = re.search(r'[一-鿿]', t) is not None
        if has_cjk:
            if hw_pat.search(t):
                punct_bad = True
                break
        elif re.search(r'[A-Za-z]', t):
            if any(ch in FULL_WIDTH_PUNCT and ch not in FULL_WIDTH_ALLOWED
                   for ch in t):
                punct_bad = True
                break
    results.append(('-3：' if punct_bad else '+0：') + rule)

    # 15 -3 封面信息表被改动
    rule = '封面信息表不是 5 行 2 列，或第一列 5 个标签不依次为 申请人：/学科门类：/学科专业：/研究方向：/指导教师：，或被加上可见框线，或被整表居中'
    cover_bad = True
    if tables:
        t0 = tables[0]
        labels = [norm(r.cells[0].text) for r in t0.rows] if t0.rows else []
        rep0 = _three_line_report(t0)
        jc0 = None
        tblPr0 = t0._tbl.find(qn('w:tblPr'))
        if tblPr0 is not None:
            j = tblPr0.find(qn('w:jc'))
            jc0 = j.get(qn('w:val')) if j is not None else None
        cover_bad = not (
            (len(t0.rows), len(t0.columns)) == (5, 2)
            and labels == COVER_LABELS
            and not rep0['top_visible'] and not rep0['bottom_visible']
            and rep0['no_vertical'] and rep0['no_extra_h']
            and jc0 != 'center')
    results.append(('-3：' if cover_bad else '+0：') + rule)

    # 16 -3 内容表格未整表居中
    rule = '17 张内容表格中有任一张未整表居中'
    jc_bad = len(content_tables) != 17
    for tbl in content_tables:
        tblPr = tbl._tbl.find(qn('w:tblPr'))
        j = tblPr.find(qn('w:jc')) if tblPr is not None else None
        if j is None or j.get(qn('w:val')) != 'center':
            jc_bad = True
            break
    results.append(('-3：' if jc_bad else '+0：') + rule)

    # 17 -3 三线表表头行未整行加粗
    rule = '17 张内容表格中有任一张的表头行（首行）存在未加粗的文字'
    bold_bad = len(content_tables) != 17
    for tbl in content_tables:
        if not tbl.rows:
            bold_bad = True
            break
        for cell in tbl.rows[0].cells:
            for p in cell.paragraphs:
                for r in p.runs:
                    if (r.text or '').strip() and not r.font.bold:
                        bold_bad = True
                        break
                if bold_bad:
                    break
            if bold_bad:
                break
        if bold_bad:
            break
    results.append(('-3：' if bold_bad else '+0：') + rule)

    # 18 -3 出现自撰说明性文字
    rule = '文档内出现「按学校相关管理规定执行」「本页为学位论文提交所需授权文本」「本表格由作者自行整理」「格式已按模板统一调整」等自撰说明性文字'
    self_written = any(m in full_text for m in SELF_WRITTEN_MARKS)
    results.append(('-3：' if self_written else '+0：') + rule)

    # 19 -3 关键词行 / Key Words 行 / 缩略语表
    rule = '缺少「关键词：城市绿地；可达性；步行行为；健康感知；多源数据」整行，或缺少以「Key Words:」起首的英文关键词行，或缩略语表不是 16 行 3 列'
    kw = norm(KEYWORDS_LINE) in all_norm_text
    kwen = any((p.text or '').strip().startswith(KEYWORDS_EN_HEAD) for p in paras)
    abbr_ok = (len(tables) > ABBR_TABLE_INDEX
               and (len(tables[ABBR_TABLE_INDEX].rows),
                    len(tables[ABBR_TABLE_INDEX].columns)) == (16, 3))
    results.append(('-3：' if not (kw and kwen and abbr_ok) else '+0：') + rule)

    # 20 -3 正文区非空段落总数非 241
    rule = '正文区非空段落总数不等于 241，存在段落被删除、合并或新增'
    results.append(('-3：' if len(nonempty) != BODY_PARA_TOTAL else '+0：') + rule)

    # 21 -3 参考文献条目数 / 自动编号 / 引文上标
    rule = '参考文献条目不为 14 条，或条目序号改用自动编号与交叉引用域，或方括号引文标注中存在非上标形式'
    ref_entries = [p for p in paras if ref_pat.search(p.text or '')]
    n_numpr = dx.count('<w:numPr')
    ref_fields = len([m for m in re.findall(r'<w:instrText[^>]*>([^<]*)</w:instrText>', dx)
                      if m.strip().startswith('REF ')])
    cite_pat = re.compile(r'^\[\d+(\s*[-,–~]\s*\d+)*\]$')
    cites = []
    for p in paras:
        if ref_pat.search(p.text or ''):
            continue
        for r in p.runs:
            if cite_pat.match((r.text or '').strip()):
                cites.append(r)
    cite_ok = bool(cites) and all(r.font.superscript for r in cites)
    results.append((
        '-3：' if (len(ref_entries) != REF_COUNT or n_numpr > 0 or ref_fields > 0
                   or not cite_ok) else '+0：') + rule)

    # 22 -3 单元格合并或同一表题下两张表
    rule = '文档内出现横向合并或纵向合并的单元格，或同一表题下被拆成两张表'
    merged = ('<w:gridSpan' in dx) or ('<w:vMerge' in dx)
    body = doc.element.body
    split_tbl = False
    prev = None
    for child in body.iterchildren():
        tag = child.tag
        if tag == qn('w:tbl'):
            if prev == 'tbl':
                split_tbl = True
                break
            prev = 'tbl'
        elif tag == qn('w:p'):
            if ''.join(child.itertext()).strip():
                prev = 'p'
    results.append(('-3：' if (merged or split_tbl) else '+0：') + rule)

    # 23 -3 各级标题字号
    rule = '论文题目不为小三 15pt，两个声明页标题不为小二 20pt，目录 / 插图和附表清单 / 摘要 / 缩略语表 / 致谢 / 作者简介 / 4 结论 不为三号 16pt，1 研究内容及方法 / 2 结果 / 3 讨论 不为小四 12pt'
    title_bad = False
    for text, want in TITLE_SIZES:
        p = plain_norm.get(text)
        if p is None:
            title_bad = True
            break
        if _resolve_effective_size(doc, p) != want:
            title_bad = True
            break
    results.append(('-3：' if title_bad else '+0：') + rule)

    # 24 -3 左对齐段落多于 8 段
    rule = '左对齐的非空段落多于 8 段，原本两端对齐的正文段被改为左对齐'
    n_left = sum(1 for p in nonempty if jc_of(p) == 'left')
    results.append(('-3：' if n_left > JC_LEFT_MAX else '+0：') + rule)

    # 25 -3 行距设置
    rule = '2.0 倍行距的段落多于 4 段，或固定行距的段落少于 24 段'
    n480 = 0
    n_exact = 0
    for p in paras:
        e = pPr(p)
        sp = e.find(qn('w:spacing')) if e is not None else None
        if sp is None:
            continue
        line = sp.get(qn('w:line'))
        rulev = sp.get(qn('w:lineRule'))
        if line == '480' and rulev == 'auto':
            n480 += 1
        if rulev == 'exact':
            n_exact += 1
    results.append((
        '-3：' if (n480 > LINE480_MAX or n_exact < LINE_EXACT_MIN) else '+0：') + rule)

    # 26 -3 目录页码序列回退
    rule = '目录条目的页码序列出现 2 次及以上回退，使目录条目失去定位功能'
    toc_pages = []
    for p in paras:
        instr = [t for t in p._p.iter(qn('w:instrText'))]
        if any((t.text or '').strip().startswith('PAGEREF') for t in instr):
            m = re.search(r'(\d+)\s*$', (p.text or '').strip())
            if m:
                toc_pages.append(int(m.group(1)))
    toc_bad = sum(1 for a, b in zip(toc_pages, toc_pages[1:]) if b < a) >= 2
    results.append(('-3：' if toc_bad else '+0：') + rule)

    # 27 -3 插图和附表清单制表位 / 句号
    rule = '插图和附表清单不是 18 条，或有任一条缺少前导点制表位，或条目末尾被加上句号'
    lo, hi = None, len(paras)
    for pi, p in enumerate(paras):
        nt = norm(p.text)
        if nt == '插图和附表清单' and not has_field(p):
            lo = pi
        elif lo is not None and nt in ('摘要', '摘  要') and pi > lo:
            hi = pi
            break
    list_bad = True
    if lo is not None:
        entries = [p for p in paras[lo + 1:hi]
                   if re.match(r'^\s*(图|表)\s*\d', p.text or '')]
        list_bad = len(entries) != LIST_ENTRY_COUNT
        for p in entries:
            e = pPr(p)
            tabs = e.find(qn('w:tabs')) if e is not None else None
            ok = False
            if tabs is not None:
                for tb in tabs.findall(qn('w:tab')):
                    if (tb.get(qn('w:leader')) == 'dot'
                            and tb.get(qn('w:pos')) == LIST_TAB_POS):
                        ok = True
                        break
            if not ok or (p.text or '').rstrip().endswith('。'):
                list_bad = True
                break
    results.append(('-3：' if list_bad else '+0：') + rule)

    # 28 -3 页眉页脚内容
    rule = '页眉中缺少「青岚大学城市发展学院硕士学位论文」或「基于多源数据的城市社区绿地可达性与居民步行行为研究」文字，或页脚中缺少自动页码'
    h_ok = (any(HEADER_ODD in t for t in hdr_txt)
            and any(HEADER_EVEN in t for t in hdr_txt))
    f_ok = any('PAGE' in x for x in ftr_xml)
    results.append(('-3：' if not (h_ok and f_ok) else '+0：') + rule)

    # 29 -3 空表格或空的标题级段落
    rule = '文档内出现一张单元格全无文字的空表格，或空的标题级段落多于 8 个'
    empty_tbl = any(
        not any((c.text or '').strip() for r in t.rows for c in r.cells)
        for t in tables)
    hstyle = re.compile(r'^([1-9]\d*|Heading)')
    n_empty_head = 0
    for p in paras:
        if (p.text or '').strip():
            continue
        sid = _para_style_id(p)
        if sid and hstyle.match(sid):
            n_empty_head += 1
    results.append((
        '-3：' if (empty_tbl or n_empty_head > EMPTY_HEADING_MAX) else '+0：') + rule)

    # 30 -1 字体中英文异名并存
    rule = '同一字体的中文名与英文名在文件内并存（宋体与 SimSun / NSimSun 混写，或黑体与 SimHei 混写）'
    fonts = set(re.findall(r'w:(?:ascii|hAnsi|eastAsia|cs)="([^"]+)"', dx))
    alias_bad = (('宋体' in fonts and ({'SimSun', 'NSimSun'} & fonts))
                 or ('黑体' in fonts and 'SimHei' in fonts))
    results.append(('-1：' if alias_bad else '+0：') + rule)

    # 31 -1 首行缩进段落数
    rule = '设置了 2 字符首行缩进的段落少于 104 段'
    n_fl = 0
    for p in paras:
        e = pPr(p)
        ind = e.find(qn('w:ind')) if e is not None else None
        if ind is None:
            continue
        if (ind.get(qn('w:firstLine')) == '480'
                or ind.get(qn('w:firstLineChars')) == '200'):
            n_fl += 1
    results.append(('-1：' if n_fl < FIRSTLINE_MIN else '+0：') + rule)

    # 32 -1 TOC / PAGEREF 域被转为静态文字
    rule = '目录被转成静态文字，页码不再可自动更新'
    n_toc = len([m for m in re.findall(r'<w:instrText[^>]*>([^<]*)</w:instrText>', dx)
                 if m.strip().startswith('TOC')])
    n_pageref = len([m for m in re.findall(r'<w:instrText[^>]*>([^<]*)</w:instrText>', dx)
                     if m.strip().startswith('PAGEREF')])
    results.append((
        '-1：' if (n_toc != TOC_FIELD_COUNT or n_pageref < PAGEREF_MIN)
        else '+0：') + rule)

    # 33 -1 标题内空格数
    rule = '「目    录」「摘    要」「致    谢」（各 4 个空格）、「缩  略  语  表」（字间各 2 个空格）、「附录」（无空格）中有任一处空格数被改动'
    sp_bad = False
    for key, want in TITLE_SPACING.items():
        p = plain_norm.get(key)
        if p is None:
            sp_bad = True
            break
        if (p.text or '').strip() != want:
            sp_bad = True
            break
    results.append(('-1：' if sp_bad else '+0：') + rule)

    # 34 -1 附录 A-E 标题缺失
    rule = '缺少或改动「附录A  问卷核心题项节选」「附录B  社区步行环境现场观察表」「附录C  居民问卷量表核心题项」「附录D  指标计算说明」「附录E  访谈提纲与编码结果」五个标题中的任意一个'
    apx_bad = any(t not in all_norm_text for t in APPENDIX_TITLES)
    results.append(('-1：' if apx_bad else '+0：') + rule)

    # ===== 把 results 里的 "+N：..." / "-N：..." 拆解为结构化 dim2_items =====
    if len(results) != len(_MAX_DELTAS_ORDER):
        raise RuntimeError('dim2 项数不一致：results=%d vs order=%d'
                           % (len(results), len(_MAX_DELTAS_ORDER)))
    dim2_items = []
    total_score = 0
    for raw, max_delta in zip(results, _MAX_DELTAS_ORDER):
        head, _, rule = raw.partition('：')
        try:
            delta = int(head.replace('+', '').strip())
        except ValueError:
            delta = 0
        hit = (delta == max_delta and max_delta != 0)
        dim2_items.append({
            "rule": rule,
            "max_delta": max_delta,
            "delta": max_delta if hit else 0,
            "hit": hit,
        })
        total_score += (max_delta if hit else 0)
    # 单分值计分：总分 = 加分项得分 - 减分项扣分总和，不做下限截断，允许为负
    result["dim2_items"] = dim2_items
    result["total_score"] = total_score
    result["max_score"] = MAX_SCORE_TOTAL
    return result


if __name__ == '__main__':
    _dir = sys.argv[1] if len(sys.argv) >= 2 else os.path.dirname(
        os.path.abspath(__file__))
    print(json.dumps(evaluate(_dir), ensure_ascii=False, indent=2))
