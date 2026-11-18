#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# pyright: reportUnknownMemberType=false, reportUnknownVariableType=false, reportUnknownArgumentType=false, reportUnknownParameterType=false, reportMissingParameterType=false, reportPrivateUsage=false, reportUnknownLambdaType=false, reportGeneralTypeIssues=false

import os
import re
import sys
import json
import glob
import shutil
import difflib
import tempfile
import zipfile
from collections import Counter
from docx import Document
from docx.oxml.ns import qn
try:
    _VERIFIER_DIR = os.path.dirname(os.path.abspath(__file__))
    if _VERIFIER_DIR not in sys.path:
        sys.path.insert(0, _VERIFIER_DIR)
    import pymupdf_compat as fitz
except ImportError:
    fitz = None

# 说明：按接口统一约定，本模块不修改全局 sys.stdout，也不 print 主结果。
# 主结果统一由 evaluate(dir_path) -> dict 返回；入参为"脚本所在目录的路径"，
# 由脚本自己在该目录内定位并打开被评估的 .docx。省略入参（evaluate() 或直接
# `python officeval_001_verifier.py`）时，自动核查脚本所在目录下的交付文档。
# 本文件仅在 __main__ 下作为本地调试入口，才把 dict 序列化为 JSON 打印。
#
# 本脚本对应细则《课时学习方案（学历案）栏目删减、编号改写与两页压缩排版》。
# 细则的计分方式有一条总则：加分项只认交付前后发生的改变——某条要求的状态在
# 交付前已经成立（例如该文件本就不设"课时学习内容分析"一栏、本就只有两页），
# 该文件此条不计分也不扣分，其分值从该文件的满分基数中扣除。因此：
#   ① 各文件的 max_score 不是常数，需按文件逐一核定，由本脚本自动计算；
#   ② 脚本会尝试定位"交付前"的同编号 .docx 作为基准（依次在被评目录、
#      被评目录/上一级/上两级下的"输入文件""交付前"子目录中按文件名前两位数字匹配）；
#   ③ 找不到基准文件时，退回内置的 BASELINE_FALLBACK（记录八个待改文件交付前的
#      客观状态），此时依赖逐字比对的减分项无法判定，一律按"不扣分"处理
#      （判定依据仅在脚本内部流转，dim2_items 不输出 detail 字段）。
# 减分项同样只认"交付方改坏了什么"：页边距、主标题格式、表格列结构、图片压字/变形
# 等一律与交付前对比，交付前即如此的不扣分；若被评文件本身位于"输入文件/交付前"
# 目录（即以交付前状态自评），则视其为自己的基准，减分项自然全部落空、总分为 0。

SCRIPT_ID = "001"
# 加分项满分基数上限（交付前无任何条目已达成时的满分）；实际 max_score 按文件核定
MAX_SCORE_TOTAL = 32
# 减分项分值合计（全部命中时的扣分下限）
MIN_SCORE_TOTAL = -22

# 维度二所有评分项按"结果追加顺序"的 max_delta 列表；末端把 results 里的
# "+N：..." / "-N：..." / "=0：..." 字符串反解为结构化 dim2_items 时用。
# 顺序须与 _evaluate_impl 中 add(...) 的顺序严格一致；共 28 项
# （加分项 14 项累加 = 32；减分项 14 项累加 = -22）。
# 细则顺序表：(分值, 返回信息里使用的规则简写)。
# 前 14 条为加分项、后 14 条为减分项；末条为目录级，其余 27 条按文件逐份判定。
# 简写只为压缩返回信息体积，判定逻辑仍严格照细则原文，详见各 score()/penal() 调用处。
_RULES = [
    (1,  '清空一、课时学习目标（保留标题）'),
    (1,  '清空二、课时评价任务（保留标题）'),
    (3,  '删除课时学习内容分析'),
    (3,  '删除课时学生实际水平'),
    (3,  '删除活动意图说明及各处活动意图文字'),
    (3,  '删除全部图片'),
    (1,  '删除后表格上移、原位无残留空段'),
    (3,  '删除（指向目标N）且环节名称保留'),
    (1,  '表格标题为三、学习过程设计'),
    (1,  '板书设计编号为四、'),
    (1,  '教学反思编号为五、'),
    (5,  '总页数恰为 2 页'),
    (3,  '两页版面利用合理、无大面积空白'),
    (3,  '图片删除后行高收缩、无空白图框'),
    (-3, '纸张不再为 A4'),
    (-1, '页边距被压缩'),
    (-1, '主标题不满足黑体、加粗、居中独占一行'),
    (-1, '主标题字号不再明显大于正文'),
    (-3, '行距被设为小于 10 磅的固定值'),
    (-1, '段落对齐方式被改动'),
    (-3, '新增指令未要求的内容文字'),
    (-3, '删除指令未要求删除的整块内容'),
    (-1, '课型行选项文字或勾选标记被改'),
    (-1, '条目文字出现漏字、错字'),
    (-1, '指令未涉及的栏目编号被删或改错'),
    (-1, '表格列数或单元格合并关系被改动'),
    (-1, '文件名与文档内课时主题不一致'),
    (-1, '交付文档不足 8 个或同名重复交付'),
]
_MAX_DELTAS_ORDER = [d for d, _ in _RULES]
_RULE_LABELS = [t for _, t in _RULES]
# 末条（交付份数与重名）是目录级，逐份判定的只有前 _N_FILE_RULES 条。
_N_FILE_RULES = len(_RULES) - 1
assert sum(d for d in _MAX_DELTAS_ORDER if d > 0) == MAX_SCORE_TOTAL
assert sum(d for d in _MAX_DELTAS_ORDER if d < 0) == MIN_SCORE_TOTAL
assert len(_MAX_DELTAS_ORDER) == 28
# 找不到"交付前"基准 .docx 时使用的兜底表：记录八个待改文件交付前的客观状态。
# 键为文件名前两位数字；用于核定该文件哪些加分项属于"交付前已成立"而须排除。
#   sections   —— 交付前是否设有 (内容分析, 学生实际水平, 活动意图) 三部分
#   images     —— 交付前是否含图片
#   zhixiang   —— 交付前是否含"（指向目标N）"
#   goal_filled/eval_filled —— 交付前"课时学习目标""课时评价任务"两栏是否有描述文字
#   proc/board/reflect —— 交付前"学习过程设计""板书设计""教学反思"的栏目编号
#   pages      —— 交付前 Word 页数
#   layout_ok  —— 交付前是否已是"两页且两页均已排满"
BASELINE_FALLBACK = {
    '01': dict(sections=(True, True, True), images=True, zhixiang=True, goal_filled=True,
               eval_filled=True, proc='五、', board='七、', reflect='八、', pages=9, layout_ok=False),
    '02': dict(sections=(True, True, True), images=True, zhixiang=True, goal_filled=True,
               eval_filled=True, proc='五、', board='六、', reflect='七、', pages=7, layout_ok=False),
    '03': dict(sections=(False, False, False), images=True, zhixiang=True, goal_filled=True,
               eval_filled=True, proc='五、', board='七、', reflect='八、', pages=2, layout_ok=True),
    '04': dict(sections=(True, True, True), images=True, zhixiang=True, goal_filled=True,
               eval_filled=True, proc='五、', board='七、', reflect='八、', pages=6, layout_ok=False),
    '05': dict(sections=(False, False, False), images=True, zhixiang=True, goal_filled=True,
               eval_filled=True, proc='五、', board='七、', reflect='八、', pages=4, layout_ok=False),
    '06': dict(sections=(True, True, True), images=True, zhixiang=True, goal_filled=True,
               eval_filled=True, proc='五、', board='七、', reflect='八、', pages=5, layout_ok=False),
    '07': dict(sections=(True, True, True), images=True, zhixiang=True, goal_filled=True,
               eval_filled=True, proc='五、', board='七、', reflect='八、', pages=6, layout_ok=False),
    '08': dict(sections=(True, True, True), images=True, zhixiang=True, goal_filled=True,
               eval_filled=True, proc='五、', board='七、', reflect='八、', pages=6, layout_ok=False),
}

# "交付前"基准文件的搜索目录名（相对被评目录及其上两级）
BASELINE_DIR_NAMES = ('输入文件', '交付前', 'input', 'baseline')

W = qn('w:p').split('}')[0][1:]           # wordprocessingml 命名空间 URI
A_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main'
VML_NS = 'urn:schemas-microsoft-com:vml'
WP_NS = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing'

# 环节标题后的"（指向目标N）"
TAG_RE = re.compile(r'[（(]\s*指向目标[^）)]*[）)]')
# 栏目编号前缀（一、二、……十、）
NUM_RE = re.compile(r'^\s*([一二三四五六七八九十]+)\s*、\s*')
# 与文字无关的符号（判断"是否只差标点"时忽略）
PUNCT_ONLY_RE = re.compile(r'^[\s：:；;，,。.、·…—\-—－()（）《》“”"\'’‘【】\[\]/|+*#~]*$')


def _t(el):
    """取元素下所有 w:t 的拼接文本。"""
    return ''.join(n.text or '' for n in el.iter(qn('w:t')))


def _ws(s):
    """只去空白的规范化（保留☑/□与标点，供逐字比对用）。"""
    return re.sub(r'\s+', '', s or '')


def _norm(s):
    """去空白并剥掉"（指向目标N）"，供文本增删比对用。"""
    return _ws(TAG_RE.sub('', s or ''))
def _styles_index(doc):
    """styleId -> w:style 元素。"""
    out = {}
    for st in doc.styles.element.findall(qn('w:style')):
        out[st.get(qn('w:styleId'))] = st
    return out


def _para_style_id(p_el):
    pPr = p_el.find(qn('w:pPr'))
    if pPr is None:
        return None
    ps = pPr.find(qn('w:pStyle'))
    return ps.get(qn('w:val')) if ps is not None else None


def _walk_styles(sidx, style_id, pick):
    """沿 w:basedOn 链在样式上查找；pick(style_el) 返回非 None 即命中。"""
    seen = set()
    while style_id and style_id not in seen:
        seen.add(style_id)
        st = sidx.get(style_id)
        if st is None:
            return None
        got = pick(st)
        if got is not None:
            return got
        base = st.find(qn('w:basedOn'))
        style_id = base.get(qn('w:val')) if base is not None else None
    return None


def _resolve_jc(sidx, p_el):
    """按 Word 继承链解析段落实际生效的对齐方式：段落 w:jc → 段落样式链 w:jc。
       都没有则返回 None（表示随样式默认，不能据此判定"被改动"）。"""
    pPr = p_el.find(qn('w:pPr'))
    if pPr is not None:
        j = pPr.find(qn('w:jc'))
        if j is not None:
            return j.get(qn('w:val'))

    def pick(st):
        spr = st.find(qn('w:pPr'))
        if spr is None:
            return None
        j = spr.find(qn('w:jc'))
        return j.get(qn('w:val')) if j is not None else None

    return _walk_styles(sidx, _para_style_id(p_el), pick)


def _resolve_ea(doc, sidx, r_el, p_el):
    """解析 run 中文字符实际生效的东亚字体：run w:rFonts/@w:eastAsia →
       段落样式链 → docDefaults；都没有返回 None。"""
    rPr = r_el.find(qn('w:rPr'))
    if rPr is not None:
        rf = rPr.find(qn('w:rFonts'))
        if rf is not None and rf.get(qn('w:eastAsia')):
            return rf.get(qn('w:eastAsia'))

    def pick(st):
        spr = st.find(qn('w:rPr'))
        if spr is None:
            return None
        rf2 = spr.find(qn('w:rFonts'))
        return rf2.get(qn('w:eastAsia')) if rf2 is not None else None

    got = _walk_styles(sidx, _para_style_id(p_el), pick)
    if got:
        return got
    dd = doc.styles.element.find(qn('w:docDefaults'))
    if dd is not None:
        rpd = dd.find(qn('w:rPrDefault'))
        if rpd is not None:
            rpr = rpd.find(qn('w:rPr'))
            if rpr is not None:
                rf3 = rpr.find(qn('w:rFonts'))
                if rf3 is not None:
                    return rf3.get(qn('w:eastAsia'))
    return None
def _render_to_pdf(file_path, tmp_dir):
    """把 .docx 渲染为 PDF，用于获取真实分页与版面空白。
       仅通过本机 Microsoft Word (COM) 的 ExportAsFixedFormat 渲染；
       未安装 pywin32 或调用失败则抛出异常。"""
    out_pdf = os.path.join(tmp_dir, os.path.splitext(os.path.basename(file_path))[0] + '.pdf')
    errors = []
    try:
        import win32com.client  # pywin32
    except ImportError as exc:
        errors.append('未安装 pywin32, 无法调用 Word 渲染 PDF: ' + str(exc))
    else:
        word = None
        doc = None
        try:
            word = win32com.client.DispatchEx('Word.Application')
            word.Visible = False
            word.DisplayAlerts = 0
            doc = word.Documents.Open(os.path.abspath(file_path), False, True)
            doc.ExportAsFixedFormat(
                OutputFileName=os.path.abspath(out_pdf),
                ExportFormat=17,  # wdExportFormatPDF
            )
            if os.path.isfile(out_pdf):
                return out_pdf
            errors.append('Word PDF 渲染未生成目标文件')
        except Exception as exc:
            errors.append('Word PDF 渲染失败: ' + str(exc))
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
    raise RuntimeError('; '.join(errors) or '无法渲染 PDF')


def _page_metrics(file_path):
    """渲染后按页测量版面：返回 {'pages': N, 'pages_info': [...]} 或 None（无法渲染）。
       每页给出 top/bottom（内容首末墨迹行，占页高比例）、band（内容区内最大连续
       空白带占页高比例）、ink（有墨迹的行数占页高比例）。
       口径与"打开 Word 看到的效果"一致：以渲染像素为准，左右各忽略 60px 页边区，
       每行墨迹像素 ≤4 视为空白行（排除表格框线的影响）。"""
    if fitz is None:
        return None
    tmp = tempfile.mkdtemp(prefix='officeval_001_pdf_')
    try:
        try:
            pdf = _render_to_pdf(file_path, tmp)
        except Exception:
            return None
        info = []
        try:
            d = fitz.open(pdf)
            for page in d:
                pix = page.get_pixmap(dpi=72, colorspace=fitz.csGRAY)
                h, w = pix.height, pix.width
                buf = pix.samples
                dark = []
                for y in range(h):
                    row = buf[y * pix.stride: y * pix.stride + w]
                    cnt = sum(1 for x in range(60, w - 60) if row[x] < 200)
                    dark.append(cnt)
                rows = [y for y, c in enumerate(dark) if c > 0]
                if not rows:
                    info.append(dict(top=1.0, bottom=1.0, band=1.0, ink=0.0, ink_rows=0))
                    continue
                top, bot = rows[0], rows[-1]
                blank = [1 if dark[y] <= 4 else 0 for y in range(top, bot + 1)]
                best = cur = 0
                for b in blank:
                    cur = cur + 1 if b else 0
                    best = max(best, cur)
                info.append(dict(top=top / h, bottom=bot / h, band=best / h,
                                 ink=sum(1 for c in dark if c > 4) / h,
                                 ink_rows=sum(1 for c in dark if c > 4)))
            n = len(info)
            d.close()
        except Exception:
            return None
        return dict(pages=n, pages_info=info)
    finally:
        shutil.rmtree(tmp, ignore_errors=True)
def _ancestors(p_el):
    """返回 (单元格嵌套层数, 是否在文本框内, 最近的 w:tc 元素)。"""
    depth, txbx, tc = 0, False, None
    cur = p_el.getparent()
    while cur is not None:
        tag = cur.tag
        if tag == qn('w:tc'):
            depth += 1
            if tc is None:
                tc = cur
        if tag.endswith('txbxContent'):
            txbx = True
        cur = cur.getparent()
    return depth, txbx, tc


def _collect(path):
    """把一份文档的客观事实一次性抽出，供各评分项判定。
       统一用 body.iter(w:p) 按文档顺序遍历（含嵌套表格与文本框内的段落，
       每段只计一次），与打开办公软件后逐段看到的内容一致。"""
    doc = Document(path)
    body = doc.element.body
    sidx = _styles_index(doc)
    z = zipfile.ZipFile(path)
    raw = z.read('word/document.xml')

    f = dict(path=path, doc=doc, sidx=sidx, paras=[], rows=[], tables=[], headings=[],
             title=[], kexing='', theme='', goal_rest=None, eval_rest=None)

    for p in body.iter(qn('w:p')):
        t = _t(p)
        depth, txbx, tc = _ancestors(p)
        pPr = p.find(qn('w:pPr'))
        line = rule = None
        if pPr is not None:
            sp = pPr.find(qn('w:spacing'))
            if sp is not None:
                line, rule = sp.get(qn('w:line')), sp.get(qn('w:lineRule'))
        sizes, eas, asciis, bolds = [], [], [], []
        for r in p.iter(qn('w:r')):
            if not _t(r).strip():
                continue
            rPr = r.find(qn('w:rPr'))
            sz = rPr.find(qn('w:sz')) if rPr is not None else None
            if sz is not None and (sz.get(qn('w:val')) or '').isdigit():
                sizes.append(int(sz.get(qn('w:val'))) / 2.0)
            rf = rPr.find(qn('w:rFonts')) if rPr is not None else None
            if rf is not None:
                for k in ('ascii', 'hAnsi'):
                    if rf.get(qn('w:' + k)):
                        asciis.append(rf.get(qn('w:' + k)))
            eas.append(_resolve_ea(doc, sidx, r, p))
            b = rPr.find(qn('w:b')) if rPr is not None else None
            bolds.append(b is not None and (b.get(qn('w:val')) or 'true') not in ('0', 'false'))
        nimg = (len(list(p.iter('{%s}blip' % A_NS))) +
                len(list(p.iter('{%s}imagedata' % VML_NS))))
        f['paras'].append(dict(t=t, ws=_ws(t), norm=_norm(t), jc=_resolve_jc(sidx, p),
                               sizes=sizes, eas=eas, asciis=asciis, bolds=bolds,
                               line=line, rule=rule, depth=depth, txbx=txbx, tc=tc, nimg=nimg))
    # 正文级（body 直属）连续空段落的最大条数——细则"表格顺次上移、不遗留连续空行"用
    runs, cur, tot = [], 0, 0
    for ch in body:
        if ch.tag == qn('w:p'):
            has_txt = bool(_t(ch).strip())
            has_img = (len(list(ch.iter('{%s}blip' % A_NS))) +
                       len(list(ch.iter('{%s}imagedata' % VML_NS)))) > 0
            if not has_txt and not has_img:
                cur += 1
                tot += 1
                continue
        runs.append(cur)
        cur = 0
    runs.append(cur)
    f['empty_body_total'] = tot
    f['empty_body_max'] = max(runs) if runs else 0

    # 表格行高（缇）与该行文字量——细则"图片删除后行高一并收缩"用
    for tr in body.iter(qn('w:tr')):
        trPr = tr.find(qn('w:trPr'))
        h = rule = None
        if trPr is not None:
            hh = trPr.find(qn('w:trHeight'))
            if hh is not None and (hh.get(qn('w:val')) or '').isdigit():
                h, rule = int(hh.get(qn('w:val'))), hh.get(qn('w:hRule'))
        txt = _ws(_t(tr))
        nimg = (len(list(tr.iter('{%s}blip' % A_NS))) +
                len(list(tr.iter('{%s}imagedata' % VML_NS))))
        f['rows'].append(dict(h=h, rule=rule, text=txt[:30], textlen=len(txt), nimg=nimg))

    # 表格列数与合并签名——细则"表格列数或单元格合并关系被改动"用
    for tbl in body.iter(qn('w:tbl')):
        grid = tbl.find(qn('w:tblGrid'))
        cols = len(grid.findall(qn('w:gridCol'))) if grid is not None else 0
        sig, effs, cells_n, row_texts = [], [], [], []
        for tr in tbl.findall(qn('w:tr')):
            cells = tr.findall(qn('w:tc'))
            spans, eff = [], 0
            for tc in cells:
                tcPr = tc.find(qn('w:tcPr'))
                gs = vm = None
                if tcPr is not None:
                    g = tcPr.find(qn('w:gridSpan'))
                    gs = g.get(qn('w:val')) if g is not None else None
                    v = tcPr.find(qn('w:vMerge'))
                    vm = (v.get(qn('w:val')) or 'continue') if v is not None else None
                spans.append('%s/%s' % (gs or 1, vm or '-'))
                eff += int(gs) if (gs or '').isdigit() else 1
            sig.append('|'.join(spans))
            effs.append(eff)
            cells_n.append(len(cells))
            row_texts.append(_ws(_t(tr))[:60])
        f['tables'].append(dict(cols=cols, rows=len(sig), sig=tuple(sig), effs=tuple(effs),
                                cells_n=tuple(cells_n), row_texts=tuple(row_texts),
                                text=_ws(_t(tbl))[:400], head=_ws(_t(tbl))[:40]))
    # 图片：嵌入型 <wp:inline> / 浮动型 <wp:anchor>；旧式 <w:pict> 亦计入
    imgs = dict(total=0, inline=0, floating=[], media=0, pict=0, boxes=0)
    for dr in body.iter(qn('w:drawing')):
        inline = dr.find('{%s}inline' % WP_NS)
        anchor = dr.find('{%s}anchor' % WP_NS)
        node = inline if inline is not None else anchor
        cx = cy = None
        if node is not None:
            ext = node.find('{%s}extent' % WP_NS)
            if ext is not None:
                cx, cy = ext.get('cx'), ext.get('cy')
        if inline is not None:
            imgs['total'] += 1
            imgs['inline'] += 1
        elif anchor is not None:
            imgs['total'] += 1
            behind = (anchor.get('behindDoc') or '0') in ('1', 'true')
            wrap_none = anchor.find('{%s}wrapNone' % WP_NS) is not None
            imgs['floating'].append(dict(cx=cx, cy=cy, behind=behind, wrap_none=wrap_none))
    for pict in body.iter(qn('w:pict')):
        imgs['pict'] += 1
        has_img = pict.find('.//{%s}imagedata' % VML_NS) is not None
        if has_img:
            imgs['total'] += 1
            style = ''
            for sh in pict.iter():
                if sh.tag.split('}')[-1] in ('shape', 'rect'):
                    style = (sh.get('style') or '').replace(' ', '').lower()
                    break
            if 'position:absolute' in style:
                imgs['floating'].append(dict(cx=None, cy=None, behind=False, wrap_none=True))
            else:
                imgs['inline'] += 1
        else:
            imgs['boxes'] += 1        # 无图数据的空白图框/占位形状
    imgs['media'] = len([n for n in z.namelist() if n.startswith('word/media/')])
    f['images'] = imgs
    f['pict_raw'] = raw.count(b'<w:pict')

    # 纸张与页边距（mm）：文档可能含多个分节符，取各边"出现最多"的设置为全文口径，
    # 免得个别节（如尾节的 0 边距占位）压过全文真实版面。
    mm = lambda v: round(int(v) / 1440.0 * 25.4, 1) if v is not None else None
    vals = dict(w=[], h=[], l=[], r=[], t=[], b=[])
    for sect in body.iter(qn('w:sectPr')):
        pg = sect.find(qn('w:pgSz'))
        mar = sect.find(qn('w:pgMar'))
        if pg is not None:
            for k, a in (('w', 'w:w'), ('h', 'w:h')):
                vals[k].append(mm(pg.get(qn(a))))
        if mar is not None:
            for k, a in (('l', 'w:left'), ('r', 'w:right'), ('t', 'w:top'), ('b', 'w:bottom')):
                vals[k].append(mm(mar.get(qn(a))))
    f['page'] = dict((k, (Counter([x for x in v if x is not None]).most_common(1)[0][0]
                          if any(x is not None for x in v) else None))
                     for k, v in vals.items())
    f['page_sects'] = len(vals['l'])
    # 全文文本：ws=仅去空白（逐字比对用）；norm=另剥"（指向目标N）"；
    # cmp=再剥掉段首栏目编号（"三、"↔"五、"这类指令许可的改写不算文字增删）
    f['full_ws'] = ''.join(p['ws'] for p in f['paras'])
    f['full_norm'] = ''.join(p['norm'] for p in f['paras'])
    f['cmp_text'] = ''.join(NUM_RE.sub('', p['norm']) for p in f['paras'])

    # 栏目编号：段首"N、栏目名"
    for p in f['paras']:
        m = NUM_RE.match(p['ws'])
        if not m:
            continue
        name = re.split(r'[：:]', p['ws'][m.end():])[0][:12]
        if name:
            f['headings'].append((m.group(1) + '、', name))

    # 主标题段（"第一课时学习方案设计"一类）
    for p in f['paras']:
        if '课时学习方案设计' in p['ws']:
            f['title'].append(p)

    # 正文字号分布（不含主标题段）
    f['body_sizes'] = Counter(s for p in f['paras']
                              if p['ws'] and '课时学习方案设计' not in p['ws'] for s in p['sizes'])

    # "课型"一行（含☑/□勾选标记，逐字比对用）
    for p in f['paras']:
        if '新授课' in p['ws']:
            f['kexing'] = p['ws']
            break

    # 文档内"课时主题"的值（取同一表格行内除"课时主题"以外的文字）
    for p in f['paras']:
        if '课时主题' in p['ws'] and p['tc'] is not None:
            tr = p['tc'].getparent()
            f['theme'] = _ws(_t(tr)).replace('课时主题', '')
            break

    # "一、课时学习目标""二、课时评价任务"两栏标题之后的残留文字：
    # 从标题所在段起，沿同一单元格（或同一层正文）向后取到下一个"N、"栏目标题为止，
    # 与打开文档后"这一栏里还剩什么字"看到的一致（不把后续嵌套表格的内容算进来）。
    def _residual_of(key):
        for i, p in enumerate(f['paras']):
            if key not in p['ws']:
                continue
            txt = p['ws'][p['ws'].find(key) + len(key):]
            for q in f['paras'][i + 1:]:
                if (q['tc'] is not p['tc']) or (q['depth'] != p['depth']):
                    break
                if NUM_RE.match(q['ws']):
                    break
                txt += q['ws']
            return True, re.sub(r'^[：:]\s*', '', txt)
        return False, None

    f['goal_found'], f['goal_rest'] = _residual_of('课时学习目标')
    f['eval_found'], f['eval_rest'] = _residual_of('课时评价任务')

    # 段落对齐比对表：规范化文本前 40 字 -> 生效对齐方式
    jc_map = {}
    for p in f['paras']:
        k = p['norm'][:40]
        if k:
            jc_map.setdefault(k, []).append(p['jc'])
    f['jc_map'] = jc_map
    # 三部分栏目 / 图片 / 指向目标 的存在性（既用于本文件判分，
    # 也用于把它当作"交付前"基准时核定哪些条目应排除）
    f['has_analysis'] = '课时学习内容分析' in f['full_ws']
    f['has_level'] = '课时学生实际水平' in f['full_ws']
    f['has_intent'] = '活动意图' in f['full_ws']
    f['has_images'] = f['images']['total'] > 0
    f['has_tag'] = bool(TAG_RE.search(''.join(p['ws'] for p in f['paras'])))
    f['num_of'] = {name: num for num, name in f['headings']}
    return f


def _heading_num(f, name):
    """取某栏目标题在文档中的编号（如"板书设计"->"四、"）；无该栏目返回 None。"""
    for num, nm in f['headings']:
        if nm.startswith(name) or name in nm:
            return num
    return None


def _has_col(f, name):
    """文档中是否出现该栏目的文字标题（不论编号）。"""
    return name in f['full_ws']


def _delivery_files(dir_path):
    """列出一个目录内的交付文档（.docx/.doc，跳过 Word 临时文件 ~$）。"""
    return [p for p in sorted(glob.glob(os.path.join(dir_path, '*.docx')) +
                              glob.glob(os.path.join(dir_path, '*.doc')))
            if not os.path.basename(p).startswith('~$')]


def _script_dir():
    """脚本自身所在目录（交互式执行时退回当前工作目录）。"""
    try:
        return os.path.dirname(os.path.abspath(__file__)) or os.getcwd()
    except NameError:
        return os.getcwd()


def _default_target():
    """默认核查对象：脚本所在目录。脚本与交付文档同放一处时，直接运行即可；
       若该目录内没有交付文档，则退到其下唯一含交付文档的子目录（排除"输入文件"
       等交付前目录）；仍不唯一时原样返回脚本所在目录，由 evaluate 报明情况。"""
    here = _script_dir()
    if _delivery_files(here):
        return here
    subs = [d for d in sorted(glob.glob(os.path.join(here, '*')))
            if os.path.isdir(d) and os.path.basename(d) not in BASELINE_DIR_NAMES
            and _delivery_files(d)]
    return subs[0] if len(subs) == 1 else here


def _find_baseline(file_path):
    """定位"交付前"的同编号 .docx：按文件名前两位数字，在被评目录及其上两级的
       "输入文件""交付前"等子目录中查找。找不到返回 None。"""
    base = os.path.basename(file_path)
    m = re.match(r'\s*(\d{2})', base)
    if not m:
        return None
    num = m.group(1)
    here = os.path.dirname(os.path.abspath(file_path))
    roots = [here, os.path.dirname(here), os.path.dirname(os.path.dirname(here))]
    cands = []
    for root in roots:
        for name in BASELINE_DIR_NAMES:
            d = os.path.join(root, name)
            if os.path.isdir(d):
                cands.extend(sorted(glob.glob(os.path.join(d, num + '*.docx'))))
    for c in cands:
        if os.path.basename(c).startswith('~$'):
            continue
        if os.path.abspath(c) != os.path.abspath(file_path):
            return c
    return None


# 任务四要求"给表格设一个标题三、学习过程设计"，该处新增文字不算"多余新增"
ALLOWED_ADD = ('学习过程设计',)
# 编号只许改这三处（三、学习过程设计 / 四、板书设计 / 五、教学反思）
RENUM_NAMES = ('学习过程', '板书设计', '教学反思')


def _bl_facts(b, num):
    """核定该文件"交付前"的事实：结构性事实优先用找到的基准 .docx 实测；
       页数与版面利用优先取内置基线（避免再把 8 份基准文件渲染一遍）。
       两者都拿不到时返回 None，此时依赖"交付前"对比的条目一律判为无法认定。"""
    d = BASELINE_FALLBACK.get(num or '')
    out = None
    if b is not None:
        out = dict(analysis=b['has_analysis'], level=b['has_level'], intent=b['has_intent'],
                   images=b['has_images'], zhixiang=b['has_tag'],
                   goal_filled=bool(_ws(b['goal_rest'] or '')),
                   eval_filled=bool(_ws(b['eval_rest'] or '')),
                   proc=_heading_num(b, '学习过程'), board=_heading_num(b, '板书设计'),
                   reflect=_heading_num(b, '教学反思'), pages=None, layout_ok=None,
                   src='交付前文件实测')
    elif d:
        a, l, i = d['sections']
        out = dict(analysis=a, level=l, intent=i, images=d['images'], zhixiang=d['zhixiang'],
                   goal_filled=d['goal_filled'], eval_filled=d['eval_filled'],
                   proc=d['proc'], board=d['board'], reflect=d['reflect'],
                   pages=None, layout_ok=None, src='内置交付前基线')
    if out is None:
        return None
    if d:
        out['pages'], out['layout_ok'] = d['pages'], d['layout_ok']
    elif b is not None:
        pm = _page_metrics(b['path'])
        out['pages'] = pm['pages'] if pm else None
    return out


def _allowed_removals(b):
    """交付前文件里"指令许可删除"的文字（按文档顺序拼接）：
       "课时学习目标""课时评价任务"两栏的内容、三部分整段、各处"活动意图：……"。
       删除比对时先扣掉这些片段，剩下的才算"删了指令未要求删除的内容"。"""
    if b is None:
        return ''
    cells = []
    for key in ('课时学习目标', '课时评价任务'):
        for p in b['paras']:
            if key in p['ws'] and p['tc'] is not None:
                cells.append(p['tc'])
                break
    blob, dropping = [], False
    for p in b['paras']:
        ws = p['ws']
        if not ws:
            continue
        if ('课时学习内容分析' in ws) or ('课时学生实际水平' in ws) or ('活动意图说明' in ws):
            dropping = True
        elif dropping and NUM_RE.match(ws):
            dropping = False
        take = dropping or ('活动意图' in ws) or (p['tc'] is not None and p['tc'] in cells)
        if not take:
            continue
        s = ws
        for key in ('课时学习目标', '课时评价任务'):
            if key in s:
                s = s[s.find(key) + len(key):]
        blob.append(NUM_RE.sub('', _norm(s)))
    return ''.join(blob)


def _diff_ops(before, after):
    """逐字比对交付前后的正文，拆成"新增片段/删除片段/短替换（疑似漏字错字）"。
       长替换（任一侧≥8字）按"删一段+加一段"处理，交给对应条目判定。"""
    ins, dele, typo = [], [], []
    sm = difflib.SequenceMatcher(None, before, after, autojunk=False)
    for tag, i1, i2, j1, j2 in sm.get_opcodes():
        if tag == 'insert':
            ins.append(after[j1:j2])
        elif tag == 'delete':
            dele.append(before[i1:i2])
        elif tag == 'replace':
            a, c = before[i1:i2], after[j1:j2]
            if max(len(a), len(c)) >= 8:
                dele.append(a)
                ins.append(c)
            else:
                typo.append((a, c))
    return ins, dele, typo


def _still_there(s, blob, win=6):
    """s 的文字是否仍在 blob 里（允许前后少量差异，如栏目编号被合并、位置前后调整）：
       按 6 字滑窗计，命中 8 成以上即认为文字仍在，只是换了位置或合并了重复行。"""
    if not s:
        return True
    if len(s) <= win:
        return s in blob
    wins = [s[i:i + win] for i in range(len(s) - win + 1)]
    return sum(1 for w in wins if w in blob) >= 0.8 * len(wins)


def _residual(chunk, blob, min_block=6):
    """从 chunk 中扣掉与 blob（许可删除文字）对齐的连续片段，返回剩余文字。"""
    if not chunk or not blob:
        return chunk
    keep = [True] * len(chunk)
    sm = difflib.SequenceMatcher(None, chunk, blob, autojunk=False)
    for blk in sm.get_matching_blocks():
        if blk.size >= min_block:
            for k in range(blk.a, blk.a + blk.size):
                keep[k] = False
    return ''.join(c for c, k in zip(chunk, keep) if k)


def _dim1(f):
    """维度1：可用与可修改性。文档能打开、正文为可点选文字、表格结构在即通过；
       版面参数（字号、行距、页边距）问题一律走维度2减分项，不在此处判 0。"""
    if len(f['full_ws']) < 200:
        return False, '正文可单独点选的文字不足 200 字（疑似整页位图或内容大面积丢失）'
    if not f['tables']:
        return False, '文档中不存在表格结构（学生活动/教师活动/评价要点表格丢失）'
    if f['images']['media'] > 0 and len(f['full_ws']) < 400:
        return False, '正文以图片为主、可编辑文字过少'
    return True, '文档可正常打开，正文与表格均为可单独点选、编辑的内容'


def _theme_core(theme):
    """取"课时主题"的核心文字（优先书名号内），供文件名一致性比对。"""
    m = re.search(r'《([^》]+)》', theme or '')
    if m:
        return m.group(1)
    s = re.sub(r'(课时主题|主题)', '', theme or '')
    s = re.split(r'[：:（(]', s)[0]
    return s.strip()


def _evaluate_impl(file_path):
    """按细则逐条判定单份交付文档，返回 27 条（末条为目录级，在 evaluate 中补齐）
       的判定串与证据串。判定串前缀：+N=加分项达成，+0=未达成，-N=命中减分项，
       -0=未命中减分项，=0=该条在交付前已成立，不计分也不扣分。"""
    name = os.path.basename(file_path)
    mnum = re.match(r'\s*(\d{2})', name)
    num = mnum.group(1) if mnum else ''
    f = _collect(file_path)
    bpath = _find_baseline(file_path)
    if bpath:
        b = _collect(bpath)
    elif os.path.basename(os.path.dirname(os.path.abspath(file_path))) in BASELINE_DIR_NAMES:
        # 被评文件本身就位于"交付前"目录：它即是自己的基准，等同于一字未改，
        # 减分项（均为"被改坏"才成立）自然全部落空。
        b, bpath = f, file_path
    else:
        b = None
    blf = _bl_facts(b, num)
    pm = _page_metrics(file_path)
    results, details = [], []

    def score(ok, mag, rule, detail, excluded=False, ex_reason=''):
        if excluded:
            results.append('=0：' + rule)
            details.append(ex_reason or '交付前已满足，本条不计分、也不扣分')
        else:
            results.append(('+%d：' % mag if ok else '+0：') + rule)
            details.append(detail)

    def penal(bad, mag, rule, detail, unknown=False):
        results.append(('-0：' if (unknown or not bad) else '-%d：' % mag) + rule)
        details.append(('无法认定：' + detail) if unknown else detail)

    # —— 加分项 01/02：清空两栏内容、保留标题 ——
    # 只看"标题还在、标题所在单元格里已无描述文字"，交付前本就为空的不计分。
    for key, found, rest, filled, tag in (
            ('一、课时学习目标', f['goal_found'], f['goal_rest'],
             None if blf is None else blf['goal_filled'], '目标描述'),
            ('二、课时评价任务', f['eval_found'], f['eval_rest'],
             None if blf is None else blf['eval_filled'], '任务描述')):
        left = _ws(rest or '')
        score(bool(found) and not left, 1,
              '保留"%s"标题文字，且该标题所在单元格内不再出现任何%s文字' % (key, tag),
              '标题%s；标题后残留 %d 字：%s' % ('保留' if found else '已丢失', len(left), left[:40]),
              excluded=(filled is False), ex_reason='交付前该栏本就为空')


    # —— 加分项 03/04/05：删除三部分内容（含标题）——
    score(not f['has_analysis'], 3,
          '全文不再出现"课时学习内容分析"标题及其下的学习内容分析文字',
          '仍存在该栏文字' if f['has_analysis'] else '已彻底清除',
          excluded=(blf is not None and not blf['analysis']), ex_reason='交付前不设此栏')
    score(not f['has_level'], 3,
          '全文不再出现"课时学生实际水平"标题及其下的学生水平描述文字',
          '仍存在该栏文字' if f['has_level'] else '已彻底清除',
          excluded=(blf is not None and not blf['level']), ex_reason='交付前不设此栏')
    intent_left = [p['ws'][:24] for p in f['paras'] if '活动意图' in p['ws']]
    score(not intent_left, 3,
          '全文不再出现"活动意图说明"标题，以及各环节表格中的"活动意图：……"整段文字',
          ('遗留 %d 处：%s' % (len(intent_left), '｜'.join(intent_left[:3]))) if intent_left else '各处均已清除',
          excluded=(blf is not None and not blf['intent']), ex_reason='交付前各环节均无"活动意图"文字')

    # —— 加分项 06：删除全部图片（嵌入式 + 浮动图形）——
    img = f['images']
    score(not f['has_images'], 3,
          '文档中不再包含任何图片——环节插图、板书示意图、流程图（含嵌入式图片与浮动图形）全部删除',
          '嵌入式 %d 张、浮动 %d 张、旧式图形 %d 处、包内 media 文件 %d 个'
          % (img['inline'], len(img['floating']), img['pict'], img['media']),
          excluded=(blf is not None and not blf['images']), ex_reason='交付前本就不含任何图片')

    # —— 加分项 07：删除后原位无残留空段、表格顺次上移 ——
    cleared = (not f['has_analysis'] and not f['has_level'] and not intent_left
               and not f['has_images'])
    score(cleared and f['empty_body_max'] <= 2 and img['boxes'] == 0, 1,
          '被删内容的原位置不遗留空段落、空行或空白图框，其余表格顺次上移衔接',
          '正文级连续空段最多 %d 个（共 %d 个）、空白图框 %d 个%s'
          % (f['empty_body_max'], f['empty_body_total'], img['boxes'],
             '' if cleared else '；仍存有图片或三部分文字，本条不得分'),
          excluded=(blf is not None and not (blf['images'] or blf['analysis']
                                             or blf['level'] or blf['intent'])),
          ex_reason='交付前本就无图片、无三部分内容，无需上移')


    # —— 加分项 08：删除"（指向目标N）"且环节名称本身完整保留 ——
    kept = True
    lost = []
    if b is not None:
        for p in b['paras']:
            if not p['ws'].startswith('环节'):
                continue
            want = TAG_RE.sub('', p['ws'])[:12]
            if want and want not in f['full_ws']:
                kept = False
                lost.append(want)
    score(not f['has_tag'] and kept, 3,
          '所有"环节一/二/三/四"标题后不再出现"（指向目标N）"，且环节名称本身完整保留',
          ('仍有"（指向目标N）"字样' if f['has_tag'] else '"（指向目标N）"已全部删除')
          + ('；环节名称缺失：' + '｜'.join(lost) if lost else '；环节名称完整'),
          excluded=(blf is not None and not blf['zhixiang']),
          ex_reason='交付前全文本就无"（指向目标N）"')

    # —— 加分项 09/10/11：栏目标题与编号改写 ——
    score('三、学习过程设计' in f['full_ws'], 1,
          '学生活动/教师活动/评价要点所在的学习过程表格，其上方标题为"三、学习过程设计"',
          '学习过程栏现编号：%s' % (_heading_num(f, '学习过程') or '未见该标题'),
          excluded=(blf is not None and blf['proc'] == '三、'),
          ex_reason='交付前该标题本就是"三、学习过程设计"')
    board_num = _heading_num(f, '板书设计')
    score(board_num == '四、', 1, '"板书设计"栏目标题的编号为"四、"（即"四、板书设计"）',
          '板书设计现编号：%s' % (board_num or ('无编号' if '板书设计' in f['full_ws']
                                            else '文档中已无"板书设计"标题')),
          excluded=(blf is not None and blf['board'] in (None, '四、')),
          ex_reason='交付前不含"板书设计"文字标题，或本就编号为"四、"')
    reflect_num = _heading_num(f, '教学反思')
    score(reflect_num == '五、', 1, '"教学反思"栏目标题的编号为"五、"（即"五、教学反思"）',
          '教学反思现编号：%s' % (reflect_num or ('无编号' if '教学反思' in f['full_ws']
                                             else '文档中已无"教学反思"标题')),
          excluded=(blf is not None and blf['reflect'] in (None, '五、')),
          ex_reason='交付前本就编号为"五、"')


    # —— 加分项 12/13：压到两页且两页版面利用合理 ——
    # 与办公软件对齐：用 Word 渲染 PDF 后数页数、量每页墨迹分布，
    # 不用估算字符数，避免与"打开 Word 看到的页数"不一致。
    pages = pm['pages'] if pm else None
    info = pm['pages_info'] if pm else []
    score(pages == 2, 5, '用 Word 打开后文档总页数恰为 2 页（既未超出 2 页，也未压成 1 页）',
          ('实测 %d 页' % pages) if pages is not None else '无法渲染，未取得页数（未安装 Word 或 PDF 兼容后端）',
          excluded=(blf is not None and blf['pages'] == 2), ex_reason='交付前本就只有 2 页')
    lay_ok, lay_note = False, '页数不为 2，版面利用不予认定'
    if pages == 2 and len(info) == 2:
        p1, p2 = info
        lay_ok = (p1['bottom'] >= 0.75 and p1['band'] <= 0.25
                  and p2['band'] <= 0.25 and p2['ink_rows'] >= 25)
        lay_note = ('第一页内容至页高 %.0f%%、最大空白带 %.0f%%；第二页墨迹 %d 行、最大空白带 %.0f%%'
                    % (p1['bottom'] * 100, p1['band'] * 100, p2['ink_rows'], p2['band'] * 100))
    elif pages is None:
        lay_note = '无法渲染，未取得版面数据'
    score(lay_ok, 3, '两页版面利用合理——第二页有实质内容，全文没有大面积空白', lay_note,
          excluded=(blf is not None and blf['pages'] == 2 and blf['layout_ok']),
          ex_reason='交付前本就只有 2 页且两页均已排满')

    # —— 加分项 14：图片删除后其所在行、单元格的固定行高一并收缩 ——
    # 口径按细则"不出现半页高的空白单元格"：只认 ≥200 磅（约半页文字高）的空行高，
    # 再叠加渲染结果——页内最大空白带超过 30% 页高即视为成片空白。
    tall = [(r['text'], r['h']) for r in f['rows']
            if r['h'] and r['h'] >= 4000 and r['textlen'] < 20]
    band_max = max([p['band'] for p in info] or [0])
    score(not f['has_images'] and not tall and img['boxes'] == 0 and band_max <= 0.30, 3,
          '图片删除后其所在表格行、单元格的固定行高一并收缩，不残留空白图框、空白段落或成片空白',
          ('残留高行 %d 行（如"%s" 行高 %.1f 磅）' % (len(tall), tall[0][0][:12], tall[0][1] / 20.0))
          if tall else ('仍存有图片，行高不予认定' if f['has_images']
                        else ('渲染后仍有 %.0f%% 页高的成片空白' % (band_max * 100)
                              if band_max > 0.30 else
                              '无半页高的空行高、无空白图框、无成片空白（最大空白带 %.0f%%）'
                              % (band_max * 100))),
          excluded=(blf is not None and not blf['images']), ex_reason='交付前本就不含图片')


    # —— 减分项 15/16：纸张与页边距（不在"字号、行距、图文间距"许可范围内）——
    pg = f['page']
    a4 = (pg['w'] is not None and pg['h'] is not None
          and abs(pg['w'] - 210) <= 1.5 and abs(pg['h'] - 297) <= 1.5)
    penal(not a4, 3, '纸张不再为 A4（宽 210mm、高 297mm）',
          '实测 %s×%s mm' % (pg['w'], pg['h']), unknown=(pg['w'] is None))
    # 只认"被压缩"：某边现在过窄、且确实比交付前更小才计；交付前本就窄的不算交付方的错。
    bpg = b['page'] if b is not None else None
    narrow = []
    for k, v, prev, lim in (('左', pg['l'], None if bpg is None else bpg['l'], 18),
                            ('右', pg['r'], None if bpg is None else bpg['r'], 18),
                            ('上', pg['t'], None if bpg is None else bpg['t'], 7)):
        if v is not None and v < lim and (prev is None or v < prev - 0.5):
            narrow.append('%s（%s→%s mm）' % (k, prev, v) if prev is not None else '%s（%s mm）' % (k, v))
    penal(bool(narrow), 1, '页边距被压缩——左右边距应保持在 18mm 以上、上边距不小于 7mm',
          '实测 左%s 右%s 上%s 下%s mm%s' % (pg['l'], pg['r'], pg['t'], pg['b'],
                                          ('；%s 被压到许可范围外' % '、'.join(narrow)) if narrow
                                          else ('；与交付前一致' if bpg is not None else '')),
          unknown=(pg['l'] is None))

    # —— 减分项 17/18：主标题的字体、加粗、独占一行居中与字号层级 ——
    # 减分项只认"被改坏"：主标题的字体名、加粗、居中与交付前一致时不扣分，
    # 只有交付后偏离了"黑体、加粗、独占一行居中"且与交付前不同才计。
    tp = f['title']
    eas = sorted({e for p in tp for e in p['eas'] if e})
    bolds = [x for p in tp for x in p['bolds']]
    b_tp = b['title'] if b is not None else []
    b_eas = sorted({e for p in b_tp for e in p['eas'] if e})
    b_bold = any(x for p in b_tp for x in p['bolds'])
    why = []
    if tp:
        if eas and not any(('黑体' in e) or ('Hei' in e) for e in eas) and eas != b_eas:
            why.append('字体被改为 %s' % '、'.join(eas))
        if bolds and not any(bolds) and (not b_tp or b_bold):
            why.append('加粗被去掉')
        if tp[0]['jc'] not in (None, 'center'):
            why.append('对齐被改为 %s' % tp[0]['jc'])
        if len(tp[0]['ws']) > 24:
            why.append('与其他文字挤在同一段，未独占一行')
        if b_tp and b_tp[0]['jc'] == 'center' and tp[0]['jc'] is None:
            why.append('原有的居中设置被去掉')
    penal(bool(why), 1,
          '"第一课时学习方案设计""第二课时学习方案设计"主标题不满足 黑体、加粗、独占一行居中显示',
          '；'.join(why) if why else '主标题的字体、加粗与居中未被改动', unknown=(not tp))
    tsz = max([s for p in tp for s in p['sizes']] or [0])
    bsz = f['body_sizes'].most_common(1)[0][0] if f['body_sizes'] else 0
    penal(tsz > 0 and bsz > 0 and tsz <= bsz * 1.15, 1,
          '主标题字号被改到不再明显大于正文字号，标题与正文的层级差消失',
          '主标题 %s 磅、正文主用 %s 磅' % (tsz or '未显式设置', bsz or '未显式设置'),
          unknown=(tsz == 0 or bsz == 0))


    # —— 减分项 19：行距被设为小于 10 磅的固定值 ——
    fixed = [(p['ws'][:12], int(p['line']) / 20.0) for p in f['paras']
             if p['ws'] and p['rule'] == 'exact' and p['line'] and p['line'].isdigit()
             and int(p['line']) / 20.0 < 10]
    penal(bool(fixed), 3, '行距被设为小于 10 磅的固定值（如固定值 6.75 磅），文字上下被压叠',
          ('%d 段命中，最小 %.2f 磅（如"%s"）'
           % (len(fixed), min(v for _, v in fixed), fixed[0][0])) if fixed
          else '无小于 10 磅的固定行距')

    # —— 减分项 20：段落对齐方式被改动 ——
    # 只认细则写出的破坏方向：交付前左对齐/随默认的条目被改成居中或两端对齐；
    # 主标题的居中另由上一条判定，不在此重复计。
    jc_diff = []
    title_keys = {p['norm'][:40] for p in tp}
    if b is not None:
        for k, v in f['jc_map'].items():
            if k not in b['jc_map'] or k in title_keys:
                continue
            was = {x or 'default' for x in b['jc_map'][k]}
            now = {x or 'default' for x in v}
            if was == now:
                continue
            if (now - was) & {'center', 'both', 'right'} and was <= {'default', 'left'}:
                jc_diff.append('"%s"：%s→%s' % (k[:14], '/'.join(sorted(was)), '/'.join(sorted(now))))
    penal(bool(jc_diff), 1, '段落对齐方式被改动（如教师活动栏内左对齐的条目被改为居中或两端对齐）',
          '；'.join(jc_diff[:3]) if jc_diff else '留存段落未出现左对齐被改为居中/两端对齐的情形',
          unknown=(b is None))

    # —— 文字增删比对：以剥掉"（指向目标N）"与段首栏目编号后的全文逐字对齐 ——
    ins, dele, typo = _diff_ops(b['cmp_text'], f['cmp_text']) if b is not None else ([], [], [])
    allowed = _allowed_removals(b)

    def _core(s):
        return ''.join(ch for ch in s if not PUNCT_ONLY_RE.match(ch))

    out_core = _core(f['cmp_text'])
    bl_core = _core(b['cmp_text']) if b is not None else ''
    add_blob = ''.join(ALLOWED_ADD)

    # —— 减分项 21：新增指令未要求的内容文字 ——
    # 任务四要求补写的"三、学习过程设计"不计；文字仅换了位置（交付前已有同样文字）也不计。
    added = []
    for s in ins:
        t = _core(_residual(s, add_blob, min_block=2))
        if len(t) >= 4 and not _still_there(t, bl_core):
            added.append(t[:30])
    penal(bool(added), 3, '新增了指令未要求的内容文字（如自行补写整段板书内容）',
          ('%d 处，如"%s"' % (len(added), added[0])) if added else '未见新增文字',
          unknown=(b is None))

    # —— 减分项 22：删除了指令未要求删除的整块内容 ——
    removed = []
    for s in dele:
        r = _core(_residual(s, allowed))
        if len(r) < 8 or _still_there(r, out_core):
            continue          # 文字仍在文档中（仅重复行被合并或位置前后调整），不算整块删除
        removed.append(r[:30])
    penal(bool(removed), 3,
          '删除了指令未要求删除的整块内容（如"板书设计""教学反思"整栏、某环节的活动条目或评价要点）',
          ('%d 处，如"%s"' % (len(removed), removed[0])) if removed else '除许可删除项外文字均在',
          unknown=(b is None))


    # —— 减分项 23："课型"一行的选项文字或勾选标记发生变化 ——
    penal(b is not None and bool(b['kexing']) and f['kexing'] != b['kexing'], 1,
          '"课型"一行的四个选项文字有增删改，或勾选标记（☑/□）发生变化',
          '交付前：%s；交付后：%s' % ((b['kexing'] if b is not None else '')[:40], f['kexing'][:40]),
          unknown=(b is None or not b['kexing']))

    # —— 减分项 24：条目文字出现漏字、错字 ——
    typos = []
    for a, c in typo:
        ca, cc = _core(a), _core(c)
        if ca == cc or (not ca and not cc):
            continue
        if not cc and ca and ca in allowed:
            continue          # 属两栏内容/三部分的许可删除残片
        if not cc and ca.isdigit():
            continue
        typos.append('"%s"→"%s"' % (a[:12], c[:12]))
    penal(bool(typos), 1, '条目文字出现漏字、错字（如"同伴提出修改建议"被写成"同伴提出修改议"）',
          '；'.join(typos[:3]) if typos else '留存文字逐字一致', unknown=(b is None))

    # —— 减分项 25：指令未涉及的栏目编号被删掉或改错 ——
    # 现有 8 份输入文件已无"板书设计""教学反思"之外的其他编号栏目（原《家乡风物介绍》
    # 里"六、课堂评价与拓展任务"整栏已不在输入文件中），故实际须照旧的编号是
    # "一、课时学习目标""二、课时评价任务"。各文件交付前编号本就跳号
    # （01/03/04/05/06/07/08 为 五、学习过程设计／七、板书设计／八、教学反思，无"六、"；
    # 02 为 五、／六、／七、），交付方不得因此自行重排或补齐序号：凡交付前带编号、
    # 交付后该栏目文字标题仍在却编号变动或消失的，一律计。
    num_bad = []
    checked = []
    if b is not None:
        for bnum, bname in b['headings']:
            if any(k in bname for k in RENUM_NAMES) or not bname:
                continue
            if bname not in f['full_ws']:
                continue      # 该栏目已整体不在（归入其他条目判定）
            got = f['num_of'].get(bname)
            checked.append('%s%s' % (got or '无编号', bname))
            if got != bnum:
                num_bad.append('"%s"：%s→%s' % (bname, bnum, got or '无编号'))
    penal(bool(num_bad), 1,
          '指令未涉及的栏目编号被删掉或改错（除"四、板书设计""五、教学反思"外一律照旧，'
          '现存须照旧的为"一、课时学习目标""二、课时评价任务"）',
          '；'.join(num_bad[:3]) if num_bad
          else (('其余栏目编号照旧（已核：%s）' % '、'.join(checked[:4])) if checked
                else '交付前不存在指令未涉及的编号栏目，本条无判定对象'),
          unknown=(b is None))

    # —— 减分项 26：表格列数或单元格合并关系被改动 ——
    # 只看细则关心的后果：学生活动/教师活动/评价要点三列的对应关系是否错位。
    # 定位表头行即"学生活动｜教师活动｜评价要点"的那张表（整份学历案外层是一张大表，
    # 学习过程表是其中的嵌套表，故按表头行文字锁定，而不是按表内是否提到这些字）。
    def _proc_tbl(g):
        for t in g['tables']:
            for i, rt in enumerate(t['row_texts']):
                if '学生活动' in rt and '教师活动' in rt and len(rt) <= 40:
                    return t, i
        return None, None

    pt, hi = _proc_tbl(f)
    bt, bhi = (_proc_tbl(b) if b is not None else (None, None))
    col_bad = []
    head_cells = pt['cells_n'][hi] if pt is not None else None
    b_head = bt['cells_n'][bhi] if bt is not None else None
    if head_cells is not None:
        if head_cells < 3:
            col_bad.append('表头行只有 %d 个单元格，三列对应关系已错位' % head_cells)
        elif b_head is not None and b_head >= 3 and head_cells < b_head:
            col_bad.append('表头行单元格数 %d→%d，列被合并掉' % (b_head, head_cells))
    penal(bool(col_bad), 1,
          '表格列数或单元格合并关系被改动，导致同一行内学生活动、教师活动、评价要点错位',
          '；'.join(col_bad) if col_bad
          else ('学习过程表格表头 %s 个单元格（交付前 %s），各行对应关系未变'
                % (head_cells, b_head if b_head is not None else '未知') if pt is not None
                else '未定位到学习过程表格'),
          unknown=(pt is None or bt is None))

    # —— 减分项 27：交付文件名与文档内"课时主题"不一致 ——
    core = _theme_core(f['theme'])
    stem = _ws(os.path.splitext(name)[0])
    hit29 = bool(core) and core not in stem
    if hit29 and core:
        same = sum(1 for ch in set(core) if ch in stem)
        hit29 = (same / float(len(set(core)))) < 0.6
    penal(hit29, 1, '交付文件名与文档内"课时主题"不一致',
          '文件名"%s"／课时主题"%s"' % (stem, core or '未识别'), unknown=(not core))

    return f, b, blf, results, details


def _parse_items(results, details, max_deltas):
    """把"+N：说明"一类判定串反解成结构化条目（excluded 仅内部使用，不对外输出）。"""
    out = []
    for raw, det, mx in zip(results, details, max_deltas):
        head, _, rule = raw.partition('：')
        head = head.strip()
        excluded = head.startswith('=')
        try:
            delta = int(head.replace('+', '').replace('=', '').strip())
        except ValueError:
            delta = 0
        out.append({'rule': rule, 'max_delta': 0 if excluded else mx, 'delta': delta,
                    'hit': (not excluded) and delta != 0, 'detail': det,
                    'excluded': excluded})
    return out


def _file_key(file_name):
    """条目前缀用的短文件标识：取文件名开头的两位序号（01—08），
       取不到时退回不含扩展名的文件名，保证仍能看出判的是哪一份。"""
    m = re.match(r'\s*(\d{1,2})', file_name or '')
    return m.group(1).zfill(2) if m else os.path.splitext(file_name or '')[0]


def _public_items(items, file_name=''):
    """对外输出的条目：rule/max_delta/delta/hit 四个字段（不输出 detail）；rule 取
       _RULE_LABELS 的简写并以 [01]—[08] 前缀标明判的是哪一份交付文档（完整文件名见
       顶层 file_name）。
       "交付前已满足、本条不计分也不扣分"的条目（分值已从该文件满分基数中扣除）
       不再对外列出，避免报告里出现一排 0 分条目；它们的 max_delta 与 delta 均为 0，
       故删去后 max_score 与 total_score 都不受影响。"""
    key = _file_key(file_name)
    out = []
    for i, it in enumerate(items):
        if it.get('excluded'):
            continue
        label = _RULE_LABELS[i] if i < len(_RULE_LABELS) else it['rule']
        out.append({'rule': ('[%s] %s' % (key, label)) if key else label,
                    'max_delta': it['max_delta'], 'delta': it['delta'], 'hit': it['hit']})
    return out


def evaluate_file(file_path):
    """判定单份交付文档，返回与目录级一致的结构（不含目录级末条）。"""
    rec = dict(id=SCRIPT_ID, file_name=os.path.basename(file_path), status='ok', error=None,
               dim1_pass=False, dim1_reason='', dim2_items=[], total_score=0, max_score=0)
    try:
        f, b, blf, results, details = _evaluate_impl(file_path)
    except Exception as exc:
        rec['status'] = 'error'
        rec['error'] = '%s: %s' % (type(exc).__name__, exc)
        # 维度一未通过时，dim1_reason 必须带上错误摘要，便于直接定位原因
        rec['dim1_reason'] = '文档无法打开或解析，完成度记 0：%s: %s' % (type(exc).__name__, exc)
        return rec
    ok, why = _dim1(f)
    base = (os.path.basename(b['path']) if b is not None
            else ('内置交付前基线' if blf is not None else None))
    rec['dim1_pass'] = ok
    rec['dim1_reason'] = '%s（对照交付前基线：%s）' % (why, base or '未找到，加分项按未达成计')
    items = _parse_items(results, details, _MAX_DELTAS_ORDER[:_N_FILE_RULES])
    if not ok:
        for it in items:
            it['delta'], it['hit'] = 0, False
    rec['dim2_items'] = _public_items(items, rec['file_name'])
    rec['max_score'] = sum(it['max_delta'] for it in items if it['max_delta'] > 0)
    rec['total_score'] = sum(it['delta'] for it in items)
    for k in ('total_score', 'max_score'):   # 把总分挪到该文件记录的最后
        rec[k] = rec.pop(k)
    return rec


def evaluate(dir_path=None):
    """判定一次交付：本任务一次交付 8 份《课时学习方案》，细则按文件累计，
       故对目录内全部 .docx 逐份判定后按条目求和；各文件满分基数因"交付前已成立
       的条目不计分"而不等，max_score 即为逐文件核定后的满分之和。
       dir_path 省略时自动取脚本所在目录（见 _default_target），即把脚本放进
       交付目录直接运行，就核查该目录下的输出文件。"""
    if dir_path is None:
        dir_path = _default_target()
    out = dict(id=SCRIPT_ID, file_name='', status='ok', error=None, dim1_pass=False,
               dim1_reason='', dim2_items=[], total_score=0, max_score=0)
    if not dir_path or not os.path.isdir(dir_path):
        # 找不到被评估文件：error 只回传纯文本原因；维度一根本没进入判定，
        # 故 dim1_pass 与 dim1_reason 一律留空（空字符串，而非 false），
        # dim2 不出项，总分与满分一并归零。
        out['status'] = 'error'
        out['error'] = '目录不存在或不是目录：%s' % dir_path
        out['dim1_pass'] = ''
        return out
    cands = _delivery_files(dir_path)
    if not cands:
        out['status'] = 'error'
        out['error'] = '目录中未找到 .docx 文件：%s' % dir_path
        out['dim1_pass'] = ''
        return out
    names = [os.path.basename(p) for p in cands]
    out['file_name'] = '；'.join(names)
    recs = [evaluate_file(p) for p in cands]
    usable = [r for r in recs if r['dim1_pass']]
    bad_names = [r['file_name'] for r in recs if not r['dim1_pass']]
    out['dim1_pass'] = bool(usable)
    if usable:
        out['dim1_reason'] = ('%d/%d 份文档可正常打开、正文与表格可单独点选编辑%s'
                              % (len(usable), len(recs),
                                 ('；以下文件维度一未通过、完成度记 0：%s' % '、'.join(bad_names))
                                 if bad_names else ''))
    else:
        # 维度一未通过：摘要里逐份带上各文件的具体原因，便于直接定位
        brief = '；'.join('%s：%s' % (r['file_name'], r['dim1_reason'] or (r['error'] or '未通过'))
                          for r in recs)
        out['dim1_reason'] = ('%d 份文档均不可用（无法打开、正文变为图片或版面严重错乱），'
                              '完成度总分记 0。%s' % (len(recs), brief))

    # 明细逐文件逐条列出：rule 前缀 [01]—[08] 标明判的是哪一份交付文档。
    # 交付前该条已成立（无需动手即已满足）的条目，max_delta 记 0，即从该文件满分基数中扣除。
    for r in recs:
        if r['dim2_items']:
            out['dim2_items'].extend(r['dim2_items'])
        else:   # 该文件解析失败：仍逐条列出，按完成度记 0
            for label, mx in zip(_RULE_LABELS[:_N_FILE_RULES],
                                 _MAX_DELTAS_ORDER[:_N_FILE_RULES]):
                out['dim2_items'].append(dict(
                    rule='[%s] %s' % (_file_key(r['file_name']), label),
                    max_delta=mx, delta=0, hit=False))
    # 末条为目录级：交付份数与重名
    stems = [re.match(r'\s*(\d{2})', n) for n in names]
    keys = [s.group(1) for s in stems if s]
    dup = sorted({k for k in keys if keys.count(k) > 1})
    hit_last = len(cands) < 8 or bool(dup) or len(set(names)) != len(names)
    out['dim2_items'].append(dict(
        rule='[交付目录] %s' % _RULE_LABELS[_N_FILE_RULES],
        max_delta=-1, delta=-1 if hit_last else 0, hit=hit_last))
    out['max_score'] = sum(it['max_delta'] for it in out['dim2_items'] if it['max_delta'] > 0)
    out['total_score'] = sum(it['delta'] for it in out['dim2_items'])
    if not out['dim1_pass']:
        out['total_score'] = 0
    return out


def _dumps_compact(out):
    """备用的紧凑打印：合法 JSON，但 dim2_items 每条占一行（约 240 行）。
       默认打印用 indent=2 的常规缩进（约 1400 行）；若显示端仍按长度截断，
       可改调用本函数。"""
    parts = []
    for k, v in out.items():
        key = json.dumps(k, ensure_ascii=False)
        if k == 'dim2_items':
            body = ',\n    '.join(json.dumps(it, ensure_ascii=False) for it in v)
            parts.append('  %s: [\n    %s\n  ]' % (key, body) if v else '  %s: []' % key)
        else:
            parts.append('  %s: %s' % (key, json.dumps(v, ensure_ascii=False)))
    return '{\n%s\n}' % ',\n'.join(parts)


if __name__ == '__main__':
    # 不带参数：核查脚本所在目录下的交付文档；带参数：核查指定目录或单个文件。
    _target = sys.argv[1] if len(sys.argv) > 1 else _default_target()
    if os.path.isdir(_target):
        _out = evaluate(_target)
    else:
        _out = evaluate_file(_target)
    print(json.dumps(_out, ensure_ascii=False, indent=2))
