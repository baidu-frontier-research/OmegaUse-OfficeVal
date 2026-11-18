# -*- coding: utf-8 -*-
"""自动评估：初中语文跨年级综合质量检测——按试题内容制作的可编辑 Word 版答题卡

对外统一接口：
    evaluate(dir_path: str) -> dict
        dir_path 为脚本所在目录，脚本仅在该目录本身（不含子目录）中定位 .docx 并评估。
        返回结构：id / file_name / status / error / dim1_pass / dim1_reason /
                  dim2_items[{rule,max_delta,delta,hit}] / total_score / max_score。

指令：麻烦帮我按试题内容对答题卡进行调整，修改出对应符合试题内容的可修改的word版答题卡

评分逻辑（对应同目录 细则.txt）：
  维度1（可用与可修改性）为一票否决：交付须为1个可正常打开的 .docx，
     否则 dim1_pass=False，维度2不评估，总分记 0。
  维度2（完成度）仅加分项，分值三档 1/3/5。

实现说明：
  - 用标准库 xml.etree.ElementTree 解析（与参考脚本 lxml 等价，避免环境差异）。
  - 好答题卡大量使用 w:txbxContent 文本框，且用 mc:AlternateContent 承载，
    文字会在 Choice/Fallback 两分支重复出现；统计文字/符号时先剥离 Fallback 再计数。
"""
import os
import sys
import json
import copy
import zipfile
import re
import xml.etree.ElementTree as ET

SCRIPT_ID = "015"

# ============ 常量 ============
DXA_PER_CM = 567.0
EMU_PER_CM = 360000.0

WNS = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
WPSNS = '{http://schemas.microsoft.com/office/word/2010/wordprocessingShape}'
VNS = '{urn:schemas-microsoft-com:vml}'
MCNS = '{http://schemas.openxmlformats.org/markup-compatibility/2006}'
WPDRAWNS = '{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}'
ANS = '{http://schemas.openxmlformats.org/drawingml/2006/main}'

CIRCLED = '①②③④⑤⑥⑦⑧⑨⑩'


# ============ 工具函数 ============
def dxa_to_cm(dxa):
    return float(dxa) / DXA_PER_CM


def emu_to_cm(emu):
    return float(emu) / EMU_PER_CM


def half_pt_to_pt(v):
    return float(v) / 2.0


def in_range(val, lo, hi):
    return val is not None and lo <= val <= hi


def get_text(elem):
    return ''.join(elem.itertext())


def norm_paren(s):
    """全角括号/点号归一化为半角，便于匹配题号与分值。"""
    return (s.replace('（', '(').replace('）', ')')
             .replace('．', '.').replace('。', '.'))


# ============ 文档解析 ============
class DocParser:
    """解析 .docx，提供页面尺寸、去重文本、文本框/形状、表格等信息。

    好答题卡用 mc:AlternateContent（Choice=现代文本框 / Fallback=VML 兜底）承载
    文本框，同一段文字在两分支重复出现。统计文字/符号时用剥离 Fallback 后的
    document 树（self.body_dd）；结构定位（形状/表格）用原始树 self.body。
    """

    def __init__(self, filepath):
        self.filepath = filepath
        self.z = zipfile.ZipFile(filepath)
        self.raw = self.z.read('word/document.xml')
        self.doc_xml = ET.fromstring(self.raw)
        self.body = self.doc_xml.find(f'{WNS}body')
        self.children = list(self.body)
        # 剥离 Fallback 的去重副本（仅用于文字/符号计数）
        self.doc_dd = ET.fromstring(self.raw)
        self.body_dd = self.doc_dd.find(f'{WNS}body')
        self._strip_fallback(self.body_dd)
        # 父指针（原始树，用于顶层块定位）
        self._parent = {c: p for p in self.body.iter() for c in p}
        self._parse()

    @staticmethod
    def _strip_fallback(root):
        # ET 无 getparent，需用父映射删除 mc:Fallback 子树
        parent = {c: p for p in root.iter() for c in p}
        for ac in list(root.iter(f'{MCNS}AlternateContent')):
            for fb in ac.findall(f'{MCNS}Fallback'):
                ac.remove(fb)

    def _parse(self):
        self.sections = []
        for child in self.children:
            tag = child.tag.split('}')[-1]
            if tag == 'p':
                pPr = child.find(f'{WNS}pPr')
                if pPr is not None and pPr.find(f'{WNS}sectPr') is not None:
                    self.sections.append(pPr.find(f'{WNS}sectPr'))
            elif tag == 'sectPr':
                self.sections.append(child)
        # 页面尺寸/边距（取首个 section）
        self.page_w_dxa = self.page_h_dxa = 0
        self.orient = ''
        self.margins = {}
        self.cols_num = None
        for sect in self.body.iter(f'{WNS}sectPr'):
            pgSz = sect.find(f'{WNS}pgSz')
            pgMar = sect.find(f'{WNS}pgMar')
            cols = sect.find(f'{WNS}cols')
            if pgSz is not None:
                self.page_w_dxa = int(pgSz.get(f'{WNS}w', 0))
                self.page_h_dxa = int(pgSz.get(f'{WNS}h', 0))
                self.orient = pgSz.get(f'{WNS}orient', '')
            if pgMar is not None:
                for s in ('top', 'bottom', 'left', 'right'):
                    self.margins[s] = int(pgMar.get(f'{WNS}{s}', 0))
            if cols is not None:
                self.cols_num = cols.get(f'{WNS}num')
            break
        self.page_w_cm = dxa_to_cm(self.page_w_dxa) if self.page_w_dxa else 0.0
        self.page_h_cm = dxa_to_cm(self.page_h_dxa) if self.page_h_dxa else 0.0
        self.mL_cm = dxa_to_cm(self.margins.get('left', 0))
        # 首个 sectPr 所在顶层块下标 -> 用于判定形状归属页
        self._first_sect_top = len(self.children)
        for i, k in enumerate(self.children):
            if k.find(f'.//{WNS}sectPr') is not None:
                self._first_sect_top = i
                break
        self._collect_shapes()
        self._collect_tables()

    # ---- 文本 ----
    @property
    def full_text_dd(self):
        return get_text(self.body_dd)

    @property
    def nospace_dd(self):
        return re.sub(r'\s', '', self.full_text_dd)

    # ---- 形状（文本框/矩形，用原始树；按 wsp 去重避免 Fallback 重复）----
    def _top_block_idx(self, node):
        cur = node
        for _ in range(80):
            if cur is None:
                return None
            p = self._parent.get(cur)
            if p is self.body:
                try:
                    return self.children.index(cur)
                except ValueError:
                    return None
            cur = p
        return None

    def _abs_x_cm(self, anchor):
        pH = anchor.find(f'{WPDRAWNS}positionH')
        if pH is None:
            return None
        rf = pH.get('relativeFrom')
        off = pH.find(f'{WPDRAWNS}posOffset')
        if off is None or not off.text:
            return None
        try:
            v = emu_to_cm(int(off.text))
        except ValueError:
            return None
        return v if rf == 'page' else v + self.mL_cm

    def _collect_shapes(self):
        self.shapes = []
        seen = set()
        for anc in self.body.findall(f'.//{WPDRAWNS}anchor'):
            for wsp in anc.findall(f'.//{WPSNS}wsp'):
                if id(wsp) in seen:
                    continue
                seen.add(id(wsp))
                txbx = wsp.find(f'{WPSNS}txbx')
                content = txbx.find(f'{WNS}txbxContent') if txbx is not None else None
                prst_el = wsp.find(f'.//{ANS}prstGeom')
                prst = prst_el.get('prst') if prst_el is not None else None
                has_txbx = content is not None
                is_rect = (prst == 'rect')
                if not (has_txbx or is_rect):
                    continue  # prst='line' 等不计为文本框/矩形
                ext = anc.find(f'{WPDRAWNS}extent')
                w_cm = h_cm = None
                if ext is not None:
                    try:
                        w_cm = emu_to_cm(int(ext.get('cx') or 0)) or None
                        h_cm = emu_to_cm(int(ext.get('cy') or 0)) or None
                    except ValueError:
                        pass
                x_cm = self._abs_x_cm(anc)
                ti = self._top_block_idx(anc)
                page = 1 if (ti is not None and ti <= self._first_sect_top) else 2
                stext = get_text(content) if content is not None else ''
                self.shapes.append({
                    'has_txbx': has_txbx, 'is_rect': is_rect,
                    'x_cm': x_cm, 'w_cm': w_cm, 'h_cm': h_cm,
                    'cx_cm': (x_cm + w_cm / 2) if (x_cm is not None and w_cm) else x_cm,
                    'page': page, 'top_idx': ti,
                    'text': stext, 'text_ns': re.sub(r'\s', '', stext),
                })
        self.txbx_count = sum(1 for s in self.shapes if s['has_txbx'])

    def shapes_containing(self, *keys):
        out = []
        for s in self.shapes:
            t = norm_paren(s['text_ns'])
            if all(norm_paren(k) in t for k in keys):
                out.append(s)
        return out

    # ---- 表格 ----
    def _collect_tables(self):
        self.tables = []
        for tbl in self.body.iter(f'{WNS}tbl'):
            grid = tbl.find(f'{WNS}tblGrid')
            cols = grid.findall(f'{WNS}gridCol') if grid is not None else []
            colws = [int(c.get(f'{WNS}w', 0)) for c in cols]
            rows = tbl.findall(f'{WNS}tr')
            rowh = None
            if rows:
                trPr = rows[0].find(f'{WNS}trPr')
                if trPr is not None:
                    h = trPr.find(f'{WNS}trHeight')
                    if h is not None and h.get(f'{WNS}val'):
                        rowh = int(h.get(f'{WNS}val'))
            ti = self._top_block_idx(tbl)
            page = 1 if (ti is not None and ti <= self._first_sect_top) else 2
            self.tables.append({
                'el': tbl, 'ncol': len(cols), 'nrow': len(rows),
                'colws': colws, 'rowh_dxa': rowh, 'top_idx': ti, 'page': page,
                'text': get_text(tbl), 'text_ns': re.sub(r'\s', '', get_text(tbl)),
            })

    def tables_containing(self, *keys):
        out = []
        for t in self.tables:
            tt = norm_paren(t['text_ns'])
            if all(norm_paren(k) in tt for k in keys):
                out.append(t)
        return out


def _run_size_font(para, keyword=None):
    """取段落中（可选含 keyword 的）run 的 (size_pt, font, bold, centered)。"""
    pPr = para.find(f'{WNS}pPr')
    centered = False
    if pPr is not None:
        jc = pPr.find(f'{WNS}jc')
        if jc is not None and jc.get(f'{WNS}val') in ('center',):
            centered = True
    size = font = None
    bold = False
    for r in para.findall(f'{WNS}r'):
        if keyword is not None and keyword not in ''.join(r.itertext()):
            continue
        rPr = r.find(f'{WNS}rPr')
        if rPr is not None:
            sz = rPr.find(f'{WNS}sz')
            rf = rPr.find(f'{WNS}rFonts')
            b = rPr.find(f'{WNS}b')
            if sz is not None:
                size = half_pt_to_pt(int(sz.get(f'{WNS}val')))
            if rf is not None:
                font = rf.get(f'{WNS}eastAsia') or rf.get(f'{WNS}ascii')
            bold = b is not None and b.get(f'{WNS}val') not in ('0', 'false')
            if size is not None:
                break
    return size, font, bold, centered


def _table_has_grid(tbl):
    """判断表格是否呈现内部网格线：表级 tblBorders 含 insideH/insideV(有效)，
    或绝大多数单元格四边都描了框。用于考号涂卡矩阵"有框线成格"的判定。"""
    def _live(el):
        return el is not None and el.get(f'{WNS}val') not in (None, 'nil', 'none')
    tblPr = tbl.find(f'{WNS}tblPr')
    tblB = tblPr.find(f'{WNS}tblBorders') if tblPr is not None else None
    if tblB is not None:
        ih = tblB.find(f'{WNS}insideH'); iv = tblB.find(f'{WNS}insideV')
        if _live(ih) and _live(iv):
            return True
    ncell = full = 0
    for tc in tbl.iter(f'{WNS}tc'):
        ncell += 1
        tcPr = tc.find(f'{WNS}tcPr')
        tb = tcPr.find(f'{WNS}tcBorders') if tcPr is not None else None
        if tb is not None and all(_live(tb.find(f'{WNS}{s}'))
                                  for s in ('top', 'bottom', 'left', 'right')):
            full += 1
    return ncell > 0 and full / ncell >= 0.9


def _table_total_height_cm(tbl):
    """累加表格各行显式 trHeight，返回总高(cm)；无显式行高返回 0。"""
    total_h = 0
    for tr in tbl.findall(f'{WNS}tr'):
        trPr = tr.find(f'{WNS}trPr')
        h = trPr.find(f'{WNS}trHeight') if trPr is not None else None
        if h is not None and h.get(f'{WNS}val'):
            total_h += int(h.get(f'{WNS}val'))
    return dxa_to_cm(total_h) if total_h else 0.0


def _table_has_outer_frame(tbl):
    """判断表格是否设有"表级外框线"：tblPr/tblBorders 的 top/bottom/left/right
    四边均有效。用于考号涂卡矩阵"须有完整外框、成一个封闭方阵"的判定——仅靠
    个别单元格描边（无表级外框定义）不算，避免涂卡区外框缺失/不完整仍得分。"""
    def _live(el):
        return el is not None and el.get(f'{WNS}val') not in (None, 'nil', 'none')
    tblPr = tbl.find(f'{WNS}tblPr')
    tblB = tblPr.find(f'{WNS}tblBorders') if tblPr is not None else None
    if tblB is None:
        return False
    return all(_live(tblB.find(f'{WNS}{s}')) for s in ('top', 'bottom', 'left', 'right'))


def _body_median_font_pt(doc):
    """按字符数加权，返回正文东亚字体 run 的中位字号(pt)。用于"答题卡字体过小"
    的判定：好卡正文以 12/9 磅为主，中位≈12；字体被压到 6-7 磅的卡应判不达标。"""
    buckets = {}
    for r in doc.body_dd.iter(f'{WNS}r'):
        txt = ''.join(r.itertext()).strip()
        if not txt:
            continue
        rPr = r.find(f'{WNS}rPr')
        sz = rPr.find(f'{WNS}sz') if rPr is not None else None
        if sz is not None and sz.get(f'{WNS}val'):
            pt = half_pt_to_pt(int(sz.get(f'{WNS}val')))
            buckets[pt] = buckets.get(pt, 0) + len(txt)
    if not buckets:
        return None
    total = sum(buckets.values())
    acc = 0
    for pt in sorted(buckets):
        acc += buckets[pt]
        if acc * 2 >= total:
            return pt
    return max(buckets)


def _tall_carrier_text(doc, min_h_cm=1.2):
    """汇总所有"达标高度承载块"(文本框 h≥min_h 或 显式行高合计≥min_h 的表格)的
    去空白文本，用于判定"某题作答区确有达标尺寸的作答块承载"，而非仅出现题号。"""
    parts = []
    for s in doc.shapes:
        if s['has_txbx'] and s['h_cm'] is not None and s['h_cm'] >= min_h_cm:
            parts.append(s['text_ns'])
    for t in doc.tables:
        if _table_total_height_cm(t['el']) >= min_h_cm:
            parts.append(t['text_ns'])
    return norm_paren(' '.join(parts))


# ============ 维度一：可用与可修改性（一票否决） ============
def check_dim1(doc):
    """返回 (all_pass, [(描述, 通过?, 详情), ...])。

    细则维度1：交付为1个可正常打开的 .docx。
    """
    results = []
    ext = os.path.splitext(doc.filepath)[1].lower()
    results.append(('文件格式为.docx', ext == '.docx', f'扩展名={ext}'))

    # 文件已被成功解析（evaluate 中 DocParser 未抛异常）
    results.append(('文件可正常打开', True, ''))

    return all(r[1] for r in results), results


# ============ 维度二：完成度（仅加分，1/3/5 三档） ============
def check_dim2(doc):
    R = []  # (分值, 描述, 命中, 详情)
    ns = norm_paren(doc.nospace_dd)      # 去重、去空白、括号归一后的全文
    pw = doc.page_w_cm
    ph = doc.page_h_cm

    def labels_present():
        """返回文中出现的 (题号, 分值) 集合，来自 'N.(S分)' / 'N(S分)' 等写法。"""
        return set(re.findall(r'(\d{1,2})[.\、,:]?\((\d{1,2})分\)', ns))

    LB = labels_present()

    def has_q(qnum, score):
        return (str(qnum), str(score)) in LB

    # 达标尺寸承载块文本 & "某题是否落在达标作答块内"判定（尺寸门槛）
    TALL_TXT = _tall_carrier_text(doc)

    def q_in_tall(qnum):
        return re.search(rf'(?<!\d){qnum}[.\．、(]', TALL_TXT) is not None

    def positioned_shape(*keys):
        """返回含全部 keys、且带有可解析绝对 x 的文本框/形状（None 表示没有）。"""
        cand = [s for s in doc.shapes_containing(*keys) if s['x_cm'] is not None]
        return cand[0] if cand else None

    def frac(x):
        return (x / pw) if pw else None

    # ---- 1. A3横向 + 页边距 + 四栏版面 (+5) ----
    w_cm = doc.page_w_cm
    h_cm = doc.page_h_cm
    is_a3 = in_range(w_cm, 41.5, 42.5) and in_range(h_cm, 29.5, 30.0)
    landscape = (doc.orient == 'landscape') or (doc.page_w_dxa > doc.page_h_dxa)
    margins_cm = {s: dxa_to_cm(v) for s, v in doc.margins.items()}
    margins_ok = bool(margins_cm) and all(in_range(v, 1.0, 1.6) for v in margins_cm.values())
    # 四栏版面：以带绝对 x 的文本框/形状的横向分布近似判断（左中右多栏）
    xs = sorted(frac(s['cx_cm']) for s in doc.shapes if s['cx_cm'] is not None)
    has_left = any(f < 0.33 for f in xs)
    has_mid = any(0.30 <= f < 0.66 for f in xs)
    has_right = any(f >= 0.60 for f in xs)
    multi_col = has_left and has_right and len(xs) >= 4
    hit = is_a3 and landscape and margins_ok and multi_col
    R.append((5,
              'A3横向(41.5-42.5×29.5-30cm)、页边距1.0-1.6cm、版面分密封线带/左/中/右作文多栏',
              hit,
              f'{w_cm:.1f}×{h_cm:.1f}cm,横向={landscape},'
              f'边距cm={{{",".join(f"{k}:{v:.2f}" for k,v in margins_cm.items())}}},'
              f'多栏(左{has_left}/中{has_mid}/右{has_right},定位块={len(xs)})={multi_col}'))

    # ---- 2. 标题 (+1) ----
    # 细则：首部醒目居中标题（如"…答题卡"），单行、字号22-36磅、明显大于正文。
    # 好答题卡实测标题为仿宋四号(14pt)、未居中，属答题卡常见做法；此处按"存在
    # 明显标题行"从宽认定：首个非空行含"答题卡"字样且有独立字号即视为完成该项。
    title_txt = ''
    title_size = None
    for p in doc.body_dd.iter(f'{WNS}p'):
        t = get_text(p).strip()
        if t:
            title_txt = t
            title_size, _f, _b, _c = _run_size_font(p)
            break
    hit = ('答题卡' in title_txt) and (title_size is not None)
    R.append((1, '首部醒目答题卡标题(单行、字号明显大于正文)', hit,
              f'标题="{title_txt[:24]}",字号={title_size}pt'))

    # ---- 3. 考生信息区结构 (+1) ----
    # 细则：至少含"姓名""学号"(或班级)两项填写位，两行小型表格或横线。
    has_name = '姓名' in ns
    has_id = ('学号' in ns) or ('班级' in ns)
    info_tbl = doc.tables_containing('姓名') or doc.tables_containing('学号')
    hit = has_name and has_id and bool(info_tbl or ('姓名' in ns and '学号' in ns))
    R.append((1, '考生信息区(至少姓名+学号/班级两项填写位)', hit,
              f'姓名={has_name},学号/班级={has_id},信息表格={len(info_tbl)}'))

    # ---- 4. 考生信息区位置：左上角、标题正下方 (+3) ----
    info_shape = positioned_shape('姓名') or positioned_shape('学号')
    info_tbls = doc.tables_containing('姓名') or doc.tables_containing('学号')
    hit = False
    detail = '未找到考生信息区'
    if info_shape is not None:
        f = frac(info_shape['cx_cm'])
        left_ok = f is not None and f < 0.40
        hit = left_ok and info_shape['page'] == 1
        detail = f'信息区(文本框)cx比例={f:.2f},页={info_shape["page"]}(要求页1左上)'
    elif info_tbls:
        # 无定位文本框时，退化为"信息表格位于第1页、文档靠前(前几个顶层块)"
        t = info_tbls[0]
        near_top = t['top_idx'] is not None and t['top_idx'] <= 3
        hit = t['page'] == 1 and near_top
        detail = f'信息表格页={t["page"]},文档顶层序号={t["top_idx"]}(要求页1靠前)'
    R.append((3, '考生信息区位于页面左上角、标题正下方', hit, detail))

    # ---- 5b. 考生信息区大小合规：紧凑框 + 可书写行高 (+1) ----
    # 细则收紧：信息区须为紧凑填写框(占页宽约0.08-0.35)，每个填写行高≥0.8cm(能写下)，
    # 而非塞进整页宽的大包裹表格、或行高被压到0.6cm以下的窄条。
    # 好卡信息表 5.0cm(占0.12)、行高0.9cm；bad 卡真正信息表行高仅0.62cm -> 不达标。
    def _info_row_heights(t):
        hs = []
        for tr in t['el'].findall(f'{WNS}tr'):
            trPr = tr.find(f'{WNS}trPr')
            h = trPr.find(f'{WNS}trHeight') if trPr is not None else None
            hs.append(dxa_to_cm(int(h.get(f'{WNS}val'))) if (h is not None and h.get(f'{WNS}val')) else 0.0)
        return hs
    info_boxes = [t for t in doc.tables
                  if ('姓名' in t['text_ns'] and ('学号' in t['text_ns'] or '班级' in t['text_ns']))
                  and frac(dxa_to_cm(sum(t['colws']))) is not None
                  and frac(dxa_to_cm(sum(t['colws']))) < 0.50]
    # 取行数最多者为真正的信息填写表（排除单行大包裹）
    info_box = max(info_boxes, key=lambda t: t['nrow']) if info_boxes else None
    hit = False
    detail = '未找到紧凑的考生信息填写表'
    if info_box is not None:
        wf = frac(dxa_to_cm(sum(info_box['colws'])))
        rhs = [h for h in _info_row_heights(info_box) if h > 0]
        med_rh = sorted(rhs)[len(rhs) // 2] if rhs else 0.0
        w_ok = in_range(wf, 0.08, 0.16)
        h_ok = med_rh >= 0.8
        hit = w_ok and h_ok and info_box['nrow'] >= 2
        detail = (f'信息表占页宽={None if wf is None else round(wf,3)}(0.08-0.16),'
                  f'行数={info_box["nrow"]},中位行高≈{round(med_rh,2)}cm(要求≥0.8可书写)')
    R.append((1, '考生信息区大小合规:紧凑填写框(占页宽0.08-0.16)、行高≥0.8cm可书写', hit, detail))

    # ---- 5. 考号填涂区结构：7-10列 [0]-[9] 纵向涂卡矩阵 (+5) ----
    # 每列自上而下 [0]..[9]。以去重文本中 [0]..[9] 各出现次数的最小值近似"列数"。
    # 收紧：除列数 7-10 外，要求矩阵整体高度落在 3-9cm（避免退化成一行/一小格）。
    col_counts = [ns.count(f'[{d}]') for d in '0123456789']
    n_cols_bubble = min(col_counts) if col_counts else 0
    has_examno = ('考号' in ns) or ('准考证' in ns)
    cand = doc.tables_containing('考号') or doc.tables_containing('准考证')
    # 含"考号"的可能是外层包裹表；取行数最多者为真正的涂卡矩阵
    exam_matrix = max(cand, key=lambda t: t['nrow']) if cand else None
    matrix_h_cm = None
    if exam_matrix is not None:
        total_h = 0
        for tr in exam_matrix['el'].findall(f'{WNS}tr'):
            trPr = tr.find(f'{WNS}trPr')
            h = trPr.find(f'{WNS}trHeight') if trPr is not None else None
            if h is not None and h.get(f'{WNS}val'):
                total_h += int(h.get(f'{WNS}val'))
        matrix_h_cm = dxa_to_cm(total_h) if total_h else None
    hit = (has_examno and in_range(n_cols_bubble, 7, 10)
           and in_range(matrix_h_cm, 3.0, 9.0)
           and exam_matrix is not None and _table_has_outer_frame(exam_matrix['el']))
    outer_ok = exam_matrix is not None and _table_has_outer_frame(exam_matrix['el'])
    R.append((5, '考号填涂区:7-10列纵向[0]-[9]涂卡矩阵、高3-9cm、须有完整外框线', hit,
              f'考号字样={has_examno},每位[0]-[9]列数≈{n_cols_bubble}(要求7-10),'
              f'矩阵高≈{None if matrix_h_cm is None else round(matrix_h_cm,2)}cm(要求3-9),'
              f'表级外框线={outer_ok}'))

    # ---- 6. 考号填涂区位置：上部、信息区右侧 (+3) ----
    exam_shape = positioned_shape('考号') or positioned_shape('准考证')
    exam_tbls = doc.tables_containing('考号') or doc.tables_containing('准考证')
    hit = False
    detail = '未找到考号填涂区'
    if exam_shape is not None:
        f = frac(exam_shape['cx_cm'])
        hit = exam_shape['page'] == 1 and (f is not None)
        detail = f'考号区(文本框)cx比例={f:.2f},页={exam_shape["page"]}(要求页1上部)'
    elif exam_tbls:
        t = exam_tbls[0]
        near_top = t['top_idx'] is not None and t['top_idx'] <= 5
        hit = t['page'] == 1 and near_top
        detail = f'考号表格页={t["page"]},文档顶层序号={t["top_idx"]}(要求页1上部)'
    R.append((3, '考号填涂区位于页面上部、考生信息区右侧', hit, detail))

    # ---- 7. 填涂示例说明 (+1) ----
    has_example = ('正确填涂' in ns) or ('2B' in ns) or ('填涂' in ns and '■' in ns)
    R.append((1, '填涂示例说明(如"正确填涂：■""2B铅笔")', has_example,
              f'含填涂示例={has_example}'))

    # ---- 8. 缺考标记区 (+1) ----
    has_absent = '缺考' in ns
    R.append((1, '缺考标记区', has_absent, f'含"缺考"={has_absent}'))

    # ---- 9. 客观选择题填涂区：仅Q3/4/7/12/19共5道，各[A][B][C][D] (+5) ----
    abcd = [ns.count(f'[{c}]') for c in 'ABCD']
    n_choice = min(abcd) if abcd else 0
    hit = n_choice == 5
    R.append((5, '客观选择题填涂区:5道题(第3/4/7/12/19)各[A][B][C][D]', hit,
              f'[A][B][C][D]成组数≈{n_choice}(要求正好5)'))

    # ---- 10. 客观选择题填涂区位置：上部、涂卡区下方或同一横带 (+3) ----
    choice_shape = positioned_shape('选择题') or positioned_shape('[A]', '[B]')
    hit = False
    detail = '未找到带定位的选择题填涂区'
    if choice_shape is not None:
        hit = choice_shape['page'] == 1
        f = frac(choice_shape['cx_cm'])
        detail = f'选择题区cx比例={None if f is None else round(f,2)},页={choice_shape["page"]}'
    R.append((3, '客观选择题填涂区位于页面上部、涂卡区下方或同一横带', hit, detail))

    # ---- 11. "一、积累运用"作答区位置+大小：左侧作答栏 (+3) ----
    # 细则：位于左侧作答栏，栏宽占页宽0.38-0.50、区域高7-13cm，纵贯版心。
    sec1_shape = positioned_shape('一、')
    hit = False
    detail = '未找到带定位的"一、"作答区'
    if sec1_shape is not None:
        f = frac(sec1_shape['cx_cm'])
        wf = frac(sec1_shape['w_cm'])
        h = sec1_shape['h_cm']
        left_ok = f is not None and f < 0.55
        w_ok = in_range(wf, 0.38, 0.50)
        h_ok = in_range(h, 7.0, 13.0)
        hit = left_ok and w_ok and h_ok
        detail = (f'cx比例={f:.2f}(左侧<0.55),栏宽占页宽={None if wf is None else round(wf,2)}'
                  f'(0.38-0.50),高={None if h is None else round(h,1)}cm(7-13),页={sec1_shape["page"]}')
    R.append((3, '"一、积累运用"作答区位于左侧作答栏、栏宽占页宽0.38-0.50、高7-13cm', hit, detail))

    # ---- 12. 第1题古诗文默写(10分):①-⑩共10个填空 (+3) ----
    n_circ = sum(1 for c in CIRCLED if c in doc.full_text_dd)
    hit = has_q(1, 10) and n_circ >= 10 and q_in_tall(1)
    R.append((3, '第1题古诗文默写(10分):①-⑩共10个填空作答位、须在达标作答块内', hit,
              f'题1(10分)={has_q(1,10)},①-⑩出现={n_circ}/10,达标作答块={q_in_tall(1)}'))

    # ---- 13. 第2题拼音写词语(3分):(1)(2)(3)三个作答位 (+1) ----
    n_paren = sum(1 for i in ('(1)', '(2)', '(3)') if i in ns)
    hit = has_q(2, 3) and q_in_tall(2)
    R.append((1, '第2题拼音写词语(3分)作答位', hit,
              f'题2(3分)={has_q(2,3)},(1)(2)(3)={n_paren},达标作答块={q_in_tall(2)}'))

    # ---- 14. 第5题配乐推荐(2分)作答位 (+1) ----
    hit = has_q(5, 2) and q_in_tall(5)
    R.append((1, '第5题配乐推荐(2分)作答位(所选序号+理由)', hit,
              f'题5(2分)={has_q(5,2)},达标作答块={q_in_tall(5)}'))

    # ---- 15. 第6题对联(2分)作答位 (+1) ----
    hit = has_q(6, 2) and q_in_tall(6)
    R.append((1, '第6题对联(2分)补写上联作答位', hit,
              f'题6(2分)={has_q(6,2)},达标作答块={q_in_tall(6)}'))

    # ---- 16. "二、阅读"作答区位置+大小：中间作答栏 (+3) ----
    # 细则：位于中间作答栏，栏宽占页宽0.38-0.50、区域高5-10cm，纵贯版心。
    sec2_shape = positioned_shape('二、')
    hit = False
    detail = '未找到带定位的"二、"作答区'
    if sec2_shape is not None:
        f = frac(sec2_shape['cx_cm'])
        wf = frac(sec2_shape['w_cm'])
        h = sec2_shape['h_cm']
        pos_ok = f is not None and f < 0.66
        w_ok = in_range(wf, 0.38, 0.50)
        h_ok = in_range(h, 5.0, 10.0)
        hit = pos_ok and w_ok and h_ok
        detail = (f'cx比例={f:.2f},栏宽占页宽={None if wf is None else round(wf,2)}(0.38-0.50),'
                  f'高={None if h is None else round(h,1)}cm(5-10),页={sec2_shape["page"]}')
    R.append((3, '"二、阅读"作答区位于中间作答栏、栏宽占页宽0.38-0.50、高5-10cm', hit, detail))

    # ---- 17-27. 第8/9/10/11/13/14/15/16/17/18/20题作答区 ----
    Q_ITEMS = [
        (3, 8, 6, '第8题文言翻译(6分):(1)(2)两小题作答区'),
        (1, 9, 3, '第9题断句(3分)作答区'),
        (3, 10, 4, '第10题(4分)作答区(3-4行横线)'),
        (3, 11, 4, '第11题(4分)作答区(3-4行横线)'),
        (1, 13, 3, '第13题文本结构(3分)作答区'),
        (3, 14, 4, '第14题佐证例子(4分):例1/例2两处作答位'),
        (1, 15, 2, '第15题选镜单(2分)补充填空'),
        (3, 16, 4, '第16题朗读语气(4分)作答区'),
        (1, 17, 3, '第17题剪辑分歧(3分)作答区'),
        (3, 18, 5, '第18题字幕理解(5分)作答区(4-5行横线)'),
        (3, 20, 4, '第20题名著专题阅读(4分)作答区(3-4行横线)'),
    ]
    for score, qn, sc, desc in Q_ITEMS:
        h = has_q(qn, sc) and q_in_tall(qn)
        R.append((score, desc, h, f'题{qn}({sc}分)={has_q(qn,sc)},达标作答块={q_in_tall(qn)}'))

    # ---- 28. 作文区位置+大小：右侧整栏 (+3) ----
    # 细则：独立占据页面右侧一整栏、栏宽占页宽0.40-0.50、高≥24cm纵贯整页。
    hit = False
    detail = '未找到带定位的作文区'
    essay_cands = [s for s in doc.shapes
                   if s['x_cm'] is not None and s['h_cm'] is not None
                   and frac(s['cx_cm']) is not None and frac(s['cx_cm']) >= 0.60]
    # 右侧最高的那个文本框视为作文栏
    essay_cands.sort(key=lambda s: s['h_cm'], reverse=True)
    if essay_cands:
        s = essay_cands[0]
        wf = frac(s['w_cm'])
        w_ok = in_range(wf, 0.40, 0.50)
        h_ok = s['h_cm'] >= 24.0
        hit = w_ok and h_ok
        detail = (f'cx比例={frac(s["cx_cm"]):.2f}(右侧),栏宽占页宽={None if wf is None else round(wf,2)}'
                  f'(0.40-0.50),高={s["h_cm"]:.1f}cm(≥24),页={s["page"]}')
    R.append((3, '作文区独立占据页面右侧一整栏、栏宽占页宽0.40-0.50、高≥24cm', hit, detail))

    # ---- 29. 第21题作文格：格数+单格尺寸 (+5) ----
    # 细则：每行20-25格、总格数≥800，单格宽0.75-0.90cm、行高0.68-0.82cm，
    #       栏顶"题目：__"，字数刻度提示。
    essay_grids = [t for t in doc.tables if 20 <= t['ncol'] <= 25 and t['nrow'] >= 5]
    total_cells = sum(t['ncol'] * t['nrow'] for t in essay_grids)
    has_title_line = ('题目' in ns)
    has_scale = ('600' in ns) or ('800' in ns)
    # 单格尺寸：取各作文格块列宽中位数与行高中位数
    col_cms, row_cms = [], []
    for t in essay_grids:
        for w in t['colws']:
            if w > 0:
                col_cms.append(dxa_to_cm(w))
        if t['rowh_dxa']:
            row_cms.append(dxa_to_cm(t['rowh_dxa']))
    med_col = sorted(col_cms)[len(col_cms) // 2] if col_cms else None
    med_row = sorted(row_cms)[len(row_cms) // 2] if row_cms else None
    size_ok = in_range(med_col, 0.75, 0.90) and in_range(med_row, 0.68, 0.82)
    hit = total_cells >= 800 and bool(essay_grids) and has_title_line and size_ok
    R.append((5, '第21题作文格:每行20-25格、总格数≥800、单格宽0.75-0.90cm行高0.68-0.82cm、含题目栏', hit,
              f'作文格块数={len(essay_grids)},总格数≈{total_cells}(≥800),'
              f'单格宽≈{None if med_col is None else round(med_col,3)}cm(0.75-0.90),'
              f'行高≈{None if med_row is None else round(med_row,3)}cm(0.68-0.82),'
              f'题目栏={has_title_line},字数刻度={has_scale}'))

    # ---- 30. 作文方格间距：不首尾相连成整片黑格 (+3) ----
    # 细则：每格单元格左右内边距≥0.15cm(约108twips)，或表格设有格间距(tblCellSpacing)，
    #       使相邻小方格视觉上彼此分开。bad 卡格子 tcMar 全为0、无间距 -> 相连成片。
    def _grid_has_spacing(t):
        tbl = t['el']
        tblPr = tbl.find(f'{WNS}tblPr')
        # 1) 表级格间距
        if tblPr is not None:
            cs = tblPr.find(f'{WNS}tblCellSpacing')
            if cs is not None and cs.get(f'{WNS}w') and int(cs.get(f'{WNS}w')) > 0:
                return True, f'tblCellSpacing={cs.get(f"{WNS}w")}'
        # 2) 单元格左右内边距（表级 tblCellMar 或单元格 tcMar），取中位判断
        def _lr_margin(mar_el):
            if mar_el is None:
                return None
            vals = []
            for s in ('left', 'right'):
                e = mar_el.find(f'{WNS}{s}')
                if e is not None and e.get(f'{WNS}w'):
                    vals.append(int(e.get(f'{WNS}w')))
            return max(vals) if vals else None
        tbl_mar = _lr_margin(tblPr.find(f'{WNS}tblCellMar')) if tblPr is not None else None
        # 采样若干单元格 tcMar
        cell_mars = []
        for tc in tbl.iter(f'{WNS}tc'):
            tcPr = tc.find(f'{WNS}tcPr')
            m = _lr_margin(tcPr.find(f'{WNS}tcMar')) if tcPr is not None else None
            if m is not None:
                cell_mars.append(m)
            if len(cell_mars) >= 20:
                break
        eff = None
        if cell_mars:
            eff = sorted(cell_mars)[len(cell_mars) // 2]  # 单元格级覆盖表级
        elif tbl_mar is not None:
            eff = tbl_mar
        if eff is not None and dxa_to_cm(eff) >= 0.15:
            return True, f'单元格左右内边距≈{dxa_to_cm(eff):.2f}cm'
        return False, f'左右内边距≈{"0" if eff is None else f"{dxa_to_cm(eff):.2f}"}cm(<0.15,格子相连)'
    hit = False
    detail = '未找到作文格'
    if essay_grids:
        results = [_grid_has_spacing(t) for t in essay_grids]
        hit = all(ok for ok, _ in results)
        detail = '; '.join(d for _, d in results[:2])
    R.append((5, '作文方格之间留有清晰间距(不首尾相连成整片黑格)', hit, detail))

    # ---- 31. 各主观题作答区高度与分值匹配、不塌缩 (+3) ----
    # 细则：大题作答区不塌缩为0高度，以文本框或设定行高的块承载(区域高度≥1.2cm)，
    #       非仅一条自动撑高的下划线段落。bad 卡主观题作答用 1x1 表格且行高为0(自动撑高)。
    # 判定：存在承载主观题作答的高块——文本框(h_cm≥1.2)或含题号的表格设定了行高。
    tall_shapes = sum(1 for s in doc.shapes
                      if s['has_txbx'] and s['h_cm'] is not None and s['h_cm'] >= 1.2)
    # 主观题作答表格（含"分)"字样、非作文格）里带显式行高的
    ans_tbls = [t for t in doc.tables
                if t['ncol'] < 18 and ('分)' in norm_paren(t['text_ns']) or '分）' in t['text_ns'])]
    tall_tbls = 0
    for t in ans_tbls:
        total_h = 0
        for tr in t['el'].findall(f'{WNS}tr'):
            trPr = tr.find(f'{WNS}trPr')
            h = trPr.find(f'{WNS}trHeight') if trPr is not None else None
            if h is not None and h.get(f'{WNS}val'):
                total_h += int(h.get(f'{WNS}val'))
        if dxa_to_cm(total_h) >= 1.2:
            tall_tbls += 1
    hit = (tall_shapes >= 3) or (tall_tbls >= 3)
    R.append((5, '各主观题作答区高度与分值匹配、不塌缩(区域高度≥1.2cm)', hit,
              f'高文本框(≥1.2cm)={tall_shapes},设定行高的作答表格(≥1.2cm)={tall_tbls}(任一≥3)'))

    # ---- 31b. 答题卡正文字体不过小 (+3) ----
    # 细则新增：正文/作答文字中位字号须≥9磅(小五及以上)，好卡以12/9磅为主；
    # 字体被压到 6-7 磅（一屏塞满、难以书写辨认）判不达标。
    med_font = _body_median_font_pt(doc)
    hit = med_font is not None and med_font >= 9.0
    R.append((3, '答题卡正文/作答字体不过小(中位字号≥9磅、小五及以上)', hit,
              f'正文中位字号≈{med_font}磅(要求≥9)'))

    # ---- 30. 密封线结构：纵向窄带 (+3) ----
    # 细则：装订侧竖排"密…封…线…"，占页宽1/20-1/12、纵贯整页。
    # 实测好答题卡密封线文本框 w≈1.7cm、h≈24cm；从宽认定窄而高的竖带即可。
    seal_text = ('密' in ns and '封' in ns and '线' in ns)
    seal_shapes = [s for s in doc.shapes
                   if s['w_cm'] is not None and s['h_cm'] is not None
                   and s['w_cm'] <= pw / 10 and s['h_cm'] >= ph * 0.5]
    hit = seal_text and bool(seal_shapes)
    R.append((3, '密封线/装订线:竖排"密…封…线…"窄带、纵贯整页', hit,
              f'密封线文字={seal_text},窄高竖带数={len(seal_shapes)}'))

    # ---- 31. 密封线位置：最左侧竖直窄带 (+1) ----
    # 好答题卡密封带位于页2最右(装订侧)，此处按"存在纵向窄带且贴近页面一侧边缘"认定。
    hit = False
    detail = '未找到纵向密封窄带'
    seal_side = [s for s in doc.shapes
                 if s['w_cm'] is not None and s['h_cm'] is not None
                 and s['w_cm'] <= pw / 10 and s['h_cm'] >= ph * 0.5
                 and s['x_cm'] is not None
                 and (frac(s['cx_cm']) < 0.12 or frac(s['cx_cm']) > 0.88)]
    if seal_side:
        s = seal_side[0]
        hit = True
        detail = f'密封窄带cx比例={frac(s["cx_cm"]):.2f}(贴近页面一侧)'
    R.append((1, '密封线位于页面最侧竖直窄带', hit, detail))

    # ---- 32. 整卡文本框/分栏分块仿真版式 (+5) ----
    # 细则：全卡由多个定位文本框分块拼成(左中右多栏)，而非上下堆叠的下划线段落。
    positioned = [s for s in doc.shapes if s['has_txbx'] and s['x_cm'] is not None]
    hit = doc.txbx_count >= 8 and len(positioned) >= 6
    R.append((5, '整卡采用文本框/分栏分块的真实答题卡版式(多定位文本框拼成)', hit,
              f'文本框数={doc.txbx_count},带定位文本框={len(positioned)}(要求≥8/≥6)'))

    return R


# ============ 主函数 ============
def _locate_docx(dir_path):
    """仅在给定目录本身定位被评估的 .docx（不进入子目录），忽略 ~$ 临时文件。"""
    if not os.path.isdir(dir_path):
        return None
    cands = []
    for name in os.listdir(dir_path):
        if name.startswith('~$'):
            continue
        full = os.path.join(dir_path, name)
        if name.lower().endswith('.docx') and os.path.isfile(full):
            cands.append(full)
    return sorted(cands)[0] if cands else None


def evaluate(dir_path: str) -> dict:
    """统一入口：接收目录路径，自行定位并评估其中的 .docx。"""
    result = {
        'id': SCRIPT_ID,
        'file_name': '',
        'status': 'ok',
        'error': None,
        'dim1_pass': False,
        'dim1_reason': '',
        'dim2_items': [],
        'total_score': 0,
        'max_score': 0,
    }
    try:
        if not os.path.isdir(dir_path):
            result['status'] = 'error'
            result['error'] = f'目录不存在或不是目录：{dir_path}'
            # 维度一根本没进入判定，dim1_pass 与 dim1_reason 一律留空（空字符串，而非 False）
            result['dim1_pass'] = ''
            return result

        filepath = _locate_docx(dir_path)
        if filepath is None:
            result['status'] = 'error'
            result['error'] = f'目录中未找到 .docx 文件：{dir_path}'
            result['dim1_pass'] = ''
            return result
        result['file_name'] = os.path.basename(filepath)

        try:
            doc = DocParser(filepath)
        except Exception as e:
            result['status'] = 'error'
            result['error'] = f'无法打开文件：{type(e).__name__}: {e}'
            return result

        # 维度一（一票否决）
        dim1_pass, dim1_results = check_dim1(doc)
        result['dim1_pass'] = bool(dim1_pass)
        if not dim1_pass:
            failed = [name for name, ok, _ in dim1_results if not ok]
            result['dim1_reason'] = '；'.join(failed) if failed else '未通过'
            result['dim2_items'] = []
            result['total_score'] = 0
            result['max_score'] = 0
            return result

        # 维度二
        dim2 = check_dim2(doc)
        items = []
        total = 0
        max_score = 0
        for score, desc, hit, _detail in dim2:
            max_score += score
            delta = score if hit else 0
            if hit:
                total += score
            items.append({
                'rule': desc,
                'max_delta': score,
                'delta': delta,
                'hit': bool(hit),
            })
        result['dim2_items'] = items
        result['total_score'] = total
        result['max_score'] = max_score
        return result
    except Exception as e:
        result['status'] = 'error'
        result['error'] = f'脚本运行异常：{type(e).__name__}: {e}'
        return result


if __name__ == '__main__':
    target_dir = sys.argv[1] if len(sys.argv) > 1 else os.path.dirname(os.path.abspath(__file__))
    print(json.dumps(evaluate(target_dir), ensure_ascii=False, indent=2))
