#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# pyright: reportUnknownMemberType=false, reportUnknownVariableType=false, reportUnknownArgumentType=false, reportMissingImports=false, reportOptionalMemberAccess=false, reportGeneralTypeIssues=false
"""
officeval_009_verifier
======================

评测任务：对 Word 文档「第三章研究结果.docx」内 8 个统计表（表 1～表 8）做长表格跨页处理——
          续页顶端重复表头，并在断页前一页的表体底部右下角加“续下页”三字。

维度1（可用与可修改性）：交付物为 1 个可正常打开的 .docx，表格与文字均为可单独点选、编辑的
          真实对象（非整表截图 / 整页位图）。不满足则完成度总分直接记 0。

维度2（完成度）：加分项 18 条（+46），减分项 24 条（-30）。

          【加减分互不重复的划界原则】同一个判定只出现一次：
          · 加分项只考核“模型必须动手新建的东西”——续页片段（第 2 段及以后）的表头行，
            以及断口处新增的“续下页”标注段落、断口前的封口线；标注“做了但没做对”
            （定位方式、是否占用单元格、锚点是否与断页位置对应）同归加分项。
          · 减分项只考核“交付前已经正确、不改即对的东西”——首个表体片段的表头与外观基线、
            全部表体数据行、正文/表题/表注、表 3 表 5 未被误动、页面设置。
          · 因此表头类判定按片段划界：首片段归减分项，续页片段归加分项；
            表体末行封口按位置划界：断口前片段末行归加分项（A15），末片段末行归减分项（D10）。

返回结构与 officeval_008_verifier 保持一致：
    {"id", "file_name", "status", "error", "dim1_pass", "dim1_reason",
     "dim2_items":[{"rule","max_delta","delta","hit"}], "total_score", "max_score"}
"""

import os
import re
import sys
import json
import hashlib
import traceback

from docx import Document
from docx.oxml.ns import qn


SCRIPT_ID = "009"

# ---------- 分值序列（顺序与细则完全一致） ----------
# A01..A18：加分项（含“续下页”的定位方式、是否占用单元格、锚点是否对应——属于“做了但没做对”）
_ADD_DELTAS_ORDER = [3, 5, 3, 3, 3, 3, 1, 5, 3, 3, 3, 1, 1, 1, 3, 3, 1, 1]
# D01..D24：减分项（只保留“指令未涉及、交付前已正确却被改坏”的部分）
_DED_DELTAS_ORDER = [-1] * 8 + [-1] * 7 + [-3, -3, -3, -1, -1, -1] + [-1, -1, -1]

MAX_SCORE_TOTAL = sum(_ADD_DELTAS_ORDER)               # +46：完成度上限
_MIN_TOTAL = MAX_SCORE_TOTAL + sum(_DED_DELTAS_ORDER)  # +16：全部减分项命中时的理论下限

assert MAX_SCORE_TOTAL == 46, MAX_SCORE_TOTAL
assert sum(_DED_DELTAS_ORDER) == -30, sum(_DED_DELTAS_ORDER)
assert len(_ADD_DELTAS_ORDER) == 18 and len(_DED_DELTAS_ORDER) == 24
assert _MIN_TOTAL == 16, _MIN_TOTAL

_MAX_DELTAS_ORDER = _ADD_DELTAS_ORDER + _DED_DELTAS_ORDER


# ---------- 版面基线（描述“不需要修改的项”，用于减分项判定） ----------
STRICT_SONGTI = ('宋体', 'SimSun', 'NSimSun')
STRICT_HEITI = ('黑体', 'SimHei')
STRICT_TNR = ('Times New Roman', 'TimesNewRoman')

# 8 个统计表的目标形态：列数 / 每列栅格宽(twips) / 表体片段数 / 断页处数 / 数据行数 / 字号(半磅)
_T_COLS = [6, 6, 5, 5, 5, 6, 6, 7]
_T_GRID = [1384, 1384, 1661, 1661, 1661, 1384, 1384, 1187]
_T_FRAGS = [2, 3, 1, 2, 1, 2, 4, 3]
_T_BREAKS = [1, 2, 0, 1, 0, 1, 3, 2]
_T_ROWS = [42, 39, 8, 11, 8, 42, 80, 36]
_T_SZ = [16, 16, 17, 17, 17, 16, 15, 15]

assert sum(_T_FRAGS) == 18 and sum(_T_BREAKS) == 10 and sum(_T_ROWS) == 266

# 各表表头逐列文字（去空白后比对）
_T_HEADERS = [
    ('变量', '总计(n=268)', '接受数字随访(n=164)', '暂不接受(n=104)', '统计量', 'P值'),
    ('变量', '总计', '接受数字随访', '暂不接受', '统计量', 'P值'),
    ('变量', '缺失数（%）', '变量类型', '编码方式', '处理方式'),
    ('变量', 'β', 'S.E', 'OR（95%CI）', 'P值'),
    ('变量', 'β', 'S.E', 'OR（95%CI）', 'P值'),
    ('变量', '18～34岁(n=31)', '35～44岁(n=58)', '45～64岁(n=179)', '统计量', 'P值'),
    ('层级', '变量', '模型', 'OR（95%CI）', 'P值', '校正内容'),
    ('序号', '分析方案', '核心变量', 'OR（95%CI）', 'P值', '方向', '结论'),
]

# 表体数据行指纹：首列文字序列 / 全部单元格文字（均去空白后按行拼接）
_ROWSEQ_MD5 = '8f11591d9be0a1aeeaf92b9f33f4be10'
_ROWALL_MD5 = '79f6904996e4b943e57d0f9a98cc01a7'

# 8 条表题
_CAPTIONS = [
    '表1研究对象基础构成及数字随访接受度分布',
    '表2服务接触与近期管理体验的组间比较',
    '表3数据完整性与变量编码检查',
    '表4数字随访接受度回归分析',
    '表5年龄层级与服务便利性感知交互作用分析',
    '表6不同年龄层级对象的综合特征比较',
    '表7分层模型的逐步校准结果',
    '表8稳健性检验与替代定义分析',
]

# 8 条表注开头
_NOTES = [
    '注：分类变量以n（%）表示',
    '注：本表中服务接触类变量已按公开材料需要重新命名和重排',
    '注：缺失值比例均低于2%',
    '注：因变量为是否接受数字随访',
    '注：模型同时校正性别、经济压力',
    '注：年龄层级结果用于解释交互效应',
    '注：Model1纳入核心解释变量',
    '注：所有稳健性检验均基于虚构汇总数据',
]

# 章 / 节 / 小节标题
_H1 = '第三章研究结果'
_H2 = [
    '3.1样本整理与变量概况',
    '3.2数字随访接受度的影响因素分析',
    '3.3分层比较与模型校准',
    '3.4稳健性检验与补充分析',
]
_H3 = [
    '3.1.1纳入对象构成与数字随访接受度',
    '3.1.2数据完整性与变量编码检查',
    '3.2.1主效应模型结果',
    '3.2.2年龄层级与便利性感知的交互作用',
    '3.3.1不同年龄层级对象的综合特征比较',
    '3.3.2分层模型的逐步校准结果',
]

# 9 段正文（含结尾总结段）开头
_BODY_HEADS = [
    '本章基于一项虚构的社区慢病康复管理资料',
    '最终纳入268名完成基线评估的服务对象',
    '在建模前，对主要变量的缺失比例',
    '在单因素比较的基础上',
    '为判断不同年龄层级中服务便利性感知',
    '考虑到年龄可能影响数字工具的使用能力',
    '进一步分别在18～44岁与45～64岁对象中',
    '为检验主要结论的稳定性',
    '综上，改写后的结果显示',
]
_SUMMARY_HEAD = '综上，改写后的结果显示'

# 段落样式名（不得被替换）
_STYLE_NAMES = {
    'body': 'Body Text',
    'caption': 'Table Caption',
    'note': 'Table Note',
    'h1': 'heading 1',
    'h2': 'heading 2',
    'h3': 'heading 3',
}

# 表格内不得被改写的数值/文字样本
_CELL_SENTINELS = ('151(56.34)', '−3.216', '<0.001', '1.36(1.17,1.58)', '加入服务接触与随访记录')

# 页面设置：A4 210mm×297mm、上下页边距 25.4mm、左右页边距 31.75mm、单栏
_PG_W, _PG_H = 11906, 16838
_PG_MAR = {'top': 1440, 'bottom': 1440, 'left': 1800, 'right': 1800}

_MARK = '续下页'
_MARK_BAD_WORDS = ('续下表', '接下页', '下页继续', '转下页', '续表', '(续)', '（续）')

# 浮动对象容器：其中的文字不计入“可直接选中编辑的正文/单元格文字”
_FLOAT_TAGS = (
    qn('w:pict'),
    qn('w:drawing'),
    '{urn:schemas-microsoft-com:vml}textbox',
    qn('w:txbxContent'),
)


# =====================================================================
# 基础工具
# =====================================================================

def _norm(text):
    """去掉所有空白字符，便于文字比对。"""
    return re.sub(r'\s+', '', text or '')


def _in_float(el):
    """判断节点是否位于浮动对象（文本框/图片）内部。"""
    cur = el
    while cur is not None:
        if cur.tag in _FLOAT_TAGS:
            return True
        cur = cur.getparent()
    return False


def _plain_text(el):
    """提取节点内“可直接选中编辑”的文字，跳过浮动对象内部的文字。"""
    out = []
    for t in el.iter(qn('w:t')):
        if not _in_float(t):
            out.append(t.text or '')
    return ''.join(out)


def _float_text(el):
    """提取节点内浮动对象所承载的文字。"""
    out = []
    for t in el.iter(qn('w:t')):
        if _in_float(t):
            out.append(t.text or '')
    return ''.join(out)


def _all_text(el):
    return ''.join(t.text or '' for t in el.iter(qn('w:t')))


def _style_id(p):
    node = p.find(qn('w:pPr') + '/' + qn('w:pStyle'))
    return node.get(qn('w:val')) if node is not None else None


def _iter_runs(p):
    """段落直属 run（不含浮动对象内部的 run）。"""
    return [r for r in p.findall(qn('w:r')) if r.find(qn('w:t')) is not None]


class _StyleBook(object):
    """样式表索引：按 styleId 提供 name / basedOn / pPr / rPr，支持沿 basedOn 链回溯。"""

    def __init__(self, doc):
        self._el = {}
        self._name = {}
        self._based = {}
        try:
            root = doc.styles.element
        except Exception:
            root = None
        if root is None:
            return
        for st in root.findall(qn('w:style')):
            sid = st.get(qn('w:styleId'))
            if not sid:
                continue
            self._el[sid] = st
            nm = st.find(qn('w:name'))
            self._name[sid] = nm.get(qn('w:val')) if nm is not None else ''
            bo = st.find(qn('w:basedOn'))
            self._based[sid] = bo.get(qn('w:val')) if bo is not None else None
        self._default_p = None
        for sid, st in self._el.items():
            if st.get(qn('w:type')) == 'paragraph' and st.get(qn('w:default')) in ('1', 'true'):
                self._default_p = sid
                break

    def name_of(self, sid):
        return self._name.get(sid, '')

    def id_by_name(self, name):
        for sid, nm in self._name.items():
            if nm == name:
                return sid
        return None

    def chain(self, sid):
        """styleId → [自身, 父, 祖父 ...]，最后追加默认段落样式。"""
        out = []
        cur = sid
        guard = 0
        while cur and cur in self._el and guard < 16:
            out.append(cur)
            cur = self._based.get(cur)
            guard += 1
        if self._default_p and self._default_p not in out:
            out.append(self._default_p)
        return out

    def find(self, sid, path):
        """沿 basedOn 链查找第一个存在的节点，path 形如 'w:pPr/w:jc'。"""
        xp = '/'.join(qn(seg) for seg in path.split('/'))
        for s in self.chain(sid):
            node = self._el[s].find(xp)
            if node is not None:
                return node
        return None


def _eff_ppr(book, p, path):
    """段落属性的有效值节点：段落直接属性优先，其次沿样式 basedOn 链。"""
    xp = '/'.join(qn(seg) for seg in path.split('/'))
    node = p.find(xp)
    if node is not None:
        return node
    return book.find(_style_id(p), path)


def _eff_rpr(book, p, run, tag):
    """run 属性的有效值节点：run rPr → 段落 pPr/rPr → 样式链 rPr。"""
    if run is not None:
        rPr = run.find(qn('w:rPr'))
        if rPr is not None:
            node = rPr.find(qn(tag))
            if node is not None:
                return node
    node = p.find(qn('w:pPr') + '/' + qn('w:rPr') + '/' + qn(tag))
    if node is not None:
        return node
    return book.find(_style_id(p), 'w:rPr/' + tag)


def _eastasia_of(book, p, run):
    node = _eff_rpr(book, p, run, 'w:rFonts')
    if node is None:
        return None
    return node.get(qn('w:eastAsia')) or node.get(qn('w:eastAsiaTheme'))


def _ascii_of(book, p, run):
    node = _eff_rpr(book, p, run, 'w:rFonts')
    if node is None:
        return None
    return (node.get(qn('w:ascii')) or node.get(qn('w:hAnsi'))
            or node.get(qn('w:asciiTheme')) or node.get(qn('w:hAnsiTheme')))


def _sz_of(book, p, run):
    node = _eff_rpr(book, p, run, 'w:sz')
    if node is None:
        return None
    try:
        return int(node.get(qn('w:val')))
    except (TypeError, ValueError):
        return None


def _bold_of(book, p, run):
    node = _eff_rpr(book, p, run, 'w:b')
    if node is None:
        return False
    return node.get(qn('w:val')) not in ('0', 'false', 'off')


def _jc_of(book, p):
    node = _eff_ppr(book, p, 'w:pPr/w:jc')
    return node.get(qn('w:val')) if node is not None else None


def _int_attr(node, tag, default=None):
    if node is None:
        return default
    val = node.get(qn(tag))
    if val is None:
        return default
    try:
        return int(val)
    except ValueError:
        return default


# =====================================================================
# 表格结构工具
# =====================================================================

def _rows(tbl):
    return tbl.findall(qn('w:tr'))


def _cells(tr):
    return tr.findall(qn('w:tc'))


def _is_header_row(tr):
    return tr.find(qn('w:trPr') + '/' + qn('w:tblHeader')) is not None


def _has_cant_split(tr):
    return tr.find(qn('w:trPr') + '/' + qn('w:cantSplit')) is not None


def _grid(tbl):
    return [_int_attr(g, 'w:w') for g in tbl.findall(qn('w:tblGrid') + '/' + qn('w:gridCol'))]


def _tblpr(tbl, tag):
    tblPr = tbl.find(qn('w:tblPr'))
    return tblPr.find(qn(tag)) if tblPr is not None else None


def _border(tc, edge):
    """返回 (val, sz)；无设置返回 (None, None)。"""
    tcPr = tc.find(qn('w:tcPr'))
    if tcPr is None:
        return (None, None)
    bs = tcPr.find(qn('w:tcBorders'))
    if bs is None:
        return (None, None)
    node = bs.find(qn('w:' + edge))
    if node is None:
        return (None, None)
    return (node.get(qn('w:val')), _int_attr(node, 'w:sz'))


def _is_line(edge, min_sz=None, max_sz=None):
    val, sz = edge
    if val in (None, 'nil', 'none'):
        return False
    if min_sz is not None and (sz is None or sz < min_sz):
        return False
    if max_sz is not None and (sz is None or sz > max_sz):
        return False
    return True


def _is_blank(edge):
    val, _ = edge
    return val in (None, 'nil', 'none')


def _shd_fill(tc):
    tcPr = tc.find(qn('w:tcPr'))
    node = tcPr.find(qn('w:shd')) if tcPr is not None else None
    return node.get(qn('w:fill')) if node is not None else None


def _valign(tc):
    tcPr = tc.find(qn('w:tcPr'))
    node = tcPr.find(qn('w:vAlign')) if tcPr is not None else None
    return node.get(qn('w:val')) if node is not None else None


# =====================================================================
# 文档解析：正文序列 → 表题分组 → 表体片段 / 片段间隔
# =====================================================================

_CAP_RE = re.compile(r'^表\s*(\d+)')


def _scan_body(body):
    """把正文顶层子节点整理成 items 序列，并按表题把表体片段分组。"""
    items = []
    for ch in body.iterchildren():
        if ch.tag == qn('w:p'):
            items.append({'kind': 'p', 'el': ch, 'text': _plain_text(ch),
                          'float': _float_text(ch)})
        elif ch.tag == qn('w:tbl'):
            items.append({'kind': 'tbl', 'el': ch, 'text': '', 'float': ''})

    groups = []
    for idx, it in enumerate(items):
        if it['kind'] != 'p':
            continue
        m = _CAP_RE.match(_norm(it['text']))
        if m:
            groups.append({'no': int(m.group(1)), 'cap_idx': idx, 'cap_el': it['el'],
                           'cap_text': _norm(it['text']), 'frag_idx': []})
    for idx, it in enumerate(items):
        if it['kind'] != 'tbl':
            continue
        owner = None
        for g in groups:
            if g['cap_idx'] < idx:
                owner = g
            else:
                break
        if owner is not None:
            owner['frag_idx'].append(idx)

    for g in groups:
        g['frags'] = [items[i]['el'] for i in g['frag_idx']]
        g['gaps'] = []
        for a, b in zip(g['frag_idx'], g['frag_idx'][1:]):
            g['gaps'].append(items[a + 1:b])
    return items, groups


def _row_texts(tr):
    return tuple(_norm(_plain_text(tc)) for tc in _cells(tr))


def _looks_like_header(tr, expected):
    if expected is None:
        return False
    return _row_texts(tr) == tuple(expected)


def _data_rows(g, expected=None):
    """片段内的表体数据行：跳过标题行（含未勾选重复但文字与表头一致的首行）。"""
    out = []
    for tbl in g['frags']:
        rows = _rows(tbl)
        for ri, tr in enumerate(rows):
            if _is_header_row(tr):
                continue
            if ri == 0 and _looks_like_header(tr, expected):
                continue
            out.append((tbl, tr))
    return out


def _header_rows(g, expected=None):
    """每个片段的表头行（若该片段首行是表头则取之，否则为 None）。"""
    out = []
    for tbl in g['frags']:
        rows = _rows(tbl)
        first = rows[0] if rows else None
        if first is not None and (_is_header_row(first) or _looks_like_header(first, expected)):
            out.append((tbl, first))
        else:
            out.append((tbl, None))
    return out


# =====================================================================
# 入口
# =====================================================================

def _open_word_document(file_path):
    return Document(file_path)


def evaluate(dir_path):
    """在 dir_path 中定位并评估 .docx，返回规范化结果 dict。

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
            raise FileNotFoundError(f"目录不存在或不是目录：{dir_path}")
        cands = sorted(
            f for f in os.listdir(dir_path)
            if f.lower().endswith(('.docx', '.doc')) and not f.startswith('~$')
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
        result["dim1_reason"] = ""
        result["dim2_items"] = []
        result["total_score"] = 0
        result["max_score"] = 0
        return result
    except Exception as exc:
        result["status"] = "error"
        result["error"] = f"脚本运行异常：{type(exc).__name__}: {exc}"
        result["dim1_pass"] = False
        result["dim1_reason"] = f"文件无法解析：{exc}"
        result["dim2_items"] = []
        result["total_score"] = 0
        return result


def _evaluate_impl(file_path, result):
    doc = _open_word_document(file_path)
    body = doc.element.body
    book = _StyleBook(doc)
    items, groups = _scan_body(body)

    # ---------------- 维度1：可用与可修改性 ----------------
    para_chars = sum(len(_norm(it['text'])) for it in items if it['kind'] == 'p')
    tables = [it['el'] for it in items if it['kind'] == 'tbl']
    cell_chars = 0
    for tbl in tables:
        for tr in _rows(tbl):
            for tc in _cells(tr):
                cell_chars += len(_norm(_plain_text(tc)))
    n_images = len(body.findall('.//' + qn('w:drawing'))) + len(body.findall('.//' + qn('w:pict')))

    if not tables or cell_chars < 2000 or para_chars < 800:
        result["dim1_pass"] = False
        result["dim1_reason"] = (
            "文档缺少可编辑的表格或正文文字（表格 %d 个、单元格文字 %d 字、正文 %d 字，"
            "疑似被转为图片或版面严重错乱）" % (len(tables), cell_chars, para_chars)
        )
        result["dim2_items"] = []
        result["total_score"] = 0
        result["max_score"] = MAX_SCORE_TOTAL
        return result

    result["dim1_pass"] = True
    result["dim1_reason"] = (
        "文件可正常打开，%d 个表格与 %d 段文字均为可单独点选编辑的对象（内嵌图片 %d 个），版面未见严重错乱"
        % (len(tables), sum(1 for it in items if it['kind'] == 'p' and _norm(it['text'])), n_images)
    )

    # ---------------- 结构预处理 ----------------
    # 按表题编号定位表 1～表 8；缺失的表以空组占位，避免索引错位
    by_no = {}
    for g in groups:
        by_no.setdefault(g['no'], g)
    tbl_groups = [by_no.get(n + 1) for n in range(8)]

    frag_counts = []
    for gi, g in enumerate(tbl_groups):
        frag_counts.append(len(g['frags']) if g else 0)

    # 每个表的表头行 / 数据行
    hdr_rows, data_rows = [], []
    for gi, g in enumerate(tbl_groups):
        exp = _T_HEADERS[gi]
        hdr_rows.append(_header_rows(g, exp) if g else [])
        data_rows.append(_data_rows(g, exp) if g else [])

    # 加减分划界用的片段切分：首片段（交付前已有，归减分项）／续页片段（模型新建，归加分项）
    long_ids = (0, 1, 3, 5, 6, 7)                 # 会跨页的 6 个表
    first_hdr = [hdr_rows[gi][0] if hdr_rows[gi] else None for gi in range(8)]
    cont_hdr = [hdr_rows[gi][1:] for gi in range(8)]
    n_cont_frags = sum(len(cont_hdr[gi]) for gi in long_ids)

    # 片段中部混入的多余表头行：表头“数量与位置是否正确”由 A02 独占判定
    # （表头行只应出现在每个片段的首行；出现在片段中部即为多余的重复表头）
    stray_hdr = [0] * 8
    for gi, g in enumerate(tbl_groups):
        if not g:
            continue
        exp = tuple(_T_HEADERS[gi])
        for tbl in g['frags']:
            for ri, tr in enumerate(_rows(tbl)):
                if ri == 0:
                    continue
                if _is_header_row(tr) or _row_texts(tr) == exp:
                    stray_hdr[gi] += 1

    # ---------------- “续下页”标注清点 ----------------
    marks = []          # 全篇所有“续下页”载体
    for idx, it in enumerate(items):
        if it['kind'] == 'p':
            if _MARK in _norm(it['text']):
                marks.append({'carrier': 'para', 'idx': idx, 'el': it['el'], 'host': it['el'],
                              'text': _norm(it['text'])})
            if _MARK in _norm(it['float']):
                marks.append({'carrier': 'float', 'idx': idx, 'el': it['el'], 'host': it['el'],
                              'text': _norm(it['float'])})
        else:
            for tc in it['el'].iter(qn('w:tc')):
                if _MARK in _norm(_plain_text(tc)):
                    marks.append({'carrier': 'cell', 'idx': idx, 'el': tc, 'host': it['el'],
                                  'text': _norm(_plain_text(tc))})
                if _MARK in _norm(_float_text(tc)):
                    marks.append({'carrier': 'float', 'idx': idx, 'el': tc, 'host': it['el'],
                                  'text': _norm(_float_text(tc))})

    n_float_shapes = 0
    n_abs_page = 0
    for pict in body.iter(qn('w:pict')):
        if _MARK not in _norm(_all_text(pict)):
            continue
        n_float_shapes += 1
        style = ''
        for sh in pict.iter():
            val = sh.get('style')
            if val:
                style += val + ';'
        low = style.lower()
        if 'position:absolute' in low and 'relative:page' in low:
            n_abs_page += 1
    for drw in body.iter(qn('w:drawing')):
        if _MARK in _norm(_all_text(drw)):
            n_float_shapes += 1
            if drw.find('.//' + qn('wp:anchor')) is not None:
                n_abs_page += 1

    # 浮动对象锚点分析：锚在表题段落 / 表头单元格 / 数据单元格 / 普通段落
    float_anchors = []
    for shp in list(body.iter(qn('w:pict'))) + list(body.iter(qn('w:drawing'))):
        if _MARK not in _norm(_all_text(shp)):
            continue
        tc = host_p = None
        cur = shp.getparent()
        while cur is not None:
            if host_p is None and cur.tag == qn('w:p'):
                host_p = cur
            if tc is None and cur.tag == qn('w:tc'):
                tc = cur
            cur = cur.getparent()
        kind = 'para'
        if tc is not None:
            kind = 'header' if _is_header_row(tc.getparent()) else 'data'
        elif host_p is not None and _CAP_RE.match(_norm(_plain_text(host_p))):
            kind = 'caption'
        float_anchors.append({'in_cell': tc is not None, 'kind': kind})

    # 断页处（相邻片段之间）合规的“续下页”段落
    gap_marks = []       # 每个表：[该间隔内恰有一处标注的间隔数]
    gap_blank = []       # 每个表：间隔内残留的空段落数
    gap_extra = []       # 每个表：间隔内的其它多余内容数
    for gi, g in enumerate(tbl_groups):
        good_n, blank_n, extra_n = 0, 0, 0
        for gap in (g['gaps'] if g else []):
            hit = 0
            for sub in gap:
                if sub['kind'] != 'p':
                    extra_n += 1
                    continue
                txt = _norm(sub['text'])
                if _MARK in txt:
                    hit += 1
                elif txt == '':
                    blank_n += 1
                else:
                    extra_n += 1
            if hit == 1:
                good_n += 1
        gap_marks.append(good_n)
        gap_blank.append(blank_n)
        gap_extra.append(extra_n)

    # 长表实际存在的断口总数：A08／A16 只按实际断口判定，段数是否正确由 A01 独占
    total_gaps = sum(len(tbl_groups[gi]['gaps']) for gi in long_ids if tbl_groups[gi])

    # 每个表的正文区间 [表题, 下一条表题)，用于把“续下页”归属到具体表格
    spans = []
    cap_idx_sorted = sorted(g['cap_idx'] for g in groups) if groups else []
    for g in tbl_groups:
        if not g:
            spans.append(None)
            continue
        nxt = len(items)
        for ci in cap_idx_sorted:
            if ci > g['cap_idx']:
                nxt = ci
                break
        spans.append((g['cap_idx'], nxt))

    def _marks_in(gi):
        span = spans[gi]
        if span is None:
            return []
        return [m for m in marks if span[0] <= m['idx'] < span[1]]

    # 断开前的“最后一段表体”起点：位于其后的标注属于“表格结尾多余标注”
    def _marks_after_last_frag(gi):
        g = tbl_groups[gi]
        if not g or not g['frag_idx']:
            return []
        last = g['frag_idx'][-1]
        return [m for m in _marks_in(gi) if m['idx'] >= last]

    results = []

    # =================================================================
    # 一、加分项
    # =================================================================

    # ---- A01 +3：6 个长表被切分成正确的段数（段数判定由本项独占） ----
    a01 = all(frag_counts[gi] == _T_FRAGS[gi] for gi in long_ids)
    results.append(("+3：" if a01 else "+0：") + "正确识别并处理了会跨页的 6 个表格——表 1、表 2、表 4、表 6、表 7、表 8 的表体依次切分为 2、3、2、2、4、3 段（长表共 16 段）")

    # ---- A02 +5：续页表头齐全（全篇应有 10 个）且只出现在片段顶端（表头数量与位置由本项独占） ----
    # 少一个断口没补表头、或某个续页片段首行不是表头，本项均不得分
    a02 = all(sum(1 for _, tr in cont_hdr[gi] if tr is not None) == _T_FRAGS[gi] - 1
              for gi in long_ids) and all(stray_hdr[gi] == 0 for gi in long_ids)
    results.append(("+5：" if a02 else "+0：") + "6 个长表的每一个续页片段顶端均出现完整表头行（全篇应有 10 个续页表头：表 1 一个、表 2 两个、表 4 一个、表 6 一个、表 7 三个、表 8 两个），且表头行只出现在片段顶端、片段中部未混入多余的重复表头行（本条独占“表头数量与位置”的判定；首片段表头为交付前既有，不计入本项）")

    # ---- A03 +3：续页表头文字与该表首页表头逐列一致 ----
    a03 = n_cont_frags > 0
    for gi in long_ids:
        fh = first_hdr[gi]
        if fh is None or fh[1] is None:
            a03 = False
            break
        base_txt = _row_texts(fh[1])
        for _, tr in cont_hdr[gi]:
            if tr is None or _row_texts(tr) != base_txt:
                a03 = False
        if not a03:
            break
    results.append(("+3：" if a03 else "+0：") + "续页表头行的文字与该表首页表头逐列一致（首页表头文字本身是否被改写由减分项 D16 判定）")

    # ---- A04 +3：续页表头行已勾选“在各页顶端以标题行形式重复出现”（tblHeader） ----
    a04 = n_cont_frags > 0
    for gi in long_ids:
        for _, tr in cont_hdr[gi]:
            if tr is None or not _is_header_row(tr):
                a04 = False
    results.append(("+3：" if a04 else "+0：") + "续页表头行已勾选“在各页顶端以标题行形式重复出现”（行属性含 tblHeader），后续再分页时可自动重复")

    # ---- A05 +3：续页表头沿用表头格式（TNR/宋体、8 磅、加粗、水平+垂直居中、白色底纹） ----
    a05 = n_cont_frags > 0
    for gi in long_ids:
        for tbl, tr in cont_hdr[gi]:
            if tr is None:
                a05 = False
                break
            for tc in _cells(tr):
                if (_shd_fill(tc) or '').upper() != 'FFFFFF' or _valign(tc) != 'center':
                    a05 = False
                for p in tc.findall(qn('w:p')):
                    if _jc_of(book, p) != 'center':
                        a05 = False
                    for r in _iter_runs(p):
                        ea = _eastasia_of(book, p, r) or ''
                        asc = _ascii_of(book, p, r) or ''
                        if ea not in STRICT_SONGTI or asc not in STRICT_TNR:
                            a05 = False
                        if _sz_of(book, p, r) != _T_SZ[gi]:
                            a05 = False
                        if not _bold_of(book, p, r):
                            a05 = False
            if not a05:
                break
        if not a05:
            break
    results.append(("+3：" if a05 else "+0：") + "续页表头行沿用表头格式——西文 Times New Roman／中文宋体、8 磅、加粗、水平居中且垂直居中、白色底纹")

    # ---- A06 +3：续页表头上框线 1.5 磅、下框线细线，列宽与首页表头逐列相同 ----
    a06 = n_cont_frags > 0
    for gi in long_ids:
        fh = first_hdr[gi]
        if fh is None:
            a06 = False
            break
        base_grid = _grid(fh[0])
        for tbl, tr in cont_hdr[gi]:
            if tr is None:
                a06 = False
                break
            if _grid(tbl) != base_grid:
                a06 = False
            for tc in _cells(tr):
                if not _is_line(_border(tc, 'top'), min_sz=12):
                    a06 = False
                if not _is_line(_border(tc, 'bottom'), max_sz=8):
                    a06 = False
            if not a06:
                break
        if not a06:
            break
    results.append(("+3：" if a06 else "+0：") + "续页表头行的上框线为 1.5 磅粗线、下框线为细线，列宽与该表首页表头逐列相同（首页列宽基线是否被改由减分项 D11／D13 判定）")

    # ---- A07 +1：续页表头单元格内的两行式写法保持原样 ----
    a07 = n_cont_frags > 0
    for gi in long_ids:
        for _, tr in cont_hdr[gi]:
            if tr is None:
                a07 = False
                break
            for tc in _cells(tr):
                txt = _norm(_plain_text(tc))
                if '(n=' in txt and tc.find('.//' + qn('w:br')) is None:
                    a07 = False
            if not a07:
                break
        if not a07:
            break
    results.append(("+1：" if a07 else "+0：") + "续页表头单元格内的两行式写法保持原样（如“总计”与“(n=268)”分两行显示）")

    # ---- A08 +5：每一处实际断口内恰有一处“续下页”（数量判定由本项独占） ----
    a08 = (total_gaps > 0 and sum(gap_marks[gi] for gi in long_ids) == total_gaps)
    results.append(("+5：" if a08 else "+0：") + "每一处表格跨页断开的位置都在断开前的表体下方加了“续下页”，每个断口恰有一处、不重不漏（段数正确时全篇共 10 处）")

    # ---- A09 +3：“续下页”以可选中编辑的文字承载（承载方式判定由本项独占） ----
    a09 = bool(marks) and all(m['carrier'] in ('para', 'cell') for m in marks) and n_float_shapes == 0
    results.append(("+3：" if a09 else "+0：") + "“续下页”为可直接选中、可编辑的文字，未使用浮动文本框、图片或艺术字承载")

    # ---- A10 +3：“续下页”段落右对齐（对齐判定由本项独占，位置数量见 A08） ----
    mark_ps = [m['el'] for m in marks if m['carrier'] == 'para']
    a10 = bool(mark_ps) and all(_jc_of(book, p) == 'right' for p in mark_ps)
    results.append(("+3：" if a10 else "+0：") + "“续下页”作为独立正文段落右对齐（贴近右侧页边），与示例图片中的位置一致")

    # ---- A11 +3：“续下页”随表格底部排版，未用相对于页面的绝对坐标定位 ----
    a11 = bool(marks) and n_abs_page == 0
    results.append(("+3：" if a11 else "+0：") + "“续下页”随表格底部排版，未采用相对于页面的绝对坐标定位（如固定在距页面顶端约 748 磅、距页面左边约 400 磅处），不会与表格底部脱钩而压住页脚、超出下页边距或与表体重叠")

    # ---- A12 +1：“续下页”位于表格之外，未占用数据单元格 ----
    a12 = bool(marks) and not (any(fa['in_cell'] for fa in float_anchors)
                               or any(m['carrier'] == 'cell' for m in marks))
    results.append(("+1：" if a12 else "+0：") + "“续下页”位于表格之外的正文层，未被置于表格单元格内部占用或覆盖数据单元格的显示位置")

    # ---- A13 +1：“续下页”的锚点与该表实际断页位置对应 ----
    a13 = bool(marks) and not any(fa['kind'] in ('caption', 'header') for fa in float_anchors)
    results.append(("+1：" if a13 else "+0：") + "“续下页”的出现位置与该表实际断页位置对应，锚点未挂在表题段落（如“表 7 分层模型的逐步校准结果”）或表头单元格上")

    # ---- A14 +1：“续下页”字体字号沿用正文（宋体 10.5 磅）且取消首行缩进 ----
    a14 = bool(marks)
    for t in body.iter(qn('w:t')):
        if _MARK not in _norm(t.text or ''):
            continue
        run = t.getparent()
        host = run.getparent() if run is not None else None
        while host is not None and host.tag != qn('w:p'):
            host = host.getparent()
        if host is None:
            a14 = False
            continue
        if _int_attr(_eff_ppr(book, host, 'w:pPr/w:ind'), 'w:firstLine', 0) != 0:
            a14 = False
        ea = _eastasia_of(book, host, run)
        if ea is not None and ea not in STRICT_SONGTI:
            a14 = False
        if _sz_of(book, host, run) != 21:
            a14 = False
    results.append(("+1：" if a14 else "+0：") + "“续下页”字体字号沿用正文（宋体 10.5 磅），且取消首行缩进")

    # ---- A15 +3：断开前的表体末行带 1.5 磅下框线封口 ----
    a15 = True
    for gi in range(8):
        g = tbl_groups[gi]
        frags = g['frags'] if g else []
        if len(frags) < 2:
            continue
        for tbl in frags[:-1]:
            rows = _rows(tbl)
            if not rows:
                a15 = False
                break
            for tc in _cells(rows[-1]):
                if not _is_line(_border(tc, 'bottom'), min_sz=12):
                    a15 = False
            if not a15:
                break
        if not a15:
            break
    results.append(("+3：" if a15 else "+0：") + "断开前的表体末行带 1.5 磅下框线，使上一页的表格底部先封口、再出现“续下页”")

    # ---- A16 +3：续页从下一页顶端开始，断口内无空段落／多余内容（空段落判定由本项独占） ----
    a16 = (total_gaps > 0 and sum(gap_blank) == 0 and sum(gap_extra) == 0)
    results.append(("+3：" if a16 else "+0：") + "表格的续页部分从下一页顶端开始，“续下页”与续页表头之间没有空段落、空行或半页空白")

    # ---- A17 +1：表格最后一段表体之后不出现“续下页” ----
    a17 = bool(marks) and not any(_marks_after_last_frag(gi) for gi in range(8))
    results.append(("+1：" if a17 else "+0：") + "表格最后一段表体之后不出现“续下页”（表 1、表 2、表 4、表 6、表 7、表 8 的结尾处以及表 3、表 5 均无）")

    # ---- A18 +1：标注文字为“续下页”三字且无多余标点 ----
    a18 = bool(marks)
    for m in marks:
        txt = _norm(m['text'])
        if txt != _MARK:
            a18 = False
    body_all = _norm(_all_text(body))
    for w in _MARK_BAD_WORDS:
        if _norm(w) in body_all:
            a18 = False
    results.append(("+1：" if a18 else "+0：") + "标注文字为“续下页”三字，未写成“续下表”“接下页”“下页继续”“(续)”等，且未附加多余标点")

    # ---- 结构测量：数据行完整性（供 D18 判定） ----
    row_counts = [len(data_rows[gi]) for gi in range(8)]
    seq_first, seq_all = [], []
    for gi in range(8):
        for _, tr in data_rows[gi]:
            cs = _cells(tr)
            seq_first.append(_norm(_plain_text(cs[0])) if cs else '')
            seq_all.append('|'.join(_norm(_plain_text(c)) for c in cs))
    md5_first = hashlib.md5('\n'.join(seq_first).encode('utf-8')).hexdigest()
    md5_all = hashlib.md5('\n'.join(seq_all).encode('utf-8')).hexdigest()
    rows_ok = (row_counts == _T_ROWS and md5_first == _ROWSEQ_MD5 and md5_all == _ROWALL_MD5)

    # ---- 结构测量：表题在表格上方、表注在表格下方（供 D22 判定） ----
    pos_ok = True
    for gi in range(8):
        g = tbl_groups[gi]
        if not g or not g['frag_idx'] or spans[gi] is None:
            pos_ok = False
            continue
        if _norm(g['cap_text']) != _CAPTIONS[gi] or g['cap_idx'] > g['frag_idx'][0]:
            pos_ok = False
        note_ok = False
        for i in range(g['frag_idx'][-1] + 1, spans[gi][1]):
            it = items[i]
            if it['kind'] == 'p' and _norm(it['text']).startswith(_NOTES[gi]):
                note_ok = True
        if not note_ok:
            pos_ok = False

    # =================================================================
    # 二、减分项
    # =================================================================

    paras = [it['el'] for it in items if it['kind'] == 'p']
    ptext = dict((id(p), _norm(_plain_text(p))) for p in paras)

    def _pick_eq(names):
        return [p for p in paras if ptext[id(p)] in names]

    def _pick_head(heads):
        return [p for p in paras if any(ptext[id(p)].startswith(h) for h in heads)]

    h1_ps = _pick_eq((_H1,))
    h2_ps = _pick_eq(tuple(_H2))
    h3_ps = _pick_eq(tuple(_H3))
    body_ps = _pick_head(tuple(_BODY_HEADS))
    sum_ps = _pick_head((_SUMMARY_HEAD,))
    cap_ps = [g['cap_el'] for g in groups]
    note_ps = [p for p in paras if ptext[id(p)].startswith('注：')]

    def _fmt_ok(ps, want_n, ea_set, asc_set, sz, bold,
                jc_set=None, first=None, line=None):
        """段落格式是否全部满足基线；数量不符或任一段不符即返回 False。"""
        if want_n is not None and len(ps) != want_n:
            return False
        if not ps:
            return False
        for p in ps:
            if jc_set is not None and _jc_of(book, p) not in jc_set:
                return False
            if first is not None:
                if _int_attr(_eff_ppr(book, p, 'w:pPr/w:ind'), 'w:firstLine', 0) != first:
                    return False
            if line is not None:
                if _int_attr(_eff_ppr(book, p, 'w:pPr/w:spacing'), 'w:line') != line:
                    return False
            runs = _iter_runs(p)
            if not runs:
                return False
            for r in runs:
                if ea_set and (_eastasia_of(book, p, r) not in ea_set):
                    return False
                if asc_set and (_ascii_of(book, p, r) not in asc_set):
                    return False
                if sz is not None and _sz_of(book, p, r) != sz:
                    return False
                if bold is not None and _bold_of(book, p, r) != bold:
                    return False
        return True

    # ---- D01 -1：章标题格式 ----
    d01 = not _fmt_ok(h1_ps, 1, STRICT_HEITI, STRICT_TNR, 32, True)
    results.append(("-1：" if d01 else "-0：") + "章标题“第三章 研究结果”不满足西文 Times New Roman／中文黑体、小三（16 磅）、加粗")

    # ---- D02 -1：节标题格式 ----
    d02 = not _fmt_ok(h2_ps, 4, STRICT_HEITI, STRICT_TNR, 28, True)
    results.append(("-1：" if d02 else "-0：") + "节标题“3.1 样本整理与变量概况”等 4 条不满足 Times New Roman／黑体、14 磅、加粗")

    # ---- D03 -1：小节标题格式 ----
    d03 = not _fmt_ok(h3_ps, 6, STRICT_HEITI, STRICT_TNR, 24, True)
    results.append(("-1：" if d03 else "-0：") + "小节标题“3.1.1 纳入对象构成与数字随访接受度”等 6 条不满足 Times New Roman／黑体、12 磅、加粗")

    # ---- D04 -1：正文段落格式 ----
    d04 = not _fmt_ok(body_ps, 9, STRICT_SONGTI, STRICT_TNR, 21, False,
                      first=420, line=324)
    results.append(("-1：" if d04 else "-0：") + "正文段落（共 9 段）不满足 Times New Roman／宋体、五号（10.5 磅）、首行缩进 2 字符、1.35 倍行距")

    # ---- D05 -1：表题格式 ----
    d05 = not _fmt_ok(cap_ps, 8, STRICT_SONGTI, STRICT_TNR, 21, True,
                      jc_set=('center',), first=0)
    results.append(("-1：" if d05 else "-0：") + "8 条表题不满足宋体、10.5 磅、加粗、居中对齐")

    # ---- D06 -1：表注格式 ----
    d06 = not _fmt_ok(note_ps, 8, STRICT_SONGTI, STRICT_TNR, 18, False,
                      jc_set=(None, 'left', 'start'), first=0)
    results.append(("-1：" if d06 else "-0：") + "8 条表注（“注：……”）不满足宋体、9 磅、左对齐、无首行缩进")

    # ---- D07 -1：结尾总结段格式 ----
    d07 = not _fmt_ok(sum_ps, 1, STRICT_SONGTI, STRICT_TNR, 21, False, first=420)
    results.append(("-1：" if d07 else "-0：") + "结尾总结段“综上，改写后的结果显示……”不满足 Times New Roman／宋体、10.5 磅、首行缩进 2 字符")

    # ---- D08 -1：段落样式名被替换 ----
    d08 = False
    for p, want in ([(p, _STYLE_NAMES['body']) for p in body_ps]
                    + [(p, _STYLE_NAMES['caption']) for p in cap_ps]
                    + [(p, _STYLE_NAMES['note']) for p in note_ps]
                    + [(p, _STYLE_NAMES['h1']) for p in h1_ps]
                    + [(p, _STYLE_NAMES['h2']) for p in h2_ps]
                    + [(p, _STYLE_NAMES['h3']) for p in h3_ps]):
        if book.name_of(_style_id(p)) != want:
            d08 = True
    results.append(("-1：" if d08 else "-0：") + "段落样式名被替换（正文为 Body Text、表题为 Table Caption、表注为 Table Note、三级标题为 heading 1／2／3）")

    # ---- D09 -1：表体数据文字格式 ----
    d09 = False
    for gi in range(8):
        for _, tr in data_rows[gi]:
            for ci, tc in enumerate(_cells(tr)):
                if _valign(tc) != 'center':
                    d09 = True
                for p in tc.findall(qn('w:p')):
                    if _jc_of(book, p) != ('left' if ci == 0 else 'center'):
                        d09 = True
                    sp = _eff_ppr(book, p, 'w:pPr/w:spacing')
                    if (_int_attr(sp, 'w:line') != 240 or _int_attr(sp, 'w:before', 0) != 0
                            or _int_attr(sp, 'w:after', 0) != 0):
                        d09 = True
                    for r in _iter_runs(p):
                        if _eastasia_of(book, p, r) not in STRICT_SONGTI:
                            d09 = True
                        if _ascii_of(book, p, r) not in STRICT_TNR:
                            d09 = True
                        if _sz_of(book, p, r) != _T_SZ[gi]:
                            d09 = True
    results.append(("-1：" if d09 else "-0：") + "表体数据文字不满足 Times New Roman／宋体、8 磅、水平居中、垂直居中、单倍行距且段前段后为 0")

    def _hdr_flags(gi, tbl):
        """片段内每一行是否为表头行（勾选 tblHeader，或首行文字与该表表头一致）。"""
        exp = _T_HEADERS[gi]
        rows = _rows(tbl)
        return [_is_header_row(tr) or (ri == 0 and _looks_like_header(tr, exp))
                for ri, tr in enumerate(rows)]

    # ---- D10 -1：三线表框线被改动（只判首片段表头与表体行，续页表头归 A06、断口前末行归 A15） ----
    d10 = False
    for gi in range(8):
        g = tbl_groups[gi]
        frags = g['frags'] if g else []
        for fi, tbl in enumerate(frags):
            rows = _rows(tbl)
            flags = _hdr_flags(gi, tbl)
            is_last_frag = (fi == len(frags) - 1)
            for ri, tr in enumerate(rows):
                is_hdr = flags[ri]
                is_last = (ri == len(rows) - 1)
                for tc in _cells(tr):
                    # 竖线：任何单元格左右都不应有线
                    if not _is_blank(_border(tc, 'left')) or not _is_blank(_border(tc, 'right')):
                        d10 = True
                    if is_hdr:
                        # 首片段表头的上粗下细是交付前基线；续页表头由 A06 加分项判定
                        if fi == 0:
                            if not _is_line(_border(tc, 'top'), min_sz=12):
                                d10 = True
                            if not is_last and not _is_line(_border(tc, 'bottom'), max_sz=8):
                                d10 = True
                        continue
                    top = _border(tc, 'top')
                    bot = _border(tc, 'bottom')
                    # 表体行上框线只允许无线或表头下方的细线
                    if not _is_blank(top) and not _is_line(top, max_sz=8):
                        d10 = True
                    if is_last:
                        # 表格最后一段的末行须 1.5 磅封口；断口前末行的封口由 A15 判定
                        if is_last_frag and not _is_line(bot, min_sz=12):
                            d10 = True
                    elif not _is_blank(bot):
                        d10 = True
    results.append(("-1：" if d10 else "-0：") + "三线表框线被改动——首片段表头上框线 1.5 磅、表头下框线为细线、每个表最后一段表体末行下框线 1.5 磅，其余单元格四边无线；出现全框线、网格线、竖线或边框加粗均计")

    # ---- D11 -1：整体水平居中与固定列宽布局被改动（首片段列宽基线；续页列宽归 A06） ----
    d11 = False
    for gi in range(8):
        g = tbl_groups[gi]
        frags = g['frags'] if g else []
        for fi, tbl in enumerate(frags):
            jc = _tblpr(tbl, 'w:jc')
            lay = _tblpr(tbl, 'w:tblLayout')
            if jc is None or jc.get(qn('w:val')) != 'center':
                d11 = True
            if lay is None or lay.get(qn('w:type')) != 'fixed':
                d11 = True
            if fi == 0 and _grid(tbl) != [_T_GRID[gi]] * _T_COLS[gi]:
                d11 = True
    results.append(("-1：" if d11 else "-0：") + "表格未保持整体水平居中与固定列宽布局，或首片段列宽被重新分配（6 列表每列约 2.44cm、5 列表每列约 2.93cm、7 列表每列约 2.09cm）")

    # ---- D12 -1：单元格白色底纹被改动（首片段表头 + 全部表体格；续页表头底纹归 A05） ----
    d12 = False
    for gi in range(8):
        g = tbl_groups[gi]
        frags = g['frags'] if g else []
        for fi, tbl in enumerate(frags):
            flags = _hdr_flags(gi, tbl)
            for ri, tr in enumerate(_rows(tbl)):
                for tc in _cells(tr):
                    fill = (_shd_fill(tc) or '').upper()
                    if flags[ri]:
                        # 首片段表头单元格应显式保持白色底纹
                        if fi == 0 and fill != 'FFFFFF':
                            d12 = True
                    elif fill not in ('', 'FFFFFF', 'AUTO'):
                        d12 = True
    results.append(("-1：" if d12 else "-0：") + "单元格白色底纹被改为其他填充色或被清除")

    # ---- D13 -1：表 1～表 8 的列数被改动（首片段基线；续页列数归 A06 的逐列一致） ----
    d13 = False
    for gi in range(8):
        g = tbl_groups[gi]
        if not g or not g['frags']:
            d13 = True
            continue
        tbl = g['frags'][0]
        if len(_grid(tbl)) != _T_COLS[gi]:
            d13 = True
        for tr in _rows(tbl):
            if len(_cells(tr)) != _T_COLS[gi]:
                d13 = True
    results.append(("-1：" if d13 else "-0：") + "表 1～表 8 的列数被改动（依次应为 6、6、5、5、5、6、6、7 列）")

    # ---- D14 -1：表 3、表 5 被拆分／被插入重复表头行／被添加“续下页” ----
    d14 = False
    for gi in (2, 4):
        g = tbl_groups[gi]
        if not g:
            continue
        if len(g['frags']) > 1:
            d14 = True
        n_hdr = sum(1 for tbl in g['frags'] for tr in _rows(tbl)
                    if _is_header_row(tr) or _looks_like_header(tr, _T_HEADERS[gi]))
        if n_hdr > 1:
            d14 = True
        if _marks_in(gi):
            d14 = True
    results.append(("-1：" if d14 else "-0：") + "表 3、表 5 被拆分成多段、被插入重复表头行或被添加“续下页”")

    # ---- D15 -1：为全部表格行统一勾选“禁止行跨页断行” ----
    n_cant = 0
    for gi in range(8):
        g = tbl_groups[gi]
        for tbl in (g['frags'] if g else []):
            for tr in _rows(tbl):
                if _has_cant_split(tr):
                    n_cant += 1
    d15 = n_cant > 0
    results.append(("-1：" if d15 else "-0：") + "为全部表格行统一勾选“禁止行跨页断行”（原文档 274 行均未设置该属性），使整篇表格的分页结果被动改变")

    # ---- D16 -3：表格内的数值与文字被删除、改写或新增（含首片段表头文字；续页表头文字归 A03） ----
    cellblob = _norm('\n'.join(seq_all))
    d16 = any(_norm(s) not in cellblob for s in _CELL_SENTINELS)
    for gi in range(8):
        fh = first_hdr[gi]
        if fh is None or fh[1] is None or _row_texts(fh[1]) != tuple(_T_HEADERS[gi]):
            d16 = True
    results.append(("-3：" if d16 else "-0：") + "删除、改写或新增表格内的数值与文字（如“151 (56.34)”“−3.216”“<0.001”“1.36 (1.17, 1.58)”“加入服务接触与随访记录”等），或改写首片段表头的逐列文字")

    # ---- D17 -3：正文段落、表题、表注的文字内容被删除或改写 ----
    parablob = '\n'.join(ptext[id(p)] for p in paras)
    d17 = False
    for s in [_H1] + list(_H2) + list(_H3) + list(_BODY_HEADS) + list(_CAPTIONS) + list(_NOTES):
        if s not in parablob:
            d17 = True
    results.append(("-3：" if d17 else "-0：") + "删除或改写正文段落、表题、表注的文字内容")

    # ---- D18 -3：表格数据行发生丢失、重复或顺序错乱 ----
    d18 = not rows_ok
    results.append(("-3：" if d18 else "-0：") + "表格数据行发生丢失、重复或顺序错乱（应为 266 行、各表内部顺序不变）")

    # ---- D19 -1：首片段表头单元格内的换行被合并或删除（续页表头换行归 A07） ----
    d19 = False
    for gi in range(8):
        fh = first_hdr[gi]
        if fh is None or fh[1] is None:
            continue
        for tc in _cells(fh[1]):
            if '(n=' in _norm(_plain_text(tc)) and tc.find('.//' + qn('w:br')) is None:
                d19 = True
    results.append(("-1：" if d19 else "-0：") + "首片段表头单元格内的换行被合并或删除，使“总计”与“(n=268)”一类两行写法挤成一行")

    # ---- D20 -1：表格编号与表题的对应关系被改动 ----
    nos = [g['no'] for g in groups]
    d20 = (nos != [1, 2, 3, 4, 5, 6, 7, 8])
    for gi in range(8):
        g = tbl_groups[gi]
        if not g or _norm(g['cap_text']) != _CAPTIONS[gi]:
            d20 = True
    results.append(("-1：" if d20 else "-0：") + "表格编号与表题的对应关系被改动（表 1～表 8 应顺序连续、编号不重复、不新增或删表）")

    # ---- D21 -1：表题未在表格上方或表注未在表格下方 ----
    d21 = not pos_ok
    results.append(("-1：" if d21 else "-0：") + "8 条表题（“表 1 研究对象基础构成及数字随访接受度分布”至“表 8 稳健性检验与替代定义分析”）未位于各自表格上方，或 8 条表注（“注：……”）未位于各自表格下方")

    # ---- D22 -1：在表题段落或正文段落中插入空白 run、多余空格或多余空段落 ----
    d22 = False
    for p in h1_ps + h2_ps + h3_ps + body_ps + cap_ps + note_ps:
        for r in p.findall(qn('w:r')):
            ts = r.findall(qn('w:t'))
            if ts and all(not (t.text or '').strip() for t in ts):
                d22 = True
        if re.search(r'　| {2,}', ''.join(t.text or '' for t in p.iter(qn('w:t')))):
            d22 = True
    prev_blank = False
    for it in items:
        blank = (it['kind'] == 'p' and not _norm(it['text']) and not _norm(it['float']))
        if blank and prev_blank:
            d22 = True
        prev_blank = blank
    results.append(("-1：" if d22 else "-0：") + "在表题段落或正文段落中插入空白 run、多余空格或多余空段落（表体片段之间的空段落归 A16）")

    # ---- D23 -1：新增页眉、页脚或页码 ----
    n_sect = len(body.findall('.//' + qn('w:sectPr')))
    n_hdr_ref = len(body.findall('.//' + qn('w:headerReference')))
    n_ftr_ref = len(body.findall('.//' + qn('w:footerReference')))
    has_page_field = False
    for node in list(body.iter(qn('w:instrText'))) + list(body.iter(qn('w:fldSimple'))):
        txt = (node.text or '') + (node.get(qn('w:instr')) or '')
        if 'PAGE' in txt.upper():
            has_page_field = True
    d23 = (n_hdr_ref > 0 or n_ftr_ref > 0 or has_page_field or n_sect != 1)
    results.append(("-1：" if d23 else "-0：") + "新增页眉、页脚或页码（全篇原本无页眉页脚，仅一个节）")

    # ---- D24 -1：页面设置被改动 ----
    sect = body.find(qn('w:sectPr'))
    pg = sect.find(qn('w:pgSz')) if sect is not None else None
    mar = sect.find(qn('w:pgMar')) if sect is not None else None
    cols = sect.find(qn('w:cols')) if sect is not None else None
    d24 = False
    if pg is None or _int_attr(pg, 'w:w') != _PG_W or _int_attr(pg, 'w:h') != _PG_H:
        d24 = True
    if mar is None or any(_int_attr(mar, 'w:' + k) != v for k, v in _PG_MAR.items()):
        d24 = True
    if cols is not None and _int_attr(cols, 'w:num', 1) != 1:
        d24 = True
    results.append(("-1：" if d24 else "-0：") + "改动页面设置——A4 纸（210mm×297mm）、上下页边距 25.4mm、左右页边距 31.75mm、单栏排版")

    # ---------------- 汇总 ----------------
    if len(results) != len(_MAX_DELTAS_ORDER):
        raise RuntimeError("细则条目数不一致：results=%d，基准=%d"
                           % (len(results), len(_MAX_DELTAS_ORDER)))

    total = 0
    dim2_items = []
    for raw, max_delta in zip(results, _MAX_DELTAS_ORDER):
        head, _, rule = raw.partition("：")
        got = int(head.replace("+", "").replace("-", "").strip())
        hit = got > 0          # 加分项：命中得满分；减分项：命中触发扣分
        delta = max_delta if hit else 0
        total += delta
        dim2_items.append({
            "rule": rule,
            "max_delta": max_delta,
            "delta": delta,
            "hit": hit,
        })

    total = max(0, min(MAX_SCORE_TOTAL, total))

    result["dim2_items"] = dim2_items
    result["total_score"] = total
    result["max_score"] = MAX_SCORE_TOTAL
    return result


if __name__ == "__main__":
    _dir = sys.argv[1] if len(sys.argv) > 1 else os.path.dirname(os.path.abspath(__file__))
    try:
        print(json.dumps(evaluate(_dir), ensure_ascii=False, indent=2))
    except Exception:
        traceback.print_exc()
        sys.exit(1)
