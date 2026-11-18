#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# pyright: reportUnknownMemberType=false, reportUnknownVariableType=false, reportUnknownArgumentType=false, reportUnknownParameterType=false, reportMissingParameterType=false, reportPrivateUsage=false, reportUnknownLambdaType=false, reportGeneralTypeIssues=false

import os
import re
import sys
import json
import math
from docx import Document
from docx.oxml.ns import qn

# 说明：按接口统一约定，本模块不修改全局 sys.stdout，也不 print 主结果。
# 主结果统一由 evaluate(dir_path) -> dict 返回；入参为"被评估文件所在目录的路径"，
# 由脚本自己在该目录内定位并打开被评估的 .docx。本文件仅在 __main__ 下作为
# 本地调试入口，才把 dict 序列化为 JSON 打印。
#
# 被评测任务（029）：按截图新建"培训记录表"Word 表格，满足指令指定的厘米级列宽、
# 仿宋小三，并交付可编辑的空白待填表。评分细则见同目录 细则.txt。

SCRIPT_ID = "029"
# 维度二所有评分项 max_delta 之和；用于维度一未通过 / 脚本异常时兜底
MAX_SCORE_TOTAL = 42

# 维度二所有评分项按"结果追加顺序"的 max_delta 列表；末端把 results 里的
# "+N：..." 字符串反解为结构化 dim2_items 时用。顺序须与 _evaluate_impl 中
# results.append(...) 的顺序严格一致；共 32 项，累加 = MAX_SCORE_TOTAL。
_MAX_DELTAS_ORDER = [
    3,  # 01 8行表格：行序与14处标签逐字一致
    1,  # 02 第1行8格：4标签+4填写空格
    1,  # 03 第2行2格
    1,  # 04 第3行6格
    1,  # 05 第4行4格（末格宽合并）
    1,  # 06 第5、6行各2格
    1,  # 07 第7、8行各2格
    1,  # 08 单线实线黑框
    1,  # 09 空白待填表（未编造数据/占位符）
    3,  # 10 第一列8行全部2.4cm
    1,  # 11 第1行第2格2.4cm
    1,  # 12 第1行第3格1.45cm
    1,  # 13 第1行第4格1.2cm
    1,  # 14 第1行第5格1.45cm
    1,  # 15 第1行第6、7、8格2.4cm
    3,  # 16 第3行6格 2.4/2.6/2.85/2.85/2.65/2.65
    1,  # 17 第4行前3格 2.4/2.6/2.85 + 末格为余量合并格
    1,  # 18 第2、5、6、7、8行第2格为通栏合并格
    3,  # 19 各行宽度合计一致、右边线对齐
    1,  # 20 列宽以固定磅值(dxa)写入
    1,  # 21 全文仿宋（中西文）
    1,  # 22 全文小三（15磅）
    1,  # 23 无加粗/倾斜/下划线/彩色
    1,  # 24 单元格文字垂直居中
    1,  # 25 标签文字水平居中
    3,  # 26 两处意见栏竖排
    1,  # 27 "学习完成情况"窄列内折为两行
    1,  # 28 第7、8行行高明显大于其他行
    1,  # 29 行高为"最小值"而非"固定值"
    1,  # 30 8行同页且表格位于页面上部
    1,  # 31 未添加原图不存在的元素
    1,  # 32 表格左边线与正文左边距对齐
]
assert sum(_MAX_DELTAS_ORDER) == MAX_SCORE_TOTAL

# ========== 目标规格（全部来自指令原文与原图） ==========
TWIPS_PER_CM = 566.9291338582677   # 1cm = 566.93 dxa(twips)
EMU_PER_CM = 360000.0
W_TOL = 0.05          # 列宽判定容差：±0.05cm 内视为达标
SUM_TOL = 0.15        # 各行宽度合计一致的极差容差（指令给的第3行 6 格合计 16.0，
                      # 第1行合计 16.1，本身相差 0.1，故极差容差放到 0.15）
IND_TOL = 0.10        # 表格左缩进容差
SMALL3_PT = 15.0      # 小三 = 15 磅
PT_TOL = 0.5
DEFAULT_TCMAR_CM = 0.19   # Word 默认单元格左右边距 108dxa
LINE_FACTOR = 1.30        # 单倍行距下一行文字占用的高度 ≈ 字号 × 1.3

FANGSONG_NAMES = ('仿宋', 'FangSong', '仿宋_GB2312', 'FangSong_GB2312', 'STFangsong')
BLACK_COLORS = (None, 'auto', '000000')

# 8 行首列标签（行序）
FIRST_COL_LABELS = ['员工姓名', '培训日期', '培训目标', '考核结果',
                    '评定方式', '学习完成情况', '直属主管意见', '人力资源意见']
# 各行应有的格数
ROW_CELL_COUNTS = [8, 2, 6, 4, 2, 2, 2, 2]
# 行内标签所在格序（0基）与文字
ROW_LABELS = {
    0: {0: '员工姓名', 2: '所属部门', 4: '性别', 6: '任职岗位'},
    1: {0: '培训日期'},
    2: {0: '培训目标', 2: '培训项目', 4: '学习时长'},
    3: {0: '考核结果', 2: '能力评定'},
    4: {0: '评定方式'},
    5: {0: '学习完成情况'},
    6: {0: '直属主管意见'},
    7: {0: '人力资源意见'},
}
ALL_LABELS = ['员工姓名', '所属部门', '性别', '任职岗位', '培训日期', '培训目标',
              '培训项目', '学习时长', '考核结果', '能力评定', '评定方式',
              '学习完成情况', '直属主管意见', '人力资源意见']
# 指令给出的列宽
FIRST_COL_W = 2.4
ROW0_W = [2.4, 2.4, 1.45, 1.2, 1.45, 2.4, 2.4, 2.4]
ROW2_W = [2.4, 2.6, 2.85, 2.85, 2.65, 2.65]
ROW3_W_HEAD = [2.4, 2.6, 2.85]
# 竖排文字方向取值
VERTICAL_DIRS = ('tbRl', 'tbRlV', 'btLr', 'tbLrV', 'lrTbV')

# ========== 通用小工具 ==========
def _twips_to_cm(val):
    """把 OOXML 的 dxa(twips) 字符串转成厘米；无法解析返回 None。"""
    try:
        return int(val) / TWIPS_PER_CM
    except (TypeError, ValueError):
        return None


def _close(a, b, tol=W_TOL):
    """两个厘米值是否在容差内相等。"""
    return a is not None and b is not None and abs(a - b) <= tol


def _fangsong(name):
    """字体名是否属于"仿宋"族。"""
    return bool(name) and any(k in name for k in FANGSONG_NAMES)


def _on(rPr, tag):
    """判定 w:b / w:i 一类开关是否真的打开：元素存在且 w:val 不是 0/false/none。"""
    if rPr is None:
        return False
    el = rPr.find(qn(tag))
    if el is None:
        return False
    v = el.get(qn('w:val'))
    return v is None or str(v).lower() not in ('0', 'false', 'none', 'off')


def _text_width_cm(text, pt):
    """按字号估算一行文字的排版宽度：中日韩字符按全宽(1em)，其余按半宽(0.5em)。"""
    em = pt / 28.3465  # 1pt = 1/28.3465 cm
    w = 0.0
    for ch in text:
        w += em if re.match(r'[\u2E80-\u9FFF\uFF00-\uFFEF\u3000-\u303F]', ch) else em * 0.5
    return w


def _line_height_cm(pt):
    """单倍行距下一行文字占用的高度（含行间余量）。"""
    return pt * LINE_FACTOR / 28.3465

# ========== 字体继承链解析（与 008 同口径：run → 段落样式(basedOn 链) → docDefaults） ==========
def _style_by_id(doc, sid):
    for st in doc.styles.element.findall(qn('w:style')):
        if st.get(qn('w:styleId')) == sid:
            return st
    return None


def _rpr_chain(doc, p_el):
    """按 Word 的继承顺序产出段落所处的各级 w:rPr：段落样式(含 basedOn 上溯) → docDefaults。
       run 自身的 rPr 由调用方先行判定，本函数只负责"上层"。"""
    style_id = None
    pPr = p_el.find(qn('w:pPr')) if p_el is not None else None
    if pPr is not None:
        pStyle = pPr.find(qn('w:pStyle'))
        if pStyle is not None:
            style_id = pStyle.get(qn('w:val'))
    seen = set()
    while style_id and style_id not in seen:
        seen.add(style_id)
        st = _style_by_id(doc, style_id)
        if st is None:
            break
        rpr = st.find(qn('w:rPr'))
        if rpr is not None:
            yield rpr
        base = st.find(qn('w:basedOn'))
        style_id = base.get(qn('w:val')) if base is not None else None

    dd = doc.styles.element.find(qn('w:docDefaults'))
    if dd is not None:
        rpd = dd.find(qn('w:rPrDefault'))
        if rpd is not None:
            rpr = rpd.find(qn('w:rPr'))
            if rpr is not None:
                yield rpr


def _resolve_effective_eastasia(doc, r_el, p_el):
    """解析 run 中文字符实际生效的东亚字体（打开办公软件看到的那个）；都没有则 None。"""
    for rpr in [r_el.find(qn('w:rPr'))] + list(_rpr_chain(doc, p_el)):
        if rpr is None:
            continue
        rf = rpr.find(qn('w:rFonts'))
        if rf is not None:
            ea = rf.get(qn('w:eastAsia'))
            if ea:
                return ea
    return None

def _resolve_effective_ascii(doc, r_el, p_el):
    """解析 run 对西文/数字实际生效的字体（w:ascii，回落 w:hAnsi）；都没有则 None。"""
    for rpr in [r_el.find(qn('w:rPr'))] + list(_rpr_chain(doc, p_el)):
        if rpr is None:
            continue
        rf = rpr.find(qn('w:rFonts'))
        if rf is not None:
            v = rf.get(qn('w:ascii')) or rf.get(qn('w:hAnsi'))
            if v:
                return v
    return None


def _resolve_effective_size_pt(doc, r_el, p_el):
    """解析 run 实际生效的字号（磅）；都没有则 None。"""
    for rpr in [r_el.find(qn('w:rPr'))] + list(_rpr_chain(doc, p_el)):
        if rpr is None:
            continue
        sz = rpr.find(qn('w:sz'))
        if sz is not None:
            try:
                return int(sz.get(qn('w:val'))) / 2.0
            except (TypeError, ValueError):
                pass
    return None


# ========== 表格结构读取（一律走 w:tc，不用 row.cells，避免合并格被展开重复） ==========
def _tcs(tr):
    return tr.findall(qn('w:tc'))


def _tc_pr(tc):
    return tc.find(qn('w:tcPr'))


def _tc_width(tc):
    """返回 (type, cm)；无 tcW 返回 (None, None)。"""
    tcPr = _tc_pr(tc)
    if tcPr is None:
        return None, None
    w = tcPr.find(qn('w:tcW'))
    if w is None:
        return None, None
    wtype = w.get(qn('w:type'))
    if wtype == 'dxa':
        return wtype, _twips_to_cm(w.get(qn('w:w')))
    return wtype, None

def _tc_span(tc):
    tcPr = _tc_pr(tc)
    if tcPr is None:
        return 1
    gs = tcPr.find(qn('w:gridSpan'))
    if gs is None:
        return 1
    try:
        return int(gs.get(qn('w:val')))
    except (TypeError, ValueError):
        return 1


def _tc_paras(tc):
    return tc.findall(qn('w:p'))


def _p_text(p_el):
    return ''.join(t.text or '' for t in p_el.findall('.//' + qn('w:t')))


def _tc_text(tc):
    return ''.join(_p_text(p) for p in _tc_paras(tc)).strip()


def _tc_val(tc, tag):
    """读取 tcPr 下某个单值元素的 w:val（如 vAlign / textDirection）。"""
    tcPr = _tc_pr(tc)
    if tcPr is None:
        return None
    el = tcPr.find(qn(tag))
    return el.get(qn('w:val')) if el is not None else None


def _tc_shd_fill(tc):
    tcPr = _tc_pr(tc)
    if tcPr is None:
        return None
    shd = tcPr.find(qn('w:shd'))
    return shd.get(qn('w:fill')) if shd is not None else None


def _tr_height(tr):
    """返回 (cm, hRule)；未显式设置行高返回 (None, None)。"""
    trPr = tr.find(qn('w:trPr'))
    if trPr is None:
        return None, None
    th = trPr.find(qn('w:trHeight'))
    if th is None:
        return None, None
    return _twips_to_cm(th.get(qn('w:val'))), th.get(qn('w:hRule'))

def _row_widths(tr):
    """一行内各格的宽度（cm）列表，缺失记 None。"""
    return [_tc_width(tc)[1] for tc in _tcs(tr)]


def _row_sum(tr):
    ws = [w for w in _row_widths(tr) if w is not None]
    return sum(ws) if ws else None


def _tbl_ind_cm(tbl):
    tblPr = tbl.find(qn('w:tblPr'))
    if tblPr is None:
        return 0.0
    ind = tblPr.find(qn('w:tblInd'))
    if ind is None:
        return 0.0
    if ind.get(qn('w:type')) != 'dxa':
        return None
    v = _twips_to_cm(ind.get(qn('w:w')))
    return 0.0 if v is None else v


def _tbl_w(tbl):
    """返回 (type, cm)。"""
    tblPr = tbl.find(qn('w:tblPr'))
    if tblPr is None:
        return None, None
    w = tblPr.find(qn('w:tblW'))
    if w is None:
        return None, None
    wtype = w.get(qn('w:type'))
    return wtype, (_twips_to_cm(w.get(qn('w:w'))) if wtype == 'dxa' else None)


def _cell_font_pt(doc, tc):
    """取单元格内首个有文字的 run 的生效字号；取不到按小三兜底（仅用于几何估算）。"""
    for p in _tc_paras(tc):
        for r in p.findall(qn('w:r')):
            if _p_text(r).strip() or r.findall(qn('w:t')):
                pt = _resolve_effective_size_pt(doc, r, p)
                if pt:
                    return pt
    return SMALL3_PT

def _p_line_texts(p_el):
    """把一个段落按手动换行(w:br)切成若干"视觉行"的文字片段。"""
    segs = ['']
    for node in p_el.iter():
        if node.tag == qn('w:t'):
            segs[-1] += node.text or ''
        elif node.tag == qn('w:br'):
            segs.append('')
    return segs


def _cell_line_texts(tc):
    """单元格内按段落 + 手动换行切分出的视觉行片段（尚未按列宽折行）。"""
    out = []
    for p in _tc_paras(tc):
        out.extend(_p_line_texts(p))
    return out


def _cell_lines(doc, tc):
    """估算单元格内文字实际占用的行数：先按段落/手动换行切片，再按可用宽度折行。"""
    _, w = _tc_width(tc)
    usable = (w - 2 * DEFAULT_TCMAR_CM) if w else None
    pt = _cell_font_pt(doc, tc)
    lines = 0
    for seg in _cell_line_texts(tc):
        if not seg.strip():
            lines += 1
            continue
        need = _text_width_cm(seg, pt)
        if usable and usable > 0:
            lines += max(1, int(math.ceil(need / usable - 1e-6)))
        else:
            lines += 1
    return max(lines, 1)


def _cell_need_height_cm(doc, tc):
    """估算单元格内容所需的最小净高度。"""
    return _cell_lines(doc, tc) * _line_height_cm(_cell_font_pt(doc, tc))


def _row_need_height_cm(doc, tr):
    need = 0.0
    for tc in _tcs(tr):
        need = max(need, _cell_need_height_cm(doc, tc))
    return need


def _open_word_document(file_path):
    """打开 .docx 文档（python-docx 仅支持 OOXML/.docx）."""
    return Document(file_path)


def evaluate(dir_path):
    """按 §2.2 约定，返回结构化 dict；不 print、不改 sys.stdout、不 sys.exit。
    入参 dir_path 为"被评估文件所在目录的路径"，脚本自己在该目录里定位并打开被评估的
    .docx（忽略以 ~$ 开头的 Word 临时锁文件）。脚本自身崩溃（含目录不存在、
    未发现文档等）由本函数捕获并返回 status='error'。"""
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
            if f.lower().endswith('.docx') and not f.startswith('~$')
            and os.path.isfile(os.path.join(dir_path, f))
        )
        if not cands:
            raise FileNotFoundError(f"目录中未找到 .docx 文件：{dir_path}")
        file_path = os.path.join(dir_path, cands[0])
        result["file_name"] = os.path.basename(file_path)
        return _evaluate_impl(file_path, result)
    except FileNotFoundError as exc:
        # 找不到被评估文件：error 只回传纯文本原因（不带异常类名），维度一根本没进入
        # 判定，dim1_pass 与 dim1_reason 一律留空（空字符串，而非 False），
        # dim2 不出项，满分一并归零。
        result["status"] = "error"
        result["error"] = str(exc)
        result["file_name"] = ""
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
        if not result["dim1_reason"]:
            result["dim1_reason"] = "文档无法打开或解析"
        result["dim2_items"] = []
        result["total_score"] = 0
        return result

def _evaluate_impl(file_path, result):
    """维度一 + 维度二的评估主体；内部构造 "+N：说明" 形式的 results 列表，
    末尾按 _MAX_DELTAS_ORDER 反解成结构化 dim2_items 填入 result。"""
    doc = _open_word_document(file_path)
    body = doc.element.body
    tables = doc.tables
    results = []

    # ---------- 维度1：可用与可修改性 ----------
    def _fail_dim1(reason):
        result["dim1_pass"] = False
        result["dim1_reason"] = reason
        result["dim2_items"] = []
        result["total_score"] = 0
        return result

    # P2：必须存在真实 Word 表格对象（可选中单元格），而非图片/文本框/伪表格
    if not tables:
        pic_cnt = len(body.findall('.//' + qn('a:blip'))) + len(body.findall('.//' + qn('w:pict')))
        if pic_cnt:
            return _fail_dim1(f"未发现真实 Word 表格对象，文档内仅有 {pic_cnt} 个图片/图形（疑似以图片代替表格）")
        return _fail_dim1("未发现任何 Word 表格对象")

    # 取行数最多的表作为主表（培训记录表）
    main_tbl = max((t._tbl for t in tables), key=lambda tb: len(tb.findall(qn('w:tr'))))
    trs = main_tbl.findall(qn('w:tr'))
    if len(trs) < 2:
        return _fail_dim1(f"主表仅 {len(trs)} 行，未构成表格")

    # P4：表格未被拆散成多个表
    if len(tables) > 1:
        return _fail_dim1(f"表格被拆散为 {len(tables)} 个独立表格，未构成一张完整的培训记录表")

    # P3：无乱码 / 无被固定行高裁切的文字
    bad_chars = [c for c in main_tbl.itertext() if c and re.search(r'[\uFFFD\uE000-\uF8FF]', c)]
    if bad_chars:
        return _fail_dim1("表格内出现乱码字符（U+FFFD 或私用区字符）")
    clipped = []
    for i, tr in enumerate(trs):
        h, rule = _tr_height(tr)
        if rule == 'exact' and h is not None and _row_need_height_cm(doc, tr) > h + 0.08:
            clipped.append(i + 1)
    if clipped:
        return _fail_dim1(f"第 {clipped} 行行高为固定值且小于文字所需高度，文字被裁切")

    # P1/P4 补充：表格右边缘不得越出纸张可打印范围
    sec = doc.sections[0]
    page_w = (sec.page_width or 0) / EMU_PER_CM
    page_h = (sec.page_height or 0) / EMU_PER_CM
    m_left = (sec.left_margin or 0) / EMU_PER_CM
    m_right = (sec.right_margin or 0) / EMU_PER_CM
    m_top = (sec.top_margin or 0) / EMU_PER_CM
    m_bottom = (sec.bottom_margin or 0) / EMU_PER_CM
    ind = _tbl_ind_cm(main_tbl) or 0.0
    row_sums = [s for s in (_row_sum(tr) for tr in trs) if s is not None]
    tbl_total_w = max(row_sums) if row_sums else (_tbl_w(main_tbl)[1] or 0.0)
    if page_w and m_left + ind + tbl_total_w > page_w - 0.5:
        return _fail_dim1(
            f"表格右边缘越出纸张可打印范围（左边距{m_left:.2f}+表宽{tbl_total_w:.2f}cm > 纸宽{page_w:.2f}cm）")

    over_margin = page_w and (m_left + ind + tbl_total_w > page_w - m_right + 0.05)
    result["dim1_pass"] = True
    result["dim1_reason"] = (
        "文件可正常打开，表格为真实 Word 表格对象（单元格可编辑），无乱码、无裁切，未拆表未断裂"
        + (f"；提示：表宽 {tbl_total_w:.2f}cm 超出正文区宽度 {page_w - m_left - m_right:.2f}cm，"
           f"表格右侧压入右页边距（仍在纸张内，不判不可用）" if over_margin else "")
    )

    # ---------- 维度2：完成度评分 ----------
    def add(hit, delta, text):
        """按 "+N：说明" 追加一条结果；未命中记 +0。顺序须与 _MAX_DELTAS_ORDER 一致。"""
        results.append(f"+{delta if hit else 0}：{text}")

    def row_tcs(i):
        return _tcs(trs[i]) if i < len(trs) else []

    def texts(i):
        return [_tc_text(tc) for tc in row_tcs(i)]

    def labels_ok(i):
        """本行标签格文字与原图逐字一致、且其余格为空。"""
        ts = texts(i)
        if len(ts) != ROW_CELL_COUNTS[i]:
            return False
        for j, t in enumerate(ts):
            want = ROW_LABELS.get(i, {}).get(j)
            if want is None:
                if t != '':
                    return False
            elif t != want:
                return False
        return True

    n_rows = len(trs)
    first_col = [(_tc_text(_tcs(tr)[0]) if _tcs(tr) else '') for tr in trs]
    all_cell_texts = [_tc_text(tc) for tr in trs for tc in _tcs(tr)]
    hit_labels = [x for x in ALL_LABELS if x in all_cell_texts]

    # 01 8 行表格：行序与 14 处标签逐字一致
    add(n_rows == 8 and first_col == FIRST_COL_LABELS and len(hit_labels) == len(ALL_LABELS), 3,
        f"生成8行表格且行序、14处标签与原图逐字一致（实测{n_rows}行、标签命中{len(hit_labels)}/14）")

    # 02-07 各行格数与标签/空格分布
    add(labels_ok(0), 1, f"第1行为8格：4个标签各带独立填写空格（实测{len(row_tcs(0))}格）")
    add(labels_ok(1), 1, f"第2行为2格：培训日期+通栏空格（实测{len(row_tcs(1))}格）")
    add(labels_ok(2), 1, f"第3行为6格：培训目标/培训项目/学习时长各带空格（实测{len(row_tcs(2))}格）")
    r3 = row_tcs(3)
    add(labels_ok(3) and len(r3) == 4 and _tc_span(r3[3]) > 1, 1,
        f"第4行为4格且末格为宽合并格（实测{len(r3)}格，末格跨{_tc_span(r3[3]) if len(r3) == 4 else 0}列）")
    add(labels_ok(4) and labels_ok(5), 1,
        f"第5、6行各为2格：评定方式、学习完成情况（实测{len(row_tcs(4))}/{len(row_tcs(5))}格）")
    add(labels_ok(6) and labels_ok(7), 1,
        f"第7、8行各为2格：直属主管意见、人力资源意见（实测{len(row_tcs(6))}/{len(row_tcs(7))}格）")

    # 08 单线实线黑框：表级六向边框均为 single，且无单元格把边框改成无线
    tblPr = main_tbl.find(qn('w:tblPr'))
    tbl_borders = tblPr.find(qn('w:tblBorders')) if tblPr is not None else None
    border_sides = ('top', 'left', 'bottom', 'right', 'insideH', 'insideV')
    if tbl_borders is not None:
        sides_ok = True
        max_sz = 0
        for side in border_sides:
            el = tbl_borders.find(qn('w:' + side))
            if el is None or el.get(qn('w:val')) != 'single':
                sides_ok = False
                break
            if el.get(qn('w:color')) not in BLACK_COLORS:
                sides_ok = False
                break
            try:
                max_sz = max(max_sz, int(el.get(qn('w:sz')) or 0))
            except ValueError:
                pass
        border_ok = sides_ok and max_sz <= 12  # 12(1.5磅)以内视为单线细框
    else:
        # 未写表级边框时，套用"网格型"表样式同样呈现单线黑框
        style_name = (tables[0].style.name if tables[0].style is not None else '') or ''
        border_ok = 'Grid' in style_name or '网格' in style_name
        max_sz = 4
    for tr in trs:
        for tc in _tcs(tr):
            tcPr = _tc_pr(tc)
            bd = tcPr.find(qn('w:tcBorders')) if tcPr is not None else None
            if bd is None:
                continue
            for el in bd:
                if el.get(qn('w:val')) in ('nil', 'none'):
                    border_ok = False
    add(border_ok, 1, f"表格内外框线均为单线实线黑框（线宽{max_sz / 8:.2f}磅）")

    # 09 空白待填表：除原图 14 处标签外未编造数据/占位符
    extra_texts = sorted({t for t in all_cell_texts if t and t not in ALL_LABELS})
    add(not extra_texts, 1,
        f"交付为空白待填表，未编造示例数据或占位符（多余文字{len(extra_texts)}处"
        + (f"：{'、'.join(extra_texts[:3])}" if extra_texts else "") + "）")

    def w_at(i, j):
        ws = _row_widths(trs[i]) if i < len(trs) else []
        return ws[j] if j < len(ws) else None

    def fmt(v):
        return f"{v:.2f}" if isinstance(v, float) else "缺"

    # 10 第一列 8 行全部 2.4cm
    first_ws = [(_tc_width(_tcs(tr)[0])[1] if _tcs(tr) else None) for tr in trs]
    add(len(first_ws) == 8 and all(_close(w, FIRST_COL_W) for w in first_ws), 3,
        "第一列8行首格宽度全部为2.4cm（实测 " + "/".join(fmt(w) for w in first_ws) + "）")

    # 11-15 第 1 行逐格宽度
    add(_close(w_at(0, 1), 2.4), 1, f"第1行第2格=2.4cm（实测{fmt(w_at(0, 1))}）")
    add(_close(w_at(0, 2), 1.45), 1, f"第1行第3格=1.45cm（实测{fmt(w_at(0, 2))}）")
    add(_close(w_at(0, 3), 1.2), 1, f"第1行第4格=1.2cm（实测{fmt(w_at(0, 3))}）")
    add(_close(w_at(0, 4), 1.45), 1, f"第1行第5格=1.45cm（实测{fmt(w_at(0, 4))}）")
    w0 = _row_widths(trs[0]) if trs else []
    add(len(w0) >= 8 and all(_close(w, 2.4) for w in w0[5:8]), 1,
        "第1行第6、7、8格均=2.4cm（实测 " + "/".join(fmt(w) for w in w0[5:8]) + "）")

    # 16 第 3 行 6 格宽度
    w2 = _row_widths(trs[2]) if len(trs) > 2 else []
    add(len(w2) == 6 and all(_close(a, b) for a, b in zip(w2, ROW2_W)), 3,
        "第3行6格宽度依次=2.4/2.6/2.85/2.85/2.65/2.65cm（实测 "
        + "/".join(fmt(w) for w in w2) + "）")

    # 17 第 4 行前 3 格 + 末格为余量合并格
    w3 = _row_widths(trs[3]) if len(trs) > 3 else []
    r3 = row_tcs(3)
    rest_ok = (len(w3) == 4 and _tc_span(r3[3]) > 1
               and _close(w3[3], tbl_total_w - sum(ROW3_W_HEAD), 0.20))
    add(len(w3) == 4 and all(_close(a, b) for a, b in zip(w3, ROW3_W_HEAD)) and rest_ok, 1,
        "第4行前3格=2.4/2.6/2.85cm且第4格为剩余宽度合并格（实测 "
        + "/".join(fmt(w) for w in w3) + "）")

    # 18 第 2、5、6、7、8 行第 2 格为通栏合并格
    merged_rows = [1, 4, 5, 6, 7]
    merged_ws = []
    merged_ok = True
    for i in merged_rows:
        tcs_i = row_tcs(i)
        wi = w_at(i, 1)
        merged_ws.append(wi)
        if len(tcs_i) != 2 or _tc_span(tcs_i[1]) <= 1 or not _close(wi, tbl_total_w - FIRST_COL_W, 0.10):
            merged_ok = False
    add(merged_ok, 1,
        f"第2、5、6、7、8行第2格为通栏合并格，宽度≈表总宽−2.4cm（应≈{tbl_total_w - FIRST_COL_W:.2f}，实测 "
        + "/".join(fmt(w) for w in merged_ws) + "）")

    # 19 各行宽度合计一致、右边线对齐
    span_ok = bool(row_sums) and (max(row_sums) - min(row_sums)) <= SUM_TOL
    add(span_ok and 15.8 <= (sum(row_sums) / len(row_sums) if row_sums else 0) <= 16.4, 3,
        f"各行宽度合计一致（约16.1cm）、右边线整齐对齐（实测极差{(max(row_sums) - min(row_sums)) if row_sums else 0:.3f}cm）")

    # 20 列宽以固定磅值写入
    wtypes = [_tc_width(tc)[0] for tr in trs for tc in _tcs(tr)]
    tblw_type = _tbl_w(main_tbl)[0]
    add(all(t == 'dxa' for t in wtypes) and tblw_type in ('dxa', None), 1,
        f"列宽以固定磅值(dxa)写入，非auto/百分比（tblW={tblw_type}，异常格{sum(1 for t in wtypes if t != 'dxa')}个）")

    # ---------- C 文字格式 ----------
    text_runs = []   # [(段落元素, run元素)]，仅统计真正有文字的 run
    for tr in trs:
        for tc in _tcs(tr):
            for p in _tc_paras(tc):
                for r in p.findall(qn('w:r')):
                    if _p_text(r).strip():
                        text_runs.append((p, r))

    # 21 全文仿宋（中文与西文/数字均设为仿宋）
    bad_font = []
    for p, r in text_runs:
        ea = _resolve_effective_eastasia(doc, r, p)
        asc = _resolve_effective_ascii(doc, r, p)
        if not _fangsong(ea) or not _fangsong(asc):
            bad_font.append(f"{_p_text(r)}:{ea}/{asc}")
    add(bool(text_runs) and not bad_font, 1,
        f"全部文字（中文与西文/数字）字体为仿宋（不合规run {len(bad_font)}个"
        + (f"，如 {bad_font[0]}" if bad_font else "") + "）")

    # 22 全文小三（15 磅）
    sizes = [_resolve_effective_size_pt(doc, r, p) for p, r in text_runs]
    bad_size = [s for s in sizes if s is None or abs(s - SMALL3_PT) > PT_TOL]
    add(bool(sizes) and not bad_size, 1,
        f"全部文字字号为小三15磅（不合规run {len(bad_size)}个"
        + (f"，实测 {sorted({s for s in bad_size if s})}" if bad_size else "") + "）")

    # 23 无加粗/倾斜/下划线/彩色
    deco = []
    for p, r in text_runs:
        rPr = r.find(qn('w:rPr'))
        for tag, name in (('w:b', '加粗'), ('w:i', '倾斜'), ('w:u', '下划线')):
            if _on(rPr, tag):
                deco.append(name)
        if rPr is not None:
            col = rPr.find(qn('w:color'))
            if col is not None and col.get(qn('w:val')) not in BLACK_COLORS:
                deco.append('彩色' + str(col.get(qn('w:val'))))
    add(not deco, 1,
        f"无加粗/倾斜/下划线/彩色等原图不存在的修饰，字色为黑色或自动（命中{sorted(set(deco))}）")

    # 24 单元格文字垂直居中
    valigns = [_tc_val(tc, 'w:vAlign') for tr in trs for tc in _tcs(tr)]
    add(bool(valigns) and all(v == 'center' for v in valigns), 1,
        f"全部单元格文字垂直居中（非居中{sum(1 for v in valigns if v != 'center')}格）")

    # 25 标签文字水平居中
    not_center = []
    for tr in trs:
        for tc in _tcs(tr):
            for p in _tc_paras(tc):
                if not _p_text(p).strip():
                    continue
                pPr = p.find(qn('w:pPr'))
                jc = pPr.find(qn('w:jc')) if pPr is not None else None
                if jc is None or jc.get(qn('w:val')) not in ('center', 'distribute'):
                    not_center.append(_p_text(p).strip())
    add(not not_center, 1,
        f"标签文字水平居中（未居中{len(not_center)}处"
        + (f"，如 {not_center[0]}" if not_center else "") + "）")

    # ---------- D 按图细节还原 ----------
    def _is_vertical(tc, label):
        """竖排判定：①设置了竖排文字方向；或②每字独占一行（分段或手动换行均可）。"""
        if _tc_val(tc, 'w:textDirection') in VERTICAL_DIRS:
            return True
        chars = [s.strip() for s in _cell_line_texts(tc) if s.strip()]
        return len(chars) == len(label) and all(len(c) == 1 for c in chars)

    # 26 两处意见栏竖排
    vert = []
    for i, label in ((6, '直属主管意见'), (7, '人力资源意见')):
        tcs_i = row_tcs(i)
        vert.append(bool(tcs_i) and _tc_text(tcs_i[0]) == label and _is_vertical(tcs_i[0], label))
    add(all(vert), 3,
        f"直属主管意见/人力资源意见两格按原图竖排、每字一行（实测 主管{'竖排' if vert[0] else '横排'}、"
        f"人资{'竖排' if vert[1] else '横排'}）")

    # 27 "学习完成情况"在 2.4cm 窄列内折为两行
    tcs5 = row_tcs(5)
    if tcs5 and _tc_text(tcs5[0]) == '学习完成情况':
        lines5 = _cell_lines(doc, tcs5[0])
        h5, rule5 = _tr_height(trs[5])
        need5 = _cell_need_height_cm(doc, tcs5[0])
        room_ok = (rule5 != 'exact') or (h5 is not None and h5 + 0.08 >= need5)
        n_br5 = len(tcs5[0].findall('.//' + qn('w:br')))
        hit27 = lines5 == 2 and room_ok
    else:
        lines5, n_br5, hit27 = 0, 0, False
    add(hit27, 1,
        f"\"学习完成情况\"在2.4cm窄列内折为两行且完整显示（实测折{lines5}行、手动换行{n_br5}处）")

    # 28 第 7、8 行行高明显大于其他行
    heights = []
    for tr in trs:
        h, _rule = _tr_height(tr)
        heights.append(h if h is not None else _row_need_height_cm(doc, tr))
    if len(heights) == 8:
        hit28 = min(heights[6], heights[7]) >= 1.2 * max(heights[:6])
    else:
        hit28 = False
    add(hit28, 1,
        "第7、8行行高明显大于其他行，为意见栏留出书写空间（实测 "
        + "/".join(f"{h:.2f}" for h in heights) + "cm）")

    # 29 行高为"最小值"而非"固定值"
    rules = [_tr_height(tr)[1] for tr in trs]
    add(all(r != 'exact' for r in rules), 1,
        f"行高设为\"最小值\"而非\"固定值\"，填写内容变多不裁切（固定值行{sum(1 for r in rules if r == 'exact')}行）")

    # 30 8 行同页、表格位于页面上部
    lead_paras = []
    for child in body.iterchildren():
        if child.tag == qn('w:tbl'):
            break
        if child.tag == qn('w:p'):
            lead_paras.append(_p_text(child).strip())
    lead_h = 0.6 * len(lead_paras)                      # 表格前空段按每段约 0.6cm 估算
    avail_h = page_h - m_top - m_bottom
    need_total = 0.0
    for idx, tr in enumerate(trs):
        h, _rule = _tr_height(tr)
        need_total += max(h or 0.0, _row_need_height_cm(doc, tr))
    # 真实分页只有排版引擎才知道；本项改用纯几何口径判定：8 行所需净高（逐行取
    # 设定行高与内容所需高度的较大值）加上表前段落估高，不超过页面可用净高即视为同页。
    # 但内容所需高度只按"字数折行 × 行高系数"估算，拿不到单元格上下边距、段前段后
    # 间距等构成，实测会明显低估未显式设置行高的行。因此只有当所有行都写明了行高时，
    # 几何口径才足以支撑判定；存在未设行高的行时按"加分项无证据不给分"处理。
    auto_rows = [i + 1 for i, tr in enumerate(trs) if _tr_height(tr)[0] is None]
    if auto_rows:
        one_page = False
        page_src = ("几何口径不足以判定：第%s行未显式设置行高，"
                    "无法据静态属性确定其真实占高" % auto_rows)
    else:
        one_page = (need_total + lead_h) <= avail_h + 0.2
        page_src = (f"几何口径：表高{need_total:.2f}+表前{lead_h:.2f}"
                    f"{'≤' if one_page else '>'}可用{avail_h:.2f}cm")
    top_ok = not any(lead_paras) and len(lead_paras) <= 1
    add(one_page and top_ok, 1,
        f"8行全部容纳在同一页内且表格位于页面上部（{page_src}；表前段落{len(lead_paras)}个）")

    # 31 未添加原图不存在的元素（标题/页眉页脚/页码/水印/说明文字/底纹）
    extras = []
    pic_cnt = len(body.findall('.//' + qn('a:blip'))) + len(body.findall('.//' + qn('w:pict')))
    if pic_cnt:
        extras.append(f"图片{pic_cnt}个")
    body_text = [_p_text(c).strip() for c in body.iterchildren()
                 if c.tag == qn('w:p') and _p_text(c).strip()]
    if body_text:
        extras.append(f"表外文字{len(body_text)}处")
    for s in doc.sections:
        for part, name in ((s.header, '页眉'), (s.footer, '页脚')):
            try:
                el = part._element
            except AttributeError:
                continue
            if any(p.text.strip() for p in part.paragraphs) \
                    or el.findall('.//' + qn('w:pict')) or el.findall('.//' + qn('a:blip')) \
                    or el.findall('.//' + qn('w:fldSimple')) or el.findall('.//' + qn('w:instrText')):
                extras.append(name)
    shades = [f for f in (_tc_shd_fill(tc) for tr in trs for tc in _tcs(tr))
              if f not in (None, 'auto', 'FFFFFF', 'ffffff')]
    if shades:
        extras.append(f"底纹{len(shades)}格")
    add(not extras, 1,
        f"未添加原图不存在的元素：标题/页眉页脚/页码/水印/说明文字/底纹（命中{sorted(set(extras))}）")

    # 32 表格左边线与正文左边距对齐
    float_pos = tblPr.find(qn('w:tblpPr')) if tblPr is not None else None
    add(ind is not None and abs(ind) <= IND_TOL and float_pos is None, 1,
        f"表格左边线与正文左边距对齐、无异常缩进或浮动定位（实测左缩进{(ind or 0):.2f}cm）")

    # ========== 将 results 反解为结构化 dim2_items ==========
    # 逐项按追加顺序对齐 _MAX_DELTAS_ORDER：解析每条 "+N：说明" 的 N 与 rule。
    if len(results) != len(_MAX_DELTAS_ORDER):
        # 顺序表与实际条目不一致：视为脚本内部错误，交给外层兜底
        raise RuntimeError(
            f"dim2 项数不一致：results={len(results)} vs order={len(_MAX_DELTAS_ORDER)}"
        )
    dim2_items = []
    total_score = 0
    for raw, max_delta in zip(results, _MAX_DELTAS_ORDER):
        # 格式约定：'+N：' 前缀 + 规则说明；命中 N=max_delta，未命中 N=0
        head, _, rule = raw.partition("：")
        try:
            delta = int(head.replace("+", "").strip())
        except ValueError:
            delta = 0
        hit = delta > 0
        dim2_items.append({
            "rule": rule,
            "max_delta": max_delta,
            "delta": max_delta if hit else 0,
            "hit": hit,
        })
        total_score += (max_delta if hit else 0)

    result["dim2_items"] = dim2_items
    result["total_score"] = total_score
    result["max_score"] = MAX_SCORE_TOTAL
    return result


if __name__ == "__main__":
    # 本地调试入口：evaluate 只接收"被评估文件所在目录的路径"，由脚本自己在该目录里
    # 定位并打开被评估的 .docx。命令行第 1 个参数可覆盖目录路径；未指定时使用
    # 脚本自身所在目录。
    # 注：evaluate() 的接口约定就是"入参为目录路径"——路径始终由外部显式传入；
    #     这里的默认值仅为本地调试便利，不影响 evaluate 的接口约定。
    if len(sys.argv) >= 2:
        _dir = sys.argv[1]
    else:
        _dir = os.path.dirname(os.path.abspath(__file__))
    _out = evaluate(_dir)
    print(json.dumps(_out, ensure_ascii=False, indent=2))

