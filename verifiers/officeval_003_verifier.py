#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# pyright: reportUnknownMemberType=false, reportUnknownVariableType=false, reportUnknownArgumentType=false, reportUnknownParameterType=false, reportMissingParameterType=false, reportPrivateUsage=false, reportUnknownLambdaType=false, reportGeneralTypeIssues=false

import os
import re
import sys
import json
from docx import Document
from docx.oxml.ns import qn

# 说明：按接口统一约定，本模块不修改全局 sys.stdout，也不 print 主结果。
# 主结果统一由 evaluate(dir_path) -> dict 返回；入参为"脚本所在目录的路径"，
# 由脚本自己在该目录内定位并打开被评估的 .docx。本文件仅在 __main__ 下作为
# 本地调试入口，才把 dict 序列化为 JSON 打印。
#
# 评测对象：把"部门协调会记录"里"会议内容"下方每一行文字放入表格并加横线、删空行的改写结果。
# 细则见同目录 细则.txt（维度1 门禁 + 维度2 加分/减分项）。

SCRIPT_ID = "003"

# 维度二"结果追加顺序"的 (max_delta, kind) 列表；末端把 results 里的 "+N：..."/"-N：..."
# 反解为结构化 dim2_items 时用。顺序须与 _evaluate_impl 中 results.append(...) 严格一致。
#   kind='bonus'  ：加分项，命中得 +mag，否则 0；输出 max_delta = +mag
#   kind='penalty'：减分项，触发扣 mag，否则 0；输出 max_delta = -mag（与其他脚本一致，
#                   减分项 max_delta 取负值表示该项能造成的最大变动；不计入维度二满分）
_ITEM_SPECS = [
    (5, 'bonus'),    # 01 各行均已添加横线
    (3, 'bonus'),    # 02 横线样式统一（上下单实线 231F20 sz2，无竖线）
    (5, 'bonus'),    # 03 一行一格（单列、未合并/未拆分）——指令"每一行文字都放到表里"的核心动作
    (3, 'bonus'),    # 04 已删除空行
    (3, 'bonus'),    # 05 单元格正文字体字号对齐与原件一致、垂直居中
    (1, 'bonus'),    # 06 会议内容标签行保留且位于表格上方
    (1, 'bonus'),    # 07 标题"部门协调会"微软雅黑20.5磅居中
    (1, 'bonus'),    # 08 表头三行微软雅黑9.5磅
    (1, 'bonus'),    # 09 表格宽度=版心、单列、纸张页边距未变
    (5, 'penalty'),  # 10 正文有遗漏/被截断（原为 +5 加分项，见下方说明改为减分项）
    (3, 'penalty'),  # 11 改动了正文文字内容
    (3, 'penalty'),  # 12 新增多余表格/文本框
    (1, 'penalty'),  # 13 添加多余底纹/颜色/高亮/下划线
    (1, 'penalty'),  # 14 新增页眉/页脚/水印/批注/多余标题说明
]
# 维度二满分 = 所有加分项 max_delta 之和（减分项只减不增，不计入满分）
MAX_SCORE_TOTAL = sum(m for m, k in _ITEM_SPECS if k == 'bonus')

# ===== 参照原"部门协调会记录"版式的期望常量（细则中"未被改动"以此为基准）=====
STRICT_SONGTI = ('宋体', 'SimSun', 'NSimSun')
BORDER_COLOR = '231F20'          # 横线颜色（深灰黑）
BORDER_SZ = 2                    # 横线线宽（1/8 磅，=0.25磅）
TITLE_TEXT = '部门协调会'
YAHEI = '微软雅黑'
# "会议内容"正文的格式基线（细则中"字体字号与对齐方式未被改动"以此为准）。
# 原件正文本就不是统一格式：主体为宋体 10.5 磅（21 半磅），但"许承柏：请各位成员发表意见？"
# 若干行为微软雅黑 10.5 磅，议题 5/6/7 的标题行为微软雅黑 10 磅（20 半磅）。这些差异属指令
# "其他格式保持不变"应原样保留的内容，不得按"统一为宋体 10.5 磅"判罚。
# 原件中出现过的 (东亚字体关键字集, 半磅) 组合白名单：正文任一 run 必须落在其中。
BODY_ALLOWED_FONTS = (
    (STRICT_SONGTI, 21),        # 宋体 10.5 磅（正文主体）
    ((YAHEI,), 21),             # 微软雅黑 10.5 磅（"请各位成员发表意见？"等行）
    ((YAHEI,), 20),             # 微软雅黑 10 磅（议题 5/6/7 标题行）
)
# 原件中文本唯一、可精确定位的行：字体字号须与原件完全一致（被统一改成宋体 10.5 即不通过）。
# "请各位成员发表意见？"一行在原件中重复出现且本身宋体/微软雅黑混用，无法逐行定位，
# 故不列入此表，只受上面白名单约束。
BODY_PINNED_FONTS = {
    '5、北苑营业点生产用水及排水设施维护事项': (YAHEI, 20),
    '6、北苑货区检测设备维修': (YAHEI, 20),
    '7、北苑货区车辆雨棚整修': (YAHEI, 20),
}
# 原件正文绝大多数行并未设置 w:jc（走默认对齐），仅少数行为两端对齐，个别行为左对齐；
# 故"对齐方式未被改动"只判定是否被改成了明显可见的居中/右对齐/分散对齐，
# 默认对齐、左对齐、两端对齐三者均视为未被改动。
ALIGN_CHANGED = ('center', 'right', 'end', 'distribute')
# 期望页面（原文件 sectPr）：pgSz 11850x16783 twips，pgMar 上下400/左846/右847
EXPECT_PGW, EXPECT_PGH = 11850, 16783
EXPECT_MAR = {'top': 400, 'bottom': 400, 'left': 846, 'right': 847}
# 完整"会议内容"的分节锚点（用于判定是否完整置入、无遗漏）
CONTENT_ANCHORS = [
    '一、专题学习', '二、回顾2月份工作情况', '三、作业偏差情况分析', '四、研究后续事项',
    '1、研究东港作业点值班室照明', '2、研究南苑作业点茶水间排风', '3、研究北苑货区仓库安装高处作业防护挂点',
    '4、南苑营业点现场办公设施维护', '5、北苑营业点生产用水及排水设施维护', '6、北苑货区检测设备维修',
    '7、北苑货区车辆雨棚整修', '8、语禾集运服务有限公司申请2026年协同装卸', '最终意见', '散会',
]
# 关键事实锚点（会议记录不可被篡改的人名/日期/地点），用于"改动正文"减分项的粗判
FACT_ANCHORS = ['许承柏', '陈予宁', '侯景澜', '魏预', '周言朔', '林若棠',
                '2026年3月18日', '综合楼二层会议室']

def _default_para_style_id(styles_el):
    """styles.xml 中默认段落样式（w:type="paragraph" w:default="1"，通常即 Normal）的 styleId。
       段落未写 w:pStyle 时，Word 按该样式取值，故继承链须从这里起算。"""
    for st in styles_el.findall(qn('w:style')):
        if st.get(qn('w:type')) == 'paragraph' and st.get(qn('w:default')) == '1':
            return st.get(qn('w:styleId'))
    return None


def _resolve_effective_eastasia(doc, run):
    """按 Word 字体继承链解析 run 中文字符实际生效的东亚字体：
       run 直接设定(w:eastAsia) → 段落样式(含 basedOn 链) → docDefaults。都没有则 None。"""
    rPr = run._element.rPr
    if rPr is not None:
        rFonts = rPr.find(qn('w:rFonts'))
        if rFonts is not None:
            ea = rFonts.get(qn('w:eastAsia'))
            if ea:
                return ea
    styles_el = doc.styles.element
    para = run._parent
    style_id = None
    try:
        pPr = para._p.find(qn('w:pPr'))
        if pPr is not None:
            pStyle = pPr.find(qn('w:pStyle'))
            if pStyle is not None:
                style_id = pStyle.get(qn('w:val'))
    except Exception:
        style_id = None
    if not style_id:            # 段落未设 w:pStyle → 从默认段落样式(Normal)起算继承链
        style_id = _default_para_style_id(styles_el)

    def _style_by_id(sid):
        for st in styles_el.findall(qn('w:style')):
            if st.get(qn('w:styleId')) == sid:
                return st
        return None

    seen = set()
    while style_id and style_id not in seen:
        seen.add(style_id)
        st = _style_by_id(style_id)
        if st is None:
            break
        rpr = st.find(qn('w:rPr'))
        if rpr is not None:
            rf = rpr.find(qn('w:rFonts'))
            if rf is not None:
                ea = rf.get(qn('w:eastAsia'))
                if ea:
                    return ea
        base = st.find(qn('w:basedOn'))
        style_id = base.get(qn('w:val')) if base is not None else None

    dd = styles_el.find(qn('w:docDefaults'))
    if dd is not None:
        rpd = dd.find(qn('w:rPrDefault'))
        if rpd is not None:
            rpr = rpd.find(qn('w:rPr'))
            if rpr is not None:
                rf = rpr.find(qn('w:rFonts'))
                if rf is not None:
                    ea = rf.get(qn('w:eastAsia'))
                    if ea:
                        return ea
    return None


def _resolve_effective_size_halfpt(doc, run):
    """按 Word 继承链解析 run 实际生效字号（半磅）：run → 段落样式(basedOn 链) → docDefaults。"""
    def _sz_of(rpr):
        if rpr is None:
            return None
        sz = rpr.find(qn('w:sz'))
        if sz is None:
            return None
        v = sz.get(qn('w:val'))
        return int(v) if v and v.lstrip('-').isdigit() else None

    v = _sz_of(run._element.rPr)
    if v is not None:
        return v
    styles_el = doc.styles.element
    para = run._parent
    style_id = None
    try:
        pPr = para._p.find(qn('w:pPr'))
        if pPr is not None:
            pStyle = pPr.find(qn('w:pStyle'))
            if pStyle is not None:
                style_id = pStyle.get(qn('w:val'))
    except Exception:
        style_id = None
    if not style_id:            # 段落未设 w:pStyle → 从默认段落样式(Normal)起算继承链
        style_id = _default_para_style_id(styles_el)

    def _style_by_id(sid):
        for st in styles_el.findall(qn('w:style')):
            if st.get(qn('w:styleId')) == sid:
                return st
        return None

    seen = set()
    while style_id and style_id not in seen:
        seen.add(style_id)
        st = _style_by_id(style_id)
        if st is None:
            break
        v = _sz_of(st.find(qn('w:rPr')))
        if v is not None:
            return v
        base = st.find(qn('w:basedOn'))
        style_id = base.get(qn('w:val')) if base is not None else None
    dd = styles_el.find(qn('w:docDefaults'))
    if dd is not None:
        rpd = dd.find(qn('w:rPrDefault'))
        if rpd is not None:
            v = _sz_of(rpd.find(qn('w:rPr')))
            if v is not None:
                return v
    return None


def _norm(s):
    """去除所有空白（含全角空格/NBSP）后的文本，用于稳健的子串/相等比较。"""
    return re.sub(r'\s+', '', (s or '')).replace('　', '').replace('\xa0', '')


def _elem_text(elem):
    """OOXML 元素内所有 w:t 文本拼接。"""
    return ''.join(t.text or '' for t in elem.iter(qn('w:t')))


def _cell_own_text(tc):
    """单元格自身文字：只取其直属 w:p，不含嵌套表格内的文字。"""
    if tc is None:
        return ''
    return ''.join(_elem_text(p) for p in tc.findall(qn('w:p')))


def _flatten_content_rows(tbl_elem):
    """按文档顺序展平表格的"内容行"，返回 [(tr, 首单元格 tc), ...]。
       单元格内若嵌有表格（原文件"一、专题学习"格内即嵌了一张 56 行的表），
       其各行也是"会议内容"的显示行，须一并纳入逐行判定，否则会被整块漏掉。"""
    out = []
    for tr in tbl_elem.findall(qn('w:tr')):
        cells = tr.findall(qn('w:tc'))
        out.append((tr, cells[0] if cells else None))
        for tc in cells:
            for nested in tc.findall(qn('w:tbl')):
                out.extend(_flatten_content_rows(nested))
    return out


def _find_content_table_elem(doc):
    """定位"会议内容"标签段之后的第一个表格元素(w:tbl)；找不到则回退为行数最多的表格。"""
    body = doc.element.body
    label_seen = False
    for child in body.iterchildren():
        tag = child.tag.split('}')[-1]
        if tag == 'p':
            if '会议内容' in _norm(_elem_text(child)):
                label_seen = True
        elif tag == 'tbl':
            if label_seen:
                return child
    # 回退：行数最多的表格
    best, best_rows = None, -1
    for tbl in body.findall(qn('w:tbl')):
        n = len(tbl.findall(qn('w:tr')))
        if n > best_rows:
            best, best_rows = tbl, n
    return best


def _border_edge(tcPr, name):
    """返回单元格某条边框的 (val, color, sz)；无则 (None,None,None)。"""
    if tcPr is None:
        return (None, None, None)
    tb = tcPr.find(qn('w:tcBorders'))
    if tb is None:
        return (None, None, None)
    e = tb.find(qn('w:' + name))
    if e is None:
        return (None, None, None)
    return (e.get(qn('w:val')), e.get(qn('w:color')), e.get(qn('w:sz')))


def _has_line(val):
    return val is not None and val not in ('none', 'nil')


def _char_em(ch):
    """单个字符的显示宽度（相对字号 em）：CJK 及全角标点/全角字符按 1 em，
    其余（ASCII 字母数字、半角标点、空格）按 0.5 em。"""
    o = ord(ch)
    if (0x4E00 <= o <= 0x9FFF or 0x3400 <= o <= 0x4DBF or 0xF900 <= o <= 0xFAFF
            or 0x3000 <= o <= 0x303F         # CJK 标点（，。、；：「」等）
            or 0xFF00 <= o <= 0xFFEF          # 全角 ASCII / 全角标点
            or 0x2018 <= o <= 0x201F          # 中文引号
            or 0x2E80 <= o <= 0x2EFF):        # CJK 部首补充
        return 1.0
    return 0.5


def _row_is_merged(text, usable_twips, half_pt, tol=1.15):
    """静态估算某行单元格文字是否明显超过"一个显示行"的宽度——用于识别把多行
    正文合并进同一格（导致中间缺横线）的情况。.docx 不存储换行位置，故据字号与
    字宽近似求文本总宽，超过版心可用宽的 tol 倍则判为"合并了多行"。"""
    if not text or usable_twips <= 0:
        return False
    em = (half_pt / 2.0) * 20.0              # 1 em 宽度（twips）≈ 字号磅数 × 20
    if em <= 0:
        return False
    width = sum(_char_em(c) for c in text) * em
    return width > usable_twips * tol


class _ParaLite:
    """给 _resolve_effective_* 用的极简段落包装：只暴露 _p（段落底层元素）。"""
    def __init__(self, doc, p_elem):
        self._p = p_elem


class _RunLite:
    """给 _resolve_effective_* 用的极简 run 包装：_element 为 w:r（CT_R，含 .rPr），
    _parent 为 _ParaLite。python-docx 解析出的元素均为注册的自定义类，故 .rPr 可用。"""
    def __init__(self, r_elem, para):
        self._element = r_elem
        self._parent = para



def _locate_error(result, message):
    """未能在目录中定位到被评估文档（目录不存在 / 目录中没有 .docx）：维度一根本没有进入判定，
    故 dim1_pass 置为空字符串表示"未判定"，dim1_reason 留空，失败摘要只写在 error；
    维度二一项未评，max_score 同时归 0。"""
    result["status"] = "error"
    result["error"] = message
    result["dim1_pass"] = ""
    result["dim1_reason"] = ""
    result["dim2_items"] = []
    result["total_score"] = 0
    result["max_score"] = 0
    return result


def evaluate(dir_path):
    """按统一约定返回结构化 dict；不 print、不改 sys.stdout、不 sys.exit。
    入参 dir_path 为目录路径，脚本自己在其中定位并打开被评估的 .docx（忽略 ~$ 临时锁文件）。
    脚本自身崩溃（目录不存在 / 未发现文档 / 文件损坏无法打开等）由本函数捕获并返回 status='error'。

    输出约定：
      ① 定位不到被评估文档（目录不存在 / 无 .docx）→ status='error'，error 写失败摘要，
         dim1_pass=""（未判定）、dim1_reason=""、dim2_items=[]、total_score=0、max_score=0；
      ② 文档存在但打不开等导致维度一不通过 → status='error'，dim1_pass=False，
         此时 dim1_reason 须写出错误摘要（与 error 同文）。"""
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
            return _locate_error(result, f"目录不存在：{dir_path}")
        cands = sorted(
            f for f in os.listdir(dir_path)
            if f.lower().endswith('.docx') and not f.startswith('~$')
        )
        if not cands:
            return _locate_error(result, f"目录中未找到 .docx 文件：{dir_path}")
        file_path = os.path.join(dir_path, cands[0])
        result["file_name"] = os.path.basename(file_path)
        return _evaluate_impl(file_path, result)
    except Exception as exc:
        # 文件损坏/无法打开等 → 维度一不通过，完成度记 0；错误摘要同时写入 error 与 dim1_reason。
        summary = f"{type(exc).__name__}: {exc}"
        result["status"] = "error"
        result["error"] = summary
        result["dim1_pass"] = False
        result["dim1_reason"] = summary
        result["dim2_items"] = []
        result["total_score"] = 0
        return result


def _finalize(result, results):
    """把 results 里的 "+N：..."/"-N：..." 按 _ITEM_SPECS 顺序反解为结构化 dim2_items。"""
    if len(results) != len(_ITEM_SPECS):
        raise RuntimeError(
            f"dim2 项数不一致：results={len(results)} vs specs={len(_ITEM_SPECS)}"
        )
    dim2_items = []
    total = 0
    for raw, (mag, kind) in zip(results, _ITEM_SPECS):
        head, _, rule = raw.partition("：")
        num = int(re.sub(r'[^0-9]', '', head) or 0)
        if kind == 'bonus':
            hit = num > 0
            delta = mag if hit else 0
            max_delta = mag
        else:  # penalty
            hit = num > 0                      # 触发扣分即视为"命中该减分项"
            delta = -mag if hit else 0
            max_delta = -mag                   # 该减分项能造成的最大变动（负值）
        dim2_items.append({
            "rule": rule,
            "max_delta": max_delta,
            "delta": delta,
            "hit": hit,
        })
        total += delta
    result["dim2_items"] = dim2_items
    result["total_score"] = max(0, total)      # 完成度不低于 0
    result["max_score"] = MAX_SCORE_TOTAL
    return result


def _evaluate_impl(file_path, result):
    doc = Document(file_path)
    results = []

    # ========== 维度1：可用与可修改性（门禁） ==========
    # 细则维度1 仅要求"交付文件应为1个可正常打开的Word文档（.docx）"：
    # 能被 Document() 成功打开即通过；打不开的情形已在 evaluate() 中捕获为不通过。
    result["dim1_pass"] = True
    result["dim1_reason"] = ""          # 通过时不填理由；不通过的情形只发生在 evaluate 的
                                        # 异常分支（status='error'），那里错误摘要同时写入
                                        # error 与 dim1_reason

    tbl_elem = _find_content_table_elem(doc)
    if tbl_elem is None:
        # 找不到可评的"会议内容"表格不再作为门禁项，维度2 各加分项均不得分。
        result["dim2_items"] = []
        result["total_score"] = 0
        return result
    rows = tbl_elem.findall(qn('w:tr'))
    table_text = _elem_text(tbl_elem)

    # 预取每行首单元格（本任务表格为单列一行一格）。嵌套表格的行一并展平纳入，
    # 否则嵌在单元格里的整块内容会绕过下面所有逐行判定。
    row_pairs = _flatten_content_rows(tbl_elem)
    rows = [tr for tr, _ in row_pairs]
    row_cells = [tr.findall(qn('w:tc')) for tr in rows]
    row_first_tc = [tc for _, tc in row_pairs]
    # 行文字只取单元格自身段落：宿主行不应把嵌套表格的文字算作自己的一行
    row_texts = [_norm(_cell_own_text(tc)) for tc in row_first_tc]
    full_table_text = _norm(table_text)
    full_doc_text = _norm('\n'.join(p.text for p in doc.paragraphs) + table_text)

    # 单元格可用文字宽（twips）：优先取实际表宽，否则用版心宽；再扣默认单元格左右内边距。
    content_w_twips = EXPECT_PGW - EXPECT_MAR['left'] - EXPECT_MAR['right']
    _tblPr0 = tbl_elem.find(qn('w:tblPr'))
    if _tblPr0 is not None:
        _wEl0 = _tblPr0.find(qn('w:tblW'))
        if _wEl0 is not None and (_wEl0.get(qn('w:w')) or '').isdigit():
            content_w_twips = int(_wEl0.get(qn('w:w')))
    usable_twips = max(1, content_w_twips - 216)   # 默认单元格左右内边距各约 108 twips

    def _cell_rep_halfpt(tc):
        """取单元格首个有文字 run 的生效字号（半磅）；取不到默认 21（=10.5磅）。"""
        if tc is None:
            return 21
        for p in tc.findall(qn('w:p')):
            para = _ParaLite(doc, p)
            for r_el in p.findall(qn('w:r')):
                run = _RunLite(r_el, para)
                t = ''.join(x.text or '' for x in r_el.iter(qn('w:t')))
                if t.strip():
                    h = _resolve_effective_size_halfpt(doc, run)
                    return h if h else 21
        return 21

    # 逐行判定是否"把多行合并进了同一格"（其内部换行处必然缺横线）
    row_merged = [_row_is_merged(t, usable_twips, _cell_rep_halfpt(tc))
                  for tc, t in zip(row_first_tc, row_texts)]

    # ========== 维度2：完成度评分 ==========

    # +5：表格各行均已添加横线，全表逐行覆盖，无遗漏或缺线行
    #     注：把多行正文合并进同一格的行，其内部换行处并无横线 → 视为"缺线行"，不计入。
    lined_rows = 0
    for tc, merged in zip(row_first_tc, row_merged):
        tcPr = tc.find(qn('w:tcPr')) if tc is not None else None
        tv, _, _ = _border_edge(tcPr, 'top')
        bv, _, _ = _border_edge(tcPr, 'bottom')
        if (_has_line(tv) or _has_line(bv)) and not merged:
            lined_rows += 1
    if rows and lined_rows == len(rows):
        results.append("+5：表格各行均已添加横线，全表逐行覆盖，无遗漏或缺线行")
    else:
        results.append("+0：表格各行均已添加横线，全表逐行覆盖，无遗漏或缺线行")

    # +3：横线样式为上、下两条水平细实线（single、#231F20、sz=2），全表统一，且不含左右竖线
    style_ok = bool(rows)
    for tc in row_first_tc:
        tcPr = tc.find(qn('w:tcPr')) if tc is not None else None
        tv, tc_, tsz = _border_edge(tcPr, 'top')
        bv, bc, bsz = _border_edge(tcPr, 'bottom')
        lv, _, _ = _border_edge(tcPr, 'left')
        rv, _, _ = _border_edge(tcPr, 'right')

        def _ok(v, c, s):
            return (v == 'single' and (c or '').upper() == BORDER_COLOR
                    and s is not None and s.isdigit() and int(s) == BORDER_SZ)
        if not (_ok(tv, tc_, tsz) and _ok(bv, bc, bsz)):
            style_ok = False
            break
        if _has_line(lv) or _has_line(rv):
            style_ok = False
            break
    if style_ok:
        results.append("+3：横线样式为上下两条水平细实线（单线、#231F20、sz=2），全表统一且不含左右竖线")
    else:
        results.append("+0：横线样式为上下两条水平细实线（单线、#231F20、sz=2），全表统一且不含左右竖线")

    # +5：每一显示行的文字对应表格中的一行，一行一格（单列；未合并多行、未拆分一行）
    #     权重与"逐行加横线"并列为 5 分：这是指令"每一行文字都放到表里面"的核心动作，
    #     也是最易做错的一项（把整段塞进一格、或把一行切到多格）。
    one_line_ok = bool(rows)
    for cs, tc, merged in zip(row_cells, row_first_tc, row_merged):
        if len(cs) != 1:                       # 单元格>1 → 一行被拆到多格
            one_line_ok = False
            break
        # 单个单元格内若含多个非空段落 → 多行被合并进同一格
        nonempty_ps = [p for p in tc.findall(qn('w:p')) if _norm(_elem_text(p))]
        if len(nonempty_ps) > 1:
            one_line_ok = False
            break
        if merged:                             # 单段落但文字超一显示行 → 多行被合并进同一格
            one_line_ok = False
            break
    if one_line_ok:
        results.append("+5：会议内容每一显示行对应表格一行，一行一格，未合并多行亦未拆分一行")
    else:
        results.append("+0：会议内容每一显示行对应表格一行，一行一格，未合并多行亦未拆分一行")

    # +3：已删除"会议内容"区域内所有无文字的空行（此处判定表格内无空行）
    empty_rows = sum(1 for t in row_texts if t == '')
    if empty_rows == 0:
        results.append("+3：已删除会议内容区域内所有无文字的空行，删除后无残留空白行")
    else:
        results.append("+0：已删除会议内容区域内所有无文字的空行，删除后无残留空白行")

    # +3：单元格正文各行的字体、字号与对齐方式与原文件一致（未被改动），且垂直居中
    #     原件正文本非统一格式（见 BODY_ALLOWED_FONTS / BODY_PINNED_FONTS 注释），故判定为：
    #     ① 每个 run 的 (字体, 字号) 必须是原件出现过的组合；② 原件中文本唯一的议题 5/6/7
    #     标题行须仍为微软雅黑 10 磅；③ 对齐未被改成居中/右对齐/分散对齐；④ 单元格垂直居中。
    font_ok = bool(rows)
    for tc in row_first_tc:
        if tc is None:
            font_ok = False
            break
        tcPr = tc.find(qn('w:tcPr'))
        vAlign = tcPr.find(qn('w:vAlign')) if tcPr is not None else None
        if vAlign is None or vAlign.get(qn('w:val')) != 'center':   # 垂直居中
            font_ok = False
            break
        cell_ok = True
        for p in tc.findall(qn('w:p')):
            line = _norm(_elem_text(p))
            if not line:
                continue
            pPr = p.find(qn('w:pPr'))
            jc = pPr.find(qn('w:jc')) if pPr is not None else None
            jc_val = jc.get(qn('w:val')) if jc is not None else None
            if jc_val in ALIGN_CHANGED:          # 对齐被改成居中/右对齐/分散对齐
                cell_ok = False
                break
            pinned = BODY_PINNED_FONTS.get(line)
            para = _ParaLite(doc, p)
            for r_el in p.findall(qn('w:r')):
                run = _RunLite(r_el, para)
                txt = ''.join(t.text or '' for t in r_el.iter(qn('w:t')))
                if not txt.strip():
                    continue
                half = _resolve_effective_size_halfpt(doc, run)
                ea = _resolve_effective_eastasia(doc, run) or ''
                is_cjk = bool(re.search(r'[一-鿿]', txt))
                if pinned is not None:                  # 唯一可定位行 → 与原件精确比对
                    if half != pinned[1] or (is_cjk and pinned[0] not in ea):
                        cell_ok = False
                        break
                    continue
                if half is None:
                    cell_ok = False
                    break
                if is_cjk:
                    if not any(half == h and any(k in ea for k in ks)
                               for ks, h in BODY_ALLOWED_FONTS):
                        cell_ok = False
                        break
                elif not any(half == h for _, h in BODY_ALLOWED_FONTS):
                    cell_ok = False
                    break
            if not cell_ok:
                break
        if not cell_ok:
            font_ok = False
            break
    if font_ok:
        results.append("+3：单元格正文各行字体、字号与对齐方式与原文件一致未被改动，且垂直居中")
    else:
        results.append("+0：单元格正文各行字体、字号与对齐方式与原文件一致未被改动，且垂直居中")

    # --- 表格之前的 body 顶层段落（标题/表头/标签所在处）---
    lead_paras = []
    for child in doc.element.body.iterchildren():
        if child is tbl_elem:
            break
        if child.tag.split('}')[-1] == 'p':
            lead_paras.append(child)

    def _para_runs(p_elem):
        para = _ParaLite(doc, p_elem)
        return [(_RunLite(r, para), ''.join(t.text or '' for t in r.iter(qn('w:t'))))
                for r in p_elem.findall(qn('w:r'))]

    def _jc_of(p_elem):
        pPr = p_elem.find(qn('w:pPr'))
        jc = pPr.find(qn('w:jc')) if pPr is not None else None
        return jc.get(qn('w:val')) if jc is not None else None

    # +1："会议内容 :"标签行予以保留，且位于表格上方
    label_ok = any('会议内容' in _norm(_elem_text(p)) for p in lead_paras)
    if label_ok:
        results.append("+1：\"会议内容 :\"标签行予以保留，且位于表格上方")
    else:
        results.append("+0：\"会议内容 :\"标签行予以保留，且位于表格上方")

    # +1：标题"部门协调会"为微软雅黑、20.5磅、居中，文字与格式均未被改动
    title_ok = False
    for p in lead_paras:
        if _norm(_elem_text(p)) != TITLE_TEXT:
            continue
        if _jc_of(p) != 'center':
            break
        ok = True
        for run, txt in _para_runs(p):
            if not txt.strip():
                continue
            ea = _resolve_effective_eastasia(doc, run)
            half = _resolve_effective_size_halfpt(doc, run)
            if not (ea and YAHEI in ea):
                ok = False
                break
            if half is None or not (40 <= half <= 42):   # 20.5磅=41半磅（容差40-42）
                ok = False
                break
        title_ok = ok
        break
    if title_ok:
        results.append("+1：标题\"部门协调会\"为微软雅黑、20.5磅、居中，文字与格式均未被改动")
    else:
        results.append("+0：标题\"部门协调会\"为微软雅黑、20.5磅、居中，文字与格式均未被改动")

    # +1：表头三行（时间/主持人/记录人、地点/应出席数/实出席数、列席人员）为微软雅黑、9.5磅
    header_keys = ('主持人', '应出席数', '列席人员')
    header_paras = [p for p in lead_paras
                    if any(k in _norm(_elem_text(p)) for k in header_keys)]
    header_ok = len(header_paras) >= 3
    for p in header_paras:
        for run, txt in _para_runs(p):
            if not txt.strip():
                continue
            ea = _resolve_effective_eastasia(doc, run)
            half = _resolve_effective_size_halfpt(doc, run)
            if not (ea and YAHEI in ea):
                header_ok = False
                break
            if half is None or not (18 <= half <= 20):   # 9.5磅=19半磅（容差18-20）
                header_ok = False
                break
        if not header_ok:
            break
    if header_ok:
        results.append("+1：表头三行（时间/主持人/记录人、地点/应出席数/实出席数、列席人员）为微软雅黑、9.5磅")
    else:
        results.append("+0：表头三行（时间/主持人/记录人、地点/应出席数/实出席数、列席人员）为微软雅黑、9.5磅")

    # +1：表格宽度与正文版心一致并保持单列，纸张大小与页边距未发生改变
    sect = doc.element.body.find(qn('w:sectPr'))
    page_ok = sect is not None
    if page_ok:
        pgSz = sect.find(qn('w:pgSz'))
        pgMar = sect.find(qn('w:pgMar'))
        cols = sect.find(qn('w:cols'))
        w = int(pgSz.get(qn('w:w'))) if pgSz is not None and pgSz.get(qn('w:w')) else 0
        h = int(pgSz.get(qn('w:h'))) if pgSz is not None and pgSz.get(qn('w:h')) else 0
        if abs(w - EXPECT_PGW) > 40 or abs(h - EXPECT_PGH) > 40:
            page_ok = False
        if pgMar is not None:
            for k, exp in EXPECT_MAR.items():
                v = pgMar.get(qn('w:' + k))
                if v is None or abs(int(v) - exp) > 30:
                    page_ok = False
                    break
        else:
            page_ok = False
        num = cols.get(qn('w:num')) if cols is not None else None
        if num is not None and num != '1':          # 多栏 → 不合规
            page_ok = False
    # 单列：所有行仅 1 格
    single_col = all(len(cs) == 1 for cs in row_cells) and bool(row_cells)
    # 表格宽度 ≈ 版心宽（页宽-左右边距）
    tblW = None
    tblPr = tbl_elem.find(qn('w:tblPr'))
    if tblPr is not None:
        wEl = tblPr.find(qn('w:tblW'))
        if wEl is not None and wEl.get(qn('w:w')) and wEl.get(qn('w:w')).isdigit():
            tblW = int(wEl.get(qn('w:w')))
    content_w = EXPECT_PGW - EXPECT_MAR['left'] - EXPECT_MAR['right']
    width_ok = tblW is not None and (0.85 * content_w) <= tblW <= (1.03 * content_w)
    if page_ok and single_col and width_ok:
        results.append("+1：表格宽度与正文版心一致并保持单列，纸张大小与页边距未发生改变")
    else:
        results.append("+0：表格宽度与正文版心一致并保持单列，纸张大小与页边距未发生改变")

    # ========== 减分项 ==========
    # 说明：减分项在仅有输出文件、无原始文件可比对时，采用可静态检出的保守代理判定，
    #       宁可漏报（默认不扣分）也不误伤；命中确凿的多余修改信号才触发扣分。

    # -5：未把"会议内容"下方的全部正文完整置入表格，存在文字遗漏或整段截断
    #     本项原为 +5 加分项，现改为减分项：原件正文本就已在表格内（部分嵌在单元格内的
    #     嵌套表格里），"完整置入"对未处理的原件同样成立，作为加分项对所有交付无差别送分、
    #     无区分度且抬高分数基线；改为减分项后只在真的丢内容时触发（如把表格逐行重建到
    #     一半就截断，丢掉议题6/7/8 与"散会"）。注意 -3"改动正文文字内容"抓不到这种情况：
    #     它的代理判据 FACT_ANCHORS 是人名/日期/地点，截断后这些锚点往往仍然齐全。
    hit_anchors = [a for a in CONTENT_ANCHORS if _norm(a) in full_table_text]
    if len(hit_anchors) != len(CONTENT_ANCHORS):
        results.append("-5：未将\"会议内容\"下方的全部正文完整置入表格，存在文字遗漏或整段截断")
    else:
        results.append("-0：未将\"会议内容\"下方的全部正文完整置入表格，存在文字遗漏或整段截断")

    # -3：改动了"会议内容"正文的文字内容（篡改数字/日期/人名/议题编号等）
    #     代理判定：正文已实际写入（维度一通过）却缺失应恒在的关键事实锚点 → 视为被篡改/破坏。
    missing_facts = [a for a in FACT_ANCHORS if _norm(a) not in full_doc_text]
    if missing_facts:
        results.append("-3：改动了会议内容正文的文字内容（关键人名/日期/地点等事实缺失或不符）")
    else:
        results.append("-0：改动了会议内容正文的文字内容（关键人名/日期/地点等事实缺失或不符）")

    # -3：新增指令未要求的表格、文本框等多余结构
    extra_tbl = len(doc.element.body.findall(qn('w:tbl'))) > 1
    has_txbx = doc.element.body.find('.//' + qn('w:txbxContent')) is not None
    if extra_tbl or has_txbx:
        results.append("-3：新增指令未要求的表格、文本框等多余结构，扰乱版面")
    else:
        results.append("-0：新增指令未要求的表格、文本框等多余结构，扰乱版面")

    # -1：为文字添加指令未要求的底纹、颜色、高亮或下划线等多余修饰（限"会议内容"单元格内）
    #     只看承载可见文字的 run：段落标记(¶)自带的 w:pPr/w:rPr 与空 run 不影响文字外观，
    #     而原文件最末段的段落标记本就带 w:u，若一并计入会对未作任何修饰的交付误扣分。
    deco = False
    for tc in row_first_tc:
        if tc is None:
            continue
        for p in tc.findall(qn('w:p')):
            for r_el in p.findall(qn('w:r')):
                if not ''.join(t.text or '' for t in r_el.iter(qn('w:t'))).strip():
                    continue
                rPr = r_el.find(qn('w:rPr'))
                if rPr is None:
                    continue
                if rPr.find(qn('w:highlight')) is not None:
                    deco = True
                    break
                u = rPr.find(qn('w:u'))
                if u is not None and (u.get(qn('w:val')) or 'single') not in ('none',):
                    deco = True
                    break
                shd = rPr.find(qn('w:shd'))
                if shd is not None:
                    fill = (shd.get(qn('w:fill')) or '').upper()
                    if fill not in ('', 'AUTO', 'FFFFFF') and \
                            (shd.get(qn('w:val')) or 'clear') != 'nil':
                        deco = True
                        break
            if deco:
                break
        if deco:
            break
    if deco:
        results.append("-1：为会议内容文字添加了指令未要求的底纹、颜色、高亮或下划线等多余修饰")
    else:
        results.append("-0：为会议内容文字添加了指令未要求的底纹、颜色、高亮或下划线等多余修饰")

    # -1：新增指令未要求的页眉、页脚、水印、批注或多余标题、说明文字
    #     代理判定：出现批注引用，或表格之后新增了含文字的多余段落。
    doc_xml = doc.element.body.xml if hasattr(doc.element.body, 'xml') else ''
    has_comment = 'commentReference' in doc_xml
    trailing_text = False
    seen_tbl = False
    for child in doc.element.body.iterchildren():
        if child is tbl_elem:
            seen_tbl = True
            continue
        if seen_tbl and child.tag.split('}')[-1] == 'p' and _norm(_elem_text(child)):
            trailing_text = True
            break
    if has_comment or trailing_text:
        results.append("-1：新增指令未要求的批注或表格之后的多余说明文字")
    else:
        results.append("-0：新增指令未要求的批注或表格之后的多余说明文字")


    return _finalize(result, results)


if __name__ == '__main__':
    # 本地调试入口：默认取脚本所在目录，或用命令行传入的目录路径。
    _dir = sys.argv[1] if len(sys.argv) > 1 else os.path.dirname(os.path.abspath(__file__))
    print(json.dumps(evaluate(_dir), ensure_ascii=False, indent=2))



