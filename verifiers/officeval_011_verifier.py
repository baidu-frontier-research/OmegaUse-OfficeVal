#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# pyright: reportUnknownMemberType=false, reportUnknownVariableType=false, reportUnknownArgumentType=false, reportUnknownParameterType=false, reportMissingParameterType=false, reportMissingTypeArgument=false, reportPrivateUsage=false, reportUnknownLambdaType=false, reportGeneralTypeIssues=false
"""
教案（BOPPPS 教学设计）格式自动评估工具（officeval_011）

统一入口：
    evaluate(dir_path: str) -> dict
    - dir_path: 脚本所在目录（内含待评估的 .docx）
    - 返回字段：id / file_name / status / error / dim1_pass / dim1_reason /
                dim2_items（命中与未命中项都返回）/ total_score / max_score
    - dim2_items 每项为 {rule, max_delta, delta, hit}，delta 严格二值

规范约束：
    * 只暴露 evaluate 一个函数供批量运行器调用
    * 不修改 sys.stdout、不打印主结果、不使用 sys.exit
    * 不硬编码文档路径，全部由 dir_path 参数驱动
    * status='error' 表示这份提交没被真正评过（目录不存在、目录内无可评估文档、
      脚本自身异常），此时 dim1_pass 为空串（未判定）、dim1_reason 为空串、
      dim2_items 为空表、total_score 与 max_score 一律为 0；
      文档能定位到但不达标一律 status='ok'，由 dim1_pass / dim2_items 承载
    * dim1_pass 为 False 时，dim1_reason 必定带有维度一未通过的错误摘要；
      dim1_pass 为 True 或空串时，dim1_reason 为空串
    * 本文件仅在 __main__ 下作为本地调试入口，才把 dict 序列化为 JSON 输出
"""

import os
import re
import sys
import json
from docx import Document
from docx.oxml.ns import qn

SCRIPT_ID = "011"

# 维度二评分项目录：(max_delta, 规则原文)，顺序与《评测细则》正文逐行一致。
# 本细则含 1 个加分项与 29 个减分项；max_score 只累计正向项。
_DIM2_RULES = [
    # 01
    (5, """第3页"学情分析""教学目标""教学重点""教学难点""思政点""教学创新点"六个标签属于同一个表格（下称"学情分析…教学创新点"表格），"素质目标"内容格所在行与"教学重点"行紧密相接，中间无表格分界线、无空段落、无空白间隙"""),
    # 02
    (-3, """"学情分析…教学创新点"表格不为9行，或9行自上而下的次序不为"知识基础"行、"能力水平"行、"行为特征"行、"知识目标+能力目标+素质目标"三个标签并列的一行、三个目标的内容行（"1.掌握场地解读图的核心作用…""1.能借助数字工具完成柳岸口袋公园更新单元的资料收集…""1.树立'生态优先、以人为本'的设计理念…"）、"教学重点"行、"教学难点"行、"思政点"行、"教学创新点"行，或出现缺行、重复行、空白行"""),
    # 03
    (-3, """"学情分析…教学创新点"表格下半部分文字不完整："教学重点"格未含"依托层级逻辑定格局""依托要素逻辑改基底""依托问题逻辑补短板""依托潜力逻辑塑特色""联动智能制图落地"5条；"教学难点"格未含"复杂场地资料的分类整合与逻辑建构""规避工具依赖，兼顾专业判断与落地可行性"2条；"思政点"格未含"1.生态文明""2.职业素养""3.协作担当""4.创新教育"4条；"教学创新点"格未含"（一）教学模式创新：BOPPPS+数字工具+虚构项目三维融合""（二）教学工具创新：智能检索+图像生成+空间资料处理的协同工作流""（三）教学内容创新：'理论—实操—应用'全链条贯通"三个小标题及各自的说明段；或以上各条出现截断、改写、重复"""),
    # 04
    (-3, """"学情分析…教学创新点"表格9行之间出现分页符或"段前分页"设置"""),
    # 05
    (-1, """"学情分析…教学创新点"表格的列宽不符：第1列各行不统一为约1.34cm，或第2列各行不统一为约2.28cm，或"知识目标""能力目标""素质目标"三格不为约4.22cm、4.25cm、5.10cm"""),
    # 06
    (-1, """"学情分析…教学创新点"表格"教学重点""教学难点""思政点"三行的内容未各为1个整格、宽约11.3cm，被竖线切成多格，或"教学创新点"行不为2格（标签格约1.34cm、内容格约13.6cm）"""),
    # 07
    (-3, """"课程名称…教学内容"表格中"主要集中于场地解读图在韧性街区绿地认知与设计中的应用…"所在单元格内，该段文字下方未留出约2行高的空白，或该段的段前间距不为5.85磅"""),
    # 08
    (-3, """"教学环节"表格"任务层""任务内容"及"L1（基础层）""L2（进阶层）""L3（挑战层）"与其任务描述各格的行距不为1.27倍，或段前间距不为13.05磅，或"任务层""任务内容"两格文字不为居中对齐"""),
    # 09
    (-3, """"教学环节"表格"评价标准"列4格（表头"评价标准"及"资料完整度（40%）、图面清晰度（30%）、操作规范性（30%）""问题定位准确性（30%）、逻辑梳理清晰度（40%）、成果美观度（30%）""策略可实施性（40%）、逻辑转换流畅度（30%）、创新点（30%）"）的行距不为固定值18磅，或表头格首行缩进不为6字符、其余3格首行缩进不为2字符，或表头格不为两端对齐、其余3格不为左对齐"""),
    # 10
    (-3, """"学情分析…教学创新点"表格"行为特征"行的内容单元格内无约11.25cm×5.04cm的插图，或该图被移出单元格、被拉伸变形"""),
    # 11
    (-3, """"学情分析…教学创新点"表格底纹配色不为：第1列标签浅绿DDEFE6、第2列标签浅绿F0F7F2、"知识基础"与"教学重点"内容格浅紫E7E1F2、其余内容格米白FBFAF4"""),
    # 12
    (-1, """"学情分析…教学创新点"表格"教学重点""教学难点"两个标签格不为垂直居中，或该表格其余单元格不为垂直顶端对齐"""),
    # 13
    (-1, """文档表格总数不为5个（"课程名称…教学内容"表格、"学情分析…教学创新点"表格、"教学启示…信息化手段及教学资源"表、"教学环节"表、"反思诊改"表）"""),
    # 14
    (-1, """"学情分析…教学创新点"表格与上方"课程名称…教学内容"表格之间、与下方"教学启示"表格之间，各自间隔的空段落不是1个"""),
    # 15
    (-3, """文档页面尺寸不为A4（约21.0cm×29.7cm）"""),
    # 16
    (-3, """正文各节页边距不为上下2.54cm、左右3.17cm，或末节页边距不为上2.79cm、下2.08cm、左2.29cm、右1.87cm"""),
    # 17
    (-3, """文档页脚少于6个，或页脚中的页码不为自动更新的页码域（被替换为静态数字或文字）"""),
    # 18
    (-3, """封面插图缺失，或其尺寸不为约21.4cm×18.2cm"""),
    # 19
    (-3, """"信息化手段及教学资源"栏中"专业为本，工具赋能，整合多元信息化手段服务教学。"与"教学资源"两段之间的单元格内无约13.5cm×4.2cm的插图"""),
    # 20
    (-3, """"板书设计"标题下方无约6.1cm×9.7cm的插图"""),
    # 21
    (-3, """除"学情分析…教学创新点"表格外，其余表格的行列数不符："课程名称…教学内容"表格应为6行4列、"教学启示…信息化手段及教学资源"表应为5行2列、"教学环节"表应为15行6列、"反思诊改"表应为2行3列"""),
    # 22
    (-3, """"课程名称：城市绿地系统规划""授课内容：知识单元四／雨洪花园更新与微地形解析／第1节 场地解读图的作用及组织逻辑""授课对象：环境设计23-1/2/3""教学时长：1学时""推荐教材：[1] 周澜员. 城市绿地系统规划方法（第2版）. 澄景教育出版社,2022.""参考教材：[1] 邵闻远. 数字化设计基础…[2] 林若禾. 公园城市空间营造…""教育教学理念：以成果导向和项目化学习为主线…""教学内容：主要集中于场地解读图在韧性街区绿地认知与设计中的应用…"八个栏目及其内容任一缺失、文字不符或次序被调换"""),
    # 23
    (-3, """"教学环节"表格15行不为自上而下依次：表头行（教学环节／教学内容／教师活动／学生活动／设计意图）、"课前"行、"课前探索"行、"课中"行、"课中导学"与"导入Bridge-in"行、"目标Objective"行、"前测Pre-assessment"行、"参与式学习Participatory Learning"行、"后测Post-assessment"行、"总结Summary"行、"课后"行、"课后拓展／任务层／任务内容／评价标准"行、"L1（基础层）"行、"L2（进阶层）"行、"L3（挑战层）"行，或其中任一行文字被改写"""),
    # 24
    (-3, """"教学环节"表格"课前探索""课中导学""目标Objective""前测Pre-assessment""参与式学习Participatory Learning""后测Post-assessment""总结Summary"各行中，"教学内容""教师活动""学生活动""设计意图"四列的单元格出现无文字的空格"""),
    # 25
    (-1, """"教学启示"格三段（"本节课融合多种教学方法，实现'教-学-做-评'一体化：""1.任务驱动法  2.演示法  3.小组协作法  4.启发式教学法  5.过程性评价法""鼓励学生在观察、讨论、绘制和复盘中形成主动学习习惯。"）任一缺失或文字不符"""),
    # 26
    (-1, """"信息化手段及教学资源"栏各段（"专业为本，工具赋能，整合多元信息化手段服务教学。""教学资源""1.校内开放课程资源库：城市绿地系统规划基础模块（课程号：UG-204）""2.学院共享素材库：街区雨洪花园更新图片、矢量底图与课堂练习数据包""3.教师自编学习单：《场地解读图制作流程与检查清单》"）任一缺失、文字不符，或三条未各自独立成段而被合并"""),
    # 27
    (-1, """"反思诊改"表"特色创新"格3条（"1.BOPPPS与'分层递进、任务贯通'深度融合…""2.数字工具链回应场地解读教学中的资料整理、图面生成和逻辑表达痛点…""3.虚构项目贯穿全程…"）与"诊断改进"格3条（"1.分层指导优化…""2.逻辑建构强化…""3.价值引导深化…"）任一缺失、文字不符，或各条未各自独立成段而被合并"""),
    # 28
    (-1, """表格内文字的中文字体大部分不为Noto Sans CJK SC，或表格内文字的字号超过12磅"""),
    # 29
    (-1, """"《城市绿地系统规划》"（56磅加粗居中）、"韧性街区绿地的认知与设计"（小初加粗居中）、封面"第1节 场地解读图的作用及组织逻辑"（一号加粗居中）、"教学设计"（48磅加粗居中）、正文"第1节 场地解读图的作用及组织逻辑"（小二加粗）、"1.课程具体安排""2.教学环节""3.板书设计""4.教学评价"（均四号加粗）任一缺失、文字不符或字号不符"""),
    # 30
    (-1, """文档中出现批注、水印，或出现"已按要求调整""修改说明"之类的交付说明文字"""),
]
assert len(_DIM2_RULES) == 30

MAX_SCORE_TOTAL = sum(d for d, _ in _DIM2_RULES if d > 0)


def _build_dim2_items(results):
    """按“每条规则一个 item”生成严格二值的评分明细，返回 (items, total_score)。

    results 是判定阶段按 _DIM2_RULES 顺序追加的命中布尔值：
    正向项命中 = 达标得分，负向项命中 = 触发扣分。
    命中时 delta = max_delta，未命中时 delta = 0；total_score 为各项 delta 之和。
    每项只输出 rule / max_delta / delta / hit 四个键，判定说明由 rule 原文承载。
    """
    if len(results) != len(_DIM2_RULES):
        raise RuntimeError(
            'dim2 项数不一致：results=%d vs rules=%d' % (len(results), len(_DIM2_RULES)))
    items = []
    total_score = 0
    for hit, (max_delta, rule) in zip(results, _DIM2_RULES):
        delta = max_delta if hit else 0
        items.append({
            'rule': rule,
            'max_delta': max_delta,
            'delta': delta,
            'hit': bool(hit),
        })
        total_score += delta
    return items, total_score


_TWIPS_PER_CM = 567.0
_EMU_PER_CM = 360000.0
_CJK = re.compile(r'[一-鿿]')


def _flat(text):
    """归一化文本：去掉所有空白与各式引号，便于用锚点串做包含比对。"""
    if not text:
        return ''
    out = re.sub(r'[\s 　]+', '', text)
    for ch in ('“', '”', '"', '‘', '’', "'", '`'):
        out = out.replace(ch, '')
    return out


def _el_text(el):
    """取元素内全部 w:t 文本（含文本框 w:txbxContent，Cell.text 看不到）。"""
    if el is None:
        return ''
    return ''.join(t.text or '' for t in el.iter(qn('w:t')))


def _rows(tbl):
    return [] if tbl is None else tbl.findall(qn('w:tr'))


def _cells(tr):
    return [] if tr is None else tr.findall(qn('w:tc'))


def _paras(el):
    return [] if el is None else el.findall(qn('w:p'))


def _near(a, b, tol):
    return a is not None and abs(a - b) <= tol


def _tcpr(tc, tag):
    pr = None if tc is None else tc.find(qn('w:tcPr'))
    return None if pr is None else pr.find(qn('w:' + tag))


def _cell_w_cm(tc):
    """单元格宽度（cm）；非 dxa 计量或缺失时返回 None。"""
    w = _tcpr(tc, 'tcW')
    if w is None or (w.get(qn('w:type')) or 'dxa') != 'dxa':
        return None
    try:
        return int(w.get(qn('w:w'))) / _TWIPS_PER_CM
    except (TypeError, ValueError):
        return None


def _fill(tc):
    shd = _tcpr(tc, 'shd')
    if shd is None:
        return None
    val = shd.get(qn('w:fill'))
    return None if not val or val == 'auto' else val.upper()


def _valign(tc):
    """单元格垂直对齐；Word 默认值为顶端对齐。"""
    va = _tcpr(tc, 'vAlign')
    return 'top' if va is None else (va.get(qn('w:val')) or 'top')


def _p_spacing(p):
    """段落间距：(段前 twips, 行距值, 行距规则)。"""
    pPr = None if p is None else p.find(qn('w:pPr'))
    sp = None if pPr is None else pPr.find(qn('w:spacing'))
    if sp is None:
        return None, None, None

    def _int(name):
        try:
            return int(sp.get(qn('w:' + name)))
        except (TypeError, ValueError):
            return None

    return _int('before'), _int('line'), sp.get(qn('w:lineRule'))


def _p_first_line_chars(p):
    pPr = None if p is None else p.find(qn('w:pPr'))
    ind = None if pPr is None else pPr.find(qn('w:ind'))
    if ind is None:
        return None
    try:
        return int(ind.get(qn('w:firstLineChars')))
    except (TypeError, ValueError):
        return None


def _p_jc(p):
    pPr = None if p is None else p.find(qn('w:pPr'))
    jc = None if pPr is None else pPr.find(qn('w:jc'))
    return None if jc is None else jc.get(qn('w:val'))


def _imgs(el):
    """元素内插图的尺寸列表 [(宽 cm, 高 cm), ...]。"""
    out = []
    if el is None:
        return out
    for ext in el.iter(qn('wp:extent')):
        try:
            out.append((int(ext.get('cx')) / _EMU_PER_CM,
                        int(ext.get('cy')) / _EMU_PER_CM))
        except (TypeError, ValueError):
            continue
    return out


def _tbl_flat(tbl):
    return _flat(_el_text(tbl))


def _tbl_shape(tbl):
    """表格(行数, 列数)；列数按 w:tblGrid 的 w:gridCol 个数统计。"""
    if tbl is None:
        return 0, 0
    grid = tbl.find(qn('w:tblGrid'))
    cols = 0 if grid is None else len(grid.findall(qn('w:gridCol')))
    return len(_rows(tbl)), cols


def _find_tbl(tbls, *keys):
    """返回文本中同时含全部关键词的第 1 个表格。"""
    for tbl in tbls:
        ft = _tbl_flat(tbl)
        if all(_flat(k) in ft for k in keys):
            return tbl
    return None


def _row_of(rows, key):
    for tr in rows:
        if _flat(key) in _flat(_el_text(tr)):
            return tr
    return None


def _content_cell(tr, label):
    """标签格右侧的内容格。"""
    cells = _cells(tr)
    for i, tc in enumerate(cells):
        if _flat(label) in _flat(_el_text(tc)):
            return cells[i + 1] if i + 1 < len(cells) else None
    return None


def _cell_eq(tbl, text):
    """返回文本恰为 text 的第 1 个单元格。"""
    for tr in _rows(tbl):
        for tc in _cells(tr):
            if _flat(_el_text(tc)) == _flat(text):
                return tc
    return None


def _cell_startswith(tbl, text):
    for tr in _rows(tbl):
        for tc in _cells(tr):
            if _flat(_el_text(tc)).startswith(_flat(text)):
                return tc
    return None


def _xq_group(tbls):
    """定位“学情分析…教学创新点”表格组。

    正常交付时该组即 1 个 9 行表格；若表格被断开成多个，则把含“学情分析”的
    表格与其后紧邻的表格合并为一组，直到覆盖“教学创新点”为止，以便按元素
    逐项抽检格式（表格是否断开由加分项单独判定）。
    """
    start = None
    for i, tbl in enumerate(tbls):
        if '学情分析' in _tbl_flat(tbl):
            start = i
            break
    if start is None:
        return []
    group = [tbls[start]]
    i = start
    while i + 1 < len(tbls) and not any('教学创新点' in _tbl_flat(t) for t in group):
        nxt = tbls[i + 1]
        if '教学启示' in _tbl_flat(nxt):
            break
        group.append(nxt)
        i += 1
    return group


def _style_by_id(doc, style_id):
    for st in doc.styles.element.findall(qn('w:style')):
        if st.get(qn('w:styleId')) == style_id:
            return st
    return None


def _para_style_id(p):
    pPr = None if p is None else p.find(qn('w:pPr'))
    pStyle = None if pPr is None else pPr.find(qn('w:pStyle'))
    return None if pStyle is None else pStyle.get(qn('w:val'))


def _style_chain_rpr(doc, p):
    """段落样式（含 basedOn 上溯）与 docDefaults 的 w:rPr 序列。"""
    out = []
    style_id = _para_style_id(p)
    seen = set()
    while style_id and style_id not in seen:
        seen.add(style_id)
        st = _style_by_id(doc, style_id)
        if st is None:
            break
        rPr = st.find(qn('w:rPr'))
        if rPr is not None:
            out.append(rPr)
        based = st.find(qn('w:basedOn'))
        style_id = None if based is None else based.get(qn('w:val'))
    dd = doc.styles.element.find(qn('w:docDefaults'))
    if dd is not None:
        rpr_dd = dd.find(qn('w:rPrDefault'))
        rPr = None if rpr_dd is None else rpr_dd.find(qn('w:rPr'))
        if rPr is not None:
            out.append(rPr)
    return out


def _resolve_effective_eastasia(doc, r, p):
    """按 Word 字体继承链解析 run 中文字符实际生效的东亚字体：
       run 直接设定(w:eastAsia) → 段落样式(含 basedOn 链) → docDefaults。"""
    rPr = None if r is None else r.find(qn('w:rPr'))
    for pr in ([rPr] if rPr is not None else []) + _style_chain_rpr(doc, p):
        rFonts = pr.find(qn('w:rFonts'))
        if rFonts is not None:
            ea = rFonts.get(qn('w:eastAsia'))
            if ea:
                return ea
    return None


def _resolve_effective_sz(doc, r, p):
    """按继承链解析 run 实际生效的字号（磅）：run → 段落样式链 → docDefaults。"""
    rPr = None if r is None else r.find(qn('w:rPr'))
    for pr in ([rPr] if rPr is not None else []) + _style_chain_rpr(doc, p):
        sz = pr.find(qn('w:sz'))
        if sz is not None:
            try:
                return int(sz.get(qn('w:val'))) / 2.0
            except (TypeError, ValueError):
                continue
    return None


def _run_sz_pt(r):
    rPr = None if r is None else r.find(qn('w:rPr'))
    sz = None if rPr is None else rPr.find(qn('w:sz'))
    if sz is None:
        return None
    try:
        return int(sz.get(qn('w:val'))) / 2.0
    except (TypeError, ValueError):
        return None


def _run_bold(r):
    rPr = None if r is None else r.find(qn('w:rPr'))
    b = None if rPr is None else rPr.find(qn('w:b'))
    if b is None:
        return False
    return (b.get(qn('w:val')) or 'true') not in ('0', 'false', 'off')


def _text_runs(p):
    """段落内含可见文字的 run。"""
    out = []
    for r in p.findall(qn('w:r')):
        if _el_text(r).strip():
            out.append(r)
    return out


def _title_ok(p, pt, need_center):
    """标题段校验：文字 run 字号为 pt 磅、加粗；need_center 时还须居中。"""
    runs = _text_runs(p)
    if not runs:
        return False
    for r in runs:
        v = _run_sz_pt(r)
        if v is None or abs(v - pt) > 0.6 or not _run_bold(r):
            return False
    if need_center and _p_jc(p) != 'center':
        return False
    return True


def _has_page_break(p):
    """段落是否含分页符或“段前分页”。"""
    for br in p.iter(qn('w:br')):
        if br.get(qn('w:type')) == 'page':
            return True
    pPr = p.find(qn('w:pPr'))
    pbb = None if pPr is None else pPr.find(qn('w:pageBreakBefore'))
    if pbb is not None and (pbb.get(qn('w:val')) or 'true') not in ('0', 'false', 'off'):
        return True
    return False


def _has_watermark(el):
    for node in el.iter():
        for val in node.attrib.values():
            if isinstance(val, str) and 'watermark' in val.lower():
                return True
    return False


def _footer_has_page_field(el):
    """页脚内是否存在自动更新的页码域（PAGE 域）。"""
    for it in el.iter(qn('w:instrText')):
        if 'PAGE' in (it.text or '').upper():
            return True
    for fs in el.iter(qn('w:fldSimple')):
        if 'PAGE' in (fs.get(qn('w:instr')) or '').upper():
            return True
    return False


def _dim1_fail(result, reason):
    """维度一未通过：status 仍为 'ok'，完成度总分记 0，不再逐条罗列（未真正逐条评过）。"""
    result['dim1_pass'] = False
    result['dim1_reason'] = reason or '维度一未通过（未给出具体原因）'
    result['dim2_items'] = []
    result['total_score'] = 0
    result['max_score'] = MAX_SCORE_TOTAL
    return result


def _evaluate_impl(file_path, result):
    """逐条判定维度二：results 按 _DIM2_RULES 顺序追加 (命中, 说明)。

    命中语义：正向项命中 = 达标得分；负向项命中 = 触发扣分。
    """
    # ---------- 维度一：可用与可修改性 ----------
    # 文档定位到了却打不开属于“评了但不达标”，交给维度一记 0，而不是 status='error'。
    try:
        doc = Document(file_path)
    except Exception as exc:
        return _dim1_fail(result, '.docx 文件损坏无法打开（%s: %s）'
                          % (type(exc).__name__, exc))
    body = doc.element.body
    kids = list(body.iterchildren())
    tbls = [t._tbl for t in doc.tables]
    results = []

    # 文档能被 python-docx 打开即为可编辑的 .docx；再排除“整页图片/扫描件”
    # 这类无可编辑文本、无表格的情形。
    doc_text = _flat(_el_text(body))
    if not tbls and len(doc_text) < 50:
        return _dim1_fail(result, '文档无可编辑文本或表格，疑为图片/扫描件')

    # 维度一通过：按统一约定 dim1_reason 保持空串（只有未通过才写原因）。
    result['dim1_pass'] = True
    result['dim1_reason'] = ''

    # ---------- 定位五个目标表格 ----------
    group = _xq_group(tbls)                                   # 学情分析…教学创新点
    grows = [tr for tbl in group for tr in _rows(tbl)]
    t_kc = _find_tbl(tbls, '课程名称', '授课内容')             # 课程名称…教学内容
    t_qs = _find_tbl(tbls, '教学启示')                        # 教学启示…教学资源
    t_hj = _find_tbl(tbls, '教学环节', '设计意图')             # 教学环节
    t_fs = _find_tbl(tbls, '反思诊改')                        # 反思诊改

    # ===== 01 (+5) 六个标签同属一个表格，素质目标内容行与教学重点行紧密相接 =====
    owner = {}
    for lab in ('学情分析', '教学目标', '教学重点', '教学难点', '思政点', '教学创新点'):
        for i, tbl in enumerate(tbls):
            if lab in _tbl_flat(tbl):
                owner[lab] = i
                break
    same_tbl = len(owner) == 6 and len(set(owner.values())) == 1
    joined = False
    if same_tbl:
        rows = _rows(tbls[owner['学情分析']])
        i_goal = i_key = -1
        for i, tr in enumerate(rows):
            ft = _flat(_el_text(tr))
            if i_goal < 0 and '1.掌握场地解读图' in ft and '1.树立' in ft:
                i_goal = i
            if i_key < 0 and '教学重点' in ft:
                i_key = i
        joined = i_goal >= 0 and i_key == i_goal + 1
    results.append(same_tbl and joined)

    # ===== 02 (-3) 表格组须为 9 行且行序正确、无缺行/重复行/空白行 =====
    expect_rows = [
        ('学情分析', '知识基础'),
        ('能力水平',),
        ('行为特征',),
        ('教学目标', '知识目标', '能力目标', '素质目标'),
        ('1.掌握场地解读图的核心作用', '1.能借助数字工具完成柳岸口袋公园更新单元的资料收集', '1.树立'),
        ('教学重点',),
        ('教学难点',),
        ('思政点',),
        ('教学创新点',),
    ]
    rows_ok = len(grows) == 9
    if rows_ok:
        for tr, keys in zip(grows, expect_rows):
            ft = _flat(_el_text(tr))
            if not ft or not all(_flat(k) in ft for k in keys):
                rows_ok = False
                break
    results.append(not rows_ok)

    # ===== 03 (-3) 下半部分文字完整性 =====
    need = {
        '教学重点': ('依托层级逻辑定格局', '依托要素逻辑改基底', '依托问题逻辑补短板',
                     '依托潜力逻辑塑特色', '联动智能制图落地'),
        '教学难点': ('复杂场地资料的分类整合与逻辑建构', '规避工具依赖，兼顾专业判断与落地可行性'),
        '思政点': ('1.生态文明', '2.职业素养', '3.协作担当', '4.创新教育'),
    }
    text_ok = True
    for label, keys in need.items():
        tc = _content_cell(_row_of(grows, label), label)
        ft = _flat(_el_text(tc))
        if not all(_flat(k) in ft for k in keys):
            text_ok = False
    # 教学创新点：3 个小标题各自独立成段，且每个小标题后紧跟说明段
    tc_new = _content_cell(_row_of(grows, '教学创新点'), '教学创新点')
    subs = ('（一）教学模式创新：BOPPPS+数字工具+虚构项目三维融合',
            '（二）教学工具创新：智能检索+图像生成+空间资料处理的协同工作流',
            '（三）教学内容创新：理论—实操—应用全链条贯通')
    ps_new = _paras(tc_new)
    for sub in subs:
        idx = -1
        for i, p in enumerate(ps_new):
            if _flat(_el_text(p)).startswith(_flat(sub)):
                idx = i
                break
        if idx < 0 or idx + 1 >= len(ps_new) or len(_flat(_el_text(ps_new[idx + 1]))) < 20:
            text_ok = False
    results.append(not text_ok)

    # ===== 04 (-3) 表格组 9 行之间出现分页符或“段前分页” =====
    broke = False
    for tr in grows:
        for p in tr.iter(qn('w:p')):
            if _has_page_break(p):
                broke = True
                break
        if broke:
            break
    if not broke and len(group) > 1:
        try:
            lo, hi = kids.index(group[0]), kids.index(group[-1])
            for el in kids[lo + 1:hi]:
                if el.tag == qn('w:p') and _has_page_break(el):
                    broke = True
                    break
        except ValueError:
            pass
    results.append(broke)

    # ===== 05 (-1) 列宽：第1列1.34cm、第2列2.28cm、三个目标 4.22/4.25/5.10cm =====
    col_ok = len(grows) == 9
    if col_ok:
        for tr in grows:
            cs = _cells(tr)
            if not cs or not _near(_cell_w_cm(cs[0]), 1.34, 0.12):
                col_ok = False
        for i in (0, 1, 2, 5, 6, 7):
            cs = _cells(grows[i])
            if len(cs) < 2 or not _near(_cell_w_cm(cs[1]), 2.28, 0.15):
                col_ok = False
        for i in (3, 4):
            cs = _cells(grows[i])
            if len(cs) < 4 or not (_near(_cell_w_cm(cs[1]), 4.22, 0.15)
                                   and _near(_cell_w_cm(cs[2]), 4.25, 0.15)
                                   and _near(_cell_w_cm(cs[3]), 5.10, 0.15)):
                col_ok = False
    results.append(not col_ok)

    # ===== 06 (-1) 教学重点/难点/思政点内容各为 1 整格约 11.3cm；教学创新点行 2 格 =====
    span_ok = len(grows) == 9
    if span_ok:
        for i in (5, 6, 7):
            cs = _cells(grows[i])
            if len(cs) != 3 or not _near(_cell_w_cm(cs[2]), 11.3, 0.25):
                span_ok = False
        cs = _cells(grows[8])
        if len(cs) != 2 or not _near(_cell_w_cm(cs[0]), 1.34, 0.12) \
                or not _near(_cell_w_cm(cs[1]), 13.6, 0.25):
            span_ok = False
    results.append(not span_ok)

    # ===== 07 (-3) “教学内容”格内该段下方约 2 行空白、段前间距 5.85 磅 =====
    blank_ok = False
    cell_jx = None
    for tr in _rows(t_kc):
        for tc in _cells(tr):
            if '主要集中于场地解读图' in _flat(_el_text(tc)):
                cell_jx = tc
                break
    if cell_jx is not None:
        ps = _paras(cell_jx)
        for i, p in enumerate(ps):
            if '主要集中于场地解读图' in _flat(_el_text(p)):
                before = _p_spacing(p)[0]
                blanks = 0
                for q in ps[i + 1:]:
                    if _flat(_el_text(q)) or _imgs(q):
                        break
                    blanks += 1
                blank_ok = _near(before, 117, 8) and blanks >= 2
                break
    results.append(not blank_ok)

    # ===== 08 (-3) 任务层/任务内容/L1-L3 及任务描述：行距 1.27 倍、段前 13.05 磅 =====
    def _line_ratio_ok(p, ratio, tol=0.03):
        _, line, rule = _p_spacing(p)
        if line is None or (rule or 'auto') != 'auto':
            return False
        return abs(line / 240.0 - ratio) <= tol

    task_ok = True
    targets = []
    for label in ('任务层', '任务内容'):
        targets.append((_cell_eq(t_hj, label), True))
    for tr in _rows(t_hj):
        cs = _cells(tr)
        for i, tc in enumerate(cs):
            if _flat(_el_text(tc)).startswith(('L1（基础层）', 'L2（进阶层）', 'L3（挑战层）')):
                targets.append((tc, False))
                if i + 1 < len(cs):
                    targets.append((cs[i + 1], False))
    if len(targets) != 8:
        task_ok = False
    for tc, need_center in targets:
        ps = _paras(tc)
        if not ps:
            task_ok = False
            continue
        p = ps[0]
        if not _line_ratio_ok(p, 1.27) or not _near(_p_spacing(p)[0], 261, 10):
            task_ok = False
        if need_center and _p_jc(p) != 'center':
            task_ok = False
    results.append(not task_ok)

    # ===== 09 (-3) “评价标准”列 4 格：固定行距 18 磅、首行缩进 6/2 字符、对齐 =====
    eval_ok = True
    eval_cells = [(_cell_eq(t_hj, '评价标准'), 600, ('both',))]
    for key in ('资料完整度（40%）', '问题定位准确性（30%）', '策略可实施性（40%）'):
        eval_cells.append((_cell_startswith(t_hj, key), 200, (None, 'left', 'start')))
    for tc, flc, jcs in eval_cells:
        ps = _paras(tc)
        if not ps:
            eval_ok = False
            continue
        p = ps[0]
        _, line, rule = _p_spacing(p)
        if rule != 'exact' or not _near(line, 360, 6):
            eval_ok = False
        if _p_first_line_chars(p) != flc or _p_jc(p) not in jcs:
            eval_ok = False
    results.append(not eval_ok)

    # ===== 10 (-3) “行为特征”行内容格中约 11.25cm×5.04cm 的插图 =====
    tc_bh = _content_cell(_row_of(grows, '行为特征'), '行为特征')
    img_ok = any(_near(cx, 11.25, 0.3) and _near(cy, 5.04, 0.3) for cx, cy in _imgs(tc_bh))
    results.append(not img_ok)

    # ===== 11 (-3) 底纹配色 =====
    shd_ok = len(grows) == 9
    if shd_ok:
        for tr in grows:
            cs = _cells(tr)
            if not cs:
                shd_ok = False
                continue
            f0 = _fill(cs[0])
            if f0 is not None and f0 != 'DDEFE6':
                shd_ok = False
        if not any(_cells(tr) and _fill(_cells(tr)[0]) == 'DDEFE6' for tr in grows):
            shd_ok = False
        for i in (0, 1, 2, 5, 6, 7):
            cs = _cells(grows[i])
            if len(cs) < 2 or _fill(cs[1]) != 'F0F7F2':
                shd_ok = False
        for i, exp in ((0, 'E7E1F2'), (1, 'FBFAF4'), (2, 'FBFAF4'),
                       (5, 'E7E1F2'), (6, 'FBFAF4'), (7, 'FBFAF4')):
            cs = _cells(grows[i])
            if len(cs) < 3 or _fill(cs[2]) != exp:
                shd_ok = False
        for i in (3, 4):
            for tc in _cells(grows[i])[1:]:
                if _fill(tc) != 'FBFAF4':
                    shd_ok = False
        cs = _cells(grows[8])
        if len(cs) < 2 or _fill(cs[1]) != 'FBFAF4':
            shd_ok = False
    results.append(not shd_ok)

    # ===== 12 (-1) 垂直对齐：教学重点/教学难点标签格居中，其余顶端 =====
    va_ok = len(grows) == 9
    if va_ok:
        for i in (5, 6):
            cs = _cells(grows[i])
            if len(cs) < 2 or _valign(cs[1]) != 'center':
                va_ok = False
        for i, tr in enumerate(grows):
            for j, tc in enumerate(_cells(tr)):
                if (i, j) in ((5, 1), (6, 1)):
                    continue
                if _valign(tc) != 'top':
                    va_ok = False
    results.append(not va_ok)

    # ===== 13 (-1) 文档表格总数为 5 个 =====
    results.append(len(tbls) != 5)

    # ===== 14 (-1) 表格组与上下相邻表格之间各间隔 1 个空段落 =====
    gap_ok = False
    if t_kc is not None and group and t_qs is not None:
        try:
            segs = (kids[kids.index(t_kc) + 1:kids.index(group[0])],
                    kids[kids.index(group[-1]) + 1:kids.index(t_qs)])
            gap_ok = all(len(seg) == 1 and seg[0].tag == qn('w:p')
                         and not _flat(_el_text(seg[0])) for seg in segs)
        except ValueError:
            gap_ok = False
    results.append(not gap_ok)

    # ===== 15 (-3) 纸张尺寸 A4 =====
    a4_ok = bool(doc.sections)
    for sec in doc.sections:
        if not (_near(sec.page_width.cm, 21.0, 0.2) and _near(sec.page_height.cm, 29.7, 0.2)):
            a4_ok = False
    results.append(not a4_ok)

    # ===== 16 (-3) 页边距：正文各节上下2.54/左右3.17；末节 2.79/2.08/2.29/1.87 =====
    secs = list(doc.sections)
    margin_ok = len(secs) >= 2
    for sec in secs[:-1]:
        if not (_near(sec.top_margin.cm, 2.54, 0.05) and _near(sec.bottom_margin.cm, 2.54, 0.05)
                and _near(sec.left_margin.cm, 3.17, 0.05) and _near(sec.right_margin.cm, 3.17, 0.05)):
            margin_ok = False
    if secs:
        last = secs[-1]
        if not (_near(last.top_margin.cm, 2.79, 0.05) and _near(last.bottom_margin.cm, 2.08, 0.05)
                and _near(last.left_margin.cm, 2.29, 0.05) and _near(last.right_margin.cm, 1.87, 0.05)):
            margin_ok = False
    results.append(not margin_ok)

    # ===== 17 (-3) 页脚数量不少于 6 个，且页码为自动更新的页码域 =====
    fparts = [pt for pt in doc.part.package.parts if 'footer' in str(pt.partname).lower()]
    foot_ok = len(fparts) >= 6
    for pt in fparts:
        el = getattr(pt, 'element', None)
        if el is None:
            continue
        if not _el_text(el).strip() and not list(el.iter(qn('w:fldChar'))):
            continue          # 完全空白的页脚部件不参与页码域判定
        if not _footer_has_page_field(el):
            foot_ok = False
    results.append(not foot_ok)

    # ===== 18 (-3) 封面插图约 21.4cm×18.2cm =====
    cover_ok = False
    for el in kids:
        if el.tag != qn('w:p'):
            continue
        if any(_near(cx, 21.4, 0.35) and _near(cy, 18.2, 0.35) for cx, cy in _imgs(el)):
            cover_ok = True
            break
    results.append(not cover_ok)

    # ===== 19 (-3) “信息化手段”两段之间约 13.5cm×4.2cm 的插图 =====
    info_ok = False
    rows_qs = _rows(t_qs)
    i_tool = i_res = -1
    for i, tr in enumerate(rows_qs):
        flats = [_flat(_el_text(p)) for tc in _cells(tr) for p in _paras(tc)]
        if i_tool < 0 and any(t.startswith(_flat('专业为本，工具赋能')) for t in flats):
            i_tool = i
        elif i_tool >= 0 and i_res < 0 and any(t == '教学资源' for t in flats):
            i_res = i
    if i_tool >= 0 and i_res > i_tool:
        for tr in rows_qs[i_tool:i_res + 1]:
            if any(_near(cx, 13.5, 0.35) and _near(cy, 4.2, 0.35) for cx, cy in _imgs(tr)):
                info_ok = True
                break
    results.append(not info_ok)

    # ===== 20 (-3) “板书设计”标题下方约 6.1cm×9.7cm 的插图 =====
    board_ok = False
    body_ps = [el for el in kids if el.tag == qn('w:p')]
    for i, p in enumerate(body_ps):
        if not _flat(_el_text(p)).startswith('3.板书设计'):
            continue
        for q in body_ps[i:]:
            if _flat(_el_text(q)).startswith('4.教学评价'):
                break
            for cx, cy in _imgs(q):
                if (_near(cx, 6.1, 0.35) and _near(cy, 9.7, 0.4)) \
                        or (_near(cx, 9.7, 0.4) and _near(cy, 6.1, 0.35)):
                    board_ok = True
        break
    results.append(not board_ok)

    # ===== 21 (-3) 其余四个表格的行列数 =====
    shape_ok = (_tbl_shape(t_kc) == (6, 4) and _tbl_shape(t_qs) == (5, 2)
                and _tbl_shape(t_hj) == (15, 6) and _tbl_shape(t_fs) == (2, 3))
    results.append(not shape_ok)

    # ===== 22 (-3) 八个栏目及其内容齐全、次序不变 =====
    kc_flat = _tbl_flat(t_kc)
    labels_kc = ('课程名称', '授课内容', '授课对象', '教学时长',
                 '推荐教材', '参考教材', '教育教学理念', '教学内容')
    order_ok = True
    cursor = 0
    for lab in labels_kc:
        pos = kc_flat.find(_flat(lab), cursor)
        if pos < 0:
            order_ok = False
            break
        cursor = pos + len(_flat(lab))
    values_kc = ('城市绿地系统规划', '知识单元四', '雨洪花园更新与微地形解析',
                 '第1节场地解读图的作用及组织逻辑', '环境设计23-1/2/3', '1学时',
                 '周澜员.城市绿地系统规划方法（第2版）.澄景教育出版社,2022.',
                 '邵闻远.数字化设计基础', '林若禾.公园城市空间营造',
                 '以成果导向和项目化学习为主线',
                 '主要集中于场地解读图在韧性街区绿地认知与设计中的应用')
    for val in values_kc:
        if _flat(val) not in kc_flat:
            order_ok = False
    results.append(not order_ok)

    # ===== 23 (-3) “教学环节”表 15 行行序 =====
    expect_hj = [
        ('教学环节', '教学内容', '教师活动', '学生活动', '设计意图'),
        ('课前',), ('课前探索',), ('课中',),
        ('课中导学', '导入Bridge-in'), ('目标Objective',),
        ('前测Pre-assessment',), ('参与式学习Participatory Learning',),
        ('后测Post-assessment',), ('总结Summary',), ('课后',),
        ('课后拓展', '任务层', '任务内容', '评价标准'),
        ('L1（基础层）',), ('L2（进阶层）',), ('L3（挑战层）',),
    ]
    rows_hj = _rows(t_hj)
    hj_ok = len(rows_hj) == 15
    if hj_ok:
        for tr, keys in zip(rows_hj, expect_hj):
            ft = _flat(_el_text(tr))
            if not all(_flat(k) in ft for k in keys):
                hj_ok = False
                break
    results.append(not hj_ok)

    # ===== 24 (-3) 七个环节行的四列内容格不得为空 =====
    fill_ok = True
    hit_rows = 0
    for tr in rows_hj:
        ft = _flat(_el_text(tr))
        if not any(_flat(k) in ft for k in ('课前探索', '课中导学', '目标Objective',
                                            '前测Pre-assessment', '参与式学习Participatory Learning',
                                            '后测Post-assessment', '总结Summary')):
            continue
        hit_rows += 1
        cs = _cells(tr)
        if len(cs) < 5:
            fill_ok = False
            continue
        for tc in cs[1:5]:
            if not _flat(_el_text(tc)):
                fill_ok = False
    if hit_rows != 7:
        fill_ok = False
    results.append(not fill_ok)

    # ===== 25 (-1) “教学启示”格三段齐全 =====
    tc_qs = _content_cell(_row_of(_rows(t_qs), '教学启示'), '教学启示')
    ft_qs = _flat(_el_text(tc_qs))
    tips = ('本节课融合多种教学方法，实现教-学-做-评一体化：',
            '1.任务驱动法', '2.演示法', '3.小组协作法', '4.启发式教学法', '5.过程性评价法',
            '鼓励学生在观察、讨论、绘制和复盘中形成主动学习习惯。')
    tips_ok = all(_flat(k) in ft_qs for k in tips)
    results.append(not tips_ok)

    # ===== 26 (-1) “信息化手段及教学资源”栏各段齐全且三条各自独立成段 =====
    ps_qs = [p for tr in _rows(t_qs) for tc in _cells(tr) for p in _paras(tc)]
    flats_qs = [_flat(_el_text(p)) for p in ps_qs]
    info_txt_ok = any(t.startswith(_flat('专业为本，工具赋能，整合多元信息化手段服务教学。'))
                      for t in flats_qs) and any(t == '教学资源' for t in flats_qs)
    for item in ('1.校内开放课程资源库：城市绿地系统规划基础模块（课程号：UG-204）',
                 '2.学院共享素材库：街区雨洪花园更新图片、矢量底图与课堂练习数据包',
                 '3.教师自编学习单：《场地解读图制作流程与检查清单》'):
        if not any(t.startswith(_flat(item)) for t in flats_qs):
            info_txt_ok = False
    results.append(not info_txt_ok)

    # ===== 27 (-1) “反思诊改”表两栏各 3 条、各自独立成段 =====
    flats_fs = [_flat(_el_text(p)) for tr in _rows(t_fs) for tc in _cells(tr) for p in _paras(tc)]
    fs_ok = True
    for item in ('1.BOPPPS与', '2.数字工具链回应场地解读教学中的资料整理', '3.虚构项目贯穿全程',
                 '1.分层指导优化', '2.逻辑建构强化', '3.价值引导深化'):
        if not any(t.startswith(_flat(item)) for t in flats_fs):
            fs_ok = False
    results.append(not fs_ok)

    # ===== 28 (-1) 表格内中文字体以 Noto Sans CJK SC 为主，字号不超过 12 磅 =====
    noto = cjk_runs = 0
    oversize = False
    for tbl in tbls:
        for p in tbl.iter(qn('w:p')):
            for r in p.findall(qn('w:r')):
                txt = _el_text(r)
                if not txt.strip():
                    continue
                sz = _resolve_effective_sz(doc, r, p)
                if sz is not None and sz > 12.5:
                    oversize = True
                if not _CJK.search(txt):
                    continue
                cjk_runs += 1
                ea = _resolve_effective_eastasia(doc, r, p)
                if ea and 'Noto Sans CJK' in ea:
                    noto += 1
    ratio = (noto / cjk_runs) if cjk_runs else 0.0
    font_ok = cjk_runs > 0 and ratio >= 0.6 and not oversize
    results.append(not font_ok)

    # ===== 29 (-1) 六组标题的文字与字号 =====
    title_specs = (('《城市绿地系统规划》', 56, True),
                   ('韧性街区绿地的认知与设计', 36, True),
                   ('第1节 场地解读图的作用及组织逻辑', 26, True),
                   ('教学设计', 48, True),
                   ('第1节 场地解读图的作用及组织逻辑', 18, False),
                   ('1.课程具体安排', 14, False), ('2.教学环节', 14, False),
                   ('3.板书设计', 14, False), ('4.教学评价', 14, False))
    title_ok = True
    for text, pt, need_center in title_specs:
        if not any(_flat(_el_text(p)) == _flat(text) and _title_ok(p, pt, need_center)
                   for p in body_ps):
            title_ok = False
    results.append(not title_ok)

    # ===== 30 (-1) 批注、水印或交付说明文字 =====
    has_comment = any('comments' in str(pt.partname).lower() for pt in doc.part.package.parts)
    has_wm = False
    for pt in doc.part.package.parts:
        pn = str(pt.partname).lower()
        if 'header' in pn or 'footer' in pn or pn.endswith('/document.xml'):
            el = getattr(pt, 'element', None)
            if el is not None and _has_watermark(el):
                has_wm = True
    has_note = any(k in doc_text for k in ('已按要求调整', '修改说明'))
    results.append(has_comment or has_wm or has_note)

    # ========== 生成结构化 dim2_items ==========
    dim2_items, total_score = _build_dim2_items(results)

    result['dim2_items'] = dim2_items
    result['total_score'] = total_score
    result['max_score'] = MAX_SCORE_TOTAL
    return result


def evaluate(dir_path: str) -> dict:
    """统一入口：接收脚本所在目录，定位并评估目录内的待评估 Word 文档。

    评分口径：维度一为门槛项，未通过则 dim2_items 置空、total_score 记 0，
    max_score 仍为正向项满分之和（MAX_SCORE_TOTAL）；维度二逐条二值判定，
    total_score 为各项 delta 之和（正向项得分、负向项扣分，不做下限截断）。

    status='error' 只在“这份提交没被真正评过”时出现——目录不存在、目录内没有
    可评估的 Word 文档、或脚本自身抛异常；此时 dim1_pass 为空串（未判定）、
    dim1_reason 为空串、dim2_items 为空表、total_score 与 max_score 一律为 0。
    文档能定位到但不达标仍是 status='ok'；维度一未通过时 dim1_pass=False 且
    dim1_reason 带有未通过的错误摘要。

    参数:
        dir_path: 脚本（及被评估 Word 文档）所在目录路径。
    返回:
        dict，字段对齐统一接口约定 §2.2。
    """
    result: dict[str, object] = {
        'id': SCRIPT_ID,
        'file_name': '',
        'status': 'ok',
        'error': None,
        # 空串表示“维度一尚未判定”；status='error' 的三条出口都保持空串，
        # 只有真正跑过维度一门槛判定才会写入 True / False。
        'dim1_pass': '',
        'dim1_reason': '',
        'dim2_items': [],
        'total_score': 0,
        'max_score': 0,
    }

    try:
        if not os.path.isdir(dir_path):
            result['status'] = 'error'
            result['error'] = f'目录不存在：{dir_path}'
            return result

        cands = sorted(f for f in os.listdir(dir_path)
                       if f.lower().endswith('.docx') and not f.startswith('~$'))
        if not cands:
            result['status'] = 'error'
            result['error'] = f'目录中未找到 .docx 文件：{dir_path}'
            return result

        file_path = os.path.join(dir_path, cands[0])
        result['file_name'] = os.path.basename(file_path)
        return _evaluate_impl(file_path, result)

    except Exception as exc:
        # _evaluate_impl 是就地写 result 的，异常可能发生在它写了一半之后，
        # 因此这里把所有承载评分结论的字段显式复位，保证 error 的返回值是干净的全 0；
        # dim1_pass 复位为空串而不是 False——异常意味着维度一的结论作废，不是判定为不通过。
        result['status'] = 'error'
        result['error'] = f'{type(exc).__name__}: {exc}'
        result['dim1_pass'] = ''
        result['dim1_reason'] = ''
        result['dim2_items'] = []
        result['total_score'] = 0
        result['max_score'] = 0
        return result


if __name__ == '__main__':
    # 本地调试入口：命令行第 1 个参数可覆盖目录路径；未指定时使用脚本所在目录。
    # 直接写 stdout 的二进制缓冲区并显式指定 utf-8，避免受控制台默认编码影响。
    _dir = sys.argv[1] if len(sys.argv) > 1 else os.path.dirname(os.path.abspath(__file__))
    _output = json.dumps(evaluate(_dir), ensure_ascii=False, indent=2) + '\n'
    sys.stdout.buffer.write(_output.encode('utf-8', errors='replace'))
