#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# pyright: reportUnknownMemberType=false, reportUnknownVariableType=false, reportUnknownArgumentType=false, reportUnknownParameterType=false, reportMissingParameterType=false, reportPrivateUsage=false, reportUnknownLambdaType=false, reportGeneralTypeIssues=false

import os
import re
import sys
import json
import traceback
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn

# 说明：按接口统一约定，本模块不修改全局 sys.stdout，也不 print 主结果。
# 主结果统一由 evaluate(dir_path) -> dict 返回；入参为"脚本所在目录的路径"，
# 由脚本自己在该目录内定位并打开被评估的 .docx。本文件仅在 __main__ 下作为
# 本地调试入口，才把 dict 序列化为 JSON 打印。
#
# 本任务交付物为 2 个 Word 文档，脚本自动按内容特征分工：
#   文档A = 由「企业经营情况说明书」改造而成的呈报材料（封面 + 目录 + 正文 + 4 表 + 1 图）
#   文档B = 主题替换后的专项核查材料（正文 + 3 表）
# 两份文档的得分合并计算：加分项来自指令下达的版式改造/封面/目录/主题替换；
# 减分项来自指令未涉及、却被改动而破坏既有内容与版式之处（max_delta 为负）。

SCRIPT_ID = "022"
# 维度二"加分项"max_delta 之和；用于维度一未通过 / 脚本异常时兜底
MAX_SCORE_TOTAL = 58
# 维度二"减分项"max_delta 之和（负值，可累计扣减）
MAX_PENALTY_TOTAL = -32

# 维度二所有评分项按"结果追加顺序"的 max_delta 列表；末端把 results 里的
# "+N：..." / "-N：..." 字符串反解为结构化 dim2_items 时用。顺序须与
# _evaluate_impl 中 results.append(...) 的顺序严格一致；共 44 项。
_MAX_DELTAS_ORDER = [
    # ---------- 加分项：页面与页脚 ----------
    3,   # 01 全文 A4
    3,   # 02 页边距 25/25/20/20mm
    3,   # 03 页脚"第X页共Y页"页码域
    1,   # 04 页脚小五号(9磅)居中
    # ---------- 加分项：封面页 ----------
    5,   # 05 单独成页封面 + 14 项信息
    3,   # 06 封面主标题 黑体二号(22磅)加粗居中
    1,   # 07 封面其余各行 宋体小三(15磅)居中
    1,   # 08 "━━━"分隔条 10磅 居中
    1,   # 09 未保留"---"占位符
    # ---------- 加分项：目录页 ----------
    5,   # 10 TOC 目录域（含页码域，可跳转）
    3,   # 11 目录标题"目录" 黑体小三加粗居中
    3,   # 12 仅一级标题层级/文字一致/页码靠右且相符
    1,   # 13 目录单独成页，其后正文另起一页
    # ---------- 加分项：标题体系 ----------
    3,   # 14 六个一级标题改为"第X部分：..."
    3,   # 15 一级标题 黑体小三(15磅)
    3,   # 16 三个二级标题改为"一、/二、/三、"
    1,   # 17 二级标题 黑体四号(14磅)左对齐
    1,   # 18 一/二级标题 1.5 倍行距
    1,   # 19 一/二级标题不再使用蓝色 2D5D91
    # ---------- 加分项：正文与图注 ----------
    3,   # 20 正文 宋体小四(12磅)
    3,   # 21 正文 1.25 倍行距 + 段后 4 磅
    1,   # 22 正文首行缩进 2 字符(24磅)
    1,   # 23 图注 黑体五号(10.5磅)居中
    1,   # 24 4 个表格单元格居中 + 小四(12磅)
    # ---------- 加分项：主题替换 ----------
    3,   # 25 首行主标题替换为新主题（四段要素齐全、逐字一致）
    1,   # 26 不再保留"专项核查报告"独立行
    # ---------- 减分项 ----------
    -3,  # 27 封面多余行"华中区域零售项目经营异常核查封面"
    -3,  # 28 正文起始处说明信息四行缺失
    -3,  # 29 表格网格框线丢失
    -1,  # 30 单元格插入多余空文本段
    -1,  # 31 表头行丢失加粗或居中
    -1,  # 32 图注文字被改写
    -1,  # 33 文末签署两行丢失右对齐
    -1,  # 34 文末签署两行被加首行缩进/段后间距
    -1,  # 35 主题文档主标题版式被改动
    -3,  # 36 主题文档八个标题不满足四号加粗
    -1,  # 37 主题文档编制人/日期两行不满足四号加粗右对齐
    -1,  # 38 主题文档西文被批量覆盖为 Times New Roman
    -1,  # 39 主题文档行距/首行缩进设定被改动
    -1,  # 40 主标题下方多余空段落或空白间距
    -3,  # 41 删除或改写正文文字内容
    -3,  # 42 删除或破坏已有表格、插图
    -3,  # 43 删除已有段落或章节
    -1,  # 44 改动姓名/联系方式/金额/日期数值
]
assert sum(v for v in _MAX_DELTAS_ORDER if v > 0) == MAX_SCORE_TOTAL
assert sum(v for v in _MAX_DELTAS_ORDER if v < 0) == MAX_PENALTY_TOTAL
assert len(_MAX_DELTAS_ORDER) == 44

# ==================== 期望文本常量 ====================

# 文档A 六个一级标题（改造后应逐字一致）
L1_TITLES_A = [
    "第一部分：项目背景",
    "第二部分：主要关联主体情况",
    "第三部分：异常资金往来情况",
    "第四部分：业务流程变化说明",
    "第五部分：库存与资金需求分析",
    "第六部分：综合说明",
]
# 文档A 三个二级标题（改造后应逐字一致）
L2_TITLES_A = [
    "一、提前付款情况",
    "二、平台收入回流滞后情况",
    "三、重复收费及费用边界不清情况",
]
# 文档A 封面页应承载的 14 项信息（取可判定的关键片段，规避空格排布差异）
COVER_ITEMS = [
    "专项核查呈报材料",
    "华中区域零售项目经营资金异常及现金流风险事项",
    "核心核查事项",
    "320万元供应链预付款提前支付及资金占用",
    "118万元线上平台结算款延迟回流",
    "76万元运营服务费重复计费",
    "210万元内部资金调拨用途待核",
    "181万元开办筹备、物流代垫及促销费用待补凭证",
    "呈报人：林远芝",
    "华中区域商业项目合作方",
    "联系方式：186****2741",
    "联系地址：华中地区羽厘商业办公区",
    "日期：2026年5月20日",
    "待核查影响资金总额：905万元",
]
COVER_MAIN_TITLE = "专项核查呈报材料"
# 文档A 正文起始处的说明信息四行
INTRO_LINES_A = [
    "说明人：林远芝",
    "联系方式：186****2741",
    "联系地址：华中地区羽厘商业办公区",
    "提交日期：2026年4月18日",
]
# 文档A 文末签署两行
SIGN_LINES_A = ["说明人：林远芝", "2026年05月20日"]
# 文档A 图注（应逐字保持）
FIG_CAPTION_A = "图1：项目资金流转示意图"
FIG_CAPTION_KEY = "项目资金流转示意图"

# 文档B 替换后的新主题
NEW_THEME_B = "华中区域零售项目财务资金异常、库存占用及现金流风险专项核查材料"
# 新主题的四段要素（地域项目 / 财务资金异常 / 库存占用 / 现金流风险 + 材料属性）
NEW_THEME_PARTS = ["华中区域零售项目", "财务资金异常", "库存占用", "现金流风险专项核查材料"]
OLD_THEME_LINE_B = "专项核查报告"
# 文档B 八个一级标题
L1_TITLES_B = [
    "一、基本概况",
    "二、关联合作主体及职责边界",
    "三、专项核查发现资金异常事项",
    "四、业务执行流程偏离及成因分析",
    "五、库存占用及现金流压力测算",
    "六、资金异常及现金流风险综合影响",
    "七、专项核查整改及风险化解建议",
    "八、核查结论",
]
# 文档B 文末两行签署
SIGN_LINES_B = ["核查编制人：林远芝", "编制日期：2026年5月20日"]

# 文档B 十六个三级小标题（应逐条保留）
L3_TITLES_B = [
    "（一）项目概况", "（二）核查范围与依据",
    "（一）预付款提前支付", "（二）线上平台收入回流滞后", "（三）运营服务费重复计费",
    "（四）费用分摊及单据管理缺失", "（五）内部资金调拨用途不明",
    "（一）原约定标准流程", "（二）实际执行流程偏离", "（三）核心成因",
    "（一）全面梳理账务", "（二）规范平台结算", "（三）重构费用科目",
    "（四）建立资金管控机制", "（五）完善业务流程及资料归档", "（六）盘活库存及应收",
]

# 文档A 正文语句锚点：逐段抽取的特征串，用于判定正文文字是否被删除或改写
BODY_ANCHORS_A = [
    "相关主体名称、资金数据及业务细节均已进行调整和重构",
    "由于部分原始凭证仍需进一步补充核验",
    "异常集中出现在付款审批依据不足",
    "项目定位为覆盖周边社区消费需求的综合零售门店",
    "部分款项在货物交付或服务确认前已提前支付",
    "信息同步不及时会直接影响资金使用效率",
    "部分交易在执行层面出现了职责边界不清的情况",
    "平台代运营方则掌握线上订单、退货、佣金及结算周期等数据",
    "若各主体之间没有建立统一的对账口径",
    "导致资金流向与业务实质之间出现偏差",
    "现有资料显示，部分款项在到货验收前已完成支付",
    "影响门店日常补货、人员支出和营销活动安排",
    "再由其根据内部规则进行延后结算",
    "造成销售数据与现金流表现不一致",
    "存在同一管理职能被重复计入不同科目的风险",
    "供应链主体完成发货和入库交接",
    "线上渠道订单收入不直接进入项目主体账户",
    "但应当配套形成书面变更记录",
    "还应同步审查资金发生前后的业务依据是否完整",
    "前期库存规模保持在较高水平",
    "但下降速度不及资金占用改善速度",
    "形成项目资金占用总表",
    "等待平台收入结算期间的周转资金",
    "也降低了各合作主体之间的对账效率",
    "重复收费及费用边界不清则可能导致成本被高估",
    "要求平台代运营主体提供完整订单流水",
    "应约定固定结算日和逾期说明机制",
    "则应将相关款项列入专项核查范围",
]
# 文档B 正文语句锚点
BODY_ANCHORS_B = [
    "远艘零售项目部为项目主体及实际经营方",
    "项目以月度对账为基准",
    "项目资金收支、预付款项、平台结算",
    "未建立统一对账口径",
    "造成账面预付款余额居高不下",
    "平台结算款形成长期往来挂账",
    "既列支固定管理服务费",
    "内部账户资金调拨无完整业务说明",
    "降幅远低于资金占用改善速度",
    "纳入专项追偿核查",
]

# 文档A/B 既有表格与插图规模（行数, 列数）
TABLE_SHAPES_A = [(6, 3), (8, 4), (5, 4), (7, 4)]
TABLE_SHAPES_B = [(6, 5), (8, 4), (7, 4)]

# 姓名 / 联系方式 / 金额 / 日期等数值：宿主片段仍在时，对应数值必须原样保留
VALUE_PAIRS_A = [
    ("说明人：", "林远芝"),
    ("说明人：林远芝", "2026年05月20日"),
    ("联系方式：", "186****2741"),
    ("联系地址：", "华中地区羽厘商业办公区"),
    ("提交日期：", "2026年4月18日"),
    ("呈报人：", "林远芝"),
    ("待核查影响资金总额", "905万元"),
    ("项目于2025年启动运营", "2025年"),
]
VALUE_ROWS_A = [
    ("开办筹备款", "85"), ("预付款项", "320"), ("平台结算挂账", "118"),
    ("物流代垫费用", "52"), ("运营服务费", "76"), ("促销费用分摊", "44"),
    ("资金调拨", "210"), ("6月", "410"), ("11月", "358"),
]
VALUE_PAIRS_B = [
    ("核查编制人：", "林远芝"),
    ("编制日期：", "2026年5月20日"),
    ("平台结算款形成长期往来挂账", "118万元"),
    ("运营服务费存在双重计费问题", "76万元"),
    ("内部账户资金调拨无完整业务说明", "210万元"),
    ("逐笔建档", "85万开办款"),
    ("逐笔建档", "320万预付款"),
    ("逐笔建档", "210万资金调拨"),
]

# 一/二级标题在参照模板中被要求不再使用的强调色
FORBIDDEN_HEADING_COLOR = "2D5D91"

# 字号容差（磅）；Word 以半磅存储，取 ±0.4 足以覆盖 21/22 半磅之类的邻值差异
_SZ_TOL = 0.4


# ==================== 通用工具 ====================

def _norm(text):
    """归一化：去掉全部空白（含全角空格），便于比较排布带空格的标题行。"""
    if not text:
        return ""
    return re.sub(r"\s+", "", text.replace("　", " "))


def _has_cjk(text):
    return bool(re.search(r"[一-鿿]", text or ""))


def _close(value, target, tol):
    return value is not None and abs(value - target) <= tol

def _child(el, path):
    """按 qn 路径逐级取子元素；任一级缺失返回 None。"""
    cur = el
    for tag in path:
        if cur is None:
            return None
        cur = cur.find(qn(tag))
    return cur


def _doc_default_rpr(doc):
    return _child(doc.styles.element, ("w:docDefaults", "w:rPrDefault", "w:rPr"))


def _doc_default_ppr(doc):
    return _child(doc.styles.element, ("w:docDefaults", "w:pPrDefault", "w:pPr"))


def _style_elements(para):
    """段落样式 -> basedOn 链上的 w:style 元素序列（Word 的继承顺序）。"""
    out = []
    try:
        st = para.style
    except Exception:
        st = None
    guard = 0
    while st is not None and guard < 12:
        el = getattr(st, "element", None)
        if el is not None:
            out.append(el)
        try:
            st = st.base_style
        except Exception:
            st = None
        guard += 1
    return out


def _rpr_sources(doc, run, para):
    """run 直接格式 -> 段落样式(含 basedOn 链) -> 文档默认，按优先级返回 rPr 列表。"""
    srcs = []
    if run is not None:
        srcs.append(run._element.find(qn("w:rPr")))
    if para is not None:
        for sel in _style_elements(para):
            srcs.append(sel.find(qn("w:rPr")))
    srcs.append(_doc_default_rpr(doc))
    return [s for s in srcs if s is not None]


def _ppr_sources(doc, para):
    srcs = [para._p.find(qn("w:pPr"))]
    for sel in _style_elements(para):
        srcs.append(sel.find(qn("w:pPr")))
    srcs.append(_doc_default_ppr(doc))
    return [s for s in srcs if s is not None]


def _first(sources, getter):
    for src in sources:
        val = getter(src)
        if val is not None:
            return val
    return None

def _g_font(attr):
    def getter(rpr):
        fonts = rpr.find(qn("w:rFonts"))
        if fonts is None:
            return None
        return fonts.get(qn("w:" + attr))
    return getter


def _g_sz(rpr):
    sz = rpr.find(qn("w:sz"))
    if sz is None:
        return None
    try:
        return float(sz.get(qn("w:val"))) / 2.0
    except (TypeError, ValueError):
        return None


def _g_bold(rpr):
    b = rpr.find(qn("w:b"))
    if b is None:
        return None
    val = b.get(qn("w:val"))
    return val not in ("0", "false", "off")


def _g_color(rpr):
    c = rpr.find(qn("w:color"))
    if c is None:
        return None
    return (c.get(qn("w:val")) or "").upper() or None


def _eastasia(doc, run, para):
    """解析 run 生效的中文字体：直接格式 -> 段落样式链 -> 文档默认。"""
    return _first(_rpr_sources(doc, run, para), _g_font("eastAsia"))


def _ascii_font(doc, run, para):
    return _first(_rpr_sources(doc, run, para), _g_font("ascii"))


def _size_pt(doc, run, para):
    return _first(_rpr_sources(doc, run, para), _g_sz)


def _bold(doc, run, para):
    return bool(_first(_rpr_sources(doc, run, para), _g_bold))


def _color(doc, run, para):
    return _first(_rpr_sources(doc, run, para), _g_color)

def _g_jc(ppr):
    jc = ppr.find(qn("w:jc"))
    return jc.get(qn("w:val")) if jc is not None else None


def _g_spacing(attr):
    def getter(ppr):
        sp = ppr.find(qn("w:spacing"))
        if sp is None:
            return None
        raw = sp.get(qn("w:" + attr))
        if raw is None:
            return None
        try:
            return float(raw)
        except ValueError:
            return None
    return getter


def _g_line_rule(ppr):
    sp = ppr.find(qn("w:spacing"))
    if sp is None:
        return None
    return sp.get(qn("w:lineRule"))


def _g_ind(attr):
    def getter(ppr):
        ind = ppr.find(qn("w:ind"))
        if ind is None:
            return None
        raw = ind.get(qn("w:" + attr))
        if raw is None:
            return None
        try:
            return float(raw)
        except ValueError:
            return None
    return getter


def _align(doc, para):
    """生效对齐；Word 默认（无 w:jc）等价于左对齐。"""
    val = _first(_ppr_sources(doc, para), _g_jc)
    if val in ("start",):
        return "left"
    if val in ("end",):
        return "right"
    return val or "left"


def _line_multiple(doc, para):
    """生效行距倍数；仅当 lineRule 为 auto（倍数行距）时有意义。"""
    srcs = _ppr_sources(doc, para)
    line = _first(srcs, _g_spacing("line"))
    if line is None:
        return None
    rule = _first(srcs, _g_line_rule) or "auto"
    if rule != "auto":
        return None
    return line / 240.0

def _space_pt(doc, para, attr):
    val = _first(_ppr_sources(doc, para), _g_spacing(attr))
    return None if val is None else val / 20.0


def _first_line_pt(doc, para):
    """生效首行缩进（磅）；同时兼容 firstLine（缇）与 firstLineChars（字符×100）。"""
    srcs = _ppr_sources(doc, para)
    hanging = _first(srcs, _g_ind("hanging"))
    if hanging:
        return 0.0
    twips = _first(srcs, _g_ind("firstLine"))
    if twips is not None:
        return twips / 20.0
    chars = _first(srcs, _g_ind("firstLineChars"))
    if chars is not None:
        size = _para_size(doc, para) or 12.0
        return chars / 100.0 * size
    return None


def _direct_ppr(para):
    return para._p.find(qn("w:pPr"))


def _direct_val(para, getter):
    """只看段落自身的直接格式（不含样式继承），用于判定"是否被额外添加"。"""
    ppr = _direct_ppr(para)
    return None if ppr is None else getter(ppr)


def _visible_runs(para):
    return [r for r in para.runs if r.text and r.text.strip()]


def _para_size(doc, para):
    """段落代表字号：取第一个可见 run 的生效字号。"""
    for run in _visible_runs(para):
        val = _size_pt(doc, run, para)
        if val is not None:
            return val
    return _size_pt(doc, None, para)


def _break_inside(para):
    for br in para._p.iter(qn("w:br")):
        if br.get(qn("w:type")) == "page":
            return True
    return False


def _break_before(para):
    ppr = _direct_ppr(para)
    if ppr is None:
        return False
    node = ppr.find(qn("w:pageBreakBefore"))
    if node is None:
        return False
    return node.get(qn("w:val")) not in ("0", "false", "off")


def _has_drawing(container):
    for tag in ("w:drawing", "w:pict", "w:object"):
        if container.find(".//" + qn(tag)) is not None:
            return True
    return False

def _field_codes(element):
    """收集域代码（含 w:instrText 与 w:fldSimple/@w:instr），大写返回便于匹配。"""
    parts = []
    for node in element.iter(qn("w:instrText")):
        parts.append(node.text or "")
    for node in element.iter(qn("w:fldSimple")):
        parts.append(node.get(qn("w:instr")) or "")
    return " ".join(parts).upper()


_BORDER_EDGES = ("top", "left", "bottom", "right", "insideH", "insideV")


def _borders_ok(tbl_pr):
    if tbl_pr is None:
        return False
    borders = tbl_pr.find(qn("w:tblBorders"))
    if borders is None:
        return False
    for edge in _BORDER_EDGES:
        node = borders.find(qn("w:" + edge))
        if node is None:
            return False
        if (node.get(qn("w:val")) or "none") in ("none", "nil"):
            return False
    return True


def _table_grid_ok(table):
    """网格框线：先看表格直接 tblPr，再沿表格样式 basedOn 链回溯。"""
    if _borders_ok(table._tbl.find(qn("w:tblPr"))):
        return True
    try:
        st = table.style
    except Exception:
        st = None
    guard = 0
    while st is not None and guard < 12:
        el = getattr(st, "element", None)
        if el is not None and _borders_ok(el.find(qn("w:tblPr"))):
            return True
        try:
            st = st.base_style
        except Exception:
            st = None
        guard += 1
    return False


def _table_shapes(doc):
    return [(len(t.rows), len(t.columns)) for t in doc.tables]


def _effective_footers(doc):
    """按节返回生效页脚（linked 的节回溯到前一节），去重后用于页脚判定。"""
    footers = []
    last = None
    for sec in doc.sections:
        ftr = sec.footer
        if ftr.is_linked_to_previous:
            if last is not None:
                footers.append(last)
        else:
            last = ftr
            footers.append(ftr)
    if not footers and doc.sections:
        footers.append(doc.sections[0].footer)
    return footers

def _doc_paragraph_texts(doc):
    return [_norm(p.text) for p in doc.paragraphs]


def _doc_table_texts(doc):
    rows = []
    for table in doc.tables:
        for row in table.rows:
            rows.append(_norm("|".join(c.text for c in row.cells)))
    return rows


def _doc_all_text(doc):
    return "".join(_doc_paragraph_texts(doc)) + "".join(_doc_table_texts(doc))


def _cell_paragraphs(table):
    for row in table.rows:
        for cell in row.cells:
            for para in cell.paragraphs:
                yield cell, para


def _has_sect_break(para):
    ppr = _direct_ppr(para)
    return ppr is not None and ppr.find(qn("w:sectPr")) is not None


def _is_toc_entry(para):
    """目录条目：TOC 域生成的 toc 样式段、含 PAGEREF 域，或"标题+制表位+页码"文本。"""
    try:
        name = (para.style.name or "").lower()
    except Exception:
        name = ""
    if name.startswith("toc") or name.startswith("目录"):
        return True
    if "PAGEREF" in _field_codes(para._p):
        return True
    raw = para.text or ""
    if re.search(r"[\t\.·…]\s*\d+\s*$", raw) and len(_norm(raw)) < 40:
        return True
    return False


def _regions_a(doc):
    """切分文档A：封面区 / 目录标题 / 目录条目区 / 正文起点。"""
    paras = doc.paragraphs
    toc_title = None
    for i, para in enumerate(paras[:60]):
        if _norm(para.text) == "目录":
            toc_title = i
            break
    first_break = None
    for i, para in enumerate(paras):
        if _break_inside(para) or _has_sect_break(para) or (i and _break_before(para)):
            first_break = i
            break
    cover_end = (toc_title - 1) if toc_title is not None else (
        first_break if first_break is not None else -1)
    body_start = None
    if toc_title is not None:
        for i in range(toc_title + 1, len(paras)):
            para = paras[i]
            if _break_inside(para) or _has_sect_break(para):
                body_start = i + 1
                break
            if _break_before(para) and not _is_toc_entry(para):
                body_start = i
                break
        if body_start is None:
            for i in range(toc_title + 1, len(paras)):
                if _norm(paras[i].text) and not _is_toc_entry(paras[i]):
                    body_start = i
                    break
    if body_start is None:
        body_start = (first_break + 1) if first_break is not None else 0
    return {"cover_end": cover_end, "toc_title": toc_title,
            "body_start": min(body_start, len(paras))}

_MARKERS_A = ("专项核查呈报材料", "项目资金流转示意图", "提前付款情况",
              "说明人：", "库存与资金需求分析")
_MARKERS_B = ("基本概况", "关联合作主体及职责边界", "核查编制人",
              "核查结论", "核查范围与依据")


def _open_word_document(file_path):
    """打开 .docx；打不开（损坏 / 非 Word 格式 / 加密）时抛出异常由上层捕获。"""
    return Document(file_path)


def _classify(candidates):
    """按内容特征把候选 .docx 分派为 文档A（呈报材料）/ 文档B（主题替换材料）。"""
    scored = []
    for path in candidates:
        try:
            doc = _open_word_document(path)
        except Exception:
            continue
        text = _doc_all_text(doc)
        sa = sum(1 for m in _MARKERS_A if m in text)
        sb = sum(1 for m in _MARKERS_B if m in text)
        scored.append({"path": path, "doc": doc, "a": sa, "b": sb})
    if not scored:
        return None, None
    pick_a = max(scored, key=lambda s: (s["a"] - s["b"], s["a"]))
    rest = [s for s in scored if s is not pick_a]
    pick_b = max(rest, key=lambda s: (s["b"] - s["a"], s["b"])) if rest else None
    if pick_a["a"] < 2:
        pick_a = None
    if pick_b is not None and pick_b["b"] < 2:
        pick_b = None
    return pick_a, pick_b


def _find_para(doc, predicate, start=0, end=None):
    paras = doc.paragraphs
    end = len(paras) if end is None else min(end, len(paras))
    for i in range(max(start, 0), end):
        if predicate(paras[i]):
            return i, paras[i]
    return None, None


def _fonts_ok(doc, para, ea_keys, size, tol=_SZ_TOL):
    """段落所有可见 run 的中文字体命中 ea_keys 之一且字号在容差内。"""
    runs = _visible_runs(para)
    if not runs:
        return False
    for run in runs:
        name = _eastasia(doc, run, para) or ""
        if ea_keys and not any(k in name for k in ea_keys):
            return False
        if size is not None and not _close(_size_pt(doc, run, para), size, tol):
            return False
    return True

def _sec_mm(section, attr):
    try:
        val = getattr(section, attr)
        return None if val is None else val.mm
    except Exception:
        return None


# ==================== 维度二：逐项判定 ====================

def _evaluate_impl(ctx):
    doc_a, doc_b = ctx["doc_a"], ctx["doc_b"]
    reg = ctx["regions_a"] or {}
    paras_a = doc_a.paragraphs if doc_a is not None else []
    paras_b = doc_b.paragraphs if doc_b is not None else []
    text_a = _doc_all_text(doc_a) if doc_a is not None else ""
    text_b = _doc_all_text(doc_b) if doc_b is not None else ""
    results = []

    def award(rule, ok, detail, delta):
        results.append("+%d：%s｜%s" % (delta if ok else 0, rule, detail))

    def penalize(rule, hit, detail, delta):
        results.append("-%d：%s｜%s" % (abs(delta) if hit else 0, rule, detail))

    # ---------- 01 纸张 A4 ----------
    # ① 呈报材料所有节的纸张一致；② 宽 210mm、高 297mm；③ 允许 ±1.2mm 换算误差。
    # 与办公软件对齐：Word「布局-纸张大小-A4」。
    rule = "呈报材料全文纸张统一设置为 A4（宽 210mm、高 297mm）"
    if doc_a is None:
        award(rule, False, "未定位到呈报材料文档", 3)
    else:
        bad = [i for i, s in enumerate(doc_a.sections)
               if not (_close(_sec_mm(s, "page_width"), 210.0, 1.2)
                       and _close(_sec_mm(s, "page_height"), 297.0, 1.2))]
        award(rule, not bad, "共 %d 节，非 A4 节序号：%s"
              % (len(doc_a.sections), bad if bad else "无"), 3)

    # ---------- 02 页边距 ----------
    # ① 左右各 25mm；② 上下各 20mm；③ 所有节一致，允许 ±1.2mm。
    # 与办公软件对齐：Word「布局-页边距-自定义边距」。
    rule = "页边距统一为左 25mm、右 25mm、上 20mm、下 20mm"
    if doc_a is None:
        award(rule, False, "未定位到呈报材料文档", 3)
    else:
        bad = []
        for i, s in enumerate(doc_a.sections):
            got = [_sec_mm(s, "left_margin"), _sec_mm(s, "right_margin"),
                   _sec_mm(s, "top_margin"), _sec_mm(s, "bottom_margin")]
            if not all(_close(g, t, 1.2) for g, t in zip(got, (25.0, 25.0, 20.0, 20.0))):
                bad.append((i, [None if g is None else round(g, 1) for g in got]))
        award(rule, not bad, "不合规节：%s" % (bad if bad else "无"), 3)

    # ---------- 03 页脚页码域 ----------
    # ① 页脚存在"第 X 页 共 Y 页"文字；② 数字由 PAGE / NUMPAGES 域生成而非手键；
    # ③ 页脚段落居中。与办公软件对齐：Word「插入-页码-页面底端」+ 域代码。
    rule = "页脚居中插入“第 X 页 共 Y 页”页码域（数字由域生成，非手工键入）"
    footer_paras = []
    if doc_a is not None:
        for ftr in _effective_footers(doc_a):
            footer_paras.extend(ftr.paragraphs)
    hit_p = None
    for para in footer_paras:
        codes = _field_codes(para._p)
        txt = _norm(para.text)
        if re.search(r"第\d*页共\d*页", txt) and "PAGE" in codes and "NUMPAGES" in codes:
            hit_p = para
            break
    ok = hit_p is not None and _align(doc_a, hit_p) == "center"
    award(rule, ok, "页脚段数=%d，命中页码域=%s，对齐=%s"
          % (len(footer_paras), hit_p is not None,
             _align(doc_a, hit_p) if hit_p is not None else "-"), 3)

    # ---------- 04 页脚字号 ----------
    # ① 页脚文字为小五号（9 磅）；② 居中；③ 以承载"第/页/共"字样的文字 run 为准，
    # 域结果数字继承正文字号不计入。与办公软件对齐：选中页脚后的字号框显示"小五"。
    rule = "页脚文字为小五号（9 磅）且居中"
    ok = False
    detail = "未找到页脚文字"
    if hit_p is not None:
        sizes = []
        for run in _visible_runs(hit_p):
            if _has_cjk(run.text):
                sizes.append(_size_pt(doc_a, run, hit_p))
        ok = bool(sizes) and all(_close(s, 9.0, _SZ_TOL) for s in sizes)
        ok = ok and _align(doc_a, hit_p) == "center"
        detail = "中文 run 字号=%s，对齐=%s" % (sizes, _align(doc_a, hit_p))
    award(rule, ok, detail, 1)

    # ---------- 封面区定位 ----------
    cover_end = reg.get("cover_end", -1)
    cover_paras = paras_a[:cover_end + 1] if cover_end >= 0 else []
    cover_text = "".join(_norm(p.text) for p in cover_paras)
    sep_paras = [p for p in cover_paras
                 if len(_norm(p.text)) >= 3 and set(_norm(p.text)) <= set("━─—-=")]
    cover_title = next((p for p in cover_paras
                        if _norm(p.text) == COVER_MAIN_TITLE), None)
    sep_ids = set(id(p._p) for p in sep_paras)
    other_cover = [p for p in cover_paras
                   if _norm(p.text) and id(p._p) not in sep_ids
                   and (cover_title is None or p._p is not cover_title._p)]

    # ---------- 05 封面单独成页且信息齐全 ----------
    # ① 封面自成一页，其后另起一页；② 主标题、括注、五条核查事项、呈报人及身份、
    # 联系方式、联系地址、日期、待核查金额等 14 项信息齐全。与办公软件对齐：
    # Word 中封面独占第 1 页；"自成一页"按 docx 内封面区后的分页符／节分隔判定。
    rule = "封面单独成页，并完整承载 14 项封面信息"
    missing = [m for m in COVER_ITEMS if m not in cover_text]
    brk = any(_break_inside(p) or _has_sect_break(p) for p in cover_paras)
    if not brk and 0 <= cover_end + 1 < len(paras_a):
        brk = _break_before(paras_a[cover_end + 1])
    award(rule, (not missing) and brk,
          "缺失项=%s，分页符=%s" % (missing or "无", brk), 5)

    # ---------- 06 封面主标题 ----------
    # ① 文字为“专项核查呈报材料”；② 黑体、二号（22 磅）、加粗；③ 居中。
    # 与办公软件对齐：字号框显示“二号”。
    rule = "封面主标题为黑体二号（22 磅）加粗居中"
    if cover_title is None:
        award(rule, False, "封面未找到主标题行", 3)
    else:
        ok = _fonts_ok(doc_a, cover_title, ("黑体",), 22.0, 0.6)
        ok = ok and all(_bold(doc_a, r, cover_title) for r in _visible_runs(cover_title))
        ok = ok and _align(doc_a, cover_title) == "center"
        award(rule, ok, "中文字体=%s 字号=%s 加粗=%s 对齐=%s"
              % (_eastasia(doc_a, _visible_runs(cover_title)[0], cover_title)
                 if _visible_runs(cover_title) else "-",
                 _para_size(doc_a, cover_title),
                 all(_bold(doc_a, r, cover_title) for r in _visible_runs(cover_title)),
                 _align(doc_a, cover_title)), 3)

    # ---------- 07 封面其余各行 ----------
    # ① 除主标题、分隔条外的封面各行；② 宋体、小三（15 磅）；③ 居中。
    rule = "封面其余各行为宋体小三（15 磅）居中"
    if not other_cover:
        award(rule, False, "封面无可判定的信息行", 1)
    else:
        bad = [_norm(p.text)[:12] for p in other_cover
               if not (_fonts_ok(doc_a, p, ("宋体",), 15.0)
                       and _align(doc_a, p) == "center")]
        award(rule, not bad, "共 %d 行，不合规：%s" % (len(other_cover), bad or "无"), 1)

    # ---------- 08 分隔条 ----------
    # ① 封面保留“━━━”分隔条；② 字号 10 磅；③ 居中。
    rule = "封面“━━━”分隔条为 10 磅并居中"
    if not sep_paras:
        award(rule, False, "封面未找到分隔条", 1)
    else:
        bad = []
        for p in sep_paras:
            sizes = [_size_pt(doc_a, r, p) for r in _visible_runs(p)]
            if not (sizes and all(_close(s, 10.0, _SZ_TOL) for s in sizes)
                    and _align(doc_a, p) == "center"):
                bad.append((sizes, _align(doc_a, p)))
        award(rule, not bad, "共 %d 条，不合规：%s" % (len(sep_paras), bad or "无"), 1)

    # ---------- 09 未保留“---”占位符 ----------
    # ① 封面来源内容中的“---”属排版占位符；② 成稿不应出现；③ 逐段检查原文字符。
    rule = "封面未保留“---”占位符"
    raw_a = "".join(p.text or "" for p in paras_a)
    award(rule, doc_a is not None and "---" not in raw_a,
          "存在“---”=%s" % ("---" in raw_a), 1)

    # ---------- 目录区定位 ----------
    toc_title_idx = reg.get("toc_title")
    body_start = reg.get("body_start", 0)
    toc_entries = []
    if toc_title_idx is not None:
        for i in range(toc_title_idx + 1, min(body_start, len(paras_a))):
            if _norm(paras_a[i].text) and _is_toc_entry(paras_a[i]):
                toc_entries.append(paras_a[i])

    def _entry_title(para):
        raw = para.text or ""
        head = raw.split("\t")[0]
        return re.sub(r"[\.·…\s]*\d+$", "", _norm(head))

    def _entry_page(para):
        m = re.search(r"(\d+)\s*$", (para.text or "").strip())
        return int(m.group(1)) if m else None

    def _has_right_tab(para):
        def getter(ppr):
            tabs = ppr.find(qn("w:tabs"))
            if tabs is None:
                return None
            for tab in tabs.findall(qn("w:tab")):
                if tab.get(qn("w:val")) == "right":
                    return True
            return None
        return bool(_first(_ppr_sources(doc_a, para), getter))

    # ---------- 10 目录域 ----------
    # ① 目录由 TOC 域自动生成；② 条目页码为 PAGEREF 域，可 Ctrl+单击跳转；
    # ③ 手工键入的目录不得分。与办公软件对齐：Word「引用-目录-自动目录」。
    rule = "目录为可自动更新的 TOC 目录域（含页码域，可跳转）"
    codes_a = _field_codes(doc_a.element.body) if doc_a is not None else ""
    ok = bool(re.search(r"\bTOC\b", codes_a)) and "PAGEREF" in codes_a
    award(rule, ok, "含 TOC 域=%s，含 PAGEREF 域=%s"
          % (bool(re.search(r"\bTOC\b", codes_a)), "PAGEREF" in codes_a), 5)

    # ---------- 11 目录标题 ----------
    # ① 独立“目录”标题行；② 黑体、小三（15 磅）、加粗；③ 居中。
    rule = "目录标题“目录”为黑体小三（15 磅）加粗居中"
    if toc_title_idx is None:
        award(rule, False, "未找到“目录”标题行", 3)
    else:
        tp = paras_a[toc_title_idx]
        ok = _fonts_ok(doc_a, tp, ("黑体",), 15.0)
        ok = ok and all(_bold(doc_a, r, tp) for r in _visible_runs(tp))
        ok = ok and _align(doc_a, tp) == "center"
        award(rule, ok, "字号=%s 加粗=%s 对齐=%s"
              % (_para_size(doc_a, tp),
                 all(_bold(doc_a, r, tp) for r in _visible_runs(tp)),
                 _align(doc_a, tp)), 3)

    # ---------- 12 目录条目 ----------
    # ① 只罗列六个一级标题，不混入二级条目；② 条目文字与正文标题逐字一致；
    # ③ 页码靠右（右对齐制表位），且页码由 PAGEREF 域生成、自上而下单调不减——
    #    页码一律取自 docx 内的域结构与制表位设置，不另行导出 PDF 核对。
    rule = "目录仅罗列六个一级标题，文字与正文一致、页码靠右且与实际页码相符"
    titles = [_entry_title(p) for p in toc_entries]
    level_ok = titles == L1_TITLES_A
    tab_ok = bool(toc_entries) and all(_has_right_tab(p) for p in toc_entries)
    page_detail = "未判定"
    page_ok = False
    if level_ok:
        nums = [_entry_page(p) for p in toc_entries]
        has_ref = all("PAGEREF" in _field_codes(p._p) for p in toc_entries)
        page_ok = has_ref and all(n is not None for n in nums) \
            and all(nums[i] <= nums[i + 1] for i in range(len(nums) - 1))
        page_detail = "结构判定：PAGEREF=%s 页码=%s" % (has_ref, nums)
    award(rule, level_ok and tab_ok and page_ok,
          "条目=%d，层级/文字一致=%s，右对齐制表位=%s，%s"
          % (len(toc_entries), level_ok, tab_ok, page_detail), 3)

    # ---------- 13 目录独立成页 ----------
    # ① 目录前有分页，与封面分离；② 目录后另起一页开始正文。
    # 分页一律按 docx 内的分页符／节分隔判定，不另行导出 PDF 复核。
    rule = "目录单独成页，其后正文另起一页"
    after_brk = False
    if toc_title_idx is not None and toc_entries:
        for i in range(toc_title_idx, min(body_start, len(paras_a))):
            if _break_inside(paras_a[i]) or _has_sect_break(paras_a[i]):
                after_brk = True
        if not after_brk and 0 <= body_start < len(paras_a):
            after_brk = _break_before(paras_a[body_start])
    award(rule, brk and after_brk,
          "目录前分页=%s，目录后分页=%s" % (brk, after_brk), 1)

    # ---------- 标题体系定位 ----------
    def _style_name(para):
        try:
            return (para.style.name or "").lower()
        except Exception:
            return ""

    h1_paras, h2_paras = [], []
    for i in range(min(body_start, len(paras_a)), len(paras_a)):
        para = paras_a[i]
        txt = _norm(para.text)
        if not txt:
            continue
        name = _style_name(para)
        if name == "heading 1" or txt in L1_TITLES_A:
            h1_paras.append(para)
        elif name == "heading 2" or txt in L2_TITLES_A:
            h2_paras.append(para)

    # ---------- 14 一级标题文字 ----------
    # ① 六个一级标题改写为“第X部分：……”；② 序号与原章节顺序一一对应；
    # ③ 文字逐字一致。
    rule = "六个一级标题改写为“第一部分：项目背景”至“第六部分：综合说明”"
    got = [_norm(p.text) for p in h1_paras]
    award(rule, got == L1_TITLES_A, "实际一级标题=%s" % (got or "无"), 3)

    # ---------- 15 一级标题字体字号 ----------
    # ① 黑体；② 小三（15 磅）；③ 六个标题全部满足。
    rule = "一级标题为黑体小三（15 磅）"
    if not h1_paras:
        award(rule, False, "未找到一级标题", 3)
    else:
        bad = [_norm(p.text)[:10] for p in h1_paras
               if not _fonts_ok(doc_a, p, ("黑体",), 15.0)]
        award(rule, not bad, "不合规标题=%s" % (bad or "无"), 3)

    # ---------- 16 二级标题文字 ----------
    # ① 三个二级标题改为“一、/二、/三、”序号；② 保留原小节名称；③ 顺序一致。
    rule = "三个二级标题改为“一、提前付款情况”等中文数字序号形式"
    got2 = [_norm(p.text) for p in h2_paras]
    award(rule, got2 == L2_TITLES_A, "实际二级标题=%s" % (got2 or "无"), 3)

    # ---------- 17 二级标题字体字号与对齐 ----------
    # ① 黑体；② 四号（14 磅）；③ 左对齐（无 w:jc 亦按左对齐处理）。
    rule = "二级标题为黑体四号（14 磅）左对齐"
    if not h2_paras:
        award(rule, False, "未找到二级标题", 1)
    else:
        bad = [(_norm(p.text)[:10], _para_size(doc_a, p), _align(doc_a, p))
               for p in h2_paras
               if not (_fonts_ok(doc_a, p, ("黑体",), 14.0)
                       and _align(doc_a, p) in ("left", "both", "justify"))]
        award(rule, not bad, "不合规标题=%s" % (bad or "无"), 1)

    # ---------- 18 标题行距 ----------
    # ① 一、二级标题行距均为 1.5 倍；② 按样式继承后的生效值判定；③ 容差 ±0.05。
    rule = "一、二级标题行距为 1.5 倍"
    heads = h1_paras + h2_paras
    if not heads:
        award(rule, False, "未找到一、二级标题", 1)
    else:
        bad = [(_norm(p.text)[:10], _line_multiple(doc_a, p)) for p in heads
               if not _close(_line_multiple(doc_a, p), 1.5, 0.05)]
        award(rule, not bad, "不合规标题=%s" % (bad or "无"), 1)

    # ---------- 19 标题取消蓝色 ----------
    # ① 一、二级标题不再使用蓝色 2D5D91；② 逐 run 解析生效颜色；
    # ③ 未显式设色（继承自动色）视为合规。
    rule = "一、二级标题不再使用蓝色（2D5D91）字体颜色"
    if not heads:
        award(rule, False, "未找到一、二级标题", 1)
    else:
        bad = [_norm(p.text)[:10] for p in heads
               for r in _visible_runs(p)
               if (_color(doc_a, r, p) or "") == FORBIDDEN_HEADING_COLOR]
        award(rule, not bad, "仍为蓝色的标题=%s" % (sorted(set(bad)) or "无"), 1)

    # ---------- 正文段落定位 ----------
    body_paras = []
    for i in range(min(body_start, len(paras_a)), len(paras_a)):
        para = paras_a[i]
        txt = _norm(para.text)
        if len(txt) < 12:
            continue
        if txt in L1_TITLES_A or txt in L2_TITLES_A:
            continue
        if _style_name(para).startswith("heading"):
            continue
        if re.match(r"^(第[一二三四五六]部分|[一二三四五六]、|（[一二三四五六]）)", txt):
            continue
        if FIG_CAPTION_KEY in txt:
            continue
        if any(txt.startswith(k) for k in ("说明人：", "联系方式：", "联系地址：", "提交日期：")):
            continue
        if len(txt) < 25 and _align(doc_a, para) == "center":
            continue
        body_paras.append(para)

    # ---------- 20 正文字体字号 ----------
    # ① 正文段落中文字体为宋体；② 字号小四（12 磅）；③ 全部正文段落一致。
    rule = "正文文字为宋体小四（12 磅）"
    if not body_paras:
        award(rule, False, "未定位到正文段落", 3)
    else:
        bad = [(_norm(p.text)[:10], _para_size(doc_a, p)) for p in body_paras
               if not _fonts_ok(doc_a, p, ("宋体",), 12.0)]
        award(rule, not bad, "共 %d 段，不合规=%s" % (len(body_paras), bad[:5] or "无"), 3)

    # ---------- 21 正文行距与段后间距 ----------
    # ① 行距 1.25 倍；② 段后间距 4 磅；③ 按继承后的生效值判定，容差 ±0.03 / ±0.6。
    rule = "正文行距 1.25 倍、段后间距 4 磅"
    if not body_paras:
        award(rule, False, "未定位到正文段落", 3)
    else:
        bad = [(_norm(p.text)[:8], _line_multiple(doc_a, p), _space_pt(doc_a, p, "after"))
               for p in body_paras
               if not (_close(_line_multiple(doc_a, p), 1.25, 0.03)
                       and _close(_space_pt(doc_a, p, "after"), 4.0, 0.6))]
        award(rule, not bad, "不合规=%s" % (bad[:5] or "无"), 3)

    # ---------- 22 正文首行缩进 ----------
    # ① 首行缩进 2 字符；② 小四字号下折算 24 磅；③ 容差 ±1.5 磅。
    rule = "正文首行缩进 2 字符（24 磅）"
    if not body_paras:
        award(rule, False, "未定位到正文段落", 1)
    else:
        bad = [(_norm(p.text)[:8], _first_line_pt(doc_a, p)) for p in body_paras
               if not _close(_first_line_pt(doc_a, p), 24.0, 1.5)]
        award(rule, not bad, "不合规=%s" % (bad[:5] or "无"), 1)

    # ---------- 23 图注格式 ----------
    # ① 图注单独成段；② 黑体、五号（10.5 磅）；③ 居中。
    rule = "图注为黑体五号（10.5 磅）居中"
    _ci, cap_para = _find_para(doc_a, lambda p: FIG_CAPTION_KEY in _norm(p.text)) \
        if doc_a is not None else (None, None)
    if cap_para is None:
        award(rule, False, "未找到图注段落", 1)
    else:
        ok = _fonts_ok(doc_a, cap_para, ("黑体",), 10.5) \
            and _align(doc_a, cap_para) == "center"
        award(rule, ok, "字号=%s 对齐=%s"
              % (_para_size(doc_a, cap_para), _align(doc_a, cap_para)), 1)

    # ---------- 24 表格文字 ----------
    # ① 四个数据表格单元格文字居中；② 字号与正文一致为小四（12 磅）；
    # ③ 表内空段不参与判定。与办公软件对齐：单元格「水平居中」。
    rule = "四个数据表格单元格文字统一居中、字号为小四（12 磅）"
    if doc_a is None or not doc_a.tables:
        award(rule, False, "未找到表格", 1)
    else:
        bad = []
        for ti, table in enumerate(doc_a.tables):
            for _cell, para in _cell_paragraphs(table):
                if not _visible_runs(para):
                    continue
                sizes = [_size_pt(doc_a, r, para) for r in _visible_runs(para)]
                if not (_align(doc_a, para) == "center"
                        and all(_close(s, 12.0, _SZ_TOL) for s in sizes)):
                    bad.append((ti, _norm(para.text)[:8], _align(doc_a, para), sizes[:2]))
        award(rule, not bad, "表数=%d，不合规单元格=%s"
              % (len(doc_a.tables), bad[:5] or "无"), 1)

    # ---------- 文档B 定位 ----------
    b_title = None
    b_title_idx = None
    for i, para in enumerate(paras_b):
        if _norm(para.text):
            b_title, b_title_idx = para, i
            break
    b_title_text = _norm(b_title.text) if b_title is not None else ""

    # ---------- 25 主题替换 ----------
    # ① 首行主标题替换为新主题，且“财务资金异常”“库存占用”“现金流风险”“专项核查材料”
    # 四段要素齐全；② 与目标主题逐字一致，无漏字、错字或语序改动；③ 以首个非空段为准。
    rule = "首行主标题替换为“华中区域零售项目财务资金异常、库存占用及现金流风险专项核查材料”"
    miss = [p for p in NEW_THEME_PARTS if p not in b_title_text]
    award(rule, (not miss) and NEW_THEME_B in b_title_text,
          "实际首行=%s，缺失要素=%s" % (b_title_text[:50] or "无", miss or "无"), 3)

    # ---------- 26 旧副标题行 ----------
    # ① 不再保留“专项核查报告”独立行；② 新主题已含材料属性；③ 逐段比对。
    rule = "不再保留“专项核查报告”独立行"
    left = [i for i, p in enumerate(paras_b)
            if _norm(p.text) == OLD_THEME_LINE_B and i != b_title_idx]
    award(rule, doc_b is not None and not left, "残留段落序号=%s" % (left or "无"), 1)

    # ---------- 27 封面多余行 ----------
    # ① 封面来源文件的文件名标题行不属于封面信息；② 成稿保留即为多余内容；
    # ③ 在封面区与全文中一并检查。
    rule = "封面出现多余的“华中区域零售项目经营异常核查封面”行"
    hit = "华中区域零售项目经营异常核查封面" in (cover_text + text_a)
    penalize(rule, hit, "存在=%s" % hit, -3)

    # ---------- 28 说明信息四行 ----------
    # ① 正文起始处的说明人、联系方式、联系地址、提交日期四行属既有内容；
    # ② 任一行缺失即扣；③ 只在正文起始区域内比对，避免与封面同名信息混淆。
    rule = "正文起始处的说明人、联系方式、联系地址、提交日期四行缺失"
    intro_zone = [p for p in paras_a[min(body_start, len(paras_a)):] if _norm(p.text)][:12]
    miss_intro = []
    for line in INTRO_LINES_A:
        if not any(_norm(p.text).startswith(line) for p in intro_zone):
            miss_intro.append(line)
    penalize(rule, bool(miss_intro) and doc_a is not None,
             "缺失行=%s" % (miss_intro or "无"), -3)

    # ---------- 29 表格框线 ----------
    # ① 四个表格应保留网格框线；② 上下左右与内部横竖线齐全；
    # ③ 直接 tblBorders 或表格样式提供均可。
    rule = "表格网格框线丢失"
    lost = [i for i, t in enumerate(doc_a.tables)
            if not _table_grid_ok(t)] if doc_a is not None else []
    penalize(rule, bool(lost), "无框线表序号=%s" % (lost or "无"), -3)

    # ---------- 30 单元格多余空文本段 ----------
    # ① 单元格内新增空段落；② 或在文字 run 旁插入空 run；③ 任一处出现即扣。
    rule = "表格单元格内被插入多余空文本段"
    anomaly = []
    if doc_a is not None:
        for ti, table in enumerate(doc_a.tables):
            for row in table.rows:
                for cell in row.cells:
                    if len(cell.paragraphs) > 1 and any(
                            not _norm(p.text) for p in cell.paragraphs):
                        anomaly.append((ti, "空段", _norm(cell.text)[:8]))
                    for para in cell.paragraphs:
                        runs = para.runs
                        if len(runs) > 1 and any(r.text == "" for r in runs) \
                                and any(r.text for r in runs):
                            anomaly.append((ti, "空run", _norm(cell.text)[:8]))
    penalize(rule, bool(anomaly), "异常=%s" % (anomaly[:5] or "无"), -1)

    # ---------- 31 表头行 ----------
    # ① 表头文字保持加粗；② 表头单元格保持居中；③ 任一表头行不合规即扣。
    rule = "表格表头行丢失加粗或居中"
    bad_head = []
    if doc_a is not None:
        for ti, table in enumerate(doc_a.tables):
            if not table.rows:
                continue
            for cell in table.rows[0].cells:
                for para in cell.paragraphs:
                    runs = _visible_runs(para)
                    if not runs:
                        continue
                    if not all(_bold(doc_a, r, para) for r in runs) \
                            or _align(doc_a, para) != "center":
                        bad_head.append((ti, _norm(para.text)[:8]))
    penalize(rule, bool(bad_head), "不合规表头=%s" % (bad_head[:5] or "无"), -1)

    # ---------- 32 图注文字 ----------
    # ① 图注文字应与既有内容逐字一致；② 冒号、序号格式不得改写；
    # ③ 图注整体缺失由“删除既有内容”类目处理，此处只判被改写。
    rule = "图注文字被改写"
    cap_raw = _norm(cap_para.text) if cap_para is not None else ""
    penalize(rule, cap_para is not None and cap_raw != FIG_CAPTION_A,
             "实际图注=%s" % (cap_raw or "无"), -1)

    # ---------- 文末签署两行定位 ----------
    sign_a = []
    for para in paras_a[-8:]:
        txt = _norm(para.text)
        if txt.startswith("说明人：林远芝") or re.match(r"^2026年0?5月20日$", txt):
            sign_a.append(para)

    # ---------- 33 签署右对齐 ----------
    # ① 文末说明人与日期两行原为右对齐；② 变为左对齐/居中即扣；
    # ③ 两行整体缺失归入删除既有内容类目。
    rule = "文末签署两行丢失右对齐"
    bad_sign = [(_norm(p.text)[:10], _align(doc_a, p)) for p in sign_a
                if _align(doc_a, p) != "right"]
    penalize(rule, bool(bad_sign), "不合规=%s" % (bad_sign or "无"), -1)

    # ---------- 34 签署段落间距 ----------
    # ① 签署两行不应被添加首行缩进；② 也不应被添加段后间距；
    # ③ 只看段落自身直接格式，继承正文样式的既有设定不计。
    rule = "文末签署两行被添加首行缩进或段后间距"
    bad_sign2 = []
    for para in sign_a:
        fi = _direct_val(para, _g_ind("firstLine")) or _direct_val(para, _g_ind("firstLineChars"))
        sa = _direct_val(para, _g_spacing("after"))
        if (fi or 0) > 0 or (sa or 0) > 0:
            bad_sign2.append((_norm(para.text)[:10], fi, sa))
    penalize(rule, bool(bad_sign2), "被改动=%s" % (bad_sign2 or "无"), -1)

    # ---------- 文档B 结构定位 ----------
    def _startswith_any(text, keys):
        return any(text.startswith(k) for k in keys)

    b_l1 = [p for p in paras_b if _startswith_any(_norm(p.text), L1_TITLES_B)]
    b_l3 = [p for p in paras_b if _startswith_any(_norm(p.text), L3_TITLES_B)]
    b_sign = [p for p in paras_b
              if _startswith_any(_norm(p.text), ("核查编制人：", "编制日期："))]
    b_body = [p for p in paras_b
              if len(_norm(p.text)) >= 12 and p is not b_title
              and p not in b_l1 and p not in b_l3 and p not in b_sign]

    # ---------- 35 文档B 主标题版式 ----------
    # ① 主标题应保持黑体三号（16 磅）加粗居中；② 不得额外添加行距、段前段后间距；
    # ③ 只看直接格式判定“是否被额外添加”。
    rule = "主题文档主标题版式被改动（非黑体三号加粗居中，或被添加行距与段前段后间距）"
    if b_title is None:
        penalize(rule, False, "未定位到主标题", -1)
    else:
        fmt_ok = _fonts_ok(doc_b, b_title, ("黑体",), 16.0, 0.6) \
            and all(_bold(doc_b, r, b_title) for r in _visible_runs(b_title)) \
            and _align(doc_b, b_title) == "center"
        extra = [_direct_val(b_title, _g_spacing("line")),
                 _direct_val(b_title, _g_spacing("before")),
                 _direct_val(b_title, _g_spacing("after"))]
        penalize(rule, not fmt_ok or any(v for v in extra),
                 "字号=%s 加粗=%s 对齐=%s 直接行距/段前/段后=%s"
                 % (_para_size(doc_b, b_title),
                    all(_bold(doc_b, r, b_title) for r in _visible_runs(b_title)),
                    _align(doc_b, b_title), extra), -1)

    # ---------- 36 文档B 八个一级标题 ----------
    # ① 八个一级标题原为四号（14 磅）加粗；② 字号或加粗任一不满足即扣；
    # ③ 只判定仍然存在的标题，整体缺失归入删除既有章节类目。
    rule = "主题文档八个一级标题不满足四号（14 磅）加粗"
    bad_b1 = []
    for para in b_l1:
        runs = [r for r in _visible_runs(para) if _has_cjk(r.text)]
        if not runs:
            continue
        if not all(_close(_size_pt(doc_b, r, para), 14.0, _SZ_TOL)
                   and _bold(doc_b, r, para) for r in runs):
            bad_b1.append((_norm(para.text)[:10], _para_size(doc_b, para)))
    penalize(rule, bool(bad_b1), "共 %d 个标题，不合规=%s"
             % (len(b_l1), bad_b1[:5] or "无"), -3)

    # ---------- 37 文档B 签署两行 ----------
    # ① 核查编制人、编制日期两行原为四号加粗右对齐；② 中文 run 为判定依据；
    # ③ 数字沿用既有西文字号不计。
    rule = "主题文档核查编制人、编制日期两行不满足四号加粗右对齐"
    bad_b2 = []
    for para in b_sign:
        runs = [r for r in _visible_runs(para) if _has_cjk(r.text)]
        ok = bool(runs) and all(_close(_size_pt(doc_b, r, para), 14.0, _SZ_TOL)
                                and _bold(doc_b, r, para) for r in runs)
        if not ok or _align(doc_b, para) != "right":
            bad_b2.append((_norm(para.text)[:10], _para_size(doc_b, para),
                           _align(doc_b, para)))
    penalize(rule, bool(bad_b2), "不合规=%s" % (bad_b2 or "无"), -1)

    # ---------- 38 西文字体被批量覆盖 ----------
    # ① 统计显式指定 ascii=Times New Roman 且含中文的 run 占全部中文 run 的比例；
    # ② 超过半数视为批量覆盖；③ 仅统计 run 级直接设置，既有零散设定不计。
    rule = "主题文档西文字体被批量覆盖为 Times New Roman"
    cnt = 0
    cjk_runs = 0
    if doc_b is not None:
        pool = list(paras_b)
        for table in doc_b.tables:
            pool.extend(p for _c, p in _cell_paragraphs(table))
        for para in pool:
            for run in para.runs:
                if not _has_cjk(run.text):
                    continue
                cjk_runs += 1
                rpr = run._element.find(qn("w:rPr"))
                if rpr is None:
                    continue
                name = _g_font("ascii")(rpr) or ""
                if "Times New Roman" in name:
                    cnt += 1
    penalize(rule, cjk_runs >= 20 and cnt > cjk_runs * 0.5,
             "中文 run=%d，其中显式 Times New Roman=%d" % (cjk_runs, cnt), -1)

    # ---------- 39 文档B 行距与缩进 ----------
    # ① 正文行距 1.5 倍；② 首行缩进为 24 磅（一级标题 28.1 磅）；
    # ③ 只在设定值被改成其他数值时扣分，未设缩进的既有段落按原样放行。
    rule = "主题文档正文行距与首行缩进设定被改动"
    bad_b3 = []
    for para in b_body + b_l1:
        ls = _line_multiple(doc_b, para)
        fi = _first_line_pt(doc_b, para)
        if ls is not None and abs(ls - 1.5) > 0.05:
            bad_b3.append((_norm(para.text)[:8], "行距", ls))
        elif fi is not None and not (_close(fi, 24.0, 1.5) or _close(fi, 28.1, 1.5)):
            bad_b3.append((_norm(para.text)[:8], "缩进", fi))
    penalize(rule, bool(bad_b3), "偏离段落=%s" % (bad_b3[:5] or "无"), -1)

    # ---------- 40 主标题下方空白 ----------
    # ① 主标题下方原有一个空段落；② 出现两个及以上连续空段即为多余空白；
    # ③ 单个空段被设置超过 12 磅的段后间距同样视为多余空白。
    rule = "主标题下方出现多余空段落或多余空白间距"
    empties = []
    if b_title_idx is not None:
        for i in range(b_title_idx + 1, len(paras_b)):
            if _norm(paras_b[i].text):
                break
            empties.append(paras_b[i])
    extra_gap = len(empties) >= 2
    if len(empties) == 1:
        sa = _direct_val(empties[0], _g_spacing("after")) or 0
        extra_gap = sa / 20.0 > 12.0
    penalize(rule, extra_gap, "主标题下连续空段数=%d" % len(empties), -1)

    # ---------- 41 正文文字完整性 ----------
    # ① 两份文档的正文语句锚点应逐条保留；② 任一锚点缺失说明被删除或改写；
    # ③ 锚点取自各章节的特征语句，不含指令要求改写的标题与主题。
    rule = "删除或改写既有正文文字内容"
    miss_anchor = [a[:14] for a in BODY_ANCHORS_A if a not in text_a] if doc_a is not None else []
    miss_anchor += [a[:14] for a in BODY_ANCHORS_B if a not in text_b] if doc_b is not None else []
    penalize(rule, bool(miss_anchor), "缺失语句=%s" % (miss_anchor[:5] or "无"), -3)

    # ---------- 42 表格与插图 ----------
    # ① 文档A 应保留 4 个表格且行列规模不变；② 应保留 1 张示意图；
    # ③ 文档B 应保留 3 个表格且规模不变。
    rule = "删除或破坏既有表格、插图"
    shape_a = _table_shapes(doc_a) if doc_a is not None else []
    shape_b = _table_shapes(doc_b) if doc_b is not None else []
    pic_ok = doc_a is not None and _has_drawing(doc_a.element.body)
    hit44 = (doc_a is not None and (sorted(shape_a) != sorted(TABLE_SHAPES_A) or not pic_ok)) \
        or (doc_b is not None and sorted(shape_b) != sorted(TABLE_SHAPES_B))
    penalize(rule, hit44, "A表格=%s 插图=%s B表格=%s" % (shape_a, pic_ok, shape_b), -3)

    # ---------- 43 段落与章节完整性 ----------
    # ① 两份文档的章节名称应全部保留；② 文档A 每个章节标题下至少保留一段正文；
    # ③ 章节名称按“去序号后的核心文字”比对，不受指令要求的序号改写影响。
    rule = "删除既有段落或章节"
    core_a = [re.sub(r"^(第[一二三四五六]部分[：:]|[一二三]、)", "", t)
              for t in L1_TITLES_A + L2_TITLES_A]
    miss_sec = [t for t in core_a if t not in text_a] if doc_a is not None else []
    core_b = [re.sub(r"^[一二三四五六七八]、", "", t) for t in L1_TITLES_B] \
        + [re.sub(r"^（[一二三四五六]）", "", t) for t in L3_TITLES_B]
    miss_sec += [t for t in core_b if t not in text_b] if doc_b is not None else []
    empty_chapter = []
    if doc_a is not None and (h1_paras or h2_paras):
        head_ids = sorted(paras_a.index(p) for p in h1_paras + h2_paras)
        for k, idx in enumerate(head_ids):
            end = head_ids[k + 1] if k + 1 < len(head_ids) else len(paras_a)
            if not any(len(_norm(paras_a[j].text)) >= 25 for j in range(idx + 1, end)):
                empty_chapter.append(_norm(paras_a[idx].text)[:10])
    penalize(rule, bool(miss_sec) or bool(empty_chapter),
             "缺失章节=%s，无正文章节=%s" % (miss_sec[:5] or "无", empty_chapter or "无"), -3)

    # ---------- 44 姓名/联系方式/金额/日期数值 ----------
    # ① 宿主文字仍在时，对应的姓名、联系方式、金额、日期必须原样保留；
    # ② 表格内金额按所在行比对；③ 宿主整体缺失时不在本项重复扣分。
    rule = "改动姓名、联系方式、金额、日期等数值"
    bad_val = []
    if doc_a is not None:
        for host, token in VALUE_PAIRS_A:
            if host in text_a and token not in text_a:
                bad_val.append((host, token))
        rows_a = _doc_table_texts(doc_a)
        for host, token in VALUE_ROWS_A:
            rows = [r for r in rows_a if host in r]
            if rows and not any(token in r for r in rows):
                bad_val.append((host, token))
    if doc_b is not None:
        for host, token in VALUE_PAIRS_B:
            if host in text_b and token not in text_b:
                bad_val.append((host, token))
    penalize(rule, bool(bad_val), "被改动=%s" % (bad_val[:5] or "无"), -1)

    return results

# ==================== 维度一：可用性 / 可修改性 ====================

def _check_dim1(pairs):
    """维度一：两份文档能正常打开、结构完整、可直接编辑；版式混乱则完成度记 0。"""
    for role, doc in pairs:
        if doc is None:
            return False, "未定位到%s文档，无法评估" % role
        paras = [p for p in doc.paragraphs if _norm(p.text)]
        if len(paras) < 15:
            return False, "%s有效段落仅 %d 段，正文内容缺失或已被转为图片" % (role, len(paras))
        if len(re.findall(r"[一-鿿]", _doc_all_text(doc))) < 1500:
            return False, "%s可编辑中文字符过少，疑似内容丢失或整页位图" % role
        for ti, table in enumerate(doc.tables):
            if len(table.rows) < 2:
                return False, "%s第 %d 个表格结构异常（仅 %d 行）" % (role, ti + 1, len(table.rows))
            widths = set(len(r.cells) for r in table.rows)
            if len(widths) > 1:
                return False, "%s第 %d 个表格错行，各行单元格数不一致：%s" % (role, ti + 1, sorted(widths))
    return True, "两份文档均可正常打开，段落与表格结构完整、文字可编辑，未出现版式错乱"


def _scan_docx(dir_path):
    """仅在给定目录本身定位被评估的 .docx（不进入子目录），忽略 ~$ 临时文件。"""
    try:
        names = os.listdir(dir_path)
    except OSError:
        return []
    return [os.path.join(dir_path, n) for n in sorted(names)
            if n.lower().endswith(".docx") and not n.startswith("~$")
            and os.path.isfile(os.path.join(dir_path, n))]

# ==================== 对外接口 ====================

def evaluate(dir_path):
    result = {
        "id": SCRIPT_ID,
        "file_name": "",
        "status": "ok",
        "error": "",
        "dim1_pass": False,
        "dim1_reason": "",
        "dim2_items": [],
        "total_score": 0,
        "max_score": MAX_SCORE_TOTAL,
    }
    try:
        if not os.path.isdir(dir_path):
            raise FileNotFoundError("目录不存在或不是目录：%s" % dir_path)
        candidates = _scan_docx(dir_path)
        if not candidates:
            raise FileNotFoundError("目录中未找到 .docx 文件：%s" % dir_path)
        pick_a, pick_b = _classify(candidates)
        names = [os.path.basename(p["path"]) for p in (pick_a, pick_b) if p]
        result["file_name"] = " | ".join(names)
        doc_a = pick_a["doc"] if pick_a else None
        doc_b = pick_b["doc"] if pick_b else None
        path_a = pick_a["path"] if pick_a else ""
        path_b = pick_b["path"] if pick_b else ""

        dim1_pass, dim1_reason = _check_dim1([("呈报材料", doc_a), ("主题核查材料", doc_b)])
        result["dim1_pass"] = dim1_pass
        result["dim1_reason"] = dim1_reason
        if not dim1_pass:
            return result

        ctx = {
            "doc_a": doc_a, "doc_b": doc_b,
            "path_a": path_a, "path_b": path_b,
            "regions_a": _regions_a(doc_a) if doc_a is not None else {},
        }
        results = _evaluate_impl(ctx)
        if len(results) != len(_MAX_DELTAS_ORDER):
            raise RuntimeError("dim2 项数不一致：results=%d vs order=%d"
                               % (len(results), len(_MAX_DELTAS_ORDER)))
        total = 0
        for text, max_delta in zip(results, _MAX_DELTAS_ORDER):
            m = re.match(r"^([+-])(\d+)：(.*)$", text, re.S)
            if m is None:
                raise RuntimeError("dim2 结果格式异常：%s" % text[:40])
            delta = int(m.group(2)) * (1 if m.group(1) == "+" else -1)
            if delta not in (0, max_delta):
                raise RuntimeError("dim2 分值与规则不符：%s（max_delta=%d）"
                                   % (text[:40], max_delta))
            rule, _sep, _detail = m.group(3).partition("｜")
            result["dim2_items"].append({
                "rule": rule,
                "max_delta": max_delta,
                "delta": delta,
                "hit": delta != 0,
            })
            total += delta
        result["total_score"] = max(0, total)
    except FileNotFoundError as exc:
        # 找不到被评估文件：只回传纯文本原因，不带异常类名；维度一根本没进入判定，
        # dim1_pass 与 dim1_reason 一律留空（空字符串，而非 False），满分归零
        result["status"] = "error"
        result["error"] = str(exc)
        result["file_name"] = ""
        result["dim1_pass"] = ""
        result["dim1_reason"] = ""
        result["dim2_items"] = []
        result["total_score"] = 0
        result["max_score"] = 0
    except Exception as exc:
        result["status"] = "error"
        result["error"] = "脚本运行异常：%s: %s" % (type(exc).__name__, exc)
        result["dim2_items"] = []
        result["total_score"] = 0
        result["dim1_pass"] = False
        if not result["dim1_reason"]:
            result["dim1_reason"] = "评估过程异常，无法完成判定"
    return result


if __name__ == "__main__":
    _target = sys.argv[1] if len(sys.argv) > 1 else os.path.dirname(os.path.abspath(__file__))
    try:
        _out = evaluate(_target)
    except Exception:
        _out = {"id": SCRIPT_ID, "file_name": "", "status": "error",
                "error": "脚本运行异常：%s" % traceback.format_exc(),
                "dim1_pass": False, "dim1_reason": "", "dim2_items": [],
                "total_score": 0, "max_score": MAX_SCORE_TOTAL}
    print(json.dumps(_out, ensure_ascii=False, indent=2))
