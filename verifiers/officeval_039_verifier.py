#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# pyright: reportUnknownMemberType=false, reportUnknownVariableType=false, reportUnknownArgumentType=false, reportUnknownParameterType=false, reportMissingImports=false, reportAttributeAccessIssue=false, reportOptionalMemberAccess=false, reportIndexIssue=false, reportGeneralTypeIssues=false, reportPrivateUsage=false
"""
说明：按接口统一约定，本模块不修改全局 sys.stdout，也不 print 主结果。主结果统一由
      evaluate(dir_path) -> dict 返回，由调用方决定如何序列化 / 落盘 / 汇总。
      __main__ 分支仅为本地自测，打印 json 便于人工核对。

评测对象：《办公设备维护与巡检服务合同》重新排版为两页的 .docx 交付件。
评分口径：完全对齐同目录《细则.txt》，维度1 为一票否决（不可用即总分 0），
      维度2 共 37 项，其中 11 项加分（满分 43）、26 项减分（最多 -40）。
      加分项只奖励"指令要求的活确实干了"，凡不动手即成立的一致性/完整性条件均归入减分项，
      因此未做任何改动的原始文件在本表下得 0 分。
版面相关判定优先走 Word COM（ComputeStatistics/Information），
      COM 不可用时退化为 PDF + PyMuPDF 兜底，两者都不可用时该项判为未达成。
"""

import os
import re
import sys
import json
import math
import shutil
import zipfile
import tempfile
import traceback

from docx import Document
from docx.oxml.ns import qn

try:
    _VERIFIER_DIR = os.path.dirname(os.path.abspath(__file__))
    if _VERIFIER_DIR not in sys.path:
        sys.path.insert(0, _VERIFIER_DIR)
    import pymupdf_compat as fitz  # 已移除 PyMuPDF(AGPL)，改用 pdfplumber+pypdfium2 兼容层
except ImportError:
    fitz = None

SCRIPT_ID = "039"
MAX_SCORE_TOTAL = 43
MIN_SCORE_TOTAL = -40

# 与《细则.txt》维度2 逐行一一对应：正数为加分项满分，负数为减分项最大扣分
_MAX_DELTAS_ORDER = [
    5,   # A01 重新分页后总页数为 2 页
    5,   # A02 标题至乙方收款账户段全部落在第1页
    5,   # A03 分页断点恰在乙方收款账户段之后、"五、"之前
    3,   # A04 第2页完整承载"五、"至"十、"六条条款
    3,   # A05 签署信息表整体在第2页且起始位置过页高一半
    3,   # A06 服务范围表整体在第1页且不跨页
    5,   # A07 服务套餐表整体在第1页且不跨页
    5,   # A08 全文压缩为同一字号且不大于 7 磅
    5,   # A09 全文正文段落统一行距且不大于单倍行距
    3,   # A10 表格行最小高度随字号同步下调
    1,   # A11 两页版面有效利用且无空白页
    -5,  # B01 删条款 / 删表格行 / 删正文段落强行压页
    -3,  # B02 字号压到 6 磅以下或行距收紧后文字被裁切
    -3,  # B03 页面设置（A4 纵向 + 页边距 + 页眉页脚距离）
    -3,  # B04 甲乙方信息表"联系号码"行
    -1,  # B05 插入空段落 / 手动分页符分节符硬拆
    -1,  # B06 页眉文字 / 页脚域 / 页眉页脚字体字号颜色对齐
    -1,  # B07 甲乙方信息表信用代码 / 邮箱 / 地址 / 末行
    -1,  # B08 甲乙方信息表 7行4列及四列列宽
    -1,  # B09 服务范围表 5行4列、列宽 4.88cm、表头
    -1,  # B10 服务套餐表 6行5列及五个套餐名称
    -1,  # B11 服务套餐表 选择列 / 年度费用 / 适用设备规模
    -1,  # B12 签署信息表 5行2列且左右列宽相同
    -5,  # B13 四张表格居中、实线网格、表级内边距与逐格内边距
    -1,  # B14 合同标题字体加粗单行居中且字号不小于正文
    -1,  # B15 十条条款标题文字序号一致且加粗
    -1,  # B16 "六、"下两组各 5 个编号小项且小标题加粗
    -1,  # B17 合同总价段落完整表述
    -1,  # B18 乙方收款账户段落完整表述
    -1,  # B19 服务期限与标准服务时间表述
    -1,  # B20 签署信息表内容
    -1,  # B21 全文统一微软雅黑
    -1,  # B22 正文段落段前段后间距 0
    -1,  # B23 5 个说明性段落保留首行缩进、其余无首行缩进
    -1,  # B24 "十、违约责任"第 4 项与签署信息表之间 2 个空段落
    -1,  # B25 无新增水印 / 批注 / 修订 / 彩色底纹 / 边框 / 说明性文字
    -1,  # B26 节数为 1 且无首页不同、奇偶页不同页眉页脚
]

assert sum(d for d in _MAX_DELTAS_ORDER if d > 0) == MAX_SCORE_TOTAL
assert sum(d for d in _MAX_DELTAS_ORDER if d < 0) == MIN_SCORE_TOTAL
assert len(_MAX_DELTAS_ORDER) == 37

# ==================== 一、评测常量（全部来自细则的显式表述） ====================

EMU_PER_TWIP = 635.0
TWIP_PER_CM = 567.0

# python-docx 的 qn() 命名空间表不含 VML 前缀 v，这里直接给出全限定标签名
VML_SHAPE_TAG = "{urn:schemas-microsoft-com:vml}shape"

TITLE_TEXT = "办公设备维护与巡检服务合同"

CLAUSE_TITLES = [
    "一、合作原则与适用依据",
    "二、服务对象与项目范围",
    "三、服务期限与响应安排",
    "四、服务套餐与费用",
    "五、资料、权限与现场配合",
    "六、双方责任与义务",
    "七、验收、交付与服务记录",
    "八、费用调整与新增事项",
    "九、保密与信息安全",
    "十、违约责任",
]
# 第1页应承载"一、"至"四、"，第2页应承载"五、"至"十、"
CLAUSE_TITLES_PAGE1 = CLAUSE_TITLES[:4]
CLAUSE_TITLES_PAGE2 = CLAUSE_TITLES[4:]

SUB_TITLE_A = "（一）甲方责任与义务"
SUB_TITLE_B = "（二）乙方责任与义务"

# 各条款下的编号小项数量（"六、"含两个小标题 + 各 5 个小项 = 12 段）
CLAUSE_BODY_COUNTS = [4, 1, 4, 3, 4, 12, 3, 3, 3, 4]
NONEMPTY_TOPLEVEL_PARAS = 53
EMPTY_TOPLEVEL_PARAS_MAX = 3

PARA_INTRO = "甲乙双方基于真实业务需求，就办公设备日常巡检、故障排查、配件更换协同、使用指导及服务记录交付等事项达成本合同。"
PARA_SCOPE = "经双方确认，服务对象为甲方办公场所内已登记在册的办公终端、打印复印设备、会议显示设备、网络辅助设备及常用外设。"
PARA_PLAN = "甲方选择的服务内容以本表勾选项目、双方盖章确认的设备清单及后续书面补充为准。"
PARA_TOTAL = "合同总价：人民币壹万捌仟捌佰元整"
PARA_ACCOUNT = "乙方收款账户：户名：岚舟金辟科技有限公司"
# 细则要求保留首行缩进的 5 个说明性段落
INDENT_PARA_HEADS = [PARA_INTRO, PARA_SCOPE, PARA_PLAN, PARA_TOTAL, PARA_ACCOUNT]

# 合同总价段落必须完整包含的三段表述
TOTAL_FRAGMENTS = [
    "人民币壹万捌仟捌佰元整（¥18,800.00）",
    "签订后7个工作日内支付首期款人民币玖仟肆佰元整",
    "履行满6个月后10个工作日内支付尾款人民币玖仟肆佰元整",
]
ACCOUNT_FRAGMENT = "户名：岚舟金辟科技有限公司；开户行：星河商业银行杭州望潮支行；账号：889012647305552108。"
TERM_FRAGMENT = "自2026年6月1日起至2027年5月31日止"
WORKTIME_FRAGMENT = "标准服务时间为工作日9:00至18:00"

# 甲乙方信息表（T0）：细则显式钉住的行
T0_ROW_CREDIT = ["统一社会信用代码", "91420372MA7QX8K59P", "统一社会信用代码", "91330118MA9LQ3C74J"]
T0_ROW_PHONE = ["联系号码", "789-781-2639转806", "联系号码", "789-925-4186转217"]
T0_ROW_MAIL = ["电子邮箱", "service@yunlingsenzhong.cn", "电子邮箱", "support@lanzhoujinpi.cn"]
T0_ROW_ADDR = ["通讯地址", "成都市锦云路118号锦澜商务中心B座18层", "通讯地址", "杭州市望潮路62号数创园5幢6层"]
T0_ROW_LAST = ["项目对接人", "周明厘/运营保障部", "项目负责人", "林知夏/区域服务部"]

# 服务范围表（T1）
T1_HEADER = ["服务模块", "包含内容", "交付记录", "不包含事项"]

# 服务套餐表（T2）
T2_HEADER = ["选择", "套餐名称", "适用设备规模", "年度费用", "主要服务内容"]
T2_PACKAGES = ["基础巡检包", "标准维护包", "增强保障包", "驻场协同包", "单项补充服务"]
T2_CHOICES = ["□", "☑", "□", "□", "□"]
T2_FEES = ["¥8,600", "¥18,800", "¥36,500", "另行确认", "按工单确认"]
T2_SCALES = ["20台以内", "21-60台", "61-120台", "120台以上", "按需"]

# 签署信息表（T3）
T3_ROWS = [
    ["甲方（盖章）：云岭森钟科技有限公司", "乙方（盖章）：岚舟金辟科技有限公司"],
    ["授权代表：陆嘉迪", "授权代表：沈若来"],
    ["签署日期：2026年5月28日", "签署日期：2026年5月28日"],
    ["签署地点：浙江省杭州市上城区", "签署地点：浙江省杭州市上城区"],
    ["", ""],
]

# 四张表格的形状：(行数, 列数)
TABLE_SHAPES = [(7, 4), (5, 4), (6, 5), (5, 2)]
# 细则给出的列宽（cm）；None 表示细则未对该表列宽提出要求
TABLE_COL_CM = [
    [2.47, 7.42, 2.47, 7.42],
    [4.88, 4.88, 4.88, 4.88],
    None,
    [8.10, 8.10],
]
COL_W_TOL_TWIP = 24

# 页面设置（twips）
PG_W_TWIP = 11906          # 21cm
PG_H_TWIP = 16838          # 29.7cm
MAR_TB_TWIP = 198          # 0.35cm
MAR_LR_TWIP = 482          # 0.85cm
HDR_FTR_TWIP = 113         # 0.20cm
PG_TOL_TWIP = 6

# 表级默认单元格内边距：左右 0.19cm = 108 twips，上下 0cm
CELL_MAR_LR_TWIP = 108
CELL_MAR_TB_TWIP = 0
CELL_MAR_TOL_TWIP = 4

# 逐格单元格内边距（left, right, top, bottom，单位 twips），键为行号，"*" 为该表其余行；
# 甲乙方信息表 0.14cm；服务范围表表头 0.12/0.16cm、其余 0.12cm；
# 服务套餐表表头 0.12/0.16cm、其余 0.11/0.12cm；签署信息表前四行 0.14/0.16cm（末行不作要求）。
CELL_MAR_PER_TABLE = [
    {"*": (80, 80, 80, 80)},
    {0: (70, 70, 90, 90), "*": (70, 70, 70, 70)},
    {0: (70, 70, 90, 90), "*": (60, 60, 70, 70)},
    {0: (80, 80, 90, 90), 1: (80, 80, 90, 90), 2: (80, 80, 90, 90), 3: (80, 80, 90, 90)},
]

# 页眉页脚
HEADER_TEXT = "办公设备维护与巡检服务合同"
HDR_FTR_FONT = "微软雅黑"
HDR_FTR_SZ_HALFPT = 16     # 8 磅
HDR_FTR_COLOR = "666666"
FOOTER_TEXT_RE = re.compile(r"^第\d*页/共\d*页$")

# 允许存在的底纹填充（合同原有的表头与勾选底色），多一种或少一种都视为改动
SHD_FILL_EXPECT = {"EDEDED": 14, "F0F0F0": 5, "FFF8D6": 5, "D9EAF7": 4, "F4F9FC": 4}
FORBIDDEN_TEXTS = ["两页排版"]

BODY_MIN_FONT_PT = 6.0     # 正文字号下限
BODY_MAX_FONT_PT = 7.0     # 压缩后字号上限（细则要求不大于 7 磅）
LINE_SINGLE_TWIP = 240     # 单倍行距 = 240 twips = 12 磅，行距不得超过该值
TRH_LINE_FACTOR = 2.0      # 行最小高度 <= 2.0 × 字号 × 估算行数 + 余量
TRH_SLACK_PT = 10.0
PAGE1_FILL_RATIO = 0.75    # 第1页底部空白不超过页高 1/4
SIGN_TABLE_MIN_RATIO = 0.50  # 签署信息表起始位置须过页高一半

# ==================== 二、通用工具 ====================


def _norm(text):
    """归一化文本：去掉所有空白（含全角空格）后比较，避免排版空格干扰内容判定。"""
    if text is None:
        return ""
    return re.sub(r"[\s　 ]+", "", str(text))


def _emu_to_twip(value):
    if value is None:
        return None
    return int(round(float(value) / EMU_PER_TWIP))


def _twip_to_pt(value):
    if value is None:
        return None
    return float(value) / 20.0


def _cm_to_twip(value):
    return int(round(float(value) * TWIP_PER_CM))


def _int_or_none(value):
    try:
        return int(str(value).strip())
    except (TypeError, ValueError):
        return None


def _para_text(p_el):
    """直接从 w:t 取文，避免 itertext 把域代码 / 嵌套内容重复计入。"""
    return "".join(node.text or "" for node in p_el.iter(qn("w:t")))


def _iter_top_level_paragraphs(doc):
    """按 body 顺序返回顶层段落元素（不含表格内段落）。"""
    for child in doc.element.body.iterchildren():
        if child.tag == qn("w:p"):
            yield child


def _iter_body_children(doc):
    """按 body 顺序返回 ('p', el) / ('tbl', el) 序列。"""
    for child in doc.element.body.iterchildren():
        if child.tag == qn("w:p"):
            yield "p", child
        elif child.tag == qn("w:tbl"):
            yield "tbl", child


def _iter_all_paragraph_elements(doc):
    """全文段落（顶层 + 表格单元格内），用于字体 / 裁切类全量校验。"""
    for p_el in doc.element.body.iter(qn("w:p")):
        yield p_el


def _iter_all_runs(doc):
    for r_el in doc.element.body.iter(qn("w:r")):
        yield r_el


def _style_map(doc):
    """styleId -> w:style 元素。"""
    out = {}
    for st in doc.styles.element.findall(qn("w:style")):
        sid = st.get(qn("w:styleId"))
        if sid:
            out[sid] = st
    return out


def _default_paragraph_style_id(doc):
    for st in doc.styles.element.findall(qn("w:style")):
        if st.get(qn("w:type")) == "paragraph" and st.get(qn("w:default")) == "1":
            return st.get(qn("w:styleId"))
    return None


def _paragraph_of_run(r_el):
    node = r_el.getparent()
    while node is not None and node.tag != qn("w:p"):
        node = node.getparent()
    return node


def _para_style_id(p_el):
    if p_el is None:
        return None
    pPr = p_el.find(qn("w:pPr"))
    if pPr is None:
        return None
    st = pPr.find(qn("w:pStyle"))
    return st.get(qn("w:val")) if st is not None else None


def _style_chain(styles, style_id, limit=12):
    """沿 basedOn 向上展开样式继承链。"""
    chain = []
    seen = set()
    sid = style_id
    while sid and sid not in seen and len(chain) < limit:
        seen.add(sid)
        st = styles.get(sid)
        if st is None:
            break
        chain.append(st)
        based = st.find(qn("w:basedOn"))
        sid = based.get(qn("w:val")) if based is not None else None
    return chain


class _FontResolver(object):
    """按 run rPr -> 段落样式继承链 -> 默认段落样式 -> docDefaults 的顺序解析生效属性。

    与办公软件对齐：文档里有 15 个 run 只写了 w:rFonts w:hint 而没写具体字体名，
    这些 run 的实际字体来自默认段落样式，若只看 run 级属性会误判为"未统一字体"。
    """

    def __init__(self, doc):
        self._styles = _style_map(doc)
        self._default_pstyle = _default_paragraph_style_id(doc)
        dd = doc.styles.element.find(qn("w:docDefaults"))
        self._doc_rpr = None
        self._doc_ppr = None
        if dd is not None:
            rpr_default = dd.find(qn("w:rPrDefault"))
            if rpr_default is not None:
                self._doc_rpr = rpr_default.find(qn("w:rPr"))
            ppr_default = dd.find(qn("w:pPrDefault"))
            if ppr_default is not None:
                self._doc_ppr = ppr_default.find(qn("w:pPr"))

    # ---------- 属性链 ----------

    def _rpr_chain(self, r_el):
        """run 生效 rPr 候选序列，优先级从高到低。"""
        out = []
        if r_el is not None:
            rPr = r_el.find(qn("w:rPr"))
            if rPr is not None:
                out.append(rPr)
        p_el = _paragraph_of_run(r_el) if r_el is not None else None
        style_ids = []
        sid = _para_style_id(p_el)
        if sid:
            style_ids.append(sid)
        if self._default_pstyle and self._default_pstyle not in style_ids:
            style_ids.append(self._default_pstyle)
        for one in style_ids:
            for st in _style_chain(self._styles, one):
                rPr = st.find(qn("w:rPr"))
                if rPr is not None:
                    out.append(rPr)
        if self._doc_rpr is not None:
            out.append(self._doc_rpr)
        return out

    def _ppr_chain(self, p_el):
        out = []
        if p_el is not None:
            pPr = p_el.find(qn("w:pPr"))
            if pPr is not None:
                out.append(pPr)
        style_ids = []
        sid = _para_style_id(p_el)
        if sid:
            style_ids.append(sid)
        if self._default_pstyle and self._default_pstyle not in style_ids:
            style_ids.append(self._default_pstyle)
        for one in style_ids:
            for st in _style_chain(self._styles, one):
                pPr = st.find(qn("w:pPr"))
                if pPr is not None:
                    out.append(pPr)
        if self._doc_ppr is not None:
            out.append(self._doc_ppr)
        return out

    # ---------- 字体 ----------

    def font_names(self, r_el):
        """返回 (eastAsia, ascii) 生效字体名，未解析到则为 None。"""
        east = ascii_ = None
        for rPr in self._rpr_chain(r_el):
            fonts = rPr.find(qn("w:rFonts"))
            if fonts is None:
                continue
            if east is None:
                east = fonts.get(qn("w:eastAsia"))
            if ascii_ is None:
                ascii_ = fonts.get(qn("w:ascii"))
            if east is not None and ascii_ is not None:
                break
        return east, ascii_

    def size_pt(self, r_el):
        for rPr in self._rpr_chain(r_el):
            sz = rPr.find(qn("w:sz"))
            if sz is not None:
                val = _int_or_none(sz.get(qn("w:val")))
                if val:
                    return val / 2.0
        return None

    def para_size_pt(self, p_el):
        """段落生效字号：取段内 run 的最大字号，无 run 时退化到段落标记。"""
        sizes = [self.size_pt(r) for r in p_el.iter(qn("w:r"))]
        sizes = [s for s in sizes if s]
        if sizes:
            return max(sizes)
        pPr = p_el.find(qn("w:pPr"))
        if pPr is not None:
            rPr = pPr.find(qn("w:rPr"))
            if rPr is not None:
                sz = rPr.find(qn("w:sz"))
                if sz is not None:
                    val = _int_or_none(sz.get(qn("w:val")))
                    if val:
                        return val / 2.0
        return None

    # ---------- 段落间距 ----------

    def spacing(self, p_el):
        """返回 (before, after, line, lineRule)，均为生效值（before/after 单位 twips）。"""
        before = after = line = rule = None
        for pPr in self._ppr_chain(p_el):
            sp = pPr.find(qn("w:spacing"))
            if sp is None:
                continue
            if before is None:
                before = _int_or_none(sp.get(qn("w:before")))
            if after is None:
                after = _int_or_none(sp.get(qn("w:after")))
            if line is None:
                line = _int_or_none(sp.get(qn("w:line")))
                if line is not None:
                    rule = sp.get(qn("w:lineRule")) or "auto"
            if None not in (before, after, line):
                break
        return before, after, line, rule

    def first_line_indent(self, p_el):
        """返回 (firstLine twips, firstLineChars)，生效值。"""
        first = chars = None
        for pPr in self._ppr_chain(p_el):
            ind = pPr.find(qn("w:ind"))
            if ind is None:
                continue
            if first is None:
                first = _int_or_none(ind.get(qn("w:firstLine")))
            if chars is None:
                chars = _int_or_none(ind.get(qn("w:firstLineChars")))
            if first is not None and chars is not None:
                break
        return first, chars

    def jc(self, p_el):
        for pPr in self._ppr_chain(p_el):
            j = pPr.find(qn("w:jc"))
            if j is not None:
                return j.get(qn("w:val"))
        return None

    def bold(self, r_el):
        """run 是否加粗；w:b 无 val 视为 True，val 为 0/false 视为 False。"""
        for rPr in self._rpr_chain(r_el):
            b = rPr.find(qn("w:b"))
            if b is not None:
                val = b.get(qn("w:val"))
                if val is None:
                    return True
                return str(val).lower() not in ("0", "false", "off")
        return False


def _para_is_bold(resolver, p_el):
    """段落所有含文字的 run 均加粗才算加粗。"""
    runs = [r for r in p_el.iter(qn("w:r")) if _para_text(r).strip()]
    if not runs:
        return False
    return all(resolver.bold(r) for r in runs)


def _line_pitch_pt(font_pt, line, rule):
    """把 w:spacing 换算成一行占用的实际高度（磅）。"""
    natural = (font_pt or 0.0) * 1.32
    if line is None:
        return natural
    if rule == "exact":
        return _twip_to_pt(line)
    if rule == "atLeast":
        return max(_twip_to_pt(line), natural)
    return natural * (line / 240.0)


def _text_width_pt(text, font_pt):
    """粗估文本宽度：CJK / 全角按 1 个字宽，其余按 0.5 个字宽。"""
    units = 0.0
    for ch in text:
        code = ord(ch)
        if code > 0x2E00 or 0xFF00 <= code <= 0xFFEF:
            units += 1.0
        elif ch.strip() == "":
            units += 0.5
        else:
            units += 0.5
    return units * (font_pt or 0.0)


def _estimate_lines(text, avail_pt, font_pt):
    if not text.strip():
        return 1
    if avail_pt is None or avail_pt <= 0:
        return 1
    return max(1, int(math.ceil(_text_width_pt(text, font_pt) / avail_pt)))


# ==================== 三、表格 / 节 / 页眉页脚 取值 ====================


def _cell_texts(table):
    return [[_norm(c.text) for c in row.cells] for row in table.rows]


def _tbl_pr(table):
    return table._tbl.find(qn("w:tblPr"))


def _table_jc(table):
    tblPr = _tbl_pr(table)
    if tblPr is None:
        return None
    j = tblPr.find(qn("w:jc"))
    return j.get(qn("w:val")) if j is not None else None


def _table_borders(table):
    """返回 {edge: (val, sz, color)}，缺失边为 None。"""
    tblPr = _tbl_pr(table)
    out = {}
    bd = tblPr.find(qn("w:tblBorders")) if tblPr is not None else None
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        el = bd.find(qn("w:" + edge)) if bd is not None else None
        if el is None:
            out[edge] = None
        else:
            out[edge] = (el.get(qn("w:val")), el.get(qn("w:sz")), el.get(qn("w:color")))
    return out


def _table_cell_margins(table):
    """表级 w:tblCellMar，返回 {top,left,bottom,right} twips。"""
    tblPr = _tbl_pr(table)
    cm = tblPr.find(qn("w:tblCellMar")) if tblPr is not None else None
    out = {}
    for edge in ("top", "left", "bottom", "right"):
        el = cm.find(qn("w:" + edge)) if cm is not None else None
        out[edge] = _int_or_none(el.get(qn("w:w"))) if el is not None else None
    return out


def _cell_margins(tc):
    """单元格级 w:tcMar，缺失项返回 None（由表级默认接管）。"""
    tcPr = tc.find(qn("w:tcPr"))
    cm = tcPr.find(qn("w:tcMar")) if tcPr is not None else None
    out = {}
    for edge in ("top", "left", "bottom", "right"):
        el = cm.find(qn("w:" + edge)) if cm is not None else None
        out[edge] = _int_or_none(el.get(qn("w:w"))) if el is not None else None
    return out


def _effective_cell_margins(table, tc):
    base = _table_cell_margins(table)
    own = _cell_margins(tc)
    out = {}
    for edge in ("top", "left", "bottom", "right"):
        out[edge] = own[edge] if own[edge] is not None else base[edge]
    return out


def _row_cell_widths_twip(table, row_index=0):
    """取指定行各单元格 w:tcW（dxa），拿不到时退回 tblGrid。"""
    widths = []
    try:
        row = table.rows[row_index]
    except IndexError:
        return widths
    for cell in row.cells:
        tcPr = cell._tc.find(qn("w:tcPr"))
        tcW = tcPr.find(qn("w:tcW")) if tcPr is not None else None
        if tcW is not None and (tcW.get(qn("w:type")) or "dxa") == "dxa":
            widths.append(_int_or_none(tcW.get(qn("w:w"))))
        else:
            widths.append(None)
    if any(w is None for w in widths):
        grid = _grid_widths_twip(table)
        if len(grid) == len(widths):
            widths = [w if w is not None else grid[i] for i, w in enumerate(widths)]
    return widths


def _grid_widths_twip(table):
    grid = table._tbl.find(qn("w:tblGrid"))
    if grid is None:
        return []
    return [_int_or_none(g.get(qn("w:w"))) for g in grid.findall(qn("w:gridCol"))]


def _row_heights(table):
    """返回每行 (val twips, hRule)，未设置为 (None, None)。"""
    out = []
    for row in table.rows:
        trPr = row._tr.find(qn("w:trPr"))
        th = trPr.find(qn("w:trHeight")) if trPr is not None else None
        if th is None:
            out.append((None, None))
        else:
            out.append((_int_or_none(th.get(qn("w:val"))), th.get(qn("w:hRule"))))
    return out


def _section_metrics(doc):
    """节的页面几何，统一换算为 twips。"""
    sec = doc.sections[0]
    return {
        "pg_w": _emu_to_twip(sec.page_width),
        "pg_h": _emu_to_twip(sec.page_height),
        "top": _emu_to_twip(sec.top_margin),
        "bottom": _emu_to_twip(sec.bottom_margin),
        "left": _emu_to_twip(sec.left_margin),
        "right": _emu_to_twip(sec.right_margin),
        "header": _emu_to_twip(sec.header_distance),
        "footer": _emu_to_twip(sec.footer_distance),
        "orient": "portrait" if (sec.page_height or 0) >= (sec.page_width or 0) else "landscape",
    }


def _hdr_ftr_runs(part):
    """页眉/页脚里所有 run 元素。"""
    if part is None:
        return []
    return list(part._element.iter(qn("w:r")))


def _hdr_ftr_run_props_ok(runs):
    """页眉页脚 run 须为 微软雅黑 / 8 磅 / 非加粗 / #666666。"""
    if not runs:
        return False
    for r in runs:
        rPr = r.find(qn("w:rPr"))
        if rPr is None:
            return False
        fonts = rPr.find(qn("w:rFonts"))
        if fonts is None:
            return False
        for attr in ("w:ascii", "w:eastAsia", "w:hAnsi"):
            name = fonts.get(qn(attr))
            if name is not None and name != HDR_FTR_FONT:
                return False
        if fonts.get(qn("w:eastAsia")) != HDR_FTR_FONT:
            return False
        sz = rPr.find(qn("w:sz"))
        if sz is None or _int_or_none(sz.get(qn("w:val"))) != HDR_FTR_SZ_HALFPT:
            return False
        color = rPr.find(qn("w:color"))
        if color is None or (color.get(qn("w:val")) or "").upper() != HDR_FTR_COLOR:
            return False
        b = rPr.find(qn("w:b"))
        if b is not None and str(b.get(qn("w:val")) or "1").lower() not in ("0", "false", "off"):
            return False
    return True


def _hdr_ftr_centered(part):
    if part is None:
        return False
    paras = [p for p in part.paragraphs if p.text.strip()]
    if not paras:
        return False
    for p in paras:
        pPr = p._p.find(qn("w:pPr"))
        j = pPr.find(qn("w:jc")) if pPr is not None else None
        if j is None or j.get(qn("w:val")) != "center":
            return False
    return True


def _footer_field_codes(part):
    if part is None:
        return []
    out = []
    for node in part._element.iter(qn("w:instrText")):
        out.append((node.text or "").strip().upper())
    for node in part._element.iter(qn("w:fldSimple")):
        out.append((node.get(qn("w:instr")) or "").strip().upper())
    return out


def _open_word_document(file_path):
    """打开 docx；损坏 / 非 docx 会抛异常，交由维度1 判定。"""
    if not os.path.isfile(file_path):
        raise FileNotFoundError(f"文件不存在：{file_path}")
    if not zipfile.is_zipfile(file_path):
        raise ValueError("文件不是有效的 .docx（OPC 包）结构")
    return Document(file_path)


# ==================== 四、真实版面信息（Word COM 优先，PDF 兜底） ====================


class _Layout(object):
    """重新分页后的真实版面信息。

    available 为 False 时，所有依赖真实分页的条目按"未达成"处理，
    避免用静态 XML 猜测分页而给出错误结论。
    """

    def __init__(self):
        self.available = False
        self.source = "none"
        self.reason = ""
        self.pages_statistic = None      # ComputeStatistics(wdStatisticPages)
        self.pages_information = None    # Information(wdNumberOfPagesInDocument)
        self.page_height_pt = None
        self.para_pages = []             # 顶层段落 -> 页码
        self.para_tops = []              # 顶层段落 -> 距页顶距离（磅）
        self.table_pages = []            # 表格 -> (起始页, 结束页)
        self.table_tops = []             # 表格 -> 起始纵向位置（磅）
        self.page1_max_top = None        # 第1页最后一个段落的纵向位置
        self.last_para_page = None

    @property
    def page_count(self):
        for value in (self.pages_statistic, self.pages_information):
            if value:
                return int(value)
        return None


def _probe_layout_via_word(file_path):
    """用 Word 重新分页后读取真实页码与纵向位置。

    与办公软件对齐：ComputeStatistics(2) = wdStatisticPages，
    Information(3) = wdActiveEndPageNumber，Information(4) = wdNumberOfPagesInDocument，
    Information(6) = wdVerticalPositionRelativeToPage。
    """
    import win32com.client  # 延迟导入：无 pywin32 时走兜底分支

    layout = _Layout()
    tmp_dir = tempfile.mkdtemp(prefix="officeval039_")
    copied = os.path.join(tmp_dir, "target.docx")
    shutil.copy2(file_path, copied)
    app = None
    doc = None
    try:
        app = win32com.client.DispatchEx("Word.Application")
        app.Visible = False
        app.DisplayAlerts = 0
        doc = app.Documents.Open(copied, ReadOnly=False, AddToRecentFiles=False)
        doc.Repaginate()
        layout.page_height_pt = float(doc.PageSetup.PageHeight)
        layout.pages_statistic = int(doc.ComputeStatistics(2))
        layout.pages_information = int(doc.Content.Information(4))

        # Word 的 Paragraphs 集合含表格内段落，按 Range.Tables.Count 过滤出顶层段落
        total = int(doc.Paragraphs.Count)
        page1_max = 0.0
        for i in range(1, total + 1):
            para = doc.Paragraphs(i)
            rng = para.Range
            page = int(rng.Information(3))
            top = float(rng.Information(6))
            if page == 1 and top > page1_max:
                page1_max = top
            if int(rng.Tables.Count) == 0:
                layout.para_pages.append(page)
                layout.para_tops.append(top)
            layout.last_para_page = page
        layout.page1_max_top = page1_max

        for k in range(1, int(doc.Tables.Count) + 1):
            tbl = doc.Tables(k)
            start = int(tbl.Range.Information(3))
            end = int(tbl.Rows(tbl.Rows.Count).Range.Information(3))
            layout.table_pages.append((start, end))
            layout.table_tops.append(float(tbl.Range.Information(6)))

        layout.available = True
        layout.source = "word-com"
        return layout
    finally:
        try:
            if doc is not None:
                doc.Close(False)
        except Exception:
            pass
        try:
            if app is not None:
                app.Quit()
        except Exception:
            pass
        shutil.rmtree(tmp_dir, ignore_errors=True)


def _render_to_pdf(file_path, tmp_dir):
    """优先用 Word 导出 PDF（ExportFormat=17），失败则尝试 LibreOffice。"""
    pdf_path = os.path.join(tmp_dir, "render.pdf")
    copied = os.path.join(tmp_dir, "target.docx")
    if not os.path.exists(copied):
        shutil.copy2(file_path, copied)
    try:
        import win32com.client

        app = None
        doc = None
        try:
            app = win32com.client.DispatchEx("Word.Application")
            app.Visible = False
            app.DisplayAlerts = 0
            doc = app.Documents.Open(copied, ReadOnly=False, AddToRecentFiles=False)
            doc.ExportAsFixedFormat(OutputFileName=pdf_path, ExportFormat=17)
        finally:
            try:
                if doc is not None:
                    doc.Close(False)
            except Exception:
                pass
            try:
                if app is not None:
                    app.Quit()
            except Exception:
                pass
        if os.path.isfile(pdf_path):
            return pdf_path
    except Exception:
        pass

    soffice = shutil.which("soffice") or shutil.which("soffice.exe")
    if soffice:
        os.system('"%s" --headless --convert-to pdf --outdir "%s" "%s"' % (soffice, tmp_dir, copied))
        cand = os.path.join(tmp_dir, "target.pdf")
        if os.path.isfile(cand):
            return cand
    return None


def _probe_layout_via_pdf(doc, file_path):
    """兜底路径：渲染 PDF 后用文本定位各段落 / 表格所在页与纵向位置。"""
    if fitz is None:
        raise RuntimeError("PyMuPDF(fitz) 不可用")
    layout = _Layout()
    tmp_dir = tempfile.mkdtemp(prefix="officeval039pdf_")
    try:
        pdf_path = _render_to_pdf(file_path, tmp_dir)
        if not pdf_path:
            raise RuntimeError("无法渲染 PDF")
        with fitz.open(pdf_path) as pdf:
            layout.pages_statistic = pdf.page_count
            layout.pages_information = pdf.page_count
            layout.page_height_pt = float(pdf[0].rect.height)
            page_texts = [_norm(page.get_text()) for page in pdf]

            def locate(text, fallback=None):
                key = _norm(text)[:24]
                if not key:
                    return fallback
                for idx, body in enumerate(page_texts):
                    if key in body:
                        return idx + 1
                return fallback

            def locate_top(text, page_no):
                if not page_no:
                    return None
                probe = (text or "").strip()[:16]
                if not probe:
                    return None
                rects = pdf[page_no - 1].search_for(probe)
                return float(min(r.y0 for r in rects)) if rects else None

            page1_max = 0.0
            for p_el in _iter_top_level_paragraphs(doc):
                text = _para_text(p_el)
                page = locate(text)
                layout.para_pages.append(page)
                top = locate_top(text, page)
                layout.para_tops.append(top)
                if page == 1 and top is not None and top > page1_max:
                    page1_max = top
                if page:
                    layout.last_para_page = page
            for table in doc.tables:
                rows = _cell_texts(table)
                first = next((c for c in rows[0] if c), "") if rows else ""
                last = ""
                for cell in reversed(rows[-1] if rows else []):
                    if cell:
                        last = cell
                        break
                start = locate(first)
                end = locate(last, start)
                layout.table_pages.append((start, end))
                top = locate_top(first, start)
                layout.table_tops.append(top)
                if start == 1 and top is not None and top > page1_max:
                    page1_max = top
            layout.page1_max_top = page1_max or None
        layout.available = True
        layout.source = "pdf-fitz"
        return layout
    finally:
        shutil.rmtree(tmp_dir, ignore_errors=True)


def _probe_layout(doc, file_path):
    try:
        return _probe_layout_via_word(file_path)
    except Exception as exc_word:
        try:
            return _probe_layout_via_pdf(doc, file_path)
        except Exception as exc_pdf:
            layout = _Layout()
            layout.reason = "Word COM 不可用（%s）；PDF 兜底不可用（%s）" % (exc_word, exc_pdf)
            return layout


# ==================== 五、细则条目原文（顺序与 _MAX_DELTAS_ORDER 严格一致） ====================

RULES = [
    # ---- 加分项 11 条 ----
    '文档在 Word 中重新分页后总页数为 2 页（ComputeStatistics 与 Information 两种页数统计均为 2），既未压缩成 1 页也未超出 2 页',
    '合同标题、合同引言段、"一、合作原则与适用依据"至"四、服务套餐与费用"的全部条款正文、合同总价段落、乙方收款账户段落均完整落在第1页，无任何一项溢出到第2页',
    '分页断点恰好落在"乙方收款账户：户名：岚舟金辟科技有限公司；开户行：星河商业银行杭州望潮支行；账号：8890 1264 7305 5521 08。"之后、"五、资料、权限与现场配合"之前',
    '第2页完整承载"五、资料、权限与现场配合"至"十、违约责任"共六条条款，条款序号连续，无任何一条被挤回第1页',
    '签署信息表（含"甲方（盖章）""乙方（盖章）""授权代表""签署日期""签署地点"）整体位于第2页，且起始纵向位置落在页面下半部（大于页高的 50%）',
    '服务范围表（表头"服务模块""包含内容""交付记录""不包含事项"）整体位于第1页且不跨页',
    '服务套餐表（6行5列，表头"选择""套餐名称""适用设备规模""年度费用""主要服务内容"）整体位于第1页且不跨页，表格中间不出现断页',
    '全文正文与表格文字压缩为同一字号，该字号不大于 7 磅，且不出现两种及以上字号混排（如一部分段落 8 磅、一部分段落 8.5 磅）',
    '全文正文段落采用统一的行距设置，该行距不大于单倍行距（即 w:line 不超过 240 twips / 12 磅），且不出现多种行距值混排（如一部分固定值 12 磅、一部分固定值 9.8 磅）',
    '四张表格的行最小高度随字号同步下调，表格内不出现由过大行高造成的成片空白',
    '两页版面均被有效利用，第1页底部不出现超过页高 1/4 的连续空白，且文档最终不含空白页，第2页之后无多余空页',
    # ---- 减分项 26 条 ----
    '删除条款、删除表格行或删除正文段落来强行压页，十条条款、四张表格及全部正文段落未能完整保留',
    '正文字号被压到 6 磅以下，或行距收紧后文字被裁切（段落行高小于该段文字高度，出现上下笔画被切、相邻两行文字重叠，打印后无法正常阅读）',
    '页面设置不满足 A4 纵向（21cm×29.7cm）、上下页边距 0.35cm、左右页边距 0.85cm、页眉距顶端 0.20cm、页脚距底端 0.20cm',
    '甲乙方信息表"联系号码"行不满足甲方为"789-781-2639 转 806"、乙方为"789-925-4186 转 217"',
    '通过插入连续空段落或手动分页符/分节符把内容硬拆成两页',
    '页眉不满足文字为"办公设备维护与巡检服务合同"、页脚不满足由"第"+PAGE 域+"页 / 共"+NUMPAGES 域+"页"组成，或页眉页脚不满足字体微软雅黑、字号 8 磅、非加粗、灰色（#666666）、水平居中',
    '甲乙方信息表"统一社会信用代码"行不满足甲方为"91420372MA7QX8K59P"、乙方为"91330118MA9LQ3C74J"，或"电子邮箱"行不满足甲方为"service@yunlingsenzhong.cn"、乙方为"support@lanzhoujinpi.cn"，或"通讯地址"行不满足甲方为"成都市锦云路 118 号锦澜商务中心 B 座 18 层"、乙方为"杭州市望潮路 62 号数创园 5 幢 6 层"，或末行不满足"项目对接人：周明厘 / 运营保障部""项目负责人：林知夏 / 区域服务部"',
    '甲乙方信息表不满足 7行4列、四列列宽依次为 2.47cm、7.42cm、2.47cm、7.42cm',
    '服务范围表不满足 5行4列、四列列宽均为 4.88cm、表头依次为"服务模块""包含内容""交付记录""不包含事项"',
    '服务套餐表不满足 6行5列，五个套餐名称不满足依次为"基础巡检包""标准维护包""增强保障包""驻场协同包""单项补充服务"',
    '服务套餐表"选择"列不满足仅"标准维护包"行为"☑"、其余四行均为"□"，或"年度费用"列不满足依次为"¥ 8,600""¥ 18,800""¥ 36,500""另行确认""按工单确认"，或"适用设备规模"列不满足依次为"20 台以内""21-60 台""61-120 台""120 台以上""按需"',
    '签署信息表不满足 5行2列、左右两列列宽相同（均为 8.10cm）',
    '四张表格不满足整表水平居中、四周及内部均为实线网格边框、表格默认单元格左右内边距 0.19cm 且上下内边距 0cm，或单元格逐格内边距不满足：甲乙方信息表全部单元格上下左右均 0.14cm，服务范围表表头行左右 0.12cm 上下 0.16cm 且其余四行上下左右均 0.12cm，服务套餐表表头行左右 0.12cm 上下 0.16cm 且其余五行左右 0.11cm 上下 0.12cm，签署信息表前四行左右 0.14cm 上下 0.16cm',
    '合同标题"办公设备维护与巡检服务合同"不满足字体微软雅黑、加粗、单行水平居中，且字号不小于正文字号',
    '条款标题不满足依次为"一、合作原则与适用依据""二、服务对象与项目范围""三、服务期限与响应安排""四、服务套餐与费用""五、资料、权限与现场配合""六、双方责任与义务""七、验收、交付与服务记录""八、费用调整与新增事项""九、保密与信息安全""十、违约责任"共十条，文字与序号完全一致且均为加粗',
    '"六、双方责任与义务"下不满足保留"（一）甲方责任与义务"5 个编号小项、"（二）乙方责任与义务"5 个编号小项，或两个小标题不满足加粗',
    '合同总价段落不满足"人民币壹万捌仟捌佰元整（¥18,800.00）""签订后 7 个工作日内支付首期款人民币玖仟肆佰元整""履行满 6 个月后 10 个工作日内支付尾款人民币玖仟肆佰元整"的完整表述',
    '乙方收款账户段落不满足"户名：岚舟金辟科技有限公司；开户行：星河商业银行杭州望潮支行；账号：8890 1264 7305 5521 08。"',
    '服务期限不满足"自 2026 年 6 月 1 日起至 2027 年 5 月 31 日止"，或标准服务时间不满足"工作日 9:00 至 18:00"',
    '签署信息表内容不满足"甲方（盖章）：云岭森钟科技有限公司""乙方（盖章）：岚舟金辟科技有限公司""授权代表：陆嘉迪""授权代表：沈若来""签署日期：2026 年 5 月 28 日""签署地点：浙江省杭州市上城区"',
    '全文正文段落与表格单元格文字不满足统一使用微软雅黑字体',
    '正文段落不满足段前间距 0、段后间距 0',
    '合同引言段、服务对象说明段、服务套餐说明段、合同总价段、乙方收款账户段共 5 个说明性段落不满足保留首行缩进，或其余条款正文段落不满足无首行缩进',
    '"十、违约责任"第 4 项与签署信息表之间不满足保留 2 个空段落作为间隔',
    '文档中出现新增的水印、批注、修订痕迹、彩色底纹、边框或"两页排版版"等说明性文字',
    '文档节数不满足为 1 节，或新增了首页不同、奇偶页不同的页眉页脚设置',
]

assert len(RULES) == len(_MAX_DELTAS_ORDER)


# ==================== 六、对外入口 ====================


class _TargetMissing(Exception):
    """目录不存在或目录中没有 .docx：连待评测对象都没有，不进入维度判定。"""


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
            raise _TargetMissing(f"目录不存在：{dir_path}")
        cands = sorted(
            f for f in os.listdir(dir_path)
            if f.lower().endswith(".docx") and not f.startswith("~$")
        )
        if not cands:
            raise _TargetMissing(f"目录中未找到 .docx 文件：{dir_path}")
        file_path = os.path.join(dir_path, cands[0])
        result["file_name"] = os.path.basename(file_path)
        return _evaluate_impl(file_path, result)
    except _TargetMissing as exc:
        # 找不到待评测文件：维度1 未参与判定，dim1_pass / dim1_reason 一律留空，满分也置 0
        result["file_name"] = ""
        result["status"] = "error"
        result["error"] = str(exc)
        result["dim1_pass"] = ""
        result["dim1_reason"] = ""
        result["dim2_items"] = []
        result["total_score"] = 0
        result["max_score"] = 0
        return result
    except Exception as exc:
        # 文件存在但打不开 / 解析失败：维度1 判不通过，dim1_reason 给出错误摘要
        result["status"] = "error"
        result["error"] = f"{type(exc).__name__}: {exc}"
        result["dim1_pass"] = False
        result["dim1_reason"] = f"文件无法解析或无法打开：{type(exc).__name__}: {exc}"
        result["dim2_items"] = []
        result["total_score"] = 0
        return result


def _check_dim1(doc, file_path):
    """维度1：可用与可修改性。返回 (pass, reason)。

    仅要求交付文件是 1 个可正常打开的 .docx：能走到本函数说明 OPC 已解析成功，
    即判为通过；文件损坏打不开的情形由 evaluate() 的异常分支判为不通过。
    """
    return True, ""


def _evaluate_impl(file_path, result):
    doc = _open_word_document(file_path)

    dim1_pass, dim1_reason = _check_dim1(doc, file_path)
    result["dim1_pass"] = dim1_pass
    result["dim1_reason"] = dim1_reason

    resolver = _FontResolver(doc)
    with zipfile.ZipFile(file_path) as zf:
        pkg_names = zf.namelist()
        pkg_document = zf.read("word/document.xml").decode("utf-8", "ignore")
        pkg_settings = ""
        if "word/settings.xml" in pkg_names:
            pkg_settings = zf.read("word/settings.xml").decode("utf-8", "ignore")
        pkg_headers = "".join(
            zf.read(n).decode("utf-8", "ignore") for n in pkg_names
            if re.match(r"word/(header|footer)\d*\.xml$", n)
        )

    # ---------- 顶层段落 ----------
    top_paras = list(_iter_top_level_paragraphs(doc))
    para_texts = [_para_text(p) for p in top_paras]
    para_norms = [_norm(t) for t in para_texts]
    nonempty_idx = [i for i, t in enumerate(para_texts) if t.strip()]
    empty_idx = [i for i, t in enumerate(para_texts) if not t.strip()]

    def find_para(prefix, exact=False):
        key = _norm(prefix)
        for i, txt in enumerate(para_norms):
            if (txt == key) if exact else txt.startswith(key):
                return i
        return -1

    i_title = find_para(TITLE_TEXT, exact=True)
    i_account = find_para(PARA_ACCOUNT)
    i_total = find_para(PARA_TOTAL)
    clause_idx = [find_para(t, exact=True) for t in CLAUSE_TITLES]
    sub_a_idx = find_para(SUB_TITLE_A, exact=True)
    sub_b_idx = find_para(SUB_TITLE_B, exact=True)
    indent_expect_idx = sorted(
        i for i in (find_para(h) for h in INDENT_PARA_HEADS) if i >= 0
    )

    # ---------- 表格 ----------
    tables = list(doc.tables)
    table_cells = [_cell_texts(t) for t in tables]
    table_shapes = [(len(t.rows), len(t.columns)) for t in tables]

    def t0_row(label):
        """按首列标签取甲乙方信息表的整行。"""
        if not table_cells:
            return None
        for row in table_cells[0]:
            if row and row[0] == _norm(label):
                return row
        return None

    # ---------- 版面 ----------
    layout = _probe_layout(doc, file_path)
    page_h_pt = layout.page_height_pt or _twip_to_pt(PG_H_TWIP)

    def page_of(idx):
        if not layout.available or idx < 0 or idx >= len(layout.para_pages):
            return None
        return layout.para_pages[idx]

    def table_span(k):
        if not layout.available or k >= len(layout.table_pages):
            return (None, None)
        return layout.table_pages[k]

    results = []
    details = []

    def mark(ok, detail):
        """加分项：ok 表示达成；减分项：ok 表示未触发扣分。"""
        idx = len(results)
        max_delta = _MAX_DELTAS_ORDER[idx]
        if max_delta > 0:
            head = "+%d" % max_delta if ok else "+0"
        else:
            head = "-0" if ok else "%d" % max_delta
        results.append("%s：%s" % (head, RULES[idx]))
        details.append(detail)

    no_layout = layout.reason or "版面信息不可用"

    # ---------------- A01 +5 总页数为 2 ----------------
    # 严格对齐细则，仅踩细则明确写出的点：①ComputeStatistics(wdStatisticPages) 为 2；
    # ②Information(wdNumberOfPagesInDocument) 为 2。与办公软件对齐：两个统计口径在
    # Word 中可能不一致（域未更新时），细则要求二者均为 2，故两者都取。
    if not layout.available:
        mark(False, no_layout)
    else:
        ok = layout.pages_statistic == 2 and layout.pages_information == 2
        mark(ok, "ComputeStatistics=%s，Information=%s（来源 %s）" % (
            layout.pages_statistic, layout.pages_information, layout.source))

    # ---------------- A02 +5 标题至乙方收款账户段全部落在第1页 ----------------
    # ①合同标题、②合同引言段、③"一、"至"四、"全部条款正文、④合同总价段、⑤乙方收款账户段
    # 在文档顺序上即顶层段落 0 至"乙方收款账户段"，逐段核对其真实页码是否为 1。
    if not layout.available:
        mark(False, no_layout)
    elif i_account < 0:
        mark(False, "未找到乙方收款账户段落，无法界定第1页应承载的范围")
    else:
        spill = [i for i in range(0, i_account + 1)
                 if i in nonempty_idx and page_of(i) not in (None, 1)]
        tbl_spill = [k for k in range(min(3, len(tables))) if table_span(k)[1] not in (None, 1)]
        ok = not spill and not tbl_spill
        mark(ok, "溢出到第2页的段落序号=%s；溢出的表格序号=%s" % (spill, tbl_spill))

    # ---------------- A03 +5 分页断点位置 ----------------
    # ①第1页最后一个有字段落为"乙方收款账户：…08。"；②第2页第一个段落为"五、资料、权限与现场配合"。
    if not layout.available:
        mark(False, no_layout)
    else:
        page1_last = None
        page2_first = None
        for i in nonempty_idx:
            pg = page_of(i)
            if pg == 1:
                page1_last = i
            elif pg == 2 and page2_first is None:
                page2_first = i
        ok = (page1_last == i_account and i_account >= 0
              and page2_first == clause_idx[4] and clause_idx[4] >= 0)
        mark(ok, "第1页末段=%r，第2页首段=%r" % (
            para_texts[page1_last][:24] if page1_last is not None else None,
            para_texts[page2_first][:24] if page2_first is not None else None))

    # ---------------- A04 +3 第2页完整承载"五、"至"十、" ----------------
    # ①六条条款标题均在第2页；②序号顺序连续递增；③"五、"之后的所有正文段落无一被挤回第1页。
    if not layout.available:
        mark(False, no_layout)
    elif min(clause_idx[4:]) < 0:
        mark(False, "第2页应承载的六条条款标题存在缺失：%s" % [
            CLAUSE_TITLES[4 + k] for k, v in enumerate(clause_idx[4:]) if v < 0])
    else:
        tail = clause_idx[4:]
        on_page2 = [page_of(i) == 2 for i in tail]
        ordered = all(tail[k] < tail[k + 1] for k in range(len(tail) - 1))
        pushed_back = [i for i in nonempty_idx if i >= tail[0] and page_of(i) == 1]
        ok = all(on_page2) and ordered and not pushed_back
        mark(ok, "六条条款页码=%s，序号连续=%s，被挤回第1页的段落=%s" % (
            [page_of(i) for i in tail], ordered, pushed_back))

    # ---------------- A05 +3 签署信息表位于第2页下半部 ----------------
    # ①表格起止行同在第2页；②表格起始纵向位置 > 页高 50%。
    if not layout.available:
        mark(False, no_layout)
    elif len(tables) < 4:
        mark(False, "未找到第4张（签署信息）表格")
    else:
        span = table_span(3)
        top = layout.table_tops[3] if len(layout.table_tops) > 3 else None
        ratio = (top / page_h_pt) if (top and page_h_pt) else None
        ok = span == (2, 2) and ratio is not None and ratio > SIGN_TABLE_MIN_RATIO
        mark(ok, "签署信息表页码=%s，起始纵向位置=%s pt（占页高 %s）" % (
            span, None if top is None else round(top, 1),
            None if ratio is None else "%.1f%%" % (ratio * 100)))

    # ---------------- A06 +3 服务范围表整体在第1页 ----------------
    if not layout.available:
        mark(False, no_layout)
    elif len(tables) < 2:
        mark(False, "未找到服务范围表")
    else:
        span = table_span(1)
        header_ok = table_cells[1][0] == [_norm(x) for x in T1_HEADER]
        ok = span == (1, 1) and header_ok
        mark(ok, "服务范围表页码=%s，表头匹配=%s" % (span, header_ok))

    # ---------------- A07 +5 服务套餐表整体在第1页且中间不断页 ----------------
    if not layout.available:
        mark(False, no_layout)
    elif len(tables) < 3:
        mark(False, "未找到服务套餐表")
    else:
        span = table_span(2)
        header_ok = table_cells[2][0] == [_norm(x) for x in T2_HEADER]
        ok = span == (1, 1) and table_shapes[2] == TABLE_SHAPES[2] and header_ok
        mark(ok, "服务套餐表页码=%s，行列=%s，表头匹配=%s" % (span, table_shapes[2], header_ok))

    # ---------------- A08 +3 全文同一字号且不大于 7 磅 ----------------
    # ①统计全文（正文段落 + 表格单元格）所有含字 run 的生效字号，取值集合必须只有 1 种；
    # ②该字号不大于 7 磅——只"统一"不"压小"不算完成排版任务，故必须同时满足上限。
    size_set = set()
    for r_el in _iter_all_runs(doc):
        if not _para_text(r_el).strip():
            continue
        val = resolver.size_pt(r_el)
        if val:
            size_set.add(round(val, 2))
    body_font_pt = min(size_set) if size_set else None
    ok = len(size_set) == 1 and body_font_pt is not None and body_font_pt <= BODY_MAX_FONT_PT
    mark(ok, "全文字号集合=%s 磅（上限 %s 磅）" % (sorted(size_set), BODY_MAX_FONT_PT))

    # ---------------- A09 +3 正文段落统一行距且不大于单倍行距 ----------------
    # ①只看正文段落（顶层段落）的生效 (w:line, w:lineRule)，取值集合必须只有 1 种；
    # ②行距值不超过单倍行距 240 twips（未显式设置行距即为单倍，视为达标）。
    line_set = set()
    for p_el in top_paras:
        _, _, line, rule = resolver.spacing(p_el)
        line_set.add((line, rule))
    line_ok = len(line_set) == 1 and all(
        line is None or line <= LINE_SINGLE_TWIP for line, _rule in line_set)
    mark(line_ok, "正文行距集合=%s（上限 %d twips）" % (
        sorted(line_set, key=lambda x: (str(x[0]), str(x[1]))), LINE_SINGLE_TWIP))

    # ---------------- A10 +3 表格行最小高度随字号同步下调 ----------------
    # 逐行核对：行最小高度 <= 2.0 × 该行字号 × 该行估算行数 + 10 磅余量。
    # 超出即说明行高没有随字号一起收紧，表格里会留出成片空白。
    over_rows = []
    for k, table in enumerate(tables):
        heights = _row_heights(table)
        for ri, (hval, _hrule) in enumerate(heights):
            if not hval:
                continue
            widths = _row_cell_widths_twip(table, ri)
            est_lines = 1
            row_font = 0.0
            for ci, cell in enumerate(table.rows[ri].cells):
                mar = _effective_cell_margins(table, cell._tc)
                width_tw = widths[ci] if ci < len(widths) else None
                avail_pt = None
                if width_tw:
                    avail_pt = _twip_to_pt(width_tw) - _twip_to_pt(
                        (mar["left"] or 0) + (mar["right"] or 0))
                cell_lines = 0
                for para in cell.paragraphs:
                    fpt = resolver.para_size_pt(para._p) or body_font_pt or BODY_MIN_FONT_PT
                    row_font = max(row_font, fpt)
                    cell_lines += _estimate_lines(para.text, avail_pt, fpt)
                est_lines = max(est_lines, cell_lines or 1)
            limit = TRH_LINE_FACTOR * (row_font or BODY_MIN_FONT_PT) * est_lines + TRH_SLACK_PT
            if _twip_to_pt(hval) > limit:
                over_rows.append((k, ri, round(_twip_to_pt(hval), 1), round(limit, 1)))
    mark(not over_rows, "行高超出合理上限的 (表,行,行高pt,上限pt)=%s" % over_rows)

    # ---------------- A11 +1 两页版面有效利用、无空白页 ----------------
    # ①总页数为 2；②第1页最后一个元素的纵向位置 >= 页高 75%（底部空白不超过页高 1/4）；
    # ③最后一个段落仍落在第2页，第2页之后无多余空页。
    if not layout.available:
        mark(False, no_layout)
    else:
        fill_ratio = None
        if layout.page1_max_top and page_h_pt:
            fill_ratio = layout.page1_max_top / page_h_pt
        ok = (layout.page_count == 2
              and fill_ratio is not None and fill_ratio >= PAGE1_FILL_RATIO
              and layout.last_para_page == 2)
        mark(ok, "总页数=%s，第1页最末元素位于页高 %s，末段所在页=%s" % (
            layout.page_count,
            None if fill_ratio is None else "%.1f%%" % (fill_ratio * 100),
            layout.last_para_page))

    # ---------------- B01 -5 删条款 / 删表格行 / 删正文段落强行压页 ----------------
    # ①十条条款标题齐全；②每条条款下的编号小项数量与合同条文一致（不得删段压页）；
    # ③四张表格齐全且行列数未被删减；④除签署表末行留白外表格无空单元格。
    missing_clause = [CLAUSE_TITLES[k] for k, v in enumerate(clause_idx) if v < 0]
    counts = []
    if not missing_clause:
        bounds = clause_idx + [len(top_paras)]
        for k in range(10):
            lo, hi = bounds[k] + 1, bounds[k + 1]
            counts.append(len([i for i in nonempty_idx if lo <= i < hi]))
    empty_cells = []
    for k, rows in enumerate(table_cells):
        for ri, row in enumerate(rows):
            for ci, cell in enumerate(row):
                if not cell and not (k == 3 and ri == len(rows) - 1):
                    empty_cells.append((k, ri, ci))
    shape_ok = table_shapes[:4] == TABLE_SHAPES
    ok = (not missing_clause
          and counts == CLAUSE_BODY_COUNTS
          and len(nonempty_idx) == NONEMPTY_TOPLEVEL_PARAS
          and shape_ok
          and not empty_cells)
    mark(ok, "缺失条款=%s，各条小项数=%s（应为 %s），有字段落数=%d（应为 %d），表格行列=%s，空单元格=%s" % (
        missing_clause, counts, CLAUSE_BODY_COUNTS, len(nonempty_idx),
        NONEMPTY_TOPLEVEL_PARAS, table_shapes[:4], empty_cells))

    # ---------------- B02 -3 字号压到 6 磅以下或文字被裁切 ----------------
    # ①最小字号 >= 6 磅；②凡使用固定行距（w:lineRule="exact"）的有字段落，
    # 其行距值必须不小于该段字号，否则上下笔画会被裁切。
    clipped = []
    for p_el in _iter_all_paragraph_elements(doc):
        if not _para_text(p_el).strip():
            continue
        _, _, line, rule = resolver.spacing(p_el)
        fpt = resolver.para_size_pt(p_el) or body_font_pt
        if rule == "exact" and line and fpt and _twip_to_pt(line) < fpt:
            clipped.append((_para_text(p_el)[:14], _twip_to_pt(line), fpt))
    size_ok = bool(size_set) and min(size_set) >= BODY_MIN_FONT_PT
    mark(size_ok and not clipped, "最小字号=%s 磅，行高小于字号的段落=%s" % (
        min(size_set) if size_set else None, clipped[:5]))

    # ---------------- B03 -3 页面设置 ----------------
    # ①A4 纵向 21cm×29.7cm；②上下页边距 0.35cm；③左右页边距 0.85cm；
    # ④页眉距顶端 0.20cm；⑤页脚距底端 0.20cm。统一换算为 twips 后按 ±6 twips 容差比较。
    sm = _section_metrics(doc)

    def near(value, target, tol=PG_TOL_TWIP):
        return value is not None and abs(value - target) <= tol

    ok = (sm["orient"] == "portrait"
          and near(sm["pg_w"], PG_W_TWIP) and near(sm["pg_h"], PG_H_TWIP)
          and near(sm["top"], MAR_TB_TWIP) and near(sm["bottom"], MAR_TB_TWIP)
          and near(sm["left"], MAR_LR_TWIP) and near(sm["right"], MAR_LR_TWIP)
          and near(sm["header"], HDR_FTR_TWIP) and near(sm["footer"], HDR_FTR_TWIP))
    mark(ok, "页面(twips) 宽高=%s×%s 上下=%s/%s 左右=%s/%s 页眉页脚=%s/%s" % (
        sm["pg_w"], sm["pg_h"], sm["top"], sm["bottom"],
        sm["left"], sm["right"], sm["header"], sm["footer"]))

    # ---------------- B04 -3 甲乙方信息表"联系号码"行 ----------------
    row = t0_row("联系号码")
    ok = row == [_norm(x) for x in T0_ROW_PHONE]
    mark(ok, "联系号码行=%s（应为 %s）" % (row, [_norm(x) for x in T0_ROW_PHONE]))

    # ---------------- B05 -1 插入空段落 / 手动分页符分节符硬拆 ----------------
    # ①正文无 w:br w:type="page" 手动分页符；②整篇仅 1 个 w:sectPr（无手动分节）；
    # ③空段落总数不超过 3 个（"十、"第4项后 2 个间隔段 + 文末 1 个段落标记）。
    manual_break = len(re.findall(r'<w:br\b[^>]*w:type="page"', pkg_document))
    sect_count = pkg_document.count("<w:sectPr")
    ok = manual_break == 0 and sect_count == 1 and len(empty_idx) <= EMPTY_TOPLEVEL_PARAS_MAX
    mark(ok, "手动分页符=%d，w:sectPr=%d，空段落数=%d" % (manual_break, sect_count, len(empty_idx)))

    # ---------------- B06 -1 页眉页脚 ----------------
    # ①页眉文字为合同名；②页脚由 "第"+PAGE 域+"页 / 共"+NUMPAGES 域+"页" 组成；
    # ③页眉页脚字体微软雅黑、字号 8 磅、非加粗、颜色 #666666；④均水平居中。
    section = doc.sections[0]
    header = section.header
    footer = section.footer
    hdr_paras = [p.text for p in header.paragraphs if p.text.strip()]
    ftr_paras = [p.text for p in footer.paragraphs if p.text.strip()]
    hdr_text_ok = len(hdr_paras) == 1 and _norm(hdr_paras[0]) == _norm(HEADER_TEXT)
    ftr_codes = " ".join(_footer_field_codes(footer))
    ftr_field_ok = "PAGE" in ftr_codes and "NUMPAGES" in ftr_codes
    ftr_text_ok = len(ftr_paras) == 1 and bool(FOOTER_TEXT_RE.match(_norm(ftr_paras[0])))
    props_ok = (_hdr_ftr_run_props_ok(_hdr_ftr_runs(header))
                and _hdr_ftr_run_props_ok(_hdr_ftr_runs(footer)))
    center_ok = _hdr_ftr_centered(header) and _hdr_ftr_centered(footer)
    ok = hdr_text_ok and ftr_field_ok and ftr_text_ok and props_ok and center_ok
    mark(ok, "页眉文字=%s，页脚域=%s，页脚文字=%s，字体字号颜色=%s，居中=%s（页眉%r 页脚%r）" % (
        hdr_text_ok, ftr_field_ok, ftr_text_ok, props_ok, center_ok,
        hdr_paras[:1], ftr_paras[:1]))

    # ---------------- B07 -1 信用代码 / 邮箱 / 地址 / 末行 ----------------
    # ①统一社会信用代码行；②电子邮箱行；③通讯地址行；④末行项目对接人 / 项目负责人。
    r_credit = t0_row("统一社会信用代码")
    r_mail = t0_row("电子邮箱")
    r_addr = t0_row("通讯地址")
    r_last = table_cells[0][-1] if table_cells and table_cells[0] else None
    checks = {
        "信用代码": r_credit == [_norm(x) for x in T0_ROW_CREDIT],
        "电子邮箱": r_mail == [_norm(x) for x in T0_ROW_MAIL],
        "通讯地址": r_addr == [_norm(x) for x in T0_ROW_ADDR],
        "末行": r_last == [_norm(x) for x in T0_ROW_LAST],
    }
    mark(all(checks.values()), "逐行核对=%s，实际末行=%s" % (checks, r_last))

    # ---------------- B08 -1 甲乙方信息表 7行4列与列宽 ----------------
    widths0 = _row_cell_widths_twip(tables[0], 0) if tables else []
    target0 = [_cm_to_twip(v) for v in TABLE_COL_CM[0]]
    width_ok = len(widths0) == len(target0) and all(
        w is not None and abs(w - t) <= COL_W_TOL_TWIP for w, t in zip(widths0, target0))
    ok = table_shapes[:1] == TABLE_SHAPES[:1] and width_ok
    mark(ok, "行列=%s，列宽(twips)=%s（应为 %s）" % (
        table_shapes[0] if table_shapes else None, widths0, target0))

    # ---------------- B09 -1 服务范围表 5行4列 / 列宽 4.88cm / 表头 ----------------
    widths1 = _row_cell_widths_twip(tables[1], 0) if len(tables) > 1 else []
    target1 = [_cm_to_twip(v) for v in TABLE_COL_CM[1]]
    width_ok = len(widths1) == len(target1) and all(
        w is not None and abs(w - t) <= COL_W_TOL_TWIP for w, t in zip(widths1, target1))
    header_ok = len(table_cells) > 1 and table_cells[1][0] == [_norm(x) for x in T1_HEADER]
    ok = table_shapes[1:2] == TABLE_SHAPES[1:2] and width_ok and header_ok
    mark(ok, "行列=%s，列宽(twips)=%s（应为 %s），表头=%s" % (
        table_shapes[1] if len(table_shapes) > 1 else None, widths1, target1,
        table_cells[1][0] if len(table_cells) > 1 else None))

    # ---------------- B10 -1 服务套餐表 6行5列与五个套餐名称 ----------------
    names = [row[1] for row in table_cells[2][1:]] if len(table_cells) > 2 else []
    ok = table_shapes[2:3] == TABLE_SHAPES[2:3] and names == [_norm(x) for x in T2_PACKAGES]
    mark(ok, "行列=%s，套餐名称=%s" % (
        table_shapes[2] if len(table_shapes) > 2 else None, names))

    # ---------------- B11 -1 服务套餐表 选择列 / 年度费用 / 适用设备规模 ----------------
    # ①"选择"列仅"标准维护包"行为 ☑，其余四行为 □；②年度费用五档；③适用设备规模五档。
    if len(table_cells) > 2 and len(table_cells[2]) == 6:
        body = table_cells[2][1:]
        choice = [r[0] for r in body]
        fees = [r[3] for r in body]
        scales = [r[2] for r in body]
    else:
        choice = fees = scales = []
    checks = {
        "选择列": choice == [_norm(x) for x in T2_CHOICES],
        "年度费用": fees == [_norm(x) for x in T2_FEES],
        "适用设备规模": scales == [_norm(x) for x in T2_SCALES],
    }
    mark(all(checks.values()), "%s；选择=%s 费用=%s 规模=%s" % (checks, choice, fees, scales))

    # ---------------- B12 -1 签署信息表 5行2列且左右列宽相同 ----------------
    widths3 = _row_cell_widths_twip(tables[3], 0) if len(tables) > 3 else []
    target3 = [_cm_to_twip(v) for v in TABLE_COL_CM[3]]
    same = len(widths3) == 2 and widths3[0] is not None and widths3[0] == widths3[1]
    width_ok = same and all(
        abs(w - t) <= COL_W_TOL_TWIP for w, t in zip(widths3, target3))
    ok = table_shapes[3:4] == TABLE_SHAPES[3:4] and width_ok
    mark(ok, "行列=%s，列宽(twips)=%s（应为 %s，且左右相同）" % (
        table_shapes[3] if len(table_shapes) > 3 else None, widths3, target3))

    # ---------------- B13 -1 四表居中 / 实线网格 / 表级内边距 / 逐格内边距 ----------------
    # ①tblPr 的 w:jc 为 center；②四周 top/left/bottom/right 与内部 insideH/insideV
    # 六条边均为 w:val="single" 实线；③表级 w:tblCellMar 左右 108 twips(0.19cm)、上下 0；
    # ④每个单元格自身的 w:tcMar 按 CELL_MAR_PER_TABLE 逐格核对（表级默认相同时视觉内边距
    # 仍可能被整体改小，故必须单独校验，签署信息表末行不作要求）。
    bad_tables = []
    bad_cell_mar = []
    for k, table in enumerate(tables[:4]):
        jc = _table_jc(table)
        borders = _table_borders(table)
        mars = _table_cell_margins(table)
        border_ok = all(v is not None and v[0] == "single" for v in borders.values())
        mar_ok = (mars["left"] is not None and abs(mars["left"] - CELL_MAR_LR_TWIP) <= CELL_MAR_TOL_TWIP
                  and mars["right"] is not None and abs(mars["right"] - CELL_MAR_LR_TWIP) <= CELL_MAR_TOL_TWIP
                  and (mars["top"] or 0) == CELL_MAR_TB_TWIP
                  and (mars["bottom"] or 0) == CELL_MAR_TB_TWIP)
        if not (jc == "center" and border_ok and mar_ok):
            bad_tables.append((k, jc, {e: (v[0] if v else None) for e, v in borders.items()}, mars))
        spec = CELL_MAR_PER_TABLE[k] if k < len(CELL_MAR_PER_TABLE) else {}
        for ri, row in enumerate(table.rows):
            want = spec.get(ri, spec.get("*"))
            if want is None:
                continue
            for ci, cell in enumerate(row.cells):
                mm = _cell_margins(cell._tc)
                got = (mm["left"], mm["right"], mm["top"], mm["bottom"])
                if any(g is None or abs(g - w) > CELL_MAR_TOL_TWIP for g, w in zip(got, want)):
                    bad_cell_mar.append((k, ri, ci, got, want))
    mark(not bad_tables and not bad_cell_mar, "不合规的表格=%s；逐格内边距异常=%s" % (
        bad_tables, bad_cell_mar[:6]))

    # ---------------- B14 -1 合同标题字体 / 加粗 / 单行居中 / 字号不小于正文 ----------------
    # ①文字为"办公设备维护与巡检服务合同"；②生效字体微软雅黑；③全部 run 加粗；
    # ④水平居中且不换行（段内无 w:br 且估算宽度不超过版心宽度）；⑤字号 >= 正文字号。
    if i_title < 0:
        mark(False, "未找到合同标题段落")
    else:
        title_el = top_paras[i_title]
        fonts = [resolver.font_names(r) for r in title_el.iter(qn("w:r"))
                 if _para_text(r).strip()]
        font_ok = bool(fonts) and all(
            east == HDR_FTR_FONT and (asc is None or asc == HDR_FTR_FONT)
            for east, asc in fonts)
        bold_ok = _para_is_bold(resolver, title_el)
        center_ok = resolver.jc(title_el) == "center"
        title_pt = resolver.para_size_pt(title_el)
        content_pt = _twip_to_pt(PG_W_TWIP - 2 * MAR_LR_TWIP)
        single_line = (len(list(title_el.iter(qn("w:br")))) == 0
                       and _text_width_pt(para_texts[i_title], title_pt or 0) <= content_pt)
        size_ok = bool(title_pt) and bool(body_font_pt) and title_pt >= body_font_pt - 0.01
        ok = font_ok and bold_ok and center_ok and single_line and size_ok
        mark(ok, "字体=%s 加粗=%s 居中=%s 单行=%s 字号=%s（正文 %s）" % (
            font_ok, bold_ok, center_ok, single_line, title_pt, body_font_pt))

    # ---------------- B15 -1 十条条款标题文字序号一致且加粗 ----------------
    not_bold = [CLAUSE_TITLES[k] for k, i in enumerate(clause_idx)
                if i >= 0 and not _para_is_bold(resolver, top_paras[i])]
    order_ok = all(clause_idx[k] >= 0 for k in range(10)) and all(
        clause_idx[k] < clause_idx[k + 1] for k in range(9))
    ok = order_ok and not not_bold
    mark(ok, "缺失或错序=%s，未加粗=%s" % (missing_clause, not_bold))

    # ---------------- B16 -1 "六、"下两组各 5 个编号小项且小标题加粗 ----------------
    # ①（一）甲方责任与义务 下 5 个"1.~5."小项；②（二）乙方责任与义务 下 5 个小项；③两小标题加粗。
    if sub_a_idx < 0 or sub_b_idx < 0 or sub_a_idx >= sub_b_idx:
        mark(False, "未按顺序找到两个小标题：（一）=%d （二）=%d" % (sub_a_idx, sub_b_idx))
    else:
        end_b = clause_idx[6] if clause_idx[6] > sub_b_idx else len(top_paras)
        group_a = [para_norms[i] for i in nonempty_idx if sub_a_idx < i < sub_b_idx]
        group_b = [para_norms[i] for i in nonempty_idx if sub_b_idx < i < end_b]
        numbered = lambda seq: [t for t in seq if re.match(r"^[1-5][.、]", t)]
        bold_ok = (_para_is_bold(resolver, top_paras[sub_a_idx])
                   and _para_is_bold(resolver, top_paras[sub_b_idx]))
        ok = (len(group_a) == 5 and len(numbered(group_a)) == 5
              and len(group_b) == 5 and len(numbered(group_b)) == 5 and bold_ok)
        mark(ok, "（一）小项=%d/%d，（二）小项=%d/%d，小标题加粗=%s" % (
            len(numbered(group_a)), len(group_a), len(numbered(group_b)), len(group_b), bold_ok))

    # ---------------- B17 -1 合同总价段落完整表述 ----------------
    total_text = para_norms[i_total] if i_total >= 0 else ""
    missing = [f for f in TOTAL_FRAGMENTS if _norm(f) not in total_text]
    mark(i_total >= 0 and not missing, "缺失表述=%s" % missing)

    # ---------------- B18 -1 乙方收款账户段落完整表述 ----------------
    account_text = para_norms[i_account] if i_account >= 0 else ""
    ok = i_account >= 0 and _norm(ACCOUNT_FRAGMENT) in account_text
    mark(ok, "账户段落=%r" % (para_texts[i_account][:80] if i_account >= 0 else None))

    # ---------------- B19 -1 服务期限与标准服务时间 ----------------
    all_norm = "".join(para_norms)
    term_ok = _norm(TERM_FRAGMENT) in all_norm
    work_ok = _norm(WORKTIME_FRAGMENT) in all_norm
    mark(term_ok and work_ok, "服务期限=%s，标准服务时间=%s" % (term_ok, work_ok))

    # ---------------- B20 -1 签署信息表内容 ----------------
    actual3 = table_cells[3] if len(table_cells) > 3 else []
    target = [[_norm(c) for c in row] for row in T3_ROWS]
    ok = actual3 == target
    mark(ok, "签署信息表=%s" % actual3)

    # ---------------- B21 -1 全文统一微软雅黑 ----------------
    # 逐个含字 run 解析生效字体（run rPr -> 段落样式链 -> 默认段落样式 -> docDefaults），
    # 中日韩字体 w:eastAsia 与西文 w:ascii 都必须是 微软雅黑。
    bad_fonts = {}
    for r_el in _iter_all_runs(doc):
        text = _para_text(r_el)
        if not text.strip():
            continue
        east, asc = resolver.font_names(r_el)
        if east != HDR_FTR_FONT or (asc is not None and asc != HDR_FTR_FONT):
            bad_fonts.setdefault((east, asc), 0)
            bad_fonts[(east, asc)] += 1
    mark(not bad_fonts, "非微软雅黑的 (eastAsia, ascii) 分布=%s" % bad_fonts)

    # ---------------- B22 -1 正文段落段前段后间距 0 ----------------
    bad_spacing = []
    for i, p_el in enumerate(top_paras):
        before, after, _, _ = resolver.spacing(p_el)
        if (before or 0) != 0 or (after or 0) != 0:
            bad_spacing.append((i, before, after))
    mark(not bad_spacing, "段前段后不为 0 的 (段序号, before, after)=%s" % bad_spacing[:8])

    # ---------------- B23 -1 5 个说明性段落保留首行缩进、其余无首行缩进 ----------------
    # 合同引言段、服务对象说明段、服务套餐说明段、合同总价段、乙方收款账户段应有首行缩进；
    # 其余条款正文段落（含条款标题与编号小项）首行缩进必须为 0。
    indented = []
    for i, p_el in enumerate(top_paras):
        first, chars = resolver.first_line_indent(p_el)
        if (first or 0) > 0 or (chars or 0) > 0:
            indented.append(i)
    ok = indented == indent_expect_idx and len(indent_expect_idx) == 5
    mark(ok, "实际缩进段落=%s，应缩进段落=%s" % (indented, indent_expect_idx))

    # ---------------- B24 -1 "十、违约责任"第 4 项与签署信息表之间 2 个空段落 ----------------
    # 按 body 子元素顺序核对：最后一个有字段落之后应恰好是 2 个空段落，紧接签署信息表。
    seq = list(_iter_body_children(doc))
    tbl3_el = tables[3]._tbl if len(tables) > 3 else None
    gap = None
    prev_text = None
    if tbl3_el is not None:
        for pos, (kind, el) in enumerate(seq):
            if kind != "tbl" or el is not tbl3_el:
                continue
            j = pos - 1
            run_back = 0
            while j >= 0 and seq[j][0] == "p" and not _para_text(seq[j][1]).strip():
                run_back += 1
                j -= 1
            gap = run_back
            if j >= 0 and seq[j][0] == "p":
                prev_text = _para_text(seq[j][1])
            break
    prev_ok = prev_text is not None and re.match(r"^4[.、]", _norm(prev_text)) is not None
    mark(gap == 2 and prev_ok, "签署表前空段落数=%s，其上一段=%r" % (gap, (prev_text or "")[:20]))

    # ---------------- B25 -1 无新增水印 / 批注 / 修订 / 彩色底纹 / 边框 / 说明性文字 ----------------
    # ①无水印（w:pict 中的 WordArt / 页眉图形、w:background）；②无批注部件；
    # ③无修订标记（w:ins/w:del/w:rPrChange/w:pPrChange）；④无荧光标记与段落边框；
    # ⑤底纹填充色种类与数量与交付要求一致（EDEDED/F0F0F0/FFF8D6/D9EAF7/F4F9FC）；
    # ⑥正文与页眉页脚中不得出现"两页排版"等说明性文字。
    issues = []
    if re.search(r'<v:shape[^>]*type="#_x0000_t136"', pkg_document + pkg_headers) \
            or "<w:background" in pkg_document or "WordArt" in pkg_headers:
        issues.append("水印/艺术字")
    if any(n.startswith("word/comments") for n in pkg_names):
        issues.append("批注部件")
    for tag in ("w:ins", "w:del", "w:rPrChange", "w:pPrChange"):
        if re.search(r"<%s[ />]" % tag, pkg_document):
            issues.append("修订标记 %s" % tag)
    if "<w:highlight" in pkg_document:
        issues.append("荧光标记")
    if "<w:pBdr" in pkg_document:
        issues.append("段落边框")
    shd_actual = {}
    for fill in re.findall(r'<w:shd\b[^>]*?w:fill="([^"]+)"', pkg_document):
        if fill.upper() in ("AUTO", "FFFFFF"):
            continue
        shd_actual[fill.upper()] = shd_actual.get(fill.upper(), 0) + 1
    if shd_actual != SHD_FILL_EXPECT:
        issues.append("底纹填充色分布=%s（应为 %s）" % (shd_actual, SHD_FILL_EXPECT))
    hit_texts = [t for t in FORBIDDEN_TEXTS
                 if _norm(t) in "".join(para_norms) or _norm(t) in _norm(pkg_headers)]
    if hit_texts:
        issues.append("说明性文字=%s" % hit_texts)
    mark(not issues, "问题=%s" % issues if issues else "未见水印/批注/修订/异常底纹/边框/说明性文字")

    # ---------------- B26 -1 节数为 1 且无首页不同、奇偶页不同页眉页脚 ----------------
    sec_issues = []
    if len(doc.sections) != 1:
        sec_issues.append("节数=%d" % len(doc.sections))
    if "<w:titlePg" in pkg_document:
        sec_issues.append("启用了首页不同页眉页脚")
    if re.search(r"<w:evenAndOddHeaders\b(?![^>]*w:val=\"(0|false)\")", pkg_settings):
        sec_issues.append("启用了奇偶页不同页眉页脚")
    mark(not sec_issues, "问题=%s" % sec_issues if sec_issues else "单节且页眉页脚全局统一")

    # ==================== 三、汇总维度2 ====================
    # 逆向解析 results 中的 "+N：<细则原文>" / "-N：<细则原文>" 头部，与 _MAX_DELTAS_ORDER
    # 逐项对齐后组装 dim2_items；加分项 hit 表示拿到分，减分项 hit 表示没有触发扣分。
    if len(results) != len(_MAX_DELTAS_ORDER):
        raise RuntimeError("dim2 项数不一致：results=%d vs order=%d" % (
            len(results), len(_MAX_DELTAS_ORDER)))

    dim2_items = []
    total_score = 0
    for raw, max_delta in zip(results, _MAX_DELTAS_ORDER):
        head, _, rule = raw.partition("：")
        try:
            delta = int(head)
        except ValueError:
            delta = 0
        if max_delta > 0:
            hit = delta > 0
            actual = max_delta if hit else 0
        else:
            hit = delta == 0
            actual = 0 if hit else max_delta
        dim2_items.append({
            "rule": rule,
            "max_delta": max_delta,
            "delta": actual,
            "hit": hit,
        })
        total_score += actual

    result["dim2_items"] = dim2_items
    result["total_score"] = 0 if not result["dim1_pass"] else max(0, total_score)
    result["max_score"] = MAX_SCORE_TOTAL
    return result


if __name__ == "__main__":
    # 未显式传目录时，默认评测脚本自身所在目录（而非当前工作目录），
    # 便于把脚本直接放进 输入文件/good输出文件/bad输出文件 里运行。
    if len(sys.argv) > 1:
        _target = sys.argv[1]
    else:
        _target = os.path.dirname(os.path.abspath(__file__))
    sys.stderr.write(f"评测目录：{_target}\n")
    _out = evaluate(_target)
    sys.stdout.write(json.dumps(_out, ensure_ascii=False, indent=2) + "\n")
