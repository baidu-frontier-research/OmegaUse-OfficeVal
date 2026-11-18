#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# pyright: reportUnknownMemberType=false, reportUnknownVariableType=false, reportUnknownArgumentType=false, reportMissingImports=false, reportPrivateUsage=false, reportOptionalMemberAccess=false, reportAttributeAccessIssue=false, reportIndexIssue=false, reportOptionalSubscript=false

"""officeval_012_verifier —— “毕业调研报告目录重建 + 全文格式规范” 任务评分脚本。

约定与 officeval_008_verifier 保持一致：
  * 本模块不修改 sys.stdout，也不打印主结果；
  * 对外只暴露 evaluate(dir_path) -> dict，dir_path 可省略（默认脚本所在目录）；
  * __main__ 分支仅用于本地调试。

评分严格对齐《细则.txt》：维度 1 为“可用与可修改性”门槛（不通过则完成度记 0），
维度 2 共 38 条——15 条加分项合计 49 分，23 条减分项合计 -59 分，总分下限截断为 0。
其中前 8 条减分项为回退检测：原稿即已达标的格式被改坏才扣分，未改动则不增不减，
使未作任何修改的原稿基线恰为 0 分。
所有判定均按“与办公软件一致”的读取语义：段落属性走
段落直接 pPr → pStyle 及其 basedOn 链 → docDefaults；
字体走 run rPr → rStyle 链 → pStyle 链 → docDefaults → theme1.xml 主题字体。
"""

import os
import re
import sys
import json
import traceback

from docx import Document
from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph

try:
    from lxml import etree
except ImportError:  # pragma: no cover - lxml 随 python-docx 一同安装
    etree = None


SCRIPT_ID = "012"
MAX_SCORE_TOTAL = 49

# 需要跳过的输入素材：格式规范文档为评测的参照物而非被评对象，
# 不应被当作输出文件打分。
SPEC_FILE_KEYWORD = "毕业论文格式"

# 每个元素对应 _evaluate_impl 中一次 results.append，正数为加分项满分，
# 负数为减分项一旦命中即扣除的分值。
# 原稿即已达标的 8 项（下标注“回退项”）不计入加分，改为回退检测：
# 模型把原本合规的格式改坏才扣分，未改动则不增不减，使未修改的原稿基线为 0 分。
_MAX_DELTAS_ORDER = [
    -5,  # 01 回退项：目录的自动更新 TOC 域被破坏
    -3,  # 02 回退项：目录条目出现缺条、重复或无关条目
    3,   # 03 目录条目按层级缩进
    -3,  # 04 回退项：目录右对齐点状前导符制表位被移除
    3,   # 05 目录条目字体字号行距
    3,   # 06 “目  录”标题二号黑体居中
    -1,  # 07 回退项：目录不再独立成页或目录页混入正文文字
    5,   # 08 7 个一级标题小三黑体居中顶格
    3,   # 09 一级标题段前段后 0.5 行、两倍行距
    5,   # 10 10 个二级标题四号黑体加粗顶格、同一段落样式
    3,   # 11 二级标题段前段后 0.5 行、单倍行距
    -5,  # 12 回退项：正文小四宋体/TNR、首行缩进、1.5 倍行距被改动
    -3,  # 13 回退项：正文两端对齐被改动
    3,   # 14 摘要与 Abstract 标题小三黑体居中
    3,   # 15 摘要正文与 Abstract 正文四号、首行缩进、单倍行距
    1,   # 16 关键词/研究方向/Key words 三行缩进与标签字体
    3,   # 17 论文题目三号黑体加粗居中
    3,   # 18 参考文献标题小三黑体居中
    5,   # 19 13 条参考文献顶格小四 1.5 倍行距、同一段落样式
    -3,  # 20 回退项：“致  谢”标题三号黑体单倍行距各 1 行被改动
    3,   # 21 致谢正文 4 段小四宋体 1.5 倍行距两端对齐
    3,   # 22 “致  谢”由分页符另起新页
    -1,  # 23 回退项：摘要/Abstract/参考文献分节分页结构被破坏
    -5,  # 24 封面“联系电话”内容被改动
    -3,  # 25 参考文献末条与致谢间残留 ≥3 空段
    -3,  # 26 页眉文字/字号被改动或丢失
    -3,  # 27 页面尺寸或页边距被改动
    -3,  # 28 页脚页码域丢失或编号体系被改动
    -3,  # 29 封面表行列结构或文字被改动
    -3,  # 30 学术诚信声明页文字被改动
    -3,  # 31 正文 13 处参考文献标注上标/位置被改动
    -3,  # 32 摘要/Abstract/参考文献/致谢正文文字被改动
    -1,  # 33 声明页标题缩进与四段声明正文排版被改动
    -1,  # 34 封面大标题与默认样式的中文字体设定被清除
    -1,  # 35 电子签名图片丢失或被替换
    -1,  # 36 页眉页脚部件被扩增且首页/偶数页部件为空
    -1,  # 37 章节序号写法被改动或标题未单独成行
    -1,  # 38 关键词行标点被改动
]

# _MAX_DELTAS_ORDER 按 _evaluate_impl 的追加顺序排列，而细则按“先加分项、
# 后减分项”成文。此表给出追加顺序 → 细则条号（0 起）的映射，
# 使返回的 dim2_items 与细则逐条对位。
_EMIT_TO_RUBRIC = [
    15,  # 01 → 减分第 1 条
    16,  # 02
    0,   # 03 → 加分第 1 条
    17,  # 04
    1,   # 05
    2,   # 06
    18,  # 07
    3,   # 08
    4,   # 09
    5,   # 10
    6,   # 11
    19,  # 12
    20,  # 13
    7,   # 14
    8,   # 15
    9,   # 16
    10,  # 17
    11,  # 18
    12,  # 19
    21,  # 20
    13,  # 21
    14,  # 22
    22,  # 23
] + list(range(23, 38))  # 24-38 顺次接在减分项之后

assert sum(d for d in _MAX_DELTAS_ORDER if d > 0) == MAX_SCORE_TOTAL
assert sorted(_EMIT_TO_RUBRIC) == list(range(len(_MAX_DELTAS_ORDER)))


# --------------------------------------------------------------------------- #
# 文本锚点常量（全部来自任务文档的固定内容，仅用于定位与文字完整性核验）
# --------------------------------------------------------------------------- #

COVER_TITLE_TEXT = "高等教育自学考试毕业调研报告封面"
DECL_TITLE_TEXT = "学术诚信声明及作品使用授权书"
DECL_SIGN_TEXT = "毕业调研报告作者签名："
DECL_DATE_TEXT = "2026年4月20日"
TOC_TITLE_TEXT = "目录"
ABSTRACT_CN_TITLE = "摘要"
ABSTRACT_EN_TITLE = "Abstract"
PAPER_TITLE_TEXT = "城市社区雨水花园参与式维护机制研究"
REF_TITLE_TEXT = "参考文献"
THANKS_TITLE_TEXT = "致谢"

L1_TITLES = (
    "一、绪论",
    "二、城市社区雨水花园的建设背景",
    "三、雨水花园维护困境的形成机制",
    "四、参与式维护机制的构建路径",
    "五、维护机制的运行评价与优化",
    "六、保障策略与推广边界",
    "七、结论",
)

L2_TITLES = (
    "（一）产权边界与责任链条不清",
    "（二）技术标准与日常经验存在断层",
    "（三）公共空间认同不足与参与疲弱",
    "（一）建立街区—物业—居民协同责任框架",
    "（二）设计低门槛的巡检与反馈流程",
    "（三）将生态教育融入社区日常场景",
    "（一）评价指标的设计",
    "（二）动态优化的实施",
    "（一）制度保障",
    "（二）推广边界",
)

# 目录应收录的条目（一级 11 条 + 二级 10 条），均以去空白文本比对
TOC_EXPECTED = (
    (ABSTRACT_CN_TITLE,)
    + (ABSTRACT_EN_TITLE,)
    + L1_TITLES
    + (REF_TITLE_TEXT, THANKS_TITLE_TEXT)
    + L2_TITLES
)

# 封面 12 行 6 列信息表中必须原样保留的文字（去空白后做包含比对）
COVER_TABLE_TEXTS = (
    "姓名", "沈若澄", "联系电话", "专业", "公共事业管理",
    "准考证号", "050426184093", "论文题目", PAPER_TITLE_TEXT, "指导教师",
    "评判项目", "评判标准", "得分", "英文摘要（10分）", "研究能力（40分）",
    "报告撰写质量（40分）", "撰写态度与规范要求（10分）",
    "毕业论文考核成绩", "指导及答辩教师签名", "主考院校签章",
)
COVER_PHONE_TEXT = "99663529048"

# 封面表 12 行的单元格合并签名（每行各 w:tc 的 gridSpan）
COVER_SPAN_SIG = [
    (1, 2, 1, 2), (1, 2, 1, 2), (1, 5), (1, 5), (1, 4, 1), (1, 4, 1),
    (1, 4, 1), (1, 4, 1), (1, 4, 1), (2, 4), (2, 4), (6,),
]

HEADER_TEXT = "东岳技术学院高等教育自学考试毕业调研报告"

# 学术诚信声明页四段声明正文的起始特征串
DECL_BODY_HEADS = (
    "本人声明所呈交的毕业调研报告是在指导教师指导下独立完成的研究成果",
    "本人了解学习中心关于学术规范与成果提交的相关要求",
    "本人了解学习中心关于保留、查阅和使用毕业调研报告的规定",
    "（涉及不宜公开内容的报告在限定期满后适用本授权书）",
)

# 摘要 / Abstract / 致谢正文的起始特征串
ABSTRACT_CN_HEADS = (
    "随着城市更新由大拆大建逐步转向微改造与精细化治理",
    "本文以城市社区雨水花园的长期维护为研究对象",
)
ABSTRACT_EN_HEADS = (
    "With the shift of urban renewal from large-scale reconstruction",
    "A mechanism based on responsibility lists",
)
THANKS_HEADS = (
    "行文至此，调研报告的写作工作告一段落。",
    "首先，感谢在资料收集、访谈安排和文本修改过程中给予帮助的各位老师与同学。",
    "其次，感谢指导教师顾明舟老师在选题、结构安排和写作规范方面给予的耐心指导。",
    "最后，也感谢坚持完成写作的自己。",
)

# 13 条参考文献著录的作者名，用于顺序与完整性核验
REF_AUTHORS = (
    "林予珂", "贺知远", "方洛宁", "许澄怀", "周栖白", "顾明舟", "叶舒禾",
    "沈稚川", "罗以辰", "宁梓遥", "程安澜", "麦承屿", "温祈年",
)

KEYWORDS_LINE = "关键词：雨水花园；社区参与；日常管护"
RESEARCH_LINE = "研究方向：城市社区公共空间治理"
KEYWORDS_EN_LINE = "Key words: rain garden; community participation; daily maintenance"

SIGN_IMAGE_LEN = 9061
SIGN_IMAGE_CX = 961390
SIGN_IMAGE_CY = 412750

# 版面参数：除正文末节外为 11850×16783 twips，末节为标准 A4
PG_MAIN = (11850, 16783)
PG_LAST = (11906, 16838)
MAR_MAIN = {"top": 1417, "bottom": 1417, "left": 1559, "right": 1417}
MAR_LAST = {"top": 1984, "bottom": 1984, "left": 1587, "right": 1587}

SONG_NAMES = ("宋体", "SimSun", "NSimSun", "宋体-简", "STSong", "Songti")
HEI_NAMES = ("黑体", "SimHei", "STHeiti", "Heiti", "黑体-简")
TNR_NAMES = ("Times New Roman", "TimesNewRoman", "Times")


def _norm(text):
    """去掉所有空白字符，用于标题、单元格等文字的同一性比对。"""
    return re.sub(r"\s+", "", text or "")


def _attr(el, name):
    return el.get(qn(name)) if el is not None else None


def _int_attr(el, name, default=None):
    raw = _attr(el, name)
    if raw is None:
        return default
    try:
        return int(round(float(raw)))
    except (TypeError, ValueError):
        return default


def _on(el):
    """w:b / w:bCs 一类开关型元素的取值语义：缺省 w:val 视为 true。"""
    if el is None:
        return False
    val = el.get(qn("w:val"))
    if val is None:
        return True
    return str(val).strip().lower() in ("1", "true", "on")


def _near(value, target, tol):
    return value is not None and abs(value - target) <= tol


def _lines_unit(spacing, hundredths, tol=5):
    """段前/段后“N行”的达标判定：须以“行”为单位设置。

    办公软件在“行”单位下写入 beforeLines/afterLines（以 1/100 行计），这种留白随
    字号与行距自动折算；换成等值的固定磅值后，一旦字号或行距变化就不再等于 N 行，
    两者不存在换算关系，故等值的固定磅值不视为达标。
    """
    return (_near(spacing.get("before_lines"), hundredths, tol)
            and _near(spacing.get("after_lines"), hundredths, tol))


def _half_line(spacing):
    """“段前/段后各0.5行”：须按“行”设置（beforeLines/afterLines=50）。"""
    return _lines_unit(spacing, 50)


def _one_line(spacing):
    """“段前/段后各1行”：须按“行”设置（beforeLines/afterLines=100）。"""
    return _lines_unit(spacing, 100)


def _two_lines(spacing):
    """“段前/段后各2行”：须按“行”设置（beforeLines/afterLines=200）。"""
    return _lines_unit(spacing, 200)


def _indent_by_chars(ind, chars=200, tol=5):
    """“首行缩进2字符”的达标判定：须以“字符”为单位设置。

    原稿这几处的首行缩进本来就是按字符设置的，这种缩进随字号自动折算；换成等值的
    固定磅值/twips 后，字号一变缩进就不再是 2 字符，两者不存在换算关系，故等值的
    固定值不视为达标。
    """
    return _near(ind.get("first_line_chars"), chars, tol)


def _in_family(name, family):
    if not name:
        return False
    low = name.strip().lower()
    for cand in family:
        if low == cand.lower() or cand.lower() in low:
            return True
    return False


def _is_song(name):
    return _in_family(name, SONG_NAMES)


def _is_hei(name):
    return _in_family(name, HEI_NAMES)


def _is_tnr(name):
    return _in_family(name, TNR_NAMES)


def _has_cjk(text):
    return bool(re.search(r"[一-鿿]", text or ""))


def _has_latin(text):
    return bool(re.search(r"[A-Za-z0-9]", text or ""))

# --------------------------------------------------------------------------- #
# 文档解析上下文：样式索引、docDefaults、主题字体
# --------------------------------------------------------------------------- #

_A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
_WP_NS = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"


def _int_attr_plain(el, name):
    """读取无命名空间前缀的属性并转为整数（如 wp:extent 的 cx/cy）。"""
    if el is None:
        return None
    try:
        return int(round(float(el.get(name))))
    except (TypeError, ValueError):
        return None



def _load_theme_fonts(doc):
    """读取 theme1.xml 的 minor/major 字体，供 *Theme 属性回溯。

    与办公软件一致：eastAsiaTheme 取主题中 script="Hans" 的 typeface，
    该项缺失时视为未定义（不再退化成拉丁字体），以便如实反映中文字体丢失。
    """
    fonts = {}
    theme_part = None
    for rel in doc.part.rels.values():
        if rel.reltype.endswith("/theme"):
            theme_part = rel.target_part
            break
    if theme_part is None or etree is None:
        return fonts
    try:
        root = etree.fromstring(theme_part.blob)
    except Exception:
        return fonts
    for tag, latin_key, ea_key in (
        ("minorFont", "minorHAnsi", "minorEastAsia"),
        ("majorFont", "majorHAnsi", "majorEastAsia"),
    ):
        node = root.find(".//{%s}%s" % (_A_NS, tag))
        if node is None:
            continue
        latin = node.find("{%s}latin" % _A_NS)
        if latin is not None and latin.get("typeface"):
            fonts[latin_key] = latin.get("typeface")
            fonts[latin_key.replace("HAnsi", "Ascii")] = latin.get("typeface")
        for font in node.findall("{%s}font" % _A_NS):
            if font.get("script") == "Hans" and font.get("typeface"):
                fonts[ea_key] = font.get("typeface")
    return fonts

class _Ctx(object):
    """一次性建好样式索引与默认值，避免逐段重复遍历 styles.xml。"""

    def __init__(self, doc):
        self.doc = doc
        self.styles = {}
        self.default_para_style = None
        for style in doc.styles.element.findall(qn("w:style")):
            sid = style.get(qn("w:styleId"))
            if sid:
                self.styles[sid] = style
            if (style.get(qn("w:type")) == "paragraph"
                    and style.get(qn("w:default")) == "1"
                    and self.default_para_style is None):
                self.default_para_style = sid
        self.dd_rpr = None
        self.dd_ppr = None
        docdefaults = doc.styles.element.find(qn("w:docDefaults"))
        if docdefaults is not None:
            rpr_default = docdefaults.find(qn("w:rPrDefault"))
            if rpr_default is not None:
                self.dd_rpr = rpr_default.find(qn("w:rPr"))
            ppr_default = docdefaults.find(qn("w:pPrDefault"))
            if ppr_default is not None:
                self.dd_ppr = ppr_default.find(qn("w:pPr"))
        self.theme = _load_theme_fonts(doc)

    def chain(self, style_id):
        """样式自身 + basedOn 继承链（就近优先）。"""
        out = []
        seen = set()
        sid = style_id
        while sid and sid not in seen and len(out) < 12:
            seen.add(sid)
            style = self.styles.get(sid)
            if style is None:
                break
            out.append(style)
            based = style.find(qn("w:basedOn"))
            sid = based.get(qn("w:val")) if based is not None else None
        return out

    def theme_font(self, key):
        return self.theme.get(key)


def _pstyle_id(para):
    ppr = para._p.pPr
    if ppr is None:
        return None
    pstyle = ppr.find(qn("w:pStyle"))
    return pstyle.get(qn("w:val")) if pstyle is not None else None

def _ppr_sources(ctx, para):
    """段落属性查找顺序：段落 pPr → pStyle 链 → 默认段落样式链 → docDefaults。"""
    sources = []
    ppr = para._p.pPr
    if ppr is not None:
        sources.append(ppr)
    sid = _pstyle_id(para) or ctx.default_para_style
    for style in ctx.chain(sid):
        style_ppr = style.find(qn("w:pPr"))
        if style_ppr is not None:
            sources.append(style_ppr)
    if ctx.dd_ppr is not None:
        sources.append(ctx.dd_ppr)
    return sources


def _rpr_sources(ctx, para, run_el):
    """字符属性查找顺序：run rPr → rStyle 链 → 段落样式链 → docDefaults。"""
    sources = []
    rpr = run_el.find(qn("w:rPr")) if run_el is not None else None
    if rpr is not None:
        sources.append(rpr)
        rstyle = rpr.find(qn("w:rStyle"))
        if rstyle is not None:
            for style in ctx.chain(rstyle.get(qn("w:val"))):
                style_rpr = style.find(qn("w:rPr"))
                if style_rpr is not None:
                    sources.append(style_rpr)
    sid = _pstyle_id(para) if para is not None else None
    for style in ctx.chain(sid or ctx.default_para_style):
        style_rpr = style.find(qn("w:rPr"))
        if style_rpr is not None:
            sources.append(style_rpr)
    if ctx.dd_rpr is not None:
        sources.append(ctx.dd_rpr)
    return sources


def _first_child(sources, tag):
    for src in sources:
        el = src.find(qn(tag))
        if el is not None:
            return el
    return None


def _first_attr(sources, tag, names):
    """按属性逐级回溯：同一属性就近生效，不同属性可来自不同层级。"""
    for src in sources:
        el = src.find(qn(tag))
        if el is None:
            continue
        for name in names:
            val = el.get(qn(name if ":" in name else "w:" + name))
            if val not in (None, ""):
                return val
    return None

def _resolve_effective_eastasia(ctx, para, run_el):
    """解析中文（东亚）字体：显式 w:eastAsia 优先，其次 eastAsiaTheme 走主题。"""
    for src in _rpr_sources(ctx, para, run_el):
        rfonts = src.find(qn("w:rFonts"))
        if rfonts is None:
            continue
        explicit = rfonts.get(qn("w:eastAsia"))
        if explicit:
            return explicit
        themed = rfonts.get(qn("w:eastAsiaTheme"))
        if themed:
            return ctx.theme_font(themed)
    return None


def _resolve_effective_ascii(ctx, para, run_el):
    """解析西文字体：显式 w:ascii 优先，其次 asciiTheme 走主题。"""
    for src in _rpr_sources(ctx, para, run_el):
        rfonts = src.find(qn("w:rFonts"))
        if rfonts is None:
            continue
        explicit = rfonts.get(qn("w:ascii"))
        if explicit:
            return explicit
        themed = rfonts.get(qn("w:asciiTheme"))
        if themed:
            return ctx.theme_font(themed)
    return None


def _style_eastasia(ctx, style_id):
    """仅沿样式链解析中文字体，用于核验默认段落样式本身。"""
    sources = []
    for style in ctx.chain(style_id):
        style_rpr = style.find(qn("w:rPr"))
        if style_rpr is not None:
            sources.append(style_rpr)
    if ctx.dd_rpr is not None:
        sources.append(ctx.dd_rpr)
    for src in sources:
        rfonts = src.find(qn("w:rFonts"))
        if rfonts is None:
            continue
        explicit = rfonts.get(qn("w:eastAsia"))
        if explicit:
            return explicit
        themed = rfonts.get(qn("w:eastAsiaTheme"))
        if themed:
            return ctx.theme_font(themed)
    return None


def _run_effective_size_halfpt(ctx, para, run_el):
    el = _first_child(_rpr_sources(ctx, para, run_el), "w:sz")
    return _int_attr(el, "w:val")


def _run_bold(ctx, para, run_el):
    for src in _rpr_sources(ctx, para, run_el):
        bold = src.find(qn("w:b"))
        if bold is not None:
            return _on(bold)
    return False


def _run_superscript(run_el):
    rpr = run_el.find(qn("w:rPr"))
    if rpr is None:
        return False
    vert = rpr.find(qn("w:vertAlign"))
    return _attr(vert, "w:val") == "superscript"

def _para_jc(ctx, para):
    """读取对齐方式；兼容新式 start/end 写法，统一归一为 left/right。"""
    el = _first_child(_ppr_sources(ctx, para), "w:jc")
    val = _attr(el, "w:val")
    return {"start": "left", "end": "right"}.get(val, val)


def _para_indent(ctx, para, size_halfpt=None):
    """返回首行缩进（twips）、左缩进（twips）与悬挂缩进。

    firstLineChars/leftChars 以 1/100 汉字符计，1 个汉字符宽 = 半磅值 × 10 twips，
    与办公软件的“按字符设置缩进”一致。
    """
    sources = _ppr_sources(ctx, para)
    char_w = (size_halfpt or 24) * 10 / 100.0
    first = _first_attr(sources, "w:ind", ("firstLine",))
    first_chars = _first_attr(sources, "w:ind", ("firstLineChars",))
    left = _first_attr(sources, "w:ind", ("left", "start"))
    left_chars = _first_attr(sources, "w:ind", ("leftChars", "startChars"))
    hanging = _first_attr(sources, "w:ind", ("hanging",))

    def _num(raw):
        try:
            return int(round(float(raw)))
        except (TypeError, ValueError):
            return None

    first_tw = _num(first)
    if first_chars is not None:
        chars_tw = _num(first_chars)
        if chars_tw is not None:
            first_tw = int(round(chars_tw * char_w))
    left_tw = _num(left)
    if left_chars is not None:
        chars_tw = _num(left_chars)
        if chars_tw is not None:
            left_tw = int(round(chars_tw * char_w))
    hang_tw = _num(hanging)
    hang_chars = _first_attr(sources, "w:ind", ("hangingChars",))
    if hang_chars is not None:
        chars_tw = _num(hang_chars)
        if chars_tw is not None:
            hang_tw = int(round(chars_tw * char_w))
    # 与办公软件一致：w:hanging 与 w:firstLine 互斥，同时出现时悬挂缩进生效，
    # firstLine 被忽略，首行位置即 left - hanging。
    if hang_tw:
        first_tw = -hang_tw
    return {
        "first_line": max(0, (left_tw or 0) + (first_tw or 0)) if hang_tw
        else (first_tw or 0),
        "first_line_raw": _num(first),
        "first_line_chars": _num(first_chars),
        "left": left_tw or 0,
        "left_chars": _num(left_chars),
        "hanging": hang_tw or 0,
    }


def _para_spacing(ctx, para):
    """返回段前、段后（twips）与行距设置；*Lines 属性按 1/100 行折算为 twips。"""
    sources = _ppr_sources(ctx, para)

    def _num(raw):
        try:
            return int(round(float(raw)))
        except (TypeError, ValueError):
            return None

    before = _num(_first_attr(sources, "w:spacing", ("before",)))
    after = _num(_first_attr(sources, "w:spacing", ("after",)))
    before_lines = _num(_first_attr(sources, "w:spacing", ("beforeLines",)))
    after_lines = _num(_first_attr(sources, "w:spacing", ("afterLines",)))
    if before_lines:
        before = int(round(before_lines * 2.4))
    if after_lines:
        after = int(round(after_lines * 2.4))
    line = _num(_first_attr(sources, "w:spacing", ("line",)))
    rule = _first_attr(sources, "w:spacing", ("lineRule",))
    return {
        "before": before or 0,
        "after": after or 0,
        "before_lines": before_lines,
        "after_lines": after_lines,
        "line": line,
        "rule": rule or ("auto" if line else None),
    }

def _line_is(spacing, target, rule="auto", tol=12):
    """行距核验：未显式设置 w:line 时，办公软件按单倍行距（240 auto）渲染。"""
    line = spacing.get("line")
    actual_rule = spacing.get("rule")
    if line is None:
        line, actual_rule = 240, "auto"
    if (actual_rule or "auto") != rule:
        return False
    return abs(line - target) <= tol


def _para_runs(para):
    """收集段落内所有承载可见文字的 run，含 w:hyperlink 内的 run。

    与 python-docx 的 Paragraph.runs 不同：后者只返回 w:p 的直接子 run，
    自动目录条目的超链接文字会被漏掉。域代码（w:instrText）不计入。
    """
    out = []
    for run_el in para._p.iter(qn("w:r")):
        if run_el.find(qn("w:instrText")) is not None:
            continue
        text = "".join(t.text or "" for t in run_el.iter(qn("w:t")))
        if text.strip():
            out.append((run_el, text))
    return out


def _para_text(para):
    """段落可见文字：含超链接内文字，排除域代码。"""
    parts = []
    for node in para._p.iter():
        if node.tag == qn("w:t"):
            parent_run = node.getparent()
            if parent_run is not None and parent_run.find(qn("w:instrText")) is not None:
                continue
            parts.append(node.text or "")
        elif node.tag == qn("w:tab"):
            if node.getparent() is not None and node.getparent().tag == qn("w:r"):
                parts.append("\t")
    return "".join(parts)


def _iter_body_paragraphs(body, doc):
    """按文档流顺序遍历段落，并递归进入 w:sdt（自动目录内容控件）。"""
    for child in body.iterchildren():
        if child.tag == qn("w:p"):
            yield Paragraph(child, doc)
        elif child.tag == qn("w:sdt"):
            for sub in child.iterchildren():
                if sub.tag == qn("w:sdtContent"):
                    for para in _iter_body_paragraphs(sub, doc):
                        yield para


def _has_page_break(para):
    for br in para._p.iter(qn("w:br")):
        if br.get(qn("w:type")) == "page":
            return True
    return False


def _has_section_break(para):
    """段落末尾承载的分节符，缺省 w:type 即 nextPage（另起新页）。"""
    ppr = para._p.pPr
    if ppr is None:
        return False
    sect = ppr.find(qn("w:sectPr"))
    if sect is None:
        return False
    stype = sect.find(qn("w:type"))
    val = _attr(stype, "w:val")
    return val in (None, "nextPage", "oddPage", "evenPage")

def _page_break_before(ctx, para):
    el = _first_child(_ppr_sources(ctx, para), "w:pageBreakBefore")
    return _on(el)


def _right_dot_tab(ctx, para):
    """段落是否设有带点状前导符、落在版心右边界的右对齐制表位。"""
    for src in _ppr_sources(ctx, para):
        tabs = src.find(qn("w:tabs"))
        if tabs is None:
            continue
        for tab in tabs.findall(qn("w:tab")):
            if _attr(tab, "w:val") != "right":
                continue
            if _attr(tab, "w:leader") != "dot":
                continue
            pos = _int_attr(tab, "w:pos", 0)
            if pos is not None and 8800 <= pos <= 8900:
                return True
    return False


def _open_word_document(file_path):
    """打开 .docx；文件损坏或非 OOXML 时抛出异常，由维度 1 兜底。"""
    return Document(file_path)


def _effective_header(doc, index):
    """向前回溯到最近一个显式定义了默认页眉的节，模拟“同前节”继承。"""
    for i in range(index, -1, -1):
        header = doc.sections[i].header
        if not header.is_linked_to_previous:
            return header
    return doc.sections[0].header


def _effective_footer(doc, index):
    for i in range(index, -1, -1):
        footer = doc.sections[i].footer
        if not footer.is_linked_to_previous:
            return footer
    return doc.sections[0].footer


def _part_has_text(part):
    if part is None:
        return False
    try:
        return bool("".join(part.element.itertext()).strip())
    except Exception:
        return False


def _field_codes(el):
    """收集元素内所有域代码文本（w:instrText 与 w:fldSimple/@w:instr）。"""
    codes = []
    for node in el.iter(qn("w:instrText")):
        codes.append((node.text or "").strip())
    for node in el.iter(qn("w:fldSimple")):
        codes.append((_attr(node, "w:instr") or "").strip())
    return codes

# --------------------------------------------------------------------------- #
# 文档结构建模：定位 9 个组成部分与各标题分组
# --------------------------------------------------------------------------- #

def _build_model(ctx, doc):
    body = doc.element.body
    paras = list(_iter_body_paragraphs(body, doc))
    texts = [_para_text(p) for p in paras]
    ntexts = [_norm(t) for t in texts]

    sdt_pids = set()
    for sdt in body.iter(qn("w:sdt")):
        for para_el in sdt.iter(qn("w:p")):
            sdt_pids.add(id(para_el))

    def _find(target, start=0, end=None):
        stop = len(paras) if end is None else end
        for i in range(start, stop):
            if ntexts[i] == target:
                return i
        return None

    model = {"paras": paras, "texts": texts, "ntexts": ntexts, "sdt_pids": sdt_pids}
    model["cover_title"] = _find(_norm(COVER_TITLE_TEXT))
    model["decl_title"] = _find(_norm(DECL_TITLE_TEXT))
    model["decl_sign"] = _find(_norm(DECL_SIGN_TEXT))
    model["decl_date"] = _find(_norm(DECL_DATE_TEXT))
    model["toc_title"] = _find(TOC_TITLE_TEXT)

    toc_start = (model["toc_title"] + 1) if model["toc_title"] is not None else 0
    entries = []
    for i in range(toc_start, len(paras)):
        style_id = (_pstyle_id(paras[i]) or "").lower().replace(" ", "")
        style_name = (paras[i].style.name or "").lower().replace(" ", "")
        is_toc_style = bool(re.match(r"^(toc|目录)\d*$", style_id)
                            or re.match(r"^(toc|目录)\d*$", style_name))
        in_sdt = id(paras[i]._p) in sdt_pids
        if not ntexts[i]:
            continue
        if is_toc_style or in_sdt:
            entries.append(i)
            continue
        break
    model["toc_entries"] = entries
    toc_end = (entries[-1] + 1) if entries else toc_start
    model["toc_end"] = toc_end
    model["abs_cn_title"] = _find(_norm(ABSTRACT_CN_TITLE), toc_end)
    model["abs_en_title"] = _find(_norm(ABSTRACT_EN_TITLE), toc_end)
    after_abs = (model["abs_en_title"] + 1) if model["abs_en_title"] is not None else toc_end
    model["paper_title"] = _find(_norm(PAPER_TITLE_TEXT), after_abs)
    model["ref_title"] = _find(_norm(REF_TITLE_TEXT), toc_end)
    model["thanks_title"] = _find(THANKS_TITLE_TEXT, toc_end)

    model["keywords"] = None
    model["research"] = None
    model["keywords_en"] = None
    for i in range(toc_end, len(paras)):
        stripped = texts[i].strip()
        if model["keywords"] is None and stripped.startswith("关键词"):
            model["keywords"] = i
        elif model["research"] is None and stripped.startswith("研究方向"):
            model["research"] = i
        elif model["keywords_en"] is None and stripped.lower().startswith("key words"):
            model["keywords_en"] = i

    body_start = (model["paper_title"] or after_abs) + 1
    body_end = model["ref_title"] if model["ref_title"] is not None else len(paras)
    title_map = {}
    for group, titles in (("l1", L1_TITLES), ("l2", L2_TITLES)):
        for title in titles:
            idx = _find(_norm(title), toc_end)
            if idx is not None:
                title_map[title] = idx
    model["title_index"] = title_map
    model["l1_idx"] = [title_map[t] for t in L1_TITLES if t in title_map]
    model["l2_idx"] = [title_map[t] for t in L2_TITLES if t in title_map]
    skip = set(model["l1_idx"]) | set(model["l2_idx"])
    model["body_idx"] = [i for i in range(body_start, body_end)
                         if ntexts[i] and i not in skip]

    refs = []
    if model["ref_title"] is not None:
        for i in range(model["ref_title"] + 1, len(paras)):
            if re.match(r"^\[\d+\]", texts[i].strip()):
                refs.append(i)
            elif ntexts[i]:
                break
    model["ref_idx"] = refs

    thanks_body = []
    if model["thanks_title"] is not None:
        for i in range(model["thanks_title"] + 1, len(paras)):
            if ntexts[i]:
                thanks_body.append(i)
    model["thanks_body_idx"] = thanks_body

    model["abs_cn_body_idx"] = _span_nonempty(
        ntexts, model["abs_cn_title"], model["keywords"])
    model["abs_en_body_idx"] = _span_nonempty(
        ntexts, model["abs_en_title"], model["keywords_en"])

    model["blanks_before_thanks"] = _count_blanks_before(ntexts, model["thanks_title"])
    return model

def _span_nonempty(ntexts, start, end):
    """取 (start, end) 之间的非空段落下标，用于摘要正文一类的正文块。"""
    if start is None:
        return []
    stop = end if end is not None else len(ntexts)
    return [i for i in range(start + 1, min(stop, len(ntexts))) if ntexts[i]]


def _count_blanks_before(ntexts, index):
    if index is None:
        return 0
    count = 0
    for i in range(index - 1, -1, -1):
        if ntexts[i]:
            break
        count += 1
    return count


def _para_mark_size(ctx, para, default=24):
    """空段落的有效字号（半磅）：段落标记 rPr → 样式链 → docDefaults。

    空段落没有 run，字号由 w:pPr/w:rPr 里的段落标记属性决定，
    不能沿用正文 run 的读取路径。
    """
    ppr = para._p.pPr
    if ppr is not None:
        mark_rpr = ppr.find(qn("w:rPr"))
        if mark_rpr is not None:
            val = _int_attr(mark_rpr.find(qn("w:sz")), "w:val")
            if val:
                return val
    val = _int_attr(_first_child(_rpr_sources(ctx, para, None), "w:sz"), "w:val")
    return val or default


def _para_height(ctx, para):
    """估算段落占据的垂直高度（twips）：行高 + 段前 + 段后。

    与办公软件一致：auto 型行距按“字号磅值 × w:line / 240”折算；
    exact/atLeast 型行距直接取 w:line（其单位即 1/20 磅）。
    """
    spacing = _para_spacing(ctx, para)
    size = _para_mark_size(ctx, para)
    line, rule = spacing["line"], spacing["rule"]
    if line is None:
        height = int(round(size / 2.0 * 20))
    elif (rule or "auto") == "auto":
        height = int(round(size / 2.0 * 20 * line / 240.0))
    else:
        height = line
    return height + spacing["before"] + spacing["after"]


def _section_body_height(section, default=13949):
    """节的版心高度（twips）= 页高 − 上页边距 − 下页边距。"""
    sect_pr = section._sectPr
    height = _int_attr(sect_pr.find(qn("w:pgSz")), "w:h")
    if not height:
        return default
    pg_mar = sect_pr.find(qn("w:pgMar"))
    body = height - _int_attr(pg_mar, "w:top", 0) - _int_attr(pg_mar, "w:bottom", 0)
    return body if body > 0 else default

# --------------------------------------------------------------------------- #
# 分组核验小工具
# --------------------------------------------------------------------------- #

def _runs_ok(ctx, para, size=None, bold=None, ea_pred=None, latin_pred=None,
             size_tol=1, skip_sup=False):
    """段落内每个可见 run 都需满足给定的字号/加粗/中西文字体要求。

    skip_sup=True 时跳过上标 run：正文中的参考文献标注以上标形式排版，
    其字号字体由标注专项条目单独核验，不参与正文字号字体的一致性判定。
    """
    runs = _para_runs(para)
    if not runs:
        return False
    for run_el, text in runs:
        if skip_sup and _run_superscript(run_el):
            continue
        if size is not None:
            if not _near(_run_effective_size_halfpt(ctx, para, run_el), size, size_tol):
                return False
        if bold is not None:
            if _run_bold(ctx, para, run_el) is not bool(bold):
                return False
        if ea_pred is not None and _has_cjk(text):
            if not ea_pred(_resolve_effective_eastasia(ctx, para, run_el)):
                return False
        if latin_pred is not None and _has_latin(text):
            if not latin_pred(_resolve_effective_ascii(ctx, para, run_el)):
                return False
    return True


def _first_run_size(ctx, para, default=24):
    for run_el, _text in _para_runs(para):
        size = _run_effective_size_halfpt(ctx, para, run_el)
        if size:
            return size
    return default


def _all_ok(indices, func):
    return bool(indices) and all(func(i) for i in indices)


def _toc_label(text):
    """目录条目的标题部分：去掉制表位与右端页码。"""
    head = text.split("\t")[0]
    head = re.sub(r"[.．·\s]*[0-9IVXivx]+\s*$", "", head)
    return _norm(head)

# --------------------------------------------------------------------------- #
# 维度 1：可用与可修改性
# --------------------------------------------------------------------------- #

def _check_dim1(ctx, doc, model):
    """交付文件须为可编辑的完整报告：9 个组成部分齐全、正文文字未被改写。"""
    missing = []
    required = (
        ("封面", model["cover_title"]),
        ("学术诚信声明及作品使用授权书", model["decl_title"]),
        ("目录", model["toc_title"]),
        ("摘要", model["abs_cn_title"]),
        ("Abstract", model["abs_en_title"]),
        ("论文题目", model["paper_title"]),
        ("参考文献", model["ref_title"]),
        ("致谢", model["thanks_title"]),
    )
    for name, idx in required:
        if idx is None:
            missing.append(name)
    if len(model["l1_idx"]) < len(L1_TITLES):
        missing.append("一至七章正文标题")
    if missing:
        return False, "缺失组成部分：" + "、".join(missing)

    if not doc.tables:
        return False, "封面信息表缺失，封面疑被整页转为图片"
    if len(model["body_idx"]) < 30:
        return False, "一至七章正文段落数量为 %d，正文疑被删减" % len(model["body_idx"])
    if len(model["ref_idx"]) != len(REF_AUTHORS):
        return False, "参考文献著录 %d 条，与 13 条不符" % len(model["ref_idx"])
    if len(model["thanks_body_idx"]) < len(THANKS_HEADS):
        return False, "致谢正文段落数量为 %d，少于 4 段" % len(model["thanks_body_idx"])

    labels = [_toc_label(model["texts"][i]) for i in model["toc_entries"]]
    expected = set(_norm(t) for t in TOC_EXPECTED)
    matched = len([lab for lab in labels if lab in expected])
    if not labels:
        return False, "目录条目缺失"
    if matched * 2 < len(expected):
        return False, "目录条目与正文标题基本对不上（匹配 %d/%d）" % (matched, len(expected))
    return True, ""

# --------------------------------------------------------------------------- #
# 维度 2：完成度 38 条
# --------------------------------------------------------------------------- #

def _evaluate_impl(file_path, result):
    doc = _open_word_document(file_path)
    ctx = _Ctx(doc)
    model = _build_model(ctx, doc)
    body = doc.element.body
    paras = model["paras"]
    texts = model["texts"]
    ntexts = model["ntexts"]
    entries = model["toc_entries"]
    sections = doc.sections

    dim1_pass, dim1_reason = _check_dim1(ctx, doc, model)
    result["dim1_pass"] = dim1_pass
    result["dim1_reason"] = dim1_reason

    results = []

    # 01 目录的自动更新 TOC 域被破坏（回退项 -5）
    # 严格对齐细则，仅踩细则明确写出的点：①“目录”内容控件（w:sdt）内含 TOC 域；
    # ②域开关含 \o "1-3" 与 \h；③每条条目为可 Ctrl+单击跳转的书签超链接。
    # 与办公软件对齐：条目跳转既可由 w:hyperlink/@w:anchor 承载，
    # 也可由 HYPERLINK \l + PAGEREF 域承载，两者等效。
    toc_field = ""
    toc_in_sdt = False
    for sdt in body.iter(qn("w:sdt")):
        for code in _field_codes(sdt):
            if code.upper().startswith("TOC"):
                toc_field, toc_in_sdt = code, True
    if not toc_field:
        for code in _field_codes(body):
            if code.upper().startswith("TOC"):
                toc_field = code

    def _entry_clickable(idx):
        para = paras[idx]
        for hyperlink in para._p.iter(qn("w:hyperlink")):
            if hyperlink.get(qn("w:anchor")):
                return True
        codes = " ".join(_field_codes(para._p)).upper()
        return "HYPERLINK \\L" in codes or "PAGEREF" in codes

    hit = not (toc_in_sdt and "\\o" in toc_field and "1-3" in toc_field
               and "\\h" in toc_field and _all_ok(entries, _entry_clickable))
    rule = ("目录的自动更新能力被破坏——“目录”内容控件内的自动目录域（TOC 域）丢失，"
            "或条目被拆成逐行手动敲入的普通文字，或条目的书签超链接失效，"
            "导致无法用F9或“更新目录”整体刷新")
    results.append(("-5：" if hit else "-0：") + rule)
    # 02 目录条目出现缺条、重复或无关条目（回退项 -3）
    # 严格对齐细则：一级 11 条（摘要、Abstract、一至七章、参考文献、致谢）
    # 与二级 10 条，须无缺条、无重复。
    # 与格式规范范例一致：论文题目条目在目录中出现属规范允许，不计为无关条目。
    labels = [_toc_label(texts[i]) for i in entries]
    expected_labels = [_norm(t) for t in TOC_EXPECTED]
    allowed_labels = expected_labels + [_norm(PAPER_TITLE_TEXT)]
    hit = not (len(labels) == len(set(labels))
               and all(x in labels for x in expected_labels)
               and all(x in allowed_labels for x in labels))
    rule = ("目录条目出现缺条、重复或无关条目，不再完整覆盖“摘要”“Abstract”“一、绪论”"
            "“二、城市社区雨水花园的建设背景”“三、雨水花园维护困境的形成机制”"
            "“四、参与式维护机制的构建路径”“五、维护机制的运行评价与优化”“六、保障策略与推广边界”"
            "“七、结论”“参考文献”“致  谢”，"
            "以及“（一）产权边界与责任链条不清”至“（二）推广边界”共10条二级标题")
    results.append(("-3：" if hit else "-0：") + rule)

    # 03 目录条目按层级缩进（+3）
    # 与办公软件对齐：条目缩进可写在段落 w:ind，也可由 toc 1/toc 2 样式给出，
    # 故按“段落 → 样式链 → docDefaults”解析后的有效左缩进核验。
    def _entry_indent_ok(idx):
        label = _toc_label(texts[idx])
        level = 2 if re.match(r"^（[一二三四五六七八九十]+）", label) else 1
        size = _first_run_size(ctx, paras[idx], 28)
        left = _para_indent(ctx, paras[idx], size)["left"]
        if level == 1:
            return _near(left, 0, 40)
        return _near(left, 280, 40)

    hit = _all_ok(entries, _entry_indent_ok)
    rule = ("目录条目按层级区分缩进——一级条目顶格（左缩进0 twips），"
            "二级条目缩进1个汉字符（280 twips±40），三级条目缩进2个汉字符（560 twips±60）")
    results.append(("+3：" if hit else "+0：") + rule)

    # 04 目录右对齐点状前导符制表位被移除（回退项 -3）
    hit = not _all_ok(entries, lambda i: _right_dot_tab(ctx, paras[i]))
    rule = ("目录条目右端的右对齐制表位或点状前导符（制表位落在版心右边界"
            "8800-8900 twips）被移除，页码不再沿右边界对齐成一竖列")
    results.append(("-3：" if hit else "-0：") + rule)

    # 05 目录条目字体字号行距（+3）
    def _entry_font_ok(idx):
        para = paras[idx]
        if not _runs_ok(ctx, para, size=28, ea_pred=_is_song, latin_pred=_is_tnr):
            return False
        return _line_is(_para_spacing(ctx, para), 360, "auto")

    hit = _all_ok(entries, _entry_font_ok)
    rule = ("目录条目文字中文为四号宋体（14磅±0.5），英文与阿拉伯数字为Times New Roman，"
            "1.5倍行距")
    results.append(("+3：" if hit else "+0：") + rule)
    # 06 “目  录”标题二号黑体居中（+3）
    idx = model["toc_title"]
    hit = False
    if idx is not None:
        para = paras[idx]
        spacing = _para_spacing(ctx, para)
        hit = (_runs_ok(ctx, para, size=44, ea_pred=_is_hei)
               and _para_jc(ctx, para) == "center"
               and _half_line(spacing)
               and _line_is(spacing, 480, "auto"))
    rule = ("“目  录”标题为二号（22磅±0.5）黑体、居中，"
            "段前段后各0.5行，两倍行距")
    results.append(("+3：" if hit else "+0：") + rule)

    # 07 目录不再独立成页或目录页混入正文文字（回退项 -1）
    # 与办公软件对齐：另起新页既可由分页符（w:br type=page / pageBreakBefore）实现，
    # 也可由 nextPage 型分节符实现，两者等效。
    # “摘要”另起新页归“摘要/Abstract/参考文献分节分页结构”一条判定，此处不重复扣分。
    hit = True
    toc_idx, abs_idx = model["toc_title"], model["abs_cn_title"]
    if toc_idx is not None and abs_idx is not None:
        entry_set = set(entries)
        clean = all((not ntexts[i]) or (i in entry_set)
                    for i in range(toc_idx + 1, abs_idx))
        before_break = any(_has_page_break(paras[i]) or _has_section_break(paras[i])
                           for i in range(max(0, toc_idx - 4), toc_idx)) \
            or _page_break_before(ctx, paras[toc_idx])
        hit = not (clean and before_break)
    rule = "目录不再独立占一整页，或目录页内混入正文文字"
    results.append(("-1：" if hit else "-0：") + rule)

    # 08 7 个一级标题小三黑体居中顶格（+5）
    def _l1_font_ok(idx):
        para = paras[idx]
        ind = _para_indent(ctx, para, 30)
        return (_runs_ok(ctx, para, size=30, ea_pred=_is_hei)
                and _para_jc(ctx, para) == "center"
                and _near(ind["first_line"], 0, 40)
                and _near(ind["left"], 0, 40))

    l1_all = len(model["l1_idx"]) == len(L1_TITLES)
    hit = l1_all and _all_ok(model["l1_idx"], _l1_font_ok)
    rule = ("“一、绪论”“二、城市社区雨水花园的建设背景”“三、雨水花园维护困境的形成机制”"
            "“四、参与式维护机制的构建路径”“五、维护机制的运行评价与优化”“六、保障策略与推广边界”"
            "“七、结论”共7个一级标题统一为小三号（15磅±0.5）黑体、居中、"
            "无首行缩进（首行缩进与左缩进均为0）")
    results.append(("+5：" if hit else "+0：") + rule)

    # 09 一级标题段前段后 0.5 行、两倍行距（+3）
    def _l1_spacing_ok(idx):
        spacing = _para_spacing(ctx, paras[idx])
        return (_half_line(spacing)
                and _line_is(spacing, 480, "auto"))

    hit = l1_all and _all_ok(model["l1_idx"], _l1_spacing_ok)
    rule = "上述7个一级标题段前段后各0.5行，两倍行距"
    results.append(("+3：" if hit else "+0：") + rule)

    # 10 10 个二级标题四号黑体加粗顶格、同一段落样式（+5）
    # 与办公软件对齐：字号、缩进、行距一致即视为同一套格式，此处再核验段落样式
    # （pStyle）也一致，避免同样的版面由两套样式拼出。
    def _l2_font_ok(idx):
        para = paras[idx]
        ind = _para_indent(ctx, para, 28)
        return (_runs_ok(ctx, para, size=28, bold=True, ea_pred=_is_hei)
                and _near(ind["first_line"], 0, 40)
                and _near(ind["left"], 0, 40))

    l2_all = len(model["l2_idx"]) == len(L2_TITLES)
    hit = (l2_all and _all_ok(model["l2_idx"], _l2_font_ok)
           and len(set(_pstyle_id(paras[i]) for i in model["l2_idx"])) == 1)
    rule = ("“（一）产权边界与责任链条不清”“（二）技术标准与日常经验存在断层”"
            "“（三）公共空间认同不足与参与疲弱”“（一）建立街区—物业—居民协同责任框架”"
            "“（二）设计低门槛的巡检与反馈流程”“（三）将生态教育融入社区日常场景”"
            "“（一）评价指标的设计”“（二）动态优化的实施”“（一）制度保障”“（二）推广边界”"
            "共10个二级标题统一为四号（14磅±0.5）黑体、加粗、顶格（首行缩进与左缩进均为0），"
            "且10条共用同一段落样式，其中“（一）评价指标的设计”“（二）动态优化的实施”"
            "“（一）制度保障”“（二）推广边界”这4条不再与其余6条呈现两套不同的字号或缩进")
    results.append(("+5：" if hit else "+0：") + rule)

    # 11 二级标题段前段后 0.5 行、单倍行距（+3）
    def _l2_spacing_ok(idx):
        spacing = _para_spacing(ctx, paras[idx])
        return (_half_line(spacing)
                and _line_is(spacing, 240, "auto"))

    hit = l2_all and _all_ok(model["l2_idx"], _l2_spacing_ok)
    rule = "上述10个二级标题段前段后各0.5行，单倍行距"
    results.append(("+3：" if hit else "+0：") + rule)

    # 12 正文小四宋体/TNR、首行缩进、1.5 倍行距被改动（回退项 -5）
    def _body_font_ok(idx):
        para = paras[idx]
        ind = _para_indent(ctx, para, 24)
        return (_runs_ok(ctx, para, size=24, ea_pred=_is_song, latin_pred=_is_tnr,
                         skip_sup=True)
                and _near(ind["first_line"], 480, 30)
                and _line_is(_para_spacing(ctx, para), 360, "auto"))

    hit = not (len(model["body_idx"]) >= 30
               and _all_ok(model["body_idx"], _body_font_ok))
    rule = ("一至七章正文段落（约50段）的小四号（12磅±0.5）、中文宋体、"
            "英文与阿拉伯数字Times New Roman、首行缩进2字符（480 twips±30）、"
            "1.5倍行距被改动")
    results.append(("-5：" if hit else "-0：") + rule)

    # 13 正文两端对齐被改动（回退项 -3）
    hit = not (len(model["body_idx"]) >= 30 and _all_ok(
        model["body_idx"], lambda i: _para_jc(ctx, paras[i]) == "both"))
    rule = "一至七章正文段落的两端对齐被改为左对齐或被清除"
    results.append(("-3：" if hit else "-0：") + rule)

    # 14 摘要与 Abstract 标题小三黑体居中（+3）
    def _abs_title_ok(idx):
        if idx is None:
            return False
        para = paras[idx]
        spacing = _para_spacing(ctx, para)
        return (_runs_ok(ctx, para, size=30, ea_pred=_is_hei, latin_pred=lambda n: True)
                and _para_jc(ctx, para) == "center"
                and _half_line(spacing)
                and _line_is(spacing, 480, "auto"))

    hit = _abs_title_ok(model["abs_cn_title"]) and _abs_title_ok(model["abs_en_title"])
    rule = ("“摘要”与“Abstract”两个标题均为小三号（15磅±0.5）黑体、居中、"
            "段前段后各0.5行、两倍行距")
    results.append(("+3：" if hit else "+0：") + rule)
    # 15 摘要正文与 Abstract 正文四号、首行缩进、单倍行距（+3）
    def _abs_body_ok(idx):
        para = paras[idx]
        ind = _para_indent(ctx, para, 28)
        return (_runs_ok(ctx, para, size=28, ea_pred=_is_song, latin_pred=_is_tnr)
                and _indent_by_chars(ind)
                and _near(ind["first_line"], 560, 40)
                and _line_is(_para_spacing(ctx, para), 240, "auto"))

    abs_body = model["abs_cn_body_idx"] + model["abs_en_body_idx"]
    hit = bool(model["abs_cn_body_idx"]) and bool(model["abs_en_body_idx"]) \
        and _all_ok(abs_body, _abs_body_ok)
    rule = ("摘要正文与Abstract正文为四号（14磅±0.5），中文宋体、英文与阿拉伯数字"
            "Times New Roman，首行缩进2字符（560 twips±40），单倍行距，"
            "Abstract正文不再沿用原稿的固定行距")
    results.append(("+3：" if hit else "+0：") + rule)

    # 16 关键词/研究方向/Key words 三行缩进与标签字体（+1）
    def _kw_line_ok(idx, label):
        if idx is None:
            return False
        para = paras[idx]
        ind = _para_indent(ctx, para, 28)
        # “缩进2字符”须按字符设置，折出的宽度按标签四号（560 twips）或按正文小四
        # （480 twips）均达标；等值的固定磅值不算。
        if not _indent_by_chars(ind):
            return False
        if not (_near(ind["first_line"], 560, 40) or _near(ind["first_line"], 480, 40)):
            return False
        runs = _para_runs(para)
        if not runs:
            return False
        label_ok = False
        for run_el, text in runs:
            if not _norm(text):
                continue
            size = _run_effective_size_halfpt(ctx, para, run_el)
            bold = _run_bold(ctx, para, run_el)
            font_ok = True
            if _has_cjk(text):
                font_ok = _is_hei(_resolve_effective_eastasia(ctx, para, run_el))
            label_ok = _near(size, 28, 1) and bold and font_ok
            break
        return label_ok and label in _norm(texts[idx])

    kw_idx = model["keywords"]
    # 关键词行的分号与行尾标点归“关键词行标点被改动”一条判定，此处只判缩进与标签字体。
    hit = (_kw_line_ok(kw_idx, _norm(KEYWORDS_LINE))
           and _kw_line_ok(model["research"], _norm(RESEARCH_LINE))
           and _kw_line_ok(model["keywords_en"], _norm(KEYWORDS_EN_LINE)))
    rule = ("“关键词：雨水花园；社区参与；日常管护”“研究方向：城市社区公共空间治理”"
            "“Key words: rain garden; community participation; daily maintenance”"
            "三行缩进2字符（按四号560或小四480 twips折算，±40），其中“关键词”“研究方向”“Key words”"
            "标签为四号黑体加粗")
    results.append(("+1：" if hit else "+0：") + rule)
    # 17 论文题目三号黑体加粗居中（+3）
    idx = model["paper_title"]
    hit = False
    if idx is not None:
        para = paras[idx]
        spacing = _para_spacing(ctx, para)
        hit = (_runs_ok(ctx, para, size=32, bold=True, ea_pred=_is_hei)
               and _para_jc(ctx, para) == "center"
               and _line_is(spacing, 480, "auto")
               and _two_lines(spacing))
    rule = ("论文题目“城市社区雨水花园参与式维护机制研究”（正文首页居中单独一行）为三号"
            "（16磅±0.5）黑体、加粗、居中，行距2倍，"
            "段前段后各2行")
    results.append(("+3：" if hit else "+0：") + rule)

    # 18 参考文献标题小三黑体居中（+3）
    idx = model["ref_title"]
    hit = False
    ref_title_spacing = None
    if idx is not None:
        para = paras[idx]
        ref_title_spacing = _para_spacing(ctx, para)
        hit = (_runs_ok(ctx, para, size=30, ea_pred=_is_hei)
               and _para_jc(ctx, para) == "center"
               and _half_line(ref_title_spacing)
               and _line_is(ref_title_spacing, 480, "auto"))
    rule = ("“参考文献”标题为小三号（15磅±0.5）黑体、居中、段前段后各0.5行、"
            "两倍行距")
    results.append(("+3：" if hit else "+0：") + rule)

    # 19 13 条参考文献顶格小四 1.5 倍行距、同一段落样式（+5）
    def _ref_ok(idx):
        para = paras[idx]
        ind = _para_indent(ctx, para, 24)
        # 顶格 = 首行起点在 0：既可为 left/firstLine 全 0，
        # 也可为著录惯用的悬挂缩进（left 与 hanging 相等，首行仍顶格）。
        flush = _near(ind["first_line"], 0, 40) and (
            _near(ind["left"], 0, 40) or _near(ind["left"], ind["hanging"], 40))
        return (_runs_ok(ctx, para, size=24, ea_pred=_is_song, latin_pred=_is_tnr)
                and flush
                and _para_jc(ctx, para) in (None, "left", "both")
                and _line_is(_para_spacing(ctx, para), 360, "auto"))

    # 与办公软件对齐：“共用同一段落样式”按有效格式判定，不比对 w:pStyle 字面值。
    # 在 Word 中选中 13 条统一设置缩进与行距并不会改动 pStyle，段落直接 pPr 会覆盖
    # 样式里的 w:ind，此时两组 pStyle 不同而排版完全一致，属同一套格式；
    # 反之原稿“[1][2][4]-[10]”left=6 而“[3][11][12][13]”left=0 的两套格式，
    # 在有效格式签名下仍分成两组，照样能如实识别。
    def _ref_signature(idx):
        para = paras[idx]
        spacing = _para_spacing(ctx, para)
        ind = _para_indent(ctx, para, 24)
        return (ind["left"], ind["first_line"], ind["hanging"],
                spacing["line"], spacing["rule"],
                spacing["before"], spacing["after"], _para_jc(ctx, para))

    refs_all = len(model["ref_idx"]) == len(REF_AUTHORS)
    hit = (refs_all and _all_ok(model["ref_idx"], _ref_ok)
           and len(set(_ref_signature(i) for i in model["ref_idx"])) == 1)
    rule = ("“[1] 林予珂”至“[13] 温祈年”共13条参考文献著录统一为顶格左对齐（首行起点为0，"
            "即首行缩进与左缩进均为0，或左缩进与悬挂缩进相等的悬挂缩进），"
            "小四号（12磅±0.5），中文宋体、英文与阿拉伯数字Times New Roman，"
            "1.5倍行距，且13条共用同一段落样式，"
            "不再出现“[1][2][4][5][6][7][8][9][10]”带列表左缩进（左缩进6 twips）"
            "而“[3][11][12][13]”无缩进、两组行距不一致的两套格式")
    results.append(("+5：" if hit else "+0：") + rule)
    # 20 “致  谢”标题三号黑体单倍行距各 1 行被改动（回退项 -3）
    idx = model["thanks_title"]
    hit = True
    if idx is not None:
        para = paras[idx]
        spacing = _para_spacing(ctx, para)
        hit = not (_runs_ok(ctx, para, size=32, ea_pred=_is_hei)
                   and _para_jc(ctx, para) == "center"
                   and _line_is(spacing, 240, "auto")
                   and _one_line(spacing))
    rule = ("“致  谢”标题的三号（16磅±0.5）黑体、居中、单倍行距、"
            "段前段后各1行被改动")
    results.append(("-3：" if hit else "-0：") + rule)

    # 21 致谢正文 4 段小四宋体 1.5 倍行距两端对齐（+3）
    def _thanks_body_ok(idx):
        para = paras[idx]
        ind = _para_indent(ctx, para, 24)
        return (_runs_ok(ctx, para, size=24, ea_pred=_is_song, latin_pred=_is_tnr)
                and _indent_by_chars(ind)
                and _near(ind["first_line"], 480, 30)
                and _line_is(_para_spacing(ctx, para), 360, "auto")
                and _para_jc(ctx, para) == "both")

    thanks_body = model["thanks_body_idx"]
    hit = len(thanks_body) >= len(THANKS_HEADS) and _all_ok(thanks_body, _thanks_body_ok)
    rule = ("致谢正文4段（“行文至此，调研报告的写作工作告一段落。”起）为小四号（12磅）宋体、"
            "首行缩进2字符（480 twips±30）、1.5倍行距、两端对齐，"
            "不再沿用原稿的固定行距")
    results.append(("+3：" if hit else "+0：") + rule)

    # 22 “致  谢”由分页符另起新页（+3）
    # 与办公软件对齐：分页符可写成段中 w:br type="page"，也可写成段落属性
    # w:pageBreakBefore；仅靠连续空段落顶页不算。
    def _hard_break_before(index):
        if index is None:
            return False
        if _page_break_before(ctx, paras[index]):
            return True
        j = index - 1
        while j >= 0:
            if _has_page_break(paras[j]):
                return True
            if ntexts[j]:
                return False
            j -= 1
        return False

    def _any_break_before(index):
        if index is None:
            return False
        if _page_break_before(ctx, paras[index]):
            return True
        j = index - 1
        while j >= 0:
            if _has_page_break(paras[j]) or _has_section_break(paras[j]):
                return True
            if ntexts[j]:
                return False
            j -= 1
        return False

    hit = _hard_break_before(model["thanks_title"])
    rule = "“致  谢”由分页符另起新页，而非仅依靠连续空段落顶到下一页"
    results.append(("+3：" if hit else "+0：") + rule)

    # 23 摘要/Abstract/参考文献分节分页结构被破坏（回退项 -1）
    hit = not (_any_break_before(model["abs_cn_title"])
               and _any_break_before(model["abs_en_title"])
               and _any_break_before(model["ref_title"]))
    rule = ("“摘要”“Abstract”“参考文献”各自另起新页的分节符或分页符结构被破坏，"
            "摘要与Abstract不再分别独立起页")
    results.append(("-1：" if hit else "-0：") + rule)
    # 24 封面“联系电话”内容被改动（-5）
    table = doc.tables[0] if doc.tables else None
    grid_texts = []
    if table is not None:
        for row in table.rows:
            row_texts = []
            for cell in row.cells:
                value = _norm(cell.text)
                if not row_texts or row_texts[-1] != value:
                    row_texts.append(value)
            grid_texts.append(row_texts)
    phone_value = None
    for row_texts in grid_texts:
        for j, value in enumerate(row_texts):
            if value == "联系电话":
                for k in range(j + 1, len(row_texts)):
                    if row_texts[k]:
                        phone_value = row_texts[k]
                        break
                break
        if phone_value is not None:
            break
    hit = phone_value != COVER_PHONE_TEXT
    rule = "封面12行6列信息表中“联系电话”一栏的内容不为11位数字“99663529048”"
    results.append(("-5：" if hit else "-0：") + rule)

    # 25 参考文献末条与致谢间残留 ≥3 空段（-3）
    # 按细则原文的落脚点“产生完全空白的多余页面”判定，需同时满足三件事：
    #   ①空段落数 ≥3；②与分页符叠加（否则“致谢”仍接在同一页，不构成空白页）；
    #   ③空段落累计高度足以在页面上留出大片空白（≥版心高度的 1/4）。
    # 加入③是因为：仅剩 3 个空段时，分页符已把“致谢”顶到新页，3 个空段留在
    # 参考文献页尾部，版面上看不出多余空白页；若无此判据，把原稿 24 个空段
    # 压缩到 3 个并补上分页符（正是本项要引导的改法）反而会被扣分。
    thanks_i = model["thanks_title"]
    blanks = model["blanks_before_thanks"]
    break_after_refs = False
    if thanks_i is not None:
        if _page_break_before(ctx, paras[thanks_i]):
            break_after_refs = True
        for j in range(thanks_i - blanks - 1, thanks_i + 1):
            if 0 <= j < len(paras) and _has_page_break(paras[j]):
                break_after_refs = True
                break
    blank_height = 0
    if thanks_i is not None:
        for j in range(max(0, thanks_i - blanks), thanks_i):
            blank_height += _para_height(ctx, paras[j])
    # 致谢位于文档末节，故以末节版心高度为参照。
    body_height = _section_body_height(sections[-1]) if len(sections) else 13949
    hit = (blanks >= 3 and break_after_refs
           and blank_height >= body_height / 4.0)
    rule = "“参考文献”末条与“致  谢”标题之间残留≥3个连续空白段落，与分页符叠加产生完全空白的多余页面"
    results.append(("-3：" if hit else "-0：") + rule)

    # 26 页眉文字/字号被改动或丢失（-3）
    # 与办公软件对齐：未显式声明 headerReference 的节按“同前节”继承，
    # 故逐节回溯到最近一个显式定义的默认页眉再比对。
    header_bad = False
    for i in range(len(sections)):
        header = _effective_header(doc, i)
        text = _norm("".join(p.text for p in header.paragraphs))
        if text != _norm(HEADER_TEXT):
            header_bad = True
            break
        for para in header.paragraphs:
            for run_el, _text in _para_runs(para):
                if not _near(_run_effective_size_halfpt(ctx, para, run_el), 21, 1):
                    header_bad = True
                    break
            if header_bad:
                break
        if header_bad:
            break
    hit = header_bad
    rule = ("页眉不为“东岳技术学院高等教育自学考试毕业调研报告”（五号，10.5磅，居中），"
            "或该页眉在任一节中丢失、被改写为“毕业论文”等其他字样")
    results.append(("-3：" if hit else "-0：") + rule)

    # 27 页面尺寸或页边距被改动（-3）
    layout_bad = False
    for i, section in enumerate(sections):
        last = (i == len(sections) - 1)
        want_size = PG_LAST if last else PG_MAIN
        want_mar = MAR_LAST if last else MAR_MAIN
        sect_pr = section._sectPr
        pg_sz = sect_pr.find(qn("w:pgSz"))
        pg_mar = sect_pr.find(qn("w:pgMar"))
        if (_int_attr(pg_sz, "w:w"), _int_attr(pg_sz, "w:h")) != want_size:
            layout_bad = True
            break
        for key, value in want_mar.items():
            if _int_attr(pg_mar, "w:" + key) != value:
                layout_bad = True
                break
        if layout_bad:
            break
    hit = layout_bad
    rule = ("页面尺寸不为宽11850×高16783 twips（约20.9cm×29.6cm，正文末节为11906×16838 twips即标准A4），"
            "或页边距不为上下各1417 twips（2.5cm）、左1559 twips（2.75cm）、右1417 twips（2.5cm）"
            "（正文末节为上下各1984 twips、左右各1587 twips）")
    results.append(("-3：" if hit else "-0：") + rule)
    # 28 页脚页码域丢失或编号体系被改动（-3）
    # 与办公软件对齐：封面页与声明页所在节页脚不含 PAGE 域即为“无页码”；
    # 目录/摘要/Abstract 所在各节由 pgNumType upperRoman start=1 给出大写罗马数字；
    # 末节页码为阿拉伯数字（decimal，或省略 w:fmt 时按 decimal 渲染）。
    def _footer_has_page(index):
        footer = _effective_footer(doc, index)
        return "PAGE" in "".join(_field_codes(footer._element)).upper()

    numbering_bad = False
    if len(sections) < 2:
        numbering_bad = True
    else:
        if _footer_has_page(0):
            numbering_bad = True
        for i in range(len(sections) - 1):
            pg_num = sections[i]._sectPr.find(qn("w:pgNumType"))
            if (_attr(pg_num, "w:fmt") != "upperRoman"
                    or _int_attr(pg_num, "w:start") != 1):
                numbering_bad = True
                break
        if not any(_footer_has_page(i) for i in range(1, len(sections) - 1)):
            numbering_bad = True
        last = len(sections) - 1
        last_num = sections[last]._sectPr.find(qn("w:pgNumType"))
        last_fmt = _attr(last_num, "w:fmt")
        if not _footer_has_page(last) or last_fmt not in (None, "decimal"):
            numbering_bad = True
    hit = numbering_bad
    rule = ("页脚页码域丢失或编号体系被改动——封面页与学术诚信声明页无页码，"
            "目录/摘要/Abstract所在节页码为大写罗马数字（页码域，自I起编），"
            "论文题目至致谢所在节页码为阿拉伯数字")
    results.append(("-3：" if hit else "-0：") + rule)

    # 29 封面表行列结构或文字被改动（-3）
    structure_ok = False
    if table is not None:
        grid = table._tbl.find(qn("w:tblGrid"))
        cols = len(grid.findall(qn("w:gridCol"))) if grid is not None else 0
        span_sig = []
        for tr in table._tbl.findall(qn("w:tr")):
            row_sig = []
            for tc in tr.findall(qn("w:tc")):
                tc_pr = tc.find(qn("w:tcPr"))
                span = 1
                if tc_pr is not None:
                    gs = tc_pr.find(qn("w:gridSpan"))
                    if gs is not None:
                        span = _int_attr(gs, "w:val", 1)
                row_sig.append(span)
            span_sig.append(tuple(row_sig))
        structure_ok = (cols == 6 and len(table.rows) == 12
                        and span_sig == COVER_SPAN_SIG)
    flat = "".join("".join(row) for row in grid_texts)
    texts_ok = all(_norm(item) in flat for item in COVER_TABLE_TEXTS)
    hit = not (structure_ok and texts_ok)
    rule = ("封面12行6列信息表的行列数、单元格合并结构被改变，或“姓名/沈若澄”“专业/公共事业管理”"
            "“准考证号/050426184093”“论文题目/城市社区雨水花园参与式维护机制研究”“指导教师”"
            "“评判项目/评判标准/得分”“英文摘要（10分）”“研究能力（40分）”“报告撰写质量（40分）”"
            "“撰写态度与规范要求（10分）”“毕业论文考核成绩”“指导及答辩教师签名”“主考院校签章”"
            "等文字被改写或删除")
    results.append(("-3：" if hit else "-0：") + rule)
    # 30 学术诚信声明页文字被改动（-3）
    decl_ok = (model["decl_title"] is not None and model["decl_sign"] is not None
               and model["decl_date"] is not None)
    if decl_ok:
        cursor = model["decl_title"]
        for head in DECL_BODY_HEADS:
            found = None
            for i in range(cursor + 1, model["decl_sign"] + 1):
                if _norm(head) in ntexts[i]:
                    found = i
                    break
            if found is None:
                decl_ok = False
                break
            cursor = found
        if decl_ok and not (model["decl_title"] < model["decl_sign"] < model["decl_date"]):
            decl_ok = False
    hit = not decl_ok
    rule = ("“学术诚信声明及作品使用授权书”页的标题（黑体，16磅，加粗，居中）与其下四段声明正文、"
            "“毕业调研报告作者签名：”行、“2026年  4月  20日”行的文字被改写、删减或调序")
    results.append(("-3：" if hit else "-0：") + rule)

    # 31 正文 13 处参考文献标注上标/位置被改动（-3）
    # 与办公软件对齐：上标由 run 级 w:vertAlign="superscript" 承载；
    # 位置要求标注落在句号之前，故标注前一字符不应为“。”。
    sup_marks = 0
    mark_pos_bad = False
    for idx in model["body_idx"]:
        para = paras[idx]
        for run_el, text in _para_runs(para):
            for _m in re.finditer(r"\[\d+\]", text):
                if _run_superscript(run_el):
                    sup_marks += 1
        for match in re.finditer(r"\[\d+\]", texts[idx]):
            start = match.start()
            if start > 0 and texts[idx][start - 1] in "。．.":
                mark_pos_bad = True
    hit = sup_marks != len(REF_AUTHORS) or mark_pos_bad
    rule = ("一至七章正文中13处参考文献标注（如“……可持续性[1]”）不再为上标形式，"
            "或标注位置由句号之前被移到句号之后")
    results.append(("-3：" if hit else "-0：") + rule)

    # 32 摘要/Abstract/参考文献/致谢正文文字被改动（-3）
    body_text_ok = True
    cn_text = _norm("".join(texts[i] for i in model["abs_cn_body_idx"]))
    en_text = "".join(texts[i] for i in model["abs_en_body_idx"])
    for head in ABSTRACT_CN_HEADS:
        if _norm(head) not in cn_text:
            body_text_ok = False
    for head in ABSTRACT_EN_HEADS:
        if head not in en_text:
            body_text_ok = False
    if len(model["ref_idx"]) != len(REF_AUTHORS):
        body_text_ok = False
    else:
        for author, idx in zip(REF_AUTHORS, model["ref_idx"]):
            if author not in texts[idx]:
                body_text_ok = False
    if len(model["thanks_body_idx"]) < len(THANKS_HEADS):
        body_text_ok = False
    else:
        for head, idx in zip(THANKS_HEADS, model["thanks_body_idx"]):
            if _norm(head) not in ntexts[idx]:
                body_text_ok = False
    hit = not body_text_ok
    rule = "摘要正文、Abstract正文、13条参考文献著录、致谢4段的文字内容被改写、删减、增补或调换顺序"
    results.append(("-3：" if hit else "-0：") + rule)
    # 33 声明页标题缩进与四段声明正文排版被改动（-1）
    # 指令只涉及目录与论文正文格式，声明页排版应原样保留；标题缩进与四段排版
    # 同属该页版式位移，命中任一处即扣该条分值，不重复累扣。
    # 首行缩进按原稿的设置方式判定：原稿这四段是按固定值设置的（560 twips），改成
    # 按字符折算出的等值缩进不视为原样保留——单位一换，缩进量就随字号浮动了。
    decl_layout_ok = True
    idx = model["decl_title"]
    if idx is None:
        decl_layout_ok = False
    else:
        ind = _para_indent(ctx, paras[idx], 32)
        if not (_near(ind["first_line"], 0, 40) and _near(ind["left"], 0, 40)):
            decl_layout_ok = False
    if model["decl_title"] is None or model["decl_sign"] is None:
        decl_layout_ok = False
    else:
        for i in range(model["decl_title"] + 1, model["decl_sign"]):
            if not ntexts[i]:
                continue
            para = paras[i]
            ind = _para_indent(ctx, para, 24)
            first = ind["first_line"]
            # 首段以两个全角空格代替缩进属原有写法，此时首行缩进为 0 亦达标。
            lead_full_width = _para_text(para).startswith("　　")
            if _para_jc(ctx, para) != "both":
                decl_layout_ok = False
            elif lead_full_width and first < 60:
                continue
            elif not (_near(first, 560, 40) and ind["first_line_chars"] is None):
                decl_layout_ok = False
    hit = not decl_layout_ok
    rule = ("“学术诚信声明及作品使用授权书”页的版式被改动——居中标题被添加首行缩进"
            "（首行缩进与左缩进应均为0）、四段声明正文的两端对齐与首行缩进"
            "（须沿用原稿按固定值设置的缩进量，改按字符折算出等值缩进视同改动；"
            "首段以两个全角空格代替缩进者视同达标）被改动")
    results.append(("-1：" if hit else "-0：") + rule)

    # 34 封面大标题与默认样式的中文字体设定被清除（-1）
    # 与办公软件对齐：中文字体经 w:eastAsia → eastAsiaTheme → theme1.xml 解析，
    # 主题中 script="Hans" 缺失时视为中文字体未定义，即已丢失宋体设定；
    # 清字体与删主题 Hans 同属一次破坏，合为一条，不重复累扣。
    cover_font_ok = False
    idx = model["cover_title"]
    if idx is not None:
        para = paras[idx]
        cover_font_ok = (_runs_ok(ctx, para, size=44, bold=True, ea_pred=_is_song)
                         and _para_jc(ctx, para) == "center")
    default_ea = _style_eastasia(ctx, ctx.default_para_style)
    theme_hans = ctx.theme_font("minorEastAsia")
    hit = not (cover_font_ok and (_is_song(default_ea) or _is_song(theme_hans)))
    rule = ("封面首行大标题“高等教育自学考试毕业调研报告封面”不满足字体格式为宋体 二号"
            "（22磅）加粗 居中，或文档默认段落样式（正文样式）的中文字体被置空"
            "或改为非宋体族，致使封面文字与未显式设定字体的空段落"
            "改由主题默认字体渲染")
    results.append(("-1：" if hit else "-0：") + rule)

    # 35 电子签名图片丢失或被替换（-1）
    # 与办公软件对齐：图片实体存于 media 部件，显示尺寸由 wp:extent 的 cx/cy 决定。
    sign_blob_ok = False
    for part in doc.part.package.iter_parts():
        content_type = getattr(part, "content_type", "") or ""
        if not content_type.startswith("image/"):
            continue
        try:
            blob = part.blob
        except Exception:
            continue
        if content_type.endswith("png") and len(blob) == SIGN_IMAGE_LEN:
            sign_blob_ok = True
    extent_ok = False
    for extent in body.iter("{%s}extent" % _WP_NS):
        if (_int_attr_plain(extent, "cx") == SIGN_IMAGE_CX
                and _int_attr_plain(extent, "cy") == SIGN_IMAGE_CY):
            extent_ok = True
    hit = not (sign_blob_ok and extent_ok)
    rule = ("学术诚信声明及作品使用授权书页签名处嵌入的电子签名图片（PNG，嵌入型环绕）"
            "丢失、被裁剪或被替换")
    results.append(("-1：" if hit else "-0：") + rule)

    # 36 页眉页脚部件被扩增且首页/偶数页部件为空（-1）
    # 与办公软件对齐：type="first" 需 w:titlePg 生效、type="even" 需
    # w:evenAndOddHeaders 生效；引用了却内容为空即渲染出空白页眉页脚。
    empty_extra = False
    for i, section in enumerate(sections):
        for tag, attr in (("w:headerReference", "first"), ("w:footerReference", "first"),
                          ("w:headerReference", "even"), ("w:footerReference", "even")):
            for ref in section._sectPr.findall(qn(tag)):
                if _attr(ref, "w:type") != attr:
                    continue
                rid = ref.get(qn("r:id"))
                part = doc.part.related_parts.get(rid) if rid else None
                if not _part_has_text(part):
                    empty_extra = True
    hit = empty_extra
    rule = ("页眉页脚部件被扩增为首页、奇数页、偶数页各自独立，且其中偶数页与首页部件为空白，"
            "一旦启用“奇偶页不同”即出现页眉与页码丢失")
    results.append(("-1：" if hit else "-0：") + rule)
    # 37 章节序号写法被改动或标题未单独成行（-1）
    order_ok = (l1_all and l2_all
                and model["l1_idx"] == sorted(model["l1_idx"])
                and model["l2_idx"] == sorted(model["l2_idx"]))
    standalone_ok = True
    for title, idx in model["title_index"].items():
        if ntexts[idx] != _norm(title):
            standalone_ok = False
    hit = not (order_ok and standalone_ok)
    rule = ("章节序号写法被改为“（一）、”“1、”等形式，或任一标题未单独成行"
            "（与正文文字挤在同一段）")
    results.append(("-1：" if hit else "-0：") + rule)

    # 38 关键词行标点被改动（-1）
    # 与办公软件对齐：仅核验“关键词”一行的标点，须为2个全角分号且行尾不带标点。
    hit = kw_idx is None or _norm(texts[kw_idx]) != _norm(KEYWORDS_LINE)
    rule = ("“关键词：雨水花园；社区参与；日常管护”末尾被添加句号、顿号等标点，"
            "或分号被改为逗号、顿号")
    results.append(("-1：" if hit else "-0：") + rule)

    return results, model, ctx

def _iter_docx(dir_path):
    """列出 dir_path 直属的 .docx，不进入任何子目录。

    只认当前目录下的文件：子目录中的同名输出文件、输入素材目录等一概不纳入候选，
    避免把参照物或其他批次的输出当作被评对象。
    """
    return sorted(
        os.path.join(dir_path, f) for f in os.listdir(dir_path)
        if f.lower().endswith(".docx") and not f.startswith("~")
        and SPEC_FILE_KEYWORD not in f
    )


def _script_dir():
    """脚本自身所在目录。无 __file__ 的场景（交互式执行）退回当前目录。"""
    try:
        return os.path.dirname(os.path.abspath(__file__))
    except NameError:  # pragma: no cover - 仅交互式粘贴执行时触发
        return os.getcwd()


def _default_target():
    """不传参时的默认评测目标：脚本所在目录本身，不再向上或向下查找。

    使脚本无论从哪个工作目录启动，都只评测与自己放在同一目录的输出文件。
    """
    return _script_dir()


def _locate_error(result, message):
    """未能定位到被评测文档（路径不存在 / 目录及子目录中没有 .docx）。

    维度一根本没有进入判定，故 dim1_pass 置为空字符串表示“未判定”、dim1_reason 留空，
    失败摘要只写在 error；维度二一项未评，max_score 同时归 0。
    """
    result["status"] = "error"
    result["error"] = message
    result["dim1_pass"] = ""
    result["dim1_reason"] = ""
    result["dim2_items"] = []
    result["total_score"] = 0
    result["max_score"] = 0
    return result


def evaluate(dir_path=None):
    """评测目录下的 .docx 输出文件，返回评分字典。

    dir_path 可为输出文件所在目录，或直接指向 .docx 文件；
    省略或传 None 时自动取脚本所在目录（见 _default_target）。
    传入目录时只识别该目录直属的 .docx，不进入子目录；
    存在多个候选时按路径排序取第一个，file_name 记录实际评测的文件。

    错误返回约定：
      ① 定位不到被评测文档（路径不存在 / 无 .docx）→ status="error"，error 写失败摘要，
         dim1_pass=""（未判定）、dim1_reason=""、dim2_items=[]、total_score=0、max_score=0；
      ② 文档存在但打不开等导致维度一不通过 → status="error"，dim1_pass=False，
         此时 dim1_reason 须写出错误摘要。
    """
    if dir_path is None:
        dir_path = _default_target()
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
        if not dir_path or not os.path.exists(dir_path):
            return _locate_error(result, "路径不存在：%s" % dir_path)
        if os.path.isfile(dir_path):
            candidates = [dir_path]
        else:
            candidates = _iter_docx(dir_path)
        if not candidates:
            return _locate_error(result, "目录中未找到 .docx 文件：%s" % dir_path)
        file_path = candidates[0]
        result["file_name"] = os.path.basename(file_path)
        results, _model, _ctx = _evaluate_impl(file_path, result)
    except Exception as exc:  # noqa: BLE001 - 任何异常都记为不可用
        # 文件损坏、无法打开等：维度一不通过，完成度记 0；
        # 错误摘要同时写入 error 与 dim1_reason，
        # 若异常前 _check_dim1 已给出不通过理由，则在其后追加摘要。
        summary = "%s: %s" % (type(exc).__name__, exc)
        result["status"] = "error"
        result["error"] = summary
        result["dim1_pass"] = False
        result["dim1_reason"] = ("%s；%s" % (result["dim1_reason"], summary)
                                 if result["dim1_reason"] else summary)
        result["dim2_items"] = []
        result["total_score"] = 0
        return result

    assert len(results) == len(_MAX_DELTAS_ORDER)

    # 反向解析每条结论：句首“+N：/-N：”即该条实际增减分，
    # 与 _MAX_DELTAS_ORDER 对位即可还原满分与命中状态。
    items = [None] * len(results)
    total = 0
    for pos, (line, max_delta) in enumerate(zip(results, _MAX_DELTAS_ORDER)):
        head, rule = line.split("：", 1)
        delta = int(head.replace("+", ""))
        if not result["dim1_pass"]:
            delta = 0
        hit = delta != 0
        total += delta
        # 按 _EMIT_TO_RUBRIC 归位，使输出顺序与细则“先加分项、后减分项”一致。
        items[_EMIT_TO_RUBRIC[pos]] = {
            "rule": rule,
            "max_delta": max_delta,
            "delta": delta,
            "hit": hit,
        }
    result["dim2_items"] = items
    result["total_score"] = max(0, total)
    return result


if __name__ == "__main__":
    # 本地调试入口：python officeval_012_verifier.py [目录或 .docx 路径]
    # 不传参时自动取脚本所在目录（而非当前工作目录），只识别该目录直属的输出文件；
    # 找到多个候选时逐个评测，输出 JSON 数组，便于同时比对多份输出文件。
    _target = sys.argv[1] if len(sys.argv) > 1 else _default_target()
    try:
        if os.path.isfile(_target):
            _paths = [_target]
        else:
            _paths = _iter_docx(_target)
        if len(_paths) > 1:
            _out = [evaluate(p) for p in _paths]
        else:
            _out = evaluate(_target)
    except Exception:  # pragma: no cover - 调试期直接暴露栈信息
        traceback.print_exc()
        raise
    print(json.dumps(_out, ensure_ascii=False, indent=2))

