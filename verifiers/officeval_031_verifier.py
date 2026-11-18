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
# 主结果统一由 evaluate(dir_path) -> dict 返回；入参为“脚本所在目录的路径”，
# 由脚本自己在该目录内定位并打开被评估的 .docx。本文件仅在 __main__ 下作为
# 本地调试入口，才把 dict 序列化为 JSON 打印。
# 返回字段：id / file_name / status / error / dim1_pass / dim1_reason /
#           dim2_items（每项只含 rule / max_delta / delta / hit）/
#           total_score / max_score
# status='error' 表示这份提交没被真正评过（目录不存在、目录内无 .docx、脚本
# 自身异常），此时 dim1_pass 为空串（维度一未判定）、total_score 与 max_score
# 一律为 0；dim1_pass 为 False 时 dim1_reason 必须给出错误摘要，通过时为空串。

SCRIPT_ID = "031"
# 维度二所有加分项 max_delta 之和；用于维度一未通过 / 脚本异常时兜底
MAX_SCORE_TOTAL = 81
# 维度二所有减分项 max_delta 之和（负值）；仅作自检用，不进入返回结果
MIN_SCORE_TOTAL = -43

# 维度二所有评分项按“结果追加顺序”的 max_delta 列表；末端把 results 里的
# “+N：...”/“-N：...”字符串反解为结构化 dim2_items 时用。顺序须与
# _evaluate_impl 中 _emit(...) 的顺序严格一致；共 48 项（27 加分 + 21 减分）。
_MAX_DELTAS_ORDER = [
    5,   # 01 九个大题齐全、题号一至九连续
    1,   # 02 卷头主标题文字
    1,   # 03 卷头拼音行文字
    1,   # 04 卷头副标题文字
    3,   # 05 卷头右栏四行填写栏
    3,   # 06 九个题块 1行1列 / 18.40cm / 细框线无网格
    3,   # 07 九个题块底纹三色分组
    3,   # 08 九个大题标题 字体字号加粗颜色
    3,   # 09 九个大题标题上方拼音行格式
    3,   # 10 题块内题干正文 字体字号
    3,   # 11 题块内嵌套表格 细线 D4E7F8/CAE5B7
    5,   # 12 一题 8 个拼音齐全且配书写区
    3,   # 13 一题书写区为田字格图片/方格表格
    3,   # 14 二题 6 组加点字及两个待选读音
    5,   # 15 二题六个受考字带着重点
    3,   # 16 三题六个待补词语及两句选词填空
    3,   # 17 四题笔画题 / 音序题 / 关联词填空
    3,   # 18 五题两关内容齐全
    3,   # 19 六题例句与比喻句
    3,   # 20 七题两关内容齐全
    5,   # 21 八题短文 4 个自然段齐全
    1,   # 22 八题短文标题单独成段居中加粗
    3,   # 23 八题 4 小题及分值齐全且未拆分版块
    3,   # 24 九题写作要求两项
    5,   # 25 九题作文格 20列×6行 格宽行高
    3,   # 26 九题作文格边框内边距对齐方式
    1,   # 27 试题内容完整且未附参考答案
    -3,  # 28 页眉文字被改写或删除
    -1,  # 29 页眉字体字号颜色对齐
    -1,  # 30 页脚页码格式
    -3,  # 31 提示块文字被改写
    -1,  # 32 “闯关提示：”格式
    -1,  # 33 提示块底纹或被加边框
    -3,  # 34 卷头左右单元格底纹
    -1,  # 35 卷头单元格垂直居中
    -3,  # 36 卷头拼音行格式
    -3,  # 37 卷头主标题格式
    -1,  # 38 卷头副标题格式
    -3,  # 39 卷头右栏填写行格式
    -1,  # 40 卷头信息块行距
    -1,  # 41 卷头信息块水平居中与列宽
    -3,  # 42 表格被加粗框线或网格线
    -3,  # 43 插入硬分页符
    -3,  # 44 额外添加“题号／得分”评分表
    -3,  # 45 额外臆造大题标题
    -3,  # 46 纸张尺寸与页边距
    -1,  # 47 页面边框
    -1,  # 48 正文默认字体
]
assert sum(d for d in _MAX_DELTAS_ORDER if d > 0) == MAX_SCORE_TOTAL
assert sum(d for d in _MAX_DELTAS_ORDER if d < 0) == MIN_SCORE_TOTAL

# ---------- 模板基线常量：细则里写死的目标文字与目标格式值 ----------
SANS = "Noto Sans CJK SC"
SERIF = "Noto Serif CJK SC"
TNR = "Times New Roman"

NINE_TITLES = (
    "一、看拼音，写词语。（12分）",
    "二、给加点字选择正确读音，画“√”。（6分）",
    "三、把词语补充完整，并选词填空。（10分）",
    "四、按要求填一填。（10分）",
    "五、词语闯关。（9分）",
    "六、按要求写句子。（8分）",
    "七、积累与运用。（10分）",
    "八、阅读直通车。（20分）",
    "九、习作。",
)
MAIN_TITLE = "三年级下册语文达标测试卷"
COVER_PINYIN = "sān nián jí xià cè yǔ wén dá biāo cè shì juàn"
SUB_TITLE_PARTS = ("第四单元测试卷", "时间：60分钟　满分：100分")
FILL_LABELS = ("学校", "班级", "姓名", "得分")

BLOCK_SHADES = ("F2F8FF", "F2F8FF", "F2F8FF", "F8FFF2", "F8FFF2", "F8FFF2",
                "FFF7F0", "FFF7F0", "FFF7F0")
TITLE_COLORS = ("2F75B5", "2F75B5", "2F75B5", "548235", "548235", "548235",
                "C55A11", "C55A11", "C55A11")
TITLE_PY_COLOR = "646464"
# 版式基线中允许出现的浅色细格线配色：卷头 9FC5E8、提示块 E2EAF2、题块 B7D4EF、
# 题块内蓝色系嵌套表 D4E7F8、绿色系嵌套表 CAE5B7；其余颜色即视为另起一套框线。
NESTED_BORDER_COLORS = ("D4E7F8", "CAE5B7", "9FC5E8")
BASE_BORDER_COLORS = ("9FC5E8", "E2EAF2", "B7D4EF", "D4E7F8", "CAE5B7")

HEADER_TEXT = ("pīn yīn xiǎo liàn xí · měi tiān jìn bù yì diǎn diǎn ｜ "
               "拼音小练习 · 每天进步一点点")
HEADER_COLOR = "767676"
FOOTER_COLOR = "787878"
TIP_TEXT = ("闯关提示：请认真拼读每一道题，书写要端正，"
            "遇到轻声和整体认读音节要多读几遍。")
TIP_PREFIX = "闯关提示："
TIP_PREFIX_COLOR = "C07000"
TIP_SHADE = "FFFDEB"
COVER_SHADES = ("EAF5FF", "FDF2E9")
COVER_PY_COLOR = "56739A"
COVER_TITLE_COLOR = "1F4E79"
COVER_SUB_COLOR = "5B5B5B"

Q1_PINYIN = ("wǔ dǎo", "zhì huì", "xīn shǎng", "bǎo cún",
             "nuò ruò", "jīng yàn", "mì fēng", "xiū jiàn")
Q2_WORDS = ("挣脱", "切开", "模样", "掠过", "闲散", "纤细")
Q2_CHARS = ("挣", "切", "模", "掠", "散", "纤")
Q3_WORDS = ("邯郸", "杯弓", "争奇", "翩翩", "学富", "守株")
Q5_PARTS = ("第一关", "第二关", "安安静静", "又细又长")
Q5_QUANTIFIERS = ("阵", "座", "本", "片", "朵", "只")
Q6_PARTS = ("花开了", "小燕子飞过田野", "比喻句")
Q7_POEMS = ("迟日江山丽", "独在异乡为异客", "观察事物要")
Q7_OPTIONS = ("认真观察", "随便猜测", "连续记录", "只看一次")
Q8_TITLE = "会写日记的银杏树"
Q9_PARTS = ("观察一种植物或一种小动物", "不少于150字")

# 页面基线：A4 11906×16838 twips；页边距 上567/下454/左737/右737 twips
PG_W, PG_H = 11906, 16838
PG_MAR = {"top": 567, "bottom": 454, "left": 737, "right": 737}
PG_BORDER = ("single", 16, "DDEAF7", 18)
# 题块网格总宽 18.40cm = 10432 twips；作文格每格 0.85cm = 482 twips
BLOCK_WIDTH = 10432
CELL_WIDTH = 482
CELL_HEIGHT = 470          # 0.83cm 行高下限
GRID_BORDER_COLOR = "737373"

# ---------- 文本与元素基础工具 ----------
_MC_FALLBACK = "{http://schemas.openxmlformats.org/markup-compatibility/2006}Fallback"
_WS_RE = re.compile(r"[\s 　]+")
_QUOTE_RE = re.compile(r"[“”\"'‘’「」『』]")
_CJK_RE = re.compile(r"[一-鿿]")
_LATIN_RE = re.compile(r"[A-Za-z]")
_PINYIN_ONLY_RE = re.compile(
    r"^[A-Za-zāáǎàēéěèīíǐìōóǒòūúǔùǖǘǚǜüńňǹ\s,，.。、·:：;；!！?？'’\-—－]+$")
_UNDERLINE_RE = "[_＿﹍﹏—─－]"


def _in_fallback(el):
    """判断元素是否落在 mc:Fallback 分支内——该分支只是旧版本兼容副本，
    与办公软件实际渲染的 mc:Choice 内容重复，统计时必须排除。"""
    node = el.getparent()
    while node is not None:
        if node.tag == _MC_FALLBACK:
            return True
        node = node.getparent()
    return False


def _t(el):
    """元素内可见文本（含文本框 w:txbxContent 内的文字，排除 mc:Fallback 副本）。"""
    if el is None:
        return ""
    return "".join(n.text or "" for n in el.iter(qn("w:t")) if not _in_fallback(n))


def _norm(s):
    """去掉所有空白（含半角空格、不换行空格、全角空格），便于跨排版比对文字。"""
    return _WS_RE.sub("", s or "")


def _nq(s):
    """在 _norm 基础上再去掉各种引号，规避直/弯引号差异导致的假失配。"""
    return _QUOTE_RE.sub("", _norm(s))


def _has_cjk(s):
    return bool(_CJK_RE.search(s or ""))


def _has_latin(s):
    return bool(_LATIN_RE.search(s or ""))


def _paras(el):
    """元素下所有段落，按文档顺序（排除 mc:Fallback 副本）。"""
    if el is None:
        return []
    return [p for p in el.iter(qn("w:p")) if not _in_fallback(p)]


def _runs(p):
    """段落自身的 run（排除嵌在段内文本框里、属于另一段的 run）。"""
    out = []
    for r in p.iter(qn("w:r")):
        if _in_fallback(r):
            continue
        anc = r.getparent()
        while anc is not None and anc.tag != qn("w:p"):
            anc = anc.getparent()
        if anc is p:
            out.append(r)
    return out


def _rt(r):
    """run 自身的文字：只取归属该 run 的 w:t，不含它内部图形/文本框里另一段落的文字。"""
    out = []
    for n in r.iter(qn("w:t")):
        if _in_fallback(n):
            continue
        anc = n.getparent()
        while anc is not None and anc.tag != qn("w:r"):
            anc = anc.getparent()
        if anc is r:
            out.append(n.text or "")
    return "".join(out)


def _jc(p):
    """段落对齐方式（w:jc/@w:val）；未设置返回 None。"""
    pPr = p.find(qn("w:pPr"))
    if pPr is None:
        return None
    el = pPr.find(qn("w:jc"))
    return el.get(qn("w:val")) if el is not None else None


def _line_spacing(p):
    """段落行距 (line, lineRule)；未设置返回 (None, None)。"""
    pPr = p.find(qn("w:pPr"))
    if pPr is None:
        return (None, None)
    el = pPr.find(qn("w:spacing"))
    if el is None:
        return (None, None)
    v = el.get(qn("w:line"))
    return (int(v) if v and v.isdigit() else None, el.get(qn("w:lineRule")))


# ---------- 字体继承链：解析办公软件里实际生效的字体/字号/颜色/加粗 ----------
def _style_by_id(styles_el, sid):
    for st in styles_el.findall(qn("w:style")):
        if st.get(qn("w:styleId")) == sid:
            return st
    return None


def _rpr_chain(doc, p, r):
    """按 Word 的继承链返回待查 rPr 列表，靠前者优先生效：
       run 级 → 段落标记 rPr → 段落样式（含 basedOn 上溯）→ 默认段落样式 → docDefaults。
       docDefaults 优先级最低，默认段落样式（Normal）可覆盖它，与办公软件解析顺序一致。"""
    out = []
    if r is not None:
        rPr = r.find(qn("w:rPr"))
        if rPr is not None:
            out.append(rPr)
    if p is not None:
        pPr = p.find(qn("w:pPr"))
        if pPr is not None:
            rPr = pPr.find(qn("w:rPr"))
            if rPr is not None:
                out.append(rPr)
    styles_el = doc.styles.element
    sid = None
    if p is not None:
        pPr = p.find(qn("w:pPr"))
        if pPr is not None:
            ps = pPr.find(qn("w:pStyle"))
            sid = ps.get(qn("w:val")) if ps is not None else None
    seen = set()
    while sid and sid not in seen:
        seen.add(sid)
        st = _style_by_id(styles_el, sid)
        if st is None:
            break
        rpr = st.find(qn("w:rPr"))
        if rpr is not None:
            out.append(rpr)
        base = st.find(qn("w:basedOn"))
        sid = base.get(qn("w:val")) if base is not None else None
    for st in styles_el.findall(qn("w:style")):
        if st.get(qn("w:type")) == "paragraph" and st.get(qn("w:default")) == "1":
            rpr = st.find(qn("w:rPr"))
            if rpr is not None:
                out.append(rpr)
            break
    dd = styles_el.find(qn("w:docDefaults"))
    if dd is not None:
        rpd = dd.find(qn("w:rPrDefault"))
        if rpd is not None:
            rpr = rpd.find(qn("w:rPr"))
            if rpr is not None:
                out.append(rpr)
    return out


def _eff_font(doc, p, r, kind):
    """kind='eastAsia' 取中文实际生效字体；kind='ascii' 取西文/数字字体（回落 hAnsi）。"""
    for rpr in _rpr_chain(doc, p, r):
        rf = rpr.find(qn("w:rFonts"))
        if rf is None:
            continue
        if kind == "eastAsia":
            v = rf.get(qn("w:eastAsia"))
        else:
            v = rf.get(qn("w:ascii")) or rf.get(qn("w:hAnsi"))
        if v:
            return v
    return None


def _eff_size(doc, p, r):
    """实际生效字号（磅）；w:sz 为半磅值。"""
    for rpr in _rpr_chain(doc, p, r):
        el = rpr.find(qn("w:sz"))
        if el is not None:
            v = el.get(qn("w:val"))
            if v and v.isdigit():
                return int(v) / 2.0
    return None


def _eff_color(doc, p, r):
    """实际生效字体颜色（大写十六进制）；auto/未设置返回 None。"""
    for rpr in _rpr_chain(doc, p, r):
        el = rpr.find(qn("w:color"))
        if el is not None:
            v = el.get(qn("w:val"))
            if v and v.lower() != "auto":
                return v.upper()
    return None


def _eff_flag(doc, p, r, tag):
    """加粗 tag='b' / 斜体 tag='i'：无 val 视为 True，val 为 0/false/off 视为 False。"""
    for rpr in _rpr_chain(doc, p, r):
        el = rpr.find(qn("w:" + tag))
        if el is not None:
            return el.get(qn("w:val")) not in ("0", "false", "off")
    return False


def _has_em(r):
    """run 是否带着重号（w:em，val 非 none）。"""
    rpr = r.find(qn("w:rPr"))
    if rpr is None:
        return False
    el = rpr.find(qn("w:em"))
    if el is None:
        return False
    return (el.get(qn("w:val")) or "dot") != "none"


# ---------- 表格工具 ----------
def _tbls(el):
    """元素下所有表格（含嵌套），按文档顺序。"""
    if el is None:
        return []
    return [t for t in el.iter(qn("w:tbl")) if not _in_fallback(t)]


def _rows(tbl):
    return tbl.findall(qn("w:tr")) if tbl is not None else []


def _cells(tr):
    return tr.findall(qn("w:tc"))


def _grid(tbl):
    """表格列宽序列（twips）。"""
    if tbl is None:
        return []
    g = tbl.find(qn("w:tblGrid"))
    if g is None:
        return []
    out = []
    for c in g.findall(qn("w:gridCol")):
        v = c.get(qn("w:w"))
        out.append(int(v) if v and v.isdigit() else 0)
    return out


def _tblPr(tbl):
    return tbl.find(qn("w:tblPr")) if tbl is not None else None


def _tcPr(tc):
    return tc.find(qn("w:tcPr")) if tc is not None else None


def _shd_fill(el):
    if el is None:
        return None
    shd = el.find(qn("w:shd"))
    if shd is None:
        return None
    v = shd.get(qn("w:fill"))
    return v.upper() if v and v.lower() != "auto" else None


def _tc_shade(tc):
    return _shd_fill(_tcPr(tc))


def _tc_valign(tc):
    pr = _tcPr(tc)
    if pr is None:
        return None
    el = pr.find(qn("w:vAlign"))
    return el.get(qn("w:val")) if el is not None else None


def _tbl_jc(tbl):
    pr = _tblPr(tbl)
    if pr is None:
        return None
    el = pr.find(qn("w:jc"))
    return el.get(qn("w:val")) if el is not None else None


_NO_LINE = (None, "nil", "none")


def _border(el, name):
    """返回 (val, sz, color)；el 为 tblBorders / tcBorders。"""
    if el is None:
        return (None, None, None)
    sub = el.find(qn("w:" + name))
    if sub is None:
        return (None, None, None)
    val = sub.get(qn("w:val"))
    szs = sub.get(qn("w:sz"))
    col = sub.get(qn("w:color"))
    return (val,
            int(szs) if szs and szs.isdigit() else None,
            col.upper() if col else None)


def _has_line(val):
    return val is not None and val not in _NO_LINE


def _merge_border(*srcs):
    """按 Word 的边框覆盖规则取生效边框：靠前的来源只要“显式声明过”就胜出，
       显式声明 nil/none 即表示该处不画线（单元格可覆盖表级 insideH/insideV）；
       全部未声明则返回未声明。"""
    for item in srcs:
        if item[0] is not None:
            return item
    return (None, None, None)


def _tbl_borders(tbl):
    pr = _tblPr(tbl)
    return pr.find(qn("w:tblBorders")) if pr is not None else None


def _tc_borders(tc):
    pr = _tcPr(tc)
    return pr.find(qn("w:tcBorders")) if pr is not None else None


def _rendered_borders(tbl):
    """按 Word/WPS 实际渲染的每条格线返回 [(kind, val, sz, color)]：
       外框 = tc 侧边框 || tblBorders 对应边；
       内部横线 = 上下相邻 tc 的 bottom/top || tblBorders.insideH（单行表不产生）；
       内部竖线 = 左右相邻 tc 的 right/left || tblBorders.insideV（单列表不产生）。
       kind ∈ {'top','bottom','left','right','insideH','insideV'}。"""
    out = []
    rows = _rows(tbl)
    if not rows:
        return out
    tb = _tbl_borders(tbl)
    b_top = _border(tb, "top")
    b_bottom = _border(tb, "bottom")
    b_left = _border(tb, "left")
    b_right = _border(tb, "right")
    b_ih = _border(tb, "insideH")
    b_iv = _border(tb, "insideV")
    grids = [_cells(tr) for tr in rows]
    for ri, cells in enumerate(grids):
        for ci, tc in enumerate(cells):
            tcb = _tc_borders(tc)
            if ri == 0:
                out.append(("top",) + _merge_border(_border(tcb, "top"), b_top))
            else:
                prev = grids[ri - 1]
                pb = _border(_tc_borders(prev[ci]), "bottom") if ci < len(prev) else (None, None, None)
                out.append(("insideH",) + _merge_border(_border(tcb, "top"), pb, b_ih))
            if ri == len(grids) - 1:
                out.append(("bottom",) + _merge_border(_border(tcb, "bottom"), b_bottom))
            if ci == 0:
                out.append(("left",) + _merge_border(_border(tcb, "left"), b_left))
            else:
                pr_ = _border(_tc_borders(cells[ci - 1]), "right")
                out.append(("insideV",) + _merge_border(_border(tcb, "left"), pr_, b_iv))
            if ci == len(cells) - 1:
                out.append(("right",) + _merge_border(_border(tcb, "right"), b_right))
    return out


def _cell_margins(tbl, tc):
    """单元格生效内边距 (top,left,bottom,right)，dxa；tcMar 优先，其次 tblCellMar，
       两者都没有则返回 None（表示未显式设定）。"""
    out = []
    pr = _tcPr(tc)
    tcmar = pr.find(qn("w:tcMar")) if pr is not None else None
    tpr = _tblPr(tbl)
    tblmar = tpr.find(qn("w:tblCellMar")) if tpr is not None else None
    for name in ("top", "left", "bottom", "right"):
        val = None
        for src in (tcmar, tblmar):
            if src is None:
                continue
            sub = src.find(qn("w:" + name))
            if sub is not None:
                v = sub.get(qn("w:w"))
                if v is not None and v.lstrip("-").isdigit():
                    val = int(v)
                    break
        out.append(val)
    return tuple(out)


def _row_height(tr):
    """行高 (val, hRule)；未设置返回 (None, None)。"""
    pr = tr.find(qn("w:trPr"))
    if pr is None:
        return (None, None)
    el = pr.find(qn("w:trHeight"))
    if el is None:
        return (None, None)
    v = el.get(qn("w:val"))
    return (int(v) if v and v.isdigit() else None, el.get(qn("w:hRule")))


def _drawing_count(el):
    """元素内图形对象数量（w:drawing / w:pict，排除 mc:Fallback 副本）。"""
    n = 0
    for tag in ("w:drawing", "w:pict"):
        n += len([x for x in el.iter(qn(tag)) if not _in_fallback(x)])
    return n


# 着重号小圆点的最大外形尺寸（EMU）：宽 ≤0.85cm、高 ≤0.4cm，超出即视为普通插图。
_DOT_MAX_CX, _DOT_MAX_CY = 306000, 144000


def _dot_mark_count(el):
    """元素内“着重点”标记数：受考字下方的小圆点图形（按外形尺寸甄别）。"""
    if el is None:
        return 0
    n = 0
    for d in el.iter(qn("w:drawing")):
        if _in_fallback(d):
            continue
        for ext in d.iter():
            if not str(ext.tag).endswith("}extent"):
                continue
            cx = ext.get("cx")
            cy = ext.get("cy")
            if (cx and cy and cx.isdigit() and cy.isdigit()
                    and int(cx) <= _DOT_MAX_CX and int(cy) <= _DOT_MAX_CY):
                n += 1
            break
    return n


# ---------- 页面设置与页眉页脚 ----------
def _sect_prs(body):
    return [s for s in body.iter(qn("w:sectPr"))]


def _pg_size(sect):
    el = sect.find(qn("w:pgSz"))
    if el is None:
        return (None, None)
    w = el.get(qn("w:w"))
    h = el.get(qn("w:h"))
    return (int(w) if w and w.isdigit() else None,
            int(h) if h and h.isdigit() else None)


def _pg_margins(sect):
    el = sect.find(qn("w:pgMar"))
    out = {}
    for name in ("top", "bottom", "left", "right"):
        v = el.get(qn("w:" + name)) if el is not None else None
        out[name] = int(v) if v and v.lstrip("-").isdigit() else None
    return out


def _pg_borders(sect):
    """页面边框四边 {(side): (val, sz, color, space)}；未设置返回空 dict。"""
    el = sect.find(qn("w:pgBorders"))
    if el is None:
        return {}
    out = {}
    for name in ("top", "left", "bottom", "right"):
        sub = el.find(qn("w:" + name))
        if sub is None:
            continue
        szs = sub.get(qn("w:sz"))
        sps = sub.get(qn("w:space"))
        col = sub.get(qn("w:color"))
        out[name] = (sub.get(qn("w:val")),
                     int(szs) if szs and szs.isdigit() else None,
                     col.upper() if col else None,
                     int(sps) if sps and sps.isdigit() else None)
    return out


def _hdrftr_elements(doc, which):
    """首节实际生效的页眉/页脚根元素列表（默认 → 首页 → 偶数页，只取有内容的）。
       首节即覆盖正文各页的节；未定义的页眉页脚（继承自前一节）跳过。"""
    out = []
    try:
        sect = doc.sections[0]
    except Exception:
        return out
    if which == "header":
        cands = (sect.header, sect.first_page_header, sect.even_page_header)
    else:
        cands = (sect.footer, sect.first_page_footer, sect.even_page_footer)
    for c in cands:
        try:
            if c.is_linked_to_previous:
                continue
            el = c._element
        except Exception:
            continue
        if el is None:
            continue
        has_field = len([x for x in el.iter(qn("w:instrText"))]) > 0
        if _norm(_t(el)) or has_field:
            out.append(el)
    return out


# ---------- 对外统一入口 ----------
def _open_word_document(file_path):
    """打开 .docx 文档（python-docx 仅支持 OOXML/.docx）。"""
    return Document(file_path)


def evaluate(dir_path):
    """按 §2.2 约定，返回结构化 dict；不 print、不改 sys.stdout、不 sys.exit。
    入参 dir_path 为“脚本所在目录的路径”，脚本自己在该目录里定位并打开被评估的
    .docx（忽略以 ~$ 开头的 Word 临时锁文件）。

    status='error' 只在“这份提交没被真正评过”时出现——目录不存在、目录内没有
    .docx、或脚本自身抛异常；此时 dim1_pass 为空串（未判定）、dim1_reason 为
    空串、dim2_items 为空表、total_score 与 max_score 一律为 0。
    文档能定位到但不达标（非可编辑文字、无表格版式），仍按 status='ok' +
    维度一未通过处理，此时 dim1_pass=False 且 dim1_reason 带有未通过的错误摘要。"""
    result = {
        "id": SCRIPT_ID,
        "file_name": "",
        "status": "ok",
        "error": None,
        # 空串表示“维度一尚未判定”；status='error' 的三条出口都保持空串，
        # 只有真正跑过维度一门禁才会写入 True / False。
        "dim1_pass": "",
        "dim1_reason": "",
        "dim2_items": [],
        "total_score": 0,
        "max_score": 0,
    }
    try:
        if not dir_path or not os.path.isdir(dir_path):
            result["status"] = "error"
            result["error"] = f"目录不存在：{dir_path}"
            return result
        cands = sorted(
            f for f in os.listdir(dir_path)
            if f.lower().endswith('.docx') and not f.startswith('~$')
        )
        if not cands:
            result["status"] = "error"
            result["error"] = f"目录中未找到 .docx 文件：{dir_path}"
            return result
        file_path = os.path.join(dir_path, cands[0])
        result["file_name"] = os.path.basename(file_path)
        return _evaluate_impl(file_path, result)
    except Exception as exc:
        # _evaluate_impl 是就地写 result 的，异常可能发生在它写了一半之后，
        # 因此这里把所有承载评分结论的字段显式复位，保证 error 的返回值是干净的全 0；
        # dim1_pass 复位为空串而不是 False——异常意味着维度一的结论作废，不是判定为不通过。
        result["status"] = "error"
        result["error"] = f"{type(exc).__name__}: {exc}"
        result["dim1_pass"] = ""
        result["dim1_reason"] = ""
        result["dim2_items"] = []
        result["total_score"] = 0
        result["max_score"] = 0
        return result


def _dim1_fail(result, reason):
    """维度一未通过：完成度总分记 0，不再输出维度二逐项。
    按统一约定，dim1_pass=False 时 dim1_reason 必须带上未通过的错误摘要。"""
    result["dim1_pass"] = False
    result["dim1_reason"] = reason or "维度一未通过（未给出具体原因）"
    result["dim2_items"] = []
    result["total_score"] = 0
    result["max_score"] = MAX_SCORE_TOTAL
    return result


def _evaluate_impl(file_path, result):
    """维度一 + 维度二的评估主体；内部按细则顺序构造 results / details 两个平行
    列表，末尾按 _MAX_DELTAS_ORDER 反解成结构化 dim2_items 填入 result。"""
    doc = _open_word_document(file_path)
    body = doc.element.body
    results = []
    details = []

    def _emit(delta, ok, rule, detail=""):
        """按“+N：说明 / -N：说明”协议登记一条评分项；未命中一律记 +0。"""
        results.append("%+d：%s" % (delta if ok else 0, rule))
        details.append(detail)

    # ========== 维度1：可用与可修改性 ==========
    # 只做门禁：能被 python-docx 正常打开（.docx/OOXML 且未损坏）、正文文字为可
    # 编辑文本而非图片、试卷版块为可编辑表格。任一不成立则完成度总分记 0。
    T = _t(body)
    NT = _norm(T)
    NQT = _nq(T)
    top_tbls = body.findall(qn("w:tbl"))
    if len(NT) < 300:
        return _dim1_fail(result, "文档内可编辑文字不足 300 字，疑为图片或扫描件")
    if not top_tbls:
        return _dim1_fail(result, "文档内未发现表格版式，试题版块不可编辑")

    # 维度一通过：按统一约定 dim1_reason 保持空串（只有未通过才写原因）。
    result["dim1_pass"] = True
    result["dim1_reason"] = ""

    # ========== 维度2：完成度评分 ==========
    # --- 版面结构定位（供各评分项复用）---
    flat = _paras(body)
    flat_norm = [_norm(_t(p)) for p in flat]
    NUMS = "一二三四五六七八九"

    # 九个大题标题段：按“汉字序号＋、”定位（拼音行里的 yī、èr 等为拉丁字母，不会误命中），
    # 同一序号只认首次出现的段落；标题与其拼音同段（中间用换行符分行）时也能定位到。
    title_pos = {}
    for i, s in enumerate(flat_norm):
        for m in re.finditer(r"([一二三四五六七八九])、", s):
            n = m.group(1)
            if n not in title_pos:
                title_pos[n] = i
    ordered = [title_pos.get(n) for n in NUMS]

    def _sec_paras(k):
        """第 k 个大题的“题区”：本大题标题段到下一个大题标题段之前的所有段落。
           按文档顺序切分，不依赖表格结构，因此模型把一个大题拆成多个版块时仍能取全。"""
        if ordered[k] is None:
            return []
        nxt = None
        for j in range(k + 1, 9):
            if ordered[j] is not None:
                nxt = ordered[j]
                break
        return flat[ordered[k]: nxt if nxt is not None else len(flat)]

    def _sec_text(k):
        return _norm("".join(_t(p) for p in _sec_paras(k)))

    def _outer_tbl(p):
        """段落所属的最外层表格；不在表格内返回 None。"""
        out = None
        node = p.getparent()
        while node is not None:
            if node.tag == qn("w:tbl"):
                out = node
            node = node.getparent()
        return out

    # 九个题块：九个大题标题段各自所属的最外层表格
    blocks = [(_outer_tbl(flat[i]) if i is not None else None) for i in ordered]

    # 卷头信息块：正文第一个顶层表格
    cover = top_tbls[0]

    # 作文格：列数不少于 15 的表格（方格作文纸）
    grid_tbl = None
    for t in _tbls(body):
        if len(_grid(t)) >= 15:
            grid_tbl = t
            break

    # 提示块：底纹为 FFFDEB 的顶层表格；找不到则退化为含“提示”二字的顶层表格
    tip = None
    for t in top_tbls:
        if TIP_SHADE in [_tc_shade(tc) for tr in _rows(t) for tc in _cells(tr)]:
            tip = t
            break
    if tip is None:
        for t in top_tbls:
            if t is not cover and "提示" in _t(t) and t not in blocks:
                tip = t
                break

    def _nested_tbls(tbl):
        """题块内用于排布词语/选项的嵌套表格（排除作文格与题块自身）。"""
        if tbl is None:
            return []
        return [x for x in _tbls(tbl)
                if x is not tbl and x is not grid_tbl and len(_grid(x)) < 15]

    def _prev_nonempty(i):
        """文档顺序上第 i 段之前最近的一个非空段落索引；没有则返回 None。"""
        j = i - 1
        while j >= 0:
            if flat_norm[j]:
                return j
            j -= 1
        return None

    # 大题标题 run 与其上方拼音行 run 的定位：
    #   办公软件里“标题上方一行拼音”有两种等效排布——拼音单独成段，或拼音与标题同段、
    #   中间用换行符（w:br）分行。两种都要认，否则会把合规排版误判为缺拼音。
    def _title_runs(k):
        """第 k 个大题标题段里属于中文标题的 run（含汉字的 run）。"""
        if ordered[k] is None:
            return []
        return [r for r in _runs(flat[ordered[k]]) if _has_cjk(_rt(r))]

    py_idx = []          # 拼音行所在段落索引（与标题同段时即标题段自身）
    py_runs = []         # 拼音行的 run 列表
    for k in range(9):
        if ordered[k] is None:
            py_idx.append(None)
            py_runs.append([])
            continue
        p = flat[ordered[k]]
        has_br = bool([x for x in p.iter(qn("w:br")) if not _in_fallback(x)])
        same = []
        for r in _runs(p):
            if _has_cjk(_rt(r)):
                break
            if _norm(_rt(r)) and _PINYIN_ONLY_RE.match(_norm(_rt(r))):
                same.append(r)
        if same and has_br:
            py_idx.append(ordered[k])
            py_runs.append(same)
            continue
        j = _prev_nonempty(ordered[k])
        if j is not None and _PINYIN_ONLY_RE.match(flat_norm[j] or "x"):
            py_idx.append(j)
            py_runs.append([r for r in _runs(flat[j]) if _norm(_rt(r))])
        else:
            py_idx.append(None)
            py_runs.append([])

    # +5：九个大题齐全且题号一至九连续无缺项，标题依次为“一、看拼音，写词语。（12分）”…“九、习作。”
    # 判定：九条标题文字逐条在正文里出现（比对前去掉全部空白与各类引号，规避直/弯引号
    #   与排版空格差异），且九条的首次出现位置严格递增，即题号一至九按序连续无缺项。
    pos = [NQT.find(_nq(x)) for x in NINE_TITLES]
    miss = [NUMS[i] for i, v in enumerate(pos) if v < 0]
    ok = (not miss) and all(pos[i] < pos[i + 1] for i in range(8))
    _emit(5, ok,
          "九个大题齐全且题号一至九连续无缺项，标题依次为“一、看拼音，写词语。（12分）”"
          "“二、给加点字选择正确读音，画“√”。（6分）”“三、把词语补充完整，并选词填空。（10分）”"
          "“四、按要求填一填。（10分）”“五、词语闯关。（9分）”“六、按要求写句子。（8分）”"
          "“七、积累与运用。（10分）”“八、阅读直通车。（20分）”“九、习作。”",
          "缺失题号=%s" % ("无" if not miss else "".join(miss)))

    # +1：卷头主标题文字为“三年级下册语文达标测试卷”
    ok = MAIN_TITLE in NT
    _emit(1, ok,
          "卷头主标题文字为“三年级下册语文达标测试卷”",
          "命中=%s" % ok)

    # +1：卷头拼音行文字为“sān nián jí xià cè yǔ wén dá biāo cè shì juàn”
    ok = _norm(COVER_PINYIN) in NT
    _emit(1, ok,
          "卷头拼音行文字为“sān nián jí xià cè yǔ wén dá biāo cè shì juàn”",
          "命中=%s" % ok)

    # +1：卷头副标题含“第四单元测试卷”与“时间：60分钟　满分：100分”文本
    # 判定：文字内容为准，空格排布差异不计入（比对前已去掉全部空白）。
    hit = [x for x in SUB_TITLE_PARTS if _norm(x) in NT]
    ok = len(hit) == len(SUB_TITLE_PARTS)
    _emit(1, ok,
          "卷头副标题含“第四单元测试卷”与“时间：60分钟　满分：100分”文本",
          "命中=%d/2" % len(hit))

    # +3：卷头右栏含“学校：____________”“班级：____________”“姓名：____________”“得分：____________”四行填写栏
    labels = set()
    for lab in FILL_LABELS:
        if re.search(lab + "[：:]" + _UNDERLINE_RE + "{6,}", NT):
            labels.add(lab)
    ok = len(labels) == 4
    _emit(3, ok,
          "卷头右栏含“学校：____________”“班级：____________”“姓名：____________”"
          "“得分：____________”四行填写栏",
          "命中填写栏=%s" % ("".join(x for x in FILL_LABELS if x in labels) or "无"))

    # +3：九个大题分别置于一个1行1列表格块内，块的网格总宽18.40cm，块的框线不粗于0.5磅（sz≤4）且单元格之间不出现网格线
    # 与办公软件对齐：网格总宽取 w:tblGrid 各列之和（18.40cm = 10432 twips，容差 ±1%）；
    #   框线取“tc 侧边框优先于 tblBorders 对应边”的生效值（单元格显式声明 nil 即不画线）；
    #   单元格之间的网格线按实际渲染判定——1 行 1 列的块不会渲染出任何内部横线/竖线，
    #   仅在表级声明了 insideH/insideV 而未渲染，不算破坏版式。
    det = []
    ok = all(b is not None for b in blocks)
    for k, b in enumerate(blocks):
        if b is None:
            det.append("%s:不在表格内" % NUMS[k])
            continue
        rows = _rows(b)
        grid = _grid(b)
        w = sum(grid)
        one_cell = len(rows) == 1 and len(_cells(rows[0])) == 1
        width_ok = abs(w - BLOCK_WIDTH) <= BLOCK_WIDTH * 0.01
        thin, inner = True, False
        for kind, val, sz, _c in _rendered_borders(b):
            if not _has_line(val):
                continue
            if kind in ("insideH", "insideV"):
                inner = True
            elif not (val == "single" and sz is not None and sz <= 4):
                thin = False
        if not (one_cell and width_ok and thin and not inner):
            ok = False
            det.append("%s:%dr%dc/%dtwips/细线=%s/内网格=%s"
                       % (NUMS[k], len(rows), len(grid), w, thin, inner))
    _emit(3, ok,
          "九个大题分别置于一个1行1列表格块内，块的网格总宽18.40cm，"
          "块的框线不粗于0.5磅（sz≤4）且单元格之间不出现网格线",
          "；".join(det) if det else "九个题块均合规")

    # +3：九个题块底纹按三题一组依次为F2F8FF（一、二、三）、F8FFF2（四、五、六）、FFF7F0（七、八、九）
    got = []
    for b in blocks:
        fill = None
        if b is not None:
            rows = _rows(b)
            if rows and _cells(rows[0]):
                fill = _tc_shade(_cells(rows[0])[0])
        got.append(fill)
    ok = tuple(got) == BLOCK_SHADES
    _emit(3, ok,
          "九个题块底纹按三题一组依次为F2F8FF（一、二、三）、F8FFF2（四、五、六）、"
          "FFF7F0（七、八、九）",
          "实测底纹=%s" % ",".join(str(x) for x in got))

    # +3：九个大题的中文标题均为Noto Sans CJK SC、13.5磅、加粗，颜色按三题一组依次为2F75B5、548235、C55A11
    # 与办公软件对齐：字体沿 w:rFonts/@w:eastAsia 继承链解析（run → 段落样式 → 默认样式），
    #   字号读 w:sz（半磅）、加粗读 w:b、颜色读 w:color，即 Word/WPS 打开后看到的效果。
    #   只校验标题段里含汉字的 run，同段的拼音 run 由下一条单独校验。
    det = []
    ok = True
    for k in range(9):
        i = ordered[k]
        if i is None:
            ok = False
            det.append("%s:标题缺失" % NUMS[k])
            continue
        p = flat[i]
        runs = _title_runs(k)
        if not runs:
            ok = False
            det.append("%s:标题无中文" % NUMS[k])
            continue
        bad = []
        for r in runs:
            sz = _eff_size(doc, p, r)
            col = _eff_color(doc, p, r)
            if sz is None or abs(sz - 13.5) > 0.25:
                bad.append("字号%s" % sz)
            if col != TITLE_COLORS[k]:
                bad.append("颜色%s" % col)
            if not _eff_flag(doc, p, r, "b"):
                bad.append("未加粗")
            if SANS not in (_eff_font(doc, p, r, "eastAsia") or ""):
                bad.append("中文字体%s" % _eff_font(doc, p, r, "eastAsia"))
        if bad:
            ok = False
            det.append("%s:%s" % (NUMS[k], "/".join(sorted(set(bad)))))
    _emit(3, ok,
          "九个大题的中文标题均为Noto Sans CJK SC、13.5磅、加粗，"
          "颜色按三题一组依次为2F75B5、548235、C55A11",
          "；".join(det) if det else "九条标题均合规")

    # +3：九个大题的中文标题上方均有一行对应拼音，为Times New Roman、8.5磅、斜体、颜色646464
    det = []
    ok = True
    for k in range(9):
        if py_idx[k] is None or not py_runs[k]:
            ok = False
            det.append("%s:标题上方无拼音行" % NUMS[k])
            continue
        p = flat[py_idx[k]]
        bad = []
        for r in py_runs[k]:
            s = _rt(r)
            sz = _eff_size(doc, p, r)
            if sz is None or abs(sz - 8.5) > 0.25:
                bad.append("字号%s" % sz)
            if _eff_color(doc, p, r) != TITLE_PY_COLOR:
                bad.append("颜色%s" % _eff_color(doc, p, r))
            if not _eff_flag(doc, p, r, "i"):
                bad.append("非斜体")
            if _has_latin(s) and TNR not in (_eff_font(doc, p, r, "ascii") or ""):
                bad.append("西文字体%s" % _eff_font(doc, p, r, "ascii"))
        if bad:
            ok = False
            det.append("%s:%s" % (NUMS[k], "/".join(sorted(set(bad)))))
    _emit(3, ok,
          "九个大题的中文标题上方均有一行对应拼音，为Times New Roman、8.5磅、斜体、颜色646464",
          "；".join(det) if det else "九行拼音均合规")

    # +3：九个题块内各题干、选项与填空文字均为Noto Serif CJK SC、10.5磅
    # 判定范围：题块内除大题标题段与其上方拼音行以外的全部段落；只校验含中文的 run
    #   （题干、选项、填空均为中文），拼音、题号域等西文 run 不纳入。
    det = []
    ok = True
    for k, b in enumerate(blocks):
        if b is None:
            ok = False
            continue
        skip = set()
        if ordered[k] is not None:
            skip.add(id(flat[ordered[k]]))
        if py_idx[k] is not None:
            skip.add(id(flat[py_idx[k]]))
        bad = []
        for p in _paras(b):
            if id(p) in skip:
                continue
            for r in _runs(p):
                s = _rt(r)
                if not _has_cjk(s):
                    continue
                sz = _eff_size(doc, p, r)
                ea = _eff_font(doc, p, r, "eastAsia")
                if sz is None or abs(sz - 10.5) > 0.25:
                    bad.append("字号%s" % sz)
                if SERIF not in (ea or ""):
                    bad.append("字体%s" % ea)
        if bad:
            ok = False
            det.append("%s:%s" % (NUMS[k], "/".join(sorted(set(bad))[:4])))
    _emit(3, ok,
          "九个题块内各题干、选项与填空文字均为Noto Serif CJK SC、10.5磅",
          "；".join(det) if det else "题块内中文文字均合规")

    # +3：九个题块内用于排布词语、选项的嵌套表格，其单元格边框为single细线、颜色沿用蓝色块D4E7F8与绿色块CAE5B7
    # 判定：题块内的嵌套表格（排除作文格）实际渲染出的每条格线都必须是 single、
    #   线宽 sz≤8（1磅及以下的细线）、颜色属于版式基线配色；至少要存在一张这样的嵌套表格。
    nested = []
    for b in blocks:
        nested += _nested_tbls(b)
    det = []
    ok = len(nested) > 0 and all(b is not None for b in blocks)
    for t in nested:
        for kind, val, sz, col in _rendered_borders(t):
            if not _has_line(val):
                continue
            if not (val == "single" and sz is not None and sz <= 8
                    and (col or "").upper() in NESTED_BORDER_COLORS):
                ok = False
                det.append("%s/sz%s/%s" % (val, sz, col))
    _emit(3, ok,
          "九个题块内用于排布词语、选项的嵌套表格，其单元格边框为single细线、"
          "颜色沿用蓝色块D4E7F8与绿色块CAE5B7",
          "嵌套表=%d张 越界格线=%s" % (len(nested), ",".join(sorted(set(det))[:4]) or "无"))

    # +5：一题8个词语拼音“wǔ dǎo”…“xiū jiàn”齐全，且每个拼音下方配田字格书写区
    # 书写区计数：一题块内的图形对象数（田字格图片），或嵌套表格里“无文字的空白单元格”
    #   数（方格表格形式的书写格），取两者较大值，需不少于 8 处。
    sec1 = _sec_text(0)
    b1 = blocks[0]
    py_hit = [x for x in Q1_PINYIN if _norm(x) in sec1]
    n_draw = _drawing_count(b1) if b1 is not None else 0
    n_empty = 0
    if b1 is not None:
        for t in _nested_tbls(b1):
            for tr in _rows(t):
                for tc in _cells(tr):
                    if not _norm(_t(tc)):
                        n_empty += 1
    n_area = max(n_draw, n_empty)
    boxes = _t(b1).count("□") if b1 is not None else 0
    ok = len(py_hit) == 8 and n_area >= 8
    _emit(5, ok,
          "一题8个词语拼音“wǔ dǎo”“zhì huì”“xīn shǎng”“bǎo cún”“nuò ruò”“jīng yàn”"
          "“mì fēng”“xiū jiàn”齐全，且每个拼音下方配田字格书写区",
          "拼音=%d/8 书写区=%d处（图形%d/空白格%d）" % (len(py_hit), n_area, n_draw, n_empty))

    # +3：一题田字格书写区为图片或方格表格形式的书写格，未用“□□□□”等文字方框替代
    ok = n_area >= 8 and boxes == 0
    _emit(3, ok,
          "一题田字格书写区为图片或方格表格形式的书写格，未用“□□□□”等文字方框替代",
          "书写区=%d处 文字方框□=%d个" % (n_area, boxes))

    # +3：二题6组加点字“挣脱”“切开”“模样”“掠过”“闲散”“纤细”齐全，每组给出两个待选读音
    # 待选读音计数：二题区内除大题拼音行以外的拼音音节 token 数不少于 12（6 组 × 2 个）。
    sec2_paras = [p for idx, p in zip(range(len(_sec_paras(1))), _sec_paras(1))]
    sec2 = _sec_text(1)
    w_hit = [x for x in Q2_WORDS if x in sec2]
    tokens = []
    for p in sec2_paras:
        if py_idx[1] is not None and p is flat[py_idx[1]]:
            continue
        for m in re.finditer(r"[A-Za-zāáǎàēéěèīíǐìōóǒòūúǔùǖǘǚǜüńňǹ]{2,}", _t(p)):
            tokens.append(m.group(0))
    ok = len(w_hit) == 6 and len(tokens) >= 12
    _emit(3, ok,
          "二题6组加点字“挣脱”“切开”“模样”“掠过”“闲散”“纤细”齐全，每组给出两个待选读音",
          "加点字=%d/6 待选读音=%d个" % (len(w_hit), len(tokens)))

    # +5：二题“挣”“切”“模”“掠”“散”“纤”六个受考字下方带着重点标记
    # 与办公软件对齐：着重号可用字符格式 w:em（着重号）实现，也可在字下方叠加小圆点图形，
    #   两种做法在页面上都是“受考字下方一个点”，故两者都认；后者按点的个数是否覆盖六组计。
    marked = set()
    for p in _sec_paras(1):
        for r in _runs(p):
            if not _has_em(r):
                continue
            for ch in Q2_CHARS:
                if ch in _rt(r):
                    marked.add(ch)
    n_dot = _dot_mark_count(blocks[1])
    ok = set(Q2_CHARS) <= marked or n_dot >= len(Q2_CHARS)
    _emit(5, ok,
          "二题“挣”“切”“模”“掠”“散”“纤”六个受考字下方带着重点标记",
          "着重号字符格式=%s 圆点标记=%d个" % (
              "".join(x for x in Q2_CHARS if x in marked) or "无", n_dot))

    # +3：三题①邯郸②杯弓③争奇④翩翩⑤学富⑥守株六个待补词语齐全，并保留其后两句选词填空
    # 填空处既认下划线段（＿＿＿＿），也认括号空位（（　）），两者都是卷面上的作答空位。
    sec3 = _sec_text(2)
    w_hit = [x for x in Q3_WORDS if x in sec3]
    tail = sec3.split("守株", 1)[-1] if "守株" in sec3 else ""
    n_blank = (len(re.findall(_UNDERLINE_RE + "{4,}", tail))
               + len(re.findall(r"[（(][\s　]*[）)]", tail)))
    n_item = len(re.findall(r"[1１][.．、]", tail)) + len(re.findall(r"[2２][.．、]", tail))
    ok = len(w_hit) == 6 and n_item >= 2
    _emit(3, ok,
          "三题①邯郸②杯弓③争奇④翩翩⑤学富⑥守株六个待补词语齐全，并保留其后两句选词填空",
          "待补词语=%d/6 填空句=%d 填空处=%d" % (len(w_hit), n_item, n_blank))

    # +3：四题含“鹿”字笔画题、用音序查字法查“骄”字题及（1）（2）两句关联词填空
    sec4 = _sec_text(3)
    c1 = "鹿" in sec4 and "笔" in sec4
    c2 = "音序" in sec4 and "骄" in sec4
    c3 = "（1）" in sec4 and "（2）" in sec4
    ok = c1 and c2 and c3
    _emit(3, ok,
          "四题含“鹿”字笔画题、用音序查字法查“骄”字题及（1）（2）两句关联词填空",
          "笔画题=%s 音序题=%s 关联词填空=%s" % (c1, c2, c3))

    # +3：五题［第一关］（4分）含“安安静静”“又细又长”两组仿写，［第二关］（5分）含①阵②座③本④片⑤朵⑥只六个量词及“一（　）花园”等5处填空
    sec5 = _sec_text(4)
    p_hit = [x for x in Q5_PARTS if x in sec5]
    q_hit = [x for x in Q5_QUANTIFIERS if x in sec5]
    n_paren = len(re.findall(r"（[^）]{0,4}）", sec5))
    ok = len(p_hit) == 4 and len(q_hit) == 6 and n_paren >= 5
    _emit(3, ok,
          "五题［第一关］（4分）含“安安静静”“又细又长”两组仿写，［第二关］（5分）含"
          "①阵②座③本④片⑤朵⑥只六个量词及“一（　）花园”等5处填空",
          "两关与仿写=%d/4 量词=%d/6 填空括号=%d处" % (len(p_hit), len(q_hit), n_paren))

    # +3：六题第1小题（4分）保留例句“花开了。”“小燕子飞过田野。”，第2小题（4分）为用“像”写一个比喻句
    sec6 = _sec_text(5)
    p_hit = [x for x in Q6_PARTS if x in sec6]
    ok = len(p_hit) == 3 and "像" in sec6
    _emit(3, ok,
          "六题第1小题（4分）保留例句“花开了。”“小燕子飞过田野。”，"
          "第2小题（4分）为用“像”写一个比喻句",
          "例句与比喻句要素=%d/3 含“像”=%s" % (len(p_hit), "像" in sec6))

    # +3：七题［第一关］（6分）含两句古诗填空及“观察事物要＿＿、＿＿”，［第二关］（4分）含四个选项
    sec7 = _sec_text(6)
    p_hit = [x for x in Q7_POEMS if x in sec7]
    o_hit = [x for x in Q7_OPTIONS if x in sec7]
    ok = len(p_hit) == 3 and len(o_hit) == 4
    _emit(3, ok,
          "七题［第一关］（6分）含“迟日江山丽，＿＿”“独在异乡为异客，＿＿”两句古诗填空及"
          "“观察事物要＿＿、＿＿”，［第二关］（4分）含①认真观察②随便猜测③连续记录④只看一次四个选项",
          "第一关要素=%d/3 第二关选项=%d/4" % (len(p_hit), len(o_hit)))

    # +5：八题短文“会写日记的银杏树”4个自然段全文齐全
    # 判定：八题区内出现短文标题，且标题之后有不少于 4 个正文长度≥60 字的自然段。
    sec8_paras = _sec_paras(7)
    title_seen = False
    seg = 0
    for p in sec8_paras:
        s = _norm(_t(p))
        if not title_seen:
            if Q8_TITLE in s:
                title_seen = True
            continue
        if len(s) >= 60:
            seg += 1
    ok = title_seen and seg >= 4
    _emit(5, ok,
          "八题短文“会写日记的银杏树”4个自然段全文齐全",
          "短文标题=%s 自然段(≥60字)=%d个" % (title_seen, seg))

    # +1：八题短文标题“会写日记的银杏树”单独成段、居中、加粗
    ok = False
    det = "未找到独立标题段"
    for p in sec8_paras:
        if _norm(_t(p)) != Q8_TITLE:
            continue
        runs = [r for r in _runs(p) if _norm(_rt(r))]
        centered = _jc(p) == "center"
        bold = bool(runs) and all(_eff_flag(doc, p, r, "b") for r in runs)
        det = "单独成段=True 居中=%s 加粗=%s" % (centered, bold)
        ok = centered and bold
        break
    _emit(1, ok, "八题短文标题“会写日记的银杏树”单独成段、居中、加粗", det)

    # +3：八题4个小题及其分值（4分、4分、6分、6分）齐全，且与短文同处一个题块内，未拆分为两个版块
    # “未拆分为两个版块”判定：短文所在的最外层表格与四个小题所在的最外层表格必须是同一个。
    sec8 = _sec_text(7)
    n4 = sec8.count("（4分）")
    n6 = sec8.count("（6分）")
    host = None
    same = True
    for p in sec8_paras:
        s = _norm(_t(p))
        if not (Q8_TITLE in s or re.match(r"^[1-4１-４][.．、]", s)):
            continue
        t = _outer_tbl(p)
        if host is None:
            host = t
        elif t is not host:
            same = False
    ok = n4 >= 2 and n6 >= 2 and same and host is not None
    _emit(3, ok,
          "八题4个小题及其分值（4分、4分、6分、6分）齐全，且与短文同处一个题块内，未拆分为两个版块",
          "（4分）=%d处 （6分）=%d处 短文与小题同块=%s" % (n4, n6, same))

    # +3：九题写作要求含“观察一种植物或一种小动物”与“不少于150字”
    sec9 = _sec_text(8)
    p_hit = [x for x in Q9_PARTS if _norm(x) in sec9]
    _emit(3, len(p_hit) == 2,
          "九题写作要求含“观察一种植物或一种小动物”与“不少于150字”",
          "写作要求=%d/2" % len(p_hit))

    # +5：九题配20列×6行的方格作文纸，每格宽0.85cm、行高不小于0.83cm
    # 与办公软件对齐：列数与格宽读 w:tblGrid（0.85cm = 482 twips，容差 ±4%）；
    #   行高读 w:trHeight（0.83cm = 470 twips），hRule 为 exact 时按固定值判定。
    if grid_tbl is None:
        _emit(5, False, "九题配20列×6行的方格作文纸，每格宽0.85cm、行高不小于0.83cm",
              "未找到方格作文纸")
        _emit(3, False,
              "九题作文格边框为single 0.25磅（sz=2）、颜色737373的细线，"
              "单元格上下左右内边距为0，格内文字分散对齐", "未找到方格作文纸")
    else:
        g = _grid(grid_tbl)
        rows = _rows(grid_tbl)
        w_bad = [x for x in g if abs(x - CELL_WIDTH) > CELL_WIDTH * 0.04]
        h_list = [_row_height(tr) for tr in rows]
        h_bad = [h for h in h_list if h[0] is None or h[0] < CELL_HEIGHT]
        ok = len(g) == 20 and len(rows) == 6 and not w_bad and not h_bad
        _emit(5, ok,
              "九题配20列×6行的方格作文纸，每格宽0.85cm、行高不小于0.83cm",
              "%d列×%d行 格宽=%s 行高=%s"
              % (len(g), len(rows), sorted(set(g)), sorted(set(h_list))))

        # +3：九题作文格边框为single 0.25磅（sz=2）、颜色737373的细线，单元格上下左右内边距为0，格内文字分散对齐
        # 判定对象是格与格之间的书写格线；作文纸最外圈的整体外框允许另用较重的线，
        #   只要求它存在，不并入细线判定。
        bd_bad = set()
        for kind, val, sz, col in _rendered_borders(grid_tbl):
            if not _has_line(val):
                bd_bad.add("缺线")
            elif kind in ("insideH", "insideV") and not (
                    val == "single" and sz is not None and sz <= 3
                    and (col or "").upper() == GRID_BORDER_COLOR):
                bd_bad.add("%s/sz%s/%s" % (val, sz, col))
        mar_bad = set()
        jc_bad = 0
        n_cell = 0
        for tr in rows:
            for tc in _cells(tr):
                n_cell += 1
                for v in _cell_margins(grid_tbl, tc):
                    if v != 0:
                        mar_bad.add(str(v))
                ps = tc.findall(qn("w:p"))
                if not ps or _jc(ps[0]) != "distribute":
                    jc_bad += 1
        ok = not bd_bad and not mar_bad and jc_bad == 0 and n_cell > 0
        _emit(3, ok,
              "九题作文格边框为single 0.25磅（sz=2）、颜色737373的细线，"
              "单元格上下左右内边距为0，格内文字分散对齐",
              "越界格线=%s 内边距异常=%s 非分散对齐格=%d/%d"
              % (",".join(sorted(bd_bad)[:3]) or "无",
                 ",".join(sorted(mar_bad)) or "无", jc_bad, n_cell))

    # +1：九个大题的试题内容完整置入，且未附带“参考答案”页及答案内容
    empty = [NUMS[k] for k in range(9)
             if len([p for p in _sec_paras(k) if _norm(_t(p))]) < 2]
    ok = not empty and "参考答案" not in NT
    _emit(1, ok,
          "九个大题的试题内容完整置入，且未附带“参考答案”页及答案内容",
          "空大题=%s 参考答案=%s" % ("无" if not empty else "".join(empty),
                                    "参考答案" in NT))

    # ---------- 减分项：指令未涉及、却改动了模板既有版式的多余修改 ----------
    # 约定：_emit 的第二个参数为“违规是否成立”，成立才计入负分，否则记 +0。
    hdrs = _hdrftr_elements(doc, "header")
    ftrs = _hdrftr_elements(doc, "footer")
    hdr_txt = _nq("".join(_t(h) for h in hdrs))

    # -3：页眉文字不为“pīn yīn xiǎo liàn xí · měi tiān jìn bù yì diǎn diǎn ｜ 拼音小练习 · 每天进步一点点”，被改写或删除
    bad_hdr = _nq(HEADER_TEXT) not in hdr_txt
    _emit(-3, bad_hdr,
          "页眉文字不为“pīn yīn xiǎo liàn xí · měi tiān jìn bù yì diǎn diǎn ｜ "
          "拼音小练习 · 每天进步一点点”，被改写或删除",
          "页眉实测=%r" % ("".join(_t(h) for h in hdrs)[:60] if hdrs else "无页眉"))

    # -1：页眉文字西文字体不为Times New Roman、中文不为Noto Serif CJK SC，或字号不为8.5磅、颜色不为767676、不为居中
    det = []
    if not hdrs:
        det.append("无页眉")
    for h in hdrs:
        for p in _paras(h):
            runs = [r for r in _runs(p) if _norm(_rt(r))]
            if not runs:
                continue
            if _jc(p) != "center":
                det.append("对齐%s" % _jc(p))
            for r in runs:
                s = _rt(r)
                sz = _eff_size(doc, p, r)
                if sz is None or abs(sz - 8.5) > 0.25:
                    det.append("字号%s" % sz)
                if _eff_color(doc, p, r) != HEADER_COLOR:
                    det.append("颜色%s" % _eff_color(doc, p, r))
                if _has_latin(s) and TNR not in (_eff_font(doc, p, r, "ascii") or ""):
                    det.append("西文%s" % _eff_font(doc, p, r, "ascii"))
                if _has_cjk(s) and SERIF not in (_eff_font(doc, p, r, "eastAsia") or ""):
                    det.append("中文%s" % _eff_font(doc, p, r, "eastAsia"))
    _emit(-1, bool(det),
          "页眉文字西文字体不为Times New Roman、中文不为Noto Serif CJK SC，"
          "或字号不为8.5磅、颜色不为767676、不为居中",
          ",".join(sorted(set(det))[:5]) or "页眉格式与模板一致")

    # -1：页脚不为居中的“第 X 页 / 共 Y 页”页码域，或字体不为Noto Serif CJK SC、8.5磅、颜色不为787878
    # 与办公软件对齐：页码为 PAGE / NUMPAGES 域，域结果随页变化，故只校验域指令本身；
    #   居中既认段落 w:jc=center，也认浮动文本框相对页边距水平居中（wp:align=center）。
    def _centered(el):
        for p in _paras(el):
            if _jc(p) == "center":
                return True
        for a in el.iter():
            if str(a.tag).endswith("}align") and (a.text or "").strip() == "center":
                return True
        return False

    det = []
    if not ftrs:
        det.append("无页脚")
    for f in ftrs:
        instr = " ".join((x.text or "") for x in f.iter(qn("w:instrText"))).upper()
        if "PAGE" not in instr:
            det.append("缺当前页码域")
        if "NUMPAGES" not in instr:
            det.append("缺总页数域")
        s = _norm(_t(f))
        if not ("第" in s and "共" in s and "页" in s):
            det.append("页码文字=%r" % s[:20])
        if not _centered(f):
            det.append("非居中")
        n_hit = 0
        for p in _paras(f):
            for r in _runs(p):
                if not _norm(_rt(r)):
                    continue
                sz = _eff_size(doc, p, r)
                if sz is None or abs(sz - 8.5) > 0.25:
                    det.append("字号%s" % sz)
                col = _eff_color(doc, p, r)
                if col == FOOTER_COLOR:
                    n_hit += 1
                elif col is not None:
                    det.append("颜色%s" % col)
                if _has_cjk(_rt(r)) and SERIF not in (_eff_font(doc, p, r, "eastAsia") or ""):
                    det.append("中文%s" % _eff_font(doc, p, r, "eastAsia"))
        if n_hit == 0:
            det.append("无787878文字")
    _emit(-1, bool(det),
          "页脚不为居中的“第 X 页 / 共 Y 页”页码域，或字体不为Noto Serif CJK SC、"
          "8.5磅、颜色不为787878",
          ",".join(sorted(set(det))[:5]) or "页脚格式与模板一致")

    # -3：底纹为FFFDEB的提示块内文字不为“闯关提示：请认真拼读每一道题，书写要端正，遇到轻声和整体认读音节要多读几遍。”
    tip_txt = _nq(_t(tip)) if tip is not None else ""
    bad_tip = _nq(TIP_TEXT) not in tip_txt
    _emit(-3, bad_tip,
          "底纹为FFFDEB的提示块内文字不为“闯关提示：请认真拼读每一道题，书写要端正，"
          "遇到轻声和整体认读音节要多读几遍。”",
          "提示块实测=%r" % (_norm(_t(tip))[:50] if tip is not None else "无提示块"))

    # -1：提示块中“闯关提示：”不为Noto Sans CJK SC、10.5磅、加粗、颜色C07000
    det = []
    found = False
    if tip is None:
        det.append("无提示块")
    else:
        for p in _paras(tip):
            for r in _runs(p):
                s = _norm(_rt(r))
                pre = _norm(TIP_PREFIX)
                if not s or (pre not in s and s not in pre):
                    continue
                found = True
                sz = _eff_size(doc, p, r)
                ea = _eff_font(doc, p, r, "eastAsia")
                if sz is None or abs(sz - 10.5) > 0.25:
                    det.append("字号%s" % sz)
                if SANS not in (ea or ""):
                    det.append("字体%s" % ea)
                if not _eff_flag(doc, p, r, "b"):
                    det.append("未加粗")
                if _eff_color(doc, p, r) != TIP_PREFIX_COLOR:
                    det.append("颜色%s" % _eff_color(doc, p, r))
        if not found:
            det.append("未找到“闯关提示：”")
    _emit(-1, bool(det),
          "提示块中“闯关提示：”不为Noto Sans CJK SC、10.5磅、加粗、颜色C07000",
          ",".join(sorted(set(det))[:5]) or "提示语格式与模板一致")

    # -1：提示块底纹不为FFFDEB，或提示块被加上可见边框线
    # 版式基线里提示块四周本就有一条与底纹几乎同色的浅描边（E2EAF2 细线），它不算“可见边框”；
    #   凡改为基线配色之外的颜色，即为另加了看得见的框线。
    det = []
    if tip is None:
        det.append("无提示块")
    else:
        shades = [_tc_shade(tc) for tr in _rows(tip) for tc in _cells(tr)]
        if TIP_SHADE not in shades:
            det.append("底纹=%s" % ",".join(str(x) for x in shades[:3]))
        for kind, val, sz, col in _rendered_borders(tip):
            if _has_line(val) and (col or "").upper() not in BASE_BORDER_COLORS:
                det.append("%s线%s/sz%s/%s" % (kind, val, sz, col))
    _emit(-1, bool(det),
          "提示块底纹不为FFFDEB，或提示块被加上可见边框线",
          ",".join(sorted(set(det))[:5]) or "提示块底纹与无边框版式保持")

    # ---------- 卷头信息块定位（左栏标题区 / 右栏填写区）----------
    cv_rows = _rows(cover)
    cv_cells = _cells(cv_rows[0]) if cv_rows else []
    cv_paras = _paras(cover)

    def _cv_pick(pred):
        """卷头信息块内第一个满足 pred(段文字) 的段落。"""
        for p in cv_paras:
            if pred(_norm(_t(p))):
                return p
        return None

    def _fmt_check(p, font, kind, size, color, bold=None, italic=None):
        """按“实际生效值”校验一个段落里所有有字 run 的字体/字号/颜色/加粗/斜体，
           返回不合规描述列表；kind 决定比 eastAsia 还是 ascii 字体。"""
        out = []
        if p is None:
            return ["段落缺失"]
        runs = [r for r in _runs(p) if _norm(_rt(r))]
        if not runs:
            return ["段落无文字"]
        for r in runs:
            s = _rt(r)
            sz = _eff_size(doc, p, r)
            if size is not None and (sz is None or abs(sz - size) > 0.25):
                out.append("字号%s" % sz)
            if color is not None and _eff_color(doc, p, r) != color:
                out.append("颜色%s" % _eff_color(doc, p, r))
            if bold is not None and _eff_flag(doc, p, r, "b") != bold:
                out.append("加粗%s" % _eff_flag(doc, p, r, "b"))
            if italic is not None and _eff_flag(doc, p, r, "i") != italic:
                out.append("斜体%s" % _eff_flag(doc, p, r, "i"))
            if font is not None:
                need = _has_cjk(s) if kind == "eastAsia" else _has_latin(s)
                if need and font not in (_eff_font(doc, p, r, kind) or ""):
                    out.append("字体%s" % _eff_font(doc, p, r, kind))
        return out

    # -3：卷头信息块左单元格底纹不为EAF5FF、右单元格底纹不为FDF2E9
    got = [_tc_shade(tc) for tc in cv_cells[:2]]
    bad_shd = tuple(got) != COVER_SHADES
    _emit(-3, bad_shd,
          "卷头信息块左单元格底纹不为EAF5FF、右单元格底纹不为FDF2E9",
          "实测底纹=%s" % ",".join(str(x) for x in got) if got else "未取到单元格")

    # -1：卷头信息块两个单元格内容不为垂直居中
    va = [_tc_valign(tc) for tc in cv_cells[:2]]
    bad_va = len(va) < 2 or any(x != "center" for x in va)
    _emit(-1, bad_va,
          "卷头信息块两个单元格内容不为垂直居中",
          "实测垂直对齐=%s" % ",".join(str(x) for x in va))

    # -3：卷头拼音行不为Times New Roman、9磅、斜体、颜色56739A
    # 定位：卷头信息块内第一个“只由拼音字母与标点组成”的段落，即主标题上方的拼音行。
    p_py = _cv_pick(lambda s: bool(s) and bool(_PINYIN_ONLY_RE.match(s)))
    det = _fmt_check(p_py, TNR, "ascii", 9.0, COVER_PY_COLOR, italic=True)
    _emit(-3, bool(det),
          "卷头拼音行不为Times New Roman、9磅、斜体、颜色56739A",
          ",".join(sorted(set(det))[:5]) or "卷头拼音行格式与模板一致")

    # -3：卷头主标题不为Noto Sans CJK SC、24磅、加粗、颜色1F4E79
    # 定位：优先取含目标标题文字的段落；若标题文字被改写，则取卷头内字号最大的中文段落，
    #   即模板主标题所在的那一行，据此判断该行原有格式是否被破坏。
    p_main = _cv_pick(lambda s: MAIN_TITLE in s)
    if p_main is None:
        best = None
        for p in cv_paras:
            if not _has_cjk(_t(p)):
                continue
            szs = [_eff_size(doc, p, r) for r in _runs(p) if _norm(_rt(r))]
            szs = [x for x in szs if x is not None]
            if szs and (best is None or max(szs) > best[0]):
                best = (max(szs), p)
        p_main = best[1] if best else None
    det = _fmt_check(p_main, SANS, "eastAsia", 24.0, COVER_TITLE_COLOR, bold=True)
    _emit(-3, bool(det),
          "卷头主标题不为Noto Sans CJK SC、24磅、加粗、颜色1F4E79",
          ",".join(sorted(set(det))[:5]) or "卷头主标题格式与模板一致")

    # -1：卷头副标题不为Noto Sans CJK SC、10.5磅、颜色5B5B5B
    # 定位：优先取含目标副标题文字的段落；副标题文字被改写时，退回主标题下方紧邻的那一行
    #   中文段落（排除“：____”填写行），据此判断该行原有格式是否被破坏。
    p_sub = _cv_pick(lambda s: any(_norm(x) in s for x in SUB_TITLE_PARTS))
    if p_sub is None:
        p_sub = _cv_pick(lambda s: "测试卷" in s and MAIN_TITLE not in s)
    if p_sub is None and p_main is not None and p_main in cv_paras:
        for p in cv_paras[cv_paras.index(p_main) + 1:]:
            s = _norm(_t(p))
            if not s or not _has_cjk(s):
                continue
            if re.search("[：:]" + _UNDERLINE_RE + "{4,}", s):
                continue
            p_sub = p
            break
    det = _fmt_check(p_sub, SANS, "eastAsia", 10.5, COVER_SUB_COLOR)
    _emit(-1, bool(det),
          "卷头副标题不为Noto Sans CJK SC、10.5磅、颜色5B5B5B",
          ",".join(sorted(set(det))[:5]) or "卷头副标题格式与模板一致")

    # -3：卷头右栏各“：____________”填写行不为Noto Sans CJK SC、9.5磅、居中、非加粗
    fills = [p for p in cv_paras
             if re.search("[：:]" + _UNDERLINE_RE + "{4,}", _norm(_t(p)))]
    det = []
    if not fills:
        det.append("未找到填写行")
    for p in fills:
        if _jc(p) != "center":
            det.append("对齐%s" % _jc(p))
        det += _fmt_check(p, SANS, "eastAsia", 9.5, None, bold=False)
    _emit(-3, bool(det),
          "卷头右栏各“：____________”填写行不为Noto Sans CJK SC、9.5磅、居中、非加粗",
          "填写行=%d行 %s" % (len(fills), ",".join(sorted(set(det))[:5]) or "格式与模板一致"))

    # -1：卷头信息块行距不为单倍
    # 与办公软件对齐：单倍行距即未设 w:spacing/@w:line，或 line=240 且 lineRule=auto。
    sp = set()
    for p in cv_paras:
        line, rule = _line_spacing(p)
        if line is None:
            continue
        if not (rule in (None, "auto") and abs(line - 240) <= 2):
            sp.add("%s/%s" % (line, rule))
    _emit(-1, bool(sp),
          "卷头信息块行距不为单倍",
          "非单倍行距=%s" % (",".join(sorted(sp)[:4]) or "无"))

    # -1：卷头信息块表格不为水平居中，或两列列宽比例失调导致填写栏文字换行错位
    # 列宽比例判定：任一列占比低于 15% 即右栏（或左栏）过窄，填写栏文字会换行错位。
    det = []
    if _tbl_jc(cover) != "center":
        det.append("表格对齐=%s" % _tbl_jc(cover))
    g = _grid(cover)
    if len(g) != 2 or sum(g) <= 0:
        det.append("列数=%d" % len(g))
    else:
        ratio = [x / float(sum(g)) for x in g]
        if any(x < 0.15 for x in ratio):
            det.append("列宽比=%s" % ":".join("%.0f%%" % (x * 100) for x in ratio))
    _emit(-1, bool(det),
          "卷头信息块表格不为水平居中，或两列列宽比例失调导致填写栏文字换行错位",
          ",".join(det) or "水平居中且两列比例均衡")

    # -3：卷头信息块、提示块或题块表格被加上0.75磅及以上（sz≥6）的框线或单元格之间的网格线，无框色块版式被破坏
    # 判定范围只含这三类“无框色块”表格自身的格线，不含其内部用于排布词语/选项的嵌套表格；
    #   网格线按实际渲染判定（1行1列的块不会渲染内部线，只声明不算破坏）；
    #   版式基线自带的浅色描边（9FC5E8/E2EAF2/B7D4EF）属于原有版式，不计为新增框线。
    det = []
    for name, t in ([("卷头信息块", cover), ("提示块", tip)]
                    + [("%s题块" % NUMS[k], blocks[k]) for k in range(9)]):
        if t is None:
            continue
        for kind, val, sz, col in _rendered_borders(t):
            if not _has_line(val):
                continue
            if (col or "").upper() in BASE_BORDER_COLORS:
                continue
            if kind in ("insideH", "insideV"):
                det.append("%s:内网格线%s" % (name, col))
            elif sz is not None and sz >= 6:
                det.append("%s:框线sz%s/%s" % (name, sz, col))
    _emit(-3, bool(det),
          "卷头信息块、提示块或题块表格被加上0.75磅及以上（sz≥6）的框线或单元格之间的网格线，"
          "无框色块版式被破坏",
          ",".join(sorted(set(det))[:5]) or "三类色块表格均未被加粗框线或网格线")

    # -3：文档中插入硬分页符，题块被强制断页
    n_br = len([x for x in body.iter(qn("w:br"))
                if x.get(qn("w:type")) == "page" and not _in_fallback(x)])
    n_pbb = 0
    for p in _paras(body):
        pPr = p.find(qn("w:pPr"))
        if pPr is None:
            continue
        el = pPr.find(qn("w:pageBreakBefore"))
        if el is not None and el.get(qn("w:val")) not in ("0", "false", "off"):
            n_pbb += 1
    _emit(-3, (n_br + n_pbb) > 0,
          "文档中插入硬分页符，题块被强制断页",
          "硬分页符=%d处 段前分页=%d处" % (n_br, n_pbb))

    # -3：额外添加“题号／得分”评分表（一至九＋总分的2行11列表格），凭空增加非试题内容
    det = []
    for t in _tbls(body):
        s = _norm(_t(t))
        if "题号" in s and "得分" in s and len(_grid(t)) >= 8:
            det.append("%d列表格含题号/得分" % len(_grid(t)))
    _emit(-3, bool(det),
          "额外添加“题号／得分”评分表（一至九＋总分的2行11列表格），凭空增加非试题内容",
          ",".join(det[:3]) or "未出现题号/得分评分表")

    # -3：额外臆造“八、阅读直通车（答题区）。（20分）”等标题，凭空增加题号
    # 判定：以“汉字序号＋、”开头的标题段总数超过九个（同一题号被重复起标题），或出现“答题区”字样。
    n_title = len([s for s in flat_norm if re.match(r"^[一二三四五六七八九]、", s)])
    has_ans = "答题区" in NT
    _emit(-3, n_title > 9 or has_ans,
          "额外臆造“八、阅读直通车（答题区）。（20分）”等标题，凭空增加题号",
          "大题标题段=%d个 出现“答题区”=%s" % (n_title, has_ans))

    # -3：纸张尺寸不为A4（21.00cm×29.70cm），或页边距不为左1.30cm、右1.30cm、上1.00cm、下0.80cm
    # 与办公软件对齐：A4 = 11906×16838 twips；页边距容差 ±10 twips（约 0.02cm，四舍五入误差）。
    sects = [s for s in _sect_prs(body) if s.find(qn("w:pgSz")) is not None]
    det = []
    for i, s in enumerate(sects):
        w, h = _pg_size(s)
        if w != PG_W or h != PG_H:
            det.append("节%d纸张=%sx%s" % (i + 1, w, h))
        mar = _pg_margins(s)
        for k, v in PG_MAR.items():
            if mar.get(k) is None or abs(mar[k] - v) > 10:
                det.append("节%d %s=%s" % (i + 1, k, mar.get(k)))
    if not sects:
        det.append("未定义页面设置")
    _emit(-3, bool(det),
          "纸张尺寸不为A4（21.00cm×29.70cm），或页边距不为左1.30cm、右1.30cm、"
          "上1.00cm、下0.80cm",
          ",".join(sorted(set(det))[:5]) or "A4与四边页边距均与模板一致")

    # -1：页面四周不保留single 2磅（sz=16）、颜色DDEAF7、距页边18磅的页面边框
    det = []
    for i, s in enumerate(sects):
        bd = _pg_borders(s)
        for side in ("top", "left", "bottom", "right"):
            item = bd.get(side)
            if item is None:
                det.append("节%d缺%s边框" % (i + 1, side))
                continue
            val, sz, col, space = item
            if val != PG_BORDER[0] or sz != PG_BORDER[1] \
                    or col != PG_BORDER[2] or space != PG_BORDER[3]:
                det.append("节%d %s=%s/sz%s/%s/space%s" % (i + 1, side, val, sz, col, space))
    if not sects:
        det.append("未定义页面设置")
    _emit(-1, bool(det),
          "页面四周不保留single 2磅（sz=16）、颜色DDEAF7、距页边18磅的页面边框",
          ",".join(sorted(set(det))[:5]) or "四周页面边框与模板一致")

    # -1：正文默认字体不为Noto Serif CJK SC、10.5磅
    # 与办公软件对齐：不传 run/段落，直接沿 docDefaults → 默认段落样式解析文档默认值。
    d_ea = _eff_font(doc, None, None, "eastAsia")
    d_sz = _eff_size(doc, None, None)
    bad_def = SERIF not in (d_ea or "") or d_sz is None or abs(d_sz - 10.5) > 0.25
    _emit(-1, bad_def,
          "正文默认字体不为Noto Serif CJK SC、10.5磅",
          "默认中文字体=%s 默认字号=%s" % (d_ea, d_sz))

    # ========== 反解为结构化 dim2_items ==========
    # results 里每项形如“+N：说明”/“-N：说明”，与 _MAX_DELTAS_ORDER 一一对应；
    # 加分项命中记 +max_delta，减分项命中记 max_delta（负值），未命中一律记 0。
    # 每项只输出 rule / max_delta / delta / hit 四个键：判定过程中收集的实测值
    # 只用于内部比对，判定说明由 rule 原文承载，不随结果返回。
    if len(results) != len(_MAX_DELTAS_ORDER):
        raise RuntimeError(
            f"dim2 项数不一致：results={len(results)} vs order={len(_MAX_DELTAS_ORDER)}")
    dim2_items = []
    total_score = 0
    for raw, max_delta in zip(results, _MAX_DELTAS_ORDER):
        head, _, rule = raw.partition("：")
        try:
            delta = int(head.replace("+", "").strip())
        except ValueError:
            delta = 0
        hit = delta != 0
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
    # 本地调试入口：默认评估脚本所在目录，也可通过命令行传入目标目录。
    if len(sys.argv) >= 2:
        _dir = sys.argv[1]
    else:
        _dir = os.path.dirname(os.path.abspath(__file__))
    _out = evaluate(_dir)
    print(json.dumps(_out, ensure_ascii=False, indent=2))
