#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# pyright: reportUnknownMemberType=false, reportUnknownVariableType=false
# pyright: reportUnknownArgumentType=false, reportMissingImports=false
# pyright: reportOptionalMemberAccess=false, reportGeneralTypeIssues=false
"""028《职称评审表》表格补边框 / 上下段对齐 / 补四个标题 —— 完成度评分脚本。"""
import os
import re
import sys
import json
import hashlib
import zipfile
import traceback

from docx import Document
from docx.oxml.ns import qn

# 说明：按接口统一约定，本模块不修改全局 sys.stdout，也不 print 主结果。
# 主结果统一由 evaluate(dir_path) -> dict 返回；入参为“脚本所在目录的路径”，
# 由脚本自己在该目录内定位并打开被评估的 .docx。本文件仅在 __main__ 下作为
# 本地调试入口，才把 dict 序列化为 JSON 打印。

SCRIPT_ID = "028"
# 维度二加分项 max_delta 之和；维度一未通过时也按该满分记 0 分，
# 仅 status='error'（未定位到文档 / 脚本异常）时 max_score 记 0
MAX_SCORE_TOTAL = 49
# 维度二减分项 max_delta 之和（负值），仅用于加载时自检
MIN_PENALTY_TOTAL = -32

# 维度二所有评分项按"结果追加顺序"的 max_delta 列表；末端把 results 里的
# "+N：..." / "-N：..." 字符串反解为结构化 dim2_items 时用。顺序须与
# _evaluate_impl 中 results.append(...) 的顺序严格一致；共 33 项，
# 正项累加 = MAX_SCORE_TOTAL，负项累加 = MIN_PENALTY_TOTAL。
_MAX_DELTAS_ORDER = [
    # ---- 任务一：文档中的所有表格补齐边框 ----
    5,   # 01 大表（国内外进修情况／指导研究生和专业技术人员进修情况）横线纵线全补齐
    5,   # 02 获现职称以来获奖情况表（20行）+ 缺失的底部外框线
    3,   # 03 发包承揽关系甲乙方项目表（18行）
    3,   # 04 获现职称之前主要专业技术工作表（12行）
    3,   # 05 获现职称以来撰写的主要论文、著作表（11行）
    3,   # 06 二、学术会议宣读的论文表（11行）
    1,   # 07 获现职称以来完成继续教育任务情况表
    1,   # 08 独立完成表及其跨页续表
    1,   # 09 多方(多人)合作完成表及其跨页续表 + 续表底部外框线
    3,   # 10 20个表格框线线型与磅值统一
    1,   # 11 不留 23 缇空白细条格
    # ---- 任务二：图片中的表格上下单元格字体对齐 ----
    5,   # 12 指导段 5 列列宽与进修段逐列一致、纵线贯通
    3,   # 13 三行各单元格 vAlign=center、前置空段落数一致
    1,   # 14 上下两段字体、字号、加粗与水平对齐方式一致
    # ---- 任务三：红圈处按顺序补 4 个标题 ----
    5,   # 15 空行内自左至右填入 4 个标题，一格一个
    3,   # 16 4 个标题落在其对应内容列的正上方
    1,   # 17 4 个标题水平居中且垂直居中
    1,   # 18 4 个标题字体、字号、加粗与同表表头一致
    1,   # 19 标题行行高由 23 缇改为可完整显示标题
    # ---- 减分项：指令未要求而擅自改动、破坏交付内容 ----
    -5,  # 20 身份证号码被改写
    -5,  # 21 主要工作简历 8 条经历的单位与证明人被改写
    -3,  # 22 工作单位名称被改写
    -3,  # 23 获奖情况 6 条记录的项目名称／授予部门被改写
    -3,  # 24 国内外进修情况三条记录的单位被改写
    -3,  # 25 独立完成表 2022.02至2022.11 行完成情况被清空／改写
    -3,  # 26 证件照被替换、重新编码或改变长宽比
    -1,  # 27 首表中未涉及项的文字被改写
    -1,  # 28 各表中机构名称被改写
    -1,  # 29 国内外进修情况竖排标签被破坏
    -1,  # 30 进修段最后一行四格内容被改动或列数被增删
    -1,  # 31 本就框线完整的 8 张表框线／列宽／合并结构被改
    -1,  # 32 加入修订痕迹、批注、额外文本框标注或强制保护
    -1,  # 33 表格之外的 127 个正文段落被增删改动
]
assert sum(d for d in _MAX_DELTAS_ORDER if d > 0) == MAX_SCORE_TOTAL
assert sum(d for d in _MAX_DELTAS_ORDER if d < 0) == MIN_PENALTY_TOTAL

# --------------------------------------------------------------------------
# 交付文档既有内容的基线（用于全部减分项的“是否被改动”判定）
# --------------------------------------------------------------------------
EXPECT_BODY_PARAS = 127
PHOTO_PART = "word/media/image1.jpeg"
PHOTO_SIZE = 5685
PHOTO_MD5 = "e8a56a075d775aaa0427812d5575cf0a"
EXPECT_TXBX = 4

# 表格下标（0 基）
T_MAIN, T_RESUME, T_STUDY, T_FOREIGN, T_EDU = 0, 1, 2, 3, 4
T_BEFORE, T_AWARD, T_SOLO, T_SOLO2 = 5, 6, 7, 8
T_COOP, T_COOP2, T_CONTRACT, T_PAPER, T_CONF = 9, 10, 11, 12, 13
# 输入文档中本就框线完整、结构不应被重排的 8 张表
T_INTACT = [0, 1, 14, 15, 16, 17, 18, 19]

# 8 张完整表的行数与每行单元格数基线
INTACT_SHAPE = {
    0: [11, 9, 5, 8, 4, 6, 6, 6, 6, 6, 6, 6, 6, 6],
    1: [5, 5, 5, 5, 5, 5, 5, 5, 5],
    14: [1, 1, 2, 2, 2],
    15: [1, 3, 3, 3, 1, 1],
    16: [1, 1, 2],
    17: [1, 2, 2, 2, 1],
    18: [1, 2, 9, 2, 5, 8],
    19: [1, 1, 2],
}

# 任务三要补的 4 个标题，以及它们各自应当对齐的内容列文字
TITLES = ["起始时间", "专业技术人员类别", "指导内容", "指导情况"]
TITLE_CONTENT_ROW = "2024.06至2024.06"
TITLE_CONTENTS = ["2024.06至2024.06", "医院专业技术人员",
                  "职称申报与职业发展规划",
                  "经辅导，职工晋升副高级4名、中级8名、初级18名；学历提升本科11名、专科6名。"]
GUIDE_ROW_KEYS = ["2023.10至2024.06", "2024.06至2024.06", "2025.05至2025.05"]
STUDY_ROW_MARK = "学时，合格"
VERT_LABEL = "国内外进修情况"

# 减分项基线文字
ID_NUMBER = "454045777852124578"
RESUME_UNITS = [
    "青禾市菟乌教育咨询有限公司综合办公室人事专员",
    "青禾市图荼人才服务有限公司招聘专员",
    "青禾市鸣齐企业管理咨询有限公司人力资源项目岗",
    "南岭人力资源服务有限公司珠海分公司薪酬专员",
    "青禾市远裄物流服务有限公司人力资源部绩效管理岗",
    "青禾市清湾城市服务有限公司人力资源部培训与员工关系岗",
    "青禾市公共事务服务中心历任组织人事岗、综合管理组负责人",
    "青禾市仁好医院人力资源科人事管理岗",
]
RESUME_WITNESS = ["陈娅锶", "赵顾钚", "周锻域", "罗静依", "许安珃", "贺明鱼", "林远夯"]
WORK_UNIT = "青禾市仁好医院"

AWARD_ITEMS = [
    ("青禾市青年人才培养计划", "青禾市人才发展服务中心"),
    ("青禾市组织效能改进案例评选", "青禾市企业管理促进中心"),
    ("青禾市公共服务数据治理专项", "青禾市公共事务管理办公室"),
    ("区域医疗机构行政管理人才培养计划", "青禾市卫生事业发展中心"),
    ("医务社会工作专业支持人才库", "青禾市卫生事业发展中心"),
    ("青禾市医院运营评估专家库成员", "青禾市医院管理研究中心"),
]
STUDY_UNITS = [
    "青禾市南粤职业发展学院",
    "青禾市人才服务园南粤人力资源发展协会",
    "青禾市人才会议中心第六届区域医院运营交流会南粤医院管理研究会",
]
SOLO_DONE_ROW = "2022.02至2022.11"
SOLO_DONE_VALUE = "完成"
MAIN_BIRTH_PLACE = "青禾市香洲区"
MAIN_DUTY = "人力资源主任2025-08-15"
SCHOOL_TEXTS = ["青禾市清皖中学", "青禾市社会服务中心", "青禾市卫生管理培训中心"]
ORG_TEXTS = ["青禾市仁好医院（研究生约24名）",
             "作为项目主管，修订青禾市远裄物流服务有限公司考勤、休假与加班管理制度。",
             "医院人力资源数据治理与薪酬核算协同机制研究",
             "铧曳公共服务研究中心", "公共治理实务编辑部"]
LAST_ROW_CELLS = ["专业实践能力考试成绩（属于考评结合专业方需填写）", "64",
                  "考试时间", "2024-07-06"]
MAIN_UNTOUCHED = [
    "梁景川", "1988.03", "汉族", "中共党员", "大学本科", "学士学位",
    "2012-7-02", "中级经济师", "2015-11-15", "岭南省专业技术资格评价中心",
    "2007.09至2011.06", "南粤财经大学", "劳动与社会保障",
]
BODY_PARA_TEXTS = [
    "申报评审（表二）", "广东省职称评审表", "梁景川",
    "工商管理专业中级经济师职称", "人力资源管理专业高级经济师职称",
    "2026年2月18日",
]
FORBIDDEN_MARKS = ["修正版", "已补齐", "边框已", "说明：本文档", "AI生成"]

# --------------------------------------------------------------------------
# 通用工具
# --------------------------------------------------------------------------
_WS_RE = re.compile(r"\s+")


def _norm(text):
    """去掉所有空白字符，便于与基线文字做稳定比对。"""
    return _WS_RE.sub("", text or "")


def _el_text(el):
    return _norm("".join(n.text or "" for n in el.iter(qn("w:t"))))


def _tables(doc):
    return doc.element.body.findall(qn("w:tbl"))


def _rows(tbl):
    return tbl.findall(qn("w:tr"))


def _cells(tr):
    return tr.findall(qn("w:tc"))


def _tcPr(tc):
    return tc.find(qn("w:tcPr"))


def _grid_span(tc):
    pr = _tcPr(tc)
    if pr is None:
        return 1
    gs = pr.find(qn("w:gridSpan"))
    if gs is None:
        return 1
    try:
        return max(1, int(gs.get(qn("w:val")) or "1"))
    except (TypeError, ValueError):
        return 1


def _is_vmerge_continue(tc):
    pr = _tcPr(tc)
    if pr is None:
        return False
    vm = pr.find(qn("w:vMerge"))
    if vm is None:
        return False
    return (vm.get(qn("w:val")) or "continue") == "continue"


def _grid(tbl):
    """把表格展开成网格：每行是 [(tc, 是否纵向合并延续), ...]，gridSpan 按列展开。"""
    grid = []
    for tr in _rows(tbl):
        line = []
        for tc in _cells(tr):
            cont = _is_vmerge_continue(tc)
            for _ in range(_grid_span(tc)):
                line.append((tc, cont))
        grid.append(line)
    return grid


def _real(val):
    """边框 val 是否是一条可见线（nil/none 视为无线）。"""
    return val is not None and val not in ("nil", "none")


def _border_map(parent, tag, names):
    out = {}
    if parent is None:
        return out
    holder = parent.find(qn(tag))
    if holder is None:
        return out
    for name in names:
        el = holder.find(qn("w:" + name))
        if el is not None:
            out[name] = (el.get(qn("w:val")), el.get(qn("w:sz")))
    return out


_TBL_SIDES = ("top", "left", "bottom", "right", "insideH", "insideV")
_TC_SIDES = ("top", "left", "bottom", "right")


def _tbl_borders(tbl):
    return _border_map(tbl.find(qn("w:tblPr")), "w:tblBorders", _TBL_SIDES)


def _tc_borders(tc):
    return _border_map(_tcPr(tc), "w:tcBorders", _TC_SIDES)


def _eff_side(tcb, tblb, side, fallback=None):
    """单元格某一侧的有效边框 val：单元格级优先，未声明时回落到表级。"""
    if side in tcb:
        return tcb[side][0]
    item = tblb.get(fallback or side)
    return item[0] if item else None


def _border_defects(tbl):
    """相邻两格之间只要一侧有可见线即算该线已补；vMerge 延续处不要求横线。

    返回 (缺失内部横线数, 缺失内部纵线数, 缺失外框线数)。
    """
    grid = _grid(tbl)
    tblb = _tbl_borders(tbl)
    cache = {}

    def borders(tc):
        key = id(tc)
        if key not in cache:
            cache[key] = _tc_borders(tc)
        return cache[key]

    miss_h = miss_v = miss_out = 0
    nrow = len(grid)
    for r in range(nrow):
        line = grid[r]
        ncol = len(line)
        for c in range(ncol):
            tc = line[c][0]
            bd = borders(tc)
            if r == 0 and not _real(_eff_side(bd, tblb, "top")):
                miss_out += 1
            if r == nrow - 1 and not _real(_eff_side(bd, tblb, "bottom")):
                miss_out += 1
            if c == 0 and not _real(_eff_side(bd, tblb, "left")):
                miss_out += 1
            if c == ncol - 1 and not _real(_eff_side(bd, tblb, "right")):
                miss_out += 1
            if r + 1 < nrow and c < len(grid[r + 1]):
                down, down_cont = grid[r + 1][c]
                if down is not tc and not down_cont:
                    up = _eff_side(bd, tblb, "bottom", "insideH")
                    dn = _eff_side(borders(down), tblb, "top", "insideH")
                    if not (_real(up) or _real(dn)):
                        miss_h += 1
            if c + 1 < ncol:
                right = line[c + 1][0]
                if right is not tc:
                    lf = _eff_side(bd, tblb, "right", "insideV")
                    rt = _eff_side(borders(right), tblb, "left", "insideV")
                    if not (_real(lf) or _real(rt)):
                        miss_v += 1
    return miss_h, miss_v, miss_out


def _borders_ok(tbl):
    return _border_defects(tbl) == (0, 0, 0)


def _bottom_edge_ok(tbl):
    """底部外框线是否闭合（获奖情况表、多方合作续表专门要求）。"""
    grid = _grid(tbl)
    if not grid:
        return False
    tblb = _tbl_borders(tbl)
    last = grid[-1]
    return all(_real(_eff_side(_tc_borders(tc), tblb, "bottom")) for tc, _ in last)


def _line_styles(tbl):
    """收集该表所有可见边框的 (val, sz)，用于线型/磅值统一性判定。"""
    vals, szs = set(), set()
    tblb = _tbl_borders(tbl)
    for val, sz in tblb.values():
        if _real(val):
            vals.add(val)
            if sz:
                szs.add(sz)
    for tr in _rows(tbl):
        for tc in _cells(tr):
            for val, sz in _tc_borders(tc).values():
                if _real(val):
                    vals.add(val)
                    if sz:
                        szs.add(sz)
    return vals, szs


def _row_height(tr):
    pr = tr.find(qn("w:trPr"))
    if pr is None:
        return None, None
    th = pr.find(qn("w:trHeight"))
    if th is None:
        return None, None
    try:
        val = int(th.get(qn("w:val")) or "0")
    except (TypeError, ValueError):
        val = None
    return val, th.get(qn("w:hRule"))


def _thin_blank_rows(tbl, limit=60):
    """高度不超过 limit 缇且不含文字的占位空行。"""
    hits = []
    for idx, tr in enumerate(_rows(tbl)):
        if _el_text(tr):
            continue
        val, _rule = _row_height(tr)
        if val is not None and val <= limit:
            hits.append(idx)
    return hits


def _grid_cols(tbl):
    """tblGrid 定义的各网格列宽（缇）；本表的列宽全部由它与 gridSpan 决定。"""
    holder = tbl.find(qn("w:tblGrid"))
    if holder is None:
        return []
    out = []
    for gc in holder.findall(qn("w:gridCol")):
        try:
            out.append(int(gc.get(qn("w:w")) or "0"))
        except (TypeError, ValueError):
            out.append(0)
    return out


def _row_bounds(tr, cols):
    """行内各单元格的累计右边界（缇），用于判定上下两段纵线是否贯通对齐。"""
    acc = 0
    gi = 0
    out = []
    for tc in _cells(tr):
        for _ in range(_grid_span(tc)):
            if gi < len(cols):
                acc += cols[gi]
            gi += 1
        out.append(acc)
    return out


def _tc_valign(tc):
    pr = _tcPr(tc)
    if pr is None:
        return None
    v = pr.find(qn("w:vAlign"))
    return v.get(qn("w:val")) if v is not None else None


def _paras(el):
    return el.findall(qn("w:p"))


def _para_jc(p):
    pr = p.find(qn("w:pPr"))
    if pr is None:
        return None
    j = pr.find(qn("w:jc"))
    return j.get(qn("w:val")) if j is not None else None


def _run_props(run):
    """返回 (eastAsia 字体, 显式字号, 是否加粗)。"""
    pr = run.find(qn("w:rPr"))
    if pr is None:
        return None, None, False
    rf = pr.find(qn("w:rFonts"))
    ea = rf.get(qn("w:eastAsia")) if rf is not None else None
    szel = pr.find(qn("w:sz"))
    sz = szel.get(qn("w:val")) if szel is not None else None
    b = pr.find(qn("w:b"))
    bold = b is not None and (b.get(qn("w:val")) or "1") in ("1", "true", "on")
    return ea, sz, bold


def _text_runs(el):
    """该元素下所有含非空白文字的 run。"""
    out = []
    for run in el.iter(qn("w:r")):
        if _norm("".join(n.text or "" for n in run.findall(qn("w:t")))):
            out.append(run)
    return out


def _study_layout(tbl):
    """定位“国内外进修情况／指导研究生和专业技术人员进修情况”合并大表的关键行。

    返回 (rows, 每行单元格文字, 进修段数据行号, 指导段数据行号, 标题行号)。
    """
    rows = _rows(tbl)
    texts = [[_el_text(tc) for tc in _cells(tr)] for tr in rows]
    study = [i for i, cs in enumerate(texts) if any(STUDY_ROW_MARK in c for c in cs)]
    guide = [i for i, cs in enumerate(texts)
             if any(k in "".join(cs) for k in GUIDE_ROW_KEYS)]
    title = (min(guide) - 1) if guide and min(guide) >= 1 else None
    return rows, texts, study, guide, title


def _leading_blank_paras(tc):
    """单元格内文字之前的空段落个数（靠空段落顶高会破坏垂直居中的一致性）。"""
    count = 0
    for p in _paras(tc):
        if _el_text(p):
            return count
        count += 1
    return count


# --------------------------------------------------------------------------
# 维度1：可用与可修改性
# --------------------------------------------------------------------------
def _check_dim1(file_path):
    try:
        with zipfile.ZipFile(file_path) as zf:
            names = zf.namelist()
    except Exception as exc:
        return False, "文件无法作为 .docx 包打开：%s" % exc
    if "word/document.xml" not in names:
        return False, "缺少 word/document.xml，不是可正常打开的 Word 文档"
    try:
        Document(file_path)
    except Exception as exc:
        return False, "文档内容损坏，无法正常打开：%s" % exc
    return True, ""


# --------------------------------------------------------------------------
# 对外接口
# --------------------------------------------------------------------------
class _InputNotFound(Exception):
    """目录不存在 / 目录中没有 .docx。

    这类"没有可评估的交付物"属预期内的输入问题，摘要文字原样写进 error，
    不加异常类名前缀。
    """


def _as_error(result, message):
    """把 result 置为 status='error' 形态。

    约定：摘要文字只写进 error；dim1_pass 置空字符串（不是 None、也不是 False），
    dim1_reason 留空，dim2_items 清空，total_score 与 max_score 均记 0。
    """
    result["status"] = "error"
    result["error"] = message
    result["dim1_pass"] = ""
    result["dim1_reason"] = ""
    result["dim2_items"] = []
    result["total_score"] = 0
    result["max_score"] = 0
    return result


def _open_word_document(file_path):
    return Document(file_path)


def evaluate(dir_path):
    """按 §2.2 约定，返回结构化 dict；不 print、不改 sys.stdout、不 sys.exit。
    入参 dir_path 为“脚本所在目录的路径”，脚本自己在该目录里定位并打开被评估的
    .docx（忽略以 ~$ 开头的 Word 临时锁文件）。脚本自身崩溃（含目录不存在、
    未发现文档等）由本函数捕获并返回 status='error'。"""
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
        # 在 dir_path 目录内定位被评估的 .docx（忽略 ~$ 打头的 Word 临时锁文件）
        if not dir_path or not os.path.isdir(dir_path):
            raise _InputNotFound(f"目录不存在：{dir_path}")
        cands = sorted(
            f for f in os.listdir(dir_path)
            if f.lower().endswith('.docx') and not f.startswith('~$')
        )
        if not cands:
            raise _InputNotFound(f"目录中未找到 .docx 文件：{dir_path}")
        file_path = os.path.join(dir_path, cands[0])
        result["file_name"] = os.path.basename(file_path)
        return _evaluate_impl(file_path, result)
    except _InputNotFound as exc:
        # 未定位到交付物：错误摘要原样输出，不带异常类名
        return _as_error(result, str(exc))
    except Exception as exc:
        # 脚本自身崩溃：保留异常类名便于定位
        return _as_error(result, f"{type(exc).__name__}: {exc}")


def _evaluate_impl(file_path, result):
    """维度一 + 维度二的评估主体；对内部构造 results 列表沿用原逻辑，
    末尾按 _MAX_DELTAS_ORDER 反解成结构化 dim2_items 填入 result。"""
    ok1, reason = _check_dim1(file_path)
    result["dim1_pass"] = ok1
    # 约定：通过时 dim1_reason 留空，仅未通过时给出摘要文字
    result["dim1_reason"] = "" if ok1 else reason
    if not ok1:
        # 维度一未通过：完成度总分直接记 0（细则维度1的硬门槛）
        result["dim2_items"] = []
        result["total_score"] = 0
        return result

    doc = _open_word_document(file_path)
    tbls = _tables(doc)
    with zipfile.ZipFile(file_path) as zf:
        names = zf.namelist()
        photo = zf.read(PHOTO_PART) if PHOTO_PART in names else b""
        doc_xml = zf.read("word/document.xml").decode("utf-8", "ignore")
        settings = (zf.read("word/settings.xml").decode("utf-8", "ignore")
                    if "word/settings.xml" in names else "")

    all_text = _el_text(doc.element.body)
    tbl_text = [_el_text(t) for t in tbls]
    body_para_text = _norm("".join(p.text for p in doc.paragraphs))
    results = []

    # ---------------- 任务一：文档中的所有表格补齐边框 ----------------
    ok = _borders_ok(tbls[T_STUDY])
    results.append(("+5：" if ok else "+0：") + "表格“国内外进修情况／指导研究生和专业技术人员进修情况”（合并大表）的内部横线、纵线全部补齐，四周外框线闭合，上下两段每个单元格均有可见框线，无一处缺口")

    ok = _borders_ok(tbls[T_AWARD]) and _bottom_edge_ok(tbls[T_AWARD])
    results.append(("+5：" if ok else "+0：") + "表格“获现职称以来获奖情况”（20行）内部横线、纵线全部补齐，并补齐该表缺失的底部外框线，全表框线闭合无缺口")

    ok = _borders_ok(tbls[T_CONTRACT])
    results.append(("+3：" if ok else "+0：") + "表格“获现职称以来完成发包承揽关系甲乙方项目的专业技术工作及取得的业绩成果”（18行）内部横线全部补齐，四周外框线闭合")

    ok = _borders_ok(tbls[T_BEFORE])
    results.append(("+3：" if ok else "+0：") + "表格“获现职称之前主要专业技术工作及取得的业绩成果情况”（12行）内部横线与纵线全部补齐，四周外框线闭合")

    ok = _borders_ok(tbls[T_PAPER])
    results.append(("+3：" if ok else "+0：") + "表格“获现职称以来撰写的主要论文、著作”（11行）内部横线全部补齐，四周外框线闭合")

    ok = _borders_ok(tbls[T_CONF])
    results.append(("+3：" if ok else "+0：") + "表格“二、学 术 会 议 宣 读 的 论 文”（11行）内部横线全部补齐，四周外框线闭合")

    ok = _borders_ok(tbls[T_EDU])
    results.append(("+1：" if ok else "+0：") + "表格“获现职称以来完成继续教育任务情况”内部横线补齐，四周外框线闭合")

    ok = _borders_ok(tbls[T_SOLO]) and _borders_ok(tbls[T_SOLO2])
    results.append(("+1：" if ok else "+0：") + "表格“获现职称以来独立完成的专业技术工作及取得的业绩成果情况”及其跨页续表（含“2022.02至2022.11”行的6行表与随后的5行表）内部横线补齐，四周外框线闭合")

    ok = (_borders_ok(tbls[T_COOP]) and _borders_ok(tbls[T_COOP2])
          and _bottom_edge_ok(tbls[T_COOP2]))
    results.append(("+1：" if ok else "+0：") + "表格“获现职称以来多方(多人)合作完成的专业技术工作及取得的业绩成果”及其跨页续表（6行表与随后的5行表）内部横线补齐，并补齐续表缺失的底部外框线")

    vals, szs = set(), set()
    for tbl in tbls:
        v, s = _line_styles(tbl)
        vals |= v
        szs |= s
    # 只有 20 个表的外框都已闭合，才谈得上“线型与磅值统一”
    ok = (vals == {"single"} and len(szs) <= 1
          and all(_border_defects(tbl)[2] == 0 for tbl in tbls))
    results.append(("+3：" if ok else "+0：") + "20个表格的框线线型与磅值统一（同为单实线、同一磅值，如0.5磅/sz=4或0.75磅/sz=6），不出现某张表框线明显粗于或细于其他表、也不出现虚线、双线、点线混用")

    ok = not any(_thin_blank_rows(tbl) for tbl in tbls)
    results.append(("+1：" if ok else "+0：") + "补齐边框后表格中不出现高度仅23缇（约0.04cm）的空白细条格——此类占位空行须删除或撑高至可容纳一行文字，不得补线后变成一道贴在数据行上的多余细格")

    # ---------------- 任务二：图片中的表格上下单元格字体对齐 ----------------
    rows3, texts3, study_rows, guide_rows, title_idx = _study_layout(tbls[T_STUDY])
    title_cells = _cells(rows3[title_idx]) if title_idx is not None else []

    cols3 = _grid_cols(tbls[T_STUDY])
    five = [_row_bounds(tr, cols3) for tr in rows3 if len(_cells(tr)) == 5]
    ok = len(five) >= 6 and all(all(abs(a - b) <= 20 for a, b in zip(bd, five[0]))
                                for bd in five)
    results.append(("+5：" if ok else "+0：") + "“指导研究生和专业技术人员进修情况”区段5列的列宽与上方“国内外进修情况”区段5列逐列一致（如统一为821/1339/1800/4501/1259缇），上下两段纵向框线一一贯通对齐，不再是1229/931/1800/4500/1260与1229/1840/1567/2879/2205两套互相错位的列宽")

    ok = bool(guide_rows)
    for ri in guide_rows:
        content = [tc for tc in _cells(rows3[ri]) if not _is_vmerge_continue(tc)][-4:]
        if len(content) < 4 or any(_tc_valign(tc) != "center" for tc in content):
            ok = False
            break
        if len({_leading_blank_paras(tc) for tc in content}) != 1:
            ok = False
            break
    results.append(("+3：" if ok else "+0：") + "“2023.10至2024.06”“2024.06至2024.06”“2025.05至2025.05”三行各单元格的文字垂直居中（vAlign=center），与上方“起止时间”“国别、省别、单位”“学 习 内 容”“学 时”各格的 vAlign=center 一致；同一行4格的前置空段落数一致，不出现某格靠多余空段落把文字顶高，同一行文字与左右邻格文字落在同一水平线上")

    ok = bool(study_rows) and bool(guide_rows)
    szset = set()
    for ri in study_rows + guide_rows:
        for tc in _cells(rows3[ri])[-4:]:
            for p in _paras(tc):
                if _el_text(p) and _para_jc(p) != "center":
                    ok = False
            for run in _text_runs(tc):
                ea, sz, bold = _run_props(run)
                if ea not in (None, "宋体") or bold:
                    ok = False
                if sz:
                    szset.add(sz)
    ok = ok and len(szset) <= 1
    results.append(("+1：" if ok else "+0：") + "上下两段正文文字的字体、字号、加粗与水平对齐方式一致——同为宋体、同一字号、均不加粗、均水平居中，“青禾市人才服务园  南粤人力资源发展协会”“青禾市人才会议中心 第六届区域医院运营交流会  南粤医院管理研究会”不为两端对齐，“经辅导，职工晋升正高级1名、副高级3名、中级10名、初级19名。”不单独加粗")

    # ---------------- 任务三：红圈处按顺序补 4 个标题 ----------------
    tvals = [_el_text(tc) for tc in title_cells]
    title_ok = len(tvals) == 5 and tvals[1:5] == TITLES
    results.append(("+5：" if title_ok else "+0：") + "“指导研究生和专业技术人员进修情况”标签右侧原本空白的那一整行，自左至右依次填入“起始时间”“专业技术人员类别”“指导内容”“指导情况”，一个单元格一个标题，无遗漏、无顺序颠倒、无一格塞两个标题、无标题写到区段外的行")

    cont_idx = None
    for ri in guide_rows:
        if TITLE_CONTENT_ROW in "".join(texts3[ri]):
            cont_idx = ri
            break
    ok = False
    if title_ok and cont_idx is not None:
        cvals = [_el_text(tc) for tc in _cells(rows3[cont_idx])]
        if len(cvals) == 5 and all(TITLE_CONTENTS[k] in cvals[k + 1] for k in range(4)):
            tb = _row_bounds(rows3[title_idx], cols3)
            cb = _row_bounds(rows3[cont_idx], cols3)
            ok = (len(tb) == len(cb) == 5
                  and all(abs(a - b) <= 20 for a, b in zip(tb, cb)))
    results.append(("+3：" if ok else "+0：") + "4个标题分别落在其对应内容列的正上方——“起始时间”对应“2023.10至2024.06”，“专业技术人员类别”对应“医院专业技术人员”，“指导内容”对应“职称申报与职业发展规划”，“指导情况”对应“经辅导，职工晋升副高级4名、中级8名、初级18名；学历提升本科11名、专科6名。”")

    ok = title_ok and all(
        _tc_valign(tc) == "center"
        and all(_para_jc(p) == "center" for p in _paras(tc) if _el_text(p))
        for tc in title_cells[1:5])
    results.append(("+1：" if ok else "+0：") + "4个标题在各自单元格内水平居中且垂直居中")

    head_props = {_run_props(r) for tc in _cells(rows3[0])[-4:] for r in _text_runs(tc)}
    title_props = ({_run_props(r) for tc in title_cells[1:5] for r in _text_runs(tc)}
                   if title_ok else set())
    ok = title_ok and bool(title_props) and title_props == head_props
    results.append(("+1：" if ok else "+0：") + "4个标题的字体、字号、加粗与同表“起止时间”“国别、省别、单位”“学 习 内 容”“学 时”表头一致")

    hval, hrule = _row_height(rows3[title_idx]) if title_idx is not None else (None, None)
    ok = title_ok and (hval is None or hval >= 340 or hrule == "atLeast")
    results.append(("+1：" if ok else "+0：") + "该标题行行高由23缇（约0.04cm）的细条改为可完整显示标题的行高（≥340缇/0.6cm，或设为“最小值”自动撑高），4个标题文字未被压扁、截断或挤到框线外")

    # ---------------- 减分项：指令未要求而擅自改动、破坏交付内容 ----------------
    bad = ID_NUMBER not in all_text
    results.append(("-5：" if bad else "-0：") + "文档第一页“身  份  证  号  码”后的号码不为“454045777852124578”")

    t = tbl_text[T_RESUME]
    bad = (any(u not in t for u in RESUME_UNITS)
           or any(w not in t for w in RESUME_WITNESS))
    results.append(("-5：" if bad else "-0：") + "“主要工作简历”表8条经历的“在何地、何单位从事何工作”与“证明人”被改写——单位不为“青禾市菟乌教育咨询有限公司综合办公室人事专员”“青禾市图荼人才服务有限公司招聘专员”“青禾市鸣齐企业管理咨询有限公司人力资源项目岗”“南岭人力资源服务有限公司珠海分公司薪酬专员”“青禾市远裄物流服务有限公司人力资源部绩效管理岗”“青禾市清湾城市服务有限公司人力资源部培训与员工关系岗”“青禾市公共事务服务中心历任组织人事岗、综合管理组负责人”“青禾市仁好医院人力资源科人事管理岗”，或证明人不为“陈娅锶”“赵顾钚”“赵顾钚”“周锻域”“罗静依”“许安珃”“贺明鱼”“林远夯”")

    bad = (("现工作单位" + WORK_UNIT) not in tbl_text[T_MAIN]
           or all_text.count(WORK_UNIT) < 2)
    results.append(("-3：" if bad else "-0：") + "工作单位名称被改写——文档第一页“工   作   单   位”后的名称或首表“现工作单位”对应单元格内容不为“青禾市仁好医院”")

    t = tbl_text[T_AWARD]
    bad = any(a not in t or b not in t for a, b in AWARD_ITEMS)
    results.append(("-3：" if bad else "-0：") + "“获现职称以来获奖情况”表6条记录的“项目名称／授予部门”不为“青禾市青年人才培养计划／青禾市人才发展服务中心”“青禾市组织效能改进案例评选／青禾市企业管理促进中心”“青禾市公共服务数据治理专项／青禾市公共事务管理办公室”“区域医疗机构行政管理人才培养计划／青禾市卫生事业发展中心”“医务社会工作专业支持人才库／青禾市卫生事业发展中心”“青禾市医院运营评估专家库成员／青禾市医院管理研究中心”")

    t = tbl_text[T_STUDY]
    bad = any(u not in t for u in STUDY_UNITS)
    results.append(("-3：" if bad else "-0：") + "“国内外进修情况”区段三条进修记录的单位不为“青禾市南粤职业发展学院”“青禾市人才服务园  南粤人力资源发展协会”“青禾市人才会议中心 第六届区域医院运营交流会  南粤医院管理研究会”")

    bad = True
    for tr in _rows(tbls[T_SOLO]):
        cs = [_el_text(tc) for tc in _cells(tr)]
        if any(SOLO_DONE_ROW in c for c in cs):
            bad = not (len(cs) >= 4 and cs[-2] == SOLO_DONE_VALUE)
            break
    results.append(("-3：" if bad else "-0：") + "“获现职称以来独立完成的专业技术工作及取得的业绩成果情况”表中“2022.02至2022.11”行“完成情况”列内容不为“完成”（已有数据被清空、被删除同样扣分）")

    bad = len(photo) != PHOTO_SIZE or hashlib.md5(photo).hexdigest() != PHOTO_MD5
    results.append(("-3：" if bad else "-0：") + "首表第一行最右侧单元格中的证件照不为 word/media/image1.jpeg（5685字节）这张图片本身——被换成另一张图、被重新编码压缩（字节数改变）或被拉伸改变长宽比")

    t = tbl_text[T_MAIN]
    bad = (MAIN_BIRTH_PLACE not in t or MAIN_DUTY not in t
           or any(s not in all_text for s in SCHOOL_TEXTS)
           or any(s not in t for s in MAIN_UNTOUCHED))
    results.append(("-1：" if bad else "-0：") + "首表中未涉及项的文字被改写——“出生地”不为“青禾市香洲区”、“现行政职务及任职时间”不为“人力资源主任 2025-08-15”，“2004.09至2007.06”行学校名称不为“青禾市清皖中学”、“2019.03至2019.06”与“2021.05至2021.08”行发证单位不为“青禾市社会服务中心”“青禾市卫生管理培训中心”，或“姓 名／梁景川”“性别／男”“出生年月／1988.03”“民族／汉族”“政治面貌／中共党员”“最高学历／大学本科”“最高学位／学士学位”“参加工作时间／2012-7-02”“现职称／中级经济师”“取得时间／2015-11-15”“现职称取得方式／考试”“现职称发证单位／岭南省专业技术资格评价中心”“2007.09至2011.06／南粤财经大学／劳动与社会保障”等文字发生改写")

    bad = any(s not in all_text for s in ORG_TEXTS)
    results.append(("-1：" if bad else "-0：") + "各表中下列机构名称被改写——“2023.10至2024.06”行“专业技术人员类别”不为“青禾市仁好医院（研究生约24名）”，“2015.05至2015.06”行工作内容不为“作为项目主管，修订青禾市远裄物流服务有限公司考勤、休假与加班管理制度。”，“医院人力资源数据治理与薪酬核算协同机制研究”行出版／编辑单位不为“铧曳公共服务研究中心／公共治理实务编辑部”")

    label = _cells(rows3[0])[0]
    lps = [p for p in _paras(label) if _el_text(p)]
    bad = (len(lps) != 7
           or "".join(_el_text(p) for p in lps) != VERT_LABEL
           or any(_para_jc(p) != "center" for p in lps)
           or any(_run_props(r)[0] not in (None, "宋体") or not _run_props(r)[2]
                  for p in lps for r in _text_runs(p)))
    results.append(("-1：" if bad else "-0：") + "“国内外进修情况”竖排标签不为“国内外进修情况”7字逐字竖排、宋体加粗、居中显示")

    bad = [_el_text(tc) for tc in _cells(rows3[-1])] != LAST_ROW_CELLS
    results.append(("-1：" if bad else "-0：") + "“国内外进修情况”区段最后一行“专业实践能力考试成绩（属于考评结合专业方需填写）”“64”“考试时间”“2024-07-06”四格的文字内容被改动或该行列数被增删")

    bad = False
    for i in T_INTACT:
        if [len(_cells(tr)) for tr in _rows(tbls[i])] != INTACT_SHAPE[i]:
            bad = True
        if not _borders_ok(tbls[i]):
            bad = True
        if _line_styles(tbls[i])[0] - {"single"}:
            bad = True
    results.append(("-1：" if bad else "-0：") + "本就框线完整的首表、“主要工作简历”表、“本人承诺”表、“单  位  考  核  及  综  合  评  价”表、“评前公示情况”表、“委  托  评  审  审  核”表、“评委会办公室审核意见”表、“评审结果公示”表的框线被改成虚线、双线或被去除，或其列宽、合并结构被重排")

    bad = (bool(re.search(r"<w:ins[ />]", doc_xml))
           or bool(re.search(r"<w:del[ />]", doc_xml))
           or "word/comments.xml" in names
           or len(re.findall(r"<w:txbxContent", doc_xml)) > EXPECT_TXBX
           or bool(re.search(r'<w:documentProtection[^>]*w:enforcement="1"', settings))
           or any(m in all_text for m in FORBIDDEN_MARKS))
    results.append(("-1：" if bad else "-0：") + "交付文档被加入指令未要求的修订痕迹（w:ins/w:del）、批注、额外文本框标注或“修正版”“已补齐”等提示语，或被设为强制保护（w:documentProtection 的 w:enforcement 由0变为1）导致需先解除保护才能继续编辑")

    bad = (len(doc.paragraphs) != EXPECT_BODY_PARAS
           or any(s not in body_para_text for s in BODY_PARA_TEXTS))
    results.append(("-1：" if bad else "-0：") + "表格之外的127个正文段落被增删改动，“申报评审（表二）”“广东省职称评审表”“姓             名                            梁景川”“现      职     称    工商管理专业中级经济师职称”“申   报   职   称        人力资源管理专业高级经济师职称”“填   表   时   间            2026年2月18日”等文字内容、数量或先后顺序发生变化")


    if len(results) != len(_MAX_DELTAS_ORDER):
        raise RuntimeError("dim2 项数不一致：results=%d vs order=%d"
                           % (len(results), len(_MAX_DELTAS_ORDER)))
    dim2_items = []
    total_score = 0
    for raw, max_delta in zip(results, _MAX_DELTAS_ORDER):
        # 格式约定：加分项 '+N：' 前缀（命中 N=max_delta，未命中 N=0）；
        #           减分项 '-N：' 前缀（命中 N=|max_delta|，未命中 N=0）。
        head, _sep, rule = raw.partition("：")
        try:
            value = int(head.replace("+", "").strip())
        except ValueError:
            value = 0
        hit = value != 0
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
    # 本地调试入口：evaluate 只接收“脚本所在目录的路径”，由脚本自己在该目录里
    # 定位并打开被评估的 .docx。命令行第 1 个参数可覆盖目录路径；未指定时使用
    # 脚本自身所在目录。
    # 注：evaluate() 的接口约定就是“入参为目录路径”——路径始终由外部显式传入；
    #     这里的默认值仅为本地调试便利，不影响 evaluate 的接口约定。
    if len(sys.argv) >= 2:
        _dir = sys.argv[1]
    else:
        _dir = os.path.dirname(os.path.abspath(__file__))
    try:
        _out = evaluate(_dir)
        print(json.dumps(_out, ensure_ascii=False, indent=2))
    except Exception:
        traceback.print_exc()

