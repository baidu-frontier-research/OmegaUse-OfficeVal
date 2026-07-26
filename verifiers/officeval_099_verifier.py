# -*- coding: utf-8 -*-
"""
自动评估脚本：对“调查结果_公开版_第4题至第25题饼图.xlsx”按打分细则评估。

评估流程：
  维度1（可用与可修改性）：任一硬性条件不满足 -> 直接 0 分，不再检查维度2。
  维度2（完成度）：逐条加分/扣分细则累计。
      - 加分细则：必须满足该细则内的“每一个点”才加分。
      - 扣分细则：满足该细则内“任意一点”即扣分。
  最终打印命中情况与总分。

不依赖人工。对于难以从 xlsx 直接判定的项（如“颜色清晰可读 / 无艺术字 / 主要扇区可辨认”），
采用可量化代理指标灵活判定，并在输出中说明判定依据。
"""

import os
import sys
import json
import zipfile
import re
import math

import openpyxl

EMU_PER_CM = 360000.0

# 与脚本文件名 officeval_099_verifier.py 中的编号一致
SCRIPT_ID = "099"

# 22 道题（第4~第25题）的“选项—小计/比例”数据区域（细则给定）。
# 结构: 题号 -> (区域起始行, 区域结束行) ，列固定 A/B/C。
# 注意：区域含表头行(选项/小计/比例)在 A_top，选项数据行，及“本题有效填写人次”行。
QUESTION_RANGES = {
    4:  (27, 30),
    5:  (35, 39),
    6:  (44, 47),
    7:  (52, 55),
    8:  (60, 64),
    9:  (69, 72),
    10: (77, 81),
    11: (86, 89),
    12: (94, 98),
    13: (103, 106),
    14: (111, 114),
    15: (119, 122),
    16: (127, 130),
    17: (135, 138),
    18: (143, 146),
    19: (151, 154),
    20: (159, 162),
    21: (167, 171),
    22: (176, 179),
    23: (184, 188),
    24: (193, 197),
    25: (202, 206),
}


def cm(emu):
    if emu is None:
        return None
    return emu / EMU_PER_CM


def parse_ref(ref):
    """把 'sheet1'!$A$27:$A$30 解析为 (sheet, col_letter, row_start, row_end)。返回 None 失败。"""
    if not ref:
        return None
    m = re.match(r"^'?([^'!]+)'?!\$?([A-Za-z]+)\$?(\d+):\$?([A-Za-z]+)\$?(\d+)$", ref)
    if not m:
        # 单点引用
        m2 = re.match(r"^'?([^'!]+)'?!\$?([A-Za-z]+)\$?(\d+)$", ref)
        if m2:
            return (m2.group(1), m2.group(2).upper(), int(m2.group(3)),
                    m2.group(2).upper(), int(m2.group(3)))
        return None
    return (m.group(1), m.group(2).upper(), int(m.group(3)), m.group(4).upper(), int(m.group(5)))


def col_to_idx(letter):
    idx = 0
    for ch in letter:
        idx = idx * 26 + (ord(ch.upper()) - ord('A') + 1)
    return idx  # 1-based


def title_text(chart):
    """提取图表标题纯文本。"""
    t = chart.title
    if t is None:
        return None
    try:
        runs = []
        for p in t.tx.rich.p:
            for r in (p.r or []):
                if r.t:
                    runs.append(r.t)
        s = "".join(runs).strip()
        return s if s else None
    except Exception:
        return None


def load_chart_xmls(path):
    """读取 xlsx 内所有 chartN.xml 原始文本，返回 {chart_filename: xml_text}。"""
    xmls = {}
    with zipfile.ZipFile(path) as z:
        for name in z.namelist():
            if re.search(r"charts/chart\d+\.xml$", name):
                xmls[name] = z.read(name).decode("utf-8", errors="replace")
    return xmls


def chart_order_key(name):
    m = re.search(r"chart(\d+)\.xml$", name)
    return int(m.group(1)) if m else 0


# 字号在 OOXML 中以“百分之一磅”存储：sz="1000" = 10磅。
FONT_SZ_RE = re.compile(r'sz="(\d+)"')
TYPEFACE_RE = re.compile(r'typeface="([^"]+)"')
COMMON_FONTS = {"宋体", "微软雅黑", "calibri", "等线", "arial", "黑体", "simsun",
                "microsoft yahei", "dengxian", "song", "yahei"}
ART_FONT_HINTS = ("华文彩云", "华文琥珀", "华文新魏", "wordart", "impact")


def extract_label_flags(xml):
    """从图表 XML 中提取数据标签开关（chart级 + series级合并取“或”）。"""
    show = {
        "showPercent": 'showPercent val="1"' in xml,
        "showVal": 'showVal val="1"' in xml,
        "showCatName": 'showCatName val="1"' in xml,
        "showSerName": 'showSerName val="1"' in xml,
        "showLegendKey": 'showLegendKey val="1"' in xml,
    }
    return show


def extract_chart_type(xml):
    if "<c:pie3DChart" in xml:
        return "pie3d"
    if "doughnutChart" in xml:
        return "doughnut"
    if "<c:barChart" in xml:
        return "bar"
    if "<c:lineChart" in xml:
        return "line"
    if "<c:pieChart" in xml:
        return "pie2d"
    return "other"


def extract_font_sizes(xml):
    return [int(x) / 100.0 for x in FONT_SZ_RE.findall(xml)]


def extract_typefaces(xml):
    return [t.strip() for t in TYPEFACE_RE.findall(xml)]


def has_error_markers(xml):
    """检测 #REF! / #VALUE! 等错误标记。"""
    return ("#REF!" in xml) or ("#VALUE!" in xml) or ("#DIV/0!" in xml) or ("#N/A" in xml)


# ======================== 维度1 ========================

def evaluate_dim1(path):
    """返回 (passed: bool, details: list[(name, ok, msg)], wb, eval_path)。

    eval_path：实际用于维度2评估的文档路径，与传入的 path 一致
    （仅支持 .xlsx/.xlsm，无需转换）。
    """
    details = []

    # 1.1 交付文件格式为 .xlsx/.xlsm，且文件可正常打开
    fmt_ok = path.lower().endswith((".xlsx", ".xlsm"))
    wb = None
    eval_path = path
    open_ok = False
    err = ""
    if fmt_ok:
        try:
            wb = openpyxl.load_workbook(eval_path)
            open_ok = True
        except Exception as e:
            err = str(e)
    cond = fmt_ok and open_ok
    details.append(("交付文件格式为xlsx或.xlsm格式，文件可正常打开",
                    cond,
                    "格式OK，工作簿成功加载" if cond else "格式或打开失败: %s %s" % (
                        "" if fmt_ok else "扩展名非xlsx/xlsm;", err)))
    if not cond:
        return False, details, None, eval_path

    return True, details, wb, eval_path


# ======================== 维度2 ========================

def collect_chart_info(wb, path):
    """汇总每个图表的可用信息，按 chartN.xml 顺序对齐。"""
    ws = wb["sheet1"]
    charts = ws._charts
    xmls = load_chart_xmls(path)
    # openpyxl 的 ws._charts 顺序与 drawing 中 graphicFrame 顺序一致；
    # chart XML 文件按 chart1..chart22。这里以 XML 文件序号为主索引（与题号映射稳定）。
    xml_names = sorted(xmls.keys(), key=chart_order_key)

    infos = []
    for idx, ch in enumerate(charts):
        info = {}
        info["py"] = ch
        info["title"] = title_text(ch)
        # 数据源
        refs = []
        for s in ch.series:
            cat_f = None
            if s.cat is not None:
                if s.cat.numRef is not None:
                    cat_f = s.cat.numRef.f
                elif s.cat.strRef is not None:
                    cat_f = s.cat.strRef.f
            val_f = s.val.numRef.f if (s.val is not None and s.val.numRef is not None) else None
            refs.append((cat_f, val_f))
        info["refs"] = refs
        # 数值缓存（用于判定主要扇区是否可辨认）：取首个系列的数值列表
        values = []
        try:
            for s in ch.series:
                if s.val is not None and s.val.numRef is not None and s.val.numRef.numCache is not None:
                    for pt in s.val.numRef.numCache.pt:
                        try:
                            values.append(float(pt.v))
                        except Exception:
                            pass
                    if values:
                        break
        except Exception:
            pass
        info["values"] = values
        # anchor 位置/尺寸
        anc = ch.anchor
        frm = getattr(anc, "_from", None)
        info["from_col"] = frm.col if frm else None     # 0-based
        info["from_row"] = frm.row if frm else None      # 0-based
        # twoCellAnchor 有明确的 to（右下角）标记；oneCellAnchor 只有 from+ext，
        # 需用 ext(EMU) 折算出右下角列/行，用于判断图表纵向跨度与是否覆盖 A:C。
        to_marker = getattr(anc, "to", None)
        info["to_col"] = to_marker.col if to_marker else None   # 0-based，含分数偏移用 colOff 忽略，取整列号
        info["to_row"] = to_marker.row if to_marker else None
        ext = getattr(anc, "ext", None)
        info["w_cm"] = cm(ext.cx) if ext else None
        info["h_cm"] = cm(ext.cy) if ext else None
        info["ext_cx"] = ext.cx if ext else None
        info["ext_cy"] = ext.cy if ext else None
        infos.append(info)

    # 把 XML 信息按顺序附加（数量一致时一一对应）
    for i, info in enumerate(infos):
        if i < len(xml_names):
            xml = xmls[xml_names[i]]
            info["xml_name"] = xml_names[i]
            info["xml"] = xml
            info["type"] = extract_chart_type(xml)
            info["labels"] = extract_label_flags(xml)
            info["font_sizes"] = extract_font_sizes(xml)
            info["typefaces"] = extract_typefaces(xml)
            info["has_error"] = has_error_markers(xml)
            info["has_legend"] = "<c:legend>" in xml
        else:
            info["xml_name"] = None
            info["xml"] = ""
            info["type"] = "unknown"
            info["labels"] = {}
            info["font_sizes"] = []
            info["typefaces"] = []
            info["has_error"] = False
            info["has_legend"] = False
    return infos


def infer_question_for_chart(info):
    """根据数据源/标题推断图表对应的题号。优先用 val 引用所在行区间匹配 QUESTION_RANGES。"""
    # 1) 用数据源行区间匹配
    for cat_f, val_f in info["refs"]:
        for ref in (val_f, cat_f):
            p = parse_ref(ref)
            if not p:
                continue
            _, _, r0, _, r1 = p
            lo, hi = min(r0, r1), max(r0, r1)
            for q, (qs, qe) in QUESTION_RANGES.items():
                # 引用区间与该题选项区间有重叠即归属该题
                if not (hi < qs or lo > qe):
                    return q
    # 2) 退而用标题中的“第N题”
    if info["title"]:
        m = re.search(r"第\s*(\d+)\s*题", info["title"])
        if m:
            q = int(m.group(1))
            if q in QUESTION_RANGES:
                return q
    return None


def ref_within_question(ref, q):
    """判断引用区间是否恰好落在第 q 题的选项区域内（不含汇总行/表头之外的行）。"""
    p = parse_ref(ref)
    if not p:
        return False, "无法解析引用"
    _, col, r0, col2, r1 = p
    lo, hi = min(r0, r1), max(r0, r1)
    qs, qe = QUESTION_RANGES[q]
    # 数据源区域（细则给定的 A.. : C..）允许范围：qs..qe
    ok = (lo >= qs and hi <= qe)
    return ok, "引用%s 行%d-%d, 题区%d-%d" % (ref, lo, hi, qs, qe)


def evaluate_dim2(wb, path):
    """逐条评估维度2，返回 (total, hits: list[(rule, delta, ok, detail)])。"""
    infos = collect_chart_info(wb, path)
    n = len(infos)

    # 把每个图表映射到题号
    for info in infos:
        info["q"] = infer_question_for_chart(info)

    # 按题号建立映射（一题可能对应多个图，理想是 1:1）
    by_q = {}
    for info in infos:
        if info["q"] is not None:
            by_q.setdefault(info["q"], []).append(info)

    all_qs = sorted(QUESTION_RANGES.keys())  # 4..25
    hits = []
    total = 0

    # ---------- 加分项 ----------

    # +5 图表总数（严格对应细则的每一个点）：
    #   点1：从第4题到第25题共新增 22 个图表
    #   点2：每道题“右侧空白处”各对应 1 个图表（右侧 = 图表起始列在数据列 A:C 右侧）
    #   点3：每道题对应的是 1 个二维饼状图（pie2d）
    #   点4：不是三维饼图/环状图/条形图/柱形图/折线图或其他图表类型
    FORBIDDEN_TYPE_LABEL = {
        "pie3d": "三维饼图",
        "doughnut": "环状图",
        "bar": "条形图/柱形图",
        "line": "折线图",
        "other": "其他图表类型",
        "unknown": "未知类型",
    }

    # 点1：图表总数 = 22
    p1_count = (n == 22)

    # 点2：每道题右侧空白处各对应 1 个图表（每题恰好 1 个，且该图表位于 A:C 数据列右侧）
    per_q_one = {q: len(by_q.get(q, [])) for q in all_qs}
    p3_each_one = all(per_q_one[q] == 1 for q in all_qs)  # 每题恰 1 个（数量1:1）
    # 右侧空白处：图表左边缘列须在 A:C（0-based 列索引 0,1,2）右侧，即 from_col >= 3
    not_right_side = []
    for q in all_qs:
        for info in by_q.get(q, []):
            fc = info["from_col"]
            if fc is None or fc < 3:
                not_right_side.append((q, info.get("xml_name"), fc))
    p2_right_side = (len(not_right_side) == 0)

    # 点3 & 点4：每题对应图表为二维饼图，且不属于任何被禁止的类型
    forbidden_hits = []   # 命中禁止类型的图表
    not_pie2d = []        # 非二维饼图的图表
    for info in infos:
        t = info["type"]
        if t != "pie2d":
            not_pie2d.append((info.get("xml_name"), FORBIDDEN_TYPE_LABEL.get(t, t)))
        if t in FORBIDDEN_TYPE_LABEL:
            forbidden_hits.append((info.get("xml_name"), FORBIDDEN_TYPE_LABEL[t]))
    p4_all_pie2d = (len(not_pie2d) == 0 and n > 0)
    p5_no_forbidden = (len(forbidden_hits) == 0)

    # 该加分项需满足细则中的“每一个点”
    rule1_ok = p1_count and p2_right_side and p3_each_one and p4_all_pie2d and p5_no_forbidden
    missing_q = [q for q in all_qs if per_q_one[q] != 1]
    add_rule(
        hits, "图表总数：第4~25题共22个，每题右侧1个二维饼图(非3D/环状/条/柱/折线等)", 5, rule1_ok,
        "图表总数=%d(需22:%s)；每题恰1个=%s(异常题=%s)；右侧空白处=%s(不在A:C右侧=%s)；"
        "全为二维饼图=%s(非二维饼图=%s)；无禁止类型=%s(禁止类型命中=%s)" % (
            n, p1_count,
            p3_each_one, missing_q,
            p2_right_side, not_right_side,
            p4_all_pie2d, not_pie2d,
            p5_no_forbidden, forbidden_hits))
    if rule1_ok:
        total += 5

    # +3 图表数据源（严格对应细则的每一个点）：
    #   点1：共 22 个图表都要逐一判定（第4~第25题）
    #   点2：每个图表使用“各自题目”的数据区域（类别引用=本题A列选项，数值引用=本题B列“小计”或C列“比例”）
    #   点3：使用的是“选项—小计”或“选项—比例”数据区域（A+B 或 A+C 组合）
    #   点4：数据区域行范围恰为细则给定的各题区域，
    #        第4题A27:C30、第5题A35:C39、……、第25题A202:C206。
    src_ok_count = 0
    src_detail = []
    for q in all_qs:
        qs, qe = QUESTION_RANGES[q]
        lst = by_q.get(q, [])
        if not lst:
            src_detail.append("第%d题:无图" % q)
            continue
        info = lst[0]

        # 收集该图所有系列的 类别引用(cat) 与 数值引用(val)
        ok_this = False
        reason = []
        for cat_f, val_f in info["refs"]:
            pc = parse_ref(cat_f)
            pv = parse_ref(val_f)
            if not pc or not pv:
                reason.append("引用无法解析(cat=%s,val=%s)" % (cat_f, val_f))
                continue
            _, c_col, c_r0, c_col2, c_r1 = pc
            _, v_col, v_r0, v_col2, v_r1 = pv
            c_lo, c_hi = min(c_r0, c_r1), max(c_r0, c_r1)
            v_lo, v_hi = min(v_r0, v_r1), max(v_r0, v_r1)

            # 点2+点4：类别引用须为本题 A 列、且行范围恰为细则区域 [qs, qe]
            cat_ok = (c_col == "A" and c_col2 == "A" and c_lo == qs and c_hi == qe)
            # 点3+点4：数值引用须为本题 B列“小计”或 C列“比例”，行范围恰为 [qs, qe]
            val_is_xiaoji = (v_col == "B" and v_col2 == "B" and v_lo == qs and v_hi == qe)
            val_is_bili = (v_col == "C" and v_col2 == "C" and v_lo == qs and v_hi == qe)
            val_ok = val_is_xiaoji or val_is_bili

            if cat_ok and val_ok:
                ok_this = True
                kind = "选项—小计(A%d:A%d + B%d:B%d)" % (qs, qe, qs, qe) if val_is_xiaoji \
                    else "选项—比例(A%d:A%d + C%d:C%d)" % (qs, qe, qs, qe)
                reason.append("命中 %s" % kind)
                break
            else:
                reason.append("cat=%s(需A%d:A%d,%s) val=%s(需B或C %d:%d,%s)" % (
                    cat_f, qs, qe, "OK" if cat_ok else "NG",
                    val_f, qs, qe, "OK" if val_ok else "NG"))

        if ok_this:
            src_ok_count += 1
        else:
            src_detail.append("第%d题: %s" % (q, " / ".join(reason)))

    rule2_ok = (src_ok_count == 22)
    add_rule(
        hits,
        "图表数据源：22个图均用本题“选项—小计”或“选项—比例”且区域恰为细则给定(如第4题A27:C30)", 3, rule2_ok,
        "数据源正确题数=%d/22；%s" % (src_ok_count, "；".join(src_detail) if src_detail else "全部正确"))
    if rule2_ok:
        total += 3

    # +1 图表位置（严格对应细则的每一个点）：
    #   点1：22 个图表均位于“对应题目右侧空白区域”——不仅左边缘在 A:C 右侧，
    #        还要求图表纵向锚点（from_row..to_row，或 from_row+ext 折算）与该题
    #        的行区间 [qs-2, qe]（含题干行与选项区）有重叠，避免“列对但行不对”
    #        即紧邻其它题目右侧、被误判为合格的情况。
    #   点2：图表左边缘位于 D 列及其右侧（D列 0-based 列索引=3，故 from_col >= 3）
    #   点3：未覆盖 A:C 列原始数据——用图表的完整水平跨度（左边缘 from_col 到
    #        右边缘 to_col，twoCellAnchor 直接取 to；oneCellAnchor 用 from_col
    #        的列偏移(colOff) + 宽度(ext.cx) 折算出右边缘所在列）确认图表整体
    #        都在 A:C 右侧，而不是仅看左边缘（若图表向左跨越/宽度过大导致左侧
    #        实际盖住 C 列及以左，仍应判定为覆盖了原始数据）。
    EMU_PER_COL_DEFAULT = 640080.0  # Excel 默认列宽约 8.43 字符对应的 EMU，用于折算 oneCellAnchor 右边缘列

    def chart_col_span(info):
        """返回 (left_col, right_col)，均为 0-based 列号；right_col 为图表右边缘所跨的列号（估算，向上取整）。"""
        fc = info["from_col"]
        if fc is None:
            return None, None
        to_c = info.get("to_col")
        if to_c is not None:
            return fc, to_c
        # oneCellAnchor：无 to，用宽度(EMU)折算跨越的列数（保守估计，按默认列宽）
        cx = info.get("ext_cx")
        if cx is not None and EMU_PER_COL_DEFAULT > 0:
            span_cols = math.ceil(cx / EMU_PER_COL_DEFAULT)
            return fc, fc + span_cols
        return fc, fc

    pos_ok_count = 0
    pos_detail = []
    for q in all_qs:
        lst = by_q.get(q, [])
        if not lst:
            pos_detail.append("第%d题:无图" % q)
            continue
        info = lst[0]
        fc = info["from_col"]   # 0-based; A=0,B=1,C=2,D=3,E=4
        left_col, _right_col = chart_col_span(info)

        # 点2：左边缘在 D 列及右侧
        p_leftD = (fc is not None and fc >= 3)
        # 点3：图表整体（左边缘~右边缘）不覆盖 A:C（0-based 0/1/2），即左边缘本身
        #      已 >=3 即可保证整体不覆盖（图表只会向右延伸，不会向左覆盖 A:C）
        p_notcover = (left_col is not None and left_col >= 3)

        # 点1：图表纵向锚点须落在“对应题目右侧”——即纵向与该题所在行区间重叠。
        #      题目所在行区间放宽到题干行（qs-2）起，覆盖题干/表头/选项/汇总行，
        #      0-based 行号 = 1-based 行号 - 1。
        qs, qe = QUESTION_RANGES[q]
        q_row_lo = max(0, qs - 2 - 1)   # 题干行(1-based qs-2)的 0-based 行号，做下界保护
        q_row_hi = qe - 1               # 该题区域结束行(1-based)的 0-based 行号
        fr = info["from_row"]
        tr = info.get("to_row")
        if fr is None:
            p_vertical = False
        else:
            chart_row_lo = fr
            chart_row_hi = tr if tr is not None else fr
            p_vertical = not (chart_row_hi < q_row_lo or chart_row_lo > q_row_hi)

        if p_vertical and p_leftD and p_notcover:
            pos_ok_count += 1
        else:
            pos_detail.append(
                "第%d题(col0=%s row0=%s~%s 需在题区行%d-%d:%s；左边缘>=3(D列):%s；未覆盖A:C:%s)" % (
                    q, fc, fr, tr, q_row_lo, q_row_hi, p_vertical, p_leftD, p_notcover))

    rule3_ok = (pos_ok_count == 22)
    add_rule(
        hits,
        "图表位置：22个图均在对应题目右侧空白区,左边缘在D列及右侧,未覆盖A:C", 1, rule3_ok,
        "位置合格题数=%d/22；%s" % (pos_ok_count, "；".join(pos_detail) if pos_detail else "全部合格"))
    if rule3_ok:
        total += 1

    # +5 图表标签（严格对应细则的每一个点）：
    #   点1：对象为“22个二维饼状图”（须为 22 个，且均为二维饼图 pie2d）
    #   点2：均“显示主要占比信息”—— 即开启了数据标签显示（每个图都显示）
    #   点3：数据标签“包含百分比、数值或两者之一”—— showPercent 或 showVal 为真
    #   点4：主要扇区可辨认 —— 不能仅凭“存在正数”判定（几乎恒真），需综合：
    #        a) 最大扇区占比达到有意义的阈值（>=5%，太小的扇区即使标了数字视觉上
    #           也难以对应到具体扇区，参考饼图常见的最小可辨认扇区角度~18度）；
    #        b) 该图确有开启数据标签(点2/点3已校验，此处复用)；
    #        c) 图表尺寸不能过小——过小的图，扇区与标签会挤在一起难以辨认，
    #           用高度(h_cm)下限代理（<3cm 视为过小，明显小于细则"约5-10cm"尺寸要求）；
    #        d) 标签字体不能过小到不可读——图表标注字号存在但全部<5磅时，即使数字
    #           显示了也无法辨认，视为不可辨认。
    MIN_MAIN_SHARE = 0.05     # 最大扇区占比阈值：低于5%视为过小、难以在图上辨认
    MIN_CHART_H_CM = 3.0      # 图表过矮，标签/扇区拥挤难辨认的下限代理

    lbl_ok_count = 0
    lbl_detail = []
    considered = 0
    for q in all_qs:
        lst = by_q.get(q, [])
        if not lst:
            lbl_detail.append("第%d题:无图" % q)
            continue
        info = lst[0]
        considered += 1
        lab = info["labels"]

        # 点1：该图为二维饼图
        p_pie2d = (info["type"] == "pie2d")
        # 点3：数据标签包含百分比或数值（两者之一即可）
        p_pct_or_val = bool(lab.get("showPercent") or lab.get("showVal"))
        # 点2：显示主要占比信息（= 已开启上述数据标签显示）
        p_show = p_pct_or_val

        # 点4a：最大扇区占比达到有意义的阈值（不再是仅 >0）
        vals = info.get("values") or []
        nums = [v for v in vals if isinstance(v, (int, float))]
        total_v = sum(nums) if nums else 0
        max_share = (max(nums) / total_v) if total_v > 0 else 0
        p_share_significant = (total_v > 0 and max_share >= MIN_MAIN_SHARE)

        # 点4c：图表尺寸足够大，扇区与标签不至于拥挤难辨认
        h = info.get("h_cm")
        p_size_ok = (h is not None and h >= MIN_CHART_H_CM)

        # 点4d：标签/图表文字字号不能全部过小（<5磅视为不可读，与图表字体细则的
        #        5-20磅下限一致）；无显式字号则视为采用主题默认字号(可读)。
        sizes = info.get("font_sizes") or []
        p_font_readable = (len(sizes) == 0) or any(s >= 5.0 for s in sizes)

        # 点4：主要扇区可辨认 = 已显示标签 且 占比显著 且 图表不过小 且 字体可读
        p_main_sector = p_show and p_share_significant and p_size_ok and p_font_readable

        if p_pie2d and p_show and p_pct_or_val and p_main_sector:
            lbl_ok_count += 1
        else:
            lbl_detail.append(
                "第%d题(%s 二维饼:%s; 显示标签:%s; 含%%/数值:%s; 主扇区可辨认:%s"
                "[最大占比=%.2f%%(需>=%.0f%%):%s, 图高=%scm(需>=%.1fcm):%s, "
                "字体可读:%s])" % (
                    q, info.get("xml_name"), p_pie2d, p_show, p_pct_or_val,
                    p_main_sector, max_share * 100, MIN_MAIN_SHARE * 100,
                    p_share_significant,
                    None if h is None else round(h, 2), MIN_CHART_H_CM, p_size_ok,
                    p_font_readable))

    rule4_ok = (lbl_ok_count == 22 and considered == 22)
    add_rule(
        hits,
        "图表标签：22个二维饼图均显示占比信息,数据标签含百分比/数值之一,主要扇区可辨认", 5, rule4_ok,
        "合格图表=%d/22；%s" % (lbl_ok_count, "；".join(lbl_detail) if lbl_detail else "全部合格"))
    if rule4_ok:
        total += 5

    # +1 图表标题：均带标题且含题号或可识别题目
    # +1 图表标题（严格对应细则的每一个点）：
    #   点1：22 个图表均“带有标题”（标题存在且非空）
    #   点2：标题“包含对应题号‘第4题’至‘第25题’” 或
    #   点3：标题“能够明确识别对应题目内容”（含本题题干/选项关键词等可识别信息）
    #        —— 点2 与点3 为“或”关系，满足其一即可；但每个图都必须满足该“或”条件。
    title_ok_count = 0
    title_detail = []
    for q in all_qs:
        lst = by_q.get(q, [])
        if not lst:
            title_detail.append("第%d题:无图" % q)
            continue
        info = lst[0]
        t = info["title"]

        # 点1：带有标题（存在且非空）
        p_has_title = bool(t and t.strip())

        # 点2：标题包含对应题号“第q题”（容许“第 q 题”空格写法）
        p_has_qnum = False
        if p_has_title:
            if ("第%d题" % q) in t:
                p_has_qnum = True
            else:
                m = re.search(r"第\s*(\d+)\s*题", t)
                if m and int(m.group(1)) == q:
                    p_has_qnum = True

        # 点3：能够明确识别对应题目内容（标题含本题题干关键词或某个选项文本）
        p_identifiable = False
        if p_has_title:
            qs, qe = QUESTION_RANGES[q]
            # 题干：题标题行(选项表头 qs-1 的上一行 qs-2)的 A 列文本
            stem_cell = ws_get_text(wb, qs - 2)
            stem_core = ""
            if stem_cell:
                # 去掉“第N题：”前缀与“[单选题]”等后缀，取核心题干
                stem_core = re.sub(r"^第\d+题[：:、.\s]*", "", stem_cell)
                stem_core = re.sub(r"\[[^\]]*\]\s*$", "", stem_core).strip()
            # 选项文本
            option_texts = [ws_get_text(wb, r) for r in range(qs, qe + 1)]
            option_texts = [x for x in option_texts if x]
            # 标题命中题干核心(取前若干字)或任一选项
            if stem_core:
                probe = stem_core[:6]
                if probe and probe in t:
                    p_identifiable = True
            if not p_identifiable:
                for opt in option_texts:
                    if opt and len(opt) >= 2 and opt in t:
                        p_identifiable = True
                        break

        # 该图需：带标题 且（含题号 或 可识别题目内容）
        ok = p_has_title and (p_has_qnum or p_identifiable)
        if ok:
            title_ok_count += 1
        else:
            title_detail.append("第%d题(标题=%r 带标题:%s 含题号:%s 可识别内容:%s)" % (
                q, t, p_has_title, p_has_qnum, p_identifiable))

    rule5_ok = (title_ok_count == 22)
    add_rule(
        hits,
        "图表标题：22个图均带标题,且含对应题号(第4题~第25题)或能明确识别题目内容", 1, rule5_ok,
        "合格题数=%d/22；%s" % (title_ok_count, "；".join(title_detail) if title_detail else "全部合格"))
    if rule5_ok:
        total += 1

    # +1 图表字体（严格对应细则的每一个点）：
    #   点1：图表“标题、数据标签、图例文字”采用常规字体（宋体、微软雅黑、Calibri 等）
    #   点2：字号 5–20 磅（所有出现的显式字号须落在 5~20 磅）
    #   点3：颜色清晰可读（文字颜色非白/非极浅，未被设为不可读的浅色）
    #   点4：无艺术字（不使用艺术字字体，且未应用 WordArt 文本变形 prstTxWarp）
    font_bad = []
    for info in infos:
        xml = info.get("xml") or ""
        sizes = info["font_sizes"]              # 磅
        faces = [f.lower() for f in info["typefaces"]]
        problems = []

        # 点2：字号 5–20 磅（细则原文区间，不放宽）。无显式字号则视为采用主题默认(不违规)。
        bad_sizes = [s for s in sizes if not (5.0 <= s <= 20.0)]
        p_size = (len(bad_sizes) == 0)
        if not p_size:
            problems.append("字号越界%s磅(需5-20)" % bad_sizes)

        # 点1：常规字体。若声明了字体名，须均属于常规字体集合；未声明则用主题字体(视为常规)。
        #   +mn-*/+mj-* 为 OOXML 主题字体占位符(minor/major latin/ea/cs)，等价于采用主题
        #   默认字体(宋体/等线等常规字体)，应视为常规字体，不算非常规。
        non_regular = [f for f in faces
                       if not f.startswith("+mn-") and not f.startswith("+mj-")
                       and not any(cf in f for cf in COMMON_FONTS)]
        p_face = (len(non_regular) == 0)
        if not p_face:
            problems.append("非常规字体%s" % non_regular)

        # 点4：无艺术字（艺术字字体名 或 文本变形 prstTxWarp）
        art_font = [f for f in faces if any(h in f for h in ART_FONT_HINTS)]
        has_warp = ("prstTxWarp" in xml)
        p_noart = (len(art_font) == 0 and not has_warp)
        if not p_noart:
            problems.append("艺术字(字体%s,文本变形=%s)" % (art_font, has_warp))

        # 点3：颜色清晰可读。检测文字颜色是否被设为白色/极浅色(不可读)。
        #      未显式设色则采用主题默认深色(视为清晰可读)。
        text_colors = re.findall(r'srgbClr val="([0-9A-Fa-f]{6})"', xml)
        unreadable = []
        for c in text_colors:
            r = int(c[0:2], 16); g = int(c[2:4], 16); b = int(c[4:6], 16)
            # 近白/极浅(亮度过高)视为不可读
            if r >= 240 and g >= 240 and b >= 240:
                unreadable.append(c)
        p_color = (len(unreadable) == 0)
        if not p_color:
            problems.append("文字颜色过浅不可读%s" % unreadable)

        if not (p_size and p_face and p_noart and p_color):
            font_bad.append("%s: %s" % (info.get("xml_name"), "; ".join(problems)))

    rule6_ok = (len(font_bad) == 0 and n > 0)
    add_rule(
        hits,
        "图表字体：标题/数据标签/图例为常规字体(宋体/微软雅黑/Calibri等),字号5-20磅,颜色清晰可读,无艺术字", 1, rule6_ok,
        "异常图表=%s" % (font_bad if font_bad else "无（字体常规、字号5-20磅、颜色可读、无艺术字）"))
    if rule6_ok:
        total += 1

    # +5 图表图例（严格对应细则的每一个点）：
    #   点1：对象为 22 个图表（逐题判定，须全部满足）
    #   点2：每个图表“能通过图例 或 数据标签”识别——
    #        即存在图例(<c:legend>) 或 数据标签显示类别名(showCatName)。
    #   点3：能识别“各扇区对应的选项内容”——
    #        类别(选项)引用须指向本题选项区，且选项文本可取得(扇区与选项一一对应)。
    leg_ok_count = 0
    leg_detail = []
    for q in all_qs:
        lst = by_q.get(q, [])
        if not lst:
            leg_detail.append("第%d题:无图" % q)
            continue
        info = lst[0]
        lab = info["labels"]

        # 点2：能通过“图例”或“数据标签”识别扇区对应的选项——二者满足其一即可。
        #      图例(<c:legend>)本身就能把颜色块与选项名对应；数据标签显示类别名/
        #      占比/数值也能起到识别作用。之前版本把 p_channel 只绑定到数据标签
        #      (p_sector_label)，导致“有图例但未开数据标签”的图表被误判不合格，
        #      与细则“通过图例或数据标签”的“或”关系不符，属过严误判，现予纠正。
        p_legend = bool(info["has_legend"])
        p_catlabel = bool(lab.get("showCatName"))
        p_sector_label = bool(
            lab.get("showCatName") or lab.get("showPercent") or lab.get("showVal"))
        p_channel = p_legend or p_sector_label

        # 点3：可识别各扇区对应的“选项内容”——类别引用指向本题选项区且能取得选项文本
        qs, qe = QUESTION_RANGES[q]
        cat_ok = False
        cat_texts = []
        for cat_f, val_f in info["refs"]:
            pc = parse_ref(cat_f)
            if not pc:
                continue
            _, c_col, c_r0, c_col2, c_r1 = pc
            c_lo, c_hi = min(c_r0, c_r1), max(c_r0, c_r1)
            if c_col == "A" and c_col2 == "A" and c_lo >= qs and c_hi <= qe:
                # 取选项文本，确认非空（可辨认选项内容）
                cat_texts = [ws_get_text(wb, r) for r in range(c_lo, c_hi + 1)]
                cat_texts = [x for x in cat_texts if x]
                if len(cat_texts) >= 1:
                    cat_ok = True
                break
        p_options = cat_ok

        if p_channel and p_options:
            leg_ok_count += 1
        else:
            leg_detail.append("第%d题(%s 有图例:%s 或类别标签:%s; 可识别选项内容:%s,选项数=%d)" % (
                q, info.get("xml_name"), p_legend, p_catlabel, p_options, len(cat_texts)))

    rule7_ok = (leg_ok_count == 22)
    add_rule(
        hits,
        "图表图例：22个图均能通过图例或数据标签识别各扇区对应的选项内容", 5, rule7_ok,
        "合格题数=%d/22；%s" % (leg_ok_count, "；".join(leg_detail) if leg_detail else "全部合格"))
    if rule7_ok:
        total += 5

    # +3 图表尺寸（严格对应细则的每一个点）：
    #   点1：每个图表高度“约 5–10cm”（逐个判定；“约”给予小幅容差，取 5~10cm，
    #        并对边界放宽 ±0.5cm 体现“约”字，即 [4.5, 10.5]cm）。
    size_ok_count = 0
    size_detail = []
    for info in infos:
        h = info["h_cm"]

        # 点1：高度约 5–10cm
        p_height = (h is not None and 4.5 <= h <= 10.5)

        if p_height:
            size_ok_count += 1
        else:
            size_detail.append("%s(高=%scm 约5-10:%s)" % (
                info.get("xml_name"), None if h is None else round(h, 2), p_height))

    rule8_ok = (size_ok_count == 22 and n == 22)
    add_rule(
        hits,
        "图表尺寸：每个图高度约5-10cm", 3, rule8_ok,
        "合格图表=%d/22；%s" % (size_ok_count, "；".join(size_detail) if size_detail else "全部合格"))
    if rule8_ok:
        total += 3

    # ---------- 扣分项 ----------
    # 两条扣分规则(-3 数据源误纳入、-1 图表顺序)已按需求删除。

    return total, hits, infos


def add_rule(hits, rule, delta, ok_or_hit, detail):
    hits.append((rule, delta, ok_or_hit, detail))


# 评分细则原文（按 hits 追加顺序一一对应），命中时按原文展示。
SPEC_TEXTS = [
    # 加分项
    "“sheet1”图表总数：从第4题到第25题共新增22个图表，每道题右侧空白处各对应1个二维饼状图，不是三维饼图、环状图、条形图、柱形图、折线图或其他图表类型。",
    "“sheet1”图表数据源：22个图表均使用各自题目的“选项—小计”或“选项—比例”数据区域作为图表数据源，分别对应第4题A27:C30、第5题A35:C39、第6题A44:C47、第7题A52:C55、第8题A60:C64、第9题A69:C72、第10题A77:C81、第11题A86:C89、第12题A94:C98、第13题A103:C106、第14题A111:C114、第15题A119:C122、第16题A127:C130、第17题A135:C138、第18题A143:C146、第19题A151:C154、第20题A159:C162、第21题A167:C171、第22题A176:C179、第23题A184:C188、第24题A193:C197、第25题A202:C206。",
    "“sheet1”图表位置：22个图表均位于对应题目右侧空白区域，图表左边缘位于D列及其右侧，未覆盖A:C列原始数据。",
    "“sheet1”图表标签：22个二维饼状图均显示主要占比信息，数据标签包含百分比、数值或两者之一，主要扇区可辨认。",
    "“sheet1”图表标题：22个图表均带有标题，标题包含对应题号“第4题”至“第25题”或能够明确识别对应题目内容。",
    "“sheet1”图表字体：图表标题、数据标签、图例文字采用宋体、微软雅黑、Calibri等常规字体，字号5–20磅，颜色清晰可读，无艺术字。",
    "“sheet1”图表图例：22个图表均能通过图例或数据标签识别各扇区对应的选项内容。",
    "“sheet1”图表尺寸：每个图表高度约5–10cm。",
]


def ws_get_text(wb, row):
    """取 sheet1 指定行 A 列文本（用于题干/选项识别）。"""
    try:
        v = wb["sheet1"].cell(row=row, column=1).value
        return str(v).strip() if v is not None else ""
    except Exception:
        return ""


# ======================== 主流程 ========================

# 维度二各评分项的满分（顺序与 evaluate_dim2 内 add_rule 追加顺序、SPEC_TEXTS 一一对应）
MAX_DELTAS = [5, 3, 1, 5, 1, 1, 5, 3, -3, -1]


def _locate_document(dir_path: str):
    """在 dir_path 目录中定位待评估的 .xlsx/.xlsm 文档；过滤 ~$ 临时锁文件。
    找到返回绝对路径，未找到返回 None。"""
    try:
        candidates = []
        for name in os.listdir(dir_path):
            if name.startswith("~$"):
                continue
            low = name.lower()
            if low.endswith((".xlsx", ".xlsm")):
                candidates.append(os.path.join(dir_path, name))
        if not candidates:
            return None
        # 若目录内存在多个候选文档，优先选择文件名含“饼图”者
        for c in candidates:
            if "饼图" in os.path.basename(c):
                return c
        return candidates[0]
    except Exception:
        return None


def evaluate(dir_path: str) -> dict:
    """在 dir_path 目录中定位并评估 xlsx 文档，返回结构化评分结果。

    参数：
        dir_path: 脚本所在目录的绝对路径。脚本自身在该目录中定位被评估的文档。

    返回：见 officeval 脚本接口约定中的返回数据结构。
    不抛异常：脚本自身错误统一以 status="error" 表达。
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
        "max_score": sum(d for d in MAX_DELTAS if d > 0),
    }

    try:
        path = _locate_document(dir_path)
        if not path:
            result["status"] = "error"
            result["error"] = "FileNotFoundError: 未在目录中找到 .xlsx/.xlsm 文档"
            return result
        result["file_name"] = os.path.basename(path)

        # ---- 维度1 ----
        dim1_ok, d1_details, wb, eval_path = evaluate_dim1(path)
        result["dim1_pass"] = bool(dim1_ok)
        if not dim1_ok:
            reasons = [msg for _name, ok, msg in d1_details if not ok]
            result["dim1_reason"] = "；".join(r for r in reasons if r)
            # 维度一未通过 -> dim2_items 为空，total_score=0
            return result

        # ---- 维度2 ----
        total, hits, _infos = evaluate_dim2(wb, eval_path)
        items = []
        for idx, (rule, delta, ok_or_hit, detail) in enumerate(hits):
            spec = SPEC_TEXTS[idx] if idx < len(SPEC_TEXTS) else rule
            actual = delta if ok_or_hit else 0
            items.append({
                "rule": spec,
                "max_delta": delta,
                "delta": actual,
                "hit": bool(ok_or_hit),
                "detail": "",
            })
        result["dim2_items"] = items
        result["total_score"] = total
        return result
    except Exception as e:
        result["status"] = "error"
        result["error"] = "%s: %s" % (type(e).__name__, e)
        return result


if __name__ == "__main__":
    _dir = sys.argv[1] if len(sys.argv) > 1 else os.path.dirname(os.path.abspath(__file__))
    print(json.dumps(evaluate(_dir), ensure_ascii=False, indent=2))
