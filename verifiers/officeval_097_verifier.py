# -*- coding: utf-8 -*-
"""
自动评估脚本：对 "349950873_1774849271pJJksQ_已插入图表.xlsx" 按照"打分细则"评分。

评估逻辑：
  1. 先检查维度1（可用与可修改性）。任意一条不满足 -> 直接 0 分，不再检查维度2。
  2. 满足维度1后，检查维度2（完成度）。维度2分为得分点(+)与扣分点(-)。
     - 加分细则：必须满足该细则内的"每一个点"才加分。
     - 扣分细则：只要满足该细则内的"任意一点"即扣分。
  3. 累加所有命中的分值（可正可负），打印命中的点与最终得分。

实现说明：
  - 直接解析 xlsx（zip + xml），读取图表 xml / drawing xml / 工作表数据，
    判断图表类型、数据源、数据标签、尺寸、位置等，全部自动化、不依赖人工。
  - 对"配色一致""文字遮挡""乱码"等难以严格量化的点，采用结构化的可量化代理判据
    （例如：同类图表 style/legend/dLbls 配置是否一致；是否存在 #REF!/#VALUE!；
    是否存在无数据源等），以贴近评估意图的方式灵活实现。
"""

import os
import re
import sys
import zipfile
import xml.etree.ElementTree as ET

SCRIPT_ID = "097"

# ----------------------------------------------------------------------------
# 配置
# ----------------------------------------------------------------------------
# 命名空间
NS = {
    "c":  "http://schemas.openxmlformats.org/drawingml/2006/chart",
    "a":  "http://schemas.openxmlformats.org/drawingml/2006/main",
    "r":  "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "xdr": "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
    "ss": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
}

EMU_PER_CM = 360000.0  # 1 cm = 360000 EMU

# 各题期望：题号 -> (期望图表类型, 期望数据区间起止行, 细则规定的完整数据源矩形)
# 类型: 'pie' 饼图 / 'bar' 横向条形图
# rng: "选项—小计—比例"数据行区间（不含题干/表头）
# src: 评分细则逐题明确写出的"选项—小计—比例"数据区域(A列..C列, 起行..止行)
QUESTION_SPEC = {
    6:  {"type": "pie", "rng": (4, 8),     "src": ("A", "C", 4, 8)},
    7:  {"type": "pie", "rng": (13, 15),   "src": ("A", "C", 13, 15)},
    8:  {"type": "bar", "rng": (20, 26),   "src": ("A", "C", 20, 26)},
    9:  {"type": "bar", "rng": (31, 38),   "src": ("A", "C", 31, 38)},
    10: {"type": "pie", "rng": (43, 46),   "src": ("A", "C", 43, 46)},
    11: {"type": "pie", "rng": (51, 55),   "src": ("A", "C", 51, 55)},
    12: {"type": "pie", "rng": (59, 63),   "src": ("A", "C", 59, 63)},
    13: {"type": "bar", "rng": (68, 72),   "src": ("A", "C", 68, 72)},
    14: {"type": "bar", "rng": (77, 83),   "src": ("A", "C", 77, 83)},
    15: {"type": "bar", "rng": (88, 94),   "src": ("A", "C", 88, 94)},
    16: {"type": "pie", "rng": (99, 103),  "src": ("A", "C", 99, 103)},
    17: {"type": "bar", "rng": (108, 114), "src": ("A", "C", 108, 114)},
    18: {"type": "bar", "rng": (119, 126), "src": ("A", "C", 119, 126)},
    19: {"type": "bar", "rng": (131, 136), "src": ("A", "C", 131, 136)},
    20: {"type": "bar", "rng": (141, 148), "src": ("A", "C", 141, 148)},
    21: {"type": "bar", "rng": (153, 158), "src": ("A", "C", 153, 158)},
}

PIE_QUESTIONS = [q for q, s in QUESTION_SPEC.items() if s["type"] == "pie"]
BAR_QUESTIONS = [q for q, s in QUESTION_SPEC.items() if s["type"] == "bar"]


# ----------------------------------------------------------------------------
# 解析底层数据
# ----------------------------------------------------------------------------
def col_letter(idx):
    """0-based 列号 -> 字母 (0->A)。"""
    s = ""
    idx += 1
    while idx:
        idx, rem = divmod(idx - 1, 26)
        s = chr(65 + rem) + s
    return s


def parse_charts(zf):
    """解析所有 chartN.xml，返回 {chart_path: info}。"""
    charts = {}
    for name in zf.namelist():
        m = re.match(r"xl/charts/(chart\d+)\.xml$", name)
        if not m:
            continue
        raw = zf.read(name).decode("utf-8", errors="replace")
        info = {
            "raw": raw,
            "title": None,
            "chart_type": None,   # pie / bar / col / line / other
            "bar_dir": None,      # bar(横) / col(纵)
            "cat_ref": None,
            "val_ref": None,
            "show_val": False,
            "show_percent": False,
            "show_cat": False,
            "has_legend": "<legend>" in raw or "<legend/>" in raw,
            "style": None,
            "has_ref_error": ("#REF!" in raw or "#VALUE!" in raw),
        }
        # 标题
        tm = re.search(r"<title>.*?<a:t>(.*?)</a:t>", raw, re.S)
        if tm:
            info["title"] = tm.group(1)
        # style
        sm = re.search(r'<style val="(\d+)"', raw)
        if sm:
            info["style"] = sm.group(1)
        # 图表类型
        if "<doughnutChart>" in raw:
            info["chart_type"] = "doughnut"   # 等效圆形占比图（环形图）
        elif "<pieChart>" in raw or "<pie3DChart>" in raw or "<ofPieChart>" in raw:
            info["chart_type"] = "pie"
        elif "<barChart>" in raw or "<bar3DChart>" in raw:
            bd = re.search(r'<barDir val="(\w+)"', raw)
            info["bar_dir"] = bd.group(1) if bd else None
            info["chart_type"] = "bar" if info["bar_dir"] == "bar" else "col"
        elif "<lineChart>" in raw:
            info["chart_type"] = "line"
        else:
            info["chart_type"] = "other"
        # 数据源（取第一个 cat/val 的引用）
        cm = re.search(r"<cat>.*?<f>(.*?)</f>", raw, re.S)
        if cm:
            info["cat_ref"] = cm.group(1)
        # 类别引用是否为字符串引用 <strRef>（而非 <numRef>）。
        # 选项文本应通过 strRef 引用 A 列文本；若用 numRef，WPS/Excel 渲染
        # 标签/类别时易出现占比不显示或异常。
        cat_seg = re.search(r"<cat>(.*?)</cat>", raw, re.S)
        info["cat_is_strref"] = bool(cat_seg and "<strRef>" in cat_seg.group(1))
        vm = re.search(r"<val>.*?<f>(.*?)</f>", raw, re.S)
        if vm:
            info["val_ref"] = vm.group(1)
        # 收集该图表内所有 <f> 引用（用于判断数据源覆盖的整体列/行范围）
        info["all_refs"] = re.findall(r"<f>(.*?)</f>", raw, re.S)
        # 类别轴/数值轴位置（横向条形图：类别轴在左 axPos=l，数值轴在底 axPos=b）
        cat_ax = re.search(r"<catAx>.*?<axPos val=\"(\w)\"", raw, re.S)
        val_ax = re.search(r"<valAx>.*?<axPos val=\"(\w)\"", raw, re.S)
        info["cat_ax_pos"] = cat_ax.group(1) if cat_ax else None
        info["val_ax_pos"] = val_ax.group(1) if val_ax else None
        # 类别引用是否指向选项文本所在的 A 列
        cat_span = ref_span([info.get("cat_ref")]) if info.get("cat_ref") else None
        info["cat_is_col_a"] = bool(cat_span and cat_span[1] == 1 and cat_span[2] == 1)
        # 值引用是否指向小计(B列)或比例(C列)
        val_span = ref_span([info.get("val_ref")]) if info.get("val_ref") else None
        info["val_is_b_or_c"] = bool(val_span and val_span[1] >= 2 and val_span[2] <= 3)
        # 数据标签
        dm = re.search(r"<dLbls>(.*?)</dLbls>", raw, re.S)
        seg = dm.group(1) if dm else raw
        info["show_val"] = '<showVal val="1"' in seg
        info["show_percent"] = '<showPercent val="1"' in seg
        info["show_cat"] = '<showCatName val="1"' in seg
        charts[name] = info
    return charts


def parse_drawing(zf):
    """解析 drawing1.xml，返回锚点列表（含图表 rId、起始行列、宽高 EMU）。"""
    anchors = []
    draw_name = None
    for name in zf.namelist():
        if re.match(r"xl/drawings/drawing\d+\.xml$", name):
            draw_name = name
            break
    if not draw_name:
        return anchors, None
    raw = zf.read(draw_name).decode("utf-8", errors="replace")
    root = ET.fromstring(raw)

    # 建立 drawing rels: rId -> chart 路径
    rels = {}
    rels_path = re.sub(r"drawings/(drawing\d+\.xml)$",
                       r"drawings/_rels/\1.rels", draw_name)
    if rels_path in zf.namelist():
        rroot = ET.fromstring(zf.read(rels_path))
        for rel in rroot.findall("rel:Relationship", NS):
            rid = rel.get("Id")
            tgt = rel.get("Target")
            # Target 相对于 drawing 文件所在目录(xl/drawings/)解析
            base_dir = os.path.dirname(draw_name)  # e.g. xl/drawings
            resolved = os.path.normpath(os.path.join(base_dir, tgt))
            resolved = resolved.replace("\\", "/").lstrip("/")
            rels[rid] = resolved

    for anchor in list(root):
        tag = anchor.tag.split("}")[-1]
        if tag not in ("oneCellAnchor", "twoCellAnchor"):
            continue
        a = {"anchor_type": tag, "from_col": None, "from_row": None,
             "to_col": None, "to_row": None, "cx": None, "cy": None,
             "chart": None, "is_chart": False, "is_picture": False}

        frm = anchor.find("xdr:from", NS)
        if frm is not None:
            c = frm.find("xdr:col", NS)
            rr = frm.find("xdr:row", NS)
            a["from_col"] = int(c.text) if c is not None else None
            a["from_row"] = int(rr.text) if rr is not None else None
        to = anchor.find("xdr:to", NS)
        if to is not None:
            c = to.find("xdr:col", NS)
            rr = to.find("xdr:row", NS)
            a["to_col"] = int(c.text) if c is not None else None
            a["to_row"] = int(rr.text) if rr is not None else None
        ext = anchor.find("xdr:ext", NS)
        if ext is not None:
            a["cx"] = int(ext.get("cx")) if ext.get("cx") else None
            a["cy"] = int(ext.get("cy")) if ext.get("cy") else None

        # 是否为图表对象
        chart_el = anchor.find(".//c:chart", NS)
        if chart_el is not None:
            rid = chart_el.get("{%s}id" % NS["r"])
            a["chart"] = rels.get(rid)
            a["is_chart"] = True
        # 是否为图片/截图对象（<xdr:pic> 内嵌位图，含粘贴的截图）
        if anchor.find(".//xdr:pic", NS) is not None:
            a["is_picture"] = True
        anchors.append(a)
    return anchors, raw


def load_worksheet_cells(zf):
    """读取 sheet1 单元格文本（含共享字符串），返回 {(row,col): text}。"""
    # 共享字符串
    shared = []
    if "xl/sharedStrings.xml" in zf.namelist():
        sroot = ET.fromstring(zf.read("xl/sharedStrings.xml"))
        for si in sroot.findall("ss:si", NS):
            txt = "".join(t.text or "" for t in si.iter("{%s}t" % NS["ss"]))
            shared.append(txt)

    # 找到 sheet1 对应的 worksheet 文件
    sheet_path = "xl/worksheets/sheet1.xml"
    if sheet_path not in zf.namelist():
        # 退化：取第一个 worksheet
        for n in zf.namelist():
            if re.match(r"xl/worksheets/sheet\d+\.xml$", n):
                sheet_path = n
                break
    cells = {}
    merged = []
    if sheet_path in zf.namelist():
        wroot = ET.fromstring(zf.read(sheet_path))
        for c in wroot.iter("{%s}c" % NS["ss"]):
            ref = c.get("r")
            t = c.get("t")
            v = c.find("ss:v", NS)
            is_el = c.find("ss:is", NS)
            text = None
            if t == "s" and v is not None:
                idx = int(v.text)
                text = shared[idx] if idx < len(shared) else None
            elif is_el is not None:
                text = "".join(x.text or "" for x in is_el.iter("{%s}t" % NS["ss"]))
            elif v is not None:
                text = v.text
            m = re.match(r"([A-Z]+)(\d+)", ref)
            if m:
                col = 0
                for ch in m.group(1):
                    col = col * 26 + (ord(ch) - 64)
                cells[(int(m.group(2)), col)] = text
        mc = wroot.find("ss:mergeCells", NS)
        if mc is not None:
            for r in mc.findall("ss:mergeCell", NS):
                merged.append(r.get("ref"))
    return cells, merged


# ----------------------------------------------------------------------------
# 把图表与题目对应起来
# ----------------------------------------------------------------------------
def ref_rows(ref):
    """从形如 'sheet1'!$A$4:$A$8 提取起止行号。"""
    if not ref:
        return None
    nums = re.findall(r"\$?[A-Z]+\$?(\d+)", ref)
    if len(nums) >= 2:
        return (int(nums[0]), int(nums[1]))
    if len(nums) == 1:
        return (int(nums[0]), int(nums[0]))
    return None


def overlaps(a1, a2, b1, b2):
    return max(a1, b1) <= min(a2, b2)


def col_to_num(letter):
    """列字母 -> 数字 (A=1)。"""
    n = 0
    for ch in letter:
        n = n * 26 + (ord(ch.upper()) - 64)
    return n


def ref_span(refs):
    """
    从一组形如 'sheet1'!$A$4:$C$8 / 'sheet1'!$B$4:$B$8 的引用中，
    汇总出覆盖的列范围(最小列..最大列)与行范围(最小行..最大行)。
    返回 (sheet, min_col, max_col, min_row, max_row) 或 None。
    """
    cols, rows, sheets = [], [], set()
    for ref in refs:
        if not ref:
            continue
        sm = re.match(r"'?([^'!]+)'?!", ref)
        if sm:
            sheets.add(sm.group(1).strip("'").lower())
        for cell in re.findall(r"\$?([A-Z]+)\$?(\d+)", ref):
            cols.append(col_to_num(cell[0]))
            rows.append(int(cell[1]))
    if not cols or not rows:
        return None
    sheet = next(iter(sheets)) if sheets else None
    return (sheet, min(cols), max(cols), min(rows), max(rows))


def match_charts_to_questions(charts, anchors):
    """
    为每个题目匹配图表。匹配策略：
      1. 优先按数据源行区间与题目期望区间重叠匹配；
      2. 退化时按标题"第N题"匹配。
    返回 {q: {anchor, chart_info, chart_path}} 以及未匹配图表列表。
    """
    chart_anchor = []  # (anchor, chart_info, path)
    for a in anchors:
        if a["is_chart"] and a["chart"] in charts:
            chart_anchor.append((a, charts[a["chart"]], a["chart"]))

    matched = {}
    used = set()

    # 第一轮：按数据源行区间匹配
    for q, spec in QUESTION_SPEC.items():
        qr = spec["rng"]
        for i, (a, ci, p) in enumerate(chart_anchor):
            if i in used:
                continue
            rr = ref_rows(ci.get("val_ref")) or ref_rows(ci.get("cat_ref"))
            if rr and overlaps(qr[0], qr[1], rr[0], rr[1]):
                matched[q] = {"anchor": a, "info": ci, "path": p}
                used.add(i)
                break

    # 第二轮：按标题匹配未匹配题目
    for q, spec in QUESTION_SPEC.items():
        if q in matched:
            continue
        for i, (a, ci, p) in enumerate(chart_anchor):
            if i in used:
                continue
            if ci.get("title") and ("第%d题" % q) in ci["title"]:
                matched[q] = {"anchor": a, "info": ci, "path": p}
                used.add(i)
                break

    unmatched = [chart_anchor[i] for i in range(len(chart_anchor)) if i not in used]
    return matched, unmatched, chart_anchor


# ----------------------------------------------------------------------------
# 维度1：可用与可修改性（任意一条不满足 -> 0 分）
# ----------------------------------------------------------------------------
def check_dimension1(path, zf, cells, charts, anchors):
    """返回 (是否通过, 详细结果列表[(描述, bool)])。"""
    results = []

    # 1) 交付文件为xlsx或.xlsm格式，文件可正常打开。
    ext_ok = path.lower().endswith((".xlsx", ".xlsm"))
    open_ok = True  # 能走到这里说明 zip 已成功打开
    try:
        _ = zf.namelist()
    except Exception:
        open_ok = False
    results.append(("交付文件为xlsx或.xlsm格式，文件可正常打开", ext_ok and open_ok))

    passed = all(ok for _, ok in results)
    return passed, results


# ----------------------------------------------------------------------------
# 维度2：完成度（加分点 + 扣分点）
# ----------------------------------------------------------------------------
def check_dimension2(matched, unmatched, all_chart_anchor, charts, anchors,
                     cells, merged):
    """返回 (得分, 命中明细列表[(分值, 描述, 命中bool)])。"""
    items = []  # (score, desc, hit)

    n_chart_total = len(all_chart_anchor)

    # +5 图表总数：第6题至第21题共新增16个图表，每个有统计数据的题目右侧空白处
    #    各对应1个图表，第22题开放题不要求绘制图表。
    #    逐点核对（每一点都要踩到）：
    #    点1) 第6–21题共新增 16 个图表 -> 这16题每题都匹配到图表，且图表对象总数为16；
    #    点2) 每题右侧空白处各对应1个图表 -> 每题恰好1个图表，且锚定在C列右侧(from_col>=3)；
    #    点3) 第22题开放题不要求绘制图表 -> 第22题区域(行161-162)右侧无新增图表。
    Q_RANGE = list(range(6, 22))  # 6..21

    # 点1：6–21题各匹配到图表，且新增图表对象总数恰为16
    p1_each_matched = all(q in matched for q in Q_RANGE)
    p1_total_16 = (len(matched) == 16 and n_chart_total == 16)
    point1 = p1_each_matched and p1_total_16

    # 点2：每个有统计数据的题目右侧各对应"1个"图表
    #   - 每题恰好1个：matched 已是每题取一个，需额外确认没有"多个图表落在同一题数据区"
    #   - 右侧空白处：图表锚定列在 C 列右侧（0-based from_col >= 3，即 D 列及以后）
    #   统计落入每个题目数据行区间的图表数量
    per_q_count = {q: 0 for q in Q_RANGE}
    for a, ci, p in all_chart_anchor:
        rr = ref_rows(ci.get("val_ref")) or ref_rows(ci.get("cat_ref"))
        if not rr:
            continue
        for q in Q_RANGE:
            qr = QUESTION_SPEC[q]["rng"]
            if overlaps(qr[0], qr[1], rr[0], rr[1]):
                per_q_count[q] += 1
                break
    p2_each_one = all(per_q_count[q] == 1 for q in Q_RANGE)
    p2_right_side = all(matched[q]["anchor"].get("from_col") is not None
                        and matched[q]["anchor"]["from_col"] >= 3
                        for q in Q_RANGE if q in matched)
    point2 = p2_each_one and p2_right_side

    # 点3：第22题开放题不要求绘制图表 -> 第22题区域(行161-162, 0-based 160-161)右侧无图表
    q22_rows = (161, 162)
    q22_has_chart = any(
        a["is_chart"] and a.get("from_row") is not None
        and q22_rows[0] - 1 <= a["from_row"] <= q22_rows[1] - 1
        for a in anchors)
    point3 = not q22_has_chart

    cond = point1 and point2 and point3
    items.append((5, "图表总数：第6–21题共新增16个图表/每个有统计数据题目右侧各1个/"
                  "第22题不绘图(共16题匹配=%s且对象总数=%d; 每题各1图=%s; "
                  "均在C列右侧=%s; 第22题无图=%s)"
                  % (p1_each_matched, n_chart_total, p2_each_one,
                     p2_right_side, point3), cond))


    # +5 饼图题型：第6题、第7题、第10题、第11题、第12题、第16题右侧各有1个饼状图，
    #    图表类型为饼图或等效圆形占比图，不是条形图、柱形图或折线图。
    #    逐点核对（每一点都要踩到）：
    #    点1) 指定6题"右侧各有1个"饼状图 -> 每题恰好匹配到1个图表且锚定在C列右侧(from_col>=3)；
    #    点2) 图表类型为饼图或等效圆形占比图 -> chart_type 属于 {pie, doughnut} 等圆形占比族；
    #    点3) 不是条形图、柱形图或折线图 -> chart_type 不属于 {bar, col, line}。
    PIE_OK_TYPES = {"pie", "doughnut"}          # 饼图/等效圆形占比图
    PIE_BAD_TYPES = {"bar", "col", "line"}      # 明确排除：条形图/柱形图/折线图

    # 点1：每个指定题目右侧恰好 1 个饼图（统计落入该题数据区的图表数）
    pie_count = {q: 0 for q in PIE_QUESTIONS}
    for a, ci, p in all_chart_anchor:
        rr = ref_rows(ci.get("val_ref")) or ref_rows(ci.get("cat_ref"))
        if not rr:
            continue
        for q in PIE_QUESTIONS:
            qr = QUESTION_SPEC[q]["rng"]
            if overlaps(qr[0], qr[1], rr[0], rr[1]):
                pie_count[q] += 1
                break
    p_each_one = all(pie_count[q] == 1 for q in PIE_QUESTIONS)
    p_right_side = all(q in matched
                       and matched[q]["anchor"].get("from_col") is not None
                       and matched[q]["anchor"]["from_col"] >= 3
                       for q in PIE_QUESTIONS)
    # 点2：类型为饼图/等效圆形占比图
    p_is_pie = all(q in matched
                   and matched[q]["info"]["chart_type"] in PIE_OK_TYPES
                   for q in PIE_QUESTIONS)
    # 点3：不是条形图/柱形图/折线图
    p_not_bad = all(q in matched
                    and matched[q]["info"]["chart_type"] not in PIE_BAD_TYPES
                    for q in PIE_QUESTIONS)
    miss_pie = [q for q in PIE_QUESTIONS
                if not (q in matched
                        and matched[q]["info"]["chart_type"] in PIE_OK_TYPES
                        and matched[q]["info"]["chart_type"] not in PIE_BAD_TYPES)]
    pie_ok = p_each_one and p_right_side and p_is_pie and p_not_bad
    items.append((5, "饼图题型：第%s题右侧各1个饼图/类型为饼图或等效圆形占比图/"
                  "非条形柱形折线(各题1图=%s; 在C列右侧=%s; 是圆形占比图=%s; "
                  "非条柱线=%s; 不达标题=%s)"
                  % ("/".join(map(str, PIE_QUESTIONS)), p_each_one, p_right_side,
                     p_is_pie, p_not_bad, miss_pie or "无"), pie_ok))


    # +5 横向条形图题型：第8/9/13/14/15/17/18/19/20/21题右侧各有1个横向条形图，
    #    类别横向排列，图表类型不是饼图或纵向柱形图。
    #    逐点核对（每一点都要踩到）：
    #    点1) 指定10题"右侧各有1个"横向条形图 -> 每题恰好1个图表且锚定在C列右侧(from_col>=3)；
    #    点2) 横向条形图 + 类别横向排列 -> chart_type=="bar"(barChart 且 barDir=bar，
    #         即条形横向延伸、类别沿纵轴逐条横向排列)；
    #    点3) 图表类型不是饼图或纵向柱形图 -> chart_type 不属于 {pie, doughnut, col}。
    BAR_BAD_TYPES = {"pie", "doughnut", "col"}   # 明确排除：饼图/等效圆形占比图/纵向柱形图

    # 点1：每个指定题目右侧恰好 1 个横向条形图
    bar_count = {q: 0 for q in BAR_QUESTIONS}
    for a, ci, p in all_chart_anchor:
        rr = ref_rows(ci.get("val_ref")) or ref_rows(ci.get("cat_ref"))
        if not rr:
            continue
        for q in BAR_QUESTIONS:
            qr = QUESTION_SPEC[q]["rng"]
            if overlaps(qr[0], qr[1], rr[0], rr[1]):
                bar_count[q] += 1
                break
    b_each_one = all(bar_count[q] == 1 for q in BAR_QUESTIONS)
    b_right_side = all(q in matched
                       and matched[q]["anchor"].get("from_col") is not None
                       and matched[q]["anchor"]["from_col"] >= 3
                       for q in BAR_QUESTIONS)
    # 点2：横向条形图，类别横向排列（barChart 且 barDir=bar）
    b_is_hbar = all(q in matched
                    and matched[q]["info"]["chart_type"] == "bar"
                    and matched[q]["info"]["bar_dir"] == "bar"
                    for q in BAR_QUESTIONS)
    # 点3：不是饼图或纵向柱形图
    b_not_bad = all(q in matched
                    and matched[q]["info"]["chart_type"] not in BAR_BAD_TYPES
                    for q in BAR_QUESTIONS)
    miss_bar = [q for q in BAR_QUESTIONS
                if not (q in matched
                        and matched[q]["info"]["chart_type"] == "bar"
                        and matched[q]["info"]["bar_dir"] == "bar"
                        and matched[q]["info"]["chart_type"] not in BAR_BAD_TYPES)]
    bar_ok = b_each_one and b_right_side and b_is_hbar and b_not_bad
    items.append((5, "横向条形图题型：第%s题右侧各1个横向条形图/类别横向排列/"
                  "非饼图或纵向柱形图(各题1图=%s; 在C列右侧=%s; 横向条形图=%s; "
                  "非饼非纵柱=%s; 不达标题=%s)"
                  % ("/".join(map(str, BAR_QUESTIONS)), b_each_one, b_right_side,
                     b_is_hbar, b_not_bad, miss_bar or "无"), bar_ok))


    # +5 图表数据源：16个图表均使用各自题目对应的"选项—小计—比例"数据区域作为数据源，
    #    分别来自第6题A4:C8、第7题A13:C15、第8题A20:C26、第9题A31:C38、第10题A43:C46、
    #    第11题A51:C55、第12题A59:C63、第13题A68:C72、第14题A77:C83、第15题A88:C94、
    #    第16题A99:C103、第17题A108:C114、第18题A119:C126、第19题A131:C136、
    #    第20题A141:C148、第21题A153:C158。
    #    逐点核对（每一点都要踩到）：
    #    点1) 16个图表均有数据源 -> 6–21每题都匹配到图表且图表内含数据源引用；
    #    点2) 数据源为各自题目对应的"选项—小计—比例"区域 -> 图表引用所在工作表为 sheet1，
    #         覆盖的列范围落在 A..C(选项A列/小计B列/比例C列)之内，行范围与该题
    #         细则规定的数据行(起..止)相符；
    #    点3) "分别来自第N题 A?:C?" 各题精确区间 -> 图表数据源行区间与细则逐题写出的
    #         起止行一致(允许 cat/val 分列引用合并后落在 A..C × 指定行 的矩形内)。
    src_ok = True
    bad_src = []
    for q, spec in QUESTION_SPEC.items():
        sc_col, ec_col, sr, er = spec["src"]
        sc, ec = col_to_num(sc_col), col_to_num(ec_col)   # A=1 .. C=3
        if q not in matched:
            src_ok = False
            bad_src.append("%d(无图表)" % q)
            continue
        ci = matched[q]["info"]
        refs = ci.get("all_refs") or [ci.get("cat_ref"), ci.get("val_ref")]
        span = ref_span(refs)
        if span is None:
            src_ok = False
            bad_src.append("%d(无数据源)" % q)
            continue
        sheet, mnc, mxc, mnr, mxr = span
        # 点2：工作表为 sheet1，列落在 A..C 之内
        sheet_ok = (sheet == "sheet1") if sheet else True
        cols_ok = (mnc >= sc and mxc <= ec)
        # 点3：行范围与该题细则规定的起止行一致（图表实际引用的数据行=细则数据行）
        rows_ok = (mnr == sr and mxr == er)
        if not (sheet_ok and cols_ok and rows_ok):
            src_ok = False
            detail = "%d(实际 sheet=%s 列%d-%d 行%d-%d; 期望 sheet1 列1-3 行%d-%d)" \
                     % (q, sheet, mnc, mxc, mnr, mxr, sr, er)
            bad_src.append(detail)
    items.append((5, "图表数据源：16图表均使用各自题目对应的选项—小计—比例区域"
                  "(A?:C? 精确区间, 不匹配=%s)" % (bad_src or "无"), src_ok))


    # +5 饼图标签：第6/7/10/11/12/16题饼图均能显示各选项占比，数据标签包含百分比、
    #    数值或两者之一，主要扇区可辨认。
    #    逐点核对（每一点都要踩到）：
    #    点1) 是这6题的饼图 -> chart_type 属于圆形占比族 {pie, doughnut}；
    #    点2) 数据标签包含百分比、数值或两者之一 -> showPercent 或 showVal 真实启用其一即可，
    #         使每个扇区都带标签，从而能体现各选项占比；
    #    点3) 主要扇区可辨认 -> 细则未规定量化阈值，不用无来源的固定比例/切片数硬阈值代理
    #         判断视觉可辨认；只要图表本身是合法的圆形占比图且开启了标签(点1/点2成立)，
    #         即视为主要扇区可辨认，避免用与实际渲染无关的代理指标误判合格图表。
    pie_bad = []
    for q in PIE_QUESTIONS:
        if q not in matched:
            pie_bad.append("%d(无图表)" % q)
            continue
        ci = matched[q]["info"]
        # 点1：是圆形占比族
        is_pie = ci["chart_type"] in {"pie", "doughnut"}
        # 点2：数据标签包含百分比、数值或两者之一——只看 showPercent/showVal 是否真实启用
        has_label = ci["show_percent"] or ci["show_val"]
        if not (is_pie and has_label):
            pie_bad.append("%d(饼图=%s 含百分比或数值标签=%s)"
                           % (q, is_pie, has_label))
    pielbl_ok = (len(pie_bad) == 0)
    items.append((5, "饼图标签：第%s题饼图显示各选项占比/标签含百分比或数值之一/"
                  "主要扇区可辨认(不达标=%s)"
                  % ("/".join(map(str, PIE_QUESTIONS)), pie_bad or "无"), pielbl_ok))


    # +5 横向条形图方向：第8/9/13/14/15/17/18/19/20/21题图表均为横向条形图，
    #    选项文本位于纵轴类别区域，小计或比例通过横向条形长度全部体现。
    #    逐点核对（每一点都要踩到）：
    #    点1) 均为横向条形图 -> chart_type=="bar" 且 barDir=="bar"；
    #    点2) 选项文本位于纵轴类别区域 -> 类别轴(catAx)位于左侧纵轴(axPos=l)，
    #         且类别引用指向选项文本所在的 A 列(cat_is_col_a)；
    #    点3) 小计或比例通过横向条形长度"全部"体现 -> 数值引用指向小计(B列)或比例(C列)
    #         (val_is_b_or_c)，且数值引用的行区间"完整覆盖"该题细则规定的全部选项行
    #         (起止行与 QUESTION_SPEC[q]["src"] 一致)，即每个选项行都对应一条横条、
    #         无遗漏，小计/比例才算被全部体现。
    bardir_bad = []
    for q in BAR_QUESTIONS:
        if q not in matched:
            bardir_bad.append("%d(无图表)" % q)
            continue
        ci = matched[q]["info"]
        # 点1
        is_hbar = (ci["chart_type"] == "bar" and ci["bar_dir"] == "bar")
        # 点2：选项文本在纵轴类别区域。横向条形图(barDir=bar)的类别轴天然为纵轴，
        #      再确认类别引用指向选项文本所在的 A 列。
        cat_on_yaxis = is_hbar and ci.get("cat_is_col_a")
        # 点3：小计或比例通过横向条形长度"全部"体现。
        #   - 数值引用指向小计(B列)或比例(C列)；
        #   - 数值引用行区间完整覆盖该题全部选项行(起止行与细则一致)，每个选项都成条；
        #   - 类别引用必须是字符串引用 <strRef>(而非 <numRef>)——只有 strRef 才能让 A 列
        #     选项文本作为合法类别供 WPS/Excel 正确渲染，否则条形/小计标签会显示不全。
        _, _, sr, er = QUESTION_SPEC[q]["src"]
        vspan = ref_span([ci.get("val_ref")])
        val_full_rows = bool(vspan and vspan[3] == sr and vspan[4] == er)
        val_horizontal = (is_hbar and ci.get("val_is_b_or_c")
                          and val_full_rows and ci.get("cat_is_strref"))
        if not (is_hbar and cat_on_yaxis and val_horizontal):
            bardir_bad.append("%d(横向条形=%s 选项在纵轴=%s 值横向全部延伸=%s)"
                              % (q, is_hbar, cat_on_yaxis, val_horizontal))
    bardir_ok = (len(bardir_bad) == 0)
    items.append((5, "横向条形图方向：第%s题均为横向条形图/选项文本位于纵轴类别区域/"
                  "小计或比例通过横向条形长度全部体现(不达标=%s)"
                  % ("/".join(map(str, BAR_QUESTIONS)), bardir_bad or "无"), bardir_ok))


    # +5 横向条形图标签：10个横向条形图能显示小计或比例的数据标签。
    #    逐点核对（每一点都要踩到）：
    #    点1) 是这10题的横向条形图 -> chart_type=="bar" 且 barDir=="bar"；
    #    点2) 能显示小计或比例的数据标签 -> showVal(小计) 或 showPercent(比例) 任一为真。
    barlbl_bad = []
    for q in BAR_QUESTIONS:
        if q not in matched:
            barlbl_bad.append("%d(无图表)" % q)
            continue
        ci = matched[q]["info"]
        # 点1
        is_hbar = (ci["chart_type"] == "bar" and ci["bar_dir"] == "bar")
        # 点2：显示小计或比例数据标签
        has_label = ci["show_val"] or ci["show_percent"]
        if not (is_hbar and has_label):
            barlbl_bad.append("%d(横向条形=%s 含小计/比例标签=%s)"
                              % (q, is_hbar, has_label))
    barlbl_ok = (len(barlbl_bad) == 0)
    items.append((5, "横向条形图标签：第%s题显示小计或比例数据标签(不达标=%s)"
                  % ("/".join(map(str, BAR_QUESTIONS)), barlbl_bad or "无"), barlbl_ok))


    # +5 图表尺寸：每个图表宽度约5–9cm、高度约3–8cm。
    #    逐点核对（每一点都要踩到）：
    #    点1) 每个图表宽度约 5–9cm -> 锚点 ext.cx 换算 cm 落在 [5,9] 区间(约：含端点)；
    #    点2) 每个图表高度约 3–8cm -> 锚点 ext.cy 换算 cm 落在 [3,8] 区间(约：含端点)。
    size_bad = []
    for q in QUESTION_SPEC:
        if q not in matched:
            size_bad.append("%d(无图表)" % q)
            continue
        a = matched[q]["anchor"]
        cx, cy = a.get("cx"), a.get("cy")
        if cx is None or cy is None:
            size_bad.append("%d(无尺寸)" % q)
            continue
        w_cm = cx / EMU_PER_CM
        h_cm = cy / EMU_PER_CM
        # 点1：宽度约 5–9cm
        w_ok = (5.0 <= w_cm <= 9.0)
        # 点2：高度约 3–8cm
        h_ok = (3.0 <= h_cm <= 8.0)
        if not (w_ok and h_ok):
            size_bad.append("%d(宽%.2fcm[%s] 高%.2fcm[%s])"
                            % (q, w_cm, "OK" if w_ok else "NG",
                               h_cm, "OK" if h_ok else "NG"))
    size_ok = (len(matched) == 16 and len(size_bad) == 0)
    items.append((5, "图表尺寸：每个宽约5–9cm、高约3–8cm"
                  "(不达标=%s)" % (size_bad or "无"), size_ok))


    # +5 图表一致性：同类型图表在配色、字体和标签样式上整体一致，
    #    饼图与横向条形图在同一工作表中排列规整。
    #    逐点核对（每一点都要踩到）：
    #    点1) 同类型图表"配色"整体一致 -> 用可视属性归一化后比较：都未显式自定义颜色时，
    #         视为统一沿用主题默认配色，等效一致，不要求 style/varyColors 字段字面相同；
    #         若显式自定义，则比较归一化后的颜色集合(大小写/顺序无关)是否相同；
    #         "都未自定义"与"都自定义同一组颜色"分别构成两种等效一致的情形。
    #    点2) 同类型图表"字体"整体一致 -> 同理归一化：都未显式自定义字体时，视为统一
    #         沿用主题默认字体，等效一致；若显式自定义，比较归一化后的字体名集合。
    #    点3) 同类型图表"标签样式"整体一致 -> 同组图表数据标签的可视表现
    #         (是否显示数值/百分比/类别名/图例) 一致，即可视配置相同即可；
    #    点4) 饼图与横向条形图"在同一工作表中" -> 两类图表的数据源/锚点均位于 sheet1；
    #    点5) 排列规整 -> 所有图表锚定起始列一致(单列对齐)，且按起始行单调递增、互不交叉。
    def color_signature(ci):
        raw = ci["raw"]
        # 归一化：提取所有显式颜色值(不区分大小写)，无自定义颜色则视为"沿用主题默认"，
        # 统一记为空集合——只要都未自定义，就等效一致，不再区分具体 style/varyColors。
        clrs = frozenset(c.upper() for c in
                         re.findall(r'<a:srgbClr val="([0-9A-Fa-f]{6})"', raw))
        scheme_clrs = frozenset(c.lower() for c in
                                re.findall(r'<a:schemeClr val="(\w+)"', raw))
        custom = bool(clrs or scheme_clrs)
        # 配色可视签名：未自定义时归一为同一个"默认主题"标记；自定义时比较实际颜色集合
        return (custom, clrs, scheme_clrs) if custom else (False, None, None)

    def font_signature(ci):
        raw = ci["raw"]
        # 归一化：提取显式指定的字体名(不区分大小写)，未自定义则视为"沿用主题默认字体"，
        # 等效一致，不再要求 style 字段字面相同。
        fonts = frozenset(f.lower() for f in
                          re.findall(r'typeface="([^"]+)"', raw))
        custom = bool(fonts)
        return (custom, fonts) if custom else (False, None)

    def label_signature(ci):
        # 标签样式的可视表现：显示数值/百分比/类别名/图例，任一维度不同即视觉可辨的不一致
        return (ci["show_val"], ci["show_percent"], ci["show_cat"], ci["has_legend"])

    def group_consistent(qs):
        infos = [matched[q]["info"] for q in qs if q in matched]
        if len(infos) < len(qs):
            return False, False, False
        color_ok = len({color_signature(i) for i in infos}) == 1   # 点1 配色一致
        font_ok = len({font_signature(i) for i in infos}) == 1     # 点2 字体一致
        label_ok = len({label_signature(i) for i in infos}) == 1   # 点3 标签样式一致
        return color_ok, font_ok, label_ok

    pie_color, pie_font, pie_label = group_consistent(PIE_QUESTIONS)
    bar_color, bar_font, bar_label = group_consistent(BAR_QUESTIONS)
    color_consistent = pie_color and bar_color
    font_consistent = pie_font and bar_font
    label_consistent = pie_label and bar_label

    # 点4：饼图与横向条形图均在同一工作表(sheet1)
    def all_on_sheet1(qs):
        for q in qs:
            if q not in matched:
                return False
            span = ref_span(matched[q]["info"].get("all_refs") or [])
            sheet = span[0] if span else None
            if sheet and sheet != "sheet1":
                return False
        return True
    same_sheet = all_on_sheet1(PIE_QUESTIONS) and all_on_sheet1(BAR_QUESTIONS)

    # 点5：排列规整——所有图表起始列一致(单列对齐)，且起始行单调递增不交叉
    from_cols = {matched[q]["anchor"]["from_col"] for q in matched
                 if matched[q]["anchor"]["from_col"] is not None}
    col_aligned = len(from_cols) <= 1
    rows = [matched[q]["anchor"]["from_row"] for q in sorted(matched)
            if matched[q]["anchor"]["from_row"] is not None]
    rows_ordered = all(rows[i] < rows[i + 1] for i in range(len(rows) - 1))
    neat = col_aligned and rows_ordered

    consistency_ok = (color_consistent and font_consistent
                      and label_consistent and same_sheet and neat)
    items.append((5, "图表一致性：同类型配色一致/字体一致/标签样式一致/同一工作表/排列规整"
                  "(配色一致=%s, 字体一致=%s, 标签一致=%s, 同表=%s, 列对齐=%s, 行有序=%s)"
                  % (color_consistent, font_consistent, label_consistent,
                     same_sheet, col_aligned, rows_ordered),
                  consistency_ok))


    # ---------- 扣分项（命中任意一点即扣） ----------
    # 三条 -5/-3/-1 扣分规则已按需求删除。

    score = sum(s for s, _, hit in items if hit)
    return score, items


# ----------------------------------------------------------------------------
# 主流程
# ----------------------------------------------------------------------------
# 各评分项分值（顺序须与 check_dimension2 中 items、以及 RUBRIC_TEXT 一致）
ITEM_SCORES = [5, 5, 5, 5, 5, 5, 5, 5, 5]

# 评分细则原文（按 check_dimension2 中 items 的顺序排列）
RUBRIC_TEXT = [
    # 得分点（+5 × 9）
    "“sheet1”图表总数：第6题至第21题共新增16个图表，每个有统计数据的题目右侧空白处各对应1个图表，第22题开放题不要求绘制图表。",
    "“sheet1”饼图题型：第6题、第7题、第10题、第11题、第12题、第16题右侧各有1个饼状图,图表类型为饼图或等效圆形占比图,不是条形图、柱形图或折线图。",
    "“sheet1”横向条形图题型：第8题、第9题、第13题、第14题、第15题、第17题、第18题、第19题、第20题、第21题右侧各有1个横向条形图,类别横向排列,图表类型不是饼图或纵向柱形图。",
    "“sheet1”图表数据源：16个图表均使用各自题目对应的“选项—小计—比例”数据区域作为数据源,分别来自第6题A4:C8、第7题A13:C15、第8题A20:C26、第9题A31:C38、第10题A43:C46、第11题A51:C55、第12题A59:C63、第13题A68:C72、第14题A77:C83、第15题A88:C94、第16题A99:C103、第17题A108:C114、第18题A119:C126、第19题A131:C136、第20题A141:C148、第21题A153:C158。",
    "“sheet1”饼图标签：第6题、第7题、第10题、第11题、第12题、第16题饼图均能显示各选项占比,数据标签包含百分比、数值或两者之一,主要扇区可辨认。",
    "“sheet1”横向条形图方向：第8题、第9题、第13题、第14题、第15题、第17题、第18题、第19题、第20题、第21题图表均为横向条形图,选项文本位于纵轴类别区域,小计或比例通过横向条形长度体现。",
    "“sheet1”横向条形图标签：10个横向条形图能显示小计或比例的数据标签。",
    "“sheet1”图表尺寸：每个图表宽度约5–9cm、高度约3–8cm。",
    "“sheet1”图表一致性：同类型图表在配色、字体和标签样式上整体一致,饼图与横向条形图在同一工作表中排列规整。",
]

MAX_SCORE = sum(s for s in ITEM_SCORES if s > 0)


def _locate_target(dir_path: str):
    """在给定目录中定位被评估的 .xlsx/.xlsm 文档。若传入的是文件路径则原样返回。"""
    if os.path.isfile(dir_path):
        return dir_path
    if not os.path.isdir(dir_path):
        return None
    # 优先选择非临时文件（跳过以 ~$ 开头的 Office 临时文件）
    candidates = []
    for name in os.listdir(dir_path):
        if name.startswith("~$"):
            continue
        if name.lower().endswith((".xlsx", ".xlsm")):
            candidates.append(os.path.join(dir_path, name))
    if not candidates:
        return None
    # 若有多个，取最先出现的一个（目录里通常只有一个待评估文档）
    return candidates[0]


def evaluate(dir_path: str) -> dict:
    """
    统一入口：接收"脚本所在目录的路径"，脚本自己在该目录里定位并打开被评估的
    .xlsx/.xlsm 文档，返回结构化评分结果字典。
    """
    result = {
        "id": "097",
        "file_name": "",
        "status": "ok",
        "error": None,
        "dim1_pass": False,
        "dim1_reason": "",
        "dim2_items": [],
        "total_score": 0,
        "max_score": MAX_SCORE,
    }

    try:
        # 在目录中定位被评估文档
        path = _locate_target(dir_path)
        if not path or not os.path.exists(path):
            result["status"] = "error"
            result["error"] = "未在目录中找到 .xlsx/.xlsm 文档：%s" % dir_path
            return result

        result["file_name"] = os.path.basename(path)

        # 打开（zip 解析失败即视为无法正常打开 -> 维度1直接不通过）
        try:
            zf = zipfile.ZipFile(path)
        except Exception as e:
            result["dim1_pass"] = False
            result["dim1_reason"] = "文件无法正常打开（%s）" % e
            result["total_score"] = 0
            return result

        charts = parse_charts(zf)
        anchors, _ = parse_drawing(zf)
        cells, merged = load_worksheet_cells(zf)

        # ---- 维度1 ----
        d1_pass, d1_results = check_dimension1(path, zf, cells, charts, anchors)
        result["dim1_pass"] = d1_pass
        if not d1_pass:
            # 维度1未通过 -> 直接 0 分，不检查维度2
            reasons = [desc for desc, ok in d1_results if not ok]
            result["dim1_reason"] = "; ".join(reasons)
            result["total_score"] = 0
            return result

        # ---- 维度2 ----
        matched, unmatched, all_chart_anchor = match_charts_to_questions(charts, anchors)
        score, items = check_dimension2(matched, unmatched, all_chart_anchor,
                                        charts, anchors, cells, merged)

        dim2_items = []
        for (s, desc, hit), rubric in zip(items, RUBRIC_TEXT):
            dim2_items.append({
                "rule": rubric,
                "max_delta": s,
                "delta": s if hit else 0,
                "hit": bool(hit),
                "detail": "",
            })
        result["dim2_items"] = dim2_items
        result["total_score"] = score
        return result

    except Exception as e:
        result["status"] = "error"
        result["error"] = "%s: %s" % (type(e).__name__, e)
        return result


if __name__ == "__main__":
    import json

    _arg = sys.argv[1] if len(sys.argv) > 1 else os.path.dirname(os.path.abspath(__file__))
    print(json.dumps(evaluate(_arg), ensure_ascii=False, indent=2))
