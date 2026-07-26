# -*- coding: utf-8 -*-
"""
自动评估脚本：对“培训记录表_按图制作_可编辑.docx”按照打分细则进行自动评分。

设计说明
--------
评分分为两个维度：
  维度1（可用与可修改性）：作为“准入门槛”。不满足 -> 直接 0 分，且不再检查维度2。
  维度2（完成度评分细则）：在维度1通过的前提下逐条检查。
      - 得分点（+N）：必须满足该条细则中的“每一个”子条件，才累计 +N。
      - 扣分点（-N）：只要满足该条细则中的“任意一个”子条件，就累计 -N。

本脚本仅依赖 Python 标准库（zipfile / xml），直接解析 docx 内部的
word/document.xml，无需安装 python-docx。

对于不便严格量化的细则（如“无文字重叠/无裁切”“边框无多余斜线”），
采用结构化的可计算近似（如行高>0、跨页判断、有无明确的斜线边框节点等）
以贴近评估意图。
"""

import os
import re
import sys
import json
import zipfile
import xml.etree.ElementTree as ET

SCRIPT_ID = "029"

# ---------------------------------------------------------------------------
# 基础常量
# ---------------------------------------------------------------------------
W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W = "{%s}" % W_NS

DXA_PER_CM = 567.0          # 1 cm = 567 dxa（二十分之一磅）
WIDTH_TOL_CM = 0.05         # 单格宽度允许偏差
SZ_XIAOSAN = 30             # 小三 = 15pt，OOXML 中 w:sz 以半磅计 -> 30
FANGSONG_KEYWORDS = ("仿宋",)   # 仿宋字体匹配关键字


def cm(dxa):
    """dxa -> cm"""
    return dxa / DXA_PER_CM


def approx(value_cm, target_cm, tol=WIDTH_TOL_CM):
    return abs(value_cm - target_cm) <= tol + 1e-9


# ---------------------------------------------------------------------------
# 文档解析
# ---------------------------------------------------------------------------
class DocModel:
    """从 docx 中解析出评估所需的结构化数据。"""

    def __init__(self, path):
        self.path = path
        self.ok_zip = False
        self.has_document = False
        self.parse_error = None
        self.tables = []           # 每个表: {'grid':[dxa...], 'rows':[...]}
        self.all_runs_fonts = []   # [(font_or_None, sz_or_None, text)]
        self.has_image = False
        self.border_info = {}
        self._load(path)

    def _load(self, path):
        if not os.path.isfile(path):
            self.parse_error = "文件不存在"
            return
        ext = os.path.splitext(path)[1].lower()
        if ext != ".docx":
            self.parse_error = "扩展名非 .docx"
            return
        try:
            zf = zipfile.ZipFile(path)
        except zipfile.BadZipFile:
            self.parse_error = "非合法 docx(zip) 结构，可能是损坏文件"
            return
        self.ok_zip = True
        names = zf.namelist()
        if "word/document.xml" not in names:
            self.parse_error = "缺少 word/document.xml"
            zf.close()
            return
        self.has_document = True
        # 检测是否使用图片代替表格（drawing / pict / OLE）
        try:
            doc_xml = zf.read("word/document.xml").decode("utf-8", "ignore")
        except Exception as e:
            self.parse_error = "无法读取 document.xml: %s" % e
            zf.close()
            return
        # 图片/对象嵌入检测
        media_imgs = [n for n in names if n.startswith("word/media/")]
        if re.search(r"<w:drawing|<w:pict|<o:OLEObject|<v:imagedata", doc_xml):
            self.has_image = True
        if media_imgs:
            self.has_image = True
        zf.close()

        # 解析 XML
        try:
            root = ET.fromstring(doc_xml)
        except ET.ParseError as e:
            self.parse_error = "document.xml 解析失败: %s" % e
            return

        body = root.find(W + "body")
        if body is None:
            self.parse_error = "缺少 body"
            return

        for tbl in body.iter(W + "tbl"):
            self.tables.append(self._parse_table(tbl))

        # 收集全部文字 run 的字体与字号
        for r in body.iter(W + "r"):
            text = "".join(t.text or "" for t in r.iter(W + "t"))
            rpr = r.find(W + "rPr")
            font = None
            sz = None
            if rpr is not None:
                rfonts = rpr.find(W + "rFonts")
                if rfonts is not None:
                    font = (rfonts.get(W + "eastAsia")
                            or rfonts.get(W + "ascii")
                            or rfonts.get(W + "hAnsi"))
                szel = rpr.find(W + "sz")
                if szel is not None:
                    try:
                        sz = int(szel.get(W + "val"))
                    except (TypeError, ValueError):
                        sz = None
            self.all_runs_fonts.append((font, sz, text))

        # 边框信息（取第一个表的 tblBorders）
        self.border_info = self._parse_borders(body)

    def _parse_table(self, tbl):
        grid = []
        tblgrid = tbl.find(W + "tblGrid")
        if tblgrid is not None:
            for gc in tblgrid.findall(W + "gridCol"):
                try:
                    grid.append(int(gc.get(W + "w")))
                except (TypeError, ValueError):
                    grid.append(0)

        rows = []
        for tr in tbl.findall(W + "tr"):
            trh = tr.find(W + "trPr")
            height = None
            if trh is not None:
                he = trh.find(W + "trHeight")
                if he is not None:
                    try:
                        height = int(he.get(W + "val"))
                    except (TypeError, ValueError):
                        height = None
            cells = []
            for tc in tr.findall(W + "tc"):
                tcpr = tc.find(W + "tcPr")
                w = None
                gridspan = 1
                vmerge = None
                if tcpr is not None:
                    tcw = tcpr.find(W + "tcW")
                    if tcw is not None:
                        try:
                            w = int(tcw.get(W + "w"))
                        except (TypeError, ValueError):
                            w = None
                    gs = tcpr.find(W + "gridSpan")
                    if gs is not None:
                        try:
                            gridspan = int(gs.get(W + "val"))
                        except (TypeError, ValueError):
                            gridspan = 1
                    vm = tcpr.find(W + "vMerge")
                    if vm is not None:
                        vmerge = vm.get(W + "val") or "continue"
                text = "".join(t.text or "" for t in tc.iter(W + "t"))
                cells.append({
                    "w": w,
                    "gridspan": gridspan,
                    "vmerge": vmerge,
                    "text": text.strip(),
                })
            rows.append({"height": height, "cells": cells})
        return {"grid": grid, "rows": rows}

    def _parse_borders(self, body):
        info = {
            "found": False,
            "all_single": True,       # 是否全部为单实线（无重线/双线）
            "colors": set(),
            "has_diagonal": False,    # 是否存在斜线边框
            "missing_or_none": False, # 是否存在 none/nil 边线
            "edges_present": set(),   # 已定义且非 none 的边：top/left/bottom/right/insideH/insideV
            "max_sz": 0,              # 最大线宽(eighths of a point)，用于“细实线”判定
            "has_double": False,      # 是否存在 double/重线类样式
        }
        tbl = body.find(W + "tbl")
        if tbl is None:
            return info
        tblpr = tbl.find(W + "tblPr")
        if tblpr is None:
            return info
        borders = tblpr.find(W + "tblBorders")
        if borders is None:
            return info
        info["found"] = True
        for edge in borders:
            tag = edge.tag.replace(W, "")
            val = edge.get(W + "val")
            color = edge.get(W + "color")
            sz = edge.get(W + "sz")
            if tag in ("tl2br", "tr2bl"):
                info["has_diagonal"] = True
            if val in (None, "none", "nil"):
                info["missing_or_none"] = True
            else:
                if tag in ("top", "left", "bottom", "right", "insideH", "insideV"):
                    info["edges_present"].add(tag)
                if val != "single":
                    info["all_single"] = False
                if val in ("double", "triple", "thick", "thickThinSmallGap",
                           "thinThickSmallGap", "dashDotStroked"):
                    info["has_double"] = True
            if sz:
                try:
                    info["max_sz"] = max(info["max_sz"], int(sz))
                except (TypeError, ValueError):
                    pass
            if color:
                info["colors"].add(color.upper())
        # 检测单元格级别是否存在斜线边框
        for tcborders in body.iter(W + "tcBorders"):
            for edge in tcborders:
                tag = edge.tag.replace(W, "")
                if tag in ("tl2br", "tr2bl"):
                    v = edge.get(W + "val")
                    if v not in (None, "none", "nil"):
                        info["has_diagonal"] = True
        return info

    # ----- 派生：把每行的单元格映射到“逻辑列宽” -----
    def effective_rows(self, table_index=0):
        """
        返回每行的有效单元格宽度（cm）列表，宽度按 gridSpan 累加 tblGrid 对应列宽得到，
        以便和细则中“合并后”的逻辑宽度对应。
        """
        if table_index >= len(self.tables):
            return []
        table = self.tables[table_index]
        grid = table["grid"]
        result = []
        for row in table["rows"]:
            col_ptr = 0
            cell_widths = []
            for c in row["cells"]:
                span = c["gridspan"] or 1
                # 优先用 tcW；若缺失再用 grid 累加
                if c["w"] is not None:
                    w_dxa = c["w"]
                else:
                    w_dxa = sum(grid[col_ptr:col_ptr + span]) if grid else 0
                cell_widths.append({
                    "w_cm": cm(w_dxa),
                    "text": c["text"],
                    "span": span,
                    "vmerge": c["vmerge"],
                })
                col_ptr += span
            result.append({"height": row["height"], "cells": cell_widths,
                           "total_cm": sum(x["w_cm"] for x in cell_widths)})
        return result


# ---------------------------------------------------------------------------
# 维度1：可用与可修改性（准入门槛）
# ---------------------------------------------------------------------------
def check_dimension1(doc):
    """
    返回 (passed: bool, details: list[str])

    按用户要求，已删除以下维度一检查：
      - "表格为 Word 原生可编辑表格，文字/单元格/边框/合并关系均可单独修改，
        不能使用图片或截图代替"
      - "文档包含完整表格，不出现表格跨页/行列错位/单元格严重变形/文字大面积
        重叠或内容被裁切"
      - "若文件格式错误、表格整体不可编辑或表格结构严重混乱，则维度1得分为0"

    仅保留：文件格式为 .docx 且能被正常打开（zip 可解析、含 word/document.xml）。
    """
    details = []
    fatal = False

    if not doc.ok_zip or not doc.has_document:
        details.append("× 文件格式错误或无法打开：%s" % (doc.parse_error or "未知"))
        fatal = True
    else:
        details.append("√ 文件为 .docx 且可正常解析打开（含 word/document.xml）")

    passed = not fatal
    return passed, details


# ---------------------------------------------------------------------------
# 维度2：完成度评分细则
# ---------------------------------------------------------------------------
# 每行第1列文本期望（去换行后的标签）
ROW_LABELS = ["员工姓名", "培训日期", "培训目标", "考核结果",
              "评定方式", "学习完成", "直属主管", "人力资源"]


def _norm(text):
    return re.sub(r"\s+", "", text or "")


def check_dimension2(doc):
    """逐条评估，返回 (total_score, hit_list)。
    hit_list: [(rule_desc, points, passed_bool, reason)]
    """
    rows = doc.effective_rows(0)
    hits = []

    def add(desc, points, passed, reason):
        hits.append({"desc": desc, "points": points,
                     "passed": passed, "reason": reason})

    # ---- +5：整体结构 8 行，依次为：基本信息、培训日期、培训信息、考核信息、
    #          评定方式、学习完成情况、直属主管意见、人力资源意见 ----
    # 细则给出的是 8 个“区块名称”。本模板中每个区块由对应行的首列标签体现，
    # 这里把细则的区块名称映射到用于识别该行的首列标签（按图制作的实际标签）。
    expect_struct = [
        ("基本信息", "员工姓名"),
        ("培训日期", "培训日期"),
        ("培训信息", "培训目标"),
        ("考核信息", "考核结果"),
        ("评定方式", "评定方式"),
        ("学习完成情况", "学习完成"),
        ("直属主管意见", "直属主管"),
        ("人力资源意见", "人力资源"),
    ]
    ok_struct = len(rows) == 8
    matched = []
    if ok_struct:
        for i, (section, label) in enumerate(expect_struct):
            first_text = _norm(rows[i]["cells"][0]["text"]) if rows[i]["cells"] else ""
            hit = label in first_text
            matched.append("行%d[%s]:%s" % (i + 1, section,
                                            "√" if hit else "×(%s)" % first_text))
            if not hit:
                ok_struct = False
    add("+5 表格整体结构：自上而下包含8行，依次为基本信息、培训日期、培训信息、考核信息、评定方式、学习完成情况、直属主管意见和人力资源意见",
        5, ok_struct,
        ("行数=%d；逐行区块：%s" % (len(rows), " ".join(matched))) if matched
        else "行数=%d，未能逐行匹配" % len(rows))

    # ---- +5：第1行 8 个基础单元格 + 合并 + 宽度 + 文字 ----
    # ---- +5：第1行基础网格 ----
    # 细则要点（逐条踩点）：
    #   (1) 包含 8 个基础单元格，并按图片效果设置必要的单元格合并关系；
    #   (2) 从左至右 8 个基础单元格宽度依次为
    #       2.4 / 2.4 / 1.45 / 1.2 / 1.45 / 2.4 / 2.4 / 2.4 厘米，单格偏差 <= 0.05cm；
    #   (3) 文字包含“员工姓名”“所属部门”“性别”“任职岗位”，且分别位于对应信息区域
    #       （即第 1、3、5、7 个基础单元格）；
    #   (4) 右侧第 2、4、6、8 个基础单元格保留可填写空白单元格。
    #
    # 说明：“8 个基础单元格”指的是合并前的基础网格。本模板按图片做了合并，
    # 第1行可见单元格数少于 8，但 8 个基础单元格的边界依然由 tblGrid 决定。
    # 这里以“2.4cm 累加边界”从 tblGrid 还原出 8 个基础单元格，并将每个基础格的
    # 文字归到它所落入的可见(合并)单元格上。
    r1_ok = False
    r1_reason = []
    if len(rows) >= 1 and doc.tables and doc.tables[0]["grid"]:
        c = rows[0]["cells"]
        expect_labels = ["员工姓名", "所属部门", "性别", "任职岗位"]
        expect_w = [2.4, 2.4, 1.45, 1.2, 1.45, 2.4, 2.4, 2.4]

        # 由细则宽度推出 8 个基础单元格的“目标累计右边界”
        target_bounds = []
        acc = 0.0
        for w in expect_w:
            acc += w
            target_bounds.append(acc)
        target_left = [0.0] + target_bounds[:-1]

        # tblGrid 的累计边界（cm）
        grid = doc.tables[0]["grid"]
        grid_bounds = []
        acc = 0.0
        for g in grid:
            acc += cm(g)
            grid_bounds.append(round(acc, 3))

        # 校验：每个基础格右边界都能在 grid 边界中找到匹配 -> 说明 8 基础格存在且宽度达标
        def near_any(val, arr, tol=WIDTH_TOL_CM):
            return any(abs(val - a) <= tol + 1e-9 for a in arr)

        cond_w = all(near_any(b, [0.0] + grid_bounds) for b in target_bounds)
        r1_reason.append("基础格目标右边界%s 命中grid边界=%s"
                         % ([round(b, 2) for b in target_bounds], cond_w))

        # 可见单元格 -> [左边界,右边界] 区间，用于把基础格的文字定位到对应可见格
        vis_spans = []
        acc = 0.0
        for vc in c:
            left = acc
            acc += vc["w_cm"]
            vis_spans.append((left, acc, _norm(vc["text"])))

        def text_at(center):
            for left, right, txt in vis_spans:
                if left - 1e-6 <= center <= right + 1e-6:
                    return txt
            return ""

        # 取每个基础格中心点所在可见格的文字
        base_texts = []
        for k in range(8):
            center = (target_left[k] + target_bounds[k]) / 2.0
            base_texts.append(text_at(center))

        cond_count = (len(grid) >= 8 and cond_w)   # 基础网格存在且达标
        # (3) 标签位于 1/3/5/7（基础格序号 0/2/4/6）
        cond_label = all(expect_labels[k] in base_texts[2 * k] for k in range(4))
        # (4) 右侧 2/4/6/8（基础格序号 1/3/5/7）为可填写空白
        cond_blank = all(base_texts[2 * k + 1] == "" for k in range(4))
        # (1) 设置了合并关系：可见单元格数 < 8
        cond_merge = (len(c) < 8)

        r1_reason.append("基础格文字%s" % base_texts)
        r1_reason.append("标签位于1/3/5/7=%s 空白位于2/4/6/8=%s 含合并=%s(可见%d格)"
                         % (cond_label, cond_blank, cond_merge, len(c)))

        r1_ok = cond_count and cond_w and cond_label and cond_blank and cond_merge
    add("+5 第1行基础网格：8个基础单元格+按图必要合并；从左至右宽度2.4/2.4/1.45/1.2/1.45/2.4/2.4/2.4cm(±0.05)；含员工姓名/所属部门/性别/任职岗位且分别位于1/3/5/7区域，右侧2/4/6/8保留空白",
        5, r1_ok, "；".join(r1_reason))

    # ---- +1：第2行单元格 ----
    # 细则要点（逐条踩点）：
    #   (1) 左侧为“培训日期”单元格；
    #   (2) 右侧 2-8 单元格横向合并，且延伸至表格右边线；
    #   (3) 第2行第1列宽度为 2.4cm，偏差 <= 0.05cm。
    r2_ok = False
    r2_reason = []
    if len(rows) >= 2 and doc.tables and doc.tables[0]["grid"]:
        c = rows[1]["cells"]
        table_total_cm = sum(cm(g) for g in doc.tables[0]["grid"])  # 表格右边线
        # (1) 左侧“培训日期”
        cond_label = bool(c) and "培训日期" in _norm(c[0]["text"])
        # (3) 第1列宽度 2.4cm
        cond_w = bool(c) and approx(c[0]["w_cm"], 2.4)
        # (2) 右侧 2-8 横向合并为一个单元格：第2行除首列外只剩 1 个可见单元格
        cond_merge = (len(c) == 2)
        # (2) 延伸至表格右边线：首列右边界 + 合并格宽度 == 表格总宽（即抵达右边线）
        cond_extend = cond_merge and approx(rows[1]["total_cm"], table_total_cm, tol=0.05)
        r2_reason.append("首列宽=%.2f(=2.4) label=%s 右侧合并为1格=%s 行总宽=%.2f(表格右边线=%.2f)"
                         % (c[0]["w_cm"] if c else 0, cond_label, cond_merge,
                            rows[1]["total_cm"], table_total_cm))
        r2_ok = cond_label and cond_w and cond_merge and cond_extend
    add("+1 第2行：左侧‘培训日期’单元格，右侧2-8横向合并并延伸至表格右边线；第1列宽度2.4cm(±0.05)",
        1, r2_ok, "；".join(r2_reason))

    # ---- +3：第3行单元格 ----
    r3_ok = False
    # 细则要点（逐条踩点）：
    #   (1) 共 6 格；
    #   (2) 按“培训目标”—填写区—“培训项目”—填写区—“学习时长”—填写区 的顺序排列
    #       （即第 1/3/5 格为标签，第 2/4/6 格为填写区）；
    #   (3) 从左至右宽度依次为 2.4/2.6/2.85/2.85/2.65/2.65cm，单格偏差 <= 0.05cm。
    r3_ok = False
    r3_reason = []
    if len(rows) >= 3:
        c = rows[2]["cells"]
        expect_w = [2.4, 2.6, 2.85, 2.85, 2.65, 2.65]
        # (1) 共6格
        cond_count = (len(c) == 6)
        cond_label = False
        cond_w = False
        if cond_count:
            # (2) 顺序：1/3/5 为对应标签，2/4/6 为填写区(空白)
            cond_label = ("培训目标" in _norm(c[0]["text"])
                          and "培训项目" in _norm(c[2]["text"])
                          and "学习时长" in _norm(c[4]["text"])
                          and _norm(c[1]["text"]) == ""
                          and _norm(c[3]["text"]) == ""
                          and _norm(c[5]["text"]) == "")
            # (3) 逐格宽度
            widths = [round(x["w_cm"], 3) for x in c]
            bad = [("第%d格%.2f≠%.2f" % (k + 1, widths[k], expect_w[k]))
                   for k in range(6) if not approx(widths[k], expect_w[k])]
            cond_w = (len(bad) == 0)
            r3_reason.append("宽度cm=%s" % widths)
            if bad:
                r3_reason.append("宽度偏差:%s" % ";".join(bad))
        r3_reason.insert(0, "格数=%d(期望6) 顺序‘培训目标-填写-培训项目-填写-学习时长-填写’=%s"
                         % (len(c), cond_label))
        r3_ok = cond_count and cond_label and cond_w
    add("+3 第3行：共6格，按‘培训目标-填写区-培训项目-填写区-学习时长-填写区’顺序排列，从左至右宽度2.4/2.6/2.85/2.85/2.65/2.65cm(±0.05)",
        3, r3_ok, "；".join(r3_reason))

    # ---- +3：第4行单元格 ----
    # 细则要点（逐条踩点）：
    #   (1) 按“考核结果—填写区—能力评定—填写区”的结构排列（4 格：
    #       第1格标签‘考核结果’、第2格填写区、第3格标签‘能力评定’、第4格填写区）；
    #   (2) 最后一个填写区横向延伸至表格右边线；
    #   (3) 从左至右宽度依次为 2.4/2.6/2.85cm（细则只给出前 3 格宽度），
    #       单格偏差 <= 0.05cm。
    r4_ok = False
    r4_reason = []
    if len(rows) >= 4 and doc.tables and doc.tables[0]["grid"]:
        c = rows[3]["cells"]
        expect_w = [2.4, 2.6, 2.85]
        table_total_cm = sum(cm(g) for g in doc.tables[0]["grid"])  # 表格右边线
        # (1) 结构 4 格，标签/填写区交替
        cond_count = (len(c) == 4)
        cond_label = False
        cond_w = False
        cond_extend = False
        if cond_count:
            cond_label = ("考核结果" in _norm(c[0]["text"])
                          and _norm(c[1]["text"]) == ""
                          and "能力评定" in _norm(c[2]["text"])
                          and _norm(c[3]["text"]) == "")
            # (3) 前 3 格宽度
            widths = [round(x["w_cm"], 3) for x in c[:3]]
            bad = [("第%d格%.2f≠%.2f" % (k + 1, widths[k], expect_w[k]))
                   for k in range(3) if not approx(widths[k], expect_w[k])]
            cond_w = (len(bad) == 0)
            # (2) 最后填写区延伸至表格右边线
            cond_extend = approx(rows[3]["total_cm"], table_total_cm, tol=0.05)
            r4_reason.append("前3格宽cm=%s 行总宽=%.2f(表格右边线=%.2f 末填写区延伸=%s)"
                             % (widths, rows[3]["total_cm"], table_total_cm, cond_extend))
            if bad:
                r4_reason.append("宽度偏差:%s" % ";".join(bad))
        r4_reason.insert(0, "格数=%d(期望4) 结构‘考核结果-填写-能力评定-填写’=%s"
                         % (len(c), cond_label))
        r4_ok = cond_count and cond_label and cond_w and cond_extend
    add("+3 第4行：按‘考核结果-填写区-能力评定-填写区’结构排列，最后填写区横向延伸至表格右边线，从左至右宽度2.4/2.6/2.85cm(±0.05)",
        3, r4_ok, "；".join(r4_reason))

    r5_ok = False
    r5_r = ""
    # ---- +1：第5行单元格 ----
    # 细则要点（逐条踩点）：
    #   (1) 左侧显示“评定方式”；
    #   (2) 第1列(评定方式)宽度为 2.4cm；
    #   (3) 右侧 2-8 格为横向合并的空白填写区。
    if len(rows) >= 5 and doc.tables and doc.tables[0]["grid"]:
        c = rows[4]["cells"]
        cond_label = bool(c) and "评定方式" in _norm(c[0]["text"])          # (1)
        cond_w = bool(c) and approx(c[0]["w_cm"], 2.4)                       # (2)
        cond_merge = (len(c) == 2)                                          # (3) 右侧2-8横向合并为一格
        cond_blank = cond_merge and _norm(c[1]["text"]) == ""               # (3) 空白填写区
        r5_r = ("首列宽=%.2f(=2.4) label=%s 右侧2-8横向合并为1格=%s 空白填写区=%s"
                % (c[0]["w_cm"] if c else 0, cond_label, cond_merge,
                   cond_blank if cond_merge else False))
        r5_ok = cond_label and cond_w and cond_merge and cond_blank
    add("+1 第5行：左侧显示‘评定方式’(宽度2.4cm)，右侧2-8格为横向合并的空白填写区", 1, r5_ok, r5_r)

    # ---- +1：第6行单元格 ----
    # 细则要点（逐条踩点）：
    #   (1) 左侧显示“学习完成情况”；
    #   (2) 第1列宽度为 2.4cm；
    #   (3) 右侧 2-8 格为“面积较大”的空白填写区。
    # 说明：“面积较大”无量化阈值，这里以可计算的近似判定——该行填写区行高明显
    # 大于普通标签行（取前5个标签行的最大行高作为基准，需更大）。
    r6_ok = False
    r6_r = ""
    if len(rows) >= 6:
        c = rows[5]["cells"]
        base_h = max((rows[k]["height"] or 0) for k in range(5))   # 普通行基准高度
        h6 = rows[5]["height"] or 0
        cond_label = bool(c) and "学习完成" in _norm(c[0]["text"])         # (1)
        cond_w = bool(c) and approx(c[0]["w_cm"], 2.4)                      # (2)
        cond_merge = (len(c) == 2)                                         # 右侧2-8合并为一格
        cond_blank = cond_merge and _norm(c[1]["text"]) == ""              # (3) 空白
        cond_big = h6 > base_h                                             # (3) 面积较大
        r6_r = ("首列宽=%.2f(=2.4) label=%s 右侧合并为1格=%s 空白=%s 行高=%s(面积较大需>普通行%s)"
                % (c[0]["w_cm"] if c else 0, cond_label, cond_merge,
                   cond_blank if cond_merge else False, h6, base_h))
        r6_ok = cond_label and cond_w and cond_merge and cond_blank and cond_big
    add("+1 第6行：左侧显示‘学习完成情况’(宽度2.4cm)，右侧2-8格为面积较大的空白填写区", 1, r6_ok, r6_r)

    # ---- +1：第7行单元格 ----
    # 细则要点（逐条踩点）：
    #   (1) 左侧显示“直属主管意见”；
    #   (2) 第1列宽度为 2.4cm；
    #   (3) 右侧 2-8 格为横向合并的大面积空白填写区。
    # 说明：“大面积”无量化阈值，沿用与第6行一致的近似——填写区行高大于普通标签行。
    r7_ok = False
    r7_r = ""
    if len(rows) >= 7:
        c = rows[6]["cells"]
        base_h = max((rows[k]["height"] or 0) for k in range(5))   # 普通行基准高度
        h7 = rows[6]["height"] or 0
        cond_label = bool(c) and "直属主管意见" in _norm(c[0]["text"])     # (1)
        cond_w = bool(c) and approx(c[0]["w_cm"], 2.4)                      # (2)
        cond_merge = (len(c) == 2)                                         # (3) 右侧2-8横向合并为一格
        cond_blank = cond_merge and _norm(c[1]["text"]) == ""              # (3) 空白
        cond_big = h7 > base_h                                             # (3) 大面积
        r7_r = ("首列宽=%.2f(=2.4) label=%s 右侧2-8横向合并为1格=%s 空白=%s 行高=%s(大面积需>普通行%s)"
                % (c[0]["w_cm"] if c else 0, cond_label, cond_merge,
                   cond_blank if cond_merge else False, h7, base_h))
        r7_ok = cond_label and cond_w and cond_merge and cond_blank and cond_big
    add("+1 第7行：左侧显示‘直属主管意见’(宽度2.4cm)，右侧2-8格为横向合并的大面积空白填写区", 1, r7_ok, r7_r)

    # ---- +1：第8行单元格 ----
    # 细则要点（逐条踩点）：
    #   (1) 左侧显示“人力资源意见”；
    #   (2) 第1列宽度为 2.4cm；
    #   (3) 右侧 2-8 格为横向合并的大面积空白填写区。
    # 说明：“大面积”无量化阈值，沿用与第6/7行一致的近似——填写区行高大于普通标签行。
    r8_ok = False
    r8_r = ""
    if len(rows) >= 8:
        c = rows[7]["cells"]
        base_h = max((rows[k]["height"] or 0) for k in range(5))   # 普通行基准高度
        h8 = rows[7]["height"] or 0
        cond_label = bool(c) and "人力资源意见" in _norm(c[0]["text"])     # (1)
        cond_w = bool(c) and approx(c[0]["w_cm"], 2.4)                      # (2)
        cond_merge = (len(c) == 2)                                         # (3) 右侧2-8横向合并为一格
        cond_blank = cond_merge and _norm(c[1]["text"]) == ""              # (3) 空白
        cond_big = h8 > base_h                                             # (3) 大面积
        r8_r = ("首列宽=%.2f(=2.4) label=%s 右侧2-8横向合并为1格=%s 空白=%s 行高=%s(大面积需>普通行%s)"
                % (c[0]["w_cm"] if c else 0, cond_label, cond_merge,
                   cond_blank if cond_merge else False, h8, base_h))
        r8_ok = cond_label and cond_w and cond_merge and cond_blank and cond_big
    add("+1 第8行：左侧显示‘人力资源意见’(宽度2.4cm)，右侧2-8格为横向合并的大面积空白填写区", 1, r8_ok, r8_r)

    # ---- +1：表格全部文字——中文字体统一为仿宋 ----
    # 细则要点（逐条踩点）：
    #   (1) 评估对象为“表格全部文字”中的中文字符；
    #   (2) 中文字体统一为仿宋（即所有含中文的文字片段，其中文字体均为仿宋，
    #       不混用其它字体）。
    # 说明：中文字符在 OOXML 中由 rFonts 的 eastAsia 属性决定字体，因此以每个
    # 含中文的 run 的 eastAsia 字体（缺省时回退 ascii/hAnsi）是否为“仿宋”来判定。
    cn_runs = [(f, s, t) for (f, s, t) in doc.all_runs_fonts
               if re.search(r"[一-鿿]", t or "")]
    if cn_runs:
        # 收集中文文字使用到的全部字体（用于“统一”判定）
        used_fonts = set(f for (f, s, t) in cn_runs)
        non_fs = [(f, t) for (f, s, t) in cn_runs
                  if not (f and any(k in f for k in FANGSONG_KEYWORDS))]
        # (2) 统一为仿宋：无任何中文 run 使用非仿宋字体
        font_ok = (len(non_fs) == 0)
        font_reason = ("含中文文字片段共%d个，使用字体集合=%s，非仿宋%d个"
                       % (len(cn_runs), sorted(x for x in used_fonts if x), len(non_fs)))
        if non_fs:
            font_reason += "：例%s" % str(non_fs[:3])
    else:
        font_ok = False
        font_reason = "未检测到中文文字"
    add("+1 表格全部文字：中文字体统一为仿宋", 1, font_ok, font_reason)

    # ---- +1：表格全部文字——字号统一为小三，不混用其他字号 ----
    # 细则要点（逐条踩点）：
    #   (1) 评估对象为“表格全部文字”；
    #   (2) 字号统一为小三（小三 = 15pt，OOXML w:sz 以半磅计 => 30）；
    #   (3) 不混用其他字号（即所有文字片段字号都只能是小三这一种）。
    # 说明：w:sz=30 即小三。逐个有文字的 run 取其字号；若某 run 未显式声明字号
    # （继承样式），记为 None，单独标识，避免漏判“混用”。
    sz_runs = [(f, s, t) for (f, s, t) in doc.all_runs_fonts if (t or "").strip()]
    explicit_sizes = set(s for (f, s, t) in sz_runs if s is not None)
    n_missing = sum(1 for (f, s, t) in sz_runs if s is None)
    # (2)(3) 统一为小三且不混用：所有文字 run 的显式字号都为 30，且无未声明字号的 run
    sz_ok = (len(sz_runs) > 0
             and explicit_sizes == {SZ_XIAOSAN}
             and n_missing == 0)
    sz_reason = ("文字片段共%d个，字号集合(半磅)=%s(小三=30)，未显式声明字号的片段=%d个"
                 % (len(sz_runs),
                    sorted(explicit_sizes) if explicit_sizes else "无",
                    n_missing))
    add("+1 表格全部文字：字号统一为小三，不混用其他字号", 1, sz_ok, sz_reason)

    # ---- +1 边框：黑/深灰细实线，连续清晰，无缺线/重线/多余斜线 ----
    # ---- +1：表格边框 ----
    # 细则要点（逐条踩点）：
    #   (1) 使用黑色或深灰色；
    #   (2) 为细实线；
    #   (3) 内外边框连续清晰（内外边框俱全：上/下/左/右四条外框 + 内部横/纵线，
    #       且均无 none/nil 缺线）；
    #   (4) 无缺线；
    #   (5) 无重线（无 double 等多重线）；
    #   (6) 无多余斜线。
    b = doc.border_info
    border_ok = False
    if b.get("found"):
        colors = b["colors"]

        # (1) 黑色或深灰色（亮度较低；auto 视为黑）
        def is_dark(hexc):
            if hexc in ("AUTO",):
                return True
            try:
                r = int(hexc[0:2], 16); g = int(hexc[2:4], 16); bl = int(hexc[4:6], 16)
                return (r + g + bl) / 3 <= 128
            except Exception:
                return False
        cond_color = bool(colors) and all(is_dark(c) for c in colors)

        # (2) 细实线：全部为 single，且线宽不过粗（sz 以 1/8 磅计，<=12 即 <=1.5pt 视为细线）
        cond_single = b["all_single"]
        cond_thin = (0 < b["max_sz"] <= 12)
        cond_fine_solid = cond_single and cond_thin

        # (3) 内外边框连续清晰：六类边线齐全
        need_edges = {"top", "left", "bottom", "right", "insideH", "insideV"}
        cond_continuous = need_edges.issubset(b["edges_present"])

        # (4) 无缺线
        cond_no_missing = not b["missing_or_none"]
        # (5) 无重线
        cond_no_double = not b["has_double"]
        # (6) 无多余斜线
        cond_no_diag = not b["has_diagonal"]

        border_ok = (cond_color and cond_fine_solid and cond_continuous
                     and cond_no_missing and cond_no_double and cond_no_diag)
        border_reason = ("颜色=%s(黑/深灰=%s) 细实线(全single=%s 线宽=%d/8pt细线=%s) "
                         "内外边框齐全=%s(已定义:%s) 无缺线=%s 无重线=%s 无斜线=%s"
                         % (sorted(colors), cond_color, cond_single, b["max_sz"],
                            cond_thin, cond_continuous,
                            sorted(b["edges_present"]), cond_no_missing,
                            cond_no_double, cond_no_diag))
    else:
        border_reason = "未找到表格 tblBorders 定义"
    add("+1 表格边框：使用黑色或深灰色细实线，内外边框连续清晰，无缺线、重线或多余斜线",
        1, border_ok, border_reason)

    total = sum(h["points"] for h in hits if h["passed"])
    return total, hits


# ---------------------------------------------------------------------------
# 主流程
# ---------------------------------------------------------------------------
def _find_target_doc(dir_path: str) -> str | None:
    """在 dir_path 目录下定位被评估的 .docx 文档（排除脚本自身等非文档文件）。"""
    candidates: list[str] = []
    for name in os.listdir(dir_path):
        ext = os.path.splitext(name)[1].lower()
        if ext == ".docx" and not name.startswith("~$"):
            candidates.append(name)
    if not candidates:
        return None
    # 若同目录下存在多个文档，优先选择非临时文件中修改时间最新的一个
    candidates.sort(key=lambda n: os.path.getmtime(os.path.join(dir_path, n)),
                    reverse=True)
    return os.path.join(dir_path, candidates[0])


def evaluate(dir_path: str) -> dict:
    """统一入口：接收脚本所在目录的路径，自行在该目录内定位并评估目标文档。

    返回结构化 dict，字段含义见项目约定文档 §2.2：
    id / file_name / status / error / dim1_pass / dim1_reason /
    dim2_items / total_score / max_score
    """
    try:
        file_path = _find_target_doc(dir_path)
        if file_path is None:
            return {
                "id": SCRIPT_ID,
                "file_name": None,
                "status": "error",
                "error": "目录中未找到 .docx 文档：%s" % dir_path,
                "dim1_pass": False,
                "dim1_reason": "未找到被评估文档",
                "dim2_items": [],
                "total_score": 0,
                "max_score": 0,
            }

        file_name = os.path.basename(file_path)
        doc = DocModel(file_path)

        # 维度1（准入门槛）
        d1_pass, d1_details = check_dimension1(doc)

        if not d1_pass:
            return {
                "id": SCRIPT_ID,
                "file_name": file_name,
                "status": "ok",
                "error": None,
                "dim1_pass": False,
                "dim1_reason": "；".join(d1_details),
                "dim2_items": [],
                "total_score": 0,
                "max_score": 0,
            }

        # 维度1通过 -> 评估维度2
        total, hits = check_dimension2(doc)
        dim2_items = []
        max_score = 0
        for h in hits:
            # desc 形如 "+5 表格整体结构：..."，去掉开头的“+N ”得到细则内容
            rule = re.sub(r"^[+\-]\d+\s*", "", h["desc"])
            dim2_items.append({
                "rule": rule,
                "max_delta": h["points"],
                "delta": h["points"] if h["passed"] else 0,
                "hit": h["passed"],
                "detail": "",
            })
            max_score += h["points"]

        return {
            "id": SCRIPT_ID,
            "file_name": file_name,
            "status": "ok",
            "error": None,
            "dim1_pass": True,
            "dim1_reason": "",
            "dim2_items": dim2_items,
            "total_score": total,
            "max_score": max_score,
        }
    except Exception as e:
        return {
            "id": SCRIPT_ID,
            "file_name": None,
            "status": "error",
            "error": str(e),
            "dim1_pass": False,
            "dim1_reason": "",
            "dim2_items": [],
            "total_score": 0,
            "max_score": 0,
        }


if __name__ == "__main__":
    # 仅用于本地调试：传入脚本所在目录路径（默认取脚本自身所在目录）
    _dir = sys.argv[1] if len(sys.argv) > 1 else os.path.dirname(os.path.abspath(__file__))
    print(json.dumps(evaluate(_dir), ensure_ascii=False, indent=2))
