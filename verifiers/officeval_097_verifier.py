#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# pyright: reportUnknownMemberType=false, reportUnknownVariableType=false, reportUnknownArgumentType=false, reportUnknownParameterType=false, reportMissingParameterType=false, reportPrivateUsage=false, reportUnknownLambdaType=false, reportGeneralTypeIssues=false

import os
import re
import sys
import json
import zipfile
import warnings

warnings.filterwarnings("ignore")
import openpyxl

# 说明：按接口统一约定，本模块不修改全局 sys.stdout，也不 print 主结果。
# 主结果统一由 evaluate(dir_path) -> dict 返回；入参为"脚本所在目录的路径"，
# 由脚本自己在该目录内定位并打开被评估的 .xlsx（忽略以 ~$ 开头的临时锁文件）。
# 本文件仅在 __main__ 下作为本地调试入口，才把 dict 序列化为 JSON 打印。
#
# 【关键前提】原始文件的"小计"数据为文本类型（非数字），会导致成品图表不可用、
# 数据显示不出来。因此所有"关于图表内容的加分检查"都以【该图表已有可见的数值数据】
# 为前提：图表数值系列所引用的单元格必须均为真正的数字（openpyxl data_type=='n'），
# 否则视为"图表无数据显示"，该题及相关加分不予计分。
# 例外：加分项 19、20 只考察图表的几何摆放（高度、重叠、列位），不看图表内容是否正常，
# 只要该题出现了图表且未超限即可通过，故不施加上述前提。

SCRIPT_ID = "097"
# 维度二"加分项"满分（也是完成度分母 / max_score）。扣分项只做减法、不计入分母。
MAX_SCORE_TOTAL = 62

# 维度二所有评分项按"结果追加顺序"的 max_delta 列表；末端把 results 里的
# "+N：..." / "-N：..." 字符串反解为结构化 dim2_items 时用。顺序须与
# _evaluate_impl 中 results.append(...) 的顺序严格一致。
# 加分项在前（正值），扣分项在后（负值）。
_MAX_DELTAS_ORDER = [
    5,   # 01 共 16 个图表（对应第 6~21 题）
    1,   # 02 所有图表均为原生可编辑对象
    3,   # 03 第6题 饼图
    3,   # 04 第7题 饼图
    3,   # 05 第8题 横向条形图
    3,   # 06 第9题 横向条形图
    3,   # 07 第10题 饼图
    3,   # 08 第11题 饼图
    3,   # 09 第12题 饼图
    3,   # 10 第13题 横向条形图
    3,   # 11 第14题 横向条形图
    3,   # 12 第15题 横向条形图
    3,   # 13 第16题 饼图
    3,   # 14 第17题 横向条形图
    3,   # 15 第18题 横向条形图
    3,   # 16 第19题 横向条形图
    3,   # 17 第20题 横向条形图
    3,   # 18 第21题 横向条形图
    3,   # 19 高度不超过题间行距、互不重叠（按合规图表占比给分）
    3,   # 20 摆放整齐、置于每题数据右侧、列位一致（按合规图表占比给分）
    1,   # 21 六张饼图显示百分比数据标签
    1,   # 22 横向条形图分类轴显示选项文本
    -5,  # 23 引用外部工作簿 / 外部链接
    -3,  # 24 另建重复选项数据区
    -3,  # 25 数据自相矛盾 / 异常文本
    -3,  # 26 图表遮挡数据 / 相互重叠 / 错位
    -3,  # 27 题目结构被破坏
    -1,  # 28 给开放填空题（第22题）强行生成图表
    -1,  # 29 遗留多余的空图表 / 临时工作表
]
# 加分项累加须等于 MAX_SCORE_TOTAL（扣分项不计入分母）
assert sum(d for d in _MAX_DELTAS_ORDER if d > 0) == MAX_SCORE_TOTAL

# ============ 题目规格（可完全由成品文件自校验，不依赖原始文件） ============
# 每题：(题号, 图表类型 'pie'/'bar', 选项数, 抽检数值, 数据起始行, 数据结束行)
# 数据列为 B 列（"小计"），选项文本在 A 列。
_QUESTIONS = [
    (6,  "pie", 5,  94,   4,   8),
    (7,  "pie", 3,  124,  13,  15),
    (8,  "bar", 7,  198,  20,  26),
    (9,  "bar", 8,  203,  31,  38),
    (10, "pie", 4,  116,  43,  46),
    (11, "pie", 5,  111,  51,  55),
    (12, "pie", 4,  126,  60,  63),
    (13, "bar", 5,  181,  68,  72),
    (14, "bar", 7,  184,  77,  83),
    (15, "bar", 7,  211,  88,  94),
    (16, "pie", 5,  101,  99,  103),
    (17, "bar", 7,  194,  108, 114),
    (18, "bar", 8,  211,  119, 126),
    (19, "bar", 6,  198,  131, 136),
    (20, "bar", 8,  191,  141, 148),
    (21, "bar", 6,  196,  153, 158),
]
_PIE_QS = {6, 7, 10, 11, 12, 16}
_Q22_ROW_LO = 159   # 开放填空题（第22题）所在区域（无选项计数、不可制图）
_DATA_COL_LIMIT = 3  # 固有数据只占 A/B/C 三列（选项 / 小计 / 比例）

# 图表数值列（B 列）在题目区间内的起始行 -> 题号，用于把图表映射到题目
_ROWBAND_TO_Q = {}
for _q, _t, _n, _v, _r1, _r2 in _QUESTIONS:
    for _r in range(_r1, _r2 + 1):
        _ROWBAND_TO_Q[_r] = _q

# 题号 -> 下一题的数据起始行（最后一题以第22题区域起始行为界）。
# "该题至下一题的行距" = 下一题起始行 - 本题起始行，用于限制图表高度。
_Q_NEXT_START = {}
_Q_STARTS = [_r1 for (_q, _t, _n, _v, _r1, _r2) in _QUESTIONS]
for _i, (_q, _t, _n, _v, _r1, _r2) in enumerate(_QUESTIONS):
    _Q_NEXT_START[_q] = _Q_STARTS[_i + 1] if _i + 1 < len(_Q_STARTS) else _Q22_ROW_LO

_NS_STRIP = re.compile(r"<(/?)(?:[a-zA-Z0-9]+:)?")  # 去命名空间前缀，统一裸标签匹配


def _col_letters_to_idx(letters):
    """列字母 -> 1-based 列号（A->1, B->2, ...）。"""
    n = 0
    for ch in letters:
        n = n * 26 + (ord(ch.upper()) - 64)
    return n


def _read_zip_text(zf, name):
    try:
        return zf.read(name).decode("utf-8", "ignore")
    except KeyError:
        return ""


def _chart_files(zf):
    """按编号排序的 xl/charts/chartN.xml 列表。"""
    names = [n for n in zf.namelist() if re.match(r"xl/charts/chart\d+\.xml$", n)]
    return sorted(names, key=lambda x: int(re.search(r"(\d+)", x.rsplit("/", 1)[-1]).group(1)))


def _parse_chart(xml):
    """从单个 chartN.xml 解析出结构化信息（对命名空间前缀不敏感）：
      type: 'pie' / 'bar' / None
      bar_horizontal: barDir==bar 时为 True
      val_ref: 数值系列引用的区间字符串（如 sheet1!$B$4:$B$8），无则 None
      cat_ref: 分类轴引用区间（选项文本，如 sheet1!$A$4:$A$8），无则 None
      pt_count: 数值系列 numCache 里的点数（可见数据点个数），无则 0
      cat_has_str: 分类轴缓存里是否有非空字符串（选项文本已显示）
      show_percent: 数据标签是否显示百分比
    """
    info = {
        "type": None, "bar_horizontal": False,
        "val_ref": None, "cat_ref": None,
        "pt_count": 0, "cat_has_str": False, "show_percent": False,
    }
    if "pieChart" in xml or "pie3DChart" in xml or "ofPieChart" in xml or "doughnutChart" in xml:
        info["type"] = "pie"
    elif "barChart" in xml or "bar3DChart" in xml:
        info["type"] = "bar"

    m = re.search(r"barDir\s+val=\"(\w+)\"", xml)
    if m and m.group(1) == "bar":
        info["bar_horizontal"] = True

    # 数值系列 <c:val>...</c:val>
    vm = re.search(r"<(?:\w+:)?val>(.*?)</(?:\w+:)?val>", xml, re.S)
    if vm:
        seg = vm.group(1)
        fm = re.search(r"<(?:\w+:)?f>(.*?)</(?:\w+:)?f>", seg)
        if fm:
            info["val_ref"] = fm.group(1)
        if "numCache" in seg:
            pc = re.search(r"ptCount\s+val=\"(\d+)\"", seg)
            # 以 numCache 里实际存在的数值点为准（真正被渲染出来的数据）
            pts = re.findall(r"<(?:\w+:)?pt\b[^>]*>\s*<(?:\w+:)?v>(.*?)</(?:\w+:)?v>", seg)
            info["pt_count"] = len(pts) if pts else (int(pc.group(1)) if pc else 0)

    # 分类轴 <c:cat>...</c:cat>
    cm = re.search(r"<(?:\w+:)?cat>(.*?)</(?:\w+:)?cat>", xml, re.S)
    if cm:
        seg = cm.group(1)
        fm = re.search(r"<(?:\w+:)?f>(.*?)</(?:\w+:)?f>", seg)
        if fm:
            info["cat_ref"] = fm.group(1)
        cats = re.findall(r"<(?:\w+:)?pt\b[^>]*>\s*<(?:\w+:)?v>(.*?)</(?:\w+:)?v>", seg)
        info["cat_has_str"] = any(c.strip() for c in cats)

    # 数据标签百分比：showPercent val="1"
    if re.search(r"showPercent\s+val=\"1\"", xml):
        info["show_percent"] = True

    return info


def _ref_range_rows(ref):
    """从引用字符串（如 'sheet1'!$B$4:$B$8）解析出 (列号, 起始行, 结束行)。
    解析失败返回 (None, None, None)。"""
    if not ref:
        return (None, None, None)
    # 去掉工作表名前缀（... ! 之后才是单元格区间），避免把 sheet1 当成列/行。
    cell_part = ref.rsplit("!", 1)[-1]
    m = re.search(r"\$?([A-Za-z]+)\$?(\d+)(?::\$?([A-Za-z]+)\$?(\d+))?", cell_part)
    if not m:
        return (None, None, None)
    col = _col_letters_to_idx(m.group(1))
    r1 = int(m.group(2))
    r2 = int(m.group(4)) if m.group(4) else r1
    if r1 > r2:
        r1, r2 = r2, r1
    return (col, r1, r2)


def _cells_all_numeric(ws, col, r1, r2):
    """区间内单元格是否全部为真正的数字（data_type=='n' 且值为数值）。
    这是"图表有可见数值数据"的判定核心——原始文本型数字会使其为 False。"""
    if col is None or r1 is None:
        return False
    seen = False
    for r in range(r1, r2 + 1):
        c = ws.cell(row=r, column=col)
        if c.value is None:
            continue
        seen = True
        if c.data_type != "n" or isinstance(c.value, bool) or not isinstance(c.value, (int, float)):
            return False
    return seen


def _drawing_anchors(zf):
    """解析 xl/drawings/drawingN.xml，返回每个图表锚点的信息列表：
      [{'chart': 'chartN.xml', 'from_col': int, 'from_row': int,
        'to_col': int|None, 'to_row': int|None, 'kind': 'one'/'two'}]
    对 xdr: 前缀与无前缀两种写法都兼容；通过 drawing 的 .rels 把 r:id 映射到 chart 文件。
    """
    out = []
    draw_names = [n for n in zf.namelist() if re.match(r"xl/drawings/drawing\d+\.xml$", n)]
    for dn in sorted(draw_names):
        dx = _read_zip_text(zf, dn)
        base = dn.rsplit("/", 1)[-1]
        rels = _read_zip_text(zf, "xl/drawings/_rels/" + base + ".rels")
        rid2chart = {}
        for m in re.finditer(r'Id="(rId\d+)"[^>]*Target="([^"]*chart\d+\.xml)"', rels):
            rid2chart[m.group(1)] = m.group(2).rsplit("/", 1)[-1]
        # 也兼容 Target 在 Id 前的顺序
        for m in re.finditer(r'Target="([^"]*chart\d+\.xml)"[^>]*Id="(rId\d+)"', rels):
            rid2chart[m.group(2)] = m.group(1).rsplit("/", 1)[-1]

        for m in re.finditer(r"<(?:xdr:)?(oneCellAnchor|twoCellAnchor)\b(.*?)</(?:xdr:)?\1>", dx, re.S):
            blk = m.group(0)
            kind = "one" if m.group(1) == "oneCellAnchor" else "two"
            rid = re.search(r'r:id="(rId\d+)"', blk)
            chart = rid2chart.get(rid.group(1)) if rid else None
            fm = re.search(
                r"<(?:xdr:)?from>.*?<(?:xdr:)?col>(\d+)</(?:xdr:)?col>.*?<(?:xdr:)?row>(\d+)</(?:xdr:)?row>",
                blk, re.S)
            from_col = int(fm.group(1)) if fm else None
            from_row = int(fm.group(2)) if fm else None
            to_col = to_row = None
            tm = re.search(
                r"<(?:xdr:)?to>.*?<(?:xdr:)?col>(\d+)</(?:xdr:)?col>.*?<(?:xdr:)?row>(\d+)</(?:xdr:)?row>",
                blk, re.S)
            if tm:
                to_col = int(tm.group(1))
                to_row = int(tm.group(2))
            # 尺寸（EMU）：oneCellAnchor 用 <xdr:ext>，twoCellAnchor 用 <a:ext>
            cx = cy = None
            em = (re.search(r"<(?:xdr:)?ext\s+cx=\"(\d+)\"\s+cy=\"(\d+)\"", blk)
                  or re.search(r"<a:ext\s+cx=\"(\d+)\"\s+cy=\"(\d+)\"", blk))
            if em:
                cx = int(em.group(1))
                cy = int(em.group(2))
            out.append({
                "chart": chart, "kind": kind,
                "from_col": from_col, "from_row": from_row,
                "to_col": to_col, "to_row": to_row,
                "cx": cx, "cy": cy,
            })
    return out


def _has_external_links(zf):
    """是否引用了外部工作簿（externalLinks/externalReferences）。"""
    if any(n.startswith("xl/externalLinks/") and n.endswith(".xml") for n in zf.namelist()):
        return True
    wb = _read_zip_text(zf, "xl/workbook.xml")
    if "externalReference" in wb:
        return True
    return False


def evaluate(dir_path):
    """按 §2.2 约定，返回结构化 dict；不 print、不改 sys.stdout、不 sys.exit。
    入参 dir_path 为"脚本所在目录的路径"，脚本自己在该目录里定位并打开被评估的
    .xlsx（忽略以 ~$ 开头的临时锁文件）。脚本自身崩溃（含目录不存在、未发现文档
    等）由本函数捕获并返回 status='error'。"""
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
        # 在 dir_path 目录内定位被评估的 .xlsx/.xlsm（忽略 ~$ 打头的临时锁文件）
        if not dir_path or not os.path.isdir(dir_path):
            raise FileNotFoundError(f"目录不存在：{dir_path}")
        cands = sorted(
            f for f in os.listdir(dir_path)
            if f.lower().endswith((".xlsx", ".xlsm")) and not f.startswith("~$")
        )
        if not cands:
            raise FileNotFoundError(f"目录内未找到 .xlsx/.xlsm 文件：{dir_path}")
        file_path = os.path.join(dir_path, cands[0])
        result["file_name"] = os.path.basename(file_path)
        return _evaluate_impl(file_path, result)
    except Exception as exc:
        # status 为 error 时未能完成判定，dim1_pass 留空，dim1_reason 也为空；
        # 具体原因只在 error 字段中给出。
        result["status"] = "error"
        result["error"] = f"{type(exc).__name__}: {exc}"
        result["dim1_pass"] = ""
        result["dim1_reason"] = ""
        result["dim2_items"] = []
        result["total_score"] = 0
        return result


def _evaluate_impl(file_path, result):
    """维度一 + 维度二的评估主体；对内部构造 results 列表沿用原逻辑，末尾按
    _MAX_DELTAS_ORDER 反解成结构化 dim2_items 填入 result。"""
    results = []

    def _emit(hit, delta, rule):
        """追加一条评分结果。rule 只写细则本身，不带实测值等补充说明。"""
        results.append(f"{delta:+d}：{rule}" if hit else f"+0：{rule}")

    # ========== 维度1：可用与可修改性 ==========
    # 交付文件须为 .xlsx/.xlsm 且能被正常解析打开。任一失败 → 维度一不通过，完成度记 0。
    try:
        zf = zipfile.ZipFile(file_path)
        wb = openpyxl.load_workbook(file_path, data_only=False)
        ws = wb.active
    except Exception as exc:
        result["dim1_pass"] = False
        result["dim1_reason"] = f"文件无法作为 Excel 打开：{type(exc).__name__}: {exc}"
        result["dim2_items"] = []
        result["total_score"] = 0
        return result
    result["dim1_pass"] = True
    result["dim1_reason"] = ""   # 维度一通过时不输出原因

    # ========== 预解析：图表、锚点、映射 ==========
    chart_names = _chart_files(zf)
    charts = {}          # 'chartN.xml' -> parsed info
    for cn in chart_names:
        charts[cn.rsplit("/", 1)[-1]] = _parse_chart(_read_zip_text(zf, cn))
    anchors = _drawing_anchors(zf)
    anchor_by_chart = {}
    for a in anchors:
        if a["chart"]:
            anchor_by_chart[a["chart"]] = a

    # 把每个图表按其数值系列引用的 B 列行区间映射到题号
    def _chart_question(info):
        col, r1, r2 = _ref_range_rows(info.get("val_ref"))
        if r1 is None:
            return None, (col, r1, r2)
        # 用区间起始行落在哪一题的数据带判定
        q = _ROWBAND_TO_Q.get(r1)
        if q is None:
            # 容错：用区间中任意行匹配
            for rr in range(r1, r2 + 1):
                if rr in _ROWBAND_TO_Q:
                    q = _ROWBAND_TO_Q[rr]
                    break
        return q, (col, r1, r2)

    q_to_chart = {}      # 题号 -> (chart_name, info, (col,r1,r2))
    q22_charts = []      # 落在第22题（开放填空）区域的图表
    for cname, info in charts.items():
        q, (col, r1, r2) = _chart_question(info)
        if q is not None and q not in q_to_chart:
            q_to_chart[q] = (cname, info, (col, r1, r2))
        if r1 is not None and r1 >= _Q22_ROW_LO:
            q22_charts.append(cname)

    def _cell_numeric_ok(col, r1, r2):
        """图表数值区间是否为可见数值数据（precondition）。"""
        return _cells_all_numeric(ws, col, r1, r2)

    # ---------- 加分项 01：共 16 个图表 ----------
    n_charts = len(charts)
    _emit(n_charts == 16, 5, "工作表中共出现 16 个图表，分别对应第 6～21 题")

    # ---------- 加分项 02：所有图表均为原生可编辑对象 ----------
    # 以 chartN.xml 承载即为 Excel 原生图表对象（非图片/截图）；存在图表即视为可编辑。
    _emit(n_charts >= 1, 1,
          "所有图表均为 Excel 原生可编辑对象（图表、数据系列可编辑），非图片或截图")

    # ---------- 加分项 03~18：逐题图表 ----------
    # 每题 +3：需同时满足（a）该题数据右侧存在图表；（b）图表已有可见数值数据（precondition）；
    #          （c）类型正确（饼图 / 横向条形图）；（d）扇区/条形数量正确；（e）抽检数值命中。
    def _q_desc(q, typ, n_opts, spot):
        opt_last = chr(ord("A") + n_opts - 1)
        if typ == "pie":
            return (f"在第{q}题数据右侧区域出现一个可编辑的饼图，包含 {n_opts} 个扇区"
                    f"（对应 A～{opt_last} {n_opts}个选项），其中一个扇区的数值为 {spot}")
        return (f"在第{q}题数据右侧区域出现一个可编辑的横向条形图，包含 {n_opts} 根条形"
                f"（对应 A～{opt_last} {n_opts}个选项），其中一根条形的长度对应数值 {spot}")

    for q, typ, n_opts, spot, r1, r2 in _QUESTIONS:
        desc = _q_desc(q, typ, n_opts, spot)
        entry = q_to_chart.get(q)
        ok = False
        if entry is not None:
            cname, info, (col, cr1, cr2) = entry
            anc = anchor_by_chart.get(cname)
            # （a）右侧：锚点起始列在数据三列（A/B/C = 0/1/2）之右
            right_ok = (anc is not None and anc.get("from_col") is not None
                        and anc["from_col"] >= _DATA_COL_LIMIT)
            # （b）precondition：数值区间为可见数值数据
            data_visible = _cell_numeric_ok(col if col else 2, cr1, cr2)
            # （c）类型
            if typ == "pie":
                type_ok = (info["type"] == "pie")
            else:
                type_ok = (info["type"] == "bar" and info["bar_horizontal"])
            # （d）数量：数值区间长度 == 选项数
            count = (cr2 - cr1 + 1) if (cr1 is not None) else 0
            count_ok = (count == n_opts)
            # （e）抽检数值命中：直接读源单元格数值，命中 spot
            spot_ok = False
            if data_visible and col:
                vals = []
                for rr in range(cr1, cr2 + 1):
                    v = ws.cell(row=rr, column=col).value
                    if isinstance(v, (int, float)) and not isinstance(v, bool):
                        vals.append(int(round(v)))
                spot_ok = spot in vals
            ok = right_ok and data_visible and type_ok and count_ok and spot_ok
        _emit(ok, 3, desc)

    # ---------- 加分项 19：高度不超过题间行距、互不重叠 ----------
    # 按合规图表占比给分。合规 = 图表纵向跨度（行数）不超过"本题起始行至下一题起始行"的行距，
    #   且底部不越入下一题区域、不与其他图表重叠。
    # 注意：本项只看几何摆放，不看图表内容是否正常——只要该题出现了图表且未超限即可通过，
    #   因此不施加"有可见数值数据"的前提（与加分项 20 一致）。
    def _rect_of(anc):
        """图表占据的 (top_row, bottom_row, left_col, right_col)，行列均为 0-based。
        twoCell 直接用 from/to；oneCell 用尺寸估算行列跨度。锚点信息不全时返回 None。"""
        if anc is None:
            return None
        top = anc.get("from_row")
        left = anc.get("from_col")
        if top is None or left is None:
            return None
        if anc["kind"] == "two" and anc.get("to_row") is not None and anc.get("to_col") is not None:
            return (top, anc["to_row"], left, anc["to_col"])
        # oneCell：用 EMU 尺寸估算跨度（默认行高≈190500EMU、列宽≈610000EMU）
        rows = max(1, int(round((anc.get("cy") or 0) / 190500.0)))
        cols = max(1, int(round((anc.get("cx") or 0) / 610000.0)))
        return (top, top + rows, left, left + cols)

    def _overlap(r1r, r2r):
        if r1r is None or r2r is None:
            return False
        return not (r1r[1] <= r2r[0] or r2r[1] <= r1r[0]
                    or r1r[3] <= r2r[2] or r2r[3] <= r1r[2])

    # 先收齐所有参评图表的矩形，再两两判重叠，保证判定对称、占比统计准确
    size_cands = []      # [(题号, 本题起始行, rect)]
    for q, typ, n_opts, spot, r1, r2 in _QUESTIONS:
        entry = q_to_chart.get(q)
        if entry is None:
            continue     # 该题没有图表才不计入分母；图表数据是否正常不影响本项
        cname, info, (col, cr1, cr2) = entry
        size_cands.append((q, r1, _rect_of(anchor_by_chart.get(cname))))

    size_total = 0
    size_ok = 0
    for q, r1, rect in size_cands:
        size_total += 1
        next_start = _Q_NEXT_START[q]
        gap_rows = next_start - r1                      # 该题至下一题的行距（行数）
        span_rows = (rect[1] - rect[0]) if rect else 0   # 图表纵向跨度（行数）
        height_ok = (1 <= span_rows <= gap_rows)
        # 不跨题遮挡：图表底部不越入下一题区域（rect 为 0-based，+1 转 1-based 行号）
        no_cross = (rect is not None and (rect[1] + 1) <= next_start)
        overlapped = any(_overlap(rect, other)
                         for q2, _r, other in size_cands if q2 != q)
        if height_ok and no_cross and not overlapped:
            size_ok += 1
    _emit(size_total > 0 and size_ok == size_total, 3,
          "每个图表高度不超过该题至下一题的行距，相邻图表互不重叠、不跨题遮挡")

    # ---------- 加分项 20：摆放整齐、统一置于每题数据右侧 ----------
    # 同加分项 19：只看摆放位置，不看图表内容是否正常。
    # "各图列位一致"是全局属性，这里折算为逐图判定——以出现最多的起始列为基准列，
    #   起始列与基准列相同才算列位一致，从而能按合规图表占比统计。
    place_cands = []     # [(题号, 本题起始行, 本题结束行, anc)]
    for q, typ, n_opts, spot, r1, r2 in _QUESTIONS:
        entry = q_to_chart.get(q)
        if entry is None:
            continue
        cname, info, (col, cr1, cr2) = entry
        place_cands.append((q, r1, r2, anchor_by_chart.get(cname)))

    # 基准列：所有图表起始列的众数（并列时取较小的列号，保证结果稳定）
    col_counts = {}
    for _q, _r1, _r2, anc in place_cands:
        c = anc.get("from_col") if anc else None
        if c is not None:
            col_counts[c] = col_counts.get(c, 0) + 1
    base_col = min(col_counts, key=lambda c: (-col_counts[c], c)) if col_counts else None

    place_total = 0
    place_ok = 0
    for q, r1, r2, anc in place_cands:
        place_total += 1
        # 置于数据右侧：起始列在 A/B/C 之右（0-based 列号 >= 3）
        col_right = (anc is not None and anc.get("from_col") is not None
                     and anc["from_col"] >= _DATA_COL_LIMIT)
        # 与本题各行水平对齐：起始行大致落在本题行区间附近
        row_near = (anc is not None and anc.get("from_row") is not None
                    and (r1 - 3) <= (anc["from_row"] + 1) <= (r2 + 3))
        # 各图列位一致：起始列与基准列相同
        col_same = (anc is not None and anc.get("from_col") == base_col)
        if col_right and row_near and col_same:
            place_ok += 1
    _emit(place_total > 0 and place_ok == place_total, 3,
          "图表摆放整齐，统一置于每题数据右侧、与本题各行水平对齐，不遮挡数据，各图列位一致")

    # ---------- 加分项 21：六张饼图显示百分比数据标签 ----------
    pie_qs = [q for (q, t, n, v, a, b) in _QUESTIONS if t == "pie"]
    pie_pct_total = 0
    pie_pct_ok = 0
    for q in pie_qs:
        entry = q_to_chart.get(q)
        if entry is None:
            continue
        cname, info, (col, cr1, cr2) = entry
        if info["type"] != "pie" or not _cell_numeric_ok(col if col else 2, cr1, cr2):
            continue
        pie_pct_total += 1
        if info["show_percent"]:
            pie_pct_ok += 1
    _emit(pie_pct_total > 0 and pie_pct_ok == pie_pct_total, 1,
          "6 张饼图（第6、7、10、11、12、16题）显示百分比数据标签")

    # ---------- 加分项 22：横向条形图分类轴显示选项文本 ----------
    bar_qs = [q for (q, t, n, v, a, b) in _QUESTIONS if t == "bar"]
    bar_cat_total = 0
    bar_cat_ok = 0
    for q in bar_qs:
        entry = q_to_chart.get(q)
        if entry is None:
            continue
        cname, info, (col, cr1, cr2) = entry
        if not (info["type"] == "bar" and info["bar_horizontal"]) or not _cell_numeric_ok(col if col else 2, cr1, cr2):
            continue
        bar_cat_total += 1
        # 分类轴引用 A 列（选项文本），或缓存里已有选项文本字符串
        cat_col = _ref_range_rows(info.get("cat_ref"))[0]
        if (cat_col == 1) or info["cat_has_str"]:
            bar_cat_ok += 1
    _emit(bar_cat_total > 0 and bar_cat_ok == bar_cat_total, 1,
          "各横向条形图的分类轴依次显示选项文本，不出现 1、2、3 等默认序号")

    # ========== 扣分项（自校验，不与原始文件比较） ==========

    # ---------- 扣分项 23：外部链接 ----------
    _emit(_has_external_links(zf), -5,
          "图表数据引用了外部工作簿（存在外部链接），打开时弹出\"更新外部链接\"警告或链接源缺失导致图表无法显示，文件不自包含")

    # ---------- 扣分项 24：另建重复选项数据区 ----------
    # 固有数据只占 A/B/C 三列；若在 D 列及之后另有成片的数值/文本数据区（供图表引用），扣分。
    extra_col_cells = 0
    max_col = ws.max_column or 0
    if max_col > _DATA_COL_LIMIT:
        for row in ws.iter_rows(min_col=_DATA_COL_LIMIT + 1, max_col=max_col, min_row=1, max_row=ws.max_row):
            for c in row:
                if c.value not in (None, ""):
                    extra_col_cells += 1
    _emit(extra_col_cells >= 3, -3,
          "工作表中除每题固有的三列（选项/小计/比例）外，另建了重复的选项数据区")

    # ---------- 扣分项 25：数据自相矛盾 / 异常文本 ----------
    # 自校验：(a) 选项/表头单元格出现残缺空括号"（）""()"、明显乱码；
    #         (b) 图表 numCache 缓存值与其所引用源单元格的当前数值不一致（喂给图表的数据对不上）。
    anomaly = False
    for row in ws.iter_rows(min_row=1, max_row=ws.max_row, min_col=1, max_col=max(max_col, 3)):
        for c in row:
            if isinstance(c.value, str):
                t = c.value
                if "（）" in t or "()" in t.replace(" ", "") or "�" in t:
                    anomaly = True
    # 缓存与源数据一致性
    for cname, info in charts.items():
        col, cr1, cr2 = _ref_range_rows(info.get("val_ref"))
        if col is None or cr1 is None:
            continue
        seg_xml = _read_zip_text(zf, "xl/charts/" + cname)
        vm = re.search(r"<(?:\w+:)?val>(.*?)</(?:\w+:)?val>", seg_xml, re.S)
        if not vm:
            continue
        pts = re.findall(r"<(?:\w+:)?pt\b[^>]*idx=\"(\d+)\"[^>]*>\s*<(?:\w+:)?v>(.*?)</(?:\w+:)?v>", vm.group(1))
        for idx_s, v_s in pts:
            try:
                cached = float(v_s)
            except ValueError:
                continue
            r = cr1 + int(idx_s)
            if r > cr2:
                continue
            cell_v = ws.cell(row=r, column=col).value
            if isinstance(cell_v, (int, float)) and not isinstance(cell_v, bool):
                if abs(float(cell_v) - cached) > 0.5:
                    anomaly = True
    _emit(anomaly, -3, "文件内数据自相矛盾或存在异常文本")

    # ---------- 扣分项 26：图表遮挡数据 / 相互重叠 / 错位 ----------
    misplaced = False
    rects_all = {}
    for cname, anc in anchor_by_chart.items():
        rects_all[cname] = _rect_of(anc)
        # 错位：起始列压在数据三列（A/B/C）上 → 遮挡数据
        if anc.get("from_col") is not None and anc["from_col"] < _DATA_COL_LIMIT:
            misplaced = True
    names_list = list(rects_all.keys())
    for i in range(len(names_list)):
        for j in range(i + 1, len(names_list)):
            if _overlap(rects_all[names_list[i]], rects_all[names_list[j]]):
                misplaced = True
    _emit(misplaced, -3, "图表遮挡了数据，或多个图表相互重叠、与所属题目错位")

    # ---------- 扣分项 27：题目结构被破坏 ----------
    # 自校验：每题的选项行（A 列）与"本题有效填写人次"行必须完整存在、非空。
    structure_broken = False
    for q, typ, n_opts, spot, r1, r2 in _QUESTIONS:
        for rr in range(r1, r2 + 1):
            if ws.cell(row=rr, column=1).value in (None, ""):
                structure_broken = True
                break
    # "本题有效填写人次"汇总行应存在（原表每题一行）
    summary_hits = 0
    for rr in range(1, (ws.max_row or 0) + 1):
        v = ws.cell(row=rr, column=1).value
        if isinstance(v, str) and "本题有效填写人次" in v:
            summary_hits += 1
    if summary_hits < len(_QUESTIONS):
        structure_broken = True
    _emit(structure_broken, -3,
          "题目结构被破坏，某题选项行或\"本题有效填写人次\"行缺失、被清空或被图表覆盖")

    # ---------- 扣分项 28：给开放填空题（第22题）强行生成图表 ----------
    _emit(bool(q22_charts), -1,
          "给开放填空题（第22题）强行生成图表，该题无选项与计数数据，图表内容无意义或为空")

    # ---------- 扣分项 29：遗留多余的空图表 / 临时工作表 ----------
    extra_sheets = [s for s in wb.sheetnames if s != ws.title]
    empty_charts = 0
    mapped = {c for (c, _i, _r) in q_to_chart.values()}
    for cname, info in charts.items():
        col, cr1, cr2 = _ref_range_rows(info.get("val_ref"))
        # 未映射到任何题目、且无数值系列引用 → 视为空图表
        if cname not in mapped and (col is None or cr1 is None):
            empty_charts += 1
    _emit(bool(extra_sheets) or empty_charts > 0, -1,
          "遗留多余的空图表、临时工作表或无关产物")

    # ========== 将 results 反解为结构化 dim2_items ==========
    if len(results) != len(_MAX_DELTAS_ORDER):
        raise RuntimeError(
            f"dim2 项数不一致：results={len(results)} vs order={len(_MAX_DELTAS_ORDER)}"
        )
    dim2_items = []
    total_score = 0
    for raw, max_delta in zip(results, _MAX_DELTAS_ORDER):
        # 格式约定：'+N：' / '-N：' 前缀 + 规则说明。命中加分项 delta=max_delta(>0)；
        # 命中扣分项 delta=max_delta(<0)；未命中一律 delta=0。
        head, _, rule = raw.partition("：")
        head = head.strip()
        try:
            parsed = int(head.replace("＋", "+").replace("－", "-"))
        except ValueError:
            parsed = 0
        if max_delta > 0:
            hit = parsed > 0
            delta = max_delta if hit else 0
        else:
            hit = parsed < 0
            delta = max_delta if hit else 0
        dim2_items.append({
            "rule": rule,
            "max_delta": max_delta,
            "delta": delta,
            "hit": hit,
        })
        total_score += delta

    # 完成度分数：加分实得 + 扣分（负），下限 0；分母为加分满分 MAX_SCORE_TOTAL。
    total_score = max(0, total_score)
    result["dim2_items"] = dim2_items
    result["total_score"] = total_score
    result["max_score"] = MAX_SCORE_TOTAL
    return result


if __name__ == "__main__":
    # 本地调试入口：evaluate 只接收"脚本所在目录的路径"，由脚本自己在该目录里定位并
    # 打开被评估的 .xlsx。命令行第 1 个参数可覆盖目录路径；未指定时使用脚本自身所在目录。
    # 注：evaluate() 的接口约定就是"入参为目录路径"——路径始终由外部显式传入，
    #     这里的默认值仅为本地调试便利，不影响 evaluate 的接口约定。
    if len(sys.argv) >= 2:
        _dir = sys.argv[1]
    else:
        _dir = os.path.dirname(os.path.abspath(__file__))
    _out = evaluate(_dir)
    print(json.dumps(_out, ensure_ascii=False, indent=2))


