"""
自动评估脚本：企业财务统计总表_已添加财务图表.xlsx

指令：我需要毛利率和净利率趋势图、柱形图，ROE的趋势图，资产负债率的趋势图，
      营业收入增长率、净利润的增长率趋势图。帮我根据表格相应数据在 sheet2 中
      绘制这些图表，规定所有图表横坐标为时间。

评估逻辑：
1. 先检查维度1（可用与可修改性）。维度1不通过，直接返回 total_score=0。
2. 维度1通过后，按维度2逐项自动检测。满足得分点则累加正分；命中扣分点则累加负分。
3. 最后以结构化字典的形式返回评估结果。

评分原则：帮用户干了多少活（加分项，全部来自指令要求的图表任务），
          惹了多少麻烦（减分项，全部来自指令未涉及、却破坏原文件的多余改动）。
维度2每张图逐张单独评分，分两条——“满足要求”（图表是否存在、类型是否正确）与
“数据引用是否正确”（数据系列是否引用了正确的列、横坐标是否为报告期时间）。
分值仅三档：1=简单点、3=中等点、5=高难度点。

依赖：openpyxl
安装：pip install openpyxl
"""
import io
import json
import os
import re
import sys
import zipfile

import openpyxl
from openpyxl.chart import LineChart, BarChart


def _load_workbook_resilient(file_path):
    """载入工作簿，绕开 openpyxl 对 <c:externalData r:id="…"/> 的解析 bug。

    某些由 WPS/Excel 保存的 xlsx，其 chartN.xml 里会写入
    <c:externalData r:id="rId1">…</c:externalData>（保存缓存数据链接用）。
    openpyxl 3.x 的 ExternalData 描述器在读 r:id 时会得到 None，
    进而在 drawings.py 里被 warn 一下并整张图丢弃，导致 ws._charts 为空。

    这里在内存中把每个 chartN.xml 里的 <c:externalData …/> 元素去掉，
    再交给 openpyxl 解析，图表就能正常读到。不修改磁盘上的原文件。
    兼容图表存放于 xl/charts/ 或 xl/drawings/charts/ 两种方式。
    """
    with open(file_path, "rb") as f:
        raw = f.read()

    buf_out = io.BytesIO()
    with zipfile.ZipFile(io.BytesIO(raw), "r") as zin, \
            zipfile.ZipFile(buf_out, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            name = item.filename
            if re.search(r"chart\d+\.xml$", name) and "colors" not in name and "style" not in name:
                text = data.decode("utf-8")
                # 同时兼容自闭合与成对写法
                text = re.sub(r"<c:externalData\b[^>]*?/>", "", text)
                text = re.sub(
                    r"<c:externalData\b[^>]*?>.*?</c:externalData>",
                    "",
                    text,
                    flags=re.DOTALL,
                )
                data = text.encode("utf-8")
            zout.writestr(item, data)

    buf_out.seek(0)
    return openpyxl.load_workbook(buf_out, data_only=True)


SCRIPT_ID = "089"
SOURCE_SHEET = "企业财务统计总表"
TARGET_SHEET = "Sheet2"
TARGET_FILE_NAME = "企业财务统计总表_已添加财务图表.xlsx"
EXPECTED_CAT_REF = "企业财务统计总表!A4:A12"  # 指令规定的横坐标（报告期时间）引用

# 原始输入表（企业财务统计总表）的关键锚点单元格，用于校验原表未被破坏/篡改
INPUT_ANCHORS = {
    "A1": "企业年度/季度财务统计总表",
    "B3": "毛利率", "C3": "净利率", "D3": "净资产收益率(ROE)",
    "H3": "资产负债率", "K3": "营业收入增长率", "L3": "净利润增长率",
    "A4": "2021年度", "A12": "2026Q4",
}

# 指令要求的 5 张图表逐张评测规格：
#   type       ：期望图表类型（line 折线 / bar 柱形）
#   cols       ：数据系列应引用的列（多列表示同一张图内需含多条系列）
#   exist_delta：满足要求（存在+类型正确）的分值
#   data_delta ：数据引用正确（列引用无误 + 横坐标为时间）的分值
#   keywords   ：按标题关键词辅助定位
CHART_SPECS = [
    {"key": "gross_net_line", "name": "毛利率和净利率趋势图", "type": "line",
     "cols": ["B", "C"], "exist_delta": 1, "data_delta": 5,
     "keywords": ("毛利率", "净利率", "趋势")},
    {"key": "gross_net_bar", "name": "毛利率和净利率柱形图", "type": "bar",
     "cols": ["B", "C"], "exist_delta": 1, "data_delta": 5,
     "keywords": ("毛利率", "净利率", "柱")},
    {"key": "roe_line", "name": "净资产收益率(ROE)趋势图", "type": "line",
     "cols": ["D"], "exist_delta": 1, "data_delta": 3,
     "keywords": ("净资产收益率",)},
    {"key": "debt_line", "name": "资产负债率趋势图", "type": "line",
     "cols": ["H"], "exist_delta": 1, "data_delta": 3,
     "keywords": ("资产负债率",)},
    {"key": "growth_line", "name": "营业收入增长率、净利润增长率趋势图", "type": "line",
     "cols": ["K", "L"], "exist_delta": 1, "data_delta": 5,
     "keywords": ("营业收入增长率", "净利润增长率")},
]


# ───────────────────────────── 基础工具函数 ─────────────────────────────
def norm_text(v):
    return "" if v is None else str(v).strip()


def normalize_ref(ref):
    """把 Excel 引用统一成便于比较的形式：去掉 $、单引号、空格。"""
    if ref is None:
        return ""
    return str(ref).replace("$", "").replace("'", "").replace(" ", "").strip()


def ref_match(actual, expected):
    return normalize_ref(actual) == normalize_ref(expected)


def formula_or_ref(obj):
    """openpyxl 不同对象版本可能使用 .f 或 .ref 存储引用。"""
    if obj is None:
        return None
    for attr in ("f", "ref"):
        try:
            value = getattr(obj, attr)
            if value:
                return value
        except Exception:
            pass
    return None


def ser_val_ref(ser):
    try:
        return formula_or_ref(ser.val.numRef)
    except Exception:
        return None


def ser_cat_ref(ser):
    try:
        if ser.cat.strRef:
            return formula_or_ref(ser.cat.strRef)
    except Exception:
        pass
    try:
        if ser.cat.numRef:
            return formula_or_ref(ser.cat.numRef)
    except Exception:
        pass
    return None


def chart_title(chart):
    """读取图表标题文本。"""
    try:
        title = chart.title
        if title is None:
            return ""
        try:
            parts = []
            for para in title.tx.rich.p:
                for run in para.r:
                    parts.append(run.t or "")
            if parts:
                return "".join(parts).strip()
        except Exception:
            pass
        return str(title).strip()
    except Exception:
        return ""


def ref_col_letter(ref):
    """从形如 企业财务统计总表!B4:B12 的引用里取出列字母。"""
    ref = normalize_ref(ref)
    tail = ref.split("!")[-1]
    m = re.search(r"([A-Z]+)\d", tail)
    return m.group(1) if m else None


def chart_val_cols(chart):
    """返回图表全部数据系列引用到的列字母集合。"""
    cols = set()
    for s in chart.series:
        c = ref_col_letter(ser_val_ref(s))
        if c:
            cols.add(c)
    return cols


def chart_cat_ok(chart):
    """横坐标（类别轴）是否全部引用报告期列 A4:A12。"""
    return bool(chart.series) and all(
        ref_match(ser_cat_ref(s), EXPECTED_CAT_REF) for s in chart.series
    )


def chart_matches_type(chart, want_type):
    if want_type == "line":
        return isinstance(chart, LineChart)
    if want_type == "bar":
        return isinstance(chart, BarChart)
    return False


def read_source_cells(file_path):
    """读取原数据表(第1张工作表)的单元格文本（含公式），用于判断原表是否被破坏/篡改。
    直接解析 xml，避免 data_only 丢失公式；返回 (工作表名, {坐标: 值字符串}) 或 (name, None)。"""
    try:
        z = zipfile.ZipFile(file_path)
        names = z.namelist()
        ns = {
            "ss": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
            "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
            "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
        }
        import xml.etree.ElementTree as ET
        wb = ET.fromstring(z.read("xl/workbook.xml"))
        sheets = wb.findall(".//ss:sheets/ss:sheet", ns)
        if not sheets:
            return None, None
        name1 = sheets[0].get("name")
        rid = sheets[0].get("{%s}id" % ns["r"])
        wbrels = ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))
        target = None
        for rel in wbrels.findall(".//rel:Relationship", ns):
            if rel.get("Id") == rid:
                t = rel.get("Target")
                target = t.lstrip("/") if t.startswith("/") else "xl/" + t
                break
        if not target or target not in names:
            return name1, None
        # sharedStrings
        sst = []
        if "xl/sharedStrings.xml" in names:
            sroot = ET.fromstring(z.read("xl/sharedStrings.xml"))
            for si in sroot.findall("ss:si", ns):
                sst.append("".join(t.text or "" for t in si.findall(".//ss:t", ns)))
        ws = ET.fromstring(z.read(target))
        cells = {}
        for c in ws.findall(".//ss:sheetData/ss:row/ss:c", ns):
            ref = c.get("r")
            ctype = c.get("t")
            f = c.find("ss:f", ns)
            v = c.find("ss:v", ns)
            if f is not None and f.text is not None:
                cells[ref] = "=" + f.text
            elif ctype == "inlineStr":
                it = c.find(".//ss:t", ns)
                cells[ref] = it.text if it is not None else ""
            elif v is not None:
                if ctype == "s":
                    try:
                        cells[ref] = sst[int(v.text)]
                    except (ValueError, IndexError):
                        cells[ref] = v.text
                else:
                    cells[ref] = v.text
        return name1, cells
    except Exception:
        return None, None


def evaluate(dir_path: str) -> dict:
    """评估入口。

    参数：
        dir_path: 脚本所在目录的路径。脚本自己在该目录里定位并打开被评估的文档。

    返回：
        结构化评估结果 dict：id / file_name / status / error / dim1_pass /
        dim1_reason / dim2_items / total_score / max_score。
    """
    result = {
        "id": SCRIPT_ID,
        "file_name": TARGET_FILE_NAME,
        "status": "ok",
        "error": None,
        "dim1_pass": False,
        "dim1_reason": "",
        "dim2_items": [],
        "total_score": 0,
        "max_score": 0,
    }

    # 每项：{"rule", "max_delta", "delta", "hit"}
    dim2_items = []
    score = 0

    def add_item(max_delta, rule, ok, _detail=""):
        # _detail 仅为保留各判定处的可读注释而存在，不再写入返回结果。
        nonlocal score
        delta = max_delta if ok else 0
        # 扣分项：max_delta 为负数，命中即代表“触发扣分”
        if ok:
            score += delta
        dim2_items.append({
            "rule": rule,
            "max_delta": max_delta,
            "delta": delta,
            "hit": ok,
        })

    def hit(max_delta, rule, _detail=""):
        add_item(max_delta, rule, True, _detail)

    def miss(max_delta, rule, _detail=""):
        add_item(max_delta, rule, False, _detail)

    try:
        # ─────────── 在目录内定位被评估文档 ───────────
        if not dir_path or not os.path.isdir(dir_path):
            raise FileNotFoundError(f"目录不存在或不是目录：{dir_path}")
        file_path = os.path.join(dir_path, TARGET_FILE_NAME)
        if os.path.isfile(file_path):
            result["file_name"] = TARGET_FILE_NAME
        else:
            candidates = []
            try:
                for name in os.listdir(dir_path):
                    ext = os.path.splitext(name)[1].lower()
                    if (ext in (".xlsx", ".xlsm") and not name.startswith("~$")
                            and os.path.isfile(os.path.join(dir_path, name))):
                        candidates.append(name)
            except OSError:
                candidates = []
            if not candidates:
                raise FileNotFoundError(f"目录中未找到 .xlsx 文件：{dir_path}")
            file_path = os.path.join(dir_path, candidates[0])
            result["file_name"] = candidates[0]

        # ─────────── 维度1：可用与可修改性 ───────────
        # 仅核验：交付文件为 .xlsx/.xlsm 且可正常打开。
        dim1_pass = True
        dim1_reason = ""

        ext = os.path.splitext(file_path)[1].lower()
        if ext not in (".xlsx", ".xlsm"):
            dim1_pass = False
            dim1_reason = "文件格式不是 .xlsx/.xlsm"

        wb = None
        if dim1_pass:
            try:
                if not zipfile.is_zipfile(file_path):
                    raise ValueError("不是有效的 Office Open XML 文件")
                wb = _load_workbook_resilient(file_path)
            except Exception as e:
                dim1_pass = False
                dim1_reason = f"文件无法正常打开：{e}"

        result["dim1_pass"] = dim1_pass
        if dim1_pass:
            result["dim1_reason"] = dim1_reason
        else:
            # dim1_reason 统一带上「维度一未通过：…」前缀与计分后果，便于直接作为错误摘要展示
            result["dim1_reason"] = f"维度一未通过：{dim1_reason}，完成度总分记 0"

        if not dim1_pass:
            # 维度1未通过，不再检查维度2，直接返回
            # max_score 按维度2 正分项满分之和填写，便于对齐
            result["max_score"] = sum(
                s["exist_delta"] + s["data_delta"] for s in CHART_SPECS
            )
            result["total_score"] = 0
            return result

        # 目标工作表缺失不再作为维度1的硬性否决项：按“无图表”进入维度2逐项评分。
        ws2 = wb[TARGET_SHEET] if TARGET_SHEET in wb.sheetnames else None
        charts = list(ws2._charts) if ws2 is not None else []

        # ─────────── 维度2：完成度评分（逐张图单独评分）───────────
        for spec in CHART_SPECS:
            want_type = spec["type"]
            want_cols = set(spec["cols"])
            name = spec["name"]

            # 定位本图：优先按“类型 + 数据列相关”匹配，退而按标题关键词
            typed = [c for c in charts if chart_matches_type(c, want_type)]
            related = [c for c in typed if chart_val_cols(c) & want_cols]
            if related:
                cand = related
            else:
                by_title = [c for c in typed
                            if all(k in chart_title(c) for k in spec["keywords"])]
                cand = by_title if by_title else typed
            best = max(cand, key=lambda c: len(chart_val_cols(c) & want_cols)) if cand else None

            # 1) 满足要求：存在 + 类型正确（图表读自 ws2._charts，天然位于 Sheet2）
            exist_ok = best is not None and chart_matches_type(best, want_type)
            if exist_ok:
                hit(spec["exist_delta"],
                    f"满足要求——Sheet2中存在“{name}”，图表类型为{'折线图' if want_type == 'line' else '柱形图'}")
            else:
                miss(spec["exist_delta"],
                     f"满足要求——Sheet2中存在“{name}”，图表类型为{'折线图' if want_type == 'line' else '柱形图'}",
                     "未找到对应类型的图表")

            # 2) 数据引用正确：所需列全部命中且不多引其它指标列 + 横坐标为报告期时间
            cols_cn = "、".join(f"{c}列" for c in spec["cols"])
            rule_data = (f"数据引用正确——“{name}”的数据系列引用{cols_cn}"
                         f"，且横坐标（类别轴）为报告期时间(A4:A12)，无错引/漏引/多引")
            if best is None:
                miss(spec["data_delta"], rule_data, "无对应图表，数据引用无从评")
                continue
            actual_cols = chart_val_cols(best)
            cols_ok = want_cols.issubset(actual_cols)
            extra_cols = actual_cols - want_cols  # 多引了指令该图不该有的其它指标列
            time_ok = chart_cat_ok(best)
            data_ok = cols_ok and not extra_cols and time_ok
            if data_ok:
                hit(spec["data_delta"], rule_data)
            else:
                miss(spec["data_delta"], rule_data,
                     f"应引用={sorted(want_cols)}，实际={sorted(actual_cols)}，"
                     f"多引={sorted(extra_cols)}，横坐标为时间={time_ok}")

        # ─────────── 减分项：指令未涉及、却破坏原文件的多余改动 ───────────
        name1, cells = read_source_cells(file_path)
        if cells is None:
            cells = {}
        tampered = [k for k, v in INPUT_ANCHORS.items() if norm_text(cells.get(k)) != v]
        if tampered:
            hit(-3, "篡改或删除了原数据表(企业财务统计总表)中的单元格数据/公式（指令仅要求新增图表，原表应原样保留）",
                f"原表被改动的锚点={tampered}")
        else:
            miss(-3, "篡改或删除了原数据表(企业财务统计总表)中的单元格数据/公式（指令仅要求新增图表，原表应原样保留）")

        if name1 != SOURCE_SHEET:
            hit(-3, "删除或重命名了原始工作表(企业财务统计总表)", f"第1张工作表名={name1!r}")
        else:
            miss(-3, "删除或重命名了原始工作表(企业财务统计总表)")

        extra_sheets = [n for n in wb.sheetnames if n not in (SOURCE_SHEET, TARGET_SHEET)]
        if extra_sheets:
            hit(-1, "新增了指令未要求的多余工作表（除必需的Sheet2外）", f"多余工作表={extra_sheets}")
        else:
            miss(-1, "新增了指令未要求的多余工作表（除必需的Sheet2外）")

        # ─────────── 汇总 ───────────
        result["dim2_items"] = dim2_items
        result["total_score"] = max(0, score)  # total 下限为 0
        result["max_score"] = sum(item["max_delta"] for item in dim2_items if item["max_delta"] > 0)
        return result

    except FileNotFoundError as e:
        # 未找到待评文件：不进入维度判定，dim1_pass 留空，分值一并归零
        result["file_name"] = ""
        result["status"] = "error"
        result["error"] = str(e)
        result["dim1_pass"] = ""
        result["dim1_reason"] = ""
        result["dim2_items"] = []
        result["total_score"] = 0
        result["max_score"] = 0
        return result

    except Exception as e:
        result["status"] = "error"
        result["error"] = f"脚本运行异常：{type(e).__name__}: {e}"
        result["dim1_pass"] = False
        result["dim1_reason"] = (f"维度一未通过：文件无法打开或解析"
                                f"（{type(e).__name__}: {e}），完成度总分记 0")
        result["dim2_items"] = []
        result["total_score"] = 0
        return result


if __name__ == "__main__":
    # 本地调试：默认使用脚本所在目录；也支持通过 sys.argv[1] 指定目录路径。
    _dir = sys.argv[1] if len(sys.argv) > 1 else os.path.dirname(os.path.abspath(__file__))
    print(json.dumps(evaluate(_dir), ensure_ascii=False, indent=2))
