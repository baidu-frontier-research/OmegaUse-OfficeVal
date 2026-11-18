#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# pyright: reportUnknownMemberType=false, reportUnknownVariableType=false, reportUnknownArgumentType=false, reportUnknownParameterType=false, reportMissingParameterType=false, reportPrivateUsage=false, reportUnknownLambdaType=false, reportGeneralTypeIssues=false

import os
import re
import sys
import json
import zipfile
import xml.etree.ElementTree as ET

from openpyxl import load_workbook
from openpyxl.utils import get_column_letter, column_index_from_string

# 说明：按接口统一约定，本模块不修改全局 sys.stdout，也不 print 主结果。
# 主结果统一由 evaluate(dir_path) -> dict 返回；入参为“脚本所在目录的路径”，
# 由脚本自己在该目录内定位并打开被评估的工作簿。本文件仅在 __main__ 下作为
# 本地调试入口，才把 dict 序列化为 JSON 打印。
#
# 任务：在“调查结果.xlsx”（sheet1，A:C 三列、207 行的问卷统计表）中，为第 4 题
# 至第 25 题各插入 1 张二维饼图，插在该题右侧（D 列及其右侧）。共 22 张。
# 维度一只判文件格式：交付物是否为可编辑的 .xlsx/.xlsm 工作簿、能否正常打开；
# 图表可见性与能否渲染出扇形归维度二对应加分项判定。
# 维度二判定一律以“图表实际引用的单元格区域内容”为依据（细则注 3）：写死的静态
# 数值、只存在于 numCache/strCache 里的数值不算；被隐藏、尺寸为 0、被完全遮挡，
# 或引用列被隐藏而扇形画不出来的图表，其对应加分项全部不得分。

SCRIPT_ID = "099"
# 维度二加分项 max_delta 之和；也是维度一未通过时回填的 max_score
# （脚本没评上时 max_score 为 0）
MAX_SCORE_TOTAL = 73
# 维度二减分项 max_delta 之和（负值），仅用于自检
MAX_DEDUCT_TOTAL = -27

# 维度二所有评分项按“结果追加顺序”的 max_delta 列表；末端把 results 里的
# “+N：...”/“-N：...”字符串反解为结构化 dim2_items 时用。顺序须与
# _evaluate_impl 中 add(...) 的顺序严格一致；共 50 项，
# 正项累加 = MAX_SCORE_TOTAL，负项累加 = MAX_DEDUCT_TOTAL。
_MAX_DELTAS_ORDER = [
    5,   # 01 22 张图表一题一图、均可见且渲染出扇形；第 1–24 行无图表
    5,   # 02 22 张全部为二维饼图
    5,   # 03 各选项扇形填充色互不相同
    5,   # 04 均显示标注百分比的数据标签
    5,   # 05 均有图例且条目文字为各选项名称
    5,   # 06 均有标题且含“第N题”与该题题干
    5,   # 07 左边界均在 D 列或其右侧
    5,   # 08 插在对应题目右侧、纵向有重合、两两不堆叠
    3,   # 09 高度≥120 磅、宽高比≤2.5:1
    3,   # 10 分类与数值均以单元格区域引用绑定且引用列可见
    3,   # 11 扇形排列顺序与选项自上而下顺序一致
    1,   # 12 标题/图例/数据标签字号≥8 磅
    1,   # 13 右边界均不超出 K 列
    # 14–35：第 4 题至第 25 题逐题的扇形数与分类/数值引用内容
    1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1,
    1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1,
    -3,  # 36 A:C 的 25 题结构与 A1/A25/A200/A75/A191 内容
    -3,  # 37 25 个题干行 A:C 跨列合并且合并区域为指定的 25 处
    -3,  # 38 题干行字体微软雅黑 10 号加粗白字、底色 2F6F73
    -3,  # 39 “选项/小计/比例”表头行字体与底色 DCECEF
    -3,  # 40 选项数据行字体与 FFFDF6 / F7F3E8 交替底色
    -3,  # 41 表头行/数据行/人次行的水平垂直对齐与自动换行
    -1,  # 42 25 个“本题有效填写人次”行字体与底色 E7F2E4
    -1,  # 43 C 列比例保留两位小数百分比、人次行显示 100%
    -1,  # 44 只占 A/B/C 三列共 207 行且 24 个指定行为空白
    -1,  # 45 A/B/C 三列均为可见状态
    -1,  # 46 A–K 列列宽均为 9 字符（±1）
    -1,  # 47 D 列及其右侧除图表外无任何内容
    -1,  # 48 只有 1 个含数据的工作表且名为 sheet1
    -1,  # 49 25 个题干行位于指定行号
    -1,  # 50 无自动筛选/冻结窗格/条件格式，A1:C207 无边框
]
assert len(_MAX_DELTAS_ORDER) == 50
assert sum(d for d in _MAX_DELTAS_ORDER if d > 0) == MAX_SCORE_TOTAL
assert sum(d for d in _MAX_DELTAS_ORDER if d < 0) == MAX_DEDUCT_TOTAL

_EXT_OK = (".xlsx", ".xlsm")
# 维度一只接受 .xlsx/.xlsm；下列后缀视为格式不符，判维度一不通过并在
# dim1_reason 里写明，而不是抛异常。
_EXT_REJECT = (".xls", ".xlsb", ".csv", ".ods", ".et")

# 第 4–25 题的基准结构，取自输入文件“调查结果 (2).xlsx”并与细则逐条核对：
#   题号: (题干行, 选项首行, “本题有效填写人次”行, [选项名], [小计], [比例文本])
# 选项行区间 = 选项首行 .. 人次行-1；细则要求扇形只能来自这段区间。
_QUESTIONS = {
    4: (25, 27, 31, ['完全不了解', '听说过但很少使用', '比较了解', '经常使用并熟悉'],
        [18, 42, 69, 39], ['10.71%', '25.00%', '41.07%', '23.21%']),
    5: (33, 35, 40, ['明显偏简单', '稍微偏简单', '难度适中', '稍微偏难', '明显偏难'],
        [14, 26, 73, 39, 16], ['8.33%', '15.48%', '43.45%', '23.21%', '9.52%']),
    6: (42, 44, 48, ['完全不能', '只能说出一部分', '基本能说清楚', '能说清楚并能安排计划'],
        [20, 48, 65, 35], ['11.90%', '28.57%', '38.69%', '20.83%']),
    7: (50, 52, 56, ['比较陌生，不知道从哪开始', '有点兴趣，但担心完成不好', '有信心，愿意尝试', '比较熟悉，能较快进入状态'],
        [25, 54, 61, 28], ['14.88%', '32.14%', '36.31%', '16.67%']),
    8: (58, 60, 65, ['听老师讲解并记录要点', '独立阅读资料并完成练习', '小组讨论、合作完成任务', '动手实践或制作成果', '观看微课或使用学习平台'],
        [38, 36, 31, 34, 29], ['22.62%', '21.43%', '18.45%', '20.24%', '17.26%']),
    9: (67, 69, 73, ['老师直接归纳结论', '老师演示过程，我跟着理解', '自己先尝试，再听老师讲解', '小组合作探究后交流'],
        [31, 49, 53, 35], ['18.45%', '29.17%', '31.55%', '20.83%']),
    10: (75, 77, 82, ['先暂时跳过，等后面再看', '主动请教老师', '和同学讨论', '查阅课本、资料或学习单', '使用电脑或移动设备检索'],
         [67, 112, 95, 86, 74], ['39.88%', '66.67%', '56.55%', '51.19%', '44.05%']),
    11: (84, 86, 90, ['从不整理', '老师要求时才整理', '偶尔自己整理', '经常主动整理'],
         [22, 46, 64, 36], ['13.10%', '27.38%', '38.10%', '21.43%']),
    12: (92, 94, 99, ['几乎没有', '比较少', '一般', '比较多', '非常多'],
         [17, 35, 52, 44, 20], ['10.12%', '20.83%', '30.95%', '26.19%', '11.90%']),
    13: (101, 103, 107, ['完全没兴趣', '兴趣一般', '比较有兴趣', '非常有兴趣'],
         [15, 52, 68, 33], ['8.93%', '30.95%', '40.48%', '19.64%']),
    14: (109, 111, 115, ['主题阅读与资料整理', '动手实验或模型制作', '生活问题分析与应用', '成果展示与交流分享'],
         [28, 47, 55, 38], ['16.67%', '27.98%', '32.74%', '22.62%']),
    15: (117, 119, 123, ['完全不愿意', '偶尔会了解', '比较愿意', '非常愿意，会主动尝试'],
         [18, 42, 70, 38], ['10.71%', '25.00%', '41.67%', '22.62%']),
    16: (125, 127, 131, ['完全没有帮助', '帮助不大', '有一定帮助', '帮助很大'],
         [13, 33, 82, 40], ['7.74%', '19.64%', '48.81%', '23.81%']),
    17: (133, 135, 139, ['完全不会', '偶尔能完成', '基本能完成', '能熟练完成并优化呈现'],
         [16, 39, 79, 34], ['9.52%', '23.21%', '47.02%', '20.24%']),
    18: (141, 143, 147, ['完全不能', '简单任务可以完成，复杂任务易出错', '大部分任务能完成，偶尔需要提示', '能准确完成并检查结果'],
         [18, 47, 73, 30], ['10.71%', '27.98%', '43.45%', '17.86%']),
    19: (149, 151, 155, ['完全不能', '能解决简单问题', '能解决大部分问题', '能解决复杂问题并总结方法'],
         [20, 52, 64, 32], ['11.90%', '30.95%', '38.10%', '19.05%']),
    20: (157, 159, 163, ['完全没有提升', '稍微有一点提升', '有明显提升', '提升很大'],
         [12, 41, 82, 33], ['7.14%', '24.40%', '48.81%', '19.64%']),
    21: (165, 167, 172, ['趣味导入或主题故事', '动手实践任务', '小组合作项目', '生活案例分析', '闯关练习或竞赛活动'],
         [24, 42, 36, 39, 27], ['14.29%', '25.00%', '21.43%', '23.21%', '16.07%']),
    22: (174, 176, 180, ['使用实物或板书演示', '用多媒体动画展示', '让学生先动手体验', '结合生活情境讲解'],
         [34, 37, 57, 40], ['20.24%', '22.02%', '33.93%', '23.81%']),
    23: (182, 184, 189, ['基础巩固练习', '实际应用任务', '拓展思考任务', '趣味闯关任务', '实践记录任务'],
         [30, 39, 28, 43, 28], ['17.86%', '23.21%', '16.67%', '25.60%', '16.67%']),
    24: (191, 193, 198, ['任务完成规范性', '探究过程与思路', '小组合作表现', '资料整理能力', '创新表达与改进'],
         [105, 119, 91, 87, 83], ['62.50%', '70.83%', '54.17%', '51.79%', '49.40%']),
    25: (200, 202, 207, ['非常不满意', '不太满意', '一般', '比较满意', '非常满意'],
         [11, 24, 42, 61, 30], ['6.55%', '14.29%', '25.00%', '36.31%', '17.86%']),
}
_QNOS = list(range(4, 26))          # 需要插图的 22 个题号
# 25 个题干行行号（含第 1–3 题）
_STEM_ROWS = [1, 7, 16, 25, 33, 42, 50, 58, 67, 75, 84, 92, 101, 109, 117, 125,
              133, 141, 149, 157, 165, 174, 182, 191, 200]
# 25 个“本题有效填写人次”行行号
_TOTAL_ROWS = [5, 14, 23, 31, 40, 48, 56, 65, 73, 82, 90, 99, 107, 115, 123, 131,
               139, 147, 155, 163, 172, 180, 189, 198, 207]
# 24 个题间空白行行号
_BLANK_ROWS = [6, 15, 24, 32, 41, 49, 57, 66, 74, 83, 91, 100, 108, 116, 124, 132,
               140, 148, 156, 164, 173, 181, 190, 199]
_MAX_ROW = 207
_TOTAL_LABEL = "本题有效填写人次"
_HEADER_TEXTS = ("选项", "小计", "比例")
_EXPECTED_A1 = "第1题：你的年级阶段"
_EXPECTED_A25 = "第4题：你对学校图书角、阅读区等学习空间的了解程度如何"
_EXPECTED_A200 = "第25题：总的来说，你对当前校园学习活动安排满意吗"
_MULTI_QNOS = (10, 24)              # A75 / A191 为多选题

_FONT_NAME = "微软雅黑"
_FONT_SZ = 10.0
_FILL_STEM = "2F6F73"               # 题干行底色
_FILL_HEADER = "DCECEF"             # 表头行底色
_FILL_DATA = ("FFFDF6", "F7F3E8")   # 数据行交替底色
_FILL_TOTAL = "E7F2E4"              # 人次行底色
_COLOR_WHITE = "FFFFFF"

_COL_W_TARGET = 9.0                 # A–K 目标列宽
_COL_W_TOL = 1.0
_MIN_H_PT = 120.0                   # 图表最小高度（磅）
_MAX_RATIO = 2.5                    # 图表最大宽高比
_MIN_FONT_SZ = 8.0                  # 图表内最小字号（磅）
_COL_LEFT_MIN = 3                   # D 列（0 基）
_COL_RIGHT_MAX = 10                 # K 列（0 基）

# EMU 换算：1 磅 = 12700 EMU；1 像素 = 9525 EMU；默认字体最大数字宽 8 像素，
# 故 1 字符列宽 = 8 × 9525 = 76200 EMU（已用 good 输出文件的 twoCellAnchor 反算校验）。
_EMU_PT = 12700
_EMU_PER_WIDTH = 76200
_DEF_COL_W = 8.43
_DEF_ROW_H = 15.0
_GRID_COLS = 64
_GRID_ROWS = 320

_NS = {
    "xdr": "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "c": "http://schemas.openxmlformats.org/drawingml/2006/chart",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "pr": "http://schemas.openxmlformats.org/package/2006/relationships",
}
# 二维饼图之外的其它图表类型标记（命中即非二维饼图）
_PIE2D = "pieChart"
_OTHER_KINDS = (
    "pie3DChart", "doughnutChart", "ofPieChart", "bar3DChart", "barChart",
    "line3DChart", "lineChart", "areaChart", "area3DChart", "scatterChart",
    "radarChart", "bubbleChart", "stockChart", "surfaceChart", "surface3DChart",
)
_ERR_TOKENS = ("#REF!", "#VALUE!", "#NAME?", "#DIV/0!", "#N/A", "#NULL!", "#NUM!")
def _pick_file(dir_path):
    """在目录内挑出被评估的工作簿：优先 .xlsx/.xlsm，其次记录被拒后缀。"""
    cands, rejects = [], []
    for name in sorted(os.listdir(dir_path)):
        if name.startswith("~$") or name.startswith("."):
            continue
        full = os.path.join(dir_path, name)
        if not os.path.isfile(full):
            continue
        ext = os.path.splitext(name)[1].lower()
        if ext in _EXT_OK:
            cands.append(full)
        elif ext in _EXT_REJECT:
            rejects.append(full)
    if cands:
        # 名字里含“调查结果”的优先，其次取体积最大的（含图表的那份）
        named = [p for p in cands if "调查结果" in os.path.basename(p)]
        pool = named or cands
        pool.sort(key=lambda p: (-os.path.getsize(p), p))
        return pool[0], None
    if rejects:
        return None, rejects[0]
    return None, None


def _rgb6(color):
    """openpyxl 颜色 -> 6 位大写 RGB；主题色/索引色/无色返回 None。"""
    if color is None:
        return None
    v = getattr(color, "rgb", None)
    if not isinstance(v, str):
        return None
    v = v.upper()
    if len(v) == 8:
        v = v[2:]
    return v if re.fullmatch(r"[0-9A-F]{6}", v) else None


def _fill_rgb(cell):
    fill = cell.fill
    if fill is None or fill.patternType not in ("solid",):
        return None
    return _rgb6(fill.fgColor)


def _font_ok(cell, bold, white=False):
    f = cell.font
    if f is None:
        return False
    if (f.name or "") != _FONT_NAME:
        return False
    if float(f.sz or 0) != _FONT_SZ:
        return False
    if bool(f.b) != bool(bold):
        return False
    if white and _rgb6(f.color) != _COLOR_WHITE:
        return False
    return True


def _align_ok(cell, horizontal, vertical="center", wrap=True):
    """horizontal / vertical 可传字符串或可选值元组；文本单元格默认左对齐，
    故“左对齐”允许 None/general 这类未显式设置的等效写法（由调用方传元组）。"""
    al = cell.alignment
    if al is None:
        return False
    hs = (horizontal,) if isinstance(horizontal, str) else tuple(horizontal)
    vs = (vertical,) if isinstance(vertical, str) else tuple(vertical)
    return ((al.horizontal or "") in hs
            and (al.vertical or "") in vs
            and bool(al.wrap_text) is bool(wrap))


_H_LEFT = ("left", "")              # 文本默认即左对齐
_H_CENTER = ("center",)
_V_CENTER = ("center",)


def _display_text(cell):
    """按单元格数字格式还原“显示出来的文字”，用于判定 C 列百分比显示形式。"""
    v = cell.value
    if v is None:
        return ""
    if isinstance(v, str):
        return v.strip()
    if isinstance(v, bool):
        return "TRUE" if v else "FALSE"
    if isinstance(v, (int, float)):
        fmt = (cell.number_format or "General")
        fmt0 = fmt.split(";")[0]
        if "%" in fmt0:
            m = re.search(r"0\.(0+)\s*%", fmt0)
            dec = len(m.group(1)) if m else 0
            return "%.*f%%" % (dec, float(v) * 100.0)
        m = re.search(r"0\.(0+)", fmt0)
        if m:
            return "%.*f" % (len(m.group(1)), float(v))
        if isinstance(v, float) and v == int(v):
            return str(int(v))
        return str(v)
    return str(v)


def _as_num(v):
    """把单元格值统一成 float：数字直接用；'82'、'48.81%' 这类文本也解析。"""
    if isinstance(v, bool):
        return None
    if isinstance(v, (int, float)):
        return float(v)
    if isinstance(v, str):
        s = v.strip().replace(",", "")
        pct = s.endswith("%")
        if pct:
            s = s[:-1].strip()
        try:
            x = float(s)
        except ValueError:
            return None
        return x / 100.0 if pct else x
    return None


def _plot_num(v):
    """按 Excel 绘图语义取单元格数值：只有真正的数字才会被画成扇形。
    文本型单元格（t="str"，例如把 18 存成字符串 '18'）在图表里等同空值，
    不参与绘制，所以不能像 _as_num 那样宽松解析。"""
    if isinstance(v, bool) or isinstance(v, str):
        return None
    if isinstance(v, (int, float)):
        return float(v)
    return None


def _col_width(ws, idx1):
    """第 idx1 列（1 基）的有效列宽（字符）；隐藏列按 0 计。"""
    dim = ws.column_dimensions.get(get_column_letter(idx1))
    if dim is not None and dim.hidden:
        return 0.0
    if dim is not None and dim.width:
        return float(dim.width)
    dw = ws.sheet_format.defaultColWidth if ws.sheet_format is not None else None
    return float(dw) if dw else _DEF_COL_W


def _row_height(ws, idx1):
    """第 idx1 行（1 基）的有效行高（磅）；隐藏行按 0 计。"""
    dim = ws.row_dimensions.get(idx1)
    if dim is not None and dim.hidden:
        return 0.0
    if dim is not None and dim.height:
        return float(dim.height)
    dh = ws.sheet_format.defaultRowHeight if ws.sheet_format is not None else None
    return float(dh) if dh else _DEF_ROW_H


def _grid(ws):
    """返回列/行的累积边界（EMU），下标为 0 基列号 / 0 基行号。"""
    cols = [0]
    for i in range(1, _GRID_COLS + 1):
        cols.append(cols[-1] + int(round(_col_width(ws, i) * _EMU_PER_WIDTH)))
    rows = [0]
    for i in range(1, _GRID_ROWS + 1):
        rows.append(rows[-1] + int(round(_row_height(ws, i) * _EMU_PT)))
    return cols, rows


def _locate(edges, pos):
    """给定累积边界与坐标（EMU），返回其所在格的 0 基下标与格内偏移。"""
    lo = 0
    while lo + 1 < len(edges) and edges[lo + 1] <= pos:
        lo += 1
    return lo, pos - edges[lo]


def _row_at(rows, pos):
    """EMU 纵坐标 -> 连续行号（1 基，带小数），用于纵向重合判定。"""
    i, off = _locate(rows, pos)
    span = rows[min(i + 1, len(rows) - 1)] - rows[i]
    frac = (off / span) if span > 0 else 0.0
    return i + 1 + frac


def _parse_ref(ref):
    """'sheet1'!$A$27:$A$30 -> (sheet, col1, row1, col2, row2)；解析不了返回 None。"""
    if not ref:
        return None
    s = ref.strip()
    if "!" not in s:
        return None
    sheet, _, addr = s.rpartition("!")
    sheet = sheet.strip()
    if sheet.startswith("'") and sheet.endswith("'"):
        sheet = sheet[1:-1].replace("''", "'")
    m = re.fullmatch(r"\$?([A-Za-z]{1,3})\$?(\d+)(?::\$?([A-Za-z]{1,3})\$?(\d+))?", addr.strip())
    if not m:
        return None
    c1 = column_index_from_string(m.group(1).upper())
    r1 = int(m.group(2))
    c2 = column_index_from_string(m.group(3).upper()) if m.group(3) else c1
    r2 = int(m.group(4)) if m.group(4) else r1
    return (sheet, min(c1, c2), min(r1, r2), max(c1, c2), max(r1, r2))


def _ref_cells(wb, ref):
    """按引用取回真实单元格对象列表（按行、再按列顺序）；引用无效返回 None。"""
    info = _parse_ref(ref)
    if info is None:
        return None
    sheet, c1, r1, c2, r2 = info
    target = None
    for ws in wb.worksheets:
        if ws.title == sheet or ws.title.lower() == sheet.lower():
            target = ws
            break
    if target is None:
        return None
    out = []
    for r in range(r1, r2 + 1):
        for c in range(c1, c2 + 1):
            out.append(target.cell(row=r, column=c))
    return out
def _lname(el):
    t = el.tag
    return t.rsplit("}", 1)[-1] if "}" in t else t


def _norm_join(base, target):
    if target.startswith("/"):
        return target.lstrip("/")
    return os.path.normpath(os.path.join(base, target)).replace("\\", "/")


def _rels(z, part):
    """part 的关系映射 rId -> zip 内绝对路径（跳过外部关系）。"""
    d, base = os.path.split(part)
    out = {}
    try:
        data = z.read("%s/_rels/%s.rels" % (d, base) if d else "_rels/%s.rels" % base)
    except KeyError:
        return out
    for rel in ET.fromstring(data).findall("pr:Relationship", _NS):
        if (rel.get("TargetMode") or "") == "External":
            continue
        out[rel.get("Id")] = _norm_join(d, rel.get("Target") or "")
    return out


def _sheet_parts(z):
    """按 workbook.xml 顺序返回 [(表名, zip 内路径, state), ...]。"""
    ns = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
    root = ET.fromstring(z.read("xl/workbook.xml"))
    rels = _rels(z, "xl/workbook.xml")
    rid_attr = "{%s}id" % _NS["r"]
    out = []
    for sh in root.findall("m:sheets/m:sheet", ns):
        out.append((sh.get("name") or "", rels.get(sh.get(rid_attr) or ""),
                    sh.get("state") or "visible"))
    return out


def _int0(el, attr):
    try:
        return int(el.get(attr) or 0)
    except (TypeError, ValueError):
        return 0


def _cell_ref(el, tag):
    """<xdr:from>/<xdr:to> -> (col, colOff, row, rowOff)，均为 0 基/EMU。"""
    sub = el.find("xdr:%s" % tag, _NS)
    if sub is None:
        return None
    def g(name):
        n = sub.find("xdr:%s" % name, _NS)
        try:
            return int((n.text or "0").strip()) if n is not None else 0
        except ValueError:
            return 0
    return (g("col"), g("colOff"), g("row"), g("rowOff"))
def _clr_key(node):
    """颜色节点 -> 可比较的字符串键（含 lumMod/tint 等修饰，避免同色误判为不同）。"""
    if node is None:
        return None
    kind = _lname(node)
    key = "%s:%s" % (kind, (node.get("val") or "").upper())
    mods = []
    for m in node:
        mods.append("%s=%s" % (_lname(m), m.get("val") or ""))
    if mods:
        key += "[" + ",".join(sorted(mods)) + "]"
    return key


def _fill_key(sp_pr):
    """spPr -> 填充键；无 spPr / 未设置填充返回 None（表示继承自动配色）。"""
    if sp_pr is None:
        return None
    if sp_pr.find("a:noFill", _NS) is not None:
        return "noFill"
    sf = sp_pr.find("a:solidFill", _NS)
    if sf is not None:
        for child in sf:
            k = _clr_key(child)
            if k:
                return "solid|" + k
        return "solid|?"
    for tag in ("gradFill", "pattFill", "blipFill", "grpFill"):
        node = sp_pr.find("a:%s" % tag, _NS)
        if node is not None:
            return tag + "|" + ET.tostring(node, encoding="unicode")[:120]
    return None


def _flag(el, tag):
    """<c:xxx val="1"/> -> True/False/None（无该元素）。"""
    node = el.find("c:%s" % tag, _NS) if el is not None else None
    if node is None:
        return None
    v = (node.get("val") or "1").strip().lower()
    return v not in ("0", "false")


def _dlbls_flags(el):
    """c:dLbls -> {'delete':bool,'percent':bool,'value':bool,...}；无则 None。"""
    if el is None:
        return None
    return {
        "delete": bool(_flag(el, "delete")),
        "percent": bool(_flag(el, "showPercent")),
        "value": bool(_flag(el, "showVal")),
        "cat": bool(_flag(el, "showCatName")),
        "ser": bool(_flag(el, "showSerName")),
        "key": bool(_flag(el, "showLegendKey")),
    }


def _font_szs(el):
    """子树内显式设置的字号（磅）列表。"""
    if el is None:
        return []
    out = []
    for node in el.iter():
        if _lname(node) in ("defRPr", "rPr") and node.get("sz"):
            try:
                out.append(float(node.get("sz")) / 100.0)
            except ValueError:
                pass
    return out


def _texts(el):
    if el is None:
        return ""
    return "".join(n.text or "" for n in el.iter() if _lname(n) == "t")
def _axis_ref(holder):
    """c:cat / c:val -> (引用式, 是否字面量, 按 pt/idx 排序的缓存值列表)。"""
    if holder is None:
        return (None, False, [])

    def cache_of(node):
        pts = []
        for pt in node.iter():
            if _lname(pt) != "pt":
                continue
            v = None
            for sub in pt:
                if _lname(sub) == "v":
                    v = sub.text or ""
                    break
            pts.append((_int0(pt, "idx"), v if v is not None else ""))
        pts.sort(key=lambda t: t[0])
        return [v for _, v in pts]

    for tag in ("strRef", "numRef", "multiLvlStrRef"):
        ref = holder.find("c:%s" % tag, _NS)
        if ref is not None:
            f = ref.find("c:f", _NS)
            return ((f.text or "").strip() if f is not None else None,
                    False, cache_of(ref))
    for tag in ("strLit", "numLit"):
        lit = holder.find("c:%s" % tag, _NS)
        if lit is not None:
            return (None, True, cache_of(lit))
    return (None, False, [])


def _ser_info(ser):
    cat_ref, cat_lit, cat_cache = _axis_ref(ser.find("c:cat", _NS))
    val_ref, val_lit, val_cache = _axis_ref(ser.find("c:val", _NS))
    dpts = []
    for dpt in ser.findall("c:dPt", _NS):
        idx = dpt.find("c:idx", _NS)
        dpts.append((_int0(idx, "val") if idx is not None else -1,
                     _fill_key(dpt.find("c:spPr", _NS))))
    tx = ser.find("c:tx", _NS)
    return {
        "cat_ref": cat_ref, "cat_lit": cat_lit, "cat_cache": cat_cache,
        "val_ref": val_ref, "val_lit": val_lit, "val_cache": val_cache,
        "dpts": dpts, "ser_fill": _fill_key(ser.find("c:spPr", _NS)),
        "dlbls": _dlbls_flags(ser.find("c:dLbls", _NS)),
        "name": _texts(tx),
        "name_ref": (lambda n: (n.text or "").strip() if n is not None else None)(
            tx.find("c:strRef/c:f", _NS) if tx is not None else None),
    }


def _parse_chart(z, part):
    """解析 chart 部件，返回图表类型、系列、标题、图例、字号等信息。"""
    info = {"part": part, "kinds": [], "sers": [], "title": "", "title_ref": None,
            "has_title": False, "legend": None, "font_szs": [], "vary": None,
            "raw_err": False}
    try:
        raw = z.read(part)
    except KeyError:
        return info
    if any(tok.encode("utf-8") in raw for tok in _ERR_TOKENS):
        info["raw_err"] = True
    root = ET.fromstring(raw)
    chart = root.find("c:chart", _NS)
    plot = chart.find("c:plotArea", _NS) if chart is not None else None
    if plot is not None:
        for child in plot:
            name = _lname(child)
            if not name.endswith("Chart"):
                continue
            info["kinds"].append(name)
            if name == _PIE2D:
                info["vary"] = _flag(child, "varyColors")
            group_dlbls = _dlbls_flags(child.find("c:dLbls", _NS))
            for ser in child.findall("c:ser", _NS):
                s = _ser_info(ser)
                if s["dlbls"] is None:
                    s["dlbls"] = group_dlbls
                s["kind"] = name
                s["vary"] = _flag(child, "varyColors")
                info["sers"].append(s)
    title = chart.find("c:title", _NS) if chart is not None else None
    deleted = bool(_flag(chart, "autoTitleDeleted")) if chart is not None else False
    if title is not None and not bool(_flag(title, "delete")):
        info["has_title"] = True
        info["title"] = _texts(title.find("c:tx", _NS)) or _texts(title)
        node = title.find("c:tx/c:strRef/c:f", _NS)
        info["title_ref"] = (node.text or "").strip() if node is not None else None
    elif title is None and not deleted:
        # 自动标题：单系列饼图会自动显示系列名，但不含题号+题干，按“无有效标题”处理
        info["has_title"] = False
    legend = chart.find("c:legend", _NS) if chart is not None else None
    if legend is not None and not bool(_flag(legend, "delete")):
        pos = legend.find("c:legendPos", _NS)
        info["legend"] = {
            "pos": (pos.get("val") if pos is not None else "r"),
            "hidden_entries": [_int0(e.find("c:idx", _NS), "val")
                               for e in legend.findall("c:legendEntry", _NS)
                               if bool(_flag(e, "delete"))],
        }
    szs = _font_szs(title) + _font_szs(legend) + _font_szs(root.find("c:txPr", _NS))
    if plot is not None:
        for node in plot.iter():
            if _lname(node) == "dLbls":
                szs += _font_szs(node)
    info["font_szs"] = szs
    pvo = _flag(chart, "plotVisOnly") if chart is not None else None
    info["plot_vis_only"] = True if pvo is None else bool(pvo)
    return info
def _edge(edges, i):
    return edges[i] if 0 <= i < len(edges) else edges[-1]


def _collect_drawings(path, wb):
    """遍历每张工作表的 drawing，返回 (charts, shapes)。
    charts 每项含所在表、几何（EMU 与行号）、隐藏标记与 chart 部件解析结果；
    shapes 为同一张表上的非图表对象（图片/文本框等），用于遮挡判定。"""
    charts, shapes = [], []
    ns_m = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
    with zipfile.ZipFile(path) as z:
        names = set(z.namelist())
        for sname, spart, state in _sheet_parts(z):
            if not spart or spart not in names:
                continue
            try:
                ws = wb[sname]
            except KeyError:
                continue
            cols, rows = _grid(ws)
            sroot = ET.fromstring(z.read(spart))
            srels = _rels(z, spart)
            dparts = []
            for node in sroot.findall("m:drawing", ns_m):
                rid = node.get("{%s}id" % _NS["r"])
                if rid and srels.get(rid):
                    dparts.append(srels[rid])
            for dpart in dparts:
                if dpart not in names:
                    continue
                drels = _rels(z, dpart)
                droot = ET.fromstring(z.read(dpart))
                for zi, anchor in enumerate(droot):
                    kind = _lname(anchor)
                    if kind not in ("twoCellAnchor", "oneCellAnchor", "absoluteAnchor"):
                        continue
                    frm = _cell_ref(anchor, "from")
                    to = _cell_ref(anchor, "to")
                    ext = anchor.find("xdr:ext", _NS)
                    pos = anchor.find("xdr:pos", _NS)
                    if frm is not None:
                        x0 = _edge(cols, frm[0]) + frm[1]
                        y0 = _edge(rows, frm[2]) + frm[3]
                    elif pos is not None:
                        x0, y0 = _int0(pos, "x"), _int0(pos, "y")
                    else:
                        x0 = y0 = 0
                    if kind == "twoCellAnchor" and to is not None:
                        x1 = _edge(cols, to[0]) + to[1]
                        y1 = _edge(rows, to[2]) + to[3]
                    elif ext is not None:
                        # oneCellAnchor / absoluteAnchor 的尺寸取锚点自身的 xdr:ext；
                        # graphicFrame 内层 a:ext 常为 0，不能用来判尺寸
                        x1 = x0 + _int0(ext, "cx")
                        y1 = y0 + _int0(ext, "cy")
                    else:
                        x1, y1 = x0, y0
                    gf = anchor.find("xdr:graphicFrame", _NS)
                    cnv = None
                    if gf is not None:
                        cnv = gf.find("xdr:nvGraphicFramePr/xdr:cNvPr", _NS)
                    x1 = max(x1, x0)
                    y1 = max(y1, y0)
                    if kind == "twoCellAnchor" and to is not None:
                        # 右边界落在 to.col 内；colOff 为 0 时右边界正好压在列左边线上
                        right_col = to[0] - 1 if to[1] == 0 else to[0]
                    else:
                        right_col = _locate(cols, max(x1 - 1, x0))[0]
                    rec = {
                        "sheet": sname, "sheet_state": state, "z": zi,
                        "anchor": kind, "from": frm, "to": to,
                        "x0": x0, "y0": y0, "x1": x1, "y1": y1,
                        "cx": x1 - x0, "cy": y1 - y0,
                        "top": _row_at(rows, y0), "bottom": _row_at(rows, y1),
                        "left_col": frm[0] if frm else _locate(cols, x0)[0],
                        "right_col": right_col,
                        "hidden": bool(cnv is not None and (cnv.get("hidden") in ("1", "true"))),
                        "name": (cnv.get("name") if cnv is not None else "") or "",
                    }
                    cref = None
                    if gf is not None:
                        gd = gf.find("a:graphic/a:graphicData", _NS)
                        if gd is not None:
                            cnode = gd.find("c:chart", _NS)
                            if cnode is not None:
                                cref = drels.get(cnode.get("{%s}id" % _NS["r"]) or "")
                    if cref and cref in names:
                        rec["chart"] = _parse_chart(z, cref)
                        charts.append(rec)
                    else:
                        sub = None
                        for tag in ("pic", "sp", "grpSp", "cxnSp", "graphicFrame"):
                            if anchor.find("xdr:%s" % tag, _NS) is not None:
                                sub = tag
                                break
                        rec["obj"] = sub or "?"
                        rec["text"] = _texts(anchor)
                        shapes.append(rec)
    charts.sort(key=lambda r: (r["sheet"], r["z"]))
    return charts, shapes
def _resolve_series(wb, rec):
    """把系列引用解析到真实单元格：得到分类文字、数值、引用行列可见性。
    draws 表示 Excel 里能画出扇形（允许用缓存/字面量）；live 表示按细则注 3
    可用于加分判定（数据必须来自可见的单元格区域引用）。"""
    ch = rec["chart"]
    for s in ch["sers"]:
        s["cat_cells"] = _ref_cells(wb, s["cat_ref"]) if s["cat_ref"] else None
        s["val_cells"] = _ref_cells(wb, s["val_ref"]) if s["val_ref"] else None
        s["cat_vals"] = ([("" if c.value is None else str(c.value).strip())
                          for c in s["cat_cells"]] if s["cat_cells"] else [])
        s["val_nums"] = ([_as_num(c.value) for c in s["val_cells"]]
                         if s["val_cells"] else [])
        s["val_plot"] = ([_plot_num(c.value) for c in s["val_cells"]]
                         if s["val_cells"] else [])
        vis = True
        for cells in (s["cat_cells"], s["val_cells"]):
            for c in (cells or []):
                if _col_width(c.parent, c.column) <= 0 or _row_height(c.parent, c.row) <= 0:
                    vis = False
        s["ref_visible"] = vis
        plot_vis_only = bool(ch.get("plot_vis_only", True))
        # 逐点判断：只有“真数字 + 所在行列没被隐藏”的点才会画出扇形
        pos_cell = []
        for c, x in zip(s["val_cells"] or [], s["val_plot"]):
            if x is None or x <= 0:
                continue
            shown = (_col_width(c.parent, c.column) > 0
                     and _row_height(c.parent, c.row) > 0)
            if shown or not plot_vis_only:
                pos_cell.append(x)
        cache_nums = [_as_num(v) for v in s["val_cache"]]
        pos_cache = [v for v in cache_nums if v is not None and v > 0]
        s["live"] = bool(s["cat_ref"] and s["val_ref"] and s["val_cells"]
                         and pos_cell and (vis or not plot_vis_only))
        # 注3：引用能落到真实单元格时，numCache 里的残留数值不能替单元格“顶账”
        if s["val_cells"]:
            s["draws"] = bool(pos_cell)
        else:
            s["draws"] = bool(pos_cache and (s["val_lit"] or not s["val_ref"]))
        s["n"] = len(s["val_cells"]) if s["val_cells"] else len(s["val_cache"])
    pies = [s for s in ch["sers"] if s.get("kind") == _PIE2D]
    ch["pie_sers"] = pies
    # Excel 只渲染饼图的第一个系列，判定一律以第一个系列为准
    ch["main"] = pies[0] if pies else (ch["sers"][0] if ch["sers"] else None)
    rec["visible"] = bool(rec["sheet_state"] == "visible" and not rec["hidden"]
                          and rec["cx"] > 0 and rec["cy"] > 0)
    rec["draws"] = bool(ch["main"] is not None and ch["main"]["draws"]
                        and not ch["raw_err"])
    rec["live"] = bool(ch["main"] is not None and ch["main"]["live"])


def _occluded(rec, others):
    """被同表内层叠在其上方（文档序更靠后）的对象完整覆盖 -> 视为不可见。"""
    for o in others:
        if o is rec or o["sheet"] != rec["sheet"] or o["z"] <= rec["z"]:
            continue
        if (o["x0"] <= rec["x0"] and o["y0"] <= rec["y0"]
                and o["x1"] >= rec["x1"] and o["y1"] >= rec["y1"]
                and o["cx"] > 0 and o["cy"] > 0):
            return True
    return False


def _overlap(a0, a1, b0, b1):
    return max(0.0, min(a1, b1) - max(a0, b0))


def _assign_questions(charts, data_sheet):
    """给每张图表定题号：优先按引用区域落在哪道题的选项行，其次标题里的“第N题”，
    最后按纵向位置与题目区域的重合度。"""
    for rec in charts:
        ch = rec["chart"]
        q, how = None, ""
        best = 0.0
        for s in ch["sers"]:
            for ref in (s.get("cat_ref"), s.get("val_ref")):
                info = _parse_ref(ref) if ref else None
                if info is None:
                    continue
                sheet, _c1, r1, _c2, r2 = info
                if sheet.lower() != data_sheet.lower():
                    continue
                for qn, item in _QUESTIONS.items():
                    ov = _overlap(r1, r2 + 1, item[1], item[2])
                    if ov > best:
                        best, q, how = ov, qn, "ref"
        if q is None:
            m = re.search(r"第\s*(\d+)\s*题", ch.get("title") or "")
            if m and int(m.group(1)) in _QUESTIONS:
                q, how = int(m.group(1)), "title"
        if q is None:
            best = 0.0
            for qn, item in _QUESTIONS.items():
                ov = _overlap(rec["top"], rec["bottom"], item[0], item[2] + 1)
                if ov > best:
                    best, q, how = ov, qn, "pos"
        rec["q"] = q
        rec["q_how"] = how
def evaluate(dir_path):
    """按 §2.2 约定，返回结构化 dict；不 print、不改 sys.stdout、不 sys.exit。
    入参为脚本所在目录，脚本自行在该目录内定位被评估的工作簿。

    返回信息约定：
      status="error"（目录不存在、目录中没有交付文件、脚本自身异常）表示这份作业
      根本没评上，此时 dim1_pass 为空串（未判定）、dim1_reason 为空串、
      dim2_items 为 []、total_score 与 max_score 均为 0；
      维度一未通过时 status 仍为 "ok"，dim1_pass 为 False、dim1_reason 写错误摘要、
      dim2_items 为 []（未真正逐条评过）、total_score 为 0、max_score 为满分；
      维度一通过时 dim1_pass 为 True、dim1_reason 保持空串。
      dim2_items 每条只含 rule / max_delta / delta / hit 四个键。"""
    result = {
        "id": SCRIPT_ID,
        "file_name": "",
        "status": "ok",
        "error": None,
        "dim1_pass": "",
        "dim1_reason": "",
        "dim2_items": [],
        "total_score": 0,
        "max_score": 0,
    }
    try:
        if not dir_path or not os.path.isdir(dir_path):
            result["status"] = "error"
            result["error"] = "目录不存在：%s" % dir_path
            return result
        file_path, rejected = _pick_file(dir_path)
        if file_path is None:
            if rejected:
                # 目录里有交付物但格式不被接受：已有交付动作，交给维度一记 0
                result["file_name"] = os.path.basename(rejected)
                return _dim1_fail(
                    result,
                    "文件格式为 %s，维度一只接受可编辑的 .xlsx/.xlsm 工作簿"
                    % os.path.splitext(rejected)[1].lower())
            # 目录里一个交付物都没有：这份作业没评上，而不是维度一判不通过
            result["status"] = "error"
            result["error"] = "目录中未找到 .xlsx/.xlsm 文件：%s" % dir_path
            return result
        result["file_name"] = os.path.basename(file_path)
        return _evaluate_impl(file_path, result)
    except Exception as exc:
        # 脚本异常同样视为“没评上”：维度一未判定，两项分数归零。
        result["status"] = "error"
        result["error"] = "%s: %s" % (type(exc).__name__, exc)
        result["dim1_pass"] = ""
        result["dim1_reason"] = ""
        result["dim2_items"] = []
        result["total_score"] = 0
        result["max_score"] = 0
        return result


def _dim1_fail(result, reason):
    """维度一未通过：完成度总分记 0，不再逐条罗列（未真正逐条评过）。"""
    result["dim1_pass"] = False
    result["dim1_reason"] = reason or "维度一未通过（未给出具体原因）"
    result["dim2_items"] = []
    result["total_score"] = 0
    result["max_score"] = MAX_SCORE_TOTAL
    return result


def _norm(text):
    """去掉所有空白，用于标题/题干包含关系比较。"""
    return re.sub(r"\s+", "", text or "")


def _strip_tag(text):
    """去掉题干末尾的 [单选题]/[多选题] 标记。"""
    return re.sub(r"\[[^\[\]]*选题\]\s*$", "", (text or "").strip()).strip()


def _txt(ws, row, col=1):
    v = ws.cell(row=row, column=col).value
    return "" if v is None else str(v).strip()


def _vals_match(vals, counts):
    """数值引用内容是否等于该题“小计”，或等于与之等价的“比例”（细则注 1）。"""
    if len(vals) != len(counts) or any(v is None for v in vals):
        return False
    if all(abs(float(v) - float(c)) < 1e-6 for v, c in zip(vals, counts)):
        return True
    sv, sc = float(sum(vals)), float(sum(counts))
    if sv <= 0 or sc <= 0:
        return False
    # 比例为四舍五入到两位百分数的近似值，按占比比较并留 0.5 个百分点余量
    return all(abs(v / sv - c / sc) <= 0.005 for v, c in zip(vals, counts))
def _evaluate_impl(file_path, result):
    results = []

    def add(hit, delta, rule):
        """记一条维度二结果：命中记 delta，否则记 0（顺序须与 _MAX_DELTAS_ORDER 一致）。"""
        results.append("%+d：%s" % ((delta if hit else 0), rule))

    # ---------- 维度一：只判文件格式 ----------
    # 后缀已在 evaluate 里限定为 .xlsx/.xlsm，此处工作簿又能正常打开、无需修复，
    # 维度一即通过。图表是否可见、能否渲染出扇形一律交给维度二对应加分项。
    try:
        wbf = load_workbook(file_path, data_only=False)   # 看公式
        wb = load_workbook(file_path, data_only=True)     # 看显示值
    except Exception as exc:
        return _dim1_fail(result, "工作簿无法正常打开（%s: %s），可能已损坏或需修复"
                                  % (type(exc).__name__, exc))
    result["dim1_pass"] = True
    # 约定：维度一通过时 dim1_reason 保持空串，只有未通过才写错误摘要。
    result["dim1_reason"] = ""

    # ---------- 选定数据表 ----------
    def content_cells(ws):
        n = 0
        for row in ws.iter_rows():
            for c in row:
                if c.value is not None and str(c.value).strip() != "":
                    n += 1
        return n

    counts = {ws.title: content_cells(ws) for ws in wb.worksheets}
    data_ws = None
    for ws in wb.worksheets:
        if ws.title.strip().lower() == "sheet1" and counts[ws.title] > 0:
            data_ws = ws
            break
    if data_ws is None and wb.worksheets:
        data_ws = max(wb.worksheets, key=lambda w: counts[w.title])
    if data_ws is None:
        # 工作簿里没有普通工作表（例如只有图表工作表）：维度一按格式判过，维度二
        # 用一张空表继续走完全部条目，加分自然全不命中。
        data_ws = wb.create_sheet("sheet1")
        wbf.create_sheet("sheet1")
        counts[data_ws.title] = 0
    ws = data_ws
    wsf = wbf[ws.title]

    # ---------- 图表与浮动对象 ----------
    charts, shapes = _collect_drawings(file_path, wb)
    for rec in charts:
        _resolve_series(wb, rec)
    _assign_questions(charts, ws.title)
    for rec in charts:
        rec["occluded"] = _occluded(rec, charts + shapes)
        rec["ok"] = bool(rec["visible"] and rec["draws"] and not rec["occluded"])
    by_q = {}
    for rec in charts:
        by_q.setdefault(rec["q"], []).append(rec)

    # ---------- 维度二：加分项 ----------
    # “均/全部”类条目按“已插入的每一张图表都要满足”判定；张数与一题一图的对应关系
    # 单独由第 01 项负责，避免少插 1 张就把 12 个条目一起判掉。
    pool = charts
    cover = bool(len(charts) == 22 and all(len(by_q.get(q, [])) == 1 for q in _QNOS))

    def main_of(rec):
        return rec["chart"]["main"]

    def exp(rec):
        return _QUESTIONS.get(rec["q"])

    hit = bool(cover and all(r["ok"] for r in pool)
               and all(r["sheet"] == ws.title for r in pool)
               and all(r["top"] >= 24.99 for r in pool))
    add(hit, 5, "图表总数为22张，分别对应第4题至第25题，每题各1张，无遗漏、无重复，"
                "且22张图表均为可见状态、图形正常渲染出扇形；第1题至第3题所在的"
                "第1行至第24行范围内没有图表")

    hit = bool(pool and all(r["chart"]["kinds"] == [_PIE2D] for r in pool))
    add(hit, 5, "22张图表全部为二维饼图，不是三维饼图、圆环图、复合饼图，"
                "也不是柱形图、条形图、折线图等其他类型")

    def color_ok(rec):
        s = main_of(rec)
        if s is None:
            return False
        if not rec["draws"]:
            return False          # 一个扇形都画不出来，谈不上颜色能区分选项
        n = s["n"] or len(s["cat_vals"]) or len(s["cat_cache"])
        keys = {}
        for idx, k in s["dpts"]:
            if k:
                keys[idx] = k
        if keys:
            if len(set(keys.values())) != len(keys):
                return False
            if len(keys) >= max(n, 1):
                return True
            return s.get("vary") is not False
        if s["ser_fill"]:
            return False          # 整个系列一种填充色，饼图变成单色
        return s.get("vary") is not False

    hit = bool(pool and all(color_ok(r) for r in pool))
    add(hit, 5, "22张饼图内各选项扇形的填充色均互不相同，可逐一区分该题的不同选项，"
                "没有整张为单一颜色的饼图")

    def dlbl_ok(rec):
        s = main_of(rec)
        d = s.get("dlbls") if s else None
        return bool(d and not d["delete"] and d["percent"])

    hit = bool(pool and all(dlbl_ok(r) for r in pool))
    add(hit, 5, "22张饼图均显示数据标签，标注各扇形所占百分比")

    def legend_ok(rec):
        leg = rec["chart"]["legend"]
        s = main_of(rec)
        it = exp(rec)
        if leg is None or s is None or it is None or leg["hidden_entries"]:
            return False
        return s["cat_vals"] == list(it[3])

    hit = bool(pool and all(legend_ok(r) for r in pool))
    add(hit, 5, "22张饼图均带有图例，图例条目文字为该题各选项名称，与扇形一一对应")

    def title_ok(rec):
        ch = rec["chart"]
        it = exp(rec)
        if it is None or not ch["has_title"]:
            return False
        raw = ch["title"] or ""
        if not _norm(raw) and ch["title_ref"]:
            cells = _ref_cells(wb, ch["title_ref"]) or []
            raw = "".join("" if c.value is None else str(c.value) for c in cells)
        t = _norm(raw)
        stem = _strip_tag(_txt(ws, it[0])) or ("第%d题" % rec["q"])
        body = _norm(re.sub(r"^第\s*\d+\s*题\s*[:：]?\s*", "", stem))
        return bool(re.search(r"第\s*%d\s*题" % rec["q"], t) and body and body in t)

    hit = bool(pool and all(title_ok(r) for r in pool))
    add(hit, 5, "22张饼图均有标题，标题同时包含题号“第N题”与该题题干文字，"
                "例如第4题的图表标题为“第4题：你对学校图书角、阅读区等学习空间的了解程度如何”")

    hit = bool(pool and all(r["left_col"] >= _COL_LEFT_MIN for r in pool))
    add(hit, 5, "22张图表的左边界均位于D列或其右侧，A列至C列的题干、选项、小计、比例数据"
                "均可完整看到，无被图表覆盖的单元格")
    _ROW_TOL = 0.35     # 允许图表上下边界压进相邻题目不超过 0.35 行

    def place_ok():
        if not pool:
            return False
        for rec in pool:
            it = exp(rec)
            if it is None or rec["sheet"] != ws.title:
                return False
            if rec["left_col"] < _COL_LEFT_MIN:
                return False
            if _overlap(rec["top"], rec["bottom"], it[0], it[2] + 1) <= 0:
                return False
            for qn, it2 in _QUESTIONS.items():
                if qn == rec["q"]:
                    continue
                if _overlap(rec["top"], rec["bottom"], it2[0], it2[2] + 1) > _ROW_TOL:
                    return False
        for i, a in enumerate(pool):
            for b in pool[i + 1:]:
                if a["sheet"] != b["sheet"]:
                    continue
                if (_overlap(a["x0"], a["x1"], b["x0"], b["x1"]) > 0
                        and _overlap(a["top"], a["bottom"], b["top"], b["bottom"]) > _ROW_TOL):
                    return False
        return True

    add(place_ok(), 5, "22张图表均插在其对应题目的右侧，图表的纵向范围与该题题干行至"
                       "“本题有效填写人次”行所在的行区间有重合，不串到其他题目的位置上，"
                       "且22张图表两两之间不相互堆叠遮挡")

    def size_ok(rec):
        h = rec["cy"] / float(_EMU_PT)
        w = rec["cx"] / float(_EMU_PT)
        if h <= 0:
            return False
        return h >= _MIN_H_PT - 0.5 and (w / h) <= _MAX_RATIO + 0.02

    hit = bool(pool and all(size_ok(r) for r in pool))
    add(hit, 3, "22张饼图的图形高度均不小于120磅，宽高比均不超过2.5:1，"
                "扇形呈正圆且大小足以辨认，不呈扁长条状")

    def bind_ok(rec):
        s = main_of(rec)
        if s is None or not rec["live"]:
            return False
        if s["cat_lit"] or s["val_lit"] or not s["cat_ref"] or not s["val_ref"]:
            return False
        ci, vi = _parse_ref(s["cat_ref"]), _parse_ref(s["val_ref"])
        if ci is None or vi is None:
            return False
        if ci[0].lower() != ws.title.lower() or vi[0].lower() != ws.title.lower():
            return False
        return bool(s["ref_visible"] and s["cat_cells"] and s["val_cells"])

    hit = bool(pool and all(bind_ok(r) for r in pool))
    add(hit, 3, "22张饼图的分类与数值均以单元格区域引用方式绑定，例如第4题的分类与数值"
                "分别引用 sheet1!$A$27:$A$30 与 sheet1!$B$27:$B$30，引用区域所在的列"
                "均为可见状态、扇形按引用数据正常绘制，修改表内数字时图表可同步更新，"
                "不是写死的静态数值")

    def order_ok(rec):
        s = main_of(rec)
        it = exp(rec)
        if s is None or it is None or not rec["live"]:
            return False
        ci, vi = _parse_ref(s["cat_ref"]), _parse_ref(s["val_ref"])
        if ci is None or vi is None or ci[2] > ci[4] or vi[2] > vi[4]:
            return False
        if (ci[2], ci[4]) != (it[1], it[2] - 1) or (vi[2], vi[4]) != (it[1], it[2] - 1):
            return False
        return s["cat_vals"] == list(it[3])

    hit = bool(pool and all(order_ok(r) for r in pool))
    add(hit, 3, "22张饼图的扇形排列顺序均与该题选项在表格中自上而下的顺序一致")

    hit = bool(pool and all(all(z >= _MIN_FONT_SZ - 1e-6 for z in r["chart"]["font_szs"])
                            for r in pool))
    add(hit, 1, "22张图表的标题、图例、数据标签字号均不小于8磅，可正常辨认")

    hit = bool(pool and all(r["right_col"] <= _COL_RIGHT_MAX for r in pool))
    add(hit, 1, "22张图表的右边界均不超出K列")

    def per_q_ok(qn):
        it = _QUESTIONS[qn]
        recs = by_q.get(qn, [])
        if len(recs) != 1:
            return False
        rec = recs[0]
        s = main_of(rec)
        if s is None or not rec["ok"] or not rec["live"]:
            return False
        if rec["chart"]["kinds"] != [_PIE2D]:
            return False
        lo, hi = it[1], it[2] - 1          # 选项行区间，不含“本题有效填写人次”行
        for ref in (s["cat_ref"], s["val_ref"]):
            info = _parse_ref(ref)
            if info is None or info[0].lower() != ws.title.lower():
                return False
            if info[2] < lo or info[4] > hi:
                return False
        if s["n"] != len(it[3]):
            return False
        if s["cat_vals"] != list(it[3]):
            return False
        return _vals_match(s["val_nums"], it[4])

    for qn in _QNOS:
        it = _QUESTIONS[qn]
        add(per_q_ok(qn), 1,
            "第%d题右侧有1张饼图，扇形数为%d，分类引用区域的内容依次为%s，"
            "数值引用区域的内容依次为%s，不含“本题有效填写人次/168”对应的扇形"
            % (qn, len(it[3]), "".join("“%s”" % o for o in it[3]),
               "、".join(str(x) for x in it[4])))
    # ---------- 维度二：减分项（不满足即扣分） ----------
    def cell(r, c):
        return ws.cell(row=r, column=c)

    def has_content(r, c):
        for w in (ws, wsf):
            v = w.cell(row=r, column=c).value
            if v is not None and str(v).strip() != "":
                return True
        return False

    def struct_ok():
        for i, (stem, hdr, o1, tot) in enumerate(_BLOCKS):
            a = _txt(ws, stem)
            if not re.match(r"^第\s*%d\s*题\s*[:：]" % (i + 1), a):
                return False
            if not re.search(r"\[[^\[\]]*选题\]$", a):
                return False
            if tuple(_txt(ws, hdr, c) for c in (1, 2, 3)) != _HEADER_TEXTS:
                return False
            if tot - o1 < 2:
                return False
            for r in range(o1, tot):
                if not all(_txt(ws, r, c) for c in (1, 2, 3)):
                    return False
            if _txt(ws, tot, 1) != _TOTAL_LABEL:
                return False
            if _as_num(ws.cell(row=tot, column=2).value) != 168.0:
                return False
        if _strip_tag(_txt(ws, 1)) != _EXPECTED_A1:
            return False
        if _strip_tag(_txt(ws, 25)) != _EXPECTED_A25:
            return False
        if _strip_tag(_txt(ws, 200)) != _EXPECTED_A200:
            return False
        return all(_txt(ws, r).endswith("[多选题]") for r in (75, 191))

    add(not struct_ok(), -3,
        "A列至C列的内容不满足以下结构：共25道题，每题依次为“第N题：题干[单选题]”或"
        "“第N题：题干[多选题]”标题行、“选项/小计/比例”表头行、各选项名称及其人数与"
        "百分比数据行、“本题有效填写人次/168/100%”行；其中A1为“第1题：你的年级阶段"
        "[单选题]”，A25为“第4题：你对学校图书角、阅读区等学习空间的了解程度如何"
        "[单选题]”，A200为“第25题：总的来说，你对当前校园学习活动安排满意吗[单选题]”，"
        "A75与A191的题干末尾为“[多选题]”")

    need_merges = set("A%d:C%d" % (r, r) for r in _STEM_ROWS)
    have_merges = set(str(m) for m in ws.merged_cells.ranges)
    add(not need_merges.issubset(have_merges), -3,
        "25个题干行不满足A列至C列跨列合并为一个单元格，合并区域依次为A1:C1、A7:C7、"
        "A16:C16、A25:C25、A33:C33、A42:C42、A50:C50、A58:C58、A67:C67、A75:C75、"
        "A84:C84、A92:C92、A101:C101、A109:C109、A117:C117、A125:C125、A133:C133、"
        "A141:C141、A149:C149、A157:C157、A165:C165、A174:C174、A182:C182、A191:C191、"
        "A200:C200")

    stem_fmt = all(_font_ok(cell(r, 1), bold=True, white=True)
                   and _fill_rgb(cell(r, 1)) == _FILL_STEM for r in _STEM_ROWS)
    add(not stem_fmt, -3,
        "25个题干行不满足字体为微软雅黑、10号、加粗、白色字，且单元格底色为深青色RGB 2F6F73")

    hdr_fmt = all(_font_ok(cell(r, c), bold=True) and _fill_rgb(cell(r, c)) == _FILL_HEADER
                  for _s, r, _o, _t in _BLOCKS for c in (1, 2, 3))
    add(not hdr_fmt, -3,
        "各题“选项”“小计”“比例”表头行不满足字体为微软雅黑、10号、加粗，"
        "且单元格底色为浅青色RGB DCECEF")

    def data_fmt_ok():
        for _s, _h, o1, tot in _BLOCKS:
            prev = None
            for r in range(o1, tot):
                fills = set()
                for c in (1, 2, 3):
                    if not _font_ok(cell(r, c), bold=False):
                        return False
                    fills.add(_fill_rgb(cell(r, c)))
                if len(fills) != 1:
                    return False
                cur = fills.pop()
                if cur not in _FILL_DATA or cur == prev:
                    return False
                prev = cur
        return True

    add(not data_fmt_ok(), -3,
        "各题选项数据行不满足字体为微软雅黑、10号、不加粗，"
        "且底色按行在米白色RGB FFFDF6与浅米色RGB F7F3E8之间交替")

    def align_ok():
        for _s, hdr, o1, tot in _BLOCKS:
            for c in (1, 2, 3):
                if not _align_ok(cell(hdr, c), _H_CENTER, _V_CENTER, True):
                    return False
            for r in list(range(o1, tot)) + [tot]:
                if not _align_ok(cell(r, 1), _H_LEFT, _V_CENTER, True):
                    return False
                for c in (2, 3):
                    if not _align_ok(cell(r, c), _H_CENTER, _V_CENTER, True):
                        return False
        return True

    add(not align_ok(), -3,
        "单元格对齐方式不满足：各题“选项”“小计”“比例”表头行的A至C列均为水平居中、"
        "垂直居中、自动换行；各题选项数据行与“本题有效填写人次”行的A列为水平左对齐、"
        "垂直居中、自动换行、B列与C列为水平居中、垂直居中、自动换行")
    total_fmt = all(_font_ok(cell(r, c), bold=True) and _fill_rgb(cell(r, c)) == _FILL_TOTAL
                    for r in _TOTAL_ROWS for c in (1, 2, 3))
    add(not total_fmt, -1,
        "第5、14、23、31、40、48、56、65、73、82、90、99、107、115、123、131、139、147、"
        "155、163、172、180、189、198、207行共25个“本题有效填写人次”行不满足字体为"
        "微软雅黑、10号、加粗，且单元格底色为浅绿色RGB E7F2E4")

    def pct_ok():
        if _display_text(cell(3, 3)) != "48.81%" or _display_text(cell(4, 3)) != "51.19%":
            return False
        for _s, _h, o1, tot in _BLOCKS:
            for r in range(o1, tot):
                if not re.fullmatch(r"\d+\.\d{2}%", _display_text(cell(r, 3))):
                    return False
            # “100%”与等值的“100.00%”都算显示为 100%
            if _display_text(cell(tot, 3)) not in ("100%", "100.00%"):
                return False
        return True

    add(not pct_ok(), -1,
        "C列“比例”数据不满足以保留两位小数的百分比形式显示，其中C3显示为“48.81%”、"
        "C4显示为“51.19%”，且“本题有效填写人次”行的C列显示为“100%”")

    def extent_ok():
        max_r = max_c = 0
        for w in (ws, wsf):
            for row in w.iter_rows():
                for c in row:
                    if c.value is not None and str(c.value).strip() != "":
                        max_r = max(max_r, c.row)
                        max_c = max(max_c, c.column)
        if max_r != _MAX_ROW or max_c > 3:
            return False
        return not any(has_content(r, c) for r in _BLANK_ROWS for c in (1, 2, 3))

    add(not extent_ok(), -1,
        "表格数据不满足只占用A、B、C三列、共207行，且第6、15、24、32、41、49、57、66、74、"
        "83、91、100、108、116、124、132、140、148、156、164、173、181、190、199行的"
        "A至C列为空白行")

    cols_visible = all(not (ws.column_dimensions.get(get_column_letter(c)) is not None
                            and ws.column_dimensions[get_column_letter(c)].hidden)
                       for c in (1, 2, 3))
    add(not cols_visible, -1, "A、B、C三列不满足均为可见状态，存在被隐藏的列")

    width_ok = all(abs(_col_width(ws, c) - _COL_W_TARGET) <= _COL_W_TOL
                   for c in range(1, _COL_RIGHT_MAX + 2))
    add(not width_ok, -1, "A列至K列列宽不满足均为9字符，允许±1字符误差")

    def right_clean():
        for w in (ws, wsf):
            for row in w.iter_rows(min_col=4):
                for c in row:
                    if c.value is not None and str(c.value).strip() != "":
                        return False
        for sp in shapes:
            if sp["sheet"] == ws.title and _norm(sp.get("text") or ""):
                return False
        return True

    add(not right_clean(), -1,
        "D列及其右侧的单元格不满足全部为空白，除图表外还存在文字、公式、辅助计算列或"
        "汇总统计数据")

    data_sheets = [t for t, n in counts.items() if n > 0]
    add(not (len(data_sheets) == 1 and data_sheets[0].strip().lower() == "sheet1"), -1,
        "工作簿不满足只有1个含数据的工作表，且该工作表名称不满足为“sheet1”")

    found_stems = [r for r in range(1, _MAX_ROW + 1)
                   if re.match(r"^第\s*\d+\s*题\s*[:：]", _txt(ws, r))]
    add(found_stems != _STEM_ROWS, -1,
        "25个题干行不满足分别位于第1、7、16、25、33、42、50、58、67、75、84、92、101、109、"
        "117、125、133、141、149、157、165、174、182、191、200行")

    def plain_ok():
        if ws.auto_filter is not None and ws.auto_filter.ref:
            return False
        if ws.freeze_panes:
            return False
        try:
            if len(ws.conditional_formatting._cf_rules) > 0:
                return False
        except Exception:
            pass
        for r in range(1, _MAX_ROW + 1):
            for c in (1, 2, 3):
                bd = cell(r, c).border
                if bd is None:
                    continue
                for side in ("left", "right", "top", "bottom", "diagonal"):
                    s = getattr(bd, side, None)
                    if s is not None and s.style:
                        return False
        return True

    add(not plain_ok(), -1,
        "工作表不满足以下任意一条：无自动筛选下拉箭头、无冻结窗格、无条件格式、"
        "A1:C207区域无单元格边框线")
    # ---------- 汇总：把 results 反解为结构化 dim2_items ----------
    if len(results) != len(_MAX_DELTAS_ORDER):
        raise RuntimeError("dim2 项数不一致：results=%d vs order=%d"
                           % (len(results), len(_MAX_DELTAS_ORDER)))
    dim2_items = []
    total_score = 0
    for raw, max_delta in zip(results, _MAX_DELTAS_ORDER):
        head, _sep, rule = raw.partition("：")
        try:
            delta = int(head.replace("+", "").strip())
        except ValueError:
            delta = 0
        hit = (delta != 0)
        dim2_items.append({
            "rule": rule,
            "max_delta": max_delta,
            "delta": max_delta if hit else 0,
            "hit": hit,
        })
        total_score += (max_delta if hit else 0)
    result["dim2_items"] = dim2_items
    result["total_score"] = max(0, total_score)   # 加分不足以抵扣分时兜底为 0
    result["max_score"] = MAX_SCORE_TOTAL
    return result
_BLOCKS = [(s, s + 1, s + 2, t) for s, t in zip(_STEM_ROWS, _TOTAL_ROWS)]


if __name__ == "__main__":
    # 本地调试入口：默认评估脚本所在目录，也可传入待评估目录
    _dir = sys.argv[1] if len(sys.argv) > 1 else os.path.dirname(os.path.abspath(__file__))
    _out = evaluate(_dir)
    print(json.dumps(_out, ensure_ascii=False, indent=2))