#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# pyright: reportUnknownMemberType=false, reportUnknownVariableType=false, reportUnknownArgumentType=false, reportUnknownParameterType=false, reportMissingParameterType=false, reportPrivateUsage=false, reportUnknownLambdaType=false, reportGeneralTypeIssues=false

import os
import re
import sys
import json
import math
import datetime

from openpyxl import load_workbook
from openpyxl.utils import get_column_letter, column_index_from_string

# 说明：按接口统一约定，本模块不修改全局 sys.stdout，也不 print 主结果。
# 主结果统一由 evaluate(dir_path) -> dict 返回；入参为“脚本所在目录的路径”，
# 由脚本自己在该目录内定位并打开被评估的工作簿。本文件仅在 __main__ 下作为
# 本地调试入口，才把 dict 序列化为 JSON 打印。
#
# 任务：在“云谷智慧农业示范园建设项目报价表”工作表里，把占一页的截图按其内容
# 还原成两张单元格清单表（“五、环控与照明物资清单”“六、物联网设备预算清单”）
# 并删掉截图；把表内字体、表头字体、表名称字体分别统一；调整行高让所有文字完整
# 显示、相似表格行高一致；打印预览中 A:J 十列同处一页宽、每张表格不跨两页。
#
# 维度一只判文件格式：交付物是否为可编辑的 .xlsx/.xlsm 工作簿、能否正常打开。
# 维度二既有加分项（指令要求做的事，做到才给分），也有减分项（指令没要求、
# 属于把原有排版改坏的地方，命中即扣分）。两张新表一律靠 A 列文本锚定行号
# （表名称行 → 下一行为表头行 → “… 小计”行 → 二者之间序号 1..5 的明细行），
# 不写死表格所在行，避免整表上下平移导致误判。

SCRIPT_ID = "092"
# 维度二加分项 max_delta 之和；用于维度一未通过 / 脚本异常时兜底
MAX_SCORE_TOTAL = 100
# 维度二减分项 max_delta 之和（负值），仅用于自检
MAX_DEDUCT_TOTAL = -33

# 维度二所有评分项按“结果追加顺序”的 max_delta 列表；末端把 results 里的
# “+N：...”/“-N：...”字符串反解为结构化 dim2_items 时用。顺序须与
# _evaluate_impl 中 add(...) 的顺序严格一致；共 47 项，
# 正项累加 = MAX_SCORE_TOTAL，负项累加 = MAX_DEDUCT_TOTAL。
_MAX_DELTAS_ORDER = [
    3,   # 01 无嵌入图片，第 69–88 行以单元格文本与数值承载清单内容
    5,   # 02 “五、环控与照明物资清单”表名称单元格格式与整行合并居中
    3,   # 03 “五”表表头行 10 列文本与字体底色
    5,   # 04 “五”表 5 行明细的序号与项目名称
    5,   # 05 “五”表项目特征描述
    3,   # 06 “五”表工程量与单位
    3,   # 07 “五”表综合单价
    3,   # 08 “五”表合价
    3,   # 09 “五”表施工/供应说明与使用部位
    3,   # 10 “五”表小计行文本、G 列金额、加粗与底色
    5,   # 11 “六、物联网设备预算清单”表名称单元格格式与整行合并居中
    3,   # 12 “六”表表头行 10 列文本与字体底色
    5,   # 13 “六”表 5 行明细的序号与项目名称
    5,   # 14 “六”表项目特征描述
    3,   # 15 “六”表工程量与单位
    3,   # 16 “六”表综合单价
    3,   # 17 “六”表合价
    3,   # 18 “六”表施工/供应说明与使用部位
    3,   # 19 “六”表小计行文本、G 列金额、加粗与底色
    5,   # 20 G76/G86 小计结果与 F122 总计取值
    1,   # 21 两张新表各行 A:J 细实线边框
    3,   # 22 C 列若干区域字体统一为微软雅黑
    3,   # 23 D/E/F/G/H 列若干区域字体统一为微软雅黑
    1,   # 24 I34:I37、G62:G65 字体统一为微软雅黑
    3,   # 25 两个新表名称与其余 7 个表名称同为 11 号
    3,   # 26 第 22 行行高≥90 磅且说明文字完整显示
    3,   # 27 全表含内容单元格行高≥1.3 倍字号、文字完整显示
    1,   # 28 各清单明细行行高互差≤1.5 磅
    5,   # 29 打印设置使 A:J 十列同处一页宽
    3,   # 30 每张清单表格不被拆到两页
    -3,  # 31 J 列“备注”指定区域应全部空白
    -3,  # 32 F2 报价日期 2026-05-25 且数字格式 yyyy-mm-dd
    -3,  # 33 H 列指定区域水平右对齐
    -3,  # 34 第 56 行“温室工程量清单合计”行字体与底色
    -3,  # 35 A26/A33/A40/A48 四个子组标题字体与底色
    -3,  # 36 “五”表表头行与小计行之间夹了整行空白
    -3,  # 37 第 31/38/46/54/96/116 行行高 48±1.5 磅
    -3,  # 38 12 个空白间隔行保持无内容/无填充/常规/微软雅黑
    -1,  # 39 D 列明细数量区域数字格式 #,##0.00
    -1,  # 40 H7:H20 占比列数字格式 0.0%
    -1,  # 41 A2:J3 填充色 F4FAF7
    -1,  # 42 A2/C2/E2/G2/I2/A3/C3/E3/G3/I3 字体颜色 1F4E3D
    -1,  # 43 A22 填充色 EAF3F8
    -1,  # 44 A1 主标题格式与 A1:J1 合并居中不换行
    -1,  # 45 C 列列宽 28±1.5 字符
    -1,  # 46 H 列列宽 12±1.5 字符
    -1,  # 47 第 1 行行高 36±1.5 磅
]
assert len(_MAX_DELTAS_ORDER) == 47
assert sum(d for d in _MAX_DELTAS_ORDER if d > 0) == MAX_SCORE_TOTAL
assert sum(d for d in _MAX_DELTAS_ORDER if d < 0) == MAX_DEDUCT_TOTAL

_EXT_OK = (".xlsx", ".xlsm")
# 维度一只接受 .xlsx/.xlsm；下列后缀视为格式不符，判维度一不通过并在
# dim1_reason 里写明，而不是抛异常。
_EXT_REJECT = (".xls", ".xlsb", ".csv", ".ods", ".et")

_SHEET = "智慧农业报价表"
_YAHEI = "微软雅黑"

# 两张新表共用的表头 10 列文本
_HEADER = ["序号", "项目名称", "项目特征描述", "工程量", "单位",
           "综合单价（元）", "合价（元）", "施工/供应说明", "使用部位", "备注"]

# 两张新增清单的基准内容（即被替换掉的那页截图所承载的清单）
_T5 = {
    "title": "五、环控与照明物资清单",
    "names": ["环流风机", "湿帘纸芯", "LED补光灯", "遮阳幕布", "电动卷膜器"],
    "feats": ["低噪声轴流，防潮电机", "150mm厚高效蒸发湿帘", "植物生长光谱，防水等级IP65",
              "外白内黑节能幕，阻燃处理", "24V直流，带限位保护"],
    "qty": [36, 220, 168, 8600, 44],
    "unit": ["台", "㎡", "盏", "㎡", "套"],
    "price": [1850, 165, 680, 22, 960],
    "amount": [66600, 36300, 114240, 189200, 42240],
    "parts": ["温室内部", "湿帘墙", "育苗区", "顶部遮阳", "侧窗通风"],
    "subtotal": 448580,
}
_T6 = {
    "title": "六、物联网设备预算清单",
    "names": ["温湿度传感器", "CO₂传感器", "土壤/基质水分传感器", "LoRa采集网关", "生产大屏看板"],
    "feats": ["温度、湿度双测点，防水外壳", "NDIR原理，量程0-5000ppm", "电容式，带温度补偿",
              "工业级，支持断点续传", "55英寸显示屏，含支架与信号盒"],
    "qty": [28, 12, 24, 6, 2],
    "unit": ["支", "支", "支", "台", "套"],
    "price": [520, 1680, 780, 3900, 6800],
    "amount": [14560, 20160, 18720, 23400, 13600],
    "parts": ["栽培区", "生产区", "种植槽", "控制间", "运营中心"],
    "subtotal": 90440,
}

# A4 纸尺寸（英寸），用于按行高/列宽模拟打印分页
_A4_W_IN = 8.27
_A4_H_IN = 11.69


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
        # 名字里含“报价表”的优先，其次取体积最大的那份
        named = [p for p in cands if "报价表" in os.path.basename(p)]
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
    try:
        v = getattr(color, "rgb", None)
    except Exception:
        return None
    if not isinstance(v, str) or len(v) < 6:
        return None
    return v[-6:].upper()


def _text(v):
    return "" if v is None else str(v).strip()


def _num(v):
    """尽量把单元格值转成 float；转不了（含错误值字符串、公式串）返回 None。"""
    if isinstance(v, bool):
        return None
    if isinstance(v, (int, float)):
        return float(v)
    s = _text(v).replace(",", "")
    try:
        return float(s)
    except Exception:
        return None


def _same_num(v, target, tol=0.51):
    """数值比较：允许 0.5 以内误差（金额/工程量均为整数口径）。"""
    n = _num(v)
    return n is not None and abs(n - float(target)) <= tol


def _units(text):
    """内容长度（列宽字符数口径）：汉字/全角按 2 个字符，其余按 1 个字符。"""
    return sum(2 if ord(ch) > 0x2E80 else 1 for ch in str(text))


def _cells(spec):
    """"C27:C29,J121" -> ['C27','C28','C29','J121']；只支持同列纵向区间与单点。"""
    out = []
    for part in spec.split(","):
        part = part.strip()
        if not part:
            continue
        if ":" in part:
            a, b = part.split(":")
            m1 = re.match(r"([A-Z]+)(\d+)$", a.strip())
            m2 = re.match(r"([A-Z]+)(\d+)$", b.strip())
            if m1 is None or m2 is None:
                continue
            for r in range(int(m1.group(2)), int(m2.group(2)) + 1):
                out.append("%s%d" % (m1.group(1), r))
        else:
            out.append(part)
    return out


def _norm_fmt(fmt):
    """数字格式归一：去掉 Excel 的转义反斜杠与引号并转小写，便于比较。"""
    return re.sub(r'[\\"]', "", fmt or "").strip().lower()


def _refs_cover(formula, targets):
    """公式里的引用（单点或区间）是否覆盖全部目标单元格；不是公式则返回 False。"""
    if not isinstance(formula, str) or not formula.startswith("="):
        return False
    need = set(targets)
    pat = r"\$?([A-Z]{1,3})\$?(\d+)(?:\s*:\s*\$?([A-Z]{1,3})\$?(\d+))?"
    for m in re.finditer(pat, formula.upper()):
        i1 = column_index_from_string(m.group(1))
        r1 = int(m.group(2))
        i2 = column_index_from_string(m.group(3)) if m.group(3) else i1
        r2 = int(m.group(4)) if m.group(4) else r1
        lo_c, hi_c = min(i1, i2), max(i1, i2)
        lo_r, hi_r = min(r1, r2), max(r1, r2)
        for t in list(need):
            mt = re.match(r"([A-Z]+)(\d+)$", t)
            if mt is None:
                continue
            if lo_c <= column_index_from_string(mt.group(1)) <= hi_c \
                    and lo_r <= int(mt.group(2)) <= hi_r:
                need.discard(t)
    return not need


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
                # 目录里有交付物但格式不对：已有交付动作，交给维度一记 0
                result["file_name"] = os.path.basename(rejected)
                return _dim1_fail(
                    result,
                    "文件格式为 %s，维度一只接受可编辑的 .xlsx/.xlsm 工作簿"
                    % os.path.splitext(rejected)[1].lower())
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


def _evaluate_impl(file_path, result):
    results = []

    def add(hit, delta, rule):
        """记一条维度二结果（顺序须与 _MAX_DELTAS_ORDER 一致）。
        加分项 delta 为正、hit 表示“按指令做到了”；
        减分项 delta 为负、hit 表示“指令没要求却被改坏了”，命中即扣分。"""
        results.append("%+d：%s" % ((delta if hit else 0), rule))

    # ---------- 维度一：只判文件格式 ----------
    # 后缀已在 evaluate 里限定为 .xlsx/.xlsm，此处工作簿又能正常打开、无需修复，
    # 维度一即通过；表格内容是否做对一律交给维度二。
    try:
        wbf = load_workbook(file_path, data_only=False)   # 看公式与格式
        wb = load_workbook(file_path, data_only=True)     # 看公式的缓存结果
    except Exception as exc:
        return _dim1_fail(
            result,
            "工作簿无法正常打开（%s: %s），可能已损坏或需修复"
            % (type(exc).__name__, exc))
    result["dim1_pass"] = True
    # 约定：维度一通过时 dim1_reason 保持空串，只有未通过才写错误摘要。
    result["dim1_reason"] = ""

    # ---------- 选定报价表所在工作表 ----------
    def content_cells(w):
        n = 0
        for row in w.iter_rows():
            for c in row:
                if c.value is not None and str(c.value).strip() != "":
                    n += 1
        return n

    if _SHEET in wbf.sheetnames:
        ws = wbf[_SHEET]
    elif wbf.worksheets:
        ws = max(wbf.worksheets, key=content_cells)
    else:
        ws = wbf.create_sheet(_SHEET)
    wsv = wb[ws.title] if ws.title in wb.sheetnames else ws

    # ---------- 行高/列宽/合并区等版式基础量 ----------
    _DEF_W = float(ws.sheet_format.defaultColWidth or 8.43)
    _DEF_H = float(ws.sheet_format.defaultRowHeight or 15.0)
    # 表格主体范围：A:J 十列、第 1 行至第 126 行（第 124 行之后已无清单内容）
    _MAXR = max(126, ws.max_row)

    def cell(r, c):
        return ws.cell(row=r, column=c)

    def vcell(r, c):
        return wsv.cell(row=r, column=c)

    def col_w(c):
        cd = ws.column_dimensions.get(get_column_letter(c))
        w = cd.width if cd is not None else None
        return float(w) if w else _DEF_W

    def row_h(r):
        rd = ws.row_dimensions.get(r)
        h = rd.height if rd is not None else None
        return float(h) if h else _DEF_H

    # 合并区映射：坐标 -> 所属合并区标识 / 可用宽度 / 是否为左上锚点
    m_of, span_w, anchor = {}, {}, {}
    for mr in ws.merged_cells.ranges:
        wsum = sum(col_w(c) for c in range(mr.min_col, mr.max_col + 1))
        for r in range(mr.min_row, mr.max_row + 1):
            for c in range(mr.min_col, mr.max_col + 1):
                m_of[(r, c)] = str(mr)
                span_w[(r, c)] = wsum
                anchor[(r, c)] = (r == mr.min_row and c == mr.min_col)

    def is_shadow(r, c):
        """是否为合并区内的“非锚点”单元格：这类格子在办公软件里不显示自身格式。"""
        return (r, c) in m_of and not anchor[(r, c)]

    def row_full_merged(r, min_span=10):
        """该行是否存在自 A 列起、至少跨 min_span 列的整行合并区。"""
        for mr in ws.merged_cells.ranges:
            if (mr.min_row == r and mr.max_row == r and mr.min_col == 1
                    and mr.max_col >= min_span):
                return True
        return False

    def fnt(r, c):
        return cell(r, c).font

    def fill_rgb(r, c):
        f = cell(r, c).fill
        if f is None or not f.patternType:
            return None
        return _rgb6(f.fgColor)

    def font_rgb(r, c):
        try:
            return _rgb6(fnt(r, c).color)
        except Exception:
            return None

    def font_sz(r, c):
        sz = fnt(r, c).sz
        return float(sz) if sz else None

    def side_style(r, c, name):
        """单元格某条边的线型（None 安全）。"""
        b = cell(r, c).border
        if b is None:
            return None
        s = getattr(b, name, None)
        return getattr(s, "style", None) if s is not None else None

    def yahei(cs):
        """一批单元格的字体是否统一为微软雅黑。"""
        for ref in cs:
            m = re.match(r"([A-Z]+)(\d+)$", ref)
            if m is None:
                return False
            r = int(m.group(2))
            c = column_index_from_string(m.group(1))
            if fnt(r, c).name != _YAHEI:
                return False
        return True

    # ---------- 两张新增清单的行号锚定 ----------
    def find_row(text):
        for r in range(1, _MAXR + 1):
            if _text(cell(r, 1).value) == text:
                return r
        return None

    def seq_no(v):
        """序号单元格取整数值（文本 "1" 与数值 1 同等对待）。"""
        n = _num(v)
        if n is None or abs(n - round(n)) > 1e-9:
            return None
        return int(round(n))

    def locate(T):
        """表名称行 → 下一行为表头行 → A 列“… 小计”行 → 二者之间序号 1..5 的明细行。"""
        t = find_row(T["title"])
        s = find_row(T["title"] + " 小计")
        if t is None or s is None or s <= t + 1:
            return None
        det = [r for r in range(t + 2, s) if seq_no(cell(r, 1).value) in (1, 2, 3, 4, 5)]
        return {"title": t, "header": t + 1, "detail": det, "sub": s}

    L5 = locate(_T5)
    L6 = locate(_T6)

    # ========== 维度二：加分项（指令要求做的事） ==========

    # 01 +3：截图已被单元格清单取代
    #   ① 工作簿内不存在任何嵌入图片/截图对象（openpyxl 的 ws._images）
    #   ② 第 69–88 行 A:J 区域确有以单元格文本与数值承载的清单内容
    img_cnt = sum(len(getattr(w, "_images", []) or []) for w in wbf.worksheets)
    filled = numeric = 0
    for r in range(69, 89):
        for c in range(1, 11):
            if _text(cell(r, c).value) != "":
                filled += 1
            if _num(vcell(r, c).value) is not None:
                numeric += 1
    add(img_cnt == 0 and filled >= 20 and numeric >= 5, 3,
        "工作表中不存在任何嵌入的图片或截图对象，第69行至第88行区域内存在以单元格文本与数值承载的清单内容")

    def check_table(T, L):
        """按细则顺序追加某张新增清单的 9 个加分项判定。"""
        t = T["title"]

        # 表名称单元格：微软雅黑、加粗、字体色 1F4E3D、填充 DDEFE7、整行合并居中
        ok = False
        if L is not None:
            r = L["title"]
            ok = (fnt(r, 1).name == _YAHEI and bool(fnt(r, 1).b)
                  and font_rgb(r, 1) == "1F4E3D" and fill_rgb(r, 1) == "DDEFE7"
                  and row_full_merged(r)
                  and (cell(r, 1).alignment.horizontal or "") == "center")
        add(ok, 5, "工作表中存在表名称单元格文本为“%s”的表格，该表名称单元格字体格式为微软雅黑、"
                   "加粗、字体颜色1F4E3D，填充色DDEFE7，整行合并居中" % t)

        # 表头行：紧邻表名称行的下一行，10 列文本与字体底色
        ok = False
        if L is not None:
            hr = L["header"]
            ok = all(_text(cell(hr, c).value) == _HEADER[c - 1] for c in range(1, 11))
            for c in range(1, 11):
                if not (fnt(hr, c).name == _YAHEI and font_sz(hr, c) == 10.0
                        and bool(fnt(hr, c).b) and font_rgb(hr, c) == "FFFFFF"
                        and fill_rgb(hr, c) == "2F7D62"):
                    ok = False
        add(ok, 3, "“%s”表名称行的下一行为表头行，其文本依次为“序号、项目名称、项目特征描述、"
                   "工程量、单位、综合单价（元）、合价（元）、施工/供应说明、使用部位、备注”，"
                   "字体格式为微软雅黑、10号、加粗、字体颜色FFFFFF，填充色2F7D62" % t)

        # 明细行：必须恰好 5 行、序号 1..5
        det = L["detail"] if L is not None else []
        ok5 = (L is not None and len(det) == 5
               and [seq_no(cell(r, 1).value) for r in det] == [1, 2, 3, 4, 5])

        add(ok5 and [_text(cell(r, 2).value) for r in det] == T["names"], 5,
            "“%s”表头行与A列文本为“%s 小计”的小计行之间的明细行，其序号依次为1至5，"
            "项目名称依次为“%s”" % (t, t, "、".join(T["names"])))

        add(ok5 and [_text(cell(r, 3).value) for r in det] == T["feats"], 5,
            "“%s”项目特征描述依次为%s" % (t, "".join("“%s”" % x for x in T["feats"])))

        add(ok5 and all(_same_num(vcell(r, 4).value, q) for r, q in zip(det, T["qty"]))
            and [_text(cell(r, 5).value) for r in det] == T["unit"], 3,
            "“%s”工程量与单位依次为%s" % (
                t, "、".join("%g%s" % (q, u) for q, u in zip(T["qty"], T["unit"]))))

        add(ok5 and all(_same_num(vcell(r, 6).value, p) for r, p in zip(det, T["price"])), 3,
            "“%s”综合单价（元）依次为%s" % (t, "、".join("%g" % p for p in T["price"])))

        add(ok5 and all(_same_num(vcell(r, 7).value, a) for r, a in zip(det, T["amount"])), 3,
            "“%s”合价（元）依次为%s" % (t, "、".join("%g" % a for a in T["amount"])))

        add(ok5 and all(_text(cell(r, 8).value) == "设备供应" for r in det)
            and [_text(cell(r, 9).value) for r in det] == T["parts"], 3,
            "“%s”5行明细的施工/供应说明均为“设备供应”，使用部位依次为“%s”"
            % (t, "、".join(T["parts"])))

        # 小计行：A 列文本、G 列金额、整行加粗与 FFF2CC 底色
        #   小计行 A:F 为一个合并区，B:F 属于“非锚点”格子，办公软件里不显示其自身
        #   格式，故只校验锚点格与未合并格。
        ok = False
        if L is not None:
            sr = L["sub"]
            ok = (_text(cell(sr, 1).value) == t + " 小计"
                  and _same_num(vcell(sr, 7).value, T["subtotal"]))
            for c in range(1, 11):
                if is_shadow(sr, c):
                    continue
                if not (bool(fnt(sr, c).b) and fill_rgb(sr, c) == "FFF2CC"):
                    ok = False
        add(ok, 3, "“%s”小计行A列文本为“%s 小计”，G列结果为%d，该行加粗且填充色为FFF2CC"
                   % (t, t, T["subtotal"]))

    check_table(_T5, L5)   # 02–10
    check_table(_T6, L6)   # 11–19

    # 20 +5：两个小计落在 G76/G86，且 F122 总计确实引用到这两格、结果非 0 非错误值
    f122_val = vcell(122, 6).value
    f122_num = _num(f122_val)
    total_ok = (f122_num is not None and abs(f122_num) > 1e-9
                and not (isinstance(f122_val, str) and f122_val.startswith("#")))
    add(_same_num(vcell(76, 7).value, 448580) and _same_num(vcell(86, 7).value, 90440)
        and _refs_cover(cell(122, 6).value, ["G76", "G86"]) and total_ok, 5,
        "G76单元格结果为448580、G86单元格结果为90440，F122总计单元格能取到G76与G86这两个"
        "小计金额，其结果不为0且不返回错误值")

    # 21 +1：两张新表（表名称行/表头行/5 行明细/小计行）A:J 的细实线边框
    #   办公软件无法在合并区内部画竖线，故落在合并区内部的左右分隔位置不作要求；
    #   上下边框一律要求细实线。两张表若没做出来，本项直接不得分。
    nb = 0
    if L5 is None or L6 is None:
        nb = 1
    else:
        for L in (L5, L6):
            for r in [L["title"], L["header"]] + L["detail"] + [L["sub"]]:
                for c in range(1, 11):
                    need = ["top", "bottom"]
                    if m_of.get((r, c)) is None or m_of.get((r, c)) != m_of.get((r, c - 1)):
                        need.append("left")
                    if m_of.get((r, c)) is None or m_of.get((r, c)) != m_of.get((r, c + 1)):
                        need.append("right")
                    for sd in need:
                        if side_style(r, c, sd) != "thin":
                            nb += 1
    add(nb == 0, 1,
        "“五、环控与照明物资清单”与“六、物联网设备预算清单”两个表格的表名称行、表头行、"
        "表头行与小计行之间的5行明细行、小计行，其A列至J列单元格的上、下边框均为细实线，"
        "左、右边框中除落在合并单元格内部的分隔位置外均为细实线")

    # 22–24 +3/+3/+1：三组“表内字体统一为微软雅黑”的区域
    add(yahei(_cells("C27:C29,C42:C44,C50:C52,C91:C95")), 3,
        "C27:C29、C42:C44、C50:C52、C91:C95的字体统一为微软雅黑")
    add(yahei(_cells("D62:D64,E101:E105,F111:F115,G111:G115,F121:F124,G121:G124,H121:H124")), 3,
        "D62:D64、E101:E105、F111:F115、G111:G115、F121:F124、G121:G124、H121:H124的"
        "字体统一为微软雅黑")
    add(yahei(_cells("I34:I37,G62:G65")), 1, "I34:I37、G62:G65的字体统一为微软雅黑")

    # 25 +3：两个新表名称字号与其余 7 个表名称一致，同为 11 号
    sz_ok = True
    for L in (L5, L6):
        if L is None or font_sz(L["title"], 1) != 11.0:
            sz_ok = False
    for r in (5, 24, 59, 89, 99, 109, 119):
        if font_sz(r, 1) != 11.0:
            sz_ok = False
    add(sz_ok, 3,
        "文本为“五、环控与照明物资清单”的表名称单元格、文本为“六、物联网设备预算清单”的"
        "表名称单元格，二者字号均为11号，与A5“一、项目报价汇总”、A24“三、温室工程量清单”、"
        "A59“四、温室主材预算清单”、A89“七、分拣包装与冷链设备清单”、A99“八、运维家具与"
        "工具清单”、A109“九、不锈钢玻璃与五金辅材清单”、A119“十、清单总计与核对”的"
        "11号字号相同")

    def need_height(r, c):
        """该单元格文字完整显示所需的最小行高（磅）。
        以列宽字符数为宽度基准（合并单元格取所跨各列列宽之和），汉字按 2 个字符
        计算内容长度；设置了自动换行且内容长度超出宽度 2 个字符以上时，按
        ceil(内容长度/宽度) 得到所需行数；行高下限 = 行数 × 字号 × 1.3。"""
        c_obj = cell(r, c)
        w = span_w.get((r, c), col_w(c))
        sz = font_sz(r, c) or 11.0
        lines = 1
        if c_obj.alignment.wrap_text and w > 0:
            lines = 0
            for seg in str(c_obj.value).split("\n"):
                u = _units(seg)
                lines += max(1, math.ceil(u / w)) if u > w + 2 else 1
            lines = max(1, lines)
        return lines * sz * 1.3

    # 26 +3：第 22 行说明单元格行高≥90 磅且多行说明文字完整显示
    h22 = row_h(22)
    add(_text(cell(22, 1).value) != "" and h22 >= 90 - 0.1
        and h22 >= need_height(22, 1) - 0.1, 3,
        "A22“说明”单元格所在第22行行高不小于90磅，单元格内多行说明文字全部完整显示、无截断")

    # 27 +3：全表凡含内容的单元格，行高均不小于其所需显示高度（文字不被压扁截断）
    #   排除第 1、22、31、38、46、54、96、116 行——这些行的行高由各自专门的
    #   条目单独判定，此处不重复计入。合并区内的非锚点格子不承载显示内容。
    bad_fit = 0
    for r in range(1, _MAXR + 1):
        if r in (1, 22, 31, 38, 46, 54, 96, 116):
            continue
        hh = row_h(r)
        for c in range(1, 11):
            if _text(cell(r, c).value) == "" or is_shadow(r, c):
                continue
            if hh < need_height(r, c) - 0.1:
                bad_fit += 1
    add(bad_fit == 0, 3,
        "工作表中除第1、22、31、38、46、54、96、116行外，凡含有内容的单元格，其所在行行高均"
        "不小于该单元格字号磅值的1.3倍；其中设置了自动换行的单元格，以列宽字符数为宽度基准、"
        "汉字按2个字符计算内容长度，合并单元格的宽度取所跨各列列宽之和，当内容长度超出宽度"
        "2个字符以上时，按内容长度除以宽度向上取整得到所需显示行数，其所在行行高不小于该行数"
        "与字号磅值1.3倍的乘积，全表文字均完整显示、无压扁或截断")

    # 28 +1：相似表格的明细行行高一致（互差≤1.5 磅）
    unif_rows = []
    for spec in ("27:30", "34:37", "41:45", "49:53", "61:65",
                 "91:95", "101:105", "111:115"):
        a, b = spec.split(":")
        unif_rows += list(range(int(a), int(b) + 1))
    unif_ok = L5 is not None and L6 is not None
    if unif_ok:
        unif_rows += L5["detail"] + L6["detail"]
        hs = [row_h(r) for r in unif_rows]
        unif_ok = (max(hs) - min(hs)) <= 1.5 + 1e-9
    add(unif_ok, 1,
        "第27:30、34:37、41:45、49:53、61:65、91:95、101:105、111:115行，以及“五、环控与"
        "照明物资清单”“六、物联网设备预算清单”两个表格表头行与小计行之间的5行明细行，"
        "行高彼此相差不超过1.5磅")

    # ---------- 打印分页模拟（29、30 共用） ----------
    #   列宽字符数 -> 像素 = 7×字符数+5，像素 -> 磅 = ×0.75；
    #   可用页宽/页高按 A4 减去页边距换算；缩放比例取“缩放比例”设置，
    #   未设缩放而启用了“调整为1页宽”时按自适应缩放计算，两者都没设则按 100%。
    ps = ws.page_setup
    pm = ws.page_margins
    w_pt = sum(7 * col_w(c) + 5 for c in range(1, 11)) * 0.75
    avail_w = (_A4_W_IN - float(pm.left or 0) - float(pm.right or 0)) * 72
    avail_h = (_A4_H_IN - float(pm.top or 0) - float(pm.bottom or 0)) * 72
    sp = ws.sheet_properties.pageSetUpPr
    fit_to_page = bool(sp is not None and sp.fitToPage) and bool(ps.fitToWidth)
    if ps.scale:
        zoom = float(ps.scale) / 100.0
    elif fit_to_page:
        zoom = min(1.0, avail_w / w_pt) if w_pt > 0 else 1.0
    else:
        zoom = 1.0
    page_limit = avail_h / zoom if zoom > 0 else avail_h

    # 29 +5：A:J 十列在打印预览中位于同一页宽内
    add((fit_to_page or bool(ps.scale)) and w_pt <= avail_w / zoom + 0.5, 5,
        "打印设置中启用“调整为1页宽”或设置了缩放比例，使A列至J列共10列在打印预览中位于同一"
        "页宽内，没有列方向的断开")

    # 30 +3：每张清单表格完整落在同一页内
    breaks = set()
    try:
        breaks = set(int(b.id) for b in ws.row_breaks.brk)
    except Exception:
        breaks = set()
    page_of, page, used = {}, 1, 0.0
    for r in range(1, _MAXR + 1):
        rh = row_h(r)
        if used + rh > page_limit or (r - 1) in breaks:
            page += 1
            used = 0.0
        page_of[r] = page
        used += rh
    tables = [(5, 21), (24, 56), (59, 66), (89, 96), (99, 106), (109, 116), (119, 124)]
    if L5 is not None:
        tables.append((L5["title"], L5["sub"]))
    if L6 is not None:
        tables.append((L6["title"], L6["sub"]))
    split_ok = (L5 is not None and L6 is not None
                and all(page_of[a] == page_of[b] for a, b in tables))
    add(split_ok, 3,
        "打印预览中每个清单表格的表名称行、表头行、明细行、小计行完整落在同一页内，"
        "没有任何一个表格被拆分到两页")

    # ========== 维度二：减分项（指令没要求、属于把原有排版改坏） ==========
    # 约定：hit 为 True 表示“细则描述的不满足情形成立”，即扣分命中。

    def ref_rc(ref):
        m = re.match(r"([A-Z]+)(\d+)$", ref)
        return int(m.group(2)), column_index_from_string(m.group(1))

    # 31 -3：J 列“备注”区域应保持全部空白
    j_spec = ("J27:J30,J34:J37,J41:J45,J49:J53,J62:J65,J91:J95,"
              "J101:J105,J111:J115,J121:J123")
    j_bad = any(_text(cell(*ref_rc(x)).value) != "" for x in _cells(j_spec))
    add(j_bad, -3,
        "J27:J30、J34:J37、J41:J45、J49:J53、J62:J65、J91:J95、J101:J105、J111:J115、"
        "J121:J123“备注”列单元格不满足全部为空白")

    # 32 -3：F2 报价日期
    f2 = cell(2, 6).value
    f2_ok = (isinstance(f2, (datetime.datetime, datetime.date))
             and f2.year == 2026 and f2.month == 5 and f2.day == 25
             and _norm_fmt(cell(2, 6).number_format) == "yyyy-mm-dd")
    add(not f2_ok, -3, "F2单元格不满足内容为日期2026-05-25、数字格式为yyyy-mm-dd")

    # 33 -3：H 列指定区域水平右对齐
    h_spec = ("H6:H21,H25,H27:H31,H34:H38,H41:H46,H49:H54,H60:H66,H90:H96,"
              "H100:H106,H110:H116,H120:H124")
    h_bad = False
    for x in _cells(h_spec):
        r, c = ref_rc(x)
        if is_shadow(r, c):
            continue
        if (cell(r, c).alignment.horizontal or "") != "right":
            h_bad = True
    add(h_bad, -3,
        "H6:H21、H25、H27:H31、H34:H38、H41:H46、H49:H54、H60:H66、H90:H96、H100:H106、"
        "H110:H116、H120:H124不满足水平对齐方式为右对齐")

    # 34 -3：第 56 行“温室工程量清单合计”行
    r56_bad = False
    for c in (1, 7, 8, 9, 10):
        if is_shadow(56, c):
            continue
        if not (fnt(56, c).name == _YAHEI and bool(fnt(56, c).b)
                and font_rgb(56, c) == "FFFFFF" and fill_rgb(56, c) == "1F4E3D"):
            r56_bad = True
    add(r56_bad, -3,
        "A56、G56、H56、I56、J56“温室工程量清单合计”行不满足字体格式为微软雅黑、加粗、"
        "字体颜色FFFFFF，填充色1F4E3D")

    # 35 -3：四个子组标题行
    sub_bad = False
    for r in (26, 33, 40, 48):
        if not (fnt(r, 1).name == _YAHEI and font_sz(r, 1) == 10.0
                and bool(fnt(r, 1).b) and font_rgb(r, 1) == "1F4E3D"
                and fill_rgb(r, 1) == "E8F5E9"):
            sub_bad = True
    add(sub_bad, -3,
        "A26“1. 温室主体与基础”、A33“2. 地面、排水与道路”、A40“3. 环控、水肥与电气”、"
        "A48“4. 物联网与分拣配套”不满足字体格式为微软雅黑、10号、加粗、字体颜色1F4E3D，"
        "填充色E8F5E9")

    # 36 -3：“五”表表头行与小计行之间夹了整行无内容的空白行
    #   表格本身没做出来时不构成“改坏”，不扣分。
    blank_in = False
    if L5 is not None:
        rows_between = list(range(L5["header"] + 1, L5["sub"]))
        if len(rows_between) != 5:
            blank_in = True
        for r in rows_between:
            if all(_text(cell(r, c).value) == "" for c in range(1, 11)):
                blank_in = True
    add(blank_in, -3,
        "“五、环控与照明物资清单”表格的表头行与小计行之间不满足为连续的5行明细，"
        "中间夹有整行无内容的空白行")

    # 37 -3：各表小计行行高 48±1.5 磅
    h48_bad = any(abs(row_h(r) - 48.0) > 1.5 for r in (31, 38, 46, 54, 96, 116))
    add(h48_bad, -3, "第31、38、46、54、96、116行行高不满足48±1.5磅")

    # 38 -3：12 个空白间隔行保持“无内容、无填充色、常规格式、微软雅黑”
    gap_rows = (57, 58, 67, 68, 87, 88, 97, 98, 107, 108, 117, 118)
    gap_bad = False
    for r in gap_rows:
        for c in range(1, 11):
            if _text(cell(r, c).value) != "" or fill_rgb(r, c) is not None \
                    or _norm_fmt(cell(r, c).number_format) not in ("general", "") \
                    or fnt(r, c).name != _YAHEI:
                gap_bad = True
    add(gap_bad, -3,
        "第57、58、67、68、87、88、97、98、107、108、117、118行的A列至J列单元格不满足为"
        "无任何内容、无填充色、数字格式为常规、字体为微软雅黑的空白间隔行")

    # 39 -1：D 列工程量区域数字格式 #,##0.00
    d_spec = ("D27:D30,D34:D37,D41:D45,D49:D53,D61:D65,D91:D95,D101:D105,D111:D115")
    d_bad = any(_norm_fmt(cell(*ref_rc(x)).number_format) != "#,##0.00"
                for x in _cells(d_spec))
    add(d_bad, -1,
        "D27:D30、D34:D37、D41:D45、D49:D53、D61:D65、D91:D95、D101:D105、D111:D115"
        "不满足数字格式为#,##0.00")

    # 40 -1：H7:H20 占比列数字格式 0.0%
    pct_bad = any(_norm_fmt(cell(r, 8).number_format) != "0.0%" for r in range(7, 21))
    add(pct_bad, -1, "H7:H20“占比”列单元格不满足数字格式为0.0%")

    # 41 -1：A2:J3 填充色 F4FAF7
    head_fill_bad = False
    for r in (2, 3):
        for c in range(1, 11):
            if is_shadow(r, c):
                continue
            if fill_rgb(r, c) != "F4FAF7":
                head_fill_bad = True
    add(head_fill_bad, -1, "A2:J3区域单元格不满足填充色为F4FAF7")

    # 42 -1：表头标签字体颜色 1F4E3D
    label_bad = any(font_rgb(*ref_rc(x)) != "1F4E3D"
                    for x in ("A2", "C2", "E2", "G2", "I2", "A3", "C3", "E3", "G3", "I3"))
    add(label_bad, -1,
        "A2“报价单位”、C2“项目地点”、E2“报价日期”、G2“币种”、I2“有效期”、A3“项目规模”、"
        "C3“建设周期”、E3“报价口径”、G3“版本”、I3“联系人”不满足字体颜色为1F4E3D")

    # 43 -1：A22 说明单元格填充色 EAF3F8
    add(fill_rgb(22, 1) != "EAF3F8", -1, "A22“说明”单元格不满足填充色为EAF3F8")

    # 44 -1：A1 主标题格式与 A1:J1 合并居中、不自动换行
    a1 = cell(1, 1)
    a1_ok = (fnt(1, 1).name == _YAHEI and font_sz(1, 1) == 18.0 and bool(fnt(1, 1).b)
             and font_rgb(1, 1) == "FFFFFF" and fill_rgb(1, 1) == "1F4E3D"
             and row_full_merged(1) and (a1.alignment.horizontal or "") == "center"
             and not a1.alignment.wrap_text)
    add(not a1_ok, -1,
        "A1“云谷智慧农业示范园建设项目报价表”不满足字体格式为微软雅黑、18号、加粗、"
        "字体颜色FFFFFF，填充色1F4E3D，A1:J1合并居中且不自动换行")

    # 45–47 -1：C 列列宽、H 列列宽、第 1 行行高
    add(abs(col_w(3) - 28.0) > 1.5, -1, "C列列宽不满足28±1.5字符")
    add(abs(col_w(8) - 12.0) > 1.5, -1, "H列列宽不满足12±1.5字符")
    add(abs(row_h(1) - 36.0) > 1.5, -1, "第1行行高不满足36±1.5磅")

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
    result["total_score"] = max(0, total_score)   # 扣分多于加分时兜底为 0
    result["max_score"] = MAX_SCORE_TOTAL
    return result


if __name__ == "__main__":
    # 本地调试入口：默认评估脚本所在目录，也可传入待评估目录；
    # 追加 --brief 只打印命中情况概览，便于人工核对。
    _args = [a for a in sys.argv[1:] if not a.startswith("--")]
    _dir = _args[0] if _args else os.path.dirname(os.path.abspath(__file__))
    _out = evaluate(_dir)
    if "--brief" in sys.argv:
        # dim1_pass 是三态：True 通过、False 不通过、空串表示这份作业没评上
        _d1 = {True: "通过", False: "不通过"}.get(_out["dim1_pass"], "未判定")
        print("%s  维度一：%s %s" % (_out["file_name"], _d1,
                                  _out["dim1_reason"] or _out["error"] or ""))
        for i, it in enumerate(_out["dim2_items"], 1):
            print("%02d  %+3d/%+3d  %s" % (i, it["delta"], it["max_delta"], it["rule"][:60]))
        print("总分：%d / %d" % (_out["total_score"], _out["max_score"]))
    else:
        print(json.dumps(_out, ensure_ascii=False, indent=2))

