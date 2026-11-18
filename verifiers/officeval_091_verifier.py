#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# pyright: reportUnknownMemberType=false, reportUnknownVariableType=false
# pyright: reportUnknownArgumentType=false, reportMissingTypeStubs=false
# pyright: reportOptionalMemberAccess=false, reportGeneralTypeIssues=false
"""
091 号任务评分器：《2026年区域渠道销量统计表.xlsx》"12月"表的单据录入 + 按客户×大类统计。

说明：按接口统一约定，本模块不修改全局 sys.stdout，也不 print 主结果。
主结果统一由 evaluate(dir_path) -> dict 返回；入参为"待评测交付文件所在目录的路径"。
__main__ 分支仅用于本地调试，直接打印 JSON。

评分口径（对齐同目录 细则.txt）：
  维度1 可用与可修改性：交付 1 个可打开的 xlsx/xls，"12月"表仍在且为可编辑单元格与公式。
  维度2 完成度：11 条加分项（满分 23）+ 29 条减分项（命中即扣，可累计）。
  max_score = 加分项之和；total_score = 命中加分项 - 命中减分项（可为负，下限截到 0）。
"""

import os
import re
import sys
import json

import openpyxl
from openpyxl.utils import get_column_letter, column_index_from_string

SCRIPT_ID = "091"
MAX_SCORE_TOTAL = 23

_MAX_DELTAS_ORDER = [
    # ===== 一、加分项（共 23 分）=====
    1,   # 01 重量 1600KG 换算为 1.6 吨写入
    3,   # 02 "育成期配合饲料"归入"保育/育肥料"列
    1,   # 03 并入第 41 行已有"连塘合作养殖社"，未新增重复行
    1,   # 04 数量(吨)2.52 / 组小计 385.48 / 全表合计 741.56 三处自洽
    5,   # 05 五组左右两栏合并为统一清单，覆盖 57 个客户无遗漏无重复
    5,   # 06 57 客户各大类数值逐格对位、未错列
    3,   # 07 六个大类总量正确
    1,   # 08 按六个大类分列且列名一致齐全
    1,   # 09 每客户"数量(吨)"行合计 + 清单总计 741.56 自洽
    1,   # 10 统计结果置于新表或数据区之外，未覆盖已有行
    1,   # 11 清单表头完整可读、数值两位小数
    # ===== 二、减分项 =====
    -1,  # 12 B1:R1 合并标题文字与格式
    -1,  # 13 第 4/18/39 行表头文字、顺序、位置
    -1,  # 14 表头行浅蓝底纹 DDEBF7、加粗、居中
    -1,  # 15 D3、M3"乳猪料"标注
    -1,  # 16 A3/A17/A38/J3/J17 组序号与蓝色底纹 5B9BD5
    -1,  # 17 五个渠道组组名
    -1,  # 18 P58"制表：渠道运营部"
    -3,  # 19 第 57 行 B~I"小计"行与七个求和公式
    -1,  # 20 第 57 行 B~I 加粗居中、E2F0D9 底纹、0.00 格式
    -1,  # 21 第 12/25/54 行小计求和公式被写死或改错区间
    -1,  # 22 右下角合计 =C12+L12+C25+L54+C57 被写死
    -1,  # 23 "小计""合计"文字被改写或整行删除、移位
    -3,  # 24 57 个客户名称、所属渠道组、所在行次
    -3,  # 25 单据涉及行以外客户的大类数值
    -3,  # 26 客户行"数量(吨)"求和公式被写死
    -1,  # 27 原本空白的大类单元格被填入占位符
    -1,  # 28 车牌/电话/编号/人名/现结等无关信息被写入
    -1,  # 29 单据"数量 64""单价 86""金额 5504"被写入
    -1,  # 30 数值区域两位小数格式被改
    -1,  # 31 字体出现明显异体字或字号不一
    -1,  # 32 有内容单元格的居中、自动换行被取消
    -1,  # 33 列宽行高被改动或行列被隐藏
    -1,  # 34 数据区插入或删除行列致对应关系错位
    -1,  # 35 B1:R1 之外新增合并或取消已有合并
    -1,  # 36 新增无关辅助列、批注、条件格式或高亮标色
    -3,  # 37 为放统计结果删除、清空、覆盖"12月"已有行
    -1,  # 38 "12月"工作表被删除或重命名
    -1,  # 39 新增空白/临时工作表未清理
    -1,  # 40 交付为 .csv/.txt 等丢失公式与格式的类型
]
assert sum(d for d in _MAX_DELTAS_ORDER if d > 0) == MAX_SCORE_TOTAL

# ===================== "12月"表既有版式基线（不得被改动的部分） =====================

SHEET_NAME = "12月"
TITLE_CELL = "B1"
TITLE_TEXT = "2026年12月份区域渠道销量统计表"
TITLE_MERGE = "B1:R1"

_LEFT_CATS = "DEFGHI"
_RIGHT_CATS = "MNOPQR"
_HEADER_LABELS = ["客户", "数量(吨)", "开口料", "保育/育肥料", "蛋禽料", "水禽料", "沉水鱼料", "浮水鱼料"]
_CATEGORIES = _HEADER_LABELS[2:]

# (序号单元格, 组名单元格, 组名, 客户列, 数量列, 六个大类列, 客户起行, 客户止行, 小计行)
_GROUPS = [
    ("A3", "B3", "东湾渠道一组", "B", "C", _LEFT_CATS, 5, 10, 12),
    ("A17", "B17", "海岚客户维护组", "B", "C", _LEFT_CATS, 19, 24, 25),
    ("A38", "B38", "云澜养殖拓展组", "B", "C", _LEFT_CATS, 40, 46, 57),
    ("J3", "K3", "北岭养殖服务组", "K", "L", _RIGHT_CATS, 5, 7, 12),
    ("J17", "K17", "西禾区域合作组", "K", "L", _RIGHT_CATS, 19, 53, 54),
]
_GROUP_SERIALS = {"A3": 1, "A17": 2, "A38": 3, "J3": 4, "J17": 5}
_PIGLET_CELLS = ["D3", "M3"]
_MAKER_CELL = "P58"
_MAKER_TEXT = "制表：渠道运营部"

# 表头块：(行号, 起始列字母) —— 自起始列起连续 8 列对应 _HEADER_LABELS
_HEADER_BLOCKS = [(4, "B"), (4, "K"), (18, "B"), (18, "K"), (39, "B")]

# 小计块：(小计行, 标签列, 数值列字母串, 求和起行, 求和止行)
_SUBTOTALS = [
    (12, "B", "C" + _LEFT_CATS, 5, 10),
    (12, "K", "L" + _RIGHT_CATS, 5, 7),
    (25, "B", "C" + _LEFT_CATS, 19, 24),
    (54, "K", "L" + _RIGHT_CATS, 19, 53),
    (57, "B", "C" + _LEFT_CATS, 40, 46),
]
_SUBTOTAL_LABEL_CELLS = ["B12", "K12", "B25", "K54", "B57"]
_SUBTOTAL_TEXT = "小计"
_TOTAL_LABEL_CELL = "K57"
_TOTAL_TEXT = "合计"
_TOTAL_FORMULA_CELL = "L57"
_TOTAL_FORMULA_REFS = ["C12", "L12", "C25", "L54", "C57"]

_FILL_TITLE = "1F4E78"
_FILL_HEADER = "DDEBF7"
_FILL_SERIAL = "5B9BD5"
_FILL_SUBTOTAL = "E2F0D9"
_FILL_TOTAL = "FCE4D6"

# 本次单据应落位的目标：第 5 组第 41 行"连塘合作养殖社"，保育/育肥料 0.92 -> 2.52
_TARGET_CUSTOMER = "连塘合作养殖社"
_TARGET_ROW = 41
_TARGET_QTY_CELL = "L41"
_TARGET_CAT_CELL = "N41"
_TARGET_BASE = 0.92
_TARGET_AFTER = 2.52
_TARGET_OTHER_CAT_CELLS = ["M41", "O41", "P41", "Q41", "R41"]
_GROUP5_SUBTOTAL_CELL = "L54"
_GROUP5_SUBTOTAL_AFTER = 385.48
_GRAND_TOTAL_AFTER = 741.56
_CATEGORY_TOTALS = [0.26, 635.98, 99.06, 1.04, 4.58, 0.64]

# 57 个客户及其在"12月"表中的既有坐标
_CUSTOMER_CELLS = {
    "B5": "暮野生态养殖场", "B6": "苍樾农牧合作社", "B7": "泠泽家庭农场",
    "B8": "芪牧饲养基地", "B9": "槿谷养殖点", "B10": "砚滨牧业服务站",
    "B19": "星溯湾养殖场", "B20": "镜渊农牧社", "B21": "岫岩育养基地",
    "B22": "渡川家庭农场", "B23": "暮岑合作养殖点", "B24": "涧埠饲养服务部",
    "B40": "泠溪养殖合作社", "B41": "攸禾家庭农场", "B42": "景隅农牧场",
    "B43": "栖林生态养殖场", "B44": "檀峰饲养中心", "B45": "青岑合作养殖社",
    "B46": "景言农牧基地",
    "K5": "芜畴联合养殖场", "K6": "沧隅牧业合作社", "K7": "汀沿饲养中心",
    "K19": "祁朔农牧场", "K20": "栀隅合作社", "K21": "阑渚养殖基地",
    "K22": "沧浔家庭农场", "K23": "苇汀育养中心", "K24": "泠漾农牧服务站",
    "K25": "茴湾饲养点", "K26": "峤南农牧场", "K27": "潇涧合作养殖场",
    "K28": "筠筱养殖社", "K29": "荞墟农牧基地", "K30": "芃洲饲养中心",
    "K31": "西墟联营农场", "K32": "檀汐家庭牧场", "K33": "漠汀生态养殖场",
    "K34": "沧漾农牧社", "K35": "岚川育养基地", "K36": "岑峦合作社",
    "K37": "桧杉养殖点", "K38": "缥涧农牧场", "K39": "西渚饲养中心",
    "K40": "玄径生态农场", "K41": "连塘合作养殖社", "K42": "汐野农牧基地",
    "K43": "琉湾牧业服务部", "K44": "湄洲家庭农场", "K45": "玦岩联合养殖场",
    "K46": "砚辞生态牧场", "K47": "峥峦合作社", "K48": "泫川农牧场",
    "K49": "曜津饲养点", "K50": "螭渊养殖基地", "K51": "安浔家庭农场",
    "K52": "聿辰农牧社", "K53": "砚漾合作养殖场",
}

# "12月"表既有的大类数值（单据涉及的 N41 亦列出，校验时单独排除）
_BASE_CATEGORY_VALUES = {
    "E5": 78.42, "F5": 5.16, "H5": 1.24, "E6": 21.36, "E7": 8.12, "F8": 14.84,
    "E9": 6.48, "E10": 9.72, "F10": 1.56,
    "D19": 0.08, "E19": 12.4, "F19": 2.36, "G19": 0.92, "H19": 1.16, "E20": 4.8,
    "E21": 43.68, "D22": 0.06, "E22": 19.84, "E23": 7.2, "F23": 0.72, "E24": 18.4,
    "E40": 9.44, "E41": 1.28, "E42": 7.92, "E43": 31.2, "E44": 10.84, "F45": 1.16,
    "E46": 5.68, "F46": 0.64,
    "N5": 24.6, "N6": 3.12, "N7": 1.68,
    "N19": 4.36, "N20": 2.88, "N21": 2.64, "N22": 3.52, "N23": 0.36, "Q23": 0.18,
    "Q24": 1.08, "P25": 0.12, "N26": 42.56, "O26": 0.08, "N27": 3.04, "R27": 0.64,
    "N28": 1.2, "N29": 1.32, "O29": 1.04, "O30": 1.72, "M31": 0.08, "N31": 22.84,
    "N32": 0.72, "N33": 0.36, "N34": 0.08, "M35": 0.04, "N35": 55.76, "O36": 62.4,
    "N37": 0.28, "O37": 0.2, "N38": 4.88, "N39": 1.76, "O39": 3.52, "Q39": 0.92,
    "N40": 3.4, "N41": 0.92, "N42": 13.88, "N43": 21.48, "N44": 15.6, "N45": 73.2,
    "N46": 2.08, "N47": 14.96, "O47": 0.12, "N48": 8.76, "N49": 2.24, "O50": 3.36,
    "N51": 0.68, "N52": 1.88, "O52": 0.18, "N53": 0.56,
}

def _customer_cell(ccol, row):
    return "%s%d" % (ccol, row)


def _build_expected_matrix():
    """由既有坐标基线推出"客户 -> 六个大类"的应有取值（含本次单据并入后的 2.52）。"""
    matrix = {}
    for _serial, _gname_cell, _gname, ccol, _qcol, cats, r1, r2, _sub in _GROUPS:
        for row in range(r1, r2 + 1):
            name = _CUSTOMER_CELLS[_customer_cell(ccol, row)]
            matrix[name] = [_BASE_CATEGORY_VALUES.get("%s%d" % (c, row)) for c in cats]
    matrix[_TARGET_CUSTOMER][1] = _TARGET_AFTER
    return matrix


_EXPECTED_MATRIX = _build_expected_matrix()
assert len(_CUSTOMER_CELLS) == 57 and len(_EXPECTED_MATRIX) == 57
assert [round(sum(v[i] or 0.0 for v in _EXPECTED_MATRIX.values()), 2) for i in range(6)] == _CATEGORY_TOTALS
assert round(sum(sum(x or 0.0 for x in v) for v in _EXPECTED_MATRIX.values()), 2) == _GRAND_TOTAL_AFTER

# 既有版式中水平居中的有内容单元格（标题、组号组名、表头、客户名、小计与合计）
_CENTER_CELLS = set((
    "B1 A3 B3 D3 J3 K3 M3 B4 C4 D4 E4 F4 G4 H4 I4 K4 L4 M4 N4 O4 P4 Q4 R4 "
    "B5 K5 B6 K6 B7 K7 B8 B9 B10 B12 C12 D12 E12 F12 G12 H12 I12 K12 L12 M12 "
    "N12 O12 P12 Q12 R12 A17 B17 J17 K17 B18 C18 D18 E18 F18 G18 H18 I18 K18 "
    "L18 M18 N18 O18 P18 Q18 R18 B19 K19 B20 K20 B21 K21 B22 K22 B23 K23 B24 "
    "K24 B25 C25 D25 E25 F25 G25 H25 I25 K25 K26 K27 K28 K29 K30 K31 K32 K33 "
    "K34 K35 K36 K37 A38 B38 K38 B39 C39 D39 E39 F39 G39 H39 I39 K39 K40 K41 "
    "K42 K43 K44 K45 K46 K47 K48 K49 K50 K51 K52 K53 K54 L54 M54 N54 O54 P54 "
    "Q54 R54 B57 C57 D57 E57 F57 G57 H57 I57 K57 L57"
).split())

# 既有版式中未设水平居中、也未设自动换行的单元格（基线本身如此，不作要求）
_ALIGN_EXEMPT = {"B40"}

# 既有列宽与行高
_COL_WIDTHS = {"A": 6.0, "B": 22.0, "C": 11.0, "D": 12.0, "J": 6.0, "K": 22.0, "L": 11.0, "M": 12.0}
_ROW_HEIGHTS = dict((r, 28.0 if r in (4, 18, 39) else (30.0 if r == 1 else 20.0)) for r in range(1, 59))

# 既有字体：中文宋体、数字 Carlito 视为统一；下列之外的字体名视为明显异体字
_FONT_OK = {None, "", "Carlito", "Calibri", "Calibri Light", "宋体", "SimSun", "NSimSun",
            "Arial", "Times New Roman", "Tahoma", "Verdana"}
_FONT_SIZES_OK = {11.0, 12.0, 15.0}

# 单据中与销量统计无关、不应写入表内的信息
_NOISE_TOKENS = ["鲁K6A12Y", "鲁k6a12y", "13962514806", "DH202604180012", "dh202604180012",
                 "沈聿白", "江叙言", "阮清禾", "夏知莞", "宋舒窈",
                 "现结", "未付款", "开单人", "提货人", "仓管", "车牌", "联系电话", "经办"]
# 单据中不应混入"吨"口径的数字字段
_WRONG_NUMBERS = [64.0, 86.0, 5504.0]
_RAW_WEIGHT = 1600.0

# ===================== 通用辅助 =====================

_REF_RE = re.compile(r"^([A-Z]{1,3})(\d+)$")
_SUM_RE = re.compile(r"^SUM\(([A-Z]{1,3})(\d+):([A-Z]{1,3})(\d+)\)$")
_REFS_RE = re.compile(r"[A-Z]{1,3}\d+")
_TERM_RE = re.compile(r"[+\-]?[^+\-]+")


def _num(value):
    """取数值：数字直接返回，纯数字文本（允许千分位、末尾"吨"）宽容解析，其它返回 None。"""
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    if isinstance(value, str):
        s = value.strip().replace(",", "").replace("，", "")
        for tail in ("吨", "t", "T", "KG", "kg"):
            if s.endswith(tail):
                s = s[: -len(tail)].strip()
                break
        try:
            return float(s)
        except ValueError:
            return None
    return None


def _txt(value):
    if value is None:
        return ""
    s = str(value).strip()
    for ch in (" ", "　", "\n", "\r", "\t"):
        s = s.replace(ch, "")
    return s


def _key(value):
    """大类/表头名归一：去掉分隔符与括号差异后比较。"""
    s = _txt(value)
    for ch in ("/", "、", "\\", "(", ")", "（", "）", "："):
        s = s.replace(ch, "")
    return s


def _is_formula(value):
    return isinstance(value, str) and value.startswith("=")


def _close(a, b, tol=0.005):
    return a is not None and b is not None and abs(a - b) <= tol


def _blank(value):
    """空白判定：None、空串，或 0 值占位一律不算"有内容的正数"。"""
    return value is None or (isinstance(value, str) and _txt(value) == "")


def _norm_formula(value):
    if not _is_formula(value):
        return ""
    return value[1:].replace(" ", "").replace("$", "").upper()

# ===================== 单元格取值（缓存值优先 + 最小公式求值） =====================


def _raw(bk, sheet, coord):
    ws = bk["wb"][sheet] if sheet in bk["wb"].sheetnames else None
    if ws is None:
        return None
    try:
        return ws[coord].value
    except Exception:
        return None


def _cached(bk, sheet, coord):
    wbc = bk["wbc"]
    if wbc is None or sheet not in wbc.sheetnames:
        return None
    try:
        return wbc[sheet][coord].value
    except Exception:
        return None


def _value(bk, sheet, coord, depth=0):
    """求单元格数值：常量直接取；公式优先取 Excel 缓存结果，无缓存时对 =SUM(区域)/=A+B 求值。"""
    raw = _raw(bk, sheet, coord)
    if _is_formula(raw):
        cv = _num(_cached(bk, sheet, coord))
        if cv is not None:
            return cv
        if depth > 12:
            return None
        return _eval_formula(bk, sheet, raw, depth + 1)
    n = _num(raw)
    if n is not None:
        return n
    if raw is None:
        return _num(_cached(bk, sheet, coord))
    return None


def _eval_formula(bk, sheet, formula, depth):
    s = _norm_formula(formula)
    if not s or "!" in s or "," in s or "*" in s or "/" in s:
        return None
    total = 0.0
    for term in _TERM_RE.findall(s):
        sign = -1.0 if term.startswith("-") else 1.0
        term = term.lstrip("+-")
        if not term:
            continue
        m = _SUM_RE.match(term)
        if m:
            c1 = column_index_from_string(m.group(1))
            c2 = column_index_from_string(m.group(3))
            r1, r2 = int(m.group(2)), int(m.group(4))
            acc = 0.0
            for ci in range(min(c1, c2), max(c1, c2) + 1):
                for ri in range(min(r1, r2), max(r1, r2) + 1):
                    v = _value(bk, sheet, "%s%d" % (get_column_letter(ci), ri), depth + 1)
                    if v:
                        acc += v
            total += sign * acc
            continue
        if _REF_RE.match(term):
            v = _value(bk, sheet, term, depth + 1)
            total += sign * (v or 0.0)
            continue
        n = _num(term)
        if n is None:
            return None
        total += sign * n
    return total

# ===================== 公式形态与格式辅助 =====================


def _is_range_sum(raw, col, r1, r2):
    """判断是否为 =SUM(col r1:col r2) 形态（区域不得改小或改错）。"""
    m = _SUM_RE.match(_norm_formula(raw))
    if not m:
        return False
    return (m.group(1) == col and m.group(3) == col
            and int(m.group(2)) == r1 and int(m.group(4)) == r2)


def _is_row_sum(raw, cols, row):
    """判断"数量(吨)"是否为对本行六个大类的求和公式（连续区域或六个引用相加均可）。"""
    s = _norm_formula(raw)
    if not s:
        return False
    m = _SUM_RE.match(s)
    if m:
        return (int(m.group(2)) == row and int(m.group(4)) == row
                and set([m.group(1), m.group(3)]) == set([cols[0], cols[-1]]))
    return set(_REFS_RE.findall(s)) == set("%s%d" % (c, row) for c in cols)


def _fill_rgb(cell):
    try:
        fg = cell.fill.fgColor
        rgb = fg.rgb if fg is not None else None
    except Exception:
        return None
    if not isinstance(rgb, str):
        return None
    return rgb[-6:].upper()


def _has_fill(cell):
    try:
        pt = cell.fill.patternType
    except Exception:
        return False
    if not pt:
        return False
    return _fill_rgb(cell) not in (None, "000000", "FFFFFF")


def _font_of(cell):
    f = cell.font
    return (f.name, float(f.sz) if f.sz else None, bool(f.b),
            (f.color.rgb if (f.color is not None and isinstance(f.color.rgb, str)) else None))


def _centered(cell, need_h=True):
    a = cell.alignment
    if need_h and a.horizontal != "center":
        return False
    return a.vertical == "center"


def _nf_2dec(cell):
    nf = cell.number_format or ""
    return "0.00" in nf


def _occupied(ws, max_row=400, max_col=60):
    """遍历工作表中所有有内容的单元格。"""
    out = []
    mr = min(ws.max_row or 1, max_row)
    mc = min(ws.max_column or 1, max_col)
    for row in ws.iter_rows(min_row=1, max_row=mr, min_col=1, max_col=mc):
        for cell in row:
            if not _blank(cell.value):
                out.append(cell)
    return out

# ===================== 由基线推导的坐标集合 =====================


def _header_cells():
    out = []
    for row, start in _HEADER_BLOCKS:
        si = column_index_from_string(start)
        for k, label in enumerate(_HEADER_LABELS):
            out.append(("%s%d" % (get_column_letter(si + k), row), label))
    return out


def _qty_cells():
    """每个客户行的"数量(吨)"单元格 -> (六个大类列, 行号)。"""
    out = {}
    for _s, _g, _n, _ccol, qcol, cats, r1, r2, _sub in _GROUPS:
        for row in range(r1, r2 + 1):
            out["%s%d" % (qcol, row)] = (cats, row)
    return out


def _cat_cells():
    """所有客户行的六个大类单元格 -> 所属客户名。"""
    out = {}
    for _s, _g, _n, ccol, _q, cats, r1, r2, _sub in _GROUPS:
        for row in range(r1, r2 + 1):
            name = _CUSTOMER_CELLS[_customer_cell(ccol, row)]
            for c in cats:
                out["%s%d" % (c, row)] = name
    return out


_HEADER_CELLS = _header_cells()
_QTY_CELLS = _qty_cells()
_CAT_CELLS = _cat_cells()

_TEXT_CELLS = {}
_TEXT_CELLS[TITLE_CELL] = TITLE_TEXT
_TEXT_CELLS[_MAKER_CELL] = _MAKER_TEXT
_TEXT_CELLS[_TOTAL_LABEL_CELL] = _TOTAL_TEXT
for _c in _SUBTOTAL_LABEL_CELLS:
    _TEXT_CELLS[_c] = _SUBTOTAL_TEXT
for _c in _PIGLET_CELLS:
    _TEXT_CELLS[_c] = "乳猪料"
for _c, _v in _GROUP_SERIALS.items():
    _TEXT_CELLS[_c] = str(_v)
for _s, _gc, _gn, _cc, _qc, _ct, _r1, _r2, _sb in _GROUPS:
    _TEXT_CELLS[_gc] = _gn
for _c, _v in _HEADER_CELLS:
    _TEXT_CELLS[_c] = _v
for _c, _v in _CUSTOMER_CELLS.items():
    _TEXT_CELLS[_c] = _v

# 既有数值/公式单元格（含求和结果），用于两位小数格式与"未被清空"校验
_NUM_CELLS = set(_BASE_CATEGORY_VALUES) | set(_QTY_CELLS) | set([_TOTAL_FORMULA_CELL])
for _row, _lab, _cols, _r1, _r2 in _SUBTOTALS:
    for _c in _cols:
        _NUM_CELLS.add("%s%d" % (_c, _row))

_BASE_OCCUPIED = set(_TEXT_CELLS) | _NUM_CELLS
# 既有带底纹的单元格：标题、组号组名与"乳猪料"、表头、各小计块、右下角合计
_FILLED_CELLS = set([TITLE_CELL, _TOTAL_LABEL_CELL, _TOTAL_FORMULA_CELL])
_FILLED_CELLS |= set(_GROUP_SERIALS) | set(_PIGLET_CELLS) | set(g[1] for g in _GROUPS)
_FILLED_CELLS |= set(c for c, _l in _HEADER_CELLS)
for _row, _lab, _cols, _r1, _r2 in _SUBTOTALS:
    _FILLED_CELLS.add("%s%d" % (_lab, _row))
    for _c in _cols:
        _FILLED_CELLS.add("%s%d" % (_c, _row))
# 原本空白的大类单元格（排除单据落位行，避免与"连塘"专项扣分重复）
_BLANK_CAT_CELLS = set(c for c in _CAT_CELLS
                       if c not in _BASE_CATEGORY_VALUES and c not in _TARGET_OTHER_CAT_CELLS)
assert len(_BASE_OCCUPIED) == 288 and len(_TEXT_CELLS) == 117 and len(_NUM_CELLS) == 171

# ===================== 统计清单区域识别 =====================

_TOTAL_WORDS = ("小计", "合计", "总计", "汇总", "总和", "总额")
_BASELINE_HEADER_POS = set([(4, 2), (4, 11), (18, 2), (18, 11), (39, 2)])


def _find_stat_region(bk):
    """在交付工作簿中寻找"统一客户 × 六个大类"的统计清单，返回区域描述或 None。"""
    best = None
    for name in bk["wb"].sheetnames:
        ws = bk["wb"][name]
        mr = min(ws.max_row or 1, 400)
        mc = min(ws.max_column or 1, 60)
        for row in range(1, mr + 1):
            texts = []
            for ci in range(1, mc + 1):
                t = _txt(ws.cell(row=row, column=ci).value)
                if t:
                    texts.append((t, ci))
            if not texts:
                continue
            cat_cols = {}
            for cat in _CATEGORIES:
                for t, ci in texts:
                    if _key(t) == _key(cat) and ci not in cat_cols.values():
                        cat_cols[cat] = ci
                        break
            if len(cat_cols) < 5:
                continue
            cust_col = None
            for t, ci in texts:
                if _key(t) == "客户" or ("客户" in t and "名" in t):
                    cust_col = ci
                    break
            if cust_col is None:
                continue
            if name == SHEET_NAME and (row, cust_col) in _BASELINE_HEADER_POS:
                continue  # "12月"表原有的五个渠道组表头，不是新增统计清单
            qty_col = None
            for t, ci in texts:
                if "数量" in t and ci != cust_col and ci not in cat_cols.values():
                    qty_col = ci
                    break
            cust_rows, total_rows = [], []
            for rr in range(row + 1, mr + 1):
                t = _txt(ws.cell(row=rr, column=cust_col).value)
                if not t:
                    continue
                if any(w in t for w in _TOTAL_WORDS):
                    total_rows.append(rr)
                else:
                    cust_rows.append(rr)
            if not cust_rows:
                continue
            cand = {"sheet": name, "header_row": row, "cust_col": cust_col,
                    "qty_col": qty_col, "cat_cols": cat_cols,
                    "cust_rows": cust_rows, "total_rows": total_rows,
                    "header_texts": dict((ci, t) for t, ci in texts)}
            if best is None or len(cand["cust_rows"]) > len(best["cust_rows"]):
                best = cand
    return best


def _stat_names(bk, st):
    ws = bk["wb"][st["sheet"]]
    return [_txt(ws.cell(row=r, column=st["cust_col"]).value) for r in st["cust_rows"]]


def _stat_row_values(bk, st, row):
    """取统计清单某行的六个大类数值（缺列返回 None）。"""
    out = []
    for cat in _CATEGORIES:
        ci = st["cat_cols"].get(cat)
        if ci is None:
            out.append(None)
        else:
            out.append(_value(bk, st["sheet"], "%s%d" % (get_column_letter(ci), row)))
    return out

# ===================== 单据落位行定位 =====================


def _target_info(bk):
    """定位"连塘合作养殖社"在"12月"表中的位置及其数量列、六个大类列。"""
    info = {"coords": [], "row": _TARGET_ROW, "qty": _TARGET_QTY_CELL,
            "cats": ["%s%d" % (c, _TARGET_ROW) for c in _RIGHT_CATS]}
    if SHEET_NAME not in bk["wb"].sheetnames:
        return info
    ws = bk["wb"][SHEET_NAME]
    for cell in _occupied(ws, max_row=200, max_col=30):
        if _txt(cell.value) == _TARGET_CUSTOMER:
            info["coords"].append(cell.coordinate)
    if len(info["coords"]) == 1:
        coord = info["coords"][0]
        m = _REF_RE.match(coord)
        if m:
            col, row = m.group(1), int(m.group(2))
            qcol, cats = ("C", _LEFT_CATS) if col == "B" else ("L", _RIGHT_CATS)
            if col in ("B", "K"):
                info["row"] = row
                info["qty"] = "%s%d" % (qcol, row)
                info["cats"] = ["%s%d" % (c, row) for c in cats]
    return info


def _row_cells(ws, row, max_col=18):
    return [ws.cell(row=row, column=ci) for ci in range(1, max_col + 1)]


# ===================== 一、加分项检查 =====================


def _p01(ctx):
    """+1 重量 1600KG 换算为 1.6 吨写入，未直接填 1600 也未误用包数/单价/金额。"""
    if ctx["ws12"] is None:
        return False
    tgt = ctx["tgt"]
    vals = [_value(ctx["bk"], SHEET_NAME, c) for c in tgt["cats"]]
    total = sum(v for v in vals if v)
    got = _close(total, _TARGET_AFTER) or any(_close(v, 1.6) for v in vals)
    for cell in _row_cells(ctx["ws12"], tgt["row"]):
        n = _num(cell.value)
        if n is not None and (_close(n, _RAW_WEIGHT) or any(_close(n, w) for w in _WRONG_NUMBERS)):
            return False
    return got


def _p02(ctx):
    """+3 "育成期配合饲料"归入"保育/育肥料"，1.6 吨落在该列（其余五个大类保持空白）。"""
    if ctx["ws12"] is None:
        return False
    cats = ctx["tgt"]["cats"]
    if not _close(_value(ctx["bk"], SHEET_NAME, cats[1]), _TARGET_AFTER):
        return False
    for c in cats[:1] + cats[2:]:
        if not _blank(_raw(ctx["bk"], SHEET_NAME, c)):
            return False
    return True


def _p03(ctx):
    """+1 并入第 41 行已有客户行，未新增重复客户行、额外行列或另起小表。"""
    ws = ctx["ws12"]
    if ws is None or len(ctx["tgt"]["coords"]) != 1:
        return False
    if ctx["tgt"]["coords"][0] != "K%d" % _TARGET_ROW:
        return False
    vals = [_value(ctx["bk"], SHEET_NAME, c) for c in ctx["tgt"]["cats"]]
    if not _close(sum(v for v in vals if v), _TARGET_AFTER):
        return False  # 该行未发生 +1.6 的录入，谈不上"并入已有行"
    for coord, name in _CUSTOMER_CELLS.items():
        if _txt(_raw(ctx["bk"], SHEET_NAME, coord)) != name:
            return False
    return not ctx["extra_in_block"]


def _p04(ctx):
    """+1 数量(吨)2.52、西禾组小计 385.48、右下角合计 741.56 三处同步且自洽。"""
    if ctx["ws12"] is None:
        return False
    bk = ctx["bk"]
    return (_close(_value(bk, SHEET_NAME, ctx["tgt"]["qty"]), _TARGET_AFTER)
            and _close(_value(bk, SHEET_NAME, _GROUP5_SUBTOTAL_CELL), _GROUP5_SUBTOTAL_AFTER)
            and _close(_value(bk, SHEET_NAME, _TOTAL_FORMULA_CELL), _GRAND_TOTAL_AFTER))

def _p05(ctx):
    """+5 五个渠道组、左右两栏合并为一张统一客户清单，覆盖 57 个客户，无遗漏无重复。"""
    st = ctx["st"]
    if st is None:
        return False
    names = _stat_names(ctx["bk"], st)
    if len(names) != 57 or len(set(names)) != 57:
        return False
    return set(names) == set(_EXPECTED_MATRIX)


def _p06(ctx):
    """+5 57 个客户各大类数值逐格对位、未发生错列（右栏 M~R 正确映射）。"""
    st = ctx["st"]
    if st is None or len(st["cat_cols"]) != 6:
        return False
    names = _stat_names(ctx["bk"], st)
    if set(names) != set(_EXPECTED_MATRIX):
        return False
    for row, name in zip(st["cust_rows"], names):
        want = _EXPECTED_MATRIX.get(name)
        got = _stat_row_values(ctx["bk"], st, row)
        for w, g in zip(want, got):
            if w is None:
                if g not in (None, 0, 0.0):
                    return False
            elif not _close(g, w):
                return False
    return True


def _p07(ctx):
    """+3 六个大类总量正确：0.26 / 635.98 / 99.06 / 1.04 / 4.58 / 0.64。"""
    st = ctx["st"]
    if st is None or len(st["cat_cols"]) != 6:
        return False
    got = []
    for cat in _CATEGORIES:
        ci = st["cat_cols"][cat]
        acc = 0.0
        for row in st["cust_rows"]:
            v = _value(ctx["bk"], st["sheet"], "%s%d" % (get_column_letter(ci), row))
            if v:
                acc += v
        got.append(round(acc, 2))
    return all(_close(g, w) for g, w in zip(got, _CATEGORY_TOTALS))


def _p08(ctx):
    """+1 统计按六个大类分列，列名与指令给出的六个大类一致且齐全。"""
    st = ctx["st"]
    return st is not None and len(st["cat_cols"]) == 6


def _p09(ctx):
    """+1 每个客户给出"数量(吨)"行合计，并有清单总计行，总计 741.56 与各列自洽。"""
    st = ctx["st"]
    if st is None or st["qty_col"] is None or len(st["cat_cols"]) != 6:
        return False
    qcol = get_column_letter(st["qty_col"])
    for row in st["cust_rows"]:
        qty = _value(ctx["bk"], st["sheet"], "%s%d" % (qcol, row))
        rowsum = sum(v for v in _stat_row_values(ctx["bk"], st, row) if v)
        if not _close(qty, rowsum, 0.011):
            return False
    for row in st["total_rows"]:
        if _close(_value(ctx["bk"], st["sheet"], "%s%d" % (qcol, row)), _GRAND_TOTAL_AFTER, 0.011):
            return True
    return False


def _p10(ctx):
    """+1 统计结果置于新增工作表或"12月"表数据区域之外，未覆盖已有客户行与小计行。"""
    st = ctx["st"]
    if st is None or ctx["ws12"] is None or ctx["cleared_rows"]:
        return False
    if st["sheet"] != SHEET_NAME:
        return True
    cols = [st["cust_col"], st["qty_col"]] + list(st["cat_cols"].values())
    rows = [st["header_row"]] + st["cust_rows"] + st["total_rows"]
    for ci in [c for c in cols if c]:
        for ri in rows:
            if ri <= 58 and ci <= 18:
                return False
    return True


def _p11(ctx):
    """+1 统计清单带完整可读表头（客户、数量(吨)及六个大类名），数值沿用两位小数格式。"""
    st = ctx["st"]
    if st is None or st["qty_col"] is None or len(st["cat_cols"]) != 6:
        return False
    if "数量" not in st["header_texts"].get(st["qty_col"], ""):
        return False
    ws = ctx["bk"]["wb"][st["sheet"]]
    cols = [st["qty_col"]] + list(st["cat_cols"].values())
    for ri in st["cust_rows"] + st["total_rows"]:
        for ci in cols:
            cell = ws.cell(row=ri, column=ci)
            if _blank(cell.value):
                continue
            if _num(cell.value) is None and not _is_formula(cell.value):
                return False
            if not _nf_2dec(cell):
                return False
    return True

# ===================== 二、减分项检查（返回 True 表示命中扣分） =====================


def _d_title(ctx):
    ws = ctx["ws12"]
    if ws is None:
        return True
    cell = ws[TITLE_CELL]
    if _txt(cell.value) != TITLE_TEXT:
        return True
    if TITLE_MERGE not in [str(r) for r in ws.merged_cells.ranges]:
        return True
    name, sz, bold, color = _font_of(cell)
    if not bold or not _close(sz, 15.0):
        return True
    if not (isinstance(color, str) and color.upper().endswith("FFFFFF")):
        return True
    if _fill_rgb(cell) != _FILL_TITLE or not _centered(cell):
        return True
    return False


def _d_header_text(ctx):
    if ctx["ws12"] is None:
        return True
    for coord, label in _HEADER_CELLS:
        if _key(_raw(ctx["bk"], SHEET_NAME, coord)) != _key(label):
            return True
    return False


def _d_header_style(ctx):
    ws = ctx["ws12"]
    if ws is None:
        return True
    for coord, _label in _HEADER_CELLS:
        cell = ws[coord]
        if _fill_rgb(cell) != _FILL_HEADER or not cell.font.b or not _centered(cell):
            return True
    return False


def _d_piglet(ctx):
    if ctx["ws12"] is None:
        return True
    return any(_txt(_raw(ctx["bk"], SHEET_NAME, c)) != "乳猪料" for c in _PIGLET_CELLS)


def _d_serial(ctx):
    ws = ctx["ws12"]
    if ws is None:
        return True
    for coord, val in _GROUP_SERIALS.items():
        cell = ws[coord]
        if _num(cell.value) != float(val):
            return True
        if _fill_rgb(cell) != _FILL_SERIAL or not cell.font.b or not _centered(cell):
            return True
    return False


def _d_group_name(ctx):
    if ctx["ws12"] is None:
        return True
    for _s, gcell, gname, _c, _q, _t, _r1, _r2, _sb in _GROUPS:
        if _txt(_raw(ctx["bk"], SHEET_NAME, gcell)) != gname:
            return True
    return False


def _d_maker(ctx):
    if ctx["ws12"] is None:
        return True
    return _txt(_raw(ctx["bk"], SHEET_NAME, _MAKER_CELL)) != _MAKER_TEXT

def _d_row57_lost(ctx):
    """-3 第 57 行 B~I"小计"行被删除或清空，"小计"字样与七个求和公式一并丢失。"""
    if ctx["ws12"] is None:
        return True
    if _txt(_raw(ctx["bk"], SHEET_NAME, "B57")) != _SUBTOTAL_TEXT:
        return True
    for col in "C" + _LEFT_CATS:
        if not _is_range_sum(_raw(ctx["bk"], SHEET_NAME, "%s57" % col), col, 40, 46):
            return True
    return False


def _d_row57_style(ctx):
    """-1 第 57 行 B~I 加粗、水平垂直居中、E2F0D9 底纹、数值列 0.00 格式。"""
    ws = ctx["ws12"]
    if ws is None:
        return True
    for col in "B" + "C" + _LEFT_CATS:
        cell = ws["%s57" % col]
        if not cell.font.b or not _centered(cell) or _fill_rgb(cell) != _FILL_SUBTOTAL:
            return True
        if col != "B" and not _nf_2dec(cell):
            return True
    return False


def _d_subtotal_formula(ctx):
    """-1 第 12、25、54 行小计求和公式被写死数值或求和区域被改小、改错。"""
    if ctx["ws12"] is None:
        return True
    for row, _lab, cols, r1, r2 in _SUBTOTALS:
        if row == 57:
            continue
        for col in cols:
            if not _is_range_sum(_raw(ctx["bk"], SHEET_NAME, "%s%d" % (col, row)), col, r1, r2):
                return True
    return False


def _d_total_formula(ctx):
    """-1 右下角合计由 =C12+L12+C25+L54+C57 被改写为写死数值。"""
    if ctx["ws12"] is None:
        return True
    raw = _raw(ctx["bk"], SHEET_NAME, _TOTAL_FORMULA_CELL)
    s = _norm_formula(raw)
    if not s:
        return True
    refs = _REFS_RE.findall(s)
    if len(refs) != 5 or set(refs) != set(_TOTAL_FORMULA_REFS):
        return True
    # 公式中不得夹带写死的数字常量
    return bool(re.search(r"(?<![A-Z])\d+(?:\.\d+)?(?![0-9])", _REFS_RE.sub("", s)))


def _d_total_label(ctx):
    """-1 "小计""合计"文字被改写为其它称谓，或小计行、合计行被整行删除、移位。"""
    if ctx["ws12"] is None:
        return True
    for coord in _SUBTOTAL_LABEL_CELLS:
        if coord == "B57":
            continue
        if _txt(_raw(ctx["bk"], SHEET_NAME, coord)) != _SUBTOTAL_TEXT:
            return True
    return _txt(_raw(ctx["bk"], SHEET_NAME, _TOTAL_LABEL_CELL)) != _TOTAL_TEXT

def _d_customer_names(ctx):
    """-3 57 个客户名称被删除、改写、增补，或所属渠道组、所在行次被改变。"""
    if ctx["ws12"] is None:
        return True
    for coord, name in _CUSTOMER_CELLS.items():
        if _txt(_raw(ctx["bk"], SHEET_NAME, coord)) != name:
            return True
    return False


def _d_other_values(ctx):
    """-3 单据涉及行以外任一客户的大类数值被改动。"""
    if ctx["ws12"] is None:
        return True
    for coord, want in _BASE_CATEGORY_VALUES.items():
        if coord == _TARGET_CAT_CELL:
            continue
        if not _close(_value(ctx["bk"], SHEET_NAME, coord), want):
            return True
    return False


def _d_qty_formula(ctx):
    """-3 客户行"数量(吨)"的求和公式被改为写死数值。"""
    if ctx["ws12"] is None:
        return True
    for coord, (cols, row) in _QTY_CELLS.items():
        if not _is_row_sum(_raw(ctx["bk"], SHEET_NAME, coord), cols, row):
            return True
    return False


def _d_blank_placeholder(ctx):
    """-1 原本空白的大类单元格被填入 0、"-"、"无"等占位符。"""
    if ctx["ws12"] is None:
        return True
    for coord in _BLANK_CAT_CELLS:
        if not _blank(_raw(ctx["bk"], SHEET_NAME, coord)):
            return True
    return False


def _d_noise(ctx):
    """-1 车牌、电话、单据编号、经办人姓名、"现结""未付款"等无关信息被写入表内。"""
    for text in ctx["all_texts"]:
        for token in _NOISE_TOKENS:
            if token in text:
                return True
    return False


def _d_wrong_numbers(ctx):
    """-1 单据"数量 64（包）""单价 86""金额 5504"被写入数量列或新增列。"""
    for cell in ctx["all_cells"]:
        n = _num(cell.value)
        if n is not None and any(_close(n, w) for w in _WRONG_NUMBERS):
            return True
    for text in ctx["all_texts"]:
        if "64包" in text or "5504" in text:
            return True
    return False

def _d_number_format(ctx):
    """-1 数值区域的两位小数格式(0.00)被改为常规、文本、百分比或科学计数。"""
    ws = ctx["ws12"]
    if ws is None:
        return True
    for coord in _NUM_CELLS:
        if not _nf_2dec(ws[coord]):
            return True
    return False


def _d_font(ctx):
    """-1 出现黑体、微软雅黑、MS PGothic 等明显异体字，或同行列内字号不一。"""
    ws = ctx["ws12"]
    if ws is None:
        return True
    for cell in _occupied(ws, max_row=200, max_col=30):
        name = cell.font.name
        if name not in _FONT_OK:
            return True
    for coord in _BASE_OCCUPIED:
        _n, sz, _b, _c = _font_of(ws[coord])
        if sz is not None and sz not in _FONT_SIZES_OK:
            return True
        if coord == TITLE_CELL and not _close(sz, 15.0):
            return True
        if coord != TITLE_CELL and sz is not None and _close(sz, 15.0):
            return True
    return False


def _d_alignment(ctx):
    """-1 有内容单元格的水平居中、垂直居中或自动换行属性被取消。"""
    ws = ctx["ws12"]
    if ws is None:
        return True
    for coord in _BASE_OCCUPIED:
        if coord in _ALIGN_EXEMPT:
            continue
        cell = ws[coord]
        a = cell.alignment
        if a.vertical != "center" or not a.wrap_text:
            return True
        if coord in _CENTER_CELLS and a.horizontal != "center":
            return True
    return False


def _d_dimension(ctx):
    """-1 列宽、行高被改动，或行、列被隐藏。"""
    ws = ctx["ws12"]
    if ws is None:
        return True
    for ci in range(1, 19):
        letter = get_column_letter(ci)
        dim = ws.column_dimensions.get(letter)
        want = _COL_WIDTHS.get(letter)
        got = dim.width if (dim is not None and dim.width) else None
        if dim is not None and dim.hidden:
            return True
        if want is None:
            if got is not None and abs(got - 8.43) > 1.0:
                return True
        elif got is None or abs(got - want) > 0.1:
            return True
    for ri, want in _ROW_HEIGHTS.items():
        dim = ws.row_dimensions.get(ri)
        if dim is not None and dim.hidden:
            return True
        got = dim.height if (dim is not None and dim.height) else None
        if got is None or abs(got - want) > 0.6:
            return True
    return False

def _d_shift(ctx):
    """-1 在数据区插入或删除行列，使客户名与其数量、大类列不再对应。"""
    ws = ctx["ws12"]
    if ws is None:
        return True
    anchors = dict((c, t) for c, t in _TEXT_CELLS.items())
    misplaced = [c for c, t in anchors.items() if _txt(_raw(ctx["bk"], SHEET_NAME, c)) != t]
    if not misplaced:
        return False
    present = {}
    for cell in _occupied(ws, max_row=200, max_col=30):
        present.setdefault(_txt(cell.value), []).append(cell.coordinate)
    for coord in misplaced:
        want = anchors[coord]
        for got in present.get(want, []):
            if got != coord:
                return True
    return False


def _d_merge(ctx):
    """-1 B1:R1 之外新增合并单元格，或取消已有合并。"""
    ws = ctx["ws12"]
    if ws is None:
        return True
    return set(str(r) for r in ws.merged_cells.ranges) != set([TITLE_MERGE])


def _d_extra(ctx):
    """-1 新增与两项任务无关的辅助列、备注列、批注、条件格式或高亮标色。"""
    ws = ctx["ws12"]
    if ws is None:
        return True
    try:
        if list(ws.conditional_formatting):
            return True
    except Exception:
        pass
    for cell in _occupied(ws, max_row=200, max_col=30):
        if cell.comment is not None:
            return True
    # 既有无底纹的单元格（客户名、数值、制表说明）被加上底纹即视为高亮标色
    for coord in _BASE_OCCUPIED:
        if coord in _FILLED_CELLS:
            continue
        if _has_fill(ws[coord]):
            return True
    return bool(ctx["extra_new"] or ctx["stray_outside"])


def _d_cleared(ctx):
    """-3 为放置统计结果而删除、清空或覆盖"12月"表已有的客户行、小计行、合计行。"""
    return ctx["ws12"] is None or bool(ctx["cleared_rows"])


def _d_sheet_missing(ctx):
    """-1 "12月"工作表被删除、重命名，或被移出为独立文件。"""
    return ctx["ws12"] is None


def _d_blank_sheet(ctx):
    """-1 新增的空白工作表、临时工作表未清理。"""
    for name in ctx["bk"]["wb"].sheetnames:
        ws = ctx["bk"]["wb"][name]
        if not _occupied(ws, max_row=200, max_col=40):
            return True
    return False


def _d_flat_file(ctx):
    """-1 交付为 .csv/.txt 等只能保存单表、丢失公式与格式的文件类型。"""
    return bool(ctx["flat_file"])

# ===================== 规则表（顺序与 _MAX_DELTAS_ORDER、细则.txt 一致） =====================

_PLUS_RULES = [
    (1, "单据“重量 1600KG”换算为 1.6 吨写入表中（未直接填 1600，也未误用“数量 64 包”“单价 86”“金额 5504”等字段）", _p01),
    (3, "“育成期配合饲料”在六个大类中被判定归入“保育/育肥料”，1.6 吨落在该列，该客户其余五个大类保持空白", _p02),
    (1, "定位到第 5 组西禾区域合作组第 41 行已有的“连塘合作养殖社”并并入该行，未新增重复客户行、额外行列或另起小表", _p03),
    (1, "“连塘合作养殖社”数量(吨)=2.52、西禾区域合作组小计=385.48、右下角全表合计=741.56，三处同步更新且相互自洽", _p04),
    (5, "将 5 个渠道组、左右两栏的客户合并为一张统一客户清单，覆盖全部 57 个客户，无遗漏、无重复计入", _p05),
    (5, "57 个客户的各大类数值逐格对位、未发生错列（右栏 M~R 列按开口料~浮水鱼料口径正确映射）", _p06),
    (3, "统计结果六个大类的总量正确：开口料 0.26、保育/育肥料 635.98、蛋禽料 99.06、水禽料 1.04、沉水鱼料 4.58、浮水鱼料 0.64", _p07),
    (1, "统计按六个大类分列，列名为“开口料”“保育/育肥料”“蛋禽料”“水禽料”“沉水鱼料”“浮水鱼料”，与指令一致且齐全", _p08),
    (1, "为每个客户给出“数量(吨)”行合计（等于该客户六个大类之和），并给出清单总计行，总计 741.56 自洽", _p09),
    (1, "统计结果置于新增工作表或“12月”表数据区域之外的空白处，未覆盖、挤占已有客户行与小计行", _p10),
    (1, "统计清单带完整可读表头（“客户”“数量(吨)”及六个大类名），数值沿用两位小数显示格式", _p11),
]

_MINUS_RULES = [
    (-1, "B1:R1 合并标题“2026年12月份区域渠道销量统计表”不满足 15 磅加粗、白色字、深蓝底纹(1F4E78)、水平垂直居中、B1:R1 合并", _d_title),
    (-1, "第 4/18/39 行表头“客户”“数量(吨)”及六个大类名（左栏 B~I、右栏 K~R）文字被改写、顺序被调换或位置被移动", _d_header_text),
    (-1, "表头行浅蓝底纹(DDEBF7)、加粗、居中被取消或改色", _d_header_style),
    (-1, "D3、M3 位置的“乳猪料”标注被删除或改写", _d_piglet),
    (-1, "A3=1、A17=2、A38=3、J3=4、J17=5 五个组序号（含蓝色底纹 5B9BD5、加粗居中）被删除、改值或改格式", _d_serial),
    (-1, "“东湾渠道一组”“海岚客户维护组”“云澜养殖拓展组”“北岭养殖服务组”“西禾区域合作组”五个组名被删除、改写或移位", _d_group_name),
    (-1, "P58“制表：渠道运营部”被删除、改写或移动到其它位置", _d_maker),
    (-3, "第 57 行 B~I 列“小计”行被删除或清空，“小计”字样与 =SUM(C40:C46)～=SUM(I40:I46) 七个求和公式一并丢失", _d_row57_lost),
    (-1, "第 57 行 B~I 列不满足加粗、水平垂直居中、浅绿底纹(E2F0D9)、数值列两位小数格式(0.00)", _d_row57_style),
    (-1, "第 12、25、54 行“小计”行的求和公式被改为写死数值，或求和区域被改小、改错", _d_subtotal_formula),
    (-1, "右下角“合计”单元格由联动公式 =C12+L12+C25+L54+C57 被改写为写死数值", _d_total_formula),
    (-1, "“小计”“合计”文字被改写为其它称谓，或小计行、合计行被整行删除、位置被移动", _d_total_label),
    (-3, "57 个客户名称被删除、改写、增补，或客户所属渠道组、所在行次被改变", _d_customer_names),
    (-3, "单据涉及行以外任一客户的大类数值被改动", _d_other_values),
    (-3, "客户行“数量(吨)”的求和公式（如 =SUM(D5:I5)、=SUM(M19:R19)）被改为写死数值", _d_qty_formula),
    (-1, "原本空白的大类单元格被填入 0、“-”、“无”等占位符，使空值变为数值或文本", _d_blank_placeholder),
    (-1, "单据中的车牌“鲁K6A12Y”、联系电话、编号 DH202604180012、开单人/会计/出纳/仓管/提货人姓名、“现结”“未付款”等无关信息被写入表内", _d_noise),
    (-1, "单据的“数量 64（包）”“单价 86”“金额 5504”被写入大类数量列或新增列，与“吨”口径混用", _d_wrong_numbers),
    (-1, "数值区域的两位小数格式(0.00)被改为常规、文本、百分比或科学计数", _d_number_format),
    (-1, "单元格字体出现可见的不统一（出现黑体、微软雅黑、MS PGothic 等明显异体字，或同一行列内字号大小不一）", _d_font),
    (-1, "有内容的单元格的水平居中、垂直居中或自动换行属性被取消", _d_alignment),
    (-1, "列宽、行高被改动或行、列被隐藏，导致数值显示为 #### 或需重新调整才能阅读", _d_dimension),
    (-1, "在数据区插入或删除行列，使客户名与其数量、大类列不再对应", _d_shift),
    (-1, "B1:R1 之外新增合并单元格，或取消已有合并，造成版面错位", _d_merge),
    (-1, "新增与两项任务无关的辅助列、备注列、批注、条件格式或高亮标色", _d_extra),
    (-3, "为放置统计结果而删除、清空或覆盖“12月”表已有的客户行、小计行、合计行", _d_cleared),
    (-1, "“12月”工作表被删除、重命名，或被移出为独立文件", _d_sheet_missing),
    (-1, "新增的空白工作表、临时工作表未清理，交付文件中留有无内容或过程性表页", _d_blank_sheet),
    (-1, "交付为 .csv/.txt 等只能保存单表、丢失公式与格式的文件类型", _d_flat_file),
]

assert [s for s, _d, _f in _PLUS_RULES] + [s for s, _d, _f in _MINUS_RULES] == _MAX_DELTAS_ORDER

# ===================== 上下文与维度1 =====================


def _build_ctx(bk, flat_file):
    ws12 = bk["wb"][SHEET_NAME] if SHEET_NAME in bk["wb"].sheetnames else None
    all_cells = []
    for name in bk["wb"].sheetnames:
        all_cells.extend(_occupied(bk["wb"][name], max_row=400, max_col=60))
    allowed = _BASE_OCCUPIED | set(_CAT_CELLS)
    base_texts = set(_TEXT_CELLS.values())
    extra = []
    extra_new = []
    stray = []
    cleared = []
    if ws12 is not None:
        st_cells = _stat_cells_in_sheet(bk, SHEET_NAME)
        for cell in _occupied(ws12, max_row=58, max_col=18):
            if cell.coordinate not in allowed:
                extra.append(cell.coordinate)
                if _txt(cell.value) not in base_texts:
                    extra_new.append(cell.coordinate)
        for cell in _occupied(ws12, max_row=400, max_col=60):
            if cell.row <= 58 and cell.column <= 18:
                continue
            if (cell.row, cell.column) not in st_cells:
                stray.append(cell.coordinate)
        by_row = {}
        for coord in _BASE_OCCUPIED:
            m = _REF_RE.match(coord)
            by_row.setdefault(int(m.group(2)), []).append(coord)
        for row, coords in by_row.items():
            if all(_blank(ws12[c].value) for c in coords):
                cleared.append(row)
    return {
        "bk": bk,
        "ws12": ws12,
        "st": _find_stat_region(bk),
        "tgt": _target_info(bk),
        "all_cells": all_cells,
        "all_texts": [_txt(c.value) for c in all_cells if isinstance(c.value, str)],
        "extra_in_block": extra,
        "extra_new": extra_new,
        "stray_outside": stray,
        "cleared_rows": sorted(cleared),
        "flat_file": flat_file,
    }


def _stat_cells_in_sheet(bk, sheet):
    """统计清单若被放在指定工作表内，返回其占用的 (行, 列) 集合。"""
    st = _find_stat_region(bk)
    if st is None or st["sheet"] != sheet:
        return set()
    cols = [c for c in [st["cust_col"], st["qty_col"]] + list(st["cat_cols"].values()) if c]
    rows = [st["header_row"]] + st["cust_rows"] + st["total_rows"]
    out = set()
    lo, hi = min(cols), max(cols)
    for ri in range(min(rows), max(rows) + 3):
        for ci in range(lo, hi + 1):
            out.add((ri, ci))
    return out


def _dim1(ctx):
    """维度1：交付物可打开、"12月"表仍在、内容为可编辑单元格与公式、行列未大面积错位。"""
    if ctx["flat_file"]:
        return False, "交付文件为 .csv/.txt 等纯文本类型，丢失工作表、公式与格式"
    ws = ctx["ws12"]
    if ws is None:
        return False, '工作簿中不存在"12月"工作表（被删除或重命名）'
    if getattr(ws, "_images", None):
        return False, '"12月"表中存在图片对象，疑似表格被截图粘贴替换'
    hit = sum(1 for coord, name in _CUSTOMER_CELLS.items()
              if _txt(_raw(ctx["bk"], SHEET_NAME, coord)) == name)
    if hit < 40:
        return False, '"12月"表 57 个客户名仅有 %d 个仍在原行列位置，行列大面积错位' % hit
    live = 0
    for coord in _NUM_CELLS:
        raw = ws[coord].value
        if _is_formula(raw) or isinstance(raw, (int, float)):
            live += 1
        elif _blank(raw):
            live += 1
    if live < int(len(_NUM_CELLS) * 0.8):
        return False, '"12月"表数值区域中 %d/%d 格变为不可计算的文本' % (len(_NUM_CELLS) - live, len(_NUM_CELLS))
    return True, '交付 1 个可正常打开的 Excel 工作簿（.xlsx/.xlsm），"12月"表仍在且为可逐格点选、可编辑的单元格与公式'

# ===================== 对外接口 =====================

_XL_EXT = (".xlsx", ".xlsm")
_FLAT_EXT = (".csv", ".txt")


def _find_target_file(dir_path):
    """在目录中定位待评测交付文件；返回 (路径, 是否为 csv/txt 纯文本交付)。

    只看该目录本层，找不到则抛 FileNotFoundError，由 evaluate() 统一转成
    "未找到文件"的返回结构。
    """
    names = [n for n in sorted(os.listdir(dir_path))
             if not n.startswith("~$") and os.path.isfile(os.path.join(dir_path, n))]
    xl = [n for n in names if n.lower().endswith(_XL_EXT)]
    if xl:
        xl.sort(key=lambda n: (0 if ("整理" in n or "统计" in n) else 1, n))
        return os.path.join(dir_path, xl[0]), False
    flat = [n for n in names if n.lower().endswith(_FLAT_EXT)]
    if flat:
        return os.path.join(dir_path, flat[0]), True
    raise FileNotFoundError("目录中未找到 .xlsx 文件：%s" % dir_path)


def evaluate(dir_path):
    """按接口统一约定返回结构化 dict。

    目录不存在或目录内没有交付文件时，返回"未找到文件"的空结果
    （各评分字段留空/归零）；其余脚本自身崩溃由通用兜底捕获并返回 status='error'。
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
        "max_score": MAX_SCORE_TOTAL,
    }
    try:
        if not dir_path or not os.path.isdir(dir_path):
            raise FileNotFoundError("目录不存在或不是目录：%s" % dir_path)
        file_path, flat_file = _find_target_file(dir_path)
        result["file_name"] = os.path.basename(file_path)
        return _evaluate_impl(file_path, flat_file, result)
    except FileNotFoundError as exc:
        # 未找到待评文件：不进入维度判定，dim1_pass 留空，分值一并归零
        result["file_name"] = ""
        result["status"] = "error"
        result["error"] = str(exc)
        result["dim1_pass"] = ""
        result["dim1_reason"] = ""
        result["dim2_items"] = []
        result["total_score"] = 0
        result["max_score"] = 0
        return result
    except Exception as exc:
        result["status"] = "error"
        result["error"] = "脚本运行异常：%s: %s" % (type(exc).__name__, exc)
        result["dim1_pass"] = False
        result["dim1_reason"] = ("维度一未通过：文件无法打开或解析（%s: %s），完成度总分记 0"
                                % (type(exc).__name__, exc))
        result["dim2_items"] = []
        result["total_score"] = 0
        return result


def _evaluate_impl(file_path, flat_file, result):
    if flat_file:
        result["dim1_pass"] = False
        result["dim1_reason"] = ("维度一未通过：交付文件为 .csv/.txt 等纯文本类型，"
                                "丢失工作表、公式与格式，完成度总分记 0")
        result["dim2_items"] = _zero_items()
        result["total_score"] = 0
        return result

    wb = openpyxl.load_workbook(file_path, data_only=False)
    try:
        wbc = openpyxl.load_workbook(file_path, data_only=True)
    except Exception:
        wbc = None
    ctx = _build_ctx({"wb": wb, "wbc": wbc}, flat_file)

    dim1_pass, dim1_reason = _dim1(ctx)
    result["dim1_pass"] = dim1_pass
    if dim1_pass:
        result["dim1_reason"] = dim1_reason
    else:
        # dim1_reason 统一带上「维度一未通过：…」前缀与计分后果，便于直接作为错误摘要展示
        result["dim1_reason"] = "维度一未通过：%s，完成度总分记 0" % dim1_reason

    results = []
    details = []
    for score, desc, fn in _PLUS_RULES:
        hit, detail = _run(fn, ctx)
        results.append("+%d：%s" % (score if hit else 0, desc))
        details.append(detail)
    for score, desc, fn in _MINUS_RULES:
        hit, detail = _run(fn, ctx)
        results.append("-%d：%s" % (abs(score) if hit else 0, desc))
        details.append(detail)

    # ========== 将 results 反解为结构化 dim2_items ==========
    if len(results) != len(_MAX_DELTAS_ORDER):
        raise RuntimeError("dim2 项数不一致：results=%d vs order=%d" % (len(results), len(_MAX_DELTAS_ORDER)))
    dim2_items = []
    total_score = 0
    # details 仅保留各判定处的取证说明供排查，不再随结果返回
    for raw, max_delta, _detail in zip(results, _MAX_DELTAS_ORDER, details):
        head, _sep, rule = raw.partition("：")
        try:
            delta = int(head.replace("+", "").strip())
        except ValueError:
            delta = 0
        hit = delta != 0
        dim2_items.append({"rule": rule, "max_delta": max_delta,
                           "delta": max_delta if hit else 0, "hit": hit})
        total_score += (max_delta if hit else 0)
    if not dim1_pass:
        for item in dim2_items:
            item["delta"] = 0
            item["hit"] = False
        total_score = 0
    result["dim2_items"] = dim2_items
    result["total_score"] = max(0, total_score)
    result["max_score"] = MAX_SCORE_TOTAL
    return result


def _run(fn, ctx):
    try:
        return bool(fn(ctx)), ""
    except Exception as exc:
        return False, "检查异常：%s: %s" % (type(exc).__name__, exc)


def _zero_items():
    items = []
    for score, desc, _fn in _PLUS_RULES + _MINUS_RULES:
        items.append({"rule": desc, "max_delta": score, "delta": 0, "hit": False})
    return items


if __name__ == "__main__":
    if len(sys.argv) >= 2:
        _dir = sys.argv[1]
    else:
        _dir = os.path.dirname(os.path.abspath(__file__))
    _out = evaluate(_dir)
    print(json.dumps(_out, ensure_ascii=False, indent=2))
