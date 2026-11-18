#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# pyright: reportUnknownMemberType=false, reportUnknownVariableType=false, reportUnknownArgumentType=false, reportUnknownParameterType=false, reportMissingParameterType=false, reportPrivateUsage=false, reportGeneralTypeIssues=false

import os
import re
import json
import zipfile
from xml.etree import ElementTree as ET

from openpyxl import load_workbook
from openpyxl.utils import get_column_letter, column_index_from_string

# 说明：按接口统一约定，本模块不修改全局标准输出，也不 print 主结果。
# 主结果统一由 evaluate() -> dict 返回；评估目录固定为"本脚本所在目录"，
# 只在该目录本层内定位并打开被评估的启用宏工作簿（.xlsm/.xlsb/.xls；
# 若目录内只有 .xlsx 则同样取用，交由维度一判负），不检索子目录、上级目录
# 或调用方的当前工作目录。evaluate() 保留 dir_path 形参仅为兼容既有调用，
# 传入任何值都会被忽略。本文件仅在 __main__ 下作为本地调试入口，才把 dict
# 序列化为 JSON 打印。
#
# 评测对象：「进账单打印.xlsx」宏按钮制作任务的交付件。指令只要求"新增/修复
# 宏按钮 + 触发第三张表指定区域的打印预览"，其余数据、公式、格式、控件、视图
# 状态均属指令未涉及部分，应保持不变：
#   加分项（13 项，合计 35 分）——仅针对宏与按钮功能；
#   扣分项（23 项）——针对指令未涉及、却被改动而破坏交付件可用性的部分。

SCRIPT_ID = "084"


def _script_dir():
    """脚本自身所在目录（交互式执行时退化为当前工作目录）。"""
    path = globals().get("__file__")
    if isinstance(path, str) and path:
        return os.path.dirname(os.path.abspath(path))
    return os.getcwd()


# 评估目标锁定为本脚本所在目录：评测时把本脚本与被评交付物放在同一个目录下即可。
# 不接受外部传入的目录，也不检索子目录与上级目录。
SCRIPT_DIR = _script_dir()

# ============================================================
# 细则基线常量：均为"指令未涉及、应保持不变"的绝对值
# ============================================================

DATA_SHEET = "支票明细"       # 第一张表：结算明细台账
FORM_SHEET = "填写"           # 第二张表：打印页
TICKET_SHEET = "划拨进账单"   # 第三张表：票面
SHEET_NAMES = (DATA_SHEET, FORM_SHEET, TICKET_SHEET)
SHEET_CODENAMES = ("Sheet1", "Sheet2", "Sheet3")

MACRO_EXTS = (".xlsm", ".xlsb", ".xls")
# 定位工作簿时的取用优先级：能承载宏的在前，.xlsx 兜底（交由维度一判负）
WORKBOOK_EXTS = MACRO_EXTS + (".xlsx",)
WORKBOOK_MACRO_CT = "application/vnd.ms-excel.sheet.macroEnabled.main+xml"
VBA_PART = "xl/vbaProject.bin"
MACROSHEET_PREFIX = "xl/macrosheets/"

# 按钮所在位置（图示红框标注）与不得遮挡的表头单元格
BUTTON_ZONE = "P2:R2"
HEADER_GUARD_ZONE = "P1:S1"
BUTTON_CAPTION = "打印进帐单"   # 去空格/换行后的标题文字（"帐"字写法须保留）

# 票面区域（图示红框）：自 A1 起至 U8/W9 一带
TICKET_PRINT_AREAS = ("$A$1:$U$8", "$A$1:$W$9")
# 输入文件既有的打印区域定义名称：仅沿用不得分，校正为 $A$1:$W$9 才算"在宏侧做了工作"
BASELINE_TICKET_PRINT_AREA = "$A$1:$U$8"
FORM_PRINT_AREA = "$A$1:$U$14"          # 填写!Print_Area

# 划拨进账单：17 处合并单元格
TICKET_MERGES = (
    "F1:G1", "H1:I1", "B2:H2", "I2:T2", "B3:H3", "I3:T3", "B4:H4", "I4:T4",
    "B5:H6", "D7:E7", "F7:G7", "J7:K7", "L7:O7", "P7:T7", "B8:G8", "B9:G9",
    "H8:T9",
)
# 划拨进账单：行高（磅）
TICKET_ROW_HEIGHTS = {1: 33.0, 2: 15.0, 4: 16.5, 5: 7.5, 6: 19.5, 7: 17.25, 8: 18.0}
# 划拨进账单：列宽（xl 存储值，即 openpyxl / OOXML 的 width）
TICKET_COL_WIDTHS = {
    "A": 12.125, "B": 7.5, "C": 4.5, "D": 3.0, "G": 3.375, "H": 15.5,
    "I": 2.0, "J": 1.75, "K": 2.0, "M": 2.125, "N": 2.0, "U": 9.0,
}
# 划拨进账单：页边距（英寸）
TICKET_MARGINS = {
    "left": 0.52, "right": 0.0, "top": 0.079, "bottom": 0.0,
    "header": 0.0, "footer": 0.0,
}
# 划拨进账单：底色与字体
TICKET_FILLS = {
    "B2": "FEF3C7", "B5": "FEF3C7", "B7": "FEF3C7", "B8": "FEF3C7", "B9": "FEF3C7",
    "I2": "0F766E", "I3": "0F766E", "I4": "0F766E",
    "J6": "E0F2FE", "T6": "E0F2FE", "H8": "DBEAFE", "W2": "0F766E",
}
TICKET_WHITE_BOLD = ("I2", "I3", "I4", "W2")   # 深绿底白色加粗字

# 支票明细：行高（磅）与列宽（存储值）
DATA_ROW_HEIGHTS = {1: 29.25, 2: 30.75}
DATA_DETAIL_HEIGHT = 24.0      # 明细行
DATA_SUMMARY_HEIGHT = 12.0     # 汇总区
DATA_COL_WIDTHS = {
    "A": 4.875, "B": 7.892, "C": 23.625, "D": 23.625, "E": 23.625,
    "F": 8.5, "G": 8.5, "H": 8.5, "I": 12.625, "J": 5.0, "K": 8.875,
    "L": 10.75, "M": 5.125, "N": 8.0, "O": 14.6, "P": 2.25, "Q": 2.5,
    "R": 6.0, "S": 12.625, "T": 7.0, "U": 8.975,
}
# 列宽允许 0.01 字符内的舍入差异（Excel"列宽"对话框显示值一致即可）
COL_WIDTH_TOL = 0.0101
ROW_HEIGHT_TOL = 0.05
MARGIN_TOL = 0.005

# 支票明细：明细区表头填充色
DATA_FILLS = {
    "A2": "0F766E", "O2": "0F766E",     # 深绿底白色加粗字
    "P1": "FDE68A", "S2": "FDE68A",     # 浅黄底
    "I1": "164E63", "O1": "164E63",     # 深蓝底
}
DATA_HEADER_WHITE_BOLD = ("A2", "O2")

# 支票明细：数据有效性
DATA_DV_COUNT = 4
DATA_DV_PRICE_LIST = "2.05,2.84,2.28"

# 单元格取值内容（字面量；比对时去空格/换行）
TICKET_FIXED_TEXTS = {
    "I2": "砾恒综合后勤服务部",
    "I3": "6274 1802 0065 3918",
    "I4": "华城商业银行云湖支行",
    "B7": "无承付",
    "H8": "7620 4381 0296 1150  汀序园区综合服务费",
}
DATA_TITLE_TEXT = "月公共能源服务结算明细"       # 支票明细!J1
DATA_DATE_TEXT = "2026-04-18"                    # 支票明细!O1
DATA_HEADER_LABELS = (
    "序号", "经办人", "缴费单位", "结算账号", "开户机构", "票据编号",
    "结算金额", "单价", "使用量", "票据金额", "附件", "票据号", "结算日期",
)
FORM_NOTE_TEXT = (
    "使用说明：在浅蓝色区域输入用户编号，并在淡橙色区域填写相应内容；"
    "“单价”用鼠标点选即可。输入完毕后请在打印页中打印输出。"
)
# 划拨进账单中"取值随 支票明细!S1 联动"的取值单元格
TICKET_LINKED_CELLS = (
    "B2", "B3", "B4", "B5", "F7", "B8", "B9", "F1", "H1",
    "J6", "K6", "L6", "M6", "N6", "O6", "P6", "Q6", "R6", "S6", "T6",
)
TICKET_SEQ_CELL = "W2"                            # =支票明细!$S$1
TICKET_CONTENT_BOUND = "A1:W9"                    # 票面之外不应新增内容
DEMO_TEXT_RE = re.compile(r"(测试|示例|样例|演示|demo|sample|test)", re.I)

# 宏名黑名单：仍指向输入文件里那个不存在的宏、或明显不可读的自动命名
BAD_MACRO_NAME_RE = re.compile(r"^(按钮\d+_单击|Button\d+_Click|Macro\d+|宏\d+)$")

# ============================================================
# 维度二各评分项的满分（max_delta）——按 §2.2 约定：
#   加分项：max_delta = +N（如 +5、+3、+1），实际 delta ∈ {0, +N}
#   扣分项：max_delta = -N，命中扣分时 delta = max_delta，未命中时 delta = 0
#   hit = (delta == max_delta)；未命中扣分项的 delta 为 0。
# 顺序与细则一致：先 13 条加分项（合计 35 分），再 23 条扣分项。
# ============================================================
RULE_MAX_DELTA = {
    # ---- 加分项：交付形态与宏可运行性 ----
    "交付为启用宏工作簿 扩展名与内部内容类型一致 打开无修复提示": 1,
    "含可正常加载的 VBA 工程(xl/vbaProject.bin) 宏代码完整可读可修改": 5,
    "宏以 VBA 过程实现 而非 Excel 4.0 宏工作表(XLM 宏表)": 3,
    # ---- 加分项：按钮控件 ----
    "支票明细 P2:R2 存在标题含打印/进帐单的按钮控件 且为新增或失效宏指派被修复": 3,
    "按钮已指派单击宏 且指派目标是工作簿内确实存在的宏过程": 3,
    # ---- 加分项：定位目标工作表 ----
    "宏按工作表名称划拨进账单定位 并在预览前激活/选中": 3,
    "目标工作表隐藏时先取消隐藏再预览 预览结束后恢复隐藏状态": 1,
    # ---- 加分项：打印预览与范围 ----
    "点击按钮进入打印预览界面(PrintPreview) 而非直接送往打印机": 5,
    "预览范围限定为票面区域(自 A1 起至 U8/W9 一带)": 3,
    "由宏在预览前显式设置打印区域 或把定义名称校正为红框范围": 3,
    "预览页面为 A4 纸 票面在一页内完整显示": 1,
    # ---- 加分项：健壮性与收尾 ----
    "宏内含错误处理 不弹出 VBA 运行时错误中断": 3,
    "宏过程为无参数的公共 Sub 命名可读便于再次指派": 1,
    # ---- 扣分项：按钮与控件对象 ----
    "按钮标题文字保持 打印+进帐单（帐字写法不变）": -1,
    "按钮文字格式保持 宋体/10磅/加粗/水平垂直居中 且无填充色与边框": -1,
    "按钮打印对象属性保持不随工作表打印": -1,
    "按钮未被移动或放大遮挡 P1/Q1/R1/S1 表头": -1,
    # ---- 扣分项：工作簿结构 ----
    "工作簿未被插入额外工作表（仍为 3 张）": -3,
    "未新增指向宏表单元格的隐藏名称 且 填写!Print_Area=$A$1:$U$14 未删改": -1,
    "工作表名称与先后顺序及代码名 Sheet1/Sheet2/Sheet3 未被改动": -1,
    # ---- 扣分项：单元格数据与公式 ----
    "支票明细内容未被改动（J1/O1/A2:O2 表头/P3 Q3 统计结果）": -3,
    "划拨进账单取值内容未被改动（B2:B5/J6:T6/F7/B8 B9/F1 H1）": -3,
    "划拨进账单 W2 仍取 支票明细!$S$1 的打印序号值": -1,
    "划拨进账单固定文本未被改动（I2/I3/I4/B7/H8）": -1,
    "填写工作表内容未被改动（A18 使用说明/E3 缴费单位）": -3,
    "未为演示或测试在任一工作表写入新内容": -1,
    # ---- 扣分项：格式与版式 ----
    "划拨进账单底色与字体未被改动": -3,
    "划拨进账单 17 处合并单元格未被拆分或改变": -1,
    "划拨进账单行高未被改动": -1,
    "划拨进账单列宽未被改动": -1,
    "划拨进账单页边距未被改动": -1,
    "支票明细行高列宽与明细区表头填充色未被改动": -1,
    # ---- 扣分项：规则与视图状态 ----
    "支票明细数据有效性规则仍为 4 条（含单价下拉列表）": -1,
    "支票明细忽略错误—列表数据有效性设置仍在": -1,
    "支票明细冻结窗格仍为冻结首列与前两行": -1,
    "保存时的活动工作表仍为支票明细": -1,
}
MAX_SCORE_TOTAL = sum(v for v in RULE_MAX_DELTA.values() if v > 0)
assert MAX_SCORE_TOTAL == 35

# ============================================================
# 通用辅助
# ============================================================

def _norm_text(value):
    """把单元格/控件文字规整为可比对形式：去掉所有空白（含不换行空格、全角
    空格、换行），其余字符原样保留（"帐"/"账"等写法差异因此仍可分辨）。"""
    if value is None:
        return ""
    if hasattr(value, "text"):          # openpyxl 的 ArrayFormula
        value = value.text
    if not isinstance(value, str):
        value = str(value)
    return re.sub(r"[\s 　]+", "", value)


def _cell_formula_or_text(cell):
    """取单元格的公式文本（含数组公式）或字面值文本；空单元格返回 ""。"""
    v = cell.value
    if v is None:
        return ""
    if hasattr(v, "text"):              # ArrayFormula
        return str(v.text)
    return str(v)


def _local(tag):
    return tag.rsplit("}", 1)[-1]


def _read_part(zf, name):
    """读取包内某个部件的文本；不存在或读取失败返回 ""。"""
    try:
        return zf.read(name).decode("utf-8", errors="ignore")
    except Exception:
        return ""


def _ref_bounds(ref):
    """"A1:W9" / "A1" -> (min_col, min_row, max_col, max_row)（1 基）。"""
    parts = ref.replace("$", "").split(":")
    out = []
    for p in parts:
        m = re.match(r"([A-Za-z]+)(\d+)$", p)
        if not m:
            return None
        out.append((column_index_from_string(m.group(1).upper()), int(m.group(2))))
    if len(out) == 1:
        out.append(out[0])
    return (out[0][0], out[0][1], out[1][0], out[1][1])

def _ranges_intersect(ref_a, ref_b):
    """两个 A1 区域是否相交（任一为 None 视为不相交）。"""
    a = _ref_bounds(ref_a) if ref_a else None
    b = _ref_bounds(ref_b) if ref_b else None
    if not a or not b:
        return False
    return not (a[2] < b[0] or b[2] < a[0] or a[3] < b[1] or b[3] < a[1])


def _bounds_to_ref(min_col, min_row, max_col, max_row):
    return "%s%d:%s%d" % (get_column_letter(min_col), min_row,
                          get_column_letter(max_col), max_row)


def _norm_area(text):
    """把 "'划拨进账单'!$A$1:$W$9" / "划拨进账单!A1:W9" 规整为 "$A$1:$W$9"。"""
    if not text:
        return None
    body = str(text).split("!")[-1].replace("'", "").strip()
    bounds = _ref_bounds(body)
    if not bounds:
        return None
    return "$%s$%d:$%s$%d" % (get_column_letter(bounds[0]), bounds[1],
                              get_column_letter(bounds[2]), bounds[3])


def _close(a, b, tol):
    return a is not None and b is not None and abs(float(a) - float(b)) <= tol


def _rgb(color):
    """openpyxl 颜色对象 -> 6 位大写 RGB（取不到返回 None）。"""
    if color is None:
        return None
    try:
        v = color.rgb
    except Exception:
        return None
    if not isinstance(v, str):
        return None
    v = v.upper()
    return v[-6:] if len(v) >= 6 else None


def _fill_rgb(cell):
    try:
        if not cell.fill or not cell.fill.patternType:
            return None
        return _rgb(cell.fill.fgColor)
    except Exception:
        return None

# ============================================================
# 包内结构解析：工作表 XML、VML 控件、drawing、VBA 工程
# ============================================================

def _sheet_col_widths(sheet_xml):
    """从工作表 XML 的 <cols><col min max width/> 展开出"列字母 -> 存储列宽"。
    openpyxl 的 column_dimensions 只记录成组区间的首列，故这里直接读原始 XML。"""
    widths = {}
    for m in re.finditer(r"<col\b[^>]*/?>", sheet_xml):
        tag = m.group(0)
        mn = re.search(r'\bmin="(\d+)"', tag)
        mx = re.search(r'\bmax="(\d+)"', tag)
        w = re.search(r'\bwidth="([\d.]+)"', tag)
        if not (mn and w):
            continue
        lo = int(mn.group(1))
        hi = int(mx.group(1)) if mx else lo
        hi = min(hi, 200)               # 忽略"整表默认宽"式的超大区间尾部
        for idx in range(lo, hi + 1):
            widths.setdefault(get_column_letter(idx), float(w.group(1)))
    return widths


def _sheet_row_heights(sheet_xml):
    """从工作表 XML 的 <row r ht/> 读出"行号 -> 行高"（未显式设置的行不收录）。"""
    heights = {}
    for m in re.finditer(r"<row\b[^>]*>", sheet_xml):
        tag = m.group(0)
        r = re.search(r'\br="(\d+)"', tag)
        h = re.search(r'\bht="([\d.]+)"', tag)
        if r and h:
            heights[int(r.group(1))] = float(h.group(1))
    return heights


def _sheet_pane(sheet_xml):
    """返回 (xSplit, ySplit, state)；无冻结窗格返回 (None, None, None)。"""
    m = re.search(r"<pane\b[^>]*>", sheet_xml)
    if not m:
        return (None, None, None)
    tag = m.group(0)
    xs = re.search(r'\bxSplit="([\d.]+)"', tag)
    ys = re.search(r'\bySplit="([\d.]+)"', tag)
    st = re.search(r'\bstate="([^"]+)"', tag)
    return (float(xs.group(1)) if xs else 0.0,
            float(ys.group(1)) if ys else 0.0,
            st.group(1) if st else None)

def _vml_buttons(vml_text):
    """解析 vmlDrawing 里的按钮控件。返回 list[dict]：
       ref/anchor（锚定区域）、caption（去空白标题）、macro（指派的宏）、
       print_object、font_faces/font_sizes/bold/halign/valign、filled/stroked。"""
    buttons = []
    for m in re.finditer(r"<v:shape\b.*?</v:shape>", vml_text, re.S):
        shape = m.group(0)
        if "ObjectType=\"Button\"" not in shape:
            continue
        anchor = re.search(r"<x:Anchor>([^<]*)</x:Anchor>", shape)
        ref = None
        if anchor:
            nums = [int(x) for x in re.findall(r"-?\d+", anchor.group(1))]
            if len(nums) >= 8:
                ref = _bounds_to_ref(nums[0] + 1, nums[2] + 1, nums[4] + 1, nums[6] + 1)
        caption_parts = re.findall(r"<font\b[^>]*>(.*?)</font>", shape, re.S)
        caption = _norm_text(re.sub(r"<[^>]+>", "", "".join(caption_parts)))
        macro = re.search(r"<x:FmlaMacro>([^<]*)</x:FmlaMacro>", shape)
        po = re.search(r"<x:PrintObject>([^<]*)</x:PrintObject>", shape)
        style = re.search(r'style="([^"]*)"', shape)
        halign = re.search(r"<x:TextHAlign>([^<]*)</x:TextHAlign>", shape)
        valign = re.search(r"<x:TextVAlign>([^<]*)</x:TextVAlign>", shape)
        filled = re.search(r'\bfilled="([^"]*)"', shape)
        stroked = re.search(r'\bstroked="([^"]*)"', shape)
        buttons.append({
            "ref": ref,
            "anchor": anchor.group(1) if anchor else "",
            "caption": caption,
            "macro": macro.group(1).strip() if macro else "",
            "print_object": (po.group(1).strip() if po else ""),
            "font_faces": re.findall(r'<font\b[^>]*face="([^"]*)"', shape),
            "font_sizes": re.findall(r'<font\b[^>]*size="([^"]*)"', shape),
            "bold": bool(re.search(r"<b>", shape)),
            "halign": halign.group(1).strip() if halign else None,
            "valign": valign.group(1).strip() if valign else None,
            "filled": filled.group(1) if filled else None,
            "stroked": stroked.group(1) if stroked else None,
            "fillcolor": bool(re.search(r'\bfillcolor="', shape)),
            "strokecolor": bool(re.search(r'\bstrokecolor="', shape)),
            "style": style.group(1) if style else "",
        })
    return buttons

def _drawing_shapes(drawing_xml):
    """解析 drawing1.xml：按锚点逐个形状收集 name/macro/可见标题文字。
       按钮标题在 Excel 里由 VML 与 drawing 两处共同承载，两处都要一致。"""
    shapes = []
    if not drawing_xml.strip():
        return shapes
    try:
        root = ET.fromstring(drawing_xml)
    except Exception:
        return shapes
    for anchor in list(root):
        texts, name, macro = [], None, None
        for el in anchor.iter():
            if _local(el.tag) == "t" and el.text:
                texts.append(el.text)
            if name is None and el.get("name"):
                name = el.get("name")
            if macro is None and el.get("macro"):
                macro = el.get("macro")
        shapes.append({
            "name": name or "",
            "macro": (macro or "").strip(),
            "caption": _norm_text("".join(texts)),
            "has_text": bool(texts),
        })
    return shapes


def _extract_vba(file_path, parts):
    """抽取 VBA 工程源码与其中的 Sub/Function 名单。

    使用 pyOpenVBA 解析出可读源码（能正确还原中文过程名）。依赖缺失属于
    verifier 运行环境错误，必须显式报错，不能静默折算成“交付文件没有宏”。
    返回 (source, subs, extract_ok)；subs 为 [(模块名, 过程名, 是否Private, 是否无参)]。
    """
    source, extract_ok = "", False
    if VBA_PART not in parts:
        return ("", [], False)
    try:
        from pyopenvba import ExcelFile
    except ImportError as exc:
        raise RuntimeError("缺少 pyOpenVBA，无法解析 xl/vbaProject.bin") from exc
    try:
        chunks = []
        with ExcelFile(str(file_path)) as workbook:
            for mod, code in workbook.vba_modules().items():
                if code:
                    chunks.append("'#### MODULE %s\n%s" % (mod, code))
        source = "\n".join(chunks)
        extract_ok = bool(source)
    except Exception:
        source, extract_ok = "", False
    return (source, _list_vba_subs(source), extract_ok)

def _list_vba_subs(source):
    """列出 VBA 源码中的过程：[(module, name, is_private, no_arg)]。"""
    subs = []
    module = ""
    for line in (source or "").splitlines():
        s = line.strip()
        m = re.match(r"'#### MODULE (.+)$", s)
        if m:
            module = m.group(1).strip()
            continue
        m = re.match(r"Attribute VB_Name = \"(.+)\"$", s)
        if m:
            module = m.group(1).strip()
            continue
        m = re.match(r"^(Public\s+|Private\s+|Friend\s+)?(Static\s+)?"
                     r"(Sub|Function)\s+([^\s(]+)\s*\(([^)]*)\)", s, re.I)
        if m:
            scope = (m.group(1) or "").strip().lower()
            subs.append({
                "module": re.sub(r"\.(cls|bas|frm)$", "", module, flags=re.I),
                "name": m.group(4).strip(),
                "kind": m.group(3).lower(),
                "is_private": scope == "private",
                "no_arg": not m.group(5).strip(),
            })
    return subs


def _macro_target_name(macro):
    """把按钮的宏指派（如 "[0]!Sheet1.打印预览划拨进账单"）拆成 (模块, 过程名)。"""
    if not macro:
        return (None, None)
    body = re.sub(r"^\[\d+\]!", "", str(macro).strip())
    body = body.replace("'", "").split("!")[-1]
    if "." in body:
        mod, _, name = body.rpartition(".")
        return (mod.strip() or None, name.strip())
    return (None, body.strip())


def _resolve_vba_sub(macro, subs):
    """按钮指派的宏是否对应工作簿内确实存在的 VBA 过程；命中返回该过程。"""
    mod, name = _macro_target_name(macro)
    if not name:
        return None
    for s in subs:
        if s["name"] != name:
            continue
        if mod and s["module"] and mod.lower() != s["module"].lower():
            continue
        return s
    return None

def _sub_body(source, name):
    """取出指定过程的代码体；取不到时返回整份源码（不影响"至少写了"的判定）。"""
    if not source or not name:
        return source or ""
    pat = re.compile(r"^\s*(?:Public\s+|Private\s+|Friend\s+)?(?:Static\s+)?"
                     r"(Sub|Function)\s+%s\s*\(" % re.escape(name),
                     re.I | re.M)
    m = pat.search(source)
    if not m:
        return source
    tail = source[m.start():]
    end = re.search(r"^\s*End\s+(Sub|Function)\b", tail, re.I | re.M)
    return tail[:end.end()] if end else tail


def _strip_vba_comments(code):
    """去掉 VBA 行注释，避免注释里的关键词被误判为代码行为。"""
    out = []
    for line in (code or "").splitlines():
        in_str = False
        cut = len(line)
        for i, ch in enumerate(line):
            if ch == '"':
                in_str = not in_str
            elif ch == "'" and not in_str:
                cut = i
                break
        out.append(line[:cut])
    return "\n".join(out)


def _macro_behavior(body):
    """从宏代码体里读出与打印预览相关的行为特征。"""
    code = _strip_vba_comments(body)
    def pos(pattern):
        m = re.search(pattern, code, re.I)
        return m.start() if m else -1
    preview_pos = pos(r"\.PrintPreview\b|\bPrintPreview\b")
    printout = re.search(r"\.PrintOut\b", code, re.I)
    printout_preview = re.search(r"\.PrintOut[^\n]*Preview\s*:=\s*True", code, re.I)
    area = re.search(r"PrintArea\s*=\s*\"([^\"]+)\"", code, re.I)
    visible_sets = [m.start() for m in re.finditer(r"\.Visible\s*=", code, re.I)]
    return {
        "code": code,
        "by_name": bool(re.search(r"(Sheets|Worksheets)\s*\(\s*\"%s\"\s*\)" % TICKET_SHEET, code)),
        "by_index": bool(re.search(r"(Sheets|Worksheets)\s*\(\s*3\s*\)", code)) or
                    bool(re.search(r"\bSheet3\b", code)),
        "preview_pos": preview_pos,
        "select_pos": max(pos(r"\.Select\b"), pos(r"\.Activate\b")),
        "printout": bool(printout) and not bool(printout_preview),
        "print_area": _norm_area(area.group(1)) if area else None,
        "visible_sets": visible_sets,
        "on_error": bool(re.search(r"On\s+Error\s+(GoTo|Resume\s+Next)", code, re.I)),
        "msgbox": bool(re.search(r"\bMsgBox\b", code, re.I)),
        "exit_sub": bool(re.search(r"\bExit\s+Sub\b", code, re.I)),
    }

def _parse_rels(zf, rels_path):
    """读取 .rels，返回 {rId: 目标部件的包内绝对路径}。"""
    out = {}
    text = _read_part(zf, rels_path)
    if not text:
        return out
    base = posix_dirname(posix_dirname(rels_path))
    try:
        root = ET.fromstring(text)
    except Exception:
        return out
    for rel in root:
        rid = rel.get("Id")
        target = rel.get("Target") or ""
        if not rid:
            continue
        if target.startswith("/"):
            out[rid] = target.lstrip("/")
        else:
            out[rid] = posix_normpath(base + "/" + target if base else target)
    return out


def posix_dirname(path):
    return path.rsplit("/", 1)[0] if "/" in path else ""


def posix_normpath(path):
    stack = []
    for seg in path.split("/"):
        if seg in ("", "."):
            continue
        if seg == "..":
            if stack:
                stack.pop()
            continue
        stack.append(seg)
    return "/".join(stack)


def _parse_package(file_path):
    """一次性解析包内所有需要的结构信息，供各评分项复用。"""
    f = {
        "path": file_path,
        "ext": os.path.splitext(file_path)[1].lower(),
        "parts": [],
        "workbook_ct": "",
        "sheets": [],            # [{name, state, rid, target, is_worksheet}]
        "defined_names": [],     # [{name, local, hidden, function, xlm, text}]
        "active_tab": 0,
        "sheet_xml": {},         # 工作表名 -> XML 文本
        "codenames": {},         # 工作表名 -> sheetPr/@codeName
        "tab_selected": {},      # 工作表名 -> bool
        "vml_buttons": [],
        "drawing_shapes": [],
        "xlm_sheets": [],
        "xlm_formulas": "",
        "vba_source": "",
        "vba_subs": [],
        "vba_extract_ok": False,
    }

    with zipfile.ZipFile(file_path) as zf:
        f["parts"] = zf.namelist()
        # ① [Content_Types].xml：工作簿主部件的内容类型（判定扩展名与内容是否一致）
        ct_text = _read_part(zf, "[Content_Types].xml")
        try:
            ct_root = ET.fromstring(ct_text) if ct_text else None
        except Exception:
            ct_root = None
        if ct_root is not None:
            for el in ct_root:
                if _local(el.tag) == "Override" and el.get("PartName") == "/xl/workbook.xml":
                    f["workbook_ct"] = el.get("ContentType") or ""
        # ② workbook.xml：工作表清单、定义名称、活动工作表
        wb_text = _read_part(zf, "xl/workbook.xml")
        rels = _parse_rels(zf, "xl/_rels/workbook.xml.rels")
        try:
            wb_root = ET.fromstring(wb_text) if wb_text else None
        except Exception:
            wb_root = None
        if wb_root is not None:
            for el in wb_root.iter():
                lname = _local(el.tag)
                if lname == "workbookView":
                    f["active_tab"] = int(el.get("activeTab") or 0)
                elif lname == "sheet":
                    rid = None
                    for k, v in el.attrib.items():
                        if _local(k) == "id":
                            rid = v
                    target = rels.get(rid or "", "")
                    f["sheets"].append({
                        "name": el.get("name") or "",
                        "state": el.get("state") or "visible",
                        "rid": rid,
                        "target": target,
                        "is_worksheet": "worksheets/" in target,
                    })
                elif lname == "definedName":
                    f["defined_names"].append({
                        "name": el.get("name") or "",
                        "local": el.get("localSheetId"),
                        "hidden": (el.get("hidden") == "1"),
                        "function": (el.get("function") == "1"),
                        "xlm": (el.get("xlm") == "1"),
                        "text": (el.text or "").strip(),
                    })
        # ③ 各工作表 XML、代码名、选中状态
        for sh in f["sheets"]:
            if not sh["target"]:
                continue
            xml = _read_part(zf, sh["target"])
            f["sheet_xml"][sh["name"]] = xml
            m = re.search(r'<sheetPr\b[^>]*codeName="([^"]*)"', xml)
            f["codenames"][sh["name"]] = m.group(1) if m else None
            f["tab_selected"][sh["name"]] = bool(re.search(r'<sheetView\b[^>]*tabSelected="1"', xml))
            if "macrosheets/" in sh["target"]:
                f["xlm_sheets"].append(sh["name"])
                f["xlm_formulas"] += " ".join(re.findall(r"<f>([^<]*)</f>", xml))

        # ④ 支票明细上的控件：VML（按钮本体/标题/宏指派）与 drawing（标题文字）
        data_target = ""
        for sh in f["sheets"]:
            if sh["name"] == DATA_SHEET:
                data_target = sh["target"]
        sheet_rels = _parse_rels(
            zf, posix_dirname(data_target) + "/_rels/" +
            data_target.rsplit("/", 1)[-1] + ".rels") if data_target else {}
        vml_parts = [t for t in sheet_rels.values() if t.lower().endswith(".vml")]
        drawing_parts = [t for t in sheet_rels.values() if "/drawings/drawing" in t]
        if not vml_parts:
            vml_parts = [p for p in f["parts"] if p.lower().endswith(".vml")]
        if not drawing_parts:
            drawing_parts = [p for p in f["parts"] if "/drawings/drawing" in p and p.endswith(".xml")]
        for p in vml_parts:
            f["vml_buttons"].extend(_vml_buttons(_read_part(zf, p)))
        for p in drawing_parts:
            f["drawing_shapes"].extend(_drawing_shapes(_read_part(zf, p)))
    # ⑤ VBA 工程
    f["vba_source"], f["vba_subs"], f["vba_extract_ok"] = _extract_vba(file_path, f["parts"])
    return f


def _pick_workbook(names, dir_path):
    """在一批文件名中按"能承载宏优先"的顺序挑一个；同优先级内按名称排序。"""
    buckets = {ext: [] for ext in WORKBOOK_EXTS}
    for name in names:
        if name.startswith("~$") or name.startswith("."):
            continue                    # Excel 临时锁文件 / 隐藏文件
        ext = os.path.splitext(name)[1].lower()
        if ext in buckets and os.path.isfile(os.path.join(dir_path, name)):
            buckets[ext].append(name)
    for ext in WORKBOOK_EXTS:
        if buckets[ext]:
            return os.path.join(dir_path, sorted(buckets[ext])[0])
    return None


def _find_workbook_in_dir(dir_path):
    """只在 dir_path 这一层目录内定位待评估的工作簿；找不到返回 None。

    选取策略：忽略 Excel 临时锁文件（~$ 前缀），优先取能承载宏的
    .xlsm/.xlsb/.xls；都没有时退而取 .xlsx（交由维度一判负）。
    不进入任何子目录，也不回退到其他目录。
    """
    if not dir_path or not os.path.isdir(dir_path):
        return None
    try:
        entries = sorted(os.listdir(dir_path))
    except OSError:
        return None
    return _pick_workbook(entries, dir_path)


def _locate_workbook(dir_path):
    """在 dir_path（恒为本脚本所在目录）这一层内定位工作簿。

    定位不到时抛出 FileNotFoundError，由 evaluate() 统一转成"未找到文件"的返回结构。
    """
    if not os.path.isdir(dir_path):
        raise FileNotFoundError("目录不存在或不是目录：%s" % dir_path)
    found = _find_workbook_in_dir(dir_path)
    if not found:
        raise FileNotFoundError("目录中未找到 .xlsm 文件：%s" % dir_path)
    return found

# ============================================================
# 维度1：可用与可修改性
# ============================================================

def _check_dimension_one(f, wb):
    """交付物须为可正常打开的启用宏工作簿，三张表的单元格可单独点选与编辑，
    宏代码可在 VBA 编辑器中查看与修改。返回 (ok, [问题描述...])。"""
    issues = []
    if f["ext"] not in MACRO_EXTS:
        issues.append("交付为 %s 格式，无法承载宏，按钮与宏全部失效" % f["ext"])
    if f["ext"] == ".xlsm" and f["workbook_ct"] and f["workbook_ct"] != WORKBOOK_MACRO_CT:
        issues.append("扩展名 .xlsm 与内部内容类型不一致（%s），Excel 打开会提示修复"
                      % f["workbook_ct"])
    names = [sh["name"] for sh in f["sheets"] if sh["is_worksheet"]]
    missing = [n for n in SHEET_NAMES if n not in names]
    if missing:
        issues.append("工作表缺失：%s" % "、".join(missing))
    for name in SHEET_NAMES:
        xml = f["sheet_xml"].get(name, "")
        if xml and re.search(r"<sheetProtection\b", xml):
            issues.append("%s 被保护，单元格无法单独点选与编辑" % name)
    if TICKET_SHEET in wb.sheetnames:
        ws = wb[TICKET_SHEET]
        bounds = _ref_bounds(TICKET_CONTENT_BOUND)
        filled = 0
        for row in ws.iter_rows(min_col=bounds[0], min_row=bounds[1],
                                max_col=bounds[2], max_row=bounds[3]):
            for cell in row:
                if _cell_formula_or_text(cell).strip():
                    filled += 1
        if filled < 10:
            issues.append("划拨进账单票面区域仅剩 %d 个非空单元格，内容疑似变为图片或"
                          "版面大面积丢失" % filled)
    return (not issues, issues)

# ============================================================
# 维度2：完成度评分
# ============================================================

def _score_bonus(f, wb, hits):
    """一、加分项（均来源于指令要求的宏按钮制作任务），合计 35 分。"""
    add = hits.append

    # ---- 交付形态与宏可运行性 ----
    ext_ok = f["ext"] in MACRO_EXTS
    ct_ok = (f["ext"] != ".xlsm") or (f["workbook_ct"] == WORKBOOK_MACRO_CT)
    add(("交付为启用宏工作簿 扩展名与内部内容类型一致 打开无修复提示",
         1 if (ext_ok and ct_ok) else 0,
         "扩展名=%s，workbook.xml 内容类型=%s" % (f["ext"], f["workbook_ct"] or "缺失")))

    vba_present = VBA_PART in f["parts"]
    vba_readable = bool(f["vba_subs"])
    add(("含可正常加载的 VBA 工程(xl/vbaProject.bin) 宏代码完整可读可修改",
         5 if (vba_present and vba_readable) else 0,
         "vbaProject.bin=%s，可读出过程 %d 个" % (vba_present, len(f["vba_subs"]))))

    # 按钮实际指派到的过程（是 VBA 过程才算真正能点得动）
    buttons = f["vml_buttons"]
    target_sub, target_button = None, None
    for b in buttons:
        sub = _resolve_vba_sub(b["macro"], f["vba_subs"])
        if sub is not None:
            target_sub, target_button = sub, b
            break
    xlm_present = bool(f["xlm_sheets"]) or any(
        p.startswith(MACROSHEET_PREFIX) for p in f["parts"])
    vba_impl = bool(target_sub) or (vba_readable and not xlm_present)
    add(("宏以 VBA 过程实现 而非 Excel 4.0 宏工作表(XLM 宏表)",
         3 if vba_impl else 0,
         "按钮指派的 VBA 过程=%s；XLM 宏表=%s"
         % ((target_sub or {}).get("name") or "无", f["xlm_sheets"] or "无")))

    # ---- 按钮控件 ----
    zone_buttons = [b for b in buttons if _ranges_intersect(b["ref"], BUTTON_ZONE)]
    def _caption_ok(text):
        return ("打印" in text) and (("进帐单" in text) or ("进账单" in text))
    ok_button = None
    for b in zone_buttons:
        if _caption_ok(b["caption"]) and _resolve_vba_sub(b["macro"], f["vba_subs"]):
            ok_button = b
            break
    add(("支票明细 P2:R2 存在标题含打印/进帐单的按钮控件 且为新增或失效宏指派被修复",
         3 if ok_button else 0,
         "P2:R2 按钮 %d 个，标题=%s，宏指派=%s"
         % (len(zone_buttons),
            [b["caption"] for b in zone_buttons] or "无",
            [b["macro"] for b in zone_buttons] or "无")))

    add(("按钮已指派单击宏 且指派目标是工作簿内确实存在的宏过程",
         3 if target_sub else 0,
         "指派目标=%s，解析到的过程=%s"
         % ((target_button or {}).get("macro") or
            ([b["macro"] for b in buttons] or "无"),
            (target_sub or {}).get("name") or "工作簿内不存在（点击会提示无法找到宏）")))

    # ---- 宏功能：以按钮真正指派到的过程为准逐条判定 ----
    body = _sub_body(f["vba_source"], (target_sub or {}).get("name")) if target_sub else ""
    beh = _macro_behavior(body) if body else None
    ws3 = wb[TICKET_SHEET] if TICKET_SHEET in wb.sheetnames else None

    add(("宏按工作表名称划拨进账单定位 并在预览前激活/选中",
         3 if (beh and beh["by_name"] and beh["select_pos"] >= 0
               and beh["preview_pos"] > beh["select_pos"]) else 0,
         "按名称定位=%s，激活/选中位置=%s，预览位置=%s"
         % (beh["by_name"], beh["select_pos"], beh["preview_pos"]) if beh else "无可运行的 VBA 宏"))

    restore_ok = False
    if beh and beh["visible_sets"] and beh["preview_pos"] >= 0:
        before = [p for p in beh["visible_sets"] if p < beh["preview_pos"]]
        after = [p for p in beh["visible_sets"] if p > beh["preview_pos"]]
        restore_ok = bool(before and after)
    add(("目标工作表隐藏时先取消隐藏再预览 预览结束后恢复隐藏状态",
         1 if restore_ok else 0,
         "Visible 赋值位置=%s，预览位置=%s"
         % (beh["visible_sets"], beh["preview_pos"]) if beh else "无可运行的 VBA 宏"))

    preview_ok = bool(beh and beh["preview_pos"] >= 0 and not beh["printout"])
    add(("点击按钮进入打印预览界面(PrintPreview) 而非直接送往打印机",
         5 if preview_ok else 0,
         "PrintPreview=%s，直接 PrintOut=%s"
         % (beh["preview_pos"] >= 0, beh["printout"]) if beh else "无可运行的 VBA 宏"))

    # 预览范围：宏内显式设置优先，其次工作表既有打印区域
    sheet_area = _norm_area(ws3.print_area) if (ws3 and ws3.print_area) else None
    macro_area = beh["print_area"] if beh else None
    eff_area = macro_area or sheet_area
    area_bounds = _ref_bounds(eff_area) if eff_area else None
    area_ok = bool(preview_ok and area_bounds and area_bounds[0] == 1 and area_bounds[1] == 1
                   and 21 <= area_bounds[2] <= 23 and 8 <= area_bounds[3] <= 9)
    add(("预览范围限定为票面区域(自 A1 起至 U8/W9 一带)",
         3 if area_ok else 0,
         "生效打印区域=%s（宏内设置=%s，工作表设置=%s）"
         % (eff_area or "未设置", macro_area or "无", sheet_area or "无")))

    corrected = bool(sheet_area and sheet_area in TICKET_PRINT_AREAS
                     and sheet_area != BASELINE_TICKET_PRINT_AREA)
    set_by_macro = bool(macro_area and macro_area in TICKET_PRINT_AREAS)
    add(("由宏在预览前显式设置打印区域 或把定义名称校正为红框范围",
         3 if (preview_ok and (set_by_macro or corrected)) else 0,
         "宏内设置=%s；定义名称=%s（输入文件既有 %s，仅沿用不得分）"
         % (macro_area or "无", sheet_area or "无", BASELINE_TICKET_PRINT_AREA)))

    # A4 一页：优先看"缩放到一页宽"设置，其次按票面尺寸与可打印区域估算
    one_page, page_detail = False, "无可运行的 VBA 宏"
    if ws3 is not None:
        ps = ws3.page_setup
        paper_ok = (ps.paperSize in (None, 9, "9"))
        fit_to_page = bool(ws3.sheet_properties.pageSetUpPr
                           and ws3.sheet_properties.pageSetUpPr.fitToPage)
        fit_w = int(ps.fitToWidth) if ps.fitToWidth not in (None, "") else 1
        fit_h = int(ps.fitToHeight) if ps.fitToHeight not in (None, "") else 1
        fit_ok = fit_to_page and fit_w <= 1 and fit_h <= 1
        est_ok = False
        if area_bounds:
            widths = _sheet_col_widths(f["sheet_xml"].get(TICKET_SHEET, ""))
            heights = _sheet_row_heights(f["sheet_xml"].get(TICKET_SHEET, ""))
            # 列宽单位 -> 像素（本簿正文字体宋体 12，MDW=8px）-> 英寸
            units = sum(widths.get(get_column_letter(c), 9.0)
                        for c in range(area_bounds[0], area_bounds[2] + 1))
            width_in = units * 8.0 / 96.0
            height_pt = sum(heights.get(r, 14.25)
                            for r in range(area_bounds[1], area_bounds[3] + 1))
            landscape = (ps.orientation == "landscape")
            paper_w, paper_h = (11.69, 8.27) if landscape else (8.27, 11.69)
            usable_w = paper_w - float(ws3.page_margins.left or 0) - float(ws3.page_margins.right or 0)
            usable_h = (paper_h - float(ws3.page_margins.top or 0)
                        - float(ws3.page_margins.bottom or 0)) * 72.0
            est_ok = (width_in <= usable_w and height_pt <= usable_h)
            page_detail = ("纸张=%s，方向=%s，缩放到一页=%s，估算票面 %.2f 英寸宽 / %.1f 磅高"
                           % (ps.paperSize, ps.orientation, fit_ok, width_in, height_pt))
        one_page = bool(preview_ok and paper_ok and (fit_ok or est_ok))
    add(("预览页面为 A4 纸 票面在一页内完整显示", 1 if one_page else 0, page_detail))

    # ---- 健壮性与收尾 ----
    err_ok = bool(beh and beh["on_error"] and (beh["msgbox"] or beh["exit_sub"]))
    add(("宏内含错误处理 不弹出 VBA 运行时错误中断",
         3 if err_ok else 0,
         "On Error=%s，MsgBox=%s，Exit Sub=%s"
         % (beh["on_error"], beh["msgbox"], beh["exit_sub"]) if beh else "无可运行的 VBA 宏"))

    name_ok = False
    if target_sub:
        name_ok = (target_sub["kind"] == "sub" and target_sub["no_arg"]
                   and not target_sub["is_private"]
                   and not BAD_MACRO_NAME_RE.match(target_sub["name"] or ""))
    add(("宏过程为无参数的公共 Sub 命名可读便于再次指派",
         1 if name_ok else 0,
         "过程=%s" % (target_sub if target_sub else "无可运行的 VBA 宏")))

def _score_penalty(f, wb, hits):
    """二、扣分项（均为指令未要求、却被改动而破坏交付件的多余修改）。"""
    add = hits.append
    buttons = f["vml_buttons"]
    draw_buttons = [s for s in f["drawing_shapes"] if s["has_text"]]

    # ---- 按钮与控件对象 ----
    captions = [b["caption"] for b in buttons] + [s["caption"] for s in draw_buttons]
    caption_bad = [c for c in captions if c != BUTTON_CAPTION]
    add(("按钮标题文字保持 打印+进帐单（帐字写法不变）",
         -1 if caption_bad else 0,
         "按钮标题（去空格/换行后）=%s，应为 %s" % (captions or "无", BUTTON_CAPTION)))

    fmt_bad = []
    for b in buttons:
        if any(face != "宋体" for face in b["font_faces"]) or not b["font_faces"]:
            fmt_bad.append("字体=%s" % (b["font_faces"] or "缺失"))
        if any(sz not in ("200", "10") for sz in b["font_sizes"]) or not b["font_sizes"]:
            fmt_bad.append("字号=%s" % (b["font_sizes"] or "缺失"))
        if not b["bold"]:
            fmt_bad.append("未加粗")
        if b["halign"] != "Center" or b["valign"] != "Center":
            fmt_bad.append("对齐=%s/%s" % (b["halign"], b["valign"]))
        if b["fillcolor"] or b["strokecolor"] or b["filled"] == "t" or b["stroked"] == "t":
            fmt_bad.append("被加上填充色或边框")
    add(("按钮文字格式保持 宋体/10磅/加粗/水平垂直居中 且无填充色与边框",
         -1 if fmt_bad else 0, "；".join(fmt_bad) if fmt_bad else "宋体/10磅/加粗/居中，无填充无边框"))

    po_bad = [b["print_object"] or "缺失" for b in buttons
              if (b["print_object"] or "").lower() != "false"]
    add(("按钮打印对象属性保持不随工作表打印",
         -1 if po_bad else 0,
         "PrintObject=%s" % (po_bad if po_bad else "False（不随工作表打印）")))

    cover = [b["ref"] for b in buttons if _ranges_intersect(b["ref"], HEADER_GUARD_ZONE)]
    add(("按钮未被移动或放大遮挡 P1/Q1/R1/S1 表头",
         -1 if cover else 0,
         "按钮锚定区域=%s" % ([b["ref"] for b in buttons] or "无")))

    # ---- 工作簿结构 ----
    sheet_names = [sh["name"] for sh in f["sheets"]]
    add(("工作簿未被插入额外工作表（仍为 3 张）",
         -3 if len(sheet_names) > 3 else 0,
         "工作表清单=%s" % [(sh["name"], sh["state"]) for sh in f["sheets"]]))

    xlm_names = [dn["name"] for dn in f["defined_names"]
                 if dn["hidden"] and (dn["function"] or dn["xlm"])]
    form_area = None
    if FORM_SHEET in wb.sheetnames and wb[FORM_SHEET].print_area:
        form_area = _norm_area(wb[FORM_SHEET].print_area)
    name_bad = bool(xlm_names) or (form_area != FORM_PRINT_AREA)
    add(("未新增指向宏表单元格的隐藏名称 且 填写!Print_Area=$A$1:$U$14 未删改",
         -1 if name_bad else 0,
         "隐藏宏名称=%s；填写!Print_Area=%s" % (xlm_names or "无", form_area or "缺失")))

    order_ok = (tuple(sheet_names[:3]) == SHEET_NAMES)
    code_now = tuple(f["codenames"].get(n) for n in SHEET_NAMES)
    code_ok = (code_now == SHEET_CODENAMES)
    add(("工作表名称与先后顺序及代码名 Sheet1/Sheet2/Sheet3 未被改动",
         -1 if not (order_ok and code_ok) else 0,
         "前三张表=%s；代码名=%s" % (sheet_names[:3], list(code_now))))

    # ---- 单元格数据与公式 ----
    ws1 = wb[DATA_SHEET] if DATA_SHEET in wb.sheetnames else None
    ws2 = wb[FORM_SHEET] if FORM_SHEET in wb.sheetnames else None
    ws3 = wb[TICKET_SHEET] if TICKET_SHEET in wb.sheetnames else None

    data_bad = []
    if ws1 is None:
        data_bad.append("支票明细缺失")
    else:
        if _norm_text(ws1["J1"].value) != DATA_TITLE_TEXT:
            data_bad.append("J1=%r" % ws1["J1"].value)
        o1 = ws1["O1"].value
        o1_text = o1.strftime("%Y-%m-%d") if hasattr(o1, "strftime") else _norm_text(o1)
        if DATA_DATE_TEXT.replace("-", "") not in o1_text.replace("-", "").replace("/", ""):
            data_bad.append("O1=%r" % o1)
        row2 = {_norm_text(ws1.cell(row=2, column=c).value) for c in range(1, 16)}
        miss = [lab for lab in DATA_HEADER_LABELS if lab not in row2]
        if miss:
            data_bad.append("A2:O2 缺少表头 %s" % miss)
        for ref in ("P3", "Q3"):
            if not _cell_formula_or_text(ws1[ref]).strip():
                data_bad.append("%s 统计结果为空" % ref)
    add(("支票明细内容未被改动（J1/O1/A2:O2 表头/P3 Q3 统计结果）",
         -3 if data_bad else 0,
         "；".join(data_bad) if data_bad else "标题/结算日期/表头/张数统计均一致"))

    # 划拨进账单取值：公式写法可不同，但必须仍从 支票明细 取到内容（不得为空/断链）
    ticket_bad = []
    if ws3 is None:
        ticket_bad.append("划拨进账单缺失")
    else:
        for ref in TICKET_LINKED_CELLS:
            text = _cell_formula_or_text(ws3[ref])
            if not text.strip():
                ticket_bad.append("%s 为空" % ref)
            elif text.startswith("=") and DATA_SHEET not in text:
                ticket_bad.append("%s 不再引用支票明细" % ref)
    add(("划拨进账单取值内容未被改动（B2:B5/J6:T6/F7/B8 B9/F1 H1）",
         -3 if ticket_bad else 0,
         "；".join(ticket_bad) if ticket_bad else "缴费单位/结算账号/开户机构/金额大写/"
         "逐位小写/票据号/用途/核验方式/结算年月日均仍从支票明细取值"))

    seq_text = _cell_formula_or_text(ws3["W2"]) if ws3 is not None else ""
    seq_ok = bool(re.search(r"%s!?\$?S\$?1\b" % DATA_SHEET, seq_text))
    add(("划拨进账单 W2 仍取 支票明细!$S$1 的打印序号值",
         -1 if not seq_ok else 0, "W2=%s" % (seq_text or "空")))

    fixed_bad = []
    if ws3 is not None:
        for ref, expect in TICKET_FIXED_TEXTS.items():
            got = _norm_text(ws3[ref].value)
            if got != _norm_text(expect):
                fixed_bad.append("%s=%r" % (ref, ws3[ref].value))
    add(("划拨进账单固定文本未被改动（I2/I3/I4/B7/H8）",
         -1 if fixed_bad else 0,
         "；".join(fixed_bad) if fixed_bad else "收款方信息/无承付/用途行文本一致"))

    form_bad = []
    if ws2 is None:
        form_bad.append("填写工作表缺失")
    else:
        if _norm_text(ws2["A18"].value) != _norm_text(FORM_NOTE_TEXT):
            form_bad.append("A18=%r" % ws2["A18"].value)
        e3 = _cell_formula_or_text(ws2["E3"])
        if not e3.strip() or (e3.startswith("=") and DATA_SHEET not in e3):
            form_bad.append("E3=%r" % e3)
    add(("填写工作表内容未被改动（A18 使用说明/E3 缴费单位）",
         -3 if form_bad else 0,
         "；".join(form_bad) if form_bad else "使用说明全文与 E3 取值一致"))

    # 为演示或测试写入的新内容：票面之外多出的内容、或含测试/示例字样的文字
    demo_bad = []
    if ws3 is not None:
        b = _ref_bounds(TICKET_CONTENT_BOUND)
        for row in ws3.iter_rows():
            for cell in row:
                text = _cell_formula_or_text(cell).strip()
                if not text:
                    continue
                if not (b[0] <= cell.column <= b[2] and b[1] <= cell.row <= b[3]):
                    demo_bad.append("划拨进账单 %s 票面之外新增内容" % cell.coordinate)
    for ws, label, limit_row in ((ws2, FORM_SHEET, None), (ws3, TICKET_SHEET, None),
                                 (ws1, DATA_SHEET, 2)):
        if ws is None:
            continue
        rows = ws.iter_rows(max_row=limit_row) if limit_row else ws.iter_rows()
        for row in rows:
            for cell in row:
                text = _cell_formula_or_text(cell)
                if text and DEMO_TEXT_RE.search(text):
                    demo_bad.append("%s %s=%r" % (label, cell.coordinate, text[:40]))
    add(("未为演示或测试在任一工作表写入新内容",
         -1 if demo_bad else 0,
         "；".join(demo_bad[:5]) if demo_bad else "三张表均未见新增说明文字或测试数据"))

    # ---- 格式与版式 ----
    style_bad = []
    if ws3 is not None:
        for ref, rgb in TICKET_FILLS.items():
            got = _fill_rgb(ws3[ref])
            if got != rgb:
                style_bad.append("%s 底色=%s（应 %s）" % (ref, got, rgb))
        for ref in TICKET_WHITE_BOLD:
            cell = ws3[ref]
            if not cell.font or not cell.font.bold or _rgb(cell.font.color) != "FFFFFF":
                style_bad.append("%s 非白色加粗字" % ref)
        if not (ws3["W2"].border and ws3["W2"].border.bottom
                and ws3["W2"].border.bottom.style):
            style_bad.append("W2 缺下框线")
        b = _ref_bounds(TICKET_CONTENT_BOUND)
        bad_fonts = set()
        for row in ws3.iter_rows(min_col=b[0], min_row=b[1], max_col=b[2], max_row=b[3]):
            for cell in row:
                name = cell.font.name if cell.font else None
                if name and name != "宋体":
                    bad_fonts.add(name)
        if bad_fonts:
            style_bad.append("正文字体出现 %s" % sorted(bad_fonts))
    add(("划拨进账单底色与字体未被改动",
         -3 if style_bad else 0,
         "；".join(style_bad[:6]) if style_bad else
         "浅黄/深绿白粗/浅蓝/W2 深绿带下框线、正文宋体均一致"))

    merges_now = sorted(str(r) for r in ws3.merged_cells.ranges) if ws3 is not None else []
    merges_ok = (merges_now == sorted(TICKET_MERGES))
    add(("划拨进账单 17 处合并单元格未被拆分或改变",
         -1 if not merges_ok else 0,
         "合并单元格 %d 处%s" % (len(merges_now),
                                "" if merges_ok else "：%s" % merges_now)))

    t_heights = _sheet_row_heights(f["sheet_xml"].get(TICKET_SHEET, ""))
    h_bad = ["第 %d 行=%s（应 %s）" % (r, t_heights.get(r), h)
             for r, h in sorted(TICKET_ROW_HEIGHTS.items())
             if not _close(t_heights.get(r), h, ROW_HEIGHT_TOL)]
    add(("划拨进账单行高未被改动",
         -1 if h_bad else 0,
         "；".join(h_bad) if h_bad else "第 1/2/4/5/6/7/8 行行高一致"))

    t_widths = _sheet_col_widths(f["sheet_xml"].get(TICKET_SHEET, ""))
    w_bad = ["%s 列=%s（应 %s）" % (c, t_widths.get(c), w)
             for c, w in sorted(TICKET_COL_WIDTHS.items())
             if not _close(t_widths.get(c), w, COL_WIDTH_TOL)]
    add(("划拨进账单列宽未被改动",
         -1 if w_bad else 0,
         "；".join(w_bad) if w_bad else "A/B/C/D/G/H/I/J/K/M/N/U 列宽一致"))

    m_bad = []
    if ws3 is not None:
        pm = ws3.page_margins
        for key, expect in TICKET_MARGINS.items():
            got = getattr(pm, key, None)
            if not _close(got, expect, MARGIN_TOL):
                m_bad.append("%s=%s（应 %s）" % (key, got, expect))
    add(("划拨进账单页边距未被改动",
         -1 if m_bad else 0,
         "；".join(m_bad) if m_bad else "左 0.52/右 0/上 0.079/下 0，页眉页脚 0"))

    # 支票明细：行高（第 1/2 行绝对值 + 明细行 24/汇总区 12 仍在）、列宽、表头填充色
    layout_bad = []
    d_heights = _sheet_row_heights(f["sheet_xml"].get(DATA_SHEET, ""))
    for r, h in sorted(DATA_ROW_HEIGHTS.items()):
        if not _close(d_heights.get(r), h, ROW_HEIGHT_TOL):
            layout_bad.append("第 %d 行=%s（应 %s）" % (r, d_heights.get(r), h))
    others = [v for k, v in d_heights.items() if k > 2]
    detail_n = sum(1 for v in others if _close(v, DATA_DETAIL_HEIGHT, ROW_HEIGHT_TOL))
    summary_n = sum(1 for v in others if _close(v, DATA_SUMMARY_HEIGHT, ROW_HEIGHT_TOL))
    if detail_n < 10 or summary_n < 10:
        layout_bad.append("明细行 24 磅 %d 行、汇总区 12 磅 %d 行" % (detail_n, summary_n))
    d_widths = _sheet_col_widths(f["sheet_xml"].get(DATA_SHEET, ""))
    for c, w in sorted(DATA_COL_WIDTHS.items(), key=lambda kv: column_index_from_string(kv[0])):
        if not _close(d_widths.get(c), w, COL_WIDTH_TOL):
            layout_bad.append("%s 列=%s（应 %s）" % (c, d_widths.get(c), w))
    if ws1 is not None:
        for ref, rgb in DATA_FILLS.items():
            got = _fill_rgb(ws1[ref])
            if got != rgb:
                layout_bad.append("%s 填充色=%s（应 %s）" % (ref, got, rgb))
        for ref in DATA_HEADER_WHITE_BOLD:
            cell = ws1[ref]
            if not cell.font or not cell.font.bold or _rgb(cell.font.color) != "FFFFFF":
                layout_bad.append("%s 非白色加粗字" % ref)
    add(("支票明细行高列宽与明细区表头填充色未被改动",
         -1 if layout_bad else 0,
         "；".join(layout_bad[:6]) if layout_bad else
         "第 1/2 行行高、A:U 列宽（容差 0.01 字符）与表头填充色均一致"))

    # ---- 规则与视图状态 ----
    dv_xml = f["sheet_xml"].get(DATA_SHEET, "")
    dv_count = len(re.findall(r"<dataValidation\b", dv_xml))
    price_ok = DATA_DV_PRICE_LIST in dv_xml.replace("&quot;", "\"")
    add(("支票明细数据有效性规则仍为 4 条（含单价下拉列表）",
         -1 if (dv_count != DATA_DV_COUNT or not price_ok) else 0,
         "数据有效性 %d 条，单价下拉列表 %s" % (dv_count, "在" if price_ok else "丢失")))

    ie_ok = bool(re.search(r"<ignoredError\b[^>]*listDataValidation=\"1\"", dv_xml))
    add(("支票明细忽略错误—列表数据有效性设置仍在",
         -1 if not ie_ok else 0,
         "ignoredError/listDataValidation=%s" % ("在" if ie_ok else "已被删除")))

    xs, ys, state = _sheet_pane(dv_xml)
    pane_ok = (_close(xs, 1, 0.001) and _close(ys, 2, 0.001) and state == "frozen")
    add(("支票明细冻结窗格仍为冻结首列与前两行",
         -1 if not pane_ok else 0, "pane xSplit=%s ySplit=%s state=%s" % (xs, ys, state)))

    active_name = f["sheets"][f["active_tab"]]["name"] if (
        f["sheets"] and 0 <= f["active_tab"] < len(f["sheets"])) else None
    active_ok = (f["active_tab"] == 0 and f["tab_selected"].get(DATA_SHEET, False))
    add(("保存时的活动工作表仍为支票明细",
         -1 if not active_ok else 0,
         "activeTab=%s（%s），支票明细 tabSelected=%s"
         % (f["active_tab"], active_name, f["tab_selected"].get(DATA_SHEET))))

# ============================================================
# 主流程
# ============================================================

def _evaluate_impl(file_path, result):
    """维度一 + 维度二的评估主体：把 hits 按 RULE_MAX_DELTA 的顺序反解为
    结构化 dim2_items 填入 result。"""
    f = _parse_package(file_path)
    wb = load_workbook(file_path, data_only=False, keep_vba=True)

    # ---------- 维度1：可用与可修改性 ----------
    ok, issues = _check_dimension_one(f, wb)
    result["dim1_pass"] = bool(ok)
    if not ok:
        # dim1_reason 统一带上「维度一未通过：…」前缀与计分后果，便于直接作为错误摘要展示
        result["dim1_reason"] = "维度一未通过：%s，完成度总分记 0" % "; ".join(issues)
        result["dim2_items"] = []
        result["total_score"] = 0
        return result

    # ---------- 维度2：完成度评分 ----------
    hits = []
    _score_bonus(f, wb, hits)
    _score_penalty(f, wb, hits)

    if len(hits) != len(RULE_MAX_DELTA):
        raise RuntimeError("dim2 项数不一致：hits=%d vs rules=%d"
                           % (len(hits), len(RULE_MAX_DELTA)))
    total = 0
    for label, delta, _detail in hits:  # _detail 仅保留各判定处的可读注释，不再写入返回结果
        max_delta = RULE_MAX_DELTA.get(label)
        if max_delta is None:
            raise RuntimeError("未登记的评分项：%s" % label)
        result["dim2_items"].append({
            "rule": label,
            "max_delta": max_delta,
            "delta": delta,
            "hit": (delta == max_delta),
        })
        total += delta
    result["total_score"] = total
    result["max_score"] = MAX_SCORE_TOTAL
    return result

def _is_damaged_package(exc):
    """判断异常是否源于交付文件自身结构损坏（缺部件、XML 不合法等）。"""
    text = str(exc)
    if isinstance(exc, KeyError) and (
        "There is no item named" in text or "no relationship of type" in text
    ):
        return True
    return type(exc).__name__ in (
        "XMLSyntaxError", "ParseError", "InvalidXmlError", "MemoryError"
    )


def evaluate(dir_path=None):
    """对外唯一入口：在本脚本所在目录内定位输出工作簿并给出评估结果。

    评估目录取 dir_path，只看该目录本层：不检索子目录、上级目录，
    也不回退到调用方的当前工作目录；dir_path 未提供时退回本脚本所在目录。
    """
    dir_path = dir_path if dir_path and os.path.isdir(dir_path) else SCRIPT_DIR
    result = {
        "id": SCRIPT_ID,
        "file_name": None,
        "status": "ok",
        "error": None,
        "dim1_pass": False,
        "dim1_reason": "",
        "dim2_items": [],
        "total_score": 0,
        "max_score": MAX_SCORE_TOTAL,
    }
    try:
        file_path = _locate_workbook(dir_path)
        result["file_name"] = os.path.basename(file_path)

        ext = os.path.splitext(file_path)[1].lower()
        if ext in (".xlsb", ".xls"):
            result["status"] = "error"
            result["error"] = ("本脚本依赖 openpyxl，无法解析 %s 格式；"
                              "请提供 .xlsm 输出文件" % ext)
            return result

        return _evaluate_impl(file_path, result)
    except FileNotFoundError as exc:
        # 目录不存在或目录内没有工作簿：不进入维度判定，各评分字段一并留空/归零
        result["file_name"] = ""
        result["status"] = "error"
        result["error"] = str(exc)
        result["dim1_pass"] = ""
        result["dim1_reason"] = ""
        result["dim2_items"] = []
        result["total_score"] = 0
        result["max_score"] = 0
        return result
    except zipfile.BadZipFile:
        result["dim1_pass"] = False
        result["dim1_reason"] = "维度一未通过：文件不是有效的 OOXML 压缩包，无法打开，完成度总分记 0"
        result["dim2_items"] = []
        result["total_score"] = 0
        return result
    except Exception as exc:  # noqa: BLE001
        if _is_damaged_package(exc):
            result["status"] = "ok"
            result["error"] = None
            result["dim1_pass"] = False
            result["dim1_reason"] = "维度一未通过：交付文件结构不完整，无法正常打开（%s: %s），完成度总分记 0" % (
                type(exc).__name__, exc)
            result["dim2_items"] = []
            result["total_score"] = 0
            return result
        result["status"] = "error"
        result["error"] = "脚本运行异常：%s: %s" % (type(exc).__name__, exc)
        result["dim1_pass"] = False
        result["dim1_reason"] = ("维度一未通过：文件无法打开或解析（%s: %s），完成度总分记 0"
                                % (type(exc).__name__, exc))
        result["dim2_items"] = []
        result["total_score"] = 0
        return result


if __name__ == "__main__":
    # 评估目录固定为本脚本所在目录，不接受命令行参数
    print(json.dumps(evaluate(), ensure_ascii=False, indent=2))
