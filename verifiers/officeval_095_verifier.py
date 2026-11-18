#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# pyright: reportUnknownMemberType=false, reportUnknownVariableType=false, reportUnknownArgumentType=false, reportUnknownParameterType=false, reportMissingParameterType=false, reportPrivateUsage=false, reportUnknownLambdaType=false, reportGeneralTypeIssues=false

import os
import re
import sys
import json
import math
import zipfile
import posixpath
import xml.etree.ElementTree as ET

import openpyxl
from openpyxl.utils import get_column_letter

# 说明：按接口统一约定，本模块不修改全局 sys.stdout，也不 print 主结果。
# 主结果统一由 evaluate(dir_path) -> dict 返回；入参为"脚本所在目录的路径"，
# 由脚本自己在该目录内定位并打开被评估的 .xlsx/.xlsm。本文件仅在 __main__ 下作为
# 本地调试入口，才把 dict 序列化为 JSON 打印。

SCRIPT_ID = "095"
# 维度二所有"加分项"max_delta 之和；也是维度一未通过时回填的 max_score
# （脚本没评上时 max_score 为 0）。
# 减分项 max_delta 为负值，不计入满分，只从 total_score 里扣。
MAX_SCORE_TOTAL = 70

# 维度二所有评分项按"结果追加顺序"的 max_delta 列表；末端把 results 里的
# "+N：..." / "-N：..." 字符串反解为结构化 dim2_items 时用。顺序须与
# _evaluate_impl 中 results.append(...) 的顺序严格一致；
# 共 42 项 = 24 个加分项（累加 = MAX_SCORE_TOTAL）+ 18 个减分项。
_MAX_DELTAS_ORDER = [
    5,   # 01 存在1个可编辑的环形图对象
    3,   # 02 只呈现第一题的5个选项
    5,   # 03 类别与数值均引用单元格且与B3:B7、D3:D7一致
    3,   # 04 置于问题右侧空白区（起始列E或其右）
    1,   # 05 图表上边缘位于第1至第8行之间
    1,   # 06 以单元格锚定方式定位
    5,   # 07 问题文字置于圆环正中空心区、水平居中
    3,   # 08 中心问题文字自动换行完整显示、未溢出扇区
    1,   # 09 圆环内径占外径不小于55%
    5,   # 10 5个扇区均为斜线/斜纹图案填充
    3,   # 11 5个扇区填充颜色各不相同
    3,   # 12 相邻扇区间白色描边3至5磅
    3,   # 13 每个扇区带柔和发光或投影
    5,   # 14 有且仅有1个扇区被向外拉出
    3,   # 15 被拉出扇区位于右上方、外移1/4至1/3半径
    1,   # 16 被拉出扇区与其余扇区样式一致
    5,   # 17 5个扇区均显示百分比数据标签且数值一致
    3,   # 18 百分比标签带白色底框、无边框线
    3,   # 19 百分比标签文字颜色随所属扇区
    1,   # 20 数据标签仅显示百分比
    5,   # 21 提供选项名称与扇区颜色的对应标识
    1,   # 22 图表区无边框线、白色/无填充、无网格线坐标轴
    1,   # 23 宽度≥8厘米、高度≥6厘米
    1,   # 24 图表已含缓存数据，打开即正常显示
    -3,  # 25 新增了额外工作表
    -3,  # 26 A1标题单元格格式被改
    -3,  # 27 A2:D2表头格式被改
    -3,  # 28 A列20个调查事项文字被增删改写或错行
    -3,  # 29 B列回答选项文字被增删改写或错行
    -3,  # 30 C列有效人数数值/字体/对齐被改
    -3,  # 31 D列比例数值或数字格式被改
    -1,  # 32 D列比例文字或C列人数文字颜色被改
    -1,  # 33 D列/C列/B列单元格底纹被改
    -1,  # 34 A列问题单元格交替底纹被改
    -1,  # 35 表格框线被删除或改为其他线型
    -3,  # 36 表格未完整占满第1至第95行
    -1,  # 37 四列列宽被改
    -1,  # 38 B列/A列文字对齐方式被改
    -1,  # 39 表格数据区被新增合并单元格
    -3,  # 40 工作簿中图表数量多于1个
    -3,  # 41 图表覆盖遮挡A至D四列表格内容
    -1,  # 42 出现额外的标注文字或说明行
]
assert len(_MAX_DELTAS_ORDER) == 42
assert sum(d for d in _MAX_DELTAS_ORDER if d > 0) == MAX_SCORE_TOTAL

# ---------------------------------------------------------------------------
# OOXML 命名空间
# ---------------------------------------------------------------------------
NS_C = "http://schemas.openxmlformats.org/drawingml/2006/chart"
NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main"
NS_XDR = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_PR = "http://schemas.openxmlformats.org/package/2006/relationships"


def _c(tag):
    return "{%s}%s" % (NS_C, tag)


def _a(tag):
    return "{%s}%s" % (NS_A, tag)


def _xdr(tag):
    return "{%s}%s" % (NS_XDR, tag)


EMU_PER_CM = 360000.0
EMU_PER_PT = 12700.0

# ---------------------------------------------------------------------------
# 被评估表格的固有内容与格式（表格本身不属于本次修改范围，全部按原样校验）
# ---------------------------------------------------------------------------
SHEET_NAME = "Sheet1"
TITLE_TEXT = "乡村公共服务调查汇总"
HEADER_TEXTS = ["调查事项", "回答选项", "有效人数", "比例"]

# 第一个调查事项（图表应呈现的那一题）
Q1_TEXT = "您对村内便民服务事项的知晓程度"
Q1_OPTIONS = ["十分清楚", "比较清楚", "有所了解", "了解较少", "从未关注"]
Q1_RATIO_STRS = ["11.89%", "23.43%", "31.82%", "24.48%", "8.39%"]
Q1_PCT_STRS = ["12%", "23%", "32%", "24%", "8%"]
SAMPLE_TOTAL = 286  # 有效样本量：各调查事项的比例 = 该行有效人数 / 286

# 20 个调查事项所在行（A 列文字所在行，同时是每个题块的首行）
Q_ROWS = [3, 8, 13, 16, 22, 27, 31, 36, 40, 45, 49, 54, 58, 63, 70, 73, 76, 80, 84, 89]
DATA_FIRST_ROW = 3
DATA_LAST_ROW = 95  # 93 个回答选项行，表格完整占满第 1 至第 95 行

QUESTION_TEXTS = [
    "您对村内便民服务事项的知晓程度",
    "您获取公共服务信息的主要方式（可多选）",
    "近一年您是否参加过村级公共活动",
    "未能持续参与公共活动的主要原因（可多选）",
    "您对不同片区公共服务便利度的感受",
    "您家门前公共区域的日常维护情况",
    "您对生活垃圾分类与清运服务的满意度",
    "您对垃圾收集点清运频率的感受",
    "您对村内公共卫生设施的使用体验",
    "村内雨污排水设施的日常运行情况",
    "村内排水与保洁仍需改进的问题（可多选）",
    "您认为村民参与公共事务的合适角色",
    "您对共同承担公共设施维护费用的态度",
    "您最希望优先完善的公共服务（可多选）",
    "您是否通过正式渠道提出过意见",
    "您所提意见的后续处理情况",
    "公共环境提升对家庭经营或就业的影响",
    "村内是否建立常态化议事与巡查机制",
    "您对村级公共服务整体情况的满意度",
    "当前公共服务推进中的主要难点（可多选）",
]
assert len(QUESTION_TEXTS) == len(Q_ROWS) == 20

# B 列 93 个回答选项（第 3 行至第 95 行，逐行对应）
OPTION_TEXTS = [
    "十分清楚", "比较清楚", "有所了解", "了解较少", "从未关注",
    "村务公示栏", "网格员通知", "社区联络群", "亲友转告", "线下咨询",
    "多次参加", "参加过一两次", "没有参加",
    "活动时间不合适", "信息获取较晚", "内容吸引力不足", "照护或务工冲突", "报名方式不便", "其他原因",
    "差异明显", "有一定差异", "差异较小", "基本一致", "暂不清楚",
    "主动维护", "偶尔维护", "较少维护", "不了解相关要求",
    "非常满意", "比较满意", "一般", "不太满意", "非常不满意",
    "非常及时", "基本及时", "偶有延迟", "经常延迟",
    "非常满意", "比较满意", "一般", "不太满意", "非常不满意",
    "运行稳定", "偶有故障", "长期异常", "尚未覆盖",
    "末端设施不足", "处理效果不稳", "覆盖范围有限", "日常维护不足", "其他问题",
    "共同决策者", "协作参与者", "信息反馈者", "暂不确定",
    "非常支持", "比较支持", "视情况而定", "不太支持", "完全不支持",
    "垃圾分类", "排水保洁", "公共卫生", "道路照明", "设施维护", "片区均衡", "其他服务",
    "多次提出", "偶尔提出", "从未提出",
    "已处理并反馈", "有反馈但未完成", "尚未获得回应",
    "带动明显", "略有帮助", "影响不明显", "暂不清楚",
    "机制健全并执行", "已有机制但不稳定", "尚未建立", "不了解",
    "非常满意", "比较满意", "一般", "不太满意", "非常不满意",
    "村民参与不足", "协同沟通不畅", "资源投入有限", "持续维护困难", "片区发展不均",
    "服务与产业衔接弱", "其他难点",
]
assert len(OPTION_TEXTS) == 93

# C 列 93 个有效人数（第 3 行至第 95 行，逐行对应）；D 列比例 = 该行人数 / 286
COUNTS = [
    34, 67, 91, 70, 24, 149, 132, 118, 76, 29, 38, 102, 146, 104, 91, 84, 129,
    72, 21, 52, 97, 83, 31, 23, 71, 109, 74, 32, 49, 126, 71, 29, 11, 58, 137,
    64, 27, 43, 121, 78, 31, 13, 91, 76, 34, 85, 117, 96, 138, 82, 27, 56, 151,
    48, 31, 33, 71, 89, 65, 28, 123, 142, 69, 96, 88, 109, 24, 26, 73, 187, 46,
    72, 168, 39, 82, 124, 41, 48, 79, 119, 40, 44, 113, 89, 29, 11, 156, 127,
    139, 148, 121, 104, 22,
]
assert len(COUNTS) == 93

# 固有格式常量
FONT_SONG = ("宋体", "SimSun")
FONT_TNR = ("Times New Roman",)
CLR_WHITE = "FFFFFF"
FILL_TITLE = "174A55"       # A1 深青底纹
FILL_HEADER = "2F6F73"      # A2:D2 表头底纹
FILL_C = "E8F1F7"           # C 列浅蓝底纹
FILL_D = "FFF1CC"           # D 列浅黄底纹
FILL_B_ALT = ("F3F8F7", "F5F8FB")   # B 列按题块交替的浅青 / 浅蓝底纹
FILL_A_ALT = ("D7ECE8", "E1EDF3")   # A 列问题单元格按题交替的底纹
CLR_C_FONT = "244B5A"       # C 列人数深蓝文字
CLR_D_FONT = "6A4B13"       # D 列比例深棕文字
NUMFMT_D = "0.00%"
EXPECT_WIDTHS = {"A": 34.125, "B": 8.375, "C": 7.625, "D": 6.875}

# 额外标注文字关键词（出现即判定新增了说明行）
ANNOT_KEYWORDS = ("图表说明", "数据来源", "占比统计", "数据说明", "图例说明", "图表标题")

# 维度一只接受的工作簿扩展名
ACCEPT_EXTS = (".xlsx", ".xlsm")
REJECT_EXTS = (".xls", ".xlt", ".csv", ".et", ".ett", ".ods", ".numbers",
               ".pdf", ".png", ".jpg", ".jpeg", ".bmp", ".gif", ".doc",
               ".docx", ".txt", ".html", ".htm", ".xml", ".wps")


# ---------------------------------------------------------------------------
# 通用小工具
# ---------------------------------------------------------------------------
def _norm(text):
    """去掉全部空白字符，便于文字比对（换行、空格不视为差异）。"""
    if text is None:
        return ""
    return re.sub(r"\s+", "", str(text))


def _hex6(value):
    """把 openpyxl / OOXML 的颜色值统一成 6 位大写 RRGGBB。"""
    if value is None:
        return None
    text = str(value).strip().upper()
    if not re.fullmatch(r"[0-9A-F]{6,8}", text):
        return None
    return text[-6:]


def _bs(border_side):
    """空单元格的 border side 可能为 None，取 style 时先兜一层。"""
    return border_side.style if border_side is not None else None


def _cell_font_rgb(cell):
    color = cell.font.color if cell.font is not None else None
    if color is None:
        return None
    if getattr(color, "type", None) == "rgb":
        return _hex6(getattr(color, "rgb", None))
    return None


def _cell_fill_rgb(cell):
    fill = cell.fill
    if fill is None or fill.patternType != "solid":
        return None
    return _hex6(getattr(fill.fgColor, "rgb", None))


def _font_name_ok(cell, names):
    name = cell.font.name if cell.font is not None else None
    return name in names


def _size_of(cell):
    size = cell.font.sz if cell.font is not None else None
    try:
        return float(size)
    except (TypeError, ValueError):
        return None


def _color_spec(parent):
    """取 DrawingML 填充/文字颜色节点下的第一个颜色定义，归一成可比较的字符串。

    srgbClr -> "srgb:4874CB"；schemeClr -> "scheme:accent1"；sysClr -> "srgb:FFFFFF"
    （sysClr 带 lastClr 时按 lastClr 归一，使"白"能与 srgbClr FFFFFF 互认）。
    """
    if parent is None:
        return None
    for node in parent.iter():
        tag = node.tag.split("}")[-1]
        if tag == "srgbClr":
            return "srgb:" + (_hex6(node.get("val")) or "")
        if tag == "sysClr":
            last = _hex6(node.get("lastClr"))
            if last:
                return "srgb:" + last
            return "sys:" + str(node.get("val"))
        if tag == "schemeClr":
            return "scheme:" + str(node.get("val"))
        if tag == "prstClr":
            return "prst:" + str(node.get("val"))
    return None


WHITE_SPECS = ("srgb:FFFFFF", "sys:window", "scheme:bg1", "scheme:lt1",
               "scheme:background1", "prst:white")


def _is_white(spec):
    return spec in WHITE_SPECS


def _read_xml(zf, name):
    try:
        return ET.fromstring(zf.read(name))
    except (KeyError, ET.ParseError):
        return None


def _rels_of(zf, part_name):
    """读取某个部件的 .rels，返回 {Id: 规范化后的部件名}。"""
    folder = posixpath.dirname(part_name)
    rels_name = posixpath.join(folder, "_rels", posixpath.basename(part_name) + ".rels")
    root = _read_xml(zf, rels_name)
    out = {}
    if root is None:
        return out
    for rel in root.findall("{%s}Relationship" % NS_PR):
        target = rel.get("Target") or ""
        if target.startswith("/"):
            resolved = target.lstrip("/")
        else:
            resolved = posixpath.normpath(posixpath.join(folder, target))
        out[rel.get("Id")] = resolved
    return out


def _layout_of(node, default_center_x=False):
    """解析 c:layout/c:manualLayout，返回 (x, y, w, h) 四个占图表区的比例。

    xMode/yMode = "edge" 时 x/y 是绝对比例；= "factor" 时 x/y 是相对默认位置的偏移量
    （标题默认水平居中，其默认左边缘为 (1-w)/2，因此 factor 模式下左边缘 = (1-w)/2 + x）。
    """
    if node is None:
        return None
    manual = node.find(_c("manualLayout"))
    if manual is None:
        return None

    def val(tag):
        el = manual.find(_c(tag))
        if el is None or el.get("val") is None:
            return None
        try:
            return float(el.get("val"))
        except ValueError:
            return None

    def mode(tag):
        el = manual.find(_c(tag))
        return el.get("val") if el is not None else None

    x, y, w, h = val("x"), val("y"), val("w"), val("h")
    if None in (x, y, w, h):
        return None
    if mode("xMode") != "edge":
        x = ((1.0 - w) / 2.0 if default_center_x else 0.0) + x
    if mode("yMode") != "edge":
        y = 0.0 + y
    return (x, y, w, h)


def _pt_cache(container):
    """取 c:cat / c:val 下 strCache/numCache 的缓存点，返回 (ptCount, [文本...])。"""
    if container is None:
        return (None, [])
    count = None
    for node in container.iter(_c("ptCount")):
        try:
            count = int(node.get("val"))
        except (TypeError, ValueError):
            count = None
        break
    pts = []
    for node in container.iter(_c("pt")):
        v = node.find(_c("v"))
        pts.append("" if v is None or v.text is None else v.text)
    return (count, pts)


def _parse_chart(root):
    """把一个 chart 部件解析成后续判分需要的结构化事实。"""
    info = {
        "is_doughnut": False, "hole": None, "first_ang": 0.0, "vary_colors": True,
        "n_pts": 0, "ser_explosion": 0.0, "pts": [],
        "cat_ref": None, "cat_is_ref": False, "cat_cache": [], "cat_count": None,
        "val_ref": None, "val_is_ref": False, "val_cache": [], "val_count": None,
        "lbl_flags": {}, "lbl_numfmt": None, "lbl_fill": None, "lbl_ln_nofill": False,
        "lbl_default_color": None, "lbl_colors": [], "lbl_texts": {},
        "has_legend": False, "legend_overlay": None,
        "title_text": "", "title_deleted": False, "title_layout": None,
        "title_overlay": False, "title_sizes": [],
        "space_fill": None, "space_nofill": False, "space_ln_nofill": False,
        "has_axes": False, "has_gridlines": False, "plot_layout": None,
    }
    if root is None:
        return info

    dough = root.find(".//" + _c("doughnutChart"))
    info["is_doughnut"] = dough is not None
    plot = root.find(".//" + _c("plotArea"))
    if plot is not None:
        info["plot_layout"] = _layout_of(plot.find(_c("layout")))
        for tag in ("catAx", "valAx", "dateAx", "serAx"):
            if plot.find(_c(tag)) is not None:
                info["has_axes"] = True
        for tag in ("majorGridlines", "minorGridlines"):
            if plot.find(".//" + _c(tag)) is not None:
                info["has_gridlines"] = True

    deleted = root.find(".//" + _c("autoTitleDeleted"))
    info["title_deleted"] = deleted is not None and deleted.get("val") in ("1", "true")
    title = root.find(".//" + _c("title"))
    if title is not None:
        info["title_text"] = _norm("".join(
            t.text or "" for t in title.iter(_a("t"))))
        info["title_layout"] = _layout_of(title.find(_c("layout")), default_center_x=True)
        ov = title.find(_c("overlay"))
        info["title_overlay"] = ov is not None and ov.get("val") in ("1", "true")
        for rpr in title.iter():
            if rpr.tag in (_a("rPr"), _a("defRPr")) and rpr.get("sz"):
                try:
                    info["title_sizes"].append(float(rpr.get("sz")) / 100.0)
                except ValueError:
                    pass

    legend = root.find(".//" + _c("legend"))
    info["has_legend"] = legend is not None

    space_sp = root.find(_c("spPr"))
    if space_sp is not None:
        info["space_nofill"] = space_sp.find(_a("noFill")) is not None
        info["space_fill"] = _color_spec(space_sp.find(_a("solidFill")))
        ln = space_sp.find(_a("ln"))
        info["space_ln_nofill"] = ln is None or ln.find(_a("noFill")) is not None
    else:
        info["space_ln_nofill"] = True
    if dough is None:
        return info

    hole = dough.find(".//" + _c("holeSize"))
    if hole is not None:
        try:
            info["hole"] = float(hole.get("val"))
        except (TypeError, ValueError):
            info["hole"] = None
    ang = dough.find(".//" + _c("firstSliceAng"))
    if ang is not None:
        try:
            info["first_ang"] = float(ang.get("val")) % 360.0
        except (TypeError, ValueError):
            info["first_ang"] = 0.0
    vary = dough.find(_c("varyColors"))
    # c:varyColors 缺省即为 true（按扇区自动分配不同颜色）
    info["vary_colors"] = True if vary is None else vary.get("val") in ("1", "true")

    ser = dough.find(_c("ser"))
    if ser is None:
        return info
    ser_ex = ser.find(_c("explosion"))
    if ser_ex is not None:
        try:
            info["ser_explosion"] = float(ser_ex.get("val"))
        except (TypeError, ValueError):
            info["ser_explosion"] = 0.0

    for key in ("cat", "val"):
        el = ser.find(_c(key))
        if el is None:
            continue
        fnode = el.find(".//" + _c("f"))
        if fnode is not None and fnode.text:
            info[key + "_ref"] = fnode.text.strip()
            info[key + "_is_ref"] = True
        count, pts = _pt_cache(el)
        info[key + "_count"] = count
        info[key + "_cache"] = pts

    ref_len = _ref_len(info["val_ref"]) or _ref_len(info["cat_ref"]) or 0
    dpts = {}
    for dp in ser.findall(_c("dPt")):
        idx = dp.find(_c("idx"))
        try:
            dpts[int(idx.get("val"))] = dp
        except (TypeError, ValueError, AttributeError):
            continue
    n_pts = max(info["val_count"] or 0, info["cat_count"] or 0, ref_len,
                (max(dpts) + 1) if dpts else 0)
    info["n_pts"] = n_pts
    ser_sp = ser.find(_c("spPr"))
    for i in range(n_pts):
        dp = dpts.get(i)
        sp = dp.find(_c("spPr")) if dp is not None else None
        if sp is None:
            sp = ser_sp
        patt = sp.find(".//" + _a("pattFill")) if sp is not None else None
        solid = sp.find(_a("solidFill")) if sp is not None else None
        ln = sp.find(_a("ln")) if sp is not None else None
        glow = sp.find(".//" + _a("glow")) if sp is not None else None
        shadow = None
        if sp is not None:
            for tag in ("outerShdw", "innerShdw", "prstShdw"):
                if sp.find(".//" + _a(tag)) is not None:
                    shadow = tag
                    break
        explosion = info["ser_explosion"]
        if dp is not None:
            dpex = dp.find(_c("explosion"))
            if dpex is not None:
                try:
                    explosion = float(dpex.get("val"))
                except (TypeError, ValueError):
                    pass
        ln_w = None
        if ln is not None and ln.get("w"):
            try:
                ln_w = float(ln.get("w"))
            except ValueError:
                ln_w = None
        info["pts"].append({
            "patt": patt.get("prst") if patt is not None else None,
            "fill": (_color_spec(patt.find(_a("fgClr"))) if patt is not None
                     else _color_spec(solid)),
            "ln_w": ln_w,
            "ln_color": _color_spec(ln.find(_a("solidFill"))) if ln is not None else None,
            "ln_nofill": ln is not None and ln.find(_a("noFill")) is not None,
            "glow": _color_spec(glow) if glow is not None else None,
            "glow_rad": glow.get("rad") if glow is not None else None,
            "shadow": shadow,
            "explosion": explosion,
        })

    dl = ser.find(_c("dLbls"))
    if dl is not None:
        for tag in ("showVal", "showPercent", "showCatName", "showSerName",
                    "showLegendKey", "showBubbleSize"):
            node = dl.find(_c(tag))
            info["lbl_flags"][tag] = (node is not None and
                                      node.get("val") in ("1", "true"))
        nf = dl.find(_c("numFmt"))
        info["lbl_numfmt"] = nf.get("formatCode") if nf is not None else None
        sp = dl.find(_c("spPr"))
        if sp is not None:
            info["lbl_fill"] = _color_spec(sp.find(_a("solidFill")))
            if sp.find(_a("noFill")) is not None:
                info["lbl_fill"] = "none"
            ln = sp.find(_a("ln"))
            info["lbl_ln_nofill"] = ln is not None and ln.find(_a("noFill")) is not None
        txpr = dl.find(_c("txPr"))
        if txpr is not None:
            defrpr = txpr.find(".//" + _a("defRPr"))
            info["lbl_default_color"] = _color_spec(
                defrpr.find(_a("solidFill")) if defrpr is not None else None)
        per_idx = {}
        for one in dl.findall(_c("dLbl")):
            idx = one.find(_c("idx"))
            try:
                i = int(idx.get("val"))
            except (TypeError, ValueError, AttributeError):
                continue
            text = _norm("".join(t.text or "" for t in one.iter(_a("t"))))
            if text:
                info["lbl_texts"][i] = text
            color = None
            for rpr in one.iter(_a("rPr")):
                color = _color_spec(rpr.find(_a("solidFill")))
                if color:
                    break
            per_idx[i] = color
        info["lbl_colors"] = [per_idx.get(i) or info["lbl_default_color"]
                              for i in range(n_pts)]
    return info


# ---------------------------------------------------------------------------
# 单元格引用解析：图表数据源允许直接指向 B3:B7/D3:D7，
# 也允许指向内容与之一致的中转数据列（中转列常以 =B3、=D3 这类公式取值）
# ---------------------------------------------------------------------------
_REF_RE = re.compile(
    r"\$?([A-Za-z]{1,3})\$?(\d+)(?::\$?([A-Za-z]{1,3})\$?(\d+))?")


def _col_to_idx(letters):
    idx = 0
    for ch in letters.upper():
        idx = idx * 26 + (ord(ch) - 64)
    return idx


def _parse_ref(ref):
    """把 "Sheet1!$B$3:$B$7" / "'Sheet1'!$M$3:$M$7" 解析成 (sheet, r1, c1, r2, c2)。"""
    if not ref:
        return None
    text = str(ref).strip().lstrip("=")
    if "," in text or ";" in text:
        text = re.split(r"[,;]", text)[0]
    sheet = None
    if "!" in text:
        sheet, _, text = text.rpartition("!")
        sheet = sheet.strip().strip("'").strip()
    m = _REF_RE.fullmatch(text.strip())
    if m is None:
        return None
    c1 = _col_to_idx(m.group(1))
    r1 = int(m.group(2))
    c2 = _col_to_idx(m.group(3) or m.group(1))
    r2 = int(m.group(4) or m.group(2))
    return (sheet or None, min(r1, r2), min(c1, c2), max(r1, r2), max(c1, c2))


def _ref_len(ref):
    parsed = _parse_ref(ref)
    if parsed is None:
        return 0
    _, r1, c1, r2, c2 = parsed
    return (r2 - r1 + 1) * (c2 - c1 + 1)


def _resolve_cell(wb, sheet_name, row, col, depth=0):
    """取单元格的值；若该单元格是"= 单个单元格"的中转公式，则跟随到源单元格取值。"""
    if depth > 6:
        return None
    try:
        ws = wb[sheet_name] if sheet_name in wb.sheetnames else wb[wb.sheetnames[0]]
    except (KeyError, IndexError):
        return None
    value = ws.cell(row=row, column=col).value
    if isinstance(value, str) and value.startswith("="):
        parsed = _parse_ref(value)
        if parsed is not None:
            nsheet, r1, c1, r2, c2 = parsed
            if (r1, c1) == (r2, c2):
                return _resolve_cell(wb, nsheet or sheet_name, r1, c1, depth + 1)
        return None
    return value


def _range_values(wb, ref):
    """把图表数据源引用解析成一串实际取值（公式中转列会被跟随到源单元格）。"""
    parsed = _parse_ref(ref)
    if parsed is None:
        return None
    sheet, r1, c1, r2, c2 = parsed
    out = []
    for row in range(r1, r2 + 1):
        for col in range(c1, c2 + 1):
            out.append(_resolve_cell(wb, sheet, row, col))
    return out


def _to_float(value):
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    if isinstance(value, str):
        text = value.strip().rstrip("%")
        try:
            num = float(text)
        except ValueError:
            return None
        return num / 100.0 if value.strip().endswith("%") else num
    return None


# ---------------------------------------------------------------------------
# 绘图层（drawing）解析：图表锚定方式、位置、尺寸，以及图表内的附加图形/文本框
# ---------------------------------------------------------------------------
def _col_width_emu(ws, col_idx):
    letter = get_column_letter(col_idx + 1)
    dim = ws.column_dimensions.get(letter)
    width = getattr(dim, "width", None) if dim is not None else None
    if not width:
        width = getattr(ws.sheet_format, "defaultColWidth", None) or 8.43
    return (round(float(width) * 7) + 5) * 9525.0


def _row_height_emu(ws, row_idx):
    dim = ws.row_dimensions.get(row_idx + 1)
    height = getattr(dim, "height", None) if dim is not None else None
    if not height:
        height = getattr(ws.sheet_format, "defaultRowHeight", None) or 13.5
    return float(height) * EMU_PER_PT


def _cell_origin(ws, col, col_off, row, row_off):
    """把 xdr:from / xdr:to 的单元格坐标换算成工作表内的绝对 EMU 坐标。"""
    x = sum(_col_width_emu(ws, i) for i in range(0, max(col, 0))) + (col_off or 0)
    y = sum(_row_height_emu(ws, i) for i in range(0, max(row, 0))) + (row_off or 0)
    return (x, y)


def _anchor_cell(node):
    if node is None:
        return None

    def get(tag):
        el = node.find(_xdr(tag))
        try:
            return int(el.text)
        except (TypeError, ValueError, AttributeError):
            return 0
    return (get("col"), get("colOff"), get("row"), get("rowOff"))


def _xfrm_box(node):
    """从 a:xfrm 取 (x, y, cx, cy)，缺失或为 0 的分量返回 None。"""
    if node is None:
        return (None, None, None, None)
    off = node.find(_a("off"))
    ext = node.find(_a("ext"))

    def num(el, key):
        if el is None or el.get(key) is None:
            return None
        try:
            value = float(el.get(key))
        except ValueError:
            return None
        return value
    return (num(off, "x"), num(off, "y"), num(ext, "cx"), num(ext, "cy"))


def _parse_drawings(zf, ws):
    """返回 (图表锚点列表, 附加图形列表)。"""
    chart_anchors = []
    shapes = []
    for name in sorted(zf.namelist()):
        if not re.fullmatch(r"xl/drawings/drawing\d+\.xml", name):
            continue
        root = _read_xml(zf, name)
        if root is None:
            continue
        rels = _rels_of(zf, name)
        for anchor in list(root):
            tag = anchor.tag.split("}")[-1]
            if tag not in ("twoCellAnchor", "oneCellAnchor", "absoluteAnchor"):
                continue
            frm = _anchor_cell(anchor.find(_xdr("from")))
            to = _anchor_cell(anchor.find(_xdr("to")))
            aext = anchor.find(_xdr("ext"))
            apos = anchor.find(_xdr("pos"))

            def resolve(xfrm):
                x, y, cx, cy = _xfrm_box(xfrm)
                if not cx or not cy:
                    if aext is not None:
                        try:
                            cx = float(aext.get("cx"))
                            cy = float(aext.get("cy"))
                        except (TypeError, ValueError):
                            pass
                if (not cx or not cy) and frm and to and ws is not None:
                    x0, y0 = _cell_origin(ws, *frm)
                    x1, y1 = _cell_origin(ws, *to)
                    cx, cy = x1 - x0, y1 - y0
                if not x or not y:
                    if frm and ws is not None:
                        x, y = _cell_origin(ws, *frm)
                    elif apos is not None:
                        try:
                            x, y = float(apos.get("x")), float(apos.get("y"))
                        except (TypeError, ValueError):
                            pass
                return (x, y, cx, cy)

            gf = anchor.find(_xdr("graphicFrame"))
            if gf is not None:
                ref = gf.find(".//" + _c("chart"))
                rid = ref.get("{%s}id" % NS_R) if ref is not None else None
                chart_anchors.append({
                    "anchor": tag,
                    "from": frm,
                    "to": to,
                    "box": resolve(gf.find(_xdr("xfrm"))),
                    "part": rels.get(rid) if rid else None,
                })
            for sp in anchor.iter(_xdr("sp")):
                spr = sp.find(_xdr("spPr"))
                body = sp.find(".//" + _a("bodyPr"))
                sizes = []
                for rpr in sp.iter():
                    if rpr.tag in (_a("rPr"), _a("defRPr")) and rpr.get("sz"):
                        try:
                            sizes.append(float(rpr.get("sz")) / 100.0)
                        except ValueError:
                            pass
                shapes.append({
                    "text": _norm("".join(t.text or "" for t in sp.iter(_a("t")))),
                    "box": resolve(spr.find(_a("xfrm")) if spr is not None else None),
                    "fill": _color_spec(spr.find(_a("solidFill")) if spr is not None else None),
                    "sizes": sizes,
                    "wrap": body.get("wrap") if body is not None else None,
                })
    return (chart_anchors, shapes)


# ---------------------------------------------------------------------------
# 圆环几何：由图表框尺寸 + 绘图区布局 + 内径比例，算出圆环圆心、外径、空心半径
# ---------------------------------------------------------------------------
def _ring_geometry(chart, box):
    if not box:
        return None
    x, y, cw, ch = box
    if not cw or not ch:
        return None
    x = x or 0.0
    y = y or 0.0
    lay = chart["plot_layout"]
    if lay is None:
        # 未写 manualLayout 时按办公软件默认留边估算；标题未叠加显示会占用顶部空间
        top = 0.03
        if chart["title_text"] and not chart["title_deleted"] and not chart["title_overlay"]:
            top = 0.15
        lay = (0.03, top, 0.94, 0.97 - top)
    px, py = x + lay[0] * cw, y + lay[1] * ch
    pw, ph = lay[2] * cw, lay[3] * ch
    r_out = min(pw, ph) / 2.0
    hole = chart["hole"] if chart["hole"] is not None else 50.0
    return {
        "cx": px + pw / 2.0, "cy": py + ph / 2.0,
        "r_out": r_out, "r_hole": r_out * hole / 100.0,
    }


def _box_corners(box):
    x, y, cx, cy = box
    return [(x, y), (x + cx, y), (x, y + cy), (x + cx, y + cy)]


def _in_hole(box, geo):
    """返回 (中心点是否落在空心区内且水平居中, 四角是否全部落在空心区内)。"""
    if geo is None or not box:
        return (False, False)
    x, y, cx, cy = box
    if None in (x, y, cx, cy) or not cx or not cy:
        return (False, False)
    mx, my = x + cx / 2.0, y + cy / 2.0
    r_hole = geo["r_hole"]
    if r_hole <= 0:
        return (False, False)
    dist_center = math.hypot(mx - geo["cx"], my - geo["cy"])
    centered = (dist_center <= r_hole) and (abs(mx - geo["cx"]) <= 0.25 * r_hole)
    inside = all(math.hypot(px - geo["cx"], py - geo["cy"]) <= r_hole
                 for px, py in _box_corners(box))
    return (centered, inside)


def _text_fits(box, char_count, sizes):
    """估算文字能否在文本框内完整换行显示（中文按全角、行高按 1.25 倍字号计）。"""
    if not box or None in box or not box[2] or not box[3]:
        return False
    size = max(sizes) if sizes else 10.0
    char_w = size * EMU_PER_PT
    line_h = size * 1.25 * EMU_PER_PT
    usable_w = max(box[2] - 76200.0, char_w)
    usable_h = max(box[3] - 45720.0, line_h)
    need = int(math.ceil(char_count * char_w / usable_w))
    allow = max(1, int(usable_h // line_h))
    return need <= allow


# ---------------------------------------------------------------------------
# 对外入口
# ---------------------------------------------------------------------------
def evaluate(dir_path):
    """入口：dir_path 为脚本所在目录，函数自行在其中定位被评估的工作簿。

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
        files = [f for f in sorted(os.listdir(dir_path))
                 if os.path.isfile(os.path.join(dir_path, f))
                 and not f.startswith("~$") and not f.startswith(".")]
        cands = [f for f in files if f.lower().endswith(ACCEPT_EXTS)]
        wrongs = [f for f in files if f.lower().endswith(REJECT_EXTS)]

        # 维度一：只检查交付文件格式与可否正常打开
        if not cands:
            if wrongs:
                # 目录里有交付物但格式不被接受：已有交付动作，交给维度一记 0
                result["file_name"] = wrongs[0]
                return _dim1_fail(
                    result,
                    "交付文件格式为 %s，不是 .xlsx 或 .xlsm"
                    % os.path.splitext(wrongs[0])[1].lower())
            # 目录里一个交付物都没有：这份作业没评上，而不是维度一判不通过
            result["status"] = "error"
            result["error"] = "目录中未找到 .xlsx/.xlsm 文件：%s" % dir_path
            return result

        file_path = os.path.join(dir_path, cands[0])
        result["file_name"] = os.path.basename(file_path)
        try:
            probe = openpyxl.load_workbook(file_path)
            probe.close()
        except Exception as exc:
            return _dim1_fail(result, "工作簿损坏、无法正常打开：%s: %s"
                                      % (type(exc).__name__, exc))
        result["dim1_pass"] = True
        # 约定：维度一通过时 dim1_reason 保持空串，只有未通过才写错误摘要。
        result["dim1_reason"] = ""
        return _evaluate_impl(file_path, result)
    except Exception as exc:
        # 脚本异常同样视为"没评上"：维度一未判定，两项分数归零。
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
    details = []

    def plus(hit, max_delta, rule, detail=""):
        results.append("+%d：%s" % (max_delta if hit else 0, rule))
        details.append(detail)

    def minus(hit, max_delta, rule, detail=""):
        results.append("-%d：%s" % (abs(max_delta) if hit else 0, rule))
        details.append(detail)

    wb = openpyxl.load_workbook(file_path, data_only=False)
    ws = wb[SHEET_NAME] if SHEET_NAME in wb.sheetnames else wb[wb.sheetnames[0]]

    zf = zipfile.ZipFile(file_path)
    try:
        chart_parts = [n for n in sorted(zf.namelist())
                       if re.search(r"charts/chart[^/]*\.xml$", n) and "_rels" not in n]
        charts = dict((n, _parse_chart(_read_xml(zf, n))) for n in chart_parts)
        chart_anchors, shapes = _parse_drawings(zf, ws)
    finally:
        zf.close()

    # 选定被判分的环形图：优先取"被绘图层锚定的圆环图"
    dough_parts = [n for n in chart_parts if charts[n]["is_doughnut"]]
    target_part, anchor = None, None
    for one in chart_anchors:
        if one["part"] in dough_parts:
            target_part, anchor = one["part"], one
            break
    if target_part is None and dough_parts:
        target_part = dough_parts[0]
    chart = charts.get(target_part) or _parse_chart(None)
    has_chart = target_part is not None
    box = anchor["box"] if anchor else None
    geo = _ring_geometry(chart, box) if box else None
    pts = chart["pts"]
    n_pts = chart["n_pts"]

    # 图表数据源：缓存优先，缺少缓存时按引用回溯取值（中转公式列会跟随到源单元格）
    cat_src = ([_norm(t) for t in chart["cat_cache"]] if len(chart["cat_cache"]) >= 5
               else [_norm(v) for v in (_range_values(wb, chart["cat_ref"]) or [])])
    val_raw = (chart["val_cache"] if len(chart["val_cache"]) >= 5
               else (_range_values(wb, chart["val_ref"]) or []))
    val_src = [_to_float(v) for v in val_raw]
    exp_ratios = [COUNTS[i] / float(SAMPLE_TOTAL) for i in range(5)]
    cat_ok = cat_src[:5] == [_norm(o) for o in Q1_OPTIONS]
    val_ok = (len(val_src) >= 5 and all(
        v is not None and abs(v - exp_ratios[i]) <= 5e-4 for i, v in enumerate(val_src[:5])))
    # ---------------- 01 存在 1 个可编辑的环形图对象 ----------------
    # ①工作簿内存在 doughnutChart 类型的图表部件；②该部件由工作表绘图层的
    # graphicFrame 引用（即真正嵌在表里的图表对象，不是插入的图片/截图）；
    # ③图表保留了系列与数据源节点，可继续编辑数据源、扇区颜色与数据标签。
    # 与办公软件对齐：读取 xl/charts/chartN.xml 的 c:doughnutChart，
    # 以及 xl/drawings/drawingN.xml 中 xdr:graphicFrame/a:graphic/c:chart 的 r:id 关系。
    ok01 = has_chart and chart["is_doughnut"] and anchor is not None
    plus(ok01, 5,
         "工作簿中存在1个环形图（圆环/doughnut）类型的图表对象，为可编辑的图表而非图片或截图，"
         "可正常选中并编辑其数据源、扇区颜色与数据标签",
         "图表部件=%s；锚定节点=%s" % (target_part, anchor["anchor"] if anchor else None))

    # ---------------- 02 只呈现第一题的 5 个选项 ----------------
    # ①系列数据点数量为 5；②类别文字恰为该题的 5 个选项且顺序一致；
    # ③未把其他调查事项的行一起纳入类别区域。
    # 与办公软件对齐：读取 c:ser/c:cat 的缓存点或其引用区域的实际取值。
    ok02 = n_pts == 5 and cat_ok
    plus(ok02, 3,
         '图表只呈现"您对村内便民服务事项的知晓程度"一题的5个选项（十分清楚、比较清楚、'
         '有所了解、了解较少、从未关注），未混入其他调查事项的数据',
         "数据点数=%s；类别=%s" % (n_pts, cat_src[:5]))

    # ---------------- 03 类别与数值均引用单元格且内容一致 ----------------
    # ①c:cat 与 c:val 都是 strRef/numRef（引用单元格区域）而非 strLit/numLit 常量数组；
    # ②类别文字与 5 个选项一致；③数值与 5 个比例一致（11.89%、23.43%、31.82%、
    # 24.48%、8.39%，即该行有效人数 / 286）。中转数据列若以公式取值，回溯到源单元格判定。
    ok03 = chart["cat_is_ref"] and chart["val_is_ref"] and cat_ok and val_ok
    plus(ok03, 5,
         "图表的类别名称与数值系列均引用单元格区域，且内容与B3:B7的5个选项文字、"
         "D3:D7的5个比例（11.89%、23.43%、31.82%、24.48%、8.39%）完全一致",
         "cat=%s；val=%s；数值=%s" % (chart["cat_ref"], chart["val_ref"],
                                   [None if v is None else round(v, 6) for v in val_src[:5]]))

    # ---------------- 04 置于问题右侧空白区（起始列 E 或其右） ----------------
    # 锚点起始列索引 ≥ 4（即 E 列），从而不压在 A 至 D 四列的表格内容上。
    from_col = anchor["from"][0] if (anchor and anchor["from"]) else None
    ok04 = from_col is not None and from_col >= 4
    plus(ok04, 3,
         '图表放置在该问题右侧的空白区域，左上角起始列为E列或其右侧，'
         '未压在"调查事项""回答选项""有效人数""比例"四列的表格内容上',
         "起始列索引=%s" % from_col)

    # ---------------- 05 图表上边缘位于第 1 至第 8 行之间 ----------------
    from_row = anchor["from"][2] if (anchor and anchor["from"]) else None
    ok05 = from_row is not None and 0 <= from_row <= 7
    plus(ok05, 1,
         '图表上边缘位于第1至第8行之间，与"您对村内便民服务事项的知晓程度"所在行相邻齐平',
         "起始行=%s" % (None if from_row is None else from_row + 1))

    # ---------------- 06 以单元格锚定方式定位 ----------------
    # twoCellAnchor / oneCellAnchor 会随行列插入删除移动；absoluteAnchor 则脱离网格漂浮。
    ok06 = anchor is not None and anchor["anchor"] in ("twoCellAnchor", "oneCellAnchor")
    plus(ok06, 1,
         "图表以单元格锚定方式定位（随行列插入删除而移动），未脱离单元格网格漂浮",
         "锚定方式=%s" % (anchor["anchor"] if anchor else None))
    # ---------------- 07 / 08 问题文字置于圆环正中的空心区 ----------------
    # 候选载体有两种：图表标题（c:title，位置由 c:layout/manualLayout 决定）与叠放在
    # 图表上的文本框（xdr:sp）。判定步骤：①先由图表框尺寸与绘图区布局算出圆环圆心、
    # 外径，再乘 c:holeSize 得到空心半径；②文字块中心点须落在空心圆内，且水平方向与
    # 圆心偏差不超过空心半径的 25%（即水平居中）；③四角全部落在空心圆内才算完整放在
    # 中心留白里、未溢出到扇区之外。注意 xMode="factor" 是相对默认位置的偏移量，
    # 居中标题的默认左边缘为 (1-w)/2，因此左边缘 = (1-w)/2 + x；xMode="edge" 才是绝对比例。
    q_norm = _norm(Q1_TEXT)
    cands = []
    if box and chart["title_text"] and q_norm in chart["title_text"] and not chart["title_deleted"]:
        lay = chart["title_layout"] or (0.35, 0.02, 0.30, 0.10)
        cands.append({"src": "图表标题",
                      "box": ((box[0] or 0) + lay[0] * box[2], (box[1] or 0) + lay[1] * box[3],
                              lay[2] * box[2], lay[3] * box[3]),
                      "sizes": chart["title_sizes"], "wrap": None})
    for sp in shapes:
        if q_norm and q_norm in sp["text"]:
            cands.append({"src": "文本框", "box": sp["box"],
                          "sizes": sp["sizes"], "wrap": sp["wrap"]})
    best = None
    for cand in cands:
        cand["centered"], cand["inside"] = _in_hole(cand["box"], geo)
        if best is None or (cand["centered"], cand["inside"]) > (best["centered"], best["inside"]):
            best = cand
    ok07 = bool(best and best["centered"])
    plus(ok07, 5,
         '问题文字"您对村内便民服务事项的知晓程度"完整放置在圆环正中间的空心区域内，'
         '水平居中、无截断、无缺字',
         "文字载体=%s；居中=%s；圆心=%s；空心半径=%s" % (
             best["src"] if best else None, ok07,
             None if geo is None else (round(geo["cx"]), round(geo["cy"])),
             None if geo is None else round(geo["r_hole"])))

    ok08 = bool(best and best["inside"] and best["wrap"] != "none"
                and _text_fits(best["box"], len(Q1_TEXT), best["sizes"]))
    plus(ok08, 3,
         "圆环中心的问题文字可在中心留白内自动换行完整显示，未溢出到扇区之外，"
         "未被扇区或数据标签压盖",
         "四角均在空心区内=%s" % (best["inside"] if best else None))

    # ---------------- 09 圆环内径占外径不小于 55% ----------------
    ok09 = chart["hole"] is not None and chart["hole"] >= 55
    plus(ok09, 1,
         "圆环内径（空心）占外径比例不小于55%，呈细环状，中心留白足以容纳问题文字",
         "holeSize=%s" % chart["hole"])

    # ---------------- 10 5 个扇区均为斜线/斜纹图案填充 ----------------
    # 每个数据点的 c:dPt/c:spPr 需为 a:pattFill 且图案为斜线类（prst 含 Diag），
    # 纯色 a:solidFill 平涂不计分。
    diag = [p for p in pts if p["patt"] and "diag" in str(p["patt"]).lower()]
    ok10 = n_pts == 5 and len(diag) == 5
    plus(ok10, 5,
         "5个扇区均为斜线/斜纹图案填充（如细密下斜线纹理），而非纯色平涂填充",
         "图案填充=%s" % [p["patt"] for p in pts])

    # ---------------- 11 5 个扇区填充颜色各不相同 ----------------
    # 显式写了 5 个扇区颜色时要求 5 色互不相同；未逐点写色时按 c:varyColors 判定
    # （该元素缺省即为 true，办公软件会自动为各扇区分配不同颜色）。
    explicit = [p["fill"] for p in pts if p["fill"]]
    if len(explicit) >= 5:
        ok11 = len(set(explicit[:5])) == 5
    else:
        ok11 = bool(chart["vary_colors"]) and n_pts == 5
    plus(ok11, 3,
         "5个扇区的填充颜色各不相同，与5个选项一一对应",
         "扇区颜色=%s；varyColors=%s" % (explicit[:5] or None, chart["vary_colors"]))

    # ---------------- 12 相邻扇区间白色描边 3 至 5 磅 ----------------
    ok12 = n_pts == 5 and all(
        p["ln_w"] is not None and 3.0 * EMU_PER_PT - 1 <= p["ln_w"] <= 5.0 * EMU_PER_PT + 1
        and _is_white(p["ln_color"]) for p in pts)
    plus(ok12, 3,
         "相邻扇区之间有白色描边作为分隔（描边宽度约3至5磅），呈现清晰的断开间隔",
         "描边=%s" % [(None if p["ln_w"] is None else round(p["ln_w"] / EMU_PER_PT, 2),
                     p["ln_color"]) for p in pts])
    # ---------------- 13 每个扇区带柔和发光或投影 ----------------
    # 每个数据点的 spPr/a:effectLst 下需有 a:glow 或 a:outerShdw/a:prstShdw；
    # 若用发光，发光色应与该扇区填充色同色（半透明由 a:alpha/satMod 控制）。
    glow_match = all((p["glow"] is None or p["fill"] is None or p["glow"] == p["fill"])
                     for p in pts)
    ok13 = n_pts == 5 and all(p["glow"] or p["shadow"] for p in pts) and glow_match
    plus(ok13, 3,
         "每个扇区外缘带柔和发光或投影效果（发光色与该扇区填充色同色、半透明，"
         "边缘呈由深到浅的光晕过渡），而非无任何边缘效果的平面扇区",
         "发光/投影=%s" % [(p["glow"], p["glow_rad"], p["shadow"]) for p in pts])

    # ---------------- 14 有且仅有 1 个扇区被向外拉出 ----------------
    # 逐点有效外移量 = 该点 c:dPt/c:explosion，缺省则继承系列级 c:ser/c:explosion。
    # 系列级 explosion 会让 5 个扇区同时外移，不构成"单独强调某一个选项"。
    exploded = [i for i, p in enumerate(pts) if (p["explosion"] or 0) > 0]
    ok14 = n_pts == 5 and len(exploded) == 1
    plus(ok14, 5,
         "圆环中有且仅有1个扇区被整体沿半径方向向外拉出，与相邻扇区之间形成明显的断开缺口，"
         "达到单独强调该选项的效果；其余4个扇区仍紧密相连构成连续圆环",
         "外移扇区=%s；系列级外移=%s；逐点外移=%s" % (
             exploded, chart["ser_explosion"], [p["explosion"] for p in pts]))

    # ---------------- 15 被拉出扇区位于右上方、外移 1/4 至 1/3 半径 ----------------
    # 起始角自 12 点方向顺时针计，扇区中心角 = firstSliceAng + 前序占比累计角 + 本扇区半角；
    # 落在 0°至 60° 即 12 点至 2 点方向。外移量按半径百分比计，需在 25% 至 33% 之间。
    mid_ang = None
    ok15 = False
    if ok14 and len(val_src) >= 5 and all(v is not None for v in val_src[:5]):
        total = sum(val_src[:5]) or 1.0
        idx = exploded[0]
        start = chart["first_ang"] + 360.0 * sum(val_src[:idx]) / total
        mid_ang = (start + 360.0 * val_src[idx] / total / 2.0) % 360.0
        pull = pts[idx]["explosion"] or 0
        ok15 = (mid_ang <= 60.0 or mid_ang >= 355.0) and 25.0 <= pull <= 33.5
    plus(ok15, 3,
         "被拉出的扇区位于圆环右上方（起始角大致在12点至2点方向），外移距离约为圆环半径的"
         "四分之一至三分之一，既能看出分离，又未远离圆环整体轮廓",
         "扇区中心角=%s；外移比例=%s%%" % (
             None if mid_ang is None else round(mid_ang, 1),
             pts[exploded[0]]["explosion"] if ok14 else None))

    # ---------------- 16 被拉出扇区与其余扇区样式一致 ----------------
    # 与其余 4 个扇区比较图案填充类型、描边宽度、描边颜色与边缘效果的有无。
    ok16 = False
    if ok14:
        one = pts[exploded[0]]
        others = [p for i, p in enumerate(pts) if i != exploded[0]]
        ok16 = all(
            p["patt"] == one["patt"]
            and abs((p["ln_w"] or 0) - (one["ln_w"] or 0)) <= 1
            and p["ln_color"] == one["ln_color"]
            and bool(p["glow"] or p["shadow"]) == bool(one["glow"] or one["shadow"])
            for p in others)
    plus(ok16, 1,
         "被拉出的扇区与其余扇区保持相同的环宽、斜线纹理、白色描边与发光效果，仅位置外移，"
         "未被单独放大、拉伸变形或改为纯色填充",
         "样式一致=%s" % ok16)
    # ---------------- 17 5 个扇区均显示百分比数据标签且数值一致 ----------------
    # ①标签开关：c:dLbls/c:showPercent 为 1，或 c:showVal 为 1 且标签数字格式为百分比
    # （c:dLbls/c:numFmt 缺省时沿用数值源单元格的数字格式）；②5 个数值与比例列一致；
    # ③若标签被改写成固定文字，文字须为 11.89%/23.43%/31.82%/24.48%/8.39%
    # 或取整后的 12%/23%/32%/24%/8%。
    def _src_numfmt(ref):
        parsed = _parse_ref(ref)
        if parsed is None:
            return None
        sheet, r1, c1, _, _ = parsed
        sheet_obj = wb[sheet] if (sheet and sheet in wb.sheetnames) else ws
        return sheet_obj.cell(row=r1, column=c1).number_format

    accept_pct = set(Q1_RATIO_STRS) | set(Q1_PCT_STRS)
    flags = chart["lbl_flags"]
    eff_fmt = chart["lbl_numfmt"] or _src_numfmt(chart["val_ref"]) or ""
    pct_shown = bool(flags.get("showPercent")) or (
        bool(flags.get("showVal")) and "%" in str(eff_fmt))
    texts_ok = all(t in accept_pct for t in chart["lbl_texts"].values())
    ok17 = n_pts == 5 and pct_shown and val_ok and texts_ok
    plus(ok17, 5,
         "5个扇区均显示百分比数据标签，数值与比例列一致（11.89%、23.43%、31.82%、24.48%、"
         "8.39%，或取整显示为12%、23%、32%、24%、8%）",
         "标签开关=%s；数字格式=%s；固定文字=%s" % (
             dict((k, v) for k, v in flags.items() if v), eff_fmt,
             sorted(chart["lbl_texts"].values())))

    # ---------------- 18 百分比标签带白色底框、无边框线 ----------------
    ok18 = _is_white(chart["lbl_fill"]) and chart["lbl_ln_nofill"]
    plus(ok18, 3,
         "百分比数据标签带白色标签底框、无边框线，压在所属扇区上呈现",
         "标签底框=%s；边框线无填充=%s" % (chart["lbl_fill"], chart["lbl_ln_nofill"]))

    # ---------------- 19 百分比标签文字颜色随所属扇区 ----------------
    # 逐个标签（c:dLbl/a:rPr/a:solidFill）的文字颜色应各不相同、与所属扇区对应，
    # 5 个标签统一一种文字颜色不计分。
    lbl_colors = [c for c in chart["lbl_colors"][:5] if c]
    ok19 = len(set(lbl_colors)) >= 4
    plus(ok19, 3,
         "百分比标签文字颜色与其所属扇区颜色一致，不是5个标签统一使用同一种文字颜色",
         "标签文字颜色=%s" % chart["lbl_colors"][:5])

    # ---------------- 20 数据标签仅显示百分比 ----------------
    ok20 = (pct_shown
            and not any(flags.get(k) for k in ("showCatName", "showSerName",
                                               "showLegendKey", "showBubbleSize"))
            and not (flags.get("showVal") and flags.get("showPercent")))
    plus(ok20, 1,
         "5个数据标签仅显示百分比数值，未同时叠加类别名称、系列名称、有效人数等造成文字重叠",
         "开启的标签内容=%s" % [k for k, v in flags.items() if v])

    # ---------------- 21 提供选项名称与扇区颜色的对应标识 ----------------
    # 两种实现都认可：①图表自带图例 c:legend；②图表上叠放色块图形 + 选项名称文字，
    # 要求 5 个选项名齐全，且色块图形的填充色不少于 5 种。
    shape_text = "".join(sp["text"] for sp in shapes)
    names_ok = all(_norm(o) in shape_text for o in Q1_OPTIONS)
    swatches = set(sp["fill"] for sp in shapes if sp["fill"] and not sp["text"])
    ok21 = bool(chart["has_legend"]) or (names_ok and len(swatches) >= 5)
    plus(ok21, 5,
         '图表中提供选项名称与扇区颜色的对应标识（图例，或色块加"十分清楚""比较清楚"'
         '"有所了解""了解较少""从未关注"文字说明），5个选项齐全，可据此分辨每个扇区',
         "图例=%s；选项名齐全=%s；色块=%s" % (chart["has_legend"], names_ok, sorted(swatches)))
    # ---------------- 22 图表区无边框线、白色/无填充、无网格线坐标轴 ----------------
    # 读 chartSpace/c:spPr：a:ln 为 a:noFill（无边框线）；填充为 a:noFill 或白色
    # （srgbClr FFFFFF / sysClr window / schemeClr bg1、lt1）；绘图区内不应出现
    # c:catAx、c:valAx 等坐标轴与 c:majorGridlines 网格线。
    fill_ok = (chart["space_nofill"] or chart["space_fill"] is None
               or _is_white(chart["space_fill"]))
    ok22 = (has_chart and chart["space_ln_nofill"] and fill_ok
            and not chart["has_axes"] and not chart["has_gridlines"])
    plus(ok22, 1,
         "环形图图表区无边框线、背景为白色或无填充，无多余网格线与坐标轴",
         "边框无填充=%s；背景=%s；坐标轴=%s；网格线=%s" % (
             chart["space_ln_nofill"], chart["space_fill"] or ("无填充" if chart["space_nofill"] else None),
             chart["has_axes"], chart["has_gridlines"]))

    # ---------------- 23 宽度 ≥ 8 厘米、高度 ≥ 6 厘米 ----------------
    # 尺寸取 xdr:graphicFrame/a:xfrm/a:ext，缺失时取 oneCellAnchor 的 xdr:ext，
    # 再退化为 from/to 单元格跨度；1 厘米 = 360000 EMU。
    width_cm = (box[2] / EMU_PER_CM) if (box and box[2]) else None
    height_cm = (box[3] / EMU_PER_CM) if (box and box[3]) else None
    ok23 = bool(width_cm and height_cm and width_cm >= 8.0 - 0.05 and height_cm >= 6.0 - 0.05)
    plus(ok23, 1,
         "环形图宽度不小于8厘米、高度不小于6厘米，圆环与全部标签完整显示不被裁切",
         "尺寸=%s×%s厘米" % (None if width_cm is None else round(width_cm, 2),
                          None if height_cm is None else round(height_cm, 2)))

    # ---------------- 24 图表已含缓存数据，打开即正常显示 ----------------
    # c:cat/c:strCache 与 c:val/c:numCache 均应写入 5 个缓存点（c:ptCount val="5"），
    # 缺少缓存时办公软件需重算才能显示扇区与标签。
    cat_cached = chart["cat_count"] == 5 and len([t for t in chart["cat_cache"] if t]) == 5
    val_cached = chart["val_count"] == 5 and len([t for t in chart["val_cache"] if t]) == 5
    ok24 = cat_cached and val_cached
    plus(ok24, 1,
         "打开工作簿即可正常显示5个扇区与百分比标签，图表已含缓存数据，无需手动重算或修复",
         "类别缓存=%s；数值缓存=%s" % (chart["cat_count"], chart["val_count"]))
    # =======================================================================
    # 以下为减分项：表格本身不属于本次作图范围，任何格式/内容改动都要扣分
    # =======================================================================
    blocks = []
    for i, top in enumerate(Q_ROWS):
        end = Q_ROWS[i + 1] - 1 if i + 1 < len(Q_ROWS) else DATA_LAST_ROW
        blocks.append((i, top, end))

    # ---------------- 25 新增了额外工作表 ----------------
    bad25 = len(wb.sheetnames) > 1
    minus(bad25, -3,
          '交付工作簿中不止Sheet1一个工作表，新增了"图表数据""Sheet2"等额外工作表',
          "工作表=%s" % wb.sheetnames)

    # ---------------- 26 A1 标题单元格格式被改 ----------------
    # 宋体、12 磅、加粗、白色文字、深青底纹 174A55、左对齐且垂直居中。
    a1 = ws["A1"]
    bad26 = not (_norm(a1.value) == _norm(TITLE_TEXT)
                 and _font_name_ok(a1, FONT_SONG) and _size_of(a1) == 12.0
                 and bool(a1.font.b) and _cell_font_rgb(a1) == CLR_WHITE
                 and _cell_fill_rgb(a1) == FILL_TITLE
                 and a1.alignment.horizontal == "left"
                 and a1.alignment.vertical == "center")
    minus(bad26, -3,
          '单元格A1"乡村公共服务调查汇总"不满足宋体、12磅、加粗、白色文字、'
          '深青底纹（174A55）、左对齐且垂直居中',
          "A1=%s/%s/%s/粗%s/字%s/底%s/%s+%s" % (
              _norm(a1.value), a1.font.name, _size_of(a1), bool(a1.font.b),
              _cell_font_rgb(a1), _cell_fill_rgb(a1),
              a1.alignment.horizontal, a1.alignment.vertical))

    # ---------------- 27 A2:D2 表头格式被改 ----------------
    # 宋体、10 磅、加粗、白色文字、2F6F73 底纹、水平垂直居中且自动换行。
    bad27_list = []
    for col in range(1, 5):
        cell = ws.cell(row=2, column=col)
        if not (_norm(cell.value) == _norm(HEADER_TEXTS[col - 1])
                and _font_name_ok(cell, FONT_SONG) and _size_of(cell) == 10.0
                and bool(cell.font.b) and _cell_font_rgb(cell) == CLR_WHITE
                and _cell_fill_rgb(cell) == FILL_HEADER
                and cell.alignment.horizontal == "center"
                and cell.alignment.vertical == "center"
                and bool(cell.alignment.wrap_text)):
            bad27_list.append(cell.coordinate)
    minus(bool(bad27_list), -3,
          '表头A2:D2"调查事项""回答选项""有效人数""比例"不满足宋体、10磅、加粗、白色文字、'
          '2F6F73底纹、水平垂直居中且自动换行',
          "不合格单元格=%s" % bad27_list)

    # ---------------- 28 A 列 20 个调查事项文字被增删改写或错行 ----------------
    a_rows = [r for r in range(DATA_FIRST_ROW, DATA_LAST_ROW + 1)
              if _norm(ws.cell(row=r, column=1).value)]
    bad28_list = [Q_ROWS[i] for i in range(len(Q_ROWS))
                  if _norm(ws.cell(row=Q_ROWS[i], column=1).value) != _norm(QUESTION_TEXTS[i])]
    bad28 = a_rows != Q_ROWS or bool(bad28_list)
    minus(bad28, -3,
          'A列20个调查事项文字（"您对村内便民服务事项的知晓程度"至'
          '"当前公共服务推进中的主要难点（可多选）"）出现增删、改写或错行',
          "文字所在行=%s；不符行=%s" % (a_rows if a_rows != Q_ROWS else "与20题一致", bad28_list))

    # ---------------- 29 B 列各回答选项文字被增删改写或错行 ----------------
    bad29_list = [r for r in range(DATA_FIRST_ROW, DATA_LAST_ROW + 1)
                  if _norm(ws.cell(row=r, column=2).value) != _norm(OPTION_TEXTS[r - DATA_FIRST_ROW])]
    minus(bool(bad29_list), -3,
          'B列各回答选项文字（如"十分清楚""从未关注""其他难点"）出现增删、改写或错行',
          "不符行=%s" % bad29_list[:12])

    # ---------------- 30 C 列有效人数数值/字体/对齐被改 ----------------
    bad30_list = []
    for r in range(DATA_FIRST_ROW, DATA_LAST_ROW + 1):
        cell = ws.cell(row=r, column=3)
        want = COUNTS[r - DATA_FIRST_ROW]
        if not (_to_float(cell.value) == float(want)
                and _font_name_ok(cell, FONT_TNR) and _size_of(cell) == 10.0
                and cell.alignment.horizontal == "center"
                and cell.alignment.vertical == "center"):
            bad30_list.append(r)
    minus(bool(bad30_list), -3,
          'C列"有效人数"数值（34、67、91、70、24……22）被改写或清空，'
          '或不为Times New Roman、10磅、水平垂直居中',
          "不符行=%s" % bad30_list[:12])

    # ---------------- 31 D 列比例数值或数字格式被改 ----------------
    # 比例应等于该行有效人数 / 286（第一题为 11.89%、23.43%、31.82%、24.48%、8.39%，
    # 末行"其他难点"为 7.69%），数字格式应为两位小数百分比 0.00%。
    bad31_list = []
    for r in range(DATA_FIRST_ROW, DATA_LAST_ROW + 1):
        cell = ws.cell(row=r, column=4)
        want = COUNTS[r - DATA_FIRST_ROW] / float(SAMPLE_TOTAL)
        got = _to_float(cell.value)
        if got is None or abs(got - want) > 5e-6 or cell.number_format != NUMFMT_D:
            bad31_list.append(r)
    minus(bool(bad31_list), -3,
          'D列"比例"出现数值被改写、清空或改成非百分比数值的情况（如第一个调查事项对应的'
          '5个比例不再为11.89%、23.43%、31.82%、24.48%、8.39%，或末行"其他难点"对应比例'
          '不再为7.69%，或被改成34、0.12、12这类与该行有效人数占286人总量不相符的数值），'
          '或数字格式不为两位小数百分比（0.00%）',
          "不符行=%s" % bad31_list[:12])
    # ---------------- 32 D 列比例文字或 C 列人数文字颜色被改 ----------------
    bad32_list = []
    for r in range(DATA_FIRST_ROW, DATA_LAST_ROW + 1):
        d = ws.cell(row=r, column=4)
        c = ws.cell(row=r, column=3)
        if not (_font_name_ok(d, FONT_TNR) and _size_of(d) == 10.0 and bool(d.font.b)
                and _cell_font_rgb(d) == CLR_D_FONT):
            bad32_list.append("D%d" % r)
        if _cell_font_rgb(c) != CLR_C_FONT:
            bad32_list.append("C%d" % r)
    minus(bool(bad32_list), -1,
          'D列比例文字不为Times New Roman、10磅、加粗、深棕色（6A4B13），'
          '或C列人数文字不为深蓝色（244B5A）',
          "不符单元格=%s" % bad32_list[:12])

    # ---------------- 33 D 列 / C 列 / B 列单元格底纹被改 ----------------
    # D 列浅黄 FFF1CC，C 列浅蓝 E8F1F7，B 列按题块在浅青 F3F8F7 与浅蓝 F5F8FB 之间交替。
    bad33_list = []
    for r in range(DATA_FIRST_ROW, DATA_LAST_ROW + 1):
        if _cell_fill_rgb(ws.cell(row=r, column=4)) != FILL_D:
            bad33_list.append("D%d" % r)
        if _cell_fill_rgb(ws.cell(row=r, column=3)) != FILL_C:
            bad33_list.append("C%d" % r)
    for i, top, end in blocks:
        want = FILL_B_ALT[i % 2]
        for r in range(top, end + 1):
            if _cell_fill_rgb(ws.cell(row=r, column=2)) != want:
                bad33_list.append("B%d" % r)
    minus(bool(bad33_list), -1,
          'D列比例单元格底纹不为浅黄（FFF1CC），或C列人数单元格底纹不为浅蓝（E8F1F7），'
          '或B列选项单元格按题块交替的浅青、浅蓝底纹（F3F8F7与F5F8FB）被改为统一颜色或被清除',
          "不符单元格=%s" % bad33_list[:12])

    # ---------------- 34 A 列问题单元格交替底纹被改 ----------------
    bad34_list = []
    for i, top, end in blocks:
        want = FILL_A_ALT[i % 2]
        if _cell_fill_rgb(ws.cell(row=top, column=1)) != want:
            bad34_list.append("A%d" % top)
    minus(bool(bad34_list), -1,
          'A列各问题单元格按题交替的底纹（D7ECE8与E1EDF3）被改为统一颜色或被清除',
          "不符单元格=%s" % bad34_list[:12])

    # ---------------- 35 表格框线被删除或改为其他线型 ----------------
    # 每个调查事项区块的上、下分界线为中粗实线（medium），区块内部框线为细实线（thin）。
    tops = set(top for _, top, _ in blocks)
    ends = set(end for _, _, end in blocks)
    bad35_list = []
    for r in range(DATA_FIRST_ROW, DATA_LAST_ROW + 1):
        for col in range(1, 5):
            border = ws.cell(row=r, column=col).border
            want_top = "medium" if (r in tops and col in (1, 2)) else "thin"
            want_bottom = "medium" if (r in ends and col in (1, 2)) else "thin"
            if (_bs(border.top) != want_top or _bs(border.bottom) != want_bottom
                    or _bs(border.left) != "thin" or _bs(border.right) != "thin"):
                bad35_list.append(ws.cell(row=r, column=col).coordinate)
    minus(bool(bad35_list), -1,
          "表格框线被删除或改为其他线型、其他颜色：每个调查事项区块的上、下分界线应为"
          "中粗实线，区块内部单元格框线应为细实线",
          "不符单元格=%s" % bad35_list[:12])

    # ---------------- 36 表格未完整占满第 1 至第 95 行 ----------------
    extra_rows = [r for r in range(DATA_LAST_ROW + 1, (ws.max_row or 0) + 1)
                  if any(_norm(ws.cell(row=r, column=col).value) for col in range(1, 5))]
    b_empty = [r for r in range(DATA_FIRST_ROW, DATA_LAST_ROW + 1)
               if not _norm(ws.cell(row=r, column=2).value)]
    bad36 = bool(extra_rows) or bool(b_empty) or a_rows != Q_ROWS
    minus(bad36, -3,
          "表格未完整占满第1至第95行（应为20个调查事项、93个回答选项行），出现新增或删除数据行",
          "第95行后仍有内容的行=%s；B列空行=%s" % (extra_rows[:6], b_empty[:6]))

    # ---------------- 37 四列列宽被改 ----------------
    bad37_list = []
    for letter in ("A", "B", "C", "D"):
        dim = ws.column_dimensions.get(letter)
        got = getattr(dim, "width", None) if dim is not None else None
        if got is None or abs(float(got) - EXPECT_WIDTHS[letter]) > 0.02:
            bad37_list.append("%s=%s" % (letter, got))
    minus(bool(bad37_list), -1,
          "A、B、C、D四列列宽不为34.125、8.375、7.625、6.875字符",
          "不符列宽=%s" % bad37_list)
    # ---------------- 38 B 列或 A 列文字对齐方式被改 ----------------
    bad38_list = []
    for r in range(DATA_FIRST_ROW, DATA_LAST_ROW + 1):
        al = ws.cell(row=r, column=2).alignment
        if al.horizontal not in (None, "left", "general") or al.vertical != "center":
            bad38_list.append("B%d" % r)
    for _, top, _ in blocks:
        al = ws.cell(row=top, column=1).alignment
        if al.horizontal != "center" or al.vertical != "center":
            bad38_list.append("A%d" % top)
    minus(bool(bad38_list), -1,
          "B列回答选项文字不为左对齐且垂直居中，或A列问题文字不为水平垂直居中",
          "不符单元格=%s" % bad38_list[:12])

    # ---------------- 39 数据区被新增合并单元格 ----------------
    merged = [str(rng) for rng in ws.merged_cells.ranges
              if rng.min_col <= 4 and rng.min_row <= DATA_LAST_ROW]
    minus(bool(merged), -1,
          "表格数据区单元格被新增合并（如将同一问题对应的A3:A7合并为一格）",
          "合并区域=%s" % merged[:8])

    # ---------------- 40 图表数量多于 1 个 ----------------
    minus(len(chart_parts) > 1, -3,
          "为其余19个调查事项也批量添加了图表，工作簿中图表数量多于1个",
          "图表部件数=%d：%s" % (len(chart_parts), chart_parts[:6]))

    # ---------------- 41 图表压住 A 至 D 四列的表格内容 ----------------
    # 任一图表锚点的左上角起始列落在 A~D（列索引 0~3）即判为遮挡表格内容。
    cover = [(one["part"], one["from"]) for one in chart_anchors
             if one["from"] is not None and one["from"][0] < 4]
    minus(bool(cover), -3,
          "图表位置覆盖或遮挡A至D四列的表格内容，导致数据被挡住无法阅读",
          "越界锚点=%s" % cover[:6])

    # ---------------- 42 单元格中出现额外的标注文字 ----------------
    hits42 = []
    for row in ws.iter_rows():
        for cell in row:
            text = _norm(cell.value) if isinstance(cell.value, str) else ""
            if text and any(k in text for k in ANNOT_KEYWORDS):
                hits42.append("%s=%s" % (cell.coordinate, text[:16]))
    minus(bool(hits42), -1,
          '单元格中出现了"图表说明""数据来源""占比统计"等额外的标注文字或说明行',
          "标注单元格=%s" % hits42[:6])

    wb.close()

    # ---------------- 汇总 ----------------
    # details 只用于本地排查与"每项恰好登记一次"的计数校验，不写入返回信息：
    # 返回的 dim2_items 每条只含 rule / max_delta / delta / hit 四个键。
    assert len(results) == len(details) == len(_MAX_DELTAS_ORDER)
    total = 0
    items = []
    for raw, max_delta in zip(results, _MAX_DELTAS_ORDER):
        head, _, rule = raw.partition("：")
        try:
            mag = int(head.strip().lstrip("+-"))
        except ValueError:
            mag = 0
        hit = mag > 0
        delta = max_delta if hit else 0
        total += delta
        items.append({
            "rule": rule,
            "max_delta": max_delta,
            "delta": delta,
            "hit": hit,
        })
    result["dim2_items"] = items
    result["total_score"] = max(0, total)   # 扣分多于加分时兜底为 0
    result["max_score"] = MAX_SCORE_TOTAL
    return result


if __name__ == "__main__":
    _dir = sys.argv[1] if len(sys.argv) >= 2 else os.path.dirname(os.path.abspath(__file__))
    print(json.dumps(evaluate(_dir), ensure_ascii=False, indent=2))
