# -*- coding: utf-8 -*-
"""
officeval_078_verifier —「区域协同汇报.pptx 第8页 Excel 超链接」任务完成度自动评测

指令：麻烦将“CRM_pipeline_restyled_reordered”文件插到“区域协同汇报”文件第8页右边的空白处做一个超链接

评测口径与 078/细则.txt 严格一一对应：
  维度1  可用与可修改性：交付文件为 1 个可正常打开的 .pptx；损坏无法打开则总分记 0
  维度2  完成度评分细则：7 条加分项（合计 21 分）+ 13 条减分项（命中即扣）

设计约束（与 officeval_008_verifier 保持一致）：
  * evaluate(dir_path) 从目录里自行定位交付文件，不接受文件路径参数；
  * 全程不 print、不写 sys.stdout、不 sys.exit，任何异常都收敛成 status='error'；
  * results 列表里每条形如 "+N：说明" / "+0：说明" / "-N：说明" / "-0：说明"，
    顺序必须与 _MAX_DELTAS_ORDER 严格对齐，最后统一反解成 dim2_items；
  * 只踩细则明确写出的点，不额外加码；读数口径与办公软件（PowerPoint/Excel）对齐。
"""

import os
import re
import sys
import json
import math
import zipfile
import traceback
import urllib.parse

SCRIPT_ID = "078"

# 加分项满分合计（减分项不计入满分，只从实得分里扣）
MAX_SCORE_TOTAL = 21

# 与 _evaluate_impl 中 results 的追加顺序严格一一对应：
# 前 7 项为加分项（+5/+3/+3/+5/+3/+1/+1），后 13 项为减分项（负值 = 命中即扣）
_MAX_DELTAS_ORDER = [
    5, 3, 3, 5, 3, 1, 1,
    -3, -3, -1, -1, -3, -3, -3, -3, -1, -3, -1, -3, -1,
]
assert sum(d for d in _MAX_DELTAS_ORDER if d > 0) == MAX_SCORE_TOTAL

# ---------------------------------------------------------------------------
# 常量：交付物、命名空间、几何容差
# ---------------------------------------------------------------------------
EXP_PPTX_NAME = "区域协同汇报.pptx"          # 细则：交付演示文稿文件名须保持不变
EXP_XLSX_NAME = "CRM_pipeline_restyled_reordered.xlsx"  # 细则：被链接的工作簿

_NS = {
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
}
_A = "{%s}" % _NS["a"]
_R = "{%s}" % _NS["r"]

EMU_PER_CM = 360000.0

TOL_POS = 0.08      # 位置/尺寸容差（cm）：容纳 PowerPoint 重存时的 EMU 取整漂移
TOL_PAGE = 0.10     # 页面尺寸容差（cm）
TOL_SZ = 0.20       # 字号容差（磅）：8.69 磅经 PowerPoint 重存会写成 8.70

# 页面：33.87cm × 19.05cm（16:9 宽屏）
PAGE_W_CM = 33.87
PAGE_H_CM = 19.05

# 第8页表格版式基线（左1.65 上4.37 宽24.0 高12.57；8 行行高 1.57；6 列列宽如下）
TBL_BOX = (1.651, 4.369, 24.003, 12.573)
TBL_ROW_H = 1.572
TBL_COL_W = (3.683, 3.937, 5.080, 3.175, 5.588, 2.540)
TBL_FILL_HEAD = "184669"
# 数据行单元格填充为 10304E 与 123758 隔行交替（第2/4/6/8行深、第3/5/7行浅）
TBL_FILL_BODY = ("10304E", "123758")
TBL_HEAD_FONT = ("Aptos", 8.69, True, "F4F9FB")   # 字体/字号/加粗/颜色

# 第8页表格文字内容基线（6 列 × 8 行）
TBL_TEXT = (
    ("Pipeline", "Sector", "Scope", "Stage", "Q2 action", "Value band"),
    ("HarborLink", "Retail & consumer", "Ocean + customs", "Proposal",
     "Validate lane cost", "Medium"),
    ("AeroPulse", "Aerospace parts", "Time-critical air", "Qualified",
     "Run service trial", "Medium"),
    ("GreenCart", "Food logistics", "Cold-chain routing", "Discovery",
     "Confirm compliance path", "Low"),
    ("AtlasWorks", "Industrial projects", "Oversized cargo", "Negotiation",
     "Finalize risk sheet", "High"),
    ("NovaCare", "Healthcare supply", "Regional replenishment", "Qualified",
     "Pilot two lanes", "Medium"),
    ("BlueGrid", "Energy equipment", "Project cargo", "Standby",
     "Recheck funding window", "High"),
    ("NorthBay", "E-commerce", "Cross-border parcel", "Discovery",
     "Build rate card", "Low"),
)

# 第8页主标题与辅助文字基线：文本 -> (左cm, 上cm, 字体, 字号, 加粗, 颜色)
TXT_TITLE = ("Strategic Pipeline Status Update", 1.397, 1.473, "Aptos", 26.0, True, "F4F9FB")
TXT_AUX = (
    ("07", 1.397, 0.864, "Aptos", 12.0, True, "F6B743"),
    ("Q1 opportunity status is organized by sector, scope, stage and Q2 action.",
     1.422, 2.616, "Aptos", 12.0, False, "B2C8D4"),
    ("Pipeline view: opportunities, next actions and priority bands.",
     1.778, 17.196, "Aptos", 10.0, False, "B2C8D4"),
    ("Regional Operating Review  |  FY2026 Q1", 1.397, 18.085, "Aptos", 8.0, False, "7D96A6"),
    ("08/10", 30.988, 18.085, "Aptos", 8.0, False, "7D96A6"),
)
PAGE_NO_TEXT = "08/10"          # 页码文本，同时用于定位第8页
PAGE_NO_LEFT = 30.988           # 页码右对齐置于 30.99cm 处

# 第8页装饰元素基线：(标识, 形状类别, 填充色, 左, 上, 宽, 高)
#   形状类别 rect/oval/pic —— rect 与 oval 用于区分同为 00C2CC 的短线与左上圆
DECOS = (
    ("深色背景矩形081226", "rect", "081226", 0.0, 0.0, 33.867, 19.05),
    ("左上青色圆00C2CC", "oval", "00C2CC", -3.048, -1.778, 10.160, 10.160),
    ("右上紫色圆7A6BFF", "oval", "7A6BFF", 22.606, -2.794, 6.604, 6.604),
    ("右下橙色圆F6B743", "oval", "F6B743", 25.400, 14.732, 11.430, 7.112),
    ("标题下方青色短线00C2CC", "rect", "00C2CC", 1.397, 3.429, 5.080, 0.102),
    ("底部横线2C5970", "rect", "2C5970", 1.397, 17.907, 30.734, 0.025),
)
DECO_PIC = (0.0, 0.0, 33.867, 19.05)   # 满版背景图片

# 表格右侧空白区（加分项 +5 的落位窗口）
BLANK_L, BLANK_R = 25.70, 32.50
BLANK_T, BLANK_B = 4.00, 14.50
RIGHT_MARGIN_LINE = 32.131        # 版心右侧界线：底部横线右端 1.397+30.734
CARRIER_W_RANGE = (1.5, 7.0)      # 超链接对象宽度合理区间（cm）
CARRIER_H_RANGE = (0.6, 5.5)      # 超链接对象高度合理区间（cm）
OLE_BLOAT_BYTES = 100000          # 内嵌工作簿导致的体积膨胀阈值（字节）

# 其余 9 页基线：页序 -> (形状数, 必须出现的文本)
OTHER_SLIDES = {
    1: (19, ("REGIONAL COORDINATION", "Q1 2026 Operating Review")),
    2: (34, ("01", "Financial Snapshot")),
    3: (23, ("02", "Performance Scorecard")),
    4: (25, ("03", "Diagnosis: Progress")),
    5: (32, ("04", "Diagnosis: Roadblocks")),
    6: (36, ("05", "Actions and Next Steps")),
    7: (16, ("06", "Leadership Alignment Requests")),
    9: (17, ("08", "Agreements and Request Table")),
    10: (30, ("09", "Rolling Forecast")),
}
TOTAL_SLIDES = 10
TARGET_INDEX = 8                  # 第8页（页脚页码 08/10）

# 工作簿基线：表名 -> (行数, 列数, 第1行标题前缀, 第2行记录数, 第3行首个字段名)
WB_SHEETS = (
    ("Strategic Pipeline", 160, 7, "Strategic Pipeline", 157, "Sales Talent"),
    ("Lead Development", 220, 24, "Lead Development", 217, "Lead ID"),
    ("Client Portfolio", 480, 43, "Client Portfolio", 477, "EP4WARD Id Customer"),
)

# ---------------------------------------------------------------------------
# 基础工具
# ---------------------------------------------------------------------------
def _cm(emu):
    """EMU -> cm；None 视为 0。"""
    try:
        return float(emu) / EMU_PER_CM
    except (TypeError, ValueError):
        return 0.0


def _near(a, b, tol=TOL_POS):
    return abs(float(a) - float(b)) <= tol


def _norm(s):
    """文本归一化：统一空白、去首尾，便于与基线逐字比对。"""
    if s is None:
        return ""
    s = s.replace(" ", " ").replace("\r", "\n").replace("\x0b", "\n")
    s = re.sub(r"[ \t]+", " ", s)
    return "\n".join(line.strip() for line in s.split("\n")).strip()


def _flat(s):
    """再把换行压成空格——PowerPoint 会把软回车/自动换行写成不同结构。"""
    return re.sub(r"\s+", " ", _norm(s)).strip()


def _box(sh):
    """返回形状外框 (左, 上, 宽, 高)，单位 cm。"""
    return (_cm(sh.left), _cm(sh.top), _cm(sh.width), _cm(sh.height))


def _rect(sh):
    """返回形状外框 (左, 上, 右, 下)，单位 cm。"""
    l, t, w, h = _box(sh)
    return (l, t, l + w, t + h)


def _overlap_area(r1, r2):
    """两矩形交叠面积（cm²）。"""
    dx = min(r1[2], r2[2]) - max(r1[0], r2[0])
    dy = min(r1[3], r2[3]) - max(r1[1], r2[1])
    if dx <= 0 or dy <= 0:
        return 0.0
    return dx * dy


def _area(r):
    return max(0.0, r[2] - r[0]) * max(0.0, r[3] - r[1])


def _rpr_font(rpr):
    """从 a:rPr / a:defRPr / a:endParaRPr 上读字体属性。"""
    out = {}
    if rpr is None:
        return out
    latin = rpr.find("a:latin", _NS)
    if latin is not None and latin.get("typeface"):
        out["name"] = latin.get("typeface")
    if rpr.get("sz"):
        try:
            out["size"] = int(rpr.get("sz")) / 100.0
        except ValueError:
            pass
    if rpr.get("b") is not None:
        out["bold"] = rpr.get("b") in ("1", "true", "on")
    clr = rpr.find("a:solidFill/a:srgbClr", _NS)
    if clr is not None and clr.get("val"):
        out["color"] = clr.get("val").upper()
    return out


def _para_font(p_elem):
    """
    段落有效字体：优先取首个 a:r/a:rPr，缺项回落到 a:pPr/a:defRPr，再回落 a:endParaRPr。
    —— 本套演示文稿的字体全写在 defRPr 上，只看 run 会全读成 None，故必须逐级回落。
    """
    chain = [r.find("a:rPr", _NS) for r in p_elem.findall("a:r", _NS)]
    chain.append(p_elem.find("a:pPr/a:defRPr", _NS))
    chain.append(p_elem.find("a:endParaRPr", _NS))
    out = {}
    for e in chain:
        for k, v in _rpr_font(e).items():
            out.setdefault(k, v)
    return out


def _para_algn(p_elem):
    pPr = p_elem.find("a:pPr", _NS)
    return pPr.get("algn") if pPr is not None else None


def _font_ok(font, name, size, bold, color):
    """字体四要素比对：字号带 TOL_SZ 容差（8.69 磅重存后为 8.70）。"""
    if name is not None and font.get("name") != name:
        return False
    if size is not None and not _near(font.get("size", -1), size, TOL_SZ):
        return False
    if bold is not None and bool(font.get("bold")) != bool(bold):
        return False
    if color is not None and font.get("color") != color.upper():
        return False
    return True


def _shape_text(sh):
    try:
        if sh.has_text_frame:
            return _norm(sh.text_frame.text)
    except Exception:
        pass
    return ""


def _shape_fill_rgb(sh):
    """纯色填充色（大写 6 位十六进制），非纯色返回 None。"""
    try:
        el = sh._element.find("p:spPr/a:solidFill/a:srgbClr", _NS)
        if el is not None and el.get("val"):
            return el.get("val").upper()
    except Exception:
        pass
    return None


def _shape_kind(sh):
    """
    粗分类：pic / tbl / ole / rect / oval / rrect / txt / other。
    装饰元素里存在同色的圆与矩形，必须靠几何类型区分。
    """
    try:
        if sh.shape_type is not None and "PICTURE" in str(sh.shape_type):
            return "pic"
    except Exception:
        pass
    try:
        if sh.has_table:
            return "tbl"
    except Exception:
        pass
    if sh._element.find(".//p:oleObj", _NS) is not None:
        return "ole"
    geom = sh._element.find("p:spPr/a:prstGeom", _NS)
    prst = geom.get("prst") if geom is not None else None
    if prst == "ellipse":
        return "oval"
    if prst == "roundRect":
        return "rrect"
    if prst == "rect":
        # 文本框（cNvSpPr txBox="1"）与纯色矩形块分开
        cnv = sh._element.find("p:nvSpPr/p:cNvSpPr", _NS)
        if cnv is not None and cnv.get("txBox") == "1":
            return "txt"
        return "rect"
    return "other"


def _shape_hlinks(sh, part):
    """
    收集形状上（含其文字 run 上）的所有 a:hlinkClick，并解析到关系目标。
    返回 [{'rid','action','target','external','reltype'}]
    """
    out = []
    for el in sh._element.iter(_A + "hlinkClick"):
        rid = el.get(_R + "id")
        item = {"rid": rid or "", "action": el.get("action") or "",
                "target": "", "external": False, "reltype": ""}
        if rid:
            try:
                rel = part.rels[rid]
                item["external"] = bool(rel.is_external)
                item["reltype"] = str(rel.reltype)
                item["target"] = (rel.target_ref if rel.is_external
                                  else str(rel.target_partname))
            except Exception:
                pass
        out.append(item)
    return out


def _target_basename(target):
    """从关系目标里取出文件名（做一次 URL 反转义，兼容 file:/// 与 %20）。"""
    t = urllib.parse.unquote(target or "")
    t = t.split("#", 1)[0].split("?", 1)[0]
    t = t.replace("\\", "/").rstrip("/")
    return t.rsplit("/", 1)[-1]


def _slide_has_pkg_rel(part):
    """幻灯片是否挂着 package/oleObject 关系（即内嵌了整份工作簿）。"""
    for rel in part.rels.values():
        rt = str(rel.reltype).rsplit("/", 1)[-1].lower()
        if rt in ("package", "oleobject"):
            return True
    return False


def _text_fits(sh):
    """
    估算文字是否溢出对象边界：按字号折算单字宽与行高，逐行累计所需高度。
    只用于“文字不溢出对象边界”这一条，阈值放宽 15%，避免误判。
    """
    txt = _shape_text(sh)
    if not txt:
        return True
    try:
        font = _para_font(sh.text_frame.paragraphs[0]._p)
    except Exception:
        font = {}
    size = font.get("size") or 18.0
    ch_w = size * 0.55 / 28.3465
    line_h = size * 1.25 / 28.3465
    l, t, w, h = _box(sh)
    inner_w, inner_h = max(0.1, w - 0.5), max(0.1, h - 0.1)
    lines = sum(max(1, int(math.ceil(len(seg) * ch_w / inner_w)))
                for seg in txt.split("\n"))
    return lines * line_h <= inner_h * 1.15


# ---------------------------------------------------------------------------
# 第8页元素归位：把交付页上的形状与基线元素一一对号，剩下的即“新增对象”
# ---------------------------------------------------------------------------
def _match_slide8(shapes):
    """
    先按“身份锚点”（文字内容 / 是否表格 / 图形类型+填充色）认领基线元素；
    认领不到时再按位置就近认领（±0.6cm）——这样元素被改文字或改颜色时
    仍能被认出是“原有元素被改动”，而不会误判成“新增了多余元素”。
    """
    left = list(shapes)
    got = {"title": None, "aux": [None] * len(TXT_AUX), "table": None,
           "pic": None, "decos": {d[0]: None for d in DECOS}}

    def take(pred):
        for sh in left:
            try:
                if pred(sh):
                    left.remove(sh)
                    return sh
            except Exception:
                continue
        return None

    got["table"] = take(lambda s: _shape_kind(s) == "tbl")
    got["pic"] = take(lambda s: _shape_kind(s) == "pic")

    # 主标题 / 辅助文字：先认文字，再认位置
    got["title"] = take(lambda s: _flat(_shape_text(s)) == _flat(TXT_TITLE[0]))
    for i, spec in enumerate(TXT_AUX):
        got["aux"][i] = take(lambda s, t=spec[0]: _flat(_shape_text(s)) == _flat(t))
    if got["title"] is None:
        got["title"] = take(lambda s: _shape_kind(s) == "txt"
                            and _near(_cm(s.left), TXT_TITLE[1], 0.6)
                            and _near(_cm(s.top), TXT_TITLE[2], 0.6))
    for i, spec in enumerate(TXT_AUX):
        if got["aux"][i] is None:
            got["aux"][i] = take(lambda s, sp=spec: _shape_kind(s) == "txt"
                                 and _near(_cm(s.left), sp[1], 0.6)
                                 and _near(_cm(s.top), sp[2], 0.6))

    # 装饰形状：先认“图形类型+填充色”，再认“图形类型+位置”
    for name, kind, fill, l, t, w, h in DECOS:
        got["decos"][name] = take(lambda s, k=kind, f=fill:
                                  _shape_kind(s) == k and _shape_fill_rgb(s) == f)
    for name, kind, fill, l, t, w, h in DECOS:
        if got["decos"][name] is None:
            got["decos"][name] = take(lambda s, k=kind, x=l, y=t:
                                      _shape_kind(s) in (k, "rrect", "other")
                                      and _near(_cm(s.left), x, 0.6)
                                      and _near(_cm(s.top), y, 0.6))
    return got, left


# ---------------------------------------------------------------------------
# 对外入口
# ---------------------------------------------------------------------------
class _TargetMissing(Exception):
    """目录不存在或目录中没有 .pptx：连待评测对象都没有，不进入维度判定。"""


def evaluate(dir_path):
    """在 dir_path 内定位交付的 .pptx 并评分；异常一律收敛为 status='error'。"""
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
        if not dir_path or not os.path.isdir(dir_path):
            raise _TargetMissing(f"目录不存在：{dir_path}")
        cands = sorted(f for f in os.listdir(dir_path)
                       if f.lower().endswith(".pptx") and not f.startswith("~$"))
        if not cands:
            raise _TargetMissing(f"目录中未找到 .pptx 文件：{dir_path}")
        # 文件名保持为“区域协同汇报.pptx”的优先作为交付件，避免多份时取错
        cands.sort(key=lambda f: (f != EXP_PPTX_NAME, f))
        file_path = os.path.join(dir_path, cands[0])
        result["file_name"] = os.path.basename(file_path)
        return _evaluate_impl(file_path, dir_path, result)
    except _TargetMissing as exc:
        # 找不到待评测文件：维度1 未参与判定，dim1_pass / dim1_reason 一律留空，满分也置 0
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
        # 文件存在但打不开 / 解析失败：维度1 判不通过，dim1_reason 给出错误摘要
        result["status"] = "error"
        result["error"] = f"{type(exc).__name__}: {exc}"
        result["dim1_pass"] = False
        result["dim1_reason"] = f"文件无法解析或无法打开：{type(exc).__name__}: {exc}"
        result["dim2_items"] = []
        result["total_score"] = 0
        return result


def _find_xlsx(dir_path):
    """在交付目录下找被链接的工作簿：返回 (同目录路径, 子目录路径, 疑似改名路径)。"""
    same, sub, renamed = None, None, None
    for f in sorted(os.listdir(dir_path)):
        p = os.path.join(dir_path, f)
        if os.path.isfile(p) and f.lower().endswith(".xlsx") and not f.startswith("~$"):
            if f.lower() == EXP_XLSX_NAME.lower():
                same = same or p
            else:
                renamed = renamed or p
    for root, _dirs, files in os.walk(dir_path):
        if os.path.abspath(root) == os.path.abspath(dir_path):
            continue
        for f in files:
            if f.lower() == EXP_XLSX_NAME.lower():
                sub = sub or os.path.join(root, f)
    return same, sub, renamed


def _stray_files(dir_path, pptx_name):
    """
    交付目录中与任务无关的残留物：脚本、截图、日志、说明文档、Office 临时锁文件等。
    评测脚本自身与本套题的指令/细则文本不计入。
    """
    keep = {pptx_name.lower(), EXP_XLSX_NAME.lower(),
            "指令.txt", "细则.txt", os.path.basename(__file__).lower()}
    strays = []
    for f in sorted(os.listdir(dir_path)):
        p = os.path.join(dir_path, f)
        if not os.path.isfile(p):
            continue
        low = f.lower()
        if low in keep:
            continue
        if re.match(r"^officeval_\d+_verifier\.py$", low):
            continue
        strays.append(f)
    return strays


def _wb_check(path):
    """校验被链接工作簿的三张工作表结构；无法读取时返回 None（不据此扣分）。"""
    try:
        import openpyxl
    except Exception:
        return None
    try:
        wb = openpyxl.load_workbook(path, data_only=True)
    except Exception:
        return ["工作簿无法打开"]
    bad = []
    try:
        if [ws.title for ws in wb.worksheets] != [s[0] for s in WB_SHEETS]:
            bad.append("工作表名称或顺序不符（应为 %s）" % "/".join(s[0] for s in WB_SHEETS))
        else:
            for title, rows, cols, t1, cnt, h1 in WB_SHEETS:
                ws = wb[title]
                if ws.max_row != rows or ws.max_column != cols:
                    bad.append("%s 应为 %d行%d列，实为 %s行%s列"
                               % (title, rows, cols, ws.max_row, ws.max_column))
                if not str(ws.cell(1, 1).value or "").startswith(t1):
                    bad.append("%s 第1行标题行被改动" % title)
                if str(cnt) not in str(ws.cell(2, 1).value or ""):
                    bad.append("%s 第2行记录数说明行被改动" % title)
                if _flat(ws.cell(3, 1).value) != h1:
                    bad.append("%s 第3行字段表头行被改动" % title)
    finally:
        try:
            wb.close()
        except Exception:
            pass
    return bad


# ---------------------------------------------------------------------------
# 主体评测
# ---------------------------------------------------------------------------
def _evaluate_impl(file_path, dir_path, result):
    from pptx import Presentation

    prs = Presentation(file_path)
    slides = list(prs.slides)
    page_w, page_h = _cm(prs.slide_width), _cm(prs.slide_height)

    # --- 定位第8页：以页脚页码文本 08/10 为准，找不到再退回第 8 张 ---
    tgt_idx, target = None, None
    for i, sl in enumerate(slides, 1):
        for sh in sl.shapes:
            if _flat(_shape_text(sh)) == PAGE_NO_TEXT:
                tgt_idx, target = i, sl
                break
        if target is not None:
            break
    if target is None and len(slides) >= TARGET_INDEX:
        tgt_idx, target = TARGET_INDEX, slides[TARGET_INDEX - 1]

    shapes = list(target.shapes) if target is not None else []
    got, added = _match_slide8(shapes)

    # --- 超链接盘点：外部 .xlsx 链接及其载体 ---
    links = []
    for sh in shapes:
        for hl in _shape_hlinks(sh, target.part):
            hl["shape"] = sh
            hl["base"] = _target_basename(hl["target"])
            links.append(hl)
    xlsx_links = [h for h in links
                  if h["external"] and h["base"].lower().endswith(".xlsx")]
    carrier = xlsx_links[0]["shape"] if xlsx_links else None
    link_target = xlsx_links[0]["target"] if xlsx_links else ""
    link_base = xlsx_links[0]["base"] if xlsx_links else ""

    # --- OLE 内嵌盘点 ---
    ole_shapes = [sh for sh in shapes if _shape_kind(sh) == "ole"]
    has_pkg_rel = _slide_has_pkg_rel(target.part) if target is not None else False
    emb_bytes = 0
    with zipfile.ZipFile(file_path) as z:
        for n in z.namelist():
            if n.startswith("ppt/embeddings/"):
                emb_bytes += z.getinfo(n).file_size
    added_tables = [sh for sh in added if _shape_kind(sh) == "tbl"]

    # --- 随附工作簿 ---
    wb_same, wb_sub, wb_renamed = _find_xlsx(dir_path)

    # =======================================================================
    # 维度1：可用与可修改性
    #   仅判一件事：交付文件为 1 个可正常打开的 PowerPoint 演示文稿（.pptx）。
    #   文件损坏/非 pptx 时 Presentation() 会抛错，由 evaluate 的兜底置为 status=error
    #   并记 0 分；能走到这里即视为维度1 通过。
    #   （页数、页面尺寸、元素可单独编辑、版面是否错乱等均不在维度1 判定，
    #     相应问题由维度2 各条减分项体现。）
    # =======================================================================
    result["dim1_pass"] = True
    result["dim1_reason"] = "交付文件为 1 个可正常打开的 PowerPoint 演示文稿（.pptx）"

    results = []

    # =======================================================================
    # 加分项 1（+5）：第8页新增了指向 .xlsx 的可点击外部文件超链接对象
    #   —— 必须是 TargetMode=External 的外部文件链接；跳转到内部页面（reltype=slide）
    #      或网页地址（http/https/mailto）都不算；纯文字/纯图片无链接也不算。
    # =======================================================================
    inner_only = [h for h in links if not h["external"]
                  and str(h["reltype"]).rsplit("/", 1)[-1].lower() == "slide"]
    web_only = [h for h in links if h["external"]
                and re.match(r"(?i)^(https?:|mailto:)", h["target"] or "")]
    if carrier is not None:
        results.append("+5：第8页新增超链接对象“%s”，挂 TargetMode=External 的外部文件超链接"
                       "（%s，action=%s），点击可调起表格软件打开工作簿"
                       % (carrier.name, link_target, xlsx_links[0]["action"] or "-"))
    elif inner_only or web_only:
        results.append("+0：第8页新增对象上的超链接指向%s，不是指向 .xlsx 工作簿的外部文件链接"
                       % ("演示文稿内部页面" if inner_only else "网页地址"))
    else:
        results.append("+0：第8页未发现指向 .xlsx 工作簿文件的可点击外部超链接对象"
                       "（新增对象 %d 个，均无外部文件链接）" % len(added))

    # =======================================================================
    # 加分项 2（+3）：超链接目标文件名与随附工作簿完全一致，未指向副本/临时目录
    #   —— 只判目标写法（文件名逐字一致、无“(2)”“副本”“copy”、不落在临时目录），
    #      工作簿是否随交付提供由下一条单独判定，避免同一事实重复计分。
    # =======================================================================
    _t = urllib.parse.unquote(link_target or "")
    copy_mark = re.search(r"(?i)\(\d+\)|副本|\bcopy\b|- ?副本", _t)
    temp_mark = re.search(r"(?i)temp|tmp|appdata|缓存|recent", _t)
    if carrier is not None and link_base.lower() == EXP_XLSX_NAME.lower() \
            and not copy_mark and not temp_mark:
        results.append("+3：超链接目标为“%s”，与随附工作簿 %s 文件名完全一致，"
                       "为可迁移的相对路径写法" % (_t, EXP_XLSX_NAME))
    elif carrier is None:
        results.append("+0：无外部工作簿超链接，目标文件名无从谈起")
    else:
        why = ("含副本编号/“副本”字样" if copy_mark else
               "指向临时目录" if temp_mark else "文件名与 %s 不一致" % EXP_XLSX_NAME)
        results.append("+0：超链接目标“%s”%s，点击会提示无法找到指定的文件" % (_t, why))

    # =======================================================================
    # 加分项 3（+3）：随交付一并提供被链接的工作簿，且与演示文稿位于同一目录
    # =======================================================================
    if wb_same:
        results.append("+3：交付目录内与演示文稿同级提供了 %s，相对路径超链接换机后仍可点开"
                       % EXP_XLSX_NAME)
    else:
        results.append("+0：交付目录内未与演示文稿同级提供 %s（%s），"
                       "相对路径超链接换机后点不开"
                       % (EXP_XLSX_NAME,
                          "仅见于子目录 %s" % os.path.relpath(wb_sub, dir_path) if wb_sub
                          else "目录内无该工作簿"))

    # =======================================================================
    # 加分项 4（+5）：超链接对象完整落在第8页表格右侧空白区
    #   窗口：左≥25.70cm、右≤32.50cm、上≥4.00cm、下≤14.50cm
    #   —— 同时排除了顶部标题区(0-3.5)、说明行(17.2起)、页脚行(18.0起)
    # =======================================================================
    if carrier is not None:
        cl, ct, cr, cb = _rect(carrier)
        ok_box = (cl >= BLANK_L - 0.02 and cr <= BLANK_R + 0.02
                  and ct >= BLANK_T - 0.02 and cb <= BLANK_B + 0.02)
        if ok_box:
            results.append("+5：超链接对象落位 左%.2f/上%.2f/右%.2f/下%.2f cm，"
                           "完整落在表格右侧空白区（左≥%.1f 右≤%.1f 上≥%.1f 下≤%.1f）"
                           % (cl, ct, cr, cb, BLANK_L, BLANK_R, BLANK_T, BLANK_B))
        else:
            results.append("+0：超链接对象落位 左%.2f/上%.2f/右%.2f/下%.2f cm，"
                           "越出表格右侧空白区窗口（左≥%.1f 右≤%.1f 上≥%.1f 下≤%.1f）"
                           % (cl, ct, cr, cb, BLANK_L, BLANK_R, BLANK_T, BLANK_B))
    else:
        results.append("+0：无超链接对象，无法判定是否落在表格右侧空白区")

    # =======================================================================
    # 加分项 5（+3）：以超链接方式实现，未把整份工作簿塞进页面
    #   —— 页面上不得有 OLE 对象、不得挂 package/oleObject 关系、不得新增整张数据表
    # =======================================================================
    if not ole_shapes and not has_pkg_rel and not added_tables:
        results.append("+3：第8页以超链接方式引用工作簿，页面内无内嵌 OLE 对象、"
                       "无 package 关系、未新增整表粘贴")
    else:
        why = []
        if ole_shapes:
            why.append("存在 %d 个内嵌 OLE 对象" % len(ole_shapes))
        if has_pkg_rel:
            why.append("幻灯片挂有 package/oleObject 关系（ppt/embeddings 合计 %d 字节）"
                       % emb_bytes)
        if added_tables:
            why.append("新增了 %d 张粘贴表格" % len(added_tables))
        results.append("+0：未以纯超链接方式实现——%s" % "；".join(why))

    # =======================================================================
    # 加分项 6（+1）：超链接对象带有可读提示文字，不是无文字的透明块/隐形热区
    # =======================================================================
    label = _flat(_shape_text(carrier)) if carrier is not None else ""
    if not label and carrier is not None:
        for sh in added:                     # 载体本身无字时，看同批新增对象上的标签
            if len(_flat(_shape_text(sh))) >= 3:
                label = _flat(_shape_text(sh))
                break
    if carrier is not None and len(label) >= 3:
        results.append("+1：超链接对象带可读提示文字“%s”" % label)
    elif carrier is None:
        results.append("+0：无超链接对象，无提示文字可言")
    else:
        results.append("+0：超链接对象无可读提示文字，是看不见的透明块/热区")

    # =======================================================================
    # 加分项 7（+1）：超链接对象尺寸适中、完整在可见范围内、文字不溢出、且不与已有元素压盖
    #   宽 1.5-7.0cm、高 0.6-5.5cm；压盖比对对象为表格、三个装饰圆、底部横线、
    #   标题下方短线以及各文字块（满版背景矩形与满版背景图片天然全覆盖，不参与比对）
    # =======================================================================
    if carrier is not None:
        cl, ct, cr, cb = _rect(carrier)
        cw, ch = cr - cl, cb - ct
        size_ok = (CARRIER_W_RANGE[0] <= cw <= CARRIER_W_RANGE[1]
                   and CARRIER_H_RANGE[0] <= ch <= CARRIER_H_RANGE[1])
        inside = (cl >= -0.02 and ct >= -0.02
                  and cr <= page_w + 0.02 and cb <= page_h + 0.02)
        fits = _text_fits(carrier)
        peers = [got["table"], got["title"]] + list(got["aux"]) + [
            got["decos"][n] for n, k, f, l, t, w, h in DECOS
            if n != "深色背景矩形081226"]
        clash = []
        for sh in peers:
            if sh is None or sh is carrier:
                continue
            if _overlap_area(_rect(sh), (cl, ct, cr, cb)) > 0.01 * max(0.01, cw * ch):
                clash.append(sh.name)
        if size_ok and inside and fits and not clash:
            results.append("+1：超链接对象宽%.2fcm、高%.2fcm 尺寸适中，完整位于幻灯片可见范围内，"
                           "文字未溢出边界，且与表格、装饰圆、底部横线等已有元素无压盖"
                           % (cw, ch))
        else:
            why = []
            if not size_ok:
                why.append("尺寸 宽%.2f×高%.2fcm 超出 宽%.1f-%.1f、高%.1f-%.1f 区间"
                           % (cw, ch, CARRIER_W_RANGE[0], CARRIER_W_RANGE[1],
                              CARRIER_H_RANGE[0], CARRIER_H_RANGE[1]))
            if not inside:
                why.append("超出幻灯片可见范围")
            if not fits:
                why.append("文字溢出对象边界")
            if clash:
                why.append("与已有元素压盖：%s" % "、".join(dict.fromkeys(clash)))
            results.append("+0：超链接对象%s" % "；".join(why))
    else:
        results.append("+0：无超链接对象，尺寸与压盖情况无从判定")

    # =======================================================================
    # 减分项 1（-3）：第8页除超链接载体本身外，还新增了指令未要求的其他图形/图片/文本框
    # =======================================================================
    extra = [sh for sh in added if sh is not carrier]
    if extra:
        results.append("-3：第8页除超链接载体外多出 %d 类新增元素：%s"
                       % (len(extra), "、".join("%s(%s)" % (sh.name, _shape_kind(sh))
                                                for sh in extra)))
    else:
        results.append("-0：第8页仅新增超链接载体 1 个对象，无多余图形/图片/文本框")

    # =======================================================================
    # 减分项 2（-3）：第8页嵌入了整份工作簿的 OLE 对象，且体积膨胀 10 万字节以上
    # =======================================================================
    if (ole_shapes or has_pkg_rel) and emb_bytes >= OLE_BLOAT_BYTES:
        results.append("-3：第8页内嵌了整份工作簿的 OLE 对象（ppt/embeddings 合计 %d 字节，"
                       "≥%d 字节），演示文稿内多出一份会与源工作簿失去同步的数据副本"
                       % (emb_bytes, OLE_BLOAT_BYTES))
    else:
        results.append("-0：演示文稿内未内嵌工作簿 OLE 对象（ppt/embeddings 合计 %d 字节）"
                       % emb_bytes)

    # =======================================================================
    # 减分项 3（-1）：新增对象观感瑕疵——圆角轮廓重复绘制形成双层描边，
    #                 或右边界越过第8页版心右侧界线 32.13cm 打破右对齐
    # =======================================================================
    flaws = []
    for sh in added:
        r = _rect(sh)
        if r[2] > RIGHT_MARGIN_LINE + 0.05:
            flaws.append("“%s”右边界 %.2fcm 越过版心右侧界线 %.2fcm"
                         % (sh.name, r[2], RIGHT_MARGIN_LINE))
    for i, a in enumerate(added):
        for b in added[i + 1:]:
            ra, rb = _rect(a), _rect(b)
            small, big = sorted((ra, rb), key=_area)
            if _area(small) > 0 and _overlap_area(ra, rb) >= 0.80 * _area(small):
                flaws.append("“%s”与“%s”外框重叠嵌套，圆角轮廓被重复绘制成双层描边"
                             % (a.name, b.name))
    if flaws:
        results.append("-1：新增对象存在观感瑕疵：%s" % "；".join(dict.fromkeys(flaws)))
    else:
        results.append("-0：新增对象无双层描边，右边界未越过版心右侧界线 %.2fcm，"
                       "与页面其余元素保持右对齐" % RIGHT_MARGIN_LINE)

    # =======================================================================
    # 减分项 4（-1）：交付的演示文稿文件名未保持为“区域协同汇报.pptx”
    # =======================================================================
    if os.path.basename(file_path) == EXP_PPTX_NAME:
        results.append("-0：交付演示文稿文件名保持为 %s" % EXP_PPTX_NAME)
    else:
        results.append("-1：交付演示文稿文件名改为“%s”，未保持为“%s”"
                       % (os.path.basename(file_path), EXP_PPTX_NAME))

    # =======================================================================
    # 减分项 5（-3）：第8页 6 列 8 行表格的文字内容被改动
    # =======================================================================
    tbl = None
    try:
        tbl = got["table"].table if got["table"] is not None else None
    except Exception:
        tbl = None
    txt_bad = []
    if tbl is None:
        txt_bad.append("第8页表格缺失")
    else:
        if len(tbl.rows) != 8 or len(tbl.columns) != 6:
            txt_bad.append("表格规格应为 6 列 8 行，实为 %d 列 %d 行"
                           % (len(tbl.columns), len(tbl.rows)))
        for ri in range(min(len(tbl.rows), 8)):
            for ci in range(min(len(tbl.columns), 6)):
                exp = TBL_TEXT[ri][ci]
                act = _flat(tbl.cell(ri, ci).text)
                if act != exp:
                    txt_bad.append("第%d行第%d列应为“%s”，实为“%s”"
                                   % (ri + 1, ci + 1, exp, act))
    if txt_bad:
        results.append("-3：第8页表格文字内容被改动：%s" % "；".join(txt_bad[:4]))
    else:
        results.append("-0：第8页表格文字内容与基线逐格一致（表头 Pipeline/Sector/Scope/"
                       "Stage/Q2 action/Value band，数据行 HarborLink 至 NorthBay 共七行）")

    # =======================================================================
    # 减分项 6（-3）：第8页表格的版式被改动（位置尺寸/行高/列宽/单元格填充/表头字体）
    # =======================================================================
    fmt_bad = []
    if tbl is None or got["table"] is None:
        fmt_bad.append("第8页表格缺失")
    else:
        gl, gt, gw, gh = _box(got["table"])
        for nm, act, exp in (("左", gl, TBL_BOX[0]), ("上", gt, TBL_BOX[1]),
                             ("宽", gw, TBL_BOX[2]), ("高", gh, TBL_BOX[3])):
            if not _near(act, exp):
                fmt_bad.append("%s应为%.2fcm，实为%.2fcm" % (nm, exp, act))
        for i, r in enumerate(tbl.rows):
            if not _near(_cm(r.height), TBL_ROW_H):
                fmt_bad.append("第%d行行高应为%.2fcm，实为%.2fcm"
                               % (i + 1, TBL_ROW_H, _cm(r.height)))
        for i, c in enumerate(tbl.columns):
            if i < len(TBL_COL_W) and not _near(_cm(c.width), TBL_COL_W[i]):
                fmt_bad.append("第%d列列宽应为%.2fcm，实为%.2fcm"
                               % (i + 1, TBL_COL_W[i], _cm(c.width)))
        for ri in range(min(len(tbl.rows), 8)):
            exp_fill = (TBL_FILL_HEAD if ri == 0
                        else TBL_FILL_BODY[0] if ri % 2 == 1 else TBL_FILL_BODY[1])
            for ci in range(min(len(tbl.columns), 6)):
                tc = tbl.cell(ri, ci)._tc
                el = tc.find("a:tcPr/a:solidFill/a:srgbClr", _NS)
                act = el.get("val").upper() if el is not None and el.get("val") else None
                if act != exp_fill:
                    fmt_bad.append("第%d行第%d列填充应为%s，实为%s"
                                   % (ri + 1, ci + 1, exp_fill, act))
        for ci in range(min(len(tbl.columns), 6)):
            p = tbl.cell(0, ci).text_frame.paragraphs[0]._p
            f = _para_font(p)
            if not _font_ok(f, *TBL_HEAD_FONT):
                fmt_bad.append("表头第%d列字体应为 Aptos %.2f磅加粗 %s，实为 %s %s磅 b=%s %s"
                               % (ci + 1, TBL_HEAD_FONT[1], TBL_HEAD_FONT[3],
                                  f.get("name"), f.get("size"), f.get("bold"), f.get("color")))
            if _para_algn(p) != "ctr":
                fmt_bad.append("表头第%d列未居中" % (ci + 1))
    if fmt_bad:
        results.append("-3：第8页表格版式被改动：%s" % "；".join(dict.fromkeys(fmt_bad))[:220])
    else:
        results.append("-0：第8页表格版式保持基线（左1.65/上4.37/宽24.00/高12.57cm，"
                       "行高均1.57cm，列宽3.68/3.94/5.08/3.17/5.59/2.54cm，"
                       "表头填充184669、数据行10304E与123758隔行交替，表头 Aptos 8.69磅加粗 F4F9FB 居中）")

    # =======================================================================
    # 减分项 7（-3）：第8页主标题内容或格式被改动
    #   基线：“Strategic Pipeline Status Update” 左1.4/上1.47，Aptos 26磅加粗 F4F9FB
    # =======================================================================
    def _check_text_block(sh, spec, tag, need_right_algn=False):
        """逐条比对文本块的内容、位置与字体；返回问题描述列表。"""
        txt, x, y, fn, sz, bd, col = spec
        bad = []
        if sh is None:
            return ["%s“%s”缺失" % (tag, txt)]
        act = _flat(_shape_text(sh))
        if act != _flat(txt):
            bad.append("%s内容应为“%s”，实为“%s”" % (tag, txt, act))
        if not _near(_cm(sh.left), x) or not _near(_cm(sh.top), y):
            bad.append("%s位置应为左%.2f/上%.2fcm，实为左%.2f/上%.2fcm"
                       % (tag, x, y, _cm(sh.left), _cm(sh.top)))
        try:
            p = sh.text_frame.paragraphs[0]._p
        except Exception:
            return bad + ["%s无法读取字体" % tag]
        f = _para_font(p)
        if not _font_ok(f, fn, sz, bd, col):
            bad.append("%s字体应为 %s %.2f磅%s %s，实为 %s %s磅 b=%s %s"
                       % (tag, fn, sz, "加粗" if bd else "", col,
                          f.get("name"), f.get("size"), f.get("bold"), f.get("color")))
        if need_right_algn and _para_algn(p) not in ("r", "just"):
            bad.append("%s未右对齐" % tag)
        return bad

    ttl_bad = _check_text_block(got["title"], TXT_TITLE, "主标题")
    if ttl_bad:
        results.append("-3：第8页主标题被改动：%s" % "；".join(ttl_bad)[:220])
    else:
        results.append("-0：第8页主标题“Strategic Pipeline Status Update”内容与格式保持基线"
                       "（左1.40/上1.47cm，Aptos 26磅加粗 F4F9FB）")

    # =======================================================================
    # 减分项 8（-3）：第8页辅助文字（栏目序号07、副标题、表格下方说明、页脚、页码）
    #                的内容或格式被改动
    # =======================================================================
    aux_bad = []
    tags = ("栏目序号", "副标题", "表格下方说明文字", "页脚", "页码")
    for i, spec in enumerate(TXT_AUX):
        aux_bad += _check_text_block(got["aux"][i], spec, tags[i],
                                     need_right_algn=(i == 4))
    if got["aux"][4] is not None and not _near(_cm(got["aux"][4].left), PAGE_NO_LEFT):
        aux_bad.append("页码未置于 %.2fcm 处" % PAGE_NO_LEFT)
    if aux_bad:
        results.append("-3：第8页辅助文字被改动：%s" % "；".join(aux_bad)[:240])
    else:
        results.append("-0：第8页栏目序号“07”、副标题、表格下方说明文字、页脚"
                       "“Regional Operating Review  |  FY2026 Q1”与页码“08/10”"
                       "的内容、位置、字体字号与颜色均保持基线")

    # =======================================================================
    # 减分项 9（-1）：第8页装饰元素被改动或删除
    #   深色背景矩形081226、满版背景图片、左上青色圆00C2CC、右上紫色圆7A6BFF、
    #   右下橙色圆F6B743、标题下方5.08cm青色短线00C2CC、底部30.73cm横线2C5970
    # =======================================================================
    deco_bad = []
    for name, kind, fill, x, y, w, h in DECOS:
        sh = got["decos"][name]
        if sh is None:
            deco_bad.append("%s缺失" % name)
            continue
        if _shape_fill_rgb(sh) != fill:
            deco_bad.append("%s填充色实为 %s" % (name, _shape_fill_rgb(sh)))
        bl, bt, bw, bh = _box(sh)
        if not (_near(bl, x) and _near(bt, y) and _near(bw, w) and _near(bh, h)):
            deco_bad.append("%s位置尺寸应为 左%.2f/上%.2f/宽%.2f/高%.2fcm，"
                            "实为 左%.2f/上%.2f/宽%.2f/高%.2fcm"
                            % (name, x, y, w, h, bl, bt, bw, bh))
    if got["pic"] is None:
        deco_bad.append("满版背景图片缺失")
    else:
        pl, pt, pw, ph = _box(got["pic"])
        if not (_near(pl, DECO_PIC[0]) and _near(pt, DECO_PIC[1])
                and _near(pw, DECO_PIC[2]) and _near(ph, DECO_PIC[3])):
            deco_bad.append("满版背景图片应铺满 0/0/%.2f/%.2fcm，实为 %.2f/%.2f/%.2f/%.2fcm"
                            % (DECO_PIC[2], DECO_PIC[3], pl, pt, pw, ph))
    if deco_bad:
        results.append("-1：第8页装饰元素被改动或删除：%s" % "；".join(deco_bad)[:220])
    else:
        results.append("-0：第8页深色背景矩形081226、满版背景图片、左上青色圆00C2CC、"
                       "右上紫色圆7A6BFF、右下橙色圆F6B743、标题下方5.08cm青色短线、"
                       "底部30.73cm横线2C5970 均在位且未被改动")

    # =======================================================================
    # 减分项 10（-3）：演示文稿页序或其余页面被改动
    #   总页数须为 10 页；08/10 这一页须仍处在第 8 页；其余 9 页形状数与关键文本不变
    # =======================================================================
    ord_bad = []
    if len(slides) != TOTAL_SLIDES:
        ord_bad.append("总页数应为 %d 页，实为 %d 页" % (TOTAL_SLIDES, len(slides)))
    if tgt_idx != TARGET_INDEX:
        ord_bad.append("标题为 Strategic Pipeline Status Update 的页面应为第 %d 页，实为第 %s 页"
                       % (TARGET_INDEX, tgt_idx))
    for idx, (cnt, texts) in OTHER_SLIDES.items():
        if idx > len(slides):
            ord_bad.append("缺少第 %d 页" % idx)
            continue
        sl = slides[idx - 1]
        shp = list(sl.shapes)
        if len(shp) != cnt:
            ord_bad.append("第%d页形状数应为%d，实为%d" % (idx, cnt, len(shp)))
        pool = [_flat(_shape_text(s)) for s in shp]
        for t in texts:
            if t not in pool:
                ord_bad.append("第%d页缺少文本“%s”" % (idx, t))
    if ord_bad:
        results.append("-3：演示文稿页序或其余页面被改动：%s" % "；".join(ord_bad)[:240])
    else:
        results.append("-0：演示文稿仍为 10 页，Strategic Pipeline Status Update 位于第 8 页，"
                       "其余 9 页（封面、Financial Snapshot、Performance Scorecard、"
                       "Diagnosis: Progress/Roadblocks、Actions and Next Steps、"
                       "Leadership Alignment Requests、Agreements and Request Table、"
                       "Rolling Forecast）内容与版式未变")

    # =======================================================================
    # 减分项 11（-1）：幻灯片页面尺寸被改动（33.87cm × 19.05cm，16:9 宽屏）
    # =======================================================================
    if _near(page_w, PAGE_W_CM, TOL_PAGE) and _near(page_h, PAGE_H_CM, TOL_PAGE):
        results.append("-0：页面尺寸保持 %.2fcm × %.2fcm（16:9 宽屏）" % (page_w, page_h))
    else:
        results.append("-1：页面尺寸应为 %.2fcm × %.2fcm，实为 %.2fcm × %.2fcm"
                       % (PAGE_W_CM, PAGE_H_CM, page_w, page_h))

    # =======================================================================
    # 减分项 12（-3）：被链接的工作簿内容、文件名或存放位置被改动
    #   工作簿完全未随交付提供时不在此条重复扣分（已由加分项 3 未得分体现）
    # =======================================================================
    wb_bad = []
    if wb_same:
        chk = _wb_check(wb_same)
        if chk is None:
            wb_bad = None                       # 环境无 openpyxl，本条不判定
        else:
            wb_bad += chk
    elif wb_sub:
        wb_bad.append("工作簿被移动到子目录 %s" % os.path.relpath(wb_sub, dir_path))
        chk = _wb_check(wb_sub)
        if chk:
            wb_bad += chk
    elif wb_renamed:
        wb_bad.append("工作簿被改名为“%s”" % os.path.basename(wb_renamed))
    else:
        wb_bad = None                           # 未随交付提供，此条不判定
    if wb_bad:
        results.append("-3：被链接工作簿 %s 被改动：%s"
                       % (EXP_XLSX_NAME, "；".join(wb_bad)[:220]))
    elif wb_bad is None:
        results.append("-0：交付目录内无可校验的 %s，本条不作扣分" % EXP_XLSX_NAME)
    else:
        results.append("-0：被链接工作簿 %s 文件名、存放位置与内容均未改动"
                       "（Strategic Pipeline 160行7列、Lead Development 220行24列、"
                       "Client Portfolio 480行43列，各表第1行标题行、第2行记录数说明行、"
                       "第3行字段表头行齐备）" % EXP_XLSX_NAME)

    # =======================================================================
    # 减分项 13（-1）：交付目录中残留与任务无关的中间产物
    # =======================================================================
    strays = _stray_files(dir_path, os.path.basename(file_path))
    if strays:
        results.append("-1：交付目录残留与任务无关的中间产物：%s"
                       % "、".join(strays[:8]))
    else:
        results.append("-0：交付目录仅含演示文稿与被链接工作簿，无脚本、截图、日志、"
                       "说明文档、Office 临时锁文件等残留")

    # =======================================================================
    # 汇总：results -> dim2_items（顺序与 _MAX_DELTAS_ORDER 严格对齐）
    #   加分项 max_delta>0，命中记满分；减分项 max_delta<0，命中即扣。
    #   维度1 不通过（文件损坏无法打开）时完成度总分直接记 0。
    # =======================================================================
    if len(results) != len(_MAX_DELTAS_ORDER):
        raise RuntimeError(
            f"dim2 项数不一致：results={len(results)} vs order={len(_MAX_DELTAS_ORDER)}"
        )
    dim2_items = []
    total_score = 0
    for raw, max_delta in zip(results, _MAX_DELTAS_ORDER):
        head, _, rule = raw.partition("：")
        try:
            delta = int(head.strip())
        except ValueError:
            delta = 0
        hit = delta != 0
        dim2_items.append({
            "rule": rule,
            "kind": "bonus" if max_delta > 0 else "penalty",
            "max_delta": max_delta,
            "delta": max_delta if hit else 0,
            "hit": hit,
        })
        total_score += (max_delta if hit else 0)
    if not result["dim1_pass"]:
        total_score = 0
    result["dim2_items"] = dim2_items
    result["total_score"] = max(0, total_score)
    result["max_score"] = MAX_SCORE_TOTAL
    return result


if __name__ == "__main__":
    if len(sys.argv) >= 2:
        _dir = sys.argv[1]
    else:
        _dir = os.path.dirname(os.path.abspath(__file__))
    _out = evaluate(_dir)
    print(json.dumps(_out, ensure_ascii=False, indent=2))
