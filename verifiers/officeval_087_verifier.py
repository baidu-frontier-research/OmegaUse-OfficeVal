#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# pyright: reportUnknownMemberType=false, reportUnknownVariableType=false, reportUnknownArgumentType=false, reportUnknownParameterType=false, reportMissingParameterType=false, reportPrivateUsage=false, reportUnknownLambdaType=false, reportGeneralTypeIssues=false

import os
import re
import sys
import json
import zipfile

from openpyxl import load_workbook
from openpyxl.utils import range_boundaries, get_column_letter

# 说明：按接口统一约定，本模块不修改全局 sys.stdout，也不 print 主结果。
# 主结果统一由 evaluate(dir_path) -> dict 返回；入参为“脚本所在目录的路径”，
# 由脚本自己在该目录内定位并打开被评估的工作簿。本文件仅在 __main__ 下作为
# 本地调试入口，才把 dict 序列化为 JSON 打印。
#
# 任务：把“识字主题分层作业.xlsx”的题目内容粘贴进“分层作业设计.xlsx”模板，
# 模板原始格式不得改变。作业内容既可以写在单元格里，也可以用浮动文本框/图片
# 叠放在单元格上（两种都是人工可接受的交付形态），故本脚本先把每个填写区域
# 还原成“视觉行”再做文本判定，见 _build_regions。

SCRIPT_ID = "087"
# 维度二加分项 max_delta 之和；也是维度一未通过时回填的 max_score（脚本没评上时为 0）
MAX_SCORE_TOTAL = 73
# 维度二减分项 max_delta 之和（负值），仅用于自检
MAX_DEDUCT_TOTAL = -36

# 维度二所有评分项按“结果追加顺序”的 max_delta 列表；末端把 results 里的
# “+N：...”/“-N：...”字符串反解为结构化 dim2_items 时用。顺序须与
# _evaluate_impl 中 results.append(...) 的顺序严格一致；共 43 项，
# 正项累加 = MAX_SCORE_TOTAL，负项累加 = MAX_DEDUCT_TOTAL。
_MAX_DELTAS_ORDER = [
    3,   # 01 标题/简介栏出现作业主题
    3,   # 02 B3:F3 分层要求
    3,   # 03 第1题题干 + 五个正确读音
    3,   # 04 干扰读音行五项且括号与上行对齐
    3,   # 05 五段词语在读音上方 + “游”处叶子图
    1,   # 06 “2.看拼音，写词语。”
    3,   # 07 qīng chén/xiǎo yǔ/hé fēng + 三个 2:1 田字格
    3,   # 08 huā duǒ/niǎo ér/fēi + 2:1、2:1、1:1 田字格
    5,   # 09 两个句子 + 5 处空缺留白
    3,   # 10 “3.你能圈出…” + 晨露/细雨/晚霞
    1,   # 11 “4.连一连。”
    3,   # 12 连一连上下两行且同序项对应
    1,   # 13 基础层 1234 顺序
    3,   # 14 “1.词语接龙。”另起一行 + 三个括号
    3,   # 15 “2.照样子，看图说一说。”
    3,   # 16 山水插图（长宽比≈2）靠左且完整在格内
    3,   # 17 插图右侧第一列 例/小鸟/柳叶/溪水
    3,   # 18 第二列 纸鸢/云朵
    1,   # 19 “3.仔细观察图片，填一填。”
    3,   # 20 山水插图（长宽比≈1.5）靠左且完整在格内
    5,   # 21 文段完整无缺字错字漏句
    5,   # 22 文段 8 处空缺各有田字格
    1,   # 23 提高层 123 顺序
    3,   # 24 拓展层两题独立成段左对齐
    3,   # 25 星星评价三段文字
    1,   # 26 三段之间 2-8 字符留白
    1,   # 27 自我评价/教师评价标签 + 可书写空行
    -3,  # 28 使用区域与 8 处合并单元格
    -3,  # 29 第二行标签 + B2/D2/F2 保持空白
    -3,  # 30 A3 两行说明文字
    -3,  # 31 A4:A6 “作业/设计/内容”及其字体
    -3,  # 32 B4/B5/B6 分层标题 + A7 反馈标题
    -3,  # 33 标题与各标签字体字号
    -3,  # 34 A-F 列宽
    -3,  # 35 各行行高
    -3,  # 36 五处填写区自动换行
    -3,  # 37 题目被概括/改写/自创表述
    -1,  # 38 C4/C5/C6 内容左对齐
    -1,  # 39 边框缺失或新增底纹/填充/彩色字体
    -1,  # 40 纸张 A4 / 纵向
    -1,  # 41 工作表数量及 Sheet2/Sheet3 留空
    -1,  # 42 批注、水印、交付说明文字
    -1,  # 43 图片变形、移出单元格、压住文字
]
assert len(_MAX_DELTAS_ORDER) == 43
assert sum(d for d in _MAX_DELTAS_ORDER if d > 0) == MAX_SCORE_TOTAL
assert sum(d for d in _MAX_DELTAS_ORDER if d < 0) == MAX_DEDUCT_TOTAL

_EXT_OK = (".xlsx", ".xlsm")
# 维度一不接受、但可能被误交的其它表格格式；命中时判维度一不通过，而不是报脚本异常。
_EXT_REJECT = (".xls", ".xlsb", ".csv", ".ods", ".et")

# 模板固有格式（减分项的比对基准）
_MAIN_SHEET_IDX = 0
_EXPECTED_DIMS = (1, 1, 6, 9)          # min_col, min_row, max_col, max_row
_EXPECTED_MERGES = {
    "A1:F1", "B3:F3", "C4:F4", "C5:F5", "C6:F6", "A4:A6", "A7:A9", "B7:F9",
}
_EXPECTED_COL_W = {"A": 13.38, "B": 15.25, "C": 13.75, "D": 18.00, "E": 12.88, "F": 16.13}
_COL_W_TOL = 1.5
_EXPECTED_ROW_H = {1: 39.0, 2: 41.0, 3: 68.0, 6: 174.0, 9: 30.0}
_ROW_H_TOL = 1.5
_ROW4_MIN_H = 162.0
_ROW5_MIN_H = 172.0

# 模板原有实线边框：每格按 左/右/上/下 记 "1"（有线）或 "0"（无线）。
# 判定规则是“不得比模板更少”，多画（例如给第 1 行补一圈框线）不扣分。
_TEMPLATE_BORDERS = {
    1: ["0000", "0000", "0000", "0000", "0000", "0000"],
    2: ["1111", "1111", "1111", "1111", "1111", "1111"],
    3: ["1111", "1111", "0011", "0011", "0011", "0111"],
    4: ["1111", "1111", "1111", "0011", "0011", "0111"],
    5: ["1100", "1111", "1111", "0011", "0011", "0111"],
    6: ["1101", "1111", "1111", "0011", "0011", "0111"],
    7: ["1111", "1111", "0010", "0010", "0010", "0110"],
    8: ["1100", "1000", "0000", "0000", "0000", "0100"],
    9: ["1101", "1001", "0001", "0001", "0001", "0101"],
}

# 模板标签文字
_LBL_ROW2 = {"A2": "班级", "C2": "任教教师", "E2": "时间"}
_LBL_A3 = ["作业设计简介", "（学习目标、题型说明等）"]
_LBL_A4 = ["作业", "设计", "内容"]
_LBL_B456 = {"B4": "基础层作业设计", "B5": "提高层作业设计", "B6": "拓展层作业设计"}
_LBL_A7 = "作业情况反馈"
# 需为宋体 11 号加粗的标签格
_BOLD11_CELLS = ("A2", "C2", "E2", "A3", "B4", "B5", "B6", "A7")

_THEME = "识字主题《四季小景》分层作业"
# 五组“词语 / 正确读音 / 干扰读音”
_WORDS5 = ["晨露", "清风", "细雨", "晚霞", "游"]
_READ_OK5 = ["chén lù", "qīng fēng", "xì yǔ", "wǎn xiá", "yóu"]
_READ_NG5 = ["chēn lù", "qīn fēng", "xī yǔ", "xiá wǎn", "yǒu"]
_PINYIN_G1 = ["qīng chén", "xiǎo yǔ", "hé fēng"]
_PINYIN_G2 = ["huā duǒ", "niǎo ér", "fēi"]
# 两个例句按空缺切分后的文字片段（第一句 2 处空缺、第二句 3 处空缺）。
# 允许第二组拼音行插在两句中间（源文件即如此排版），故按顺序逐段比对。
_SENT_SEGS = ["太阳从山后慢慢升起", "轻轻落在窗前", "送来阵阵", "清香",
              "花园里的", "开得正好", "从树梢掠过", "进了安静的小院"]
_LINK_ROW1 = ["晨露", "午阳", "晚霞", "夜风"]
_LINK_ROW2 = ["亮", "暖", "红", "轻"]
_COL1_ITEMS = ["小鸟", "柳叶", "溪水"]
_COL2_ITEMS = ["纸鸢", "云朵"]
# 观察填空文段按 8 处空缺切分后的 9 个片段（顺序比对，允许空缺处夹田字格）
_PARA_SEGS = [
    "清风吹过河岸边铺满了嫩嫩的",
    "草地上点缀着五颜六色的",
    "清亮的",
    "中",
    "慢慢游来游去还轻轻划出了",
    "纹远处的",
    "上开着浅浅的",
    "一群小鸟正快活地",
    "林间",
]
_EXT_TASKS = [
    ["和家人", "公园", "校园", "记录三种自然景物", "班级分享"],
    ["搜集两个描写季节的词语", "同学交流"],
]
_STAR_SEGS = ["今天我完成了", "颗星我已经累计了", "颗星请继续保持"]
# 改写/自创表述关键词（命中即认为未逐题照录）
_REWRITE_WORDS = ["辨音勾选", "情境填空", "参考要点", "看拼音写词", "圈景物字",
                  "看图表达", "观察记录", "辨音、"]
# 必须逐字出现的原始题干（缺 3 条以上视为概括改写）
_ORIG_STEMS = [
    "给词语选择正确的读音", "看拼音，写词语", "你能圈出这几个词中表示自然景物的字吗",
    "连一连", "词语接龙", "照样子，看图说一说", "仔细观察图片，填一填",
    "和家人到附近公园或校园走一走", "搜集两个描写季节的词语",
]
# 交付说明类文字
_DELIVERY_WORDS = ["填充完成", "已按要求粘贴", "修改说明", "已完成填充", "填写说明",
                   "交付说明", "水印"]

# 几何换算：EMU -> 磅
_EMU_PT = 12700.0
# 半角字符宽度（宋体/微软雅黑 11 号约 5.5 磅），用于把浮动文本框的横向位置
# 折算成“字符列”，从而与单元格内文字使用同一套坐标做对齐判定。
_HALF_PT = 5.5
_W_SPACE = 0.55        # 一个半角空格约占 0.55 个半角字符宽
_LINE_H_DEF = 16.0     # 文本框默认行高（磅）
_INSET_PT = 3.6        # 文本框默认上/左内边距（磅）

# 田字格长宽比容差：约 2:1 记 2 个字格，约 1:1 记 1 个字格
# （细则注：一个 2:1 田字格可用两个 1:1 田字格替换，反之亦然）
_RATIO_2 = (1.60, 2.60)
_RATIO_1 = (0.70, 1.40)
_TIAN_MAX_PT = 70.0    # 田字格图片的最大边长，用来和插图区分


def _norm(s):
    """统一空白与常见等价字符，保留标点。"""
    if s is None:
        return ""
    t = str(s)
    t = t.replace(chr(0x00A0), " ").replace(chr(0x2002), " ").replace(chr(0x2003), chr(0x3000))
    t = t.replace(" ", " ").replace("　", "　")
    t = t.replace("\r\n", "\n").replace("\r", "\n").replace("\t", " ")
    t = t.replace("（", "（").replace("）", "）")
    t = t.replace("✓", "√").replace("✔", "√")
    t = re.sub(r"[ ]+\n", "\n", t)
    return t


def _compact(s):
    """去掉所有空白，保留标点；用于“文本原样出现”类判定。"""
    return re.sub(r"[\s　]+", "", _norm(s))


def _loose(s):
    """去掉空白与全部标点/括号/下划线，只留文字与拼音；用于宽松比对。"""
    t = _compact(s)
    t = re.sub(r"[，。、；：？！“”‘’（）()\[\]【】《》…—\-_~·．.,;:!?\"'／/\\|]+", "", t)
    return t


def _contains_ordered(hay, tokens):
    """tokens 是否按顺序出现在 hay 中。"""
    pos = 0
    for tk in tokens:
        idx = hay.find(tk, pos)
        if idx < 0:
            return False
        pos = idx + len(tk)
    return True


def _missing_ordered(hay, tokens):
    """返回未按顺序命中的片段列表（用于把缺字/漏句写进说明）。"""
    miss = []
    pos = 0
    for tk in tokens:
        idx = hay.find(tk, pos)
        if idx < 0:
            miss.append(tk)
        else:
            pos = idx + len(tk)
    return miss


def _is_wide(ch):
    o = ord(ch)
    if ch == "　":
        return True
    return (0x1100 <= o <= 0x115F or 0x2190 <= o <= 0x2BFF or 0x2E80 <= o <= 0xA4CF
            or 0xAC00 <= o <= 0xD7A3 or 0xF900 <= o <= 0xFAFF or 0xFE30 <= o <= 0xFE6F
            or 0xFF00 <= o <= 0xFF60 or 0xFFE0 <= o <= 0xFFE6)
def _vis_w(s):
    """文本的视觉宽度，单位＝半角字符。全角字 2、半角字 1、空格 0.55。"""
    w = 0.0
    for ch in _norm(s):
        if ch == " ":
            w += _W_SPACE
        elif _is_wide(ch):
            w += 2.0
        else:
            w += 1.0
    return w


class _Line(object):
    """一条“视觉行”：由若干片段组成，片段可能来自单元格文字或浮动文本框。"""

    __slots__ = ("y", "frags")

    def __init__(self, y):
        self.y = y
        self.frags = []            # [(x_半角列, 文本)]

    def sort(self):
        self.frags.sort(key=lambda f: f[0])

    @property
    def text(self):
        """按横向位置渲染成一行字符串（片段之间按间距补空格）。"""
        out = ""
        cur = 0.0
        for x, t in self.frags:
            pad = int(round(x - cur))
            if out and pad < 1:
                pad = 1
            if not out and pad < 0:
                pad = 0
            out += " " * max(0, pad)
            out += t
            cur = x + _vis_w(t)
        return out

    def x_of(self, token, start=0.0):
        """token 左边缘的横向位置（半角列），找不到返回 None。"""
        for x, t in self.frags:
            if x + _vis_w(t) < start:
                continue
            idx = t.find(token)
            while idx >= 0:
                xx = x + _vis_w(t[:idx])
                if xx >= start - 0.01:
                    return xx
                idx = t.find(token, idx + 1)
        return None
def _spans(line, tokens):
    """依次取出 tokens 在该行的 (左边缘, 宽度)；任一缺失则返回 None。"""
    out = []
    cursor = 0.0
    for tk in tokens:
        x = line.x_of(tk, cursor)
        if x is None:
            return None
        out.append((x, _vis_w(tk)))
        cursor = x + _vis_w(tk)
    return out


def _gaps_ok(spans, min_gap=1.05):
    """相邻 token 之间的净间距是否都不小于 min_gap（约 2 个半角空格）。"""
    for i in range(1, len(spans)):
        if spans[i][0] - (spans[i - 1][0] + spans[i - 1][1]) < min_gap:
            return False
    return True


class _Obj(object):
    """一个浮动对象（图片或文本框），坐标已折算成磅。"""

    __slots__ = ("kind", "x", "y", "w", "h", "paras", "col", "row")

    def __init__(self, kind, x, y, w, h, paras, col, row):
        self.kind = kind           # "pic" / "txt" / "shape"
        self.x = x
        self.y = y
        self.w = w
        self.h = h
        self.paras = paras
        self.col = col
        self.row = row

    @property
    def ratio(self):
        return (self.w / self.h) if self.h > 0 else 0.0

    @property
    def cells(self):
        """当作田字格时折算的“字格数”：约 2:1 记 2，约 1:1 记 1，其余记 0。"""
        r = self.ratio
        if _RATIO_2[0] <= r <= _RATIO_2[1]:
            return 2
        if _RATIO_1[0] <= r <= _RATIO_1[1]:
            return 1
        return 0

    @property
    def is_tian(self):
        return (self.kind != "txt" and self.cells > 0
                and max(self.w, self.h) <= _TIAN_MAX_PT)
def _mdw(wb):
    """默认字体的最大数字宽度（像素）：中文字体 11 号为 8，Calibri 11 号为 7。
    列宽单位就是它，浮动对象的横向定位必须用同一把尺子。"""
    name, size = "", 11.0
    try:
        f = wb._fonts[0]
        name = f.name or ""
        size = float(f.sz or 11)
    except Exception:
        pass
    if size >= 11 and not re.match(r"^(Calibri|Arial|Times|Aptos)", name, re.I):
        return 8.0
    return 7.0


def _col_edges(ws, mdw, ncol=32):
    """返回 0..ncol 列左边界的累计磅值（下标 0 基）。"""
    default = ws.sheet_format.defaultColWidth or 8.43
    edges = [0.0]
    for i in range(1, ncol + 1):
        letter = get_column_letter(i)
        dim = ws.column_dimensions.get(letter)
        w = dim.width if (dim is not None and dim.width) else default
        edges.append(edges[-1] + int(float(w) * mdw + 5) * 0.75)
    return edges


def _row_edges(ws, nrow=64):
    """返回 0..nrow 行上边界的累计磅值（下标 0 基）。"""
    default = ws.sheet_format.defaultRowHeight or 14.4
    edges = [0.0]
    for r in range(1, nrow + 1):
        dim = ws.row_dimensions.get(r)
        h = dim.height if (dim is not None and dim.height) else default
        edges.append(edges[-1] + float(h))
    return edges


def _unesc(s):
    s = s.replace("&lt;", "<").replace("&gt;", ">").replace("&quot;", '"')
    s = s.replace("&apos;", "'").replace("&#10;", "\n").replace("&#13;", "")
    return s.replace("&amp;", "&")


_RE_ANCHOR = re.compile(r"<xdr:(oneCellAnchor|twoCellAnchor|absoluteAnchor)\b(.*?)</xdr:\1>", re.S)
_RE_FROM = re.compile(r"<xdr:from>\s*<xdr:col>(-?\d+)</xdr:col>\s*<xdr:colOff>(-?\d+)</xdr:colOff>"
                      r"\s*<xdr:row>(-?\d+)</xdr:row>\s*<xdr:rowOff>(-?\d+)</xdr:rowOff>", re.S)
_RE_TO = re.compile(r"<xdr:to>\s*<xdr:col>(-?\d+)</xdr:col>\s*<xdr:colOff>(-?\d+)</xdr:colOff>"
                    r"\s*<xdr:row>(-?\d+)</xdr:row>\s*<xdr:rowOff>(-?\d+)</xdr:rowOff>", re.S)
_RE_EXT = re.compile(r"<xdr:ext\s+cx=\"(\d+)\"\s+cy=\"(\d+)\"")
_RE_PARA = re.compile(r"<a:p\b.*?</a:p>|<a:p\s*/>", re.S)
_RE_T = re.compile(r"<a:t>(.*?)</a:t>", re.S)
def _sheet_parts(zf):
    """按工作簿顺序返回每个 sheet 的 xml 部件名。"""
    try:
        wbx = zf.read("xl/workbook.xml").decode("utf-8", "ignore")
        rels = zf.read("xl/_rels/workbook.xml.rels").decode("utf-8", "ignore")
    except KeyError:
        return []
    rmap = dict(re.findall(r"Id=\"([^\"]+)\"[^>]*?Target=\"([^\"]+)\"", rels))
    out = []
    for rid in re.findall(r"<sheet\b[^>]*r:id=\"([^\"]+)\"", wbx):
        tgt = rmap.get(rid, "")
        if not tgt:
            out.append("")
            continue
        tgt = tgt.lstrip("/")
        out.append(tgt if tgt.startswith("xl/") else "xl/" + tgt)
    return out


def _drawing_part(zf, sheet_part):
    """某个 sheet 关联的 drawing 部件名，没有则返回 ""。"""
    if not sheet_part:
        return ""
    rel = sheet_part.rsplit("/", 1)
    rel_name = rel[0] + "/_rels/" + rel[1] + ".rels"
    try:
        rels = zf.read(rel_name).decode("utf-8", "ignore")
    except KeyError:
        return ""
    m = re.search(r"Target=\"([^\"]*drawings/drawing\d+\.xml)\"", rels)
    if not m:
        return ""
    tgt = m.group(1).replace("../", "")
    return tgt if tgt.startswith("xl/") else "xl/" + tgt


def _abs_xy(cedge, redge, col, coloff, row, rowoff):
    cx = cedge[min(col, len(cedge) - 1)] + coloff / _EMU_PT
    ry = redge[min(row, len(redge) - 1)] + rowoff / _EMU_PT
    return cx, ry


def _parse_objs(zf, part, cedge, redge):
    """把 drawing xml 解析成 _Obj 列表（坐标单位：磅）。"""
    objs = []
    if not part:
        return objs
    try:
        xml = zf.read(part).decode("utf-8", "ignore")
    except KeyError:
        return objs
    for m in _RE_ANCHOR.finditer(xml):
        body = m.group(2)
        mf = _RE_FROM.search(body)
        if not mf:
            continue
        col, coloff, row, rowoff = (int(g) for g in mf.groups())
        x, y = _abs_xy(cedge, redge, col, coloff, row, rowoff)
        me = _RE_EXT.search(body)
        if me:
            w, h = int(me.group(1)) / _EMU_PT, int(me.group(2)) / _EMU_PT
        else:
            mt = _RE_TO.search(body)
            if not mt:
                continue
            tc, tco, tr, tro = (int(g) for g in mt.groups())
            x2, y2 = _abs_xy(cedge, redge, tc, tco, tr, tro)
            w, h = x2 - x, y2 - y
        paras = []
        for p in _RE_PARA.findall(body):
            paras.append(_unesc("".join(_RE_T.findall(p))))
        if paras and any(t.strip() for t in paras):
            kind = "txt"
        elif "<xdr:pic" in body:
            kind = "pic"
        else:
            kind = "shape"
        objs.append(_Obj(kind, x, y, w, h, paras, col, row))
    return objs
class _Region(object):
    """一个填写区域（合并单元格）：把单元格文字和叠放在它上面的浮动文本框
    统一还原成若干条“视觉行”，这样两种交付形态可以走同一套文本判定。"""

    __slots__ = ("name", "ref", "x0", "y0", "x1", "y1", "cell", "cell_lines", "objs", "lines")

    def __init__(self, name, ref, ws, objs, cedge, redge):
        self.name = name
        self.ref = ref
        c0, r0, c1, r1 = range_boundaries(ref)
        self.x0, self.x1 = cedge[c0 - 1], cedge[c1]
        self.y0, self.y1 = redge[r0 - 1], redge[r1]
        self.cell = ws.cell(row=r0, column=c0)
        val = self.cell.value
        self.cell_lines = [l for l in _norm(val).split("\n")] if val is not None else []
        self.objs = [o for o in objs
                     if self.x0 - 1.0 <= o.x + 2.0 < self.x1 and self.y0 - 1.0 <= o.y + 2.0 < self.y1]
        self.lines = self._build_lines()

    def _build_lines(self):
        frags = []
        lh = (self.cell.font.size or 11.0) * 1.45
        for i, ln in enumerate(self.cell_lines):
            if ln.strip():
                frags.append((0.0, self.y0 + _INSET_PT + (i + 0.5) * lh, ln))
        for o in self.objs:
            if o.kind != "txt":
                continue
            n = max(1, len(o.paras))
            bh = (o.h - 2 * _INSET_PT) / n
            # 文本框内行高：正常 11 号字约 16 磅；段落很多时按框高均摊，
            # 段落很少时不能把行摊开，否则纵向顺序会失真。
            bh = _LINE_H_DEF if bh <= 0 else min(_LINE_H_DEF, max(11.0, bh))
            for i, p in enumerate(o.paras):
                if p.strip():
                    frags.append(((o.x - self.x0) / _HALF_PT,
                                  o.y + _INSET_PT + (i + 0.5) * bh, _norm(p)))
        frags.sort(key=lambda f: (f[1], f[0]))
        lines, cur = [], None
        for x, y, t in frags:
            if cur is None or y - cur.y > 9.0:
                cur = _Line(y)
                lines.append(cur)
            cur.frags.append((x, t))
        for ln in lines:
            ln.sort()
        return lines

    @property
    def width(self):
        return self.x1 - self.x0

    @property
    def text(self):
        return "\n".join(l.text for l in self.lines)
    @property
    def raw_paras(self):
        """区域内所有“原始段落序列”：单元格自身一组，每个文本框各一组。
        空段落也保留，用于判定“留出可书写的空行”。"""
        out = []
        if self.cell_lines:
            out.append(list(self.cell_lines))
        for o in sorted(self.objs, key=lambda z: (z.y, z.x)):
            if o.kind == "txt":
                out.append([_norm(p) for p in o.paras])
        return out

    @property
    def compact(self):
        return _compact(self.text)

    @property
    def loose(self):
        return _loose(self.text)

    def idx_of(self, *tokens):
        """含全部 tokens 的第一条视觉行的行号，找不到返回 -1。"""
        for i, ln in enumerate(self.lines):
            t = _compact(ln.text)
            if all(_compact(tk) in t for tk in tokens):
                return i
        return -1

    def line_of(self, *tokens):
        i = self.idx_of(*tokens)
        return self.lines[i] if i >= 0 else None

    def num_lines(self):
        """题号行：返回 [(行号, 序号)]，只认行首的“N.”“N、”。"""
        out = []
        for i, ln in enumerate(self.lines):
            m = re.match(r"^\s*([1-9])\s*[.、．]", ln.text.lstrip())
            if m:
                out.append((i, int(m.group(1))))
        return out

    @property
    def pics(self):
        return [o for o in self.objs if o.kind != "txt"]

    @property
    def tians(self):
        return [o for o in self.objs if o.is_tian]

    @property
    def illus(self):
        """插图：非文本框且宽度明显大于田字格的图片。"""
        return [o for o in self.pics if o.w >= 80.0 and o.h >= 40.0]

    def contains(self, o, tol=3.0):
        return (o.x >= self.x0 - tol and o.y >= self.y0 - tol
                and o.x + o.w <= self.x1 + tol and o.y + o.h <= self.y1 + tol)


def _tian_rows(objs, tol=12.0):
    """把田字格按纵向位置聚成“行组”，只保留 2 个以上、合计 3 个字格以上的行组。"""
    rows, cur = [], []
    for o in sorted(objs, key=lambda z: z.y):
        if cur and o.y - cur[0].y > tol:
            rows.append(cur)
            cur = []
        cur.append(o)
    if cur:
        rows.append(cur)
    return [sorted(r, key=lambda z: z.x) for r in rows
            if len(r) >= 2 and sum(z.cells for z in r) >= 3]
_SONG = ("宋体", "SimSun", "NSimSun", "宋体-简", "STSong")


def _cell(ws, ref):
    return ws[ref]


def _ctext(ws, ref):
    v = ws[ref].value
    return _norm(v) if v is not None else ""


def _clines(ws, ref):
    t = _ctext(ws, ref)
    return [x for x in t.split("\n")]


def _font_ok(cell, name=None, size=None, bold=None):
    f = cell.font
    if name is not None:
        fn = (f.name or "")
        if not any(fn.startswith(n) for n in (_SONG if name == "宋体" else (name,))):
            return False
    if size is not None:
        try:
            if abs(float(f.sz or 0) - size) > 0.6:
                return False
        except Exception:
            return False
    if bold is not None and bool(f.bold) != bold:
        return False
    return True


def _align_ok(cell, horiz=None, vert=None, none_ok=()):
    a = cell.alignment
    if horiz is not None:
        h = a.horizontal
        if h != horiz and not (h is None and horiz in none_ok):
            return False
    if vert is not None:
        v = a.vertical
        if v != vert and not (v is None and vert in none_ok):
            return False
    return True


def _font_colored(cell):
    """字体是否被改成了黑/白以外的彩色。"""
    c = cell.font.color
    if c is None:
        return False
    rgb = None
    try:
        rgb = c.rgb
    except Exception:
        rgb = None
    if isinstance(rgb, str) and re.match(r"^[0-9A-Fa-f]{8}$", rgb):
        if rgb[2:].upper() not in ("000000", "FFFFFF"):
            return True
    try:
        if c.theme is not None and int(c.theme) not in (0, 1):
            return True
    except Exception:
        pass
    return False


def _bflags(cell):
    b = cell.border
    out = ""
    for side in (b.left, b.right, b.top, b.bottom):
        out += "1" if (side is not None and side.style) else "0"
    return out


def _has_fill(cell):
    f = cell.fill
    try:
        if f is None or f.fill_type in (None, "none"):
            return False
    except Exception:
        return False
    return True
def _token_xs(line, tokens):
    """依次取出 tokens 的左边缘位置（缺失记 None）。"""
    xs, cursor = [], 0.0
    for tk in tokens:
        x = line.x_of(tk, cursor)
        xs.append(x)
        if x is not None:
            cursor = x + _vis_w(tk)
    return xs


def _find_all_x(line, pattern):
    """该行中匹配 pattern 的全部 (左边缘, 匹配文本)，按位置排序。"""
    out = []
    for x, t in line.frags:
        for m in re.finditer(pattern, t):
            out.append((x + _vis_w(t[:m.start()]), m.group(0)))
    out.sort(key=lambda z: z[0])
    return out


def _paren_xs(line, tokens):
    """沿着一条视觉行依次找“读音 + 紧跟的括号”，返回每个括号左边缘的位置。
    找不到的项记 None。半角、全角括号都接受。"""
    xs = []
    cursor = 0.0
    for tk in tokens:
        x = line.x_of(tk, cursor)
        if x is None:
            xs.append(None)
            continue
        cursor = x + _vis_w(tk)
        best = None
        for p in ("（", "("):
            px = line.x_of(p, cursor)
            if px is not None and (best is None or px < best):
                best = px
        if best is None or best - cursor > 6.0:
            xs.append(None)
        else:
            xs.append(best)
            cursor = best + 1.0
    return xs


def _blank_slots(text, min_half=6):
    """行内“前后都有字、长度足够”的空白段数量，用作可书写空缺的计数。"""
    n = len(re.findall(r"(?<=\S)[ ]{%d,}(?=\S)" % min_half, text))
    n += len(re.findall(r"(?<=\S)　{%d,}(?=\S)" % max(2, min_half // 2), text))
    n += len(re.findall(r"(?<=\S)[_﹏﹍…]{4,}(?=\S)", text))
    return n


def _pick_file(dir_path):
    """在目录内定位被评估的工作簿：优先文件名含“分层作业设计”的。
    返回 (工作簿路径, 格式不被接受的候选文件名列表)。"""
    files = [f for f in sorted(os.listdir(dir_path)) if not f.startswith("~$")]
    cands = [f for f in files if f.lower().endswith(_EXT_OK)]
    rejects = [f for f in files if f.lower().endswith(_EXT_REJECT)]
    if not cands:
        return "", rejects
    pref = [f for f in cands if "分层作业设计" in f]
    return os.path.join(dir_path, (pref or cands)[0]), rejects
def _left_aligned(region):
    """区域内容是否左对齐：单元格文字看对齐属性，浮动文本框看题号行的左边缘。"""
    if any(l.strip() for l in region.cell_lines):
        h = region.cell.alignment.horizontal
        if h not in (None, "left", "justify", "general"):
            return False
    nums = region.num_lines()
    if not nums:
        return True
    xs_all = [min(x for x, _ in ln.frags) for ln in region.lines if ln.frags]
    base = min(xs_all) if xs_all else 0.0
    for i, _ in nums:
        ln = region.lines[i]
        if min(x for x, _ in ln.frags) - base > 4.0:
            return False
    return True


# add() 登记的字符串尾部常带一段“（实测值…）”，那是本地排查用的实测摘要，不写入
# 返回信息：返回的 rule 只保留规则原文，命中与未命中完全一致，不随判定结果变化。
# 只剥离结尾处、内部不再嵌套括号的一段全角括号，规则原文中部的括号不受影响。
_RULE_ONLY = re.compile(r"（[^（）]*）$")


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
        file_path, rejects = _pick_file(dir_path)
        if not file_path:
            if rejects:
                # 目录里只有 .xls 之类不被接受的格式：维度一直接判不通过。
                result["file_name"] = rejects[0]
                return _dim1_fail(result, "交付物格式为 %s，维度一仅接受 .xlsx、.xlsm"
                                          % os.path.splitext(rejects[0])[1].lower())
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
    try:
        wb = load_workbook(file_path, data_only=True)
    except Exception as exc:
        return _dim1_fail(result, "工作簿无法打开：%s: %s" % (type(exc).__name__, exc))
    if not wb.worksheets:
        return _dim1_fail(result, "工作簿内没有工作表")
    result["dim1_pass"] = True
    # 约定：维度一通过时 dim1_reason 保持空串，只有未通过才写错误摘要。
    result["dim1_reason"] = ""

    ws = wb.worksheets[_MAIN_SHEET_IDX]
    mdw = _mdw(wb)
    cedge = _col_edges(ws, mdw)
    redge = _row_edges(ws)
    with zipfile.ZipFile(file_path) as zf:
        names = zf.namelist()
        parts = _sheet_parts(zf)
        main_part = parts[_MAIN_SHEET_IDX] if len(parts) > _MAIN_SHEET_IDX else ""
        objs = _parse_objs(zf, _drawing_part(zf, main_part), cedge, redge)
        has_comment = any(n.startswith("xl/comments") or n.startswith("xl/threadedComments")
                          for n in names)
        other_draw = any(_drawing_part(zf, p)
                         for i, p in enumerate(parts) if i != _MAIN_SHEET_IDX and p)
    RT = _Region("A1", "A1:F1", ws, objs, cedge, redge)
    R3 = _Region("B3", "B3:F3", ws, objs, cedge, redge)
    R4 = _Region("C4", "C4:F4", ws, objs, cedge, redge)
    R5 = _Region("C5", "C5:F5", ws, objs, cedge, redge)
    R6 = _Region("C6", "C6:F6", ws, objs, cedge, redge)
    R7 = _Region("B7", "B7:F9", ws, objs, cedge, redge)
    flat4 = R4.text.replace("\n", " ")
    all_text = "\n".join([_ctext(ws, "A1"), R3.text, R4.text, R5.text, R6.text, R7.text])

    results = []

    def add(hit, delta, rule):
        # rule 尾部允许带“（实测值…）”，汇总时会剥掉，只把规则原文写进返回信息。
        results.append("%+d：%s" % ((delta if hit else 0), rule))

    # ---- 01 作业主题 ----
    hit = _compact(_THEME) in (_compact(RT.text) + R3.compact)
    add(hit, 3, "标题行或“作业设计简介”栏出现作业主题“%s”" % _THEME)

    # ---- 02 分层要求 ----
    b3 = R3.compact
    hit = ("基础" in b3 and ("提升" in b3 or "提高" in b3) and "必做" in b3
           and "拓展" in b3 and "选做" in b3)
    add(hit, 3, "B3:F3 说明分层要求：基础层/提高层必做、拓展层选做")

    # ---- 03 第 1 题题干 + 五个正确读音 ----
    i_stem1 = R4.idx_of("给词语选择正确的读音")
    i_read_ok = R4.idx_of("chén lù")
    got_ok = [r for r in _READ_OK5 if re.search(re.escape(r) + r"\s*[（(]", flat4)]
    hit = (i_stem1 >= 0 and "√" in R4.compact and len(got_ok) == 5
           and i_read_ok > i_stem1)
    add(hit, 3, "C4:F4 有“1.给词语选择正确的读音画“√”。”并另起一行给出五个正确读音"
                "（命中 %d/5）" % len(got_ok))

    # ---- 04 干扰读音行且括号与上行对齐 ----
    l_ok = R4.line_of("chén lù")
    l_ng = R4.line_of("chēn lù")
    n_ng = 0
    off = -1.0
    hit = False
    if l_ok is not None and l_ng is not None:
        xa = _paren_xs(l_ok, _READ_OK5)
        xb = _paren_xs(l_ng, _READ_NG5)
        n_ng = sum(1 for v in xb if v is not None)
        pairs = [(a, b) for a, b in zip(xa, xb) if a is not None and b is not None]
        if pairs:
            off = max(abs(a - b) for a, b in pairs)
        bad_pair = sum(1 for a, b in pairs if abs(a - b) > 2.5)
        hit = (n_ng == 5 and len(pairs) >= 4 and bad_pair <= 1)
    add(hit, 3, "C4:F4 有五个干扰读音且与上行括号横向基本对齐"
                "（干扰读音 %d/5，最大偏差 %.1f 字符）" % (n_ng, off))
    # ---- 05 五段词语位于读音上方 + “游”处叶子图 ----
    i_word = R4.idx_of(*_WORDS5[:4])
    hit = False
    dmax = -1.0
    if i_word >= 0 and l_ok is not None and i_word < i_read_ok:
        lw = R4.lines[i_word]
        sp = _spans(lw, _WORDS5[:4])
        xr = _token_xs(l_ok, _READ_OK5)
        if sp and _gaps_ok(sp) and all(v is not None for v in xr[:4]):
            dmax = max(abs(sp[k][0] - xr[k]) for k in range(4))
            x_you = lw.x_of("游")
            leaf = None
            if x_you is not None:
                for o in R4.pics:
                    if o.w < 15 or o.h < 15:
                        continue
                    if o.y - 8.0 <= lw.y <= o.y + o.h + 8.0 and \
                            abs((o.x - R4.x0) / _HALF_PT - x_you) <= 9.0:
                        leaf = o
                        break
            hit = (dmax <= 6.0 and x_you is not None and leaf is not None)
    add(hit, 3, "C4:F4 读音上方有“晨露/清风/细雨/晚霞/游”五段文本、间距分明且各自对位，"
                "“游”处有叶子图片（对位最大偏差 %.1f 字符）" % dmax)

    # ---- 06 第 2 题题干 ----
    hit = "2.看拼音，写词语" in R4.compact
    add(hit, 1, "C4:F4 有文本“2.看拼音，写词语。”")

    # ---- 07 第一组拼音 + 三个 2:1 田字格 ----
    rows4 = _tian_rows(R4.tians)
    g1 = [p for p in _PINYIN_G1 if p in flat4]
    n_g1 = sum(o.cells for o in rows4[0]) if rows4 else 0
    hit = (len(g1) == 3 and n_g1 >= 5)
    add(hit, 3, "C4:F4 有“qīng chén/xiǎo yǔ/hé fēng”且下一行有三个 2:1 田字格"
                "（拼音 %d/3，田字格合计 %d 字格）" % (len(g1), n_g1))

    # ---- 08 第二组拼音 + 2:1、2:1、1:1 田字格 ----
    g2 = [p for p in _PINYIN_G2 if p in flat4]
    n_g2 = sum(o.cells for o in rows4[1]) if len(rows4) >= 2 else 0
    hit = (len(g2) == 3 and 4 <= n_g2 <= 7)
    add(hit, 3, "C4:F4 有“huā duǒ/niǎo ér/fēi”且下一行有 2:1、2:1、1:1 田字格"
                "（拼音 %d/3，田字格合计 %d 字格）" % (len(g2), n_g2))

    # ---- 09 两个例句 + 5 处留白 ----
    miss_sent = _missing_ordered(R4.loose, _SENT_SEGS)
    slots = 0
    for ln in R4.lines:
        tc = _compact(ln.text)
        if any(_compact(p) in tc for p in _SENT_SEGS):
            slots += _blank_slots(ln.text, 6)
    hit = (not miss_sent and slots >= 5)
    add(hit, 5, "C4:F4 两个例句文字完整且共 5 处不少于 6 空格的留白"
                "（缺失片段 %d 个，留白 %d 处）" % (len(miss_sent), slots))
    # ---- 10 第 3 题题干 + 三个词语 ----
    i_stem3 = R4.idx_of("3.你能圈出这几个词中表示自然景物的字吗")
    hit = False
    if i_stem3 >= 0:
        for ln in R4.lines[i_stem3 + 1:i_stem3 + 3]:
            sp = _spans(ln, ["晨露", "细雨", "晚霞"])
            if sp and _gaps_ok(sp):
                hit = True
                break
    add(hit, 3, "C4:F4 有“3.你能圈出这几个词中表示自然景物的字吗？”"
                "且下一行为间距分明的“晨露/细雨/晚霞”")

    # ---- 11 第 4 题题干 ----
    i_stem4 = R4.idx_of("4.连一连")
    add(i_stem4 >= 0, 1, "C4:F4 有文本“4.连一连。”")

    # ---- 12 连一连上下两行同序项对应 ----
    hit = False
    dmax = -1.0
    if i_stem4 >= 0:
        cand = R4.lines[i_stem4 + 1:i_stem4 + 6]
        l1 = i1 = None
        for k, ln in enumerate(cand):
            if _spans(ln, _LINK_ROW1):
                l1, i1 = ln, k
                break
        if l1 is not None:
            for ln in cand[i1 + 1:]:
                s1 = _spans(l1, _LINK_ROW1)
                s2 = _spans(ln, _LINK_ROW2)
                if not s2:
                    continue
                dmax = max(abs(s1[k][0] - s2[k][0]) for k in range(4))
                hit = (_gaps_ok(s1) and _gaps_ok(s2) and dmax <= 6.0)
                break
    add(hit, 3, "C4:F4 连一连有“晨露 午阳 晚霞 夜风”与“亮 暖 红 轻”上下两行、"
                "项间有间距且同序项上下对应（最大错位 %.1f 字符）" % dmax)

    # ---- 13 基础层题号 1234 ----
    seq4 = [n for _, n in R4.num_lines() if n <= 4]
    hit = (seq4 == [1, 2, 3, 4])
    add(hit, 1, "C4:F4 内 1、2、3、4 四道题依次自上而下排列（实测序号 %s）"
                % ("".join(str(n) for n in seq4) or "无"))
    # ---- 14 词语接龙 ----
    i14 = R5.idx_of("1.词语接龙")
    hit = False
    gmin = -1.0
    if i14 >= 0 and "清晨" not in _compact(R5.lines[i14].text):
        for ln in R5.lines[i14 + 1:i14 + 3]:
            if "清晨" not in ln.text or "←" in ln.text:
                continue
            grp = _find_all_x(ln, r"[（(][\s　]*[）)]")
            if len(grp) >= 3 and ln.text.count("→") >= 2:
                gaps = [grp[k][0] - (grp[k - 1][0] + _vis_w(grp[k - 1][1]))
                        for k in range(1, len(grp))]
                gmin = min(gaps)
                hit = gmin > 2.0
            break
    add(hit, 3, "C5:F5 第一行“1.词语接龙。”、第二行“清晨→(  )→(  )→(  )”，箭头统一向右、"
                "括号间隔大于 2 字符（最小间隔 %.1f 字符）" % gmin)

    # ---- 15 第 2 题题干 ----
    i15 = R5.idx_of("2.照样子，看图说一说")
    add(i15 >= 0, 3, "C5:F5 有文本“2.照样子，看图说一说。”")

    # ---- 16 山水插图（长宽比≈2）----
    y15 = R5.lines[i15].y if i15 >= 0 else -1.0e9
    illu1 = None
    for o in sorted(R5.illus, key=lambda z: z.y):
        if 1.60 <= o.ratio <= 2.45 and o.y >= y15 - 4.0:
            illu1 = o
            break
    hit = (illu1 is not None
           and (illu1.x - R5.x0) <= 0.25 * R5.width and R5.contains(illu1))
    add(hit, 3, "C5:F5 “2.照样子”下方有长宽比≈2 的山水插图，靠左且完整在格内"
                "（%s）" % ("长宽比 %.2f，左偏移 %.0f 磅" % (illu1.ratio, illu1.x - R5.x0)
                           if illu1 is not None else "未找到插图"))

    # ---- 17 插图右侧第一列 ----
    i_ex = R5.idx_of("例", "微风轻")
    rows1 = []
    if i_ex >= 0:
        cur = i_ex
        for tk in _COL1_ITEMS:
            found = -1
            for j in range(cur + 1, min(len(R5.lines), i_ex + 9)):
                if (R5.lines[j].x_of(tk) is not None
                        and re.search(re.escape(tk) + r"[\s　]*[（(]", R5.lines[j].text)):
                    found = j
                    break
            if found < 0:
                rows1 = []
                break
            rows1.append((found, R5.lines[found].x_of(tk)))
            cur = found
    hit = False
    if rows1:
        xs1 = [x for _, x in rows1]
        right_ok = (illu1 is None
                    or min(xs1) * _HALF_PT >= (illu1.x - R5.x0) + illu1.w * 0.6)
        hit = (max(xs1) - min(xs1) <= 3.0 and right_ok)
    add(hit, 3, "C5:F5 插图右侧第一列依次为“例：微风轻/小鸟(  )/柳叶(  )/溪水(  )”")

    # ---- 18 插图右侧第二列 ----
    hit = False
    if rows1:
        ok2 = 0
        for tk in _COL2_ITEMS:
            for j in range(i_ex, min(len(R5.lines), i_ex + 9)):
                x = R5.lines[j].x_of(tk)
                if x is None or not re.search(re.escape(tk) + r"[\s　]*[（(]", R5.lines[j].text):
                    continue
                if x > min(x2 for _, x2 in rows1) + 2.0 or j > rows1[-1][0]:
                    ok2 += 1
                    break
        hit = (ok2 == 2)
    add(hit, 3, "C5:F5 第二列（或紧随第一列之后）为“纸鸢(  )”“云朵(  )”")
    # ---- 19 第 3 题题干 ----
    i19 = R5.idx_of("3.仔细观察图片，填一填")
    add(i19 >= 0, 1, "C5:F5 有文本“3.仔细观察图片，填一填。”")

    # ---- 20 山水插图（长宽比≈1.5）----
    y19 = R5.lines[i19].y if i19 >= 0 else -1.0e9
    illu2 = None
    for o in sorted(R5.illus, key=lambda z: z.y):
        if o is illu1:
            continue
        if 1.20 <= o.ratio <= 1.80 and o.y >= y19 - 4.0:
            illu2 = o
            break
    hit = (illu2 is not None
           and (illu2.x - R5.x0) <= 0.25 * R5.width and R5.contains(illu2))
    add(hit, 3, "C5:F5 “3.仔细观察图片”下方有长宽比≈1.5 的山水插图，靠左且完整在格内"
                "（%s）" % ("长宽比 %.2f，左偏移 %.0f 磅" % (illu2.ratio, illu2.x - R5.x0)
                           if illu2 is not None else "未找到插图"))

    # ---- 21 观察填空文段完整 ----
    miss_para = _missing_ordered(R5.loose, _PARA_SEGS)
    hit = not miss_para
    add(hit, 5, "C5:F5 第二张插图下方文段完整、无缺字错字漏句（缺失片段 %d 个%s）"
                % (len(miss_para), ("：" + "/".join(miss_para[:3])) if miss_para else ""))

    # ---- 22 文段 8 处空缺各有田字格 ----
    i_para = R5.idx_of("清风吹过")
    y_ref = illu2.y if illu2 is not None else (
        R5.lines[i_para].y - 24.0 if i_para >= 0 else -1.0e9)
    tian5 = [o for o in R5.tians if o.y >= y_ref - 4.0]
    inside = all(R5.contains(o) for o in tian5)
    hit = (8 <= len(tian5) <= 18 and sum(o.cells for o in tian5) >= 8 and inside)
    add(hit, 5, "C5:F5 文段 8 处空缺各有田字格图片、与前后文字同行且未压住文字"
                "（实测田字格 %d 个 / %d 字格）" % (len(tian5), sum(o.cells for o in tian5)))

    # ---- 23 提高层题号 123 ----
    seq5 = [n for _, n in R5.num_lines() if n <= 3]
    hit = (seq5 == [1, 2, 3])
    add(hit, 1, "C5:F5 内 1、2、3 三道题依次自上而下排列（实测序号 %s）"
                % ("".join(str(n) for n in seq5) or "无"))
    # ---- 24 拓展层两题 ----
    t6 = R6.compact
    have6 = [all(k in t6 for k in grp) for grp in _EXT_TASKS]
    i6a, i6b = R6.idx_of("和家人"), R6.idx_of("搜集两个描写季节的词语")
    seq6 = [n for _, n in R6.num_lines() if n <= 2]
    hit = (all(have6) and 0 <= i6a < i6b and seq6 == [1, 2] and _left_aligned(R6))
    add(hit, 3, "C6:F6 两道拓展题文字完整、序号 1/2 各自独立成段且左侧对齐")

    # ---- 25 星星评价三段文字 ----
    miss_star = _missing_ordered(R7.loose, _STAR_SEGS)
    hit = not miss_star
    add(hit, 3, "B7:F9 有“今天我完成了”“颗星，我已经累计了”“颗星，请继续保持！”三段文字"
                "（缺失 %d 段）" % len(miss_star))

    # ---- 26 三段之间 2-8 字符留白 ----
    l_star = R7.line_of("今天我完成了")
    hit = False
    if l_star is not None:
        t = l_star.text
        m1 = re.search(r"今天我完成了([ 　_﹏﹍…\-]{2,8})颗星", t)
        m2 = re.search(r"累计了?([ 　_﹏﹍…\-]{2,8})颗星", t)
        hit = bool(m1 and m2)
    add(hit, 1, "B7:F9 三段文字之间各有一处 2-8 字符的下划线或空白留白")

    # ---- 27 自我评价/教师评价标签 + 可书写空行 ----
    t7 = R7.compact
    blank_after = False
    for paras in R7.raw_paras:
        for k, p in enumerate(paras):
            if "教师评价" not in p:
                continue
            if re.search(r"[ 　_﹏﹍…\-]{6,}\s*$", p):
                blank_after = True
            if k + 1 < len(paras) and not paras[k + 1].strip():
                blank_after = True
            if k + 1 == len(paras) and R7.lines:
                blank_after = blank_after or (R7.y1 - R7.lines[-1].y) >= 24.0
    hit = ("自我评价" in t7 and "教师评价" in t7 and blank_after)
    add(hit, 1, "B7:F9 含“自我评价”“教师评价”两项标签，并为教师评价留出可书写空行")
    # ================= 减分项 =================
    # ---- 28 使用区域与 8 处合并单元格 ----
    merges = set(str(r) for r in ws.merged_cells.ranges)
    dims_ok = (ws.max_row <= _EXPECTED_DIMS[3] and ws.max_column <= _EXPECTED_DIMS[2])
    hit = not (dims_ok and merges == _EXPECTED_MERGES)
    add(hit, -3, "工作表使用区域或 8 处合并单元格被改动（实测区域 %s，合并 %d 处）"
                 % (ws.dimensions, len(merges)))

    # ---- 29 第二行标签与空格 ----
    bad29 = [k for k, v in _LBL_ROW2.items() if _compact(_ctext(ws, k)) != v]
    filled = [k for k in ("B2", "D2", "F2") if _ctext(ws, k).strip()]
    hit = bool(bad29 or filled)
    add(hit, -3, "第二行标签被改动或 B2/D2/F2 被填入文字（异常标签 %s，被填 %s）"
                 % ("/".join(bad29) or "无", "/".join(filled) or "无"))

    # ---- 30 A3 说明文字 ----
    a3 = [x for x in _clines(ws, "A3") if x.strip()]
    hit = not (len(a3) == 2 and _compact(a3[0]) == _LBL_A3[0]
               and _loose(a3[1]) == _loose(_LBL_A3[1]))
    add(hit, -3, "A3 不是“作业设计简介”与“（学习目标、题型说明等）”两行")

    # ---- 31 A4:A6 三行文字与字体 ----
    a4raw = _clines(ws, "A4")
    a4 = [x for x in a4raw if x.strip()]
    c_a4 = ws["A4"]
    hit = not (a4 == _LBL_A4 and len(a4raw) >= 5
               and _font_ok(c_a4, "宋体", 14, True)
               and _align_ok(c_a4, "center", "center"))
    add(hit, -3, "A4:A6 不是“作业/设计/内容”三行（行间留空）或不是宋体 14 号加粗居中")

    # ---- 32 分层标题与反馈标题 ----
    bad32 = [k for k, v in _LBL_B456.items() if _compact(_ctext(ws, k)) != v]
    if _compact(_ctext(ws, "A7")) != _LBL_A7:
        bad32.append("A7")
    hit = bool(bad32)
    add(hit, -3, "B4/B5/B6 分层标题或 A7“作业情况反馈”被改动（异常 %s）"
                 % ("/".join(bad32) or "无"))

    # ---- 33 标题与标签字体字号 ----
    c_a1 = ws["A1"]
    bad33 = []
    if not (_font_ok(c_a1, "宋体", 18, True) and _align_ok(c_a1, "center", "center")):
        bad33.append("A1")
    for k in _BOLD11_CELLS:
        if not _font_ok(ws[k], "宋体", 11, True):
            bad33.append(k)
    hit = bool(bad33)
    add(hit, -3, "标题或标签字体字号不符（异常 %s）" % ("/".join(bad33) or "无"))
    # ---- 34 列宽 ----
    bad34 = []
    for col, want in _EXPECTED_COL_W.items():
        dim = ws.column_dimensions.get(col)
        got = float(dim.width) if (dim is not None and dim.width) else 0.0
        if abs(got - want) > _COL_W_TOL:
            bad34.append("%s=%.2f" % (col, got))
    hit = bool(bad34)
    add(hit, -3, "A 至 F 列列宽超出 ±1.5 字符（异常 %s）" % ("/".join(bad34) or "无"))

    # ---- 35 行高 ----
    bad35 = []

    def _rh(r):
        dim = ws.row_dimensions.get(r)
        return float(dim.height) if (dim is not None and dim.height) else 0.0

    for r, want in _EXPECTED_ROW_H.items():
        if abs(_rh(r) - want) > _ROW_H_TOL:
            bad35.append("%d=%.0f" % (r, _rh(r)))
    if _rh(4) < _ROW4_MIN_H:
        bad35.append("4=%.0f" % _rh(4))
    if _rh(5) < _ROW5_MIN_H:
        bad35.append("5=%.0f" % _rh(5))
    hit = bool(bad35)
    add(hit, -3, "行高不符（异常 %s）" % ("/".join(bad35) or "无"))

    # ---- 36 五处填写区自动换行 ----
    bad36 = []
    for R in (R3, R4, R5, R6, R7):
        if not any(l.strip() for l in R.cell_lines):
            continue                      # 内容放在浮动文本框里，不涉及单元格换行
        if bool(R.cell.alignment.wrap_text):
            continue
        if len(R.cell_lines) > 1:
            bad36.append(R.name)
        elif max(_vis_w(l) for l in R.cell_lines) * _HALF_PT > R.width * 1.15:
            bad36.append(R.name)
    hit = bool(bad36)
    add(hit, -3, "填写区未自动换行、文字挤在一行或超出右边界（异常 %s）"
                 % ("/".join(bad36) or "无"))

    # ---- 37 题目被概括/改写 ----
    ct_all = _compact(all_text)
    hit_words = [w for w in _REWRITE_WORDS if w in ct_all]
    miss_stem = [s for s in _ORIG_STEMS if _compact(s) not in ct_all]
    hit = bool(hit_words) or len(miss_stem) >= 3
    add(hit, -3, "题目被概括/改写/自创表述（可疑表述 %s；缺失原题干 %d 条）"
                 % ("/".join(hit_words) or "无", len(miss_stem)))

    # ---- 38 内容左对齐 ----
    bad38 = [R.name for R in (R4, R5, R6) if not _left_aligned(R)]
    hit = bool(bad38)
    add(hit, -1, "C4/C5/C6 内容未左对齐、题号未沿左边界对齐（异常 %s）"
                 % ("/".join(bad38) or "无"))
    # ---- 39 边框与装饰 ----
    lost_border, deco = [], []
    for r in range(1, _EXPECTED_DIMS[3] + 1):
        want_row = _TEMPLATE_BORDERS.get(r)
        for c in range(1, _EXPECTED_DIMS[2] + 1):
            cell = ws.cell(row=r, column=c)
            if want_row is not None:
                want, got = want_row[c - 1], _bflags(cell)
                if any(want[k] == "1" and got[k] == "0" for k in range(4)):
                    lost_border.append(cell.coordinate)
            if _has_fill(cell) or _font_colored(cell):
                deco.append(cell.coordinate)
    hit = bool(lost_border or deco)
    add(hit, -1, "实线边框缺失或新增底纹/填充/彩色字体（缺框 %s；装饰 %s）"
                 % ("/".join(lost_border[:4]) or "无", "/".join(deco[:4]) or "无"))

    # ---- 40 纸张与方向 ----
    ps = ws.page_setup
    try:
        paper = int(ps.paperSize) if ps.paperSize else 0
    except Exception:
        paper = 0
    orient = ps.orientation
    hit = not (paper == 9 and orient in (None, "portrait", "default"))
    add(hit, -1, "页面设置被破坏（纸张代码 %s，方向 %s）" % (paper or "缺失", orient or "缺失"))

    # ---- 41 工作表数量与 Sheet2/Sheet3 ----
    dirty_sheets = []
    for i, wsx in enumerate(wb.worksheets):
        if i == _MAIN_SHEET_IDX:
            continue
        if wsx.max_row > 1 or wsx.max_column > 1 or wsx["A1"].value is not None:
            dirty_sheets.append(wsx.title)
    hit = (len(wb.sheetnames) < 3) or bool(dirty_sheets) or other_draw
    add(hit, -1, "工作表少于 3 个或 Sheet2/Sheet3 被写入内容（工作表 %d 个，被写入 %s）"
                 % (len(wb.sheetnames), "/".join(dirty_sheets) or "无"))

    # ---- 42 批注、水印、交付说明文字 ----
    n_cmt = 0
    for row in ws.iter_rows(min_row=1, max_row=9, min_col=1, max_col=6):
        for cell in row:
            if getattr(cell, "comment", None) is not None:
                n_cmt += 1
    hdr = ""
    try:
        for part in (ws.oddHeader, ws.oddFooter):
            for pos in (part.left, part.center, part.right):
                hdr += (pos.text or "")
    except Exception:
        hdr = ""
    dw = [w for w in _DELIVERY_WORDS if w in ct_all or w in _compact(hdr)]
    hit = bool(has_comment or n_cmt or dw)
    add(hit, -1, "出现批注、水印或交付说明文字（批注 %d 处；可疑文字 %s）"
                 % (n_cmt, "/".join(dw) or "无"))

    # ---- 43 图片变形、移出单元格 ----
    regions_all = (RT, R3, R4, R5, R6, R7)
    bad43 = []
    for o in objs:
        host = None
        for R in regions_all:
            if R.x0 - 1.0 <= o.x + 2.0 < R.x1 and R.y0 - 1.0 <= o.y + 2.0 < R.y1:
                host = R
                break
        if host is None or not host.contains(o):
            bad43.append("越界%.0fx%.0f" % (o.w, o.h))
        elif o.kind != "txt":
            if not (0.30 <= o.ratio <= 6.0):
                bad43.append("变形%.0fx%.0f" % (o.w, o.h))
            elif max(o.w, o.h) <= _TIAN_MAX_PT and o.cells == 0:
                bad43.append("方格变形%.0fx%.0f" % (o.w, o.h))
    hit = bool(bad43)
    add(hit, -1, "图片被拉伸变形或移出所属单元格（异常 %d 处 %s）"
                 % (len(bad43), "/".join(bad43[:3])))
    # ---- 汇总：把 results 字符串反解为结构化 dim2_items ----
    if len(results) != len(_MAX_DELTAS_ORDER):
        raise RuntimeError(f"dim2 项数不一致：results={len(results)} vs order={len(_MAX_DELTAS_ORDER)}")
    dim2_items = []
    total_score = 0
    for raw, max_delta in zip(results, _MAX_DELTAS_ORDER):
        head, _, rule = raw.partition("：")
        try:
            delta = int(head.replace("+", "").strip())
        except ValueError:
            delta = 0
        # 加分项：delta>0 表示拿到分；减分项：delta<0 表示扣分成立。
        # 两者统一用 hit 表示“该评分项判定成立”。
        hit = (delta != 0)
        dim2_items.append({"rule": _RULE_ONLY.sub("", rule), "max_delta": max_delta,
                           "delta": max_delta if hit else 0, "hit": hit})
        total_score += (max_delta if hit else 0)
    result["dim2_items"] = dim2_items
    # 扣分不倒扣：完成度总分下限为 0
    result["total_score"] = max(0, total_score)
    result["max_score"] = MAX_SCORE_TOTAL
    return result


if __name__ == "__main__":
    _dir = sys.argv[1] if len(sys.argv) > 1 else os.path.dirname(os.path.abspath(__file__))
    _out = evaluate(_dir)
    print(json.dumps(_out, ensure_ascii=False, indent=2))
