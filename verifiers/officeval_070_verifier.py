#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# pyright: reportUnknownMemberType=false, reportUnknownVariableType=false, reportUnknownArgumentType=false, reportPrivateUsage=false, reportOptionalMemberAccess=false, reportAttributeAccessIssue=false, reportMissingImports=false, reportUnknownParameterType=false, reportMissingParameterType=false, reportGeneralTypeIssues=false

import os
import re
import sys
import json
import math

from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.oxml.ns import qn

# 说明：按接口统一约定，本模块不修改全局 sys.stdout，也不 print 主结果。
# 主结果统一由 evaluate(dir_path) -> dict 返回；入参为“脚本所在目录的路径”，
# 由本模块自行在该目录内定位并打开被评估的 .pptx。只有把本文件当命令行工具
# 直接运行（__main__）时，才把返回的 dict 序列化成 JSON 打印，便于本地调试。

SCRIPT_ID = "070"
# 维度二全部“加分项”额度之和；维度一未通过或脚本异常时 total_score 记 0
MAX_SCORE_TOTAL = 21

# 维度二各评分项按 results.append(...) 的追加顺序排列的额度表：
#   正数 = 加分项，命中记满额、未命中记 0；
#   负数 = 减分项，合规记 0、触发则扣该额度（不计入 max_score）。
# 划分口径：加分项只放“指令明确下达、且输入文件本身并不满足、必须动手才能做到”
# 的任务（把六块文字放进板块框内、版式改成参照两块的样子）；输入文件本来就满足
# 的项一律不计加分，改为减分项——它们衡量的是“改动过程有没有把原来好的地方弄坏”。
# 顺序必须与 _evaluate_impl 中 results.append(...) 的书写顺序严格一致。
_DELTAS_ORDER = [
    3,    # 01 编号01板块文字落在板块框内
    3,    # 02 编号02板块文字落在板块框内
    3,    # 03 编号03板块文字落在板块框内
    3,    # 04 编号04板块文字落在板块框内
    3,    # 05 编号05板块文字落在板块框内
    3,    # 06 编号06板块文字落在板块框内
    3,    # 07 六块版式改成与“临床问题”“科学假说”一致
    -3,   # 08 六个板块文字出现叠压
    -1,   # 09 页面尺寸被改动
    -3,   # 10 六块正文字号不统一
    -1,   # 11 六块标题字号不统一
    -3,   # 12 六块正文字号低于8磅
    -1,   # 13 六块标题字号低于10磅或掉粗
    -3,   # 14 六块正文行距超过单倍或段前后不为0磅
    -1,   # 15 六块文字内容不完整
    -3,   # 16 标题与正文文字内容被插空格或增删字符
    -3,   # 17 编号与标题文字颜色被改写
    -3,   # 18 未同时指定 a:latin 与 a:ea 为微软雅黑
    -1,   # 19 编号框与标题框文本垂直对齐被改动
    -1,   # 20 底部说明“输出目标：…”规格被改动
    -1,   # 21 底部注释“所有流程框为背景图片…”规格被改动
    -3,   # 22 “临床问题”“科学假说”参照两块规格被改动
    -3,   # 23 其余固定文字规格被改动
    -3,   # 24 页面内23张图片被移动、缩放、替换或删除
]
assert sum(d for d in _DELTAS_ORDER if d > 0) == MAX_SCORE_TOTAL

# ===== 页面与几何基准（单位 cm，来源为交付版式本身的固定版面尺寸） =====
SLIDE_W_EMU = 7200900
SLIDE_H_EMU = 10184130
EMU_PER_CM = 360000.0
PT_PER_CM = 72.0 / 2.54          # 1cm ≈ 28.3465 磅

# 微软雅黑 (ascender+descender)/em ≈ 1.32，即“单倍行距”的行高倍数；
# 中日韩字形实际着墨高度约 1.05em，用于判断上下行是否“粘连压字”。
LINE_H_EM = 1.32
INK_H_EM = 1.05
POS_TOL_CM = 0.06                # 位置/尺寸比对容差
SIZE_TOL_PT = 0.05               # 字号比对容差
GAP_MIN_CM = 0.05                # 标题行与首条正文之间要求的可见间隙
METRIC_TOL_CM = 0.03             # 字形度量为估算值，比对间隙时给出的容差

# 六个编号板块的板块框（左、上、右、下），与细则中标注的框线坐标一致
_BLOCK_FRAMES = {
    "01": (1.422, 12.497, 9.093, 15.240),
    "02": (1.422, 15.875, 9.093, 18.618),
    "03": (1.422, 19.253, 9.093, 21.996),
    "04": (10.795, 12.497, 18.466, 15.240),
    "05": (10.795, 15.875, 18.466, 18.618),
    "06": (10.795, 19.253, 18.466, 21.996),
}

# 板块之间的 3 个箭头图标（左、上、右、下）
_ARROW_RECTS = [
    (9.068, 13.411, 10.465, 13.970),
    (9.068, 16.789, 10.465, 17.348),
    (9.068, 20.168, 10.465, 20.727),
]

# 六个编号板块应有的标题文字、4 条正文条目文字、编号与标题的文字颜色
_BLOCK_SPECS = {
    "01": (
        "研究内容一：关键病理事件界定",
        ["临床分层与样本库建立",
         "LH/hCG刺激模型构建",
         "黄体化基因：STAR / CYP11A1 / LHCGR",
         "排卵执行因子：COX2 / PTX3 / HAS2"],
        "0F5B5B",
    ),
    "02": (
        "研究内容二：色氨酸代谢异常定位",
        ["卵泡液与颗粒细胞代谢组检测",
         "筛选差异代谢物并锁定关键节点",
         "重点观察：IA及相关吲哚类代谢物",
         "建立代谢物—黄体化表型关联"],
        "253E83",
    ),
    "03": (
        "研究内容三：ATF4核内调控解析",
        ["ATF4表达、磷酸化与核定位检测",
         "核/胞分离与免疫荧光验证",
         "ChIP-qPCR / CUT&Tag定位靶基因",
         "解析“代谢异常—核内转录”桥梁"],
        "55328D",
    ),
    "04": (
        "研究内容四：表观遗传持续化机制",
        ["H3K4me3与H3K27ac开放状态检测",
         "关键位点：STAR / COX2 / PTX3",
         "刺激撤除后恢复观察",
         "判断程序异常是否具备持续化特征"],
        "8A4C10",
    ),
    "05": (
        "研究内容五：轴向因果验证",
        ["IA梯度干预与时间窗干预",
         "ATF4沉默/过表达与救援实验",
         "表观修饰干预与逆转观察",
         "形成“IA—ATF4—表观遗传”证据链"],
        "2D6C34",
    ),
    "06": (
        "预期结果与科学意义",
        ["揭示PCOS持续性排卵障碍的新病理机制",
         "建立“代谢异常—核内调控—程序启动失败”框架",
         "提出可分层评估的促排反应预测标志",
         "为靶向代谢—表观干预提供实验依据"],
        "8B274B",
    ),
}
_BLOCK_NUMS = ["01", "02", "03", "04", "05", "06"]

# “微软雅黑”在文件里可能以中文名或英文名书写，二者是同一款字体
_YAHEI_NAMES = ("微软雅黑", "Microsoft YaHei")

# 底部两段说明文字的规格：
# (标签, 文字, 字号磅, 是否加粗, 颜色, 段落对齐, 文本垂直对齐, 左cm, 上cm, 宽cm, 高cm)
_NOTE_MAIN = ("底部说明", "输出目标：从临床异质性出发，锁定可验证、可干预、可转化的排卵障碍机制节点。",
              9.3, False, "465C67", "ctr", "ctr", 1.829, 25.451, 16.383, 0.914)
_NOTE_SUB = ("底部注释", "所有流程框为背景图片；研究内容以可编辑文本框单独叠加，并故意略大于底层流程框。",
             6.8, False, "83909A", "ctr", "ctr", 1.829, 26.492, 16.383, 0.660)

# 作为版式参照的“临床问题”“科学假说”两块文字规格（同上元组结构，垂直对齐留 None 表示不校验）
_REF_TEXT_SPECS = [
    ("编号A", "A", 12.0, True, "17375A", "ctr", None, 11.049, 4.726, None, None),
    ("标题临床问题", "临床问题", 12.2, True, "17375A", "l", None, 11.989, 4.726, None, None),
    ("编号B", "B", 12.0, True, "6B351E", "ctr", None, 1.675, 5.870, None, None),
    ("标题科学假说", "科学假说", 12.2, True, "6B351E", "l", None, 2.515, 5.818, None, None),
]
# “临床问题”3 条要点（同一文本框内 3 段）与“科学假说”正文段
_REF_BULLETS = ["• PCOS患者排卵窗口延迟，卵泡成熟与激素峰值不同步",
                "• 体重、胰岛素抵抗改善后，周期性排卵仍不足",
                "• 同一促排方案下卵泡反应差异大，缺少可解释的标志"]
_REF_BULLET_BOX = (8.6, "27384A", 11.303, 5.334, 6.334, 1.732)
_REF_HYPO_TEXT = ("色氨酸代谢失衡可能触发 ATF4 依赖的核内应激转录网络，抬高颗粒细胞黄体化启动阈值；"
                  "若表观开放状态被持续重塑，排卵执行程序将长期偏离并形成促排反应异质性。")
_REF_HYPO_BOX = (8.6, "3A2B20", 1.838, 6.452, 7.269, 2.357)
# 两块底图（左、上、宽、高）
_REF_PIC_RECTS = [(10.795, 4.521, 7.163, 2.997), (1.448, 5.436, 7.925, 3.200)]

# 六个编号板块与“临床问题”“科学假说”之外的固定文字规格
_FIXED_TEXT_SPECS = [
    ("主标题", "技术路线图", 18.0, True, "1A4050", None, None, 1.067, 0.559, None, None),
    ("副标题", "PCOS持续性排卵障碍机制研究", 11.5, False, "3C5263", None, None, 1.067, 1.600, None, None),
    ("右上导语", "代谢异常 → 核内转录调控 → 表观遗传持续化 → 排卵执行偏离",
     10.2, False, "52636D", "r", None, 6.502, 0.660, None, None),
    ("节点主标签1", "提出问题", 8.7, True, "203444", "ctr", None, None, None, None, None),
    ("节点主标签2", "界定问题", 8.7, True, "203444", "ctr", None, None, None, None, None),
    ("节点主标签3", "解析问题", 8.7, True, "203444", "ctr", None, None, None, None, None),
    ("节点主标签4", "验证机制", 8.7, True, "203444", "ctr", None, None, None, None, None),
    ("节点主标签5", "凝练意义", 8.7, True, "203444", "ctr", None, None, None, None, None),
    ("节点副标签1", "临床表型归纳", 6.8, False, "52636D", "ctr", None, None, None, None, None),
    ("节点副标签2", "黄体化启动时窗", 6.8, False, "52636D", "ctr", None, None, None, None, None),
    ("节点副标签3", "代谢—转录耦合", 6.8, False, "52636D", "ctr", None, None, None, None, None),
    ("节点副标签4", "轴向干预救援", 6.8, False, "52636D", "ctr", None, None, None, None, None),
    ("节点副标签5", "靶点与策略", 6.8, False, "52636D", "ctr", None, None, None, None, None),
    ("小节标题", "研究方法与技术路线", 14.0, True, "1E4352", None, None, 1.270, 11.303, 7.620, 0.838),
    ("中部横条文字", "核心判断：颗粒细胞黄体化启动阈值异常，是持续性排卵障碍由“代谢扰动”走向“程序失序”的关键节点。",
     9.3, True, "24435A", "ctr", None, 1.092, 9.627, None, None),
    ("下部横条文字", "因果链验证：IA干预 · ATF4干预 · 表观修饰干预 · 救援实验 · 表型恢复观察",
     9.4, True, "263D55", "ctr", None, 1.067, 23.393, None, None),
]

# 页面内 23 张图片应有的 (左cm, 上cm, 宽cm, 高cm)
_PIC_RECTS = [
    (0.000, 0.000, 20.003, 28.289),                                     # 满页背景底图
    (1.168, 2.667, 2.921, 1.372), (4.775, 2.667, 2.921, 1.372),         # 顶部5个节点框图
    (8.382, 2.667, 2.921, 1.372), (11.989, 2.667, 2.921, 1.372),
    (15.596, 2.667, 2.921, 1.372),
    (3.886, 3.048, 0.940, 0.406), (7.493, 3.048, 0.940, 0.406),         # 节点间4个箭头图标
    (11.100, 3.048, 0.940, 0.406), (14.707, 3.048, 0.940, 0.406),
    (10.795, 4.521, 7.163, 2.997), (1.448, 5.436, 7.925, 3.200),        # “临床问题”“科学假说”底图
    (1.473, 9.195, 17.221, 1.829),                                      # 中部横条图
    (1.422, 12.497, 7.671, 2.743), (1.422, 15.875, 7.671, 2.743),       # 六个编号板块底框图
    (1.422, 19.253, 7.671, 2.743), (10.795, 12.497, 7.671, 2.743),
    (10.795, 15.875, 7.671, 2.743), (10.795, 19.253, 7.671, 2.743),
    (9.068, 13.411, 1.397, 0.559), (9.068, 16.789, 1.397, 0.559),       # 板块间3个箭头图标
    (9.068, 20.168, 1.397, 0.559),
    (1.524, 22.911, 17.120, 1.880),                                     # 底部横条图
]
assert len(_PIC_RECTS) == 23


# ============================ 基础度量工具 ============================

def _cm(emu):
    if emu is None:
        return None
    return float(emu) / EMU_PER_CM


# 全角/等宽字符区间：中日韩文字、中日韩标点、全角形式等，占 1.0 em；
# 半角拉丁字母、数字、空格、ASCII 标点占约 0.5 em。
_WIDE_RANGES = (
    (0x1100, 0x115F), (0x2E80, 0xA4CF), (0xA960, 0xA97F), (0xAC00, 0xD7A3),
    (0xF900, 0xFAFF), (0xFE10, 0xFE19), (0xFE30, 0xFE6F),
    (0xFF00, 0xFF60), (0xFFE0, 0xFFE6),
)
_WIDE_EXTRA = set("—–…‘’“”·→←↑↓")


def _char_width_em(ch):
    if ch == "\t":
        return 2.0
    if ch in _WIDE_EXTRA:
        return 1.0
    o = ord(ch)
    for lo, hi in _WIDE_RANGES:
        if lo <= o <= hi:
            return 1.0
    return 0.5


def _visual_width_em(text):
    return sum(_char_width_em(c) for c in text)


def _norm_space(text):
    """把连续空白折叠为单个半角空格并去首尾，用于“字面内容是否一致”的宽松比对。"""
    return re.sub(r"\s+", " ", (text or "")).strip()


def _body_pr(shape):
    try:
        return shape.text_frame._txBody.find(qn("a:bodyPr"))
    except Exception:
        return None


def _anchor_of(shape):
    bp = _body_pr(shape)
    val = bp.get("anchor") if bp is not None else None
    return val or "t"          # OOXML 默认顶端对齐


def _insets_cm(shape):
    bp = _body_pr(shape)
    d = {"l": 91440, "t": 45720, "r": 91440, "b": 45720}   # OOXML 默认内边距
    if bp is not None:
        for key, attr in (("l", "lIns"), ("t", "tIns"), ("r", "rIns"), ("b", "bIns")):
            raw = bp.get(attr)
            if raw is not None:
                try:
                    d[key] = int(raw)
                except ValueError:
                    pass
    return {k: v / EMU_PER_CM for k, v in d.items()}


def _autofit(shape):
    """返回 (字号缩放比例, 行距缩减比例)，对应 <a:normAutofit fontScale lnSpcReduction>。"""
    bp = _body_pr(shape)
    na = bp.find(qn("a:normAutofit")) if bp is not None else None
    if na is None:
        return 1.0, 0.0
    fs, lr = na.get("fontScale"), na.get("lnSpcReduction")
    scale = int(fs) / 100000.0 if fs else 1.0
    red = int(lr) / 100000.0 if lr else 0.0
    return scale, red


def _run_size_pt(run, scale=1.0):
    """字号 = 显式 sz × 自动缩放比例；未显式设置时按 PPT 默认 18 磅兜底。"""
    try:
        if run.font.size is not None:
            return float(run.font.size.pt) * scale
    except Exception:
        pass
    return 18.0 * scale


def _run_bold(run):
    try:
        return run.font.bold is True
    except Exception:
        return False


def _run_color(run):
    """返回 6 位大写十六进制 RGB；非 srgbClr（主题色/继承）返回 None。"""
    try:
        rPr = run._r.find(qn("a:rPr"))
        if rPr is None:
            return None
        fill = rPr.find(qn("a:solidFill"))
        if fill is None:
            return None
        clr = fill.find(qn("a:srgbClr"))
        if clr is None:
            return None
        return (clr.get("val") or "").upper()
    except Exception:
        return None


def _run_typeface(run, tag):
    """读取 <a:rPr> 下 a:latin / a:ea 的 typeface；未设置返回 None。"""
    try:
        rPr = run._r.find(qn("a:rPr"))
        if rPr is None:
            return None
        el = rPr.find(qn(tag))
        return el.get("typeface") if el is not None else None
    except Exception:
        return None


def _is_yahei(name):
    return name is not None and name.strip() in _YAHEI_NAMES


def _para_props(para):
    """段落对齐与行距/段间距：lnSpc 支持百分比与固定磅值，spcBef/spcAft 折算为磅。"""
    out = {"algn": None, "ln_pct": None, "ln_pts": None, "spc_bef": 0.0, "spc_aft": 0.0}
    pPr = para._p.find(qn("a:pPr"))
    if pPr is None:
        return out
    out["algn"] = pPr.get("algn")
    ln = pPr.find(qn("a:lnSpc"))
    if ln is not None and len(ln):
        child = ln[0]
        val = child.get("val")
        if val is not None:
            if child.tag == qn("a:spcPct"):
                out["ln_pct"] = int(val) / 100000.0
            elif child.tag == qn("a:spcPts"):
                out["ln_pts"] = int(val) / 100.0
    for key, tag in (("spc_bef", "a:spcBef"), ("spc_aft", "a:spcAft")):
        el = pPr.find(qn(tag))
        if el is not None and len(el):
            child = el[0]
            val = child.get("val")
            if val is not None and child.tag == qn("a:spcPts"):
                out[key] = int(val) / 100.0
    return out


def _layout(shape):
    """
    按“办公软件里看到的样子”估算文本框内文字的实际排布，而不是用文本框矩形代替文字：
      · 每段行数 = ceil(该段视觉字宽 / 可用行宽)，按 wrap 折行；
      · 行高 = 字号 × 1.32 × 行距百分比（微软雅黑单倍行距）；
      · 竖直起点由 anchor 决定（t=顶端+上内边距，ctr=框内垂直居中，b=底端-下内边距）；
      · 每行左端由段落对齐 algn 决定（l/None 左对齐、ctr 居中、r 右对齐）。
    返回 None 表示该形状没有文字。
    """
    if not shape.has_text_frame:
        return None
    tf = shape.text_frame
    if not (tf.text or "").strip():
        return None
    scale, ln_red = _autofit(shape)
    ins = _insets_cm(shape)
    box_l, box_t = _cm(shape.left), _cm(shape.top)
    box_w, box_h = _cm(shape.width), _cm(shape.height)
    avail_w = max(box_w - ins["l"] - ins["r"], 0.05)

    paras, total_h = [], 0.0
    for para in tf.paragraphs:
        runs = [r for r in para.runs if r.text]
        text = "".join(r.text for r in para.runs)
        size_pt = max([_run_size_pt(r, scale) for r in runs]) if runs else 18.0 * scale
        em_cm = size_pt / PT_PER_CM
        pp = _para_props(para)
        pct = max((pp["ln_pct"] if pp["ln_pct"] is not None else 1.0) - ln_red, 0.1)
        adv = pp["ln_pts"] / PT_PER_CM if pp["ln_pts"] else size_pt * LINE_H_EM * pct / PT_PER_CM
        units = _visual_width_em(text)
        cap = max(avail_w / em_cm, 1.0) if em_cm > 0 else 1.0
        n_lines = 1 if not text.strip() else max(1, int(math.ceil(units / cap - 1e-9)))
        info = {"text": text, "runs": runs, "size_pt": size_pt, "lines": n_lines,
                "adv": adv, "ink_h": size_pt * INK_H_EM / PT_PER_CM,
                "width": min(units, cap) * em_cm, "algn": pp["algn"],
                "ln_pct": pct, "ln_pts": pp["ln_pts"],
                "spc_bef": pp["spc_bef"], "spc_aft": pp["spc_aft"]}
        paras.append(info)
        total_h += (pp["spc_bef"] + pp["spc_aft"]) / PT_PER_CM + n_lines * adv

    anchor = _anchor_of(shape)
    if anchor == "ctr":
        y = box_t + (box_h - total_h) / 2.0
    elif anchor == "b":
        y = box_t + box_h - ins["b"] - total_h
    else:
        y = box_t + ins["t"]

    lines = []
    for p in paras:
        y += p["spc_bef"] / PT_PER_CM
        for _ in range(p["lines"]):
            if p["algn"] == "ctr":
                lx = box_l + ins["l"] + (avail_w - p["width"]) / 2.0
            elif p["algn"] == "r":
                lx = box_l + box_w - ins["r"] - p["width"]
            else:
                lx = box_l + ins["l"]
            lines.append({"top": y, "bottom": y + p["adv"], "left": lx,
                          "right": lx + p["width"], "para": p})
            y += p["adv"]
        y += p["spc_aft"] / PT_PER_CM
    return {"shape": shape, "paras": paras, "lines": lines, "anchor": anchor,
            "top": lines[0]["top"], "bottom": lines[-1]["bottom"],
            "left": min(l["left"] for l in lines),
            "right": max(l["right"] for l in lines)}


def _rect_of(shape):
    l, t = _cm(shape.left), _cm(shape.top)
    return (l, t, l + _cm(shape.width), t + _cm(shape.height))


def _overlap(a, b):
    """两矩形是否有实际重叠面积（(左,上,右,下) 口径）。"""
    return a[0] < b[2] and b[0] < a[2] and a[1] < b[3] and b[1] < a[3]


def _dist_to_rect(cx, cy, rect):
    dx = max(rect[0] - cx, 0.0, cx - rect[2])
    dy = max(rect[1] - cy, 0.0, cy - rect[3])
    return math.hypot(dx, dy)


def _collect_text_items(slide):
    """采集全部承载文字的形状，附带渲染排布与常用派生信息。"""
    items = []
    for sh in slide.shapes:
        lay = _layout(sh)
        if lay is None:
            continue
        text = sh.text_frame.text
        rect = _rect_of(sh)
        items.append({
            "shape": sh, "layout": lay, "text": text,
            "flat": _norm_space(text.replace("\n", " ")),
            "paras_text": [p["text"] for p in lay["paras"]],
            "rect": rect, "cx": (rect[0] + rect[2]) / 2.0, "cy": (rect[1] + rect[3]) / 2.0,
            "runs": [r for p in lay["paras"] for r in p["runs"]],
        })
    return items


def _split_blocks(text_items):
    """
    把六个编号板块的文字形状归位并分层。兼容两种常见做法：
      · 编号单独一个文本框 + 标题一个文本框 + 正文一个文本框；
      · 编号并入标题文本框（“01 研究内容一：…”，编号为首个 run），编号框被清空/删除。
    """
    assigned = {n: [] for n in _BLOCK_NUMS}
    for it in text_items:
        inside = [n for n in _BLOCK_NUMS
                  if _BLOCK_FRAMES[n][0] <= it["cx"] <= _BLOCK_FRAMES[n][2]
                  and _BLOCK_FRAMES[n][1] <= it["cy"] <= _BLOCK_FRAMES[n][3]]
        if len(inside) == 1:
            assigned[inside[0]].append(it)
            continue
        cands = inside or _BLOCK_NUMS
        best = min(cands, key=lambda n: _dist_to_rect(it["cx"], it["cy"], _BLOCK_FRAMES[n]))
        if _dist_to_rect(it["cx"], it["cy"], _BLOCK_FRAMES[best]) <= 0.6:
            assigned[best].append(it)

    blocks = {}
    for num in _BLOCK_NUMS:
        cands = assigned[num]
        body = None
        multi = [c for c in cands if len(c["paras_text"]) >= 2]
        if multi:
            body = max(multi, key=lambda c: len(c["paras_text"]))
        badge = title = None
        for c in cands:
            if c is body:
                continue
            if re.fullmatch(r"0[1-6]", c["flat"]):
                if badge is None:
                    badge = c
            elif title is None or len(c["flat"]) > len(title["flat"]):
                title = c
        merged = bool(title and re.match(r"^\s*" + num + r"\b", title["flat"]))
        blocks[num] = {"all": cands, "badge": badge, "title": title,
                       "body": body, "merged_number": merged}
    return blocks


def _title_text_raw(blk, num):
    """剥掉可能并入标题框的编号前缀后的纯标题文字，逐字保留（空格不折叠）。"""
    if blk["title"] is None:
        return ""
    txt = (blk["title"]["text"] or "").replace("\n", "").strip()
    if blk["merged_number"]:
        txt = re.sub(r"^" + num + r"\s*", "", txt)
    return txt.strip()


def _title_runs_of(blk):
    """标题文本框中真正承载标题文字的 run（排除只写了编号的 run）。"""
    if blk["title"] is None:
        return []
    out = [r for r in blk["title"]["runs"]
           if r.text.strip() and not re.fullmatch(r"0[1-6]", r.text.strip())]
    return out or list(blk["title"]["runs"])


def _number_line_of(blk):
    """编号所在的那一行（独立编号框 → 其首行；并入标题框 → 标题首行）。"""
    src = blk["badge"] or blk["title"]
    if src is None:
        return None
    return src["layout"]["lines"][0]


def _title_line_of(blk):
    if blk["title"] is None:
        return None
    return blk["title"]["layout"]["lines"][0]


def _find_by_text(text_items, text):
    """按“折叠空白后的字面内容”定位文本框，找不到返回 None。"""
    target = _norm_space(text)
    for it in text_items:
        if it["flat"] == target:
            return it
    return None


def _check_text_spec(text_items, spec):
    """
    校验一处固定文字的字体族/字号/加粗/颜色/段落对齐/文本垂直对齐/位置/尺寸。
    返回不合规原因列表，空列表表示完全一致。
    """
    label, text, size_pt, bold, color, algn, anchor, left, top, width, height = spec
    it = _find_by_text(text_items, text)
    if it is None:
        return ["%s文字缺失或被改写" % label]
    bad = []
    runs = [r for r in it["runs"] if r.text.strip()]
    scale, _ = _autofit(it["shape"])
    if not runs:
        return ["%s文字缺失" % label]
    for r in runs:
        if not (_is_yahei(_run_typeface(r, "a:latin")) and _is_yahei(_run_typeface(r, "a:ea"))):
            bad.append("%s字体非微软雅黑" % label)
            break
    if any(abs(_run_size_pt(r, scale) - size_pt) > SIZE_TOL_PT for r in runs):
        bad.append("%s字号非%s磅（实为%s）"
                   % (label, size_pt, round(_run_size_pt(runs[0], scale), 2)))
    if any(_run_bold(r) != bold for r in runs):
        bad.append("%s加粗状态不符" % label)
    if any(_run_color(r) != color for r in runs):
        bad.append("%s颜色非%s（实为%s）" % (label, color, _run_color(runs[0])))
    if algn is not None:
        got = it["layout"]["paras"][0]["algn"] or "l"
        if got != algn:
            bad.append("%s段落对齐非%s（实为%s）" % (label, algn, got))
    if anchor is not None and it["layout"]["anchor"] != anchor:
        bad.append("%s文本垂直对齐非%s（实为%s）" % (label, anchor, it["layout"]["anchor"]))
    for name, want, got in (("左", left, _cm(it["shape"].left)),
                            ("上", top, _cm(it["shape"].top)),
                            ("宽", width, _cm(it["shape"].width)),
                            ("高", height, _cm(it["shape"].height))):
        if want is not None and abs(got - want) > POS_TOL_CM:
            bad.append("%s%s为%.3fcm（应%.3fcm）" % (label, name, got, want))
    return bad


def _check_para_box(text_items, label, want_paras, box_spec):
    """
    校验一个多段落文本框（“临床问题”3条要点 / “科学假说”正文段）：
    逐段文字、字号、颜色、字体族，以及文本框的左/上/宽/高。
    """
    size_pt, color, left, top, width, height = box_spec
    target = _norm_space(want_paras[0])
    it = None
    for cand in text_items:
        if cand["paras_text"] and _norm_space(cand["paras_text"][0]) == target:
            it = cand
            break
    if it is None:
        return ["%s文字缺失或被改写" % label]
    bad = []
    got = [t.strip() for t in it["paras_text"] if t.strip()]
    if got != [t.strip() for t in want_paras]:
        bad.append("%s共%d段、文字内容与原规格不一致" % (label, len(got)))
    scale, _ = _autofit(it["shape"])
    runs = [r for r in it["runs"] if r.text.strip()]
    if any(not (_is_yahei(_run_typeface(r, "a:latin")) and _is_yahei(_run_typeface(r, "a:ea")))
           for r in runs):
        bad.append("%s字体非微软雅黑" % label)
    # 此处按“字号框里设置的磅值”比对（不乘 normAutofit 缩放）：原版式本身就
    # 对要点框用了 90% 自动缩小，缩放比例属于原有设置的一部分。
    if any(abs(_run_size_pt(r, 1.0) - size_pt) > SIZE_TOL_PT for r in runs):
        bad.append("%s字号非%s磅（实为%s）" % (label, size_pt, round(_run_size_pt(runs[0], 1.0), 2)))
    if any(_run_color(r) != color for r in runs):
        bad.append("%s颜色非%s（实为%s）" % (label, color, _run_color(runs[0]) if runs else None))
    for name, want, real in (("左", left, _cm(it["shape"].left)),
                             ("上", top, _cm(it["shape"].top)),
                             ("宽", width, _cm(it["shape"].width)),
                             ("高", height, _cm(it["shape"].height))):
        if want is not None and abs(real - want) > POS_TOL_CM:
            bad.append("%s%s为%.3fcm（应%.3fcm）" % (label, name, real, want))
    return bad


class _TargetFileMissing(Exception):
    """待评估文件根本不存在（目录缺失或目录内没有 .pptx），区别于“文件在但打不开”。"""


def evaluate(dir_path):
    """入参为脚本所在目录；在该目录内定位待评估的 .pptx，返回评测结果 dict。"""
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
        if not os.path.isdir(dir_path):
            raise _TargetFileMissing("目录不存在：%s" % dir_path)
        cands = [n for n in sorted(os.listdir(dir_path))
                 if n.lower().endswith(".pptx") and not n.startswith("~$")]
        if not cands:
            raise _TargetFileMissing("目录中未找到 .pptx 文件：%s" % dir_path)
        file_name = cands[0]
        result["file_name"] = file_name
        _evaluate_impl(os.path.join(dir_path, file_name), result)
    except _TargetFileMissing as exc:
        # 没有待评估文件 → 维度一无从判定：dim1_pass / dim1_reason 均留空，
        # 且不给出满分基数（max_score=0），避免被计入“得0分”的样本。
        result["status"] = "error"
        result["error"] = "%s" % exc
        result["dim1_pass"] = ""
        result["dim1_reason"] = ""
        result["dim2_items"] = []
        result["total_score"] = 0
        result["max_score"] = 0
    except Exception as exc:
        # 文件存在但无法作为 .pptx 打开/解析 → 属于交付文件问题，
        # 记为维度一不通过、正常 0 分，而不是脚本运行错误。
        result["status"] = "ok"
        result["error"] = None
        result["dim1_pass"] = False
        result["dim1_reason"] = ("文件未能作为PPT演示文稿（.pptx）正常打开并完成解析：%s: %s"
                                % (type(exc).__name__, exc))
        result["dim2_items"] = []
        result["total_score"] = 0
        result["max_score"] = MAX_SCORE_TOTAL
    return result


# 维度二 24 个评分项的细则原文（顺序与 _DELTAS_ORDER 一一对应）
_RULE_TEXTS = (
    # 01—06 加分项：各编号板块的文字是否被放进了板块框内
    "编号01“研究内容一：关键病理事件界定”板块的标题行与4条正文（临床分层与样本库建立／LH/hCG刺激模型构建／黄体化基因：STAR / CYP11A1 / LHCGR／排卵执行因子：COX2 / PTX3 / HAS2）全部落在该板块框（横1.42—9.09cm、纵12.50—15.24cm）内，末行文字不越出框下边线、行尾不越出框右边线",
    "编号02“研究内容二：色氨酸代谢异常定位”板块的标题行与4条正文（卵泡液与颗粒细胞代谢组检测／筛选差异代谢物并锁定关键节点／重点观察：IA及相关吲哚类代谢物／建立代谢物—黄体化表型关联）全部落在该板块框（横1.42—9.09cm、纵15.88—18.62cm）内，末行文字不越出框下边线、行尾不越出框右边线",
    "编号03“研究内容三：ATF4核内调控解析”板块的标题行与4条正文（ATF4表达、磷酸化与核定位检测／核/胞分离与免疫荧光验证／ChIP-qPCR / CUT&Tag定位靶基因／解析“代谢异常—核内转录”桥梁）全部落在该板块框（横1.42—9.09cm、纵19.25—21.99cm）内，末行文字不越出框下边线、行尾不越出框右边线",
    "编号04“研究内容四：表观遗传持续化机制”板块的标题行与4条正文（H3K4me3与H3K27ac开放状态检测／关键位点：STAR / COX2 / PTX3／刺激撤除后恢复观察／判断程序异常是否具备持续化特征）全部落在该板块框（横10.79—18.46cm、纵12.50—15.24cm）内，末行文字不越出框下边线、行尾不越出框右边线",
    "编号05“研究内容五：轴向因果验证”板块的标题行与4条正文（IA梯度干预与时间窗干预／ATF4沉默/过表达与救援实验／表观修饰干预与逆转观察／形成“IA—ATF4—表观遗传”证据链）全部落在该板块框（横10.79—18.46cm、纵15.88—18.62cm）内，末行文字不越出框下边线、行尾不越出框右边线",
    "编号06“预期结果与科学意义”板块的标题行与4条正文（揭示PCOS持续性排卵障碍的新病理机制／建立“代谢异常—核内调控—程序启动失败”框架／提出可分层评估的促排反应预测标志／为靶向代谢—表观干预提供实验依据）全部落在该板块框（横10.79—18.46cm、纵19.25—21.99cm）内，末行文字不越出框下边线、行尾不越出框右边线",
    # 07 加分项：版式是否照参照两块改过来了
    "六个编号板块的版式调整为与“临床问题”“科学假说”两块一致：编号位于板块左上角，板块标题与编号同一水平行、左对齐紧随编号之后，4条正文自标题下方起左对齐纵向排列，正文左端相对板块框左边线内缩0.3—1.0cm",
    # 08—15 减分项：调整过程中把原本已经成立的地方弄坏
    "六个编号板块的文字出现叠压：块内“编号—板块标题—4条正文”三层元素相互压字，标题行底部与首条正文顶部之间没有可见间隙（小于0.05cm），或4条正文上下行字形交叠粘连；或编号、标题与正文文字压到板块之间的3个箭头图标（左端9.07cm、宽1.40cm、顶端分别13.41/16.79/20.17cm）或相邻板块框上",
    "页面尺寸不再为宽20.0cm×高28.29cm（7200900×10184130 EMU），被改为16:9、A4或其他规格",
    "六个编号板块的正文字号不统一为同一磅值，六块之间出现正文字号不一（如7.25磅与7.35磅并存）",
    "六个编号板块的标题字号不统一为同一磅值，六块之间出现标题字号不一（如其余为10磅而某一块为9.6磅）",
    "六个编号板块的正文字号低于8磅，被压到7.5磅以下导致文字过密难辨",
    "六个编号板块的标题字号低于10磅，或未保持加粗",
    "六个编号板块的正文行距超过单倍行距，或段前段后间距不为0磅",
    "六个编号板块的文字内容不完整：编号01、02、03、04、05、06未全部保留可见，被删除、清空或被其他元素遮挡；或正文未保持4条条目、每条独立成段，出现删减条目内容、合并为整段、折行拆成5行及以上",
    # 16—24 减分项：指令未涉及、却被顺手改坏的地方
    "六个编号板块的标题与正文文字内容不满足“研究内容三：ATF4核内调控解析”“ChIP-qPCR / CUT&Tag定位靶基因”“H3K4me3与H3K27ac开放状态检测”“IA梯度干预与时间窗干预”“ATF4沉默/过表达与救援实验”（中英文之间及“/”两侧均不插入空格，全部文字不增删任何字符）",
    "六个编号板块的编号与标题文字颜色不满足01为0F5B5B、02为253E83、03为55328D、04为8A4C10、05为2D6C34、06为8B274B",
    "六个编号板块的编号、标题与正文文字不满足同时将拉丁文字体与东亚文字字体（a:latin 与 a:ea）指定为“微软雅黑”，出现仅保留拉丁字体名“Microsoft YaHei”而丢失东亚文字字体设置、致中文回落到主题默认东亚字体",
    "六个编号板块的编号框与标题框不满足文本垂直对齐为垂直居中（anchor=\"ctr\"）",
    "页面底部说明“输出目标：从临床异质性出发，锁定可验证、可干预、可转化的排卵障碍机制节点。”不满足微软雅黑 9.3磅 非加粗 颜色465C67、段落居中、文本垂直居中，且文本框位于左1.83cm/上25.45cm、尺寸16.38×0.91cm",
    "页面底部注释“所有流程框为背景图片；研究内容以可编辑文本框单独叠加，并故意略大于底层流程框。”不满足微软雅黑 6.8磅 颜色83909A、段落居中、文本垂直居中，且文本框位于左1.83cm/上26.49cm、尺寸16.38×0.66cm",
    "作为参照样式的“临床问题”“科学假说”两块不满足原有规格：两块底图位置尺寸、编号“A”“B”与标题“临床问题”“科学假说”的字号加粗颜色位置、“临床问题”3条要点与“科学假说”正文段的字号颜色位置尺寸",
    "页面上六个编号板块与“临床问题”“科学假说”之外的固定文字不满足原有规格：主标题、副标题、右上导语、顶部5个节点主标签与5个副标签、小节标题、中部横条文字、下部横条文字的字体字号加粗颜色对齐与位置尺寸",
    "页面内23张图片（满页背景底图、顶部5个节点框图与其间4个箭头图标、“临床问题”“科学假说”两块底图、中部横条图、六个编号板块底框图、板块间3个箭头图标、底部横条图）不满足全部保留且位置与尺寸不发生移动、缩放、替换或删除",
)
assert len(_RULE_TEXTS) == len(_DELTAS_ORDER)


def _evaluate_impl(file_path, result):
    prs = Presentation(file_path)
    results = []

    def emit(idx, ok, detail=""):
        """按固定顺序追加一条评分结果；加分项写 '+N/+0'，减分项写 '-0/-N'。
        detail 为不合规原因，仅供本地调试查看，不写入返回结果。"""
        if len(results) != idx:
            raise RuntimeError("dim2 追加顺序错位：期望第%d项，实际已有%d项" % (idx, len(results)))
        delta = _DELTAS_ORDER[idx]
        if delta > 0:
            results.append("+%d：%s" % (delta if ok else 0, _RULE_TEXTS[idx]))
        else:
            results.append("-%d：%s" % (0 if ok else -delta, _RULE_TEXTS[idx]))

    # ======================= 维度一：可用与可修改性 =======================
    # 与办公软件对齐：只判“这个文件双击后能不能作为 PPT 演示文稿正常打开”。
    # 文件损坏、格式错误、无法解析时，Presentation() 会直接抛异常，由 evaluate()
    # 统一兜到 status="error" / dim1_pass=False / total_score=0。
    page_w = _cm(prs.slide_width)
    page_h = _cm(prs.slide_height)
    if len(prs.slides) < 1:
        raise RuntimeError("演示文稿内没有任何幻灯片，无法作为交付文件评估")
    slide = prs.slides[0]
    text_items = _collect_text_items(slide)
    blocks = _split_blocks(text_items)
    result["dim1_pass"] = True
    result["dim1_reason"] = "交付文件为1个可正常打开的PPT演示文稿（.pptx）"

    # 六块的正文/标题 run 集合，后续多条细则复用
    body_sizes = {}
    title_sizes = {}
    for num in _BLOCK_NUMS:
        blk = blocks[num]
        if blk["body"] is not None:
            scale, _ = _autofit(blk["body"]["shape"])
            sizes = [round(_run_size_pt(r, scale), 2)
                     for r in blk["body"]["runs"] if r.text.strip()]
            body_sizes[num] = sizes
        if blk["title"] is not None:
            scale, _ = _autofit(blk["title"]["shape"])
            title_sizes[num] = [round(_run_size_pt(r, scale), 2) for r in _title_runs_of(blk)]

    # ============================ 维度二：完成度 ============================
    # 加分项只放“指令明确下达、且输入文件本身并不满足、必须动手改才能做到”的任务：
    # 把六块的文字放进板块框内（6×+3）、把版式改成参照两块的样子（+3）。
    # 输入文件本来就成立的项一律不给加分，改为减分项——它们衡量的是“动手过程中
    # 有没有把原来已经好的地方弄坏”。

    # +3 ×6：编号01—06各板块的标题行与4条正文全部落在板块框内
    # 严格对齐细则，仅踩细则明确写出的点，不额外约束：
    #   ① 末行文字不越出框下边线；② 行尾文字不越出框右边线；
    #   ③ 同时不越出框的上边线与左边线（“全部落在框内”的另一半含义）。
    # 与办公软件对齐：判的是“文字实际渲染出来占的那块地方”——竖直起点由文本
    # 垂直对齐决定，行高按字号×1.32×行距，行数按框宽折行；不是拿文本框矩形
    # 顶替文字。本版式原本就故意让文本框略大于底层流程框，用矩形判会误伤。
    for _i, _num in enumerate(_BLOCK_NUMS):
        _frame = _BLOCK_FRAMES[_num]
        _bad = []
        for _key, _label in (("badge", "编号"), ("title", "标题"), ("body", "正文")):
            _it = blocks[_num][_key]
            if _it is None:
                continue
            _lay = _it["layout"]
            if _lay["bottom"] > _frame[3] + POS_TOL_CM:
                _bad.append("%s末行越出框下边线%.2fcm" % (_label, _lay["bottom"] - _frame[3]))
            if _lay["right"] > _frame[2] + POS_TOL_CM:
                _bad.append("%s行尾越出框右边线%.2fcm" % (_label, _lay["right"] - _frame[2]))
            if _lay["top"] < _frame[1] - POS_TOL_CM:
                _bad.append("%s首行越出框上边线%.2fcm" % (_label, _frame[1] - _lay["top"]))
            if _lay["left"] < _frame[0] - POS_TOL_CM:
                _bad.append("%s行首越出框左边线%.2fcm" % (_label, _frame[0] - _lay["left"]))
        emit(_i, not _bad, "；".join(_bad))

    # +3：六个编号板块的版式调整为与“临床问题”“科学假说”两块一致
    # 严格对齐细则，仅踩细则明确写出的点，不额外约束：
    #   ① 编号位于板块左上角；② 板块标题与编号同一水平行、左对齐紧随编号之后；
    #   ③ 4条正文自标题下方起左对齐纵向排列；
    #   ④ 正文左端相对板块框左边线内缩 0.3—1.0cm。
    # 与办公软件对齐：这就是参照两块“A + 临床问题 + 3条要点”的读法——编号在
    # 左上、标题同行跟在编号后面、要点一条一行缩在标题下方。编号并入标题框写成
    # “01 研究内容一：…”时，①②天然成立，同样算一致。
    _lay_bad = []
    for _num in _BLOCK_NUMS:
        _blk = blocks[_num]
        _frame = _BLOCK_FRAMES[_num]
        _nline, _tline, _body = _number_line_of(_blk), _title_line_of(_blk), _blk["body"]
        if _nline is None or _tline is None or _body is None:
            _lay_bad.append("编号%s板块缺少编号/标题/正文层" % _num)
            continue
        if _nline["left"] > _frame[0] + 1.2 or _nline["top"] > _frame[1] + 0.9:
            _lay_bad.append("编号%s的编号未处于板块左上角（左偏%.2fcm、下偏%.2fcm）"
                            % (_num, _nline["left"] - _frame[0], _nline["top"] - _frame[1]))
        if not _blk["merged_number"]:
            _nc = (_nline["top"] + _nline["bottom"]) / 2.0
            _tc = (_tline["top"] + _tline["bottom"]) / 2.0
            if abs(_nc - _tc) > 0.35:
                _lay_bad.append("编号%s的标题与编号不在同一水平行（错开%.2fcm）"
                                % (_num, abs(_nc - _tc)))
            if _tline["left"] < _nline["left"] - POS_TOL_CM:
                _lay_bad.append("编号%s的标题未紧随编号之后（跑到编号左侧）" % _num)
        if (_tline["para"]["algn"] or "l") != "l":
            _lay_bad.append("编号%s的标题未左对齐" % _num)
        if _body["layout"]["lines"][0]["top"] < max(_tline["bottom"], _nline["bottom"]) - 0.1:
            _lay_bad.append("编号%s的正文未自标题下方起排" % _num)
        if any((_p["algn"] or "l") != "l" for _p in _body["layout"]["paras"]):
            _lay_bad.append("编号%s的正文未左对齐纵向排列" % _num)
        _indent = _body["layout"]["left"] - _frame[0]
        if not (0.30 - 0.02 <= _indent <= 1.00 + 0.02):
            _lay_bad.append("编号%s的正文左端内缩%.2fcm（应0.3—1.0cm）" % (_num, _indent))
    emit(6, not _lay_bad, "；".join(_lay_bad[:4]))

    # -3：六个编号板块的文字出现叠压
    # 严格对齐细则，仅踩细则明确写出的点，不额外约束：
    #   ① 块内“编号—标题—4条正文”三层元素不相互压字；
    #   ② 标题行底部与首条正文顶部之间留有可见间隙（≥0.05cm）；
    #   ③ 4条正文逐行分开、无上下行字形交叠；
    #   ④ 三层文字都没有压到板块间的3个箭头图标或相邻板块框上。
    # 与办公软件对齐：②比的是肉眼能看见的那道白缝，字形度量为估算值，故给
    # 0.03cm 容差；③用“行进距离 ≥ 中日韩字形着墨高度(1.05em)”判是否粘字。
    _ov = []
    for _num in _BLOCK_NUMS:
        _blk = blocks[_num]
        _nline = _number_line_of(_blk)
        _tline = _title_line_of(_blk)
        _body = _blk["body"]
        if _blk["badge"] is not None and _tline is not None:
            _b = _blk["badge"]["layout"]
            if _overlap((_b["left"], _b["top"], _b["right"], _b["bottom"]),
                        (_tline["left"], _tline["top"], _tline["right"], _tline["bottom"])):
                _ov.append("编号%s的编号与标题文字相互压字" % _num)
        if _tline is not None and _body is not None:
            _ref_bot = _tline["bottom"]
            if _nline is not None:
                _ref_bot = max(_ref_bot, _nline["bottom"])
            _gap = _body["layout"]["lines"][0]["top"] - _ref_bot
            if _gap < GAP_MIN_CM - METRIC_TOL_CM:
                _ov.append("编号%s的标题行与首条正文之间仅%.3fcm" % (_num, _gap))
        if _body is not None:
            for _p in _body["layout"]["paras"]:
                if _p["adv"] < _p["ink_h"] - 1e-6:
                    _ov.append("编号%s的正文行进%.3fcm小于字形高度%.3fcm、上下行粘连"
                               % (_num, _p["adv"], _p["ink_h"]))
                    break
        for _key in ("badge", "title", "body"):
            _it = _blk[_key]
            if _it is None:
                continue
            _lay = _it["layout"]
            _r = (_lay["left"], _lay["top"], _lay["right"], _lay["bottom"])
            if any(_overlap(_r, _a) for _a in _ARROW_RECTS):
                _ov.append("编号%s的文字压到板块之间的箭头图标上" % _num)
            for _other in _BLOCK_NUMS:
                if _other != _num and _overlap(_r, _BLOCK_FRAMES[_other]):
                    _ov.append("编号%s的文字压到编号%s板块框上" % (_num, _other))
                    break
    emit(7, not _ov, "；".join(_ov[:4]))

    # -1：页面尺寸不再为宽20.0cm×高28.29cm（7200900×10184130 EMU）
    # 严格对齐细则，仅踩细则明确写出的点，不额外约束：
    #   ① 页面宽 7200900 EMU；② 页面高 10184130 EMU。
    # 与办公软件对齐：即“设计 → 幻灯片大小 → 自定义”里的纸张宽高；指令明确要求
    # 尺寸不变，一旦被改成 16:9 或 A4，整张竖版长图的坐标基准全部失效，所有元素
    # 都得重排。
    size_ok = (abs(prs.slide_width - SLIDE_W_EMU) <= 2000
               and abs(prs.slide_height - SLIDE_H_EMU) <= 2000)
    emit(8, size_ok, "页面尺寸为%.2f×%.2fcm" % (page_w, page_h))

    # -3：六个编号板块的正文字号不统一为同一磅值
    # 严格对齐细则，仅踩细则明确写出的点，不额外约束：
    #   ① 六块正文的实际磅值只有一种（含 normAutofit 字号缩放后的等效磅值）。
    # 与办公软件对齐：选中不同板块正文时“字号”框里读到的数字必须相同；
    # 7.25 磅与 7.35 磅并存这种差值肉眼就是“这几块字明显更小一点”。
    _all_body = sorted({s for _s in body_sizes.values() for s in _s})
    _uni_body = (not _all_body) or (max(_all_body) - min(_all_body) <= SIZE_TOL_PT)
    emit(9, _uni_body, "六块正文出现%s磅并存" % "/".join("%.2f" % s for s in _all_body))

    # -1：六个编号板块的标题字号不统一为同一磅值
    # 严格对齐细则，仅踩细则明确写出的点，不额外约束：
    #   ① 六块板块标题的实际磅值只有一种（编号本身的磅值不在此项内）。
    # 与办公软件对齐：六个标题横排在一页上，其中一块小 0.4 磅即“这行标题矮一截”。
    _all_title = sorted({s for _s in title_sizes.values() for s in _s})
    _uni_title = (not _all_title) or (max(_all_title) - min(_all_title) <= SIZE_TOL_PT)
    emit(10, _uni_title, "六块标题出现%s磅并存" % "/".join("%.2f" % s for s in _all_title))

    # -3：六个编号板块的正文字号低于8磅
    # 严格对齐细则，仅踩细则明确写出的点，不额外约束：
    #   ① 每一块正文的每个 run 磅值 ≥ 8.0 磅。
    # 与办公软件对齐：为了塞进框里而把正文压到 7 磅出头，投屏和打印都糊成一片。
    _min_body = min(_all_body) if _all_body else 0.0
    emit(11, bool(_all_body) and _min_body >= 8.0 - SIZE_TOL_PT,
         "正文最小字号仅%.2f磅" % _min_body)

    # -1：六个编号板块的标题字号低于10磅，或未保持加粗
    # 严格对齐细则，仅踩细则明确写出的点，不额外约束：
    #   ① 每块标题磅值 ≥ 10.0 磅；② 每块标题保持加粗。
    # 与办公软件对齐：标题靠“比正文大 + 加粗”才立得住，掉到 9.6 磅或掉粗
    # 就会和正文糊在一起，板块层级读不出来。
    _min_title = min(_all_title) if _all_title else 0.0
    _title_bold = all(all(_run_bold(r) for r in _title_runs_of(blocks[n]))
                      for n in _BLOCK_NUMS if blocks[n]["title"] is not None)
    _t_bad = []
    if not _all_title or _min_title < 10.0 - SIZE_TOL_PT:
        _t_bad.append("标题最小字号仅%.2f磅" % _min_title)
    if not _title_bold:
        _t_bad.append("存在未加粗的板块标题")
    emit(12, not _t_bad, "；".join(_t_bad))

    # -3：六个编号板块的正文行距超过单倍行距，或段前段后间距不为0磅
    # 严格对齐细则，仅踩细则明确写出的点，不额外约束：
    #   ① 行距百分比 ≤ 100%（或固定磅值 ≤ 字号×1.32）；
    #   ② 段前间距 = 0 磅；③ 段后间距 = 0 磅。
    # 与办公软件对齐：对应“段落”对话框里的行距与段前/段后；这三项一放大，
    # 4 条正文就顶出板块框，是这类长图最常见的溢出源头。
    _ln_bad = []
    for _num in _BLOCK_NUMS:
        _body = blocks[_num]["body"]
        if _body is None:
            continue
        for _p in _body["layout"]["paras"]:
            if _p["ln_pts"] is not None:
                if _p["ln_pts"] > _p["size_pt"] * LINE_H_EM + 0.05:
                    _ln_bad.append("编号%s正文固定行距%.2f磅超过单倍" % (_num, _p["ln_pts"]))
                    break
            elif _p["ln_pct"] > 1.0 + 1e-6:
                _ln_bad.append("编号%s正文行距%.0f%%超过单倍" % (_num, _p["ln_pct"] * 100))
                break
            if _p["spc_bef"] > 0.01 or _p["spc_aft"] > 0.01:
                _ln_bad.append("编号%s正文段前%.2f磅/段后%.2f磅不为0"
                               % (_num, _p["spc_bef"], _p["spc_aft"]))
                break
    emit(13, not _ln_bad, "；".join(_ln_bad))

    # -1：六个编号板块的文字内容不完整
    # 严格对齐细则，仅踩细则明确写出的点，不额外约束：
    #   ① 编号 01—06 均保留可见（独立编号框或并入标题框的编号前缀均可）；
    #   ② 正文保持 4 条条目、每条独立成段；
    #   ③ 未合并为整段、未折行拆成 5 行及以上。
    # 与办公软件对齐：编号被清空、4 条被并成一段、或因框太窄折成 5 行以上，
    # 用户都得回去重新拆段重排。
    _cp_bad = []
    for _num in _BLOCK_NUMS:
        _blk = blocks[_num]
        _has_num = ((_blk["badge"] is not None and _blk["badge"]["flat"] == _num)
                    or _blk["merged_number"])
        if not _has_num:
            _cp_bad.append("编号%s未保留可见编号" % _num)
        _body = _blk["body"]
        if _body is None:
            _cp_bad.append("编号%s板块正文缺失" % _num)
            continue
        _paras = [t for t in _body["paras_text"] if t.strip()]
        if len(_paras) != 4:
            _cp_bad.append("编号%s板块正文为%d段（应4条各自独立成段）" % (_num, len(_paras)))
        _nl = sum(_p["lines"] for _p in _body["layout"]["paras"] if _p["text"].strip())
        if _nl >= 5:
            _cp_bad.append("编号%s板块正文折行后共%d行（≥5行）" % (_num, _nl))
    emit(14, not _cp_bad, "；".join(_cp_bad[:4]))

    # -3：六个编号板块的标题与正文文字内容（中英文之间及“/”两侧均不插入空格，不增删字符）
    # 严格对齐细则，仅踩细则明确写出的点，不额外约束：
    #   ① 6 个板块标题逐字一致；② 每块 4 条正文逐字一致（空格算字符）。
    # 与办公软件对齐：把“ATF4核内调控解析”改成“ATF4 核内调控解析”、把
    # “ATF4沉默/过表达”改成“ATF4沉默 / 过表达”，都是指令没要求的改字，
    # 用户要逐条比对着删空格才能还原。
    _txt_bad = []
    for _num in _BLOCK_NUMS:
        _blk = blocks[_num]
        _want_title, _want_items, _ = _BLOCK_SPECS[_num]
        _got_title = _title_text_raw(_blk, _num)
        if _got_title != _want_title:
            _txt_bad.append("编号%s标题为“%s”（应“%s”）" % (_num, _got_title, _want_title))
        _body = _blk["body"]
        _got_items = [t.strip() for t in (_body["paras_text"] if _body else []) if t.strip()]
        for _k, _want in enumerate(_want_items):
            _got = _got_items[_k] if _k < len(_got_items) else ""
            if _got != _want:
                _txt_bad.append("编号%s第%d条正文为“%s”（应“%s”）" % (_num, _k + 1, _got, _want))
    emit(15, not _txt_bad, "；".join(_txt_bad[:4]))

    # -3：六个编号板块的编号与标题文字颜色（01=0F5B5B 02=253E83 03=55328D
    #     04=8A4C10 05=2D6C34 06=8B274B）
    # 严格对齐细则，仅踩细则明确写出的点，不额外约束：
    #   ① 编号文字颜色为该块指定色值；② 板块标题文字颜色为同一色值。
    # 与办公软件对齐：这六个色值是每块与底层流程框配色对应的品牌色，改成
    # 邻近色（如 0F5B5B→006156）后六块与底图不再同色系，用户得逐块取色改回。
    _clr_bad = []
    for _num in _BLOCK_NUMS:
        _want = _BLOCK_SPECS[_num][2]
        for _key, _label in (("badge", "编号"), ("title", "标题")):
            _it = blocks[_num][_key]
            if _it is None:
                continue
            for _r in _it["runs"]:
                if not _r.text.strip():
                    continue
                _got = _run_color(_r)
                if _got != _want:
                    _clr_bad.append("编号%s的%s颜色为%s（应%s）" % (_num, _label, _got, _want))
                    break
    emit(16, not _clr_bad, "；".join(_clr_bad[:4]))

    # -3：六个编号板块的编号、标题与正文文字须同时指定 a:latin 与 a:ea 为“微软雅黑”
    # 严格对齐细则，仅踩细则明确写出的点，不额外约束：
    #   ① 每个 run 的拉丁文字体为微软雅黑；② 每个 run 的东亚文字字体也为微软雅黑。
    # 与办公软件对齐：只写拉丁字体名而丢掉东亚文字字体，中文会回落到主题的
    # 默认中文字体（本主题的次要东亚字体为空，实际就是系统兜底字体），
    # 屏幕上表现为中文字形、字宽、笔画粗细整体变样，且无法靠“字体”框一键改回。
    _fnt_bad = []
    for _num in _BLOCK_NUMS:
        for _key, _label in (("badge", "编号"), ("title", "标题"), ("body", "正文")):
            _it = blocks[_num][_key]
            if _it is None:
                continue
            for _r in _it["runs"]:
                if not _r.text.strip():
                    continue
                _lat, _ea = _run_typeface(_r, "a:latin"), _run_typeface(_r, "a:ea")
                if not (_is_yahei(_lat) and _is_yahei(_ea)):
                    _fnt_bad.append("编号%s的%s拉丁字体=%s、东亚字体=%s"
                                    % (_num, _label, _lat, _ea))
                    break
    emit(17, not _fnt_bad, "；".join(_fnt_bad[:4]))

    # -1：六个编号板块的编号框与标题框文本垂直对齐须为垂直居中（anchor="ctr"）
    # 严格对齐细则，仅踩细则明确写出的点，不额外约束：
    #   ① 承载编号的文本框 anchor="ctr"；② 承载标题的文本框 anchor="ctr"。
    # 与办公软件对齐：即“设置形状格式 → 文本选项 → 垂直对齐方式：中部对齐”。
    # 改成顶端对齐后编号与标题在框内整体上移，与底层流程框的圆角标签错位。
    _anc_bad = []
    for _num in _BLOCK_NUMS:
        for _key, _label in (("badge", "编号框"), ("title", "标题框")):
            _it = blocks[_num][_key]
            if _it is None:
                continue
            if _it["layout"]["anchor"] != "ctr":
                _anc_bad.append("编号%s的%s垂直对齐为%s" % (_num, _label, _it["layout"]["anchor"]))
    emit(18, not _anc_bad, "；".join(_anc_bad[:4]))

    # -1：页面底部说明“输出目标：…”须为微软雅黑 9.3磅 非加粗 465C67、段落居中、
    #     文本垂直居中，文本框位于左1.83cm/上25.45cm、尺寸16.38×0.91cm
    # 严格对齐细则，仅踩细则明确写出的点，不额外约束：
    #   ① 字体族；② 9.3磅；③ 非加粗；④ 颜色465C67；⑤ 段落居中；
    #   ⑥ 文本垂直居中；⑦ 左/上位置；⑧ 宽/高尺寸。
    # 与办公软件对齐：这句是全页收尾的说明条，字号被放大到 10.5 磅、垂直对齐
    # 改成顶端后，它与底部注释之间的呼吸感被破坏，整页收口显得头重脚轻。
    _nm_bad = _check_text_spec(text_items, _NOTE_MAIN)
    emit(19, not _nm_bad, "；".join(_nm_bad[:4]))

    # -1：页面底部注释“所有流程框为背景图片；…”须为微软雅黑 6.8磅 83909A、
    #     段落居中、文本垂直居中，文本框位于左1.83cm/上26.49cm、尺寸16.38×0.66cm
    # 严格对齐细则，仅踩细则明确写出的点，不额外约束：同上 8 项。
    # 与办公软件对齐：这是最小号的脚注灰字，再缩小到 6.3 磅基本不可读。
    _ns_bad = _check_text_spec(text_items, _NOTE_SUB)
    emit(20, not _ns_bad, "；".join(_ns_bad[:4]))

    # -3：作为参照样式的“临床问题”“科学假说”两块须保持原有规格
    # 严格对齐细则，仅踩细则明确写出的点，不额外约束：
    #   ① 两块底图的位置与尺寸；② 编号“A”“B”的字号/加粗/颜色/位置；
    #   ③ 标题“临床问题”“科学假说”的字号/加粗/颜色/位置；
    #   ④ “临床问题”3条要点的字号/颜色/位置尺寸且保留“• ”前缀；
    #   ⑤ “科学假说”正文段的字号/颜色/位置尺寸。
    # 与办公软件对齐：这两块是六个编号板块要对齐的“样板间”，它一被改动，
    # 下面六块即使照它排也对不上，整页的版式基准就没了。
    _ref_bad = []
    _pics = [sh for sh in prs.slides[0].shapes if sh.shape_type == MSO_SHAPE_TYPE.PICTURE]
    for _l, _t, _w, _h in _REF_PIC_RECTS:
        if not any(abs(_cm(p.left) - _l) <= POS_TOL_CM and abs(_cm(p.top) - _t) <= POS_TOL_CM
                   and abs(_cm(p.width) - _w) <= POS_TOL_CM
                   and abs(_cm(p.height) - _h) <= POS_TOL_CM for p in _pics):
            _ref_bad.append("参照块底图（左%.2f/上%.2f、%.2f×%.2fcm）被移动、缩放或删除"
                            % (_l, _t, _w, _h))
    for _spec in _REF_TEXT_SPECS:
        _ref_bad.extend(_check_text_spec(text_items, _spec))
    _ref_bad.extend(_check_para_box(text_items, "临床问题要点框", _REF_BULLETS, _REF_BULLET_BOX))
    _ref_bad.extend(_check_para_box(text_items, "科学假说正文段", [_REF_HYPO_TEXT], _REF_HYPO_BOX))
    emit(21, not _ref_bad, "；".join(_ref_bad[:4]))

    # -3：六个编号板块与“临床问题”“科学假说”之外的固定文字须保持原有规格
    # 严格对齐细则，仅踩细则明确写出的点，不额外约束：
    #   主标题、副标题、右上导语、顶部5个节点主标签与5个副标签、小节标题、
    #   中部横条文字、下部横条文字，各自的字体族/字号/加粗/颜色/段落对齐，
    #   以及细则中给出坐标的那几处的位置与尺寸。
    # 与办公软件对齐：这些是指令没让动的页面骨架文字，动了就要用户逐处对着
    # 原样重设字号和颜色。
    _fix_bad = []
    for _spec in _FIXED_TEXT_SPECS:
        _fix_bad.extend(_check_text_spec(text_items, _spec))
    emit(22, not _fix_bad, "；".join(_fix_bad[:4]))

    # -3：页面内23张图片须全部保留且位置与尺寸不发生移动、缩放、替换或删除
    # 严格对齐细则，仅踩细则明确写出的点，不额外约束：
    #   ① 图片总数为 23；② 23 个既定位置尺寸各有一张图片一一对上。
    # 与办公软件对齐：所有流程框、箭头、横条都是背景图片，文字只是叠在上面；
    # 任一张被拖动或缩放，文字与底框的对位关系立刻全错。
    _pic_bad = []
    if len(_pics) != len(_PIC_RECTS):
        _pic_bad.append("图片共%d张（应%d张）" % (len(_pics), len(_PIC_RECTS)))
    _pool = list(_pics)
    for _l, _t, _w, _h in _PIC_RECTS:
        _hit = None
        for _p in _pool:
            if (abs(_cm(_p.left) - _l) <= POS_TOL_CM and abs(_cm(_p.top) - _t) <= POS_TOL_CM
                    and abs(_cm(_p.width) - _w) <= POS_TOL_CM
                    and abs(_cm(_p.height) - _h) <= POS_TOL_CM):
                _hit = _p
                break
        if _hit is None:
            _pic_bad.append("缺少位于左%.2f/上%.2f、尺寸%.2f×%.2fcm的图片" % (_l, _t, _w, _h))
        else:
            _pool.remove(_hit)
    emit(23, not _pic_bad, "；".join(_pic_bad[:4]))

    # ==================== 汇总：反解 results 得到 dim2_items ====================
    if len(results) != len(_DELTAS_ORDER):
        raise RuntimeError("dim2 项数不一致：results=%d vs order=%d"
                           % (len(results), len(_DELTAS_ORDER)))
    dim2_items = []
    total_score = 0
    for raw, max_delta in zip(results, _DELTAS_ORDER):
        head, _, rule = raw.partition("：")
        try:
            delta = int(head)
        except ValueError:
            delta = 0
        if max_delta > 0:
            # 加分项：命中记满额，未命中记 0
            hit = delta > 0
            got = max_delta if hit else 0
            shown_max = max_delta
        else:
            # 减分项：合规记 0，触发则扣该额度；max_delta 保留负额度，
            # 以满足「delta 只能是 0 或 max_delta」的统一结果契约（扣分项不计入 max_score）
            hit = delta == 0
            got = 0 if hit else max_delta
            shown_max = max_delta
        dim2_items.append({"rule": rule, "max_delta": shown_max, "delta": got, "hit": hit})
        total_score += got
    result["dim2_items"] = dim2_items
    result["total_score"] = max(total_score, 0)
    result["max_score"] = MAX_SCORE_TOTAL
    return result


if __name__ == "__main__":
    if len(sys.argv) >= 2:
        _dir = sys.argv[1]
    else:
        _dir = os.path.dirname(os.path.abspath(__file__))
    _out = evaluate(_dir)
    print(json.dumps(_out, ensure_ascii=False, indent=2))
