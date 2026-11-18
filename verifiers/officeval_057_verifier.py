#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""自动评估 editable_chemical_scheme.pptx（057：按图片做一页可编辑化学反应式PPT）。

对外统一接口：
    evaluate(dir_path: str = None) -> dict
        评估目标锁定为本脚本所在目录：脚本只在该目录内定位 .pptx 文档并评估，
        dir_path 参数仅为保持接口统一而保留，传入任何值都会被忽略。
        返回结构：id / file_name / status / error / dim1_pass / dim1_reason /
                  dim2_items(rule / max_delta / delta / hit) / total_score / max_score。

指令：给我按照图片做一页PPT，每个部分都要可编辑，大小和图片一样，画圈那要改成直线

评分逻辑：
  维度1（文件可正常打开）为一票否决：文件打不开 -> dim1_pass=False，维度2不评估。
  维度2（完成度）仅加分项，满分 60。
本脚本在计数之外，进一步按“办公软件打开后的实际观感”评估线条位置与相连情况：
  - 键线角度是否规整落在 30° 栅格（骨架式画法）
  - 键长是否均匀（变异系数）
  - 五个分子块的整体质心是否落在原图对应的相对位置（元素位置评测，仅对已满足结构/角度要求的合格部件判定，各3分）
  - 中部三唑环芳香键是否以双线表示、有无误画的叁键
这些项能把“线条错位/断裂/悬空/整体排布错乱/键型错误”的输出显著拉开分差。
"""
import json
import math
import os
import re
import statistics

from pptx import Presentation

SCRIPT_ID = "057"
DEFAULT_PPTX_NAME = "editable_chemical_scheme.pptx"

# 评估目标锁定为本脚本所在目录：评测时把本脚本与被评交付物放在同一个目录下即可。
# 不接受外部传入的目录，也不检索子目录与上级目录。
_SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

IMG_RATIO = 1.9403  # 原图 image.png 877x452 的宽高比
A = '{http://schemas.openxmlformats.org/drawingml/2006/main}'


def _emu(el, name):
    """读取 OOXML 的 EMU 整数属性。

    部分生成工具会写出带小数的值（如 cx="2066350.784"），按四舍五入取整，
    避免整脚本因单个非整数属性中断。
    """
    if el is None:
        return 0
    raw = el.get(name)
    if raw is None:
        return 0
    try:
        return int(round(float(raw)))
    except (TypeError, ValueError):
        return 0


def line_endpoints(sp):
    """从 <a:off>/<a:ext>(+flipH/flipV) 计算线段两端点。返回 ((x0,y0),(x1,y1)) 或 None。"""
    xfrm = sp._element.find('.//' + A + 'xfrm')
    if xfrm is None:
        return None
    off = xfrm.find(A + 'off'); ext = xfrm.find(A + 'ext')
    if off is None or ext is None:
        return None
    x = _emu(off, 'x'); y = _emu(off, 'y')
    cx = _emu(ext, 'cx'); cy = _emu(ext, 'cy')
    fH = xfrm.get('flipH') == '1'; fV = xfrm.get('flipV') == '1'
    x0, y0, x1, y1 = x, y, x + cx, y + cy
    if fH:
        x0, x1 = x + cx, x
    if fV:
        y0, y1 = y + cy, y
    return (x0, y0), (x1, y1)


def gather(slide):
    """收集幻灯片上的线条(段端点)、文本框、图片。"""
    lines, texts, pics = [], [], 0
    segs = []  # 线段端点列表

    def walk(shapes):
        nonlocal pics
        for sp in shapes:
            st = str(sp.shape_type)
            try:
                l, t, w, h = sp.left, sp.top, sp.width, sp.height
            except Exception:
                l = t = w = h = None
            xml = sp._element.xml
            if 'PICTURE' in st or '<p:pic>' in xml:
                pics += 1
            if l is not None:
                cx = l + (w or 0) / 2
                cy = t + (h or 0) / 2
                if sp.has_text_frame and sp.text_frame.text.strip():
                    texts.append((cx, cy, sp.text_frame.text.strip()))
            is_line = ('prst="line"' in xml) or (sp.shape_type is not None and 'CONNECTOR' in st)
            if is_line:
                ep = line_endpoints(sp)
                if ep:
                    segs.append(ep)
                    lines.append(((ep[0][0] + ep[1][0]) / 2, (ep[0][1] + ep[1][1]) / 2))
            if hasattr(sp, "shapes"):
                walk(sp.shapes)

    walk(slide.shapes)
    return lines, texts, pics, segs


def region_counts(items_lines, items_texts, W, H, box):
    x0, x1, y0, y1 = box
    lc = sum(1 for cx, cy in items_lines if x0 * W <= cx <= x1 * W and y0 * H <= cy <= y1 * H)
    tc = [t for cx, cy, t in items_texts if x0 * W <= cx <= x1 * W and y0 * H <= cy <= y1 * H]
    return lc, tc


def segs_in_region(segs, W, H, box):
    """按线段中心点筛选落在某分区内的键线。"""
    x0, x1, y0, y1 = box
    out = []
    for (a, b) in segs:
        cx = (a[0] + b[0]) / 2; cy = (a[1] + b[1]) / 2
        if x0 * W <= cx <= x1 * W and y0 * H <= cy <= y1 * H:
            out.append((a, b))
    return out


def region_centroid(items_lines, W, H, box):
    """返回落在分区内所有键线中心点的“相对质心”(relx, rely) 及数量。无键线时返回 None。"""
    x0, x1, y0, y1 = box
    pts = [(cx, cy) for cx, cy in items_lines
           if x0 * W <= cx <= x1 * W and y0 * H <= cy <= y1 * H]
    if not pts:
        return None
    rx = sum(p[0] for p in pts) / len(pts) / W
    ry = sum(p[1] for p in pts) / len(pts) / H
    return (rx, ry, len(pts))


def geometry_quality(segs):
    """返回(角度规整比例, 键长变异系数CV)。用于判断线条位置是否正确、结构是否相连。"""
    angs_ok = 0
    lens = []
    n = 0
    for (a, b) in segs:
        dx = b[0] - a[0]; dy = b[1] - a[1]
        d = math.hypot(dx, dy)
        if d <= 0:
            continue
        lens.append(d)
        n += 1
        ang = math.degrees(math.atan2(dy, dx)) % 180.0
        dev = min(abs(ang - m) for m in (0, 30, 60, 90, 120, 150, 180))
        if dev <= 7.0:
            angs_ok += 1
    ang_ratio = angs_ok / n if n else 0.0
    cv = (statistics.pstdev(lens) / statistics.mean(lens)) if len(lens) >= 2 else 1.0
    return ang_ratio, cv


def count_multibonds(segs, W, H):
    """在一组键线里，统计“互相近平行且靠得很近、长度相仿”的短线簇。
    返回 (双键簇数, 疑似叁键簇数)。叁键=3条及以上平行短线并排——本题红圈处
    以及仍未改直线的部件会出现这种簇。"""
    n = len(segs)
    used = [False] * n
    diag = math.hypot(W, H)
    triples = doubles = 0

    def mid(s):
        return ((s[0][0] + s[1][0]) / 2, (s[0][1] + s[1][1]) / 2)

    def ang(s):
        return math.degrees(math.atan2(s[1][1] - s[0][1], s[1][0] - s[0][0])) % 180.0

    def ln(s):
        return math.hypot(s[1][0] - s[0][0], s[1][1] - s[0][1])

    for i in range(n):
        if used[i]:
            continue
        grp = [i]; used[i] = True
        ai = ang(segs[i]); li = ln(segs[i]); mi = mid(segs[i])
        for j in range(n):
            if used[j]:
                continue
            aj = ang(segs[j]); lj = ln(segs[j]); mj = mid(segs[j])
            da = abs(ai - aj); da = min(da, 180 - da)
            if da > 10:                                    # 需近平行
                continue
            if max(li, lj) / max(1e-9, min(li, lj)) > 1.8:  # 长度相仿
                continue
            gap = math.hypot(mi[0] - mj[0], mi[1] - mj[1]) / diag
            if gap <= 0.03:                                # 中点靠得很近
                grp.append(j); used[j] = True
        if len(grp) >= 3:
            triples += 1
        elif len(grp) == 2:
            doubles += 1
    return doubles, triples


# ── 五个分子组件所在的相对区域（relx0, relx1, rely0, rely1，页面归一化）────────────
REGIONS = {
    "左上反应物": (0.0, 0.30, 0.0, 0.36),
    "右上产物":   (0.62, 1.0, 0.0, 0.36),
    "左下反应物": (0.0, 0.30, 0.55, 0.92),
    "中部三唑酯": (0.28, 0.53, 0.40, 0.92),
    "右下终产物": (0.63, 1.0, 0.38, 0.92),
}

# 各分子块在原图中的目标相对质心(relx, rely)，坐标从页面左上角算起。
POS_TARGET = {
    "左上反应物": (0.13, 0.18),
    "右上产物":   (0.78, 0.18),
    "左下反应物": (0.13, 0.70),
    "中部三唑酯": (0.42, 0.61),
    "右下终产物": (0.84, 0.64),
}
POS_TOL = 0.12  # 质心与目标位置的最大相对距离（约页宽/页高的12%）

# ── 维度二评分项定义（rule + max_delta），保持稳定顺序 ──────────────────────────
DIM2_RULES = [
    {"key": "single_page", "rule": "+1：整个演示文稿只含1张幻灯片", "max_delta": 1},
    {"key": "aspect_ratio", "rule": "+1：幻灯片宽高比为1.94:1（宽:高与原图偏差不超过3%），页面尺寸与原图一致", "max_delta": 1},
    {"key": "vector_editable", "rule": "+3：整页由可单独选中编辑的矢量线段与文本框构成（键线≥20根、文本框≥10个、图片0张），非整张位图", "max_delta": 3},
    {"key": "circle_to_line", "rule": "+5：红圈处叁键改为直线（单键），全图键线总数≤90根", "max_delta": 5},
    {"key": "top_plus", "rule": "+1：上排反应物之间出现1个“+”号文本", "max_delta": 1},
    {"key": "top_arrow", "rule": "+1：上排出现1条水平指向右方的反应箭头“→”", "max_delta": 1},
    {"key": "reagent_text", "rule": "+1：上排出现二甲胺试剂文本(CH3)2NH（含CH与NH字样）", "max_delta": 1},
    {"key": "bottom_plus", "rule": "+1：下排反应物之间出现1个“+”号文本", "max_delta": 1},
    {"key": "bottom_arrow", "rule": "+1：下排出现1条水平指向右方的反应箭头“→”", "max_delta": 1},
    {"key": "comp1_struct", "rule": "+1：左上反应物5-氯-2-(氯甲基)吡嗪——氯代吡嗪环+CH2Cl侧链，结构完整（该区键线≥8根）", "max_delta": 1},
    {"key": "comp1_geom", "rule": "+3：左上反应物——键线方向规整（标准方向±7°占比≥90%）、键长均匀（CV≤0.30），六边形端正相接无错位", "max_delta": 3},
    {"key": "comp2_struct", "rule": "+1：右上产物5-氯-2-(二甲氨基甲基)吡嗪——吡嗪环+二甲氨基侧链，结构完整（键线≥8根）", "max_delta": 1},
    {"key": "comp2_geom", "rule": "+3：右上产物——键线方向规整（±7°占比≥90%）、键长均匀（CV≤0.30），六边形端正相接无错位", "max_delta": 3},
    {"key": "comp3_struct", "rule": "+1：左下反应物（与右上产物同一分子）——吡嗪环+二甲氨基侧链，结构完整（键线≥8根）", "max_delta": 1},
    {"key": "comp3_geom", "rule": "+3：左下反应物——键线方向规整（±7°占比≥90%）、键长均匀（CV≤0.33），六边形端正相接无错位", "max_delta": 3},
    {"key": "comp4_struct", "rule": "+3：中部三唑-乙酸酯——五元三唑环(3N)+乙酸酯(酯基O)，结构完整（区域内N≥2、O≥1，键线≥8根）", "max_delta": 3},
    {"key": "comp4_bond", "rule": "+3：中部三唑-乙酸酯——芳香键以双线表示，且不出现3条平行短线误画的叁键(≡)", "max_delta": 3},
    {"key": "comp5_struct", "rule": "+3：右下终产物吡嗪-CH2-N(CH3)-C(=O)-三唑——吡嗪环+酰胺(N+羰基O)+三唑，结构完整（键线≥15、N≥4、O≥1）", "max_delta": 3},
    {"key": "comp5_geom", "rule": "+5：右下终产物——键线方向规整（±7°占比≥75%）、键长均匀（CV≤0.30），两环加主链首尾相接不错位", "max_delta": 5},
    {"key": "aromatic_double", "rule": "+3：吡嗪环、三唑环内芳香双键以“双线”（两条相邻平行短线）表示（全图键线≥40根）", "max_delta": 3},
    {"key": "subscript", "rule": "+1：下标书写正确，CH3中的“3”、(CH3)2中的“2”为下标小字", "max_delta": 1},
    {"key": "pos_comp1", "rule": "+3：元素位置——左上反应物整体落在原图对应区域（页面左上、上排左侧）", "max_delta": 3},
    {"key": "pos_comp2", "rule": "+3：元素位置——右上产物整体落在原图对应区域（页面右上、上排右侧）", "max_delta": 3},
    {"key": "pos_comp3", "rule": "+3：元素位置——左下反应物整体落在原图对应区域（页面左下、下排左侧）", "max_delta": 3},
    {"key": "pos_comp4", "rule": "+3：元素位置——中部三唑-乙酸酯整体落在原图对应区域（页面中部、下排中间偏上）", "max_delta": 3},
    {"key": "pos_comp5", "rule": "+3：元素位置——右下终产物整体落在原图对应区域（页面右下、下排右侧）", "max_delta": 3},
]

# 各分子块的位置评测项 -> 对应组件名（位置项以该组件“结构/角度”合格为前提）
POS_KEY_TO_COMP = {
    "pos_comp1": "左上反应物",
    "pos_comp2": "右上产物",
    "pos_comp3": "左下反应物",
    "pos_comp4": "中部三唑酯",
    "pos_comp5": "右下终产物",
}


def _new_dim2_items():
    """按 DIM2_RULES 生成默认未命中的 dim2_items 列表。"""
    items = []
    for r in DIM2_RULES:
        items.append({
            "rule": r["rule"],
            "max_delta": r["max_delta"],
            "delta": 0,
            "hit": False,
        })
    return items


def _dim2_item(items, key):
    """按 key 从 DIM2_RULES 找到 items 中对应条目。"""
    idx = next(i for i, r in enumerate(DIM2_RULES) if r["key"] == key)
    return items[idx]


def check_dim1(prs, reasons):
    """维度1：文件可正常打开。

    文件已能被 python-pptx 打开（否则 evaluate 走错误分支）；此处仅再确认
    文档中至少存在1张幻灯片，可进入后续评估。
    """
    if len(prs.slides) < 1:
        reasons.append("文档内没有任何幻灯片")
        return False
    return True


def check_dim2(prs):
    """维度二：返回 dim2_items 列表（命中与未命中都在其中）。"""
    items = _new_dim2_items()

    n_slides = len(prs.slides)
    W, H = prs.slide_width, prs.slide_height
    slide = prs.slides[0]
    lines, texts, pics, segs = gather(slide)
    alltext = ' '.join(t for _, _, t in texts)
    slide_xml = slide._element.xml

    def hit(key, cond, delta, _detail=""):
        # _detail 仅为保留各判定处的可读注释而存在，不再写入返回结果。
        it = _dim2_item(items, key)
        it["hit"] = bool(cond)
        it["delta"] = delta if cond else 0
        return bool(cond)

    # ========== 整体：单页与尺寸 ==========
    hit("single_page", n_slides == 1, 1, f"slides={n_slides}")
    ratio = W / H
    hit("aspect_ratio", abs(ratio - IMG_RATIO) / IMG_RATIO <= 0.03, 1,
        f"ratio={ratio:.4f} vs {IMG_RATIO:.4f}")
    editable = (len(lines) >= 20 and len(texts) >= 10 and pics == 0)
    hit("vector_editable", editable, 3, f"lines={len(lines)}, texts={len(texts)}, pics={pics}")

    # ========== 关键任务：画圈改直线 ==========
    hit("circle_to_line", len(lines) <= 90, 5,
        f"总线条={len(lines)}（>90 疑似仍保留叁键/多重平行短线）")

    # ========== 反应式通用元素 ==========
    plus_cnt = alltext.count('+')
    arrow_cnt = slide_xml.count('tailEnd type="triangle"') + slide_xml.count('<a:custGeom>')
    hit("top_plus", plus_cnt >= 2, 1, f"文档内+号数={plus_cnt}")
    hit("top_arrow", arrow_cnt >= 1, 1, f"箭头数≈{arrow_cnt}")
    has_reagent = bool(re.search(r'\(?CH.{0,3}\)?.{0,3}2.{0,3}NH', alltext.replace(' ', '')) or
                       ('NH' in alltext and 'CH' in alltext))
    hit("reagent_text", has_reagent, 1, f"text片段含CH/NH={has_reagent}")
    hit("bottom_plus", plus_cnt >= 2, 1, f"文档内+号数={plus_cnt}")
    hit("bottom_arrow", arrow_cnt >= 2, 1, f"箭头数≈{arrow_cnt}")

    # ========== 五个分子组件 ==========
    rc = {name: region_counts(lines, texts, W, H, box) for name, box in REGIONS.items()}
    rg = {name: geometry_quality(segs_in_region(segs, W, H, box)) for name, box in REGIONS.items()}
    cen = {name: region_centroid(lines, W, H, box) for name, box in REGIONS.items()}
    comp_ok = {}  # 组件是否满足“结构完整/角度/键型”等要求

    def geom_detail(name):
        r, c = rg[name]
        return f"角度规整={r*100:.0f}%, CV={c:.2f}"

    # 组件1：左上反应物
    lc, tc = rc["左上反应物"]; ar, cv = rg["左上反应物"]
    s1 = lc >= 8 and any('N' in x for x in tc) and any('C' in x.upper() for x in tc)
    g1 = ar >= 0.90 and cv <= 0.30
    hit("comp1_struct", s1, 1, f"lines={lc}")
    hit("comp1_geom", g1, 3, geom_detail("左上反应物"))
    comp_ok["左上反应物"] = s1 and g1

    # 组件2：右上产物
    lc, tc = rc["右上产物"]; ar, cv = rg["右上产物"]
    s2 = lc >= 8 and any('N' in x for x in tc)
    g2 = ar >= 0.90 and cv <= 0.30
    hit("comp2_struct", s2, 1, f"lines={lc}")
    hit("comp2_geom", g2, 3, geom_detail("右上产物"))
    comp_ok["右上产物"] = s2 and g2

    # 组件3：左下反应物
    lc, tc = rc["左下反应物"]; ar, cv = rg["左下反应物"]
    s3 = lc >= 8 and any('N' in x for x in tc)
    g3 = ar >= 0.90 and cv <= 0.33
    hit("comp3_struct", s3, 1, f"lines={lc}")
    hit("comp3_geom", g3, 3, geom_detail("左下反应物"))
    comp_ok["左下反应物"] = s3 and g3

    # 组件4：中部三唑-乙酸酯
    lc, tc = rc["中部三唑酯"]
    n_mid = sum(x.count('N') for x in tc)
    o_mid = sum(x.count('O') for x in tc)
    s4 = lc >= 8 and n_mid >= 2 and o_mid >= 1
    hit("comp4_struct", s4, 3, f"lines={lc}, N计={n_mid}, O计={o_mid}")
    dbl_mid, tri_mid = count_multibonds(segs_in_region(segs, W, H, REGIONS["中部三唑酯"]), W, H)
    b4 = (tri_mid == 0 and dbl_mid >= 1)
    hit("comp4_bond", b4, 3, f"双键簇={dbl_mid}, 疑似叁键簇={tri_mid}")
    comp_ok["中部三唑酯"] = s4 and b4

    # 组件5：右下终产物
    lc, tc = rc["右下终产物"]; ar, cv = rg["右下终产物"]
    n_end = sum(x.count('N') for x in tc)
    o_end = sum(x.count('O') for x in tc)
    s5 = lc >= 15 and n_end >= 4 and o_end >= 1
    g5 = ar >= 0.75 and cv <= 0.30
    hit("comp5_struct", s5, 3, f"lines={lc}, N计={n_end}, O计={o_end}")
    hit("comp5_geom", g5, 5, geom_detail("右下终产物"))
    comp_ok["右下终产物"] = s5 and g5

    # ========== 结构细节 ==========
    hit("aromatic_double", len(lines) >= 40, 3, f"总线条={len(lines)}")
    has_sub = ('₃' in alltext) or ('₂' in alltext) or ('baseline="-' in slide_xml)
    hit("subscript", has_sub, 1, "含下标字符或baseline标记")

    # ========== 元素位置评测（前提：该组件已满足结构/角度/键型要求）==========
    for key, name in POS_KEY_TO_COMP.items():
        c = cen[name]
        if not comp_ok.get(name):
            hit(key, False, 3, "组件未满足结构/角度要求，位置不评")
            continue
        if c is None:
            hit(key, False, 3, "该区无键线，无法定位")
            continue
        rx, ry, npt = c
        tx, ty = POS_TARGET[name]
        d = math.hypot(rx - tx, ry - ty)
        hit(key, d <= POS_TOL, 3,
            f"质心=({rx:.2f},{ry:.2f}) 目标=({tx:.2f},{ty:.2f}) 距离={d:.2f} n={npt}")

    return items


def _find_target_file(dir_path):
    """仅在给定目录本身定位待评估的 .pptx 文档，不进入任何子目录。

    优先取默认文件名 editable_chemical_scheme.pptx；否则回退到目录内首个 .pptx。
    找不到时抛 FileNotFoundError，由 evaluate 统一转成“未找到文件”的返回格式。
    """
    if not dir_path or not os.path.isdir(dir_path):
        raise FileNotFoundError(f"目录不存在或不是目录：{dir_path}")
    preferred = os.path.join(dir_path, DEFAULT_PPTX_NAME)
    if os.path.isfile(preferred):
        return preferred
    candidates = sorted(
        os.path.join(dir_path, name)
        for name in os.listdir(dir_path)
        if name.lower().endswith(".pptx")
        and not name.startswith("~$")
        and os.path.isfile(os.path.join(dir_path, name))
    )
    if not candidates:
        raise FileNotFoundError(f"目录中未找到 .pptx 文件：{dir_path}")
    return candidates[0]


def evaluate(dir_path: str = None) -> dict:
    """统一入口：只评估本脚本所在目录内的 .pptx 文档。

    dir_path 参数仅为保持各脚本接口一致而保留，传入的值不会被使用。

    两类异常分开回传：
      · 目录不存在 / 目录内没有 .pptx——维度一根本没进入判定，error 只带纯文本原因
        （不带异常类名），file_name、dim1_pass、dim1_reason 一律留空（dim1_pass 为空
        字符串而非 False），dim2 不出项，总分与满分一并归零；
      · 其他情况（文件打不开、运行期异常）——dim1_pass 记 False 并在 dim1_reason 中
        给出错误摘要，max_score 仍保留真实满分。
    """
    dir_path = dir_path if dir_path and os.path.isdir(dir_path) else _SCRIPT_DIR
    max_score = sum(r["max_delta"] for r in DIM2_RULES if r["max_delta"] > 0)
    result = {
        "id": SCRIPT_ID,
        "file_name": "",
        "status": "ok",
        "error": None,
        "dim1_pass": False,
        "dim1_reason": "",
        "dim2_items": _new_dim2_items(),
        "total_score": 0,
        "max_score": max_score,
    }

    try:
        target = _find_target_file(dir_path)
        result["file_name"] = os.path.basename(target)

        try:
            prs = Presentation(target)
        except Exception as e:
            result["status"] = "error"
            result["error"] = f"文件无法打开: {e}"
            result["dim1_pass"] = False
            result["dim1_reason"] = (f"维度一未通过：文件无法打开或解析"
                                     f"（{type(e).__name__}: {e}），完成度总分记 0")
            result["total_score"] = 0
            return result

        dim1_reasons = []
        dim1_ok = check_dim1(prs, dim1_reasons)
        result["dim1_pass"] = dim1_ok
        if not dim1_ok:
            # 维度一未通过：dim2_items 保留结构但不评估，总分为0
            result["dim1_reason"] = ("维度一未通过：%s，完成度总分记 0"
                                     % ("；".join(dim1_reasons) or "文件无法正常打开"))
            result["total_score"] = 0
            return result

        items = check_dim2(prs)
        result["dim2_items"] = items
        result["total_score"] = sum(i["delta"] for i in items)
        return result
    except FileNotFoundError as e:
        # 找不到被评估文件：只回传纯文本原因，维度一未进入判定，各字段留空、分数归零。
        result["status"] = "error"
        result["error"] = str(e)
        result["file_name"] = ""
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
        result["dim1_reason"] = (f"评测过程异常终止：{type(e).__name__}: {e}")
        result["total_score"] = 0
        return result


if __name__ == "__main__":
    # 本地调试入口：固定评估本脚本所在目录，不接受命令行参数。
    print(json.dumps(evaluate(), ensure_ascii=False, indent=2))


