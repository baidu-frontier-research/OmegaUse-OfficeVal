#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# pyright: reportUnknownMemberType=false, reportUnknownVariableType=false, reportUnknownArgumentType=false, reportUnknownParameterType=false, reportMissingParameterType=false, reportMissingImports=false, reportOptionalSubscript=false, reportOptionalMemberAccess=false
"""officeval 074：「农业县低空经济发展方案.pptx」中译英交付稿完成度评测。

指令要的活是「把 PPT 里所有中文翻译成英语」，禁止动的是「排版与图片背景」。
因此加分项只判各类文字元素的英译覆盖面与译文质量，减分项只判指令未要求、
却被改动而破坏版面的属性（字号、字色、加粗、对齐、填充、遮罩、图像数据、
页数页序、编号文字等）。同一属性只按一条细则计扣，不跨条重复计分。

按接口统一约定，本模块不修改全局 sys.stdout，也不 print 主结果。主结果统一由
evaluate(dir_path) -> dict 返回；入参为「脚本所在目录的路径」，由脚本自己在该
目录内定位并打开被评估的 .pptx。本文件仅在 __main__ 下作为本地调试入口，才把
dict 序列化为 JSON 打印。

原稿基线（46 页每页的形状/图片数量、每个文本形状的磅值·加粗·色值·对齐、
色块与遮罩填充签名、配图几何与 srcRect、媒体 md5、编号文字等）已内嵌在
_BASELINE_JSON 中，评测时无需另行提供原稿文件。
"""

import os
import re
import sys
import json
import zipfile
import hashlib
from xml.etree import ElementTree as ET

SCRIPT_ID = "074"

# 满分只统计加分项：+5 全篇覆盖 + 11×(+3) 各类元素 + 2×(+1) 专项 = 40
MAX_SCORE_TOTAL = 40

# dim2 各条的满分增量，必须与 _evaluate_impl 内 results 的追加顺序严格一致。
# 正值为加分项（计入满分），负值为减分项（命中即扣，不计入满分）。
_MAX_DELTAS_ORDER = [
    5,    # 01 全篇 46 页无中文残留
    3,    # 02 第1页封面主标题·副标题·右上角徽章英译
    3,    # 03 第2页目录 7 条 PART 条目标题及说明文字英译
    3,    # 04 章节分隔页章节标题与一句话导语英译
    3,    # 05 各内容页顶部主标题英译
    3,    # 06 主标题下方过渡副标题英译
    3,    # 07 卡片小标题与卡片正文段落英译
    3,    # 08 指标磁贴文字英译
    3,    # 09 流程步骤条文字英译
    3,    # 10 第16与第36页对比行行标题·说明·结论句英译
    3,    # 11 全部 46 条底部页脚导航英译
    3,    # 12 译文完整通顺（无中英混排、无漏译、无乱码）
    1,    # 13 第46页"感谢观看"与署名行英译
    1,    # 14 第18页横幅句·第5页强调句·第24页放大路径英译
    -3,   # 15 页数页序与内容对应关系被改动
    -3,   # 16 画布不再是 13.333 × 7.5 英寸
    -3,   # 17 任一页面形状对象数量减少
    -1,   # 18 封面主标题字号/字色被改动
    -1,   # 19 封面副标题字号/字色被改动
    -1,   # 20 满版底图上方深色遮罩被改动
    -1,   # 21 目录 PART 条目标题与说明文字字号/字色被改动
    -3,   # 22 章节分隔页两位数序号字号/字色被改动
    -3,   # 23 章节分隔页标题与导语字号/字色被改动
    -5,   # 24 内容页顶部主标题字号/字色被改动
    -5,   # 25 卡片正文段落字号/字色被改动
    -3,   # 26 过渡副标题字号/字色被改动
    -3,   # 27 卡片小标题字号/字色被改动
    -3,   # 28 加粗状态被改动
    -3,   # 29 对齐方式被改动
    -3,   # 30 指标磁贴字号/字色被改动
    -3,   # 31 第16页对比行与结论句字号/字色被改动
    -3,   # 32 第36页对比行与结论句字号被改动
    -3,   # 33 流程步骤序号与步骤名称字号被改动
    -1,   # 34 第18页横幅句与横幅底色被改动
    -1,   # 35 第5页强调句被改动
    -1,   # 36 第24页"放大路径"标题与说明被改动
    -3,   # 37 底部页脚导航字号/字色被改动
    -1,   # 38 页面底色·顶部色条·卡片底板填充或圆角被改动
    -5,   # 39 配图图像数据未逐字节保持一致
    -3,   # 40 配图裁剪范围·摆放·叠放层次被改动
    -3,   # 41 说明性文字被删除或改写
    -1,   # 42 编号类文字被改写或丢失
    -1,   # 43 署名·机构名·地名等实体名称取值被改动
    -5,   # 44 英文使行距被压缩（写入 lnSpcReduction）
]
assert sum(d for d in _MAX_DELTAS_ORDER if d > 0) == MAX_SCORE_TOTAL

_NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main"
_NS_P = "http://schemas.openxmlformats.org/presentationml/2006/main"
_NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

_EMU_PER_INCH = 914400.0

# 细则允许「位置偏差不超过 0.01 英寸」的另存等价改写，判定时取 0.011 作容差
_POS_TOL = 0.011

# 原稿主题的配色映射：细则允许「仅重写颜色写法而渲染色值不变」（如白色由 FFFFFF
# 改写为主题背景色），故主题色一律先归一化为具体色值再比较
_SCHEME_COLOR = {"bg1": "FFFFFF", "lt1": "FFFFFF", "tx1": "000000",
                 "dk1": "000000", "bg2": "FFFFFF", "tx2": "000000"}

_CJK_RE = re.compile(r"[㐀-䶿一-鿿豈-﫿]")
# 中文标点（含全角标点）；不含 ‘’“” 等英文也常用的弯引号，避免误判正常英文排印
_CJK_PUNCT_RE = re.compile(r"[　-〿！-～]")
# 判定一个 run 是否承载可感知字形：含字母、数字或汉字才算，纯空白/纯标点不算
_GLYPH_RE = re.compile(r"[0-9A-Za-z㐀-䶿一-鿿豈-﫿]")

# 章节分隔页、指标磁贴页、流程步骤条页（按原稿页号）
_SEC_PAGES = frozenset({3, 6, 14, 17, 19, 23, 42})
_METRIC_PAGES = frozenset({15, 24})
_STEP_PAGES = frozenset({8, 9, 22, 33, 34, 37, 44})

# 实体名一致性判定时忽略的通用词（保留"青岚"音译一类专有名词参与比较）
_ENTITY_STOP = frozenset({
    "the", "and", "for", "with", "from", "into", "county", "countys", "county's",
    "agricultural", "agriculture", "low", "altitude", "lowaltitude", "aerial",
    "application", "applications", "implementation", "plan", "project", "program",
    "research", "group", "team", "work", "task", "force", "study", "office",
})

# 细则中「为容纳英文而调整文本框的位置、尺寸、内边距或垂直锚点」一条已在评测口径
# 中删除（英文本身比中文长，允许为容纳英文调整文本框）。经确认，【封面主标题】
# 【内容页主标题】【底部页脚】等条目内含的文本框几何子句（如 L0.65/T0.52/W9.30/
# H0.50）随该条一并豁免，一律不参与判定；本开关仅为留档，正常评测不应改动
_JUDGE_TEXTBOX_GEOMETRY = False

# 细则规定的文本框几何（英寸：左/上/宽/高），已随文本框调整条豁免，仅留档备查
_BOX_COVER_TITLE = (0.66, 1.70, 5.10, 1.30)
_BOX_MAIN_TITLE = (0.65, 0.52, 9.30, 0.50)
_BOX_FOOTER = (0.55, 7.14, 12.15, 0.22)

# 第33页配图应有的几何；封面与结束页配图应满版
_PIC_P33 = (0.70, 1.45, 5.60, 4.90)
_PIC_FULL = (0.00, 0.00, 13.333, 7.50)

def _a(tag):
    return "{%s}%s" % (_NS_A, tag)


def _p(tag):
    return "{%s}%s" % (_NS_P, tag)


def _inch(value):
    """EMU 转英寸，保留 4 位小数。"""
    try:
        return round(int(value) / _EMU_PER_INCH, 4)
    except (TypeError, ValueError):
        return None


def _solid_fill(node):
    """把 <a:solidFill> 解析为 'RRGGBB' 或 'RRGGBB/alpha'；主题色归一化为具体色值。"""
    if node is None:
        return None
    srgb = node.find(_a("srgbClr"))
    if srgb is not None:
        alpha = srgb.find(_a("alpha"))
        val = (srgb.get("val") or "").upper()
        return val + ("/" + alpha.get("val") if alpha is not None else "")
    scheme = node.find(_a("schemeClr"))
    if scheme is not None:
        raw = scheme.get("val") or ""
        val = _SCHEME_COLOR.get(raw, raw)
        alpha = scheme.find(_a("alpha"))
        return val + ("/" + alpha.get("val") if alpha is not None else "")
    return None


def _shape_box(sp_pr):
    """读取形状的 a:xfrm 偏移与尺寸，返回 [左, 上, 宽, 高]（英寸）。"""
    if sp_pr is None:
        return None
    xfrm = sp_pr.find(_a("xfrm"))
    if xfrm is None:
        return None
    off, ext = xfrm.find(_a("off")), xfrm.find(_a("ext"))
    if off is None or ext is None:
        return None
    return [_inch(off.get("x")), _inch(off.get("y")),
            _inch(ext.get("cx")), _inch(ext.get("cy"))]


def _parse_slide(xml_bytes):
    """解析单页幻灯片，抽取与基线同构的属性集合。

    形状按 cNvPr/@name 归档（同名形状按文档顺序排队），另存一份文档顺序的扁平表
    seq 作为配对兜底；不含文字的形状只贡献填充签名与计数。
    """
    root = ET.fromstring(xml_bytes)
    info = {"sp": 0, "pic": 0, "t": {}, "seq": [], "f": [], "p": [],
            "lnsp": 0, "fsc": 0}
    for sp in root.iter(_p("sp")):
        info["sp"] += 1
        cnv = sp.find(".//" + _p("cNvPr"))
        name = (cnv.get("name") if cnv is not None else None) or "?"
        sp_pr = sp.find(_p("spPr"))
        box = _shape_box(sp_pr)
        if sp_pr is not None:
            color = _solid_fill(sp_pr.find(_a("solidFill")))
            if color:
                geom = sp_pr.find(_a("prstGeom"))
                info["f"].append("%s:%s:%s" % (
                    geom.get("prst") if geom is not None else "custom", color,
                    "%.2fx%.2f" % (box[2], box[3]) if box else "-"))
        runs, npara, text = [], 0, ""
        for para in sp.iter(_a("p")):
            p_pr = para.find(_a("pPr"))
            algn = (p_pr.get("algn") if p_pr is not None else None) or "l"
            ptext = "".join(t.text or "" for t in para.iter(_a("t")))
            if ptext.strip():
                npara += 1
                text += ptext
            for run in para.iter(_a("r")):
                r_pr = run.find(_a("rPr"))
                if r_pr is None or not r_pr.get("sz"):
                    continue
                rtext = "".join(t.text or "" for t in run.iter(_a("t")))
                runs.append([int(r_pr.get("sz")), 1 if r_pr.get("b") == "1" else 0,
                             _solid_fill(r_pr.find(_a("solidFill"))), algn, rtext])
        for autofit in sp.iter(_a("normAutofit")):
            if autofit.get("lnSpcReduction"):
                info["lnsp"] += 1
            if autofit.get("fontScale"):
                info["fsc"] += 1
        if runs:
            item = {"r": runs, "np": npara, "b": box, "txt": text,
                    "name": name, "idx": len(info["seq"])}
            info["seq"].append(item)
            info["t"].setdefault(name, []).append(item)
    for pic in root.iter(_p("pic")):
        info["pic"] += 1
        box = _shape_box(pic.find(_p("spPr")))
        src = pic.find(".//" + _a("srcRect"))
        crop = {k: int(v) for k, v in (src.attrib.items() if src is not None else [])}
        info["p"].append(((box or [None] * 4) + [crop])[:5])
    info["f"].sort()
    tree = root.find("%s/%s" % (_p("cSld"), _p("spTree")))
    zmap = {_p("sp"): "s", _p("pic"): "p", _p("graphicFrame"): "g",
            _p("grpSp"): "G", _p("cxnSp"): "c"}
    info["z"] = "".join(zmap[c.tag] for c in (list(tree) if tree is not None else [])
                        if c.tag in zmap)
    return info


def _slide_targets(pres_xml, rels_xml):
    """按 sldIdLst 的真实放映顺序返回幻灯片部件路径（不依赖 slideN.xml 的编号顺序）。"""
    rel_map = {}
    for rel in ET.fromstring(rels_xml):
        rel_map[rel.get("Id")] = rel.get("Target") or ""
    targets = []
    root = ET.fromstring(pres_xml)
    lst = root.find(_p("sldIdLst"))
    for sld in (list(lst) if lst is not None else []):
        target = rel_map.get(sld.get("{%s}id" % _NS_R), "")
        target = target.replace("../", "").lstrip("/")
        if not target.startswith("ppt/"):
            target = "ppt/" + target
        targets.append(target)
    return targets


def _load_deck(file_path):
    """打开 .pptx（zip 直读，不落地解压），返回画布尺寸、逐页解析结果与媒体 md5。"""
    with zipfile.ZipFile(file_path) as zf:
        names = zf.namelist()
        pres_xml = zf.read("ppt/presentation.xml")
        rels_xml = zf.read("ppt/_rels/presentation.xml.rels")
        targets = [t for t in _slide_targets(pres_xml, rels_xml) if t in names]
        if not targets:
            targets = sorted((n for n in names if re.match(r"ppt/slides/slide\d+\.xml$", n)),
                             key=lambda n: int(re.search(r"(\d+)", n.split("/")[-1]).group(1)))
        slides = []
        for pos, target in enumerate(targets, 1):
            info = _parse_slide(zf.read(target))
            info["pos"] = pos
            slides.append(info)
        md5 = sorted(hashlib.md5(zf.read(n)).hexdigest()
                     for n in names if n.startswith("ppt/media/") and not n.endswith("/"))
    sld_sz = ET.fromstring(pres_xml).find(_p("sldSz"))
    return {"cx": _inch(sld_sz.get("cx")) if sld_sz is not None else None,
            "cy": _inch(sld_sz.get("cy")) if sld_sz is not None else None,
            "slides": slides, "md5": md5}


def _page_order(deck, total):
    """由页脚导航行开头的页码反推放映顺序对应的原稿页号。

    页脚是全篇唯一自带页号的元素（形如「10 涉农场景 · 上页 … → 下页 …」），
    比 slideN.xml 编号可靠；若页码不构成 1..total 的排列，则退化为按位置顺序对应。
    """
    order = []
    for slide in deck["slides"]:
        cand = None
        for item in slide["seq"]:
            top = item["b"][1] if item["b"] else None
            is_footer = any(r[0] == 780 for r in item["r"]) or (top is not None and top >= 6.8)
            if not is_footer:
                continue
            if cand is None or (top is not None and cand["b"] and top > cand["b"][1]):
                cand = item
        num = None
        if cand is not None:
            hit = re.search(r"\d{1,2}", cand["txt"][:28])
            if hit:
                num = int(hit.group(0))
        order.append(num)
    if len(order) == total and sorted(n for n in order if n) == list(range(1, total + 1)):
        return order, True
    return [i + 1 for i in range(len(deck["slides"]))], False


def _pair_shapes(exp_slide, got_slide):
    """页内按 cNvPr/@name 配对（同名按文档顺序消费），名称对不上的再按文档顺序兜底。

    不按 a:off 几何配对：每页都有多个位于 (0,0) 的形状，且交付稿常有位移，
    以几何为键会静默错配。
    """
    pool = {k: list(v) for k, v in got_slide["t"].items()}
    pairs, pending, used = [], [], set()
    for name, exp in exp_slide["t"].items():
        bucket = pool.get(name)
        if bucket:
            got = bucket.pop(0)
            used.add(got["idx"])
            pairs.append((name, exp, got))
        else:
            pending.append((name, exp))
    if pending:
        rest = [it for it in got_slide["seq"] if it["idx"] not in used]
        for (name, exp), got in zip(pending, rest):
            pairs.append((name, exp, got))
        pending = pending[len(rest):]
    return pairs, [name for name, _ in pending]


def _run_pairs(page, pairs):
    """逐 run 配对；交付稿把一段拆成多个 run 时，多出的 run 与基线末个 run 比较。

    只承载空白或标点的 run 不参与字号/字色/加粗比对：译文常把中文句子重排成
    「正文 run + 句末标点 run」，这类碎片 run 没有可感知的字形，属另存产生的等价
    改写，不作为字号被缩小或字色被改动的证据。
    """
    out = []
    for name, exp, got in pairs:
        exp_runs = exp["r"]
        if not exp_runs:
            continue
        for idx, got_run in enumerate(got["r"]):
            if len(got_run) > 4 and not _GLYPH_RE.search(got_run[4]):
                continue
            exp_run = exp_runs[idx] if idx < len(exp_runs) else exp_runs[-1]
            out.append((page, name, idx, exp_run, got_run))
    return out


def _fmt(items, limit=5):
    """把命中明细压缩成一行；超出 limit 只列前几处并给出总数。"""
    if not items:
        return ""
    text = "；".join(items[:limit])
    return text + ("…（共 %d 处）" % len(items) if len(items) > limit else "")


def _is_translated(text):
    """英译判定：非空、不含中日韩汉字、且至少含两个拉丁字母。"""
    body = (text or "").strip()
    if not body or _CJK_RE.search(body):
        return False
    return len(re.findall(r"[A-Za-z]", body)) >= 2


def _text_ok(exp, got):
    """按基线原文判断该形状的文字是否已妥善处理。

    原文含中文的，必须变成不含中文且有实义字母的英文；原文本就是数字或英文的
    （章节两位数序号、流程步骤序号、指标数字、已有的英文行）只要求未被删空、
    未混入中文，其取值另由编号类细则核查。
    """
    if got is None:
        return False
    text = (got["txt"] or "").strip()
    if not text or _CJK_RE.search(text):
        return False
    if exp.get("c"):
        return len(re.findall(r"[A-Za-z]", text)) >= 2
    return True


def _diff_runs(run_pairs, pred, sz=False, col=False, bold=False, algn=False):
    """在满足 pred(页号, 基线run) 的 run 上比较指定属性，返回可读的差异明细。"""
    bad = []
    for page, name, _idx, exp, got in run_pairs:
        if not pred(page, exp):
            continue
        tag = "p%d/%s" % (page, name)
        if sz and got[0] != exp[0]:
            bad.append("%s 字号 %.1f→%.1f" % (tag, exp[0] / 100.0, got[0] / 100.0))
        if col and (got[2] or "") != (exp[2] or ""):
            bad.append("%s 字色 %s→%s" % (tag, exp[2] or "继承", got[2] or "继承"))
        if bold and got[1] != exp[1]:
            bad.append("%s 加粗 %s→%s" % (tag, "是" if exp[1] else "否",
                                          "是" if got[1] else "否"))
        if algn and got[3] != exp[3]:
            bad.append("%s 对齐 %s→%s" % (tag, exp[3], got[3]))
    return bad


def _geo_bad(box, want):
    """几何比较：任一维超过 0.01 英寸容差即算改动。"""
    if box is None or any(v is None for v in box):
        return True
    return any(abs(box[i] - want[i]) > _POS_TOL for i in range(4))


def _baseline():
    """载入内嵌的原稿基线常量。"""
    return json.loads(_BASELINE_JSON)


# 交付物只接受 .pptx：细则要求交付「1 个可正常打开的 PowerPoint 演示文稿（.pptx）」，
# 其余演示文稿/导出格式（.ppt / .pptm / .odp / .pdf 等）一律在维度一判不通过。
_ACCEPTED_EXT = ".pptx"
# 目录里没有 .pptx 时，这些扩展名仍收作被评对象，交给维度一判不通过（总分记 0），
# 而不是当成「没有交付物」直接跳过评测。
_WRONG_FMT_EXTS = (".ppt", ".pptm", ".ppsx", ".ppsm", ".pps", ".potx", ".pot",
                   ".odp", ".key", ".pdf")


def _fmt_reject(file_path):
    """维度1的格式门禁：交付物必须是可正常打开的 .pptx（OOXML 演示文稿包）。

    不通过返回原因文本，通过返回 None。判定与 PowerPoint 的「能不能双击打开」对齐：
    扩展名须为 .pptx，文件须是 zip 容器，且包内含 ppt/presentation.xml。
    """
    ext = os.path.splitext(file_path)[1].lower()
    if ext != _ACCEPTED_EXT:
        return "交付物为 %s 格式，只接受 .pptx 演示文稿" % (ext or "无扩展名")
    if not zipfile.is_zipfile(file_path):
        return "交付物不是有效的 .pptx 包（无法作为 zip 容器打开）"
    try:
        with zipfile.ZipFile(file_path) as zf:
            names = set(zf.namelist())
    except Exception as exc:
        return "交付物无法作为 .pptx 打开：%s: %s" % (type(exc).__name__, exc)
    if "ppt/presentation.xml" not in names:
        return "交付物包内缺少 ppt/presentation.xml，不是 PowerPoint 演示文稿"
    return None


def _check_dim1(baseline, deck):
    """维度1：可用与可修改性（格式门禁另见 _fmt_reject）。

    只拦「无法使用」的情形：页数残缺、整页被导出为图片（文本对象大面积消失）、
    文字整体跑出画布。细则口径中不把局部包围盒轻微交叠算作版面严重错乱。
    """
    fails = []
    total = baseline["n"]
    if len(deck["slides"]) != total:
        fails.append("幻灯片为 %d 页，应为 %d 页" % (len(deck["slides"]), total))
    blank = [s["pos"] for s in deck["slides"] if not s["seq"]]
    if blank:
        fails.append("第 %s 页无可单独点选编辑的文本对象" % ",".join(str(x) for x in blank[:6]))
    exp_shapes = sum(len(s["t"]) for s in baseline["s"].values())
    got_shapes = sum(len(s["seq"]) for s in deck["slides"])
    if exp_shapes and got_shapes < exp_shapes * 0.6:
        fails.append("可编辑文本对象仅 %d 个（原稿 %d 个），疑似整页转为位图"
                     % (got_shapes, exp_shapes))
    cx = deck["cx"] or baseline["cx"]
    cy = deck["cy"] or baseline["cy"]
    off_canvas = []
    for slide in deck["slides"]:
        for item in slide["seq"]:
            box = item["b"]
            if not box or any(v is None for v in box):
                continue
            if box[0] >= cx or box[1] >= cy or box[0] + box[2] <= 0 or box[1] + box[3] <= 0:
                off_canvas.append("p%d/%s" % (slide["pos"], item["name"]))
    if off_canvas:
        fails.append("文本框完全超出画布：%s" % _fmt(off_canvas, 4))
    return fails


def _evaluate_impl(file_path, result):
    # 维度一的第一道门禁：交付物格式只接受 .pptx，不合格即判 0，不再解析内容
    reject = _fmt_reject(file_path)
    if reject:
        result["dim1_pass"] = False
        result["dim1_reason"] = "维度一未通过：%s，完成度总分记 0" % reject
        result["dim2_items"] = []
        result["total_score"] = 0
        result["max_score"] = 0
        return result

    baseline = _baseline()
    deck = _load_deck(file_path)

    dim1 = _check_dim1(baseline, deck)
    if dim1:
        result["dim1_pass"] = False
        result["dim1_reason"] = "维度一未通过：%s，完成度总分记 0" % "；".join(dim1)
        result["dim2_items"] = []
        result["total_score"] = 0
        result["max_score"] = 0
        return result
    result["dim1_pass"] = True

    # 先由页脚页码把交付稿的放映顺序还原成原稿页号，再逐页配对形状
    order, order_ok = _page_order(deck, baseline["n"])
    pmap = {}
    for pos, page in enumerate(order):
        if page is not None:
            pmap.setdefault(page, pos)

    page_pairs, shape_pairs, run_pairs, missing = [], [], [], []
    for key in sorted(baseline["s"], key=int):
        page, exp_slide = int(key), baseline["s"][key]
        pos = pmap.get(page)
        got_slide = deck["slides"][pos] if pos is not None and pos < len(deck["slides"]) else None
        if got_slide is None:
            missing.extend((page, nm) for nm in exp_slide["t"])
            continue
        page_pairs.append((page, exp_slide, got_slide))
        pairs, lost = _pair_shapes(exp_slide, got_slide)
        missing.extend((page, nm) for nm in lost)
        shape_pairs.extend((page, nm, exp, got) for nm, exp, got in pairs)
        run_pairs.extend(_run_pairs(page, pairs))

    got_of = {(page, name): got for page, name, _exp, got in shape_pairs}
    slide_of = {page: got_slide for page, _exp, got_slide in page_pairs}
    mask_pages = {1} | set(_SEC_PAGES)

    def cls(sizes, pages=None, color=None):
        """按基线属性挑出某类文字元素：[(页号, 形状名, 基线项, 交付项或 None)]。"""
        picked = []
        for key in sorted(baseline["s"], key=int):
            page = int(key)
            if pages is not None and page not in pages:
                continue
            for name, exp in baseline["s"][key]["t"].items():
                if any(r[0] in sizes and (color is None or r[2] == color) for r in exp["r"]):
                    picked.append((page, name, exp, got_of.get((page, name))))
        return picked

    def untranslated(group):
        return ["p%d/%s" % (page, name) for page, name, exp, got in group
                if not _text_ok(exp, got)]

    def geo_bad_list(group, want):
        """文本框几何核查；口径中该条已删，默认不参与判定。"""
        if not _JUDGE_TEXTBOX_GEOMETRY:
            return []
        return ["p%d/%s 文本框几何偏离 %s" % (page, name, want)
                for page, name, _exp, got in group
                if got is None or _geo_bad(got["b"], want)]

    results = []

    def bonus(score, rule, ok, detail=""):
        results.append(("+%d：%s" % (score if ok else 0, rule), "" if ok else detail))

    def penalty(score, rule, hit, detail=""):
        results.append(("-%d：%s" % (score if hit else 0, rule), detail if hit else ""))

    # 01 +5：全部 46 页幻灯片内不残留任何中文字符（含标题、副标题、卡片、指标、
    # 流程标签、结论句、页脚导航）。逐形状扫描交付稿全部文字，不依赖配对结果。
    cjk = ["p%d/%s" % (s["pos"], it["name"]) for s in deck["slides"] for it in s["seq"]
           if _CJK_RE.search(it["txt"])]
    bonus(5, "全部 46 页幻灯片内不残留任何中文字符", not cjk, _fmt(cjk))

    # 02 +3：第1页封面主标题（32磅）、副标题（13磅）与右上角徽章（8.2磅）均译为英文
    group = cls({3200, 1300, 820}, pages={1})
    bad = untranslated(group)
    bonus(3, "第1页封面主标题、副标题与右上角徽章均译为英文", not bad, _fmt(bad))

    # 03 +3：第2页目录 7 条 PART 条目标题（11.5磅）及其下方说明文字（8.4磅）均译为英文
    group = cls({1150, 840}, pages={2})
    bad = untranslated(group)
    bonus(3, "第2页目录 7 条 PART 条目标题及其说明文字均译为英文", not bad, _fmt(bad))

    # 04 +3：第3、6、14、17、19、23、42 页章节标题（34磅）与一句话导语（14磅）均译为英文
    group = cls({3400}, pages=_SEC_PAGES) + cls({1400}, pages=_SEC_PAGES)
    bad = untranslated(group)
    bonus(3, "7 个章节分隔页的章节标题与一句话导语均译为英文", not bad, _fmt(bad))

    # 05 +3：各内容页顶部主标题（26磅）全部译为英文
    group = cls({2600})
    bad = untranslated(group)
    bonus(3, "各内容页顶部主标题全部译为英文", not bad, _fmt(bad))

    # 06 +3：主标题下方过渡副标题（11磅、6A7F7D）全部译为英文
    group = cls({1100}, color="6A7F7D")
    bad = untranslated(group)
    bonus(3, "各内容页主标题下方的过渡副标题全部译为英文", not bad, _fmt(bad))

    # 07 +3：卡片小标题（12.5磅、063D3B）与卡片正文段落（9.2磅）全部译为英文
    group = cls({1250}, color="063D3B") + cls({920})
    bad = untranslated(group)
    bonus(3, "各卡片的小标题与卡片正文段落全部译为英文", not bad, _fmt(bad))

    # 08 +3：第15、24 页指标磁贴的程度词（25磅）、指标名（11磅）与注解（8.2磅）全部译为英文
    group = cls({2500, 1100, 820}, pages=_METRIC_PAGES)
    bad = untranslated(group)
    bonus(3, "指标磁贴文字全部译为英文", not bad, _fmt(bad))

    # 09 +3：第8、9、22、33、34、37、44 页流程步骤条的序号（10磅）与步骤名（8.4磅）译为英文
    group = cls({1000}) + cls({840}, pages=_STEP_PAGES)
    bad = untranslated(group)
    bonus(3, "流程步骤条文字全部译为英文", not bad, _fmt(bad))

    # 10 +3：第16页「植保环节/传统方式—低空方式」与第36页「人工踏勘/特点—适用建议」
    # 的行标题、行内说明及底部结论句全部译为英文
    group = cls({1300, 1050, 870}, pages={16}) + cls({1250, 1050, 880}, pages={36})
    bad = untranslated(group)
    bonus(3, "第16与第36页对比行的行标题、行内说明及底部结论句全部译为英文",
          not bad, _fmt(bad))

    # 11 +3：全部 46 条底部页脚导航文字（7.8磅）译为英文，且 46 条一条不缺
    group = cls({780})
    bad = untranslated(group)
    if len([1 for _p2, _n, _e, got in group if got is not None]) != baseline["n"]:
        bad.append("页脚条数不足 %d 条" % baseline["n"])
    bonus(3, "全部 46 条底部页脚导航文字译为英文", not bad, _fmt(bad))

    # 12 +3：译文完整通顺——无中英混排、无漏译半句、无空文本、无中文标点残留、无乱码
    quality = ["p%d/%s 文本框缺失" % t for t in missing]
    for page, name, exp, got in shape_pairs:
        tag = "p%d/%s" % (page, name)
        if not _text_ok(exp, got):
            quality.append("%s 为空或未译" % tag)
            continue
        if _CJK_PUNCT_RE.search(got["txt"]):
            quality.append("%s 残留中文标点" % tag)
        if "�" in got["txt"] or "???" in got["txt"]:
            quality.append("%s 含乱码或占位符" % tag)
    bonus(3, "译文完整通顺，无中英混排、无漏译半句、无断句错乱", not quality, _fmt(quality))

    # 13 +1：第46页「感谢观看」（26磅）与署名行（12磅）译为英文；
    # 「THANKS FOR YOUR TIME」本已为英文，保持不变即可
    group = cls({2600, 1200}, pages={46})
    bad = untranslated(group)
    for _pg, name, exp, got in cls({1700}, pages={46}):
        text = re.sub(r"\s+", " ", (got["txt"] if got else "")).strip().upper()
        if text != "THANKS FOR YOUR TIME":
            bad.append("p46/%s 原有英文行被改写：%s" % (name, text or "空"))
    bonus(1, "第46页「感谢观看」与署名行译为英文，原有英文行保持不变", not bad, _fmt(bad))

    # 14 +1：第18页横幅句（16磅）、第5页强调句（20磅）、第24页「放大路径」标题（15磅）
    # 及其说明（12磅）均译为英文
    group = (cls({1600}, pages={18}) + cls({2000}, pages={5})
             + cls({1500, 1200}, pages={24}))
    bad = untranslated(group)
    bonus(1, "第18页横幅句、第5页强调句与第24页「放大路径」标题及说明均译为英文",
          not bad, _fmt(bad))

    # 15 -3：交付稿不满足 46 页、页序与内容对应关系不变（缺页、并页、加页、调序均计）。
    # 页序以页脚自带页码为准；页码若整体无法解析则只判页数，不臆断页序。
    hit = len(deck["slides"]) != baseline["n"]
    detail = "幻灯片 %d 页" % len(deck["slides"]) if hit else ""
    if order_ok and order != list(range(1, baseline["n"] + 1)):
        moved = ["第%d位=原第%d页" % (i + 1, n) for i, n in enumerate(order) if n != i + 1]
        hit, detail = True, "页序错位：%s" % _fmt(moved, 6)
    penalty(3, "交付稿不满足 46 页、页序与内容对应关系不变", hit, detail)

    # 16 -3：幻灯片画布不满足 13.333 英寸 × 7.5 英寸（16:9）
    hit = (deck["cx"] is None or deck["cy"] is None
           or abs(deck["cx"] - baseline["cx"]) > _POS_TOL
           or abs(deck["cy"] - baseline["cy"]) > _POS_TOL)
    penalty(3, "幻灯片画布不满足 13.333 × 7.5 英寸（16:9）", hit,
            "画布 %s × %s 英寸" % (deck["cx"], deck["cy"]))

    # 17 -3：任一页面中的形状对象数量减少（色块、卡片、分隔线、遮罩、图片、文本框被删或合并）
    shrink = []
    for page, exp_slide, got_slide in page_pairs:
        if got_slide["sp"] < exp_slide["sp"]:
            shrink.append("p%d 形状 %d→%d" % (page, exp_slide["sp"], got_slide["sp"]))
        if got_slide["pic"] < exp_slide["pic"]:
            shrink.append("p%d 图片 %d→%d" % (page, exp_slide["pic"], got_slide["pic"]))
    shrink += ["p%d/%s 文本框缺失" % t for t in missing]
    penalty(3, "任一页面中的形状对象数量减少", bool(shrink), _fmt(shrink))

    # 18 -1：封面主标题不满足 32 磅、颜色 FFFFFF（加粗另按【加粗状态】条计；
    # 文本框几何按口径不再判定）
    bad = _diff_runs(run_pairs, lambda p, r: p == 1 and r[0] == 3200, sz=True, col=True)
    bad += geo_bad_list(cls({3200}, pages={1}), _BOX_COVER_TITLE)
    penalty(1, "封面主标题不满足 32 磅、颜色 FFFFFF", bool(bad), _fmt(bad))

    # 19 -1：封面副标题不满足 13 磅、颜色 EAF7F4
    bad = _diff_runs(run_pairs, lambda p, r: p == 1 and r[0] == 1300, sz=True, col=True)
    penalty(1, "封面副标题不满足 13 磅、颜色 EAF7F4", bool(bad), _fmt(bad))

    # 20 -1：满版底图上方的深色遮罩透明度被改动（封面 000000/82000 + 063D3B/92000；
    # 章节分隔页 000000/62000 + 063D3B/78000。遮罩被删除、加深、变浅或换色均计；
    # 只比较色值与透明度，不比较遮罩尺寸）
    def alpha_fills(sig_list):
        out = []
        for sig in sig_list:
            parts = sig.split(":")
            if len(parts) >= 2 and "/" in parts[1]:
                out.append(parts[1])
        return sorted(out)

    bad = []
    for page, exp_slide, got_slide in page_pairs:
        if page not in mask_pages:
            continue
        want, have = alpha_fills(exp_slide["f"]), alpha_fills(got_slide["f"])
        if want != have:
            bad.append("p%d 遮罩 %s→%s" % (page, want, have))
    penalty(1, "满版底图上方的深色遮罩透明度被改动", bool(bad), _fmt(bad))

    # 21 -1：7 条 PART 条目标题不满足 11.5 磅、颜色 0E7C7B 与 1F6E86 逐条交替，
    # 其下方说明文字不满足 8.4 磅
    bad = _diff_runs(run_pairs, lambda p, r: p == 2 and r[0] == 1150, sz=True, col=True)
    bad += _diff_runs(run_pairs, lambda p, r: p == 2 and r[0] == 840, sz=True)
    penalty(1, "目录 PART 条目标题不满足 11.5 磅与 0E7C7B/1F6E86 交替、说明不满足 8.4 磅",
            bool(bad), _fmt(bad))

    # 22 -3：章节分隔页的两位数序号不满足 48 磅、颜色 E5A63B
    bad = _diff_runs(run_pairs, lambda p, r: r[0] == 4800, sz=True, col=True)
    penalty(3, "章节分隔页两位数序号不满足 48 磅、颜色 E5A63B", bool(bad), _fmt(bad))

    # 23 -3：章节分隔页标题不满足 34 磅、颜色 FFFFFF；其下导语不满足 14 磅、颜色 EAF7F4
    bad = _diff_runs(run_pairs, lambda p, r: r[0] == 3400, sz=True, col=True)
    bad += _diff_runs(run_pairs, lambda p, r: p in _SEC_PAGES and r[0] == 1400,
                      sz=True, col=True)
    penalty(3, "章节分隔页标题不满足 34 磅 FFFFFF、导语不满足 14 磅 EAF7F4",
            bool(bad), _fmt(bad))

    # 24 -5：内容页顶部主标题不满足 26 磅、颜色 112B2D（被整体缩放字号计）
    bad = _diff_runs(run_pairs, lambda p, r: r[0] == 2600, sz=True, col=True)
    bad += geo_bad_list(cls({2600}), _BOX_MAIN_TITLE)
    penalty(5, "内容页顶部主标题不满足 26 磅、颜色 112B2D", bool(bad), _fmt(bad))

    # 25 -5：卡片正文段落不满足 9.2 磅、颜色 112B2D（被整体缩放字号计）
    bad = _diff_runs(run_pairs, lambda p, r: r[0] == 920, sz=True, col=True)
    penalty(5, "卡片正文段落不满足 9.2 磅、颜色 112B2D", bool(bad), _fmt(bad))

    # 26 -3：主标题下方过渡副标题不满足 11 磅、颜色 6A7F7D
    bad = _diff_runs(run_pairs, lambda p, r: r[0] == 1100 and r[2] == "6A7F7D",
                     sz=True, col=True)
    penalty(3, "主标题下方过渡副标题不满足 11 磅、颜色 6A7F7D", bool(bad), _fmt(bad))

    # 27 -3：卡片小标题不满足 12.5 磅、颜色 063D3B
    bad = _diff_runs(run_pairs, lambda p, r: r[0] == 1250 and r[2] == "063D3B",
                     sz=True, col=True)
    penalty(3, "卡片小标题不满足 12.5 磅、颜色 063D3B", bool(bad), _fmt(bad))

    # 28 -3：各页标题与正文的加粗状态被改动。按计分约定，全篇加粗差异只在本条计扣，
    # 其余各条不再重复判定加粗
    bad = _diff_runs(run_pairs, lambda p, r: True, bold=True)
    penalty(3, "各页标题与正文的加粗状态被改动", bool(bad), _fmt(bad))

    # 29 -3：各页文字的对齐方式被改动。按计分约定只在本条计扣；补写与继承值一致的
    # 对齐属性视为等价改写（缺省 algn 一律按继承值 l 处理），不计
    bad = _diff_runs(run_pairs, lambda p, r: True, algn=True)
    penalty(3, "各页文字的对齐方式被改动", bool(bad), _fmt(bad))

    # 30 -3：指标磁贴的大号数字/程度词不满足 25 磅及 0E7C7B/1F6E86/E5A63B/C96B3C 的
    # 分列取值，其下指标名不满足 11 磅、注解不满足 8.2 磅（三行居中另按【对齐方式】条计）
    bad = _diff_runs(run_pairs, lambda p, r: p in _METRIC_PAGES and r[0] == 2500,
                     sz=True, col=True)
    bad += _diff_runs(run_pairs, lambda p, r: p in _METRIC_PAGES and r[0] in (1100, 820),
                      sz=True)
    penalty(3, "指标磁贴数字/指标名/注解不满足 25 磅分列取色与 11、8.2 磅", bool(bad), _fmt(bad))

    # 31 -3：第16页对比行不满足 13 磅（行标题）/ 10.5 磅（方式名）/ 8.7 磅（说明），
    # 底部结论句不满足 13 磅、颜色 0E7C7B
    bad = _diff_runs(run_pairs, lambda p, r: p == 16 and r[0] in (1300, 1050, 870), sz=True)
    bad += _diff_runs(run_pairs, lambda p, r: p == 16 and r[2] == "0E7C7B", col=True)
    penalty(3, "第16页对比行不满足 13/10.5/8.7 磅或结论句不满足 13 磅 0E7C7B",
            bool(bad), _fmt(bad))

    # 32 -3：第36页对比行不满足 12.5 磅 / 10.5 磅 / 8.8 磅，底部结论句不满足 12.5 磅
    bad = _diff_runs(run_pairs, lambda p, r: p == 36 and r[0] in (1250, 1050, 880), sz=True)
    penalty(3, "第36页对比行不满足 12.5/10.5/8.8 磅", bool(bad), _fmt(bad))

    # 33 -3：流程步骤序号不满足 10 磅、步骤名称不满足 8.4 磅。按口径「框被拉宽」一项
    # 已随文本框调整条删除，本条只判字号
    bad = _diff_runs(run_pairs, lambda p, r: r[0] == 1000, sz=True)
    bad += _diff_runs(run_pairs, lambda p, r: p in _STEP_PAGES and r[0] == 840, sz=True)
    penalty(3, "流程步骤序号不满足 10 磅、步骤名称不满足 8.4 磅", bool(bad), _fmt(bad))

    # 34 -1：第18页横幅句不满足 16 磅、颜色 FFFFFF，其下横幅底色不满足 000000 透明度 65000
    bad = _diff_runs(run_pairs, lambda p, r: p == 18 and r[0] == 1600, sz=True, col=True)
    banner = slide_of.get(18)
    if banner is not None and not any(sig.split(":")[1] == "000000/65000"
                                      for sig in banner["f"] if len(sig.split(":")) >= 2):
        bad.append("p18 横幅底色 000000/65000 缺失或被改写")
    penalty(1, "第18页横幅句不满足 16 磅 FFFFFF 或横幅底色不满足 000000/65000",
            bool(bad), _fmt(bad))

    # 35 -1：第5页「政策可行 → 场景可用 → 运营可持续」一行不满足 20 磅、颜色 063D3B
    bad = _diff_runs(run_pairs, lambda p, r: p == 5 and r[0] == 2000, sz=True, col=True)
    penalty(1, "第5页强调句不满足 20 磅、颜色 063D3B", bool(bad), _fmt(bad))

    # 36 -1：第24页「放大路径」标题不满足 15 磅、其说明不满足 12 磅
    bad = _diff_runs(run_pairs, lambda p, r: p == 24 and r[0] in (1500, 1200), sz=True)
    penalty(1, "第24页「放大路径」标题不满足 15 磅、说明不满足 12 磅", bool(bad), _fmt(bad))

    # 37 -3：底部页脚导航不满足 7.8 磅，深色页不满足 DDF3EC、浅色页不满足 6A7F7D
    # （文本框几何按口径不再判定）
    bad = _diff_runs(run_pairs, lambda p, r: r[0] == 780, sz=True, col=True)
    bad += geo_bad_list(cls({780}), _BOX_FOOTER)
    penalty(3, "底部页脚导航不满足 7.8 磅或深/浅页色值 DDF3EC/6A7F7D", bool(bad), _fmt(bad))

    # 38 -1：页面底色不满足 FFF9EE、顶部色条不满足 0E7C7B 且高 0.22 英寸，卡片与底板
    # 填充不满足 FFFFFF/F8FFFC/F4FBFF/FFF7E8/E8F6F2 的分区取值，或圆角矩形被改为直角。
    # 逐页比较「几何形状类型 + 色值」的多重集合；遮罩与第18页横幅底色分别由第20、34条判定
    def plain_fills(page, sig_list):
        out = []
        for sig in sig_list:
            parts = sig.split(":")
            if len(parts) < 2:
                continue
            color = parts[1]
            if "/" in color and (page in mask_pages or (page == 18 and color == "000000/65000")):
                continue
            out.append(parts[0] + ":" + color)
        return sorted(out)

    bad = []
    for page, exp_slide, got_slide in page_pairs:
        want, have = plain_fills(page, exp_slide["f"]), plain_fills(page, got_slide["f"])
        if want != have:
            only_exp = sorted(set(want) - set(have))
            only_got = sorted(set(have) - set(want))
            bad.append("p%d 填充差异 -%s +%s" % (page, only_exp or "无", only_got or "无"))
        for sig in got_slide["f"]:
            parts = sig.split(":")
            if len(parts) == 3 and parts[0] == "rect" and parts[1] == "0E7C7B" and "x" in parts[2]:
                height = float(parts[2].split("x")[1])
                if abs(height - 0.22) > _POS_TOL:
                    bad.append("p%d 顶部色条高 %.2f 英寸" % (page, height))
    penalty(1, "页面底色、顶部色条、卡片底板填充或圆角矩形被改动", bool(bad), _fmt(bad))

    # 39 -5：7 张配图的图像数据未保持逐字节一致（以 ppt/media 下各文件的 md5 比对，
    # 被替换、重新编码、调色、加滤镜、去背或改为占位灰底都会改变 md5）
    hit = deck["md5"] != baseline["md5"]
    detail = ""
    if hit:
        detail = "媒体文件 %d 个（应为 %d 个），md5 不一致 %d 处" % (
            len(deck["md5"]), len(baseline["md5"]),
            len(set(baseline["md5"]) ^ set(deck["md5"])))
    penalty(5, "7 张配图的图像数据未保持逐字节一致", hit, detail)

    # 40 -3：配图的裁剪范围、摆放与叠放层次被改动（第33页配图应位于 0.7/1.45/5.6/4.9；
    # 封面与结束页配图应满版 13.333×7.5 并置于遮罩之下）
    bad = []
    for page, exp_slide, got_slide in page_pairs:
        exp_pics, got_pics = exp_slide["p"], got_slide["p"]
        if len(exp_pics) != len(got_pics):
            bad.append("p%d 图片数 %d→%d" % (page, len(exp_pics), len(got_pics)))
            continue
        for i, (exp_pic, got_pic) in enumerate(zip(exp_pics, got_pics), 1):
            if _geo_bad(got_pic[:4], exp_pic[:4]):
                bad.append("p%d 图%d 几何 %s→%s" % (page, i, exp_pic[:4], got_pic[:4]))
            if (exp_pic[4] or {}) != (got_pic[4] or {}):
                bad.append("p%d 图%d 裁剪 %s→%s" % (page, i, exp_pic[4], got_pic[4]))
        if exp_slide.get("z") and got_slide.get("z") != exp_slide["z"]:
            bad.append("p%d 叠放层次 %s→%s" % (page, exp_slide["z"], got_slide["z"]))
    for page, want in ((33, _PIC_P33), (1, _PIC_FULL), (46, _PIC_FULL)):
        slide = slide_of.get(page)
        if slide is None or not slide["p"]:
            bad.append("p%d 配图缺失" % page)
        elif all(_geo_bad(pic[:4], want) for pic in slide["p"]):
            bad.append("p%d 配图未位于 %s 英寸" % (page, want))
    penalty(3, "配图的裁剪范围、摆放与叠放层次被改动", bool(bad), _fmt(bad))

    # 41 -3：删除或改写页面上的说明性文字内容（非翻译需要的增删、改写、缩写、合并句子）。
    # 可机械判定的是「文字被删空」与「段落数减少（合并句子）」，以及整块文本框丢失
    bad = ["p%d/%s 文本框缺失" % t for t in missing]
    for page, name, exp, got in shape_pairs:
        if exp["np"] and not got["np"]:
            bad.append("p%d/%s 文字被删空" % (page, name))
        elif got["np"] < exp["np"]:
            bad.append("p%d/%s 段落 %d→%d（疑似合并句子）" % (page, name, exp["np"], got["np"]))
    penalty(3, "删除或改写页面上的说明性文字内容", bool(bad), _fmt(bad))

    # 42 -1：编号类文字被改写或丢失——目录 PART 序号、章节页两位数序号、流程步骤序号、
    # 页脚页码。按数值比较，"01"→"1" 这类零填充写法差异不计
    bad = []
    for group in (cls({4800}), cls({1150}, pages={2}), cls({1000}), cls({780})):
        for page, name, exp, got in group:
            want = {int(x) for x in exp.get("x", [])}
            have = {int(x) for x in re.findall(r"\d+", got["txt"])} if got else set()
            if want - have:
                bad.append("p%d/%s 编号 %s 丢失" % (page, name,
                                                   sorted(want - have)))
    penalty(1, "编号类文字被改写或丢失", bool(bad), _fmt(bad))

    # 43 -1：改动署名「青岚县农业低空应用课题组」、机构名、地名等信息中的实体名称取值。
    # 判据：原文出现县名的各处文字都须存在且已英译，且这些文字之间仍共享同一专有名词
    # （通用词如 county/agricultural/plan 等不参与比较）
    bad, token_sets = [], []
    for key in sorted(baseline["s"], key=int):
        page = int(key)
        for name, exp in baseline["s"][key]["t"].items():
            if not exp.get("e"):
                continue
            got = got_of.get((page, name))
            if got is None or not (got["txt"] or "").strip():
                bad.append("p%d/%s 含实体名的文字缺失或被删空" % (page, name))
                continue
            if _CJK_RE.search(got["txt"]):
                continue    # 仍是中文，属英译覆盖面问题，由加分项判定，本条不重复计扣
            words = re.findall(r"[A-Za-z][A-Za-z'\-]{3,}", got["txt"])
            token_sets.append({w.lower() for w in words} - _ENTITY_STOP)
    if token_sets and not set.intersection(*token_sets):
        bad.append("含地名/署名的 %d 处文字之间已无共同专有名词" % len(token_sets))
    penalty(1, "署名、机构名、地名等实体名称取值被改动", bool(bad), _fmt(bad))

    # 44 -5：英文使行距被压缩——全篇文本体不得被写入行距压缩比例 lnSpcReduction，
    # 34/12.5/10.5/9.2/8.2/7.8 磅等各类文字的行间距须保持原值
    total_lnsp = sum(slide["lnsp"] for slide in deck["slides"])
    penalty(5, "英文使行距被压缩（写入 lnSpcReduction）", total_lnsp > 0,
            "%d 处文本体写入 lnSpcReduction" % total_lnsp)

    return _finalize(result, results)


def _finalize(result, results):
    """把 results 反解为 dim2_items：加分项计入满分，减分项命中即扣、不计入满分。"""
    if len(results) != len(_MAX_DELTAS_ORDER):
        raise RuntimeError("dim2 项数不一致：results=%d vs order=%d"
                           % (len(results), len(_MAX_DELTAS_ORDER)))
    dim2_items, total_score, max_score = [], 0, 0
    for (raw, _detail), max_delta in zip(results, _MAX_DELTAS_ORDER):
        head, _, rule = raw.partition("：")
        try:
            magnitude = int(head.replace("+", "").replace("-", "").strip())
        except ValueError:
            magnitude = 0
        hit = magnitude > 0
        delta = max_delta if hit else 0
        total_score += delta
        if max_delta > 0:
            max_score += max_delta
        dim2_items.append({"rule": rule, "max_delta": max_delta, "delta": delta,
                           "hit": hit})
    result["dim2_items"] = dim2_items
    result["total_score"] = total_score
    result["max_score"] = max_score
    return result


def evaluate(dir_path):
    """评测入口：在 dir_path 内定位交付的 .pptx 并给出完成度评分。

    交付物只接受 .pptx：目录内优先取 .pptx；若只交了 .ppt / .pptm / .odp / .pdf 等
    其他格式，仍把它收作被评对象，由维度一判不通过（total_score 记 0），而不是当成
    「没有交付物」跳过评测。

    返回固定 9 个字段：id / file_name / status / error / dim1_pass / dim1_reason /
    dim2_items / total_score / max_score，其中 dim2_items 每项为
    {rule, max_delta, delta, hit}。

    两类异常分开回传：
      · 目录不存在 / 目录内没有任何演示文稿——维度一根本没进入判定，error 只带纯文本
        原因（不带异常类名），file_name、dim1_pass、dim1_reason 一律留空（dim1_pass 为
        空字符串而非 False），dim2 不出项，总分与满分一并归零；
      · 其他运行期异常——error 以「脚本运行异常：」打头，dim1_pass 记 False 并在
        dim1_reason 中给出错误摘要。
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
        files = sorted(f for f in os.listdir(dir_path)
                       if not f.startswith("~$")
                       and os.path.isfile(os.path.join(dir_path, f)))
        cands = [f for f in files if f.lower().endswith(_ACCEPTED_EXT)]
        if not cands:
            # 没有 .pptx，但交了 .ppt/.odp/.pdf 等其他格式：仍收作被评对象，由维度一判不通过
            cands = [f for f in files if f.lower().endswith(_WRONG_FMT_EXTS)]
        if not cands:
            raise FileNotFoundError("目录中未找到 .pptx 文件：%s" % dir_path)
        file_path = os.path.join(dir_path, cands[0])
        result["file_name"] = os.path.basename(file_path)
        return _evaluate_impl(file_path, result)
    except FileNotFoundError as exc:
        # 找不到被评估文件：error 只回传纯文本原因，维度一未进入判定，各字段留空、分数归零。
        result["status"] = "error"
        result["error"] = str(exc)
        result["file_name"] = ""
        result["dim1_pass"] = ""
        result["dim1_reason"] = ""
        result["dim2_items"] = []
        result["total_score"] = 0
        result["max_score"] = 0
        return result
    except Exception as exc:
        if type(exc).__name__ in ("ParseError", "XMLSyntaxError", "BadZipFile") or (
            isinstance(exc, KeyError) and "There is no item named" in str(exc)
        ):
            # 交付文件本身结构/XML 损坏：记为维度一不通过、正常 0 分
            result["status"] = "ok"
            result["error"] = None
            result["dim1_pass"] = False
            result["dim1_reason"] = "文件无法打开或解析：%s: %s" % (
                type(exc).__name__, exc,
            )
            result["dim2_items"] = []
            result["total_score"] = 0
            return result
        result["status"] = "error"
        result["error"] = "脚本运行异常：%s: %s" % (type(exc).__name__, exc)
        result["dim1_pass"] = False
        result["dim1_reason"] = "文件无法打开或解析：%s: %s" % (type(exc).__name__, exc)
        result["dim2_items"] = []
        result["total_score"] = 0
        result["max_score"] = 0
        return result


_BASELINE_JSON = r"""{"cx":13.3326,"cy":7.5,"md5":["247b0d69ea34a5bd009ab93dcb7071fc","6b0d9e0926904e53ad8e8b5a0ce9e072","81c4399df273b45b134df34b8c76cb38","b1a5fa435c5661dd9ff4d9f838cbf38b","c226db814c0e67a1775e12ed89efaffa","dd40b802097fa2990c055c1d5ddee6cc","f6e35fe11d2ca8d634ec0c83ef25e750"],"n":46,"s":{"1":{"f":["rect:000000/82000:13.33x7.50","rect:063D3B/92000:6.10x6.95","roundRect:E5A63B:1.85x0.28"],"no":1,"p":[[0.0,0.0,13.333,7.5,{"l":1,"r":1}]],"pic":1,"sp":8,"t":{"Text 2":{"b":[0.66,1.7,5.1,1.3],"c":1,"e":1,"np":1,"r":[[3200,1,"FFFFFF","l"]]},"Text 3":{"c":1,"np":1,"r":[[1300,0,"EAF7F4","l"]]},"Text 5":{"c":1,"np":1,"r":[[820,1,"FFFFFF","ctr"]]},"Text 7":{"b":[0.55,7.14,12.15,0.22],"c":1,"np":1,"r":[[780,0,"DDF3EC","l"]],"x":["01"]}},"z":"pssssssss"},"10":{"f":["rect:0E7C7B:13.33x0.22","rect:FFF9EE:13.33x7.50","roundRect:F8FFFC:5.30x0.90","roundRect:F8FFFC:5.30x0.90","roundRect:F8FFFC:5.30x0.90","roundRect:FFF7E8:5.30x0.90"],"no":10,"p":[[6.75,1.45,5.85,4.65,{"l":14617,"r":14617}]],"pic":1,"sp":18,"t":{"Text 11":{"c":1,"np":1,"r":[[1250,1,"063D3B","l"]]},"Text 12":{"c":1,"np":1,"r":[[920,0,"112B2D","l"]]},"Text 14":{"c":1,"np":1,"r":[[1250,1,"063D3B","l"]]},"Text 15":{"c":1,"np":1,"r":[[920,0,"112B2D","l"]]},"Text 17":{"b":[0.55,7.14,12.15,0.22],"c":1,"np":1,"r":[[780,0,"6A7F7D","l"]],"x":["10"]},"Text 2":{"b":[0.65,0.52,9.3,0.5],"c":1,"np":1,"r":[[2600,1,"112B2D","l"]]},"Text 3":{"c":1,"np":1,"r":[[1100,0,"6A7F7D","l"]]},"Text 5":{"c":1,"np":1,"r":[[1250,1,"063D3B","l"]]},"Text 6":{"c":1,"np":1,"r":[[920,0,"112B2D","l"]]},"Text 8":{"c":1,"np":1,"r":[[1250,1,"063D3B","l"]]},"Text 9":{"c":1,"np":1,"r":[[920,0,"112B2D","l"]]}},"z":"sssspssssssssssssss"},"11":{"f":["rect:0E7C7B:13.33x0.22","rect:FFF9EE:13.33x7.50","roundRect:FFFFFF:4.80x0.75","roundRect:FFFFFF:4.80x0.75","roundRect:FFFFFF:4.80x0.75","roundRect:FFFFFF:4.80x0.75"],"no":11,"p":[[0.7,1.45,6.4,4.95,{"l":13636,"r":13636}]],"pic":1,"sp":19,"t":{"Text 11":{"c":1,"np":1,"r":[[1250,1,"063D3B","l"]]},"Text 12":{"c":1,"np":1,"r":[[920,0,"112B2D","l"]]},"Text 14":{"c":1,"np":1,"r":[[1250,1,"063D3B","l"]]},"Text 15":{"c":1,"np":1,"r":[[920,0,"112B2D","l"]]},"Text 16":{"c":1,"np":1,"r":[[1100,1,"0E7C7B","l"]]},"Text 18":{"b":[0.55,7.14,12.15,0.22],"c":1,"np":1,"r":[[780,0,"6A7F7D","l"]],"x":["11"]},"Text 2":{"b":[0.65,0.52,9.3,0.5],"c":1,"np":1,"r":[[2600,1,"112B2D","l"]]},"Text 3":{"c":1,"np":1,"r":[[1100,0,"6A7F7D","l"]]},"Text 5":{"c":1,"np":1,"r":[[1250,1,"063D3B","l"]]},"Text 6":{"c":1,"np":1,"r":[[920,0,"112B2D","l"]]},"Text 8":{"c":1,"np":1,"r":[[1250,1,"063D3B","l"]]},"Text 9":{"c":1,"np":1,"r":[[920,0,"112B2D","l"]]}},"z":"sssspsssssssssssssss"},"12":{"f":["rect:0E7C7B:13.33x0.22","rect:FFF9EE:13.33x7.50","roundRect:FFF7E8:6.17x1.55","roundRect:FFFFFF:2.90x1.14","roundRect:FFFFFF:2.90x1.14","roundRect:FFFFFF:2.90x1.14","roundRect:FFFFFF:2.90x1.14"],"no":12,"p":[[0.72,1.45,5.2,4.8,{"l":19531,"r":19531}]],"pic":1,"sp":21,"t":{"Text 11":{"c":1,"np":1,"r":[[1250,1,"063D3B","l"]]},"Text 12":{"c":1,"np":1,"r":[[920,0,"112B2D","l"]]},"Text 14":{"c":1,"np":1,"r":[[1250,1,"063D3B","l"]]},"Text 15":{"c":1,"np":1,"r":[[920,0,"112B2D","l"]]},"Text 17":{"c":1,"np":1,"r":[[1300,1,"063D3B","l"]]},"Text 18":{"c":1,"np":1,"r":[[1050,0,"112B2D","l"]]},"Text 2":{"b":[0.65,0.52,9.3,0.5],"c":1,"np":1,"r":[[2600,1,"112B2D","l"]]},"Text 20":{"b":[0.55,7.14,12.15,0.22],"c":1,"np":1,"r":[[780,0,"6A7F7D","l"]],"x":["12"]},"Text 3":{"c":1,"np":1,"r":[[1100,0,"6A7F7D","l"]]},"Text 5":{"c":1,"np":1,"r":[[1250,1,"063D3B","l"]]},"Text 6":{"c":1,"np":1,"r":[[920,0,"112B2D","l"]]},"Text 8":{"c":1,"np":1,"r":[[1250,1,"063D3B","l"]]},"Text 9":{"c":1,"np":1,"r":[[920,0,"112B2D","l"]]}},"z":"sssspsssssssssssssssss"},"13":{"f":["rect:0E7C7B:13.33x0.22","rect:FFF9EE:13.33x7.50","roundRect:F8FFFC:5.40x0.80","roundRect:FFFFFF:5.40x1.00","roundRect:FFFFFF:5.40x1.00","roundRect:FFFFFF:5.40x1.00"],"no":13,"p":[[0.72,1.45,5.8,4.6,{"l":14538,"r":14538}]],"pic":1,"sp":18,"t":{"Text 11":{"c":1,"np":1,"r":[[1250,1,"063D3B","l"]]},"Text 12":{"c":1,"np":1,"r":[[920,0,"112B2D","l"]]},"Text 14":{"c":1,"np":1,"r":[[1250,1,"063D3B","l"]]},"Text 15":{"c":1,"np":1,"r":[[920,0,"112B2D","l"]]},"Text 17":{"b":[0.55,7.14,12.15,0.22],"c":1,"np":1,"r":[[780,0,"6A7F7D","l"]],"x":["13"]},"Text 2":{"b":[0.65,0.52,9.3,0.5],"c":1,"np":1,"r":[[2600,1,"112B2D","l"]]},"Text 3":{"c":1,"np":1,"r":[[1100,0,"6A7F7D","l"]]},"Text 5":{"c":1,"np":1,"r":[[1250,1,"063D3B","l"]]},"Text 6":{"c":1,"np":1,"r":[[920,0,"112B2D","l"]]},"Text 8":{"c":1,"np":1,"r":[[1250,1,"063D3B","l"]]},"Text 9":{"c":1,"np":1,"r":[[920,0,"112B2D","l"]]}},"z":"sssspssssssssssssss"},"14":{"f":["rect:000000/62000:13.33x7.50","rect:063D3B/78000:13.33x7.50"],"no":14,"p":[[0.0,0.0,13.333,7.5,{"l":1,"r":1}]],"pic":1,"sp":8,"t":{"Text 2":{"np":1,"r":[[4800,1,"E5A63B","l"]],"x":["03"]},"Text 3":{"c":1,"np":1,"r":[[3400,1,"FFFFFF","l"]]},"Text 5":{"c":1,"np":1,"r":[[1400,0,"EAF7F4","l"]]},"Text 7":{"b":[0.55,7.14,12.15,0.22],"c":1,"np":1,"r":[[780,0,"DDF3EC","l"]],"x":["14"]}},"z":"pssssssss"},"15":{"f":["rect:0E7C7B:13.33x0.22","rect:FFF9EE:13.33x7.50","roundRect:F8FFFC:3.60x1.30","roundRect:F8FFFC:3.60x1.30","roundRect:F8FFFC:3.60x1.30","roundRect:FFFFFF:2.60x1.62","roundRect:FFFFFF:2.60x1.62","roundRect:FFFFFF:2.60x1.62","roundRect:FFFFFF:2.60x1.62"],"no":15,"p":[],"pic":0,"sp":31,"t":{"Text 10":{"c":1,"np":1,"r":[[1100,1,"112B2D","ctr"]]},"Text 11":{"c":1,"np":1,"r":[[820,0,"6A7F7D","ctr"]]},"Text 13":{"c":1,"np":1,"r":[[2500,1,"E5A63B","ctr"]]},"Text 14":{"c":1,"np":1,"r":[[1100,1,"112B2D","ctr"]]},"Text 15":{"c":1,"np":1,"r":[[820,0,"6A7F7D","ctr"]]},"Text 17":{"c":1,"np":1,"r":[[2500,1,"C96B3C","ctr"]]},"Text 18":{"c":1,"np":1,"r":[[1100,1,"112B2D","ctr"]]},"Text 19":{"c":1,"np":1,"r":[[820,0,"6A7F7D","ctr"]]},"Text 2":{"b":[0.65,0.52,9.3,0.5],"c":1,"np":1,"r":[[2600,1,"112B2D","l"]]},"Text 21":{"c":1,"np":1,"r":[[1250,1,"063D3B","l"]]},"Text 22":{"c":1,"np":1,"r":[[920,0,"112B2D","l"]]},"Text 24":{"c":1,"np":1,"r":[[1250,1,"063D3B","l"]]},"Text 25":{"c":1,"np":1,"r":[[920,0,"112B2D","l"]]},"Text 27":{"c":1,"np":1,"r":[[1250,1,"063D3B","l"]]},"Text 28":{"c":1,"np":1,"r":[[920,0,"112B2D","l"]]},"Text 3":{"c":1,"np":1,"r":[[1100,0,"6A7F7D","l"]]},"Text 30":{"b":[0.55,7.14,12.15,0.22],"c":1,"np":1,"r":[[780,0,"6A7F7D","l"]],"x":["15"]},"Text 5":{"c":1,"np":1,"r":[[2500,1,"0E7C7B","ctr"]]},"Text 6":{"c":1,"np":1,"r":[[1100,1,"112B2D","ctr"]]},"Text 7":{"c":1,"np":1,"r":[[820,0,"6A7F7D","ctr"]]},"Text 9":{"c":1,"np":1,"r":[[2500,1,"1F6E86","ctr"]]}},"z":"sssssssssssssssssssssssssssssss"},"16":{"f":["rect:0E7C7B:13.33x0.22","rect:FFF9EE:13.33x7.50","roundRect:F4FBFF:2.40x0.82","roundRect:F4FBFF:2.40x0.82","roundRect:F8FFFC:2.40x0.82","roundRect:F8FFFC:2.40x0.82","roundRect:FFF7E8:4.75x0.82","roundRect:FFF7E8:4.75x0.82","roundRect:FFF7E8:4.75x0.82","roundRect:FFF7E8:4.75x0.82","roundRect:FFFFFF:4.20x0.82","roundRect:FFFFFF:4.20x0.82","roundRect:FFFFFF:4.20x0.82","roundRect:FFFFFF:4.20x0.82"],"no":16,"p":[],"pic":0,"sp":43,"t":{"Text 11":{"c":1,"np":1,"r":[[1050,1,"063D3B","l"]]},"Text 12":{"c":1,"np":1,"r":[[870,0,"112B2D","l"]]},"Text 14":{"c":1,"np":1,"r":[[1300,1,"063D3B","l"]]},"Text 17":{"c":1,"np":1,"r":[[1050,1,"063D3B","l"]]},"Text 18":{"c":1,"np":1,"r":[[870,0,"112B2D","l"]]},"Text 2":{"b":[0.65,0.52,9.3,0.5],"c":1,"np":1,"r":[[2600,1,"112B2D","l"]]},"Text 20":{"c":1,"np":1,"r":[[1050,1,"063D3B","l"]]},"Text 21":{"c":1,"np":1,"r":[[870,0,"112B2D","l"]]},"Text 23":{"c":1,"np":1,"r":[[1300,1,"063D3B","l"]]},"Text 26":{"c":1,"np":1,"r":[[1050,1,"063D3B","l"]]},"Text 27":{"c":1,"np":1,"r":[[870,0,"112B2D","l"]]},"Text 29":{"c":1,"np":1,"r":[[1050,1,"063D3B","l"]]},"Text 3":{"c":1,"np":1,"r":[[1100,0,"6A7F7D","l"]]},"Text 30":{"c":1,"np":1,"r":[[870,0,"112B2D","l"]]},"Text 32":{"c":1,"np":1,"r":[[1300,1,"063D3B","l"]]},"Text 35":{"c":1,"np":1,"r":[[1050,1,"063D3B","l"]]},"Text 36":{"c":1,"np":1,"r":[[870,0,"112B2D","l"]]},"Text 38":{"c":1,"np":1,"r":[[1050,1,"063D3B","l"]]},"Text 39":{"c":1,"np":1,"r":[[870,0,"112B2D","l"]]},"Text 40":{"c":1,"np":1,"r":[[1300,1,"0E7C7B","ctr"]]},"Text 42":{"b":[0.55,7.14,12.15,0.22],"c":1,"np":1,"r":[[780,0,"6A7F7D","l"]],"x":["16"]},"Text 5":{"c":1,"np":1,"r":[[1300,1,"063D3B","l"]]},"Text 8":{"c":1,"np":1,"r":[[1050,1,"063D3B","l"]]},"Text 9":{"c":1,"np":1,"r":[[870,0,"112B2D","l"]]}},"z":"sssssssssssssssssssssssssssssssssssssssssss"},"17":{"f":["rect:000000/62000:13.33x7.50","rect:063D3B/78000:13.33x7.50"],"no":17,"p":[[0.0,0.0,13.333,7.5,{"l":1,"r":1}]],"pic":1,"sp":8,"t":{"Text 2":{"np":1,"r":[[4800,1,"E5A63B","l"]],"x":["04"]},"Text 3":{"c":1,"np":1,"r":[[3400,1,"FFFFFF","l"]]},"Text 5":{"c":1,"np":1,"r":[[1400,0,"EAF7F4","l"]]},"Text 7":{"b":[0.55,7.14,12.15,0.22],"c":1,"np":1,"r":[[780,0,"DDF3EC","l"]],"x":["17"]}},"z":"pssssssss"},"18":{"f":["rect:000000/65000:11.62x1.00","rect:0E7C7B:13.33x0.22","rect:FFF9EE:13.33x7.50","roundRect:FFFFFF:5.55x1.10","roundRect:FFFFFF:5.55x1.10","roundRect:FFFFFF:5.55x1.10","roundRect:FFFFFF:5.55x1.10"],"no":18,"p":[[0.85,5.0,11.62,1.0,{"b":42350,"t":42350}]],"pic":1,"sp":20,"t":{"Text 11":{"c":1,"np":1,"r":[[1250,1,"063D3B","l"]]},"Text 12":{"c":1,"np":1,"r":[[920,0,"112B2D","l"]]},"Text 14":{"c":1,"np":1,"r":[[1250,1,"063D3B","l"]]},"Text 15":{"c":1,"np":1,"r":[[920,0,"112B2D","l"]]},"Text 17":{"c":1,"np":1,"r":[[1600,1,"FFFFFF","ctr"]]},"Text 19":{"b":[0.55,7.14,12.15,0.22],"c":1,"np":1,"r":[[780,0,"6A7F7D","l"]],"x":["18"]},"Text 2":{"b":[0.65,0.52,9.3,0.5],"c":1,"np":1,"r":[[2600,1,"112B2D","l"]]},"Text 3":{"c":1,"np":1,"r":[[1100,0,"6A7F7D","l"]]},"Text 5":{"c":1,"np":1,"r":[[1250,1,"063D3B","l"]]},"Text 6":{"c":1,"np":1,"r":[[920,0,"112B2D","l"]]},"Text 8":{"c":1,"np":1,"r":[[1250,1,"063D3B","l"]]},"Text 9":{"c":1,"np":1,"r":[[920,0,"112B2D","l"]]}},"z":"sssssssssssssssspssss"},"19":{"f":["rect:000000/62000:13.33x7.50","rect:063D3B/78000:13.33x7.50"],"no":19,"p":[[0.0,0.0,13.333,7.5,{"l":1,"r":1}]],"pic":1,"sp":8,"t":{"Text 2":{"np":1,"r":[[4800,1,"E5A63B","l"]],"x":["05"]},"Text 3":{"c":1,"np":1,"r":[[3400,1,"FFFFFF","l"]]},"Text 5":{"c":1,"np":1,"r":[[1400,0,"EAF7F4","l"]]},"Text 7":{"b":[0.55,7.14,12.15,0.22],"c":1,"np":1,"r":[[780,0,"DDF3EC","l"]],"x":["19"]}},"z":"pssssssss"},"2":{"f":["rect:0E7C7B:13.33x0.22","rect:FFF9EE:13.33x7.50","roundRect:F8FFFC:5.20x0.78","roundRect:F8FFFC:5.20x0.78","roundRect:F8FFFC:5.20x0.78","roundRect:FFFFFF:5.20x0.78","roundRect:FFFFFF:5.20x0.78","roundRect:FFFFFF:5.20x0.78","roundRect:FFFFFF:5.20x0.78"],"no":2,"p":[],"pic":0,"sp":27,"t":{"Text 11":{"c":1,"np":1,"r":[[1150,1,"0E7C7B","l"]],"x":["03"]},"Text 12":{"c":1,"np":1,"r":[[840,0,"112B2D","l"]]},"Text 14":{"c":1,"np":1,"r":[[1150,1,"1F6E86","l"]],"x":["04"]},"Text 15":{"c":1,"np":1,"r":[[840,0,"112B2D","l"]]},"Text 17":{"c":1,"np":1,"r":[[1150,1,"0E7C7B","l"]],"x":["05"]},"Text 18":{"c":1,"np":1,"r":[[840,0,"112B2D","l"]]},"Text 2":{"b":[0.65,0.52,9.3,0.5],"c":1,"np":1,"r":[[2600,1,"112B2D","l"]]},"Text 20":{"c":1,"np":1,"r":[[1150,1,"1F6E86","l"]],"x":["06"]},"Text 21":{"c":1,"np":1,"r":[[840,0,"112B2D","l"]]},"Text 23":{"c":1,"np":1,"r":[[1150,1,"0E7C7B","l"]],"x":["07"]},"Text 24":{"c":1,"np":1,"r":[[840,0,"112B2D","l"]]},"Text 26":{"b":[0.55,7.14,12.15,0.22],"c":1,"np":1,"r":[[780,0,"6A7F7D","l"]],"x":["02"]},"Text 3":{"c":1,"np":1,"r":[[1100,0,"6A7F7D","l"]]},"Text 5":{"c":1,"np":1,"r":[[1150,1,"0E7C7B","l"]],"x":["01"]},"Text 6":{"c":1,"np":1,"r":[[840,0,"112B2D","l"]]},"Text 8":{"c":1,"np":1,"r":[[1150,1,"1F6E86","l"]],"x":["02"]},"Text 9":{"c":1,"np":1,"r":[[840,0,"112B2D","l"]]}},"z":"sssssssssssssssssssssssssss"},"20":{"f":["rect:0E7C7B:13.33x0.22","rect:FFF9EE:13.33x7.50","roundRect:FFF7E8:5.30x0.70","roundRect:FFFFFF:5.30x0.95","roundRect:FFFFFF:5.30x0.95","roundRect:FFFFFF:5.30x0.95"],"no":20,"p":[[6.75,1.4,5.75,4.9,{"l":16996,"r":16996}]],"pic":1,"sp":18,"t":{"Text 11":{"c":1,"np":1,"r":[[1250,1,"063D3B","l"]]},"Text 12":{"c":1,"np":1,"r":[[920,0,"112B2D","l"]]},"Text 14":{"c":1,"np":1,"r":[[1250,1,"063D3B","l"]]},"Text 15":{"c":1,"np":1,"r":[[880,0,"112B2D","l"]]},"Text 17":{"b":[0.55,7.14,12.15,0.22],"c":1,"np":1,"r":[[780,0,"6A7F7D","l"]],"x":["20"]},"Text 2":{"b":[0.65,0.52,9.3,0.5],"c":1,"np":1,"r":[[2600,1,"112B2D","l"]]},"Text 3":{"c":1,"np":1,"r":[[1100,0,"6A7F7D","l"]]},"Text 5":{"c":1,"np":1,"r":[[1250,1,"063D3B","l"]]},"Text 6":{"c":1,"np":1,"r":[[920,0,"112B2D","l"]]},"Text 8":{"c":1,"np":1,"r":[[1250,1,"063D3B","l"]]},"Text 9":{"c":1,"np":1,"r":[[920,0,"112B2D","l"]]}},"z":"sssspssssssssssssss"},"21":{"f":["rect:0E7C7B:13.33x0.22","rect:FFF9EE:13.33x7.50","roundRect:F8FFFC:5.40x0.62","roundRect:FFFFFF:5.40x1.10","roundRect:FFFFFF:5.40x1.10","roundRect:FFFFFF:5.40x1.10"],"no":21,"p":[[0.75,1.45,5.7,4.9,{"l":17283,"r":17283}]],"pic":1,"sp":18,"t":{"Text 11":{"c":1,"np":1,"r":[[1250,1,"063D3B","l"]]},"Text 12":{"c":1,"np":1,"r":[[920,0,"112B2D","l"]]},"Text 14":{"c":1,"np":1,"r":[[1250,1,"063D3B","l"]]},"Text 15":{"c":1,"np":1,"r":[[920,0,"112B2D","l"]]},"Text 17":{"b":[0.55,7.14,12.15,0.22],"c":1,"np":1,"r":[[780,0,"6A7F7D","l"]],"x":["21"]},"Text 2":{"b":[0.65,0.52,9.3,0.5],"c":1,"np":1,"r":[[2600,1,"112B2D","l"]]},"Text 3":{"c":1,"np":1,"r":[[1100,0,"6A7F7D","l"]]},"Text 5":{"c":1,"np":1,"r":[[1250,1,"063D3B","l"]]},"Text 6":{"c":1,"np":1,"r":[[920,0,"112B2D","l"]]},"Text 8":{"c":1,"np":1,"r":[[1250,1,"063D3B","l"]]},"Text 9":{"c":1,"np":1,"r":[[920,0,"112B2D","l"]]}},"z":"sssspssssssssssssss"},"22":{"f":["rect:0E7C7B:13.33x0.22","rect:FFF9EE:13.33x7.50","roundRect:E8F6F2:2.20x0.82","roundRect:E8F6F2:2.20x0.82","roundRect:FFF7E8:5.60x1.20","roundRect:FFFFFF:2.20x0.82","roundRect:FFFFFF:2.20x0.82","roundRect:FFFFFF:2.20x0.82","roundRect:FFFFFF:2.65x1.20","roundRect:FFFFFF:2.65x1.20","roundRect:FFFFFF:2.65x1.20","roundRect:FFFFFF:2.65x1.20"],"no":22,"p":[],"pic":0,"sp":36,"t":{"Text 11":{"c":1,"np":1,"r":[[1250,1,"063D3B","l"]]},"Text 12":{"c":1,"np":1,"r":[[920,0,"112B2D","l"]]},"Text 14":{"c":1,"np":1,"r":[[1250,1,"063D3B","l"]]},"Text 15":{"c":1,"np":1,"r":[[920,0,"112B2D","l"]]},"Text 17":{"c":1,"np":1,"r":[[1250,1,"063D3B","l"]]},"Text 18":{"c":1,"np":1,"r":[[920,0,"112B2D","l"]]},"Text 2":{"b":[0.65,0.52,9.3,0.5],"c":1,"np":1,"r":[[2600,1,"112B2D","l"]]},"Text 20":{"np":1,"r":[[1000,1,"E5A63B","l"]],"x":["1"]},"Text 21":{"c":1,"np":1,"r":[[1000,0,"112B2D","l"]]},"Text 23":{"np":1,"r":[[1000,1,"E5A63B","l"]],"x":["2"]},"Text 24":{"c":1,"np":1,"r":[[1000,0,"112B2D","l"]]},"Text 26":{"np":1,"r":[[1000,1,"E5A63B","l"]],"x":["3"]},"Text 27":{"c":1,"np":1,"r":[[1000,0,"112B2D","l"]]},"Text 29":{"np":1,"r":[[1000,1,"E5A63B","l"]],"x":["4"]},"Text 3":{"c":1,"np":1,"r":[[1100,0,"6A7F7D","l"]]},"Text 30":{"c":1,"np":1,"r":[[1000,0,"112B2D","l"]]},"Text 32":{"np":1,"r":[[1000,1,"E5A63B","l"]],"x":["5"]},"Text 33":{"c":1,"np":1,"r":[[1000,0,"112B2D","l"]]},"Text 35":{"b":[0.55,7.14,12.15,0.22],"c":1,"np":1,"r":[[780,0,"6A7F7D","l"]],"x":["22"]},"Text 5":{"c":1,"np":1,"r":[[1250,1,"063D3B","l"]]},"Text 6":{"c":1,"np":1,"r":[[920,0,"112B2D","l"]]},"Text 8":{"c":1,"np":1,"r":[[1250,1,"063D3B","l"]]},"Text 9":{"c":1,"np":1,"r":[[920,0,"112B2D","l"]]}},"z":"ssssssssssssssssssssssssssssssssssss"},"23":{"f":["rect:000000/62000:13.33x7.50","rect:063D3B/78000:13.33x7.50"],"no":23,"p":[[0.0,0.0,13.333,7.5,{"l":1,"r":1}]],"pic":1,"sp":8,"t":{"Text 2":{"np":1,"r":[[4800,1,"E5A63B","l"]],"x":["06"]},"Text 3":{"c":1,"np":1,"r":[[3400,1,"FFFFFF","l"]]},"Text 5":{"c":1,"np":1,"r":[[1400,0,"EAF7F4","l"]]},"Text 7":{"b":[0.55,7.14,12.15,0.22],"c":1,"np":1,"r":[[780,0,"DDF3EC","l"]],"x":["23"]}},"z":"pssssssss"},"24":{"f":["rect:0E7C7B:13.33x0.22","rect:FFF9EE:13.33x7.50","roundRect:F8FFFC:11.30x1.35","roundRect:FFFFFF:2.60x1.62","roundRect:FFFFFF:2.60x1.62","roundRect:FFFFFF:2.60x1.62","roundRect:FFFFFF:2.60x1.62"],"no":24,"p":[],"pic":0,"sp":25,"t":{"Text 10":{"c":1,"np":1,"r":[[1100,1,"112B2D","ctr"]]},"Text 11":{"c":1,"np":1,"r":[[820,0,"6A7F7D","ctr"]]},"Text 13":{"c":1,"np":1,"r":[[2500,1,"E5A63B","ctr"]]},"Text 14":{"c":1,"np":1,"r":[[1100,1,"112B2D","ctr"]]},"Text 15":{"c":1,"np":1,"r":[[820,0,"6A7F7D","ctr"]]},"Text 17":{"c":1,"np":1,"r":[[2500,1,"C96B3C","ctr"]]},"Text 18":{"c":1,"np":1,"r":[[1100,1,"112B2D","ctr"]]},"Text 19":{"c":1,"np":1,"r":[[820,0,"6A7F7D","ctr"]]},"Text 2":{"b":[0.65,0.52,9.3,0.5],"c":1,"np":1,"r":[[2600,1,"112B2D","l"]]},"Text 21":{"c":1,"np":1,"r":[[1500,1,"063D3B","l"]]},"Text 22":{"c":1,"np":1,"r":[[1200,0,"112B2D","l"]]},"Text 24":{"b":[0.55,7.14,12.15,0.22],"c":1,"np":1,"r":[[780,0,"6A7F7D","l"]],"x":["24"]},"Text 3":{"c":1,"np":1,"r":[[1100,0,"6A7F7D","l"]]},"Text 5":{"c":1,"np":1,"r":[[2500,1,"0E7C7B","ctr"]]},"Text 6":{"c":1,"np":1,"r":[[1100,1,"112B2D","ctr"]]},"Text 7":{"c":1,"np":1,"r":[[820,0,"6A7F7D","ctr"]]},"Text 9":{"c":1,"np":1,"r":[[2500,1,"1F6E86","ctr"]]}},"z":"sssssssssssssssssssssssss"},"25":{"f":["rect:0E7C7B:13.33x0.22","rect:FFF9EE:13.33x7.50","roundRect:FFF7E8:5.35x0.90","roundRect:FFFFFF:5.35x0.90","roundRect:FFFFFF:5.35x0.90","roundRect:FFFFFF:5.35x0.90"],"no":25,"p":[[0.7,1.45,5.85,4.95,{"l":16761,"r":16761}]],"pic":1,"sp":18,"t":{"Text 11":{"c":1,"np":1,"r":[[1250,1,"063D3B","l"]]},"Text 12":{"c":1,"np":1,"r":[[920,0,"112B2D","l"]]},"Text 14":{"c":1,"np":1,"r":[[1250,1,"063D3B","l"]]},"Text 15":{"c":1,"np":1,"r":[[920,0,"112B2D","l"]]},"Text 17":{"b":[0.55,7.14,12.15,0.22],"c":1,"np":1,"r":[[780,0,"6A7F7D","l"]],"x":["25"]},"Text 2":{"b":[0.65,0.52,9.3,0.5],"c":1,"np":1,"r":[[2600,1,"112B2D","l"]]},"Text 3":{"c":1,"np":1,"r":[[1100,0,"6A7F7D","l"]]},"Text 5":{"c":1,"np":1,"r":[[1250,1,"063D3B","l"]]},"Text 6":{"c":1,"np":1,"r":[[920,0,"112B2D","l"]]},"Text 8":{"c":1,"np":1,"r":[[1250,1,"063D3B","l"]]},"Text 9":{"c":1,"np":1,"r":[[920,0,"112B2D","l"]]}},"z":"sssspssssssssssssss"},"26":{"f":["rect:0E7C7B:13.33x0.22","rect:FFF9EE:13.33x7.50","roundRect:F8FFFC:5.75x0.72","roundRect:FFFFFF:5.75x1.00","roundRect:FFFFFF:5.75x1.00","roundRect:FFFFFF:5.75x1.00"],"no":26,"p":[[0.7,1.45,5.5,4.95,{"l":18750,"r":18750}]],"pic":1,"sp":18,"t":{"Text 11":{"c":1,"np":1,"r":[[1250,1,"063D3B","l"]]},"Text 12":{"c":1,"np":1,"r":[[920,0,"112B2D","l"]]},"Text 14":{"c":1,"np":1,"r":[[1250,1,"063D3B","l"]]},"Text 15":{"c":1,"np":1,"r":[[920,0,"112B2D","l"]]},"Text 17":{"b":[0.55,7.14,12.15,0.22],"c":1,"np":1,"r":[[780,0,"6A7F7D","l"]],"x":["26"]},"Text 2":{"b":[0.65,0.52,9.3,0.5],"c":1,"np":1,"r":[[2600,1,"112B2D","l"]]},"Text 3":{"c":1,"np":1,"r":[[1100,0,"6A7F7D","l"]]},"Text 5":{"c":1,"np":1,"r":[[1250,1,"063D3B","l"]]},"Text 6":{"c":1,"np":1,"r":[[920,0,"112B2D","l"]]},"Text 8":{"c":1,"np":1,"r":[[1250,1,"063D3B","l"]]},"Text 9":{"c":1,"np":1,"r":[[920,0,"112B2D","l"]]}},"z":"sssspssssssssssssss"},"27":{"f":["rect:0E7C7B:13.33x0.22","rect:FFF9EE:13.33x7.50","roundRect:FFF7E8:5.50x0.70","roundRect:FFFFFF:5.50x1.00","roundRect:FFFFFF:5.50x1.00","roundRect:FFFFFF:5.50x1.00"],"no":27,"p":[[7.0,1.45,5.35,4.95,{"l":19602,"r":19602}]],"pic":1,"sp":18,"t":{"Text 11":{"c":1,"np":1,"r":[[1250,1,"063D3B","l"]]},"Text 12":{"c":1,"np":1,"r":[[920,0,"112B2D","l"]]},"Text 14":{"c":1,"np":1,"r":[[1250,1,"063D3B","l"]]},"Text 15":{"c":1,"np":1,"r":[[3600,0,"112B2D","l"],[920,0,"112B2D","l"]]},"Text 17":{"b":[0.55,7.14,12.15,0.22],"c":1,"np":1,"r":[[780,0,"6A7F7D","l"]],"x":["27"]},"Text 2":{"b":[0.65,0.52,9.3,0.5],"c":1,"np":1,"r":[[2600,1,"112B2D","l"]]},"Text 3":{"c":1,"np":1,"r":[[1100,0,"6A7F7D","l"]]},"Text 5":{"c":1,"np":1,"r":[[1250,1,"063D3B","l"]]},"Text 6":{"c":1,"np":1,"r":[[920,0,"112B2D","l"]]},"Text 8":{"c":1,"np":1,"r":[[1250,1,"063D3B","l"]]},"Text 9":{"c":1,"np":1,"r":[[920,0,"112B2D","l"]]}},"z":"sssspssssssssssssss"},"28":{"f":["rect:0E7C7B:13.33x0.22","rect:FFF9EE:13.33x7.50","roundRect:F8FFFC:5.75x0.92","roundRect:FFFFFF:5.75x0.92","roundRect:FFFFFF:5.75x0.92","roundRect:FFFFFF:5.75x0.92"],"no":28,"p":[[0.75,1.45,5.4,4.95,{"l":19318,"r":19318}]],"pic":1,"sp":18,"t":{"Text 11":{"c":1,"np":1,"r":[[1250,1,"063D3B","l"]]},"Text 12":{"c":1,"np":1,"r":[[920,0,"112B2D","l"]]},"Text 14":{"c":1,"np":1,"r":[[1250,1,"063D3B","l"]]},"Text 15":{"c":1,"np":1,"r":[[920,0,"112B2D","l"]]},"Text 17":{"b":[0.55,7.14,12.15,0.22],"c":1,"np":1,"r":[[780,0,"6A7F7D","l"]],"x":["28"]},"Text 2":{"b":[0.65,0.52,9.3,0.5],"c":1,"np":1,"r":[[2600,1,"112B2D","l"]]},"Text 3":{"c":1,"np":1,"r":[[1100,0,"6A7F7D","l"]]},"Text 5":{"c":1,"np":1,"r":[[1250,1,"063D3B","l"]]},"Text 6":{"c":1,"np":1,"r":[[920,0,"112B2D","l"]]},"Text 8":{"c":1,"np":1,"r":[[1250,1,"063D3B","l"]]},"Text 9":{"c":1,"np":1,"r":[[920,0,"112B2D","l"]]}},"z":"sssspssssssssssssss"},"29":{"f":["rect:0E7C7B:13.33x0.22","rect:FFF9EE:13.33x7.50","roundRect:FFF7E8:5.50x0.90","roundRect:FFFFFF:5.50x0.90","roundRect:FFFFFF:5.50x0.90","roundRect:FFFFFF:5.50x0.90"],"no":29,"p":[[7.0,1.45,5.35,4.95,{"l":19602,"r":19602}]],"pic":1,"sp":18,"t":{"Text 11":{"c":1,"np":1,"r":[[1250,1,"063D3B","l"]]},"Text 12":{"c":1,"np":1,"r":[[920,0,"112B2D","l"]]},"Text 14":{"c":1,"np":1,"r":[[1250,1,"063D3B","l"]]},"Text 15":{"c":1,"np":1,"r":[[920,0,"112B2D","l"]]},"Text 17":{"b":[0.55,7.14,12.15,0.22],"c":1,"np":1,"r":[[780,0,"6A7F7D","l"]],"x":["29"]},"Text 2":{"b":[0.65,0.52,9.3,0.5],"c":1,"np":1,"r":[[2600,1,"112B2D","l"]]},"Text 3":{"c":1,"np":1,"r":[[1100,0,"6A7F7D","l"]]},"Text 5":{"c":1,"np":1,"r":[[1250,1,"063D3B","l"]]},"Text 6":{"c":1,"np":1,"r":[[920,0,"112B2D","l"]]},"Text 8":{"c":1,"np":1,"r":[[1250,1,"063D3B","l"]]},"Text 9":{"c":1,"np":1,"r":[[920,0,"112B2D","l"]]}},"z":"sssspssssssssssssss"},"3":{"f":["rect:000000/62000:13.33x7.50","rect:063D3B/78000:13.33x7.50"],"no":3,"p":[[0.0,0.0,13.333,7.5,{"l":1,"r":1}]],"pic":1,"sp":8,"t":{"Text 2":{"np":1,"r":[[4800,1,"E5A63B","l"]],"x":["01"]},"Text 3":{"c":1,"np":1,"r":[[3400,1,"FFFFFF","l"]]},"Text 5":{"c":1,"np":1,"r":[[1400,0,"EAF7F4","l"]]},"Text 7":{"b":[0.55,7.14,12.15,0.22],"c":1,"np":1,"r":[[780,0,"DDF3EC","l"]],"x":["03"]}},"z":"pssssssss"},"30":{"f":["rect:0E7C7B:13.33x0.22","rect:FFF9EE:13.33x7.50","roundRect:F8FFFC:5.80x0.72","roundRect:FFFFFF:5.80x0.96","roundRect:FFFFFF:5.80x0.96","roundRect:FFFFFF:5.80x0.96"],"no":30,"p":[[0.75,1.45,5.35,4.95,{"l":19602,"r":19602}]],"pic":1,"sp":18,"t":{"Text 11":{"c":1,"np":1,"r":[[1250,1,"063D3B","l"]]},"Text 12":{"c":1,"np":1,"r":[[920,0,"112B2D","l"]]},"Text 14":{"c":1,"np":1,"r":[[1250,1,"063D3B","l"]]},"Text 15":{"c":1,"np":1,"r":[[840,0,"112B2D","l"]]},"Text 17":{"b":[0.55,7.14,12.15,0.22],"c":1,"np":1,"r":[[780,0,"6A7F7D","l"]],"x":["30"]},"Text 2":{"b":[0.65,0.52,9.3,0.5],"c":1,"np":1,"r":[[2600,1,"112B2D","l"]]},"Text 3":{"c":1,"np":1,"r":[[1100,0,"6A7F7D","l"]]},"Text 5":{"c":1,"np":1,"r":[[1250,1,"063D3B","l"]]},"Text 6":{"c":1,"np":1,"r":[[920,0,"112B2D","l"]]},"Text 8":{"c":1,"np":1,"r":[[1250,1,"063D3B","l"]]},"Text 9":{"c":1,"np":1,"r":[[920,0,"112B2D","l"]]}},"z":"sssspssssssssssssss"},"31":{"f":["rect:0E7C7B:13.33x0.22","rect:FFF9EE:13.33x7.50","roundRect:FFF7E8:5.20x0.75","roundRect:FFFFFF:5.20x0.95","roundRect:FFFFFF:5.20x0.95","roundRect:FFFFFF:5.20x0.95"],"no":31,"p":[[0.75,1.45,5.9,4.95,{"l":16477,"r":16477}]],"pic":1,"sp":18,"t":{"Text 11":{"c":1,"np":1,"r":[[1250,1,"063D3B","l"]]},"Text 12":{"c":1,"np":1,"r":[[920,0,"112B2D","l"]]},"Text 14":{"c":1,"np":1,"r":[[1250,1,"063D3B","l"]]},"Text 15":{"c":1,"np":1,"r":[[840,0,"112B2D","l"]]},"Text 17":{"b":[0.55,7.14,12.15,0.22],"c":1,"np":1,"r":[[780,0,"6A7F7D","l"]],"x":["31"]},"Text 2":{"b":[0.65,0.52,9.3,0.5],"c":1,"np":1,"r":[[2600,1,"112B2D","l"]]},"Text 3":{"c":1,"np":1,"r":[[1100,0,"6A7F7D","l"]]},"Text 5":{"c":1,"np":1,"r":[[1250,1,"063D3B","l"]]},"Text 6":{"c":1,"np":1,"r":[[920,0,"112B2D","l"]]},"Text 8":{"c":1,"np":1,"r":[[1250,1,"063D3B","l"]]},"Text 9":{"c":1,"np":1,"r":[[920,0,"112B2D","l"]]}},"z":"sssspssssssssssssss"},"32":{"f":["rect:0E7C7B:13.33x0.22","rect:FFF9EE:13.33x7.50","roundRect:F8FFFC:5.50x0.90","roundRect:FFFFFF:5.50x0.90","roundRect:FFFFFF:5.50x0.90","roundRect:FFFFFF:5.50x0.90"],"no":32,"p":[[7.0,1.45,5.35,4.95,{"l":19602,"r":19602}]],"pic":1,"sp":18,"t":{"Text 11":{"c":1,"np":1,"r":[[1250,1,"063D3B","l"]]},"Text 12":{"c":1,"np":1,"r":[[920,0,"112B2D","l"]]},"Text 14":{"c":1,"np":1,"r":[[1250,1,"063D3B","l"]]},"Text 15":{"c":1,"np":1,"r":[[920,0,"112B2D","l"]]},"Text 17":{"b":[0.55,7.14,12.15,0.22],"c":1,"np":1,"r":[[780,0,"6A7F7D","l"]],"x":["32"]},"Text 2":{"b":[0.65,0.52,9.3,0.5],"c":1,"np":1,"r":[[2600,1,"112B2D","l"]]},"Text 3":{"c":1,"np":1,"r":[[1100,0,"6A7F7D","l"]]},"Text 5":{"c":1,"np":1,"r":[[1250,1,"063D3B","l"]]},"Text 6":{"c":1,"np":1,"r":[[920,0,"112B2D","l"]]},"Text 8":{"c":1,"np":1,"r":[[1250,1,"063D3B","l"]]},"Text 9":{"c":1,"np":1,"r":[[920,0,"112B2D","l"]]}},"z":"sssspssssssssssssss"},"33":{"f":["rect:0E7C7B:13.33x0.22","rect:FFF9EE:13.33x7.50","roundRect:E8F6F2:1.04x2.60","roundRect:E8F6F2:1.04x2.60","roundRect:FFF7E8:5.70x1.25","roundRect:FFFFFF:1.04x2.60","roundRect:FFFFFF:1.04x2.60","roundRect:FFFFFF:1.04x2.60"],"no":33,"p":[[0.7,1.45,5.6,4.9,{"l":17857,"r":17857}]],"pic":1,"sp":24,"t":{"Text 11":{"np":1,"r":[[1000,1,"1F6E86","l"]],"x":["3"]},"Text 12":{"c":1,"np":1,"r":[[840,0,"112B2D","l"]]},"Text 14":{"np":1,"r":[[1000,1,"1F6E86","l"]],"x":["4"]},"Text 15":{"c":1,"np":1,"r":[[840,0,"112B2D","l"]]},"Text 17":{"np":1,"r":[[1000,1,"1F6E86","l"]],"x":["5"]},"Text 18":{"c":1,"np":1,"r":[[840,0,"112B2D","l"]]},"Text 2":{"b":[0.65,0.52,9.3,0.5],"c":1,"np":1,"r":[[2600,1,"112B2D","l"]]},"Text 20":{"c":1,"np":1,"r":[[1300,1,"063D3B","l"]]},"Text 21":{"c":1,"np":1,"r":[[1020,0,"112B2D","l"]]},"Text 23":{"b":[0.55,7.14,12.15,0.22],"c":1,"np":1,"r":[[780,0,"6A7F7D","l"]],"x":["33"]},"Text 3":{"c":1,"np":1,"r":[[1100,0,"6A7F7D","l"]]},"Text 5":{"np":1,"r":[[1000,1,"1F6E86","l"]],"x":["1"]},"Text 6":{"c":1,"np":1,"r":[[840,0,"112B2D","l"]]},"Text 8":{"np":1,"r":[[1000,1,"1F6E86","l"]],"x":["2"]},"Text 9":{"c":1,"np":1,"r":[[840,0,"112B2D","l"]]}},"z":"sssspssssssssssssssssssss"},"34":{"f":["rect:0E7C7B:13.33x0.22","rect:FFF9EE:13.33x7.50","roundRect:E8F6F2:0.99x2.65","roundRect:E8F6F2:0.99x2.65","roundRect:F8FFFC:5.45x1.25","roundRect:FFFFFF:0.99x2.65","roundRect:FFFFFF:0.99x2.65","roundRect:FFFFFF:0.99x2.65"],"no":34,"p":[[6.8,1.4,5.75,4.9,{"l":16996,"r":16996}]],"pic":1,"sp":24,"t":{"Text 11":{"np":1,"r":[[1000,1,"E5A63B","l"]],"x":["3"]},"Text 12":{"c":1,"np":1,"r":[[840,0,"112B2D","l"]]},"Text 14":{"np":1,"r":[[1000,1,"E5A63B","l"]],"x":["4"]},"Text 15":{"c":1,"np":1,"r":[[840,0,"112B2D","l"]]},"Text 17":{"np":1,"r":[[1000,1,"E5A63B","l"]],"x":["5"]},"Text 18":{"c":1,"np":1,"r":[[840,0,"112B2D","l"]]},"Text 2":{"b":[0.65,0.52,9.3,0.5],"c":1,"np":1,"r":[[2600,1,"112B2D","l"]]},"Text 20":{"c":1,"np":1,"r":[[1300,1,"063D3B","l"]]},"Text 21":{"c":1,"np":1,"r":[[1050,0,"112B2D","l"]]},"Text 23":{"b":[0.55,7.14,12.15,0.22],"c":1,"np":1,"r":[[780,0,"6A7F7D","l"]],"x":["34"]},"Text 3":{"c":1,"np":1,"r":[[1100,0,"6A7F7D","l"]]},"Text 5":{"np":1,"r":[[1000,1,"E5A63B","l"]],"x":["1"]},"Text 6":{"c":1,"np":1,"r":[[840,0,"112B2D","l"]]},"Text 8":{"np":1,"r":[[1000,1,"E5A63B","l"]],"x":["2"]},"Text 9":{"c":1,"np":1,"r":[[840,0,"112B2D","l"]]}},"z":"sssspssssssssssssssssssss"},"35":{"f":["rect:0E7C7B:13.33x0.22","rect:FFF9EE:13.33x7.50","roundRect:FFF7E8:5.25x0.88","roundRect:FFFFFF:5.25x0.88","roundRect:FFFFFF:5.25x0.88","roundRect:FFFFFF:5.25x0.88"],"no":35,"p":[[0.75,1.45,5.85,4.9,{"l":16422,"r":16422}]],"pic":1,"sp":18,"t":{"Text 11":{"c":1,"np":1,"r":[[1250,1,"063D3B","l"]]},"Text 12":{"c":1,"np":1,"r":[[920,0,"112B2D","l"]]},"Text 14":{"c":1,"np":1,"r":[[1250,1,"063D3B","l"]]},"Text 15":{"c":1,"np":1,"r":[[920,0,"112B2D","l"]]},"Text 17":{"b":[0.55,7.14,12.15,0.22],"c":1,"np":1,"r":[[780,0,"6A7F7D","l"]],"x":["35"]},"Text 2":{"b":[0.65,0.52,9.3,0.5],"c":1,"np":1,"r":[[2600,1,"112B2D","l"]]},"Text 3":{"c":1,"np":1,"r":[[1100,0,"6A7F7D","l"]]},"Text 5":{"c":1,"np":1,"r":[[1250,1,"063D3B","l"]]},"Text 6":{"c":1,"np":1,"r":[[920,0,"112B2D","l"]]},"Text 8":{"c":1,"np":1,"r":[[1250,1,"063D3B","l"]]},"Text 9":{"c":1,"np":1,"r":[[920,0,"112B2D","l"]]}},"z":"sssspssssssssssssss"},"36":{"f":["rect:0E7C7B:13.33x0.22","rect:FFF9EE:13.33x7.50","roundRect:F4FBFF:2.60x0.82","roundRect:F4FBFF:2.60x0.82","roundRect:F8FFFC:2.60x0.82","roundRect:F8FFFC:2.60x0.82","roundRect:FFF7E8:4.30x0.82","roundRect:FFF7E8:4.30x0.82","roundRect:FFF7E8:4.30x0.82","roundRect:FFF7E8:4.30x0.82","roundRect:FFFFFF:4.20x0.82","roundRect:FFFFFF:4.20x0.82","roundRect:FFFFFF:4.20x0.82","roundRect:FFFFFF:4.20x0.82"],"no":36,"p":[],"pic":0,"sp":43,"t":{"Text 11":{"c":1,"np":1,"r":[[1050,1,"063D3B","l"]]},"Text 12":{"c":1,"np":1,"r":[[880,0,"112B2D","l"]]},"Text 14":{"c":1,"np":1,"r":[[1250,1,"063D3B","l"]]},"Text 17":{"c":1,"np":1,"r":[[1050,1,"063D3B","l"]]},"Text 18":{"c":1,"np":1,"r":[[880,0,"112B2D","l"]]},"Text 2":{"b":[0.65,0.52,9.3,0.5],"c":1,"np":1,"r":[[2600,1,"112B2D","l"]]},"Text 20":{"c":1,"np":1,"r":[[1050,1,"063D3B","l"]]},"Text 21":{"c":1,"np":1,"r":[[880,0,"112B2D","l"]]},"Text 23":{"c":1,"np":1,"r":[[1250,1,"063D3B","l"]]},"Text 26":{"c":1,"np":1,"r":[[1050,1,"063D3B","l"]]},"Text 27":{"c":1,"np":1,"r":[[880,0,"112B2D","l"]]},"Text 29":{"c":1,"np":1,"r":[[1050,1,"063D3B","l"]]},"Text 3":{"c":1,"np":1,"r":[[1100,0,"6A7F7D","l"]]},"Text 30":{"c":1,"np":1,"r":[[880,0,"112B2D","l"]]},"Text 32":{"c":1,"np":1,"r":[[1250,1,"063D3B","l"]]},"Text 35":{"c":1,"np":1,"r":[[1050,1,"063D3B","l"]]},"Text 36":{"c":1,"np":1,"r":[[880,0,"112B2D","l"]]},"Text 38":{"c":1,"np":1,"r":[[1050,1,"063D3B","l"]]},"Text 39":{"c":1,"np":1,"r":[[880,0,"112B2D","l"]]},"Text 40":{"c":1,"np":1,"r":[[1250,1,"0E7C7B","ctr"]]},"Text 42":{"b":[0.55,7.14,12.15,0.22],"c":1,"np":1,"r":[[780,0,"6A7F7D","l"]],"x":["36"]},"Text 5":{"c":1,"np":1,"r":[[1250,1,"063D3B","l"]]},"Text 8":{"c":1,"np":1,"r":[[1050,1,"063D3B","l"]]},"Text 9":{"c":1,"np":1,"r":[[880,0,"112B2D","l"]]}},"z":"sssssssssssssssssssssssssssssssssssssssssss"},"37":{"f":["rect:0E7C7B:13.33x0.22","rect:FFF9EE:13.33x7.50","roundRect:E8F6F2:1.06x2.40","roundRect:E8F6F2:1.06x2.40","roundRect:F8FFFC:5.80x1.35","roundRect:FFFFFF:1.06x2.40","roundRect:FFFFFF:1.06x2.40","roundRect:FFFFFF:1.06x2.40"],"no":37,"p":[[7.05,1.45,5.2,4.9,{"l":20153,"r":20153}]],"pic":1,"sp":24,"t":{"Text 11":{"np":1,"r":[[1000,1,"0E7C7B","l"]],"x":["3"]},"Text 12":{"c":1,"np":1,"r":[[840,0,"112B2D","l"]]},"Text 14":{"np":1,"r":[[1000,1,"0E7C7B","l"]],"x":["4"]},"Text 15":{"c":1,"np":1,"r":[[840,0,"112B2D","l"]]},"Text 17":{"np":1,"r":[[1000,1,"0E7C7B","l"]],"x":["5"]},"Text 18":{"c":1,"np":1,"r":[[840,0,"112B2D","l"]]},"Text 2":{"b":[0.65,0.52,9.3,0.5],"c":1,"np":1,"r":[[2600,1,"112B2D","l"]]},"Text 20":{"c":1,"np":1,"r":[[1300,1,"063D3B","l"]]},"Text 21":{"c":1,"np":1,"r":[[1020,0,"112B2D","l"]]},"Text 23":{"b":[0.55,7.14,12.15,0.22],"c":1,"np":1,"r":[[780,0,"6A7F7D","l"]],"x":["37"]},"Text 3":{"c":1,"np":1,"r":[[1100,0,"6A7F7D","l"]]},"Text 5":{"np":1,"r":[[1000,1,"0E7C7B","l"]],"x":["1"]},"Text 6":{"c":1,"np":1,"r":[[840,0,"112B2D","l"]]},"Text 8":{"np":1,"r":[[1000,1,"0E7C7B","l"]],"x":["2"]},"Text 9":{"c":1,"np":1,"r":[[840,0,"112B2D","l"]]}},"z":"sssspssssssssssssssssssss"},"38":{"f":["rect:0E7C7B:13.33x0.22","rect:FFF9EE:13.33x7.50","roundRect:FFF7E8:5.65x0.90","roundRect:FFFFFF:5.65x0.90","roundRect:FFFFFF:5.65x0.90","roundRect:FFFFFF:5.65x0.90"],"no":38,"p":[[0.72,1.45,5.5,4.95,{"l":18750,"r":18750}]],"pic":1,"sp":18,"t":{"Text 11":{"c":1,"np":1,"r":[[1250,1,"063D3B","l"]]},"Text 12":{"c":1,"np":1,"r":[[920,0,"112B2D","l"]]},"Text 14":{"c":1,"np":1,"r":[[1250,1,"063D3B","l"]]},"Text 15":{"c":1,"np":1,"r":[[920,0,"112B2D","l"]]},"Text 17":{"b":[0.55,7.14,12.15,0.22],"c":1,"np":1,"r":[[780,0,"6A7F7D","l"]],"x":["38"]},"Text 2":{"b":[0.65,0.52,9.3,0.5],"c":1,"np":1,"r":[[2600,1,"112B2D","l"]]},"Text 3":{"c":1,"np":1,"r":[[1100,0,"6A7F7D","l"]]},"Text 5":{"c":1,"np":1,"r":[[1250,1,"063D3B","l"]]},"Text 6":{"c":1,"np":1,"r":[[920,0,"112B2D","l"]]},"Text 8":{"c":1,"np":1,"r":[[1250,1,"063D3B","l"]]},"Text 9":{"c":1,"np":1,"r":[[920,0,"112B2D","l"]]}},"z":"sssspssssssssssssss"},"39":{"f":["rect:0E7C7B:13.33x0.22","rect:FFF9EE:13.33x7.50","roundRect:F8FFFC:5.55x0.90","roundRect:FFFFFF:5.55x0.90","roundRect:FFFFFF:5.55x0.90","roundRect:FFFFFF:5.55x0.90"],"no":39,"p":[[6.8,1.45,5.55,4.95,{"l":18466,"r":18466}]],"pic":1,"sp":18,"t":{"Text 11":{"c":1,"np":1,"r":[[1250,1,"063D3B","l"]]},"Text 12":{"c":1,"np":1,"r":[[920,0,"112B2D","l"]]},"Text 14":{"c":1,"np":1,"r":[[1250,1,"063D3B","l"]]},"Text 15":{"c":1,"np":1,"r":[[920,0,"112B2D","l"]]},"Text 17":{"b":[0.55,7.14,12.15,0.22],"c":1,"np":1,"r":[[780,0,"6A7F7D","l"]],"x":["39"]},"Text 2":{"b":[0.65,0.52,9.3,0.5],"c":1,"np":1,"r":[[2600,1,"112B2D","l"]]},"Text 3":{"c":1,"np":1,"r":[[1100,0,"6A7F7D","l"]]},"Text 5":{"c":1,"np":1,"r":[[1250,1,"063D3B","l"]]},"Text 6":{"c":1,"np":1,"r":[[920,0,"112B2D","l"]]},"Text 8":{"c":1,"np":1,"r":[[1250,1,"063D3B","l"]]},"Text 9":{"c":1,"np":1,"r":[[920,0,"112B2D","l"]]}},"z":"sssspssssssssssssss"},"4":{"f":["rect:0E7C7B:13.33x0.22","rect:FFF9EE:13.33x7.50","roundRect:FFFFFF:5.35x0.72","roundRect:FFFFFF:5.35x1.05","roundRect:FFFFFF:5.35x1.05","roundRect:FFFFFF:5.35x1.05"],"no":4,"p":[[6.65,1.45,6.0,4.95,{"l":15909,"r":15909}]],"pic":1,"sp":18,"t":{"Text 11":{"c":1,"np":1,"r":[[1250,1,"063D3B","l"]]},"Text 12":{"c":1,"np":1,"r":[[920,0,"112B2D","l"]]},"Text 14":{"c":1,"np":1,"r":[[1100,1,"063D3B","l"]]},"Text 15":{"c":1,"np":1,"r":[[900,0,"112B2D","l"]]},"Text 17":{"b":[0.55,7.14,12.15,0.22],"c":1,"np":1,"r":[[780,0,"6A7F7D","l"]],"x":["04"]},"Text 2":{"b":[0.65,0.52,9.3,0.5],"c":1,"np":1,"r":[[2600,1,"112B2D","l"]]},"Text 3":{"c":1,"np":1,"r":[[1100,0,"6A7F7D","l"]]},"Text 5":{"c":1,"np":1,"r":[[1250,1,"063D3B","l"]]},"Text 6":{"c":1,"np":1,"r":[[920,0,"112B2D","l"]]},"Text 8":{"c":1,"np":1,"r":[[1250,1,"063D3B","l"]]},"Text 9":{"c":1,"np":1,"r":[[920,0,"112B2D","l"]]}},"z":"sssspssssssssssssss"},"40":{"f":["rect:0E7C7B:13.33x0.22","rect:FFF9EE:13.33x7.50","roundRect:FFF7E8:5.55x0.90","roundRect:FFFFFF:5.55x0.90","roundRect:FFFFFF:5.55x0.90","roundRect:FFFFFF:5.55x0.90"],"no":40,"p":[[0.75,1.45,5.55,4.95,{"l":18466,"r":18466}]],"pic":1,"sp":18,"t":{"Text 11":{"c":1,"np":1,"r":[[1250,1,"063D3B","l"]]},"Text 12":{"c":1,"np":1,"r":[[920,0,"112B2D","l"]]},"Text 14":{"c":1,"np":1,"r":[[1250,1,"063D3B","l"]]},"Text 15":{"c":1,"np":1,"r":[[920,0,"112B2D","l"]]},"Text 17":{"b":[0.55,7.14,12.15,0.22],"c":1,"np":1,"r":[[780,0,"6A7F7D","l"]],"x":["40"]},"Text 2":{"b":[0.65,0.52,9.3,0.5],"c":1,"np":1,"r":[[2600,1,"112B2D","l"]]},"Text 3":{"c":1,"np":1,"r":[[1100,0,"6A7F7D","l"]]},"Text 5":{"c":1,"np":1,"r":[[1250,1,"063D3B","l"]]},"Text 6":{"c":1,"np":1,"r":[[920,0,"112B2D","l"]]},"Text 8":{"c":1,"np":1,"r":[[1250,1,"063D3B","l"]]},"Text 9":{"c":1,"np":1,"r":[[920,0,"112B2D","l"]]}},"z":"sssspssssssssssssss"},"41":{"f":["rect:0E7C7B:13.33x0.22","rect:FFF9EE:13.33x7.50","roundRect:F8FFFC:5.55x0.70","roundRect:FFFFFF:5.55x1.00","roundRect:FFFFFF:5.55x1.00","roundRect:FFFFFF:5.55x1.00"],"no":41,"p":[[6.9,1.45,5.4,4.95,{"l":19318,"r":19318}]],"pic":1,"sp":18,"t":{"Text 11":{"c":1,"np":1,"r":[[1250,1,"063D3B","l"]]},"Text 12":{"c":1,"np":1,"r":[[920,0,"112B2D","l"]]},"Text 14":{"c":1,"np":1,"r":[[1250,1,"063D3B","l"]]},"Text 15":{"c":1,"np":1,"r":[[900,0,"112B2D","l"]]},"Text 17":{"b":[0.55,7.14,12.15,0.22],"c":1,"np":1,"r":[[780,0,"6A7F7D","l"]],"x":["41"]},"Text 2":{"b":[0.65,0.52,9.3,0.5],"c":1,"np":1,"r":[[2600,1,"112B2D","l"]]},"Text 3":{"c":1,"np":1,"r":[[1100,0,"6A7F7D","l"]]},"Text 5":{"c":1,"np":1,"r":[[1250,1,"063D3B","l"]]},"Text 6":{"c":1,"np":1,"r":[[920,0,"112B2D","l"]]},"Text 8":{"c":1,"np":1,"r":[[1250,1,"063D3B","l"]]},"Text 9":{"c":1,"np":1,"r":[[920,0,"112B2D","l"]]}},"z":"sssspssssssssssssss"},"42":{"f":["rect:000000/62000:13.33x7.50","rect:063D3B/78000:13.33x7.50"],"no":42,"p":[[0.0,0.0,13.333,7.5,{"l":1,"r":1}]],"pic":1,"sp":8,"t":{"Text 2":{"np":1,"r":[[4800,1,"E5A63B","l"]],"x":["07"]},"Text 3":{"c":1,"np":1,"r":[[3400,1,"FFFFFF","l"]]},"Text 5":{"c":1,"np":1,"r":[[1400,0,"EAF7F4","l"]]},"Text 7":{"b":[0.55,7.14,12.15,0.22],"c":1,"np":1,"r":[[780,0,"DDF3EC","l"]],"x":["42"]}},"z":"pssssssss"},"43":{"f":["rect:0E7C7B:13.33x0.22","rect:FFF9EE:13.33x7.50","roundRect:FFF7E8:5.20x0.65","roundRect:FFFFFF:5.20x0.95","roundRect:FFFFFF:5.20x0.95","roundRect:FFFFFF:5.20x0.95"],"no":43,"p":[[0.75,1.45,5.9,4.95,{"l":16477,"r":16477}]],"pic":1,"sp":18,"t":{"Text 11":{"c":1,"np":1,"r":[[1250,1,"063D3B","l"]]},"Text 12":{"c":1,"np":1,"r":[[920,0,"112B2D","l"]]},"Text 14":{"c":1,"np":1,"r":[[1250,1,"063D3B","l"]]},"Text 15":{"c":1,"np":1,"r":[[820,0,"112B2D","l"]]},"Text 17":{"b":[0.55,7.14,12.15,0.22],"c":1,"np":1,"r":[[780,0,"6A7F7D","l"]],"x":["43"]},"Text 2":{"b":[0.65,0.52,9.3,0.5],"c":1,"np":1,"r":[[2600,1,"112B2D","l"]]},"Text 3":{"c":1,"np":1,"r":[[1100,0,"6A7F7D","l"]]},"Text 5":{"c":1,"np":1,"r":[[1250,1,"063D3B","l"]]},"Text 6":{"c":1,"np":1,"r":[[920,0,"112B2D","l"]]},"Text 8":{"c":1,"np":1,"r":[[1250,1,"063D3B","l"]]},"Text 9":{"c":1,"np":1,"r":[[920,0,"112B2D","l"]]}},"z":"sssspssssssssssssss"},"44":{"f":["rect:0E7C7B:13.33x0.22","rect:FFF9EE:13.33x7.50","roundRect:E8F6F2:2.22x1.00","roundRect:E8F6F2:2.22x1.00","roundRect:FFF7E8:9.10x0.80","roundRect:FFFFFF:2.22x1.00","roundRect:FFFFFF:2.22x1.00","roundRect:FFFFFF:2.22x1.00","roundRect:FFFFFF:3.55x1.35","roundRect:FFFFFF:3.55x1.35","roundRect:FFFFFF:3.55x1.35"],"no":44,"p":[],"pic":0,"sp":33,"t":{"Text 11":{"np":1,"r":[[1000,1,"0E7C7B","l"]],"x":["3"]},"Text 12":{"c":1,"np":1,"r":[[1000,0,"112B2D","l"]]},"Text 14":{"np":1,"r":[[1000,1,"0E7C7B","l"]],"x":["4"]},"Text 15":{"c":1,"np":1,"r":[[1000,0,"112B2D","l"]]},"Text 17":{"np":1,"r":[[1000,1,"0E7C7B","l"]],"x":["5"]},"Text 18":{"c":1,"np":1,"r":[[1000,0,"112B2D","l"]]},"Text 2":{"b":[0.65,0.52,9.3,0.5],"c":1,"np":1,"r":[[2600,1,"112B2D","l"]]},"Text 20":{"c":1,"np":1,"r":[[1250,1,"063D3B","l"]]},"Text 21":{"c":1,"np":1,"r":[[920,0,"112B2D","l"]]},"Text 23":{"c":1,"np":1,"r":[[1250,1,"063D3B","l"]]},"Text 24":{"c":1,"np":1,"r":[[920,0,"112B2D","l"]]},"Text 26":{"c":1,"np":1,"r":[[1250,1,"063D3B","l"]]},"Text 27":{"c":1,"np":1,"r":[[920,0,"112B2D","l"]]},"Text 29":{"c":1,"np":1,"r":[[1300,1,"063D3B","l"]]},"Text 3":{"c":1,"np":1,"r":[[1100,0,"6A7F7D","l"]]},"Text 30":{"c":1,"np":1,"r":[[920,0,"112B2D","l"]]},"Text 32":{"b":[0.55,7.14,12.15,0.22],"c":1,"np":1,"r":[[780,0,"6A7F7D","l"]],"x":["44"]},"Text 5":{"np":1,"r":[[1000,1,"0E7C7B","l"]],"x":["1"]},"Text 6":{"c":1,"np":1,"r":[[1000,0,"112B2D","l"]]},"Text 8":{"np":1,"r":[[1000,1,"0E7C7B","l"]],"x":["2"]},"Text 9":{"c":1,"np":1,"r":[[1000,0,"112B2D","l"]]}},"z":"sssssssssssssssssssssssssssssssss"},"45":{"f":["rect:000000/80000:11.90x2.15","rect:0E7C7B:13.33x0.22","rect:FFF9EE:13.33x7.50","roundRect:FFFFFF:3.70x1.25","roundRect:FFFFFF:3.70x1.25","roundRect:FFFFFF:3.70x1.25"],"no":45,"p":[[0.75,3.25,11.9,2.15,{"b":33940,"t":33940}]],"pic":1,"sp":17,"t":{"Text 11":{"c":1,"np":1,"r":[[1250,1,"063D3B","l"]]},"Text 12":{"c":1,"np":1,"r":[[920,0,"112B2D","l"]]},"Text 14":{"c":1,"np":1,"r":[[2300,1,"FFFFFF","ctr"]]},"Text 16":{"b":[0.55,7.14,12.15,0.22],"c":1,"np":1,"r":[[780,0,"6A7F7D","l"]],"x":["45"]},"Text 2":{"b":[0.65,0.52,9.3,0.5],"c":1,"np":1,"r":[[2600,1,"112B2D","l"]]},"Text 3":{"c":1,"np":1,"r":[[1100,0,"6A7F7D","l"]]},"Text 5":{"c":1,"np":1,"r":[[1250,1,"063D3B","l"]]},"Text 6":{"c":1,"np":1,"r":[[920,0,"112B2D","l"]]},"Text 8":{"c":1,"np":1,"r":[[1250,1,"063D3B","l"]]},"Text 9":{"c":1,"np":1,"r":[[920,0,"112B2D","l"]]}},"z":"ssssssssssssspssss"},"46":{"f":["rect:000000/67000:13.33x7.50","rect:063D3B/75000:13.33x7.50","rect:063D3B:13.33x7.50","rect:25A7A1:13.33x0.22"],"no":46,"p":[[0.0,0.0,13.333,7.5,{"l":1,"r":1}]],"pic":1,"sp":10,"t":{"Text 2":{"b":[0.65,0.52,9.3,0.5],"c":1,"np":1,"r":[[2600,1,"FFFFFF","l"]]},"Text 3":{"c":1,"np":1,"r":[[1100,0,"CFEDE7","l"]]},"Text 6":{"np":1,"r":[[1700,1,"DDF3EC","ctr"]]},"Text 7":{"c":1,"e":1,"np":1,"r":[[1200,0,"EAF7F4","ctr"]]},"Text 9":{"b":[0.55,7.14,12.15,0.22],"c":1,"np":1,"r":[[780,0,"DDF3EC","l"]],"x":["46"]}},"z":"sssspssssss"},"5":{"f":["rect:0E7C7B:13.33x0.22","rect:FFF9EE:13.33x7.50","roundRect:F4FBFF:3.55x2.85","roundRect:F8FFFC:3.55x2.85","roundRect:FFF7E8:3.55x2.85"],"no":5,"p":[],"pic":0,"sp":17,"t":{"Text 11":{"c":1,"np":1,"r":[[1250,1,"063D3B","l"]]},"Text 12":{"c":1,"np":1,"r":[[920,0,"112B2D","l"]]},"Text 14":{"c":1,"np":1,"r":[[2000,1,"063D3B","ctr"]]},"Text 16":{"b":[0.55,7.14,12.15,0.22],"c":1,"np":1,"r":[[780,0,"6A7F7D","l"]],"x":["05"]},"Text 2":{"b":[0.65,0.52,9.3,0.5],"c":1,"np":1,"r":[[2600,1,"112B2D","l"]]},"Text 3":{"c":1,"np":1,"r":[[1100,0,"6A7F7D","l"]]},"Text 5":{"c":1,"np":1,"r":[[1250,1,"063D3B","l"]]},"Text 6":{"c":1,"np":1,"r":[[920,0,"112B2D","l"]]},"Text 8":{"c":1,"np":1,"r":[[1250,1,"063D3B","l"]]},"Text 9":{"c":1,"e":1,"np":1,"r":[[920,0,"112B2D","l"]]}},"z":"sssssssssssssssss"},"6":{"f":["rect:000000/62000:13.33x7.50","rect:063D3B/78000:13.33x7.50"],"no":6,"p":[[0.0,0.0,13.333,7.5,{"l":1,"r":1}]],"pic":1,"sp":8,"t":{"Text 2":{"np":1,"r":[[4800,1,"E5A63B","l"]],"x":["02"]},"Text 3":{"c":1,"np":1,"r":[[3400,1,"FFFFFF","l"]]},"Text 5":{"c":1,"np":1,"r":[[1400,0,"EAF7F4","l"]]},"Text 7":{"b":[0.55,7.14,12.15,0.22],"c":1,"np":1,"r":[[780,0,"DDF3EC","l"]],"x":["06"]}},"z":"pssssssss"},"7":{"f":["rect:0E7C7B:13.33x0.22","rect:FFF9EE:13.33x7.50","roundRect:F8FFFC:2.75x1.05","roundRect:F8FFFC:2.75x1.05","roundRect:FFFFFF:3.55x1.05","roundRect:FFFFFF:3.55x1.05","roundRect:FFFFFF:3.55x1.05"],"no":7,"p":[[0.75,1.55,4.25,2.45,{"l":1212,"r":1212}],[8.35,1.55,4.25,2.45,{"l":1212,"r":1212}]],"pic":2,"sp":21,"t":{"Text 11":{"c":1,"np":1,"r":[[1250,1,"063D3B","l"]]},"Text 12":{"c":1,"np":1,"r":[[920,0,"112B2D","l"]]},"Text 14":{"c":1,"np":1,"r":[[1250,1,"063D3B","l"]]},"Text 15":{"c":1,"np":1,"r":[[920,0,"112B2D","l"]]},"Text 17":{"c":1,"np":1,"r":[[1250,1,"063D3B","l"]]},"Text 18":{"c":1,"np":1,"r":[[920,0,"112B2D","l"]]},"Text 2":{"b":[0.65,0.52,9.3,0.5],"c":1,"np":1,"r":[[2600,1,"112B2D","l"]]},"Text 20":{"b":[0.55,7.14,12.15,0.22],"c":1,"np":1,"r":[[780,0,"6A7F7D","l"]],"x":["07"]},"Text 3":{"c":1,"np":1,"r":[[1100,0,"6A7F7D","l"]]},"Text 5":{"c":1,"np":1,"r":[[1250,1,"063D3B","l"]]},"Text 6":{"c":1,"np":1,"r":[[920,0,"112B2D","l"]]},"Text 8":{"c":1,"np":1,"r":[[1250,1,"063D3B","l"]]},"Text 9":{"c":1,"np":1,"r":[[920,0,"112B2D","l"]]}},"z":"ssssppsssssssssssssssss"},"8":{"f":["rect:0E7C7B:13.33x0.22","rect:FFF9EE:13.33x7.50","roundRect:E8F6F2:2.24x0.92","roundRect:E8F6F2:2.24x0.92","roundRect:FFF7E8:5.70x1.75","roundRect:FFFFFF:2.24x0.92","roundRect:FFFFFF:2.24x0.92","roundRect:FFFFFF:2.24x0.92","roundRect:FFFFFF:3.70x1.28","roundRect:FFFFFF:3.70x1.28","roundRect:FFFFFF:3.70x1.28"],"no":8,"p":[[0.85,4.72,5.5,1.75,{"b":21717,"t":21717}]],"pic":1,"sp":33,"t":{"Text 11":{"np":1,"r":[[1000,1,"0E7C7B","l"]],"x":["3"]},"Text 12":{"c":1,"np":1,"r":[[840,0,"112B2D","l"]]},"Text 14":{"np":1,"r":[[1000,1,"0E7C7B","l"]],"x":["4"]},"Text 15":{"c":1,"np":1,"r":[[840,0,"112B2D","l"]]},"Text 17":{"np":1,"r":[[1000,1,"0E7C7B","l"]],"x":["5"]},"Text 18":{"c":1,"np":1,"r":[[840,0,"112B2D","l"]]},"Text 2":{"b":[0.65,0.52,9.3,0.5],"c":1,"np":1,"r":[[2600,1,"112B2D","l"]]},"Text 20":{"c":1,"np":1,"r":[[1250,1,"063D3B","l"]]},"Text 21":{"c":1,"np":1,"r":[[920,0,"112B2D","l"]]},"Text 23":{"c":1,"np":1,"r":[[1250,1,"063D3B","l"]]},"Text 24":{"c":1,"np":1,"r":[[920,0,"112B2D","l"]]},"Text 26":{"c":1,"np":1,"r":[[1250,1,"063D3B","l"]]},"Text 27":{"c":1,"np":1,"r":[[920,0,"112B2D","l"]]},"Text 29":{"c":1,"np":1,"r":[[1400,1,"063D3B","l"]]},"Text 3":{"c":1,"np":1,"r":[[1100,0,"6A7F7D","l"]]},"Text 30":{"c":1,"np":1,"r":[[1050,0,"112B2D","l"]]},"Text 32":{"b":[0.55,7.14,12.15,0.22],"c":1,"np":1,"r":[[780,0,"6A7F7D","l"]],"x":["08"]},"Text 5":{"np":1,"r":[[1000,1,"0E7C7B","l"]],"x":["1"]},"Text 6":{"c":1,"np":1,"r":[[840,0,"112B2D","l"]]},"Text 8":{"np":1,"r":[[1000,1,"0E7C7B","l"]],"x":["2"]},"Text 9":{"c":1,"np":1,"r":[[840,0,"112B2D","l"]]}},"z":"sssssssssssssssssssssssssssspsssss"},"9":{"f":["rect:0E7C7B:13.33x0.22","rect:FFF9EE:13.33x7.50","roundRect:E8F6F2:2.28x1.05","roundRect:E8F6F2:2.28x1.05","roundRect:FFFFFF:2.28x1.05","roundRect:FFFFFF:2.28x1.05","roundRect:FFFFFF:2.28x1.05","roundRect:FFFFFF:2.70x1.15","roundRect:FFFFFF:2.70x1.15","roundRect:FFFFFF:2.70x1.15","roundRect:FFFFFF:2.70x1.15"],"no":9,"p":[[0.8,3.0,5.7,2.55,{"b":10234,"t":10234}]],"pic":1,"sp":33,"t":{"Text 11":{"np":1,"r":[[1000,1,"1F6E86","l"]],"x":["3"]},"Text 12":{"c":1,"np":1,"r":[[840,0,"112B2D","l"]]},"Text 14":{"np":1,"r":[[1000,1,"1F6E86","l"]],"x":["4"]},"Text 15":{"c":1,"np":1,"r":[[840,0,"112B2D","l"]]},"Text 17":{"np":1,"r":[[1000,1,"1F6E86","l"]],"x":["5"]},"Text 18":{"c":1,"np":1,"r":[[840,0,"112B2D","l"]]},"Text 2":{"b":[0.65,0.52,9.3,0.5],"c":1,"np":1,"r":[[2600,1,"112B2D","l"]]},"Text 20":{"c":1,"np":1,"r":[[1250,1,"063D3B","l"]]},"Text 21":{"c":1,"np":1,"r":[[920,0,"112B2D","l"]]},"Text 23":{"c":1,"np":1,"r":[[1250,1,"063D3B","l"]]},"Text 24":{"c":1,"np":1,"r":[[920,0,"112B2D","l"]]},"Text 26":{"c":1,"np":1,"r":[[1250,1,"063D3B","l"]]},"Text 27":{"c":1,"np":1,"r":[[920,0,"112B2D","l"]]},"Text 29":{"c":1,"np":1,"r":[[1250,1,"063D3B","l"]]},"Text 3":{"c":1,"np":1,"r":[[1100,0,"6A7F7D","l"]]},"Text 30":{"c":1,"np":1,"r":[[920,0,"112B2D","l"]]},"Text 32":{"b":[0.55,7.14,12.15,0.22],"c":1,"np":1,"r":[[780,0,"6A7F7D","l"]],"x":["09"]},"Text 5":{"np":1,"r":[[1000,1,"1F6E86","l"]],"x":["1"]},"Text 6":{"c":1,"np":1,"r":[[840,0,"112B2D","l"]]},"Text 8":{"np":1,"r":[[1000,1,"1F6E86","l"]],"x":["2"]},"Text 9":{"c":1,"np":1,"r":[[840,0,"112B2D","l"]]}},"z":"ssssssssssssssssssspssssssssssssss"}}}"""


if __name__ == "__main__":
    if len(sys.argv) >= 2:
        _dir = sys.argv[1]
    else:
        _dir = os.path.dirname(os.path.abspath(__file__))
    _out = evaluate(_dir)
    print(json.dumps(_out, ensure_ascii=False, indent=2))
