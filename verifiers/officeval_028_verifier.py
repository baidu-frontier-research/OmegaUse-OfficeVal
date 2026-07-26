# -*- coding: utf-8 -*-
"""
职称评审表_表格修正版.docx 自动评估脚本
==========================================
按照"打分细则"对目标 .docx 文件进行无人工干预的自动评分。

评分模型：
  维度1（可用与可修改性）：不满足任何一条 -> 直接 0 分，不再检查维度2。
  维度2（完成度）：得分点（每一条须全部子点满足才加分）+ 扣分点（任一子点命中即扣分），累加得到最终分。

依赖：仅使用 Python 标准库（zipfile + re + xml）。
用法：python officeval_028_verifier.py [目录路径]
"""

import sys
import os
import re
import json
import glob
import zipfile

# ---------------------------------------------------------------------------
# Word XML 命名空间常量
# ---------------------------------------------------------------------------
DXA_PER_CM = 567.0          # 1cm = 567 twips(dxa)
SZ_PER_PT = 8.0             # w:sz 单位为 1/8 磅，故 8 = 1磅
SCRIPT_ID = "028"


def load_part(zf, name):
    """读取 docx 内某个 xml 部件文本（utf-8）。不存在返回 None。"""
    try:
        with zf.open(name) as fp:
            return fp.read().decode("utf-8", errors="replace")
    except KeyError:
        return None


def split_tables(xml):
    """返回文档中所有 <w:tbl>...</w:tbl> 字符串（顺序即文档顺序）。"""
    return re.findall(r"<w:tbl>.*?</w:tbl>", xml, re.S)


def split_rows(tbl_xml):
    return re.findall(r"<w:tr\b.*?</w:tr>", tbl_xml, re.S)


def split_cells(row_xml):
    return re.findall(r"<w:tc>.*?</w:tc>", row_xml, re.S)


def cell_text(cell_xml):
    """提取单元格内所有 <w:t> 文本并拼接。"""
    parts = re.findall(r"<w:t[ >].*?</w:t>", cell_xml, re.S)
    out = []
    for p in parts:
        m = re.search(r">(.*?)</w:t>", p, re.S)
        if m:
            out.append(m.group(1))
    return "".join(out)


# ===========================================================================
# 维度 1：可用与可修改性
#   任一子项不满足 -> 维度1 整体不通过 -> 总分 0
# ===========================================================================
def check_dimension1(path, zf, doc_xml, pages):
    """返回 (passed: bool, details: list[(ok, msg)])。

    按用户口径，维度一仅保留"文件格式正确且能正常打开"一项：
      · 原"20个表格均为原生可编辑表格/不能截图覆盖" —— 已删除；
      · 原"保持25页、内容与章节顺序完整" —— 已删除；
      · 原"无表格断裂/整页错位/文字重叠/裁切/新增空白页" —— 已删除；
      · 原"格式错误/无法打开/不可编辑/结构损坏则维度1=0"门槛 —— 已删除。
    """
    del zf, pages  # 保留形参以维持调用点签名不变
    details = []

    # 文件为 .docx，且能正常被解析（能 unzip + 含 document.xml 即视为可打开）
    ext = os.path.splitext(path)[1].lower()
    fmt_ok = ext == ".docx" and doc_xml is not None
    details.append((fmt_ok, f"文件格式为 {ext}，且可被正常解析打开" if fmt_ok
                    else f"文件格式 {ext} 非法或文档无法打开/解析（仅支持 .docx）"))

    passed = all(ok for ok, _ in details)
    return passed, details


# ===========================================================================
# 维度 2 - 得分点
# ===========================================================================
def _parse_one_border(attrs: str) -> "dict[str, object]":
    """从边线元素属性串解析 {val,color,sz}。w:sz 单位为 1/8 磅。"""
    val = re.search(r'w:val="([^"]+)"', attrs)
    col = re.search(r'w:color="([^"]+)"', attrs)
    sz = re.search(r'w:sz="(\d+)"', attrs)
    return {
        "val": val.group(1) if val else None,
        "color": col.group(1) if col else None,
        "sz": int(sz.group(1)) if sz else None,
    }


def _parse_borders_block(block: str) -> "dict[str, dict[str, object]]":
    """解析一个 <w:tblBorders>/<w:tcBorders> 块，返回 {side: border}。
    start/end 归一化为 left/right。"""
    d: "dict[str, dict[str, object]]" = {}
    if not block:
        return d
    for m in re.finditer(
            r'<w:(top|bottom|left|right|start|end|insideH|insideV)\b([^>]*)>', block):
        side = {"start": "left", "end": "right"}.get(m.group(1), m.group(1))
        d[side] = _parse_one_border(m.group(2))
    return d


_NONE_VALS = {None, "none", "nil"}


def _border_present(b) -> bool:
    """边线是否"存在且连续显示"：有定义且 w:val 非 none/nil。"""
    return b is not None and b.get("val") not in _NONE_VALS


def _tbl_border_block(t: str) -> "dict[str, dict[str, object]]":
    m = re.search(r"<w:tblBorders\b.*?</w:tblBorders>", t, re.S)
    return _parse_borders_block(m.group(0) if m else "")


def _cell_border_block(cell_xml: str) -> "dict[str, dict[str, object]]":
    """取单元格自身 tcPr 内的 tcBorders（避免误取嵌套表格的边框）。"""
    pr = re.search(r"<w:tcPr>.*?</w:tcPr>", cell_xml, re.S)
    scope = pr.group(0) if pr else cell_xml
    m = re.search(r"<w:tcBorders\b.*?</w:tcBorders>", scope, re.S)
    return _parse_borders_block(m.group(0) if m else "")


def _split_cells_with_props(row_xml):
    """解析一行的单元格，返回 list[dict]：
      span    —— gridSpan（横向跨列数，默认1）
      vmerge  —— None / "restart" / "continue"（纵向合并状态）
      borders —— 该单元格 tcBorders 解析出的 {side: border}
    """
    cells = []
    for c in re.findall(r"<w:tc>.*?</w:tc>", row_xml, re.S):
        pr_m = re.search(r"<w:tcPr>.*?</w:tcPr>", c, re.S)
        pr_xml = pr_m.group(0) if pr_m else ""
        sm = re.search(r'<w:gridSpan w:val="(\d+)"', pr_xml)
        span = int(sm.group(1)) if sm else 1
        vm = re.search(r'<w:vMerge\b([^>]*?)/?>', pr_xml)
        if vm:
            vv = re.search(r'w:val="([^"]+)"', vm.group(1))
            vmerge = vv.group(1) if vv else "continue"
        else:
            vmerge = None
        cells.append({
            "span": span,
            "vmerge": vmerge,
            "borders": _cell_border_block(c),
        })
    return cells


def _build_grid(t):
    """把一张表展开成逻辑网格 occ[row][gcol]，处理 gridSpan(横向)与 vMerge(纵向)。
    合并区域内的所有网格位置共享同一个 cell 对象（用对象身份判断"是否同一逻辑单元格"）。
    返回 (tbl_borders, occ, ncols, nrows)。
    """
    tbl_b = _tbl_border_block(t)
    rows = split_rows(t)
    parsed = [_split_cells_with_props(r) for r in rows]
    ncols = max((sum(c["span"] for c in cells) for cells in parsed), default=0)
    nrows = len(rows)
    occ = [[None] * ncols for _ in range(nrows)]
    for ri, cells in enumerate(parsed):
        gi = 0
        for c in cells:
            span = c["span"]
            if c["vmerge"] == "continue" and ri > 0:
                # 纵向合并延续：与正上方同属一个逻辑单元格，复用其对象
                for k in range(span):
                    g = gi + k
                    if g < ncols:
                        occ[ri][g] = occ[ri - 1][g]
                gi += span
                continue
            cell_obj = {"borders": c["borders"]}
            for k in range(span):
                if gi + k < ncols:
                    occ[ri][gi + k] = cell_obj
            gi += span
    return tbl_b, occ, ncols, nrows


def _effective_side(cell, side, tbl_b, default_key):
    """单元格某条边的有效边线：单元格自身显式定义优先，否则继承表级缺省。
      cell        —— occ 中的单元格对象（可能为 None）
      side        —— 该单元格上的边名(top/bottom/left/right)
      default_key —— 表级缺省键：外框用同名边(top/bottom/left/right)，
                     内部横线用 insideH、内部竖线用 insideV。
    返回 border dict 或 None。
    """
    if cell is not None:
        b = cell["borders"].get(side)
        if b is not None:
            return b
    return tbl_b.get(default_key)


def _is_dark(color) -> bool:
    """黑/深灰：None(默认黑)/auto，或 RGB 三通道均 ≤ 0x80。"""
    if color is None:
        return True
    c = str(color).lower()
    if c == "auto":
        return True
    if len(c) == 6:
        try:
            return all(int(c[i:i + 2], 16) <= 0x80 for i in (0, 2, 4))
        except ValueError:
            return False
    return False


def score_borders(tables):
    """
    +3：严格对应细则的 5 个点（每个点都必须满足才加分）：
      (1) 第1~20个表格【外边框】：上、下、左、右四条边线全部连续显示；
      (2) 第1~20个表格【内部横线】：所有上下相邻单元格之间均有连续边框；
      (3) 第1~20个表格【内部竖线】：所有左右相邻单元格之间均有连续边框；
      (4) 线型统一为单实线，不混用虚线、点线或双线(val 仅 single)；
      (5) 黑色或深灰色线条，线宽统一且清晰可见(建议约0.5~1磅，即 w:sz 4~8)。

    连续性判据（修正点）：不再依赖表级 insideH/insideV 是否存在，而是把每张表
    展开为逻辑网格(处理 gridSpan 横向合并 + vMerge 纵向合并)，对【每一对实际相邻
    单元格】逐格判断它们之间那条边是否显示：
      · 内部横线：网格位置 (r, c) 与 (r+1, c)，若二者属于不同逻辑单元格(真·上下相邻)，
        则该分界处的边线 = 上格 bottom 或 下格 top，任一显示即"连续"，都缺失则失败；
        缺省继承表级 insideH。
      · 内部竖线：网格位置 (r, c) 与 (r, c+1)，若属于不同逻辑单元格，
        分界边 = 左格 right 或 右格 left，缺省继承表级 insideV。
      · 外框：每张表最外圈网格单元格朝外那条边(top/bottom/left/right)，
        缺省继承表级同名边。
    这样即便边框全部由单元格 top/bottom/left/right 逐格定义、而没有表级
    insideH/insideV，也能正确判为连续。

    线宽单位：w:sz = 1/8 磅，故 0.5~1 磅 = sz 4~8。
    """
    tbls = tables[:20]  # 准确覆盖第 1~20 个表格

    bad_outer, bad_h, bad_v, bad_type, bad_color = [], [], [], [], []
    all_sz: "set[int]" = set()

    def _both_missing(b1, b2):
        """两条候选边线都不显示(缺失或 none/nil)才判"不连续"。"""
        return not _border_present(b1) and not _border_present(b2)

    for idx, t in enumerate(tbls, 1):
        tbl_b, occ, ncols, nrows = _build_grid(t)
        if nrows == 0 or ncols == 0:
            continue

        # (1) 外边框：最外圈朝外的四条边（缺省继承表级同名边）
        outer_bad_sides = set()
        for c in range(ncols):
            if not _border_present(_effective_side(occ[0][c], "top", tbl_b, "top")):
                outer_bad_sides.add("top")
            if not _border_present(_effective_side(occ[nrows - 1][c], "bottom", tbl_b, "bottom")):
                outer_bad_sides.add("bottom")
        for r in range(nrows):
            if not _border_present(_effective_side(occ[r][0], "left", tbl_b, "left")):
                outer_bad_sides.add("left")
            if not _border_present(_effective_side(occ[r][ncols - 1], "right", tbl_b, "right")):
                outer_bad_sides.add("right")
        for s in ("top", "bottom", "left", "right"):
            if s in outer_bad_sides:
                bad_outer.append((idx, s))

        # (2) 内部横线：逐格判断上下相邻(不同逻辑单元格)之间的分界边
        h_broken = False
        for r in range(nrows - 1):
            for c in range(ncols):
                up, down = occ[r][c], occ[r + 1][c]
                if up is None or down is None or up is down:
                    continue  # 同一逻辑单元格(纵向合并)内部无分界，不要求边线
                b_bottom = _effective_side(up, "bottom", tbl_b, "insideH")
                b_top = _effective_side(down, "top", tbl_b, "insideH")
                if _both_missing(b_bottom, b_top):
                    h_broken = True
                    break
            if h_broken:
                break
        if h_broken:
            bad_h.append(idx)

        # (3) 内部竖线：逐格判断左右相邻(不同逻辑单元格)之间的分界边
        v_broken = False
        for r in range(nrows):
            for c in range(ncols - 1):
                left, right = occ[r][c], occ[r][c + 1]
                if left is None or right is None or left is right:
                    continue  # 同一逻辑单元格(横向合并)内部无分界
                b_right = _effective_side(left, "right", tbl_b, "insideV")
                b_left = _effective_side(right, "left", tbl_b, "insideV")
                if _both_missing(b_right, b_left):
                    v_broken = True
                    break
            if v_broken:
                break
        if v_broken:
            bad_v.append(idx)

        # (4)(5) 线型/颜色/线宽：遍历该表所有实际显示的边线(表级 + 单元格级)
        seen_cells = set()
        borders_iter = list(tbl_b.values())
        for row in occ:
            for cell in row:
                if cell is None or id(cell) in seen_cells:
                    continue
                seen_cells.add(id(cell))
                borders_iter.extend(cell["borders"].values())

        vals_here = set()
        color_bad = None
        width_bad = None
        for b in borders_iter:
            if not _border_present(b):
                continue
            v = b.get("val")
            vals_here.add(v)
            if color_bad is None and not _is_dark(b.get("color")):
                color_bad = f"颜色{b.get('color')}"
            sz = b.get("sz")
            if isinstance(sz, int):
                all_sz.add(sz)
                if width_bad is None and not (4 <= sz <= 8):
                    width_bad = f"线宽sz={sz}不在4~8(0.5~1磅)"
        # (4) 线型统一 single
        non_single = vals_here - {"single"}
        if non_single:
            bad_type.append((idx, sorted(str(x) for x in non_single)))
        # (5) 颜色/线宽越界
        if color_bad or width_bad:
            bad_color.append((idx, color_bad or width_bad))

    # (5) 线宽"统一"：全文档边线 sz 取值集合应唯一（不混用不同线宽）
    width_uniform = len(all_sz) <= 1

    points = [
        ("外边框上/下/左/右四条边线全部连续显示", len(bad_outer) == 0,
         "全部满足" if not bad_outer else "不连续: " + ",".join(f"T{i}-{s}" for i, s in bad_outer[:6])),
        ("内部横线：所有上下相邻单元格之间均有连续边框", len(bad_h) == 0,
         "全部满足" if not bad_h else "缺失/不连续: " + ",".join(f"T{i}" for i in bad_h[:8])),
        ("内部竖线：所有左右相邻单元格之间均有连续边框", len(bad_v) == 0,
         "全部满足" if not bad_v else "缺失/不连续: " + ",".join(f"T{i}" for i in bad_v[:8])),
        ("线型统一为单实线，不混用虚线/点线/双线", len(bad_type) == 0,
         "全部 single" if not bad_type else "混用: " + ",".join(f"T{i}{v}" for i, v in bad_type[:6])),
        ("黑色或深灰色，线宽统一且清晰可见(约0.5~1磅=sz4~8)", len(bad_color) == 0 and width_uniform,
         (f"颜色合规且线宽统一(sz={sorted(all_sz)} ≈ {[round(s/8,2) for s in sorted(all_sz)]}磅)"
          if (not bad_color and width_uniform)
          else ("颜色/线宽问题: " + ",".join(f"T{i}{r}" for i, r in bad_color[:6])
                + ("" if width_uniform else f" 线宽不统一sz={sorted(all_sz)}")))),
    ]

    ok = all(p_ok for _, p_ok, _ in points)
    lines = [f"      ({'OK' if p_ok else 'NG'}) {name} -> {info}" for name, p_ok, info in points]
    detail = ("细则5个点逐项核对：\n" + "\n".join(lines))
    return ok, detail


def _count_page_breaks(s: str) -> int:
    """统计一段 document.xml 文本中的"分页数"（无需 COM/渲染）。

    依据 Word/WPS 保存文档时写入的分页标记：
      · <w:lastRenderedPageBreak/> —— 上次渲染得到的【自动分页】位置（最可靠）；
      · <w:br w:type="page"/>       —— 手动分页符；
      · <w:pageBreakBefore/>        —— 段落"段前分页"（w:val 为 false/0/off 时不生效）。
    为避免重复计数：若存在 lastRenderedPageBreak（说明文档被真实渲染过，
    自动分页已被记录），以它 + 手动分页符为准；否则退回手动分页符 + 段前分页。
    """
    lrpb = len(re.findall(r'<w:lastRenderedPageBreak\b', s))
    manual = len(re.findall(r'<w:br\b[^>]*w:type="page"', s))
    if lrpb > 0:
        return lrpb + manual
    pbb = len(re.findall(
        r'<w:pageBreakBefore\b(?![^>]*w:val="(?:false|0|off)")', s))
    return manual + pbb


def _table_page_number(table_xml: str, doc_xml: str):
    """估算 table_xml 在 doc_xml 中所处的 1-based 页码。

    页码 = 1 + 该表格【起始位置之前】出现的分页数。
    返回 (page:int|None, determinable:bool)：
      · determinable=False 表示文档完全没有任何分页标记（未被办公软件渲染过、
        无法在不借助 COM/渲染的情况下判定页码），此时 page=None，调用方应
        降级为"仅按内容匹配"，不因缺信息而误判。
    """
    if not doc_xml:
        return None, False
    if _count_page_breaks(doc_xml) == 0:
        return None, False
    pos = doc_xml.find(table_xml)
    if pos < 0:
        return None, False
    return _count_page_breaks(doc_xml[:pos]) + 1, True


def find_target_table(tables, doc_xml: str = "", required_page: int = 5):
    """定位"含起始时间/专业技术人员类别/指导内容/指导情况"的目标表（第5页十行表）。

    定位策略（在不使用 COM/渲染的前提下）：
      1) 先按内容标记（"起始时间"+"指导内容"+"指导情况"）筛出候选表；
      2) 若提供了 doc_xml 且文档存在分页标记（lastRenderedPageBreak / 手动分页 /
         pageBreakBefore），估算每个候选表所处页码，优先返回位于 required_page 页
         的候选；若均不在该页，返回 None（表明"第5页表"缺失/位置不符）。
      3) 若 doc_xml 无分页信息（未被办公软件渲染过），降级为仅按内容返回首个候选，
         避免因分页信号缺失产生误判。
    """
    candidates = [t for t in tables
                  if "起始时间" in t and "指导内容" in t and "指导情况" in t]
    if not candidates:
        return None
    if not doc_xml:
        return candidates[0]
    # 若文档没有任何分页信号，降级为内容匹配
    any_break = _count_page_breaks(doc_xml) > 0
    if not any_break:
        return candidates[0]
    for t in candidates:
        page, ok = _table_page_number(t, doc_xml)
        if ok and page == required_page:
            return t
    return None


def score_row6_text(tables, doc_xml: str = ""):
    """
    +3：目标表共 10 行，第 6 行的第 2/3/4/5 格依次为
        "起始时间""专业技术人员类别""指导内容""指导情况"，
        宋体五号(sz=21)，单元格内水平居中、垂直居中，文字完整不压边框。
    每个子点必须全部满足才加分。
    目标表须为第5页表格；若 doc_xml 无分页信号则降级为内容匹配。
    """
    t = find_target_table(tables, doc_xml, required_page=5)
    reasons = []
    if t is None:
        return False, "未在第5页找到目标表格（含起始时间/指导内容/指导情况的十行表）"
    rows = split_rows(t)
    if len(rows) != 10:
        reasons.append(f"目标表行数为 {len(rows)}（要求 10 行）")
    expected = ["起始时间", "专业技术人员类别", "指导内容", "指导情况"]
    if len(rows) >= 6:
        cells = split_cells(rows[5])  # 第6行
        # 第 2~5 格（index 1..4）
        for i, exp in enumerate(expected, start=1):
            if i >= len(cells):
                reasons.append(f"第6行缺少第{i+1}格")
                continue
            c = cells[i]
            txt = cell_text(c).strip()
            if txt != exp:
                reasons.append(f"第6行第{i+1}格文字为'{txt}'(应为'{exp}')")
                continue
            # 字体宋体
            if "宋体" not in c:
                reasons.append(f"第{i+1}格'{exp}'非宋体")
            # 五号 = sz 21（半磅，宋体五号≈10.5磅 -> w:sz=21）
            sz = re.search(r'<w:sz w:val="(\d+)"', c)
            if not sz or sz.group(1) != "21":
                reasons.append(f"第{i+1}格'{exp}'字号非五号(sz应为21,实为{sz.group(1) if sz else '无'})")
            # 水平居中
            if not re.search(r'<w:jc w:val="center"', c):
                reasons.append(f"第{i+1}格'{exp}'非水平居中")
            # 垂直居中
            if not re.search(r'<w:vAlign w:val="center"', c):
                reasons.append(f"第{i+1}格'{exp}'非垂直居中")
            # 文字完整显示且不压住边框：单元格不得有明显负缩进将文字挤出边框
            # （沿用本项目判据：left/start 负缩进 <= -200dxa 视为压线）
            neg = re.search(r'<w:ind\b[^>]*w:(?:left|start)="(-\d+)"', c)
            if neg and int(neg.group(1)) <= -200:
                reasons.append(f"第{i+1}格'{exp}'负缩进{neg.group(1)}压住边框")
    else:
        reasons.append("目标表不足6行，无法检查第6行")
    ok = len(reasons) == 0
    detail = ("目标表为10行，第6行第2~5格依次为'起始时间/专业技术人员类别/指导内容/指导情况'，"
              "均为宋体五号、水平且垂直居中"
              if ok else "; ".join(reasons))
    return ok, detail


def score_col_widths(tables, doc_xml: str = ""):
    """
    +5：目标表（第5页十行表）第2列列宽均为1.64cm、第3列均为3.18cm、
        第4列均为7.94cm、第5列均为2.22cm。
    逻辑列 → 底层 gridCol 区间(0-based, 闭区间)：
        第2列=网格[1]; 第3列=网格[2,3]; 第4列=网格[4..8]; 第5列=网格[9]。

    判定策略（避免"合并即失败"的过严误判）：
      1) 主判据 —— 按 <w:tblGrid> 计算该表整体的视觉列边界，
         逐一比较逻辑列的目标宽度。tblGrid 是全表所有行共享的物理网格，
         合并单元格并不改变网格边界，故这一步足以覆盖 rubric 的"均为"要求。
      2) 辅助判据 —— 若某行的单元格显式给出 tcW(w:type="dxa") 与其所跨
         gridCol 区间对应的目标累计宽度冲突，则该行视觉宽度确实不符，
         记为该行该列不达标；仅在"视觉宽度确实不符"时失败，
         不再仅因跨列合并本身而失败。
      3) tcW 非 dxa（auto/pct）不作为绝对宽度比较，跳过。
    判据阈值：宽度与目标值误差 <= 0.06cm（约 2 像素，允许排版四舍五入）。
    目标表须为第5页表格；若 doc_xml 无分页信号则降级为内容匹配。
    """
    t = find_target_table(tables, doc_xml, required_page=5)
    if t is None:
        return False, "未在第5页找到目标表格，无法校验列宽"
    grid = [int(x) for x in re.findall(r'<w:gridCol w:w="(\d+)"', t)]
    if len(grid) < 10:
        return False, f"目标表底层网格列数为 {len(grid)}，无法定位逻辑列"

    # 逻辑列 -> (底层 gridCol 区间[a,b], 目标宽度cm)
    logical = [
        ("第2列", (1, 1), 1.64),
        ("第3列", (2, 3), 3.18),
        ("第4列", (4, 8), 7.94),
        ("第5列", (9, 9), 2.22),
    ]
    tol = 0.06
    reasons = []

    # 1) 主判据：tblGrid 视觉列边界
    for name, (a, b), target_cm in logical:
        w = sum(grid[a:b + 1])
        cm = w / DXA_PER_CM
        if abs(cm - target_cm) > tol:
            reasons.append(f"tblGrid {name}={cm:.2f}cm(应{target_cm}cm)")

    # 2) 辅助判据：逐行核查显式 tcW(dxa) 与所跨 gridCol 区间目标之和是否冲突
    logical_map = {(a, b): (name, tcm) for name, (a, b), tcm in logical}
    for ri, r in enumerate(split_rows(t), 1):
        gi = 0
        for c in split_cells(r):
            sm = re.search(r'<w:gridSpan w:val="(\d+)"', c)
            span = int(sm.group(1)) if sm else 1
            a, b = gi, gi + span - 1
            gi += span
            wm = re.search(r'<w:tcW w:w="(-?\d+)"', c)
            if not wm:
                continue
            # 仅在明确为 dxa 时比较绝对宽度；pct/auto/nil 不适用绝对判定
            tm = re.search(r'<w:tcW\b[^>]*w:type="(\w+)"', c)
            if tm and tm.group(1) != "dxa":
                continue
            w = int(wm.group(1))
            cm = w / DXA_PER_CM
            if (a, b) in logical_map:
                name, tcm = logical_map[(a, b)]
                if abs(cm - tcm) > tol:
                    reasons.append(f"第{ri}行{name}(tcW)={cm:.2f}cm(应{tcm}cm)")
            else:
                covered = [(nm, tcm) for (la, lb), (nm, tcm) in logical_map.items()
                           if la >= a and lb <= b]
                if covered:
                    target_sum = sum(tcm for _, tcm in covered)
                    if abs(cm - target_sum) > tol:
                        names = "/".join(nm for nm, _ in covered)
                        reasons.append(
                            f"第{ri}行合并列({names})(tcW)={cm:.2f}cm(应{target_sum:.2f}cm)")

    ok = len(reasons) == 0
    detail = ("目标表 tblGrid 第2/3/4/5列列宽均为 1.64/3.18/7.94/2.22cm，"
              "且各行 tcW 与之一致"
              if ok else "列宽不符: " + "; ".join(reasons[:8])
              + (f" 等共{len(reasons)}处" if len(reasons) > 8 else ""))
    return ok, detail


# ===========================================================================
# 维度 2 - 扣分点（任一子点命中即扣分）
# ===========================================================================
def penalty_image_cover(tables, doc_xml):
    """-5：文档表格被转换为图片或使用截图覆盖，导致内容无法编辑。
    细则两种情形，命中任一即扣分：
      (1) 表格被转换为图片：表格内含图片元素(<w:drawing>/<w:pict>)；
      (2) 使用截图覆盖：表格被一张图替代/盖住。
    二者共同后果是【内容无法编辑】——即该表几乎没有可编辑文字。
    判据：某表格含图片(被转换/被截图覆盖) 且 其可编辑文字总长 < 5
          (基本无可编辑内容) -> 视为内容无法编辑，命中扣分。
    """
    covered = []
    for ti, t in enumerate(tables, 1):
        has_img = ("<w:drawing>" in t) or ("<w:pict>" in t)
        txt_len = len("".join(cell_text(c) for c in split_cells(t)).strip())
        if has_img and txt_len < 5:
            covered.append(ti)
    hit = len(covered) > 0
    detail = (f"表 {covered} 被转换为图片/截图覆盖，内容无法编辑" if hit
              else "未发现表格被转换为图片或被截图覆盖，内容均可编辑")
    return hit, detail


def plain_text(xml):
    """去掉所有标签，得到文档纯文本（用于跨 run 的整体文字检索）。"""
    return re.sub(r"<[^>]+>", "", xml)


def penalty_pages_and_marks(pages, doc_xml, settings_xml):
    """
    -3：文档页数不是25页，或出现无关空白页、批注、修订标记及临时说明文字。
    细则 5 种情形，命中任一即扣分，逐点核对：
      (1) 页数不是25页：Pages != 25；
      (2) 无关空白页：文档中存在分页符 <w:br w:type="page"> / <w:lastRenderedPageBreak>
          之后紧跟无任何可见文字的整页内容（以"显式分页符"出现作为空白页判据）；
      (3) 批注：存在 <w:commentReference> / <w:commentRangeStart>；
      (4) 修订标记：存在插入/删除 <w:ins>/<w:del> 或文档处于修订(trackChanges)状态；
      (5) 临时说明文字：正文出现"批注/说明：/备注/TODO/待补充/示例/请填写/占位"等临时性说明字样。
    """
    reasons = []
    # (1) 页数
    if pages != 25:
        reasons.append(f"页数为 {pages}(应为25页)")
    # (2) 无关空白页：以显式分页符 <w:br w:type="page"> 作为存在额外/空白分页的判据
    if re.search(r'<w:br\b[^>]*w:type="page"', doc_xml):
        reasons.append("出现显式分页符，疑似无关空白页")
    # (3) 批注
    if re.search(r"<w:commentReference\b", doc_xml) or "<w:commentRangeStart" in doc_xml:
        reasons.append("出现批注")
    # (4) 修订标记
    if re.search(r"<w:(ins|del)\b", doc_xml):
        reasons.append("出现修订标记(插入/删除)")
    if settings_xml and "<w:trackChanges" in settings_xml:
        reasons.append("文档处于修订(trackChanges)状态")
    # (5) 临时说明文字（指人为加入的临时性提示/占位字样，不含表格本身的固定标题）
    flat = plain_text(doc_xml)
    note_kw = ["TODO", "待补充", "请填写", "占位", "此处填写", "临时说明", "批注："]
    hit_kw = [k for k in note_kw if k in flat]
    if hit_kw:
        reasons.append(f"出现临时说明文字: {hit_kw}")

    hit = len(reasons) > 0
    detail = ("; ".join(reasons) if hit
              else f"页数为 {pages} 页，无空白页/批注/修订标记/临时说明文字")
    return hit, detail


# ===========================================================================
# 主流程
# ===========================================================================
def get_pages(zf):
    app = load_part(zf, "docProps/app.xml")
    if not app:
        return None
    m = re.search(r"<Pages>(\d+)</Pages>", app)
    return int(m.group(1)) if m else None


def find_target_doc(dir_path):
    """在 dir_path 目录内定位被评估的 .docx 文件（排除 Word 临时文件 ~$）。"""
    candidates = [
        p for p in glob.glob(os.path.join(dir_path, "*.docx"))
        if not os.path.basename(p).startswith("~$")
    ]
    return candidates[0] if candidates else None


def evaluate(dir_path: str) -> dict:
    """
    统一入口：接收"脚本所在目录的路径"，脚本自己在该目录内定位并打开被评估的文档。
    返回结构化字典，字段含义见项目约定文档 §2.2。
    """
    file_name = None
    try:
        path = find_target_doc(dir_path)
        if path is None:
            return {
                "id": SCRIPT_ID,
                "file_name": None,
                "status": "error",
                "error": f"目录 {dir_path} 内未找到待评估的 .docx 文件",
                "dim1_pass": False,
                "dim1_reason": "",
                "dim2_items": [],
                "total_score": 0,
                "max_score": 11,
            }
        file_name = os.path.basename(path)

        zf = zipfile.ZipFile(path)
        try:
            doc_xml = load_part(zf, "word/document.xml")
            settings_xml = load_part(zf, "word/settings.xml")
            pages = get_pages(zf)
            tables = split_tables(doc_xml) if doc_xml else []

            # ---------- 维度1 ----------
            d1_pass, d1_details = check_dimension1(path, zf, doc_xml, pages)

            if not d1_pass:
                d1_reason = "；".join(msg for ok, msg in d1_details if not ok)
                return {
                    "id": SCRIPT_ID,
                    "file_name": file_name,
                    "status": "ok",
                    "error": None,
                    "dim1_pass": False,
                    "dim1_reason": d1_reason,
                    "dim2_items": [],
                    "total_score": 0,
                    "max_score": 11,
                }

            # ---------- 维度2 ----------
            dim2_items = []
            total = 0

            # 得分点：(满分, 细则内容, 判定结果)
            add_rules = [
                (3, "表格的外边框（上、下、左、右四条边线）全部连续显示，内部所有相邻单元格之间均有连续边框，线型统一为单实线，黑色或深灰色线条，线宽统一且清晰可见。",
                 score_borders(tables)),
                (3, "第5页表格一共十行，第六行的第二格、第三格、第四格、第五格依次添加文字“起始时间”“专业技术人员类别”“指导内容”“指导情况”字体为宋体五号。在单元格内水平居中、垂直居中，文字完整显示且不压住边框。",
                 score_row6_text(tables, doc_xml or "")),
                (5, "第5页表格第二列列宽均为1.64cm、第三列列宽均为3.18cm、第四列列宽均为7.94cm、第五列列宽均为2.22cm。",
                 score_col_widths(tables, doc_xml or "")),
            ]
            for pts, rule, (ok, _detail) in add_rules:
                delta = pts if ok else 0
                total += delta
                dim2_items.append({
                    "rule": rule,
                    "max_delta": pts,
                    "delta": delta,
                    "hit": ok,
                    "detail": "",
                })

            # 扣分点：(扣分, 细则内容, 判定结果)
            pen_rules = [
                (-5, "文档表格被转换为图片或使用截图覆盖，导致内容无法编辑。",
                 penalty_image_cover(tables, doc_xml)),
                (-3, "页数不是25页，或出现无关空白页、批注、修订标记及临时说明文字。",
                 penalty_pages_and_marks(pages, doc_xml, settings_xml)),
            ]
            for pts, rule, (hit, _detail) in pen_rules:
                delta = pts if hit else 0
                total += delta
                dim2_items.append({
                    "rule": rule,
                    "max_delta": pts,
                    "delta": delta,
                    "hit": hit,
                    "detail": "",
                })

            return {
                "id": SCRIPT_ID,
                "file_name": file_name,
                "status": "ok",
                "error": None,
                "dim1_pass": True,
                "dim1_reason": "",
                "dim2_items": dim2_items,
                "total_score": total,
                "max_score": 11,
            }
        finally:
            zf.close()
    except Exception as e:
        return {
            "id": SCRIPT_ID,
            "file_name": file_name,
            "status": "error",
            "error": str(e),
            "dim1_pass": False,
            "dim1_reason": "",
            "dim2_items": [],
            "total_score": 0,
            "max_score": 11,
        }


if __name__ == "__main__":
    target_dir = sys.argv[1] if len(sys.argv) > 1 else os.path.dirname(os.path.abspath(__file__))
    result = evaluate(target_dir)
    print(json.dumps(result, ensure_ascii=False, indent=2))
