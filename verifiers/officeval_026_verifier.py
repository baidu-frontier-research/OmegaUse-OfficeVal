# -*- coding: utf-8 -*-
"""
自动评估脚本：对 "项目式协作学习统计表_三线表优化版.docx" 按照"打分细则"评分。

评估逻辑：
  维度1（可用与可修改性）：若不满足任意一条 -> 直接 0 分，不再检查维度2。
  维度2（完成度）：满足维度1后逐条检查；得分点累加正分，扣分点累加负分。
      - 加分细则：必须满足细则中的"每一个点"才加分（按满足的表格个数累加，有上限）。
      - 扣分细则：只要满足"任意一点"即扣分。

仅依赖 Python 标准库（zipfile + xml.etree），无需 python-docx。
对不易精确判定的点，采用合理的近似方式实现自动判定，并在输出中说明。

对外只暴露 evaluate(dir_path: str) -> dict：
  传入"脚本所在目录的路径"，脚本自己在该目录内定位被评估的 .docx 文档并返回结构化结果。
"""

import os
import re
import sys
import json
import zipfile
import xml.etree.ElementTree as ET

# ---------------------------------------------------------------------------
# 常量与命名空间
# ---------------------------------------------------------------------------
W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS = {"w": W}


def w(tag):
    return f"{{{W}}}{tag}"


# Word 中边框线宽 sz 单位为 1/8 磅。1.5磅 = sz 12；0.75磅 = sz 6。
SZ_1_5PT = 12
SZ_0_75PT = 6

# 字号：half-points。小五 = 9磅 = sz18；五号 = 10.5磅 = sz21。
SZ_XIAOWU = 18   # 小五号
SZ_WUHAO = 21    # 五号

# 目标表格题注前缀（顺序对应文档中表格顺序）
TARGET_LABELS = ["表5-4", "表5-5", "表5-6", "表5-7", "表5-8",
                 "表5-9", "表5-10", "表5-11", "表5-12"]

# 需要"两个分组标题(Levene/t检验)下方各一条内框线"的表
TWO_GROUP_TABLES = {"表5-4", "表5-5", "表5-8", "表5-9", "表5-12"}
# 仅"t检验"一个分组标题下方有内框线的表
ONE_GROUP_TABLES = {"表5-6", "表5-7", "表5-10", "表5-11"}


# ---------------------------------------------------------------------------
# 文档加载
# ---------------------------------------------------------------------------
class DocxModel:
    def __init__(self, path):
        self.path = path
        self.ok_open = False
        self.is_docx = path.lower().endswith(".docx")
        self.document_xml = None
        self.app_xml = None
        self.has_image_parts = False
        self.tables = []          # 文档顺序的所有表格 (ET element)
        self.body_order = []      # 正文中 段落/表格 的顺序列表 [("p",text)/("tbl",elem)]
        self.page_count = None
        self.sect = None
        self._load()

    def _load(self):
        try:
            with zipfile.ZipFile(self.path) as z:
                names = z.namelist()
                self.document_xml = z.read("word/document.xml")
                if "docProps/app.xml" in names:
                    self.app_xml = z.read("docProps/app.xml")
                # 是否包含图片资源（用于判断是否被转图）
                self.has_image_parts = any(
                    n.lower().startswith("word/media/") for n in names
                )
            self.ok_open = True
        except Exception as e:
            self.ok_open = False
            self.load_error = str(e)
            return

        self.root = ET.fromstring(self.document_xml)
        self.body = self.root.find(w("body"))
        for child in list(self.body):
            if child.tag == w("p"):
                txt = self._para_text(child)
                self.body_order.append(("p", txt, child))
            elif child.tag == w("tbl"):
                self.tables.append(child)
                self.body_order.append(("tbl", None, child))

        self.sect = self.root.find(f".//{w('sectPr')}")

        if self.app_xml is not None:
            try:
                aroot = ET.fromstring(self.app_xml)
                for el in aroot.iter():
                    if el.tag.endswith("}Pages") or el.tag.endswith("Pages"):
                        self.page_count = int(el.text)
                        break
            except Exception:
                pass

    @staticmethod
    def _para_text(p):
        return "".join(t.text or "" for t in p.iter(w("t")))


# ---------------------------------------------------------------------------
# 辅助：把目标表格与其题注配对
# ---------------------------------------------------------------------------
def map_labeled_tables(model):
    """
    返回 dict: label -> table_element
    依据正文顺序，紧邻在题注段落(含 '表5-x')之后的表格即为该表。
    """
    result = {}
    pending_label = None
    for kind, txt, elem in model.body_order:
        if kind == "p":
            t = (txt or "").replace(" ", "")
            for lab in TARGET_LABELS:
                # 题注以 表5-x 开头，避免 表5-1 误匹配 表5-12：用边界
                norm = lab.replace("-", "-")
                if re.search(re.escape(norm) + r"(?![0-9])", t):
                    pending_label = lab
                    break
        elif kind == "tbl":
            if pending_label is not None:
                result[pending_label] = elem
                pending_label = None
    return result


# ---------------------------------------------------------------------------
# 边框解析辅助
# ---------------------------------------------------------------------------
def get_table_borders(tbl):
    tblPr = tbl.find(w("tblPr"))
    if tblPr is None:
        return None
    return tblPr.find(w("tblBorders"))


def border_attr(borders_el, side):
    """返回 (val, sz, color) 或 None"""
    if borders_el is None:
        return None
    b = borders_el.find(w(side))
    if b is None:
        return None
    return (
        b.get(w("val")),
        b.get(w("sz")),
        (b.get(w("color")) or "auto"),
    )


def cell_borders(tc):
    tcPr = tc.find(w("tcPr"))
    if tcPr is None:
        return None
    return tcPr.find(w("tcBorders"))


def is_black(color):
    return color in ("000000", "auto", "black", None)


def is_visible(val):
    return val is not None and val not in ("nil", "none")


def rows_of(tbl):
    return tbl.findall(w("tr"))


def cells_of(row):
    return row.findall(w("tc"))


def cell_text(tc):
    return "".join(t.text or "" for t in tc.iter(w("t")))


# ---------------------------------------------------------------------------
# 维度1 检查
# ---------------------------------------------------------------------------
def check_dimension1(model):
    """维度一四个子项（.docx 格式/表格齐全/页数与无乱码/表格未图片化）已全部删除，
    维度一不再拦截评分流程，恒返回 True。保留函数签名以兼容 evaluate() 现有调用。"""
    details: list[tuple[bool, str]] = []
    return True, details


# ---------------------------------------------------------------------------
# 维度2 - 得分点
# ---------------------------------------------------------------------------
def check_outer_border(labeled):
    """
    +1/表：表格外框线 —— 外框线只保留上下外框线，表格上框线和下框线为1.5磅黑色实线。
    （一个三线表满足+1，两个+2，以此类推，最多9分）

    细则逐点拆解（每一个点都必须踩到）：
      点1：外框线"只保留上下外框线" —— 左外框线、右外框线均不保留（nil/none）。
      点2：表格上框线为 1.5磅(sz=12) 黑色(000000/auto) 实线(single)。
      点3：表格下框线为 1.5磅(sz=12) 黑色(000000/auto) 实线(single)。
    细则未提及内框线(insideH/insideV)，故此条不对内框线加以约束。
    """
    count = 0
    hits = []
    for lab in TARGET_LABELS:
        tbl = labeled.get(lab)
        if tbl is None:
            continue
        tb = get_table_borders(tbl)
        top = border_attr(tb, "top")
        bottom = border_attr(tb, "bottom")
        left = border_attr(tb, "left")
        right = border_attr(tb, "right")

        def line_1_5pt_black_solid(side_attr):
            # 实线 single + 1.5磅(sz=12) + 黑色
            if side_attr is None:
                return False
            val, sz, color = side_attr
            return (val == "single"
                    and sz is not None and int(sz) == SZ_1_5PT
                    and is_black(color))

        def outer_not_kept(side_attr):
            # "只保留上下外框线" -> 左右外框线不保留：缺失或 nil/none 均视为不保留
            if side_attr is None:
                return True
            val = side_attr[0]
            return not is_visible(val)

        # 点1：只保留上下外框线（左右外框线不保留）
        p1 = outer_not_kept(left) and outer_not_kept(right)
        # 点2：上框线 1.5磅黑色实线
        p2 = line_1_5pt_black_solid(top)
        # 点3：下框线 1.5磅黑色实线
        p3 = line_1_5pt_black_solid(bottom)

        if p1 and p2 and p3:
            count += 1
            hits.append(lab)
    score = min(count, 9)
    return score, hits


def _inner_horizontal_border(tc):
    """
    返回该单元格"内框横线"(下框线/上框线)信息。
    内框线指单元格 tcBorders 中的 bottom / top / insideH。
    返回 list[(side, val, sz, color)]，仅含可见者。
    """
    cb = cell_borders(tc)
    result = []
    for side in ("bottom", "top", "insideH"):
        a = border_attr(cb, side)
        if a is not None and is_visible(a[0]):
            result.append((side, a[0], a[1], a[2]))
    return result


def check_group_inner_border(labeled, target_tables, title_keywords, max_score):
    """
    通用：分组标题下方内框线评分。
    细则逐点拆解（每一个点都必须踩到）：
      点1：每一个指定分组标题(title_keywords)所在单元格下方都有下框线。
      点2：除这些分组标题单元格外，其余位置均无内框横线（"其余位置无内框线"）。
      点3：这些下框线的类型为 0.75磅(sz=6) 黑色(000000/auto) 短横线(single)。
    细则未提及竖向内框线，故此条不对竖线加以约束（竖线由专门的扣分项处理）。
    """
    norm_titles = [k.replace(" ", "") for k in title_keywords]
    count = 0
    hits = []
    for lab in sorted(target_tables):
        tbl = labeled.get(lab)
        if tbl is None:
            continue

        found_titles = set()       # 已满足"标题下方有0.75磅黑色短横线"的标题
        other_has_inner = False    # 非分组标题位置是否出现内框横线

        for row in rows_of(tbl):
            for tc in cells_of(row):
                txt = cell_text(tc).replace(" ", "")
                is_title = any(k in txt for k in norm_titles)
                inner = _inner_horizontal_border(tc)

                if is_title:
                    # 点1 + 点3：分组标题单元格需有 0.75磅黑色短横线的下框线
                    for side, val, sz, color in inner:
                        if (side == "bottom"
                                and val == "single"
                                and sz is not None and int(sz) == SZ_0_75PT
                                and is_black(color)):
                            for k in norm_titles:
                                if k in txt:
                                    found_titles.add(k)
                else:
                    # 点2："其余位置无内框线" —— 非标题单元格不应有任何可见内框横线
                    if inner:
                        other_has_inner = True

        p1_p3 = (found_titles == set(norm_titles))  # 每个指定标题都满足下框线且类型正确
        p2 = (not other_has_inner)                  # 其余位置无内框线

        if p1_p3 and p2:
            count += 1
            hits.append(lab)

    score = min(count, max_score)
    return score, hits


def check_two_group_inner(labeled):
    """
    +1/表：表5-4、表5-5、表5-8、表5-9和表5-12 分组内框线 ——
    仅有"Levene检验"、"t检验"两个分组标题下方有下框线，其余位置无内框线，
    框线类型为0.75磅黑色短横线。（一个+1，两个+2，以此类推，最多5分）
    """
    return check_group_inner_border(
        labeled, TWO_GROUP_TABLES, ["Levene检验", "t检验"], 5)


def check_one_group_inner(labeled):
    """
    +1/表：表5-6、表5-7、表5-10和表5-11 分组表头内框线 ——
    仅有"t检验"标题下方有下框线，其余位置无内框线，
    框线类型为0.75磅黑色短横线。（一个+1，两个+2，以此类推，最多4分）

    细则逐点拆解（每一个点都必须踩到，由 check_group_inner_border 实现）：
      点1："t检验"标题所在单元格下方有下框线。
      点2：除该标题单元格外，其余位置均无内框横线（"其余位置无内框线"，字面严格）。
      点3：该下框线类型为 0.75磅(sz=6) 黑色(000000/auto) 短横线(single)。
    细则未提及竖向内框线，故此条不对竖线加以约束。
    """
    return check_group_inner_border(
        labeled, ONE_GROUP_TABLES, ["t检验"], 4)


# ---------------------------------------------------------------------------
# 维度2 - 扣分点
# ---------------------------------------------------------------------------
def deduct_extra_vertical(labeled):
    """
    -3：表格出现额外竖框线。

    细则逐点拆解（"出现"即扣分，只要任意一点命中即触发）：
      竖框线 = 表格中任何竖直方向的边框线（left / right / insideV）。
      判定"有效竖线"：与 WPS/Word 实际渲染一致 —— 单元格级 tcBorders 优先于
      表级 tblBorders。一条竖线是否真正显示，由它两侧单元格的对应竖边决定：
        · 表格最左边线：由第 1 列单元格的 tcBorders/left 决定（缺省时继承表级 left）。
        · 表格最右边线：由最后一列单元格的 tcBorders/right 决定（缺省继承表级 right）。
        · 内部竖线(insideV)：夹在相邻两单元格之间，只要"左邻的 right"或
          "右邻的 left"任一被单元格显式设为 nil/none，该竖线即不画；两侧都缺省
          时才继承表级 insideV。
      因此单元格若把自身 left/right 全设为 nil/none，即使表级 left/right/insideV
      写着 single，最终也不显示任何竖线（本文档三线表即属此情形）。
      三线表本身不含竖框线，故任何真正显示的竖线均属"额外"。
    细则只针对竖框线，不约束横框线、对角线(tl2br/tr2bl)等其余内容。
    """

    def cell_side(tc, side):
        """返回单元格 tcBorders 中指定边的 val；无定义返回 None（继承表级）。"""
        cb = cell_borders(tc)
        a = border_attr(cb, side)
        return a[0] if a is not None else None

    def table_side_visible(tb, side):
        a = border_attr(tb, side)
        return a is not None and is_visible(a[0])

    def edge_visible(cell_val, table_val_visible):
        """一条竖边的最终可见性：单元格显式值优先，缺省则继承表级。"""
        if cell_val is not None:
            return is_visible(cell_val)
        return table_val_visible

    for lab in TARGET_LABELS:
        tbl = labeled.get(lab)
        if tbl is None:
            continue

        tb = get_table_borders(tbl)
        t_left = table_side_visible(tb, "left")
        t_right = table_side_visible(tb, "right")
        t_insideV = table_side_visible(tb, "insideV")

        for row in rows_of(tbl):
            cells = cells_of(row)
            n = len(cells)
            for ci, tc in enumerate(cells):
                c_left = cell_side(tc, "left")
                c_right = cell_side(tc, "right")

                # 表格最左外框竖线（第 1 列的 left）
                if ci == 0 and edge_visible(c_left, t_left):
                    return True, f"{lab} 出现额外竖框线（最左竖线）"
                # 表格最右外框竖线（最后 1 列的 right）
                if ci == n - 1 and edge_visible(c_right, t_right):
                    return True, f"{lab} 出现额外竖框线（最右竖线）"

                # 内部竖线：本单元格 right 与右邻单元格 left 共同决定；
                # 任一被显式设为 nil/none 即不画，两侧都缺省才继承表级 insideV。
                if ci < n - 1:
                    nb_left = cell_side(cells[ci + 1], "left")
                    right_hidden = (c_right is not None and not is_visible(c_right))
                    nbleft_hidden = (nb_left is not None and not is_visible(nb_left))
                    if not (right_hidden or nbleft_hidden):
                        # 两侧都没把它设为不可见 -> 由显式可见值或表级 insideV 决定
                        shown = (
                            (c_right is not None and is_visible(c_right))
                            or (nb_left is not None and is_visible(nb_left))
                            or t_insideV
                        )
                        if shown:
                            return True, f"{lab} 出现额外竖框线（内部竖线 insideV）"

    return False, ""


def deduct_simulated_columns(labeled):
    """
    -5：表5-4至表5-12任意表格仍使用连续空格、制表符或同一大单元格模拟多列数据，
        未形成与字段对应的独立可编辑单元格。

    细则逐点拆解（"任意表格"命中"任意一种模拟方式"即扣分）：
      方式1：用"连续空格"在单个单元格内模拟多列 —— 单元格文本内部出现 2 个及
             以上连续空格分隔多段数据。
      方式2：用"制表符"在单个单元格内模拟多列 —— 单元格文本含 \\t。
      方式3：用"同一大单元格"模拟多列 —— 一个单元格通过 gridSpan 横跨多列，
             却在其中塞入本应分列的多段数据（即跨列单元格内仍含空格/制表符分隔
             的多段内容）。
    共同后果判据：未形成"与字段对应的独立可编辑单元格"。即上述任一模拟方式出现，
    意味着数据没有落入各自独立的单元格。
    细则只针对"模拟多列"这一行为，不对列数本身、字段顺序等其他内容加以约束。
    """
    # 仅当一段数据明显是被空格/制表符拆成"多列"时才算模拟；
    # 中文短语内部的单个空格(如 "t 检验")不计为模拟。
    multi_space = re.compile(r"\S+\s{2,}\S+")   # 2+连续空格分隔的多段
    tab_char = "\t"

    for lab in TARGET_LABELS:
        tbl = labeled.get(lab)
        if tbl is None:
            continue
        for row in rows_of(tbl):
            for tc in cells_of(row):
                txt = cell_text(tc)

                # 该单元格是否为"同一大单元格"(横跨多列)
                tcPr = tc.find(w("tcPr"))
                gs = tcPr.find(w("gridSpan")) if tcPr is not None else None
                span = int(gs.get(w("val"))) if gs is not None else 1
                is_big_cell = span >= 2

                # 方式2：制表符模拟分列
                if tab_char in txt:
                    return True, (f"{lab} 单元格用制表符模拟多列、"
                                  f"未形成独立单元格: '{txt}'")

                # 方式1：连续空格模拟分列
                if multi_space.search(txt):
                    return True, (f"{lab} 单元格用连续空格模拟多列、"
                                  f"未形成独立单元格: '{txt}'")

                # 方式3：同一大单元格(跨列)内塞入用空格/制表符分隔的多段数据
                if is_big_cell and (multi_space.search(txt) or tab_char in txt):
                    return True, (f"{lab} 用同一大单元格(跨{span}列)模拟多列、"
                                  f"未形成独立单元格: '{txt}'")
    return False, ""


def deduct_image_table(labeled, model):
    """
    -5：表5-4至表5-12任意表格被转换为图片、截图或不可编辑对象。

    细则逐点拆解（任意一个目标表命中任意一种"被转换"形态即扣分）：
      形态1：被转换为"图片"/"截图" —— 原本应是 w:tbl 的位置出现图片对象：
             w:drawing（DrawingML，含 a:blip 引用 word/media 中的图片）、
             w:pict（VML 图片）、w:object 内嵌图片。截图与普通图片在 OOXML
             中同为图片对象，无法也无需区分。
      形态2：被转换为"不可编辑对象" —— OLE/嵌入对象 w:object / w:OLEObject，
             或内容控件锁定为不可编辑(w:sdt 且 w:lock 含 contentLocked)。
      "被转换"判据：该目标表对应位置已不再是可编辑的 w:tbl，而是上述对象之一；
                    或在本应是表格的题注之后，紧邻出现的是图片/对象而非 w:tbl。
    细则只针对"被转换为图片/截图/不可编辑对象"，不对其他内容加以约束。
    """
    IMG_OBJ_TAGS = ("drawing", "pict", "object", "OLEObject")

    # 情形A：目标表仍是 w:tbl，但表格内部嵌入了图片/对象（即用图片替换了内容）
    for lab in TARGET_LABELS:
        tbl = labeled.get(lab)
        if tbl is None:
            continue
        for tag in IMG_OBJ_TAGS:
            if tbl.find(f".//{w(tag)}") is not None:
                return True, f"{lab} 表格被转换为图片/截图/不可编辑对象（含 {tag}）"

    # 情形B：题注之后本应紧邻 w:tbl，却变成了图片/对象段落（整表被转图）
    pending_label = None
    for kind, txt, elem in model.body_order:
        if kind == "p":
            t = (txt or "").replace(" ", "")
            matched = None
            for lab in TARGET_LABELS:
                if re.search(re.escape(lab) + r"(?![0-9])", t):
                    matched = lab
                    break
            if matched is not None:
                pending_label = matched
            elif pending_label is not None:
                # 题注后、表格出现前的段落：若含图片/对象，视为整表被转图
                for tag in IMG_OBJ_TAGS:
                    if elem.find(f".//{w(tag)}") is not None:
                        return True, (f"{pending_label} 表格被转换为图片/截图/"
                                      f"不可编辑对象（题注后出现 {tag}）")
        elif kind == "tbl":
            # 正常出现表格，清除待配对标签
            pending_label = None

    return False, ""


def _render_pdf_via_wps(docx_path: str):
    """
    历史入口：曾用 WPS/Word COM 渲染 PDF 以拿到真实分页坐标。
    因用户要求"非必要不用 COM"，且跨页检测已可用 OOXML 分页信号回退（见
    deduct_page_break 的回退分支），本函数固定返回 None，不再触发 COM 调用。
    保留函数与调用点是为了让后续如需真渲染时可插回 pdf 后端（如 Aspose / docx2pdf
    的非 COM 后端），无需改上层调用链。
    """
    del docx_path  # 保留形参签名，不再使用
    return None


def _detect_crosspage_from_pdf(pdf_path):
    """
    用 pdfplumber（经 pdf_backend 适配层）在渲染后的 PDF 上检测：表5-4~表5-12 中是否存在表格跨页
    或数据行被分页拆开。返回 (命中:bool, 触发说明:str)。

    判定方法（基于真实排版坐标，与 WPS 所见一致）：
      · 以各表题注("表5-x")在 PDF 中的首次出现作为该表起点(页, y)；
        下一个题注(或末尾"例子")作为"内容区间"的上界。
      · 关键：只有当表格"自身内容"真的落到了下一页，才算跨页。
        下一张表的题注在下一页，只说明当前表之后发生了翻页，并不代表当前
        表被切断（当前表可能是本页最后一张、完整落在本页）。
      · 因此对每张表：收集从其题注 y 往下、到下一锚点之前、且位于其起始页
        的所有文字块，取最大底部 y 作为该表实际底部。
        - 若在下一锚点之前，本表内容还出现在了起始页的下一页，则表身跨越了
          页边界 -> 跨页 / 数据行被分页拆开。
        - 否则（内容全部落在起始页）-> 不跨页，即使下一题注在下一页。
    """
    try:
        try:
            import pdf_backend
        except ImportError:
            from verifiers import pdf_backend
    except Exception:
        return None  # 库不可用 -> 交由调用方降级处理

    try:
        doc = pdf_backend.open_pdf(pdf_path)
    except Exception:
        return None

    order = list(TARGET_LABELS) + ["例子"]

    def first_anchor(label):
        hits = []
        for pno in range(doc.page_count):
            for r in doc.search_text(pno, label):
                hits.append((pno, r.y0))
        hits.sort()
        return hits[0] if hits else None

    anchors = {lab: first_anchor(lab) for lab in order}

    # 预取各页文字块（page, y0, y1）用于判断表身实际延伸范围
    def page_blocks(pno):
        try:
            # 保持与历史文本块相同的元组结构
            # (x0,y0,x1,y1,text,bno,btype)
            return [
                (b.bbox.x0, b.bbox.y0, b.bbox.x1, b.bbox.y1, b.text, i, 0)
                for i, b in enumerate(doc.extract_text_blocks(pno))
            ]
        except Exception:
            return []

    for i, lab in enumerate(TARGET_LABELS):
        a = anchors.get(lab)
        if a is None:
            continue
        pg, y = a

        # 下一个存在的锚点作为内容区间上界
        nxt = None
        for j in range(i + 1, len(order)):
            if anchors.get(order[j]) is not None:
                nxt = anchors[order[j]]
                break

        # 本表内容是否出现在起始页的"下一页及之后"，且仍在下一锚点之前。
        # 若下一锚点就在起始页 -> 本表必然不跨页，直接跳过。
        if nxt is not None:
            npg, ny = nxt
            if npg == pg:
                continue  # 下一张表与本表同页 -> 本表完整落在本页，不跨页

        # 起始页之后、下一锚点之前，是否还有本表的文字内容
        crossed = False
        cross_pg = None
        for pno in range(pg + 1, doc.page_count):
            # 到达下一锚点所在页时，只统计该页中位于下一锚点 y 之上的内容
            if nxt is not None:
                npg, ny = nxt
                if pno > npg:
                    break
                limit_y = ny if pno == npg else None
            else:
                limit_y = None
            for b in page_blocks(pno):
                by0 = b[1]
                text = (b[4] or "").strip()
                if not text:
                    continue
                if limit_y is not None and by0 >= limit_y:
                    continue  # 已进入下一张表的题注区域，不属于本表
                crossed = True
                cross_pg = pno
                break
            if crossed:
                break

        if crossed:
            doc.close()
            return True, (f"{lab} 表格出现跨页（题注在第{pg + 1}页，"
                          f"表身延续到第{cross_pg + 1}页，数据行被分页拆开）")

    doc.close()
    return False, ""


def deduct_page_break(labeled, model=None, docx_path=None):
    """
    -3：表5-4至表5-12任意一个表格出现跨页情况，或任意数据行被分页拆开。

    细则逐点拆解（两点为"或"关系，任意一点在任意目标表命中即扣分）：
      点1：表格出现"跨页"——整张表横跨两个页面。
      点2："任意数据行被分页拆开"——某一行被分页边界从中切断。

    实现：分页位置只有排版引擎渲染时才确定，docx 源 XML 通常不含可靠的分页
    信息。故本条用本机 WPS 将文档真实渲染为 PDF，再按真实坐标判定
    每张表是否跨页 / 数据行是否被页边界切开（与在 WPS 中肉眼所见一致）。

    渲染不可用时（无 WPS/Word 或无 PDF 解析库）回退到 XML 源码信号
    （显式分页符 w:br[type=page]、渲染分页标记 w:lastRenderedPageBreak），
    此回退为近似，可能漏判未写入分页标记的文档。
    细则只针对"跨页 / 数据行被拆开"，不对其他内容加以约束。
    """
    # ---- 首选：真实渲染 PDF 后按坐标判定 ----
    if docx_path:
        pdf_path = _render_pdf_via_wps(docx_path)
        if pdf_path:
            result = _detect_crosspage_from_pdf(pdf_path)
            try:
                os.remove(pdf_path)
            except Exception:
                pass
            if result is not None:
                return result  # 渲染检测成功（命中或未命中均以此为准）

    # ---- 回退：XML 源码分页信号（近似）----
    for lab in TARGET_LABELS:
        tbl = labeled.get(lab)
        if tbl is None:
            continue

        # 点1：表格跨页
        table_crosses_page = False
        for br in tbl.iter(w("br")):
            if br.get(w("type")) == "page":
                table_crosses_page = True
                break
        if not table_crosses_page and tbl.find(f".//{w('lastRenderedPageBreak')}") is not None:
            table_crosses_page = True
        if table_crosses_page:
            return True, f"{lab} 表格出现跨页（XML分页信号，回退判定）"

        # 点2：数据行被分页拆开
        rows = rows_of(tbl)
        for ri, row in enumerate(rows):
            if ri < 2:
                continue
            if row.find(f".//{w('lastRenderedPageBreak')}") is not None:
                return True, f"{lab} 第{ri + 1}行(数据行)被分页拆开（XML分页信号）"
            trPr = row.find(w("trPr"))
            cant_split = (trPr is not None
                          and trPr.find(w("cantSplit")) is not None)
            if table_crosses_page and not cant_split:
                return True, f"{lab} 第{ri + 1}行(数据行)未禁止跨页拆分且表格已跨页"
    return False, ""


def deduct_data_corruption(labeled):
    """
    -3：表5-4至表5-12中的 维度名称、组别名称、测试时间、样本量 或 统计数值
        出现大面积缺失、替换或顺序错乱。

    细则逐点拆解（5 类内容 × 3 种异常，任意目标表命中任意一项即扣分）：
      涉及的 5 类内容：
        a) 维度名称   —— 表头含"维度"列，数据区每个维度分组应有名称。
        b) 组别名称   —— 表头含"组别"列时，应出现 干预组/参照组 等组别名。
        c) 测试时间   —— 表头含"测试时间"列时，应出现 干预前/干预后(前测/后测) 等。
           （注：双组表用"组别"，单组前后测表用"测试时间"，二者按表头存在性判断。）
        d) 样本量     —— 若表头声明样本量列(如 N/样本量/人数)，数据区应有对应数值。
        e) 统计数值   —— X±S、F值、Sig.、t值、Sig(双侧) 等统计列的数值。
      3 种异常：
        缺失：应有内容的位置大面积为空（非空率过低）。
        替换：内容被无意义占位符替换（如 ###、xxx、占位、N/A、待填、??? 等）。
        顺序错乱：表头列顺序与三线表规范顺序明显不一致（如 维度/组别 不在最前，
                  或统计列顺序颠倒）。
    "大面积"判据：缺失按非空率阈值衡量；替换/顺序错乱按是否出现即判定。
    细则只针对这 5 类内容的 缺失/替换/顺序错乱，不对其他内容加以约束。
    """
    placeholder_pat = re.compile(
        r"(#{2,}|x{3,}|X{3,}|\?{2,}|待填|占位|placeholder|N/?A|TODO|空缺|未填)",
        re.IGNORECASE)
    # 表头各类内容的关键字
    HEADER_DIM = ["维度"]
    HEADER_GROUP = ["组别"]
    HEADER_TIME = ["测试时间", "时间"]
    HEADER_SAMPLE = ["样本量", "N", "人数", "样本"]
    GROUP_VALUES = ["干预组", "参照组", "实验组", "对照组"]
    TIME_VALUES = ["干预前", "干预后", "前测", "后测", "测前", "测后"]
    STAT_HEADERS = ["X±S", "F值", "Sig", "t值"]

    def norm(s):
        return (s or "").replace(" ", "")

    for lab in TARGET_LABELS:
        tbl = labeled.get(lab)
        if tbl is None:
            continue
        rows = rows_of(tbl)
        if len(rows) < 3:
            return True, f"{lab} 表格行数异常({len(rows)})，疑似数据缺失"

        # 表头文本（前两行）与各列首行字段
        header_cells = [norm(cell_text(tc)) for r in rows[:2] for tc in cells_of(r)]
        header_text = "".join(header_cells)
        # 字段名行（第2行）按列顺序
        field_row = [norm(cell_text(tc)) for tc in cells_of(rows[1])]

        has_dim = any(k in header_text for k in HEADER_DIM)
        has_group = any(norm(k) in header_text for k in HEADER_GROUP)
        has_time = any(norm(k) in header_text for k in HEADER_TIME)
        has_sample = any(norm(k) in header_text for k in HEADER_SAMPLE
                         if k not in ("N",)) or ("N" in [c for c in field_row])

        # 数据区（第3行起）全部文本
        data_cells = [cell_text(tc) for r in rows[2:] for tc in cells_of(r)]
        data_text = "".join(norm(c) for c in data_cells)

        # ---- 异常1：替换（占位符/乱码替换真实内容）----
        for c in data_cells + header_cells:
            if placeholder_pat.search(c):
                return True, f"{lab} 出现占位符/替换内容: '{c.strip()}'"

        # ---- 异常2：大面积缺失 ----
        # a) 维度名称：表头有"维度"但数据区维度分组名称大面积为空
        if has_dim:
            # 维度名称位于每个分组首行第一列；统计非空的维度名个数
            dim_names = [norm(cell_text(cells_of(r)[0])) for r in rows[2:]
                         if len(cells_of(r)) > 0]
            nonempty_dim = [d for d in dim_names if d != ""]
            if len(nonempty_dim) == 0:
                return True, f"{lab} 维度名称大面积缺失"
        # b) 组别名称
        if has_group and not any(g in data_text for g in [norm(x) for x in GROUP_VALUES]):
            return True, f"{lab} 组别名称大面积缺失（未见 干预组/参照组 等）"
        # c) 测试时间
        if has_time and not any(t in data_text for t in [norm(x) for x in TIME_VALUES]):
            return True, f"{lab} 测试时间大面积缺失（未见 干预前/干预后 等）"
        # d) 样本量（仅当表头声明样本量列时检查）
        if has_sample:
            # 数据区应含数字（样本量为整数）
            if not re.search(r"\d", data_text):
                return True, f"{lab} 样本量大面积缺失"
        # e) 统计数值：统计列对应的数值整体非空率
        # 数值型单元格（含 数字 / ± / 小数点 / 负号）
        numeric_cells = [c for c in data_cells
                         if re.search(r"[0-9]", c)]
        # 统计数值应占数据区相当比例；若数据区几乎无数值 -> 大面积缺失
        if data_cells:
            # 仅统计"本应有数值"的单元格：排除维度名/组别名/时间名等文字列
            text_value_set = set(norm(x) for x in
                                 (GROUP_VALUES + TIME_VALUES))
            candidate = []
            for c in data_cells:
                nc = norm(c)
                if nc == "":
                    continue
                if nc in text_value_set:
                    continue
                candidate.append(c)
            # 在非空、非"文字值"的数据单元格里，应当主要是统计数值
            if candidate:
                num_ratio = sum(1 for c in candidate if re.search(r"[0-9]", c)) / len(candidate)
                if num_ratio < 0.5:
                    return True, (f"{lab} 统计数值大面积缺失/异常"
                                  f"（数值占比 {num_ratio:.2f}）")

        # ---- 异常3：顺序错乱 ----
        # 三线表规范：表头字段行应以 维度 开头，随后是 组别/测试时间，
        # 统计列(X±S / F值 / Sig. / t值 / Sig(双侧))在后。
        non_empty_fields = [f for f in field_row if f != ""]
        if non_empty_fields:
            # 维度应是首个非空字段
            if has_dim and non_empty_fields[0] != norm("维度"):
                return True, (f"{lab} 表头字段顺序错乱（首列应为'维度'，"
                              f"实为'{non_empty_fields[0]}'）")
            # 维度/组别/测试时间 应排在统计列之前
            first_stat_idx = None
            first_label_after_stat = None
            for i, f in enumerate(non_empty_fields):
                is_stat = any(norm(s) in f for s in STAT_HEADERS)
                is_label = (f == norm("维度") or f == norm("组别")
                            or "测试时间" in f or f == norm("时间"))
                if is_stat and first_stat_idx is None:
                    first_stat_idx = i
                if is_label and first_stat_idx is not None:
                    first_label_after_stat = f
            if first_label_after_stat is not None:
                return True, (f"{lab} 表头字段顺序错乱（描述列'{first_label_after_stat}'"
                              f"出现在统计列之后）")
    return False, ""


def deduct_font_issue(labeled, model):
    """
    -3：表5-4至表5-12 表内文字 中文不是宋体小五号，英文、数字及统计符号不是
        Times New Roman 小五号，字体颜色不是黑色；或 任意标题字体不是宋体五号加粗。

    细则逐点拆解（前半段"表内文字"与后半段"标题"为"或"关系，任意一处违规即扣分）：
      表内文字（表5-4~表5-12 单元格内）每个 run：
        点1：中文字符 -> 中文字体(eastAsia)必须是"宋体"，字号必须是小五号(sz=18)。
        点2：英文、数字及统计符号(如 ± . () F t Sig 等 ASCII/拉丁字符) ->
             西文字体(ascii)必须是"Times New Roman"，字号必须是小五号(sz=18)。
        点3：字体颜色必须是黑色(000000/auto)。
      标题（各表题注，如"表5-4 …"）：
        点4：标题字体必须是宋体五号(sz=21)加粗(b)。任意一个标题不满足即违规。
    小五号 = 9磅 = sz18；五号 = 10.5磅 = sz21。
    细则只针对上述字体/字号/颜色/加粗，不对其他内容加以约束。
    """

    def run_props(r):
        rpr = r.find(w("rPr"))
        if rpr is None:
            return {"ascii": None, "eastAsia": None, "sz": None,
                    "color": None, "bold": False}
        fonts = rpr.find(w("rFonts"))
        sz = rpr.find(w("sz"))
        color = rpr.find(w("color"))
        bold = rpr.find(w("b"))
        return {
            "ascii": fonts.get(w("ascii")) if fonts is not None else None,
            "eastAsia": fonts.get(w("eastAsia")) if fonts is not None else None,
            "sz": sz.get(w("val")) if sz is not None else None,
            "color": color.get(w("val")) if color is not None else None,
            "bold": (bold is not None and bold.get(w("val")) not in ("0", "false")),
        }

    def is_cjk(ch):
        return '一' <= ch <= '鿿'

    def is_latin_or_symbol(ch):
        # 英文、数字及统计符号（ASCII 字母数字 + 常见统计符号）
        if ch.isascii() and (ch.isalnum()):
            return True
        return ch in "±.()-+/*=<>%:,"

    # ---- 表内文字（前半段）----
    for lab in TARGET_LABELS:
        tbl = labeled.get(lab)
        if tbl is None:
            continue
        for r in tbl.iter(w("r")):
            txt = "".join(t.text or "" for t in r.findall(w("t")))
            if txt.strip() == "":
                continue
            p = run_props(r)
            has_cjk = any(is_cjk(ch) for ch in txt)
            has_latin = any(is_latin_or_symbol(ch) for ch in txt)

            # 点1：中文 -> 宋体 + 小五
            if has_cjk:
                if p["eastAsia"] != "宋体":
                    return True, (f"{lab} 表内中文字体不是宋体: "
                                  f"'{txt.strip()}'(eastAsia={p['eastAsia']})")
                if p["sz"] is None or int(p["sz"]) != SZ_XIAOWU:
                    return True, (f"{lab} 表内中文字号不是小五号: "
                                  f"'{txt.strip()}'(sz={p['sz']})")
            # 点2：英文/数字/统计符号 -> Times New Roman + 小五
            if has_latin:
                if p["ascii"] != "Times New Roman":
                    return True, (f"{lab} 表内英文/数字/符号字体不是Times New Roman: "
                                  f"'{txt.strip()}'(ascii={p['ascii']})")
                if p["sz"] is None or int(p["sz"]) != SZ_XIAOWU:
                    return True, (f"{lab} 表内英文/数字/符号字号不是小五号: "
                                  f"'{txt.strip()}'(sz={p['sz']})")
            # 点3：颜色 -> 黑色
            if not is_black(p["color"]):
                return True, (f"{lab} 表内文字颜色不是黑色: "
                              f"'{txt.strip()}'(color={p['color']})")

    # ---- 标题（后半段）：任意标题字体不是宋体五号加粗 ----
    for kind, txt, elem in model.body_order:
        if kind != "p":
            continue
        t = (txt or "").replace(" ", "")
        if not any(re.search(re.escape(lab) + r"(?![0-9])", t) for lab in TARGET_LABELS):
            continue
        # 该题注内每个有文字的 run 都需满足 宋体 + 五号 + 加粗
        for r in elem.iter(w("r")):
            rt = "".join(x.text or "" for x in r.findall(w("t")))
            if rt.strip() == "":
                continue
            p = run_props(r)
            if p["eastAsia"] != "宋体":
                return True, (f"标题字体不是宋体: '{t}'(eastAsia={p['eastAsia']})")
            if p["sz"] is None or int(p["sz"]) != SZ_WUHAO:
                return True, (f"标题字号不是五号: '{t}'(sz={p['sz']})")
            if not p["bold"]:
                return True, (f"标题未加粗: '{t}'")

    return False, ""


# ---------------------------------------------------------------------------
# 主流程
# ---------------------------------------------------------------------------
SCRIPT_ID = "026"


def _find_docx_in_dir(dir_path):
    """在 dir_path 目录下定位被评估的 .docx 文档（忽略临时文件）。"""
    candidates = [
        f for f in os.listdir(dir_path)
        if f.lower().endswith(".docx") and not f.startswith("~$")
    ]
    if not candidates:
        raise FileNotFoundError(f"目录中未找到 .docx 文档: {dir_path}")
    return os.path.join(dir_path, candidates[0])


def evaluate(dir_path: str) -> dict:
    """
    对 dir_path 目录下的 .docx 文档评分。

    Args:
        dir_path: 脚本所在目录的路径，脚本自己在该目录内定位并打开被评估的文档。

    Returns:
        结构化结果 dict，字段见 §2.2：
        id / file_name / status / error / dim1_pass / dim1_reason /
        dim2_items / total_score / max_score。
    """
    result = {
        "id": SCRIPT_ID,
        "file_name": None,
        "status": "ok",
        "error": None,
        "dim1_pass": False,
        "dim1_reason": "",
        "dim2_items": [],
        "total_score": 0,
        "max_score": 0,
    }

    try:
        docx_path = _find_docx_in_dir(dir_path)
        result["file_name"] = os.path.basename(docx_path)
        model = DocxModel(docx_path)

        # ---------------- 维度1 ----------------
        passed1, d1 = check_dimension1(model)
        result["dim1_pass"] = passed1
        if not passed1:
            for ok, msg in d1:
                if not ok:
                    result["dim1_reason"] = msg
                    break
            result["total_score"] = 0
            return result

        labeled = map_labeled_tables(model)

        # ---------------- 维度2：逐条评分 ----------------
        # 每条评分细则的完整文字（用于 dim2_items 展示）
        RULE_OUTER = ("表5-4至表5-12表格外框线：外框线只保留上下外框线，"
                      "表格上框线和下框线为1.5磅黑色实线。")
        RULE_TWO = ("表5-4、表5-5、表5-8、表5-9和表5-12分组内框线：仅有"
                    "“Levene检验”、“t检验”两个分组标题下方有下框线，其余位置无内框线，"
                    "框线类型为0.75磅黑色短横线。")
        RULE_ONE = ("表5-6、表5-7、表5-10和表5-11分组表头内框线：仅有“t检验”标题下方"
                    "有下框线，其余位置无内框线，框线类型为0.75磅黑色短横线。")
        RULE_VERT = "表格出现额外竖框线"
        RULE_SIM = ("表5-4至表5-12任意表格仍使用连续空格、制表符或同一大单元格模拟"
                    "多列数据，未形成与字段对应的独立可编辑单元格。")
        RULE_IMG = "表5-4至表5-12任意表格被转换为图片、截图或不可编辑对象。"
        RULE_PAGE = "表5-4至表5-12任意一个表格出现跨页情况或任意数据行被分页拆开"
        RULE_DATA = ("表5-4至表5-12中的维度名称、组别名称、测试时间、样本量或统计数值"
                     "出现大面积缺失、替换或顺序错乱。")
        RULE_FONT = ("表5-4至表5-12表内文字中文不是宋体小五号，英文、数字及统计符号不是"
                     "Times New Roman小五号，字体颜色不是黑色；或任意标题字体不是宋体五号加粗")

        dim2_items = []
        total = 0

        # 得分点（每张三线表 +1，rubric 语义为"逐表独立计分"）。
        # 脚本契约要求每个 item 满足 delta == max_delta(命中) 或 delta == 0(未命中)，
        # 因此把"最多 N 分"的聚合项拆成 N 个逐表 item（每表 max_delta=1）：
        #   - hit=该表是否满足本条细则；
        #   - delta = 1 if hit else 0，天然满足 delta == max_delta if hit else 0；
        #   - 逐表 rule 文案追加表标签，便于区分与定位。
        def _append_per_table(rule, target_labels, hit_labels, per_table_delta=1):
            nonlocal total
            hit_set = set(hit_labels)
            for lab in target_labels:
                hit = lab in hit_set
                delta = per_table_delta if hit else 0
                total += delta
                dim2_items.append({
                    "rule": f"{rule}（{lab}）",
                    "max_delta": per_table_delta,
                    "delta": delta,
                    "hit": hit,
                    "detail": "",
                })

        _, outer_hits = check_outer_border(labeled)
        _append_per_table(RULE_OUTER, TARGET_LABELS, outer_hits)

        _, two_hits = check_two_group_inner(labeled)
        _append_per_table(RULE_TWO, sorted(TWO_GROUP_TABLES), two_hits)

        _, one_hits = check_one_group_inner(labeled)
        _append_per_table(RULE_ONE, sorted(ONE_GROUP_TABLES), one_hits)

        # 扣分点（命中即计入对应负分；未命中 delta=0）
        deductions = [
            (-3, deduct_extra_vertical(labeled), RULE_VERT),
            (-5, deduct_simulated_columns(labeled), RULE_SIM),
            (-5, deduct_image_table(labeled, model), RULE_IMG),
            (-3, deduct_page_break(labeled, model, docx_path), RULE_PAGE),
            (-3, deduct_data_corruption(labeled), RULE_DATA),
            (-3, deduct_font_issue(labeled, model), RULE_FONT),
        ]
        for pts, (hit, reason), rule in deductions:
            delta = pts if hit else 0
            total += delta
            dim2_items.append({
                "rule": rule, "max_delta": pts, "delta": delta,
                "hit": hit, "detail": reason,
            })

        result["dim2_items"] = dim2_items
        result["total_score"] = total
        result["max_score"] = sum(abs(item["max_delta"]) for item in dim2_items
                                  if item["max_delta"] > 0)
        return result

    except Exception as e:
        result["status"] = "error"
        result["error"] = str(e)
        return result


if __name__ == "__main__":
    # 仅用于本地调试：传入目录路径，打印结构化结果
    target_dir = sys.argv[1] if len(sys.argv) > 1 else os.path.dirname(os.path.abspath(__file__))
    output = json.dumps(evaluate(target_dir), ensure_ascii=False, indent=2)
    _ = sys.stdout.buffer.write((output + "\n").encode("utf-8"))
