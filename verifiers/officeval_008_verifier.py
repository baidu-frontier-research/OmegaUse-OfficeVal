#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
文献翻译译文（Word 排版）自动评估脚本

评估逻辑：
  维度一：可用与可修改性 → 任一条不满足直接判 0 分，不再检查维度二。
  维度二：完成度 → 各加分项累计（39 条 / 满分 127 分）。
"""

import os
import sys
import json
import zipfile
import re
from docx import Document
from docx.oxml.ns import qn

# 本脚本编号（用于统一返回结果中的 "id" 字段）
SCRIPT_ID = "008"

# ============================================================
# 常量
# ============================================================

EMU_PER_CM = 360000.0
TWIP_PER_CM = 567.0
A4_TWIPS = (11906, 16838)        # A4 纵向：宽 21cm、高 29.7cm
EXACT_LINE = 550                 # 固定行距 27.5 磅 = 550 twips
SIMSUN, SIMHEI, TNR = "宋体", "黑体", "Times New Roman"
M_NS = "{http://schemas.openxmlformats.org/officeDocument/2006/math}"
A_NS = "{http://schemas.openxmlformats.org/drawingml/2006/main}"

# 字体名归一（同一字体的不同书写形式视为同一字体）
_FONT_ALIAS = {
    "simsun": SIMSUN, "宋体": SIMSUN, "nsimsun": SIMSUN, "宋体-简": SIMSUN,
    "simhei": SIMHEI, "黑体": SIMHEI, "黑体-简": SIMHEI,
    "times new roman": TNR, "timesnewroman": TNR, "times": TNR,
}

# Unicode 形似角标字符（要求使用真实上下标而非这些字符）
UNI_SUPSUB = "₀₁₂₃₄₅₆₇₈₉₊₋₍₎ⁿ⁰¹²³⁴⁵⁶⁷⁸⁹⁺⁻⁽⁾"

CJK_RE = re.compile(r"[一-鿿]")
ALNUM_RE = re.compile(r"[A-Za-z0-9]")

# 正文文献引用标注：方括号 + 序号（可含区间与并列）
CIT_RE = re.compile(r"\[\s*\d{1,3}\s*(?:[-–,，]\s*\d{1,3}\s*)*\]")

# 单位与量符号的指数被写成下划线 / 连字符 / 普通文字（应为真实上标）
UP_BAD_RE = re.compile(
    r"(?:mg|kg|g|mol|L|min|cm|nm|K|h|J)\s*[·・.]?\s*[_\-−]\s*\d"
    r"|(?<![A-Za-z0-9])[RrΔ∆][A-Za-z]?\s*[02](?![0-9A-Za-z])")

# 参考文献条目内的空格缺失
GLUE_RE = re.compile(r"[A-Za-z][.,;:][A-Za-z0-9]")
LONG_RE = re.compile(r"[A-Za-z][A-Za-z.,;:\-]{24,}")

# 中文段落里出现的半角标点（充当中文标点）
HALF_PUNCT_RE = re.compile(r"[一-鿿][,;:?!]|[,;:?!][一-鿿]")

# 章节标题编号形态
H1_RE = re.compile(r"^[1-9]\s+\S")
H2_RE = re.compile(r"^(\d+\.\d+)(?!\.)\s+\S")
H3_RE = re.compile(r"^(\d+\.\d+\.\d+)\s+\S")
FIG_RE = re.compile(r"^图\s*(\d{1,3})")
TAB_RE = re.compile(r"^表\s*(\d{1,3})")
REF_ITEM_RE = re.compile(r"^\s*(?:\[\s*(\d{1,3})\s*\]|(\d{1,3})[\.、]\s)")
FORMULA_NO_RE = re.compile(r"\(\s*(\d{1,3})\s*\)\s*$")

# 期望结构（规则原文给定的绝对值）
H2_EXPECT = ["2.%d" % i for i in range(1, 11)] + ["3.%d" % i for i in range(1, 7)]
H3_EXPECT = ["2.1.1", "2.1.2", "2.3.1", "2.3.2", "2.3.3", "2.3.4"]
TABLE_DIMS = [(6, 4), (5, 5), (4, 4), (6, 5), (7, 4), (6, 4)]
HEADER_TEXT = "华语青月大学本科生毕业设计（论文）用纸"
FIG_COUNT = 11
TAB_COUNT = 6
FORMULA_COUNT = 8
REF_COUNT = 40
KEYWORD_COUNT = 6

# ============================================================
# 基础工具函数
# ============================================================


def _f(v, default=None):
    """安全转 float"""
    try:
        return float(v)
    except Exception:
        return default


def emu_to_cm(v):
    try:
        return round(float(v) / EMU_PER_CM, 3)
    except Exception:
        return None


def twip_to_cm(v):
    try:
        return round(float(v) / TWIP_PER_CM, 3)
    except Exception:
        return None


def emu_to_twip(v):
    try:
        return int(round(float(v) / 635.0))
    except Exception:
        return None


def norm_font(name):
    """字体名归一：宋体≡SimSun、黑体≡SimHei、Times New Roman 多种写法视为同一字体"""
    if not name:
        return None
    key = str(name).strip()
    return _FONT_ALIAS.get(key.lower(), key)


def on_off(el):
    """开关型属性：标签存在且 val 非 0/false/off 才算开启（val="0" 表示关闭）"""
    if el is None:
        return None
    v = el.get(qn("w:val"))
    if v is None:
        return True
    return str(v).lower() not in ("0", "false", "off")

def has_cjk(text):
    return bool(CJK_RE.search(text or ""))


def has_alnum(text):
    return bool(ALNUM_RE.search(text or ""))


def align_ok(pf, want):
    """对齐方式判定：CENTER 居中、JUSTIFY 两端对齐、LEFT 左对齐（含未设置）"""
    v = (pf.get("align") or "").lower()
    if want == "CENTER":
        return v == "center"
    if want == "JUSTIFY":
        return v in ("both", "justify", "distribute")
    if want == "LEFT":
        return v in ("", "left", "start")
    return True


def spacing_ok(pf, rule, twips=None):
    """行距判定：rule="exact" 固定值；rule="auto" 单倍行距"""
    lr = (pf.get("lineRule") or "auto").lower()
    ln = _f(pf.get("line"))
    if rule == "exact":
        return lr == "exact" and ln is not None and abs(ln - (twips or EXACT_LINE)) <= 2
    if lr != "auto":
        return False
    return ln is None or abs(ln - 240) <= 2


def first_line_chars(pf, size_pt):
    """首行缩进（字符）：优先字符单位，其次磅值折算；悬挂缩进折算为负值"""
    base = size_pt or 12.0
    if pf.get("hang_chars") is not None:
        return -_f(pf["hang_chars"], 0.0) / 100.0
    if pf.get("hang") is not None:
        return -_f(pf["hang"], 0.0) / 20.0 / base
    if pf.get("fl_chars") is not None:
        return _f(pf["fl_chars"], 0.0) / 100.0
    if pf.get("fl") is not None:
        return _f(pf["fl"], 0.0) / 20.0 / base
    return 0.0

def hanging_chars(pf, size_pt):
    """悬挂缩进（字符）"""
    base = size_pt or 12.0
    if pf.get("hang_chars") is not None:
        return _f(pf["hang_chars"], 0.0) / 100.0
    if pf.get("hang") is not None:
        return _f(pf["hang"], 0.0) / 20.0 / base
    return 0.0


def left_chars(pf, size_pt):
    """左缩进（字符）"""
    base = size_pt or 12.0
    if pf.get("left_chars") is not None:
        return _f(pf["left_chars"], 0.0) / 100.0
    if pf.get("left") is not None:
        return _f(pf["left"], 0.0) / 20.0 / base
    return 0.0


def ref_item_no(text):
    """参考文献条目序号：支持"[1]xxx"与"1. xxx"两种写法"""
    m = REF_ITEM_RE.match(text or "")
    if not m:
        return None
    return int(m.group(1) or m.group(2))


def near(a, b, tol=0.05):
    if a is None or b is None:
        return False
    return abs(a - b) <= tol


# ============================================================
# 格式解析器（字体 / 段落属性的完整继承链）
#   解析顺序：docDefaults → 段落样式 basedOn 链 → 字符样式 rStyle 链 → 直接格式
#   主题字体（*Theme）需到 theme1.xml 换算，中文取 minorFont/majorFont 的 script="Hans"
# ============================================================


class FormatResolver(object):

    def __init__(self, doc, docx_path):
        self.styles = {}
        self.default_pstyle = None
        self.dd_rpr = None
        self.dd_ppr = None
        self.theme = {}
        self._load_styles(doc)
        self._load_theme(docx_path)

    def _load_styles(self, doc):
        try:
            root = doc.styles.element
        except Exception:
            return
        for st in root.findall(qn("w:style")):
            sid = st.get(qn("w:styleId"))
            if sid:
                self.styles[sid] = st
            if (st.get(qn("w:type")) == "paragraph"
                    and st.get(qn("w:default")) in ("1", "true", "on")
                    and self.default_pstyle is None):
                self.default_pstyle = sid
        dd = root.find(qn("w:docDefaults"))
        if dd is not None:
            rd = dd.find(qn("w:rPrDefault"))
            if rd is not None:
                self.dd_rpr = rd.find(qn("w:rPr"))
            pd = dd.find(qn("w:pPrDefault"))
            if pd is not None:
                self.dd_ppr = pd.find(qn("w:pPr"))

    def _load_theme(self, docx_path):
        try:
            with zipfile.ZipFile(docx_path) as z:
                data = z.read("word/theme/theme1.xml")
        except Exception:
            return
        try:
            from lxml import etree
            root = etree.fromstring(data)
        except Exception:
            return
        for tag, key in (("majorFont", "major"), ("minorFont", "minor")):
            fs = root.find(".//" + A_NS + tag)
            if fs is None:
                continue
            latin = fs.find(A_NS + "latin")
            if latin is not None and latin.get("typeface"):
                self.theme[key + "Latin"] = latin.get("typeface")
            ea = fs.find(A_NS + "ea")
            if ea is not None and ea.get("typeface"):
                self.theme[key + "EastAsia"] = ea.get("typeface")
            for fo in fs.findall(A_NS + "font"):
                if fo.get("script") == "Hans" and fo.get("typeface"):
                    self.theme[key + "Hans"] = fo.get("typeface")

    def theme_font(self, theme_val):
        """把 minorEastAsia / minorHAnsi 等主题引用换算为真实字体名"""
        if not theme_val:
            return None
        kind = "major" if str(theme_val).startswith("major") else "minor"
        if "EastAsia" in theme_val:
            return self.theme.get(kind + "Hans") or self.theme.get(kind + "EastAsia")
        return self.theme.get(kind + "Latin")

    def _style_chain(self, sid):
        """样式及其 basedOn 链，根样式在前（便于由低到高逐层覆盖）"""
        chain, seen = [], set()
        while sid and sid in self.styles and sid not in seen:
            seen.add(sid)
            st = self.styles[sid]
            chain.append(st)
            b = st.find(qn("w:basedOn"))
            sid = b.get(qn("w:val")) if b is not None else None
        chain.reverse()
        return chain

    def _para_style_id(self, para):
        pPr = para._p.find(qn("w:pPr"))
        if pPr is not None:
            ps = pPr.find(qn("w:pStyle"))
            if ps is not None and ps.get(qn("w:val")):
                return ps.get(qn("w:val"))
        return self.default_pstyle

    def rpr_chain(self, para, run):
        """run 生效的 rPr 序列，由低到高"""
        out = []
        if self.dd_rpr is not None:
            out.append(self.dd_rpr)
        for st in self._style_chain(self._para_style_id(para)):
            r = st.find(qn("w:rPr"))
            if r is not None:
                out.append(r)
        rPr = run._element.find(qn("w:rPr"))
        if rPr is not None:
            rs = rPr.find(qn("w:rStyle"))
            if rs is not None:
                for st in self._style_chain(rs.get(qn("w:val"))):
                    r = st.find(qn("w:rPr"))
                    if r is not None:
                        out.append(r)
            out.append(rPr)
        return out

    def font(self, para, run):
        """返回 (中文字体, 西文字体, 字号磅, 是否加粗)"""
        ea = asc = None
        size = None
        bold = None
        for rpr in self.rpr_chain(para, run):
            rf = rpr.find(qn("w:rFonts"))
            if rf is not None:
                v = rf.get(qn("w:eastAsia"))
                if v:
                    ea = v
                else:
                    t = rf.get(qn("w:eastAsiaTheme"))
                    if t:
                        ea = self.theme_font(t) or ea
                v = rf.get(qn("w:ascii"))
                if v:
                    asc = v
                else:
                    t = rf.get(qn("w:asciiTheme"))
                    if t:
                        asc = self.theme_font(t) or asc
            sz = rpr.find(qn("w:sz"))
            if sz is not None and sz.get(qn("w:val")):
                v = _f(sz.get(qn("w:val")))
                if v is not None:
                    size = v / 2.0
            b = rpr.find(qn("w:b"))
            if b is not None:
                bold = on_off(b)
        return norm_font(ea), norm_font(asc), size, bool(bold)

    def vert_align(self, para, run):
        """返回 superscript / subscript / baseline / None"""
        va = None
        for rpr in self.rpr_chain(para, run):
            v = rpr.find(qn("w:vertAlign"))
            if v is not None and v.get(qn("w:val")):
                va = v.get(qn("w:val"))
        return va

    def highlighted(self, para, run):
        hl = None
        for rpr in self.rpr_chain(para, run):
            v = rpr.find(qn("w:highlight"))
            if v is not None:
                hl = v.get(qn("w:val"))
        return hl not in (None, "none")

    def pf(self, para):
        """段落属性：对齐、行距、缩进（字符单位优先，其次 twips）"""
        out = {"align": None, "lineRule": None, "line": None,
               "fl_chars": None, "fl": None, "left_chars": None, "left": None,
               "hang_chars": None, "hang": None}
        chain = []
        if self.dd_ppr is not None:
            chain.append(self.dd_ppr)
        for st in self._style_chain(self._para_style_id(para)):
            p = st.find(qn("w:pPr"))
            if p is not None:
                chain.append(p)
        pPr = para._p.find(qn("w:pPr"))
        if pPr is not None:
            chain.append(pPr)
        for p in chain:
            jc = p.find(qn("w:jc"))
            if jc is not None and jc.get(qn("w:val")):
                out["align"] = jc.get(qn("w:val"))
            sp = p.find(qn("w:spacing"))
            if sp is not None:
                if sp.get(qn("w:line")):
                    out["line"] = sp.get(qn("w:line"))
                if sp.get(qn("w:lineRule")):
                    out["lineRule"] = sp.get(qn("w:lineRule"))
            ind = p.find(qn("w:ind"))
            if ind is not None:
                pairs = (("w:firstLineChars", "fl_chars"), ("w:firstLine", "fl"),
                         ("w:leftChars", "left_chars"), ("w:left", "left"),
                         ("w:hangingChars", "hang_chars"), ("w:hanging", "hang"))
                got_first = ind.get(qn("w:firstLineChars")) or ind.get(qn("w:firstLine"))
                got_hang = ind.get(qn("w:hangingChars")) or ind.get(qn("w:hanging"))
                if got_first:
                    out["hang_chars"] = out["hang"] = None
                if got_hang:
                    out["fl_chars"] = out["fl"] = None
                for attr, key in pairs:
                    v = ind.get(qn(attr))
                    if v is not None:
                        out[key] = v
        return out

# ============================================================
# 维度一：可用与可修改性（门槛项，不计分）
#   点1：文件为 .docx 且能正常打开、包结构完整、无损坏
#   点2：正文非空、可编辑（未加密、未只读、未启用编辑限制）
#   点3：内容为中文译文，汉字占多数
#   任一条不满足 → 完成度直接判 0 分
# ============================================================


def check_dimension_one(doc, docx_path):
    issues = []

    # 点1：包结构完整
    try:
        with zipfile.ZipFile(docx_path) as z:
            names = z.namelist()
            if "word/document.xml" not in names:
                issues.append("文档包内缺少 word/document.xml，文件可能已损坏")
            bad = z.testzip()
            if bad:
                issues.append("文档包内文件损坏: %s" % bad)
    except Exception as e:
        issues.append("文档无法作为 .docx 包读取（可能损坏或加密）: %s" % e)

    # 点2：正文非空
    paras = [p.text for p in doc.paragraphs]
    body_text = "".join(paras)
    if len(body_text.strip()) < 50 and not doc.tables:
        issues.append("文档正文为空或内容过少，无法评估")

    # 点2：可编辑（无编辑限制 / 无只读推荐 / 无强制保护）
    try:
        st = doc.settings.element
    except Exception:
        st = None
    if st is not None:
        dp = st.find(qn("w:documentProtection"))
        if dp is not None:
            edit = dp.get(qn("w:edit"))
            enforce = dp.get(qn("w:enforcement"))
            if edit in ("readOnly", "comments", "forms", "trackedChanges") and \
                    str(enforce).lower() in ("1", "true", "on"):
                issues.append("文档启用了编辑限制（%s），无法修改" % edit)
        if st.find(qn("w:writeProtection")) is not None:
            issues.append("文档设置了写保护，无法修改")

    # 点2：文件系统只读属性
    try:
        if os.path.exists(docx_path) and not os.access(docx_path, os.W_OK):
            issues.append("文件为只读，无法保存修改")
    except Exception:
        pass

    # 点3：中文译文，汉字占多数（参考文献著录按西文规范，不计入统计）
    ref_at = None
    for i, t in enumerate(paras):
        if re.sub(r"\s", "", t) == "参考文献":
            ref_at = i
            break
    all_text = "".join(paras[:ref_at] if ref_at is not None else paras)
    for t in doc.tables:
        for row in t.rows:
            for c in row.cells:
                all_text += c.text
    solid = re.sub(r"\s", "", all_text)
    cjk = len(CJK_RE.findall(solid))
    if solid:
        if cjk * 2 <= len(solid):
            issues.append("汉字占比 %.1f%%，未达半数，不符合中文译文要求"
                          % (100.0 * cjk / len(solid)))
    else:
        issues.append("文档无有效文字内容")

    return (len(issues) == 0), issues


# ============================================================
# 维度二：完成度
# ============================================================


class Evaluator(object):

    def __init__(self, doc, docx_path):
        self.doc = doc
        self.path = docx_path
        self.res = FormatResolver(doc, docx_path)
        self.hits = []                 # [(label, delta), ...]
        self.details = {}              # label -> 说明文字
        self.paras = list(doc.paragraphs)
        self.texts = [p.text.strip() for p in self.paras]
        self.ftexts = [self._full_text(p) for p in self.paras]
        self.tables = list(doc.tables)
        self.raw = ""
        try:
            with zipfile.ZipFile(docx_path) as z:
                self.raw = z.read("word/document.xml").decode("utf-8", "ignore")
                self.parts = z.namelist()
        except Exception:
            self.parts = []
        self._classify()

    # --------------------------------------------------------
    # 结构定位：题头区、章节标题、正文段、图注、表题、公式段、参考文献条目
    # --------------------------------------------------------
    def _classify(self):
        n = len(self.paras)
        ts = self.texts

        self.ref_idx = None
        for i, t in enumerate(ts):
            if re.sub(r"\s", "", t) == "参考文献":
                self.ref_idx = i
                break
        limit = self.ref_idx if self.ref_idx is not None else n

        nonempty = [i for i, t in enumerate(ts) if t]
        self.title_idx = nonempty[0] if nonempty else None
        self.author_idx = nonempty[1] if len(nonempty) > 1 else None

        self.abs_idx = None
        for i, t in enumerate(ts[:limit]):
            if re.sub(r"\s", "", t) == "摘要":
                self.abs_idx = i
                break
        self.abs_body_idx = None
        if self.abs_idx is not None:
            for i in range(self.abs_idx + 1, limit):
                if ts[i]:
                    self.abs_body_idx = i
                    break

        self.kw_idx = None
        for i, t in enumerate(ts[:limit]):
            if t.startswith("关键词"):
                self.kw_idx = i
                break

        head_end = max([x for x in (self.kw_idx, self.abs_body_idx,
                                    self.author_idx, self.title_idx) if x is not None] or [-1])
        self.head_end = head_end

        # 图片段：含 w:drawing 或 w:pict 的段落
        self.img_paras = []
        self.img_count = 0
        for i, p in enumerate(self.paras):
            k = len(p._p.findall(".//" + qn("w:drawing"))) + \
                len(p._p.findall(".//" + qn("w:pict")))
            if k:
                self.img_paras.append(i)
                self.img_count += k

        # 正文流顺序（段落与表格的先后关系）
        self.body_seq = []
        pi = ti = 0
        for ch in self.doc.element.body.iterchildren():
            if ch.tag == qn("w:p"):
                if pi < len(self.paras):
                    self.body_seq.append(("p", pi))
                pi += 1
            elif ch.tag == qn("w:tbl"):
                if ti < len(self.tables):
                    self.body_seq.append(("t", ti))
                ti += 1

        # 图注：图片段之后的第一个非空段落且以"图"加序号开头
        self.fig_idx = []
        for k in self.img_paras:
            for j in range(k + 1, n):
                if ts[j]:
                    if FIG_RE.match(ts[j]):
                        self.fig_idx.append(j)
                    break

        # 表题：表格上方最近的非空段落且以"表"加序号开头
        self.prev_para = {}
        last_p = None
        for kind, idx in self.body_seq:
            if kind == "p":
                if ts[idx]:
                    last_p = idx
            else:
                self.prev_para[idx] = last_p
        self.tab_idx = []
        for k in range(len(self.tables)):
            pi2 = self.prev_para.get(k)
            if pi2 is not None and TAB_RE.match(ts[pi2]):
                self.tab_idx.append(pi2)

        self.formula_idx = []
        for i, p in enumerate(self.paras):
            if i >= limit:
                continue
            if self._has_omath(p):
                self.formula_idx.append(i)
            elif FORMULA_NO_RE.search(self.ftexts[i]) \
                    and re.search(r"[=＝＋+×/∑√－−]", self.ftexts[i]) \
                    and len(self.ftexts[i]) < 200:
                self.formula_idx.append(i)

        self.h1_idx, self.h2_idx, self.h3_idx = [], [], []
        for i, t in enumerate(ts[:limit]):
            if i in self.fig_idx or i in self.tab_idx or i in self.formula_idx:
                continue
            if H3_RE.match(t) and len(t) < 40:
                self.h3_idx.append(i)
            elif H2_RE.match(t) and len(t) < 40:
                self.h2_idx.append(i)
            elif H1_RE.match(t) and len(t) < 24:
                self.h1_idx.append(i)

        skip = set(self.fig_idx) | set(self.tab_idx) | set(self.formula_idx) \
            | set(self.h1_idx) | set(self.h2_idx) | set(self.h3_idx)
        self.body_idx = [i for i, t in enumerate(ts[:limit])
                         if i > head_end and i not in skip and len(t) > 60]

        # 参考文献条目
        self.ref_items = []
        if self.ref_idx is not None:
            for i in range(self.ref_idx + 1, n):
                if ref_item_no(ts[i]) is not None:
                    self.ref_items.append(i)

        # 关键词
        self.keywords = []
        if self.kw_idx is not None:
            tail = re.sub(r"^关键词\s*[:：]?\s*", "", ts[self.kw_idx])
            self.keywords = [x for x in re.split(r"[;；]", tail) if x.strip()]

    @staticmethod
    def _has_omath(para):
        for _ in para._p.iter(M_NS + "oMath"):
            return True
        return False

    @staticmethod
    def _full_text(para):
        """段落完整文字：含公式对象内部字符（m:t），制表符按空格计"""
        out = []
        for el in para._p.iter():
            if el.tag == qn("w:t") or el.tag == M_NS + "t":
                out.append(el.text or "")
            elif el.tag == qn("w:tab"):
                out.append(" ")
        return "".join(out).strip()

    def _plain_text(self, para):
        """去掉真实上下标之后的文字（用于判断角标是否被普通字符替代）"""
        out = []
        for r in para.runs:
            if self.res.vert_align(para, r) in ("superscript", "subscript"):
                continue
            out.append(r.text)
        return "".join(out)

    def para_size(self, para):
        """段落主字号（取首个有文字 run 的字号）"""
        for r in para.runs:
            if r.text.strip():
                return self.res.font(para, r)[2]
        return None

    @staticmethod
    def _uniq_cells(table):
        """按元素身份去重（合并单元格会重复出现；不可用 id() 去重）"""
        out = []
        for row in table.rows:
            for c in row.cells:
                if not any(x._tc is c._tc for x in out):
                    out.append(c)
        return out

    # --------------------------------------------------------
    # 通用判定与记分
    # --------------------------------------------------------
    def rep(self, label, max_delta, ok, detail=""):
        """记录一条检查项：满足记满分，不满足记 0"""
        self.hits.append((label, max_delta if ok else 0))
        self.details[label] = detail

    def font_check(self, para, run, ea=None, asc=None, size=None, bold=None):
        f_ea, f_asc, sz, bd = self.res.font(para, run)
        if size is not None and not near(sz, size):
            return False
        if bold is not None and bool(bd) != bool(bold):
            return False
        if ea is not None and has_cjk(run.text) and f_ea != ea:
            return False
        if asc is not None and has_alnum(run.text) and f_asc != asc:
            return False
        return True

    def pf_check(self, para, align=None, lr=None, ln=None,
                 fi=None, hang=None, left=None, size_hint=None):
        pf = self.res.pf(para)
        if align is not None and not align_ok(pf, align):
            return False
        if lr is not None and not spacing_ok(pf, lr, ln):
            return False
        base = size_hint or self.para_size(para) or 12.0
        if fi is not None and not near(first_line_chars(pf, base), fi):
            return False
        if hang is not None and not near(hanging_chars(pf, base), hang):
            return False
        if left is not None and not near(left_chars(pf, base), left):
            return False
        return True

    def chk(self, para, align=None, ea=None, asc=None, size=None, bold=None,
            lr=None, ln=None, fi=None, hang=None, left=None):
        """段落版式 + 段内所有非空 run 的字体，全部满足才返回 True"""
        if para is None:
            return False
        if not self.pf_check(para, align=align, lr=lr, ln=ln, fi=fi,
                             hang=hang, left=left, size_hint=size):
            return False
        seen = False
        for r in para.runs:
            if not r.text.strip():
                continue
            seen = True
            if not self.font_check(para, r, ea=ea, asc=asc, size=size, bold=bold):
                return False
        return seen

    # --------------------------------------------------------
    # 逐条检查（顺序与评分细则一致）
    # --------------------------------------------------------
    def check_all(self):
        self.eval_page_setup()
        self.eval_header_text()
        self.eval_doc_title()
        self.eval_author_line()
        self.eval_abstract_title()
        self.eval_abstract_body()
        self.eval_keyword_line()
        self.eval_h1_exist()
        self.eval_h1_format()
        self.eval_h2_numbers()
        self.eval_h3_numbers()
        self.eval_body_layout()
        self.eval_body_font()
        self.eval_citation_superscript()
        self.eval_unit_superscript()
        self.eval_image_paragraphs()
        self.eval_figcap_format()
        self.eval_figcap_numbers()
        self.eval_table_dims()
        self.eval_three_line_tables()
        self.eval_table_inner_format()
        self.eval_tabcap_format()
        self.eval_table_cells_filled()
        self.eval_tabcap_coverage()
        self.eval_table1_vmerge()
        self.eval_formula_numbers()
        self.eval_formula_scripts()
        self.eval_ref_title()
        self.eval_ref_items_format()
        self.eval_ref_items_spacing()
        self.eval_global_cn_fonts()
        self.eval_global_latin_font()
        self.eval_global_real_scripts()
        self.eval_no_dup_para()
        self.eval_no_markup()
        self.eval_fullwidth_punct()
        self.eval_object_counts()
        self.eval_no_english_block()
        self.eval_structure_kept()

    # --------------------------------------------------------
    # 分值：+1
    # 规则原文：页面为 A4 纵向（21cm × 29.7cm），页边距上 3.3cm、下 2.6cm、左 2.1cm、右 2.2cm
    # 定位口径：文档首节页面设置
    # 点1：纸张宽高为 A4 纵向 11906 × 16838 twips
    # 点2：上下左右页边距依次为 3.3 / 2.6 / 2.1 / 2.2 厘米
    # 全部满足 → +1；否则 0
    # --------------------------------------------------------
    def eval_page_setup(self):
        LABEL = "页面A4纵向21×29.7厘米，页边距上3.3下2.6左2.1右2.2厘米"
        ok, detail = False, "未取到页面设置"
        sec = self.doc.sections[0] if self.doc.sections else None
        if sec is not None:
            w, h = emu_to_twip(sec.page_width), emu_to_twip(sec.page_height)
            mg = (emu_to_cm(sec.top_margin), emu_to_cm(sec.bottom_margin),
                  emu_to_cm(sec.left_margin), emu_to_cm(sec.right_margin))
            size_ok = (w is not None and h is not None
                       and abs(w - A4_TWIPS[0]) <= 20 and abs(h - A4_TWIPS[1]) <= 20)
            mg_ok = all(near(a, b, 0.03) for a, b in zip(mg, (3.3, 2.6, 2.1, 2.2)))
            ok = bool(size_ok and mg_ok)
            detail = "纸张 %s×%s twips；页边距 上%s 下%s 左%s 右%s cm" % ((w, h) + mg)
        self.rep(LABEL, 1, ok, detail)

    # --------------------------------------------------------
    # 分值：+1
    # 规则原文：页眉文字为"华语青月大学本科生毕业设计（论文）用纸"
    # 定位口径：全部页眉部件
    # 点1：至少一个页眉含该文字
    # 点2：所有非空页眉都含该文字
    # 全部满足 → +1；否则 0
    # --------------------------------------------------------
    def eval_header_text(self):
        LABEL = "页眉文字为华语青月大学本科生毕业设计（论文）用纸"
        target = re.sub(r"\s", "", HEADER_TEXT)
        found, others = 0, 0
        try:
            with zipfile.ZipFile(self.path) as z:
                for nm in z.namelist():
                    if not re.match(r"word/header\d*\.xml$", nm):
                        continue
                    txt = re.sub(r"<[^>]+>", "", z.read(nm).decode("utf-8", "ignore"))
                    txt = re.sub(r"\s", "", txt)
                    if not txt:
                        continue
                    if target in txt:
                        found += 1
                    else:
                        others += 1
        except Exception:
            pass
        self.rep(LABEL, 1, found > 0 and others == 0,
                 "含目标文字的页眉 %d 个，其他非空页眉 %d 个" % (found, others))

    # --------------------------------------------------------
    # 分值：+3
    # 规则原文：文章标题独立成段，居中，中文黑体、西文新罗马，18 磅（小二），不加粗，
    #           固定行距 27.5 磅，无首行缩进
    # 定位口径：正文第一个非空段落
    # 点1：居中且无首行缩进；点2：固定行距 27.5 磅；
    # 点3：中文黑体、西文 Times New Roman、18 磅、不加粗
    # 全部满足 → +3；否则 0
    # --------------------------------------------------------
    def eval_doc_title(self):
        LABEL = "文章标题独立成段居中，黑体+新罗马18磅不加粗，固定27.5磅无缩进"
        p = self.paras[self.title_idx] if self.title_idx is not None else None
        ok = self.chk(p, align="CENTER", ea=SIMHEI, asc=TNR, size=18.0, bold=False,
                      lr="exact", ln=EXACT_LINE, fi=0)
        self.rep(LABEL, 3, ok, "标题段序号 %s" % self.title_idx)

    # --------------------------------------------------------
    # 分值：+1
    # 规则原文：作者姓名段居中，中文宋体、西文新罗马，12 磅（小四），不加粗，
    #           固定行距 27.5 磅，无首行缩进
    # 定位口径：文章标题之后的第一个非空段落
    # 全部满足 → +1；否则 0
    # --------------------------------------------------------
    def eval_author_line(self):
        LABEL = "作者姓名段居中，宋体+新罗马12磅不加粗，固定27.5磅无缩进"
        p = self.paras[self.author_idx] if self.author_idx is not None else None
        ok = self.chk(p, align="CENTER", ea=SIMSUN, asc=TNR, size=12.0, bold=False,
                      lr="exact", ln=EXACT_LINE, fi=0)
        self.rep(LABEL, 1, ok, "作者段序号 %s" % self.author_idx)

    # --------------------------------------------------------
    # 分值：+3
    # 规则原文："摘  要"标题段居中，中文黑体，16 磅（三号），不加粗，
    #           固定行距 27.5 磅，无首行缩进
    # 定位口径：去空格后文字为"摘要"的段落
    # 全部满足 → +3；否则 0
    # --------------------------------------------------------
    def eval_abstract_title(self):
        LABEL = "摘要标题段居中，黑体16磅不加粗，固定27.5磅无缩进"
        p = self.paras[self.abs_idx] if self.abs_idx is not None else None
        ok = self.chk(p, align="CENTER", ea=SIMHEI, size=16.0, bold=False,
                      lr="exact", ln=EXACT_LINE, fi=0)
        self.rep(LABEL, 3, ok, "摘要标题段序号 %s" % self.abs_idx)

    # --------------------------------------------------------
    # 分值：+3
    # 规则原文：摘要正文独立成段，两端对齐，首行缩进 2 字符，中文宋体、西文新罗马，
    #           12 磅（小四），不加粗，固定行距 27.5 磅
    # 定位口径：摘要标题之后的第一个非空段落
    # 全部满足 → +3；否则 0
    # --------------------------------------------------------
    def eval_abstract_body(self):
        LABEL = "摘要正文两端对齐首行缩进2字符，宋体+新罗马12磅不加粗，固定27.5磅"
        p = self.paras[self.abs_body_idx] if self.abs_body_idx is not None else None
        ok = self.chk(p, align="JUSTIFY", ea=SIMSUN, asc=TNR, size=12.0, bold=False,
                      lr="exact", ln=EXACT_LINE, fi=2)
        self.rep(LABEL, 3, ok, "摘要正文段序号 %s" % self.abs_body_idx)

    # --------------------------------------------------------
    # 分值：+3
    # 规则原文：关键词段整行（含"关键词："标签）中文黑体 10.5 磅（五号），不加粗，
    #           两端对齐，无首行缩进，固定行距 27.5 磅，关键词不少于 5 个且用分号分隔
    # 定位口径：以"关键词"开头的段落
    # 点1：版式与字体；点2：分号分隔且不少于 5 个
    # 全部满足 → +3；否则 0
    # --------------------------------------------------------
    def eval_keyword_line(self):
        LABEL = "关键词段黑体五号不加粗两端对齐无缩进固定27.5磅且不少于5个分号分隔"
        p = self.paras[self.kw_idx] if self.kw_idx is not None else None
        fmt = self.chk(p, align="JUSTIFY", ea=SIMHEI, size=10.5, bold=False,
                       lr="exact", ln=EXACT_LINE, fi=0)
        cnt_ok = len(self.keywords) >= 5
        self.rep(LABEL, 3, bool(fmt and cnt_ok),
                 "格式%s；关键词 %d 个" % ("符合" if fmt else "不符合", len(self.keywords)))

    # --------------------------------------------------------
    # 分值：+5
    # 规则原文：存在 4 个一级章节标题段，各自独立成段，形式为一位阿拉伯数字加空格加标题名
    # 定位口径：参考文献之前、形如"1 引言"的独立段落
    # 全部满足 → +5；否则 0
    # --------------------------------------------------------
    def eval_h1_exist(self):
        LABEL = "存在4个一级章节标题段，形式为一位数字加空格加标题名"
        nums = [self.texts[i].split()[0] for i in self.h1_idx]
        ok = len(self.h1_idx) == 4
        self.rep(LABEL, 5, ok, "一级标题 %d 个：%s" % (len(self.h1_idx), nums))

    # --------------------------------------------------------
    # 分值：+5
    # 规则原文：4 个一级标题均居中，中文黑体，16 磅（三号），不加粗，
    #           固定行距 27.5 磅，无首行缩进
    # 定位口径：全部一级标题段
    # 全部满足 → +5；否则 0
    # --------------------------------------------------------
    def eval_h1_format(self):
        LABEL = "一级标题均居中，黑体16磅不加粗，固定27.5磅无缩进"
        bad = [self.texts[i][:12] for i in self.h1_idx
               if not self.chk(self.paras[i], align="CENTER", ea=SIMHEI, size=16.0,
                               bold=False, lr="exact", ln=EXACT_LINE, fi=0)]
        ok = bool(self.h1_idx) and not bad
        self.rep(LABEL, 5, ok, "不合格 %d 个：%s" % (len(bad), bad))

    # --------------------------------------------------------
    # 分值：+3
    # 规则原文：存在 16 个二级标题段，编号 2.1–2.10 与 3.1–3.6 完整
    # 定位口径：形如"2.1 xxx"的独立段落
    # 点1：数量 16；点2：编号序列与 2.1–2.10、3.1–3.6 完全一致
    # 全部满足 → +3；否则 0
    # --------------------------------------------------------
    def eval_h2_numbers(self):
        LABEL = "16个二级标题段，编号2.1至2.10与3.1至3.6完整"
        nums = [H2_RE.match(self.texts[i]).group(1) for i in self.h2_idx]
        ok = nums == H2_EXPECT
        miss = [x for x in H2_EXPECT if x not in nums]
        self.rep(LABEL, 3, ok, "实得 %d 个，缺失 %s" % (len(nums), miss or "无"))

    # --------------------------------------------------------
    # 分值：+3
    # 规则原文：存在 6 个三级标题段，编号 2.1.1、2.1.2、2.3.1–2.3.4 完整
    # 定位口径：形如"2.1.1 xxx"的独立段落
    # 全部满足 → +3；否则 0
    # --------------------------------------------------------
    def eval_h3_numbers(self):
        LABEL = "6个三级标题段，编号2.1.1、2.1.2、2.3.1至2.3.4完整"
        nums = [H3_RE.match(self.texts[i]).group(1) for i in self.h3_idx]
        ok = nums == H3_EXPECT
        miss = [x for x in H3_EXPECT if x not in nums]
        self.rep(LABEL, 3, ok, "实得 %d 个，缺失 %s" % (len(nums), miss or "无"))

    # --------------------------------------------------------
    # 分值：+5
    # 规则原文：正文段落两端对齐、首行缩进 2 字符、固定行距 27.5 磅
    # 定位口径：参考文献之前、题头区之后、字数多于 60 且非标题/图注/表题/公式的段落
    # 全部满足 → +5；否则 0
    # --------------------------------------------------------
    def eval_body_layout(self):
        LABEL = "正文段落两端对齐、首行缩进2字符、固定行距27.5磅"
        bad = [i for i in self.body_idx
               if not self.pf_check(self.paras[i], align="JUSTIFY", lr="exact",
                                    ln=EXACT_LINE, fi=2, size_hint=12.0)]
        ok = bool(self.body_idx) and not bad
        self.rep(LABEL, 5, ok, "正文 %d 段，不合格段序号 %s"
                 % (len(self.body_idx), bad[:10] or "无"))

    # --------------------------------------------------------
    # 分值：+5
    # 规则原文：正文中文宋体、数字与拉丁字母新罗马、12 磅（小四）、不加粗
    # 定位口径：全部正文段的非空 run
    # 全部满足 → +5；否则 0
    # --------------------------------------------------------
    def eval_body_font(self):
        LABEL = "正文中文宋体、数字字母新罗马、12磅不加粗"
        bad = []
        for i in self.body_idx:
            p = self.paras[i]
            for r in p.runs:
                if not r.text.strip():
                    continue
                if not self.font_check(p, r, ea=SIMSUN, asc=TNR, size=12.0, bold=False):
                    bad.append(i)
                    break
        ok = bool(self.body_idx) and not bad
        self.rep(LABEL, 5, ok, "不合格段序号 %s" % (bad[:10] or "无"))

    # --------------------------------------------------------
    # 分值：+3
    # 规则原文：正文文献引用为方括号加序号且全部为真实上标
    # 定位口径：参考文献之前正文段落中的方括号序号
    # 点1：存在引用标注；点2：每处引用覆盖的 run 均为上标
    # 全部满足 → +3；否则 0
    # --------------------------------------------------------
    def eval_citation_superscript(self):
        LABEL = "正文文献引用为方括号序号且全部为真实上标"
        total, bad = 0, []
        limit = self.ref_idx if self.ref_idx is not None else len(self.paras)
        for i in range(limit):
            p = self.paras[i]
            text = p.text
            if "[" not in text:
                continue
            spans = [(m.start(), m.end()) for m in CIT_RE.finditer(text)]
            if not spans:
                continue
            pos, marks = 0, []
            for r in p.runs:
                ln = len(r.text)
                marks.append((pos, pos + ln, self.res.vert_align(p, r)))
                pos += ln
            for s, e in spans:
                total += 1
                covers = [va for (a, b, va) in marks if a < e and b > s]
                if not covers or any(va != "superscript" for va in covers):
                    bad.append((i, text[s:e]))
        self.rep(LABEL, 3, total > 0 and not bad,
                 "引用 %d 处，非上标 %d 处：%s" % (total, len(bad), bad[:5] or "无"))

    # --------------------------------------------------------
    # 分值：+3
    # 规则原文：单位与量符号的指数为真实上标，符号与角标之间不夹空格或下划线
    # 定位口径：正文与表格内文字
    # 点1：去掉真实上下标后的文字中不出现"mg·g-1""R2"一类替代写法
    # 点2：上标或下标 run 之前不夹下划线
    # 全部满足 → +3；否则 0
    # --------------------------------------------------------
    def eval_unit_superscript(self):
        LABEL = "单位与量符号指数为真实上标，符号与角标间无空格或下划线"
        bad = []
        limit = self.ref_idx if self.ref_idx is not None else len(self.paras)
        for i in range(limit):
            p = self.paras[i]
            m = UP_BAD_RE.search(self._plain_text(p))
            if m:
                bad.append((i, m.group(0)))
                continue
            prev = ""
            for r in p.runs:
                va = self.res.vert_align(p, r)
                if va in ("superscript", "subscript") and (
                        prev.endswith("_") or (r.text or "").startswith("_")):
                    bad.append((i, "角标与符号之间有下划线"))
                    break
                prev = r.text or prev
        for t in self.tables:
            for c in self._uniq_cells(t):
                for p in c.paragraphs:
                    m = UP_BAD_RE.search(self._plain_text(p))
                    if m:
                        bad.append(("表格", m.group(0)))
        self.rep(LABEL, 3, not bad, "疑似问题 %d 处：%s" % (len(bad), bad[:5] or "无"))

    # --------------------------------------------------------
    # 分值：+5
    # 规则原文：文中含 11 张图片，每张图片单独占一段
    # 定位口径：含 w:drawing / w:pict 的段落
    # 点1：图片总数 11；点2：每个图片段只有图片、无正文文字；点3：一段只放一张图
    # 全部满足 → +5；否则 0
    # --------------------------------------------------------
    def eval_image_paragraphs(self):
        LABEL = "含11张图片且每张图片单独占一段"
        multi = [i for i in self.img_paras
                 if len(self.paras[i]._p.findall(".//" + qn("w:drawing")))
                 + len(self.paras[i]._p.findall(".//" + qn("w:pict"))) > 1]
        withtext = [i for i in self.img_paras if self.texts[i]]
        ok = (self.img_count == FIG_COUNT and len(self.img_paras) == FIG_COUNT
              and not multi and not withtext)
        self.rep(LABEL, 5, ok, "图片 %d 张，占 %d 段；多图段 %s；带文字段 %s"
                 % (self.img_count, len(self.img_paras), multi or "无", withtext or "无"))

    # --------------------------------------------------------
    # 分值：+5
    # 规则原文：图注段居中，中文宋体、西文新罗马，10.5 磅（五号），加粗，
    #           单倍行距，无首行缩进
    # 定位口径：以"图"加序号开头的段落
    # 全部满足 → +5；否则 0（本条不判颜色）
    # --------------------------------------------------------
    def eval_figcap_format(self):
        LABEL = "图注段居中，宋体+新罗马10.5磅加粗，单倍行距无缩进"
        bad = [i for i in self.fig_idx
               if not self.chk(self.paras[i], align="CENTER", ea=SIMSUN, asc=TNR,
                               size=10.5, bold=True, lr="auto", fi=0)]
        ok = bool(self.fig_idx) and not bad
        self.rep(LABEL, 5, ok, "图注 %d 段，不合格 %s"
                 % (len(self.fig_idx), [self.texts[i][:10] for i in bad][:5] or "无"))

    # --------------------------------------------------------
    # 分值：+3
    # 规则原文：图注为"图"加序号加图题，序号 1–11 连续、无缺号无重号
    # 定位口径：全部图注段
    # 全部满足 → +3；否则 0
    # --------------------------------------------------------
    def eval_figcap_numbers(self):
        LABEL = "图注为图加序号加图题，序号1至11连续无缺无重"
        nums = [int(FIG_RE.match(self.texts[i]).group(1)) for i in self.fig_idx]
        titled = [i for i in self.fig_idx
                  if len(re.sub(r"^图\s*\d+\s*", "", self.texts[i])) >= 2]
        ok = nums == list(range(1, FIG_COUNT + 1)) and len(titled) == len(self.fig_idx)
        self.rep(LABEL, 3, ok, "序号序列 %s；有图题 %d/%d"
                 % (nums, len(titled), len(self.fig_idx)))

    # --------------------------------------------------------
    # 分值：+5
    # 规则原文：文中含 6 个表格，行列数依次为 6×4、5×5、4×4、6×5、7×4、6×4
    # 定位口径：文档中全部表格，按出现顺序
    # 全部满足 → +5；否则 0
    # --------------------------------------------------------
    def eval_table_dims(self):
        LABEL = "含6个表格且行列数依次为6×4、5×5、4×4、6×5、7×4、6×4"
        dims = [(len(t.rows), len(t.columns)) for t in self.tables]
        self.rep(LABEL, 5, dims == TABLE_DIMS, "实测行列 %s" % dims)

    # --------------------------------------------------------
    # 框线解析：w:tcBorders 覆盖 w:tblBorders 覆盖表样式；
    # 内部横竖线看 insideH / insideV，w:sz 为 1/8 磅
    # --------------------------------------------------------
    def _tbl_borders(self, tbl):
        own = sty = None
        tblPr = tbl._tbl.find(qn("w:tblPr"))
        if tblPr is not None:
            own = tblPr.find(qn("w:tblBorders"))
            ts = tblPr.find(qn("w:tblStyle"))
            if ts is not None:
                st = self.res.styles.get(ts.get(qn("w:val")))
                if st is not None:
                    p = st.find(qn("w:tblPr"))
                    if p is not None:
                        sty = p.find(qn("w:tblBorders"))
        return own, sty

    @staticmethod
    def _border_pt(bel, side):
        """返回磅值；"na" 表示该层未定义，None 表示无线"""
        if bel is None:
            return "na"
        b = bel.find(qn("w:" + side))
        if b is None:
            return "na"
        if (b.get(qn("w:val")) or "nil") in ("nil", "none"):
            return None
        v = _f(b.get(qn("w:sz")))
        return round(v / 8.0, 3) if v is not None else 0.5

    def _edge_tc(self, tbl, tc, side, ri, ci, nrows, ncols, row_el=None):
        """按 tcBorders > 行级 tblPrEx > tblBorders > 表样式 的覆盖链解析某条边"""
        outer = ((ri == 0 and side == "top") or (ri == nrows - 1 and side == "bottom")
                 or (ci == 0 and side == "left") or (ci == ncols - 1 and side == "right"))
        tside = side if outer else ("insideH" if side in ("top", "bottom") else "insideV")
        tcPr = tc.find(qn("w:tcPr")) if tc is not None else None
        v = self._border_pt(tcPr.find(qn("w:tcBorders")) if tcPr is not None else None, side)
        if v != "na":
            return v
        if row_el is not None:
            ex = row_el.find(qn("w:tblPrEx"))
            if ex is not None:
                v = self._border_pt(ex.find(qn("w:tblBorders")), tside)
                if v != "na":
                    return v
        own, sty = self._tbl_borders(tbl)
        v = self._border_pt(own, tside)
        if v != "na":
            return v
        v = self._border_pt(sty, tside)
        return None if v == "na" else v

    @staticmethod
    def _vmerge_state(tc):
        """None＝非纵向合并；"restart"＝合并起始；"continue"＝合并延续"""
        if tc is None:
            return None
        tcPr = tc.find(qn("w:tcPr"))
        vm = tcPr.find(qn("w:vMerge")) if tcPr is not None else None
        if vm is None:
            return None
        return "restart" if vm.get(qn("w:val")) == "restart" else "continue"

    # --------------------------------------------------------
    # 分值：+5
    # 规则原文：6 个表格均为三线表：上下框线 1.5 磅单实线，表头行下 0.5 磅或 0.75 磅单实线，
    #           无竖线，无其他横线
    # 定位口径：逐单元格解析框线覆盖链
    # 点1：首行上边线全为 1.5 磅；点2：末行下边线全为 1.5 磅；
    # 点3：首行下边线（或次行上边线）为 0.5/0.75 磅且不缺；
    # 点4：其余横线全部无；点5：内部竖线全部无
    # 全部满足 → +5；否则 0
    # --------------------------------------------------------
    def eval_three_line_tables(self):
        LABEL = "6个表格均为三线表：上下框线1.5磅、表头线0.5或0.75磅、无竖线无其他横线"
        bad = []
        for k, tbl in enumerate(self.tables):
            gm, trs, ncols = self._grid_map(tbl)
            nrows = len(trs)
            if nrows < 2 or ncols < 1:
                bad.append(k + 1)
                continue
            top, bot, hdr, mid, ver = set(), set(), set(), set(), set()
            for ci in range(ncols):
                t0, t1 = gm.get((0, ci)), gm.get((1, ci))
                tl = gm.get((nrows - 1, ci))
                top.add(self._edge_tc(tbl, t0, "top", 0, ci, nrows, ncols, trs[0]))
                a = self._edge_tc(tbl, t0, "bottom", 0, ci, nrows, ncols, trs[0])
                b = self._edge_tc(tbl, t1, "top", 1, ci, nrows, ncols, trs[1])
                hdr.add(a if a is not None else b)
                bot.add(self._edge_tc(tbl, tl, "bottom", nrows - 1, ci,
                                      nrows, ncols, trs[nrows - 1]))
                for ri in range(1, nrows - 1):
                    if self._vmerge_state(gm.get((ri + 1, ci))) == "continue":
                        continue          # 纵向合并区域内部不存在行边界
                    mid.add(self._edge_tc(tbl, gm.get((ri, ci)), "bottom", ri, ci,
                                          nrows, ncols, trs[ri]))
                    mid.add(self._edge_tc(tbl, gm.get((ri + 1, ci)), "top", ri + 1, ci,
                                          nrows, ncols, trs[ri + 1]))
                for ri in range(nrows):
                    tc = gm.get((ri, ci))
                    if ci > 0:
                        ver.add(self._edge_tc(tbl, tc, "left", ri, ci,
                                              nrows, ncols, trs[ri]))
                    if ci < ncols - 1:
                        ver.add(self._edge_tc(tbl, tc, "right", ri, ci,
                                              nrows, ncols, trs[ri]))
            good = (top == {1.5} and bot == {1.5} and hdr <= {0.5, 0.75}
                    and None not in hdr and hdr and mid <= {None} and ver <= {None})
            if not good:
                bad.append(k + 1)
        ok = len(self.tables) == TAB_COUNT and not bad
        self.rep(LABEL, 5, ok, "不合格表 %s" % (bad or "无"))

    # --------------------------------------------------------
    # 分值：+3
    # 规则原文：表格整体居中，表内非空段落居中，10.5 磅（五号），
    #           中文宋体或黑体，西文新罗马
    # 定位口径：表格属性 w:jc 与表内全部非空单元格段落
    # 全部满足 → +3；否则 0
    # --------------------------------------------------------
    def eval_table_inner_format(self):
        LABEL = "表格整体居中，表内文字居中、五号、中文宋体或黑体、西文新罗马"
        bad = []
        for k, tbl in enumerate(self.tables):
            tblPr = tbl._tbl.find(qn("w:tblPr"))
            jc = tblPr.find(qn("w:jc")) if tblPr is not None else None
            if jc is None or (jc.get(qn("w:val")) or "").lower() != "center":
                bad.append((k + 1, "表格未居中"))
                continue
            for c in self._uniq_cells(tbl):
                for p in c.paragraphs:
                    if not p.text.strip():
                        continue
                    if not self.pf_check(p, align="CENTER", size_hint=10.5):
                        bad.append((k + 1, "段落未居中: " + p.text[:10]))
                        break
                    okrun = True
                    for r in p.runs:
                        if not r.text.strip():
                            continue
                        f_ea, f_asc, sz, _ = self.res.font(p, r)
                        if not near(sz, 10.5):
                            okrun = False
                        if has_cjk(r.text) and f_ea not in (SIMSUN, SIMHEI):
                            okrun = False
                        if has_alnum(r.text) and f_asc != TNR:
                            okrun = False
                        if not okrun:
                            bad.append((k + 1, "字体字号: " + r.text[:10]))
                            break
                    if not okrun:
                        break
        ok = bool(self.tables) and not bad
        self.rep(LABEL, 3, ok, "问题 %d 处：%s" % (len(bad), bad[:5] or "无"))

    # --------------------------------------------------------
    # 分值：+5
    # 规则原文：表题段居中，中文宋体、西文新罗马，10.5 磅（五号），加粗，
    #           固定行距 27.5 磅，无首行缩进，以"表"加序号开头，序号 1–6 连续
    # 定位口径：以"表"加序号开头的段落
    # 全部满足 → +5；否则 0
    # --------------------------------------------------------
    def eval_tabcap_format(self):
        LABEL = "表题段居中，宋体+新罗马10.5磅加粗，固定27.5磅无缩进，序号表1至表6连续"
        bad = [i for i in self.tab_idx
               if not self.chk(self.paras[i], align="CENTER", ea=SIMSUN, asc=TNR,
                               size=10.5, bold=True, lr="exact", ln=EXACT_LINE, fi=0)]
        nums = [int(TAB_RE.match(self.texts[i]).group(1)) for i in self.tab_idx]
        ok = not bad and nums == list(range(1, TAB_COUNT + 1))
        self.rep(LABEL, 5, ok, "表题 %d 段，序号 %s，格式不合格 %s"
                 % (len(self.tab_idx), nums,
                    [self.texts[i][:10] for i in bad][:5] or "无"))

    # --------------------------------------------------------
    # 分值：+3
    # 规则原文：所有表格单元格均有内容，无空单元格
    # 定位口径：按元素身份去重后的全部单元格
    # 全部满足 → +3；否则 0
    # --------------------------------------------------------
    def eval_table_cells_filled(self):
        LABEL = "所有表格单元格均有内容，无空单元格"
        total = empty = 0
        loc = []
        for k, tbl in enumerate(self.tables):
            for c in self._uniq_cells(tbl):
                total += 1
                if not c.text.strip():
                    empty += 1
                    if len(loc) < 5:
                        loc.append("表%d" % (k + 1))
        ok = total > 0 and empty == 0
        self.rep(LABEL, 3, ok, "单元格 %d 个，空 %d 个：%s" % (total, empty, loc or "无"))

    # --------------------------------------------------------
    # 分值：+3
    # 规则原文：每张表格上方都有表题段；无表题的表格视为续表，
    #           其首行文字必须与上一表格首行完全相同
    # 定位口径：正文流顺序中表格上方最近的非空段落
    # 点1：有表题 → 通过；点2：无表题 → 首行与上一表格首行逐格相同
    # 全部满足 → +3；否则 0
    # --------------------------------------------------------
    def eval_tabcap_coverage(self):
        LABEL = "每张表格上方有表题，无表题者首行与上一表格首行完全相同"
        prev_para = {}
        last_p = None
        for kind, idx in self.body_seq:
            if kind == "p":
                if self.texts[idx]:
                    last_p = idx
            else:
                prev_para[idx] = last_p
        bad = []
        for k, tbl in enumerate(self.tables):
            pi = prev_para.get(k)
            if pi is not None and TAB_RE.match(self.texts[pi]):
                continue
            if k == 0:
                bad.append(k + 1)
                continue
            cur = [c.text.strip() for c in tbl.rows[0].cells] if tbl.rows else []
            pv = self.tables[k - 1]
            head = [c.text.strip() for c in pv.rows[0].cells] if pv.rows else []
            if cur != head or not cur:
                bad.append(k + 1)
        ok = bool(self.tables) and not bad
        self.rep(LABEL, 3, ok, "无表题且首行不匹配的表 %s" % (bad or "无"))

    def _grid_map(self, tbl):
        """(行, 列) → tc 元素；纵向合并的续行保留自身 tc（不被上方单元格顶替）"""
        gm, ncols = {}, 0
        trs = tbl._tbl.findall(qn("w:tr"))
        for ri, tr in enumerate(trs):
            ci = 0
            for tc in tr.findall(qn("w:tc")):
                tcPr = tc.find(qn("w:tcPr"))
                gs = 1
                if tcPr is not None:
                    g = tcPr.find(qn("w:gridSpan"))
                    if g is not None:
                        gs = int(_f(g.get(qn("w:val")), 1) or 1)
                for k in range(gs):
                    gm[(ri, ci + k)] = tc
                ci += gs
            ncols = max(ncols, ci)
        return gm, trs, ncols

    # --------------------------------------------------------
    # 分值：+3
    # 规则原文：第一个表格末两列各跨 5 个数据行纵向合并
    # 定位口径：第一个表格最后两列的 w:vMerge
    # 点1：末两列各存在一个纵向合并区域；点2：该区域覆盖行数不少于 5
    # 全部满足 → +3；否则 0
    # --------------------------------------------------------
    def eval_table1_vmerge(self):
        LABEL = "第一个表格末两列各跨5个数据行纵向合并"
        ok, detail = False, "无表格"
        if self.tables:
            gm, trs, ncols = self._grid_map(self.tables[0])
            nrows = len(trs)
            spans = []
            for ci in (ncols - 2, ncols - 1):
                best = cur = 0
                for ri in range(nrows):
                    st = self._vmerge_state(gm.get((ri, ci)))
                    if st is None:
                        cur = 0
                    else:
                        cur = 1 if st == "restart" else cur + 1
                        best = max(best, cur)
                spans.append(best)
            ok = ncols >= 2 and all(s >= 5 for s in spans)
            detail = "末两列纵向合并跨行 %s" % spans
        self.rep(LABEL, 3, ok, detail)

    # --------------------------------------------------------
    # 分值：+5
    # 规则原文：8 个公式各自独立成段，编号 (1)–(8) 完整连续、无重号
    # 定位口径：含公式对象的段落，或以"(序号)"结尾且含运算符的段落
    # 点1：公式段共 8 个；点2：编号序列为 1–8
    # 全部满足 → +5；否则 0
    # --------------------------------------------------------
    def eval_formula_numbers(self):
        LABEL = "8个公式各自独立成段，编号(1)至(8)完整连续无重号"
        nums = []
        for i in self.formula_idx:
            m = FORMULA_NO_RE.search(self.ftexts[i])
            if m:
                nums.append(int(m.group(1)))
        ok = (len(self.formula_idx) == FORMULA_COUNT
              and nums == list(range(1, FORMULA_COUNT + 1)))
        self.rep(LABEL, 5, ok, "公式段 %d 个，编号 %s" % (len(self.formula_idx), nums))

    # --------------------------------------------------------
    # 分值：+3
    # 规则原文：公式中的上下标为真实上标下标，减号为标准减号，
    #           不用下划线、脱字符或普通文字替代
    # 定位口径：全部公式段
    # 点1：不含 _ ^ ** 等替代写法；点2：不含相似 Unicode 角标字符；
    # 点3：不用短横线以外的连接符充当减号
    # 全部满足 → +3；否则 0
    # --------------------------------------------------------
    def eval_formula_scripts(self):
        LABEL = "公式上下标为真实上下标、减号为标准减号，不用替代写法"
        bad = []
        for i in self.formula_idx:
            t = self.ftexts[i]
            if re.search(r"[_\^]|\*\*", t):
                bad.append((i, "含替代角标写法"))
                continue
            if any(ch in t for ch in UNI_SUPSUB):
                bad.append((i, "含相似角标字符"))
                continue
            if re.search(r"[0-9A-Za-z]\s*[–—~]\s*[0-9A-Za-z]", t):
                bad.append((i, "减号非标准"))
        self.rep(LABEL, 3, bool(self.formula_idx) and not bad,
                 "问题 %d 处：%s" % (len(bad), bad[:5] or "无"))

    # --------------------------------------------------------
    # 分值：+3
    # 规则原文："参考文献"标题段独立成段、居中、中文黑体、14 磅（四号）、不加粗、
    #           固定行距 27.5 磅，条目序号 1–40 连续
    # 定位口径：去空格后文字为"参考文献"的段落及其后的条目段
    # 全部满足 → +3；否则 0
    # --------------------------------------------------------
    def eval_ref_title(self):
        LABEL = "参考文献标题居中、黑体14磅不加粗、固定27.5磅，条目序号1至40连续"
        p = self.paras[self.ref_idx] if self.ref_idx is not None else None
        fmt = self.chk(p, align="CENTER", ea=SIMHEI, size=14.0, bold=False,
                       lr="exact", ln=EXACT_LINE, fi=0)
        nums = [ref_item_no(self.texts[i]) for i in self.ref_items]
        num_ok = nums == list(range(1, REF_COUNT + 1))
        self.rep(LABEL, 3, bool(fmt and num_ok),
                 "标题格式%s；条目 %d 条" % ("符合" if fmt else "不符合", len(nums)))

    # --------------------------------------------------------
    # 分值：+5
    # 规则原文：40 条参考文献各自独立成段，两端对齐，左缩进 0，悬挂缩进 2 字符，
    #           12 磅（小四），固定行距 27.5 磅
    # 定位口径：参考文献标题之后的全部条目段
    # 全部满足 → +5；否则 0
    # --------------------------------------------------------
    def eval_ref_items_format(self):
        LABEL = "40条参考文献两端对齐、左缩进0、悬挂缩进2字符、12磅、固定27.5磅"
        bad = []
        for i in self.ref_items:
            p = self.paras[i]
            if not self.pf_check(p, align="JUSTIFY", lr="exact", ln=EXACT_LINE,
                                 hang=2, left=0, size_hint=12.0):
                bad.append(i)
                continue
            for r in p.runs:
                if r.text.strip() and not near(self.res.font(p, r)[2], 12.0):
                    bad.append(i)
                    break
        ok = len(self.ref_items) == REF_COUNT and not bad
        self.rep(LABEL, 5, ok, "条目 %d 条，不合格 %s"
                 % (len(self.ref_items), bad[:10] or "无"))

    # --------------------------------------------------------
    # 分值：+3
    # 规则原文：条目内作者、题名、刊名、年份、卷期页码之间空格完整，
    #           无无空格长串、无标点后缺空格
    # 定位口径：全部参考文献条目段
    # 全部满足 → +3；否则 0
    # --------------------------------------------------------
    def eval_ref_items_spacing(self):
        LABEL = "参考文献条目内空格完整，无标点后缺空格与无空格长串"
        bad = []
        for i in self.ref_items:
            t = self.texts[i]
            m = GLUE_RE.search(t) or LONG_RE.search(t)
            if m:
                bad.append((i, m.group(0)))
        self.rep(LABEL, 3, bool(self.ref_items) and not bad,
                 "问题 %d 处：%s" % (len(bad), bad[:5] or "无"))

    def _all_paras(self):
        """全文段落：正文段 + 表格单元格内段落"""
        for p in self.paras:
            yield p
        for t in self.tables:
            for c in self._uniq_cells(t):
                for p in c.paragraphs:
                    yield p

    # --------------------------------------------------------
    # 分值：+5
    # 规则原文：全文中文仅使用宋体与黑体
    # 定位口径：全文含汉字的 run（含表格内）
    # 全部满足 → +5；否则 0
    # --------------------------------------------------------
    def eval_global_cn_fonts(self):
        LABEL = "全文中文仅使用宋体与黑体"
        bad = {}
        for p in self._all_paras():
            for r in p.runs:
                if not has_cjk(r.text):
                    continue
                f_ea = self.res.font(p, r)[0]
                if f_ea not in (SIMSUN, SIMHEI):
                    bad[f_ea] = bad.get(f_ea, 0) + 1
        self.rep(LABEL, 5, not bad, "异常中文字体 %s" % (bad or "无"))

    # --------------------------------------------------------
    # 分值：+3
    # 规则原文：全文数字与拉丁字母均为新罗马字体
    # 定位口径：全文含数字或拉丁字母的 run（含表格内）
    # 全部满足 → +3；否则 0
    # --------------------------------------------------------
    def eval_global_latin_font(self):
        LABEL = "全文数字与拉丁字母均为新罗马字体"
        bad = {}
        for p in self._all_paras():
            for r in p.runs:
                if not has_alnum(r.text):
                    continue
                f_asc = self.res.font(p, r)[1]
                if f_asc != TNR:
                    bad[f_asc] = bad.get(f_asc, 0) + 1
        self.rep(LABEL, 3, not bad, "异常西文字体 %s" % (bad or "无"))

    # --------------------------------------------------------
    # 分值：+3
    # 规则原文：全文角标均为真实上下标，不用相似 Unicode 字符替代
    # 定位口径：全文文字（含表格内）
    # 全部满足 → +3；否则 0
    # --------------------------------------------------------
    def eval_global_real_scripts(self):
        LABEL = "全文角标均为真实上下标，不用相似Unicode字符替代"
        bad = []
        for p in self._all_paras():
            t = p.text
            hit = [ch for ch in t if ch in UNI_SUPSUB]
            if hit:
                bad.append((t[:16], "".join(hit)))
        self.rep(LABEL, 3, not bad, "含相似角标字符 %d 段：%s" % (len(bad), bad[:5] or "无"))

    # --------------------------------------------------------
    # 分值：+1
    # 规则原文：全文无文字完全相同的整段重复
    # 定位口径：字数多于 40 的正文段落
    # 全部满足 → +1；否则 0
    # --------------------------------------------------------
    def eval_no_dup_para(self):
        LABEL = "全文无文字完全相同的整段重复"
        seen, dup = {}, []
        for i, t in enumerate(self.texts):
            key = re.sub(r"\s", "", t)
            if len(key) <= 40:
                continue
            if key in seen:
                dup.append((seen[key], i))
            else:
                seen[key] = i
        self.rep(LABEL, 1, not dup, "重复段 %d 对：%s" % (len(dup), dup[:5] or "无"))

    # --------------------------------------------------------
    # 分值：+1
    # 规则原文：全文无批注、无修订标记、无文字高亮
    # 定位口径：document.xml 原始标记与批注部件
    # 点1：无 w:ins / w:del 修订；点2：无批注引用与批注部件；点3：无 w:highlight
    # 全部满足 → +1；否则 0
    # --------------------------------------------------------
    def eval_no_markup(self):
        LABEL = "全文无批注、无修订标记、无文字高亮"
        ins = len(re.findall(r"<w:ins[ >]", self.raw))
        dele = len(re.findall(r"<w:del[ >]", self.raw))
        cmt = len(re.findall(r"commentReference", self.raw))
        hl = 0
        for p in self._all_paras():
            for r in p.runs:
                if r.text.strip() and self.res.highlighted(p, r):
                    hl += 1
        part = any(nm.startswith("word/comments") for nm in self.parts)
        ok = not (ins or dele or cmt or hl or part)
        self.rep(LABEL, 1, ok, "修订 %d/%d，批注引用 %d，批注部件 %s，高亮 run %d"
                 % (ins, dele, cmt, "有" if part else "无", hl))

    # --------------------------------------------------------
    # 分值：+1
    # 规则原文：中文段落使用全角标点，不用半角逗号句号充当中文标点
    # 定位口径：含汉字的段落（参考文献条目按西文著录规范不计入）
    # 全部满足 → +1；否则 0
    # --------------------------------------------------------
    def eval_fullwidth_punct(self):
        LABEL = "中文段落使用全角标点，不用半角标点充当中文标点"
        bad = []
        refs = set(self.ref_items)
        for i, t in enumerate(self.texts):
            if i in refs or not has_cjk(t):
                continue
            m = HALF_PUNCT_RE.search(t)
            if m:
                bad.append((i, m.group(0)))
        self.rep(LABEL, 1, not bad, "半角标点 %d 处：%s" % (len(bad), bad[:5] or "无"))

    # --------------------------------------------------------
    # 分值：+3
    # 规则原文：图表公式全部保留：11 张图、6 个表格、8 个公式
    # 定位口径：全文图片、表格、公式对象数量
    # 全部满足 → +3；否则 0
    # --------------------------------------------------------
    def eval_object_counts(self):
        LABEL = "图表公式全部保留：11张图、6个表格、8个公式"
        omath = len(re.findall(r"<m:oMath[ >]", self.raw))
        fcnt = max(len(self.formula_idx), omath)
        ok = (self.img_count == FIG_COUNT and len(self.tables) == TAB_COUNT
              and fcnt >= FORMULA_COUNT)
        self.rep(LABEL, 3, ok, "图 %d，表 %d，公式 %d"
                 % (self.img_count, len(self.tables), fcnt))

    # --------------------------------------------------------
    # 分值：+3
    # 规则原文：正文无成段英文残留
    # 定位口径：全部正文段（题头区与参考文献不计入）
    # 点1：同段内 3 个字母以上的英文词不足 6 个，或汉字占比不低于 20%
    # 全部满足 → +3；否则 0
    # --------------------------------------------------------
    def eval_no_english_block(self):
        LABEL = "正文无成段英文残留"
        bad = []
        for i in self.body_idx:
            t = self.texts[i]
            if len(t) < 20:
                continue
            words = re.findall(r"[A-Za-z]{3,}", t)
            solid = re.sub(r"\s", "", t)
            ratio = len(CJK_RE.findall(solid)) / float(len(solid) or 1)
            if len(words) >= 6 and ratio < 0.2:
                bad.append((i, t[:20]))
        self.rep(LABEL, 3, not bad, "疑似英文残留 %d 段：%s" % (len(bad), bad[:5] or "无"))

    # --------------------------------------------------------
    # 分值：+1
    # 规则原文：关键词与章节层级完整保留：关键词 6 个、二级标题 16 个、三级标题 6 个
    # 定位口径：关键词段与章节标题段
    # 全部满足 → +1；否则 0
    # --------------------------------------------------------
    def eval_structure_kept(self):
        LABEL = "关键词6个、二级标题16个、三级标题6个完整保留"
        ok = (len(self.keywords) == KEYWORD_COUNT and len(self.h2_idx) == 16
              and len(self.h3_idx) == 6)
        self.rep(LABEL, 1, ok, "关键词 %d，二级 %d，三级 %d"
                 % (len(self.keywords), len(self.h2_idx), len(self.h3_idx)))

# ============================================================
# 各检查项满分（hit = (delta == max_delta)）
# 全部为正向加分项，不设扣分项；合计 127 分
# ============================================================

RULE_MAX_DELTA = {
    "页面A4纵向21×29.7厘米，页边距上3.3下2.6左2.1右2.2厘米": 1,
    "页眉文字为华语青月大学本科生毕业设计（论文）用纸": 1,
    "文章标题独立成段居中，黑体+新罗马18磅不加粗，固定27.5磅无缩进": 3,
    "作者姓名段居中，宋体+新罗马12磅不加粗，固定27.5磅无缩进": 1,
    "摘要标题段居中，黑体16磅不加粗，固定27.5磅无缩进": 3,
    "摘要正文两端对齐首行缩进2字符，宋体+新罗马12磅不加粗，固定27.5磅": 3,
    "关键词段黑体五号不加粗两端对齐无缩进固定27.5磅且不少于5个分号分隔": 3,
    "存在4个一级章节标题段，形式为一位数字加空格加标题名": 5,
    "一级标题均居中，黑体16磅不加粗，固定27.5磅无缩进": 5,
    "16个二级标题段，编号2.1至2.10与3.1至3.6完整": 3,
    "6个三级标题段，编号2.1.1、2.1.2、2.3.1至2.3.4完整": 3,
    "正文段落两端对齐、首行缩进2字符、固定行距27.5磅": 5,
    "正文中文宋体、数字字母新罗马、12磅不加粗": 5,
    "正文文献引用为方括号序号且全部为真实上标": 3,
    "单位与量符号指数为真实上标，符号与角标间无空格或下划线": 3,
    "含11张图片且每张图片单独占一段": 5,
    "图注段居中，宋体+新罗马10.5磅加粗，单倍行距无缩进": 5,
    "图注为图加序号加图题，序号1至11连续无缺无重": 3,
    "含6个表格且行列数依次为6×4、5×5、4×4、6×5、7×4、6×4": 5,
    "6个表格均为三线表：上下框线1.5磅、表头线0.5或0.75磅、无竖线无其他横线": 5,
    "表格整体居中，表内文字居中、五号、中文宋体或黑体、西文新罗马": 3,
    "表题段居中，宋体+新罗马10.5磅加粗，固定27.5磅无缩进，序号表1至表6连续": 5,
    "所有表格单元格均有内容，无空单元格": 3,
    "每张表格上方有表题，无表题者首行与上一表格首行完全相同": 3,
    "第一个表格末两列各跨5个数据行纵向合并": 3,
    "8个公式各自独立成段，编号(1)至(8)完整连续无重号": 5,
    "公式上下标为真实上下标、减号为标准减号，不用替代写法": 3,
    "参考文献标题居中、黑体14磅不加粗、固定27.5磅，条目序号1至40连续": 3,
    "40条参考文献两端对齐、左缩进0、悬挂缩进2字符、12磅、固定27.5磅": 5,
    "参考文献条目内空格完整，无标点后缺空格与无空格长串": 3,
    "全文中文仅使用宋体与黑体": 5,
    "全文数字与拉丁字母均为新罗马字体": 3,
    "全文角标均为真实上下标，不用相似Unicode字符替代": 3,
    "全文无文字完全相同的整段重复": 1,
    "全文无批注、无修订标记、无文字高亮": 1,
    "中文段落使用全角标点，不用半角标点充当中文标点": 1,
    "图表公式全部保留：11张图、6个表格、8个公式": 3,
    "正文无成段英文残留": 3,
    "关键词6个、二级标题16个、三级标题6个完整保留": 1,
}

def _find_docx_in_dir(dir_path):
    """在目录下查找待评估的 .docx（跳过 Word 临时文件）"""
    try:
        names = os.listdir(dir_path)
    except Exception:
        return None
    cands = [n for n in names
             if n.lower().endswith(".docx") and not n.startswith("~$")]
    cands.sort()
    for n in cands:
        full = os.path.join(dir_path, n)
        if os.path.isfile(full):
            return full
    return None


def _fail(result, exc):
    """脚本未能完成判定：置 status=error，dim1_pass 留空，dim1_reason 也为空。"""
    result["status"] = "error"
    result["error"] = "%s: %s" % (type(exc).__name__, exc)
    result["dim1_pass"] = ""
    result["dim1_reason"] = ""
    result["dim2_items"] = []
    result["total_score"] = 0
    return result


def _is_damaged_package(exc):
    """判断异常是否源于交付文件自身结构损坏（缺部件、XML 不合法等）。"""
    text = str(exc)
    if isinstance(exc, KeyError) and "There is no item named" in text:
        return True
    return type(exc).__name__ in ("XMLSyntaxError", "ParseError", "BadZipFile")


def _damaged(result, exc):
    """交付文件损坏属于维度一不通过（正常 0 分），不是脚本运行错误。"""
    result["status"] = "ok"
    result["error"] = None
    result["dim1_pass"] = False
    result["dim1_reason"] = "交付文件结构不完整，无法正常打开：%s: %s" % (
        type(exc).__name__, exc,
    )
    result["dim2_items"] = []
    result["total_score"] = 0
    return result


def evaluate(dir_path: str) -> dict:
    result = {
        "id": SCRIPT_ID,
        "file_name": "",
        "status": "ok",
        "error": None,
        "dim1_pass": False,
        "dim1_reason": "",
        "dim2_items": [],
        "total_score": 0,
        "max_score": sum(v for v in RULE_MAX_DELTA.values() if v > 0),
    }

    if not dir_path or not os.path.isdir(dir_path):
        return _fail(result, FileNotFoundError("目录不存在：%s" % dir_path))

    docx_path = _find_docx_in_dir(dir_path)
    if not docx_path:
        return _fail(result, FileNotFoundError("目录内未找到 .docx 文件：%s" % dir_path))
    result["file_name"] = os.path.basename(docx_path)

    try:
        doc = Document(docx_path)
    except Exception as e:
        if _is_damaged_package(e):
            return _damaged(result, e)
        return _fail(result, e)

    try:
        ok, issues = check_dimension_one(doc, docx_path)
        result["dim1_pass"] = ok
        # dim1_reason 仅在维度一未通过时给出原因摘要；通过则为空
        result["dim1_reason"] = "" if ok else "; ".join(issues)
        if not ok:
            # 维度一未通过：完成度直接判 0，不再检查维度二
            return result

        ev = Evaluator(doc, docx_path)
        ev.check_all()

        total = 0
        for label, delta in ev.hits:
            max_delta = RULE_MAX_DELTA.get(label, 0)
            total += delta
            result["dim2_items"].append({
                "rule": label,
                "max_delta": max_delta,
                "delta": delta,
                "hit": (delta == max_delta),
            })
        result["total_score"] = total
    except Exception as e:
        if _is_damaged_package(e):
            return _damaged(result, e)
        return _fail(result, e)

    return result


if __name__ == "__main__":
    _dir = sys.argv[1] if len(sys.argv) > 1 else os.path.dirname(os.path.abspath(__file__))
    _payload = json.dumps(evaluate(_dir), ensure_ascii=False, indent=2)
    try:
        sys.stdout.buffer.write((_payload + "\n").encode("utf-8"))
    except AttributeError:
        print(_payload)

