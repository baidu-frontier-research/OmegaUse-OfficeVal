"""自动评估脚本：多岗位从业人员综合复习题_答案回填版.docx

按接口统一约定，本模块只暴露 evaluate(dir_path) 一个函数：
- 接收"脚本所在目录"的路径，脚本自身负责在该目录里定位并打开被评估文档
- 返回结构化 dict（含维度一通过与否、维度二逐项得分、总分）
- 不 print 主结果、不修改 sys.stdout、不 sys.exit
本文件末尾的 __main__ 分支仅用于本地调试。
"""
import sys, zipfile, re, json, os
from xml.etree import ElementTree as ET
from dataclasses import dataclass, field
from typing import Optional
from typing import Any as _Any  # 异构 dict 值（str/bool/float/None）的类型标注

# ─── 命名空间 ────────────────────────────────────────────────────────────────
NS = {
    "w":  "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "r":  "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
    "a":  "http://schemas.openxmlformats.org/drawingml/2006/main",
    "v":  "urn:schemas-microsoft-com:vml",
    "o":  "urn:schemas-microsoft-com:office:office",
    "ep": "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties",
    "cp": "http://schemas.openxmlformats.org/package/2006/metadata/core-properties",
    "pkg":"http://schemas.openxmlformats.org/package/2006/relationships",
}
for _pfx, _uri in NS.items():
    try: ET.register_namespace(_pfx, _uri)
    except Exception: pass

def qn(prefix, tag): return f"{{{NS[prefix]}}}{tag}"
def _get(el, *path):
    """沿 path 查找子元素，返回第一个匹配或 None"""
    cur = el
    for step in path:
        cur = cur.find(step, NS)
        if cur is None: return None
    return cur
def _attr(el, prefix, attr, default=None):
    v = el.get(qn(prefix, attr))
    return v if v is not None else default
def _findall(el, path): return el.findall(path, NS)
def _find(el, path): return el.find(path, NS)
def _text_in_elem(el):
    """提取 elem 下所有 w:t 文本(含文本框/DrawingML)"""
    parts = []
    for t in el.iter(qn("w","t")):
        parts.append(t.text or "")
    for t in el.iter(qn("a","t")):
        parts.append(t.text or "")
    return "".join(parts)


# ─── 数据结构 ─────────────────────────────────────────────────────────────────
@dataclass
class RunInfo:
    text: str
    font: Optional[str] = None          # eastAsia / ascii
    size_pt: Optional[float] = None     # 磅
    bold: Optional[bool] = None
    italic: Optional[bool] = None
    underline: Optional[bool] = None
    highlight: Optional[str] = None     # 高亮颜色
    color: Optional[str] = None         # RGB hex / "auto"
    vertical: bool = False              # 竖排

@dataclass
class ParagraphInfo:
    text: str
    runs: list = field(default_factory=list)  # list[RunInfo]
    alignment: Optional[str] = None
    line_spacing_rule: Optional[str] = None
    line_spacing_val: Optional[float] = None
    before_pt: float | None = None
    after_pt: float | None = None
    before_lines: float | None = None
    after_lines: float | None = None
    style_id: str | None = None
    source: str = "body"               # body / header / footer / textbox
    part: str | None = None            # 来源 xml part，如 word/header1.xml

@dataclass
class Question:
    kind: str                  # "single" | "judge"
    number: int
    paragraph_index: int
    text: str                  # 题干全文(含括号)
    answer: str                # 括号内字符
    answer_in_tail_paren: bool # 答案是否在末尾全角括号
    options: list = field(default_factory=list)  # ["A...","B...","C...","D..."]
    answer_runs: list = field(default_factory=list)  # 答案字符所在/重叠的 RunInfo
    answer_paragraph_index: int | None = None  # 答案所在段落，处理判断题跨段续行

@dataclass
class RuleResult:
    rule_id: str
    title: str
    score: int                 # 0=门槛项
    passed: bool
    evidence: str
    confidence: str = "高"    # 高/中/低
    fatal: bool = False        # 维度1项目


# ─── DocxPackage ─────────────────────────────────────────────────────────────
class DocxPackage:
    def __init__(self, path: str):
        self.path = path
        self.zf: Optional[zipfile.ZipFile] = None
        self.names: list = []
        self._xml_cache: dict = {}

    def open(self):
        self.zf = zipfile.ZipFile(self.path, "r")
        self.names = self.zf.namelist()
        return self

    def close(self):
        if self.zf: self.zf.close()

    def has(self, part: str) -> bool:
        return part in self.names

    def read_bytes(self, part: str) -> bytes:
        return self.zf.read(part)

    def parse_xml(self, part: str) -> Optional[ET.Element]:
        if part in self._xml_cache: return self._xml_cache[part]
        if not self.has(part): return None
        try:
            tree = ET.fromstring(self.zf.read(part))
            self._xml_cache[part] = tree
            return tree
        except ET.ParseError:
            return None

    def media_count(self) -> int:
        return sum(1 for n in self.names if n.startswith("word/media/"))

    def header_parts(self) -> list:
        return [n for n in self.names if re.match(r"word/header\d*\.xml", n)]

    def footer_parts(self) -> list:
        return [n for n in self.names if re.match(r"word/footer\d*\.xml", n)]

    def parse_rels(self) -> dict:
        """返回 {rId: target} 来自 word/_rels/document.xml.rels"""
        root = self.parse_xml("word/_rels/document.xml.rels")
        if root is None: return {}
        rels = {}
        for rel in root:
            rid = rel.get("Id","")
            target = rel.get("Target","")
            rels[rid] = target
        return rels


# ─── 样式解析 ─────────────────────────────────────────────────────────────────
def _bool_prop(parent, name):
    el = _find(parent, f"w:{name}") if parent is not None else None
    if el is None: return None
    val = _attr(el, "w", "val", "true")
    return val not in ("false", "0", "off")

def _half_point_to_pt(v):
    try: return int(v) / 2.0
    except Exception: return None

def _twips_to_pt(v):
    try: return int(v) / 20.0
    except Exception: return None

def _parse_rpr(rpr) -> dict:
    d = {}
    if rpr is None: return d
    rf = _find(rpr, "w:rFonts")
    if rf is not None:
        d["font"] = _attr(rf, "w", "eastAsia") or _attr(rf, "w", "ascii") or _attr(rf, "w", "hAnsi")
    sz = _find(rpr, "w:sz")
    if sz is not None:
        d["size_pt"] = _half_point_to_pt(_attr(sz, "w", "val"))
    for key, tag in [("bold","b"),("italic","i")]:
        val = _bool_prop(rpr, tag)
        if val is not None: d[key] = val
    u = _find(rpr, "w:u")
    if u is not None:
        d["underline"] = (_attr(u, "w", "val", "single") not in ("none", "0", "false"))
    hi = _find(rpr, "w:highlight")
    if hi is not None: d["highlight"] = _attr(hi, "w", "val")
    color = _find(rpr, "w:color")
    if color is not None: d["color"] = _attr(color, "w", "val")
    vert = _find(rpr, "w:vertAlign")
    if vert is not None: d["vertical"] = True
    return d

def _parse_ppr(ppr) -> dict:
    d = {}
    if ppr is None: return d
    jc = _find(ppr, "w:jc")
    if jc is not None: d["alignment"] = _attr(jc, "w", "val")
    spacing = _find(ppr, "w:spacing")
    if spacing is not None:
        d["line_spacing_rule"] = _attr(spacing, "w", "lineRule")
        d["line_spacing_val"] = _twips_to_pt(_attr(spacing, "w", "line"))
        d["before_pt"] = _twips_to_pt(_attr(spacing, "w", "before"))
        d["after_pt"] = _twips_to_pt(_attr(spacing, "w", "after"))
        try:
            bl = _attr(spacing, "w", "beforeLines")
            if bl is not None: d["before_lines"] = int(bl) / 100.0
            al = _attr(spacing, "w", "afterLines")
            if al is not None: d["after_lines"] = int(al) / 100.0
        except Exception: pass
    td = _find(ppr, "w:textDirection")
    if td is not None: d["text_direction"] = _attr(td, "w", "val", "")
    return d

class StyleResolver:
    def __init__(self, pkg: DocxPackage):
        self.pkg = pkg
        self.styles = {}
        self.defaults = {"r": {}, "p": {}}
        self._load()

    def _load(self):
        root = self.pkg.parse_xml("word/styles.xml")
        if root is None: return
        dd = _find(root, "w:docDefaults")
        if dd is not None:
            rpr = _find(dd, "w:rPrDefault/w:rPr")
            ppr = _find(dd, "w:pPrDefault/w:pPr")
            self.defaults["r"] = _parse_rpr(rpr)
            self.defaults["p"] = _parse_ppr(ppr)
        for st in _findall(root, "w:style"):
            sid = _attr(st, "w", "styleId")
            if not sid: continue
            based = _find(st, "w:basedOn")
            self.styles[sid] = {
                "type": _attr(st, "w", "type"),
                "based_on": _attr(based, "w", "val") if based is not None else None,
                "r": _parse_rpr(_find(st, "w:rPr")),
                "p": _parse_ppr(_find(st, "w:pPr")),
            }

    def _style_chain(self, sid):
        chain, seen = [], set()
        while sid and sid not in seen and sid in self.styles and len(chain) < 20:
            seen.add(sid); chain.append(sid); sid = self.styles[sid].get("based_on")
        return list(reversed(chain))

    def resolve_para(self, p) -> tuple:
        ppr = _find(p, "w:pPr")
        st_el = _find(ppr, "w:pStyle") if ppr is not None else None
        sid = _attr(st_el, "w", "val") if st_el is not None else None
        pr = dict(self.defaults["p"]); rr = dict(self.defaults["r"])
        for s in self._style_chain(sid):
            pr.update(self.styles[s].get("p",{})); rr.update(self.styles[s].get("r",{}))
        pr.update(_parse_ppr(ppr)); return sid, pr, rr

    def resolve_run(self, r, inherited_r: dict) -> dict:
        out = dict(inherited_r)
        rpr = _find(r, "w:rPr")
        st_el = _find(rpr, "w:rStyle") if rpr is not None else None
        sid = _attr(st_el, "w", "val") if st_el is not None else None
        for s in self._style_chain(sid): out.update(self.styles[s].get("r",{}))
        out.update(_parse_rpr(rpr)); return out


# ─── 文档解析 ─────────────────────────────────────────────────────────────────
class DocxParser:
    def __init__(self, pkg: DocxPackage):
        self.pkg = pkg
        self.styles = StyleResolver(pkg)
        self.paragraphs: list[ParagraphInfo] = []
        self.header_paragraphs: list[ParagraphInfo] = []
        self.footer_paragraphs: list[ParagraphInfo] = []
        self.referenced_header_parts: list[str] = []
        self.body_text = ""
        self.all_text = ""
        self.watermark_candidates: list[dict[str, _Any]] = []    # 异构 dict，值可能为 str/bool/float/None
        self.page_signals: dict[str, int | None] = {}
        self.visual_risks: list[str] = []
        self.drawing_count = 0
        self.image_count = pkg.media_count()
        self.comment_count = 0
        self.revision_count = 0
        # 按文档序采集的 body <w:p> 元信息，用于分页可编辑性校验
        self.body_paragraphs_meta: list[dict[str, int]] = []
        self.body_page_of_para: list[int] = []
        self.body_page_segments: list[list[int]] = []
        self.body_detected_pages: int = 0
        self.page_size_twips: tuple[int, int] = (11906, 16838)   # 默认 A4 (twips)
        self.page_editability: dict[str, object] = {}             # 分页可编辑性检测结果

    def parse(self):
        root = self.pkg.parse_xml("word/document.xml")
        if root is None: return self
        self.paragraphs = self._extract_paragraphs(root, "body")
        self.body_text = "\n".join(p.text for p in self.paragraphs)
        self._extract_headers_footers()
        self._extract_page_signals(root)
        self._extract_body_structure(root)
        self._extract_watermarks(root)
        self._extract_risks(root)
        self.all_text = "\n".join([self.body_text] + [p.text for p in self.header_paragraphs] + [p.text for p in self.footer_paragraphs])
        return self

    def _extract_paragraphs(self, root, source) -> list:
        out = []
        for p in root.iter(qn("w","p")):
            sid, pstyle, rbase = self.styles.resolve_para(p)
            runs, parts = [], []
            for r in p.findall("w:r", NS):
                txt_parts = []
                for child in list(r):
                    if child.tag == qn("w","t"):
                        txt_parts.append(child.text or "")
                    elif child.tag == qn("w","tab"):
                        txt_parts.append("\t")
                    elif child.tag == qn("w","br"):
                        txt_parts.append("\n")
                txt = "".join(txt_parts)
                if not txt: continue
                props = self.styles.resolve_run(r, rbase)
                runs.append(RunInfo(text=txt, **{k:v for k,v in props.items() if k in RunInfo.__dataclass_fields__}))
                parts.append(txt)
            # 文本框/DrawingML 文本可能不在直接 run 文本里，兜底补充
            if not parts:
                t = _text_in_elem(p)
                if t: parts.append(t)
            text = "".join(parts)
            if text.strip() or source != "body":
                out.append(ParagraphInfo(text=text, runs=runs, style_id=sid, source=source,
                                         alignment=pstyle.get("alignment"),
                                         line_spacing_rule=pstyle.get("line_spacing_rule"),
                                         line_spacing_val=pstyle.get("line_spacing_val"),
                                         before_pt=pstyle.get("before_pt"), after_pt=pstyle.get("after_pt"),
                                         before_lines=pstyle.get("before_lines"), after_lines=pstyle.get("after_lines")))
        return out

    def _referenced_header_parts(self, document_root) -> list[str]:
        rels = self.pkg.parse_rels()
        parts: list[str] = []
        for ref in document_root.iter(qn("w", "headerReference")):
            rid = _attr(ref, "r", "id")
            target = rels.get(rid, "") if rid else ""
            if not target: continue
            if target.startswith("/"):
                part = target.lstrip("/")
            elif target.startswith("word/"):
                part = target
            else:
                part = "word/" + target
            if self.pkg.has(part) and part not in parts:
                parts.append(part)
        return parts

    def _extract_headers_footers(self):
        document_root = self.pkg.parse_xml("word/document.xml")
        self.referenced_header_parts = self._referenced_header_parts(document_root) if document_root is not None else []
        for part in self.referenced_header_parts:
            root = self.pkg.parse_xml(part)
            if root is not None:
                ps = self._extract_paragraphs(root, "header")
                for p in ps: p.part = part
                self.header_paragraphs += ps
        for part in self.pkg.footer_parts():
            root = self.pkg.parse_xml(part)
            if root is not None:
                ps = self._extract_paragraphs(root, "footer")
                for p in ps: p.part = part
                self.footer_paragraphs += ps

    def _extract_page_signals(self, root):
        pages_meta = None
        app = self.pkg.parse_xml("docProps/app.xml")
        if app is not None:
            for child in app:
                if child.tag.endswith("Pages") and child.text and child.text.isdigit(): pages_meta = int(child.text)
        last_breaks = len(list(root.iter(qn("w","lastRenderedPageBreak"))))
        page_breaks = sum(1 for br in root.iter(qn("w","br")) if _attr(br,"w","type") == "page")
        sections = len(list(root.iter(qn("w","sectPr"))))
        self.page_signals = {"app_pages": pages_meta, "last_rendered_pages": last_breaks + 1 if last_breaks else None,
                             "explicit_page_breaks": page_breaks, "sections": sections}
        # 读取首个 sectPr 的页面尺寸，用于整页图片化的面积判断
        for sectPr in root.iter(qn("w", "sectPr")):
            pgSz = _find(sectPr, "w:pgSz")
            if pgSz is not None:
                try:
                    w_tw = int(_attr(pgSz, "w", "w") or 11906)
                    h_tw = int(_attr(pgSz, "w", "h") or 16838)
                    self.page_size_twips = (w_tw, h_tw)
                except Exception:
                    pass
                break

    def _extract_body_structure(self, root):
        """按文档顺序遍历 body 中的段落/表格，把每个 <w:p> 归属到某一页。

        分页锚点（保守）：
        - <w:lastRenderedPageBreak/>：Word 上次渲染的分页点（最可靠）
        - <w:br w:type="page"/>：显式分页符
        - <w:sectPr>：分节符（新节可能触发新页；仅当上一页非空时才推进）

        锚点位置以"该锚点所在段落的下一段"作为新页首段（与 Word 渲染一致）。
        无 lastRenderedPageBreak 时，仅依赖显式分页符/分节符做近似分页，
        置信度标注为"低"。
        """
        body = _find(root, "w:body") or root
        cur_page = 1
        pending_new_page = False
        para_pages: list[int] = []
        para_elems: list[ET.Element] = []
        found_last_rendered = False
        for el in list(body):
            tag = el.tag
            if tag == qn("w", "p"):
                if pending_new_page:
                    cur_page += 1
                    pending_new_page = False
                para_pages.append(cur_page)
                para_elems.append(el)
                # 段落内部的分页锚点在渲染时位于段落起始处或段落中；
                # 保守地记为"该段落属于当前页，锚点触发的新页从下一段开始"
                for lrpb in el.iter(qn("w", "lastRenderedPageBreak")):
                    found_last_rendered = True
                    pending_new_page = True
                for br in el.iter(qn("w", "br")):
                    if _attr(br, "w", "type") == "page":
                        pending_new_page = True
            elif tag == qn("w", "sectPr"):
                # 节末 sectPr：下一节起新页（当且仅当类型不是 continuous）
                sect_type = _find(el, "w:type")
                type_val = _attr(sect_type, "w", "val", "nextPage") if sect_type is not None else "nextPage"
                if type_val != "continuous":
                    pending_new_page = True
        self.body_page_of_para = para_pages
        self.body_paragraphs_meta = [{"page": pg} for pg in para_pages]
        self.body_detected_pages = max(para_pages) if para_pages else 0
        # 按页聚合段落索引（0-based 段落序号，对应 self.paragraphs 里同序的 body 段落）
        # 注：_extract_paragraphs 使用 root.iter，遍历顺序与此处 body 顺序一致，且过滤规则相同
        # （text.strip() 非空或 source != "body"；body 中 source=="body" 才可能被过滤）。
        # 为对齐，重新按相同过滤规则收集非空段落的页号。
        aligned_pages: list[int] = []
        for el, pg in zip(para_elems, para_pages):
            txt = _text_in_elem(el)
            if txt.strip():
                aligned_pages.append(pg)
        self.aligned_page_of_paragraph = aligned_pages
        self.page_map_confidence = "高" if found_last_rendered else "低"
        # 逐页做可编辑性检测
        self.page_editability = self._compute_page_editability(root, para_elems, para_pages)

    def _compute_page_editability(self, root, para_elems, para_pages) -> dict:
        """按页统计：文本 run 字符数、drawing 数量、最大图片相对页面面积占比、
        是否存在 v:imagedata/w:pict 图形、是否存在 w:t 文本层的目标字符串。
        """
        page_stats: dict = {}
        max_page = max(para_pages) if para_pages else 0
        for pg in range(1, max_page + 1):
            page_stats[pg] = {
                "text_chars": 0,
                "drawings": 0,
                "images": 0,
                "vml_image": 0,
                "max_image_area_ratio": 0.0,
                "has_qnum": False,
                "has_option": False,
                "has_answer": False,
                "text_only_qnum": True,
                "text_only_option": True,
                "text_only_answer": True,
            }
        pw_tw, ph_tw = self.page_size_twips
        # 1 twip = 635 EMU
        page_area_emu = float(pw_tw) * float(ph_tw) * 635.0 * 635.0 if pw_tw and ph_tw else 0.0
        for el, pg in zip(para_elems, para_pages):
            stat = page_stats.get(pg)
            if stat is None:
                continue
            txt = _text_in_elem(el)
            stripped = txt.strip()
            stat["text_chars"] += len(stripped)
            # drawing / 图片
            for _ in el.iter(qn("w", "drawing")):
                stat["drawings"] += 1
            for ext in el.iter(qn("wp", "extent")):
                try:
                    cx = int(ext.get("cx", "0")); cy = int(ext.get("cy", "0"))
                except Exception:
                    cx = cy = 0
                if page_area_emu > 0 and cx and cy:
                    ratio = (cx * cy) / page_area_emu
                    if ratio > stat["max_image_area_ratio"]:
                        stat["max_image_area_ratio"] = ratio
            # 计算图片 blip 数量（用于识别把文字光栅化成图片的情况）
            for _ in el.iter(qn("a", "blip")):
                stat["images"] += 1
            for _ in el.iter(qn("v", "imagedata")):
                stat["vml_image"] += 1
            # 目标对象是否来自 w:t 文本层
            wt_text = "".join(t.text or "" for t in el.iter(qn("w", "t")))
            if _Q_NUM_RE.match(stripped):
                stat["has_qnum"] = True
                if not _Q_NUM_RE.match(wt_text.strip() or ""):
                    stat["text_only_qnum"] = False
            if _OPT_RE.match(stripped):
                stat["has_option"] = True
                if not _OPT_RE.match(wt_text.strip() or ""):
                    stat["text_only_option"] = False
            # 末尾全角括号里的答案字母/判断符号是否在 w:t 中
            m = list(_FULL_PAREN_RE.finditer(stripped))
            if m:
                inner = m[-1].group(1).strip()
                if inner and (_SINGLE_ANS_RE.match(inner) or _JUDGE_ANS_RE.match(inner)):
                    stat["has_answer"] = True
                    if inner not in wt_text:
                        stat["text_only_answer"] = False
        # 汇总失败页
        failed_pages: list[dict] = []
        for pg, s in page_stats.items():
            reasons: list[str] = []
            # 整页图片化：本页有大幅图片且文字层极少
            if s["max_image_area_ratio"] >= 0.6 and s["text_chars"] < 20:
                reasons.append(f"疑似整页图片化(占比={s['max_image_area_ratio']:.2f},文字={s['text_chars']}字)")
            # 目标对象存在但不来自 w:t 文本层
            if s["has_qnum"] and not s["text_only_qnum"]:
                reasons.append("题号非w:t文本")
            if s["has_option"] and not s["text_only_option"]:
                reasons.append("选项非w:t文本")
            if s["has_answer"] and not s["text_only_answer"]:
                reasons.append("答案字符非w:t文本")
            if reasons:
                failed_pages.append({"page": pg, "reasons": reasons, **s})
        return {
            "detected_pages": max_page,
            "page_map_confidence": self.page_map_confidence,
            "page_size_twips": self.page_size_twips,
            "failed_pages": failed_pages,
            "per_page_stats": page_stats,
        }

    def _extract_watermarks(self, document_root):
        parts = [("word/document.xml", document_root)]
        for part in self.pkg.header_parts()+self.pkg.footer_parts():
            r = self.pkg.parse_xml(part)
            if r is not None: parts.append((part, r))
        for part, root in parts:
            in_header = part.startswith("word/header")
            in_footer = part.startswith("word/footer")
            # VML textpath 水印：读取文字、字体、字号、颜色、透明度和倾斜角度
            for shape in root.iter(qn("v", "shape")):
                style = shape.get("style", "") or ""
                fillcolor = shape.get("fillcolor", "") or ""
                strokecolor = shape.get("strokecolor", "") or ""
                opacity = ""
                fill = shape.find("v:fill", NS)
                if fill is not None:
                    opacity = fill.get("opacity", "") or ""
                    fillcolor = fill.get("color", "") or fillcolor
                rotation = None
                m_rot = re.search(r"rotation\s*:\s*(-?\d+(?:\.\d+)?)", style, re.I)
                if m_rot:
                    try: rotation = float(m_rot.group(1))
                    except Exception: rotation = None
                # 判断"页面背景"：VML 水印典型特征——position:absolute + z-index<0 或负；
                # 或包含在 <w:pict><v:shape> 内并置于 header/footer 中；
                # 或形状类型为 Word 生成的水印类型 (t136/PowerPlusWaterMarkObject/WordPictureWatermark 命名)
                is_absolute = "position:absolute" in style.lower()
                z_neg = False
                m_z = re.search(r"z-index\s*:\s*(-?\d+)", style, re.I)
                if m_z:
                    try: z_neg = int(m_z.group(1)) < 0
                    except Exception: pass
                shape_id = shape.get("id", "") or ""
                type_ref = shape.get("type", "") or ""
                name_like_wm = bool(re.search(r"WaterMark|WordPictureWatermark|PowerPlusWaterMarkObject", shape_id + " " + type_ref, re.I))
                behind_doc = is_absolute and (z_neg or name_like_wm)
                for tp in shape.iter(qn("v", "textpath")):
                    s = tp.get("string", "") or tp.get(qn("o", "title"), "")
                    if s:
                        tp_style = tp.get("style", "") or ""
                        font = tp.get("font-family", "") or ""
                        if not font:
                            m_font = re.search(r"font-family\s*:\s*([^;]+)", tp_style, re.I)
                            if m_font: font = m_font.group(1).strip(" '\"")
                        size_pt = None
                        m_size = re.search(r"font-size\s*:\s*(\d+(?:\.\d+)?)\s*pt", tp_style, re.I)
                        if m_size:
                            try: size_pt = float(m_size.group(1))
                            except Exception: size_pt = None
                        self.watermark_candidates.append({
                            "part": part, "type": "vml_textpath", "text": s, "confidence": "高",
                            "font": font, "size_pt": size_pt, "color": fillcolor or strokecolor,
                            "opacity": opacity, "rotation": rotation,
                            "in_header": in_header, "in_footer": in_footer,
                            "behind_doc": behind_doc, "is_absolute": is_absolute,
                            "shape_id": shape_id, "shape_type": type_ref,
                        })
            # 兼容没有 shape 外层信息的 VML textpath
            for tp in root.iter(qn("v", "textpath")):
                s = tp.get("string", "") or tp.get(qn("o", "title"), "")
                if s and not any(c.get("part") == part and c.get("type") == "vml_textpath" and c.get("text") == s for c in self.watermark_candidates):
                    self.watermark_candidates.append({"part":part, "type":"vml_textpath", "text":s, "confidence":"中",
                                                      "in_header": in_header, "in_footer": in_footer, "behind_doc": False})
            # DrawingML anchor 水印（wp:anchor@behindDoc="1"）
            for anchor in root.iter(qn("wp", "anchor")):
                behind = anchor.get("behindDoc", "0") in ("1", "true")
                for at in anchor.iter(qn("a", "t")):
                    if at.text and "菟思学院" in at.text:
                        self.watermark_candidates.append({
                            "part": part, "type": "drawing_anchor", "text": at.text,
                            "confidence": "高" if behind else "中",
                            "in_header": in_header, "in_footer": in_footer,
                            "behind_doc": behind, "is_absolute": True,
                        })
            # DrawingML/shape 中的可编辑文字（inline 或其他，兜底）
            for t in root.iter(qn("a","t")):
                if t.text and "菟思学院" in t.text:
                    # 避免与 drawing_anchor 重复
                    already = any(c.get("part") == part and c.get("text") == t.text and c.get("type") == "drawing_anchor"
                                  for c in self.watermark_candidates)
                    if not already:
                        self.watermark_candidates.append({
                            "part": part, "type": "drawing_text", "text": t.text, "confidence": "中",
                            "in_header": in_header, "in_footer": in_footer, "behind_doc": False,
                        })

    def _extract_risks(self, root):
        self.drawing_count = len(list(root.iter(qn("w","drawing")))) + len(list(root.iter(qn("v","imagedata"))))
        self.comment_count = 1 if self.pkg.has("word/comments.xml") else 0
        self.revision_count = len(list(root.iter(qn("w","ins")))) + len(list(root.iter(qn("w","del"))))
        text = _text_in_elem(root)
        # ① 大面积乱码：异常字符占比超过 0.5%
        bad_chars = sum(text.count(c) for c in "�□■")
        if len(text) and bad_chars / max(len(text), 1) > 0.005:
            self.visual_risks.append(f"大面积乱码({bad_chars}/{len(text)})")
        # ② 逐字竖排：出现 Word 竖排方向标记
        if any(td is not None for td in root.iter(qn("w", "textDirection"))):
            self.visual_risks.append("逐字竖排")
        # ③ 文字重叠：大量绘图/图片对象或极端字号，作为可自动识别风险信号
        extreme = 0
        for r in root.iter(qn("w", "r")):
            rpr = _find(r, "w:rPr")
            sz = _find(rpr, "w:sz") if rpr is not None else None
            try:
                pt = int(_attr(sz, "w", "val")) / 2.0 if sz is not None else None
            except Exception:
                pt = None
            if pt and (pt < 6 or pt > 80):
                extreme += 1
        if self.drawing_count > 50 or extreme > 20:
            self.visual_risks.append(f"文字重叠风险(绘图对象={self.drawing_count},极端字号run={extreme})")
        # ④ 题目内容超出页面边界：段落文本异常长且无明显分隔，作为自动越界风险信号
        long_question_paras = 0
        for p in root.iter(qn("w", "p")):
            t = _text_in_elem(p).strip()
            if _Q_NUM_RE.match(t) and len(t) > 260:
                long_question_paras += 1
        if long_question_paras:
            self.visual_risks.append(f"题目内容疑似超出页面边界({long_question_paras}段)")


# 已知样例答案（来自打分细则，用于兜底校验）
_SINGLE_SAMPLES = {1:"A", 2:"D", 3:"C", 4:"B", 50:"D", 100:"B"}
_JUDGE_SAMPLES = {
    1:"√",2:"√",3:"√",4:"×",5:"×",
    51:"√",52:"√",53:"√",54:"×",55:"×",
    96:"√",97:"√",98:"√",99:"×",100:"×",
    401:"√",402:"√",403:"√",404:"×",405:"×",
    496:"√",497:"√",498:"√",499:"×",500:"×",
    501:"√",502:"√",503:"√",504:"×",505:"×",
    506:"√",507:"√",508:"√",
}

# ─── 题目/答案提取 ────────────────────────────────────────────────────────────
_Q_NUM_RE = re.compile(r"^\s*(\d{1,3})[\.．、\s]")
_FULL_PAREN_RE = re.compile(r"（([^（）]*)）")
_SINGLE_ANS_RE = re.compile(r"^\s*([A-Da-d])\s*$")
_JUDGE_ANS_RE  = re.compile(r"^\s*([√×])\s*$")
_OPT_RE = re.compile(r"^([A-D])[\.．、\s]")

class QuestionExtractor:
    def __init__(self, parser: DocxParser):
        self.parser = parser
        self.single_questions: list[Question] = []
        self.judge_questions: list[Question] = []
        self.single_answers: dict[int,str] = {}   # 从答案表
        self.judge_answers: dict[int,str] = {}    # 从答案表
        self.answer_table_found = False

    def extract(self):
        self._scan_paragraphs()
        self._extract_answer_tables()

    def _scan_paragraphs(self):
        paras = self.parser.paragraphs
        mode = None          # "single" | "judge" | "answer" | None
        seen_single, seen_judge = set(), set()
        for i, p in enumerate(paras):
            t = p.text.strip()
            # 切换题型上下文
            if re.search(r"一[、．.]*\s*单选题", t):
                mode = "single"; continue
            if re.search(r"二[、．.]*\s*判断题", t):
                mode = "judge"; continue
            if re.search(r"答案|参考答案", t):
                mode = "answer"; continue
            if mode in ("single","judge"):
                m = _Q_NUM_RE.match(t)
                if m:
                    num = int(m.group(1))
                    ans, ok, ans_runs, full_text, ans_para_idx = self._find_answer_with_continuation(paras, i, mode)
                    opts = self._find_options_after(paras, i) if mode == "single" else []
                    q = Question(kind=mode, number=num, paragraph_index=i, text=full_text,
                                 answer=ans, answer_in_tail_paren=ok, answer_runs=ans_runs,
                                 answer_paragraph_index=ans_para_idx, options=opts)
                    if mode == "single":
                        if num not in seen_single:
                            self.single_questions.append(q); seen_single.add(num)
                    else:
                        if num not in seen_judge:
                            self.judge_questions.append(q); seen_judge.add(num)

    def _find_options_after(self, paras, idx):
        opts = []
        for j in range(idx + 1, min(len(paras), idx + 10)):
            nt = paras[j].text.strip()
            if not nt or nt == "菟思学院":
                continue
            if _Q_NUM_RE.match(nt):
                break
            if _OPT_RE.match(nt):
                opts.append(nt)
                if len(opts) == 4:
                    break
        return opts

    def _find_answer_with_continuation(self, paras, idx, mode):
        """题干可能被页眉或自动换行拆成多个段落；向后拼到答案括号或下个题号/选项。"""
        p = paras[idx]
        text_parts = [p.text]
        ans_runs = []
        # 先检查当前段落
        ans, ok, runs = self._find_answer_in_text(p.text, p.runs, mode)
        if ok: return ans, ok, runs, p.text, idx
        for j in range(idx+1, min(len(paras), idx+6)):
            nt = paras[j].text.strip()
            if not nt or nt == "菟思学院":
                continue
            # 单选题在选项前通常必须已经出现答案；遇到 A 选项停止
            if _OPT_RE.match(nt): break
            # 下一个题号出现则停止
            if _Q_NUM_RE.match(nt): break
            text_parts.append(paras[j].text)
            ans, ok, runs = self._find_answer_in_text(paras[j].text, paras[j].runs, mode)
            if ok: return ans, ok, runs, "".join(text_parts), j
        return ans, False, ans_runs, "".join(text_parts), None

    def _find_answer_in_text(self, t: str, runs, mode):
        """从段落末尾全角括号中提取答案，并定位答案字符所在的 run。"""
        matches = list(_FULL_PAREN_RE.finditer(t))
        if not matches:
            return "", False, []
        last = matches[-1]
        inner = last.group(1).strip()
        after = t[last.end():].strip()
        if after and not re.match(r"^[\s\n\t。.，,]*$", after):
            return inner, False, []
        if mode == "single" and not _SINGLE_ANS_RE.match(inner):
            return inner, False, []
        if mode == "judge" and not _JUDGE_ANS_RE.match(inner):
            return inner, False, []
        answer = inner.upper() if inner.upper() in "ABCD" else inner
        ans_start = -1
        for cand in {inner, inner.upper(), inner.lower(), answer}:
            if cand:
                ans_start = t.find(cand, last.start(1), last.end(1))
                if ans_start >= 0: break
        ans_end = ans_start + len(answer) if ans_start >= 0 else -1
        ans_runs, pos = [], 0
        for r in runs:
            start, end = pos, pos + len(r.text)
            if ans_start >= 0 and start < ans_end and end > ans_start:
                ans_runs.append(r)
            pos = end
        return answer, True, ans_runs

    def _extract_answer_tables(self):
        """从正文中寻找答案表(表格或连续段落)"""
        root = self.parser.pkg.parse_xml("word/document.xml")
        if root is None: return
        body = _find(root, "w:body")
        if body is None: body = root
        single_kv, judge_kv = {}, {}
        for tbl in body.findall(".//w:tbl", NS):
            cells = [_text_in_elem(c).strip() for row in tbl.findall(".//w:tr", NS)
                     for c in row.findall(".//w:tc", NS)]
            text = " ".join(cells)
            # 判断是否答案表
            if re.search(r"题号|答案|单选|判断", text[:100]):
                self._parse_answer_text(text, single_kv, judge_kv)
        # 也从段落文本中找连续答案行
        body_text = "\n".join(p.text for p in self.parser.paragraphs)
        # 找到"单选题答案"或"答案表"后的区域
        for section in re.split(r"单选题?答案|单选答案", body_text):
            if len(single_kv) >= 510: break
            self._parse_answer_text(section[:2000], single_kv, {})
        for section in re.split(r"判断题?答案|判断答案", body_text):
            if len(judge_kv) >= 508: break
            self._parse_answer_text(section[:2000], {}, judge_kv)
        self.single_answers = single_kv
        self.judge_answers = judge_kv
        self.answer_table_found = bool(single_kv or judge_kv)

    def _parse_answer_text(self, text, single_kv, judge_kv):
        # 匹配 "1.A" "1、A" "1 A" "1：A" 等，包括判断 √/×
        for m in re.finditer(r"(\d{1,3})\s*[、.．:：]?\s*([A-Da-d√×])", text):
            num, ans = int(m.group(1)), m.group(2)
            if ans.upper() in "ABCD": single_kv.setdefault(num, ans.upper())
            elif ans in "√×": judge_kv.setdefault(num, ans)


# ─── 辅助函数 ─────────────────────────────────────────────────────────────────
def _is_songti(font):
    if not font: return None
    ok = re.search(r"宋体|SimSun|Sungti|Songti|FangSong|仿宋|等线|Microsoft YaHei|微软雅黑|楷体|KaiTi", font, re.I)
    return bool(ok)

def _color_is_light_gray(c):
    if not c or c.lower() in ("auto",): return None
    try:
        r = int(c[0:2],16); g = int(c[2:4],16); b = int(c[4:6],16)
        return r > 150 and g > 150 and b > 150 and max(r,g,b) - min(r,g,b) < 30
    except Exception: return None

def _normalize_hex_color(c):
    if not c or c.lower() in ("auto", "none"):
        return None
    c = c.strip().lstrip("#")
    names = {"black": "000000", "gray": "808080", "grey": "808080", "darkgray": "404040", "darkgrey": "404040"}
    c = names.get(c.lower(), c)
    return c.upper() if re.fullmatch(r"[0-9A-Fa-f]{6}", c) else None

def _opacity_is_transparent(v):
    if not v:
        return None
    s = str(v).strip().lower()
    try:
        if s.endswith("%"):
            return float(s[:-1]) < 100
        if s.endswith("f"):
            return int(s[:-1], 16) < 0x10000
        return float(s) < 1
    except Exception:
        return None

def _color_is_dark(c):
    c = _normalize_hex_color(c)
    if not c: return None
    try:
        r = int(c[0:2],16); g = int(c[2:4],16); b = int(c[4:6],16)
        return max(r,g,b) < 100
    except Exception: return None

def _color_is_red(c):
    if not c: return False
    try:
        r = int(c[0:2],16); g = int(c[2:4],16); b = int(c[4:6],16)
        return r > 180 and g < 100 and b < 100
    except Exception: return False

def _color_is_plain(c):
    """答案文字应为默认/黑色系；明显彩色或浅色视为特殊颜色。"""
    if not c or c.lower() in ("auto",): return True
    return bool(_color_is_dark(c))

def _is_single_line_spacing(p):
    """单倍行距：lineRule=auto 且 line 约 240 twips(11.5-14.5pt)；或 rule/val 均缺失（继承默认）。"""
    if p.line_spacing_rule is None and p.line_spacing_val is None: return True
    if p.line_spacing_rule in (None, "auto"):
        return p.line_spacing_val is None or 11.5 <= p.line_spacing_val <= 14.5
    return False

def R(rule_id, title, score, passed, evidence, confidence="高", fatal=False):
    return RuleResult(rule_id=rule_id, title=title, score=score, passed=passed,
                      evidence=evidence, confidence=confidence, fatal=fatal)


# ─── 维度1门槛检查 ───────────────────────────────────────────────────────────
class GateChecker:
    def __init__(self, pkg: DocxPackage, parser: DocxParser, extractor: QuestionExtractor):
        self.pkg = pkg; self.parser = parser; self.ext = extractor

    def check(self) -> list[RuleResult]:
        results = []
        # D1-1 文件格式与可解析性（唯一保留的门槛项；其余门槛/扣分项按 rubric 修订已删除）
        if not self.pkg.has("word/document.xml"):
            return [R("D1-1","文件格式/可解析性",0,False,"缺少 word/document.xml",fatal=True)]
        results.append(R("D1-1","文件格式/可解析性",0,True,"ZIP结构正常，document.xml 可解析",fatal=True))
        return results


# ─── 维度2评分 ───────────────────────────────────────────────────────────────
class ScoreChecker:
    def __init__(self, pkg: DocxPackage, parser: DocxParser, extractor: QuestionExtractor):
        self.pkg = pkg; self.parser = parser; self.ext = extractor
        self.singles = {q.number:q for q in extractor.single_questions}
        self.judges  = {q.number:q for q in extractor.judge_questions}

    def score(self) -> list[RuleResult]:
        rs = []
        rs += [self._structure(), self._editable_objects(), self._title_type_format(),
               self._header_format(), self._watermark_format(), self._body_format(),
               self._single_complete(), self._single_position(), self._single_format(),
               self._single_correct(), self._judge_complete(), self._judge_position(),
               self._judge_format(), self._judge_correct()]
        rs += self._deductions()
        return rs

    def _pages_ok(self) -> tuple[bool, dict[str, int | None]]:
        pg = self.parser.page_signals
        vals: list[int] = [v for v in (pg.get("app_pages"), pg.get("last_rendered_pages")) if v is not None]
        explicit = pg.get("explicit_page_breaks")
        if explicit is not None:
            vals.append(explicit + 1)
        return any(88 <= v <= 104 for v in vals), pg

    def _structure(self):
        pages_ok, pg = self._pages_ok()
        title_ok = "多岗位从业人员综合复习题" in self.parser.all_text
        single_title_ok = any(re.search(r"一[、．.]\s*单选题", p.text.strip()) for p in self.parser.paragraphs)
        judge_title_ok  = any(re.search(r"二[、．.]\s*判断题",  p.text.strip()) for p in self.parser.paragraphs)
        single_count_ok = len(self.ext.single_questions) == 510
        judge_count_ok  = len(self.ext.judge_questions)  == 508
        single_ans_ok   = len(self.ext.single_answers)   == 510
        judge_ans_ok    = len(self.ext.judge_answers)    == 508
        ok = (pages_ok and title_ok and single_title_ok and judge_title_ok and
              single_count_ok and judge_count_ok and single_ans_ok and judge_ans_ok)
        return R("D2+5-STRUCT","文档完整保留96页左右内容结构",5,ok,
                 f"页数={pages_ok}，标题={title_ok}，一单选={single_title_ok}，二判断={judge_title_ok}，单选={len(self.ext.single_questions)}/510，判断={len(self.ext.judge_questions)}/508，单选答案页={len(self.ext.single_answers)}/510，判断答案页={len(self.ext.judge_answers)}/508，页数信号={pg}","中")

    def _editable_objects(self):
        pages_ok, pg = self._pages_ok()
        single_qs = self.ext.single_questions
        judge_qs = self.ext.judge_questions
        question_numbers_ok = len(single_qs) == 510 and len(judge_qs) == 508
        question_stems_ok = all(q.text.strip() for q in single_qs + judge_qs)
        option_counts = {letter: 0 for letter in "ABCD"}
        for q in single_qs:
            seen = {opt[0].upper() for opt in q.options if _OPT_RE.match(opt)}
            for letter in option_counts:
                if letter in seen:
                    option_counts[letter] += 1
        options_ok = all(option_counts[letter] == 510 for letter in "ABCD")
        single_answers_ok = sum(1 for q in single_qs if _SINGLE_ANS_RE.match(q.answer or "")) == 510
        judge_answers_ok = sum(1 for q in judge_qs if _JUDGE_ANS_RE.match(q.answer or "")) == 508

        # ── 页眉可编辑性：不仅要包含"菟思学院"，还须来自 <w:t> 文本层（非图片/DrawingML 光栅）
        header_ok = False
        header_text_source_ok = False
        referenced = self.parser.referenced_header_parts or self.pkg.header_parts()
        for part in referenced:
            root = self.pkg.parse_xml(part)
            if root is None: continue
            # 直接 w:t 文本中出现"菟思学院"
            wt_text = "".join((t.text or "") for t in root.iter(qn("w", "t")))
            if "菟思学院" in wt_text:
                header_ok = True
                header_text_source_ok = True
                break
        if not header_ok:
            header_ok = any("菟思学院" in p.text for p in self.parser.header_paragraphs)

        # ── 水印可编辑性：VML v:textpath@string 或 w:t 承载"菟思学院"，且不是纯图片水印
        wm_editable = False
        wm_present = False
        for c in self.parser.watermark_candidates:
            if "菟思学院" not in c.get("text", ""):
                continue
            wm_present = True
            # vml_textpath / drawing_text 均属于矢量可编辑文字层
            if c.get("type") in ("vml_textpath", "drawing_text"):
                wm_editable = True

        # ── 逐页逐类对象覆盖检测（不用 COM，基于 lastRenderedPageBreak / 分页符 / 分节符）
        page_info = self.parser.page_editability or {}
        _dp = page_info.get("detected_pages", 0)
        detected_pages: int = _dp if isinstance(_dp, int) else 0
        _pmc = page_info.get("page_map_confidence", "低")
        page_map_conf: str = _pmc if isinstance(_pmc, str) else "低"
        _fp = page_info.get("failed_pages", [])
        failed_pages: list[object] = _fp if isinstance(_fp, list) else []
        _pp = page_info.get("per_page_stats", {})
        per_page: dict[object, object] = _pp if isinstance(_pp, dict) else {}

        # 目标覆盖：第 1—96 页必须每页都至少出现"题号 / 题干"文字对象来自 w:t，
        # 且不允许任一页被判定为整页图片化或关键对象非文字层。
        target_pages = range(1, 97)
        missing_pages: list[int] = []
        image_only_pages: list[int] = []
        for pg_no in target_pages:
            s = per_page.get(pg_no)
            if not isinstance(s, dict):
                # 检测到的总页数不足 96 时，若 lastRenderedPageBreak 缺失只能标为低置信度而非直接判失败
                if page_map_conf == "高":
                    missing_pages.append(pg_no)
                continue
            # 该页文字层字符数 < 20 且存在大幅图片：视为整页图片化
            text_chars = s.get("text_chars", 0)
            ratio = s.get("max_image_area_ratio", 0.0)
            if isinstance(text_chars, int) and isinstance(ratio, (int, float)) and text_chars < 20 and ratio >= 0.6:
                image_only_pages.append(pg_no)
        # 从 failed_pages 中过滤出 1—96 页的失败原因
        target_failed: list[dict] = []
        for fp in failed_pages:
            if isinstance(fp, dict):
                pg_val = fp.get("page", 0)
                if isinstance(pg_val, int) and 1 <= pg_val <= 96:
                    target_failed.append(fp)

        # 综合判定：
        # - 分页置信度高时，要求 1—96 页无一页被判失败、无缺失
        # - 分页置信度低时，只能退化为"整篇不存在整页图片化"的宽松判定，并在证据里标注
        if page_map_conf == "高":
            no_full_page_image_ok = not target_failed and not missing_pages
        else:
            has_image_fail = False
            for fp in failed_pages:
                if not isinstance(fp, dict):
                    continue
                pg_val = fp.get("page", 0)
                reasons = fp.get("reasons", [])
                if not (isinstance(pg_val, int) and pg_val <= detected_pages):
                    continue
                if isinstance(reasons, list) and any(isinstance(r, str) and "整页图片化" in r for r in reasons):
                    has_image_fail = True
                    break
            no_full_page_image_ok = not has_image_fail

        ok = (pages_ok and question_numbers_ok and question_stems_ok and options_ok and
              single_answers_ok and judge_answers_ok and header_ok and header_text_source_ok and
              wm_present and wm_editable and no_full_page_image_ok)
        evidence = (
            f"页数={pages_ok}，题号=单选{len(single_qs)}/510+判断{len(judge_qs)}/508，题干={question_stems_ok}，"
            f"A-D选项={option_counts}，单选答案字母={sum(1 for q in single_qs if _SINGLE_ANS_RE.match(q.answer or ''))}/510，"
            f"判断答案符号={sum(1 for q in judge_qs if _JUDGE_ANS_RE.match(q.answer or ''))}/508，"
            f"菟思学院页眉={header_ok}(文字层={header_text_source_ok})，"
            f"菟思学院水印存在={wm_present},可编辑={wm_editable}，"
            f"分页可编辑性检测={{'检测页数':{detected_pages},'置信度':'{page_map_conf}',"
            f"'1-96页失败':{target_failed[:5]},'缺失页':{missing_pages[:5]},'疑似整页图片化':{image_only_pages[:5]}}}，"
            f"图片={self.parser.image_count}，绘图={self.parser.drawing_count}，页数信号={pg}"
        )
        return R("D2+5-EDITABLE", "第1—96页正文对象、页眉和水印均可编辑", 5, ok, evidence,
                 "高" if page_map_conf == "高" else "中")

    def _spacing_half_line_ok(self, p, attr):
        val = getattr(p, attr)
        if val is not None:
            return 0.45 <= val <= 0.55
        pt_attr = "before_pt" if attr == "before_lines" else "after_pt"
        pt = getattr(p, pt_attr)
        return pt is not None and 5.5 <= pt <= 7.5

    def _page_of_paragraph(self, p_obj):
        """定位 body 段落对象在文档中的页号；aligned_page_of_paragraph 未覆盖时返回 None。"""
        pages = getattr(self.parser, "aligned_page_of_paragraph", []) or []
        for i, x in enumerate(self.parser.paragraphs):
            if x is p_obj:
                return pages[i] if i < len(pages) else None
        return None

    def _title_type_format(self):
        title_p = next((p for p in self.parser.paragraphs if "多岗位从业人员综合复习题" in p.text), None)
        single_p = next((p for p in self.parser.paragraphs if re.search(r"一[、．.]*\s*单选题", p.text)), None)
        judge_p = next((p for p in self.parser.paragraphs if re.search(r"二[、．.]*\s*判断题", p.text)), None)
        issues = []

        # 分页置信度：低置信度时"位于第1页"退化为"未见其他分页锚点在其之前"
        page_map_conf = getattr(self.parser, "page_map_confidence", "低")

        def run_fonts(p):
            # 保留 None，以便区分"显式非宋体"和"未设置(继承)"
            return [r.font for r in p.runs]

        def run_sizes(p):
            return [r.size_pt for r in p.runs]

        def fonts_ok(fonts):
            """所有显式设置的字体都必须是宋体或相近字体；全部未设置则视为继承默认，通过。"""
            explicit = [f for f in fonts if f]
            if not explicit:
                return True
            return all(_is_songti(f) for f in explicit)

        def sizes_ok(sizes, lo, hi):
            explicit = [s for s in sizes if s is not None]
            if not explicit:
                return True  # 继承默认字号也算通过
            return all(lo <= s <= hi for s in explicit)

        def title_check(p):
            if not p:
                issues.append("标题不存在")
                return {"存在": False}
            fonts = run_fonts(p)
            sizes = run_sizes(p)
            bolds = [r.bold for r in p.runs if r.bold is not None]
            # 位于第 1 页
            page = self._page_of_paragraph(p)
            if page is None:
                on_page1 = True  # 未能定位到页时不阻断
            else:
                on_page1 = page == 1
            detail = {
                "存在": True,
                "位于第1页": on_page1,
                "宋体或相近字体": fonts_ok(fonts),
                "字号23-24磅": sizes_ok(sizes, 23, 24),
                "不加粗": not any(bolds),
                "居中对齐": p.alignment == "center",
                "标题下方空一行": p.after_lines == 1 or (p.after_pt is not None and 11.5 <= p.after_pt <= 14.5),
                "fonts": fonts,
                "sizes": sizes,
                "alignment": p.alignment,
                "after_lines": p.after_lines,
                "after_pt": p.after_pt,
                "page": page,
                "page_map_confidence": page_map_conf,
            }
            for key in ["位于第1页", "宋体或相近字体", "字号23-24磅", "不加粗", "居中对齐", "标题下方空一行"]:
                if not detail[key]:
                    issues.append(f"标题{key}不符合(page={page},置信度={page_map_conf})" if key == "位于第1页"
                                  else f"标题{key}不符合")
            return detail

        def type_check(p, label):
            if not p:
                issues.append(f"{label}不存在")
                return {"存在": False}
            fonts = run_fonts(p)
            sizes = run_sizes(p)
            detail = {
                "存在": True,
                "左对齐": p.alignment in (None, "left"),
                "宋体或相近字体": fonts_ok(fonts),
                "字号15-17磅": sizes_ok(sizes, 15, 17),
                "段前0.5行": self._spacing_half_line_ok(p, "before_lines"),
                "段后0.5行": self._spacing_half_line_ok(p, "after_lines"),
                "fonts": fonts,
                "sizes": sizes,
                "alignment": p.alignment,
                "before_lines": p.before_lines,
                "before_pt": p.before_pt,
                "after_lines": p.after_lines,
                "after_pt": p.after_pt,
            }
            for key in ["左对齐", "宋体或相近字体", "字号15-17磅", "段前0.5行", "段后0.5行"]:
                if not detail[key]:
                    issues.append(f"{label}{key}不符合")
            return detail

        title_detail = title_check(title_p)
        single_detail = type_check(single_p, "一、单选题")
        judge_detail = type_check(judge_p, "二、判断题")
        title_ok = all(title_detail.get(k) for k in ["存在", "位于第1页", "宋体或相近字体", "字号23-24磅", "不加粗", "居中对齐", "标题下方空一行"])
        single_ok = all(single_detail.get(k) for k in ["存在", "左对齐", "宋体或相近字体", "字号15-17磅", "段前0.5行", "段后0.5行"])
        judge_ok = all(judge_detail.get(k) for k in ["存在", "左对齐", "宋体或相近字体", "字号15-17磅", "段前0.5行", "段后0.5行"])
        ok = title_ok and single_ok and judge_ok
        return R("D2+3-TITLEFMT", "第1页标题与题型标题格式", 3, ok,
                 f"标题={title_detail}；一、单选题={single_detail}；二、判断题={judge_detail}；问题={issues}",
                 "高" if page_map_conf == "高" else "中")

    def _section_page_headers(self):
        """遍历文档所有 section 的 default/first/even 页眉引用，返回每页应使用的页眉信息。

        返回 (page_headers, sections_meta)：
        - page_headers: {page_no: {'expected_type': 'default'|'first'|'even', 'part': str|None, 'has_text': bool}}
        - sections_meta: list[dict]，逐节的 titlePg / evenAndOddHeaders / 页范围 / 引用
        """
        pkg = self.parser.pkg
        rels = pkg.parse_rels()
        # 全局设置：settings.xml 里的 evenAndOddHeaders
        settings_root = pkg.parse_xml("word/settings.xml")
        even_odd_global = False
        if settings_root is not None:
            eo = _find(settings_root, "w:evenAndOddHeaders")
            if eo is not None:
                even_odd_global = _attr(eo, "w", "val", "true") not in ("false", "0", "off")

        doc_root = pkg.parse_xml("word/document.xml")
        if doc_root is None:
            return {}, []
        body = _find(doc_root, "w:body") or doc_root

        # 收集所有 sectPr（包括段落 pPr 内的 sectPr 以及 body 末尾的 sectPr），按文档顺序
        # 与 _extract_body_structure 的段落顺序对齐，从而得到每节的起止页
        sect_boundaries = []  # list of {"sectPr": Element, "end_page": int}
        cur_page = 1
        pending_new_page = False
        para_idx = 0
        para_pages = self.parser.body_page_of_para or []
        for el in list(body):
            tag = el.tag
            if tag == qn("w", "p"):
                if pending_new_page:
                    cur_page += 1
                    pending_new_page = False
                if para_idx < len(para_pages):
                    cur_page = para_pages[para_idx]
                para_idx += 1
                # 段落 pPr 内嵌 sectPr（本段末尾结束一个 section）
                ppr = _find(el, "w:pPr")
                inner_sect = _find(ppr, "w:sectPr") if ppr is not None else None
                if inner_sect is not None:
                    sect_boundaries.append({"sectPr": inner_sect, "end_page": cur_page})
                for lrpb in el.iter(qn("w", "lastRenderedPageBreak")):
                    pending_new_page = True
                for br in el.iter(qn("w", "br")):
                    if _attr(br, "w", "type") == "page":
                        pending_new_page = True
            elif tag == qn("w", "sectPr"):
                sect_boundaries.append({"sectPr": el, "end_page": cur_page})

        # 若最后一节的 sectPr 没被识别（罕见），兜底：把 body 剩余页归入最后一节
        detected_pages = max(para_pages) if para_pages else 0
        if sect_boundaries and sect_boundaries[-1]["end_page"] < detected_pages:
            sect_boundaries[-1]["end_page"] = detected_pages

        # 计算每节起止页
        sections_meta: list[dict] = []
        start_page = 1
        for sb in sect_boundaries:
            end_page = max(sb["end_page"], start_page)
            sectPr = sb["sectPr"]
            title_pg = _find(sectPr, "w:titlePg") is not None
            refs = {}
            for ref in sectPr.iter(qn("w", "headerReference")):
                rtype = _attr(ref, "w", "type", "default") or "default"
                rid = _attr(ref, "r", "id")
                target = rels.get(rid, "") if rid else ""
                if not target: continue
                if target.startswith("/"): part = target.lstrip("/")
                elif target.startswith("word/"): part = target
                else: part = "word/" + target
                if pkg.has(part):
                    refs[rtype] = part
            sections_meta.append({
                "start_page": start_page, "end_page": end_page,
                "titlePg": title_pg, "headers": refs,
            })
            start_page = end_page + 1

        # 预读各 header part 里是否包含"菟思学院"（w:t 文本层）
        header_has_text: dict[str, bool] = {}
        for sect in sections_meta:
            for part in sect["headers"].values():
                if part in header_has_text: continue
                root = pkg.parse_xml(part)
                if root is None:
                    header_has_text[part] = False
                else:
                    wt = "".join((t.text or "") for t in root.iter(qn("w", "t")))
                    header_has_text[part] = "菟思学院" in wt

        # 为每页决定使用哪个 header：first (titlePg + 节首页) / even (evenAndOddHeaders + 偶数页) / default
        page_headers: dict[int, dict] = {}
        for sect in sections_meta:
            hdrs = sect["headers"]
            for pg in range(sect["start_page"], sect["end_page"] + 1):
                is_first = sect["titlePg"] and pg == sect["start_page"]
                is_even = even_odd_global and (pg % 2 == 0)
                if is_first and "first" in hdrs:
                    expected_type, part = "first", hdrs["first"]
                elif is_even and "even" in hdrs:
                    expected_type, part = "even", hdrs["even"]
                elif "default" in hdrs:
                    expected_type, part = "default", hdrs["default"]
                else:
                    # 该节没引用任何 header：视为该页无页眉
                    expected_type, part = "missing", None
                page_headers[pg] = {
                    "section_start": sect["start_page"],
                    "expected_type": expected_type,
                    "part": part,
                    "has_text": bool(part) and header_has_text.get(part, False),
                }
        return page_headers, sections_meta

    def _header_format(self):
        # 细则：每页顶部都需要出现"菟思学院"文字；且字体/字号/对齐/颜色需符合
        page_headers, sections_meta = self._section_page_headers()
        header_parts = self.pkg.header_parts()
        # ── 每页顶部覆盖：逐页检查是否引用到含"菟思学院"文字的页眉 part
        page_map_conf = getattr(self.parser, "page_map_confidence", "低")
        target_pages = list(page_headers.keys())
        pages_missing: list[int] = [pg for pg, info in page_headers.items() if not info["has_text"]]
        # 若分页置信度高，覆盖判定必须每页命中；否则退化为"每个 section 的 default 页眉均含菟思学院"
        if page_map_conf == "高":
            present_ok = bool(page_headers) and not pages_missing
        else:
            default_parts = [s["headers"].get("default") for s in sections_meta if s["headers"].get("default")]
            present_ok = bool(default_parts) and all(
                self.pkg.parse_xml(part) is not None and
                "菟思学院" in "".join((t.text or "") for t in self.pkg.parse_xml(part).iter(qn("w", "t")))
                for part in default_parts
            )

        # ── 字体/字号/对齐/颜色：以命中的页眉段落做为样本，若多种页眉都存在，则所有涉及"菟思学院"的段落都要通过
        hit_paras = [p for p in self.parser.header_paragraphs if "菟思学院" in p.text]
        font_ok = size_ok = align_ok = color_ok = False
        fmt_details: list[dict] = []
        if hit_paras:
            font_ok = size_ok = align_ok = color_ok = True
            for p in hit_paras:
                runs = p.runs or []
                # 只取包含"菟思学院"的 run 及其相邻 run 参与格式判断（简化处理：整段所有 run）
                fonts_all = [r.font for r in runs]
                sizes_all = [r.size_pt for r in runs]
                colors_all = [r.color for r in runs]
                explicit_fonts = [f for f in fonts_all if f]
                explicit_sizes = [s for s in sizes_all if s is not None]
                explicit_colors = [c for c in colors_all if c]
                # 字体：未显式设置视为继承默认，通过；显式设置则必须宋体或相近
                p_font_ok = not explicit_fonts or all(_is_songti(f) for f in explicit_fonts)
                # 字号 10-11 磅：未显式时不能确定，视为通过（继承）
                p_size_ok = not explicit_sizes or all(10 <= s <= 11 for s in explicit_sizes)
                # 左对齐：None/left/both 通过
                p_align_ok = p.alignment in (None, "left", "both")
                # 浅灰色：必须至少显式一次浅灰色
                p_color_ok = bool(explicit_colors) and any(_color_is_light_gray(c) for c in explicit_colors)
                fmt_details.append({
                    "part": p.part, "text": p.text[:30],
                    "font_ok": p_font_ok, "size_ok": p_size_ok, "align_ok": p_align_ok, "color_ok": p_color_ok,
                    "fonts": explicit_fonts, "sizes": explicit_sizes, "alignment": p.alignment, "colors": explicit_colors,
                })
                font_ok = font_ok and p_font_ok
                size_ok = size_ok and p_size_ok
                align_ok = align_ok and p_align_ok
                color_ok = color_ok and p_color_ok
        ok = present_ok and font_ok and size_ok and align_ok and color_ok
        section_summary = [
            {"pages": f"{s['start_page']}-{s['end_page']}", "titlePg": s["titlePg"],
             "headers": {k: v for k, v in s["headers"].items()}} for s in sections_meta
        ]
        evidence = (
            f"header部件={header_parts}，section数={len(sections_meta)}，section覆盖={section_summary}，"
            f"覆盖页数={len(target_pages)}，未出现菟思学院的页={pages_missing[:10]}(共{len(pages_missing)}页)，"
            f"分页置信度={page_map_conf}，格式检查={fmt_details}"
        )
        return R("D2+3-HEADER", "每页顶部出现菟思学院页眉", 3, ok, evidence,
                 "高" if page_map_conf == "高" else "中")

    def _watermark_format(self):
        # rubric：页面背景中的"菟思学院"水印文字（深灰/黑半透明、倾斜、70-74磅、宋体）
        # 判定要点：
        # ① 必须是"页面背景"水印 —— 位于页眉部件（in_header）、且 behindDoc 或 wp:anchor@behindDoc="1"
        # ② 逐 section 校验：每节的 default 页眉必须引用到至少一个合格候选
        # ③ 分页置信度高时进一步要求每页对应的 header part 都含合格候选
        # ④ 格式（字号/字体/颜色/透明度/旋转）逐候选校验
        all_cands = [c for c in self.parser.watermark_candidates if "菟思学院" in str(c.get("text", ""))]

        def _is_background_wm(c) -> bool:
            in_header = bool(c.get("in_header"))
            t = c.get("type", "")
            behind = bool(c.get("behind_doc"))
            # VML textpath: 需 in_header + behind_doc
            # drawing_anchor: 由 wp:anchor@behindDoc=1 生成时 behind_doc 已为 True
            # drawing_text: 仅表明存在可编辑文字，不足以证明为背景水印 → 不计入
            if t == "vml_textpath": return in_header and behind
            if t == "drawing_anchor": return in_header and behind
            return False

        bg_cands = [c for c in all_cands if _is_background_wm(c)]

        # 逐候选做格式校验
        def _check_format(c):
            color_ok = bool(_color_is_dark(str(c.get("color") or "")))
            transparent_ok = bool(_opacity_is_transparent(str(c.get("opacity") or "")))
            rotation = c.get("rotation")
            try:
                slant_ok = rotation is not None and abs(float(rotation)) >= 10
            except (TypeError, ValueError):
                slant_ok = False
            size = c.get("size_pt")
            try:
                size_val = float(size) if size is not None else None
            except (TypeError, ValueError):
                size_val = None
            size_ok = size_val is not None and 70 <= size_val <= 74
            font = str(c.get("font") or "")
            font_ok = bool(font) and bool(_is_songti(font))
            passed = color_ok and transparent_ok and slant_ok and size_ok and font_ok
            return {
                "text": str(c.get("text") or ""), "part": str(c.get("part") or ""),
                "type": str(c.get("type") or ""),
                "in_header": bool(c.get("in_header")), "behind_doc": bool(c.get("behind_doc")),
                "深灰或黑色": color_ok, "半透明": transparent_ok, "倾斜版式": slant_ok,
                "字号70-74磅": size_ok, "宋体或相近字体": font_ok,
                "color": c.get("color", ""), "opacity": c.get("opacity", ""),
                "rotation": rotation, "size_pt": size, "font": font, "passed": passed,
            }

        checked = [_check_format(c) for c in bg_cands]
        # part → 是否有"格式全部合格"的候选
        parts_with_valid_wm = {c["part"] for c in checked if c["passed"]}
        # part → 是否至少含"菟思学院"背景水印（不论格式）——用于诊断
        parts_with_any_bg = {str(c.get("part") or "") for c in bg_cands}

        # 逐 section 覆盖判定：每节的 default 页眉必须引用到 parts_with_valid_wm 中的 part
        page_headers, sections_meta = self._section_page_headers()
        page_map_conf = getattr(self.parser, "page_map_confidence", "低")
        section_coverage = []
        for s in sections_meta:
            default_part = s["headers"].get("default")
            covered = bool(default_part) and default_part in parts_with_valid_wm
            section_coverage.append({
                "pages": f"{s['start_page']}-{s['end_page']}",
                "default_header": default_part,
                "has_valid_wm": covered,
                "has_any_bg_wm": bool(default_part) and default_part in parts_with_any_bg,
            })
        sections_ok = bool(section_coverage) and all(x["has_valid_wm"] for x in section_coverage)

        # 高置信度：更进一步逐页命中——每页 expected header part 必须在 parts_with_valid_wm 中
        pages_missing_wm: list[int] = []
        if page_map_conf == "高" and page_headers:
            for pg, info in page_headers.items():
                part = info.get("part")
                if not part or part not in parts_with_valid_wm:
                    pages_missing_wm.append(pg)
            coverage_ok = sections_ok and not pages_missing_wm
        else:
            # 低置信度：退化为节级覆盖
            coverage_ok = sections_ok

        # 至少一个候选格式全通过（若节全覆盖，隐含存在）
        any_valid = bool(parts_with_valid_wm)
        ok = coverage_ok and any_valid

        # 若一个背景水印候选都没有，落回"任一候选格式通过"作为最低兜底诊断输出
        fallback_note = ""
        if not bg_cands and all_cands:
            fallback_note = "未识别到页面背景水印（缺少 in_header+behindDoc 信号）"
        elif not all_cands:
            fallback_note = "未检测到菟思学院文字水印"

        evidence = (
            f"背景水印候选数={len(bg_cands)}，格式合格 part={sorted(parts_with_valid_wm)}，"
            f"section覆盖={section_coverage}，分页置信度={page_map_conf}，"
            f"高置信度下未覆盖水印的页={pages_missing_wm[:10]}(共{len(pages_missing_wm)}页)，"
            f"候选详情={checked[:3]}" + (f"，备注={fallback_note}" if fallback_note else "")
        )
        return R("D2+3-WATERMARK", "页面背景菟思学院水印文字格式", 3, ok, evidence,
                 "高" if page_map_conf == "高" else "中")

    def _body_format(self):
        # 细则：宋体或其他相近字体、左对齐、13-14磅、单倍行距
        # 取全部单选题+判断题题干段落进行检查
        qparas = [self.parser.paragraphs[q.paragraph_index]
                  for q in (self.ext.single_questions + self.ext.judge_questions)
                  if 0 <= q.paragraph_index < len(self.parser.paragraphs)]
        if not qparas:
            return R("D2+3-BODYFMT", "正文对应格式", 3, False, "未找到题目段落", "低")
        issues_map: dict[str, list] = {"非宋体/相近字体": [], "非左对齐": [], "字号非13-14磅": [], "非单倍行距": []}
        for p in qparas:
            # ① 字体：宋体或相近字体
            fonts = [r.font for r in p.runs if r.font]
            if fonts and not all(_is_songti(f) for f in fonts):
                issues_map["非宋体/相近字体"].append(p.text[:20])
            # ② 左对齐（None 表示未显式设置，默认左对齐，视为通过）
            if p.alignment not in (None, "left", "both"):
                issues_map["非左对齐"].append(p.text[:20])
            # ③ 字号 13-14 磅
            sizes = [r.size_pt for r in p.runs if r.size_pt]
            if sizes and not all(13 <= s <= 14 for s in sizes):
                issues_map["字号非13-14磅"].append(p.text[:20])
            # ④ 单倍行距：lineRule=auto 且 line≈240 twips(=12pt)；或 lineRule=exact/atLeast 且 line 在 13-14pt；
            #    或 lineRule/line 均缺失时视为通过（继承默认单倍行距）
            rule = p.line_spacing_rule
            val  = p.line_spacing_val
            if rule is not None or val is not None:
                if rule in (None, "auto"):
                    # auto 模式：line 值含义为"每行高度/240"，240=单倍，允许 ±0.5pt(10 twips)
                    # line_spacing_val 已转换为 pt，单倍行距 ≈ 12pt（240/20）
                    single_ok = (val is None or 11.5 <= val <= 14.5)
                else:
                    # exact/atLeast 模式：line 值直接为磅数
                    single_ok = (val is None or 13 <= val <= 14)
                if not single_ok:
                    issues_map["非单倍行距"].append(f"{p.text[:20]}(rule={rule},val={val})")
        total = len(qparas)
        bad_counts = {k: len(v) for k, v in issues_map.items()}
        ok = all(cnt == 0 for cnt in bad_counts.values())
        evidence_parts = [f"{k}={cnt}道" for k, cnt in bad_counts.items() if cnt]
        evidence = f"检查题干段落={total}，" + ("全部通过" if ok else "问题：" + "，".join(evidence_parts))
        if not ok:
            for k, v in issues_map.items():
                if v:
                    evidence += f"；{k}样例={v[:3]}"
        return R("D2+3-BODYFMT", "正文对应格式", 3, ok, evidence, "中")

    def _single_complete(self):
        # 细则：510道单选题末尾括号内均填写一个答案字母
        # 括号内不能出现：空白、两个以上字母、"√""×"或其他无关字符
        qs = self.ext.single_questions
        bad = []
        for q in qs:
            ans = q.answer
            # 必须恰好是一个 A-D 字母，无前后空白，不能是空、√×或多字符
            if not (q.answer_in_tail_paren and len(ans) == 1 and ans in "ABCD"):
                reason = []
                if not q.answer_in_tail_paren:
                    reason.append("括号未在末尾或无括号")
                elif ans == "":
                    reason.append("括号内空白")
                elif ans in "√×":
                    reason.append(f"括号内出现判断符号({ans!r})")
                elif len(ans) > 1:
                    reason.append(f"括号内多于一个字符({ans!r})")
                else:
                    reason.append(f"括号内非A-D字母({ans!r})")
                bad.append(f"{q.number}:{','.join(reason)}")
        ok = len(qs) == 510 and not bad
        return R("D2+5-SINGLE-COMPLETE", "单选题1—510题答案完整性", 5, ok,
                 f"合格={510 - len(bad)}/510，检测题数={len(qs)}" + (f"，不合格样例={bad[:5]}" if bad else ""))

    def _single_position(self):
        # 细则：每道单选题答案填写在题干末尾原有全角括号内
        # 显示为（A）（B）（C）或（D），答案字母与左右括号位于同一题目段落内
        qs = self.ext.single_questions
        bad: list[str] = []
        # 保存 rubric 要求"答案 run 应解析完整继承格式并与同题题干/正文基准比较"的诊断
        format_mismatches: list[str] = []
        for q in qs:
            ans = q.answer
            # ① 答案必须在末尾全角括号内
            if not q.answer_in_tail_paren:
                bad.append(f"{q.number}:答案不在末尾全角括号内")
                continue
            # ② 括号内显示为（A/B/C/D）
            if not (len(ans) == 1 and ans in "ABCD"):
                bad.append(f"{q.number}:括号内非（A/B/C/D）({ans!r})")
                continue
            # ③ 答案字母与左右括号位于同一题目段落内
            #    严格要求 answer_paragraph_index == paragraph_index；
            #    若提取器把答案定位到题干后续段落，视为不合格
            if q.answer_paragraph_index is None:
                bad.append(f"{q.number}:无法定位答案所在段落")
                continue
            if q.answer_paragraph_index != q.paragraph_index:
                bad.append(f"{q.number}:答案与题干不在同一段落"
                           f"(题干段={q.paragraph_index},答案段={q.answer_paragraph_index})")
                continue
            # ④ 答案 run 存在——用于后续继承格式解析；缺失说明无法确认，不应通过
            if not q.answer_runs:
                bad.append(f"{q.number}:无法定位答案字符 run")
                continue
            # ⑤ 与"同题题干/正文基准"对比继承后的运行时格式：
            #    以同段落内除答案 run 外的其它 run 作为题干基准；若无则回退到全局正文基准
            baseline = self._paragraph_body_baseline(q)
            mismatch = self._answer_run_inheritance_mismatch(q, baseline)
            if mismatch:
                # 不直接判为 bad（rubric 主要判位置），但记录以便证据可见
                format_mismatches.append(f"{q.number}:{'/'.join(mismatch)}")
        ok = len(qs) == 510 and not bad
        evidence = (
            f"题末全角括号内同段落匹配={510 - len(bad)}/510，检测题数={len(qs)}"
            + (f"，不合格样例={bad[:5]}" if bad else "")
            + (f"，继承格式与题干基准不一致样例={format_mismatches[:5]}(共{len(format_mismatches)}道)" if format_mismatches else "")
        )
        return R("D2+5-SINGLE-POS", "单选题答案所在位置", 5, ok, evidence)

    def _paragraph_body_baseline(self, q):
        """从题干所在段落（q.paragraph_index）里，除答案 run 外的其它 run 取一个格式基准。
        返回 dict: {font, size_pt, bold, italic, underline, highlight, color}
        缺失字段为 None。若同段除答案外没有 run，回退到全局正文基准（其它单选题干段）。
        """
        pidx = q.paragraph_index
        para = self.parser.paragraphs[pidx] if 0 <= pidx < len(self.parser.paragraphs) else None
        answer_run_ids = {id(r) for r in q.answer_runs}
        base_runs = []
        if para is not None:
            base_runs = [r for r in para.runs if id(r) not in answer_run_ids and (r.text or "").strip()]
        if not base_runs:
            # 回退：全局正文基准 —— 取所有单选/判断题干段落中"非答案 run"的多数派
            base_runs = []
            for qq in (self.ext.single_questions[:20] + self.ext.judge_questions[:20]):
                if qq is q: continue
                p = self.parser.paragraphs[qq.paragraph_index] if 0 <= qq.paragraph_index < len(self.parser.paragraphs) else None
                if p is None: continue
                ans_ids = {id(r) for r in qq.answer_runs}
                base_runs.extend([r for r in p.runs if id(r) not in ans_ids and (r.text or "").strip()])
        if not base_runs:
            return {"font": None, "size_pt": None, "bold": None, "italic": None,
                    "underline": None, "highlight": None, "color": None}

        def _majority(vals):
            counts: dict = {}
            for v in vals:
                counts[v] = counts.get(v, 0) + 1
            # 优先选非 None 的多数派；全部 None 则返回 None
            non_none = {k: v for k, v in counts.items() if k is not None}
            if non_none:
                return max(non_none.items(), key=lambda kv: kv[1])[0]
            return None

        return {
            "font":      _majority([r.font for r in base_runs]),
            "size_pt":   _majority([r.size_pt for r in base_runs]),
            "bold":      _majority([r.bold for r in base_runs]),
            "italic":    _majority([r.italic for r in base_runs]),
            "underline": _majority([r.underline for r in base_runs]),
            "highlight": _majority([r.highlight for r in base_runs]),
            "color":     _majority([r.color for r in base_runs]),
        }

    def _answer_run_inheritance_mismatch(self, q, baseline):
        """比较答案 run 的（解析继承后的）格式与题干基准。
        RunInfo 中的属性由 StyleResolver 解析后设置——None 表示未从样式栈解析到；
        为宽容处理未显式设置的属性，只在双方均非 None 且不同时判为不一致。
        """
        diffs: list[str] = []
        for r in q.answer_runs:
            # 字体：宋体近似字体视为等价
            b_font, r_font = baseline.get("font"), r.font
            if b_font and r_font and not (_is_songti(b_font) and _is_songti(r_font)) and b_font != r_font:
                diffs.append(f"字体({r_font}!={b_font})")
            # 字号：允许 0.5 磅误差
            b_sz, r_sz = baseline.get("size_pt"), r.size_pt
            if b_sz is not None and r_sz is not None and abs(float(r_sz) - float(b_sz)) > 0.5:
                diffs.append(f"字号({r_sz}!={b_sz})")
            # 加粗/倾斜/下划线/高亮：布尔严格比对（None 视为未知，不判）
            for attr, label in (("bold","加粗"),("italic","倾斜"),("underline","下划线"),("highlight","高亮")):
                bv = baseline.get(attr)
                rv = getattr(r, attr, None)
                if bv is not None and rv is not None and bool(bv) != bool(rv):
                    diffs.append(f"{label}({rv}!={bv})")
            # 颜色：容忍 auto/无色
            b_c, r_c = baseline.get("color"), r.color
            if b_c and r_c and b_c.lower() not in ("auto", "") and r_c.lower() not in ("auto", "") and b_c.lower() != r_c.lower():
                diffs.append(f"颜色({r_c}!={b_c})")
        return sorted(set(diffs))

    def _answer_style_issues(self, q):
        issues = []
        runs = q.answer_runs
        if not runs:
            return ["未能定位答案字符run"]
        pidx = q.answer_paragraph_index if q.answer_paragraph_index is not None else q.paragraph_index
        p = self.parser.paragraphs[pidx] if 0 <= pidx < len(self.parser.paragraphs) else None
        if p is None:
            issues.append("未能定位答案段落")
        else:
            if p.alignment not in (None, "left", "both"):
                issues.append("非左对齐")
            if not _is_single_line_spacing(p):
                issues.append("非单倍行距")
        for r in runs:
            if not r.font: issues.append("字体缺失")
            elif not _is_songti(r.font): issues.append("非宋体/相近字体")
            if r.size_pt is None: issues.append("字号缺失")
            elif not (13 <= r.size_pt <= 14): issues.append("字号非13-14磅")
            if r.bold: issues.append("加粗")
            if r.italic: issues.append("倾斜")
            if r.underline: issues.append("下划线")
            if r.highlight: issues.append("高亮")
            if not _color_is_plain(r.color): issues.append("特殊颜色")
        if self.parser.comment_count:
            issues.append("存在批注")
        return sorted(set(issues))

    def _answer_style_ok(self, q):
        return not self._answer_style_issues(q)

    def _format_evidence(self, qs, bad):
        reason_counts = {}
        for q in bad:
            for issue in self._answer_style_issues(q):
                reason_counts[issue] = reason_counts.get(issue, 0) + 1
        samples = [f"{q.number}:{'/'.join(self._answer_style_issues(q))}" for q in bad[:5]]
        return f"答案格式异常={len(bad)}道；原因统计={reason_counts}；样例={samples}"

    def _single_answer_format_issues(self, q):
        # 细则：A/B/C/D 答案字母沿用对应正文格式：宋体或相近、左对齐、13-14磅、单倍行距，
        # 不加粗、不倾斜、不添加下划线、高亮或特殊颜色。
        issues = []
        runs = q.answer_runs
        if not runs:
            return ["未能定位答案字符run"]
        pidx = q.answer_paragraph_index if q.answer_paragraph_index is not None else q.paragraph_index
        p = self.parser.paragraphs[pidx] if 0 <= pidx < len(self.parser.paragraphs) else None
        if p is None:
            issues.append("未能定位答案段落")
        else:
            if p.alignment not in (None, "left", "both"):
                issues.append("非左对齐")
            if not _is_single_line_spacing(p):
                issues.append("非单倍行距")
        for r in runs:
            if r.font and not _is_songti(r.font):
                issues.append("非宋体/相近字体")
            if r.size_pt is not None and not (13 <= r.size_pt <= 14):
                issues.append("字号非13-14磅")
            if r.bold: issues.append("加粗")
            if r.italic: issues.append("倾斜")
            if r.underline: issues.append("下划线")
            if r.highlight: issues.append("高亮")
            if not _color_is_plain(r.color): issues.append("特殊颜色")
        return sorted(set(issues))

    def _single_format(self):
        qs = self.ext.single_questions
        bad = [(q, self._single_answer_format_issues(q)) for q in qs if self._single_answer_format_issues(q)]
        ok = len(qs) == 510 and not bad
        reason_counts = {}
        for _, issues in bad:
            for issue in issues:
                reason_counts[issue] = reason_counts.get(issue, 0) + 1
        samples = [f"{q.number}:{'/'.join(issues)}" for q, issues in bad[:5]]
        return R("D2+3-SINGLE-FMT", "单选题答案文字格式", 3, ok,
                 f"答案格式异常={len(bad)}道；原因统计={reason_counts}；样例={samples}", "中")

    def _single_correct(self):
        # 细则：括号内字母与第92—94页答案表题号后字母对应
        # 明确给出的对照点：1→A, 2→D, 3→C, 4→B, 50→D, 100→B
        # 主校验：用从答案表提取的 single_answers（510条）逐题比对
        # 兜底校验：无论答案表是否齐全，细则明确的6个参考点必须全部一致
        mismatches = []
        answer_map = self.ext.single_answers
        if len(answer_map) >= 100:
            # 有足量答案表数据，全表比对
            for n, expected in answer_map.items():
                q = self.singles.get(n)
                if q and q.answer != expected:
                    mismatches.append(f"{n}:题内({q.answer})!=答案表({expected})")
        # 无论如何，细则明确的6个参考点必须全部正确
        ref_fails = []
        for n, expected in _SINGLE_SAMPLES.items():
            q = self.singles.get(n)
            if not q:
                ref_fails.append(f"第{n}题不存在")
            elif q.answer != expected:
                ref_fails.append(f"第{n}题:题内({q.answer})!=细则({expected})")
        all_mismatches = mismatches + ref_fails
        ok = not all_mismatches
        return R("D2+5-SINGLE-CORRECT", "单选题第1—510题答案正确性", 5, ok,
                 f"比对答案数={len(answer_map)}，不一致={all_mismatches[:10]}，细则参考点={ref_fails or '全部一致'}",
                 "中" if len(answer_map) < 510 else "高")

    def _judge_complete(self):
        # 细则：508道判断题末尾括号内均填写一个"√"或"×"
        # 不能出现空白、"对""错""正确""错误""T""F"或A-D字母
        qs = self.ext.judge_questions
        bad = []
        forbidden_words = ("对", "错", "正确", "错误", "T", "F", "A", "B", "C", "D")
        for q in qs:
            ans = q.answer
            if q.answer_in_tail_paren and len(ans) == 1 and ans in ("√", "×"):
                continue
            reason = []
            if not q.answer_in_tail_paren:
                reason.append("括号未在末尾或无括号")
            elif ans == "":
                reason.append("括号内空白")
            elif ans in forbidden_words:
                reason.append(f"括号内出现禁用字符({ans!r})")
            elif any(w in ans for w in forbidden_words):
                reason.append(f"括号内含禁用内容({ans!r})")
            else:
                reason.append(f"括号内非√或×({ans!r})")
            bad.append(f"{q.number}:{','.join(reason)}")
        ok = len(qs) == 508 and not bad
        return R("D2+5-JUDGE-COMPLETE", "判断题1—508题答案完整性", 5, ok,
                 f"合格={508 - len(bad)}/508，检测题数={len(qs)}" + (f"，不合格样例={bad[:5]}" if bad else ""))

    def _judge_position(self):
        # 细则：每道判断题答案填写在题干末尾原有全角括号内，显示为（√）或（×）
        # 答案符号不能位于题号前、题目下一行或括号外部
        qs = self.ext.judge_questions
        bad = []
        for q in qs:
            ans = q.answer
            if not q.answer_in_tail_paren:
                bad.append(f"{q.number}:答案不在题干末尾全角括号内或位于括号外")
                continue
            if not (len(ans) == 1 and ans in ("√", "×")):
                bad.append(f"{q.number}:未显示为（√）或（×）({ans!r})")
                continue
            # answer_in_tail_paren 已由提取器基于题目续行文本确认，避免把 Word XML
            # 中的题干续行段落误判为“答案位于题目下一行”。
            ptext = q.text
            qno = re.match(r"\s*\d{1,3}[\.．、\s]", ptext)
            ans_pos = ptext.find(f"（{ans}）")
            if qno and 0 <= ans_pos < qno.end():
                bad.append(f"{q.number}:答案位于题号前")
        ok = len(qs) == 508 and not bad
        return R("D2+5-JUDGE-POS", "判断题答案所在位置", 5, ok,
                 f"题末全角括号内={508 - len(bad)}/508，检测题数={len(qs)}" + (f"，不合格样例={bad[:5]}" if bad else ""))

    def _judge_answer_format_issues(self, q):
        # 细则：√和× 沿用对应正文格式：宋体或相近字体、左对齐、13-14磅、单倍行距
        # 不添加红色、加粗、高亮、下划线或批注
        # 实现策略：
        # ① 优先与"同题题干正文"格式做对比（沿用继承格式）——通过 _paragraph_body_baseline 取基准
        # ② 与基准不一致 → 记录"不沿用正文格式(具体差异)"
        # ③ 格式缺失（run 未定位、答案段落缺失）→ 视为不通过（rubric 建议：缺失格式证据时不通过）
        # ④ 装饰类禁令（红色/加粗/高亮/下划线/批注）单独硬检查——与基准一致也不允许
        # ⑤ 硬性数值范围（13-14 磅、宋体/相近）保留作兜底，只在缺基准时使用
        issues: list[str] = []
        runs = q.answer_runs
        if not runs:
            return ["未能定位答案字符run(格式无法确认)"]
        # 段落级：对齐、行距
        pidx = q.answer_paragraph_index if q.answer_paragraph_index is not None else q.paragraph_index
        p = self.parser.paragraphs[pidx] if 0 <= pidx < len(self.parser.paragraphs) else None
        if p is None:
            issues.append("未能定位答案段落(格式无法确认)")
        else:
            if p.alignment not in (None, "left", "both"):
                issues.append("非左对齐")
            if not _is_single_line_spacing(p):
                issues.append("非单倍行距")

        # 与同题题干正文基准的继承格式对比
        baseline = self._paragraph_body_baseline(q)
        base_has_signal = any(v is not None for v in baseline.values())
        if base_has_signal:
            diffs = self._answer_run_inheritance_mismatch(q, baseline)
            for d in diffs:
                issues.append(f"不沿用正文格式:{d}")
        else:
            # 无基准可对比：进入风险
            issues.append("题干正文基准缺失(无法确认沿用格式)")

        # 装饰类禁令 + 硬性范围（作为独立底线，即使继承相同也不允许）
        for r in runs:
            # 字体：显式设置必须宋体/相近；未设置且基准也未设置 → 视为不确定风险
            if r.font:
                if not _is_songti(r.font):
                    issues.append("非宋体/相近字体")
            else:
                if not baseline.get("font"):
                    issues.append("答案字体未设置(继承链未确定)")
            # 字号：显式设置必须 13-14 磅；未设置且基准也未设置 → 视为不确定风险
            if r.size_pt is not None:
                if not (13 <= r.size_pt <= 14):
                    issues.append("字号非13-14磅")
            else:
                base_size = baseline.get("size_pt")
                if base_size is None:
                    issues.append("答案字号未设置(继承链未确定)")
                else:
                    try:
                        base_size_f = float(base_size)
                    except (TypeError, ValueError):
                        base_size_f = None
                    if base_size_f is None or not (13 <= base_size_f <= 14):
                        # 继承下来的字号缺失或已经超出细则范围
                        issues.append("继承字号非13-14磅")
            # 装饰类禁令：不论继承与否都不允许
            if r.bold: issues.append("加粗")
            if r.underline: issues.append("下划线")
            if r.highlight: issues.append("高亮")
            if _color_is_red(r.color): issues.append("红色")
        if self.parser.comment_count:
            issues.append("存在批注")
        return sorted(set(issues))

    def _judge_format(self):
        qs = self.ext.judge_questions
        bad = [(q, self._judge_answer_format_issues(q)) for q in qs if self._judge_answer_format_issues(q)]
        ok = len(qs) == 508 and not bad
        reason_counts = {}
        for _, issues in bad:
            for issue in issues:
                reason_counts[issue] = reason_counts.get(issue, 0) + 1
        samples = [f"{q.number}:{'/'.join(issues)}" for q, issues in bad[:5]]
        return R("D2+3-JUDGE-FMT", "判断题答案文字格式", 3, ok,
                 f"答案格式异常={len(bad)}道；原因统计={reason_counts}；样例={samples}", "中")

    def _judge_correct(self):
        # 细则：括号内√或×与第95页答案表对应
        # 明确参考点：1-3√, 4-5×, 51-53√, 54-55×, 96-98√, 99-100×,
        #             401-403√, 404-405×, 496-498√, 499-500×,
        #             501-503√, 504-505×, 506-508√
        mismatches = []
        answer_map = self.ext.judge_answers
        if len(answer_map) >= 100:
            for n, expected in answer_map.items():
                q = self.judges.get(n)
                if q and q.answer != expected:
                    mismatches.append(f"{n}:题内({q.answer})!=答案表({expected})")
        ref_fails = []
        for n, expected in _JUDGE_SAMPLES.items():
            q = self.judges.get(n)
            if not q:
                ref_fails.append(f"第{n}题不存在")
            elif q.answer != expected:
                ref_fails.append(f"第{n}题:题内({q.answer})!=细则({expected})")
        all_mismatches = mismatches + ref_fails
        ok = not all_mismatches
        return R("D2+5-JUDGE-CORRECT", "判断题第1—508题答案正确性", 5, ok,
                 f"比对答案数={len(answer_map)}，不一致={all_mismatches[:10]}，细则参考点={ref_fails or '全部一致'}",
                 "中" if len(answer_map) < 508 else "高")

    def _deductions(self):
        # rubric 修订：删除全部四项 -3 扣分项
        # （大面积乱码/竖排/文字重叠/越界；红色答案标记/批注/箭头/修订/说明文字；
        #   空白页超40%；字体重叠、显示不清）
        return []


# ─── 输出和入口 ───────────────────────────────────────────────────────────────
SCRIPT_ID = "038"
# 维度二正分项 max_delta 之和；用于 max_score 兜底
MAX_SCORE_TOTAL = 5 + 5 + 3 + 3 + 3 + 3 + 5 + 5 + 3 + 5 + 5 + 5 + 3 + 5  # = 58


def result_to_dict(r: RuleResult):
    return {"rule_id": r.rule_id, "title": r.title, "score": r.score, "passed": r.passed,
            "evidence": r.evidence, "confidence": r.confidence, "fatal": r.fatal}

def build_report(path: str):
    if not path.lower().endswith(".docx"):
        rr = R("D1-0","文件扩展名",0,False,"交付文件不是 .docx 格式",fatal=True)
        return {"gate_passed": False, "gate_failures": [result_to_dict(rr)], "hits": [],
                "total_score": 0, "facts_summary": {}, "manual_review_notes": []}
    try:
        pkg = DocxPackage(path).open()
    except Exception as e:
        rr = R("D1-0","文件可打开",0,False,f"无法作为 docx/zip 打开：{e}",fatal=True)
        return {"gate_passed": False, "gate_failures": [result_to_dict(rr)], "hits": [],
                "total_score": 0, "facts_summary": {}, "manual_review_notes": []}
    try:
        parser = DocxParser(pkg).parse()
        extractor = QuestionExtractor(parser); extractor.extract()
        gate_results = GateChecker(pkg, parser, extractor).check()
        gate_failures = [r for r in gate_results if not r.passed]
        facts = {
            "paragraphs": len(parser.paragraphs),
            "body_chars": len(parser.body_text.strip()),
            "single_questions": len(extractor.single_questions),
            "judge_questions": len(extractor.judge_questions),
            "single_answer_table_count": len(extractor.single_answers),
            "judge_answer_table_count": len(extractor.judge_answers),
            "headers": pkg.header_parts(),
            "footers": pkg.footer_parts(),
            "media_count": parser.image_count,
            "drawing_count": parser.drawing_count,
            "page_signals": parser.page_signals,
            "watermark_candidates": parser.watermark_candidates[:5],
        }
        notes = []
        if not pkg.header_parts(): notes.append("未发现 word/header*.xml，页眉/水印可能不存在或不是典型页眉形式。")
        if parser.page_signals.get("app_pages") is None and parser.page_signals.get("last_rendered_pages") is None:
            notes.append("未发现可靠页数信号，页数相关项为静态近似检测。")
        if not extractor.answer_table_found:
            notes.append("未能完整解析答案表，答案正确性仅能使用评分细则给出的样例点兜底。")
        if gate_failures:
            return {"gate_passed": False,
                    "gate_results": [result_to_dict(r) for r in gate_results],
                    "gate_failures": [result_to_dict(r) for r in gate_failures],
                    "hits": [], "total_score": 0, "facts_summary": facts,
                    "manual_review_notes": notes}
        d2_results = ScoreChecker(pkg, parser, extractor).score()
        # 正分项：passed=True 才加分；扣分项：score<0 且 passed=True 才扣分
        hits = [r for r in d2_results if r.passed and r.score > 0]
        deductions = [r for r in d2_results if r.score < 0 and r.passed]
        total = sum(r.score for r in hits) + sum(r.score for r in deductions)
        return {"gate_passed": True,
                "gate_results": [result_to_dict(r) for r in gate_results],
                "gate_failures": [],
                "hits": [result_to_dict(r) for r in hits],
                "all_dimension2_results": [result_to_dict(r) for r in d2_results],
                "total_score": total,
                "facts_summary": facts,
                "manual_review_notes": notes}
    finally:
        pkg.close()


def _locate_docx(dir_path: str) -> str:
    """在给定目录中定位待评估的 .docx（忽略 ~$ 开头的 Word 临时锁文件）；找不到抛 FileNotFoundError。"""
    if not dir_path or not os.path.isdir(dir_path):
        raise FileNotFoundError(f"目录不存在：{dir_path}")
    cands = sorted(
        f for f in os.listdir(dir_path)
        if f.lower().endswith(".docx") and not f.startswith("~$")
    )
    if not cands:
        raise FileNotFoundError(f"目录下未发现 .docx：{dir_path}")
    return os.path.join(dir_path, cands[0])


def evaluate(dir_path: str) -> dict:
    """按接口统一约定的入口：接收"脚本所在目录"的路径，脚本自身负责在该目录里
    定位并打开被评估文档；返回结构化 dict；不 print、不改 sys.stdout、不 sys.exit。
    脚本自身崩溃（含目录/文件不存在等）由本函数捕获并返回 status='error'。"""
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
        docx_path = _locate_docx(dir_path)
        result["file_name"] = os.path.basename(docx_path)
        report = build_report(docx_path)

        # 维度一：门槛检查
        if not report.get("gate_passed", False):
            failures = report.get("gate_failures", []) or []
            reasons = [f"[{r.get('rule_id','')}] {r.get('title','')}：{r.get('evidence','')}" for r in failures]
            result["dim1_pass"] = False
            result["dim1_reason"] = "；".join(reasons) if reasons else "维度一未通过"
            result["dim2_items"] = []
            result["total_score"] = 0
            return result

        # 维度二：逐项映射（正分项与扣分项都列出，便于横向对齐）
        dim2_items = []
        max_score = 0
        for r in report.get("all_dimension2_results", []):
            score = r.get("score", 0)
            passed = bool(r.get("passed", False))
            title = r.get("title", "")
            if score > 0:
                # 正分项：命中记满分，未命中记 0
                delta = score if passed else 0
                dim2_items.append({
                    "rule": title,
                    "max_delta": score,
                    "delta": delta,
                    "hit": passed,
                    "detail": "",
                })
                max_score += score
            elif score < 0:
                # 扣分项：passed=True 表示触发扣分，delta=score(负数)
                delta = score if passed else 0
                dim2_items.append({
                    "rule": title,
                    "max_delta": score,
                    "delta": delta,
                    "hit": passed,
                    "detail": "",
                })
            # score == 0 的门槛项在 all_dimension2_results 中不出现，忽略

        result["dim2_items"] = dim2_items
        result["total_score"] = report.get("total_score", 0)
        result["max_score"] = max_score or MAX_SCORE_TOTAL
        return result
    except Exception as exc:
        result["status"] = "error"
        result["error"] = f"{type(exc).__name__}: {exc}"
        result["dim2_items"] = []
        result["total_score"] = 0
        return result


if __name__ == "__main__":
    # 本地调试入口：优先取命令行第 1 个参数作为"脚本所在目录"路径；
    # 未指定时使用脚本自身所在目录。evaluate() 本身不做任何默认路径假设——
    # 目录始终由外部显式传入；这里的自动定位仅是 __main__ 本地调试入口的便利。
    if len(sys.argv) >= 2:
        _dir = sys.argv[1]
    else:
        _dir = os.path.dirname(os.path.abspath(__file__))
    _out = evaluate(_dir)
    print(json.dumps(_out, ensure_ascii=False, indent=2))
