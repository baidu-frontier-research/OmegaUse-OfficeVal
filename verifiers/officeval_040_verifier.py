# -*- coding: utf-8 -*-
"""
自动评估脚本：对 "简介素材_格式设置_水印版.docx" 进行评估

评估逻辑：
    1) 维度1（可用与可修改性）为门槛项，任何一条不满足则总分 0，不再评估维度2；
    2) 维度2 分为加分点（须满足条目内每一子项才加分）和扣分点（满足任一子项即扣分），
       将命中的分值累加得到最终得分。

对外接口：仅暴露 evaluate(dir_path: str) -> dict —— 接收脚本所在目录的路径，
由脚本自行在该目录中定位 .docx 文件并打开，返回统一结构的字典（详见 evaluate 文档串）。

说明：
- 仅支持 .docx（Office Open XML）；不支持二进制 .doc。
- 整套评估基于 zipfile + lxml 静态解析，不依赖 Microsoft Word COM、LibreOffice 等外部 Office。
"""

from __future__ import annotations

import io
import json
import os

SCRIPT_ID = "040"
import re
import sys
import zipfile
from dataclasses import dataclass, field
from typing import Callable

# python-docx 只用于打开/读取的便捷断言；核心解析基于底层 XML，更精确。
try:
    from docx import Document  # type: ignore
except Exception:  # pragma: no cover
    Document = None

from lxml import etree
from PIL import Image

# ---------------------------------------------------------------------------
# 命名空间与常量
# ---------------------------------------------------------------------------

NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "w14": "http://schemas.microsoft.com/office/word/2010/wordml",
    "v": "urn:schemas-microsoft-com:vml",
    "o": "urn:schemas-microsoft-com:office:office",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
}


def q(ns: str, tag: str) -> str:
    return f"{{{NS[ns]}}}{tag}"


# 尺寸换算
HALF_POINT_PER_PT = 2       # size in Word XML is half-points (sz="44" => 22pt)
TWIPS_PER_PT = 20           # space/line values in 二十分之一磅
EMU_PER_PT = 12700          # DrawingML EMU
POINTS_PER_CM = 28.3464567

# 中文字号对照 (磅)
FONT_SIZE_NAME_TO_PT = {
    "初号": 42, "小初": 36, "一号": 26, "小一": 24,
    "二号": 22, "小二": 18, "三号": 16, "小三": 15,
    "四号": 14, "小四": 12, "五号": 10.5, "小五": 9,
}

# 橄榄色候选：Word 主题 accent3 (Office 2007+) 为 #9BBB59；
# 也接受经典"橄榄色"#808000 及其相近色。
OLIVE_HEX_CANDIDATES = ["9BBB59", "808000", "6B8E23", "A2B463"]

# 蓝色候选（标题颜色）
BLUE_HEX_CANDIDATES = ["0000FF", "0070C0", "1F3864", "2F5496", "4472C4"]

# 绿色候选（"智能软件工程系"文字色）
GREEN_HEX_CANDIDATES = ["008000", "00B050", "00B04F", "008040", "339933"]


# ---------------------------------------------------------------------------
# XML 载入
# ---------------------------------------------------------------------------


@dataclass
class DocxBundle:
    path: str
    zf: zipfile.ZipFile
    document: etree._Element
    headers: dict[str, etree._Element] = field(default_factory=dict)
    theme: etree._Element | None = None
    rels_map: dict[str, dict[str, str]] = field(default_factory=dict)
    accent_map: dict[str, str] = field(default_factory=dict)  # accentN -> hex

    def close(self) -> None:
        self.zf.close()


def load_bundle(path: str) -> DocxBundle:
    zf = zipfile.ZipFile(path)
    document = etree.fromstring(zf.read("word/document.xml"))

    bundle = DocxBundle(path=path, zf=zf, document=document)

    # 头部（水印通常在页眉）
    for name in zf.namelist():
        if re.fullmatch(r"word/header\d+\.xml", name):
            bundle.headers[name] = etree.fromstring(zf.read(name))

    # 主题色（用于将 schemeClr accent3 解析为 sRGB）
    if "word/theme/theme1.xml" in zf.namelist():
        theme = etree.fromstring(zf.read("word/theme/theme1.xml"))
        bundle.theme = theme
        for accent_name in ("accent1", "accent2", "accent3", "accent4", "accent5", "accent6"):
            el = theme.find(f".//{{{NS['a']}}}{accent_name}")
            if el is not None:
                srgb = el.find(f"{{{NS['a']}}}srgbClr")
                if srgb is not None and "val" in srgb.attrib:
                    bundle.accent_map[accent_name] = srgb.attrib["val"].upper()

    return bundle


# ---------------------------------------------------------------------------
# XML 便捷访问工具
# ---------------------------------------------------------------------------


def get_body_paragraphs(bundle: DocxBundle) -> list[etree._Element]:
    return bundle.document.findall(f".//{q('w','body')}/{q('w','p')}")


def paragraph_text(p: etree._Element) -> str:
    return "".join(t.text or "" for t in p.findall(f".//{q('w','t')}"))


def run_font_east_asia(r: etree._Element) -> str | None:
    rfonts = r.find(f"{q('w','rPr')}/{q('w','rFonts')}")
    if rfonts is None:
        return None
    return rfonts.get(q("w", "eastAsia")) or rfonts.get(q("w", "ascii"))


def run_size_pt(r: etree._Element) -> float | None:
    sz = r.find(f"{q('w','rPr')}/{q('w','sz')}")
    if sz is None or "val" not in sz.attrib.get(q("w", "val"), "") and q("w", "val") not in sz.attrib:
        # 兼容不同属性写法
        val = sz.attrib.get(q("w", "val")) if sz is not None else None
        if val is None:
            return None
        return float(val) / HALF_POINT_PER_PT
    return float(sz.attrib[q("w", "val")]) / HALF_POINT_PER_PT


def run_bool(r: etree._Element, tag: str) -> bool:
    el = r.find(f"{q('w','rPr')}/{q('w', tag)}")
    if el is None:
        return False
    val = el.get(q("w", "val"))
    if val is None:
        return True
    return val not in ("0", "false", "off")


def run_color_hex(r: etree._Element) -> str | None:
    c = r.find(f"{q('w','rPr')}/{q('w','color')}")
    if c is None:
        return None
    return (c.get(q("w", "val")) or "").upper() or None


def run_em(r: etree._Element) -> str | None:
    em = r.find(f"{q('w','rPr')}/{q('w','em')}")
    return None if em is None else em.get(q("w", "val"))


def para_alignment(p: etree._Element) -> str | None:
    jc = p.find(f"{q('w','pPr')}/{q('w','jc')}")
    return None if jc is None else jc.get(q("w", "val"))


def para_spacing(p: etree._Element) -> dict[str, str]:
    sp = p.find(f"{q('w','pPr')}/{q('w','spacing')}")
    return {} if sp is None else {re.sub(r".*}", "", k): v for k, v in sp.attrib.items()}


def para_indent(p: etree._Element) -> dict[str, str]:
    ind = p.find(f"{q('w','pPr')}/{q('w','ind')}")
    return {} if ind is None else {re.sub(r".*}", "", k): v for k, v in ind.attrib.items()}


def para_framepr(p: etree._Element) -> dict[str, str] | None:
    fr = p.find(f"{q('w','pPr')}/{q('w','framePr')}")
    if fr is None:
        return None
    return {re.sub(r".*}", "", k): v for k, v in fr.attrib.items()}


def resolve_glow_color(rpr: etree._Element, accent_map: dict[str, str]) -> str | None:
    glow = rpr.find(f"{q('w14','glow')}")
    if glow is None:
        return None
    srgb = glow.find(f"{q('w14','srgbClr')}")
    if srgb is not None:
        return (srgb.get("val") or "").upper()
    scheme = glow.find(f"{q('w14','schemeClr')}")
    if scheme is not None:
        val = scheme.get("val")
        return accent_map.get(val)
    return None


def hex_distance(a: str, b: str) -> float:
    """两个十六进制颜色的欧氏距离。"""
    ai = [int(a[i:i + 2], 16) for i in (0, 2, 4)]
    bi = [int(b[i:i + 2], 16) for i in (0, 2, 4)]
    return sum((x - y) ** 2 for x, y in zip(ai, bi)) ** 0.5


def is_near(color: str | None, candidates: list[str], threshold: float = 80.0) -> bool:
    if not color:
        return False
    color = color.upper()
    return any(hex_distance(color, c) <= threshold for c in candidates)


# ---------------------------------------------------------------------------
# Office 属性生效解析（run rPr > 段落 pPr/rPr > 段落样式链）
# 这样即使属性通过"样式"设置而非直接在 run 上，也能被检测到（贴合 Word/WPS 渲染）。
# ---------------------------------------------------------------------------


def _styles_map(bundle: "DocxBundle") -> dict[str, etree._Element]:
    if not hasattr(bundle, "_styles_cache"):
        cache: dict[str, etree._Element] = {}
        if "word/styles.xml" in bundle.zf.namelist():
            root = etree.fromstring(bundle.zf.read("word/styles.xml"))
            for st in root.findall(q("w", "style")):
                sid = st.get(q("w", "styleId"))
                if sid:
                    cache[sid] = st
        bundle._styles_cache = cache  # type: ignore[attr-defined]
    return bundle._styles_cache  # type: ignore[attr-defined]


def _pstyle_id(p: etree._Element) -> str | None:
    ps = p.find(f"{q('w','pPr')}/{q('w','pStyle')}")
    return ps.get(q("w", "val")) if ps is not None else None


def _style_chain(styles: dict[str, etree._Element], style_id: str | None) -> list[etree._Element]:
    chain: list[etree._Element] = []
    cur = style_id
    seen: set[str] = set()
    while cur and cur not in seen and cur in styles:
        seen.add(cur)
        chain.append(styles[cur])
        based = styles[cur].find(q("w", "basedOn"))
        cur = based.get(q("w", "val")) if based is not None else None
    return chain


def _run_prop_sources(r: etree._Element, p: etree._Element, bundle) -> list[etree._Element]:
    """按 Word 生效优先级（高→低）返回可查询的 rPr 元素列表。"""
    sources: list[etree._Element] = []
    r_rpr = r.find(q("w", "rPr"))
    if r_rpr is not None:
        sources.append(r_rpr)
    p_rpr = p.find(f"{q('w','pPr')}/{q('w','rPr')}")
    if p_rpr is not None:
        sources.append(p_rpr)
    for st in _style_chain(_styles_map(bundle), _pstyle_id(p)):
        st_rpr = st.find(q("w", "rPr"))
        if st_rpr is not None:
            sources.append(st_rpr)
    return sources


def _p_prop_sources(p: etree._Element, bundle) -> list[etree._Element]:
    sources: list[etree._Element] = []
    p_pr = p.find(q("w", "pPr"))
    if p_pr is not None:
        sources.append(p_pr)
    for st in _style_chain(_styles_map(bundle), _pstyle_id(p)):
        st_ppr = st.find(q("w", "pPr"))
        if st_ppr is not None:
            sources.append(st_ppr)
    return sources


def effective_font_eastasia(r, p, bundle) -> str | None:
    for src in _run_prop_sources(r, p, bundle):
        rfonts = src.find(q("w", "rFonts"))
        if rfonts is not None:
            v = rfonts.get(q("w", "eastAsia")) or rfonts.get(q("w", "ascii"))
            if v:
                return v
    return None


def effective_size_pt(r, p, bundle) -> float | None:
    for src in _run_prop_sources(r, p, bundle):
        sz = src.find(q("w", "sz"))
        if sz is not None:
            v = sz.get(q("w", "val"))
            if v:
                return float(v) / 2.0  # w:sz 单位为半磅
    return None


def effective_bold(r, p, bundle) -> bool:
    # Word 的加粗为 toggle：<w:b/> 表示开；<w:b w:val="0"/> 表示关。取最高优先级来源。
    for src in _run_prop_sources(r, p, bundle):
        b = src.find(q("w", "b"))
        if b is not None:
            v = b.get(q("w", "val"))
            return (v is None) or (v not in ("0", "false", "off"))
    return False


def effective_italic(r, p, bundle) -> bool:
    """倾斜同样是 toggle：<w:i/> 开；<w:i w:val="0"/> 关。"""
    for src in _run_prop_sources(r, p, bundle):
        it = src.find(q("w", "i"))
        if it is not None:
            v = it.get(q("w", "val"))
            return (v is None) or (v not in ("0", "false", "off"))
    return False


def effective_em(r, p, bundle) -> str | None:
    """w:em 表示着重号，取值：none / dot / comma / circle / underDot。
    Office 中"字体→着重号"选择任一非 none 即为"添加着重号"。"""
    for src in _run_prop_sources(r, p, bundle):
        em = src.find(q("w", "em"))
        if em is not None:
            return em.get(q("w", "val"))
    return None


def effective_color(r, p, bundle) -> tuple[str | None, str | None]:
    for src in _run_prop_sources(r, p, bundle):
        c = src.find(q("w", "color"))
        if c is not None:
            hex_v = c.get(q("w", "val"))
            theme_v = c.get(q("w", "themeColor"))
            if hex_v or theme_v:
                return (hex_v.upper() if hex_v else None), theme_v
    return None, None


def effective_alignment(p, bundle) -> str | None:
    for src in _p_prop_sources(p, bundle):
        jc = src.find(q("w", "jc"))
        if jc is not None:
            v = jc.get(q("w", "val"))
            if v:
                return v
    return None


def effective_space_after(p, bundle) -> tuple[float | None, bool]:
    """返回 (段后磅数, afterAutospacing 是否开启)。autospacing 开启时 Office 忽略数值。"""
    after_pt: float | None = None
    auto_on = False
    for src in _p_prop_sources(p, bundle):
        sp = src.find(q("w", "spacing"))
        if sp is None:
            continue
        if after_pt is None:
            a = sp.get(q("w", "after"))
            if a is not None:
                after_pt = float(a) / 20.0  # twips → pt
        auto = sp.get(q("w", "afterAutospacing"))
        if auto in ("1", "true", "on"):
            auto_on = True
        if after_pt is not None:
            break
    return after_pt, auto_on


def effective_line_spacing(p, bundle) -> tuple[float | None, str | None]:
    """返回 (行距倍数, lineRule)。倍数依据 lineRule：
       - auto  : line / 240
       - atLeast/exact : 视为磅数比例（不返回倍数，multiplier=None）
    """
    for src in _p_prop_sources(p, bundle):
        sp = src.find(q("w", "spacing"))
        if sp is None:
            continue
        line = sp.get(q("w", "line"))
        rule = sp.get(q("w", "lineRule")) or "auto"
        if line is not None:
            if rule == "auto":
                return float(line) / 240.0, rule
            return None, rule
    return None, None


def effective_first_line_indent(p, bundle) -> tuple[str | None, int | None, int | None]:
    """返回 (firstLineChars, firstLine_twips, hanging_twips)。
    办公软件"首行缩进 N 字符" 对应 firstLineChars=N*100；
    也可用绝对量 firstLine (twips)。若 hanging 存在，则为悬挂缩进而非首行缩进。
    """
    for src in _p_prop_sources(p, bundle):
        ind = src.find(q("w", "ind"))
        if ind is None:
            continue
        fl_chars = ind.get(q("w", "firstLineChars"))
        fl = ind.get(q("w", "firstLine"))
        hanging = ind.get(q("w", "hanging"))
        return (
            fl_chars,
            int(fl) if fl is not None else None,
            int(hanging) if hanging is not None else None,
        )
    return None, None, None


def _theme_hex(bundle, theme_name: str | None) -> str | None:
    if not theme_name or not bundle.accent_map:
        return None
    return bundle.accent_map.get(theme_name)


def _is_office_blue(hex_val: str | None, theme_val: str | None, bundle) -> bool:
    """按 Word/WPS 实际显示判断是否为蓝色：B 通道主导且明显高于 R、G。
    覆盖标准色"蓝色"#0000FF、"深蓝"#002060、"浅蓝"#00B0F0 等常见蓝色系。"""
    if hex_val is None and theme_val:
        hex_val = _theme_hex(bundle, theme_val)
    if not hex_val or hex_val.upper() == "AUTO":
        return False
    try:
        r = int(hex_val[0:2], 16); g = int(hex_val[2:4], 16); b = int(hex_val[4:6], 16)
    except Exception:
        return False
    return b >= 100 and b >= r + 30 and b >= g + 30


# ---------------------------------------------------------------------------
# 评估结果收集
# ---------------------------------------------------------------------------


@dataclass
class RuleResult:
    rule_id: str
    score: int
    passed: bool
    detail: str


@dataclass
class Report:
    dim1_passed: bool = True
    dim1_reasons: list[str] = field(default_factory=list)
    rules: list[RuleResult] = field(default_factory=list)

    def add_dim1_fail(self, reason: str) -> None:
        self.dim1_passed = False
        self.dim1_reasons.append(reason)

    def add(self, rid: str, score: int, passed: bool, detail: str) -> None:
        self.rules.append(RuleResult(rid, score, passed, detail))

    def final_score(self) -> int:
        if not self.dim1_passed:
            return 0
        return sum(r.score for r in self.rules if r.passed)

    def render(self) -> str:
        # 维度 1 未通过：仅输出维度一失败原因 + 总分 0
        if not self.dim1_passed:
            lines = ["维度一：未通过"]
            for r in self.dim1_reasons:
                lines.append(f"  - {r}")
            lines.append(f"总分：0")
            return "\n".join(lines)

        # 维度 1 通过：按用户指定格式输出
        lines = ["维度一：通过", "维度二：评分结果"]
        for r in self.rules:
            if not r.passed:
                continue  # 只显示命中项
            sign = "+" if r.score >= 0 else ""
            lines.append(f"{sign}{r.score}：{r.rule_id}")
        lines.append(f"总分：{self.final_score()}")
        return "\n".join(lines)


# ---------------------------------------------------------------------------
# 维度 1 检查
# ---------------------------------------------------------------------------


def check_dimension1(bundle: DocxBundle, report: Report) -> None:
    # 1.1 .docx 格式，正常打开：能到这里说明能打开
    if not bundle.path.lower().endswith(".docx"):
        report.add_dim1_fail("文件不是 .docx 扩展名")
    try:
        if Document is not None:
            Document(bundle.path)
    except Exception as e:
        report.add_dim1_fail(f"python-docx 无法打开文档：{e}")

    # 原 1.2 (标题/正文可编辑文本)、1.3 (连续空白页/乱码/图片遮挡)、
    # 1.4 (水印为图片水印且正文未被转成图片) 三项已按用户要求删除。


# ---------------------------------------------------------------------------
# 维度 2 检查
# ---------------------------------------------------------------------------


def find_title_run(bundle: DocxBundle) -> tuple[etree._Element | None, etree._Element | None]:
    """返回 (标题段落, 标题 run)。"""
    for p in get_body_paragraphs(bundle):
        for r in p.findall(f"{q('w','r')}"):
            for t in r.findall(f"{q('w','t')}"):
                if t.text and "智能软件工程系简介" in t.text:
                    return p, r
    return None, None


def rule_title_basic_format(bundle: DocxBundle) -> tuple[bool, str]:
    """+1: 文档标题段落："智能软件工程系简介"位于文档首行，采用隶书、二号、加粗、蓝色，
    居中对齐，段后间距20磅。

    每一条细则的检测思路（贴合 Word/WPS 实际显示）：
      · 位于文档首行     -> 正文首段的可见文本 == 标题
      · 采用隶书         -> run/pPr/样式链上 rFonts eastAsia 生效值为"隶书"
      · 二号             -> w:sz 生效值 = 44 半磅 = 22pt
      · 加粗             -> w:b toggle 生效为 True
      · 蓝色             -> w:color 生效值判定为蓝色系（含 themeColor 解析）
      · 居中对齐         -> w:jc 生效值 = center
      · 段后间距 20 磅   -> w:spacing/@after = 400 twips 且未启用 afterAutospacing
    """
    body = get_body_paragraphs(bundle)
    if not body:
        return False, "文档无正文段落"

    title_text = "智能软件工程系简介"

    # 1) "位于文档首行"：第一段（跳过纯空段）文本必须严格等于标题
    first_nonempty_idx = next(
        (i for i, p in enumerate(body) if paragraph_text(p).strip()), None
    )
    checks: list[tuple[str, bool]] = []
    if first_nonempty_idx is None:
        return False, "文档无可见文本"
    first_para = body[first_nonempty_idx]
    first_text = paragraph_text(first_para).strip()
    checks.append((
        f"'{title_text}' 位于文档首行  实际首行='{first_text}'",
        first_text == title_text,
    ))

    # 定位承载标题文本的 run（用于字符级属性判定）
    title_para = first_para
    title_runs = [
        r for r in title_para.findall(q("w", "r"))
        if title_text in "".join((t.text or "") for t in r.findall(q("w", "t")))
    ]
    if not title_runs:
        # 兜底：任何包含"智能软件工程系简介"文字片段的 run 集合
        title_runs = [
            r for r in title_para.findall(q("w", "r"))
            if "".join((t.text or "") for t in r.findall(q("w", "t"))).strip()
        ]
    if not title_runs:
        return False, "标题段落无可用 run"
    r0 = title_runs[0]

    # 2) 字体：采用隶书
    font = effective_font_eastasia(r0, title_para, bundle)
    checks.append((f"采用隶书  实际={font}", font == "隶书"))

    # 3) 字号：二号 (= 22pt)
    size_pt = effective_size_pt(r0, title_para, bundle)
    checks.append((
        f"二号(22磅)  实际={size_pt}pt",
        size_pt is not None and abs(size_pt - 22.0) < 1e-6,
    ))

    # 4) 加粗
    bold = effective_bold(r0, title_para, bundle)
    checks.append(("加粗", bold))

    # 5) 蓝色
    hex_val, theme_val = effective_color(r0, title_para, bundle)
    is_blue = _is_office_blue(hex_val, theme_val, bundle)
    checks.append((
        f"蓝色  实际=#{hex_val} themeColor={theme_val}",
        is_blue,
    ))

    # 6) 居中对齐
    align = effective_alignment(title_para, bundle)
    checks.append((f"居中对齐  实际={align}", align == "center"))

    # 7) 段后间距 20 磅（Word 中 w:spacing/@after 单位为 二十分之一磅；20 磅 = 400）
    after_pt, auto_on = effective_space_after(title_para, bundle)
    space_ok = (
        after_pt is not None
        and abs(after_pt - 20.0) < 1e-6
        and not auto_on
    )
    checks.append((
        f"段后间距 20 磅  实际={after_pt}pt afterAutospacing={auto_on}",
        space_ok,
    ))

    passed = all(ok for _, ok in checks)
    return passed, "\n".join(f"{'✓' if ok else '✗'} {msg}" for msg, ok in checks)


def rule_title_glow_shadow(bundle: DocxBundle) -> tuple[bool, str]:
    """+3: 文档标题段落："智能软件工程系简介"设置发光效果，发光颜色为橄榄色或接近橄榄色，
    发光大小约7磅，使用强调文字颜色3或接近效果。"智能软件工程系简介"设置向上偏移
    阴影效果，标题文字仍清晰可读。

    每一条细则的检测思路（对齐 Word/WPS "开始→字体→文字效果和版式" 的行为）：
      · 设置发光效果             -> 标题 run 存在 w14:glow
      · 发光颜色为橄榄色或接近橄榄色 -> glow 颜色解析后与橄榄色系(#9BBB59/#808000)接近
      · 发光大小约 7 磅          -> w14:glow/@w14:rad ÷ 12700 ≈ 7pt (约 ±2 磅)
      · 使用强调文字颜色 3 或接近效果 -> glow 使用 schemeClr accent3；或颜色接近主题 accent3
      · 设置向上偏移阴影效果      -> 标题 run 存在 w14:shadow，dir=16200000 (Office"向上偏移"预设)
                                     或方向落在向上区间(225°~315°)以匹配"或接近"
      · 标题文字仍清晰可读        -> 标题字体颜色存在、非白、且与发光颜色不同（不被吞没）
    """
    _, title_run = find_title_run(bundle)
    if title_run is None:
        return False, "未找到标题 run"
    rpr = title_run.find(q("w", "rPr"))
    if rpr is None:
        return False, "标题 run 缺 rPr"

    checks: list[tuple[str, bool]] = []

    # ---- 发光相关 4 点 ----
    glow = rpr.find(q("w14", "glow"))
    # 1) 设置发光效果
    checks.append(("设置发光效果  存在 w14:glow", glow is not None))

    # 2) 发光颜色为橄榄色或接近橄榄色
    glow_color_hex: str | None = None
    scheme_val: str | None = None
    if glow is not None:
        srgb = glow.find(q("w14", "srgbClr"))
        scheme = glow.find(q("w14", "schemeClr"))
        if srgb is not None:
            glow_color_hex = (srgb.get(q("w14", "val")) or srgb.get("val") or "").upper() or None
        if scheme is not None:
            scheme_val = scheme.get(q("w14", "val")) or scheme.get("val")
            if not glow_color_hex and scheme_val:
                glow_color_hex = bundle.accent_map.get(scheme_val)
    olive_ok = is_near(glow_color_hex, OLIVE_HEX_CANDIDATES, threshold=90.0)
    checks.append((
        f"发光颜色为橄榄色或接近橄榄色  实际颜色=#{glow_color_hex}",
        olive_ok,
    ))

    # 3) 发光大小约 7 磅
    rad_pt: float | None = None
    if glow is not None:
        rad_emu_str = glow.get(q("w14", "rad")) or glow.get("rad")
        if rad_emu_str is not None:
            rad_pt = int(rad_emu_str) / EMU_PER_PT
    checks.append((
        f"发光大小约 7 磅  实际={rad_pt}pt" if rad_pt is not None else "发光大小约 7 磅  未设置 rad",
        rad_pt is not None and abs(rad_pt - 7.0) <= 2.0,
    ))

    # 4) 使用强调文字颜色 3 或接近效果
    accent3_hex = bundle.accent_map.get("accent3")
    uses_accent3_scheme = scheme_val == "accent3"
    close_to_accent3 = (
        accent3_hex is not None
        and glow_color_hex is not None
        and hex_distance(glow_color_hex.upper(), accent3_hex.upper()) <= 90.0
    )
    accent3_ok = uses_accent3_scheme or close_to_accent3
    checks.append((
        f"使用强调文字颜色 3 或接近效果  schemeClr=accent3? {uses_accent3_scheme}; "
        f"主题 accent3=#{accent3_hex}",
        accent3_ok,
    ))

    # ---- 阴影相关 1 点 ----
    shadow = rpr.find(q("w14", "shadow"))
    # 5) 设置向上偏移阴影效果
    dir_val: int | None = None
    if shadow is not None:
        dir_str = shadow.get(q("w14", "dir")) or shadow.get("dir")
        if dir_str is not None:
            dir_val = int(dir_str)
    # Office "向上偏移" 预设精确值：dir=16200000 (=270°)；允许"或接近"：225°~315°
    shadow_up_ok = (
        shadow is not None
        and dir_val is not None
        and 13500000 <= dir_val <= 18900000
    )
    checks.append((
        f"设置向上偏移阴影效果  存在 w14:shadow={shadow is not None} dir={dir_val}",
        shadow_up_ok,
    ))

    # ---- 可读性 1 点 ----
    # 6) 标题文字仍清晰可读：颜色已设定、非白色、且与发光颜色区分
    text_color = run_color_hex(title_run)
    readable = (
        text_color is not None
        and text_color.upper() != "FFFFFF"
        and (glow_color_hex is None or text_color.upper() != glow_color_hex.upper())
    )
    checks.append((
        f"标题文字仍清晰可读  文字颜色=#{text_color} 发光颜色=#{glow_color_hex}",
        readable,
    ))

    passed = all(ok for _, ok in checks)
    return passed, "\n".join(f"{'✓' if ok else '✗'} {msg}" for msg, ok in checks)


def locate_special_dropcap(
    body: "list[etree._Element]",
) -> tuple[int | None, int | None]:
    """定位"专业设置如下："段的首字下沉结构。

    返回 (first_char_idx, others_idx)，两者均为 body 中的下标：
      · 结构 A（framePr 挂在原段，首字与其余文字同段）：
            first_char_idx = None
            others_idx     = 该段（含完整"专业设置如下："文字）
      · 结构 B（Office 将首字拆到独立段，其余文字放到下一段，仅"首字段"挂 framePr）：
            first_char_idx = 挂 framePr 的"首字段"下标
            others_idx     = 紧邻的"其余文字段"下标
      · 未找到"专业设置如下"结构时两者均为 None。

    先按 framePr + dropCap 属性识别 Office 的首字下沉容器；仅当"首字段(1字)+其余
    文字段"拼接后仍以"专业设置如下"起始时才视为结构 B，否则退回按文本前缀匹配的
    结构 A（此时可能只是普通段落，framePr 不存在）。
    """
    # 1) 优先按 framePr / dropCap 定位 Office 首字下沉容器
    for i, p in enumerate(body):
        fr = p.find(f"{q('w','pPr')}/{q('w','framePr')}")
        if fr is None or fr.get(q("w", "dropCap")) not in ("margin", "drop"):
            continue
        p_text = paragraph_text(p).strip()
        # 结构 B：首字段独占单字，与下一段拼接后仍以"专业设置如下"开头
        if len(p_text) == 1 and i + 1 < len(body):
            next_text = paragraph_text(body[i + 1]).strip()
            if (p_text + next_text).startswith("专业设置如下"):
                return i, i + 1
        # 结构 A：framePr 与"专业设置如下"文字挂在同一段
        if p_text.startswith("专业设置如下"):
            return None, i
    # 2) framePr 不存在或位置异常时，按文本前缀兜底找出内容段
    for i, p in enumerate(body):
        if paragraph_text(p).strip().startswith("专业设置如下"):
            return None, i
    return None, None


def rule_body_para_format(bundle: DocxBundle) -> tuple[bool, str]:
    """+1: 正文段落格式：除标题、"专业设置如下："、"计算机应用技术"内容段落外，
    其余正文采用首行缩进2字符，段后间距6磅，1.1倍行距。

    每一条细则的检测思路（对齐 Word/WPS 段落格式面板的行为）：
      · 首行缩进 2 字符  -> 每个正文段 w:ind/@w:firstLineChars == "200"
                            （Office UI "首行缩进 2 字符" 精确对应 200，即 2×100）
                            并且不存在悬挂缩进（hanging 应为 0/缺省）
      · 段后间距 6 磅    -> w:spacing/@w:after == 120 twips (=6pt) 且未启用 afterAutospacing
      · 1.1 倍行距       -> w:spacing/@w:line == 264 且 @w:lineRule == "auto"
                            （多倍行距 1.1 → 240*1.1 = 264）

    段落例外（整段跳过，不参与判定）：
      · 标题段（首个非空段）
      · "专业设置如下："内容段：使用 locate_special_dropcap 定位；若 Office 将首字
        拆为独立段（结构 B），首字段与其余文字段共同构成该"内容段"，两段一起跳过；
        若首字与其余文字同段（结构 A），跳过该段。
      · 文本包含"计算机应用技术"的段落
      · 空段
    """
    body = get_body_paragraphs(bundle)
    if len(body) < 2:
        return False, "正文段落不足"

    # 定位标题段：正文首个非空段
    title_idx = next(
        (i for i, p in enumerate(body) if paragraph_text(p).strip()), 0
    )

    # "专业设置如下："内容段：结构 A 一段、结构 B 两段，一并从段落格式判定中排除
    first_char_idx, others_idx = locate_special_dropcap(body)
    special_para_indices: set[int] = {
        idx for idx in (first_char_idx, others_idx) if idx is not None
    }

    failures: list[str] = []
    stats = dict(indent_ok=0, after_ok=0, line_ok=0, total=0)

    for i, p in enumerate(body):
        if i == title_idx:
            continue
        if i in special_para_indices:
            continue
        text = paragraph_text(p).strip()
        if not text:
            continue  # 空段不参与判定
        if "计算机应用技术" in text:
            continue
        stats["total"] += 1

        # 首行缩进 2 字符：firstLineChars == "200"，且无 hanging
        fl_chars, fl_twips, hanging = effective_first_line_indent(p, bundle)
        indent_ok = (fl_chars == "200") and (not hanging or hanging == 0)
        if indent_ok:
            stats["indent_ok"] += 1
        else:
            failures.append(
                f"段{i}('{text[:10]}...')首行缩进异常：firstLineChars={fl_chars} "
                f"firstLine={fl_twips} hanging={hanging}"
            )

        # 段后间距 6 磅：after==120twips 且 afterAutospacing 关闭
        after_pt, auto_on = effective_space_after(p, bundle)
        after_ok = (after_pt is not None) and abs(after_pt - 6.0) < 1e-6 and (not auto_on)
        if after_ok:
            stats["after_ok"] += 1
        else:
            failures.append(
                f"段{i}('{text[:10]}...')段后异常：after={after_pt}pt afterAutospacing={auto_on}"
            )

        # 1.1 倍行距：line==264 且 lineRule=="auto"
        multiplier, rule = effective_line_spacing(p, bundle)
        line_ok = (rule == "auto") and (multiplier is not None) and abs(multiplier - 1.1) < 1e-6
        if line_ok:
            stats["line_ok"] += 1
        else:
            failures.append(
                f"段{i}('{text[:10]}...')行距异常：倍数={multiplier} lineRule={rule}"
            )

    total = stats["total"]
    detail_lines = [
        f"共评估 {total} 个正文段落（排除标题段、\"专业设置如下：\"内容段、\"计算机应用技术\"段及空段）",
        f"首行缩进2字符 {stats['indent_ok']}/{total}  "
        f"段后6磅 {stats['after_ok']}/{total}  1.1倍行距 {stats['line_ok']}/{total}",
    ]
    # 细则要求段落格式整体符合，故每个非例外段落的每一项都必须通过。
    passed = (
        total > 0
        and stats["indent_ok"] == total and stats["after_ok"] == total
        and stats["line_ok"] == total
    )
    if not passed and failures:
        detail_lines.append("失败明细：")
        for msg in failures[:8]:
            detail_lines.append(f"  - {msg}")
        if len(failures) > 8:
            detail_lines.append(f"  ...（共 {len(failures)} 条）")
    return passed, "\n".join(detail_lines)


def rule_body_font_format(bundle: DocxBundle) -> tuple[bool, str]:
    """+1: 正文字体格式：除标题、"专业设置如下："内容段落特殊首字和被替换文字
    "智能软件工程系"特殊格式外，正文采用楷体、小四号。

    每一条细则的检测思路（对齐 Word/WPS 段落格式面板的行为）：
      · 楷体             -> 每个非例外 run 的 rFonts eastAsia 生效值 == "楷体"
      · 小四号           -> 每个非例外 run 的 w:sz 生效值 == 24 半磅 (= 12pt)

    Run 级例外：
      · 标题段整段跳过
      · "专业设置如下："内容段的"特殊首字"：使用 locate_special_dropcap 定位。
          - 结构 A（framePr 在同段）：跳过该段第一个可见 run 的首字 run
          - 结构 B（首字被拆为独立段）：整段（=特殊首字）跳过；其余文字段照常参与
      · run 文本 == "智能软件工程系" 的 run（保留其特殊格式，不参与 run 属性判定）
    """
    body = get_body_paragraphs(bundle)
    if len(body) < 2:
        return False, "正文段落不足"

    # 定位标题段：正文首个非空段
    title_idx = next(
        (i for i, p in enumerate(body) if paragraph_text(p).strip()), 0
    )

    # 使用与 rule_drop_cap 一致的首字下沉定位逻辑
    first_char_idx, others_idx = locate_special_dropcap(body)
    # 结构 A：首字与其余文字同段 → 该段第一个可见 run 是特殊首字，需按 run 排除
    dropcap_first_run_para_idx: int | None = others_idx if first_char_idx is None else None
    # 结构 B：首字独占一段 → 整段（长度为 1 的字符）就是特殊首字，需按段排除
    isolated_first_char_para_idx: int | None = first_char_idx

    failures: list[str] = []
    stats = dict(font_ok=0, size_ok=0, total=0)

    for i, p in enumerate(body):
        if i == title_idx:
            continue
        if i == isolated_first_char_para_idx:
            continue  # 结构 B 的"首字段" = 特殊首字本身
        text = paragraph_text(p).strip()
        if not text:
            continue  # 空段不参与判定
        stats["total"] += 1

        # 排除：run 文本 == "智能软件工程系"；结构 A 下再排除该段首字下沉的首字 run。
        runs = p.findall(f"{q('w','r')}")
        para_font_ok = True
        para_size_ok = True
        first_visible_run_seen = False
        for r in runs:
            txt = "".join((t.text or "") for t in r.findall(f"{q('w','t')}"))
            if not txt:
                continue
            # 例外 1：被替换文字"智能软件工程系"
            if txt == "智能软件工程系":
                continue
            # 例外 2：结构 A 首字下沉段的"特殊首字"（该段第一个可见 run）
            if i == dropcap_first_run_para_idx and not first_visible_run_seen:
                first_visible_run_seen = True
                continue
            first_visible_run_seen = True
            font = effective_font_eastasia(r, p, bundle)
            size = effective_size_pt(r, p, bundle)
            if font != "楷体":
                para_font_ok = False
                failures.append(f"段{i}('{text[:10]}...')run 字体非楷体：'{txt[:10]}' font={font}")
            if size is None or abs(size - 12.0) > 1e-6:
                para_size_ok = False
                failures.append(f"段{i}('{text[:10]}...')run 字号非小四：'{txt[:10]}' size={size}")
        if para_font_ok:
            stats["font_ok"] += 1
        if para_size_ok:
            stats["size_ok"] += 1

    total = stats["total"]
    detail_lines = [
        f"共评估 {total} 个正文段落（排除标题段、结构 B 首字段、空段）",
        f"楷体 {stats['font_ok']}/{total}  小四 {stats['size_ok']}/{total}",
    ]
    # 细则要求"正文字体格式"整体符合，故每个非例外段落的每一项都必须通过。
    passed = (
        total > 0
        and stats["font_ok"] == total and stats["size_ok"] == total
    )
    if not passed and failures:
        detail_lines.append("失败明细：")
        for msg in failures[:8]:
            detail_lines.append(f"  - {msg}")
        if len(failures) > 8:
            detail_lines.append(f"  ...（共 {len(failures)} 条）")
    return passed, "\n".join(detail_lines)


def rule_drop_cap(bundle: DocxBundle) -> tuple[bool, str]:
    """+5: 第三段"专业设置如下："所在段落：段首文字设置首字下沉，首字下沉方式为悬挂，
    下沉2行，距正文0.3厘米。首字下沉后段落其余文字仍保持正文楷体小四号、1.1倍行距，
    未造成文字重叠或段落断裂。

    判定按 rubric 逐项验证，不再要求非 rubric 的硬性 XML 结构（如 wrap/vAnchor/
    hAnchor 齐备或"首字必须独占一段"）——这些是 Office UI 保存首字下沉的常见形态
    之一，但并非唯一形态。Word/WPS 里 framePr 也可能与整段挂在一起。

    验证项（对齐 rubric 原文的 7 点）：
      1) 目标位于第三段（按"用户可见段"计数：结构 B 中"首字段+其余文字段"合并计 1 段）
      2) 段首文字设置首字下沉
      3) 悬挂        → w:framePr/@w:dropCap == "margin"
      4) 下沉 2 行   → w:framePr/@w:lines == "2"
      5) 距正文 0.3 厘米 → w:framePr/@w:hSpace == 170 twips (= 0.30 cm)
      6-8) 其余文字仍保持正文楷体 / 小四 / 1.1 倍行距
      9) 未造成文字重叠或段落断裂
    """
    body = get_body_paragraphs(bundle)
    first_char_idx, others_idx = locate_special_dropcap(body)
    if others_idx is None:
        return False, "未找到 '专业设置如下：' 段落"

    target: etree._Element = body[others_idx]
    target_idx: int = others_idx

    # 计算"用户可见段"顺序：跳过空段；结构 B 的首字段与其余文字段合并为一段
    visible_para_start: list[int] = []
    i = 0
    while i < len(body):
        text_i = paragraph_text(body[i]).strip()
        if not text_i:
            i += 1
            continue
        # 结构 B：当前是"首字段"且下一段是"其余文字段"，两段视为同一"用户可见段"
        if i == first_char_idx and others_idx == i + 1:
            visible_para_start.append(i)
            i += 2
            continue
        visible_para_start.append(i)
        i += 1

    # 目标用户可见段的位置（1-based）
    target_start_body_idx = first_char_idx if first_char_idx is not None else others_idx
    try:
        target_visible_pos = visible_para_start.index(target_start_body_idx) + 1
    except ValueError:
        target_visible_pos = -1

    checks: list[tuple[str, bool]] = []

    # 1) 目标位于第三段
    checks.append((
        f"目标位于第三段  实际位置={target_visible_pos}",
        target_visible_pos == 3,
    ))

    # 定位承载 framePr 的段落：结构 A 挂在 target 上，结构 B 挂在其前的"首字段"上
    fr: etree._Element | None = None
    fr_para: etree._Element | None = None
    if first_char_idx is not None:
        fr_para = body[first_char_idx]
    if fr_para is None:
        fr_para = target
    fr = fr_para.find(f"{q('w','pPr')}/{q('w','framePr')}") if fr_para is not None else None
    # 若 framePr 不在预期位置，兜底再看 target 前一段（少数分拆场景）
    if (fr is None or fr.get(q("w", "dropCap")) not in ("margin", "drop")) and target_idx - 1 >= 0:
        prev = body[target_idx - 1]
        pf = prev.find(f"{q('w','pPr')}/{q('w','framePr')}")
        if pf is not None and pf.get(q("w", "dropCap")) in ("margin", "drop"):
            fr = pf
            fr_para = prev

    # 2) 段首文字设置首字下沉
    dropcap_attr = fr.get(q("w", "dropCap")) if fr is not None else None
    checks.append((
        f"段首文字设置首字下沉  framePr={fr is not None} dropCap={dropcap_attr}",
        fr is not None and dropcap_attr in ("margin", "drop"),
    ))

    # 3) 首字下沉方式为悬挂 (dropCap=margin)
    checks.append((
        f"首字下沉方式为悬挂  实际 dropCap={dropcap_attr}",
        dropcap_attr == "margin",
    ))

    # 4) 下沉 2 行
    lines_val = fr.get(q("w", "lines")) if fr is not None else None
    checks.append((
        f"下沉 2 行  实际 lines={lines_val}",
        lines_val == "2",
    ))

    # 5) 距正文 0.3 厘米
    hspace_str = fr.get(q("w", "hSpace")) if fr is not None else None
    hspace_twips: int | None = int(hspace_str) if hspace_str is not None else None
    hspace_ok = hspace_twips == 170
    hspace_cm = (hspace_twips / 20.0 / POINTS_PER_CM) if hspace_twips is not None else None
    checks.append((
        f"距正文 0.3 厘米  实际 hSpace={hspace_twips}twips"
        f"({hspace_cm:.2f}cm)" if hspace_cm is not None else "距正文 0.3 厘米  hSpace 未设置",
        hspace_ok,
    ))

    # 6~8) 其余文字保持楷体 / 小四 / 1.1 倍行距
    # "其余文字段"即 target（结构 A 中即整段；结构 B 中即其余文字所在段）
    others_para: etree._Element = target
    other_runs: list[etree._Element] = [
        r for r in others_para.findall(f"{q('w','r')}")
        if "".join((t.text or "") for t in r.findall(f"{q('w','t')}"))
    ]
    # 结构 A：first_char_idx is None，target 段本身还包含"特殊首字"这一 run，
    # 参照 rubric"其余文字"应剔除首字 run 后再判定。
    if first_char_idx is None and other_runs:
        other_runs = other_runs[1:]

    font_ok = bool(other_runs) and all(
        effective_font_eastasia(r, others_para, bundle) == "楷体" for r in other_runs
    )
    checks.append(("其余文字仍保持正文楷体", font_ok))

    size_ok = bool(other_runs) and all(
        (effective_size_pt(r, others_para, bundle) is not None
         and abs(effective_size_pt(r, others_para, bundle) - 12.0) < 1e-6)
        for r in other_runs
    )
    checks.append(("其余文字仍保持小四号", size_ok))

    multiplier, rule = effective_line_spacing(others_para, bundle)
    line_ok = (rule == "auto") and (multiplier is not None) and abs(multiplier - 1.1) < 1e-6
    checks.append((
        f"其余文字仍保持 1.1 倍行距  实际倍数={multiplier} lineRule={rule}",
        line_ok,
    ))

    # 9) 未造成文字重叠或段落断裂 —— 轻量启发式，不做非 rubric 的结构硬约束：
    #    · hSpace > 0：首字与正文有间距，避免像素重叠
    #    · 其余文字长度不少于下沉行数：Word 下沉高度约等于 lines 行，其余文字过短
    #      时下沉框底部会与后续段拼接出空白/断裂
    #    · 紧邻的下一段不得再挂 dropCap framePr（避免与当前浮动框冲突）
    lines_num = int(lines_val) if lines_val and lines_val.isdigit() else 0
    remainder_len = len(paragraph_text(others_para))
    next_i = target_idx + 1
    next_frame_conflict = False
    if 0 < next_i < len(body):
        next_fr = body[next_i].find(f"{q('w','pPr')}/{q('w','framePr')}")
        next_frame_conflict = (
            next_fr is not None and next_fr.get(q("w", "dropCap")) is not None
        )
    no_overlap_break = (
        (hspace_twips is not None and hspace_twips > 0)
        and (lines_num == 0 or remainder_len >= lines_num)
        and (not next_frame_conflict)
    )
    checks.append((
        f"未造成文字重叠或段落断裂  hSpace>0={hspace_twips not in (None, 0)}, "
        f"其余字符≥下沉行数({remainder_len}≥{lines_num})={lines_num == 0 or remainder_len >= lines_num}, "
        f"后段无 framePr 冲突={not next_frame_conflict}",
        no_overlap_break,
    ))

    passed = all(ok for _, ok in checks)
    return passed, "\n".join(f"{'✓' if ok else '✗'} {msg}" for msg, ok in checks)


def rule_no_wrong_name(bundle: DocxBundle) -> tuple[bool, str]:
    """+1: 文档中没有出现"智能工程系"文本。

    检测思路（对齐 Word/WPS "开始→查找 (Ctrl+F)"的行为）：
      Office 的"查找"只匹配文档可见文本（正文、页眉页脚、脚注/尾注、文本框
      等所有 w:t 节点），不匹配图像里的像素文字。此处严格判定：任何 w:t
      文本片段之间跨节点拼接后不出现子串"智能工程系"。
      注："智能软件工程系"包含"软件"作为中缀，本身不含子串"智能工程系"，
      因此无需人为剔除即可分辨。
    """
    needle = "智能工程系"

    # 收集文档所有 XML 部件中的 w:t 文本，模拟 Office 全文查找覆盖范围。
    parts: list[str] = ["word/document.xml"]
    for name in bundle.zf.namelist():
        if re.fullmatch(
            r"word/(header\d+|footer\d+|footnotes|endnotes|comments)\.xml",
            name,
        ):
            parts.append(name)

    hits: list[str] = []
    for name in parts:
        try:
            root = etree.fromstring(bundle.zf.read(name))
        except Exception:
            continue
        # 逐段落拼接 w:t —— 对齐 Word 查找"跨 run 匹配"的行为；跨段落不算连续。
        for p in root.iter(q("w", "p")):
            joined = "".join((t.text or "") for t in p.iter(q("w", "t")))
            if needle in joined:
                hits.append(f"{name}: '{joined[:60]}'")

    triggered = bool(hits)
    if triggered:
        detail = "仍出现'智能工程系'文本：" + "; ".join(hits[:5])
        if len(hits) > 5:
            detail += f"; ...（共 {len(hits)} 处）"
    else:
        detail = "未在任何 w:t 文本中检出子串'智能工程系'"
    return (not triggered), detail


def _paragraph_char_run_map(
    p: "etree._Element",
) -> tuple[str, list["etree._Element"]]:
    """把段落的可见文本按字符位置映射回 run。

    只拼接 w:r/w:t 文本节点（不含 tab/br 等结构化字符），因此段落文本索引与
    每个字符所属的 run 严格对齐。返回 (text, char_to_run)，长度一致：
    char_to_run[i] 即 text[i] 所属的 w:r 元素。
    """
    text_parts: list[str] = []
    char_to_run: list[etree._Element] = []
    for r in p.findall(q("w", "r")):
        run_text = "".join((t.text or "") for t in r.findall(q("w", "t")))
        if not run_text:
            continue
        text_parts.append(run_text)
        char_to_run.extend([r] * len(run_text))
    return "".join(text_parts), char_to_run


def rule_special_run_format(bundle: DocxBundle) -> tuple[bool, str]:
    """+1: 正文中的"智能软件工程系"：采用宋体、四号、加粗、倾斜、绿色，并添加着重号。

    每一条细则的检测思路（对齐 Word/WPS "开始→字体"面板的行为）：
      · 宋体         -> run 的 rFonts eastAsia 生效值 == "宋体"
      · 四号         -> w:sz 生效值 == 28 半磅 (=14pt，Office 中"四号"精确对应值)
      · 加粗         -> w:b toggle 生效为真
      · 倾斜         -> w:i toggle 生效为真
      · 绿色         -> w:color 生效值判定为绿色（G 通道显著高于 R、B），
                        并支持通过 themeColor 解析主题色（对齐 Office UI 里
                        "标准色→绿色"与"主题颜色"两条路径）
      · 添加着重号   -> w:em 生效值存在且不为 "none"（Office UI "着重号"下拉的
                        任一非"(无)"选项：dot/comma/circle/underDot）

    "正文中的"智能软件工程系"":按段落可见文本拼接后定位所有出现位置，再映射
    回覆盖该子串的 run 集合；每一处出现被覆盖的每个 run 都必须同时满足上述 6 项。
    """
    target_text = "智能软件工程系"
    body = get_body_paragraphs(bundle)

    # 收集所有"出现"，每个"出现"记录 (段落, 覆盖它的 run 集合[保序去重])
    occurrences: list[tuple[etree._Element, list[etree._Element]]] = []
    for p in body:
        para_text, char_to_run = _paragraph_char_run_map(p)
        if not para_text:
            continue
        start = 0
        while True:
            pos = para_text.find(target_text, start)
            if pos < 0:
                break
            covered: list[etree._Element] = []
            seen: set[int] = set()
            for i in range(pos, pos + len(target_text)):
                r = char_to_run[i]
                rid = id(r)
                if rid not in seen:
                    seen.add(rid)
                    covered.append(r)
            occurrences.append((p, covered))
            start = pos + len(target_text)

    if not occurrences:
        return False, "正文中未找到 '智能软件工程系'"

    total = len(occurrences)
    stats = dict(font=0, size=0, bold=0, italic=0, color=0, em=0)
    failures: list[str] = []

    for idx, (p, runs) in enumerate(occurrences, 1):
        # 每一处出现：所覆盖的每个 run 都必须满足全部 6 项属性
        per_ok = dict(font=True, size=True, bold=True, italic=True, color=True, em=True)
        for r in runs:
            # 1) 宋体
            font = effective_font_eastasia(r, p, bundle)
            if font != "宋体":
                per_ok["font"] = False
                failures.append(f"第{idx}处 有 run 字体非宋体：{font}")

            # 2) 四号 (=14pt，严格)
            size = effective_size_pt(r, p, bundle)
            if size is None or abs(size - 14.0) > 1e-6:
                per_ok["size"] = False
                failures.append(f"第{idx}处 有 run 字号非四号：{size}pt")

            # 3) 加粗
            if not effective_bold(r, p, bundle):
                per_ok["bold"] = False
                failures.append(f"第{idx}处 有 run 未加粗")

            # 4) 倾斜
            if not effective_italic(r, p, bundle):
                per_ok["italic"] = False
                failures.append(f"第{idx}处 有 run 未倾斜")

            # 5) 绿色：G 通道明显高于 R、B（覆盖 Office"标准色→绿色"#00B050、
            #    "深绿"#006100、经典"绿色"#008000 等），并支持 themeColor 解析
            hex_val, theme_val = effective_color(r, p, bundle)
            color_hex = hex_val if hex_val and hex_val.upper() != "AUTO" else _theme_hex(bundle, theme_val)
            color_ok = False
            if color_hex:
                try:
                    rr = int(color_hex[0:2], 16)
                    gg = int(color_hex[2:4], 16)
                    bb = int(color_hex[4:6], 16)
                    color_ok = gg >= 100 and gg >= rr + 30 and gg >= bb + 30
                except Exception:
                    color_ok = False
            if not color_ok:
                per_ok["color"] = False
                failures.append(f"第{idx}处 有 run 非绿色：#{hex_val} themeColor={theme_val}")

            # 6) 添加着重号：w:em 存在且非 "none"
            em_val = effective_em(r, p, bundle)
            if not (em_val is not None and em_val != "none"):
                per_ok["em"] = False
                failures.append(f"第{idx}处 有 run 未添加着重号：w:em={em_val}")

        for k, v in per_ok.items():
            if v:
                stats[k] += 1

    detail_lines = [
        f"正文中共发现 {total} 处 '智能软件工程系'（按段落文本拼接后定位，可跨 run）：",
        f"  宋体 {stats['font']}/{total}, 四号 {stats['size']}/{total}, "
        f"加粗 {stats['bold']}/{total}, 倾斜 {stats['italic']}/{total}, "
        f"绿色 {stats['color']}/{total}, 着重号 {stats['em']}/{total}",
    ]
    passed = all(v == total for v in stats.values())
    if not passed and failures:
        detail_lines.append("失败明细：")
        for msg in failures[:8]:
            detail_lines.append(f"  - {msg}")
        if len(failures) > 8:
            detail_lines.append(f"  ...（共 {len(failures)} 条）")
    return passed, "\n".join(detail_lines)


def find_watermark_imagedata(bundle: DocxBundle) -> tuple[etree._Element | None, str | None, str | None]:
    """在页眉中查找水印 v:shape，返回 (shape, image_target, header_name)。"""
    for hname, hxml in bundle.headers.items():
        # 传统 VML 水印：<v:shape><v:imagedata r:id="..."/>
        for shape in hxml.findall(f".//{q('v','shape')}"):
            imagedata = shape.find(f"{q('v','imagedata')}")
            if imagedata is None:
                continue
            rid = imagedata.get(q("r", "id"))
            # 解析 rId -> target
            rels_name = f"word/_rels/{os.path.basename(hname)}.rels"
            target = None
            if rels_name in bundle.zf.namelist():
                rels = etree.fromstring(bundle.zf.read(rels_name))
                for rel in rels.findall(".//{http://schemas.openxmlformats.org/package/2006/relationships}Relationship"):
                    if rel.get("Id") == rid:
                        target = rel.get("Target")
                        break
            return shape, target, hname
        # DrawingML 水印
        for drawing in hxml.findall(f".//{q('w','drawing')}"):
            blip = drawing.find(f".//{{{NS['a']}}}blip")
            if blip is not None:
                return drawing, None, hname
    return None, None, None


def rule_watermark_scale(bundle: DocxBundle) -> tuple[bool, str]:
    """+5: 文档水印背景：校园横幅图片被设置为整篇简介文档的水印背景，
    显示在正文后方，缩放比例为100%或接近100%。

    每一条细则的检测思路（对齐 Word/WPS "设计→水印→自定义水印→图片水印" 面板的行为）：
      · 校园横幅图片            -> 水印引用的图片存在、可解析为图像，且长宽比呈"横幅"
                                    形态（宽/高 ≥ 1.6，Office 中"横幅"素材的典型比例）
      · 被设置为整篇简介文档的水印背景 -> 水印在页眉里(w:hdr)，通过节属性(sectPr)
                                    被文档的每个 section 引用（Word 图片水印的实现方式）
      · 显示在正文后方          -> VML 水印 z-index < 0，或 DrawingML behindDoc="1"
                                    （对应 Office UI 中"衬于文字下方"）
      · 缩放比例为100%或接近100% -> Office"图片水印"对话框里"缩放"下拉的值。
                                    从 XML 反算：display_size / native_size × 100%
                                    (native = 像素 ÷ DPI × 72pt/inch)；接受 85%~115%
    """
    shape, target, hname = find_watermark_imagedata(bundle)
    if shape is None:
        return False, "未在页眉中找到水印图片对象"

    checks: list[tuple[str, bool]] = []

    # 解析显示尺寸 (VML style 或 DrawingML wp:extent)
    style = shape.get("style") or ""
    m_w = re.search(r"width:([\d.]+)pt", style)
    m_h = re.search(r"height:([\d.]+)pt", style)
    display_w_pt: float | None = float(m_w.group(1)) if m_w else None
    display_h_pt: float | None = float(m_h.group(1)) if m_h else None
    # DrawingML 兜底：<wp:extent cx="EMU" cy="EMU"/>
    if display_w_pt is None:
        extent = shape.find(f".//{q('wp','extent')}")
        if extent is not None:
            cx = extent.get("cx"); cy = extent.get("cy")
            if cx and cy:
                display_w_pt = int(cx) / EMU_PER_PT
                display_h_pt = int(cy) / EMU_PER_PT

    # 打开图片，读取像素与 DPI
    orig_w = orig_h = None
    orig_dpi_x = orig_dpi_y = None
    orig_ratio: float | None = None
    if target:
        for name in (f"word/{target.lstrip('/')}", f"word/{target}"):
            name = name.replace("word/word/", "word/")
            if name in bundle.zf.namelist():
                with bundle.zf.open(name) as f:
                    img = Image.open(io.BytesIO(f.read()))
                    orig_w, orig_h = img.size
                    orig_dpi_x, orig_dpi_y = img.info.get("dpi", (96, 96))
                orig_ratio = orig_w / orig_h if orig_w and orig_h else None
                break

    # 1) 校园横幅图片：水印引用的图片可解析，且长宽比呈横幅形态
    banner_ok = (
        target is not None
        and orig_w is not None and orig_h is not None
        and orig_ratio is not None
        and orig_ratio >= 1.6  # Office 中"横幅/banner"类图片的典型宽高比
    )
    checks.append((
        f"校园横幅图片  目标={target} 尺寸={orig_w}x{orig_h}px 长宽比={orig_ratio}",
        banner_ok,
    ))

    # 2) 被设置为整篇简介文档的水印背景：水印在页眉，且所有节都引用该页眉
    #    Word "图片水印" 就是在页眉插入 v:shape/v:imagedata，通过 sectPr→headerReference 生效。
    in_header = hname is not None and re.fullmatch(r"word/header\d+\.xml", hname) is not None

    header_rid = None
    if hname is not None:
        # 通过反查 document.xml.rels 找到指向该页眉的 rId
        if "word/_rels/document.xml.rels" in bundle.zf.namelist():
            doc_rels = etree.fromstring(bundle.zf.read("word/_rels/document.xml.rels"))
            for rel in doc_rels.findall(
                ".//{http://schemas.openxmlformats.org/package/2006/relationships}Relationship"
            ):
                if rel.get("Target") in (os.path.basename(hname), f"/{hname}", hname.replace("word/", "")):
                    header_rid = rel.get("Id")
                    break

    sect_prs = bundle.document.findall(f".//{q('w','sectPr')}")
    if not sect_prs:
        sect_prs = []
    all_sections_ref = False
    if sect_prs and header_rid:
        all_sections_ref = all(
            any(
                hr.get(q("r", "id")) == header_rid
                for hr in sp.findall(f"{q('w','headerReference')}")
            )
            for sp in sect_prs
        )
    else:
        # 兼容单节且未显式引用的情形：Word 在保存图片水印时会自动写入 headerReference；
        # 若只有一个 section 且页眉是唯一 header，视为覆盖整篇。
        header_files = [n for n in bundle.zf.namelist() if re.fullmatch(r"word/header\d+\.xml", n)]
        all_sections_ref = len(sect_prs) <= 1 and len(header_files) == 1 and in_header

    background_ok = in_header and all_sections_ref
    checks.append((
        f"被设置为整篇简介文档的水印背景  水印在页眉={in_header} 每节引用该页眉={all_sections_ref}",
        background_ok,
    ))

    # 3) 显示在正文后方：VML z-index<0，或 DrawingML behindDoc="1"
    behind_doc = False
    z_match = re.search(r"z-index:\s*(-?\d+)", style)
    if z_match:
        behind_doc = int(z_match.group(1)) < 0
    if not behind_doc:
        anchor = shape.find(f".//{q('wp','anchor')}") if shape is not None else None
        if anchor is not None:
            behind_doc = anchor.get("behindDoc") in ("1", "true")
    checks.append((
        f"显示在正文后方  z-index={z_match.group(1) if z_match else None} behindDoc/负z-index生效={behind_doc}",
        behind_doc,
    ))

    # 4) 缩放比例为 100% 或接近 100%
    #    对齐 Office UI "设置图片格式→大小→缩放" 显示的数值，即：
    #        缩放% = 显示尺寸(pt) / 原生尺寸(pt) × 100
    #        原生尺寸(pt) = 像素 / DPI × 72
    #    该值就是用户在办公软件里肉眼看到的百分比（本文档 Office 显示 41%）。
    #    "接近 100%" 保守取 ±15%。
    scale_note = "无法计算缩放比例"
    scale_ok = False
    if display_w_pt and orig_w and orig_h and orig_dpi_x and orig_dpi_y:
        native_w_pt = orig_w / float(orig_dpi_x) * 72.0
        native_h_pt = orig_h / float(orig_dpi_y) * 72.0
        scale_w = display_w_pt / native_w_pt * 100
        scale_h = (display_h_pt / native_h_pt * 100) if display_h_pt else scale_w
        scale_ok = 85.0 <= scale_w <= 115.0 and 85.0 <= scale_h <= 115.0
        scale_note = (
            f"原图 {orig_w}x{orig_h}px @ {orig_dpi_x}dpi -> 原生 "
            f"{native_w_pt:.1f}x{native_h_pt:.1f}pt；显示 {display_w_pt:.1f}x{display_h_pt}pt；"
            f"Office 缩放≈{scale_w:.1f}%(宽)/{scale_h:.1f}%(高)"
        )
    checks.append((f"缩放比例为100%或接近100%  {scale_note}", scale_ok))

    passed = all(ok for _, ok in checks)
    return passed, "\n".join(f"{'✓' if ok else '✗'} {msg}" for msg, ok in checks)


def rule_watermark_washout(bundle: DocxBundle) -> tuple[bool, str]:
    """+3: 文档水印效果：水印图片应用"冲蚀"或明显淡化效果，
    正文文字在水印上方仍清晰可读。

    对齐 Office/WPS "图片水印"对话框里"冲蚀"复选框的实际效果：
      · Word 勾选"冲蚀"后，会在 VML <v:imagedata> 上写入
            o:gain="19661f"       (=0.4，亮度提升)
            o:blacklevel="22938f" (=0.35，黑阶抬升)
        这两个属性是 Office 内部把"冲蚀"落到 XML 的**唯一渠道**。
        如果这两个属性缺失，Office 在打开时不会应用冲蚀滤镜，
        用户在 UI 上就看不到冲蚀效果（即便图片本身颜色较浅）。
      · WPS 保存"图片水印→冲蚀"输出等价的 gain/blacklevel。

    因此本项**只**认这两个 Office 属性作为"应用了冲蚀"的证据；
    图片文件本身预先偏亮偏灰不等于 Office 应用了冲蚀效果。
    """
    shape, _target, _ = find_watermark_imagedata(bundle)
    if shape is None:
        return False, "未找到水印图片"

    checks: list[tuple[str, bool]] = []

    # ---------- 点 1：水印图片应用"冲蚀"或明显淡化效果 ----------
    imagedata = shape.find(f"{q('v','imagedata')}")
    gain = blacklevel = chromakey = None
    if imagedata is not None:
        gain = imagedata.get(f"{{{NS['o']}}}gain") or imagedata.get("gain")
        blacklevel = imagedata.get(f"{{{NS['o']}}}blacklevel") or imagedata.get("blacklevel")
        chromakey = imagedata.get(f"{{{NS['o']}}}chromakey") or imagedata.get("chromakey")

    # Office 冲蚀的判定：至少有 gain 或 blacklevel（Word "冲蚀"预设）
    washout_applied = gain is not None or blacklevel is not None
    msg = (
        f"水印图片应用'冲蚀'或明显淡化效果  "
        f"Office 冲蚀属性 o:gain={gain} o:blacklevel={blacklevel} o:chromakey={chromakey}"
    )
    checks.append((msg, washout_applied))

    # ---------- 点 2：正文文字在水印上方仍清晰可读 ----------
    body = get_body_paragraphs(bundle)
    has_text = False
    fg_ok = False
    sample_color: str | None = None
    for p in body:
        text = paragraph_text(p).strip()
        if not text:
            continue
        has_text = True
        for r in p.findall(f"{q('w','r')}"):
            if "".join((t.text or "") for t in r.findall(f"{q('w','t')}")):
                color = (run_color_hex(r) or "000000").upper()
                sample_color = color
                try:
                    rr = int(color[0:2], 16); gg = int(color[2:4], 16); bb = int(color[4:6], 16)
                    fg_ok = not (rr >= 240 and gg >= 240 and bb >= 240)
                except Exception:
                    fg_ok = color not in ("FFFFFF",)
                break
        if fg_ok:
            break
    readable = has_text and fg_ok
    checks.append((
        f"正文文字在水印上方仍清晰可读  正文有文字={has_text} 首个正文run颜色=#{sample_color} 前景非白={fg_ok}",
        readable,
    ))

    passed = all(ok for _, ok in checks)
    return passed, "\n".join(f"{'✓' if ok else '✗'} {msg}" for msg, ok in checks)


def rule_watermark_range(bundle: DocxBundle) -> tuple[bool, str]:
    """+1: 文档水印范围：水印出现在页面背景区域，
    不作为正文中的普通插入图片占用段落位置，不挤压正文排版。

    对齐 Office/WPS 中"图片水印"的行为：
      · 页面背景区域 —— Word/WPS 的图片水印固定放在页眉(w:hdr)里；
                        VML 通过 v:shape 的 position:absolute + mso-position-*-relative
                        绝对定位到页面/页边距，脱离正文流。
      · 不作为正文中的普通插入图片占用段落位置
                     —— 正文 (w:body) 直接子层的段落 run 中不包含指向同一水印图
                        的 <w:pict>/<w:drawing>（若有也不得是 inline 而必须是 anchor 且
                        behindDoc="1"）。
      · 不挤压正文排版
                     —— VML: <w10:wrap type="none"/> 或不存在 wrap（文字不环绕）；
                        DrawingML: <wp:anchor behindDoc="1"> + <wp:wrapNone/>。
    """
    shape, target, hname = find_watermark_imagedata(bundle)
    if shape is None:
        return False, "未找到水印"

    checks: list[tuple[str, bool]] = []

    # ---------- 点 1：水印出现在页面背景区域 ----------
    #   a) 位于页眉 xml（Office 图片水印的固定载体）
    #   b) VML style 中 position:absolute，且 mso-position-*-relative 绑到 page/margin
    in_header = hname is not None and re.fullmatch(r"word/header\d+\.xml", hname) is not None
    style = shape.get("style") or ""
    style_l = style.lower().replace(" ", "")
    is_absolute = "position:absolute" in style_l
    hrel = re.search(r"mso-position-horizontal-relative:([a-z\-]+)", style_l)
    vrel = re.search(r"mso-position-vertical-relative:([a-z\-]+)", style_l)
    hrel_v = hrel.group(1) if hrel else None
    vrel_v = vrel.group(1) if vrel else None
    # Office 保存图片水印时把水平/垂直相对参考写成 page 或 margin
    rel_ok = (
        (hrel_v in ("page", "margin", "left-margin", "right-margin") or hrel_v is None)
        and (vrel_v in ("page", "margin", "top-margin", "bottom-margin") or vrel_v is None)
    )
    bg_ok = in_header and (is_absolute or rel_ok)
    checks.append((
        f"水印出现在页面背景区域  页眉={hname} position:absolute={is_absolute} "
        f"mso-position-horizontal-relative={hrel_v} mso-position-vertical-relative={vrel_v}",
        bg_ok,
    ))

    # ---------- 点 2：不作为正文中的普通插入图片占用段落位置 ----------
    #    正文 w:body 里不得存在与水印图相同 rId 目标、且是"inline"或"anchor 但非 behindDoc"
    #    的图片 —— 那样才会像普通插图一样占据段落位置。
    body = bundle.document.find(f".//{q('w','body')}")
    body_pics: list[etree._Element] = []
    if body is not None:
        body_pics = body.findall(f".//{q('w','pict')}") + body.findall(f".//{q('w','drawing')}")
    # 过滤掉 body 内声明为背景层的图片（behindDoc="1" 或 z-index<0），
    # 只保留"占段落位置"的普通插入图片
    occupying_paragraph = 0
    for pic in body_pics:
        # DrawingML: <w:drawing><wp:inline .../></w:drawing> 一律占位
        if pic.find(f".//{q('wp','inline')}") is not None:
            occupying_paragraph += 1
            continue
        # DrawingML anchor：不是 behindDoc 且没有 wrapNone —— 会挤压正文
        anchor = pic.find(f".//{q('wp','anchor')}")
        if anchor is not None:
            behind = anchor.get("behindDoc") in ("1", "true")
            wrap_none = anchor.find(f"{q('wp','wrapNone')}") is not None
            if not behind and not wrap_none:
                occupying_paragraph += 1
            continue
        # VML pict：z-index 非负 或 wrap 非 none => 占位
        s = pic.find(f".//{q('v','shape')}")
        if s is not None:
            s_style = (s.get("style") or "").lower()
            z_m = re.search(r"z-index:\s*(-?\d+)", s_style)
            z = int(z_m.group(1)) if z_m else 0
            wrap = pic.find(f".//{{{NS.get('w10','urn:schemas-microsoft-com:office:word')}}}wrap")
            wrap_ok = wrap is not None and (wrap.get("type") == "none")
            if z >= 0 and not wrap_ok:
                occupying_paragraph += 1
    not_inline_in_body = occupying_paragraph == 0
    checks.append((
        f"不作为正文中的普通插入图片占用段落位置  "
        f"body 内 pict/drawing 总数={len(body_pics)} 其中占段落位置的={occupying_paragraph}",
        not_inline_in_body,
    ))

    # ---------- 点 3：不挤压正文排版 ----------
    #    水印 shape 上：VML <w10:wrap type="none"/> 或不存在 wrap；DrawingML wp:wrapNone
    w10_ns = NS.get("w10", "urn:schemas-microsoft-com:office:word")
    w10_wrap = shape.find(f".//{{{w10_ns}}}wrap")
    wrap_type = w10_wrap.get("type") if w10_wrap is not None else None
    dml_wrap_none = shape.find(f".//{q('wp','wrapNone')}") is not None
    # "无环绕" 意味着不挤压排版：type="none" / 缺省 / DrawingML wrapNone
    no_squeeze = (
        (w10_wrap is None and not dml_wrap_none)  # VML 无 wrap 元素默认不影响正文
        or wrap_type == "none"
        or dml_wrap_none
    )
    checks.append((
        f"不挤压正文排版  w10:wrap type={wrap_type} wp:wrapNone={dml_wrap_none}",
        no_squeeze,
    ))

    passed = all(ok for _, ok in checks)
    return passed, "\n".join(f"{'✓' if ok else '✗'} {msg}" for msg, ok in checks)


# ---- 扣分点 -----------------------------------------------------------------
# （原 -5 漏字/错字 与 -3 水印裁切/变形/覆盖 两条按用户要求删除）


# ---------------------------------------------------------------------------
# 主流程
# ---------------------------------------------------------------------------


def evaluate(dir_path: str) -> dict[str, object]:
    """统一入口：接收脚本所在目录的路径，脚本自行在该目录内定位 .docx 并评估。

    返回结构（详见《脚本接口差异与统一建议》§2.2）：
        {
            "id": "040",
            "file_name": <被评估文档名>,
            "status": "ok" | "error",
            "error": None | <错误信息>,
            "dim1_pass": bool,
            "dim1_reason": "" | <未通过原因>,
            "dim2_items": [ {rule, max_delta, delta, hit, detail}, ... ],
            "total_score": int,
            "max_score": int,
        }
    """
    # 加分项元数据：rule_id 使用打分细则的原文；max_delta 为该项满分。
    plus_rules: list[tuple[str, int, Callable[[DocxBundle], tuple[bool, str]]]] = [
        ("文档标题段落：\"智能软件工程系简介\"位于文档首行，采用隶书、二号、加粗、蓝色、居中，段后间距为20磅。", 1, rule_title_basic_format),
        ("文档标题段落：\"智能软件工程系简介\"设置发光效果，颜色为橄榄色或接近橄榄色，大小约7磅，使用强调文字颜色3或接近效果；设置向上偏移阴影效果；标题文字仍清晰可读。", 3, rule_title_glow_shadow),
        ("正文段落格式：除标题、\"专业设置如下：\"、\"计算机应用技术\"内容段落外，其余正文首行缩进2字符、段后6磅、1.1倍行距。", 1, rule_body_para_format),
        ("正文字体格式：除标题、\"专业设置如下：\"内容段落特殊首字和被替换文字\"智能软件工程系\"特殊格式外，正文采用楷体、小四号。", 1, rule_body_font_format),
        ("第三段\"专业设置如下：\"所在段落：段首文字设置首字下沉，方式为悬挂，下沉2行，距正文0.3厘米；其余文字仍保持正文楷体、小四号、1.1倍行距；未造成文字重叠或段落断裂。", 5, rule_drop_cap),
        ("文档中没有出现\"智能工程系\"文本。", 1, rule_no_wrong_name),
        ("正文中的\"智能软件工程系\"：采用宋体、四号、加粗、倾斜、绿色，并添加着重号。", 1, rule_special_run_format),
        ("文档水印背景：校园横幅图片被设置为整篇简介文档的水印背景，显示在正文后方，缩放比例为100%或接近100%。", 5, rule_watermark_scale),
        ("文档水印效果：水印图片应用\"冲蚀\"或明显淡化效果，正文文字在水印上方仍清晰可读。", 3, rule_watermark_washout),
        ("文档水印范围：水印出现在页面背景区域，不作为正文中的普通插入图片占用段落位置，不挤压正文排版。", 1, rule_watermark_range),
    ]

    dim2_items: list[dict[str, object]] = []
    result: dict[str, object] = {
        "id": "040",
        "file_name": "",
        "status": "ok",
        "error": None,
        "dim1_pass": True,
        "dim1_reason": "",
        "dim2_items": dim2_items,
        "total_score": 0,
        "max_score": sum(s for _, s, _ in plus_rules),
    }

    try:
        # 在目录中定位 .docx（排除 Office 临时锁文件 ~$*.docx）
        if not os.path.isdir(dir_path):
            raise FileNotFoundError(f"目录不存在：{dir_path}")
        docx_files = [
            f for f in os.listdir(dir_path)
            if f.lower().endswith(".docx") and not f.startswith("~$")
        ]
        if not docx_files:
            raise FileNotFoundError(f"目录中未找到 .docx 文件：{dir_path}")
        file_name = docx_files[0]
        result["file_name"] = file_name
        file_path = os.path.join(dir_path, file_name)

        report = Report()
        bundle = load_bundle(file_path)
        try:
            check_dimension1(bundle, report)
            if not report.dim1_passed:
                result["dim1_pass"] = False
                result["dim1_reason"] = "；".join(report.dim1_reasons)
                result["total_score"] = 0
                return result

            total = 0
            for rid, max_delta, fn in plus_rules:
                passed, detail = fn(bundle)
                delta = max_delta if passed else 0
                total += delta
                # detail 字段按需求置空；上方仍保留 fn(bundle) 的完整计算，
                # 以保证评分逻辑与结构不变。
                _ = detail
                dim2_items.append({
                    "rule": rid,
                    "max_delta": max_delta,
                    "delta": delta,
                    "hit": passed,
                    "detail": "",
                })
            result["total_score"] = total
            # 扣分项已按用户要求全部移除
        finally:
            bundle.close()
    except Exception as e:
        result["status"] = "error"
        result["error"] = f"{type(e).__name__}: {e}"
        dim2_items.clear()
        result["total_score"] = 0
    return result


if __name__ == "__main__":
    # 仅用于本地调试：入参为脚本所在目录（缺省则取本脚本所在目录）
    target = sys.argv[1] if len(sys.argv) > 1 else os.path.dirname(os.path.abspath(__file__))
    _payload = json.dumps(evaluate(target), ensure_ascii=False, indent=2)
    # 直接写 UTF-8 字节，避免 Windows 终端 cp1252 打印中文时抛 UnicodeEncodeError；
    # 不修改 sys.stdout（遵循接口规约 §2.3）。
    _ = sys.stdout.buffer.write(_payload.encode("utf-8"))
    _ = sys.stdout.buffer.write(b"\n")
