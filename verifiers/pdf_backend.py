#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
PDF 解析适配层。

基于宽松许可证库实现：
  - pdfplumber (MIT)：文本、字符、坐标、字体、颜色、矢量对象；
  - pypdfium2 (Apache-2.0 / BSD-3-Clause)：页面渲染、旋转等 PDFium 能力
    （按需在后续迁移阶段接入，本模块当前不强制依赖）。

坐标约定：原点在页面左上角，y 向下增大，单位 pt。
pdfplumber 的 chars/words 坐标本身即为 top-left 原点（top/bottom 字段），
与本模块的 y0/y1 直接对应。

验证脚本不应直接 import pdfplumber / pypdfium2，统一经由本模块访问。
"""

from __future__ import annotations

from dataclasses import dataclass


@dataclass(frozen=True)
class PdfRect:
    """页面矩形，pt 单位，top-left 原点。"""
    x0: float
    y0: float
    x1: float
    y1: float

    @property
    def width(self) -> float:
        return self.x1 - self.x0

    @property
    def height(self) -> float:
        return self.y1 - self.y0


@dataclass(frozen=True)
class PdfTextLine:
    """一条视觉文本行：按 y 坐标聚合后的行文本与包围盒。"""
    text: str
    bbox: PdfRect


@dataclass(frozen=True)
class PdfTextBlock:
    """文本块：由相邻视觉行聚合而成的段落级文本区域。"""
    text: str
    bbox: PdfRect


@dataclass(frozen=True)
class PdfSpanLine:
    """
    按内容流重建的文本行（带字体/字号/颜色属性）。

    同一位置重复绘制的文本（伪粗体/阴影的常用实现）保留为多条独立行，
    由调用方自行去重或统计重复次数。

    font/color 取行内首个 span；color 为 0xRRGGBB 整数；
    flags 为兼容保留字段，恒为 0。
    """
    text: str
    bbox: PdfRect
    size: float
    font: str
    color: int
    flags: int = 0


@dataclass(frozen=True)
class PdfDrawing:
    """矢量绘图对象：外接矩形、填充色与描边属性（颜色为 0~1 浮点 RGB）。"""
    rect: PdfRect
    fill: tuple[float, float, float] | None
    kind: str = "rect"
    stroke_color: tuple[float, float, float] | None = None
    line_width: float = 0.0


@dataclass(frozen=True)
class PdfRawChar:
    """单个字符：文本与精确包围盒。"""
    c: str
    bbox: PdfRect


@dataclass(frozen=True)
class PdfRawSpan:
    """span：同字体/字号/颜色的连续字符串，含字符级 bbox。"""
    text: str
    bbox: PdfRect
    font: str
    size: float
    color: int
    chars: tuple[PdfRawChar, ...]


@dataclass(frozen=True)
class PdfRawLine:
    """raw 文本行：span 序列及整行 bbox。"""
    text: str
    bbox: PdfRect
    spans: tuple[PdfRawSpan, ...]


@dataclass(frozen=True)
class PdfImage:
    """页面嵌入图像的一次放置：放置矩形与源像素尺寸。"""
    rect: PdfRect
    src_width: int
    src_height: int


@dataclass(frozen=True)
class PdfPath:
    """
    路径级矢量对象：外接矩形、填充/描边色与构成 items。

    items 为元组序列：
      ("re", PdfRect)              轴对齐矩形子路径
      ("l", (x0, y0), (x1, y1))    线段（含曲线锚点间的近似线段）
    坐标为 top-left 原点 pt。
    """
    rect: PdfRect
    fill: tuple[float, float, float] | None
    stroke: tuple[float, float, float] | None
    line_width: float
    items: tuple


def normalize_font_name(name: str) -> str:
    """去掉子集前缀（如 'AAAAAA+SimSun' -> 'SimSun'）。"""
    if not name:
        return ""
    return name.split("+", 1)[-1]


def color_to_float_rgb(color) -> tuple[float, float, float] | None:
    """把底层颜色值统一为 (r, g, b) 浮点三元组，范围 0~1。"""
    if color is None:
        return None
    if isinstance(color, (int, float)):
        v = float(color)
        return (v, v, v)
    try:
        vals = [float(v) for v in color]
    except (TypeError, ValueError):
        return None
    if len(vals) == 1:
        return (vals[0], vals[0], vals[0])
    if len(vals) == 3:
        return (vals[0], vals[1], vals[2])
    if len(vals) == 4:
        c, m, y, k = vals
        return ((1 - c) * (1 - k), (1 - m) * (1 - k), (1 - y) * (1 - k))
    return None


def color_to_int(color) -> int:
    """把底层颜色值统一为 0xRRGGBB 整数。无法识别时返回 0（黑色）。"""
    rgb = color_to_float_rgb(color)
    if rgb is None:
        return 0
    r, g, b = (max(0, min(255, round(v * 255))) for v in rgb)
    return (r << 16) | (g << 8) | b


class PdfDocument:
    """PDF 文档句柄。用上下文管理器确保关闭底层文件。"""

    def __init__(self, path: str):
        import pdfplumber
        self._path = path
        self._pdf = pdfplumber.open(path)
        self._pdfium = None  # 渲染用 pypdfium2 句柄，首次渲染时打开

    # ---- 生命周期 ----

    def close(self) -> None:
        try:
            self._pdf.close()
        except Exception:
            pass
        if self._pdfium is not None:
            try:
                self._pdfium.close()
            except Exception:
                pass
            self._pdfium = None

    def __enter__(self) -> "PdfDocument":
        return self

    def __exit__(self, exc_type, exc, tb) -> None:
        self.close()

    # ---- 基本信息 ----

    @property
    def page_count(self) -> int:
        return len(self._pdf.pages)

    def page_size(self, page_index: int) -> tuple[float, float]:
        """(width_pt, height_pt)"""
        page = self._pdf.pages[page_index]
        return float(page.width), float(page.height)

    def page_text(self, page_index: int) -> str:
        """页面纯文本（触发一次完整解析，可用于损坏检测）。"""
        page = self._pdf.pages[page_index]
        return page.extract_text() or ""

    def page_rotation(self, page_index: int) -> int:
        """页面 /Rotate 值（0/90/180/270）。"""
        return int(self._pdf.pages[page_index].rotation or 0)

    def count_annotations(self, page_index: int) -> int:
        """页面注释（批注/标注等 annotation 对象）数量。"""
        try:
            return len(self._pdf.pages[page_index].annots or [])
        except Exception:
            return 0

    def extract_images(self, page_index: int) -> list[PdfImage]:
        """页面嵌入图像的全部放置（每次放置一条，含源像素尺寸）。"""
        page = self._pdf.pages[page_index]
        out: list[PdfImage] = []
        for im in page.images:
            src = im.get("srcsize") or (0, 0)
            out.append(PdfImage(
                rect=PdfRect(float(im["x0"]), float(im["top"]),
                             float(im["x1"]), float(im["bottom"])),
                src_width=int(src[0] or 0),
                src_height=int(src[1] or 0),
            ))
        return out

    def render_page(self, page_index: int, dpi: float = 150.0):
        """
        用 pypdfium2 渲染页面为 RGB ndarray (H, W, 3)，已计入页面旋转。

        渲染是唯一需要 pypdfium2 的能力，延迟到首次调用时才导入并打开。
        """
        import pypdfium2 as pdfium
        if self._pdfium is None:
            self._pdfium = pdfium.PdfDocument(self._path)
        page = self._pdfium[page_index]
        bitmap = page.render(scale=dpi / 72.0, rev_byteorder=True)  # RGB 通道序
        arr = bitmap.to_numpy()
        if arr.shape[2] == 4:
            arr = arr[:, :, :3]
        return arr.copy()

    # ---- 文本提取 ----

    def extract_text_lines(self, page_index: int,
                           y_tol: float = 3.0) -> list[PdfTextLine]:
        """
        按 y 坐标聚合的视觉行。

        以字符 bottom（基线附近）坐标聚类：同一视觉行内不同字号的字符
        top 值可能差异较大，但底部基本对齐，因此按 bottom 差值在 y_tol
        内视为同一行；行内字符按 x0 排序拼接，行 bbox 为字符 bbox 并集。
        """
        page = self._pdf.pages[page_index]
        chars = page.chars
        if not chars:
            return []

        # 先按 bottom 再按 x0 排序，逐字符聚类到行
        chars = sorted(chars, key=lambda c: (c["bottom"], c["x0"]))
        rows: list[list[dict]] = []
        for ch in chars:
            if rows and abs(rows[-1][0]["bottom"] - ch["bottom"]) <= y_tol:
                rows[-1].append(ch)
            else:
                rows.append([ch])

        lines: list[PdfTextLine] = []
        for row in rows:
            row.sort(key=lambda c: c["x0"])
            text = "".join(c["text"] for c in row)
            if not text.strip():
                continue
            bbox = PdfRect(
                x0=min(float(c["x0"]) for c in row),
                y0=min(float(c["top"]) for c in row),
                x1=max(float(c["x1"]) for c in row),
                y1=max(float(c["bottom"]) for c in row),
            )
            lines.append(PdfTextLine(text=text, bbox=bbox))
        return lines

    def extract_text_blocks(self, page_index: int,
                            y_gap: float = 6.0) -> list[PdfTextBlock]:
        """
        文本块提取：将相邻视觉行聚合为块。

        以视觉行为基础：行间距不超过 y_gap 的相邻行合并为一个块。
        块文本为行文本以换行符拼接，bbox 为行 bbox 并集。
        """
        lines = self.extract_text_lines(page_index)
        if not lines:
            return []

        blocks: list[list[PdfTextLine]] = [[lines[0]]]
        for ln in lines[1:]:
            prev = blocks[-1][-1]
            if ln.bbox.y0 - prev.bbox.y1 <= y_gap:
                blocks[-1].append(ln)
            else:
                blocks.append([ln])

        result: list[PdfTextBlock] = []
        for blk in blocks:
            bbox = PdfRect(
                x0=min(l.bbox.x0 for l in blk),
                y0=min(l.bbox.y0 for l in blk),
                x1=max(l.bbox.x1 for l in blk),
                y1=max(l.bbox.y1 for l in blk),
            )
            result.append(PdfTextBlock(
                text="\n".join(l.text for l in blk), bbox=bbox))
        return result

    def search_text(self, page_index: int, needle: str) -> list[PdfRect]:
        """
        页面文本搜索。

        在视觉行文本中做子串匹配（忽略空白差异），返回每处命中的近似
        bbox（按命中字符区间的字符 bbox 并集计算）。跨行文本不匹配。
        """
        if not needle:
            return []
        page = self._pdf.pages[page_index]
        chars = page.chars
        if not chars:
            return []

        chars = sorted(chars, key=lambda c: (c["top"], c["x0"]))
        rows: list[list[dict]] = []
        for ch in chars:
            if rows and abs(rows[-1][0]["top"] - ch["top"]) <= 3.0:
                rows[-1].append(ch)
            else:
                rows.append([ch])

        target = "".join(needle.split())
        hits: list[PdfRect] = []
        for row in rows:
            row.sort(key=lambda c: c["x0"])
            # 剔除空白字符后做子串匹配，保留字符索引以还原 bbox
            visible = [c for c in row if c["text"].strip()]
            row_text = "".join(c["text"] for c in visible)
            start = 0
            while True:
                idx = row_text.find(target, start)
                if idx < 0:
                    break
                seg = visible[idx:idx + len(target)]
                hits.append(PdfRect(
                    x0=min(float(c["x0"]) for c in seg),
                    y0=min(float(c["top"]) for c in seg),
                    x1=max(float(c["x1"]) for c in seg),
                    y1=max(float(c["bottom"]) for c in seg),
                ))
                start = idx + 1
        return hits

    def extract_span_lines(self, page_index: int,
                           y_tol: float = 2.0) -> list[PdfSpanLine]:
        """
        按内容流顺序切分的文本行（span 粒度，带字体/字号/颜色属性）。

        沿字符在内容流中的出现顺序扫描，遇到以下情况开新行：
        - 字体（去子集前缀）、字号或颜色变化；
        - 基线（bottom）变化超过 y_tol；
        - x 坐标回退（重复绘制文本副本的开始）；
        - 与前一字符的水平间隙超过约 1.6 个字宽（另一列/另一段）。

        同一位置重复绘制的文本（伪粗体/阴影的常用实现）因 x 回退被切分
        为多条几乎重叠的独立行，由调用方去重或统计重复次数。
        """
        page = self._pdf.pages[page_index]
        chars = page.chars
        if not chars:
            return []

        runs: list[list[dict]] = []
        prev = None
        prev_key = None
        for c in chars:
            key = (normalize_font_name(c.get("fontname", "")),
                   round(float(c.get("size", 0)), 2),
                   str(c.get("non_stroking_color")))
            new_run = prev is None or key != prev_key
            if not new_run:
                if abs(float(prev["bottom"]) - float(c["bottom"])) > y_tol:
                    new_run = True
                else:
                    px1 = float(prev["x1"])
                    cx0 = float(c["x0"])
                    char_w = max(px1 - float(prev["x0"]), 1.5)
                    # x 回退（重绘副本开始）或间隙过大（跨列）
                    if cx0 < px1 - 0.5 * char_w or cx0 - px1 > char_w * 1.6:
                        new_run = True
            if new_run:
                runs.append([c])
            else:
                runs[-1].append(c)
            prev = c
            prev_key = key

        lines: list[PdfSpanLine] = []
        for run in runs:
            text = "".join(c["text"] for c in run)
            if not text.strip():
                continue
            lines.append(PdfSpanLine(
                text=text,
                bbox=PdfRect(
                    x0=min(float(c["x0"]) for c in run),
                    y0=min(float(c["top"]) for c in run),
                    x1=max(float(c["x1"]) for c in run),
                    y1=max(float(c["bottom"]) for c in run),
                ),
                size=float(run[0].get("size", 0)),
                font=normalize_font_name(run[0].get("fontname", "")),
                color=color_to_int(run[0].get("non_stroking_color")),
            ))
        lines.sort(key=lambda l: (round(l.bbox.y0, 1), l.bbox.x0))
        return lines

    # ---- 矢量对象 ----

    def extract_drawings(self, page_index: int) -> list[PdfDrawing]:
        """
        页面矢量绘图对象（矩形、曲线、线段），带填充色。

        圆角矩形通常以曲线路径表示，因此曲线对象与矩形一并返回，
        调用方按外接框和填充色判断即可。
        """
        page = self._pdf.pages[page_index]
        out: list[PdfDrawing] = []
        for kind, objs in (("rect", page.rects), ("curve", page.curves),
                           ("line", page.lines)):
            for o in objs:
                fill = None
                if o.get("fill"):
                    fill = color_to_float_rgb(o.get("non_stroking_color"))
                stroke_color = None
                if o.get("stroke"):
                    stroke_color = color_to_float_rgb(o.get("stroking_color"))
                out.append(PdfDrawing(
                    rect=PdfRect(
                        x0=float(o["x0"]), y0=float(o["top"]),
                        x1=float(o["x1"]), y1=float(o["bottom"]),
                    ),
                    fill=fill,
                    kind=kind,
                    stroke_color=stroke_color,
                    line_width=float(o.get("linewidth") or 0.0),
                ))
        return out

    def extract_raw_spans(self, page_index: int,
                          y_tol: float = 2.0,
                          gap_chars: float = 1.6) -> list[PdfRawSpan]:
        """
        平铺的 raw span（含字符级 bbox），不做行聚合。

        字符按内容流顺序切分为 span：字体/字号/颜色变化、基线（bottom）
        变化超过 y_tol、x 回退、或水平间隙超过 gap_chars 倍字宽即开新
        span。gap_chars 控制同 span 内允许的最大字间隙（以前一字符宽度
        为基准），用于贴合不同渲染器的 span 切分粒度。
        """
        page = self._pdf.pages[page_index]
        chars = page.chars
        if not chars:
            return []

        runs: list[list[dict]] = []
        prev = None
        prev_key = None
        for c in chars:
            key = (normalize_font_name(c.get("fontname", "")),
                   round(float(c.get("size", 0)), 2),
                   str(c.get("non_stroking_color")))
            new_run = prev is None or key != prev_key
            if not new_run:
                if abs(float(prev["bottom"]) - float(c["bottom"])) > y_tol:
                    new_run = True
                else:
                    px1 = float(prev["x1"])
                    cx0 = float(c["x0"])
                    char_w = max(px1 - float(prev["x0"]), 1.5)
                    if cx0 < px1 - 0.5 * char_w or cx0 - px1 > char_w * gap_chars:
                        new_run = True
            if new_run:
                runs.append([c])
            else:
                runs[-1].append(c)
            prev = c
            prev_key = key

        spans: list[PdfRawSpan] = []
        for run in runs:
            text = "".join(c["text"] for c in run)
            if not text.strip():
                continue
            raw_chars = tuple(
                PdfRawChar(
                    c=c["text"],
                    bbox=PdfRect(float(c["x0"]), float(c["top"]),
                                 float(c["x1"]), float(c["bottom"])),
                ) for c in run
            )
            spans.append(PdfRawSpan(
                text=text,
                bbox=PdfRect(
                    x0=min(float(c["x0"]) for c in run),
                    y0=min(float(c["top"]) for c in run),
                    x1=max(float(c["x1"]) for c in run),
                    y1=max(float(c["bottom"]) for c in run),
                ),
                font=str(run[0].get("fontname", "")),
                size=float(run[0].get("size", 0)),
                color=color_to_int(run[0].get("non_stroking_color")),
                chars=raw_chars,
            ))
        return spans

    @staticmethod
    def _path_items(obj: dict, kind: str) -> tuple:
        """把 pdfplumber 对象转为 ("re", rect) / ("l", p0, p1) items。"""
        x0, top = float(obj["x0"]), float(obj["top"])
        x1, bottom = float(obj["x1"]), float(obj["bottom"])
        if kind == "rect":
            return (("re", PdfRect(x0, top, x1, bottom)),)
        pts = [(float(p[0]), float(p[1])) for p in (obj.get("pts") or [])]
        if kind == "line":
            if len(pts) >= 2:
                return (("l", pts[0], pts[1]),)
            return (("l", (x0, top), (x1, bottom)),)
        # curve：先判是否闭合轴对齐矩形（办公软件矩形常导出为闭合路径）
        dedup: list[tuple[float, float]] = []
        for p in pts:
            if not dedup or abs(p[0] - dedup[-1][0]) > 0.01 or abs(p[1] - dedup[-1][1]) > 0.01:
                dedup.append(p)
        if len(dedup) >= 2 and abs(dedup[0][0] - dedup[-1][0]) <= 0.01                 and abs(dedup[0][1] - dedup[-1][1]) <= 0.01:
            dedup = dedup[:-1]
        if len(dedup) == 4:
            closed = dedup + [dedup[0]]
            axis_aligned = all(
                abs(a[0] - b[0]) <= 0.01 or abs(a[1] - b[1]) <= 0.01
                for a, b in zip(closed, closed[1:])
            )
            if axis_aligned:
                return (("re", PdfRect(x0, top, x1, bottom)),)
        items = []
        for a, b in zip(pts, pts[1:]):
            if abs(a[0] - b[0]) <= 0.01 and abs(a[1] - b[1]) <= 0.01:
                continue
            items.append(("l", a, b))
        return tuple(items)

    def extract_paths(self, page_index: int) -> list[PdfPath]:
        """
        页面路径级矢量对象（矩形/线段/曲线），带 items 结构。

        与 extract_drawings 的区别：保留每条路径的构成（矩形子路径、
        线段序列），供需要判断形状结构（斜条、箭头尖角等）的调用方使用。
        曲线的贝塞尔段以锚点间线段近似。
        """
        page = self._pdf.pages[page_index]
        out: list[PdfPath] = []
        for kind, objs in (("rect", page.rects), ("curve", page.curves),
                           ("line", page.lines)):
            for o in objs:
                fill = None
                if o.get("fill"):
                    fill = color_to_float_rgb(o.get("non_stroking_color"))
                stroke = None
                if o.get("stroke"):
                    stroke = color_to_float_rgb(o.get("stroking_color"))
                out.append(PdfPath(
                    rect=PdfRect(
                        x0=float(o["x0"]), y0=float(o["top"]),
                        x1=float(o["x1"]), y1=float(o["bottom"]),
                    ),
                    fill=fill,
                    stroke=stroke,
                    line_width=float(o.get("linewidth") or 0.0),
                    items=self._path_items(o, kind),
                ))
        return out

    def extract_raw_lines(self, page_index: int,
                          y_tol: float = 2.0,
                          line_gap: float = 12.0) -> list[PdfRawLine]:
        """
        raw 文本行：span 序列 + 字符级 bbox。

        1. 按内容流顺序把同 (字体, 字号, 颜色) 且同基线、x 相邻的字符
           串成 span（x 回退即换 span，保持重复绘制文本独立）；
        2. 同基线（bottom 差 <= y_tol）且水平间隙 <= line_gap 的相邻
           span 归入同一行；间隙更大视为另一列（如表格相邻单元格）。
        """
        page = self._pdf.pages[page_index]
        chars = page.chars
        if not chars:
            return []

        # ---- 1. 字符 -> span（内容流顺序切分）----
        runs: list[list[dict]] = []
        prev = None
        prev_key = None
        for c in chars:
            key = (normalize_font_name(c.get("fontname", "")),
                   round(float(c.get("size", 0)), 2),
                   str(c.get("non_stroking_color")))
            new_run = prev is None or key != prev_key
            if not new_run:
                if abs(float(prev["bottom"]) - float(c["bottom"])) > y_tol:
                    new_run = True
                else:
                    px1 = float(prev["x1"])
                    cx0 = float(c["x0"])
                    char_w = max(px1 - float(prev["x0"]), 1.5)
                    if cx0 < px1 - 0.5 * char_w or cx0 - px1 > char_w * 1.6:
                        new_run = True
            if new_run:
                runs.append([c])
            else:
                runs[-1].append(c)
            prev = c
            prev_key = key

        spans: list[PdfRawSpan] = []
        for run in runs:
            text = "".join(c["text"] for c in run)
            if not text.strip():
                continue
            raw_chars = tuple(
                PdfRawChar(
                    c=c["text"],
                    bbox=PdfRect(float(c["x0"]), float(c["top"]),
                                 float(c["x1"]), float(c["bottom"])),
                ) for c in run
            )
            spans.append(PdfRawSpan(
                text=text,
                bbox=PdfRect(
                    x0=min(float(c["x0"]) for c in run),
                    y0=min(float(c["top"]) for c in run),
                    x1=max(float(c["x1"]) for c in run),
                    y1=max(float(c["bottom"]) for c in run),
                ),
                font=normalize_font_name(run[0].get("fontname", "")),
                size=float(run[0].get("size", 0)),
                color=color_to_int(run[0].get("non_stroking_color")),
                chars=raw_chars,
            ))

        # ---- 2. span -> 行（同基线 + 间隙上限）----
        spans.sort(key=lambda s: (round(s.bbox.y1, 1), s.bbox.x0))
        lines: list[list[PdfRawSpan]] = []
        for sp in spans:
            merged = False
            if lines:
                last = lines[-1]
                if abs(last[0].bbox.y1 - sp.bbox.y1) <= y_tol:
                    gap = sp.bbox.x0 - max(s.bbox.x1 for s in last)
                    if -1.0 <= gap <= line_gap:
                        last.append(sp)
                        merged = True
            if not merged:
                lines.append([sp])

        out: list[PdfRawLine] = []
        for group in lines:
            group.sort(key=lambda s: s.bbox.x0)
            out.append(PdfRawLine(
                text="".join(s.text for s in group),
                bbox=PdfRect(
                    x0=min(s.bbox.x0 for s in group),
                    y0=min(s.bbox.y0 for s in group),
                    x1=max(s.bbox.x1 for s in group),
                    y1=max(s.bbox.y1 for s in group),
                ),
                spans=tuple(group),
            ))
        out.sort(key=lambda l: (l.bbox.y0, l.bbox.x0))
        return out


def open_pdf(path: str) -> PdfDocument:
    """打开 PDF，返回文档句柄。失败抛出异常，由调用方处理。"""
    return PdfDocument(path)
