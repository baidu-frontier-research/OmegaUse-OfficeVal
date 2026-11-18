#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
PyMuPDF（fitz / pymupdf）兼容层 —— 只用宽松许可证库实现其所需子集。

本项目已移除 PyMuPDF（AGPL 与本项目许可证不兼容），运行环境不安装 fitz。
少数验证脚本按 PyMuPDF 的 API 写成，逐条改写评分逻辑风险大，因此在此提供
等价实现：脚本只需把 import 换成本模块（import pymupdf_compat as pymupdf /
as fitz），其余代码一行不动。

底层实现：
  - pdfplumber (MIT)：文本 / 字符 / 坐标 / 字体 / 颜色 / 矢量对象 / 表格 / 注释；
  - pypdfium2 (Apache-2.0 / BSD-3-Clause)：页面渲染、嵌入图像字节提取；
  - 文本行 / span / 文本块的切分一律复用同目录 pdf_backend 适配层的现成逻辑。

坐标约定与 PyMuPDF 一致：原点在页面左上角，y 向下增大，单位 pt。

已实现子集
  模块级：open / Document / Page / Pixmap / Rect / IRect / Point / Matrix /
          Colorspace / csGRAY / csRGB / csCMYK / PDF_PERM_* / PDF_ANNOT_*
  Document：page_count / len() / 下标 / 迭代 / metadata / permissions /
            needs_pass / is_encrypted / extract_image / extract_font /
            close / with 语句
  Page：rect / rotation / number / get_text("text"|"dict"|"rawdict"|"blocks") /
        get_drawings / get_images / get_image_info / get_image_rects /
        get_pixmap / find_tables / annots / get_fonts

无法由 pdfplumber / pypdfium2 提供的能力（一律优雅降级，绝不抛异常）
  - Document.extract_font()：拿不到内嵌字体二进制，恒返回空 buffer；
  - Page.get_fonts()：拿不到页面字体资源表，恒返回空列表；
  - get_pixmap 的非等比缩放：pypdfium2 只支持等比，x/y 缩放不同时取均值；
  - Pixmap.tobytes()：无 PIL 时返回 b""。
"""

from __future__ import annotations

import io
import math
import os
import sys

# verifier 由 core/worker.py 以文件路径动态加载（模块名 = 文件名，不作为包），
# 也可能作为 verifiers 包的一部分被导入。把本文件所在目录放进 sys.path，
# 两种加载方式下 "import pdf_backend" 都必定可用。
_HERE = os.path.dirname(os.path.abspath(__file__))
if _HERE not in sys.path:
    sys.path.insert(0, _HERE)

import pdf_backend  # noqa: E402 —— 需先补 sys.path

try:
    import numpy as _np
except ImportError:                       # 渲染缓冲区退化为逐行手工拷贝
    _np = None


# --------------------------------------------------------------------------
# 常量（与 PyMuPDF 同名同值）
# --------------------------------------------------------------------------

# /Encrypt 的 /P 权限位（PDF 32000-1 表 22），PyMuPDF 沿用同一组数值
PDF_PERM_PRINT = 4
PDF_PERM_MODIFY = 8
PDF_PERM_COPY = 16
PDF_PERM_ANNOTATE = 32
PDF_PERM_FORM = 256
PDF_PERM_ACCESSIBILITY = 512
PDF_PERM_ASSEMBLE = 1024
PDF_PERM_PRINT_HQ = 2048
PDF_PERM_ALL = 0xFFFFFFFF

# span flags 位（1 上标 / 2 斜体 / 4 衬线 / 8 等宽 / 16 粗体）
TEXT_FONT_SUPERSCRIPT = 1
TEXT_FONT_ITALIC = 2
TEXT_FONT_SERIFED = 4
TEXT_FONT_MONOSPACED = 8
TEXT_FONT_BOLD = 16

CS_GRAY = 0
CS_RGB = 1
CS_CMYK = 2

# 注释类型编号（PyMuPDF PDF_ANNOT_*）
_ANNOT_TYPES = {
    "Text": 0, "Link": 1, "FreeText": 2, "Line": 3, "Square": 4,
    "Circle": 5, "Polygon": 6, "PolyLine": 7, "Highlight": 8,
    "Underline": 9, "Squiggly": 10, "StrikeOut": 11, "Redact": 12,
    "Stamp": 13, "Caret": 14, "Ink": 15, "Popup": 16, "FileAttachment": 17,
    "Sound": 18, "Movie": 19, "Widget": 20, "Screen": 21,
    "PrinterMark": 22, "TrapNet": 23, "Watermark": 24, "3D": 25,
    "Projection": 26, "RichMedia": 27,
}
PDF_ANNOT_TEXT = 0
PDF_ANNOT_LINK = 1
PDF_ANNOT_HIGHLIGHT = 8
PDF_ANNOT_POPUP = 16
PDF_ANNOT_WIDGET = 20
PDF_ANNOT_UNKNOWN = -1

# get_text("dict") 的文本块聚合阈值：与 pdf_backend.extract_text_blocks 一致
_BLOCK_Y_GAP = 6.0

# --------------------------------------------------------------------------
# 几何：Point / Rect / IRect / Matrix
# --------------------------------------------------------------------------

class Point:
    """PyMuPDF Point 的子集：属性 x/y，可迭代、可索引。"""

    __slots__ = ("x", "y")

    def __init__(self, *args):
        if not args:
            x = y = 0.0
        elif len(args) == 2:
            x, y = args
        elif len(args) == 1:
            a = args[0]
            if isinstance(a, Point):
                x, y = a.x, a.y
            else:
                seq = tuple(a)
                if len(seq) < 2:
                    raise ValueError("Point 需要 2 个坐标")
                x, y = seq[0], seq[1]
        else:
            raise ValueError("Point 参数个数非法：%d" % len(args))
        self.x, self.y = float(x), float(y)

    def __iter__(self):
        return iter((self.x, self.y))

    def __len__(self):
        return 2

    def __getitem__(self, i):
        return (self.x, self.y)[i]

    def __eq__(self, other):
        try:
            o = Point(other)
        except (TypeError, ValueError):
            return NotImplemented
        return self.x == o.x and self.y == o.y

    def __hash__(self):
        return hash((self.x, self.y))

    def __repr__(self):
        return "Point(%g, %g)" % (self.x, self.y)


def _rect_args(args):
    """把 Rect(...) 的各种入参形式统一成 (x0, y0, x1, y1)。"""
    if not args:
        return (0.0, 0.0, 0.0, 0.0)
    if len(args) == 4:
        return tuple(float(v) for v in args)
    if len(args) == 2:                                  # Rect(点, 点)
        p0, p1 = Point(args[0]), Point(args[1])
        return (p0.x, p0.y, p1.x, p1.y)
    if len(args) == 1:
        a = args[0]
        if hasattr(a, "x0") and hasattr(a, "y1"):       # Rect / PdfRect
            return (float(a.x0), float(a.y0), float(a.x1), float(a.y1))
        seq = tuple(a)
        if len(seq) == 4:
            return tuple(float(v) for v in seq)
        if len(seq) == 2:
            p0, p1 = Point(seq[0]), Point(seq[1])
            return (p0.x, p0.y, p1.x, p1.y)
    raise ValueError("Rect 参数非法：%r" % (args,))


class Rect:
    """PyMuPDF Rect 的子集：pt 单位、左上原点，支持交并、包含与面积。"""

    __slots__ = ("x0", "y0", "x1", "y1")

    _AREA_UNIT = {"pt": 1.0, "in": 1.0 / 72.0, "mm": 25.4 / 72.0,
                  "cm": 2.54 / 72.0}

    def __init__(self, *args):
        self.x0, self.y0, self.x1, self.y1 = _rect_args(args)

    @property
    def width(self):
        return abs(self.x1 - self.x0)

    @property
    def height(self):
        return abs(self.y1 - self.y0)

    @property
    def top_left(self):
        return Point(self.x0, self.y0)

    @property
    def bottom_right(self):
        return Point(self.x1, self.y1)

    @property
    def is_empty(self):
        return self.x0 >= self.x1 or self.y0 >= self.y1

    @property
    def is_valid(self):
        return self.x0 <= self.x1 and self.y0 <= self.y1

    @property
    def irect(self):
        return IRect(math.floor(self.x0), math.floor(self.y0),
                     math.ceil(self.x1), math.ceil(self.y1))

    def round(self):
        return self.irect

    def normalize(self):
        """就地规范化（x0<=x1、y0<=y1），与 PyMuPDF 一致返回自身。"""
        if self.x0 > self.x1:
            self.x0, self.x1 = self.x1, self.x0
        if self.y0 > self.y1:
            self.y0, self.y1 = self.y1, self.y0
        return self

    def _normalized(self):
        return Rect(min(self.x0, self.x1), min(self.y0, self.y1),
                    max(self.x0, self.x1), max(self.y0, self.y1))

    def get_area(self, unit="pt"):
        f = self._AREA_UNIT.get(unit, 1.0)
        a = max(0.0, self.x1 - self.x0) * max(0.0, self.y1 - self.y0)
        return a * f * f
    def _has_point(self, x, y):
        return self.x0 <= x <= self.x1 and self.y0 <= y <= self.y1

    def _has_rect(self, r):
        return (self.x0 <= r.x0 and r.x1 <= self.x1
                and self.y0 <= r.y0 and r.y1 <= self.y1)

    def contains(self, other):
        """点（Point / 2 元组）或矩形（Rect / 4 元组）是否被本矩形包含。"""
        if self.is_empty:
            return False
        if isinstance(other, Point):
            return self._has_point(other.x, other.y)
        if isinstance(other, Rect):
            return self._has_rect(other._normalized())
        if hasattr(other, "x0") and hasattr(other, "y1"):
            return self._has_rect(Rect(other)._normalized())
        try:
            seq = tuple(float(v) for v in other)
        except (TypeError, ValueError):
            return False
        if len(seq) == 2:
            return self._has_point(seq[0], seq[1])
        if len(seq) == 4:
            return self._has_rect(Rect(seq)._normalized())
        return False

    def __contains__(self, other):
        return self.contains(other)

    def intersects(self, other):
        return not (self & other).is_empty

    def __and__(self, other):
        """交集；无重叠时返回空矩形（get_area() == 0），不抛异常。"""
        a = self._normalized()
        b = Rect(other)._normalized()
        x0, y0 = max(a.x0, b.x0), max(a.y0, b.y0)
        x1, y1 = min(a.x1, b.x1), min(a.y1, b.y1)
        if x1 <= x0 or y1 <= y0:
            return Rect(0.0, 0.0, 0.0, 0.0)
        return Rect(x0, y0, x1, y1)

    def __or__(self, other):
        """并集（点亦可并入）。"""
        a = self._normalized()
        try:
            b = Rect(other)._normalized()
        except ValueError:
            p = Point(other)
            b = Rect(p.x, p.y, p.x, p.y)
        return Rect(min(a.x0, b.x0), min(a.y0, b.y0),
                    max(a.x1, b.x1), max(a.y1, b.y1))

    def __iter__(self):
        return iter((self.x0, self.y0, self.x1, self.y1))

    def __len__(self):
        return 4

    def __getitem__(self, i):
        return (self.x0, self.y0, self.x1, self.y1)[i]

    def __eq__(self, other):
        try:
            o = Rect(other)
        except (TypeError, ValueError):
            return NotImplemented
        return tuple(self) == tuple(o)

    def __hash__(self):
        return hash(tuple(self))

    def __repr__(self):
        return "Rect(%g, %g, %g, %g)" % (self.x0, self.y0, self.x1, self.y1)


class IRect(Rect):
    """整数矩形（PyMuPDF IRect 的最小子集）。"""

    def __init__(self, *args):
        Rect.__init__(self, *args)
        self.x0, self.y0 = int(round(self.x0)), int(round(self.y0))
        self.x1, self.y1 = int(round(self.x1)), int(round(self.y1))

    def __repr__(self):
        return "IRect(%d, %d, %d, %d)" % (self.x0, self.y0, self.x1, self.y1)


class Matrix:
    """PyMuPDF Matrix 的子集。本兼容层只用它的缩放分量 a（x）与 d（y）。"""

    __slots__ = ("a", "b", "c", "d", "e", "f")

    def __init__(self, *args):
        if not args:
            vals = (1.0, 0.0, 0.0, 1.0, 0.0, 0.0)
        elif len(args) == 6:
            vals = tuple(float(v) for v in args)
        elif len(args) == 2:                        # Matrix(sx, sy)：缩放
            vals = (float(args[0]), 0.0, 0.0, float(args[1]), 0.0, 0.0)
        elif len(args) == 1:
            a = args[0]
            if isinstance(a, Matrix):
                vals = tuple(a)
            elif isinstance(a, (int, float)):       # Matrix(度数)：旋转
                r = math.radians(float(a))
                vals = (math.cos(r), math.sin(r), -math.sin(r),
                        math.cos(r), 0.0, 0.0)
            else:
                vals = tuple(float(v) for v in a)
                if len(vals) != 6:
                    raise ValueError("Matrix 需要 6 个分量")
        else:
            raise ValueError("Matrix 参数个数非法：%d" % len(args))
        self.a, self.b, self.c, self.d, self.e, self.f = vals

    def __iter__(self):
        return iter((self.a, self.b, self.c, self.d, self.e, self.f))

    def __len__(self):
        return 6

    def __getitem__(self, i):
        return (self.a, self.b, self.c, self.d, self.e, self.f)[i]

    def __repr__(self):
        return "Matrix(%g, %g, %g, %g, %g, %g)" % tuple(self)


class Colorspace:
    """色彩空间标记（PyMuPDF Colorspace 的子集），只需要 n 与名字。"""

    __slots__ = ("type", "n", "name")

    _INFO = {CS_GRAY: (1, "DeviceGray"), CS_RGB: (3, "DeviceRGB"),
             CS_CMYK: (4, "DeviceCMYK")}

    def __init__(self, type_=CS_RGB):
        n, name = self._INFO.get(int(type_), (3, "DeviceRGB"))
        self.type, self.n, self.name = int(type_), n, name

    def __repr__(self):
        return "Colorspace(%s)" % self.name


csGRAY = Colorspace(CS_GRAY)
csRGB = Colorspace(CS_RGB)
csCMYK = Colorspace(CS_CMYK)

# --------------------------------------------------------------------------
# Pixmap
# --------------------------------------------------------------------------

class Pixmap:
    """
    渲染结果位图（PyMuPDF Pixmap 的子集）。

    samples 为无行填充的紧凑字节串，因此灰度用法
    np.frombuffer(pix.samples, np.uint8).reshape(pix.height, pix.width)
    必然成立。
    """

    __slots__ = ("samples", "width", "height", "n", "alpha", "colorspace",
                 "stride", "x", "y", "xres", "yres")

    def __init__(self, samples, width, height, n, alpha=False,
                 colorspace=None, xres=72, yres=72):
        self.samples = bytes(samples)
        self.width, self.height, self.n = int(width), int(height), int(n)
        self.alpha = bool(alpha)
        self.colorspace = colorspace
        self.stride = self.width * self.n
        self.x = self.y = 0
        self.xres, self.yres = int(xres), int(yres)

    @property
    def w(self):
        return self.width

    @property
    def h(self):
        return self.height

    @property
    def irect(self):
        return IRect(0, 0, self.width, self.height)

    @property
    def samples_mv(self):
        return memoryview(self.samples)

    def pixel(self, x, y):
        """(x, y) 处像素，返回长度为 n 的整数元组；越界抛 IndexError。"""
        x, y = int(x), int(y)
        if not (0 <= x < self.width and 0 <= y < self.height):
            raise IndexError("像素坐标越界：(%d, %d)" % (x, y))
        off = (y * self.width + x) * self.n
        return tuple(self.samples[off:off + self.n])

    def tobytes(self, output="png", jpg_quality=95):
        """编码为图片字节串；无 PIL 时降级返回 b""（不抛异常）。"""
        try:
            from PIL import Image
        except ImportError:
            return b""
        try:
            mode = {1: "L", 3: "RGB",
                    4: "RGBA" if self.alpha else "CMYK"}[self.n]
            img = Image.frombytes(mode, (self.width, self.height), self.samples)
            buf = io.BytesIO()
            fmt = (output or "png").upper()
            fmt = "JPEG" if fmt in ("JPG", "JPEG") else fmt
            if fmt == "JPEG":
                img.save(buf, fmt, quality=int(jpg_quality))
            else:
                img.save(buf, fmt)
            return buf.getvalue()
        except Exception:
            return b""

    def __repr__(self):
        return "Pixmap(%dx%d, n=%d)" % (self.width, self.height, self.n)


def _bitmap_samples(bitmap, want_gray, keep_alpha):
    """pypdfium2 位图 -> (samples, width, height, n)；samples 无行填充。"""
    width, height = int(bitmap.width), int(bitmap.height)
    if _np is not None:
        arr = bitmap.to_numpy()
        n = 1 if arr.ndim == 2 else int(arr.shape[2])
        if want_gray and n != 1:
            gray = (arr[:, :, 0] * 0.299 + arr[:, :, 1] * 0.587
                    + arr[:, :, 2] * 0.114)
            arr, n = gray.round().clip(0, 255).astype("uint8"), 1
        elif n == 4 and not keep_alpha:
            arr, n = arr[:, :, :3], 3
        return _np.ascontiguousarray(arr).tobytes(), width, height, n

    # 无 numpy 的兜底：按 stride 逐行裁掉行填充
    raw = bytes(bitmap.buffer)
    stride = int(getattr(bitmap, "stride", width))
    n = int(getattr(bitmap, "n_channels", max(1, stride // max(1, width))))
    data = b"".join(raw[y * stride:y * stride + width * n]
                    for y in range(height))
    if want_gray and n != 1:
        data = bytes(int(round(data[i] * 0.299 + data[i + 1] * 0.587
                               + data[i + 2] * 0.114))
                     for i in range(0, len(data), n))
        n = 1
    elif n == 4 and not keep_alpha:
        data = bytes(b for i in range(0, len(data), 4) for b in data[i:i + 3])
        n = 3
    return data, width, height, n


# --------------------------------------------------------------------------
# 小工具
# --------------------------------------------------------------------------

_BOLD_HINTS = ("bold", "heavy", "black", "-bd", ",bd")
_ITALIC_HINTS = ("italic", "oblique", "-it", ",it")
_SERIF_HINTS = ("times", "serif", "songti", "simsun", "mincho", "georgia",
                "garamond", "cambria", "宋", "仿宋", "楷")
_MONO_HINTS = ("mono", "courier", "consol")


def _font_flags(font_name):
    """由字体名推断 span flags（拿不到字体描述符，只能靠名字，与 PyMuPDF 近似）。"""
    name = font_name or ""
    low = name.lower()
    flags = 0
    if any(h in low for h in _MONO_HINTS):
        flags |= TEXT_FONT_MONOSPACED
    if any(h in low for h in _SERIF_HINTS):
        flags |= TEXT_FONT_SERIFED
    if any(h in low for h in _ITALIC_HINTS):
        flags |= TEXT_FONT_ITALIC
    if any(h in low for h in _BOLD_HINTS) or "黑" in name:
        flags |= TEXT_FONT_BOLD
    return flags


def _dash_str(dash):
    """pdfplumber dash 字段 -> PyMuPDF dashes 字符串（"[] 0" 即实线）。"""
    if not dash:
        return "[] 0"
    try:
        arr, phase = dash
        vals = list(arr or [])
        if not vals:
            return "[] 0"
        return "[%s] %g" % (" ".join("%g" % float(v) for v in vals),
                            float(phase or 0))
    except Exception:
        return "[] 0"


def _rect_of(pdf_rect):
    """pdf_backend.PdfRect -> 本模块 Rect。"""
    return Rect(pdf_rect.x0, pdf_rect.y0, pdf_rect.x1, pdf_rect.y1)


def _bbox_of(pdf_rect):
    return (float(pdf_rect.x0), float(pdf_rect.y0),
            float(pdf_rect.x1), float(pdf_rect.y1))


def _as_text(value):
    """元数据 / 注释文本统一成 str（pdfplumber 一般已解码，此处只做兜底）。"""
    if value is None:
        return ""
    if isinstance(value, bytes):
        for enc in ("utf-16", "utf-8", "latin-1"):
            try:
                return value.decode(enc).strip("\x00")
            except (UnicodeDecodeError, LookupError):
                continue
        return ""
    return value if isinstance(value, str) else str(value)

# --------------------------------------------------------------------------
# 注释
# --------------------------------------------------------------------------

def _annot_subtype(raw):
    sub = (raw.get("data") or {}).get("Subtype")
    name = getattr(sub, "name", None)
    if name:
        return str(name)
    if isinstance(sub, bytes):
        return sub.decode("latin-1", "ignore")
    if isinstance(sub, str):
        return sub.lstrip("/")
    return "Unknown"


class Annot:
    """注释对象（PyMuPDF Annot 的子集）：rect / type / info。"""

    __slots__ = ("rect", "type", "info", "_raw")

    def __init__(self, raw):
        raw = raw or {}
        self._raw = raw
        self.rect = Rect(float(raw.get("x0") or 0.0),
                         float(raw.get("top") or 0.0),
                         float(raw.get("x1") or 0.0),
                         float(raw.get("bottom") or 0.0))
        subtype = _annot_subtype(raw)
        self.type = (_ANNOT_TYPES.get(subtype, PDF_ANNOT_UNKNOWN), subtype)
        self.info = {"title": _as_text(raw.get("title")),
                     "content": _as_text(raw.get("contents")),
                     "subject": "", "name": ""}

    def __repr__(self):
        return "Annot(%s, %r)" % (self.type[1], self.rect)


class _AnnotList(list):
    """既能被迭代（page.annots）又能被调用（page.annots()），两种写法都可用。"""

    def __call__(self, types=None):
        if types is None:
            return _AnnotList(self)
        want = set(types)
        return _AnnotList(a for a in self if a.type[0] in want)


# --------------------------------------------------------------------------
# 表格：pdfplumber find_tables 的 PyMuPDF 外观
# --------------------------------------------------------------------------

class _TableRow:
    """表格行：cells 为 4 元组（跨行跨列的空位为 None），bbox 为整行外框。"""

    __slots__ = ("cells", "bbox")

    def __init__(self, row):
        self.cells = list(getattr(row, "cells", None) or [])
        bbox = getattr(row, "bbox", None)
        self.bbox = tuple(float(v) for v in bbox) if bbox else (0.0,) * 4


class Table:
    """PyMuPDF Table 的子集：bbox / row_count / col_count / rows / cells /
    extract()，全部由 pdfplumber 的 Table 转出。"""

    def __init__(self, table):
        self._t = table
        bbox = getattr(table, "bbox", None) or (0.0,) * 4
        self.bbox = tuple(float(v) for v in bbox)
        self.rows = [_TableRow(r) for r in (getattr(table, "rows", None) or [])]
        self.row_count = len(self.rows)
        cols = getattr(table, "columns", None)
        self.col_count = (len(cols) if cols else
                          max((len(r.cells) for r in self.rows), default=0))
        self.cells = list(getattr(table, "cells", None) or [])
        self.header = None

    def extract(self, **kwargs):
        """单元格文本矩阵，与 PyMuPDF 同形：list[list[str | None]]。"""
        try:
            return self._t.extract(**kwargs)
        except Exception:
            return []

    def to_markdown(self, **kwargs):
        return "\n".join(
            "|" + "|".join("" if c is None else str(c) for c in row) + "|"
            for row in self.extract())

    def __repr__(self):
        return "Table(%dx%d, bbox=%s)" % (self.row_count, self.col_count,
                                          self.bbox)


class TableFinder:
    """PyMuPDF find_tables() 的返回对象，主要提供 .tables。"""

    def __init__(self, tables):
        self.tables = list(tables)

    def __len__(self):
        return len(self.tables)

    def __iter__(self):
        return iter(self.tables)

    def __getitem__(self, i):
        return self.tables[i]

    def __repr__(self):
        return "TableFinder(%d tables)" % len(self.tables)


# PyMuPDF strategy -> pdfplumber 的 (vertical_strategy, horizontal_strategy)
_STRATEGY_MAP = {
    "lines": ("lines", "lines"),
    "lines_strict": ("lines_strict", "lines_strict"),
    "text": ("text", "text"),
    "explicit": ("explicit", "explicit"),
}

# --------------------------------------------------------------------------
# Page
# --------------------------------------------------------------------------

class Page:
    """页面句柄（PyMuPDF Page 的子集）。"""

    def __init__(self, doc, index):
        self._doc = doc
        self.number = int(index)
        w, h = doc._backend.page_size(self.number)
        self._width, self._height = float(w), float(h)
        self._dict_cache = {}
        self._drawings = None
        self._annots = None
        self._pdfium_page = None

    # ---- 基本属性 ----

    @property
    def parent(self):
        return self._doc

    @property
    def rect(self):
        return Rect(0.0, 0.0, self._width, self._height)

    @property
    def mediabox(self):
        return self.rect

    @property
    def cropbox(self):
        return self.rect

    @property
    def rotation(self):
        try:
            return self._doc._backend.page_rotation(self.number)
        except Exception:
            return 0

    def bound(self):
        return self.rect

    def __repr__(self):
        return "Page(%d, %gx%g)" % (self.number, self._width, self._height)

    # ---- 内部句柄 ----

    def _plumber(self):
        return self._doc._plumber_page(self.number)

    def _pdfium(self):
        if self._pdfium_page is None:
            self._pdfium_page = self._doc._pdfium_doc()[self.number]
        return self._pdfium_page

    # ---- 文本 ----

    def get_text(self, option="text", **kwargs):
        """支持 "text" / "dict" / "rawdict" / "blocks"；其他类型降级为空。"""
        opt = str(option or "text").lower()
        if opt == "text":
            try:
                return self._doc._backend.page_text(self.number)
            except Exception:
                return ""
        if opt in ("dict", "rawdict"):
            return self._text_dict(raw=(opt == "rawdict"))
        if opt == "blocks":
            try:
                blocks = self._doc._backend.extract_text_blocks(self.number)
            except Exception:
                return []
            return [(b.bbox.x0, b.bbox.y0, b.bbox.x1, b.bbox.y1, b.text, i, 0)
                    for i, b in enumerate(blocks)]
        if opt == "words":                     # 形状对齐：PyMuPDF 返回元组列表
            try:
                lines = self._doc._backend.extract_text_lines(self.number)
            except Exception:
                return []
            return [(l.bbox.x0, l.bbox.y0, l.bbox.x1, l.bbox.y1, l.text,
                     i, 0, 0) for i, l in enumerate(lines)]
        return ""

    def _text_dict(self, raw=False):
        """构造 PyMuPDF 的 dict / rawdict 结构。

        文本块由 pdf_backend.extract_raw_lines 的行按 extract_text_blocks
        的行间距规则聚合；图片块（type=1）来自 extract_images 的放置矩形，
        width / height 取源像素尺寸（与 PyMuPDF 一致）。
        """
        key = bool(raw)
        if key in self._dict_cache:
            return self._dict_cache[key]

        be = self._doc._backend
        try:
            lines = be.extract_raw_lines(self.number)
        except Exception:
            lines = []
        try:
            images = be.extract_images(self.number)
        except Exception:
            images = []

        groups = []
        for ln in lines:
            if groups and ln.bbox.y0 - groups[-1][-1].bbox.y1 <= _BLOCK_Y_GAP:
                groups[-1].append(ln)
            else:
                groups.append([ln])

        blocks = []
        for grp in groups:
            blocks.append({
                "type": 0,
                "bbox": (min(l.bbox.x0 for l in grp),
                         min(l.bbox.y0 for l in grp),
                         max(l.bbox.x1 for l in grp),
                         max(l.bbox.y1 for l in grp)),
                "lines": [self._line_dict(l, raw) for l in grp],
            })
        for im in images:
            blocks.append({
                "type": 1,
                "bbox": _bbox_of(im.rect),
                "width": int(im.src_width),
                "height": int(im.src_height),
                "ext": "", "colorspace": 3, "bpc": 8,
                "xres": 96, "yres": 96, "size": 0,
                "image": b"", "digest": b"",
                "transform": (im.rect.width, 0.0, 0.0, im.rect.height,
                              im.rect.x0, im.rect.y0),
            })

        blocks.sort(key=lambda b: (round(b["bbox"][1], 2), b["bbox"][0]))
        for i, b in enumerate(blocks):
            b["number"] = i
        out = {"width": self._width, "height": self._height, "blocks": blocks}
        self._dict_cache[key] = out
        return out

    @staticmethod
    def _line_dict(line, raw):
        return {"bbox": _bbox_of(line.bbox), "wmode": 0, "dir": (1.0, 0.0),
                "spans": [Page._span_dict(sp, raw) for sp in line.spans]}

    @staticmethod
    def _span_dict(span, raw):
        d = {
            "size": float(span.size),
            "flags": _font_flags(span.font),
            "bidi": 0, "char_flags": 0,
            "font": span.font or "",
            "color": int(span.color),
            "alpha": 255, "ascender": 0.9, "descender": -0.2,
            "text": span.text,
            "origin": (float(span.bbox.x0), float(span.bbox.y1)),
            "bbox": _bbox_of(span.bbox),
        }
        if raw:
            d["chars"] = [{"origin": (float(c.bbox.x0), float(c.bbox.y1)),
                           "bbox": _bbox_of(c.bbox), "c": c.c}
                          for c in span.chars]
        return d
    # ---- 矢量对象 ----

    def get_drawings(self, extended=False):
        """矢量绘图对象，字段与 PyMuPDF 对齐：
        rect / items / type / fill / color / width / dashes。

        items 为 ("re", Rect) 或 ("l", Point, Point)；颜色为 0~1 浮点 RGB。
        """
        if self._drawings is not None:
            return self._drawings
        try:
            paths = self._doc._backend.extract_paths(self.number)
        except Exception:
            self._drawings = []
            return self._drawings

        # extract_paths 按 (rects, curves, lines) 顺序平铺，按同序取回 dash
        dashes = []
        try:
            pg = self._plumber()
            for objs in (pg.rects, pg.curves, pg.lines):
                for o in objs:
                    dashes.append(_dash_str(o.get("dash")))
        except Exception:
            dashes = []

        out = []
        for i, p in enumerate(paths):
            items = []
            for it in p.items:
                if it[0] == "re":
                    items.append(("re", _rect_of(it[1])))
                elif it[0] == "l":
                    items.append(("l", Point(it[1]), Point(it[2])))
                else:
                    items.append(tuple(it))
            fill = tuple(p.fill) if p.fill is not None else None
            color = tuple(p.stroke) if p.stroke is not None else None
            if fill is not None and color is not None:
                typ = "fs"
            elif fill is not None:
                typ = "f"
            elif color is not None:
                typ = "s"
            else:
                typ = ""
            out.append({
                "rect": _rect_of(p.rect), "items": items, "type": typ,
                "fill": fill, "color": color, "width": float(p.line_width),
                "dashes": dashes[i] if i < len(dashes) else "[] 0",
                "even_odd": False, "closePath": False,
                "fill_opacity": 1.0, "stroke_opacity": 1.0,
                "lineCap": (0, 0, 0), "lineJoin": 0,
                "layer": "", "seqno": i, "level": 0,
            })
        self._drawings = out
        return out

    def get_cdrawings(self, extended=False):
        return self.get_drawings()

    # ---- 嵌入图像 ----

    def get_images(self, full=False):
        """PyMuPDF 形状的图像清单：
        (xref, smask, w, h, bpc, colorspace, alt_cs, name, filter, referencer)。
        同一 xref 在本页多次放置只出现一次（与 PyMuPDF 一致）。
        """
        out, seen = [], set()
        for e in self._doc._image_entries(self.number):
            if e["xref"] in seen:
                continue
            seen.add(e["xref"])
            out.append((e["xref"], 0, e["src"][0], e["src"][1], e["bits"],
                        e["colorspace"], "", e["name"], e["filter"], 0))
        return out

    def get_image_info(self, hashes=False, xrefs=False):
        """页面图像的逐次放置信息，至少含 bbox / width / height / xref。"""
        out = []
        for i, e in enumerate(self._doc._image_entries(self.number)):
            r = e["rect"]
            out.append({
                "number": i,
                "bbox": (r.x0, r.y0, r.x1, r.y1),
                "transform": (r.width, 0.0, 0.0, r.height, r.x0, r.y0),
                "width": e["src"][0], "height": e["src"][1],
                "colorspace": 3, "cs-name": e["colorspace"],
                "bpc": e["bits"], "xres": 96, "yres": 96,
                "size": 0, "digest": b"", "has-mask": False,
                "xref": e["xref"] if xrefs else 0,
            })
        return out

    def get_image_rects(self, xref, transform=False):
        """给定 xref 在本页的全部放置矩形。"""
        try:
            want = int(xref)
        except (TypeError, ValueError):
            return []
        return [e["rect"] for e in self._doc._image_entries(self.number)
                if e["xref"] == want]

    def get_fonts(self, full=False):
        """降级实现：pdfplumber / pypdfium2 拿不到页面字体资源表，恒返回空表。"""
        return []
    # ---- 渲染 ----

    def get_pixmap(self, matrix=None, dpi=None, colorspace=None, clip=None,
                   alpha=False, annots=True, **kwargs):
        """用 pypdfium2 渲染页面（或 clip 指定的 pt 矩形）为 Pixmap。

        - dpi 优先于 matrix；matrix 只取缩放分量，x/y 不等时取均值
          （pypdfium2 只支持等比缩放）；
        - colorspace=csGRAY 时输出单通道灰度（n == 1）；
        - alpha=False 时丢弃 alpha 通道；
        - annots=False 时不绘制注释与表单。
        """
        if dpi:
            sx = sy = float(dpi) / 72.0
        elif matrix is not None:
            m = matrix if isinstance(matrix, Matrix) else Matrix(matrix)
            sx, sy = abs(float(m.a)) or 1.0, abs(float(m.d)) or 1.0
        else:
            sx = sy = 1.0
        scale = (sx + sy) / 2.0
        if scale <= 0:
            scale = 1.0

        pw, ph = self._width, self._height
        crop = (0.0, 0.0, 0.0, 0.0)
        if clip is not None:
            c = Rect(clip) & Rect(0.0, 0.0, pw, ph)
            if not c.is_empty:
                cand = (max(0.0, c.x0), max(0.0, ph - c.y1),
                        max(0.0, pw - c.x1), max(0.0, c.y0))
                # pypdfium2 按 ceil(v*scale) 换算裁边，先确认至少剩 1 个像素
                src_w, src_h = math.ceil(pw * scale), math.ceil(ph * scale)
                cl = [math.ceil(v * scale) for v in cand]
                if (src_w - cl[0] - cl[2]) >= 1 and (src_h - cl[1] - cl[3]) >= 1:
                    crop = cand

        want_gray = colorspace is not None and getattr(colorspace, "n", 3) == 1
        bitmap = self._pdfium().render(
            scale=scale, crop=crop, grayscale=want_gray,
            draw_annots=bool(annots), may_draw_forms=bool(annots),
            rev_byteorder=True)
        samples, width, height, n = _bitmap_samples(bitmap, want_gray, alpha)
        cs = colorspace if colorspace is not None else (
            csGRAY if n == 1 else csRGB)
        res = int(round(72.0 * scale))
        return Pixmap(samples, width, height, n, alpha=(n == 4 and alpha),
                      colorspace=cs, xres=res, yres=res)

    # ---- 表格 ----

    def find_tables(self, clip=None, strategy=None, vertical_strategy=None,
                    horizontal_strategy=None, **kwargs):
        """用 pdfplumber 的 find_tables 实现；strategy 映射到 pdfplumber 设置。"""
        settings = {}
        if strategy:
            v, h = _STRATEGY_MAP.get(str(strategy).lower(), (None, None))
            if v:
                settings["vertical_strategy"] = v
                settings["horizontal_strategy"] = h
        if vertical_strategy:
            settings["vertical_strategy"] = str(vertical_strategy)
        if horizontal_strategy:
            settings["horizontal_strategy"] = str(horizontal_strategy)
        try:
            pg = self._plumber()
            if clip is not None:
                c = Rect(clip) & self.rect
                if not c.is_empty:
                    pg = pg.crop((c.x0, c.y0, c.x1, c.y1))
            found = pg.find_tables(settings) if settings else pg.find_tables()
        except Exception:
            return TableFinder([])
        return TableFinder([Table(t) for t in found])

    # ---- 注释 ----

    @property
    def annots(self):
        """注释集合。既支持 for a in page.annots，也支持 page.annots()。

        与 PyMuPDF 一致：Link 与 Popup 不计为注释。
        """
        if self._annots is None:
            try:
                raws = self._plumber().annots or []
            except Exception:
                raws = []
            items = [Annot(r) for r in raws]
            self._annots = _AnnotList(
                a for a in items if a.type[1] not in ("Link", "Popup"))
        return self._annots

    def first_annot(self):
        items = self.annots
        return items[0] if items else None

    def get_links(self):
        try:
            raws = self._plumber().annots or []
        except Exception:
            return []
        return [{"kind": 2 if r.get("uri") else 0,
                 "uri": r.get("uri") or "",
                 "from": Annot(r).rect}
                for r in raws if _annot_subtype(r) == "Link"]

# --------------------------------------------------------------------------
# Document
# --------------------------------------------------------------------------

_META_KEYS = ("format", "title", "author", "subject", "keywords", "creator",
              "producer", "creationDate", "modDate", "trapped", "encryption")

_META_MAP = {"Title": "title", "Author": "author", "Subject": "subject",
             "Keywords": "keywords", "Creator": "creator",
             "Producer": "producer", "CreationDate": "creationDate",
             "ModDate": "modDate", "Trapped": "trapped"}


class Document:
    """文档句柄（PyMuPDF Document 的子集）。只支持按路径打开本地 PDF。"""

    def __init__(self, filename=None, *args, **kwargs):
        if filename is None:
            raise ValueError("pymupdf_compat 只支持按文件路径打开 PDF")
        self.name = str(filename)
        self._backend = pdf_backend.open_pdf(self.name)
        self._pdfium = None
        self._pages = {}
        self._images = None
        self._closed = False

    # ---- 生命周期 ----

    def close(self):
        if self._closed:
            return
        self._closed = True
        self._pages.clear()
        try:
            self._backend.close()
        except Exception:
            pass
        if self._pdfium is not None:
            try:
                self._pdfium.close()
            except Exception:
                pass
            self._pdfium = None

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc, tb):
        self.close()
        return False

    @property
    def is_closed(self):
        return self._closed

    @property
    def is_pdf(self):
        return True

    # 打开成功即说明无需口令；pdfplumber 打不开会在 open_pdf 阶段抛异常。
    @property
    def needs_pass(self):
        return False

    @property
    def is_encrypted(self):
        return False

    def __repr__(self):
        return "Document(%r, %d pages)" % (self.name, self.page_count)

    # ---- 页面访问 ----

    @property
    def page_count(self):
        return self._backend.page_count

    def __len__(self):
        return self.page_count

    def load_page(self, index=0):
        n = self.page_count
        i = int(index)
        if i < 0:
            i += n
        if not (0 <= i < n):
            raise IndexError("页码越界：%r（共 %d 页）" % (index, n))
        if i not in self._pages:
            self._pages[i] = Page(self, i)
        return self._pages[i]

    def __getitem__(self, index):
        if isinstance(index, slice):
            return [self.load_page(i)
                    for i in range(*index.indices(self.page_count))]
        return self.load_page(index)

    def __iter__(self):
        for i in range(self.page_count):
            yield self.load_page(i)

    # ---- 内部句柄 ----

    def _plumber_doc(self):
        """pdf_backend 内部持有的 pdfplumber 文档句柄（表格 / 注释 / 图像流用）。"""
        pdf = getattr(self._backend, "_pdf", None)
        if pdf is None:
            raise RuntimeError("pdf_backend 未暴露 pdfplumber 句柄")
        return pdf

    def _plumber_page(self, index):
        return self._plumber_doc().pages[index]

    def _pdfium_doc(self):
        if self._pdfium is None:
            import pypdfium2 as pdfium
            self._pdfium = pdfium.PdfDocument(self.name)
        return self._pdfium

    # ---- 元数据与权限 ----

    @property
    def metadata(self):
        """PyMuPDF 风格的小写键元数据字典（由 pdfplumber 的 pdf.metadata 映射）。"""
        md = dict.fromkeys(_META_KEYS, "")
        md["format"] = "PDF"
        md["encryption"] = None
        try:
            raw = self._plumber_doc().metadata or {}
        except Exception:
            raw = {}
        for k, v in raw.items():
            key = _META_MAP.get(str(k))
            if key:
                md[key] = _as_text(v)
        return md

    @property
    def permissions(self):
        """权限位掩码。未加密或读不到 /P 时按"全部允许"返回。"""
        try:
            enc = getattr(self._plumber_doc().doc, "encryption", None)
            if not enc:
                return PDF_PERM_ALL
            params = enc[1] if len(enc) > 1 else {}
            value = params.get("P") if hasattr(params, "get") else None
            if value is None:
                return PDF_PERM_ALL
            try:
                from pdfminer.pdftypes import resolve1
                value = resolve1(value)
            except Exception:
                pass
            return int(value) & 0xFFFFFFFF
        except Exception:
            return PDF_PERM_ALL

    # ---- 字体（降级） ----

    def extract_font(self, xref=0, info_only=False, named=False):
        """与 PyMuPDF 同形的四元组 (basefont, ext, type, buffer)。

        pdfplumber / pypdfium2 都拿不到内嵌字体的二进制流，因此 buffer 恒为空，
        调用方据此得到"字体流体积 0"，不会抛异常。
        """
        if named:
            return {"name": "", "ext": "n/a", "type": "", "content": b""}
        return ("", "n/a", "", b"")

    # ---- 嵌入图像 ----

    def _build_image_index(self):
        """建立 页码 -> [图像放置] 索引。

        xref 取 pdfminer 图像流的对象号（即 PDF 真实 xref），使"同一 xref
        在一页里多次放置"的语义与 PyMuPDF 保持一致；个别拿不到对象号的流
        分配唯一负数占位，保证索引不串。
        """
        index = {}
        synthetic = -1
        for pno in range(self.page_count):
            entries = []
            try:
                images = self._plumber_page(pno).images or []
            except Exception:
                images = []
            for im in images:
                stream = im.get("stream")
                try:
                    xref = int(getattr(stream, "objid", 0) or 0)
                except (TypeError, ValueError):
                    xref = 0
                if xref <= 0:
                    xref, synthetic = synthetic, synthetic - 1
                src = im.get("srcsize") or (im.get("width"), im.get("height"))
                entries.append({
                    "page": pno,
                    "xref": xref,
                    "rect": Rect(float(im.get("x0") or 0.0),
                                 float(im.get("top") or 0.0),
                                 float(im.get("x1") or 0.0),
                                 float(im.get("bottom") or 0.0)),
                    "src": (int((src[0] if src else 0) or 0),
                            int((src[1] if src else 0) or 0)),
                    "bits": int(im.get("bits") or 8),
                    "colorspace": _cs_name(im.get("colorspace")),
                    "name": str(im.get("name") or ""),
                    "filter": _filter_name(stream),
                    "stream": stream,
                })
            index[pno] = entries
        self._images = index

    def _image_entries(self, page_index):
        if self._images is None:
            self._build_image_index()
        return self._images.get(int(page_index), [])

    def _find_image_entry(self, xref):
        try:
            want = int(xref)
        except (TypeError, ValueError):
            return None
        if self._images is None:
            self._build_image_index()
        for entries in self._images.values():
            for e in entries:
                if e["xref"] == want:
                    return e
        return None

    def extract_image(self, xref):
        """与 PyMuPDF 同形的图像字典（至少含 image / ext / width / height）。

        字节由 pypdfium2 提取（JPEG / JPEG2000 原样取出，其余重编码为 PNG）；
        pypdfium2 取不到时退回 pdfminer 的原始流。取不到则返回空 dict，
        调用方的 try/except 会把该项当作"无此图"处理。
        """
        entry = self._find_image_entry(xref)
        if entry is None:
            return {}
        data, ext = self._image_bytes(entry)
        if not data:
            return {}
        return {"ext": ext, "smask": 0, "colorspace": 3,
                "cs-name": entry["colorspace"], "bpc": entry["bits"],
                "width": entry["src"][0], "height": entry["src"][1],
                "xres": 96, "yres": 96, "image": data}

    def _image_bytes(self, entry):
        obj = self._pdfium_image_obj(entry)
        if obj is not None:
            try:
                buf = io.BytesIO()
                obj.extract(buf)
                data = buf.getvalue()
                if data:
                    return data, _guess_ext(data)
            except Exception:
                pass
        stream = entry.get("stream")
        name = entry.get("filter") or ""
        if stream is not None and name in ("DCTDecode", "JPXDecode"):
            try:
                data = bytes(stream.rawdata or b"")
                if data:
                    return data, ("jpeg" if name == "DCTDecode" else "jpx")
            except Exception:
                pass
        return b"", ""

    def _pdfium_image_obj(self, entry):
        """在 pypdfium2 页面上找到与该放置对应的图像对象（按源像素尺寸 + 位置）。"""
        try:
            import pypdfium2 as pdfium
            page = self._pdfium_doc()[entry["page"]]
            page_h = float(page.get_size()[1])
            objs = page.get_objects(filter=[pdfium.raw.FPDF_PAGEOBJ_IMAGE])
        except Exception:
            return None
        want, best, best_score = entry["rect"], None, None
        for obj in objs:
            try:
                left, bottom, right, top = obj.get_bounds()
            except Exception:
                continue
            got = Rect(left, page_h - top, right, page_h - bottom)
            score = (abs(got.x0 - want.x0) + abs(got.y0 - want.y0)
                     + abs(got.x1 - want.x1) + abs(got.y1 - want.y1))
            try:
                if tuple(int(v) for v in obj.get_px_size()) != entry["src"]:
                    score += 1.0e6          # 源像素尺寸不符，仅作最后兜底
            except Exception:
                score += 1.0e6
            if best_score is None or score < best_score:
                best, best_score = obj, score
        return best

# --------------------------------------------------------------------------
# 图像流的辅助信息
# --------------------------------------------------------------------------

def _ps_name(value):
    """PSLiteral / bytes / str -> 纯名字（去掉 / 前缀）。"""
    name = getattr(value, "name", None)
    if name:
        return str(name)
    if isinstance(value, bytes):
        return value.decode("latin-1", "ignore").lstrip("/")
    if isinstance(value, str):
        return value.lstrip("/")
    return ""


def _cs_name(colorspace):
    """pdfplumber 的 colorspace 字段 -> 色彩空间名（取不到时为空串）。"""
    if colorspace is None:
        return ""
    if isinstance(colorspace, (list, tuple)):
        for item in colorspace:
            name = _ps_name(item)
            if name:
                return name
        return ""
    return _ps_name(colorspace)


def _filter_name(stream):
    """图像流的最外层压缩过滤器名（DCTDecode / FlateDecode / ...）。"""
    if stream is None:
        return ""
    try:
        filters = stream.get_filters() or []
    except Exception:
        return ""
    if len(filters) != 1:                  # 多层过滤器时 rawdata 不可直接当图片
        return ""
    return _ps_name(filters[0][0])


def _guess_ext(data):
    """按魔数猜图片扩展名（与 PyMuPDF extract_image 的 ext 语义一致）。"""
    if data[:2] == b"\xff\xd8":
        return "jpeg"
    if data[:8] == b"\x89PNG\r\n\x1a\n":
        return "png"
    if data[:4] in (b"II*\x00", b"MM\x00*"):
        return "tiff"
    if data[:6] in (b"GIF87a", b"GIF89a"):
        return "gif"
    if data[:2] == b"BM":
        return "bmp"
    if data[:4] == b"\x00\x00\x00\x0c":
        return "jpx"
    return "png"


# --------------------------------------------------------------------------
# 模块级入口
# --------------------------------------------------------------------------

def open(filename=None, *args, **kwargs):        # noqa: A001 —— 与 PyMuPDF 同名
    """打开本地 PDF，返回 Document（等价于 pymupdf.open / fitz.open）。"""
    return Document(filename, *args, **kwargs)


# PyMuPDF 的别名
Pdf = Document
open_pdf = open

__all__ = [
    "open", "open_pdf", "Document", "Pdf", "Page", "Pixmap", "Annot",
    "Table", "TableFinder", "Rect", "IRect", "Point", "Matrix", "Colorspace",
    "csGRAY", "csRGB", "csCMYK", "CS_GRAY", "CS_RGB", "CS_CMYK",
    "PDF_PERM_PRINT", "PDF_PERM_MODIFY", "PDF_PERM_COPY", "PDF_PERM_ANNOTATE",
    "PDF_PERM_FORM", "PDF_PERM_ACCESSIBILITY", "PDF_PERM_ASSEMBLE",
    "PDF_PERM_PRINT_HQ", "PDF_PERM_ALL",
    "PDF_ANNOT_TEXT", "PDF_ANNOT_LINK", "PDF_ANNOT_HIGHLIGHT",
    "PDF_ANNOT_POPUP", "PDF_ANNOT_WIDGET", "PDF_ANNOT_UNKNOWN",
    "TEXT_FONT_SUPERSCRIPT", "TEXT_FONT_ITALIC", "TEXT_FONT_SERIFED",
    "TEXT_FONT_MONOSPACED", "TEXT_FONT_BOLD",
]