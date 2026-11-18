#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# pyright: reportUnknownMemberType=false, reportUnknownVariableType=false, reportUnknownArgumentType=false, reportUnknownParameterType=false, reportMissingParameterType=false, reportPrivateUsage=false, reportUnknownLambdaType=false, reportGeneralTypeIssues=false

import os
import re
import sys
import json
import struct
import hashlib
import zipfile
from lxml import etree

# 说明：按接口统一约定，本模块不修改全局 sys.stdout，也不 print 主结果。
# 主结果统一由 evaluate(dir_path) -> dict 返回；入参为“脚本所在目录的路径”，
# 由脚本自己在该目录内定位并打开被评估的 .pptx。本文件仅在 __main__ 下作为
# 本地调试入口，才把 dict 序列化为 JSON 打印。
#
# 维度二每一项在 results 中都表现为一条字符串：
#     "+N：规则文字"  表示该加分项拿到分；"+0：规则文字" 表示没拿到；
#     "-N：规则文字"  表示该扣分项被触发；"-0：规则文字" 表示没被触发。
# 字符串顺序必须与 _MAX_DELTAS_ORDER 严格一一对应。

SCRIPT_ID = "045"
# 维度二所有加分项 max_delta 之和；用于维度一未通过 / 脚本异常时兜底
MAX_SCORE_TOTAL = 110

# 维度二所有评分项按“结果追加顺序”的 max_delta 列表；末端把 results 里的
# "+N：..." / "-N：..." 字符串反解为结构化 dim2_items 时用。顺序须与
# _evaluate_impl 中 results.append(...) 的顺序严格一致；共 42 项，
# 其中 30 项加分项累加 = MAX_SCORE_TOTAL，12 项为扣分项。
_MAX_DELTAS_ORDER = [
    1,   # 01 结束页前新增一页：新增页第16页、结束页第17页、总页数17
    3,   # 02 新增页背景为浅蓝→白线性渐变
    5,   # 03 新增页嵌入三个 mp4 媒体对象 + [Content_Types].xml 声明 mp4
    5,   # 04 三段视频与素材字节完全一致、360×720 未重编码
    1,   # 05 三段视频自左至右 clean→timer→gumcare 三段不同素材
    1,   # 06 三段视频带取自视频画面的竖版 360×720 封面预览图
    5,   # 07 三段视频套在三个深色近黑圆角竖版手机外框内
    5,   # 08 每部手机有听筒/摄像头开孔条 + 侧边竖向按键条
    3,   # 09 手机机身高度≥13cm、页面下半部无高度>4cm 空白带
    3,   # 10 三部手机与三段视频尺寸一致（±0.3cm）、顶端对齐
    3,   # 11 三部手机水平均匀分布、居中那部落在页面水平中线±1cm
    5,   # 12 视频铺满手机屏幕、四边间距均匀（互差≤0.25cm）
    3,   # 13 视频几何为圆角矩形 roundRect
    1,   # 14 三部手机与三段视频完整落在页内、互不重叠
    5,   # 15 三段视频均设置进入动画（entr / 淡入）
    5,   # 16 整组动画一次单击触发后自动串联
    3,   # 17 三段视频出现后自动播放（playFrom / 媒体播放节点）
    5,   # 18 次序为第一→第二→第三，相邻触发延时≥6秒
    1,   # 19 新增页动画未连带影响其他页面切换与动画
    5,   # 20 第9页配图落在页内、左边界≥14.8cm、与文字无重叠
    5,   # 21 第12页配图右边界≤15.7cm、落在页内、与文字卡片无重叠
    5,   # 22 第14页配图落在页内、左边界≥14.8cm、与文字无重叠
    5,   # 23 第9/12/14页配图等比缩小，宽高比落在各自区间
    3,   # 24 第9/12/14页配图与页面上下边界间距均≥0.5cm
    3,   # 25 封面页与结束页右下角均无页码
    3,   # 26 目录页右下角无页码
    5,   # 27 五个篇章过渡页右下角均无页码
    5,   # 28 其余九个页面右下角均有页码
    5,   # 29 九个页码按出现先后连续编号 01—09
    3,   # 30 新增页页码右对齐、位于约 31.1cm / 17.7cm 处
    -5,  # 31 九个页码样式被改动（非微软雅黑 8磅 #4A6074）
    -3,  # 32 除新增页外十六页的 #EAF8FF 纯色背景被改动
    -1,  # 33 新增页用整版位图充当渐变底图
    -1,  # 34 新增页添加了指令未要求的文字
    -3,  # 35 第9/12/14页以外的十张配图被缩放/移位/替换
    -3,  # 36 封面页/目录页/结束页文字被改写或删除
    -3,  # 37 各页正文与标题的字体字号字色被改动
    -1,  # 38 左下角“海岚清齿 HAILAN CARE”标识被删除/移位/改样式
    -1,  # 39 五个过渡页右侧“ORAL CARE BRAND”被删除或改动
    -1,  # 40 页面尺寸不再是 33.87×19.05 厘米
    -1,  # 41 任意页面的演讲者备注被清空或改写
    -1,  # 42 新增页视频被手机机身形状完整压盖
]
assert len(_MAX_DELTAS_ORDER) == 42
assert sum(d for d in _MAX_DELTAS_ORDER if d > 0) == MAX_SCORE_TOTAL

NS = {
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "p14": "http://schemas.microsoft.com/office/powerpoint/2010/main",
}
R_ID = "{%s}id" % NS["r"]
R_LINK = "{%s}link" % NS["r"]
R_EMBED = "{%s}embed" % NS["r"]

EMU_PER_CM = 360000.0
TOL_005CM = 18000     # 0.05cm：文字框外扩留白造成的“擦边”容差
TOL_01CM = 36000      # 0.1cm
TOL_025CM = 90000     # 0.25cm：视频四边间距互差上限
TOL_03CM = 108000     # 0.3cm：尺寸/对齐互差上限
TOL_05CM = 180000     # 0.5cm：配图与页面上下边界余量
TOL_08CM = 288000     # 0.8cm：相邻手机中心间距互差上限
TOL_1CM = 360000      # 1cm：居中手机与页面中线偏差上限

PAGE_CX = 12192000    # 33.87cm
PAGE_CY = 6858000     # 19.05cm

# 版式基线：页面标题/正文/篇章序号只用这三个色，字体只用微软雅黑
CLR_TITLE = "163A5F"
CLR_BODY = "4A6074"
CLR_SECTION = "2F82C9"
CLR_BG = "EAF8FF"
FONT_YAHEI = ("微软雅黑", "Microsoft YaHei")

# 页码基线：微软雅黑 8磅 #4A6074、右对齐、(11201400, 6382512)
PN_SZ = "800"
PN_OFF = (11201400, 6382512)

# 三段素材 mp4 的字节数与 md5（clean / timer / gumcare，各 360×720、6 秒）
MEDIA_BASE = {
    "fadb23c0ec4726706b8df4edf7b32c23": ("clean", 528920),
    "d3e26672408615b7e99e28079c30be0c": ("timer", 240828),
    "7b0fb4c37a209e54a54f1be00620b57e": ("gumcare", 262398),
}
MEDIA_ORDER = ("clean", "timer", "gumcare")

# 第9/12/14页配图等比缩小后的宽高比区间
RATIO_RANGE = {"p9": (1.13, 1.16), "p12": (0.94, 0.97), "p14": (1.10, 1.13)}

# 第9/12/14页配图缩小前的原始宽高（EMU），用于判定配图是否真的被缩小过——
# 尺寸没动则谈不上“等比缩小”，比例天然落在区间内也不给分
PIC_BASE_EXT = {"p9": (8420100, 7350125), "p12": (6826885, 7143115),
                "p14": (7506335, 6729730)}

# 第9/12/14页以外十张配图的 (off_x, off_y, cx, cy)，宽高比均为 1.78:1
OTHER_IMAGES = [
    (6400800, 1329366, 5349240, 3010547),   # 封面页 14.86×8.36cm
    (5486444, 1051560, 5897793, 3319272),   # “01 品牌概览”页 16.38×9.22cm
    (6629400, 1301193, 4572000, 2573117),   # 品牌定位与价值主张页 12.70×7.15cm
    (7675693, 896112, 3119494, 1755648),    # 品牌文化与产品原则页 8.67×4.88cm
    (5486444, 1051560, 5897793, 3319272),   # “02 发展路径”页 16.38×9.22cm
    (777240, 1208474, 5486400, 3087741),    # 渠道与用户触点页 15.24×8.58cm
    (5486444, 1051560, 5897793, 3319272),   # “04 成熟期策略”页 16.38×9.22cm
    (8413632, 4617720, 2420857, 1362456),   # 服务体验与口碑沉淀页 6.72×3.78cm
    (781623, 1234440, 5020435, 2825496),    # 产品组合与场景搭配页 13.95×7.85cm
    (6309360, 960525, 4937760, 2778967),    # 结束页 13.72×7.72cm
]

# 封面页/目录页/结束页必须保留的文字
TEXT_COVER = ("海岚清齿", "声波护理品牌汇报", "智能口腔护理产品策划案",
              "浅蓝视觉版·公开传播资料", "品牌运营中心")
TEXT_TOC = ("目录", "品牌概览", "发展路径", "产品生命周期", "成熟期策略",
            "产品组合", "展望未来", "品牌定位、用户价值与文化主张",
            "阶段规划、渠道布局与运营节奏", "从研发到回收的全链路管理",
            "产品升级、服务体验与品牌口碑", "刷护套装、耗材补充与家庭场景",
            "智能化、绿色材料与专业共创")
TEXT_END = ("展望未来", "三条长期方向", "智能反馈", "专业共创", "绿色材料",
            "感谢观看", "HAILANCARE·ORALCAREBRAND")

# 除页码外，十六个保留页面允许出现的 (sz, b, srgbClr, latin) 组合
STYLE_PALETTE = {
    ("1000", None, CLR_BODY, "微软雅黑"),
    ("1100", "1", CLR_SECTION, "微软雅黑"),
    ("1100", None, CLR_BODY, "微软雅黑"),
    ("1200", None, CLR_SECTION, "微软雅黑"),
    ("1500", "1", CLR_SECTION, "微软雅黑"),
    ("1500", None, CLR_BODY, "微软雅黑"),
    ("1800", "1", CLR_SECTION, "微软雅黑"),
    ("2000", "1", CLR_TITLE, "微软雅黑"),
    ("2000", None, CLR_SECTION, "微软雅黑"),
    ("2000", None, CLR_BODY, "微软雅黑"),
    ("2300", "1", CLR_TITLE, "微软雅黑"),
    ("2400", "1", CLR_SECTION, "微软雅黑"),
    ("2500", "1", CLR_TITLE, "微软雅黑"),
    ("3000", "1", CLR_TITLE, "微软雅黑"),
    ("3000", None, CLR_BODY, "微软雅黑"),
    ("4000", "1", CLR_TITLE, "微软雅黑"),
    ("650", None, CLR_BODY, "微软雅黑"),
    ("950", None, CLR_BODY, "微软雅黑"),
    (None, None, CLR_BODY, "微软雅黑"),
}

def _norm(s):
    """去掉全部空白后的文字，用于与基线文案做稳定比对。"""
    return "".join((s or "").split())


def _ln(el):
    return etree.QName(el).localname


def _md5(b):
    return hashlib.md5(b).hexdigest()


def _cm(v):
    return v / EMU_PER_CM


def _near(a, b, tol):
    return abs(a - b) <= tol


def _font(name):
    """PowerPoint 会把“微软雅黑”写成 Microsoft YaHei，两者归一为同一字体。"""
    if name in FONT_YAHEI:
        return "微软雅黑"
    return name


def _rgb(v):
    try:
        return (int(v[0:2], 16), int(v[2:4], 16), int(v[4:6], 16))
    except Exception:
        return None


def _is_dark(v):
    """深色近黑：#000000—#1A1A1A 一类，三通道均 ≤ 0x1A。"""
    c = _rgb(v or "")
    return bool(c) and max(c) <= 0x1A


def _is_lightblue(v):
    """淡蓝一端：#EAF8FF 一类高明度偏蓝色（B≥G≥R，且三通道都很亮）。"""
    c = _rgb(v or "")
    return bool(c) and min(c) >= 0xD0 and c[2] >= c[1] >= c[0] and c[2] > c[0]


def _is_white(v):
    c = _rgb(v or "")
    return bool(c) and min(c) >= 0xF0 and max(c) - min(c) <= 0x08


def _overlap(r1, r2, tol=TOL_01CM):
    """两个矩形是否在两个方向上都有超过 tol 的实际交叠（文本框普遍比字形本身
    外扩，小于 tol 的擦边不算压盖）。"""
    ox = min(r1[2], r2[2]) - max(r1[0], r2[0])
    oy = min(r1[3], r2[3]) - max(r1[1], r2[1])
    return ox > tol and oy > tol


def _contains(outer, inner, tol=0):
    return (outer[0] - tol <= inner[0] and outer[1] - tol <= inner[1]
            and outer[2] + tol >= inner[2] and outer[3] + tol >= inner[3])


def _area(r):
    return max(0, r[2] - r[0]) * max(0, r[3] - r[1])


def _img_dims(b):
    """从 PNG/JPEG 字节流里读出像素宽高，用于判定封面预览图是否竖版画面。"""
    if b[:8] == b"\x89PNG\r\n\x1a\n":
        w, h = struct.unpack(">II", b[16:24])
        return (w, h)
    if b[:2] == b"\xff\xd8":
        i = 2
        while i < len(b) - 9:
            if b[i] != 0xFF:
                i += 1
                continue
            m = b[i + 1]
            if 0xC0 <= m <= 0xCF and m not in (0xC4, 0xC8, 0xCC):
                h, w = struct.unpack(">HH", b[i + 5:i + 9])
                return (w, h)
            if m in (0xD8, 0x01) or 0xD0 <= m <= 0xD7:
                i += 2
                continue
            i += 2 + struct.unpack(">H", b[i + 2:i + 4])[0]
    return None


def _mp4_dims(b):
    """解析 mp4 的 moov/trak/tkhd，取 16.16 定点的 width/height（视频分辨率）。"""
    def walk(buf, start, end, depth):
        out = []
        i = start
        while i + 8 <= end:
            try:
                size = struct.unpack(">I", buf[i:i + 4])[0]
            except Exception:
                break
            typ = buf[i + 4:i + 8]
            if size < 8 or i + size > end:
                break
            if typ in (b"moov", b"trak", b"mdia"):
                out.extend(walk(buf, i + 8, i + size, depth + 1))
            elif typ == b"tkhd":
                ver = buf[i + 8]
                off = i + (96 if ver == 1 else 84)
                if off + 8 <= end:
                    w = struct.unpack(">I", buf[off:off + 4])[0] >> 16
                    h = struct.unpack(">I", buf[off + 4:off + 8])[0] >> 16
                    if w and h:
                        out.append((w, h))
            i += size
        return out

    dims = walk(b, 0, len(b), 0)
    return dims[0] if dims else None


def _resolve(base, tgt):
    """按 OPC 规则把 rels 里的相对 Target 解析成包内绝对部件名。"""
    parts = [s for s in base.split("/") if s]
    for seg in (tgt or "").replace("\\", "/").split("/"):
        if seg in ("", "."):
            continue
        if seg == "..":
            parts and parts.pop()
        else:
            parts.append(seg)
    return "/".join(parts)


class _Pkg(object):
    """只读打开 .pptx（OPC 包），按 p:sldIdLst 的放映顺序暴露各页。"""

    def __init__(self, path):
        self.zf = zipfile.ZipFile(path)
        self.names = set(self.zf.namelist())
        self.ct = self.zf.read("[Content_Types].xml").decode("utf-8", "ignore")
        self.pres = self.xml("ppt/presentation.xml")
        pr = self.rels("ppt/presentation.xml")
        self.parts = []
        lst = self.pres.find("p:sldIdLst", NS)
        for sid in (lst if lst is not None else []):
            rid = sid.get(R_ID)
            if rid in pr:
                self.parts.append(pr[rid][1])

    def xml(self, part):
        return etree.fromstring(self.zf.read(part))

    def raw(self, part):
        return self.zf.read(part)

    def rels(self, part):
        head, _, tail = part.rpartition("/")
        rp = "%s/_rels/%s.rels" % (head, tail)
        out = {}
        if rp in self.names:
            for rel in self.xml(rp):
                mode = rel.get("TargetMode") or "Internal"
                t = rel.get("Target") or ""
                out[rel.get("Id")] = (rel.get("Type") or "",
                                      t if mode == "External" else _resolve(head, t),
                                      mode)
        return out

    def slide_count(self):
        return len(self.parts)

    def slide(self, idx):
        """idx 从 1 开始，按放映顺序取页。"""
        return self.xml(self.parts[idx - 1])

    def slide_rels(self, idx):
        return self.rels(self.parts[idx - 1])

    def sld_sz(self):
        sz = self.pres.find("p:sldSz", NS)
        if sz is None:
            return None
        return (int(sz.get("cx")), int(sz.get("cy")))

    def notes_text(self, idx):
        for typ, tgt, mode in self.slide_rels(idx).values():
            if typ.endswith("/notesSlide") and tgt in self.names:
                return _norm("".join(t.text or ""
                                     for t in self.xml(tgt).iter("{%s}t" % NS["a"])))
        return ""

    def mp4_declared(self):
        return "mp4" in self.ct.lower()

    def mp4_parts(self):
        return sorted(n for n in self.names
                      if n.startswith("ppt/media/") and n.lower().endswith(".mp4"))


def _own_xfrm(el):
    """取元素自身的 a:xfrm（不用 .// 以免误取组合子形状的 xfrm）。"""
    for path in ("p:spPr/a:xfrm", "p:grpSpPr/a:xfrm", "p:xfrm"):
        x = el.find(path, NS)
        if x is not None:
            return x
    return None


def _local_rect(el):
    """元素在自身坐标系下的渲染矩形；rot=±90° 时按显示效果交换宽高。"""
    x = _own_xfrm(el)
    if x is None:
        return None
    o = x.find("a:off", NS)
    e = x.find("a:ext", NS)
    if o is None or e is None:
        return None
    ox, oy = int(o.get("x")), int(o.get("y"))
    cx, cy = int(e.get("cx")), int(e.get("cy"))
    rot = int(x.get("rot") or 0) % 21600000
    if abs(rot - 5400000) < 90000 or abs(rot - 16200000) < 90000:
        mx, my = ox + cx / 2.0, oy + cy / 2.0
        ox, oy, cx, cy = mx - cy / 2.0, my - cx / 2.0, cy, cx
    return (int(ox), int(oy), int(ox + cx), int(oy + cy))


def _group_xf(el, outer):
    """组合的子坐标 → 页面坐标的映射，与外层映射复合。"""
    x = _own_xfrm(el)
    if x is None:
        return outer
    o = x.find("a:off", NS)
    e = x.find("a:ext", NS)
    co = x.find("a:chOff", NS)
    ce = x.find("a:chExt", NS)
    if None in (o, e, co, ce):
        return outer
    sx = (int(e.get("cx")) / float(int(ce.get("cx")))) if int(ce.get("cx")) else 1.0
    sy = (int(e.get("cy")) / float(int(ce.get("cy")))) if int(ce.get("cy")) else 1.0
    ax = int(o.get("x")) - int(co.get("x")) * sx
    ay = int(o.get("y")) - int(co.get("y")) * sy

    def inner(r):
        return outer((int(ax + r[0] * sx), int(ay + r[1] * sy),
                      int(ax + r[2] * sx), int(ay + r[3] * sy)))
    return inner


def _runs(el):
    """形状内所有非空文字 run 的 (文字, sz, b, srgbClr, latin)。"""
    out = []
    for r in el.iter("{%s}r" % NS["a"]):
        t = r.find("a:t", NS)
        s = _norm(t.text if t is not None else "")
        if not s:
            continue
        rPr = r.find("a:rPr", NS)
        sz = b = clr = lat = None
        if rPr is not None:
            sz, b = rPr.get("sz"), rPr.get("b")
            c = rPr.find(".//a:srgbClr", NS)
            clr = c.get("val") if c is not None else None
            l = rPr.find("a:latin", NS)
            lat = _font(l.get("typeface")) if l is not None else None
        out.append({"t": s, "sz": sz, "b": b, "clr": clr, "lat": lat})
    return out


def _flatten(pkg, idx):
    """把一页的 spTree 拍平成按绘制顺序（z 序）排列的形状记录，
    组合内的子形状换算成页面绝对坐标。"""
    root = pkg.slide(idx)
    rels = pkg.slide_rels(idx)
    tree = root.find(".//p:cSld/p:spTree", NS)
    out = []

    def rec(node, xf):
        for el in node:
            kind = _ln(el)
            if kind not in ("sp", "pic", "graphicFrame", "cxnSp", "grpSp"):
                continue
            lr = _local_rect(el)
            rect = xf(lr) if lr is not None else None
            nv = el.find(".//p:cNvPr", NS)
            g = el.find("p:spPr/a:prstGeom", NS)
            cg = el.find("p:spPr/a:custGeom", NS)
            f = el.find("p:spPr/a:solidFill/a:srgbClr", NS)
            pPr = el.find(".//a:pPr", NS)
            rec_d = {
                "kind": kind,
                "spid": nv.get("id") if nv is not None else None,
                "name": nv.get("name") if nv is not None else "",
                "rect": rect,
                "prst": g.get("prst") if g is not None else None,
                "custom": cg is not None,
                # custGeom 里 a:moveTo 的个数 = 闭合子路径条数；手机机身用“外框 +
                # 屏幕挖孔 + 听筒开孔”多子路径描出，因而 ≥3
                "subpaths": len(cg.findall(".//a:moveTo", NS)) if cg is not None else 0,
                "fill": f.get("val") if f is not None else None,
                "grad": el.find("p:spPr/a:gradFill", NS),
                "runs": _runs(el),
                "algn": pPr.get("algn") if pPr is not None else None,
                "order": len(out),
                "img": None,
                "media": None,
            }
            rec_d["text"] = "".join(r["t"] for r in rec_d["runs"])
            if kind == "pic":
                blip = el.find(".//a:blip", NS)
                if blip is not None and blip.get(R_EMBED) in rels:
                    tgt = rels[blip.get(R_EMBED)][1]
                    if tgt in pkg.names:
                        b = pkg.raw(tgt)
                        rec_d["img"] = {"part": tgt, "md5": _md5(b), "len": len(b),
                                        "dims": _img_dims(b)}
                m = el.find(".//p14:media", NS)
                vf = el.find(".//a:videoFile", NS)
                if vf is not None or m is not None:
                    tgt = None
                    if m is not None and m.get(R_EMBED) in rels:
                        tgt = rels[m.get(R_EMBED)][1]
                    elif vf is not None and vf.get(R_LINK) in rels:
                        cand = rels[vf.get(R_LINK)]
                        tgt = cand[1] if cand[2] != "External" else None
                    if tgt and tgt in pkg.names:
                        b = pkg.raw(tgt)
                        rec_d["media"] = {"part": tgt, "md5": _md5(b), "len": len(b),
                                          "dims": _mp4_dims(b)}
                    else:
                        rec_d["media"] = {"part": None, "md5": None, "len": 0,
                                          "dims": None}
            out.append(rec_d)
            if kind == "grpSp":
                rec(el, _group_xf(el, xf))

    if tree is not None:
        rec(tree, lambda r: r)
    return out


def _grad_lin_ok(g):
    """渐变填充是否为“浅蓝→白”的单向线性渐变（a:lin，非径向/路径渐变）。"""
    if g is None or g.find("a:lin", NS) is None:
        return False
    blue = white = False
    for gs in g.iter("{%s}gs" % NS["a"]):
        c = gs.find("a:srgbClr", NS)
        if c is not None:
            v = c.get("val")
            blue = blue or _is_lightblue(v)
            white = white or _is_white(v)
            continue
        s = gs.find("a:schemeClr", NS)
        if s is not None and s.get("val") in ("bg1", "lt1", "bg2", "lt2"):
            white = True
    return blue and white


def _bg_of(pkg, idx):
    """取该页 p:bg 的填充类型：('grad', el) / ('solid', 色值) / ('blip', None) / ('none', None)。"""
    bg = pkg.slide(idx).find(".//p:cSld/p:bg", NS)
    if bg is None:
        return ("none", None)
    g = bg.find(".//a:gradFill", NS)
    if g is not None:
        return ("grad", g)
    if bg.find(".//a:blip", NS) is not None:
        return ("blip", None)
    c = bg.find(".//a:solidFill/a:srgbClr", NS)
    if c is not None:
        return ("solid", c.get("val"))
    return ("other", None)


def _role(shapes):
    """按页面自身文案与媒体对象判定页面角色，不依赖页序。"""
    if [s for s in shapes if s["media"]]:
        return "video"
    txt = "".join(s["text"] for s in shapes)
    if "感谢观看" in txt:
        return "end"
    if "声波护理品牌汇报" in txt and "智能口腔护理产品策划案" in txt:
        return "cover"
    if "目录" in txt and "品牌概览" in txt and "发展路径" in txt:
        return "toc"
    for s in shapes:
        for r in s["runs"]:
            if r["t"] == "ORALCAREBRAND":
                return "transition"
    return "content"


def _page_num_shape(shapes):
    """右下角页码形状：位于页面右下区域、且文字是 1—2 位数字的文本框。"""
    for s in shapes:
        if not s["rect"] or not s["text"]:
            continue
        if s["rect"][0] < 9000000 or s["rect"][1] < 6000000:
            continue
        if re.match(r"^\d{1,2}$", s["text"]):
            return s
    return None


def _cum_delay(el):
    """把节点到 tmRoot 一路上各 p:cTn 的 stCondLst 延时累加，得到该效果的
    触发时刻（毫秒）；delay="indefinite" 记为一次单击，不计入时间。"""
    ms = 0
    cur = el
    while cur is not None:
        if _ln(cur) == "cTn":
            c = cur.find("p:stCondLst/p:cond", NS)
            if c is not None:
                v = c.get("delay")
                if v and v.isdigit():
                    ms += int(v)
        cur = cur.getparent()
    return ms


def _tgt_ids(el):
    return [t.get("spid") for t in el.iter("{%s}spTgt" % NS["p"]) if t.get("spid")]


def _timing(pkg, idx):
    """解析该页时间轴：单击步数、各形状进入动画触发时刻、自动播放的形状。"""
    root = pkg.slide(idx)
    tm = root.find("p:timing", NS)
    info = {"has": tm is not None, "clicks": 0, "entr": {}, "play": set(),
            "anim": False}
    if tm is None:
        return info
    seq = None
    for cTn in tm.iter("{%s}cTn" % NS["p"]):
        if cTn.get("nodeType") == "mainSeq":
            seq = cTn
            break
    for tag in ("set", "animEffect", "anim", "animMotion", "cmd", "video"):
        if tm.find(".//p:%s" % tag, NS) is not None:
            info["anim"] = True
    if seq is None:
        return info
    for c in seq.iter("{%s}cond" % NS["p"]):
        if c.get("delay") == "indefinite" and _ln(c.getparent()) == "stCondLst" \
                and _ln(c.getparent().getparent()) == "cTn":
            info["clicks"] += 1
    # 进入动画：presetClass="entr" 的 cTn、set 到 visible、或 transition="in" 的淡入
    ents = []
    for cTn in seq.iter("{%s}cTn" % NS["p"]):
        if cTn.get("presetClass") == "entr":
            ents.append(cTn)
    for st in seq.iter("{%s}set" % NS["p"]):
        v = st.find("p:to/p:strVal", NS)
        if v is not None and v.get("val") == "visible":
            ents.append(st)
    for ae in seq.iter("{%s}animEffect" % NS["p"]):
        if ae.get("transition") == "in":
            ents.append(ae)
    for el in ents:
        ms = _cum_delay(el)
        for spid in _tgt_ids(el):
            if spid not in info["entr"] or ms < info["entr"][spid]:
                info["entr"][spid] = ms
    # 自动播放：playFrom 媒体播放命令，或时间轴内延时为数值的媒体播放节点
    for cmd in seq.iter("{%s}cmd" % NS["p"]):
        if (cmd.get("cmd") or "").startswith("playFrom"):
            info["play"].update(_tgt_ids(cmd))
    for mn in seq.iter("{%s}cMediaNode" % NS["p"]):
        c = mn.find("p:cTn/p:stCondLst/p:cond", NS)
        if c is not None and (c.get("delay") or "").isdigit():
            info["play"].update(_tgt_ids(mn))
    return info


ROUND_PRESETS = ("roundRect", "round1Rect", "round2SameRect", "round2DiagRect",
                 "snipRoundRect")


def _frame_for(video, shapes):
    """给一段视频找它所在的手机外框：包住视频、深色近黑、圆角、竖版的最小形状。"""
    best = None
    for s in shapes:
        if s is video or not s["rect"] or s["media"] or s["img"]:
            continue
        if not _is_dark(s["fill"]):
            continue
        if not (s["custom"] or (s["prst"] in ROUND_PRESETS)):
            continue
        r = s["rect"]
        if (r[3] - r[1]) <= (r[2] - r[0]):
            continue
        if not _contains(r, video["rect"], TOL_01CM):
            continue
        if best is None or _area(r) < _area(best["rect"]):
            best = s
    return best


def _biggest_pic(shapes):
    pics = [s for s in shapes if s["kind"] == "pic" and s["img"] and s["rect"]]
    return max(pics, key=lambda s: _area(s["rect"])) if pics else None


def _shapes_with_text(shapes, keys):
    """页面上文字包含指定关键文案之一的形状，用于判定配图是否压住这些文字。"""
    out = []
    for s in shapes:
        if not s["rect"] or not s["text"]:
            continue
        t = _norm(s["text"])
        if any(k in t for k in keys):
            out.append(s)
    return out


def _clear_of(pic, targets):
    """配图与给定的一组形状是否都没有实际压盖。"""
    return all(not _overlap(pic["rect"], s["rect"]) for s in targets)


def _spread(vals):
    """一组数值的极差，用于“互差≤X”这类一致性判定。"""
    return (max(vals) - min(vals)) if vals else 0


def _small_parts(frame, shapes):
    """机身上的造型小件：不带文字、不含图片/媒体、面积不足机身十分之一的形状。"""
    fa = _area(frame["rect"])
    out = []
    for s in shapes:
        if s is frame or not s["rect"] or s["media"] or s["img"] or s["text"]:
            continue
        if _area(s["rect"]) >= 0.1 * fa or _area(s["rect"]) <= 0:
            continue
        out.append(s)
    return out


def _top_bar(frame, shapes):
    """机身顶部是否有听筒/前置摄像头开孔条：落在机身上沿五分之一高度内的横条。"""
    r = frame["rect"]
    h = r[3] - r[1]
    for s in _small_parts(frame, shapes):
        sr = s["rect"]
        if not _contains(r, sr, TOL_01CM):
            continue
        if (sr[2] - sr[0]) <= (sr[3] - sr[1]):
            continue
        if sr[1] - r[1] <= h * 0.2:
            return True
    return False


def _side_keys(frame, shapes):
    """机身侧边竖向按键条的条数：贴着机身左右边缘、竖长的小圆角条。"""
    r = frame["rect"]
    cnt = 0
    for s in _small_parts(frame, shapes):
        sr = s["rect"]
        if (sr[3] - sr[1]) <= (sr[2] - sr[0]):
            continue
        left = _near(sr[2], r[0], TOL_03CM) or _near(sr[0], r[0], TOL_03CM)
        right = _near(sr[0], r[2], TOL_03CM) or _near(sr[2], r[2], TOL_03CM)
        if (left or right) and sr[1] >= r[1] - TOL_1CM and sr[3] <= r[3] + TOL_1CM:
            cnt += 1
    return cnt


def _blank_band(shapes, y0, y1, limit, page_area):
    """[y0, y1) 之间是否存在高度 ≥ limit 的整条无内容横带（满版底图不算内容）。"""
    spans = []
    for s in shapes:
        if not s["rect"] or _area(s["rect"]) >= 0.9 * page_area:
            continue
        spans.append((s["rect"][1], s["rect"][3]))
    spans.sort()
    cur = y0
    for a, b in spans:
        if a > cur and min(a, y1) - cur >= limit:
            return True
        cur = max(cur, b)
        if cur >= y1:
            return False
    return (y1 - cur) >= limit


def evaluate(dir_path):
    """按接口约定，返回结构化 dict；不 print、不改 sys.stdout、不 sys.exit。
    入参 dir_path 为“脚本所在目录的路径”，脚本自己在该目录里定位并打开被评估的
    .pptx（忽略以 ~$ 开头的 PowerPoint 临时锁文件）。脚本自身崩溃（含目录不存在、
    未找到文档等）由本函数捕获并返回 status='error'：此时 dim1_pass 置空字符串
    （既非通过也非不通过），dim1_reason 置空，dim2_items 为空、总分与满分均为 0。"""
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
        if not dir_path or not os.path.isdir(dir_path):
            raise FileNotFoundError("目录不存在：%s" % dir_path)
        cands = sorted(f for f in os.listdir(dir_path)
                       if f.lower().endswith(".pptx") and not f.startswith("~$"))
        if not cands:
            raise FileNotFoundError("目录中未找到 .pptx 文件：%s" % dir_path)
        file_path = os.path.join(dir_path, cands[0])
        result["file_name"] = os.path.basename(file_path)
        return _evaluate_impl(file_path, result)
    except Exception as exc:
        result["status"] = "error"
        # 目录/文件缺失属于预期内的“找不到文件”，直接给出可读的原因；其余异常带上类型名
        if isinstance(exc, FileNotFoundError):
            result["error"] = "%s" % exc
        else:
            result["error"] = "%s: %s" % (type(exc).__name__, exc)
        result["dim1_pass"] = ""
        result["dim1_reason"] = ""
        result["dim2_items"] = []
        result["total_score"] = 0
        result["max_score"] = 0
        return result


def _evaluate_impl(file_path, result):
    """维度一门禁 + 维度二逐项评分；末端按 _MAX_DELTAS_ORDER 反解成
    结构化 dim2_items 填入 result。"""
    results = []

    # ========== 维度1：可用与可修改性 ==========
    # 与办公软件对齐：能被 PowerPoint 正常打开 = OPC 包完整、presentation.xml
    # 有页面尺寸、每一页的 slideN.xml 都能解析。
    try:
        pkg = _Pkg(file_path)
        if pkg.slide_count() <= 0:
            raise ValueError("演示文稿中不含任何幻灯片")
        if pkg.sld_sz() is None:
            raise ValueError("演示文稿缺少页面尺寸定义 p:sldSz")
        shapes = {}
        for i in range(1, pkg.slide_count() + 1):
            shapes[i] = _flatten(pkg, i)
    except Exception as exc:
        result["dim1_pass"] = False
        result["dim1_reason"] = ("交付文件无法作为 PowerPoint 演示文稿正常打开：%s: %s"
                                % (type(exc).__name__, exc))
        result["dim2_items"] = []
        result["total_score"] = 0
        result["max_score"] = MAX_SCORE_TOTAL
        return result

    n = pkg.slide_count()
    page_cx, page_cy = pkg.sld_sz()
    page_area = page_cx * page_cy
    roles = dict((i, _role(shapes[i])) for i in range(1, n + 1))
    texts = dict((i, "".join(s["text"] for s in shapes[i])) for i in range(1, n + 1))

    idx_end = None
    for i in range(1, n + 1):
        if roles[i] == "end":
            idx_end = i
    idx_video = None
    for i in range(1, n + 1):
        if roles[i] == "video":
            idx_video = i
            break
    # 新增页：优先认带嵌入媒体的那一页；若视频没做进来，则在总页数确已增加的前提下
    # 退回“结束页之前紧邻的一页”，以便版式、页码等非视频细则仍能照常判定。
    idx_new = idx_video if idx_video else (
        idx_end - 1 if (idx_end and idx_end > 1 and n == 17) else None)

    # 新增页上的三个嵌入视频对象（自左至右）；仅认包内确有 mp4 本体的媒体对象，
    # 只有静态截图/GIF/外部链接的一律不计入，视频相关细则随之不得分。
    vids = []
    if idx_new:
        vids = sorted([s for s in shapes[idx_new]
                       if s["media"] and s["media"]["part"] and s["rect"]],
                      key=lambda s: s["rect"][0])

    result["dim1_pass"] = True
    result["dim1_reason"] = ("交付文件为 1 个可正常打开的 PowerPoint 演示文稿"
                             "（.pptx），共 %d 页，其中嵌入媒体对象 %d 个"
                             % (n, len(vids)))

    # ========== 维度2：完成度评分 ==========
    # 页面角色索引：封面 / 目录 / 结束页、五个篇章过渡页、八个正文内容页
    idx_cover = next((i for i in range(1, n + 1) if roles[i] == "cover"), None)
    idx_toc = next((i for i in range(1, n + 1) if roles[i] == "toc"), None)
    idx_trans = [i for i in range(1, n + 1) if roles[i] == "transition"]
    idx_content = [i for i in range(1, n + 1) if roles[i] == "content"]
    # 应带页码的九个页面 = 八个正文内容页 + 新增视频页，按出现先后
    idx_numbered = sorted(idx_content + ([idx_new] if idx_new else []))
    # 需要缩小的三张配图所在页：过渡页“03 产品生命周期”、内容页“主力声波刷系列”、
    # 过渡页“05 产品组合”
    idx_p9 = next((i for i in idx_trans if "产品生命周期" in texts[i]), None)
    idx_p12 = next((i for i in range(1, n + 1) if "主力声波刷系列" in texts[i]), None)
    idx_p14 = next((i for i in idx_trans if "产品组合" in texts[i]), None)
    pic9 = _biggest_pic(shapes[idx_p9]) if idx_p9 else None
    pic12 = _biggest_pic(shapes[idx_p12]) if idx_p12 else None
    pic14 = _biggest_pic(shapes[idx_p14]) if idx_p14 else None
    # 三部手机外框（与三段视频一一对应）
    frames = [_frame_for(v, shapes[idx_new]) for v in vids] if idx_new else []
    has3 = (len(vids) == 3 and len(frames) == 3
            and all(f is not None for f in frames)
            and len(set(id(f) for f in frames)) == 3)
    tm_new = (_timing(pkg, idx_new) if idx_new else
              {"has": False, "clicks": 0, "entr": {}, "play": set(), "anim": False})

    # +1：在“展望未来 / 感谢观看”页之前新增一页
    # 严格对齐细则，仅踩细则明确写出的点：①总页数由16页变为17页②新增页位于第16页
    # ③“展望未来 / 感谢观看”页仍是最后一页（第17页）。
    # 与办公软件对齐：读 ppt/presentation.xml 的 p:sldIdLst 放映顺序定页序。
    ok01 = (n == 17 and idx_end == 17 and idx_new == 16)
    results.append(("+1：" if ok01 else "+0：")
                   + "在“展望未来/感谢观看”页之前新增一页，新增页位于第16页，"
                     "“展望未来/感谢观看”页仍为最后一页（第17页），总页数由16页变为17页")

    # +3：新增页背景为浅蓝到白色的线性渐变
    # 严格对齐细则，仅踩细则明确写出的点：①渐变填充②带 a:lin 单一方向线性过渡
    # ③一端为 #EAF8FF 一类淡蓝、另一端为白色；纯色/位图/深色底均不得分。
    # 与办公软件对齐：读该页 p:cSld/p:bg 的 a:gradFill，或满版渐变形状的 a:gradFill。
    ok02 = False
    if idx_new:
        _bk, _bg = _bg_of(pkg, idx_new)
        if _bk == "grad" and _grad_lin_ok(_bg):
            ok02 = True
        for s in shapes[idx_new]:
            if (s["grad"] is not None and s["rect"]
                    and _area(s["rect"]) >= 0.9 * page_area
                    and _grad_lin_ok(s["grad"])):
                ok02 = True
    results.append(("+3：" if ok02 else "+0：")
                   + "新增页背景为浅蓝到白色的线性渐变（浅蓝一端约 #EAF8FF 一类淡蓝，"
                     "白色一端约 #FFFFFF），单一方向线性过渡，非纯色填充、非径向渐变、"
                     "非深色底")

    # +5：新增页插入三个 mp4 视频媒体对象，视频本体随文稿嵌入包内
    # 严格对齐细则，仅踩细则明确写出的点：①新增页上有 3 个媒体对象
    # ②ppt/media 下有 3 个 mp4 本体③[Content_Types].xml 已声明 mp4。
    # 与办公软件对齐：读 p14:media r:embed / a:videoFile r:link 指向的包内部件。
    ok03 = (len(vids) == 3 and len(pkg.mp4_parts()) >= 3 and pkg.mp4_declared())
    results.append(("+5：" if ok03 else "+0：")
                   + "新增页插入三个 mp4 视频媒体对象，视频本体随演示文稿一并嵌入包内"
                     "（ppt/media 下新增3个 mp4，且 [Content_Types].xml 已声明 mp4 "
                     "内容类型）")

    # +5：三段嵌入视频与所给素材字节完全一致
    # 严格对齐细则，仅踩细则明确写出的点：①三段媒体的 md5 命中素材基线
    # ②字节数分别为 528920 / 240828 / 262398③分辨率仍为 360×720 竖版。
    # 与办公软件对齐：读 ppt/media/*.mp4 原始字节 + moov/trak/tkhd 的宽高。
    ok04 = len(vids) == 3
    for v in vids:
        m = v["media"]
        base = MEDIA_BASE.get(m["md5"])
        if base is None or m["len"] != base[1]:
            ok04 = False
        elif m["dims"] is not None and m["dims"] != (360, 720):
            ok04 = False
    results.append(("+5：" if ok04 else "+0：")
                   + "三段嵌入视频与所给素材字节完全一致（分别为 528920 字节、"
                     "240828 字节、262398 字节），未被重新编码、未改变 360×720 竖版"
                     "分辨率、未裁剪画面、未加模糊填充边条")

    # +1：三段视频自左至右依次为“清洁”“定时”“护龈”
    # 严格对齐细则，仅踩细则明确写出的点：按页面横向位置排序后，三段媒体依次命中
    # clean / timer / gumcare 三段不同素材，无重复、无遗漏。
    # 与办公软件对齐：按各视频 a:off 的 x 排序，再逐个比对媒体字节 md5。
    _got = [MEDIA_BASE.get(v["media"]["md5"], (None, 0))[0] for v in vids]
    ok05 = (tuple(_got) == MEDIA_ORDER)
    results.append(("+1：" if ok05 else "+0：")
                   + "三段视频自左至右依次为“清洁”（clean）、“定时”（timer）、"
                     "“护龈”（gumcare）三段不同素材，无重复使用同一段、无遗漏")

    # +1：三段视频均带有取自视频自身画面的封面预览图
    # 严格对齐细则，仅踩细则明确写出的点：①每段视频都有封面预览位图
    # ②该位图为竖版画面且长宽比与 360×720 素材一致（约 1:2）。
    # 与办公软件对齐：读视频对象 blipFill 里 a:blip r:embed 指向的位图并解析像素宽高。
    ok06 = len(vids) == 3
    for v in vids:
        d = v["img"]["dims"] if v["img"] else None
        if not d or d[1] <= d[0] or abs(d[0] / float(d[1]) - 0.5) > 0.05:
            ok06 = False
    results.append(("+1：" if ok06 else "+0：")
                   + "三段视频均带有取自视频自身画面的封面预览图（竖版 360×720 画面），"
                     "非黑屏、非灰底占位图、非无关图片")

    # +5：三段视频分别被套在三个深色近黑、四角圆角的竖版手机外框内
    # 严格对齐细则，仅踩细则明确写出的点：①三个外框各自完整包住一段视频②填充色落在
    # #000000—#1A1A1A 深色区间③圆角几何④竖版（高>宽）⑤三个外框互不复用。
    # 与办公软件对齐：读形状 a:solidFill/a:srgbClr、a:prstGeom/a:custGeom 与 a:off/a:ext。
    ok07 = has3
    results.append(("+5：" if ok07 else "+0：")
                   + "三段视频分别被套在三个深色近黑（填充色落在 #000000—#1A1A1A 一类"
                     "深色区间）、四角圆角的竖版手机外框内，构成“三部手机横向并排展示”"
                     "的样机版式")

    # +5：每部手机带有手机造型细节
    # 严格对齐细则，仅踩细则明确写出的点：①顶部有听筒/前置摄像头开孔条（机身自定义
    # 几何多子路径挖孔，或机身上沿另置横条）②机身侧边至少有一条竖向小圆角按键条；
    # 只用一个光秃圆角矩形代替则不得分。
    # 与办公软件对齐：读机身 a:custGeom/a:pathLst 的子路径条数，以及贴边小形状的
    # a:off/a:ext 与 a:xfrm rot（±90° 按显示效果换算宽高）。
    ok08 = has3
    if ok08:
        for f in frames:
            notch = (f["subpaths"] >= 3) or _top_bar(f, shapes[idx_new])
            if not notch or _side_keys(f, shapes[idx_new]) < 1:
                ok08 = False
    results.append(("+5：" if ok08 else "+0：")
                   + "每部手机带有手机造型细节——顶部听筒/前置摄像头开孔条，以及机身"
                     "侧边的音量键、静音键或电源键等竖向小圆角条，而不是仅用一个光秃的"
                     "圆角矩形代替")

    # +3：三部手机机身高度≥13厘米，页面下半部无高度超过4厘米的整条空白带
    # 严格对齐细则，仅踩细则明确写出的点：①三部机身高度均 ≥13cm
    # ②页面下半部（19.05cm 的下半段）不存在高度 >4cm 的整条无内容横带。
    # 与办公软件对齐：机身 a:ext cy 换算厘米；按各形状纵向占位区间扫描空白横带。
    ok09 = has3
    if ok09:
        for f in frames:
            if (f["rect"][3] - f["rect"][1]) < 13 * EMU_PER_CM:
                ok09 = False
        if _blank_band(shapes[idx_new], page_cy // 2, page_cy,
                       int(4 * EMU_PER_CM), page_area):
            ok09 = False
    results.append(("+3：" if ok09 else "+0：")
                   + "三部手机机身高度≥13厘米（约占页面高度70%以上），三部手机构成页面"
                     "主体，页面下半部无高度超过4厘米的整条空白带")

    # +3：三部手机与三段视频各自尺寸一致、顶端对齐
    # 严格对齐细则，仅踩细则明确写出的点：手机组与视频组内部，宽度互差≤0.3cm、
    # 高度互差≤0.3cm、上边界互差≤0.3cm。
    # 与办公软件对齐：读各形状 a:off y 与 a:ext cx/cy 求极差。
    ok10 = has3
    if ok10:
        for grp in (frames, vids):
            ws = [s["rect"][2] - s["rect"][0] for s in grp]
            hs = [s["rect"][3] - s["rect"][1] for s in grp]
            ts = [s["rect"][1] for s in grp]
            if (_spread(ws) > TOL_03CM or _spread(hs) > TOL_03CM
                    or _spread(ts) > TOL_03CM):
                ok10 = False
    results.append(("+3：" if ok10 else "+0：")
                   + "三部手机与三段视频各自尺寸一致（宽度互差≤0.3厘米、高度互差"
                     "≤0.3厘米）、顶端对齐（上边界互差≤0.3厘米）")

    # +3：三部手机水平方向均匀分布，居中那部落在页面水平中线±1厘米内
    # 严格对齐细则，仅踩细则明确写出的点：①相邻手机水平中心间距互差≤0.8cm
    # ②中间那部手机的水平中心与页面水平中线相差≤1cm。
    # 与办公软件对齐：由 a:off x + a:ext cx/2 求水平中心，与 p:sldSz cx/2 比较。
    ok11 = has3
    if ok11:
        cs = sorted((f["rect"][0] + f["rect"][2]) / 2.0 for f in frames)
        ok11 = (_spread([cs[1] - cs[0], cs[2] - cs[1]]) <= TOL_08CM
                and abs(cs[1] - page_cx / 2.0) <= TOL_1CM)
    results.append(("+3：" if ok11 else "+0：")
                   + "三部手机水平方向均匀分布（相邻手机水平中心间距互差≤0.8厘米），"
                     "居中那部手机的水平中心落在页面水平中线±1厘米内")

    # +5：每段视频完整铺满所在手机的屏幕区域，四边与机身外缘间距均匀
    # 严格对齐细则，仅踩细则明确写出的点：①视频不越出机身②左右上下四个间距互差
    # ≤0.25cm，构成“手机屏幕正在播放”的观感。
    # 与办公软件对齐：视频与机身的 a:off/a:ext 相减得到四边间距。
    ok12 = has3
    if ok12:
        for v, f in zip(vids, frames):
            vr, fr = v["rect"], f["rect"]
            ins = [vr[0] - fr[0], fr[2] - vr[2], vr[1] - fr[1], fr[3] - vr[3]]
            if min(ins) < -TOL_005CM or _spread(ins) > TOL_025CM:
                ok12 = False
    results.append(("+5：" if ok12 else "+0：")
                   + "每段视频完整铺满所在手机的屏幕区域，视频四边与机身外缘的间距均匀"
                     "（各边间距互差≤0.25厘米），构成“手机屏幕正在播放”的观感")

    # +3：视频形状做了圆角处理
    # 严格对齐细则，仅踩细则明确写出的点：三段视频的几何均为圆角矩形一类
    # （roundRect 等），而不是直角矩形 rect。
    # 与办公软件对齐：读视频对象 p:spPr/a:prstGeom 的 prst 值。
    ok13 = (len(vids) == 3 and all(v["prst"] in ROUND_PRESETS for v in vids))
    results.append(("+3：" if ok13 else "+0：")
                   + "视频形状做了圆角处理（视频对象几何为圆角矩形 roundRect，而非"
                     "直角矩形），与手机屏幕圆角贴合")

    # +1：三部手机与三段视频均完整落在页面范围内，彼此互不重叠
    # 严格对齐细则，仅踩细则明确写出的点：①左/上边界≥0，右边界≤33.87cm、
    # 下边界≤19.05cm②三部手机之间、三段视频之间均无实际压盖。
    # 与办公软件对齐：以 p:sldSz 为页框，比对各形状的页面绝对坐标矩形。
    ok14 = has3
    if ok14:
        for s in list(frames) + list(vids):
            r = s["rect"]
            if (r[0] < -TOL_005CM or r[1] < -TOL_005CM
                    or r[2] > page_cx + TOL_005CM or r[3] > page_cy + TOL_005CM):
                ok14 = False
        for a in range(3):
            for b in range(a + 1, 3):
                if (_overlap(frames[a]["rect"], frames[b]["rect"])
                        or _overlap(vids[a]["rect"], vids[b]["rect"])):
                    ok14 = False
    results.append(("+1：" if ok14 else "+0：")
                   + "三部手机与三段视频均完整落在页面范围内（左边界≥0厘米、右边界"
                     "≤33.87厘米、上边界≥0厘米、下边界≤19.05厘米），彼此互不重叠、"
                     "互不压盖")

    # +5：三段视频均设置了进入动画，放映进入该页时尚未显示
    # 严格对齐细则，仅踩细则明确写出的点：三段视频在该页时间轴 mainSeq 中都有进入
    # 效果（presetClass="entr" / p:set 到 visible / animEffect transition="in"）；
    # 三段一开始就静态摊在页面上则不得分。
    # 与办公软件对齐：读该页 p:timing/p:tnLst 里 mainSeq 下的效果节点与 p:spTgt spid。
    ok15 = (len(vids) == 3 and all(v["spid"] in tm_new["entr"] for v in vids))
    results.append(("+5：" if ok15 else "+0：")
                   + "三段视频均设置了进入动画（出现/淡入一类 entr 效果），放映进入该页"
                     "时视频尚未显示，随动画依次登场，而非三段视频一开始就全部静态摊在"
                     "页面上")

    # +5：整组动画由一次单击触发后自动串联完成
    # 严格对齐细则，仅踩细则明确写出的点：①三段视频都有进入动画②该页时间轴里
    # delay="indefinite" 的单击步只有 1 次，其余靠 after/with 自动衔接。
    # 与办公软件对齐：数 mainSeq 下 p:cTn/p:stCondLst/p:cond[@delay="indefinite"] 的个数。
    ok16 = (ok15 and tm_new["clicks"] == 1)
    results.append(("+5：" if ok16 else "+0：")
                   + "整组动画由一次单击触发后自动串联完成——单击一次后第一段视频出现"
                     "并播放，第二段、第三段在前一段播放结束后自动出现并播放，全程无需"
                     "第二次、第三次点击")

    # +3：三段视频均设置为出现后自动播放
    # 严格对齐细则，仅踩细则明确写出的点：三段视频在时间轴里都带媒体播放命令
    # （<p:cmd cmd="playFrom(...)">）或带数值延时的媒体播放节点 p:cMediaNode。
    # 与办公软件对齐：读 mainSeq 下 p:cmd/p:cMediaNode 及其 p:spTgt spid。
    ok17 = (len(vids) == 3 and all(v["spid"] in tm_new["play"] for v in vids))
    results.append(("+3：" if ok17 else "+0：")
                   + "三段视频均设置为出现后自动播放（时间轴中带 playFrom 媒体播放命令"
                     "或媒体播放节点），放映时无需再点击播放按钮或拖动媒体控制条")

    # +5：出现与播放次序严格为第一段→第二段→第三段，相邻触发间隔≥6秒
    # 严格对齐细则，仅踩细则明确写出的点：按自左至右的三段视频，其进入效果触发时刻
    # 依次递增，且相邻两段相差≥6000毫秒（与 6 秒素材时长匹配）。
    # 与办公软件对齐：把效果节点到 tmRoot 一路上各 p:cTn 的 stCondLst 数值延时累加。
    ok18 = ok15
    if ok18:
        ts = [tm_new["entr"][v["spid"]] for v in vids]
        ok18 = all((ts[k + 1] - ts[k]) >= 6000 for k in range(2))
    results.append(("+5：" if ok18 else "+0：")
                   + "三段视频的出现与播放次序严格为第一段→第二段→第三段，后一段的"
                     "触发时间不早于前一段播放结束（相邻两段触发延时间隔≥6秒，与6秒"
                     "素材时长匹配）")

    # +1：新增页三段视频的动画未连带影响其他页面的切换与动画设置
    # 严格对齐细则，仅踩细则明确写出的点：除新增页外的十六个页面既没有被加上切换效果
    # p:transition，也没有被加上任何动画效果节点。
    # 前提是新增页确实做了动画：没有加过任何动画就谈不上“未连带影响”，此时不给分。
    # 与办公软件对齐：逐页读 p:transition 与 p:timing 内的效果节点。
    ok19 = tm_new["anim"] or bool(tm_new["entr"]) or bool(tm_new["play"])
    for i in range(1, n + 1):
        if i == idx_new:
            continue
        if pkg.slide(i).find("p:transition", NS) is not None:
            ok19 = False
        if _timing(pkg, i)["anim"]:
            ok19 = False
    results.append(("+1：" if ok19 else "+0：")
                   + "新增页三段视频的动画未连带影响其他页面的切换与动画设置")

    # +5：第9页配图缩小后完整落在页面内、整体位于版面右侧、与文字无重叠
    # 严格对齐细则，仅踩细则明确写出的点：①上边界≥0、下边界≤19.05cm②左边界≥14.8cm
    # ③与“03”“产品生命周期”标题、说明文字、“ORAL CARE BRAND”字样均无实际压盖。
    # 与办公软件对齐：读该页配图与各文本框的 a:off/a:ext 做矩形交叠判定。
    ok20 = pic9 is not None
    if ok20:
        r = pic9["rect"]
        tg = _shapes_with_text(shapes[idx_p9],
                               ("03", "产品生命周期", "围绕研发、上市、维护、回收",
                                "ORALCAREBRAND"))
        ok20 = (r[1] >= -TOL_005CM and r[3] <= page_cy + TOL_005CM
                and r[0] >= 14.8 * EMU_PER_CM - TOL_005CM and _clear_of(pic9, tg))
    results.append(("+5：" if ok20 else "+0：")
                   + "第9页配图缩小后完整落在页面内（上边界≥0厘米、下边界≤19.05厘米）"
                     "不再向页面下方溢出，且整体位于版面右侧、左边界≥14.8厘米，与“03”、"
                     "“产品生命周期”标题、“围绕研发、上市、维护、回收形成完整产品管理"
                     "链路。”说明文字及“ORAL CARE BRAND”字样均无重叠")

    # +5：第12页配图缩小后右边界≤15.7厘米、落在页面内、与文字卡片无重叠
    # 严格对齐细则，仅踩细则明确写出的点：①右边界≤15.7cm②上边界≥0、下边界≤19.05cm
    # ③与“主力声波刷系列”标题、说明文字及四个卡片文案均无实际压盖。
    # 与办公软件对齐：读该页配图与各文本框的 a:off/a:ext 做矩形交叠判定。
    ok21 = pic12 is not None
    if ok21:
        r = pic12["rect"]
        tg = _shapes_with_text(shapes[idx_p12],
                              ("主力声波刷系列", "面向日常清洁与敏感护理场景",
                               "舒适刷护", "轻量操作", "耐用设计", "内容服务"))
        ok21 = (r[1] >= -TOL_005CM and r[3] <= page_cy + TOL_005CM
                and r[2] <= 15.7 * EMU_PER_CM + TOL_005CM and _clear_of(pic12, tg))
    results.append(("+5：" if ok21 else "+0：")
                   + "第12页配图缩小后右边界≤15.7厘米、上边界≥0厘米、下边界≤19.05厘米，"
                     "与“主力声波刷系列”标题、“面向日常清洁与敏感护理场景，重点不再堆叠"
                     "参数，而是把易用性、舒适度和维护服务打磨到稳定。”说明文字，以及"
                     "“舒适刷护”“轻量操作”“耐用设计”“内容服务”四个卡片完全无重叠")

    # +5：第14页配图缩小后完整落在页面内、整体位于版面右侧、与文字无重叠
    # 严格对齐细则，仅踩细则明确写出的点：①上边界≥0、下边界≤19.05cm②左边界≥14.8cm
    # ③与“05”“产品组合”标题、说明文字、“ORAL CARE BRAND”字样均无实际压盖。
    # 与办公软件对齐：读该页配图与各文本框的 a:off/a:ext 做矩形交叠判定。
    ok22 = pic14 is not None
    if ok22:
        r = pic14["rect"]
        tg = _shapes_with_text(shapes[idx_p14],
                               ("05", "产品组合", "围绕同一刷护习惯",
                                "ORALCAREBRAND"))
        ok22 = (r[1] >= -TOL_005CM and r[3] <= page_cy + TOL_005CM
                and r[0] >= 14.8 * EMU_PER_CM - TOL_005CM and _clear_of(pic14, tg))
    results.append(("+5：" if ok22 else "+0：")
                   + "第14页配图缩小后完整落在页面内（上边界≥0厘米、下边界≤19.05厘米），"
                     "左边界≥14.8厘米，与“05”、“产品组合”标题、“围绕同一刷护习惯，形成"
                     "主机、耗材、护理液与收纳配件组合。”说明文字及“ORAL CARE BRAND”"
                     "字样均无重叠")

    # +5：第9、12、14页三张配图均为等比缩小，缩小后宽高比落在各自区间
    # 严格对齐细则，仅踩细则明确写出的点：三张配图缩小后的宽高比分别落在
    # 1.13—1.16、0.94—0.97、1.10—1.13 区间内（未被横向拉伸或纵向压扁）。
    # 前提是配图尺寸确实被改小过：三张图原始宽高见 PIC_BASE_EXT，没缩小就谈不上
    # “等比缩小”，此时即便比例落在区间内也不给分。
    # 与办公软件对齐：读配图 a:ext 的 cx/cy 求比值，并与原始 a:ext 比大小。
    ok23 = True
    for key, pic in (("p9", pic9), ("p12", pic12), ("p14", pic14)):
        if pic is None:
            ok23 = False
            continue
        r = pic["rect"]
        h = float(r[3] - r[1])
        ratio = ((r[2] - r[0]) / h) if h else 0.0
        lo, hi = RATIO_RANGE[key]
        base_cx, base_cy = PIC_BASE_EXT[key]
        if not (lo <= ratio <= hi):
            ok23 = False
        if (r[2] - r[0]) > base_cx - TOL_01CM or (r[3] - r[1]) > base_cy - TOL_01CM:
            ok23 = False
    results.append(("+5：" if ok23 else "+0：")
                   + "第9、12、14页三张配图均为等比缩小，缩小后宽高比分别落在 "
                     "1.13—1.16 : 1、0.94—0.97 : 1、1.10—1.13 : 1 区间内，画面未被"
                     "二次横向拉伸或纵向压扁")

    # +3：第9、12、14页三张配图与页面上下边界均留有余量
    # 严格对齐细则，仅踩细则明确写出的点：三张配图与页面上边界、下边界的间距均≥0.5cm。
    # 与办公软件对齐：配图 a:off y 与 p:sldSz cy 相减求上下留白。
    ok24 = True
    for pic in (pic9, pic12, pic14):
        if pic is None:
            ok24 = False
            continue
        r = pic["rect"]
        if r[1] < TOL_05CM or (page_cy - r[3]) < TOL_05CM:
            ok24 = False
    results.append(("+3：" if ok24 else "+0：")
                   + "第9、12、14页三张配图缩小后与页面上下边界均留有余量（与页面上"
                     "边界、下边界的间距均≥0.5厘米），不贴边、不顶格，观感整齐")

    # 页码普查：右下角、文字为 1—2 位数字的文本框
    pns = [(i, _page_num_shape(shapes[i])) for i in idx_numbered]
    pn_new = _page_num_shape(shapes[idx_new]) if idx_new else None

    # +3：封面页与结束页右下角均无页码
    # 严格对齐细则，仅踩细则明确写出的点：封面页（“海岚清齿”“声波护理品牌汇报”页）与
    # 结束页（“展望未来”“感谢观看”页）右下角都不存在页码文本框。
    # 与办公软件对齐：按页面自身文案定位这两页，再在右下角区域找纯数字文本框。
    ok25 = (idx_cover is not None and idx_end is not None
            and _page_num_shape(shapes[idx_cover]) is None
            and _page_num_shape(shapes[idx_end]) is None)
    results.append(("+3：" if ok25 else "+0：")
                   + "首尾两页——封面页（“海岚清齿”“声波护理品牌汇报”页）与结束页"
                     "（“展望未来”“感谢观看”页）右下角均无页码")

    # +3：目录页右下角无页码
    # 严格对齐细则，仅踩细则明确写出的点：列出“01 品牌概览”至“06 展望未来”六个条目
    # 的那一页右下角不存在页码文本框。
    # 与办公软件对齐：按页面自身文案定位目录页，再在右下角区域找纯数字文本框。
    ok26 = (idx_toc is not None and _page_num_shape(shapes[idx_toc]) is None)
    results.append(("+3：" if ok26 else "+0：")
                   + "目录页（列出“01 品牌概览”至“06 展望未来”六个条目的页面）右下角"
                     "无页码")

    # +5：五个篇章过渡页右下角均无页码
    # 严格对齐细则，仅踩细则明确写出的点：①识别出的篇章过渡页恰为 5 页
    # ②这 5 页右下角都不存在页码文本框。
    # 与办公软件对齐：以页面右侧“ORAL CARE BRAND”字样识别过渡页，再查右下角数字框。
    ok27 = (len(idx_trans) == 5
            and all(_page_num_shape(shapes[i]) is None for i in idx_trans))
    results.append(("+5：" if ok27 else "+0：")
                   + "五个篇章过渡页（“01 品牌概览”“02 发展路径”“03 产品生命周期”"
                     "“04 成熟期策略”“05 产品组合”页）右下角均无页码")

    # +5：其余九个页面右下角均有页码
    # 严格对齐细则，仅踩细则明确写出的点：①八个正文内容页加新增视频页共 9 页
    # ②这 9 页右下角都有页码文本框。
    # 与办公软件对齐：按页面角色取九页，再在右下角区域找纯数字文本框。
    ok28 = (len(idx_numbered) == 9 and all(s is not None for _i, s in pns))
    results.append(("+5：" if ok28 else "+0：")
                   + "其余九个页面（品牌定位与价值主张、品牌文化与产品原则、阶段化发展"
                     "路径、渠道与用户触点、产品生命周期管理、成熟期核心产品策略、服务"
                     "体验与口碑沉淀、产品组合与场景搭配、新增视频页）右下角均有页码")

    # +5：九个带页码页面按出现先后连续编号为 01—09
    # 严格对齐细则，仅踩细则明确写出的点：按放映顺序取这九页的页码文字，须恰为
    # 01、02、03、04、05、06、07、08、09，不跳号、不重号、不保留旧编号。
    # 与办公软件对齐：读页码文本框内 a:t 文字并按页序排列比对。
    ok29 = ok28 and ([s["text"] for _i, s in pns]
                     == ["0%d" % k for k in range(1, 10)])
    results.append(("+5：" if ok29 else "+0：")
                   + "上述九个带页码页面按出现先后连续编号为 01、02、03、04、05、06、"
                     "07、08、09，不跳号、不重号、不保留与实际序位不符的旧编号")

    # +3：新增视频页页码的位置与对齐方式与其他内容页一致
    # 严格对齐细则，仅踩细则明确写出的点：①段落对齐为右对齐②文本框左上角位于距页面
    # 左边界约31.1厘米、上边界约17.7厘米处（0.25厘米容差）。
    # 与办公软件对齐：读页码文本框 a:off 与段落 a:pPr 的 algn 属性。
    ok30 = pn_new is not None
    if ok30:
        ok30 = (_near(pn_new["rect"][0], PN_OFF[0], TOL_025CM)
                and _near(pn_new["rect"][1], PN_OFF[1], TOL_025CM)
                and pn_new["algn"] == "r")
    results.append(("+3：" if ok30 else "+0：")
                   + "新增视频页右下角页码的位置与对齐方式与其他内容页一致：右对齐，"
                     "位于距页面左边界约31.1厘米、上边界约17.7厘米处")

    # -5：九个带页码页面的页码样式被改动
    # 严格对齐细则，仅踩细则明确写出的点：页码 run 上显式写出的字号≠8磅、显式字体
    # 非微软雅黑、或显式颜色≠#4A6074 才算被改动；未显式写出（随母版继承）不算。
    # 与办公软件对齐：读页码 run 的 a:rPr@sz、a:rPr/a:latin@typeface、a:srgbClr@val。
    bad31 = False
    for _i, s in pns:
        if s is None:
            continue
        for r in s["runs"]:
            if ((r["sz"] is not None and r["sz"] != PN_SZ)
                    or (r["lat"] is not None and r["lat"] != "微软雅黑")
                    or (r["clr"] is not None and r["clr"] != CLR_BODY)):
                bad31 = True
    results.append(("-5：" if bad31 else "-0：")
                   + "九个带页码页面的页码样式被改动，不再是微软雅黑、8磅、颜色 "
                     "#4A6074（如被改为 Arial、6磅、#365C80 一类字体字号字色）")

    # -3：除新增页外的十六个页面，其 #EAF8FF 纯色页面背景被改动
    # 严格对齐细则，仅踩细则明确写出的点：这十六页的页面背景须仍是 #EAF8FF 纯色填充，
    # 被改为渐变、位图或其他颜色即触发。
    # 与办公软件对齐：逐页读 p:cSld/p:bg 的 a:solidFill/a:srgbClr@val。
    bad32 = False
    for i in range(1, n + 1):
        if i == idx_new:
            continue
        kind, val = _bg_of(pkg, i)
        if kind != "solid" or (val or "").upper() != CLR_BG:
            bad32 = True
    results.append(("-3：" if bad32 else "-0：")
                   + "除新增页外的十六个页面，其浅蓝色（#EAF8FF）纯色页面背景被改为"
                     "渐变、位图或其他颜色")

    # -1：新增页以整版位图图片充当底图来模拟渐变
    # 严格对齐细则，仅踩细则明确写出的点：新增页存在满版位图（页面背景位图或覆盖
    # 九成以上页面的图片），且该页既无 p:bg 渐变也无满版渐变形状。
    # 与办公软件对齐：读该页 p:bg 的 a:blip/a:gradFill 与满版 pic 的 a:blip。
    bad33 = False
    if idx_new:
        kind, _v = _bg_of(pkg, idx_new)
        full_bmp = (kind == "blip") or any(
            s["kind"] == "pic" and s["img"] and s["rect"]
            and _area(s["rect"]) >= 0.9 * page_area for s in shapes[idx_new])
        bad33 = full_bmp and not ok02
    results.append(("-1：" if bad33 else "-0：")
                   + "新增页以整版位图图片充当底图来模拟渐变，而非用可继续调整的渐变"
                     "填充（p:bg 的 gradFill 或满版渐变形状），导致背景颜色与方向无法在"
                     "软件内修改")

    # -1：新增页添加了指令未要求的文字
    # 严格对齐细则，仅踩细则明确写出的点：新增页上除右下角页码、左下角“海岚清齿
    # HAILAN CARE”标识以外，还出现了页面大标题、操作提示语或视频旁说明标签等文字。
    # 与办公软件对齐：遍历该页 spTree 内所有 a:t 文字并逐个形状归类。
    bad34 = False
    if idx_new:
        for s in shapes[idx_new]:
            t = _norm(s["text"])
            if not t or (pn_new is not None and s is pn_new):
                continue
            if re.match(r"^\d{1,2}$", t):
                continue
            if "海岚清齿" in t and "HAILANCARE" in t.upper():
                continue
            bad34 = True
    results.append(("-1：" if bad34 else "-0：")
                   + "新增页添加了指令未要求的文字（如“功能视频演示”一类页面大标题、"
                     "“点击后依次出现并播放：清洁 → 定时 → 护龈”一类操作提示语，或"
                     "“清洁模式”“定时提醒”“护龈护理”一类视频旁说明标签）")

    # -3：第9、12、14页以外的其余十张配图被缩放、移位、裁剪或替换
    # 严格对齐细则，仅踩细则明确写出的点：这十张配图的位置与尺寸须逐一命中基线
    # （封面页 14.86×8.36cm 起共十项，容差 0.1cm），张数不符或有一张对不上即触发。
    # 与办公软件对齐：读各页 pic 的 a:off/a:ext，与基线四元组逐一比对。
    kept = []
    for i in range(1, n + 1):
        if i in (idx_p9, idx_p12, idx_p14) or i == idx_new:
            continue
        for s in shapes[i]:
            if s["kind"] == "pic" and s["img"] and s["rect"] and not s["media"]:
                kept.append(s["rect"])
    bad35 = (len(kept) != len(OTHER_IMAGES))
    for r in kept:
        if not any(_near(r[0], ox, TOL_01CM) and _near(r[1], oy, TOL_01CM)
                   and _near(r[2] - r[0], cx, TOL_01CM)
                   and _near(r[3] - r[1], cy, TOL_01CM)
                   for ox, oy, cx, cy in OTHER_IMAGES):
            bad35 = True
    results.append(("-3：" if bad35 else "-0：")
                   + "第9、12、14页以外的其余十张配图（封面页 14.86×8.36厘米、"
                     "“01 品牌概览”页 16.38×9.22厘米、品牌定位与价值主张页 "
                     "12.70×7.15厘米、品牌文化与产品原则页 8.67×4.88厘米、"
                     "“02 发展路径”页 16.38×9.22厘米、渠道与用户触点页 15.24×8.58厘米、"
                     "“04 成熟期策略”页 16.38×9.22厘米、服务体验与口碑沉淀页 "
                     "6.72×3.78厘米、产品组合与场景搭配页 13.95×7.85厘米、结束页 "
                     "13.72×7.72厘米，宽高比均为 1.78 : 1）被缩放、移位、裁剪或替换")

    # -3：封面页、目录页、结束页的文字被改写、删除或新增
    # 严格对齐细则，仅踩细则明确写出的点：这三页必须完整保留细则列出的每一条文案
    # （封面五条、目录六条目及其说明、结束页六条），缺任一条即触发。
    # 与办公软件对齐：拼接该页 spTree 内全部 a:t 文字后做去空白包含判定。
    bad36 = False
    for i, need in ((idx_cover, TEXT_COVER), (idx_toc, TEXT_TOC),
                    (idx_end, TEXT_END)):
        if i is None:
            bad36 = True
            continue
        for k in need:
            if k not in texts[i]:
                bad36 = True
    results.append(("-3：" if bad36 else "-0：")
                   + "封面页“海岚清齿”“声波护理品牌汇报”“智能口腔护理产品策划案”"
                     "“浅蓝视觉版 · 公开传播资料”“品牌运营中心”，目录页六个条目及其"
                     "说明文字，或结束页“三条长期方向”“智能反馈”“专业共创”“绿色材料”"
                     "“感谢观看”“HAILAN CARE · ORAL CARE BRAND”等文字被改写、删除或新增")

    # -3：各页正文与标题的字体字号字色被改动
    # 严格对齐细则，仅踩细则明确写出的点：十六个保留页面（不含新增页、不含页码文本框）
    # 的每个文字 run，其“字号+加粗+颜色+字体”组合须落在基线版式配色内——页面标题
    # 微软雅黑加粗 #163A5F、说明正文微软雅黑 #4A6074、篇章序号微软雅黑加粗 #2F82C9。
    # 与办公软件对齐：读 run 的 a:rPr@sz/@b、a:srgbClr@val、a:latin@typeface
    #（Microsoft YaHei 与微软雅黑视为同一字体）。
    bad37 = False
    for i in range(1, n + 1):
        if i == idx_new:
            continue
        pn = _page_num_shape(shapes[i])
        for s in shapes[i]:
            if pn is not None and s is pn:
                continue
            for r in s["runs"]:
                if (r["sz"], r["b"], r["clr"], r["lat"]) not in STYLE_PALETTE:
                    bad37 = True
    results.append(("-3：" if bad37 else "-0：")
                   + "各页正文与标题的字体字号字色被改动（页面标题为微软雅黑加粗、颜色 "
                     "#163A5F，说明正文为微软雅黑、颜色 #4A6074，篇章序号“01”—“06”为"
                     "微软雅黑加粗、颜色 #2F82C9）")

    # -1：各页左下角“海岚清齿 HAILAN CARE”标识被删除、移位或改变样式
    # 严格对齐细则，仅踩细则明确写出的点：十六个保留页面每页都要有该标识，位置为
    # 距页面左边界约1.65厘米、上边界约17.93厘米（0.1厘米容差），样式为微软雅黑
    # 6.5磅、颜色 #4A6074。
    # 与办公软件对齐：读该文本框 a:off 与 run 的 a:rPr@sz、a:latin、a:srgbClr。
    bad38 = False
    for i in range(1, n + 1):
        if i == idx_new:
            continue
        hit = None
        for s in shapes[i]:
            t = _norm(s["text"]).upper()
            if s["rect"] and "海岚清齿" in _norm(s["text"]) and "HAILANCARE" in t:
                hit = s
                break
        if hit is None:
            bad38 = True
            continue
        if not (_near(hit["rect"][0], 594360, TOL_01CM)
                and _near(hit["rect"][1], 6455664, TOL_01CM)):
            bad38 = True
        for r in hit["runs"]:
            if (r["sz"] != "650" or _font(r["lat"]) != "微软雅黑"
                    or r["clr"] != CLR_BODY):
                bad38 = True
    results.append(("-1：" if bad38 else "-0：")
                   + "各页左下角“海岚清齿  HAILAN CARE”标识文字（微软雅黑6.5磅、颜色 "
                     "#4A6074、位于距页面左边界约1.65厘米、上边界约17.93厘米处）被删除、"
                     "移位或改变样式")

    # -1：五个篇章过渡页右侧“ORAL CARE BRAND”字样被删除或改动
    # 严格对齐细则，仅踩细则明确写出的点：每个过渡页都要有文字恰为 ORAL CARE BRAND
    # 的文本框，且样式为微软雅黑 10磅、颜色 #4A6074。
    # 与办公软件对齐：读该文本框 run 的 a:t 与 a:rPr@sz、a:latin、a:srgbClr。
    bad39 = (len(idx_trans) != 5)
    for i in idx_trans:
        hit = None
        for s in shapes[i]:
            if _norm(s["text"]) == "ORALCAREBRAND":
                hit = s
                break
        if hit is None:
            bad39 = True
            continue
        for r in hit["runs"]:
            if (r["sz"] != "1000" or _font(r["lat"]) != "微软雅黑"
                    or r["clr"] != CLR_BODY):
                bad39 = True
    results.append(("-1：" if bad39 else "-0：")
                   + "五个篇章过渡页右侧“ORAL CARE BRAND”字样（微软雅黑10磅、颜色 "
                     "#4A6074）被删除或改动")

    # -1：页面尺寸不再是 33.87×19.05 厘米
    # 严格对齐细则，仅踩细则明确写出的点：p:sldSz 须为 12192000×6858000 EMU
    # （留 0.1 厘米容差，容纳办公软件回写时的整数取舍）。
    # 与办公软件对齐：读 ppt/presentation.xml 的 p:sldSz@cx/@cy。
    bad40 = not (_near(page_cx, PAGE_CX, TOL_01CM) and _near(page_cy, PAGE_CY, TOL_01CM))
    results.append(("-1：" if bad40 else "-0：")
                   + "页面尺寸不再是 33.87×19.05 厘米（12192000×6858000 EMU）")

    # -1：任意页面的演讲者备注内容被清空或改写
    # 严格对齐细则，仅踩细则明确写出的点：十六个保留页面的备注内容须与基线一致
    # （基线为无备注文字），任一页备注文字发生变化即触发。
    # 与办公软件对齐：顺 slide 的 rels 找 notesSlide 部件并拼接其 a:t 文字。
    bad41 = False
    for i in range(1, n + 1):
        if i == idx_new:
            continue
        if pkg.notes_text(i) != "":
            bad41 = True
    results.append(("-1：" if bad41 else "-0：")
                   + "任意页面的演讲者备注内容被清空或改写")

    # -1：新增页的视频对象被手机机身形状完整压盖（层级过低）
    # 严格对齐细则，仅踩细则明确写出的点：绘制顺序在视频之后、带纯色填充、且几何为
    # 预设形状（非自定义挖孔几何）的形状完整盖住视频，放映时看不到画面；机身用
    # 自定义几何挖出屏幕孔洞的情况画面照常可见，不触发。
    # 与办公软件对齐：按 spTree 内的先后顺序即 z 序，比对 a:off/a:ext 的包含关系。
    bad42 = False
    if idx_new:
        for v in vids:
            for s in shapes[idx_new]:
                if (s["order"] <= v["order"] or not s["rect"] or s["media"]
                        or s["img"] or s["custom"] or s["prst"] is None
                        or not s["fill"]):
                    continue
                if _contains(s["rect"], v["rect"]):
                    bad42 = True
    results.append(("-1：" if bad42 else "-0：")
                   + "新增页的视频对象被手机机身形状完整压盖（层级过低），放映时看不到"
                     "画面")

    # ========== 将 results 反解为结构化 dim2_items ==========
    if len(results) != len(_MAX_DELTAS_ORDER):
        raise RuntimeError("dim2 项数不一致：results=%d vs order=%d"
                           % (len(results), len(_MAX_DELTAS_ORDER)))
    dim2_items = []
    total_score = 0
    for raw, max_delta in zip(results, _MAX_DELTAS_ORDER):
        head, _, rule = raw.partition("：")
        try:
            delta = int(head.replace("+", "").strip())
        except ValueError:
            delta = 0
        hit = delta != 0
        dim2_items.append({
            "rule": rule,
            "max_delta": max_delta,
            "delta": max_delta if hit else 0,
            "hit": hit,
        })
        total_score += (max_delta if hit else 0)

    result["dim2_items"] = dim2_items
    result["total_score"] = total_score
    result["max_score"] = MAX_SCORE_TOTAL
    return result


if __name__ == "__main__":
    # 本地调试入口：evaluate 只接收“脚本所在目录的路径”，由脚本自己在该目录里
    # 定位并打开被评估的 .pptx。命令行第 1 个参数可覆盖目录路径；未指定时使用
    # 脚本自身所在目录。
    if len(sys.argv) >= 2:
        _dir = sys.argv[1]
    else:
        _dir = os.path.dirname(os.path.abspath(__file__))
    _out = evaluate(_dir)
    print(json.dumps(_out, ensure_ascii=False, indent=2))


