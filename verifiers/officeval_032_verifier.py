#!/usr/bin/env python
# -*- coding: utf-8 -*-
# pyright: reportUnknownMemberType=false, reportUnknownVariableType=false, reportUnknownArgumentType=false, reportMissingModuleSource=false, reportAttributeAccessIssue=false, reportOptionalMemberAccess=false, reportIndexIssue=false, reportArgumentType=false, reportCallIssue=false, reportGeneralTypeIssues=false
"""
officeval_032_verifier —「求学申请材料_林绯绯」简历改动评测脚本

说明
----
对外只暴露一个入口：

    evaluate(dir_path) -> dict

dir_path 为存放交付文档的目录；交付物只接受可正常打开的 .pdf——目录内若只有
.docx / .doc / .pptx 等其他格式，仍收作被评对象，但在维度一直接判不通过。

返回值结构：

    {
        "id":          "032",
        "file_name":   "xxx.pdf",
        "status":      "ok" | "error",
        "error":       None | "可读原因" | "脚本运行异常：ExcType: msg",
        "dim1_pass":   True | False | ""，  # 维度1 可用与可修改性；"" = 未进入判定
        "dim1_reason": "……",
        "dim2_items":  [ {rule, max_delta, delta, hit}, … ],
        "total_score": 0,
        "max_score":   34,
    }

维度1 不通过时 dim2_items 为空、total_score 记 0，dim1_reason 带错误摘要；
定位不到交付文件时 dim1_pass 与 dim1_reason 留空、max_score 一并记 0。
本模块不打印任何内容、不改动 sys.stdout、不调用 sys.exit；
仅在 __main__ 下以 JSON 形式输出，便于命令行调试。
"""

import io
import json
import os
import re
import shutil
import sys
import tempfile

# PyMuPDF 已从本项目移除（AGPL 与本项目许可证不兼容），运行环境不安装 fitz。
# 改用同目录的 pymupdf_compat：用 pdfplumber + pypdfium2 实现本脚本用到的
# PyMuPDF 子集。verifier 由 core/worker.py 按文件路径加载（其所在目录已被插入
# sys.path），作为 verifiers 包导入时下面这两行会补齐目录，两种方式都成立。
_VERIFIER_DIR = os.path.dirname(os.path.abspath(__file__))
if _VERIFIER_DIR not in sys.path:
    sys.path.insert(0, _VERIFIER_DIR)
import pymupdf_compat as fitz

SCRIPT_ID = "032"

# ---------------------------------------------------------------------------
# 分值序列：顺序必须与 _check_all() 中 results 的追加顺序严格一致
# ---------------------------------------------------------------------------
_ADD_DELTAS_ORDER = [
    3,  # A01 被横线标注的整句被完整删除，未残留孤立句号或空行
    1,  # A02 该段剩余文字措辞与用字未被改写或缩写
    1,  # A03 该段排版属性保持（9磅/深灰/非加粗/行宽351/2行/行距13.5）
    5,  # A04 新增独立"志愿服务"模块，置于正文流模块之间、间距正常、未重叠溢出
    3,  # A05 标题栏与其他模块同款（深绿横条 355×18 + 橙色竖条 4×12）
    1,  # A06 标题文字"志愿服务"白色加粗 10.5 磅
    3,  # A07 两条内容分成两个独立行、文字无错漏
    1,  # A08 两行内容 9磅/深灰/非加粗
    1,  # A09 两行内容带橙色圆点且左边界与模块条目对齐
    1,  # A10 两行内容之间行距与其他模块条目一致
    1,  # A11 两行内容句末标点完整
    3,  # A12 "技能证书与获奖"末尾另起一行新增条目，未并入上一条
    1,  # A13 新增条目带橙色圆点且与前三条严格左对齐
    1,  # A14 新增条目 9磅/深灰/非加粗且行距与前三条一致
    1,  # A15 新增条目以句号收尾
    5,  # A16 已添加新增内容且仍为 1 页 A4 纵向
    1,  # A17 页脚带与页脚文字仍在页面下沿原位置
    1,  # A18 "调整。"与页脚带之间仍留可见空白
]

_DED_DELTAS_ORDER = [
    -1,  # D01 大标题"求学申请材料" 24磅/#173C3B/加粗/左对齐右栏
    -1,  # D02 副标题"面向能源信息与智能控制方向" 11磅/#6B7474/非加粗
    -1,  # D03 标题下金色分隔横线 #D4A557 宽约352磅 位于约75磅处
    -1,  # D04 "报考与初试情况"四组标签/取值格式与左右两列对齐
    -1,  # D05 初试成绩表 5列2行/单元格71×24/表头底色/框线/字号
    -1,  # D06 "教育背景"年月·院校·专业行 深绿加粗 左中右
    -1,  # D07 "课程结构以算法设计…"段落 9.1磅/深灰/3行/行宽355/行距14
    -1,  # D08 "实践经历"年月·公司·岗位行 深绿加粗 左中右
    -1,  # D09 "课程项目""楼宇能耗预测与调度小组项目" 深绿加粗
    -1,  # D10 "围绕温湿度…"段落 9磅/深灰/2行/行宽351
    -1,  # D11 "课程项目"行上下垂直间距未出现紧贴、重叠或大片空洞
    -1,  # D12 "技能证书与获奖"原有三条 9磅/深灰/橙点/行距16.5/缩进12
    -1,  # D13 "研修规划"段落 9磅/深灰/首行缩进18/5行/行距14
    -1,  # D14 五个模块标题栏 深绿355×18 + 白色加粗10.5 + 橙色竖条4×12
    -1,  # D15 页脚文字 7.6磅/#6B7474/非加粗/右对齐/位于页脚带内
    -1,  # D16 左侧深绿通栏 175磅满高 + 最左深色窄条 18磅
    -1,  # D17 圆形头像 圆形裁切/直径84磅/水平居中于左栏
    -1,  # D18 "林绯绯"/"本科 · 计算机与自动化技术"/"跨学科申请"胶囊
    -1,  # D19 左栏四个分组标题 白色加粗10.2 位于120×24 #3F7774 圆角条内
    -1,  # D20 左栏五行标签 8.8磅浅绿 / 取值 8.8磅白色 / 两列对齐
    -1,  # D21 "联系方式"三行 8.5磅/白色/非加粗/行距17
    -1,  # D22 "能力概览"五条 8.6磅/白色/橙点4.4/行距18
    -1,  # D23 "申请信息"三行标签 8.6磅浅绿 / 取值 8.6磅白色
    -1,  # D24 左栏底部浅色装饰圆弧 #C8DDD6 未被删除移位改色
    -3,  # D25 删除或改写指令未指定改动的文字内容
    -3,  # D26 删除或破坏头像、成绩表、模块标题栏、圆点、装饰弧形等既有对象
    -3,  # D27 删除任一既有模块或段落
    -1,  # D28 改动姓名/出生年月/成绩/院校/专业/起止年月等具体数值
    -1,  # D29 在指令未要求处新增多余内容
]

MAX_SCORE_TOTAL = sum(_ADD_DELTAS_ORDER)
_MIN_TOTAL = MAX_SCORE_TOTAL + sum(_DED_DELTAS_ORDER)
assert MAX_SCORE_TOTAL == 34, MAX_SCORE_TOTAL
assert sum(_DED_DELTAS_ORDER) == -35, sum(_DED_DELTAS_ORDER)
assert len(_ADD_DELTAS_ORDER) == 18
assert len(_DED_DELTAS_ORDER) == 29

_MAX_DELTAS_ORDER = _ADD_DELTAS_ORDER + _DED_DELTAS_ORDER
_ADD_COUNT = len(_ADD_DELTAS_ORDER)

# ---------------------------------------------------------------------------
# 版面基准常量（单位：磅；颜色为 0xRRGGBB）
# ---------------------------------------------------------------------------
C_ACCENT = 0x173C3B      # 右栏标题/小标题深绿
C_BODY = 0x2F3333        # 正文深灰
C_BAR = 0x235D5B         # 模块标题栏 / 左侧通栏深绿
C_STRIP = 0x163B3B       # 左侧最窄深色条
C_ORANGE = 0xD4A557      # 橙色（竖条、圆点、金线、胶囊）
C_WHITE = 0xFFFFFF
C_MUTED = 0x6B7474       # 副标题 / 页脚文字灰
C_SUBTITLE_L = 0xD8E7E4  # 左栏"本科 · …"浅绿
C_LABEL_L = 0xCFE1DE     # 左栏标签浅绿
C_GROUP_BAR = 0x3F7774   # 左栏分组标题条
C_TBL_HEAD = 0xDDEBE6    # 成绩表表头底色
C_TBL_LINE = 0xC3D1CD    # 成绩表框线
C_FOOTER = 0xE8EFE9      # 页脚带
C_ARC = 0xC8DDD6         # 左下装饰弧

RIGHT_X = 205.0          # 右栏左边缘
RIGHT_W = 355.3          # 右栏栏宽 / 模块标题栏宽
COL_SPLIT = 200.0        # 左右栏切分线
A4_W, A4_H = 595.28, 841.89

BAND_X0 = 175.0          # 页脚色带左端（贴左栏右边缘）
BAND_H = 22.0            # 页脚色带高
ARC_D = 230.0            # 左下装饰弧直径

_SAMPLE_DPI = 150
_SIZE_TOL = 0.12         # 字号相对容差（吸收 docx↔pdf 转换的全局缩放）
_TOL_VEC = 18            # 矢量填充色比较容差
_TOL_IMG = 40            # 位图取样色比较容差（抗锯齿）

_BOLD_RE = re.compile(r"bold|black|heavy|semibold|demi|\bbd\b", re.I)

_PUNCT_MAP = {
    "，": ",", "；": ";", "：": ":", "、": ",", "。": ".",
    "（": "(", "）": ")", "【": "[", "】": "]", "％": "%",
    "－": "-", "—": "-", "–": "-", "﹣": "-", "～": "~",
    "！": "!", "？": "?", "／": "/", "＋": "+",
    " ": "", "　": "", "​": "",
}

_END_PUNCT = ("。", "；", ".", ";")
_END_PERIOD = ("。", ".")

# 指令未要求改动、必须原样保留的内容锚点（用于 D25 / D27）
_ANCHORS_RIGHT = [
    "求学申请材料", "面向能源信息与智能控制方向", "报考与初试情况",
    "报考院校", "南港理工学院", "报考方向", "新能源控制工程",
    "培养类型", "专业型硕士", "当前学历", "本科在读",
    "思政", "外语", "专业基础", "数学", "总分",
    "教育背景", "2022.09-2026.06", "东序信息学院", "计算机与自动化技术(本科)",
    "课程结构以算法设计", "本化处理与版本协作",
    "实践经历", "2025.02-2025.07", "青封智能装配有限公司", "智能监测开发助理",
    "参与传感数据采集", "和模型推理接口联调",
    "课程项目", "楼宇能耗预测与调度小组项目",
    "围绕温湿度", "项目说明书与演示看板",
    "技能证书与获奖", "英语能力:通过大学英语四级", "工具能力:熟悉Git",
    "综合表现:曾获校级学业进步奖",
    "研修规划", "进入研究生阶段后", "后续方向将根据导师课题与团队安排动态调整",
]
_ANCHORS_LEFT = [
    "林绯绯", "本科·计算机与自动化技术", "跨学科申请",
    "基本信息", "出生年月", "2001.11", "政治面貌", "共青团员",
    "民族", "汉族", "最高学历", "本科学历", "毕业院校", "东序信息学院",
    "联系方式", "校内申请系统转接", "联系邮箱:申请系统转接", "常驻城市:江南片区",
    "能力概览", "Python/C++/C#", "数据处理与可视化", "边缘计算部署",
    "控制系统仿真", "技术文档写作",
    "申请信息", "报考院校", "报考专业", "报考属性",
]
_MODULE_TITLES = ["报考与初试情况", "教育背景", "实践经历", "技能证书与获奖", "研修规划"]
_GROUP_TITLES = ["基本信息", "联系方式", "能力概览", "申请信息"]
_DELETED_SENTENCE = ["项目内容聚焦设备运行状态识别与远程诊断", "未保留原始业务数据"]

# ---------------------------------------------------------------------------
# 通用小工具
# ---------------------------------------------------------------------------


def _norm(text):
    """归一化：去空白、全角标点折叠为半角，用于内容比对。"""
    out = []
    for ch in text or "":
        if ch.isspace() or ord(ch) in (0x200B, 0x200E, 0x200F, 0xFEFF, 0xA0, 0x3000):
            continue
        out.append(_PUNCT_MAP.get(ch, ch))
    return "".join(out)


def _fill_geom(mark):
    mark["w"] = mark["x1"] - mark["x0"]
    mark["h"] = mark["y1"] - mark["y0"]
    mark["cx"] = (mark["x0"] + mark["x1"]) / 2.0
    mark["cy"] = (mark["y0"] + mark["y1"]) / 2.0
    return mark


def _rgb(value):
    return ((value >> 16) & 0xFF, (value >> 8) & 0xFF, value & 0xFF)


def _cdiff(a, b):
    """两个 (r,g,b) 的最大通道差。"""
    return max(abs(a[0] - b[0]), abs(a[1] - b[1]), abs(a[2] - b[2]))


def _color_ok(got, ref, tol=_TOL_VEC):
    """got 为 int 或 (r,g,b)；ref 为 int 常量。"""
    if got is None:
        return False
    g = _rgb(got) if isinstance(got, int) else tuple(got)
    return _cdiff(g, _rgb(ref)) <= tol


def _size_ok(got, ref, tol=_SIZE_TOL):
    if not got:
        return False
    return abs(float(got) - ref) <= max(ref * tol, 0.35)


def _near(a, b, tol):
    return abs(a - b) <= tol


def _median(values):
    vals = sorted(v for v in values if v is not None)
    if not vals:
        return None
    mid = len(vals) // 2
    if len(vals) % 2:
        return vals[mid]
    return (vals[mid - 1] + vals[mid]) / 2.0


def _all(*flags):
    return all(bool(f) for f in flags)


# ---------------------------------------------------------------------------
# 交付文件定位与格式门禁
# ---------------------------------------------------------------------------
# 交付物只接受 .pdf：细则要求交的是可正常打开的 PDF，其余版式/源文件格式
# （.docx / .doc / .pptx / .odt 等）一律在维度一判不通过。
_ACCEPTED_EXT = ".pdf"
_PDF_EXT = (".pdf",)
# 目录里没有 .pdf 时，这些扩展名仍收作被评对象，交给维度一判不通过（完成度记 0），
# 而不是当成「没有交付物」直接跳过评测。
_WRONG_FMT_EXTS = (".docx", ".doc", ".docm", ".dotx", ".dot", ".rtf", ".odt", ".wps",
                   ".pages", ".pptx", ".ppt", ".pptm", ".ppsx", ".potx", ".odp", ".key")


def _list_candidates(dir_path, exts):
    out = []
    for name in sorted(os.listdir(dir_path)):
        if name.startswith("~$") or name.startswith("."):
            continue
        if name.lower().endswith(exts):
            out.append(os.path.join(dir_path, name))
    return out


def _fmt_reject(file_path):
    """维度1的格式门禁：交付物必须是可正常打开的 .pdf。

    不通过返回原因文本，通过返回 None。判定与阅读器的「能不能双击打开」对齐：
    扩展名须为 .pdf，文件头须含 %PDF- 标识，且能被 PDF 解析后端打开、至少有 1 页、
    不需要口令。
    """
    ext = os.path.splitext(file_path)[1].lower()
    if ext != _ACCEPTED_EXT:
        return "交付物为 %s 格式，只接受 .pdf 文档" % (ext or "无扩展名")
    try:
        with open(file_path, "rb") as fh:
            head = fh.read(1024)
    except OSError as exc:
        return "交付物无法读取：%s: %s" % (type(exc).__name__, exc)
    if b"%PDF-" not in head:
        return "交付物不是有效的 PDF（文件头缺少 %PDF- 标识）"
    try:
        doc = fitz.open(file_path)
    except Exception as exc:
        return "交付物无法作为 PDF 打开：%s: %s" % (type(exc).__name__, exc)
    try:
        if doc.needs_pass:
            return "交付物 PDF 已加密，需要口令才能打开"
        if doc.page_count < 1:
            return "交付物 PDF 内没有任何页面"
    finally:
        doc.close()
    return None


def _resolve_pdf(dir_path, work_dir):
    """返回 (用于展示的交付文件名, 实际用于解析的 PDF 路径)。

    交付物只接受 .pdf：目录内优先取 .pdf；没有 .pdf 时把 .docx/.pptx 等其他格式
    原样收作被评对象（不再转成 PDF 后放行），由 _fmt_reject 在维度一判不通过。
    work_dir 保留在签名里，供放宽门禁、需要转档时复用。
    """
    pdfs = _list_candidates(dir_path, _PDF_EXT)
    if pdfs:
        return os.path.basename(pdfs[0]), pdfs[0]
    others = _list_candidates(dir_path, _WRONG_FMT_EXTS)
    if others:
        return os.path.basename(others[0]), others[0]
    raise FileNotFoundError(f"目录中未找到 .pdf 文件：{dir_path}")


# ---------------------------------------------------------------------------
# 页面模型：文本行 + 图形元素（矢量与位图统一看待）
# ---------------------------------------------------------------------------
class _Doc(object):
    """把一页 PDF 拍平成"文本行列表 + 图形块列表"，供细则逐条比对。

    好文件与坏文件出自不同渲染管线：一方把圆点/竖条/细线画成矢量，另一方
    压成小位图。因此图形块统一由 get_drawings() 与 get_image_info() 汇总，
    颜色对矢量取 fill/stroke、对位图取渲染像素众数，并把同色相邻碎块合并
    （某些管线会把一根圆角横条拆成左中右三段）。
    """

    def __init__(self, pdf_path):
        self.doc = fitz.open(pdf_path)
        self.n_pages = self.doc.page_count
        self.page = self.doc[0]
        self.width = self.page.rect.width
        self.height = self.page.rect.height
        pix = self.page.get_pixmap(dpi=_SAMPLE_DPI)
        self._pw, self._ph, self._n = pix.width, pix.height, pix.n
        self._buf = pix.samples
        self._scale = _SAMPLE_DPI / 72.0
        self.bg = self._page_bg()
        self.lines = self._collect_lines(self.page)
        self.right = [ln for ln in self.lines if ln["x0"] >= COL_SPLIT]
        self.left = [ln for ln in self.lines if ln["x0"] < COL_SPLIT]
        self.flat_right = "".join(ln["ntext"] for ln in self.right)
        self.flat_left = "".join(ln["ntext"] for ln in self.left)
        self.flat_all = self.flat_right + self.flat_left
        self.full_text = "".join(
            _norm(self.doc[i].get_text("text")) for i in range(self.n_pages))
        self.images = self._collect_images(self.page)
        self.marks = self._collect_marks(self.page)

    def close(self):
        try:
            self.doc.close()
        except Exception:
            pass

    # -- 像素取样 ---------------------------------------------------------
    def px(self, x, y):
        ix = min(max(int(x * self._scale), 0), self._pw - 1)
        iy = min(max(int(y * self._scale), 0), self._ph - 1)
        off = (iy * self._pw + ix) * self._n
        return (self._buf[off], self._buf[off + 1], self._buf[off + 2])

    def _page_bg(self):
        """全页粗网格取样的众数色，即纸张底色。"""
        tally = {}
        for i in range(24):
            for j in range(24):
                c = self.px(self.width * (i + 0.5) / 24.0, self.height * (j + 0.5) / 24.0)
                slot = tally.setdefault((c[0] // 16, c[1] // 16, c[2] // 16), [0, c])
                slot[0] += 1
        return max(tally.values(), key=lambda s: s[0])[1]

    def _dominant(self, x0, y0, x1, y1):
        """块内向内收缩后密集取样的众数色；忽略与纸张底色接近的取样点。

        小尺寸装饰件（圆点、竖条）边缘抗锯齿严重，直接取中心或不加过滤的
        众数都会把纸张底色误当成对象颜色，故此处先剔除底色再取众数。
        """
        w, h = x1 - x0, y1 - y0
        fx0 = x0 + (w * 0.18 if w >= 1.5 else w * 0.5)
        fx1 = x1 - (w * 0.18 if w >= 1.5 else w * 0.5)
        fy0 = y0 + (h * 0.18 if h >= 1.5 else h * 0.5)
        fy1 = y1 - (h * 0.18 if h >= 1.5 else h * 0.5)
        tally = {}
        fallback = {}
        for i in range(9):
            for j in range(9):
                x = fx0 + (fx1 - fx0) * (i / 8.0)
                y = fy0 + (fy1 - fy0) * (j / 8.0)
                c = self.px(x, y)
                key = (c[0] // 16, c[1] // 16, c[2] // 16)
                slot = fallback.setdefault(key, [0, c])
                slot[0] += 1
                if _cdiff(c, self.bg) <= 14:
                    continue
                slot2 = tally.setdefault(key, [0, c])
                slot2[0] += 1
        pool = tally if tally else fallback
        return max(pool.values(), key=lambda s: (s[0], _cdiff(s[1], self.bg)))[1]

    # -- 文本行 -----------------------------------------------------------
    @staticmethod
    def _collect_lines(page):
        raw = []
        for blk in page.get_text("dict").get("blocks", []):
            if blk.get("type") != 0:
                continue
            for ln in blk.get("lines", []):
                spans = [sp for sp in ln.get("spans", []) if (sp.get("text") or "").strip()]
                if not spans:
                    continue
                text = "".join(sp["text"] for sp in spans)
                bbox = ln["bbox"]
                raw.append({
                    "text": text.strip(),
                    "ntext": _norm(text),
                    "x0": bbox[0], "y0": bbox[1], "x1": bbox[2], "y1": bbox[3],
                    "size": max(sp.get("size", 0.0) for sp in spans),
                    "color": spans[0].get("color", 0),
                    "font": spans[0].get("font", "") or "",
                    "bold": any(_BOLD_RE.search(sp.get("font", "") or "") for sp in spans),
                })
        raw.sort(key=lambda ln: (round(ln["y0"], 1), ln["x0"]))
        # 部分管线用"同字重叠描两遍"模拟加粗：去重并据此判定 bold
        out = []
        for ln in raw:
            dup = None
            for kept in out:
                if (kept["ntext"] == ln["ntext"]
                        and _near(kept["y0"], ln["y0"], 1.5)
                        and _near(kept["x0"], ln["x0"], 1.5)):
                    dup = kept
                    break
            if dup is not None:
                dup["bold"] = True
                continue
            out.append(ln)
        return out

    @staticmethod
    def _collect_images(page):
        out = []
        try:
            infos = page.get_image_info(xrefs=True)
        except Exception:
            infos = []
        for im in infos:
            b = im.get("bbox")
            if not b:
                continue
            out.append({"x0": b[0], "y0": b[1], "x1": b[2], "y1": b[3],
                        "w": b[2] - b[0], "h": b[3] - b[1]})
        return out

    # -- 图形块 -----------------------------------------------------------
    def _collect_marks(self, page):
        raw = []
        for g in page.get_drawings():
            r = g["rect"]
            fill = g.get("fill")
            stroke = g.get("color")
            base = fill if fill else stroke
            if base is None:
                continue
            color = tuple(int(round(c * 255)) for c in base)
            raw.append({
                "kind": "V", "x0": r.x0, "y0": r.y0, "x1": r.x1, "y1": r.y1,
                "color": color, "tol": _TOL_VEC,
                "stroke": (tuple(int(round(c * 255)) for c in stroke) if stroke else None),
            })
        for im in self.images:
            raw.append({
                "kind": "I", "x0": im["x0"], "y0": im["y0"], "x1": im["x1"], "y1": im["y1"],
                "color": self._dominant(im["x0"], im["y0"], im["x1"], im["y1"]),
                "tol": _TOL_IMG, "stroke": None,
            })
        for m in raw:
            _fill_geom(m)
            m["strokes"] = [m["stroke"]] if m["stroke"] else []
            m["parts"] = 1
        self.raw = sorted(raw, key=lambda m: (round(m["y0"], 1), m["x0"]))
        merged = self._merge(raw, self.width * self.height * 0.5)
        for m in merged:
            _fill_geom(m)
        return merged

    @staticmethod
    def _merge(marks, area_cap):
        """同色且相邻/相交的碎块合并为一个整体（横条被拆成三段时尤为必要）。

        整页底色块之类的巨幅色块不参与合并，否则会把其上的浅色小块吞掉。
        """
        n = len(marks)
        parent = list(range(n))
        big = [((m["x1"] - m["x0"]) * (m["y1"] - m["y0"])) >= area_cap for m in marks]

        def find(i):
            while parent[i] != i:
                parent[i] = parent[parent[i]]
                i = parent[i]
            return i

        for i in range(n):
            a = marks[i]
            if big[i]:
                continue
            for j in range(i + 1, n):
                b = marks[j]
                if big[j]:
                    continue
                if _cdiff(a["color"], b["color"]) > max(a["tol"], b["tol"]):
                    continue
                if a["x0"] - 2.0 > b["x1"] or b["x0"] - 2.0 > a["x1"]:
                    continue
                if a["y0"] - 2.0 > b["y1"] or b["y0"] - 2.0 > a["y1"]:
                    continue
                ra, rb = find(i), find(j)
                if ra != rb:
                    parent[rb] = ra
        groups = {}
        for i, m in enumerate(marks):
            groups.setdefault(find(i), []).append(m)
        out = []
        for members in groups.values():
            head = max(members, key=lambda m: (m["x1"] - m["x0"]) * (m["y1"] - m["y0"]) + 1e-6)
            out.append({
                "kind": head["kind"], "tol": head["tol"],
                "color": head["color"], "stroke": head["stroke"],
                "x0": min(m["x0"] for m in members), "y0": min(m["y0"] for m in members),
                "x1": max(m["x1"] for m in members), "y1": max(m["y1"] for m in members),
                "strokes": [m["stroke"] for m in members if m["stroke"]],
                "parts": len(members),
            })
        out.sort(key=lambda m: (round(m["y0"], 1), m["x0"]))
        return out

    # -- 查询辅助 ---------------------------------------------------------
    def find(self, needle, side=None, y_range=None):
        got = self.finds(needle, side, y_range)
        return got[0] if got else None

    def finds(self, needle, side=None, y_range=None):
        key = _norm(needle)
        pool = self.lines
        if side == "R":
            pool = self.right
        elif side == "L":
            pool = self.left
        out = []
        for ln in pool:
            if key not in ln["ntext"]:
                continue
            if y_range and not (y_range[0] <= ln["y0"] <= y_range[1]):
                continue
            out.append(ln)
        return out

    def exact(self, text, side=None):
        key = _norm(text)
        pool = self.right if side == "R" else (self.left if side == "L" else self.lines)
        for ln in pool:
            if ln["ntext"] == key:
                return ln
        return None

    def has(self, needle):
        return _norm(needle) in self.flat_all or _norm(needle) in self.full_text

    def marks_of(self, ref, wrange=None, hrange=None, xrange=None, yrange=None, pool="merged"):
        out = []
        for m in (self.raw if pool == "raw" else self.marks):
            if not _color_ok(m["color"], ref, m["tol"]):
                continue
            if wrange and not (wrange[0] <= m["w"] <= wrange[1]):
                continue
            if hrange and not (hrange[0] <= m["h"] <= hrange[1]):
                continue
            if xrange and not (xrange[0] <= m["x0"] <= xrange[1]):
                continue
            if yrange and not (yrange[0] <= m["y0"] <= yrange[1]):
                continue
            out.append(m)
        return out

    def module_bars(self):
        """右栏模块标题横条（深绿、宽约 355、高约 18）。"""
        return self.marks_of(C_BAR, wrange=(330.0, 380.0), hrange=(14.0, 23.0),
                             xrange=(196.0, 214.0))

    def ticks(self):
        """模块标题栏左端的橙色竖条（约 4×12）。"""
        return self.marks_of(C_ORANGE, wrange=(2.5, 6.5), hrange=(8.0, 16.0),
                             xrange=(204.0, 224.0))

    def dots_right(self):
        """右栏条目的橙色圆点（约 4.6×4.6）；排除标题栏内竖条被拆出的小碎片。"""
        cand = self.marks_of(C_ORANGE, wrange=(3.0, 6.0), hrange=(3.0, 6.0),
                             xrange=(198.0, 218.0), pool="raw")
        bars = self.module_bars()
        return [m for m in cand
                if not any(b["y0"] - 1.0 <= m["cy"] <= b["y1"] + 1.0
                           and b["x0"] - 1.0 <= m["cx"] <= b["x1"] + 1.0 for b in bars)]

    def dots_left(self):
        return self.marks_of(C_ORANGE, wrange=(3.0, 6.0), hrange=(3.0, 6.0),
                             xrange=(28.0, 46.0), pool="raw")

    def footer_band(self):
        bands = self.marks_of(C_FOOTER, wrange=(300.0, 460.0), hrange=(15.0, 30.0),
                              yrange=(self.height - 60.0, self.height), pool="raw")
        return bands[0] if bands else None

    def dot_for(self, line, dots, tol=6.0):
        cy = (line["y0"] + line["y1"]) / 2.0
        for d in dots:
            if abs(d["cy"] - cy) <= tol and d["x1"] <= line["x0"] + 3.0:
                return d
        return None

    def bar_of_title(self, line):
        for m in self.module_bars():
            if (m["y0"] - 4.0 <= line["y0"] and line["y1"] <= m["y1"] + 6.0
                    and m["x0"] - 4.0 <= line["x0"] and line["x1"] <= m["x1"] + 4.0):
                return m
        return None

    def gap_above(self, bar):
        """横条顶边与其上方最近的右栏正文底边之间的净空白。"""
        best = None
        for ln in self.right:
            if ln["y1"] <= bar["y0"] + 0.5 and (best is None or ln["y1"] > best):
                best = ln["y1"]
        if best is None:
            return bar["y0"]
        return bar["y0"] - best


# ---------------------------------------------------------------------------
# 维度1：可用与可修改性
# ---------------------------------------------------------------------------


def _dim1(d):
    # 格式门禁（必须是可正常打开的 .pdf）已由 _fmt_reject 在 _evaluate_impl 里先判过
    if d.n_pages < 1 or not d.lines:
        return False, "文档无有效页面或页面内无可选取文字"
    if len(d.flat_all) < 350:
        return False, "页面可选取文字过少，疑为整页位图截图或扫描件"
    for im in d.images:
        if im["w"] * im["h"] >= d.width * d.height * 0.85:
            return False, "页面主体为整页位图，正文不可编辑"
    if any(ln["x1"] > COL_SPLIT - 4.0 for ln in d.left):
        return False, "左栏文字越界压入右栏，左右两栏互相压盖"
    if any(ln["x1"] > d.width + 1.0 or ln["y1"] > d.height + 1.0 for ln in d.lines):
        return False, "存在超出纸张边界、被裁切丢失的文字"
    lines = d.lines
    clashes = 0
    for i, a in enumerate(lines):
        for b in lines[i + 1:]:
            if b["y0"] >= a["y1"] - 0.5:
                break
            ov_y = min(a["y1"], b["y1"]) - max(a["y0"], b["y0"])
            ov_x = min(a["x1"], b["x1"]) - max(a["x0"], b["x0"])
            hmin = min(a["y1"] - a["y0"], b["y1"] - b["y0"]) or 1.0
            wmin = min(a["x1"] - a["x0"], b["x1"] - b["x0"]) or 1.0
            if ov_y > hmin * 0.5 and ov_x > wmin * 0.4:
                clashes += 1
    if clashes > 8:
        return False, f"版面严重错乱：检出 {clashes} 处文字大面积重叠"
    return True, "文件可正常打开，正文为可选取文字与矢量/图形对象，版面未见严重错乱"


# ---------------------------------------------------------------------------
# 维度2：逐条细则
# ---------------------------------------------------------------------------
_VOL1 = "2022年1-2月在社区做抗击疫情先进志愿者"
_VOL2 = "2023年7-9月在山底村做志愿者"
_SKILL_NEW = "2021年11月-2023年9月参加青年马克思主义者培养工程，培训合格"
_KEEP_PARA = ("参与传感数据采集、状态识别与可视化看板搭建，"
              "负责部分边缘网关脚本、异常日志清洗和模型推理接口联调。")


def _fmt_ok(line, size=None, color=None, bold=None):
    if line is None:
        return False
    if size is not None and not _size_ok(line["size"], size):
        return False
    if color is not None and not _color_ok(line["color"], color):
        return False
    if bold is not None and bool(line["bold"]) != bool(bold):
        return False
    return True


def _consecutive(pool, lines):
    try:
        idx = [pool.index(ln) for ln in lines]
    except ValueError:
        return False
    return all(idx[i + 1] == idx[i] + 1 for i in range(len(idx) - 1))


def _ends_with(line, endings):
    return bool(line) and line["text"].rstrip().endswith(endings)


def _has_color_inside(d, mark, ref, tol):
    for fx in (0.15, 0.35, 0.5, 0.65, 0.85):
        for fy in (0.15, 0.35, 0.5, 0.65, 0.85):
            if _color_ok(d.px(mark["x0"] + mark["w"] * fx, mark["y0"] + mark["h"] * fy),
                         ref, tol):
                return True
    return False


def _band_intact(d, band):
    """页脚色带是否仍是原样：左端贴左栏右边缘、右端贯通到页面右边缘、高约22磅、位于页面下沿。

    footer_band() 的筛选窗口很宽（宽 300~460、高 15~30），只够"找到"色带；
    尺寸与位置是否被改动要在这里判。
    """
    if band is None:
        return False
    return (_near(band["x0"], BAND_X0, 12.0)
            and band["x1"] >= d.width - 6.0
            and _near(band["h"], BAND_H, 5.0)
            and band["y0"] >= d.height - 60.0)


def _find_arc(d):
    """左下装饰弧：贴页面左下角、尺寸在弧形量级内（原稿 230×230，转换后常被页面裁去大半）。

    候选先按"填充色即弧色"命中，再退回像素取样。整页底色块与左栏通栏都够大且
    位置也贴着左下角，若不设尺寸上限，它们会抢在真弧之前命中，把"弧形已被删除"
    掩盖成通过。
    """
    cand = []
    for m in d.raw:
        if m["x0"] > 78.0 or m["y1"] < d.height - 95.0:
            continue
        if not (40.0 <= m["w"] <= ARC_D * 1.15 and 40.0 <= m["h"] <= ARC_D * 1.15):
            continue
        cand.append(m)
    for m in cand:
        if _color_ok(m["color"], C_ARC, max(m["tol"], 30)):
            return m
    for m in cand:
        if _has_color_inside(d, m, C_ARC, 30):
            return m
    return None


def _find_avatar(d):
    for im in d.images:
        if im["x0"] >= COL_SPLIT:
            continue
        if not (74.0 <= im["w"] <= 95.0 and 74.0 <= im["h"] <= 95.0):
            continue
        if abs(im["w"] - im["h"]) > 4.0:
            continue
        return im
    return None


def _add(results, ok, delta, rule):
    results.append(("+%d：%s" % (delta, rule)) if ok else ("+0：%s" % rule))


def _ded(results, hit, delta, rule):
    results.append(("-%d：%s" % (delta, rule)) if hit else ("-0：%s" % rule))


def _check_additions(d, results, st):
    """一、加分项（三项指令任务 + 篇幅控制）。"""
    # ---- 【删除被横线标注的语句】 ------------------------------------
    p1 = d.find("参与传感数据采集", side="R")
    p2 = d.find("和模型推理接口联调", side="R")
    orphan = any(ln["ntext"] in (".", ",", ";", "。") for ln in d.right)
    a01 = (not any(d.has(s) for s in _DELETED_SENTENCE)) and not orphan
    _add(results, a01, 3, '"实践经历"首段中被横线标注的整句被完整删除，未残留孤立句号或多余空行')

    # 后两条以"整句已删除"为前提，未删除者不得凭原样未动获分
    a02 = a01 and _norm(_KEEP_PARA) in d.flat_right
    _add(results, a02, 1, "该段删除后剩余文字措辞与用字未被改写或缩写")

    a03 = False
    if a01 and a02 and p1 and p2:
        pair = [p1, p2]
        a03 = (all(_fmt_ok(ln, 9.0, C_BODY, False) for ln in pair)
               and _consecutive(d.right, pair)
               and _near(p1["x1"] - p1["x0"], 351.0, 30.0)
               and 10.0 <= p2["y0"] - p1["y0"] <= 19.0)
    _add(results, a03, 1, "该段排版属性保持——约9磅、深灰 #2F3333、非加粗、行宽约351磅、缩为2行且行距约13.5磅")

    # ---- 【新增"志愿服务"模块】 --------------------------------------
    vol_title = d.exact("志愿服务", side="R")
    vol_bar = d.bar_of_title(vol_title) if vol_title else None
    bars = sorted(d.module_bars(), key=lambda m: m["y0"])
    others = [b for b in bars if b is not vol_bar]
    ref_w = _median([b["w"] for b in others]) or RIGHT_W
    ref_gap = _median([d.gap_above(b) for b in others]) or 20.0
    mod_titles = [ln for ln in (d.find(t, side="R") for t in _MODULE_TITLES) if ln]
    ref_title_size = _median([ln["size"] for ln in mod_titles]) or 10.5
    skill_items = [d.find(t, side="R") for t in ("英语能力：", "工具能力：", "综合表现：")]
    skill_items = [ln for ln in skill_items if ln]
    ref_item_x0 = _median([ln["x0"] for ln in skill_items])
    ref_item_gap = _median([skill_items[i + 1]["y0"] - skill_items[i]["y0"]
                            for i in range(len(skill_items) - 1)]) or 16.5
    st.update(bars=bars, ref_item_x0=ref_item_x0, ref_item_gap=ref_item_gap,
              skill_items=skill_items, p2=p2)

    v1 = d.finds(_VOL1, side="R")
    v2 = d.finds(_VOL2, side="R")
    a07 = (len(v1) == 1 and len(v2) == 1 and v1[0] is not v2[0]
           and _norm(_VOL2) not in v1[0]["ntext"])
    vl = sorted([v1[0], v2[0]], key=lambda ln: ln["y0"]) if a07 else []

    a04 = False
    if vol_bar is not None and vl:
        below = [b["y0"] for b in others if b["y0"] > vol_bar["y1"]]
        next_top = min(below) if below else d.height
        inside = all(vol_bar["y1"] < ln["y0"] and ln["y1"] < next_top for ln in vl)
        overlap = any(not (b["y1"] <= vol_bar["y0"] + 0.5 or b["y0"] >= vol_bar["y1"] - 0.5)
                      for b in others)
        gap = d.gap_above(vol_bar)
        a04 = (inside and not overlap
               and vol_bar["x0"] >= COL_SPLIT - 6.0 and vol_bar["x1"] <= d.width - 12.0
               and gap >= max(6.0, ref_gap * 0.5) and gap <= ref_gap * 3.0 + 12.0
               and next_top - max(ln["y1"] for ln in vl) >= 3.0)
    _add(results, a04, 5, '新增1个独立的"志愿服务"模块，置于右栏正文流的模块之间，与相邻模块间距正常、未重叠、未溢出右边界')

    a05 = False
    if vol_bar is not None:
        tick = None
        for tk in d.ticks():
            if (vol_bar["y0"] - 2.0 <= tk["cy"] <= vol_bar["y1"] + 2.0
                    and 1.0 <= tk["x0"] - vol_bar["x0"] <= 20.0):
                tick = tk
                break
        a05 = (_near(vol_bar["w"], ref_w, max(ref_w * 0.06, 6.0))
               and 14.0 <= vol_bar["h"] <= 23.0
               and _color_ok(vol_bar["color"], C_BAR, vol_bar["tol"])
               and tick is not None)
    _add(results, a05, 3, '该模块标题栏与其他模块同款——深绿 #235D5B 横条约355×18磅、左端带橙色 #D4A557 竖条约4×12磅')

    a06 = (vol_title is not None and vol_title["ntext"] == "志愿服务"
           and _fmt_ok(vol_title, None, C_WHITE, True)
           and _size_ok(vol_title["size"], ref_title_size))
    _add(results, a06, 1, '标题文字为"志愿服务"四字，白色、加粗、字号约10.5磅，与其他模块标题一致')

    _add(results, a07, 3, "模块内容为两条志愿服务经历，分成两个独立行呈现、未合并串行、文字无错漏")

    a08 = bool(vl) and all(_fmt_ok(ln, 9.0, C_BODY, False) for ln in vl)
    _add(results, a08, 1, "两行内容字号约9磅、颜色深灰 #2F3333、非加粗，与相邻模块条目一致")

    a09 = False
    if vl and ref_item_x0 is not None:
        dots = d.dots_right()
        a09 = (all(d.dot_for(ln, dots) is not None for ln in vl)
               and all(_near(ln["x0"], ref_item_x0, 3.0) for ln in vl)
               and all(ln["x0"] - RIGHT_X >= 8.0 for ln in vl))
    _add(results, a09, 1, '两行内容带与"技能证书与获奖"条目同款的橙色圆点项目符号，文字左边界与该模块条目对齐')

    a10 = len(vl) == 2 and _near(vl[1]["y0"] - vl[0]["y0"], ref_item_gap, 2.0)
    _add(results, a10, 1, "两行内容之间的行距与其他模块条目行距一致（约16.5磅）")

    a11 = bool(vl) and all(_ends_with(ln, _END_PUNCT) for ln in vl)
    _add(results, a11, 1, "两行内容句末标点风格与文档其他条目一致，未出现无标点截断")

    st.update(vol_bar=vol_bar, vl=vl, a07=a07)


def _check_additions_tail(d, results, st):
    """一、加分项（"技能证书与获奖"补充条目 + 篇幅控制）。"""
    ref_item_x0 = st.get("ref_item_x0")
    ref_item_gap = st.get("ref_item_gap") or 16.5
    last_item = None
    for ln in st.get("skill_items") or []:
        if "综合表现" in ln["ntext"] and (last_item is None or ln["y0"] > last_item["y0"]):
            last_item = ln

    hits = [ln for ln in d.finds(_SKILL_NEW, side="R") if "综合表现" not in ln["ntext"]]
    new_item = hits[0] if len(hits) == 1 else None
    a12 = False
    if new_item is not None and last_item is not None:
        below = [b["y0"] for b in st.get("bars") or [] if b["y0"] > last_item["y1"]]
        next_top = min(below) if below else d.height
        a12 = (new_item["y0"] > last_item["y0"] and new_item["y1"] < next_top
               and _norm(_SKILL_NEW) not in _norm(last_item["text"]))
    _add(results, a12, 3, '"技能证书与获奖"模块末尾另起一行新增1条培养工程培训条目，未并入上一条也未覆盖既有条目')

    a13 = False
    if new_item is not None and ref_item_x0 is not None:
        a13 = (d.dot_for(new_item, d.dots_right()) is not None
               and _near(new_item["x0"], ref_item_x0, 2.0))
    _add(results, a13, 1, "该新增条目带同款橙色圆点项目符号，且文字左边界与同模块其余条目严格对齐")

    a14 = False
    if new_item is not None and last_item is not None:
        a14 = (_fmt_ok(new_item, 9.0, C_BODY, False)
               and _near(new_item["y0"] - last_item["y0"], ref_item_gap, 2.0))
    _add(results, a14, 1, "该新增条目字号约9磅、颜色深灰 #2F3333、非加粗，行距与同模块其余条目一致")

    a15 = _ends_with(new_item, _END_PERIOD)
    _add(results, a15, 1, "该新增条目以句号收尾，与同模块其余条目标点风格一致")

    # ---- 【篇幅控制】：仅在两处内容确已新增的前提下计分 ----------------
    added = bool(st.get("a07")) and a12
    band = d.footer_band()

    a16 = (added and d.n_pages == 1
           and _near(d.width, A4_W, 8.0) and _near(d.height, A4_H, 8.0))
    _add(results, a16, 5, "新增上述内容后全文仍为1页，未溢出到第2页，纸张仍为A4纵向")

    a17 = False
    if added and band is not None:
        ftr = [ln for ln in d.right
               if band["y0"] - 4.0 <= ln["y0"] and ln["y1"] <= band["y1"] + 4.0]
        a17 = (bool(ftr) and max(ln["x1"] for ln in ftr) >= 500.0
               and _band_intact(d, band))
    _add(results, a17, 1, "页面下沿的浅色页脚色带与其中的页码文字仍位于原位置，未被正文挤出页面或覆盖")

    a18 = False
    tail = None
    for ln in d.right:
        if ln["y1"] <= d.height - 60.0 and (tail is None or ln["y1"] > tail["y1"]):
            tail = ln
    if added and band is not None and tail is not None:
        a18 = tail["y1"] <= band["y0"] - 4.0
    _add(results, a18, 1, "正文末句与页脚色带之间仍留有可见空白，正文未紧贴或压住页脚")


def _check_deductions_right(d, results, st):
    """二、减分项（右栏既有版式）。"""
    # D01 大标题
    title = d.find("求学申请材料", side="R", y_range=(0.0, 60.0))
    d01 = not (_fmt_ok(title, 24.0, C_ACCENT, True) and title is not None
               and _near(title["x0"], RIGHT_X, 8.0))
    _ded(results, d01, 1, '右栏顶部标题"求学申请材料"不满足字号约24磅、颜色深绿 #173C3B、加粗、左端与右栏正文左对齐')

    # D02 副标题
    sub = d.find("面向能源信息与智能控制方向", side="R")
    d02 = not (_fmt_ok(sub, 11.0, C_MUTED, False)
               and title is not None and sub is not None
               and sub["size"] < title["size"] * 0.7)
    _ded(results, d02, 1, '标题下方一行"面向能源信息与智能控制方向"不满足字号约11磅、颜色灰 #6B7474、非加粗且明显小于主标题')

    # D03 金色分隔线
    d03 = not d.marks_of(C_ORANGE, wrange=(330.0, 375.0), hrange=(0.0, 5.0),
                         yrange=(60.0, 92.0))
    _ded(results, d03, 1, "标题区下方那条金色 #D4A557 细分隔横线不满足宽约352磅、位于标题区下沿、贯通右栏")

    # D04 报考与初试情况 四组标签/取值
    labels = [d.find(t, side="R") for t in ("报考院校：", "报考方向：", "培养类型：", "当前学历：")]
    values = [d.find(t, side="R") for t in ("南港理工学院", "新能源控制工程", "专业型硕士", "本科在读")]
    d04 = True
    if all(labels) and all(values):
        lx = sorted(set(round(ln["x0"]) for ln in labels))
        cols = []
        for x in lx:
            if not cols or x - cols[-1][-1] > 4.0:
                cols.append([x])
            else:
                cols[-1].append(x)
        fmt = (all(_fmt_ok(ln, 9.2, C_ACCENT, True) for ln in labels)
               and all(_fmt_ok(ln, 9.2, C_BODY, False) for ln in values))
        two_cols = (len(cols) == 2 and (cols[1][0] - cols[0][0]) > 100.0
                    and _near(cols[0][0], RIGHT_X, 8.0))
        d04 = not (fmt and two_cols)
    _ded(results, d04, 1, '"报考与初试情况"四组"标签+取值"不满足标签约9.2磅深绿 #173C3B 加粗、取值约9.2磅深灰 #2F3333 非加粗，或未保持左右两列各两行的对齐关系')

    # D05 初试成绩表
    heads = [d.exact(t, side="R") for t in ("思政", "外语", "专业基础", "数学", "总分")]
    vals = [d.exact(t, side="R") for t in ("64", "71", "118", "96", "349")]
    cells = d.marks_of(C_TBL_HEAD, wrange=(60.0, 82.0), hrange=(19.0, 29.0), pool="raw")
    grid = [m for m in d.raw
            if m["stroke"] and _color_ok(m["stroke"], C_TBL_LINE, _TOL_VEC)]
    d05 = not (len(cells) == 5 and len(grid) >= 4 and all(heads) and all(vals)
               and all(_fmt_ok(ln, 9.5, C_ACCENT, True) for ln in heads)
               and all(_fmt_ok(ln, 9.5, C_BODY) for ln in vals))
    _ded(results, d05, 1, "初试成绩表不满足5列表头带浅绿 #DDEBE6 底色、单元格约71×24磅、细框线 #C3D1CD 完整、表头约9.5磅深绿加粗、数值约9.5磅深灰")

    # D06 教育背景行
    row = [d.find("2022.09", side="R"), d.find("东序信息学院", side="R"),
           d.find("计算机与自动化技术（本科）", side="R")]
    d06 = True
    if all(row):
        cys = [(ln["y0"] + ln["y1"]) / 2.0 for ln in row]
        d06 = not (_fmt_ok(row[0], 9.6, C_ACCENT, True)
                   and _fmt_ok(row[1], 10.2, C_ACCENT, True)
                   and _fmt_ok(row[2], 9.6, C_ACCENT, True)
                   and max(cys) - min(cys) <= 5.0
                   and _near(row[0]["x0"], RIGHT_X, 8.0)
                   and 300.0 <= (row[1]["x0"] + row[1]["x1"]) / 2.0 <= 400.0
                   and row[2]["x1"] >= 540.0)
    _ded(results, d06, 1, '"教育背景"中"起止年月/院校名称/专业与学历"一行不满足三段同一基线、字号约9.6/10.2/9.6磅、颜色深绿 #173C3B 加粗、左中右三点分布')

    # D07 课程结构段落
    para = [d.find("课程结构以算法设计", side="R"), d.find("与工程实践为主", side="R"),
            d.find("本化处理与版本协作", side="R")]
    d07 = True
    if all(para):
        gaps = [para[i + 1]["y0"] - para[i]["y0"] for i in range(2)]
        d07 = not (all(_fmt_ok(ln, 9.1, C_BODY, False) for ln in para)
                   and _consecutive(d.right, para)
                   and _near(para[0]["x1"] - para[0]["x0"], RIGHT_W, 28.0)
                   and all(10.0 <= g <= 19.0 for g in gaps))
    _ded(results, d07, 1, '"课程结构以算法设计…"整段不满足约9.1磅、深灰 #2F3333、非加粗、排为3行且行宽约355磅、行距约14磅')

    # D08 实践经历行
    row2 = [d.find("2025.02", side="R"), d.find("青封智能装配有限公司", side="R"),
            d.find("智能监测开发助理", side="R")]
    d08 = True
    if all(row2):
        cys = [(ln["y0"] + ln["y1"]) / 2.0 for ln in row2]
        d08 = not (_fmt_ok(row2[0], 9.4, C_ACCENT, True)
                   and _fmt_ok(row2[1], 10.0, C_ACCENT, True)
                   and _fmt_ok(row2[2], 9.4, C_ACCENT, True)
                   and max(cys) - min(cys) <= 5.0
                   and _near(row2[0]["x0"], RIGHT_X, 8.0)
                   and 300.0 <= (row2[1]["x0"] + row2[1]["x1"]) / 2.0 <= 400.0
                   and row2[2]["x1"] >= 500.0)
    _ded(results, d08, 1, '"实践经历"中"起止年月/单位名称/岗位名称"一行不满足三段同一基线、字号约9.4/10.0/9.4磅、颜色深绿 #173C3B 加粗、左中右三点分布')

    # D09 课程项目 小标题行
    proj_l = d.find("课程项目", side="R")
    proj_r = d.find("楼宇能耗预测与调度小组项目", side="R")
    d09 = not (_fmt_ok(proj_l, 9.4, C_ACCENT, True) and _fmt_ok(proj_r, 10.0, C_ACCENT, True))
    _ded(results, d09, 1, '"课程项目"与"楼宇能耗预测与调度小组项目"两处小标题不满足字号约9.4/10.0磅、颜色深绿 #173C3B、加粗')

    # D10 项目说明段落
    pj = [d.find("围绕温湿度", side="R"), d.find("项目说明书与演示看板", side="R")]
    d10 = True
    if all(pj):
        d10 = not (all(_fmt_ok(ln, 9.0, C_BODY, False) for ln in pj)
                   and _consecutive(d.right, pj)
                   and _near(pj[0]["x1"] - pj[0]["x0"], 351.0, 30.0))
    _ded(results, d10, 1, '"围绕温湿度…"整段不满足约9磅、深灰 #2F3333、非加粗、排为2行且行宽约351磅')

    # D11 课程项目行上下间距
    d11 = True
    prev = st.get("p2")
    if proj_l is not None and prev is not None and pj[0] is not None:
        g1 = proj_l["y0"] - prev["y1"]
        g2 = pj[0]["y0"] - proj_l["y1"]
        d11 = not (-1.0 <= g1 <= 45.0 and -1.0 <= g2 <= 45.0)
    _ded(results, d11, 1, '"课程项目"小标题行与其上下文字之间的垂直间距未保持正常，出现紧贴、重叠或大片空洞')

    # D12 技能证书与获奖 原有三条
    items = st.get("skill_items") or []
    d12 = True
    if len(items) == 3:
        items = sorted(items, key=lambda ln: ln["y0"])
        dots = d.dots_right()
        gaps = [items[i + 1]["y0"] - items[i]["y0"] for i in range(2)]
        xs = [ln["x0"] for ln in items]
        d12 = not (all(_fmt_ok(ln, 9.0, C_BODY, False) for ln in items)
                   and all(d.dot_for(ln, dots) is not None for ln in items)
                   and max(xs) - min(xs) <= 2.5 and min(xs) - RIGHT_X >= 8.0
                   and all(13.5 <= g <= 20.0 for g in gaps))
    _ded(results, d12, 1, '"技能证书与获奖"模块原有三条条目不满足约9磅深灰 #2F3333 非加粗、每条带橙色圆点、左缩进约12磅且行距约16.5磅')

    # D13 研修规划段落
    head = d.find("进入研究生阶段后", side="R")
    d13 = True
    if head is not None:
        plan = sorted([ln for ln in d.right
                       if ln["y0"] >= head["y0"] - 0.5 and ln["y1"] <= d.height - 60.0],
                      key=lambda ln: ln["y0"])
        if len(plan) == 5:
            gaps = [plan[i + 1]["y0"] - plan[i]["y0"] for i in range(4)]
            base = _median([ln["x0"] for ln in plan[1:]])
            d13 = not (all(_fmt_ok(ln, 9.0, C_BODY, False) for ln in plan)
                       and 10.0 <= plan[0]["x0"] - base <= 28.0
                       and all(10.0 <= g <= 19.0 for g in gaps))
        st["plan_last"] = plan[-1] if plan else None
    _ded(results, d13, 1, '"研修规划"整段不满足约9磅、深灰 #2F3333、非加粗、首行缩进约18磅、排为5行且行距约14磅')

    # D14 五个模块标题栏
    d14 = False
    for name in _MODULE_TITLES:
        ln = d.find(name, side="R")
        bar = d.bar_of_title(ln) if ln else None
        ok = (_fmt_ok(ln, 10.5, C_WHITE, True) and bar is not None
              and 14.0 <= bar["h"] <= 23.0 and _near(bar["w"], RIGHT_W, 26.0)
              and any(bar["y0"] - 2.0 <= tk["cy"] <= bar["y1"] + 2.0
                      and 1.0 <= tk["x0"] - bar["x0"] <= 20.0 for tk in d.ticks()))
        if not ok:
            d14 = True
            break
    _ded(results, d14, 1, "右栏五个既有模块标题栏不满足深绿 #235D5B 横条约355×18磅、标题白色加粗约10.5磅、左端橙色 #D4A557 竖条约4×12磅")

    # D15 页脚文字
    band = d.footer_band()
    ftr = [ln for ln in d.right if ln["y0"] >= d.height - 60.0]
    d15 = True
    if band is not None and ftr:
        d15 = not (all(_fmt_ok(ln, 7.6, C_MUTED, False) for ln in ftr)
                   and max(ln["x1"] for ln in ftr) >= 500.0
                   and all(band["y0"] - 4.0 <= ln["y0"] and ln["y1"] <= band["y1"] + 4.0
                           for ln in ftr))
    _ded(results, d15, 1, '页面底部"求学申请材料 · 2026"一行不满足约7.6磅、颜色灰 #6B7474、非加粗、靠右排布并位于浅色页脚色带内')


def _check_deductions_left(d, results, st):
    """二、减分项（左栏侧边栏既有版式）。"""
    # D16 左侧通栏与最窄深色条
    panel = d.marks_of(C_BAR, wrange=(163.0, 188.0), hrange=(780.0, 1200.0),
                       xrange=(-2.0, 3.0), pool="raw")
    strip = d.marks_of(C_STRIP, wrange=(14.0, 22.0), hrange=(780.0, 1200.0),
                       xrange=(-2.0, 3.0), pool="raw")
    d16 = not (panel and strip)
    _ded(results, d16, 1, "左侧深绿 #235D5B 竖向通栏不满足宽约175磅、自页面顶端贯通至底端，或其最左侧那条约18磅宽的更深色 #163B3B 窄条缺失")

    # D17 圆形头像
    av = _find_avatar(d)
    d17 = True
    if av is not None:
        cs = [d.px(av["x0"] + 2.5, av["y0"] + 2.5), d.px(av["x1"] - 2.5, av["y0"] + 2.5),
              d.px(av["x0"] + 2.5, av["y1"] - 2.5), d.px(av["x1"] - 2.5, av["y1"] - 2.5)]
        flat = max(max(abs(a[i] - b[i]) for i in range(3)) for a in cs for b in cs)
        d17 = not (_size_ok(av["w"], 84.0) and flat <= 30
                   and _near((av["x0"] + av["x1"]) / 2.0, 87.0, 10.0))
    _ded(results, d17, 1, "左栏上部那张圆形裁切的人物头像不满足直径约84磅、正圆裁切、水平居中于左栏，或已被删除/拉伸变形")

    # D18 姓名区与胶囊标签
    name = d.find("林绯绯", side="L")
    major = d.find("计算机与自动化技术", side="L")
    tagline = d.find("跨学科申请", side="L")
    pill = d.marks_of(C_ORANGE, wrange=(84.0, 102.0), hrange=(16.0, 24.0),
                      xrange=(20.0, 60.0), pool="raw")
    d18 = not (_fmt_ok(name, 18.0, C_WHITE, True)
               and _fmt_ok(major, 9.5, C_SUBTITLE_L)
               and _fmt_ok(tagline, 9.0, C_ACCENT)
               and bool(pill))
    _ded(results, d18, 1, '左栏"林绯绯"不满足约18磅白色加粗，或"本科 · 计算机与自动化技术"不满足约9.5磅浅绿 #D8E7E4，或"跨学科申请"及其橙色 #D4A557 胶囊底块（约93×20磅）缺失变形')

    # D19 左栏四个分组标题条
    gbars = d.marks_of(C_GROUP_BAR, wrange=(110.0, 132.0), hrange=(20.0, 28.0), pool="raw")
    d19 = len(gbars) != 4
    if not d19:
        for name_t in _GROUP_TITLES:
            ln = d.find(name_t, side="L")
            if not _fmt_ok(ln, 10.2, C_WHITE, True):
                d19 = True
                break
            if not any(m["y0"] - 4.0 <= ln["y0"] and ln["y1"] <= m["y1"] + 4.0
                       and m["x0"] - 4.0 <= ln["x0"] and ln["x1"] <= m["x1"] + 4.0
                       for m in gbars):
                d19 = True
                break
    _ded(results, d19, 1, '左栏"基本信息""联系方式""能力概览""申请信息"四个分组标题不满足白色加粗约10.2磅并位于约120×24磅的浅绿 #3F7774 色块内')

    # D20 基本信息 五行标签与取值
    b_lab = [d.find(t, side="L") for t in ("出生年月", "政治面貌", "民族", "最高学历", "毕业院校")]
    b_val = [d.find(t, side="L") for t in ("2001.11", "共青团员", "汉族", "本科学历", "东序信息学院")]
    d20 = True
    if all(b_lab) and all(b_val):
        d20 = not (all(_fmt_ok(ln, 8.8, C_LABEL_L) for ln in b_lab)
                   and all(_fmt_ok(ln, 8.8, C_WHITE) for ln in b_val)
                   and all(_near(ln["x0"], 33.0, 5.0) for ln in b_lab)
                   and all(_near(ln["x0"], 86.0, 6.0) for ln in b_val))
    _ded(results, d20, 1, '"基本信息"五行不满足标签约8.8磅浅绿 #CFE1DE 起于左栏约33磅处、取值约8.8磅白色起于约86磅处的两列对齐')

    # D21 联系方式 三行
    ct = [d.find("校内申请系统转接", side="L"), d.find("联系邮箱", side="L"),
          d.find("常驻城市", side="L")]
    d21 = True
    if all(ct):
        ct = sorted(ct, key=lambda ln: ln["y0"])
        gaps = [ct[i + 1]["y0"] - ct[i]["y0"] for i in range(2)]
        d21 = not (all(_fmt_ok(ln, 8.5, C_WHITE, False) for ln in ct)
                   and all(_near(ln["x0"], 33.0, 5.0) for ln in ct)
                   and all(12.0 <= g <= 22.0 for g in gaps))
    _ded(results, d21, 1, '"联系方式"三行不满足约8.5磅白色非加粗、左端对齐于约33磅处、行距约17磅')

    # D22 能力概览 五条
    caps = [d.find(t, side="L") for t in ("Python / C++ / C#", "数据处理与可视化",
                                          "边缘计算部署", "控制系统仿真", "技术文档写作")]
    d22 = True
    if all(caps):
        caps = sorted(caps, key=lambda ln: ln["y0"])
        ldots = d.dots_left()
        gaps = [caps[i + 1]["y0"] - caps[i]["y0"] for i in range(4)]
        d22 = not (all(_fmt_ok(ln, 8.6, C_WHITE) for ln in caps)
                   and len(ldots) >= 5
                   and all(d.dot_for(ln, ldots) is not None for ln in caps)
                   and all(14.0 <= g <= 23.0 for g in gaps))
    _ded(results, d22, 1, '"能力概览"五条不满足约8.6磅白色、每条前带约4.4磅橙色 #D4A557 圆点、行距约18磅')

    # D23 申请信息 三行
    a_lab = [d.find(t, side="L") for t in ("报考院校", "报考专业", "报考属性")]
    a_val = [d.find(t, side="L") for t in ("南港理工学院", "新能源控制工程", "专业型硕士")]
    d23 = True
    if all(a_lab) and all(a_val):
        d23 = not (all(_fmt_ok(ln, 8.6, C_LABEL_L) for ln in a_lab)
                   and all(_fmt_ok(ln, 8.6, C_WHITE) for ln in a_val)
                   and all(_near(ln["x0"], 33.0, 5.0) for ln in a_lab)
                   and all(_near(ln["x0"], 86.0, 6.0) for ln in a_val))
    _ded(results, d23, 1, '"申请信息"三行不满足标签约8.6磅浅绿 #CFE1DE、取值约8.6磅白色的两列对齐')

    # D24 左下装饰弧
    d24 = _find_arc(d) is None
    _ded(results, d24, 1, "左栏左下角那片浅色 #C8DDD6 装饰圆弧被删除、移位或改色")


def _check_deductions_global(d, results, st):
    """二、减分项（整体性破坏）。"""
    # D25 文字内容
    miss = [t for t in _ANCHORS_RIGHT if _norm(t) not in d.flat_right and not d.has(t)]
    miss += [t for t in _ANCHORS_LEFT if _norm(t) not in d.flat_left and not d.has(t)]
    _ded(results, bool(miss), 3, "除被横线标注的那一句之外，其余原有文字段落、条目与标签存在删除、改写或错字")

    # D26 图形对象
    objs = {
        "头像": _find_avatar(d) is not None,
        "成绩表底色": len(d.marks_of(C_TBL_HEAD, wrange=(60.0, 82.0), hrange=(19.0, 29.0),
                                     pool="raw")) >= 5,
        "模块标题栏": len(d.module_bars()) >= 5,
        "右栏圆点": len(d.dots_right()) >= 3,
        "左栏圆点": len(d.dots_left()) >= 5,
        "装饰弧": _find_arc(d) is not None,
        "页脚色带": d.footer_band() is not None,
        "胶囊底块": bool(d.marks_of(C_ORANGE, wrange=(84.0, 102.0), hrange=(16.0, 24.0),
                                    xrange=(20.0, 60.0), pool="raw")),
        "分组标题条": len(d.marks_of(C_GROUP_BAR, wrange=(110.0, 132.0),
                                     hrange=(20.0, 28.0), pool="raw")) >= 4,
    }
    _ded(results, not all(objs.values()), 3, "头像、初试成绩表框线与底色、模块标题栏色块、条目圆点、左下装饰弧形、页脚色带等既有图形对象存在删除或破坏")

    # D27 模块与段落存废
    lost = [t for t in _MODULE_TITLES if d.find(t, side="R") is None]
    lost += [t for t in _GROUP_TITLES if d.find(t, side="L") is None]
    for anchor in ("课程结构以算法设计", "围绕温湿度", "进入研究生阶段后", "英语能力"):
        if d.find(anchor, side="R") is None:
            lost.append(anchor)
    _ded(results, bool(lost), 3, "右栏五个既有模块或左栏四个既有分组中的任一整块被删除、合并或整段搬移")

    # D28 具体数值
    nums = ["林绯绯", "2001.11", "南港理工学院", "新能源控制工程", "专业型硕士", "本科在读",
            "东序信息学院", "2022.09", "2026.06", "2025.02", "2025.07"]
    bad_num = [t for t in nums if not d.has(t)]
    if not bad_num:
        bad_num = [v for v in ("64", "71", "118", "96", "349") if d.exact(v, side="R") is None]
    _ded(results, bool(bad_num), 1, "姓名、出生年月、初试各科分数与总分、院校与专业名称、各段起止年月等具体数值被改动")

    # D29 多余新增
    extra = False
    if d.n_pages != 1:
        extra = True
    if len(d.module_bars()) > 6 or len(d.marks_of(C_GROUP_BAR, wrange=(110.0, 132.0),
                                                 hrange=(20.0, 28.0), pool="raw")) > 4:
        extra = True
    if any(ln["y1"] <= 16.0 or ln["y0"] >= d.height - 8.0 for ln in d.lines):
        extra = True
    for anchor, limit in (("技能证书与获奖", 4), ("志愿服务", 2)):
        title = d.exact(anchor, side="R")
        bar = d.bar_of_title(title) if title else None
        if bar is None:
            continue
        below = [b["y0"] for b in d.module_bars() if b["y0"] > bar["y1"]]
        top = min(below) if below else d.height - 60.0
        inner = [ln for ln in d.right if bar["y1"] <= ln["y0"] and ln["y1"] <= top]
        if len(inner) > limit:
            extra = True
    _ded(results, extra, 1, "在指令未要求处新增了多余的页面、模块、分组、条目、页眉页脚或装饰元素")


# ---------------------------------------------------------------------------
# 汇总与对外入口
# ---------------------------------------------------------------------------


def _check_all(d):
    """按 _MAX_DELTAS_ORDER 的顺序产出 47 条结果字符串。"""
    results = []
    st = {}
    _check_additions(d, results, st)
    _check_additions_tail(d, results, st)
    _check_deductions_right(d, results, st)
    _check_deductions_left(d, results, st)
    _check_deductions_global(d, results, st)
    return results


def _evaluate_impl(dir_path, work_dir):
    display_name, pdf_path = _resolve_pdf(dir_path, work_dir)
    # 维度一的第一道门禁：交付物格式只接受 .pdf，不合格即判 0，不再解析版面
    reject = _fmt_reject(pdf_path)
    if reject:
        return {
            "id": SCRIPT_ID,
            "file_name": display_name,
            "status": "ok",
            "error": None,
            "dim1_pass": False,
            # 维度一未通过：dim1_reason 必须带上错误摘要，便于直接定位原因
            "dim1_reason": "维度一未通过：%s，完成度记 0" % reject,
            "dim2_items": [],
            "total_score": 0,
            "max_score": MAX_SCORE_TOTAL,
        }
    d = _Doc(pdf_path)
    try:
        ok1, reason1 = _dim1(d)
        if not ok1:
            return {
                "id": SCRIPT_ID,
                "file_name": display_name,
                "status": "ok",
                "error": None,
                "dim1_pass": False,
                # 维度一未通过：dim1_reason 必须带上错误摘要，便于直接定位原因
                "dim1_reason": "维度一未通过：%s，完成度记 0" % reason1,
                "dim2_items": [],
                "total_score": 0,
                "max_score": MAX_SCORE_TOTAL,
            }
        results = _check_all(d)
    finally:
        d.close()

    if len(results) != len(_MAX_DELTAS_ORDER):
        raise RuntimeError(
            f"结果条目数({len(results)})与 max_delta 序列({len(_MAX_DELTAS_ORDER)})不一致")

    dim2_items = []
    total = 0
    for raw, max_delta in zip(results, _MAX_DELTAS_ORDER):
        head, _, rule = raw.partition("：")
        got = int(head.replace("+", "").replace("-", "").strip())
        hit = got > 0
        delta = max_delta if hit else 0
        total += delta
        dim2_items.append({
            "rule": rule,
            "max_delta": max_delta,
            "delta": delta,
            "hit": hit,
        })
    total = max(total, 0)
    return {
        "id": SCRIPT_ID,
        "file_name": display_name,
        "status": "ok",
        "error": None,
        "dim1_pass": True,
        "dim1_reason": reason1,
        "dim2_items": dim2_items,
        "total_score": total,
        "max_score": MAX_SCORE_TOTAL,
    }


def evaluate(dir_path):
    """对 dir_path 下的交付文档评分，返回结果字典（不抛异常、不打印）。

    - 交付物只接受 .pdf：扩展名不是 .pdf、或文件打不开（缺 %PDF- 头、加密、零页），
      一律在维度一判不通过（dim1_pass=False、total_score=0），而不是当成没有交付物。
      目录里没有 .pdf 但有 .docx/.pptx 等其他格式时，仍收作被评对象走这条判负路径。
    - 定位不到交付文件时：error 只回纯文本原因（不带异常类名），
      dim1_pass 与 dim1_reason 一律留空（空字符串，表示维度一根本没进入判定），
      dim2_items=[]、total_score=max_score=0。
    - 维度一未通过时：dim1_pass=False，且 dim1_reason 带出错误摘要。
    """
    work_dir = tempfile.mkdtemp(prefix="officeval_%s_" % SCRIPT_ID)
    try:
        if not dir_path or not os.path.isdir(dir_path):
            raise FileNotFoundError("目录不存在或不是目录：%s" % dir_path)
        return _evaluate_impl(dir_path, work_dir)
    except FileNotFoundError as exc:
        # 找不到被评估文件：error 只回传纯文本原因，不带异常类名；维度一根本没进入
        # 判定，dim1_pass 与 dim1_reason 一律留空（空字符串，而非 False），满分归零
        return {
            "id": SCRIPT_ID,
            "file_name": "",
            "status": "error",
            "error": str(exc),
            "dim1_pass": "",
            "dim1_reason": "",
            "dim2_items": [],
            "total_score": 0,
            "max_score": 0,
        }
    except Exception as exc:  # noqa: BLE001 — 其余异常都折叠为 error 状态
        # 文件在但打不开/解析失败：维度一未通过，dim1_reason 带上错误摘要便于定位
        return {
            "id": SCRIPT_ID,
            "file_name": "",
            "status": "error",
            "error": "脚本运行异常：%s: %s" % (type(exc).__name__, exc),
            "dim1_pass": False,
            "dim1_reason": "评测过程异常终止：%s: %s" % (type(exc).__name__, exc),
            "dim2_items": [],
            "total_score": 0,
            "max_score": MAX_SCORE_TOTAL,
        }
    finally:
        shutil.rmtree(work_dir, ignore_errors=True)


if __name__ == "__main__":
    # 本地调试入口：默认评估脚本所在目录，可用 argv[1] 覆盖目标目录。
    _target = sys.argv[1] if len(sys.argv) > 1 else os.path.dirname(os.path.abspath(__file__))
    _out = evaluate(_target)
    _buf = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", newline="")
    _buf.write(json.dumps(_out, ensure_ascii=False, indent=2))
    _buf.write("\n")
    _buf.flush()









