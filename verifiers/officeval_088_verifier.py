#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# pyright: reportUnknownMemberType=false, reportUnknownVariableType=false, reportUnknownArgumentType=false, reportUnknownParameterType=false, reportMissingParameterType=false, reportPrivateUsage=false, reportUnknownLambdaType=false, reportGeneralTypeIssues=false, reportOptionalMemberAccess=false, reportAttributeAccessIssue=false, reportIndexIssue=false, reportArgumentType=false, reportCallIssue=false, reportOptionalSubscript=false

import os
import re
import io
import sys
import json
import zipfile

from docx import Document
from docx.oxml.ns import qn
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

try:
    from PIL import Image
except ImportError:
    Image = None
try:
    import numpy as np
except ImportError:
    np = None
try:
    import cv2
except ImportError:
    cv2 = None

# 说明：按接口统一约定，本模块不修改全局 sys.stdout，也不 print 主结果。
# 主结果统一由 evaluate(dir_path) -> dict 返回；入参为“交付物所在目录的路径”，
# 由脚本自己在该目录内定位并打开被评估的 Word 文档与 Excel 工作簿。本文件仅在
# __main__ 下作为本地调试入口，才把 dict 序列化为 JSON 打印。
#
# 本任务交付物为“两个文件”：1 个 Word 文档（产品分析）+ 1 个 Excel 工作簿
# （选货汇总表）。维度一要求两者齐备、格式在接受范围内、可正常打开且内容可编辑；
# 任一不满足则 dim1_pass=False 且 total_score=0。
#
# 【图片内容项的可机检边界】——以下两点属于“需要人眼语义判断”的内容，纯代码无法
# 判定，脚本只判定其中可机检的结构/像素条件，并在对应评分项注释中标明：
#   ① “图片机型与产品名称一致”：无法由像素判定机型归属，脚本改判“张数、每个产品
#      各 1 张、位置落在展示图列/右列、5(3) 张互不相同且非空白画面”。
#   ② “机身完整、四边未被画面边界截断”：白底白机身使边缘着墨、背景差分、暗区占比
#      等代理指标在样本上完全无区分度（已实测），脚本只判定同一评分项内可机检的
#      “像素宽高比 0.55–0.90”与“左右纯白空白带各 ≤10% 宽度”两个条件。
#   ③ “画面内无文字类宣传元素”部分可机检：宣传横幅/海报标题的判别特征是“高饱和
#      色块内部存在成规模的白色字形团”，见 _has_promo_badge；标注容量、品类的小
#      尺寸规格角标（如“12.5KG／洗衣机”）按细则不属于宣传文案，已按规模阈值排除，
#      而紧凑型宣传角标与紧凑型规格角标的区分需要 OCR（运行环境无 OCR），故这类
#      小角标一律不判命中。

SCRIPT_ID = "088"
# 维度二加分项 max_delta 之和；用于维度一未通过 / 脚本异常时兜底
MAX_SCORE_TOTAL = 62
# 维度二减分项 max_delta 之和（负值），仅用于自检
MAX_DEDUCT_TOTAL = -38

# 维度二所有评分项按“结果追加顺序”的 max_delta 列表；末端把 results 里的
# “+N：...”/“-N：...”字符串反解为结构化 dim2_items 时用。顺序须与 _evaluate_impl
# 中 add(...) 的调用顺序严格一致；共 44 项，正项累加 = MAX_SCORE_TOTAL，
# 负项累加 = MAX_DEDUCT_TOTAL。
_MAX_DELTAS_ORDER = [
    3,   # 01 Excel 5 个产品数据区块及产品名称清单
    3,   # 02 Excel 指导价格列 5 个纯价格值
    5,   # 03 Excel 核心参数 5 条内容准确完整
    3,   # 04 Excel 核心参数逐项分行
    3,   # 05 Excel 适用场景整段描述（字数/开头/结尾/关键词）
    1,   # 06 Excel 核心参数与适用场景自动换行 + 垂直居中
    1,   # 07 Excel 适用场景列列宽 ≥30 字符
    5,   # 08 Excel 展示图每区块 1 张、共 5 张、与产品对应
    3,   # 09 Excel 展示图画面内无文字类宣传元素
    3,   # 10 Excel 展示图机身完整（像素宽高比 + 左右白边）
    3,   # 11 Excel 展示图嵌入单元格显示
    3,   # 12 Excel 5 个区块行数一致且五项横向对齐
    5,   # 13 Word 3 行按指导价格从高到低排列
    1,   # 14 Word 三行产品名称文字
    3,   # 15 Word 三行核心参数内容 + 逐项分行
    3,   # 16 Word 三行适用场景整段描述
    1,   # 17 Word 三行指导价格纯价格值
    5,   # 18 Word 右列 3 个单元格各 1 张图片、与同行产品对应
    3,   # 19 Word 右列 3 张图片显示尺寸与不变形
    3,   # 20 Word 页眉居中“产品分析”+ 下方横线
    1,   # 21 Word 页脚居中页码域、自第 1 页连续编号
    1,   # 22 Word 文件名主体为“产品分析”
    -3,  # 23 Word 表格 6 个单元格四周浅灰细实线边框
    -3,  # 24 Word 页边距 2.0 厘米 / 页眉页脚距边界 0 厘米
    -3,  # 25 Word 标题“产品摘要”小初、不加粗
    -1,  # 26 Word 标题段前段后间距 0 磅、水平居中
    -1,  # 27 Word 标题与表格之间的空行段落
    -3,  # 28 Word 表格左列 12.00 厘米 / 右列 4.40 厘米
    -3,  # 29 Word 表格行高最小值 5.10 厘米
    -1,  # 30 Word 表格禁止跨页断行 / 无分页符分节符
    -1,  # 31 Word 表格整体水平居中 + 固定列宽
    -1,  # 32 Word 表格单元格内边距 0.19 / 0 厘米
    -1,  # 33 Word 单元格垂直居中 + 右列图片段落水平居中
    -1,  # 34 Word“产品一/二/三”宋体小三加粗
    -1,  # 35 Word 四个字段提示词宋体五号加粗
    -1,  # 36 Word 填入的产品信息宋体五号不加粗
    -3,  # 37 Word 正文无多余非空段落或第二个表格
    -1,  # 38 Word 纸张尺寸 A4
    -3,  # 39 Excel 第一行五个表头文字及其后无多余表头
    -3,  # 40 Excel 四列列宽
    -1,  # 41 Excel 表头字体字号居中底纹边框
    -1,  # 42 Excel 工作表名称与数量
    -1,  # 43 Excel 展示图像素宽度 ≥200
    -1,  # 44 Excel/Word 核心参数无多余参数字段
]
assert len(_MAX_DELTAS_ORDER) == 44
assert sum(d for d in _MAX_DELTAS_ORDER if d > 0) == MAX_SCORE_TOTAL
assert sum(d for d in _MAX_DELTAS_ORDER if d < 0) == MAX_DEDUCT_TOTAL

_EXT_WORD_OK = (".docx",)
# 维度一不接受、但可能被误交的其它文档格式；命中时判维度一不通过，而非报脚本异常。
_EXT_WORD_REJECT = (".doc", ".wps", ".rtf", ".odt", ".pdf", ".txt", ".md")
_EXT_XL_OK = (".xlsx", ".xlsm")
_EXT_XL_REJECT = (".xls", ".xlsb", ".csv", ".ods", ".et")

# ===== 交付内容基准常量（判定文字内容是否齐全、准确） =====
_NAMES = ("云溪5.0 D3滚筒洗衣机", "统帅三桶懒人洗衣机", "小天鹅12KG滚筒洗衣机",
          "米家8KG滚筒洗衣机", "美的8KG波轮洗衣机")
_PRICES = ("4500元", "3999元", "3039元", "999元", "846元")
# 每款产品的核心参数键值对（按“键+值”压缩后子串匹配）
_PARAMS = (
    (("洗涤容量", "12KG"), ("宽度", "595mm"), ("类型", "滚筒洗衣机"),
     ("能效等级", "一级"), ("深度", "593mm537mm"), ("高度", "850mm")),
    (("洗涤容量", "大筒10.5KG+双小筒各1KG"), ("宽度", "600mm"), ("类型", "滚筒洗衣机"),
     ("能效等级", "一级"), ("深度", "570mm")),
    (("洗涤容量", "12KG"), ("宽度", "595mm"), ("类型", "滚筒洗衣机"),
     ("能效等级", "一级"), ("深度", "560mm585mm")),
    (("洗涤容量", "8KG"), ("宽度", "598mm"), ("类型", "滚筒洗衣机"),
     ("能效等级", "一级"), ("深度", "495mm")),
    (("洗涤容量", "8KG"), ("宽度", "515mm"), ("类型", "波轮洗衣机"),
     ("能效等级", "一级"), ("深度", "525mm")),
)
# 核心参数单元格内每项参数各占一行时，各产品应有的行数与行首提示词
_PARAM_LINES = (6, 5, 5, 5, 5)
_PARAM_KEYS = ("洗涤容量", "宽度", "类型", "能效等级", "深度", "高度")
# 只允许出现上述参数字段；出现其它字段（如“产品型号”）计入减分项 44
_EXTRA_PARAM_KEYS = ("产品型号", "型号", "品牌", "颜色", "转速", "重量", "净重",
                     "毛重", "内筒材质", "电机", "噪音", "水效", "功率", "洗净比",
                     "烘干容量", "上市时间", "产地", "保修")
_SCENE_HEAD = "本产品适配多种场景，核心适用于"
_SCENE_MIN_CHARS = 60
_SCENE_KEYS_XL = (("多人口家庭", "大件家纺", "小型民宿", "公寓出租房"),
                  ("多人口家庭", "分类洗护", "谁的衣服先洗", "公寓出租房"),
                  ("多人口家庭", "大件家纺", "母婴家庭", "高温除菌"),
                  ("小人口家庭", "薄家纺", "员工宿舍", "高性价比"),
                  ("小人口家庭", "薄家纺", "员工宿舍", "高性价比"))
_SCENE_KEYS_WD = (("多人口家庭", "大件家纺", "公寓出租房"),
                  ("分类洗护", "谁的衣服先洗"),
                  ("母婴家庭", "高温除菌"))
_HEADERS = ("产品名称", "核心参数", "适用场景", "指导价格", "展示图")
_FIELD_LABELS = ("产品名称：", "核心参数：", "适用场景：", "指导价格：")
_PROD_LABELS = ("产品一：", "产品二：", "产品三：")
# 宣传文案关键词（用于文字类宣传元素落在可提取文本时的直接判定）。
# 标注容量、品类的规格角标（如“12.5KG”“洗衣机”字样）不属于宣传文案，不列入本表。
_PROMO_WORDS = ("巅峰之作", "AI全能旗舰", "高端净护选直驱", "一级健康洁净标准",
                "10年包修", "纯平自由嵌", "预混高能洁净力", "PTC柔烘科技",
                "超薄省空间", "高效除菌洗", "汽车级防锈工艺",
                "防腐防生锈")

# ===== 交付格式基准常量（减分项判定阈值） =====
_CM_PER_TWIP = 1.0 / 567.0
_CM_PER_EMU = 1.0 / 360000.0
_A4_CM = (21.0, 29.7)
_MARGIN_CM = 2.0
_HDR_FTR_CM = 0.0
_TITLE_TEXT = "产品摘要"
_TITLE_PT = 36.0          # 小初
_PROD_LABEL_PT = 15.0     # 小三
_BODY_PT = 10.5           # 五号
_SONGTI = ("宋体", "SimSun", "NSimSun")
_COL_LEFT_CM = 12.00
_COL_RIGHT_CM = 4.40
_ROW_H_CM = 5.10
_CELL_MAR_LR_CM = 0.19
_CELL_MAR_TB_CM = 0.0
_BORDER_COLOR = "D9D9D9"
_BORDER_SZ = 6            # 0.75 磅 = 6 个 1/8 磅
_XL_COL_W = {"产品名称": 24.13, "核心参数": 33.52, "指导价格": 13.08, "展示图": 12.79}
_XL_COL_W_TOL = 1.5
_XL_HDR_FONT = "Noto Sans CJK SC"
_XL_HDR_PT = 10.0
_XL_SCENE_MIN_W = 30.0
_XL_SHEET_NAME = "工作表1"
_IMG_MIN_PX_W = 200
_WD_IMG_H_CM = (4.90, 5.20)
_WD_IMG_W_CM = (3.30, 4.50)
_WD_IMG_RATIO_TOL = 0.08
_XL_IMG_RATIO = (0.55, 0.90)
_XL_IMG_WHITE_BAND = 0.10
_HDR_LINE_MIN = 0.70

# ===== 通用文本工具 =====
def _txt(v):
    """单元格/段落原始值 -> 字符串（None 视为空串），并把全角空格换成半角。"""
    if v is None:
        return ""
    return str(v).replace("　", " ").replace("\xa0", " ")


def _norm(s):
    """归一化：去首尾空白、把连续空白压成一个半角空格。用于“看得见的文字”比对。"""
    return re.sub(r"[ \t\r]+", " ", _txt(s).replace("\n", "\n")).strip()


def _compact(s):
    """压缩：去掉所有空白与常见分隔标点，用于“内容是否齐全”的子串匹配。
    分隔符归一使 593mm,537mm / 593mm/537mm / 593mm、537mm 视为同一内容。"""
    return re.sub(r"[\s：:；;，,、/｜|]+", "", _txt(s))


def _lines(s):
    """按换行拆成非空行列表（单元格内的软回车与硬回车都算行）。"""
    return [ln.strip() for ln in _txt(s).replace("\r", "\n").split("\n") if ln.strip()]


def _char_len(s):
    """字数：去掉空白后的字符数。"""
    return len(re.sub(r"\s+", "", _txt(s)))


def _is_pure_price(s):
    """单元格内只有价格数值与“元”字（允许数字与“元”之间有空格）。"""
    return bool(re.match(r"^\d+(\.\d+)?元$", _compact(s)))


def _param_line_ok(cell_text, idx):
    """核心参数是否逐项分行：行数正确、各行以规定提示词起头、且没有一行里塞进多项参数。"""
    lns = _lines(cell_text)
    if len(lns) != _PARAM_LINES[idx]:
        return False
    for ln in lns:
        head = re.match(r"^([一-鿿]+)\s*[：:]", ln)
        if head is None or head.group(1) not in _PARAM_KEYS:
            return False
        # 一行内出现第 2 个参数提示词 => 用分号/逗号把多项连成了一段
        if sum(1 for k in _PARAM_KEYS if (k + "：") in ln or (k + ":") in ln) > 1:
            return False
    return True


def _param_content_ok(cell_text, idx):
    """核心参数内容是否准确完整：该产品的每个键值对都能在压缩文本里找到。"""
    flat = _compact(cell_text)
    return all((k + v) in flat for k, v in _PARAMS[idx])


def _has_extra_param(cell_text):
    """核心参数里是否出现规定六项之外的参数字段。"""
    flat = _compact(cell_text)
    return any(k in flat for k in _EXTRA_PARAM_KEYS)


def _scene_ok(text, keys):
    """适用场景：整段、字数达标、规定开头、句号结尾、关键词齐全。"""
    t = _norm(text).replace("\n", "")
    if not t.startswith(_SCENE_HEAD):
        return False
    if not t.endswith("。"):
        return False
    if _char_len(t) < _SCENE_MIN_CHARS:
        return False
    return all(k in t for k in keys)


# ===== 图片像素工具 =====
def _load_bgr(data):
    """PNG/JPG 字节 -> (BGR ndarray, 宽, 高)；缺库或解码失败返回 (None, 0, 0)。"""
    if Image is None or np is None:
        return None, 0, 0
    try:
        im = Image.open(io.BytesIO(data)).convert("RGB")
    except Exception:
        return None, 0, 0
    arr = np.array(im)
    return arr[:, :, ::-1].copy(), im.width, im.height


def _gray(bgr):
    if cv2 is not None:
        return cv2.cvtColor(bgr, cv2.COLOR_BGR2GRAY)
    return (0.114 * bgr[:, :, 0] + 0.587 * bgr[:, :, 1]
            + 0.299 * bgr[:, :, 2]).astype("uint8")


def _white_bands(bgr):
    """画面左右两侧纯白空白带宽度占图片宽度的比例。"""
    g = _gray(bgr)
    h, w = g.shape
    colwhite = (g >= 245).mean(0)
    left = 0
    while left < w and colwhite[left] >= 0.98:
        left += 1
    right = 0
    while right < w and colwhite[w - 1 - right] >= 0.98:
        right += 1
    return left / float(w), right / float(w)


def _is_blank_image(bgr):
    """整幅近乎单色（无实际画面内容）时视为空白图。"""
    g = _gray(bgr)
    return bool(g.std() < 6.0)


def _has_promo_badge(bgr):
    """画面内是否存在文字类宣传元素（宣传文案 / 海报标题 / 品牌广告横幅）。

    判别特征：高饱和度色块（S≥0.40、V≥0.30，面积≥画面 0.4%）内部存在白色字形团
    （灰度≥205、团高占色块高 10%–75%、长宽比 0.25–4.0）。纯色块（家具、木纹、
    彩色机身）内部没有白色字形团，故不会误判。

    仅当文字块达到“宣传横幅/海报标题”的规模才判为命中：字形团不少于 12 个，或
    色块横向跨度不小于画面宽度 45%。标注容量、品类的小尺寸规格角标（如画面右上角
    “12.5KG／洗衣机”两行小字）按细则不计入本项，此处即以规模阈值将其排除；把这类
    紧凑角标与同样紧凑的宣传角标（如四字标语）区分开需要文字识别，运行环境无
    OCR，故对紧凑角标一律不判命中。缺 cv2 时无法判定，返回 None（调用处按“未检出”
    处理）。"""
    if cv2 is None or np is None:
        return None
    h, w = bgr.shape[:2]
    hsv = cv2.cvtColor(bgr, cv2.COLOR_BGR2HSV)
    s = hsv[:, :, 1].astype("float32") / 255.0
    v = hsv[:, :, 2].astype("float32") / 255.0
    g = _gray(bgr)
    strong = ((s >= 0.40) & (v >= 0.30)).astype("uint8")
    strong = cv2.morphologyEx(strong, cv2.MORPH_CLOSE, np.ones((5, 5), "uint8"))
    n, lab, stats, _c = cv2.connectedComponentsWithStats(strong, 8)
    for i in range(1, n):
        x, y, bw, bh, area = stats[i]
        if area < 0.004 * h * w or bw < 8 or bh < 6:
            continue
        sub = (lab[y:y + bh, x:x + bw] == i)
        white = ((g[y:y + bh, x:x + bw] >= 205) & sub).astype("uint8")
        m, _l2, st2, _c2 = cv2.connectedComponentsWithStats(white, 8)
        cnt = 0
        for j in range(1, m):
            _gx, _gy, gw, gh, ga = st2[j]
            if ga < 6 or gh < 0.10 * bh or gh > 0.75 * bh:
                continue
            if gw < 0.25 * gh or gw > 4.0 * gh:
                continue
            cnt += 1
        if cnt >= 12 or (cnt >= 3 and bw >= 0.45 * w):
            return True
    return False


# ===== Word XML 工具 =====
# VML 命名空间未在 python-docx 的 nsmap 中注册，直接用完整 URI。
_NS_V = "{urn:schemas-microsoft-com:vml}"


def _twip_cm(v, default=None):
    try:
        return float(v) * _CM_PER_TWIP
    except (TypeError, ValueError):
        return default


def _near(a, b, tol):
    return a is not None and abs(a - b) <= tol


def _cell_borders_ok(cell):
    """单元格四周是否为浅灰(D9D9D9) 0.75 磅细实线。"""
    tcPr = cell._tc.find(qn('w:tcPr'))
    if tcPr is None:
        return False
    bd = tcPr.find(qn('w:tcBorders'))
    if bd is None:
        return False
    for side in ('top', 'left', 'bottom', 'right'):
        el = bd.find(qn('w:' + side))
        if el is None:
            return False
        if el.get(qn('w:val')) != 'single':
            return False
        color = (el.get(qn('w:color')) or '').upper()
        if color != _BORDER_COLOR:
            return False
        try:
            if int(el.get(qn('w:sz'))) != _BORDER_SZ:
                return False
        except (TypeError, ValueError):
            return False
    return True


def _para_font(para):
    """段落内非空 run 的 (东亚字体集合, 字号集合, 加粗集合)。"""
    fonts, sizes, bolds = set(), set(), set()
    for r in para.runs:
        if not (r.text or '').strip():
            continue
        rPr = r._element.rPr
        name = None
        if rPr is not None:
            rf = rPr.find(qn('w:rFonts'))
            if rf is not None:
                name = rf.get(qn('w:eastAsia')) or rf.get(qn('w:ascii'))
        fonts.add(name or (r.font.name or ''))
        sizes.add(round(r.font.size.pt, 1) if r.font.size else None)
        bolds.add(bool(r.font.bold))
    return fonts, sizes, bolds


def _is_songti(fonts):
    return bool(fonts) and all(any(k in (f or '') for k in _SONGTI) for f in fonts)


def _spacing_pt(para):
    """段前/段后间距（磅）；未设置时按 0 处理（Word 默认继承值另行判断）。"""
    pPr = para._p.find(qn('w:pPr'))
    if pPr is None:
        return 0.0, 0.0
    sp = pPr.find(qn('w:spacing'))
    if sp is None:
        return 0.0, 0.0
    before = sp.get(qn('w:before'))
    after = sp.get(qn('w:after'))
    b = (float(before) / 20.0) if before is not None else 0.0
    a = (float(after) / 20.0) if after is not None else 0.0
    return b, a


def _is_centered(para):
    """段落是否水平居中（含未显式设置但样式为居中的情形按显式值判定）。"""
    pPr = para._p.find(qn('w:pPr'))
    if pPr is not None:
        jc = pPr.find(qn('w:jc'))
        if jc is not None:
            return jc.get(qn('w:val')) in ('center', 'both-center')
    return False


def _inline_images(cell, rels):
    """单元格内图片的 [(显示宽cm, 显示高cm, 像素宽高比 或 None)]。"""
    out = []
    for para in cell.paragraphs:
        for drawing in para._p.findall('.//' + qn('w:drawing')):
            ext = drawing.find('.//' + qn('wp:extent'))
            if ext is None:
                continue
            wcm = int(ext.get('cx')) * _CM_PER_EMU
            hcm = int(ext.get('cy')) * _CM_PER_EMU
            ratio = None
            blip = drawing.find('.//' + qn('a:blip'))
            if blip is not None:
                rid = blip.get(qn('r:embed'))
                data = rels.get(rid)
                if data:
                    _bgr, pw, ph = _load_bgr(data)
                    if ph:
                        ratio = pw / float(ph)
            out.append((wcm, hcm, ratio))
    return out


def _xmlstr(el):
    """取任意 XML 元素的字符串形式（未注册的 VML 元素没有 .xml 属性）。"""
    if el is None:
        return ""
    try:
        return el.xml
    except AttributeError:
        try:
            from lxml import etree
            return etree.tostring(el, encoding='unicode')
        except Exception:
            return ""


def _footer_xml_element(footer):
    """读取 footer part 的原始 XML，兼容 python-docx 的通用 Part 对象。"""
    try:
        from lxml import etree
        blob = getattr(footer.part, "blob", b"")
        return etree.fromstring(blob) if blob else None
    except Exception:
        return None


def _rels_map(doc):
    """rId -> 图片字节；用于把单元格里的图片引用还原成像素。"""
    out = {}
    try:
        for rid, part in doc.part.related_parts.items():
            blob = getattr(part, 'blob', None)
            if blob:
                out[rid] = blob
    except Exception:
        pass
    return out


def _header_line_ok(section):
    """页眉内是否水平居中显示“产品分析”，且文字下方有覆盖正文宽度 70% 以上的横线。

    段落下框线(w:pBdr/w:bottom)天然横贯整个段落宽度，即正文宽度的 100%，满足
    “覆盖 70% 以上”；此外也接受用横线形状(v:line/v:rect)画出的页眉线。"""
    parts = [section.header]
    try:
        if section.different_first_page_header_footer:
            parts.append(section.first_page_header)
    except Exception:
        pass
    for part in parts:
        header_xml = _footer_xml_element(part)
        text_ok = False
        line_ok = False
        for para in (header_xml.findall('.//' + qn('w:p'))
                     if header_xml is not None else []):
            text = "".join(t.text or "" for t in para.findall('.//' + qn('w:t')))
            ppr = para.find(qn('w:pPr'))
            jc = ppr.find(qn('w:jc')) if ppr is not None else None
            centered = jc is not None and jc.get(qn('w:val')) in (
                'center', 'both', 'distribute'
            )
            if "产品分析" in _norm(text) and centered:
                text_ok = True
            if ppr is not None:
                bdr = ppr.find(qn('w:pBdr'))
                if bdr is not None:
                    bot = bdr.find(qn('w:bottom'))
                    if bot is not None and bot.get(qn('w:val')) not in (None, 'none', 'nil'):
                        line_ok = True
            xml = _xmlstr(para)
            if ('v:line' in xml) or ('v:rect' in xml):
                line_ok = True
        if not (text_ok and line_ok):
            return False
    return True


def _est_text_pt(text, size_pt):
    """按“中文/全角=1 字宽、半角=0.5 字宽”估算一段文字的渲染宽度（磅）。"""
    w = 0.0
    for ch in text:
        if ord(ch) > 0x2E80:
            w += size_pt
        else:
            w += size_pt / 2.0
    return w


def _footer_page_no_ok(section, page_w_pt, text_w_pt, left_margin_pt):
    """页脚内是否有水平居中的页码域，且自第 1 页起连续编号。

    两种落地方式都接受：
      ① 页码域所在段落本身水平居中；
      ② 页码域放在页脚浮动文本框里——此时按文本框的水平偏移与框内文字的估算
         宽度算出页码文字的水平中心，与页面水平中心相差不超过正文宽度的 8%
         即视为居中（纯 XML 无法取得渲染坐标，此为等效几何判定）。
    连续编号：使用 PAGE 域即为自动连续；若显式指定起始页码，须为 1。"""
    footer = section.footer
    footer_xml = _footer_xml_element(footer)
    xml = _xmlstr(footer_xml)
    if 'PAGE' not in xml:
        return False
    # 起始页码
    sectPr = section._sectPr
    pn = sectPr.find(qn('w:pgNumType'))
    if pn is not None:
        start = pn.get(qn('w:start'))
        if start is not None and start.strip() not in ('1',):
            return False
    # ① 页码域直接在居中段落里
    for para in (footer_xml.findall('.//' + qn('w:p'))
                 if footer_xml is not None else []):
        pxml = _xmlstr(para)
        ppr = para.find(qn('w:pPr'))
        jc = ppr.find(qn('w:jc')) if ppr is not None else None
        centered = jc is not None and jc.get(qn('w:val')) in (
            'center', 'both', 'distribute'
        )
        if 'PAGE' in pxml and centered and 'v:shape' not in pxml:
            return True
    # ② 页码域在浮动文本框里：按几何位置判定
    for shape in (footer_xml.findall('.//' + _NS_V + 'shape')
                  if footer_xml is not None else []):
        if 'PAGE' not in _xmlstr(shape):
            continue
        style = shape.get('style') or ''
        st = dict(re.findall(r'([a-zA-Z\-]+)\s*:\s*([^;]+)', style))

        def _pt(key):
            m = re.match(r'^(-?[\d.]+)pt$', (st.get(key) or '').strip())
            return float(m.group(1)) if m else None

        box_w = _pt('width')
        left = _pt('margin-left')
        if left is None:
            left = _pt('left')
        if left is None or box_w is None:
            continue
        if 'margin' in (st.get('mso-position-horizontal-relative') or 'margin'):
            left += left_margin_pt
        if (st.get('mso-position-horizontal') or '').strip() == 'center':
            return True
        # 框内文字：域结果按 1 位数字估算
        inner_txt = "".join(t.text or '' for t in shape.findall('.//' + qn('w:t')))
        sz = 21.0
        szel = shape.find('.//' + qn('w:sz'))
        if szel is not None:
            try:
                sz = float(szel.get(qn('w:val')))
            except (TypeError, ValueError):
                sz = 21.0
        size_pt = sz / 2.0
        inner_para_centered = any(
            (jc.get(qn('w:val')) == 'center')
            for jc in shape.findall('.//' + qn('w:jc')))
        if inner_para_centered:
            cx = left + box_w / 2.0
        else:
            cx = left + _est_text_pt(inner_txt, size_pt) / 2.0
        if abs(cx - page_w_pt / 2.0) <= 0.08 * text_w_pt:
            return True
    return False


# ===== Excel 结构工具 =====
def _merge_of(ws, col, row):
    """返回覆盖 (col,row) 的合并区域 (min_row, max_row)；未合并返回 (row, row)。"""
    for rng in ws.merged_cells.ranges:
        if rng.min_col <= col <= rng.max_col and rng.min_row <= row <= rng.max_row:
            return rng.min_row, rng.max_row
    return row, row


def _block_value(ws, col, top, bottom):
    """取区块内某列的内容：优先区块首行，其次向下找第一个非空。"""
    for r in range(top, bottom + 1):
        v = ws.cell(row=r, column=col).value
        if v is not None and _norm(v) != "":
            return _norm(v)
    return ""


def _header_cols(ws):
    """按第一行表头文字定位五列列号；缺失的表头回落到 A–E 的固定次序。"""
    found = {}
    for c in range(1, min(ws.max_column, 30) + 1):
        t = _norm(ws.cell(row=1, column=c).value)
        if t in _HEADERS and t not in found:
            found[t] = c
    for i, name in enumerate(_HEADERS):
        found.setdefault(name, i + 1)
    return found


def _product_blocks(ws, name_col):
    """按“产品名称”列的非空锚点行切分产品数据区块，返回 [(top, bottom)]。
    合并单元格区块与单行区块都能切出：锚点即区块首行，区块延伸到下一个锚点前一行。"""
    anchors = []
    for r in range(2, ws.max_row + 1):
        v = ws.cell(row=r, column=name_col).value
        if v is not None and _norm(v) != "":
            anchors.append(r)
    blocks = []
    for i, top in enumerate(anchors):
        nxt = anchors[i + 1] if i + 1 < len(anchors) else ws.max_row + 1
        m_top, m_bottom = _merge_of(ws, name_col, top)
        bottom = max(m_bottom, nxt - 1)
        blocks.append((top, min(bottom, nxt - 1)))
    return blocks


def _xl_media(file_path):
    """工作簿内所有展示图的字节（嵌入单元格图片与浮动图片都在 xl/media 下）。"""
    out = []
    try:
        with zipfile.ZipFile(file_path) as zf:
            for n in sorted(zf.namelist()):
                if n.startswith('xl/media/') and not n.endswith('/'):
                    out.append((n.split('/')[-1], zf.read(n)))
    except Exception:
        pass
    return out


def _xl_cell_images(file_path):
    """是否使用“嵌入单元格”方式放图：存在 xl/cellimages.xml 图片仓。"""
    try:
        with zipfile.ZipFile(file_path) as zf:
            return 'xl/cellimages.xml' in zf.namelist()
    except Exception:
        return False


def _dispimg_rows(ws, col):
    """展示图列中以“嵌入单元格图片”落地的行号（单元格公式为 DISPIMG）。"""
    rows = []
    for r in range(2, ws.max_row + 1):
        v = ws.cell(row=r, column=col).value
        if isinstance(v, str) and 'DISPIMG' in v.upper():
            rows.append(r)
    return rows


def _float_img_rows(ws):
    """浮动图片（叠放在单元格上方）的 (行号, 列号) 列表。"""
    out = []
    for img in getattr(ws, '_images', []) or []:
        anc = getattr(img, 'anchor', None)
        frm = getattr(anc, '_from', None)
        if frm is None:
            continue
        out.append((frm.row + 1, frm.col + 1))
    return out


def _word_fit(path):
    """Word 候选文件与“产品摘要单页表格文档”这一交付物形态的契合度打分。
    目录内可能同时存在其它 .docx（如各产品的说明文档），只靠文件名不足以定位，
    因此按正文特征打分：标题“产品摘要”、“产品一/二/三”表格、四个字段提示词。"""
    score = 0
    try:
        with zipfile.ZipFile(path) as z:
            xml = z.read('word/document.xml').decode('utf-8', 'ignore')
    except Exception:
        return -1
    body = re.sub(r'<[^>]+>', '', xml)
    if '产品摘要' in body:
        score += 4
    if '产品一' in body and '产品二' in body and '产品三' in body:
        score += 3
    if '产品名称' in body and '核心参数' in body and '指导价格' in body:
        score += 2
    if xml.count('<w:tbl>') == 1:
        score += 1
    name = os.path.basename(path)
    if '产品分析' in name:
        score += 2
    return score


def _xl_fit(path):
    """Excel 候选文件与“选货汇总表”这一交付物形态的契合度打分。"""
    score = 0
    try:
        wb = load_workbook(path, read_only=True)
    except Exception:
        return -1
    try:
        ws = wb.worksheets[0] if wb.worksheets else None
        head = ""
        if ws is not None:
            for row in ws.iter_rows(min_row=1, max_row=1, max_col=8):
                head = "".join(_txt(c.value) for c in row if c.value is not None)
                break
        score += sum(2 for h in _HEADERS if h in head)
    finally:
        try:
            wb.close()
        except Exception:
            pass
    if '选货汇总表' in os.path.basename(path):
        score += 1
    return score


def _pick_files(dir_path):
    """在目录内定位两个交付物：Word 文档与 Excel 工作簿。
    返回 (word 路径, excel 路径, 格式不被接受的候选文件名列表)。"""
    files = [f for f in sorted(os.listdir(dir_path))
             if not f.startswith('~$') and not f.startswith('.')]
    wd = [f for f in files if f.lower().endswith(_EXT_WORD_OK)]
    xl = [f for f in files if f.lower().endswith(_EXT_XL_OK)]
    rej = [f for f in files if f.lower().endswith(_EXT_WORD_REJECT + _EXT_XL_REJECT)]
    wp = xp = ""
    if wd:
        wp = os.path.join(dir_path, max(
            wd, key=lambda f: (_word_fit(os.path.join(dir_path, f)), -wd.index(f))))
    if xl:
        xp = os.path.join(dir_path, max(
            xl, key=lambda f: (_xl_fit(os.path.join(dir_path, f)), -xl.index(f))))
    return wp, xp, rej


# add() 登记的字符串尾部常带一段“（实测值…）”，那是本地排查用的实测摘要，不写入
# 返回信息：返回的 rule 只保留规则原文，命中与未命中完全一致，不随判定结果变化。
# 只剥离结尾处、内部不再嵌套括号的一段全角括号，规则原文中部的括号不受影响。
_RULE_ONLY = re.compile(r"（[^（）]*）$")


def evaluate(dir_path):
    """按接口约定返回结构化 dict；不 print、不改 sys.stdout、不 sys.exit。
    入参 dir_path 为“交付物所在目录的路径”，脚本自己在该目录内定位并打开被评估的
    Word 文档与 Excel 工作簿（忽略 ~$ 开头的临时锁文件）。

    返回信息约定：
      status="error"（目录不存在、目录中没有任何交付文件、脚本自身异常）表示这份
      作业根本没评上，此时 dim1_pass 为空串（未判定）、dim1_reason 为空串、
      dim2_items 为 []、total_score 与 max_score 均为 0；
      维度一未通过时 status 仍为 "ok"，dim1_pass 为 False、dim1_reason 写错误摘要、
      dim2_items 为 []（未真正逐条评过）、total_score 为 0、max_score 为满分；
      维度一通过时 dim1_pass 为 True、dim1_reason 保持空串。
      dim2_items 每条只含 rule / max_delta / delta / hit 四个键。"""
    result = {
        "id": SCRIPT_ID,
        "file_name": "",
        "status": "ok",
        "error": None,
        "dim1_pass": "",
        "dim1_reason": "",
        "dim2_items": [],
        "total_score": 0,
        "max_score": 0,
    }
    try:
        if not dir_path or not os.path.isdir(dir_path):
            result["status"] = "error"
            result["error"] = "目录不存在：%s" % dir_path
            return result
        word_path, xl_path, rejects = _pick_files(dir_path)
        result["file_name"] = " | ".join(
            os.path.basename(p) for p in (word_path, xl_path) if p)
        if not word_path and not xl_path and not rejects:
            # 目录里一个交付物都没有：这份作业没评上，而不是维度一判不通过
            result["status"] = "error"
            result["error"] = "目录中未找到 .docx/.xlsx 文件：%s" % dir_path
            return result
        if not word_path or not xl_path:
            # 交付物残缺或格式不被接受：已有交付动作，交给维度一记 0
            miss = []
            if not word_path:
                miss.append("Word 文档（接受 .docx）")
            if not xl_path:
                miss.append("Excel 工作簿（接受 .xlsx/.xlsm）")
            reason = "缺少交付物：" + "、".join(miss)
            if rejects:
                reason += "；目录内格式不被接受的文件：" + "、".join(rejects)
            return _dim1_fail(result, reason)
        return _evaluate_impl(word_path, xl_path, result)
    except Exception as exc:
        # 脚本异常同样视为“没评上”：维度一未判定，两项分数归零。
        result["status"] = "error"
        result["error"] = "%s: %s" % (type(exc).__name__, exc)
        result["dim1_pass"] = ""
        result["dim1_reason"] = ""
        result["dim2_items"] = []
        result["total_score"] = 0
        result["max_score"] = 0
        return result


def _dim1_fail(result, reason):
    """维度一未通过：完成度总分记 0，不再逐条罗列（未真正逐条评过）。"""
    result["dim1_pass"] = False
    result["dim1_reason"] = reason or "维度一未通过（未给出具体原因）"
    result["dim2_items"] = []
    result["total_score"] = 0
    result["max_score"] = MAX_SCORE_TOTAL
    return result


def _dim1(word_path, xl_path, result):
    """维度一：两个交付物齐备、格式在接受范围内、可正常打开、内容为可编辑文本。
    返回 (doc, wb) 或 (None, None)（此时 result 已填好不通过原因）。"""
    def _fail(reason):
        _dim1_fail(result, reason)
        return None, None

    try:
        doc = Document(word_path)
    except Exception as exc:
        return _fail("Word 文档无法打开：%s: %s" % (type(exc).__name__, exc))
    try:
        wb = load_workbook(xl_path)
    except Exception as exc:
        return _fail("Excel 工作簿无法打开：%s: %s" % (type(exc).__name__, exc))
    if not wb.worksheets:
        return _fail("Excel 工作簿内没有工作表")
    # 内容可编辑：文档正文/表格里要有可编辑文本，而不是整页截图
    doc_text = "".join(p.text for p in doc.paragraphs)
    for t in doc.tables:
        for row in t.rows:
            for c in row.cells:
                doc_text += c.text
    if _char_len(doc_text) < 4:
        return _fail("Word 文档内没有可编辑文本（疑为图片/扫描件/整页截图）")
    ws = wb.worksheets[0]
    cell_text = ""
    for row in ws.iter_rows(min_row=1, max_row=min(ws.max_row, 200)):
        for c in row:
            if c.value is not None:
                cell_text += _txt(c.value)
    if _char_len(cell_text) < 4:
        return _fail("Excel 工作表单元格内没有可编辑文本（疑为图片/扫描件/整页截图）")
    result["dim1_pass"] = True
    # 约定：维度一通过时 dim1_reason 保持空串，只有未通过才写错误摘要。
    result["dim1_reason"] = ""
    return doc, wb


_PROD_KEYS = ("云溪", "统帅", "小天鹅", "米家", "美的")


def _prod_index(text):
    """按品牌/系列关键词判断这条内容属于哪一款产品；判不出返回 -1。"""
    flat = _compact(text)
    for i, k in enumerate(_PROD_KEYS):
        if k in flat:
            return i
    return -1


def _cell_fields(cell):
    """把 Word 左列单元格拆成 {字段提示词: [值行...]}，另用 '__prod__' 存产品序号行。
    字段值可以续写在后面的独立段落里（核心参数逐项分行就是这种写法）。"""
    fields = {}
    cur = None
    for para in cell.paragraphs:
        t = _norm(para.text)
        if not t:
            continue
        if re.match(r'^产品[一二三四五六七八九十]+[：:]$', t):
            cur = '__prod__'
            fields.setdefault(cur, []).append(t)
            continue
        matched = None
        for lab in _FIELD_LABELS:
            if t.startswith(lab):
                matched = lab
                break
        if matched:
            cur = matched
            fields.setdefault(cur, [])
            rest = t[len(matched):].strip()
            if rest:
                fields[cur].append(rest)
        elif cur:
            fields[cur].append(t)
    return fields


def _cell_run_groups(cell):
    """把左列单元格的 run 按角色分组：产品序号行、字段提示词、填入的信息值。"""
    prod_runs, label_runs, value_runs = [], [], []
    for para in cell.paragraphs:
        t = _norm(para.text)
        if not t:
            continue
        if re.match(r'^产品[一二三四五六七八九十]+[：:]$', t):
            prod_runs.extend(r for r in para.runs if (r.text or '').strip())
            continue
        cut = 0
        for lab in _FIELD_LABELS:
            if para.text.lstrip().startswith(lab):
                cut = len(lab) + (len(para.text) - len(para.text.lstrip()))
                break
        pos = 0
        for r in para.runs:
            length = len(r.text or '')
            if (r.text or '').strip():
                (label_runs if pos < cut else value_runs).append(r)
            pos += length
    return prod_runs, label_runs, value_runs


def _runs_font_ok(runs, pt, bold):
    """一组 run 是否都为宋体、指定字号、指定加粗状态；空组视为合规（无内容可判）。"""
    if not runs:
        return True
    for r in runs:
        rPr = r._element.rPr
        name = None
        if rPr is not None:
            rf = rPr.find(qn('w:rFonts'))
            if rf is not None:
                name = rf.get(qn('w:eastAsia')) or rf.get(qn('w:ascii'))
        name = name or (r.font.name or '')
        if not any(k in name for k in _SONGTI):
            return False
        size = r.font.size.pt if r.font.size else None
        if size is None or abs(size - pt) > 0.3:
            return False
        if bool(r.font.bold) != bold:
            return False
    return True


def _evaluate_impl(word_path, xl_path, result):
    """维度一 + 维度二的评估主体；按细则顺序追加 results，末尾反解为 dim2_items。"""
    doc, wb = _dim1(word_path, xl_path, result)
    if doc is None:
        return result

    results = []

    def add(hit, delta, rule):
        # rule 尾部允许带“（实测值…）”，汇总时会剥掉，只把规则原文写进返回信息。
        results.append("%+d：%s" % ((delta if hit else 0), rule))

    # ================= Excel 侧数据准备 =================
    ws = wb.worksheets[0]
    cols = _header_cols(ws)
    c_name, c_par, c_scene, c_price, c_img = (cols[h] for h in _HEADERS)
    blocks = _product_blocks(ws, c_name)
    xl_names = [_block_value(ws, c_name, t, b) for t, b in blocks]
    xl_pars = [_block_value(ws, c_par, t, b) for t, b in blocks]
    xl_scenes = [_block_value(ws, c_scene, t, b) for t, b in blocks]
    xl_prices = [_block_value(ws, c_price, t, b) for t, b in blocks]
    xl_idx = [_prod_index(n) for n in xl_names]
    five_blocks = (len(blocks) == 5 and sorted(xl_idx) == [0, 1, 2, 3, 4])

    media = _xl_media(xl_path)
    disp_rows = _dispimg_rows(ws, c_img)
    float_in_col = [r for (r, c) in _float_img_rows(ws) if c == c_img]
    img_rows = sorted(disp_rows + float_in_col)

    def _imgs_per_block():
        return [sum(1 for r in img_rows if t <= r <= b) for t, b in blocks]

    # ---- 01 Excel 产品清单 ----
    hit = ([_compact(n) for n in xl_names] == [_compact(n) for n in _NAMES])
    add(hit, 3, "Excel 第一行以下共 5 个产品数据区块，“产品名称”列自上而下依次为"
                "“云溪5.0 D3滚筒洗衣机”“统帅三桶懒人洗衣机”“小天鹅12KG滚筒洗衣机”"
                "“米家8KG滚筒洗衣机”“美的8KG波轮洗衣机”，5 项齐全、无重复、无多余产品")

    # ---- 02 Excel 指导价格 ----
    hit = (len(xl_prices) == 5
           and all(_compact(a) == _compact(b) for a, b in zip(xl_prices, _PRICES))
           and all(_is_pure_price(v) for v in xl_prices))
    add(hit, 3, "Excel“指导价格”列 5 个值自上而下依次为“4500元”“3999元”“3039元”"
                "“999元”“846元”，单元格内只有价格数值与“元”字，不含说明性文字")

    # ---- 03 Excel 核心参数内容 ----
    hit = five_blocks and all(_param_content_ok(xl_pars[i], xl_idx[i])
                              for i in range(len(blocks)))
    add(hit, 5, "Excel“核心参数”列 5 条内容分别准确完整（洗涤容量、宽度、类型、"
                "能效等级、深度，云溪另含高度 850mm）")

    # ---- 04 Excel 核心参数逐项分行 ----
    hit = five_blocks and all(_param_line_ok(xl_pars[i], xl_idx[i])
                              for i in range(len(blocks)))
    add(hit, 3, "Excel“核心参数”单元格内每项参数各占一行（云溪 6 行、其余 4 款各 5 行），"
                "各行分别以“洗涤容量：”“宽度：”“类型：”“能效等级：”“深度：”起头，"
                "不出现用分号或逗号把多项参数连成一整段的写法")

    # ---- 05 Excel 适用场景 ----
    hit = five_blocks and all(_scene_ok(xl_scenes[i], _SCENE_KEYS_XL[xl_idx[i]])
                              for i in range(len(blocks)))
    add(hit, 3, "Excel“适用场景”列 5 条均为整段场景描述，字数均不少于 60 字，均以"
                "“本产品适配多种场景，核心适用于”开头并以句号结尾，各条场景关键词齐全")

    # ---- 06 Excel 自动换行 + 垂直居中 ----
    ok = bool(blocks)
    for (t, _b) in blocks:
        for c in (c_par, c_scene):
            al = ws.cell(row=t, column=c).alignment
            if not (al.wrap_text and al.vertical == 'center'):
                ok = False
    add(ok, 1, "Excel 5 个产品区块的“核心参数”“适用场景”单元格均开启自动换行且垂直居中，"
               "长文本在本单元格内完整折行显示，不溢出单元格、不遮挡相邻列")

    # ---- 07 Excel 适用场景列宽 ----
    dim = ws.column_dimensions.get(get_column_letter(c_scene))
    w_scene = dim.width if (dim is not None and dim.width) else None
    add(bool(w_scene and w_scene >= _XL_SCENE_MIN_W), 1,
        "Excel“适用场景”列列宽不小于 30 字符，整段场景描述可在本列内完整换行显示")

    # ---- 08 Excel 展示图张数与产品对应 ----
    # 【可机检边界】“图片机型与产品名称一致”需人眼语义判断，此处判定其可机检部分：
    # 每个产品区块的展示图列内恰好 1 张图、共 5 张、5 张互不相同且都不是空白画面。
    per_block = _imgs_per_block()
    distinct = len({bytes(b) for _n, b in media})
    blank_free = True
    for _n, b in media:
        bgr, _w, _h = _load_bgr(b)
        if bgr is None or _is_blank_image(bgr):
            blank_free = False
    hit = (len(blocks) == 5 and per_block == [1] * 5 and len(media) == 5
           and distinct == 5 and blank_free)
    add(hit, 5, "Excel“展示图”列 5 个产品区块各有 1 张洗衣机图片、共 5 张，"
                "且每张图片与同区块“产品名称”对应（云溪、统帅、小天鹅、米家、美的各 1 张）")

    # ---- 09 Excel 展示图无文字类宣传元素 ----
    promo_text = any(w in _compact("".join(xl_names + xl_pars + xl_scenes))
                     for w in _PROMO_WORDS)
    promo_img = False
    for _n, b in media:
        bgr, _w, _h = _load_bgr(b)
        if bgr is None:
            continue
        if _has_promo_badge(bgr) is True:
            promo_img = True
    add(bool(media) and not (promo_img or promo_text), 3,
        "Excel 5 张展示图画面内均无文字类宣传元素，不出现“巅峰之作”“AI全能旗舰 高端净护选直驱”"
        "“一级健康洁净标准”“10年包修”“超薄省空间”等宣传文案、海报标题与品牌广告角标；"
        "标注容量、品类的规格角标不计入本项")

    # ---- 10 Excel 展示图机身完整 ----
    # 【可机检边界】“四边未被画面边界截断”无可靠像素代理指标（白底白机身使边缘着墨、
    # 背景差分、暗区占比在样本上均无区分度），此处判定同一评分项内可机检的两个条件：
    # 像素宽高比 0.55–0.90，画面左右两侧纯白空白带各不超过图片宽度的 10%。
    ok = bool(media)
    for _n, b in media:
        bgr, pw, ph = _load_bgr(b)
        if bgr is None or not ph:
            ok = False
            continue
        ratio = pw / float(ph)
        if not (_XL_IMG_RATIO[0] <= ratio <= _XL_IMG_RATIO[1]):
            ok = False
        wl, wr = _white_bands(bgr)
        if wl > _XL_IMG_WHITE_BAND or wr > _XL_IMG_WHITE_BAND:
            ok = False
    add(ok, 3, "Excel 5 张展示图中洗衣机机身完整，上下左右四边均未被画面边界截断，"
               "图片自身像素宽高比在 0.55 至 0.90 之间，画面左右两侧的纯白空白带宽度"
               "各不超过图片宽度的 10%")

    # ---- 11 Excel 展示图嵌入单元格 ----
    hit = (_xl_cell_images(xl_path) and len(disp_rows) == 5 and not float_in_col
           and len(blocks) == 5
           and [sum(1 for r in disp_rows if t <= r <= b) for t, b in blocks] == [1] * 5)
    add(hit, 3, "Excel 5 张展示图均嵌入单元格显示、随行高列宽缩放，图片完整落在“展示图”"
                "列内，高度不超过所在区块总行高、宽度不超过该列列宽，未纵向压到下一产品"
                "区块、未横向越出表格右边界")

    # ---- 12 Excel 区块行数一致且五项横向对齐 ----
    sizes = [b - t + 1 for t, b in blocks]
    aligned = True
    for (t, b) in blocks:
        for c in (c_name, c_par, c_scene, c_price):
            if _merge_of(ws, c, t)[0] != t or not _block_value(ws, c, t, b):
                aligned = False
        if sum(1 for r in img_rows if t <= r <= b) != 1:
            aligned = False
    hit = (len(blocks) == 5 and len(set(sizes)) == 1 and aligned)
    add(hit, 3, "Excel 5 个产品各占一个数据区块，5 个区块所含行数完全相同，"
                "“产品名称”“核心参数”“适用场景”“指导价格”“展示图”五项在同一区块内横向对齐")

    # ================= Word 侧数据准备 =================
    rels = _rels_map(doc)
    tbl = doc.tables[0] if doc.tables else None
    wd_rows = list(tbl.rows) if tbl is not None else []
    wd_fields, wd_imgs, wd_left, wd_right = [], [], [], []
    for row in wd_rows:
        cells = row.cells
        left = cells[0] if len(cells) > 0 else None
        right = cells[-1] if len(cells) > 1 else None
        wd_left.append(left)
        wd_right.append(right)
        wd_fields.append(_cell_fields(left) if left is not None else {})
        wd_imgs.append(_inline_images(right, rels) if right is not None else [])

    def _fv(i, lab):
        """第 i 行某字段的值（多行合并为一段）。"""
        if i >= len(wd_fields):
            return ""
        return " ".join(wd_fields[i].get(lab, []))

    def _fl(i, lab):
        """第 i 行某字段的值行列表。"""
        if i >= len(wd_fields):
            return []
        return wd_fields[i].get(lab, [])

    wd_idx = [_prod_index(_fv(i, "产品名称：")) for i in range(len(wd_rows))]
    three_rows = (len(wd_rows) == 3)

    # ---- 13 Word 按指导价格从高到低排列 ----
    hit = (three_rows and wd_idx == [0, 1, 2]
           and all(_compact(_fv(i, "指导价格：")) == _compact(_PRICES[i])
                   for i in range(3)))
    add(hit, 5, "Word 表格 3 行自上而下按指导价格从高到低排列，第 1 行为云溪机型且指导价格"
                "4500 元、第 2 行为统帅机型且指导价格 3999 元、第 3 行为小天鹅机型且指导价格"
                "3039 元")

    # ---- 14 Word 三行产品名称 ----
    hit = (three_rows and all(_compact(_fv(i, "产品名称：")) == _compact(_NAMES[i])
                              for i in range(3)))
    add(hit, 1, "Word 三行“产品名称：”后的内容分别为“云溪5.0 D3滚筒洗衣机”"
                "“统帅三桶懒人洗衣机”“小天鹅12KG滚筒洗衣机”")

    # ---- 15 Word 三行核心参数 + 逐项分行 ----
    ok = three_rows
    for i in range(min(3, len(wd_rows))):
        text = "\n".join(_fl(i, "核心参数："))
        if not (_param_content_ok(text, i) and _param_line_ok(text, i)):
            ok = False
    add(ok, 3, "Word 三行“核心参数：”后内容分别准确完整，且每项参数各占一行"
               "（云溪 6 行，统帅、小天鹅各 5 行）")

    # ---- 16 Word 三行适用场景 ----
    ok = three_rows
    for i in range(min(3, len(wd_rows))):
        if not _scene_ok(_fv(i, "适用场景："), _SCENE_KEYS_WD[i]):
            ok = False
    add(ok, 3, "Word 三行“适用场景：”后均为整段场景描述，字数均不少于 60 字，均以"
               "“本产品适配多种场景，核心适用于”开头并以句号结尾，各条场景关键词齐全")

    # ---- 17 Word 三行指导价格 ----
    hit = (three_rows
           and all(_compact(_fv(i, "指导价格：")) == _compact(_PRICES[i])
                   and _is_pure_price(_fv(i, "指导价格：")) for i in range(3)))
    add(hit, 1, "Word 三行“指导价格：”后的内容分别为“4500元”“3999元”“3039元”，"
                "只有价格数值与“元”字，不含促销、补贴等说明文字")

    # ---- 18 Word 右列图片张数与产品对应 ----
    # 【可机检边界】“图片机型与同行产品名称一致”需人眼语义判断，此处判定其可机检部分：
    # 右列 3 个单元格各 1 张、共 3 张、3 张互不相同且都不是空白画面。
    wd_media = []
    if tbl is not None:
        for right in wd_right:
            for para in (right.paragraphs if right is not None else []):
                for blip in para._p.findall('.//' + qn('a:blip')):
                    data = rels.get(blip.get(qn('r:embed')))
                    if data:
                        wd_media.append(data)
    blank_free = True
    for b in wd_media:
        bgr, _w, _h = _load_bgr(b)
        if bgr is None or _is_blank_image(bgr):
            blank_free = False
    hit = (three_rows and [len(x) for x in wd_imgs] == [1, 1, 1]
           and len({bytes(b) for b in wd_media}) == 3 and blank_free)
    add(hit, 5, "Word 表格右列 3 个单元格各有 1 张洗衣机图片、共 3 张，"
                "且每张图片与同行左列的产品名称对应")

    # ---- 19 Word 右列图片显示尺寸与不变形 ----
    ok = three_rows and all(len(x) == 1 for x in wd_imgs)
    for imgs in wd_imgs:
        for (wcm, hcm, ratio) in imgs:
            if not (_WD_IMG_H_CM[0] <= hcm <= _WD_IMG_H_CM[1]):
                ok = False
            if not (_WD_IMG_W_CM[0] <= wcm <= _WD_IMG_W_CM[1]):
                ok = False
            if ratio is None or hcm <= 0:
                ok = False
            elif abs((wcm / hcm) - ratio) > _WD_IMG_RATIO_TOL:
                ok = False
    add(ok, 3, "Word 右列 3 张图片显示高度均在 4.90 至 5.20 厘米之间、显示宽度均在"
               "3.30 至 4.50 厘米之间，且显示宽高比与图片自身像素宽高比的偏差不超过 0.08")

    # ---- 20 Word 页眉 ----
    sec = doc.sections[0]
    add(_header_line_ok(sec), 3,
        "Word 每页页眉内水平居中显示“产品分析”，页眉文字下方有一条居中横线，"
        "横线长度覆盖正文宽度 70% 以上")

    # ---- 21 Word 页脚页码 ----
    page_w_pt = (sec.page_width.pt if sec.page_width else 595.3)
    left_pt = (sec.left_margin.pt if sec.left_margin else 56.7)
    right_pt = (sec.right_margin.pt if sec.right_margin else 56.7)
    add(_footer_page_no_ok(sec, page_w_pt, page_w_pt - left_pt - right_pt, left_pt), 1,
        "Word 页脚内有水平居中的页码域，自第 1 页起连续编号，不跳号、不重复")

    # ---- 22 Word 文件名 ----
    base = os.path.splitext(os.path.basename(word_path))[0]
    stem = re.sub(r'[\s(（][^()（）]*[)）]?\s*$', '', base).strip() or base
    hit = ("产品分析" in base
           and not any(k in base for k in ("模板", "副本", "未命名", "Copy", "copy"))
           and stem.replace(" ", "") == "产品分析")
    add(hit, 1, "交付 Word 文件名主体为“产品分析”，不含“模板”“副本”“未命名”等字样")

    # ================= 减分项：Word 固有格式是否被破坏 =================
    all_cells = []
    for row in wd_rows:
        all_cells.extend(row.cells)

    # ---- 23 单元格边框 ----
    ok = bool(all_cells) and all(_cell_borders_ok(c) for c in all_cells)
    add(not ok, -3, "Word 表格 3 行 2 列共 6 个单元格四周的边框不是浅灰色（D9D9D9）"
                    "0.75 磅细实线")

    # ---- 24 页边距与页眉页脚距边界 ----
    def _cm(v):
        return v.cm if v is not None else None

    ok = (all(_near(_cm(x), _MARGIN_CM, 0.05) for x in
              (sec.top_margin, sec.bottom_margin, sec.left_margin, sec.right_margin))
          and _near(_cm(sec.header_distance), _HDR_FTR_CM, 0.05)
          and _near(_cm(sec.footer_distance), _HDR_FTR_CM, 0.05))
    add(not ok, -3, "Word 页面上、下、左、右页边距不为 2.0 厘米，或页眉、页脚距页边界"
                    "不为 0 厘米")

    # ---- 25/26/27 标题“产品摘要” ----
    title_para = None
    for para in doc.paragraphs:
        if _norm(para.text):
            title_para = para
            break
    t_fonts, t_sizes, t_bolds = _para_font(title_para) if title_para is not None else \
        (set(), set(), set())
    ok = (title_para is not None and _norm(title_para.text) == _TITLE_TEXT
          and t_sizes and all(s is not None and abs(s - _TITLE_PT) <= 0.6 for s in t_sizes)
          and not any(t_bolds))
    add(not ok, -3, "Word 标题“产品摘要”字号不为小初，或为加粗")

    if title_para is not None:
        sb, sa = _spacing_pt(title_para)
        ok = (abs(sb) <= 0.5 and abs(sa) <= 0.5 and _is_centered(title_para))
    else:
        ok = False
    add(not ok, -1, "Word 标题“产品摘要”段落的段前或段后间距不为 0 磅，或对齐方式不为水平居中")

    body_children = list(doc.element.body)
    tbl_pos = next((i for i, el in enumerate(body_children)
                    if el.tag == qn('w:tbl')), None)
    title_pos = next((i for i, el in enumerate(body_children)
                      if el.tag == qn('w:p') and _norm("".join(
                          t.text or '' for t in el.findall('.//' + qn('w:t')))) == _TITLE_TEXT),
                     None)
    blank_between = False
    if title_pos is not None and tbl_pos is not None and tbl_pos > title_pos:
        for el in body_children[title_pos + 1:tbl_pos]:
            if el.tag == qn('w:p') and not _norm("".join(
                    t.text or '' for t in el.findall('.//' + qn('w:t')))):
                blank_between = True
    add(not blank_between, -1, "Word 标题“产品摘要”与表格之间没有空行段落，标题与表格直接相连")

    # ---- 28 左右列列宽 ----
    def _tcw_cm(cell):
        tcPr = cell._tc.find(qn('w:tcPr'))
        if tcPr is None:
            return None
        tcW = tcPr.find(qn('w:tcW'))
        if tcW is None or tcW.get(qn('w:type')) != 'dxa':
            return None
        return _twip_cm(tcW.get(qn('w:w')))

    ok = bool(wd_rows)
    for row in wd_rows:
        cs = row.cells
        if len(cs) < 2:
            ok = False
            continue
        if not _near(_tcw_cm(cs[0]), _COL_LEFT_CM, 0.05):
            ok = False
        if not _near(_tcw_cm(cs[-1]), _COL_RIGHT_CM, 0.05):
            ok = False
    add(not ok, -3, "Word 表格左列列宽不为 12.00 厘米，或右列列宽不为 4.40 厘米")

    # ---- 29 行高 ----
    ok = bool(wd_rows)
    for row in wd_rows:
        trPr = row._tr.find(qn('w:trPr'))
        th = trPr.find(qn('w:trHeight')) if trPr is not None else None
        if th is None or th.get(qn('w:hRule')) != 'atLeast':
            ok = False
            continue
        if not _near(_twip_cm(th.get(qn('w:val'))), _ROW_H_CM, 0.05):
            ok = False
    add(not ok, -3, "Word 表格 3 行的行高不为“最小值 5.10 厘米”")

    # ---- 30 禁止跨页断行 / 分页符分节符 ----
    doc_xml = _xmlstr(doc.element.body)
    ok = (bool(wd_rows)
          and all((row._tr.find(qn('w:trPr')) is not None
                   and row._tr.find(qn('w:trPr')).find(qn('w:cantSplit')) is not None)
                  for row in wd_rows)
          and 'w:type="page"' not in doc_xml
          and len(doc.sections) == 1)
    add(not ok, -1, "Word 表格 3 行中任一行的“允许跨页断行”处于勾选状态，或文档中存在分页符、"
                    "分节符，使某一个产品的文字信息与图片被拆分到两页显示")

    # ---- 31 表格居中 + 固定列宽 ----
    tblPr = tbl._tbl.tblPr if tbl is not None else None
    jc = tblPr.find(qn('w:jc')) if tblPr is not None else None
    layout = tblPr.find(qn('w:tblLayout')) if tblPr is not None else None
    ok = (jc is not None and jc.get(qn('w:val')) == 'center'
          and layout is not None and layout.get(qn('w:type')) == 'fixed')
    add(not ok, -1, "Word 表格未整体水平居中，或表格布局不为“固定列宽”而为“根据内容自动调整”")

    # ---- 32 单元格内边距 ----
    # 表格未显式写 tblCellMar 时沿用 Word 默认（左右 0.19 厘米、上下 0 厘米），视为合规。
    ok = True
    mar = tblPr.find(qn('w:tblCellMar')) if tblPr is not None else None
    if mar is not None:
        for side, want in (('left', _CELL_MAR_LR_CM), ('right', _CELL_MAR_LR_CM),
                           ('top', _CELL_MAR_TB_CM), ('bottom', _CELL_MAR_TB_CM)):
            el = mar.find(qn('w:' + side))
            if el is None:
                continue
            if not _near(_twip_cm(el.get(qn('w:w'))), want, 0.02):
                ok = False
    add(not ok, -1, "Word 表格单元格左右内边距不为 0.19 厘米，或上下内边距不为 0 厘米")

    # ---- 33 垂直居中 + 右列图片段落居中 ----
    ok = bool(all_cells)
    for c in all_cells:
        tcPr = c._tc.find(qn('w:tcPr'))
        va = tcPr.find(qn('w:vAlign')) if tcPr is not None else None
        if va is None or va.get(qn('w:val')) != 'center':
            ok = False
    for right in wd_right:
        for para in (right.paragraphs if right is not None else []):
            if para._p.findall('.//' + qn('w:drawing')) and not _is_centered(para):
                ok = False
    add(not ok, -1, "Word 表格单元格内容垂直对齐方式不为居中，或右列图片所在段落不为水平居中")

    # ---- 34/35/36 单元格内文字格式 ----
    prod_runs, label_runs, value_runs = [], [], []
    for left in wd_left:
        if left is None:
            continue
        p, l, v = _cell_run_groups(left)
        prod_runs.extend(p)
        label_runs.extend(l)
        value_runs.extend(v)
    add(not (bool(prod_runs) and _runs_font_ok(prod_runs, _PROD_LABEL_PT, True)), -1,
        "Word“产品一：”“产品二：”“产品三：”三个段落不为宋体、小三、加粗")
    add(not (bool(label_runs) and _runs_font_ok(label_runs, _BODY_PT, True)), -1,
        "Word“产品名称：”“核心参数：”“适用场景：”“指导价格：”四个字段提示词不为宋体、五号、加粗")
    add(not _runs_font_ok(value_runs, _BODY_PT, False), -1,
        "Word 三行填入的产品信息文字不为宋体、五号，或为加粗")

    # ---- 37 正文多余内容 ----
    extra = (len(doc.tables) > 1)
    for i, el in enumerate(body_children):
        if el.tag != qn('w:p'):
            continue
        txt = _norm("".join(t.text or '' for t in el.findall('.//' + qn('w:t'))))
        if not txt:
            continue
        if i == title_pos and txt == _TITLE_TEXT:
            continue
        extra = True
    add(extra, -3, "Word 正文中除“产品摘要”标题、其后的空行段落与该 3 行 2 列表格之外，"
                   "还有其他非空段落或第二个表格")

    # ---- 38 纸张尺寸 ----
    ok = (_near(_cm(sec.page_width), _A4_CM[0], 0.05)
          and _near(_cm(sec.page_height), _A4_CM[1], 0.05))
    add(not ok, -1, "Word 纸张尺寸不为 A4（宽 21.0 厘米×高 29.7 厘米）")

    # ================= 减分项：Excel 固有格式是否被破坏 =================
    # ---- 39 表头文字 ----
    hdr_txt = [_norm(ws.cell(row=1, column=i + 1).value) for i in range(5)]
    tail = [_norm(ws.cell(row=1, column=c).value)
            for c in range(6, min(ws.max_column, 30) + 1)]
    ok = (hdr_txt == list(_HEADERS) and not any(tail))
    add(not ok, -3, "Excel 第一行五个表头自左至右不依次为“产品名称”“核心参数”“适用场景”"
                    "“指导价格”“展示图”，或第一行“展示图”之后仍有非空表头单元格")

    # ---- 40 四列列宽 ----
    ok = True
    for name, want in _XL_COL_W.items():
        d = ws.column_dimensions.get(get_column_letter(cols[name]))
        got = d.width if (d is not None and d.width) else None
        if not _near(got, want, _XL_COL_W_TOL):
            ok = False
    add(not ok, -3, "Excel“产品名称”列列宽不为 24.13 字符、“核心参数”列列宽不为 33.52 字符、"
                    "“指导价格”列列宽不为 13.08 字符、“展示图”列列宽不为 12.79 字符"
                    "（各允许±1.5 字符误差）")

    # ---- 41 表头字体字号居中底纹边框 ----
    ok = True
    for i in range(5):
        c = ws.cell(row=1, column=i + 1)
        if (c.font.name or '') != _XL_HDR_FONT:
            ok = False
        if not (c.font.sz and abs(float(c.font.sz) - _XL_HDR_PT) <= 0.3):
            ok = False
        if c.alignment.horizontal != 'center':
            ok = False
        if c.fill is None or c.fill.fill_type != 'solid':
            ok = False
        for side in ('left', 'right', 'top', 'bottom'):
            if getattr(getattr(c.border, side), 'style', None) not in ('thin', 'hair'):
                ok = False
    add(not ok, -1, "Excel 第一行 5 个表头单元格不为 Noto Sans CJK SC 字体、10 磅、水平居中，"
                    "或无底纹填充色、四周无细实线边框")

    # ---- 42 工作表名称与数量 ----
    ok = (wb.sheetnames == [_XL_SHEET_NAME])
    add(not ok, -1, "Excel 工作表名称不为“工作表1”，或工作簿内工作表数量多于 1 个")

    # ---- 43 展示图像素宽度 ----
    narrow = False
    for _n, b in media:
        _bgr, pw, _ph = _load_bgr(b)
        if pw and pw < _IMG_MIN_PX_W:
            narrow = True
    add(narrow, -1, "Excel 中任意一张展示图的像素宽度小于 200 像素，画面模糊")

    # ---- 44 多余参数字段 ----
    extra_param = any(_has_extra_param(t) for t in xl_pars)
    for i in range(len(wd_rows)):
        if _has_extra_param("\n".join(_fl(i, "核心参数："))):
            extra_param = True
    add(extra_param, -1, "Excel 或 Word 的“核心参数”内容中出现洗涤容量、宽度、类型、"
                         "能效等级、深度、高度之外的其他参数字段")

    # ========== 将 results 反解为结构化 dim2_items ==========
    # 逐项按追加顺序对齐 _MAX_DELTAS_ORDER：解析每条 “+N：说明”/“-N：说明” 的 N 与 rule。
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
        # 加分项：delta>0 表示拿到分；减分项：delta<0 表示扣分成立。
        # 两者统一用 hit 表示“该评分项判定成立”。
        hit = (delta != 0)
        dim2_items.append({"rule": _RULE_ONLY.sub("", rule), "max_delta": max_delta,
                           "delta": max_delta if hit else 0, "hit": hit})
        total_score += (max_delta if hit else 0)

    result["dim2_items"] = dim2_items
    result["total_score"] = max(0, total_score)
    result["max_score"] = MAX_SCORE_TOTAL
    return result


if __name__ == "__main__":
    # 本地调试入口：evaluate 只接收“交付物所在目录的路径”，由脚本自己在该目录里定位
    # 并打开被评估的 Word 文档与 Excel 工作簿。命令行第 1 个参数可覆盖目录路径；
    # 未指定时使用脚本自身所在目录。
    if len(sys.argv) >= 2:
        _dir = sys.argv[1]
    else:
        _dir = os.path.dirname(os.path.abspath(__file__))
    _out = evaluate(_dir)
    print(json.dumps(_out, ensure_ascii=False, indent=2))
