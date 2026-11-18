#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# pyright: reportUnknownMemberType=false, reportUnknownVariableType=false, reportUnknownArgumentType=false, reportUnknownParameterType=false, reportMissingParameterType=false, reportPrivateUsage=false, reportGeneralTypeIssues=false
"""对"日常信息处理能力观察表"的自动评估脚本（评测细则见同目录 评测细则.txt）。

评估逻辑：
  维度1（可用与可修改性）：门槛维度。任意一条不满足 -> 直接判 0 分，不再评维度2。
  维度2（完成度评分细则）：加分项 + 扣分项。
      - 加分细则：必须满足细则内的【每一个点】才加分；
      - 扣分细则：只要满足细则内的【任意一点】即扣分；
      - 累计各条命中的分值（可正可负）即为最终得分，满分只统计加分项。

实现说明：
  1) .xlsx/.xlsm 走 OOXML 深解析（zipfile + ElementTree，仅标准库）：单元格文字来自
     sharedStrings/inlineStr，字体/填充/边框/对齐来自 styles.xml 的 cellXfs 索引链，
     图形几何来自 xl/drawings/drawing*.xml。两种常见写法（默认命名空间与 x: 前缀、
     indexed 调色板与 srgbClr）都做了兼容。
  2) .xls 走 BIFF 解析（可选依赖 xlrd，formatting_info=True），可覆盖全部单元格类
     细则；但 xlrd 无法读取 Escher 图形记录，故 .xls 交付物的"图形"类加分项一律按
     未命中处理（detail 中注明不可判定），此限制在脚本内显式标注，不静默放过。

单位换算：
  1 磅(pt) = 12700 EMU；1 像素(px) = 9525 EMU；1 字符列宽 = 7 px（标准 MDW）。
"""

import os
import re
import sys
import json
import zipfile
import shutil
import tempfile
import xml.etree.ElementTree as ET

SCRIPT_ID = "093"

# 维度二所有加分项 max_delta 之和；也是维度一未通过时回填的 max_score
# （脚本没评上时 max_score 为 0）
MAX_SCORE_TOTAL = 33

# 维度二全部评分项按"结果追加顺序"的 (max_delta, rule) 表。顺序须与
# check_dimension2 中 hits.append(...) 的顺序严格一致：前 9 条为加分项
# （累加 = MAX_SCORE_TOTAL），其后 26 条为扣分项。
_RULE_ORDER = [
    (5, '"图形组织"标签右侧、含"30、"序号的C32单元格内含有图形，图样为两个五边形相互交叉的组合'),
    (3, '该图形为可选中、可拖动、可调整大小与线条颜色的可编辑矢量自选图形，非位图图片、非艺术字、非文本符号拼凑'),
    (3, '两个图形均为闭合的正五边形几何形状，不是菱形、六边形、圆形、星形等其他形状'),
    (3, '两个五边形沿水平方向并排放置并部分重叠，重叠区域在中部形成交叉的菱形空腔'),
    (5, '两个五边形均为无填充的透明图形，重叠处两个五边形的轮廓线彼此可见、互不遮挡'),
    (3, '两个五边形各自带有倾斜旋转角度，左侧五边形约顺时针旋转20度、右侧五边形约逆时针旋转10度，两者旋转方向相反、姿态相差约30度，而非两个完全正立、姿态相同的五边形'),
    (1, '两个五边形轮廓为RGB 000000黑色或近黑色的单线实线，线宽约1—1.5磅，无虚线、无阴影、无渐变、无三维效果'),
    (5, '图形水平起点位于"30、"序号文字右侧，"30、"字样完整可见、未被图形遮挡'),
    (5, '图形整体完整落在C32单元格内部，左右不越出C列左右边界、上下不越出第32行上下边界，未压住任何单元格框线'),
    # ---- 以下为扣分项，max_delta 为负 ----
    (-3, '工作簿不为且仅为排序在前的"Sheet0"与随后的"Sheet1"两个工作表，或Sheet0数据区域不为33行7列的A1:G33'),
    (-3, 'Sheet0不含以下33处合并单元格：A1:C2、D1:E1、F1:F2、G1:G2、A3:A12、A13:A15、A16:A20、A21:A23、A24:A32、B24:B25、B27:B29、B3:C3至B23:C23共21处、B33:C33'),
    (-3, 'A1、D1、F1三格文字不依次为"观察维度""观察记录""完成情况"'),
    (-1, 'A3、A13、A16、A21、A24五格文字不依次为含单元格内换行的两行文本"情境识别（10项）""即时记忆（3项）""专注处理（5项）""延迟回想（3项）""表达与执行（9项）"，或A33文字不为"综合结果"'),
    (-3, 'B3至B12共10格文字不依次为"1、请说出当前所在的年度。""2、请判断现在处于哪个季度。""3、请说出当前月份。""4、请说出今天的日期。""5、请判断今天是星期几。""6、请说出目前所在的城市。""7、请说出目前所在的城区。""8、请描述所在场所的主要用途。""9、请判断自己所在的楼层。""10、请说出当前房间的功能。"'),
    (-3, 'B13至B23共11格文字不依次为"11、复述词语：灯塔""12、复述词语：河流""13、复述词语：松果""14、从80开始连续减6。""15、说出第一次计算结果。""16、在上一步结果上再减6。""17、继续进行一次相同计算。""18、再完成最后一次减6。""19、回想词语：灯塔""20、回想词语：河流""21、回想词语：松果"'),
    (-3, 'C24至C31共8格文字不依次为"22、辨认：钥匙扣""23、辨认：便签本""24、复述：清晨的风吹过安静的长桥。""25、用左手拿起桌面的卡片。""26、将卡片翻到背面。""27、把卡片放回指定位置。""28、按文字要求完成动作"+单元格内换行+带中文引号的"轻触桌面两次""29、写一句包含时间和地点的话。"'),
    (-1, 'B24、B26、B27、B30、B31、B32六格文字不依次为"物品识别""语句复现""连续指令""文字执行""自主表达""图形组织"'),
    (-3, 'C32单元格文字不为含顿号的"30、"'),
    (-1, 'F3至F32共30格文字不均为否与是之间空两格的"否  是"'),
    (-1, '第1、2行除A1、D1、F1外的单元格、D列与E列第3至32行、G列第1至32行、B33至G33中出现文字内容'),
    (-3, 'A1至G2表头区文字不为10磅加粗、颜色不为白色FFFFFF、填充不为深青色1F4E5F、边框不为D6E3E7细实线，或A1、D1、F1三格字体不为DejaVu Sans'),
    (-3, 'A33至G33文字不为10磅加粗、颜色不为白色FFFFFF、填充不为2D6672、边框不为D6E3E7细实线'),
    (-3, 'A3至A32五个维度标签单元格文字不为DejaVu Sans 10磅加粗、颜色不为164B55、填充不为D7ECE8、边框不为8EA8B1细实线'),
    (-3, 'B24至B32小类标签单元格文字不为DejaVu Sans 10磅加粗、颜色不为234F5C、填充不为E7F0F4、边框不为8EA8B1细实线'),
    (-3, '题目与记录区B3:B23、C3:C32、D3:E32、G3:G32文字不为Arial 10磅不加粗、颜色不为26383F，或边框不为8EA8B1细实线'),
    (-3, '题目与记录区B3:B23、C3:C32、D3:E32隔行底色不为奇数行填充FFFFFF、偶数行填充F4F8F8，或G3至G32底色不为统一FFFFFF'),
    (-3, 'F3至F32单元格文字不为DejaVu Sans 10磅加粗、颜色不为694A12、填充不为FFF1CC、边框不为C6A762细实线'),
    (-1, '水平对齐不为——A1至G2表头区、A列维度标签、B24至B32小类标签、F3至F32、A33至G33居中，B3:B23与C24:C32题目文字常规靠左'),
    (-1, '单元格不为垂直居中，或未开启自动换行'),
    (-3, '行高不为——第1至29行及第31行约19.2磅、第30行约38.4磅、第32行42磅、第33行44磅'),
    (-3, '列宽不为——A列约7.91、B列约10.70、C列约32.71、D列约7.61、E列约9.35、F列约16.60、G列约7.98字符'),
    (-1, 'C32以外的题号行、"完成情况"列或Sheet1中存在图形对象'),
    (-1, '图形内部含有文字内容，或图形上叠加"30"等字符'),
    (-1, '图形高度超过第32行行高42磅，或宽度超过C列列宽'),
    (-3, '若已插入图形，其锚定方式不为随单元格位置与大小变化的两点锚定，行高或列宽变动时图形不随单元格同步移动与缩放'),
]

assert sum(d for d, _ in _RULE_ORDER if d > 0) == MAX_SCORE_TOTAL, "加分项合计与 MAX_SCORE_TOTAL 不一致"

# ---------------------------------------------------------------- 常量 ----

EMU_PER_PT = 12700.0
EMU_PER_PX = 9525.0
PX_PER_CHAR = 7.0           # 标准 MDW（Calibri 11 / 96dpi）
CHAR_PADDING_PX = 5.0

NS = {
    'x': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    'rel': 'http://schemas.openxmlformats.org/package/2006/relationships',
    'ct': 'http://schemas.openxmlformats.org/package/2006/content-types',
    'xdr': 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing',
    'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
}

# BIFF8 默认 indexed 调色板（索引 0..63），用于 styles.xml 未自带 indexedColors 时
_DEFAULT_INDEXED = [
    '000000', 'FFFFFF', 'FF0000', '00FF00', '0000FF', 'FFFF00', 'FF00FF', '00FFFF',
    '000000', 'FFFFFF', 'FF0000', '00FF00', '0000FF', 'FFFF00', 'FF00FF', '00FFFF',
    '800000', '008000', '000080', '808000', '800080', '008080', 'C0C0C0', '808080',
    '9999FF', '993366', 'FFFFCC', 'CCFFFF', '660066', 'FF8080', '0066CC', 'CCCCFF',
    '000080', 'FF00FF', 'FFFF00', '00FFFF', '800080', '800000', '008080', '0000FF',
    '00CCFF', 'CCFFFF', 'CCFFCC', 'FFFF99', '99CCFF', 'FF99CC', 'CC99FF', 'FFCC99',
    '3366FF', '33CCCC', '99CC00', 'FFCC00', 'FF9900', 'FF6600', '666699', '969696',
    '003366', '339966', '003300', '333300', '993300', '993366', '333399', '333333',
]

# 目标区域颜色 / 字体
C_WHITE, C_HDR_FILL, C_FOOT_FILL = 'FFFFFF', '1F4E5F', '2D6672'
C_HDR_BORDER, C_LABEL_BORDER = 'D6E3E7', '8EA8B1'
C_DIMLABEL_FONT, C_DIMLABEL_FILL = '164B55', 'D7ECE8'
C_SUBLABEL_FONT, C_SUBLABEL_FILL = '234F5C', 'E7F0F4'
C_BODY_FONT, C_STRIPE_EVEN = '26383F', 'F4F8F8'
C_F_FONT, C_F_FILL, C_F_BORDER = '694A12', 'FFF1CC', 'C6A762'
FONT_LABEL, FONT_BODY = 'DejaVu Sans', 'Arial'

# 目标行高（磅）与列宽（字符）
EXPECT_ROW_HT = dict([(r, 19.2) for r in range(1, 30)] + [(30, 38.4), (31, 19.2), (32, 42.0), (33, 44.0)])
EXPECT_COL_W = {1: 7.91, 2: 10.70, 3: 32.71, 4: 7.61, 5: 9.35, 6: 16.60, 7: 7.98}
ROW_HT_TOL, COL_W_TOL = 0.6, 0.25

# 目标合并单元格（33 处）
EXPECT_MERGES = set(
    ['A1:C2', 'D1:E1', 'F1:F2', 'G1:G2',
     'A3:A12', 'A13:A15', 'A16:A20', 'A21:A23', 'A24:A32',
     'B24:B25', 'B27:B29', 'B33:C33']
    + ['B%d:C%d' % (r, r) for r in range(3, 24)]
)

# 目标文字
EXPECT_HEAD = {'A1': '观察维度', 'D1': '观察记录', 'F1': '完成情况'}
EXPECT_DIM_LABELS = [
    ('A3', '情境识别\n（10项）'), ('A13', '即时记忆\n（3项）'), ('A16', '专注处理\n（5项）'),
    ('A21', '延迟回想\n（3项）'), ('A24', '表达与执行\n（9项）'),
]
EXPECT_B3_B12 = [
    '1、请说出当前所在的年度。', '2、请判断现在处于哪个季度。', '3、请说出当前月份。',
    '4、请说出今天的日期。', '5、请判断今天是星期几。', '6、请说出目前所在的城市。',
    '7、请说出目前所在的城区。', '8、请描述所在场所的主要用途。', '9、请判断自己所在的楼层。',
    '10、请说出当前房间的功能。',
]
EXPECT_B13_B23 = [
    '11、复述词语：灯塔', '12、复述词语：河流', '13、复述词语：松果', '14、从80开始连续减6。',
    '15、说出第一次计算结果。', '16、在上一步结果上再减6。', '17、继续进行一次相同计算。',
    '18、再完成最后一次减6。', '19、回想词语：灯塔', '20、回想词语：河流', '21、回想词语：松果',
]
EXPECT_C24_C31 = [
    '22、辨认：钥匙扣', '23、辨认：便签本', '24、复述：清晨的风吹过安静的长桥。',
    '25、用左手拿起桌面的卡片。', '26、将卡片翻到背面。', '27、把卡片放回指定位置。',
    '28、按文字要求完成动作\n“轻触桌面两次”', '29、写一句包含时间和地点的话。',
]
EXPECT_SUB_LABELS = [
    ('B24', '物品识别'), ('B26', '语句复现'), ('B27', '连续指令'),
    ('B30', '文字执行'), ('B31', '自主表达'), ('B32', '图形组织'),
]
EXPECT_C32_TEXT = '30、'
EXPECT_F_TEXT = '否  是'

# ------------------------------------------------------------ 通用工具 ----


def _col_to_idx(letters):
    """'A' -> 1, 'AA' -> 27。"""
    n = 0
    for ch in letters:
        n = n * 26 + (ord(ch.upper()) - 64)
    return n


def _idx_to_col(idx):
    s = ''
    while idx > 0:
        idx, rem = divmod(idx - 1, 26)
        s = chr(65 + rem) + s
    return s


def _split_ref(ref):
    """'C32' -> (32, 3)；解析失败返回 (0, 0)。"""
    m = re.match(r'^\$?([A-Za-z]+)\$?(\d+)$', ref or '')
    if not m:
        return (0, 0)
    return (int(m.group(2)), _col_to_idx(m.group(1)))


def _norm_text(val):
    """统一换行符并去掉首尾空白；None -> ''。"""
    if val is None:
        return ''
    s = val if isinstance(val, str) else str(val)
    return s.replace('\r\n', '\n').replace('\r', '\n').strip()


def _col_width_to_emu(chars):
    """字符列宽 -> EMU。"""
    if chars is None:
        return 0.0
    return (float(chars) * PX_PER_CHAR + CHAR_PADDING_PX) * EMU_PER_PX


def _est_text_width_emu(text, size_pt):
    """粗估一段文字的显示宽度（EMU）：ASCII 记 0.5em，全角/中文记 1em。"""
    w = 0.0
    for ch in text or '':
        w += 0.5 if ord(ch) < 0x2E80 else 1.0
    return w * float(size_pt or 10) * EMU_PER_PT + 2 * EMU_PER_PX


def _hex6(rgb):
    """'FF1F4E5F' / '1f4e5f' -> '1F4E5F'；无效返回 None。"""
    if not rgb:
        return None
    s = str(rgb).strip().upper().lstrip('#')
    if len(s) == 8:
        s = s[2:]
    return s if re.match(r'^[0-9A-F]{6}$', s) else None


def _apply_tint(hex6, tint):
    """按 OOXML 规则对颜色施加 tint（-1..1）。"""
    if not hex6 or not tint:
        return hex6
    try:
        t = float(tint)
    except (TypeError, ValueError):
        return hex6
    out = ''
    for i in range(0, 6, 2):
        v = int(hex6[i:i + 2], 16)
        if t < 0:
            v = int(round(v * (1.0 + t)))
        else:
            v = int(round(v * (1.0 - t) + 255.0 * t))
        out += '%02X' % max(0, min(255, v))
    return out


def _color_dist(a, b):
    """两个 6 位十六进制颜色的最大通道差；任一为空返回 999。"""
    if not a or not b:
        return 999
    return max(abs(int(a[i:i + 2], 16) - int(b[i:i + 2], 16)) for i in (0, 2, 4))


def _same_color(a, b, tol=6):
    return _color_dist(a, b) <= tol


# ------------------------------------------------------------ 数据容器 ----


class WorkbookData(object):
    """归一化后的工作簿数据；.xlsx 与 .xls 两条解析路径填充同一套字段。"""

    def __init__(self):
        self.path = ''
        self.ext = ''
        self.ok_open = False
        self.open_error = ''
        self.sheet_names = []
        self.merged = set()          # {'A1:C2', ...}
        self.cells = {}              # 'C32' -> 单元格字典（仅首个工作表）
        self.max_row = 0
        self.max_col = 0
        self.row_ht = {}             # 行号 -> 磅
        self.col_w = {}              # 列号 -> 字符
        self.default_col_w = 8.43
        self.default_row_ht = 15.0
        self.shapes = []             # 首个工作表内的图形（已展开组合）
        self.other_shapes = 0        # 除首个工作表外其它工作表的图形数
        self.shape_readable = True   # .xls 下为 False：图形记录不可解析
        self.has_drawing_hint = False  # .xls 下按 Escher 签名推断是否存在图形
        self.pictures = 0            # 首个工作表内的位图数量

    def cell(self, ref):
        return self.cells.get(ref)

    def text(self, ref):
        c = self.cells.get(ref)
        return c['text'] if c else ''


def _blank_cell():
    return {
        'text': '', 'fname': None, 'sz': None, 'bold': False, 'fcolor': None,
        'eff': None,          # 富文本片段实际生效的字号/加粗/字色
        'fill': None, 'pattern': None,
        'borders': {'left': (None, None), 'right': (None, None),
                    'top': (None, None), 'bottom': (None, None)},
        'halign': None, 'valign': None, 'wrap': False,
    }


# ------------------------------------------------------- OOXML 解析路径 ----


def _xml(tmp, rel_path):
    """读取解包目录下的 XML；不存在返回 None。"""
    full = os.path.join(tmp, rel_path.replace('/', os.sep))
    if not os.path.isfile(full):
        return None
    try:
        return ET.parse(full).getroot()
    except ET.ParseError:
        return None


def _rels_map(tmp, part):
    """part='xl/worksheets/sheet1.xml' -> {rId: 解析后的包内路径}。"""
    d, base = os.path.split(part)
    root = _xml(tmp, '%s/_rels/%s.rels' % (d, base))
    out = {}
    if root is None:
        return out
    for rel in root.findall('rel:Relationship', NS):
        rid, tgt = rel.get('Id'), rel.get('Target') or ''
        if not rid:
            continue
        if tgt.startswith('/'):
            out[rid] = tgt.lstrip('/')
            continue
        cur = d.split('/')
        for seg in tgt.split('/'):
            if seg == '..':
                if cur:
                    cur.pop()
            elif seg not in ('', '.'):
                cur.append(seg)
        out[rid] = '/'.join(cur)
    return out


def _run_format(st, pr):
    """<rPr> -> 字体/字号/加粗/字色；未声明的项为 None，表示沿用单元格字体。"""
    item = {'fname': None, 'sz': None, 'bold': None, 'fcolor': None}
    if pr is None:
        return item
    name = pr.find('x:rFont', NS)
    if name is not None:
        item['fname'] = name.get('val')
    sz = pr.find('x:sz', NS)
    if sz is not None:
        try:
            item['sz'] = float(sz.get('val'))
        except (TypeError, ValueError):
            pass
    b = pr.find('x:b', NS)
    if b is not None:
        item['bold'] = b.get('val') in (None, '1', 'true', 'True')
    color = pr.find('x:color', NS)
    if color is not None:
        item['fcolor'] = st.resolve(color)
    return item


def _rich_text(node, st):
    """解析 <si> / <is> 容器 -> (文本, 富文本 run 格式列表)。

    run 上的 rPr 会覆盖单元格字体，屏幕上看到的字号/加粗/字色由它决定，
    因此一并取出，供格式判定按实际显示效果核对。
    """
    parts = []
    runs = []
    for child in node:
        tag = child.tag.split('}')[-1]
        if tag == 't':
            parts.append(child.text or '')
        elif tag == 'r':          # 富文本 run
            text = ''
            for sub in child:
                if sub.tag.split('}')[-1] == 't':
                    text += sub.text or ''
            parts.append(text)
            if text:
                runs.append(_run_format(st, child.find('x:rPr', NS)))
        # rPh / phoneticPr 为拼音注音，不计入正文
    return ''.join(parts), runs


def _shared_strings(tmp, st):
    """共享字符串表 -> [{'text': 文本, 'runs': [run 格式, ...]}]。"""
    root = _xml(tmp, 'xl/sharedStrings.xml')
    out = []
    if root is None:
        return out
    for si in root.findall('x:si', NS):
        text, runs = _rich_text(si, st)
        out.append({'text': text, 'runs': runs})
    return out


def _theme_colors(tmp):
    """xl/theme/theme1.xml 的 clrScheme -> 按 xlsx theme 索引排列的颜色表。"""
    root = _xml(tmp, 'xl/theme/theme1.xml')
    if root is None:
        return []
    scheme = {}
    for node in root.iter('{%s}clrScheme' % NS['a']):
        for slot in node:
            key = slot.tag.split('}')[-1]
            val = None
            for sub in slot:
                sub_tag = sub.tag.split('}')[-1]
                if sub_tag == 'srgbClr':
                    val = _hex6(sub.get('val'))
                elif sub_tag == 'sysClr':
                    val = _hex6(sub.get('lastClr')) or ('FFFFFF' if 'window' in (sub.get('val') or '') else '000000')
            scheme[key] = val
        break
    order = ['lt1', 'dk1', 'lt2', 'dk2', 'accent1', 'accent2', 'accent3',
             'accent4', 'accent5', 'accent6', 'hlink', 'folHlink']
    return [scheme.get(k) for k in order]


class _StyleTables(object):
    """styles.xml 的 fonts/fills/borders/cellXfs 索引链。"""

    def __init__(self):
        self.palette = list(_DEFAULT_INDEXED)
        self.theme = []
        self.fonts = []
        self.fills = []
        self.borders = []
        self.xfs = []
        self.style_xfs = []

    def resolve(self, node):
        """<color .../> -> 6 位十六进制；auto / 缺省返回 None。"""
        if node is None:
            return None
        rgb = _hex6(node.get('rgb'))
        if rgb:
            return _apply_tint(rgb, node.get('tint'))
        idx = node.get('indexed')
        if idx is not None:
            try:
                i = int(idx)
            except ValueError:
                return None
            if 0 <= i < len(self.palette):
                return _hex6(self.palette[i])
            return {64: '000000', 65: 'FFFFFF'}.get(i)
        th = node.get('theme')
        if th is not None:
            try:
                i = int(th)
            except ValueError:
                return None
            if 0 <= i < len(self.theme):
                return _apply_tint(self.theme[i], node.get('tint'))
        return None


def _parse_styles(tmp):
    st = _StyleTables()
    st.theme = _theme_colors(tmp)
    root = _xml(tmp, 'xl/styles.xml')
    if root is None:
        return st
    pal = [c.get('rgb') for c in root.iterfind('.//x:indexedColors/x:rgbColor', NS)]
    if pal:
        st.palette = pal
    for f in root.iterfind('x:fonts/x:font', NS):
        name = f.find('x:name', NS)
        sz = f.find('x:sz', NS)
        b = f.find('x:b', NS)
        try:
            size = float(sz.get('val')) if sz is not None else None
        except (TypeError, ValueError):
            size = None
        st.fonts.append({
            'name': name.get('val') if name is not None else None,
            'sz': size,
            'bold': b is not None and (b.get('val') in (None, '1', 'true', 'True')),
            'color': st.resolve(f.find('x:color', NS)),
        })
    for fl in root.iterfind('x:fills/x:fill', NS):
        pf = fl.find('x:patternFill', NS)
        if pf is None:
            st.fills.append({'pattern': None, 'fg': None})
            continue
        st.fills.append({'pattern': pf.get('patternType'),
                         'fg': st.resolve(pf.find('x:fgColor', NS))})
    for bd in root.iterfind('x:borders/x:border', NS):
        item = {}
        for side in ('left', 'right', 'top', 'bottom'):
            node = bd.find('x:%s' % side, NS)
            if node is None:
                item[side] = (None, None)
            else:
                item[side] = (node.get('style'), st.resolve(node.find('x:color', NS)))
        st.borders.append(item)
    for tag, sink in (('cellStyleXfs', st.style_xfs), ('cellXfs', st.xfs)):
        for xf in root.iterfind('x:%s/x:xf' % tag, NS):
            al = xf.find('x:alignment', NS)
            sink.append({
                'fontId': xf.get('fontId'), 'fillId': xf.get('fillId'),
                'borderId': xf.get('borderId'), 'xfId': xf.get('xfId'),
                'align': (None if al is None else {
                    'h': al.get('horizontal'), 'v': al.get('vertical'),
                    'wrap': al.get('wrapText') in ('1', 'true', 'True')}),
            })
    return st


def _cell_format(st, s_attr):
    """按 cellXfs 索引组装单元格格式；缺省属性回落到 cellStyleXfs。"""
    cell = _blank_cell()
    try:
        si = int(s_attr) if s_attr is not None else 0
    except ValueError:
        si = 0
    xf = st.xfs[si] if 0 <= si < len(st.xfs) else None
    if xf is None:
        return cell
    base = None
    if xf.get('xfId') is not None:
        try:
            bi = int(xf['xfId'])
        except ValueError:
            bi = -1
        if 0 <= bi < len(st.style_xfs):
            base = st.style_xfs[bi]

    def pick(key):
        v = xf.get(key)
        if v is None and base is not None:
            v = base.get(key)
        try:
            return int(v)
        except (TypeError, ValueError):
            return None

    fi, li, bi2 = pick('fontId'), pick('fillId'), pick('borderId')
    if fi is not None and 0 <= fi < len(st.fonts):
        f = st.fonts[fi]
        cell['fname'], cell['sz'] = f['name'], f['sz']
        cell['bold'], cell['fcolor'] = f['bold'], f['color']
    if li is not None and 0 <= li < len(st.fills):
        fl = st.fills[li]
        cell['pattern'] = fl['pattern']
        cell['fill'] = fl['fg'] if fl['pattern'] == 'solid' else (
            'FFFFFF' if fl['pattern'] in (None, 'none') else fl['fg'])
    if bi2 is not None and 0 <= bi2 < len(st.borders):
        cell['borders'] = dict(st.borders[bi2])
    al = xf.get('align') or (base.get('align') if base else None)
    if al:
        cell['halign'], cell['valign'], cell['wrap'] = al['h'], al['v'], al['wrap']
    return cell


def _cell_value(node, shared, st):
    """返回 (单元格文本, 富文本 run 格式列表)。"""
    t = node.get('t')
    if t == 's':
        v = node.find('x:v', NS)
        try:
            item = shared[int(v.text)] if v is not None and v.text else None
        except (ValueError, IndexError):
            item = None
        return ('', []) if item is None else (item['text'], item['runs'])
    if t == 'inlineStr':
        node_is = node.find('x:is', NS)
        if node_is is None:
            return '', []
        return _rich_text(node_is, st)
    v = node.find('x:v', NS)
    if v is None or v.text is None:
        return '', []
    if t in ('str', 'e', 'b'):
        return v.text, []
    try:
        num = float(v.text)
        return (str(int(num)) if num == int(num) else repr(num)), []
    except ValueError:
        return v.text, []


def _apply_runs(cell, runs):
    """把富文本片段的生效字号/加粗/字色写入 cell['eff']。

    字体名不参与：中英文混排时同一段文字本身就会按字符集拆成多个 run
    （数字 Arial、汉字另一字体），字体名以单元格声明为基准更可核对；
    而字号/加粗/字色一旦被 run 改写，屏幕显示会立刻变形，必须按 run 判定。
    """
    if not runs:
        return
    eff = {}
    for key in ('sz', 'bold', 'fcolor'):
        eff[key] = [cell[key] if run[key] is None else run[key] for run in runs]
    cell['eff'] = eff


def _eff(cell, key):
    """该属性实际生效的取值列表：富文本按片段展开，普通文本即单元格自身取值。"""
    vals = (cell.get('eff') or {}).get(key)
    return vals if vals else [cell[key]]


def _parse_sheet(tmp, part, st, shared, data):
    """解析首个工作表：单元格、合并、行高、列宽；返回 drawing 部件路径或 None。"""
    root = _xml(tmp, part)
    if root is None:
        return None
    fmt = root.find('x:sheetFormatPr', NS)
    if fmt is not None:
        for attr, key in (('defaultColWidth', 'default_col_w'), ('defaultRowHeight', 'default_row_ht')):
            try:
                if fmt.get(attr) is not None:
                    setattr(data, key, float(fmt.get(attr)))
            except ValueError:
                pass
    for col in root.iterfind('x:cols/x:col', NS):
        try:
            lo, hi = int(col.get('min')), int(col.get('max'))
            w = float(col.get('width'))
        except (TypeError, ValueError):
            continue
        for c in range(lo, min(hi, 64) + 1):
            data.col_w[c] = w
    for row in root.iterfind('x:sheetData/x:row', NS):
        try:
            rno = int(row.get('r'))
        except (TypeError, ValueError):
            continue
        if row.get('ht') is not None:
            try:
                data.row_ht[rno] = float(row.get('ht'))
            except ValueError:
                pass
        for c in row.findall('x:c', NS):
            ref = c.get('r') or ''
            rr, cc = _split_ref(ref)
            if not rr:
                continue
            cell = _cell_format(st, c.get('s'))
            text, runs = _cell_value(c, shared, st)
            cell['text'] = _norm_text(text)
            _apply_runs(cell, runs)
            data.cells[ref] = cell
            # 数据区域按"工作表内实际存在的单元格"统计（含仅设格式的空单元格）
            data.max_row = max(data.max_row, rr)
            data.max_col = max(data.max_col, cc)
    for mc in root.iterfind('x:mergeCells/x:mergeCell', NS):
        ref = (mc.get('ref') or '').replace('$', '').upper()
        if ref:
            data.merged.add(ref)
    dr = root.find('x:drawing', NS)
    if dr is None:
        return None
    rid = dr.get('{%s}id' % NS['r'])
    return _rels_map(tmp, part).get(rid)


def _fill_kind(node):
    """从 spPr（或 a:ln）子节点判断填充类型与颜色。"""
    if node is None:
        return (None, None)
    for child in node:
        tag = child.tag.split('}')[-1]
        if tag == 'noFill':
            return ('none', None)
        if tag == 'solidFill':
            rgb = None
            for sub in child.iter():
                stag = sub.tag.split('}')[-1]
                if stag == 'srgbClr':
                    rgb = _hex6(sub.get('val'))
                    break
                if stag == 'sysClr':
                    rgb = _hex6(sub.get('lastClr'))
                    break
                if stag == 'schemeClr':
                    rgb = None
                    break
            return ('solid', rgb)
        if tag in ('gradFill', 'blipFill', 'pattFill', 'grpFill'):
            return (tag[:-4] if tag.endswith('Fill') else tag, None)
    return (None, None)


def _shape_geom(sp):
    """返回 (prst, 顶点数, 是否闭合)；prst 为 None 表示非预设几何。"""
    spPr = sp.find('xdr:spPr', NS)
    if spPr is None:
        return (None, 0, False)
    prst = spPr.find('a:prstGeom', NS)
    if prst is not None:
        return (prst.get('prst'), 0, True)
    cust = spPr.find('a:custGeom', NS)
    if cust is None:
        return (None, 0, False)
    npts, closed = 0, False
    for path in cust.iterfind('a:pathLst/a:path', NS):
        for seg in path:
            tag = seg.tag.split('}')[-1]
            if tag in ('moveTo', 'lnTo', 'cubicBezTo', 'quadBezTo', 'arcTo'):
                npts += 1
            elif tag == 'close':
                closed = True
        break
    return (None, npts, closed)


def _int_attr(node, name, default=0):
    if node is None:
        return default
    try:
        return int(float(node.get(name)))
    except (TypeError, ValueError):
        return default


def _anchor_pos(anchor):
    """返回 (from_col, from_colOff, from_row, from_rowOff, to_*，锚定类型)。"""
    kind = anchor.tag.split('}')[-1]
    kind = {'twoCellAnchor': 'two', 'oneCellAnchor': 'one', 'absoluteAnchor': 'abs'}.get(kind, kind)
    fr = anchor.find('xdr:from', NS)
    to = anchor.find('xdr:to', NS)

    def grab(node):
        if node is None:
            return (None, 0, None, 0)
        def sub(tag):
            el = node.find('xdr:%s' % tag, NS)
            try:
                return int(float(el.text))
            except (TypeError, ValueError, AttributeError):
                return None
        return (sub('col'), sub('colOff') or 0, sub('row'), sub('rowOff') or 0)

    f = grab(fr)
    t = grab(to)
    ext = anchor.find('xdr:ext', NS)
    return {
        'anchor': kind,
        'col': f[0], 'colOff': f[1], 'row': f[2], 'rowOff': f[3],
        'to_col': t[0], 'to_colOff': t[1], 'to_row': t[2], 'to_rowOff': t[3],
        'ext_cx': _int_attr(ext, 'cx', 0), 'ext_cy': _int_attr(ext, 'cy', 0),
    }


def _shape_name(obj, kind):
    for tag in ('xdr:nvSpPr/xdr:cNvPr', 'xdr:nvPicPr/xdr:cNvPr',
                'xdr:nvGrpSpPr/xdr:cNvPr', 'xdr:nvGraphicFramePr/xdr:cNvPr',
                'xdr:nvCxnSpPr/xdr:cNvPr'):
        el = obj.find(tag, NS)
        if el is not None:
            return el.get('name') or ''
    return kind


def _mk_shape(obj, kind, anchor, box):
    """把一个绘图对象归一化成评估用记录。"""
    spPr = obj.find('xdr:spPr', NS)
    rec = dict(anchor)
    rec.update({'kind': kind, 'name': _shape_name(obj, kind),
                'x': box['x'], 'y': box['y'], 'cx': box['cx'], 'cy': box['cy'],
                'rot': 0.0, 'flip': False, 'prst': None, 'npts': 0, 'closed': False,
                'fill': None, 'fill_rgb': None, 'ln_w': None, 'ln_rgb': None,
                'ln_fill': None, 'ln_dash': None, 'ln_cmpd': None,
                'has_effect': False, 'has_3d': False, 'text': '', 'txbox': False,
                'in_group': box.get('in_group', False)})
    xfrm = spPr.find('a:xfrm', NS) if spPr is not None else None
    if xfrm is not None:
        try:
            rec['rot'] = ((float(xfrm.get('rot') or 0) / 60000.0) + 180.0) % 360.0 - 180.0
        except ValueError:
            rec['rot'] = 0.0
        rec['flip'] = (xfrm.get('flipH') in ('1', 'true')) or (xfrm.get('flipV') in ('1', 'true'))
        ext = xfrm.find('a:ext', NS)
        cx, cy = _int_attr(ext, 'cx', 0), _int_attr(ext, 'cy', 0)
        if cx > 0 and cy > 0 and not box.get('in_group'):
            rec['cx'], rec['cy'] = cx, cy
    rec['prst'], rec['npts'], rec['closed'] = _shape_geom(obj) if kind == 'sp' else (None, 0, False)
    rec['fill'], rec['fill_rgb'] = _fill_kind(spPr)
    ln = spPr.find('a:ln', NS) if spPr is not None else None
    if ln is not None:
        rec['ln_w'] = _int_attr(ln, 'w', None)
        rec['ln_cmpd'] = ln.get('cmpd')
        rec['ln_fill'], rec['ln_rgb'] = _fill_kind(ln)
        dash = ln.find('a:prstDash', NS)
        rec['ln_dash'] = dash.get('val') if dash is not None else None
    if spPr is not None:
        eff = spPr.find('a:effectLst', NS)
        rec['has_effect'] = eff is not None and len(list(eff)) > 0
        rec['has_3d'] = (spPr.find('a:scene3d', NS) is not None) or (spPr.find('a:sp3d', NS) is not None)
    cnv = obj.find('xdr:nvSpPr/xdr:cNvSpPr', NS)
    rec['txbox'] = cnv is not None and cnv.get('txBox') in ('1', 'true')
    tx = obj.find('xdr:txBody', NS)
    if tx is not None:
        rec['text'] = _norm_text(''.join(t.text or '' for t in tx.iter('{%s}t' % NS['a'])))
    return rec


def _expand_group(grp, anchor, box, depth=0):
    """展开组合图形：把子图形换算到与单元格同一坐标系下。"""
    out = []
    if depth > 3:
        return out
    gx = grp.find('xdr:grpSpPr/a:xfrm', NS)
    sx = sy = 1.0
    chx = chy = 0
    if gx is not None:
        ext, chExt = gx.find('a:ext', NS), gx.find('a:chExt', NS)
        chOff = gx.find('a:chOff', NS)
        chx, chy = _int_attr(chOff, 'x', 0), _int_attr(chOff, 'y', 0)
        ecx, ecy = _int_attr(ext, 'cx', 0), _int_attr(ext, 'cy', 0)
        ccx, ccy = _int_attr(chExt, 'cx', 0), _int_attr(chExt, 'cy', 0)
        if ecx > 0 and ccx > 0:
            sx = float(ecx) / ccx
        if ecy > 0 and ccy > 0:
            sy = float(ecy) / ccy
    for child in grp:
        tag = child.tag.split('}')[-1]
        if tag not in ('sp', 'pic', 'grpSp', 'graphicFrame', 'cxnSp'):
            continue
        cxf = child.find('xdr:spPr/a:xfrm', NS) or child.find('xdr:grpSpPr/a:xfrm', NS)
        off, ext = (None, None) if cxf is None else (cxf.find('a:off', NS), cxf.find('a:ext', NS))
        sub_box = {
            'x': box['x'] + (_int_attr(off, 'x', 0) - chx) * sx,
            'y': box['y'] + (_int_attr(off, 'y', 0) - chy) * sy,
            'cx': _int_attr(ext, 'cx', 0) * sx,
            'cy': _int_attr(ext, 'cy', 0) * sy,
            'in_group': True,
        }
        if tag == 'grpSp':
            out.extend(_expand_group(child, anchor, sub_box, depth + 1))
        else:
            kind = {'sp': 'sp', 'pic': 'pic', 'graphicFrame': 'frame', 'cxnSp': 'cxn'}[tag]
            out.append(_mk_shape(child, kind, anchor, sub_box))
    return out


def _parse_drawing(tmp, part, xy):
    """解析 drawing 部件，返回归一化图形列表。"""
    root = _xml(tmp, part)
    shapes = []
    if root is None:
        return shapes
    for anchor in root:
        tag = anchor.tag.split('}')[-1]
        if tag not in ('twoCellAnchor', 'oneCellAnchor', 'absoluteAnchor'):
            continue
        a = _anchor_pos(anchor)
        x0, y0 = xy(a['col'], a['colOff'], a['row'], a['rowOff'])
        if a['anchor'] == 'two' and a['to_col'] is not None:
            x1, y1 = xy(a['to_col'], a['to_colOff'], a['to_row'], a['to_rowOff'])
            cx, cy = x1 - x0, y1 - y0
        else:
            cx, cy = a['ext_cx'], a['ext_cy']
        box = {'x': x0, 'y': y0, 'cx': cx, 'cy': cy}
        for obj in anchor:
            otag = obj.tag.split('}')[-1]
            if otag == 'grpSp':
                shapes.extend(_expand_group(obj, a, dict(box)))
            elif otag in ('sp', 'pic', 'graphicFrame', 'cxnSp'):
                kind = {'sp': 'sp', 'pic': 'pic', 'graphicFrame': 'frame', 'cxnSp': 'cxn'}[otag]
                shapes.append(_mk_shape(obj, kind, a, dict(box)))
    return shapes


def _make_xy(data):
    """生成 (col, colOff, row, rowOff) -> (x_emu, y_emu) 的换算函数（0 基下标）。"""
    def xy(col, col_off, row, row_off):
        x = float(col_off or 0)
        for c in range(1, int(col or 0) + 1):
            x += _col_width_to_emu(data.col_w.get(c, data.default_col_w))
        y = float(row_off or 0)
        for r in range(1, int(row or 0) + 1):
            y += data.row_ht.get(r, data.default_row_ht) * EMU_PER_PT
        return (x, y)
    return xy


def _load_xlsx(path, data):
    if not zipfile.is_zipfile(path):
        data.open_error = '文件不是有效的 OOXML(zip) 包，无法作为 .xlsx 打开'
        return data
    tmp = tempfile.mkdtemp(prefix='ov093_')
    try:
        with zipfile.ZipFile(path) as zf:
            bad = zf.testzip()
            if bad:
                data.open_error = '压缩包损坏：%s' % bad
                return data
            for name in zf.namelist():
                if name.startswith('/') or '..' in name.replace('\\', '/').split('/'):
                    continue
                zf.extract(name, tmp)
        wb = _xml(tmp, 'xl/workbook.xml')
        if wb is None:
            data.open_error = '缺少 xl/workbook.xml，不是 Excel 工作簿'
            return data
        rels = _rels_map(tmp, 'xl/workbook.xml')
        sheets = []
        for sh in wb.iterfind('x:sheets/x:sheet', NS):
            part = rels.get(sh.get('{%s}id' % NS['r'])) or ''
            sheets.append((sh.get('name') or '', part))
        data.sheet_names = [n for n, _ in sheets]
        if not sheets:
            data.open_error = '工作簿内未发现任何工作表'
            return data
        st = _parse_styles(tmp)
        shared = _shared_strings(tmp, st)
        first_part = sheets[0][1] or 'xl/worksheets/sheet1.xml'
        dpart = _parse_sheet(tmp, first_part, st, shared, data)
        xy = _make_xy(data)
        if dpart:
            data.shapes = _parse_drawing(tmp, dpart, xy)
        for _, part in sheets[1:]:
            if not part:
                continue
            other = _rels_map(tmp, part)
            root = _xml(tmp, part)
            if root is None:
                continue
            dr = root.find('x:drawing', NS)
            if dr is None:
                continue
            op = other.get(dr.get('{%s}id' % NS['r']))
            if op:
                data.other_shapes += len(_parse_drawing(tmp, op, xy))
        data.pictures = sum(1 for s in data.shapes if s['kind'] == 'pic')
        data.has_drawing_hint = bool(data.shapes) or data.other_shapes > 0
        data.ok_open = True
        return data
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


# --------------------------------------------------------- BIFF 解析路径 ----

_XLS_HALIGN = {0: None, 1: 'left', 2: 'center', 3: 'right', 5: 'justify', 6: 'centerContinuous'}
_XLS_VALIGN = {0: 'top', 1: 'center', 2: 'bottom', 3: 'justify'}


def _load_xls(path, data):
    try:
        import xlrd
    except ImportError:
        data.open_error = '缺少可选依赖 xlrd，无法解析 .xls（BIFF）格式'
        return data
    try:
        bk = xlrd.open_workbook(path, formatting_info=True)
    except Exception as exc:                                  # noqa: BLE001
        data.open_error = '%s: %s' % (type(exc).__name__, exc)
        return data
    cmap = getattr(bk, 'colour_map', {}) or {}

    def col_of(idx):
        v = cmap.get(idx)
        return ('%02X%02X%02X' % v) if v else None

    data.sheet_names = list(bk.sheet_names())
    sh = bk.sheet_by_index(0)
    for r in range(sh.nrows):
        for c in range(sh.ncols):
            xf = bk.xf_list[sh.cell_xf_index(r, c)]
            fnt = bk.font_list[xf.font_index]
            cell = _blank_cell()
            cell['text'] = _norm_text(sh.cell_value(r, c))
            cell['fname'] = fnt.name
            cell['sz'] = fnt.height / 20.0
            cell['bold'] = fnt.weight >= 700
            cell['fcolor'] = col_of(fnt.colour_index)
            solid = xf.background.fill_pattern == 1
            cell['pattern'] = 'solid' if solid else ('none' if xf.background.fill_pattern == 0 else 'other')
            cell['fill'] = col_of(xf.background.pattern_colour_index) if solid else 'FFFFFF'
            for side in ('left', 'right', 'top', 'bottom'):
                ls = getattr(xf.border, '%s_line_style' % side)
                ci = getattr(xf.border, '%s_colour_index' % side)
                cell['borders'][side] = ('thin' if ls == 1 else (None if ls == 0 else 'other'), col_of(ci))
            cell['halign'] = _XLS_HALIGN.get(xf.alignment.hor_align, 'other')
            cell['valign'] = _XLS_VALIGN.get(xf.alignment.vert_align, 'other')
            cell['wrap'] = bool(xf.alignment.text_wrapped)
            ref = '%s%d' % (_idx_to_col(c + 1), r + 1)
            data.cells[ref] = cell
    data.max_row, data.max_col = sh.nrows, sh.ncols
    return _load_xls_tail(bk, sh, path, data)


def _load_xls_tail(bk, sh, path, data):
    """.xls 的合并区域、行高列宽，以及图形记录存在性推断。"""
    for (rlo, rhi, clo, chi) in sh.merged_cells:
        data.merged.add('%s%d:%s%d' % (_idx_to_col(clo + 1), rlo + 1,
                                       _idx_to_col(chi), rhi))
    try:
        data.default_row_ht = sh.default_row_height / 20.0
    except (AttributeError, TypeError):
        pass
    for r, info in (getattr(sh, 'rowinfo_map', {}) or {}).items():
        if getattr(info, 'height', None):
            data.row_ht[r + 1] = info.height / 20.0
    for c, info in (getattr(sh, 'colinfo_map', {}) or {}).items():
        if getattr(info, 'width', None):
            data.col_w[c + 1] = info.width / 256.0
    # xlrd 不解析 Escher/OBJ 记录：只能按 DggContainer 签名判断"是否存在图形"，
    # 具体几何（形状类型/旋转/填充/线条/位置）在该格式下不可判定。
    data.shape_readable = False
    try:
        with open(path, 'rb') as fh:
            raw = fh.read()
        data.has_drawing_hint = b'\x0f\x00\x00\xf0' in raw
    except OSError:
        data.has_drawing_hint = False
    data.ok_open = True
    return data


def load_workbook(path):
    data = WorkbookData()
    data.path = path
    data.ext = os.path.splitext(path)[1].lower()
    try:
        with open(path, 'rb') as fh:
            head = fh.read(8)
        # 以文件头而非扩展名判定真实格式：zipfile.is_zipfile 对二进制文件可能误判
        if head.startswith(b'PK\x03\x04'):
            return _load_xlsx(path, data)
        if head.startswith(b'\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1'):
            return _load_xls(path, data)
        if data.ext in ('.xlsx', '.xlsm', '.xltx', '.xltm'):
            return _load_xlsx(path, data)
        if data.ext == '.xls':
            return _load_xls(path, data)
        data.open_error = '不受支持的扩展名：%s，需 .xlsx 或 .xlsm' % data.ext
        return data
    except Exception as exc:                                  # noqa: BLE001
        data.ok_open = False
        data.open_error = '%s: %s' % (type(exc).__name__, exc)
        return data


# -------------------------------------------------------------- 区域定义 ----


def _refs(c1, r1, c2, r2):
    return ['%s%d' % (_idx_to_col(c), r) for r in range(r1, r2 + 1) for c in range(c1, c2 + 1)]


REG_HEADER = _refs(1, 1, 7, 2)
REG_FOOTER = _refs(1, 33, 7, 33)
REG_DIMLAB = _refs(1, 3, 1, 32)
REG_SUBLAB = _refs(2, 24, 2, 32)
REG_FCOL = _refs(6, 3, 6, 32)
REG_STRIPE = _refs(2, 3, 2, 23) + _refs(3, 3, 3, 32) + _refs(4, 3, 5, 32)
REG_BODY = REG_STRIPE + _refs(7, 3, 7, 32)
REG_ALL = _refs(1, 1, 7, 33)
REG_BLANK = ([r for r in _refs(1, 1, 7, 2) if r not in ('A1', 'D1', 'F1')]
             + _refs(4, 3, 5, 32) + _refs(7, 1, 7, 32) + _refs(2, 33, 7, 33))
C32_COL0, C32_ROW0 = 2, 31          # C32 在绘图坐标系下的 0 基列/行号


def _in_c32(shape):
    return shape.get('col') == C32_COL0 and shape.get('row') == C32_ROW0


def _thin(cell, side, color):
    st, cc = cell['borders'][side]
    return st == 'thin' and _same_color(cc, color)


def _thin4(cell, color):
    return all(_thin(cell, s, color) for s in ('left', 'right', 'top', 'bottom'))


def _bad(data, refs, pred):
    """返回区域内不满足 pred(cell) 的单元格引用列表。"""
    out = []
    for ref in refs:
        cell = data.cells.get(ref) or _blank_cell()
        try:
            ok = pred(cell, ref)
        except Exception:                                     # noqa: BLE001
            ok = False
        if not ok:
            out.append(ref)
    return out


def _brief(items, limit=6):
    items = list(items)
    if not items:
        return '无'
    head = '、'.join(str(i) for i in items[:limit])
    return head + ('…等%d处' % len(items) if len(items) > limit else '')


# ---------------------------------------------------------------- 维度一 ----


def check_dimension1(data):
    """门槛维度：仅检查交付文件格式，返回 (是否通过, [(条目, 是否通过, 说明)])。"""
    items = []
    ok_fmt = data.ok_open and data.ext in ('.xlsx', '.xlsm')
    items.append(('交付1个可正常打开的.xlsx或.xlsm格式Excel表格文件',
                  bool(ok_fmt),
                  '已打开：%s' % os.path.basename(data.path) if ok_fmt
                  else (data.open_error
                        or ('实际格式为 %s，非 .xlsx 或 .xlsm' % data.ext if data.ok_open
                            else '文件缺失或非 .xlsx、.xlsm 格式'))))
    return (bool(ok_fmt), items)


# ---------------------------------------------------------------- 维度二 ----


def check_dimension2(data):
    """返回 (得分, hits)；hits 元素为 (分值, 条目, 是否命中, 子项列表)。"""
    hits = []
    xy = _make_xy(data)
    c32_x0, c32_y0 = xy(C32_COL0, 0, C32_ROW0, 0)
    c32_w = _col_width_to_emu(data.col_w.get(3, data.default_col_w))
    c32_h = data.row_ht.get(32, data.default_row_ht) * EMU_PER_PT
    c32_x1, c32_y1 = c32_x0 + c32_w, c32_y0 + c32_h

    in_c32 = [s for s in data.shapes if _in_c32(s)]
    vec = [s for s in in_c32 if s['kind'] in ('sp', 'cxn')]
    unreadable = (not data.shape_readable) and data.has_drawing_hint

    def gnote():
        if not unreadable:
            return []
        return [('图形几何可解析', False,
                 '当前为 .xls（BIFF）格式，xlrd 无法读取 Escher 图形记录，图形类加分项不予判定')]

    geo = {}
    if len(vec) == 2:
        a, b = sorted(vec, key=lambda s: s['x'])
        geo = {
            'a': a, 'b': b,
            'ax0': a['x'], 'ax1': a['x'] + a['cx'], 'ay0': a['y'], 'ay1': a['y'] + a['cy'],
            'bx0': b['x'], 'bx1': b['x'] + b['cx'], 'by0': b['y'], 'by1': b['y'] + b['cy'],
        }
        geo['ox'] = min(geo['ax1'], geo['bx1']) - max(geo['ax0'], geo['bx0'])
        geo['oy'] = min(geo['ay1'], geo['by1']) - max(geo['ay0'], geo['by0'])
        geo['acx'] = (geo['ax0'] + geo['ax1']) / 2.0
        geo['bcx'] = (geo['bx0'] + geo['bx1']) / 2.0
        geo['acy'] = (geo['ay0'] + geo['ay1']) / 2.0
        geo['bcy'] = (geo['by0'] + geo['by1']) / 2.0
        geo['avg_cx'] = (a['cx'] + b['cx']) / 2.0
        geo['avg_cy'] = (a['cy'] + b['cy']) / 2.0

    # +5 C32 内含有图形，图样为两个五边形相互交叉的组合
    p = gnote() + [
        ('C32单元格内存在图形对象', len(in_c32) > 0,
         'C32 内图形 %d 个' % len(in_c32)),
        ('图形数量为两个', len(vec) == 2,
         'C32 内矢量图形 %d 个：%s' % (len(vec), _brief([s['name'] for s in vec]))),
        ('两个图形相互交叉（存在重叠区）', bool(geo) and geo['ox'] > 0 and geo['oy'] > 0,
         '水平重叠 %.2f 磅 / 垂直重叠 %.2f 磅' % (geo['ox'] / EMU_PER_PT, geo['oy'] / EMU_PER_PT)
         if geo else '不足两个图形，无法判断'),
    ]
    hits.append((5, _RULE_ORDER[0][1], all(x[1] for x in p), p))

    # +3 可编辑的矢量自选图形
    p = gnote() + [
        ('为矢量自选图形（sp/prstGeom 或 custGeom）', len(vec) == 2 and all(
            s['prst'] or s['npts'] >= 3 for s in vec),
         '几何类型：%s' % _brief([s['prst'] or ('custGeom(%d点)' % s['npts']) for s in vec])),
        ('非位图图片', not [s for s in in_c32 if s['kind'] == 'pic'] and len(vec) == 2 and all(
            s['fill'] != 'blip' for s in vec),
         'C32 内位图 %d 个' % len([s for s in in_c32 if s['kind'] == 'pic'])),
        ('非艺术字/文本框拼凑（非 txBox、无文字符号替代）', len(vec) == 2 and not any(
            s['txbox'] or s['text'] for s in vec),
         '文本框标记 %s，图形内文字 %s' % (
             [s['txbox'] for s in vec], _brief([s['text'] for s in vec if s['text']]))),
        ('线条与大小属性可编辑（具备 ln/几何属性）', len(vec) == 2 and all(
            s['ln_w'] is not None or s['ln_fill'] for s in vec),
         '线条属性：%s' % _brief(['w=%s' % s['ln_w'] for s in vec])),
    ]
    hits.append((3, _RULE_ORDER[1][1], all(x[1] for x in p), p))

    # +3 两个图形均为闭合五边形
    pent = [s for s in vec if (s['prst'] or '').lower() == 'pentagon' or
            (s['prst'] is None and s['npts'] == 5 and s['closed'])]
    p = gnote() + [
        ('两个图形的几何均为五边形（pentagon）', len(vec) == 2 and len(pent) == 2,
         '几何：%s' % _brief([s['prst'] or ('custGeom(%d点,闭合=%s)' % (s['npts'], s['closed'])) for s in vec])),
        ('不是菱形/六边形/圆形/星形等其它形状', len(vec) == 2 and not [
            s for s in vec if (s['prst'] or '').lower() in (
                'diamond', 'hexagon', 'ellipse', 'star4', 'star5', 'star6', 'star7',
                'star8', 'star10', 'star12', 'star16', 'star24', 'star32',
                'rect', 'roundrect', 'triangle', 'octagon', 'heptagon', 'homeplate')],
         '异常几何：%s' % _brief([s['prst'] for s in vec if (s['prst'] or '').lower() != 'pentagon'])),
    ]
    hits.append((3, _RULE_ORDER[2][1], all(x[1] for x in p), p))

    # +3 水平并排、部分重叠、中部形成交叉菱形空腔
    p = gnote() + [
        ('两图形垂直方向大致齐平', bool(geo) and abs(geo['acy'] - geo['bcy']) <= 0.4 * geo['avg_cy'],
         '纵向中心差 %.2f 磅' % (abs(geo['acy'] - geo['bcy']) / EMU_PER_PT) if geo else '不足两个图形'),
        ('沿水平方向并排（横向错开）', bool(geo) and abs(geo['acx'] - geo['bcx']) >= 0.2 * geo['avg_cx'],
         '横向中心差 %.2f 磅' % (abs(geo['acx'] - geo['bcx']) / EMU_PER_PT) if geo else '不足两个图形'),
        ('部分重叠而非完全分离或完全叠合', bool(geo) and 0 < geo['ox'] < 0.8 * min(geo['a']['cx'], geo['b']['cx']),
         '重叠宽度 %.2f 磅（较窄图形宽 %.2f 磅）' % (
             geo['ox'] / EMU_PER_PT, min(geo['a']['cx'], geo['b']['cx']) / EMU_PER_PT) if geo else '不足两个图形'),
        ('重叠区位于两图形中部并有垂直交叠', bool(geo) and geo['oy'] > 0 and
         min(geo['acx'], geo['bcx']) <= max(geo['ax0'], geo['bx0']) + geo['ox'] / 2.0 <= max(geo['acx'], geo['bcx']),
         '垂直重叠 %.2f 磅' % (geo['oy'] / EMU_PER_PT) if geo else '不足两个图形'),
    ]
    hits.append((3, _RULE_ORDER[3][1], all(x[1] for x in p), p))

    # +5 两个五边形均为无填充（透明）
    fills = [s['fill'] for s in vec]
    p = gnote() + [
        ('两个五边形均为无填充（noFill）', len(vec) == 2 and all(f == 'none' for f in fills),
         '填充类型：%s' % _brief(['%s%s' % (s['fill'], '/' + (s['fill_rgb'] or '') if s['fill_rgb'] else '')
                              for s in vec]) if vec else '不足两个图形'),
        ('无实心/渐变/图片填充遮挡对方轮廓', len(vec) == 2 and not any(
            f in ('solid', 'grad', 'blip', 'patt') for f in fills),
         '遮挡性填充：%s' % _brief([f for f in fills if f in ('solid', 'grad', 'blip', 'patt')])),
    ]
    hits.append((5, _RULE_ORDER[4][1], all(x[1] for x in p), p))

    # +3 各自带有倾斜旋转角度、相互交错
    rots = [s['rot'] for s in vec]
    rot_l = geo['a']['rot'] if geo else None
    rot_r = geo['b']['rot'] if geo else None
    p = gnote() + [
        ('两个五边形均带有倾斜旋转角度（非正立）', len(vec) == 2 and all(abs(r) >= 5.0 for r in rots),
         '旋转角：%s' % _brief(['%.2f°' % r for r in rots]) if vec else '不足两个图形'),
        ('左侧五边形约顺时针旋转20度', rot_l is not None and 10.0 <= rot_l <= 30.0,
         '左侧旋转角 %.2f°' % rot_l if rot_l is not None else '不足两个图形'),
        ('右侧五边形约逆时针旋转10度', rot_r is not None and -20.0 <= rot_r <= -5.0,
         '右侧旋转角 %.2f°' % rot_r if rot_r is not None else '不足两个图形'),
        ('两者旋转方向相反、姿态相差约30度', rot_l is not None and rot_l * rot_r < 0
         and 20.0 <= abs(rot_l - rot_r) <= 45.0,
         '角度差 %.2f°' % abs(rot_l - rot_r) if rot_l is not None else '不足两个图形'),
    ]
    hits.append((3, _RULE_ORDER[5][1], all(x[1] for x in p), p))

    # +1 黑色单线实线、线宽约 1—1.5 磅、无虚线/阴影/渐变/三维
    p = gnote() + [
        ('轮廓为黑色或近黑色', len(vec) == 2 and all(
            _color_dist(s['ln_rgb'], '000000') <= 64 for s in vec),
         '轮廓色：%s' % _brief([s['ln_rgb'] or '未设置' for s in vec])),
        ('线宽约1—1.5磅', len(vec) == 2 and all(
            s['ln_w'] is not None and 11430 <= s['ln_w'] <= 20955 for s in vec),
         '线宽：%s' % _brief(['%.2f磅' % (s['ln_w'] / EMU_PER_PT) if s['ln_w'] else '未设置' for s in vec])),
        ('单线实线（非复合线、非虚线）', len(vec) == 2 and all(
            (s['ln_cmpd'] in (None, 'sng')) and (s['ln_dash'] in (None, 'solid')) for s in vec),
         '线型：%s' % _brief(['cmpd=%s dash=%s' % (s['ln_cmpd'], s['ln_dash']) for s in vec])),
        ('无阴影、无渐变、无三维效果', len(vec) == 2 and not any(
            s['has_effect'] or s['has_3d'] or s['ln_fill'] not in (None, 'solid') for s in vec),
         '阴影/三维：%s' % _brief(['effect=%s 3d=%s lnFill=%s' % (
             s['has_effect'], s['has_3d'], s['ln_fill']) for s in vec])),
    ]
    hits.append((1, _RULE_ORDER[6][1], all(x[1] for x in p), p))

    # +5 图形水平起点位于"30、"序号文字右侧
    c32 = data.cells.get('C32') or _blank_cell()
    text_w = _est_text_width_emu(EXPECT_C32_TEXT, c32['sz'] or 10)
    left = min([s['x'] for s in in_c32]) if in_c32 else None
    p = gnote() + [
        ('C32内含"30、"序号文字', EXPECT_C32_TEXT in c32['text'],
         'C32 文字=%r' % c32['text']),
        ('图形水平起点在序号文字右侧', left is not None and left >= c32_x0 + text_w,
         '图形左边距单元格左边界 %.2f 磅，"30、"预计占宽 %.2f 磅' % (
             (left - c32_x0) / EMU_PER_PT, text_w / EMU_PER_PT) if left is not None else '未发现图形'),
    ]
    hits.append((5, _RULE_ORDER[7][1], all(x[1] for x in p), p))

    # +5 图形整体完整落在 C32 单元格内部
    anchored = bool(in_c32) and len(data.shapes) == len(in_c32)
    inside_x = bool(in_c32) and min(s['x'] for s in in_c32) > c32_x0 and \
        max(s['x'] + s['cx'] for s in in_c32) < c32_x1
    inside_y = bool(in_c32) and min(s['y'] for s in in_c32) > c32_y0 and \
        max(s['y'] + s['cy'] for s in in_c32) < c32_y1
    same_cell = bool(in_c32) and all(
        (s['to_col'] is None or s['to_col'] == C32_COL0) and
        (s['to_row'] is None or s['to_row'] == C32_ROW0) for s in in_c32)
    p = gnote() + [
        ('图形锚定于C32单元格', anchored and same_cell,
         '锚定单元格：%s' % _brief(['%s%d' % (_idx_to_col((s['col'] or 0) + 1), (s['row'] or 0) + 1)
                                for s in data.shapes]) if data.shapes else '未发现图形'),
        ('左右不越出C列边界', inside_x,
         '左边距 %.2f 磅、右边距 %.2f 磅' % (
             (min(s['x'] for s in in_c32) - c32_x0) / EMU_PER_PT,
             (c32_x1 - max(s['x'] + s['cx'] for s in in_c32)) / EMU_PER_PT) if in_c32 else '未发现图形'),
        ('上下不越出第32行边界', inside_y,
         '上边距 %.2f 磅、下边距 %.2f 磅' % (
             (min(s['y'] for s in in_c32) - c32_y0) / EMU_PER_PT,
             (c32_y1 - max(s['y'] + s['cy'] for s in in_c32)) / EMU_PER_PT) if in_c32 else '未发现图形'),
    ]
    hits.append((5, _RULE_ORDER[8][1], all(x[1] for x in p), p))

    # ------------------------------------------------------------ 扣分项 ----
    # 子项统一采用"True = 符合要求"的极性；只要有一项不符合即命中扣分。
    def add_minus(idx, subs):
        hits.append((_RULE_ORDER[idx][0], _RULE_ORDER[idx][1],
                     not all(s[1] for s in subs), subs))

    add_minus(9, [
        ('工作簿为且仅为"Sheet0""Sheet1"两个工作表（Sheet0在前）',
         data.sheet_names == ['Sheet0', 'Sheet1'], '实际工作表：%s' % _brief(data.sheet_names)),
        ('Sheet0数据区域为A1:G33（33行7列）', data.max_row == 33 and data.max_col == 7,
         '实际数据区域 %d行×%d列' % (data.max_row, data.max_col)),
    ])

    miss = sorted(EXPECT_MERGES - data.merged)
    extra = sorted(data.merged - EXPECT_MERGES)
    add_minus(10, [
        ('33处合并单元格齐备', not miss, '缺失：%s' % _brief(miss)),
        ('未出现额外的合并区域', not extra, '多余：%s' % _brief(extra)),
    ])

    diff = [k for k, v in sorted(EXPECT_HEAD.items()) if data.text(k) != v]
    add_minus(11, [
        ('A1、D1、F1三格文字依次为"观察维度""观察记录""完成情况"', not diff,
         '不符：%s' % _brief(['%s=%r' % (k, data.text(k)) for k in diff])),
    ])

    diff = [(k, v) for k, v in EXPECT_DIM_LABELS if data.text(k) != v]
    add_minus(12, [
        ('A3、A13、A16、A21、A24五格为对应的两行维度标签文字', not diff,
         '不符：%s' % _brief(['%s=%r' % (k, data.text(k)) for k, _ in diff])),
        ('A33文字为"综合结果"', data.text('A33') == '综合结果', 'A33=%r' % data.text('A33')),
    ])

    for idx, col, r1, expect in ((13, 2, 3, EXPECT_B3_B12),
                                 (14, 2, 13, EXPECT_B13_B23),
                                 (15, 3, 24, EXPECT_C24_C31)):
        refs = ['%s%d' % (_idx_to_col(col), r1 + i) for i in range(len(expect))]

        diff = [(ref, expect[i]) for i, ref in enumerate(refs) if data.text(ref) != expect[i]]
        add_minus(idx, [
            ('%s至%s共%d格题目文字与细则完全一致' % (refs[0], refs[-1], len(refs)), not diff,
             '不符：%s' % _brief(['%s=%r' % (ref, data.text(ref)) for ref, _ in diff])),
        ])

    diff = [(k, v) for k, v in EXPECT_SUB_LABELS if data.text(k) != v]
    add_minus(16, [
        ('B24、B26、B27、B30、B31、B32六格小类标签文字正确', not diff,
         '不符：%s' % _brief(['%s=%r' % (k, data.text(k)) for k, _ in diff])),
    ])

    add_minus(17, [
        ('C32单元格文字为"30、"（含顿号）', data.text('C32') == EXPECT_C32_TEXT,
         'C32=%r' % data.text('C32')),
    ])

    diff = [r for r in REG_FCOL if data.text(r) != EXPECT_F_TEXT]
    add_minus(18, [
        ('F3至F32共30格文字均为"否  是"（两个空格）', not diff,
         '不符：%s' % _brief(['%s=%r' % (r, data.text(r)) for r in diff])),
    ])

    filled = [r for r in REG_BLANK if data.text(r)]
    add_minus(19, [
        ('应为空白的表头空位、D/E列、G列与末行区域内无文字', not filled,
         '出现文字：%s' % _brief(['%s=%r' % (r, data.text(r)) for r in filled])),
    ])

    def fmt_subs(refs, fname, size, bold, fcolor, fill, border, label):
        """生成"字体/字号/加粗/字色/填充/边框"六项通用格式子项。

        字号/加粗/字色按实际显示效果判定：单元格内若有富文本片段，
        以片段自身声明的属性为准（片段属性会覆盖单元格字体）。
        """
        subs = []
        if fname is not None:
            b = _bad(data, refs, lambda c, _r: (c['fname'] or '') == fname)
            subs.append(('%s字体为%s' % (label, fname), not b, '不符：%s' % _brief(b)))
        b = _bad(data, refs, lambda c, _r: all(
            v is not None and abs(v - size) < 0.05 for v in _eff(c, 'sz')))
        subs.append(('%s字号为%g磅' % (label, size), not b, '不符：%s' % _brief(b)))
        b = _bad(data, refs, lambda c, _r: all(bool(v) == bold for v in _eff(c, 'bold')))
        subs.append(('%s%s' % (label, '加粗' if bold else '不加粗'), not b, '不符：%s' % _brief(b)))
        if fcolor is not None:
            b = _bad(data, refs, lambda c, _r: all(
                _same_color(v, fcolor) for v in _eff(c, 'fcolor')))
            subs.append(('%s文字颜色为%s' % (label, fcolor), not b, '不符：%s' % _brief(b)))
        if fill is not None:
            b = _bad(data, refs, lambda c, _r: _same_color(c['fill'], fill))
            subs.append(('%s填充为%s' % (label, fill), not b, '不符：%s' % _brief(b)))
        if border is not None:
            b = _bad(data, refs, lambda c, _r: _thin4(c, border))
            subs.append(('%s边框为%s细实线' % (label, border), not b, '不符：%s' % _brief(b)))
        return subs

    subs = fmt_subs(REG_HEADER, None, 10, True, C_WHITE, C_HDR_FILL, C_HDR_BORDER, 'A1:G2表头区')
    b = [r for r in ('A1', 'D1', 'F1') if ((data.cells.get(r) or _blank_cell())['fname'] or '') != FONT_LABEL]
    subs.append(('A1、D1、F1三格字体为%s' % FONT_LABEL, not b, '不符：%s' % _brief(b)))
    add_minus(20, subs)

    add_minus(21, fmt_subs(REG_FOOTER, None, 10, True, C_WHITE, C_FOOT_FILL, C_HDR_BORDER, 'A33:G33'))
    add_minus(22, fmt_subs(REG_DIMLAB, FONT_LABEL, 10, True, C_DIMLABEL_FONT, C_DIMLABEL_FILL,
                           C_LABEL_BORDER, 'A3:A32维度标签'))
    add_minus(23, fmt_subs(REG_SUBLAB, FONT_LABEL, 10, True, C_SUBLABEL_FONT, C_SUBLABEL_FILL,
                           C_LABEL_BORDER, 'B24:B32小类标签'))
    add_minus(24, fmt_subs(REG_BODY, FONT_BODY, 10, False, C_BODY_FONT, None,
                           C_LABEL_BORDER, '题目与记录区'))

    b = _bad(data, REG_STRIPE, lambda c, r: _same_color(
        c['fill'], C_WHITE if _split_ref(r)[0] % 2 else C_STRIPE_EVEN))
    b2 = _bad(data, _refs(7, 3, 7, 32), lambda c, _r: _same_color(c['fill'], C_WHITE))
    add_minus(25, [
        ('B3:B23、C3:C32、D3:E32隔行底色为奇数行FFFFFF、偶数行F4F8F8', not b, '不符：%s' % _brief(b)),
        ('G3至G32底色统一为FFFFFF', not b2, '不符：%s' % _brief(b2)),
    ])

    add_minus(26, fmt_subs(REG_FCOL, FONT_LABEL, 10, True, C_F_FONT, C_F_FILL, C_F_BORDER, 'F3:F32'))

    center = REG_HEADER + REG_DIMLAB + REG_SUBLAB + REG_FCOL + REG_FOOTER
    general = _refs(2, 3, 2, 23) + _refs(3, 24, 3, 32)
    b = _bad(data, center, lambda c, _r: c['halign'] == 'center')
    b2 = _bad(data, general, lambda c, _r: c['halign'] in (None, 'general', 'left'))
    add_minus(27, [
        ('表头区、A列维度标签、小类标签、F列、末行水平居中', not b, '不符：%s' % _brief(b)),
        ('B3:B23与C24:C32题目文字为常规（靠左）', not b2, '不符：%s' % _brief(b2)),
    ])

    b = _bad(data, REG_ALL, lambda c, _r: c['valign'] == 'center')
    b2 = _bad(data, REG_ALL, lambda c, _r: bool(c['wrap']))
    add_minus(28, [
        ('全表单元格垂直居中', not b, '不符：%s' % _brief(b)),
        ('全表单元格开启自动换行', not b2, '不符：%s' % _brief(b2)),
    ])

    bad_ht = []
    for rno, want in sorted(EXPECT_ROW_HT.items()):
        got = data.row_ht.get(rno, data.default_row_ht)
        if abs(got - want) > ROW_HT_TOL:
            bad_ht.append('第%d行=%.2f磅(应%.1f)' % (rno, got, want))
    add_minus(29, [
        ('第1至29行及第31行约19.2磅、第30行约38.4磅、第32行42磅、第33行44磅', not bad_ht,
         '不符：%s' % _brief(bad_ht)),
    ])

    bad_w = []
    for cno, want in sorted(EXPECT_COL_W.items()):
        got = data.col_w.get(cno, data.default_col_w)
        if abs(got - want) > COL_W_TOL:
            bad_w.append('%s列=%.2f(应%.2f)' % (_idx_to_col(cno), got, want))
    add_minus(30, [
        ('A至G列列宽依次约7.91/10.70/32.71/7.61/9.35/16.60/7.98字符', not bad_w,
         '不符：%s' % _brief(bad_w)),
    ])

    # 图形类扣分项：.xls 下几何不可解析时不做"违规"判定，避免误扣
    judgeable = data.shape_readable
    outside = [s for s in data.shapes if not _in_c32(s)]
    add_minus(31, [
        ('C32以外的题号行与"完成情况"列不存在图形对象',
         (not judgeable) or not outside,
         '越位图形：%s' % _brief(['%s@%s%d' % (s['name'], _idx_to_col((s['col'] or 0) + 1),
                                          (s['row'] or 0) + 1) for s in outside])),
        ('Sheet1中不存在图形对象', (not judgeable) or data.other_shapes == 0,
         '其它工作表图形 %d 个' % data.other_shapes),
    ])

    texted = [s for s in in_c32 if s['text']]
    add_minus(32, [
        ('图形内部不含文字内容，未叠加"30"等字符',
         (not judgeable) or not texted,
         '含文字图形：%s' % _brief(['%s=%r' % (s['name'], s['text']) for s in texted])),
    ])

    over = ['%s(%.2f×%.2f磅)' % (s['name'], s['cx'] / EMU_PER_PT, s['cy'] / EMU_PER_PT)
            for s in in_c32 if s['cy'] > c32_h * 1.02 or s['cx'] > c32_w * 1.02]
    add_minus(33, [
        ('图形高度不超过第32行行高42磅、宽度不超过C列列宽',
         (not judgeable) or not over,
         '超限图形：%s（第32行%.1f磅／C列%.1f磅）' % (_brief(over), c32_h / EMU_PER_PT, c32_w / EMU_PER_PT)),
    ])

    not_two = [s for s in data.shapes if s['anchor'] != 'two']
    add_minus(34, [
        ('若已插入图形，其锚定方式为随单元格位置与大小变化（两点锚定）',
         (not judgeable) or not data.shapes or not not_two,
         '非两点锚定：%s' % _brief(['%s=%s' % (s['name'], s['anchor']) for s in not_two])),
    ])

    score = sum(pts for pts, _rule, hit, _sub in hits if hit)
    return (score, hits)


# ------------------------------------------------------------ 结果组装 ----


def _build_dim2_items(hits):
    """hits -> dim2_items；顺序与分值必须与 _RULE_ORDER 完全一致。"""
    if len(hits) != len(_RULE_ORDER):
        raise RuntimeError('dim2 项数不一致：hits=%d vs order=%d' % (len(hits), len(_RULE_ORDER)))
    items = []
    for (pts, rule, hit, _subs), (max_delta, _rule_text) in zip(hits, _RULE_ORDER):
        if pts != max_delta:
            raise RuntimeError('dim2 分值错位：%r 实为 %d，应为 %d' % (rule, pts, max_delta))
        items.append({'rule': rule, 'max_delta': max_delta,
                      'delta': max_delta if hit else 0, 'hit': bool(hit)})
    return items


def _find_target_file(dir_path):
    """在目录内定位被评估的工作簿；找不到返回空串（由调用方判定为“没评上”）。
    .xls 也会被拾取，只为把它交给维度一记 0，而不是当作“目录里没有交付物”。"""
    exts = ('.xlsx', '.xlsm', '.xls')
    cands = [f for f in sorted(os.listdir(dir_path))
             if f.lower().endswith(exts) and not f.startswith('~$')]
    if not cands:
        return ''
    # 只要交付中出现 .xls 就优先取它：.xls 格式一律维度1不通过
    cands.sort(key=lambda f: (0 if f.lower().endswith('.xls') else 1, f))
    return os.path.join(dir_path, cands[0])


def _dim1_fail(result, reason):
    """维度一未通过：完成度总分记 0，不再逐条罗列（未真正逐条评过）。"""
    result['dim1_pass'] = False
    result['dim1_reason'] = reason or '维度一未通过（未给出具体原因）'
    result['dim2_items'] = []
    result['total_score'] = 0
    result['max_score'] = MAX_SCORE_TOTAL
    return result


def evaluate(dir_path):
    """按接口统一约定返回结构化 dict；不 print、不改 sys.stdout、不 sys.exit。

    返回信息约定：
      status="error"（目录不存在、目录中没有交付文件、脚本自身异常）表示这份作业
      根本没评上，此时 dim1_pass 为空串（未判定）、dim1_reason 为空串、
      dim2_items 为 []、total_score 与 max_score 均为 0；
      维度一未通过时 status 仍为 "ok"，dim1_pass 为 False、dim1_reason 写错误摘要、
      dim2_items 为 []（未真正逐条评过）、total_score 为 0、max_score 为满分；
      维度一通过时 dim1_pass 为 True、dim1_reason 保持空串。
      dim2_items 每条只含 rule / max_delta / delta / hit 四个键。"""
    result = {
        'id': SCRIPT_ID, 'file_name': '', 'status': 'ok', 'error': None,
        'dim1_pass': '', 'dim1_reason': '', 'dim2_items': [],
        'total_score': 0, 'max_score': 0,
    }
    try:
        if not dir_path or not os.path.isdir(dir_path):
            result['status'] = 'error'
            result['error'] = '目录不存在：%s' % dir_path
            return result
        file_path = _find_target_file(dir_path)
        if not file_path:
            result['status'] = 'error'
            result['error'] = '目录中未找到 .xlsx/.xlsm 文件：%s' % dir_path
            return result
        result['file_name'] = os.path.basename(file_path)
        data = load_workbook(file_path)
        passed, d1 = check_dimension1(data)
        if not passed:
            return _dim1_fail(result, '；'.join(
                '%s：%s' % (name, info) for name, ok, info in d1 if not ok))
        result['dim1_pass'] = True
        # 约定：维度一通过时 dim1_reason 保持空串，只有未通过才写错误摘要。
        result['dim1_reason'] = ''
        score, hits = check_dimension2(data)
        result['dim2_items'] = _build_dim2_items(hits)
        result['total_score'] = max(0, score)   # 扣分多于加分时兜底为 0
        result['max_score'] = MAX_SCORE_TOTAL
        return result
    except Exception as exc:                                  # noqa: BLE001
        # 脚本异常同样视为“没评上”：维度一未判定，两项分数归零。
        result['status'] = 'error'
        result['error'] = '%s: %s' % (type(exc).__name__, exc)
        result['dim1_pass'] = ''
        result['dim1_reason'] = ''
        result['dim2_items'] = []
        result['total_score'] = 0
        result['max_score'] = 0
        return result


# 说明：按接口统一约定，本模块不修改全局 sys.stdout，也不 print 主结果。
# 主结果统一由 evaluate(dir_path) -> dict 返回；以下 __main__ 仅用于本地调试。
if __name__ == '__main__':
    _args = [a for a in sys.argv[1:] if not a.startswith('-')]
    _dir = _args[0] if _args else os.path.dirname(os.path.abspath(__file__))
    _res = evaluate(_dir)
    if '-v' in sys.argv or '--verbose' in sys.argv:
        _target = _find_target_file(_dir) if os.path.isdir(_dir) else ''
        if not _target:
            # 没评上：只有 error 可看，维度一未判定
            print('未评估：%s' % _res['error'])
            sys.exit(0)
        _data = load_workbook(_target)
        _p1, _d1 = check_dimension1(_data)
        print('文件：%s' % _res['file_name'])
        for _n, _ok, _info in _d1:
            print('  [%s] %s —— %s' % ('通过' if _ok else '不通过', _n, _info))
        if _p1:
            _sc, _hits = check_dimension2(_data)
            for _pts, _rule, _hit, _subs in _hits:
                print('%s %+d %s' % ('★' if _hit else '·', _pts if _hit else 0, _rule[:48]))
                for _sn, _sok, _sinfo in _subs:
                    print('      [%s] %s —— %s' % ('√' if _sok else '×', _sn, _sinfo))
            print('合计：%d / %d' % (_sc, MAX_SCORE_TOTAL))
    else:
        print(json.dumps(_res, ensure_ascii=False, indent=2))
