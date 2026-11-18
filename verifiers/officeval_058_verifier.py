#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# pyright: reportUnknownMemberType=false, reportUnknownVariableType=false, reportUnknownArgumentType=false, reportUnknownParameterType=false, reportMissingParameterType=false, reportPrivateUsage=false, reportUnknownLambdaType=false, reportGeneralTypeIssues=false

import os
import re
import sys
import json
import zipfile
import colorsys
import xml.etree.ElementTree as ET

# 说明：按接口统一约定，本模块不修改全局 sys.stdout，也不 print 主结果。
# 主结果统一由 evaluate(dir_path) -> dict 返回；入参为"脚本所在目录的路径"，
# 由脚本自己在该目录内定位并打开被评估的 .pptx。本文件仅在 __main__ 下作为
# 本地调试入口，才把 dict 序列化为 JSON 打印。
#
# 被评估任务：给"中国电影业面临的挑战"演示文稿的目录页（第2页）四组小标题下方
# 各补 3 个小标签（共 12 个），圆角胶囊形、深绿底白字、MiSans、10.5 磅、
# 自左至右等距排列、大小一致。

SCRIPT_ID = "058"
# 维度二所有加分项 max_delta 之和；用于维度一未通过 / 脚本异常时兜底
MAX_SCORE_TOTAL = 45

# 维度二所有评分项按"结果追加顺序"的 max_delta 列表；末端把 results 里的
# "+N：..." / "-N：..." 字符串反解为结构化 dim2_items 时用。顺序须与
# _evaluate_impl 中 results.append(...) 的顺序严格一致；共 35 项，
# 其中正值 15 项累加 = MAX_SCORE_TOTAL，负值 20 项为减分项。
_MAX_DELTAS_ORDER = [
    5,   # 01 12个标签齐全、分组对应、无错漏、全在目录页
    5,   # 02 中文实际生效 MiSans（a:ea 与 a:latin 同时指定）
    3,   # 03 标签为可单独点选编辑的形状
    3,   # 04 字号统一 10.5 磅
    3,   # 05 底色深绿且12个统一
    3,   # 06 文字白色
    3,   # 07 圆角胶囊形（adj≥40%）
    3,   # 08 同组自左至右、顶边齐平
    3,   # 09 等距 + 同列对齐 + 行距与小标题一致（规整矩阵）
    3,   # 10 12个尺寸一致
    5,   # 11 落在该组小标题下方空白区
    3,   # 12 与既有元素互不重叠
    1,   # 13 色块尺寸与文字匹配、文字居中
    1,   # 14 扁平纯色块（无描边、无效果）
    1,   # 15 常规字重纯文本
    -1,  # 16 目录页页首 CONTENTS / 目录
    -1,  # 17 目录页 01-04 序号
    -1,  # 18 目录页金色短横线与竖向分隔线
    -3,  # 19 目录页四个小标题
    -3,  # 20 目录页四条说明文字
    -3,  # 21 目录页右侧配图
    -3,  # 22 目录页右下角总结句
    -1,  # 23 目录页页脚细线与页码 02
    -3,  # 24 添加了指令未要求的元素
    -3,  # 25 第1页封面元素
    -3,  # 26 第3页标题与四张卡片
    -3,  # 27 第4页标题与四张卡片
    -3,  # 28 第5页标题与三张卡片
    -3,  # 29 第6页标题与三个圆点
    -3,  # 30 第7页深色页元素
    -3,  # 31 第8页资料条目
    -3,  # 32 页数/页序/页码标注
    -1,  # 33 页面底色与页面尺寸
    -1,  # 34 字体主题 / 配色主题
    -3,  # 35 出现乱码
]
assert sum(d for d in _MAX_DELTAS_ORDER if d > 0) == MAX_SCORE_TOTAL


# ==================== OOXML 命名空间与单位 ====================
NS_A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
NS_P = 'http://schemas.openxmlformats.org/presentationml/2006/main'
NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
NS_PKG = 'http://schemas.openxmlformats.org/package/2006/relationships'

EMU_CM = 360000.0          # 1cm = 360000 EMU
EMU_PT = 12700.0           # 1磅 = 12700 EMU
TOL_002 = int(0.02 * EMU_CM)   # 尺寸一致容差 0.02cm
TOL_005 = int(0.05 * EMU_CM)   # 对齐/等距容差 0.05cm
TOL_015 = int(0.15 * EMU_CM)   # 行距一致容差 0.15cm
TOL_POS = int(0.05 * EMU_CM)   # 既有元素"未被移动"的位置容差

# ==================== 指令要求的 12 个小标签 ====================
# (小标题, [自左至右三个标签文字])
LABEL_GROUPS = [
    ('文化沟通', ['儒家伦理', '叙事逻辑', '文化折扣']),
    ('工业体系', ['特效能力', '宣发困境', '产业链断点']),
    ('全球竞争', ['好莱坞', '韩国电影', '发行网络']),
    ('认知偏差', ['预设立场', '误解误读', '题材受限']),
]
LABEL_WORDS = [w for _, ws in LABEL_GROUPS for w in ws]
LABEL_SLIDE = 2            # 目录页为第2页
LABEL_SZ = 1050            # 10.5磅
LABEL_FONT = 'MiSans'
LABEL_ADJ_MIN = 40000      # 圆角调节值 ≥40%
LABEL_CY_MIN = int(0.55 * EMU_CM)
LABEL_CY_MAX = int(0.90 * EMU_CM)
BAND_GAP_MIN = int(0.35 * EMU_CM)
BAND_GAP_MAX = int(2.00 * EMU_CM)

# ==================== 目录页四组小标题 / 分隔线基准坐标 ====================
HEADING_XY = [(2788920, 1737360), (2788920, 2697480),
              (2788920, 3657600), (2788920, 4617720)]
HEADING_EXT = (1691640, 292608)
HAIRLINE_Y = 6473952       # 页面底部横向细分隔线纵坐标

# ==================== 字体别名（不同办公软件保存后的同义名） ====================
FONT_ALIAS = {
    '微软雅黑': 'Microsoft YaHei',
    'microsoft yahei': 'Microsoft YaHei',
    'msyh': 'Microsoft YaHei',
    'misans': 'MiSans',
    '小米兰亭': 'MiSans',
}

# ==================== 基准：8页页面底色 / 顶层形状数 / 主题 ====================
BG_COLORS = ['F4EFE6', 'E9F0EC', 'F4EFE6', 'F1E8DA',
             'EAF1F2', 'F5EEE8', '112A35', 'EFE6D7']
SHAPE_COUNTS = [9, 23, 19, 20, 16, 17, 24, 15]
THEME_SCHEME = {
    'dk1': '000000', 'lt1': 'FFFFFF', 'dk2': '44546A', 'lt2': 'E7E6E6',
    'accent1': '4472C4', 'accent2': 'ED7D31', 'accent3': 'A5A5A5',
    'accent4': 'FFC000', 'accent5': '5B9BD5', 'accent6': '70AD47',
    'hlink': '0563C1', 'folHlink': '954F72',
}
THEME_FONT = 'Microsoft YaHei'
SLD_CX_CM = (33.5, 34.2)
SLD_CY_CM = (18.8, 19.3)

# ==================== 基准：各页页码文本框 ====================
# (页序, x, y, cx, cy, 文字, 字号, 字色)
PAGE_NUMBERS = [
    (1, 11247120, 6345936, 411480, 228600, '01', 950, '65737B'),
    (2, 11247120, 6345936, 411480, 228600, '02', 950, '65737B'),
    (3, 11247120, 6345936, 411480, 228600, '03', 950, '65737B'),
    (4, 11247120, 6345936, 411480, 228600, '04', 950, '65737B'),
    (5, 11247120, 6345936, 411480, 228600, '05', 950, '65737B'),
    (6, 11247120, 6345936, 411480, 228600, '06', 950, '65737B'),
    (7, 5532120, 6345936, 320040, 228600, '07', 950, '9FB4B6'),
    (8, 11247120, 6345936, 411480, 228600, '08', 950, '65737B'),
]


# ==================== 基准：必须保持原样的既有元素签名 ====================
# 每条签名 = (页序, x, y, cx, cy, 期望属性)；按 (x, y, cx, cy) 在该页定位形状，
# 再逐项比对期望属性。属性键：text 文字、sz 字号(百分磅)、b 加粗、clr 字色、
# font 字体、fill 填充色、lnw 线宽、lnclr 线色、geom 几何、kind 形状类别。
def _sig(slide, x, y, cx, cy, **props):
    return (slide, x, y, cx, cy, props)


_S2_ROW_Y = [1755648, 2715768, 3675888, 4636008]     # 序号/说明文字行纵坐标
_S2_RULE_Y = [1975104, 2935224, 3895344, 4855464]    # 金色短横线纵坐标
_S2_DESCS = ['本土表达跨语境传播的第一道门槛', '创作热情与标准流程之间的磨合',
             '成熟市场、平台渠道与类型片规则的挤压', '外部视角如何影响作品理解与进入机会']

PROTECTED = {
    # ---- 目录页页首 ----
    's2_header': [
        _sig(2, 658368, 502920, 1828800, 256032, text='CONTENTS', sz=1100, b=1,
             clr='C77A42', font=THEME_FONT),
        _sig(2, 658368, 868680, 2011680, 502920, text='目录', sz=3100, b=1,
             clr='112A35', font=THEME_FONT),
    ],
    # ---- 目录页 01-04 序号 ----
    's2_numbers': [
        _sig(2, 786384, y, 640080, 384048, text='%02d' % (i + 1), sz=2000, b=1,
             clr='D9A441', font=THEME_FONT)
        for i, y in enumerate(_S2_ROW_Y)
    ],
    # ---- 目录页金色短横线 + 右侧竖向金色分隔线 ----
    's2_goldlines': [
        _sig(2, 1572768, y, 1051560, 0, geom='line', lnw=19050, lnclr='D9A441')
        for y in _S2_RULE_Y
    ] + [_sig(2, 6510528, 1627632, 0, 4389120, geom='line', lnw=25400, lnclr='D9A441')],
    # ---- 目录页四个小标题 ----
    's2_headings': [
        _sig(2, x, y, HEADING_EXT[0], HEADING_EXT[1], text=LABEL_GROUPS[i][0],
             sz=1700, b=1, clr='112A35', font=THEME_FONT)
        for i, (x, y) in enumerate(HEADING_XY)
    ],
    # ---- 目录页四条说明文字 ----
    's2_descs': [
        _sig(2, 4526280, _S2_ROW_Y[i], 1920240, 301752, text=_S2_DESCS[i],
             sz=1200, clr='65737B', font=THEME_FONT)
        for i in range(4)
    ],
    # ---- 目录页右侧配图 ----
    's2_pic': [_sig(2, 6748272, 1611493, 4736592, 2665749, kind='pic')],
    # ---- 目录页右下角总结句 ----
    's2_summary': [
        _sig(2, 6839712, 4718304, 4434840, 603504, sz=1600, b=1, clr='1E5B64',
             font=THEME_FONT,
             text='四个维度共同构成影片海外传播的主要压力面：理解、生产、渠道、评价。'),
    ],
    # ---- 目录页页脚：底部细分隔线 + 页码 02 ----
    's2_footer': [
        _sig(2, 594360, HAIRLINE_Y, 10469880, 0, geom='line', lnw=7620, lnclr='D6C9B8'),
        _sig(2, 11247120, 6345936, 411480, 228600, text='02', sz=950, clr='65737B',
             font=THEME_FONT, algn='r'),
    ],
}

# ---- 第1页封面 ----
PROTECTED['s1'] = [
    _sig(1, 658368, 676656, 5212080, 274320, sz=1100, b=1, clr='C77A42',
         font=THEME_FONT, text='CHINESE FILM INDUSTRY CHALLENGES'),
    _sig(1, 640080, 1417320, 5257800, 1325880, sz=3400, b=1, clr='112A35',
         font=THEME_FONT, text='中国电影业面临的挑战'),
    _sig(1, 658368, 3346704, 5212080, 329184, sz=1450, clr='1E5B64',
         font=THEME_FONT, text='文化沟通 · 工业体系 · 全球竞争 · 认知偏差'),
    _sig(1, 658368, 5422392, 2011680, 274320, sz=1100, clr='65737B',
         font=THEME_FONT, text='行业观察稿'),
    _sig(1, 6336792, 0, 146304, 6858000, fill='D9A441'),
    _sig(1, 6492240, 1825176, 5699455, 3207648, kind='pic'),
]

# ---- 第3页：01 CULTURAL COMMUNICATION + 四张 FFF9EF 圆角卡片 ----
_S3_CARD_XY = [(685800, 1755648), (3977640, 1755648), (685800, 3767328), (3977640, 3767328)]
_S3_TITLE_XY = [(850392, 1883664), (4142232, 1883664), (850392, 3895344), (4142232, 3895344)]
_S3_TITLES = ['文化语境落差', '叙事节奏差异', '语言幽默损耗', '情感转译压力']
PROTECTED['s3'] = [
    _sig(3, 658368, 384048, 1143000, 274320, sz=1100, b=1, clr='D9A441',
         font=THEME_FONT, text='01  CULTURAL COMMUNICATION'),
] + [
    _sig(3, x, y, 3063240, 1600200, geom='roundRect', fill='FFF9EF')
    for x, y in _S3_CARD_XY
] + [
    _sig(3, x, y, 2734056, 292608, sz=1300, b=1, clr='112A35', font=THEME_FONT,
         text=_S3_TITLES[i])
    for i, (x, y) in enumerate(_S3_TITLE_XY)
]

# ---- 第4页：02 INDUSTRIAL CAPABILITY + 四张 FFF5E8 圆角卡片 + 核心判断 ----
_S4_CARD_XY = [(5742432, 1682496), (8668512, 1682496), (5742432, 3255264), (8668512, 3255264)]
_S4_TITLE_XY = [(5907024, 1810512), (8833104, 1810512), (5907024, 3383280), (8833104, 3383280)]
_S4_TITLES = ['制作端', '宣发端', '触达端', '协同端']
PROTECTED['s4'] = [
    _sig(4, 658368, 384048, 1143000, 274320, sz=1100, b=1, clr='D9A441',
         font=THEME_FONT, text='02  INDUSTRIAL CAPABILITY'),
] + [
    _sig(4, x, y, 2697480, 1280160, geom='roundRect', fill='FFF5E8')
    for x, y in _S4_CARD_XY
] + [
    _sig(4, x, y, 2368296, 292608, sz=1300, b=1, clr='112A35', font=THEME_FONT,
         text=_S4_TITLES[i])
    for i, (x, y) in enumerate(_S4_TITLE_XY)
] + [
    _sig(4, 5806440, 5138928, 5349240, 530352, sz=1500, b=1, clr='1E5B64',
         font=THEME_FONT,
         text='核心判断：不是缺少优秀单点能力，而是需要更稳定、可复制、可协作的工业链条。'),
]

# ---- 第5页：03 GLOBAL COMPETITION + 三张 FFF9EF 圆角卡片 + 左侧配图 ----
_S5_TITLES = ['成熟工业体系', '区域市场崛起', '中国内容压力']
PROTECTED['s5'] = [
    _sig(5, 658368, 384048, 1143000, 274320, sz=1100, b=1, clr='D9A441',
         font=THEME_FONT, text='03  GLOBAL COMPETITION'),
    _sig(5, 658368, 1573659, 5504688, 3098033, kind='pic'),
] + [
    _sig(5, 6446520, y, 5074920, 1097280, geom='roundRect', fill='FFF9EF')
    for y in (1554480, 2889504, 4224528)
] + [
    _sig(5, 6611112, y, 4745736, 292608, sz=1300, b=1, clr='1E5B64',
         font=THEME_FONT, text=_S5_TITLES[i])
    for i, y in enumerate((1682496, 3017520, 4352544))
]

# ---- 第6页：04 PERCEPTION BARRIERS + 三个圆点 + 三段小标题 ----
_S6_TITLES = ['预设立场', '文化陌生', '双重审视']
_S6_DOT_FILL = ['D9A441', '1E5B64', 'C85C4A']
PROTECTED['s6'] = [
    _sig(6, 658368, 384048, 1143000, 274320, sz=1100, b=1, clr='D9A441',
         font=THEME_FONT, text='04  PERCEPTION BARRIERS'),
] + [
    _sig(6, 749808, y, 310896, 310896, geom='ellipse', fill=_S6_DOT_FILL[i])
    for i, y in enumerate((1728216, 2889504, 4050792))
] + [
    _sig(6, 1225296, y, 1600200, 256032, sz=1500, b=1, clr='112A35',
         font=THEME_FONT, text=_S6_TITLES[i])
    for i, y in enumerate((1682496, 2843784, 4005072))
]

# ---- 第7页：突破之道（深色页）+ 金色竖条 + 四条编号条目 ----
_S7_TITLES = ['情感转译', '工业协同', '渠道前置', '类型创新']
PROTECTED['s7'] = [
    _sig(7, 676656, 932688, 2926080, 512064, sz=3000, b=1, clr='FFFFFF',
         font=THEME_FONT, text='突破之道'),
    _sig(7, 6181344, 0, 164592, 6858000, fill='D9A441'),
] + [
    _sig(7, 1810512, y, 1371600, 256032, sz=1420, b=1, clr='FFFFFF',
         font=THEME_FONT, text=_S7_TITLES[i])
    for i, y in enumerate((2359152, 3200400, 4041648, 4882896))
]

# ---- 第8页：资料依据 + 五条圆角条目（底色交替）+ 谢谢观看 ----
_S8_BAR_FILL = ['FFF9EF', 'F7F0E4', 'FFF9EF', 'F7F0E4', 'FFF9EF']
PROTECTED['s8'] = [
    _sig(8, 713232, 658368, 2377440, 457200, sz=2900, b=1, clr='112A35',
         font=THEME_FONT, text='资料依据'),
    _sig(8, 713232, 5742432, 2011680, 347472, sz=1900, b=1, clr='1E5B64',
         font=THEME_FONT, text='谢谢观看'),
] + [
    _sig(8, 932688, y, 9555480, 603504, geom='roundRect', fill=_S8_BAR_FILL[i])
    for i, y in enumerate((1874520, 2660904, 3447288, 4233672, 5020056))
]
# ==================== 基础工具：命名空间 / 文本 / 字体 ====================
def _qa(tag):
    return '{%s}%s' % (NS_A, tag)


def _qp(tag):
    return '{%s}%s' % (NS_P, tag)


def _qr(tag):
    return '{%s}%s' % (NS_R, tag)


def _norm_text(s):
    """归一化文本：连续空白折叠成单个空格并去除首尾空白。"""
    return re.sub(r'\s+', ' ', s or '').strip()


def _norm_font(name):
    """归一化字体名：处理不同办公软件保存后的同义名（微软雅黑 ≡ Microsoft YaHei）。"""
    if not name:
        return ''
    key = name.strip()
    if key in FONT_ALIAS:
        return FONT_ALIAS[key]
    return FONT_ALIAS.get(key.lower(), key)


def _hex6(v):
    if not v:
        return None
    v = v.strip().lstrip('#').upper()
    return v if re.match(r'^[0-9A-F]{6}$', v) else None


def _hls(hexv):
    r = int(hexv[0:2], 16) / 255.0
    g = int(hexv[2:4], 16) / 255.0
    b = int(hexv[4:6], 16) / 255.0
    return colorsys.rgb_to_hls(r, g, b)


def _to_hex(h, l, s):
    l = min(1.0, max(0.0, l))
    s = min(1.0, max(0.0, s))
    r, g, b = colorsys.hls_to_rgb(h, l, s)
    return '%02X%02X%02X' % (int(round(r * 255)), int(round(g * 255)),
                             int(round(b * 255)))

# ==================== 基础工具：颜色解析 ====================
def _apply_clr_mods(hexv, node):
    """把 a:lumMod / a:lumOff / a:shade / a:tint / a:satMod 作用到基色上。
    例：accent6(70AD47) + lumMod 75% -> 548235。"""
    h, l, s = _hls(hexv)
    for c in list(node):
        tag = c.tag.split('}')[-1]
        try:
            val = int(c.get('val')) / 100000.0
        except (TypeError, ValueError):
            continue
        if tag == 'lumMod':
            l *= val
        elif tag == 'lumOff':
            l += val
        elif tag == 'shade':
            l *= val
        elif tag == 'tint':
            l = l * val + (1.0 - val)
        elif tag == 'satMod':
            s *= val
    return _to_hex(h, l, s)


def _clr_of(node, ctx):
    """把 a:srgbClr / a:schemeClr / a:sysClr 解析为 6 位十六进制色值。"""
    if node is None:
        return None
    tag = node.tag.split('}')[-1]
    if tag == 'srgbClr':
        base = _hex6(node.get('val'))
    elif tag == 'sysClr':
        base = _hex6(node.get('lastClr')) or 'FFFFFF'
    elif tag == 'schemeClr':
        name = node.get('val') or ''
        name = ctx['clrmap'].get(name, name)
        base = ctx['theme'].get(name)
    else:
        return None
    if not base:
        return None
    return _apply_clr_mods(base, node)


def _solid_of(parent, ctx):
    """取 parent 下的填充色：a:noFill -> 'noFill'，渐变/图片/图案 -> 类型名，
    a:solidFill -> 色值，未指定 -> None。"""
    if parent is None:
        return None
    if parent.find(_qa('noFill')) is not None:
        return 'noFill'
    for tag in ('gradFill', 'blipFill', 'pattFill'):
        if parent.find(_qa(tag)) is not None:
            return tag
    sf = parent.find(_qa('solidFill'))
    if sf is None:
        return None
    for c in list(sf):
        v = _clr_of(c, ctx)
        if v:
            return v
    return None

def _is_dark_green(hexv):
    """深绿色系：色相落在黄绿-青绿之间、饱和度足够、明度偏低。"""
    if not hexv or not re.match(r'^[0-9A-F]{6}$', hexv):
        return False
    h, l, s = _hls(hexv)
    return 0.20 <= h <= 0.53 and s >= 0.15 and l <= 0.45


def _is_white(hexv):
    """白色或等效浅色1。"""
    if not hexv or not re.match(r'^[0-9A-F]{6}$', hexv):
        return False
    if hexv == 'FFFFFF':
        return True
    h, l, s = _hls(hexv)
    return l >= 0.92 and s <= 0.12


# ==================== 基础工具：主题 / 颜色映射 / 页部件 ====================
def _parse_theme(z, part):
    """返回 (配色方案 dict, 字体方案 dict)。"""
    scheme, fonts = {}, {}
    try:
        root = ET.fromstring(z.read(part))
    except Exception:
        return scheme, fonts
    cs = root.find('.//' + _qa('clrScheme'))
    if cs is not None:
        for c in list(cs):
            name = c.tag.split('}')[-1]
            srgb = c.find(_qa('srgbClr'))
            if srgb is not None:
                scheme[name] = _hex6(srgb.get('val'))
                continue
            sysc = c.find(_qa('sysClr'))
            if sysc is not None:
                scheme[name] = _hex6(sysc.get('lastClr'))
    fs = root.find('.//' + _qa('fontScheme'))
    if fs is not None:
        for kind in ('majorFont', 'minorFont'):
            f = fs.find(_qa(kind))
            if f is None:
                continue
            la, ea = f.find(_qa('latin')), f.find(_qa('ea'))
            fonts[kind] = (
                _norm_font(la.get('typeface') if la is not None else ''),
                _norm_font(ea.get('typeface') if ea is not None else ''))
    return scheme, fonts

def _rels(z, part):
    """读取某部件的 .rels，返回 {Id: 归一化后的部件名}。"""
    base = part.rsplit('/', 1)
    rel_name = (base[0] + '/_rels/' + base[1] + '.rels') if len(base) == 2 \
        else '_rels/' + part + '.rels'
    out = {}
    try:
        root = ET.fromstring(z.read(rel_name))
    except Exception:
        return out
    prefix = base[0] if len(base) == 2 else ''
    for r in root:
        tgt = (r.get('Target') or '').replace('\\', '/')
        if not tgt or tgt.startswith('http'):
            continue
        if tgt.startswith('/'):
            name = tgt.lstrip('/')
        else:
            cur = prefix
            while tgt.startswith('../'):
                tgt = tgt[3:]
                cur = cur.rsplit('/', 1)[0] if '/' in cur else ''
            name = (cur + '/' + tgt) if cur else tgt
        out[r.get('Id')] = name
    return out


def _slide_parts(z):
    """按 p:sldIdLst 的放映顺序返回幻灯片部件名列表。"""
    root = ET.fromstring(z.read('ppt/presentation.xml'))
    rmap = _rels(z, 'ppt/presentation.xml')
    parts = []
    lst = root.find(_qp('sldIdLst'))
    if lst is not None:
        for sid in lst.findall(_qp('sldId')):
            name = rmap.get(sid.get(_qr('id')))
            if name:
                parts.append(name)
    return parts


def _master_clrmap(z):
    """从 slideMaster1 读取 p:clrMap（bg1->lt1 等）。"""
    try:
        mst = ET.fromstring(z.read('ppt/slideMasters/slideMaster1.xml'))
    except Exception:
        return {}
    cm = mst.find(_qp('clrMap'))
    return dict(cm.attrib) if cm is not None else {}

# ==================== 基础工具：形状遍历（含组合坐标换算） ====================
SHAPE_KINDS = ('sp', 'pic', 'cxnSp', 'graphicFrame', 'grpSp')


def _shape_xfrm(el):
    for tag in ('spPr', 'grpSpPr'):
        pr = el.find(_qp(tag))
        if pr is not None:
            xf = pr.find(_qa('xfrm'))
            if xf is not None:
                return xf
    return el.find(_qp('xfrm'))


def _compose_tf(tf, grp):
    """把组合的子坐标系叠加到已有变换上；tf = (dx, dy, sx, sy)。"""
    xf = _shape_xfrm(grp)
    if xf is None:
        return tf
    off, ext = xf.find(_qa('off')), xf.find(_qa('ext'))
    ch_off, ch_ext = xf.find(_qa('chOff')), xf.find(_qa('chExt'))
    if off is None or ext is None or ch_off is None or ch_ext is None:
        return tf
    try:
        ox, oy = int(off.get('x')), int(off.get('y'))
        ecx, ecy = int(ext.get('cx')), int(ext.get('cy'))
        cox, coy = int(ch_off.get('x')), int(ch_off.get('y'))
        ccx, ccy = int(ch_ext.get('cx')), int(ch_ext.get('cy'))
    except (TypeError, ValueError):
        return tf
    gx = (ecx / float(ccx)) if ccx else 1.0
    gy = (ecy / float(ccy)) if ccy else 1.0
    sx, sy = gx * tf[2], gy * tf[3]
    return (tf[0] + ox * tf[2] - cox * sx, tf[1] + oy * tf[3] - coy * sy, sx, sy)


def _norm_geom(g):
    return 'line' if g in ('line', 'straightConnector1') else g

def _mk_shape(el, kind, tf, ctx, in_group, top):
    """把一个形状元素抽成便于比对的记录。坐标已换算为页面绝对坐标（EMU）。"""
    spPr = el.find(_qp('spPr'))
    xf = _shape_xfrm(el)
    rec = {'el': el, 'kind': kind, 'x': None, 'y': None, 'cx': None, 'cy': None,
           'rot': 0, 'in_group': in_group, 'top': top, 'is_label': False}
    if xf is not None:
        try:
            rec['rot'] = int(xf.get('rot') or 0)
        except ValueError:
            rec['rot'] = 0
        off, ext = xf.find(_qa('off')), xf.find(_qa('ext'))
        if off is not None:
            try:
                rec['x'] = tf[0] + int(off.get('x') or 0) * tf[2]
                rec['y'] = tf[1] + int(off.get('y') or 0) * tf[3]
            except ValueError:
                pass
        if ext is not None:
            try:
                rec['cx'] = int(ext.get('cx') or 0) * tf[2]
                rec['cy'] = int(ext.get('cy') or 0) * tf[3]
            except ValueError:
                pass
    rec['geom'] = None
    rec['adj'] = {}
    if spPr is not None:
        pg = spPr.find(_qa('prstGeom'))
        if pg is not None:
            rec['geom'] = _norm_geom(pg.get('prst'))
            for gd in pg.iter(_qa('gd')):
                m = re.search(r'val\s+(-?\d+)', gd.get('fmla') or '')
                if m:
                    rec['adj'][gd.get('name')] = int(m.group(1))
        elif spPr.find(_qa('custGeom')) is not None:
            rec['geom'] = 'custGeom'
    style = el.find(_qp('style'))
    rec['fill'] = _solid_of(spPr, ctx)
    if rec['fill'] is None and style is not None:
        fr = style.find(_qa('fillRef'))
        if fr is not None and (fr.get('idx') or '0') != '0':
            for c in list(fr):
                v = _clr_of(c, ctx)
                if v:
                    rec['fill'] = v
                    break
    _fill_line_and_effect(rec, el, spPr, style, ctx)
    _fill_text(rec, el, style, ctx)
    return rec

def _fill_line_and_effect(rec, el, spPr, style, ctx):
    """描边与效果：ln_w / ln_clr / ln_explicit_nofill / effects / eff_ref。"""
    rec['ln_w'] = None
    rec['ln_clr'] = None
    rec['ln_nofill'] = False
    ln = spPr.find(_qa('ln')) if spPr is not None else None
    rec['ln_present'] = ln is not None
    if ln is not None:
        if ln.get('w'):
            try:
                rec['ln_w'] = int(ln.get('w'))
            except ValueError:
                rec['ln_w'] = None
        rec['ln_clr'] = _solid_of(ln, ctx)
        rec['ln_nofill'] = rec['ln_clr'] == 'noFill'
    rec['effects'] = []
    if spPr is not None:
        lst = spPr.find(_qa('effectLst'))
        if lst is not None:
            rec['effects'] = [c.tag.split('}')[-1] for c in list(lst)]
        for tag in ('scene3d', 'sp3d'):
            if spPr.find(_qa(tag)) is not None:
                rec['effects'].append(tag)
    rec['eff_ref'] = None
    rec['ln_ref'] = None
    rec['font_ref_clr'] = None
    if style is not None:
        er = style.find(_qa('effectRef'))
        if er is not None:
            try:
                rec['eff_ref'] = int(er.get('idx') or 0)
            except ValueError:
                rec['eff_ref'] = None
        lr = style.find(_qa('lnRef'))
        if lr is not None:
            try:
                rec['ln_ref'] = int(lr.get('idx') or 0)
            except ValueError:
                rec['ln_ref'] = None
        fr = style.find(_qa('fontRef'))
        if fr is not None:
            for c in list(fr):
                v = _clr_of(c, ctx)
                if v and v != 'noFill':
                    rec['font_ref_clr'] = v
                    break

def _fill_text(rec, el, style, ctx):
    """文本层：runs（字号/字重/字体/字色）、段落对齐、bodyPr 参数、项目符号。"""
    rec.update({'text': '', 'runs': [], 'algn': None, 'anchor': None,
                'anchor_ctr': None, 'wrap': None, 'l_ins': None, 'r_ins': None,
                'autofit': None, 'bullet': False})
    body = el.find(_qp('txBody'))
    if body is None:
        return
    bodyPr = body.find(_qa('bodyPr'))
    if bodyPr is not None:
        rec['anchor'] = bodyPr.get('anchor')
        rec['anchor_ctr'] = bodyPr.get('anchorCtr')
        rec['wrap'] = bodyPr.get('wrap')
        rec['l_ins'] = bodyPr.get('lIns')
        rec['r_ins'] = bodyPr.get('rIns')
        na = bodyPr.find(_qa('normAutofit'))
        if na is not None:
            try:
                rec['autofit'] = int(na.get('fontScale') or 100000)
            except ValueError:
                rec['autofit'] = None
    for p in body.findall(_qa('p')):
        pPr = p.find(_qa('pPr'))
        if pPr is not None:
            if rec['algn'] is None:
                rec['algn'] = pPr.get('algn')
            if pPr.find(_qa('buChar')) is not None or \
                    pPr.find(_qa('buAutoNum')) is not None:
                rec['bullet'] = True
        for r in p.findall(_qa('r')):
            rPr = r.find(_qa('rPr'))
            t = r.find(_qa('t'))
            la = rPr.find(_qa('latin')) if rPr is not None else None
            ea = rPr.find(_qa('ea')) if rPr is not None else None
            clr = _solid_of(rPr, ctx)
            sz_raw = rPr.get('sz') if rPr is not None else None
            rec['runs'].append({
                'text': (t.text or '') if t is not None else '',
                'sz': int(sz_raw) if (sz_raw or '').isdigit() else None,
                'b': rPr is not None and rPr.get('b') == '1',
                'i': rPr is not None and rPr.get('i') == '1',
                'u': rPr is not None and (rPr.get('u') or 'none') != 'none',
                'latin': _norm_font(la.get('typeface') if la is not None else ''),
                'ea': _norm_font(ea.get('typeface') if ea is not None else ''),
                'clr': clr if clr not in ('noFill',) else None,
            })
    rec['text'] = _norm_text(''.join(t.text or '' for t in body.iter(_qa('t'))))

def _walk_tree(parent, ctx, tf, in_group, out, top):
    for c in list(parent):
        kind = c.tag.split('}')[-1]
        if kind not in SHAPE_KINDS:
            continue
        out.append(_mk_shape(c, kind, tf, ctx, in_group, top))
        if kind == 'grpSp':
            _walk_tree(c, ctx, _compose_tf(tf, c), True, out, False)


def _shapes_of(slide_root, ctx):
    tree = slide_root.find(_qp('cSld') + '/' + _qp('spTree'))
    out = []
    if tree is not None:
        _walk_tree(tree, ctx, (0, 0, 1.0, 1.0), False, out, True)
    return out


def _slide_bg(slide_root, ctx):
    """页面底色：p:sld/p:cSld/p:bg 下的 p:bgPr 或 p:bgRef。"""
    bg = slide_root.find(_qp('cSld') + '/' + _qp('bg'))
    if bg is None:
        return None
    bgPr = bg.find(_qp('bgPr'))
    if bgPr is not None:
        return _solid_of(bgPr, ctx)
    ref = bg.find(_qp('bgRef'))
    if ref is not None:
        for c in list(ref):
            v = _clr_of(c, ctx)
            if v:
                return v
    return None


def _eff_run_clr(rec, run, ctx):
    """字色的生效值：显式 run 填充 -> 形状 fontRef -> 主题正文色。"""
    if run.get('clr'):
        return run['clr']
    if rec.get('font_ref_clr'):
        return rec['font_ref_clr']
    name = ctx['clrmap'].get('tx1', 'dk1')
    return ctx['theme'].get(name) or '000000'


def _eff_run_font(run):
    """生效字体名集合：东亚字体与西文字体都算（中文看 a:ea，西文看 a:latin）。"""
    return set(f for f in (run.get('ea'), run.get('latin')) if f)

# ==================== 基础工具：定位 / 几何 / 签名比对 ====================
def _near(rec, x, y, cx, cy, tol):
    for got, want in ((rec['x'], x), (rec['y'], y), (rec['cx'], cx), (rec['cy'], cy)):
        if got is None or abs(got - want) > tol:
            return False
    return True


def _rect_overlap(r1, r2, tol=TOL_005):
    """矩形是否真正交叠；细线等零厚度元素在该方向上不给容差。"""
    x1, y1, w1, h1 = r1
    x2, y2, w2, h2 = r2
    tx = 0 if (w1 <= 0 or w2 <= 0) else tol
    ty = 0 if (h1 <= 0 or h2 <= 0) else tol
    iw = min(x1 + max(w1, 0), x2 + max(w2, 0)) - max(x1, x2)
    ih = min(y1 + max(h1, 0), y2 + max(h2, 0)) - max(y1, y2)
    return iw > tx and ih > ty


def _text_width_emu(text, sz):
    """单行文字宽度估算：中日韩全宽字符按 1.0 em，其余按 0.55 em。"""
    em = sz / 100.0 * EMU_PT
    total = 0.0
    for ch in text:
        total += em if ord(ch) > 0x2E7F else em * 0.55
    return total


def _real_runs(rec):
    return [r for r in rec['runs'] if (r.get('text') or '').strip()]


def _props_ok(rec, props, ctx):
    """逐项比对期望属性；只比对签名里写明的键。"""
    if 'kind' in props and rec['kind'] != props['kind']:
        return False
    if 'geom' in props and rec['geom'] != props['geom']:
        return False
    if 'fill' in props and rec['fill'] != props['fill']:
        return False
    if 'lnw' in props and rec['ln_w'] != props['lnw']:
        return False
    if 'lnclr' in props and rec['ln_clr'] != props['lnclr']:
        return False
    if 'algn' in props and rec['algn'] != props['algn']:
        return False
    if 'text' in props and rec['text'] != _norm_text(props['text']):
        return False
    runs = _real_runs(rec)
    if any(k in props for k in ('sz', 'b', 'clr', 'font')) and not runs:
        return False
    for run in runs:
        if 'sz' in props and run['sz'] != props['sz']:
            return False
        if 'b' in props and run['b'] != bool(props['b']):
            return False
        if 'clr' in props and _eff_run_clr(rec, run, ctx) != props['clr']:
            return False
        if 'font' in props and props['font'] not in _eff_run_font(run):
            return False
    return True


def _sig_ok(by_slide, sig, ctx):
    """签名命中：先按 (x, y, cx, cy) 定位，再比对属性。"""
    slide, x, y, cx, cy, props = sig
    for rec in by_slide.get(slide, []):
        if _near(rec, x, y, cx, cy, TOL_POS) and _props_ok(rec, props, ctx):
            return True
    return False


def _group_kept(by_slide, key, ctx):
    """PROTECTED 里某一组签名是否全部保持原样。"""
    return all(_sig_ok(by_slide, sig, ctx) for sig in PROTECTED[key])

def _theme_effects(z, part):
    """主题 effectStyleLst 各级是否真的带可见效果（本模板 1、2 级为空效果）。"""
    out = []
    try:
        root = ET.fromstring(z.read(part))
    except Exception:
        return out
    lst = root.find('.//' + _qa('effectStyleLst'))
    if lst is None:
        return out
    for es in list(lst):
        lst2 = es.find(_qa('effectLst'))
        visible = bool(lst2 is not None and len(list(lst2)))
        if es.find(_qa('scene3d')) is not None or es.find(_qa('sp3d')) is not None:
            visible = True
        out.append(visible)
    return out


def _theme_effect_visible(ctx, idx):
    effs = ctx.get('effects') or []
    if not idx or idx < 1 or idx > len(effs):
        return False
    return effs[idx - 1]


def _is_protected_pos(slide, rec):
    """该形状是否落在某条基准签名的位置上（用于识别既有元素）。"""
    for sigs in PROTECTED.values():
        for sig in sigs:
            if sig[0] == slide and _near(rec, sig[1], sig[2], sig[3], sig[4], TOL_POS):
                return True
    for pg, x, y, cx, cy, _t, _sz, _clr in PAGE_NUMBERS:
        if pg == slide and _near(rec, x, y, cx, cy, TOL_POS):
            return True
    return False


def _label_container(rec):
    """组合形状内是否只装了指令要求的小标签（用于统计新增元素时不误判）。"""
    if rec['kind'] != 'grpSp':
        return False
    texts = []
    for sp in rec['el'].iter(_qp('sp')):
        body = sp.find(_qp('txBody'))
        if body is None:
            continue
        t = _norm_text(''.join(x.text or '' for x in body.iter(_qa('t'))))
        if t:
            texts.append(t)
    return bool(texts) and all(t in LABEL_WORDS for t in texts)

# ==================== 对外接口 ====================
class _TargetMissing(Exception):
    """目录不存在或目录中没有 .pptx：连待评测对象都没有，不进入维度判定。"""


def evaluate(dir_path):
    """入参为脚本所在目录；在该目录内定位 .pptx 并返回评分结果 dict。"""
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
            raise _TargetMissing("目录不存在：%s" % dir_path)
        cands = sorted(f for f in os.listdir(dir_path)
                       if f.lower().endswith('.pptx') and not f.startswith('~$'))
        if not cands:
            raise _TargetMissing("目录中未找到 .pptx 文件：%s" % dir_path)
        file_path = os.path.join(dir_path, cands[0])
        result["file_name"] = os.path.basename(file_path)
        return _evaluate_impl(file_path, result)
    except _TargetMissing as exc:
        # 找不到待评测文件：维度1 未参与判定，dim1_pass / dim1_reason 一律留空，满分也置 0
        result["file_name"] = ""
        result["status"] = "error"
        result["error"] = str(exc)
        result["dim1_pass"] = ""
        result["dim1_reason"] = ""
        result["dim2_items"] = []
        result["total_score"] = 0
        result["max_score"] = 0
        return result
    except Exception as exc:
        # 文件存在但打不开 / 解析失败：维度1 判不通过，dim1_reason 给出错误摘要
        result["status"] = "error"
        result["error"] = "%s: %s" % (type(exc).__name__, exc)
        result["dim1_pass"] = False
        result["dim1_reason"] = "文件无法解析或无法打开：%s: %s" % (type(exc).__name__, exc)
        result["dim2_items"] = []
        result["total_score"] = 0
        result["max_score"] = MAX_SCORE_TOTAL
        return result


def _open_deck(file_path):
    """打开并解析演示文稿；返回 (zip, presentation 根, 幻灯片根列表)。"""
    z = zipfile.ZipFile(file_path)
    broken = z.testzip()
    if broken:
        raise zipfile.BadZipFile("包内文件损坏：%s" % broken)
    pres = ET.fromstring(z.read('ppt/presentation.xml'))
    parts = _slide_parts(z)
    if not parts:
        raise ValueError("演示文稿内没有可用的幻灯片")
    roots = [ET.fromstring(z.read(p)) for p in parts]
    return z, pres, roots

def _build_ctx(z):
    """组装配色 / 字体 / 效果解析上下文。"""
    theme_part = 'ppt/theme/theme1.xml'
    for name in _rels(z, 'ppt/slideMasters/slideMaster1.xml').values():
        if name.startswith('ppt/theme/') and name.endswith('.xml'):
            theme_part = name
            break
    theme, fonts = _parse_theme(z, theme_part)
    if not theme:
        theme = dict(THEME_SCHEME)
    clrmap = _master_clrmap(z) or {
        'bg1': 'lt1', 'tx1': 'dk1', 'bg2': 'lt2', 'tx2': 'dk2'}
    return {'theme': theme, 'fonts': fonts, 'clrmap': clrmap,
            'effects': _theme_effects(z, theme_part), 'theme_part': theme_part}


def _collect_labels(by_slide, ctx):
    """在目录页找出指令要求的小标签，并按四组归位。"""
    labels = []
    for rec in by_slide.get(LABEL_SLIDE, []):
        if rec['kind'] == 'grpSp' or rec['x'] is None:
            continue
        if rec['text'] in LABEL_WORDS:
            rec['is_label'] = True
            labels.append(rec)
    heads = []
    for i, (hx, hy) in enumerate(HEADING_XY):
        found = None
        for rec in by_slide.get(LABEL_SLIDE, []):
            if rec['is_label'] or rec['x'] is None:
                continue
            if rec['text'] == LABEL_GROUPS[i][0]:
                found = rec
                break
        if found is not None and found['cy']:
            heads.append((found['x'], found['y'], found['cx'], found['cy']))
        else:
            heads.append((hx, hy, HEADING_EXT[0], HEADING_EXT[1]))
    groups = dict((i, []) for i in range(4))
    for rec in labels:
        idx = -1
        for i, head in enumerate(heads):
            if rec['y'] >= head[1]:
                idx = i
        if idx >= 0:
            groups[idx].append(rec)
    for i in range(4):
        groups[i].sort(key=lambda r: r['x'])
    return labels, heads, groups

# ==================== 维度二：加分项 01-15 ====================
def _check_bonus(results, by_slide, labels, heads, groups, ctx):
    n12 = len(labels) == 12

    # 01 12个标签齐全、分组对应、无错漏、全在目录页
    ok = n12 and sorted(r['text'] for r in labels) == sorted(LABEL_WORDS)
    if ok:
        for i in range(4):
            grp = groups[i]
            if len(grp) != 3 or sorted(r['text'] for r in grp) != \
                    sorted(LABEL_GROUPS[i][1]):
                ok = False
                break
    if ok:
        for si, recs in by_slide.items():
            if si == LABEL_SLIDE:
                continue
            for rec in recs:
                if rec['text'] in LABEL_WORDS and rec['x'] is not None and \
                        not _is_protected_pos(si, rec):
                    ok = False
                    break
    results.append(("+5：" if ok else "+0：") +
                   "目录页四组小标题下方共新增12个小标签（每组3个），文字与所属组严格对应、"
                   "无错字无漏项无多余标签，且全部位于目录页")

    # 02 中文实际生效 MiSans（a:ea 与 a:latin 同时指定）
    ok = n12
    for rec in labels:
        runs = _real_runs(rec)
        if not runs:
            ok = False
        for run in runs:
            if run['ea'] != LABEL_FONT or run['latin'] != LABEL_FONT:
                ok = False
    results.append(("+5：" if ok else "+0：") +
                   "标签中文字符实际生效为 MiSans——东亚字体（a:ea）与西文字体（a:latin）"
                   "同时指定为 MiSans")

    # 03 标签为可单独点选编辑的形状
    ok = n12 and all(rec['kind'] == 'sp' and not rec['in_group'] for rec in labels)
    results.append(("+3：" if ok else "+0：") +
                   "12个小标签均为可单独点选、单独编辑的形状，未以图片形式贴入、"
                   "未被合并进无法单独拆分的组合")

    # 04 字号统一 10.5 磅
    ok = n12
    for rec in labels:
        if rec['autofit'] not in (None, 100000):
            ok = False
        for run in _real_runs(rec):
            if run['sz'] != LABEL_SZ:
                ok = False
    results.append(("+3：" if ok else "+0：") +
                   "12个标签文字字号统一为10.5磅，未出现放大、缩小或自动缩放导致的字号不一致")

    # 05 底色深绿且12个统一
    fills = set(rec['fill'] for rec in labels)
    ok = n12 and len(fills) == 1 and _is_dark_green(list(fills)[0])
    results.append(("+3：" if ok else "+0：") +
                   "标签底色为深绿色系且12个完全统一")
    # 06 文字白色
    ok = n12
    for rec in labels:
        runs = _real_runs(rec)
        if not runs:
            ok = False
        for run in runs:
            if not _is_white(_eff_run_clr(rec, run, ctx)):
                ok = False
    results.append(("+3：" if ok else "+0：") +
                   "标签文字为白色（#FFFFFF 或等效浅色1），与深绿底形成清晰对比")

    # 07 圆角胶囊形（adj≥40%）
    ok = n12 and all(rec['geom'] == 'roundRect' and
                     rec['adj'].get('adj', 0) >= LABEL_ADJ_MIN for rec in labels)
    results.append(("+3：" if ok else "+0：") +
                   "标签为圆角胶囊形（圆角矩形且圆角调节值≥40%），不是直角矩形、椭圆或其他形状")

    # 08 同组自左至右、顶边齐平
    ok = n12
    for i in range(4):
        grp = groups[i]
        if len(grp) != 3:
            ok = False
            continue
        if [r['text'] for r in grp] != list(LABEL_GROUPS[i][1]):
            ok = False
        ys = [r['y'] for r in grp]
        if max(ys) - min(ys) > TOL_005:
            ok = False
    results.append(("+3：" if ok else "+0：") +
                   "同组三个标签自左至右的次序与指令词序完全一致，且三者顶边纵坐标一致（≤0.05cm）")

    # 09 等距 + 同列对齐 + 行距与小标题一致（规整矩阵）
    ok = n12
    cols = []
    for i in range(4):
        grp = groups[i]
        if len(grp) != 3:
            ok = False
            continue
        xs = [r['x'] for r in grp]
        if abs((xs[1] - xs[0]) - (xs[2] - xs[1])) > TOL_005:
            ok = False
        cols.append(xs)
    if ok and len(cols) == 4:
        for j in range(3):
            vals = [c[j] for c in cols]
            if max(vals) - min(vals) > TOL_005:
                ok = False
        rows = [min(r['y'] for r in groups[i]) for i in range(4)]
        row_gaps = [rows[i + 1] - rows[i] for i in range(3)]
        head_gaps = [heads[i + 1][1] - heads[i][1] for i in range(3)]
        if max(row_gaps) - min(row_gaps) > TOL_005:
            ok = False
        for got, want in zip(row_gaps, head_gaps):
            if abs(got - want) > TOL_015:
                ok = False
    else:
        ok = False
    results.append(("+3：" if ok else "+0：") +
                   "四行三列对齐成规整矩阵——同组水平间距相等、四组共用同一套横向坐标、"
                   "相邻两行标签顶边间距与相邻两组小标题顶边间距一致")
    # 10 12个尺寸一致
    ok = n12
    if ok:
        cxs = [r['cx'] or 0 for r in labels]
        cys = [r['cy'] or 0 for r in labels]
        ok = (max(cxs) - min(cxs) <= TOL_002) and (max(cys) - min(cys) <= TOL_002)
    results.append(("+3：" if ok else "+0：") +
                   "12个标签尺寸完全一致（同一宽度、同一高度，偏差≤0.02cm）")

    # 11 落在该组小标题下方空白区
    ok = n12
    for i in range(4):
        grp = groups[i]
        if len(grp) != 3:
            ok = False
            continue
        top = min(r['y'] for r in grp)
        bottom = max(r['y'] + (r['cy'] or 0) for r in grp)
        gap = top - (heads[i][1] + heads[i][3])
        if not (BAND_GAP_MIN <= gap <= BAND_GAP_MAX):
            ok = False
        limit = heads[i + 1][1] if i < 3 else HAIRLINE_Y
        if bottom > limit:
            ok = False
    results.append(("+5：" if ok else "+0：") +
                   "每组三个标签整体落在该组小标题下方空白区——上边缘位于小标题底边下方"
                   "0.35–2.0cm，下边缘不越过下一组小标题顶边（第四组不越过页脚细分隔线）")

    # 12 与既有元素互不重叠
    others = [r for r in by_slide.get(LABEL_SLIDE, [])
              if not r['is_label'] and r['kind'] != 'grpSp' and r['x'] is not None]
    ok = n12
    for rec in labels:
        r1 = (rec['x'], rec['y'], rec['cx'] or 0, rec['cy'] or 0)
        for other in others:
            r2 = (other['x'], other['y'], other['cx'] or 0, other['cy'] or 0)
            if _rect_overlap(r1, r2):
                ok = False
                break
    results.append(("+3：" if ok else "+0：") +
                   "标签块与目录页既有元素互不重叠——未覆盖01–04序号、金色短横线、四个小标题、"
                   "右侧说明文字、竖向金色分隔线、右侧配图及右下角总结句")

    # 13 色块尺寸与文字匹配、文字居中
    ok = n12
    for rec in labels:
        cy = rec['cy'] or 0
        if not (LABEL_CY_MIN <= cy <= LABEL_CY_MAX):
            ok = False
        if rec['algn'] != 'ctr' or (rec['anchor'] or '') != 'ctr':
            ok = False
        runs = _real_runs(rec)
        sz = (runs[0]['sz'] if runs and runs[0]['sz'] else LABEL_SZ)
        l_ins = int(rec['l_ins']) if (rec['l_ins'] or '').isdigit() else 91440
        r_ins = int(rec['r_ins']) if (rec['r_ins'] or '').isdigit() else 91440
        if (rec['cx'] or 0) - l_ins - r_ins < _text_width_emu(rec['text'], sz):
            ok = False
    results.append(("+1：" if ok else "+0：") +
                   "色块尺寸与10.5磅文字匹配——最长的“产业链断点”不截断不换行不溢出，"
                   "标签高度在0.55–0.90cm之间，文字在色块内上下左右居中")
    # 14 扁平纯色块（无描边、无效果）
    ok = n12
    for rec in labels:
        if rec['fill'] in ('noFill', 'gradFill', 'blipFill', 'pattFill'):
            ok = False
        if not rec['ln_nofill']:
            if rec['ln_clr'] and rec['ln_clr'] != rec['fill']:
                ok = False
            elif rec['ln_clr'] is None and (rec['ln_ref'] or 0) != 0:
                ok = False
        if rec['effects']:
            ok = False
        if _theme_effect_visible(ctx, rec['eff_ref']):
            ok = False
    results.append(("+1：" if ok else "+0：") +
                   "标签为扁平纯色块——无可见描边，且未附加阴影、发光、三维、渐变等效果")

    # 15 常规字重纯文本
    ok = n12
    for rec in labels:
        if rec['bullet']:
            ok = False
        raw = ''.join(r['text'] for r in rec['runs'])
        if raw.strip() != rec['text'] or rec['text'] not in LABEL_WORDS:
            ok = False
        for run in _real_runs(rec):
            if run['b'] or run['i'] or run['u']:
                ok = False
    results.append(("+1：" if ok else "+0：") +
                   "标签文字为常规字重的纯文本——未加粗未倾斜未加下划线，"
                   "无项目符号、编号、多余空格或标点")

# ==================== 维度二：减分项 16-35 ====================
def _check_deduct(results, by_slide, roots, pres, labels, ctx):
    def kept(key):
        return _group_kept(by_slide, key, ctx)

    # 16 目录页页首 CONTENTS / 目录
    bad = not kept('s2_header')
    results.append(("-1：" if bad else "-0：") +
                   "第2页页首“CONTENTS”（Microsoft YaHei、11磅、加粗、#C77A42）与“目录”"
                   "（31磅、加粗、#112A35）被改写、改变字体字号字色或移动位置")

    # 17 目录页 01-04 序号
    bad = not kept('s2_numbers')
    results.append(("-1：" if bad else "-0：") +
                   "第2页左侧“01”“02”“03”“04”四个序号不满足 Microsoft YaHei、20磅、加粗、"
                   "#D9A441，或四者纵向间距被改动")

    # 18 目录页金色短横线与竖向分隔线
    bad = not kept('s2_goldlines')
    results.append(("-1：" if bad else "-0：") +
                   "第2页四条金色短横线（线宽1.5磅）与右侧竖向金色分隔线（线宽2磅）不满足"
                   "#D9A441，或出现缺失、变长、位置偏移")

    # 19 目录页四个小标题
    bad = not kept('s2_headings')
    results.append(("-3：" if bad else "-0：") +
                   "第2页四个小标题“文化沟通”“工业体系”“全球竞争”“认知偏差”被改写，"
                   "或不满足 Microsoft YaHei、17磅、加粗、#112A35，或位置被移动")

    # 20 目录页四条说明文字
    bad = not kept('s2_descs')
    results.append(("-3：" if bad else "-0：") +
                   "第2页四条说明文字被删除、改写，或不满足 Microsoft YaHei、12磅、#65737B")

    # 21 目录页右侧配图
    bad = not kept('s2_pic')
    results.append(("-3：" if bad else "-0：") +
                   "第2页右侧配图（约13.16cm×7.40cm）被替换、缩放、移动、裁剪、加边框或删除")

    # 22 目录页右下角总结句
    bad = not kept('s2_summary')
    results.append(("-3：" if bad else "-0：") +
                   "第2页右下角总结句被改写，或不满足 Microsoft YaHei、16磅、加粗、#1E5B64")

    # 23 目录页页脚细线与页码 02
    bad = not kept('s2_footer')
    results.append(("-1：" if bad else "-0：") +
                   "第2页页脚被改动——底部横向细分隔线（线宽0.6磅、#D6C9B8）被删除变色移动，"
                   "或右下角页码“02”不满足9.5磅、#65737B、右对齐")
    # 24 添加了指令未要求的元素
    bad = False
    for idx in range(1, len(SHAPE_COUNTS) + 1):
        tops = [r for r in by_slide.get(idx, []) if r['top']]
        expect = SHAPE_COUNTS[idx - 1]
        if idx == LABEL_SLIDE:
            expect += len([r for r in tops if r['is_label'] or _label_container(r)])
        if len(tops) != expect:
            bad = True
            break
    if len(roots) > len(SHAPE_COUNTS):
        bad = True
    results.append(("-3：" if bad else "-0：") +
                   "添加了指令未要求的元素——目录页出现额外说明文字、图标、装饰线条、背景色块、"
                   "多余标签，或在其他页面新增小标签、色块、重复元素")

    # 25 第1页封面元素
    bad = not kept('s1')
    results.append(("-3：" if bad else "-0：") +
                   "第1页封面英文行、主标题“中国电影业面临的挑战”（34磅、加粗、#112A35）、"
                   "副标题行（14.5磅、#1E5B64）、“行业观察稿”、右侧通高金色竖条（#D9A441）"
                   "或右侧配图被改写、移动、缩放、删除")

    # 26 第3页标题与四张卡片
    bad = not kept('s3')
    results.append(("-3：" if bad else "-0：") +
                   "第3页“01 CULTURAL COMMUNICATION”标题及四张底色 #FFF9EF 的圆角卡片"
                   "（约8.51cm×4.44cm）的文字、底色、尺寸或位置被改动")

    # 27 第4页标题与四张卡片
    bad = not kept('s4')
    results.append(("-3：" if bad else "-0：") +
                   "第4页“02 INDUSTRIAL CAPABILITY”标题及四张底色 #FFF5E8 的圆角卡片"
                   "（约7.49cm×3.56cm）与“核心判断”整句的文字、底色、尺寸或位置被改动")

    # 28 第5页标题与三张卡片
    bad = not kept('s5')
    results.append(("-3：" if bad else "-0：") +
                   "第5页“03 GLOBAL COMPETITION”标题及三张底色 #FFF9EF 的圆角卡片与页内配图的"
                   "文字、底色、尺寸或位置被改动")

    # 29 第6页标题与三个圆点
    bad = not kept('s6')
    results.append(("-3：" if bad else "-0：") +
                   "第6页“04 PERCEPTION BARRIERS”标题、三个圆点（直径约0.86cm，"
                   "#D9A441／#1E5B64／#C85C4A）及“预设立场”“文化陌生”“双重审视”被改动、变色或删除")

    # 30 第7页深色页元素
    bad = not kept('s7')
    results.append(("-3：" if bad else "-0：") +
                   "第7页“突破之道”白色标题（30磅）、左侧金色竖条（#D9A441）、"
                   "四条编号条目（“情感转译”“工业协同”“渠道前置”“类型创新”）被改动")

    # 31 第8页资料条目
    bad = not kept('s8')
    results.append(("-3：" if bad else "-0：") +
                   "第8页“资料依据”（29磅）标题、五条底色交替为 #FFF9EF／#F7F0E4 的圆角条目"
                   "（约26.51cm×1.68cm）、“谢谢观看”（19磅、#1E5B64）被改写、变色、删除")
    # 32 页数/页序/页码标注
    bad = len(roots) != len(SHAPE_COUNTS)
    if not bad:
        for pg, x, y, cx, cy, txt, sz, clr in PAGE_NUMBERS:
            sig = _sig(pg, x, y, cx, cy, text=txt, sz=sz, clr=clr)
            if not _sig_ok(by_slide, sig, ctx):
                bad = True
                break
    results.append(("-3：" if bad else "-0：") +
                   "文稿页数不再是8页，或页序被调换，或各页右下角页码标注（01–08）"
                   "与实际页序不一致")

    # 33 页面底色与页面尺寸
    bad = False
    for idx, want in enumerate(BG_COLORS, 1):
        got = _slide_bg(roots[idx - 1], ctx) if idx <= len(roots) else None
        if got != want:
            bad = True
            break
    sld_sz = pres.find(_qp('sldSz'))
    try:
        sz_cx = int(sld_sz.get('cx'))
        sz_cy = int(sld_sz.get('cy'))
    except (AttributeError, TypeError, ValueError):
        sz_cx = sz_cy = 0
    if not (SLD_CX_CM[0] <= sz_cx / EMU_CM <= SLD_CX_CM[1] and
            SLD_CY_CM[0] <= sz_cy / EMU_CM <= SLD_CY_CM[1]):
        bad = True
    if not sz_cy or abs(sz_cx / float(sz_cy) - 16.0 / 9.0) > 0.02:
        bad = True
    results.append(("-1：" if bad else "-0：") +
                   "整体版式参数被改动——任一页面底色不再依次为 #F4EFE6、#E9F0EC、#F4EFE6、"
                   "#F1E8DA、#EAF1F2、#F5EEE8、#112A35、#EFE6D7，或页面尺寸不再是16:9宽屏")

    # 34 字体主题 / 配色主题
    bad = False
    for name, want in THEME_SCHEME.items():
        if (ctx['theme'].get(name) or '') != want:
            bad = True
            break
    for kind in ('majorFont', 'minorFont'):
        pair = ctx['fonts'].get(kind)
        if not pair or THEME_FONT not in pair:
            bad = True
            break
    results.append(("-1：" if bad else "-0：") +
                   "整份文稿的字体主题、配色主题被整体替换，导致既有文字字体或色彩连带发生变化")

    # 35 出现乱码
    bad = False
    for recs in by_slide.values():
        for rec in recs:
            text = rec['text']
            if not text:
                continue
            if any(m in text for m in ('�', '□', 'ï¿½')):
                bad = True
            if re.search('[À-ÿ][-¿]', text):
                bad = True
    results.append(("-3：" if bad else "-0：") +
                   "文稿中出现乱码——既有中文标题、说明文字或新增标签文字显示为乱码、方框、问号")

# ==================== 主评估流程 ====================
def _evaluate_impl(file_path, result):
    # ---- 维度一：交付文件能否作为 .pptx 正常打开 ----
    try:
        z, pres, roots = _open_deck(file_path)
    except Exception as exc:
        result["status"] = "error"
        result["error"] = "%s: %s" % (type(exc).__name__, exc)
        result["dim1_pass"] = False
        result["dim1_reason"] = "交付文件无法作为 PowerPoint 演示文稿正常打开（%s: %s），" \
                                "完成度总分记0" % (type(exc).__name__, exc)
        result["dim2_items"] = []
        result["total_score"] = 0
        result["max_score"] = MAX_SCORE_TOTAL
        return result
    result["dim1_pass"] = True
    result["dim1_reason"] = "交付文件为1个可正常打开的 PowerPoint 演示文稿（.pptx，共%d页）" \
                            % len(roots)

    # ---- 维度二：完成度 ----
    ctx = _build_ctx(z)
    by_slide = {}
    for idx, root in enumerate(roots, 1):
        by_slide[idx] = _shapes_of(root, ctx)
    labels, heads, groups = _collect_labels(by_slide, ctx)

    results = []
    _check_bonus(results, by_slide, labels, heads, groups, ctx)
    _check_deduct(results, by_slide, roots, pres, labels, ctx)
    z.close()

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
        dim2_items.append({"rule": rule, "max_delta": max_delta,
                           "delta": max_delta if hit else 0,
                           "hit": hit})
        total_score += (max_delta if hit else 0)
    result["dim2_items"] = dim2_items
    result["total_score"] = max(0, min(MAX_SCORE_TOTAL, total_score))
    result["max_score"] = MAX_SCORE_TOTAL
    return result


if __name__ == "__main__":
    if len(sys.argv) >= 2:
        _dir = sys.argv[1]
    else:
        _dir = os.path.dirname(os.path.abspath(__file__))
    _out = evaluate(_dir)
    print(json.dumps(_out, ensure_ascii=False, indent=2))
