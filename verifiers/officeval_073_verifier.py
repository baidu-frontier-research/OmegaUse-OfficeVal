#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# pyright: reportUnknownMemberType=false, reportUnknownVariableType=false, reportUnknownArgumentType=false, reportUnknownParameterType=false, reportMissingParameterType=false, reportPrivateUsage=false, reportUnknownLambdaType=false, reportGeneralTypeIssues=false

import os
import sys
import json
from pptx import Presentation
from pptx.util import Emu

# 说明：按接口统一约定，本模块不修改全局 sys.stdout，也不 print 主结果。
# 主结果统一由 evaluate(dir_path) -> dict 返回；入参为“脚本所在目录的路径”，
# 由脚本自己在该目录内定位并打开被评估的 .pptx。本文件仅在 __main__ 下作为
# 本地调试入口，才把 dict 序列化为 JSON 打印。

SCRIPT_ID = "073"
# 维度二加分项 max_delta 之和；用于维度一未通过 / 脚本异常时兜底
MAX_SCORE_TOTAL = 27
# 维度二减分项 max_delta 之和（负值）；加分项与减分项合计计算，合计为负数时
# 按实际负分记，不归零。
MIN_SCORE_TOTAL = -61

# 维度二加分项按“结果追加顺序”的 max_delta 列表；共 7 项，累加 = MAX_SCORE_TOTAL。
# 加分项只覆盖“指令要求动手做的活”（改字号、改字体为宋体、收窄至 8cm 以内、
# 收窄后重新对齐）；未做任何修改的交付不触发任何加分项，合计为 0 分。
_MAX_DELTAS_ORDER = [
    5,  # A01 第1页全部文字 6 号字
    5,  # A02 第2页全部文字 6 号字
    3,  # A03 第1页三处楷体文字改为宋体
    5,  # A04 第1页长条圆角矩形收窄至 8cm 内
    3,  # A05 第2页顶部圆角矩形收窄至 8cm 内
    3,  # A06 第1页三处条目形状收窄至 8cm 内
    3,  # A07 收窄后仍居中、未越界、未交叠（无形状收窄则不得分）
]
# 维度二减分项按“结果追加顺序”的 max_delta 列表（负值）；共 27 项，累加 = MIN_SCORE_TOTAL。
# 分值按“恢复被破坏内容所需的工作量”给定：1 = 三个以内形状或一处全局属性（5 分钟内）；
# 3 = 四个以上形状成组错乱或整套填充线条（5-10 分钟）；5 = 整页版式、15 个形状、
# 全部文字重排或文字丢失（大于 10 分钟）。
_MIN_DELTAS_ORDER = [
    -5,  # D01 两页图形整体覆盖范围
    -1,  # D02 第1页三个层级标签圆角矩形
    -1,  # D03 第1页顶部三个矩形
    -3,  # D04 第1页金字塔四个等腰三角形
    -3,  # D05 第1页六个素养矩形
    -3,  # D06 第1页五处文本框
    -1,  # D07 第1页视角/支撑/理论三个文本框
    -1,  # D08 第1页左右两个内容底框矩形
    -1,  # D09 第1页底部虚线边框矩形
    -3,  # D10 第1页各辅助图元
    -1,  # D11 第2页三个竖排层级圆角矩形
    -1,  # D12 第2页联动贯通/持续推进两矩形
    -3,  # D13 第2页五个措施圆角矩形
    -3,  # D14 第2页五个环节圆角矩形
    -5,  # D15 第2页15个竖排圆角矩形
    -3,  # D16 第2页五个成效圆角矩形
    -1,  # D17 第2页三个大边框矩形
    -3,  # D18 第2页各辅助图元
    -3,  # D19 各形状填充与线条设置
    -1,  # D20 第1页其余文字中文（东亚）字体非宋体
    -1,  # D21 第2页文字中文（东亚）字体非宋体
    -1,  # D22 数字与西文字体非宋体
    -5,  # D23 文字在 6 磅下溢出/被截断/线宽不相称
    -5,  # D24 123 个图形对象与全部文字未完整保留
    -1,  # D25 页数/页序/页面尺寸/备注页
    -1,  # D26 底部虚线框内出现多余字符
    -1,  # D27 四处条目文字被加粗
]
_DELTAS_ORDER = _MAX_DELTAS_ORDER + _MIN_DELTAS_ORDER
assert sum(_MAX_DELTAS_ORDER) == MAX_SCORE_TOTAL
assert sum(_MIN_DELTAS_ORDER) == MIN_SCORE_TOTAL

_A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
_P = 'http://schemas.openxmlformats.org/presentationml/2006/main'

# 严格“宋体”判定（仅认这几种字体名，含其英文别名）
STRICT_SONGTI = ('宋体', 'SimSun', 'NSimSun')
# 目标字号：6 号字 = 6 磅 = sz 600（百分之一磅）
TARGET_SZ = 600
# 几何量比对容差（cm）：形状定位/尺寸
TOL_POS = 0.06
# 辅助图元跨度比对容差（cm）：细则中以“约”描述的斜向箭头、连接符长度
TOL_SPAN = 0.09
# 整页图形覆盖范围比对容差（cm）
TOL_BBOX = 0.5
# 收窄后水平居中比对容差（cm）
TOL_CENTER = 0.25
# 宽度不超过 8cm 的判定上限（含 0.02cm 余量，容许换算舍入）
LIMIT_8CM = 8.02
# 页面尺寸（cm）：16:9 宽屏
PAGE_W, PAGE_H = 33.867, 19.05
# 每页顶层图形对象数量
SHAPE_COUNT = {1: 45, 2: 78}
# 6 磅字号下单个全角字的字面尺寸（cm）与行距系数
EM_6PT = 6.0 * 2.54 / 72.0
LINE_FACTOR = 1.2
# 文本框默认左右（上下）内边距合计（cm）：91440 EMU * 2
INSET_SUM = 0.254

# 第1页“1.…2.…”四处条目文字所在形状：既是收窄目标，也是禁止加粗的对象
ITEM_SHAPES = [
    (1, '矩形 51'),
    (1, '文本框 81'),
    (1, '文本框 84'),
    (1, '圆角矩形 37'),
]
# 第1页由楷体改为宋体的三处文字
KAITI_SHAPES = [
    (1, '文本框 63'),
    (1, '文本框 6'),
    (1, '文本框 74'),
]
# 应当保持无文字的底部虚线边框矩形
EMPTY_DASH_BOX = (1, '矩形 52')

# 交付文件应完整保留的全部文字：{(页码, 形状名): 文字}，共 62 处
# 其中空格数量、序号“1.”“2.”、逗号/分号/句号等标点均须逐字一致。
EXPECTED_TEXT = {
    (1, '矩形 9'): '价值引领根本目标',
    (1, '矩形 10'): '学习者整体成长',
    (1, '矩形 51'): '1.儿童认知发展理论     2.自主学习支持理论',
    (1, '文本框 63'): '教育学\n视角',
    (1, '文本框 74'): '心理发展理论',
    (1, '文本框 75'): '立规范',
    (1, '文本框 81'): '1.差异化教学思想        2.支架式成长理论',
    (1, '文本框 84'): '1.智能协作理论       2.生成技术应用理论',
    (1, '文本框 3'): '定目标',
    (1, '文本框 4'): '划边界',
    (1, '文本框 6'): '技术学\n支撑',
    (1, '矩形 7'): '价值涵养',
    (1, '矩形 8'): '智慧启发',
    (1, '矩形 13'): '品格践行',
    (1, '文本框 14'): '育德',
    (1, '矩形 24'): '知识奠基',
    (1, '矩形 26'): '能力提升',
    (1, '矩形 27'): '素养生成',
    (1, '文本框 30'): '成才',
    (1, '圆角矩形 37'): '1.儿童成长规律与学习特征   2.课程标准素养目标与育人要求',
    (1, '矩形 1'): '个性学习落地',
    (1, '圆角矩形 15'): '目标层',
    (1, '圆角矩形 23'): '核心层',
    (1, '圆角矩形 29'): '基础层',
    (2, '矩形 11'): '联动贯通',
    (2, '圆角矩形 21'): '智能技术支持小学生差异化学习',
    (2, '矩形 23'): '持续推进',
    (2, '圆角矩形 36'): '机制层',
    (2, '圆角矩形 44'): '教师统筹审核',
    (2, '圆角矩形 48'): '持续复盘改进',
    (2, '圆角矩形 50'): '发展规律匹配',
    (2, '圆角矩形 51'): '家校联动支持',
    (2, '圆角矩形 52'): '智能工具协助',
    (2, '圆角矩形 64'): '学情识别',
    (2, '圆角矩形 65'): '内容匹配',
    (2, '圆角矩形 66'): '路径调整',
    (2, '圆角矩形 67'): '反馈评价',
    (2, '圆角矩形 74'): '信息汇集',
    (2, '圆角矩形 75'): '综合分析',
    (2, '圆角矩形 76'): '教师确认',
    (2, '圆角矩形 77'): '特征匹配',
    (2, '圆角矩形 78'): '资源生成',
    (2, '圆角矩形 79'): '教师审看',
    (2, '圆角矩形 80'): '层级安排',
    (2, '圆角矩形 81'): '及时调整',
    (2, '圆角矩形 82'): '人工修订',
    (2, '圆角矩形 83'): '多维评价',
    (2, '圆角矩形 84'): '鼓励反馈',
    (2, '圆角矩形 85'): '家校互通',
    (2, '圆角矩形 96'): '准确把握差异，\n降低判断偏差；',
    (2, '圆角矩形 98'): '形成匹配资源，\n守住内容规范；',
    (2, '圆角矩形 100'): '对接成长区间，\n促进逐级提升；',
    (2, '圆角矩形 102'): '汇聚支持力量，\n维护学习信心；',
    (2, '圆角矩形 126'): '实施层',
    (2, '圆角矩形 127'): '目标层',
    (2, '文本框 130'): '规范内容',
    (2, '圆角矩形 1'): '循环改进',
    (2, '圆角矩形 2'): '状态更新',
    (2, '圆角矩形 3'): '策略调整',
    (2, '圆角矩形 4'): '规则完善',
    (2, '圆角矩形 8'): '不断优化流程，\n回应成长需要。',
    (2, '文本框 26'): '运行逻辑',
}
assert len(EXPECTED_TEXT) == 62

# ===== 各减分项的形状几何规格：形状名 -> (左, 上, 宽, 高)，None 表示该项不比对 =====
# D02 第1页“目标层”“核心层”“基础层”三个层级标签圆角矩形
_SPEC_D02 = {
    '圆角矩形 15': (3.34, 1.94, 3.53, 1.28),
    '圆角矩形 23': (3.34, 9.43, 3.53, 1.28),
    '圆角矩形 29': (3.34, 15.14, 3.53, 1.28),
}
# D03 第1页“价值引领根本目标”“个性学习落地”“学习者整体成长”三个矩形
_SPEC_D03 = {
    '矩形 9': (8.02, 1.94, 5.34, 1.28),
    '矩形 1': (16.71, 2.00, 5.34, 1.28),
    '矩形 10': (25.83, 1.97, 4.71, 1.28),
}
# D04 第1页金字塔外层与内部三个小等腰三角形
_SPEC_D04 = {
    '等腰三角形 47': (13.55, 4.77, 11.60, 9.40),
    '等腰三角形 48': (13.55, None, 4.80, 3.90),
    '等腰三角形 49': (20.33, None, 4.80, 3.90),
    '等腰三角形 50': (16.91, None, 4.80, 3.90),
}
# D05 第1页“价值涵养”等六个矩形
_SPEC_D05 = {
    '矩形 7': (8.93, 8.71, 3.84, 0.90),
    '矩形 8': (8.93, 9.80, 3.84, 0.90),
    '矩形 13': (8.93, 10.89, 3.84, 0.89),
    '矩形 24': (25.84, 8.71, 3.84, 0.90),
    '矩形 26': (25.84, 9.80, 3.84, 0.90),
    '矩形 27': (25.84, 10.89, 3.84, 0.89),
}
# D06 第1页“育德”“定目标”“划边界”“立规范”“成才”五处文本框（细则只约束宽度与左边距）
_SPEC_D06 = {
    '文本框 14': (9.53, None, 2.58, None),
    '文本框 3': (16.68, None, 2.58, None),
    '文本框 4': (19.26, None, 2.71, None),
    '文本框 75': (16.68, None, 5.34, None),
    '文本框 30': (26.75, None, 2.58, None),
}
# D07 第1页“教育学视角”“技术学支撑”“心理发展理论”三个文本框
_SPEC_D07 = {
    '文本框 63': (14.48, 5.02, 2.99, 2.10),
    '文本框 6': (21.23, 5.02, 2.99, 2.10),
    '文本框 74': (17.83, 10.54, 2.99, 1.62),
}
# D08 第1页左右两个内容底框矩形
_SPEC_D08 = {
    '矩形 46': (8.72, 8.42, 4.24, 3.61),
    '矩形 22': (25.64, 8.42, 4.24, 3.61),
}
# D09 第1页底部虚线边框矩形
_SPEC_D09 = {
    '矩形 52': (8.02, 14.92, 22.54, 1.79),
}
# D10 第1页各辅助图元（尺寸与跨度）
_SPEC_D10 = {
    '任意多边形 17': (8.02, None, 6.54, 6.17),
    '任意多边形 20': (24.00, None, 6.54, 6.17),
    '右箭头 18': (7.11, None, 0.78, 0.65),
    '右箭头 28': (7.11, None, 0.78, 0.65),
    '右箭头 32': (7.11, None, 0.78, 0.65),
    '加号 5': (None, None, 1.07, 1.14),
    '直接箭头连接符 11': (None, None, 3.41, None),
    '直接箭头连接符 2': (None, None, 3.78, None),
    '直接连接符 53': (None, None, None, 3.44),
    '直接连接符 54': (None, None, None, 3.29),
    '直接连接符 19': (None, None, None, 2.08),
    '直接连接符 21': (None, None, None, 2.16),
}
# D10 中以“约”描述的斜向箭头跨度，单独放宽容差
_SPEC_D10_SPAN = {
    '直接箭头连接符 34': (None, None, 5.03, None),
    '直接箭头连接符 35': (None, None, 5.21, None),
}
# D11 第2页“机制层”“实施层”“目标层”三个竖排圆角矩形
_SPEC_D11 = {
    '圆角矩形 36': (2.34, None, 0.80, 2.83),
    '圆角矩形 126': (2.34, None, 0.80, 2.71),
    '圆角矩形 127': (2.34, None, 0.80, 2.78),
}
# D12 第2页“联动贯通”“持续推进”两个矩形
_SPEC_D12 = {
    '矩形 11': (6.56, 1.38, 4.28, 1.07),
    '矩形 23': (25.64, 1.38, 4.28, 1.07),
}
# D13 第2页“教师统筹审核”等五个圆角矩形
_SPEC_D13 = {
    '圆角矩形 44': (5.66, 3.59, 4.97, 0.90),
    '圆角矩形 52': (11.09, 3.59, 4.28, 0.90),
    '圆角矩形 51': (15.85, 3.59, 4.82, 0.90),
    '圆角矩形 50': (21.19, 3.59, 3.73, 0.90),
    '圆角矩形 48': (25.73, 3.59, 3.79, 0.90),
}
# D14 第2页“学情识别”等五个圆角矩形
_SPEC_D14 = {
    '圆角矩形 64': (5.66, None, 2.80, 1.01),
    '圆角矩形 65': (11.17, None, 2.80, 1.01),
    '圆角矩形 66': (16.68, None, 2.80, 1.01),
    '圆角矩形 67': (22.12, None, 2.80, 1.01),
    '圆角矩形 1': (27.68, None, 2.80, 1.01),
}
# D15 第2页“信息汇集”至“规则完善”共15个竖排圆角矩形
_SPEC_D15 = dict(
    (n, (None, None, 0.80, 2.40)) for n in (
        '圆角矩形 74', '圆角矩形 75', '圆角矩形 76', '圆角矩形 77', '圆角矩形 78',
        '圆角矩形 79', '圆角矩形 80', '圆角矩形 81', '圆角矩形 82', '圆角矩形 83',
        '圆角矩形 84', '圆角矩形 85', '圆角矩形 2', '圆角矩形 3', '圆角矩形 4',
    )
)
assert len(_SPEC_D15) == 15
# D16 第2页五个成效圆角矩形
_SPEC_D16 = dict(
    (n, (None, None, 5.15, 2.00)) for n in (
        '圆角矩形 96', '圆角矩形 98', '圆角矩形 100', '圆角矩形 102', '圆角矩形 8',
    )
)
# D17 第2页三个大边框矩形
_SPEC_D17 = {
    '矩形 53': (None, None, 26.77, 5.70),
    '矩形 35': (None, None, 27.74, 8.12),
    '矩形 125': (None, None, 27.75, 2.79),
}
# D18 第2页各辅助图元（尺寸与跨度）
_SPEC_D18 = {
    '左大括号 25': (17.90, None, 0.70, 22.39),
    '文本框 26': (None, None, 2.91, None),
    '文本框 130': (None, None, 1.58, None),
    '左箭头 32': (None, None, 1.30, 0.70),
    '左箭头 33': (None, None, 1.30, 0.70),
    '左箭头 34': (None, None, 0.70, 1.20),
}
for _n in ('肘形连接符 6', '肘形连接符 12', '肘形连接符 90', '肘形连接符 91',
           '肘形连接符 92', '肘形连接符 93', '肘形连接符 120', '肘形连接符 121',
           '肘形连接符 122', '肘形连接符 123'):
    _SPEC_D18[_n] = (None, None, None, 2.73)
for _n in ('直接箭头连接符 9', '直接箭头连接符 111', '直接箭头连接符 113',
           '直接箭头连接符 115', '直接箭头连接符 117'):
    _SPEC_D18[_n] = (None, None, None, 1.40)
for _n in ('直接箭头连接符 13', '直接箭头连接符 14', '直接箭头连接符 128'):
    _SPEC_D18[_n] = (None, None, 0.98, None)
# D18 中以“约”描述的竖向短连接符与横向箭头跨度，单独放宽容差
_SPEC_D18_SPAN = {
    '直接连接符 5': (None, None, None, 0.61),
    '直接连接符 86': (None, None, None, 0.61),
    '直接连接符 87': (None, None, None, 0.61),
    '直接连接符 88': (None, None, None, 0.61),
    '直接连接符 89': (None, None, None, 0.61),
    '直接箭头连接符 27': (None, None, 2.71, None),
    '直接箭头连接符 28': (None, None, 2.71, None),
    '直接箭头连接符 29': (None, None, 2.61, None),
    '直接箭头连接符 30': (None, None, 2.71, None),
}

# D01 两页图形整体覆盖范围：(左, 右, 上, 下)，单位 cm
_SPEC_BBOX = {
    1: (3.34, 30.56, 1.94, 16.72),
    2: (2.34, 31.87, 1.38, 14.59),
}

# D19 各形状既有的填充与线条设置：形状名 -> (填充, 线条填充, 线型, 首端型, 末端型)
# 填充/线条填充取值：'none' 无填充、'solid:scheme=xxx' 主题色、'solid:srgb=XXXXXX' 固定色、
# '-' 表示该形状未显式声明（沿用主题默认）。
_SPEC_D19 = {
    (1, '矩形 52'): ('none', 'solid:scheme=tx1', 'dash', '-', '-'),
    (1, '圆角矩形 15'): ('solid:srgb=D9D9D9', 'solid:scheme=tx1', '-', '-', '-'),
    (1, '圆角矩形 23'): ('solid:srgb=D9D9D9', 'solid:scheme=tx1', '-', '-', '-'),
    (1, '圆角矩形 29'): ('solid:srgb=D9D9D9', 'solid:scheme=tx1', '-', '-', '-'),
    (1, '等腰三角形 47'): ('none', 'solid:scheme=tx1', '-', '-', '-'),
    (1, '等腰三角形 48'): ('none', 'solid:scheme=tx1', '-', '-', '-'),
    (1, '等腰三角形 49'): ('none', 'solid:scheme=tx1', '-', '-', '-'),
    (1, '等腰三角形 50'): ('none', 'solid:scheme=tx1', '-', '-', '-'),
    (1, '直接箭头连接符 11'): ('-', 'solid:scheme=tx1', '-', '-', 'stealth'),
    (1, '直接箭头连接符 2'): ('-', 'solid:scheme=tx1', '-', '-', 'stealth'),
    (1, '直接箭头连接符 34'): ('-', 'solid:scheme=tx1', '-', 'stealth', 'none'),
    (2, '矩形 53'): ('none', 'solid:scheme=tx1', 'solid', '-', '-'),
    (2, '矩形 35'): ('none', 'solid:scheme=tx1', 'lgDashDot', '-', '-'),
    (2, '矩形 125'): ('none', 'solid:scheme=tx1', 'lgDashDot', '-', '-'),
    (2, '圆角矩形 36'): ('solid:srgb=E2E0E1', 'solid:scheme=tx1', '-', '-', '-'),
    (2, '圆角矩形 126'): ('solid:srgb=E2E0E1', 'solid:scheme=tx1', '-', '-', '-'),
    (2, '圆角矩形 127'): ('solid:srgb=E2E0E1', 'solid:scheme=tx1', '-', '-', '-'),
    (2, '肘形连接符 90'): ('solid:scheme=bg1', 'solid:scheme=tx1', '-', 'none', 'none'),
    (2, '左大括号 25'): ('-', 'solid:scheme=tx1', '-', '-', '-'),
    (2, '左箭头 32'): ('solid:srgb=B7B7B7', 'none', '-', '-', '-'),
}


def _q(tag):
    """DrawingML 命名空间限定名"""
    return '{%s}%s' % (_A, tag)


def _cm(value):
    """EMU -> 厘米"""
    return 0.0 if value is None else Emu(int(value)).cm


def _open_presentation(file_path):
    """打开 .pptx 演示文档（python-pptx 仅支持 OOXML/.pptx）"""
    return Presentation(file_path)


def _shape_map(prs):
    """建立 {(页码, 形状名): 形状} 索引，并统计每页顶层图形对象个数"""
    smap = {}
    counts = {}
    for si, slide in enumerate(prs.slides, 1):
        n = 0
        for sh in slide.shapes:
            n += 1
            smap.setdefault((si, sh.name), sh)
        counts[si] = n
    return smap, counts


def _sp_pr(shape):
    """取形状的 p:spPr（形状属性）节点；连接符/占位符同样适用"""
    el = shape._element
    node = el.find('{%s}spPr' % _P)
    if node is None:
        node = el.find(_q('spPr'))
    return node


def _color_of(node):
    """解析 solidFill 的颜色引用：主题色 scheme=xxx / 固定色 srgb=XXXXXX"""
    if node is None:
        return '?'
    sc = node.find(_q('schemeClr'))
    if sc is not None:
        return 'scheme=%s' % sc.get('val')
    sr = node.find(_q('srgbClr'))
    if sr is not None:
        return 'srgb=%s' % sr.get('val')
    return '?'


def _fill_line_sig(shape):
    """返回 (填充, 线条填充, 线型, 首端型, 末端型) 五元组，与 _SPEC_D19 同构"""
    spPr = _sp_pr(shape)
    if spPr is None:
        return ('?', '?', '?', '?', '?')
    fill = '-'
    for tag in ('noFill', 'solidFill', 'gradFill', 'pattFill', 'blipFill', 'grpFill'):
        node = spPr.find(_q(tag))
        if node is None:
            continue
        if tag == 'noFill':
            fill = 'none'
        elif tag == 'solidFill':
            fill = 'solid:%s' % _color_of(node)
        else:
            fill = tag
        break
    ln = spPr.find(_q('ln'))
    if ln is None:
        return (fill, '-', '-', '-', '-')
    lnfill = '-'
    if ln.find(_q('noFill')) is not None:
        lnfill = 'none'
    else:
        node = ln.find(_q('solidFill'))
        if node is not None:
            lnfill = 'solid:%s' % _color_of(node)
    dash = ln.find(_q('prstDash'))
    head = ln.find(_q('headEnd'))
    tail = ln.find(_q('tailEnd'))
    return (fill, lnfill,
            dash.get('val') if dash is not None else '-',
            head.get('type') if head is not None else '-',
            tail.get('type') if tail is not None else '-')


def _line_width_cm(shape):
    """形状边框显式线宽（cm）；未显式声明返回 None（沿用主题默认线宽）"""
    spPr = _sp_pr(shape)
    if spPr is None:
        return None
    ln = spPr.find(_q('ln'))
    if ln is None or not ln.get('w'):
        return None
    return _cm(int(ln.get('w')))


def _runs(shape):
    """形状内的全部文字 run（按段落顺序展开）"""
    if not shape.has_text_frame:
        return []
    out = []
    for para in shape.text_frame.paragraphs:
        for run in para.runs:
            out.append(run)
    return out


def _rpr(run):
    return run._r.find(_q('rPr'))


def _run_size(run):
    """run 实际生效的字号（百分之一磅）：run 级 sz → 所属段落 defRPr/endParaRPr；
    都没有则返回 None。"""
    rPr = _rpr(run)
    if rPr is not None and rPr.get('sz'):
        return int(rPr.get('sz'))
    para = run._r.getparent()
    if para is not None:
        for tag in ('pPr', 'endParaRPr'):
            node = para.find(_q(tag))
            if node is None:
                continue
            target = node.find(_q('defRPr')) if tag == 'pPr' else node
            if target is not None and target.get('sz'):
                return int(target.get('sz'))
    return None


def _run_font(run, kind):
    """取 run 上显式声明的字体名：kind='ea' 中文（东亚）字体，kind='latin' 西文字体"""
    rPr = _rpr(run)
    if rPr is None:
        return None
    node = rPr.find(_q(kind))
    return node.get('typeface') if node is not None else None


def _run_bold(run):
    rPr = _rpr(run)
    return bool(rPr is not None and rPr.get('b') == '1')


def _visible_runs(shape):
    """仅保留含非空白字符的 run：纯空格 run 不承载可见文字，不参与字号/字体判定"""
    return [r for r in _runs(shape) if r.text and r.text.strip()]


def _char_w(ch):
    """6 磅字号下单字的字面宽度（cm）：全角按 1 字宽，半角（数字/西文/空格）按 0.5 字宽"""
    return EM_6PT if ord(ch) >= 0x2E80 else 0.5 * EM_6PT


def _wrap_lines(text, avail_w):
    """按可用宽度做贪心断行，返回 (最长行宽 cm, 总行数)"""
    longest = 0.0
    total = 0
    for src in text.split('\n'):
        cur = 0.0
        cnt = 1
        for ch in src:
            cw = _char_w(ch)
            if cur > 0 and cur + cw > avail_w + 0.02:
                longest = max(longest, cur)
                cur = cw
                cnt += 1
            else:
                cur += cw
        longest = max(longest, cur)
        total += cnt
    return longest, total


def _text_fits(shape, text):
    """6 磅字号下文字能否完整显示在形状内：
       行内方向按扣除默认左右内边距后的宽度断行（可用宽度不足一个字时视为内边距被压缩）；
       跨行方向按 1.2 倍行距累计行高，与形状高度比较。"""
    w = _cm(shape.width)
    h = _cm(shape.height)
    avail = w - INSET_SUM
    if avail < EM_6PT:
        avail = w
    if avail <= 0:
        return False, '可用宽度为0'
    longest, lines = _wrap_lines(text, avail)
    if longest > avail + 0.02:
        return False, '行宽%.2f>可用宽%.2f' % (longest, avail)
    need_h = lines * EM_6PT * LINE_FACTOR
    if need_h > h + 0.02:
        return False, '%d行需高%.2f>形状高%.2f' % (lines, need_h, h)
    return True, ''


def _check_spec(smap, si, spec, tol):
    """按 (页, 形状名) 逐项比对几何规格（cm），返回不符合项的说明列表"""
    bad = []
    for name in sorted(spec):
        L, T, W, H = spec[name]
        sh = smap.get((si, name))
        if sh is None:
            bad.append('%s缺失' % name)
            continue
        for tag, exp, act in (('左', L, _cm(sh.left)), ('上', T, _cm(sh.top)),
                              ('宽', W, _cm(sh.width)), ('高', H, _cm(sh.height))):
            if exp is not None and abs(act - exp) > tol:
                bad.append('%s%s%.2f≠%.2f' % (name, tag, act, exp))
    return bad


def _error_result(result, message):
    """脚本未能完成评估（目录不存在、未找到文档、运行异常）时的统一返回：
    维度一根本没有开始评估，因此 dim1_pass 置空字符串、dim1_reason 留空，
    错误摘要只写在 error 字段；分值与满分一并记 0。"""
    result["status"] = "error"
    result["error"] = message
    result["dim1_pass"] = ""
    result["dim1_reason"] = ""
    result["dim2_items"] = []
    result["total_score"] = 0
    result["max_score"] = 0
    return result


def evaluate(dir_path):
    """按接口约定返回结构化 dict；不 print、不改 sys.stdout、不 sys.exit。
    入参 dir_path 为“脚本所在目录的路径”，脚本自己在该目录里定位并打开被评估的
    .pptx（忽略以 ~$ 开头的 PowerPoint 临时锁文件）。脚本自身崩溃（含目录不存在、
    未找到文档等）由本函数捕获并返回 status='error'，此时 dim1_pass 为空字符串。"""
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
            raise FileNotFoundError(f"目录不存在：{dir_path}")
        cands = sorted(
            f for f in os.listdir(dir_path)
            if f.lower().endswith('.pptx') and not f.startswith('~$')
        )
        if not cands:
            raise FileNotFoundError(f"目录中未找到 .pptx 文件：{dir_path}")
        file_path = os.path.join(dir_path, cands[0])
        result["file_name"] = os.path.basename(file_path)
        return _evaluate_impl(file_path, result)
    except FileNotFoundError as exc:
        # 目录不存在 / 目录中没有 .pptx：错误信息直接给出可读的中文摘要
        return _error_result(result, str(exc))
    except Exception as exc:
        return _error_result(result, f"{type(exc).__name__}: {exc}")


def _evaluate_impl(file_path, result):
    """维度一 + 维度二的评估主体；内部构造 results 字符串列表，末尾按
    _DELTAS_ORDER 反解成结构化 dim2_items 填入 result。加分项写成 "+N：..."，
    减分项命中写成 "-N：..."，未命中一律写成 "+0：..." / "-0：..."。"""
    # ========== 维度1：可用性 ==========
    # 维度一只保留“可打开 pptx”一项：能被 python-pptx 正常打开即通过；
    # 文件损坏、非 OOXML/.pptx 等导致打不开时判 dim1_pass=False，dim1_reason 写明
    # 错误摘要，完成度总分记 0。
    try:
        prs = _open_presentation(file_path)
    except Exception as exc:
        result["dim1_pass"] = False
        result["dim1_reason"] = '文件无法作为 .pptx 打开：%s: %s' % (type(exc).__name__, exc)
        result["dim2_items"] = []
        result["total_score"] = 0
        result["max_score"] = MAX_SCORE_TOTAL
        return result

    smap, counts = _shape_map(prs)
    n_slides = len(prs.slides)
    results = []

    # ========== 维度2：完成度评分 ==========

    # --- 辅助闭包 ---
    def _keys_of(si):
        return [k for k in EXPECTED_TEXT if k[0] == si]

    def _size_bad(keys):
        """逐 run 检查字号是否为 6 号字（sz=600）；仅统计含可见字符的 run"""
        bad = []
        for key in sorted(keys, key=lambda k: (k[0], k[1])):
            sh = smap.get(key)
            if sh is None:
                bad.append('第%d页%s缺失' % key)
                continue
            runs = _visible_runs(sh)
            if not runs:
                bad.append('第%d页%s无文字' % key)
                continue
            for r in runs:
                sz = _run_size(r)
                if sz != TARGET_SZ:
                    bad.append('第%d页%s“%s”字号%s' % (key[0], key[1], r.text.strip()[:6], sz))
        return bad

    def _ea_bad(keys):
        """逐 run 检查中文（东亚）字体是否为宋体"""
        bad = []
        for key in sorted(keys, key=lambda k: (k[0], k[1])):
            sh = smap.get(key)
            if sh is None:
                bad.append('第%d页%s缺失' % key)
                continue
            for r in _visible_runs(sh):
                ea = _run_font(r, 'ea')
                if ea not in STRICT_SONGTI:
                    bad.append('第%d页%s“%s”中文字体%s'
                               % (key[0], key[1], r.text.strip()[:6], ea))
        return bad

    def _hcenter_gap(a, b):
        """两形状水平中心线的偏差（cm）"""
        return abs((_cm(a.left) + _cm(a.width) / 2.0)
                   - (_cm(b.left) + _cm(b.width) / 2.0))

    def _span(sh):
        return _cm(sh.left), _cm(sh.left) + _cm(sh.width)

    def _vspan(sh):
        return _cm(sh.top), _cm(sh.top) + _cm(sh.height)

    def _add(delta, text, neg=False):
        """追加一条评分结果；delta 为实得分（加分项 0 或正、减分项 0 或负）。
        加分项未命中写成 "+0：..."，减分项未命中写成 "-0：..."。"""
        if neg:
            head = '%d' % delta if delta else '-0'
        else:
            head = '+%d' % delta
        results.append('%s：%s' % (head, text))

    # --- 加分项 A01：第1页各形状与文本框内的全部文字字号均为 6 号字 ---
    bad = _size_bad(_keys_of(1))
    _add(5 if not bad else 0,
         '第1页各形状与文本框内的全部文字（“目标层”“核心层”“基础层”“价值引领根本目标”'
         '“个性学习落地”“学习者整体成长”“育德”“成才”“定目标”“划边界”“立规范”“价值涵养”'
         '“智慧启发”“品格践行”“知识奠基”“能力提升”“素养生成”“教育学视角”“技术学支撑”'
         '“心理发展理论”以及各“1.…2.…”条目）字号均为6号字（6磅，sz=600），无一处遗漏')

    # --- 加分项 A02：第2页各形状与文本框内的全部文字字号均为 6 号字 ---
    bad = _size_bad(_keys_of(2))
    _add(5 if not bad else 0,
         '第2页各形状与文本框内的全部文字（“机制层”“实施层”“目标层”“联动贯通”'
         '“智能技术支持小学生差异化学习”“持续推进”“教师统筹审核”“智能工具协助”'
         '“家校联动支持”“发展规律匹配”“持续复盘改进”“学情识别”“内容匹配”“路径调整”'
         '“反馈评价”“循环改进”、“信息汇集”至“规则完善”共15个竖排小框、5个成效框文字、'
         '“运行逻辑”“规范内容”）字号均为6号字（6磅，sz=600），无一处遗漏')

    # --- 加分项 A03：第1页三处文字的中文字体改为宋体 ---
    bad = _ea_bad(KAITI_SHAPES)
    _add(3 if not bad else 0,
         '第1页“教育学视角”“技术学支撑”“心理发展理论”三处文字的中文字体改为宋体，'
         '未保留楷体等其他中文字体')

    # --- 加分项 A04：第1页长条圆角矩形宽度不超过 8cm ---
    sh = smap.get((1, '圆角矩形 37'))
    w = _cm(sh.width) if sh is not None else -1.0
    _add(5 if (sh is not None and w <= LIMIT_8CM) else 0,
         '第1页“1.儿童成长规律与学习特征 2.课程标准素养目标与育人要求”所在圆角矩形'
         '宽度由18.89cm收窄至8cm以内')

    # --- 加分项 A05：第2页顶部圆角矩形宽度不超过 8cm ---
    sh = smap.get((2, '圆角矩形 21'))
    w = _cm(sh.width) if sh is not None else -1.0
    _add(3 if (sh is not None and w <= LIMIT_8CM) else 0,
         '第2页“智能技术支持小学生差异化学习”所在圆角矩形宽度由10.87cm收窄至8cm以内')

    # --- 加分项 A06：第1页三处条目形状宽度均不超过 8cm ---
    bad = []
    for key in ITEM_SHAPES[:3]:
        sh = smap.get(key)
        if sh is None:
            bad.append('%s缺失' % key[1])
        elif _cm(sh.width) > LIMIT_8CM:
            bad.append('%s宽%.2fcm' % (key[1], _cm(sh.width)))
    _add(3 if not bad else 0,
         '第1页“1.儿童认知发展理论 2.自主学习支持理论”“1.差异化教学思想 2.支架式成长理论”'
         '“1.智能协作理论 2.生成技术应用理论”三处文字所在形状与文本框宽度'
         '由11.60cm、10.05cm、9.60cm均收窄至8cm以内')

    # --- 加分项 A07：收窄后仍水平居中、未越出所属框体、未与相邻图形交叠 ---
    # 只针对“确实已收窄到 8cm 以内”的目标形状做检查；若无任何形状完成收窄，
    # 说明并未动手改宽度，本项不得分。
    narrowed = set()
    for key in ((1, '圆角矩形 37'), (2, '圆角矩形 21'),
                (1, '矩形 51'), (1, '文本框 81'), (1, '文本框 84')):
        sh = smap.get(key)
        if sh is not None and _cm(sh.width) <= LIMIT_8CM:
            narrowed.add(key)
    bad = []
    pairs = [((1, '圆角矩形 37'), (1, '矩形 52'), '底部虚线框'),
             ((1, '矩形 51'), (1, '等腰三角形 47'), '外层三角形'),
             ((1, '文本框 81'), (1, '等腰三角形 48'), '左侧小三角形'),
             ((1, '文本框 84'), (1, '等腰三角形 49'), '右侧小三角形')]
    for key, ref_key, ref_name in pairs:
        if key not in narrowed:
            continue
        sh, ref = smap.get(key), smap.get(ref_key)
        if sh is None or ref is None:
            bad.append('%s或%s缺失' % (key[1], ref_key[1]))
            continue
        gap = _hcenter_gap(sh, ref)
        if gap > TOL_CENTER:
            bad.append('%s与%s中心线偏差%.2fcm' % (key[1], ref_name, gap))
    # 底部条目框须完整落在虚线框内（水平与垂直方向）
    sh, ref = smap.get((1, '圆角矩形 37')), smap.get((1, '矩形 52'))
    if (1, '圆角矩形 37') in narrowed and sh is not None and ref is not None:
        (l1, r1), (l2, r2) = _span(sh), _span(ref)
        (t1, b1), (t2, b2) = _vspan(sh), _vspan(ref)
        if l1 < l2 - 0.05 or r1 > r2 + 0.05 or t1 < t2 - 0.05 or b1 > b2 + 0.05:
            bad.append('圆角矩形 37越出底部虚线框')
    # 顶部条目框须留在同排两矩形之间且不与左右箭头交叠
    sh = smap.get((2, '圆角矩形 21'))
    left_box, right_box = smap.get((2, '矩形 11')), smap.get((2, '矩形 23'))
    if ((2, '圆角矩形 21') in narrowed and sh is not None
            and left_box is not None and right_box is not None):
        l1, r1 = _span(sh)
        row_c = (_span(left_box)[0] + _span(right_box)[1]) / 2.0
        if abs((l1 + r1) / 2.0 - row_c) > TOL_CENTER:
            bad.append('圆角矩形 21未与同排保持水平居中')
        if l1 < _span(left_box)[1] - 0.05 or r1 > _span(right_box)[0] + 0.05:
            bad.append('圆角矩形 21与同排矩形交叠')
        for nb in ('左箭头 32', '左箭头 33'):
            arrow = smap.get((2, nb))
            if arrow is None:
                continue
            al, ar = _span(arrow)
            if l1 < ar - 0.05 and r1 > al + 0.05:
                bad.append('圆角矩形 21与%s交叠' % nb)
    _add(3 if (narrowed and not bad) else 0,
         '完成收窄的形状仍与其所属虚线框、背景矩形或所在横排保持水平居中，'
         '未越出所属框体边界，未与左右相邻的箭头、形状发生交叠（无任何形状完成收窄则不得分）')

    # --- 减分项 D01：两页图形整体覆盖范围 ---
    bad = []
    for si, slide in enumerate(prs.slides, 1):
        xs, ys = [], []
        for sh in slide.shapes:
            l, t = _cm(sh.left), _cm(sh.top)
            r, b = l + _cm(sh.width), t + _cm(sh.height)
            # 本身即跨出页面的图元（如超高的大括号）不参与整体覆盖范围统计
            if l < -0.05 or t < -0.05 or r > PAGE_W + 0.05 or b > PAGE_H + 0.05:
                continue
            xs.extend([l, r])
            ys.extend([t, b])
        exp = _SPEC_BBOX[si]
        act = (min(xs), max(xs), min(ys), max(ys)) if xs else (0.0, 0.0, 0.0, 0.0)
        for i, tag in enumerate(('左', '右', '上', '下')):
            if abs(act[i] - exp[i]) > TOL_BBOX:
                bad.append('第%d页%s边界%.2f≠%.2f' % (si, tag, act[i], exp[i]))
    _add(-5 if bad else 0,
         '两页图形整体覆盖范围不满足第1页横向自左3.34cm延伸至30.56cm（约27.2cm宽）、'
         '纵向自上1.94cm延伸至16.72cm（约14.8cm高），第2页横向自左2.34cm延伸至31.87cm'
         '（约29.5cm宽）、纵向自上1.38cm延伸至14.59cm，页面四周出现大面积空白', neg=True)

    # --- 减分项 D02：第1页三个层级标签圆角矩形 ---
    bad = _check_spec(smap, 1, _SPEC_D02, TOL_POS)
    _add(-1 if bad else 0,
         '第1页“目标层”“核心层”“基础层”三个圆角矩形不满足宽3.53cm、高1.28cm，'
         '左边距均为3.34cm，上边距分别为1.94cm、9.43cm、15.14cm', neg=True)

    # --- 减分项 D03：第1页顶部三个矩形 ---
    bad = _check_spec(smap, 1, _SPEC_D03, TOL_POS)
    _add(-1 if bad else 0,
         '第1页“价值引领根本目标”“个性学习落地”“学习者整体成长”三个矩形不满足高1.28cm、'
         '宽分别为5.34cm、5.34cm、4.71cm，且沿页面上部同一横排（上边距1.94-2.00cm）'
         '自左8.02cm、16.71cm、25.83cm排布', neg=True)

    # --- 减分项 D04：第1页金字塔四个等腰三角形 ---
    bad = _check_spec(smap, 1, _SPEC_D04, TOL_POS)
    _add(-3 if bad else 0,
         '第1页金字塔结构不满足外层等腰三角形宽11.60cm、高9.40cm（位于左13.55cm、上4.77cm），'
         '内部三个小等腰三角形宽4.80cm、高3.90cm（分别位于左13.55cm、20.33cm、16.91cm）', neg=True)

    # --- 减分项 D05：第1页六个素养矩形 ---
    bad = _check_spec(smap, 1, _SPEC_D05, TOL_POS)
    _add(-3 if bad else 0,
         '第1页“价值涵养”“智慧启发”“品格践行”“知识奠基”“能力提升”“素养生成”六个矩形'
         '不满足宽3.84cm、高0.90cm（末行0.89cm），左侧三个左对齐于8.93cm、'
         '右侧三个左对齐于25.84cm，上边距依次为8.71cm、9.80cm、10.89cm', neg=True)

    # --- 减分项 D06：第1页五处文本框 ---
    bad = _check_spec(smap, 1, _SPEC_D06, TOL_POS)
    _add(-3 if bad else 0,
         '第1页“育德”“定目标”“划边界”“立规范”“成才”五处文本框不满足宽度依次为2.58cm、'
         '2.58cm、2.71cm、5.34cm、2.58cm，左边距依次为9.53cm、16.68cm、19.26cm、'
         '16.68cm、26.75cm', neg=True)

    # --- 减分项 D07：第1页视角/支撑/理论三个文本框 ---
    bad = _check_spec(smap, 1, _SPEC_D07, TOL_POS)
    _add(-1 if bad else 0,
         '第1页“教育学视角”“技术学支撑”两个文本框不满足宽2.99cm、高2.10cm'
         '（分别位于左14.48cm与21.23cm、上5.02cm），“心理发展理论”文本框不满足宽2.99cm、'
         '高1.62cm（位于左17.83cm、上10.54cm）', neg=True)

    # --- 减分项 D08：第1页左右两个内容底框矩形 ---
    bad = _check_spec(smap, 1, _SPEC_D08, TOL_POS)
    _add(-1 if bad else 0,
         '第1页左右两个内容底框矩形不满足宽4.24cm、高3.61cm'
         '（分别位于左8.72cm与25.63cm、上8.42cm）', neg=True)

    # --- 减分项 D09：第1页底部虚线边框矩形 ---
    bad = _check_spec(smap, 1, _SPEC_D09, TOL_POS)
    _add(-1 if bad else 0,
         '第1页底部虚线边框矩形不满足宽22.54cm、高1.79cm，位于左8.02cm、上14.92cm', neg=True)

    # --- 减分项 D10：第1页各辅助图元 ---
    bad = _check_spec(smap, 1, _SPEC_D10, TOL_POS)
    bad += _check_spec(smap, 1, _SPEC_D10_SPAN, TOL_SPAN)
    _add(-3 if bad else 0,
         '第1页各辅助图元不满足其尺寸与跨度（两个任意多边形宽6.54cm高6.17cm、'
         '三个右箭头宽0.78cm高0.65cm且左对齐于7.11cm、加号形状宽1.07cm高1.14cm、'
         '顶部两条水平箭头长3.41cm与3.78cm、两条斜向箭头跨度约5.03cm与5.21cm、'
         '上部两条竖直连接符长约3.44cm与3.29cm、下部两条竖直连接符长约2.08cm与2.16cm）', neg=True)

    # --- 减分项 D11：第2页三个竖排层级圆角矩形 ---
    bad = _check_spec(smap, 2, _SPEC_D11, TOL_POS)
    _add(-1 if bad else 0,
         '第2页“机制层”“实施层”“目标层”三个竖排圆角矩形不满足宽0.80cm、'
         '高分别为2.83cm、2.71cm、2.78cm，左边距均为2.34cm', neg=True)

    # --- 减分项 D12：第2页联动贯通/持续推进两矩形 ---
    bad = _check_spec(smap, 2, _SPEC_D12, TOL_POS)
    _add(-1 if bad else 0,
         '第2页“联动贯通”“持续推进”两个矩形不满足宽4.28cm、高1.07cm，'
         '分别位于左6.56cm与25.64cm、上1.38cm', neg=True)

    # --- 减分项 D13：第2页五个措施圆角矩形 ---
    bad = _check_spec(smap, 2, _SPEC_D13, TOL_POS)
    _add(-3 if bad else 0,
         '第2页“教师统筹审核”“智能工具协助”“家校联动支持”“发展规律匹配”“持续复盘改进”'
         '五个圆角矩形不满足高0.90cm、宽依次为4.97cm、4.28cm、4.82cm、3.73cm、3.79cm，'
         '且沿上3.59cm同一横排自左5.66cm、11.09cm、15.85cm、21.19cm、25.73cm均匀排布', neg=True)

    # --- 减分项 D14：第2页五个环节圆角矩形 ---
    bad = _check_spec(smap, 2, _SPEC_D14, TOL_POS)
    _add(-3 if bad else 0,
         '第2页“学情识别”“内容匹配”“路径调整”“反馈评价”“循环改进”五个圆角矩形'
         '不满足宽2.80cm、高1.01cm，自左5.66cm、11.17cm、16.68cm、22.12cm、27.68cm排布', neg=True)

    # --- 减分项 D15：第2页15个竖排圆角矩形 ---
    bad = _check_spec(smap, 2, _SPEC_D15, TOL_POS)
    for name in sorted(_SPEC_D15):
        sh = smap.get((2, name))
        if sh is None:
            continue
        ok, why = _text_fits(sh, sh.text_frame.text if sh.has_text_frame else '')
        if not ok:
            bad.append('%s四字未能竖向完整显示（%s）' % (name, why))
    _add(-5 if bad else 0,
         '第2页“信息汇集”“综合分析”“教师确认”“特征匹配”“资源生成”“教师审看”“层级安排”'
         '“及时调整”“人工修订”“多维评价”“鼓励反馈”“家校互通”“状态更新”“策略调整”'
         '“规则完善”共15个竖排圆角矩形不满足宽0.80cm、高2.40cm，四字竖向单列完整显示', neg=True)

    # --- 减分项 D16：第2页五个成效圆角矩形 ---
    bad = _check_spec(smap, 2, _SPEC_D16, TOL_POS)
    for name in sorted(_SPEC_D16):
        sh = smap.get((2, name))
        if sh is None:
            continue
        text = sh.text_frame.text if sh.has_text_frame else ''
        if len(text.split('\n')) != 2:
            bad.append('%s文字未分两行' % name)
            continue
        ok, why = _text_fits(sh, text)
        if not ok:
            bad.append('%s文字未能完整显示（%s）' % (name, why))
    _add(-3 if bad else 0,
         '第2页“准确把握差异，降低判断偏差；”“形成匹配资源，守住内容规范；”'
         '“对接成长区间，促进逐级提升；”“汇聚支持力量，维护学习信心；”'
         '“不断优化流程，回应成长需要。”五个圆角矩形不满足宽5.15cm、高2.00cm，'
         '且每框文字分两行完整显示', neg=True)

    # --- 减分项 D17：第2页三个大边框矩形 ---
    bad = _check_spec(smap, 2, _SPEC_D17, TOL_POS)
    _add(-1 if bad else 0,
         '第2页三个大边框矩形不满足尺寸为宽26.77cm高5.70cm、宽27.74cm高8.12cm、'
         '宽27.75cm高2.79cm', neg=True)

    # --- 减分项 D18：第2页各辅助图元 ---
    bad = _check_spec(smap, 2, _SPEC_D18, TOL_POS)
    bad += _check_spec(smap, 2, _SPEC_D18_SPAN, TOL_SPAN)
    _add(-3 if bad else 0,
         '第2页各辅助图元不满足其尺寸与跨度（左大括号宽0.70cm高22.39cm且位于左17.90cm、'
         '“运行逻辑”文本框宽2.91cm、“规范内容”文本框宽1.58cm、'
         '三个左箭头为宽1.30cm高0.70cm两个与宽0.70cm高1.20cm一个、肘形连接符高2.73cm、'
         '竖向箭头高1.40cm、竖向短连接符高约0.61cm、横向箭头长2.61-2.71cm、'
         '层级引入箭头长0.98cm）', neg=True)

    # --- 减分项 D19：各形状既有的填充与线条设置 ---
    bad = []
    for key in sorted(_SPEC_D19, key=lambda k: (k[0], k[1])):
        sh = smap.get(key)
        if sh is None:
            bad.append('第%d页%s缺失' % key)
            continue
        got = _fill_line_sig(sh)
        if got != _SPEC_D19[key]:
            bad.append('第%d页%s为%s≠%s' % (key[0], key[1], got, _SPEC_D19[key]))
    _add(-3 if bad else 0,
         '各形状不满足既有的填充与线条设置（底部大矩形为无填充加黑色虚线边框、'
         '层级标签圆角矩形为浅灰色填充、金字塔等腰三角形为无填充加黑色实线边框、'
         '连接符为黑色实线且箭头端型不变），出现改色、改线型、改形状类型等多余修改', neg=True)

    # --- 减分项 D20：第1页其余文字的中文（东亚）字体被改成非宋体 ---
    bad = _ea_bad([k for k in _keys_of(1) if k not in KAITI_SHAPES])
    _add(-1 if bad else 0,
         '第1页除“教育学视角”“技术学支撑”“心理发展理论”外的其余文字不满足'
         '中文（东亚）字体为宋体，出现改用其他中文字体的多余修改', neg=True)

    # --- 减分项 D21：第2页文字的中文（东亚）字体被改成非宋体 ---
    bad = _ea_bad(_keys_of(2))
    _add(-1 if bad else 0,
         '第2页全部文字不满足中文（东亚）字体为宋体，出现改用其他中文字体的多余修改', neg=True)

    # --- 减分项 D22：数字与西文字体被改成非宋体 ---
    bad = []
    for key in sorted(EXPECTED_TEXT, key=lambda k: (k[0], k[1])):
        sh = smap.get(key)
        if sh is None:
            continue
        for r in _visible_runs(sh):
            if not any(('0' <= c <= '9') or ('A' <= c <= 'Z') or ('a' <= c <= 'z')
                       for c in r.text):
                continue
            la = _run_font(r, 'latin')
            if la not in STRICT_SONGTI:
                bad.append('第%d页%s“%s”西文字体%s'
                           % (key[0], key[1], r.text.strip()[:6], la))
    _add(-1 if bad else 0,
         '文字中的数字与西文（各条目前的“1.”“2.”等）不满足字体为宋体，'
         '出现改用Times New Roman、Arial、Calibri等西文字体的多余修改', neg=True)

    # --- 减分项 D23：文字在 6 磅下溢出/被截断，或边框线宽与形状尺寸不相称 ---
    bad = []
    for key in sorted(EXPECTED_TEXT, key=lambda k: (k[0], k[1])):
        sh = smap.get(key)
        if sh is None:
            bad.append('第%d页%s缺失' % key)
            continue
        ok, why = _text_fits(sh, sh.text_frame.text)
        if not ok:
            bad.append('第%d页%s文字溢出（%s）' % (key[0], key[1], why))
    for key in sorted(smap, key=lambda k: (k[0], k[1])):
        lw = _line_width_cm(smap[key])
        if lw is None:
            continue
        short = min(_cm(smap[key].width), _cm(smap[key].height))
        if short > 0 and lw > short * 0.20:
            bad.append('第%d页%s线宽%.2fcm占形状短边%.0f%%'
                       % (key[0], key[1], lw, lw / short * 100))
    _add(-5 if bad else 0,
         '全部文字不满足在6磅字号下完整显示于所属形状内部，出现文字溢出形状上下左右边界、'
         '被形状边框截断、换行行数超出形状高度而只能显示局部，或形状边框线宽与形状尺寸'
         '不相称而使形状糊成实心块、文字被边框挤压覆盖', neg=True)

    # --- 减分项 D24：123 个图形对象与全部文字未完整保留 ---
    # 文字判定范围为交付文档应有的 62 处文字：逐字（含空格、序号、标点）比对，
    # 不得删除、改写；本应无文字的形状里新增字符另由减分项单列。
    bad = []
    for si in (1, 2):
        if counts.get(si) != SHAPE_COUNT[si]:
            bad.append('第%d页图形对象%d个≠%d个' % (si, counts.get(si, 0), SHAPE_COUNT[si]))
    for key in sorted(EXPECTED_TEXT, key=lambda k: (k[0], k[1])):
        sh = smap.get(key)
        if sh is None:
            bad.append('第%d页%s缺失' % key)
            continue
        got = sh.text_frame.text if sh.has_text_frame else ''
        if got != EXPECTED_TEXT[key]:
            bad.append('第%d页%s文字为“%s”' % (key[0], key[1], got.replace('\n', '')[:14]))
    _add(-5 if bad else 0,
         '两页不满足共123个图形对象（第1页45个、第2页78个）数量完整且全部文字表述完整保留，'
         '出现删除文字、改写文字或改变图形对象数量（文字含序号“1.”“2.”，'
         '逗号、分号、句号等标点，以及条目之间的空格）', neg=True)

    # --- 减分项 D25：页数/页序/页面尺寸/备注页 ---
    bad = []
    if n_slides != 2:
        bad.append('共%d页' % n_slides)
    pw, ph = _cm(prs.slide_width), _cm(prs.slide_height)
    if abs(pw - PAGE_W) > 0.05 or abs(ph - PAGE_H) > 0.05:
        bad.append('页面尺寸%.2f×%.2fcm' % (pw, ph))
    if (1, '等腰三角形 47') not in smap or (1, '圆角矩形 15') not in smap:
        bad.append('第1页不是三层育人结构图')
    if (2, '圆角矩形 36') not in smap or (2, '圆角矩形 126') not in smap:
        bad.append('第2页不是机制层/实施层/目标层流程图')
    for si, slide in enumerate(prs.slides, 1):
        if slide.has_notes_slide and slide.notes_slide.notes_text_frame.text.strip():
            bad.append('第%d页备注页出现文字' % si)
    _add(-1 if bad else 0,
         '文档不满足共2页、页序为第1页三层育人结构图与第2页机制层/实施层/目标层流程图、'
         '页面尺寸为宽33.867cm×高19.05cm（16:9宽屏），或两个备注页文字内容出现新增删改', neg=True)

    # --- 减分项 D26：底部虚线框内出现多余字符 ---
    sh = smap.get(EMPTY_DASH_BOX)
    stray = ''
    if sh is not None and sh.has_text_frame:
        stray = sh.text_frame.text.strip()
    _add(-1 if stray else 0,
         '第1页底部虚线边框矩形（宽22.54cm、高1.79cm）内不满足无任何文字内容，'
         '出现了多余字符', neg=True)

    # --- 减分项 D27：四处条目文字被加粗 ---
    bad = []
    for key in ITEM_SHAPES:
        sh = smap.get(key)
        if sh is None:
            continue
        for r in _visible_runs(sh):
            if _run_bold(r):
                bad.append('第%d页%s“%s”为加粗' % (key[0], key[1], r.text.strip()[:6]))
                break
    _add(-1 if bad else 0,
         '第1页“1.儿童认知发展理论 2.自主学习支持理论”“1.差异化教学思想 2.支架式成长理论”'
         '“1.智能协作理论 2.生成技术应用理论”“1.儿童成长规律与学习特征 '
         '2.课程标准素养目标与育人要求”四处文字不满足字形为常规（非加粗）', neg=True)

    # --- 结果反解：把 results 里的 "+N：..." / "-N：..." 还原为结构化 dim2_items ---
    if len(results) != len(_DELTAS_ORDER):
        raise RuntimeError(
            f"dim2 项数不一致：results={len(results)} vs order={len(_DELTAS_ORDER)}")
    dim2_items = []
    total_score = 0
    for raw, max_delta in zip(results, _DELTAS_ORDER):
        head, _, rule = raw.partition("：")
        try:
            delta = int(head.replace("+", "").strip())
        except ValueError:
            delta = 0
        hit = delta != 0
        dim2_items.append({"rule": rule, "max_delta": max_delta,
                           "delta": max_delta if hit else 0, "hit": hit})
        total_score += (max_delta if hit else 0)
    result["dim2_items"] = dim2_items
    result["total_score"] = total_score
    result["max_score"] = MAX_SCORE_TOTAL
    return result


if __name__ == "__main__":
    if len(sys.argv) >= 2:
        _dir = sys.argv[1]
    else:
        _dir = os.path.dirname(os.path.abspath(__file__))
    _out = evaluate(_dir)
    print(json.dumps(_out, ensure_ascii=False, indent=2))
