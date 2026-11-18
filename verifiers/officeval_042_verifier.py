#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# pyright: reportUnknownMemberType=false, reportUnknownVariableType=false
# pyright: reportUnknownArgumentType=false, reportMissingImports=false
# pyright: reportPrivateUsage=false, reportGeneralTypeIssues=false
"""officeval_042_verifier
《幼儿游戏活动课程设计与实施》PPTX 改版任务完成度评分脚本。

依据同目录《细则.txt》实现：
  维度1 可用与可修改性 —— 不通过则完成度总分直接记 0；
  维度2 完成度评分细则 —— 16 项加分项（满分 44）＋13 项减分项（合计 -23）。

用法：
  python officeval_042_verifier.py <存放 .pptx 的目录>
返回：JSON（字段与 officeval_008_verifier 一致）。
"""

import os
import re
import io
import sys
import json
import struct
import hashlib
import zipfile

from pptx import Presentation
from PIL import Image, ImageStat

SCRIPT_ID = "042"

# 加分项满分（减分项不计入 max_score，只从实得分中扣除）
MAX_SCORE_TOTAL = 44
PENALTY_TOTAL = 23

_MAX_DELTAS_ORDER = [
    3,    # 01 第1页整页背景替换为幼儿园卡通插画且约50%透明度
    3,    # 02 第2页同一插画背景且约85%透明度
    1,    # 03 第1、2页背景满版铺满且左右未被裁切
    1,    # 04 第2页红色标题字体 Arial 且加粗
    3,    # 05 六个小标题文字改蓝色并全部加下划线
    5,    # 06 六个小标题全部改竖排且色块改为竖长方形
    5,    # 07 学生/任务/资源三组改为横向并排三栏
    3,    # 08 三栏水平间距相等
    3,    # 09 三个标题位于内容块正上方且水平居中
    3,    # 10 三栏内容自动换行完整显示、无堆叠重叠
    3,    # 11 “评价”标题及其内容已删除
    1,    # 12 三栏整体左右均衡
    3,    # 13 第6页音频为所给钢琴曲（34-38秒/双声道/48kHz）
    3,    # 14 第6页设有可点击播放该音频的按钮
    3,    # 15 播放按钮位于第6页左下角区域
    1,    # 16 播放按钮尺寸适中且未压盖页脚
    -1,   # 17 音频额外内嵌专辑封面图
    -3,   # 18 左下角叠放第二个播放按钮
    -3,   # 19 全部6页添加慢速2秒页面切换
    -1,   # 20 透明度烧进位图像素并放大重采样
    -1,   # 21 学生/任务/资源标签字号底色被改
    -1,   # 22 三栏内容字号字色被改
    -1,   # 23 底部“核心主线”样式与位置被改
    -1,   # 24 六个圆形数字标号字体被改为 Arial
    -1,   # 25 第2页其余元素（课程逻辑总览、页脚）被改
    -3,   # 26 第1页除背景外内容被改
    -3,   # 27 第3、4、5页内容与版式被改
    -1,   # 28 第6页除音频与按钮外内容被改
    -3,   # 29 主题字体回退表/默认文本样式/视图/标记部件/文档属性被破坏
]

assert sum(d for d in _MAX_DELTAS_ORDER if d > 0) == MAX_SCORE_TOTAL
assert -sum(d for d in _MAX_DELTAS_ORDER if d < 0) == PENALTY_TOTAL

NS = {
    'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
    'p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    'p14': 'http://schemas.microsoft.com/office/powerpoint/2010/main',
    'mc': 'http://schemas.openxmlformats.org/markup-compatibility/2006',
}
A = '{%s}' % NS['a']
P = '{%s}' % NS['p']
R = '{%s}' % NS['r']

EMU_PER_CM = 360000.0
PAGE_W_CM = 33.87
PAGE_H_CM = 19.05

POS_TOL = 0.3          # 位置/尺寸容差（cm）
PT_TOL = 0.5           # 字号容差（磅）
RGB_TOL = 8            # 文字颜色单通道容差
FILL_TOL = 16          # 填充色单通道容差
YAHEI = ('微软雅黑', 'Microsoft YaHei', 'Microsoft YaHei UI', 'MicrosoftYaHei')

# 所给幼儿园卡通插画（米黄底色、4名幼儿与积木/玩具/树木、960×539）的指纹。
# dHash 对“整体变亮/降对比度”（按透明度与白底混合）及重采样均不敏感，
# 因此淡化后的位图仍可据此判定是否同一幅插画。
SRC_IMG_DHASH = 0xA8AAAA9BE2AAAA22
SRC_IMG_MEAN = 220.95
SRC_IMG_STD = 52.17
SRC_IMG_W = 960
SRC_IMG_H = 539
SRC_IMG_ASPECT = 1.78
DHASH_MAX_HAM = 10

# 所给钢琴曲：36.0 秒、双声道、48kHz
SRC_WAV_SHA1 = '06ab342ce195fd9e5828ad176021a18723d48d54'
AUDIO_DUR_RANGE = (34.0, 38.0)
AUDIO_CH = 2
AUDIO_RATE = 48000
COVER_MAX_BYTES = 200000

TITLE_S2 = '整体教学设计：从学习基础到实践改进'
SUB_LABELS = ('学习诊断', '任务转化', '资源整合', '情境演练', '表现评价', '迭代成长')
COL_TITLES = ('学生', '任务', '资源')
COL_TEXTS = (
    '已有理论认知较清晰；组织经验与现场应变需要继续加强',
    '围绕游戏方案、环境创设、组织实施、入园观察形成连续任务',
    '数字课程、观察记录、游戏材料包、园所协同场景共同支撑',
)
DROPPED_TITLE = '评价'
DROPPED_TEXT = '关注安全意识、幼儿回应、组织节奏、反思修正四类表现'
FOOTER_S6 = '幼儿游戏活动课程设计与实施·6'

# 元素锚点表：(定位文字, 是否整体匹配, 字号磅, 加粗, 字色, 左边界cm, 上边界cm, 宽度cm)
# 字段为 None 表示该项不校验；位置尺寸单位 cm，容差 POS_TOL。
SLIDE2_KEEP = [
    ('课程逻辑总览', True, 10.5, None, '5D6470', 23.11, 1.02, 8.64),
    ('幼儿游戏活动课程设计与实施·2', True, 7.8, None, '7B8790', 1.57, 18.11, 12.19),
]

SLIDE1_KEEP = [
    ('幼儿游戏活动课程设计与实施', True, 32.0, True, '1F3A5F', 1.98, 2.67, 13.34),
    ('以岗位任务为牵引，融合体育游戏、音乐游戏与观察评价，形成可迁移的活动组织能力。',
     True, 15.5, None, '5D6470', 2.08, 7.06, 13.34),
    ('学情洞察', True, 10.5, True, 'FFFFFF', 2.08, 10.67, 2.79),
    ('任务驱动', True, 10.5, True, 'FFFFFF', 5.21, 10.67, 2.79),
    ('场景实践', True, 10.5, True, 'FFFFFF', 8.33, 10.67, 2.79),
    ('评价成长', True, 10.5, True, 'FFFFFF', 11.46, 10.67, 2.79),
]
SLIDE1_LABEL_H = 0.86

SLIDE3_KEEP = [
    ('教学目标：品格、知识与能力同步生长', True, 23.0, True, 'FF0000', 1.57, 0.86, 20.83),
    ('目标重构', True, 10.5, None, '5D6470', 23.11, 1.02, 8.64),
    ('品格目标', True, 13.0, True, 'FFFFFF', 1.93, 3.00, 5.28),
    ('知识目标', True, 13.0, True, 'FFFFFF', 8.41, 3.00, 5.28),
    ('能力目标', True, 13.0, True, 'FFFFFF', 14.88, 3.00, 5.28),
    ('尊重儿童兴趣与个体差异', False, 11.0, None, '5D6470', 1.93, 4.62, 5.59),
    ('理解体育游戏与音乐游戏的教育价值', False, 11.0, None, '5D6470', 8.41, 4.62, 5.59),
    ('能够完成完整游戏活动方案设计', False, 11.0, None, '5D6470', 14.88, 4.62, 5.59),
    ('会观察', False, 10.8, True, '1F3A5F', 2.54, 11.94, 6.05),
    ('会设计', False, 10.8, True, '1F3A5F', 10.31, 11.94, 6.05),
    ('会组织', False, 10.8, True, '1F3A5F', 18.08, 11.94, 6.05),
    ('会复盘', False, 10.8, True, '1F3A5F', 25.86, 11.94, 6.05),
    ('幼儿游戏活动课程设计与实施·3', True, 7.8, None, '7B8790', 1.57, 18.11, 12.19),
]

SLIDE4_KEEP = [
    ('重难点与破解路径：把“会做”变成“做得稳”', True, 23.0, True, 'FF0000', 1.57, 0.86, 20.83),
    ('教学支点', True, 10.5, None, '5D6470', 23.11, 1.02, 8.64),
    ('教学重点', True, 13.0, True, 'FFFFFF', 2.08, 3.00, 4.01),
    ('掌握游戏类型、教育价值、活动安全和观察评价方法', False, 11.5, None, '5D6470', 6.71, 3.00, 23.88),
    ('教学难点', True, 13.0, True, 'FFFFFF', 2.08, 8.31, 4.01),
    ('融合体育与音乐元素', False, 11.5, None, '5D6470', 6.71, 8.31, 23.88),
    ('案例拆解', False, 11.2, True, '1F3A5F', 6.71, 4.83, 6.91),
    ('微型演练', False, 11.2, True, '1F3A5F', 14.58, 4.83, 6.91),
    ('观察训练', False, 11.2, True, '1F3A5F', 22.45, 4.83, 6.91),
    ('融合工作坊', False, 11.2, True, '1F3A5F', 6.71, 10.13, 6.91),
    ('应变演练', False, 11.2, True, '1F3A5F', 14.58, 10.13, 6.91),
    ('同伴评议', False, 11.2, True, '1F3A5F', 22.45, 10.13, 6.91),
    ('课堂闭环', False, 12.2, True, '1F3A5F', 13.77, 13.67, 16.81),
    ('幼儿游戏活动课程设计与实施·4', True, 7.8, None, '7B8790', 1.57, 18.11, 12.19),
]
SLIDE4_PIC = (2.08, 13.67, 10.79, 2.84)

SLIDE5_KEEP = [
    ('基于工作流程的学习任务：双线并进，四步完成', True, 23.0, True, 'FF0000', 1.57, 0.86, 20.83),
    ('任务结构', True, 10.5, None, '5D6470', 23.11, 1.02, 8.64),
    ('方案设计', True, 12.5, True, 'FFFFFF', 2.64, 4.32, 4.62),
    ('环境创设', True, 12.5, True, 'FFFFFF', 9.98, 4.32, 4.62),
    ('组织实施', True, 12.5, True, 'FFFFFF', 17.32, 4.32, 4.62),
    ('实践反馈', True, 12.5, True, 'FFFFFF', 24.66, 4.32, 4.62),
    ('体能探索', True, 12.5, True, 'FFFFFF', 2.64, 9.91, 4.62),
    ('节奏表达', True, 12.5, True, '1F3A5F', 9.98, 9.91, 4.62),
    ('情境合作', True, 12.5, True, 'FFFFFF', 17.32, 9.91, 4.62),
    ('观察优化', True, 12.5, True, 'FFFFFF', 24.66, 9.91, 4.62),
    ('依据儿童发展需求设定活动目标与玩法规则', False, 10.4, None, '5D6470', 2.64, 6.22, 4.62),
    ('布置空间、材料与安全边界', False, 10.4, None, '5D6470', 9.98, 6.22, 4.62),
    ('运用口令、音乐、分组与示范推动活动流动', False, 10.4, None, '5D6470', 17.32, 6.22, 4.62),
    ('记录幼儿表现，调整难度并形成改进清单', False, 10.4, None, '5D6470', 24.66, 6.22, 4.62),
    ('把跑、跳、钻、平衡等动作转为趣味挑战', False, 10.4, None, '5D6470', 2.64, 11.81, 4.62),
    ('将节拍、旋律、强弱变化融入游戏情节', False, 10.4, None, '5D6470', 9.98, 11.81, 4.62),
    ('通过角色分工与规则协商提升互动质量', False, 10.4, None, '5D6470', 17.32, 11.81, 4.62),
    ('用观察证据支持评价与下一轮活动设计', False, 10.4, None, '5D6470', 24.66, 11.81, 4.62),
    ('体育游戏线', True, 10.5, True, 'FFFFFF', 0.30, 5.08, 2.08),
    ('音乐游戏线', True, 10.5, True, 'FFFFFF', 0.30, 10.67, 2.08),
    ('输出成果：一份完整活动方案、一段组织演练记录、一张观察评价表、一份改进说明。',
     True, 12.1, True, '1F3A5F', 2.64, 15.93, 26.54),
    ('幼儿游戏活动课程设计与实施·5', True, 7.8, None, '7B8790', 1.57, 18.11, 12.19),
]

SLIDE6_KEEP = [
    ('资源支撑与评价方式：让成长看得见', True, 23.0, True, 'FF0000', 1.57, 0.86, 20.83),
    ('评价闭环', True, 10.5, None, '5D6470', 23.11, 1.02, 8.64),
    ('安全与适配', True, 11.5, True, 'FFFFFF', 2.08, 3.17, 4.70),
    ('组织与沟通', True, 11.5, True, 'FFFFFF', 2.08, 5.16, 4.70),
    ('观察与分析', True, 11.5, True, 'FFFFFF', 2.08, 7.14, 4.70),
    ('反思与迁移', True, 11.5, True, 'FFFFFF', 2.08, 9.12, 4.70),
    ('能识别场地、材料和活动强度的风险点', True, 11.0, None, '5D6470', 7.42, 3.17, 10.92),
    ('指令清楚，示范准确，能稳定维持秩序', True, 11.0, None, '5D6470', 7.42, 5.16, 10.92),
    ('能基于儿童行为记录做出专业判断', True, 11.0, None, '5D6470', 7.42, 7.14, 10.92),
    ('能把反馈转化为下一轮活动优化', True, 11.0, None, '5D6470', 7.42, 9.12, 10.92),
    ('形成性评价', True, 12.0, True, 'FFFFFF', 20.32, 3.17, 4.52),
    ('资源支撑', True, 12.0, True, 'FFFFFF', 25.78, 3.17, 4.52),
    ('课堂观察', False, 13.0, None, '5D6470', 20.32, 4.78, 4.52),
    ('活动材料包', False, 13.0, None, '5D6470', 25.78, 4.78, 4.52),
    ('最终指向：从“完成任务”走向“理解儿童”，从“会组织活动”走向“能持续优化教育经验”。',
     True, 11.7, True, '1F3A5F', 21.03, 13.46, 9.19),
    ('幼儿游戏活动课程设计与实施·6', True, 7.8, None, '7B8790', 1.57, 18.11, 12.19),
]

# ---------------------------------------------------------------- 基础工具


def _cm(v):
    """EMU 转厘米。"""
    return 0.0 if v is None else float(v) / EMU_PER_CM


def _near(val, target, tol=POS_TOL):
    return val is not None and target is not None and abs(val - target) <= tol


def _norm(text):
    """去掉全部空白，便于文字定位比较。"""
    return re.sub(r'\s+', '', text or '')


def _rect(sh):
    """返回 (左, 上, 宽, 高)，单位 cm。"""
    return (_cm(sh.left), _cm(sh.top), _cm(sh.width), _cm(sh.height))


def _inter_area(r1, r2):
    x = min(r1[0] + r1[2], r2[0] + r2[2]) - max(r1[0], r2[0])
    y = min(r1[1] + r1[3], r2[1] + r2[3]) - max(r1[1], r2[1])
    return (x * y) if (x > 0 and y > 0) else 0.0


def _union_rect(r1, r2):
    l = min(r1[0], r2[0])
    t = min(r1[1], r2[1])
    r = max(r1[0] + r1[2], r2[0] + r2[2])
    b = max(r1[1] + r1[3], r2[1] + r2[3])
    return (l, t, r - l, b - t)


def _iter_shapes(container):
    """递归遍历形状（含组合内部）。"""
    for sh in container.shapes:
        yield sh
        if sh.shape_type is not None and str(sh.shape_type).startswith('GROUP'):
            try:
                for sub in _iter_shapes(sh):
                    yield sub
            except Exception:
                pass


def _shape_text(sh):
    if not getattr(sh, 'has_text_frame', False):
        return ''
    try:
        return sh.text_frame.text or ''
    except Exception:
        return ''


def _text_shapes(slide):
    return [sh for sh in _iter_shapes(slide) if _norm(_shape_text(sh))]


def _locate(slide, key, exact=True):
    """按文字定位形状：exact=True 整体相等，否则前缀匹配（优先整体相等）。"""
    key = _norm(key)
    cands = []
    for sh in _iter_shapes(slide):
        t = _norm(_shape_text(sh))
        if not t:
            continue
        if t == key:
            return sh
        if not exact and t.startswith(key):
            cands.append((len(t), sh))
    if cands:
        cands.sort(key=lambda x: x[0])
        return cands[0][1]
    return None


def _runs(sh):
    out = []
    if not getattr(sh, 'has_text_frame', False):
        return out
    try:
        for para in sh.text_frame.paragraphs:
            for run in para.runs:
                out.append(run)
    except Exception:
        pass
    return out


def _rpr(run):
    try:
        return run._r.find(A + 'rPr')
    except Exception:
        return None


def _size_pt(run):
    rpr = _rpr(run)
    if rpr is not None and rpr.get('sz'):
        try:
            return int(rpr.get('sz')) / 100.0
        except Exception:
            pass
    try:
        return None if run.font.size is None else run.font.size.pt
    except Exception:
        return None


def _bold(run):
    rpr = _rpr(run)
    if rpr is not None and rpr.get('b') is not None:
        return rpr.get('b') in ('1', 'true', 'on')
    try:
        return run.font.bold
    except Exception:
        return None


def _underlined(run):
    rpr = _rpr(run)
    u = rpr.get('u') if rpr is not None else None
    return bool(u) and u != 'none'


def _typefaces(run):
    """返回该 run 的 latin/ea/cs 字体名集合。"""
    out = []
    rpr = _rpr(run)
    if rpr is None:
        return out
    for tag in ('latin', 'ea', 'cs'):
        el = rpr.find(A + tag)
        if el is not None and el.get('typeface'):
            out.append(el.get('typeface'))
    return out


def _hex_of(el):
    """从 a:srgbClr / a:schemeClr / a:sysClr 元素解析出 RRGGBB。"""
    if el is None:
        return None
    tag = el.tag.split('}')[-1]
    if tag == 'srgbClr':
        return (el.get('val') or '').upper()
    if tag == 'sysClr':
        return (el.get('lastClr') or '').upper()
    return None


def _apply_mods(hexv, el):
    """套用 lumMod/lumOff/shade/tint 等子元素（近似换算）。"""
    if not hexv or len(hexv) != 6:
        return hexv
    try:
        r, g, b = (float(int(hexv[i:i + 2], 16)) for i in (0, 2, 4))
    except Exception:
        return hexv
    for ch in list(el):
        name = ch.tag.split('}')[-1]
        try:
            v = float(ch.get('val') or 100000) / 100000.0
        except Exception:
            continue
        if name in ('lumMod', 'shade'):
            r, g, b = r * v, g * v, b * v
        elif name == 'lumOff':
            r, g, b = r + 255 * v, g + 255 * v, b + 255 * v
        elif name == 'tint':
            r, g, b = r * v + 255 * (1 - v), g * v + 255 * (1 - v), b * v + 255 * (1 - v)
    cl = lambda x: max(0, min(255, int(round(x))))
    return '%02X%02X%02X' % (cl(r), cl(g), cl(b))


def _color_hex(fill_parent, scheme):
    """从 a:solidFill 容器解析颜色（支持主题色与亮度修正）。"""
    if fill_parent is None:
        return None
    sf = fill_parent.find(A + 'solidFill')
    if sf is None:
        return None
    for child in list(sf):
        tag = child.tag.split('}')[-1]
        if tag in ('srgbClr', 'sysClr'):
            return _apply_mods(_hex_of(child), child)
        if tag == 'schemeClr':
            base = (scheme or {}).get(child.get('val'))
            if base:
                return _apply_mods(base, child)
            return None
    return None


def _run_color(run, scheme):
    return _color_hex(_rpr(run), scheme)


def _fill_color(sh, scheme):
    try:
        sp_pr = sh._element.find(A + 'spPr')
        if sp_pr is None:
            sp_pr = sh._element.find(P + 'spPr')
        return _color_hex(sp_pr, scheme)
    except Exception:
        return None


def _rgb_near(hexv, target, tol=RGB_TOL):
    if not hexv or not target or len(hexv) != 6 or len(target) != 6:
        return False
    try:
        for i in (0, 2, 4):
            if abs(int(hexv[i:i + 2], 16) - int(target[i:i + 2], 16)) > tol:
                return False
    except Exception:
        return False
    return True


def _is_blue(hexv):
    """蓝色系判定：蓝通道明显强于红、绿通道。"""
    if not hexv or len(hexv) != 6:
        return False
    try:
        r, g, b = (int(hexv[i:i + 2], 16) for i in (0, 2, 4))
    except Exception:
        return False
    return b >= 96 and (b - r) >= 40 and (b - g) >= 18


def _scheme_colors(zf):
    """读取主题配色（accent1-6/dk/lt 等），用于解析 schemeClr。"""
    out = {}
    try:
        xml = zf.read('ppt/theme/theme1.xml').decode('utf8', 'ignore')
    except Exception:
        return out
    blk = re.search(r'<a:clrScheme\b.*?</a:clrScheme>', xml, re.S)
    if not blk:
        return out
    for m in re.finditer(
            r'<a:(\w+)>\s*<a:(?:srgbClr val="([0-9A-Fa-f]{6})"|sysClr[^>]*?lastClr="([0-9A-Fa-f]{6})")',
            blk.group()):
        out[m.group(1)] = (m.group(2) or m.group(3)).upper()
    out.setdefault('tx1', out.get('dk1', '000000'))
    out.setdefault('tx2', out.get('dk2', '000000'))
    out.setdefault('bg1', out.get('lt1', 'FFFFFF'))
    out.setdefault('bg2', out.get('lt2', 'FFFFFF'))
    return out


def _first_run_props(sh, scheme):
    """返回首个 run 的 (字号, 加粗, 颜色)。"""
    for run in _runs(sh):
        if not (run.text or '').strip():
            continue
        return _size_pt(run), _bold(run), _run_color(run, scheme)
    return None, None, None


def _check_anchors(slide, table, scheme):
    """按锚点表逐项校验文字/字号/字色/位置，返回问题列表。"""
    bad = []
    for key, exact, sz, bold, rgb, left, top, width in table:
        sh = _locate(slide, key, exact=exact)
        if sh is None:
            bad.append('缺少“%s”' % key[:12])
            continue
        r_sz, r_bold, r_rgb = _first_run_props(sh, scheme)
        if sz is not None and not _near(r_sz, sz, PT_TOL):
            bad.append('“%s”字号%s≠%.1f磅' % (key[:8], r_sz, sz))
        if bold is True and r_bold is not True:
            bad.append('“%s”未加粗' % key[:8])
        if bold is None and r_bold is True:
            bad.append('“%s”被改为加粗' % key[:8])
        if rgb is not None and not _rgb_near(r_rgb, rgb):
            bad.append('“%s”字色%s≠%s' % (key[:8], r_rgb, rgb))
        l, t, w, _h = _rect(sh)
        if left is not None and not _near(l, left):
            bad.append('“%s”左边界%.2f≠%.2fcm' % (key[:8], l, left))
        if top is not None and not _near(t, top):
            bad.append('“%s”上边界%.2f≠%.2fcm' % (key[:8], t, top))
        if width is not None and not _near(w, width):
            bad.append('“%s”宽度%.2f≠%.2fcm' % (key[:8], w, width))
    return bad


def _pictures(slide):
    out = []
    for sh in _iter_shapes(slide):
        if str(getattr(sh, 'shape_type', '')).startswith('PICTURE'):
            out.append(sh)
    return out


def _dhash(im, w=9, h=8):
    """行内相邻像素比较哈希：对整体明暗/对比度线性变化与缩放不敏感。"""
    g = im.convert('L').resize((w, h), Image.LANCZOS)
    px = bytearray(g.tobytes())
    bits = 0
    for y in range(h):
        for x in range(w - 1):
            bits = (bits << 1) | (1 if px[y * w + x] > px[y * w + x + 1] else 0)
    return bits


def _ham(a, b):
    return bin((a or 0) ^ (b or 0)).count('1')


def _img_stats(blob):
    """返回 (宽, 高, 灰度均值, 灰度标准差, dhash)。"""
    im = Image.open(io.BytesIO(blob))
    im.load()
    w, h = im.size
    g = im.convert('L')
    st = ImageStat.Stat(g)
    return w, h, st.mean[0], st.stddev[0], _dhash(im)


def _blip_fill_candidates(slide):
    """收集可作为整页背景的图片填充：页面背景填充或铺满页面的形状填充。"""
    out = []
    el = slide._element
    bg = el.find('.//' + P + 'bg')
    if bg is not None:
        bf = bg.find('.//' + A + 'blipFill')
        if bf is not None:
            out.append(('页面背景填充', bf, (0.0, 0.0, PAGE_W_CM, PAGE_H_CM)))
    for sh in _iter_shapes(slide):
        l, t, w, h = _rect(sh)
        if w < PAGE_W_CM * 0.9 or h < PAGE_H_CM * 0.9:
            continue
        bf = sh._element.find('.//' + A + 'blipFill')
        if bf is None:
            bf = sh._element.find('.//' + P + 'blipFill')
        if bf is not None:
            out.append(('图片填充形状', bf, (l, t, w, h)))
    return out


def _bg_info(slide):
    """解析第1/2页整页背景：图片指纹、透明度、裁切、铺满情况。"""
    info = {'found': False, 'kind': '', 'rect': None, 'amt': None, 'crop': (0, 0, 0, 0),
            'tile': False, 'px': None, 'match': False, 'transp': None, 'baked': False,
            'note': ''}
    best = None
    for kind, bf, rect in _blip_fill_candidates(slide):
        blip = bf.find(A + 'blip')
        rid = blip.get(R + 'embed') if blip is not None else None
        blob = None
        if rid:
            try:
                blob = slide.part.related_part(rid).blob
            except Exception:
                blob = None
        stats = None
        if blob:
            try:
                stats = _img_stats(blob)
            except Exception:
                stats = None
        ham = _ham(stats[4], SRC_IMG_DHASH) if stats else 64
        cand = (ham, kind, bf, blip, rect, stats)
        if best is None or ham < best[0]:
            best = cand
    if best is None:
        info['note'] = '未找到铺满页面的图片填充'
        return info
    ham, kind, bf, blip, rect, stats = best
    info['found'] = True
    info['kind'] = kind
    info['rect'] = rect
    info['match'] = ham <= DHASH_MAX_HAM
    info['ham'] = ham
    sr = bf.find(A + 'srcRect')
    if sr is not None:
        info['crop'] = tuple(int(sr.get(k) or 0) for k in ('l', 't', 'r', 'b'))
    info['tile'] = bf.find(A + 'tile') is not None
    amf = blip.find(A + 'alphaModFix') if blip is not None else None
    if amf is not None:
        try:
            info['amt'] = int(amf.get('amt') or 100000)
        except Exception:
            info['amt'] = None
    if stats:
        info['px'] = (stats[0], stats[1])
        info['mean'] = stats[2]
        info['std'] = stats[3]
    return info


def _bg_transparency(info):
    """得出背景透明度（0-1）与取得方式；无 alphaModFix 时按像素统计反解。"""
    if not info.get('found'):
        return None, '未找到背景图片', False
    if info.get('amt') is not None:
        opacity = info['amt'] / 100000.0
        return 1.0 - opacity, '透明度参数 alphaModFix amt=%d（不透明度%.0f%%）' % (
            info['amt'], opacity * 100), False
    if info.get('std') is None:
        return None, '无透明度参数且无法读取像素', False
    a_std = info['std'] / SRC_IMG_STD if SRC_IMG_STD else 1.0
    a_mean = (255.0 - info['mean']) / (255.0 - SRC_IMG_MEAN)
    opacity = max(0.0, min(1.0, (a_std + a_mean) / 2.0))
    baked = opacity <= 0.95
    return 1.0 - opacity, '无透明度参数，按像素反解不透明度≈%.0f%%（已烧入位图）' % (
        opacity * 100), baked


def _full_bleed(info):
    """背景是否满版且左右未被裁切。"""
    if not info.get('found'):
        return False, '未找到整页背景图片'
    l, t, w, h = info['rect']
    if info['kind'] != '页面背景填充':
        if not (l <= POS_TOL and t <= POS_TOL and w >= PAGE_W_CM - POS_TOL
                and h >= PAGE_H_CM - POS_TOL):
            return False, '背景仅覆盖 %.2f×%.2fcm（左上 %.2f,%.2f），未满版' % (w, h, l, t)
    if info['tile']:
        return False, '背景为重复平铺填充'
    cl, ct, cr, cb = info['crop']
    if max(cl, cr) > 1000 or max(ct, cb) > 1000:
        return False, '插画被裁切（srcRect 左%.1f%% 右%.1f%% 上%.1f%% 下%.1f%%）' % (
            cl / 1000.0, cr / 1000.0, ct / 1000.0, cb / 1000.0)
    return True, ''


def _transitions(zf):
    """返回带页面切换效果的幻灯片编号与切换参数。"""
    hits = []
    for name in sorted(zf.namelist()):
        m = re.match(r'ppt/slides/slide(\d+)\.xml$', name)
        if not m:
            continue
        xml = zf.read(name).decode('utf8', 'ignore')
        for t in re.findall(r'<p(?:14)?:transition[^>]*>', xml):
            spd = re.search(r'spd="(\w+)"', t)
            dur = re.search(r'dur="(\d+)"', t)
            hits.append((int(m.group(1)),
                         spd.group(1) if spd else '',
                         int(dur.group(1)) if dur else 0))
            break
    return hits


def _id3_cover_bytes(blob):
    """统计 ID3v2 标签中 APIC（专辑封面）帧的字节数。"""
    if len(blob) < 10 or blob[:3] != b'ID3':
        return 0, 0
    ver = blob[3]
    size = ((blob[6] & 0x7f) << 21) | ((blob[7] & 0x7f) << 14) | \
           ((blob[8] & 0x7f) << 7) | (blob[9] & 0x7f)
    end = min(len(blob), 10 + size)
    i = 10
    cover = 0
    while i + 10 <= end:
        fid = blob[i:i + 4]
        if not fid.strip(b'\x00'):
            break
        try:
            if ver >= 4:
                fs = ((blob[i + 4] & 0x7f) << 21) | ((blob[i + 5] & 0x7f) << 14) | \
                     ((blob[i + 6] & 0x7f) << 7) | (blob[i + 7] & 0x7f)
            else:
                fs = struct.unpack('>I', blob[i + 4:i + 8])[0]
        except Exception:
            break
        if fs <= 0:
            break
        if fid == b'APIC':
            cover += fs
        i += 10 + fs
    if cover == 0 and size > COVER_MAX_BYTES:
        cover = size
    return cover, 10 + size


def _mp3_meta(blob):
    """遍历 MPEG 帧累计时长，返回 (声道, 采样率, 时长)。"""
    br_tab = [0, 32, 40, 48, 56, 64, 80, 96, 112, 128, 160, 192, 224, 256, 320, 0]
    sr_tab = [44100, 48000, 32000]
    _cover, i = _id3_cover_bytes(blob)
    ch = rate = None
    dur = 0.0
    n = len(blob)
    while i < n - 4:
        if blob[i] == 0xFF and (blob[i + 1] & 0xE0) == 0xE0:
            ver = (blob[i + 1] >> 3) & 3
            layer = (blob[i + 1] >> 1) & 3
            bri = (blob[i + 2] >> 4) & 0xF
            sri = (blob[i + 2] >> 2) & 3
            pad = (blob[i + 2] >> 1) & 1
            mode = (blob[i + 3] >> 6) & 3
            if layer != 1 or bri in (0, 15) or sri == 3:
                i += 1
                continue
            samplerate = sr_tab[sri] // (2 if ver == 2 else (4 if ver == 0 else 1))
            flen = (144 * br_tab[bri] * 1000) // samplerate + pad
            if flen <= 0:
                i += 1
                continue
            rate = samplerate
            ch = 1 if mode == 3 else 2
            dur += 1152.0 / samplerate
            i += flen
        else:
            i += 1
    return ch, rate, (dur if dur > 0 else None)


def _audio_meta(blob):
    """解析音频：格式、声道、采样率、时长（秒）、内嵌封面字节数、sha1。"""
    meta = {'fmt': '未知', 'ch': None, 'rate': None, 'dur': None, 'cover': 0,
            'bytes': len(blob or b''), 'sha1': ''}
    if not blob:
        return meta
    meta['sha1'] = hashlib.sha1(blob).hexdigest()
    if blob[:4] == b'RIFF' and blob[8:12] == b'WAVE':
        meta['fmt'] = 'wav'
        j = 12
        byte_rate = None
        while j + 8 <= len(blob):
            cid = blob[j:j + 4]
            try:
                cs = struct.unpack('<I', blob[j + 4:j + 8])[0]
            except Exception:
                break
            if cid == b'fmt ' and j + 24 <= len(blob):
                _f, ch, rate, byte_rate, _ba, _bits = struct.unpack(
                    '<HHIIHH', blob[j + 8:j + 24])
                meta['ch'] = ch
                meta['rate'] = rate
            elif cid == b'data' and byte_rate:
                meta['dur'] = cs / float(byte_rate)
                break
            j += 8 + cs + (cs & 1)
        return meta
    if blob[:3] == b'ID3' or (blob[0] == 0xFF and (blob[1] & 0xE0) == 0xE0):
        meta['fmt'] = 'mp3'
        meta['cover'] = _id3_cover_bytes(blob)[0]
        meta['ch'], meta['rate'], meta['dur'] = _mp3_meta(blob)
        return meta
    if len(blob) > 12 and blob[4:8] == b'ftyp':
        meta['fmt'] = 'mp4/m4a'
        k = blob.find(b'mvhd')
        if k > 0:
            try:
                ver = blob[k + 4]
                if ver == 1:
                    ts = struct.unpack('>I', blob[k + 20:k + 24])[0]
                    du = struct.unpack('>Q', blob[k + 24:k + 32])[0]
                else:
                    ts = struct.unpack('>I', blob[k + 12:k + 16])[0]
                    du = struct.unpack('>I', blob[k + 16:k + 20])[0]
                if ts:
                    meta['dur'] = du / float(ts)
            except Exception:
                pass
        c = blob.find(b'covr')
        if c > 0:
            try:
                meta['cover'] = struct.unpack('>I', blob[c - 4:c])[0]
            except Exception:
                pass
        return meta
    return meta


def _audio_bindings(slide):
    """找出承载音频的形状：(shape, rId, 绑定方式)。"""
    out = []
    for sh in _iter_shapes(slide):
        el = sh._element
        for af in el.iter(A + 'audioFile'):
            rid = af.get(R + 'link') or af.get(R + 'embed')
            out.append((sh, rid, 'media'))
        for hl in el.iter(A + 'hlinkClick'):
            snd = hl.find(A + 'snd')
            if snd is not None:
                out.append((sh, snd.get(R + 'embed') or snd.get(R + 'link'), 'snd'))
    return out


def _rel_info(slide, rid):
    """返回 (部件名, 字节内容, 是否外部链接)。"""
    if not rid:
        return None, None, False
    try:
        rel = slide.part.rels[rid]
    except Exception:
        return None, None, False
    try:
        if rel.is_external:
            return rel.target_ref, None, True
        part = rel.target_part
        return str(part.partname), part.blob, False
    except Exception:
        return None, None, False


def _is_clickable(slide, sh, kind):
    """形状是否绑定了点击即播放的动作。"""
    if kind == 'snd':
        return True
    for hl in sh._element.iter(A + 'hlinkClick'):
        if 'media' in (hl.get('action') or ''):
            return True
    xml = slide._element
    for node in xml.iter():
        if node.tag.split('}')[-1] in ('cMediaNode', 'audio', 'video'):
            return True
    return False


def _button_cluster(slide, seed):
    """把与播放对象相互压盖的无文字形状合并为一个按钮整体。"""
    if seed is None:
        return None, []
    cur = _rect(seed)
    picked = [seed]
    cands = []
    for sh in _iter_shapes(slide):
        if sh is seed:
            continue
        l, t, w, h = _rect(sh)
        if w >= PAGE_W_CM * 0.9 or h >= PAGE_H_CM * 0.9:
            continue
        if _norm(_shape_text(sh)):
            continue
        if w <= 0 or h <= 0:
            continue
        cands.append(sh)
    changed = True
    while changed:
        changed = False
        for sh in cands:
            if sh in picked:
                continue
            if _inter_area(cur, _rect(sh)) > 0.02:
                picked.append(sh)
                cur = _union_rect(cur, _rect(sh))
                changed = True
    return cur, picked


def _package_issues(zf):
    """检查随文件保存的主题回退表、默认文本样式、视图、标记部件与文档属性。"""
    issues = []
    names = zf.namelist()
    try:
        theme = zf.read('ppt/theme/theme1.xml')
    except Exception:
        theme = b''
    scripts = len(re.findall(rb'<a:font script=', theme))
    if scripts < 20 or len(theme) < 5000:
        issues.append('主题多语种字体映射回退表被精简（script 映射%d条、主题%d字节）'
                      % (scripts, len(theme)))
    try:
        pres = zf.read('ppt/presentation.xml').decode('utf8', 'ignore')
    except Exception:
        pres = ''
    lvls = len(re.findall(r'<a:lvl\dpPr', pres))
    if 'defaultTextStyle' not in pres or lvls < 9:
        issues.append('演示文稿级默认文本样式缺失（9级默认段落仅剩%d级）' % lvls)
    if 'ppt/viewProps.xml' not in names:
        issues.append('视图设置部件 ppt/viewProps.xml 缺失')
    tags = [n for n in names if re.match(r'ppt/tags/tag\d+\.xml$', n)]
    if len(tags) < 28:
        issues.append('自定义标记数据部件仅剩%d个（应不少于28个）' % len(tags))
    refs = 0
    for n in names:
        if re.match(r'ppt/slides/slide\d+\.xml$', n):
            refs += len(re.findall(rb'<p:custDataLst>', zf.read(n)))
    if refs < 28:
        issues.append('形状上的标记引用仅剩%d处（应不少于28处）' % refs)
    try:
        app = zf.read('docProps/app.xml')
    except Exception:
        app = b''
    if len(app) < 1200 or b'TitlesOfParts' not in app:
        issues.append('文档应用程序属性被精简至%d字节' % len(app))
    try:
        core = zf.read('docProps/core.xml').decode('utf8', 'ignore')
    except Exception:
        core = ''
    if 'lastModifiedBy' not in core or 'dc:creator' not in core:
        issues.append('文档属性中的作者/最后修改者信息缺失')
    return issues


def _open_presentation(file_path):
    try:
        return Presentation(file_path), ''
    except Exception as exc:
        return None, '%s: %s' % (type(exc).__name__, exc)


# ---------------------------------------------------------------- 维度1


def _check_dim1(prs, zf):
    """维度1：可用与可修改性。返回 (pass, reason)。"""
    slides = list(prs.slides)
    if len(slides) < 6:
        return False, '全文共%d页，不满足6页（页数缺失）' % len(slides)
    for idx, slide in enumerate(slides, start=1):
        tshapes = _text_shapes(slide)
        pics = _pictures(slide)
        if not tshapes:
            return False, '第%d页无任何可点选编辑的文字，疑似整页被替换为位图截图' % idx
        for pic in pics:
            l, t, w, h = _rect(pic)
            if w >= PAGE_W_CM * 0.95 and h >= PAGE_H_CM * 0.95 and len(tshapes) <= 1:
                return False, '第%d页整体为一张满版图片且文字不可编辑' % idx
    for idx in (1, 2):
        slide = slides[idx - 1]
        if len(_text_shapes(slide)) < 3:
            return False, '第%d页可单独点选编辑的文本元素不足3个，疑似整页被粘贴为图片' % idx
        if not _bg_info(slide).get('found'):
            return False, '第%d页背景不是可再调整的图片填充或页面背景填充' % idx
    page_area = PAGE_W_CM * PAGE_H_CM
    for idx, slide in enumerate(slides, start=1):
        rects = []
        for sh in _text_shapes(slide):
            l, t, w, h = _rect(sh)
            if w <= 0 or h <= 0:
                continue
            if l + w <= 0.2 or t + h <= 0.2 or l >= PAGE_W_CM - 0.2 or t >= PAGE_H_CM - 0.2:
                return False, '第%d页存在文字元素整体溢出页面边界被截断（左%.2f 上%.2f）' % (idx, l, t)
            rects.append((l, t, w, h))
        overlap = 0.0
        for i in range(len(rects)):
            for j in range(i + 1, len(rects)):
                overlap += _inter_area(rects[i], rects[j])
        if overlap > page_area * 0.25:
            return False, '第%d页文字大面积重叠压盖（重叠面积%.1fcm²）' % (idx, overlap)
    return True, ''


# ---------------------------------------------------------------- 维度2


def _pct(x):
    return '无法判定' if x is None else '%.0f%%' % (x * 100.0)


def _cmf(x):
    return '无' if x is None else '%.2fcm' % x


def _evaluate_impl(file_path, result):
    prs, err = _open_presentation(file_path)
    if prs is None:
        # err 形如 "InvalidXmlError: ..." / "KeyError: There is no item named ..."：
        # 这些都是交付文件自身结构损坏（缺部件、XML 非法、OPC 关系不完整），
        # 属于维度一不通过，不是脚本运行错误。
        if any(tag in err for tag in (
                'InvalidXmlError', 'There is no item named',
                'no relationship of type', 'BadZipFile', 'XMLSyntaxError',
                'ParseError')):
            result['status'] = 'ok'
            result['error'] = None
            result['dim1_pass'] = False
            result['dim1_reason'] = '交付文件结构不完整，无法正常解析：%s' % err
            result['dim2_items'] = []
            result['total_score'] = 0
            return result
        result['status'] = 'error'
        result['error'] = err
        result['dim1_pass'] = False
        result['dim1_reason'] = '文件损坏或无法正常打开（%s）' % err
        result['dim2_items'] = []
        result['total_score'] = 0
        return result

    zf = zipfile.ZipFile(file_path)
    try:
        dim1_pass, dim1_reason = _check_dim1(prs, zf)
        result['dim1_pass'] = dim1_pass
        result['dim1_reason'] = dim1_reason
        if not dim1_pass:
            result['dim2_items'] = []
            result['total_score'] = 0
            return result

        scheme = _scheme_colors(zf)
        slides = list(prs.slides)
        s1, s2 = slides[0], slides[1]
        s6 = slides[5] if len(slides) >= 6 else slides[-1]
        results = []

        # ---- 01/02 第1、2页背景插画与透明度
        bg1 = _bg_info(s1)
        bg2 = _bg_info(s2)
        tr1, note1, baked1 = _bg_transparency(bg1)
        tr2, note2, baked2 = _bg_transparency(bg2)
        for idx, bg, tr, note, lo, hi, score in (
                (1, bg1, tr1, note1, 0.45, 0.55, 3),
                (2, bg2, tr2, note2, 0.80, 0.90, 3)):
            rule = ('第%d页整页背景替换为所给幼儿园卡通插画（米黄底色、4名幼儿与积木/玩具/树木、'
                    '宽高比约1.78:1），且以约%d%%透明度（%d%%-%d%%）呈现'
                    % (idx, round((lo + hi) * 50), lo * 100, hi * 100))
            if not bg.get('found'):
                results.append('+0：%s（%s）' % (rule, bg.get('note') or '未找到整页背景图片'))
            elif not bg.get('match'):
                results.append('+0：%s（背景图片与所给插画不一致，指纹汉明距离%d）'
                               % (rule, bg.get('ham', 64)))
            elif tr is None or not (lo <= tr <= hi):
                results.append('+0：%s（%s，实测透明度%s）' % (rule, note, _pct(tr)))
            else:
                results.append('+%d：%s' % (score, rule))

        # ---- 03 背景满版且左右未裁切
        rule = '第1、2页背景铺满整个页面（33.87cm×19.05cm满版、不重复平铺），且插画左右两侧未被裁切'
        ok_a, why_a = _full_bleed(bg1)
        ok_b, why_b = _full_bleed(bg2)
        if ok_a and ok_b:
            results.append('+1：%s' % rule)
        else:
            why = '；'.join(x for x in ('第1页%s' % why_a if why_a else '',
                                       '第2页%s' % why_b if why_b else '') if x)
            results.append('+0：%s（%s）' % (rule, why))

        # ---- 04 第2页红色标题 Arial 加粗
        rule = '第2页红色标题“整体教学设计：从学习基础到实践改进”字体设为Arial且加粗'
        title2 = _locate(s2, TITLE_S2, exact=True)
        if title2 is None:
            results.append('+0：%s（未找到该标题文本框）' % rule)
        else:
            t_runs = [r for r in _runs(title2) if (r.text or '').strip()]
            fonts = set()
            for r in t_runs:
                fonts.update(_typefaces(r))
            has_arial = any('Arial' in (f or '') for f in fonts)
            all_bold = bool(t_runs) and all(_bold(r) is True for r in t_runs)
            if has_arial and all_bold:
                results.append('+1：%s' % rule)
            else:
                results.append('+0：%s（字体%s，%s）'
                               % (rule, '/'.join(sorted(fonts)) or '未设置',
                                  '已加粗' if all_bold else '未加粗'))

        # ---- 05 六个小标题蓝色＋下划线
        rule = ('“学习诊断”“任务转化”“资源整合”“情境演练”“表现评价”“迭代成长”'
                '六个小标题文字颜色改为蓝色系并全部添加下划线')
        labels = [(k, _locate(s2, k, exact=True)) for k in SUB_LABELS]
        bad = []
        for key, sh in labels:
            if sh is None:
                bad.append('缺少“%s”' % key)
                continue
            rs = [r for r in _runs(sh) if (r.text or '').strip()]
            cols = [_run_color(r, scheme) for r in rs]
            if not rs or not all(_is_blue(c) for c in cols):
                bad.append('“%s”字色%s非蓝色系' % (key, '/'.join(c or '默认' for c in cols)))
            if not rs or not all(_underlined(r) for r in rs):
                bad.append('“%s”未加下划线' % key)
        if bad:
            results.append('+0：%s（%s）' % (rule, '；'.join(bad[:4])))
        else:
            results.append('+3：%s' % rule)

        # ---- 06 六个小标题竖排＋色块改竖长方形
        rule = ('上述六个小标题全部由横排改为竖排显示（文本方向eaVert），且承载它们的色块'
                '由宽3.17cm×高1.65cm的横长方形改为宽1.5-2.2cm、高3.0-4.0cm的竖长方形')
        bad = []
        for key, sh in labels:
            if sh is None:
                bad.append('缺少“%s”' % key)
                continue
            vert = None
            try:
                bp = sh.text_frame._txBody.find(A + 'bodyPr')
                vert = bp.get('vert') if bp is not None else None
            except Exception:
                vert = None
            if vert not in ('eaVert', 'vert', 'vert270', 'mongolianVert'):
                bad.append('“%s”文本方向为%s（未改竖排）' % (key, vert or '横排'))
            _l, _t, w, h = _rect(sh)
            if not (1.5 - POS_TOL <= w <= 2.2 + POS_TOL and 3.0 - POS_TOL <= h <= 4.0 + POS_TOL):
                bad.append('“%s”色块%.2f×%.2fcm不在竖长方形区间' % (key, w, h))
                continue
            sz = _size_pt(_runs(sh)[0]) if _runs(sh) else 12.0
            need_h = 4 * (sz or 12.0) * 0.03528
            if h < need_h or w < (sz or 12.0) * 0.03528 + 0.05:
                bad.append('“%s”四字竖排后溢出色块（需高≥%.2fcm）' % (key, need_h))
        if bad:
            results.append('+0：%s（%s）' % (rule, '；'.join(bad[:4])))
        else:
            results.append('+5：%s' % rule)

        # ---- 07 三栏并排
        rule = ('“学生”“任务”“资源”三组标题及其对应内容由自上而下逐行堆叠改为横向并排三栏，'
                '三个内容块尺寸一致、顶端对齐，三段内容分别落在第一、二、三栏内')
        tshapes = [_locate(s2, k, exact=True) for k in COL_TITLES]
        cshapes = [_locate(s2, t, exact=False) for t in COL_TEXTS]
        bad = []
        for i, sh in enumerate(cshapes):
            if sh is None:
                bad.append('未找到第%d栏内容文字' % (i + 1))
        for i, sh in enumerate(tshapes):
            if sh is None:
                bad.append('未找到“%s”标题' % COL_TITLES[i])
        if not bad:
            cr = [_rect(sh) for sh in cshapes]
            tr = [_rect(sh) for sh in tshapes]
            if max(r[2] for r in cr) - min(r[2] for r in cr) > POS_TOL or \
                    max(r[3] for r in cr) - min(r[3] for r in cr) > POS_TOL:
                bad.append('三个内容块尺寸不一致（%s）'
                           % ' / '.join('%.2f×%.2f' % (r[2], r[3]) for r in cr))
            if max(r[1] for r in cr) - min(r[1] for r in cr) > POS_TOL:
                bad.append('三个内容块顶端未对齐（上边界%s）'
                           % ' / '.join('%.2f' % r[1] for r in cr))
            for i in range(2):
                if cr[i + 1][0] <= cr[i][0] or cr[i + 1][0] + POS_TOL < cr[i][0] + cr[i][2]:
                    bad.append('第%d、%d栏未横向并排（左边界%.2f→%.2f）'
                               % (i + 1, i + 2, cr[i][0], cr[i + 1][0]))
        if bad:
            results.append('+0：%s（%s）' % (rule, '；'.join(bad[:4])))
        else:
            results.append('+5：%s' % rule)

        cols_ok = all(x is not None for x in cshapes) and all(x is not None for x in tshapes)
        cr = [_rect(sh) for sh in cshapes] if cols_ok else None
        trr = [_rect(sh) for sh in tshapes] if cols_ok else None
        # 后续“间距相等/左右均衡”等项以三栏确实横向并排为前提
        side_by_side = bool(cols_ok) and all(
            cr[i + 1][0] > cr[i][0] and cr[i + 1][0] + POS_TOL >= cr[i][0] + cr[i][2]
            for i in range(2))

        # ---- 08 三栏水平间距相等
        rule = '三栏水平间距相等（相邻两栏空隙差值≤0.3cm，或三栏左边界等差、步长差≤0.3cm）'
        if not side_by_side:
            results.append('+0：%s（三栏结构不完整或未横向并排，无法判定）' % rule)
        else:
            g1 = cr[1][0] - (cr[0][0] + cr[0][2])
            g2 = cr[2][0] - (cr[1][0] + cr[1][2])
            p1 = cr[1][0] - cr[0][0]
            p2 = cr[2][0] - cr[1][0]
            if (abs(g1 - g2) <= POS_TOL and min(g1, g2) >= -POS_TOL) or abs(p1 - p2) <= POS_TOL:
                results.append('+3：%s' % rule)
            else:
                results.append('+0：%s（空隙%.2fcm与%.2fcm相差%.2fcm）'
                               % (rule, g1, g2, abs(g1 - g2)))

        # ---- 09 标题位于内容块正上方且水平居中
        rule = ('“学生”“任务”“资源”三个标题分别位于各自内容块正上方且水平居中'
                '（水平中心偏差≤0.5cm、标题底边与内容块顶边间距差值≤0.3cm）')
        if not cols_ok:
            results.append('+0：%s（三栏结构不完整，无法判定）' % rule)
        else:
            bad = []
            gaps = []
            for i in range(3):
                dc = abs((trr[i][0] + trr[i][2] / 2.0) - (cr[i][0] + cr[i][2] / 2.0))
                if dc > 0.5:
                    bad.append('“%s”水平中心偏差%.2fcm' % (COL_TITLES[i], dc))
                gap = cr[i][1] - (trr[i][1] + trr[i][3])
                gaps.append(gap)
                if gap < -POS_TOL:
                    bad.append('“%s”未位于内容块上方' % COL_TITLES[i])
            if max(gaps) - min(gaps) > POS_TOL:
                bad.append('标题与内容块间距不一致（%s）' % ' / '.join('%.2f' % g for g in gaps))
            if bad:
                results.append('+0：%s（%s）' % (rule, '；'.join(bad[:3])))
            else:
                results.append('+3：%s' % rule)

        # ---- 10 三栏内容自动换行完整显示
        rule = ('三栏内容文字在各自块内自动换行完整显示，不出现文字堆叠、跨栏重叠、'
                '压住上方标题或溢出块外被截断（每个内容块高度≥3cm）')
        if not cols_ok:
            results.append('+0：%s（三栏结构不完整，无法判定）' % rule)
        else:
            bad = []
            for i in range(3):
                _l, _t, w, h = cr[i]
                if h < 3.0 - POS_TOL:
                    bad.append('第%d栏内容块高度仅%.2fcm' % (i + 1, h))
                rs = _runs(cshapes[i])
                sz = _size_pt(rs[0]) if rs else 11.6
                cw = (sz or 11.6) * 0.03528
                cpl = max(1, int((w - 0.2) / cw))
                lines = int(len(_norm(_shape_text(cshapes[i]))) / cpl) + 1
                need = lines * cw * 1.25
                if need > h + POS_TOL:
                    bad.append('第%d栏%d磅文字需%.2fcm高、块高仅%.2fcm' % (i + 1, sz or 0, need, h))
            allr = [(COL_TITLES[i], trr[i]) for i in range(3)] + \
                   [('第%d栏内容' % (i + 1), cr[i]) for i in range(3)]
            for i in range(len(allr)):
                for j in range(i + 1, len(allr)):
                    ov = _inter_area(allr[i][1], allr[j][1])
                    if ov > 0.1:
                        bad.append('“%s”与“%s”重叠%.2fcm²' % (allr[i][0], allr[j][0], ov))
            if bad:
                results.append('+0：%s（%s）' % (rule, '；'.join(bad[:3])))
            else:
                results.append('+3：%s' % rule)

        # ---- 11 删除“评价”标题及其内容
        rule = ('“评价”标题及其对应内容“关注安全意识、幼儿回应、组织节奏、反思修正四类表现”'
                '已删除，第2页不再出现这两处文字')
        left_title = _locate(s2, DROPPED_TITLE, exact=True)
        s2_text = ''.join(_norm(_shape_text(sh)) for sh in _iter_shapes(s2))
        left_text = _norm(DROPPED_TEXT) in s2_text
        if left_title is None and not left_text:
            results.append('+3：%s' % rule)
        else:
            why = []
            if left_title is not None:
                why.append('“评价”标签仍在')
            if left_text:
                why.append('“关注安全意识…四类表现”仍在')
            results.append('+0：%s（%s）' % (rule, '、'.join(why)))

        # ---- 12 三栏整体左右均衡
        rule = '三栏整体在页面上左右均衡（最左栏左侧留白与最右栏右侧留白之差≤2cm）'
        if not side_by_side:
            results.append('+0：%s（三栏结构不完整或未横向并排，无法判定）' % rule)
        else:
            ml = min(r[0] for r in cr)
            mr = PAGE_W_CM - max(r[0] + r[2] for r in cr)
            if abs(ml - mr) <= 2.0:
                results.append('+1：%s' % rule)
            else:
                results.append('+0：%s（左留白%.2fcm、右留白%.2fcm，相差%.2fcm）'
                               % (rule, ml, mr, abs(ml - mr)))

        # ---- 13~16 第6页音频与播放按钮
        slast = slides[-1]
        binds = _audio_bindings(slast)
        audio_shape = audio_kind = None
        audio_meta = None
        audio_name = ''
        audio_external = False
        for sh, rid, kind in binds:
            name, blob, ext = _rel_info(slast, rid)
            if ext:
                audio_external = True
            if blob and audio_meta is None:
                audio_shape, audio_kind = sh, kind
                audio_name = name or ''
                audio_meta = _audio_meta(blob)
        if audio_shape is None and binds:
            audio_shape, _rid0, audio_kind = binds[0]
        gate = '音频仅为指向本机磁盘路径的外部链接，第三项任务相关加分项不得分'

        rule = ('第6页（最后一页）的音频内容为所给钢琴曲：时长约36秒（34-38秒）、'
                '双声道48kHz的钢琴独奏，与所给音乐一致')
        if audio_external:
            results.append('+0：%s（%s）' % (rule, gate))
        elif audio_meta is None:
            results.append('+0：%s（最后一页未发现随文件嵌入的音频）' % rule)
        else:
            dur = audio_meta['dur']
            same = audio_meta['sha1'] == SRC_WAV_SHA1
            fit = (dur is not None and AUDIO_DUR_RANGE[0] <= dur <= AUDIO_DUR_RANGE[1]
                   and audio_meta['ch'] == AUDIO_CH and audio_meta['rate'] == AUDIO_RATE)
            if same or fit:
                results.append('+3：%s' % rule)
            else:
                results.append('+0：%s（实测%s：时长%s、%s声道、%sHz）'
                               % (rule, audio_meta['fmt'],
                                  '未知' if dur is None else '%.1f秒' % dur,
                                  audio_meta['ch'], audio_meta['rate']))

        rule = ('第6页设有可点击播放该音频的按钮或图标（形状/图标上绑定播放动作或媒体触发，'
                '放映时点击即出声）')
        if audio_external:
            results.append('+0：%s（%s）' % (rule, gate))
        elif audio_shape is None:
            results.append('+0：%s（页面上未发现与音频绑定的形状或图标）' % rule)
        elif not _is_clickable(slast, audio_shape, audio_kind):
            results.append('+0：%s（音频对象未绑定点击播放动作）' % rule)
        else:
            results.append('+3：%s' % rule)

        btn_rect, btn_parts = _button_cluster(slast, audio_shape)

        rule = ('播放按钮位于第6页左下角区域（按钮左边界距页面左侧≤4cm、'
                '上边界距页面顶端≥15cm）')
        if audio_external:
            results.append('+0：%s（%s）' % (rule, gate))
        elif btn_rect is None:
            results.append('+0：%s（未找到播放按钮）' % rule)
        elif btn_rect[0] <= 4.0 + POS_TOL and btn_rect[1] >= 15.0 - POS_TOL:
            results.append('+3：%s' % rule)
        else:
            results.append('+0：%s（按钮左边界距页左%.2fcm、上边界距页顶%.2fcm）'
                           % (rule, btn_rect[0], btn_rect[1]))

        rule = ('播放按钮尺寸适中（宽高均在1.0-2.5cm之间），且未压盖第6页页脚文字'
                '“幼儿游戏活动课程设计与实施  ·  6”')
        foot = _locate(slast, FOOTER_S6, exact=True)
        foot_rect = _rect(foot) if foot is not None else (1.57, 18.11, 12.19, 0.56)
        if audio_external:
            results.append('+0：%s（%s）' % (rule, gate))
        elif btn_rect is None:
            results.append('+0：%s（未找到播放按钮）' % rule)
        else:
            bw, bh = btn_rect[2], btn_rect[3]
            cover = _inter_area(btn_rect, foot_rect)
            if not (1.0 - POS_TOL <= bw <= 2.5 + POS_TOL and
                    1.0 - POS_TOL <= bh <= 2.5 + POS_TOL):
                results.append('+0：%s（按钮实测%.2f×%.2fcm）' % (rule, bw, bh))
            elif cover > 0.02:
                results.append('+0：%s（按钮与页脚文本框重叠%.2f平方厘米）' % (rule, cover))
            else:
                results.append('+1：%s' % rule)

        # ---- 17 音频不应额外内嵌专辑封面图
        rule = ('第6页嵌入的音频文件仅含音频数据，未额外内嵌约1.1MB的900×900专辑封面图'
                '使文件体积无谓增大')
        cov = (audio_meta or {}).get('cover') or 0
        if audio_meta is None:
            results.append('-0：%s' % rule)
        elif cov > COVER_MAX_BYTES:
            results.append('-1：%s（音频内嵌封面图%.2fMB）' % (rule, cov / 1048576.0))
        else:
            results.append('-0：%s' % rule)

        # ---- 18 左下角只保留一个播放按钮
        rule = '第6页左下角只保留一个播放按钮，未在其上再叠放第二个图标造成两个按钮相互压盖'
        bodies = []
        for sh in (btn_parts or []):
            r = _rect(sh)
            if min(r[2], r[3]) >= 0.8:
                bodies.append((sh, r))
        stacked = []
        for i in range(len(bodies)):
            for j in range(i + 1, len(bodies)):
                if _inter_area(bodies[i][1], bodies[j][1]) > 0.05:
                    stacked.append('%.2f×%.2fcm 与 %.2f×%.2fcm 重叠'
                                   % (bodies[i][1][2], bodies[i][1][3],
                                      bodies[j][1][2], bodies[j][1][3]))
        if stacked:
            results.append('-3：%s（左下角有%d个按钮主体：%s）'
                           % (rule, len(bodies), '；'.join(stacked[:2])))
        else:
            results.append('-0：%s' % rule)

        # ---- 19 无页面切换动画
        rule = '第1至第6页无页面切换动画，未给各页统一添加“慢速、时长2秒”的页面切换效果'
        trans = _transitions(zf)
        if trans:
            spd, dur = trans[0][1], trans[0][2]
            desc = '，'.join(x for x in (('速度%s' % spd) if spd else '',
                                        ('时长%.1f秒' % (dur / 1000.0)) if dur else '') if x)
            results.append('-3：%s（第%s页存在切换效果%s）'
                           % (rule, '、'.join(str(t[0]) for t in trans[:6]),
                              '：%s' % desc if desc else ''))
        else:
            results.append('-0：%s' % rule)

        # ---- 20 背景以图片＋透明度实现且未放大重采样
        rule = ('第1、2页背景以图片配合透明度设置实现、透明度可再次调整，未把淡化效果烧进'
                '位图像素，也未把960×539的插画重采样放大后嵌入')
        why = []
        for idx, bg, baked in ((1, bg1, baked1), (2, bg2, baked2)):
            if not bg.get('found') or not bg.get('match'):
                continue  # 背景并非所给插画，由 01/02 项判定，不在此重复扣分
            if baked:
                why.append('第%d页透明度已烧入位图（无alphaModFix参数）' % idx)
            px = bg.get('px')
            if px and px[0] >= SRC_IMG_W * 1.5:
                why.append('第%d页背景图被放大为%d×%d像素' % (idx, px[0], px[1]))
        if why:
            results.append('-1：%s（%s）' % (rule, '；'.join(why[:3])))
        else:
            results.append('-0：%s' % rule)

        # ---- 21 三个栏标题标签样式保持
        rule = ('第2页“学生”“任务”“资源”三个标题保持字号11.5磅、加粗、白色FFFFFF文字的'
                '小标签（宽2.84cm×高1.22cm），未放大字号也未拉宽成与内容块同宽的整条色带')
        why = []
        for i, sh in enumerate(tshapes):
            if sh is None:
                continue
            sz, bd, col = _first_run_props(sh, scheme)
            r = _rect(sh)
            if sz is not None and abs(sz - 11.5) > PT_TOL:
                why.append('“%s”字号%.1f磅' % (COL_TITLES[i], sz))
            if bd is False:
                why.append('“%s”未加粗' % COL_TITLES[i])
            if col and not _rgb_near(col, 'FFFFFF'):
                why.append('“%s”字色%s' % (COL_TITLES[i], col))
            if not _near(r[2], 2.84, 0.5):
                why.append('“%s”标签宽%.2fcm' % (COL_TITLES[i], r[2]))
        if why:
            results.append('-1：%s（%s）' % (rule, '；'.join(why[:4])))
        else:
            results.append('-0：%s' % rule)

        # ---- 22 三栏内容文字样式保持
        rule = '第2页三栏内容文字保持字号11.6磅、字色5D6470，未放大字号也未改色'
        why = []
        for i, sh in enumerate(cshapes):
            if sh is None:
                continue
            sz, _bd, col = _first_run_props(sh, scheme)
            if sz is not None and abs(sz - 11.6) > PT_TOL:
                why.append('第%d栏字号%.1f磅' % (i + 1, sz))
            if col and not _rgb_near(col, '5D6470'):
                why.append('第%d栏字色%s' % (i + 1, col))
        if why:
            results.append('-1：%s（%s）' % (rule, '；'.join(why[:4])))
        else:
            results.append('-0：%s' % rule)

        # ---- 23 底部“核心主线”文本框样式与位置保持
        rule = ('第2页底部“核心主线：让学生在“看见儿童—设计游戏—组织活动—观察反馈”的循环中'
                '形成专业判断。”保持字号12.5磅、加粗、字色1F3A5F、底色F0FBF8配2F9B98边框，'
                '且文本框宽28.45cm、左边界距页左2.34cm、上边界距页顶16.0-16.4cm')
        main = _locate(s2, '核心主线', exact=False)
        why = []
        if main is None:
            why.append('未找到该文本框')
        else:
            sz, bd, col = _first_run_props(main, scheme)
            l, t, w, _h = _rect(main)
            if sz is not None and abs(sz - 12.5) > PT_TOL:
                why.append('字号%.1f磅' % sz)
            if bd is False:
                why.append('未加粗')
            if col and not _rgb_near(col, '1F3A5F'):
                why.append('字色%s' % col)
            fillc = _fill_color(main, scheme)
            if fillc and not _rgb_near(fillc, 'F0FBF8', FILL_TOL):
                why.append('底色%s' % fillc)
            linec = None
            try:
                sp_pr = main._element.find(A + 'spPr')
                linec = _color_hex(sp_pr.find(A + 'ln'), scheme) if sp_pr is not None else None
            except Exception:
                linec = None
            if linec and not _rgb_near(linec, '2F9B98', FILL_TOL):
                why.append('边框色%s' % linec)
            if not _near(w, 28.45):
                why.append('宽%.2fcm' % w)
            if not _near(l, 2.34):
                why.append('左边界%.2fcm' % l)
            if not (16.0 - POS_TOL <= t <= 16.4 + POS_TOL):
                why.append('上边界距页顶%.2fcm' % t)
        if why:
            results.append('-1：%s（%s）' % (rule, '；'.join(why[:4])))
        else:
            results.append('-0：%s' % rule)

        # ---- 24 六个圆形数字标号字体保持微软雅黑
        rule = ('第2页六个圆形数字标号“1”“2”“3”“4”“5”“6”（各0.97cm×0.97cm，白色12磅加粗）'
                '的数字字体保持微软雅黑，未改为Arial')
        badges = []
        for sh in _text_shapes(s2):
            txt = _norm(_shape_text(sh))
            if txt in ('1', '2', '3', '4', '5', '6'):
                r = _rect(sh)
                if _near(r[2], 0.97, 0.4) and _near(r[3], 0.97, 0.4):
                    badges.append((txt, sh))
        why = []
        for txt, sh in badges:
            rs = [r for r in _runs(sh) if (r.text or '').strip()]
            for r in rs:
                rpr = _rpr(r)
                el = rpr.find(A + 'latin') if rpr is not None else None
                face = el.get('typeface') if el is not None else None
                if face and face != '微软雅黑':
                    why.append('“%s”数字字体为%s' % (txt, face))
                break
        if why:
            results.append('-1：%s（%s）' % (rule, '；'.join(sorted(set(why))[:3])))
        else:
            results.append('-0：%s' % rule)

        # ---- 25~28 各页其余元素保持不变
        for rule, groups, score in (
                ('第2页其余元素保持不变：右上角“课程逻辑总览”须为微软雅黑10.5磅、字色5D6470，'
                 '左边界23.11cm、上边界1.02cm、宽8.64cm；页脚须为7.8磅、字色7B8790，'
                 '左边界1.57cm、上边界18.11cm',
                 [('', s2, SLIDE2_KEEP, None)], -1),
                ('第1页除背景外内容一律不动：主标题、副标题与“学情洞察”“任务驱动”“场景实践”'
                 '“评价成长”四个白色10.5磅加粗标签均须保留且字号、字色、位置不变',
                 [('', s1, SLIDE1_KEEP, None)], -3),
                ('第3、4、5页内容与版式一律不动：三页的红色标题、各组白字标签与说明、卡片、'
                 '第4页左下角图片（10.79cm×2.84cm，左边界2.08cm/上边界13.67cm）、'
                 '第5页左侧竖排标签与底部“输出成果……”的文字、字号、字色与位置均保持不变',
                 [('第3页', slides[2], SLIDE3_KEEP, None),
                  ('第4页', slides[3], SLIDE4_KEEP, SLIDE4_PIC),
                  ('第5页', slides[4], SLIDE5_KEEP, None)], -3),
                ('第6页除新增音频与播放按钮外内容一律不动：标题、四行标签与说明、右侧'
                 '“形成性评价”“资源支撑”两栏、结语均须保持字号、字色与位置',
                 [('', s6, SLIDE6_KEEP, None)], -1)):
            bad = []
            for tag, slide, table, pic in groups:
                for msg in _check_anchors(slide, table, scheme):
                    bad.append(tag + msg)
                if pic is not None:
                    pics = [p for p in _pictures(slide)
                            if _near(_cm(p.left), pic[0]) and _near(_cm(p.top), pic[1])
                            and _near(_cm(p.width), pic[2]) and _near(_cm(p.height), pic[3])]
                    if not pics:
                        bad.append('%s左下角图片位置或尺寸已改变（应为%.2f×%.2fcm，左%.2f/上%.2f）'
                                   % (tag, pic[2], pic[3], pic[0], pic[1]))
            if bad:
                results.append('%d：%s（%s）' % (score, rule, '；'.join(bad[:4])))
            else:
                results.append('-0：%s' % rule)

        # ---- 29 随文件保存的主题、默认样式与自定义部件保持
        rule = ('文稿保留随文件保存的主题字体多语种映射回退表、演示文稿级默认文本样式（9级）、'
                '视图设置部件、自定义标记数据部件（保留不少于28个ppt/tags/tag*.xml及形状引用）、'
                '以及文档属性中的作者/最后修改者/制作应用信息')
        issues = _package_issues(zf)
        if issues:
            results.append('-3：%s（%s）' % (rule, '；'.join(issues[:5])))
        else:
            results.append('-0：%s' % rule)

        # ========== 将 results 反解为结构化 dim2_items ==========
        # 逐项按追加顺序对齐 _MAX_DELTAS_ORDER：解析每条 "+N：说明" / "-N：说明"
        # 的 N 与 rule；加分项 N>0 表示命中，减分项 N<0 表示触发扣分，N=0 均为未命中。
        if len(results) != len(_MAX_DELTAS_ORDER):
            raise RuntimeError(
                'dim2 项数不一致：results=%d vs order=%d'
                % (len(results), len(_MAX_DELTAS_ORDER)))
        dim2_items = []
        total_score = 0
        for raw, max_delta in zip(results, _MAX_DELTAS_ORDER):
            head, _, rule = raw.partition('：')
            try:
                delta = int(head.replace('+', '').strip())
            except ValueError:
                delta = 0
            hit = delta != 0
            dim2_items.append({
                'rule': rule,
                'max_delta': max_delta,
                'delta': max_delta if hit else 0,
                'hit': hit,
            })
            total_score += (max_delta if hit else 0)

        result['dim2_items'] = dim2_items
        result['total_score'] = max(0, total_score)
        result['max_score'] = MAX_SCORE_TOTAL
        return result
    finally:
        try:
            zf.close()
        except Exception:
            pass


class _TargetMissing(Exception):
    """目录不存在或目录中没有 .pptx：连待评测对象都没有，不进入维度判定。"""


def _is_damaged_package(exc):
    """判断异常是否源于交付文件自身结构损坏（缺部件、XML 不合法、OPC 关系
    不完整等），而非脚本逻辑缺陷。"""
    text = str(exc)
    if isinstance(exc, KeyError) and (
        "There is no item named" in text or "no relationship of type" in text
    ):
        return True
    return type(exc).__name__ in ("XMLSyntaxError", "ParseError", "BadZipFile",
                                  "InvalidXmlError")


def evaluate(dir_path):
    """按接口约定返回结构化 dict；不 print、不改 sys.stdout、不 sys.exit。
    入参 dir_path 为“脚本所在目录的路径”，脚本自己在该目录里定位并打开被评估的
    .pptx（忽略以 ~$ 开头的 PowerPoint 临时锁文件）。脚本自身崩溃（含目录不存在、
    未发现文档等）由本函数捕获并返回 status='error'。"""
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
            raise _TargetMissing('目录不存在：%s' % dir_path)
        cands = sorted(
            f for f in os.listdir(dir_path)
            if f.lower().endswith('.pptx') and not f.startswith('~$')
        )
        if not cands:
            raise _TargetMissing('目录中未找到 .pptx 文件：%s' % dir_path)
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
        if _is_damaged_package(exc):
            result["status"] = "ok"
            result["error"] = None
            result["dim1_pass"] = False
            result["dim1_reason"] = "文件结构不完整，无法正常解析：%s: %s" % (
                type(exc).__name__, exc)
            result["dim2_items"] = []
            result["total_score"] = 0
            return result
        # 文件存在但打不开 / 解析失败：维度1 判不通过，dim1_reason 给出错误摘要
        result["status"] = "error"
        result["error"] = '%s: %s' % (type(exc).__name__, exc)
        result["dim1_pass"] = False
        result["dim1_reason"] = '文件无法解析或无法打开：%s: %s' % (type(exc).__name__, exc)
        result["dim2_items"] = []
        result["total_score"] = 0
        return result


if __name__ == "__main__":
    # 本地调试入口：evaluate 只接收“脚本所在目录的路径”，由脚本自己在该目录里
    # 定位并打开被评估的 .pptx。命令行第 1 个参数可覆盖目录路径；未指定时使用
    # 脚本自身所在目录。
    if len(sys.argv) > 1:
        _dir = sys.argv[1]
    else:
        _dir = os.path.dirname(os.path.abspath(__file__))
    _out = evaluate(_dir)
    print(json.dumps(_out, ensure_ascii=False, indent=2))
