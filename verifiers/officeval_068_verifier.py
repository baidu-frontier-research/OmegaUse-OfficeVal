#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# pyright: reportUnknownMemberType=false, reportUnknownVariableType=false, reportUnknownArgumentType=false, reportMissingImports=false
"""
068 —— 汇报 PDF 与调度坐席 PPT 合并为浅蓝可编辑演示文稿 完成度校验器

指令：我现在要把两个文件合在一起后修改一下，把"综合调度坐席系统操作介绍"ppt里的
      第13/16/18页删掉不用了，之后把"综合调度坐席系统操作介绍"ppt剩余的页面放在
      第二个PDF的第8页后面，最后给我一个做好的PPT文件。所有PPT页面背景及字体改为
      蓝色系，都做成可编辑的

说明：按接口统一约定，本模块不修改全局 sys.stdout，也不 print 主结果。
主结果统一由 evaluate(dir_path) -> dict 返回；入参为"脚本所在目录的路径"，
由脚本自己在该目录内定位并打开被评估的 .pptx。本文件仅在 __main__ 下作为
本地调试入口，才把 dict 序列化为 JSON 打印。

判定一律不依赖任何外部基准文件：细则里写明的每一处常量（页数、逐页页序锚点、
须保留的逐条文案、须缺席的三页标题、页面尺寸、蓝色相区间、可编辑门槛、
矢量轮廓门槛、正文字号门槛）都以常量形式内联在本文件中，脚本只读被评估的
那一个 .pptx。
"""

import os
import re
import sys
import json
import zipfile
import colorsys
import collections
from xml.etree import ElementTree as ET

SCRIPT_ID = "068"
MAX_SCORE_TOTAL = 50

# 加分项满分序列，顺序与 results 的 append 顺序严格一致
_MAX_DELTAS_ORDER = [
    5,   # 01 交付 1 个 .pptx，含两部分内容共 25 页，无空白页无重复页
    3,   # 02 不含"预约会议：把参会方和发起流程提前准备好"
    3,   # 03 不含"工单列表与预约回呼：让未完成事项可追踪"
    3,   # 04 不含"值守检查清单"
    5,   # 05 第8页为"典型设备组成"，操作介绍紧随其后插入
    5,   # 06 第24/25页为"运行流程与管理要点""谢谢聆听"
    3,   # 07 第9-23页共15页依次就位
    5,   # 08 25页页面背景底色均为蓝色系
    5,   # 09 25页文字着色均为蓝色系
    3,   # 10 蓝底与蓝字之间保持可读对比
    5,   # 11 全部页面文字为可编辑文本
    5,   # 12 全部示意图形为可编辑矢量形状
]
assert sum(_MAX_DELTAS_ORDER) == MAX_SCORE_TOTAL

# 减分项扣分序列，顺序与 results 中减分项的 append 顺序严格一致（共 24 项）
_PENALTY_DELTAS_ORDER = [-3, -3, -3, -1, -3, -3, -1, -5, -3, -1, -1] + [-1] * 13

# ============================ 命名空间 ============================
NS_P = 'http://schemas.openxmlformats.org/presentationml/2006/main'
NS_A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'

# ============================ 版式常量 ============================
SLIDE_TOTAL = 25
EMU_PER_IN = 914400.0
EMU_PER_CM = 360000.0
# 幻灯片尺寸：宽 14.22 英寸 × 高 8.00 英寸（1024×576 磅）
SLIDE_W_IN, SLIDE_H_IN = 14.222, 8.0
SIZE_TOL_IN = 0.15

# 逐页页序锚点：该页归一化文本必须含有下列全部片段，页序据此认定（不看页码标注）
ORDER_ANCHORS = {
    1:  ['应急有线人工通信保障网', '定义场景价值与台站体系解析'],
    2:  ['CONTENTS'],
    3:  ['什么是人工通信保障网'],
    4:  ['从常态值守到突发联络的备用语音链路'],
    5:  ['核心能力价值一', '强调可用可控可审计'],
    6:  ['核心能力价值二', '让应急联络有顺序有冗余有恢复路径'],
    7:  ['台站体系与协作架构'],
    8:  ['典型设备组成'],
    9:  ['操作介绍与值守流程'],
    10: ['登录与界面', '值守规范'],
    11: ['系统总览调度坐席的四层协作'],
    12: ['登录进入值守环境'],
    13: ['主界面把高频操作放在同一屏'],
    14: ['焦点规则先确认操作再确认语音'],
    15: ['工单窗口先记录再执行呼叫动作'],
    16: ['线路状态用可视化面板降低误操作'],
    17: ['接续流程从来话到三方沟通'],
    18: ['转接与交叉操作语音不一定跟着窗口走'],
    19: ['合并与拆分处理复杂通话关系'],
    20: ['查号用更少输入找到正确联系人'],
    21: ['即时会议临时协同但不污染会议列表'],
    22: ['座席互联用于班组内快速沟通'],
    23: ['主菜单按业务动作重新归类'],
    24: ['运行流程与管理要点', '让人工链路形成可追踪的闭环'],
    25: ['谢谢聆听', 'THANKSFORLISTENING'],
}
# 操作介绍内容所占的页位区间（第9页封面 + 第10-23页正文）
AGENT_PAGES = list(range(9, 24))
# 带眉标与页脚的操作介绍正文页（封面页无眉标页脚，不纳入该项）
AGENT_BODY_PAGES = list(range(10, 24))
PLACEHOLDER = 8            # 操作介绍内容须紧随其后插入的页位
TAIL_PAGES = (24, 25)      # 专题汇报末两页所在页位

# 须整页缺席的三页标题（指令明确要求删除）
DELETED_TITLES = [
    '预约会议把参会方和发起流程提前准备好',
    '工单列表与预约回呼让未完成事项可追踪',
    '值守检查清单',
]
# 被删除三页的专属文案：若出现在任何保留页，说明删除页内容被挪并进了别的页面
DELETED_BODY_TOKENS = ['把参会方', '发起流程', '预约回呼', '工单列表',
                       '值守检查', '检查清单']
# 两部分内容各自的识别标识
PART_A_KEY = '应急有线人工通信保障网'
PART_B_KEY = '云岚调度坐席平台'
# 页面家具（页眉眉标 / 页脚），做重复文本判定时须排除
FURNITURE = ('操作培训', '云岚调度坐席平台')
# 指令未涉及、属于凭空新增的说明文字与小标题
ADDED_HEADINGS = ['四个部分构成完整汇报框架', '协作说明', '运行说明',
                  '操作说明', '快捷键', '会场窗口能力']
# 标注页面来源的附加文字
PROVENANCE_RE = re.compile(r'(?:PDF|PPT)原第\d+页')
# 第9页封面须分行独立呈现、不得被"｜"并成一行的两处文字
COVER_SPLIT_PAIR = ('操作介绍与值守流程', '面向调度中心服务热线园区值守等场景')

# 逐页须保留的文案（键为页位，值为该页归一化文本须含有的全部片段）
KEEP_TOC_A = (2, ['网络定位与构成', 'NetworkPositioningandComposition',
                  '典型保障场景', 'TypicalSupportScenarios',
                  '核心能力价值', 'CoreCapabilitiesandValue',
                  '台站层级与设备', 'StationHierarchyandEquipment'])
KEEP_P3 = (3, ['关键词人工接续有线传输层级汇接闭环记录', '四个基础组成',
               '人工接续台', '有线传输链路', '用户终端', '安全与记录'])
KEEP_P4 = (4, ['从常态值守到突发联络的备用语音链路', '应急指挥备份', '园区专线联络',
               '偏远站点接入', '跨班值守交接', '大型活动保障', '灾害抢险协同'])
KEEP_P56 = [(5, ['强调可用可控可审计', '稳定性', '可控性']),
            (6, ['让应急联络有顺序有冗余有恢复路径', '优先级调度', '快速恢复',
                 '低依赖运行', '易培训维护'])]
KEEP_P7 = (7, ['层级汇接分区管理按需接入', '中心汇接站', '区域转接站',
               '现场保障站', '终端接入点'])
KEEP_P8 = (8, ['人工交换坐席', '线路分配架', '有线传输介质',
               '安全记录设备', '值守终端'])
KEEP_P24 = (24, ['让人工链路形成可追踪的闭环', '接入确认', '记录需求', '选择通道',
                 '呼出转接', '归档复核', '可靠接入', '身份核对', '闭环记录'])
KEEP_P25 = (25, ['THANKSFORLISTENING', '稳定可控可恢复',
                 '人工有线通信链路为应急联络保留可靠的组织保障'])
KEEP_TOC_B = (10, ['登录与界面', '工单处理', '接续操作',
                   '检索查号', '会议协作', '值守规范'])
KEEP_P11 = (11, ['来话接入', '坐席工作区', '路由与交换', '协作与归档'])
KEEP_P16 = (16, ['L%02d' % n for n in range(1, 21)])

# ============================ 判定门槛 ============================
# 蓝色系：色相落在 185°-255° 视为蓝色相（深蓝、中蓝、浅蓝、蓝灰均计入）
BLUE_HUE_LO, BLUE_HUE_HI = 185.0, 255.0
NEUTRAL_S, NEUTRAL_L_HI, NEUTRAL_L_LO = 0.12, 0.94, 0.06
# 深蓝底上的白字与近白字视为达标：明度高于此值即认定为白色或近白色
NEAR_WHITE_L = 0.90
# 【通用判定门槛】1）"蓝色系"只针对页面背景底色与文字着色两处，图形填充、
#   图形描边、装饰细线、表格单元格边框均不参与判定。
# 文字着色判定的容差：异色文字 run 数占全篇有文字 run 的比例上限
FONT_OFF_RATIO = 0.02
# 文字与其所在底色的对比度下限（低于此值视为压底不可辨识）
CONTRAST_MIN = 2.0
# 说明性正文字号下限（磅）；只对归一化后长度≥12 的成句文本设限，页码/眉标/页脚不计
BODY_SIZE_MIN = 10.0
BODY_TEXT_MIN_LEN = 12
# 矢量轮廓门槛：全篇自由绘制轮廓（custGeom）总数下限
CUSTGEOM_MIN = 20
# 同页重复承载判定：单页出现的"文本包含/重复"关系数超过此值即认定重复叠压
#   （标题与页脚眉标之间的一处正常回显不计入）
DUP_REL_MAX = 1
DUP_MIN_LEN = 6

# 归一化时剔除的空白与标点（保留数字，页码/线路编号判定要用）
_PUNCT = set('　 \t\r\n·、，。；：？！“”‘’（）()《》〈〉【】[]{},.;:?!|｜/\\-—–~•*"\'')


# ============================ 基础工具 ============================
def _q(uri, name):
    return '{%s}%s' % (uri, name)


def _norm(s):
    """去掉空白与标点后的比较用文本（保留数字与中英文字符）。"""
    return ''.join(ch for ch in (s or '') if ch not in _PUNCT)


def _key(s):
    """重复文本判定用的键：在 _norm 基础上再去掉数字（页码/序号不参与比较）。"""
    return ''.join(ch for ch in _norm(s) if not ch.isdigit())


class _Package(object):
    """按 OPC 包读取 .pptx；只读，不改写被评估文件。"""

    def __init__(self, path):
        self.zf = zipfile.ZipFile(path)
        self.names = list(self.zf.namelist())
        self.nameset = set(self.names)
        self._theme_cache = {}

    def close(self):
        try:
            self.zf.close()
        except Exception:
            pass

    def has(self, part):
        return part in self.nameset

    def text(self, part):
        try:
            return self.zf.read(part).decode('utf-8', 'ignore')
        except KeyError:
            return ''

    def xml(self, part):
        return ET.fromstring(self.zf.read(part))

    def match(self, pattern):
        rx = re.compile(pattern)
        return [n for n in self.names if rx.match(n)]

    def _rels(self, part):
        d, b = os.path.split(part)
        rp = (d + '/_rels/' + b + '.rels').lstrip('/')
        if not self.has(rp):
            return None
        return self.xml(rp)

    def follow(self, part, type_suffix):
        """沿关系链找到 part 引用的指定类型部件名；找不到返回 None。"""
        if not part:
            return None
        rels = self._rels(part)
        if rels is None:
            return None
        for rel in rels:
            if (rel.get('Type') or '').endswith(type_suffix):
                tgt = (rel.get('Target') or '').replace('\\', '/')
                base = os.path.dirname(part)
                while tgt.startswith('../'):
                    tgt = tgt[3:]
                    base = os.path.dirname(base)
                return (base + '/' + tgt).lstrip('/')
        return None

    def slide_parts(self):
        """按 sldIdLst 的放映顺序返回幻灯片部件名列表。"""
        pres = self.xml('ppt/presentation.xml')
        rels = self.xml('ppt/_rels/presentation.xml.rels')
        rmap = {}
        for rel in rels:
            rmap[rel.get('Id')] = rel.get('Target') or ''
        out = []
        lst = pres.find(_q(NS_P, 'sldIdLst'))
        if lst is None:
            return out
        for sid in lst:
            tgt = rmap.get(sid.get(_q(NS_R, 'id')), '').replace('\\', '/')
            if not tgt:
                continue
            while tgt.startswith('../'):
                tgt = tgt[3:]
            if not tgt.startswith('ppt/'):
                tgt = 'ppt/' + tgt.lstrip('/')
            out.append(tgt)
        return out

    def slide_size(self):
        pres = self.xml('ppt/presentation.xml')
        sz = pres.find(_q(NS_P, 'sldSz'))
        if sz is None:
            return (None, None)
        try:
            return (int(sz.get('cx')), int(sz.get('cy')))
        except (TypeError, ValueError):
            return (None, None)

    def clr_scheme(self, slide_part):
        """与办公软件对齐：主题色取自该页所绑定 版式→母版→主题 的 clrScheme。"""
        if slide_part in self._theme_cache:
            return self._theme_cache[slide_part]
        lay = self.follow(slide_part, 'slideLayout')
        mas = self.follow(lay, 'slideMaster')
        th = self.follow(mas, 'theme')
        scheme = {}
        if th and self.has(th):
            root = self.xml(th)
            for sch in root.iter(_q(NS_A, 'clrScheme')):
                for child in sch:
                    name = child.tag.split('}')[-1]
                    srgb = child.find(_q(NS_A, 'srgbClr'))
                    if srgb is not None:
                        scheme[name] = (srgb.get('val') or '').upper()
                        continue
                    sysc = child.find(_q(NS_A, 'sysClr'))
                    if sysc is not None:
                        scheme[name] = (sysc.get('lastClr') or '').upper()
                break
        self._theme_cache[slide_part] = scheme
        return scheme


# ============================ 颜色解析 ============================
# 与办公软件对齐：schemeClr 先按 clrMap 的常用别名折回主题槽位，再套用
#   lumMod/lumOff 亮度修饰，得到 PowerPoint 里眼睛看到的那个颜色。
_SCHEME_ALIAS = {'bg1': 'lt1', 'tx1': 'dk1', 'bg2': 'lt2', 'tx2': 'dk2',
                 'phClr': 'accent1'}
_HEX6 = re.compile(r'^[0-9A-F]{6}$')


def _resolve_color(el, scheme):
    """把 <a:srgbClr>/<a:schemeClr> 解析成 6 位十六进制；无法解析返回 None。"""
    if el is None:
        return None
    tag = el.tag.split('}')[-1]
    if tag == 'srgbClr':
        hexv = (el.get('val') or '').upper()
    elif tag == 'schemeClr':
        slot = el.get('val') or ''
        hexv = scheme.get(_SCHEME_ALIAS.get(slot, slot)) or ''
    else:
        return None
    if not _HEX6.match(hexv):
        return None
    r, g, b = (int(hexv[i:i + 2], 16) / 255.0 for i in (0, 2, 4))
    h, l, s = colorsys.rgb_to_hls(r, g, b)
    lm = el.find(_q(NS_A, 'lumMod'))
    lo = el.find(_q(NS_A, 'lumOff'))
    if lm is not None and (lm.get('val') or '').isdigit():
        l *= int(lm.get('val')) / 100000.0
    if lo is not None and (lo.get('val') or '').isdigit():
        l += int(lo.get('val')) / 100000.0
    l = max(0.0, min(1.0, l))
    r, g, b = colorsys.hls_to_rgb(h, l, s)
    return '%02X%02X%02X' % tuple(int(round(v * 255)) for v in (r, g, b))


def _hsl(hexv):
    r, g, b = (int(hexv[i:i + 2], 16) / 255.0 for i in (0, 2, 4))
    h, l, s = colorsys.rgb_to_hls(r, g, b)
    return (h * 360.0, s, l)


def _is_neutral(hexv):
    h, s, l = _hsl(hexv)
    return s < NEUTRAL_S or l > NEUTRAL_L_HI or l < NEUTRAL_L_LO


def _is_blue(hexv):
    h, s, l = _hsl(hexv)
    return BLUE_HUE_LO <= h <= BLUE_HUE_HI


def _is_near_white(hexv):
    """白色与近白色：深蓝底上的白字属达标着色，不计为异色。"""
    return _hsl(hexv)[2] >= NEAR_WHITE_L


def _rel_lum(hexv):
    def _f(c):
        c /= 255.0
        return c / 12.92 if c <= 0.03928 else ((c + 0.055) / 1.055) ** 2.4
    r, g, b = (int(hexv[i:i + 2], 16) for i in (0, 2, 4))
    return 0.2126 * _f(r) + 0.7152 * _f(g) + 0.0722 * _f(b)


def _contrast(a, b):
    la, lb = _rel_lum(a), _rel_lum(b)
    hi, lo = max(la, lb), min(la, lb)
    return (hi + 0.05) / (lo + 0.05)


# ============================ 形状与文本采集 ============================
def _shapes(root, scheme):
    """采集页内每个形状：位置尺寸、整块文本、逐 run 的字号与文字色、自身填充色。

    与办公软件对齐：位置尺寸取 a:xfrm/a:off、a:ext（EMU）；文字取 a:t 的可见字符；
      字号取 a:rPr/@sz（1/100 磅）；文字色与形状填充取各自的 a:solidFill 首子元素。
    """
    out = []
    for sp in root.iter(_q(NS_P, 'sp')):
        off = next(iter(sp.iter(_q(NS_A, 'off'))), None)
        ext = next(iter(sp.iter(_q(NS_A, 'ext'))), None)
        x = int(off.get('x')) if off is not None and off.get('x') else None
        y = int(off.get('y')) if off is not None and off.get('y') else None
        cx = int(ext.get('cx')) if ext is not None and ext.get('cx') else 0
        cy = int(ext.get('cy')) if ext is not None and ext.get('cy') else 0
        runs = []
        for r in sp.iter(_q(NS_A, 'r')):
            t = r.find(_q(NS_A, 't'))
            txt = (t.text or '') if t is not None else ''
            rpr = r.find(_q(NS_A, 'rPr'))
            sz = None
            color = None
            if rpr is not None:
                if (rpr.get('sz') or '').isdigit():
                    sz = int(rpr.get('sz')) / 100.0
                sf = rpr.find(_q(NS_A, 'solidFill'))
                if sf is not None and len(sf):
                    color = _resolve_color(sf[0], scheme)
            runs.append({'sz': sz, 'text': txt, 'color': color})
        fill = None
        spPr = sp.find(_q(NS_P, 'spPr'))
        if spPr is not None:
            sf = spPr.find(_q(NS_A, 'solidFill'))
            if sf is not None and len(sf):
                fill = _resolve_color(sf[0], scheme)
        out.append({
            'x': x, 'y': y, 'cx': cx, 'cy': cy, 'fill': fill, 'runs': runs,
            'text': ''.join(r['text'] for r in runs),
        })
    return out


def _bg_of_part(root, scheme):
    """某个部件自身声明的底色：p:bg 的纯色填充或 p:bgRef 引用的主题色。"""
    for bg in root.iter(_q(NS_P, 'bg')):
        for sf in bg.iter(_q(NS_A, 'solidFill')):
            if len(sf):
                c = _resolve_color(sf[0], scheme)
                if c:
                    return c
        for br in bg.iter(_q(NS_P, 'bgRef')):
            for child in br:
                c = _resolve_color(child, scheme)
                if c:
                    return c
    return None


def _eff_bg(pkg, slide_part, scheme):
    """与办公软件对齐：页面实际生效的底色 —— 本页自设底色优先，未自设则
    依次向其版式、母版继承（PowerPoint 的背景继承链就是这三级）。"""
    lay = pkg.follow(slide_part, 'slideLayout')
    mas = pkg.follow(lay, 'slideMaster') if lay else None
    for part in (slide_part, lay, mas):
        if not part or not pkg.has(part):
            continue
        c = _bg_of_part(pkg.xml(part), scheme)
        if c:
            return c
    return None


def _table_texts(root):
    """表格单元格文本；GraphicFrame 里的 a:tbl 也承载正文，须一并纳入页面文本。"""
    out = []
    for tbl in root.iter(_q(NS_A, 'tbl')):
        for tc in tbl.iter(_q(NS_A, 'tc')):
            out.append(''.join((t.text or '') for t in tc.iter(_q(NS_A, 't'))))
    return out


# ============================ 对外接口 ============================
class _TargetMissing(Exception):
    """目录不存在或目录中没有 .pptx：连待评测对象都没有，不进入维度判定。"""


def evaluate(dir_path):
    """入参为目录路径，在该目录内定位 .pptx 并给出完成度评分。"""
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
            raise _TargetMissing(f"目录不存在：{dir_path}")
        cands = sorted(f for f in os.listdir(dir_path)
                       if f.lower().endswith('.pptx') and not f.startswith('~$'))
        if not cands:
            raise _TargetMissing(f"目录中未找到 .pptx 文件：{dir_path}")
        file_path = os.path.join(dir_path, cands[0])
        result["file_name"] = os.path.basename(file_path)
        return _evaluate_impl(file_path, dir_path, result)
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
        result["error"] = f"{type(exc).__name__}: {exc}"
        result["dim1_pass"] = False
        result["dim1_reason"] = f"文件无法解析或无法打开：{type(exc).__name__}: {exc}"
        result["dim2_items"] = []
        result["total_score"] = 0
        return result


def _dim1_gate(file_path):
    """维度1：交付文件为1个可正常打开的 PowerPoint 演示文稿（.pptx）。

    与办公软件对齐：PowerPoint / WPS 打开 .pptx 时先校验 OPC 包结构与
      [Content_Types].xml，再逐部件解析 XML；任一环节失败即"文件损坏无法打开"
      或弹出"发现不可读取的内容"。此处按同一条链路检查：
      ① 是 zip（OPC）包；② 有 [Content_Types].xml 与 ppt/presentation.xml；
      ③ presentation.xml、每张幻灯片、母版、版式的 XML 均可解析；
      ④ sldIdLst 指向的幻灯片部件都真实存在。
    可编辑性与版面可读性不在本维度设卡，改由维度2 的对应加分项计分。
    """
    if not zipfile.is_zipfile(file_path):
        return (None, '文件不是有效的 OPC(zip) 包，无法作为 .pptx 打开')
    try:
        pkg = _Package(file_path)
    except Exception as exc:
        return (None, '演示文稿包无法打开：%s' % exc)
    for part in ('[Content_Types].xml', 'ppt/presentation.xml'):
        if not pkg.has(part):
            pkg.close()
            return (None, '缺少必需部件 %s，打开时会提示"发现不可读取的内容"' % part)
    try:
        slides = pkg.slide_parts()
        if not slides:
            raise ValueError('sldIdLst 为空，演示文稿内没有可放映的幻灯片')
        for part in slides:
            if not pkg.has(part):
                raise ValueError('幻灯片部件缺失：%s' % part)
            pkg.xml(part)
        for part in (pkg.match(r'ppt/slideMasters/slideMaster\d+\.xml$')
                     + pkg.match(r'ppt/slideLayouts/slideLayout\d+\.xml$')):
            pkg.xml(part)
    except Exception as exc:
        pkg.close()
        return (None, '演示文稿内容无法正常解析：%s: %s' % (type(exc).__name__, exc))
    return (pkg, '')


def _evaluate_impl(file_path, dir_path, result):
    # ---------------- 维度1 ----------------
    pkg, reason = _dim1_gate(file_path)
    if pkg is None:
        result["status"] = "error"
        result["error"] = reason
        result["dim1_pass"] = False
        result["dim1_reason"] = reason
        result["dim2_items"] = []
        result["total_score"] = 0
        result["max_score"] = MAX_SCORE_TOTAL
        return result
    result["dim1_pass"] = True
    result["dim1_reason"] = "交付物为1个可正常打开的 .pptx，包结构与各部件 XML 均可解析"
    try:
        return _score(pkg, file_path, dir_path, result)
    finally:
        pkg.close()


def _score(pkg, file_path, dir_path, result):
    results = []      # ["+N：规则说明", ...]，与 _MAX_DELTAS_ORDER 逐项对齐
    details = []      # 与 results 等长的取证说明

    def emit(delta, rule, detail=""):
        results.append("%+d：%s" % (delta, rule))
        details.append(detail)

    # ---------------- 数据采集 ----------------
    parts = pkg.slide_parts()
    n = len(parts)
    pages = range(1, n + 1)
    roots, raws, shapes, ptext, schemes = {}, {}, {}, {}, {}
    for i, part in enumerate(parts, 1):
        roots[i] = pkg.xml(part)
        raws[i] = pkg.text(part)
        sch = pkg.clr_scheme(part)
        schemes[i] = sch
        shapes[i] = _shapes(roots[i], sch)
        segs = [sp['text'] for sp in shapes[i]] + _table_texts(roots[i])
        ptext[i] = _norm(''.join(segs))
    alltext = ''.join(ptext.values())
    cust_total = sum(len(re.findall(r'<a:custGeom', raws[i])) for i in pages)
    pic_total = sum(len(re.findall(r'<p:pic>', raws[i])) for i in pages)
    cx, cy = pkg.slide_size()

    def has_all(i, tokens):
        """第 i 页缺失的片段列表。"""
        if i not in ptext:
            return list(tokens)
        return [t for t in tokens if t not in ptext[i]]

    def page_number_present(i):
        """第 i 页是否带页码标注：存在文本恰为该页序号，或以该页序号结尾。"""
        pat_exact = re.compile(r'^0?%d$' % i)
        pat_tail = re.compile(r'\s0?%d$' % i)
        for sp in shapes[i]:
            t = (sp['text'] or '').strip()
            if pat_exact.match(t) or pat_tail.search(t):
                return True
        return False

    # ==================== 维度2：加分项 ====================
    # +5：交付1个 .pptx，同时包含专题汇报与操作介绍两部分内容，合计25页，
    #     无空白页、无重复页
    pptx_cnt = len([f for f in os.listdir(dir_path)
                    if f.lower().endswith('.pptx') and not f.startswith('~$')])
    blank = [i for i in pages if not ptext[i]]
    seen_pages = collections.Counter(ptext[i] for i in pages if ptext[i])
    repeat = [t[:14] for t, c in seen_pages.items() if c >= 2]
    a01 = (pptx_cnt == 1 and n == SLIDE_TOTAL and not blank and not repeat
           and PART_A_KEY in alltext and PART_B_KEY in alltext)
    emit(5 if a01 else 0,
         "交付1个 .pptx 文件，同时包含\"应急有线人工通信保障网\"专题汇报内容与\"云岚调度坐席平台\"操作介绍内容，合计25页，无空白页、无重复页",
         ".pptx 数=%d 页数=%d 空白页=%s 重复页=%s 两部分内容=%s/%s"
         % (pptx_cnt, n, blank or '无', repeat or '无',
            PART_A_KEY in alltext, PART_B_KEY in alltext))

    # +3 ×3：三页整页缺席（指令要求删除的三页）
    for title, label in zip(DELETED_TITLES,
                            ('预约会议：把参会方和发起流程提前准备好',
                             '工单列表与预约回呼：让未完成事项可追踪',
                             '值守检查清单')):
        hit_pages = [i for i in pages if title in ptext[i]]
        emit(3 if not hit_pages else 0,
             "全篇不包含标题为\"%s\"的页面" % label,
             "残留页：%s" % (hit_pages or '无'))

    # +5：第8页为"典型设备组成"页，操作介绍内容整体紧随该页之后插入
    a05_bad = []
    if n < 9:
        a05_bad.append('总页数不足9页')
    else:
        if has_all(PLACEHOLDER, ORDER_ANCHORS[PLACEHOLDER]):
            a05_bad.append('第8页不是"典型设备组成"页')
        if has_all(9, ORDER_ANCHORS[9]):
            a05_bad.append('第9页不是"云岚调度坐席平台"封面页')
        for i in range(1, PLACEHOLDER):
            if has_all(i, ORDER_ANCHORS[i]):
                a05_bad.append('第%d页不在原有汇报页位上' % i)
    emit(5 if not a05_bad else 0,
         "第8页为\"典型设备组成\"页，操作介绍内容整体紧随该页之后插入",
         "异常：%s" % ('；'.join(a05_bad[:5]) or '无'))

    # +5：第24页"运行流程与管理要点"、第25页"谢谢聆听"位于全篇最后
    a06_bad = []
    for i in TAIL_PAGES:
        miss = has_all(i, ORDER_ANCHORS[i]) if i <= n else ORDER_ANCHORS[i]
        if miss:
            a06_bad.append('第%d页缺%s' % (i, '、'.join(miss)))
    if n != SLIDE_TOTAL:
        a06_bad.append('总页数为%d，末两页未落在第24、25页' % n)
    emit(5 if not a06_bad else 0,
         "第24页为\"运行流程与管理要点\"，第25页为\"谢谢聆听 / THANKS FOR LISTENING\"，即专题汇报末两页位于全篇最后",
         "异常：%s" % ('；'.join(a06_bad[:5]) or '无'))

    # +3：第9至第23页共15页依次就位且相对先后顺序一致
    a07_bad = []
    for i in AGENT_PAGES:
        miss = has_all(i, ORDER_ANCHORS[i]) if i <= n else ORDER_ANCHORS[i]
        if miss:
            a07_bad.append('第%d页应为"%s"' % (i, ORDER_ANCHORS[i][0]))
    emit(3 if not a07_bad else 0,
         "第9至第23页依次为\"云岚调度坐席平台\"封面、\"内容结构\"、\"系统总览：调度坐席的四层协作\"、\"登录：进入值守环境\"、\"主界面：把高频操作放在同一屏\"、\"焦点规则：先确认'操作'，再确认'语音'\"、\"工单窗口：先记录，再执行呼叫动作\"、\"线路状态：用可视化面板降低误操作\"、\"接续流程：从来话到三方沟通\"、\"转接与交叉操作：语音不一定跟着窗口走\"、\"合并与拆分：处理复杂通话关系\"、\"查号：用更少输入找到正确联系人\"、\"即时会议：临时协同但不污染会议列表\"、\"座席互联：用于班组内快速沟通\"、\"主菜单：按业务动作重新归类\"，共15页且相对先后顺序一致",
         "错位页：%s" % ('；'.join(a07_bad[:6]) or '无'))

    # +5：25页的页面背景底色均为蓝色系
    #    按【通用判定门槛】1）只看"页面背景"这一处：本页自设底色优先，未自设
    #    则取版式与母版继承的底色；图形填充、描边、表格边框一律不参与判定。
    eff_bg, bg_bad = {}, []
    for i in pages:
        c = _eff_bg(pkg, parts[i - 1], schemes[i])
        eff_bg[i] = c
        if c is None:
            bg_bad.append('第%d页未取到生效底色' % i)
        elif not _is_blue(c):
            bg_bad.append('第%d页底色#%s非蓝色系' % (i, c))
    emit(5 if not bg_bad else 0,
         "25页的页面背景底色均为蓝色系（色相落在蓝色区间，如 #F2F8FF / #E5EFF8 / #BDD7EE / #123D5A 一类），无白底、米黄底、浅灰底、绿底等非蓝色背景残留，专题汇报部分与操作介绍部分的背景同属蓝色系",
         "生效底色：%s；异常%d处：%s"
         % (' '.join('%s×%d页' % (h or '未设置', c) for h, c in
                     collections.Counter(eff_bg.values()).most_common(4)),
            len(bg_bad), '；'.join(bg_bad[:4]) or '无'))

    # +5：25页的文字着色均为蓝色系（白/近白字与中性灰字不计异色）
    font_tot, font_off, off_run = 0, 0, []
    for i in pages:
        for sp in shapes[i]:
            for r in sp['runs']:
                if not r['color'] or len(_norm(r['text'])) == 0:
                    continue
                font_tot += 1
                if (_is_blue(r['color']) or _is_near_white(r['color'])
                        or _is_neutral(r['color'])):
                    continue
                font_off += 1
                off_run.append('第%d页"%s"字色#%s'
                               % (i, _norm(r['text'])[:8], r['color']))
    off_share = (font_off / float(font_tot)) if font_tot else 0.0
    emit(5 if off_share <= FONT_OFF_RATIO else 0,
         "25页的文字着色均为蓝色系（深蓝、中蓝、蓝灰均可，如 #123D5A / #2C6FAA / #35A3DC / #5F7381 一类），深蓝底页面上的白色或近白色文字亦计达标，不出现橙色、青绿、草绿、正红等异色文字",
         "显式着色文字%d处，其中异色%d处（占比%.1f%%，容差%.0f%%）：%s"
         % (font_tot, font_off, off_share * 100, FONT_OFF_RATIO * 100,
            '；'.join(off_run[:3]) or '无'))

    # +3：蓝色背景与蓝色文字之间保持可读对比
    #    只判"蓝底 × 蓝字/白字"这一组合：文字所在形状有填充时压形状填充，
    #    否则压该页生效底色；异色文字的问题归上一条，不在此重复计分。
    low_ct = []
    for i in pages:
        for sp in shapes[i]:
            back = sp['fill'] or eff_bg[i]
            if not back or not _is_blue(back):
                continue
            for r in sp['runs']:
                if not r['color'] or len(_norm(r['text'])) < 2:
                    continue
                if not (_is_blue(r['color']) or _is_near_white(r['color'])):
                    continue
                cv = _contrast(r['color'], back)
                if cv < CONTRAST_MIN:
                    low_ct.append('第%d页"%s"字色#%s压底#%s对比度%.2f'
                                  % (i, _norm(r['text'])[:8], r['color'],
                                     back, cv))
    emit(3 if not low_ct else 0,
         "蓝色背景与蓝色文字之间保持可读对比（浅蓝底配深蓝或深蓝灰字，深蓝底配白字或浅蓝字），不出现蓝字压同明度蓝底而无法辨识的文字",
         "对比度低于%.1f的文字%d处：%s"
         % (CONTRAST_MIN, len(low_ct), '；'.join(low_ct[:3]) or '无'))

    # +5：全部页面文字均为可编辑文本
    txt_pages = [i for i in pages if not ptext[i]]
    a11 = (pic_total == 0 and not txt_pages)
    emit(5 if a11 else 0,
         "全部页面文字均为可编辑文本（可选中、可改字、可改字号字色），无以图片形式呈现的文字",
         "图片对象%d个；无可编辑文字的页面=%s" % (pic_total, txt_pages or '无'))

    # +5：全部示意图形与图标均为可编辑矢量形状，关键面板可逐格逐项改字
    a12_bad = []
    if pic_total:
        a12_bad.append('存在%d个图片对象' % pic_total)
    lp, fp = KEEP_P16[0], KEEP_P24[0]
    miss_l = has_all(lp, KEEP_P16[1])
    if miss_l:
        a12_bad.append('线路面板缺可改字标识%s' % '、'.join(miss_l[:4]))
    miss_f = has_all(fp, ['接入确认', '记录需求', '选择通道', '呼出转接', '归档复核'])
    if miss_f:
        a12_bad.append('流程节点缺可改字条目%s' % '、'.join(miss_f))
    emit(5 if not a12_bad else 0,
         "全部示意图形与图标均为可编辑矢量形状或形状组合（可单独选中、可改填充与轮廓），\"线路状态：用可视化面板降低误操作\"页的线路面板与\"运行流程与管理要点\"页的流程节点可逐格逐项改字，无以图片形式嵌入的插图",
         "异常：%s" % ('；'.join(a12_bad) or '无'))

    # ==================== 维度2：减分项 ====================
    def pen(delta, rule, violated, detail=""):
        results.append("%s：%s" % (("%d" % delta) if violated else "-0", rule))
        details.append(detail)

    # -3：页面上不得有"PDF原第N页""PPT原第N页"一类标注页面来源的附加文字
    prov = []
    for i in pages:
        m = PROVENANCE_RE.findall(''.join(sp['text'] for sp in shapes[i]))
        if m:
            prov.append('第%d页"%s"' % (i, m[0]))
    pen(-3, "25页任一页面不满足页面上无\"PDF原第1页\"至\"PDF原第10页\"\"PPT原第1页\"至\"PPT原第17页\"一类标注页面来源的附加文字",
        bool(prov), "命中%d页：%s" % (len(prov), '；'.join(prov[:4]) or '无'))

    # -3：页面尺寸须为宽 14.22 英寸 × 高 8.00 英寸
    w_in = (cx / EMU_PER_IN) if cx else 0.0
    h_in = (cy / EMU_PER_IN) if cy else 0.0
    size_bad = (abs(w_in - SLIDE_W_IN) > SIZE_TOL_IN
                or abs(h_in - SLIDE_H_IN) > SIZE_TOL_IN)
    pen(-3, "幻灯片页面尺寸不满足宽14.22英寸×高8.00英寸（36.12厘米×20.32厘米，1024×576 磅）",
        size_bad, "实测 %.3f×%.3f 英寸（%.2f×%.2f 厘米）"
        % (w_in, h_in, cx / EMU_PER_CM if cx else 0, cy / EMU_PER_CM if cy else 0))

    # -3：首页须保留署名"通信保障研究组"
    sign_bad = ('通信保障研究组' not in ptext.get(1, ''))
    pen(-3, "首页不满足在\"应急有线人工通信保障网\"\"定义、场景、价值与台站体系解析\"下方保留署名\"通信保障研究组\"",
        sign_bad, "首页署名=%s" % ('缺失' if sign_bad else '已保留'))

    # -1：首页不得新增"稳定""可控""可恢复""2026"等标签文字
    p1 = ptext.get(1, '')
    add_tag = [t for t in ('稳定', '可控', '可恢复', '2026') if t in p1]
    pen(-1, "首页不满足除\"应急有线人工通信保障网\"\"定义、场景、价值与台站体系解析\"\"通信保障研究组\"外无新增的\"稳定\"\"可控\"\"可恢复\"\"2026\"等标签文字",
        bool(add_tag), "新增标签=%s" % (add_tag or '无'))

    # -3："典型设备组成"页副标题须为指定文案
    sub8_bad = ('用自绘图形替代实拍设备保持表达一致' not in ptext.get(8, ''))
    pen(-3, "\"典型设备组成\"页副标题不满足文字为\"用自绘图形替代实拍设备，保持表达一致\"",
        sub8_bad, "第8页副标题=%s" % ('已被改写或缺失' if sub8_bad else '与要求一致'))

    # -3：各页不得新增原本不存在的说明文字与小标题
    add_head = []
    for t in ADDED_HEADINGS:
        hp = [i for i in pages if t in ptext[i]]
        if hp:
            add_head.append('"%s"(第%s页)' % (t, hp[0]))
    pen(-3, "各页不满足无新增的\"四个部分构成完整汇报框架\"\"协作说明\"\"运行说明\"\"操作说明\"\"快捷键\"\"会场窗口能力\"等原本不存在的说明文字与小标题",
        bool(add_head), "新增说明文字%d处：%s" % (len(add_head), '；'.join(add_head[:4]) or '无'))

    # -1："云岚调度坐席平台"封面页副标题与四项标签须分行独立呈现
    merged = []
    for sp in shapes.get(9, []):
        k = _norm(sp['text'])
        if COVER_SPLIT_PAIR[0] in k and COVER_SPLIT_PAIR[1] in k:
            merged.append(sp['text'][:26])
    pen(-1, "\"云岚调度坐席平台\"封面页不满足其副标题与\"登录\"\"接续\"\"会议\"\"回呼\"四项标签分行独立呈现，被合并为以\"｜\"分隔的单行",
        bool(merged), "合并成一行的文本框=%s" % (merged[:1] or '无'))

    # -5：示意图形须保留原有轮廓形态，不得被拍平为纯矩形/圆角矩形色块
    flat_bad = (cust_total < CUSTGEOM_MIN)
    pen(-5, "25页的示意图形不满足保留原有轮廓形态（含首页与各页装饰弧线、\"四个基础组成\"结构图示、六类保障场景图标、能力价值图标、台站层级示意图、五类设备示意图形、调度坐席界面模拟面板、人物头像圆形、流程箭头与连接线），被简化为纯矩形/圆角矩形色块或整体缺失",
        flat_bad, "全篇自由绘制轮廓(custGeom)=%d 个，门槛%d 个"
        % (cust_total, CUSTGEOM_MIN))

    # -3：说明性正文字号不得小于10磅
    small = []
    for i in pages:
        for sp in shapes[i]:
            for r in sp['runs']:
                if (r['sz'] is not None and r['sz'] < BODY_SIZE_MIN
                        and len(_norm(r['text'])) >= BODY_TEXT_MIN_LEN):
                    small.append('第%d页%.1f磅"%s"' % (i, r['sz'], r['text'][:12]))
    pen(-3, "说明性正文文字不满足字号不小于10磅（不出现5.6磅、6.2磅、7.2磅、8.0磅、8.5磅一类难以辨识的字号）",
        bool(small), "小字号正文%d处：%s" % (len(small), '；'.join(small[:3]) or '无'))

    # -1：同一条目文字只出现一次，不得重复叠压
    dup_pages = []
    for i in pages:
        keys = [_key(sp['text']) for sp in shapes[i]]
        keys = [k for k in keys
                if len(k) >= DUP_MIN_LEN and not any(f in k for f in FURNITURE)]
        rel = 0
        for a in range(len(keys)):
            for b in range(len(keys)):
                if a == b:
                    continue
                if keys[a] in keys[b] and (keys[a] != keys[b] or a < b):
                    rel += 1
                    break
        if rel > DUP_REL_MAX:
            dup_pages.append('第%d页%d处' % (i, rel))
    pen(-1, "任一页面不满足同一条目文字只出现一次，出现\"大型活动保障\"\"灾害抢险协同\"\"值守终端\"等条目重复叠压的同名图形或文字",
        bool(dup_pages), "重复叠压页：%s" % ('；'.join(dup_pages[:5]) or '无'))

    # -1：页码标注须连续完整无空缺
    no_num = [i for i in pages if not page_number_present(i)]
    pen(-1, "25页的页码标注不满足连续完整无空缺，出现部分页面页码框为空白的情况",
        bool(no_num), "缺页码的页面%d个：%s"
        % (len(no_num), no_num[:8] or '无'))

    # -1 ×13：逐页须原样保留的文案
    def keep(spec, rule):
        pg, tokens = spec
        miss = has_all(pg, tokens) if pg <= n else list(tokens)
        pen(-1, rule, bool(miss),
            "第%d页缺失%d处：%s" % (pg, len(miss), '、'.join(miss[:5]) or '无'))

    keep(KEEP_TOC_A,
         "\"内容结构 / 目录 CONTENTS\"页不满足保留\"01 网络定位与构成 Network Positioning and Composition\"\"02 典型保障场景 Typical Support Scenarios\"\"03 核心能力价值 Core Capabilities and Value\"\"04 台站层级与设备 Station Hierarchy and Equipment\"四条中英文目录项")
    keep(KEEP_P3,
         "\"什么是人工通信保障网？\"页不满足保留\"关键词：人工接续 · 有线传输 · 层级汇接 · 闭环记录\"整行文字及\"人工接续台\"\"有线传输链路\"\"用户终端\"\"安全与记录\"四个基础组成条目")
    keep(KEEP_P4,
         "\"02 典型保障场景\"页不满足保留副标题\"从常态值守到突发联络的备用语音链路\"及\"应急指挥备份\"\"园区专线联络\"\"偏远站点接入\"\"跨班值守交接\"\"大型活动保障\"\"灾害抢险协同\"六个场景条目")

    miss56 = []
    for pg, tokens in KEEP_P56:
        m = has_all(pg, tokens) if pg <= n else list(tokens)
        if m:
            miss56.append('第%d页缺%s' % (pg, '、'.join(m[:4])))
    pen(-1, "\"03 核心能力价值（一）（二）\"两页不满足保留副标题\"强调可用、可控、可审计\"\"让应急联络有顺序、有冗余、有恢复路径\"及\"稳定性\"\"可控性\"\"优先级调度\"\"快速恢复\"\"低依赖运行\"\"易培训维护\"六个能力条目",
        bool(miss56), "；".join(miss56) or '两页文案齐备')

    keep(KEEP_P7,
         "\"04 台站体系与协作架构\"页不满足保留副标题\"层级汇接、分区管理、按需接入\"及\"中心汇接站\"\"区域转接站\"\"现场保障站\"\"终端接入点\"四级台站条目")
    keep(KEEP_P8,
         "\"典型设备组成\"页不满足保留\"人工交换坐席\"\"线路分配架\"\"有线传输介质\"\"安全记录设备\"\"值守终端\"五类设备条目")
    keep(KEEP_P24,
         "\"运行流程与管理要点\"页不满足保留副标题\"让人工链路形成可追踪的闭环\"、\"1 接入确认\"\"2 记录需求\"\"3 选择通道\"\"4 呼出转接\"\"5 归档复核\"五步流程及\"可靠接入\"\"身份核对\"\"闭环记录\"三条值守原则")
    keep(KEEP_P25,
         "\"谢谢聆听\"页不满足保留\"THANKS FOR LISTENING\"\"稳定 · 可控 · 可恢复\"与\"人工有线通信链路为应急联络保留可靠的组织保障。\"三处文字")
    keep(KEEP_TOC_B,
         "\"内容结构\"（操作介绍部分）页不满足保留\"01 登录与界面\"至\"06 值守规范\"六条目录项")
    keep(KEEP_P11,
         "\"系统总览：调度坐席的四层协作\"页不满足保留四层协作的层级图示结构与各层名称文字")
    keep(KEEP_P16,
         "\"线路状态：用可视化面板降低误操作\"页不满足完整保留 L01 至 L20 共20个线路状态标识")

    furn_bad = [i for i in AGENT_BODY_PAGES
                if i <= n and not all(f in ptext[i] for f in FURNITURE)]
    pen(-1, "操作介绍各页不满足保留眉标\"操作培训\"字样与页脚\"云岚调度坐席平台\"字样",
        bool(furn_bad), "缺眉标或页脚的页面：%s" % (furn_bad or '无'))

    moved = []
    for t in DELETED_BODY_TOKENS:
        hp = [i for i in pages if t in ptext[i]]
        if hp:
            moved.append('"%s"(第%d页)' % (t, hp[0]))
    pen(-1, "被删除的三页内容不满足未被并入其他页面，出现将删除页文字挪入保留页的情况",
        bool(moved), "挪入保留页的删除页文案%d处：%s"
        % (len(moved), '；'.join(moved[:4]) or '无'))

    # ==================== 汇总 ====================
    order = list(_MAX_DELTAS_ORDER) + list(_PENALTY_DELTAS_ORDER)
    if len(results) != len(order):
        raise RuntimeError("评分项数量与满分表不一致：results=%d, order=%d"
                           % (len(results), len(order)))
    items = []
    total_score = 0
    for text, max_delta in zip(results, order):
        head, _, rule = text.partition('：')
        delta = int(head)
        total_score += delta
        items.append({
            "rule": rule,
            "max_delta": max_delta,
            "delta": delta,
            "hit": (delta != 0),
        })
    result["dim2_items"] = items
    result["total_score"] = max(0, total_score)
    result["max_score"] = MAX_SCORE_TOTAL
    return result


if __name__ == "__main__":
    target = sys.argv[1] if len(sys.argv) > 1 else os.path.dirname(
        os.path.abspath(__file__))
    print(json.dumps(evaluate(target), ensure_ascii=False, indent=2))
