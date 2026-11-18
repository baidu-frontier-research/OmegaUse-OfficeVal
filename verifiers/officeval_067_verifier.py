#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# pyright: reportUnknownMemberType=false, reportUnknownVariableType=false, reportUnknownArgumentType=false, reportMissingImports=false
"""
067 —— PPTX 幻灯片切换动画完成度校验器

指令：修改一下PPT，每页设置不同的切换动画，但是开头结尾要一样

说明：按接口统一约定，本模块不修改全局 sys.stdout，也不 print 主结果。
主结果统一由 evaluate(dir_path) -> dict 返回；入参为"脚本所在目录的路径"，
由脚本自己在该目录内定位并打开被评估的 .pptx。本文件仅在 __main__ 下作为
本地调试入口，才把 dict 序列化为 JSON 打印。

判定一律不依赖任何外部基准文件：细则里写明的每一处版式常量（页数、页序锚点、
逐页形状数、底色、配色、字体、字号、文案、母版/版式/备注/标签部件数量、
文档属性）都以常量形式内联在本文件中，脚本只读被评估的那一个 .pptx。
"""

import os
import re
import sys
import json
import zipfile
from xml.etree import ElementTree as ET

SCRIPT_ID = "067"
MAX_SCORE_TOTAL = 38

# 加分项满分序列，顺序与 results 的 append 顺序严格一致
_MAX_DELTAS_ORDER = [
    5,   # 01 全部35页均在幻灯片自身设置了切换效果
    5,   # 02 第2-34页共33张的切换配置两两不重复
    5,   # 03 第2-34页所用切换效果类型不少于18种
    5,   # 04 增强类切换效果不少于5种且均有 mc:Fallback 兜底
    5,   # 05 第1页与第35页切换效果四项全部一致
    3,   # 06 首尾所用效果未在第2-34页再次出现
    1,   # 07 首尾两页所选效果为平缓、庄重类
    3,   # 08 明确持续时长并分出快慢档次
    3,   # 09 五张章节过渡页切换互不相同且异于紧邻正文页
    3,   # 10 全部为效果库可识别可替换的标准效果
]
assert sum(_MAX_DELTAS_ORDER) == MAX_SCORE_TOTAL

# 减分项扣分序列，顺序与 results 中减分项的 append 顺序严格一致
_PENALTY_DELTAS_ORDER = [-5] * 6 + [-3] * 11 + [-1] * 6

# ============================ 命名空间 ============================
# 与办公软件对齐：PowerPoint / WPS 的"切换"选项卡写在幻灯片部件里，
#   经典效果位于 PresentationML 主命名空间，2010 以后的增强效果分别挂在
#   p14 / p15 / p159 扩展命名空间下，并由 mc:AlternateContent 做版本回退。
NS_P = 'http://schemas.openxmlformats.org/presentationml/2006/main'
NS_A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
NS_MC = 'http://schemas.openxmlformats.org/markup-compatibility/2006'
NS_P14 = 'http://schemas.microsoft.com/office/powerpoint/2010/main'
NS_P15 = 'http://schemas.microsoft.com/office/powerpoint/2012/main'
NS_P159 = 'http://schemas.microsoft.com/office/powerpoint/2015/09/main'
EXT_NS = (NS_P14, NS_P15, NS_P159)

# ECMA-376 经典切换效果（"切换"选项卡效果库中 2007 版即有的那批）
LEGACY_EFFECTS = {
    'blinds', 'checker', 'circle', 'comb', 'cover', 'cut', 'diamond',
    'dissolve', 'fade', 'newsflash', 'plus', 'pull', 'push', 'random',
    'randomBar', 'split', 'strips', 'wedge', 'wheel', 'wipe', 'zoom',
    'uncover', 'wheelReverse',
}

# PowerPoint 2010 以后新增的增强类切换效果（细则第4项所指的那一批）
EXT_EFFECTS = {
    'morph', 'ripple', 'prism', 'doors', 'window', 'glitter', 'honeycomb',
    'flash', 'vortex', 'ferris', 'gallery', 'conveyor', 'switch', 'flip',
    'reveal', 'shred', 'warp', 'wind', 'prestige', 'fracture', 'crush',
    'peelOff', 'pageCurlDouble', 'pageCurlSingle', 'airplane', 'origami',
    'drape', 'curtains', 'rotate', 'orbit', 'cube', 'box', 'pan',
}
STANDARD_EFFECTS = LEGACY_EFFECTS | EXT_EFFECTS

# 平缓、庄重类效果（细则第7项：与开场页/致谢页场景相称）
CALM_EFFECTS = {
    'morph', 'fade', 'dissolve', 'wipe', 'push', 'pull', 'cover',
    'uncover', 'circle', 'wedge', 'curtains', 'drape',
    'pageCurlDouble', 'pageCurlSingle',
}

# 必须带方向/选项参数才算"参数完整可读"的效果
DIR_REQUIRED = {'strips': ('dir',), 'split': ('orient', 'dir')}

# ======================= 版式常量（细则内联） =======================
SLIDE_TOTAL = 35
SLD_CX, SLD_CY = 12191365, 6858000        # 33.87cm × 19.05cm，16:9
BOUND_TOL = 120000                        # 约 0.33cm 的溢出容差

DARK_BG = '0F3554'
LIGHT_BG = 'F6FAFC'
DARK_PAGES = (1, 3, 11, 16, 28, 32, 35)   # 封面 + 五张章节过渡页 + 致谢
SECTION_PAGES = (3, 11, 16, 28, 32)
CONTENT_PAGES = tuple(i for i in range(1, SLIDE_TOTAL + 1) if i not in DARK_PAGES)

# 逐页形状数（细则"任一页的形状构成发生增减"的完整落地口径）
SHAPE_COUNTS = (
    6, 20, 13, 33, 36, 22, 27, 15, 27, 32, 12, 33, 25, 38, 23, 13, 38,
    75, 25, 38, 34, 27, 37, 26, 52, 43, 35, 10, 19, 25, 22, 8, 19, 15, 3,
)

# 全篇配色体系（每色在35页内的出现次数下限）
PALETTE_MIN = {
    'FFFFFF': 200, '1E2F3B': 150, 'F39B4A': 150, '2E8D97': 80,
    '103A5E': 60, 'AFC3CC': 50, 'D2E2E9': 50,
}

# 字体：中文微软雅黑 + 西文 Arial（"+mn-*"/"+mj-*" 为主题字体占位，属正常写法）
ALLOWED_FONTS = {'微软雅黑', 'Arial'}
THEME_FONT_RE = re.compile(r'^\+(mj|mn)-(ea|lt|cs)$')

# 字号锚点（DrawingML 中 sz 为百分之一磅）
SZ_COVER_TITLE = 3100      # 封面主标题 31 磅
SZ_SECTION_NUM = 4400      # 章节过渡页序号大字 44 磅
SZ_SECTION_TITLE = 3100    # 章节过渡页 Part 标题 31 磅
SZ_THANKS = 3300           # 致谢页主字 33 磅
SZ_SUBSECTION = 2300       # 正文页小节标题 23 磅
SZ_BODY_MIN, SZ_BODY_MAX = 640, 1500

# 封面页六处文案（前三处为主标题的三行分行）
COVER_TITLE_LINES = ('智慧园区背景下', '再生水回用管网更新项目', '风险因素识别与治理策略研究')
COVER_OTHER_TEXTS = (
    '毕业论文预答辩',
    '答辩人：周昱辰',
    '指导教师：陈若澜',
    'Urban Reclaimed Water Network Renewal · Risk Governance',
)
# 致谢页三处文案（前两处为同一形状的两行分行）
THANKS_LINES = ('敬请各位评审老师', '批评指正')
THANKS_OTHER = ('Thank you',)

# 目录页文案：标题 + 01-05 五个条目
TOC_TEXTS = (
    '目录', 'Contents', '01', '02', '03', '04', '05',
    '研究概述', '利益相关者识别及风险因素分析', '风险网络分析及关键因素检验',
    '治理策略分析及模型构建', '结论与展望',
)

# 五张章节过渡页的序号与 Part 标题
SECTION_TITLES = {
    3: ('01', '研究概述'),
    11: ('02', '利益相关者识别及风险因素分析'),
    16: ('03', '风险网络分析及关键因素检验'),
    28: ('04', '治理策略分析及模型构建'),
    32: ('05', '结论与展望'),
}

# 小节标题的编号与中英对照体系：页码 -> (中文标题, 英文副标题)
HEADINGS = {
    4: ('1.1 研究背景', 'Research background'),
    5: ('1.2 研究现状', 'Research status'),
    6: ('1.3 研究目的、内容和意义', 'Purpose & value'),
    7: ('1.4 研究内容', 'Research scope'),
    8: ('1.5 研究意义', 'Theoretical & practical meaning'),
    9: ('1.6 研究方法', 'Methods'),
    10: ('1.7 技术路线', 'Technical route'),
    12: ('2.1 利益相关者分析', 'Stakeholder map'),
    13: ('2.2 风险因素初始清单', 'Initial factor pool'),
    14: ('2.3 访谈修订后的因素编码', 'Revised coding system'),
    15: ('2.4 调查问卷设计', 'Questionnaire design'),
    17: ('3.1 风险网络可视化分析', 'Network visualization'),
    18: ('3.2 块模型分析', 'Block model'),
    19: ('3.3 指标分析框架', 'Metric framework'),
    20: ('3.4 点度中心度分析', 'Degree centrality'),
    21: ('3.5 接近中心度分析', 'Closeness centrality'),
    22: ('3.6 节点中介中心度分析', 'Betweenness centrality'),
    23: ('3.7 链中介中心度分析', 'Link betweenness'),
    24: ('3.8 熵权-TOPSIS 法确定关键因素', 'Entropy weight TOPSIS'),
    25: ('3.9 关键风险因素排序', 'Key risk ranking'),
    26: ('3.10 关键因素检验', 'Validation by node removal'),
    27: ('3.11 网络指标检验结果', 'Validation metrics'),
    29: ('4.1 分阶段治理策略', 'Stage-based actions'),
    30: ('4.2 利益相关者视角的治理策略', 'Stakeholder-oriented actions'),
    31: ('4.3 治理策略模型构建', 'Governance model'),
    33: ('5.1 主要工作结论', 'Main conclusions'),
    34: ('5.2 研究不足与展望', 'Limitations & outlook'),
}

# 既有页序的逐页锚点文本（用于"页面顺序未被调序或重排"的判定）
ORDER_ANCHORS = {1: '智慧园区背景下', 2: '目录', 35: 'Thank you'}
for _n, (_zh, _en) in HEADINGS.items():
    ORDER_ANCHORS[_n] = _zh
for _n, (_num, _title) in SECTION_TITLES.items():
    ORDER_ANCHORS[_n] = _title

PUBLIC_MARK = '公开交流版'

# 附属部件规模
MASTER_COUNT = 1
LAYOUT_COUNT = 2
NOTES_COUNT = 35
TAG_COUNT = 23
TAG_NAME = 'KSO_WM_DIAGRAM_VIRTUALLY_FRAME'
TAG_HOST_PAGE = 10          # 技术路线页承载 23 处 WPS 图示框标签

# 文件标识信息
FILE_STEM_KEY = '智慧园区再生水管网'
BAD_NAME_TOKENS = (
    '更新版', '修改版', '最终版', '终版', '新版', '论文预答辩', '预答辩',
    '切换动画', '动画版', '切换版', '副本', 'copy', 'final', 'new', 'v2',
)
DOC_TITLE = '智慧园区再生水回用管网更新项目风险因素识别与治理策略研究'
DOC_SUBJECT = '论文预答辩'
DOC_CREATOR = 'OpenAI'
DOC_LASTMOD = '张增慧'

# 允许与演示文稿同目录存在、不算"额外交付文件"的辅助文件
ALLOWED_SIDECAR_EXT = ('.py', '.pyc')
ALLOWED_SIDECAR_DIR = ('__pycache__',)


# ============================ 基础工具 ============================
def _ln(tag):
    """取 XML 标签的本地名。"""
    return tag.split('}')[-1] if '}' in tag else tag


def _uri(tag):
    """取 XML 标签的命名空间 URI。"""
    return tag.split('}')[0][1:] if tag.startswith('{') else ''


def _q(uri, name):
    return '{%s}%s' % (uri, name)


class _Package(object):
    """按 OPC 包读取 .pptx；只读，不改写被评估文件。"""

    def __init__(self, path):
        self.zf = zipfile.ZipFile(path)
        self.names = list(self.zf.namelist())
        self.nameset = set(self.names)

    def close(self):
        try:
            self.zf.close()
        except Exception:
            pass

    def has(self, part):
        return part in self.nameset

    def read(self, part):
        return self.zf.read(part)

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


def _transition_record(tr):
    """把一个 <p:transition> 解析为可比较的记录。

    与办公软件对齐：
      · 效果名 = <p:transition> 下唯一的效果子节点本地名；p15 的增强效果统一
        写作 <p15:prstTrans prst="airplane"/>，故以 @prst 作为效果名；
      · 方向/选项 = 该效果子节点上的属性（dir / orient / spokes / thruBlk /
        isInverted / option 等），即"切换"选项卡里"效果选项"下拉的取值；
      · 速度 = @spd（slow/med/fast），时长 = @p14:dur（毫秒）；
      · @advTm 为自动换片时间，<p:sndAc> 为切换声音。
    """
    rec = {
        'spd': tr.get('spd'), 'dur': None,
        'advTm': tr.get('advTm'), 'advClick': tr.get('advClick'),
        'snd': False, 'effect': None, 'ns': '', 'opts': {}, 'children': 0,
    }
    dur = tr.get(_q(NS_P14, 'dur'))
    if dur is not None and dur.lstrip('-').isdigit():
        rec['dur'] = int(dur)
    for ch in tr:
        lname = _ln(ch.tag)
        if lname == 'sndAc':
            rec['snd'] = True
            continue
        if lname == 'extLst':
            continue
        attrs = dict(ch.attrib)
        name = lname
        if lname == 'prstTrans':
            name = attrs.pop('prst', '') or 'prstTrans'
        rec['effect'] = name
        rec['ns'] = _uri(ch.tag)
        rec['opts'] = attrs
        rec['children'] += 1
    return rec


def _is_ext_effect(rec):
    """是否为 PowerPoint 2010 以后的增强类切换效果。"""
    if not rec or not rec.get('effect'):
        return False
    return rec['ns'] in EXT_NS or rec['effect'] in EXT_EFFECTS


def _sig(rec):
    """效果 + 方向/选项 的配置签名。"""
    if not rec or not rec.get('effect'):
        return None
    return (rec['effect'], tuple(sorted(rec['opts'].items())))


def _full_sig(rec):
    """效果 + 方向/选项 + 速度 + 时长 四项签名。"""
    s = _sig(rec)
    if s is None:
        return None
    return (s, rec.get('spd'), rec.get('dur'))


def _slide_transitions(root):
    """扫描 <p:sld> 的直接子节点，取出切换配置。

    返回 (clrMapOvr 的位置索引, [切换条目...])。切换条目形如
      {'pos': 子节点序号, 'wrapped': 是否被 mc:AlternateContent 包裹,
       'requires': mc:Choice/@Requires, 'choice': 主用效果记录,
       'fallback': mc:Fallback 内的兜底效果记录}
    """
    children = list(root)
    clr = -1
    for i, c in enumerate(children):
        if _ln(c.tag) == 'clrMapOvr':
            clr = i
            break
    out = []
    for i, c in enumerate(children):
        if _ln(c.tag) == 'transition' and _uri(c.tag) == NS_P:
            out.append({'pos': i, 'wrapped': False, 'requires': None,
                        'choice': _transition_record(c), 'fallback': None})
        elif _ln(c.tag) == 'AlternateContent' and _uri(c.tag) == NS_MC:
            choice = c.find(_q(NS_MC, 'Choice'))
            fallback = c.find(_q(NS_MC, 'Fallback'))
            ci = fi = None
            req = None
            if choice is not None:
                req = choice.get('Requires')
                t = choice.find(_q(NS_P, 'transition'))
                if t is not None:
                    ci = _transition_record(t)
            if fallback is not None:
                t = fallback.find(_q(NS_P, 'transition'))
                if t is not None:
                    fi = _transition_record(t)
            if ci is not None or fi is not None:
                out.append({'pos': i, 'wrapped': True, 'requires': req,
                            'choice': ci if ci is not None else fi,
                            'fallback': fi if ci is not None else None})
    return clr, out


def _shape_texts(root):
    """按形状返回文本段落列表：[[段1, 段2, ...], ...]，只取有文字的形状。"""
    out = []
    for sp in root.iter(_q(NS_P, 'sp')):
        tx = sp.find(_q(NS_P, 'txBody'))
        if tx is None:
            continue
        paras = []
        for para in tx.findall(_q(NS_A, 'p')):
            s = ''.join(t.text or '' for t in para.iter(_q(NS_A, 't')))
            if s.strip():
                paras.append(s.strip())
        if paras:
            out.append(paras)
    return out


def _shape_count(root):
    """整页形状数量（<p:sp>，含组内形状）。"""
    return len(list(root.iter(_q(NS_P, 'sp'))))


def _bg_color(root):
    """整页底色 srgbClr；取 <p:bg> 内第一个 srgbClr。"""
    bg = root.find('%s/%s' % (_q(NS_P, 'cSld'), _q(NS_P, 'bg')))
    if bg is None:
        return None
    for el in bg.iter(_q(NS_A, 'srgbClr')):
        v = el.get('val')
        if v:
            return v.upper()
    return None


def _font_sizes(root):
    """整页出现过的字号集合（sz 属性，百分之一磅）。"""
    out = set()
    for el in root.iter():
        v = el.get('sz')
        if v is None:
            continue
        if _ln(el.tag) not in ('rPr', 'defRPr', 'endParaRPr'):
            continue
        if v.isdigit():
            out.add(int(v))
    return out


def _shape_boxes(root):
    """返回所有显式设定了 xfrm 的形状包围盒 [(x, y, cx, cy), ...]。"""
    boxes = []
    for sp in root.iter(_q(NS_P, 'sp')):
        spPr = sp.find(_q(NS_P, 'spPr'))
        if spPr is None:
            continue
        xfrm = spPr.find(_q(NS_A, 'xfrm'))
        if xfrm is None:
            continue
        off = xfrm.find(_q(NS_A, 'off'))
        ext = xfrm.find(_q(NS_A, 'ext'))
        if off is None or ext is None:
            continue
        try:
            boxes.append((int(off.get('x')), int(off.get('y')),
                          int(ext.get('cx')), int(ext.get('cy'))))
        except (TypeError, ValueError):
            continue
    return boxes


def _flat(texts):
    """把 _shape_texts 的结果拍平成一个整页文本串。"""
    return '\n'.join('\n'.join(paras) for paras in texts)


# ============================ 对外接口 ============================
def evaluate(dir_path):
    """入参为目录路径，在该目录内定位 .pptx 并给出完成度评分。"""
    result = {
        "id": SCRIPT_ID,
        "file_name": "",
        "status": "ok",
        "error": "",
        "dim1_pass": "",
        "dim1_reason": "",
        "dim2_items": [],
        "total_score": 0,
        "max_score": MAX_SCORE_TOTAL,
    }
    try:
        if not dir_path or not os.path.isdir(dir_path):
            raise FileNotFoundError(f"目录不存在：{dir_path}")
        cands = sorted(f for f in os.listdir(dir_path)
                       if f.lower().endswith('.pptx') and not f.startswith('~$'))
        if not cands:
            raise FileNotFoundError(f"目录中未找到 .pptx 文件：{dir_path}")
        file_path = os.path.join(dir_path, cands[0])
        result["file_name"] = os.path.basename(file_path)
        return _evaluate_impl(file_path, dir_path, result)
    except FileNotFoundError as exc:
        # 找不到目录 / 找不到待评估文件：两个维度都没能开始判定，
        # dim1_pass 与 dim1_reason 一律留空，满分基数也归零。
        return _error_result(result, str(exc))
    except Exception as exc:
        return _error_result(result, f"{type(exc).__name__}: {exc}")


def _error_result(result, message):
    """未能完成判定时的统一返回：dim1 字段留空、分数与满分均为 0。"""
    result["status"] = "error"
    result["error"] = message
    result["dim1_pass"] = ""
    result["dim1_reason"] = ""
    result["dim2_items"] = []
    result["total_score"] = 0
    result["max_score"] = 0
    return result


def _dim1_gate(file_path):
    """维度1：交付文件为1个可正常打开的 PowerPoint 演示文稿（.pptx）。

    与办公软件对齐：PowerPoint / WPS 打开 .pptx 时先校验 OPC 包结构与
      [Content_Types].xml，再逐部件解析 XML；任一环节失败即"文件损坏无法打开"
      或弹出"发现不可读取的内容"。此处按同一条链路做检查：
      ① 是 zip（OPC）包；② 有 [Content_Types].xml 与 ppt/presentation.xml；
      ③ presentation.xml、每张幻灯片、母版、版式的 XML 均可解析；
      ④ sldIdLst 指向的幻灯片部件都真实存在。
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
        # 维度一不通过：dim1_pass 为 False，dim1_reason 给出不通过的错误摘要。
        result["dim1_pass"] = False
        result["dim1_reason"] = "维度一不通过：%s" % reason
        result["dim2_items"] = []
        result["total_score"] = 0
        result["max_score"] = MAX_SCORE_TOTAL
        return result
    result["dim1_pass"] = True
    result["dim1_reason"] = "演示文稿可正常打开，包结构与各部件 XML 均可解析"

    try:
        return _score(pkg, file_path, dir_path, result)
    finally:
        pkg.close()


def _score(pkg, file_path, dir_path, result):
    results = []      # ["+N：规则说明", ...]，与 _MAX_DELTAS_ORDER 逐项对齐

    def emit(delta, rule, detail=""):
        # detail 仅作为判定过程中的现场说明，不再随结果输出。
        results.append("%+d：%s" % (delta, rule))

    # ---------------- 数据采集 ----------------
    slide_parts = pkg.slide_parts()
    n_slides = len(slide_parts)
    roots, raws = {}, {}
    for idx, part in enumerate(slide_parts, 1):
        raws[idx] = pkg.text(part)
        roots[idx] = pkg.xml(part)

    pages = range(1, n_slides + 1)
    clrmap = {}
    trans = {}                # 页码 -> 切换条目（取第一条）
    trans_all = {}            # 页码 -> 全部切换条目
    for i in pages:
        clr, items = _slide_transitions(roots[i])
        clrmap[i] = clr
        trans_all[i] = items
        trans[i] = items[0] if items else None

    def eff(i):
        t = trans.get(i)
        return t['choice']['effect'] if (t and t['choice']) else None

    def rec(i):
        t = trans.get(i)
        return t['choice'] if t else None

    mid = [i for i in range(2, min(34, n_slides) + 1)]
    mid_recs = [rec(i) for i in mid]
    mid_names = set(r['effect'] for r in mid_recs if r and r['effect'])
    mid_sigs = [_sig(r) for r in mid_recs]

    texts = {i: _shape_texts(roots[i]) for i in pages}
    flats = {i: _flat(texts[i]) for i in pages}

        # ==================== 维度2：加分项 ====================
    # +5：全部35页均在幻灯片自身设置了切换效果（写在每页颜色映射节点之后），
    #     逐页在"切换"选项卡可查、顺序放映时35页均实际触发，不存在任何一页为
    #     "无切换"或仅依赖母版/版式继承
    # 严格对齐细则，仅踩细则明确写出的点：
    #   ① 35 页：sldIdLst 内的幻灯片总数为 35；
    #   ② 在幻灯片自身设置：<p:transition> 位于该页 <p:sld> 的直接子层（含被
    #      mc:AlternateContent 包裹的写法），而不是继承母版/版式；
    #   ③ 写在颜色映射节点之后：节点位置索引大于 <p:clrMapOvr> 的索引，符合
    #      CT_Slide 的 cSld → clrMapOvr → transition → timing 顺序；
    #   ④ 不为"无切换"：<p:transition> 下必须挂着效果子节点（空节点在放映时
    #      等同于"无"）。
    # 与办公软件对齐：PowerPoint/WPS 的"切换"选项卡逐页读的就是本页
    #   <p:transition>；节点顺序不合规时会被判为不可读内容或直接忽略。
    a1_hit = 0
    for i in pages:
        t = trans.get(i)
        if not t or not t['choice'] or not t['choice']['effect']:
            continue
        if t['choice']['children'] < 1:
            continue
        if clrmap.get(i, -1) >= 0 and t['pos'] <= clrmap[i]:
            continue
        a1_hit += 1
    a1_ok = (n_slides == SLIDE_TOTAL and a1_hit == SLIDE_TOTAL)
    emit(5 if a1_ok else 0,
         "全部35页均在幻灯片自身设置了切换效果（写在颜色映射节点之后），逐页可查、放映均实际触发，无一页为无切换或仅靠母版继承",
         "共%d页，页内切换命中 %d/%d 页" % (n_slides, a1_hit, SLIDE_TOTAL))

    # +5：第2页至第34页共33张页面之间的切换配置两两不重复，同一效果族在多页
    #     复用时以不同方向/选项加以区分，不存在两页使用完全相同的
    #     "效果+方向/选项"组合
    # 严格对齐细则：签名 = 效果名 + 该效果子节点上的全部方向/选项属性；
    #   33 张页面必须给出 33 个互不相同的签名（任一页缺切换即视为重复失败）。
    a2_total = len(mid)
    a2_valid = [s for s in mid_sigs if s is not None]
    a2_ok = (a2_total == 33 and len(a2_valid) == 33 and len(set(a2_valid)) == 33)
    emit(5 if a2_ok else 0,
         "第2页至第34页共33张页面之间的切换配置两两不重复，同一效果族以不同方向/选项区分，不存在完全相同的\"效果+方向/选项\"组合",
         "第2-34页有效配置 %d 个，去重后 %d 个" % (len(a2_valid), len(set(a2_valid))))

    # +5：第2页至第34页所用的切换效果类型（按效果名计）不少于18种
    # 严格对齐细则：按效果名去重计数，同一效果仅换方向/选项不另计一种；
    #   门槛为 ≥18 种，靠少数效果反复换方向凑页数的（<15 种）本项不得分。
    a3_ok = (len(mid_names) >= 18)
    emit(5 if a3_ok else 0,
         "第2页至第34页所用的切换效果类型（按效果名计）不少于18种，覆盖淡化、擦除、推入、覆盖、劈裂、百叶窗、棋盘、梳理、随机线条、轮辐、条纹、形状等多个效果族",
         "效果类型 %d 种：%s" % (len(mid_names), '、'.join(sorted(mid_names))))

        # +5：使用了平滑(morph)、涟漪、棱柱、门、页面卷曲、飞机、剥离、闪光等
    #     PowerPoint 2010 以后的增强类切换效果不少于5种，且每处增强效果均以
    #     兼容性容器（mc:AlternateContent 的 mc:Choice/mc:Fallback）包裹并配置
    #     经典效果兜底，保证在不支持该效果的版本或 WPS 中放映时不会退化为无切换
    # 严格对齐细则，仅踩细则明确写出的点：
    #   ① 增强类效果种数（按效果名去重）≥5；
    #   ② 每一处增强效果都在 mc:AlternateContent 里，且 mc:Choice 声明了
    #      Requires（p14/p15/p159）；
    #   ③ 同容器内必须有 mc:Fallback，且其中 <p:transition> 挂的是经典效果，
    #      这样低版本 / WPS 才会退化为经典切换而不是"无切换"。
    ext_names = set()
    ext_wrap_bad = []
    for i in pages:
        t = trans.get(i)
        if not t or not _is_ext_effect(t['choice']):
            continue
        ext_names.add(t['choice']['effect'])
        fb = t.get('fallback')
        ok = bool(t.get('wrapped')) and bool(t.get('requires'))
        ok = ok and fb is not None and fb.get('effect') in LEGACY_EFFECTS
        if not ok:
            ext_wrap_bad.append(i)
    a4_ok = (len(ext_names) >= 5 and not ext_wrap_bad)
    emit(5 if a4_ok else 0,
         "使用了平滑、涟漪、棱柱、门、页面卷曲、飞机、剥离、闪光等 PowerPoint 2010 以后的增强类切换效果不少于5种，且每处均以 mc:AlternateContent 的 Choice/Fallback 包裹并配置经典效果兜底",
         "增强效果 %d 种：%s；未做兜底的页：%s"
         % (len(ext_names), '、'.join(sorted(ext_names)) or '无',
            ext_wrap_bad or '无'))

    # +5：第1页封面页与第35页致谢页使用完全相同的切换效果，效果名、方向/选项、
    #     速度与时长四项全部一致，构成首尾呼应
    # 严格对齐细则：比较 (效果名, 方向/选项属性集, @spd, @p14:dur) 四元签名；
    #   两页都必须真的有切换，四项逐一相等才算一致。
    first_rec, last_rec = rec(1), rec(n_slides)
    fs, ls = _full_sig(first_rec), _full_sig(last_rec)
    a5_ok = (n_slides == SLIDE_TOTAL and fs is not None and fs == ls)
    emit(5 if a5_ok else 0,
         "第1页封面页与第35页致谢页使用完全相同的切换效果，效果名、方向/选项、速度与时长四项全部一致",
         "首页=%s / 末页=%s" % (fs, ls))

    # +3：第1页与第35页所用的这一种切换效果未在第2页至第34页中再次出现
    # 严格对齐细则：按效果名判断唯一性（首尾用的这一种效果名不出现在 2-34 页）。
    head_names = set(x for x in (eff(1), eff(n_slides)) if x)
    a6_ok = bool(head_names) and not (head_names & mid_names)
    emit(3 if a6_ok else 0,
         "第1页与第35页所用的这一种切换效果未在第2页至第34页中再次出现，使首尾切换在整套放映中具有唯一性",
         "首尾效果=%s；与中间页重复的=%s"
         % ('、'.join(sorted(head_names)) or '无',
            '、'.join(sorted(head_names & mid_names)) or '无'))

        # +1：首尾两页所选效果为平缓、庄重类（如平滑、淡入淡出、淡出至黑），未使用
    #     蜂巢、爆炸、随机等跳脱效果
    # 严格对齐细则：首尾两页的效果名都必须落在平缓/庄重白名单内；
    #   蜂巢(honeycomb)、涡流(vortex)、随机(random)、闪光(glitter) 等不在白名单。
    a7_ok = bool(head_names) and head_names.issubset(CALM_EFFECTS)
    emit(1 if a7_ok else 0,
         "首尾两页所选效果为平缓、庄重类（如平滑、淡入淡出、淡出至黑），与开场页和致谢页场景相称，未使用蜂巢、爆炸、随机等跳脱效果",
         "首尾效果=%s" % ('、'.join(sorted(head_names)) or '无'))

    # +3：为切换设置了明确的持续时长（不少于30页带有时长参数、且落在0.5-2.0秒
    #     区间），并分出快慢档次——首尾两页取1.2-1.5秒的较慢档，全篇至少出现
    #     3种不同的时长取值，而非35页全部沿用同一档默认速度
    # 严格对齐细则，仅踩细则明确写出的三点：
    #   ① 带时长参数且落在 [500, 2000] 毫秒的页数 ≥30；
    #   ② 第1页与第35页的时长落在 [1200, 1500] 毫秒的较慢档；
    #   ③ 全篇有效时长取值去重后 ≥3 种（不是 35 页共用一档默认速度）。
    # 与办公软件对齐："切换"选项卡右侧的"持续时间"框写入 <p:transition
    #   p14:dur="1500">，单位毫秒；只写 @spd 属于沿用旧的三档默认速度。
    durs = [rec(i)['dur'] for i in pages if rec(i) and rec(i)['dur'] is not None]
    in_band = [d for d in durs if 500 <= d <= 2000]
    head_durs = [r['dur'] for r in (first_rec, last_rec) if r and r['dur'] is not None]
    a8_ok = (len(in_band) >= 30
             and len(head_durs) == 2
             and all(1200 <= d <= 1500 for d in head_durs)
             and len(set(in_band)) >= 3)
    emit(3 if a8_ok else 0,
         "为切换设置了明确的持续时长（不少于30页带时长参数且落在0.5-2.0秒），并分出快慢档次——首尾两页取1.2-1.5秒的较慢档，全篇至少3种不同时长取值",
         "带时长且在区间内 %d 页，时长取值 %d 种：%s；首尾时长=%s"
         % (len(in_band), len(set(in_band)),
            sorted(set(in_band)) or '无', head_durs or '无'))

    # +3：第3、11、16、28、32页五张章节过渡页所用切换效果彼此互不相同，且与
    #     各自紧邻的前后正文页不同
    # 严格对齐细则，仅踩细则明确写出的两点（按"效果+方向/选项"配置比较，
    #   同族换方向即算不同，这与"切换"选项卡里两页效果选项不同的观感一致）：
    #   ① 五张章节过渡页的配置两两不同；
    #   ② 每张章节过渡页的配置与其前一页、后一页的配置都不同。
    sec_sigs = [_sig(rec(i)) for i in SECTION_PAGES]
    a9_ok = all(s is not None for s in sec_sigs) and len(set(sec_sigs)) == len(SECTION_PAGES)
    a9_bad = []
    if a9_ok:
        for i in SECTION_PAGES:
            for j in (i - 1, i + 1):
                if 1 <= j <= n_slides and _sig(rec(i)) == _sig(rec(j)):
                    a9_bad.append('第%d页与第%d页相同' % (i, j))
        a9_ok = not a9_bad
    emit(3 if a9_ok else 0,
         "第3、11、16、28、32页五张章节过渡页所用切换效果彼此互不相同，且与各自紧邻的前后正文页不同",
         "五页配置去重 %d/5；相邻冲突：%s"
         % (len(set(sec_sigs)), '；'.join(a9_bad) or '无'))

        # +3：35页所用切换全部为可在"切换"选项卡效果库中直接识别、选中并替换的标准
    #     效果，方向/选项等参数完整可读，不存在自造效果名、缺参数的空切换节点、
    #     未闭合的切换节点或未声明的命名空间前缀
    # 严格对齐细则，仅踩细则明确写出的四点：
    #   ① 标准效果：主用效果名与兜底效果名都在效果库名单内（无自造名）；
    #   ② 参数完整：效果子节点上的属性值非空，条纹/劈裂等必须带方向、劈裂还要
    #      带 orient，避免"缺参数"；
    #   ③ 非空切换节点：<p:transition> 下确实挂了效果子节点；
    #   ④ 未闭合节点 / 未声明前缀：XML 若有这两类问题，维度1 的逐部件解析已经
    #      抛错，能走到这里说明每张幻灯片都能被解析（此处再确认一遍已解析）。
    a10_bad = []
    for i in pages:
        t = trans.get(i)
        if not t or not t['choice'] or not t['choice']['effect']:
            a10_bad.append('第%d页无有效切换节点' % i)
            continue
        for tag, r in (('主用', t['choice']), ('兜底', t.get('fallback'))):
            if r is None:
                continue
            name = r.get('effect')
            if name not in STANDARD_EFFECTS:
                a10_bad.append('第%d页%s效果名"%s"不在效果库内' % (i, tag, name))
                continue
            if r.get('children', 1) < 1:
                a10_bad.append('第%d页%s为空切换节点' % (i, tag))
                continue
            for k, v in (r.get('opts') or {}).items():
                if v is None or str(v).strip() == '':
                    a10_bad.append('第%d页%s效果的%s参数为空' % (i, tag, k))
            for need in DIR_REQUIRED.get(name, ()):
                if need not in (r.get('opts') or {}):
                    a10_bad.append('第%d页%s效果"%s"缺%s参数' % (i, tag, name, need))
    a10_ok = (a1_hit == SLIDE_TOTAL and not a10_bad)
    emit(3 if a10_ok else 0,
         "35页所用切换全部为效果库中可直接识别、选中并替换的标准效果，方向/选项参数完整可读，不存在自造效果名、空切换节点、未闭合节点或未声明的命名空间前缀",
         "异常项：%s" % ('；'.join(a10_bad[:6]) if a10_bad else '无'))

    if len(results) != len(_MAX_DELTAS_ORDER):
        raise RuntimeError("加分项条数不一致：results=%d vs order=%d"
                           % (len(results), len(_MAX_DELTAS_ORDER)))

        # ==================== 维度2：减分项 ====================
    # 减分项描述的是"违规事实"：命中（violated=True）即按该条扣分。
    def pen(delta, rule, violated, detail=""):
        # detail 仅作为判定过程中的现场说明，不再随结果输出。
        results.append("%s：%s" % (("%d" % delta) if violated else "-0", rule))

    # -5：交付稿的幻灯片总数不为35页，出现增页、删页或把某一页拆成多页
    p_count_bad = (n_slides != SLIDE_TOTAL)
    pen(-5, "交付稿的幻灯片总数不为35页，出现增页、删页或把某一页拆成多页",
        p_count_bad, "实际 %d 页" % n_slides)

    # -5：页面顺序不为"封面页→目录页→01 Part…→致谢页"的既有页序，出现调序或重排
    # 与办公软件对齐：放映顺序取自 presentation.xml 的 sldIdLst，逐页用该页
    #   唯一的锚点文案（封面主标题首行、目录、各 Part 标题、各小节编号标题、
    #   致谢页 Thank you）核对是否仍坐在原本的位次上。
    order_bad = []
    for i, anchor in sorted(ORDER_ANCHORS.items()):
        if i > n_slides or anchor not in flats.get(i, ''):
            order_bad.append('第%d页应为"%s"' % (i, anchor))
    pen(-5, "页面顺序不为\"封面页→目录页→01 Part 研究概述→1.1…1.7→02 Part→2.1…2.4→03 Part→3.1…3.11→04 Part→4.1…4.3→05 Part→5.1、5.2→致谢页\"的既有页序，出现调序或重排",
        bool(order_bad), "错位页：%s" % ('；'.join(order_bad[:6]) or '无'))

    # -5：任一页的形状构成发生增减（如目录页20个、技术路线页32个、块模型分析页
    #     75个、关键风险因素排序页52个、关键因素检验页43个形状的数量变化）
    shape_bad = []
    for i in pages:
        if i > len(SHAPE_COUNTS):
            shape_bad.append('第%d页为新增页' % i)
            continue
        got = _shape_count(roots[i])
        if got != SHAPE_COUNTS[i - 1]:
            shape_bad.append('第%d页 %d→%d' % (i, SHAPE_COUNTS[i - 1], got))
    pen(-5, "任一页的形状构成发生增减（如目录页20个、技术路线页32个、块模型分析页75个、关键风险因素排序页52个、关键因素检验页43个形状的数量变化），出现删形状、加装饰形状或拆分合并形状",
        bool(shape_bad), "形状数变化：%s" % ('；'.join(shape_bad[:6]) or '无'))

    # -5：任一页的文字被改写、删减或新增
    # 口径：把细则逐条写明的文案（封面六处、致谢三处、目录标题与五个条目、
    #   五张 Part 页的序号与标题、27 个小节的中文编号标题与英文副标题、
    #   28 张正文页的"公开交流版"标识）逐页核对，缺一即为改写/删减。
    text_bad = []
    expect = []
    for s in COVER_TITLE_LINES + COVER_OTHER_TEXTS:
        expect.append((1, s))
    for s in THANKS_LINES + THANKS_OTHER:
        expect.append((SLIDE_TOTAL, s))
    for s in TOC_TEXTS:
        expect.append((2, s))
    for i, (num, title) in SECTION_TITLES.items():
        expect.append((i, num))
        expect.append((i, title))
    for i, (zh, en) in HEADINGS.items():
        expect.append((i, zh))
        expect.append((i, en))
    for i in CONTENT_PAGES:
        expect.append((i, PUBLIC_MARK))
    for i, s in expect:
        if i > n_slides or s not in flats.get(i, ''):
            text_bad.append('第%d页缺"%s"' % (i, s))
    pen(-5, "任一页的文字被改写、删减或新增，如把\"3.8 熵权-TOPSIS 法确定关键因素\"改为其他表述、删掉小节的英文副标题、替换研究结论表述或补写未经要求的说明文字",
        bool(text_bad), "缺失/改写：%s" % ('；'.join(text_bad[:6]) or '无'))

        # -5：为页内元素添加了未被要求的进入/强调/退出动画或动画时间轴
    # 与办公软件对齐："动画"选项卡里的进入/强调/退出效果一律落在幻灯片的
    #   <p:timing> 动画时间轴里（<p:par>/<p:seq>/<p:anim*>/<p:animEffect>），
    #   与"切换"选项卡的 <p:transition> 是两套互不相干的节点。
    anim_bad = []
    for i in pages:
        raw = raws[i]
        if '<p:timing' in raw or '<p:anim' in raw or '<p:animEffect' in raw:
            anim_bad.append(i)
    pen(-5, "为页内元素添加了未被要求的进入/强调/退出动画或动画时间轴，使放映时标题、色块、正文条目逐个出现，改变了整页一次呈现的放映方式",
        bool(anim_bad), "含动画时间轴的页：%s" % (anim_bad[:8] or '无'))

    # -5：任一页被整体转成一张位图、或被组合压平，导致其中的标题、色块、序号
    #     圆点、正文文本框不能单独点选编辑
    # 口径：① 页内出现覆盖整页 80% 以上面积的图片（<p:pic>）即视为整页位图；
    #       ② 页内不存在任何带文字的形状，也视为已被压平（文字进了图里）。
    flat_bad = []
    page_area = float(SLD_CX) * float(SLD_CY)
    for i in pages:
        for pic in roots[i].iter(_q(NS_P, 'pic')):
            ext = None
            for e in pic.iter(_q(NS_A, 'ext')):
                ext = e
                break
            if ext is None:
                continue
            try:
                area = int(ext.get('cx')) * int(ext.get('cy'))
            except (TypeError, ValueError):
                continue
            if page_area > 0 and area >= 0.8 * page_area:
                flat_bad.append('第%d页含满版位图' % i)
                break
        if not texts.get(i):
            flat_bad.append('第%d页无可单独编辑的文字形状' % i)
    pen(-5, "任一页被整体转成一张位图、或被组合压平，导致其中的标题、色块、序号圆点、正文文本框不能单独点选编辑",
        bool(flat_bad), "异常页：%s" % ('；'.join(flat_bad[:6]) or '无'))

    # -3：幻灯片页面尺寸不为33.87cm×19.05cm（12191365×6858000 EMU，16:9）
    cx, cy = pkg.slide_size()
    size_bad = (cx != SLD_CX or cy != SLD_CY)
    pen(-3, "幻灯片页面尺寸不为33.87cm×19.05cm（12191365×6858000 EMU，16:9），被改为4:3或其他自定义尺寸",
        size_bad, "实际 %sx%s EMU" % (cx, cy))

    # -3：页面底色被改动——封面页、五张章节过渡页与致谢页共7页的深蓝底色不为
    #     0F3554，或28张正文页的浅色底不为 F6FAFC
    bg_bad = []
    for i in pages:
        got = _bg_color(roots[i])
        want = DARK_BG if i in DARK_PAGES else LIGHT_BG
        if got != want:
            bg_bad.append('第%d页 %s→%s' % (i, want, got))
    pen(-3, "页面底色被改动——封面页、五张章节过渡页与致谢页共7页的深蓝底色不为 0F3554，或28张正文页的浅色底不为 F6FAFC",
        bool(bg_bad), "底色异常：%s" % ('；'.join(bg_bad[:6]) or '无'))

        # -3：全篇配色不再以橙 F39B4A、青 2E8D97、深蓝 103A5E 与 1E2F3B、浅灰蓝
    #     AFC3CC 与 D2E2E9、白 FFFFFF 为体系，出现整体换色或色块被统一改色
    color_hits = {}
    for i in pages:
        for v in re.findall(r'srgbClr val="([0-9A-Fa-f]{6})"', raws[i]):
            key = v.upper()
            color_hits[key] = color_hits.get(key, 0) + 1
    palette_bad = ['%s(%d<%d)' % (k, color_hits.get(k, 0), v)
                   for k, v in sorted(PALETTE_MIN.items()) if color_hits.get(k, 0) < v]
    pen(-3, "全篇配色不再以橙 F39B4A、青 2E8D97、深蓝 103A5E 与 1E2F3B、浅灰蓝 AFC3CC 与 D2E2E9、白 FFFFFF 为体系，出现整体换色或色块被统一改色",
        bool(palette_bad), "配色缺口：%s" % ('；'.join(palette_bad) or '无'))

    # -3：字体不为"中文微软雅黑＋西文 Arial"的搭配，被替换为宋体、等线或其他字体
    # 与办公软件对齐：DrawingML 里 latin/ea/cs 的 @typeface 即"字体"下拉的取值；
    #   "+mn-ea""+mj-lt"这类是主题字体占位符，属正常写法不算被换字体。
    used_fonts = set()
    for i in pages:
        used_fonts.update(re.findall(r'typeface="([^"]*)"', raws[i]))
    font_bad = sorted(f for f in used_fonts
                      if f and f not in ALLOWED_FONTS and not THEME_FONT_RE.match(f))
    font_bad = font_bad or ([] if '微软雅黑' in used_fonts else ['缺微软雅黑'])
    pen(-3, "字体不为\"中文微软雅黑＋西文 Arial\"的搭配，被替换为宋体、等线或其他字体",
        bool(font_bad), "越界字体：%s" % ('、'.join(font_bad) or '无'))

    # -3：各级文字字号层级被整体放大或缩小，封面主标题31磅、章节过渡页序号大字
    #     44磅与 Part 标题31磅、致谢页主字33磅、28张正文页小节标题23磅、
    #     页内正文6.4-15磅的比例关系被破坏
    sz_bad = []
    if 1 <= n_slides:
        if SZ_COVER_TITLE not in _font_sizes(roots[1]):
            sz_bad.append('封面页缺31磅主标题')
    for i in SECTION_PAGES:
        if i > n_slides:
            continue
        s = _font_sizes(roots[i])
        if SZ_SECTION_NUM not in s or SZ_SECTION_TITLE not in s:
            sz_bad.append('第%d页缺44磅序号或31磅 Part 标题' % i)
    if n_slides >= SLIDE_TOTAL and SZ_THANKS not in _font_sizes(roots[SLIDE_TOTAL]):
        sz_bad.append('致谢页缺33磅主字')
    for i in CONTENT_PAGES:
        if i > n_slides:
            continue
        s = _font_sizes(roots[i])
        if SZ_SUBSECTION not in s:
            sz_bad.append('第%d页缺23磅小节标题' % i)
        body = [v for v in s if v != SZ_SUBSECTION]
        if body and (min(body) < SZ_BODY_MIN or max(body) > 2200):
            sz_bad.append('第%d页正文字号越界(%d-%d)' % (i, min(body), max(body)))
    pen(-3, "各级文字字号层级被整体放大或缩小，封面主标题31磅、章节过渡页序号大字44磅与 Part 标题31磅、致谢页主字33磅、28张正文页小节标题23磅、页内正文6.4-15磅的比例关系被破坏",
        bool(sz_bad), "字号异常：%s" % ('；'.join(sz_bad[:6]) or '无'))

        # -3：封面页与致谢页的文字未保持原有分行、措辞与摆放位置——封面六处、
    #     致谢页三处文案
    # 口径：① 封面主标题必须仍是"智慧园区背景下/再生水回用管网更新项目/
    #        风险因素识别与治理策略研究"三行、同一形状内按此顺序分行；
    #       ② 封面另外三个文本形状（毕业论文预答辩、答辩人/指导教师、英文副标题）
    #        齐备，封面共 4 个带文字的形状；
    #       ③ 致谢页"敬请各位评审老师/批评指正"仍为同一形状两行分行，
    #        "Thank you"独立成一个形状，致谢页共 2 个带文字的形状。
    cover_bad = []
    if n_slides >= 1:
        cover_shapes = texts.get(1, [])
        if len(cover_shapes) != 4:
            cover_bad.append('封面带文字形状 %d 个（应为4个）' % len(cover_shapes))
        if not any(tuple(p) == COVER_TITLE_LINES for p in cover_shapes):
            cover_bad.append('封面主标题三行分行被改动')
        for s in COVER_OTHER_TEXTS:
            if s not in flats.get(1, ''):
                cover_bad.append('封面缺"%s"' % s)
    if n_slides >= SLIDE_TOTAL:
        thanks_shapes = texts.get(SLIDE_TOTAL, [])
        if len(thanks_shapes) != 2:
            cover_bad.append('致谢页带文字形状 %d 个（应为2个）' % len(thanks_shapes))
        if not any(tuple(p) == THANKS_LINES for p in thanks_shapes):
            cover_bad.append('致谢页两行分行被改动')
        if 'Thank you' not in flats.get(SLIDE_TOTAL, ''):
            cover_bad.append('致谢页缺"Thank you"')
    pen(-3, "封面页与致谢页的文字未保持原有分行、措辞与摆放位置——封面六处\"智慧园区背景下\"\"再生水回用管网更新项目\"\"风险因素识别与治理策略研究\"\"毕业论文预答辩\"\"答辩人：周昱辰    指导教师：陈若澜\"\"Urban Reclaimed Water Network Renewal · Risk Governance\"，致谢页三处\"敬请各位评审老师\"\"批评指正\"\"Thank you\"",
        bool(cover_bad), "异常：%s" % ('；'.join(cover_bad[:6]) or '无'))

    # -3：目录页"目录""Contents"及 01-05 五个条目出现缺项、改字或编号错乱
    toc_bad = []
    toc_flat = flats.get(2, '')
    for s in TOC_TEXTS:
        if s not in toc_flat:
            toc_bad.append('缺"%s"' % s)
    if not toc_bad:
        pos = [toc_flat.find(k) for k in ('01', '02', '03', '04', '05')]
        if pos != sorted(pos):
            toc_bad.append('01-05 编号顺序错乱')
        for num, title in (('01', '研究概述'), ('02', '利益相关者识别及风险因素分析'),
                           ('03', '风险网络分析及关键因素检验'),
                           ('04', '治理策略分析及模型构建'), ('05', '结论与展望')):
            if toc_flat.find(num) > toc_flat.find(title):
                toc_bad.append('%s 与"%s"错位' % (num, title))
    pen(-3, "目录页\"目录\"\"Contents\"及 01-05 五个条目\"研究概述\"\"利益相关者识别及风险因素分析\"\"风险网络分析及关键因素检验\"\"治理策略分析及模型构建\"\"结论与展望\"出现缺项、改字或编号错乱",
        bool(toc_bad), "目录异常：%s" % ('；'.join(toc_bad[:6]) or '无'))

        # -3：小节标题的编号与中英对照体系出现改号、跳号、缺失或中英文不对应
    # 口径：① 全篇出现的"X.Y 标题"编号集合必须恰好等于 1.1-1.7、2.1-2.4、
    #        3.1-3.11、4.1-4.3、5.1-5.2 这 27 个，不多不少（无改号/跳号）；
    #       ② 每个编号标题必须落在它原本那一页上，且同页有对应的英文副标题
    #        （中英一一对应）。
    num_bad = []
    found_nums = set()
    for i in pages:
        for paras in texts.get(i, []):
            for s in paras:
                m = re.match(r'^([1-5]\.\d{1,2})\s', s)
                if m:
                    found_nums.add(m.group(1))
    want_nums = set(zh.split(' ')[0] for zh in (v[0] for v in HEADINGS.values()))
    if found_nums != want_nums:
        num_bad.append('编号集合异常：多%s 少%s'
                       % (sorted(found_nums - want_nums) or '无',
                          sorted(want_nums - found_nums) or '无'))
    for i, (zh, en) in sorted(HEADINGS.items()):
        page_flat = flats.get(i, '')
        if zh not in page_flat:
            num_bad.append('第%d页缺"%s"' % (i, zh))
        elif en not in page_flat:
            num_bad.append('第%d页"%s"缺英文副标题"%s"' % (i, zh, en))
    pen(-3, "小节标题的编号与中英对照体系（1.1 研究背景 Research background、1.7 技术路线 Technical route、2.3 访谈修订后的因素编码 Revised coding system、3.2 块模型分析 Block model、3.8 熵权-TOPSIS 法确定关键因素 Entropy weight TOPSIS、4.3 治理策略模型构建 Governance model、5.2 研究不足与展望 Limitations & outlook 等）出现改号、跳号、缺失或中英文不对应",
        bool(num_bad), "编号/对照异常：%s" % ('；'.join(num_bad[:6]) or '无'))

    # -3：任一页形状的位置或尺寸被位移、缩放，出现文字与色块错位、文字重叠、
    #     内容溢出页面边界
    # 口径：逐页读每个形状的 a:off/a:ext，凡包围盒越出版心（含约0.33cm容差）
    #   即判为位移/缩放导致的溢出；负坐标同样计入。
    geo_bad = []
    for i in pages:
        for (x, y, w, h) in _shape_boxes(roots[i]):
            over = max(-x, -y, x + w - SLD_CX, y + h - SLD_CY)
            if over > BOUND_TOL:
                geo_bad.append('第%d页超出版心%dEMU' % (i, over))
                break
    pen(-3, "任一页形状的位置或尺寸被位移、缩放，出现文字与色块错位、文字重叠、内容溢出页面边界",
        bool(geo_bad), "越界页：%s" % ('；'.join(geo_bad[:6]) or '无'))

    # -3：为切换添加了未经要求的附加设置——设置自动换片时间使放映按秒自动推进，
    #     或添加风铃、鼓掌、打字机等切换声音
    # 与办公软件对齐："切换"选项卡的"设置自动换片时间"写 <p:transition
    #   advTm="…">（并把 advClick 置 0），"声音"下拉写 <p:sndAc>。
    extra_bad = []
    for i in pages:
        for item in trans_all.get(i, []):
            for r in (item.get('choice'), item.get('fallback')):
                if not r:
                    continue
                if r.get('advTm') is not None:
                    extra_bad.append('第%d页设了自动换片时间' % i)
                if r.get('advClick') == '0':
                    extra_bad.append('第%d页关闭了单击换片' % i)
                if r.get('snd'):
                    extra_bad.append('第%d页加了切换声音' % i)
    pen(-3, "为切换添加了未经要求的附加设置——设置自动换片时间使放映按秒自动推进（答辩人无法按讲述节奏单击控制翻页），或添加风铃、鼓掌、打字机等切换声音",
        bool(extra_bad), "附加设置：%s" % ('；'.join(sorted(set(extra_bad))[:6]) or '无'))

        # -3：1个幻灯片母版与2个幻灯片版式被删除、替换或额外新增
    masters = pkg.match(r'ppt/slideMasters/slideMaster\d+\.xml$')
    layouts = pkg.match(r'ppt/slideLayouts/slideLayout\d+\.xml$')
    ml_bad = (len(masters) != MASTER_COUNT or len(layouts) != LAYOUT_COUNT)
    pen(-3, "1个幻灯片母版与2个幻灯片版式被删除、替换或额外新增",
        ml_bad, "母版 %d 个、版式 %d 个" % (len(masters), len(layouts)))

    # -1：28张正文页左下角的"公开交流版"标识被删除、改字或移位
    # 口径：28 张浅色正文页每页都要有"公开交流版"文本，且其形状仍落在页面
    #   左下角象限（x 中心 < 页宽一半、y 中心 > 页高一半）。
    mark_bad = []
    for i in CONTENT_PAGES:
        if i > n_slides:
            continue
        if PUBLIC_MARK not in flats.get(i, ''):
            mark_bad.append('第%d页缺标识' % i)
            continue
        placed = False
        for sp in roots[i].iter(_q(NS_P, 'sp')):
            tx = sp.find(_q(NS_P, 'txBody'))
            if tx is None:
                continue
            if PUBLIC_MARK not in ''.join(t.text or '' for t in tx.iter(_q(NS_A, 't'))):
                continue
            spPr = sp.find(_q(NS_P, 'spPr'))
            xfrm = spPr.find(_q(NS_A, 'xfrm')) if spPr is not None else None
            off = xfrm.find(_q(NS_A, 'off')) if xfrm is not None else None
            ext = xfrm.find(_q(NS_A, 'ext')) if xfrm is not None else None
            if off is None or ext is None:
                continue
            try:
                cxm = int(off.get('x')) + int(ext.get('cx')) / 2.0
                cym = int(off.get('y')) + int(ext.get('cy')) / 2.0
            except (TypeError, ValueError):
                continue
            if cxm < SLD_CX / 2.0 and cym > SLD_CY / 2.0:
                placed = True
            break
        if not placed:
            mark_bad.append('第%d页标识移位' % i)
    pen(-1, "28张正文页左下角的\"公开交流版\"标识被删除、改字或移位",
        bool(mark_bad), "标识异常：%s" % ('；'.join(mark_bad[:6]) or '无'))

    # -1：35页所用的自动页码域被删除，或被替换为手工输入的固定数字
    # 与办公软件对齐：自动页码在 OOXML 里是 <a:fld type="slidenum">，删域改成
    #   手打数字后该节点即消失，增删页后页码不再自动更新。
    fld_bad = [i for i in pages if 'type="slidenum"' not in raws[i]]
    pen(-1, "35页所用的自动页码域被删除，或被替换为手工输入的固定数字，导致增删页后页码不再自动更新",
        bool(fld_bad), "缺自动页码域的页：%s" % (fld_bad[:8] or '无'))

    # -1：把切换设置一并写入幻灯片母版或版式，造成后续新增页面被动继承同一切换
    inherit_bad = []
    for part in masters + layouts:
        if '<p:transition' in pkg.text(part) or ':transition' in pkg.text(part):
            inherit_bad.append(part)
    pen(-1, "把切换设置一并写入幻灯片母版或版式，造成后续新增页面被动继承同一切换",
        bool(inherit_bad), "含切换的部件：%s" % ('；'.join(inherit_bad) or '无'))

        # -1：附属部件被清除或篡改——35张备注页或备注母版被删除、被写入未经要求的
    #     讲稿文字，或技术路线页所关联的23处 WPS 图示框标签
    #     （KSO_WM_DIAGRAM_VIRTUALLY_FRAME）被清除致该页图示失去可编辑框架
    aux_bad = []
    notes = pkg.match(r'ppt/notesSlides/notesSlide\d+\.xml$')
    if len(notes) != NOTES_COUNT:
        aux_bad.append('备注页 %d 张（应为35张）' % len(notes))
    if not pkg.match(r'ppt/notesMasters/notesMaster\d+\.xml$'):
        aux_bad.append('备注母版被删除')
    for part in notes:
        body = ''.join(re.findall(r'<a:t>(.*?)</a:t>', pkg.text(part), re.S))
        if re.sub(r'[\s　]', '', body):
            aux_bad.append('%s 被写入讲稿文字' % part.split('/')[-1])
    tags = pkg.match(r'ppt/tags/tag\d+\.xml$')
    if len(tags) != TAG_COUNT:
        aux_bad.append('WPS 图示框标签 %d 处（应为23处）' % len(tags))
    else:
        for part in tags:
            if TAG_NAME not in pkg.text(part):
                aux_bad.append('%s 标签名被篡改' % part.split('/')[-1])
                break
    host_rels = 'ppt/slides/_rels/slide%d.xml.rels' % TAG_HOST_PAGE
    if len(re.findall(r'tags/tag\d+\.xml', pkg.text(host_rels))) != TAG_COUNT:
        aux_bad.append('技术路线页与23处图示框标签的关联被切断')
    pen(-1, "附属部件被清除或篡改——35张备注页或备注母版被删除、被写入未经要求的讲稿文字，或技术路线页所关联的23处 WPS 图示框标签（KSO_WM_DIAGRAM_VIRTUALLY_FRAME）被清除致该页图示失去可编辑框架",
        bool(aux_bad), "附属部件异常：%s" % ('；'.join(aux_bad[:6]) or '无'))

    # -1：文件标识信息被改动——文件名被叠加"更新版""论文预答辩""切换动画版"等
    #     描述性前后缀而未沿用"智慧园区再生水管网"，或文档属性中的标题、主题、
    #     作者与最后修改人被改写清空
    id_bad = []
    stem = os.path.splitext(os.path.basename(file_path))[0]
    if FILE_STEM_KEY not in stem:
        id_bad.append('文件名未沿用"%s"' % FILE_STEM_KEY)
    low = stem.lower()
    hit_tokens = [t for t in BAD_NAME_TOKENS if t.lower() in low]
    if hit_tokens:
        id_bad.append('文件名叠加了描述性前后缀：%s' % '、'.join(hit_tokens))
    core = pkg.text('docProps/core.xml')

    def _core(tag):
        m = re.search(r'<(?:[a-zA-Z0-9]+:)?%s[^>]*>(.*?)</' % tag, core, re.S)
        return (m.group(1).strip() if m else '')

    for tag, want, label in (('title', DOC_TITLE, '标题'),
                             ('subject', DOC_SUBJECT, '主题'),
                             ('creator', DOC_CREATOR, '作者'),
                             ('lastModifiedBy', DOC_LASTMOD, '最后修改人')):
        got = _core(tag)
        if got != want:
            id_bad.append('文档属性%s被改为"%s"' % (label, got or '空'))
    pen(-1, "文件标识信息被改动——文件名被叠加\"更新版\"\"论文预答辩\"\"切换动画版\"等描述性前后缀而未沿用\"智慧园区再生水管网\"，或文档属性中的标题\"智慧园区再生水回用管网更新项目风险因素识别与治理策略研究\"、主题\"论文预答辩\"、作者与最后修改人被改写清空",
        bool(id_bad), "标识异常：%s" % ('；'.join(id_bad[:6]) or '无'))

        # -1：随演示文稿一并交付了未被要求的额外文件，如说明文档、脚本、日志、
    #     临时文件或拆分出的分段演示文稿
    # 口径：交付目录内只应有 1 个 .pptx；校验脚本自身及其 __pycache__ 不算交付物。
    self_name = os.path.basename(os.path.abspath(__file__))
    extra = []
    for name in sorted(os.listdir(dir_path)):
        if name == self_name or name in ALLOWED_SIDECAR_DIR:
            continue
        if name.startswith('~$'):
            continue
        full = os.path.join(dir_path, name)
        if os.path.isdir(full):
            extra.append(name + '/')
            continue
        ext = os.path.splitext(name)[1].lower()
        if ext == '.pptx':
            if os.path.abspath(full) != os.path.abspath(file_path):
                extra.append(name)
            continue
        if ext in ALLOWED_SIDECAR_EXT:
            continue
        extra.append(name)
    pen(-1, "随演示文稿一并交付了未被要求的额外文件，如说明文档、脚本、日志、临时文件或拆分出的分段演示文稿",
        bool(extra), "额外文件：%s" % ('、'.join(extra[:6]) or '无'))

    # ========== 将 results 反解为结构化 dim2_items ==========
    # 逐项按追加顺序对齐 _MAX_DELTAS_ORDER + _PENALTY_DELTAS_ORDER：
    #   · 加分项："+N：说明" 命中 N=满分，未命中 N=0；
    #   · 减分项："-N：说明" 命中即触发扣分，"-0：说明" 表示未踩该项。
    #   hit=True 对加分项表示"做到了"，对减分项表示"踩到了该违规事实"。
    order = list(_MAX_DELTAS_ORDER) + list(_PENALTY_DELTAS_ORDER)
    if len(results) != len(order):
        raise RuntimeError("dim2 项数不一致：results=%d vs order=%d"
                           % (len(results), len(order)))
    dim2_items = []
    total_score = 0
    for raw, max_delta in zip(results, order):
        head, _, rule = raw.partition("：")
        try:
            delta = int(head.replace("+", "").strip())
        except ValueError:
            delta = 0
        hit = (delta != 0)
        dim2_items.append({
            "rule": rule,
            "max_delta": max_delta,
            "delta": max_delta if hit else 0,
            "hit": hit,
        })
        total_score += (max_delta if hit else 0)

    result["dim2_items"] = dim2_items
    result["total_score"] = max(0, total_score)
    result["max_score"] = MAX_SCORE_TOTAL
    return result


if __name__ == "__main__":
    # 本地调试入口：evaluate 只接收"脚本所在目录的路径"，由脚本自己在该目录里
    # 定位并打开被评估的 .pptx。命令行第 1 个参数可覆盖目录路径；未指定时使用
    # 脚本自身所在目录。
    # 注：evaluate() 的接口约定就是"入参为目录路径"——路径始终由外部显式传入；
    #     这里的默认值仅为本地调试便利，不影响 evaluate 的接口约定。
    if len(sys.argv) >= 2:
        _dir = sys.argv[1]
    else:
        _dir = os.path.dirname(os.path.abspath(__file__))
    _out = evaluate(_dir)
    print(json.dumps(_out, ensure_ascii=False, indent=2))
