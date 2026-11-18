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
import xml.etree.ElementTree as ET

# 说明：按接口统一约定，本模块不修改全局 sys.stdout，也不 print 主结果。
# 主结果统一由 evaluate(dir_path) -> dict 返回；入参为"脚本所在目录的路径"，
# 由脚本自己在该目录内定位并打开被评估的 .pptx。本文件仅在 __main__ 下作为
# 本地调试入口，才把 dict 序列化为 JSON 打印。
#
# 本脚本只读解析 OOXML 包（zipfile + ElementTree），不依赖 python-pptx，
# 避免任何"打开即重写"造成的误判。
#
# 评分前提：加分项与减分项都建立在"交付件确实做了改动"之上——加分项对应指令要求
# 干的活，减分项对应指令之外惹的麻烦。交付件原样未动时既无功也无过，全部条目一律
# 记 0 分，总分 0（见 S4_UNADJUSTED_PIC / S4_UNADJUSTED_SIG）。

SCRIPT_ID = "060"
# 维度二加分项 max_delta 之和（满分）；维度一未通过 / 脚本异常时兜底为 0
MAX_SCORE_TOTAL = 22
# 维度二减分项 max_delta 之和（全部踩中时的扣分总额）
MIN_PENALTY_TOTAL = -58

# 维度二所有评分项按"结果追加顺序"的 max_delta 列表；末端把 results 里的
# "+N：..." / "-N：..." 字符串反解为结构化 dim2_items 时用。顺序须与
# _evaluate_impl 中 results.append(...) 的顺序严格一致；共 24 项，
# 正项累加 = MAX_SCORE_TOTAL，负项累加 = MIN_PENALTY_TOTAL。
_MAX_DELTAS_ORDER = [
    5,   # 01 第4页照片显示宽高比 1.75-1.80
    5,   # 02 第4页照片横向满版
    5,   # 03 第4页照片纵向满版且高度收敛到页面高度量级
    3,   # 04 第4页完整呈现三段工艺链路，未被页面边界裁掉
    3,   # 05 第4页照片位于最底层，文字全部叠压其上且完整可读
    1,   # 06 第4页照片仍为独立图片对象并保留等比缩放锁定
    -5,  # 07 照片被替换成重新生成的位图
    -5,  # 08 交付文件体积超过 2.6MB
    -5,  # 09 第4页三条要点文字内容/格式/文本框被改动
    -5,  # 10 第4页4个青色标签文字/数量/两行两列排布被改动
    -5,  # 11 第1、2、3、5页被改动
    -3,  # 12 image4.jpeg 成为冗余资源且残留失效图片关系
    -3,  # 13 第4页右侧深色圆角信息卡被改动
    -3,  # 14 第4页标题被改动
    -3,  # 15 第4页副标题被改动
    -3,  # 16 第4页4个青色标签的尺寸/位置/颜色/文字格式被改动
    -3,  # 17 第4页底部说明文字被改动
    -3,  # 18 第4页新增了指令未要求的元素 / 备注批注宏
    -3,  # 19 第4页删除了 15 个既有对象中的任一个
    -3,  # 20 第4页对象叠放次序被打乱
    -3,  # 21 页数/版式/页面尺寸/主题母版版式被改动
    -1,  # 22 第4页青色装饰条被改动
    -1,  # 23 第4页文本框 spAutoFit 与内边距被改动
    -1,  # 24 docProps/thumbnail.jpeg 预览图未随第4页版式更新
]
assert sum(d for d in _MAX_DELTAS_ORDER if d > 0) == MAX_SCORE_TOTAL
assert sum(d for d in _MAX_DELTAS_ORDER if d < 0) == MIN_PENALTY_TOTAL

# ---------- OOXML 命名空间 ----------
NS_A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
NS_P = 'http://schemas.openxmlformats.org/presentationml/2006/main'
NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
NS_PR = 'http://schemas.openxmlformats.org/package/2006/relationships'


def _a(tag):
    return '{%s}%s' % (NS_A, tag)


def _p(tag):
    return '{%s}%s' % (NS_P, tag)


def _r(tag):
    return '{%s}%s' % (NS_R, tag)


# ---------- 细则里写死的绝对值（全部为 EMU / 百分点 / 十六进制色值） ----------
PAGE_W, PAGE_H = 12192000, 6858000     # 页面尺寸 12192000×6858000 EMU（16:9）
EMU_CM = 360000                        # 1cm = 360000 EMU
TOL_BLANK = int(0.2 * EMU_CM)          # 0.2cm 空白条阈值 = 72000 EMU
FONT = '微软雅黑'
SLIDE_COUNT = 5
SLIDE4_LAYOUT = '../slideLayouts/slideLayout2.xml'
SIZE_LIMIT = int(2.6 * 1024 * 1024)    # 交付文件体积上限 2.6MB

# 第4页背景照片：必须仍是包内 ppt/media/image4.jpeg 这张 1467×941 的 JPEG，
# 且图形编号 id="2"、名称 "Picture 1"、替换文字 "crop_filter.jpg"
S4_PIC_TARGET = '../media/image4.jpeg'
S4_PIC_PX = (1467, 941)
S4_PIC_ID = '2'
S4_PIC_NAME = 'Picture 1'
S4_PIC_DESCR = 'crop_filter.jpg'
MEDIA_COUNT = 5                        # ppt/media 下图片数量

# 第4页各既有对象的绝对几何/颜色/文字格式（off_x, off_y, cx, cy）
S4_CARD = (7022592, 502920, 4754880, 5852160)
S4_CARD_FILL, S4_CARD_LINE = '0D3B42', 'FFFFFF'
S4_TITLE_BOX = (7278624, 758952, 4242816, 640080)
S4_TITLE_TEXT = '多级膜分离与清水回用链路'
S4_BAR = (7333488, 1472184, 1828800, 27432)
S4_ACCENT = '2BB2B7'
S4_SUB_BOX = (7278624, 1581912, 4242816, 777240)
S4_SUB_TEXT = '预处理、膜组件与透明水箱构成连续式中试平台'
S4_SUB_COLOR = 'E9F2F6'
S4_BUL_BOX = (7333488, 2468880, 4133088, 2057400)
S4_BULLETS = (
    '• 左侧高浊度水进入预处理罐，去除大颗粒与悬浮物',
    '• 中段膜堆完成深度分离，压差信号用于判断清洗周期',
    '• 右侧清水箱作为回用缓冲，便于后续稳定供水',
)
S4_TAG_EXT = (1984248, 393192)
S4_TAG_X = (7315200, 9500616)           # 左列 / 右列
S4_TAG_Y = (4892040, 5458968)           # 上行 / 下行
S4_TAGT_EXT = (1892807, 356616)
S4_TAGT_X = (7360920, 9546336)
S4_TAGT_Y = (4910327, 5477255)
S4_TAGS = ('预处理', '膜分离', '压差判读', '回用缓冲')
S4_FOOT_BOX = (411480, 6446520, 11384280, 256032)
S4_FOOT_TEXT = '工艺说明：预处理、膜组件与透明水箱构成连续式中试平台'
WHITE = 'FFFFFF'
INSETS = ('54864', '18288', '54864', '18288')   # lIns / tIns / rIns / bIns

# 第1、2、3、5页：不需要修改的页，逐页写死背景照片摆放与全部文字/版面绝对值。
# pic=(off_x, off_y, cx, cy, descr)；card_x/title_x/bar_x 为该页信息卡一侧的横向基准；
# tag_x/tagt_x 为标签左右两列横向偏移（纵向与第4页同为 4892040 / 5458968、
# 标签文字同为 4910327 / 5477255）。
OTHER_SLIDES = {
    1: {
        'pic': (0, 635, 12192635, 6966585, 'crop_wetland.jpg'),
        'card_fill': '123E2C', 'accent': '66AD65',
        'card_x': 384048, 'title_x': 640080, 'bar_x': 694944,
        'tag_x': (676656, 2862072), 'tagt_x': (722376, 2907791),
        'title': '湿地物联网水质净化系统',
        'sub': '太阳能传感器、分级植物床与在线水样检测协同运行',
        'bullets': ('• 浑浊进水经沉降、植物根区与微生物膜逐级削减颗粒物',
                    '• 边缘传感器实时回传水位、浊度与溶氧变化',
                    '• 末端清水区用于复核采样与生态补水回用'),
        'tags': ('进水', '根区反应', '在线监测', '出水回用'),
        'foot': '工艺说明：太阳能传感器、分级植物床与在线水样检测协同运行',
    },
    2: {
        'pic': (0, 0, 12221845, 6858000, 'crop_biogas.jpg'),
        'card_fill': '3F2816', 'accent': 'FFA64D',
        'card_x': 7022592, 'title_x': 7278624, 'bar_x': 7333488,
        'tag_x': (7315200, 9500616), 'tagt_x': (7360920, 9546336),
        'title': '餐厨残渣厌氧发酵与沼气追踪',
        'sub': '密闭反应器、气囊储存与数据面板形成闭环监控',
        'bullets': ('• 有机残渣在恒温罐内完成水解、酸化与产甲烷过程',
                    '• 压力、温度与产气量通过节点设备同步采集',
                    '• 沼气进入柔性储袋，异常波动触发运行参数调整'),
        'tags': ('投料', '发酵', '储气', '调参'),
        'foot': '工艺说明：密闭反应器、气囊储存与数据面板形成闭环监控',
    },
    3: {
        'pic': (0, -674370, 12085320, 8206105, 'crop_electro.jpg'),
        'card_fill': '0F2A44', 'accent': '5BAFE4',
        'card_x': 384048, 'title_x': 640080, 'bar_x': 694944,
        'tag_x': (676656, 2862072), 'tagt_x': (722376, 2907791),
        'title': '电化学反应器强化去污',
        'sub': '透明流道中电极板调控颗粒絮凝与溶解污染物剥离',
        'bullets': ('• 电极阵列形成稳定微电场，促进胶体颗粒聚集沉降',
                    '• 循环流道保持剪切强度，避免局部堵塞和短流',
                    '• 控制端根据电导率与流量变化调节脉冲强度'),
        'tags': ('电极阵列', '循环流道', '脉冲控制', '稳定出水'),
        'foot': '工艺说明：透明流道中电极板调控颗粒絮凝与溶解污染物剥离',
    },
    5: {
        'pic': (635, -635, 12713970, 6858635, 'crop_adsorb.jpg'),
        'card_fill': '293021', 'accent': '778F4D',
        'card_x': 384048, 'title_x': 640080, 'bar_x': 694944,
        'tag_x': (676656, 2862072), 'tagt_x': (722376, 2907791),
        'title': '吸附柱捕集与显微验证',
        'sub': '炭基填料、样品瓶与显微观察用于确认微颗粒去除效果',
        'bullets': ('• 河道来水通过双柱填料，微塑料与金属颗粒被截留吸附',
                    '• 分级样品瓶展示浊度逐步下降的过程变化',
                    '• 显微图像复核填料表面负载状态，支撑工艺优化'),
        'tags': ('河道取样', '炭基吸附', '样品对比', '显微复核'),
        'foot': '工艺说明：炭基填料、样品瓶与显微观察用于确认微颗粒去除效果',
    },
}

# 主题 / 母版 / 11 个版式的规范化内容指纹（_canon_sig 的 MD5）。这些部件属于
# "不需要修改"的内容，指纹变化即说明被改动。规范化会剔除 XML 声明差异、
# 属性顺序、空白文本节点等纯序列化噪声，只对语义内容取指纹。
STATIC_PART_SIG = {
    'ppt/theme/theme1.xml': '2280c79c393f730102f1d2089cb0ec60',
    'ppt/slideMasters/slideMaster1.xml': '29c92092c21769b7d8b636fdef66128f',
    'ppt/slideLayouts/slideLayout1.xml': '816c78e76c6417fbac3577e01bf1e6b9',
    'ppt/slideLayouts/slideLayout2.xml': 'bd44692cbe830898544d8679fb5455e6',
    'ppt/slideLayouts/slideLayout3.xml': '78cd360338bffb7d937a4b9d25d7a919',
    'ppt/slideLayouts/slideLayout4.xml': 'cc0e37cb897a8e7bf957be3b0f1cdae9',
    'ppt/slideLayouts/slideLayout5.xml': 'ef0008aa95c6eea1ed091dd4bdcbc5d2',
    'ppt/slideLayouts/slideLayout6.xml': '369db3f53ba8d74ecd218d8aa3a22121',
    'ppt/slideLayouts/slideLayout7.xml': 'f67a4745b4c36b882a94ae5079ad9a5d',
    'ppt/slideLayouts/slideLayout8.xml': '4971c275cb0583044f7c7a683415df3d',
    'ppt/slideLayouts/slideLayout9.xml': '26c82441e0cdbcca95c9be17747b1692',
    'ppt/slideLayouts/slideLayout10.xml': 'ecb1a466a02d1b6876d48f1b2d4386c8',
    'ppt/slideLayouts/slideLayout11.xml': 'd950802d47ec9fb7647d4a128bfb788a',
}

# 与"第4页调整前"完全一致的那张预览图字节签名：交付件的 docProps/thumbnail.jpeg
# 若仍等于该签名（或整个预览图缺失），说明预览图没有随第4页版式一同更新。
STALE_THUMB_MD5 = 'cea0e726227c3926e7edee78433306b8'

# 第4页"未作任何调整"的状态签名。加分项与减分项都以"交付件确实动过第4页"为前提：
# 原样交付=既没干活也没惹麻烦，全部条目一律记 0 分，不加也不减。
#   S4_UNADJUSTED_PIC：背景照片仍停在近方形、上下各溢出百万级 EMU 的未调整摆放
#     （横向232410、纵向-1483360、尺寸11742420×11650980 EMU），照片摆放等于该值
#     即视为未着手指令要求的排版工作，6 条加分项全部不计分；
#   S4_UNADJUSTED_SIG：整页 slide4.xml 的语义指纹仍等于未调整状态，说明第4页版式
#     没有任何变化，"预览图未随第4页版式一同更新"这条也就无从谈起，不予扣分。
S4_UNADJUSTED_PIC = (232410, -1483360, 11742420, 11650980)
S4_UNADJUSTED_SIG = 'fba75d674a94eb0e58b0a52413f4d3f6'


# ==================== 只读解析工具 ====================

def _canon_sig(data):
    """对 XML 部件取"语义指纹"：递归序列化 标签 + 排序后的属性 + 去空白文本，
    从而忽略 XML 声明引号、属性书写顺序、空元素写法、缩进空白等纯序列化差异。"""
    def walk(el, out):
        out.append('<' + el.tag)
        for k in sorted(el.attrib):
            out.append(' %s=%r' % (k, el.attrib[k]))
        out.append('>')
        if el.text and el.text.strip():
            out.append(el.text.strip())
        for ch in el:
            walk(ch, out)
        out.append('</' + el.tag + '>')
        if el.tail and el.tail.strip():
            out.append(el.tail.strip())
    buf = []
    walk(ET.fromstring(data), buf)
    return hashlib.md5(''.join(buf).encode('utf-8')).hexdigest()


def _img_px(data):
    """返回位图的像素尺寸 (w, h) 与格式标识；无法识别时返回 (None, None, fmt)。"""
    if data[:3] == b'\xff\xd8\xff':
        i, n = 2, len(data)
        while i < n - 9:
            if data[i] != 0xFF:
                i += 1
                continue
            m = data[i + 1]
            if m in (0xC0, 0xC1, 0xC2, 0xC3, 0xC5, 0xC6, 0xC7,
                     0xC9, 0xCA, 0xCB, 0xCD, 0xCE, 0xCF):
                h, w = struct.unpack('>HH', data[i + 5:i + 9])
                return (w, h, 'jpeg')
            if m == 0xD8 or m == 0xD9 or 0xD0 <= m <= 0xD7:
                i += 2
                continue
            i += 2 + struct.unpack('>H', data[i + 2:i + 4])[0]
        return (None, None, 'jpeg')
    if data[:8] == b'\x89PNG\r\n\x1a\n':
        w, h = struct.unpack('>II', data[16:24])
        return (w, h, 'png')
    return (None, None, 'other')


class _Deck(object):
    """把 .pptx 当成只读 OOXML 包访问：部件字节、幻灯片 XML、关系表。"""

    def __init__(self, file_path):
        self.path = file_path
        self.size = os.path.getsize(file_path)
        self.zf = zipfile.ZipFile(file_path)
        self.names = set(self.zf.namelist())

    def has(self, name):
        return name in self.names

    def raw(self, name):
        return self.zf.read(name)

    def xml(self, name):
        return ET.fromstring(self.zf.read(name))

    def rels(self, part):
        """返回 {rId: (Type简名, Target)}。"""
        d, base = {}, os.path.basename(part)
        rp = part[:-len(base)] + '_rels/' + base + '.rels'
        if not self.has(rp):
            return d
        for rel in self.xml(rp):
            d[rel.get('Id')] = (rel.get('Type', '').rsplit('/', 1)[-1], rel.get('Target', ''))
        return d

    def slide_count(self):
        pres = self.xml('ppt/presentation.xml')
        lst = pres.find(_p('sldIdLst'))
        return 0 if lst is None else len(list(lst))

    def page_size(self):
        sz = self.xml('ppt/presentation.xml').find(_p('sldSz'))
        if sz is None:
            return (None, None)
        return (int(sz.get('cx', 0)), int(sz.get('cy', 0)))

    def sp_tree(self, idx):
        return self.xml('ppt/slides/slide%d.xml' % idx).find(_p('cSld')).find(_p('spTree'))

    def media(self):
        return sorted(n for n in self.names
                      if n.startswith('ppt/media/') and not n.endswith('/'))


def _shapes(sp_tree):
    """返回 spTree 中"可单独点选"的对象序列：[(kind, element), ...]，
    顺序即 PowerPoint 的叠放次序（文档序，越靠前越靠下）。"""
    out = []
    for ch in sp_tree:
        tag = ch.tag.split('}')[-1]
        if tag in ('sp', 'pic', 'graphicFrame', 'grpSp', 'cxnSp'):
            out.append((tag, ch))
    return out


def _cnvpr(kind, sh):
    nv = sh.find(_p('nvPicPr') if kind == 'pic' else _p('nvSpPr'))
    if nv is None:
        for tag in ('nvGraphicFramePr', 'nvGrpSpPr', 'nvCxnSpPr'):
            nv = sh.find(_p(tag))
            if nv is not None:
                break
    return None if nv is None else nv.find(_p('cNvPr'))


def _sppr(sh):
    return sh.find(_p('spPr'))


def _xfrm(sh):
    """返回 (x, y, cx, cy)；缺 xfrm 时返回 None。"""
    spr = _sppr(sh)
    xf = None if spr is None else spr.find(_a('xfrm'))
    if xf is None:
        return None
    off, ext = xf.find(_a('off')), xf.find(_a('ext'))
    if off is None or ext is None:
        return None
    return (int(off.get('x', 0)), int(off.get('y', 0)),
            int(ext.get('cx', 0)), int(ext.get('cy', 0)))


def _prst(sh):
    spr = _sppr(sh)
    g = None if spr is None else spr.find(_a('prstGeom'))
    return None if g is None else g.get('prst')


def _solid_srgb(el):
    """读 <a:solidFill><a:srgbClr val=".."/>；非 srgb 或无填充返回 None。"""
    if el is None:
        return None
    sf = el.find(_a('solidFill'))
    if sf is None:
        return None
    c = sf.find(_a('srgbClr'))
    return None if c is None else (c.get('val') or '').upper()


def _fill(sh):
    return _solid_srgb(_sppr(sh))


def _line(sh):
    """返回 ('none', None) / ('solid', 'FFFFFF') / (None, None)。"""
    spr = _sppr(sh)
    ln = None if spr is None else spr.find(_a('ln'))
    if ln is None:
        return (None, None)
    if ln.find(_a('noFill')) is not None:
        return ('none', None)
    c = _solid_srgb(ln)
    return ('solid', c) if c else (None, None)


def _is_txbox(kind, sh):
    if kind != 'sp':
        return False
    nv = sh.find(_p('nvSpPr'))
    cn = None if nv is None else nv.find(_p('cNvSpPr'))
    return cn is not None and cn.get('txBox') == '1'


def _body_pr(sh):
    """返回 (lIns, tIns, rIns, bIns, has_spAutoFit)；无 txBody 返回 None。"""
    tb = sh.find(_p('txBody'))
    bp = None if tb is None else tb.find(_a('bodyPr'))
    if bp is None:
        return None
    return (bp.get('lIns'), bp.get('tIns'), bp.get('rIns'), bp.get('bIns'),
            bp.find(_a('spAutoFit')) is not None)


def _paras(sh):
    """把形状文字解析成段落列表：
    [{'algn':.., 'spc_aft':.., 'text':.., 'segs':[('r'|'br', text, latin, sz, bold, color)]}]
    （细则里三条要点在同一段内用 <a:br> 换行，因此 br 也要单独记录。）"""
    tb = sh.find(_p('txBody'))
    if tb is None:
        return []
    out = []
    for par in tb.findall(_a('p')):
        ppr = par.find(_a('pPr'))
        algn = None if ppr is None else ppr.get('algn')
        spc = None
        if ppr is not None:
            sa = ppr.find(_a('spcAft'))
            if sa is not None:
                pts = sa.find(_a('spcPts'))
                spc = None if pts is None else pts.get('val')
        segs, txt = [], []
        for node in par:
            ln = node.tag.split('}')[-1]
            if ln not in ('r', 'br'):
                continue
            rpr = node.find(_a('rPr'))
            sz = b = col = lat = None
            if rpr is not None:
                sz, b = rpr.get('sz'), rpr.get('b')
                col = _solid_srgb(rpr)
                la = rpr.find(_a('latin'))
                lat = None if la is None else la.get('typeface')
            if ln == 'br':
                segs.append(('br', '', lat, sz, b, col))
                txt.append('\n')
            else:
                t = node.find(_a('t'))
                t = '' if t is None or t.text is None else t.text
                segs.append(('r', t, lat, sz, b, col))
                txt.append(t)
        out.append({'algn': algn, 'spc_aft': spc, 'text': ''.join(txt), 'segs': segs})
    return out


def _text_of(sh):
    return '\n'.join(pa['text'] for pa in _paras(sh)).strip()


def _runs_ok(paras, sz, bold, color, font=FONT):
    """所有文字 run（含 <a:br> 上继承的 rPr）均满足 字号/加粗/颜色/字体。
    字号单位为百分点（15磅=1500）；加粗按 b="1" 判定，非加粗接受缺省或 b="0"。"""
    seen = False
    for pa in paras:
        for kind, _t, lat, s, b, col in pa['segs']:
            if kind == 'r':
                seen = True
            if s != sz:
                return False
            if bold and b != '1':
                return False
            if (not bold) and b not in (None, '0'):
                return False
            if (col or '').upper() != color:
                return False
            if lat != font:
                return False
    return seen


def _one_para_text(sh, text, algn, sz, bold, color):
    """单段落文本框：内容、对齐、字号、加粗、颜色全部对齐细则写死的值。"""
    paras = [pa for pa in _paras(sh) if pa['segs']]
    if len(paras) != 1 or paras[0]['text'] != text:
        return False
    if paras[0]['algn'] != algn:
        return False
    return _runs_ok(paras, sz, bold, color)


def _pic_blip_rid(sh):
    bf = sh.find(_p('blipFill'))
    blip = None if bf is None else bf.find(_a('blip'))
    return None if blip is None else blip.get(_r('embed'))


def _pic_srcrect(sh):
    """返回 blipFill 的裁剪量 (l, t, r, b)（千分之一百分比），未裁剪为 (0,0,0,0)。"""
    bf = sh.find(_p('blipFill'))
    sr = None if bf is None else bf.find(_a('srcRect'))
    if sr is None:
        return (0, 0, 0, 0)
    return tuple(int(sr.get(k, 0)) for k in ('l', 't', 'r', 'b'))


def _pic_lock_aspect(sh):
    nv = sh.find(_p('nvPicPr'))
    cp = None if nv is None else nv.find(_p('cNvPicPr'))
    lk = None if cp is None else cp.find(_a('picLocks'))
    return lk is not None and lk.get('noChangeAspect') == '1'


def _inside(rect, box, tol=0):
    """rect 是否完整落在 box(=x,y,cx,cy) 内（允许 tol 容差）。"""
    x, y, cx, cy = rect
    bx, by, bcx, bcy = box
    return (x >= bx - tol and y >= by - tol
            and x + cx <= bx + bcx + tol and y + cy <= by + bcy + tol)


def _classify_s4(shapes):
    """按"对象自身特征"（而不是位置索引）把第4页对象归类到 15 个角色，
    使叠放次序、删除对象、几何改动三类判定互不串扰。
    返回 (roles, order)：roles 为 {角色名: (kind, element)}；
    order 为按叠放次序排列的角色类别序列（未能归类者记 'unknown'）。"""
    roles, order = {}, []
    tag_i = 0
    tagt_seen = []
    for kind, sh in shapes:
        role = None
        if kind == 'pic':
            role = 'photo'
        elif _is_txbox(kind, sh):
            t = _text_of(sh)
            xf = _xfrm(sh)
            if t == S4_TITLE_TEXT or xf == S4_TITLE_BOX:
                role = 'title'
            elif t == S4_SUB_TEXT or xf == S4_SUB_BOX:
                role = 'subtitle'
            elif t.replace('\n', '').startswith(S4_BULLETS[0][:12]) or xf == S4_BUL_BOX:
                role = 'bullets'
            elif t.startswith('工艺说明') or xf == S4_FOOT_BOX:
                role = 'footer'
            elif (t in S4_TAGS and t not in tagt_seen) or (
                    xf and xf[0] in S4_TAGT_X and xf[1] in S4_TAGT_Y):
                tagt_seen.append(t)
                role = 'tagtext%d' % len(tagt_seen)
        elif kind == 'sp':
            xf = _xfrm(sh)
            prst, fill = _prst(sh), _fill(sh)
            if prst == 'roundRect' and xf and xf[2] > 3000000:
                role = 'card'
            elif prst == 'roundRect':
                tag_i += 1
                role = 'tag%d' % tag_i
            elif prst == 'rect' and fill and xf and xf[3] < 200000:
                role = 'bar'
        if role is None or role in roles:
            order.append('unknown')
            continue
        roles[role] = (kind, sh)
        order.append(re.sub(r'\d+$', '', role))
    return roles, order


# 第4页叠放次序：照片在最底层，其后依次为深色信息卡、标题、青色装饰条、副标题、
# 要点文字、4 组"标签+标签文字"、底部说明文字
S4_ORDER = (['photo', 'card', 'title', 'bar', 'subtitle', 'bullets']
            + ['tag', 'tagtext'] * 4 + ['footer'])
S4_ROLES = (['photo', 'card', 'title', 'bar', 'subtitle', 'bullets', 'footer']
            + ['tag%d' % i for i in (1, 2, 3, 4)]
            + ['tagtext%d' % i for i in (1, 2, 3, 4)])


def _bullets_ok(sh, texts):
    """三条要点：同一段落内 3 个 run 以 2 个 <a:br> 分隔、内容逐条相同、
    段后间距 6 磅（spcPts val="600"）、左对齐、微软雅黑 15 磅非加粗白色。"""
    paras = [pa for pa in _paras(sh) if pa['segs']]
    if len(paras) != 1:
        return False
    pa = paras[0]
    if pa['algn'] != 'l' or pa['spc_aft'] != '600':
        return False
    kinds = [s[0] for s in pa['segs']]
    if kinds != ['r', 'br', 'r', 'br', 'r']:
        return False
    if tuple(s[1] for s in pa['segs'] if s[0] == 'r') != tuple(texts):
        return False
    return _runs_ok(paras, '1500', False, WHITE)


def _deco_slide_ok(shapes, spec):
    """第1、2、3、5页：15 个对象的类型、叠放次序、位置尺寸、填充描边、
    文字内容、字体、字号、加粗、颜色、对齐方式全部与写死的绝对值一致。"""
    if len(shapes) != 15:
        return False
    kinds = [k for k, _ in shapes]
    if kinds != ['pic'] + ['sp'] * 14:
        return False
    sh = [s for _, s in shapes]

    px, py, pcx, pcy, pdescr = spec['pic']
    cpr = _cnvpr('pic', sh[0])
    if _xfrm(sh[0]) != (px, py, pcx, pcy):
        return False
    if cpr is None or cpr.get('descr') != pdescr:
        return False

    if _xfrm(sh[1]) != (spec['card_x'], 502920, 4754880, 5852160):
        return False
    if _prst(sh[1]) != 'roundRect' or _fill(sh[1]) != spec['card_fill']:
        return False
    if _line(sh[1]) != ('solid', WHITE):
        return False

    if _xfrm(sh[2]) != (spec['title_x'], 758952, 4242816, 640080):
        return False
    if not _one_para_text(sh[2], spec['title'], 'l', '2800', True, WHITE):
        return False

    if _xfrm(sh[3]) != (spec['bar_x'], 1472184, 1828800, 27432):
        return False
    if _prst(sh[3]) != 'rect' or _fill(sh[3]) != spec['accent'] or _line(sh[3])[0] != 'none':
        return False

    if _xfrm(sh[4]) != (spec['title_x'], 1581912, 4242816, 777240):
        return False
    if not _one_para_text(sh[4], spec['sub'], 'l', '1400', False, S4_SUB_COLOR):
        return False

    if _xfrm(sh[5]) != (spec['bar_x'], 2468880, 4133088, 2057400):
        return False
    if not _bullets_ok(sh[5], spec['bullets']):
        return False

    for i in range(4):
        col, row = i % 2, i // 2
        rect, text = sh[6 + i * 2], sh[7 + i * 2]
        if _xfrm(rect) != (spec['tag_x'][col], S4_TAG_Y[row]) + S4_TAG_EXT:
            return False
        if _prst(rect) != 'roundRect' or _fill(rect) != spec['accent']:
            return False
        if _line(rect)[0] != 'none':
            return False
        if _xfrm(text) != (spec['tagt_x'][col], S4_TAGT_Y[row]) + S4_TAGT_EXT:
            return False
        if not _one_para_text(text, spec['tags'][i], 'ctr', '1200', True, WHITE):
            return False

    if _xfrm(sh[14]) != S4_FOOT_BOX:
        return False
    return _one_para_text(sh[14], spec['foot'], 'r', '800', False, WHITE)


class _TargetMissing(Exception):
    """目录不存在或目录中没有 .pptx：连待评测对象都没有，不进入维度判定。"""


def evaluate(dir_path):
    """按 §2.2 约定，返回结构化 dict；不 print、不改 sys.stdout、不 sys.exit。
    入参 dir_path 为"脚本所在目录的路径"，脚本自己在该目录里定位并打开被评估的
    .pptx（忽略以 ~$ 开头的 PowerPoint 临时锁文件）。目录不存在 / 未发现文档时
    返回 status='error' 且维度一未参与判定（dim1_pass 留空）；文件存在但打不开
    或脚本崩溃时返回 status='error' 且 dim1_pass=False 并给出错误摘要。"""
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
        cands = sorted(
            f for f in os.listdir(dir_path)
            if f.lower().endswith('.pptx') and not f.startswith('~$')
        )
        if not cands:
            raise _TargetMissing(f"目录中未找到 .pptx 文件：{dir_path}")
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
        result["error"] = f"{type(exc).__name__}: {exc}"
        result["dim1_pass"] = False
        result["dim1_reason"] = f"文件无法解析或无法打开：{type(exc).__name__}: {exc}"
        result["dim2_items"] = []
        result["total_score"] = 0
        return result


def _dim1_check(file_path):
    """维度1：可用与可修改性。返回 (deck, reason)；reason 非空即门禁未通过，
    完成度总分直接记 0。
    细则只保留一条：交付文件为 1 个可正常打开的 PowerPoint 演示文稿（.pptx）——
    zip 包完整，且 ppt/presentation.xml 与第 4 页 slide 部件存在并可解析；文件
    损坏无法打开即判总分 0。
    页数、第 4 页版面是否错乱、15 个对象是否可单独点选编辑等不再作为门禁条件，
    统一交由维度二的减分项体现（换图 -5、删除既有对象 -3、叠放次序 -3、
    页数/版式/页面尺寸 -3 等）。"""
    try:
        deck = _Deck(file_path)
    except Exception as exc:
        return (None, f"文件损坏无法打开：{type(exc).__name__}: {exc}")
    for need in ('ppt/presentation.xml', 'ppt/slides/slide4.xml'):
        if not deck.has(need):
            return (None, f"文件损坏无法打开：包内缺少必需部件 {need}")
    try:
        deck.slide_count()
        deck.sp_tree(4)
    except Exception as exc:
        return (None, f"文件损坏无法打开：演示文稿主体无法解析：{type(exc).__name__}: {exc}")
    return (deck, "")


def _evaluate_impl(file_path, result):
    """维度一门禁 + 维度二逐条评分；末尾按 _MAX_DELTAS_ORDER 把 results 里的
    "+N：..." / "-N：..." 反解成结构化 dim2_items 填入 result。"""
    deck, reason = _dim1_check(file_path)
    if reason:
        # 维度一未通过：完成度总分直接记 0
        result["status"] = "error"
        result["error"] = reason
        result["dim1_pass"] = False
        result["dim1_reason"] = reason
        result["dim2_items"] = []
        result["total_score"] = 0
        return result

    results = []
    shapes = _shapes(deck.sp_tree(4))
    roles, order = _classify_s4(shapes)
    rels4 = deck.rels('ppt/slides/slide4.xml')
    photo = roles.get('photo')
    pic = None if photo is None else photo[1]
    box = None if pic is None else _xfrm(pic)
    text_shapes = [s for k, s in shapes if _is_txbox(k, s) and _text_of(s)]

    # 评分前提：加减分项都建立在"交付件确实动过第4页"之上。
    #   s4_touched：整页 slide4.xml 的语义指纹已不同于未调整状态；
    #   pic_adjusted：背景照片的摆放（位置+尺寸）已不同于未调整摆放，即确实着手了
    #     指令要求的"重新排版图片"这件事——6 条加分项均以此为前提。
    # 原样未动的交付件：加分项无功可记、减分项无过可罚，全部条目记 0，总分 0。
    s4_touched = _canon_sig(deck.raw('ppt/slides/slide4.xml')) != S4_UNADJUSTED_SIG
    pic_adjusted = box != S4_UNADJUSTED_PIC

    # ========== 维度2：完成度评分 ==========

    # +5：第4页背景照片的显示宽高比与页面16:9一致，宽高比值落在1.75-1.80区间
    # 严格对齐细则，仅踩细则明确写出的点：
    #   ① 照片显示宽高比 = 显示宽度 cx / 显示高度 cy，落在 [1.75, 1.80]
    #   ② 不呈宽高比约 1.0 的近方形摆放（由 ① 的区间同时排除）
    # 与办公软件对齐：读 p:pic/p:spPr/a:xfrm/a:ext 的 cx、cy，即 PowerPoint
    #   "大小"面板里看到的图片显示尺寸（而不是位图自身的像素比）。
    ratio = (box[2] / box[3]) if (box and box[3]) else None
    if pic_adjusted and ratio is not None and 1.75 <= ratio <= 1.80:
        results.append(f"+5：第4页背景照片显示宽高比 {ratio:.4f} 落在 1.75-1.80，与页面16:9一致")
    else:
        results.append("+0：第4页背景照片显示宽高比落在 1.75-1.80，与页面16:9一致")

    # +5：第4页背景照片横向满版铺开——横向偏移落在0-60000 EMU、显示宽度落在
    #     12100000-12260000 EMU，页面左右两侧均不出现宽于0.2cm的竖向空白条
    # 严格对齐细则，仅踩细则明确写出的三点：
    #   ① 横向偏移 x ∈ [0, 60000]
    #   ② 显示宽度 cx ∈ [12100000, 12260000]
    #   ③ 左侧空白 x、右侧空白 12192000-(x+cx) 均 ≤ 0.2cm（72000 EMU）
    # 与办公软件对齐：a:off/@x 与 a:ext/@cx 就是"位置-水平"与"大小-宽度"。
    h_ok = pic_adjusted and bool(box) and 0 <= box[0] <= 60000 and 12100000 <= box[2] <= 12260000
    if h_ok:
        h_ok = box[0] <= TOL_BLANK and (PAGE_W - (box[0] + box[2])) <= TOL_BLANK
    if h_ok:
        results.append(f"+5：第4页背景照片横向满版铺开（横向偏移{box[0]}、显示宽度{box[2]} EMU，左右无宽于0.2cm的竖向空白条）")
    else:
        results.append("+0：第4页背景照片横向满版铺开（横向偏移0-60000、显示宽度12100000-12260000 EMU，左右无宽于0.2cm的竖向空白条）")

    # +5：第4页背景照片纵向满版铺开且显示高度收敛到页面高度量级——照片顶边纵向偏移
    #     落在-60000至60000 EMU、显示高度落在6800000-6900000 EMU，页面上下均不出现
    #     高于0.2cm的横向空白带，不出现显示高度达页面高度1.5倍以上并向上下各溢出百万级EMU
    # 严格对齐细则，仅踩细则明确写出的三点：
    #   ① 纵向偏移 y ∈ [-60000, 60000]
    #   ② 显示高度 cy ∈ [6800000, 6900000]（因此自然排除 cy ≥ 1.5×6858000 的溢出摆放）
    #   ③ 上方空白 max(y,0)、下方空白 6858000-(y+cy) 均 ≤ 0.2cm（72000 EMU）
    # 与办公软件对齐：a:off/@y 与 a:ext/@cy 就是"位置-垂直"与"大小-高度"。
    v_ok = pic_adjusted and bool(box) and -60000 <= box[1] <= 60000 and 6800000 <= box[3] <= 6900000
    if v_ok:
        v_ok = (max(box[1], 0) <= TOL_BLANK
                and (PAGE_H - (box[1] + box[3])) <= TOL_BLANK
                and box[3] < int(PAGE_H * 1.5))
    if v_ok:
        results.append(f"+5：第4页背景照片纵向满版铺开（纵向偏移{box[1]}、显示高度{box[3]} EMU，上下无高于0.2cm的横向空白带）")
    else:
        results.append("+0：第4页背景照片纵向满版铺开且显示高度收敛到6800000-6900000 EMU（纵向偏移-60000至60000，上下无高于0.2cm的横向空白带）")

    # +3：第4页画面完整呈现"左侧黄色高浊度水箱—中部白蓝色膜堆—右侧透明清水箱"
    #     整条工艺链路，三段设备主体均未被页面边界裁掉
    # 严格对齐细则，仅踩细则明确写出的点：
    #   ① 照片显示矩形四边均落在页面内（可见比例≥0.98），横向三段设备主体与纵向
    #      设备本体都不被页面边界切掉；
    #   ② blipFill 未设 a:srcRect 裁剪（裁剪量为 0），画面未被人为裁去左/中/右任一段。
    # 与办公软件对齐：页面边界之外的部分在放映与导出中都不显示，等价于被裁掉；
    #   a:srcRect 即"图片格式-裁剪"的四边裁剪量（千分之一百分比）。
    chain_ok = False
    if pic_adjusted and box:
        vis_w = min(box[0] + box[2], PAGE_W) - max(box[0], 0)
        vis_h = min(box[1] + box[3], PAGE_H) - max(box[1], 0)
        frac_w = vis_w / box[2] if box[2] else 0
        frac_h = vis_h / box[3] if box[3] else 0
        chain_ok = (frac_w >= 0.98 and frac_h >= 0.98
                    and _pic_srcrect(pic) == (0, 0, 0, 0))
    if chain_ok:
        results.append("+3：第4页完整呈现左侧高浊度水箱—中部膜堆—右侧清水箱整条工艺链路，三段设备主体未被页面边界裁掉")
    else:
        results.append("+0：第4页完整呈现左侧高浊度水箱—中部膜堆—右侧清水箱整条工艺链路，三段设备主体未被页面边界裁掉")

    # +3：第4页背景照片位于该页图层最底层（spTree中的第一个图形），右侧深色信息卡、
    #     标题、副标题、三条要点、4个青色标签及底部说明文字全部叠压在照片之上且
    #     完整可读、无被照片盖住
    # 严格对齐细则，仅踩细则明确写出的两点：
    #   ① 照片是 spTree 中的第一个图形（叠放次序最底层），其余 14 个对象都在其之后；
    #   ② 这些文字对象完整落在页面范围内（完整可读），未被照片压住。
    # 与办公软件对齐：PowerPoint 的叠放次序即 spTree 文档序，越靠前越靠下。
    bottom_ok = pic_adjusted and bool(shapes) and shapes[0][0] == 'pic'
    if bottom_ok:
        bottom_ok = all(_inside(_xfrm(s) or (0, 0, 0, 0), (0, 0, PAGE_W, PAGE_H),
                                tol=EMU_CM // 10) for s in text_shapes)
    if bottom_ok:
        results.append("+3：第4页背景照片位于图层最底层，信息卡与全部文字叠压在照片之上且完整可读")
    else:
        results.append("+0：第4页背景照片位于图层最底层，信息卡与全部文字叠压在照片之上且完整可读")

    # +1：第4页背景照片仍为可单独点选、可继续拖动与缩放的独立图片对象，
    #     保留等比缩放锁定（picLocks noChangeAspect="1"）
    # 严格对齐细则，仅踩细则明确写出的两点：
    #   ① 照片仍是独立的 p:pic 图片对象（未被合并进组合/图形框）；
    #   ② p:cNvPicPr/a:picLocks/@noChangeAspect = "1"。
    # 与办公软件对齐：noChangeAspect 即"设置图片格式-锁定纵横比"。
    lock_ok = pic_adjusted and pic is not None and _pic_lock_aspect(pic)
    if lock_ok:
        results.append("+1：第4页背景照片仍为独立图片对象并保留等比缩放锁定 picLocks noChangeAspect=\"1\"")
    else:
        results.append("+0：第4页背景照片仍为独立图片对象并保留等比缩放锁定 picLocks noChangeAspect=\"1\"")

    # -5：第4页背景照片被替换成重新生成的位图——图片引用不指向包内 ppt/media/image4.jpeg、
    #     ppt/media 下图片数量不为5个、嵌入像素尺寸不满足1467×941、改以PNG等格式重新编码、
    #     图形编号不为 id="2" 或名称不为"Picture 1"或替换文字不为"crop_filter.jpg"，
    #     上述任一项成立即扣
    # 严格对齐细则，仅踩细则明确写出的五点（任一成立即扣）：
    #   ① a:blip/@r:embed 经 slide4.xml.rels 解析后的 Target 为 ../media/image4.jpeg；
    #   ② ppt/media 下图片数量为 5 个；
    #   ③ 被引用位图的嵌入像素尺寸为 1467×941；
    #   ④ 被引用位图仍为 JPEG 编码（未改成 PNG 等格式重新编码）；
    #   ⑤ p:cNvPr 的 id="2"、name="Picture 1"、descr="crop_filter.jpg"。
    # 与办公软件对齐：①②④对应包内实际引用与媒体资源，③为位图自身像素尺寸
    #   （PowerPoint"压缩图片/裁剪"会改变它），⑤为选择窗格中的对象名与替换文字。
    rid = None if pic is None else _pic_blip_rid(pic)
    tgt = rels4.get(rid, ('', ''))[1] if rid else ''
    media = deck.media()
    px = fmt = None
    if tgt:
        part = os.path.normpath(os.path.join('ppt/slides', tgt)).replace('\\', '/')
        if deck.has(part):
            w, h, fmt = _img_px(deck.raw(part))
            px = (w, h)
    cpr4 = None if pic is None else _cnvpr('pic', pic)
    id_ok = (cpr4 is not None and cpr4.get('id') == S4_PIC_ID
             and cpr4.get('name') == S4_PIC_NAME and cpr4.get('descr') == S4_PIC_DESCR)
    pic_kept = (tgt == S4_PIC_TARGET and len(media) == MEDIA_COUNT
                and px == S4_PIC_PX and fmt == 'jpeg' and id_ok)
    if pic_kept:
        results.append("-0：第4页背景照片未被替换成重新生成的位图（仍引用 ppt/media/image4.jpeg 的 1467×941 JPEG，媒体资源 5 个，id=\"2\"/名称 Picture 1/替换文字 crop_filter.jpg）")
    else:
        results.append(f"-5：第4页背景照片被替换成重新生成的位图（引用 {tgt or '缺失'}、媒体资源 {len(media)} 个、像素 {px}、编码 {fmt}）")

    # -5：交付文件体积超过2.6MB（第4页仅需调整图片摆放，合理结果应保持在2.2-2.5MB），
    #     因重复嵌入大体积位图导致文件膨胀
    # 严格对齐细则，仅踩细则明确写出的点：交付 .pptx 的文件体积 ≤ 2.6MB。
    # 与办公软件对齐：即资源管理器中看到的文件大小（2.6MB = 2.6×1024×1024 字节）。
    if deck.size <= SIZE_LIMIT:
        results.append(f"-0：交付文件体积未超过2.6MB（实际 {deck.size / 1048576:.2f}MB）")
    else:
        results.append(f"-5：交付文件体积超过2.6MB（实际 {deck.size / 1048576:.2f}MB）")

    # -5：第4页三条要点文字的内容、条数或换行结构被改动，不满足微软雅黑、15磅、
    #     非加粗、白色FFFFFF、段后间距6磅，或其文本框不在横向7333488、纵向2468880、
    #     尺寸4133088×2057400 EMU
    # 严格对齐细则，仅踩细则明确写出的四点：
    #   ① 三条要点内容逐条一致、条数为 3、换行结构为同段内 2 个 <a:br>；
    #   ② 字体微软雅黑、字号 15 磅（sz="1500"）、非加粗（无 b 或 b="0"）、白色 FFFFFF；
    #   ③ 段后间距 6 磅（a:spcAft/a:spcPts@val="600"）；
    #   ④ 文本框 a:off=(7333488, 2468880)、a:ext=(4133088, 2057400)。
    # 与办公软件对齐：sz 为百分点、spcPts 为百分磅，均与"字体/段落"对话框一致。
    bul = roles.get('bullets')
    bul_ok = bul is not None and _xfrm(bul[1]) == S4_BUL_BOX and _bullets_ok(bul[1], S4_BULLETS)
    if bul_ok:
        results.append("-0：第4页三条要点文字的内容、条数与换行结构未改动，满足微软雅黑15磅非加粗白色FFFFFF、段后6磅，文本框在7333488/2468880、4133088×2057400 EMU")
    else:
        results.append("-5：第4页三条要点文字的内容、条数、换行结构或微软雅黑15磅非加粗白色FFFFFF/段后6磅/文本框7333488·2468880·4133088×2057400 EMU 被改动")

    # -5：第4页4个青色圆角标签的文字不再分别为"预处理""膜分离""压差判读""回用缓冲"，
    #     或数量不为4个、不再按上下两行左右两列排布
    # 严格对齐细则，仅踩细则明确写出的三点：
    #   ① 4 个标签文字恰为 预处理 / 膜分离 / 压差判读 / 回用缓冲；
    #   ② 青色圆角标签数量为 4 个（标签文字同为 4 个）；
    #   ③ 4 个标签按上下两行左右两列排布：横向取 2 个不同值、纵向取 2 个不同值，
    #      且 2×2 四个位置各占其一。
    # 与办公软件对齐：a:off 的 x/y 即标签在页面上的实际位置。
    tags = [roles.get('tag%d' % i) for i in (1, 2, 3, 4)]
    tagts = [roles.get('tagtext%d' % i) for i in (1, 2, 3, 4)]
    tag_layout_ok = all(t is not None for t in tags) and all(t is not None for t in tagts)
    if tag_layout_ok:
        pos = [_xfrm(t[1])[:2] for t in tags]
        xs, ys = sorted({p[0] for p in pos}), sorted({p[1] for p in pos})
        tag_layout_ok = (len(xs) == 2 and len(ys) == 2
                         and sorted(pos) == sorted((x, y) for y in ys for x in xs))
    tag_text_ok = tag_layout_ok and sorted(_text_of(t[1]) for t in tagts) == sorted(S4_TAGS)
    if tag_text_ok:
        results.append("-0：第4页4个青色圆角标签的文字仍分别为预处理/膜分离/压差判读/回用缓冲，数量为4个且按上下两行左右两列排布")
    else:
        results.append("-5：第4页4个青色圆角标签的文字、数量或上下两行左右两列排布被改动")

    # -5：第1、2、3、5页被改动——背景照片未分别保持为第1页横向0/纵向635/尺寸
    #     12192635×6966585 EMU、第2页横向0/纵向0/尺寸12221845×6858000 EMU、第3页横向0/
    #     纵向-674370/尺寸12085320×8206105 EMU、第5页横向635/纵向-635/尺寸
    #     12713970×6858635 EMU，或这4页的文字内容、字体、字号、颜色、对齐方式、
    #     文本框位置尺寸发生任何改动
    # 严格对齐细则，仅踩细则明确写出的两点（逐页比对写死的绝对值）：
    #   ① 4 页背景照片的 a:off/a:ext 分别等于上述值；
    #   ② 4 页每个对象的文字内容、微软雅黑字体、字号、颜色、对齐方式、
    #      文本框 a:off/a:ext 均与写死值一致（顺带核对信息卡/装饰条/标签的填充描边）。
    # 与办公软件对齐：逐项对应"大小和位置"面板与"字体/段落"对话框中的设置。
    others_ok = True
    bad_pages = []
    for idx in sorted(OTHER_SLIDES):
        try:
            ok = _deco_slide_ok(_shapes(deck.sp_tree(idx)), OTHER_SLIDES[idx])
        except Exception:
            ok = False
        if not ok:
            others_ok = False
            bad_pages.append(idx)
    if others_ok:
        results.append("-0：第1、2、3、5页的背景照片摆放与文字内容/字体/字号/颜色/对齐方式/文本框位置尺寸均未改动")
    else:
        results.append(f"-5：第{'、'.join(str(i) for i in bad_pages)}页的背景照片摆放或文字内容/字体/字号/颜色/对齐方式/文本框位置尺寸被改动")

    # -3：ppt/media/image4.jpeg 在包内不再被任何图形引用而成为冗余资源，
    #     且 slide4.xml.rels 中残留无对应引用的失效图片关系
    # 严格对齐细则，仅踩细则明确写出的两点（须同时成立才扣）：
    #   ① ppt/media/image4.jpeg 存在，但全包 5 张幻灯片里没有任何图形通过
    #      r:embed/r:link 引用它 → 成为冗余资源；
    #   ② slide4.xml.rels 中存在 Type=image 的关系，其 Id 未被 slide4.xml 里
    #      任何 r:embed/r:link 使用 → 失效图片关系残留。
    # 与办公软件对齐：这类冗余资源与失效关系在 PowerPoint 界面上看不到，
    #   却会一直留在文件里增大体积、并在"检查文档"时报出问题。
    def _used_rids(root):
        used = set()
        for el in root.iter():
            for k, v in el.attrib.items():
                if k.startswith('{%s}' % NS_R):
                    used.add(v)
        return used

    refs_all = set()
    for i in range(1, deck.slide_count() + 1):
        part = 'ppt/slides/slide%d.xml' % i
        if not deck.has(part):
            continue
        used_i = _used_rids(deck.xml(part))
        for rid_i, (typ, tg) in deck.rels(part).items():
            if typ == 'image' and rid_i in used_i:
                refs_all.add(os.path.normpath(os.path.join('ppt/slides', tg)).replace('\\', '/'))
    used4 = _used_rids(deck.xml('ppt/slides/slide4.xml'))
    orphan_img = ('ppt/media/image4.jpeg' in media
                  and 'ppt/media/image4.jpeg' not in refs_all)
    stale_rels = [rid_i for rid_i, (typ, _tg) in rels4.items()
                  if typ == 'image' and rid_i not in used4]
    if not (orphan_img and stale_rels):
        results.append("-0：ppt/media/image4.jpeg 仍被图形引用，slide4.xml.rels 中无失效图片关系残留")
    else:
        results.append(f"-3：ppt/media/image4.jpeg 成为无人引用的冗余资源，且 slide4.xml.rels 残留失效图片关系 {'/'.join(sorted(stale_rels))}")

    # -3：第4页右侧深色圆角矩形信息卡不满足横向偏移7022592、纵向偏移502920、
    #     尺寸4754880×5852160 EMU，或填充色不满足0D3B42、描边不满足白色FFFFFF实线
    # 严格对齐细则，仅踩细则明确写出的三点：
    #   ① a:off=(7022592, 502920)、a:ext=(4754880, 5852160)；
    #   ② 填充 a:solidFill/a:srgbClr@val = 0D3B42；
    #   ③ 描边 a:ln 为实线且颜色 FFFFFF。
    # 与办公软件对齐：即"设置形状格式"中的填充色与轮廓色。
    card = roles.get('card')
    card_ok = (card is not None and _xfrm(card[1]) == S4_CARD
               and _prst(card[1]) == 'roundRect'
               and _fill(card[1]) == S4_CARD_FILL
               and _line(card[1]) == ('solid', S4_CARD_LINE))
    if card_ok:
        results.append("-0：第4页右侧深色圆角信息卡仍为7022592/502920、4754880×5852160 EMU，填充0D3B42、白色FFFFFF实线描边")
    else:
        results.append("-3：第4页右侧深色圆角信息卡的位置尺寸7022592·502920·4754880×5852160 EMU、填充0D3B42或白色FFFFFF实线描边被改动")

    # -3：第4页标题"多级膜分离与清水回用链路"不满足微软雅黑、28磅、加粗、白色FFFFFF、
    #     左对齐，或其文本框不在横向7278624、纵向758952、尺寸4242816×640080 EMU
    # 严格对齐细则，仅踩细则明确写出的三点：
    #   ① 文字内容为"多级膜分离与清水回用链路"；
    #   ② 微软雅黑、sz="2800"、b="1"、颜色 FFFFFF、a:pPr@algn="l"；
    #   ③ 文本框 a:off=(7278624, 758952)、a:ext=(4242816, 640080)。
    # 与办公软件对齐：sz 为百分点（28磅=2800），algn="l" 即左对齐。
    title = roles.get('title')
    title_ok = (title is not None and _xfrm(title[1]) == S4_TITLE_BOX
                and _one_para_text(title[1], S4_TITLE_TEXT, 'l', '2800', True, WHITE))
    if title_ok:
        results.append("-0：第4页标题\"多级膜分离与清水回用链路\"仍为微软雅黑28磅加粗白色FFFFFF左对齐，文本框在7278624/758952、4242816×640080 EMU")
    else:
        results.append("-3：第4页标题\"多级膜分离与清水回用链路\"的微软雅黑28磅加粗白色FFFFFF左对齐或文本框7278624·758952·4242816×640080 EMU 被改动")

    # -3：第4页副标题"预处理、膜组件与透明水箱构成连续式中试平台"不满足微软雅黑、
    #     14磅、非加粗、颜色E9F2F6、左对齐，或其文本框不在横向7278624、纵向1581912、
    #     尺寸4242816×777240 EMU
    # 严格对齐细则，仅踩细则明确写出的三点：
    #   ① 文字内容为"预处理、膜组件与透明水箱构成连续式中试平台"；
    #   ② 微软雅黑、sz="1400"、非加粗、颜色 E9F2F6、algn="l"；
    #   ③ 文本框 a:off=(7278624, 1581912)、a:ext=(4242816, 777240)。
    # 与办公软件对齐：颜色取 a:solidFill/a:srgbClr@val，即字体颜色的 RGB 值。
    sub = roles.get('subtitle')
    sub_ok = (sub is not None and _xfrm(sub[1]) == S4_SUB_BOX
              and _one_para_text(sub[1], S4_SUB_TEXT, 'l', '1400', False, S4_SUB_COLOR))
    if sub_ok:
        results.append("-0：第4页副标题仍为微软雅黑14磅非加粗E9F2F6左对齐，文本框在7278624/1581912、4242816×777240 EMU")
    else:
        results.append("-3：第4页副标题的微软雅黑14磅非加粗E9F2F6左对齐或文本框7278624·1581912·4242816×777240 EMU 被改动")

    # -3：第4页4个青色圆角标签的尺寸不统一为1984248×393192 EMU，左列横向偏移不为
    #     7315200、右列横向偏移不为9500616、上行纵向偏移不为4892040、下行纵向偏移不为
    #     5458968，填充色不满足2BB2B7、无描边，标签文字不满足微软雅黑、12磅、加粗、
    #     白色FFFFFF、居中，或4个标签文字框不分别位于横向7360920/9546336、
    #     纵向4910327/5477255、尺寸1892807×356616 EMU
    # 严格对齐细则，仅踩细则明确写出的四点：
    #   ① 4 个标签 a:ext 统一为 1984248×393192，横向取 7315200/9500616、
    #      纵向取 4892040/5458968，2×2 四个位置各占其一；
    #   ② 标签填充 2BB2B7 且 a:ln 为 noFill（无描边）；
    #   ③ 标签文字微软雅黑、sz="1200"、b="1"、白色 FFFFFF、algn="ctr"；
    #   ④ 4 个标签文字框 a:off 取 (7360920/9546336)×(4910327/5477255)、
    #      a:ext = 1892807×356616。
    # 与办公软件对齐：以上全部对应"大小和位置"与"设置形状格式/字体"面板中的值。
    tag_geo_ok = all(t is not None for t in tags) and all(t is not None for t in tagts)
    if tag_geo_ok:
        want_tag = sorted((x, y) + S4_TAG_EXT for y in S4_TAG_Y for x in S4_TAG_X)
        want_txt = sorted((x, y) + S4_TAGT_EXT for y in S4_TAGT_Y for x in S4_TAGT_X)
        tag_geo_ok = (sorted(_xfrm(t[1]) for t in tags) == want_tag
                      and sorted(_xfrm(t[1]) for t in tagts) == want_txt)
    if tag_geo_ok:
        for t in tags:
            if _prst(t[1]) != 'roundRect' or _fill(t[1]) != S4_ACCENT or _line(t[1])[0] != 'none':
                tag_geo_ok = False
                break
    if tag_geo_ok:
        for t in tagts:
            if not _one_para_text(t[1], _text_of(t[1]), 'ctr', '1200', True, WHITE):
                tag_geo_ok = False
                break
    if tag_geo_ok:
        results.append("-0：第4页4个青色圆角标签仍统一为1984248×393192 EMU、左右列7315200/9500616、上下行4892040/5458968、填充2BB2B7无描边，标签文字微软雅黑12磅加粗白色居中且文字框为1892807×356616 EMU")
    else:
        results.append("-3：第4页4个青色圆角标签的尺寸1984248×393192、左右列7315200/9500616、上下行4892040/5458968、填充2BB2B7无描边，或标签文字微软雅黑12磅加粗白色FFFFFF居中、文字框7360920/9546336·4910327/5477255·1892807×356616 EMU 被改动")

    # -3：第4页底部说明文字"工艺说明：预处理、膜组件与透明水箱构成连续式中试平台"的
    #     内容被改动，不满足微软雅黑、8磅、非加粗、白色FFFFFF、右对齐，或其文本框不在
    #     横向411480、纵向6446520、尺寸11384280×256032 EMU
    # 严格对齐细则，仅踩细则明确写出的三点：
    #   ① 文字内容为"工艺说明：预处理、膜组件与透明水箱构成连续式中试平台"；
    #   ② 微软雅黑、sz="800"、非加粗、白色 FFFFFF、algn="r"；
    #   ③ 文本框 a:off=(411480, 6446520)、a:ext=(11384280, 256032)。
    # 与办公软件对齐：algn="r" 即右对齐。
    foot = roles.get('footer')
    foot_ok = (foot is not None and _xfrm(foot[1]) == S4_FOOT_BOX
               and _one_para_text(foot[1], S4_FOOT_TEXT, 'r', '800', False, WHITE))
    if foot_ok:
        results.append("-0：第4页底部说明文字内容未改动，仍为微软雅黑8磅非加粗白色FFFFFF右对齐，文本框在411480/6446520、11384280×256032 EMU")
    else:
        results.append("-3：第4页底部说明文字的内容、微软雅黑8磅非加粗白色FFFFFF右对齐或文本框411480·6446520·11384280×256032 EMU 被改动")

    # -3：第4页新增了指令未要求的元素（新增文本框、色块、半透明蒙版、遮罩、外框线、
    #     水印、页码、箭头或图标等），或文档中新增了备注页文字、批注、宏
    # 严格对齐细则，仅踩细则明确写出的三点：
    #   ① 第4页可单独点选的对象数未超过 15（多出的文本框/色块/蒙版/遮罩/外框线/水印/
    #      页码/箭头/图标都会使该数增加；少于 15 属于"删除对象"，由后一条判定）；
    #   ② 包内无带文字的备注页（notesSlide）；
    #   ③ 包内无批注部件、无 vbaProject.bin 宏工程。
    # 与办公软件对齐：对象数即选择窗格中的条目数；备注/批注/宏均为独立包部件。
    extra_notes = []
    for n in sorted(deck.names):
        if re.match(r'ppt/notesSlides/notesSlide\d+\.xml$', n):
            body = re.sub(r'<[^>]+>', '', deck.raw(n).decode('utf-8', 'ignore'))
            if body.strip():
                extra_notes.append(n)
    has_comment = any(('comment' in n.lower()) for n in deck.names)
    has_macro = any(n.lower().endswith('vbaproject.bin') for n in deck.names)
    added_ok = (len(shapes) <= 15 and not extra_notes
                and not has_comment and not has_macro)
    if added_ok:
        results.append("-0：第4页未新增指令未要求的元素，文档未新增备注页文字、批注或宏")
    else:
        results.append(f"-3：第4页新增了指令未要求的元素（对象数 {len(shapes)}，超过15个）或新增了备注页文字/批注/宏")

    # -3：第4页删除了15个既有对象中的任一个（满版照片、深色圆角信息卡、标题文本框、
    #     青色装饰条、副标题文本框、要点文本框、4个青色圆角标签、4个标签文字框、
    #     底部说明文本框）
    # 严格对齐细则，仅踩细则明确写出的点：15 个角色对象逐一仍在（按对象自身特征
    #   识别：图片、圆角信息卡、细长色条、4 个圆角标签，以及按内容或原位置识别的
    #   标题/副标题/要点/4 个标签文字/底部说明文本框）。
    # 与办公软件对齐：对应选择窗格里 15 个条目一个都不少。
    missing = [r for r in S4_ROLES if r not in roles]
    if not missing and len(shapes) >= 15:
        results.append("-0：第4页15个既有对象（照片、信息卡、标题、装饰条、副标题、要点、4个标签及标签文字、底部说明）均未被删除")
    else:
        results.append(f"-3：第4页删除了既有对象（缺失 {'、'.join(missing) or '对象数不足15'}）")

    # -3：第4页对象的叠放次序被打乱，未保持照片在最底层、其后依次为深色信息卡、标题、
    #     青色装饰条、副标题、要点文字、4组标签与标签文字、底部说明文字
    # 严格对齐细则，仅踩细则明确写出的点：spTree 文档序对应的角色序列须为
    #   照片→信息卡→标题→装饰条→副标题→要点→(标签+标签文字)×4→底部说明
    #   （新增的额外对象不参与本条判定，由"新增元素"一条负责）。
    # 与办公软件对齐：spTree 文档序即 PowerPoint 的叠放次序（前=底层）。
    known_order = [o for o in order if o != 'unknown']
    if missing:
        # 对象已缺失时不重复扣叠放次序（缺失由上一条判定）
        results.append("-0：第4页对象叠放次序（照片在最底层，其后依次为信息卡、标题、装饰条、副标题、要点、4组标签与标签文字、底部说明）")
    elif known_order == S4_ORDER:
        results.append("-0：第4页对象叠放次序仍为照片在最底层，其后依次为信息卡、标题、装饰条、副标题、要点、4组标签与标签文字、底部说明")
    else:
        results.append("-3：第4页对象的叠放次序被打乱，未保持照片在最底层及其后信息卡、标题、装饰条、副标题、要点、4组标签与标签文字、底部说明的次序")

    # -3：演示文稿页数不为5页、第4页不再套用 slideLayout2 版式、页面尺寸不为
    #     12192000×6858000 EMU，或主题 theme1、母版 slideMaster1、11个版式
    #     slideLayout1-slideLayout11 的内容被改动
    # 严格对齐细则，仅踩细则明确写出的四点：
    #   ① p:sldIdLst 中幻灯片数为 5；
    #   ② slide4.xml.rels 的 slideLayout 关系指向 ../slideLayouts/slideLayout2.xml；
    #   ③ p:sldSz 为 cx=12192000、cy=6858000；
    #   ④ theme1.xml、slideMaster1.xml、slideLayout1-11.xml 的语义指纹不变
    #      （规范化后忽略 XML 声明、属性顺序、空白等纯序列化差异）。
    # 与办公软件对齐：②即"版式"下拉中选中的版式，③即"幻灯片大小"，
    #   ④对应"幻灯片母版"视图中的主题与 11 个版式。
    layout_tgt = [tg for typ, tg in rels4.values() if typ == 'slideLayout']
    struct_ok = (deck.slide_count() == SLIDE_COUNT
                 and layout_tgt == [SLIDE4_LAYOUT]
                 and deck.page_size() == (PAGE_W, PAGE_H))
    if struct_ok:
        for part, sig in STATIC_PART_SIG.items():
            if not deck.has(part) or _canon_sig(deck.raw(part)) != sig:
                struct_ok = False
                break
    if struct_ok:
        results.append("-0：演示文稿仍为5页、第4页套用 slideLayout2、页面尺寸12192000×6858000 EMU，主题 theme1、母版 slideMaster1 与11个版式内容未改动")
    else:
        results.append("-3：演示文稿页数、第4页 slideLayout2 版式、页面尺寸12192000×6858000 EMU 或主题 theme1/母版 slideMaster1/11个版式的内容被改动")

    # -1：第4页标题下方青色装饰条不满足横向7333488、纵向1472184、尺寸1828800×27432 EMU、
    #     填充色2BB2B7、无描边
    # 严格对齐细则，仅踩细则明确写出的三点：
    #   ① a:off=(7333488, 1472184)、a:ext=(1828800, 27432)；
    #   ② 填充 2BB2B7；③ a:ln 为 noFill（无描边）。
    # 与办公软件对齐：即该细长矩形的"大小和位置"与"填充/轮廓"设置。
    bar = roles.get('bar')
    bar_ok = (bar is not None and _xfrm(bar[1]) == S4_BAR
              and _fill(bar[1]) == S4_ACCENT and _line(bar[1])[0] == 'none')
    if bar_ok:
        results.append("-0：第4页标题下方青色装饰条仍为7333488/1472184、1828800×27432 EMU、填充2BB2B7、无描边")
    else:
        results.append("-1：第4页标题下方青色装饰条的7333488·1472184·1828800×27432 EMU、填充2BB2B7或无描边被改动")

    # -1：第4页各文本框的自动调整与内边距设置被改动，未保留 spAutoFit 及
    #     左54864、上18288、右54864、下18288 EMU 的内边距
    # 严格对齐细则，仅踩细则明确写出的两点（对标题、副标题、要点、4个标签文字、
    #   底部说明共 8 个文本框逐个核对）：
    #   ① a:bodyPr 下存在 a:spAutoFit；
    #   ② a:bodyPr 的 lIns/tIns/rIns/bIns 分别为 54864/18288/54864/18288。
    # 与办公软件对齐：对应"设置形状格式-文本选项-根据文字调整形状大小"
    #   与"内部边距"四项。
    inset_targets = [roles.get(k) for k in
                     ('title', 'subtitle', 'bullets', 'footer',
                      'tagtext1', 'tagtext2', 'tagtext3', 'tagtext4')]
    inset_ok = all(t is not None for t in inset_targets)
    if inset_ok:
        for t in inset_targets:
            bp = _body_pr(t[1])
            if bp is None or bp[:4] != INSETS or not bp[4]:
                inset_ok = False
                break
    if inset_ok:
        results.append("-0：第4页各文本框保留 spAutoFit 及左54864、上18288、右54864、下18288 EMU 的内边距")
    else:
        results.append("-1：第4页文本框的 spAutoFit 或左54864/上18288/右54864/下18288 EMU 内边距被改动")

    # -1：docProps/thumbnail.jpeg 预览图未随第4页版式一同更新，文档缩略图与实际页面不一致
    # 严格对齐细则，仅踩细则明确写出的点：docProps/thumbnail.jpeg 存在，且其字节
    #   不再等于"第4页调整前"的那张预览图（STALE_THUMB_MD5）。
    # 前提：第4页版式确有改动（s4_touched）。第4页原样未动时预览图与实际页面本就
    #   一致，不存在"未随之更新"的问题，不予扣分。
    # 与办公软件对齐：PowerPoint 正常保存会重新生成该预览图，资源管理器/
    #   文件属性里看到的缩略图即取自该部件。
    thumb = 'docProps/thumbnail.jpeg'
    thumb_stale = s4_touched and not (
        deck.has(thumb) and hashlib.md5(deck.raw(thumb)).hexdigest() != STALE_THUMB_MD5)
    if not thumb_stale:
        results.append("-0：docProps/thumbnail.jpeg 预览图已随第4页版式一同更新")
    else:
        results.append("-1：docProps/thumbnail.jpeg 预览图未随第4页版式一同更新，文档缩略图与实际页面不一致")

    # ========== 将 results 反解为结构化 dim2_items ==========
    # 逐项按追加顺序对齐 _MAX_DELTAS_ORDER：解析每条 "±N：说明" 的 N 与 rule。
    # 加分项：命中记 N=max_delta，未命中记 N=0；
    # 减分项：踩中记 N=max_delta（负值），未踩中记 N=0；
    # hit 统一表示"本条的 delta 被计入"（加分项=拿到分，减分项=被扣分）。
    if len(results) != len(_MAX_DELTAS_ORDER):
        # 顺序表与实际条目不一致：视为脚本内部错误，交给外层兜底
        raise RuntimeError(
            f"dim2 项数不一致：results={len(results)} vs order={len(_MAX_DELTAS_ORDER)}"
        )
    dim2_items = []
    total_score = 0
    for raw, max_delta in zip(results, _MAX_DELTAS_ORDER):
        head, _, rule = raw.partition("：")
        try:
            delta = int(head.strip().lstrip('+'))
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
