#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# pyright: reportUnknownMemberType=false, reportUnknownVariableType=false, reportUnknownArgumentType=false, reportUnknownParameterType=false, reportMissingParameterType=false, reportPrivateUsage=false, reportUnknownLambdaType=false, reportGeneralTypeIssues=false

import io
import os
import re
import sys
import json
import struct
import shutil
import hashlib
import zipfile
import tempfile
import traceback
from docx import Document
from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph

# 说明：按接口统一约定，本模块不修改全局 sys.stdout，也不 print 主结果。
# 主结果统一由 evaluate(dir_path) -> dict 返回；入参为“脚本所在目录的路径”，
# 由脚本自己在该目录内定位并打开被评估的 .docx；若目录内没有 .docx 而只有
# .zip／.rar 压缩包，则解压到临时目录后在包内定位 .docx，并记下“交付物为压缩包”
# 这一事实供减分项使用。本文件仅在 __main__ 下作为本地调试入口，才把 dict
# 序列化为 JSON 打印。

SCRIPT_ID = "059"
# 维度二所有加分项 max_delta 之和；用于维度一未通过 / 脚本异常时兜底
MAX_SCORE_TOTAL = 30
# 维度二所有减分项 max_delta 之和，即全部踩坑时的下限
MIN_SCORE_TOTAL = -44

# 维度二所有评分项按“结果追加顺序”的 max_delta 列表；末端把 results 里的
# "+N：..." / "-N：..." 字符串反解为结构化 dim2_items 时用。顺序须与
# _evaluate_impl 中 results.append(...) 的顺序严格一致；共 30 项，
# 正项累加 = MAX_SCORE_TOTAL，负项累加 = MIN_SCORE_TOTAL。
_MAX_DELTAS_ORDER = [
    5,   # 01 单元格内是真嵌入对象，PPT 字节写入文档包
    5,   # 02 载荷是完整可用的 12,275,624 字节／18 张幻灯片演示文稿
    5,   # 03 双击即开：ProgID／ShapeID／ObjectID／r:id 与关系表、内容类型齐备
    3,   # 04 图标方式显示 + 标准矢量图标 + 中文文件名标注
    3,   # 05 插入位置落在“项目介绍材料”行右侧“申报内容”单元格
    3,   # 06 对象尺寸 72-115pt × 66-85pt 且等比不变形
    3,   # 07 单元格内排版：居中／9 磅非加粗／≥3 空段落／行高不塌陷
    3,   # 08 末行不跨页断裂、不增空白页
    -3,  # 09 以链接对象或域链接方式插入
    -3,  # 10 交付物为压缩包
    -1,  # 11 Word 文档文件名被改写
    -3,  # 12 演示文稿被改名或字节被改动
    -5,  # 13 四处项目名不是“青梧谷灵 MR 共创计划”
    -3,  # 14 封面副标题被改动
    -3,  # 15 单元格内新增指令未要求的文字，或新增文字不满足 9 磅非加粗
    -3,  # 16 额外插入指向外部 .pptx 的超链接
    -3,  # 17 图标不是 PowerPoint 标准矢量图标
    -3,  # 18 申报表行列结构被改动
    -3,  # 19 24 个申报事项名称或正文被删改
    -1,  # 20 表头行底色／字色／字号／居中
    -1,  # 21 封面首行“项目申报书”字体格式
    -1,  # 22 三色块一行文字／字色／底色
    -1,  # 23 “申报单位”“申报日期”两行
    -1,  # 24 左列申报事项标签格式
    -1,  # 25 右列申报内容正文格式
    -1,  # 26 表格框线颜色与单元格内边距
    -1,  # 27 页面设置
    -1,  # 28 页脚“第 X 页”页码域
    -1,  # 29 新增指令未要求的页面／章节
    -1,  # 30 遗留临时占位内容
]
assert len(_MAX_DELTAS_ORDER) == 30
assert sum(d for d in _MAX_DELTAS_ORDER if d > 0) == MAX_SCORE_TOTAL
assert sum(d for d in _MAX_DELTAS_ORDER if d < 0) == MIN_SCORE_TOTAL

# ---------- 裸命名空间（w:object / v:shape / o:OLEObject 需按原始 XML 取属性） ----------
_W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
_R = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'
_O = '{urn:schemas-microsoft-com:office:office}'
_V = '{urn:schemas-microsoft-com:vml}'

# ---------- 交付物应有的硬编码期望值 ----------
DOCX_NAME = '青梧谷灵MR共创计划_项目申报书.docx'
PPTX_NAME = '青梧谷灵MR共创计划.pptx'
PPTX_SIZE = 12275624
PPTX_MD5 = '17636ecb181d591c6db01915f537eb81'
PPTX_SLIDES = 18
PPTX_SLIDE1_KEYS = ('青梧谷灵', 'MR共创计划', '以滨星手作、互动叙事与数字展演，重构区域文化体验')
PROJECT_NAME = '青梧谷灵MR共创计划'
COVER_TITLE = '项目申报书'
COVER_SUBTITLE = '以滨星手作、互动叙事与数字展演，重构区域文化体验'
FORM_TITLE = '青梧谷灵MR共创计划项目申报表'
HEADER_TEXT = '青梧谷灵MR共创计划·项目申报书'
TARGET_LABEL = '项目介绍材料'
UNIT_LINE = '申报单位：' + '_' * 26
DATE_LINE = '申报日期：2026年5月'
CHIP_TEXTS = ('文化科技融合', '沉浸式文旅', '共创工坊')
CHIP_FILLS = ('EAF4F2', 'D9EAF7', 'EAF4F2')
HEAD_CELLS = ('申报事项', '申报内容')
# 申报表隔行浅色底纹：表头 1F4E79；此后奇数行浅色、偶数行更浅；末行“项目介绍材料”单独配色
LABEL_FILLS = tuple(['1F4E79']
                    + ['EAF4F2' if i % 2 else 'F7FBFA' for i in range(1, 24)]
                    + ['D9EAF7'])
BODY_FILLS = tuple(['1F4E79']
                   + ['FFFFFF' if i % 2 else 'FBFBFB' for i in range(1, 24)]
                   + ['FFFFFF'])
SERIF = 'Noto Serif CJK SC'
SANS = 'Noto Sans CJK SC'
# 申报内容列／封面正文里的中文 run 有两种落地写法：显式 Noto Serif CJK SC，或
# w:hint="eastAsia" 落到 宋体；两者在同一列中并存、视觉一致，均按达标处理，
# 字体项的判定重心因此放在字号、加粗、底色、段后间距这些会被改坏的属性上。
CJK_SERIF_FONTS = (SERIF, '宋体')
FORM_GRID = [2381, 6917]          # 左列 4.2cm / 右列 12.2cm（twips）
CELL_MARGIN = 120                 # 0.21cm（twips）
INNER_BORDER = ('C9D6DF', '6')    # 内部框线 0.5 磅
LAST_TOP_BORDER = ('1F4E79', '12')  # 末行上框线 1 磅
PAGE_SIZE = (11906, 16838)        # A4 纵向 21cm × 29.7cm
PAGE_MARGIN = {'top': 1247, 'right': 1304, 'bottom': 1134, 'left': 1304}
BODY_SIGNATURE = ['p', 'p', 'p', 'tbl', 'p', 'p', 'p', 'p', 'tbl', 'p']
PAGE_BREAKS = 1
EXPECT_PAGES = 4
PLACEHOLDERS = ('待补充', '待完善', '此处插入', '见附件', '占位', 'TODO', 'todo')

# 24 个申报事项：(左列标签, 右列正文去空白后的字符数, 右列正文去空白后的 md5)。
# “项目介绍材料”一行是插入位置，右列正文只校标签、不校正文（正文另由单元格新增
# 文字类减分项约束）。
_EXPECT_ROWS = [
    ('项目名称', 10, '6c821056d7e17e7ebc1ae0582487465e'),
    ('项目类别', 21, '32cceb5e40f79e070f67b15a292696b3'),
    ('申报方向', 33, '9d552d4e631263f9b93e2c9b173bd7e9'),
    ('项目定位', 74, '70523e5a57fee841b89e0b2cc6253194'),
    ('项目背景', 146, 'd752f8b98c7e8cdc910d30927a1fa022'),
    ('痛点分析', 118, 'f8f109d3cce9662c4ea0a9affd69007a'),
    ('建设目标', 75, '60c63f7f2cc0289d49832e15c2014363'),
    ('建设内容', 178, '5009cfa4ea25e6b331af0af39385b3af'),
    ('产品体系', 149, '9fff05f5e928e7b8205f049e55a0ba08'),
    ('目标客群', 75, '892a65c6593124631343e17edd7d220b'),
    ('体验流程', 83, '121cbd86dca384a8f3d0ec4ac17afb12'),
    ('技术路线', 106, '8f72a9939144c216c27a5f2780924c37'),
    ('技术选型', 135, 'bc673fc8eb19b3cb1a796bb5cf118749'),
    ('竞争优势', 93, 'f322a98d6cf58383640da29d6fc051ec'),
    ('差异化能力', 111, '76867a83d79dfb389c37d11b42c4e491'),
    ('商业模式', 146, 'b985ef8910108b8e607e5b3090078ba7'),
    ('资金使用方向', 103, '086d5b8846424c192d7744c9a245bc58'),
    ('资源整合', 95, '6497f403573091896c3f14300efb98c1'),
    ('团队配置', 133, 'b7b6be4ce7898f4cd38b95bee63ac4a0'),
    ('实施计划', 92, '76ace7611ecb85b70e108aeceba81c23'),
    ('风险控制', 141, '35cd418355811331f9a1393cb7a7f581'),
    ('预期成果', 87, '8de7353b672ee0d7e4e4439f4f9e0b4d'),
    ('社会价值', 107, 'a7a5f053d6ea02852fa2f93418554bc4'),
    ('项目介绍材料', 0, 'd41d8cd98f00b204e9800998ecf8427e'),
]
assert len(_EXPECT_ROWS) == 24


def _norm(text):
    """去掉所有空白字符，用于文本比对（对齐办公软件里看到的可见字符序列）。"""
    return re.sub(r'\s+', '', text or '')


def _hexval(value):
    return (value or '').strip().upper()


def _md5(text):
    return hashlib.md5((text or '').encode('utf-8')).hexdigest()

def _fix_zip_name(name):
    """修正“以 UTF-8 写入条目名但未置 UTF-8 标志位”的压缩包条目名（中文名会乱码）。"""
    try:
        fixed = name.encode('cp437').decode('utf-8')
        return fixed
    except Exception:
        return name


class _TargetMissing(Exception):
    """目录不存在或目录中没有 .docx：连待评测对象都没有，不进入维度判定。"""


def _locate_delivery(dir_path):
    """在 dir_path 内定位被评估的 .docx（忽略 ~$ 打头的 Word 临时锁文件）。
    目录内没有 .docx 而只有 .zip／.rar 压缩包时，解压到临时目录后在包内定位。
    返回 dict：docx_path／docx_name／tmp_dir／archive_name／sibling_pptx
    （sibling_pptx 为与文档一同交付的 .pptx 的 [(名称, 字节), ...]）。"""
    if not dir_path or not os.path.isdir(dir_path):
        raise _TargetMissing("目录不存在：%s" % dir_path)
    info = {'docx_path': '', 'docx_name': '', 'tmp_dir': None,
            'archive_name': None, 'sibling_pptx': []}
    entries = sorted(os.listdir(dir_path))

    def _pick(folder, names):
        docxs = sorted(f for f in names
                       if f.lower().endswith('.docx') and not f.startswith('~$'))
        if not docxs:
            return False
        info['docx_name'] = docxs[0]
        info['docx_path'] = os.path.join(folder, docxs[0])
        for pptx in sorted(f for f in names
                           if f.lower().endswith('.pptx') and not f.startswith('~$')):
            with open(os.path.join(folder, pptx), 'rb') as fh:
                info['sibling_pptx'].append((pptx, fh.read()))
        return True

    if _pick(dir_path, entries):
        return info

    archives = [f for f in entries if f.lower().endswith(('.zip', '.rar', '.7z'))]
    if not archives:
        raise _TargetMissing("目录中未找到 .docx 文件：%s" % dir_path)
    info['archive_name'] = archives[0]
    archive = os.path.join(dir_path, archives[0])
    if not zipfile.is_zipfile(archive):
        raise RuntimeError("交付物为压缩包且非 zip 格式，无法取得 .docx：%s" % archives[0])
    info['tmp_dir'] = tempfile.mkdtemp(prefix='officeval_%s_' % SCRIPT_ID)
    with zipfile.ZipFile(archive) as zf:
        for item in zf.infolist():
            if item.filename.endswith('/'):
                continue
            out = os.path.join(info['tmp_dir'],
                               os.path.basename(_fix_zip_name(item.filename)))
            with open(out, 'wb') as fh:
                fh.write(zf.read(item))
    if not _pick(info['tmp_dir'], sorted(os.listdir(info['tmp_dir']))):
        raise _TargetMissing("压缩包内未找到 .docx 文件：%s" % archives[0])
    return info

class _Package(object):
    """按部件名读取 .docx 包内容，并解析 document.xml 的关系表与内容类型声明。"""

    _KEEP = ('[Content_Types].xml', 'word/_rels/document.xml.rels',
             'word/document.xml', 'docProps/app.xml')

    def __init__(self, docx_path):
        self._blobs = {}
        with zipfile.ZipFile(docx_path) as zf:
            self.names = [n for n in zf.namelist() if not n.endswith('/')]
            for name in self.names:
                if name in self._KEEP or name.startswith('word/embeddings/') \
                        or name.startswith('word/media/'):
                    self._blobs[name] = zf.read(name)
        self.content_types = self.text('[Content_Types].xml')
        self.document_xml = self.text('word/document.xml')
        self.app_xml = self.text('docProps/app.xml')
        self.rels = {}
        for tag in re.findall(r'<Relationship\b[^>]*>',
                              self.text('word/_rels/document.xml.rels')):
            rid = re.search(r'\bId="([^"]+)"', tag)
            typ = re.search(r'\bType="([^"]+)"', tag)
            tgt = re.search(r'\bTarget="([^"]*)"', tag)
            mode = re.search(r'\bTargetMode="([^"]+)"', tag)
            if rid is not None:
                self.rels[rid.group(1)] = {
                    'type': typ.group(1).rsplit('/', 1)[-1] if typ else '',
                    'target': tgt.group(1) if tgt else '',
                    'mode': mode.group(1) if mode else 'Internal',
                }

    def text(self, name):
        return self._blobs.get(name, b'').decode('utf-8', 'ignore')

    def blob(self, name):
        return self._blobs.get(name)

    def part_of(self, rid):
        """r:id → 包内部件名；外部关系或部件不存在时返回 None。"""
        rel = self.rels.get(rid or '')
        if not rel or rel['mode'] == 'External':
            return None
        target = rel['target'].lstrip('/')
        name = target if target.startswith('word/') else 'word/' + target
        return name if name in self.names else None

    def declares_extension(self, ext):
        """[Content_Types].xml 是否以 Default 声明了该扩展名。"""
        if not ext:
            return False
        return re.search(r'<Default[^>]+Extension="%s"' % re.escape(ext),
                         self.content_types, re.I) is not None

    def declares_part(self, part_name):
        """[Content_Types].xml 是否为该部件单独声明了 Override。"""
        if not part_name:
            return False
        return re.search(r'<Override[^>]+PartName="/%s"' % re.escape(part_name),
                         self.content_types, re.I) is not None

    def declared_pages(self):
        found = re.search(r'<Pages>(\d+)</Pages>', self.app_xml)
        return int(found.group(1)) if found else None

def _emf_texts(blob):
    """遍历 EMF 记录，抽取 EMR_EXTTEXTOUTA(83)／EMR_EXTTEXTOUTW(84) 绘制的文本。
    图标下方的文件名标注在 EMF 里按行拆成多条记录，且记录内文本之后紧跟字符间距
    数组，所以必须按记录定位 offString/Chars 取出再拼接——直接在字节流里找连续的
    文件名串是找不到的。"""
    texts, pos, total = [], 0, len(blob)
    while pos + 8 <= total:
        try:
            itype, rec_size = struct.unpack_from('<II', blob, pos)
        except struct.error:
            break
        if rec_size < 8 or pos + rec_size > total:
            break
        if itype in (83, 84):
            rec = blob[pos:pos + rec_size]
            if len(rec) >= 52:
                chars, off_str = struct.unpack_from('<II', rec, 44)
                width = 2 if itype == 84 else 1
                if chars and off_str and off_str + chars * width <= len(rec):
                    raw = rec[off_str:off_str + chars * width]
                    try:
                        texts.append(raw.decode('utf-16-le' if width == 2 else 'gbk',
                                                'ignore'))
                    except Exception:
                        pass
        pos += rec_size
    return texts


def _wmf_texts(blob):
    """WMF 图元的 META_EXTTEXTOUT 文本为 ANSI，按可读片段粗取即可满足标注判定。"""
    out = []
    for enc in ('gbk', 'latin-1'):
        try:
            decoded = blob.decode(enc, 'ignore')
        except Exception:
            continue
        out.extend(re.findall(r'[0-9A-Za-z一-鿿._()（）-]{2,60}', decoded))
    return out


def _icon_caption(blob, ext):
    """取图标预览图内绘制的标注文字（位图无可提取文字，返回空串）。"""
    if not blob:
        return ''
    if ext == 'emf':
        return ''.join(_emf_texts(blob))
    if ext == 'wmf':
        return ''.join(_wmf_texts(blob))
    return ''

def _image_ratio(blob, ext):
    """预览图自身的宽高比，用于判定对象是否等比、未被拉伸变形；取不到返回 None。"""
    try:
        if ext == 'emf' and len(blob) >= 40:
            left, top, right, bottom = struct.unpack_from('<4i', blob, 24)  # rclFrame
            if right > left and bottom > top:
                return (right - left) / float(bottom - top)
        if ext == 'wmf' and len(blob) >= 22:
            left, top, right, bottom = struct.unpack_from('<4h', blob, 6)
            if abs(right - left) and abs(bottom - top):
                return abs(right - left) / float(abs(bottom - top))
        if ext == 'png' and len(blob) >= 24:
            width, height = struct.unpack_from('>II', blob, 16)
            if width and height:
                return width / float(height)
        if ext in ('jpg', 'jpeg'):
            pos = 2
            while pos + 9 < len(blob):
                if blob[pos] != 0xFF:
                    pos += 1
                    continue
                marker, seg = blob[pos + 1], struct.unpack_from('>H', blob, pos + 2)[0]
                if marker in (0xC0, 0xC1, 0xC2):
                    height, width = struct.unpack_from('>HH', blob, pos + 5)
                    if width and height:
                        return width / float(height)
                pos += 2 + seg
    except Exception:
        pass
    return None


def _zip_payload(blob):
    """从字节流中切出 pptx（zip）载荷：
    - word/embeddings/*.pptx 这种原生部件本身就是 zip，直接返回；
    - oleObject*.bin 是 OLE 复合文档（D0CF11E0…），其 Ole10Native 流承载原始文件
      字节，从第一个能解析成演示文稿包的 PK\\x03\\x04 处切出即为 pptx。"""
    if not blob:
        return None
    starts = [0] + [m.start() for m in re.finditer(rb'PK\x03\x04', blob)]
    for start in starts:
        chunk = blob[start:]
        if chunk[:4] != b'PK\x03\x04':
            continue
        try:
            with zipfile.ZipFile(io.BytesIO(chunk)) as zf:
                if 'ppt/presentation.xml' in zf.namelist():
                    return chunk
        except Exception:
            continue
    return None

def _pptx_info(blob):
    """解析 pptx 载荷，返回 (能否正常打开, 幻灯片张数, 首页文字)。"""
    if not blob:
        return (False, 0, '')
    try:
        with zipfile.ZipFile(io.BytesIO(blob)) as zf:
            names = zf.namelist()
            if 'ppt/presentation.xml' not in names:
                return (False, 0, '')
            slides = [n for n in names if re.match(r'ppt/slides/slide\d+\.xml$', n)]
            first = ''
            if 'ppt/slides/slide1.xml' in names:
                xml = zf.read('ppt/slides/slide1.xml').decode('utf-8', 'ignore')
                first = _norm(''.join(re.findall(r'<a:t>(.*?)</a:t>', xml, re.S)))
            return (bool(slides), len(slides), first)
    except Exception:
        return (False, 0, '')


def _payload_is_expected_bytes(blob):
    """载荷字节是否与“青梧谷灵MR共创计划.pptx”（12,275,624 字节）逐字节一致。
    OLE Package 流会在文件字节之后带少量流内填充，故允许按期望长度截取后比对。"""
    if not blob:
        return False
    if hashlib.md5(blob).hexdigest() == PPTX_MD5:
        return True
    if len(blob) >= PPTX_SIZE and hashlib.md5(blob[:PPTX_SIZE]).hexdigest() == PPTX_MD5:
        return True
    return False


def _collect_objects(root):
    """收集 root 下所有 w:object，取出判定要用的 OLE／VML 属性。"""
    out = []
    for obj in root.iter(_W + 'object'):
        ole = obj.find(_O + 'OLEObject')
        shape = None
        for node in obj.iter(_V + 'shape'):
            shape = node
            break
        image = shape.find(_V + 'imagedata') if shape is not None else None
        get = (lambda k: (ole.get(k) or '')) if ole is not None else (lambda k: '')
        out.append({
            'element': obj,
            'has_ole': ole is not None,
            'type': get('Type'),
            'prog_id': get('ProgID'),
            'draw_aspect': get('DrawAspect'),
            'shape_ref': get('ShapeID'),
            'object_id': get('ObjectID'),
            'ole_rid': (ole.get(_R + 'id') or '') if ole is not None else '',
            'shape_id': (shape.get('id') or '') if shape is not None else '',
            'style': (shape.get('style') or '') if shape is not None else '',
            'image_rid': (image.get(_R + 'id') or '') if image is not None else '',
        })
    return out

def _shape_size_pt(style):
    """v:shape 的 style 里 width／height 换算成磅；缺失返回 None。"""
    def _one(key):
        found = re.search(r'(?:^|;)\s*%s\s*:\s*([0-9.]+)\s*(pt|px|in|cm|mm)?' % key,
                          style or '', re.I)
        if not found:
            return None
        factor = {'pt': 1.0, 'px': 0.75, 'in': 72.0,
                  'cm': 28.3464567, 'mm': 2.83464567}[(found.group(2) or 'pt').lower()]
        return float(found.group(1)) * factor
    return (_one('width'), _one('height'))


def _resolve_effective_font(doc, run, want):
    """按 Word 的字体继承链解析 run 实际生效的字体（打开办公软件看到的）：
       run 直接设定 → 段落样式（含 basedOn 链）→ docDefaults；任一层拿到即返回。
       want='eastAsia' 取中文字体，want='ascii' 取西文字体（缺失时回落 w:hAnsi）。"""
    def _pick(rfonts):
        if rfonts is None:
            return None
        if want == 'eastAsia':
            return rfonts.get(qn('w:eastAsia'))
        return rfonts.get(qn('w:ascii')) or rfonts.get(qn('w:hAnsi'))

    # 1) run 级
    rPr = run._element.rPr
    if rPr is not None:
        name = _pick(rPr.find(qn('w:rFonts')))
        if name:
            return name
    # 2) 段落样式链（w:pStyle → styles.xml，沿 basedOn 上溯）
    styles_el = doc.styles.element
    style_id = None
    try:
        pPr = run._parent._p.find(qn('w:pPr'))
        if pPr is not None:
            pStyle = pPr.find(qn('w:pStyle'))
            if pStyle is not None:
                style_id = pStyle.get(qn('w:val'))
    except Exception:
        style_id = None

    def _style_by_id(sid):
        for st in styles_el.findall(qn('w:style')):
            if st.get(qn('w:styleId')) == sid:
                return st
        return None

    seen = set()
    while style_id and style_id not in seen:
        seen.add(style_id)
        style = _style_by_id(style_id)
        if style is None:
            break
        rpr = style.find(qn('w:rPr'))
        if rpr is not None:
            name = _pick(rpr.find(qn('w:rFonts')))
            if name:
                return name
        base = style.find(qn('w:basedOn'))
        style_id = base.get(qn('w:val')) if base is not None else None
    # 3) docDefaults
    dd = styles_el.find(qn('w:docDefaults'))
    rpd = dd.find(qn('w:rPrDefault')) if dd is not None else None
    rpr = rpd.find(qn('w:rPr')) if rpd is not None else None
    return _pick(rpr.find(qn('w:rFonts'))) if rpr is not None else None

def _resolve_effective_eastasia(doc, run):
    """run 中文字符实际生效的东亚字体。"""
    return _resolve_effective_font(doc, run, 'eastAsia')


def _resolve_effective_ascii(doc, run):
    """run 西文／数字实际生效的字体。"""
    return _resolve_effective_font(doc, run, 'ascii')


def _font_ok(doc, run, allowed):
    """run 的中文或西文字体命中允许集合之一即算达标（中西文混排 run 的常规口径）。"""
    names = (_resolve_effective_eastasia(doc, run), _resolve_effective_ascii(doc, run))
    return any(name in allowed for name in names if name)


def _size_pt(run):
    return run.font.size.pt if run.font.size is not None else None


def _size_ok(run, expect, tol=0.3):
    size = _size_pt(run)
    return size is not None and abs(size - expect) <= tol


def _is_bold(run):
    """run 是否加粗（w:b 存在且未显式置 0／false）。"""
    rPr = run._element.rPr
    if rPr is None:
        return False
    node = rPr.find(qn('w:b'))
    if node is None:
        return False
    return (node.get(qn('w:val')) or 'true') not in ('0', 'false', 'off', 'none')


def _run_color(run):
    rPr = run._element.rPr
    node = rPr.find(qn('w:color')) if rPr is not None else None
    return _hexval(node.get(qn('w:val'))) if node is not None else None


def _para_align(para):
    pPr = para._p.find(qn('w:pPr'))
    node = pPr.find(qn('w:jc')) if pPr is not None else None
    return node.get(qn('w:val')) if node is not None else None


def _para_space_after(para):
    pPr = para._p.find(qn('w:pPr'))
    node = pPr.find(qn('w:spacing')) if pPr is not None else None
    value = node.get(qn('w:after')) if node is not None else None
    return int(value) if value is not None and value.lstrip('-').isdigit() else None

def _cell_fill(cell):
    tcPr = cell._tc.find(qn('w:tcPr'))
    node = tcPr.find(qn('w:shd')) if tcPr is not None else None
    return _hexval(node.get(qn('w:fill'))) if node is not None else None


def _cell_margins(cell):
    """单元格四边内边距（twips）；未设定的边为 None。"""
    tcPr = cell._tc.find(qn('w:tcPr'))
    mar = tcPr.find(qn('w:tcMar')) if tcPr is not None else None
    if mar is None:
        return {}
    out = {}
    for side in ('top', 'left', 'bottom', 'right'):
        node = mar.find(qn('w:' + side))
        value = node.get(qn('w:w')) if node is not None else None
        out[side] = int(value) if value is not None and value.isdigit() else None
    return out


def _cell_borders(cell):
    """单元格四边框线的 (颜色, w:sz)；未设定的边为 (None, None)。"""
    tcPr = cell._tc.find(qn('w:tcPr'))
    borders = tcPr.find(qn('w:tcBorders')) if tcPr is not None else None
    out = {}
    for side in ('top', 'left', 'bottom', 'right'):
        node = borders.find(qn('w:' + side)) if borders is not None else None
        out[side] = ((_hexval(node.get(qn('w:color'))), node.get(qn('w:sz')))
                     if node is not None else (None, None))
    return out


def _row_height(row):
    """行高设定 (hRule, 行高 twips)；未设定时返回 (None, None)。
    hRule 为 exact 表示固定行高，会压住超出部分的内容。"""
    trPr = row._tr.find(qn('w:trPr'))
    node = trPr.find(qn('w:trHeight')) if trPr is not None else None
    if node is None:
        return (None, None)
    try:
        value = int(node.get(qn('w:val')))
    except (TypeError, ValueError):
        value = None
    return (node.get(qn('w:hRule')), value)


def _cell_has_object(cell):
    return next(cell._tc.iter(_W + 'object'), None) is not None


def _para_has_object(para):
    return next(para._p.iter(_W + 'object'), None) is not None


def _text_runs(para):
    """段落内承载可见文字的 run（对象所在 run 无文字，不参与文字格式判定）。"""
    return [r for r in para.runs if (r.text or '').strip()]


def _open_word_document(file_path):
    """打开 .docx 文档（python-docx 仅支持 OOXML/.docx）."""
    return Document(file_path)

def _locate_form_table(doc):
    """定位申报表：取行数最多的 2 列表格；再定位左侧标签为“项目介绍材料”的那一行
    （找不到该标签时退回最后一行，以便位置类评分项照常判负）。"""
    table = None
    for candidate in doc.tables:
        if len(candidate.columns) != 2:
            continue
        if table is None or len(candidate.rows) > len(table.rows):
            table = candidate
    if table is None:
        return (None, -1)
    for index, row in enumerate(table.rows):
        try:
            if _norm(row.cells[0].text) == TARGET_LABEL:
                return (table, index)
        except Exception:
            continue
    return (table, len(table.rows) - 1)


def _body_signature(doc):
    """正文顶层子元素的 p／tbl 序列，用于判定是否新增了页面、章节或表格。"""
    out = []
    for child in doc.element.body.iterchildren():
        tag = child.tag.split('}')[-1]
        if tag in ('p', 'tbl'):
            out.append(tag)
    return out


def _all_text(doc):
    """全文可见文字：正文段落 + 表格单元格 + 页眉 + 页脚。"""
    parts = [p.text for p in doc.paragraphs]
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                parts.append(cell.text)
    for section in doc.sections:
        for container in (section.header, section.footer):
            parts.extend(p.text for p in container.paragraphs)
    return '\n'.join(parts)


def _dim1_fail(result, reason):
    result["dim1_pass"] = False
    result["dim1_reason"] = reason
    result["dim2_items"] = []
    result["total_score"] = 0
    result["max_score"] = MAX_SCORE_TOTAL
    return result

def evaluate(dir_path):
    """按 §2.2 约定，返回结构化 dict；不 print、不改 sys.stdout、不 sys.exit。
    入参 dir_path 为“脚本所在目录的路径”，脚本自己在该目录里定位并打开被评估的
    .docx（目录内只有压缩包时先解压）。脚本自身崩溃（含目录不存在、未发现文档等）
    由本函数捕获并返回 status='error'。"""
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
    delivery = None
    try:
        delivery = _locate_delivery(dir_path)
        result["file_name"] = delivery['docx_name']
        return _evaluate_impl(delivery, result)
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
    finally:
        if delivery and delivery.get('tmp_dir'):
            shutil.rmtree(delivery['tmp_dir'], ignore_errors=True)

def _evaluate_impl(delivery, result):
    """维度一 + 维度二的评估主体；内部构造 results 列表（每项 "+N：说明" 或
    "-N：说明"，未命中写 "+0：" / "-0："），末尾按 _MAX_DELTAS_ORDER 反解成结构化
    dim2_items 填入 result。"""
    file_path = delivery['docx_path']
    results = []

    # ========== 维度1：可用与可修改性 ==========
    # 门禁两条：① 交付物里的 Word 文档能被正常打开（能解析出正文与表格结构）；
    #           ② 文档中“项目介绍材料”一行嵌入的 PPT 能正常打开（嵌入部件里能切出
    #              演示文稿包且含幻灯片）。任一条不成立 → 完成度总分直接记 0。
    try:
        doc = _open_word_document(file_path)
        pkg = _Package(file_path)
        _ = (len(doc.paragraphs), len(doc.tables))
    except Exception as exc:
        return _dim1_fail(result, 'Word 文档无法正常打开：%s: %s' % (type(exc).__name__, exc))

    form_table, target_index = _locate_form_table(doc)
    target_cell = None
    if form_table is not None and target_index >= 0:
        target_cell = form_table.rows[target_index].cells[1]

    body_objects = _collect_objects(doc.element.body)
    hf_objects = []
    for section in doc.sections:
        for container in (section.header, section.footer):
            hf_objects.extend(_collect_objects(container._element))
    cell_objects = _collect_objects(target_cell._tc) if target_cell is not None else []
    # 判定用的主对象：优先取目标单元格内的，其次取文档内任意一个（位置类评分项另判）
    obj = cell_objects[0] if cell_objects else (
        body_objects[0] if body_objects else (hf_objects[0] if hf_objects else None))

    embed_part = pkg.part_of(obj['ole_rid']) if obj else None
    embed_blob = pkg.blob(embed_part) if embed_part else None
    payload = _zip_payload(embed_blob)
    pptx_ok, slide_count, slide1_text = _pptx_info(payload)
    image_part = pkg.part_of(obj['image_rid']) if obj else None
    image_ext = image_part.rsplit('.', 1)[-1].lower() if image_part and '.' in image_part else ''
    image_blob = pkg.blob(image_part) if image_part else None
    caption = _norm(_icon_caption(image_blob, image_ext))

    if not (payload is not None and pptx_ok):
        return _dim1_fail(result, '“项目介绍材料”一行嵌入的 PPT 无法正常打开：'
                                  '未能从文档包内取出可解析的演示文稿载荷')
    result["dim1_reason"] = ('Word 文档可正常打开；嵌入的 PPT 可正常打开'
                             '（%d 张幻灯片）' % slide_count)

    # ========== 维度2：完成度评分 ==========

    # +5：申报表“项目介绍材料”一行右侧“申报内容”单元格内存在真正的嵌入式对象
    #      （w:object 内含 o:OLEObject，Type 为“Embed”而非“Link”），PPT 字节完整写入
    #      文档包内（word/embeddings/ 下的 .bin 或 .pptx 部件），文档被单独复制到其它
    #      目录或其它电脑后双击仍能打开 PPT，不依赖同目录存在同名 .pptx
    # 严格对齐细则，仅踩细则明确写出的点：
    #   ① 目标单元格内存在 w:object，且其中含 o:OLEObject
    #   ② o:OLEObject 的 Type == "Embed"
    #   ③ 其 r:id 指向 word/embeddings/ 下真实存在的部件
    #   ④ 该部件里能切出可解析的演示文稿载荷（= PPT 字节确实写进了包内，脱离目录
    #        也能打开；反之只写关系不写字节、或指向包外文件都取不到载荷）
    embed_ok = bool(
        cell_objects
        and cell_objects[0]['has_ole']
        and cell_objects[0]['type'] == 'Embed'
        and embed_part and embed_part.startswith('word/embeddings/')
        and payload is not None and pptx_ok
    )
    results.append(("+5：" if embed_ok else "+0：")
                   + "“项目介绍材料”单元格内为真嵌入对象（Type=Embed），PPT 字节写入 word/embeddings/")

    # +5：嵌入对象的载荷是一个完整可用的 .pptx 演示文稿（约 12.28MB／12,275,624 字节，
    #      18 张幻灯片，首页含“青梧谷灵 / MR共创计划 / 以滨星手作、互动叙事与数字展演，
    #      重构区域文化体验”），提取后可直接打开且页数内容齐全，不是空壳文件、缩水重建版
    #      或仅含目录页的替代品
    # 严格对齐细则，仅踩细则明确写出的点：
    #   ① 载荷能作为演示文稿包打开
    #   ② 张数 == 18（页数齐全，排除仅含目录页／缩水重建）
    #   ③ 首页文字包含细则列出的三处字样
    #   ④ 体量落在 12,275,624 字节的 ±5%（排除空壳文件）
    payload_ok = bool(
        pptx_ok
        and slide_count == PPTX_SLIDES
        and all(key in slide1_text for key in PPTX_SLIDE1_KEYS)
        and payload is not None
        and abs(len(payload) - PPTX_SIZE) <= PPTX_SIZE * 0.05
    )
    results.append(("+5：" if payload_ok else "+0：")
                   + "载荷是完整可用的演示文稿（约 12,275,624 字节、18 张幻灯片、首页三处字样齐全）")

    # +5：在文档中双击该对象即可自动打开这份 PPT——ProgID 为 Package（OLE Package）或
    #      PowerPoint.Show.12，ShapeID、ObjectID、r:id 与关系表一一对应，document.xml.rels
    #      中 oleObject 与 image 两条关系齐备、[Content_Types].xml 已声明嵌入部件与预览图
    #      扩展名（如 bin／emf），在 Word/WPS 中激活时不报“对象不可用”“无法启动源应用程序”
    # 严格对齐细则，仅踩细则明确写出的点：
    #   ① ProgID ∈ {Package, PowerPoint.Show.12}
    #   ② ShapeID 与 v:shape/@id 一致；ObjectID、r:id 均非空
    #   ③ 两条 r:id 在 document.xml.rels 中分别是 oleObject 与 image 类型，且目标部件真实存在
    #   ④ [Content_Types].xml 声明了嵌入部件（Override 或其扩展名 Default）与预览图扩展名
    #   ⑤ 激活时不报错的可验证前提：载荷能被解析为演示文稿（③④成立且载荷可用即不会
    #        出现“对象不可用／无法启动源应用程序”）
    ole_rel = pkg.rels.get(obj['ole_rid'], {}) if obj else {}
    img_rel = pkg.rels.get(obj['image_rid'], {}) if obj else {}
    embed_ext = embed_part.rsplit('.', 1)[-1].lower() if embed_part and '.' in embed_part else ''
    activate_ok = bool(
        obj
        and obj['prog_id'] in ('Package', 'PowerPoint.Show.12')
        and obj['shape_id'] and obj['shape_id'] == obj['shape_ref']
        and obj['object_id'] and obj['ole_rid'] and obj['image_rid']
        and ole_rel.get('type') == 'oleObject' and ole_rel.get('mode') == 'Internal'
        and img_rel.get('type') == 'image' and img_rel.get('mode') == 'Internal'
        and embed_part and image_part
        and (pkg.declares_part(embed_part) or pkg.declares_extension(embed_ext))
        and pkg.declares_extension(image_ext)
        and payload is not None and pptx_ok
    )
    results.append(("+5：" if activate_ok else "+0：")
                   + "双击即开：ProgID／ShapeID／ObjectID／r:id 与关系表一一对应，oleObject 与 image 关系及内容类型声明齐备")

    # +3：嵌入对象以图标方式显示（DrawAspect="Icon"），采用 PowerPoint 演示文稿的标准橙色
    #      矢量图标，图标下方标注该演示文稿的中文文件名“青梧谷灵MR共创计划.pptx”，与在 Word 中
    #      “插入对象—显示为图标”的结果一致
    # 严格对齐细则，仅踩细则明确写出的点：
    #   ① DrawAspect == "Icon"
    #   ② 预览图是矢量图元（emf／wmf，即 Word“显示为图标”落地的图元格式），而非位图
    #   ③ 图元内绘制的标注文字含中文文件名“青梧谷灵MR共创计划.pptx”
    #        （标注在 EMF 里按行拆成多条 ExtTextOut 记录，故按记录抽取后拼接再比对）
    icon_ok = bool(
        obj and obj['draw_aspect'] == 'Icon'
        and image_ext in ('emf', 'wmf')
        and PPTX_NAME in caption
    )
    results.append(("+3：" if icon_ok else "+0：")
                   + "以图标方式显示（DrawAspect=Icon）的标准矢量图标，下方标注“青梧谷灵MR共创计划.pptx”")

    # +3：对象插入位置准确——落在申报表最后一行、左侧标签为“项目介绍材料”的那一行的右侧
    #      “申报内容”单元格内，而不是文末正文、页眉页脚、新增行或新增页
    # 严格对齐细则，仅踩细则明确写出的点：
    #   ① 对象在“项目介绍材料”行右侧单元格内
    #   ② 该行是申报表最后一行，且表仍是 25 行（没有为放对象新增行）
    #   ③ 正文其它位置、页眉、页脚都没有另外的 w:object（排除放在文末正文／页眉页脚）
    position_ok = bool(
        cell_objects
        and form_table is not None
        and len(form_table.rows) == 25
        and target_index == len(form_table.rows) - 1
        and _norm(form_table.rows[target_index].cells[0].text) == TARGET_LABEL
        and len(body_objects) == len(cell_objects)
        and not hf_objects
    )
    results.append(("+3：" if position_ok else "+0：")
                   + "对象落在申报表最后一行“项目介绍材料”右侧“申报内容”单元格内，正文其它位置与页眉页脚均无对象")

    # +3：对象尺寸与单元格匹配：宽 2.5-4cm、高 2-3cm（约 72-115pt×66-85pt），等比不拉伸变形，
    #      完整显示在 12.2cm 宽的申报内容列内，不越出单元格、不压住 C9D6DF 色框线
    # 严格对齐细则，仅踩细则明确写出的点：
    #   ① v:shape 的 width 落在 72-115pt
    #   ② v:shape 的 height 落在 66-85pt
    #   ③ 等比不拉伸：shape 宽高比与预览图自身宽高比一致（容差 5%）
    #   ④ 不越出单元格、不压住框线：宽度 ≤ 申报内容列净宽（6917 twips 去掉左右各 120
    #        twips 内边距 ≈ 333.85pt），①满足即必然成立，这里一并核算
    shape_w, shape_h = _shape_size_pt(obj['style']) if obj else (None, None)
    image_ratio = _image_ratio(image_blob, image_ext) if image_blob else None
    column_pt = (FORM_GRID[1] - 2 * CELL_MARGIN) / 20.0
    size_ok = bool(shape_w and shape_h and 72 <= shape_w <= 115 and 66 <= shape_h <= 85
                   and shape_w <= column_pt)
    if size_ok and image_ratio:
        size_ok = abs(shape_w / shape_h - image_ratio) <= max(0.05 * image_ratio, 0.03)
    results.append(("+3：" if size_ok else "+0：")
                   + "对象尺寸 72-115pt × 66-85pt、等比不变形且完整落在 12.2cm 宽的申报内容列内")

    # +3：对象在单元格内的排版与全表协调：水平居中，承载段落沿用申报内容列的
    #      Noto Serif CJK SC 9 磅非加粗字符格式，该单元格保留 3 个及以上空段落形成的留白，
    #      行高与其它申报事项行一致不塌陷
    # 严格对齐细则，仅踩细则明确写出的点：
    #   ① 承载对象的段落 jc == center
    #   ② 承载对象的 run 字体为 Noto Serif CJK SC、字号 9 磅（w:sz=18）、非加粗
    #   ③ 该单元格内空段落（无文字、无对象）数量 ≥ 3
    #   ④ 行高不塌陷：该行未被设成固定行高（hRule=exact）压住图标；若设了最小行高，
    #        其值须不小于图标高度，图标才不会被压扁裁切
    obj_para = None
    empty_paras = 0
    if target_cell is not None:
        for para in target_cell.paragraphs:
            if _para_has_object(para):
                obj_para = obj_para or para
            elif not (para.text or '').strip():
                empty_paras += 1
    obj_runs = ([r for r in obj_para.runs
                 if next(r._element.iter(_W + 'object'), None) is not None]
                if obj_para is not None else [])
    height_ok = False
    if form_table is not None:
        rule, value = _row_height(form_table.rows[target_index])
        height_ok = (rule != 'exact'
                     and (value is None or shape_h is None or value >= shape_h * 20))
    layout_ok = bool(
        obj_para is not None
        and _para_align(obj_para) == 'center'
        and obj_runs
        and all(_font_ok(doc, r, (SERIF,)) and _size_ok(r, 9.0) and not _is_bold(r)
                for r in obj_runs)
        and empty_paras >= 3
        and height_ok
    )
    results.append(("+3：" if layout_ok else "+0：")
                   + "单元格内水平居中、承载段落 Noto Serif CJK SC 9 磅非加粗、保留 3 个及以上空段落、行高不塌陷")

    # +3：插入后申报表末行不因对象撑高而跨页断裂，全文页数不增加空白页
    # 严格对齐细则，仅踩细则明确写出的点：
    #   ① 全文页数仍为 4 页（未多出空白页）；页数取包内 docProps/app.xml 的
    #        <Pages> 声明值
    #   ② 末行不跨页断裂的可验证前提：文档内 w:br type="page" 仍只有封面后那 1 个
    #        （没有为避让／撑高而额外插入分页），且末行允许跨页断行的 cantSplit 未被关掉
    pages = pkg.declared_pages()
    page_breaks = len(re.findall(r'<w:br\b[^>]*w:type="page"', pkg.document_xml))
    paging_ok = bool(
        page_breaks == PAGE_BREAKS
        and pages == EXPECT_PAGES
    )
    results.append(("+3：" if paging_ok else "+0：")
                   + "申报表末行未跨页断裂，全文仍为 4 页、未增加空白页")

    # -3：以链接对象（o:OLEObject Type="Link"）或 INCLUDEPICTURE／域链接方式插入，
    #      PPT 未真正写入文档包，脱离特定路径即失效
    # 严格对齐细则，命中任一情形即扣：
    #   ① 文档内（正文或页眉页脚）存在 Type == "Link" 的 o:OLEObject
    #   ② 域代码里出现 INCLUDEPICTURE／LINK／DDE 这类链接域
    #   ③ 存在对象但包内取不到 PPT 字节（只有关系没有部件，等同于未真正写入）
    instr_text = ' '.join(re.findall(r'<w:instrText[^>]*>(.*?)</w:instrText>',
                                     pkg.document_xml, re.S)).upper()
    link_hit = bool(
        any(o['type'] == 'Link' for o in body_objects + hf_objects)
        or re.search(r'INCLUDEPICTURE|DDEAUTO|\bDDE\b|\bLINK\b', instr_text)
        or (obj is not None and (embed_part is None or payload is None))
    )
    results.append(("-3：" if link_hit else "-0：")
                   + "以链接对象（Type=Link）或 INCLUDEPICTURE／域链接方式插入，PPT 未真正写入文档包")

    # -3：交付物为 .zip／.rar 等压缩包，需先解压才能取得可用文档，未直接交付 .docx 文件
    # 严格对齐细则：目录内没有可直接取用的 .docx，文档被装在压缩包里即扣。
    archive_hit = bool(delivery.get('archive_name'))
    results.append(("-3：" if archive_hit else "-0：")
                   + "交付物为压缩包（需先解压才能取得 .docx）")

    # -1：Word 文档文件名被改写（如追加“_已插入PPT”“_含PPT附件”等后缀），
    #      未保持“青梧谷灵MR共创计划_项目申报书.docx”
    # 严格对齐细则：文件名逐字符比对（含扩展名），不一致即扣。
    docx_name_hit = (delivery['docx_name'] != DOCX_NAME)
    results.append(("-1：" if docx_name_hit else "-0：")
                   + "Word 文档文件名不是“青梧谷灵MR共创计划_项目申报书.docx”")

    # -3：随附或嵌入的演示文稿被改名（如改成“××MR共创计划_项目介绍材料.pptx”）或字节内容
    #      被改动，与“青梧谷灵MR共创计划.pptx”（12,275,624 字节）不一致，导致图标标注、
    #      文档引用与实际交付的 PPT 对不上
    # 严格对齐细则，命中任一情形即扣：
    #   ① 一同交付的 .pptx 文件名不是“青梧谷灵MR共创计划.pptx”，或其字节 md5 不一致
    #   ② 嵌入载荷的字节与 12,275,624 字节的原始 pptx 不一致
    #   ③ 图标标注里写的是别的 .pptx 文件名
    pptx_changed = False
    for name, blob in delivery['sibling_pptx']:
        if _norm(name) != PPTX_NAME or hashlib.md5(blob).hexdigest() != PPTX_MD5:
            pptx_changed = True
    if payload is not None and not _payload_is_expected_bytes(payload):
        pptx_changed = True
    if caption and PPTX_NAME not in caption and re.search(r'\.pptx', caption, re.I):
        pptx_changed = True
    results.append(("-3：" if pptx_changed else "-0：")
                   + "交付／嵌入的演示文稿被改名或字节被改动，与“青梧谷灵MR共创计划.pptx”（12,275,624 字节）不一致")

    # -5：封面第二行主标题、第二页表格上方标题、申报表“项目名称”一行的申报内容、页眉右侧
    #      共四处的项目名不是“青梧谷灵 MR 共创计划”（页眉完整内容应为“青梧谷灵 MR 共创计划 ·
    #      项目申报书”，Noto Sans CJK SC 9 磅、下方 BFBFBF 单细线）
    # 严格对齐细则，四处任一不符即扣：
    #   ① 封面第二行主标题 == “青梧谷灵MR共创计划”
    #   ② 申报表上方标题 == “青梧谷灵MR共创计划项目申报表”（取表格前最后一个非空段落）
    #   ③ 申报表“项目名称”行右侧申报内容 == “青梧谷灵MR共创计划”
    #   ④ 页眉内容 == “青梧谷灵MR共创计划·项目申报书”，且为 Noto Sans CJK SC 9 磅、
    #        段落下框线 single BFBFBF
    paras = doc.paragraphs
    body_children = [c for c in doc.element.body.iterchildren()
                     if c.tag.split('}')[-1] in ('p', 'tbl')]
    form_title = ''
    if form_table is not None and form_table._tbl in body_children:
        for child in reversed(body_children[:body_children.index(form_table._tbl)]):
            if child.tag.split('}')[-1] != 'p':
                continue
            text = _norm(Paragraph(child, doc).text)
            if text:
                form_title = text
                break
    header = doc.sections[0].header
    header_text = _norm(''.join(p.text for p in header.paragraphs))
    header_runs = [r for p in header.paragraphs for r in _text_runs(p)]
    header_border = None
    for para in header.paragraphs:
        pPr = para._p.find(qn('w:pPr'))
        pBdr = pPr.find(qn('w:pBdr')) if pPr is not None else None
        bottom = pBdr.find(qn('w:bottom')) if pBdr is not None else None
        if bottom is not None:
            header_border = (bottom.get(qn('w:val')), _hexval(bottom.get(qn('w:color'))))
    name_cell = _norm(form_table.rows[1].cells[1].text) if form_table is not None \
        and len(form_table.rows) > 1 else ''
    name_hit = not (
        len(paras) > 1 and _norm(paras[1].text) == PROJECT_NAME
        and form_title == FORM_TITLE
        and name_cell == PROJECT_NAME
        and header_text == HEADER_TEXT
        and header_runs
        and all(_font_ok(doc, r, (SANS,)) and _size_ok(r, 9.0) for r in header_runs)
        and header_border == ('single', 'BFBFBF')
    )
    results.append(("-5：" if name_hit else "-0：")
                   + "封面主标题／申报表上方标题／“项目名称”申报内容／页眉四处项目名不是“青梧谷灵 MR 共创计划”")

    # -3：封面第三行副标题不是“以滨星手作、互动叙事与数字展演，重构区域文化体验”
    # 严格对齐细则：封面第三行文字逐字符比对。
    subtitle_hit = not (len(paras) > 2 and _norm(paras[2].text) == COVER_SUBTITLE)
    results.append(("-3：" if subtitle_hit else "-0：")
                   + "封面第三行副标题不是“以滨星手作、互动叙事与数字展演，重构区域文化体验”")

    # -3：“项目介绍材料”单元格内新增了指令未要求的文字（如把 PPT 文件名另行写成一行正文，
    #      或“说明：可双击上方 PPT 图标打开嵌入文件……”“需与同名 PPT 放在同一文件夹”之类的
    #      使用提示与前置条件说明），或新增文字不满足申报内容列统一的 Noto Serif CJK SC
    #      9 磅非加粗（如改用宋体、字号设为 10／10.5 磅、设为加粗）
    # 严格对齐细则：两种情形同源——该单元格本应只承载嵌入对象、不含任何可见文字，字符格式
    #   问题只可能出现在新增出来的文字上，所以“该单元格出现可见文字”即为本项判据；承载对象
    #   段落自身的 Noto Serif CJK SC 9 磅非加粗由排版加分项（+3）单独核对。
    cell_text = _norm(target_cell.text) if target_cell is not None else ''
    extra_text_hit = bool(cell_text)
    results.append(("-3：" if extra_text_hit else "-0：")
                   + "“项目介绍材料”单元格内新增了指令未要求的文字（文件名正文行／使用提示／"
                     "前置条件说明），或新增文字不满足 Noto Serif CJK SC 9 磅非加粗")

    # -3：除嵌入对象外，额外插入了指向外部同目录 .pptx 的超链接（TargetMode="External"），
    #      文档单独流转或改名后该链接点击报错
    # 严格对齐细则：关系表中存在 TargetMode="External" 且目标为 .pptx 的关系即扣。
    ext_link_hit = any(rel['mode'] == 'External' and rel['target'].lower().endswith('.pptx')
                       for rel in pkg.rels.values())
    results.append(("-3：" if ext_link_hit else "-0：")
                   + "额外插入了指向外部 .pptx 的超链接（TargetMode=External）")

    # -3：对象图标不是 PowerPoint 演示文稿的标准矢量图标——自行绘制示意图形（如自画的
    #      显示器＋箭头方框）、图标内文字互相重叠或被边界裁切、标注使用英文（如“Project Intro”）、
    #      改用 PNG／JPG 位图致打印放大后发虚锯齿，任一情形即扣
    # 严格对齐细则，只在有“确凿缺陷证据”时才扣（DrawAspect 不是 Icon 而显示幻灯片内容的
    #   情形只失掉图标加分项，不在此重复扣分）：
    #   ① 预览图是 PNG／JPG／GIF／BMP 位图 → 打印放大发虚锯齿
    #   ② 预览图是矢量图元，但图元内标注不含中文文件名，且出现成串英文字母（英文标注）
    #        或整段标注无中文（自画图形通常也取不到中文文件名标注）
    icon_bad = False
    if obj is not None and image_part:
        if image_ext in ('png', 'jpg', 'jpeg', 'gif', 'bmp', 'tiff'):
            icon_bad = True
        elif PPTX_NAME not in caption:
            if re.search(r'[A-Za-z]{3,}', caption) or not re.search(r'[一-鿿]', caption):
                icon_bad = True
    results.append(("-3：" if icon_bad else "-0：")
                   + "对象图标不是 PowerPoint 标准矢量图标（位图图标／英文或缺失中文文件名标注／自画图形）")

    # -3：申报表的行列结构被改动（不再是 25 行 2 列，出现新增行、删除行、合并单元格或
    #      列宽调整），左列宽 4.2cm、右列宽 12.2cm 的比例被破坏
    # 严格对齐细则，命中任一情形即扣：
    #   ① 行数 != 25 或列数 != 2
    #   ② tblGrid 不是 2381／6917 twips（= 4.2cm／12.2cm）
    #   ③ 表内出现 w:gridSpan 或 w:vMerge（合并单元格）
    grid = []
    merged = False
    if form_table is not None:
        tbl_grid = form_table._tbl.find(qn('w:tblGrid'))
        if tbl_grid is not None:
            for col in tbl_grid.findall(qn('w:gridCol')):
                value = col.get(qn('w:w'))
                grid.append(int(value) if value and value.isdigit() else -1)
        merged = (next(form_table._tbl.iter(qn('w:gridSpan')), None) is not None
                  or next(form_table._tbl.iter(qn('w:vMerge')), None) is not None)
    struct_hit = not (
        form_table is not None
        and len(form_table.rows) == 25 and len(form_table.columns) == 2
        and grid == FORM_GRID and not merged
    )
    results.append(("-3：" if struct_hit else "-0：")
                   + "申报表行列结构被改动（不再是 25 行 2 列／左列 2381 右列 6917 twips／出现合并单元格）")

    # -3：24 个申报事项（项目名称、项目类别、申报方向、项目定位、项目背景、痛点分析、建设目标、
    #      建设内容、产品体系、目标客群、体验流程、技术路线、技术选型、竞争优势、差异化能力、
    #      商业模式、资金使用方向、资源整合、团队配置、实施计划、风险控制、预期成果、社会价值、
    #      项目介绍材料）的名称或正文内容被删改、精简、改写、调序
    # 严格对齐细则：逐行按顺序比对左列标签，并按去空白后的字符数 + md5 比对右列正文
    #   （“项目介绍材料”行是插入位置，只比标签不比正文）。任一行不符即扣。
    rows_hit = False
    if form_table is None or len(form_table.rows) < len(_EXPECT_ROWS) + 1:
        rows_hit = True
    else:
        for offset, (label, length, digest) in enumerate(_EXPECT_ROWS, start=1):
            cells = form_table.rows[offset].cells
            if _norm(cells[0].text) != label:
                rows_hit = True
                break
            if label == TARGET_LABEL:
                continue
            body = _norm(cells[1].text)
            if len(body) != length or _md5(body) != digest:
                rows_hit = True
                break
    results.append(("-3：" if rows_hit else "-0：")
                   + "24 个申报事项的名称或正文内容被删改、精简、改写、调序")

    # -1：表头行“申报事项”“申报内容”不满足 1F4E79 深蓝底色＋白色加粗 10.5 磅
    #      Noto Sans CJK SC 居中
    # 严格对齐细则：两个表头单元格文字、底色、字体、字号、加粗、字色、居中逐项核对。
    head_hit = True
    if form_table is not None and len(form_table.rows) >= 1:
        head_hit = False
        head_cells = form_table.rows[0].cells
        if len(head_cells) != 2:
            head_hit = True
        for index, cell in enumerate(head_cells[:2]):
            if _norm(cell.text) != HEAD_CELLS[index] or _cell_fill(cell) != '1F4E79':
                head_hit = True
                break
            for para in cell.paragraphs:
                if _text_runs(para) and _para_align(para) != 'center':
                    head_hit = True
                for run in _text_runs(para):
                    if not (_font_ok(doc, run, (SANS,)) and _size_ok(run, 10.5)
                            and _is_bold(run) and _run_color(run) == 'FFFFFF'):
                        head_hit = True
                        break
    results.append(("-1：" if head_hit else "-0：")
                   + "表头行“申报事项”“申报内容”不满足 1F4E79 底色＋白色加粗 10.5 磅 Noto Sans CJK SC 居中")

    # -1：封面首行“项目申报书”不满足 Noto Sans CJK SC 28 磅 加粗 1F4E79 居中
    # 严格对齐细则：文字、字体、字号、加粗、字色、居中逐项核对。
    cover_hit = True
    if paras and _norm(paras[0].text) == COVER_TITLE:
        runs = _text_runs(paras[0])
        cover_hit = not (runs and _para_align(paras[0]) == 'center'
                         and all(_font_ok(doc, r, (SANS,)) and _size_ok(r, 28.0)
                                 and _is_bold(r) and _run_color(r) == '1F4E79'
                                 for r in runs))
    results.append(("-1：" if cover_hit else "-0：")
                   + "封面首行“项目申报书”不满足 Noto Sans CJK SC 28 磅 加粗 1F4E79 居中")

    # -1：“文化科技融合／沉浸式文旅／共创工坊”三色块一行的文字、11 磅加粗 1F4E79 字色或
    #      EAF4F2／D9EAF7／EAF4F2 底色被改动
    # 严格对齐细则：三个色块的文字、底色、字号、加粗、字色逐项核对（取 3 列的那张表）。
    chip_table = None
    for candidate in doc.tables:
        if len(candidate.columns) == 3 and len(candidate.rows) == 1:
            chip_table = candidate
            break
    chip_hit = True
    if chip_table is not None:
        chip_hit = False
        cells = chip_table.rows[0].cells
        if len(cells) != 3:
            chip_hit = True
        for index, cell in enumerate(cells[:3]):
            if _norm(cell.text) != CHIP_TEXTS[index] or _cell_fill(cell) != CHIP_FILLS[index]:
                chip_hit = True
                break
            for para in cell.paragraphs:
                for run in _text_runs(para):
                    if not (_size_ok(run, 11.0) and _is_bold(run)
                            and _run_color(run) == '1F4E79'):
                        chip_hit = True
                        break
    results.append(("-1：" if chip_hit else "-0：")
                   + "三色块一行的文字、11 磅加粗 1F4E79 字色或 EAF4F2／D9EAF7／EAF4F2 底色被改动")

    # -1：“申报单位：__________________________”与“申报日期：2026 年 5 月”两行不满足
    #      Noto Serif CJK SC 12 磅居中，或下划线长度、日期文字被改动
    # 严格对齐细则：两行文字（含 26 个下划线）、字体、字号、居中逐项核对。
    unit_hit = True
    if len(paras) > 4:
        unit_para, date_para = paras[3], paras[4]
        runs = _text_runs(unit_para) + _text_runs(date_para)
        unit_hit = not (
            _norm(unit_para.text) == UNIT_LINE
            and _norm(date_para.text) == DATE_LINE
            and _para_align(unit_para) == 'center' and _para_align(date_para) == 'center'
            and runs
            and all(_font_ok(doc, r, CJK_SERIF_FONTS) and _size_ok(r, 12.0) for r in runs)
        )
    results.append(("-1：" if unit_hit else "-0：")
                   + "“申报单位：”（26 个下划线）与“申报日期：2026 年 5 月”两行不满足 Noto Serif CJK SC 12 磅居中")

    # -1：左列各申报事项标签不满足 Noto Serif CJK SC 9.5 磅 加粗 1F4E79 居中、EAF4F2 底色
    #      （末行“项目介绍材料”标签沿用 D9EAF7 底色）
    # 严格对齐细则：24 个标签单元格的底色、字体、字号、加粗、字色、居中逐项核对。
    label_hit = True
    if form_table is not None and len(form_table.rows) >= 25:
        label_hit = False
        for index in range(1, 25):
            cell = form_table.rows[index].cells[0]
            if _cell_fill(cell) != LABEL_FILLS[index]:
                label_hit = True
                break
            for para in cell.paragraphs:
                if _text_runs(para) and _para_align(para) != 'center':
                    label_hit = True
                for run in _text_runs(para):
                    if not (_font_ok(doc, run, CJK_SERIF_FONTS) and _size_ok(run, 9.5)
                            and _is_bold(run) and _run_color(run) == '1F4E79'):
                        label_hit = True
                        break
            if label_hit:
                break
    results.append(("-1：" if label_hit else "-0：")
                   + "左列申报事项标签不满足 Noto Serif CJK SC 9.5 磅 加粗 1F4E79 居中、EAF4F2／F7FBFA 隔行底色（末行 D9EAF7）")

    # -1：右列申报内容正文不满足 Noto Serif CJK SC 9 磅、FFFFFF／FBFBFB 隔行底色、段后间距 0
    # 严格对齐细则：24 个申报内容单元格的底色、承载文字段落的段后间距，以及每个文字 run 的
    #   字号与字体逐项核对（空段落不承载文字，不纳入段后间距与字体判定）。
    content_hit = True
    if form_table is not None and len(form_table.rows) >= 25:
        content_hit = False
        for index in range(1, 25):
            cell = form_table.rows[index].cells[1]
            if _cell_fill(cell) != BODY_FILLS[index]:
                content_hit = True
                break
            for para in cell.paragraphs:
                runs = _text_runs(para)
                if not runs:
                    continue
                if _para_space_after(para) != 0:
                    content_hit = True
                    break
                for run in runs:
                    if not (_font_ok(doc, run, CJK_SERIF_FONTS) and _size_ok(run, 9.0)):
                        content_hit = True
                        break
                if content_hit:
                    break
            if content_hit:
                break
    results.append(("-1：" if content_hit else "-0：")
                   + "右列申报内容正文不满足 Noto Serif CJK SC 9 磅、FFFFFF／FBFBFB 隔行底色（末行 FFFFFF）、段后间距 0")

    # -1：表格框线颜色被改动（内部框线 C9D6DF 0.5 磅、末行上框线 1F4E79 1 磅），
    #      或单元格内边距不再是 0.21cm
    # 严格对齐细则：申报表每个单元格四边框线的颜色与磅宽、四边内边距 120 twips 逐项核对；
    #   末行两个单元格的上框线应为 1F4E79 sz=12，其余边与其余行为 C9D6DF sz=6。
    border_hit = True
    if form_table is not None and len(form_table.rows) == 25:
        border_hit = False
        for index, row in enumerate(form_table.rows):
            for cell in row.cells:
                margins = _cell_margins(cell)
                if any(margins.get(side) != CELL_MARGIN
                       for side in ('top', 'left', 'bottom', 'right')):
                    border_hit = True
                    break
                borders = _cell_borders(cell)
                expect_top = LAST_TOP_BORDER if index == 24 else INNER_BORDER
                if borders.get('top') != expect_top:
                    border_hit = True
                    break
                if any(borders.get(side) != INNER_BORDER
                       for side in ('left', 'bottom', 'right')):
                    border_hit = True
                    break
            if border_hit:
                break
    results.append(("-1：" if border_hit else "-0：")
                   + "表格框线颜色（内部 C9D6DF 0.5 磅、末行上框线 1F4E79 1 磅）或单元格内边距 0.21cm 被改动")

    # -1：页面设置被改动（不再是 A4 纵向 21cm×29.7cm，上边距 2.2cm、下边距 2.0cm、
    #      左右边距 2.3cm）
    # 严格对齐细则：sectPr 的 pgSz（11906×16838 twips、纵向）与 pgMar 四边逐项核对。
    page_hit = True
    try:
        sect_pr = doc.sections[0]._sectPr
        pg_sz = sect_pr.find(_W + 'pgSz')
        pg_mar = sect_pr.find(_W + 'pgMar')
        page_hit = not (
            pg_sz is not None and pg_mar is not None
            and (int(pg_sz.get(_W + 'w')), int(pg_sz.get(_W + 'h'))) == PAGE_SIZE
            and (pg_sz.get(_W + 'orient') or 'portrait') == 'portrait'
            and all(int(pg_mar.get(_W + side) or -1) == value
                    for side, value in PAGE_MARGIN.items())
        )
    except Exception:
        page_hit = True
    results.append(("-1：" if page_hit else "-0：")
                   + "页面设置被改动（A4 纵向 21cm×29.7cm，上 2.2cm／下 2.0cm／左右 2.3cm 边距）")

    # -1：页脚居中的“第 X 页”页码域被删除、改成静态数字或格式
    #      （Noto Serif CJK SC 9 磅 666666）被改动
    # 严格对齐细则：页脚里必须仍是 PAGE 域（fldChar+instrText 或 fldSimple），
    #   文字为“第…页”，承载段落居中，且各 run 为 Noto Serif CJK SC 9 磅 666666。
    footer = doc.sections[0].footer
    footer_paras = list(footer.paragraphs)
    footer_text = _norm(''.join(para.text for para in footer_paras))
    footer_runs = [run for para in footer_paras for run in _text_runs(para)]
    page_field = False
    for para in footer_paras:
        instr = ''.join(node.text or '' for node in para._p.iter(_W + 'instrText'))
        instr += ' '.join(node.get(_W + 'instr') or ''
                          for node in para._p.iter(_W + 'fldSimple'))
        if 'PAGE' in instr.upper():
            page_field = True
            break
    footer_hit = not (
        page_field
        and '第' in footer_text and '页' in footer_text
        and footer_runs
        and all(_para_align(para) == 'center'
                for para in footer_paras if _text_runs(para))
        and all(_font_ok(doc, run, CJK_SERIF_FONTS) and _size_ok(run, 9.0)
                and _run_color(run) == '666666' for run in footer_runs)
    )
    results.append(("-1：" if footer_hit else "-0：")
                   + "页脚居中“第 X 页”页码域被删除／改成静态数字，或 Noto Serif CJK SC 9 磅 666666 格式被改动")

    # -1：新增了指令未要求的页面、附录页、封底、目录或章节标题
    # 严格对齐细则：正文顶层结构应仍是“3 段 + 三色块表 + 4 段 + 申报表 + 1 段”，
    #   分页符仍为 1 处，全文页数不超过 4 页（新增页／附录／封底／目录／章节标题都会
    #   体现为顶层多出段落或表格、多出分页符、页数增加）。
    pages = pkg.declared_pages()
    extra_page_hit = not (
        _body_signature(doc) == BODY_SIGNATURE
        and len(re.findall(r'<w:br\b[^>]*w:type="page"', pkg.document_xml)) == PAGE_BREAKS
        and (pages is None or pages <= EXPECT_PAGES)
    )
    results.append(("-1：" if extra_page_hit else "-0：")
                   + "新增了指令未要求的页面、附录页、封底、目录或章节标题")

    # -1：文档中留下临时占位内容（如“待补充”“此处插入PPT”“见附件”等未清理的提示文字）
    full_text = _all_text(doc)
    placeholder_hit = any(word in full_text for word in PLACEHOLDERS)
    results.append(("-1：" if placeholder_hit else "-0：")
                   + "文档中留下临时占位内容（“待补充”“此处插入PPT”“见附件”等未清理的提示文字）")

    # ---------- 反解成结构化 dim2_items ----------
    # results 内每项形如 "+5：说明" / "+0：说明" / "-1：说明" / "-0：说明"，
    # 与 _MAX_DELTAS_ORDER 逐一对应：前缀非 0 即命中该项（加分项得满分，减分项扣满分）。
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
