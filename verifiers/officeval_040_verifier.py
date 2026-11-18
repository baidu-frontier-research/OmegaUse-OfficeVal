#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# pyright: reportUnknownMemberType=false, reportUnknownVariableType=false, reportUnknownArgumentType=false, reportUnknownParameterType=false, reportMissingParameterType=false, reportPrivateUsage=false, reportUnknownLambdaType=false, reportGeneralTypeIssues=false

import os
import re
import sys
import json
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import parse_xml
from docx.oxml.ns import qn
try:
    from docx.image.image import Image as _DocxImage  # 读取所嵌位图的像素尺寸/DPI
except ImportError:
    _DocxImage = None

# 说明：按接口统一约定，本模块不修改全局 sys.stdout，也不 print 主结果。
# 主结果统一由 evaluate(dir_path) -> dict 返回；入参为“脚本所在目录的路径”，
# 由脚本自己在该目录内定位并打开被评估的 .docx。本文件仅在 __main__ 下作为
# 本地调试入口，才把 dict 序列化为 JSON 打印。

SCRIPT_ID = "040"
# 维度二“加分项”max_delta 之和（= 满分）；维度一未通过 / 脚本异常时兜底
MAX_SCORE_TOTAL = 72

# 维度二所有评分项按“结果追加顺序”的 max_delta 列表；末端把 results 里的
# "+N：..." / "-N：..." 字符串反解为结构化 dim2_items 时用。顺序须与
# _evaluate_impl 中 results.append(...) 的顺序严格一致；共 36 项：前 22 项为
# 加分项（累加 = MAX_SCORE_TOTAL），后 14 项为减分项（命中即扣，不计入满分）。
_MAX_DELTAS_ORDER = [
    3,   # 01 首行标题中文字体隶书
    3,   # 02 首行标题二号/加粗/蓝色
    5,   # 03 首行标题发光（橄榄色、7磅）
    3,   # 04 首行标题阴影（向上偏移）
    1,   # 05 首行标题居中对齐
    3,   # 06 首行标题段后20磅
    3,   # 07 正文楷体小四
    3,   # 08 正文首行缩进2字符
    1,   # 09 正文段后6磅
    3,   # 10 正文1.1倍行距
    3,   # 11 “专业设置如下：”拆分为独立段落
    5,   # 12 首字下沉—悬挂
    5,   # 13 下沉行数2行
    3,   # 14 下沉距正文0.3厘米
    3,   # 15 全文4处改为“智能软件工程系”
    3,   # 16 正文3处“智能软件工程系”宋体
    3,   # 17 正文3处四号/加粗倾斜/绿色
    3,   # 18 正文3处着重号
    5,   # 19 图片作水印衬于文字下方
    5,   # 20 水印冲蚀效果
    5,   # 21 水印缩放100%铺满整页
    1,   # 22 水印水平垂直居中且仅1张
    -3,  # 23 正文文字颜色被改（非深灰333333）
    -3,  # 24 正文出现未要求的加粗/倾斜/下划线/着重号等
    -3,  # 25 正文段前间距非0磅
    -1,  # 26 正文段落非两端对齐
    -1,  # 27 纸张非A4纵向或页边距被改
    -3,  # 28 5行专业名称未独立成行/顺序被调整
    -3,  # 29 正文数据被改写
    -1,  # 30 机构名称/引号内固定表述被改写
    -1,  # 31 新增未要求的内容（小标题/空段/分页符等）
    -1,  # 32 页眉页脚出现除水印外的文字/页码/横线
    -3,  # 33 “专业设置如下：”整行被置入下沉框
    -3,  # 34 正文被图片/艺术字/文本框/表格承载
    -3,  # 35 水印被拉伸变形或被替换
    -1,  # 36 正文文字被加发光/阴影等文本效果
]
assert sum(d for d in _MAX_DELTAS_ORDER if d > 0) == MAX_SCORE_TOTAL

# ===== 命名空间 =====
W14_NS = 'http://schemas.microsoft.com/office/word/2010/wordml'
A_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main'
WP_NS = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing'
V_NS = 'urn:schemas-microsoft-com:vml'
O_NS = 'urn:schemas-microsoft-com:office:office'

# ===== 文本锚点常量（细则中逐条点名、不得改动的文字） =====
TITLE_TEXT = '智能软件工程系简介'
OLD_TERM = '智能工程系'
NEW_TERM = '智能软件工程系'
SPLIT_LINE = '专业设置如下：'
PREV_TAIL = '科学设置专业。'
DROP_CHAR = '专'
LAST_ANCHOR = '面向岗位，面向技能，面向未来，'
MAJORS = (
    '计算机应用技术', '软件开发技术', '道路安全与智能运维',
    '城市轨道交通运维', '数字媒体艺术设计',
)
DATA_KEYS = (
    '专、兼职教师48名', '专职教师24名', '正高级职称2名', '副高级职称6名',
    '讲师15名', '建筑面积5200平方米', '10个专业实验实训室', '4个普通机房',
    '校外实训基地9个',
)
NAME_KEYS = (
    '南澜智信信息技术有限公司', '南澜云行科技发展有限公司', '系办智联技术服务中心',
    '“项目实训”', '“双证融合”', '“岗位实训”', '“企业+平台”',
    '“育人为先、重德精技、务实创新、追求卓越”',
)
# 交付文档应有的全文内容（{T} 为“智能工程系”/“智能软件工程系”占位）；
# 仅用于判定“是否新增了指令未要求的文字”，按拼接后的长串做包含性比对。
EXPECTED_PARTS = (
    '{T}简介',
    '{T}是学院设立较早的系部之一，本系师资队伍结构合理，教学、科研基础较扎实。'
    '现有专、兼职教师48名。专职教师24名，有正高级职称2名，副高级职称6名，讲师15名。'
    '校内实训中心建筑面积5200平方米，共有10个专业实验实训室和4个普通机房，'
    '一个综合协作工作室和一个技术服务中心；另有校外实训基地9个。',
    '{T}坚持以岗位能力为导向，结合社会、区域和行业对技术人才的需求科学设置专业。',
    '专业设置如下：',
    '计算机应用技术', '软件开发技术', '道路安全与智能运维',
    '城市轨道交通运维', '数字媒体艺术设计',
    '本系在持续办学过程中，持续优化和完善复合型技术人才培养模式。'
    '基本形成了以职业岗位需求为导向，以提升学生的职业技能和职业素养为重点，'
    '以学生成长为本位的人才培养理念。同时，在教学上将职业资格证书训练融入专业课程'
    '安排之中，稳步推进与“项目实训”教学模式相衔接的“双证融合”课程体系建设。'
    '该培养模式把课堂学习和岗位实践贯穿于教学全过程，整体提升学生的专业能力、'
    '实践能力和协作能力。',
    '为拓宽学生“岗位实训”和就业的渠道，持续加强校内、外实训基地建设。'
    '在校内实训基地建设上，实施三位一体“企业+平台”的建设模式。该模式通过与'
    '南澜智信信息技术有限公司合作成立的南澜云行科技发展有限公司、系办智联技术服务中心'
    '和由系实训中心参照企业运行方式而建立的训练型工作室，为学生提供生产性实习和'
    '具有职业氛围的综合实训环境。在校外实训基地的建设上，通过与相关单位建立稳定和'
    '长期的合作关系来不断扩大基地的数量和规模。',
    '面向岗位，面向技能，面向未来，{T}以培养学生的职业能力为核心，'
    '按照“育人为先、重德精技、务实创新、追求卓越”的办学理念，着力培养市场上需要的、'
    '适应性强的、具备一技之长的技术型人才，为区域产业发展培养可用之材。',
)
_EXPECTED_NEW = ''.join(EXPECTED_PARTS).replace('{T}', NEW_TERM)
_EXPECTED_OLD = ''.join(EXPECTED_PARTS).replace('{T}', OLD_TERM)

# ===== 字体/度量常量 =====
LISHU_KEYS = ('隶书', 'LiSu')
KAITI_KEYS = ('楷体', 'KaiTi', 'STKaiti')
STRICT_SONGTI = ('宋体', 'SimSun', 'NSimSun')
BODY_COLOR = '333333'          # 正文原有深灰
ZH_RE = re.compile(r'[一-鿿]')
EMU_PER_PT = 12700.0
EMU_PER_CM = 360000.0
PX_DPI = 96.0                  # 细则按 96DPI 折算位图的 100% 显示尺寸
_THEME_CACHE = {}


def _hex_rgb(val):
    """把 w:color/@w:val 之类的十六进制颜色转成 (r,g,b)；auto/空/非法返回 None。"""
    if not val:
        return None
    v = str(val).strip().lstrip('#').upper()
    if v in ('AUTO', 'NONE'):
        return None
    if len(v) != 6 or re.search(r'[^0-9A-F]', v):
        return None
    return (int(v[0:2], 16), int(v[2:4], 16), int(v[4:6], 16))


def _is_blue(val):
    """蓝色系（0000FF / 0070C0 等）。"""
    rgb = _hex_rgb(val)
    if rgb is None:
        return False
    r, g, b = rgb
    return b >= 110 and b > r + 40 and b > g + 20


def _is_green(val):
    """绿色系（00B050 / 008000 / 9BBB59 等）。"""
    rgb = _hex_rgb(val)
    if rgb is None:
        return False
    r, g, b = rgb
    return g >= 100 and g > r + 20 and g > b + 20


def _is_olive(val):
    """橄榄色系（9BBB59 / 6B8E23 / 808000 等：红绿相近且明显高于蓝）。"""
    rgb = _hex_rgb(val)
    if rgb is None:
        return False
    r, g, b = rgb
    return g >= 90 and g > b + 40 and r > b + 20 and abs(r - g) <= 100

def _load_theme_fonts(doc):
    """读取 word/theme/theme1.xml 的字体方案，返回
       {'majorAscii','majorEastAsia','minorAscii','minorEastAsia'}；
       东亚字体优先取 <a:ea>，为空时回落 <a:font script="Hans">（与 Word 一致）。"""
    fonts = {'majorAscii': None, 'majorEastAsia': None,
             'minorAscii': None, 'minorEastAsia': None}
    root = None
    try:
        for part in doc.part.package.iter_parts():
            if str(part.partname).lower().endswith('theme1.xml'):
                root = getattr(part, 'element', None)
                if root is None:
                    root = parse_xml(part.blob)
                break
    except Exception:
        root = None
    if root is None:
        return fonts
    for tag, key in (('majorFont', 'major'), ('minorFont', 'minor')):
        node = root.find('.//{%s}%s' % (A_NS, tag))
        if node is None:
            continue
        latin = node.find('{%s}latin' % A_NS)
        if latin is not None and latin.get('typeface'):
            fonts[key + 'Ascii'] = latin.get('typeface')
        ea = node.find('{%s}ea' % A_NS)
        face = ea.get('typeface') if ea is not None else None
        if not face:
            for f in node.findall('{%s}font' % A_NS):
                if f.get('script') == 'Hans' and f.get('typeface'):
                    face = f.get('typeface')
                    break
        if face:
            fonts[key + 'EastAsia'] = face
    return fonts


def _theme_fonts(doc):
    key = id(doc)
    if key not in _THEME_CACHE:
        _THEME_CACHE[key] = _load_theme_fonts(doc)
    return _THEME_CACHE[key]


def _resolve_theme(doc, theme_ref, slot):
    """把 w:eastAsiaTheme / w:asciiTheme 等主题字体引用解析为真实字体名。"""
    if not theme_ref:
        return None
    ref = str(theme_ref)
    kind = 'major' if ref.startswith('major') else 'minor'
    return _theme_fonts(doc).get(kind + ('EastAsia' if slot == 'ea' else 'Ascii'))

def _style_chain_rprs(doc, para_el):
    """返回段落样式链（w:pStyle → styles.xml，沿 basedOn 上溯）上的 w:rPr 列表。"""
    out = []
    styles_el = doc.styles.element
    style_id = None
    pPr = para_el.find(qn('w:pPr')) if para_el is not None else None
    if pPr is not None:
        pStyle = pPr.find(qn('w:pStyle'))
        if pStyle is not None:
            style_id = pStyle.get(qn('w:val'))

    def _style_by_id(sid):
        for st in styles_el.findall(qn('w:style')):
            if st.get(qn('w:styleId')) == sid:
                return st
        return None

    seen = set()
    while style_id and style_id not in seen:
        seen.add(style_id)
        st = _style_by_id(style_id)
        if st is None:
            break
        rpr = st.find(qn('w:rPr'))
        if rpr is not None:
            out.append(rpr)
        base = st.find(qn('w:basedOn'))
        style_id = base.get(qn('w:val')) if base is not None else None
    return out


def _doc_default_rpr(doc):
    """docDefaults/rPrDefault/rPr。"""
    dd = doc.styles.element.find(qn('w:docDefaults'))
    if dd is None:
        return None
    rpd = dd.find(qn('w:rPrDefault'))
    if rpd is None:
        return None
    return rpd.find(qn('w:rPr'))


def _rpr_chain(doc, run):
    """run 实际生效属性的查找链：run 级 → 段落样式链 → docDefaults。"""
    chain = []
    rPr = run._element.rPr
    if rPr is not None:
        chain.append(rPr)
    para_el = getattr(getattr(run, '_parent', None), '_p', None)
    chain.extend(_style_chain_rprs(doc, para_el))
    dflt = _doc_default_rpr(doc)
    if dflt is not None:
        chain.append(dflt)
    return chain

def _resolve_effective_eastasia(doc, run):
    """按 Word 的字体继承链，解析 run 中文字符实际生效的东亚字体（打开办公软件看到的）：
       run 直接设定(w:eastAsia / w:eastAsiaTheme) → 段落样式(含 basedOn 链) → docDefaults。
       任一层拿到即返回；都没有则返回 None。"""
    for rpr in _rpr_chain(doc, run):
        rf = rpr.find(qn('w:rFonts'))
        if rf is None:
            continue
        ea = rf.get(qn('w:eastAsia'))
        if ea:
            return ea
        themed = _resolve_theme(doc, rf.get(qn('w:eastAsiaTheme')), 'ea')
        if themed:
            return themed
    return None


def _resolve_effective_ascii(doc, run):
    """按 Word 的字体继承链，解析 run 对西文/数字实际生效的字体：
       run 直接设定(w:ascii → 回落 w:hAnsi / 主题引用) → 段落样式链 → docDefaults。"""
    for rpr in _rpr_chain(doc, run):
        rf = rpr.find(qn('w:rFonts'))
        if rf is None:
            continue
        v = rf.get(qn('w:ascii')) or rf.get(qn('w:hAnsi'))
        if v:
            return v
        themed = _resolve_theme(
            doc, rf.get(qn('w:asciiTheme')) or rf.get(qn('w:hAnsiTheme')), 'ascii')
        if themed:
            return themed
    return None


def _resolve_effective_size_pt(doc, run):
    """解析 run 实际生效的字号（磅）：w:sz 为半磅值。"""
    for rpr in _rpr_chain(doc, run):
        sz = rpr.find(qn('w:sz'))
        if sz is not None and sz.get(qn('w:val')):
            try:
                return int(sz.get(qn('w:val'))) / 2.0
            except ValueError:
                return None
    return None


def _resolve_effective_color(doc, run):
    """解析 run 实际生效的字体颜色，返回大写十六进制、'AUTO' 或 None（未设定）。"""
    for rpr in _rpr_chain(doc, run):
        col = rpr.find(qn('w:color'))
        if col is None:
            continue
        val = col.get(qn('w:val'))
        if val:
            return str(val).upper()
    return None

def _toggle_on(el):
    """开关型属性（w:b/w:i/w:em/w:u...）是否为“开”：无 val 视为开，val=0/false 为关。"""
    if el is None:
        return False
    val = el.get(qn('w:val'))
    if val is None:
        return True
    return str(val).lower() not in ('0', 'false', 'off', 'none')


def _resolve_effective_toggle(doc, run, tag):
    """解析 run 的开关型属性（如 'w:b'/'w:i'）：沿继承链取第一个显式声明。"""
    for rpr in _rpr_chain(doc, run):
        el = rpr.find(qn(tag))
        if el is not None:
            return _toggle_on(el)
    return False


def _run_rpr(run):
    return run._element.rPr


def _w14_children(run, names):
    """run 直接 rPr 下是否存在指定的 Word 2010 文本效果子元素（如 glow/shadow）。"""
    rPr = _run_rpr(run)
    if rPr is None:
        return []
    return [rPr.find('{%s}%s' % (W14_NS, n)) for n in names
            if rPr.find('{%s}%s' % (W14_NS, n)) is not None]


def _int_attr(el, name):
    """读取属性并转 int（w: 前缀走 qn，其余按原名）；缺失/非法返回 None。"""
    if el is None:
        return None
    raw = el.get(qn(name)) if name.startswith('w:') else el.get(name)
    if raw is None:
        return None
    try:
        return int(round(float(str(raw).rstrip('f'))))
    except ValueError:
        return None


def _pPr(para):
    return para._element.find(qn('w:pPr'))


def _frame_pr(para):
    pPr = _pPr(para)
    return pPr.find(qn('w:framePr')) if pPr is not None else None

def _para_prop(doc, para, tag, attr):
    """解析段落级属性（w:spacing/@w:after、w:ind/@w:firstLineChars、w:jc/@w:val 等）：
       段落直接设定 → 段落样式链 → docDefaults/pPrDefault；返回字符串或 None。"""
    pPr = _pPr(para)
    if pPr is not None:
        el = pPr.find(qn(tag))
        if el is not None and el.get(qn(attr)) is not None:
            return el.get(qn(attr))
    styles_el = doc.styles.element
    style_id = None
    if pPr is not None:
        pStyle = pPr.find(qn('w:pStyle'))
        if pStyle is not None:
            style_id = pStyle.get(qn('w:val'))
    seen = set()
    while style_id and style_id not in seen:
        seen.add(style_id)
        st = None
        for s in styles_el.findall(qn('w:style')):
            if s.get(qn('w:styleId')) == style_id:
                st = s
                break
        if st is None:
            break
        spr = st.find(qn('w:pPr'))
        if spr is not None:
            el = spr.find(qn(tag))
            if el is not None and el.get(qn(attr)) is not None:
                return el.get(qn(attr))
        base = st.find(qn('w:basedOn'))
        style_id = base.get(qn('w:val')) if base is not None else None
    dd = styles_el.find(qn('w:docDefaults'))
    if dd is not None:
        ppd = dd.find(qn('w:pPrDefault'))
        spr = ppd.find(qn('w:pPr')) if ppd is not None else None
        if spr is not None:
            el = spr.find(qn(tag))
            if el is not None and el.get(qn(attr)) is not None:
                return el.get(qn(attr))
    return None


def _para_prop_int(doc, para, tag, attr, default=None):
    raw = _para_prop(doc, para, tag, attr)
    if raw is None:
        return default
    try:
        return int(round(float(raw)))
    except ValueError:
        return default

def _run_spans(para):
    """返回 [(起, 止, 序号, run)]，与 para.text 的字符下标对齐。"""
    spans = []
    pos = 0
    for i, r in enumerate(para.runs):
        t = r.text or ''
        spans.append((pos, pos + len(t), i, r))
        pos += len(t)
    return spans


def _runs_covering(para, sub):
    """返回段落中每一处 sub 出现位置所覆盖的 run 列表（按出现顺序）。"""
    groups = []
    text = para.text or ''
    spans = _run_spans(para)
    for m in re.finditer(re.escape(sub), text):
        grp = [r for (s, e, _i, r) in spans if s < m.end() and e > m.start()]
        if grp:
            groups.append(grp)
    return groups


def _term_run_ids(para, sub=NEW_TERM):
    """返回覆盖 sub 的 run 在段落内的序号集合（lxml 代理对象的 id 不稳定，故用序号）。"""
    idx = set()
    text = para.text or ''
    for m in re.finditer(re.escape(sub), text):
        for (s, e, i, _r) in _run_spans(para):
            if s < m.end() and e > m.start():
                idx.add(i)
    return idx


def _visible_border(el):
    """w:pBdr 是否有可见边框（任一侧 val 非 none/nil）。"""
    if el is None:
        return False
    for side in ('top', 'bottom', 'left', 'right', 'between'):
        b = el.find(qn('w:' + side))
        if b is None:
            continue
        val = b.get(qn('w:val'))
        if val and str(val).lower() not in ('none', 'nil'):
            return True
    return False


def _open_word_document(file_path):
    """打开 .docx 文档（python-docx 仅支持 OOXML/.docx）."""
    return Document(file_path)

def _len_to_emu(raw):
    """把 VML style 里的长度（595.3pt / 21cm / 8.27in / 1916px）折成 EMU。"""
    if not raw:
        return None
    m = re.match(r'^\s*(-?[\d.]+)\s*(pt|cm|mm|in|px|pc)?\s*$', str(raw))
    if not m:
        return None
    try:
        num = float(m.group(1))
    except ValueError:
        return None
    unit = (m.group(2) or 'pt').lower()
    factor = {
        'pt': EMU_PER_PT, 'pc': EMU_PER_PT * 12.0, 'cm': EMU_PER_CM,
        'mm': EMU_PER_CM / 10.0, 'in': 914400.0, 'px': 914400.0 / PX_DPI,
    }.get(unit, EMU_PER_PT)
    return int(round(num * factor))


def _vml_style(shape):
    """把 v:shape/@style 解析成小写键的 dict。"""
    out = {}
    for seg in (shape.get('style') or '').split(';'):
        if ':' not in seg:
            continue
        k, _, v = seg.partition(':')
        out[k.strip().lower()] = v.strip()
    return out


def _image_px(part, rid):
    """按关系 id 取所嵌位图的像素尺寸 (px_w, px_h)；取不到返回 (None, None)。"""
    try:
        img_part = part.related_parts[rid]
    except Exception:
        return (None, None)
    img = getattr(img_part, 'image', None)
    if img is None and _DocxImage is not None:
        try:
            img = _DocxImage.from_blob(img_part.blob)
        except Exception:
            img = None
    if img is None:
        return (None, None)
    return (getattr(img, 'px_width', None), getattr(img, 'px_height', None))


def _hf_parts(doc):
    """返回文档的全部页眉/页脚部件。"""
    parts = []
    try:
        for part in doc.part.package.iter_parts():
            name = str(part.partname).lower()
            if re.search(r'/(header|footer)\d*\.xml$', name):
                parts.append(part)
    except Exception:
        pass
    return parts

R_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'


def _collect_watermarks(doc):
    """收集页眉/页脚中的水印对象（新式 DrawingML 与旧式 VML 两种落地方式都认）。
       每项：kind/behind/wash/cx/cy/hcenter/vcenter/px_w/px_h/is_text。"""
    out = []
    for part in _hf_parts(doc):
        root = getattr(part, 'element', None)
        if root is None:
            try:
                root = parse_xml(part.blob)
            except Exception:
                continue
        # ① 新式：w:drawing/wp:anchor（衬于文字下方 = behindDoc="1"）
        nodes = list(root.findall('.//{%s}anchor' % WP_NS))
        nodes += list(root.findall('.//{%s}inline' % WP_NS))
        for nd in nodes:
            blip = nd.find('.//{%s}blip' % A_NS)
            if blip is None:
                continue
            ext = nd.find('{%s}extent' % WP_NS)
            lum = nd.find('.//{%s}lum' % A_NS)
            bright = _int_attr(lum, 'bright') if lum is not None else None
            contrast = _int_attr(lum, 'contrast') if lum is not None else None
            aligns = [(a.getparent().tag, (a.text or '').strip().lower())
                      for a in nd.findall('.//{%s}align' % WP_NS)]
            px_w, px_h = _image_px(part, blip.get('{%s}embed' % R_NS))
            out.append({
                'kind': 'dml',
                'behind': str(nd.get('behindDoc') or '').lower() in ('1', 'true'),
                'wash': (bright is not None and contrast is not None
                         and bright >= 50000 and contrast <= -50000),
                'cx': _int_attr(ext, 'cx'), 'cy': _int_attr(ext, 'cy'),
                'hcenter': any(t.endswith('positionH') and v == 'center' for t, v in aligns),
                'vcenter': any(t.endswith('positionV') and v == 'center' for t, v in aligns),
                'px_w': px_w, 'px_h': px_h, 'is_text': False,
            })
        # ② 旧式：w:pict/v:shape（衬于文字下方 = z-index 为负）
        for sp in root.findall('.//{%s}shape' % V_NS):
            style = _vml_style(sp)
            imagedata = sp.find('{%s}imagedata' % V_NS)
            is_text = sp.find('{%s}textpath' % V_NS) is not None
            if imagedata is None and not is_text:
                continue
            rid = imagedata.get('{%s}id' % R_NS) if imagedata is not None else None
            gain = None
            black = None
            if imagedata is not None:
                gain = imagedata.get('gain') or imagedata.get('{%s}gain' % O_NS)
                black = (imagedata.get('blacklevel')
                         or imagedata.get('{%s}blacklevel' % O_NS))
            px_w, px_h = _image_px(part, rid) if rid else (None, None)
            zi = (style.get('z-index') or '').strip()
            out.append({
                'kind': 'vml',
                'behind': zi.startswith('-'),
                'wash': bool(gain or black),
                'cx': _len_to_emu(style.get('width')),
                'cy': _len_to_emu(style.get('height')),
                'hcenter': style.get('mso-position-horizontal', '').lower() == 'center',
                'vcenter': style.get('mso-position-vertical', '').lower() == 'center',
                'px_w': px_w, 'px_h': px_h, 'is_text': is_text,
            })
    return out

class _FileMissingError(Exception):
    """目录不存在或目录中没有 .docx：连交付文件都没有，不进入维度一判定，
    此时 dim1_pass / dim1_reason 留空、max_score 归 0。"""


def evaluate(dir_path):
    """按 §2.2 约定，返回结构化 dict；不 print、不改 sys.stdout、不 sys.exit。
    入参 dir_path 为“脚本所在目录的路径”，脚本自己在该目录里定位并打开被评估的
    .docx（忽略以 ~$ 开头的 Word 临时锁文件）。脚本自身崩溃（含目录不存在、
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
            raise _FileMissingError(f"目录不存在：{dir_path}")
        cands = sorted(
            f for f in os.listdir(dir_path)
            if f.lower().endswith('.docx') and not f.startswith('~$')
        )
        if not cands:
            raise _FileMissingError(f"目录中未找到 .docx 文件：{dir_path}")
        file_path = os.path.join(dir_path, cands[0])
        result["file_name"] = os.path.basename(file_path)
        return _evaluate_impl(file_path, result)
    except _FileMissingError as exc:
        # 找不到交付文件：dim1_pass / dim1_reason 留空，满分也归 0
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
        # 文件存在但打不开/解析失败：维度一判为 false，dim1_reason 给出错误摘要
        result["status"] = "error"
        result["error"] = f"{type(exc).__name__}: {exc}"
        result["dim1_pass"] = False
        result["dim1_reason"] = f"文档无法正常打开或解析：{type(exc).__name__}: {exc}"
        result["dim2_items"] = []
        result["total_score"] = 0
        return result

def _evaluate_impl(file_path, result):
    """维度一 + 维度二的评估主体；内部构造 results 列表（加分项 "+N：…"、
    减分项 "-N：…"，未命中一律写 "+0：…"/"-0：…"），末尾按 _MAX_DELTAS_ORDER
    反解成结构化 dim2_items 填入 result。"""
    doc = _open_word_document(file_path)
    results = []

    body_el = doc.element.body
    all_paras = list(doc.paragraphs)
    paras = [p for p in all_paras if (p.text or '').strip()]
    para_texts = [(p.text or '').strip() for p in paras]
    title_para = paras[0] if paras else None
    body_paras = paras[1:]
    flat_all = re.sub(r'\s+', '', ''.join((p.text or '') for p in all_paras))

    # ---- 首字下沉框定位（多条细则复用） ----
    drop_para, drop_idx, drop_fp = None, -1, None
    for i, p in enumerate(body_paras):
        fp = _frame_pr(p)
        if fp is not None and fp.get(qn('w:dropCap')):
            drop_para, drop_idx, drop_fp = p, i, fp
            break
    drop_text = (drop_para.text or '').strip() if drop_para is not None else ''
    drop_lines = _int_attr(drop_fp, 'w:lines') if drop_fp is not None else None
    drop_next = ((body_paras[drop_idx + 1].text or '').strip()
                 if 0 <= drop_idx < len(body_paras) - 1 else '')
    # 下沉框自身的行距/段后由下沉版式决定，且下沉首字必然放大：这些段落在
    # “正文段后6磅/1.1倍行距/小四”等项里按办公软件口径排除。
    drop_only = {drop_idx} if drop_idx >= 0 else set()
    # 环绕在下沉字右侧的行（下沉行数范围内）由下沉版式接管其首行缩进，
    # 故“首行缩进2字符”一项排除下沉框及其后 lines 行。
    drop_wrapped = set(drop_only)
    if drop_idx >= 0:
        for k in range(1, (drop_lines or 2) + 1):
            drop_wrapped.add(drop_idx + k)

    def _normal_runs(para):
        """段落内“普通正文 run”：含中文、且不属于 3 处替换文字的 run。"""
        tids = _term_run_ids(para)
        return [r for i, r in enumerate(para.runs)
                if ZH_RE.search(r.text or '') and i not in tids]

    # ========== 维度1：可用与可修改性 ==========
    # 严格对齐细则：本维度只判定“交付文件为 1 个可正常打开的 Word 文档（.docx）”这一点。
    # 与办公软件对齐：能被 OPC/WordprocessingML 解析器成功打开即视为可正常打开；
    #   打开失败（文件损坏、需修复、格式非法）由 evaluate 兜底为 status="error"、
    #   dim1_pass=False、总分 0。走到这里说明文档已成功打开。
    result["dim1_pass"] = True
    result["dim1_reason"] = "交付文件为可正常打开的 Word 文档（.docx）"

    # ========== 维度2：完成度评分 ==========

    # +3：首行标题“智能软件工程系简介”中文字体为隶书
    # 严格对齐细则，仅踩细则写出的点：
    #   ① 判定对象为首行标题段落中承载中文的所有 run
    #   ② 中文字体（东亚字体）为隶书
    # 与办公软件对齐：沿 w:rFonts/@w:eastAsia（含主题字体引用）→ 段落样式链 →
    #   docDefaults 解析实际生效的中文字体，即打开 Word/WPS 看到的效果。
    title_runs = [r for r in title_para.runs if ZH_RE.search(r.text or '')] if title_para else []
    lishu_ok = bool(title_runs) and all(
        (_resolve_effective_eastasia(doc, r) or '') and
        any(k in (_resolve_effective_eastasia(doc, r) or '') for k in LISHU_KEYS)
        for r in title_runs)
    if lishu_ok:
        results.append(f"+3：首行标题中文字体为隶书（{_resolve_effective_eastasia(doc, title_runs[0])}）")
    else:
        results.append("+0：首行标题中文字体为隶书")

    # +3：首行标题字号为二号（22磅 / sz=44）、设置加粗、字体颜色为蓝色
    # 严格对齐细则三点：① 字号二号（22磅，容差21.5-22.5）② 加粗 ③ 蓝色系
    # 与办公软件对齐：w:sz（半磅）、w:b、w:color/@w:val。
    t_sizes = [_resolve_effective_size_pt(doc, r) for r in title_runs]
    size2_ok = bool(t_sizes) and all(s is not None and 21.5 <= s <= 22.5 for s in t_sizes)
    bold_ok = bool(title_runs) and all(_resolve_effective_toggle(doc, r, 'w:b') for r in title_runs)
    blue_ok = bool(title_runs) and all(_is_blue(_resolve_effective_color(doc, r)) for r in title_runs)
    if size2_ok and bold_ok and blue_ok:
        results.append(
            f"+3：首行标题二号/加粗/蓝色（{t_sizes[0]}磅、加粗、"
            f"{_resolve_effective_color(doc, title_runs[0])}）")
    else:
        results.append("+0：首行标题二号/加粗/蓝色")

    # +5：首行标题添加“发光”文本效果，发光颜色为橄榄色（强调文字颜色3）、半径7磅
    # 严格对齐细则三点：① 有发光效果 ② 颜色为橄榄色/强调文字颜色3
    #   ③ 发光半径7磅（w14:glow/@w14:rad = 88900 EMU，容差 6.5-7.5磅）
    # 与办公软件对齐：Word 2010 文本效果命名空间 w14 下的 w14:glow 及其子颜色节点
    #   （srgbClr 直接给色值、schemeClr accent3 即“强调文字颜色3”）。
    glow_ok = False
    glow_detail = ''
    for r in (title_runs or []):
        rPr = _run_rpr(r)
        glow = rPr.find('{%s}glow' % W14_NS) if rPr is not None else None
        if glow is None:
            glow_ok = False
            break
        rad = _int_attr(glow, '{%s}rad' % W14_NS)
        rad_ok = rad is not None and 82550 <= rad <= 95250
        srgb = glow.find('{%s}srgbClr' % W14_NS)
        scheme = glow.find('{%s}schemeClr' % W14_NS)
        cval = srgb.get('{%s}val' % W14_NS) if srgb is not None else None
        sval = scheme.get('{%s}val' % W14_NS) if scheme is not None else None
        color_ok = _is_olive(cval) or (sval or '').lower() in ('accent3',)
        glow_ok = bool(rad_ok and color_ok)
        glow_detail = f"rad={rad}、色={cval or sval}"
        if not glow_ok:
            break
    if glow_ok:
        results.append(f"+5：首行标题发光效果（橄榄色/强调文字颜色3、7磅；{glow_detail}）")
    else:
        results.append("+0：首行标题发光效果（橄榄色/强调文字颜色3、7磅）")

    # +3：首行标题添加“阴影”文本效果且方向为向上偏移（阴影角度270°、w14:dir=16200000）
    # 严格对齐细则两点：① 有阴影效果 ② 方向为向上偏移（dir≈16200000，即270°；
    #   偏移距离约3磅、dist>0），向下/向左右等其他方向不得分。
    # 与办公软件对齐：w14:shadow/@w14:dir（1/60000 度）与 @w14:dist（EMU）。
    shadow_ok = False
    shadow_detail = ''
    for r in (title_runs or []):
        rPr = _run_rpr(r)
        sh = rPr.find('{%s}shadow' % W14_NS) if rPr is not None else None
        if sh is None:
            shadow_ok = False
            break
        d = _int_attr(sh, '{%s}dir' % W14_NS)
        dist = _int_attr(sh, '{%s}dist' % W14_NS)
        shadow_ok = (d is not None and 15900000 <= d <= 16500000
                     and dist is not None and dist > 0)
        shadow_detail = f"dir={d}、dist={dist}"
        if not shadow_ok:
            break
    if shadow_ok:
        results.append(f"+3：首行标题阴影效果为向上偏移（{shadow_detail}）")
    else:
        results.append("+0：首行标题阴影效果为向上偏移")

    # +1：首行标题段落为居中对齐
    # 与办公软件对齐：段落 w:jc/@w:val = center（含样式链继承）。
    t_jc = _para_prop(doc, title_para, 'w:jc', 'w:val') if title_para is not None else None
    if t_jc == 'center' or (title_para is not None
                            and title_para.alignment == WD_ALIGN_PARAGRAPH.CENTER):
        results.append("+1：首行标题段落居中对齐（w:jc=center）")
    else:
        results.append("+0：首行标题段落居中对齐")

    # +3：首行标题段落段后间距为20磅（w:after=400）
    # 与办公软件对齐：w:spacing/@w:after 为 twips（20磅 = 400 twips，容差 390-410）。
    t_after = _para_prop_int(doc, title_para, 'w:spacing', 'w:after') if title_para is not None else None
    if t_after is not None and 390 <= t_after <= 410:
        results.append(f"+3：首行标题段后间距20磅（w:after={t_after}）")
    else:
        results.append("+0：首行标题段后间距20磅")

    # +3：正文所有段落（第二段至末段，含5行专业名称）中文字体为楷体、字号为小四
    # 严格对齐细则两点：① 中文字体楷体 ② 字号小四（12磅，容差11.5-12.5）
    # 判定范围按办公软件口径：正文所有段落中承载中文的 run，排除 3 处
    #   “智能软件工程系”（细则另有宋体四号要求）；下沉首字必然被放大，故其
    #   字号不纳入本项（字体仍需为楷体）。
    kai_bad, size4_bad = [], []
    for i, p in enumerate(body_paras):
        for r in _normal_runs(p):
            ea = _resolve_effective_eastasia(doc, r) or ''
            if not any(k in ea for k in KAITI_KEYS):
                kai_bad.append(f"第{i + 2}段“{(r.text or '')[:6]}”={ea or '未设定'}")
            if i in drop_only:
                continue
            sz = _resolve_effective_size_pt(doc, r)
            if sz is None or not (11.5 <= sz <= 12.5):
                size4_bad.append(f"第{i + 2}段={sz}磅")
    if body_paras and not kai_bad and not size4_bad:
        results.append(f"+3：正文所有段落楷体小四（共{len(body_paras)}段）")
    else:
        results.append("+0：正文所有段落楷体小四")

    # +3：正文所有段落设置首行缩进2字符（firstLineChars=200，非空格或制表符模拟）
    # 严格对齐细则三点：① 首行缩进 ② 2字符 ③ 非空格/制表符模拟
    # 与办公软件对齐：w:ind/@w:firstLineChars（1/100 字符，Word“2字符”的原生存储）
    #   或 @w:firstLine（twips）÷ 段落字号折算 2 字符；并检查段首无空格/Tab 充数。
    ind_bad = []
    for i, p in enumerate(body_paras):
        if i in drop_wrapped:
            continue        # 下沉框及其环绕行的首行缩进由下沉版式接管
        chars = _para_prop_int(doc, p, 'w:ind', 'w:firstLineChars')
        first = _para_prop_int(doc, p, 'w:ind', 'w:firstLine')
        ok = chars is not None and 180 <= chars <= 220
        if not ok and first:
            runs = _normal_runs(p)
            base = None
            for r in runs:
                base = _resolve_effective_size_pt(doc, r)
                if base:
                    break
            base = base or 12.0
            ok = 1.8 <= (first / 20.0) / base <= 2.2
        raw = p.text or ''
        if raw[:2] in ('  ', '　　', '\t\t') or raw.startswith('\t'):
            ok = False      # 用空格/制表符模拟缩进不算
        if not ok:
            ind_bad.append(f"第{i + 2}段(firstLineChars={chars}, firstLine={first})")
    if body_paras and not ind_bad:
        results.append("+3：正文所有段落首行缩进2字符（firstLineChars=200）")
    else:
        results.append("+0：正文所有段落首行缩进2字符")

    # +1：正文所有段落段后间距为6磅（w:after=120）
    # 与办公软件对齐：w:spacing/@w:after（twips，6磅=120）；下沉框段落的段后由
    #   下沉版式决定，不纳入判定。
    after_bad = []
    for i, p in enumerate(body_paras):
        if i in drop_only:
            continue
        after = _para_prop_int(doc, p, 'w:spacing', 'w:after')
        if after is None or not (110 <= after <= 130):
            after_bad.append(f"第{i + 2}段={after}")
    if body_paras and not after_bad:
        results.append("+1：正文所有段落段后间距6磅（w:after=120）")
    else:
        results.append("+0：正文所有段落段后间距6磅")

    # +3：正文所有段落行距为1.1倍行距（w:line=264 且 lineRule=auto）
    # 与办公软件对齐：w:spacing/@w:line=264（240×1.1）且 @w:lineRule=auto；
    #   下沉框段落的行距为下沉版式专用（lineRule=exact），不纳入判定。
    line_bad = []
    for i, p in enumerate(body_paras):
        if i in drop_only:
            continue
        line = _para_prop_int(doc, p, 'w:spacing', 'w:line')
        rule = _para_prop(doc, p, 'w:spacing', 'w:lineRule')
        if line is None or not (258 <= line <= 270) or rule != 'auto':
            line_bad.append(f"第{i + 2}段(line={line}, rule={rule})")
    if body_paras and not line_bad:
        results.append("+3：正文所有段落1.1倍行距（w:line=264、lineRule=auto）")
    else:
        results.append("+0：正文所有段落1.1倍行距")

    # +3：“专业设置如下：”已从“…科学设置专业。”所在段落中拆分为独立段落，
    #      且“…科学设置专业。”仍完整留在上一段落末尾，文字无丢失、无重复
    # 与办公软件对齐：办公软件里看到的是两个独立段落；若已做悬挂首字下沉，
    #   该行会被 Word 拆成“专”（下沉框段）+“业设置如下：”（环绕段）两段，
    #   两种落地形态都算“已独立成段”。
    split_idx = -1
    for i, t in enumerate(para_texts):
        if t == SPLIT_LINE:
            split_idx = i
            break
        if t == DROP_CHAR and i + 1 < len(para_texts) and para_texts[i + 1] == SPLIT_LINE[1:]:
            split_idx = i
            break
    prev_ok = split_idx > 0 and para_texts[split_idx - 1].endswith(PREV_TAIL)
    once_ok = flat_all.count(SPLIT_LINE) == 1 and flat_all.count(PREV_TAIL) == 1
    if split_idx > 0 and prev_ok and once_ok:
        results.append("+3：“专业设置如下：”已独立成段且“…科学设置专业。”完整留在上一段末尾")
    else:
        results.append("+0：“专业设置如下：”已独立成段且“…科学设置专业。”完整留在上一段末尾")

    # +5：“专业设置如下：”的首字“专”设置“首字下沉”，下沉方式为“悬挂”
    # 严格对齐细则三点：① 有首字下沉 ② 方式为悬挂（framePr dropCap="margin"，
    #   即单字置于版心左侧页边距一侧）③ 仅该单字放大下沉、后续“业设置如下：”
    #   以小四楷体环绕其右侧正常显示。
    # 与办公软件对齐：w:pPr/w:framePr/@w:dropCap=margin，下沉框段落只含 1 个字符，
    #   紧随其后的段落承接该行余下文字。
    hang_ok = (drop_fp is not None
               and drop_fp.get(qn('w:dropCap')) == 'margin'
               and drop_text == DROP_CHAR
               and drop_next.startswith(SPLIT_LINE[1:]))
    if hang_ok:
        results.append("+5：首字“专”首字下沉且方式为悬挂（dropCap=margin，仅单字下沉）")
    else:
        results.append("+0：首字“专”首字下沉且方式为悬挂")

    # +5：首字下沉的下沉行数为2行（下沉字高度约等于两行正文高度，字号约36-40磅）
    # 与办公软件对齐：w:framePr/@w:lines=2，且下沉首字 run 的 w:sz 被放大到
    #   约 36-40磅（sz≈72-80），即版面上实际占两行高。
    drop_size = None
    if drop_para is not None:
        sizes = [_resolve_effective_size_pt(doc, r) for r in drop_para.runs
                 if ZH_RE.search(r.text or '')]
        sizes = [s for s in sizes if s]
        drop_size = max(sizes) if sizes else None
    if drop_lines == 2 and drop_size is not None and 35.0 <= drop_size <= 41.0:
        results.append(f"+5：首字下沉行数为2行（w:lines=2、下沉字{drop_size}磅）")
    else:
        results.append("+0：首字下沉行数为2行")

    # +3：首字下沉的“距正文”距离为0.3厘米（约170 twips，hSpace=170）
    # 与办公软件对齐：w:framePr/@w:hSpace（twips；0.3厘米≈170，容差150-190）。
    hspace = _int_attr(drop_fp, 'w:hSpace') if drop_fp is not None else None
    if hspace is not None and 150 <= hspace <= 190:
        results.append(f"+3：首字下沉距正文0.3厘米（hSpace={hspace}）")
    else:
        results.append("+0：首字下沉距正文0.3厘米")

    # +3：全文4处“智能工程系”（首行标题内、第二段段首、第三段段首、末段
    #      “面向岗位，面向技能，面向未来，”之后）均已改为“智能软件工程系”，
    #      无遗漏、无重复插入，且未误改“道路安全与智能运维”等其他含“智能”的文字
    # 与办公软件对齐：按可见文本判定（全文去空白后统计出现次数与位置锚点）。
    cnt_new = flat_all.count(NEW_TERM)
    cnt_old = flat_all.count(OLD_TERM)
    title_ok = bool(para_texts) and para_texts[0] == TITLE_TEXT
    head_ok = sum(1 for t in para_texts[1:] if t.startswith(NEW_TERM)) >= 2
    tail_ok = (LAST_ANCHOR + NEW_TERM) in flat_all
    other_ok = MAJORS[2] in flat_all      # “道路安全与智能运维”未被误改
    if cnt_new == 4 and cnt_old == 0 and title_ok and head_ok and tail_ok and other_ok:
        results.append("+3：全文4处均已改为“智能软件工程系”（标题/第二段首/第三段首/末段中部）")
    else:
        results.append("+0：全文4处均已改为“智能软件工程系”")

    # 3 处正文替换文字的 run 分组（下面 3 条细则复用）
    term_groups = []
    for p in body_paras:
        term_groups.extend(_runs_covering(p, NEW_TERM))

    # +3：正文3处“智能软件工程系”中文字体为宋体
    # 与办公软件对齐：沿 w:rFonts/@w:eastAsia 或主题引用（minorEastAsia→theme1.xml
    #   中 Hans 字体）解析实际生效的中文字体。
    song_ok = len(term_groups) == 3
    song_detail = ''
    for grp in term_groups:
        for r in grp:
            ea = _resolve_effective_eastasia(doc, r) or ''
            song_detail = ea or '未设定'
            if not any(k in ea for k in STRICT_SONGTI):
                song_ok = False
    if song_ok:
        results.append(f"+3：正文3处“智能软件工程系”中文字体为宋体（{song_detail}）")
    else:
        results.append("+0：正文3处“智能软件工程系”中文字体为宋体")

    # +3：上述3处字号为四号（14磅 / sz=28）、同时设置加粗与倾斜、字体颜色为绿色
    # 与办公软件对齐：w:sz、w:b、w:i、w:color/@w:val（绿色系，含 9BBB59 强调色3）。
    fmt_ok = len(term_groups) == 3
    fmt_detail = ''
    for grp in term_groups:
        for r in grp:
            sz = _resolve_effective_size_pt(doc, r)
            b = _resolve_effective_toggle(doc, r, 'w:b')
            it = _resolve_effective_toggle(doc, r, 'w:i')
            col = _resolve_effective_color(doc, r)
            fmt_detail = f"{sz}磅、加粗={b}、倾斜={it}、{col}"
            if not (sz is not None and 13.5 <= sz <= 14.5 and b and it and _is_green(col)):
                fmt_ok = False
    if fmt_ok:
        results.append(f"+3：正文3处“智能软件工程系”四号/加粗倾斜/绿色（{fmt_detail}）")
    else:
        results.append("+0：正文3处“智能软件工程系”四号/加粗倾斜/绿色")

    # +3：上述3处添加着重号（每字下方一个小圆点，w:em="dot"），而非下划线或删除线
    # 与办公软件对齐：w:rPr/w:em/@w:val=dot（“着重号”），w:u/w:strike 不算。
    em_ok = len(term_groups) == 3
    for grp in term_groups:
        for r in grp:
            rPr = _run_rpr(r)
            em = rPr.find(qn('w:em')) if rPr is not None else None
            if em is None or (em.get(qn('w:val')) or '') != 'dot':
                em_ok = False
    if em_ok:
        results.append("+3：正文3处“智能软件工程系”添加着重号（w:em=dot）")
    else:
        results.append("+0：正文3处“智能软件工程系”添加着重号")

    # ---- 水印相关（下面 4 条加分项 + 1 条减分项复用） ----
    wms = _collect_watermarks(doc)
    pic_wms = [w for w in wms if not w['is_text']]
    txt_wms = [w for w in wms if w['is_text']]
    wm = None
    for w in pic_wms:
        if w['behind']:
            wm = w
            break
    if wm is None and pic_wms:
        wm = pic_wms[0]

    # +5：给定图片以水印形式置于文档背景（插入页眉并设为衬于文字下方），
    #      正文全部文字显示在图片之上且可正常选中编辑
    # 与办公软件对齐：图片位于页眉部件中，且新式 wp:anchor/@behindDoc=1 或
    #   旧式 v:shape 的 z-index 为负 —— 即 Word 的“衬于文字下方”。
    if wm is not None and wm['behind']:
        results.append(f"+5：图片以水印形式衬于文字下方置于页眉背景（{wm['kind']}）")
    else:
        results.append("+0：图片以水印形式衬于文字下方置于页眉背景")

    # +5：水印图片应用“冲蚀”效果（亮度+70%、对比度-70%）
    # 与办公软件对齐：新式 a:blip/a:lum@bright≈70000 且 @contrast≈-70000；
    #   旧式 v:imagedata 的 gain/blacklevel 冲蚀参数。把冲蚀预先烘进位图像素
    #   （文档内无冲蚀参数）不算，因为办公软件里图片格式仍显示为未冲蚀。
    if wm is not None and wm['wash']:
        results.append("+5：水印图片应用“冲蚀”效果（亮度+70%、对比度-70%）")
    else:
        results.append("+0：水印图片应用“冲蚀”效果")

    # +5：水印缩放比例为100%（按所嵌位图像素尺寸1:1呈现，96DPI折算显示宽度≥30厘米、
    #      高度≥21厘米，纵向铺满整页版面）；被等比缩小到只覆盖页面中部一条横带不达标
    # 与办公软件对齐：显示尺寸取 wp:extent（EMU）或 v:shape style 的 width/height；
    #   位图 100% 尺寸 = 像素数 ÷ 96 × 914400 EMU。
    scale_ok = False
    scale_detail = ''
    if wm is not None and wm['cx'] and wm['cy'] and wm['px_w'] and wm['px_h']:
        nat_cx = wm['px_w'] / PX_DPI * 914400.0
        nat_cy = wm['px_h'] / PX_DPI * 914400.0
        ratio_w = wm['cx'] / nat_cx
        ratio_h = wm['cy'] / nat_cy
        scale_detail = (f"{wm['cx'] / EMU_PER_CM:.1f}×{wm['cy'] / EMU_PER_CM:.1f}厘米、"
                        f"缩放{ratio_w * 100:.0f}%")
        scale_ok = (0.95 <= ratio_w <= 1.05 and 0.95 <= ratio_h <= 1.05
                    and wm['cx'] >= 30 * EMU_PER_CM and wm['cy'] >= 21 * EMU_PER_CM)
    if scale_ok:
        results.append(f"+5：水印缩放比例100%且纵向铺满整页（{scale_detail}）")
    else:
        results.append("+0：水印缩放比例100%且纵向铺满整页")

    # +1：水印图片在页面上水平居中且垂直居中放置，且全文仅插入1次、未重复叠放多张图片
    # 与办公软件对齐：新式 wp:positionH/positionV 的 wp:align=center；
    #   旧式 mso-position-horizontal/vertical:center；并统计页眉页脚中图片水印总数。
    if wm is not None and wm['hcenter'] and wm['vcenter'] and len(pic_wms) == 1:
        results.append("+1：水印水平垂直居中且全文仅插入1张")
    else:
        results.append("+0：水印水平垂直居中且全文仅插入1张")

    # ---- 以下为减分项：命中即扣（对应细则中的 -N 条目） ----

    # -3：除首行标题及3处“智能软件工程系”外，正文各段落文字颜色不为深灰色（333333）
    # 与办公软件对齐：沿继承链解析 w:color/@w:val；未设定或 auto 即渲染为黑色，同样计入。
    color_bad = []
    for i, p in enumerate(body_paras):
        for r in _normal_runs(p):
            col = _resolve_effective_color(doc, r)
            if (col or '') != BODY_COLOR:
                color_bad.append(f"第{i + 2}段={col or '未设定'}")
    if color_bad:
        results.append(f"-3：正文文字颜色被改为非深灰333333（{color_bad[0]}，共{len(color_bad)}处）")
    else:
        results.append("-0：正文各段落文字颜色保持深灰333333")

    # -3：除3处“智能软件工程系”外，正文中出现未要求的加粗、倾斜、下划线、着重号、
    #      删除线、突出显示或字符底纹
    # 与办公软件对齐：w:b/w:i（开关型，val=0 为关不算）、w:u（val 非 none）、
    #   w:em、w:strike/w:dstrike、w:highlight（非 none）、w:shd（底纹填充非 auto）。
    deco_bad = []
    for i, p in enumerate(body_paras):
        for r in _normal_runs(p):
            rPr = _run_rpr(r)
            if rPr is None:
                continue
            hits = []
            if _toggle_on(rPr.find(qn('w:b'))):
                hits.append('加粗')
            if _toggle_on(rPr.find(qn('w:i'))):
                hits.append('倾斜')
            u = rPr.find(qn('w:u'))
            if u is not None and (u.get(qn('w:val')) or 'single') != 'none':
                hits.append('下划线')
            em = rPr.find(qn('w:em'))
            if em is not None and (em.get(qn('w:val')) or 'dot') != 'none':
                hits.append('着重号')
            if _toggle_on(rPr.find(qn('w:strike'))) or _toggle_on(rPr.find(qn('w:dstrike'))):
                hits.append('删除线')
            hl = rPr.find(qn('w:highlight'))
            if hl is not None and (hl.get(qn('w:val')) or 'none') != 'none':
                hits.append('突出显示')
            shd = rPr.find(qn('w:shd'))
            if shd is not None:
                fill = (shd.get(qn('w:fill')) or 'auto').upper()
                val = (shd.get(qn('w:val')) or 'clear').lower()
                if fill not in ('AUTO', 'FFFFFF') or val not in ('clear', 'nil'):
                    hits.append('字符底纹')
            if hits:
                deco_bad.append(f"第{i + 2}段{'/'.join(hits)}")
    if deco_bad:
        results.append(f"-3：正文出现未要求的字符格式（{deco_bad[0]}，共{len(deco_bad)}处）")
    else:
        results.append("-0：正文未出现未要求的加粗/倾斜/下划线/着重号/删除线/突出显示/底纹")

    # -3：正文各段落段前间距不为0磅
    # 与办公软件对齐：w:spacing/@w:before（twips），未设定按 0 处理。
    before_bad = []
    for i, p in enumerate(body_paras):
        before = _para_prop_int(doc, p, 'w:spacing', 'w:before', 0) or 0
        if before != 0:
            before_bad.append(f"第{i + 2}段={before / 20.0:.1f}磅")
    if before_bad:
        results.append(f"-3：正文段前间距非0磅（{before_bad[0]}，共{len(before_bad)}段）")
    else:
        results.append("-0：正文各段落段前间距为0磅")

    # -1：除首行标题居中外，其余各段落（含5行专业名称）对齐方式不为两端对齐
    # 与办公软件对齐：w:jc/@w:val=both（含样式链继承）。
    jc_bad = []
    for i, p in enumerate(body_paras):
        jc = _para_prop(doc, p, 'w:jc', 'w:val')
        if jc != 'both':
            jc_bad.append(f"第{i + 2}段={jc or '未设定(左对齐)'}")
    if jc_bad:
        results.append(f"-1：正文段落对齐方式非两端对齐（{jc_bad[0]}，共{len(jc_bad)}段）")
    else:
        results.append("-0：除标题居中外其余段落均为两端对齐")

    # -1：页面纸张不为A4纵向（11906×16838 twips），或页边距不为上下2.54厘米、
    #      左右约3.17厘米
    # 与办公软件对齐：w:sectPr/w:pgSz（@w:w/@w:h/@w:orient）与 w:pgMar；
    #   未写 w:pgMar 视为未改动页面设置，不计扣分。
    page_bad = []
    sectPrs = body_el.findall(qn('w:sectPr'))
    sect = sectPrs[-1] if sectPrs else None
    pgSz = sect.find(qn('w:pgSz')) if sect is not None else None
    if pgSz is None:
        page_bad.append('未设置纸张')
    else:
        w = _int_attr(pgSz, 'w:w')
        h = _int_attr(pgSz, 'w:h')
        orient = pgSz.get(qn('w:orient'))
        if not (w and h and abs(w - 11906) <= 60 and abs(h - 16838) <= 60):
            page_bad.append(f'纸张={w}×{h}')
        if orient == 'landscape':
            page_bad.append('横向')
    pgMar = sect.find(qn('w:pgMar')) if sect is not None else None
    if pgMar is not None:
        top = _int_attr(pgMar, 'w:top')
        bottom = _int_attr(pgMar, 'w:bottom')
        left = _int_attr(pgMar, 'w:left')
        right = _int_attr(pgMar, 'w:right')
        for name, got, want, tol in (('上', top, 1440, 40), ('下', bottom, 1440, 40),
                                     ('左', left, 1797, 60), ('右', right, 1797, 60)):
            if got is None or abs(got - want) > tol:
                page_bad.append(f'{name}边距={got}')
    if page_bad:
        results.append(f"-1：页面设置被改（{'、'.join(page_bad)}）")
    else:
        results.append("-0：页面为A4纵向且页边距为上下2.54厘米、左右约3.17厘米")

    # -3：“计算机应用技术”“软件开发技术”“道路安全与智能运维”“城市轨道交通运维”
    #      “数字媒体艺术设计”未各占独立一行、先后顺序被调整，或被合并进相邻段落
    # 与办公软件对齐：5 个专业名称须各自单独成段（段落文本恰好等于该名称）、
    #   且按给定先后顺序连续排列。
    major_idx = []
    for m in MAJORS:
        major_idx.append(para_texts.index(m) if m in para_texts else -1)
    majors_ok = (all(i >= 0 for i in major_idx)
                 and major_idx == sorted(major_idx)
                 and major_idx[-1] - major_idx[0] == len(MAJORS) - 1)
    if majors_ok:
        results.append("-0：5行专业名称各占独立一行且先后顺序未变")
    else:
        results.append(f"-3：5行专业名称未各占独立一行或顺序被调整（段落序号{major_idx}）")

    # -3：正文数据被改写（教师/职称/面积/实训室/机房/实训基地等数字）
    # 与办公软件对齐：按可见文本逐条核对细则点名的数据串是否原样存在。
    data_bad = [k for k in DATA_KEYS if k not in flat_all]
    if data_bad:
        results.append(f"-3：正文数据被改写（缺失{'、'.join(data_bad)}）")
    else:
        results.append("-0：正文各项数据未被改写")

    # -1：机构名称被改写删除，或中文引号内的固定表述被改写、引号被替换为英文引号
    # 与办公软件对齐：按可见文本核对机构名与带中文引号的固定表述是否原样存在。
    name_bad = [k for k in NAME_KEYS if k not in flat_all]
    if name_bad:
        results.append(f"-1：机构名称或中文引号内表述被改写（缺失{'、'.join(name_bad)}）")
    else:
        results.append("-0：机构名称与中文引号内固定表述均未被改写")

    # -1：文档中新增指令未要求的内容，如“专业介绍”“办学理念”等小标题、自动编号、
    #      项目符号、分隔线、总结语、空白段落、空行、分页符、分节符，或末段之后
    #      产生多余空白页
    # 与办公软件对齐：空段落（既无文字也无图片的 w:p）、w:numPr（自动编号/项目符号）、
    #   w:br@type=page 与 pageBreakBefore（分页符）、段落内 w:sectPr（分节符）、
    #   w:pBdr（分隔线），以及出现了应有内容之外的新文字段落。
    extra = []
    for p in all_paras:
        if (p.text or '').strip():
            continue
        if p._element.findall('.//' + qn('w:drawing')) or p._element.findall('.//' + qn('w:pict')):
            continue
        extra.append('空段落/空行')
        break
    if body_el.findall('.//' + qn('w:numPr')):
        extra.append('自动编号/项目符号')
    if any((br.get(qn('w:type')) or '') == 'page' for br in body_el.findall('.//' + qn('w:br'))):
        extra.append('分页符')
    for p in all_paras:
        pPr = _pPr(p)
        if pPr is None:
            continue
        if pPr.find(qn('w:sectPr')) is not None:
            extra.append('分节符')
            break
    for p in all_paras:
        pPr = _pPr(p)
        if pPr is not None and _visible_border(pPr.find(qn('w:pBdr'))):
            extra.append('段落分隔线/边框')
            break
    for t in para_texts:
        flat = re.sub(r'\s+', '', t)
        if flat and flat not in _EXPECTED_NEW and flat not in _EXPECTED_OLD:
            extra.append(f'新增文字“{flat[:12]}”')
            break
    if extra:
        results.append(f"-1：新增了指令未要求的内容（{'、'.join(extra)}）")
    else:
        results.append("-0：未新增小标题/编号/空段/分页符等指令未要求的内容")

    # -1：页眉或页脚中出现除水印图片以外的文字、页码、横线或边框
    # 与办公软件对齐：遍历页眉/页脚部件，检查是否有可见 w:t 文字、PAGE 域（页码），
    #   或段落 w:pBdr 可见边框（页眉横线）。
    hf_bad = []
    for part in _hf_parts(doc):
        root = getattr(part, 'element', None)
        if root is None:
            continue
        for t in root.findall('.//' + qn('w:t')):
            if (t.text or '').strip():
                hf_bad.append(f"文字“{(t.text or '').strip()[:8]}”")
                break
        for it in root.findall('.//' + qn('w:instrText')):
            if 'PAGE' in (it.text or '').upper():
                hf_bad.append('页码域')
                break
        for pPr in root.findall('.//' + qn('w:pPr')):
            if _visible_border(pPr.find(qn('w:pBdr'))):
                hf_bad.append('页眉横线/边框')
                break
    if hf_bad:
        results.append(f"-1：页眉页脚出现除水印图片以外的内容（{'、'.join(hf_bad[:3])}）")
    else:
        results.append("-0：页眉页脚除水印图片外无文字/页码/横线/边框")

    # -3：将“专业设置如下：”整行文字置入首字下沉框内，导致该行以超大字号单独占据版面
    # 与办公软件对齐：带 w:framePr@dropCap 的段落内只应保留 1 个下沉字符，
    #   其余文字须留在紧随其后的正文段落中。
    drop_flat = re.sub(r'\s+', '', drop_text)
    if drop_fp is not None and len(drop_flat) > 1:
        results.append("-3：整行文字被置入首字下沉框（框内文字“%s”）" % drop_flat[:10])
    else:
        results.append("-0：首字下沉框内未置入整行文字")

    # -3：正文文字被转为图片、艺术字、文本框或表格承载，导致文字不可编辑
    # 与办公软件对齐：正文中出现 w:tbl（表格）、w:txbxContent（文本框）或
    #   v:textpath（艺术字）且其中含中文字符，即判定正文被改用其他载体承载。
    carrier = []
    for tbl in body_el.findall('.//' + qn('w:tbl')):
        if ZH_RE.search(''.join(t.text or '' for t in tbl.findall('.//' + qn('w:t')))):
            carrier.append('表格')
            break
    for tbx in body_el.findall('.//' + qn('w:txbxContent')):
        if ZH_RE.search(''.join(t.text or '' for t in tbx.findall('.//' + qn('w:t')))):
            carrier.append('文本框')
            break
    for tp in body_el.findall('.//{%s}textpath' % V_NS):
        if ZH_RE.search(tp.get('string') or ''):
            carrier.append('艺术字')
            break
    if carrier:
        results.append(f"-3：正文文字被其他载体承载（{'、'.join(carrier)}）")
    else:
        results.append("-0：正文文字未被图片/艺术字/文本框/表格承载")

    # -3：水印图片被拉伸变形（长宽比与图片自身长宽比不一致），或被替换为其他图片、文字水印
    # 与办公软件对齐：比较显示尺寸长宽比（wp:extent 或 v:shape@style 的 width/height）
    #   与位图像素长宽比，偏差超过 5% 判为变形；出现 v:textpath 判为文字水印。
    wm_bad = []
    for wm in wms:
        if wm['is_text']:
            wm_bad.append('被替换为文字水印')
            continue
        if wm['cx'] and wm['cy'] and wm['px_w'] and wm['px_h']:
            shown = float(wm['cx']) / float(wm['cy'])
            native = float(wm['px_w']) / float(wm['px_h'])
            if native > 0 and abs(shown - native) / native > 0.05:
                wm_bad.append(f'长宽比失真（显示{shown:.3f} vs 图片{native:.3f}）')
    if wm_bad:
        results.append(f"-3：水印被拉伸变形或被替换（{'；'.join(wm_bad[:2])}）")
    else:
        results.append("-0：水印保持图片自身长宽比且未被替换")

    # -1：正文文字被添加发光、阴影、映像、三维等首行标题以外的文本效果
    # 与办公软件对齐：正文段落 run 的 w:rPr 下不应出现 w14:glow / w14:shadow /
    #   w14:reflection / w14:scene3d / w14:props3d / w14:textOutline；
    #   w14:textFill 仅为文字填充色，不属于发光/阴影类效果，不计入。
    EFFECT_TAGS = ('glow', 'shadow', 'reflection', 'scene3d', 'props3d', 'textOutline')
    eff_bad = []
    for idx, para in enumerate(all_paras):
        if title_para is not None and para._element is title_para._element:
            continue
        for run in para.runs:
            rpr = _run_rpr(run)
            if rpr is None:
                continue
            for name in EFFECT_TAGS:
                if rpr.find('{%s}%s' % (W14_NS, name)) is not None:
                    eff_bad.append(f'第{idx + 1}段“{name}”')
                    break
            if eff_bad:
                break
        if eff_bad:
            break
    if eff_bad:
        results.append(f"-1：正文文字被添加标题以外的文本效果（{'、'.join(eff_bad[:2])}）")
    else:
        results.append("-0：正文文字未被添加发光/阴影/映像/三维等文本效果")

    # ---------------------------------------------------------------- 汇总
    # 按接口统一约定：results 按 _MAX_DELTAS_ORDER 的顺序逐条位置对齐反解析。
    # 加分项 "+N：…" 命中记 N 分、"+0：…" 未命中记 0 分；
    # 减分项 "-N：…" 表示扣分触发（delta 取负值），"-0：…" 表示未触发。
    if len(results) != len(_MAX_DELTAS_ORDER):
        raise RuntimeError(f"dim2 项数不一致：results={len(results)} vs order={len(_MAX_DELTAS_ORDER)}")
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
    if not result["dim1_pass"]:
        total_score = 0
    result["dim2_items"] = dim2_items
    result["total_score"] = max(0, total_score)
    result["max_score"] = MAX_SCORE_TOTAL
    return result


if __name__ == "__main__":
    if len(sys.argv) >= 2:
        _dir = sys.argv[1]
    else:
        _dir = os.path.dirname(os.path.abspath(__file__))
    _out = evaluate(_dir)
    print(json.dumps(_out, ensure_ascii=False, indent=2))
