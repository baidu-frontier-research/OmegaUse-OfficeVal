#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# pyright: reportUnknownMemberType=false, reportUnknownVariableType=false, reportUnknownArgumentType=false, reportUnknownParameterType=false, reportMissingParameterType=false, reportPrivateUsage=false, reportUnknownLambdaType=false, reportGeneralTypeIssues=false

import os
import re
import json
import hashlib
import zipfile
from pptx import Presentation

# 说明：按接口统一约定，本模块不修改全局 sys.stdout，也不 print 主结果、不 sys.exit。
# 主结果统一由 evaluate() -> dict 返回；评估目标锁定为本脚本所在目录，脚本自己在该
# 目录内定位并打开被评估的 .pptx，不接受外部传入的目录，也不检索子目录与上级目录。
# 本文件仅在 __main__ 下作为本地调试入口，才把 dict 序列化为 JSON 打印。
#
# 与 008 号脚本的差异：本任务（「启蒙运动思想解放.pptx」第1页分步单击动画）细则同时含
# “加分项”（指令下达的动画改造与素材搬运任务，22 条、满分 +70）与“减分项”（指令未涉及
# 却破坏第1页既有内容/版面、或改动素材三页的多余修改，28 条、合计 -44）。故 dim2 项的
# max_delta 允许为负：正值=加分项（命中记满分，未命中记 0）；负值=减分项（触发“破坏”记
# 该负分，未触发记 0）。total_score 为全部 delta 之和，取值区间 [-44, +70]；max_score 仍
# 为所有加分项满分之和（+70），作为完成度上限。
#
# 判定不依赖 Office / LibreOffice 渲染：全部基于 python-pptx 读取的形状树（含组合内部的
# 绝对坐标换算）与 ppt/slides/slideN.xml 的 p:timing 单击序列。

SCRIPT_ID = "077"

# 评估目标锁定为本脚本所在目录：评测时把本脚本与被评交付物放在同一个目录下即可。
_SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

# —— 加分项按“结果追加顺序”的 max_delta（正值）——
_ADD_DELTAS_ORDER = [
    3,   # A01 “思想脉络”版块及三行内容存在且未设任何动画（前提：已加分步单击动画）
    5,   # A02 整套单击过程中三行内容始终可见，未被整页遮挡层/材料卡片压盖
    3,   # A03 七张金框肖像均设淡化进入动画、起始不显示
    3,   # A04 七条竖线短句均设淡化进入动画、起始不显示
    1,   # A05 上述进入动画触发方式均为“单击时”
    3,   # A06 第1次单击使洛克“自然权利｜政府同意”淡入
    1,   # A07 洛克不设材料显示/消失步骤，仅占1次单击
    5,   # A08 孟德斯鸠三步单击序列
    5,   # A09 伏尔泰三步单击序列
    5,   # A10 卢梭三步单击序列
    5,   # A11 休谟三步单击序列
    5,   # A12 亚当·斯密三步单击序列
    5,   # A13 康德三步单击序列
    3,   # A14 全页 19 步、18 步无交叉无错位无缺步多步
    3,   # A15 六组材料在第1页内淡入，未新增幻灯片/跳页/超链接
    3,   # A16 每组材料构件完整
    3,   # A17 材料与人物一一对应无错配
    3,   # A18 材料区域背景洁净、与版块文字无重叠串行
    1,   # A19 每组材料统一编入同一次单击，一次淡入/淡出整体出现消失
    1,   # A20 进入与退出统一使用“淡化/淡出”，未混用其他效果
    1,   # A21 材料淡出后无残留，页面回到卡片与“思想脉络”同时可见
    3,   # A22 动画仅加在第1页，素材三页保持不变（前提：已加分步单击动画）
]

# —— 减分项按“结果追加顺序”的 max_delta（负值）——
_DED_DELTAS_ORDER = [
    -3,  # D01 七张金框肖像图片缺失/被删/被替换为文字单元
    -3,  # D02 七个人物卡片版式被改写
    -3,  # D03 七条竖线短句被改写或格式被改
    -1,  # D04 人物排列顺序或姓名用字被改动
    -3,  # D05 “思想脉络”版块被删除或改为表格式版面
    -1,  # D06 “思想脉络”标签文字格式被改
    -1,  # D07 三个 0.18×0.18 英寸小方块缺失或被改色
    -1,  # D08 三个时期标签与三个人物栏的文字或格式被改动
    -1,  # D09 三行说明文字被改写或格式被改
    -1,  # D10 为版块新增本不存在的栏目文字
    -1,  # D11 页面主标题文字或格式被改
    -3,  # D12 副标题文字被改写
    -1,  # D13 副标题格式被改
    -1,  # D14 左上角“理性主义的觉醒”格式被改
    -1,  # D15 右下角“第8课 欧洲的思想解放运动”格式被改
    -1,  # D16 上下两条金色横线缺失或被改为填充色块
    -3,  # D17 底部四个深青底标签缺失或被删除
    -1,  # D18 四个标签文字格式被改
    -1,  # D19 中文（东亚）字体设定丢失或被改名
    -1,  # D20 新增覆盖整页的纯色矩形底板
    -1,  # D21 幻灯片尺寸/母版版式被替换
    -1,  # D22 新增操作说明类文字
    -1,  # D23 材料要点被加项目符号或被重新分段改写
    -1,  # D24 材料中的人物身份说明与人物不符
    -1,  # D25 材料肖像或文本框超出页面边界
    -3,  # D26 素材三页被删除、清空或改写内容
    -3,  # D27 删除第1页既有图片/图形/文本框导致内容缺失
    -1,  # D28 为第1页以外的页面添加动画或页面切换效果
]

MAX_SCORE_TOTAL = sum(_ADD_DELTAS_ORDER)               # +70：完成度上限
_MIN_TOTAL = MAX_SCORE_TOTAL + sum(_DED_DELTAS_ORDER)  # +26：全部命中且全部踩坑
assert MAX_SCORE_TOTAL == 70
assert sum(_DED_DELTAS_ORDER) == -44

# 结果追加顺序 = 先加分项、后减分项
_MAX_DELTAS_ORDER = _ADD_DELTAS_ORDER + _DED_DELTAS_ORDER
_ADD_COUNT = len(_ADD_DELTAS_ORDER)

NS = {
    'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
    'p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
}
EMU_IN = 914400.0                 # 1 英寸 = 914400 EMU
_PAGE_W_IN = 13.3333              # 宽屏 13.33×7.5 英寸
_PAGE_H_IN = 7.5
_PAGE_TOL_IN = 0.02               # 幻灯片尺寸容差
_OOB_TOL_IN = 0.05                # 越界容差
_COVER_TOL = 0.15                 # 材料构件对“思想脉络”三行说明文字的不透明压盖面积比上限

# —— 第1页七位思想家（自左向右横排顺序）——
_PERSONS = ['洛克', '孟德斯鸠', '伏尔泰', '卢梭', '休谟', '亚当·斯密', '康德']

# —— 卡片下方七条竖线短句（“各自主张”）——
_CLAIMS = {
    '洛克': '自然权利｜政府同意',
    '孟德斯鸠': '分权制衡｜法律理性',
    '伏尔泰': '思想自由｜批判专制',
    '卢梭': '社会契约｜主权在民',
    '休谟': '经验主义｜理性边界',
    '亚当·斯密': '分工市场｜自由竞争',
    '康德': '独立思考｜自由自律',
}

# —— 六组材料的身份说明 ——
_IDENTITY = {
    '孟德斯鸠': '国家学说与法学理论的奠基者之一',
    '伏尔泰': '法国启蒙运动的重要代表',
    '卢梭': '社会契约与人民主权的重要思想家',
    '休谟': '苏格兰启蒙运动与经验主义代表',
    '亚当·斯密': '现代经济学的重要奠基者',
    '康德': '启蒙思想的总结者之一',
}

# —— 六组材料的引文（不含引号）——
_QUOTES = {
    '孟德斯鸠': '要避免权力失衡，就必须让权力与权力相互约束。',
    '伏尔泰': '自由不是任性，而是在法律保护下表达、选择并承担责任。',
    '卢梭': '合法的政治秩序来自人民共同认可，而不是少数人的命令。',
    '休谟': '理性需要经验校验，制度变革也需要现实条件支撑。',
    '亚当·斯密': '财富增长不仅依赖资源，更依赖分工、交换与制度秩序。',
    '康德': '自由不是没有边界，而是能够以理性为自己立法。',
}

# —— 六组材料“观点提炼”下的要点句（源自素材三页，按段落顺序）——
_KEYPOINTS = {
    '孟德斯鸠': [
        '提出分权思想，将立法、行政与司法权分开运行。',
        '强调权力之间相互制约，避免集中过度导致滥用。',
        '认为法律应体现理性，并与社会环境相适应。',
    ],
    '伏尔泰': [
        '倡导思想与言论自由，强调自然权利和平等人格。',
        '批判宗教专断与封建特权，主张以法律约束权力。',
        '重视财产权与文明多元，反对单一价值压制。',
    ],
    '卢梭': [
        '强调自由和平等是人的基本权利。',
        '主张政治共同体应建立在社会契约与共同意志之上。',
        '认为主权属于人民，政府必须回应公共意志。',
        '批评贫富分化造成的社会不平等。',
    ],
    '休谟': [
        '认为知识起于经验，反对把理性神化。',
        '对因果必然性保持怀疑，强调习惯和经验判断。',
        '将情感纳入道德解释，突出同情与社会秩序。',
        '政治上更重视渐进改良与制度稳定。',
    ],
    '亚当·斯密': [
        '把劳动视为财富形成的关键来源之一。',
        '强调分工能够提高劳动生产率。',
        '主张自由竞争，反思重商主义的限制。',
        '认为市场机制需要信任、规则与道德基础配合。',
    ],
    '康德': [
        '倡导独立思考，强调人应运用自己的理性。',
        '提出“人非工具”的伦理立场，确认人的主体地位。',
        '支持自由、平等与法治，但自由必须与自律相结合。',
        '把启蒙理解为摆脱不成熟状态的勇气。',
    ],
}

# —— “思想脉络”版块基线 ——
_PANEL_LABEL = '思想脉络'
_PANEL_FILL = 'F6EFE0'
_PANEL_LINE = 'D7B46A'
_SQUARE_COLORS = ['2AA6A5', 'D7B46A', 'A65A45']       # 三行行首小方块，自上而下
_PERIODS = ['17世纪英国', '18世纪法国', '18世纪后期']
_PERSON_COLS = ['洛克', '孟德斯鸠、伏尔泰、卢梭', '休谟、亚当·斯密、康德']
_PANEL_ROWS = [
    '自然权利、财产权与有限政府，为后来权利政治提供理论起点。',
    '分权制衡、言论自由、社会契约与主权在民共同推动法国成为思想中心。',
    '经验主义、自由市场与理性批判继续扩展启蒙运动的影响。',
]

# —— 标题、角标与底部标签基线 ——
_TITLE = '四、启蒙运动——理性主义的觉醒'
_SUBTITLE = '代表人物及思想主张：从“权利”到“法治”的思想谱系'
_CORNER_TL = '理性主义的觉醒'
_CORNER_BR = '第8课  欧洲的思想解放运动'
_TAGS = ['理性至上', '天赋人权', '批判专制与教权', '民主与法治']
_TAG_FILL = '0F3C45'
_TAG_LINE = '2AA6A5'

# —— 卡片／版块／标题／标签的基线尺寸与文字格式（sz 为百分之一磅，b 为加粗，末位为字色）——
_LINE_W = 12700                            # 1 磅描边
_CARD_FRAME = (1.63, 1.95, 'D7B46A')       # 金框：宽、高、填充色
_CARD_PORTRAIT = (1.56, 1.88)              # 金框内肖像图片尺寸
_CARD_BAR = (1.56, 0.33, '000000')         # 姓名条底色块
_CARD_NAME_FMT = (870, True, 'FFFFFF')     # 卡片姓名
_CARD_CLAIM_FMT = (720, False, 'F7F3E8')   # 卡片下方竖线短句
_PANEL_BOX = (0.55, 4.06, 12.20, 2.20)     # “思想脉络”版块底板：左、上、宽、高
_PANEL_LABEL_FMT = (1050, True, 'A65A45')
_SQUARE_SIDE = 0.18                        # 行首小方块边长
_PERIOD_FMT = (860, True, '163038')        # 时期标签
_PERSON_COL_FMT = (860, True, 'A65A45')    # 人物栏
_PANEL_ROW_FMT = (790, False, '163038')    # 三行说明文字
_TITLE_FMT = (1900, True, 'F7F3E8')
_SUBTITLE_FMT = (880, True, 'D7B46A')
_CORNER_TL_FMT = (800, True, '9CB7B8')
_CORNER_BR_FMT = (750, False, '9CB7B8')
_TAG_BOX = (2.87, 0.50)                    # 底部标签底板尺寸
_TAG_FMT = (820, True, 'F7F3E8')
_GOLD_LINE_W = 13.33                       # 上下两条金色横线长度
_BOX_TOL = 0.05                            # 尺寸比对容差（英寸）

# —— 中文（东亚）字体：Microsoft YaHei 与其中文名视为同一字体 ——
_EA_OK = ('Microsoft YaHei', 'MicrosoftYaHei', 'Microsoft Yahei', '微软雅黑')

# —— 素材三页（第2—4页）基线签名：(去空白正文 md5, p:sp 计数, p:pic 计数) ——
# 仅比对“内容”，不比对组合层级：重新保存/重新组合不改变文字与图形计数。
_SRC_SIG = {
    2: ('f2e747159bf2479e531bb2620200fe53', 29, 2),
    3: ('92d6ba11eff8f6c9d4c757e72ecf6695', 29, 2),
    4: ('c6718e3ed45444c73589a4c5749ecf64', 36, 2),
}
_SRC_TITLES = ['法国中心：理性批判与制度设计', '从人民主权到经验理性', '向欧洲其他地区扩展']

# —— 指令未要求的“操作说明类”文字特征 ——
_INSTR_MARKS = (
    '单击顺序', '点击顺序', '动画顺序', '初始隐藏', '材料阅读', '单击后淡出',
    '点击逐步呈现', '逐步呈现', '点击查看', '单击查看', '操作说明', '放映说明',
)
# —— “表格式版面”特征：为版块新增的分栏栏目文字 ——
_TABLE_WORDS = ('时间', '国家', '代表人物', '主要主张', '17世纪', '18世纪',
                '英国', '法国', '德国', '苏格兰')

def _norm(text):
    """去掉所有空白后的文本，用于与基线做“文字是否被改写”的比对。"""
    return re.sub(r'\s+', '', text or '')


def _tag(el):
    """返回元素的本地名（去掉命名空间）。"""
    return el.tag.rsplit('}', 1)[-1] if isinstance(el.tag, str) else ''


def _grp_xfrm(sp):
    """读取组合的 a:xfrm（off/ext/chOff/chExt），用于把组内子元素换算到页面绝对坐标。"""
    x = sp._element.find('./p:grpSpPr/a:xfrm', NS)
    if x is None:
        return None
    off = x.find('./a:off', NS)
    ext = x.find('./a:ext', NS)
    cho = x.find('./a:chOff', NS)
    che = x.find('./a:chExt', NS)
    if off is None or ext is None or cho is None or che is None:
        return None
    return (int(off.get('x')), int(off.get('y')), int(ext.get('cx')), int(ext.get('cy')),
            int(cho.get('x')), int(cho.get('y')), int(che.get('cx')), int(che.get('cy')))


def _map_box(box, tf):
    """把组内坐标 box=(l, t, w, h) 按父组变换 tf 换算为绝对坐标。"""
    if tf is None or box[0] is None:
        return box
    ox, oy, ex, ey, cx, cy, cw, ch = tf
    sx = (ex / cw) if cw else 1.0
    sy = (ey / ch) if ch else 1.0
    return (ox + (box[0] - cx) * sx, oy + (box[1] - cy) * sy, box[2] * sx, box[3] * sy)


def _walk(shapes, tf=None, chain=()):
    """深度遍历形状树，逐个产出 (shape, 绝对 l/t/w/h, 自身及全部祖先组合的 id 链)。
    组合内部元素的动画通常挂在外层组合上，故用 id 链判断“是否被某步动画命中”。"""
    for sp in shapes:
        try:
            box = (sp.left, sp.top, sp.width, sp.height)
        except Exception:
            box = (None, None, None, None)
        box = _map_box(box, tf)
        ids = tuple(chain) + (sp.shape_id,)
        yield sp, box, ids
        if str(sp.shape_type) == 'GROUP (6)':
            g = _grp_xfrm(sp)
            if g is not None and tf is not None:
                mapped = _map_box((g[0], g[1], g[2], g[3]), tf)
                g = (mapped[0], mapped[1], mapped[2], mapped[3], g[4], g[5], g[6], g[7])
            for item in _walk(sp.shapes, g, ids):
                yield item


def _kind_of(sp):
    st = str(sp.shape_type)
    if st == 'GROUP (6)':
        return 'grp'
    if st == 'PICTURE (13)':
        return 'pic'
    return 'sp'


def _fill_of(el):
    """返回 (纯色填充色值 或 None, 是否显式无填充)。"""
    pr = el.find('./p:spPr', NS)
    if pr is None:
        return None, False
    clr = pr.find('./a:solidFill/a:srgbClr', NS)
    return (clr.get('val').upper() if clr is not None and clr.get('val') else None,
            pr.find('./a:noFill', NS) is not None)


def _ln_of(el):
    """返回 (线宽 EMU 或 None, 线色 或 None)。"""
    ln = el.find('./p:spPr/a:ln', NS)
    if ln is None:
        return None, None
    clr = ln.find('./a:solidFill/a:srgbClr', NS)
    w = ln.get('w')
    return (int(w) if w and w.isdigit() else None,
            clr.get('val').upper() if clr is not None and clr.get('val') else None)


def _prst_of(el):
    g = el.find('./p:spPr/a:prstGeom', NS)
    return g.get('prst') if g is not None else None


def _runs_of(sp):
    """收集全部 run 的字号（百分之一磅）、加粗、东亚字体名与字色。"""
    out = []
    for para in sp.text_frame.paragraphs:
        for run in para.runs:
            rPr = run._r.find('./a:rPr', NS)
            sz = b = ea = color = None
            if rPr is not None:
                sz = int(rPr.get('sz')) if (rPr.get('sz') or '').isdigit() else None
                b = rPr.get('b') == '1'
                ea_el = rPr.find('./a:ea', NS)
                ea = ea_el.get('typeface') if ea_el is not None else None
                clr = rPr.find('./a:solidFill/a:srgbClr', NS)
                color = clr.get('val').upper() if clr is not None and clr.get('val') else None
            out.append({'sz': sz, 'b': bool(b), 'ea': ea, 'color': color,
                        'text': run.text or ''})
    return out


def _recs(slide):
    """把一页的形状树拍平为便于比对的记录列表；坐标单位换算为英寸，z 为绘制层序。"""
    out = []
    for z, (sp, box, ids) in enumerate(_walk(slide.shapes)):
        el = sp._element
        kind = _kind_of(sp)
        fill, nofill = (None, False) if kind == 'grp' else _fill_of(el)
        lw, lc = (None, None) if kind == 'grp' else _ln_of(el)
        has_tf = False
        try:
            has_tf = bool(sp.has_text_frame)
        except Exception:
            has_tf = False
        txt = sp.text_frame.text if has_tf else ''
        algn = None
        if has_tf and sp.text_frame.paragraphs:
            pPr = sp.text_frame.paragraphs[0]._p.find('./a:pPr', NS)
            algn = pPr.get('algn') if pPr is not None else None
        out.append({
            'z': z, 'id': sp.shape_id, 'name': sp.name or '', 'kind': kind, 'ids': ids,
            'l': (box[0] / EMU_IN) if box[0] is not None else None,
            't': (box[1] / EMU_IN) if box[1] is not None else None,
            'w': (box[2] / EMU_IN) if box[2] is not None else None,
            'h': (box[3] / EMU_IN) if box[3] is not None else None,
            'text': txt, 'norm': _norm(txt),
            'paras': [p.text or '' for p in sp.text_frame.paragraphs] if has_tf else [],
            'algn': algn,
            'fill': fill, 'nofill': nofill, 'lw': lw, 'lc': lc,
            'prst': None if kind == 'grp' else _prst_of(el),
            'runs': _runs_of(sp) if has_tf else [],
        })
    return out


def _rect(r):
    """记录的外框矩形 (x0, y0, x1, y1)，英寸；无坐标时返回 None。"""
    if r['l'] is None or r['w'] is None:
        return None
    return (r['l'], r['t'], r['l'] + r['w'], r['t'] + r['h'])


def _area(rc):
    return max(0.0, rc[2] - rc[0]) * max(0.0, rc[3] - rc[1]) if rc else 0.0


def _est_rect(r):
    """估算左对齐文本的实际占位矩形：文本框常远宽于文字，按“一个中文字≈一个字宽”折算。
    仅用于“思想脉络”三行说明文字这类左对齐长句的压盖判断。"""
    rc = _rect(r)
    if rc is None or not r['norm']:
        return rc
    sz = next((run['sz'] for run in r['runs'] if run['sz']), None)
    if not sz:
        return rc
    em = sz / 100.0 / 72.0                       # 百分之一磅 → 英寸
    widest = 0.0
    for line in (r['paras'] or [r['text']]):
        n = 0.0
        for ch in (line or ''):
            n += 1.0 if ord(ch) > 0x2E80 else 0.55
        widest = max(widest, n * em)
    return (rc[0], rc[1], min(rc[2], rc[0] + widest), rc[3])


def _cover_ratio(target, covers):
    """网格采样估算 target 被 covers 并集覆盖的面积比（避免多矩形重叠时重复计面积）。"""
    if target is None or not covers:
        return 0.0
    x0, y0, x1, y1 = target
    if x1 <= x0 or y1 <= y0:
        return 0.0
    nx, ny = 80, 12
    hit = 0
    for i in range(nx):
        px = x0 + (i + 0.5) * (x1 - x0) / nx
        for j in range(ny):
            py = y0 + (j + 0.5) * (y1 - y0) / ny
            for c in covers:
                if c and c[0] <= px <= c[2] and c[1] <= py <= c[3]:
                    hit += 1
                    break
    return hit / float(nx * ny)


def _opaque(r):
    """是否属于会遮挡下层内容的不透明构件：图片，或带纯色填充的图形。"""
    if r['kind'] == 'pic':
        return True
    if r['kind'] == 'grp':
        return False
    return bool(r['fill']) and not r['nofill']


def _spid_of(el):
    """取动画行为的目标形状 id（p:cBhvr/p:tgtEl/p:spTgt@spid）。"""
    t = el.find('./p:cBhvr/p:tgtEl/p:spTgt', NS)
    if t is None:
        t = el.find('.//p:spTgt', NS)
    return t.get('spid') if t is not None else None


_MOVE_TAGS = ('anim', 'animMotion', 'animRot', 'animScale', 'animClr')


def _click_steps(slide):
    """解析 p:timing 主序列，按单击次序返回每步的效果。
    每项：{'in','out','all': set(spid), 'ontap': 是否“单击时”触发,
           'fade': 淡化/淡出效果数, 'nonfade': 其他效果数, 'other': [(效果名, spid)]}"""
    tm = slide._element.find('./p:timing', NS)
    if tm is None:
        return []
    seq = next((s for s in tm.iter('{%s}seq' % NS['p'])
                if s.get('nodeType') in (None, 'mainSeq')), None)
    if seq is None:
        return []
    lst = seq.find('./p:cTn/p:childTnLst', NS)
    if lst is None:
        return []
    steps = []
    for par in lst.findall('./p:par', NS):
        step = {'in': set(), 'out': set(), 'all': set(), 'ontap': False,
                'fade': 0, 'nonfade': 0, 'other': []}
        ctn = par.find('./p:cTn', NS)
        if ctn is not None:
            for cond in ctn.findall('./p:stCondLst/p:cond', NS):
                if cond.get('delay') == 'indefinite':
                    step['ontap'] = True
        for el in par.iter():
            t = _tag(el)
            spid = _spid_of(el) if t in (('animEffect', 'set') + _MOVE_TAGS) else None
            if t == 'animEffect':
                out_dir = (el.get('transition') or 'in') == 'out'
                filt = (el.get('filter') or '')
                if 'fade' in filt.lower():
                    step['fade'] += 1
                else:
                    step['nonfade'] += 1
                    step['other'].append((filt or 'animEffect', spid))
                if spid:
                    (step['out'] if out_dir else step['in']).add(spid)
            elif t == 'set':
                val = el.find('./p:to/p:strVal', NS)
                v = val.get('val') if val is not None else None
                if spid and v in ('visible', 'hidden'):
                    (step['in'] if v == 'visible' else step['out']).add(spid)
            elif t in _MOVE_TAGS:
                step['nonfade'] += 1
                step['other'].append((t, spid))
            if spid:
                step['all'].add(spid)
        steps.append(step)
    return steps


def _targets(recs, spids):
    """把动画目标 spid 集合展开为记录索引集合：动画挂在组合上时，组内子元素同样被命中。"""
    ids = set()
    for s in spids or ():
        try:
            ids.add(int(s))
        except (TypeError, ValueError):
            pass
    return set(i for i, r in enumerate(recs) if ids & set(r['ids']))


def _find(recs, needle, exact=False, kind=None):
    """按去空白文本查找记录：exact=True 要求整体相等，否则按包含匹配。"""
    key = _norm(needle)
    out = []
    for r in recs:
        if not r['norm'] or (kind and r['kind'] != kind):
            continue
        if (r['norm'] == key) if exact else (key in r['norm']):
            out.append(r)
    return out


def _one(recs, needle, exact=False):
    """取首个匹配记录，找不到返回 None。"""
    hits = _find(recs, needle, exact)
    return hits[0] if hits else None


def _has_anim(slide):
    """该页是否存在任何动画行为。"""
    tm = slide._element.find('./p:timing', NS)
    if tm is None:
        return False
    for el in tm.iter():
        if _tag(el) in (('animEffect', 'set', 'cmd') + _MOVE_TAGS):
            return True
    return False


def _has_transition(slide):
    """该页是否设置了切换效果；重新保存产生的空 <p:transition/> 标记不计。"""
    tr = slide._element.find('./p:transition', NS)
    if tr is None:
        return False
    if len(tr) > 0:
        return True
    return any(k.rsplit('}', 1)[-1] not in ('spd', 'dur', 'advClick', 'advTm')
               for k in tr.attrib)


def _slide_sig(slide):
    """页面内容签名：(去空白正文 md5, p:sp 计数, p:pic 计数)；不比对组合层级。"""
    el = slide._element
    txt = _norm(''.join(t.text or '' for t in el.findall('.//a:t', NS)))
    return (hashlib.md5(txt.encode('utf-8')).hexdigest(),
            len(el.findall('.//p:sp', NS)), len(el.findall('.//p:pic', NS)))


def _near(a, b, tol=_BOX_TOL):
    return a is not None and abs(a - b) <= tol


def _fmt_ok(r, fmt=None, algn=None):
    """校验记录内所有非空 run 的 (字号, 加粗, 字色) 与首段对齐方式是否等于基线。"""
    runs = [x for x in r['runs'] if (x['text'] or '').strip()]
    if not runs:
        return False
    if fmt is not None:
        sz, b, color = fmt
        for x in runs:
            if x['sz'] != sz or bool(x['b']) != bool(b) or x['color'] != color:
                return False
    if algn is not None and r['algn'] != algn:
        return False
    return True


def _paras_of(r):
    """记录内去空白后的非空段落文本列表。"""
    return [_norm(p) for p in r['paras'] if _norm(p)]


def _analyze(prs, slides):
    """一次性提取后续 50 条细则共用的版面事实与单击序列事实。"""
    s1 = slides[0]
    recs = _recs(s1)
    steps = _click_steps(s1)
    animated = set()
    for st in steps:
        animated |= _targets(recs, st['all'])

    # 六组材料：以“带引号的引文”所在的淡入步为锚点，把该步的目标集合视为整组构件
    units = {}
    for person, quote in _QUOTES.items():
        qkey = _norm(quote)
        qi = next((i for i, r in enumerate(recs) if qkey in r['norm']), None)
        if qi is None:
            continue
        hit = next(((si, _targets(recs, st['in'])) for si, st in enumerate(steps)
                    if st['fade'] > 0 and qi in _targets(recs, st['in'])), None)
        if hit is None:
            top = recs[qi]['ids'][0]
            units[person] = {'in_step': None, 'out_step': None,
                             'idx': set(i for i, r in enumerate(recs) if r['ids'][0] == top)}
            continue
        si, tin = hit
        oj = next((k for k in range(si + 1, len(steps))
                   if _targets(recs, steps[k]['out']) & tin), None)
        units[person] = {'in_step': si, 'out_step': oj, 'idx': tin}
    mat = set()
    for u in units.values():
        mat |= u['idx']

    ctx = {'prs': prs, 'slides': slides, 's1': s1, 'recs': recs, 'steps': steps,
           'animated': animated, 'units': units, 'mat': mat,
           'has_click': len(steps) >= 2 and any(st['ontap'] for st in steps)}
    _analyze_layout(ctx)
    return ctx


def _analyze_layout(ctx):
    """在第1页定位各基线构件（材料构件除外），结果写回 ctx。"""
    recs, mat = ctx['recs'], ctx['mat']
    free = [(i, r) for i, r in enumerate(recs) if i not in mat]
    ctx['free'] = free

    def pick(pred):
        return [i for i, r in free if pred(r)]

    def first(needle):
        key = _norm(needle)
        return next((i for i, r in free if r['norm'] == key), None)

    px0, py0 = _PANEL_BOX[0], _PANEL_BOX[1]
    px1, py1 = px0 + _PANEL_BOX[2], py0 + _PANEL_BOX[3]

    def in_panel(r):
        """记录是否落在“思想脉络”版块基线区域内（留 0.3 英寸余量）。"""
        if r['l'] is None or r['t'] is None:
            return False
        return (r['l'] >= px0 - 0.3 and r['t'] >= py0 - 0.3
                and r['l'] + r['w'] <= px1 + 0.3 and r['t'] + r['h'] <= py1 + 0.3)

    def in_panel_first(needle):
        key = _norm(needle)
        return next((i for i, r in free if r['norm'] == key and in_panel(r)), None)

    ctx['in_panel'] = in_panel
    # —— 七个人物卡片 ——
    ctx['card_pics'] = pick(lambda r: r['kind'] == 'pic')
    ctx['card_portraits'] = pick(
        lambda r: r['kind'] == 'pic' and _near(r['w'], _CARD_PORTRAIT[0])
        and _near(r['h'], _CARD_PORTRAIT[1]))
    ctx['card_frames'] = pick(
        lambda r: r['fill'] == _CARD_FRAME[2] and r['prst'] == 'rect'
        and r['lw'] == _LINE_W and r['lc'] == _CARD_FRAME[2]
        and _near(r['w'], _CARD_FRAME[0]) and _near(r['h'], _CARD_FRAME[1]))
    ctx['card_bars'] = pick(
        lambda r: r['fill'] == _CARD_BAR[2] and _near(r['w'], _CARD_BAR[0])
        and _near(r['h'], _CARD_BAR[1]))
    ctx['card_names'], ctx['claims'], ctx['claims_fmt'] = {}, {}, {}
    ctx['name_order'] = []
    for i, r in sorted(free, key=lambda t: ((t[1]['t'] or 0.0), (t[1]['l'] or 0.0))):
        for p in _PERSONS:
            if r['norm'] == _norm(p):
                if p not in ctx['name_order']:
                    ctx['name_order'].append(p)
                if p not in ctx['card_names'] and _fmt_ok(r, _CARD_NAME_FMT, 'ctr'):
                    ctx['card_names'][p] = i
            if r['norm'] == _norm(_CLAIMS[p]) and p not in ctx['claims']:
                ctx['claims'][p] = i
                if _fmt_ok(r, _CARD_CLAIM_FMT, 'ctr'):
                    ctx['claims_fmt'][p] = i
    _analyze_panel(ctx, free, pick, first, in_panel_first)


def _analyze_panel(ctx, free, pick, first, in_panel_first):
    """定位“思想脉络”版块、标题组、金色横线与底部标签，结果写回 ctx。"""
    recs = ctx['recs']
    ctx['panel'] = next((i for i in pick(
        lambda r: r['fill'] == _PANEL_FILL and r['lc'] == _PANEL_LINE
        and r['prst'] == 'roundRect' and _near(r['w'], _PANEL_BOX[2])
        and _near(r['h'], _PANEL_BOX[3]))), None)
    ctx['panel_label'] = first(_PANEL_LABEL)
    ctx['squares'] = pick(
        lambda r: r['prst'] == 'rect' and r['fill'] in _SQUARE_COLORS
        and _near(r['w'], _SQUARE_SIDE, 0.02) and _near(r['h'], _SQUARE_SIDE, 0.02))
    ctx['periods'] = [in_panel_first(t) for t in _PERIODS]
    ctx['person_cols'] = [in_panel_first(t) for t in _PERSON_COLS]
    ctx['rows'] = [in_panel_first(t) for t in _PANEL_ROWS]
    # —— 标题组与四角文字 ——
    ctx['title'] = first(_TITLE)
    ctx['subtitle'] = first(_SUBTITLE)
    ctx['corner_tl'] = first(_CORNER_TL)
    ctx['corner_br'] = first(_CORNER_BR)
    # —— 上下两条金色横线 ——
    ctx['gold_lines'] = pick(
        lambda r: r['prst'] == 'line' and r['nofill'] and r['lw'] == _LINE_W
        and r['lc'] == _PANEL_LINE and _near(r['w'], _GOLD_LINE_W, 0.1))
    # —— 底部四个标签 ——
    ctx['tag_plates'] = pick(
        lambda r: r['fill'] == _TAG_FILL and r['lc'] == _TAG_LINE
        and r['prst'] == 'roundRect' and _near(r['w'], _TAG_BOX[0])
        and _near(r['h'], _TAG_BOX[1]))
    ctx['tag_texts'] = [first(t) for t in _TAGS]
    # —— 铺满页面的不透明色块（多余的整页底板）——
    page = _PAGE_W_IN * _PAGE_H_IN
    ctx['full_page'] = pick(
        lambda r: r['kind'] == 'sp' and _opaque(r) and _area(_rect(r)) >= 0.9 * page)
    # —— 版块相关记录集合：用于判断说明文字是否被压盖 ——
    idx = [ctx['panel'], ctx['panel_label']] + ctx['periods'] + ctx['person_cols'] + ctx['rows']
    ctx['panel_idx'] = set(i for i in idx if i is not None) | set(ctx['squares'])
    ctx['row_recs'] = [recs[i] for i in ctx['rows'] if i is not None]


# —— 需要“材料淡入→淡出→主张淡入”三步的六位人物（与细则条目顺序一致）——
_MAT_PERSONS = ['孟德斯鸠', '伏尔泰', '卢梭', '休谟', '亚当·斯密', '康德']


def _emit(results, delta, rule, hit):
    """按 008 号脚本的结果串格式追加一条评分项：命中记实际增减分，未命中记 0。"""
    results.append('%s%d：%s' % ('+' if delta > 0 else '-',
                                abs(delta) if hit else 0, rule))


def _fade_sets(ctx):
    """逐步给出“淡化型”进入/退出效果命中的记录索引集合；非淡化步返回空集。"""
    recs = ctx['recs']
    fin, fout = [], []
    for st in ctx['steps']:
        f = st['fade'] > 0
        fin.append(_targets(recs, st['in']) if f else set())
        fout.append(_targets(recs, st['out']) if f else set())
    return fin, fout


def _blob(recs, idx):
    """把一组记录的去空白文本拼成一个串，用于“是否包含某段基线文字”的比对。"""
    return ''.join(recs[i]['norm'] for i in sorted(idx))


def _in_page(rc):
    """矩形是否完整落在页面可见区域内（留 _OOB_TOL_IN 容差）。"""
    if rc is None:
        return True
    return (rc[0] >= -_OOB_TOL_IN and rc[1] >= -_OOB_TOL_IN
            and rc[2] <= _PAGE_W_IN + _OOB_TOL_IN
            and rc[3] <= _PAGE_H_IN + _OOB_TOL_IN)


def _overlap(a, b):
    """两个矩形的相交面积（英寸²）；用于文字串行判断。"""
    if a is None or b is None:
        return 0.0
    return _area((max(a[0], b[0]), max(a[1], b[1]), min(a[2], b[2]), min(a[3], b[3])))


def _add_rules(ctx, results):
    """按细则加分项 A01–A22 的顺序逐条判定，结果写入 results。"""
    recs, steps, units = ctx['recs'], ctx['steps'], ctx['units']
    animated, click = ctx['animated'], ctx['has_click']
    fin, fout = _fade_sets(ctx)
    fade_in_all = set()
    for s in fin:
        fade_in_all |= s
    rows = [i for i in ctx['rows'] if i is not None]
    panel_ok = (ctx['panel'] is not None and ctx['panel_label'] is not None
                and len(rows) == 3)
    pics, claims = ctx['card_pics'], ctx['claims']

    # A01 —— 思想脉络版块完整且无任何动画（前提：已添加分步单击动画）
    _emit(results, 3,
          '“思想脉络”浅米色版块（#F6EFE0 填充、#D7B46A 描边）及其下三行内容完整存在，'
          '且未被设置任何进入/退出/消失类动画（前提：第1页已添加分步单击动画）',
          click and panel_ok and not (ctx['panel_idx'] & animated))

    # A02 —— 三行内容全程可见，未被遮挡
    cov_ok = panel_ok and not ctx['full_page']
    if cov_ok:
        for i in rows:
            covers = [_rect(recs[j]) for j in range(len(recs))
                      if recs[j]['z'] > recs[i]['z'] and _opaque(recs[j])]
            if _cover_ratio(_est_rect(recs[i]), covers) > _COVER_TOL:
                cov_ok = False
                break
    _emit(results, 5,
          '整套单击过程中该版块三行内容始终保持可见，未被整页遮挡层覆盖、'
          '也未被淡入的材料卡片压盖',
          click and cov_ok)

    # A03 —— 七张肖像均有淡化进入
    _emit(results, 3,
          '洛克、孟德斯鸠、伏尔泰、卢梭、休谟、亚当·斯密、康德七张金框肖像均设置了'
          '淡化型进入动画，放映起始状态均不显示',
          click and len(pics) >= 7 and all(i in fade_in_all for i in pics))

    # A04 —— 七条竖线短句均有淡化进入
    _emit(results, 3,
          '“自然权利｜政府同意”等七条竖线主张短句均设置了淡化型进入动画，'
          '放映起始状态均不显示',
          click and len(claims) == 7
          and all(i in fade_in_all for i in claims.values()))

    # A05 —— 进入动画触发方式均为“单击时”
    entr = set(pics) | set(claims.values())
    entr_steps = [k for k, s in enumerate(fin) if s & entr]
    _emit(results, 1,
          '上述肖像与主张短句的进入动画触发方式均为“单击时”，而非随上一动画自动播放',
          bool(entr_steps) and all(steps[k]['ontap'] for k in entr_steps)
          and len(pics) >= 7 and len(claims) == 7)

    # A06 —— 第1次单击淡入洛克
    lk = claims.get('洛克')
    _emit(results, 3,
          '第1次单击即让洛克的肖像与“自然权利｜政府同意”淡入显示，作为整页动画的起点',
          lk is not None and bool(fin) and lk in fin[0])

    # A07 —— 洛克仅占 1 次单击、无消失步骤
    lk_ids = set()
    if lk is not None:
        lk_ids.add(lk)
        for i in pics:
            if (recs[i]['l'] is not None and recs[lk]['l'] is not None
                    and abs(recs[i]['l'] - recs[lk]['l']) < 0.3):
                lk_ids.add(i)
    lk_steps = [k for k, st in enumerate(steps) if _targets(recs, st['all']) & lk_ids]
    _emit(results, 1,
          '洛克不设置“肖像＋引文材料”的显示与消失步骤，整页动画中仅占 1 次单击',
          lk is not None and len(lk_steps) == 1
          and not any(s & lk_ids for s in fout))

    # A08–A13 —— 六位人物各自的三步顺序
    for person in _MAT_PERSONS:
        u, ci = units.get(person), claims.get(person)
        ok = False
        if u is not None and ci is not None and u['in_step'] is not None:
            j = u['in_step']
            ok = (u['out_step'] == j + 1 and j + 2 < len(steps)
                  and ci in fin[j + 2]
                  and all(steps[k]['nonfade'] == 0 and steps[k]['fade'] > 0
                          for k in (j, j + 1, j + 2)))
        _emit(results, 5,
              '%s按“单击1→肖像与引文材料淡入显示；单击2→该材料淡出消失；'
              '单击3→“%s”淡入”的三步顺序设置' % (person, _CLAIMS[person]), ok)

    # A14 —— 总步数 19 且顺序无交叉
    order_ok = (len(steps) == 19 and lk is not None and bool(fin) and lk in fin[0])
    if order_ok:
        for k, person in enumerate(_MAT_PERSONS, 1):
            u, ci = units.get(person), claims.get(person)
            if (u is None or ci is None or u['in_step'] != 3 * k - 2
                    or u['out_step'] != 3 * k - 1 or ci not in fin[3 * k]):
                order_ok = False
                break
    _emit(results, 3,
          '全页单击总步数为 19 步（洛克 1 步＋六人各 3 步），六组材料的淡入、淡出与'
          '主张淡入共 18 步顺序无交叉、无错位、无缺步或多步',
          order_ok)

    # A15 —— 六组材料均在第1页内实现
    xml1 = ctx['s1']._element.xml
    _emit(results, 3,
          '六组引文材料均在第1页内直接淡入显示，未新增幻灯片、未跳转到其他页面，'
          '也未借助超链接或换页切换实现',
          len(units) == 6 and len(ctx['slides']) == 4
          and 'hlinkClick' not in xml1 and 'hlinkHover' not in xml1)

    # A16 —— 每组材料构成完整
    full_ok = len(units) == 6
    for person in _MAT_PERSONS:
        u = units.get(person)
        if u is None:
            full_ok = False
            break
        b = _blob(recs, u['idx'])
        if not any(recs[i]['kind'] == 'pic' for i in u['idx']):
            full_ok = False
        if not any(recs[i]['fill'] == _CARD_BAR[2] for i in u['idx']):
            full_ok = False
        for need in ([person, _IDENTITY[person], '观点提炼', _QUOTES[person]]
                     + _KEYPOINTS[person]):
            if _norm(need) not in b:
                full_ok = False
        if not full_ok:
            break
    _emit(results, 3,
          '每组材料构成完整——含该人物金框肖像、黑底姓名条、身份说明、“观点提炼”标签、'
          '要点文字与带引号的引文',
          full_ok)

    # A17 —— 材料与人物一一对应
    pair_ok = len(units) == 6
    for person in _MAT_PERSONS:
        u = units.get(person)
        if u is None:
            pair_ok = False
            break
        b = _blob(recs, u['idx'])
        if _norm(_QUOTES[person]) not in b or _norm(_IDENTITY[person]) not in b:
            pair_ok = False
            break
        for other in _MAT_PERSONS:
            if other != person and (_norm(_QUOTES[other]) in b
                                    or _norm(_IDENTITY[other]) in b):
                pair_ok = False
                break
        if not pair_ok:
            break
    _emit(results, 3,
          '六组材料与人物一一对应无错配，引文内容、身份说明均与本人相符',
          pair_ok)

    # A18 —— 材料区域洁净、文字不串行、不出界
    ptexts = [recs[i] for i in ([ctx['panel_label']] + ctx['periods']
                                + ctx['person_cols'] + ctx['rows']) if i is not None]
    clean = panel_ok and len(units) == 6
    for u in (units.values() if clean else []):
        for i in u['idx']:
            r = recs[i]
            if not _in_page(_rect(r)):
                clean = False
                break
            if not r['norm']:
                continue
            a = _est_rect(r)
            if any(_overlap(a, _est_rect(pr)) > 0.001 for pr in ptexts):
                clean = False
                break
        if not clean:
            break
    _emit(results, 3,
          '材料所在区域背景洁净、文字清晰可读，材料文字与下层版块的时期标签、'
          '人物栏、说明文字无重叠串行',
          clean)

    # A19 —— 每组材料整体一次淡入、一次淡出
    once_ok = bool(units)
    for u in (units.values() if once_ok else []):
        if (sum(1 for s in fin if s & u['idx']) != 1
                or sum(1 for s in fout if s & u['idx']) != 1):
            once_ok = False
            break
    _emit(results, 1,
          '每组材料的全部构件被组合为一个整体或统一编入同一次单击，'
          '一次淡入/淡出即完整出现与完整消失',
          once_ok)

    # A20 —— 效果统一为淡化
    _emit(results, 1,
          '全部进入与退出效果统一使用“淡化/淡出”，未混用飞入、缩放、擦除、轮子等其他效果',
          bool(steps) and all(st['nonfade'] == 0 for st in steps)
          and sum(st['fade'] for st in steps) > 0)

    # A21 —— 淡出后无残留
    residue_ok = (panel_ok and bool(units)
                  and not (ctx['panel_idx'] & animated))
    for u in (units.values() if residue_ok else []):
        got_in, got_out = set(), set()
        for k in range(len(steps)):
            got_in |= (fin[k] & u['idx'])
            got_out |= (fout[k] & u['idx'])
        if not got_in or not (got_in <= got_out):
            residue_ok = False
            break
    _emit(results, 1,
          '材料淡出后不残留半透明影像、空白底板或占位边框，页面回到七人卡片与'
          '“思想脉络”版块同时可见的状态',
          residue_ok)

    # A22 —— 动画仅加在第1页，素材三页不变（前提：已添加分步单击动画）
    scope_ok = click and len(ctx['slides']) >= 4
    if scope_ok:
        for n in (2, 3, 4):
            sl = ctx['slides'][n - 1]
            if (_slide_sig(sl) != _SRC_SIG[n] or _has_anim(sl)
                    or _has_transition(sl)):
                scope_ok = False
                break
    _emit(results, 3,
          '动画仅添加在第1页；素材三页的内容、版式与无动画状态保持不变'
          '（前提：第1页已添加分步单击动画）',
          scope_ok)


def _ded_rules(ctx, results):
    """按细则减分项 D01–D28 的顺序逐条判定，结果写入 results。"""
    recs, units = ctx['recs'], ctx['units']
    prs, slides = ctx['prs'], ctx['slides']
    rows = [i for i in ctx['rows'] if i is not None]
    xml1 = ctx['s1']._element.xml
    allnorm = set(r['norm'] for r in recs if r['norm'])

    # D01 —— 七张肖像缺失
    _emit(results, -3,
          '洛克、孟德斯鸠、伏尔泰、卢梭、休谟、亚当·斯密、康德七张肖像图片存在缺失、'
          '被删除，或被替换为无图形的文字单元',
          len(ctx['card_portraits']) < 7)

    # D02 —— 卡片版式被改写
    _emit(results, -3,
          '七个人物卡片的版式被改写（应为 1.63×1.95 英寸 #D7B46A 金色边框、内嵌 '
          '1.56×1.88 英寸肖像、下方黑色底条内 8.7 磅白色加粗居中姓名）',
          not (len(ctx['card_frames']) >= 7 and len(ctx['card_bars']) >= 7
               and len(ctx['card_names']) == 7))

    # D03 —— 竖线短句被改写或格式变化
    _emit(results, -3,
          '七条竖线主张短句被改写为其他表述，或不满足 7.2 磅、居中、颜色 #F7F3E8',
          len(ctx['claims_fmt']) < 7)

    # D04 —— 人物顺序或姓名用字被改动
    _emit(results, -1,
          '人物排列顺序或姓名用字被改动（应为洛克、孟德斯鸠、伏尔泰、卢梭、休谟、'
          '亚当·斯密、康德依次横排，“亚当·斯密”含间隔号）',
          ctx['name_order'] != _PERSONS)

    # D05 —— 版块被删除或改为表格式版面
    _emit(results, -3,
          '“思想脉络”版块被删除，或被改为“时间／国家／代表人物／主要主张”之类的'
          '表格式版面',
          ctx['panel'] is None or ctx['panel_label'] is None or len(rows) < 3
          or '<a:tbl' in xml1)

    # D06 —— 标签格式
    _emit(results, -1,
          '“思想脉络”标签文字不满足 10.5 磅加粗、颜色 #A65A45',
          ctx['panel_label'] is None
          or not _fmt_ok(recs[ctx['panel_label']], _PANEL_LABEL_FMT))

    # D07 —— 三个小方块
    sq = sorted((recs[i] for i in ctx['squares']), key=lambda r: (r['t'] or 0.0))
    _emit(results, -1,
          '三行行首的三个 0.18×0.18 英寸小方块（自上而下青色 #2AA6A5、金色 #D7B46A、'
          '赭红 #A65A45）缺失或被改色',
          len(sq) < 3 or [r['fill'] for r in sq[:3]] != _SQUARE_COLORS)

    # D08 —— 时期标签与人物栏
    _emit(results, -1,
          '三个时期标签“17世纪英国”“18世纪法国”“18世纪后期”（8.6 磅加粗 #163038）'
          '与三个人物栏（8.6 磅加粗 #A65A45）的文字或格式被改动',
          any(i is None or not _fmt_ok(recs[i], _PERIOD_FMT)
              for i in ctx['periods'])
          or any(i is None or not _fmt_ok(recs[i], _PERSON_COL_FMT)
                 for i in ctx['person_cols']))

    # D09 —— 三行说明文字
    _emit(results, -1,
          '三行说明文字被改写，或不满足 7.9 磅、颜色 #163038',
          any(i is None or not _fmt_ok(recs[i], _PANEL_ROW_FMT)
              for i in ctx['rows']))

    # D10 —— 新增分栏栏目文字
    allow = set(_norm(t) for t in ([_PANEL_LABEL] + _PERIODS + _PERSON_COLS
                                   + _PANEL_ROWS))
    tabwords = set(_norm(w) for w in _TABLE_WORDS)
    extra = [i for i, r in ctx['free']
             if r['norm'] and ctx['in_panel'](r) and r['norm'] not in allow]
    extra += [i for i, r in ctx['free'] if r['norm'] in tabwords]
    _emit(results, -1,
          '为该版块新增本不存在的栏目文字（如“时间”“国家”“代表人物”“主要主张”'
          '等分栏表述）',
          bool(extra))

    # D11 —— 主标题
    _emit(results, -1,
          '页面主标题文字不是“四、启蒙运动——理性主义的觉醒”，或不满足 19 磅加粗、'
          '颜色 #F7F3E8',
          ctx['title'] is None or not _fmt_ok(recs[ctx['title']], _TITLE_FMT))

    # D12 —— 副标题文字被改写
    _emit(results, -3,
          '副标题文字不是“代表人物及思想主张：从“权利”到“法治”的思想谱系”'
          '（被改写为操作提示类文字即计）',
          ctx['subtitle'] is None)

    # D13 —— 副标题格式
    _emit(results, -1,
          '副标题不满足 8.8 磅加粗、颜色 #D7B46A',
          ctx['subtitle'] is None
          or not _fmt_ok(recs[ctx['subtitle']], _SUBTITLE_FMT))

    # D14 —— 左上角角标
    _emit(results, -1,
          '左上角“理性主义的觉醒”不满足 8 磅加粗、颜色 #9CB7B8',
          ctx['corner_tl'] is None
          or not _fmt_ok(recs[ctx['corner_tl']], _CORNER_TL_FMT))

    # D15 —— 右下角角标
    _emit(results, -1,
          '右下角“第8课  欧洲的思想解放运动”不满足 7.5 磅、右对齐、颜色 #9CB7B8',
          ctx['corner_br'] is None
          or not _fmt_ok(recs[ctx['corner_br']], _CORNER_BR_FMT, 'r'))

    # D16 —— 上下两条金色横线
    _emit(results, -1,
          '页面上下两条金色横线（纵向 0.18 与 7.38 英寸、线宽 1 磅、颜色 #D7B46A 的'
          '无填充线条）缺失，或被改为填充色块',
          len(ctx['gold_lines']) < 2)

    # D17 —— 底部四个标签底板
    _emit(results, -3,
          '底部四个 2.87×0.50 英寸深青底标签（#0F3C45 填充、#2AA6A5 描边）'
          '“理性至上”“天赋人权”“批判专制与教权”“民主与法治”缺失或被删除',
          len(ctx['tag_plates']) < 4 or any(i is None for i in ctx['tag_texts']))

    # D18 —— 标签文字格式
    _emit(results, -1,
          '四个标签文字不满足 8.2 磅加粗、居中、颜色 #F7F3E8',
          any(i is None or not _fmt_ok(recs[i], _TAG_FMT, 'ctr')
              for i in ctx['tag_texts']))

    # D19 —— 东亚字体
    bad_ea = False
    for r in recs:
        for x in r['runs']:
            if (x['text'] or '').strip() and (x['ea'] or '') not in _EA_OK:
                bad_ea = True
                break
        if bad_ea:
            break
    _emit(results, -1,
          '页面文字的中文（东亚）字体设定丢失，或字体名称被由 Microsoft YaHei '
          '改为其他名称',
          bad_ea)

    # D20 —— 新增整页色块
    _emit(results, -1,
          '新增覆盖整页的纯色矩形底板，改变页面底色',
          bool(ctx['full_page']))

    # D21 —— 页面尺寸与母版
    try:
        pw = prs.slide_width / EMU_IN
        ph = prs.slide_height / EMU_IN
    except Exception:
        pw = ph = 0.0
    lays = set()
    for sl in slides:
        try:
            lays.add((sl.slide_layout.name, sl.slide_layout.slide_master.name))
        except Exception:
            lays.add(None)
    _emit(results, -1,
          '幻灯片尺寸不再是 13.33×7.5 英寸的宽屏规格（宽高偏差超过 0.02 英寸即计），'
          '或所用母版、版式被替换',
          abs(pw - _PAGE_W_IN) > _PAGE_TOL_IN or abs(ph - _PAGE_H_IN) > _PAGE_TOL_IN
          or len(lays) > 1)

    # D22 —— 新增操作说明类文字
    _emit(results, -1,
          '新增指令未要求的操作说明类文字（如“单击顺序：……”“主要主张（内容初始隐藏）”'
          '“材料阅读：×××”等）',
          any(m in r['norm'] for r in recs for m in _INSTR_MARKS))

    # D23 —— 要点文字被加符号或重新分段
    kp_bad = False
    for person, u in units.items():
        best, n = None, 0
        for i in u['idx']:
            ps = _paras_of(recs[i])
            if len(ps) > n:
                best, n = ps, len(ps)
        if n < 2 or best != [_norm(x) for x in _KEYPOINTS.get(person, [])]:
            kp_bad = True
            break
    _emit(results, -1,
          '材料要点文字被加上“•”等项目符号，或被重新分段、改写措辞',
          kp_bad)

    # D24 —— 身份说明错配
    id_bad = False
    for person, u in units.items():
        b = _blob(recs, u['idx'])
        if (_norm(_IDENTITY.get(person, '')) not in b
                or any(o != person and _norm(_IDENTITY[o]) in b for o in _IDENTITY)):
            id_bad = True
            break
    _emit(results, -1, '材料中的人物身份说明与人物不符', id_bad)

    # D25 —— 材料出界
    oob = False
    for u in units.values():
        if any(not _in_page(_rect(recs[i])) for i in u['idx']):
            oob = True
            break
    _emit(results, -1,
          '材料肖像或文本框超出页面边界，部分内容落在可见区域之外',
          oob)

    # D26 —— 素材三页被改动
    src_bad = len(slides) < 4
    if not src_bad:
        for n in (2, 3, 4):
            if _slide_sig(slides[n - 1]) != _SRC_SIG[n]:
                src_bad = True
                break
    _emit(results, -3,
          '“法国中心：理性批判与制度设计”“从人民主权到经验理性”“向欧洲其他地区扩展”'
          '三页被删除、被清空或被改写内容',
          src_bad)

    # D27 —— 第1页已有对象被删除
    baseline = (_PERSONS + [_CLAIMS[p] for p in _PERSONS] + [_PANEL_LABEL]
                + _PERIODS + _PERSON_COLS + _PANEL_ROWS
                + [_TITLE, _SUBTITLE, _CORNER_TL, _CORNER_BR] + _TAGS)
    _emit(results, -3,
          '删除第1页中已有的图片、图形或文本框，导致原有内容缺失',
          any(_norm(t) not in allnorm for t in baseline)
          or len(ctx['card_portraits']) < 7 or len(ctx['gold_lines']) < 2
          or len(ctx['tag_plates']) < 4)

    # D28 —— 其他页面新增动画或切换
    _emit(results, -1,
          '为除第1页以外的页面添加了指令未要求的动画或页面切换效果',
          any(_has_anim(sl) or _has_transition(sl) for sl in slides[1:]))


# 交付物只接受 .pptx：细则要求交付「1 个可正常打开的 PowerPoint 演示文稿（.pptx）」，
# 其余演示文稿/导出格式（.ppt / .pptm / .odp / .pdf 等）一律在维度一判不通过。
_ACCEPTED_EXT = '.pptx'
# 目录里没有 .pptx 时，这些扩展名仍收作被评对象，交给维度一判不通过（总分记 0），
# 而不是当成「没有交付物」直接跳过评测。
_WRONG_FMT_EXTS = ('.ppt', '.pptm', '.ppsx', '.ppsm', '.pps', '.potx', '.pot',
                   '.odp', '.key', '.pdf')


def _fmt_reject(file_path):
    """维度1的格式门禁：交付物必须是可正常打开的 .pptx（OOXML 演示文稿包）。

    不通过返回原因文本，通过返回 None。判定与 PowerPoint 的“能不能双击打开”对齐：
    扩展名须为 .pptx，文件须是 zip 容器，且包内含 ppt/presentation.xml。
    """
    ext = os.path.splitext(file_path)[1].lower()
    if ext != _ACCEPTED_EXT:
        return '交付物为 %s 格式，只接受 .pptx 演示文稿' % (ext or '无扩展名')
    if not zipfile.is_zipfile(file_path):
        return '交付物不是有效的 .pptx 包（无法作为 zip 容器打开）'
    try:
        with zipfile.ZipFile(file_path) as zf:
            names = set(zf.namelist())
    except Exception as exc:
        return '交付物无法作为 .pptx 打开：%s: %s' % (type(exc).__name__, exc)
    if 'ppt/presentation.xml' not in names:
        return '交付物包内缺少 ppt/presentation.xml，不是 PowerPoint 演示文稿'
    return None


def _dim1(prs, slides):
    """维度1：交付文件格式正确——可正常解析、第1页对象可独立编辑、版面未整体丢失。

    “只接受 .pptx”这道格式门禁由 _fmt_reject 在解析前先行判定，本函数只看内容。"""
    if not slides:
        return False, '文档内没有任何幻灯片'
    recs = _recs(slides[0])
    texts = [r for r in recs if r['norm']]
    if not texts:
        return False, '第1页没有任何可编辑文字对象（疑似被压平为图片或截图粘贴）'
    if not any(_norm(_TITLE) == r['norm'] or '启蒙运动' in r['norm'] for r in texts):
        return False, '第1页缺少可识别的标题文字，原有版面内容已丢失'
    return True, ''


def _evaluate_impl(file_path, result):
    """解析文件并按细则完成维度1、维度2的判定，回填 result 后返回。"""
    # 维度一的第一道门禁：交付物格式只接受 .pptx，不合格即判 0，不再解析内容
    reject = _fmt_reject(file_path)
    if reject:
        result['dim1_pass'] = False
        result['dim1_reason'] = '维度一未通过：%s，完成度总分记 0' % reject
        result['dim2_items'] = []
        result['total_score'] = 0
        return result

    prs = Presentation(file_path)
    slides = list(prs.slides)
    ok, reason = _dim1(prs, slides)
    result['dim1_pass'] = ok
    if not ok:
        result['dim1_reason'] = '维度一未通过：%s，完成度总分记 0' % reason
        result['dim2_items'] = []
        result['total_score'] = 0
        return result
    result['dim1_reason'] = reason

    ctx = _analyze(prs, slides)
    results = []
    _add_rules(ctx, results)
    _ded_rules(ctx, results)

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
        dim2_items.append({'rule': rule, 'max_delta': max_delta,
                           'delta': max_delta if hit else 0, 'hit': hit})
        total_score += (max_delta if hit else 0)
    result['dim2_items'] = dim2_items
    result['total_score'] = total_score
    result['max_score'] = MAX_SCORE_TOTAL
    return result


def evaluate(dir_path=None):
    """对本脚本所在目录内的输出 PPTX 评分，返回统一约定格式的结果字典。

    评估目标固定为 _SCRIPT_DIR，dir_path 参数仅为保持各脚本接口一致而保留，
    传入的值不会被使用；也不会进入子目录或检索上级目录。

    交付物只接受 .pptx：目录内优先取 .pptx；若只交了 .ppt / .pptm / .odp / .pdf 等
    其他格式，仍把它收作被评对象，由维度一判不通过（total_score 记 0、max_score 仍
    保留真实满分），而不是当成「没有交付物」跳过评测。

    返回固定 9 个字段：id / file_name / status / error / dim1_pass / dim1_reason /
    dim2_items / total_score / max_score，其中 dim2_items 每项为
    {rule, max_delta, delta, hit}。

    两类异常分开回传：
      · 目录不存在 / 目录内没有任何演示文稿——维度一根本没进入判定，error 只带纯文本
        原因（不带异常类名），file_name、dim1_pass、dim1_reason 一律留空（dim1_pass 为
        空字符串而非 False），dim2 不出项，总分与满分一并归零；
      · 其他运行期异常——error 以「脚本运行异常：」打头，dim1_pass 记 False 并在
        dim1_reason 中给出错误摘要，max_score 仍保留真实满分。
    """
    dir_path = dir_path if dir_path and os.path.isdir(dir_path) else _SCRIPT_DIR
    result = {
        'id': SCRIPT_ID,
        'file_name': '',
        'status': 'ok',
        'error': '',
        'dim1_pass': False,
        'dim1_reason': '',
        'dim2_items': [],
        'total_score': 0,
        'max_score': MAX_SCORE_TOTAL,
    }
    try:
        if not dir_path or not os.path.isdir(dir_path):
            raise FileNotFoundError('目录不存在或不是目录：%s' % dir_path)

        files = sorted(
            n for n in os.listdir(dir_path)
            if not n.startswith('~$') and os.path.isfile(os.path.join(dir_path, n)))
        cands = [n for n in files if n.lower().endswith(_ACCEPTED_EXT)]
        if not cands:
            # 没有 .pptx，但交了 .ppt/.pptm/.odp/.pdf 等其他格式：仍收作被评对象，
            # 由维度一判不通过，而不是当成「没有交付物」跳过评测。
            cands = [n for n in files if n.lower().endswith(_WRONG_FMT_EXTS)]
        if not cands:
            raise FileNotFoundError('目录中未找到 .pptx 文件：%s' % dir_path)

        file_name = cands[0]
        result['file_name'] = file_name
        return _evaluate_impl(os.path.join(dir_path, file_name), result)
    except FileNotFoundError as exc:
        # 找不到被评估文件：error 只回传纯文本原因，维度一未进入判定，各字段留空、分数归零。
        result['status'] = 'error'
        result['error'] = str(exc)
        result['file_name'] = ''
        result['dim1_pass'] = ''
        result['dim1_reason'] = ''
        result['dim2_items'] = []
        result['total_score'] = 0
        result['max_score'] = 0
        return result
    except Exception as exc:
        result['status'] = 'error'
        result['error'] = '脚本运行异常：%s: %s' % (type(exc).__name__, exc)
        result['dim1_pass'] = False
        result['dim1_reason'] = ('文件无法打开或解析：%s: %s'
                                 % (type(exc).__name__, exc))
        result['dim2_items'] = []
        result['total_score'] = 0
        return result


if __name__ == '__main__':
    # 本地调试入口：固定评估本脚本所在目录，不接受命令行参数。
    print(json.dumps(evaluate(), ensure_ascii=False, indent=2))

