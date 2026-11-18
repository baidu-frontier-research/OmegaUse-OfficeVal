#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# pyright: reportUnknownMemberType=false, reportUnknownVariableType=false, reportUnknownArgumentType=false, reportUnknownParameterType=false, reportMissingParameterType=false, reportPrivateUsage=false, reportUnknownLambdaType=false, reportGeneralTypeIssues=false

import os
import re
import sys
import json
import zipfile
import traceback
from docx import Document
from docx.oxml.ns import qn

# 说明：按接口统一约定，本模块不修改全局 sys.stdout，也不 print 主结果。
# 主结果统一由 evaluate(dir_path) -> dict 返回；入参为"脚本所在目录的路径"，
# 由脚本自己在该目录内定位并打开被评估的 .docx。本文件仅在 __main__ 下作为
# 本地调试入口，才把 dict 序列化为 JSON 打印。
#
# 本脚本对应任务：按"V2毕业设计论文格式要求"中的批注调整"V2成果说明书_A"格式、
# 在目录页重建可自动跳转的目录、并添加页面下方居中的页码（目录页及以前不编页码）。
# 维度一仅判"交付文件是否为1个可正常打开的 .docx"；维度二分为 21 条加分项（+）
# 与 13 条减分项（-）；加分项来自指令下达的任务，减分项来自指令未涉及却被改动、
# 破坏交付文档既有内容的多余修改。

SCRIPT_ID = "030"
# 维度二加分项 max_delta 之和；用于维度一未通过 / 脚本异常时兜底
MAX_SCORE_TOTAL = 69
# 维度二减分项 max_delta 之和（负值），仅用于自检
MIN_PENALTY_TOTAL = -35

# 维度二所有评分项按"结果追加顺序"的 max_delta 列表；末端把 results 里的
# "+N：..." / "-N：..." 字符串反解为结构化 dim2_items 时用。顺序须与
# _evaluate_impl 中 results.append(...) 的顺序严格一致；共 34 项。
_MAX_DELTAS_ORDER = [
    3,   # 01 A4 纵向 + 页边距 3.6/3.4/2.7/2.5cm
    3,   # 02 封面标题区三行字体字号
    3,   # 03 封面信息表六行：宋体四号 + 左缩进1400 + 固定25磅
    3,   # 04 独创性声明标题：黑体三号加粗居中 + 段前后 + 固定28磅
    3,   # 05 声明正文 + 学生签字行
    3,   # 06 "目 录"标题：黑体小三居中 + 空两字符 + 1.5倍行距
    3,   # 07 一级标题：黑体四号顶格 + 段前约18磅段后约12磅
    5,   # 08 正文段落：宋体小四 + 首行缩进2字符 + 固定20磅
    3,   # 09 图注：黑体五号 + 位于图片下方
    3,   # 10 "六、参考资料"标题：黑体小四居中
    3,   # 11 参考资料条目：宋体五号顶格
    5,   # 12 目录域（TOC 域 / SDT）——前置检查项
    3,   # 13 一级标题大纲级别1 + 目录条目与标题一一对应（受门槛约束）
    3,   # 14 目录条目可跳转
    3,   # 15 目录条目页码由 PAGEREF 域生成（受门槛约束）
    3,   # 16 目录条目字体：宋体小四 + 页码TNR + 加粗 + 空两字符
    3,   # 17 目录条目段后5磅 + 固定19磅
    3,   # 18 目录条目右对齐点线制表位（版心右边界）
    5,   # 19 分节实现页码控制 + 正文节起始页码为1
    3,   # 20 页脚 PAGE 域 + 水平居中
    3,   # 21 封面/声明/目录三页无页码
    -5,  # 22 文档整体结构（章节顺序/缺失）
    -3,  # 23 一级标题文字被改写
    -3,  # 24 图注文字被改写
    -3,  # 25 正文段落文字被改写
    -3,  # 26 插图数量或尺寸被改动
    -3,  # 27 模块功能表行列结构被改动
    -3,  # 28 封面信息表/声明/参考资料条目文字被改写
    -3,  # 29 被插入批注/修订/高亮/底纹/彩色字体
    -3,  # 30 被插入页眉/水印/分隔线/额外空白页/空段落堆叠换页
    -3,  # 31 图注被附加加粗
    -1,  # 32 封面副标题排版与既有配色被改动
    -1,  # 33 封面信息表标签列加粗右对齐被破坏
    -1,  # 34 目录页被插入额外说明文字
]
assert sum(d for d in _MAX_DELTAS_ORDER if d > 0) == MAX_SCORE_TOTAL
assert sum(d for d in _MAX_DELTAS_ORDER if d < 0) == MIN_PENALTY_TOTAL

# ---------------------------------------------------------------------------
# 交付文档必须保持不变的既有文字（用于减分项的"文字保真"判定）
# ---------------------------------------------------------------------------

# 中文字体族名（含常见等价西文名与 Linux 端等价字族），只做"族"判定不做精确串比
STRICT_SONGTI = ('宋体', 'SimSun', 'NSimSun', 'STSong', 'Songti SC',
                 'Noto Serif CJK SC', 'Source Han Serif SC', 'Source Han Serif CN')
STRICT_HEITI = ('黑体', 'SimHei', 'STHeiti', 'Heiti SC',
                'Noto Sans CJK SC', 'Source Han Sans SC', 'Source Han Sans CN')
STRICT_TNR = ('Times New Roman', 'TimesNewRoman', 'Times New Roman PS', 'Liberation Serif')

COVER_SCHOOL = '江海应用技术学院'
COVER_TITLE = '毕业设计成果说明书'
COVER_KIND = '（全日制专科生）'
COVER_SUBTITLE = '结构方案、控制布置与使用性能说明'
COVER_BLUEGRAY = '36475C'
COVER_SUBTITLE_GRAY = '646E78'
# 封面"（全日制专科生）"一行既有的灰色，属交付文档原有配色，不计入"彩色字体"
COVER_KIND_GRAY = '4B555F'

# 封面信息表六行（标签, 取值）
COVER_TABLE_ROWS = [
    ('毕业设计题目：', '折叠式仓储搬运车机械结构优化设计'),
    ('二 级 学 院：', '智能装备学院'),
    ('专 业 班 级：', '机电一体化技术 24-3'),
    ('学 生 学 号：', '24A0318426'),
    ('学 生 姓 名：', '林若航'),
    ('指 导 老 师：', '周雁宁'),
]

DECL_TITLE = '毕业设计独创性声明'
# 独创性声明两段正文（逐字保持）
DECL_BODY = [
    '本人郑重声明：所提交的毕业设计成果为本人在指导教师指导下完成的综合实践成果，'
    '文中所述设计过程、训练记录、图表和分析内容均根据项目资料整理形成。',
    '除文内已注明的参考资料外，本文未直接使用其他个人或集体已经发表、撰写或制作的成果。'
    '若后续发现成果内容存在学术不端或资料使用不当情形，本人愿意按学院相关规定承担相应责任。',
]
# 学生签字行：内含大段占位空格，比对时按"折叠空白后逐字一致 + 占位空格未被压缩"两步判定
DECL_SIGN_CORE = '学生签字： 日期： 年 月 日'
DECL_SIGN_MIN_LEN = 34

TOC_TITLE_KEY = '目录'

# 六个一级标题（含中文序号）；"六、参考资料"另有单独字号要求，故单列出来
HEADINGS = ['一、设计思路', '二、设计内容', '三、设计过程',
            '四、作品及特点', '五、致谢', '六、参考资料']
HEADING_KEYS = ['设计思路', '设计内容', '设计过程', '作品及特点', '致谢', '参考资料']
REF_HEADING = '六、参考资料'

# 七条图注（逐字保持，"图"与序号之间无空格）
CAPTIONS = [
    '图1 折叠式仓储搬运车外观结构',
    '图2 模块化底盘与副车架布局',
    '图3 机械转向机构',
    '图4 辅助动力与传动布置',
    '图5 集中式控制面板',
    '图6 悬架轮端结构',
    '图7 整机模块关系',
]

# 七张插图的 EMU 尺寸（cx, cy），按图1..图7 顺序
IMAGE_EXTENTS = [
    (4392000, 2781600), (4392000, 2781600), (4392000, 2781600), (4392000, 2781600),
    (4392000, 2781600), (4140000, 2622000), (4392000, 2781600),
]

# 模块功能表：表头三列 + 首列四行
TABLE_HEADER = ['模块', '主要功能', '设计要点']
TABLE_FIRST_COL = ['承载箱体', '副车架', '悬挂轮组', '控制面板']

# 四条参考资料条目（逐字保持，含中文逗号与句号）
REFERENCES = [
    '[1] 机械结构设计基础课程资料，智能装备类专业教学资料，2024。',
    '[2] 机电设备装配与维护训练手册，职业教育实训资料，2024。',
    '[3] 低速搬运设备安全使用要点，校园设备管理资料，2025。',
    '[4] 产品结构表达与三维建模实践指导，工程制图训练资料，2024。',
]

# 六章正文段落（逐段保持原有叙述文字），共 21 段
BODY_PARAGRAPHS = [
    '本设计面向校园维修、仓储转运和短距离物料搬运场景，目标是形成一款便于折叠收纳、操控稳定、能适应一般室外路面的轻型搬运车。方案重点从箱体轻量化、底盘缓冲、转向灵活性和低噪声助力四个方面展开。',
    '整体造型采用低重心箱体和外置把手的组合，使主要载荷集中在轮距范围内，降低转弯或越过小台阶时的侧倾风险。框架采用模块化思路，便于后期维护、拆装和更换零件，同时保留简洁的外观线条。',
    '在使用体验上，设计强调“轻推、好转、易收纳”。把手角度、前轮转向间隙和轮组接地宽度均以日常搬运习惯为依据进行调整，使操作者在走廊、库房门口和室外平整路面都能获得较稳定的反馈。',
    '设计内容包括承载箱体、底盘框架、轮端悬挂、机械转向、辅助动力、制动与控制面板六个部分。各部分之间通过螺栓连接与定位槽配合，既保证装配强度，也便于后期维修。',
    '承载箱体采用浅箱式轮廓，并在底部设置纵横加强筋。箱体开口边缘设置圆角翻边，以减少搬运时的磕碰风险。把手为可折叠结构，收纳后不明显增加整车长度，适合放置在工具间或车辆后备区域。',
    '底盘部分采用一体副车架承接轮端冲击，将轮组、悬挂和传动部件集中在下方模块。助力系统不追求高速度，而强调起步平稳、低噪声和低维护；控制面板集中显示电量、档位与助力状态，减少操作者的判断成本。',
    '第一阶段完成使用场景分析和结构分解。根据库房搬运、教学实训和户外短距离转运的需求，确定箱体尺寸以中小型工具箱、维修包和散件盒为主要载荷对象，并将整机重心控制在轮轴连线附近。',
    '第二阶段进行主体结构建模。承载箱体外壁保持简洁，底部采用加强筋网格分担集中载荷；电池仓、控制器和传动件的位置预留在副车架内部，使上部箱体尽量少开孔，避免削弱承载刚性。',
    '第三阶段优化转向机构。前轮采用机械连杆式转向，结构由转向柱、横拉杆、轮端转向节和限位件组成。该方案取消复杂助力机构，减少故障点，转弯半径满足室内通道和仓储货架间移动需求。',
    '转向连接处使用可调节拉杆，便于装配后校正前束。限位结构控制最大转角，防止急转时轮胎与箱体侧壁发生干涉。把手与转向柱通过折叠铰链连接，收纳状态下可贴近箱体侧面。',
    '第四阶段完成辅助动力布置。动力单元采取低速助力方案，电池组布置在底盘中部偏下位置，电机与减速单元靠近后轮，短传动路径能够降低能量损耗。该布置同时将质量集中在车体中央，使转向和停车更稳定。',
    '为避免控制复杂度过高，助力系统以低、中、高三个档位切换为主，并配置前进、后退和断电保护功能。操作者可根据坡道、地面阻力和载荷情况选择合适档位。',
    '第五阶段设计控制与制动系统。控制面板集中在把手下方，显示电量、档位、工作状态和故障提醒。按钮采用分区排布，减少误触；手刹与把手一体化，方便在坡道或临时停放时快速制动。',
    '制动结构采用手捏式机械—液压联动方案，轮端制动力优先满足低速搬运场景下的短距离停车需求。线束沿副车架内侧布置，并采用卡扣固定，避免在轮组运动时被拉扯。',
    '第六阶段校核轮端悬挂。悬挂结构以短行程缓冲为主，能够过滤普通路面缝隙、地砖落差和小坡面带来的冲击。轮端支架与下控制臂共同承受侧向力，减少箱体直接受力。',
    '悬挂设计不追求复杂调校，而侧重可靠性和易维护性。常用紧固件集中布置在外侧，日常检查时无需拆卸箱体即可观察连接件状态。',
    '作品的主要特点是结构清晰、重量分布合理、操作界面集中。承载箱体以浅箱式结构为核心，在不增加过多零件的情况下兼顾刚性与容量；底盘副车架将悬挂、动力和制动整合在下部模块，便于维护。',
    '轮组采用较宽胎面并配合短行程缓冲，适合校园路面、仓储通道和普通室外平整区域。相较于普通平板推车，该方案在起步助力、转向反馈和通过小障碍时更友好。',
    '控制系统的优势在于信息集中。电量、档位和运行状态一目了然，前进、后退、助力调节与断电开关分区明显，能够减少新操作者的学习成本。整车外观避免过度装饰，强调清爽、耐用和便于清洁。',
    '本次毕业设计从需求分析到结构方案完善，得到了指导教师和实训室同学的帮助。老师在设计定位、结构表达和装配逻辑方面给予了细致建议，使方案从初步构想到较完整的产品说明逐步成形。',
    '在资料整理和图形绘制过程中，我进一步理解了机械结构设计需要兼顾使用场景、加工装配和后期维护。今后将继续在机械制图、结构分析和设备调试方面积累经验。',
]

# 空段落基线：交付文档本身在封面/前置区域已有 7 个空段落，最长连续 2 个；
# 只有出现 >=5 个连续空段落时才判定为"用堆叠空段落实现换页"
EMPTY_RUN_LIMIT = 5

_A_NS = '{http://schemas.openxmlformats.org/drawingml/2006/main}'
_WP_NS = '{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}'
_R_ID = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id'


# ---------------------------------------------------------------------------
# 通用小工具
# ---------------------------------------------------------------------------

def _ptext(el):
    """取元素内所有 w:t 拼成的可见文字（不含域代码 instrText）。"""
    if el is None:
        return ''
    return ''.join(t.text or '' for t in el.iter(qn('w:t')))


def _norm(s):
    """去掉全部空白（含全角空格），用于"文字是否被改写"的逐字比对。"""
    return re.sub(r'[\s　]+', '', s or '')


def _collapse(s):
    """把连续空白折叠为单个半角空格，用于含占位空格的行的比对。"""
    return re.sub(r'[\s　]+', ' ', s or '').strip()


def _blocks(body):
    """body 的直接子块：段落 / 表格 / 结构化文档标记。"""
    return [ch for ch in body.iterchildren()
            if ch.tag in (qn('w:p'), qn('w:tbl'), qn('w:sdt'))]


def _own_runs(p_el):
    """只属于该段落自身的 run（排除嵌套文本框、嵌套表格里的 run）。"""
    out = []
    for r in p_el.iter(qn('w:r')):
        anc = r.getparent()
        while anc is not None and anc.tag != qn('w:p'):
            anc = anc.getparent()
        if anc is p_el:
            out.append(r)
    return out


def _visible_runs(p_el):
    return [r for r in _own_runs(p_el) if (_ptext(r) or '').strip()]


def _int_or_none(v):
    try:
        return int(round(float(v)))
    except (TypeError, ValueError):
        return None


def _font_in(name, family):
    if not name:
        return False
    low = str(name).strip().lower()
    return any(low == f.lower() or f.lower() in low for f in family)


# ---------------------------------------------------------------------------
# 样式继承链：run rPr -> 字符样式 -> 段落样式(basedOn 链) -> 默认段落样式 -> docDefaults
# ---------------------------------------------------------------------------

def _styles_element(doc):
    """读取 styles.xml，兼容 python-docx 将样式部件加载为通用 Part。"""
    try:
        return doc.styles.element
    except AttributeError:
        try:
            part = doc.part._styles_part
            return _part_element(part)
        except Exception:
            return None


def _style_element(doc, style_id):
    if not style_id:
        return None
    styles = _styles_element(doc)
    for st in (styles.findall(qn('w:style')) if styles is not None else []):
        if st.get(qn('w:styleId')) == style_id:
            return st
    return None


def _style_chain(doc, style_id, limit=8):
    out, sid, seen = [], style_id, set()
    while sid and sid not in seen and len(out) < limit:
        seen.add(sid)
        st = _style_element(doc, sid)
        if st is None:
            break
        out.append(st)
        based = st.find(qn('w:basedOn'))
        sid = based.get(qn('w:val')) if based is not None else None
    return out


def _p_style_id(p_el):
    pPr = p_el.find(qn('w:pPr'))
    if pPr is None:
        return None
    st = pPr.find(qn('w:pStyle'))
    return st.get(qn('w:val')) if st is not None else None


def _default_para_style_id(doc):
    styles = _styles_element(doc)
    for st in (styles.findall(qn('w:style')) if styles is not None else []):
        if st.get(qn('w:type')) == 'paragraph' and st.get(qn('w:default')) in ('1', 'true', 'on'):
            return st.get(qn('w:styleId'))
    return None


def _doc_defaults(doc, which):
    """which='rPr' 取 docDefaults/rPrDefault/rPr；'pPr' 取 pPrDefault/pPr。"""
    styles = _styles_element(doc)
    dd = styles.find(qn('w:docDefaults')) if styles is not None else None
    if dd is None:
        return None
    holder = dd.find(qn('w:%sDefault' % which))
    if holder is None:
        return None
    return holder.find(qn('w:' + which))


def _rpr_sources(doc, p_el, r_el):
    """按 Word 的优先级从高到低列出所有可能承载字符格式的 rPr 容器。"""
    src = []
    if r_el is not None:
        rPr = r_el.find(qn('w:rPr'))
        if rPr is not None:
            src.append(rPr)
            rs = rPr.find(qn('w:rStyle'))
            if rs is not None:
                for st in _style_chain(doc, rs.get(qn('w:val'))):
                    e = st.find(qn('w:rPr'))
                    if e is not None:
                        src.append(e)
    if p_el is not None:
        sid = _p_style_id(p_el)
        for st in _style_chain(doc, sid):
            e = st.find(qn('w:rPr'))
            if e is not None:
                src.append(e)
        if not sid:
            for st in _style_chain(doc, _default_para_style_id(doc)):
                e = st.find(qn('w:rPr'))
                if e is not None:
                    src.append(e)
    e = _doc_defaults(doc, 'rPr')
    if e is not None:
        src.append(e)
    return src


def _resolve_effective_eastasia(doc, p_el, r_el):
    """解析 run 实际生效的中文字体名；无 w:eastAsia 时回落到 w:ascii。"""
    fallback = None
    for rPr in _rpr_sources(doc, p_el, r_el):
        rf = rPr.find(qn('w:rFonts'))
        if rf is None:
            continue
        v = rf.get(qn('w:eastAsia'))
        if v:
            return v
        if fallback is None:
            fallback = rf.get(qn('w:ascii')) or rf.get(qn('w:hAnsi'))
    return fallback


def _resolve_effective_ascii(doc, p_el, r_el):
    """解析 run 实际生效的西文字体名；无 w:ascii 时回落到 w:hAnsi。"""
    for rPr in _rpr_sources(doc, p_el, r_el):
        rf = rPr.find(qn('w:rFonts'))
        if rf is None:
            continue
        v = rf.get(qn('w:ascii')) or rf.get(qn('w:hAnsi'))
        if v:
            return v
    return None


def _resolve_effective_sz(doc, p_el, r_el):
    """解析 run 实际生效的字号（单位：磅）。"""
    for rPr in _rpr_sources(doc, p_el, r_el):
        e = rPr.find(qn('w:sz'))
        if e is not None:
            v = _int_or_none(e.get(qn('w:val')))
            if v:
                return v / 2.0
    return None


def _resolve_effective_bold(doc, p_el, r_el):
    for rPr in _rpr_sources(doc, p_el, r_el):
        e = rPr.find(qn('w:b'))
        if e is not None:
            return (e.get(qn('w:val')) or '1') not in ('0', 'false', 'off')
    return False


def _resolve_effective_color(doc, p_el, r_el):
    for rPr in _rpr_sources(doc, p_el, r_el):
        e = rPr.find(qn('w:color'))
        if e is not None:
            v = (e.get(qn('w:val')) or '').strip().upper()
            if v and v != 'AUTO':
                return v
            return None
    return None


def _ppr_sources(doc, p_el):
    src = []
    pPr = p_el.find(qn('w:pPr'))
    if pPr is not None:
        src.append(pPr)
    sid = _p_style_id(p_el)
    for st in _style_chain(doc, sid):
        e = st.find(qn('w:pPr'))
        if e is not None:
            src.append(e)
    if not sid:
        for st in _style_chain(doc, _default_para_style_id(doc)):
            e = st.find(qn('w:pPr'))
            if e is not None:
                src.append(e)
    e = _doc_defaults(doc, 'pPr')
    if e is not None:
        src.append(e)
    return src


def _p_val(doc, p_el, tag):
    """解析段落级简单属性（jc / outlineLvl / pageBreakBefore ...）的 w:val。"""
    for pPr in _ppr_sources(doc, p_el):
        e = pPr.find(qn('w:' + tag))
        if e is not None:
            return e.get(qn('w:val')) if e.get(qn('w:val')) is not None else '1'
    return None


def _p_child_attr(doc, p_el, tag, attr):
    """逐属性解析 w:ind / w:spacing 的某个属性（Word 按属性粒度合并样式链）。"""
    for pPr in _ppr_sources(doc, p_el):
        e = pPr.find(qn('w:' + tag))
        if e is None:
            continue
        v = e.get(qn('w:' + attr))
        if v is not None:
            return v
    return None


def _p_ind(doc, p_el):
    keys = ('left', 'leftChars', 'right', 'firstLine', 'firstLineChars', 'hanging', 'hangingChars')
    return dict((k, _int_or_none(_p_child_attr(doc, p_el, 'ind', k))) for k in keys)


def _p_spacing(doc, p_el):
    out = {}
    for k in ('before', 'after', 'beforeLines', 'afterLines', 'line'):
        out[k] = _int_or_none(_p_child_attr(doc, p_el, 'spacing', k))
    out['lineRule'] = _p_child_attr(doc, p_el, 'spacing', 'lineRule')
    return out


def _p_tabs(doc, p_el):
    for pPr in _ppr_sources(doc, p_el):
        tabs = pPr.find(qn('w:tabs'))
        if tabs is None:
            continue
        out = []
        for t in tabs.findall(qn('w:tab')):
            out.append((t.get(qn('w:val')), t.get(qn('w:leader')),
                        _int_or_none(t.get(qn('w:pos')))))
        if out:
            return out
    return []


def _p_outline_level(doc, p_el):
    """段落大纲级别：显式 outlineLvl 优先，其次由 heading 1 样式给出 0 级。"""
    v = _int_or_none(_p_val(doc, p_el, 'outlineLvl'))
    if v is not None:
        return v
    for st in _style_chain(doc, _p_style_id(p_el)):
        name = st.find(qn('w:name'))
        nm = (name.get(qn('w:val')) or '').strip().lower() if name is not None else ''
        if nm in ('heading 1', 'heading1', '标题 1', '标题1'):
            return 0
    return None


def _p_no_indent(ind):
    """顶格：无左缩进、无首行缩进（None 或 0 都算顶格）。"""
    return (not ind.get('left') and not ind.get('leftChars')
            and not ind.get('firstLine') and not ind.get('firstLineChars'))


def _is_fixed_line(spacing, pt, tol=1.0):
    """行距是否为固定 pt 磅（lineRule=exact/atLeast 皆按固定值判定）。"""
    if (spacing.get('lineRule') or '') not in ('exact', 'atLeast'):
        return False
    line = spacing.get('line')
    return line is not None and abs(line - pt * 20.0) <= tol * 20.0


def _runs_font_ok(doc, p_el, family, pt, tol=0.3, runs=None):
    """段落内所有可见 run 的中文字体族与字号是否全部达标。"""
    rs = _visible_runs(p_el) if runs is None else runs
    if not rs:
        return False
    for r in rs:
        if family is not None:
            if not _font_in(_resolve_effective_eastasia(doc, p_el, r), family):
                return False
        if pt is not None:
            sz = _resolve_effective_sz(doc, p_el, r)
            if sz is None or abs(sz - pt) > tol:
                return False
    return True


def _has_any_bold(doc, p_el):
    return any(_resolve_effective_bold(doc, p_el, r) for r in _visible_runs(p_el))


def _all_bold(doc, p_el):
    rs = _visible_runs(p_el)
    return bool(rs) and all(_resolve_effective_bold(doc, p_el, r) for r in rs)


# ---------------------------------------------------------------------------
# 域 / 页码 / 分节
# ---------------------------------------------------------------------------

def _field_codes(el):
    """收集元素内所有域代码文本（w:instrText 与 w:fldSimple@instr）。"""
    out = []
    for t in el.iter(qn('w:instrText')):
        out.append(t.text or '')
    for f in el.iter(qn('w:fldSimple')):
        out.append(f.get(qn('w:instr')) or '')
    return out


def _has_field(el, keyword):
    kw = keyword.upper()
    return any(kw in (c or '').upper() for c in _field_codes(el))


def _open_word_document(file_path):
    return Document(file_path)


def _sectPr_list(doc):
    return [s._sectPr for s in doc.sections]


def _part_element(part):
    """读取关联 Part XML，兼容 python-docx 返回通用 Part 的情况。"""
    element = getattr(part, "element", None)
    if element is not None:
        return element
    try:
        from lxml import etree
        blob = getattr(part, "blob", b"")
        return etree.fromstring(blob) if blob else None
    except Exception:
        return None


def _footer_element(doc, sec_idx):
    """取该节实际生效的默认页脚 XML；无显式引用时沿用前一节（linkToPrevious）。"""
    for i in range(sec_idx, -1, -1):
        sectPr = doc.sections[i]._sectPr
        for ref in sectPr.findall(qn('w:footerReference')):
            if (ref.get(qn('w:type')) or 'default') != 'default':
                continue
            rid = ref.get(_R_ID)
            try:
                return _part_element(doc.part.related_parts[rid])
            except KeyError:
                return None
    return None


def _footer_is_explicit(doc, sec_idx):
    """该节是否自带 footerReference(default)，即已断开与前一节的链接。"""
    sectPr = doc.sections[sec_idx]._sectPr
    for ref in sectPr.findall(qn('w:footerReference')):
        if (ref.get(qn('w:type')) or 'default') == 'default':
            return True
    return False


def _nearest_p(el):
    cur = el.getparent()
    while cur is not None and cur.tag != qn('w:p'):
        cur = cur.getparent()
    return cur


def _p_has_page_field(p_el):
    """该段落自身是否含 PAGE 域（PAGEREF / NUMPAGES 不算）。"""
    pat = re.compile(r'(?<![A-Z])PAGE(?![A-Z])')
    for t in p_el.iter(qn('w:instrText')):
        if _nearest_p(t) is p_el and pat.search((t.text or '').upper()):
            return True
    for f in p_el.iter(qn('w:fldSimple')):
        if _nearest_p(f) is p_el and pat.search((f.get(qn('w:instr')) or '').upper()):
            return True
    return False


def _p_has_static_number(p_el):
    """该段落是否只是手工敲入的静态数字。"""
    return bool(re.fullmatch(r'[\s　]*[-—第]?\s*\d{1,3}\s*[页\-—]?[\s　]*', _ptext(p_el) or ''))


def _anchor_centered(el):
    """上溯查找图形锚点：True=锚点水平居中；False=锚点存在但未居中；None=非锚点图形。"""
    cur = el
    while cur is not None:
        if cur.tag == _WP_NS + 'anchor':
            ph = cur.find(_WP_NS + 'positionH')
            if ph is not None:
                al = ph.find(_WP_NS + 'align')
                if al is not None and (al.text or '').strip() == 'center':
                    return True
            return False
        if cur.tag == _WP_NS + 'inline':
            return None
        cur = cur.getparent()
    return None


def _page_field_centered(doc, p_el):
    """PAGE 域所在段落是否水平居中（含文本框/图形锚点居中的情形）。"""
    if (_p_val(doc, p_el, 'jc') or '') == 'center':
        return True
    if _anchor_centered(p_el) is True:
        return True
    outer, cur = p_el, p_el.getparent()
    while cur is not None:
        if cur.tag == qn('w:p'):
            outer = cur
        cur = cur.getparent()
    return outer is not p_el and (_p_val(doc, outer, 'jc') or '') == 'center'


def _top_block(body, el):
    cur = el
    while cur is not None and cur.getparent() is not body:
        cur = cur.getparent()
    return cur


def _block_section_index(doc):
    """每个 body 直接子块所属的节序号（携带 sectPr 的段落归属该节本身）。"""
    body = doc.element.body
    mapping, sec, nsec = {}, 0, max(len(doc.sections), 1)
    for ch in _blocks(body):
        mapping[id(ch)] = min(sec, nsec - 1)
        if ch.tag == qn('w:p'):
            pPr = ch.find(qn('w:pPr'))
            if pPr is not None and pPr.find(qn('w:sectPr')) is not None:
                sec += 1
    return mapping


# ---------------------------------------------------------------------------
# 版面区域定位：封面 / 声明 / 目录 / 六章正文 / 图注 / 表格 / 参考资料
# ---------------------------------------------------------------------------

def _locate(doc):
    body = doc.element.body
    blocks = _blocks(body)
    loc = {'body': body, 'blocks': blocks}

    def find_p(norm_target):
        for ch in blocks:
            if ch.tag == qn('w:p') and _norm(_ptext(ch)) == norm_target:
                return ch
        return None

    loc['school'] = find_p(_norm(COVER_SCHOOL))
    loc['title'] = find_p(_norm(COVER_TITLE))
    loc['kind'] = find_p(_norm(COVER_KIND))
    loc['subtitle'] = find_p(_norm(COVER_SUBTITLE))
    loc['decl_title'] = find_p(_norm(DECL_TITLE))

    # 声明两段正文 + 签字行（按内容定位，避免依赖块序号）
    loc['decl_body'] = []
    for ch in blocks:
        if ch.tag != qn('w:p'):
            continue
        n = _norm(_ptext(ch))
        if any(n == _norm(x) for x in DECL_BODY):
            loc['decl_body'].append(ch)
    loc['decl_sign'] = None
    for ch in blocks:
        if ch.tag == qn('w:p') and '学生签字' in _ptext(ch):
            loc['decl_sign'] = ch
            break

    # "目 录"标题（可能被包进 w:sdt，故在整棵树里找）
    loc['toc_title'] = None
    for p in body.iter(qn('w:p')):
        if _norm(_ptext(p)) == TOC_TITLE_KEY:
            loc['toc_title'] = p
            break

    # 目录域：docPartGallery="Table of Contents" 的 sdt，或 TOC 域代码
    loc['toc_sdt'] = None
    for sdt in body.iter(qn('w:sdt')):
        gal = sdt.find('.//' + qn('w:docPartGallery'))
        if gal is not None and 'table of contents' in (gal.get(qn('w:val')) or '').lower():
            loc['toc_sdt'] = sdt
            break
    if loc['toc_sdt'] is None:
        for sdt in body.iter(qn('w:sdt')):
            if _has_field(sdt, 'TOC'):
                loc['toc_sdt'] = sdt
                break
    loc['toc_field'] = bool(loc['toc_sdt'] is not None or
                            any(re.search(r'\bTOC\b', (c or '').upper()) for c in _field_codes(body)))

    # 六个一级标题：目录之后、正文区内的短段落，按关键词逐个定位
    toc_block = _top_block(body, loc['toc_title']) if loc['toc_title'] is not None else None
    start = blocks.index(toc_block) + 1 if toc_block is not None and toc_block in blocks else 0
    heads = []
    for kw in HEADING_KEYS:
        got = None
        for ch in blocks[start:]:
            if ch.tag != qn('w:p'):
                continue
            n = _norm(_ptext(ch))
            if n and len(n) <= 12 and n.endswith(kw):
                got = ch
                break
        heads.append(got)
    loc['heads'] = heads
    present = [h for h in heads if h is not None]
    body_start = min((blocks.index(h) for h in present if h in blocks), default=len(blocks))
    loc['body_start'] = body_start

    # 目录条目：目录域内的条目段落，或"目 录"标题与第一个一级标题之间的段落
    entries = []
    if loc['toc_sdt'] is not None:
        for p in loc['toc_sdt'].iter(qn('w:p')):
            if p is loc['toc_title'] or not _norm(_ptext(p)):
                continue
            entries.append(p)
    elif toc_block is not None:
        for ch in blocks[start:body_start]:
            if ch.tag == qn('w:p') and _norm(_ptext(ch)):
                entries.append(ch)
    loc['toc_entries'] = entries

    # 图注 / 参考资料条目 / 正文段落
    caps, refs, bodies = [], [], []
    for ch in blocks[body_start:]:
        if ch.tag != qn('w:p'):
            continue
        raw = _ptext(ch)
        n = _norm(raw)
        if not n:
            continue
        if ch in present:
            continue
        if re.match(r'^图\s*\d', raw.strip()):
            caps.append(ch)
            continue
        if re.match(r'^\[\d+\]', raw.strip()):
            refs.append(ch)
            continue
        if re.match(r'^表[\s　]*\d', raw.strip()) or (
                raw.strip().startswith('表') and len(n) <= 24):
            continue
        if list(ch.iter(qn('w:drawing'))) or list(ch.iter(qn('w:pict'))):
            continue
        bodies.append(ch)
    loc['captions'] = caps
    loc['refs'] = refs
    loc['bodies'] = bodies

    # 表格：封面信息表 / 模块功能表
    loc['cover_tbl'] = None
    loc['data_tbl'] = None
    for tbl in body.iter(qn('w:tbl')):
        txt = _ptext(tbl)
        if loc['cover_tbl'] is None and '毕业设计题目' in txt:
            loc['cover_tbl'] = tbl
        if loc['data_tbl'] is None and '主要功能' in txt:
            loc['data_tbl'] = tbl

    # 插图：按文档顺序收集 a:ext 尺寸
    loc['extents'] = []
    for dr in body.iter(qn('w:drawing')):
        ext = None
        for e in dr.iter(_A_NS + 'ext'):
            ext = e
            break
        if ext is not None:
            loc['extents'].append((_int_or_none(ext.get('cx')), _int_or_none(ext.get('cy'))))
        else:
            loc['extents'].append((None, None))
    loc['pict_count'] = len(list(body.iter(qn('w:pict'))))
    return loc


def _cell_paragraphs(tbl):
    """返回 [[ [段落, ...], ... 每列 ], ... 每行 ]。"""
    rows = []
    for tr in tbl.findall(qn('w:tr')):
        row = []
        for tc in tr.findall(qn('w:tc')):
            row.append(tc.findall(qn('w:p')))
        rows.append(row)
    return rows


def _cell_texts(tbl):
    return [[_ptext(tc) for tc in tr.findall(qn('w:tc'))] for tr in tbl.findall(qn('w:tr'))]


# ---------------------------------------------------------------------------
# 维度一：可正常打开的 .docx
# ---------------------------------------------------------------------------

def _check_dim1(file_path, doc, loc):
    """返回 (pass, reason)。仅判"是否为可正常打开的 .docx"，不掺入内容与格式分。"""
    try:
        with zipfile.ZipFile(file_path) as zf:
            names = set(zf.namelist())
    except Exception as exc:
        return False, '文件无法作为 .docx 压缩包打开：%s' % exc
    if 'word/document.xml' not in names:
        return False, '缺少 word/document.xml，不是有效的 .docx'
    return True, ''


# ---------------------------------------------------------------------------
# 主入口
# ---------------------------------------------------------------------------

def _not_found_result(result, message):
    """找不到待评估文件时的返回：dim1_pass 置空字符串，分值与满分均为 0。"""
    result["file_name"] = ""
    result["status"] = "error"
    result["error"] = message
    result["dim1_pass"] = ""
    result["dim1_reason"] = ""
    result["dim2_items"] = []
    result["total_score"] = 0
    result["max_score"] = 0
    return result


def evaluate(dir_path):
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
            return _not_found_result(result, "目录不存在：%s" % dir_path)
        cands = sorted(f for f in os.listdir(dir_path)
                       if f.lower().endswith('.docx') and not f.startswith('~$'))
        if not cands:
            return _not_found_result(result, "目录中未找到 .docx 文件：%s" % dir_path)
        file_path = os.path.join(dir_path, cands[0])
        result["file_name"] = os.path.basename(file_path)
        return _evaluate_impl(file_path, result)
    except Exception as exc:
        result["status"] = "error"
        result["error"] = "%s: %s" % (type(exc).__name__, exc)
        result["dim2_items"] = []
        result["total_score"] = 0
        return result


def _evaluate_impl(file_path, result):
    doc = _open_word_document(file_path)
    loc = _locate(doc)
    results = []

    dim1_pass, dim1_reason = _check_dim1(file_path, doc, loc)
    result["dim1_pass"] = dim1_pass
    if not dim1_pass:
        # 维度一未通过：dim1_pass 为 False，dim1_reason 必须给出错误摘要
        result["dim1_reason"] = dim1_reason or "交付文件不是一个可正常打开的 .docx"
        result["dim2_items"] = []
        result["total_score"] = 0
        result["max_score"] = MAX_SCORE_TOTAL
        return result
    result["dim1_reason"] = ""

    body = loc['body']
    blocks = loc['blocks']
    heads = loc['heads']
    all_paras = list(body.iter(qn('w:p')))

    def pos(el):
        try:
            return all_paras.index(el)
        except ValueError:
            return -1

    # --- 01 纸张与页边距 ------------------------------------------------------
    # 严格对齐细则：① A4 纵向 11906×16838 ② 上2041 下1928 左1531 右1417 ③ 允许±20 twips
    ok = bool(doc.sections)
    for sec in doc.sections:
        sz = sec._sectPr.find(qn('w:pgSz'))
        mg = sec._sectPr.find(qn('w:pgMar'))
        if sz is None or mg is None:
            ok = False
            break
        w, h = _int_or_none(sz.get(qn('w:w'))), _int_or_none(sz.get(qn('w:h')))
        orient = sz.get(qn('w:orient')) or 'portrait'
        if w is None or h is None or abs(w - 11906) > 20 or abs(h - 16838) > 20 or orient != 'portrait':
            ok = False
            break
        for key, exp in (('top', 2041), ('bottom', 1928), ('left', 1531), ('right', 1417)):
            v = _int_or_none(mg.get(qn('w:' + key)))
            if v is None or abs(v - exp) > 20:
                ok = False
                break
        if not ok:
            break
    results.append(("+3：" if ok else "+0：") +
                   "纸张为A4纵向（11906×16838 twips），页边距上3.6cm、下3.4cm、左2.7cm、右2.5cm（±20 twips）")

    # --- 02 封面标题区三行 ----------------------------------------------------
    ok = all(loc[k] is not None for k in ('school', 'title', 'kind'))
    if ok:
        ok = (_runs_font_ok(doc, loc['school'], STRICT_SONGTI, 22.0)
              and _all_bold(doc, loc['school'])
              and (_p_val(doc, loc['school'], 'jc') or '') == 'center'
              and _runs_font_ok(doc, loc['title'], STRICT_HEITI, 24.0)
              and _all_bold(doc, loc['title'])
              and (_p_val(doc, loc['title'], 'jc') or '') == 'center'
              and _runs_font_ok(doc, loc['kind'], STRICT_SONGTI, 15.0)
              and (_p_val(doc, loc['kind'], 'jc') or '') == 'center')
    results.append(("+3：" if ok else "+0：") +
                   "封面标题区三行满足校名宋体二号加粗居中、主标题黑体24磅加粗居中、"
                   "\"（全日制专科生）\"宋体小三居中")

    # --- 03 封面信息表六行 ----------------------------------------------------
    # 行高与整块缩进由标签列段落承载（Word 中行高取同行段落的最大值），
    # 故"左缩进1400 + 固定25磅"核查标签列，"宋体四号"核查全部单元格。
    tbl = loc['cover_tbl']
    rows = _cell_paragraphs(tbl) if tbl is not None else []
    ok = len(rows) == 6
    if ok:
        tblInd = None
        tblPr = tbl.find(qn('w:tblPr'))
        if tblPr is not None:
            ti = tblPr.find(qn('w:tblInd'))
            if ti is not None:
                tblInd = _int_or_none(ti.get(qn('w:w')))
        block_indent = tblInd is not None and 1250 <= tblInd <= 1600
        for row in rows:
            if len(row) < 2:
                ok = False
                break
            for cell in row[:2]:
                for p in cell:
                    if not _norm(_ptext(p)):
                        continue
                    if not _runs_font_ok(doc, p, STRICT_SONGTI, 14.0):
                        ok = False
            row_indent, row_line = block_indent, False
            for p in row[0]:
                if not _norm(_ptext(p)):
                    continue
                lv = _p_ind(doc, p).get('left') or 0
                if 1250 <= lv <= 1600:
                    row_indent = True
                if _is_fixed_line(_p_spacing(doc, p), 25.0):
                    row_line = True
            if not (row_indent and row_line):
                ok = False
                break
    results.append(("+3：" if ok else "+0：") +
                   "封面信息表六行满足宋体、四号（sz=28）、整块左缩进约1400 twips、行距固定25磅")

    # --- 04 独创性声明标题 ----------------------------------------------------
    p = loc['decl_title']
    ok = p is not None
    if ok:
        sp = _p_spacing(doc, p)
        ok = (_runs_font_ok(doc, p, STRICT_HEITI, 16.0) and _all_bold(doc, p)
              and (_p_val(doc, p, 'jc') or '') == 'center'
              and sp.get('before') is not None and abs(sp['before'] - 160) <= 60
              and sp.get('after') is not None and abs(sp['after'] - 360) <= 120
              and _is_fixed_line(sp, 28.0))
    results.append(("+3：" if ok else "+0：") +
                   "\"毕业设计独创性声明\"标题满足黑体、三号、加粗、居中，段前约8磅、段后约18磅、行距固定28磅")

    # --- 05 声明正文与签字行 --------------------------------------------------
    ok = len(loc['decl_body']) == 2 and loc['decl_sign'] is not None
    if ok:
        for p in loc['decl_body']:
            ind, sp = _p_ind(doc, p), _p_spacing(doc, p)
            first_ok = ((ind.get('firstLine') is not None and abs(ind['firstLine'] - 480) <= 40)
                        or (ind.get('firstLineChars') is not None
                            and abs(ind['firstLineChars'] - 200) <= 20))
            if not (_runs_font_ok(doc, p, STRICT_SONGTI, 12.0) and first_ok
                    and _is_fixed_line(sp, 20.0)):
                ok = False
        sign = loc['decl_sign']
        sp = _p_spacing(doc, sign)
        if not (_runs_font_ok(doc, sign, STRICT_SONGTI, 12.0)
                and sp.get('before') is not None and abs(sp['before'] - 480) <= 120):
            ok = False
    results.append(("+3：" if ok else "+0：") +
                   "独创性声明两段正文与学生签字行满足宋体、小四，声明正文首行缩进2字符、"
                   "行距固定20磅，签字行段前约24磅")

    # --- 06 "目 录"标题 -------------------------------------------------------
    p = loc['toc_title']
    ok = p is not None
    if ok:
        raw = (_ptext(p) or '').strip()
        sp = _p_spacing(doc, p)
        line = sp.get('line')
        rule = sp.get('lineRule') or 'auto'
        line_ok = ((rule == 'auto' and line is not None and abs(line - 360) <= 40)
                   or (rule in ('exact', 'atLeast') and line is not None and line >= 405))
        ok = (_runs_font_ok(doc, p, STRICT_HEITI, 15.0)
              and (_p_val(doc, p, 'jc') or '') == 'center'
              and bool(re.fullmatch(r'目[\s　]{2,}录', raw))
              and line_ok)
    results.append(("+3：" if ok else "+0：") +
                   "\"目 录\"标题满足黑体、小三、居中，\"目\"与\"录\"之间空两个字符，行距1.5倍或等效固定行距")

    # --- 07 六个一级标题的字体与段间距 ----------------------------------------
    # 注："六、参考资料"标题另有"黑体小四居中"的单独要求（见第 11 项），
    #     故本项只核查前五个一级标题的黑体四号顶格与段前段后。
    chapter_heads = heads[:5]
    ok = all(h is not None for h in chapter_heads)
    if ok:
        for h in chapter_heads:
            ind, sp = _p_ind(doc, h), _p_spacing(doc, h)
            if not (_runs_font_ok(doc, h, STRICT_HEITI, 14.0) and _p_no_indent(ind)
                    and sp.get('before') is not None and 240 <= sp['before'] <= 480
                    and sp.get('after') is not None and 120 <= sp['after'] <= 360):
                ok = False
                break
    results.append(("+3：" if ok else "+0：") +
                   "六个一级标题满足黑体、四号、顶格无缩进，段前约18磅（360±120）、段后约12磅（240±120）")

    # --- 08 六章正文段落 ------------------------------------------------------
    bodies = loc['bodies']
    ok = len(bodies) >= 15
    if ok:
        for p in bodies:
            ind, sp = _p_ind(doc, p), _p_spacing(doc, p)
            first_ok = ((ind.get('firstLine') is not None and abs(ind['firstLine'] - 480) <= 40)
                        or (ind.get('firstLineChars') is not None
                            and abs(ind['firstLineChars'] - 200) <= 20))
            if not (_runs_font_ok(doc, p, STRICT_SONGTI, 12.0) and first_ok
                    and _is_fixed_line(sp, 20.0)):
                ok = False
                break
            for r in _visible_runs(p):
                if re.search(r'[A-Za-z0-9]', _ptext(r) or ''):
                    if not _font_in(_resolve_effective_ascii(doc, p, r), STRICT_TNR):
                        ok = False
                        break
            if not ok:
                break
    results.append(("+5：" if ok else "+0：") +
                   "六章正文段落全部满足宋体、小四、首行缩进2字符（firstLine=480 或 firstLineChars=200）、"
                   "行距固定20磅，西文与数字为 Times New Roman")

    # --- 09 七条图注 ----------------------------------------------------------
    caps = loc['captions']
    ok = len(caps) == 7
    if ok:
        for cap in caps:
            if not _runs_font_ok(doc, cap, STRICT_HEITI, 10.5):
                ok = False
                break
            top = _top_block(body, cap)
            below_image = False
            if top in blocks:
                for prev in reversed(blocks[:blocks.index(top)]):
                    if prev.tag == qn('w:p') and not _norm(_ptext(prev)) and \
                            not list(prev.iter(qn('w:drawing'))):
                        continue
                    below_image = bool(list(prev.iter(qn('w:drawing'))))
                    break
            if not below_image:
                ok = False
                break
    results.append(("+3：" if ok else "+0：") +
                   "七条图注全部满足黑体、五号（sz=21），且均位于对应插图下方")

    # --- 10 "六、参考资料"标题 -------------------------------------------------
    p = heads[5]
    ok = p is not None and _runs_font_ok(doc, p, STRICT_HEITI, 12.0) \
        and (_p_val(doc, p, 'jc') or '') == 'center'
    results.append(("+3：" if ok else "+0：") +
                   "\"六、参考资料\"标题满足黑体、小四（sz=24）、居中")

    # --- 11 参考资料条目 ------------------------------------------------------
    refs = loc['refs']
    ok = len(refs) == 4
    if ok:
        for p in refs:
            if not (_runs_font_ok(doc, p, STRICT_SONGTI, 10.5) and _p_no_indent(_p_ind(doc, p))):
                ok = False
                break
    results.append(("+3：" if ok else "+0：") +
                   "四条参考资料条目满足宋体、五号（sz=21）、顶格无缩进（无首行缩进、无左缩进）")

    # --- 12 目录域（前置检查项）------------------------------------------------
    toc_field = bool(loc['toc_field'])
    results.append(("+5：" if toc_field else "+0：") +
                   "目录由可自动更新的目录域承载——含 w:sdt（docPartGallery=\"Table of Contents\"）"
                   "或 TOC 域代码，可在Word中通过\"更新域\"整体刷新")

    entries = loc['toc_entries']

    def _core(text):
        t = _norm(text)
        t = re.sub(r'[·．.．…\-—_]+', '', t)
        t = re.sub(r'\d+$', '', t)
        t = re.sub(r'^[一二三四五六七八九十]+、', '', t)
        t = re.sub(r'^\d+[、.]?', '', t)
        return t

    head_cores = [_core(x) for x in HEADINGS]
    entry_cores = [_core(_ptext(p)) for p in entries]

    # --- 13 大纲级别与目录条目一一对应 ----------------------------------------
    ok = toc_field and len(entries) == 6
    if ok:
        for h in heads:
            if h is None or _p_outline_level(doc, h) != 0:
                ok = False
                break
    if ok:
        ok = entry_cores == head_cores
    results.append(("+3：" if ok else "+0：") +
                   "六个一级标题具备大纲级别1并被目录域自动收录为恰好6条条目、与六个标题一一对应，"
                   "封面文字、封面信息表、声明、图注、参考资料条目均未被误纳入目录")

    # --- 14 目录条目可跳转 ----------------------------------------------------
    ok = len(entries) == 6
    if ok:
        for p in entries:
            jump = _has_field(p, 'HYPERLINK')
            if not jump:
                for hl in p.iter(qn('w:hyperlink')):
                    if hl.get(qn('w:anchor')):
                        jump = True
                        break
            if not jump:
                ok = False
                break
    results.append(("+3：" if ok else "+0：") +
                   "6条目录条目均可按住Ctrl单击跳转至正文对应标题（含 HYPERLINK \\l 域或指向标题书签的超链接）")

    # --- 15 目录页码由 PAGEREF 域生成 ------------------------------------------
    bookmarks = set()
    for bm in body.iter(qn('w:bookmarkStart')):
        nm = bm.get(qn('w:name'))
        if nm:
            bookmarks.add(nm)
    ok = toc_field and len(entries) == 6
    if ok:
        for p in entries:
            refs_in = []
            for code in _field_codes(p):
                m = re.search(r'PAGEREF\s+([^\s\\]+)', (code or '').upper())
                if m:
                    refs_in.append(m.group(1))
            if not refs_in:
                ok = False
                break
            upper_bm = set(b.upper() for b in bookmarks)
            if not all(r in upper_bm for r in refs_in):
                ok = False
                break
    results.append(("+3：" if ok else "+0：") +
                   "目录条目页码由 PAGEREF 域生成而非手工数字，且所引书签存在、刷新后与标题实际所在页一致")

    # --- 16 目录条目字体 ------------------------------------------------------
    ok = len(entries) == 6
    if ok:
        for p in entries:
            raw = (_ptext(p) or '').strip()
            gap_ok = bool(re.match(r'^(?:[一二三四五六七八九十]+、|\d+[、.]?)[\s　]{2,}\S', raw))
            if not gap_ok:
                ok = False
                break
            for r in _visible_runs(p):
                txt = (_ptext(r) or '').strip()
                if re.fullmatch(r'[\d\s]+', txt):
                    # 页码与数字：Times New Roman
                    if not _font_in(_resolve_effective_ascii(doc, p, r), STRICT_TNR):
                        ok = False
                        break
                    continue
                if re.fullmatch(r'[·．.…\s　]+', txt):
                    continue
                if not _font_in(_resolve_effective_eastasia(doc, p, r), STRICT_SONGTI):
                    ok = False
                    break
                sz = _resolve_effective_sz(doc, p, r)
                if sz is None or abs(sz - 12.0) > 0.3:
                    ok = False
                    break
                if not _resolve_effective_bold(doc, p, r):
                    ok = False
                    break
            if not ok:
                break
    results.append(("+3：" if ok else "+0：") +
                   "目录条目满足宋体、小四（sz=24），页码与数字为 Times New Roman，一级条目加粗，"
                   "序号与标题文字之间空两个字符")

    # --- 17 目录条目段后与行距 ------------------------------------------------
    ok = len(entries) == 6
    if ok:
        for p in entries:
            sp = _p_spacing(doc, p)
            if not (sp.get('after') is not None and abs(sp['after'] - 100) <= 40
                    and _is_fixed_line(sp, 19.0)):
                ok = False
                break
    results.append(("+3：" if ok else "+0：") +
                   "目录条目满足段后5磅（after=100）、行距固定19磅（line=380、lineRule=exact）")

    # --- 18 目录条目点线制表位 ------------------------------------------------
    ok = len(entries) == 6
    if ok:
        for p in entries:
            hit = False
            for val, leader, tpos in _p_tabs(doc, p):
                if val == 'right' and (leader or '') == 'dot' and tpos is not None \
                        and 8700 <= tpos <= 9200:
                    hit = True
                    break
            if not hit or len(_norm(_ptext(p))) > 40:
                ok = False
                break
    results.append(("+3：" if ok else "+0：") +
                   "目录条目设右对齐制表位并带点线前导符（tab val=right、leader=dot、pos≈8958 落在版心右边界），"
                   "标题与页码之间由连续点线连通且每条单行显示不折行")

    # --- 19 分节实现页码分段控制 ----------------------------------------------
    secmap = _block_section_index(doc)

    def sec_of(el):
        if el is None:
            return None
        top = _top_block(body, el)
        return secmap.get(id(top))

    front_secs = [s for s in (sec_of(loc['school']), sec_of(loc['decl_title']),
                              sec_of(loc['toc_title'])) if s is not None]
    body_sec = sec_of(heads[0]) if heads[0] is not None else None
    ok = (len(doc.sections) >= 2 and body_sec is not None and front_secs
          and body_sec > max(front_secs))
    if ok:
        ok = _footer_is_explicit(doc, body_sec)
    if ok:
        pnt = doc.sections[body_sec]._sectPr.find(qn('w:pgNumType'))
        start = _int_or_none(pnt.get(qn('w:start'))) if pnt is not None else None
        ok = start == 1
    results.append(("+5：" if ok else "+0：") +
                   "通过分节实现页码分段控制——封面、独创性声明、目录页所在节与正文所在节为不同节，"
                   "正文节页脚已断开与前一节的链接，且正文节 pgNumType 起始页码为1")

    # --- 20 正文节页脚 PAGE 域并水平居中 --------------------------------------
    ok = False
    if body_sec is not None:
        ftr = _footer_element(doc, body_sec)
        if ftr is not None:
            for p in ftr.iter(qn('w:p')):
                if _p_has_page_field(p):
                    ok = _page_field_centered(doc, p)
                    if ok:
                        break
    results.append(("+3：" if ok else "+0：") +
                   "正文节页脚中的页码为可自动更新的 PAGE 域（非手工静态数字），并在页面下方水平居中显示")

    # --- 21 封面/声明/目录三页无页码 ------------------------------------------
    ok = bool(front_secs) and body_sec is not None and body_sec not in front_secs
    if ok:
        for s in set(front_secs):
            ftr = _footer_element(doc, s)
            if ftr is None:
                continue
            for p in ftr.iter(qn('w:p')):
                if _p_has_page_field(p) or _p_has_static_number(p):
                    ok = False
                    break
            if not ok:
                break
    results.append(("+3：" if ok else "+0：") +
                   "封面页、独创性声明页、目录页三页均不显示任何页码")

    # --- 22 文档整体结构 ------------------------------------------------------
    anchors = [loc['school'], loc['decl_title'], loc['toc_title']] + list(heads)
    bad = any(a is None for a in anchors)
    if not bad:
        seq = [pos(a) for a in anchors]
        bad = any(x < 0 for x in seq) or any(seq[i] >= seq[i + 1] for i in range(len(seq) - 1))
    results.append(("-5：" if bad else "-0：") +
                   "文档整体结构被破坏——\"封面→毕业设计独创性声明→目 录→一、设计思路→二、设计内容→"
                   "三、设计过程→四、作品及特点→五、致谢→六、参考资料\"的先后顺序未完整保持，"
                   "或出现章节缺失、章节被合并")

    # --- 23 一级标题文字 ------------------------------------------------------
    bad = any(h is None for h in heads)
    if not bad:
        bad = any(_norm(_ptext(h)) != _norm(HEADINGS[i]) for i, h in enumerate(heads))
    results.append(("-3：" if bad else "-0：") +
                   "六个一级标题的文字被改写，未保持为\"一、设计思路\"\"二、设计内容\"\"三、设计过程\""
                   "\"四、作品及特点\"\"五、致谢\"\"六、参考资料\"")

    # --- 24 图注文字 ----------------------------------------------------------
    bad = len(caps) != 7
    if not bad:
        bad = any((_ptext(caps[i]) or '').strip() != CAPTIONS[i] for i in range(7))
    results.append(("-3：" if bad else "-0：") +
                   "七条图注的文字被改写，未逐字保持为\"图1 折叠式仓储搬运车外观结构\"至"
                   "\"图7 整机模块关系\"（如在\"图\"与序号之间插入空格）")

    # --- 25 正文段落文字 ------------------------------------------------------
    bad = len(bodies) != len(BODY_PARAGRAPHS)
    if not bad:
        bad = any(_norm(_ptext(bodies[i])) != _norm(BODY_PARAGRAPHS[i])
                  for i in range(len(BODY_PARAGRAPHS)))
    results.append(("-3：" if bad else "-0：") +
                   "六章正文段落的文字内容被增写、删减、合并、拆分或改写，未逐段保持原有叙述文字")

    # --- 26 插图数量与尺寸 ----------------------------------------------------
    exts = loc['extents']
    bad = len(exts) != 7 or loc['pict_count'] > 0
    if not bad:
        for i, (cx, cy) in enumerate(exts):
            ecx, ecy = IMAGE_EXTENTS[i]
            if cx is None or cy is None or abs(cx - ecx) > ecx * 0.02 or abs(cy - ecy) > ecy * 0.02:
                bad = True
                break
    results.append(("-3：" if bad else "-0：") +
                   "插图数量不为7张，或插图被删除、替换、裁剪、调整长宽比／尺寸／位置，"
                   "或图片与其下方图注的对应关系被打乱")

    # --- 27 模块功能表行列结构 ------------------------------------------------
    texts = _cell_texts(loc['data_tbl']) if loc['data_tbl'] is not None else []
    bad = len(texts) != 5 or any(len(r) != 3 for r in texts)
    if not bad:
        bad = ([_norm(x) for x in texts[0]] != [_norm(x) for x in TABLE_HEADER]
               or [_norm(texts[i + 1][0]) for i in range(4)] !=
               [_norm(x) for x in TABLE_FIRST_COL])
    results.append(("-3：" if bad else "-0：") +
                   "模块功能表的行列结构未完整保持——表头\"模块\"\"主要功能\"\"设计要点\"三列，"
                   "或\"承载箱体\"\"副车架\"\"悬挂轮组\"\"控制面板\"四行被增删、合并、拆分、改序")

    # --- 28 封面信息表/声明/参考资料条目文字 ----------------------------------
    ctexts = _cell_texts(loc['cover_tbl']) if loc['cover_tbl'] is not None else []
    bad = len(ctexts) != 6 or any(len(r) < 2 for r in ctexts)
    if not bad:
        for i, (lab, val) in enumerate(COVER_TABLE_ROWS):
            if (ctexts[i][0] or '').strip() != lab or (ctexts[i][1] or '').strip() != val:
                bad = True
                break
    if not bad:
        if len(loc['decl_body']) != 2:
            bad = True
        else:
            for i, p in enumerate(loc['decl_body']):
                if _norm(_ptext(p)) != _norm(DECL_BODY[i]):
                    bad = True
                    break
    if not bad:
        sign_raw = _ptext(loc['decl_sign']) if loc['decl_sign'] is not None else ''
        if _collapse(sign_raw) != DECL_SIGN_CORE or len(sign_raw) < DECL_SIGN_MIN_LEN:
            bad = True
    if not bad:
        if len(refs) != 4:
            bad = True
        else:
            for i, p in enumerate(refs):
                if (_ptext(p) or '').strip() != REFERENCES[i]:
                    bad = True
                    break
    results.append(("-3：" if bad else "-0：") +
                   "封面信息表六行文字（含标签中的空格间隔）、独创性声明两段正文与"
                   "\"学生签字：　　日期：　年　月　日\"一行、四条参考资料条目中任一处文字被改写、删除或压缩")

    # --- 29 批注/修订/高亮/底纹/彩色字体 --------------------------------------
    with zipfile.ZipFile(file_path) as zf:
        part_names = set(zf.namelist())
    bad = 'word/comments.xml' in part_names
    if not bad:
        for tag in ('commentReference', 'commentRangeStart', 'ins', 'del',
                    'moveFrom', 'moveTo', 'pPrChange', 'rPrChange'):
            if list(body.iter(qn('w:' + tag))):
                bad = True
                break
    if not bad:
        for hl in body.iter(qn('w:highlight')):
            if (hl.get(qn('w:val')) or 'none') != 'none':
                bad = True
                break
    if not bad:
        # 交付文档的封面信息表与模块功能表本就带有淡蓝底纹，属既有排版，不计入"多余底纹"
        existing_fills = {'AUTO', 'FFFFFF', 'E8EEF4', 'F7F9FB', 'DCEAF4'}
        for shd in body.iter(qn('w:shd')):
            anc, in_table = shd, False
            while anc is not None:
                if anc.tag == qn('w:tbl'):
                    in_table = True
                    break
                anc = anc.getparent()
            if in_table:
                continue
            fill = (shd.get(qn('w:fill')) or 'auto').upper()
            pat = (shd.get(qn('w:val')) or 'clear').lower()
            if fill not in existing_fills or pat not in ('clear', 'nil', 'none'):
                bad = True
                break
    if not bad:
        allowed = {'000000', COVER_BLUEGRAY, COVER_SUBTITLE_GRAY, COVER_KIND_GRAY}
        toc_ids = set(id(x) for x in entries)
        if loc['toc_title'] is not None:
            toc_ids.add(id(loc['toc_title']))
        for p in body.iter(qn('w:p')):
            if id(p) in toc_ids:
                continue
            for r in _own_runs(p):
                rPr = r.find(qn('w:rPr'))
                if rPr is None:
                    continue
                c = rPr.find(qn('w:color'))
                if c is None:
                    continue
                v = (c.get(qn('w:val')) or '').strip().upper()
                if v and v != 'AUTO' and v not in allowed:
                    bad = True
                    break
            if bad:
                break
    results.append(("-3：" if bad else "-0：") +
                   "文档中被插入未经要求的批注、修订标记、文字高亮、底纹或彩色字体")

    # --- 30 页眉/水印/分隔线/空白页/空段落堆叠 --------------------------------
    bad = False
    for i in range(len(doc.sections)):
        sectPr = doc.sections[i]._sectPr
        if sectPr.find(qn('w:pgBorders')) is not None:
            bad = True
            break
        for ref in sectPr.findall(qn('w:headerReference')):
            rid = ref.get(_R_ID)
            try:
                hdr = _part_element(doc.part.related_parts[rid])
            except KeyError:
                continue
            if hdr is None:
                continue
            if _norm(_ptext(hdr)) or list(hdr.iter(qn('w:pict'))) or list(hdr.iter(qn('w:drawing'))):
                bad = True
                break
        if bad:
            break
    if not bad and loc['pict_count'] > 0:
        bad = True
    if not bad:
        for p in blocks:
            if p.tag != qn('w:p'):
                continue
            for pPr in _ppr_sources(doc, p)[:1]:
                if pPr.find(qn('w:pBdr')) is not None:
                    bad = True
                    break
            if bad:
                break
    if not bad:
        run_len = 0
        for ch in blocks:
            if ch.tag == qn('w:p') and not _norm(_ptext(ch)) and \
                    not list(ch.iter(qn('w:drawing'))):
                run_len += 1
                if run_len >= EMPTY_RUN_LIMIT:
                    bad = True
                    break
            else:
                run_len = 0
    results.append(("-3：" if bad else "-0：") +
                   "文档中被插入未经要求的页眉文字、水印、横向分隔线、装饰边框、额外空白页，"
                   "或封面、独创性声明、目录、正文之间的换页被改为堆叠连续空段落实现")

    # --- 31 图注被附加加粗 ----------------------------------------------------
    bad = any(_has_any_bold(doc, cap) for cap in caps)
    results.append(("-3：" if bad else "-0：") +
                   "七条图注被附加加粗属性（黑体五号图注不应加粗）")

    # --- 32 封面副标题排版与既有配色 ------------------------------------------
    sub = loc['subtitle']
    bad = sub is None
    if not bad:
        szs = [_resolve_effective_sz(doc, sub, r) for r in _visible_runs(sub)]
        col = [_resolve_effective_color(doc, sub, r) for r in _visible_runs(sub)]
        bad = ((_p_val(doc, sub, 'jc') or '') != 'center'
               or not szs or any(s is None or abs(s - 11.0) > 0.5 for s in szs)
               or any(c != COVER_SUBTITLE_GRAY for c in col))
    if not bad:
        for p in [loc['school'], loc['title']] + list(heads):
            if p is None:
                bad = True
                break
            cols = [_resolve_effective_color(doc, p, r) for r in _visible_runs(p)]
            if not cols or any(c != COVER_BLUEGRAY for c in cols):
                bad = True
                break
    results.append(("-1：" if bad else "-0：") +
                   "封面\"结构方案、控制布置与使用性能说明\"一行不满足居中、字号约11磅、深灰色（646E78）的"
                   "副标题排版，或封面校名、主标题与六个一级标题的深蓝灰色（36475C）字体颜色被去除、改色")

    # --- 33 封面信息表标签列排版 ----------------------------------------------
    bad = len(rows) != 6
    if not bad:
        for row in rows:
            if len(row) < 2:
                bad = True
                break
            for p in row[0]:
                if not _norm(_ptext(p)):
                    continue
                if not _all_bold(doc, p) or (_p_val(doc, p, 'jc') or '') != 'right':
                    bad = True
                    break
            for p in row[1]:
                if not _norm(_ptext(p)):
                    continue
                if (_p_val(doc, p, 'jc') or 'left') not in ('left', 'start'):
                    bad = True
                    break
            if bad:
                break
    results.append(("-1：" if bad else "-0：") +
                   "封面信息表不满足标签列加粗右对齐、取值列左对齐的对照排版")

    # --- 34 目录页额外说明文字 ------------------------------------------------
    bad = any(c not in head_cores for c in entry_cores)
    if not bad:
        scope = list(loc['toc_sdt'].iter(qn('w:p'))) if loc['toc_sdt'] is not None else \
            [ch for ch in blocks[:loc['body_start']] if ch.tag == qn('w:p')
             and loc['toc_title'] is not None and pos(ch) >= pos(loc['toc_title'])]
        for p in scope:
            t = _ptext(p) or ''
            if re.search(r'contents', t, re.I) or '目录说明' in t or re.search(r'共\s*\d+\s*页', t):
                bad = True
                break
    results.append(("-1：" if bad else "-0：") +
                   "\"目 录\"所在页除\"目 录\"二字与6条目录条目外，被插入\"Contents\"\"目录说明\"\"共X页\"等"
                   "额外说明性文字或空白占位段落")

    # --- 汇总：把 results 反解为结构化 dim2_items ------------------------------
    if len(results) != len(_MAX_DELTAS_ORDER):
        raise RuntimeError("dim2 项数不一致：results=%d vs order=%d"
                           % (len(results), len(_MAX_DELTAS_ORDER)))
    dim2_items = []
    total_score = 0
    for raw, max_delta in zip(results, _MAX_DELTAS_ORDER):
        head, _, rule = raw.partition("：")
        try:
            value = int(head.replace("+", "").strip())
        except ValueError:
            value = 0
        if max_delta > 0:
            hit = value > 0
            delta = max_delta if hit else 0
        else:
            hit = value != 0
            delta = max_delta if hit else 0
        dim2_items.append({"rule": rule, "max_delta": max_delta,
                           "delta": delta, "hit": hit})
        total_score += delta
    result["dim2_items"] = dim2_items
    result["total_score"] = total_score
    result["max_score"] = MAX_SCORE_TOTAL
    return result


if __name__ == "__main__":
    _dir = sys.argv[1] if len(sys.argv) >= 2 else os.path.dirname(os.path.abspath(__file__))
    try:
        print(json.dumps(evaluate(_dir), ensure_ascii=False, indent=2))
    except Exception:
        traceback.print_exc()
        raise
