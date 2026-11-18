#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# pyright: reportUnknownMemberType=false, reportUnknownVariableType=false, reportUnknownArgumentType=false, reportUnknownParameterType=false, reportMissingParameterType=false, reportPrivateUsage=false, reportUnknownLambdaType=false, reportGeneralTypeIssues=false

import os
import re
import json
import zipfile
from openpyxl import load_workbook
from openpyxl.worksheet.formula import ArrayFormula

# 说明：按接口统一约定，本模块不修改全局 sys.stdout，也不 print 主结果。
# 主结果统一由 evaluate() -> dict 返回；评估目标锁定为本脚本所在目录，脚本自己在该
# 目录内定位并打开被评估的 .xlsx（接受 .xlsm），不接受外部传入的目录，也不检索子目录
# 与上级目录。本文件仅在 __main__ 下作为本地调试入口，才把 dict 序列化为 JSON 打印。
#
# 与 008 号脚本的差异：本任务（标签维度表 勾选联动）细则同时含“加分项”（指令要求
# 的勾选交互 + 右侧自动汇总）与“减分项”（指令未涉及却被改动的标题/表头/标签文字/
# 样式/空白区等）。故 dim2 项的 max_delta 允许为负：正值=加分项（命中记满分，未命
# 中记0）；负值=减分项（触发“破坏”记该负分，未触发记0）。total_score 为全部 delta
# 之和，取值区间 [-17, +40]；max_score 仍为所有加分项满分之和（+40），作为完成度上限。
#
# 关于宏：指令未限制实现手段，用工作簿内 VBA 事件宏实现“勾选后右侧自动显示”同样计分，
# .xlsm 不作减分；仅当勾选/汇总必须手动运行宏或点按钮才刷新时，才不满足 A06 的“自动”。
#
# 计分前置条件（细则原文）：加分项只计“为实现勾选联动而新增的工作量”。若“可勾选
# 入口”与“右侧自动汇总”两项核心功能均未实现（表格仍为静态文字、右侧无任何公式输
# 出），则 A01~A14 全部记 0，避免未做任何改动的文件凭“能打开、样式未变”得分。

SCRIPT_ID = "082"

# 评估目标锁定为本脚本所在目录：评测时把本脚本与被评交付物放在同一个目录下即可。
_SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

# —— 加分项按“结果追加顺序”的 max_delta（正值）——
_ADD_DELTAS_ORDER = [
    5,  # A01 B5:C68 单击即可切换选中/取消（复选框类控件）
    5,  # A02 勾选入口覆盖 126 个有文字标签单元格，B32/B58 不设入口
    3,  # A03 控件绑定独立状态单元格，TRUE/FALSE 可被公式引用
    3,  # A04 控件与标签单元格一一对齐、落在单元格边框内
    1,  # A05 标签单元格不再残留「□」静态占位符号
    5,  # A06 I 列表头下方用公式自动输出已勾选内容
    3,  # A07 汇总条目给出“分类维度 + 标签名称”的对应关系
    3,  # A08 汇总结果连续排列、无空行、取消勾选自动上移
    3,  # A09 汇总公式引用范围覆盖第 5 行至第 68 行全部数据
    3,  # A10 汇总区行数足以容纳 126 个标签，末行不被截断
    1,  # A11 未勾选时汇总区为空白，无 #N/A 等错误值
    1,  # A12 汇总单元格自动换行或足够列宽，文本不被截断
    1,  # A13 汇总区表头沿用 115E59 深绿底 + 白色加粗 + 居中
    3,  # A14 全部由工作簿内控件 + 公式（或工作簿内事件宏）实现，无需外部脚本/手动刷新
]

# —— 减分项按“结果追加顺序”的 max_delta（负值）——
_DED_DELTAS_ORDER = [
    -1,  # D01 A1:C1 合并标题行格式
    -1,  # D02 A2:C2 浅色分隔条格式
    -1,  # D03 第 3 行被写入内容或行高被加高
    -1,  # D04 A4/B4/C4 表头文字与格式
    -3,  # D05 B5:C68 标签文字被增删改写
    -1,  # D06 A5:A68 维度列格式与分组填充色
    -1,  # D07 B5:C68 标签单元格格式与交替填充色
    -1,  # D08 第 5~68 行行高 22
    -1,  # D09 A5:C68 被插入/删除行列导致错行
    -1,  # D10 A/B/C 列列宽被调整到无法显示标签文字
    -1,  # D11 I2:J3 空白区被写入文字或填充色
    -1,  # D12 I4 被与 J4 合并成跨列单元格
    -1,  # D13 I4 填充色/字号字体/居中/边框被改
    -1,  # D14 工作表数量或名称发生变化
    -1,  # D15 冻结窗格/网格线/标签颜色等视图设置被改动
]

MAX_SCORE_TOTAL = sum(_ADD_DELTAS_ORDER)               # +40：完成度上限
_MIN_TOTAL = MAX_SCORE_TOTAL + sum(_DED_DELTAS_ORDER)  # +23：全部命中且全部踩坑
assert MAX_SCORE_TOTAL == 40
assert sum(_DED_DELTAS_ORDER) == -17

# 结果追加顺序 = 先加分项、后减分项
_MAX_DELTAS_ORDER = _ADD_DELTAS_ORDER + _DED_DELTAS_ORDER
_ADD_COUNT = len(_ADD_DELTAS_ORDER)

# ============================================================================
# 基线常量：左侧标签表的既有文字与版式（细则中“应保持不变”的部分）
# ============================================================================
SHEET_NAME = '标签维度表'
DATA_R1, DATA_R2 = 5, 68            # 标签数据区行范围
LABEL_COLS = (2, 3)                 # B、C 两列为标签列
DIM_COL = 1                         # A 列为维度列
SUM_COL_CANDIDATES = (9, 10, 11, 12)  # 右侧汇总区候选列 I~L
HELPER_COL_CANDIDATES = tuple(range(4, 9))  # D~H：辅助列候选（须隐藏）
EMU_PER_PX = 9525

_TITLE_TEXT = '用户标签维度表'
_HDR_TEXTS = {1: '分类维度', 2: '标签选项一', 3: '标签选项二'}
_HDR_FILL = '115E59'                # 左侧第 4 行表头与右侧汇总区表头共用的深绿底
_TICK_CHARS = '□☑☐■◻◼✓✔√'
_CHECKED_CHARS = '☑■◼✓✔√'

# A 列 5 个维度分组：(起始行, 结束行, 维度名, 分段填充色)
_DIM_GROUPS = (
    (5, 18, '基础信息', 'E0F2FE'),
    (19, 32, '花销去向', 'ECFDF5'),
    (33, 45, '使用场景', 'FEF3C7'),
    (46, 58, '日常爱好', 'F5F3FF'),
    (59, 68, '购买动因', 'FFE4E6'),
)

# B/C 标签单元格分段交替填充色：奇数行 F8FAFC、偶数行 FFFFFF
_LABEL_FILL_ODD = 'F8FAFC'
_LABEL_FILL_EVEN = 'FFFFFF'
_LABEL_BORDER_RGB = 'CBD5E1'

# 标签基线文字（已去掉「□ 」前缀）：行号 -> (B 列标签, C 列标签)
# 其中 B32、B58 无文字，为空串——该两处不应出现任何标签文字或勾选入口。
_LABELS = {
    5: ('新用户', '线上购物偏好'),
    6: ('老用户', '线下体验偏好'),
    7: ('单人使用', '喜欢提前规划'),
    8: ('多人共享', '临时决策较多'),
    9: ('学生群体', '注重口碑评价'),
    10: ('初入职场', '常看测评内容'),
    11: ('稳定职场', '偏好品牌直营'),
    12: ('家庭用户', '接受新品牌'),
    13: ('城市通勤', '喜欢套装组合'),
    14: ('近郊生活', '偏好单品购买'),
    15: ('租房居住', '关注售后服务'),
    16: ('自有住房', '关注配送时效'),
    17: ('偶尔购买', '会员权益敏感'),
    18: ('高频购买', '活动促销敏感'),
    19: ('行政办公', '入门预算'),
    20: ('市场运营', '中等预算'),
    21: ('产品策划', '高预算'),
    22: ('设计创意', '注重长期成本'),
    23: ('客户服务', '关注一次性投入'),
    24: ('数据分析', '喜欢分期付款'),
    25: ('电商运营', '经常比价'),
    26: ('内容编辑', '看重品牌溢价'),
    27: ('门店管理', '重视实用功能'),
    28: ('物流配送', '偏好高颜值设计'),
    29: ('财务会计', '愿意为服务付费'),
    30: ('人力资源', '常用企业采购'),
    31: ('项目协调', '常用个人采购'),
    32: ('', '关注发票报销'),
    33: ('日常通勤', '室内使用'),
    34: ('周末短途', '室外使用'),
    35: ('长途出差', '移动携带'),
    36: ('家庭聚会', '固定摆放'),
    37: ('校园活动', '小空间适用'),
    38: ('办公会议', '大空间适用'),
    39: ('社群活动', '多设备协同'),
    40: ('亲友送礼', '夜间使用'),
    41: ('新居布置', '高频清洁维护'),
    42: ('节日采购', '低维护需求'),
    43: ('临时应急', '轻量化需求'),
    44: ('计划囤货', '耐用性需求'),
    45: ('户外休闲', '可收纳需求'),
    46: ('数码产品', '新品资讯'),
    47: ('家居整理', '购物清单'),
    48: ('运动健身', '使用教程'),
    49: ('美食烹饪', '深度评测'),
    50: ('旅行攻略', '对比测评'),
    51: ('亲子教育', '真实案例'),
    52: ('职场技能', '清单模板'),
    53: ('财务规划', '经验分享'),
    54: ('影音娱乐', '短视频内容'),
    55: ('手作创意', '图文内容'),
    56: ('绿色生活', '直播讲解'),
    57: ('宠物护理', '社群讨论'),
    58: ('', '专家建议'),
    59: ('节省时间', '品质稳定'),
    60: ('降低成本', '操作简单'),
    61: ('提升体验', '外观协调'),
    62: ('减少麻烦', '售后可靠'),
    63: ('增强安全感', '使用灵活'),
    64: ('获得成就感', '便于携带'),
    65: ('保持整洁', '适合长期用'),
    66: ('提高效率', '适合多人用'),
    67: ('改善舒适度', '环保节能'),
    68: ('方便分享', '可扩展升级'),
}
# 126 个“有标签文字”的单元格（应各有一个勾选入口）：(行, 列)
_LABEL_CELLS = frozenset(
    (r, c) for r in range(DATA_R1, DATA_R2 + 1)
    for c in LABEL_COLS if _LABELS[r][c - 2]
)
# 两个无文字单元格（不应设勾选入口）
_EMPTY_CELLS = frozenset(
    (r, c) for r in range(DATA_R1, DATA_R2 + 1)
    for c in LABEL_COLS if not _LABELS[r][c - 2]
)
assert len(_LABEL_CELLS) == 126 and _EMPTY_CELLS == frozenset({(32, 2), (58, 2)})

# 基线维度名与标签名集合：用于识别“宏实现”在右侧写下的静态汇总值
_DIM_NAMES = frozenset(name for _, _, name, _ in _DIM_GROUPS)
_ALL_LABELS = frozenset(_LABELS[r][c - 2] for (r, c) in _LABEL_CELLS)
# 最后一个维度分组（第 59~68 行）的标签：出现即说明汇总覆盖到了表格末行
_LAST_GROUP_LABELS = frozenset(
    _LABELS[r][c - 2] for (r, c) in _LABEL_CELLS if r >= _DIM_GROUPS[-1][0])

# 汇总公式中“连续排列/自动上移”所依赖的取第 k 个函数
_RANK_FUNCS = ('AGGREGATE', 'SMALL(', 'LARGE(', 'FILTER(', 'SORT(', 'TEXTJOIN',
               'TOCOL(', 'MATCH(')
_ERR_GUARDS = ('IFERROR', 'IFNA', 'ISERROR', 'ISNA')


# ============================================================================
# 通用取值/取格式工具
# ============================================================================
def _rgb6(color):
    """把 openpyxl 颜色对象归一为 6 位十六进制（丢掉 alpha 前缀）；主题色/索引色返回 None。"""
    if color is None:
        return None
    v = getattr(color, 'rgb', None)
    if not isinstance(v, str):
        return None
    v = v.upper()
    return v[-6:] if len(v) >= 6 else None


def _fill6(cell):
    """单元格填充色（无填充返回 None）。"""
    try:
        f = cell.fill
        if f is None or not f.patternType:
            return None
        return _rgb6(f.fgColor)
    except Exception:
        return None


def _txt(v):
    """单元格文本（None/公式对象归一为字符串）。"""
    if v is None:
        return ''
    if isinstance(v, ArrayFormula):
        return str(v.text or '')
    return str(v)


def _label_text(v):
    """标签文字归一：去掉勾选符号与首尾空白，便于与基线文字比对。"""
    s = _txt(v)
    for ch in _TICK_CHARS:
        s = s.replace(ch, '')
    return s.strip()


def _formula(cell):
    """返回单元格中的公式文本（含 = 前缀）；非公式返回 None。
       兼容普通公式（str）与动态数组公式（ArrayFormula）。"""
    v = cell.value
    if isinstance(v, ArrayFormula):
        t = v.text or ''
        return t if t.startswith('=') else None
    if isinstance(v, str) and v.startswith('='):
        # data_type == 's' 说明是被当成文本存下来的“假公式”，不计为公式
        return v if getattr(cell, 'data_type', 'f') == 'f' else None
    return None


# Excel 的“未来函数”（Excel 2007 之后新增的函数）在 xlsx 内部必须写成 _xlfn.XXX，
# 见 [MS-XLSX] Functions 的 future function 表（含 _xlfn.AGGREGATE、_xlfn.TEXTJOIN 等）。
# 若文件里直接写成 AGGREGATE(...)，Excel 打开后会把它当作未定义名称 → 返回 #NAME?；
# 外层再套 IFERROR/IFNA 时错误被吞成空字符串，汇总区将永远显示空白，等同于未实现。
# 因此这类公式不计为“可用的汇总公式”。
_FUTURE_FUNCS = (
    'AGGREGATE', 'IFNA', 'IFS', 'TEXTJOIN', 'CONCAT', 'MAXIFS', 'MINIFS', 'SWITCH',
    'XLOOKUP', 'FILTER', 'SORT', 'SORTBY', 'UNIQUE', 'SEQUENCE', 'RANDARRAY',
    'TOCOL', 'TOROW', 'LET', 'LAMBDA', 'TAKE', 'DROP', 'HSTACK', 'VSTACK',
    'CHOOSECOLS', 'CHOOSEROWS', 'TEXTSPLIT', 'TEXTBEFORE', 'TEXTAFTER',
    'FORMULATEXT', 'ISFORMULA', 'XOR',
)
# 前置的 (?<![A-Za-z0-9_.]) 保证已带 _xlfn. / _xlfn._xlws. 前缀的写法不会被误判
_FUTURE_RE = re.compile(r'(?<![A-Za-z0-9_.])(' + '|'.join(_FUTURE_FUNCS) + r')\s*\(')


def _formula_broken(text):
    """公式在 Excel 中是否必然算不出结果（未来函数缺 _xlfn. 前缀 → #NAME?）。"""
    return bool(text) and bool(_FUTURE_RE.search(text))


# 指令未限制实现手段，用工作簿内 VBA 实现“勾选后自动显示”同样计分。事件宏（勾选即触发）
# 满足“自动”；只有必须手动点按钮/运行宏才刷新的做法不满足。vbaProject.bin 中模块源码是
# 压缩存放的，事件过程名并不总是明文可见，故：能在原始字节里找到事件过程名即判为事件驱动；
# 找不到时按“无法静态确认”从宽处理（仍判为自动），并在该项的规则说明里写明未能确认。
_VBA_EVENT_HINTS = (
    b'Worksheet_Change', b'Worksheet_SelectionChange', b'Worksheet_BeforeDoubleClick',
    b'Workbook_SheetChange', b'Workbook_SheetSelectionChange', b'_Click', b'_Change',
)
_VBA_MANUAL_HINTS = (b'CommandButton', b'Shapes(', b'Button_Click')


def _vba_profile(parts):
    """返回 (是否含宏, 是否检出事件过程名, 是否检出手动按钮痕迹)。"""
    raw = parts.get('vba_raw') or b''
    if not parts['has_vba']:
        return (False, False, False)
    ev = any(h in raw for h in _VBA_EVENT_HINTS)
    mn = any(h in raw for h in _VBA_MANUAL_HINTS)
    return (True, ev, mn)


def _side(cell, name):
    """边框某一侧的 (线型, 颜色6位)；未设置返回 (None, None)。"""
    try:
        s = getattr(cell.border, name, None)
        if s is None:
            return (None, None)
        return (s.style, _rgb6(s.color))
    except Exception:
        return (None, None)


def _centered(cell):
    al = cell.alignment
    return bool(al) and al.horizontal == 'center' and al.vertical == 'center'


def _white_or_theme(cell):
    """白色字（或主题色/未显式设色——基线表头即以主题白字保存，不据此判失分）。"""
    c = _rgb6(cell.font.color) if cell.font else None
    return c in (None, 'FFFFFF')


def _sz(cell):
    try:
        return float(cell.font.sz) if cell.font and cell.font.sz else None
    except Exception:
        return None
def _col_dims(ws):
    """解析列宽/隐藏：openpyxl 把 <col min=".." max=".."/> 合并存放（如 min=5,max=6
       同时描述 E、F 两列），必须按 min~max 展开，否则 F 列会被误判为“未隐藏”。
       返回 {列号: (列宽, 是否隐藏)} 与默认列宽。"""
    try:
        default = ws.sheet_format.defaultColWidth
    except Exception:
        default = None
    default = float(default) if default else 8.43
    out = {}
    try:
        dims = list(ws.column_dimensions.values())
    except Exception:
        dims = []
    for cd in dims:
        lo = getattr(cd, 'min', None)
        hi = getattr(cd, 'max', None)
        if not lo:
            continue
        for i in range(int(lo), int(hi or lo) + 1):
            out[i] = (float(cd.width) if cd.width else default, bool(cd.hidden))
    return out, default


def _col_w(ws, idx, cache=None):
    m, default = cache if cache else _col_dims(ws)
    return m.get(idx, (default, False))[0]


def _col_hidden(ws, idx, cache=None):
    m, default = cache if cache else _col_dims(ws)
    return m.get(idx, (default, False))[1]


def _row_h(ws, r):
    try:
        rd = ws.row_dimensions.get(r)
    except Exception:
        rd = None
    if rd is None or rd.height is None:
        return None
    return float(rd.height)


def _merged_ranges(ws):
    try:
        return [(rng.min_row, rng.min_col, rng.max_row, rng.max_col, str(rng))
                for rng in ws.merged_cells.ranges]
    except Exception:
        return []


def _has_merge(mrs, r1, c1, r2, c2):
    """判断给定矩形区域内是否存在（哪怕部分重叠的）合并单元格。"""
    for mr, mc, xr, xc, _ in mrs:
        if not (xr < r1 or mr > r2 or xc < c1 or mc > c2):
            return True
    return False


def _is_merged_exact(mrs, name):
    return any(s == name for *_, s in mrs)


def _read_parts(file_path):
    """读取包内需要的原始部件：勾选控件、VML、宏等。"""
    p = {'names': [], 'sheet_xml': '', 'rels': '', 'ctrlprops': {}, 'vml': '',
         'has_vba': False, 'vba_raw': b'', 'activex': 0}
    with zipfile.ZipFile(file_path) as zf:
        p['names'] = zf.namelist()
        for n in p['names']:
            low = n.lower()
            if low.endswith('vbaproject.bin'):
                p['has_vba'] = True
                p['vba_raw'] = zf.read(n)
            if low.startswith('xl/activex/') and low.endswith('.xml'):
                p['activex'] += 1
        for n in ('xl/worksheets/sheet1.xml',):
            if n in p['names']:
                p['sheet_xml'] = zf.read(n).decode('utf-8', 'ignore')
        if 'xl/worksheets/_rels/sheet1.xml.rels' in p['names']:
            p['rels'] = zf.read('xl/worksheets/_rels/sheet1.xml.rels').decode('utf-8', 'ignore')
        for n in p['names']:
            if n.startswith('xl/ctrlProps/') and n.endswith('.xml'):
                p['ctrlprops'][os.path.basename(n)] = zf.read(n).decode('utf-8', 'ignore')
            elif n.startswith('xl/drawings/') and n.endswith('.vml'):
                p['vml'] += zf.read(n).decode('utf-8', 'ignore')
    return p
_CTRL_BLOCK_RE = re.compile(r'<control\s[^>]*/>|<control\s[^>]*>.*?</control>', re.S)
_ANCH_RE = re.compile(
    r'<(?:\w+:)?from>\s*<(?:\w+:)?col>(-?\d+)</(?:\w+:)?col>\s*'
    r'<(?:\w+:)?colOff>(-?\d+)</(?:\w+:)?colOff>\s*'
    r'<(?:\w+:)?row>(-?\d+)</(?:\w+:)?row>\s*'
    r'<(?:\w+:)?rowOff>(-?\d+)</(?:\w+:)?rowOff>\s*</(?:\w+:)?from>\s*'
    r'<(?:\w+:)?to>\s*<(?:\w+:)?col>(-?\d+)</(?:\w+:)?col>\s*'
    r'<(?:\w+:)?colOff>(-?\d+)</(?:\w+:)?colOff>\s*'
    r'<(?:\w+:)?row>(-?\d+)</(?:\w+:)?row>\s*'
    r'<(?:\w+:)?rowOff>(-?\d+)</(?:\w+:)?rowOff>\s*</(?:\w+:)?to>', re.S)


def _parse_checkbox_controls(parts):
    """解析工作表内的“复选框”表单控件：锚点单元格 + 绑定的状态单元格。
       主路径取 sheet1.xml 的 <controls>（锚点）+ 关系表 -> ctrlProps（objectType/
       fmlaLink）；无 <controls> 时退回 VML 的 x:ClientData（锚点 + FmlaLink）。"""
    # rId -> ctrlPropN.xml（属性顺序不固定，逐个 <Relationship> 标签取）
    rid2prop = {}
    for m in re.finditer(r'<Relationship\b[^>]*>', parts['rels']):
        tag = m.group(0)
        rid = re.search(r'Id="([^"]+)"', tag)
        tgt = re.search(r'ctrlProps/([^"/]+)"', tag)
        if rid and tgt:
            rid2prop[rid.group(1)] = tgt.group(1)

    def prop_of(rid):
        name = rid2prop.get(rid)
        xml = parts['ctrlprops'].get(name, '') if name else ''
        objtype = re.search(r'objectType="([^"]+)"', xml)
        link = re.search(r'fmlaLink="([^"]+)"', xml)
        return ((objtype.group(1) if objtype else None),
                (link.group(1) if link else None),
                ('checked="Checked"' in xml))

    out = {}
    for block in _CTRL_BLOCK_RE.findall(parts['sheet_xml']):
        rid = re.search(r'r:id="([^"]+)"', block)
        if not rid:
            continue
        rid = rid.group(1)
        anchor = None
        m = _ANCH_RE.search(block)
        if m:
            g = [int(x) for x in m.groups()]
            anchor = {'fc': g[0], 'fco': g[1], 'fr': g[2], 'fro': g[3],
                      'tc': g[4], 'tco': g[5], 'tr': g[6], 'tro': g[7]}
        objtype, link, checked = prop_of(rid)
        cur = out.get(rid)
        if cur is None:
            out[rid] = {'rid': rid, 'anchor': anchor, 'objtype': objtype,
                        'link': link, 'checked': checked}
        elif cur['anchor'] is None and anchor is not None:
            cur['anchor'] = anchor
    # objectType 缺失时按“复选框”处理（VML/旧版控件可能不写该属性）
    ctrls = [c for c in out.values()
             if (c['objtype'] or 'CheckBox').replace(' ', '').lower() == 'checkbox']

    if not ctrls and parts['vml']:
        for shape in re.findall(r'<v:shape\b.*?</v:shape>', parts['vml'], re.S):
            if not re.search(r'ObjectType="Checkbox"', shape, re.I):
                continue
            a = re.search(r'<x:Anchor>([^<]+)</x:Anchor>', shape)
            link = re.search(r'<x:FmlaLink>([^<]+)</x:FmlaLink>', shape)
            anchor = None
            if a:
                nums = [int(x.strip()) for x in a.group(1).split(',') if x.strip().lstrip('-').isdigit()]
                if len(nums) >= 8:
                    anchor = {'fc': nums[0], 'fco': 0, 'fr': nums[2], 'fro': 0,
                              'tc': nums[4], 'tco': 0, 'tr': nums[6], 'tro': 0}
            ctrls.append({'rid': None, 'anchor': anchor, 'objtype': 'CheckBox',
                          'link': link.group(1) if link else None, 'checked': False})
    for c in ctrls:
        a = c['anchor']
        c['cell'] = (a['fr'] + 1, a['fc'] + 1) if a else None
    return ctrls
_RANGE_REF_RE = re.compile(r'\$?([A-Z]{1,3})\$?(\d+)\s*:\s*\$?([A-Z]{1,3})\$?(\d+)')
_CELL_REF_RE = re.compile(r'(?<![A-Z0-9_.$!])\$?([A-Z]{1,3})\$?(\d+)(?![\d(])')


def _col_idx(letter):
    n = 0
    for ch in letter:
        n = n * 26 + (ord(ch) - 64)
    return n


def _refs(formula):
    """从公式文本中解析引用：返回 [(起列, 止列, 起行, 止行), ...]（列为数字）。"""
    out = []
    up = (formula or '').upper()
    for m in _RANGE_REF_RE.finditer(up):
        out.append((_col_idx(m.group(1)), _col_idx(m.group(3)),
                    int(m.group(2)), int(m.group(4))))
    for m in _CELL_REF_RE.finditer(up):
        c = _col_idx(m.group(1))
        r = int(m.group(2))
        out.append((c, c, r, r))
    return out


def _covers_col(ref, col):
    c1, c2 = min(ref[0], ref[1]), max(ref[0], ref[1])
    return c1 <= col <= c2


def _label_dv(ws):
    """统计覆盖到标签单元格的“下拉/序列”数据验证：
       返回 (被覆盖的标签单元格数, 序列选项里是否内嵌了勾选符号)。"""
    n_cells = 0
    tick_in_list = False
    try:
        dvs = list(ws.data_validations.dataValidation)
    except Exception:
        dvs = []
    for dv in dvs:
        f1 = str(dv.formula1 or '')
        hit = 0
        try:
            rngs = list(dv.sqref.ranges)
        except Exception:
            rngs = []
        for rng in rngs:
            for r in range(rng.min_row, rng.max_row + 1):
                for c in range(rng.min_col, rng.max_col + 1):
                    if (r, c) in _LABEL_CELLS or (r, c) in _EMPTY_CELLS:
                        hit += 1
        if hit:
            n_cells += hit
            if any(ch in f1 for ch in _CHECKED_CHARS):
                tick_in_list = True
    return n_cells, tick_in_list


def _summary_block(ws):
    """定位右侧汇总区：在候选列 I~L 中，rows>=5 个“可计算”公式单元格的列视为汇总输出列。
       缺 _xlfn. 前缀的未来函数公式在 Excel 中恒为 #NAME?（或被 IFERROR 吞成空白），
       不产生任何汇总结果，故不计入。返回 {列号: [公式所在行, ...]}（按列升序）。"""
    cols = {}
    for c in SUM_COL_CANDIDATES:
        rows = [r for r in range(4, 201)
                if _formula(ws.cell(r, c)) and not _formula_broken(_formula(ws.cell(r, c)))]
        if len(rows) >= 5:
            cols[c] = rows
    return dict(sorted(cols.items()))


def _broken_summary_cells(ws):
    """候选汇总列中“写了公式但在 Excel 中必然算不出结果”的单元格列表。"""
    out = []
    for c in SUM_COL_CANDIDATES:
        for r in range(4, 201):
            t = _formula(ws.cell(r, c))
            if t and _formula_broken(t):
                out.append((r, c))
    return out


def _summary_profile(ws, sum_cols):
    """汇总区画像：每列的公式文本、是否引用维度列/标签列、引用到的标签行范围。"""
    prof = {}
    for c, rows in sum_cols.items():
        txts = [_formula(ws.cell(r, c)) or '' for r in rows]
        refs = []
        for t in txts:
            refs.extend(_refs(t))
        lab_refs = [x for x in refs
                    if any(_covers_col(x, k) for k in (DIM_COL,) + LABEL_COLS)]
        rmin = min((x[2] for x in lab_refs), default=None)
        rmax = max((x[3] for x in lab_refs), default=None)
        prof[c] = {
            'rows': rows,
            'txts': txts,
            'ref_dim': any(_covers_col(x, DIM_COL) for x in refs),
            'ref_label': any(_covers_col(x, k) for x in refs for k in LABEL_COLS),
            'row_min': rmin,
            'row_max': rmax,
        }
    return prof


def _summary_values(ws):
    """宏实现时右侧输出是运行期写入的静态值：在候选列 I~L 中找出与基线维度名/标签名
       匹配的文本单元格，返回 {列号: [行, ...]}（按列升序）。"""
    cols = {}
    for c in SUM_COL_CANDIDATES:
        rows = []
        for r in range(DATA_R1, 201):
            s = _txt(ws.cell(r, c).value).strip()
            if not s or _formula(ws.cell(r, c)):
                continue
            parts_ = [x.strip() for x in re.split(r'[：:]', s) if x.strip()]
            if (s in _DIM_NAMES or s in _ALL_LABELS
                    or any(x in _DIM_NAMES or x in _ALL_LABELS for x in parts_)):
                rows.append(r)
        if rows:
            cols[c] = rows
    return dict(sorted(cols.items()))


def _value_profile(ws, val_cols):
    """静态值汇总区画像：每列是维度列 / 标签列 / 同格「维度：标签」列，以及是否含末组标签。"""
    prof = {}
    for c, rows in val_cols.items():
        txts = [_txt(ws.cell(r, c).value).strip() for r in rows]
        segs = [[x.strip() for x in re.split(r'[：:]', t) if x.strip()] for t in txts]
        prof[c] = {
            'rows': rows,
            'txts': txts,
            'combined': any(len(g) >= 2 and any(x in _DIM_NAMES for x in g)
                            and any(x in _ALL_LABELS for x in g) for g in segs),
            'is_dim': all(t in _DIM_NAMES for t in txts),
            'is_lab': all(t in _ALL_LABELS for t in txts),
            'last_group': any(x in _LAST_GROUP_LABELS for g in segs for x in g),
        }
    return prof
# ============================================================================
# 对外接口
# ============================================================================
def evaluate(dir_path=None):
    """按接口约定返回结构化 dict；不 print、不改 sys.stdout、不 sys.exit。

    评估目标固定为本脚本所在目录 _SCRIPT_DIR：只在该目录内定位并打开被评估的
    .xlsx/.xlsm（忽略以 ~$ 打头的临时锁文件），不进入子目录、不检索上级目录；
    dir_path 参数仅为保持各脚本接口一致而保留，传入的值不会被使用。

    返回固定 9 个字段：id / file_name / status / error / dim1_pass / dim1_reason /
    dim2_items / total_score / max_score，其中 dim2_items 每项为
    {rule, max_delta, delta, hit}。

    两类异常分开回传：
      · 目录不存在 / 目录内没有 .xlsx/.xlsm——维度一根本没进入判定，error 只带纯文本
        原因（不带异常类名），file_name、dim1_pass、dim1_reason 一律留空（dim1_pass 为
        空字符串而非 False），dim2 不出项，总分与满分一并归零；
      · 其他运行期异常——error 以「脚本运行异常：」打头，dim1_pass 记 False 并在
        dim1_reason 中给出错误摘要，max_score 仍保留真实满分。
    """
    dir_path = dir_path if dir_path and os.path.isdir(dir_path) else _SCRIPT_DIR
    result = {
        "id": SCRIPT_ID,
        "file_name": "",
        "status": "ok",
        "error": None,
        "dim1_pass": False,
        "dim1_reason": "",
        "dim2_items": [],
        "total_score": 0,
        "max_score": MAX_SCORE_TOTAL,
    }
    try:
        if not dir_path or not os.path.isdir(dir_path):
            raise FileNotFoundError(f"目录不存在或不是目录：{dir_path}")
        cands = sorted(
            f for f in os.listdir(dir_path)
            if f.lower().endswith(('.xlsx', '.xlsm')) and not f.startswith('~$')
            and os.path.isfile(os.path.join(dir_path, f))
        )
        if not cands:
            raise FileNotFoundError(f"目录中未找到 .xlsx 文件：{dir_path}")
        file_path = os.path.join(dir_path, cands[0])
        result["file_name"] = os.path.basename(file_path)
        return _evaluate_impl(file_path, result)
    except FileNotFoundError as exc:
        # 找不到被评估文件：error 只回传纯文本原因，维度一未进入判定，各字段留空、分数归零。
        result["status"] = "error"
        result["error"] = str(exc)
        result["file_name"] = ""
        result["dim1_pass"] = ""
        result["dim1_reason"] = ""
        result["dim2_items"] = []
        result["total_score"] = 0
        result["max_score"] = 0
        return result
    except Exception as exc:
        result["status"] = "error"
        result["error"] = f"脚本运行异常：{type(exc).__name__}: {exc}"
        result["dim1_pass"] = False
        result["dim1_reason"] = f"文件无法打开或解析：{type(exc).__name__}: {exc}"
        result["dim2_items"] = []
        result["total_score"] = 0
        return result


def _fail_dim1(result, reason):
    """维度一不通过：完成度总分直接记 0，dim1_reason 带上错误摘要与计分后果。"""
    result["dim1_pass"] = False
    result["dim1_reason"] = f"维度一未通过：{reason}，完成度总分记 0"
    result["dim2_items"] = []
    result["total_score"] = 0
    return result


def _evaluate_impl(file_path, result):
    """维度一 + 维度二评估主体。加分项、减分项分别构造 "+N："/"-N：" 结果串，
       末尾按 _MAX_DELTAS_ORDER 反解成结构化 dim2_items。"""
    results = []  # 先加分项，后减分项

    # ---------- 维度1：可用与可修改性 ----------
    try:
        wb = load_workbook(file_path)
    except Exception as exc:
        return _fail_dim1(result, f"文件需修复才能打开或格式损坏（{type(exc).__name__}）")
    try:
        parts = _read_parts(file_path)
    except Exception as exc:
        return _fail_dim1(result, f"工作簿包结构损坏，无法读取（{type(exc).__name__}）")

    ws = wb[SHEET_NAME] if SHEET_NAME in wb.sheetnames else wb.worksheets[0]
    coldims = _col_dims(ws)
    mrs = _merged_ranges(ws)

    n_label_text = sum(1 for r in range(DATA_R1, DATA_R2 + 1) for c in LABEL_COLS
                       if _label_text(ws.cell(r, c).value))
    if n_label_text < 100:
        return _fail_dim1(result, f"左侧标签表可编辑文字大量缺失（仅 {n_label_text} 个标签单元格有文字），疑为图片粘贴或版面丢失")

    fake_formula = 0
    for r in range(1, 201):
        for c in range(1, 15):
            cell = ws.cell(r, c)
            v = cell.value
            if isinstance(v, str) and v.startswith('=') and getattr(cell, 'data_type', 'f') != 'f':
                fake_formula += 1
    if fake_formula > 5:
        return _fail_dim1(result, f"{fake_formula} 个单元格的公式以文本形式存放，打开后不计算")

    ctrls = _parse_checkbox_controls(parts)
    anchored = [c for c in ctrls if c['cell']]
    if len(anchored) >= 5 and len(set(c['cell'] for c in anchored)) < 0.6 * len(anchored):
        return _fail_dim1(result, "勾选控件大量堆叠在同一位置，版面不可用")

    result["dim1_pass"] = True
    # ---------- 维度二公共事实 ----------
    dv_cells, _ = _label_dv(ws)
    sum_cols = _summary_block(ws)
    broken_cells = _broken_summary_cells(ws)
    prof = _summary_profile(ws, sum_cols)
    # 文件本身可打开、可编辑（维度1 通过）；但公式能否算出结果影响维度2 的汇总类加分
    result["dim1_reason"] = (
        "文件可正常打开，左侧标签表与右侧汇总区均为可点选编辑的真实单元格/控件，公式正常计算"
        if not broken_cells else
        f"文件可正常打开、内容为可编辑的真实单元格，但右侧 {len(broken_cells)} 个公式的未来函数"
        f"缺 _xlfn. 前缀，在 Excel 中恒为 #NAME?（被 IFERROR 吞成空白），汇总按未实现计分")
    anchored_cells = [c['cell'] for c in ctrls if c['cell']]
    has_click_entry = bool(ctrls) or parts['activex'] > 0 or dv_cells > 0
    # 宏实现：右侧输出由工作簿内 VBA 在运行期写入，静态看不到公式。指令未限制实现手段，
    # 故“含宏 + 左侧有勾选入口”或“含宏 + 右侧已写下与基线标签匹配的汇总值”同样视为已实现。
    vba, vba_event, vba_manual = _vba_profile(parts)
    val_cols = _summary_values(ws) if not sum_cols else {}
    vprof = _value_profile(ws, val_cols) if vba else {}
    val_mode = bool(vprof)
    has_summary = bool(sum_cols) or (vba and (has_click_entry or val_mode))
    # 计分前置条件：两项核心功能均未实现 -> 加分项全部记 0
    core_done = has_click_entry or has_summary

    sum_ref_cols = set()
    for c in prof:
        for t in prof[c]['txts']:
            for ref in _refs(t):
                for k in range(min(ref[0], ref[1]), max(ref[0], ref[1]) + 1):
                    sum_ref_cols.add(k)
    combined_cols = [c for c in prof if prof[c]['ref_dim'] and prof[c]['ref_label']]
    split_dim_cols = [c for c in prof if prof[c]['ref_dim'] and not prof[c]['ref_label']]
    split_lab_cols = [c for c in prof if prof[c]['ref_label'] and not prof[c]['ref_dim']]
    first_out_row = min((min(prof[c]['rows']) for c in prof), default=None)
    last_out_row = max((max(prof[c]['rows']) for c in prof), default=None)

    # 缺 _xlfn. 前缀而算不出结果的公式虽不计加分，但仍物理占据版面，
    # 故 A10/A13 等“输出范围”类判断按物理存在的公式格取首末行。
    phys_rows = [r for c in prof for r in prof[c]['rows']] + [r for r, _ in broken_cells]
    phys_first = min(phys_rows, default=None)
    phys_last = max(phys_rows, default=None)

    # ======================= 加分项 =======================
    # 约定：加分项命中追加 "+N："，未命中追加 "+0："。

    # A01 +3：单击即可切换选中/取消（复选框类控件，非下拉菜单/手工输入符号）
    a01 = (any(cell in _LABEL_CELLS for cell in anchored_cells)
           or (not anchored_cells and bool(ctrls))
           or parts['activex'] > 0)
    a01 = a01 and core_done
    results.append(("+5" if a01 else "+0") + f"：B5:C68 单击即可切换勾选（复选框控件 {len(ctrls) + parts['activex']} 个）")

    # A02 +5：勾选入口完整覆盖 126 个有文字标签单元格，B32/B58 不设入口
    covered = set(anchored_cells)
    a02 = (core_done and covered == _LABEL_CELLS
           and len(anchored_cells) == len(covered) == 126)
    results.append(("+5" if a02 else "+0") + f"：勾选入口覆盖 126 个有文字标签单元格且无重复（实际覆盖 {len(covered)}）")

    # A03 +3：每个控件绑定独立状态单元格，勾选结果以 TRUE/FALSE 记录并可被公式引用
    links = [c['link'] for c in ctrls if c['link']]
    single = [x for x in links if re.match(r'^\$?[A-Z]{1,3}\$?\d+$', x.strip())]
    link_cols = set()
    for x in single:
        m = re.match(r'^\$?([A-Z]{1,3})\$?(\d+)$', x.strip())
        if m:
            link_cols.add(_col_idx(m.group(1)))
    bool_ok = True
    for x in single:
        cell = ws[x.replace('$', '')]
        if cell.value is not None and not isinstance(cell.value, bool):
            bool_ok = False
            break
    a03 = (core_done and bool(ctrls) and len(single) == len(ctrls)
           and len(set(single)) == len(single) and bool_ok
           and not (link_cols & ({DIM_COL} | set(LABEL_COLS)))
           and (not sum_ref_cols or link_cols <= sum_ref_cols))
    results.append(("+3" if a03 else "+0") + f"：控件绑定独立状态单元格并以 TRUE/FALSE 记录（{len(set(single))} 个互不重复的链接单元格）")

    # A04 +3：控件与标签单元格一一对齐、落在单元格边框内、不遮挡标签文字
    def _anchor_ok(a):
        if a['fc'] != a['tc']:
            return False
        r, c = a['fr'] + 1, a['fc'] + 1
        if (r, c) not in _LABEL_CELLS:
            return False
        if a['tr'] - a['fr'] > 1 or a['tr'] < a['fr']:
            return False
        if a['tr'] == a['fr'] + 1 and a['tro'] > 100000:
            return False
        if a['fco'] < 0:
            return False
        px = _col_w(ws, c, coldims) * 7 + 5
        return (a['tco'] - a['fco']) <= 0.6 * px * EMU_PER_PX
    aligned = [c for c in ctrls if c['anchor'] and _anchor_ok(c['anchor'])]
    a04 = core_done and bool(ctrls) and len(aligned) == len(ctrls)
    results.append(("+3" if a04 else "+0") + f"：控件与标签单元格一一对齐且落在边框内（{len(aligned)}/{len(ctrls)}）")

    # A05 +1：标签单元格中不再残留「□」等静态占位符号
    residual = [(r, c) for r in range(DATA_R1, DATA_R2 + 1) for c in LABEL_COLS
                if any(ch in _txt(ws.cell(r, c).value) for ch in _TICK_CHARS)]
    a05 = core_done and not residual
    results.append(("+1" if a05 else "+0") + f"：标签单元格无「□」等静态勾选符号残留（残留 {len(residual)} 处）")
    # A06 +5：I 列表头下方自动输出已勾选内容（勾选后立即刷新，不需手工录入/点按钮）
    n_sum_cells = sum(len(prof[c]['rows']) for c in prof)
    a06_formula = n_sum_cells >= 20 and any(prof[c]['ref_label'] for c in prof)
    # 宏实现：事件宏（勾选即触发）满足“自动”；只有明显是按钮驱动、又查不到事件过程时不满足
    a06_vba = bool(vba) and has_summary and not (vba_manual and not vba_event)
    a06 = core_done and (a06_formula or a06_vba)
    detail06 = f"（{n_sum_cells} 个可计算公式单元格）"
    if broken_cells:
        detail06 = (f"（{n_sum_cells} 个可计算公式单元格；另有 {len(broken_cells)} 个公式的未来函数"
                    f"缺 _xlfn. 前缀，在 Excel 中恒为 #NAME? 并被 IFERROR 吞成空白，汇总永不显示）")
    elif a06_vba and not a06_formula:
        detail06 = ("（由工作簿内 VBA 实现；%s）"
                    % ("已检出事件过程名，勾选即触发" if vba_event else "未能在 vbaProject.bin 中确认事件过程名，按自动刷新从宽计分"))
    results.append(("+5" if a06 else "+0") + "：右侧 I 列表头下方自动输出已勾选内容" + detail06)

    # A07 +3：汇总条目给出“分类维度 + 标签名称”的对应关系（同格或分两列）
    a07_f = bool(combined_cols) or (bool(split_dim_cols) and bool(split_lab_cols))
    a07_v = (any(vprof[c]['combined'] for c in vprof)
             or (any(vprof[c]['is_dim'] for c in vprof) and any(vprof[c]['is_lab'] for c in vprof)))
    a07 = core_done and (a07_f or a07_v)
    results.append(("+3" if a07 else "+0") + "：汇总条目给出“分类维度 + 标签名称”的对应关系")

    # A08 +3：结果连续排列、未勾选不产生空行、取消勾选后自动上移
    a08 = core_done and has_summary
    for c in prof:
        joined = ' '.join(prof[c]['txts']).upper()
        if not (any(f in joined for f in _RANK_FUNCS)
                and ('ROWS(' in joined or 'ROW()' in joined)):
            a08 = False
            break
    if val_mode and not prof:
        # 宏实现：按交付时可见输出判断——自表头下方首行起连续、中间无空行
        a08 = core_done and all(
            vprof[c]['rows'] == list(range(min(vprof[c]['rows']), max(vprof[c]['rows']) + 1))
            for c in vprof)
    results.append(("+3" if a08 else "+0") + "：汇总结果连续排列、无空行且取消勾选后自动上移")

    # A09 +3：汇总引用范围覆盖第 5~68 行全部维度分组
    rmin = min((prof[c]['row_min'] for c in prof if prof[c]['row_min']), default=None)
    rmax = max((prof[c]['row_max'] for c in prof if prof[c]['row_max']), default=None)
    a09 = core_done and rmin is not None and rmin <= DATA_R1 and rmax >= DATA_R2
    detail09 = f"（实际 {rmin}~{rmax}）"
    if val_mode and not prof:
        # 宏实现无公式可查：以“输出中出现末组（第 59~68 行）标签”作为覆盖到末行的直接证据
        a09 = core_done and any(vprof[c]['last_group'] for c in vprof)
        detail09 = ("（宏实现：输出中含第 59~68 行分组的标签，可确认覆盖至末行）" if a09
                    else "（宏实现：交付时的可见输出未出现第 59~68 行分组的标签，无法确认覆盖至末行）")
    results.append(("+3" if a09 else "+0") + "：汇总引用范围覆盖第 5~68 行" + detail09)

    # A10 +3：汇总区行数足以容纳 126 个标签，勾选到第 68 行时结果不被截断
    capacity = sum(len(prof[c]['rows']) for c in combined_cols)
    if split_lab_cols:
        capacity += max(len(prof[c]['rows']) for c in split_lab_cols)
    a10 = core_done and capacity >= 126
    detail10 = f"（容量 {capacity} 条）"
    if val_mode and not prof:
        # 宏实现：输出条数取决于运行时勾选数，改判“表头下方是否留出 >=126 行无占用空间”
        free = all(_txt(ws.cell(r, c).value).strip() == '' or r in vprof.get(c, {'rows': []})['rows']
                   for c in vprof for r in range(DATA_R1, DATA_R1 + 126))
        a10 = core_done and free
        detail10 = f"（宏实现：表头下方第 {DATA_R1}~{DATA_R1 + 125} 行%s被其他内容占用）" % ('未' if free else '已')
    results.append(("+3" if a10 else "+0") + "：汇总区可容纳全部 126 个标签" + detail10)

    # A11 +1：未勾选时汇总区为空白，不出现 #N/A、#VALUE!、0 等占位
    a11 = core_done and has_summary
    for c in prof:
        for t in prof[c]['txts']:
            up = t.upper()
            if not (any(g in up for g in _ERR_GUARDS) and '""' in t):
                a11 = False
                break
        if not a11:
            break
    if val_mode and not prof:
        # 宏实现：区域内不得残留错误值/占位字符
        junk = ('#N/A', '#VALUE!', '#DIV/0!', '#REF!', '#NAME?', '#NUM!', '#NULL!')
        a11 = core_done and not any(
            any(j in _txt(ws.cell(r, c).value).upper() for j in junk)
            for c in SUM_COL_CANDIDATES for r in range(DATA_R1, 201))
    results.append(("+1" if a11 else "+0") + "：未勾选时汇总区为空白，无 #N/A 等错误值")

    # A12 +1：汇总单元格自动换行或列宽足够，「维度：标签」可完整显示
    out_cols = sum_cols if sum_cols else val_cols
    out_prof = prof if prof else vprof
    a12 = core_done and has_summary
    if out_cols:
        last_col = max(out_cols)
        out_last = phys_last or max((max(out_prof[c]['rows']) for c in out_prof), default=4)
        spill_free = all(_txt(ws.cell(r, last_col + 1).value) == ''
                         for r in range(4, out_last + 1))
        for c in out_prof:
            cell = ws.cell(min(out_prof[c]['rows']), c)
            wrap = bool(cell.alignment and cell.alignment.wrap_text)
            wide = _col_w(ws, c, coldims) >= 18
            # 未换行且列宽不足时，只要右侧相邻列为空，文本仍可溢出完整显示
            if not (wrap or wide or (c == last_col and spill_free)):
                a12 = False
                break
    results.append(("+1" if a12 else "+0") + "：汇总单元格自动换行或列宽足够，文本不被截断")

    # A13 +1：汇总区表头沿用 115E59 深绿底 + 白色加粗 + 水平垂直居中
    a13 = False
    hdr_first = phys_first if phys_first else (min(min(vprof[c]['rows']) for c in vprof) if vprof else None)
    if has_summary and hdr_first and hdr_first > 4:
        hdrs = [ws.cell(r, c) for c in (out_cols or SUM_COL_CANDIDATES[:1]) for r in range(4, hdr_first)]
        hdrs = [x for x in hdrs if _txt(x.value).strip() or _fill6(x)]
        a13 = bool(hdrs) and all(
            _fill6(x) == _HDR_FILL and bool(x.font and x.font.b)
            and _white_or_theme(x) and _centered(x) for x in hdrs)
    a13 = a13 and core_done
    results.append(("+1" if a13 else "+0") + "：汇总区表头沿用 115E59 深绿底 + 白色加粗 + 居中")
    # A14 +3：勾选与汇总全部在单个工作簿内实现（控件+公式 或 工作簿内 VBA 事件宏均可），
    #         无需外部脚本、无需另存或重新打开、无需手动刷新
    a14 = core_done and has_click_entry and has_summary
    results.append(("+3" if a14 else "+0")
                   + "：勾选与汇总在单个工作簿内实现（控件+公式或工作簿内事件宏），无需外部脚本或手动刷新")

    # 说明：“交付文件为单个可直接使用的工作簿”已由维度一把关，此处不再重复计加分项。

    # ======================= 减分项 =======================
    # 约定：减分项触发（=指令未涉及却被改动）追加 "-N："，未触发追加 "-0："。

    # D01 -1：A1:C1 合并标题「用户标签维度表」16 磅加粗白字 / 0F766E / 居中 / 行高 30
    a1 = ws['A1']
    d01 = not (_is_merged_exact(mrs, 'A1:C1') and _txt(a1.value).strip() == _TITLE_TEXT
               and _sz(a1) == 16 and bool(a1.font and a1.font.b) and _white_or_theme(a1)
               and _fill6(a1) == '0F766E' and _centered(a1)
               and (_row_h(ws, 1) is not None and abs(_row_h(ws, 1) - 30) <= 0.6))
    results.append(("-1" if d01 else "-0") + "：A1:C1 合并标题行格式（16 磅加粗白字 / 0F766E / 居中 / 行高 30）")

    # D02 -1：A2:C2 浅色分隔条 CCFBF1、行高 22、无文字
    a2 = ws['A2']
    d02 = not (_is_merged_exact(mrs, 'A2:C2') and _txt(a2.value).strip() == ''
               and _fill6(a2) == 'CCFBF1'
               and (_row_h(ws, 2) is not None and abs(_row_h(ws, 2) - 22) <= 0.6))
    results.append(("-1" if d02 else "-0") + "：A2:C2 浅色分隔条格式（CCFBF1 / 行高 22 / 无文字）")

    # D03 -1：第 3 行被写入内容，或行高被加高至 22 以上
    row3_text = any(_txt(ws.cell(3, c).value).strip() for c in range(1, 4))
    h3 = _row_h(ws, 3)
    d03 = row3_text or (h3 is not None and h3 > 22)
    results.append(("-1" if d03 else "-0") + "：第 3 行被写入内容或行高被加高至 22 以上")

    # D04 -1：A4/B4/C4 表头文字与 11 磅加粗白字 / 115E59 / 居中 / 行高 24
    d04 = not (_row_h(ws, 4) is not None and abs(_row_h(ws, 4) - 24) <= 0.6)
    for c, text in _HDR_TEXTS.items():
        cell = ws.cell(4, c)
        if not (_txt(cell.value).strip() == text and _sz(cell) == 11
                and bool(cell.font and cell.font.b) and _white_or_theme(cell)
                and _fill6(cell) == _HDR_FILL and _centered(cell)):
            d04 = True
            break
    results.append(("-1" if d04 else "-0") + "：A4/B4/C4 表头文字与格式（11 磅加粗白字 / 115E59 / 居中 / 行高 24）")

    # D05 -3：B5:C68 标签文字被增删、改写或覆盖，126 个标签内容不完整
    bad_labels = []
    for r in range(DATA_R1, DATA_R2 + 1):
        for c in LABEL_COLS:
            if _label_text(ws.cell(r, c).value) != _LABELS[r][c - 2]:
                bad_labels.append((r, c))
    d05 = bool(bad_labels)
    results.append(("-3" if d05 else "-0") + f"：B5:C68 标签文字被增删改写（{len(bad_labels)} 处不符）")

    # 说明：勾选状态写在标签单元格自身（下拉选「☑ 新用户」替换文字）属“新增功能没做好”，
    #       其后果已由 A03（状态是否落在独立单元格）与 A05（是否残留静态符号）在加分项体现，
    #       不属于对既有内容的破坏，故不另计减分项。

    # D06 -1：A5:A68 维度列 10 磅加粗、居中、自动换行，及 5 个分组的分段填充色
    d06 = False
    for r1, r2, name, fill in _DIM_GROUPS:
        for r in range(r1, r2 + 1):
            cell = ws.cell(r, DIM_COL)
            if not (_sz(cell) == 10 and bool(cell.font and cell.font.b) and _centered(cell)
                    and bool(cell.alignment and cell.alignment.wrap_text)
                    and _fill6(cell) == fill):
                d06 = True
                break
        if d06:
            break
    results.append(("-1" if d06 else "-0") + "：A5:A68 维度列格式与 5 个维度分组的分段填充色")

    # D07 -1：B5:C68 标签单元格 10 磅、左对齐、自动换行、CBD5E1 细边框、交替填充色
    d07 = False
    for r in range(DATA_R1, DATA_R2 + 1):
        want_fill = _LABEL_FILL_ODD if r % 2 else _LABEL_FILL_EVEN
        for c in LABEL_COLS:
            cell = ws.cell(r, c)
            al = cell.alignment
            ls, lc = _side(cell, 'left')
            rs, rc = _side(cell, 'right')
            if not (_sz(cell) == 10 and bool(al) and al.horizontal == 'left'
                    and bool(al.wrap_text) and _fill6(cell) == want_fill
                    and ls == 'thin' and lc == _LABEL_BORDER_RGB
                    and rs == 'thin' and rc == _LABEL_BORDER_RGB):
                d07 = True
                break
        if d07:
            break
    results.append(("-1" if d07 else "-0") + "：B5:C68 标签单元格格式（10 磅 / 左对齐 / 换行 / CBD5E1 细边框 / 交替填充）")

    # D08 -1：第 5~68 行行高 22（未显式设置行高即回落到默认 14.25，同样计为不符）
    bad_h = [r for r in range(DATA_R1, DATA_R2 + 1)
             if _row_h(ws, r) is None or abs(_row_h(ws, r) - 22) > 0.6]
    d08 = bool(bad_h)
    results.append(("-1" if d08 else "-0")
                   + "：第 5~68 行行高 22（%d 行不符%s）"
                   % (len(bad_h), ("：" + ",".join(str(r) for r in bad_h[:8])) if bad_h else ""))

    # D09 -1：A5:C68 被插入/删除行列，导致维度分组与标签错行
    dim_shift = False
    for r1, r2, name, fill in _DIM_GROUPS:
        for r in range(r1, r2 + 1):
            if _txt(ws.cell(r, DIM_COL).value).strip() != name:
                dim_shift = True
                break
        if dim_shift:
            break
    spill_below = any(_txt(ws.cell(r, c).value).strip()
                      for r in range(DATA_R2 + 1, DATA_R2 + 41) for c in (1, 2, 3))
    d09 = dim_shift or spill_below or _has_merge(mrs, DATA_R1, 1, DATA_R2, 3)
    results.append(("-1" if d09 else "-0") + "：A5:C68 被插入/删除行列导致维度分组与标签错行")

    # D10 -1：A/B/C 列列宽被调到无法完整显示标签文字（自动换行下仅极窄列宽才会截断）
    d10 = False
    for c in (1, 2, 3):
        if _col_hidden(ws, c, coldims) or _col_w(ws, c, coldims) < 6:
            d10 = True
            break
    results.append(("-1" if d10 else "-0") + "：A/B/C 列列宽被调整到无法完整显示标签文字")

    # D11 -1：I2:J3 区域被写入文字或填充色（“操作说明/使用提示”等）
    d11 = _has_merge(mrs, 2, 9, 3, 10)
    for r in (2, 3):
        for c in (9, 10):
            cell = ws.cell(r, c)
            if _txt(cell.value).strip() or _fill6(cell):
                d11 = True
    results.append(("-1" if d11 else "-0") + "：I2:J3 空白区被写入文字或填充色")

    # D12 -1：I4 被与 J4 合并成跨列单元格
    d12 = _has_merge(mrs, 4, 9, 4, 9)
    results.append(("-1" if d12 else "-0") + "：I4 被合并成跨列单元格")

    # D13 -1：I4 填充色 115E59、白色 11 磅加粗、居中、上/左/右 CCCCCC 中粗边框
    i4 = ws['I4']
    i4_border = all(_side(i4, s) == ('medium', 'CCCCCC') for s in ('top', 'left', 'right'))
    d13 = not (_fill6(i4) == _HDR_FILL and _sz(i4) == 11
               and bool(i4.font and i4.font.b) and _white_or_theme(i4)
               and (i4.font.name in ('宋体', 'SimSun', 'NSimSun') if i4.font else False)
               and _centered(i4) and i4_border)
    results.append(("-1" if d13 else "-0") + "：I4 填充色/字号字体/居中/CCCCCC 中粗边框被改")

    # 说明：“需修复才能打开”已由维度一直接判 0 分，“既有样式大面积丢失”亦已由
    #       D01~D07（标题行/分隔条/表头/维度列/标签格式）逐项覆盖，此处不再重复计减分项。

    # D14 -1：工作表数量或名称发生变化
    d14 = wb.sheetnames != [SHEET_NAME]
    results.append(("-1" if d14 else "-0") + f"：工作表数量或名称变化（当前 {wb.sheetnames}）")

    # D15 -1：冻结窗格、网格线显示、工作表标签颜色等视图设置被改动
    tab_color = ws.sheet_properties.tabColor
    d15 = bool(ws.freeze_panes) or (ws.sheet_view.showGridLines is False) or bool(tab_color)
    results.append(("-1" if d15 else "-0") + "：冻结窗格/网格线/工作表标签颜色等视图设置被改动")

    # ---------- 反解：results（"+N："/"-N：" 串）-> 结构化 dim2_items ----------
    if len(results) != len(_MAX_DELTAS_ORDER):
        raise RuntimeError(
            f"结果条目数({len(results)})与 max_delta 序列({len(_MAX_DELTAS_ORDER)})不一致"
        )
    dim2_items = []
    total = 0
    for raw, max_delta in zip(results, _MAX_DELTAS_ORDER):
        head, _, rule = raw.partition("：")
        got = int(head.replace("+", "").replace("-", "").strip())  # 该项写入的绝对分值
        hit = got > 0
        # 加分项（max_delta>0）：命中记满分，否则 0
        # 减分项（max_delta<0）：触发“破坏”记该负分，否则 0
        delta = max_delta if hit else 0
        total += delta
        dim2_items.append({
            "rule": rule,
            "max_delta": max_delta,
            "delta": delta,
            "hit": hit,
        })

    result["dim2_items"] = dim2_items
    result["total_score"] = total
    result["max_score"] = MAX_SCORE_TOTAL
    return result


if __name__ == "__main__":
    # 本地调试入口：固定评估本脚本所在目录，不接受命令行参数。
    print(json.dumps(evaluate(), ensure_ascii=False, indent=2))
