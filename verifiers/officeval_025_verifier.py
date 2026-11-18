#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# pyright: reportUnknownMemberType=false, reportUnknownVariableType=false
# pyright: reportUnknownArgumentType=false, reportMissingImports=false
# pyright: reportPrivateUsage=false, reportGeneralTypeIssues=false, reportOptionalMemberAccess=false
"""025 号任务评分器：《青岭山居旅居营造计划》行距 / 错别字 / 标题格式 / 表格分页排版

评分口径（对齐同目录 细则.txt）：
  维度1 可用与可修改性：交付物为 1 个可正常打开的 .docx，正文与表格为可单独点选编辑的
        文字与表格（非整页位图），版面未严重错乱；不通过则完成度总分直接记 0。
  维度2 完成度：19 条加分项（满分 63）+ 23 条减分项（命中即扣，可累计）。
        加分项只覆盖指令要求且需要实际动手的改动；原文已达标、无需干活的项一律作为减分项，
        改坏即扣分。max_score = 加分项之和；total_score = 命中加分项 - 命中减分项（下限截到 0）。

说明：按接口统一约定，本模块不修改全局 sys.stdout，也不 print 主结果，也不调用 sys.exit。
      主结果统一由 evaluate() -> dict 返回；评估目标固定为“本脚本所在的那个目录”，
      只认该目录下的 .docx/.doc 交付物，不进入子目录、不检索上级目录、
      也不接受命令行或入参指定的其它路径。
表格分页判定优先走 Word COM 真实重排（win32com + Word），取每个表格首/末位置所在页码；
Word 不可用时退化为按 A4 可用高度 + 单元格文字行数的估算（仅内部记录，返回值固定为
      id / file_name / status / error / dim1_pass / dim1_reason / dim2_items /
      total_score / max_score 九个字段，不额外输出 notes）。
"""
import os
import re
import json
import zipfile
import collections

from docx import Document
from docx.oxml.ns import qn
from docx.table import Table
from docx.text.paragraph import Paragraph

try:
    import win32com.client as _wc  # Word COM：用于按真实重排结果判定表格跨页
except Exception:
    _wc = None

SCRIPT_ID = "025"
MAX_SCORE_TOTAL = 63

# 本脚本只评估“脚本自身所在的那一个目录”里的交付物：不进子目录、不看上级目录、
# 也不接受外部传入的其它路径。评测时把本脚本与被评 .docx 放在同一个目录下即可。
_SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

# ==== 输入文档基线常量（用于判定"指令未涉及、却被改动"的部分）====
STRICT_SONGTI = ('宋体', 'SimSun', 'NSimSun')
BASE_TABLE_COUNT = 28          # 表格数
BASE_TABLE_ROWS = 275          # 全部表格行数合计
BASE_NONEMPTY_PARAS = 281      # 非空段落数
BASE_SECTIONS = 3              # 分节数
BASE_PAGE_FIELDS = 2           # 页眉页脚中 PAGE 域片段数
BASE_DRAWING = 14              # <w:drawing> 图片数
BASE_PICT = 9                  # <w:pict> 图片数
BASE_SONGTI_ASCII = 190        # 非标题正文/表格中西文字体沿用宋体的 run 数
BASE_DOTTED_PCT_MIN = 55       # 小数点百分比写法应达到的条数（基线 59 处）
BASE_PAGE_MARGIN_LR = 1800     # 左右页边距 3.18 厘米 = 1800 twips
BASE_PAGE_MARGIN_TB = 1440     # 上下页边距 2.54 厘米 = 1440 twips
BASE_PAGE_W = 11906            # A4 宽 210 毫米
BASE_PAGE_H = 16838            # A4 高 297 毫米
BASE_EXTENTS = [
    (9187815, 13303250), (2736215, 1580515), (6818630, 1748790), (2716530, 1030605),
    (6297295, 1530350), (3558540, 1240155), (6778625, 2430780), (3169920, 833755),
    (7512050, 5302250), (2373630, 2373630), (3126105, 3118485), (5106670, 5261610),
    (5274310, 3515995), (5271135, 1637665),
]

# ==== 表头签名（用于定位各张表格）====
H_RES = ('主类', '亚类', '基本类型', '旅游资源单体名称')
H_LEDGER = ('月份', '重点工作', '责任主体', '交付物', '备注')
H_INNOV = ('创新维度', '具体体现', '不可复制性来源')
H_SURVEY = ('选项', '小计', '比例')
H_FUND = ('项目', '金额（万元）', '占比', '资金使用说明')
H_REV = ('业务线', '营收（万元）', '占比', '核心驱动因素')
H_COST = ('项目', '第一年（万元）', '第二年（万元）', '第三年（万元）', '说明')
H_RISK = ('序号', '阶段', '风险预测', '管理对策', '其他')
H_PRICE = ('业务类型', '目标客群', '业务特点', '价格范围', '典型例子')
H_TEAM = ('姓名', '专业班级', '学院', '职务', '核心职责')
H_COURT = ('院落类型', '空间特征与改造要点', '资源依托', '目标客群', '特色服务')
H_SEASON = ('季节', '核心体验项目', '资源依托', '特色亮点')

# ==== 标题分级（按核心标题文字匹配，容忍编号形式被改写）====
TITLE_L1 = ['市场分析', '项目简介', '产品与服务', '商业模式', '财务分析',
            '风险识别与管理对策', '未来前景分析', '团队介绍']
TITLE_L2 = ['环境分析', '旅游资源分析', '数据分析', '启动背景', '项目简述', '产品架构',
            '产品内容', '创新点分析', '商业模式', '营销策略', '财务规划',
            '项目可行性分析', '风险识别与评估', '风险管理对策', '前期', '中期', '后期']
TITLE_L3 = ['政治环境（Politics）', '经济环境（Economy）', '社会环境（Society）',
            '技术环境（Technology）', '样本构成分析', '旅居市场现状与潜力',
            '消费能力与预算结构', '旅居地选择因素分析', '住宿形式偏好', '服务需求偏好',
            '决策行为分析', '开放题内容分析', '分年龄段交叉分析（补充）', '主要调研结论汇总',
            '政策层面：', '省域层面：', '市级层面：', '青岭湾村现状', '目标人群',
            '产品介绍', '项目定位', '产品核心：山居特色院落群', '四大体验板块',
            '三个阶段', '多方利益分配机制', '扁平化组织架构', '线上获客（核心渠道）',
            '线下获客（补充渠道）', '财务规划原则', '阶段发展假设', '启动资金预算与使用规划',
            '定价体系与业务线梳理', '三年收入预测（按业务线拆分）', '年度总成本与费用表',
            '投资回报与现金流测算', '政策可行性', '市场可行性', '经济可行性', '运营可行性']
TITLE_L3_ALIAS = {'政策层面': '政策层面：', '省域层面': '省域层面：', '市级层面': '市级层面：'}

# 四级编号小标题（与正文同级，应为宋体小四不加粗）
LV4_NUMBERED = ['地域分布', '年龄结构', '旅居经历', '对山地村落旅居项目的兴趣度',
                '每月旅居预算（含住宿+基础餐饮）', '最看重的三项因素（多选）',
                '倾向的住宿形式', '最有吸引力的服务（多选）', '决定长住前的准备行为（多选）',
                '关键词分类统计', '预算与年龄的交叉', '市场基本面']
# 冒号型小标题与提示行（应为宋体小四不加粗）
LV4_COLON = ['宏观市场规模：', '客群地域来源：', '用户画像：', '核心：山地院落民宿群',
             '四大体验板块：', '核心定位：', '定位阐释：', '投资回收期测算', '现金流预测',
             '三大核心因素清晰可见：', '服务需求呈现清晰层次：', '现金流保障措施：']
BOLD_HINT = ['成员分工：', '指导教师：']

MONTH_LABELS = ['%s月·第%d轮' % (m, k)
                for m in ['一', '二', '三', '四', '五', '六', '七', '八', '九', '十', '十一', '十二']
                for k in (1, 2, 3)]
INNOV_NAMES = ['资源深度转化', '主客共创', '弹性组合设计', '小体量运营', '在地物产增值', '差异化场景']
CELL_NOTES = ['跨页表格，便于季度复盘', '院落巡检、活动复盘、农产品清单校对', '低于警戒值暂停扩容']
KEEP_NUMBERS = ['165', '128', '3.2', '19公里', '28分钟', '32%', '61%', '199', '172',
                '300', '800', '1450']
KEEP_NAMES = ['陆修远', '何修晴', '周一行', '赵清澜', '林若楠', '许知瑾', '梁冥香',
              '旅游管理23级1班']

# ============================ 基础工具 ============================

def _norm(s):
    """归一化文字：去掉空白与换行，全角空格一并去掉。"""
    return re.sub(r'[\s　]+', '', s or '')


def _style_chain(style):
    """沿 w:basedOn 链自底向上返回样式列表。"""
    out, guard = [], 0
    while style is not None and guard < 24:
        out.append(style)
        guard += 1
        try:
            style = style.base_style
        except Exception:
            style = None
    return out


def _rpr_attr(el, tag, attr):
    rPr = el.find(qn('w:rPr'))
    if rPr is None:
        return None
    node = rPr.find(qn(tag))
    if node is None:
        return None
    return node.get(qn('w:' + attr)) if attr else ''


def _docdefault_rpr(doc, tag, attr):
    dd = doc.styles.element.find(qn('w:docDefaults'))
    if dd is None:
        return None
    node = dd.find(qn('w:rPrDefault') + '/' + qn('w:rPr') + '/' + qn(tag))
    if node is None:
        return None
    return node.get(qn('w:' + attr)) if attr else ''


def _eff_run_attr(doc, run, tag, attr):
    """run 直接属性 -> 段落样式链 -> docDefaults，逐级回退取有效值。"""
    v = _rpr_attr(run._element, tag, attr)
    if v:
        return v
    try:
        base = run._parent.style
    except Exception:
        base = None
    for st in _style_chain(base):
        v = _rpr_attr(st.element, tag, attr)
        if v:
            return v
    return _docdefault_rpr(doc, tag, attr)


def _eff_size(doc, run):
    """有效字号（磅）。"""
    v = _eff_run_attr(doc, run, 'w:sz', 'val')
    try:
        return int(v) / 2.0
    except (TypeError, ValueError):
        return None

def _eff_bold(doc, run):
    """有效加粗：w:b 存在且 val 不为 0/false 视为加粗。"""
    for el in [run._element] + [st.element for st in _style_chain(_safe_style(run))]:
        rPr = el.find(qn('w:rPr'))
        if rPr is None:
            continue
        node = rPr.find(qn('w:b'))
        if node is not None:
            return node.get(qn('w:val')) not in ('0', 'false', 'off')
    return False


def _safe_style(run):
    try:
        return run._parent.style
    except Exception:
        return None


def _eff_eastasia(doc, run):
    return _eff_run_attr(doc, run, 'w:rFonts', 'eastAsia')


def _eff_ascii(doc, run):
    for attr in ('ascii', 'hAnsi'):
        v = _rpr_attr(run._element, 'w:rFonts', attr)
        if v:
            return v
    for st in _style_chain(_safe_style(run)):
        for attr in ('ascii', 'hAnsi'):
            v = _rpr_attr(st.element, 'w:rFonts', attr)
            if v:
                return v
    for attr in ('ascii', 'hAnsi'):
        v = _docdefault_rpr(doc, 'w:rFonts', attr)
        if v:
            return v
    return None


def _ppr_attr(el, tag, attr):
    pPr = el.find(qn('w:pPr'))
    if pPr is None:
        return None
    node = pPr.find(qn(tag))
    if node is None:
        return None
    return node.get(qn('w:' + attr))


def _eff_para_attr(doc, para, tag, attr):
    v = _ppr_attr(para._element, tag, attr)
    if v is not None:
        return v
    try:
        base = para.style
    except Exception:
        base = None
    for st in _style_chain(base):
        v = _ppr_attr(st.element, tag, attr)
        if v is not None:
            return v
    dd = doc.styles.element.find(qn('w:docDefaults'))
    if dd is not None:
        pd = dd.find(qn('w:pPrDefault') + '/' + qn('w:pPr'))
        if pd is not None:
            node = pd.find(qn(tag))
            if node is not None:
                return node.get(qn('w:' + attr))
    return None

def _line_mult(doc, para):
    """有效行距倍数；返回 (倍数, 规则)。规则为 auto 时倍数 = w:line / 240。"""
    rule = _eff_para_attr(doc, para, 'w:spacing', 'lineRule')
    line = _eff_para_attr(doc, para, 'w:spacing', 'line')
    if line is None:
        return (1.0, rule or 'auto')
    try:
        line = int(line)
    except ValueError:
        return (None, rule)
    if rule in (None, 'auto'):
        return (line / 240.0, 'auto')
    return (line / 240.0, rule)


def _is_mult(doc, para, target, tol=0.06):
    mult, rule = _line_mult(doc, para)
    return rule == 'auto' and mult is not None and abs(mult - target) <= tol


def _runs(para):
    return [r for r in para.runs if r.text.strip()]


def _dedup_cells(row):
    """按 tc 去重，横向合并单元格只算一次。"""
    seen, out = set(), []
    try:
        cells = row.cells
    except Exception:
        return out
    for c in cells:
        if id(c._tc) in seen:
            continue
        seen.add(id(c._tc))
        out.append(c)
    return out


def _core_title(text):
    """剥掉各种编号前缀，得到核心标题文字。"""
    pat = re.compile(r'^\s*(?:[（(][一二三四五六七八九十\d]+[）)]'
                     r'|[一二三四五六七八九十]+[、.．]'
                     r'|\d+(?:[.．]\d+)*[、.．]?)\s*')
    s = (text or '').strip()
    prev = None
    while s != prev:
        prev = s
        s = pat.sub('', s, count=1).strip()
    return _norm(s)


def _has_visible_border(doc, tbl_el):
    """表格是否呈现可见边框：显式 tblBorders 优先，其次看表格样式名。"""
    sides = ('top', 'left', 'bottom', 'right', 'insideH', 'insideV')
    tblPr = tbl_el.find(qn('w:tblPr'))
    if tblPr is not None:
        bd = tblPr.find(qn('w:tblBorders'))
        if bd is not None:
            for node in bd:
                if node.tag.split('}')[-1] not in sides:
                    continue
                if node.get(qn('w:val')) not in (None, 'nil', 'none'):
                    return True
            return False
    name = ''
    try:
        name = Table(tbl_el, doc).style.name or ''
    except Exception:
        name = ''
    return ('Grid' in name) or ('网格' in name)

def _table_cells(tb):
    """整表去重后的单元格列表（横向、纵向合并只算一次）。"""
    seen, out = set(), []
    try:
        rows = list(tb.rows)
    except Exception:
        return out
    for row in rows:
        for c in _dedup_cells(row):
            if id(c._tc) in seen:
                continue
            seen.add(id(c._tc))
            out.append(c)
    return out


def _grid_cols(tbl_el):
    g = tbl_el.find(qn('w:tblGrid'))
    if g is None:
        return []
    out = []
    for c in g.findall(qn('w:gridCol')):
        try:
            out.append(int(c.get(qn('w:w')) or 0))
        except ValueError:
            out.append(0)
    return out


def _table_pages_by_word(file_path):
    """用 Word COM 重排后取每个表格首/末所在页码，返回 [(起页, 止页), ...] 或 None。"""
    if _wc is None:
        return None
    app = doc = None
    try:
        app = _wc.DispatchEx('Word.Application')
        app.Visible = False
        app.DisplayAlerts = 0
        doc = app.Documents.Open(os.path.abspath(file_path), ReadOnly=True,
                                 AddToRecentFiles=False)
        spans = []
        for i in range(1, int(doc.Tables.Count) + 1):
            rng = doc.Tables(i).Range
            start = int(doc.Range(rng.Start, rng.Start).Information(3))
            end = int(doc.Range(rng.End, rng.End).Information(3))
            spans.append((start, end))
        return spans
    except Exception:
        return None
    finally:
        try:
            if doc is not None:
                doc.Close(False)
        except Exception:
            pass
        try:
            if app is not None:
                app.Quit()
        except Exception:
            pass


def _estimate_spans(ctx):
    """Word 不可用时的退化估算：按 A4 可用高度与单元格文字行数推算表格占几页。"""
    sec = ctx['doc'].sections[0]
    usable_pt = (sec.page_height.pt - sec.top_margin.pt - sec.bottom_margin.pt)
    spans = []
    for info in ctx['tables']:
        cols = [w for w in info['gridcols'] if w >= 170] or [8300]
        total_pt = 0.0
        for row in info['rowcells']:
            maxline = 1
            for idx, cell in enumerate(row):
                width_pt = cols[min(idx, len(cols) - 1)] / 20.0 - 10.8
                sz = info['maxsize'] or 12.0
                per = max(1.0, width_pt / max(1.0, sz))
                txt = _norm(cell.text)
                need = max(1, int(len(txt) / per + 0.999))
                maxline = max(maxline, need)
            total_pt += maxline * (info['maxsize'] or 12.0) * 1.25 * 1.2 + 2.0
        spans.append((1, max(1, int(total_pt / usable_pt + 0.999))))
    return [(1, s[1]) for s in spans]

# ============================ 上下文构建 ============================

def _build_ctx(file_path):
    doc = Document(file_path)
    ctx = {'doc': doc, 'file_path': file_path, 'notes': []}

    # --- 表格信息 ---
    tables = []
    for child in doc.element.body.iterchildren():
        if child.tag != qn('w:tbl'):
            continue
        tb = Table(child, doc)
        try:
            rows = list(tb.rows)
        except Exception:
            rows = []
        rowcells = [_dedup_cells(r) for r in rows]
        header = tuple(_norm(c.text) for c in (rowcells[0] if rowcells else []))
        cells = _table_cells(tb)
        cparas, cruns, sizes = [], [], []
        for c in cells:
            for p in c.paragraphs:
                cparas.append(p)
                for r in _runs(p):
                    cruns.append(r)
                    sz = _eff_size(doc, r)
                    if sz:
                        sizes.append(sz)
        trPr0 = rows[0]._tr.find(qn('w:trPr')) if rows else None
        tables.append({
            'el': child, 'tb': tb, 'nrows': len(rows), 'header': header,
            'rowcells': rowcells, 'cells': cells, 'cparas': cparas, 'cruns': cruns,
            'sizes': sizes, 'minsize': min(sizes) if sizes else None,
            'maxsize': max(sizes) if sizes else None,
            'gridcols': _grid_cols(child),
            'bordered': _has_visible_border(doc, child),
            'tblheader': trPr0 is not None and trPr0.find(qn('w:tblHeader')) is not None,
            'cantsplit': len(child.findall('.//' + qn('w:cantSplit'))),
        })
    ctx['tables'] = tables
    ctx['total_rows'] = sum(t['nrows'] for t in tables)

    # --- 标题分级（按核心文字有序匹配，容忍编号被改写）---
    cmap = collections.defaultdict(list)
    for lv, pool in ((1, TITLE_L1), (2, TITLE_L2), (3, TITLE_L3)):
        for name in pool:
            cmap[_norm(name)].append(lv)
    for alias, real in TITLE_L3_ALIAS.items():
        if _norm(alias) not in cmap:
            cmap[_norm(alias)].append(3)
    heads = {1: [], 2: [], 3: []}
    hset = set()
    used = collections.Counter()
    for p in doc.paragraphs:
        t = p.text.strip()
        if not t or len(t) > 40 or t.endswith('。'):
            continue
        c = _core_title(t)
        if c not in cmap:
            continue
        lvls = cmap[c]
        lv = lvls[min(used[c], len(lvls) - 1)]
        used[c] += 1
        heads[lv].append(p)
        hset.add(id(p._element))
    ctx['heads'] = heads
    ctx['hset'] = hset
    # --- 段落分类 ---
    lv4_pool = [_norm(x) for x in LV4_NUMBERED]
    colon_pool = [_norm(x) for x in LV4_COLON]
    bold_pool = [_norm(x) for x in BOLD_HINT]
    narrative, lv4num, lv4colon, boldhint, nonempty = [], [], [], [], []
    for p in doc.paragraphs:
        t = p.text.strip()
        if not t:
            continue
        nonempty.append(p)
        if id(p._element) in hset:
            continue
        n = _norm(t)
        if len(n) <= 40 and any(n.endswith(k) for k in lv4_pool):
            lv4num.append(p)
        elif len(n) <= 40 and any(n.endswith(k) for k in colon_pool):
            lv4colon.append(p)
        elif len(n) <= 40 and any(n.endswith(k) for k in bold_pool):
            boldhint.append(p)
        elif len(t) >= 25 and not t.endswith('：'):
            narrative.append(p)
    ctx['narrative'] = narrative
    ctx['lv4num'] = lv4num
    ctx['lv4colon'] = lv4colon
    ctx['boldhint'] = boldhint
    ctx['nonempty'] = nonempty

    # --- 全文文字（正文 + 表格）---
    parts = [p.text for p in doc.paragraphs]
    for info in ctx['tables']:
        for c in info['cells']:
            parts.append(c.text)
    ctx['text'] = '\n'.join(parts)

    # --- 图片、分节、页眉页脚 ---
    body_xml = doc.element.body.xml
    ctx['n_drawing'] = body_xml.count('<w:drawing>')
    ctx['n_pict'] = body_xml.count('<w:pict>')
    ctx['extents'] = [(int(a), int(b)) for a, b in
                      re.findall(r'<wp:extent cx="(\d+)" cy="(\d+)"', body_xml)]
    hf_page, hf_text = 0, []
    for s in doc.sections:
        for part in (s.header, s.footer):
            hf_page += len(re.findall(r'\bPAGE\b', part._element.xml))
        hf_text += [p.text.strip() for p in s.header.paragraphs if p.text.strip()]
        hf_text += [p.text.strip() for p in s.footer.paragraphs if p.text.strip()]
    ctx['hf_page_fields'] = hf_page
    ctx['hf_text'] = hf_text
    ctx['n_sections'] = len(doc.sections)

    # --- 表格分页跨页数 ---
    spans = _table_pages_by_word(file_path)
    if spans is not None and len(spans) == len(tables):
        ctx['span_source'] = 'word'
    else:
        spans = _estimate_spans(ctx)
        ctx['span_source'] = 'estimate'
        ctx['notes'].append('Word 不可用，表格跨页数由 A4 可用高度估算')
    for info, sp in zip(tables, spans):
        info['page_from'], info['page_to'] = sp
        info['span'] = max(1, sp[1] - sp[0] + 1)
    return ctx

# ============================ 表格定位 ============================

def _hp(header, sig):
    """表头前若干列是否与签名一致。"""
    sig = tuple(_norm(x) for x in sig)
    return tuple(header[:len(sig)]) == sig


def _find(ctx, sig):
    return [t for t in ctx['tables'] if _hp(t['header'], sig)]


def _res_tables(ctx):
    """旅游资源表（含被拆分后失去表头的续表）。"""
    out = []
    for t in ctx['tables']:
        h = t['header']
        if _hp(h, H_RES):
            out.append(t)
        elif (len(h) == 4 and t['nrows'] >= 10
              and re.match(r'^[A-F][一-龥]', h[0] or '')):
            out.append(t)
    return out


def _age_court_team(ctx):
    out = []
    for t in ctx['tables']:
        h = t['header']
        if not h:
            continue
        if h[0] == '年龄段' or _hp(h, H_COURT) or _hp(h, H_SEASON) or _hp(h, H_TEAM):
            out.append(t)
    return out


def _grid_group(ctx):
    out = []
    for t in ctx['tables']:
        h = t['header']
        if not h:
            continue
        if (_hp(h, H_SURVEY) or _hp(h, H_FUND) or _hp(h, H_REV)
                or (h[0] == '年度' and len(h) > 1 and '期初现金' in h[1])):
            out.append(t)
    return out


# ============================ 维度1 ============================

# 交付物只接受 .docx：细则要求交付「1 个可正常打开的 Word 文档（.docx）」，其余
# 格式（.doc / .docm / .rtf / .odt / .pdf 等）一律在维度一判不通过。
_ACCEPTED_EXT = '.docx'
# 目录里没有 .docx 时，这些扩展名仍收作被评对象，交给维度一判不通过（总分记 0），
# 而不是当成「没有交付物」直接跳过评测。
_WRONG_FMT_EXTS = ('.doc', '.docm', '.dotx', '.dotm', '.dot', '.rtf', '.odt',
                   '.wps', '.pages', '.pdf')


def _fmt_reject(file_path):
    """维度1的格式门禁：交付物必须是可正常打开的 .docx（OOXML 文档包）。

    不通过返回原因文本，通过返回 None。判定与 Word 的“能不能双击打开”对齐：
    扩展名须为 .docx，文件须是 zip 容器，且包内含 word/document.xml。
    """
    ext = os.path.splitext(file_path)[1].lower()
    if ext != _ACCEPTED_EXT:
        return '交付物为 %s 格式，只接受 .docx 文档' % (ext or '无扩展名')
    if not zipfile.is_zipfile(file_path):
        return '交付物不是有效的 .docx 包（无法作为 zip 容器打开）'
    try:
        with zipfile.ZipFile(file_path) as zf:
            names = set(zf.namelist())
    except Exception as exc:
        return '交付物无法作为 .docx 打开：%s: %s' % (type(exc).__name__, exc)
    if 'word/document.xml' not in names:
        return '交付物包内缺少 word/document.xml，不是 Word 文档'
    return None


def _dim1(ctx):
    n_para = len(ctx['nonempty'])
    n_tbl = len(ctx['tables'])
    n_char = len(_norm(ctx['text']))
    if n_para < 60:
        return False, '可单独点选的正文段落仅 %d 个，正文疑似被转为位图或大面积丢失' % n_para
    if n_tbl < 5:
        return False, '可编辑表格仅 %d 张（基线 28 张），表格疑似被转为图片或大量丢失' % n_tbl
    if n_char < 5000:
        return False, '可提取文字仅 %d 字，内容疑似严重缺失' % n_char
    for i, t in enumerate(ctx['tables']):
        if t['nrows'] == 0:
            return False, '第 %d 张表格行结构无法解析，版面严重错乱' % (i + 1)
    return True, ''

# ============================ 加分项 ============================

def _p01(ctx):
    """全部正文叙述段落为单倍行距。"""
    ps = ctx['narrative']
    if len(ps) < 120:
        return False
    ok = sum(1 for p in ps if _is_mult(ctx['doc'], p, 1.0))
    return ok >= len(ps) * 0.95


def _p02(ctx):
    """提示行与说明段落同样单倍，全文无遗漏的 1.5 倍行距正文段落。"""
    d = ctx['doc']
    hints = ctx['lv4num'] + ctx['lv4colon'] + ctx['boldhint']
    if len(hints) < 20:
        return False
    if any(not _is_mult(d, p, 1.0) for p in hints):
        return False
    lv2 = {id(p._element) for p in ctx['heads'][2]}
    for p in ctx['nonempty']:
        if id(p._element) in lv2:
            continue
        mult, rule = _line_mult(d, p)
        if rule == 'auto' and mult is not None and mult >= 1.4:
            return False
    return True


def _p03(ctx):
    """全部表格单元格文字为 1.25 倍行距。"""
    d, tot, ok = ctx['doc'], 0, 0
    for t in ctx['tables']:
        for p in t['cparas']:
            if not p.text.strip():
                continue
            tot += 1
            if _is_mult(d, p, 1.25, 0.05):
                ok += 1
    return tot > 0 and ok >= tot * 0.98


def _check_heads(ctx, level, count, size=None, bold=None, songti=True):
    d, hs = ctx['doc'], ctx['heads'][level]
    if len(hs) != count:
        return False
    for p in hs:
        rs = _runs(p)
        if not rs:
            return False
        for r in rs:
            if songti and (_eff_eastasia(d, r) or '') not in STRICT_SONGTI:
                return False
            if size is not None:
                sz = _eff_size(d, r)
                if sz is None or abs(sz - size) > 0.26:
                    return False
            if bold is not None and _eff_bold(d, r) != bold:
                return False
    return True


def _head_sizes(ctx, level):
    d, out = ctx['doc'], set()
    for p in ctx['heads'][level]:
        for r in _runs(p):
            sz = _eff_size(d, r)
            if sz:
                out.add(round(sz, 1))
    return out


def _p04(ctx):
    """8 处一级标题为宋体、小三（15磅）。"""
    return _check_heads(ctx, 1, 8, size=15.0)


def _p05(ctx):
    """8 处一级标题为单倍行距。"""
    hs = ctx['heads'][1]
    return len(hs) == 8 and all(_is_mult(ctx['doc'], p, 1.0) for p in hs)


def _p06(ctx):
    """8 处一级标题段前 0.5 行、段后 0.5 行（按行设置）。"""
    d, hs = ctx['doc'], ctx['heads'][1]
    if len(hs) != 8:
        return False
    for p in hs:
        before = _eff_para_attr(d, p, 'w:spacing', 'beforeLines')
        after = _eff_para_attr(d, p, 'w:spacing', 'afterLines')
        try:
            if abs(int(before) - 50) > 2 or abs(int(after) - 50) > 2:
                return False
        except (TypeError, ValueError):
            return False
    return True


def _p07(ctx):
    """17 处二级标题为宋体、四号（14磅）。"""
    return _check_heads(ctx, 2, 17, size=14.0)


def _p08(ctx):
    """39 处三级标题为宋体、小四（12磅）。"""
    return _check_heads(ctx, 3, 39, size=12.0)


def _p09(ctx):
    """三级标题为加粗。"""
    return _check_heads(ctx, 3, 39, bold=True, songti=False)


def _p10(ctx):
    """三级标题为单倍行距。"""
    hs = ctx['heads'][3]
    return len(hs) == 39 and all(_is_mult(ctx['doc'], p, 1.0) for p in hs)


def _p11(ctx):
    """正文与表格中的"运盈"均改为"运营"。"""
    t = ctx['text']
    return '运盈' not in t and '运营重点' in t and '小体量运营' in t


def _p12(ctx):
    """"测祘"改为"测算"。"""
    t = ctx['text']
    return '测祘' not in t and '测算口径' in t


def _p13(ctx):
    """以顿号充当小数点的百分比均改为小数点写法。"""
    t = ctx['text']
    dun = re.findall(r'\d+、\d+\s*%', t)
    dot = re.findall(r'\d+\.\d+\s*%', t)
    return not dun and len(dot) >= BASE_DOTTED_PCT_MIN


def _p14(ctx):
    """资源表为一个连续表格（36行），跨两页排布。"""
    rs = _res_tables(ctx)
    return len(rs) == 1 and rs[0]['nrows'] == 36 and rs[0]['span'] == 2


def _p15(ctx):
    """资源表在续排处重复显示表头行。"""
    rs = _res_tables(ctx)
    return len(rs) == 1 and rs[0]['tblheader']

def _p16(ctx):
    """月度台账 37 行与创新维度 7 行内容无删减，台账完整排布在三页上。"""
    led, inn = _find(ctx, H_LEDGER), _find(ctx, H_INNOV)
    if len(led) != 1 or len(inn) != 1:
        return False
    if led[0]['nrows'] != 37 or inn[0]['nrows'] != 7:
        return False
    t = ctx['text']
    if any(m not in t for m in MONTH_LABELS):
        return False
    if any(n not in t for n in INNOV_NAMES):
        return False
    return led[0]['span'] == 3


def _p17(ctx):
    """月度台账 5 列、创新维度 3 列，不再共用同一组列宽产生空列错位。"""
    led, inn = _find(ctx, H_LEDGER), _find(ctx, H_INNOV)
    if len(led) != 1 or len(inn) != 1:
        return False
    vl = [w for w in led[0]['gridcols'] if w >= 170]
    vi = [w for w in inn[0]['gridcols'] if w >= 170]
    return (len(vl) == 5 and len(vi) == 3
            and len(led[0]['header']) == 5 and len(inn[0]['header']) == 3)


def _p18(ctx):
    """除资源表与创新点分析下方两表外，其余表格均完整落在同一页内。"""
    exempt = {id(t['el']) for t in _res_tables(ctx) + _find(ctx, H_LEDGER) + _find(ctx, H_INNOV)}
    others = [t for t in ctx['tables'] if id(t['el']) not in exempt]
    return bool(others) and all(t['span'] == 1 for t in others)


def _p19(ctx):
    """采用跨页断行/重复标题行/分页位置等排版手段，未删减表格行列或内容。"""
    if len(ctx['tables']) != BASE_TABLE_COUNT or ctx['total_rows'] != BASE_TABLE_ROWS:
        return False
    rs, led, inn = _res_tables(ctx), _find(ctx, H_LEDGER), _find(ctx, H_INNOV)
    if len(rs) != 1 or rs[0]['nrows'] != 36:
        return False
    if len(led) != 1 or led[0]['nrows'] != 37 or len(inn) != 1 or inn[0]['nrows'] != 7:
        return False
    n_th = sum(1 for t in ctx['tables'] if t['tblheader'])
    n_cs = sum(t['cantsplit'] for t in ctx['tables'])
    n_keep = len(ctx['doc'].element.body.findall('.//' + qn('w:keepNext')))
    return n_cs > 0 or n_th >= 5 or n_keep >= 600


# ============================ 减分项 ============================

def _d01(ctx):
    """左右页边距不满足 3.18 厘米。"""
    for s in ctx['doc'].sections:
        for v in (s.left_margin, s.right_margin):
            if v is None or abs(int(v.twips) - BASE_PAGE_MARGIN_LR) > 30:
                return True
    return False


def _d02(ctx):
    """纸张与上下页边距不满足 A4、上下 2.54 厘米。"""
    for s in ctx['doc'].sections:
        if s.page_width is None or abs(int(s.page_width.twips) - BASE_PAGE_W) > 40:
            return True
        if s.page_height is None or abs(int(s.page_height.twips) - BASE_PAGE_H) > 40:
            return True
        for v in (s.top_margin, s.bottom_margin):
            if v is None or abs(int(v.twips) - BASE_PAGE_MARGIN_TB) > 30:
                return True
    return False


def _d03(ctx):
    """17 处二级标题不满足 1.5 倍行距（原文已达标，被改动即为破坏）。"""
    hs = ctx['heads'][2]
    if len(hs) != 17:
        return True
    return any(not _is_mult(ctx['doc'], p, 1.5, 0.05) for p in hs)


def _small_title_bad(ctx, paras, size, bold, least, font=True):
    d = ctx['doc']
    if len(paras) < least:
        return True
    for p in paras:
        for r in _runs(p):
            sz = _eff_size(d, r)
            if sz is None or abs(sz - size) > 0.26:
                return True
            if _eff_bold(d, r) != bold:
                return True
            if font and (_eff_eastasia(d, r) or '') not in STRICT_SONGTI:
                return True
    return False


def _d04(ctx):
    """四级编号小标题不满足宋体小四（12磅）不加粗。"""
    return _small_title_bad(ctx, ctx['lv4num'], 12.0, False, 8)


def _d05(ctx):
    """冒号型小标题与提示行不满足宋体小四（12磅）不加粗。"""
    return _small_title_bad(ctx, ctx['lv4colon'], 12.0, False, 8)


def _d06(ctx):
    """"成员分工：""指导教师："不满足五号（10.5磅）加粗。"""
    return _small_title_bad(ctx, ctx['boldhint'], 10.5, True, 2, font=False)


def _d07(ctx):
    """西文字体不满足 Times New Roman 与宋体并存（被统一替换）。"""
    d, n = ctx['doc'], 0
    for p in d.paragraphs:
        if id(p._element) in ctx['hset']:
            continue
        for r in p.runs:
            if re.search(r'[A-Za-z0-9]', r.text) and (_eff_ascii(d, r) or '') in STRICT_SONGTI:
                n += 1
    for t in ctx['tables']:
        for r in t['cruns']:
            if re.search(r'[A-Za-z0-9]', r.text) and (_eff_ascii(d, r) or '') in STRICT_SONGTI:
                n += 1
    return n < BASE_SONGTI_ASCII * 0.5


def _d08(ctx):
    """正文叙述段落不满足宋体小四（12磅）、首行缩进 2 字符、两端对齐。"""
    d, ps = ctx['doc'], ctx['narrative']
    if not ps:
        return True
    bad = 0
    for p in ps:
        okf = True
        for r in _runs(p):
            sz = _eff_size(d, r)
            if (_eff_eastasia(d, r) or '') not in STRICT_SONGTI or sz is None \
                    or abs(sz - 12.0) > 0.26:
                okf = False
                break
        okj = _eff_para_attr(d, p, 'w:jc', 'val') == 'both'
        oki = (_eff_para_attr(d, p, 'w:ind', 'firstLineChars') == '200'
               or (_eff_para_attr(d, p, 'w:ind', 'firstLine') or '') in ('420', '480'))
        if not (okf and okj and oki):
            bad += 1
    return bad > len(ps) * 0.2


def _d09(ctx):
    """表格文字出现小于五号（9磅）的字号。"""
    return any(t['minsize'] is not None and t['minsize'] < 8.9 for t in ctx['tables'])


def _d10(ctx):
    """资源表、月度台账表、创新维度表不满足无边框排版。"""
    group = _res_tables(ctx) + _find(ctx, H_LEDGER) + _find(ctx, H_INNOV)
    return any(t['bordered'] for t in group)


def _d11(ctx):
    """交叉分析表、院落类型表、季节表、团队表不满足无边框排版。"""
    return any(t['bordered'] for t in _age_court_team(ctx))


def _d12(ctx):
    """调研表、启动资金表、三张收入表、现金流表不满足带网格边框排版。"""
    group = _grid_group(ctx)
    return (not group) or any(not t['bordered'] for t in group)

def _d13(ctx):
    """资源表不满足仅含一行表头、总计 36 行（被拆表并新增重复表头行）。"""
    rs = _res_tables(ctx)
    if not rs:
        return True
    heads = sum(1 for t in rs if _hp(t['header'], H_RES))
    return heads != 1 or sum(t['nrows'] for t in rs) != 36


def _d14(ctx):
    """表格列宽与列数不满足与表头项数一致（空列、列被合并或整体缩窄）。"""
    sec = ctx['doc'].sections[0]
    usable = int(sec.page_width.twips) - int(sec.left_margin.twips) - int(sec.right_margin.twips)
    for t in ctx['tables']:
        if not t['header'] or _hp(t['header'], H_LEDGER) or _hp(t['header'], H_INNOV):
            continue  # 台账/创新维度两表的列结构由加分项 P18 单独评判
        vis = [w for w in t['gridcols'] if w >= 170]
        if len(vis) != len(t['header']):
            return True
        if sum(t['gridcols']) < usable * 0.6:
            return True
    return False


def _d15(ctx):
    """表格单元格文字被压行截断，或表头与数据行行距不一致。"""
    d = ctx['doc']
    for t in ctx['tables']:
        head, data = set(), set()
        for i, row in enumerate(t['rowcells']):
            for c in row:
                for p in c.paragraphs:
                    if not p.text.strip():
                        continue
                    mult, rule = _line_mult(d, p)
                    if rule != 'auto':
                        return True
                    (head if i == 0 else data).add(round(mult or 0, 2))
        if head and data and head != data:
            return True
    return False


def _d16(ctx):
    """资源表分类与资源名称被改写。"""
    rows = []
    for t in _res_tables(ctx):
        for row in t['rowcells']:
            rows.append([_norm(c.text) for c in row])

    def pick(key):
        return [r for r in rows if len(r) >= 4 and key in r[2] and '·' not in r[3]]

    a, b, c = pick('EAA土木院落'), pick('EBB老粮仓广场'), pick('FAA驿站遗址')
    if not (a and b and c):
        return True
    if not a[0][0].startswith('F'):
        return True
    if not b[0][0].startswith('F') or b[0][3] != '松火节、水源节':
        return True
    return c[0][3] != '松火节、水源节'


def _d17(ctx):
    """表格单元格内的说明文字未完整保留。"""
    t = _norm(ctx['text'])
    return any(_norm(k) not in t for k in CELL_NOTES)


def _d18(ctx):
    """正文文字、段落顺序与章节结构未完整保留。"""
    if len(ctx['nonempty']) < BASE_NONEMPTY_PARAS * 0.95:
        return True
    if len(ctx['heads'][1]) != 8 or len(ctx['heads'][2]) != 17:
        return True
    return len(ctx['heads'][3]) < 38


def _d19(ctx):
    """封面图片与数据图表插图未全部保留、位置与大小改变。"""
    if ctx['n_drawing'] != BASE_DRAWING or ctx['n_pict'] != BASE_PICT:
        return True
    return sorted(ctx['extents']) != sorted(BASE_EXTENTS)


def _d20(ctx):
    """表格数量与行数未完整保留（存在被删除的表格或被删减的数据行）。"""
    return len(ctx['tables']) < BASE_TABLE_COUNT or ctx['total_rows'] < BASE_TABLE_ROWS


def _d21(ctx):
    """数值、姓名与班级信息未保持不变。"""
    t = ctx['text']
    return any(k not in t for k in KEEP_NUMBERS + KEEP_NAMES)


def _d22(ctx):
    """页眉、页脚与页码未保持无自定义内容的既有设置。"""
    return bool(ctx['hf_text']) or ctx['hf_page_fields'] != BASE_PAGE_FIELDS


def _d23(ctx):
    """未保持全文 3 个节的分节设置。"""
    return ctx['n_sections'] != BASE_SECTIONS

# ============================ 规则表 ============================

_PLUS_RULES = [
    (5, '全部正文叙述段落（约200个段落）为单倍行距（1.0倍）', _p01),
    (1, '"成员分工：""指导教师：""三大核心因素清晰可见：""现金流保障措施："等提示行与图表下方说明段落同样落实单倍行距，全文无遗漏的1.5倍行距正文段落', _p02),
    (5, '全部表格的单元格文字为1.25倍行距（含调研表、交叉分析表、成本费用表、风险表、团队表等全部表格）', _p03),
    (5, '"市场分析""项目简介""产品与服务""商业模式""财务分析""风险识别与管理对策""未来前景分析""团队介绍"共8处一级标题为宋体、小三（15磅）', _p04),
    (3, '上述8处一级标题为单倍行距', _p05),
    (3, '上述8处一级标题段前0.5行、段后0.5行（按"行"设置，非固定磅值）', _p06),
    (5, '"环境分析""旅游资源分析""数据分析""启动背景"等共17处二级标题为宋体、四号（14磅）', _p07),
    (5, '"政治环境（Politics）""样本构成分析""政策层面""线下获客（补充渠道）""运营可行性"等约38处三级标题为宋体、小四（12磅）', _p08),
    (3, '上述三级标题为加粗', _p09),
    (3, '上述三级标题为单倍行距', _p10),
    (3, '正文"本项目的运盈重点..."与表格"小体量运盈"中的"运盈"均改为"运营"', _p11),
    (1, '正文"避免测祘口径不一致"中的"测祘"改为"测算"', _p12),
    (3, '调研统计表与财务表中以顿号充当小数点的百分比（"24、12%""33、67%"等约60处）均改为小数点写法', _p13),
    (5, '"（二）旅游资源分析"下方"主类|亚类|基本类型|旅游资源单体名称"资源表为一个连续表格（1行表头+35行数据，共36行），跨两页排布', _p14),
    (1, '该资源表在第二页续排处重复显示"主类|亚类|基本类型|旅游资源单体名称"表头行', _p15),
    (5, '"创新点分析"下方表格内容完整排布在三页上——月度台账37行（一月·第1轮至十二月·第3轮）与创新维度7行行数与文字均无删减', _p16),
    (3, '"创新点分析"下方表格的列结构与表头一一对应——月度台账部分为5列、创新维度部分为3列，不再出现空列、错位与并入同一表格框的情形', _p17),
    (3, '除"（二）旅游资源分析"与"创新点分析"下方两处表格外，其余表格均完整落在同一页内，未被拆到两页显示', _p18),
    (1, '为实现指定的分页效果，采用了关闭"允许跨页断行"、设置重复标题行、调整表格前后分页位置等排版手段，未删减表格行列或表格内容', _p19),
]
_MINUS_RULES = [
    (-3, '正文页左右页边距不满足 3.18 厘米（被改为 2.5 厘米一类数值）', _d01),
    (-1, '纸张与上下页边距不满足 A4（宽210毫米、高297毫米）、上下 2.54 厘米', _d02),
    (-3, '"环境分析""旅游资源分析""数据分析""启动背景"等17处二级标题不满足1.5倍行距', _d03),
    (-1, '"1.1 地域分布""1.2 年龄结构""2.1 旅居经历"等编号小标题不满足宋体小四（12磅）不加粗（被放大为四号14磅一类字号）', _d04),
    (-1, '"2.1宏观市场规模：""3.2四大体验板块：""7.2现金流预测"等冒号型小标题与"三大核心因素清晰可见：""现金流保障措施："等提示行不满足宋体小四（12磅）不加粗', _d05),
    (-1, '"成员分工：""指导教师："不满足五号（10.5磅）加粗', _d06),
    (-3, '正文与表格中的数字、英文（如"48""199""19公里""32%""Politics""Wi-Fi""OTA""IP"）不满足西文字体保持 Times New Roman 与宋体两种并存（约190处随中文一并使用宋体的数字与英文被统一替换为 Times New Roman 一类西文字体）', _d07),
    (-3, '正文叙述段落不满足宋体小四（12磅）、首行缩进2字符、两端对齐', _d08),
    (-5, '表格文字出现小于五号（9磅）的字号——如 4.5磅、6.5磅、7磅、7.5磅、8.5磅（为使表格落在一页内将字号适度下调至五号9磅及以上不扣分）', _d09),
    (-3, '"主类|亚类|基本类型|旅游资源单体名称"资源表、"月份|重点工作|责任主体|交付物|备注"月度台账表、"创新维度|具体体现|不可复制性来源"表不满足无边框排版（被添加网格实线边框）', _d10),
    (-1, '"年龄段|1800元以下|..."等交叉分析表、"院落类型|..."表、"季节|核心体验项目|..."表、"姓名|专业班级|学院|职务|核心职责"团队表不满足无边框排版', _d11),
    (-1, '"选项|小计|比例"系列调研表、"项目|金额（万元）|占比|资金使用说明"启动资金表、"业务线|营收（万元）|占比|核心驱动因素"三张收入表、"年度|期初现金（万元）|..."表不满足带网格边框排版', _d12),
    (-1, '资源表不满足全表仅含一行"主类|亚类|基本类型|旅游资源单体名称"表头、共36行（被拆成两个独立表格并新增一行重复表头，总行数变为37行）', _d13),
    (-1, '表格列宽与列数不满足与表头项数一致（出现空列、列被合并或表格被整体缩窄留白）', _d14),
    (-1, '表格单元格文字被压行截断，或同一表格内表头行与数据行的行距不一致', _d15),
    (-3, '资源表中"EA 传统民居|EAA 土木院落|青岭老院群"与"EB 公共空间|EBB 老粮仓广场|松火节、水源节"两行的主类列不满足以"F"开头，"FA 物质遗存|FAA 驿站遗址"行的旅游资源单体名称不满足为"松火节、水源节"（分类与资源名称被改写为"E 建筑设施""旧驿厅遗址"一类内容）', _d16),
    (-3, '表格单元格内的说明文字不满足完整保留（如"跨页表格，便于季度复盘""院落巡检、活动复盘、农产品清单校对""低于警戒值暂停扩容"等被删改或压缩改写）', _d17),
    (-3, '正文文字、段落顺序与章节结构不满足完整保留（存在删除、改写、合并或重新排序的段落与章节）', _d18),
    (-3, '封面图片与"1.1 地域分布""2.1 旅居经历"等小标题下方的数据图表插图不满足全部保留、位置与大小不变', _d19),
    (-3, '表格数量与行数不满足完整保留（存在被删除的表格或被删减的数据行）', _d20),
    (-1, '"165万元""128万元""3.2年""19公里""28分钟""32%""61%"等数值与"陆修远""何修晴""周一行""赵清澜""林若楠""许知瑾""梁冥香"等姓名、"旅游管理23级1班"等班级信息不满足保持不变', _d21),
    (-1, '页眉、页脚与页码不满足保持无自定义内容的既有设置（新增页眉文字、页码或删除既有设置）', _d22),
    (-1, '不满足全文3个节的分节设置（新增分节符改变了页面方向、纸张或页码编排）', _d23),
]

_MAX_DELTAS_ORDER = [s for s, _d, _f in _PLUS_RULES] + [s for s, _d, _f in _MINUS_RULES]
assert sum(d for d in _MAX_DELTAS_ORDER if d > 0) == MAX_SCORE_TOTAL
assert len(_PLUS_RULES) == 19 and len(_MINUS_RULES) == 23

# ============================ 主流程 ============================

def _run(fn, ctx):
    try:
        return bool(fn(ctx)), ''
    except Exception as exc:
        return False, '检查异常：%s: %s' % (type(exc).__name__, exc)


def _zero_items():
    return [{'rule': desc, 'max_delta': score, 'delta': 0, 'hit': False}
            for score, desc, _fn in _PLUS_RULES + _MINUS_RULES]


def _find_target_file(dir_path):
    """仅在给定目录本身定位交付物（不进入子目录），忽略 ~$/. 开头的临时隐藏文件。

    优先取 .docx；一个都没有时退而取 .doc/.rtf/.odt/.pdf 等其他格式，好让维度一
    能就“格式不对”判不通过，而不是报“没找到交付物”。"""
    try:
        names = os.listdir(dir_path)
    except OSError:
        names = []
    files = sorted(f for f in names
                   if not f.startswith('~$') and not f.startswith('.')
                   and os.path.isfile(os.path.join(dir_path, f)))
    cands = [f for f in files if f.lower().endswith(_ACCEPTED_EXT)]
    if not cands:
        cands = [f for f in files if f.lower().endswith(_WRONG_FMT_EXTS)]
    if not cands:
        raise FileNotFoundError('目录中未找到 .docx 文件：%s' % dir_path)
    return os.path.join(dir_path, cands[0])


def _evaluate_impl(file_path, result):
    # 维度一的第一道门禁：交付物格式只接受 .docx，不合格就地判 0，不再解析内容
    reject = _fmt_reject(file_path)
    if reject:
        result['dim1_pass'] = False
        result['dim1_reason'] = reject
        result['dim2_items'] = _zero_items()
        result['total_score'] = 0
        return result

    ctx = _build_ctx(file_path)
    dim1_pass, dim1_reason = _dim1(ctx)
    result['dim1_pass'] = dim1_pass
    result['dim1_reason'] = dim1_reason

    results, details = [], []
    for score, desc, fn in _PLUS_RULES:
        hit, detail = _run(fn, ctx)
        results.append('+%d：%s' % (score if hit else 0, desc))
        details.append(detail)
    for score, desc, fn in _MINUS_RULES:
        hit, detail = _run(fn, ctx)
        results.append('-%d：%s' % (abs(score) if hit else 0, desc))
        details.append(detail)

    if len(results) != len(_MAX_DELTAS_ORDER):
        raise RuntimeError('dim2 项数不一致：results=%d vs order=%d'
                           % (len(results), len(_MAX_DELTAS_ORDER)))

    dim2_items, total_score = [], 0
    for raw, max_delta, _detail in zip(results, _MAX_DELTAS_ORDER, details):
        head, _sep, rule = raw.partition('：')
        try:
            delta = int(head.replace('+', '').strip())
        except ValueError:
            delta = 0
        hit = delta != 0
        dim2_items.append({'rule': rule, 'max_delta': max_delta,
                           'delta': max_delta if hit else 0, 'hit': hit})
        total_score += (max_delta if hit else 0)

    if not dim1_pass:
        for item in dim2_items:
            item['delta'] = 0
            item['hit'] = False
        total_score = 0

    result['dim2_items'] = dim2_items
    result['total_score'] = max(0, total_score)
    return result

def evaluate(dir_path=None):
    """评估交付目录下的唯一交付物。

    dir_path 为待评估目录；未传入时退回本脚本所在目录（_SCRIPT_DIR），
    既不进入子目录，也不会去别的目录找文件。

    交付物只接受 .docx：目录内优先取 .docx；若只交了 .doc / .rtf / .odt / .pdf 等
    其他格式，仍把它收作被评对象，由维度一判不通过（dim1_pass=False、total_score
    记 0、max_score 仍取满分），而不是当成“没有交付物”跳过评测。
    """
    dir_path = dir_path if dir_path and os.path.isdir(dir_path) else _SCRIPT_DIR
    result = {
        'id': SCRIPT_ID,
        'file_name': '',
        'status': 'ok',
        'error': None,
        'dim1_pass': False,
        'dim1_reason': '',
        'dim2_items': [],
        'total_score': 0,
        'max_score': MAX_SCORE_TOTAL,
    }
    try:
        if not dir_path or not os.path.isdir(dir_path):
            raise FileNotFoundError('目录不存在或不是目录：%s' % dir_path)
        file_path = _find_target_file(dir_path)
        result['file_name'] = os.path.basename(file_path)
        return _evaluate_impl(file_path, result)
    except FileNotFoundError as exc:
        # 找不到被评估文件：error 只回传纯文本原因（不带异常类名），维度一根本没进入
        # 判定，dim1_pass 与 dim1_reason 一律留空（空字符串，而非 False），
        # dim2 不出项，满分一并归零。
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
        if type(exc).__name__ in ('XMLSyntaxError', 'ParseError', 'BadZipFile') or (
            isinstance(exc, KeyError) and 'There is no item named' in str(exc)
        ):
            # 交付文件自身结构/XML 损坏：属于维度一不通过（正常 0 分），不是脚本错误
            result['status'] = 'ok'
            result['error'] = None
            result['dim1_pass'] = False
            result['dim1_reason'] = '交付文件结构不完整，无法正常打开：%s: %s' % (
                type(exc).__name__, exc,
            )
            result['dim2_items'] = _zero_items()
            result['total_score'] = 0
            return result
        result['status'] = 'error'
        result['error'] = '脚本运行异常：%s: %s' % (type(exc).__name__, exc)
        result['dim1_pass'] = False
        result['dim1_reason'] = result['error']
        result['dim2_items'] = _zero_items()
        result['total_score'] = 0
        return result


if __name__ == '__main__':
    # 不接受命令行目录参数：评估目标固定为本脚本所在目录
    print(json.dumps(evaluate(), ensure_ascii=False, indent=2))
