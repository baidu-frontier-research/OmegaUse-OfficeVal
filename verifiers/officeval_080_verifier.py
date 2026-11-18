#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# pyright: reportUnknownMemberType=false, reportUnknownVariableType=false, reportUnknownArgumentType=false, reportUnknownParameterType=false, reportMissingParameterType=false, reportPrivateUsage=false, reportUnknownLambdaType=false, reportGeneralTypeIssues=false

import os
import re
import sys
import json
import math
from pptx import Presentation

# 说明：按接口统一约定，本模块不修改全局 sys.stdout，也不 print 主结果。
# 主结果统一由 evaluate(dir_path) -> dict 返回；入参为"脚本所在目录的路径"，
# 由脚本自己在该目录内定位并打开被评估的 .pptx。本文件仅在 __main__ 下作为
# 本地调试入口，才把 dict 序列化为 JSON 打印。
#
# 本脚本不做渲染，全部判定都落在 OOXML/PresentationML 里可读可改的属性上：
# 形状几何（EMU→厘米）、run 字号/西文字体/字体颜色、页面背景 srgbClr、图片
# 原始像素与 a:srcRect 裁剪、p:timing 动画树（p:animEffect/@filter、
# p:cTn/@presetClass|@nodeType、p:spTgt/@spid）、p:transition 切换效果，
# 与在 PowerPoint/WPS 里打开后看到的效果口径一致。

SCRIPT_ID = "080"

# 维度二全部评分项；元组顺序 == _evaluate_impl 中 _score(...) 的调用顺序。
# 正分为加分项（命中记满分），负分为减分项（触发记负分）。文案与细则逐条一致。
_RULES = (
    (3, '第一页"UNIT 3""Comics & Animation""Story Writing Lab""Build a story from panels, bubbles, and action."及"observe visual clues／organize events／write a vivid adventure"三条要点的字号统一增大2-3磅，实测分别为15-16磅、28-29磅、44-45磅、18-19磅、19-20磅'),
    (1, '第一页字号增大后各段文字仍完整显示在各自文本框内，未与右侧插图重叠、未溢出页面边界'),
    (5, '第二页背景填充改为与第三页完全一致的米黄色，色值精确为FFF7E8'),
    (3, '第二页改为浅色底后，"Writing""From comic strip to complete story""Turn pictures into paragraphs.""• notice character, place, and action"等文字颜色同步改为深色，与米黄底对比清晰可读，未出现白色/浅青色字压米黄底而不可见'),
    (1, '第二页右侧插图、标题下方装饰分隔线与右下角"02"页码均完整保留'),
    (3, '第三页页面右侧新增1张插图（教室黑板"Today\'s Plan"场景），图片宽度为11.8-12.2厘米，且未被横向或纵向拉伸变形'),
    (1, '第三页新增图片整体位于页面右半区（左边界≥18厘米）并完整落在页面内'),
    (1, '第三页"Learning objectives"标题为25磅、副标题为10.5磅，"01-05"五条学习目标文字为18磅且各自单行显示未折行'),
    (5, '第四页完成左右版式对调——插图移至页面左侧（左边界≤2.5厘米），"What kind of story is this?"、三条提问要点及"A strong prediction helps you write a stronger plot."整体移至页面右侧（左边界≥17厘米）'),
    (1, '第四页"Warm-up: meet the explorer"标题、"Observe the character and predict the adventure."副标题及其下方装饰分隔线随正文一并移至页面右侧，与右栏文字左对齐'),
    (1, '第四页图文对调后图片与文字区互不重叠，文字右边界未超出33.9厘米'),
    (3, '第五页新增2张插图（漫画分镜页面素材）'),
    (3, '第五页两张新增图片放置于页面空白区，未压盖"1""2""3"三个序号、"The Cloud Compass""Island Under the Waves""Clockwork Cave"三个标题、三段说明文字及三条分隔线'),
    (1, '第五页两张新增图片完整落在页面内，未超出上、下、左、右页面边界'),
    (5, '第六页"Clear sky""Storm wind""Huge wave""New island"四行中箭头后的文字（start the trip／face danger／make a choice／find a way out）全部删除，页面上不再出现这四段文字'),
    (3, '第六页四行箭头后各补充上下两条横线（共8条），横线为可单独选中的直线形状或下划线，同行两条横线上下分离不重叠'),
    (1, '第六页每条新增横线横向长度≥5厘米（约6.5厘米），留有可书写填空的空间'),
    (1, '第六页左侧插图、"Clear sky""Storm wind""Huge wave""New island"四个标签文字及四个箭头符号"→"均完整保留'),
    (5, '第八页4组（A／Speech bubble、B／Caption、C／Sound effect、D／Thought bubble）的小标题与其对应释义文字均添加了进入动画'),
    (5, '第九页4组（Caption、Speech bubble、Thought bubble、Sound effect）的小标题与其对应例句（The scene begins at sunset.／"Watch out!"／I hope the boat is strong.／Crash!）均添加了进入动画'),
    (5, '第八页所添加的各条动画效果两两互不相同，且每组小标题与其对应释义使用的效果也不相同（8个对象对应8种不同效果）'),
    (5, '第九页所添加的各条动画效果两两互不相同，且每组小标题与其对应例句使用的效果也不相同（8个对象对应8种不同效果）'),
    (1, '第八、九页动画均为单击触发，并按"小标题→对应内容"的先后顺序依次播放'),
    (5, '标题为"Build a clear structure"、含"Character → Change → Sequence"及三条编号提问的整页内容已删除，并在第10页位置新增一页，交付稿仍为18页'),
    (3, '新增第10页的标题文字为"Match five stages of the story with frames"，位于页面顶部标题位置（顶边纵向≤1.5厘米）'),
    (5, '新增第10页在标题下方插入故事发展曲线图（含Beginning/Opening、Development/Rising action、Climax/Turning point、Falling action、Ending五个阶段），图片宽度为19.8-20.2厘米且未被拉伸变形'),
    (1, '新增第10页图片顶边位于标题下方（纵向≥3.5厘米），并在页面内水平居中（左右留白差≤1厘米）'),
    (1, '新增第10页标题字号与其余内容页标题一致，为25磅'),
    (1, '新增第10页右下角保留与页序一致的"10"页码标记'),
    (1, '新增第10页背景与其余内容页一致，为米黄色FFF7E8'),
    (5, '第十一页6个对话文本框（"A storm is coming!""We are in trouble.""Look! There is an island.""Quick! Row faster!""Thank goodness!""We made it!"）全部添加了进入动画'),
    (3, '第十一页6个对话文本框使用的进入动画为同一种效果，且为单击逐个触发，便于逐条讲解'),
    (1, '第十一页主标题"While writing"未被添加动画'),
    (3, '第十六页右下角插图（位于横向21.3厘米、纵向13.5厘米，尺寸约8.9×3.4厘米）已删除，该位置无残留图片或空白图片框'),
    (1, '第十六页"Tense／Person／Voice"三行共9个文本框的文字内容、字号与位置完整保留，未随图片删除而丢失'),
    (5, '全部18页均设置了页面切换动画'),
    (5, '18页的切换动画效果两两互不相同，共出现18种不同的切换效果'),
    (1, '交付稿仍为18页，页序与各页主题（封面→Writing→Learning objectives→…→Final checklist）保持一致'),
    (-5, '第二、三、四、五、六、八、九、十、十一、十二、十三、十四、十五页的标题与正文文字不满足西文字体随主题字体（Aptos Display／Aptos）设置，被统一写死为"Arial"'),
    (-3, '第三页"Identify frames, bubbles, captions, and sound words""Organize the plot: beginning, problem, action, ending"等5条学习目标文本框不满足宽度25.15厘米、文字单行显示（被压窄至13.8厘米，导致文字折行并超出0.81厘米的框高）'),
    (-3, '第五页底部"Use title words to guess the setting, conflict, and mood."不满足字号22磅、顶边位于纵向14.35厘米（被改为11.5磅并上移至纵向11.78厘米）'),
    (-1, '第一页右下角页码"01"不满足字号8磅'),
    (-1, '第九页位于纵向7.44厘米与11.76厘米处的4条装饰分隔线不满足不参与动画（被一并加入动画播放序列）'),
    (-1, '第十页不满足无空白元素残留（纵向2.39厘米处遗留1个无任何文字的空文本框）'),
)

# 维度二各评分项按"结果追加顺序"的 max_delta 列表；末端把 results 里的
# "+N：..." / "-N：..." 字符串反解为结构化 dim2_items 时用。
_MAX_DELTAS_ORDER = [p for p, _ in _RULES]
# 维度二满分 = 全部加分项之和（减分项只在触发时扣，不计入满分）
MAX_SCORE_TOTAL = sum(p for p in _MAX_DELTAS_ORDER if p > 0)
MIN_PENALTY_TOTAL = sum(p for p in _MAX_DELTAS_ORDER if p < 0)
assert len(_RULES) == 44
assert MAX_SCORE_TOTAL == 104 and MIN_PENALTY_TOTAL == -14

# ---------------------------------------------------------------- 基础常量/工具

_A_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main'
_P_NS = 'http://schemas.openxmlformats.org/presentationml/2006/main'
_MC_NS = 'http://schemas.openxmlformats.org/markup-compatibility/2006'

EMU_PER_CM = 360000.0
PT_PER_CM = 28.3465
PAGE_W_CM = 33.87
PAGE_H_CM = 19.05

# PowerPoint 里"动画窗格"能列出的行为节点；p:cTn 的 childTnLst 直接挂这些标签
# 时才算一条真实动画效果（cBhvr 内部的 p:cTn 不带 childTnLst，自动被排除）。
_BEHAVIOR_TAGS = {'animEffect', 'set', 'anim', 'animScale', 'animRot',
                  'animMotion', 'animClr'}
# a:latin/@typeface 取这些值表示"跟随主题字体"，不算写死字体
_THEME_FONT_REFS = {'+mj-lt', '+mn-lt', '+mj-ea', '+mn-ea', '+mj-cs', '+mn-cs'}


def _a(tag):
    return '{%s}%s' % (_A_NS, tag)


def _p(tag):
    return '{%s}%s' % (_P_NS, tag)


def _local(elem):
    """去掉命名空间的标签名。"""
    return elem.tag.split('}')[-1] if isinstance(elem.tag, str) else ''


def _cm(emu):
    """EMU → 厘米；360000 EMU = 1cm，与 PowerPoint 标尺一致。"""
    return None if emu is None else float(emu) / EMU_PER_CM


def _norm(s):
    """折叠空白后的可见文本，用于按文案定位形状。"""
    return re.sub(r'\s+', ' ', (s or '').replace(' ', ' ')).strip()


def _squash(s):
    """去掉全部空白，用于比对含空格差异的短语。"""
    return re.sub(r'\s+', '', s or '')

def _shape_text(shape):
    """形状里的可见文字；无文本框/纯装饰形状返回空串。"""
    try:
        if not shape.has_text_frame:
            return ''
        return _norm(shape.text_frame.text)
    except Exception:
        return ''


def _sid(shape):
    """形状 id；与 p:timing 里 p:spTgt/@spid 同一套编号。"""
    try:
        return int(shape.shape_id)
    except Exception:
        return -1


def _box(shape):
    """(left, top, right, bottom) 单位厘米；几何缺失返回 None。"""
    try:
        if None in (shape.left, shape.top, shape.width, shape.height):
            return None
        l = _cm(shape.left)
        t = _cm(shape.top)
        return (l, t, l + _cm(shape.width), t + _cm(shape.height))
    except Exception:
        return None


def _area(b):
    return 0.0 if b is None else max(0.0, b[2] - b[0]) * max(0.0, b[3] - b[1])


def _overlap(b1, b2):
    """两个矩形的相交面积（cm²）。"""
    if b1 is None or b2 is None:
        return 0.0
    ow = min(b1[2], b2[2]) - max(b1[0], b2[0])
    oh = min(b1[3], b2[3]) - max(b1[1], b2[1])
    return ow * oh if ow > 0 and oh > 0 else 0.0

def _defrpr_sz(elem):
    """从 a:pPr/a:defRPr 或 a:lstStyle//a:defRPr 取字号（半点→磅）。"""
    if elem is None:
        return None
    for d in elem.iter(_a('defRPr')):
        v = d.get('sz')
        if v and v.isdigit():
            return int(v) / 100.0
    return None


def _run_sizes(para, shape):
    """段落内各非空 run 的实际字号（磅）：run 级 sz → 段落 defRPr → 形状 lstStyle。"""
    out = []
    try:
        runs = list(para.runs)
    except Exception:
        return out
    for run in runs:
        if not (run.text or '').strip():
            continue
        v = None
        try:
            if run.font.size is not None:
                v = run.font.size.pt
        except Exception:
            v = None
        if v is None:
            v = _defrpr_sz(para._p.find(_a('pPr')))
        if v is None:
            try:
                v = _defrpr_sz(shape.text_frame._txBody.find(_a('lstStyle')))
            except Exception:
                v = None
        out.append(v)
    return out


def _latin_typefaces(shape):
    """形状内全部 a:latin/@typeface（含 defRPr），对应"西文字体"下拉框的取值。"""
    out = []
    try:
        if not shape.has_text_frame:
            return out
        for el in shape.text_frame._txBody.iter(_a('latin')):
            tf = el.get('typeface')
            if tf:
                out.append(tf)
    except Exception:
        pass
    return out

def _run_color_hex(run):
    """run 的显式字体颜色（a:rPr/a:solidFill）；schemeClr 归并为黑/白，未显式设置返回 None。"""
    try:
        rPr = run._r.find(_a('rPr'))
    except Exception:
        return None
    if rPr is None:
        return None
    fill = rPr.find(_a('solidFill'))
    if fill is None:
        return None
    srgb = fill.find(_a('srgbClr'))
    if srgb is not None and srgb.get('val'):
        return srgb.get('val').upper()
    sch = fill.find(_a('schemeClr'))
    if sch is not None:
        v = (sch.get('val') or '').lower()
        if v in ('tx1', 'dk1', 'tx2', 'dk2'):
            return '000000'
        if v in ('bg1', 'lt1', 'bg2', 'lt2'):
            return 'FFFFFF'
    return None


def _brightness(hex_val):
    """感知亮度（0-255）；用于判定"深色字"是否压得住米黄底。"""
    try:
        r = int(hex_val[0:2], 16)
        g = int(hex_val[2:4], 16)
        b = int(hex_val[4:6], 16)
    except Exception:
        return None
    return 0.299 * r + 0.587 * g + 0.114 * b


def _bg_hex(slide):
    """页面背景色：先看幻灯片自身 p:cSld/p:bg，再回退版式/母版。"""
    if slide is None:
        return None
    try:
        holders = [slide._element, slide.slide_layout._element,
                   slide.slide_layout.slide_master._element]
    except Exception:
        holders = [slide._element]
    for el in holders:
        cSld = el.find(_p('cSld'))
        if cSld is None:
            continue
        bg = cSld.find(_p('bg'))
        if bg is None:
            continue
        srgb = bg.find('.//' + _a('srgbClr'))
        if srgb is not None and srgb.get('val'):
            return srgb.get('val').upper()
        sch = bg.find('.//' + _a('schemeClr'))
        if sch is not None:
            return 'SCHEME:' + (sch.get('val') or '')
    return None

def _pictures(slide):
    """页面上的图片形状（p:pic）。"""
    out = []
    if slide is None:
        return out
    for sh in slide.shapes:
        if _local(sh._element) == 'pic':
            out.append(sh)
    return out


def _pic_aspect_delta(pic):
    """图片显示宽高比与"原始像素经 a:srcRect 裁剪后"宽高比的相对偏差。
    与办公软件对齐：只要不勾选"锁定纵横比"就可能被拉伸，此处按裁剪后有效像素
    计算真实比例，偏差>阈值即判为横向/纵向拉伸变形。无法取到像素时返回 None。"""
    try:
        px_w, px_h = pic.image.size
    except Exception:
        return None
    try:
        kw = 1.0 - (pic.crop_left or 0.0) - (pic.crop_right or 0.0)
        kh = 1.0 - (pic.crop_top or 0.0) - (pic.crop_bottom or 0.0)
    except Exception:
        kw = kh = 1.0
    kw = max(1e-6, kw)
    kh = max(1e-6, kh)
    if not px_w or not px_h or not pic.width or not pic.height:
        return None
    native = (px_w * kw) / (px_h * kh)
    shown = float(pic.width) / float(pic.height)
    return abs(shown - native) / native


def _is_line_shape(shape):
    """可单独选中的直线/装饰分隔线：连接符，或一维方向极扁的形状。"""
    if _local(shape._element) == 'cxnSp':
        return True
    b = _box(shape)
    if b is None:
        return False
    w, h = b[2] - b[0], b[3] - b[1]
    return (h <= 0.25 and w >= 1.0) or (w <= 0.25 and h >= 1.0)


def _line_length_cm(shape):
    b = _box(shape)
    if b is None:
        return 0.0
    return max(b[2] - b[0], b[3] - b[1])

def _est_width_cm(text, size_pt, factor=0.5):
    """按西文平均字宽 0.5em 估算一行文字的排版宽度（厘米）。
    与办公软件对齐：不渲染时无法拿到精确字宽，用 0.5em 这一保守均值判断"该行
    是否放得下"，只用来识别明显折行/明显溢出，不做像素级结论。"""
    if not text or not size_pt:
        return 0.0
    return len(text) * factor * float(size_pt) / PT_PER_CM


def _single_line(shape):
    """形状内每个段落估算宽度是否都不超过框宽（即文字单行显示未折行）。"""
    b = _box(shape)
    if b is None:
        return True
    box_w = max(0.1, b[2] - b[0])
    try:
        paras = list(shape.text_frame.paragraphs)
    except Exception:
        return True
    for para in paras:
        t = _norm(para.text)
        if not t:
            continue
        szs = [v for v in _run_sizes(para, shape) if v]
        sz = max(szs) if szs else 18.0
        if _est_width_cm(t, sz) > box_w:
            return False
    return True


def _text_fits_box(shape, slack=2.0):
    """粗估文本总高是否落在文本框内（含自动换行后的行数）。slack 给足余量，
    只拦明显放不下的情形。"""
    b = _box(shape)
    if b is None:
        return True
    box_w = max(0.1, b[2] - b[0])
    box_h = max(0.01, b[3] - b[1])
    try:
        paras = list(shape.text_frame.paragraphs)
    except Exception:
        return True
    total = 0.0
    for para in paras:
        t = _norm(para.text)
        szs = [v for v in _run_sizes(para, shape) if v]
        sz = max(szs) if szs else 18.0
        if not t:
            total += sz * 0.6 / PT_PER_CM
            continue
        lines = max(1, int(math.ceil(_est_width_cm(t, sz) / box_w)))
        total += lines * sz * 1.25 / PT_PER_CM
    return total <= box_h * slack

def _anim_nodes(slide):
    """解析 p:timing，按播放顺序返回动画效果节点。
    与办公软件对齐：动画窗格里的每一行对应一个带 childTnLst 行为的 p:cTn；
    mainSeq 下的每个直接子节点是"一次单击放出的一组动画"，故用该组序号
    （group）判断是否单击逐个触发以及先后顺序，比读 @nodeType 稳（部分生成
    器只在组级写 clickEffect）。效果种类优先取 @presetID（动画窗格里的效果名
    就是它），无 presetID 时退化用 p:animEffect/@filter；进入动画由
    @presetClass="entr"、animEffect/@transition="in" 或 p:set 把 visibility
    置 visible 来识别。"""
    out = []
    if slide is None:
        return out
    timing = slide._element.find(_p('timing'))
    if timing is None:
        return out
    gmap = {}
    groups = []
    mains = [c for c in timing.iter(_p('cTn'))
             if (c.get('nodeType') or '') == 'mainSeq']
    if mains:
        top = mains[0].find(_p('childTnLst'))
        if top is not None:
            # 注意：groups 必须持有这些元素的引用，否则 lxml 代理对象会被回收、
            # id() 复用，导致组号错乱。
            groups = [list(g.iter(_p('cTn'))) for g in list(top)]
            for gi, ctns in enumerate(groups):
                for ctn in ctns:
                    gmap[id(ctn)] = gi
    for ctn in timing.iter(_p('cTn')):
        ctl = ctn.find(_p('childTnLst'))
        if ctl is None:
            continue
        behaviors = [c for c in ctl if _local(c) in _BEHAVIOR_TAGS]
        if not behaviors:
            continue
        spids = set()
        for tgt in ctn.iter(_p('spTgt')):
            sid = tgt.get('spid')
            if sid:
                try:
                    spids.add(int(sid))
                except ValueError:
                    pass
        if not spids:
            continue
        filt = None
        is_entr = (ctn.get('presetClass') or '') == 'entr'
        for ae in ctn.iter(_p('animEffect')):
            if filt is None and ae.get('filter'):
                filt = ae.get('filter').split('(')[0].strip().lower()
            if (ae.get('transition') or '') == 'in':
                is_entr = True
        if not is_entr:
            for st in ctn.iter(_p('set')):
                for sv in st.iter(_a('strVal')):
                    if (sv.get('val') or '') == 'visible':
                        is_entr = True
        preset_id = ctn.get('presetID')
        if preset_id:
            key = 'preset:' + preset_id
        elif filt:
            key = 'filter:' + filt
        else:
            key = 'behavior:' + '|'.join(sorted({_local(b) for b in behaviors}))
        grp = gmap.get(id(ctn))
        out.append({'spids': spids, 'effect': key, 'is_entr': is_entr,
                    'node_type': ctn.get('nodeType') or '',
                    'group': grp if grp is not None else 1000 + len(out)})
    return out

def _entr_index(slide):
    """(进入动画节点列表, spid→该形状首次出现的播放序号)。"""
    nodes = [n for n in _anim_nodes(slide) if n['is_entr']]
    first = {}
    for i, nd in enumerate(nodes):
        for spid in nd['spids']:
            first.setdefault(spid, i)
    return nodes, first


def _all_anim_spids(slide):
    """页面上被任何动画（不限进入动画）引用的形状 id 集合。"""
    out = set()
    for nd in _anim_nodes(slide):
        out |= nd['spids']
    return out


def _transition_key(slide):
    """页面切换效果标识；无切换返回 None。
    与办公软件对齐：新式切换（平滑/涟漪/翻页等）写在 mc:AlternateContent 的
    mc:Choice 里（p14/p159 命名空间），mc:Fallback 里放一个老版本兜底效果；
    切换选项卡显示的是 Choice 的效果，故优先取非 Fallback 的那个。"""
    if slide is None:
        return None
    found = []
    for tr in slide._element.iter():
        if _local(tr) != 'transition':
            continue
        in_fb = False
        anc = tr.getparent()
        while anc is not None:
            if anc.tag == '{%s}Fallback' % _MC_NS:
                in_fb = True
                break
            anc = anc.getparent()
        eff = None
        for ch in tr:
            ln = _local(ch)
            if ln in ('sndAc', 'extLst'):
                continue
            eff = ln
            if ch.get('prst'):
                eff += '(' + ch.get('prst') + ')'
            break
        if eff:
            found.append((in_fb, eff))
    if not found:
        return None
    for in_fb, eff in found:
        if not in_fb:
            return eff
    return found[0][1]

class _TargetMissing(Exception):
    """目录不存在或目录中没有 .pptx：连待评测对象都没有，不进入维度判定。"""

def evaluate(dir_path):
    """按接口约定返回结构化 dict；不 print、不改 sys.stdout、不 sys.exit。
    入参 dir_path 为"脚本所在目录的路径"，脚本自己在该目录里定位并打开被评估的
    .pptx（忽略以 ~$ 开头的 PowerPoint 临时锁文件）。脚本自身崩溃（含目录不
    存在、未发现文档等）由本函数捕获并返回 status='error'。"""
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
            raise _TargetMissing("目录不存在：%s" % dir_path)
        cands = sorted(
            f for f in os.listdir(dir_path)
            if f.lower().endswith('.pptx') and not f.startswith('~$')
        )
        if not cands:
            raise _TargetMissing("目录中未找到 .pptx 文件：%s" % dir_path)
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
        result["error"] = "%s: %s" % (type(exc).__name__, exc)
        result["dim1_pass"] = False
        result["dim1_reason"] = "文件无法解析或无法打开：%s: %s" % (type(exc).__name__, exc)
        result["dim2_items"] = []
        result["total_score"] = 0
        return result

def _evaluate_impl(file_path, result):
    prs = Presentation(file_path)
    slides = list(prs.slides)
    page_w = _cm(prs.slide_width) or PAGE_W_CM
    page_h = _cm(prs.slide_height) or PAGE_H_CM

    results = []
    details = []
    _cursor = [0]

    def _score(ok, detail=""):
        """按 _RULES 顺序追加一条结果。加分项 ok=命中；减分项 ok=触发扣分。"""
        points, rule = _RULES[_cursor[0]]
        _cursor[0] += 1
        if points > 0:
            results.append(("+%d：%s" % (points, rule)) if ok else ("+0：%s" % rule))
        else:
            results.append(("%d：%s" % (points, rule)) if ok else ("-0：%s" % rule))
        details.append(detail)

    # ---------------------------------------------------------- 页内定位小工具
    def sl(n):
        """第 n 页（1 基）。"""
        return slides[n - 1] if 0 < n <= len(slides) else None

    def texts(slide):
        """[(shape, 可见文本)]，只含有文字的形状。"""
        out = []
        if slide is None:
            return out
        for sh in slide.shapes:
            t = _shape_text(sh)
            if t:
                out.append((sh, t))
        return out

    def paras_with(slide, needle):
        """按文案定位段落：[(shape, paragraph, 段落文本)]，忽略空白与大小写差异。"""
        out = []
        if slide is None:
            return out
        key = _squash(needle).lower()
        for sh in slide.shapes:
            try:
                if not sh.has_text_frame:
                    continue
                paras = list(sh.text_frame.paragraphs)
            except Exception:
                continue
            for para in paras:
                t = _norm(para.text)
                if key and key in _squash(t).lower():
                    out.append((sh, para, t))
        return out

    def shapes_with(slide, needle):
        """按文案定位形状（去重，保持页面 z 序）。"""
        seen, out = set(), []
        for sh, _para, _t in paras_with(slide, needle):
            if _sid(sh) not in seen:
                seen.add(_sid(sh))
                out.append(sh)
        return out

    def page_num_ids(slide):
        """右下角页码标记的形状 id 集合（纯数字且位于右下角）。"""
        out = set()
        for sh, t in texts(slide):
            b = _box(sh)
            if b and b[0] >= 27.5 and b[1] >= 15.5 and re.fullmatch(r'\d{1,2}', t):
                out.add(_sid(sh))
        return out

    def header_ids(slide):
        """页眉区（顶边 <3.5cm）的文本形状 id：页标题与副标题。"""
        out = set()
        for sh, _t in texts(slide):
            b = _box(sh)
            if b and b[1] < 3.5:
                out.add(_sid(sh))
        return out

    def sizes_of(shapes_list):
        """一批形状里全部非空 run 的字号。"""
        out = []
        for sh in shapes_list:
            try:
                paras = list(sh.text_frame.paragraphs)
            except Exception:
                continue
            for para in paras:
                out.extend(_run_sizes(para, sh))
        return out

    def near(vals, lo, hi):
        """所有实测值都落在 [lo, hi]（含 0.05 磅/厘米容差）且至少有一个实测值。"""
        vals = [v for v in vals if v is not None]
        return bool(vals) and all(lo - 0.05 <= v <= hi + 0.05 for v in vals)

    def in_page(b, tol=0.1):
        return (b is not None and b[0] >= -tol and b[1] >= -tol
                and b[2] <= page_w + tol and b[3] <= page_h + tol)

    # ========== 维度1：可用与可修改性 ==========
    # 只判一件事：交付文件是 1 个可正常打开的 .pptx。打不开（文件损坏、格式不
    # 符、目录内无文档）会在 evaluate 里被捕获并落到 status='error'、
    # dim1_pass=False；能走到这里即视为通过。
    result["dim1_reason"] = "交付文件为1个可正常打开的 PowerPoint 演示文稿（.pptx）"

    # ========== 维度2：完成度评分 ==========
    S1, S2, S3, S4, S5, S6 = sl(1), sl(2), sl(3), sl(4), sl(5), sl(6)
    S8, S9, S10, S11, S16 = sl(8), sl(9), sl(10), sl(11), sl(16)

    # +3：第一页五处文字字号统一增大 2-3 磅
    # ①"UNIT 3" 15-16磅；②"Comics & Animation" 28-29磅；③"Story Writing Lab"
    # 44-45磅；④"Build a story from panels, bubbles, and action." 18-19磅；
    # ⑤observe visual clues／organize events／write a vivid adventure 19-20磅。
    # 与办公软件对齐：字号取 a:rPr/@sz（半点），缺省时回退段落 a:defRPr 与形状
    # a:lstStyle，与"开始→字号"框里显示的值一致。
    _SPEC1 = (("UNIT 3", 15.0, 16.0),
              ("Comics & Animation", 28.0, 29.0),
              ("Story Writing Lab", 44.0, 45.0),
              ("Build a story from panels", 18.0, 19.0),
              ("observe visual clues", 19.0, 20.0),
              ("organize events", 19.0, 20.0),
              ("write a vivid adventure", 19.0, 20.0))
    r1_ok = True
    r1_msg = []
    for needle, lo, hi in _SPEC1:
        hits = paras_with(S1, needle)
        if not hits:
            r1_ok = False
            r1_msg.append("%s=未找到" % needle)
            continue
        vals = []
        for shp, para, _t in hits:
            vals.extend(_run_sizes(para, shp))
        r1_msg.append("%s=%s磅" % (needle, "/".join(
            "?" if v is None else ("%g" % v) for v in vals)))
        if not near(vals, lo, hi):
            r1_ok = False
    _score(r1_ok, "；".join(r1_msg))

    # +1：第一页放大后文字仍完整显示在框内、未压右侧插图、未溢出页面
    # ①每个文本框的估算文本高度不超出框高；②文本框与插图的相交面积不超过其自身
    # 面积的 5%；③文本框四边都在页面内。
    r2_msg = []
    for sh, t in texts(S1):
        b = _box(sh)
        if b is None:
            continue
        if not in_page(b):
            r2_msg.append("\"%s\"超出页面边界" % t[:14])
        if not _text_fits_box(sh):
            r2_msg.append("\"%s\"文字溢出文本框" % t[:14])
        for pic in _pictures(S1):
            if _area(b) > 0 and _overlap(b, _box(pic)) / _area(b) > 0.05:
                r2_msg.append("\"%s\"与右侧插图重叠" % t[:14])
    _score(not r2_msg, "；".join(r2_msg) or "各段文字均在框内且未与插图重叠")

    # +5：第二页背景改为与第三页完全一致的米黄色 FFF7E8
    # 与办公软件对齐：读 p:cSld/p:bg//a:srgbClr/@val，即"设置背景格式→纯色填充"
    # 里的十六进制色值；必须精确等于 FFF7E8，且与第三页取值一致。
    bg2, bg3 = _bg_hex(S2), _bg_hex(S3)
    _score(bg2 == 'FFF7E8' and bg3 == 'FFF7E8',
           "第二页背景=%s，第三页背景=%s" % (bg2, bg3))

    # +3：第二页文字同步改为深色，浅底上清晰可读
    # ①"Writing"②"From comic strip to complete story"③"Turn pictures into
    # paragraphs."④"notice character, place, and action" 四处文字的显式字体色
    # 感知亮度 ≤180（即不再是白色/浅青色）。
    # 与办公软件对齐：颜色取 a:rPr/a:solidFill/a:srgbClr；schemeClr 归并为深/浅。
    _DARK_NEEDLES = ("Writing", "From comic strip to complete story",
                     "Turn pictures into paragraphs", "notice character, place, and action")
    r4_seen, r4_msg = 0, []
    for needle in _DARK_NEEDLES:
        for shp, para, t in paras_with(S2, needle):
            for run in para.runs:
                if not (run.text or '').strip():
                    continue
                hx = _run_color_hex(run)
                if hx is None:
                    continue
                r4_seen += 1
                br = _brightness(hx)
                if br is None or br > 180:
                    r4_msg.append("\"%s\"=%s" % (t[:16], hx))
    _score(r4_seen > 0 and not r4_msg,
           ("仍为浅色字：" + "；".join(r4_msg)) if r4_msg
           else "已识别%d处显式深色字体色" % r4_seen)

    # +1：第二页右侧插图、标题下方装饰分隔线、右下角"02"页码均保留
    has_pic2 = len(_pictures(S2)) >= 1
    has_div2 = any(_is_line_shape(s) and not _shape_text(s) for s in S2.shapes)
    has_pn2 = any(t == '02' for _s, t in texts(S2))
    _score(has_pic2 and has_div2 and has_pn2,
           "插图=%s，分隔线=%s，\"02\"页码=%s" % (has_pic2, has_div2, has_pn2))

    # ---- 第三页元素归类：标题 / 副标题 / 01-05 序号 / 5条学习目标 / 页码
    t3 = texts(S3)
    title3 = shapes_with(S3, 'Learning objectives')
    title3_ids = {_sid(s) for s in title3}
    pn3_ids = page_num_ids(S3)
    mark3_ids = {_sid(sh) for sh, t in t3
                 if re.fullmatch(r'0[1-5]', t) and _sid(sh) not in pn3_ids}
    # 页顶（顶边<3.5厘米）的文本框属于标题／副标题区，不计入 01-05 学习目标
    head3_ids = header_ids(S3)
    obj3 = [(sh, t) for sh, t in t3
            if len(t) >= 15
            and _sid(sh) not in title3_ids | pn3_ids | mark3_ids | head3_ids]
    obj3_ids = {_sid(sh) for sh, _t in obj3}
    sub3 = [(sh, t) for sh, t in t3
            if _sid(sh) in head3_ids
            and _sid(sh) not in title3_ids | pn3_ids | mark3_ids | obj3_ids]

    # +3：第三页右侧新增1张插图，宽 11.8-12.2 厘米且未被拉伸变形
    # 与办公软件对齐：宽度取 p:spPr/a:xfrm/a:ext/@cx；变形判定按图片原始像素经
    # a:srcRect 裁剪后的真实宽高比与显示宽高比之差（相对偏差>5% 视为拉伸）。
    pics3 = _pictures(S3)
    new_pic3 = None
    r6_ok = False
    r6_msg = []
    for pp in pics3:
        w = _cm(pp.width)
        d = _pic_aspect_delta(pp)
        r6_msg.append("宽%.2fcm，宽高比偏差%s"
                      % (w or 0.0, "未知" if d is None else "%.1f%%" % (d * 100)))
        if w is not None and 11.8 <= w <= 12.2:
            if new_pic3 is None:
                new_pic3 = pp
            if d is None or d <= 0.05:
                new_pic3 = pp
                r6_ok = True
                break
    if new_pic3 is None and pics3:
        new_pic3 = pics3[0]
    _score(r6_ok, "；".join(r6_msg) or "该页未发现图片")

    # +1：第三页新增图片整体位于右半区（左边界≥18厘米）并完整落在页面内
    b3 = _box(new_pic3) if new_pic3 is not None else None
    _score(b3 is not None and b3[0] >= 18.0 and in_page(b3),
           "无图片" if b3 is None else "左边界%.2fcm，右边界%.2fcm，下边界%.2fcm"
           % (b3[0], b3[2], b3[3]))

    # +1：第三页标题25磅、副标题10.5磅、5条学习目标18磅且各自单行显示未折行
    sz_t3 = sizes_of(title3)
    sz_s3 = sizes_of([s for s, _t in sub3])
    sz_o3 = sizes_of([s for s, _t in obj3])
    fold3 = [t[:18] for sh, t in obj3 if not _single_line(sh)]
    r8_ok = (near(sz_t3, 25.0, 25.0) and near(sz_s3, 10.5, 10.5)
             and len(obj3) == 5 and near(sz_o3, 18.0, 18.0) and not fold3)
    _score(r8_ok, "标题%s磅，副标题%s磅，%d条目标%s磅%s"
           % ("/".join("%g" % v for v in sz_t3 if v is not None) or "?",
              "/".join("%g" % v for v in sz_s3 if v is not None) or "?",
              len(obj3),
              "/".join(sorted({"%g" % v for v in sz_o3 if v is not None})) or "?",
              ("，折行：" + "；".join(fold3)) if fold3 else "，均单行显示"))

    # +5：第四页左右版式对调——插图移至左侧（左边界≤2.5厘米），正文移至右侧
    # （左边界≥17厘米）
    # 与办公软件对齐：位置取 a:off/@x，即"设置形状格式→位置→水平位置"。
    pics4 = _pictures(S4)
    pic4_left = min([_cm(p.left) for p in pics4 if p.left is not None] or [None]) \
        if pics4 else None
    body4 = []
    for needle in ("What kind of story is this?",
                   "A strong prediction helps you write a stronger plot"):
        body4.extend(shapes_with(S4, needle))
    body4_lefts = [_box(s)[0] for s in body4 if _box(s) is not None]
    right_cnt4 = sum(1 for sh, _t in texts(S4)
                     if _box(sh) is not None and _box(sh)[0] >= 17.0)
    r9_ok = (pic4_left is not None and pic4_left <= 2.5
             and bool(body4_lefts) and min(body4_lefts) >= 17.0
             and right_cnt4 >= 3)
    _score(r9_ok, "插图左边界%s，正文左边界%s，页面右半区文本框%d个"
           % ("无图" if pic4_left is None else "%.2fcm" % pic4_left,
              "/".join("%.2fcm" % v for v in body4_lefts) or "未找到", right_cnt4))

    # +1：第四页标题、副标题及其下方装饰分隔线随正文一并右移，并与右栏文字左对齐
    tt4 = shapes_with(S4, 'Warm-up: meet the explorer')
    st4 = shapes_with(S4, 'Observe the character and predict the adventure')
    div4 = [s for s in S4.shapes
            if _is_line_shape(s) and not _shape_text(s)
            and _box(s) is not None and _box(s)[1] < 6.0]
    ref_left4 = min(body4_lefts) if body4_lefts else None
    head4 = [s for s in (tt4 + st4 + div4)]
    head4_lefts = [_box(s)[0] for s in head4 if _box(s) is not None]
    r10_ok = (bool(tt4) and bool(st4) and bool(div4) and ref_left4 is not None
              and bool(head4_lefts) and min(head4_lefts) >= 17.0
              and all(abs(v - ref_left4) <= 0.6 for v in head4_lefts))
    _score(r10_ok, "标题/副标题/分隔线左边界=%s，右栏正文左边界=%s"
           % ("/".join("%.2f" % v for v in head4_lefts) or "未找到",
              "?" if ref_left4 is None else "%.2f" % ref_left4))

    # +1：第四页图文互不重叠，文字右边界未超出33.9厘米
    r11_msg = []
    for sh, t in texts(S4):
        b = _box(sh)
        if b is None:
            continue
        if b[2] > 33.9:
            r11_msg.append("\"%s\"右边界%.2fcm" % (t[:14], b[2]))
        for pic in _pictures(S4):
            if _area(b) > 0 and _overlap(b, _box(pic)) / _area(b) > 0.05:
                r11_msg.append("\"%s\"与插图重叠" % t[:14])
    _score(not r11_msg, "；".join(r11_msg) or "图文分列且文字右边界均≤33.9厘米")

    # +3：第五页新增2张插图（漫画分镜页面素材）
    pics5 = _pictures(S5)
    _score(len(pics5) >= 2, "该页图片数=%d" % len(pics5))

    # +3：第五页两张新增图片放置在空白区，未压盖三个序号、三个标题、三段说明文字
    # 及三条分隔线
    # 与办公软件对齐：不做渲染，按形状外框相交面积判断"压盖"；受保护对象为该页
    # 除页标题/副标题（顶边<3.5cm）、右下角页码、底部提示句以外的文字形状，以及
    # 位于内容区（顶边≥4cm）的分隔线。
    bottom5 = shapes_with(S5, 'Use title words to guess the setting')
    excl5 = page_num_ids(S5) | header_ids(S5) | {_sid(s) for s in bottom5}
    protected5 = [(sh, t) for sh, t in texts(S5) if _sid(sh) not in excl5]
    protected5 += [(s, '分隔线') for s in S5.shapes
                   if _is_line_shape(s) and not _shape_text(s)
                   and _box(s) is not None and _box(s)[1] >= 4.0]
    r13_msg = []
    for pic in pics5:
        pb = _box(pic)
        for sh, t in protected5:
            b = _box(sh)
            if _area(b) > 0 and _overlap(pb, b) / _area(b) > 0.15:
                r13_msg.append("压盖\"%s\"（相交%.0f%%）"
                               % (t[:16], _overlap(pb, b) / _area(b) * 100))
    _score(len(pics5) >= 2 and not r13_msg,
           "；".join(sorted(set(r13_msg))) or "两张图片均未压盖序号/标题/说明/分隔线")

    # +1：第五页两张新增图片完整落在页面内
    out5 = ["图%d越界" % (i + 1) for i, p in enumerate(pics5) if not in_page(_box(p))]
    _score(len(pics5) >= 2 and not out5,
           "；".join(out5) or "；".join(
               "图%d(%.2f,%.2f)-(%.2f,%.2f)" % ((i + 1,) + tuple(_box(p)))
               for i, p in enumerate(pics5) if _box(p) is not None))

    # +5：第六页箭头后的四段文字全部删除
    all6 = " ".join(t for _s, t in texts(S6)).lower()
    gone6 = [k for k in ('start the trip', 'face danger', 'make a choice',
                         'find a way out') if k in all6]
    _score(not gone6, ("仍存在：" + "／".join(gone6)) if gone6 else "四段文字均已删除")

    # +3：第六页四行箭头后各补充上下两条横线（共8条），可单独选中，同行上下分离
    # 与办公软件对齐：既接受"插入→形状→直线"（p:cxnSp / 极扁形状，可单独点选），
    # 也接受在文本框里用连续下划线字符排出的横线；两者都只统计位于箭头右侧
    # （左边界≥15厘米）的对象，避免把标题下方装饰分隔线算进来。
    lines6 = [s for s in S6.shapes
              if _is_line_shape(s) and _box(s) is not None and _box(s)[0] >= 15.0]
    us6 = []
    for sh in S6.shapes:
        b = _box(sh)
        try:
            if not sh.has_text_frame or b is None or b[0] < 15.0:
                continue
            paras = list(sh.text_frame.paragraphs)
        except Exception:
            continue
        for para in paras:
            m = re.search(r'_{3,}', para.text or '')
            if m:
                szs = [v for v in _run_sizes(para, sh) if v]
                us6.append((m.group(0), max(szs) if szs else 16.0, b))
    n6 = len(lines6) + len(us6)
    tops6 = sorted({round(_box(s)[1], 2) for s in lines6})
    sep6 = (len(us6) >= 8) or (len(tops6) >= 8)
    _score(n6 >= 8 and sep6,
           "直线形状%d条、下划线%d条，独立纵向位置%d个" % (len(lines6), len(us6), len(tops6)))

    # +1：第六页每条新增横线横向长度≥5厘米
    len6 = [_line_length_cm(s) for s in lines6]
    len6 += [min(_est_width_cm(u[0], u[1]), u[2][2] - u[2][0]) for u in us6]
    _score(bool(len6) and all(v >= 5.0 for v in len6),
           "各横线长度=%s（厘米）" % "/".join("%.2f" % v for v in len6) if len6
           else "未发现新增横线")

    # +1：第六页左侧插图、四个标签文字与四个"→"箭头符号均完整保留
    labels6 = ('Clear sky', 'Storm wind', 'Huge wave', 'New island')
    miss6 = [k for k in labels6 if not shapes_with(S6, k)]
    n_arrow6 = sum(t.count('→') for _s, t in texts(S6))
    _score(len(_pictures(S6)) >= 1 and not miss6 and n_arrow6 >= 4,
           "插图%d张，缺失标签=%s，箭头%d个"
           % (len(_pictures(S6)), "／".join(miss6) or "无", n_arrow6))

    # ---- 第八/九页：按"小标题→对应内容"配对，再查动画
    def group_pairs(slide, sub_needles):
        """[(小标题文案, 小标题形状, 对应内容形状)]。
        内容形状按版式二选一：①同一行且在小标题右侧（顶边差≤0.5厘米）的最近文本
        框（左右分栏排法）；②小标题正下方、左边界对齐（±2.5厘米）的最近文本框
        （上下堆叠排法）。字数<3 的对象一并排除，从而排除 A/B/C/D 字母标记与
        页标题/副标题/页码。"""
        tx = texts(slide)
        excl = page_num_ids(slide) | header_ids(slide)
        subs = []
        for n in sub_needles:
            for sh, t in tx:
                if t.lower() == n.lower() and _sid(sh) not in excl:
                    subs.append((n, sh))
                    break
        sub_ids = {_sid(sh) for _n, sh in subs}
        pairs = []
        for n, sub in subs:
            sb = _box(sub)
            cands = []
            for sh, t in tx:
                if _sid(sh) in excl or _sid(sh) in sub_ids or len(t) < 3:
                    continue
                b = _box(sh)
                if b is None or sb is None:
                    continue
                if abs(b[1] - sb[1]) <= 0.5 and b[0] > sb[0] + 0.1:
                    cands.append((0, b[0], sh))          # 同行右侧
                elif b[1] > sb[1] + 0.1 and abs(b[0] - sb[0]) <= 2.5:
                    cands.append((1, b[1], sh))          # 正下方
            cands.sort(key=lambda c: (c[0], c[1]))
            pairs.append((n, sub, cands[0][2] if cands else None))
        return pairs

    def anim_report(slide, sub_needles):
        """返回 (8个对象的动画明细, 缺动画说明, 触发方式/顺序说明)。
        明细项为 (组名, 角色, 效果标识, mainSeq 单击组序号)。"""
        nodes, first = _entr_index(slide)
        objs, miss, seq = [], [], []
        for n, sub, cont in group_pairs(slide, sub_needles):
            idxs = {}
            for role, shp in (("小标题", sub), ("内容", cont)):
                if shp is None:
                    miss.append("%s-%s未找到" % (n, role))
                    continue
                i = first.get(_sid(shp))
                if i is None:
                    miss.append("%s-%s无进入动画" % (n, role))
                    continue
                idxs[role] = i
                objs.append((n, role, nodes[i]['effect'], nodes[i]['group']))
            if "小标题" in idxs and "内容" in idxs and idxs["小标题"] > idxs["内容"]:
                seq.append("%s的内容早于小标题播放" % n)
        return objs, miss, seq

    _SUBS8 = ('Speech bubble', 'Caption', 'Sound effect', 'Thought bubble')
    _SUBS9 = ('Caption', 'Speech bubble', 'Thought bubble', 'Sound effect')
    objs8, miss8, seq8 = anim_report(S8, _SUBS8)
    objs9, miss9, seq9 = anim_report(S9, _SUBS9)

    # +5：第八页4组小标题与其对应释义文字均添加了进入动画
    # 与办公软件对齐：进入动画=p:timing 里 p:cTn/@presetClass="entr"（或
    # animEffect/@transition="in"），作用对象由 p:spTgt/@spid 指向的形状 id 确定。
    _score(len(objs8) == 8 and not miss8,
           ("；".join(miss8)) if miss8 else "8个对象均已添加进入动画")

    # +5：第九页4组小标题与其对应例句均添加了进入动画
    _score(len(objs9) == 8 and not miss9,
           ("；".join(miss9)) if miss9 else "8个对象均已添加进入动画")

    # +5：第八页8个对象对应8种互不相同的效果
    # 与办公软件对齐：效果种类取 p:cTn/@presetID（动画窗格里显示的效果名就是它），
    # 无 presetID 时退化取 p:animEffect/@filter 基名；同组小标题与释义若同名即不算
    # 不同。
    eff8 = [e for _n, _r, e, _g in objs8]
    _score(len(objs8) == 8 and not miss8 and len(set(eff8)) == 8,
           "实测效果=%s（去重后%d种）" % ("／".join(eff8) or "无", len(set(eff8))))

    # +5：第九页8个对象对应8种互不相同的效果
    eff9 = [e for _n, _r, e, _g in objs9]
    _score(len(objs9) == 8 and not miss9 and len(set(eff9)) == 8,
           "实测效果=%s（去重后%d种）" % ("／".join(eff9) or "无", len(set(eff9))))

    # +1：第八、九页动画均为单击触发，并按"小标题→对应内容"的先后顺序依次播放
    # 与办公软件对齐：p:timing 里 mainSeq 的每个直接子节点＝"一次单击放出的一组
    # 动画"，故 8 个对象必须落在 8 个不同的单击组里才算逐个单击触发；被打包进同
    # 一组的对象会随一次单击同时出现。顺序按小标题所在组序号早于其内容判定。
    def click_check(objs, tag):
        msg = []
        gs = [g for _n, _r, _e, g in objs]
        merged = sum(1 for g in set(gs) if gs.count(g) > 1)
        if len(objs) == 8 and merged:
            msg.append("%s有%d次单击同时放出多个对象" % (tag, merged))
        by = {}
        for n, r, _e, g in objs:
            by.setdefault(n, {})[r] = g
        for n, d in by.items():
            if "小标题" in d and "内容" in d and d["小标题"] >= d["内容"]:
                msg.append("%s%s未按小标题→内容顺序" % (tag, n))
        return msg

    click_msg = click_check(objs8, "第八页") + click_check(objs9, "第九页")
    r23_ok = (len(objs8) == 8 and len(objs9) == 8 and not miss8 and not miss9
              and not click_msg and not seq8 and not seq9)
    _score(r23_ok, "；".join(click_msg[:6] + seq8 + seq9) or "均为单击逐个触发且顺序正确")

    # +5：旧"Build a clear structure"整页已删除，并在第10页位置新增一页，仍18页
    page_text = {i: " ".join(t for _s, t in texts(s)) for i, s in enumerate(slides, 1)}
    old_pages = [i for i, v in page_text.items()
                 if 'build a clear structure' in v.lower()
                 or 'character→change→sequence' in _squash(v).lower()]
    new10 = shapes_with(S10, 'Match five stages of the story with frames')
    _score(len(slides) == 18 and not old_pages and bool(new10),
           "总页数%d，旧页残留=%s，第10页新标题=%s"
           % (len(slides), old_pages or "无", "有" if new10 else "无"))

    # +3：新增第10页标题文字正确且位于页面顶部标题位置（顶边≤1.5厘米）
    b10t = _box(new10[0]) if new10 else None
    _score(bool(new10) and b10t is not None and b10t[1] <= 1.5,
           "未找到标题" if b10t is None else "标题顶边%.2fcm" % b10t[1])

    # +5：新增第10页在标题下方插入故事发展曲线图，宽19.8-20.2厘米且未被拉伸变形
    pics10 = _pictures(S10)
    pic10, r26_ok, r26_msg = None, False, []
    for pp in pics10:
        w = _cm(pp.width)
        d = _pic_aspect_delta(pp)
        r26_msg.append("宽%.2fcm，宽高比偏差%s"
                       % (w or 0.0, "未知" if d is None else "%.1f%%" % (d * 100)))
        if w is not None and 19.8 <= w <= 20.2:
            if pic10 is None:
                pic10 = pp
            if d is None or d <= 0.05:
                pic10, r26_ok = pp, True
                break
    if pic10 is None and pics10:
        pic10 = pics10[0]
    _score(r26_ok, "；".join(r26_msg) or "该页未发现图片")

    # +1：新增第10页图片顶边≥3.5厘米（在标题下方）且在页面内水平居中
    b10p = _box(pic10) if pic10 is not None else None
    r27_ok = (b10p is not None and b10p[1] >= 3.5 and in_page(b10p)
              and abs(b10p[0] - (page_w - b10p[2])) <= 1.0)
    _score(r27_ok, "无图片" if b10p is None else
           "图片顶边%.2fcm，左留白%.2fcm，右留白%.2fcm"
           % (b10p[1], b10p[0], page_w - b10p[2]))

    # +1：新增第10页标题字号与其余内容页标题一致，为25磅
    sz10 = sizes_of(new10)
    _score(near(sz10, 25.0, 25.0),
           "标题字号=%s磅" % ("/".join("%g" % v for v in sz10 if v is not None) or "?"))

    # +1：新增第10页右下角保留与页序一致的"10"页码标记
    pn10 = [t for sh, t in texts(S10) if _sid(sh) in page_num_ids(S10)]
    _score('10' in pn10, "右下角页码=%s" % ("／".join(pn10) or "无"))

    # +1：新增第10页背景与其余内容页一致，为米黄色FFF7E8
    bg10 = _bg_hex(S10)
    _score(bg10 == 'FFF7E8', "第10页背景=%s" % bg10)

    # +5：第十一页6个对话文本框全部添加了进入动画
    _BUBBLES = ('A storm is coming!', 'We are in trouble.',
                'Look! There is an island.', 'Quick! Row faster!',
                'Thank goodness!', 'We made it!')
    nodes11, first11 = _entr_index(S11)
    bub11, miss11 = [], []
    for n in _BUBBLES:
        shp = None
        for sh in shapes_with(S11, n):
            shp = sh
            break
        if shp is None:
            miss11.append("\"%s\"未找到" % n)
            continue
        i = first11.get(_sid(shp))
        if i is None:
            miss11.append("\"%s\"无进入动画" % n)
            continue
        bub11.append((n, nodes11[i]['effect'], nodes11[i]['group']))
    _score(len(bub11) == 6 and not miss11,
           "；".join(miss11) or "6个对话文本框均已添加进入动画")

    # +3：第十一页6个对话文本框使用同一种进入动画效果，且为单击逐个触发
    # 与办公软件对齐：逐个单击=6个对话框分别落在 mainSeq 的 6 个不同单击组里。
    eff11 = {e for _n, e, _g in bub11}
    grp11 = [g for _n, _e, g in bub11]
    solo11 = (len(set(grp11)) == len(grp11))
    _score(len(bub11) == 6 and len(eff11) == 1 and solo11,
           "效果=%s；单击组序号=%s"
           % ("／".join(sorted(eff11)) or "无",
              "/".join(str(g) for g in grp11) or "无"))

    # +1：第十一页主标题"While writing"未被添加动画
    anim11 = _all_anim_spids(S11)
    title11 = shapes_with(S11, 'While writing')
    _score(bool(title11) and all(_sid(s) not in anim11 for s in title11),
           "未找到主标题" if not title11 else
           ("主标题被加入动画" if any(_sid(s) in anim11 for s in title11)
            else "主标题未参与动画"))

    # +3：第十六页右下角插图已删除，该位置无残留图片或空白图片框
    # 与办公软件对齐：删除干净=该页不再有 p:pic，也没有留在原位（横向21.3厘米、
    # 纵向13.5厘米附近）的空图片占位框。
    res16 = []
    for pp in _pictures(S16):
        b = _box(pp)
        res16.append("残留图片(%.2f,%.2f)" % (b[0], b[1]) if b else "残留图片")
    for sh in S16.shapes:
        b = _box(sh)
        if b is None or _shape_text(sh) or _local(sh._element) != 'sp':
            continue
        if 20.0 <= b[0] <= 23.0 and 12.0 <= b[1] <= 15.0 \
                and (b[2] - b[0]) >= 6.0 and (b[3] - b[1]) >= 2.0:
            res16.append("残留空图片框(%.2f,%.2f)" % (b[0], b[1]))
    _score(not res16, "；".join(res16) or "右下角插图已删除且无残留")

    # +1：第十六页"Tense／Person／Voice"三行共9个文本框完整保留
    miss16 = [k for k in ('Tense', 'Person', 'Voice') if not shapes_with(S16, k)]
    grid16 = [sh for sh, _t in texts(S16)
              if _sid(sh) not in page_num_ids(S16) | header_ids(S16)]
    _score(not miss16 and len(grid16) >= 9,
           "缺失行标签=%s，内容区文本框%d个" % ("／".join(miss16) or "无", len(grid16)))

    # +5：全部18页均设置了页面切换动画
    # 与办公软件对齐：切换效果写在 p:sld/p:transition 下（新式效果放在
    # mc:AlternateContent 的 mc:Choice 里），即"切换"选项卡当前选中的效果。
    trans = [_transition_key(s) for s in slides]
    no_tr = [i for i, v in enumerate(trans, 1) if not v]
    _score(len(slides) == 18 and not no_tr,
           ("无切换的页：" + "、".join("第%d页" % i for i in no_tr)) if no_tr
           else "18页均已设置切换效果")

    # +5：18页的切换动画效果两两互不相同，共18种
    uniq_tr = {v for v in trans if v}
    dup_tr = sorted({v for v in trans if v and trans.count(v) > 1})
    _score(len(slides) == 18 and not no_tr and len(uniq_tr) == 18,
           "去重后%d种；重复效果=%s" % (len(uniq_tr), "／".join(dup_tr) or "无"))

    # +1：交付稿仍为18页，页序与各页主题保持一致
    order_ok = (len(slides) == 18
                and 'story writing lab' in page_text.get(1, '').lower()
                and 'writing' in page_text.get(2, '').lower()
                and 'learning objectives' in page_text.get(3, '').lower()
                and 'final checklist' in page_text.get(18, '').lower())
    _score(order_ok, "共%d页；首页/第2页/第3页/末页主题匹配=%s"
           % (len(slides), order_ok))

    # ---------- 减分项：指令未涉及、却被一并改坏的项 ----------
    # -5：第二至十五页正文西文字体被写死为"Arial"，未随主题字体（Aptos Display／
    # Aptos）
    # 与办公软件对齐：随主题=a:latin/@typeface 取 "+mj-lt"/"+mn-lt" 或干脆不写；
    # 一旦写成具体字体名，改主题字体就不再生效。
    hard_font = []
    for n in (2, 3, 4, 5, 6, 8, 9, 10, 11, 12, 13, 14, 15):
        s = sl(n)
        if s is None:
            continue
        fonts = set()
        for sh in s.shapes:
            for tf in _latin_typefaces(sh):
                if tf not in _THEME_FONT_REFS:
                    fonts.add(tf)
        if fonts:
            hard_font.append("第%d页=%s" % (n, "／".join(sorted(fonts))))
    _score(bool(hard_font), "；".join(hard_font) or "各页西文字体均随主题字体")

    # -3：第三页5条学习目标文本框被压窄，宽度不再是25.15厘米且文字折行
    w_bad = ["%.2fcm" % (_cm(sh.width) or 0.0) for sh, _t in obj3
             if sh.width is None or not (25.0 <= (_cm(sh.width) or 0.0) <= 25.3)]
    trig40 = len(obj3) != 5 or bool(w_bad) or bool(fold3)
    _score(trig40, "实测框宽异常=%s；折行=%s"
           % ("／".join(w_bad) or "无", "；".join(fold3) or "无")
           if trig40 else "5条目标框宽25.15厘米且均单行显示")

    # -3：第五页底部提示句被改字号并上移（应为22磅、顶边纵向14.35厘米）
    trig41, r41_msg = True, "未找到该句"
    if bottom5:
        shb = bottom5[0]
        bb = _box(shb)
        szb = sizes_of([shb])
        r41_msg = "字号=%s磅，顶边=%s" % (
            "/".join("%g" % v for v in szb if v is not None) or "?",
            "?" if bb is None else "%.2fcm" % bb[1])
        trig41 = not (near(szb, 22.0, 22.0) and bb is not None
                      and abs(bb[1] - 14.35) <= 0.15)
    _score(trig41, r41_msg)

    # -1：第一页右下角页码"01"字号被改（应为8磅）
    pn1 = [sh for sh, t in texts(S1) if t == '01']
    sz_pn1 = sizes_of(pn1)
    trig42 = not near(sz_pn1, 8.0, 8.0)
    _score(trig42, "\"01\"字号=%s磅"
           % ("/".join("%g" % v for v in sz_pn1 if v is not None) or "未找到"))

    # -1：第九页4条装饰分隔线被一并加入动画播放序列
    # 与办公软件对齐：分隔线是无文字的直线/极扁形状；只要其 shape id 出现在
    # p:timing 的任一 p:spTgt/@spid 上，就会在动画窗格里占一行并参与播放。
    anim9 = _all_anim_spids(S9)
    div9 = [s for s in S9.shapes if _is_line_shape(s) and not _shape_text(s)]
    hit9 = [s for s in div9 if _sid(s) in anim9]
    trig43 = bool(hit9)
    _score(trig43, "参与动画的分隔线%d条（该页分隔线共%d条）%s"
           % (len(hit9), len(div9),
              "，纵向位置=" + "／".join("%.2fcm" % _box(s)[1] for s in hit9
                                    if _box(s) is not None) if hit9 else ""))

    # -1：第十页遗留无任何文字的空文本框
    # 与办公软件对齐：只统计"真正的空文本框"——有 txBody 但无可见文字，且尺寸
    # 达到正常文本框量级（宽≥1厘米、高≥0.5厘米），从而排除装饰用的极扁分隔线。
    empty10 = []
    for sh in S10.shapes:
        b = _box(sh)
        try:
            if not sh.has_text_frame:
                continue
        except Exception:
            continue
        if _shape_text(sh) or b is None:
            continue
        if (b[2] - b[0]) >= 1.0 and (b[3] - b[1]) >= 0.5:
            empty10.append("顶边%.2fcm、%.2f×%.2f厘米"
                           % (b[1], b[2] - b[0], b[3] - b[1]))
    trig44 = bool(empty10)
    _score(trig44, "；".join(empty10) or "无空白元素残留")

    # ========== 将 results 反解为结构化 dim2_items ==========
    # 逐项按追加顺序对齐 _MAX_DELTAS_ORDER：解析每条 "+N：说明" / "-N：说明" 的
    # N 与 rule。加分项 N=max_delta 表示命中；减分项 N=max_delta（负值）表示触发
    # 扣分；两者未命中/未触发均写作 "+0"/"-0"。
    if len(results) != len(_MAX_DELTAS_ORDER):
        raise RuntimeError(
            "dim2 项数不一致：results=%d vs order=%d"
            % (len(results), len(_MAX_DELTAS_ORDER)))
    dim2_items = []
    total_score = 0
    for raw, max_delta in zip(results, _MAX_DELTAS_ORDER):
        head, _sep, rule = raw.partition("：")
        try:
            delta = int(head.strip())
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
    result["total_score"] = total_score
    result["max_score"] = MAX_SCORE_TOTAL
    return result

if __name__ == "__main__":
    # 本地调试入口：evaluate 只接收"脚本所在目录的路径"，由脚本自己在该目录里
    # 定位并打开被评估的 .pptx。命令行第 1 个参数可覆盖目录路径；未指定时使用
    # 脚本自身所在目录。
    if len(sys.argv) >= 2:
        _dir = sys.argv[1]
    else:
        _dir = os.path.dirname(os.path.abspath(__file__))
    _out = evaluate(_dir)
    print(json.dumps(_out, ensure_ascii=False, indent=2))




























