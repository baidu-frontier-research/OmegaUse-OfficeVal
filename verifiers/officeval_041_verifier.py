#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# pyright: reportUnknownMemberType=false, reportUnknownVariableType=false, reportUnknownArgumentType=false, reportUnknownParameterType=false, reportMissingParameterType=false, reportPrivateUsage=false, reportUnknownLambdaType=false, reportGeneralTypeIssues=false

import os
import re
import sys
import json
import zipfile
import traceback
from pptx import Presentation
from lxml import etree

# 说明：按接口统一约定，本模块不修改全局 sys.stdout，也不 print 主结果。
# 主结果统一由 evaluate(dir_path) -> dict 返回；入参为"脚本所在目录的路径"，
# 由脚本自己在该目录内定位并打开被评估的 .pptx。本文件仅在 __main__ 下作为
# 本地调试入口，才把 dict 序列化为 JSON 打印。

SCRIPT_ID = "041"
# 维度二加分项 max_delta 之和；用于维度一未通过 / 脚本异常时兜底
MAX_SCORE_TOTAL = 78
# 维度二减分项 max_delta 绝对值之和（仅用于自检，不参与 max_score）
PENALTY_TOTAL = 21

# 维度二所有评分项按"结果追加顺序"的 max_delta 列表；末端把 results 里的
# "+N：..." / "-N：..." 字符串反解为结构化 dim2_items 时用。顺序须与
# _evaluate_impl 中 results.append(...) 的顺序严格一致；共 47 项
# （34 条加分项累加 = MAX_SCORE_TOTAL，13 条减分项累加 = -PENALTY_TOTAL）。
_MAX_DELTAS_ORDER = [
    3,   # 01 共23页且目录页在第2页、其余页序不乱
    1,   # 02 目录页"目录"40磅 + 六板块30磅
    1,   # 03 三处"01  背景"角标20磅
    1,   # 04 "02  目标"角标20磅
    1,   # 05 两处"03  需求"角标20磅
    1,   # 06 两处"04  目标"角标20磅
    3,   # 07 两处角标改"05  验证"且20磅
    3,   # 08 两处角标改"06  方案"且20磅
    3,   # 09 四处角标改"07  实施"且20磅
    3,   # 10 两处角标改"08  验证"且20磅
    1,   # 11 一处角标改"09  巩固"且20磅
    3,   # 12 风险机制页插入三张配图且等大
    1,   # 13 三张配图淡蓝色边框
    1,   # 14 三页表格文字微软雅黑10磅
    5,   # 15 视频进入本页自动播放
    1,   # 16 "分离""观察""维护"15磅
    3,   # 17 创新借鉴页三图统一8.00×5.00cm
    3,   # 18 目标设定页底句扩至50字且20磅
    5,   # 19 可行性论证页图片进入动画
    3,   # 20 验证结论扩至50字
    5,   # 21 筛分路径验证页图片"轮子"动画
    1,   # 22 模拟试验二说明句20磅
    3,   # 23 设计原则扩为五条
    1,   # 24 研发指标页表格20磅
    1,   # 25 研发指标页说明句20磅
    1,   # 26 "项目协调负责人：沈亦然"
    1,   # 27 对策实施一页三句15磅
    5,   # 28 对策实施二页图片"渐变"动画
    5,   # 29 对策实施三页图片"溶解"动画
    5,   # 30 试运行页表格数据替换为Excel内容
    1,   # 31 试运行页表格24.07×10.59cm
    1,   # 32 效果验证页红、黄两圆形
    1,   # 33 两处说明语句20磅
    1,   # 34 "总结与展望""感谢聆听"40磅
    -3,  # 35 风险机制页须保留14.35×13.54cm示意图
    -3,  # 36 试运行页须保留监测画面图片
    -3,  # 37 试运行页长句内容与宋体12.8磅加粗111827
    -1,  # 38 试运行页方案C列两格须空白
    -3,  # 39 十一处改号角标须加粗+00A6A6
    -1,  # 40 筛分路径验证页不得新增小圆形
    -1,  # 41 验证结论正文须13.0磅+334155
    -1,  # 42 验证结论文本框须≥9.00×2.00cm
    -1,  # 43 目标设定页底句须加粗+00A6A6
    -1,  # 44 设计原则须12.5磅加粗0A1F33
    -1,  # 45 制定对策页表格须8.7磅
    -1,  # 46 处置侧圆形须2.08cm+10B981
    -1,  # 47 改造前页须保留视频对象
]
assert sum(d for d in _MAX_DELTAS_ORDER if d > 0) == MAX_SCORE_TOTAL
assert -sum(d for d in _MAX_DELTAS_ORDER if d < 0) == PENALTY_TOTAL

NS = {
    'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
    'p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
}
A = '{%s}' % NS['a']
P = '{%s}' % NS['p']

EMU_PER_CM = 360000.0
# 字体别名：办公软件里显示为"微软雅黑"/"宋体"的实际 typeface 名
YAHEI = ('微软雅黑', 'Microsoft YaHei', 'Microsoft YaHei UI', 'MicrosoftYaHei')
SONGTI = ('宋体', 'SimSun', 'NSimSun')

# 23 页页序基准（用于维度一页数与第1项页序判定）；封面按前缀匹配
SLIDE_TITLES = [
    '高速成型线碎屑分离',
    '目录',
    '项目定位',
    '风险机制：三个要素相互叠加',
    '改造前：回收路径长、观察节点少',
    '改造目标：把风险链拆开',
    '需求转换：从现场声音到关键需求',
    '创新借鉴：分割与预先感知',
    '目标设定：用可量化指标约束设计',
    '可行性论证：从工艺、技术、资源三方面评估',
    '模拟试验一：筛分路径验证',
    '模拟试验二：粉料含量与异常蔓延趋势',
    '总体方案：物理分路',
    '研发指标：功能、性能与兼容性',
    '制定对策：任务、目标与交付物',
    '对策实施一：筛分系统设计与加工',
    '对策实施二：可开合观察窗制作',
    '对策实施三：监测画面接入操作位',
    '整体试运行：连续班次结果',
    '效果验证：风险链条被削弱',
    '标准化与巩固措施',
    '总结与展望',
    '感谢聆听',
]

# 试运行页表格应有的内容（来自所给 Excel）；None 表示该格应保持空白
TRIAL_TABLE = [
    ['评估项目', '方案A：紧凑型产线', '方案B：标准化产线', '方案C：智能柔性产线'],
    ['单线产能', '约7,200件/小时', '约9,600件/小时', None],
    ['工艺配置', '在线检测+分段除尘', '常规传送+集中抽吸', '视觉识别+闭环调节+分级回收'],
    ['维护空间', '侧向检修通道约0.8米', '后部集中维护，需错峰停机', None],
    ['运行稳定性', '适合中等批量连续生产', '高负荷下需关注物料波动', '可按订单自动切换参数，波动响应较快'],
    ['综合关注点', '能耗与噪声控制', '备件兼容性与清洁周期', '粉尘积聚、温升预警与盲区巡检联动'],
]

def _cm(emu):
    """EMU → 厘米；None 安全。"""
    if emu is None:
        return None
    return float(emu) / EMU_PER_CM


def _near(val, target, tol):
    return val is not None and abs(val - target) <= tol


def _norm(text):
    """规范化文本：去掉所有空白与不可见字符，便于内容比对。"""
    if not text:
        return ''
    return re.sub(r'[\s ​]+', '', text)


def _cjk_len(text):
    """语句长度按"汉字+可见标点"计，用于"扩充至五十字"判定。"""
    return len(re.sub(r'\s+', '', text or ''))


def _iter_shapes(container):
    """递归遍历（含组合内的）所有形状。"""
    for shp in container.shapes:
        yield shp
        if shp.shape_type == 6 or shp.__class__.__name__ == 'GroupShape':
            try:
                for sub in _iter_shapes(shp):
                    yield sub
            except Exception:
                pass


def _shape_text(shape):
    """形状可见文字（含表格）。"""
    try:
        if getattr(shape, 'has_table', False) and shape.has_table:
            cells = []
            for row in shape.table.rows:
                for cell in row.cells:
                    cells.append(cell.text)
            return '\n'.join(cells)
        if getattr(shape, 'has_text_frame', False) and shape.has_text_frame:
            return shape.text_frame.text
    except Exception:
        pass
    return ''


def _slide_text(slide):
    return '\n'.join(_shape_text(s) for s in _iter_shapes(slide))


def _text_shapes(slide):
    """仅含文字（非表格）的形状列表。"""
    out = []
    for shp in _iter_shapes(slide):
        try:
            if getattr(shp, 'has_text_frame', False) and shp.has_text_frame and shp.text_frame.text.strip():
                out.append(shp)
        except Exception:
            continue
    return out

def _runs(shape):
    """形状内所有 run（文本框）。"""
    out = []
    try:
        if getattr(shape, 'has_text_frame', False) and shape.has_text_frame:
            for para in shape.text_frame.paragraphs:
                out.extend(para.runs)
    except Exception:
        pass
    return out


def _table_runs(table):
    out = []
    for row in table.rows:
        for cell in row.cells:
            for para in cell.text_frame.paragraphs:
                out.extend(para.runs)
    return out


def _slide_table_runs(slide):
    """页面内所有表格的 run 汇总。"""
    out = []
    for tbl in _tables(slide):
        out.extend(_table_runs(tbl))
    return out


def _pixels(picture):
    """图片的像素尺寸 (宽, 高)，取不到时返回 (0, 0)。"""
    try:
        return tuple(picture.image.size)
    except Exception:
        return (0, 0)


def _sha1(picture):
    """图片二进制的 sha1，用于判断三张配图互不相同。"""
    try:
        return picture.image.sha1
    except Exception:
        return ''


def _run_typefaces(run):
    """run 实际写明的 typeface 集合（latin / ea / cs 三处），空集表示继承。"""
    names = set()
    try:
        rPr = run._r.find(A + 'rPr')
        if rPr is not None:
            for tag in ('latin', 'ea', 'cs'):
                el = rPr.find(A + tag)
                if el is not None and el.get('typeface'):
                    names.add(el.get('typeface'))
    except Exception:
        pass
    try:
        if run.font.name:
            names.add(run.font.name)
    except Exception:
        pass
    return names


def _font_ok(run, family):
    """run 的字体是否属于给定字体族（别名等价）。未写明字体视为不达标。"""
    names = _run_typefaces(run)
    if not names:
        return False
    return any(n in family for n in names)


def _size_pt(run):
    try:
        return None if run.font.size is None else round(run.font.size.pt, 2)
    except Exception:
        return None


def _bold(run):
    try:
        return run.font.bold
    except Exception:
        return None


def _rgb(run):
    """run 字体颜色的 RRGGBB；主题色/继承色返回 None。"""
    try:
        color = run.font.color
        if color is None or color.type is None:
            return None
        return str(color.rgb).upper()
    except Exception:
        return None

def _runs_all(runs, family=None, size=None, bold=None, rgb=None, size_tol=0.05):
    """批量校验一组 run 的字体/字号/加粗/颜色；runs 为空一律不达标。"""
    if not runs:
        return False
    for run in runs:
        if not run.text.strip():
            continue
        if family is not None and not _font_ok(run, family):
            return False
        if size is not None and not _near(_size_pt(run), size, size_tol):
            return False
        if bold is not None and _bold(run) is not bold:
            return False
        if rgb is not None and _rgb(run) != rgb:
            return False
    return True


def _sizes_of(runs):
    """一组 run 的字号取值集合（用于把实测值写进结果说明）。"""
    return sorted({s for s in (_size_pt(r) for r in runs if r.text.strip()) if s is not None})


def _find_shape_by_text(slide, key, exact=False):
    """按文字定位形状：exact 时要求整体文字相等，否则按包含匹配。"""
    for shp in _text_shapes(slide):
        txt = _shape_text(shp)
        if exact:
            if _norm(txt) == _norm(key):
                return shp
        elif _norm(key) in _norm(txt):
            return shp
    return None


def _locate(slides, key):
    """按标题文字定位页序号（0 起）；优先整段等于 key 的形状，其次包含匹配。"""
    for idx, slide in enumerate(slides):
        for shp in _text_shapes(slide):
            if _norm(_shape_text(shp)) == _norm(key):
                return idx
    for idx, slide in enumerate(slides):
        if _norm(key) in _norm(_slide_text(slide)):
            return idx
    return -1


def _pictures(slide):
    """页内图片（不含视频占位图）。"""
    out = []
    for shp in _iter_shapes(slide):
        try:
            if shp.shape_type is not None and shp.element.find('.//' + A + 'blip') is not None:
                if shp.element.find('.//' + A + 'videoFile') is None and shp.element.tag == P + 'pic':
                    out.append(shp)
        except Exception:
            continue
    return out


def _media_shapes(slide):
    """页内视频/媒体对象。"""
    out = []
    for shp in _iter_shapes(slide):
        try:
            if shp.element.find('.//' + A + 'videoFile') is not None:
                out.append(shp)
        except Exception:
            continue
    return out


def _tables(slide):
    return [s.table for s in _iter_shapes(slide) if getattr(s, 'has_table', False) and s.has_table]

def _table_shape(slide):
    for shp in _iter_shapes(slide):
        if getattr(shp, 'has_table', False) and shp.has_table:
            return shp
    return None


def _corner_shape(slide):
    """右上角标形状：文字形如"01  背景"（两位数字 + 空格 + 两字），且位于页面右上。"""
    best = None
    for shp in _text_shapes(slide):
        txt = _shape_text(shp).strip()
        if re.match(r'^\d{2}\s+\S{2}$', txt):
            if best is None or (shp.top or 0) < (best.top or 0):
                best = shp
    return best


def _theme_colors(pptx_path):
    """解析 ppt/theme/theme1.xml 的配色方案：{'accent1':'RRGGBB', ...}。"""
    out = {}
    try:
        with zipfile.ZipFile(pptx_path) as zf:
            names = [n for n in zf.namelist() if re.match(r'ppt/theme/theme\d+\.xml$', n)]
            if not names:
                return out
            root = etree.fromstring(zf.read(sorted(names)[0]))
            scheme = root.find('.//' + A + 'clrScheme')
            if scheme is None:
                return out
            for child in scheme:
                key = etree.QName(child).localname
                srgb = child.find(A + 'srgbClr')
                sysc = child.find(A + 'sysClr')
                if srgb is not None:
                    out[key] = (srgb.get('val') or '').upper()
                elif sysc is not None:
                    out[key] = (sysc.get('lastClr') or '').upper()
    except Exception:
        pass
    return out


def _fill_rgb(shape, theme):
    """形状填充色 RRGGBB：srgbClr 直接取值，schemeClr 用主题配色解析（忽略明暗微调）。"""
    try:
        spPr = shape.element.find(P + 'spPr')
        if spPr is None:
            spPr = shape.element.find('.//' + A + 'spPr')
        if spPr is None:
            return None
        solid = spPr.find(A + 'solidFill')
        if solid is None:
            return None
        srgb = solid.find(A + 'srgbClr')
        if srgb is not None:
            return (srgb.get('val') or '').upper()
        sch = solid.find(A + 'schemeClr')
        if sch is not None:
            key = sch.get('val') or ''
            key = {'tx1': 'dk1', 'tx2': 'dk2', 'bg1': 'lt1', 'bg2': 'lt2'}.get(key, key)
            return theme.get(key)
    except Exception:
        pass
    return None

def _hex_to_rgb(hexstr):
    if not hexstr or len(hexstr) != 6:
        return None
    try:
        return (int(hexstr[0:2], 16), int(hexstr[2:4], 16), int(hexstr[4:6], 16))
    except ValueError:
        return None


def _is_red(hexstr):
    rgb = _hex_to_rgb(hexstr)
    if rgb is None:
        return False
    r, g, b = rgb
    return r >= 140 and r >= g * 2 and r >= b * 2


def _is_yellow(hexstr):
    rgb = _hex_to_rgb(hexstr)
    if rgb is None:
        return False
    r, g, b = rgb
    return r >= 190 and g >= 140 and b <= 130 and abs(r - g) <= 90


def _is_light_blue(hexstr):
    rgb = _hex_to_rgb(hexstr)
    if rgb is None:
        return False
    r, g, b = rgb
    return b >= 190 and b >= r + 20 and b >= g + 5 and r >= 90 and g >= 130


def _line_light_blue(shape, theme):
    """形状描边是否为淡蓝色实线、线宽 0.75-3 磅（9525-38100 EMU）。"""
    try:
        ln = shape.element.find('.//' + A + 'ln')
        if ln is None:
            return False, '未设置边框'
        if ln.find(A + 'noFill') is not None:
            return False, '边框为无填充'
        width = ln.get('w')
        width_pt = round(int(width) / 12700.0, 2) if width else None
        if width_pt is not None and not (0.75 - 0.01 <= width_pt <= 3.0 + 0.01):
            return False, '线宽%.2f磅' % width_pt
        solid = ln.find(A + 'solidFill')
        if solid is None:
            return False, '边框未设置实线颜色'
        srgb = solid.find(A + 'srgbClr')
        if srgb is not None:
            val = (srgb.get('val') or '').upper()
            return _is_light_blue(val), val
        sch = solid.find(A + 'schemeClr')
        if sch is not None:
            key = sch.get('val') or ''
            base = theme.get({'tx1': 'dk1', 'tx2': 'dk2', 'bg1': 'lt1', 'bg2': 'lt2'}.get(key, key))
            # 主题色 + lumMod/lumOff 提亮 → 视为该主题色的淡色变体
            lighten = sch.find(A + 'lumOff') is not None or sch.find(A + 'lumMod') is not None
            if base and _is_light_blue(base):
                return True, base
            if base and lighten and _hex_to_rgb(base) and _hex_to_rgb(base)[2] >= 140:
                return True, base + '(淡色变体)'
            return False, str(base)
    except Exception as exc:
        return False, 'ln解析失败:%s' % exc
    return False, '未设置边框'

def _slide_xml_map(pptx_path):
    """{幻灯片序号(1起): 该页 slideN.xml 的根元素}；p:timing 只能从原始 XML 读。"""
    out = {}
    with zipfile.ZipFile(pptx_path) as zf:
        for name in zf.namelist():
            m = re.match(r'ppt/slides/slide(\d+)\.xml$', name)
            if m:
                try:
                    out[int(m.group(1))] = etree.fromstring(zf.read(name))
                except Exception:
                    continue
    return out


def _slide_order(pptx_path):
    """按 presentation.xml 的 sldIdLst 顺序返回 slideN.xml 的 N 列表。"""
    order = []
    with zipfile.ZipFile(pptx_path) as zf:
        try:
            pres = etree.fromstring(zf.read('ppt/presentation.xml'))
            rels = etree.fromstring(zf.read('ppt/_rels/presentation.xml.rels'))
        except Exception:
            return order
        rid_map = {}
        for rel in rels:
            rid_map[rel.get('Id')] = rel.get('Target') or ''
        for sld_id in pres.findall('.//' + P + 'sldIdLst/' + P + 'sldId'):
            rid = sld_id.get('{%s}id' % NS['r'])
            target = rid_map.get(rid, '')
            m = re.search(r'slide(\d+)\.xml$', target)
            if m:
                order.append(int(m.group(1)))
    return order


def _entr_effects(root):
    """页内所有进入类动画：[{'spid','presetID','filters'}]。"""
    out = []
    if root is None:
        return out
    timing = root.find(P + 'timing')
    if timing is None:
        return out
    for ctn in timing.iter(P + 'cTn'):
        if ctn.get('presetClass') != 'entr':
            continue
        spids = [t.get('spid') for t in ctn.iter(P + 'spTgt') if t.get('spid')]
        filters = [e.get('filter') for e in ctn.iter(P + 'animEffect') if e.get('filter')]
        out.append({
            'spid': spids[0] if spids else None,
            'presetID': ctn.get('presetID'),
            'filters': filters,
        })
    return out


def _autoplay_media_spids(root):
    """挂在 mainSeq 下、随页面自动触发播放的媒体目标 spid 集合。"""
    spids = set()
    if root is None:
        return spids
    timing = root.find(P + 'timing')
    if timing is None:
        return spids
    for ctn in timing.iter(P + 'cTn'):
        if ctn.get('nodeType') != 'mainSeq':
            continue
        for sub in ctn.iter(P + 'cTn'):
            if sub.get('presetClass') != 'mediacall':
                continue
            for tgt in sub.iter(P + 'spTgt'):
                if tgt.get('spid'):
                    spids.add(tgt.get('spid'))
    return spids

def _open_presentation(file_path):
    """打开 .pptx 演示文稿（python-pptx 仅支持 OOXML/.pptx）。"""
    return Presentation(file_path)


def evaluate(dir_path):
    """按接口约定返回结构化 dict；不 print、不改 sys.stdout、不 sys.exit。
    入参 dir_path 为"脚本所在目录的路径"，脚本自己在该目录里定位并打开被评估的
    .pptx（忽略以 ~$ 开头的临时锁文件）。脚本自身崩溃（含目录不存在、未发现
    文档等）由本函数捕获并返回 status='error'，此时 dim1_pass 置为空字符串、
    dim1_reason 为空、max_score 归零；维度一未通过时 dim1_pass 为 False 且
    dim1_reason 给出错误摘要。"""
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
    def _as_error(message):
        """统一的错误返回：dim1_pass 置空（未进入维度一判定），max_score 归零。"""
        result["status"] = "error"
        result["error"] = message
        result["dim1_pass"] = ""
        result["dim1_reason"] = ""
        result["dim2_items"] = []
        result["total_score"] = 0
        result["max_score"] = 0
        return result

    try:
        if not dir_path or not os.path.isdir(dir_path):
            return _as_error(f"目录不存在：{dir_path}")
        cands = sorted(
            f for f in os.listdir(dir_path)
            if f.lower().endswith('.pptx') and not f.startswith('~$')
        )
        if not cands:
            return _as_error(f"目录中未找到 .pptx 文件：{dir_path}")
        file_path = os.path.join(dir_path, cands[0])
        result["file_name"] = os.path.basename(file_path)
        return _evaluate_impl(file_path, result)
    except Exception as exc:
        return _as_error(f"{type(exc).__name__}: {exc}")


def _check_dim1(prs, file_path):
    """维度1：可用与可修改性。返回 (pass, reason)。"""
    slides = list(prs.slides)
    if len(slides) != 23:
        return False, "全文共%d页，不满足23页（页数缺失或多增）" % len(slides)
    page_w = _cm(prs.slide_width) or 0.0
    page_h = _cm(prs.slide_height) or 0.0
    for idx, slide in enumerate(slides, start=1):
        texts = _text_shapes(slide)
        tables = _tables(slide)
        pics = _pictures(slide)
        if not texts and not tables:
            return False, "第%d页无任何可编辑文字对象（疑整页粘贴为图片）" % idx
        # 整页位图：单张图片铺满 ≥90% 版面且该页几乎没有可编辑文字
        for pic in pics:
            area = (_cm(pic.width) or 0) * (_cm(pic.height) or 0)
            if page_w and page_h and area >= 0.90 * page_w * page_h and len(texts) <= 1:
                return False, "第%d页整页为一张位图，内容不可单独编辑" % idx
        # 版面严重错乱：形状整体跑出页面之外
        for shp in _iter_shapes(slide):
            left, top = _cm(shp.left), _cm(shp.top)
            width, height = _cm(shp.width), _cm(shp.height)
            if None in (left, top, width, height):
                continue
            if left > page_w or top > page_h or left + width < 0 or top + height < 0:
                return False, "第%d页存在完全位于页面之外的对象，版面错乱" % idx
    return True, ""

def _evaluate_impl(file_path, result):
    """维度一 + 维度二的评估主体；维度二逐项把 "+N：..." / "-N：..." 追加进
    results，末尾按 _MAX_DELTAS_ORDER 反解成结构化 dim2_items 填入 result。"""
    prs = _open_presentation(file_path)
    slides = list(prs.slides)
    theme = _theme_colors(file_path)
    results = []

    # ========== 维度1：可用与可修改性 ==========
    dim1_pass, dim1_reason = _check_dim1(prs, file_path)
    result["dim1_pass"] = dim1_pass
    # 维度一未通过时 dim1_reason 必须给出错误摘要
    result["dim1_reason"] = dim1_reason if dim1_pass else (dim1_reason or "维度一未通过（文档不可用或不可修改）")
    if not dim1_pass:
        result["dim2_items"] = []
        result["total_score"] = 0
        return result

    # ========== 维度2：完成度评分 ==========
    # 每页原始 XML（p:timing / 视频节点只能从 XML 读）
    order = _slide_order(file_path)
    xml_map = _slide_xml_map(file_path)
    roots = [xml_map.get(n) for n in order] if len(order) == len(slides) else [None] * len(slides)

    def page(key):
        """按标题定位页；返回 (index, slide, xml_root)，未找到时 slide 为 None。"""
        idx = _locate(slides, key)
        if idx < 0:
            return -1, None, None
        return idx, slides[idx], (roots[idx] if idx < len(roots) else None)

    def corner_ok(keys, label, size=20.0):
        """一组页面的右上角标：文字等于 label 且为微软雅黑 size 磅。"""
        bad = []
        for key in keys:
            _, slide, _ = page(key)
            if slide is None:
                bad.append('%s(缺页)' % key)
                continue
            shp = _corner_shape(slide)
            if shp is None:
                bad.append('%s(无角标)' % key)
                continue
            txt = _shape_text(shp).strip()
            runs = _runs(shp)
            if _norm(txt) != _norm(label):
                bad.append('%s(实测"%s")' % (key, txt))
            elif not _runs_all(runs, family=YAHEI, size=size):
                bad.append('%s(字号%s)' % (key, _sizes_of(runs) or '继承'))
        return (not bad), '；'.join(bad)

    def entr_hit(root, slide, preset=None, filter_kw=None):
        """页内图片是否被指定的进入动画命中。"""
        pic_ids = {str(p.shape_id) for p in _pictures(slide)} if slide is not None else set()
        for eff in _entr_effects(root):
            if eff['spid'] not in pic_ids:
                continue
            if preset is not None and eff['presetID'] != preset:
                continue
            if filter_kw is not None and not any(filter_kw in (f or '') for f in eff['filters']):
                continue
            return True, 'presetID=%s filter=%s' % (eff['presetID'], ','.join(eff['filters']) or '-')
        return False, '未发现命中该图片的进入动画'

    # +3：全文共23页，"目录"页位于第2页，其余各页保持既定先后顺序
    order_bad = []
    for idx, key in enumerate(SLIDE_TITLES):
        if idx >= len(slides):
            order_bad.append('缺第%d页(%s)' % (idx + 1, key))
            continue
        if _norm(key) not in _norm(_slide_text(slides[idx])):
            order_bad.append('第%d页非"%s"' % (idx + 1, key))
    if len(slides) == 23 and not order_bad:
        results.append("+3：全文共23页，目录页位于第2页，其余各页顺序未打乱、未重复、未遗漏")
    else:
        results.append("+0：全文共23页，目录页位于第2页，其余各页顺序未打乱、未重复、未遗漏（%s）"
                       % ('；'.join(order_bad[:4]) or '共%d页' % len(slides)))

    # +1："目录"二字微软雅黑40磅，六个板块文字微软雅黑30磅
    _, cat_slide, _ = page('CONTENTS')
    cat_bad = []
    if cat_slide is None:
        cat_bad.append('未定位到目录页')
    else:
        shp = _find_shape_by_text(cat_slide, '目录', exact=True)
        if shp is None or not _runs_all(_runs(shp), family=YAHEI, size=40.0):
            cat_bad.append('"目录"二字字号%s' % (_sizes_of(_runs(shp)) if shp is not None else '未找到'))
        for name in ('项目背景与需求', '目标设定与可行性', '方案设计与优选', '实施过程', '效果验证', '标准化与展望'):
            sec = _find_shape_by_text(cat_slide, name, exact=True)
            if sec is None or not _runs_all(_runs(sec), family=YAHEI, size=30.0):
                cat_bad.append('%s(%s)' % (name, _sizes_of(_runs(sec)) if sec is not None else '未找到'))
    if not cat_bad:
        results.append("+1：目录页\"目录\"为微软雅黑40磅、六个板块文字均为微软雅黑30磅")
    else:
        results.append("+0：目录页\"目录\"为微软雅黑40磅、六个板块文字均为微软雅黑30磅（%s）"
                       % '；'.join(cat_bad[:3]))

    # +1 / +1 / +1 / +1：四组只改字号、不改编号的右上角标
    for label, keys, note in (
        ('01  背景', ['项目定位', '风险机制：三个要素相互叠加', '改造前：回收路径长、观察节点少'], '三页'),
        ('02  目标', ['改造目标：把风险链拆开'], ''),
        ('03  需求', ['需求转换：从现场声音到关键需求', '创新借鉴：分割与预先感知'], '两页'),
        ('04  目标', ['目标设定：用可量化指标约束设计', '可行性论证：从工艺、技术、资源三方面评估'], '两页'),
    ):
        ok, why = corner_ok(keys, label)
        rule = '%s右上角标"%s"均为微软雅黑20磅' % (note, label)
        results.append(("+1：%s" % rule) if ok else ("+0：%s（%s）" % (rule, why)))

    # +3 / +3 / +3 / +3 / +1：五组需改写编号并设为20磅的右上角标
    for label, keys, old in (
        ('05  验证', ['模拟试验一：筛分路径验证', '模拟试验二：粉料含量与异常蔓延趋势'], '04  验证'),
        ('06  方案', ['总体方案：物理分路', '研发指标：功能、性能与兼容性'], '05  方案'),
        ('07  实施', ['制定对策：任务、目标与交付物', '对策实施一：筛分系统设计与加工',
                      '对策实施二：可开合观察窗制作', '对策实施三：监测画面接入操作位'], '06  实施'),
        ('08  验证', ['整体试运行：连续班次结果', '效果验证：风险链条被削弱'], '07  验证'),
        ('09  巩固', ['标准化与巩固措施'], '08  巩固'),
    ):
        ok, why = corner_ok(keys, label)
        delta = 1 if label == '09  巩固' else 3
        rule = '%d处右上角标由"%s"改为"%s"且均为微软雅黑20磅' % (len(keys), old, label)
        results.append(("+%d：%s" % (delta, rule)) if ok else ("+0：%s（%s）" % (rule, why)))

    def bottom_note(slide, min_len=8):
        """取页面最下方的一段正文文本框（排除角标、页码等短标记）。"""
        best = None
        for shp in _text_shapes(slide):
            txt = _norm(_shape_text(shp))
            if len(txt) < min_len or re.match(r'^\d{2}\s+\S{2}$', txt):
                continue
            if shp.top is None:
                continue
            if best is None or shp.top > best.top:
                best = shp
        return best

    def longest_text(slide, skip=()):
        """取页面文字最长的一段正文文本框（用于定位被扩充改写的段落）。"""
        best, best_len = None, -1
        for shp in _text_shapes(slide):
            txt = _norm(_shape_text(shp))
            if any(word in txt for word in skip):
                continue
            if len(txt) > best_len:
                best, best_len = shp, len(txt)
        return best

    # +3："风险机制"页按序插入三张提供的PNG配图且尺寸统一
    _, risk_slide, _ = page('风险机制：三个要素相互叠加')
    inserted, risk_why = [], ''
    if risk_slide is not None:
        inserted = [p for p in _pictures(risk_slide) if _pixels(p) == (1448, 1086)]
    shas = {_sha1(p) for p in inserted}
    if len(inserted) < 3:
        risk_why = '页内1448×1086配图仅%d张' % len(inserted)
    elif len(shas) < 3:
        risk_why = '三张配图存在重复（互异%d张）' % len(shas)
    else:
        ws = [_cm(p.width) for p in inserted]
        hs = [_cm(p.height) for p in inserted]
        if max(ws) - min(ws) > 0.10 + 1e-6 or max(hs) - min(hs) > 0.10 + 1e-6:
            risk_why = '三张尺寸不统一（宽差%.2fcm 高差%.2fcm）' % (max(ws) - min(ws), max(hs) - min(hs))
    rule = '"风险机制"页插入三张互不相同的1448×1086配图且三者尺寸统一（宽高差≤0.10cm）'
    results.append(("+3：%s" % rule) if not risk_why else ("+0：%s（%s）" % (rule, risk_why)))

    # +1：新插入的三张配图均设有淡蓝色边框线
    ln_bad = []
    for i, pic in enumerate(inserted, 1):
        ln_ok, ln_why = _line_light_blue(pic, theme)
        if not ln_ok:
            ln_bad.append('第%d张%s' % (i, ln_why))
    rule = '"风险机制"页新插入的三张配图均设有淡蓝色边框线（实线0.75-3磅）'
    if inserted and not ln_bad:
        results.append("+1：%s" % rule)
    else:
        results.append("+0：%s（%s）" % (rule, '缺配图' if not inserted else '；'.join(ln_bad[:3])))

    # +1：三页表格内全部文字均为微软雅黑10磅
    tbl10_bad = []
    for key in ('改造前：回收路径长、观察节点少', '需求转换：从现场声音到关键需求', '对策实施二：可开合观察窗制作'):
        _, sld, _ = page(key)
        runs = _slide_table_runs(sld) if sld is not None else []
        if not runs or not _runs_all(runs, family=YAHEI, size=10.0):
            tbl10_bad.append('%s(%s)' % (key.split('：')[0], _sizes_of(runs) or '无表格'))
    rule = '"改造前""需求转换""对策实施二"三页表格内全部文字均为微软雅黑10磅'
    results.append(("+1：%s" % rule) if not tbl10_bad else ("+0：%s（%s）" % (rule, '；'.join(tbl10_bad))))

    # +5："改造前"页视频设置为进入本页后自动播放
    _, before_slide, before_root = page('改造前：回收路径长、观察节点少')
    media_ids = {str(s.shape_id) for s in _media_shapes(before_slide)} if before_slide is not None else set()
    auto_ids = _autoplay_media_spids(before_root) if before_root is not None else set()
    rule = '"改造前"页视频挂在主序列mainSeq下、进入本页即自动播放'
    if media_ids and (media_ids & auto_ids):
        results.append("+5：%s" % rule)
    else:
        results.append("+0：%s（%s）" % (rule, '页内无视频对象' if not media_ids
                                       else '媒体节点未挂到mainSeq（仅单击播放）'))

    # +1："改造目标"页"分离""观察""维护"三个词均为微软雅黑15磅
    _, goal_slide, _ = page('改造目标：把风险链拆开')
    w15_bad = []
    for word in ('分离', '观察', '维护'):
        shp = _find_shape_by_text(goal_slide, word, exact=True) if goal_slide is not None else None
        if shp is None or not _runs_all(_runs(shp), family=YAHEI, size=15.0):
            w15_bad.append('%s(%s)' % (word, _sizes_of(_runs(shp)) if shp is not None else '未找到'))
    rule = '"改造目标"页"分离""观察""维护"三词均为微软雅黑15磅'
    results.append(("+1：%s" % rule) if not w15_bad else ("+0：%s（%s）" % (rule, '；'.join(w15_bad))))

    # +3："创新借鉴"页三张配图统一为宽8.00cm×高5.00cm
    _, inno_slide, _ = page('创新借鉴：分割与预先感知')
    inno_pics = _pictures(inno_slide) if inno_slide is not None else []
    size_bad = ['%.2f×%.2f' % (_cm(p.width), _cm(p.height)) for p in inno_pics
                if not (_near(_cm(p.width), 8.00, 0.05) and _near(_cm(p.height), 5.00, 0.05))]
    rule = '"创新借鉴"页三张配图统一为宽8.00cm×高5.00cm（±0.05cm）'
    if len(inno_pics) >= 3 and not size_bad:
        results.append("+3：%s" % rule)
    else:
        results.append("+0：%s（%s）" % (rule, '页内配图%d张' % len(inno_pics) if len(inno_pics) < 3
                                       else '实际' + '、'.join(size_bad[:3])))

    # +3："目标设定"页底部居中语句扩充至50字以上并设为微软雅黑20磅
    _, tgt_slide, _ = page('目标设定：用可量化指标约束设计')
    tgt_note = bottom_note(tgt_slide) if tgt_slide is not None else None
    tgt_txt = _norm(_shape_text(tgt_note)) if tgt_note is not None else ''
    tgt_len = _cjk_len(tgt_txt)
    tgt_font = tgt_note is not None and _runs_all(_runs(tgt_note), family=YAHEI, size=20.0)
    rule = '"目标设定"页底部语句扩充至50字以上且为微软雅黑20磅'
    if tgt_len >= 50 and tgt_font:
        results.append("+3：%s" % rule)
    else:
        results.append("+0：%s（实际%d字，字号%s）" % (rule, tgt_len, _sizes_of(_runs(tgt_note)) if tgt_note is not None else '未找到'))

    # +5："可行性论证"页右侧图片添加进入类动态效果
    _, feas_slide, feas_root = page('可行性论证：从工艺、技术、资源三方面评估')
    ok, why = entr_hit(feas_root, feas_slide)
    rule = '"可行性论证"页右侧图片添加了进入类动态效果'
    results.append(("+5：%s" % rule) if ok else ("+0：%s（%s）" % (rule, why)))

    def split_items(txt):
        """把"设计原则"文字拆成并列条目（优先按编号，其次按顿号分条）。"""
        marks = set(re.findall(r'([1-5])[.、）)．]', txt))
        if len(marks) >= 5:
            segs = re.split(r'[1-5][.、）)．]', txt)
        else:
            segs = re.split(r'[、；;\n]', txt)
        segs = [_norm(s).strip('。，； ') for s in segs]
        return [s for s in segs if len(s) >= 3]

    # +3："模拟试验一"页"验证结论"正文扩充至50字以上
    _, pilot_slide, pilot_root = page('模拟试验一：筛分路径验证')
    pilot_body = longest_text(pilot_slide, skip=('PILOT TEST 1',)) if pilot_slide is not None else None
    pilot_txt = _norm(_shape_text(pilot_body)) if pilot_body is not None else ''
    pilot_len = _cjk_len(pilot_txt)
    rule = '"模拟试验一"页"验证结论"正文扩充至50字以上且语义连贯'
    results.append(("+3：%s" % rule) if pilot_len >= 50 else ("+0：%s（实际%d字）" % (rule, pilot_len)))

    # +5："模拟试验一"页图片添加"轮子"进入动画
    ok, why = entr_hit(pilot_root, pilot_slide, preset='21')
    if not ok:
        ok, why = entr_hit(pilot_root, pilot_slide, filter_kw='wheel')
    rule = '"模拟试验一"页图片添加了"轮子"进入动画（presetID=21 / filter=wheel）'
    results.append(("+5：%s" % rule) if ok else ("+0：%s（%s）" % (rule, why)))

    # +1："模拟试验二"页免责说明语句为微软雅黑20磅
    _, pilot2_slide, _ = page('模拟试验二：粉料含量与异常蔓延趋势')
    note2 = _find_shape_by_text(pilot2_slide, '本试验仅用于工艺趋势判断') if pilot2_slide is not None else None
    rule = '"模拟试验二"页"本试验仅用于工艺趋势判断…"为微软雅黑20磅'
    if note2 is not None and _runs_all(_runs(note2), family=YAHEI, size=20.0):
        results.append("+1：%s" % rule)
    else:
        results.append("+0：%s（%s）" % (rule, '未找到该语句' if note2 is None else '字号%s' % _sizes_of(_runs(note2))))

    # +3："总体方案"页"设计原则"扩充为五条并列条目
    _, plan_slide, _ = page('总体方案：物理分路')
    principle = _find_shape_by_text(plan_slide, '不干扰主线控制') if plan_slide is not None else None
    if principle is None and plan_slide is not None:
        principle = longest_text(plan_slide, skip=('设计原则',))
    items = split_items(_norm(_shape_text(principle))) if principle is not None else []
    rule = '"总体方案"页"设计原则"扩充为五条互不重复的并列条目'
    if len(items) >= 5 and len(set(items)) >= 5:
        results.append("+3：%s" % rule)
    else:
        results.append("+0：%s（实际%d条，去重后%d条）" % (rule, len(items), len(set(items))))

    # +1："研发指标"页表格内全部文字均为微软雅黑20磅
    _, spec_slide, _ = page('研发指标：功能、性能与兼容性')
    spec_runs = _slide_table_runs(spec_slide) if spec_slide is not None else []
    rule = '"研发指标"页表格内全部文字均为微软雅黑20磅'
    if spec_runs and _runs_all(spec_runs, family=YAHEI, size=20.0):
        results.append("+1：%s" % rule)
    else:
        results.append("+0：%s（%s）" % (rule, _sizes_of(spec_runs) or '未找到表格'))

    # +1："研发指标"页底部说明语句为微软雅黑20磅
    spec_note = _find_shape_by_text(spec_slide, '指标已经转为公开交流口径') if spec_slide is not None else None
    rule = '"研发指标"页以"指标已经转为公开交流口径"开头的说明语句为微软雅黑20磅'
    if spec_note is not None and _runs_all(_runs(spec_note), family=YAHEI, size=20.0):
        results.append("+1：%s" % rule)
    else:
        results.append("+0：%s（%s）" % (rule, '未找到该语句' if spec_note is None else '字号%s' % _sizes_of(_runs(spec_note))))

    # +1："制定对策"页表格"责任角色"列末格改为"项目协调负责人：沈亦然"
    _, plan2_slide, _ = page('制定对策：任务、目标与交付物')
    plan2_tbl = _table_shape(plan2_slide) if plan2_slide is not None else None
    role_txt = ''
    if plan2_tbl is not None:
        _t = plan2_tbl.table
        role_txt = _norm(_t.cell(len(_t.rows) - 1, len(_t.columns) - 1).text)
    rule = '"制定对策"页表格"责任角色"列末格文字改为"项目协调负责人：沈亦然"'
    if '项目协调负责人：沈亦然' in role_txt:
        results.append("+1：%s" % rule)
    else:
        results.append("+0：%s（实际：%s）" % (rule, role_txt or '未找到表格'))

    # +1："对策实施一"页三句说明均为微软雅黑15磅
    _, act1_slide, act1_root = page('对策实施一：筛分系统设计与加工')
    s15_bad = []
    for sent in ('长孔筛面兼顾分离效率与堵塞风险', '让物料在自重与振动下自然前移', '减少粉料外逸和箱体边缘堆积'):
        shp = _find_shape_by_text(act1_slide, sent) if act1_slide is not None else None
        if shp is None or not _runs_all(_runs(shp), family=YAHEI, size=15.0):
            s15_bad.append('%s(%s)' % (sent[:6], _sizes_of(_runs(shp)) if shp is not None else '未找到'))
    rule = '"对策实施一"页三句说明文字均为微软雅黑15磅'
    results.append(("+1：%s" % rule) if not s15_bad else ("+0：%s（%s）" % (rule, '；'.join(s15_bad))))

    # +5："对策实施二"页图片添加"渐变（淡入）"进入动画
    _, act2_slide, act2_root = page('对策实施二：可开合观察窗制作')
    ok, why = entr_hit(act2_root, act2_slide, preset='10')
    if not ok:
        ok, why = entr_hit(act2_root, act2_slide, filter_kw='fade')
    rule = '"对策实施二"页图片添加了"渐变（淡入）"进入动画（presetID=10 / filter=fade）'
    results.append(("+5：%s" % rule) if ok else ("+0：%s（%s）" % (rule, why)))

    # +5："对策实施三"页图片添加"溶解"进入动画
    _, act3_slide, act3_root = page('对策实施三：监测画面接入操作位')
    ok, why = entr_hit(act3_root, act3_slide, preset='9')
    if not ok:
        ok, why = entr_hit(act3_root, act3_slide, filter_kw='dissolve')
    rule = '"对策实施三"页图片添加了"溶解"进入动画（presetID=9 / filter=dissolve）'
    results.append(("+5：%s" % rule) if ok else ("+0：%s（%s）" % (rule, why)))

    def circle_above(slide, label):
        """定位某个文字标签正上方的圆形。"""
        lab = _find_shape_by_text(slide, label, exact=True)
        if lab is None or lab.left is None:
            return None
        cx = lab.left + lab.width / 2.0
        best, best_dx = None, None
        for shp in _iter_shapes(slide):
            if shp.left is None or shp.width is None or shp.height is None:
                continue
            w, h = _cm(shp.width), _cm(shp.height)
            if not (1.0 <= w <= 3.5) or abs(w - h) > 0.15:
                continue
            if shp.top is None or shp.top + shp.height > lab.top + lab.height:
                continue
            dx = abs(shp.left + shp.width / 2.0 - cx)
            if best_dx is None or dx < best_dx:
                best, best_dx = shp, dx
        if best is not None and best_dx is not None and best_dx > 2 * lab.width:
            return None
        return best

    # +5："整体试运行"页表格改为6行4列且内容与所给Excel一致
    _, trial_slide, _ = page('整体试运行：连续班次结果')
    trial_tbl = _table_shape(trial_slide) if trial_slide is not None else None
    trial_bad = []
    if trial_tbl is None:
        trial_bad.append('未找到表格')
    else:
        tbl = trial_tbl.table
        if len(tbl.rows) != 6 or len(tbl.columns) != 4:
            trial_bad.append('实际%d行%d列' % (len(tbl.rows), len(tbl.columns)))
        else:
            for r in range(6):
                for c in range(4):
                    want = TRIAL_TABLE[r][c]
                    if want is None:
                        continue
                    got = _norm(tbl.cell(r, c).text)
                    if got != _norm(want):
                        trial_bad.append('第%d行%d列(%s)' % (r + 1, c + 1, got or '空'))
    rule = '"整体试运行"页表格改为6行4列且22处有值单元格内容与所给Excel一致'
    if not trial_bad:
        results.append("+5：%s" % rule)
    else:
        results.append("+0：%s（%s）" % (rule, '；'.join(trial_bad[:4])))

    # +1："整体试运行"页表格尺寸为宽24.07cm×高10.59cm
    rule = '"整体试运行"页表格尺寸为宽24.07cm×高10.59cm（±0.10cm）'
    if trial_tbl is not None and _near(_cm(trial_tbl.width), 24.07, 0.10) and _near(_cm(trial_tbl.height), 10.59, 0.10):
        results.append("+1：%s" % rule)
    else:
        results.append("+0：%s（实际%s）" % (rule, '未找到表格' if trial_tbl is None
                                          else '%.2f×%.2fcm' % (_cm(trial_tbl.width), _cm(trial_tbl.height))))

    # +1："效果验证"页"燃料侧"圆形为红色、"发现侧"圆形为黄色
    _, eff_slide, _ = page('效果验证：风险链条被削弱')
    fuel = circle_above(eff_slide, '燃料侧') if eff_slide is not None else None
    found = circle_above(eff_slide, '发现侧') if eff_slide is not None else None
    fuel_hex = _fill_rgb(fuel, theme) if fuel is not None else None
    found_hex = _fill_rgb(found, theme) if found is not None else None
    rule = '"效果验证"页"燃料侧"上方圆形填充为红色、"发现侧"上方圆形填充为黄色'
    if _is_red(fuel_hex) and _is_yellow(found_hex):
        results.append("+1：%s" % rule)
    else:
        results.append("+0：%s（燃料侧%s、发现侧%s）" % (rule, fuel_hex or '未取到', found_hex or '未取到'))

    # +1："标准化与巩固措施"页"下一步…"与"总结与展望"页"后续方向…"均为微软雅黑20磅
    _, std_slide, _ = page('标准化与巩固措施')
    _, sum_slide, _ = page('总结与展望')
    next_shp = _find_shape_by_text(std_slide, '下一步：') if std_slide is not None else None
    dir_shp = _find_shape_by_text(sum_slide, '后续方向：') if sum_slide is not None else None
    tail20_bad = []
    if next_shp is None or not _runs_all(_runs(next_shp), family=YAHEI, size=20.0):
        tail20_bad.append('下一步(%s)' % (_sizes_of(_runs(next_shp)) if next_shp is not None else '未找到'))
    if dir_shp is None or not _runs_all(_runs(dir_shp), family=YAHEI, size=20.0):
        tail20_bad.append('后续方向(%s)' % (_sizes_of(_runs(dir_shp)) if dir_shp is not None else '未找到'))
    rule = '"标准化与巩固措施"页"下一步…"与"总结与展望"页"后续方向…"均为微软雅黑20磅'
    results.append(("+1：%s" % rule) if not tail20_bad else ("+0：%s（%s）" % (rule, '；'.join(tail20_bad))))

    # +1："总结与展望""感谢聆听"两处四字标题均为微软雅黑40磅
    _, thx_slide, _ = page('感谢聆听')
    t40_bad = []
    for key, sld in (('总结与展望', sum_slide), ('感谢聆听', thx_slide)):
        shp = _find_shape_by_text(sld, key, exact=True) if sld is not None else None
        if shp is None or not _runs_all(_runs(shp), family=YAHEI, size=40.0):
            t40_bad.append('%s(%s)' % (key, _sizes_of(_runs(shp)) if shp is not None else '未找到'))
    rule = '"总结与展望""感谢聆听"两处四字标题均为微软雅黑40磅'
    results.append(("+1：%s" % rule) if not t40_bad else ("+0：%s（%s）" % (rule, '；'.join(t40_bad))))

    # -3："风险机制"页未保留左侧宽14.35cm×高13.54cm的装置示意图片
    keep = None
    for pic in (_pictures(risk_slide) if risk_slide is not None else []):
        if (_near(_cm(pic.width), 14.35, 0.10) and _near(_cm(pic.height), 13.54, 0.10)
                and _near(_cm(pic.left), 1.40, 0.20) and _near(_cm(pic.top), 3.35, 0.20)):
            keep = pic
            break
    rule = '"风险机制"页不满足在页面左侧（1.40cm, 3.35cm）保留一张宽14.35cm×高13.54cm的装置示意图片'
    results.append(("-0：%s" % rule) if keep is not None else ("-3：%s" % rule))

    # -3："整体试运行"页未保留表格右侧的现场监测画面图片
    trial_pic = None
    for pic in (_pictures(trial_slide) if trial_slide is not None else []):
        if 8.80 - 1e-6 <= _cm(pic.width) <= 10.70 + 1e-6 and 9.00 - 1e-6 <= _cm(pic.height) <= 10.70 + 1e-6:
            trial_pic = pic
            break
    rule = '"整体试运行"页不满足在表格右侧保留一张宽8.80-10.70cm、高9.00-10.70cm的现场监测画面图片'
    results.append(("-0：%s" % rule) if trial_pic is not None else ("-3：%s" % rule))

    # -3："整体试运行"页"结果显示："右侧语句被改写或格式被破坏
    WANT = '平均分离率约89.8%，带出情况稳定；监测画面连续覆盖关键区域，装置运行达到设计预期。'
    res_shp = None
    for shp in (_text_shapes(trial_slide) if trial_slide is not None else []):
        if '平均分离率约89.8%' in _norm(_shape_text(shp)):
            res_shp = shp
            break
    res_ok = (res_shp is not None and _norm(WANT) in _norm(_shape_text(res_shp))
              and _runs_all([r for r in _runs(res_shp) if _norm(r.text) and '结果显示' not in r.text],
                            family=SONGTI, size=12.8, bold=True, rgb='111827'))
    rule = '"整体试运行"页"结果显示："右侧不满足为"平均分离率约89.8%…"原句且宋体12.8磅 加粗 颜色111827'
    results.append(("-0：%s" % rule) if res_ok else ("-3：%s" % rule))

    # -1："整体试运行"页"方案C"列两格未保持空白（自行编造Excel中不存在的内容）
    blank_bad = []
    if (
        trial_tbl is not None
        and len(trial_tbl.table.rows) >= 4
        and all(len(row.cells) >= 4 for row in trial_tbl.table.rows)
    ):
        for r, name in ((1, '单线产能'), (3, '维护空间')):
            got = _norm(trial_tbl.table.cell(r, 3).text)
            if got:
                blank_bad.append('%s填了"%s"' % (name, got))
    rule = '"整体试运行"页表格"方案C：智能柔性产线"列的"单线产能""维护空间"两格不满足保持空白'
    results.append(("-0：%s" % rule) if not blank_bad else ("-1：%s（%s）" % (rule, '；'.join(blank_bad))))

    # -3：11处改写后的右上角标未保持微软雅黑 加粗 颜色00A6A6
    corner_fmt_bad = []
    for key in ('模拟试验一：筛分路径验证', '模拟试验二：粉料含量与异常蔓延趋势', '总体方案：物理分路',
                '研发指标：功能、性能与兼容性', '制定对策：任务、目标与交付物', '对策实施一：筛分系统设计与加工',
                '对策实施二：可开合观察窗制作', '对策实施三：监测画面接入操作位', '整体试运行：连续班次结果',
                '效果验证：风险链条被削弱', '标准化与巩固措施'):
        _, sld, _ = page(key)
        shp = _corner_shape(sld) if sld is not None else None
        if shp is None or not _runs_all(_runs(shp), family=YAHEI, bold=True, rgb='00A6A6'):
            corner_fmt_bad.append(key.split('：')[0])
    rule = '"05  验证""06  方案""07  实施""08  验证""09  巩固"共11处右上角标不满足微软雅黑 加粗 颜色00A6A6'
    if not corner_fmt_bad:
        results.append("-0：%s" % rule)
    else:
        results.append("-3：%s（%d处不符：%s）" % (rule, len(corner_fmt_bad), '、'.join(corner_fmt_bad[:4])))

    # -1："模拟试验一"页新增了叠放在图片上的多余小圆形
    extra_shapes = []
    for shp in (_iter_shapes(pilot_slide) if pilot_slide is not None else []):
        if shp.width is None or shp.height is None or shp.element.tag == P + 'pic':
            continue
        if getattr(shp, 'has_table', False) and shp.has_table:
            continue
        w, h = _cm(shp.width), _cm(shp.height)
        if 0.30 <= w <= 1.10 and 0.30 <= h <= 1.10 and not _norm(_shape_text(shp)):
            extra_shapes.append('%.2f×%.2fcm' % (w, h))
    rule = '"模拟试验一"页不满足页内不含其他图形（新增了叠放在图片上的0.70cm×0.70cm圆形）'
    if not extra_shapes:
        results.append("-0：%s" % rule)
    else:
        results.append("-1：%s（多出%d个：%s）" % (rule, len(extra_shapes), '、'.join(extra_shapes[:3])))

    # -1："模拟试验一"页"验证结论"正文字体格式被破坏
    body_fmt = pilot_body is not None and _runs_all(_runs(pilot_body), family=YAHEI, size=13.0, rgb='334155')
    rule = '"模拟试验一"页"验证结论"正文不满足微软雅黑13.0磅 颜色334155'
    results.append(("-0：%s" % rule) if body_fmt else ("-1：%s" % rule))

    # -1："模拟试验一"页"验证结论"文本框被压缩
    box_ok = (pilot_body is not None and pilot_body.width is not None
              and _cm(pilot_body.width) >= 9.00 - 1e-6 and _cm(pilot_body.height) >= 2.00 - 1e-6)
    rule = '"模拟试验一"页"验证结论"正文文本框不满足宽度≥9.00cm、高度≥2.00cm'
    if box_ok:
        results.append("-0：%s" % rule)
    else:
        results.append("-1：%s（实际%s）" % (rule, '未找到' if pilot_body is None
                                        else '%.2f×%.2fcm' % (_cm(pilot_body.width), _cm(pilot_body.height))))

    # -1："目标设定"页底部语句字体格式被破坏
    tgt_fmt = tgt_note is not None and _runs_all(_runs(tgt_note), family=YAHEI, bold=True, rgb='00A6A6')
    rule = '"目标设定"页底部居中语句不满足微软雅黑 加粗 颜色00A6A6'
    results.append(("-0：%s" % rule) if tgt_fmt else ("-1：%s" % rule))

    # -1："总体方案"页"设计原则"文字格式被破坏
    pri_fmt = principle is not None and _runs_all(_runs(principle), family=YAHEI, size=12.5, bold=True, rgb='0A1F33')
    rule = '"总体方案"页"设计原则"文字不满足微软雅黑12.5磅 加粗 颜色0A1F33'
    results.append(("-0：%s" % rule) if pri_fmt else ("-1：%s" % rule))

    # -1："制定对策"页表格内文字字号被改动
    plan2_runs = _slide_table_runs(plan2_slide) if plan2_slide is not None else []
    p2_ok = bool(plan2_runs) and _runs_all(plan2_runs, family=YAHEI, size=8.7)
    rule = '"制定对策"页表格内全部文字（含"项目协调负责人：沈亦然"单元格）不满足微软雅黑8.7磅'
    results.append(("-0：%s" % rule) if p2_ok else ("-1：%s（实际%s）" % (rule, _sizes_of(plan2_runs) or '未找到表格')))

    # -1："效果验证"页"处置侧"上方圆形被改动
    disp = circle_above(eff_slide, '处置侧') if eff_slide is not None else None
    disp_hex = _fill_rgb(disp, theme) if disp is not None else None
    disp_ok = (disp is not None and _near(_cm(disp.width), 2.08, 0.10) and _near(_cm(disp.height), 2.08, 0.10)
               and (disp_hex or '').upper() == '10B981')
    rule = '"效果验证"页"处置侧"上方圆形不满足为直径2.08cm、填充色10B981的正圆'
    if disp_ok:
        results.append("-0：%s" % rule)
    else:
        results.append("-1：%s（实际%s）" % (rule, '未找到' if disp is None
                                        else '%.2f×%.2fcm %s' % (_cm(disp.width), _cm(disp.height), disp_hex or '无填充')))

    # -1："改造前"页视频对象或其媒体链接被破坏
    vid_ok = False
    for shp in (_media_shapes(before_slide) if before_slide is not None else []):
        vf = shp.element.find('.//' + A + 'videoFile')
        linked = vf is not None and any(k.endswith('}link') or k.endswith('}embed') for k in vf.attrib)
        if linked and _near(_cm(shp.width), 11.06, 0.10) and _near(_cm(shp.height), 14.59, 0.10):
            vid_ok = True
            break
    rule = '"改造前"页不满足保留一个宽11.06cm×高14.59cm的视频对象及其可播放的媒体链接'
    results.append(("-0：%s" % rule) if vid_ok else ("-1：%s" % rule))

    dim2_items = result["dim2_items"]
    if len(results) != len(_MAX_DELTAS_ORDER):
        raise RuntimeError("评分项数量与 _MAX_DELTAS_ORDER 不一致：%d vs %d"
                           % (len(results), len(_MAX_DELTAS_ORDER)))
    total = 0
    for raw, max_delta in zip(results, _MAX_DELTAS_ORDER):
        head, _, rule_text = raw.partition("：")
        delta = int(head.replace("+", "").strip())
        hit = delta != 0
        real = max_delta if hit else 0
        total += real
        dim2_items.append({"rule": rule_text, "max_delta": max_delta,
                           "delta": real, "hit": hit})
    result["total_score"] = max(0, total)
    return result


if __name__ == "__main__":
    _dir = sys.argv[1] if len(sys.argv) >= 2 else os.path.dirname(os.path.abspath(__file__))
    print(json.dumps(evaluate(_dir), ensure_ascii=False, indent=2))
