# -*- coding: utf-8 -*-
"""
对 "日常信息处理能力观察表_图形组织可编辑版.xlsx" 的自动评估脚本

评估逻辑：
  维度1（可用与可修改性）：门槛维度。任意一条不满足 -> 直接判 0 分，不再检查维度2。
  维度2（完成度评分细则）：得分点 + 扣分点。
      - 加分细则：必须满足细则内的【每一个点】才加分。
      - 扣分细则：只要满足细则内的【任意一点】即扣分。
      - 累计各条命中的分值（可正可负），即为最终得分。

实现说明：
  Excel 形状/图形的几何、尺寸、轮廓信息以 EMU(English Metric Unit) 存储于
  xl/drawings/drawing*.xml 中。本脚本直接解析 OOXML（解压 xlsx）来读取这些底层信息，
  对于无法纯程序判定的主观项（如"是否近似镜像""菱形交叠区域是否清晰"），
  采用基于几何坐标的可量化近似规则来体现评估意图。

单位换算：
  1 cm = 360000 EMU
  1 pt(磅) = 12700 EMU
"""

import os
import sys
import json
import zipfile
import shutil
import tempfile
import xml.etree.ElementTree as ET

# ----------------------------------------------------------------------------
# 单位与常量
# ----------------------------------------------------------------------------
EMU_PER_CM = 360000.0
EMU_PER_PT = 12700.0

NS = {
    'xdr': 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing',
    'a':   'http://schemas.openxmlformats.org/drawingml/2006/main',
    'x':   'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
    'r':   'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    'rel': 'http://schemas.openxmlformats.org/package/2006/relationships',
    'ct':  'http://schemas.openxmlformats.org/package/2006/content-types',
}


def emu_to_cm(v):
    return v / EMU_PER_CM


def emu_to_pt(v):
    return v / EMU_PER_PT


# ----------------------------------------------------------------------------
# 解析辅助
# ----------------------------------------------------------------------------
class WorkbookData(object):
    """承载从 xlsx 解压后解析出的全部信息。"""

    def __init__(self):
        self.ok_open = False              # 文件能否作为 zip/ooxml 正常打开
        self.ext = None                   # 文件扩展名
        self.sheet_names = []             # 工作表名列表
        self.sheet_files = {}             # 表名 -> worksheet xml 路径
        self.shared_strings = []          # 共享字符串（已拼接 rich text）
        self.shapes = []                  # C32 所在绘图中的形状列表（dict）
        self.cell_text = {}               # 'Sheet0' 单元格 -> 文本，形如 cell_text[('A1')]
        self.merged = []                  # Sheet0 合并单元格
        self.cols = []                    # Sheet0 列宽信息
        self.rows = {}                    # Sheet0 行高信息 r -> ht
        self.cell_style = {}              # Sheet0 单元格 -> style index
        self.xf_font = {}                 # cellXfs 索引 -> fontId
        self.xf_fill = {}                 # cellXfs 索引 -> fillId
        self.xf_wrap = {}                 # cellXfs 索引 -> 是否自动换行
        self.xf_shrink = {}               # cellXfs 索引 -> 是否缩小字体填充
        self.font_color = {}              # fontId -> 字体颜色 RGB(8位含alpha或6位)
        self.font_size = {}               # fontId -> 字号
        self.fill_fg = {}                 # fillId -> 前景填充色 RGB / None
        self.parse_error = None


def _strip_text_from_si(si):
    """从 sharedStrings 的一个 <si> 节点提取纯文本（处理 rich text run）。"""
    texts = []
    for t in si.iter('{%s}t' % NS['x']):
        texts.append(t.text or '')
    return ''.join(texts)


def load_workbook(path):
    data = WorkbookData()
    data.ext = os.path.splitext(path)[1].lower()

    if not os.path.exists(path):
        data.parse_error = '文件不存在'
        return data

    # 这里只支持 .xlsx/.xlsm 的 OOXML zip 容器深度解析。
    if not zipfile.is_zipfile(path):
        data.parse_error = '文件不是合法的 .xlsx/.xlsm OOXML zip 容器或已损坏'
        return data

    tmp = tempfile.mkdtemp(prefix='xlsx_eval_')
    try:
        with zipfile.ZipFile(path) as z:
            z.extractall(tmp)
        data.ok_open = True
        _parse_extracted(tmp, data)
    except Exception as e:  # noqa
        data.parse_error = '解析失败: %r' % e
    finally:
        shutil.rmtree(tmp, ignore_errors=True)
    return data


def _parse_extracted(root, data):
    # 1) workbook.xml -> 工作表名 + rId
    wb_path = os.path.join(root, 'xl', 'workbook.xml')
    rels_path = os.path.join(root, 'xl', '_rels', 'workbook.xml.rels')
    rid_to_target = {}
    if os.path.exists(rels_path):
        rtree = ET.parse(rels_path).getroot()
        for rel in rtree.findall('rel:Relationship', NS):
            rid_to_target[rel.get('Id')] = rel.get('Target')

    sheet_rid = {}
    if os.path.exists(wb_path):
        wtree = ET.parse(wb_path).getroot()
        for sh in wtree.iter('{%s}sheet' % NS['x']):
            name = sh.get('name')
            rid = sh.get('{%s}id' % NS['r'])
            data.sheet_names.append(name)
            sheet_rid[name] = rid

    # 2) shared strings
    ss_path = os.path.join(root, 'xl', 'sharedStrings.xml')
    if os.path.exists(ss_path):
        stree = ET.parse(ss_path).getroot()
        for si in stree.findall('x:si', NS):
            data.shared_strings.append(_strip_text_from_si(si))

    # 3) 找到 Sheet0 的 worksheet xml
    sheet0_target = None
    for name in ('Sheet0',):
        rid = sheet_rid.get(name)
        if rid and rid in rid_to_target:
            tgt = rid_to_target[rid]
            sheet0_target = tgt.lstrip('/')
            if not sheet0_target.startswith('xl/'):
                sheet0_target = 'xl/' + sheet0_target
    # 兜底：若关系解析不到，默认 sheet1.xml
    if sheet0_target is None:
        sheet0_target = 'xl/worksheets/sheet1.xml'

    ws_path = os.path.join(root, *sheet0_target.split('/'))
    if os.path.exists(ws_path):
        _parse_worksheet(ws_path, root, data)


def _col_letter_to_idx(col):
    idx = 0
    for ch in col:
        idx = idx * 26 + (ord(ch) - ord('A') + 1)
    return idx - 1


def _split_ref(ref):
    """'C32' -> ('C', 32)"""
    col = ''.join(ch for ch in ref if ch.isalpha())
    row = int(''.join(ch for ch in ref if ch.isdigit()))
    return col, row


# Excel 传统 64 色索引调色板（index -> RRGGBB），0-7/8-15 为基础色的两份拷贝（历史遗留）。
INDEXED_PALETTE = [
    '000000', 'FFFFFF', 'FF0000', '00FF00', '0000FF', 'FFFF00', 'FF00FF', '00FFFF',
    '000000', 'FFFFFF', 'FF0000', '00FF00', '0000FF', 'FFFF00', 'FF00FF', '00FFFF',
    '800000', '008000', '000080', '808000', '800080', '008080', 'C0C0C0', '808080',
    '9999FF', '993366', 'FFFFCC', 'CCFFFF', '660066', 'FF8080', '0066CC', 'CCCCFF',
    '000080', 'FF00FF', 'FFFF00', '00FFFF', '800080', '800000', '008080', '0000FF',
    '00CCFF', 'CCFFFF', 'CCFFCC', 'FFFF99', '99CCFF', 'FF99CC', 'CC99FF', 'FFCC99',
    '3366FF', '33CCCC', '99CC00', 'FFCC00', 'FF9900', 'FF6600', '666699', '969696',
    '003366', '339966', '003300', '333300', '993300', '993366', '333399', '333333',
]

# SpreadsheetML theme 索引 -> clrScheme 槽位名，顺序固定（ECMA-376/Excel 约定）。
THEME_SLOT_ORDER = ('lt1', 'dk1', 'lt2', 'dk2', 'accent1', 'accent2', 'accent3',
                     'accent4', 'accent5', 'accent6', 'hlink', 'folHlink')

# 主题文件缺失/槽位缺失时的 Office 默认主题色兜底。
DEFAULT_THEME_RGB = {
    'lt1': 'FFFFFF', 'dk1': '000000', 'lt2': 'EEECE1', 'dk2': '1F497D',
    'accent1': '4472C4', 'accent2': 'ED7D31', 'accent3': 'A5A5A5',
    'accent4': 'FFC000', 'accent5': '5B9BD5', 'accent6': '70AD47',
    'hlink': '0563C1', 'folHlink': '954F72',
}


def _normalize_rgb_hex(value):
    """将 rgb/ARGB 十六进制值归一化为大写 RRGGBB；无法解析返回 None。"""
    if not value:
        return None
    v = value.strip().upper()
    if len(v) == 8:      # ARGB -> 去掉 alpha
        v = v[2:]
    if len(v) != 6:
        return None
    try:
        int(v, 16)
    except ValueError:
        return None
    return v


def _apply_tint(rgb_hex, tint):
    """按 Excel 惯用的按通道近似公式对 rgb_hex 施加 tint（与本仓库其他 verifier 一致）。
    tint<0：向黑变暗；tint>0：向白变亮；tint 为 None/0 时原样返回。
    """
    if not rgb_hex or not tint:
        return rgb_hex
    try:
        t = float(tint)
    except (TypeError, ValueError):
        return rgb_hex
    if t == 0:
        return rgb_hex
    rr = int(rgb_hex[0:2], 16); gg = int(rgb_hex[2:4], 16); bb = int(rgb_hex[4:6], 16)

    def _ch(c):
        if t < 0:
            v = c * (1 + t)
        else:
            v = c * (1 - t) + 255 * t
        return max(0, min(255, int(round(v))))

    return '%02X%02X%02X' % (_ch(rr), _ch(gg), _ch(bb))


def _load_theme_colors(root):
    """解析 xl/theme/theme1.xml 的 <a:clrScheme>，返回按 THEME_SLOT_ORDER 顺序的
    12 个 RRGGBB 颜色列表；缺失主题文件或个别槽位时用 Office 默认主题色兜底。"""
    colors = dict(DEFAULT_THEME_RGB)
    tpath = os.path.join(root, 'xl', 'theme', 'theme1.xml')
    if os.path.exists(tpath):
        try:
            ttree = ET.parse(tpath).getroot()
        except ET.ParseError:
            ttree = None
        if ttree is not None:
            scheme = ttree.find('.//a:clrScheme', NS)
            if scheme is not None:
                for node in list(scheme):
                    key = node.tag.split('}', 1)[-1]
                    if key not in THEME_SLOT_ORDER:
                        continue
                    srgb = node.find('a:srgbClr', NS)
                    sysclr = node.find('a:sysClr', NS)
                    raw = None
                    if srgb is not None and srgb.get('val'):
                        raw = srgb.get('val')
                    elif sysclr is not None and sysclr.get('lastClr'):
                        raw = sysclr.get('lastClr')
                    norm = _normalize_rgb_hex(raw) if raw else None
                    if norm:
                        colors[key] = norm
    return [colors[name] for name in THEME_SLOT_ORDER]


def _resolve_color_element(color_el, theme_colors):
    """解析 SpreadsheetML <x:color>/<x:fgColor> 元素为最终 RRGGBB。
    优先级：rgb > theme(+tint) > indexed；均无法解析时返回 None（不臆造证据）。
    """
    if color_el is None:
        return None
    rgb = _normalize_rgb_hex(color_el.get('rgb'))
    if rgb:
        return rgb
    theme_attr = color_el.get('theme')
    if theme_attr is not None:
        try:
            idx = int(theme_attr)
        except ValueError:
            idx = None
        if idx is not None and 0 <= idx < len(theme_colors):
            base = theme_colors[idx]
            tint = color_el.get('tint')
            return _apply_tint(base, tint) if tint else base
    indexed_attr = color_el.get('indexed')
    if indexed_attr is not None:
        try:
            idx = int(indexed_attr)
        except ValueError:
            idx = None
        # 64/65 为 Excel 遗留的“系统前景/背景（自动）”特殊索引，不代表具体颜色
        if idx is not None and 0 <= idx < len(INDEXED_PALETTE):
            return INDEXED_PALETTE[idx]
    return None


def _parse_styles(root, data):
    """解析 styles.xml：cellXfs -> fontId/fillId；fonts -> 颜色/字号；fills -> 背景色。
    颜色解析同时支持 rgb/theme(+tint)/indexed 三种取值方式（theme 需配合 xl/theme/theme1.xml）。
    """
    spath = os.path.join(root, 'xl', 'styles.xml')
    if not os.path.exists(spath):
        return
    stree = ET.parse(spath).getroot()
    theme_colors = _load_theme_colors(root)

    # fonts
    fonts = stree.find('x:fonts', NS)
    if fonts is not None:
        for i, f in enumerate(fonts.findall('x:font', NS)):
            clr = f.find('x:color', NS)
            rgb = _resolve_color_element(clr, theme_colors)
            sz = f.find('x:sz', NS)
            data.font_color[i] = rgb
            data.font_size[i] = float(sz.get('val')) if (sz is not None and sz.get('val')) else None

    # fills
    fills = stree.find('x:fills', NS)
    if fills is not None:
        for i, fl in enumerate(fills.findall('x:fill', NS)):
            pf = fl.find('x:patternFill', NS)
            fg = None
            if pf is not None:
                fgc = pf.find('x:fgColor', NS)
                fg = _resolve_color_element(fgc, theme_colors)
            data.fill_fg[i] = fg

    # cellXfs
    cellxfs = stree.find('x:cellXfs', NS)
    if cellxfs is not None:
        for i, xf in enumerate(cellxfs.findall('x:xf', NS)):
            data.xf_font[i] = int(xf.get('fontId')) if xf.get('fontId') else 0
            data.xf_fill[i] = int(xf.get('fillId')) if xf.get('fillId') else 0
            # 对齐：是否自动换行(wrapText)、是否缩小字体填充(shrinkToFit)
            al = xf.find('x:alignment', NS)
            wrap = (al is not None and al.get('wrapText') in ('1', 'true'))
            shrink = (al is not None and al.get('shrinkToFit') in ('1', 'true'))
            data.xf_wrap[i] = wrap
            data.xf_shrink[i] = shrink


def _parse_worksheet(ws_path, root, data):
    wtree = ET.parse(ws_path).getroot()

    # 样式表：建立 cellXfs 索引 -> (字体色, 填充背景色, 字号)
    _parse_styles(root, data)

    # 列宽
    for col in wtree.iter('{%s}col' % NS['x']):
        data.cols.append({
            'min': int(col.get('min')),
            'max': int(col.get('max')),
            'width': float(col.get('width')) if col.get('width') else None,
        })

    # 单元格文本 + 行高 + 样式
    for row in wtree.iter('{%s}row' % NS['x']):
        r = row.get('r')
        ht = row.get('ht')
        if r:
            data.rows[int(r)] = float(ht) if ht else None
        for c in row.findall('x:c', NS):
            ref = c.get('r')
            if not ref:
                continue
            s = c.get('s')
            if s is not None:
                data.cell_style[ref] = s
            t = c.get('t')
            v = c.find('x:v', NS)
            text = ''
            if t == 's' and v is not None:
                try:
                    text = data.shared_strings[int(v.text)]
                except (ValueError, IndexError):
                    text = ''
            elif t == 'str' and v is not None:
                text = v.text or ''
            elif v is not None:
                text = v.text or ''
            else:
                inline = c.find('x:is', NS)
                if inline is not None:
                    text = _strip_text_from_si(inline)
            data.cell_text[ref] = text

    # 合并单元格
    for mc in wtree.iter('{%s}mergeCell' % NS['x']):
        data.merged.append(mc.get('ref'))

    # drawing 关系 -> drawing xml
    ws_rels = os.path.join(os.path.dirname(ws_path), '_rels',
                           os.path.basename(ws_path) + '.rels')
    drawing_target = None
    if os.path.exists(ws_rels):
        rtree = ET.parse(ws_rels).getroot()
        for rel in rtree.findall('rel:Relationship', NS):
            if rel.get('Type', '').endswith('/drawing'):
                drawing_target = rel.get('Target').lstrip('/')
                if not drawing_target.startswith('xl/'):
                    drawing_target = 'xl/' + drawing_target.replace('../', '')
    if drawing_target:
        dpath = os.path.join(root, *drawing_target.split('/'))
        if os.path.exists(dpath):
            _parse_drawing(dpath, data)


def _parse_anchor_from(anch):
    """提取 anchor 的 xdr:from -> (col, row, colOff, rowOff)，缺失时返回 (None, None, 0, 0)。"""
    frm = anch.find('xdr:from', NS)
    if frm is None:
        return None, None, 0, 0
    col = int(frm.find('xdr:col', NS).text)
    row = int(frm.find('xdr:row', NS).text)
    colOff_n = frm.find('xdr:colOff', NS)
    rowOff_n = frm.find('xdr:rowOff', NS)
    colOff = int(colOff_n.text) if colOff_n is not None else 0
    rowOff = int(rowOff_n.text) if rowOff_n is not None else 0
    return col, row, colOff, rowOff


def _group_transform(anchor_origin_x, anchor_origin_y, grp_node):
    """
    计算 xdr:grpSp 的子坐标系(chOff/chExt) -> anchor-local 绝对坐标(EMU) 的线性映射。
    返回 dict: origin_x/origin_y（group 显示矩形左上角，相对 anchor 起点）、
    ch_off_x/ch_off_y（子坐标系原点）、scale_x/scale_y（子坐标系 -> 显示坐标的缩放）。
    任意环节缺失时按恒等映射（scale=1，off=0）降级，不伪造尺寸。
    """
    off_x = off_y = 0
    ext_cx = ext_cy = None
    ch_off_x = ch_off_y = 0
    ch_ext_cx = ch_ext_cy = None

    grpSpPr = grp_node.find('xdr:grpSpPr', NS)
    xfrm = grpSpPr.find('a:xfrm', NS) if grpSpPr is not None else None
    if xfrm is not None:
        off = xfrm.find('a:off', NS)
        if off is not None:
            off_x = int(off.get('x', 0))
            off_y = int(off.get('y', 0))
        ext = xfrm.find('a:ext', NS)
        if ext is not None:
            ext_cx = int(ext.get('cx'))
            ext_cy = int(ext.get('cy'))
        chOff = xfrm.find('a:chOff', NS)
        if chOff is not None:
            ch_off_x = int(chOff.get('x', 0))
            ch_off_y = int(chOff.get('y', 0))
        chExt = xfrm.find('a:chExt', NS)
        if chExt is not None:
            ch_ext_cx = int(chExt.get('cx'))
            ch_ext_cy = int(chExt.get('cy'))

    if not ch_ext_cx:
        ch_ext_cx = ext_cx
    if not ch_ext_cy:
        ch_ext_cy = ext_cy

    scale_x = (ext_cx / ch_ext_cx) if (ext_cx and ch_ext_cx) else 1.0
    scale_y = (ext_cy / ch_ext_cy) if (ext_cy and ch_ext_cy) else 1.0

    return {
        'origin_x': anchor_origin_x + off_x,
        'origin_y': anchor_origin_y + off_y,
        'ch_off_x': ch_off_x,
        'ch_off_y': ch_off_y,
        'scale_x': scale_x,
        'scale_y': scale_y,
    }


def _parse_sp_shape(sp, col, row, colOff, rowOff, cx, cy, in_group=False):
    """
    解析单个 xdr:sp 节点为统一的 shape dict。
    col/row/colOff/rowOff/cx/cy 由调用方按坐标系（顶层 anchor 或 group 换算后）传入，
    均已是相对 anchor 起点的绝对显示坐标/尺寸（EMU）。
    """
    shape = {'kind': 'sp', 'col': col, 'row': row,
             'colOff': colOff, 'rowOff': rowOff,
             'cx': cx, 'cy': cy, 'in_group': in_group}

    spPr = sp.find('xdr:spPr', NS)
    shape['geom'] = None
    shape['poly_pts'] = 0
    shape['is_closed'] = False
    shape['path_pts'] = []      # 自由多边形各顶点 (x, y)，单位 EMU（相对路径坐标系）
    shape['path_w'] = None      # custGeom 路径坐标系宽 (a:path @w)
    shape['path_h'] = None      # custGeom 路径坐标系高 (a:path @h)
    if spPr is not None:
        prst = spPr.find('a:prstGeom', NS)
        cust = spPr.find('a:custGeom', NS)
        if cust is not None:
            shape['geom'] = 'custGeom'  # 自由曲线/多边形（可编辑顶点）
            # 统计路径中的顶点数量、是否闭合，并提取每个顶点坐标
            path = cust.find('.//a:path', NS)
            if path is not None:
                if path.get('w'):
                    shape['path_w'] = int(path.get('w'))
                if path.get('h'):
                    shape['path_h'] = int(path.get('h'))
                pts = []
                # 按文档顺序遍历 moveTo/lnTo，提取顶点
                for node in list(path):
                    tag = node.tag.split('}')[-1]
                    if tag in ('moveTo', 'lnTo'):
                        ptn = node.find('a:pt', NS)
                        if ptn is not None:
                            pts.append((int(ptn.get('x')), int(ptn.get('y'))))
                shape['path_pts'] = pts
                shape['poly_pts'] = len(pts)
                shape['is_closed'] = path.find('a:close', NS) is not None
        elif prst is not None:
            shape['geom'] = prst.get('prst')  # 预设形状（如 pentagon）

        # ext 兜底：从 xfrm 读取（仅当调用方未提供尺寸时）
        if shape.get('cx') is None:
            xfrm = spPr.find('a:xfrm', NS)
            if xfrm is not None:
                ex = xfrm.find('a:ext', NS)
                if ex is not None:
                    shape['cx'] = int(ex.get('cx'))
                    shape['cy'] = int(ex.get('cy'))

        # 填充
        shape['fill'] = None
        if spPr.find('a:noFill', NS) is not None:
            shape['fill'] = 'none'
        else:
            sf = spPr.find('a:solidFill', NS)
            if sf is not None:
                clr = sf.find('a:srgbClr', NS)
                shape['fill'] = clr.get('val') if clr is not None else 'solid'

        # 轮廓线
        ln = spPr.find('a:ln', NS)
        shape['line_w'] = None
        shape['line_color'] = None
        shape['dash'] = None
        if ln is not None:
            if ln.get('w'):
                shape['line_w'] = int(ln.get('w'))
            lf = ln.find('a:solidFill', NS)
            if lf is not None:
                lc = lf.find('a:srgbClr', NS)
                shape['line_color'] = lc.get('val') if lc is not None else None
            dash = ln.find('a:prstDash', NS)
            shape['dash'] = dash.get('val') if dash is not None else None

        # 特效（阴影/发光/立体）
        effects = []
        lst = spPr.find('a:effectLst', NS)
        if lst is not None and len(list(lst)):
            effects.append('effectLst')
        if spPr.find('a:scene3d', NS) is not None:
            effects.append('scene3d')
        if spPr.find('a:sp3d', NS) is not None:
            effects.append('sp3d')
        shape['effects'] = effects

    # 名称
    nv = sp.find('xdr:nvSpPr/xdr:cNvPr', NS)
    shape['name'] = nv.get('name') if nv is not None else ''

    return shape


def _parse_group_children(grp, anchor_col, anchor_row, transform, data, depth=0):
    """
    递归展开 xdr:grpSp 内的子 xdr:sp（及嵌套 xdr:grpSp），换算到 anchor-local 绝对坐标后
    追加到 data.shapes（kind='sp', in_group=True）。
    """
    if depth > 4:   # 防御性限制递归深度，避免异常文件导致死循环
        return

    scale_x = transform['scale_x']
    scale_y = transform['scale_y']

    for child_sp in grp.findall('xdr:sp', NS):
        spPr = child_sp.find('xdr:spPr', NS)
        loc_x = loc_y = 0
        ext_cx = ext_cy = None
        if spPr is not None:
            xfrm = spPr.find('a:xfrm', NS)
            if xfrm is not None:
                off = xfrm.find('a:off', NS)
                if off is not None:
                    loc_x = int(off.get('x', 0))
                    loc_y = int(off.get('y', 0))
                ext = xfrm.find('a:ext', NS)
                if ext is not None:
                    ext_cx = int(ext.get('cx'))
                    ext_cy = int(ext.get('cy'))

        abs_colOff = int(round(transform['origin_x'] +
                                (loc_x - transform['ch_off_x']) * scale_x))
        abs_rowOff = int(round(transform['origin_y'] +
                                (loc_y - transform['ch_off_y']) * scale_y))
        abs_cx = int(round(ext_cx * scale_x)) if ext_cx is not None else None
        abs_cy = int(round(ext_cy * scale_y)) if ext_cy is not None else None

        data.shapes.append(_parse_sp_shape(
            child_sp, anchor_col, anchor_row, abs_colOff, abs_rowOff,
            abs_cx, abs_cy, in_group=True))

    for nested_grp in grp.findall('xdr:grpSp', NS):
        nested_transform = _group_transform(
            transform['origin_x'], transform['origin_y'], nested_grp)
        # 嵌套 group 的 off 是相对父 group 显示坐标系的，需要先按父 transform 映射一次；
        # 简化为在父 origin 基础上继续叠加，保持与顶层 _group_transform 相同语义。
        _parse_group_children(nested_grp, anchor_col, anchor_row,
                               nested_transform, data, depth=depth + 1)


def _parse_drawing(dpath, data):
    """解析 drawing xml，提取所有锚定在 C32（col=2,row=31）附近的形状。"""
    dtree = ET.parse(dpath).getroot()

    anchors = []
    anchors += dtree.findall('xdr:oneCellAnchor', NS)
    anchors += dtree.findall('xdr:twoCellAnchor', NS)
    anchors += dtree.findall('xdr:absoluteAnchor', NS)

    for anch in anchors:
        col, row, colOff, rowOff = _parse_anchor_from(anch)

        ext = anch.find('xdr:ext', NS)
        anchor_cx = int(ext.get('cx')) if ext is not None else None
        anchor_cy = int(ext.get('cy')) if ext is not None else None

        # 顶层 sp（可能存在多个）
        for sp in anch.findall('xdr:sp', NS):
            data.shapes.append(_parse_sp_shape(
                sp, col, row, colOff, rowOff, anchor_cx, anchor_cy, in_group=False))

        # 图片：只记录占位信息用于维度1"非图片"判定
        for pic in anch.findall('xdr:pic', NS):
            data.shapes.append({
                'kind': 'pic', 'col': col, 'row': row,
                'cx': None, 'cy': None, 'geom': None,
                'fill': None, 'line_w': None, 'line_color': None,
                'dash': None, 'effects': [], 'name': '',
            })

        # 图形组：记录组占位信息，并展开组内子形状为可评估的 sp
        for grp in anch.findall('xdr:grpSp', NS):
            data.shapes.append({
                'kind': 'group', 'col': col, 'row': row,
                'cx': anchor_cx, 'cy': anchor_cy, 'geom': None,
                'fill': None, 'line_w': None, 'line_color': None,
                'dash': None, 'effects': [], 'name': '',
            })
            transform = _group_transform(colOff, rowOff, grp)
            _parse_group_children(grp, col, row, transform, data)


# ----------------------------------------------------------------------------
# 维度1：可用与可修改性（门槛）
# ----------------------------------------------------------------------------
def check_dimension1(data):
    """返回 (passed: bool, details: list[(条目, 通过?, 说明)])"""
    details = []
    passed = True

    # 1.1 格式 .xlsx/.xlsm 且可正常打开
    fmt_ok = data.ext in ('.xlsx', '.xlsm') and data.ok_open
    details.append((
        '交付文件为xlsx或.xlsm格式，文件可正常打开。',
        fmt_ok,
        '扩展名=%s, 可打开=%s%s' % (
            data.ext, data.ok_open,
            ('（%s）' % data.parse_error) if data.parse_error else '')
    ))
    if not fmt_ok:
        passed = False

    return passed, details


def _shapes_in_c32(data):
    """C32 = col index 2, row index 31。返回锚定于该单元格的形状。"""
    return [s for s in data.shapes
            if s.get('col') == 2 and s.get('row') == 31]


# ----------------------------------------------------------------------------
# 自由多边形几何辅助（用于 +3 第一条逐句判定）
# ----------------------------------------------------------------------------
def _dedup_closed_pts(pts):
    """去掉闭合路径中重复的首尾点，返回有序顶点列表。"""
    if len(pts) >= 2 and pts[0] == pts[-1]:
        pts = pts[:-1]
    return pts


def _is_equilateral_polygon(pts, tol=0.20):
    """等边判定：各边长度的相对极差 <= tol（默认20%容差）。"""
    p = _dedup_closed_pts(list(pts))
    if len(p) < 3:
        return False
    lens = []
    n = len(p)
    for i in range(n):
        x1, y1 = p[i]
        x2, y2 = p[(i + 1) % n]
        lens.append(((x2 - x1) ** 2 + (y2 - y1) ** 2) ** 0.5)
    if not lens or min(lens) == 0:
        return False
    return (max(lens) - min(lens)) / max(lens) <= tol


def _polygon_tilt_about_30(pts, target=30.0, tol=10.0):
    """
    判定多边形整体倾斜约 target 度（默认30°，容差±tol）。
    取多边形最长边的方向角作为整体朝向的近似度量。
    容差收紧为±10度（即20°~40°），此前±20度(10°~50°)过于宽松，
    会把明显不是"约30度"的倾斜也判定通过。
    """
    p = _dedup_closed_pts(list(pts))
    if len(p) < 3:
        return False
    import math
    n = len(p)
    best_len = -1
    best_ang = None
    for i in range(n):
        x1, y1 = p[i]
        x2, y2 = p[(i + 1) % n]
        d = ((x2 - x1) ** 2 + (y2 - y1) ** 2) ** 0.5
        if d > best_len:
            best_len = d
            ang = math.degrees(math.atan2(abs(y2 - y1), abs(x2 - x1)))
            best_ang = ang
    if best_ang is None:
        return False
    return abs(best_ang - target) <= tol


def _apex_on_side(pts, side='right'):
    """
    判定多边形“最尖的顶点”（内角最小者）是否位于图形的左/右半部。
    side='right' -> 尖角在右半（x 大于中点）；'left' -> 在左半。
    """
    import math
    p = _dedup_closed_pts(list(pts))
    if len(p) < 3:
        return False
    xs = [pt[0] for pt in p]
    mid_x = (min(xs) + max(xs)) / 2.0
    n = len(p)
    min_angle = None
    apex_x = None
    for i in range(n):
        ax, ay = p[(i - 1) % n]
        bx, by = p[i]
        cx, cy = p[(i + 1) % n]
        v1 = (ax - bx, ay - by)
        v2 = (cx - bx, cy - by)
        n1 = math.hypot(*v1)
        n2 = math.hypot(*v2)
        if n1 == 0 or n2 == 0:
            continue
        cosv = max(-1.0, min(1.0, (v1[0] * v2[0] + v1[1] * v2[1]) / (n1 * n2)))
        ang = math.degrees(math.acos(cosv))
        if min_angle is None or ang < min_angle:
            min_angle = ang
            apex_x = bx
    if apex_x is None:
        return False
    if side == 'right':
        return apex_x >= mid_x
    return apex_x <= mid_x


def _has_30_sequence(text):
    """精确检查文本是否包含序号"30、"（数字30 + 中文顿号），而非仅包含子串'30'。"""
    return '30、' in (text or '')


def _normalize_points(pts):
    """将点集按自身 bbox 归一化到 [0,1]x[0,1]。bbox 退化（宽或高为0）时返回 None。"""
    p = _dedup_closed_pts(list(pts))
    if len(p) < 3:
        return None
    xs = [x for x, _ in p]
    ys = [y for _, y in p]
    min_x, max_x = min(xs), max(xs)
    min_y, max_y = min(ys), max(ys)
    w = max_x - min_x
    h = max_y - min_y
    if w <= 0 or h <= 0:
        return None
    return [((x - min_x) / w, (y - min_y) / h) for x, y in p]


def _mirror_reflection_error(left_pts, right_pts):
    """
    计算 left 关于竖直中线镜像后与 right 的最佳匹配误差。
    分别归一化两组点到各自 bbox 后比较，可容忍两侧尺寸的小幅差异；
    同时枚举顶点起点偏移与正/反两种绕行方向，取误差最小的匹配。
    返回 (rms_error, max_error)；无法判定时返回 (None, None)。
    """
    L = _normalize_points(left_pts)
    R = _normalize_points(right_pts)
    if L is None or R is None or len(L) != len(R):
        return None, None

    n = len(L)
    Lm = [(1.0 - x, y) for x, y in L]   # 关于 x=0.5 竖直镜像

    best_rms = None
    best_max = None
    for direction in (R, list(reversed(R))):
        for k in range(n):
            candidate = direction[k:] + direction[:k]
            d2_sum = 0.0
            max_d = 0.0
            for (x1, y1), (x2, y2) in zip(Lm, candidate):
                d = ((x1 - x2) ** 2 + (y1 - y2) ** 2) ** 0.5
                d2_sum += d * d
                if d > max_d:
                    max_d = d
            rms = (d2_sum / n) ** 0.5
            if best_rms is None or rms < best_rms:
                best_rms = rms
                best_max = max_d

    return best_rms, best_max


def _is_mirror_reflection(left_pts, right_pts, tol=0.08):
    """
    判定 left_pts 关于竖直轴镜像后是否与 right_pts 近似重合（真正的几何镜像判定，
    而非仅比较高度）。tol 为归一化坐标下的容差（bbox 尺度的比例）。
    """
    rms, max_err = _mirror_reflection_error(left_pts, right_pts)
    if rms is None:
        return False
    return (rms <= tol) and (max_err <= tol * 2.0)


# ----------------------------------------------------------------------------
# 维度2：完成度评分细则
# ----------------------------------------------------------------------------
def check_dimension2(data):
    """返回 (score, hits: list[(分值, 条目, 命中?, 说明)])"""
    hits = []
    score = 0

    c32 = _shapes_in_c32(data)
    # 待评估的自由多边形：既包含顶层独立 sp，也包含已组合进 group 内的子 sp
    # （_parse_drawing 已将 group 子形状换算为相对 C32 anchor 的绝对坐标并展开为 kind='sp'）
    sp_shapes = [s for s in c32 if s['kind'] == 'sp']

    # ---- 公共度量 ----
    # 以 colOff 排序区分左/右多边形
    sp_sorted = sorted(sp_shapes, key=lambda s: s.get('colOff', 0))

    # ============================================================
    # +3 ：C32 可编辑图形
    #   严格逐句对应细则原文。子项全部满足才 +3。
    # ============================================================
    plus3a_pts = []   # 记录每一个细则点

    # ---- 细则点① ：“30、”之后存在2个 宽1.1-1.3cm 高0.9-1.1cm 的
    #                可编辑的闭合自由多边形对象 ----
    cust_closed = [s for s in sp_shapes
                   if s.get('geom') == 'custGeom' and s.get('is_closed')]
    # 数量恰为2，且均为闭合自由多边形(custGeom + a:close)
    cond_count = (len(sp_shapes) == 2 and len(cust_closed) == 2)
    # 每个对象尺寸：宽1.1-1.3cm、高0.9-1.1cm
    size_ok = cond_count and all(
        (s.get('cx') and 1.1 <= emu_to_cm(s['cx']) <= 1.3) and
        (s.get('cy') and 0.9 <= emu_to_cm(s['cy']) <= 1.1)
        for s in sp_shapes)
    # “30、”之后：图形位于 C32 中“30、”文本右侧（colOff > 0，即不在单元格最左）
    c32_text = data.cell_text.get('C32', '')
    after_30 = _has_30_sequence(c32_text) and all(s.get('colOff', 0) > 0 for s in sp_shapes)
    cond1 = bool(cond_count and size_ok and after_30)
    plus3a_pts.append((
        '“30、”后存在2个宽1.1-1.3cm高0.9-1.1cm的可编辑闭合自由多边形',
        cond1,
        '形状数=%d, 闭合自由多边形数=%d, 各尺寸=%s, C32文本=%r, 均在“30、”右侧=%s' % (
            len(sp_shapes), len(cust_closed),
            [('%.2fx%.2fcm' % (emu_to_cm(s.get('cx') or 0), emu_to_cm(s.get('cy') or 0)))
             for s in sp_shapes],
            c32_text, all(s.get('colOff', 0) > 0 for s in sp_shapes))))

    # ---- 细则点② ：每个对象均可单独选择并修改顶点、轮廓和尺寸 ----
    # 可单独选择：各为独立 sp 对象（非锁定为不可选）；可改顶点：custGeom 有顶点；
    # 可改尺寸：存在 ext/xfrm 尺寸；可改轮廓：存在 a:ln 轮廓定义。
    each_editable = cond_count and all(
        (s.get('geom') == 'custGeom') and (len(s.get('path_pts', [])) > 0) and
        (s.get('cx') is not None and s.get('cy') is not None) and
        (s.get('line_w') is not None or s.get('line_color') is not None)
        for s in sp_shapes)
    cond2 = bool(each_editable)
    plus3a_pts.append((
        '每个对象均可单独选择并修改顶点、轮廓和尺寸',
        cond2,
        '各对象[独立=%s,顶点数=%s,有尺寸=%s,有轮廓=%s]' % (
            len(sp_shapes),
            [len(s.get('path_pts', [])) for s in sp_shapes],
            [s.get('cx') is not None for s in sp_shapes],
            [(s.get('line_w') is not None or s.get('line_color') is not None)
             for s in sp_shapes])))

    # ---- 细则点③ ：两个对象可组合为一个图形组 ----
    # “可组合”既包括两个独立形状（数量为2、均为可选形状，几何上可被组合），
    # 也包括已经组合进同一个 xdr:grpSp 的两个子形状（此时顶层不会再有独立 sp，
    # 但 sp_shapes 已通过 group 展开纳入这两个子形状，且需确认它们确实同属一个组）。
    group_shapes_here = [s for s in c32 if s.get('kind') == 'group']
    has_group = bool(group_shapes_here)
    group_child_count = sum(1 for s in sp_shapes if s.get('in_group'))
    independent_count = sum(1 for s in sp_shapes if not s.get('in_group'))
    cond3 = bool(cond_count and (
        (has_group and group_child_count == 2) or
        (not has_group and independent_count == 2)
    ))
    plus3a_pts.append((
        '两个对象可组合为一个图形组',
        cond3,
        '待评估sp=%d, 已存在图形组=%s, 组内sp=%d, 独立sp=%d' % (
            len(sp_shapes), has_group, group_child_count, independent_count)))

    # 排序：左/右多边形（按 colOff）
    if len(sp_shapes) == 2:
        left, right = sorted(sp_shapes, key=lambda s: s.get('colOff', 0))
    else:
        left = right = None

    # ---- 细则点④ ：左侧多边形位于序号“30、”右侧，由5个主要折点组成，
    #                轮廓为倾斜约30度的等边五边形，右侧尖角伸向单元格中部 ----
    if left is not None:
        l_after30 = (left.get('colOff', 0) > 0) and _has_30_sequence(c32_text)
        l_5pts = (left.get('poly_pts', 0) == 5)
        l_equilateral = _is_equilateral_polygon(left.get('path_pts', []))
        l_tilt = _polygon_tilt_about_30(left.get('path_pts', []))
        # 右侧尖角伸向单元格中部：最“尖”的顶点（夹角最小）位于多边形右半部
        l_apex_right = _apex_on_side(left.get('path_pts', []), side='right')
        cond4 = bool(l_after30 and l_5pts and l_equilateral and l_tilt and l_apex_right)
        c4_desc = ('在“30、”右侧=%s, 折点=%d, 等边=%s, 倾斜约30度=%s, 右尖角朝中部=%s'
                   % (l_after30, left.get('poly_pts', 0), l_equilateral, l_tilt, l_apex_right))
    else:
        cond4 = False
        c4_desc = '左侧多边形不存在（形状数!=2）'
    plus3a_pts.append((
        '左侧：在“30、”右侧/5折点/倾斜约30度等边五边形/右尖角伸向中部',
        cond4, c4_desc))

    # ---- 细则点⑤ ：右侧多边形位于左侧多边形右边，与左侧近似镜像排列，
    #                由5个主要折点组成，左侧尖角伸向单元格中部 ----
    if left is not None and right is not None:
        r_right_of_left = (right.get('colOff', 0) > left.get('colOff', 0))
        r_5pts = (right.get('poly_pts', 0) == 5)
        # 近似镜像：比较左右路径坐标的真实几何镜像关系（而非仅高度相近）
        r_mirror_rms, r_mirror_max = _mirror_reflection_error(
            left.get('path_pts', []), right.get('path_pts', []))
        r_mirror = _is_mirror_reflection(left.get('path_pts', []), right.get('path_pts', []))
        # 左侧尖角伸向单元格中部：右多边形最尖顶点位于其左半部
        r_apex_left = _apex_on_side(right.get('path_pts', []), side='left')
        cond5 = bool(r_right_of_left and r_5pts and r_mirror and r_apex_left)
        c5_desc = ('在左侧右边=%s, 折点=%d, 近似镜像=%s(rms=%s, max=%s), 左尖角朝中部=%s'
                   % (r_right_of_left, right.get('poly_pts', 0), bool(r_mirror),
                      ('%.4f' % r_mirror_rms) if r_mirror_rms is not None else 'N/A',
                      ('%.4f' % r_mirror_max) if r_mirror_max is not None else 'N/A',
                      r_apex_left))
    else:
        cond5 = False
        c5_desc = '右侧多边形不存在（形状数!=2）'
    plus3a_pts.append((
        '右侧：在左侧右边/与左近似镜像/5折点/左尖角伸向中部',
        cond5, c5_desc))

    # ---- 细则点⑥ ：左右两个多边形相互交叠，交叉轮廓在图形中部
    #                形成一个清晰的竖向菱形区域 ----
    if left is not None and right is not None:
        left_end = left.get('colOff', 0) + (left.get('cx') or 0)
        right_start = right.get('colOff', 0)
        overlap = left_end - right_start                 # 水平交叠量
        overlapped = overlap > 0
        # “竖向菱形”：交叠区竖向(高度)大于横向(交叠宽度) => 竖长菱形
        overlap_h = min(left.get('cy') or 0, right.get('cy') or 0)
        vertical_diamond = overlapped and (overlap_h > overlap)
        # “清晰”：交叠量适中（占单个多边形宽度的合理比例，约5%~80%）
        avg_w = ((left.get('cx') or 0) + (right.get('cx') or 0)) / 2.0
        clear = overlapped and avg_w > 0 and (0.05 * avg_w <= overlap <= 0.80 * avg_w)
        cond6 = bool(overlapped and vertical_diamond and clear)
        c6_desc = ('水平交叠=%.2fcm, 交叠区高=%.2fcm, 竖向菱形=%s, 清晰(占宽%.0f%%)=%s'
                   % (emu_to_cm(overlap), emu_to_cm(overlap_h), vertical_diamond,
                      (overlap / avg_w * 100) if avg_w else 0, clear))
    else:
        cond6 = False
        c6_desc = '形状数!=2，无法判定交叠'
    plus3a_pts.append((
        '左右相互交叠/中部形成清晰的竖向菱形区域',
        cond6, c6_desc))

    plus3a_all = all(p[1] for p in plus3a_pts)
    hits.append((3, '+3 C32 可编辑图形（双自由五边形/镜像/中部竖向菱形交叠）',
                 plus3a_all, plus3a_pts))
    if plus3a_all:
        score += 3

    # ============================================================
    # +3 ：C32 图形尺寸
    #   严格逐句对应细则原文。子项全部满足才 +3。
    # ============================================================
    plus3b_pts = []

    # ---- 细则点① ：组合图形整体宽 2.2–2.5cm、高 0.9–1.1cm ----
    if sp_sorted:
        min_x = min(s.get('colOff', 0) for s in sp_sorted)
        max_x = max(s.get('colOff', 0) + (s.get('cx') or 0) for s in sp_sorted)
        total_w = emu_to_cm(max_x - min_x)
        total_h = emu_to_cm(max((s.get('cy') or 0) for s in sp_sorted))
    else:
        total_w = total_h = 0.0
    cond_size = bool(sp_sorted) and (2.2 <= total_w <= 2.5) and (0.9 <= total_h <= 1.1)
    plus3b_pts.append(('组合图形整体宽2.2–2.5cm、高0.9–1.1cm', cond_size,
                       '整体宽=%.2fcm, 整体高=%.2fcm' % (total_w, total_h)))

    # ---- 细则点② ：两个多边形高度基本一致 ----
    if len(sp_sorted) == 2 and sp_sorted[0].get('cy') and sp_sorted[1].get('cy'):
        h1, h2 = sp_sorted[0]['cy'], sp_sorted[1]['cy']
        cond_same_h = abs(h1 - h2) <= 0.1 * max(h1, h2)
        same_h_desc = '高1=%.2fcm,高2=%.2fcm' % (emu_to_cm(h1), emu_to_cm(h2))
    else:
        cond_same_h = False
        same_h_desc = '形状数!=2'
    plus3b_pts.append(('两个多边形高度基本一致', cond_same_h, same_h_desc))

    # ---- 细则点③ ：两个多边形均使用黑色或深灰色实线轮廓，线宽 1.5–2.5 磅 ----
    def _is_dark(hexc):
        """黑色或深灰色判定：平均亮度较低。"""
        if not hexc or len(hexc) < 6:
            return False
        try:
            rr = int(hexc[-6:-4], 16); gg = int(hexc[-4:-2], 16); bb = int(hexc[-2:], 16)
        except ValueError:
            return False
        return (rr + gg + bb) / 3.0 <= 90
    line_ok = bool(sp_shapes)
    line_descs = []
    for s in sp_shapes:
        w_pt = emu_to_pt(s['line_w']) if s.get('line_w') else None
        dark = _is_dark(s.get('line_color'))           # 黑/深灰
        solid_line = (s.get('dash') in (None, 'solid'))  # 实线
        w_ok = (w_pt is not None and 1.5 <= w_pt <= 2.5)  # 线宽1.5–2.5磅
        line_descs.append('色=%s(黑/深灰=%s),实线=%s,线宽=%s磅' % (
            s.get('line_color'), dark, solid_line,
            ('%.2f' % w_pt) if w_pt is not None else 'N/A'))
        if not (dark and solid_line and w_ok):
            line_ok = False
    cond_line = line_ok
    plus3b_pts.append(('两多边形均黑/深灰实线轮廓、线宽1.5–2.5磅', cond_line,
                       ' | '.join(line_descs) if line_descs else '无形状'))

    # ---- 细则点④ ：无虚线、阴影、发光或立体效果 ----
    no_dash = bool(sp_shapes) and all(
        s.get('dash') in (None, 'solid') for s in sp_shapes)
    # 阴影/发光归于 effectLst，立体归于 scene3d/sp3d
    no_effects = all(not s.get('effects') for s in sp_shapes)
    cond_noeffect = bool(sp_shapes) and no_dash and no_effects
    plus3b_pts.append(('无虚线、阴影、发光或立体效果', cond_noeffect,
                       '虚线=%s, 特效(阴影/发光/立体)=%s' % (
                           [s.get('dash') for s in sp_shapes],
                           [s.get('effects') for s in sp_shapes])))

    # ---- 细则点⑤ ：无填充或白色透明填充，单元格浅色底纹能够正常显示 ----
    fill_ok = bool(sp_shapes)
    fill_descs = []
    for s in sp_shapes:
        f = s.get('fill')
        # 无填充(none) 或 白色填充(FFFFFF) => 不遮挡单元格浅色底纹
        ok = (f == 'none') or (f and f.upper() in ('FFFFFF', 'FFFFFFFF'))
        fill_descs.append(str(f))
        if not ok:
            fill_ok = False
    cond_fill = fill_ok
    plus3b_pts.append(('无填充或白色透明填充，浅色底纹可正常显示', cond_fill,
                       '填充=%s' % (fill_descs if fill_descs else '无形状')))

    plus3b_all = all(p[1] for p in plus3b_pts)
    hits.append((3, '+3 C32 图形尺寸/轮廓/填充', plus3b_all, plus3b_pts))
    if plus3b_all:
        score += 3

    # ============================================================
    # 扣分项（任意一点命中即扣）
    # ============================================================

    # -5 ：A3:C32 间存在“文本不显示”或“显示不清”的情况
    #   细则两个点：① 文本不显示 ② 文本显示不清，A3:C32 区域内任意一点命中即扣 -5。
    #   程序判定（基于 OOXML 可量化指标，对模糊措辞做灵活变通）：
    #     · 文本不显示：本应有文本的单元格内容为空（排除被合并覆盖的非左上角格），
    #       或字体颜色与所在单元格底纹色对比度过低到近乎不可辨（WCAG 对比度 <= 1.10）。
    #     · 显示不清：字体色与底纹色对比度低于常规可读标准（WCAG 对比度阈值，
    #       正常字号 4.50，18pt 及以上大字号放宽到 3.00）。
    #   空文本、截断、对比度三类判据分别使用独立的、有明确出处的阈值常量，互不混用。
    a3c32_pts = []

    # 截断判定阈值（估算行高/列宽是否放得下文本）
    CJK_CHAR_PX = 14.0           # 10pt 中文字符约占 14 像素宽
    DEFAULT_ROW_HEIGHT_PT = 15.0  # Excel 默认行高
    EST_LINE_HEIGHT_PT = 13.5     # 10pt 文本单行估算所需高度
    TRUNCATION_EPS_PT = 1.0       # 竖向截断判定的容差余量

    # 对比度判定阈值（WCAG 2.1 相对亮度/对比度公式）
    CONTRAST_NOT_DISPLAYED_MAX = 1.10   # 对比度<=此值：字体几乎被底纹“吞掉”，文本不显示
    CONTRAST_UNCLEAR_MIN_NORMAL = 4.50  # 正常字号可读所需的最低对比度（WCAG AA 正文标准）
    CONTRAST_UNCLEAR_MIN_LARGE = 3.00   # 大字号（>=18pt）可读所需的最低对比度（WCAG AA 大字标准）
    LARGE_TEXT_PT = 18.0

    def _relative_luminance(hexc):
        """WCAG 2.1 相对亮度：对 sRGB 做 gamma 校正后按 0.2126/0.7152/0.0722 加权。
        hexc 可为 8 位(含alpha)或 6 位；None/非法值返回 None。"""
        if not hexc or len(hexc) < 6:
            return None
        try:
            rr = int(hexc[-6:-4], 16); gg = int(hexc[-4:-2], 16); bb = int(hexc[-2:], 16)
        except ValueError:
            return None

        def _chan(c):
            c = c / 255.0
            return c / 12.92 if c <= 0.03928 else ((c + 0.055) / 1.055) ** 2.4

        r, g, b = _chan(rr), _chan(gg), _chan(bb)
        return 0.2126 * r + 0.7152 * g + 0.0722 * b

    def _contrast_ratio(fg_hex, bg_hex):
        """WCAG 对比度 = (较亮亮度+0.05) / (较暗亮度+0.05)；任一方不可解析返回 None。"""
        lf = _relative_luminance(fg_hex)
        lb = _relative_luminance(bg_hex)
        if lf is None or lb is None:
            return None
        lighter, darker = (lf, lb) if lf >= lb else (lb, lf)
        return (lighter + 0.05) / (darker + 0.05)

    def _cell_colors(ref):
        """返回 (字体色, 底纹色)。底纹为 none 时按白底处理；字体色未显式设置时按
        Excel“自动”默认的黑色处理（否则无法命中黑字黑底吞掉的情形）。"""
        s = data.cell_style.get(ref)
        if s is None:
            return '000000', 'FFFFFFFF'
        si = int(s)
        fid = data.xf_font.get(si, 0)
        flid = data.xf_fill.get(si, 0)
        fcolor = data.font_color.get(fid) or '000000'
        fill = data.fill_fg.get(flid)
        if not fill:                 # 无填充 -> 白底
            fill = 'FFFFFFFF'
        return fcolor, fill

    def _cell_font_size(ref):
        """解析单元格实际字号；无法解析时按 11pt（Excel 常见默认字号）兜底。"""
        s = data.cell_style.get(ref)
        if s is None:
            return 11.0
        fid = data.xf_font.get(int(s), 0)
        sz = data.font_size.get(fid)
        return sz if sz else 11.0

    def _merged_span(ref):
        """返回 ref 作为合并区左上角时的 (跨列数, 跨行数)；非左上角返回 (0,0)；独立格 (1,1)。"""
        col, row = _split_ref(ref)
        ci = _col_letter_to_idx(col)
        for m in data.merged:
            try:
                tl, br = m.split(':')
            except ValueError:
                continue
            tlc, tlr = _split_ref(tl)
            brc, brr = _split_ref(br)
            tlci, brci = _col_letter_to_idx(tlc), _col_letter_to_idx(brc)
            if tlci <= ci <= brci and tlr <= row <= brr:
                if ci == tlci and row == tlr:
                    return (brci - tlci + 1, brr - tlr + 1)
                return (0, 0)   # 被合并覆盖的非左上角
        return (1, 1)

    def _col_px(col_letter):
        """列字符宽 -> 像素。"""
        idx = _col_letter_to_idx(col_letter) + 1
        w = None
        for c in data.cols:
            if c['min'] <= idx <= c['max'] and c['width'] is not None:
                w = c['width']; break
        if w is None:
            w = 8.43
        return int(round(w * 7 + 5))

    def _check_truncation(ref):
        """
        判定单元格文本是否因行高放不下、且未开启自动换行/缩小填充而被截断
        （即 WPS/Excel 中“双击才完整显示”的情形）。
        返回 (是否截断, 说明)。
        """
        import math
        txt = data.cell_text.get(ref, '')
        if not txt.strip():
            return False, ''
        col, row = _split_ref(ref)
        span_c, span_r = _merged_span(ref)
        if span_c == 0:                      # 被合并覆盖的非左上角，跳过
            return False, ''

        si = data.cell_style.get(ref)
        wrap = data.xf_wrap.get(int(si)) if si is not None else False
        shrink = data.xf_shrink.get(int(si)) if si is not None else False
        if shrink:                           # 缩小字体填充 -> 始终可见，不截断
            return False, ''

        # 可用像素宽 = 该格起跨列的列宽之和
        total_px = 0
        ci0 = _col_letter_to_idx(col)
        for k in range(span_c):
            total_px += _col_px(chr(ord('A') + ci0 + k))
        per_line = max(1, int(total_px / CJK_CHAR_PX))

        # 文本所需总行数（硬换行 + 按列宽自动折行）
        if wrap:
            need_lines = 0
            for seg in txt.split('\n'):
                need_lines += max(1, math.ceil(len(seg) / per_line))
        else:
            # 未开自动换行：只按硬换行计行数，超宽部分横向溢出（可能被右侧非空格遮挡）
            need_lines = txt.count('\n') + 1

        # 可用行高 = 跨行行高之和
        total_ht = 0.0
        for k in range(span_r):
            total_ht += (data.rows.get(row + k) or DEFAULT_ROW_HEIGHT_PT)
        need_ht = need_lines * EST_LINE_HEIGHT_PT
        # 竖向截断：所需高度超过可用行高
        vertical_cut = need_ht > total_ht + TRUNCATION_EPS_PT

        # 横向截断：未开自动换行且单行宽度超过可用列宽，
        # 且右侧相邻格非空（溢出文字被遮挡 -> 双击才显示）
        horizontal_cut = False
        if not wrap:
            longest = max((len(seg) for seg in txt.split('\n')), default=0)
            if longest > per_line:
                # 右侧相邻列同行是否有内容
                right_letter = chr(ord('A') + ci0 + span_c)
                right_ref = '%s%d' % (right_letter, row)
                right_txt = data.cell_text.get(right_ref, '')
                if right_txt.strip():
                    horizontal_cut = True

        cut = vertical_cut or horizontal_cut
        if cut:
            return True, ('%s[换行=%s,需行=%d,需高≈%.0f/可用%.0f,竖截=%s,横截=%s]'
                          % (ref, wrap, need_lines, need_ht, total_ht,
                             vertical_cut, horizontal_cut))
        return False, ''

    not_displayed = []   # 文本不显示
    unclear = []         # 显示不清
    for r in range(3, 33):
        row_empty_refs = []
        row_has_text = False
        for col in ('A', 'B', 'C'):
            ref = '%s%d' % (col, r)
            # 被合并覆盖的非左上角格本就该为空，不参与判定
            if _is_inside_merge_non_topleft(data, ref):
                continue
            txt = data.cell_text.get(ref, '')
            has_text = (txt.strip() != '')

            # —— ① 文本不显示：逐单元格检查缺失（而非仅整行全空才算）——
            if not has_text:
                row_empty_refs.append(ref)
                continue
            row_has_text = True

            # 有文本：① 字体色与底纹对比度过低（几乎同色，被吞掉 -> 不显示）
            fcolor, fill = _cell_colors(ref)
            contrast = _contrast_ratio(fcolor, fill)
            if contrast is not None:
                font_size = _cell_font_size(ref)
                unclear_min = (CONTRAST_UNCLEAR_MIN_LARGE if font_size >= LARGE_TEXT_PT
                               else CONTRAST_UNCLEAR_MIN_NORMAL)
                if contrast <= CONTRAST_NOT_DISPLAYED_MAX:
                    not_displayed.append('%s(字体%s/底纹%s,对比度%.2f)' % (ref, fcolor, fill, contrast))
                    continue
                elif contrast < unclear_min:
                    unclear.append('%s(字体%s/底纹%s,对比度%.2f,阈值%.2f)'
                                    % (ref, fcolor, fill, contrast, unclear_min))

            # ① 文本因行高/列宽放不下被截断（双击才显示）-> 文本不显示
            cut, cut_desc = _check_truncation(ref)
            if cut:
                not_displayed.append(cut_desc)

        # 该行 A/B/C（除合并覆盖格外）全部为空 -> 折叠为一条整行提示，避免逐格重复；
        # 否则该行内确有文本却仍有格子缺失 -> 逐格上报，体现“逐单元格检查缺失”。
        if row_empty_refs and not row_has_text:
            not_displayed.append('第%d行A/B/C全空' % r)
        elif row_empty_refs:
            for ref in row_empty_refs:
                not_displayed.append('%s为空' % ref)

    # 点① 文本不显示
    p_notshow = len(not_displayed) > 0
    a3c32_pts.append(('A3:C32 存在文本不显示', p_notshow,
                      ('不显示单元格: %s' % not_displayed) if not_displayed
                      else 'A3:C32 文本均可正常显示'))
    # 点② 显示不清
    p_unclear = len(unclear) > 0
    a3c32_pts.append(('A3:C32 存在文本显示不清', p_unclear,
                      ('显示不清单元格: %s' % unclear) if unclear
                      else 'A3:C32 文本对比度正常'))

    text_unclear = p_notshow or p_unclear
    hits.append((-5, '-5 A3:C32 存在文本不显示或显示不清', text_unclear, a3c32_pts))
    if text_unclear:
        score -= 5

    return score, hits


# ----------------------------------------------------------------------------
# 列宽/行高工具函数已随边框/图形范围扣分项一起删除
# ----------------------------------------------------------------------------


def _is_inside_merge_non_topleft(data, ref):
    """判断 ref 是否位于某合并区域内但不是左上角（这种格子本就该为空）。"""
    col, row = _split_ref(ref)
    ci = _col_letter_to_idx(col)
    for m in data.merged:
        try:
            tl, br = m.split(':')
        except ValueError:
            continue
        tlc, tlr = _split_ref(tl)
        brc, brr = _split_ref(br)
        tlci, brci = _col_letter_to_idx(tlc), _col_letter_to_idx(brc)
        if tlci <= ci <= brci and tlr <= row <= brr:
            if not (ci == tlci and row == tlr):
                return True
    return False


# ----------------------------------------------------------------------------
# 主流程与报告
# ----------------------------------------------------------------------------
# 该脚本编号与被评估文档名（不含路径）
SCRIPT_ID = '093'
DOC_FILENAME = '日常信息处理能力观察表_图形组织可编辑版.xlsx'


def _build_dim2_items(hits):
    """将 check_dimension2 的 hits 转换为统一约定的 dim2_items 列表。

    hits 中每一项为 (pts, name, hit, subpts)；pts 为该项的分值（可正可负）。
    统一约定字段：rule / max_delta / delta / hit / detail。
    """
    items = []
    for pts, name, hit, subpts in hits:
        detail_parts = []
        for sub in subpts:
            # subpts 元素形如 (子项名, 是否命中, 说明)
            try:
                sname, sok, sdesc = sub
            except (ValueError, TypeError):
                continue
            detail_parts.append('%s=%s(%s)' % (sname, sok, sdesc))
        items.append({
            'rule': name,
            'max_delta': pts,
            'delta': pts if hit else 0,
            'hit': bool(hit),
            'detail': '',
        })
    return items


def evaluate(dir_path: str) -> dict:
    """统一入口：接收"脚本所在目录的路径"，脚本自己在该目录里定位并打开被评估文档。

    返回结构见《脚本接口差异与统一建议.md》§2.2。
    """
    result = {
        'id': SCRIPT_ID,
        'file_name': DOC_FILENAME,
        'status': 'ok',
        'error': None,
        'dim1_pass': False,
        'dim1_reason': '',
        'dim2_items': [],
        'total_score': 0,
        'max_score': 0,
    }

    try:
        # 在指定目录中定位被评估文档
        if not os.path.isdir(dir_path):
            raise FileNotFoundError(f"目录不存在: {dir_path}")
        file_path = os.path.join(dir_path, DOC_FILENAME)
        if not os.path.isfile(file_path):
            raise FileNotFoundError(f"目录内未找到必需文件: {DOC_FILENAME}")

        data = load_workbook(file_path)

        # ---- 维度1（门槛维度） ----
        d1_pass, d1_details = check_dimension1(data)
        result['dim1_pass'] = bool(d1_pass)

        if not d1_pass:
            # 汇总所有未通过条目的说明作为维度一未通过原因
            reasons = ['%s -> %s' % (name, desc)
                       for name, ok, desc in d1_details if not ok]
            result['dim1_reason'] = '; '.join(reasons)
            result['dim2_items'] = []
            result['total_score'] = 0
            # 满分固定为维度二各加分项 max_delta 之和（两条 +3 = 6）
            result['max_score'] = 6
            return result

        # ---- 维度2 ----
        score, hits = check_dimension2(data)
        dim2_items = _build_dim2_items(hits)
        result['dim2_items'] = dim2_items
        result['total_score'] = score
        # 满分：仅统计正向加分项的 max_delta 之和
        result['max_score'] = sum(it['max_delta'] for it in dim2_items
                                  if it['max_delta'] > 0)
        return result
    except Exception as e:  # noqa: BLE001  # 顶层兜底，避免影响批量运行
        result['status'] = 'error'
        result['error'] = '%s: %s' % (type(e).__name__, e)
        return result


if __name__ == '__main__':
    # 本地调试：接收"脚本所在目录路径"，默认取脚本自身所在目录
    default_dir = os.path.dirname(os.path.abspath(__file__))
    target_dir = sys.argv[1] if len(sys.argv) > 1 else default_dir
    print(json.dumps(evaluate(target_dir), ensure_ascii=False, indent=2))
