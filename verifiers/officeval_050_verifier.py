"""
PPT 自动评估脚本（officeval_050）

对外接口：
    evaluate(dir_path: str) -> dict
        接收脚本所在目录的路径；脚本自己在该目录内定位 .pptx 文档，
        返回结构化评估结果字典（含维度一通过与否、维度二逐项得分、总分）。
        字段规范见项目《脚本接口差异与统一建议》§2.2。

本地调试：
    python officeval_050_verifier.py [目录路径]
    未传参时默认使用脚本所在目录。主结果只走 return；此处仅将 dict
    以 JSON 打印到 stdout 供作者自测。
"""
import sys, json, re, math, zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

# ─── OOXML 命名空间 ───────────────────────────────────────────────────────────
NS = {
    'a':   'http://schemas.openxmlformats.org/drawingml/2006/main',
    'p':   'http://schemas.openxmlformats.org/presentationml/2006/main',
    'r':   'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    'xfrm':'http://schemas.openxmlformats.org/drawingml/2006/main',
}
def tag(ns, local): return '{'+NS[ns]+'}'+local

# ─── 单位转换 (EMU → cm) ─────────────────────────────────────────────────────
EMU_PER_CM = 360000
def emu2cm(e): return int(e) / EMU_PER_CM if e is not None else 0

# ─── 颜色工具 ────────────────────────────────────────────────────────────────
def hex2rgb(h):
    h = h.lstrip('#').upper()
    if len(h) != 6: return (0,0,0)
    return (int(h[0:2],16), int(h[2:4],16), int(h[4:6],16))

def color_dist(r1, r2):
    return math.sqrt(sum((a-b)**2 for a,b in zip(r1,r2)))

def is_black(rgb):   return color_dist(rgb,(0,0,0))<60
def is_white(rgb):   return color_dist(rgb,(255,255,255))<60
def is_blue(rgb):    r,g,b=rgb; return b>100 and b>(r+30) and b>(g+30)
def is_dark_blue(rgb): r,g,b=rgb; return b>80 and r<100 and g<130 and b>(r+40)
def is_light_blue(rgb): r,g,b=rgb; return b>150 and r>140 and g>170
def is_teal_blue(rgb): r,g,b=rgb; return g>140 and b>130 and r<140  # 蓝绿/青
def is_teal(rgb):
    """蓝绿色(teal/青)：绿、蓝分量都明显高于红，蓝≥绿，办公软件里呈蓝绿调，
    含深蓝绿如 107182(16,113,130)。"""
    r, g, b = rgb
    return g > r + 30 and b > r + 30 and b >= g - 20 and g >= 60
def is_green(rgb):   r,g,b=rgb; return g>100 and g>(r+20) and g>(b+20)
def is_dark_green(rgb): r,g,b=rgb; return g>80 and r<100 and b<100
def is_light_green(rgb): r,g,b=rgb; return g>160 and r>130 and b>130
def is_orange(rgb):  r,g,b=rgb; return r>180 and g>80 and b<100
def is_gray(rgb):    r,g,b=rgb; return abs(r-g)<30 and abs(g-b)<30 and r<200
def is_strict_gray(rgb):
    r, g, b = rgb
    return abs(r-g) < 25 and abs(g-b) < 25 and 80 <= r <= 180
def is_dark_gray(rgb):
    """深灰色：近中性(各通道接近)且整体偏暗，办公软件里呈深灰/近黑灰，含 202020 等。"""
    r, g, b = rgb
    return abs(r-g) < 30 and abs(g-b) < 30 and abs(r-b) < 30 and max(r, g, b) <= 130

# ─── PPTX 解析 ───────────────────────────────────────────────────────────────
class Shape:
    """单个形状的属性容器"""
    def __init__(self):
        self.name: str = ''
        self.descr: str = ''   # cNvPr@descr（alt text）
        self.shape_type = ''   # sp | pic | cxnSp | grpSp
        self.x = 0.0           # cm, 距左
        self.y = 0.0           # cm, 距上
        self.w = 0.0           # cm, 宽
        self.h = 0.0           # cm, 高
        self.rot = 0           # 1/60000 度; 正值顺时针
        self.texts = []        # [(run_text, font_name, font_size_pt, bold, color_hex)]
        self.text_aligns = []   # paragraph alignments, e.g. ctr/l/r
        self.paragraph_texts = [] # 每个段落/行的文本
        self.text_insets = None # (lIns,tIns,rIns,bIns) cm，None=未解析
        self.fill_color = None # hex str or None
        self.line_color = None
        self.line_width_pt = None
        self.line_dash = None  # solid / dash / dashDot / sysDash etc.
        self.has_tail_arrow = False
        self.has_head_arrow = False
        self.geom = None       # preset geometry name
        self.has_shadow: bool = False  # spPr/a:effectLst 下存在 outerShdw/innerShdw/prstShdw
        self.round_adj: float | None = None  # roundRect 圆角调整值(占较短边比例)，None=用默认
        self.is_pic = False
        self.slide_idx = 0
    @property
    def left(self): return min(self.x, self.x + self.w)
    @property
    def right(self): return max(self.x, self.x + self.w)
    @property
    def top(self): return min(self.y, self.y + self.h)
    @property
    def bottom(self): return max(self.y, self.y + self.h)
    @property
    def abs_w(self): return abs(self.w)
    @property
    def abs_h(self): return abs(self.h)
    def rotation_deg(self): return (self.rot / 60000) % 360
    def full_text(self): return ' '.join(t[0] for t in self.texts).strip()

def _parse_color_node(nd):
    """从 a:solidFill 或 a:ln 下的颜色节点提取颜色 hex"""
    if nd is None: return None
    sf = nd.find(f'.//{{{NS["a"]}}}solidFill')
    if sf is None: return None
    srgb = sf.find(f'{{{NS["a"]}}}srgbClr')
    if srgb is not None: return srgb.get('val','').upper()
    pst = sf.find(f'{{{NS["a"]}}}prstClr')
    if pst is not None:
        pmap={'black':'000000','white':'FFFFFF','blue':'0000FF','red':'FF0000',
              'green':'00FF00','yellow':'FFFF00','orange':'FFA500','gray':'808080',
              'darkBlue':'00008B','lightBlue':'ADD8E6','teal':'008080'}
        return pmap.get(pst.get('val',''),'000000').upper()
    return None

def _parse_shape(sp_el, shape_type, slide_idx):
    s = Shape()
    s.shape_type = shape_type
    s.slide_idx = slide_idx
    s.is_pic = (shape_type == 'pic')
    nva = sp_el.find(f'.//{{{NS["p"]}}}nvPr', NS)
    cnvp = sp_el.find(f'.//{{{NS["p"]}}}cNvPr', NS)
    if cnvp is None:
        cnvp = sp_el.find(f'.//{{"http://schemas.openxmlformats.org/drawingml/2006/main"}}cNvPr', NS)
    # name / descr from cNvPr (works for sp and pic)
    for ns_uri in ['http://schemas.openxmlformats.org/presentationml/2006/main',
                   'http://schemas.openxmlformats.org/drawingml/2006/main']:
        el = sp_el.find(f'.//{{{ns_uri}}}cNvPr')
        if el is not None:
            s.name = el.get('name', '') or ''
            s.descr = el.get('descr', '') or ''
            break

    # xfrm
    xfrm = sp_el.find(f'.//{{{NS["a"]}}}xfrm')
    if xfrm is None:
        xfrm = sp_el.find(f'.//{{{NS["p"]}}}xfrm')
    if xfrm is not None:
        s.rot = int(xfrm.get('rot','0') or 0)
        off = xfrm.find(f'{{{NS["a"]}}}off')
        ext = xfrm.find(f'{{{NS["a"]}}}ext')
        if off is not None: s.x=emu2cm(off.get('x','0')); s.y=emu2cm(off.get('y','0'))
        if ext is not None: s.w=emu2cm(ext.get('cx','0')); s.h=emu2cm(ext.get('cy','0'))

    # fill color
    spPr = sp_el.find(f'{{{NS["p"]}}}spPr')
    if spPr is None:
        spPr = sp_el.find(f'.//{{{NS["a"]}}}spPr')
    if spPr is not None:
        s.fill_color = _parse_color_node(spPr)
        # 阴影：spPr/a:effectLst 下若存在 outerShdw/innerShdw/prstShdw 等即视为带阴影
        effLst = spPr.find(f'{{{NS["a"]}}}effectLst')
        if effLst is not None:
            for shd_tag in ('outerShdw', 'innerShdw', 'prstShdw'):
                if effLst.find(f'{{{NS["a"]}}}{shd_tag}') is not None:
                    s.has_shadow = True
                    break
        ln = spPr.find(f'{{{NS["a"]}}}ln')
        if ln is not None:
            s.line_color = _parse_color_node(ln)
            w = ln.get('w')
            if w: s.line_width_pt = int(w)/12700  # EMU/pt
            pd = ln.find(f'{{{NS["a"]}}}prstDash')
            if pd is not None: s.line_dash = pd.get('val','solid')
            else: s.line_dash = 'solid'
            # arrows
            th = ln.find(f'{{{NS["a"]}}}tailEnd')
            if th is not None and th.get('type','none') not in ('none',''):
                s.has_tail_arrow = True
            hd = ln.find(f'{{{NS["a"]}}}headEnd')
            if hd is not None and hd.get('type','none') not in ('none',''):
                s.has_head_arrow = True
        prstGeom = spPr.find(f'{{{NS["a"]}}}prstGeom')
        if prstGeom is not None:
            s.geom = prstGeom.get('prst')
            # roundRect 圆角比例：avLst/gd@fmla="val N"，N/100000 为占较短边的比例
            gd = prstGeom.find(f'{{{NS["a"]}}}avLst/{{{NS["a"]}}}gd')
            if gd is not None:
                fmla = gd.get('fmla', '')
                if fmla.startswith('val '):
                    try: s.round_adj = int(fmla.split()[1]) / 100000
                    except (ValueError, IndexError): pass

    # text
    a_ns = NS['a']
    for txBody in sp_el.findall(f'.//{{{NS["p"]}}}txBody') + sp_el.findall(f'.//{{{a_ns}}}txBody'):
        bodyPr = txBody.find(f'{{{a_ns}}}bodyPr')
        if bodyPr is not None and s.text_insets is None:
            # 默认 PowerPoint 内边距: l/r=0.25cm(91440EMU), t/b=0.13cm(45720EMU)
            lIns = emu2cm(bodyPr.get('lIns', '91440'))
            tIns = emu2cm(bodyPr.get('tIns', '45720'))
            rIns = emu2cm(bodyPr.get('rIns', '91440'))
            bIns = emu2cm(bodyPr.get('bIns', '45720'))
            s.text_insets = (lIns, tIns, rIns, bIns)
        for p in txBody.findall(f'{{{a_ns}}}p'):
            pPr = p.find(f'{{{a_ns}}}pPr')
            if pPr is not None:
                s.text_aligns.append(pPr.get('algn', ''))
            para_runs = []
            for r in p.findall(f'{{{a_ns}}}r'):
                t_el = r.find(f'{{{a_ns}}}t')
                rPr  = r.find(f'{{{a_ns}}}rPr')
                if t_el is None: continue
                txt = t_el.text or ''
                para_runs.append(txt)
                font = ''; size_pt = 0; bold = False; col = None
                if rPr is not None:
                    sz = rPr.get('sz')
                    if sz: size_pt = int(sz)/100
                    b  = rPr.get('b')
                    if b in ('1','true'): bold = True
                    latin = rPr.find(f'{{{a_ns}}}latin')
                    if latin is not None: font = latin.get('typeface','')
                    col = _parse_color_node(rPr)
                s.texts.append((txt, font, size_pt, bold, col))
            if para_runs:
                s.paragraph_texts.append(''.join(para_runs).strip())
    return s

def load_pptx(path: 'str | Path') -> tuple[float, float, list[list['Shape']], dict[str, str]]:
    """返回 (page_w_cm, page_h_cm, slides, theme_fonts)
    slides 是 list[list[Shape]]；theme_fonts 是 {'minor':..,'major':..}，
    用于把 +mn-lt/+mj-lt 及未显式设置字体的文本解析为办公软件实际渲染的字体。"""
    slides_shapes: list[list[Shape]] = []
    page_w = page_h = 0.0
    theme_fonts = {'minor': '', 'major': ''}
    with zipfile.ZipFile(path) as z:
        # presentation dims
        prs_xml = z.read('ppt/presentation.xml')
        prs = ET.fromstring(prs_xml)
        sldSz = prs.find(f'{{{NS["p"]}}}sldSz')
        if sldSz is not None:
            page_w = emu2cm(sldSz.get('cx','0'))
            page_h = emu2cm(sldSz.get('cy','0'))
        # theme 主题字体（办公软件在文本未显式指定字体时按此渲染）
        theme_files = sorted([n for n in z.namelist()
                              if re.match(r'ppt/theme/theme\d+\.xml', n)])
        if theme_files:
            try:
                troot = ET.fromstring(z.read(theme_files[0]))
                fs = troot.find(f'.//{{{NS["a"]}}}fontScheme')
                if fs is not None:
                    mj = fs.find(f'{{{NS["a"]}}}majorFont/{{{NS["a"]}}}latin')
                    mn = fs.find(f'{{{NS["a"]}}}minorFont/{{{NS["a"]}}}latin')
                    if mj is not None: theme_fonts['major'] = mj.get('typeface', '')
                    if mn is not None: theme_fonts['minor'] = mn.get('typeface', '')
            except Exception:
                pass
        # slides
        slide_files = sorted([n for n in z.namelist()
                               if re.match(r'ppt/slides/slide\d+\.xml', n)],
                              key=lambda n: int(re.search(r'\d+', Path(n).stem).group()))
        for idx, sf in enumerate(slide_files):
            root = ET.fromstring(z.read(sf))
            shapes: list[Shape] = []
            spTree = root.find(f'.//{{{NS["p"]}}}spTree')
            if spTree is None: slides_shapes.append(shapes); continue
            # sp
            for el in spTree.findall(f'{{{NS["p"]}}}sp'):
                shapes.append(_parse_shape(el, 'sp', idx))
            # pic
            for el in spTree.findall(f'{{{NS["p"]}}}pic'):
                shapes.append(_parse_shape(el, 'pic', idx))
            # cxnSp (connector)
            for el in spTree.findall(f'{{{NS["p"]}}}cxnSp'):
                shapes.append(_parse_shape(el, 'cxnSp', idx))
            # grpSp — flatten one level
            for grp in spTree.findall(f'{{{NS["p"]}}}grpSp'):
                for el in grp.findall(f'{{{NS["p"]}}}sp'):
                    shapes.append(_parse_shape(el, 'sp', idx))
                for el in grp.findall(f'{{{NS["p"]}}}cxnSp'):
                    shapes.append(_parse_shape(el, 'cxnSp', idx))
                for el in grp.findall(f'{{{NS["p"]}}}pic'):
                    shapes.append(_parse_shape(el, 'pic', idx))
            slides_shapes.append(shapes)
    return page_w, page_h, slides_shapes, theme_fonts

# ─── 辅助筛选 ────────────────────────────────────────────────────────────────
def in_range(v, lo, hi): return lo <= v <= hi
def shape_in_box(s, xl, xr, yt, yb):
    """形状中心或左上角在矩形框内"""
    cx, cy = s.x + s.w/2, s.y + s.h/2
    return (xl <= s.x <= xr or xl <= cx <= xr) and (yt <= s.y <= yb or yt <= cy <= yb)
def shapes_in(shapes, xl, xr, yt, yb):
    return [s for s in shapes if shape_in_box(s, xl, xr, yt, yb)]
def texts_contain(shapes, keyword, exact=False):
    for s in shapes:
        ft = s.full_text()
        if exact and ft.strip()==keyword.strip(): return True
        if not exact and keyword.lower() in ft.lower(): return True
    return False
def get_texts(shapes, keyword=None):
    result=[]
    for s in shapes:
        ft=s.full_text()
        if keyword is None or keyword.lower() in ft.lower(): result.append(ft)
    return result

# ─── 颜色提取 ────────────────────────────────────────────────────────────────
def shape_fill_rgb(s):
    if s.fill_color: return hex2rgb(s.fill_color)
    return None
def shape_line_rgb(s):
    if s.line_color: return hex2rgb(s.line_color)
    return None
def text_run_rgb(tr):  # tr = (txt,font,size,bold,col_hex)
    if tr[4]: return hex2rgb(tr[4])
    return None

# ─── 旋转角度工具 ────────────────────────────────────────────────────────────
def near_angle(s, target_deg, tol=15):
    d = s.rotation_deg()
    diff = abs(d - target_deg) % 360
    if diff > 180: diff = 360 - diff
    return diff <= tol

def rendered_bbox(s):
    """返回旋转后的可见包围盒 (left, top, right, bottom)，单位 cm"""
    theta = math.radians(s.rotation_deg())
    bw = abs(s.w * math.cos(theta)) + abs(s.h * math.sin(theta))
    bh = abs(s.w * math.sin(theta)) + abs(s.h * math.cos(theta))
    cx, cy = s.x + s.w/2, s.y + s.h/2
    return cx - bw/2, cy - bh/2, cx + bw/2, cy + bh/2

def bbox_in_box(s, xl, xr, yt, yb):
    left, top, right, bottom = rendered_bbox(s)
    return xl <= left and right <= xr and yt <= top and bottom <= yb

def center_x(s):
    left, _, right, _ = rendered_bbox(s)
    return (left + right) / 2

def centered_in_x_range(s, xl, xr, tol=0.2):
    return abs(center_x(s) - (xl + xr) / 2) <= tol

def is_center_aligned(s):
    return bool(s.text_aligns) and all(a in ('ctr', 'center') for a in s.text_aligns if a)

def non_empty_runs(s):
    return [tr for tr in s.texts if tr[0].strip()]

# ─── 文本自动换行估算(反映办公软件所见的视觉行数/列数) ─────────────────────
# 优先用系统真实字体度量文字宽度；无字体文件时退回经验系数。
_CHAR_W_FACTOR = 0.52          # 退化估算：平均字宽 ≈ 字号 × 该系数
_FONT_FILES = {                # (font_lower, bold) -> Windows 字体文件
    ('arial', False): 'C:/Windows/Fonts/arial.ttf',
    ('arial', True):  'C:/Windows/Fonts/arialbd.ttf',
    ('calibri', False): 'C:/Windows/Fonts/calibri.ttf',
    ('calibri', True):  'C:/Windows/Fonts/calibrib.ttf',
}
_font_cache = {}

def _text_width_cm(text, font_name, size_pt, bold):
    """返回文本在给定字体/字号下的渲染宽度(cm)。有真实字体则用 PIL 度量。"""
    if not text:
        return 0.0
    key = ((font_name or 'arial').lower().split()[0], bool(bold), round(size_pt, 1))
    fpath = _FONT_FILES.get((key[0], key[1])) or _FONT_FILES.get(('arial', key[1]))
    if fpath:
        try:
            from PIL import ImageFont
            import os
            if not os.path.exists(fpath):
                raise FileNotFoundError
            DPI = 600
            fkey = (fpath, size_pt)
            font = _font_cache.get(fkey)
            if font is None:
                font = ImageFont.truetype(fpath, int(round(size_pt / 72 * DPI)))
                _font_cache[fkey] = font
            box = font.getbbox(text)
            return (box[2] - box[0]) / DPI * 2.54
        except Exception:
            pass
    # 退化：字符数 × 经验字宽
    return len(text) * (size_pt / 72 * 2.54 * _CHAR_W_FACTOR)

def estimate_wrapped_lines(s):
    """估算文本在其文本框内经自动换行后的可见行数(即竖排时的列数)。
    办公软件按“文本框可用宽度”(box 宽 - 左右内边距)对每个段落做贪心折行：
    单词依次排入，一行放不下就换行。返回视觉行数(≥ 段落数)。"""
    runs = non_empty_runs(s)
    if not runs:
        return len(s.paragraph_texts)
    lIns, _, rIns, _ = s.text_insets or (0.25, 0.13, 0.25, 0.13)
    avail_cm = abs(s.w) - lIns - rIns
    if avail_cm <= 0:
        return len(s.paragraph_texts)
    # 该形状字号/字体/加粗(取首个非空 run 为准)
    tr0 = runs[0]
    font_name, size_pt, bold = tr0[1], (tr0[2] or 12.0), tr0[3]
    total_lines = 0
    for para in (s.paragraph_texts or [s.full_text()]):
        words = [w for w in para.split() if w]
        if not words:
            continue
        space_w = _text_width_cm(' ', font_name, size_pt, bold) or \
                  _text_width_cm('n', font_name, size_pt, bold)
        lines, cur = 1, 0.0
        for i, w in enumerate(words):
            wlen = _text_width_cm(w, font_name, size_pt, bold)
            add = wlen if i == 0 else space_w + wlen
            if cur + add > avail_cm and cur > 0:
                lines += 1
                cur = wlen
            else:
                cur += add
        total_lines += lines
    return max(total_lines, len(s.paragraph_texts))

def text_props_ok(s, size_lo, size_hi, color_check, require_bold=True):
    runs = non_empty_runs(s)
    if not runs:
        return False
    for tr in runs:
        rgb = text_run_rgb(tr)
        if not run_font_ok(tr): return False
        if not (size_lo <= tr[2] <= size_hi): return False
        if require_bold and not tr[3]: return False
        if rgb is not None and not color_check(rgb): return False
    return True

def is_strict_black(rgb):
    return color_dist(rgb, (0, 0, 0)) < 35

def is_pale_blue_fill(rgb):
    r, g, b = rgb
    return b > 200 and g > 190 and b > r + 12 and b > g + 5 and color_dist(rgb, (255, 255, 255)) > 18

def is_light_blue_fill(rgb):
    """浅蓝色填充：整体明亮(高亮度)且蓝色分量占优，办公软件里呈浅蓝调。
    覆盖较饱和的浅蓝，也覆盖 F7FAFD 这类极浅的蓝色调。"""
    r, g, b = rgb
    return r > 180 and g > 190 and b > 200 and b >= g >= r and b > r

def is_strict_dark_blue(rgb):
    r, g, b = rgb
    return b >= 80 and r <= 60 and g <= 100 and b > g + 25 and b > r + 40

def is_white_or_blue_white(rgb):
    r, g, b = rgb
    return (r > 245 and g > 245 and b > 245) or (r > 230 and g > 240 and b > 245 and b >= g >= r - 10)

def is_pale_green_fill(rgb):
    r, g, b = rgb
    # 淡绿：整体明亮但绿色分量占优；排除纯白/近中性色。
    return (g > 235 and r > 225 and b > 225 and g >= r and g >= b
            and (g - min(r, b) >= 4 or g - r >= 2)
            and color_dist(rgb, (255, 255, 255)) > 6)

def is_pale_orange_fill(rgb):
    r, g, b = rgb
    # 淡橘：整体明亮但红分量占优、蓝分量较低；排除纯白/近中性色。
    return (r > 240 and g > 220 and b > 190 and r >= g >= b
            and (r - b >= 10) and color_dist(rgb, (255, 255, 255)) > 6)

def is_pale_any(rgb):
    """极浅色：各通道都很高、接近白，办公软件里呈几乎白的浅色调。"""
    r, g, b = rgb
    return min(r, g, b) >= 235

def group_bbox(shapes):
    left = min(s.left for s in shapes)
    top = min(s.top for s in shapes)
    right = max(s.right for s in shapes)
    bottom = max(s.bottom for s in shapes)
    return left, top, right, bottom

# ─── 箭头朝向 ────────────────────────────────────────────────────────────────
# "一体箭头"包含两种：
#   (A) 预设几何箭头形状（rightArrow/downArrow/…），整个 sp 就是一根箭头；
#   (B) 线条(prstGeom=line)/直线连接符(cxnSp) 的 <a:ln> 上带 tailEnd/headEnd 箭头标记。
# 两种在办公软件里都渲染为"一根整体的箭头"，判定时一视同仁。
ARROW_GEOMS_ONE_DIR = {
    'rightArrow', 'leftArrow', 'upArrow', 'downArrow',
    'notchedRightArrow', 'stripedRightArrow', 'bentArrow',
    'curvedRightArrow', 'curvedLeftArrow',
    'curvedUpArrow', 'curvedDownArrow',
    'circularArrow',
}
ARROW_GEOMS_TWO_DIR = {'leftRightArrow', 'upDownArrow', 'bentUpArrow'}
ARROW_GEOMS_QUAD   = {'quadArrow', 'leftRightUpArrow'}
ARROW_GEOMS_ALL = ARROW_GEOMS_ONE_DIR | ARROW_GEOMS_TWO_DIR | ARROW_GEOMS_QUAD

# 单向预设箭头(几何未旋转时)的基础朝向
_PRESET_ARROW_BASE_DIR = {
    'rightArrow': 'r', 'notchedRightArrow': 'r', 'stripedRightArrow': 'r',
    'bentArrow': 'r', 'curvedRightArrow': 'r', 'circularArrow': 'r',
    'leftArrow': 'l', 'curvedLeftArrow': 'l',
    'upArrow': 'u', 'curvedUpArrow': 'u',
    'downArrow': 'd', 'curvedDownArrow': 'd',
}
# PowerPoint 旋转按顺时针，rot=90° 时 "→" 变 "↓"
_ROT_DIR_TABLE = {
    ('r', 0): 'r', ('r', 1): 'd', ('r', 2): 'l', ('r', 3): 'u',
    ('l', 0): 'l', ('l', 1): 'u', ('l', 2): 'r', ('l', 3): 'd',
    ('u', 0): 'u', ('u', 1): 'r', ('u', 2): 'd', ('u', 3): 'l',
    ('d', 0): 'd', ('d', 1): 'l', ('d', 2): 'u', ('d', 3): 'r',
}

def is_preset_arrow(s):
    """形状本身是预设箭头几何(如 rightArrow/downArrow/leftRightArrow…)。"""
    return s.geom in ARROW_GEOMS_ALL

def is_line_arrow(s):
    """线条/连接符上带端点箭头标记(tailEnd/headEnd)。"""
    return bool(s.has_tail_arrow) or bool(s.has_head_arrow)

def is_arrow_shape(s):
    """一体箭头：预设箭头几何 或 线条+端点箭头标记。"""
    return is_preset_arrow(s) or is_line_arrow(s)

def has_any_arrow(s):
    """兼容旧调用：只要是一体箭头即算有箭头。"""
    return is_arrow_shape(s)

def arrow_count(s):
    """箭头个数：单向=1、双向=2、四向=4；不是一体箭头返回 0。
    与办公软件所见一致：leftRightArrow/upDownArrow 视为双箭头。"""
    if s.geom in ARROW_GEOMS_QUAD: return 4
    if s.geom in ARROW_GEOMS_TWO_DIR: return 2
    if s.geom in ARROW_GEOMS_ONE_DIR: return 1
    return int(bool(s.has_tail_arrow)) + int(bool(s.has_head_arrow))

def preset_arrow_dir(s):
    """单向预设箭头在当前旋转下的朝向 → 'u'/'d'/'l'/'r'；不是单向预设箭头返回 None。"""
    base = _PRESET_ARROW_BASE_DIR.get(s.geom)
    if not base: return None
    rot_quarter = int(round(s.rotation_deg() / 90)) % 4
    return _ROT_DIR_TABLE.get((base, rot_quarter))

def line_angle_from_horizontal(s):
    """线条与水平方向的夹角(度)，返回0–90。0=水平，90=竖直；含形状旋转，
    与办公软件里看到的线条朝向一致。"""
    if not (s.w or s.h):
        base = 0.0
    else:
        base = math.degrees(math.atan2(s.h, s.w))
    a = (base + s.rotation_deg()) % 180
    if a > 90: a = 180 - a
    return a

def arrow_points_up(s):
    """箭头是否朝上。
    (A) 预设箭头几何：按其基础朝向 + 形状旋转推算(单向箭头精确朝向；双向/四向含向上分量则视为朝上)。
    (B) 线条+端点箭头：按有箭头那一端相对可见包围盒中线的位置判定。
    办公软件在两种情况下都渲染为一根整体箭头，此处一视同仁。"""
    if is_preset_arrow(s):
        d = preset_arrow_dir(s)
        if d is not None: return d == 'u'
        return s.geom in ARROW_GEOMS_TWO_DIR and s.geom == 'upDownArrow' \
               or s.geom in ARROW_GEOMS_QUAD
    top_y = min(s.y, s.y + s.h)     # 可见顶部
    bottom_y = max(s.y, s.y + s.h)  # 可见底部
    if bottom_y - top_y < 0.05:      # 近水平线：不视为朝上/朝下
        return False
    mid = (top_y + bottom_y) / 2
    ends = []
    if s.has_tail_arrow: ends.append(bottom_y)
    if s.has_head_arrow: ends.append(top_y)
    return any(e < mid for e in ends)

def arrow_points_right(s):
    """箭头是否朝右。判定逻辑同 arrow_points_up。"""
    if is_preset_arrow(s):
        d = preset_arrow_dir(s)
        if d is not None: return d == 'r'
        return s.geom in ARROW_GEOMS_TWO_DIR and s.geom == 'leftRightArrow' \
               or s.geom in ARROW_GEOMS_QUAD
    left_x = min(s.x, s.x + s.w)
    right_x = max(s.x, s.x + s.w)
    if right_x - left_x < 0.05:      # 近垂直线：不视为朝左/朝右
        return False
    mid = (left_x + right_x) / 2
    ends = []
    if s.has_tail_arrow: ends.append(right_x)
    if s.has_head_arrow: ends.append(left_x)
    return any(e > mid for e in ends)

def arrow_points_down(s):
    """箭头是否朝下。规则同 arrow_points_up 对称。"""
    if is_preset_arrow(s):
        d = preset_arrow_dir(s)
        if d is not None: return d == 'd'
        return s.geom == 'upDownArrow' or s.geom in ARROW_GEOMS_QUAD
    top_y = min(s.y, s.y + s.h)
    bottom_y = max(s.y, s.y + s.h)
    if bottom_y - top_y < 0.05:
        return False
    mid = (top_y + bottom_y) / 2
    ends = []
    if s.has_tail_arrow: ends.append(bottom_y)
    if s.has_head_arrow: ends.append(top_y)
    return any(e > mid for e in ends)

def arrow_points_left(s):
    """箭头是否朝左。规则同 arrow_points_right 对称。"""
    if is_preset_arrow(s):
        d = preset_arrow_dir(s)
        if d is not None: return d == 'l'
        return s.geom == 'leftRightArrow' or s.geom in ARROW_GEOMS_QUAD
    left_x = min(s.x, s.x + s.w)
    right_x = max(s.x, s.x + s.w)
    if right_x - left_x < 0.05:
        return False
    mid = (left_x + right_x) / 2
    ends = []
    if s.has_tail_arrow: ends.append(right_x)
    if s.has_head_arrow: ends.append(left_x)
    return any(e < mid for e in ends)

def is_visually_black(rgb):
    """办公软件中视觉呈黑色：各通道均较暗（含纯黑000000与深灰404040等）。"""
    return max(rgb) <= 96

def round_rect_radius_cm(s):
    """roundRect 的圆角半径(cm)，与办公软件渲染一致：
    半径 = 圆角调整比例 × 较短边长。PowerPoint 默认调整值约 0.16667。"""
    adj = s.round_adj if s.round_adj is not None else 1/6
    short_side = min(s.abs_w, s.abs_h)
    return adj * short_side


# ─── 线宽工具 ────────────────────────────────────────────────────────────────
# 未显式设置线宽时，PowerPoint 默认按 1.0 磅渲染。
DEFAULT_LINE_WIDTH_PT = 1.0
def effective_line_width_pt(s):
    return s.line_width_pt if s.line_width_pt is not None else DEFAULT_LINE_WIDTH_PT
def lw_in(s, lo, hi):
    lw = effective_line_width_pt(s)
    return lo <= lw <= hi

# ─── 虚线工具 ────────────────────────────────────────────────────────────────
def is_dashed(s):
    if s.line_dash is None: return False
    return s.line_dash.lower() not in ('solid','')

def is_solid_line(s):
    if s.line_dash is None: return True
    return s.line_dash.lower() in ('solid','')

# ─── 字体工具 ────────────────────────────────────────────────────────────────
GOOD_FONTS = {'arial', 'calibri'}
# 主题字体（load_pptx 解析后写入）。办公软件在文本未显式指定字体、
# 或字体名为 +mn-lt / +mj-lt 时，按主题的次要/主要拉丁字体渲染。
THEME_FONTS = {'minor': '', 'major': ''}

def resolve_font(f: str) -> str:
    """把文本 run 的字体名解析为办公软件实际渲染的拉丁字体名。
    - 空字体：继承主题次要字体(+mn-lt)
    - +mn-lt / +mj-lt：主题字体占位符
    - 其余：原样返回"""
    f = (f or '').strip()
    low = f.lower()
    if not f or low in ('+mn-lt', '+mn-ea', '+mn-cs'):
        return (THEME_FONTS.get('minor') or '').strip()
    if low in ('+mj-lt', '+mj-ea', '+mj-cs'):
        return (THEME_FONTS.get('major') or '').strip()
    return f

def run_font_ok(tr):
    f = resolve_font(tr[1]).lower()
    if not f:  # 主题字体也为空时无法判定，宽松通过
        return True
    return f in GOOD_FONTS

def shape_font_ok(s):
    if not s.texts: return True
    return all(run_font_ok(t) for t in s.texts)

def all_fonts_ok(shapes):
    for s in shapes:
        if s.is_pic: continue
        if not shape_font_ok(s): return False
    return True

# ─── 重叠检测 ────────────────────────────────────────────────────────────────
def overlap_area(a, b):
    ox = max(0, min(a.right,b.right)-max(a.x,b.x))
    oy = max(0, min(a.bottom,b.bottom)-max(a.y,b.y))
    return ox*oy

# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
#  维度1 门槛检查
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
def check_dim1(filepath: str, slides: list[list['Shape']]) -> list[str]:
    failures: list[str] = []
    # D1-1: 交付文件为 .pptx 格式
    suf = Path(filepath).suffix.lower()
    if suf != '.pptx':
        failures.append('D1-1: 文件后缀不是 .pptx')

    # D1-2: 幻灯片数量必须为2
    if len(slides) != 2:
        failures.append(f'D1-2: 幻灯片数量为 {len(slides)}，要求2页')

    return failures

# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
#  维度2 规则检测器
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# 每个检测函数返回 (hit: bool, evidence: str)
# ─────────────────────────────────────────────────────────────────────────────
def chk_fonts(slides, **_):
    """+5: PPT中的可编辑文本皆为 Arial 或 Calibri 字体
    以办公软件实际渲染的字体为准：文本未显式指定字体时继承主题次要字体，
    +mn-lt/+mj-lt 解析为主题字体，再判断是否为 Arial/Calibri。"""
    bad = []
    for idx, shapes in enumerate(slides):
        for s in shapes:
            if s.is_pic: continue
            for tr in s.texts:
                if not tr[0].strip():  # 空白run不含可见文本，跳过
                    continue
                f = resolve_font(tr[1])
                if f and f.lower() not in GOOD_FONTS:
                    bad.append(f'Slide{idx+1}:"{s.name}" run "{tr[0][:15]}" font={f}')
    if not bad: return True, '全部可编辑文本字体为 Arial/Calibri（含继承主题字体）'
    return False, f'发现 {len(bad)} 处非Arial/Calibri字体: ' + '; '.join(bad[:3])

# --- 第1页 ---
def chk_p1_y_arrow(slides, **_):
    """+1: 第1页左侧纵轴箭头：位于距左5.0cm–5.7cm、距上0.8cm–18cm范围内，
    黑色竖向单箭头线，线宽1.5–2.5磅，箭头朝上，线条角度约90度。"""
    shapes = slides[0]
    tol = 0.2  # cm，容忍 EMU 取整误差
    for s in shapes:
        # 单箭头线
        if arrow_count(s) != 1: continue
        # 距左 5.0–5.7cm（竖线 x 恒定，取左边界）
        left = min(s.x, s.x + s.w)
        if not (5.0 - tol <= left <= 5.7 + tol): continue
        # 距上 0.8–18cm 范围内（线条整体纵向落在该带内）
        top = min(s.y, s.y + s.h)
        bottom = max(s.y, s.y + s.h)
        if not (0.8 - tol <= top and bottom <= 18.0 + tol): continue
        # 竖向 / 线条角度约90度
        if not (85 <= line_angle_from_horizontal(s) <= 95): continue
        # 黑色（线条颜色，办公软件视觉呈黑）
        rgb = shape_line_rgb(s) or shape_fill_rgb(s)
        if not (rgb and is_visually_black(rgb)): continue
        # 线宽 1.5–2.5 磅
        if not lw_in(s, 1.5, 2.5): continue
        # 箭头朝上
        if not arrow_points_up(s): continue
        return True, (f'找到纵轴箭头: {s.name} x={left:.2f} top={top:.2f} bottom={bottom:.2f} '
                      f'angle={line_angle_from_horizontal(s):.0f}° lw={s.line_width_pt}')
    return False, '未找到符合细则的黑色竖向单箭头(距左5.0-5.7cm/距上0.8-18cm/线宽1.5-2.5磅/朝上/约90°)'

def chk_p1_y_label(slides, **_):
    """+1: 第1页左侧纵轴标题文本：位于距左4cm–5.2cm、距上6.5cm–12.0cm范围内，
    文本为"System Conditions"，处于1列，文字旋转90度竖向排布，
    字体为Arial或Calibri，字号13–15磅，加粗，颜色为黑色。"""
    shapes = slides[0]
    tol = 0.2  # cm，容忍 EMU→cm 取整与屏幕测量误差
    failures = []
    for s in shapes:
        # 文本为"System Conditions"
        if s.full_text().strip() != 'System Conditions':
            continue

        # 位置：以办公软件所见的旋转后可见包围盒为准
        left, top, right, bottom = rendered_bbox(s)
        text_runs = [tr for tr in s.texts if tr[0].strip()]

        # 处于1列：以办公软件实际渲染为准——文本在文本框内自动换行后
        # 的可见行数(竖排时即列数)为 1，才算单列。
        visual_lines = estimate_wrapped_lines(s)
        one_column = (visual_lines == 1)
        # 文字旋转90度竖向排布（90或270均为竖向）
        vertical = near_angle(s, 90, tol=10) or near_angle(s, 270, tol=10)
        # 距左4–5.2cm、距上6.5–12.0cm
        in_pos = (4.0 - tol <= left <= 5.2 + tol) and (6.5 - tol <= top <= 12.0 + tol)
        # 字体 Arial/Calibri（含继承主题字体）
        font_ok = bool(text_runs) and all(run_font_ok(tr) for tr in text_runs)
        # 字号13–15磅
        size_ok = bool(text_runs) and all(13 <= tr[2] <= 15 for tr in text_runs)
        # 加粗
        bold_ok = bool(text_runs) and all(tr[3] for tr in text_runs)
        # 颜色为黑色（办公软件视觉呈黑，含纯黑与深灰）
        color_ok = bool(text_runs) and all(
            text_run_rgb(tr) is None or is_visually_black(text_run_rgb(tr)) for tr in text_runs)

        if all([one_column, vertical, in_pos, font_ok, size_ok, bold_ok, color_ok]):
            return True, (f'找到纵轴标题: {s.name} bbox=({left:.2f},{top:.2f})-({right:.2f},{bottom:.2f}) '
                          f'rot={s.rotation_deg():.0f}°')

        reasons = []
        if not in_pos:
            reasons.append(f'位置left={left:.2f}/top={top:.2f}不在距左4-5.2/距上6.5-12范围内')
        if not one_column:
            reasons.append(f'不是1列文本(视觉行数={visual_lines})')
        if not vertical:
            reasons.append(f'旋转角度{s.rotation_deg():.0f}°不是竖向')
        if not font_ok:
            reasons.append('字体不是Arial/Calibri')
        if not size_ok:
            reasons.append('字号不在13-15磅')
        if not bold_ok:
            reasons.append('未加粗')
        if not color_ok:
            colors = [tr[4] for tr in text_runs]
            reasons.append(f'颜色不是黑色: {colors}')
        failures.append(f'{s.name}: ' + '；'.join(reasons))
    if failures:
        return False, '找到"System Conditions"但不符合要求：' + ' | '.join(failures[:2])
    return False, '未找到符合要求的"System Conditions"文本'

def chk_p1_x_arrow(slides, **_):
    """+1: 第1页底部横轴箭头：位于距左5.0cm–29cm、距上18.0cm–18.5cm范围内，
    黑色横向单箭头线，线宽1.5–2.5磅，箭头朝右，线条角度约0度。"""
    shapes = slides[0]
    tol = 0.2  # cm，容忍 EMU→cm 取整误差
    for s in shapes:
        # 单箭头线
        if arrow_count(s) != 1: continue
        # 距左5.0–29cm（横线整体横向落在该带内）
        left = min(s.x, s.x + s.w)
        right = max(s.x, s.x + s.w)
        if not (5.0 - tol <= left and right <= 29.0 + tol): continue
        # 距上18.0–18.5cm（横线 y 恒定）
        top = min(s.y, s.y + s.h)
        bottom = max(s.y, s.y + s.h)
        if not (18.0 - tol <= top and bottom <= 18.5 + tol): continue
        # 横向 / 线条角度约0度
        if not (0 <= line_angle_from_horizontal(s) <= 5): continue
        # 黑色（办公软件视觉呈黑，含纯黑与深灰）
        rgb = shape_line_rgb(s) or shape_fill_rgb(s)
        if not (rgb and is_visually_black(rgb)): continue
        # 线宽1.5–2.5磅
        if not lw_in(s, 1.5, 2.5): continue
        # 箭头朝右
        if not arrow_points_right(s): continue
        return True, (f'找到横轴箭头: {s.name} left={left:.2f} right={right:.2f} top={top:.2f} '
                      f'angle={line_angle_from_horizontal(s):.0f}° lw={s.line_width_pt}')
    return False, '未找到符合细则的黑色横向单箭头(距左5.0-29cm/距上18.0-18.5cm/线宽1.5-2.5磅/朝右/约0°)'

def chk_p1_x_label(slides, **_):
    """+1: 第1页底部横轴标题文本：位于距左15.0cm–19.0cm、距上17.9cm–18.8cm范围内，
    文本为"Time (T)"，字体为Arial或Calibri，字号13–15磅，加粗，颜色为黑色，水平居中。"""
    shapes = slides[0]
    tol = 0.2  # cm，容忍 EMU→cm 取整与屏幕测量误差
    failures = []
    for s in shapes:
        # 文本为"Time (T)"
        if s.full_text().strip() != 'Time (T)':
            continue

        left, top, right, bottom = rendered_bbox(s)
        text_runs = [tr for tr in s.texts if tr[0].strip()]
        reasons = []

        # 距左15.0–19.0cm、距上17.9–18.8cm（以办公软件所见的文本框位置起点为准）
        if not ((15.0 - tol <= left <= 19.0 + tol) and (17.9 - tol <= top <= 18.8 + tol)):
            reasons.append(f'位置left={left:.2f}/top={top:.2f}不在距左15-19/距上17.9-18.8范围内')
        # 字体为Arial或Calibri
        if not (text_runs and all(run_font_ok(tr) for tr in text_runs)):
            reasons.append('字体不是Arial/Calibri')
        # 字号13–15磅
        if not (text_runs and all(13 <= tr[2] <= 15 for tr in text_runs)):
            reasons.append(f'字号不在13-15磅: {[tr[2] for tr in text_runs]}')
        # 加粗
        if not (text_runs and all(tr[3] for tr in text_runs)):
            reasons.append('未加粗')
        # 颜色为黑色（办公软件视觉呈黑，含纯黑与深灰）
        if not (text_runs and all(
                text_run_rgb(tr) is None or is_visually_black(text_run_rgb(tr)) for tr in text_runs)):
            reasons.append(f'颜色不是黑色: {[tr[4] for tr in text_runs]}')
        # 水平居中（段落对齐为居中）
        if not is_center_aligned(s):
            reasons.append('文本非水平居中')

        if not reasons:
            return True, (f'找到横轴标题: {s.name} "{s.full_text()}" '
                          f'left={left:.2f} top={top:.2f}')
        failures.append(f'{s.name}: ' + '；'.join(reasons))
    if failures:
        return False, '找到"Time (T)"但不符合要求：' + ' | '.join(failures[:2])
    return False, '未找到符合要求的"Time (T)"文本'

def chk_p1_top_rrect(slides, **_):
    """+1: 第1页顶部浅蓝色圆角矩形：位于距左7.0cm–28.5cm、距上0.25cm–4.4cm范围内，
    宽20.5cm–22cm，高3.8cm–4.4cm，填充为浅蓝色，边线为蓝色单实线，线宽1–1.5磅，
    圆角半径约0.2cm–0.5cm。"""
    shapes = slides[0]
    failures = []
    for s in shapes_in(shapes, 7.0, 28.5, 0.25, 4.4):
        # 圆角矩形
        if s.shape_type != 'sp' or s.geom != 'roundRect':
            continue
        reasons = []
        # 位于距左7.0–28.5cm、距上0.25–4.4cm范围内（旋转后可见包围盒落在该区域）
        if not bbox_in_box(s, 7.0, 28.5, 0.25, 4.4):
            left, top, right, bottom = rendered_bbox(s)
            reasons.append(f'位置bbox=({left:.2f},{top:.2f})-({right:.2f},{bottom:.2f})不在范围内')
        # 宽20.5–22cm
        if not in_range(s.abs_w, 20.5, 22.0):
            reasons.append(f'宽度{s.abs_w:.2f}cm不在20.5-22cm')
        # 高3.8–4.4cm
        if not in_range(s.abs_h, 3.8, 4.4):
            reasons.append(f'高度{s.abs_h:.2f}cm不在3.8-4.4cm')
        # 填充为浅蓝色
        fill = shape_fill_rgb(s)
        if not (fill and is_light_blue_fill(fill)):
            reasons.append(f'填充不是浅蓝色: {s.fill_color}')
        # 边线为蓝色单实线
        line = shape_line_rgb(s)
        if not (line and is_blue(line) and is_solid_line(s)):
            reasons.append(f'边线不是蓝色单实线: color={s.line_color}, dash={s.line_dash}')
        # 线宽1–1.5磅
        if not lw_in(s, 1.0, 1.5):
            reasons.append(f'线宽{s.line_width_pt}不在1-1.5磅')
        # 圆角半径约0.2–0.5cm（办公软件按 adj 比例 × 较短边渲染）
        radius = round_rect_radius_cm(s)
        if not (0.2 <= radius <= 0.5):
            reasons.append(f'圆角半径{radius:.2f}cm不在0.2-0.5cm')
        if not reasons:
            return True, (f'找到顶部浅蓝圆角矩形: {s.name} w={s.abs_w:.1f} h={s.abs_h:.1f} '
                          f'圆角={radius:.2f}cm')
        failures.append(f'{s.name}: ' + '；'.join(reasons))
    if failures:
        return False, '找到圆角矩形但不符合要求：' + ' | '.join(failures[:2])
    return False, '未找到符合要求的顶部浅蓝色圆角矩形'

def chk_p1_top_title(slides, **_):
    """+1: 第1页顶部圆角矩形标题文本：位于距左14.0cm–21.5cm、距上0.55cm–1.3cm范围内，
    文本为"Macro Environment"，字体为Arial或Calibri，字号16–18磅，文本横向一行排列，
    加粗，颜色为深蓝色，水平居中。"""
    shapes = slides[0]
    tol = 0.2  # cm，容忍 EMU→cm 取整与屏幕测量误差
    failures = []
    for s in shapes:
        # 文本为"Macro Environment"
        if s.full_text().strip() != 'Macro Environment':
            continue

        left, top, right, bottom = rendered_bbox(s)
        text_runs = [tr for tr in s.texts if tr[0].strip()]
        reasons = []

        # 距左14.0–21.5cm、距上0.55–1.3cm
        if not ((14.0 - tol <= left <= 21.5 + tol) and (0.55 - tol <= top <= 1.3 + tol)):
            reasons.append(f'位置left={left:.2f}/top={top:.2f}不在距左14-21.5/距上0.55-1.3范围内')
        # 字体为Arial或Calibri
        if not (text_runs and all(run_font_ok(tr) for tr in text_runs)):
            reasons.append('字体不是Arial/Calibri')
        # 字号16–18磅
        if not (text_runs and all(16 <= tr[2] <= 18 for tr in text_runs)):
            reasons.append(f'字号不在16-18磅: {[tr[2] for tr in text_runs]}')
        # 文本横向排列（形状未旋转，约0度）
        if not near_angle(s, 0, 5):
            reasons.append(f'文本不是横向排列(rot={s.rotation_deg():.0f}°)')
        # 一行排列：以办公软件实际渲染为准，文本框内自动换行后仅 1 行
        lines = estimate_wrapped_lines(s)
        if lines != 1:
            reasons.append(f'文本不是一行排列(视觉行数={lines})')
        # 加粗
        if not (text_runs and all(tr[3] for tr in text_runs)):
            reasons.append('未加粗')
        # 颜色为深蓝色
        if not (text_runs and all(
                text_run_rgb(tr) is not None and is_strict_dark_blue(text_run_rgb(tr)) for tr in text_runs)):
            reasons.append(f'颜色不是深蓝色: {[tr[4] for tr in text_runs]}')
        # 水平居中（段落对齐为居中）
        if not is_center_aligned(s):
            reasons.append('文本非水平居中')

        if not reasons:
            return True, f'找到"Macro Environment": {s.name} left={left:.2f} top={top:.2f}'
        failures.append(f'{s.name}: ' + '；'.join(reasons))
    if failures:
        return False, '找到"Macro Environment"但不符合要求：' + ' | '.join(failures[:2])
    return False, '未找到符合要求的"Macro Environment"文本'

def chk_p1_six_texts(slides, **_):
    """+3: 第1页顶部圆角矩形内部六组文本：位于距左8.0cm–27.3cm、距上1.54cm–2.4cm范围内，
    从左到右出现"Technological advances""Policy & reforms""Funding landscape"
    "Demographic shifts""Global health priorities""Socioeconomic trends"，
    一个单词占一行，每个词组之间的间距约为1.4cm-2cm，字体为Arial或Calibri，
    字号7–9磅、加粗，颜色为黑色。"""
    shapes = slides[0]
    tol = 0.2  # cm，容忍 EMU→cm 取整与屏幕测量误差
    required = [
        'Technological advances',
        'Policy & reforms',
        'Funding landscape',
        'Demographic shifts',
        'Global health priorities',
        'Socioeconomic trends',
    ]
    failures = []
    matches = []
    for phrase in required:
        cands = [s for s in shapes if s.full_text().strip() == phrase]
        if not cands:
            failures.append(f'缺少"{phrase}"')
            continue
        s = cands[0]
        matches.append(s)
        reasons = []
        # 位于距左8.0–27.3cm、距上1.54–2.4cm范围内
        left, top, right, bottom = rendered_bbox(s)
        if not (8.0 - tol <= left and right <= 27.3 + tol
                and 1.54 - tol <= top and bottom <= 2.4 + tol):
            reasons.append(f'位置bbox=({left:.2f},{top:.2f})-({right:.2f},{bottom:.2f})不在范围内')
        # 一个单词占一行（文本分多行竖直堆叠排布）
        lines = [p for p in s.paragraph_texts if p.strip()]
        if len(lines) < 2:
            reasons.append(f'未按一个单词占一行分行排布: {s.paragraph_texts}')
        # 字体Arial/Calibri、字号7–9磅、加粗、颜色黑色
        runs = non_empty_runs(s)
        if not (runs and all(run_font_ok(tr) for tr in runs)):
            reasons.append('字体不是Arial/Calibri')
        if not (runs and all(7 <= tr[2] <= 9 for tr in runs)):
            reasons.append(f'字号不在7-9磅: {[tr[2] for tr in runs]}')
        if not (runs and all(tr[3] for tr in runs)):
            reasons.append('未加粗')
        if not (runs and all(
                text_run_rgb(tr) is None or is_visually_black(text_run_rgb(tr)) for tr in runs)):
            reasons.append(f'颜色不是黑色: {[tr[4] for tr in runs]}')
        if reasons:
            failures.append(f'{phrase}: ' + '；'.join(reasons))
    if len(matches) == 6:
        # 从左到右出现（按 required 顺序，x 递增）
        if not all(matches[i].x < matches[i+1].x for i in range(5)):
            failures.append('六组文本未从左到右排列')
        # 每个词组之间的间距约为1.4-2cm（相邻词组可见边缘间距）
        gaps = [rendered_bbox(matches[i+1])[0] - rendered_bbox(matches[i])[2] for i in range(5)]
        if not all(1.4 - tol <= gap <= 2.0 + tol for gap in gaps):
            failures.append(f'词组间距不在1.4-2cm: {[round(g, 2) for g in gaps]}')
    if not failures:
        return True, '六组顶部环境文本均符合要求'
    return False, '六组顶部环境文本不符合要求：' + ' | '.join(failures[:4])

def chk_p1_curves(slides, **_):
    """+3: 第1页顶部矩形内3条蓝色曲线，逐条验证指定线型序列
    位置距左7.3–28cm、距上2.3–4.0cm。要求3条从左向右延伸的蓝色曲线，
    右端带蓝色箭头，线宽0.75–1.25磅；三条曲线的线型序列（任意匹配一条）：
      · 实线-点划线-虚线
      · 虚线-点划线-虚线-点划线
      · 虚线
    """
    shapes = slides[0]
    cands = [s for s in shapes_in(shapes, 7.3, 28.0, 2.3, 4.0)
             if s.shape_type in ('sp', 'cxnSp')
             and shape_line_rgb(s) and is_blue(shape_line_rgb(s))
             and lw_in(s, 0.75, 1.25)]
    if not cands:
        return False, '未找到符合条件的蓝色曲线段(色/线宽/位置)'

    def dash_kind(seg):
        d = (seg.line_dash or 'solid').lower()
        if d in ('solid', ''):
            return 'solid'
        if 'dashdot' in d:
            return 'dashDot'
        return 'dash'

    def y_center(seg):
        _, t, _, b = rendered_bbox(seg)
        return (t + b) / 2

    # 按 y 中心聚类为3行 —— 每一行代表一条曲线
    sorted_by_y = sorted(cands, key=y_center)
    rows: list[list] = []
    for seg in sorted_by_y:
        if rows and abs(y_center(rows[-1][-1]) - y_center(seg)) <= 0.4:
            rows[-1].append(seg)
        else:
            rows.append([seg])

    if len(rows) != 3:
        return False, (f'蓝色曲线未聚成3条，实际得到{len(rows)}条 '
                       f'(每行段数: {[len(r) for r in rows]})')

    # 每条曲线按 x 从左到右排序，得到线型序列
    per_curve = []
    for row in rows:
        row.sort(key=lambda s: rendered_bbox(s)[0])
        seq = [dash_kind(s) for s in row]
        left = rendered_bbox(row[0])[0]
        right = rendered_bbox(row[-1])[2]
        per_curve.append((seq, left, right, row))

    expected = [
        ['solid', 'dashDot', 'dash'],           # 实线-点划线-虚线
        ['dash', 'dashDot', 'dash', 'dashDot'], # 虚线-点划线-虚线-点划线
        ['dash'],                                # 虚线
    ]

    # 允许三条曲线在页面上的顺序任意，逐一匹配
    used_row = [False, False, False]
    matched_map = {}
    for exp_idx, exp_seq in enumerate(expected):
        found = -1
        for row_idx, (seq, *_ignore) in enumerate(per_curve):
            if used_row[row_idx]:
                continue
            if seq == exp_seq:
                found = row_idx
                break
        if found < 0:
            seqs = [s for s, *_ in per_curve]
            return False, (f'未匹配到期望线型序列 {exp_seq}；'
                           f'当前3条曲线线型序列: {seqs}')
        used_row[found] = True
        matched_map[exp_idx] = found

    # 每条曲线从左向右延伸 + 右端带箭头
    reasons = []
    for exp_idx in range(3):
        row_idx = matched_map[exp_idx]
        seq, left, right, row = per_curve[row_idx]
        if left > 13.0:
            reasons.append(f'{"-".join(seq)} 左端{left:.2f}cm 未从左侧起点延伸')
        if right < 27.0:
            reasons.append(f'{"-".join(seq)} 右端{right:.2f}cm 未延伸至右侧')
        if not has_any_arrow(row[-1]):
            reasons.append(f'{"-".join(seq)} 右端未带箭头')

    if reasons:
        return False, '顶部曲线组不符合要求：' + '；'.join(reasons)
    return True, ('三条曲线线型序列匹配成功: '
                  + '; '.join('-'.join(seq) for seq, *_ in per_curve))

def chk_p1_big_ellipse(slides, **_):
    """+1: 第1页中部浅绿色大椭圆：位于距左9.2cm–24.9cm、距上4.8cm–10.2cm范围内，
    宽15cm–15.5cm，高4.8cm–5.5cm，填充为浅绿色，边线为绿色单实线，线宽1–1.5磅。"""
    shapes = slides[0]
    tol = 0.2  # cm，容忍 EMU→cm 取整与屏幕测量误差
    failures = []
    for s in shapes:
        # 椭圆
        if s.geom != 'ellipse':
            continue
        reasons = []
        # 位于距左9.2–24.9cm、距上4.8–10.2cm范围内（旋转后可见包围盒落在该区域）
        if not bbox_in_box(s, 9.2, 24.9, 4.8, 10.2):
            left, top, right, bottom = rendered_bbox(s)
            reasons.append(f'位置bbox=({left:.2f},{top:.2f})-({right:.2f},{bottom:.2f})不在范围内')
        # 宽15–15.5cm
        if not in_range(s.abs_w, 15.0 - tol, 15.5 + tol):
            reasons.append(f'宽度{s.abs_w:.2f}cm不在15-15.5cm')
        # 高4.8–5.5cm
        if not in_range(s.abs_h, 4.8 - tol, 5.5 + tol):
            reasons.append(f'高度{s.abs_h:.2f}cm不在4.8-5.5cm')
        # 填充为浅绿色
        fill = shape_fill_rgb(s)
        if not (fill and is_pale_green_fill(fill)):
            reasons.append(f'填充不是浅绿色: {s.fill_color}')
        # 边线为绿色单实线
        line = shape_line_rgb(s)
        if not (line and is_green(line) and is_solid_line(s)):
            reasons.append(f'边线不是绿色单实线: color={s.line_color}, dash={s.line_dash}')
        # 线宽1–1.5磅
        if not lw_in(s, 1.0, 1.5):
            reasons.append(f'线宽{s.line_width_pt}不在1-1.5磅')
        if not reasons:
            return True, f'找到中部浅绿色大椭圆: {s.name} w={s.abs_w:.2f} h={s.abs_h:.2f}'
        failures.append(f'{s.name}: ' + '；'.join(reasons))
    if failures:
        return False, '找到椭圆但不符合要求：' + ' | '.join(failures[:2])
    return False, '未找到符合要求的中部浅绿色大椭圆'

def chk_p1_big_ellipse_title(slides, **_):
    """+1: 第1页中部大椭圆标题文本：位于距左13.5cm–21.5cm、距上5.4cm–6.2cm范围内，
    文本为"Local Health Innovation Context"，字体为Arial或Calibri，内容横向排列，
    放置在一行，字号13–15磅，加粗，颜色为深绿色，水平居中。"""
    shapes = slides[0]
    tol = 0.2  # cm，容忍 EMU→cm 取整与屏幕测量误差
    failures = []
    for s in shapes:
        # 文本为"Local Health Innovation Context"
        if s.full_text().strip() != 'Local Health Innovation Context':
            continue

        left, top, right, bottom = rendered_bbox(s)
        runs = non_empty_runs(s)
        reasons = []

        # 距左13.5–21.5cm、距上5.4–6.2cm
        if not ((13.5 - tol <= left <= 21.5 + tol) and (5.4 - tol <= top <= 6.2 + tol)):
            reasons.append(f'位置left={left:.2f}/top={top:.2f}不在距左13.5-21.5/距上5.4-6.2范围内')
        # 字体为Arial或Calibri
        if not (runs and all(run_font_ok(tr) for tr in runs)):
            reasons.append('字体不是Arial/Calibri')
        # 内容横向排列（形状未旋转，约0度）
        if not near_angle(s, 0, 5):
            reasons.append(f'内容不是横向排列(rot={s.rotation_deg():.0f}°)')
        # 放置在一行（单个非空段落且无换行符）
        paras = [p for p in s.paragraph_texts if p.strip()]
        if len(paras) != 1 or any('\n' in tr[0] or '\r' in tr[0] for tr in runs):
            reasons.append(f'不是一行文本: {s.paragraph_texts}')
        # 字号13–15磅
        if not (runs and all(13 <= tr[2] <= 15 for tr in runs)):
            reasons.append(f'字号不在13-15磅: {[tr[2] for tr in runs]}')
        # 加粗
        if not (runs and all(tr[3] for tr in runs)):
            reasons.append('未加粗')
        # 颜色为深绿色
        if not (runs and all(
                text_run_rgb(tr) is not None and is_dark_green(text_run_rgb(tr)) for tr in runs)):
            reasons.append(f'颜色不是深绿色: {[tr[4] for tr in runs]}')
        # 水平居中（段落对齐为居中）
        if not is_center_aligned(s):
            reasons.append('文本非水平居中')

        if not reasons:
            return True, f'找到大椭圆标题: {s.name} left={left:.2f} top={top:.2f}'
        failures.append(f'{s.name}: ' + '；'.join(reasons))
    if failures:
        return False, '找到"Local Health Innovation Context"但不符合要求：' + ' | '.join(failures[:2])
    return False, '未找到符合要求的"Local Health Innovation Context"'

def chk_p1_small_ellipses_top(slides, **_):
    """+3: 第1页中部大椭圆上排小椭圆组：位于距左10.2cm–23.8cm、距上6.4cm–7.9cm范围内，
    包含"clinical infrastructure""data resources""regulatory support"三个浅色小椭圆，
    三个小椭圆从左到右排列，单个宽2.6cm–3.2cm，高0.9cm–1.4cm，
    填充为白色或极浅绿色，边线为绿色单实线，线宽0.75–1.25磅，
    字体为Arial或Calibri，文本字号8–10磅，颜色为黑色。"""
    return _p1_small_ellipse_group(
        slides[0], 10.2, 23.8, 6.4, 7.9,
        ['clinical infrastructure', 'data resources', 'regulatory support'], '上排')

def _p1_small_ellipse_group(shapes, xl, xr, yt, yb, phrases, tag_name):
    tol = 0.2  # cm，容忍 EMU→cm 取整与屏幕测量误差
    failures = []
    # 三个浅色小椭圆（几何尺寸/填充/边线/线宽符合细则）
    ellipses = []
    for s in shapes:
        if s.geom != 'ellipse': continue
        if not shape_in_box(s, xl, xr, yt, yb): continue
        if not in_range(s.abs_w, 2.6 - tol, 3.2 + tol): continue
        if not in_range(s.abs_h, 0.9 - tol, 1.4 + tol): continue
        fill = shape_fill_rgb(s)
        # 填充为白色或极浅绿色
        if not (fill and (is_white(fill) or is_pale_green_fill(fill))): continue
        line = shape_line_rgb(s)
        # 边线为绿色单实线，线宽0.75-1.25磅
        if not (line and is_green(line) and is_solid_line(s)): continue
        if not lw_in(s, 0.75, 1.25): continue
        ellipses.append(s)
    if len(ellipses) < 3:
        failures.append(f'符合几何/填充/边线要求的小椭圆不足3个: {len(ellipses)}')
    else:
        # 从左到右排列
        ordered = sorted(ellipses[:3], key=lambda s: s.x)
        if not all(ordered[i].x < ordered[i+1].x for i in range(2)):
            failures.append('三个小椭圆未从左到右排列')

    # 三个标签文本（内容 + 字体/字号/颜色）
    for phrase in phrases:
        cands = [s for s in shapes if s.full_text().strip().lower() == phrase.lower()]
        if not cands:
            failures.append(f'缺少"{phrase}"文本')
            continue
        s = cands[0]
        # 文本须位于该椭圆组区域内
        left, top, right, bottom = rendered_bbox(s)
        if not (xl - tol <= left and right <= xr + tol and yt - tol <= top and bottom <= yb + tol):
            failures.append(f'"{phrase}"不在椭圆组区域内')
        runs = non_empty_runs(s)
        if not (runs and all(run_font_ok(tr) for tr in runs)):
            failures.append(f'"{phrase}"字体不是Arial/Calibri')
        if not (runs and all(8 <= tr[2] <= 10 for tr in runs)):
            failures.append(f'"{phrase}"字号不在8-10磅: {[tr[2] for tr in runs]}')
        if not (runs and all(
                text_run_rgb(tr) is None or is_visually_black(text_run_rgb(tr)) for tr in runs)):
            failures.append(f'"{phrase}"颜色不是黑色: {[tr[4] for tr in runs]}')

    if not failures:
        return True, f'{tag_name}小椭圆组三项均符合要求'
    return False, f'{tag_name}小椭圆组不符合要求：' + ' | '.join(failures[:4])

def chk_p1_small_ellipses_bot(slides, **_):
    """+3: 第1页中部大椭圆下排小椭圆组：位于距左12.0cm–22.4cm、距上8.0cm–9.5cm范围内，
    包含"talent base""public trust""research networks"三个浅色小椭圆，
    三个小椭圆从左到右排列，宽2.6cm–3.2cm，高0.9cm–1.4cm，
    填充为白色或极浅绿色，边线为绿色单实线，线宽0.75–1.25磅，
    字体为Arial或Calibri，文本字号8–10磅，颜色为黑色。"""
    return _p1_small_ellipse_group(
        slides[0], 12.0, 22.4, 8.0, 9.5,
        ['talent base', 'public trust', 'research networks'], '下排')

def chk_p1_blue_arrows_top(slides, **_):
    """+3: 第1页顶部圆角矩形到中部大椭圆
    左侧向下箭头：位于距左11.5cm–13.0cm、距上4.1cm–5.4cm范围内，形状为蓝色粗箭头，
    箭头朝下，宽0.5cm–1cm，高1cm–1.4cm。
    右侧向下箭头：位于距左22.0cm–23.5cm、距上4.1cm–5.4cm范围内，形状为蓝色粗箭头，
    箭头朝下，宽0.5cm–1cm，高1cm–1.4cm。"""
    shapes = slides[0]
    tol = 0.2  # cm，容忍 EMU→cm 取整与屏幕测量误差

    def find_blue_down_arrow(xl, xr, yt, yb):
        for s in shapes:
            # 位置：形状落在指定范围内
            left, top, right, bottom = rendered_bbox(s)
            if not (xl - tol <= left and right <= xr + tol
                    and yt - tol <= top and bottom <= yb + tol):
                continue
            # 一体箭头(预设箭头几何 或 线条+端点箭头)，且视觉朝下
            if not is_arrow_shape(s): continue
            if not arrow_points_down(s): continue
            # 蓝色（填充或线条呈蓝，办公软件视觉为蓝）
            rgb = shape_fill_rgb(s) or shape_line_rgb(s)
            if not (rgb and is_blue(rgb)): continue
            # 宽0.5–1cm、高1–1.4cm
            if not in_range(s.abs_w, 0.5 - tol, 1.0 + tol): continue
            if not in_range(s.abs_h, 1.0 - tol, 1.4 + tol): continue
            return s
        return None

    left = find_blue_down_arrow(11.5, 13.0, 4.1, 5.4)
    right = find_blue_down_arrow(22.0, 23.5, 4.1, 5.4)
    if left and right:
        return True, f'找到左右蓝色向下粗箭头: {left.name}, {right.name}'
    missing = []
    if not left: missing.append('左侧(11.5-13.0,4.1-5.4)')
    if not right: missing.append('右侧(22.0-23.5,4.1-5.4)')
    return False, '未找到符合细则的蓝色向下粗箭头(宽0.5-1cm/高1-1.4cm/朝下): ' + '、'.join(missing)

def chk_p1_green_arrows_mid(slides, **_):
    """+3: 第1页中部大椭圆到下方蓝色圆角矩形边框间有两个绿色竖向向下箭头，
    左侧向下箭头位于距左11.5cm–13.0cm、距上9.2cm–11cm范围内，
    右侧向下箭头位于距左21.3cm–22.5cm、距上9.2cm–11cm范围内，
    宽0.5cm–0.9cm，高1cm–1.3cm。"""
    shapes = slides[0]
    tol = 0.2  # cm，容忍 EMU→cm 取整与屏幕测量误差

    def find_green_down_arrow(xl, xr, yt, yb):
        for s in shapes:
            # 位置：形状落在指定范围内
            left, top, right, bottom = rendered_bbox(s)
            if not (xl - tol <= left and right <= xr + tol
                    and yt - tol <= top and bottom <= yb + tol):
                continue
            # 一体箭头(预设箭头几何 或 线条+端点箭头)，且视觉朝下
            if not is_arrow_shape(s): continue
            if not arrow_points_down(s): continue
            # 绿色（填充或线条呈绿，办公软件视觉为绿）
            rgb = shape_fill_rgb(s) or shape_line_rgb(s)
            if not (rgb and is_green(rgb)): continue
            # 宽0.5–0.9cm、高1–1.3cm
            if not in_range(s.abs_w, 0.5 - tol, 0.9 + tol): continue
            if not in_range(s.abs_h, 1.0 - tol, 1.3 + tol): continue
            return s
        return None

    left = find_green_down_arrow(11.5, 13.0, 9.2, 11.0)
    right = find_green_down_arrow(21.3, 22.5, 9.2, 11.0)
    if left and right:
        return True, f'找到左右绿色竖向向下箭头: {left.name}, {right.name}'
    missing = []
    if not left: missing.append('左侧(11.5-13.0,9.2-11)')
    if not right: missing.append('右侧(21.3-22.5,9.2-11)')
    return False, '未找到符合细则的绿色竖向向下箭头(宽0.5-0.9cm/高1-1.3cm): ' + '、'.join(missing)

def chk_p1_arena_border(slides, **_):
    """+1: 第1页下方蓝色边框圆角矩形外框：位于距左8.8cm–25.4cm、距上10.3cm–17.3cm范围内，
    宽16cm–16.5cm，高6.8cm–7.3cm，填充为白色或极浅色，边线为蓝色单实线，
    线宽1.2–1.8磅，圆角半径约0.2cm–0.5cm。"""
    shapes = slides[0]
    tol = 0.2  # cm，容忍 EMU→cm 取整与屏幕测量误差
    failures = []
    for s in shapes:
        # 圆角矩形
        if s.geom != 'roundRect':
            continue
        # 只考察落在指定区域内的圆角矩形
        if not shape_in_box(s, 8.8, 25.4, 10.3, 17.3):
            continue
        reasons = []
        # 位于距左8.8–25.4cm、距上10.3–17.3cm范围内
        if not bbox_in_box(s, 8.8, 25.4, 10.3, 17.3):
            left, top, right, bottom = rendered_bbox(s)
            reasons.append(f'位置bbox=({left:.2f},{top:.2f})-({right:.2f},{bottom:.2f})不在范围内')
        # 宽16–16.5cm
        if not in_range(s.abs_w, 16.0 - tol, 16.5 + tol):
            reasons.append(f'宽度{s.abs_w:.2f}cm不在16-16.5cm')
        # 高6.8–7.3cm
        if not in_range(s.abs_h, 6.8 - tol, 7.3 + tol):
            reasons.append(f'高度{s.abs_h:.2f}cm不在6.8-7.3cm')
        # 填充为白色或极浅色
        fill = shape_fill_rgb(s)
        if not (fill and (is_white(fill) or is_pale_any(fill))):
            reasons.append(f'填充不是白色或极浅色: {s.fill_color}')
        # 边线为蓝色单实线
        line = shape_line_rgb(s)
        if not (line and is_blue(line) and is_solid_line(s)):
            reasons.append(f'边线不是蓝色单实线: color={s.line_color}, dash={s.line_dash}')
        # 线宽1.2–1.8磅
        if not lw_in(s, 1.2, 1.8):
            reasons.append(f'线宽{s.line_width_pt}不在1.2-1.8磅')
        # 圆角半径约0.2–0.5cm
        radius = round_rect_radius_cm(s)
        if not (0.2 <= radius <= 0.5):
            reasons.append(f'圆角半径{radius:.2f}cm不在0.2-0.5cm')
        if not reasons:
            return True, (f'找到arena蓝色边框圆角矩形: {s.name} w={s.abs_w:.2f} h={s.abs_h:.2f} '
                          f'圆角={radius:.2f}cm')
        failures.append(f'{s.name}: ' + '；'.join(reasons))
    if failures:
        return False, '找到圆角矩形但不符合要求：' + ' | '.join(failures[:2])
    return False, '未找到符合要求的下方蓝色边框圆角矩形'

def chk_p1_arena_title_bar(slides, **_):
    """+1: 第1页下方蓝色边框圆角矩形外框顶部深蓝标题条：位于距左13.8cm–20.3cm、
    距上10.1cm–11.2cm范围内，形状为深蓝色圆角矩形，宽5.5cm–6.5cm，高0.6cm–1cm，
    内部出现"Innovation Dynamics Arena"，文字为白色，字体为Arial或Calibri，字号10–12磅，加粗。"""
    shapes = slides[0]
    tol = 0.2  # cm，容忍 EMU→cm 取整与屏幕测量误差
    failures = []
    for s in shapes:
        # 形状为圆角矩形
        if s.geom != 'roundRect':
            continue
        left, top, right, bottom = rendered_bbox(s)
        # 位于距左13.8–20.3cm、距上10.1–11.2cm范围内（以标题条左上角为准）
        if not ((13.8 - tol <= left <= 20.3 + tol) and (10.1 - tol <= top <= 11.2 + tol)):
            continue
        reasons = []
        # 深蓝色
        fill = shape_fill_rgb(s)
        if not (fill and is_dark_blue(fill)):
            reasons.append(f'填充不是深蓝色: {s.fill_color}')
        # 宽5.5–6.5cm
        if not in_range(s.abs_w, 5.5 - tol, 6.5 + tol):
            reasons.append(f'宽度{s.abs_w:.2f}cm不在5.5-6.5cm')
        # 高0.6–1cm
        if not in_range(s.abs_h, 0.6 - tol, 1.0 + tol):
            reasons.append(f'高度{s.abs_h:.2f}cm不在0.6-1cm')
        # 内部出现"Innovation Dynamics Arena"文本（文本位于标题条范围内）
        label = None
        for t in shapes:
            if t.full_text().strip() != 'Innovation Dynamics Arena':
                continue
            tl, tt, tr_, tb = rendered_bbox(t)
            cx, cy = (tl + tr_) / 2, (tt + tb) / 2
            if left - tol <= cx <= right + tol and top - tol <= cy <= bottom + tol:
                label = t
                break
        if label is None:
            reasons.append('标题条内未出现"Innovation Dynamics Arena"文本')
        else:
            runs = non_empty_runs(label)
            # 文字为白色
            if not (runs and all(
                    text_run_rgb(tr) is not None and is_white(text_run_rgb(tr)) for tr in runs)):
                reasons.append(f'文字不是白色: {[tr[4] for tr in runs]}')
            # 字体为Arial或Calibri
            if not (runs and all(run_font_ok(tr) for tr in runs)):
                reasons.append('字体不是Arial/Calibri')
            # 字号10–12磅
            if not (runs and all(10 <= tr[2] <= 12 for tr in runs)):
                reasons.append(f'字号不在10-12磅: {[tr[2] for tr in runs]}')
            # 加粗
            if not (runs and all(tr[3] for tr in runs)):
                reasons.append('未加粗')

        if not reasons:
            return True, f'找到arena深蓝标题条: {s.name} w={s.abs_w:.2f} h={s.abs_h:.2f}'
        failures.append(f'{s.name}: ' + '；'.join(reasons))
    if failures:
        return False, '找到标题条圆角矩形但不符合要求：' + ' | '.join(failures[:2])
    return False, '未找到符合要求的arena顶部深蓝标题条'

def chk_p1_demand_module(slides, **_):
    """+3: 第1页下方蓝色边框圆角矩形外框左侧需求模块：位于距左9.5cm–13.8cm、
    距上11.2cm–16.5cm范围内，包含两个蓝色虚线圆角矩形，文本为"Service demand"
    "Equity needs"，字体为Arial或Calibri，字号9–11磅，加粗，两个框上下排列；
    单框宽3.5cm–4cm，高1.3cm–1.9cm，两个矩形竖向排列，间距为1.4cm–1.6cm，
    边线为蓝色虚线，线宽0.75–1.25磅，填充为白色或浅蓝白色。"""
    shapes = slides[0]
    xl, xr, yt, yb = 9.5, 13.8, 11.2, 16.5
    tol = 0.2  # cm，容忍 EMU→cm 取整与屏幕测量误差
    reasons = []

    # 两个蓝色虚线圆角矩形（位置/尺寸/边线/线宽/填充符合细则）
    valid_rects = []
    rect_failures = []
    for s in shapes:
        if s.geom != 'roundRect': continue
        if not bbox_in_box(s, xl, xr, yt, yb): continue
        rf = []
        # 单框宽3.5–4cm、高1.3–1.9cm
        if not in_range(s.abs_w, 3.5 - tol, 4.0 + tol): rf.append(f'宽{s.abs_w:.2f}cm不在3.5-4cm')
        if not in_range(s.abs_h, 1.3 - tol, 1.9 + tol): rf.append(f'高{s.abs_h:.2f}cm不在1.3-1.9cm')
        # 边线为蓝色虚线
        line = shape_line_rgb(s)
        if not (line and is_blue(line) and is_dashed(s)):
            rf.append(f'边线不是蓝色虚线: color={s.line_color}, dash={s.line_dash}')
        # 线宽0.75–1.25磅
        if not lw_in(s, 0.75, 1.25): rf.append(f'线宽{s.line_width_pt}不在0.75-1.25磅')
        # 填充为白色或浅蓝白色
        fill = shape_fill_rgb(s)
        if not (fill and is_white_or_blue_white(fill)):
            rf.append(f'填充不是白色或浅蓝白色: {s.fill_color}')
        if not rf:
            valid_rects.append(s)
        else:
            rect_failures.append(f'{s.name}: ' + '；'.join(rf))
    if len(valid_rects) != 2:
        reasons.append(f'符合要求的蓝色虚线圆角矩形数量为{len(valid_rects)}，需要2个；'
                       + ' | '.join(rect_failures[:2]))
    else:
        # 两个矩形竖向排列（上下），间距1.4–1.6cm
        ordered = sorted(valid_rects, key=lambda s: s.top)
        gap = ordered[1].top - ordered[0].bottom
        if not (ordered[0].bottom <= ordered[1].top + tol and 1.4 - tol <= gap <= 1.6 + tol):
            reasons.append(f'两个矩形不是竖向上下排列或间距不在1.4-1.6cm: gap={gap:.2f}')

    # 两个文本"Service demand""Equity needs"（字体/字号/加粗），且上下排列
    texts = []
    for phrase in ['Service demand', 'Equity needs']:
        matches = [s for s in shapes if s.full_text().strip() == phrase
                   and shape_in_box(s, xl, xr, yt, yb)]
        if not matches:
            reasons.append(f'缺少文本"{phrase}"')
            continue
        s = matches[0]
        texts.append(s)
        runs = non_empty_runs(s)
        if not (runs and all(run_font_ok(tr) for tr in runs)):
            reasons.append(f'{phrase}字体不是Arial/Calibri')
        if not (runs and all(9 <= tr[2] <= 11 for tr in runs)):
            reasons.append(f'{phrase}字号不在9-11磅: {[tr[2] for tr in runs]}')
        if not (runs and all(tr[3] for tr in runs)):
            reasons.append(f'{phrase}未加粗')
    if len(texts) == 2 and not (texts[0].top < texts[1].top):
        reasons.append('两个文本未上下排列')

    if not reasons:
        return True, '需求模块两个蓝色虚线圆角矩形和文本均符合要求'
    return False, '需求模块不符合要求：' + '；'.join(reasons[:4])

def chk_p1_demand_icons(slides, **_):
    """+3: 第1页下方蓝色边框圆角矩形外框左侧图标组：位于距左9.2cm–13.7cm、
    距上11.4cm–16.1cm范围内，包含蓝色柱状图图标、人群图标和双向竖箭头，顺序为
    蓝色虚线圆角矩形内蓝色柱状图图标-双向箭头-蓝色虚线圆角矩形内人群图标；
    柱状图由3–4个蓝色竖向矩形组成，组合图高0.6cm-0.8cm、宽0.8cm-1cm；
    人群图标为蓝色圆形头部与身体线条，组合图高0.6cm-1cm、宽1.0-1.4cm；
    双向箭头为蓝色竖向双箭头，高1.4-1.6cm，线宽1–1.5磅。"""
    shapes = slides[0]
    cands = shapes_in(shapes, 9.2, 13.7, 11.4, 16.1)
    tol = 0.2  # cm，容忍 EMU→cm 取整与屏幕测量误差

    def blue(s):
        rgb = shape_line_rgb(s) or shape_fill_rgb(s)
        return bool(rgb and is_blue(rgb))

    # 柱状图：3-4个蓝色竖向矩形 + 底部基线，组合图高0.6-0.8cm、宽0.8-1cm
    bars = [s for s in cands if s.geom == 'rect' and blue(s)
            and s.abs_h > s.abs_w and s.abs_w <= 0.35]
    bar_parts = list(bars)
    if bars:
        blo = min(s.top for s in bars); bhi = max(s.bottom for s in bars)
        # 柱状图基线：与柱子相邻的蓝色横线
        for s in cands:
            if s.geom == 'line' and blue(s) and s.abs_w > s.abs_h:
                if blo - 0.4 <= s.top <= bhi + 0.4:
                    bar_parts.append(s)
    bar_ok = False
    if 3 <= len(bars) <= 4 and bar_parts:
        left, top, right, bottom = group_bbox(bar_parts)
        bar_ok = (0.8 - tol <= right - left <= 1.0 + tol) and (0.6 - tol <= bottom - top <= 0.8 + tol)

    # 人群图标：蓝色圆形头部(ellipse) + 身体线条(arc)，组合图高0.6-1cm、宽1.0-1.4cm
    heads = [s for s in cands if s.geom == 'ellipse' and blue(s)]
    bodies = [s for s in cands if s.geom == 'arc' and blue(s)]
    crowd_ok = False
    crowd_center_y = None
    if heads and bodies:
        left, top, right, bottom = group_bbox(heads + bodies)
        crowd_center_y = (top + bottom) / 2
        crowd_ok = (1.0 - tol <= right - left <= 1.4 + tol) and (0.6 - tol <= bottom - top <= 1.0 + tol)

    # 双向箭头：蓝色竖向双箭头(可由两段带箭头竖线组成)，高1.4-1.6cm，线宽1-1.5磅
    arrow_lines = [s for s in cands if has_any_arrow(s) and blue(s) and s.abs_w < 0.2]
    double_arrow_ok = False
    arrow_center_y = None
    if arrow_lines:
        left, top, right, bottom = group_bbox(arrow_lines)
        arrow_center_y = (top + bottom) / 2
        height_ok = 1.4 - tol <= bottom - top <= 1.6 + tol
        lw_ok = all(lw_in(s, 1.0, 1.5) for s in arrow_lines)
        # 双箭头：两端都有箭头（单形状含首尾箭头，或多段合起来含上下两个箭头）
        both_ends = (any(s.has_head_arrow for s in arrow_lines)
                     and any(s.has_tail_arrow for s in arrow_lines)) or len(arrow_lines) >= 2
        double_arrow_ok = height_ok and lw_ok and both_ends

    reasons = []
    if not bar_ok:
        if bar_parts:
            left, top, right, bottom = group_bbox(bar_parts)
            reasons.append(f'柱状图尺寸或数量不符: 竖矩形={len(bars)}, 组合尺寸=({right-left:.2f},{bottom-top:.2f})')
        else:
            reasons.append('未找到3-4个蓝色竖向矩形柱状图')
    if not crowd_ok:
        reasons.append(f'人群图标尺寸或组成不符: 头={len(heads)}, 身体={len(bodies)}')
    if not double_arrow_ok:
        reasons.append(f'蓝色竖向双箭头不符(高1.4-1.6cm/线宽1-1.5磅): '
                       f'{[(s.name, round(s.abs_h,2), s.line_width_pt) for s in arrow_lines]}')
    if not (bar_ok and crowd_ok and double_arrow_ok):
        return False, '需求模块图标组不符合要求：' + '；'.join(reasons)

    # 顺序：柱状图(上) - 双向箭头(中) - 人群图标(下)
    bar_center_y = sum(s.y + s.h / 2 for s in bars) / len(bars)
    if not (bar_center_y < arrow_center_y < crowd_center_y):
        return False, (f'图标顺序不符合柱状图-双向箭头-人群: '
                       f'y={bar_center_y:.2f},{arrow_center_y:.2f},{crowd_center_y:.2f}')
    return True, '需求模块柱状图、人群图标和蓝色竖向双箭头均符合要求'

def chk_p1_arena_vline(slides, **_):
    """+1: 第1页下方蓝色边框圆角矩形外框中部竖向虚线：位于距左14.5cm–15.1cm、
    距上11.0cm–16.9cm范围内，为蓝色竖向虚线，高5.3-5.7cm，线宽0.75–1.25磅，
    线条与水平方向角度约90度。"""
    shapes = slides[0]
    tol = 0.2  # cm，容忍 EMU→cm 取整与屏幕测量误差
    failures = []
    for s in shapes:
        if s.shape_type not in ('sp', 'cxnSp'):
            continue
        # 位于距左14.5–15.1cm、距上11.0–16.9cm范围内
        left = min(s.x, s.x + s.w)
        cy = s.y + s.h / 2
        if not (14.5 - tol <= left <= 15.1 + tol): continue
        if not (11.0 - tol <= cy <= 16.9 + tol): continue
        reasons = []
        # 蓝色
        line = shape_line_rgb(s)
        if not (line and is_blue(line)):
            reasons.append(f'不是蓝色: {s.line_color}')
        # 虚线
        if not is_dashed(s):
            reasons.append(f'不是虚线: dash={s.line_dash}')
        # 高5.3–5.7cm
        if not in_range(s.abs_h, 5.3 - tol, 5.7 + tol):
            reasons.append(f'高{s.abs_h:.2f}cm不在5.3-5.7cm')
        # 线宽0.75–1.25磅
        if not lw_in(s, 0.75, 1.25):
            reasons.append(f'线宽{s.line_width_pt}不在0.75-1.25磅')
        # 竖向 / 线条与水平方向角度约90度
        if not (85 <= line_angle_from_horizontal(s) <= 95):
            reasons.append(f'线条角度{line_angle_from_horizontal(s):.0f}°不约90度')
        if not reasons:
            return True, (f'找到蓝色竖向虚线: {s.name} h={s.abs_h:.2f} lw={s.line_width_pt} '
                          f'angle={line_angle_from_horizontal(s):.0f}°')
        failures.append(f'{s.name}: ' + '；'.join(reasons))
    if failures:
        return False, '找到竖向线但不符合要求：' + ' | '.join(failures[:3])
    return False, '未在(距左14.5-15.1, 距上11.0-16.9)找到符合要求的蓝色竖向虚线'

def chk_p1_arena_h_arrow(slides, **_):
    """+1: 第1页下方蓝色边框圆角矩形外框中部向右箭头：位于距左13.0cm–14.7cm、
    距上13.2cm–14.6cm范围内，为蓝色水平箭头，线宽1–1.5磅，箭头朝右，
    连接左侧需求模块（x≈9.5-13.8cm）与右侧主体网络区域（x≈15.0-24.5cm）。"""
    shapes = slides[0]
    tol = 0.2  # cm，容忍 EMU→cm 取整与屏幕测量误差

    # 端点连接的对端参照：左侧需求模块和右侧主体网络（四椭圆）
    demand_xl, demand_xr, demand_yt, demand_yb = 9.5, 13.8, 11.2, 16.5
    arena_xl, arena_xr, arena_yt, arena_yb = 15.0, 24.5, 11.1, 17.0

    demand_rects = [s for s in shapes
                    if s.geom == 'roundRect'
                    and bbox_in_box(s, demand_xl, demand_xr, demand_yt, demand_yb)
                    and shape_line_rgb(s) and is_blue(shape_line_rgb(s))
                    and is_dashed(s)]
    arena_ellipses = [s for s in shapes
                      if s.geom == 'ellipse'
                      and bbox_in_box(s, arena_xl, arena_xr, arena_yt, arena_yb)]

    def arrow_connects(seg):
        """箭头两端是否几何连接左需求模块与右主体网络区。"""
        left_x, top_y, right_x, bottom_y = rendered_bbox(seg)
        y_mid = (top_y + bottom_y) / 2
        connect_tol = 0.6  # cm，允许约 6mm 的接头缝隙

        # 左端点：与某个需求矩形的右边缘接触，且垂直方向在其内部
        left_ok = False
        if demand_rects:
            for r in demand_rects:
                rleft, rtop, rright, rbot = rendered_bbox(r)
                if (abs(left_x - rright) <= connect_tol
                        and rtop - tol <= y_mid <= rbot + tol):
                    left_ok = True; break
        else:
            # 兜底：只要左端进入需求模块 x 区域即视为连接
            left_ok = left_x <= demand_xr + connect_tol

        # 右端点：与主体网络区左侧椭圆的左边缘接触，或直接进入主体区域
        right_ok = False
        if arena_ellipses:
            for e in arena_ellipses:
                eleft, etop, eright, ebot = rendered_bbox(e)
                if (abs(right_x - eleft) <= connect_tol
                        and etop - tol <= y_mid <= ebot + tol):
                    right_ok = True; break
        if not right_ok:
            right_ok = right_x >= arena_xl - connect_tol

        return left_ok, right_ok

    failures = []
    for s in shapes:
        # 位于距左13.0–14.7cm、距上13.2–14.6cm范围内
        left, top, right, bottom = rendered_bbox(s)
        if not (13.0 - tol <= left and right <= 14.7 + tol
                and 13.2 - tol <= top and bottom <= 14.6 + tol):
            continue
        reasons = []
        # 箭头朝右的水平箭头：一体箭头(预设箭头几何 或 线条+端点箭头)且视觉朝右
        is_right_arrow = is_arrow_shape(s) and arrow_points_right(s)
        if not is_right_arrow:
            reasons.append(f'不是朝右的水平箭头: geom={s.geom}')
        # 水平（形状未旋转，约0度）
        if not near_angle(s, 0, 5):
            reasons.append(f'不是水平方向(rot={s.rotation_deg():.0f}°)')
        # 蓝色
        rgb = shape_fill_rgb(s) or shape_line_rgb(s)
        if not (rgb and is_blue(rgb)):
            reasons.append(f'不是蓝色: fill={s.fill_color}, line={s.line_color}')
        # 线宽1–1.5磅
        if not lw_in(s, 1.0, 1.5):
            reasons.append(f'线宽{s.line_width_pt}不在1-1.5磅')
        # 端点连接左侧需求模块与右侧主体网络区
        left_ok, right_ok = arrow_connects(s)
        if not left_ok:
            reasons.append(f'左端未连接需求模块(左端x={left:.2f}cm，需求模块右缘参考{demand_xr}cm)')
        if not right_ok:
            reasons.append(f'右端未连接主体网络区(右端x={right:.2f}cm，主体网络左缘参考{arena_xl}cm)')
        if not reasons:
            return True, (f'找到arena中部蓝色向右箭头: {s.name} geom={s.geom} '
                          f'lw={s.line_width_pt}，端点连接需求模块与主体网络区')
        failures.append(f'{s.name}: ' + '；'.join(reasons))
    if failures:
        return False, '找到候选但不符合要求：' + ' | '.join(failures[:3])
    return False, '未在(距左13.0-14.7, 距上13.2-14.6)找到蓝色水平向右箭头'

def chk_p1_four_ellipses(slides, **_):
    """+5: 第1页下方蓝色边框圆角矩形外框内右侧四个主体椭圆：位于距左15.0cm–24.5cm、
    距上11.1cm–17cm范围内，包含内容为"Hospitals""Start-ups""Government""Citizens"
    四个浅色椭圆，文本字体为Arial或Calibri，字号9–11磅，四个椭圆呈2行2列排列，
    单个椭圆宽3cm–3.5cm，高1.5cm–1.8cm，填充为浅蓝白色，边线为蓝色单实线，
    线宽0.75–1.25磅。"""
    shapes = slides[0]
    xl, xr, yt, yb = 15.0, 24.5, 11.1, 17.0
    tol = 0.2  # cm，容忍 EMU→cm 取整与屏幕测量误差
    keywords = ['Hospitals', 'Start-ups', 'Government', 'Citizens']
    reasons = []

    # 四个内容文本
    labels = {k: [s for s in shapes if s.full_text().strip() == k
                  and shape_in_box(s, xl, xr, yt, yb)] for k in keywords}
    missing = [k for k, m in labels.items() if not m]
    if missing:
        reasons.append(f'缺少文本: {missing}')
    # 文本字体Arial/Calibri，字号9–11磅
    text_failures = []
    for k, m in labels.items():
        if not m: continue
        for tr in non_empty_runs(m[0]):
            if not run_font_ok(tr): text_failures.append(f'{k}字体不是Arial/Calibri: {tr[1]}')
            if not (9 <= tr[2] <= 11): text_failures.append(f'{k}字号{tr[2]}不在9-11磅')
    if text_failures:
        reasons.append('；'.join(text_failures[:4]))

    # 四个浅色椭圆（尺寸/填充/边线/线宽符合细则）
    valid_ellipses = []
    ellipse_failures = []
    for s in shapes:
        if s.geom != 'ellipse': continue
        left, top, right, bottom = rendered_bbox(s)
        if not (xl - tol <= left and right <= xr + tol
                and yt - tol <= top and bottom <= yb + tol): continue
        rf = []
        # 单个椭圆宽3–3.5cm、高1.5–1.8cm（宽高容差±0.05cm）
        size_tol = 0.05
        if not in_range(s.abs_w, 3.0 - size_tol, 3.5 + size_tol): rf.append(f'宽{s.abs_w:.2f}cm不在3-3.5cm')
        if not in_range(s.abs_h, 1.5 - size_tol, 1.8 + size_tol): rf.append(f'高{s.abs_h:.2f}cm不在1.5-1.8cm')
        # 填充为浅蓝白色
        fill = shape_fill_rgb(s)
        if not (fill and is_white_or_blue_white(fill)): rf.append(f'填充不是浅蓝白色: {s.fill_color}')
        # 边线为蓝色单实线
        line = shape_line_rgb(s)
        if not (line and is_blue(line) and is_solid_line(s)):
            rf.append(f'边线不是蓝色单实线: color={s.line_color}, dash={s.line_dash}')
        # 线宽0.75–1.25磅
        if not lw_in(s, 0.75, 1.25): rf.append(f'线宽{s.line_width_pt}不在0.75-1.25磅')
        if rf:
            ellipse_failures.append(f'{s.name}: ' + '；'.join(rf))
        else:
            valid_ellipses.append(s)
    if len(valid_ellipses) != 4:
        reasons.append(f'符合要求的椭圆{len(valid_ellipses)}个，需要4个；{ellipse_failures[:3]}')
    else:
        # 2行2列排列：x 有两个不同列、y 有两个不同行
        xs = sorted({round(s.x, 1) for s in valid_ellipses})
        ys = sorted({round(s.y, 1) for s in valid_ellipses})
        def cluster(vals):
            groups = []
            for v in vals:
                if groups and v - groups[-1][-1] <= 1.0:
                    groups[-1].append(v)
                else:
                    groups.append([v])
            return len(groups)
        if not (cluster(xs) == 2 and cluster(ys) == 2):
            reasons.append(f'四个椭圆不是2行2列排列: xs={xs}, ys={ys}')

    if not reasons:
        return True, '四主体椭圆和文本均符合要求'
    return False, '四主体椭圆不符合要求：' + '；'.join(reasons[:4])

def chk_p1_four_icons(slides, **_):
    """+3: 第1页下方蓝色边框圆角矩形外框内右侧四个主体图标：位于四个主体椭圆内部
    文本下方，分别包含医院建筑图标、火箭或创业图标、政府建筑图标、公民人群图标，
    图标颜色为蓝色，高0.45-0.7cm，宽0.4cm–1cm。

    语义识别：优先按形状 name/descr（alt text）关键词判断；无关键词时按几何组合特征
    兜底判断（医院=矩形+十字；火箭=三角+竖矩形；政府=多柱+屋顶三角；公民=多个小圆/人形）。
    """
    shapes = slides[0]
    xl, xr, yt, yb = 15.0, 24.5, 11.1, 17.0
    tol = 0.2  # cm，容忍 EMU→cm 取整与屏幕测量误差
    keywords = ['Hospitals', 'Start-ups', 'Government', 'Citizens']

    # 关键词映射：形状 name / alt text 里若命中则视为该图标
    NAME_HINTS = {
        'Hospitals':   ('hospital', 'medical', '医院', 'clinic', 'cross', '红十字'),
        'Start-ups':   ('rocket', 'launch', 'startup', 'start-up', '火箭', '创业'),
        'Government':  ('government', 'capitol', 'building', 'bank', '政府', '建筑'),
        'Citizens':    ('citizen', 'people', 'person', 'group', 'crowd', 'user',
                        '公民', '人群', '群众', '人物'),
    }

    labels = {k: [s for s in shapes if s.full_text().strip() == k
                  and shape_in_box(s, xl, xr, yt, yb)] for k in keywords}
    # 四个主体椭圆（按尺寸筛出大椭圆）
    big_ellipses = [s for s in shapes if s.geom == 'ellipse'
                    and in_range(s.abs_w, 3.0 - tol, 3.5 + tol)
                    and in_range(s.abs_h, 1.5 - tol, 1.8 + tol)
                    and shape_in_box(s, xl, xr, yt, yb)]

    def _semantic_from_names(parts):
        """按 name/descr 关键词判断图标类别；返回 keyword 或 None。"""
        blob = ' '.join((getattr(p, 'name', '') + ' ' + getattr(p, 'descr', '')).lower()
                        for p in parts)
        if not blob.strip():
            return None
        for k, hints in NAME_HINTS.items():
            if any(h.lower() in blob for h in hints):
                return k
        return None

    def _semantic_from_geometry(parts, keyword):
        """按几何组合兜底判断图标语义。返回 (ok, reason)。"""
        # 分组统计部件类型
        rects = [p for p in parts if p.geom in ('rect', 'roundRect')]
        tris  = [p for p in parts if p.geom in ('triangle', 'rtTriangle')]
        ellips = [p for p in parts if p.geom == 'ellipse']
        lines  = [p for p in parts if p.shape_type == 'cxnSp' or p.geom in ('line', 'straightConnector1')]
        crosses = [p for p in parts if p.geom in ('plus', 'mathPlus')]
        n_parts = len(parts)

        if keyword == 'Hospitals':
            # 建筑主体（矩形/圆角矩形）+ 十字（plus 几何 或 两条正交短线 或 白色小矩形叠加）
            has_body = len(rects) >= 1
            has_cross = bool(crosses) or (
                # 两条正交细线且接近正中
                any(near_angle(l, 0, 15) for l in lines)
                and any(near_angle(l, 90, 15) for l in lines)
            )
            if has_body and (has_cross or n_parts >= 2):
                return True, 'ok'
            return False, f'医院图标缺矩形建筑或十字标记(rects={len(rects)}, crosses={len(crosses)}, lines={len(lines)})'

        if keyword == 'Start-ups':
            # 火箭：竖直主体（roundRect/rect）+ 三角顶（triangle）；或整体带三角组合
            has_body = any(p.geom in ('rect', 'roundRect') and p.abs_h >= p.abs_w * 0.8 for p in parts)
            has_nose = bool(tris)
            if (has_body and has_nose) or n_parts >= 2:
                return True, 'ok'
            return False, f'火箭图标缺竖直主体或三角顶(rects={len(rects)}, tris={len(tris)})'

        if keyword == 'Government':
            # 政府建筑：≥2 个竖矩形柱 + 屋顶（三角形或宽扁矩形）
            columns = [p for p in rects if p.abs_h > p.abs_w * 0.8]
            has_roof = bool(tris) or any(p.abs_w > p.abs_h * 1.6 for p in rects)
            if len(columns) >= 2 and has_roof:
                return True, 'ok'
            if len(rects) >= 3:
                return True, 'ok'  # 多矩形组合，视为建筑
            return False, f'政府图标缺多柱建筑与屋顶(columns={len(columns)}, tris={len(tris)}, rects={len(rects)})'

        if keyword == 'Citizens':
            # 公民人群：多个小圆（头部）或者多个复合小图形组合
            heads = [p for p in ellips if p.abs_w <= 0.35 and p.abs_h <= 0.35]
            if len(heads) >= 2 or len(ellips) >= 2:
                return True, 'ok'
            if n_parts >= 3:
                return True, 'ok'
            return False, f'公民图标缺多个人形/小圆(ellips={len(ellips)}, heads={len(heads)}, parts={n_parts})'

        return False, 'unknown'

    found = {}
    details = []
    for k in keywords:
        matches = labels[k]
        if not matches:
            details.append(f'{k}: 缺少文本定位')
            continue
        label = matches[0]
        if not big_ellipses:
            details.append(f'{k}: 缺少椭圆定位')
            continue
        lcx, lcy = label.x + label.w / 2, label.y + label.h / 2
        # 与该标签同属一个椭圆（中心最近）
        ellipse = min(big_ellipses,
                      key=lambda e: ((e.x + e.w / 2) - lcx) ** 2 + ((e.y + e.h / 2) - lcy) ** 2)
        # 图标 = 椭圆内、文本下方的蓝色非文本部件的组合
        icon_parts = []
        for s in shapes:
            if s is ellipse or s.texts: continue
            cx, cy = s.x + s.w / 2, s.y + s.h / 2
            if not (ellipse.left <= cx <= ellipse.right and ellipse.top <= cy <= ellipse.bottom):
                continue
            if cy < label.bottom - 0.15:  # 须位于文本下方
                continue
            rgb = shape_line_rgb(s) or shape_fill_rgb(s)
            if rgb and is_blue(rgb):
                icon_parts.append(s)
        if not icon_parts:
            details.append(f'{k}: 未找到文本下方的蓝色图标部件')
            continue
        left, top, right, bottom = group_bbox(icon_parts)
        w, h = right - left, bottom - top
        # 尺寸校验：宽0.4-1cm、高0.45-0.7cm
        size_ok = (0.4 - tol <= w <= 1.0 + tol) and (0.45 - tol <= h <= 0.7 + tol)
        if not size_ok:
            details.append(f'{k}: 图标组合尺寸({w:.2f},{h:.2f})不在宽0.4-1cm、高0.45-0.7cm')
            continue

        # 语义判定：先按 name/alt text 匹配；否则用几何组合兜底
        sem_key = _semantic_from_names(icon_parts)
        if sem_key is not None:
            if sem_key != k:
                details.append(f'{k}: 图标语义按命名判定为"{sem_key}"，与所在椭圆标签不一致')
                continue
            sem_reason = '命名/alt text命中'
        else:
            ok, reason = _semantic_from_geometry(icon_parts, k)
            if not ok:
                details.append(f'{k}: 图形语义判定失败({reason})')
                continue
            sem_reason = '几何组合命中'

        found[k] = (round(w, 2), round(h, 2), sem_reason)

    missing = [k for k in keywords if k not in found]
    if not missing:
        return True, f'四主体图标(含语义)均符合要求: {found}'
    return False, '四主体图标不符合要求：' + '；'.join(details[:6])

def chk_p1_four_bidir_arrows(slides, **_):
    """+3: 第1页下方蓝色边框圆角矩形外框右侧四个主体图标外的蓝色椭圆间有四个双向箭头，
    "Hospitals"-"Start-ups"、"Government"-"Citizens"间的横向双向箭头
    宽2.45-2.65cm、高0.25-0.45cm；
    "Start-ups"-"Government"、"Citizens"-"Hospitals"间的竖向双向箭头
    高1.75-1.95cm、宽0.25-0.45cm。"""
    shapes = slides[0]
    cands = shapes_in(shapes, 15.0, 25.0, 11.0, 17.5)
    tol = 0.2  # cm，容忍 EMU→cm 取整与屏幕测量误差
    # 蓝色带箭头线段
    arrow_lines = [s for s in cands if has_any_arrow(s)
                   and shape_line_rgb(s) and is_blue(shape_line_rgb(s))]

    def group_double_arrows(segs, horizontal):
        """把两条重叠共线的半箭头线段合并为一个双向箭头，返回其可见包围盒列表。"""
        used = [False] * len(segs)
        result = []
        for i, a in enumerate(segs):
            if used[i]: continue
            la, ta, ra, ba = rendered_bbox(a)
            group = [a]
            for j in range(i + 1, len(segs)):
                if used[j]: continue
                b = segs[j]
                lb, tb, rb, bb = rendered_bbox(b)
                if horizontal:
                    # 同一水平轴(y接近)且 x 跨度重叠
                    same_axis = abs((ta + ba) / 2 - (tb + bb) / 2) <= 0.3
                    overlap = min(ra, rb) - max(la, lb) > 0
                else:
                    same_axis = abs((la + ra) / 2 - (lb + rb) / 2) <= 0.3
                    overlap = min(ba, bb) - max(ta, tb) > 0
                if same_axis and overlap:
                    group.append(b); used[j] = True
            used[i] = True
            result.append(group_bbox(group))
        return result

    # 只保留轴对齐（近水平/近竖直）的箭头线段，排除对角装饰线
    horiz_segs = [s for s in arrow_lines if line_angle_from_horizontal(s) <= 5]
    vert_segs = [s for s in arrow_lines if line_angle_from_horizontal(s) >= 85]
    h_arrows = group_double_arrows(horiz_segs, True)
    v_arrows = group_double_arrows(vert_segs, False)

    # 横向双向箭头：宽2.45-2.65cm、高0.25-0.45cm
    h_ok = [(l, t, r, b) for (l, t, r, b) in h_arrows
            if (2.45 - tol <= r - l <= 2.65 + tol) and (0.25 - tol <= b - t <= 0.45 + tol)]
    # 竖向双向箭头：高1.75-1.95cm、宽0.25-0.45cm
    v_ok = [(l, t, r, b) for (l, t, r, b) in v_arrows
            if (1.75 - tol <= b - t <= 1.95 + tol) and (0.25 - tol <= r - l <= 0.45 + tol)]

    reasons = []
    if len(h_ok) < 2:
        reasons.append(f'符合要求的横向双向箭头{len(h_ok)}个，需要2个(宽2.45-2.65cm/高0.25-0.45cm)；'
                       f'实测横向={[(round(r-l,2), round(b-t,2)) for (l,t,r,b) in h_arrows]}')
    if len(v_ok) < 2:
        reasons.append(f'符合要求的竖向双向箭头{len(v_ok)}个，需要2个(高1.75-1.95cm/宽0.25-0.45cm)；'
                       f'实测竖向={[(round(r-l,2), round(b-t,2)) for (l,t,r,b) in v_arrows]}')
    if not reasons:
        return True, f'四个双向箭头符合要求: 横向{len(h_ok)}个, 竖向{len(v_ok)}个'
    return False, '四主体双向箭头不符合要求：' + '；'.join(reasons)

def chk_p1_center_text(slides, **_):
    """+1: 第1页下方蓝色边框圆角矩形外框右侧中心说明文本：位于距左17.0cm–22.4cm、
    距上13.1cm–14.7cm范围内，文本为"alignment, negotiation, and shared experimentation"，
    字体为Arial或Calibri，字号9–11磅，颜色为蓝色或深灰色，位于四个主体椭圆之间。"""
    shapes = slides[0]
    tol = 0.2  # cm，容忍 EMU→cm 取整与屏幕测量误差
    failures = []
    for s in shapes:
        # 文本为"alignment, negotiation, and shared experimentation"（可能分多段/多行，
        # 规范化空白后精确匹配）
        ft = ' '.join(s.full_text().split())
        if ft.lower() != 'alignment, negotiation, and shared experimentation':
            continue
        left, top, right, bottom = rendered_bbox(s)
        runs = non_empty_runs(s)
        reasons = []
        # 距左17.0–22.4cm、距上13.1–14.7cm（以文本框左上角为准）
        if not ((17.0 - tol <= left <= 22.4 + tol) and (13.1 - tol <= top <= 14.7 + tol)):
            reasons.append(f'位置left={left:.2f}/top={top:.2f}不在距左17-22.4/距上13.1-14.7范围内')
        # 字体为Arial或Calibri
        if not (runs and all(run_font_ok(tr) for tr in runs)):
            reasons.append('字体不是Arial/Calibri')
        # 字号9–11磅
        if not (runs and all(9 <= tr[2] <= 11 for tr in runs)):
            reasons.append(f'字号不在9-11磅: {[tr[2] for tr in runs]}')
        # 颜色为蓝色或深灰色
        if not (runs and all(
                text_run_rgb(tr) is None
                or is_blue(text_run_rgb(tr)) or is_dark_gray(text_run_rgb(tr)) for tr in runs)):
            reasons.append(f'颜色不是蓝色或深灰色: {[tr[4] for tr in runs]}')
        if not reasons:
            return True, f'找到中心说明文本: {s.name} left={left:.2f} top={top:.2f}'
        failures.append(f'{s.name}: ' + '；'.join(reasons))
    if failures:
        return False, '找到中心说明文本但不符合要求：' + ' | '.join(failures[:2])
    return False, '未找到"alignment, negotiation, and shared experimentation"'

def chk_p1_dashed_arrows_to_ellipses(slides, **_):
    """+3: "alignment, negotiation, and shared experimentation"文本框与四个蓝色椭圆间
    有四个蓝色虚线箭头，长度0.8-1.5cm，由该文本分别指向内容为"Hospitals""Start-ups"
    "Government""Citizens"的四个浅色椭圆。"""
    shapes = slides[0]
    tol = 0.2  # cm，容忍 EMU→cm 取整与屏幕测量误差
    # 定位中心说明文本
    center = next((s for s in shapes
                   if ' '.join(s.full_text().split()).lower()
                   == 'alignment, negotiation, and shared experimentation'), None)
    if center is None:
        return False, '未找到中心说明文本'
    ccx, ccy = center.x + center.w / 2, center.y + center.h / 2

    # 定位四个主体椭圆（按标签文本关联最近的大椭圆）
    keywords = ['Hospitals', 'Start-ups', 'Government', 'Citizens']
    big_ellipses = [s for s in shapes if s.geom == 'ellipse'
                    and in_range(s.abs_w, 3.0 - tol, 3.5 + tol)
                    and in_range(s.abs_h, 1.5 - tol, 1.8 + tol)]
    targets = []
    for k in keywords:
        lab = next((s for s in shapes if s.full_text().strip() == k), None)
        if lab is None or not big_ellipses:
            continue
        lcx, lcy = lab.x + lab.w / 2, lab.y + lab.h / 2
        el = min(big_ellipses, key=lambda e: ((e.x+e.w/2)-lcx)**2 + ((e.y+e.h/2)-lcy)**2)
        targets.append((k, el.x + el.w/2, el.y + el.h/2))

    # 蓝色虚线箭头
    arrows = [s for s in shapes if has_any_arrow(s)
              and shape_line_rgb(s) and is_blue(shape_line_rgb(s)) and is_dashed(s)]
    # 长度0.8-1.5cm
    dashed_in_range = [s for s in arrows
                       if 0.8 - tol <= (s.abs_w**2 + s.abs_h**2)**0.5 <= 1.5 + tol]

    # 每个箭头须从中心文本指向某个主体椭圆（起点近中心、终点近某椭圆）
    def links_center_to_target(s):
        p1 = (s.x, s.y)
        p2 = (s.x + s.w, s.y + s.h)
        near_center = min((abs(p[0]-ccx)+abs(p[1]-ccy)) for p in (p1, p2))
        for _, tx, ty in targets:
            near_target = min((abs(p[0]-tx)+abs(p[1]-ty)) for p in (p1, p2))
            if near_center <= 1.5 and near_target <= 1.5:
                return True
        return False

    linking = [s for s in dashed_in_range if links_center_to_target(s)]

    if len(linking) >= 4:
        return True, f'找到{len(linking)}条由中心文本指向四椭圆的蓝色虚线短箭头(0.8-1.5cm)'

    reasons = []
    if not arrows:
        reasons.append('区域内未找到蓝色虚线箭头(现有箭头为实线 dash=solid)')
    else:
        if not dashed_in_range:
            lens = [round((s.abs_w**2+s.abs_h**2)**0.5, 2) for s in arrows]
            reasons.append(f'蓝色虚线箭头长度均不在0.8-1.5cm: {lens}')
        reasons.append(f'由中心文本指向四椭圆的合规箭头{len(linking)}条(需要4条)')
    return False, '中心到四椭圆虚线箭头不符合要求：' + '；'.join(reasons)

def chk_p1_left_digital(slides, **_):
    """+1: 第1页左侧数字健康文字和箭头：位于距左5.4cm–9cm、距上12.7cm–15.2cm范围内，
    出现"Digital health initiatives (T1)"文本，颜色为蓝绿色，字体为Arial或Calibri，
    字号9-11磅，加粗；旁边有蓝绿色水平箭头，线宽1–1.5磅，箭头指向右侧。"""
    return _p1_digital_text_arrow(
        slides[0], 5.4, 9.0, 12.7, 15.2, 'Digital health initiatives (T1)', '左侧')

def _p1_digital_text_arrow(shapes, xl, xr, yt, yb, phrase, tag_name):
    tol = 0.2  # cm，容忍 EMU→cm 取整与屏幕测量误差
    reasons = []

    # 文本："Digital health initiatives (T1)"（可能分多行拼接，规范化空白后匹配）
    text = None
    for s in shapes:
        ft = ' '.join(s.full_text().split())
        if ft == phrase and shape_in_box(s, xl, xr, yt, yb):
            text = s
            break
    if text is None:
        reasons.append(f'未在({xl}-{xr},{yt}-{yb})找到"{phrase}"文本')
    else:
        left, top, right, bottom = rendered_bbox(text)
        runs = non_empty_runs(text)
        # 位于距左xl-xr、距上yt-yb
        if not ((xl - tol <= left <= xr + tol) and (yt - tol <= top <= yb + tol)):
            reasons.append(f'文本位置left={left:.2f}/top={top:.2f}不在范围内')
        # 颜色为蓝绿色
        if not (runs and all(text_run_rgb(tr) is not None and is_teal(text_run_rgb(tr)) for tr in runs)):
            reasons.append(f'文本颜色不是蓝绿色: {[tr[4] for tr in runs]}')
        # 字体为Arial或Calibri
        if not (runs and all(run_font_ok(tr) for tr in runs)):
            reasons.append('文本字体不是Arial/Calibri')
        # 字号9-11磅
        if not (runs and all(9 <= tr[2] <= 11 for tr in runs)):
            reasons.append(f'文本字号不在9-11磅: {[tr[2] for tr in runs]}')
        # 加粗
        if not (runs and all(tr[3] for tr in runs)):
            reasons.append('文本未加粗')

    # 旁边的蓝绿色水平箭头：线宽1-1.5磅，箭头指向右侧
    arrow = None
    for s in shapes:
        if not shape_in_box(s, xl, xr + 0.5, yt, yb): continue
        # 一体箭头(预设箭头几何 或 线条+端点箭头)，视觉朝右
        if not (is_arrow_shape(s) and arrow_points_right(s)): continue
        if not near_angle(s, 0, 5): continue
        rgb = shape_fill_rgb(s) or shape_line_rgb(s)
        if not (rgb and is_teal(rgb)): continue
        if not lw_in(s, 1.0, 1.5): continue
        arrow = s
        break
    if arrow is None:
        reasons.append('未找到蓝绿色水平向右箭头(线宽1-1.5磅)')

    if not reasons:
        return True, f'{tag_name}数字健康文字和蓝绿色向右箭头均符合要求'
    return False, f'{tag_name}数字健康文字/箭头不符合要求：' + '；'.join(reasons[:4])

def chk_p1_right_digital(slides, **_):
    """+1: 第1页右侧数字健康文字和箭头：位于距左25cm–29.5cm、距上13cm–15.4cm范围内，
    出现"Digital health outcomes (T2)"文本，颜色为蓝绿色，字体为Arial或Calibri，
    字号9–11磅，加粗；旁边有蓝绿色水平箭头，线宽1–1.5磅，箭头指向右侧或指向输出方向。"""
    return _p1_digital_text_arrow(
        slides[0], 25.0, 29.5, 13.0, 15.4, 'Digital health outcomes (T2)', '右侧')

# --- 第2页 ---
def chk_p2_y_arrow(slides, **_):
    """+1: 第2页左侧纵轴箭头：位于距左5.4cm–6.0cm、距上2.0cm–18cm范围内，
    为蓝色或深蓝色竖向单箭头线，线宽1.5–2.5磅，箭头朝上，线条角度为水平方向90度。"""
    shapes = slides[1]
    tol = 0.2  # cm，容忍 EMU→cm 取整误差
    failures = []
    for s in shapes:
        # 单箭头线
        if arrow_count(s) != 1: continue
        # 距左5.4–6.0cm（竖线 x 恒定，取左边界）
        left = min(s.x, s.x + s.w)
        if not (5.4 - tol <= left <= 6.0 + tol): continue
        reasons = []
        # 距上2.0–18cm范围内
        top = min(s.y, s.y + s.h)
        bottom = max(s.y, s.y + s.h)
        if not (2.0 - tol <= top and bottom <= 18.0 + tol):
            reasons.append(f'位置top={top:.2f}/bottom={bottom:.2f}不在距上2.0-18范围内')
        # 蓝色或深蓝色
        rgb = shape_line_rgb(s) or shape_fill_rgb(s)
        if not (rgb and (is_blue(rgb) or is_dark_blue(rgb))):
            reasons.append(f'颜色不是蓝色或深蓝色: {s.line_color or s.fill_color}')
        # 线宽1.5–2.5磅
        if not lw_in(s, 1.5, 2.5):
            reasons.append(f'线宽{s.line_width_pt}不在1.5-2.5磅')
        # 竖向 / 线条角度约90度
        if not (85 <= line_angle_from_horizontal(s) <= 95):
            reasons.append(f'线条角度{line_angle_from_horizontal(s):.0f}°不约90度')
        # 箭头朝上
        if not arrow_points_up(s):
            reasons.append('箭头不朝上')
        if not reasons:
            return True, (f'找到第2页纵轴箭头: {s.name} x={left:.2f} top={top:.2f} bottom={bottom:.2f} '
                          f'angle={line_angle_from_horizontal(s):.0f}° lw={s.line_width_pt}')
        failures.append(f'{s.name}: ' + '；'.join(reasons))
    if failures:
        return False, '找到纵轴箭头但不符合要求：' + ' | '.join(failures[:3])
    return False, '未找到符合细则的第2页蓝色/深蓝色竖向单箭头(距左5.4-6.0cm/距上2.0-18cm/线宽1.5-2.5磅/朝上/约90°)'

def chk_p2_y_labels(slides, **_):
    """+1: 第2页左侧纵轴文字组：位于距左4cm–6.5cm、距上0.8cm–17.4cm范围内，
    包含"Industry Maturity""High""Low"文本。
      · Industry Maturity: 位于箭头上方，Arial/Calibri，字号11–13磅，加粗，深蓝
      · High:            位于纵轴左侧上部，Arial/Calibri，字号12–14磅，深蓝
      · Low:             位于纵轴左侧下部，Arial/Calibri，字号12–14磅，深蓝
    """
    shapes = slides[1]
    tol = 0.2  # cm，容忍 EMU→cm 取整与屏幕测量误差
    failures = []
    found = {}
    for kw in ['Industry Maturity', 'High', 'Low']:
        matches = [s for s in shapes if s.full_text().strip() == kw]
        if not matches:
            failures.append(f'缺少文本"{kw}"')
            continue
        found[kw] = matches[0]

    # 三个 label 的 rubric 规格：(kw, size_lo, size_hi, need_bold, need_dark_blue)
    SPEC = [
        ('Industry Maturity', 11, 13, True,  True),
        ('High',              12, 14, False, True),
        ('Low',               12, 14, False, True),
    ]

    if len(found) == 3:
        # 整体位于距左4–6.5cm、距上0.8–17.4cm（以文本框左上角为准）
        for kw, s in found.items():
            left, top, _r, _b = rendered_bbox(s)
            if not ((4.0 - tol <= left <= 6.5 + tol) and (0.8 - tol <= top <= 17.4 + tol)):
                failures.append(f'{kw}位置left={left:.2f}/top={top:.2f}不在距左4-6.5/距上0.8-17.4范围内')

        # 逐一按 rubric 校验字体/字号/加粗/颜色
        for kw, lo_sz, hi_sz, need_bold, need_dark_blue in SPEC:
            s = found[kw]
            runs = non_empty_runs(s)
            if not runs:
                failures.append(f'{kw}: 无有效文本run')
                continue
            # 字体 Arial/Calibri
            if not all(run_font_ok(tr) for tr in runs):
                failures.append(f'{kw}字体不是Arial/Calibri: {[tr[1] for tr in runs]}')
            # 字号
            if not all(lo_sz <= tr[2] <= hi_sz for tr in runs):
                failures.append(f'{kw}字号不在{lo_sz}-{hi_sz}磅: {[tr[2] for tr in runs]}')
            # 加粗
            if need_bold and not all(tr[3] for tr in runs):
                failures.append(f'{kw}未加粗')
            # 颜色深蓝
            if need_dark_blue:
                if not all(text_run_rgb(tr) is not None
                           and is_dark_blue(text_run_rgb(tr)) for tr in runs):
                    failures.append(f'{kw}颜色不是深蓝: {[tr[4] for tr in runs]}')

        # Industry Maturity 位于 High 上方（即"箭头上方"更上一层）
        im = found['Industry Maturity']
        if not (im.bottom <= found['High'].top + tol):
            failures.append(f'Industry Maturity未位于箭头/High上方: bottom={im.bottom:.2f}')

        # High 位于上部、Low 位于下部（High 在 Low 之上）
        hi, lo = found['High'], found['Low']
        if not (hi.top < lo.top):
            failures.append(f'High/Low上下位置不符: High.top={hi.top:.2f}, Low.top={lo.top:.2f}')
    if not failures:
        return True, f'找到第2页纵轴全部文字且符合要求: {list(found)}'
    return False, '第2页纵轴文字组不符合要求：' + '；'.join(failures[:6])

def chk_p2_x_arrow(slides, **_):
    """+1: 第2页底部横轴箭头：位于距左5.6cm–29.6cm、距上17.6cm–17.9cm范围内，
    为黑色或深灰色横向单箭头线，线宽1.5–2.5磅，箭头朝右，线条角度为水平方向0度。"""
    shapes = slides[1]
    tol = 0.2  # cm，容忍 EMU→cm 取整误差
    for s in shapes:
        # 单箭头线
        if arrow_count(s) != 1: continue
        # 距左5.6–29.6cm（横线左右边界落在该带内）
        left = min(s.x, s.x + s.w)
        right = max(s.x, s.x + s.w)
        if not (5.6 - tol <= left and right <= 29.6 + tol): continue
        # 距上17.6–17.9cm（横线 y 恒定）
        top = min(s.y, s.y + s.h)
        bottom = max(s.y, s.y + s.h)
        if not (17.6 - tol <= top and bottom <= 17.9 + tol): continue
        # 横向 / 线条角度约0度
        if not (0 <= line_angle_from_horizontal(s) <= 5): continue
        # 黑色或深灰色（办公软件视觉呈黑/深灰）
        rgb = shape_line_rgb(s) or shape_fill_rgb(s)
        if not (rgb and is_visually_black(rgb)): continue
        # 线宽1.5–2.5磅
        if not lw_in(s, 1.5, 2.5): continue
        # 箭头朝右
        if not arrow_points_right(s): continue
        return True, (f'找到第2页横轴箭头: {s.name} left={left:.2f} right={right:.2f} top={top:.2f} '
                      f'angle={line_angle_from_horizontal(s):.0f}° lw={s.line_width_pt}')
    return False, '未找到符合细则的第2页黑色/深灰色横向单箭头(距左5.6-29.6cm/距上17.6-17.9cm/线宽1.5-2.5磅/朝右/约0°)'

def chk_p2_x_label(slides, **_):
    """+1: 第2页底部横轴标题文本：位于距左16cm–19cm、距上17.7cm–18.7cm范围内，
    文本为"Time"，字体为Arial或Calibri，字号13–15磅，加粗，颜色为黑色。"""
    shapes = slides[1]
    tol = 0.2  # cm，容忍 EMU→cm 取整与屏幕测量误差
    failures = []
    for s in shapes:
        # 文本为"Time"
        if s.full_text().strip() != 'Time':
            continue
        left, top, right, bottom = rendered_bbox(s)
        runs = non_empty_runs(s)
        reasons = []
        # 距左16–19cm、距上17.7–18.7cm（以文本框左上角为准）
        if not ((16.0 - tol <= left <= 19.0 + tol) and (17.7 - tol <= top <= 18.7 + tol)):
            reasons.append(f'位置left={left:.2f}/top={top:.2f}不在距左16-19/距上17.7-18.7范围内')
        # 字体为Arial或Calibri
        if not (runs and all(run_font_ok(tr) for tr in runs)):
            reasons.append('字体不是Arial/Calibri')
        # 字号13–15磅
        if not (runs and all(13 <= tr[2] <= 15 for tr in runs)):
            reasons.append(f'字号不在13-15磅: {[tr[2] for tr in runs]}')
        # 加粗
        if not (runs and all(tr[3] for tr in runs)):
            reasons.append('未加粗')
        # 颜色为黑色（办公软件视觉呈黑，含纯黑与深灰）
        if not (runs and all(
                text_run_rgb(tr) is None or is_visually_black(text_run_rgb(tr)) for tr in runs)):
            reasons.append(f'颜色不是黑色: {[tr[4] for tr in runs]}')
        if not reasons:
            return True, f'找到第2页横轴标题"Time": {s.name} left={left:.2f} top={top:.2f}'
        failures.append(f'{s.name}: ' + '；'.join(reasons))
    if failures:
        return False, '找到"Time"但不符合要求：' + ' | '.join(failures[:2])
    return False, '未找到符合要求的单独"Time"标题文本'

def chk_p2_maturity_curve(slides, **_):
    """+1: 第2页顶部成熟度曲线箭头：位于距左6.4cm–29.5cm、距上1.1cm–4.2cm范围内，
    为向右上方延伸的曲线箭头，线宽1.8–3磅，左段为蓝色，中段为青绿色，右段为橙色，
    末端箭头朝右上方。"""
    shapes = slides[1]
    tol = 0.2  # cm，容忍 EMU→cm 取整与屏幕测量误差
    # 曲线由多段线/曲线连接器拼接，取落在区域内、线宽1.8-3磅的线段
    curve_geoms = {'line', 'arc', 'curvedConnector2', 'curvedConnector3',
                   'curvedConnector4', 'curvedConnector5'}
    segs = [s for s in shapes if s.geom in curve_geoms and shape_line_rgb(s)
            and lw_in(s, 1.8, 3.0)
            and shape_in_box(s, 6.4, 29.5, 1.1, 4.2)]

    def seg_color(s): return shape_line_rgb(s)
    blue = [s for s in segs if is_blue(seg_color(s))]
    teal = [s for s in segs if is_teal(seg_color(s))]
    orange = [s for s in segs if is_orange(seg_color(s))]

    reasons = []
    if not segs:
        return False, '未在(距左6.4-29.5,距上1.1-4.2)找到线宽1.8-3磅的曲线线段'
    # 连续曲线：整条应由曲线段(arc/curvedConnector/freeform)构成，
    # 而不是由多段直线首尾拼接。全部 geom='line' 视为"直线段拼接"，不符合。
    non_line_segs = [s for s in segs if s.geom != 'line']
    if not non_line_segs and len(segs) > 1:
        reasons.append(f'不是一条连续的曲线，为{len(segs)}段直线段拼接')
    # 向右上方延伸：整体从左下到右上（x 增大、y 减小）
    seg_sorted = sorted(segs, key=lambda s: s.x)
    lft = seg_sorted[0]; rgt = seg_sorted[-1]
    left_low_y = min(lft.y, lft.y + lft.h)
    right_high_y = min(rgt.y, rgt.y + rgt.h)
    if not (min(s.x for s in segs) <= 6.8 and max(s.right for s in segs) >= 28.8):
        reasons.append('曲线未覆盖距左6.4-29.5范围的整体跨度')
    if not (right_high_y < left_low_y - tol):
        reasons.append(f'整体未向右上方延伸: 左端top={left_low_y:.2f}, 右端top={right_high_y:.2f}')
    # 左段蓝色、中段青绿色、右段橙色
    if not blue: reasons.append('缺少左段蓝色线段')
    if not teal: reasons.append('缺少中段青绿色线段')
    if not orange: reasons.append('缺少右段橙色线段')
    if blue and teal and orange:
        bx = sum(s.x for s in blue) / len(blue)
        tx = sum(s.x for s in teal) / len(teal)
        ox = sum(s.x for s in orange) / len(orange)
        if not (bx < tx < ox):
            reasons.append(f'颜色分段顺序不是蓝-青绿-橙(从左到右): 蓝x={bx:.1f},青绿x={tx:.1f},橙x={ox:.1f}')
    # 末端箭头朝右上方：以办公软件视觉朝向为准——末端箭头形状同时"朝右"且"朝上"
    arrow_segs = [s for s in segs if has_any_arrow(s)]
    # 取最右侧的带箭头线段；其可见箭头端应位于 bbox 右上角(既朝右又朝上)
    end_arrow_ok = any(arrow_points_right(s) and arrow_points_up(s) for s in arrow_segs)
    if not end_arrow_ok:
        reasons.append('末端未找到朝右上方的箭头')

    if not reasons:
        return True, f'找到成熟度曲线箭头: 蓝{len(blue)}段 青绿{len(teal)}段 橙{len(orange)}段'
    return False, '成熟度曲线箭头不符合要求：' + '；'.join(reasons)

def chk_p2_curve_nodes(slides, **_):
    """+1: 第2页顶部三个圆形节点：位于成熟度曲线上方或曲线节点处，分别位于
    距左9.2cm–10.0cm、距左16cm–18.5cm、距左24cm–27cm的水平范围内，
    距上1.5cm–4cm范围内；三个节点为圆形，直径0.4cm–0.8cm，
    颜色分别为蓝色、青绿色、橙色。"""
    shapes = slides[1]
    tol = 0.2   # cm，位置容差
    size_tol = 0.05  # cm，节点宽/高容差
    # 圆形节点：ellipse，直径0.4-0.8cm（近正圆），距上1.5-4cm
    nodes = [s for s in shapes if s.geom == 'ellipse'
             and in_range(s.abs_w, 0.4 - size_tol, 0.8 + size_tol)
             and in_range(s.abs_h, 0.4 - size_tol, 0.8 + size_tol)
             and abs(s.abs_w - s.abs_h) <= 0.15
             and (1.5 - tol <= s.y <= 4.0 + tol)]

    def node_color(s):
        # 节点颜色可能体现在填充或边线（本例外圈白填充+彩色边）
        return shape_fill_rgb(s), shape_line_rgb(s)

    def has_color(s, color_fn):
        f, l = node_color(s)
        return (f and color_fn(f)) or (l and color_fn(l))

    # 采用与 chk_p2_maturity_curve 相同的筛选取出曲线线段，用于位置校验
    curve_geoms = {'line', 'arc', 'curvedConnector2', 'curvedConnector3',
                   'curvedConnector4', 'curvedConnector5'}
    curve_segs = [s for s in shapes if s.geom in curve_geoms and shape_line_rgb(s)
                  and lw_in(s, 1.8, 3.0)
                  and shape_in_box(s, 6.4, 29.5, 1.1, 4.2)]

    def curve_y_at(x):
        """返回成熟度曲线在横坐标 x 处的 y (cm)。取覆盖该 x 的线段做线性插值，
        多段覆盖时取更靠上(更小 y)的一段。"""
        ys = []
        for seg in curve_segs:
            xl = min(seg.x, seg.x + seg.w)
            xr = max(seg.x, seg.x + seg.w)
            if xl - 0.05 <= x <= xr + 0.05 and abs(seg.w) > 1e-6:
                t = (x - seg.x) / seg.w
                ys.append(seg.y + t * seg.h)
        return min(ys) if ys else None

    def node_pos_ok(node):
        """节点位于曲线上方：其可见 bbox 底部应严格高于(y 更小)曲线在该 x 处的 y。
        若节点与线段视觉重叠(即 bbox 底部 y >= 曲线 y)，视为不符。"""
        cx = node.x + node.abs_w / 2
        y_curve = curve_y_at(cx)
        if y_curve is None:
            return False
        # 节点底部：y+h（cm，y 向下增）；容忍 0.02cm 的贴边情况
        return node.y + node.abs_h <= y_curve + 0.02

    # 蓝色节点(距左9.2-10.0)、青绿色节点(距左16-18.5)、橙色节点(距左24-27)
    def group(xl, xr, color_fn):
        cands = [s for s in nodes if xl - tol <= s.x <= xr + tol and has_color(s, color_fn)]
        return cands, [s for s in cands if node_pos_ok(s)]

    blue_cands, blue_ok = group(9.2, 10.0, is_blue)
    teal_cands, teal_ok = group(16.0, 18.5, is_teal)
    orange_cands, orange_ok = group(24.0, 27.0, is_orange)

    reasons = []
    if not blue_cands: reasons.append('缺少距左9.2-10.0、直径0.4-0.8cm的蓝色圆形节点')
    elif not blue_ok:  reasons.append('蓝色节点未位于成熟度曲线上方(与线段重叠)')
    if not teal_cands: reasons.append('缺少距左16-18.5、直径0.4-0.8cm的青绿色圆形节点')
    elif not teal_ok:  reasons.append('青绿色节点未位于成熟度曲线上方(与线段重叠)')
    if not orange_cands: reasons.append('缺少距左24-27、直径0.4-0.8cm的橙色圆形节点')
    elif not orange_ok:  reasons.append('橙色节点未位于成熟度曲线上方(与线段重叠)')
    if not reasons:
        return True, f'找到三个圆形节点: {blue_ok[0].name}, {teal_ok[0].name}, {orange_ok[0].name}'
    return False, '顶部三个圆形节点不符合要求：' + '；'.join(reasons)

def _chk_stage_card(shapes, stage_num, xl,xr,yt,yb, title_color_fn, kw_title, kw_profile, milestones):
    """通用阶段卡片检测"""
    cands = shapes_in(shapes, xl, xr, yt, yb)
    # 外层卡片: 大圆角矩形
    cards = [s for s in cands if s.w>5.5 and s.h>8]
    title_found = any(kw_title.lower() in s.full_text().lower() for s in cands) or \
                  any(kw_title.lower() in s.full_text().lower() for s in shapes)
    profile_found = any(kw_profile.lower() in s.full_text().lower() for s in shapes)
    ms_found = [m for m in milestones if any(str(m) in s.full_text() for s in shapes)]
    num_found = any(str(stage_num).zfill(2) in s.full_text() for s in cands)
    evidence = f'卡片:{len(cards)}, 标题:{"✓" if title_found else "✗"}, profile:{"✓" if profile_found else "✗"}, 里程碑:{ms_found}/{milestones}, 编号:{"✓" if num_found else "✗"}'
    hit = title_found and profile_found and len(ms_found)>=2
    return hit, evidence

def chk_p2_stage1_card(slides, **_):
    """+1: 第2页左侧第一阶段外层卡片：位于距左6.1cm–13.2cm、距上6.1cm–17.2cm范围内，
    形状为圆角矩形，边线为蓝色单实线，线宽1–1.5磅，宽6.5cm–7.2cm，高10.5cm–11.6cm，
    带轻微阴影或白色底板，背景填充为淡蓝色。"""
    return _p2_stage1_outer_card(slides[1], 6.1, 13.2, 6.1, 17.2, '第一阶段')

def _p2_stage1_outer_card(shapes, xl, xr, yt, yb, tag_name, line_fn=is_blue,
                          fill_required=True, fill_fn=None):
    tol = 0.2  # cm，容忍 EMU→cm 取整与屏幕测量误差
    # 默认仅接受淡蓝色（第一阶段：细则要求背景填充为淡蓝色）
    if fill_fn is None:
        fill_fn = lambda rgb: is_pale_blue_fill(rgb) or is_light_blue_fill(rgb)
    failures = []
    for s in shapes:
        # 形状为圆角矩形
        if s.geom != 'roundRect':
            continue
        # 只考察落在指定区域、且尺寸达卡片级的圆角矩形
        left, top, right, bottom = rendered_bbox(s)
        if not (xl - tol <= left and right <= xr + tol
                and yt - tol <= top and bottom <= yb + tol):
            continue
        if not (s.abs_w > 5.0 and s.abs_h > 8.0):
            continue
        reasons = []
        # 宽6.5–7.2cm
        if not in_range(s.abs_w, 6.5 - tol, 7.2 + tol):
            reasons.append(f'宽{s.abs_w:.2f}cm不在6.5-7.2cm')
        # 高10.5–11.6cm
        if not in_range(s.abs_h, 10.5 - tol, 11.6 + tol):
            reasons.append(f'高{s.abs_h:.2f}cm不在10.5-11.6cm')
        # 边线为指定颜色单实线
        line = shape_line_rgb(s)
        if not (line and line_fn(line) and is_solid_line(s)):
            reasons.append(f'边线不是指定颜色单实线: color={s.line_color}, dash={s.line_dash}')
        # 线宽1–1.5磅
        if not lw_in(s, 1.0, 1.5):
            reasons.append(f'线宽{s.line_width_pt}不在1-1.5磅')
        # 背景填充（仅当细则要求填充时校验）
        if fill_required:
            fill = shape_fill_rgb(s)
            if not (fill and fill_fn(fill)):
                reasons.append(f'背景填充不符: {s.fill_color}')
        # 轻微阴影 或 白色底板：任一满足即可
        # a) 卡片自身 spPr/a:effectLst 带阴影
        # b) 卡片下方（同区域，尺寸接近或略大）存在一块白色/近白填充的圆角矩形作为底板
        card_left, card_top, card_right, card_bottom = left, top, right, bottom
        has_backing = False
        backing_reason = ''
        if s.has_shadow:
            has_backing = True
            backing_reason = '带阴影(effectLst/outerShdw等)'
        else:
            for cand in shapes:
                if cand is s: continue
                if cand.geom not in ('roundRect', 'rect'): continue
                cf = shape_fill_rgb(cand)
                if not (cf and is_white(cf)): continue
                cl, ct, cr, cb = rendered_bbox(cand)
                # 底板与卡片高度重合度 ≥ 80%，且底板略大或等大
                overlap_w = max(0.0, min(card_right, cr) - max(card_left, cl))
                overlap_h = max(0.0, min(card_bottom, cb) - max(card_top, ct))
                card_area = max((card_right - card_left) * (card_bottom - card_top), 0.01)
                if overlap_w * overlap_h / card_area >= 0.8 and \
                        (cr - cl) >= (card_right - card_left) - tol and \
                        (cb - ct) >= (card_bottom - card_top) - tol:
                    has_backing = True
                    backing_reason = f'底部存在白色底板 {cand.name}'
                    break
        if not has_backing:
            reasons.append('未检测到轻微阴影(effectLst)或白色底板')
        if not reasons:
            return True, f'找到{tag_name}外层卡片: {s.name} w={s.abs_w:.2f} h={s.abs_h:.2f}；{backing_reason}'
        failures.append(f'{s.name}: ' + '；'.join(reasons))
    if failures:
        return False, f'找到{tag_name}圆角矩形但不符合要求：' + ' | '.join(failures[:2])
    return False, f'未找到符合要求的{tag_name}外层卡片'

def chk_p2_stage1_num(slides, **_):
    """第2页左侧第一阶段编号圆形：位于距左6.1cm–7.8cm、距上4.3cm–6.2cm范围内，
    形状为蓝色圆形，直径1.2cm–1.5cm(±0.2)，内部文本为"01"，文字为白色，
    字号19–21磅，加粗，在序号外围有一个宽度为0.1cm-0.3cm的白色圆环。"""
    return _p2_stage_num_circle(slides[1], 6.1, 7.8, 4.3, 6.2, is_blue, '01', '第一阶段')

def _p2_stage_num_circle(shapes, xl, xr, yt, yb, fill_fn, expected_text, tag_name):
    tol = 0.2  # cm，容忍 EMU→cm 取整与屏幕测量误差
    # 区域内的圆形
    circles = [s for s in shapes if s.geom == 'ellipse'
               and shape_in_box(s, xl, xr, yt, yb)
               and abs(s.abs_w - s.abs_h) <= 0.15]
    reasons = []

    # 编号圆形：填充为指定颜色(蓝)，直径1.2-1.5cm
    color_circle = next((s for s in circles
                         if shape_fill_rgb(s) and fill_fn(shape_fill_rgb(s))
                         and in_range(s.abs_w, 1.2 - tol, 1.5 + tol)), None)
    if color_circle is None:
        reasons.append('未找到直径1.2-1.5cm的蓝色圆形编号')

    # 内部文本"01"：白色，字号19-21磅，加粗
    text = None
    for s in shapes:
        if s.full_text().strip() != expected_text: continue
        if not shape_in_box(s, xl, xr, yt, yb): continue
        text = s
        break
    if text is None:
        reasons.append(f'未找到编号文本"{expected_text}"')
    else:
        runs = non_empty_runs(text)
        if not (runs and all(run_font_ok(tr) for tr in runs)):
            reasons.append('编号文本字体不是Arial/Calibri')
        if not (runs and all(text_run_rgb(tr) is not None and is_white(text_run_rgb(tr)) for tr in runs)):
            reasons.append(f'编号文本不是白色: {[tr[4] for tr in runs]}')
        if not (runs and all(19 <= tr[2] <= 21 for tr in runs)):
            reasons.append(f'编号文本字号不在19-21磅: {[tr[2] for tr in runs]}')
        if not (runs and all(tr[3] for tr in runs)):
            reasons.append('编号文本未加粗')

    # 外围白色圆环：白填充、直径比编号圆形大，环宽0.1-0.3cm
    if color_circle is not None:
        ring = None
        for s in circles:
            if s is color_circle: continue
            f = shape_fill_rgb(s)
            if not (f and is_white(f)): continue
            ring_w = (s.abs_w - color_circle.abs_w) / 2
            if 0.1 - 0.05 <= ring_w <= 0.3 + 0.05:
                ring = s
                break
        if ring is None:
            reasons.append('未找到宽0.1-0.3cm的外围白色圆环')

    if not reasons:
        return True, f'{tag_name}编号圆形"{expected_text}"符合要求'
    return False, f'{tag_name}编号圆形不符合要求：' + '；'.join(reasons)

def _p2_runs_ok(s, size_lo, size_hi, color_fn=None, require_bold=None):
    runs = non_empty_runs(s)
    if not runs:
        return False
    for tr in runs:
        rgb = text_run_rgb(tr)
        if not run_font_ok(tr): return False
        if not (size_lo <= tr[2] <= size_hi): return False
        if require_bold is not None and tr[3] != require_bold: return False
        if color_fn and not (rgb and color_fn(rgb)): return False
    return True

def _p2_stage_title(shapes, xl, xr, yt, yb, color_fn, text_parts, label):
    """第X阶段标题条：胶囊形圆角矩形，左侧与编号圆形拼接，
    宽5-5.5cm、高1.3-1.7cm（以办公软件所选中的形状本体尺寸为准），
    填充为指定色，内部文本为指定内容，白色、字号11-13磅、加粗。"""
    tol = 0.2  # cm，容忍 EMU→cm 取整与屏幕测量误差
    reasons = []
    # 胶囊标题条：圆角矩形，填充为指定色，位于该区域
    bar = None
    bar_failures = []
    for s in shapes:
        if s.geom != 'roundRect': continue
        if not shape_in_box(s, xl, xr, yt, yb): continue
        fill = shape_fill_rgb(s)
        if not (fill and color_fn(fill)):
            continue
        # 高1.3-1.7cm
        if not in_range(s.abs_h, 1.3 - tol, 1.7 + tol):
            bar_failures.append(f'{s.name}高{s.abs_h:.2f}cm不在1.3-1.7cm')
            continue
        bar = s
        break

    if bar is None:
        reasons.append('未找到符合颜色/高度的胶囊标题条；' + ' | '.join(bar_failures[:2]))
    else:
        # 左侧和编号圆形拼接：需在同一区域内找到近正圆的徽章椭圆
        badge = [s for s in shapes if s.geom == 'ellipse'
                 and shape_in_box(s, xl, xr, yt, yb)
                 and abs(s.abs_w - s.abs_h) <= 0.15
                 and s.x + s.abs_w <= bar.x + bar.abs_w]
        if not badge:
            reasons.append('未找到与标题条拼接的编号圆形')
        # 胶囊本体宽度：以办公软件里选中形状显示的宽度为准
        if not in_range(bar.abs_w, 5.0 - tol, 5.5 + tol):
            reasons.append(f'胶囊宽度{bar.abs_w:.2f}cm不在5-5.5cm')

    # 内部文本为指定内容，白色、字号11-13磅、加粗
    title_texts = [s for s in shapes if shape_in_box(s, xl, xr, yt, yb) and s.texts
                   and any(part in s.full_text() for part in text_parts)]
    region_text = ' '.join(s.full_text() for s in title_texts)
    if not all(part in region_text for part in text_parts):
        reasons.append(f'标题文本缺失，应含{text_parts}，实际="{region_text}"')
    else:
        for s in title_texts:
            runs = non_empty_runs(s)
            if not (runs and all(run_font_ok(tr) for tr in runs)):
                reasons.append(f'"{s.full_text()}"字体不是Arial/Calibri')
            if not (runs and all(text_run_rgb(tr) is not None and is_white(text_run_rgb(tr)) for tr in runs)):
                reasons.append(f'"{s.full_text()}"文字不是白色')
            if not (runs and all(11 <= tr[2] <= 13 for tr in runs)):
                reasons.append(f'"{s.full_text()}"字号不在11-13磅: {[tr[2] for tr in runs]}')
            if not (runs and all(tr[3] for tr in runs)):
                reasons.append(f'"{s.full_text()}"未加粗')

    if not reasons:
        return True, f'{label}标题条符合要求'
    return False, f'{label}标题条不符合要求：' + '；'.join(reasons[:4])

def _p2_milestones(shapes, xl, xr, yt, yb, color_fn, years, label):
    """第X阶段里程碑列表：标题"Representative milestones"(字号10-12磅)，
    标题下方带对应色横线，列表含各年份及英文说明；年份加粗，
    正文黑色、字号10-12磅，年份左侧带对应色圆点。"""
    tol = 0.2  # cm，容忍 EMU→cm 取整与屏幕测量误差
    def inbox(s): return shape_in_box(s, xl, xr, yt, yb)
    reasons = []

    # 标题"Representative milestones"，字号10-12磅
    title = next((s for s in shapes if inbox(s)
                  and s.full_text().strip() == 'Representative milestones'), None)
    if title is None:
        reasons.append('缺少标题"Representative milestones"')
    else:
        t_runs = non_empty_runs(title)
        if not (t_runs and all(10 <= tr[2] <= 12 for tr in t_runs)):
            reasons.append(f'标题字号不在10-12磅: {[tr[2] for tr in t_runs]}')
        # 标题下方带对应色(蓝色)横线
        line = next((s for s in shapes if inbox(s) and s.geom == 'line'
                     and s.abs_w > 2 and s.abs_h < 0.2
                     and shape_line_rgb(s) and color_fn(shape_line_rgb(s))
                     and min(s.y, s.y + s.h) >= title.bottom - 0.6), None)
        if line is None:
            reasons.append('标题下方缺少对应颜色横线')

    # 列表年份：含全部指定年份，年份加粗
    year_shapes = [s for s in shapes if inbox(s) and s.full_text().strip() in years]
    found_years = sorted({s.full_text().strip() for s in year_shapes})
    if found_years != sorted(years):
        reasons.append(f'年份缺失: 实际{found_years}，应为{sorted(years)}')
    else:
        for s in year_shapes:
            runs = non_empty_runs(s)
            if not (runs and all(tr[3] for tr in runs)):
                reasons.append(f'年份"{s.full_text().strip()}"未加粗')

    # 对应英文说明(正文)：黑色，字号10-12磅
    # 仅取真正的说明文本（含英文单词、长度>3），排除"›"等装饰/导航按钮
    def is_desc(s):
        t = s.full_text().strip()
        return len(t) >= 3 and any(ch.isascii() and ch.isalpha() for ch in t)
    body_shapes = [s for s in shapes if inbox(s) and s.texts and s.full_text().strip()
                   and s.full_text().strip() not in years
                   and s.full_text().strip() != 'Representative milestones'
                   and is_desc(s)]
    if not body_shapes:
        reasons.append('缺少年份对应的英文说明正文')
    else:
        for s in body_shapes:
            runs = non_empty_runs(s)
            if not (runs and all(run_font_ok(tr) for tr in runs)):
                reasons.append(f'正文"{s.full_text()[:12]}"字体不是Arial/Calibri')
            if not (runs and all(10 <= tr[2] <= 12 for tr in runs)):
                reasons.append(f'正文"{s.full_text()[:12]}"字号不在10-12磅: {[tr[2] for tr in runs]}')
            if not (runs and all(
                    text_run_rgb(tr) is None or is_visually_black(text_run_rgb(tr)) for tr in runs)):
                reasons.append(f'正文"{s.full_text()[:12]}"颜色不是黑色')

    # 年份左侧对应色(蓝色)圆点，数量不少于年份数
    dots = [s for s in shapes if inbox(s) and s.geom == 'ellipse'
            and in_range(s.abs_w, 0.12 - 0.05, 0.35 + 0.05)
            and abs(s.abs_w - s.abs_h) <= 0.1
            and shape_fill_rgb(s) and color_fn(shape_fill_rgb(s))]
    if len(dots) < len(years):
        reasons.append(f'年份左侧圆点不足: {len(dots)}/{len(years)}')

    if not reasons:
        return True, f'{label}里程碑列表符合要求'
    return False, f'{label}里程碑列表不符合要求：' + '；'.join(reasons[:5])

def _p2_profile(shapes, xl, xr, yt, yb, fill_fn, color_fn, keywords, label,
                icon_kind='clipboard_gear'):
    """第X阶段 Stage profile 框：浅色圆角矩形，宽6.2-6.6cm、高3-3.4cm，
    放在阶段卡片内下方，含圆形图标背景、指定语义的图标、
    "Stage profile"(10-12磅)及说明关键词文本(8-10磅)。

    icon_kind: 图标语义
      · 'clipboard_gear'  第一阶段 剪贴板齿轮
      · 'charging_pile'   第二阶段 充电桩
      · 'display_chart'   第三阶段 显示器图表
    """
    tol = 0.2  # cm，容忍 EMU→cm 取整与屏幕测量误差
    def inbox(s): return shape_in_box(s, xl, xr, yt, yb)
    reasons = []

    # 浅色圆角矩形：宽6.2-6.6cm、高3-3.4cm、填充为指定浅色
    box = None
    box_failures = []
    for s in shapes:
        if s.geom != 'roundRect': continue
        if not inbox(s): continue
        if not (s.abs_w > 4 and s.abs_h > 2): continue
        rf = []
        if not in_range(s.abs_w, 6.2 - tol, 6.6 + tol): rf.append(f'宽{s.abs_w:.2f}cm不在6.2-6.6cm')
        if not in_range(s.abs_h, 3.0 - tol, 3.4 + tol): rf.append(f'高{s.abs_h:.2f}cm不在3-3.4cm')
        fill = shape_fill_rgb(s)
        if not (fill and (fill_fn(fill) or is_white(fill))): rf.append(f'填充色不符: {s.fill_color}')
        if not rf:
            box = s; break
        box_failures.append(f'{s.name}: ' + '；'.join(rf))
    if box is None:
        reasons.append('未找到符合尺寸/颜色的profile圆角矩形；' + ' | '.join(box_failures[:2]))

    # 圆形图标背景（profile 框内的填充圆）
    icon_bg = next((s for s in shapes if inbox(s) and s.geom == 'ellipse'
                    and shape_fill_rgb(s)), None)
    if icon_bg is None:
        reasons.append('缺少圆形图标背景')

    # 图标语义校验：优先命名/alt text；否则几何组合兜底
    ICON_HINTS = {
        'clipboard_gear': ('clipboard', 'clip-board', 'gear', 'cog', 'settings',
                           '剪贴板', '齿轮', '设置'),
        'charging_pile':  ('charging', 'charger', 'ev-charge', 'ev_station', 'pile',
                           'plug', 'battery', '充电', '充电桩', '插头', '电池'),
        'display_chart':  ('display', 'monitor', 'screen', 'chart', 'graph',
                           'dashboard', 'stat', '显示器', '屏幕', '图表', '仪表盘'),
    }

    if icon_bg is not None:
        # 收集图标背景圆内的部件（中心落在其包围盒内，非文本，非背景圆本身，非 profile 大框）
        ic_left, ic_top, ic_right, ic_bottom = rendered_bbox(icon_bg)
        icon_parts = []
        for s in shapes:
            if s is icon_bg or s is box: continue
            if s.texts: continue
            cx, cy = s.x + s.w / 2, s.y + s.h / 2
            if ic_left - tol <= cx <= ic_right + tol and ic_top - tol <= cy <= ic_bottom + tol:
                icon_parts.append(s)

        # 命名/alt text 匹配
        blob = ' '.join((getattr(p, 'name', '') + ' ' + getattr(p, 'descr', '')).lower()
                        for p in icon_parts)
        hits = ICON_HINTS.get(icon_kind, ())
        sem_ok = False
        sem_reason = ''
        if blob.strip() and any(h.lower() in blob for h in hits):
            sem_ok = True
            sem_reason = f'{icon_kind}: 命名/alt text命中'
        else:
            # 几何组合兜底
            rects = [p for p in icon_parts if p.geom in ('rect', 'roundRect')]
            ellips = [p for p in icon_parts if p.geom == 'ellipse']
            tris  = [p for p in icon_parts if p.geom in ('triangle', 'rtTriangle')]
            lines  = [p for p in icon_parts
                      if p.shape_type == 'cxnSp' or p.geom in ('line', 'straightConnector1')]
            gears  = [p for p in icon_parts if p.geom in ('gear6', 'gear9')]
            plaques = [p for p in icon_parts if p.geom == 'plaque']  # 圆矩形卡片
            n = len(icon_parts)

            if icon_kind == 'clipboard_gear':
                # 剪贴板：偏高的圆角矩形/矩形；齿轮：gear 几何 或 圆+周围小矩形/线
                has_board = any(p.abs_h >= p.abs_w * 0.9 for p in rects)
                has_gear = bool(gears) or (ellips and (len(ellips) >= 2 or len(rects) + len(lines) >= 2))
                if (has_board and has_gear) or n >= 2:
                    sem_ok = True; sem_reason = 'clipboard_gear: 剪贴板+齿轮几何'
                else:
                    sem_reason = (f'clipboard_gear 缺剪贴板矩形或齿轮 '
                                  f'(rects={len(rects)}, ellips={len(ellips)}, gears={len(gears)}, '
                                  f'lines={len(lines)}, parts={n})')

            elif icon_kind == 'charging_pile':
                # 充电桩：竖矩形主体 + 顶部小矩形/短线（插头/电线）
                pole = [p for p in rects if p.abs_h >= p.abs_w * 1.2]
                top_bits = [p for p in icon_parts if p is not None and (p.geom in ('rect', 'roundRect')
                            or p.shape_type == 'cxnSp')]
                if pole and (len(top_bits) >= 2 or lines):
                    sem_ok = True; sem_reason = 'charging_pile: 竖矩形主体+顶部部件'
                elif n >= 2 and (rects or lines):
                    sem_ok = True; sem_reason = 'charging_pile: 多部件组合'
                else:
                    sem_reason = (f'charging_pile 缺竖矩形主体或顶部部件 '
                                  f'(pole={len(pole)}, rects={len(rects)}, lines={len(lines)}, parts={n})')

            elif icon_kind == 'display_chart':
                # 显示器+图表：宽扁矩形屏幕 + 底座（细矩形/三角）+ 屏幕内图表（折线/柱）
                screens = [p for p in rects if p.abs_w >= p.abs_h * 1.2]
                bars = [p for p in rects if p.abs_h > p.abs_w and p.abs_w < 0.4]
                has_screen = bool(screens) or bool(plaques)
                has_chart = bool(lines) or len(bars) >= 2
                if has_screen and (has_chart or n >= 2):
                    sem_ok = True; sem_reason = 'display_chart: 屏幕+图表元素'
                elif n >= 3:
                    sem_ok = True; sem_reason = 'display_chart: 多部件组合'
                else:
                    sem_reason = (f'display_chart 缺屏幕或图表 '
                                  f'(screens={len(screens)}, lines={len(lines)}, bars={len(bars)}, parts={n})')

        if not sem_ok:
            reasons.append(f'图标语义不符({sem_reason})')

    # "Stage profile" 文本，字号10-12磅
    title = next((s for s in shapes if inbox(s) and s.full_text().strip() == 'Stage profile'), None)
    if title is None:
        reasons.append('缺少"Stage profile"文本')
    else:
        t_runs = non_empty_runs(title)
        if not (t_runs and all(10 <= tr[2] <= 12 for tr in t_runs)):
            reasons.append(f'"Stage profile"字号不在10-12磅: {[tr[2] for tr in t_runs]}')

    # 说明关键词文本，字号8-10磅
    all_text = ' '.join(s.full_text() for s in shapes if inbox(s)).lower()
    if not all(k in all_text for k in keywords):
        reasons.append(f'profile说明文本缺失关键词: {keywords}')
    else:
        detail_texts = [s for s in shapes if inbox(s)
                        and any(k in s.full_text().lower() for k in keywords)]
        for s in detail_texts:
            runs = non_empty_runs(s)
            if not (runs and all(8 <= tr[2] <= 10 for tr in runs)):
                reasons.append(f'说明文本字号不在8-10磅: {[tr[2] for tr in runs]}')

    if not reasons:
        return True, f'{label}Stage profile框符合要求'
    return False, f'{label}Stage profile框不符合要求：' + '；'.join(reasons[:4])

def _p2_stage_card(shapes, xl, xr, yt, yb, line_fn, fill_fn, label):
    cands = shapes_in(shapes, xl, xr, yt, yb)
    failures = []
    for s in cands:
        if s.geom != 'roundRect' or s.h < 8:
            continue
        line = shape_line_rgb(s); fill = shape_fill_rgb(s)
        reasons = []
        if not bbox_in_box(s, xl, xr, yt, yb): reasons.append('位置不在范围内')
        if not (6.5 <= s.w <= 7.2): reasons.append(f'宽{s.w:.2f}cm不在6.5-7.2cm')
        if not (10.5 <= s.h <= 11.6): reasons.append(f'高{s.h:.2f}cm不在10.5-11.6cm')
        if not (line and line_fn(line) and is_solid_line(s) and lw_in(s, 1.0, 1.5)): reasons.append(f'边线不符: {s.line_color}, lw={s.line_width_pt}')
        if not (fill and fill_fn(fill)): reasons.append(f'背景填充不符: {s.fill_color}')
        if not reasons: return True, f'{label}外层卡片符合要求'
        failures.append(f'{s.name}: ' + '；'.join(reasons))
    return False, f'{label}外层卡片不符合要求：' + ' | '.join(failures[:2])

def _p2_stage_num(shapes, xl, xr, yt, yb, color_fn, expected_text, label):
    cands = shapes_in(shapes, xl, xr, yt, yb)
    circles = [s for s in cands if s.geom == 'ellipse']
    colored = [s for s in circles if 1.2 <= s.w <= 1.5 and 1.2 <= s.h <= 1.5 and shape_fill_rgb(s) and color_fn(shape_fill_rgb(s))]
    text = next((s for s in cands if s.full_text().strip() == expected_text), None)
    text_ok = bool(text and _p2_runs_ok(text, 19, 21, is_white, True))
    ring_ok = any(s.geom == 'ellipse' and shape_fill_rgb(s) and is_white(shape_fill_rgb(s)) and 1.4 <= s.w <= 1.8 for s in circles)
    reasons = []
    if not colored: reasons.append(f'未找到直径1.2-1.5cm的{label}色圆形编号')
    if not text_ok: reasons.append(f'编号文本不是{expected_text}或字号/白色/加粗不符')
    if not ring_ok: reasons.append('未找到外围白色圆环')
    if not reasons: return True, f'{label}编号圆形符合要求'
    return False, f'{label}编号圆形不符合要求：' + '；'.join(reasons)

def chk_p2_stage1_title(slides, **_):
    return _p2_stage_title(slides[1], 6.5, 13.0, 4.3, 6.2, is_blue, ['Pilot Exploration Stage', '(1998–2008)'], '第一阶段')

def chk_p2_stage1_milestones(slides, **_):
    return _p2_milestones(slides[1], 6.2, 13.1, 6.5, 12.4, is_blue, ['1999','2002','2005','2007'], '第一阶段')

def chk_p2_stage1_profile(slides, **_):
    return _p2_profile(slides[1], 6.0, 13.0, 13.2, 16.8, is_pale_blue_fill, is_blue,
                       ['technology testing', 'limited routes', 'institutional learning'],
                       '第一阶段', icon_kind='clipboard_gear')

def chk_p2_stage2_card(slides, **_):
    """+1: 第2页中间第二阶段外层卡片：位于距左13.5cm–21cm、距上6.1cm–17.2cm范围内，
    形状为圆角矩形，边线为青绿色单实线，线宽1–1.5磅，宽6.5cm–7.2cm，高10.5cm–11.6cm，
    带轻微阴影或白色底板，背景填充为淡绿色。"""
    return _p2_stage1_outer_card(slides[1], 13.5, 21.0, 6.1, 17.2, '第二阶段',
                                 line_fn=is_teal, fill_required=True,
                                 fill_fn=is_pale_green_fill)

def chk_p2_stage2_num(slides, **_):
    """+1: 第2页中间第二阶段编号圆形：位于距左13.5cm–15.5cm、距上4.3cm–6.2cm范围内，
    形状为青绿色圆形，直径1.2cm–1.5cm(±0.2)，内部文本为"02"，文字为白色，
    字号19–21磅，加粗，在序号外围有一个宽度为0.1cm-0.3cm的白色圆环。"""
    return _p2_stage_num_circle(slides[1], 13.5, 15.5, 4.3, 6.2, is_teal, '02', '第二阶段')

def chk_p2_stage2_title(slides, **_):
    """+3: 第2页中间第二阶段标题条：位于距左14.3cm–21cm、距上4.3cm–6.2cm范围内，
    形状类似胶囊，宽5cm-5.5cm，高1.3cm-1.7cm，左侧和第二阶段圆形编号拼接，填充为青绿色，
    内部文本为"Expansion & Standardization Stage (2009–2018)"，文字为白色，字号11–13磅，加粗。"""
    return _p2_stage_title(slides[1], 14.3, 21.0, 4.3, 6.2, is_teal,
                           ['Expansion & Standardization Stage', '(2009–2018)'], '第二阶段')

def chk_p2_stage2_milestones(slides, **_):
    """+3: 第2页中间第二阶段里程碑列表：位于距左13.5cm–21.0cm、距上6.5cm–12.4cm范围内，
    标题为"Representative milestones"，标题字体大小为10-12磅，标题下方带有青绿色横线，
    列表包含2009、2011、2013、2016、2018及对应英文说明；年份加粗，正文为黑色，
    字号10–12磅，年份左侧带有青绿色圆点。"""
    return _p2_milestones(slides[1], 13.5, 21.0, 6.5, 12.4, is_teal,
                          ['2009', '2011', '2013', '2016', '2018'], '第二阶段')

def chk_p2_stage2_profile(slides, **_):
    """+3: 第2页中间第二阶段Stage profile框：位于距左13.5cm–21.0cm、距上13.2cm–16.8cm范围内，
    宽6.2cm-6.6cm，高3cm-3.4cm，形状为浅青绿色或白色圆角矩形，放置在第二阶段外层卡片内下方，
    包含浅青绿色圆形图标背景、充电桩图标、"Stage profile"(10-12磅)和
    "policy support, scale-up, infrastructure coordination"(8-10磅)等文本。"""
    return _p2_profile(slides[1], 13.5, 21.0, 13.2, 16.8,
                       lambda rgb: is_teal(rgb) or is_white(rgb), is_teal,
                       ['policy support', 'scale-up', 'infrastructure coordination'],
                       '第二阶段', icon_kind='charging_pile')

def chk_p2_stage3_card(slides, **_):
    """+1: 第2页右侧第三阶段外层卡片：位于距左21.5cm–30cm、距上6.1cm–17.2cm范围内，
    形状为白色圆角矩形，边线为橙色单实线，线宽1–1.5磅，宽6.5cm–7.2cm，高10.5cm–11.6cm，
    带轻微阴影或白色底板，背景填充为淡橘色。"""
    return _p2_stage1_outer_card(slides[1], 21.5, 30.0, 6.1, 17.2, '第三阶段',
                                 line_fn=is_orange, fill_required=True,
                                 fill_fn=is_pale_orange_fill)

def chk_p2_stage3_num(slides, **_):
    """+1: 第2页右侧第三阶段编号圆形：位于距左21.5cm–23.5cm、距上4.3cm–6.2cm范围内，
    形状为橙色圆形，直径1.2cm–1.5cm(±0.2)，内部文本为"02"（按 rubric 原文），
    文字为白色，字号19–21磅，加粗，在序号外围有一个宽度为0.1cm-0.3cm的白色圆环。"""
    return _p2_stage_num_circle(slides[1], 21.5, 23.5, 4.3, 6.2, is_orange, '02', '第三阶段')

def chk_p2_stage3_title(slides, **_):
    """+3: 第2页右侧第三阶段标题条：位于距左22.5cm–29.5cm、距上4.3cm–6.2cm范围内，
    形状类似胶囊，宽5cm-5.5cm，高1.3cm-1.7cm，左侧和第三阶段圆形编号拼接，填充为橙色，
    内部文本为"Integrated Optimization Stage (2019–Present)"，文字为白色，字号11–13磅，加粗。"""
    return _p2_stage_title(slides[1], 22.5, 29.5, 4.3, 6.2, is_orange,
                           ['Integrated Optimization Stage', '(2019–Present)'], '第三阶段')

def chk_p2_stage3_milestones(slides, **_):
    """+3: 第2页右侧第三阶段里程碑列表：位于距左21.5cm–29.5cm、距上6.5cm–12.4cm范围内，
    标题为"Representative milestones"，标题字体大小为10-12磅，标题下方带有橙色横线，
    列表包含2019、2021、2023、2024、2025及对应英文说明；年份加粗，正文为黑色，
    字号10–12磅，年份左侧带有橙色圆点。"""
    return _p2_milestones(slides[1], 21.5, 29.5, 6.5, 12.4, is_orange,
                          ['2019', '2021', '2023', '2024', '2025'], '第三阶段')

def chk_p2_stage3_profile(slides, **_):
    """+3: 第2页右侧第三阶段Stage profile框：位于距左21.5cm–29.5cm、距上13.2cm–16.8cm范围内，
    宽6.2cm-6.6cm，高3cm-3.4cm，放置在第三阶段外层卡片内下方，形状为浅橙色圆角矩形，
    包含浅橙色圆形图标背景、显示器图表图标、"Stage profile"(10-12磅)和
    "market maturity, efficiency optimization, system integration"(8-10磅)等文本。"""
    return _p2_profile(slides[1], 21.5, 29.5, 13.2, 16.8,
                       lambda rgb: is_pale_orange_fill(rgb) or is_white(rgb), is_orange,
                       ['market maturity', 'efficiency optimization', 'system integration'],
                       '第三阶段', icon_kind='display_chart')

def chk_p2_left_vline(slides, **_):
    """+1: 第2页左中竖向虚线分隔线：位于距左13.3cm–13.6cm、距上4.0cm–17.2cm范围内，
    为灰色竖向虚线，线宽0.5–1磅，线条角度水平方向90度。"""
    return _p2_divider_vline(slides[1], 13.3, 13.6, 4.0, 17.2, '左中')

def _p2_divider_vline(shapes, xl, xr, yt, yb, tag_name):
    tol = 0.2  # cm，容忍 EMU→cm 取整与屏幕测量误差
    failures = []
    for s in shapes:
        if s.shape_type not in ('sp', 'cxnSp'):
            continue
        # 距左xl-xr（竖线 x 恒定，取左边界）
        left = min(s.x, s.x + s.w)
        if not (xl - tol <= left <= xr + tol):
            continue
        top = min(s.y, s.y + s.h)
        bottom = max(s.y, s.y + s.h)
        reasons = []
        # 距上yt-yb范围内
        if not (yt - tol <= top and bottom <= yb + tol):
            reasons.append(f'位置top={top:.2f}/bottom={bottom:.2f}不在距上{yt}-{yb}范围内')
        # 灰色
        line = shape_line_rgb(s)
        if not (line and is_strict_gray(line)):
            reasons.append(f'不是灰色: {s.line_color}')
        # 虚线
        if not is_dashed(s):
            reasons.append(f'不是虚线: dash={s.line_dash}')
        # 线宽0.5-1磅
        if not lw_in(s, 0.5, 1.0):
            reasons.append(f'线宽{s.line_width_pt}不在0.5-1磅')
        # 竖向 / 线条角度约90度
        if not (85 <= line_angle_from_horizontal(s) <= 95):
            reasons.append(f'线条角度{line_angle_from_horizontal(s):.0f}°不约90度')
        if not reasons:
            return True, f'找到{tag_name}灰色竖向虚线分隔线: {s.name} lw={s.line_width_pt}'
        failures.append(f'{s.name}: ' + '；'.join(reasons))
    if failures:
        return False, f'{tag_name}竖向虚线分隔线不符合要求：' + ' | '.join(failures[:3])
    return False, f'未在(距左{xl}-{xr},距上{yt}-{yb})找到{tag_name}灰色竖向虚线分隔线'

def chk_p2_right_vline(slides, **_):
    """+1: 第2页右中竖向虚线分隔线：位于距左21cm–21.5cm、距上4.0cm–17.2cm范围内，
    为灰色竖向虚线，线宽0.5–1磅，线条角度水平方向90度。"""
    return _p2_divider_vline(slides[1], 21.0, 21.5, 4.0, 17.2, '右中')

def chk_p2_stage_buttons(slides, **_):
    """+1: 第2页阶段间圆形箭头按钮组：分别位于距左12.5cm–14.2cm、距上9.5cm–11cm范围内
    和距左20.5cm–22cm、距上9.5cm–11cm范围内；两个按钮为圆形，从左到右填充分别为
    蓝色或青绿色，内部为白色向右箭头，箭头水平方向0度。"""
    shapes = slides[1]

    def is_circular(s):
        # 圆形：ellipse 工具且宽高近似相等（办公软件里呈正圆，非拉长椭圆）
        if s.abs_w <= 0 or s.abs_h <= 0:
            return False
        return min(s.abs_w, s.abs_h) / max(s.abs_w, s.abs_h) >= 0.8

    def check_button(xl, xr, yt, yb, color_fn, tag):
        # 圆形按钮：ellipse 且为正圆，填充为指定色
        circle = next((s for s in shapes if s.geom == 'ellipse' and is_circular(s)
                       and shape_in_box(s, xl, xr, yt, yb)
                       and shape_fill_rgb(s) and color_fn(shape_fill_rgb(s))), None)
        if circle is None:
            return f'{tag}未找到指定填充色的圆形按钮'
        # 内部白色向右箭头（箭头图形或白色右向指示符），水平0度
        arrow = None
        for s in shapes:
            if not shape_in_box(s, xl, xr, yt, yb): continue
            if s is circle: continue
            cx, cy = s.x + s.w / 2, s.y + s.h / 2
            if not (circle.left <= cx <= circle.right and circle.top <= cy <= circle.bottom):
                continue
            # 白色向右箭头：可能是一体箭头图形(rightArrow/带端点箭头且朝右)，
            # 或白色右向箭头字符(›/>)文本
            runs = non_empty_runs(s)
            is_white_text_arrow = (runs
                                   and all(text_run_rgb(tr) is not None and is_white(text_run_rgb(tr)) for tr in runs)
                                   and s.full_text().strip() in ('›', '>', '❯', '▶', '►'))
            is_white_shape_arrow = (is_arrow_shape(s) and arrow_points_right(s)
                                    and shape_fill_rgb(s) and is_white(shape_fill_rgb(s)))
            if not (is_white_text_arrow or is_white_shape_arrow):
                continue
            if not near_angle(s, 0, 5):  # 箭头水平方向0度
                continue
            arrow = s
            break
        if arrow is None:
            return f'{tag}圆形内未找到白色水平向右箭头'
        return None

    # 填充色：按 rubric "从左到右分别为蓝色或青绿色" 的位置对应关系严格校验——
    # 左按钮为蓝色，右按钮为青绿色。
    left_err = check_button(12.5, 14.2, 9.5, 11.0, is_blue, '左按钮(蓝色)')
    right_err = check_button(20.5, 22.0, 9.5, 11.0, is_teal, '右按钮(青绿色)')
    errs = [e for e in (left_err, right_err) if e]
    if not errs:
        return True, '两个阶段间圆形箭头按钮均符合要求'
    return False, '阶段间圆形箭头按钮不符合要求：' + '；'.join(errs)

def chk_layout(slides, page_w, page_h, **_):
    """+3: 第1页和第2页整体排版：
    (1) 两页所有文本位于页面可视范围内；
    (2) 主要图形没有超出页面边界；
    (3) 箭头不穿过主要正文；
    (4) 卡片、椭圆、曲线、坐标轴之间无明显重叠导致无法阅读的问题。
    所有几何均按办公软件渲染(旋转后的可见包围盒/线段实际端点)判断。"""
    ARROW_GEOMS = ARROW_GEOMS_ALL
    TOL = 0.3          # 页面边界容差(cm)，容忍渲染取整误差
    TEXT_INSET = 0.2   # 正文框内缩(cm)，避免坐标轴标签仅边缘相切被误判为"穿过"
    STRUCT_MIN_AREA = 1.5  # 结构元素最小面积(cm²)，小于此为装饰/图标零件，不计重叠

    def seg_endpoints(s):
        """线条(含箭头)两端点(cm)，考虑形状旋转，与办公软件所见一致。"""
        cx, cy = s.x + s.w / 2, s.y + s.h / 2
        th = math.radians(s.rotation_deg())
        pts = []
        for px, py in ((s.x, s.y), (s.x + s.w, s.y + s.h)):
            dx, dy = px - cx, py - cy
            rx = cx + dx * math.cos(th) - dy * math.sin(th)
            ry = cy + dx * math.sin(th) + dy * math.cos(th)
            pts.append((rx, ry))
        return pts[0], pts[1]

    def seg_cross_rect(p1, p2, r):
        """线段是否穿过矩形内部(Liang-Barsky裁剪)。r=(l,t,r,b)。"""
        x1, y1 = p1; x2, y2 = p2
        dx, dy = x2 - x1, y2 - y1
        ps = [-dx, dx, -dy, dy]
        qs = [x1 - r[0], r[2] - x1, y1 - r[1], r[3] - y1]
        u1, u2 = 0.0, 1.0
        for pi, qi in zip(ps, qs):
            if pi == 0:
                if qi < 0: return False
            else:
                t = qi / pi
                if pi < 0: u1 = max(u1, t)
                else:      u2 = min(u2, t)
        return u1 <= u2

    def bb_area(b): return max(0, b[2]-b[0]) * max(0, b[3]-b[1])
    def bb_overlap(a, b):
        return (max(0, min(a[2], b[2]) - max(a[0], b[0]))
                * max(0, min(a[3], b[3]) - max(a[1], b[1])))

    issues = []
    for idx, shapes in enumerate(slides[:2]):
        pg = idx + 1
        texts = [s for s in shapes if non_empty_runs(s)]

        # (1) 所有文本位于页面可视范围内
        for s in texts:
            l, t, r, b = rendered_bbox(s)
            if l < -TOL or t < -TOL or r > page_w + TOL or b > page_h + TOL:
                issues.append(f'第{pg}页文本"{s.full_text()[:12]}"超出页面可视范围')

        # (2) 主要图形没有超出页面边界(主要图形=非文本且面积较大，或较长坐标轴线)
        for s in shapes:
            if non_empty_runs(s): continue
            if not (s.abs_w * s.abs_h >= 3.0 or max(s.abs_w, s.abs_h) >= 5.0):
                continue
            l, t, r, b = rendered_bbox(s)
            if l < -TOL or t < -TOL or r > page_w + TOL or b > page_h + TOL:
                issues.append(f'第{pg}页主要图形{s.name}超出页面边界')

        # (3) 箭头不穿过主要正文
        arrows = [s for s in shapes if has_any_arrow(s) or (s.geom in ARROW_GEOMS)]
        for ar in arrows:
            is_block = ar.geom in ARROW_GEOMS
            ab = rendered_bbox(ar) if is_block else None
            seg = None if is_block else seg_endpoints(ar)
            for tx in texts:
                l, t, r, b = rendered_bbox(tx)
                ri = (l + TEXT_INSET, t + TEXT_INSET, r - TEXT_INSET, b - TEXT_INSET)
                if ri[0] >= ri[2] or ri[1] >= ri[3]:
                    continue
                if is_block:
                    ta = bb_area(ri)
                    hit = ta > 0 and bb_overlap(ab, ri) / ta > 0.25
                else:
                    hit = seg_cross_rect(seg[0], seg[1], ri)
                if hit:
                    issues.append(f'第{pg}页箭头{ar.name}穿过正文"{tx.full_text()[:12]}"')
                    break

        # (4) 卡片、椭圆、曲线、坐标轴之间无明显重叠导致无法阅读
        #     "无法阅读"的实质是带填充的面状元素(卡片roundRect/椭圆/矩形)相互遮挡；
        #     线条、曲线、坐标轴为细描边，包围盒相交不构成遮挡，不计入。
        #     尺寸悬殊(包含关系)与同位叠加(重复描边=同一视觉元素)也不算问题。
        FILLED_GEOMS = {'roundRect', 'rect', 'ellipse'}
        struct = [s for s in shapes if not non_empty_runs(s)
                  and s.geom in FILLED_GEOMS and s.fill_color]
        boxes = [(s, rendered_bbox(s)) for s in struct]
        boxes = [(s, bb) for s, bb in boxes if bb_area(bb) >= STRUCT_MIN_AREA]
        for i, (a, ba) in enumerate(boxes):
            aa = bb_area(ba)
            for c, bc in boxes[i+1:]:
                ac = bb_area(bc)
                ratio = aa / ac
                if ratio < 0.4 or ratio > 2.5:      # 尺寸悬殊=包含关系
                    continue
                if (abs((ba[0]+ba[2])/2 - (bc[0]+bc[2])/2) < 0.1
                        and abs((ba[1]+ba[3])/2 - (bc[1]+bc[3])/2) < 0.1):
                    continue                          # 同位叠加=同一视觉元素
                if bb_overlap(ba, bc) / min(aa, ac) > 0.5:
                    issues.append(f'第{pg}页{a.name}与{c.name}明显重叠影响阅读')

    if not issues:
        return True, '两页排版：文本在页面内、主要图形无越界、箭头不穿正文、结构元素无明显重叠'
    return False, '排版问题: ' + '; '.join(issues[:6])

# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# 规则表
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
RULES = [
    ('+5',  chk_fonts,                '+5 所有文本字体为 Arial 或 Calibri'),
    ('+1',  chk_p1_y_arrow,           '+1 第1页左侧纵轴箭头'),
    ('+1',  chk_p1_y_label,           '+1 第1页纵轴标题"System Conditions"'),
    ('+1',  chk_p1_x_arrow,           '+1 第1页底部横轴箭头'),
    ('+1',  chk_p1_x_label,           '+1 第1页横轴标题"Time (T)"'),
    ('+1',  chk_p1_top_rrect,         '+1 第1页顶部浅蓝圆角矩形'),
    ('+1',  chk_p1_top_title,         '+1 第1页"Macro Environment"标题'),
    ('+3',  chk_p1_six_texts,         '+3 第1页六组环境文本'),
    ('+3',  chk_p1_curves,            '+3 第1页顶部三条蓝色曲线'),
    ('+1',  chk_p1_big_ellipse,       '+1 第1页中部大椭圆'),
    ('+1',  chk_p1_big_ellipse_title, '+1 第1页大椭圆标题"Local Health Innovation Context"'),
    ('+3',  chk_p1_small_ellipses_top,'+3 第1页大椭圆上排三小椭圆'),
    ('+3',  chk_p1_small_ellipses_bot,'+3 第1页大椭圆下排三小椭圆'),
    ('+3',  chk_p1_blue_arrows_top,   '+3 第1页顶部矩形到椭圆蓝色下箭头'),
    ('+3',  chk_p1_green_arrows_mid,  '+3 第1页椭圆到arena框绿色下箭头'),
    ('+1',  chk_p1_arena_border,      '+1 第1页下方蓝色边框圆角矩形'),
    ('+1',  chk_p1_arena_title_bar,   '+1 第1页"Innovation Dynamics Arena"标题条'),
    ('+3',  chk_p1_demand_module,     '+3 第1页需求模块两个虚线矩形'),
    ('+3',  chk_p1_demand_icons,      '+3 第1页需求模块图标组'),
    ('+1',  chk_p1_arena_vline,       '+1 第1页arena中部竖向虚线'),
    ('+1',  chk_p1_arena_h_arrow,     '+1 第1页arena中部向右箭头'),
    ('+5',  chk_p1_four_ellipses,     '+5 第1页四主体椭圆 Hospitals/Start-ups/Government/Citizens'),
    ('+3',  chk_p1_four_icons,        '+3 第1页四主体图标'),
    ('+3',  chk_p1_four_bidir_arrows, '+3 第1页四主体双向箭头'),
    ('+1',  chk_p1_center_text,       '+1 第1页中心说明文本'),
    ('+3',  chk_p1_dashed_arrows_to_ellipses, '+3 中心文本到四椭圆虚线箭头'),
    ('+1',  chk_p1_left_digital,      '+1 第1页左侧"Digital health initiatives (T1)"'),
    ('+1',  chk_p1_right_digital,     '+1 第1页右侧"Digital health outcomes (T2)"'),
    ('+1',  chk_p2_y_arrow,           '+1 第2页左侧纵轴蓝色箭头'),
    ('+1',  chk_p2_y_labels,          '+1 第2页纵轴文字组'),
    ('+1',  chk_p2_x_arrow,           '+1 第2页底部横轴箭头'),
    ('+1',  chk_p2_x_label,           '+1 第2页横轴标题"Time"'),
    ('+1',  chk_p2_maturity_curve,    '+1 第2页成熟度曲线箭头'),
    ('+1',  chk_p2_curve_nodes,       '+1 第2页三个圆形节点'),
    ('+1',  chk_p2_stage1_card,       '+1 第2页第一阶段外层卡片'),
    ('+1',  chk_p2_stage1_num,        '+1 第2页第一阶段编号"01"'),
    ('+3',  chk_p2_stage1_title,      '+3 第2页第一阶段标题条'),
    ('+3',  chk_p2_stage1_milestones, '+3 第2页第一阶段里程碑'),
    ('+3',  chk_p2_stage1_profile,    '+3 第2页第一阶段Stage profile'),
    ('+1',  chk_p2_stage2_card,       '+1 第2页第二阶段外层卡片'),
    ('+1',  chk_p2_stage2_num,        '+1 第2页第二阶段编号"02"'),
    ('+3',  chk_p2_stage2_title,      '+3 第2页第二阶段标题条'),
    ('+3',  chk_p2_stage2_milestones, '+3 第2页第二阶段里程碑'),
    ('+3',  chk_p2_stage2_profile,    '+3 第2页第二阶段Stage profile'),
    ('+1',  chk_p2_stage3_card,       '+1 第2页第三阶段外层卡片'),
    ('+1',  chk_p2_stage3_num,        '+1 第2页第三阶段编号圆形'),
    ('+3',  chk_p2_stage3_title,      '+3 第2页第三阶段标题条'),
    ('+3',  chk_p2_stage3_milestones, '+3 第2页第三阶段里程碑'),
    ('+3',  chk_p2_stage3_profile,    '+3 第2页第三阶段Stage profile'),
    ('+1',  chk_p2_left_vline,        '+1 第2页左中竖向虚线分隔线'),
    ('+1',  chk_p2_right_vline,       '+1 第2页右中竖向虚线分隔线'),
    ('+1',  chk_p2_stage_buttons,     '+1 第2页阶段间圆形箭头按钮'),
]

# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# 主评估流程
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
SCRIPT_ID = '050'
MAX_SCORE = sum(int(r[0]) for r in RULES)


def _locate_pptx(dir_path: Path) -> Path | None:
    """在给定目录内定位待评估的 .pptx 文档。
    若目录内有多个，取修改时间最新的一个；找不到返回 None。"""
    if not dir_path.is_dir():
        return None
    candidates = [p for p in dir_path.iterdir()
                  if p.is_file() and p.suffix.lower() == '.pptx'
                  and not p.name.startswith('~$')]
    if not candidates:
        return None
    return max(candidates, key=lambda p: p.stat().st_mtime)


def _empty_dim2_items() -> list[dict]:
    """按 RULES 顺序生成未命中的空评分项，用于错误或维度一未通过时对齐列。"""
    return [{'rule': desc, 'max_delta': int(score_str),
             'delta': 0, 'hit': False, 'detail': ''} for score_str, _, desc in RULES]


def evaluate(dir_path: str) -> dict:
    """评估入口：接收脚本所在目录路径，脚本自行在目录中定位 .pptx 文档并评估。

    返回结构化 dict，字段见项目《脚本接口差异与统一建议》§2.2。
    脚本内不改 sys.stdout、不 sys.exit、不硬编码路径。
    """
    try:
        d = Path(dir_path)
        pptx_path = _locate_pptx(d)
        if pptx_path is None:
            return {
                'id': SCRIPT_ID, 'file_name': '', 'status': 'error',
                'error': f'目录内未找到 .pptx 文件: {dir_path}',
                'dim1_pass': False, 'dim1_reason': '',
                'dim2_items': _empty_dim2_items(),
                'total_score': 0, 'max_score': MAX_SCORE,
            }

        # 格式检查（维度一的先决条件）—— 仅按后缀判断，不再单独判"是否能打开"
        try:
            page_w, page_h, slides, theme_fonts = load_pptx(pptx_path)
        except Exception as e:
            return {
                'id': SCRIPT_ID, 'file_name': pptx_path.name, 'status': 'ok',
                'error': None,
                'dim1_pass': False, 'dim1_reason': f'解析PPTX失败: {e}',
                'dim2_items': [], 'total_score': 0, 'max_score': MAX_SCORE,
            }

        # 记录主题字体，供字体检测在文本未显式指定字体（继承主题）或引用
        # +mn-lt/+mj-lt 时，解析为办公软件实际渲染的字体。
        THEME_FONTS.update(theme_fonts)

        # ─── 维度1 ───
        d1_failures = check_dim1(str(pptx_path), slides)
        if d1_failures:
            return {
                'id': SCRIPT_ID, 'file_name': pptx_path.name, 'status': 'ok',
                'error': None,
                'dim1_pass': False, 'dim1_reason': '; '.join(d1_failures),
                'dim2_items': [], 'total_score': 0, 'max_score': MAX_SCORE,
            }

        # ─── 维度2 ───
        total = 0
        dim2_items: list[dict] = []
        for score_str, fn, desc in RULES:
            score = int(score_str)
            try:
                hit, _ = fn(slides=slides, page_w=page_w, page_h=page_h)
            except Exception:
                hit = False
            dim2_items.append({
                'rule': desc, 'max_delta': score,
                'delta': score if hit else 0,
                'hit': bool(hit),
                # 统一置空：不影响 hit/delta 判定与总分计算
                'detail': '',
            })
            if hit:
                total += score

        return {
            'id': SCRIPT_ID, 'file_name': pptx_path.name, 'status': 'ok',
            'error': None,
            'dim1_pass': True, 'dim1_reason': '',
            'dim2_items': dim2_items,
            'total_score': total, 'max_score': MAX_SCORE,
        }
    except Exception as e:
        return {
            'id': SCRIPT_ID, 'file_name': '', 'status': 'error',
            'error': f'{type(e).__name__}: {e}',
            'dim1_pass': False, 'dim1_reason': '',
            'dim2_items': _empty_dim2_items(),
            'total_score': 0, 'max_score': MAX_SCORE,
        }


if __name__ == '__main__':
    # 本地调试用：默认使用脚本所在目录；也可显式传入目录路径
    target_dir = sys.argv[1] if len(sys.argv) > 1 else str(Path(__file__).resolve().parent)
    result = evaluate(target_dir)
    # 主结果通过 evaluate 的 return 值返回；此处仅为脚本作者自测的可视化输出。
    # 直接写 UTF-8 字节到 stdout.buffer，避免 Windows 控制台默认 GBK 编码打印中文
    # 报 UnicodeEncodeError，同时不修改 sys.stdout 本身（遵守统一约定 §2.3）。
    payload = json.dumps(result, ensure_ascii=False, indent=2) + '\n'
    try:
        _ = sys.stdout.buffer.write(payload.encode('utf-8'))
    except AttributeError:
        # 无 buffer 属性(极少见)时退回带转义的 ASCII 输出
        print(json.dumps(result, ensure_ascii=True, indent=2))
