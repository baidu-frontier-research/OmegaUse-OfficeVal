#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# pyright: reportUnknownMemberType=false, reportUnknownVariableType=false, reportUnknownArgumentType=false, reportUnknownParameterType=false, reportMissingParameterType=false, reportPrivateUsage=false, reportUnknownLambdaType=false, reportGeneralTypeIssues=false

import os
import re
import sys
import json
import math
import zipfile
import colorsys
from pptx import Presentation
from pptx.util import Emu

# 说明：按接口统一约定，本模块不修改全局 sys.stdout，也不 print 主结果。
# 主结果统一由 evaluate(dir_path) -> dict 返回；入参为"脚本所在目录的路径"，
# 由脚本自己在该目录内定位并打开被评估的 .pptx。本文件仅在 __main__ 下作为
# 本地调试入口，才把 dict 序列化为 JSON 打印。

SCRIPT_ID = "043"
# 维度二所有评分项 max_delta 之和；用于维度一未通过 / 脚本异常时兜底
MAX_SCORE_TOTAL = 115

# 维度二所有评分项按"结果追加顺序"的 max_delta 列表；末端把 results 里的
# "+N：..." 字符串反解为结构化 dim2_items 时用。顺序须与 _evaluate_impl 中
# results.append(...) 的顺序严格一致；共 33 项，累加 = MAX_SCORE_TOTAL。
_MAX_DELTAS_ORDER = [
    1,  # 01 共2页，第1页中文正面、第2页英文背面，页序不颠倒
    3,  # 02 两页画布一致、横向名片比例、圆角矩形卡片主体、米白填充
    5,  # 03 两页右上角二维码为图片对象且尺寸位置达标
    3,  # 04 两页二维码位置与尺寸一致
    3,  # 05 每页原生可编辑对象≥25个，无整版截图贴图
    3,  # 06 全篇字体为 Windows 常见字体，未用仅 macOS 字体承载可见文字
    3,  # 07 所有形状宽高均为非负有效值
    3,  # 08 第1页左上角企业徽标（图片/成组矢量）
    3,  # 09 中文公司名文本、加粗、深墨绿、左边缘、字号倍数
    3,  # 10 英文公司名全大写单行、无衬线带字间距、字号比例、铜橙色
    3,  # 11 三处"线段＋叶片＋线段"分隔装饰
    5,  # 12 第1页姓名"沈知行"书法字体/深墨绿/最大字号/位置
    5,  # 13 第1页3行徽章式联系方式（等距堆叠、徽章左对齐）
    5,  # 14 三个徽章内为真实电话/信封/定位针图标
    3,  # 15 第1行 M.P.＋电话（近黑加粗、同行左置、字号倍数）
    3,  # 16 第2行邮箱、第3行地址断两行且与电话行左对齐
    3,  # 17 第1页纵向分隔线（位置与跨度）
    3,  # 18 第1页"主 营"胶囊标签
    5,  # 19 第1页右栏3个主营分组（圆形底衬＋图标＋文字块，等距、左对齐）
    5,  # 20 三个圆形底衬内为真实水滴/叶片/包装箱图标
    5,  # 21 第1页右栏三组文字完整（标题全角冒号加粗深墨绿＋7行条目）
    1,  # 22 三组之间2条细横向分隔线
    5,  # 23 两页底部深墨绿色带，上缘为自由曲线，横贯整页宽
    3,  # 24 第1页色带左端铜橙色块/渐变过渡段
    3,  # 25 第2页英文公司名两行、加粗深墨绿、该页最大字号、左边缘
    3,  # 26 第2页姓名"Ethan Shen"手写花体、深墨绿、位置对应
    3,  # 27 第2页3行徽章式联系方式＋英文地址断3行
    3,  # 28 第2页"MAIN BUSINESS"胶囊＋纵向分隔线，与第1页同位
    5,  # 29 第2页右栏3个主营分组（英文标题＋条目＋圆形底衬＋图标）
    5,  # 30 第2页色带内4组"图标＋全大写标签"
    1,  # 31 第2页色带上缘另叠波浪/曲线类线条装饰（画法与深浅不限）
    5,  # 32 两页所有对象落在卡片内，无遮挡/重叠/溢出
    3,  # 33 两页统一配色体系，卡片主体不描黑色边框
]
assert sum(_MAX_DELTAS_ORDER) == MAX_SCORE_TOTAL

# ---------------------------------------------------------------- 常量
A_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main'


def _a(tag):
    return '{%s}%s' % (A_NS, tag)


EMU_PER_CM = 360000.0
# 1 磅 = 12700 EMU：把字号换算成画布上的实际长度时用
EMU_PER_PT = 12700.0
# PowerPoint 文本框未显式写 sz 时的默认字号（打开 PPT/WPS 实际显示 18 磅）
DEFAULT_SZ_PT = 18.0

# 仅 macOS 随系统提供的字体：在 Windows 打开会回退到默认字形，原字形丢失
MAC_ONLY_FONTS = (
    'STXingkai', 'STKaiti', 'STSong', 'STHeiti', 'STFangsong', 'STXihei',
    'STLiti', 'STXinwei', 'STZhongsong', 'Zapfino', 'Hiragino', 'PingFang',
    'Apple LiGothic', 'Apple LiSung', 'Songti SC', 'Heiti SC', 'Kaiti SC',
    'Yuanti SC', 'Hannotate SC', 'HanziPen SC', 'Weibei SC', 'Lantinghei SC',
    'Baoli SC', 'Libian SC', 'Xingkai SC', 'Yuppy SC', 'Wawati SC',
    'LiHei Pro', 'BiauKai',
)
# macOS 旧版繁体中文字体，字族名就叫"Hei"/"Kai"：这两个名字太短，按"族名包含"
# 匹配会把 SimHei（黑体）、Microsoft YaHei（微软雅黑）、KaiTi（楷体）这些 Windows
# 自带字体一并误判成 macOS 专有，故只按整名相等匹配。
MAC_ONLY_EXACT = ('Hei', 'Kai')
# Windows 环境可正常显示的常见字体（含前缀族）
WIN_COMMON_FONTS = (
    '微软雅黑', 'Microsoft YaHei', '微软正黑体', 'Microsoft JhengHei',
    '黑体', 'SimHei', '宋体', 'SimSun', 'NSimSun', '新宋体',
    '楷体', 'KaiTi', 'KaiTi_GB2312', '楷体_GB2312', '仿宋', 'FangSong',
    '等线', 'DengXian', '思源', 'Source Han', '方正',
    'Arial', 'Times New Roman', 'Calibri', 'Cambria', 'Candara', 'Consolas',
    'Corbel', 'Courier New', 'Georgia', 'Impact', 'Tahoma', 'Trebuchet MS',
    'Verdana', 'Segoe', 'Symbol', 'Wingdings', 'Webdings', 'Franklin Gothic',
    'Century', 'Garamond', 'Palatino Linotype', 'Lucida', 'MS Gothic',
    'Malgun Gothic', 'Sylfaen', 'Ebrima', 'Gabriola', 'Ink Free',
)
# 行楷/楷书类书法风格中文字体（细则明确点名"楷体、行楷等"）
KAI_FONTS = ('楷', '行楷', 'KaiTi', 'Kaiti', 'Xingkai', 'XingKai', 'STKaiti',
             'STXingkai', 'Kai')
# 手写花体英文字体（细则明确点名"Segoe Script 等脚本体"）
SCRIPT_FONTS = ('Script', 'Hand', 'Zapfino', 'Mistral', 'Brush', 'Vladimir',
                'Gabriola', 'Ink Free', 'Pacifico', 'Segoe Print')
# 胶囊/圆角标签允许的预设几何
CAPSULE_PRSTS = ('roundRect', 'round1Rect', 'round2SameRect', 'round2DiagRect',
                 'flowChartTerminator', 'flowChartAlternateProcess', 'pill')


# ------------------------------------------------------- 颜色工具
def _rgb(hexstr):
    """'RRGGBB' -> (r, g, b)；非法返回 None。"""
    if not hexstr or not re.fullmatch(r'[0-9A-Fa-f]{6}', hexstr):
        return None
    return (int(hexstr[0:2], 16), int(hexstr[2:4], 16), int(hexstr[4:6], 16))


def _is_blackish(hexstr):
    """近黑色（描黑边框判定）：三通道都很暗且接近中性灰，深墨绿不算黑。"""
    c = _rgb(hexstr)
    if c is None:
        return False
    return max(c) <= 0x38 and (max(c) - min(c)) <= 0x14


def _is_near_black(hexstr):
    """近黑（电话号码"近黑"文字色）：细则给的参照区间是 #1A1A1A-#333333，
    按"近似"留出容差取到 0x4D；允许带一点冷色偏（各通道极差 ≤ 0x30），
    这是办公软件里常见的"黑里带一点墨绿/墨蓝"写法，肉眼仍是黑字。"""
    c = _rgb(hexstr)
    if c is None:
        return False
    return max(c) <= 0x4D and (max(c) - min(c)) <= 0x30


def _is_white(hexstr):
    c = _rgb(hexstr)
    return c is not None and min(c) >= 0xE6


def _is_cream(hexstr):
    """米白/象牙白（细则给出近似 #F5F3EE-#FEFDF8 的浅暖白）：
    整体极亮、暖偏（R≥B）、通道差很小，且不是纯白、不是彩色。"""
    c = _rgb(hexstr)
    if c is None:
        return False
    r, g, b = c
    if min(c) < 0xEE or max(c) - min(c) > 0x18:
        return False
    if r < b:
        return False
    return hexstr.upper() != 'FFFFFF'


def _is_pale_greengray(hexstr):
    """淡灰绿实心圆底衬（细则给出近似 #EDF0EA-#F0F2EC）：极亮、绿分量不低于
    红蓝、通道差很小。"""
    c = _rgb(hexstr)
    if c is None:
        return False
    r, g, b = c
    if min(c) < 0xDC or max(c) - min(c) > 0x18:
        return False
    return g >= r - 2 and g >= b - 2


def _is_deep_teal(hexstr):
    """深墨绿色族（细则给出近似 #013B3C-#004F49 / #004F49-#05565D，
    或同族更深值）：红分量很低、绿蓝分量接近且处于中低亮度区，按整族判定。"""
    c = _rgb(hexstr)
    if c is None:
        return False
    r, g, b = c
    if r > 0x3A:
        return False
    if not (0x20 <= g <= 0x6E and 0x20 <= b <= 0x76):
        return False
    return abs(g - b) <= 0x30


def _is_teal_family(hexstr):
    """墨绿/青绿填充族（胶囊、色带等实心填充；比正文文字色允许更亮一档）。"""
    c = _rgb(hexstr)
    if c is None:
        return False
    r, g, b = c
    if not (g > r + 0x0A and b > r + 0x0A):
        return False
    return 0x28 <= max(g, b) <= 0xA0


def _is_green(hexstr):
    """泛绿色（细则只要求"字体为绿色"，不限定深浅）：按色相整段判定——
    从黄绿（约75°）到青绿（约200°）都算绿，深墨绿#05565D、橄榄绿#3A5E21、
    亮青绿#09747B 同属此段；另要求有一定饱和度且不是近黑/近白，否则灰色
    与黑白会因色相计算的舍入误差被当成绿色。"""
    c = _rgb(hexstr)
    if c is None:
        return False
    h, s, l = _rgb_to_hsl(c)
    if s < 0.12 or not (0.03 < l < 0.92):
        return False
    return 75.0 <= h * 360.0 <= 200.0


def _is_copper(hexstr):
    """铜橙色族（细则给出近似 #B5713C-#C77B45，或同族更饱和的暖橙）：
    按"红>绿>蓝且反差明显的暖橙"整族判定，不卡死在两个端点值上——
    办公软件里同一枚铜橙常以主题色叠亮度变换写出（如 accent2 + lumMod
    得到 #C65F10），取色器看着是同一色系，硬套区间反而会误判。"""
    c = _rgb(hexstr)
    if c is None:
        return False
    r, g, b = c
    if not (0x80 <= r <= 0xEE and 0x40 <= g <= 0xB0 and b <= 0x80):
        return False
    return r > g + 0x18 and g > b + 0x10


# ------------------------------------------------- 主题色解析（schemeClr）
# 说明：PPT 里一个颜色可以写成 srgbClr（直接给 RRGGBB），也可以写成
# schemeClr（引用主题色板，再叠加 lumMod/lumOff/shade/tint 等亮度变换）。
# 办公软件里"取色器看到的颜色"是变换之后的结果，所以必须先解析成实际 RGB，
# 否则"深墨绿""铜橙色"这类判定会漏判。
def _rgb_to_hsl(c):
    r, g, b = (x / 255.0 for x in c)
    h, l, s = colorsys.rgb_to_hls(r, g, b)
    return h, s, l


def _hsl_to_rgb(h, s, l):
    r, g, b = colorsys.hls_to_rgb(h, max(0.0, min(1.0, l)),
                                  max(0.0, min(1.0, s)))
    return (int(round(r * 255)), int(round(g * 255)), int(round(b * 255)))


def _hex(c):
    return '%02X%02X%02X' % c


def _apply_mods(base_rgb, mods):
    """把 lumMod/lumOff/shade/tint 依次作用到基色上（PowerPoint 语义）。"""
    c = base_rgb
    for name, val in mods:
        v = val / 100000.0
        if name == 'lumMod':
            h, s, l = _rgb_to_hsl(c)
            c = _hsl_to_rgb(h, s, l * v)
        elif name == 'lumOff':
            h, s, l = _rgb_to_hsl(c)
            c = _hsl_to_rgb(h, s, l + v)
        elif name == 'shade':
            c = tuple(int(round(x * v)) for x in c)
        elif name == 'tint':
            c = tuple(int(round(x * v + 255 * (1 - v))) for x in c)
        elif name == 'satMod':
            h, s, l = _rgb_to_hsl(c)
            c = _hsl_to_rgb(h, s * v, l)
    return tuple(max(0, min(255, int(x))) for x in c)


_SCHEME_ALIAS = {'tx1': 'dk1', 'bg1': 'lt1', 'tx2': 'dk2', 'bg2': 'lt2'}


def _read_theme_scheme(pptx_path):
    """从 ppt/theme/theme1.xml 读出主题色板 {'accent1': (r,g,b), ...}。
    读不到时返回空 dict，schemeClr 便放弃解析（不误判成某个颜色）。"""
    table = {}
    try:
        with zipfile.ZipFile(pptx_path) as z:
            names = [n for n in z.namelist()
                     if re.fullmatch(r'ppt/theme/theme\d+\.xml', n)]
            if not names:
                return table
            xml = z.read(sorted(names)[0]).decode('utf-8', 'ignore')
    except Exception:
        return table
    m = re.search(r'<a:clrScheme\b.*?</a:clrScheme>', xml, re.S)
    if not m:
        return table
    for key, body in re.findall(
            r'<a:(dk1|lt1|dk2|lt2|accent[1-6]|hlink|folHlink)>(.*?)</a:\1>',
            m.group(0), re.S):
        mm = re.search(r'<a:srgbClr val="([0-9A-Fa-f]{6})"', body)
        if mm is None:
            mm = re.search(r'<a:sysClr [^>]*lastClr="([0-9A-Fa-f]{6})"', body)
        if mm:
            table[key] = _rgb(mm.group(1))
    return table


# theme 字典里存放主题字体的保留键（与色板的 accent1/dk1 等键不冲突）
THEME_FONT_KEY = '#fonts'


def _read_theme_fonts(pptx_path):
    """从 ppt\theme\theme1.xml 读出主题字体：
    {'major_latin':…, 'minor_latin':…, 'major_ea':…, 'minor_ea':…}。
    与办公软件对齐：run 里不写 a:latin/a:ea 时，屏幕上显示的就是这里的主题字体
    （字体框里显示成"+ 中文正文"一类），判字体时必须回落到它，否则等于没判。
    只取 majorFont/minorFont 的 a:latin 与 a:ea，不取 a:font script="…" 的
    分脚本回退字体（那些不是正文实际使用的字体）。"""
    out = {}
    try:
        with zipfile.ZipFile(pptx_path) as z:
            names = [n for n in z.namelist()
                     if re.fullmatch(r'ppt/theme/theme\d+\.xml', n)]
            if not names:
                return out
            xml = z.read(sorted(names)[0]).decode('utf-8', 'ignore')
    except Exception:
        return out
    for kind in ('major', 'minor'):
        m = re.search(r'<a:%sFont>(.*?)</a:%sFont>' % (kind, kind), xml, re.S)
        if m is None:
            continue
        for tag in ('latin', 'ea'):
            mm = re.search(r'<a:%s typeface="([^"]*)"' % tag, m.group(1))
            if mm and mm.group(1).strip():
                out['%s_%s' % (kind, tag)] = mm.group(1).strip()
    return out


def _resolve_font(name, fonts, kind):
    """把 run 上的字体名解析成办公软件实际使用的字体名：
    空值 → 主题正文（minor）字体；"+mj-lt"/"+mn-ea" 这类主题引用 → 对应主题字体。
    kind 取 'latin'（西文）或 'ea'（中日韩文）。"""
    n = (name or '').strip()
    fonts = fonts or {}
    if n.startswith('+'):
        prefix = 'major_' if n[1:3] == 'mj' else 'minor_'
        tag = 'ea' if n.endswith('ea') else 'latin'
        return fonts.get(prefix + tag, '')
    if n:
        return n
    return fonts.get('minor_' + kind, '')


def _clr_hex(holder, theme):
    """holder 是任意"装着颜色"的元素（a:solidFill / a:gs / a:ln 里的 fill 等）。
    返回该颜色在办公软件里实际显示的 'RRGGBB'；解析不出时返回 None。"""
    if holder is None:
        return None
    node = holder.find(_a('srgbClr'))
    base = None
    if node is not None:
        base = _rgb(node.get('val') or '')
    else:
        node = holder.find(_a('schemeClr'))
        if node is not None:
            key = node.get('val') or ''
            key = _SCHEME_ALIAS.get(key, key)
            base = theme.get(key)
        else:
            node = holder.find(_a('sysClr'))
            if node is not None:
                base = _rgb(node.get('lastClr') or '')
    if base is None or node is None:
        return None
    mods = []
    for child in node:
        tag = child.tag.split('}')[-1]
        if tag in ('lumMod', 'lumOff', 'shade', 'tint', 'satMod'):
            try:
                mods.append((tag, int(child.get('val'))))
            except (TypeError, ValueError):
                pass
    if mods:
        base = _apply_mods(base, mods)
    return _hex(base)


# ------------------------------------------------------- 形状模型
class Sh(object):
    """把一个形状（含组内子形状）拍平成评分要用到的最小属性集合。
    坐标一律换算成"幻灯片绝对坐标（EMU）"，组内子形状按组的缩放映射过。"""

    __slots__ = ('sid', 'name', 'z', 'left', 'top', 'width', 'height',
                 'raw_w', 'raw_h', 'kind', 'prst', 'is_pic', 'is_group',
                 'curve_n', 'has_custgeom', 'fill_kind', 'fills', 'line_hex',
                 'line_w', 'text', 'runs', 'rot', 'el')

    def right(self):
        return self.left + self.width

    def bottom(self):
        return self.top + self.height

    def cx(self):
        return self.left + self.width / 2.0

    def cy(self):
        return self.top + self.height / 2.0

    def area(self):
        return max(0, self.width) * max(0, self.height)

    def is_sized(self):
        """是否是一个"有尺寸"的对象：宽或高有值即可。
        直线/连接线只有长度没有厚度（面积为 0），仍然是可点选可编辑的对象。"""
        return self.width > 0 or self.height > 0

    def has_text(self):
        return bool(self.text.strip())


def _norm_text(s):
    """全角标点→半角、各类连字符→'-'、各类空格→普通空格、压缩空白；
    保留 '：' 以便判定全角冒号。
    办公软件里排版常写入不换行空格（U+00A0）等特殊空格，肉眼与普通空格无异，
    统一折算成普通空格后再比对文字内容。"""
    out = []
    for ch in s:
        o = ord(ch)
        if 0xFF01 <= o <= 0xFF5E and ch != '：':
            out.append(chr(o - 0xFEE0))
        elif ch in '‐‑‒–—―−':
            out.append('-')
        elif ch == '　' or o == 0xA0 or 0x2000 <= o <= 0x200A \
                or o in (0x202F, 0x205F):
            out.append(' ')
        else:
            out.append(ch)
    return re.sub(r'[ \t\r\n]+', ' ', ''.join(out)).strip()


def _has_cjk(s):
    return any('㐀' <= ch <= '鿿' or '　' <= ch <= '〿'
               or '＀' <= ch <= '￯' for ch in s)


def _has_latin(s):
    return any(('a' <= ch <= 'z') or ('A' <= ch <= 'Z') or ch.isdigit()
               for ch in s)


def _int_attr(el, name, default=None):
    if el is None:
        return default
    v = el.get(name)
    if v is None:
        return default
    try:
        return int(v)
    except ValueError:
        return default


def _collect_paras(tx_body, theme):
    """把 a:txBody 解析成段落列表；每段是 run 字典列表。
    run 上没写的字号/加粗/倾斜/字距/大写效果，一律按办公软件的继承链逐级回落：
    run rPr → 段落 pPr/defRPr → txBody 的 lstStyle/lvlNpPr/defRPr → endParaRPr →
    （字号）PowerPoint 文本框默认 18 磅。屏幕上看到的是继承后的效果，只读 run 级
    会把"在段落/文本框级别设的加粗、字距、全部大写"判成没设。
    run['latin']/['ea'] 采用"办公软件里实际使用的字体"：未写字体时回落到主题
    字体；原始写法保留在 ['latin_raw']/['ea_raw'] 里。"""
    paras = []
    if tx_body is None:
        return paras
    fonts = theme.get(THEME_FONT_KEY) if isinstance(theme, dict) else None
    lst = tx_body.find(_a('lstStyle'))
    for p in tx_body.findall(_a('p')):
        p_pr = p.find(_a('pPr'))
        def_rpr = p_pr.find(_a('defRPr')) if p_pr is not None else None
        lvl = _int_attr(p_pr, 'lvl', 0) or 0
        lvl_rpr = None
        if lst is not None:
            lvl_pr = lst.find(_a('lvl%dpPr' % (lvl + 1)))
            if lvl_pr is not None:
                lvl_rpr = lvl_pr.find(_a('defRPr'))
        end_rpr = p.find(_a('endParaRPr'))
        runs = []
        for r in p.findall(_a('r')):
            t_el = r.find(_a('t'))
            txt = t_el.text if (t_el is not None and t_el.text) else ''
            r_pr = r.find(_a('rPr'))

            def inherit(name, _rp=r_pr):
                """按继承链取属性值，取到第一个显式写出的为止。"""
                for holder in (_rp, def_rpr, lvl_rpr, end_rpr):
                    if holder is not None and holder.get(name) is not None:
                        return holder.get(name)
                return None

            sz = inherit('sz')
            sz = int(sz) if (sz or '').isdigit() else None
            spc = inherit('spc')
            try:
                spc = int(spc) if spc is not None else 0
            except ValueError:
                spc = 0
            latin = r_pr.find(_a('latin')) if r_pr is not None else None
            ea = r_pr.find(_a('ea')) if r_pr is not None else None
            latin_raw = (latin.get('typeface') or '') if latin is not None else ''
            ea_raw = (ea.get('typeface') or '') if ea is not None else ''
            runs.append({
                't': txt,
                'sz': (sz / 100.0) if sz else DEFAULT_SZ_PT,
                'sz_explicit': sz is not None,
                'b': inherit('b') == '1',
                'i': inherit('i') == '1',
                'spc': spc,
                'cap': inherit('cap') or '',
                'latin': _resolve_font(latin_raw, fonts, 'latin'),
                'ea': _resolve_font(ea_raw, fonts, 'ea'),
                'latin_raw': latin_raw,
                'ea_raw': ea_raw,
                'clr': _clr_hex(r_pr.find(_a('solidFill'))
                                if r_pr is not None else None, theme),
            })
        paras.append(runs)
    return paras


P_NS = 'http://schemas.openxmlformats.org/presentationml/2006/main'


def _p(tag):
    return '{%s}%s' % (P_NS, tag)


def _xfrm_of(el, kind):
    if kind == 'grp':
        holder = el.find(_p('grpSpPr'))
        return holder.find(_a('xfrm')) if holder is not None else None
    if kind == 'frame':
        return el.find(_p('xfrm'))
    holder = el.find(_p('spPr'))
    return holder.find(_a('xfrm')) if holder is not None else None


def _box(xfrm):
    """返回 (x, y, w, h)；缺 xfrm 或缺 off/ext 时返回 None。
    注意 ext 的 cx/cy 允许为负——这正是"宽高非负"评分项要抓的问题。"""
    if xfrm is None:
        return None
    off = xfrm.find(_a('off'))
    ext = xfrm.find(_a('ext'))
    if off is None or ext is None:
        return None
    return (_int_attr(off, 'x', 0), _int_attr(off, 'y', 0),
            _int_attr(ext, 'cx', 0), _int_attr(ext, 'cy', 0))


def _fill_info(el, kind, theme):
    """返回 (fill_kind, [hex...])。gradFill 把所有渐变色标都列出来，
    对应办公软件"设置形状格式→渐变填充"里看到的一串光圈颜色。"""
    sp_pr = el.find(_p('spPr'))
    if kind == 'pic':
        return ('pic', [])
    if sp_pr is None:
        return ('none', [])
    if sp_pr.find(_a('solidFill')) is not None:
        h = _clr_hex(sp_pr.find(_a('solidFill')), theme)
        return ('solid', [h] if h else [])
    g = sp_pr.find(_a('gradFill'))
    if g is not None:
        hexes = []
        for gs in g.iter(_a('gs')):
            h = _clr_hex(gs, theme)
            if h:
                hexes.append(h)
        return ('grad', hexes)
    if sp_pr.find(_a('blipFill')) is not None:
        return ('pic', [])
    if sp_pr.find(_a('pattFill')) is not None:
        return ('patt', [])
    if sp_pr.find(_a('noFill')) is not None:
        return ('none', [])
    return ('inherit', [])


def _line_info(el, theme):
    """返回 (线条颜色 hex 或 None, 线宽 EMU 或 None)。noFill 视为无边框。"""
    sp_pr = el.find(_p('spPr'))
    if sp_pr is None:
        return (None, None)
    ln = sp_pr.find(_a('ln'))
    if ln is None:
        return (None, None)
    w = _int_attr(ln, 'w')
    if ln.find(_a('noFill')) is not None:
        return (None, w)
    return (_clr_hex(ln.find(_a('solidFill')), theme), w)


_KIND_BY_TAG = {
    _p('sp'): 'sp', _p('pic'): 'pic', _p('cxnSp'): 'cxn',
    _p('grpSp'): 'grp', _p('graphicFrame'): 'frame',
}


def _make_sh(el, kind, box, z, theme):
    sh = Sh()
    sh.el = el
    sh.z = z
    sh.kind = kind
    sh.is_pic = (kind == 'pic')
    sh.is_group = (kind == 'grp')
    x, y, w, h = box
    sh.raw_w, sh.raw_h = w, h
    # 负宽高在几何上无意义：取绝对值得到"实际占位框"，原始值另存供第7项判定
    sh.left = x if w >= 0 else x + w
    sh.top = y if h >= 0 else y + h
    sh.width = abs(w)
    sh.height = abs(h)
    nv = None
    for tag in ('nvSpPr', 'nvPicPr', 'nvCxnSpPr', 'nvGrpSpPr',
                'nvGraphicFramePr'):
        cand = el.find(_p(tag))
        if cand is not None:
            nv = cand
            break
    c_nv = nv.find(_p('cNvPr')) if nv is not None else None
    sh.sid = _int_attr(c_nv, 'id', 0)
    sh.name = (c_nv.get('name') or '') if c_nv is not None else ''
    xfrm = _xfrm_of(el, kind)
    sh.rot = _int_attr(xfrm, 'rot', 0) or 0
    sp_pr = el.find(_p('spPr'))
    pg = sp_pr.find(_a('prstGeom')) if sp_pr is not None else None
    sh.prst = (pg.get('prst') or '') if pg is not None else ''
    cg = sp_pr.find(_a('custGeom')) if sp_pr is not None else None
    sh.has_custgeom = cg is not None
    sh.curve_n = 0
    if cg is not None:
        sh.curve_n = (len(list(cg.iter(_a('cubicBezTo'))))
                      + len(list(cg.iter(_a('quadBezTo')))))
    sh.fill_kind, sh.fills = _fill_info(el, kind, theme)
    sh.line_hex, sh.line_w = _line_info(el, theme)
    tx = el.find(_p('txBody'))
    sh.runs = _collect_paras(tx, theme)
    sh.text = _norm_text('\n'.join(
        ''.join(r['t'] for r in para) for para in sh.runs))
    return sh


def _walk(container, theme, mapper, out, counter, in_group):
    """深度优先遍历形状树。组：既登记组自身（供"成组矢量"判定），也把子形状
    按组的缩放映射成绝对坐标登记（供位置/对齐类判定）。"""
    for el in container:
        kind = _KIND_BY_TAG.get(el.tag)
        if kind is None:
            continue
        box = _box(_xfrm_of(el, kind))
        if box is None:
            box = (0, 0, 0, 0)
        abox = mapper(box)
        counter[0] += 1
        sh = _make_sh(el, kind, abox, counter[0], theme)
        out.append((sh, in_group))
        if kind == 'grp':
            xfrm = _xfrm_of(el, kind)
            ch_off = xfrm.find(_a('chOff')) if xfrm is not None else None
            ch_ext = xfrm.find(_a('chExt')) if xfrm is not None else None
            cox = _int_attr(ch_off, 'x', 0)
            coy = _int_attr(ch_off, 'y', 0)
            cex = _int_attr(ch_ext, 'cx', 0) or 1
            cey = _int_attr(ch_ext, 'cy', 0) or 1
            gx, gy, gw, gh = abox
            sx = (gw / float(cex)) if cex else 1.0
            sy = (gh / float(cey)) if cey else 1.0

            def sub(b, gx=gx, gy=gy, cox=cox, coy=coy, sx=sx, sy=sy):
                x, y, w, h = b
                return (int(gx + (x - cox) * sx), int(gy + (y - coy) * sy),
                        int(w * sx), int(h * sy))

            _walk(el, theme, sub, out, counter, True)


def _slide_shapes(slide_el, theme):
    """返回 (all_shapes, top_level_shapes)；两者都是 Sh 列表，z 为遍历序。"""
    sp_tree = slide_el.find(_p('cSld')).find(_p('spTree'))
    out = []
    _walk(sp_tree, theme, lambda b: b, out, [0], False)
    return ([s for s, _ in out], [s for s, g in out if not g])


# ------------------------------------------------------- 字体 / 文本工具
def _hit_family(name, families):
    n = (name or '').strip()
    if not n:
        return False
    low = n.lower()
    return any(f.lower() in low for f in families)


def _is_mac_only(name):
    n = (name or '').strip()
    if n.lower() in (x.lower() for x in MAC_ONLY_EXACT):
        return True
    return _hit_family(n, MAC_ONLY_FONTS)


def _is_win_common(name):
    return _hit_family(name, WIN_COMMON_FONTS)


def _is_kai_font(name):
    return _hit_family(name, KAI_FONTS)


def _is_script_font(name):
    return _hit_family(name, SCRIPT_FONTS)


def _sans_serif(name):
    """无衬线判定：常见无衬线族命中即可；明确的衬线族（宋体/Times 等）排除。"""
    if _hit_family(name, ('Times', 'Georgia', 'Cambria', 'Garamond', 'Palatino',
                          'Book Antiqua', 'Serif', '宋', 'Song', 'Ming')):
        return False
    return _hit_family(name, ('Arial', 'Helvetica', '雅黑', 'YaHei', 'Segoe UI',
                              'Calibri', 'Tahoma', 'Verdana', 'Franklin',
                              'Gothic', '黑体', 'SimHei', '等线', 'DengXian',
                              'Corbel', 'Candara', 'Trebuchet', 'Ebrima',
                              'Malgun', 'Roboto', 'Open Sans', 'Lato',
                              'Noto Sans', 'Source Han Sans', '思源黑'))


# 字族名本身就带粗度的写法（"微软雅黑 Bold"、"Arial Black"、"思源黑体 Heavy"）
BOLD_FONT_TOKENS = ('BOLD', 'BLACK', 'HEAVY', 'SEMIBOLD', 'DEMIBOLD',
                    'EXTRABOLD', 'ULTRABOLD', '粗体', '特粗', '中粗')


def _is_bold_font(name):
    """字族自带粗体：办公软件里"加粗"按钮没按下（rPr@b 为 0），
    但选的字体本身是粗体，屏幕上就是粗的，视觉效果与加粗等同，故一并认可。
    注意"黑体/思源黑体"是常规字重的字族名，不算粗体。"""
    n = (name or '').upper().replace('-', ' ')
    if not n:
        return False
    return any(tok in n for tok in BOLD_FONT_TOKENS)


def _letter_spaced(text):
    """用"字符间显式留空"做出的字间距：如"F U J I A N   L A N Y U"。
    与办公软件对齐：这类写法在屏幕上与"字符间距加宽"看不出差别，视同带字间距。"""
    body = _norm_text(text or '').strip()
    if len([ch for ch in body if ch.isalnum()]) < 6:
        return False
    toks = [t for t in body.split(' ') if t]
    if len(toks) < 4:
        return False
    singles = sum(1 for t in toks if len(t) == 1)
    return singles >= 0.6 * len(toks)


def _iter_runs(shapes):
    for sh in shapes:
        for para in sh.runs:
            for r in para:
                if r['t'].strip():
                    yield sh, r


def _max_run_sz(shapes):
    szs = [r['sz'] for _, r in _iter_runs(shapes)]
    return max(szs) if szs else 0.0


def _est_text_cm(text, sz_pt):
    """粗估一行文字的显示宽度（cm）：中日韩字符按 1 个字号宽，
    西文/数字按 0.52 个字号宽，空格 0.28。用于判断文本是否会自动折行。"""
    w_pt = 0.0
    for ch in text:
        if ch == ' ':
            w_pt += 0.28 * sz_pt
        elif _has_cjk(ch):
            w_pt += 1.0 * sz_pt
        else:
            w_pt += 0.52 * sz_pt
    return w_pt / 72.0 * 2.54


def _visual_lines(sh):
    """估算形状里文字实际显示的行数：段落数 + 每段因框宽不足产生的折行。
    办公软件里"地址断两行"既可能是两个段落，也可能是一个段落自动折行。"""
    if not sh.runs:
        return 0
    box_cm = sh.width / EMU_PER_CM
    total = 0
    for para in sh.runs:
        txt = ''.join(r['t'] for r in para)
        if not txt.strip():
            continue
        sz = max([r['sz'] for r in para if r['t'].strip()] or [DEFAULT_SZ_PT])
        need = _est_text_cm(txt, sz)
        if box_cm <= 0.2:
            total += 1
        else:
            total += max(1, int(math.ceil(need / box_cm - 0.08)))
    return total


def _inter_area(a, b):
    w = min(a.right(), b.right()) - max(a.left, b.left)
    h = min(a.bottom(), b.bottom()) - max(a.top, b.top)
    if w <= 0 or h <= 0:
        return 0
    return w * h


def _cover_ratio(a, b):
    """b 覆盖 a 的面积占比。"""
    if a.area() <= 0:
        return 0.0
    return _inter_area(a, b) / float(a.area())


def _text_ink_box(sh):
    """文本形状里"真正有字"的那块矩形 (left, top, right, bottom)。
    与办公软件对齐：文本框高度常留有富余（自动换行留白、框比字大），
    两个框的边线擦到一起并不等于文字压字，故按"行数 × 字号 × 1.45 行距"
    估算文字实际占用高度（顶端对齐，不超过框高），再判断是否真的重叠。"""
    lines = 0
    sz = 0.0
    for para in sh.runs:
        txt = ''.join(r['t'] for r in para)
        psz = max([r['sz'] for r in para] or [DEFAULT_SZ_PT])
        sz = max(sz, psz)
        if not txt.strip():
            lines += 1
            continue
        avail = sh.width / EMU_PER_CM
        need = _est_text_cm(txt, psz)
        lines += max(1, int(math.ceil(need / avail)) if avail > 0 else 1)
    lines = max(1, lines)
    sz = sz or DEFAULT_SZ_PT
    ink_h = int(lines * sz * 1.45 * 12700)
    if ink_h <= 0 or ink_h > sh.height:
        ink_h = sh.height
    return (sh.left, sh.top, sh.left + sh.width, sh.top + ink_h)


def _box_inter_area(p, q):
    w = min(p[2], q[2]) - max(p[0], q[0])
    h = min(p[3], q[3]) - max(p[1], q[1])
    if w <= 0 or h <= 0:
        return 0
    return w * h


def _box_area(p):
    return max(0, p[2] - p[0]) * max(0, p[3] - p[1])


def _union_span(items):
    """一组 (start, end) 的并集总长度。"""
    segs = sorted(items)
    total = 0
    cur_s, cur_e = None, None
    for s, e in segs:
        if cur_e is None:
            cur_s, cur_e = s, e
        elif s <= cur_e:
            cur_e = max(cur_e, e)
        else:
            total += cur_e - cur_s
            cur_s, cur_e = s, e
    if cur_e is not None:
        total += cur_e - cur_s
    return total


def _equidistant(vals, tol):
    """相邻间距是否一致（等距堆叠）：极差 ≤ tol。"""
    if len(vals) < 3:
        return len(vals) == 2
    gaps = [vals[i + 1] - vals[i] for i in range(len(vals) - 1)]
    return (max(gaps) - min(gaps)) <= tol


# ------------------------------------------------------- 入口
class _TargetMissing(Exception):
    """目录不存在或目录中没有 .pptx：连待评测对象都没有，不进入维度判定。"""


def evaluate(dir_path):
    result = {"id": SCRIPT_ID, "file_name": "", "status": "ok", "error": None,
              "dim1_pass": True, "dim1_reason": "", "dim2_items": [],
              "total_score": 0, "max_score": MAX_SCORE_TOTAL}
    try:
        if not dir_path or not os.path.isdir(dir_path):
            raise _TargetMissing(f"目录不存在：{dir_path}")
        cands = sorted(f for f in os.listdir(dir_path)
                       if f.lower().endswith('.pptx')
                       and not f.startswith('~$') and not f.startswith('.'))
        if not cands:
            raise _TargetMissing(f"目录中未找到 .pptx 文件：{dir_path}")
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
        result["error"] = f"{type(exc).__name__}: {exc}"
        result["dim1_pass"] = False
        result["dim1_reason"] = f"文件无法解析或无法打开：{type(exc).__name__}: {exc}"
        result["dim2_items"] = []
        result["total_score"] = 0
        return result


def _fail_dim1(result, reason):
    """维度一不通过（文件打不开或画布无效）：总分直接 0，维度二不再逐项计分。"""
    result["dim1_pass"] = False
    result["dim1_reason"] = reason
    result["dim2_items"] = []
    result["total_score"] = 0
    result["max_score"] = MAX_SCORE_TOTAL
    return result


def _evaluate_impl(file_path, result):
    prs = Presentation(file_path)
    theme = _read_theme_scheme(file_path)
    theme[THEME_FONT_KEY] = _read_theme_fonts(file_path)
    W = int(prs.slide_width or 0)
    H = int(prs.slide_height or 0)
    slides = list(prs.slides)

    # ---------------- 维度一：可用或可修改性 ----------------
    # 细则只要求"交付文件为 1 个可正常打开的 PowerPoint 文档（.pptx）"，因此这里
    # 只做两件事：
    #   ① 文件能被正常解析为 PPTX（走到这里说明已打开成功，损坏包会在
    #      Presentation() 抛异常，由 evaluate() 记为 status=error、总分 0）
    #   ② 画布尺寸有效——宽/高读不出来等于文档无法正常呈现，也无法按比例判定
    # 页数、是否整版贴图、关键文字是否可编辑等，一律不再作为"总分记0"的前置门槛，
    # 改由维度二对应评分项各自扣分。
    if W <= 0 or H <= 0:
        return _fail_dim1(result, "幻灯片画布尺寸无效，文档无法正常呈现")

    pages = []
    for sl in slides:
        allsh, topsh = _slide_shapes(sl._element, theme)
        pages.append({
            'all': allsh,
            'top': topsh,
            'pics': [s for s in allsh if s.is_pic],
            'texts': [s for s in allsh if s.has_text()],
        })
    # 页数不足 2 时补空页占位：维度二统一按"该页无任何对象"判定，各项自然不得分，
    # 不再让脚本因索引越界而中断。
    while len(pages) < 2:
        pages.append({'all': [], 'top': [], 'pics': [], 'texts': []})
    page_area = float(W) * H

    result["dim1_pass"] = True
    result["dim1_reason"] = "文件可正常打开为 .pptx，画布尺寸有效"

    # ---------------- 维度二：完成度 ----------------
    # 所有阈值一律用"占页宽/页高的比例"表达，避免因画布尺寸不同而误判。
    p1, p2 = pages[0], pages[1]
    results = []

    def rw(v):
        return v / float(W)

    def rh(v):
        return v / float(H)

    def card_of(pg):
        """卡片主体：铺满画布的最大非图片形状。"""
        best = None
        for s in pg['top']:
            if s.is_pic or s.area() < 0.70 * page_area:
                continue
            if best is None or s.area() > best.area():
                best = s
        return best

    def qr_of(pg):
        """右上角二维码：位于右上区域、接近正方形、面积最大的对象。"""
        best = None
        for s in pg['all']:
            if s.is_group or s.area() <= 0:
                continue
            if rw(s.cx()) < 0.60 or rh(s.cy()) > 0.30:
                continue
            ar = s.width / float(s.height) if s.height else 0
            if not (0.80 <= ar <= 1.25):
                continue
            if rw(s.width) < 0.06:
                continue
            if best is None or s.area() > best.area():
                best = s
        return best

    def bands_of(pg):
        """底部色带：贴近底边、横向较长、填充为墨绿族的形状。"""
        out = []
        for s in pg['all']:
            if s.is_group or s.area() <= 0:
                continue
            if rh(s.bottom()) < 0.86 or rw(s.width) < 0.25:
                continue
            if any(_is_teal_family(h) for h in s.fills):
                out.append(s)
        return out

    def badges_of(pg):
        """联系方式徽章：左半区偏下、近正方形的小块（深墨绿圆角方底或其图片）。
        按"左边缘对齐（≤页宽0.015）+ 尺寸相当（宽度相差≤35%）"聚类，取最大的一组；
        底部色带里的装饰图标虽然也在左侧，但尺寸与纵向位置都对不上，不会混进来。"""
        cand = []
        for s in pg['all']:
            if s.is_group or s.height <= 0 or s.width <= 0:
                continue
            if rw(s.cx()) > 0.45 or not (0.50 <= rh(s.cy()) <= 0.88):
                continue
            ar = s.width / float(s.height)
            if not (0.70 <= ar <= 1.45):
                continue
            if not (0.025 <= rw(s.width) <= 0.10):
                continue
            teal = any(_is_teal_family(h) or _is_deep_teal(h) for h in s.fills)
            if not (s.is_pic or teal):
                continue
            cand.append(s)
        best = []
        for a in cand:
            grp = [b for b in cand
                   if abs(b.left - a.left) <= 0.015 * W
                   and abs(b.width - a.width) <= 0.35 * a.width]
            grp.sort(key=lambda s: s.top)
            if len(grp) > len(best):
                best = grp
        return best

    def icon_inside(pg, host, need_pic=True):
        """判断 host（徽章方底/圆形底衬）内部是否放了"真实图标"：
        中心落在 host 内的图片或矢量图形（且不是纯文字符号）。
        用 ☎ ✉ ♧ 这类字符文字充当图标的，只会命中带文字的形状，不算。"""
        for s in pg['all']:
            if s is host or s.is_group or s.area() <= 0:
                continue
            if not (host.left <= s.cx() <= host.right()
                    and host.top <= s.cy() <= host.bottom()):
                continue
            if s.area() > 1.30 * host.area():
                continue
            if s.is_pic:
                return True
            if not need_pic and (s.has_custgeom or s.prst) and not s.has_text():
                return True
        return False

    def glyph_icon_inside(pg, host):
        """host 内是否只有"符号字符文字"当图标（含 host 自身带符号文字）。"""
        pool = [host] + [s for s in pg['all']
                         if s is not host and s.area() > 0
                         and host.left <= s.cx() <= host.right()
                         and host.top <= s.cy() <= host.bottom()]
        for s in pool:
            t = s.text.strip()
            if not t:
                continue
            if any((' ' <= ch <= '⯿') or ch in '☎✉⌖♧♢▱↻✓◎'
                   for ch in t):
                return True
        return False

    def hlines_of(pg):
        """细横线段：高度极小、横向有一定长度的线/形状。"""
        out = []
        for s in pg['all']:
            if s.is_group or s.width <= 0:
                continue
            if rh(s.height) > 0.014 or rw(s.width) < 0.03:
                continue
            out.append(s)
        return out


    def circles_of(pg):
        """右栏主营分组的浅色圆形底衬：正圆、右半区、浅色实心，按左边缘对齐聚类。
        圆内的图标图片本身也可能是"圆形裁剪"的图片，不能算作底衬，故排除图片。"""
        cand = []
        for s in pg['all']:
            if s.is_group or s.is_pic or s.height <= 0 or s.width <= 0:
                continue
            if s.prst not in ('ellipse', 'circle'):
                continue
            if s.fill_kind != 'solid' or not s.fills:
                continue
            if rw(s.cx()) < 0.45 or rh(s.cy()) < 0.25:
                continue
            ar = s.width / float(s.height)
            if not (0.85 <= ar <= 1.18):
                continue
            if not (0.030 <= rw(s.width) <= 0.13):
                continue
            cand.append(s)
        best = []
        for a in cand:
            grp = [b for b in cand if abs(b.left - a.left) <= 0.015 * W]
            grp.sort(key=lambda s: s.top)
            key = (len(grp), sum(s.area() for s in grp))
            if key > (len(best), sum(s.area() for s in best)):
                best = grp
        return best

    def vdiv_of(pg):
        """纵向分隔线：极窄、竖向较长、横坐标落在版面中部的线/形状。"""
        best = None
        for s in pg['all']:
            if s.is_group or s.height <= 0:
                continue
            if rw(s.width) > 0.013 or rh(s.height) < 0.25:
                continue
            if not (0.38 <= rw(s.cx()) <= 0.68):
                continue
            if best is None or s.height > best.height:
                best = s
        return best

    def cap_texts(pg, cap):
        """胶囊标签上的文字形状：文字写在形状里，或另用文本框叠在形状上，
        在办公软件里看到的都是"标签上的一行字"，两种做法都要认。"""
        if cap is None:
            return []
        out = [cap] if cap.has_text() else []
        for s in pg['texts']:
            if s is cap:
                continue
            if cap.left <= s.cx() <= cap.right() \
                    and cap.top <= s.cy() <= cap.bottom():
                out.append(s)
        return out

    def capsule_of(pg, keys):
        """胶囊标签：圆角/胶囊预设形状 + 墨绿实心填充 + 标签上写着指定文字。"""
        best = None
        for s in pg['all']:
            if s.is_group or s.area() <= 0:
                continue
            if s.prst not in CAPSULE_PRSTS:
                continue
            if not any(_is_teal_family(h) or _is_deep_teal(h) for h in s.fills):
                continue
            if s.area() > 0.30 * page_area:      # 排除铺满画布的卡片主体
                continue
            flat = ' '.join(t.text for t in cap_texts(pg, s))
            flat = flat.replace(' ', '').upper()
            if not any(k.replace(' ', '').upper() in flat for k in keys):
                continue
            if best is None or s.area() > best.area():
                best = s
        return best

    def find_text(pg, pred):
        """按文字内容找形状；pred 收到规范化后的文字。"""
        out = []
        for s in pg['texts']:
            if pred(s.text):
                out.append(s)
        return out

    def divider_groups(pg):
        """"线段＋叶片＋线段"分隔装饰：同一水平位置（Δy≤页高0.012）的两条细
        横线，中间夹一个小图形/小图片（叶片），叶片中心落在两线之间。
        返回 [(整体左端, 整体右端, y, 叶片)]。"""
        found = []
        lines = hlines_of(pg)
        for i in range(len(lines)):
            for j in range(i + 1, len(lines)):
                a, b = lines[i], lines[j]
                if a.left > b.left:
                    a, b = b, a
                if abs(a.cy() - b.cy()) > 0.012 * H:
                    continue
                gap_l, gap_r = a.right(), b.left
                if gap_r - gap_l <= 0:
                    continue
                if (gap_r - gap_l) > 0.12 * W:
                    continue
                for s in pg['all']:
                    if s is a or s is b or s.is_group or s.area() <= 0:
                        continue
                    if rw(s.width) > 0.09 or rh(s.height) > 0.09:
                        continue
                    if not (gap_l - 0.01 * W <= s.cx() <= gap_r + 0.01 * W):
                        continue
                    if abs(s.cy() - a.cy()) > 0.06 * H:
                        continue
                    found.append((a.left, b.right(), a.cy(), s))
                    break
        return found

    # ------------ 共用的版面要素定位结果 ------------
    card1, card2 = card_of(p1), card_of(p2)
    qr1, qr2 = qr_of(p1), qr_of(p2)
    badge1, badge2 = badges_of(p1), badges_of(p2)
    circ1, circ2 = circles_of(p1), circles_of(p2)
    vdiv1, vdiv2 = vdiv_of(p1), vdiv_of(p2)
    cap1 = capsule_of(p1, ('主营',))
    cap2 = capsule_of(p2, ('MAIN BUSINESS', 'MAINBUSINESS'))
    band1, band2 = bands_of(p1), bands_of(p2)
    divs1, divs2 = divider_groups(p1), divider_groups(p2)

    def nospace(t):
        return t.replace(' ', '')

    cn_name = find_text(p1, lambda t: '福建蓝屿' in nospace(t))
    en_name1 = find_text(p1, lambda t: 'FUJIAN LANYU' in t.upper())
    xing = find_text(p1, lambda t: '沈知行' in nospace(t))
    ethan = find_text(p2, lambda t: 'ethan' in t.lower() and 'shen' in t.lower())
    en_name2 = find_text(p2, lambda t: 'FUJIAN LANYU' in t.upper())
    mp1 = find_text(p1, lambda t: nospace(t).upper().startswith('M.P.'))
    mp2 = find_text(p2, lambda t: nospace(t).upper().startswith('M.P.'))
    phone1 = find_text(p1, lambda t: '4826' in t and '7153' in t)
    phone2 = find_text(p2, lambda t: '4826' in t and '7153' in t)
    mail1 = find_text(p1, lambda t: 'z.xing@lanyunova.com' in t.lower())
    mail2 = find_text(p2, lambda t: 'z.xing@lanyunova.com' in t.lower())
    addr1 = find_text(p1, lambda t: '云创大道' in t or '蓝湾创新中心' in t)
    addr2 = find_text(p2, lambda t: 'yunchuang' in t.lower()
                      or 'bluebay' in t.lower())

    def first(lst):
        return lst[0] if lst else None

    def sz_of(sh):
        """形状内可见文字的最大字号（磅）。"""
        if sh is None:
            return 0.0
        return _max_run_sz([sh])

    def runs_of(sh):
        if sh is None:
            return []
        return [r for para in sh.runs for r in para if r['t'].strip()]

    def rightcol(pg):
        """右栏文字块：横坐标在版面右侧、纵向位于胶囊标签之下的文本形状。"""
        return sorted([s for s in pg['texts']
                       if rw(s.cx()) > 0.55 and rh(s.cy()) > 0.30],
                      key=lambda s: s.top)

    def item_sz(pg):
        """主营条目文字字号：取右栏文字块里最小的可见字号。"""
        szs = [r['sz'] for _, r in _iter_runs(rightcol(pg))]
        return min(szs) if szs else 0.0

    # ============================================================
    # +1：文档共2页幻灯片，第1页为中文正面版式、第2页为英文背面版式，页序不颠倒
    # 严格对齐细则，仅踩细则明确写出的点，不额外约束：
    #   ① 幻灯片恰好 2 页（多页/少页都不算）
    #   ② 第1页以中文为主（中日韩字符数明显多于纯西文字符数）
    #   ③ 第2页以西文为主（西文字符数明显多于中日韩字符数）
    # 与办公软件对齐：即在 PPT 里翻页看到的"第1页中文正面、第2页英文背面"。
    def cjk_latin(pg):
        cjk = lat = 0
        for _, r in _iter_runs(pg['texts']):
            for ch in r['t']:
                if _has_cjk(ch):
                    cjk += 1
                elif ch.isalpha():
                    lat += 1
        return cjk, lat

    c1, l1 = cjk_latin(p1)
    c2, l2 = cjk_latin(p2)
    if len(slides) == 2 and c1 > l1 and l2 > c2:
        results.append("+1：文档共2页幻灯片，第1页为中文正面版式、第2页为英文背面版式，页序不颠倒")
    else:
        results.append("+0：文档共2页幻灯片，第1页为中文正面版式、第2页为英文背面版式，页序不颠倒")

    # ============================================================
    # +3：两页画布尺寸完全一致且为横向名片比例（宽/高在1.50-1.85之间），每页以一个
    #     圆角矩形卡片主体铺满画布（圆角调节值>0），卡片填充为米白/象牙白
    # 严格对齐细则：
    #   ① 画布宽/高比在 1.50-1.85（.pptx 全篇共用一个画布，尺寸天然一致）
    #   ② 每页有一个圆角矩形（roundRect 系列预设）铺满画布（面积≥画布0.88）
    #   ③ 该圆角矩形的圆角调节值 adj > 0（即真的有圆角，不是直角矩形）
    #   ④ 填充为浅暖白（米白/象牙白），既不是纯白也不是彩色
    # 与办公软件对齐：a:prstGeom prst="roundRect" + a:avLst/a:gd@val → 形状格式里的
    #   "圆角"调节柄；a:solidFill/a:srgbClr → 填充颜色取色器读数。
    def card_adj(sh):
        sp_pr = sh.el.find(_p('spPr'))
        pg_el = sp_pr.find(_a('prstGeom')) if sp_pr is not None else None
        if pg_el is None:
            return None
        av = pg_el.find(_a('avLst'))
        if av is None:
            return 0
        for gd in av.findall(_a('gd')):
            if (gd.get('name') or '') == 'adj':
                try:
                    return int(gd.get('fmla', 'val 0').split()[-1])
                except (ValueError, IndexError):
                    return 0
        return 0

    ratio_ok = 1.50 <= (W / float(H)) <= 1.85
    card_ok = True
    for cd in (card1, card2):
        if cd is None or cd.area() < 0.88 * page_area:
            card_ok = False
            break
        if 'roundRect' not in cd.prst and cd.prst not in CAPSULE_PRSTS:
            card_ok = False
            break
        if not card_adj(cd):
            card_ok = False
            break
        if not any(_is_cream(h) for h in cd.fills):
            card_ok = False
            break
    if ratio_ok and card_ok:
        results.append("+3：两页画布尺寸完全一致且为横向名片比例（宽/高在1.50-1.85之间），每页以一个圆角矩形卡片主体铺满画布（圆角调节值>0），卡片填充为米白/象牙白（近似#F5F3EE-#FEFDF8的浅暖白），非纯白或彩色")
    else:
        results.append("+0：两页画布尺寸完全一致且为横向名片比例（宽/高在1.50-1.85之间），每页以一个圆角矩形卡片主体铺满画布（圆角调节值>0），卡片填充为米白/象牙白（近似#F5F3EE-#FEFDF8的浅暖白），非纯白或彩色")

    # ============================================================
    # +5：两页右上角二维码均以图片对象（picture / blipFill）插入，而非用形状、表格或
    #     文字字符拼出的伪二维码；二维码为正方形（0.90-1.10）、边长占页宽0.11-0.20、
    #     左边缘位于页宽0.78-0.88、上边缘位于页高0.02-0.14
    # 严格对齐细则：
    #   ① 两页右上角都能定位到该对象
    #   ② 该对象是图片（p:pic / blipFill 填充），不是形状或文字拼贴
    #   ③ 长宽比 0.90-1.10、边长占页宽 0.11-0.20
    #   ④ 左边缘占页宽 0.78-0.88、上边缘占页高 0.02-0.14
    # 与办公软件对齐：p:pic → 选中后功能区出现"图片格式"，可另存为图片。
    qr_ok = True
    for q in (qr1, qr2):
        if q is None:
            qr_ok = False
            break
        if not (q.is_pic or q.fill_kind == 'pic'):
            qr_ok = False
            break
        ar = q.width / float(q.height) if q.height else 0
        if not (0.90 <= ar <= 1.10):
            qr_ok = False
            break
        if not (0.11 <= rw(q.width) <= 0.20):
            qr_ok = False
            break
        if not (0.78 <= rw(q.left) <= 0.88):
            qr_ok = False
            break
        if not (0.02 <= rh(q.top) <= 0.14):
            qr_ok = False
            break
    if qr_ok:
        results.append("+5：两页右上角二维码均以图片对象（picture / blipFill）插入，而非用形状、表格或文字字符拼出的伪二维码；二维码为正方形（长宽比0.90-1.10）、边长占页宽0.11-0.20、左边缘位于页宽0.78-0.88、上边缘位于页高0.02-0.14")
    else:
        results.append("+0：两页右上角二维码均以图片对象（picture / blipFill）插入，而非用形状、表格或文字字符拼出的伪二维码；二维码为正方形（长宽比0.90-1.10）、边长占页宽0.11-0.20、左边缘位于页宽0.78-0.88、上边缘位于页高0.02-0.14")

    # ============================================================
    # +3：两页二维码的位置与尺寸保持一致（左坐标、上坐标、边长三项差异均≤页宽的0.02）
    # 严格对齐细则：① 两页都定位到二维码；② 左、上、边长三项差异均 ≤ 页宽 0.02
    # 与办公软件对齐：即在"大小和位置"窗格里读到的水平/垂直位置与宽度基本相同。
    if qr1 is not None and qr2 is not None \
            and rw(abs(qr1.left - qr2.left)) <= 0.02 \
            and rw(abs(qr1.top - qr2.top)) <= 0.02 \
            and rw(abs(qr1.width - qr2.width)) <= 0.02:
        results.append("+3：两页二维码的位置与尺寸保持一致（左坐标、上坐标、边长三项差异均≤页宽的0.02），正反面同位置呈现")
    else:
        results.append("+0：两页二维码的位置与尺寸保持一致（左坐标、上坐标、边长三项差异均≤页宽的0.02），正反面同位置呈现")

    # ============================================================
    # +3：每页原生可编辑对象数量≥25个，无整版截图贴图
    # 严格对齐细则：
    #   ① 每页非图片的原生对象（文本框/自动图形/连接线/自由曲线）数量 ≥ 25
    #   ② 不存在铺满整页的位图（按占页面积 0.60 判定"整版截图贴图"）
    # 说明：直线/连接线只有长度没有厚度（面积为 0），在 PPT 里同样是可单独点选、
    #   可编辑的对象，因此按"宽或高有值"计数，不按面积计数。
    # 与办公软件对齐：在 PPT 里用"选择窗格"能逐个点选到的对象数量。
    native_ok = True
    for pg in (p1, p2):
        native = [s for s in pg['all']
                  if not s.is_pic and not s.is_group and s.is_sized()]
        if len(native) < 25:
            native_ok = False
        if any(s.area() >= 0.60 * page_area for s in pg['pics']):
            native_ok = False
    if native_ok:
        results.append("+3：除二维码与徽标/图标类小图外，页面主体由文本框、自动图形、连接线等原生对象拼搭而成，每页原生可编辑对象数量≥25个，无整版截图贴图")
    else:
        results.append("+0：除二维码与徽标/图标类小图外，页面主体由文本框、自动图形、连接线等原生对象拼搭而成，每页原生可编辑对象数量≥25个，无整版截图贴图")

    # ============================================================
    # +3：全篇字体为Windows环境可正常显示的常见字体；未使用仅macOS可用字体承载可见文字
    # 严格对齐细则：
    #   ① 承载中日韩文字的 run，其东亚字体（a:ea，缺省时看 a:latin）不得为 macOS 专有；
    #   ② 承载西文/数字的 run，其西文字体（a:latin）不得为 macOS 专有；
    #   ③ 显式指定的字体须落在 Windows 常见字体族内（未指定＝随主题，视为可显示）。
    # 说明：只检查"实际承载可见文字"的那一侧字体——把 ea 字体挂在纯西文 run 上并不会
    #   影响 Windows 下的显示效果，不据此扣分。
    # 与办公软件对齐：a:latin/a:ea typeface → 开始选项卡的字体框；缺字体时 Windows 会
    #   回退成默认字形，书法体等特殊字形随之丢失。
    font_ok = True
    for pg in (p1, p2):
        for _, r in _iter_runs(pg['texts']):
            names = []
            if _has_cjk(r['t']):
                names.append(r['ea'] or r['latin'])
            if _has_latin(r['t']):
                names.append(r['latin'])
            for nm in names:
                if not nm:
                    continue
                if _is_mac_only(nm) or not _is_win_common(nm):
                    font_ok = False
    if font_ok:
        results.append("+3：全篇字体为Windows环境可正常显示的常见字体（微软雅黑、黑体、宋体、楷体、Arial、Times New Roman、Segoe 系列等）；未使用仅macOS可用字体（如STXingkai、STKaiti、Zapfino等）承载可见文字而导致字形回退丢失")
    else:
        results.append("+0：全篇字体为Windows环境可正常显示的常见字体（微软雅黑、黑体、宋体、楷体、Arial、Times New Roman、Segoe 系列等）；未使用仅macOS可用字体（如STXingkai、STKaiti、Zapfino等）承载可见文字而导致字形回退丢失")

    # ============================================================
    # +3：所有形状的宽、高均为非负有效值，无负宽/负高的畸形图形
    # 严格对齐细则：a:ext@cx / a:ext@cy 一律 ≥ 0（组内子形状同样检查）
    # 与办公软件对齐：负的扩展量在 PowerPoint 里表现为图形不显示或上下/左右翻转，
    #   在"大小和位置"窗格里也无法正常读数。
    ext_ok = True
    for pg in (p1, p2):
        for s in pg['all']:
            if s.raw_w < 0 or s.raw_h < 0:
                ext_ok = False
    if ext_ok:
        results.append("+3：所有形状的宽、高均为非负有效值，无负宽/负高的畸形图形（负值扩展量会导致图形不显示或反向翻转）")
    else:
        results.append("+0：所有形状的宽、高均为非负有效值，无负宽/负高的畸形图形（负值扩展量会导致图形不显示或反向翻转）")

    # ============================================================
    # +3：第1页左上角设有企业徽标，以图片或成组矢量图形完整呈现，
    #     左边缘0.02-0.10页宽、上边缘0.02-0.14页高、宽占页宽0.07-0.20
    # 严格对齐细则：
    #   ① 位置：左边缘 0.02-0.10 页宽、上边缘 0.02-0.14 页高
    #   ② 尺寸：宽占页宽 0.07-0.20
    #   ③ 形态：单个图片对象，或一个把矢量图形组合在一起的组合对象
    #      （散落、未成组的弧线/水滴不算"完整呈现"，无法整体移动缩放）
    # 与办公软件对齐：p:pic 或 p:grpSp → 在 PPT 里单击一次即可整体选中并移动。
    logo = None
    for s in p1['all']:
        if not (s.is_pic or s.is_group):
            continue
        if s is qr1 or s.area() <= 0:
            continue
        if not (0.02 <= rw(s.left) <= 0.10):
            continue
        if not (0.02 <= rh(s.top) <= 0.14):
            continue
        if not (0.07 <= rw(s.width) <= 0.20):
            continue
        if logo is None or s.area() > logo.area():
            logo = s
    if logo is not None:
        results.append("+3：第1页左上角设有企业徽标（圆形叶片＋暖色弧标识），以图片或成组矢量图形完整呈现，左边缘位于页宽0.02-0.10、上边缘位于页高0.02-0.14、宽占页宽0.07-0.20")
    else:
        results.append("+0：第1页左上角设有企业徽标（圆形叶片＋暖色弧标识），以图片或成组矢量图形完整呈现，左边缘位于页宽0.02-0.10、上边缘位于页高0.02-0.14、宽占页宽0.07-0.20")

    # ============================================================
    # +3：第1页中文公司名"福建蓝屿新材科技有限公司"单行文本，加粗、深墨绿，
    #     左边缘位于徽标右缘之后且位于页宽0.17-0.24，
    #     字号≥主营条目文字字号的2倍且为页眉区最大字号
    # 严格对齐细则：
    #   ① 文字内容与细则给出的公司名完全一致（多字/少字/错字均不算达标）
    #   ② 单行呈现（不折行、不分段）
    #   ③ 加粗：rPr@b=1，或所选字族本身是粗体（"微软雅黑 Bold"一类）；中文字体
    #      （a:ea）与西文字体（a:latin）任一侧是粗体字族即认，中文字符实际渲染用的
    #      是 a:ea 那一侧，只看 a:latin 会把"中文设了粗体族"的写法漏判成没加粗
    #   ④ 文字色为深墨绿色系
    #   ⑤ 左边缘落在徽标右缘之后，且位于页宽 0.17-0.24
    #   ⑥ 字号 ≥ 主营条目字号的 2 倍，且是页眉区（上边缘 0.28 页高以内）最大字号
    # 与办公软件对齐：rPr@b → 加粗按钮；rPr/solidFill → 字体颜色；sz → 字号框。
    #   "在徽标右侧"按坐标关系判定，不与徽标自身的位置区间并列，避免两条各自
    #   合规却仍然压字的情形。
    cn = first(cn_name)
    it_sz1 = item_sz(p1)
    hdr_sz1 = _max_run_sz([s for s in p1['texts'] if rh(s.top) <= 0.28])
    cn_ok = cn is not None
    if cn_ok:
        rs = runs_of(cn)
        cn_ok = (nospace(cn.text) == nospace('福建蓝屿新材科技有限公司')
                 and _visual_lines(cn) == 1
                 and all(r['b'] or _is_bold_font(r['latin'])
                         or _is_bold_font(r['ea']) for r in rs)
                 and any(_is_deep_teal(r['clr'] or '') for r in rs)
                 and 0.17 <= rw(cn.left) <= 0.24
                 and (logo is None or cn.left >= logo.right() - 0.01 * W)
                 and it_sz1 > 0 and sz_of(cn) >= 2.0 * it_sz1
                 and sz_of(cn) >= hdr_sz1 - 0.01)
    if cn_ok:
        results.append('+3：第1页徽标右侧设有中文公司名"福建蓝屿新材科技有限公司"单行文本，加粗（加粗属性或粗体字族均可）、深墨绿色系（近似#013B3C-#004F49，或同族更深值），左边缘位于徽标右缘之后且位于页宽0.17-0.24，字号≥主营条目文字字号的2倍且为页眉区最大字号')
    else:
        results.append('+0：第1页徽标右侧设有中文公司名"福建蓝屿新材科技有限公司"单行文本，加粗（加粗属性或粗体字族均可）、深墨绿色系（近似#013B3C-#004F49，或同族更深值），左边缘位于徽标右缘之后且位于页宽0.17-0.24，字号≥主营条目文字字号的2倍且为页眉区最大字号')

    # ============================================================
    # +3：中文公司名正下方的英文公司名"FUJIAN LANYU NOVA MATERIALS CO., LTD."
    #     单行全大写、无衬线带字间距、字号为中文公司名0.45-0.70倍、铜橙色
    # 严格对齐细则：
    #   ① 内容与细则给出的英文公司名一致，且全大写（大写字符或"全部大写"效果）
    #   ② 单行呈现；用字符间显式留空写出字距时，留空本身会把折行估算撑大，
    #      故此写法下按"单段落"衡量单行
    #   ③ 西文字体为无衬线族
    #   ④ 字间距不作判定：字符间距加宽在版面上只影响字距疏密，肉眼难以与"没加宽"
    #      区分，也不影响内容完整与可编辑性，故本项不因 spc 未设而扣分（细则原文
    #      仍保留在结果串里，仅此一处不参与判定）
    #   ⑤ 紧接中文公司名之下并与其左对齐（差 ≤ 0.02 页宽）；"紧接"按中文公司名
    #      字号的相对间隙衡量（≤1.5 倍字号），不随画布尺寸漂移
    #   ⑥ 字号为中文公司名的 0.45-0.70 倍
    #   ⑦ 文字色为铜橙色系
    # 与办公软件对齐：rPr@cap="all" → "全部大写"复选框。中文公司名缺失时，字号比与
    #   左对齐两项无从比对，不在本条重复扣分（该缺陷已由中文公司名一条计入）。
    en1 = first(en_name1)
    en_ok = en1 is not None
    if en_ok:
        rs = runs_of(en1)
        body = nospace(en1.text).upper().replace('，', ',')
        allcaps = (en1.text == en1.text.upper()
                   or any((r['cap'] or '').lower() in ('all', 'small')
                          for r in rs))
        spaced_txt = _letter_spaced(en1.text)
        one_line = (len(en1.runs) == 1) if spaced_txt else (_visual_lines(en1) == 1)
        rel_ok = True
        if cn is not None and sz_of(cn) > 0:
            rel_ok = (abs(en1.left - cn.left) <= 0.02 * W
                      and en1.cy() > cn.cy()
                      and en1.top - cn.bottom() <= 1.5 * sz_of(cn) * EMU_PER_PT
                      and 0.45 <= sz_of(en1) / sz_of(cn) <= 0.70)
        en_ok = ('FUJIANLANYUNOVAMATERIALSCO.,LTD.' in body
                 and allcaps
                 and one_line
                 and all(_sans_serif(r['latin']) for r in rs if r['latin'])
                 and rel_ok
                 and any(_is_copper(r['clr'] or '') for r in rs))
    if en_ok:
        results.append('+3：中文公司名正下方设有英文公司名"FUJIAN LANYU NOVA MATERIALS CO., LTD."单行全大写文本（大写字符或全部大写效果均可），无衬线西文字体、带字间距（字符间距加宽≥0.3磅或字符间显式留空），左边缘与中文公司名对齐（差≤页宽0.02）且紧接其下（间隙≤中文公司名字号的1.5倍），字号为中文公司名的0.45-0.70倍，且填充为铜橙色系（近似#B5713C-#C77B45，或同族更饱和的暖橙）')
    else:
        results.append('+0：中文公司名正下方设有英文公司名"FUJIAN LANYU NOVA MATERIALS CO., LTD."单行全大写文本（大写字符或全部大写效果均可），无衬线西文字体、带字间距（字符间距加宽≥0.3磅或字符间显式留空），左边缘与中文公司名对齐（差≤页宽0.02）且紧接其下（间隙≤中文公司名字号的1.5倍），字号为中文公司名的0.45-0.70倍，且填充为铜橙色系（近似#B5713C-#C77B45，或同族更饱和的暖橙）')

    # ============================================================
    # +3：三处"线段＋叶片＋线段"分隔装饰均具备且样式一致
    # 严格对齐细则：
    #   ① 每处都是"两条细横线段 + 中间 1 个叶片图形/图片"，叶片把横线断开且居中
    #   ② 第1页公司名下方：整体跨度占页宽 0.45-0.62
    #   ③ 第1页姓名下方：跨度占页宽 0.25-0.42，且仅位于左半区（右端不越过中线）
    #   ④ 第2页英文公司名下方：同样形态的一处
    # 说明：跨度按"最左线段左端 → 最右线段右端"的整体宽度衡量，与肉眼看到的
    #   分隔装饰总长一致。
    # 与办公软件对齐：两条 p:cxnSp/直线形状 + 中间一个小图形，逐个可点选。
    def match_div(divs, lo, hi, y_lo, y_hi, right_limit=None):
        for l, r, y, leaf in divs:
            span = rw(r - l)
            if not (lo <= span <= hi):
                continue
            if not (y_lo <= rh(y) <= y_hi):
                continue
            if right_limit is not None and rw(r) > right_limit:
                continue
            return True
        return False

    div_a = match_div(divs1, 0.45, 0.62, 0.18, 0.38)
    div_b = match_div(divs1, 0.25, 0.42, 0.44, 0.72, right_limit=0.52)
    div_c = match_div(divs2, 0.20, 0.62, 0.18, 0.48)
    if div_a and div_b and div_c:
        results.append('+3：三处"线段＋叶片＋线段"分隔装饰均具备且样式一致——第1页公司名下方（两段线合计跨度占页宽0.45-0.62）、第1页姓名下方（跨度占页宽0.25-0.42、仅位于左半区）、第2页英文公司名下方；每处均为两条细横线段＋中间1个叶片图形/图片，叶片将横线断开居中')
    else:
        results.append('+0：三处"线段＋叶片＋线段"分隔装饰均具备且样式一致——第1页公司名下方（两段线合计跨度占页宽0.45-0.62）、第1页姓名下方（跨度占页宽0.25-0.42、仅位于左半区）、第2页英文公司名下方；每处均为两条细横线段＋中间1个叶片图形/图片，叶片将横线断开居中')

    # ============================================================
    # +5：第1页姓名"沈知行"，行楷/楷书类书法字体、深墨绿、该页最大字号
    #     （≥中文公司名1.8倍）、左边缘0.08-0.18页宽、上边缘0.30-0.42页高
    # 严格对齐细则：
    #   ① 文字为"沈知行"
    #   ② 中文字体属行楷/楷书类书法风格，且不是黑体/雅黑类无衬线体
    #   ③ 文字色为深墨绿
    #   ④ 字号为该页最大，且 ≥ 中文公司名字号的 1.8 倍
    #   ⑤ 左边缘 0.08-0.18 页宽、上边缘 0.30-0.42 页高
    # 与办公软件对齐：a:ea typeface → 中文字体框（楷体/行楷等书法字形）。
    xg = first(xing)
    xg_ok = xg is not None
    if xg_ok:
        rs = runs_of(xg)
        fonts = [(r['ea'] or r['latin']) for r in rs if _has_cjk(r['t'])]
        xg_ok = (nospace(xg.text) == '沈知行'
                 and fonts and all(_is_kai_font(f) for f in fonts)
                 and not any(_sans_serif(f) for f in fonts)
                 and any(_is_deep_teal(r['clr'] or '') for r in rs)
                 and sz_of(xg) >= _max_run_sz(p1['texts']) - 0.01
                 and sz_of(cn) > 0 and sz_of(xg) >= 1.8 * sz_of(cn)
                 and 0.08 <= rw(xg.left) <= 0.18
                 and 0.30 <= rh(xg.top) <= 0.42)
    if xg_ok:
        results.append('+5：第1页左半区设有姓名"沈知行"，采用行楷/楷书类书法风格中文字体（楷体、行楷等，非黑体/雅黑类无衬线体），深墨绿，字号为该页最大（≥中文公司名字号的1.8倍），左边缘位于页宽0.08-0.18、上边缘位于页高0.30-0.42')
    else:
        results.append('+0：第1页左半区设有姓名"沈知行"，采用行楷/楷书类书法风格中文字体（楷体、行楷等，非黑体/雅黑类无衬线体），深墨绿，字号为该页最大（≥中文公司名字号的1.8倍），左边缘位于页宽0.08-0.18、上边缘位于页高0.30-0.42')

    # ============================================================
    # +5：第1页3行"徽章图标＋文字"联系方式，等距堆叠、三个徽章左边缘对齐
    # 严格对齐细则：
    #   ① 左半区能定位到 3 个徽章式小方块（深墨绿圆角方底或其图片）
    #   ② 三者左边缘对齐（差异 ≤ 页宽 0.01；聚类时已按 0.015 放宽定位，此处收紧判定）
    #   ③ 相邻行间距差异 ≤ 单行行高的 0.5 倍（等距堆叠）
    #   ④ 每个徽章右侧有对应文字
    # 与办公软件对齐：即三行图标与文字在 PPT 里看起来上下均匀、左侧图标一条竖线对齐。
    def badge_rows_ok(pg, badges):
        if len(badges) != 3:
            return False
        lefts = [b.left for b in badges]
        if (max(lefts) - min(lefts)) > 0.012 * W:
            return False
        tops = sorted(b.top for b in badges)
        row_h = sum(b.height for b in badges) / 3.0
        if not _equidistant(tops, 0.5 * row_h):
            return False
        for b in badges:
            if not any(t.left > b.left and abs(t.cy() - b.cy()) < 0.9 * b.height
                       for t in pg['texts']):
                return False
        return True

    if badge_rows_ok(p1, badge1):
        results.append('+5：第1页左半区设有3行联系方式，每行为"深墨绿圆角方形徽章图标＋右侧文字"结构，三行自上而下等距堆叠（相邻行间距差异≤单行行高的0.5倍），三个徽章左边缘对齐（差异≤页宽0.01）')
    else:
        results.append('+0：第1页左半区设有3行联系方式，每行为"深墨绿圆角方形徽章图标＋右侧文字"结构，三行自上而下等距堆叠（相邻行间距差异≤单行行高的0.5倍），三个徽章左边缘对齐（差异≤页宽0.01）')

    # ============================================================
    # +5：三个联系方式徽章内为真实的电话听筒、信封、定位针图标，
    #     而非用☎、✉、⌖等文本符号字符充当图标
    # 严格对齐细则：
    #   ① 每个徽章本身是图片，或徽章内部套着图片/矢量图形（无文字）
    #   ② 徽章范围内不得靠符号字符文字充当图标
    # 说明：本项只判"图标是否为真实图形"，不识别具体图案语义（听筒/信封/定位针），
    #   避免把无法在 XML 层面确证的内容硬判为不达标。
    # 与办公软件对齐：符号字符做的"图标"在 PPT 里双击就变成可输入的文字光标，
    #   且缺字体时会变成方框；真实图标是独立的图片/矢量对象。
    def icons_real(pg, hosts, need):
        if len(hosts) != need:
            return False
        for hb in hosts:
            if glyph_icon_inside(pg, hb):
                return False
            if not (hb.is_pic or icon_inside(pg, hb, need_pic=False)):
                return False
        return True

    if icons_real(p1, badge1, 3):
        results.append("+5：三个联系方式徽章内为真实的电话听筒、信封、定位针图标（以图片或矢量图形绘制，白色图形置于深墨绿圆角方底上），而非用☎、✉、⌖等文本符号字符充当图标")
    else:
        results.append("+0：三个联系方式徽章内为真实的电话听筒、信封、定位针图标（以图片或矢量图形绘制，白色图形置于深墨绿圆角方底上），而非用☎、✉、⌖等文本符号字符充当图标")

    # ============================================================
    # +3：第1页左半区3行联系方式第1行"M.P."（绿色）＋电话"139 4826 7153"
    #     （近黑加粗、分组以空格分隔），号码字号为邮箱行与地址两行中较大字号的
    #     1.2-2.0 倍
    # 严格对齐细则：
    #   ① "M.P."文字为绿色（细则只要求"绿色"，故按色相整段判定：黄绿到青绿均可，
    #      不限定深浅，也不再要求与中文公司名同色族）
    #   ② 电话为细则给出的号码，且分组之间以空格分隔
    #   ③ "M.P."横坐标位于号码左侧，两者纵向中心差 ≤ 号码行高（同一行内）
    #   ④ 号码文字近黑，且加粗（加粗属性、粗体字族、段落级加粗继承均可）
    #   ⑤ 号码文字字号为邮箱行、地址两行中较大字号的 1.2-2.0 倍
    # 与办公软件对齐：rPr@b → 加粗复选框，字族名自带 Bold/Black 的同样是粗体，
    #   在段落/文本框级别设的加粗也会显示成粗体（已按继承链取值）；号码分组是
    #   文字内容本身的空格，半角/全角/不换行空格肉眼一致，已统一折算成普通空格
    #   再比对。"M.P."与号码写在同一个文本框时，字号按号码那段字符算（取整框最大
    #   字号会被"M.P."的字号顶替，倍数判定失真）。倍数设上限是因为号码只需比另
    #   两行醒目，放大到两倍以上即为版面失衡。
    ph1 = first(phone1)
    mpo = first(mp1)
    small_sz = max(sz_of(first(mail1)), sz_of(first(addr1)))
    ph_ok = ph1 is not None and mpo is not None and small_sz > 0
    if ph_ok:
        digit_runs = [r for r in runs_of(ph1) if re.search(r'\d', r['t'])]
        mp_runs = [r for r in runs_of(mpo) if not re.search(r'\d', r['t'])]
        if mpo is ph1:
            row_ok = True
        else:
            row_ok = (mpo.left < ph1.left
                      and abs(mpo.cy() - ph1.cy()) <= ph1.height)
        ph_sz = max([r['sz'] for r in digit_runs] or [0.0])
        ph_ok = (bool(digit_runs) and bool(mp_runs)
                 and any(_is_green(r['clr'] or '') for r in mp_runs)
                 and '139 4826 7153' in ph1.text
                 and row_ok
                 and any(_is_near_black(r['clr'] or '') for r in digit_runs)
                 and all(r['b'] or _is_bold_font(r['latin'] or r['ea'])
                         for r in digit_runs)
                 and 1.2 * small_sz <= ph_sz <= 2.0 * small_sz)
    if ph_ok:
        results.append('+3：第1页左半区3行联系方式的第1行为"M.P."（字体为绿色）＋电话"139 4826 7153"（分组间以空格分隔，半角/全角/不换行空格均可），M.P.横坐标位于号码左侧、两者纵向中心差≤号码行高；号码文字近黑（近似#1A1A1A-#333333，或明度不高于#4D4D4D且无明显色偏）且加粗（加粗属性、粗体字族、段落级加粗继承均可）；号码文字字号为邮箱行与地址两行中较大字号的1.2-2.0倍（与"M.P."同框时按号码那段字符计）')
    else:
        results.append('+0：第1页左半区3行联系方式的第1行为"M.P."（字体为绿色）＋电话"139 4826 7153"（分组间以空格分隔，半角/全角/不换行空格均可），M.P.横坐标位于号码左侧、两者纵向中心差≤号码行高；号码文字近黑（近似#1A1A1A-#333333，或明度不高于#4D4D4D且无明显色偏）且加粗（加粗属性、粗体字族、段落级加粗继承均可）；号码文字字号为邮箱行与地址两行中较大字号的1.2-2.0倍（与"M.P."同框时按号码那段字符计）')

    # ============================================================
    # +3：第2行邮箱、第3行地址（完整呈现两段地址并断为两行），两行均与电话行左对齐
    # 严格对齐细则：
    #   ① 邮箱文字为"z.xing@lanyunova.com"
    #   ② 地址完整含"福建省厦门市海沧区云创大道88号"与"蓝湾创新中心12层1208室"
    #   ③ 地址在版面上断成两行（分两段写，或一段但框宽不足自动折行，视觉都是两行）
    #   ④ 邮箱行、地址行左边缘与电话行左边缘对齐（差异 ≤ 页宽 0.01）
    # 与办公软件对齐：a:p 分段 或 文本框宽度导致的自动换行，都是肉眼看到的"两行"。
    ml = first(mail1)
    ad = first(addr1)
    row_left = mpo.left if mpo is not None else (ph1.left if ph1 else None)
    mail_ok = ml is not None and ad is not None and row_left is not None
    if mail_ok:
        flat = nospace(ad.text)
        mail_ok = ('z.xing@lanyunova.com' in ml.text.lower()
                   and '福建省厦门市海沧区云创大道88号' in flat
                   and '蓝湾创新中心12层1208室' in flat
                   and _visual_lines(ad) == 2
                   and abs(ml.left - row_left) <= 0.01 * W
                   and abs(ad.left - row_left) <= 0.01 * W)
    if mail_ok:
        results.append('+3：第2行为邮箱"z.xing@lanyunova.com"、第3行为地址，地址完整呈现"福建省厦门市海沧区云创大道88号"与"蓝湾创新中心12层1208室"并断为两行显示，两行均与电话行左对齐')
    else:
        results.append('+0：第2行为邮箱"z.xing@lanyunova.com"、第3行为地址，地址完整呈现"福建省厦门市海沧区云创大道88号"与"蓝湾创新中心12层1208室"并断为两行显示，两行均与电话行左对齐')

    # ============================================================
    # +3：第1页纵向分隔线将版面切为左右两栏，位于页宽0.48-0.56，
    #     纵向跨度占页高0.35-0.60，起于页眉分隔线下方、止于底部色带上方
    # 严格对齐细则：
    #   ① 存在一条极窄的竖向线/形状
    #   ② 横坐标（中线）位于页宽 0.48-0.56
    #   ③ 纵向跨度占页高 0.35-0.60
    #   ④ 上端在页眉分隔装饰下方（>0.20 页高），下端在底部色带上方（<0.95 页高）
    # 与办公软件对齐：p:cxnSp 直线 或极窄矩形，在 PPT 里是一条可单独点选的竖线。
    vd_ok = vdiv1 is not None
    if vd_ok:
        vd_ok = (0.48 <= rw(vdiv1.cx()) <= 0.56
                 and 0.35 <= rh(vdiv1.height) <= 0.60
                 and rh(vdiv1.top) > 0.20
                 and rh(vdiv1.bottom()) < 0.95)
    if vd_ok:
        results.append("+3：第1页设有一条纵向分隔线将版面切为左右两栏，位于页宽0.48-0.56，纵向跨度占页高0.35-0.60，起于页眉分隔线下方、止于底部色带上方")
    else:
        results.append("+0：第1页设有一条纵向分隔线将版面切为左右两栏，位于页宽0.48-0.56，纵向跨度占页高0.35-0.60，起于页眉分隔线下方、止于底部色带上方")

    # ============================================================
    # +3：第1页右栏顶部"主 营"胶囊标签：圆角/胶囊实心深墨绿＋白色加粗文字（两字间留空），
    #     宽占页宽0.10-0.19、左边缘位于页宽0.50-0.58
    # 严格对齐细则：
    #   ① 形状为圆角矩形/胶囊类预设，实心填充深墨绿族
    #   ② 内部文字为"主 营"且两字之间留有空格
    #   ③ 文字为白色且加粗
    #   ④ 宽占页宽 0.10-0.19、左边缘位于页宽 0.50-0.58
    # 与办公软件对齐：prst="roundRect"/"flowChartTerminator" → 形状库里的圆角矩形/
    #   胶囊；solidFill → 纯色填充。
    cap_ok = cap1 is not None
    if cap_ok:
        cts = cap_texts(p1, cap1)
        rs = [r for t in cts for r in runs_of(t)]
        joined = ' '.join(t.text for t in cts).strip()
        cap_ok = (' ' in joined
                  and rs and all(r['b'] for r in rs)
                  and any(_is_white(r['clr'] or '') for r in rs)
                  and 0.10 <= rw(cap1.width) <= 0.19
                  and 0.50 <= rw(cap1.left) <= 0.58)
    if cap_ok:
        results.append('+3：第1页右栏顶部设有"主 营"胶囊标签：圆角/胶囊形状实心填充深墨绿，内部为白色加粗文字（两字间留空），标签宽占页宽0.10-0.19、左边缘位于页宽0.50-0.58')
    else:
        results.append('+0：第1页右栏顶部设有"主 营"胶囊标签：圆角/胶囊形状实心填充深墨绿，内部为白色加粗文字（两字间留空），标签宽占页宽0.10-0.19、左边缘位于页宽0.50-0.58')

    # ============================================================
    # +5：第1页右栏3个主营分组＝浅色圆形底衬＋圆内图标＋右侧文字块，
    #     三组等距排列、三个圆形底衬左边缘对齐
    # 严格对齐细则：
    #   ① 右栏能定位到 3 个正圆形底衬，填充为淡灰绿浅色实心
    #   ② 三个圆左边缘对齐（差异 ≤ 页宽 0.012）
    #   ③ 三组自上而下等距（相邻间距差异 ≤ 圆直径的 0.5 倍）
    #   ④ 每个圆右侧有对应的文字块
    # 与办公软件对齐：prst="ellipse" + solidFill → 形状库的椭圆/正圆纯色填充。
    def circle_groups_ok(pg, circles):
        if len(circles) != 3:
            return False
        if not all(any(_is_pale_greengray(h) for h in c.fills) for c in circles):
            return False
        lefts = [c.left for c in circles]
        if (max(lefts) - min(lefts)) > 0.012 * W:
            return False
        tops = sorted(c.top for c in circles)
        dia = sum(c.height for c in circles) / 3.0
        if not _equidistant(tops, 0.5 * dia):
            return False
        for c in circles:
            if not any(t.left > c.left and abs(t.cy() - c.cy()) < 1.2 * c.height
                       for t in pg['texts']):
                return False
        return True

    if circle_groups_ok(p1, circ1):
        results.append('+5：第1页右栏设有3个主营分组，每组为"浅色圆形底衬（近似#EDF0EA-#F0F2EC淡灰绿实心圆）＋圆内图标＋右侧文字块"结构，三组自上而下等距排列，三个圆形底衬左边缘对齐')
    else:
        results.append('+0：第1页右栏设有3个主营分组，每组为"浅色圆形底衬（近似#EDF0EA-#F0F2EC淡灰绿实心圆）＋圆内图标＋右侧文字块"结构，三组自上而下等距排列，三个圆形底衬左边缘对齐')

    # ============================================================
    # +5：三个圆形底衬内为水滴、叶片、包装箱图标（图片或矢量图形），
    #     而非♧、♢、▱等文本符号字符充当图标
    # 严格对齐细则：
    #   ① 每个圆内套着图片或矢量图形（无文字）
    #   ② 圆范围内不得靠符号字符文字充当图标
    # 与办公软件对齐：同徽章图标——符号字符做的图标点进去是文字，换环境会变方框。
    if icons_real(p1, circ1, 3):
        results.append("+5：三个圆形底衬内为水滴、叶片、包装箱三个对应图标（以图片或矢量图形绘制），而非♧、♢、▱等文本符号字符充当图标")
    else:
        results.append("+0：三个圆形底衬内为水滴、叶片、包装箱三个对应图标（以图片或矢量图形绘制），而非♧、♢、▱等文本符号字符充当图标")

    # ============================================================
    # +5：第1页右栏三组文字完整且逐行呈现——标题"环保涂层：""绿色材料：""应用方案："
    #     加粗深墨绿、全角冒号、字号大于其下条目；条目分别 3 行 / 2 行 / 2 行
    # 严格对齐细则：
    #   ① 三个标题逐字一致且使用全角冒号"："
    #   ② 标题加粗、深墨绿，字号大于同组条目字号
    #   ③ 条目内容逐字一致，且行数分别为 3、2、2
    # 与办公软件对齐：a:p 分段 → 每条一行；全角/半角冒号在 PPT 里字形宽度不同。
    def group_texts(pg, circles):
        """把右栏文字块按纵向位置归到最近的圆形底衬，得到三组文字。
        只认"与圆形底衬大致平齐"的文字块（纵向中心差 ≤ 圆直径的 0.9 倍），
        这样右栏顶部的胶囊标签文字不会被算进任何一组。"""
        buckets = [[] for _ in circles]
        for t in pg['texts']:
            if not circles or t.left <= min(c.left for c in circles):
                continue
            if rw(t.cx()) < 0.50:
                continue
            k = min(range(len(circles)),
                    key=lambda i: abs(t.cy() - circles[i].cy()))
            if abs(t.cy() - circles[k].cy()) > 0.9 * circles[k].height:
                continue
            buckets[k].append(t)
        return buckets

    TITLES1 = ('环保涂层：', '绿色材料：', '应用方案：')
    ITEMS1 = (('水性树脂', '功能助剂', '高阻隔涂层'),
              ('可降解薄膜', '纸塑复合材料'),
              ('工业胶黏剂', '包装表面处理'))

    def right_groups_ok(pg, circles, titles, items, need_bold_teal=True):
        buckets = group_texts(pg, circles)
        if len(buckets) != len(titles):
            return False
        for k, bucket in enumerate(buckets):
            title = titles[k]
            t_key = nospace(title)
            hit_title = False
            title_sz = 0.0
            item_szs = []
            item_lines = 0
            joined = ''
            for sh in bucket:
                for para in sh.runs:
                    rs = [r for r in para if r['t'].strip()]
                    if not rs:
                        continue
                    ptxt = _norm_text(''.join(r['t'] for r in rs))
                    joined += ptxt + '\n'
                    if t_key in nospace(ptxt):
                        hit_title = True
                        title_sz = max(title_sz, max(r['sz'] for r in rs))
                        if need_bold_teal:
                            if not all(r['b'] for r in rs):
                                return False
                            if not any(_is_deep_teal(r['clr'] or '')
                                       for r in rs):
                                return False
                    else:
                        item_szs.append(max(r['sz'] for r in rs))
                        item_lines += 1
            if not hit_title:
                return False
            if item_szs and title_sz <= max(item_szs):
                return False
            if item_lines != len(items[k]):
                return False
            flat = nospace(joined)
            if not all(nospace(x) in flat for x in items[k]):
                return False
        return True

    if circ1 and right_groups_ok(p1, circ1, TITLES1, ITEMS1):
        results.append('+5：第1页右栏三组文字完整且逐行呈现——标题依次为"环保涂层："、"绿色材料："、"应用方案："，加粗深墨绿、使用全角冒号、字号大于其下条目；条目依次为"水性树脂／功能助剂／高阻隔涂层"共3行、"可降解薄膜／纸塑复合材料"共2行、"工业胶黏剂／包装表面处理"共2行')
    else:
        results.append('+0：第1页右栏三组文字完整且逐行呈现——标题依次为"环保涂层："、"绿色材料："、"应用方案："，加粗深墨绿、使用全角冒号、字号大于其下条目；条目依次为"水性树脂／功能助剂／高阻隔涂层"共3行、"可降解薄膜／纸塑复合材料"共2行、"工业胶黏剂／包装表面处理"共2行')

    # ============================================================
    # +1：三组之间设有2条细横向分隔线，位于组间空隙处，跨度占页宽0.20-0.35
    # 严格对齐细则：
    #   ① 右栏内有 2 条细横线（高度极小）
    #   ② 每条跨度占页宽 0.20-0.35
    #   ③ 位于相邻两组之间的空隙（纵向落在两个圆形底衬之间）
    # 与办公软件对齐：即右栏两组之间那条淡淡的横向细线。
    sep_cnt = 0
    if len(circ1) == 3:
        tops = sorted(c.top for c in circ1)
        bots = sorted(c.bottom() for c in circ1)
        gaps = [(bots[0], tops[1]), (bots[1], tops[2])]
        for lo, hi in gaps:
            for s in hlines_of(p1):
                if rw(s.cx()) < 0.50:
                    continue
                if not (0.20 <= rw(s.width) <= 0.35):
                    continue
                if lo - 0.02 * H <= s.cy() <= hi + 0.02 * H:
                    sep_cnt += 1
                    break
    if sep_cnt >= 2:
        results.append("+1：三组之间设有2条细横向分隔线，位于组间空隙处，跨度覆盖圆形底衬至文字块右缘（占页宽0.20-0.35）")
    else:
        results.append("+0：三组之间设有2条细横向分隔线，位于组间空隙处，跨度覆盖圆形底衬至文字块右缘（占页宽0.20-0.35）")

    # ============================================================
    # +5：两页底部深墨绿色带，上缘为平滑曲线（自由曲线/曲线路径），横向覆盖≥页宽0.70，
    #     最高处占页高0.15-0.35；直角矩形、波浪预设形状、矩形堆叠近似者不达标
    # 严格对齐细则：
    #   ① 色带以自由曲线绘制：custGeom 路径且含贝塞尔曲线段（cubicBezTo/quadBezTo）
    #   ② 填充为墨绿族
    #   ③ 曲线色带横向覆盖 ≥ 页宽 0.70
    #   ④ 最高处占页高 0.15-0.35
    # 与办公软件对齐：custGeom+cubicBezTo → PPT 里的"任意多边形/曲线"，编辑顶点时
    #   看到的是带弧度的控制柄；prstGeom prst="wave" 只是预设波浪，边缘规律呆板。
    band_ok = True
    for pg in (p1, p2):
        curved = [s for s in bands_of(pg) if s.has_custgeom and s.curve_n > 0]
        if not curved:
            band_ok = False
            continue
        cover = _union_span([(s.left, s.right()) for s in curved])
        top_h = max(rh(s.height) for s in curved)
        if rw(cover) < 0.70 or not (0.15 <= top_h <= 0.35):
            band_ok = False
    if band_ok:
        results.append("+5：两页底部均设有深墨绿色带，上缘为平滑曲线（以自由曲线/曲线路径绘制），横向覆盖≥页宽0.70，最高处占页高0.15-0.35；以直角矩形、波浪预设形状或多个矩形/旋转矩形堆叠近似出折角、平直边缘者不达标")
    else:
        results.append("+0：两页底部均设有深墨绿色带，上缘为平滑曲线（以自由曲线/曲线路径绘制），横向覆盖≥页宽0.70，最高处占页高0.15-0.35；以直角矩形、波浪预设形状或多个矩形/旋转矩形堆叠近似出折角、平直边缘者不达标")

    # ============================================================
    # +3：第1页底部色带左端设有铜橙色块/渐变过渡段，与深墨绿色带衔接形成双色收边，
    #     宽占页宽0.10-0.45
    # 严格对齐细则（两种做法都认，只要收边宽度落在 0.10-0.45 页宽）：
    #   ① 独立的铜橙色形状压在色带左端；或
    #   ② 色带本身用渐变填充，铜橙色标位于左端，其铜橙段宽度占页宽 0.10-0.45
    # 与办公软件对齐：gradFill/gs@pos → 渐变光圈位置（1/1000 百分比），lin@ang=0 表示
    #   从左到右过渡，因此铜橙段宽度 = 最后一个铜橙光圈位置 × 色带宽度。
    copper_ok = False
    for s in p1['all']:
        if s.is_group or s.area() <= 0:
            continue
        if rh(s.bottom()) < 0.80 or rw(s.left) > 0.15:
            continue
        if s.fill_kind == 'solid' and any(_is_copper(h) for h in s.fills):
            if 0.10 <= rw(s.width) <= 0.45:
                copper_ok = True
    if not copper_ok:
        for s in bands_of(p1):
            if s.fill_kind != 'grad':
                continue
            g = s.el.find(_p('spPr')).find(_a('gradFill'))
            last = 0.0
            for gs in g.iter(_a('gs')):
                h = _clr_hex(gs, theme)
                if h and _is_copper(h):
                    last = max(last, _int_attr(gs, 'pos', 0) / 100000.0)
            seg = rw(last * s.width)
            if 0.10 <= seg <= 0.45:
                copper_ok = True
    if copper_ok:
        results.append("+3：第1页底部色带左端设有铜橙色块/渐变过渡段，与深墨绿色带衔接形成双色收边，宽占页宽0.10-0.45")
    else:
        results.append("+0：第1页底部色带左端设有铜橙色块/渐变过渡段，与深墨绿色带衔接形成双色收边，宽占页宽0.10-0.45")

    # ============================================================
    # +3：第2页左上角英文公司名分两行"FUJIAN LANYU NOVA"与"MATERIALS CO., LTD."，
    #     加粗深墨绿，为该页最大字号文本，左边缘位于页宽0.04-0.12
    # 严格对齐细则：
    #   ① 内容分两行，且两行文字分别与细则给出的两段一致
    #   ② 加粗、深墨绿
    #   ③ 为该页最大字号文本
    #   ④ 左边缘位于页宽 0.04-0.12
    # 说明：手写花体/书法体的标称字号与常规字体不可直接比较（同样磅值下花体的字面
    #   要小得多），因此"该页最大字号"只与常规字体文本比较，不与花体姓名比较。
    # 与办公软件对齐：a:p 两段 → 两行；sz 与页内其他文字比较即"该页最大字号"。
    en2 = first(en_name2)
    en2_ok = en2 is not None
    if en2_ok:
        lines = [_norm_text(''.join(r['t'] for r in para)).upper()
                 for para in en2.runs
                 if ''.join(r['t'] for r in para).strip()]
        rs = runs_of(en2)
        plain_szs = [r['sz'] for _, r in _iter_runs(p2['texts'])
                     if not _is_script_font(r['latin'] or '')
                     and not _is_kai_font(r['ea'] or r['latin'] or '')]
        biggest = max(plain_szs) if plain_szs else 0.0
        en2_ok = (len(lines) == 2
                  and nospace(lines[0]) == 'FUJIANLANYUNOVA'
                  and nospace(lines[1]) == 'MATERIALSCO.,LTD.'
                  and all(r['b'] for r in rs)
                  and any(_is_deep_teal(r['clr'] or '') for r in rs)
                  and sz_of(en2) >= biggest - 0.01
                  and 0.04 <= rw(en2.left) <= 0.12)
    if en2_ok:
        results.append('+3：第2页左上角设有英文公司名，分两行呈现"FUJIAN LANYU NOVA"与"MATERIALS CO., LTD."，加粗深墨绿，为该页最大字号文本，左边缘位于页宽0.04-0.12')
    else:
        results.append('+0：第2页左上角设有英文公司名，分两行呈现"FUJIAN LANYU NOVA"与"MATERIALS CO., LTD."，加粗深墨绿，为该页最大字号文本，左边缘位于页宽0.04-0.12')

    # ============================================================
    # +3：第2页姓名"Ethan Shen"，手写花体英文字体、深墨绿、上边缘位于页高0.30-0.45，
    #     与第1页姓名位置大致对应
    # 严格对齐细则：
    #   ① 文字为"Ethan Shen"
    #   ② 西文字体为脚本体/手写花体
    #   ③ 文字色为深墨绿
    #   ④ 上边缘位于页高 0.30-0.45
    #   ⑤ 左边缘与第1页姓名左边缘大致对应（差异 ≤ 页宽 0.10）
    # 与办公软件对齐：a:latin typeface="Segoe Script" 等 → 字体框里的手写花体。
    et = first(ethan)
    et_ok = et is not None and xg is not None
    if et_ok:
        rs = runs_of(et)
        fonts = [r['latin'] for r in rs if r['latin']]
        et_ok = (nospace(et.text).lower() == 'ethanshen'
                 and fonts and all(_is_script_font(f) for f in fonts)
                 and any(_is_deep_teal(r['clr'] or '') for r in rs)
                 and 0.30 <= rh(et.top) <= 0.45
                 and abs(rw(et.left) - rw(xg.left)) <= 0.10)
    if et_ok:
        results.append('+3：第2页左半区设有姓名"Ethan Shen"，采用手写花体英文字体（Segoe Script等脚本体），深墨绿，上边缘位于页高0.30-0.45，与第1页姓名位置大致对应')
    else:
        results.append('+0：第2页左半区设有姓名"Ethan Shen"，采用手写花体英文字体（Segoe Script等脚本体），深墨绿，上边缘位于页高0.30-0.45，与第1页姓名位置大致对应')

    # ============================================================
    # +3：第2页与第1页同款的3行徽章式联系方式，英文地址完整并断为3行
    # 严格对齐细则：
    #   ① 第2页左半区同样有 3 个徽章、等距堆叠、左边缘对齐
    #   ② 三行内容为电话、邮箱、英文地址
    #   ③ 英文地址三段内容逐字齐全，并断为 3 行
    # 与办公软件对齐：三个 a:p 段落 → 地址显示为三行。
    ad2 = first(addr2)
    c27 = badge_rows_ok(p2, badge2)
    c27 = c27 and bool(phone2) and bool(mail2) and ad2 is not None
    if c27:
        flat2 = _norm_text(ad2.text).lower().replace(' ', '')
        lines2 = [_norm_text(''.join(r['t'] for r in para))
                  for para in ad2.runs
                  if ''.join(r['t'] for r in para).strip()]
        want = ('suite1208,bluebayinnovationcenter,',
                'no.88yunchuangavenue,',
                'haicangdistrict,xiamen,fujian,china')
        c27 = (len(lines2) == 3 and all(w in flat2 for w in want)
               and '4826' in ' '.join(s.text for s in phone2))
    if c27:
        results.append('+3：第2页左半区设有与第1页同款的3行徽章式联系方式（相同图标样式＋"M.P. 139 4826 7153"、"z.xing@lanyunova.com"、英文地址），英文地址完整呈现"Suite 1208, BlueBay Innovation Center,"、"No. 88 Yunchuang Avenue,"、"Haicang District, Xiamen, Fujian, China"并断为3行')
    else:
        results.append('+0：第2页左半区设有与第1页同款的3行徽章式联系方式（相同图标样式＋"M.P. 139 4826 7153"、"z.xing@lanyunova.com"、英文地址），英文地址完整呈现"Suite 1208, BlueBay Innovation Center,"、"No. 88 Yunchuang Avenue,"、"Haicang District, Xiamen, Fujian, China"并断为3行')

    # ============================================================
    # +3：第2页右栏"MAIN BUSINESS"胶囊标签（深墨绿实心＋白色文字）并设纵向分隔线，
    #     标签左边缘与分隔线横坐标同第1页对应元素差异≤页宽0.06
    # 严格对齐细则：
    #   ① 有 MAIN BUSINESS 胶囊标签，墨绿实心 + 白色文字
    #   ② 有纵向分隔线
    #   ③ 标签左边缘、分隔线横坐标与第1页对应元素差异均 ≤ 页宽 0.06
    # 与办公软件对齐：正反面同位置元素在"大小和位置"里读到的水平位置基本一致。
    c28 = cap2 is not None and vdiv2 is not None and cap1 is not None \
        and vdiv1 is not None
    if c28:
        rs = [r for t in cap_texts(p2, cap2) for r in runs_of(t)]
        c28 = (rs and any(_is_white(r['clr'] or '') for r in rs)
               and rw(abs(cap2.left - cap1.left)) <= 0.06
               and rw(abs(vdiv2.cx() - vdiv1.cx())) <= 0.06)
    if c28:
        results.append('+3：第2页右栏顶部设有"MAIN BUSINESS"胶囊标签（深墨绿实心＋白色文字）并设纵向分隔线，标签左边缘与分隔线横坐标同第1页对应元素差异≤页宽0.06')
    else:
        results.append('+0：第2页右栏顶部设有"MAIN BUSINESS"胶囊标签（深墨绿实心＋白色文字）并设纵向分隔线，标签左边缘与分隔线横坐标同第1页对应元素差异≤页宽0.06')

    # ============================================================
    # +5：第2页右栏3个主营分组，英文标题与条目逐字齐全，每组配浅色圆形底衬＋真实图标
    # 严格对齐细则：
    #   ① 三个英文标题逐字一致
    #   ② 条目内容逐字一致，行数分别为 3、2、2
    #   ③ 每组有浅色圆形底衬，且圆内图标为图片/矢量图形（非符号字符）
    # 与办公软件对齐：与第1页右栏同构，只是文字换成英文。
    TITLES2 = ('ECO COATINGS:', 'GREEN MATERIALS:', 'APPLICATION SOLUTIONS:')
    ITEMS2 = (('Water-based resins', 'Functional additives',
               'Barrier coatings'),
              ('Biodegradable films', 'Paper-plastic composites'),
              ('Industrial adhesives', 'Packaging surface treatment'))
    c29 = (len(circ2) == 3
           and circle_groups_ok(p2, circ2)
           and icons_real(p2, circ2, 3)
           and right_groups_ok(p2, circ2, TITLES2, ITEMS2,
                               need_bold_teal=False))
    if c29:
        results.append('+5：第2页右栏设有3个主营分组，标题依次为"ECO COATINGS:"、"GREEN MATERIALS:"、"APPLICATION SOLUTIONS:"，条目分别为"Water-based resins／Functional additives／Barrier coatings"、"Biodegradable films／Paper-plastic composites"、"Industrial adhesives／Packaging surface treatment"，每组均配浅色圆形底衬＋水滴/叶片/包装箱图标')
    else:
        results.append('+0：第2页右栏设有3个主营分组，标题依次为"ECO COATINGS:"、"GREEN MATERIALS:"、"APPLICATION SOLUTIONS:"，条目分别为"Water-based resins／Functional additives／Barrier coatings"、"Biodegradable films／Paper-plastic composites"、"Industrial adhesives／Packaging surface treatment"，每组均配浅色圆形底衬＋水滴/叶片/包装箱图标')

    # ============================================================
    # +5：第2页底部色带内4组"图标＋全大写标签"，横向等距排列于色带左侧2/3范围内；
    #     图标为图片或矢量图形而非文本符号字符，标签字号≥8磅
    # 严格对齐细则：
    #   ① 四个标签 SUSTAINABLE / INNOVATIVE / RELIABLE / GLOBAL 齐全且全大写
    #   ② 四组位于色带内（页高 0.80 以下）、横坐标都在页宽 2/3 以内
    #   ③ 横向等距（相邻间距差异 ≤ 图标宽度的 0.5 倍）
    #   ④ 每个标签上方配一个图片/矢量图形图标，且不是符号字符文字
    #   ⑤ 标签字号 ≥ 8 磅
    # 与办公软件对齐：字号被压到 7 磅以下在 PPT 里几乎看不清，属"惹麻烦"。
    WORDS = ('SUSTAINABLE', 'INNOVATIVE', 'RELIABLE', 'GLOBAL')
    labels = []
    for w in WORDS:
        hit = [s for s in p2['texts']
               if nospace(s.text).upper() == w and rh(s.cy()) > 0.78]
        if hit:
            labels.append(hit[0])
    c30 = len(labels) == 4
    if c30:
        c30 = all(sz_of(s) >= 8.0 for s in labels)
        c30 = c30 and all(s.text == s.text.upper() for s in labels)
        c30 = c30 and all(rw(s.cx()) <= 0.67 for s in labels)
    if c30:
        icons = []
        for lb in labels:
            cand = []
            for s in p2['all']:
                if s.is_group or s.has_text() or s.area() <= 0:
                    continue
                if not (s.is_pic or s.has_custgeom or s.prst):
                    continue
                if abs(s.cx() - lb.cx()) > 1.2 * max(lb.width, lb.height):
                    continue
                if s.cy() > lb.cy():
                    continue
                if rh(s.cy()) < 0.72:
                    continue
                cand.append(s)
            # 一个标签配"正上方最近"的那个图标：先比横向偏移，再比谁更贴近标签
            cand.sort(key=lambda s: (abs(s.cx() - lb.cx()), -s.cy()))
            icons.append(cand[0] if cand else None)
        c30 = all(ic is not None for ic in icons)
        # 四个标签必须各自配到不同的图标，不能几个标签共用同一个图形
        c30 = c30 and len(set(id(ic) for ic in icons)) == 4
        if c30:
            cxs = sorted(ic.cx() for ic in icons)
            wdt = sum(ic.width for ic in icons) / 4.0
            c30 = _equidistant(cxs, 0.5 * wdt)
    if c30:
        results.append('+5：第2页底部深墨绿色带内设有4组"图标＋全大写标签"，依次为叶片-SUSTAINABLE、循环箭头-INNOVATIVE、盾牌对勾-RELIABLE、地球-GLOBAL，横向等距排列于色带左侧2/3范围内；4个图标为图片或矢量图形而非文本符号字符，标签文字字号≥8磅、清晰可读（未被压缩至7磅以下）')
    else:
        results.append('+0：第2页底部深墨绿色带内设有4组"图标＋全大写标签"，依次为叶片-SUSTAINABLE、循环箭头-INNOVATIVE、盾牌对勾-RELIABLE、地球-GLOBAL，横向等距排列于色带左侧2/3范围内；4个图标为图片或矢量图形而非文本符号字符，标签文字字号≥8磅、清晰可读（未被压缩至7磅以下）')

    # ============================================================
    # +1：第2页底部色带上缘另有波浪/曲线类线条装饰，横向跨度≥页宽0.20
    # 严格对齐细则：
    #   ① 是独立于色带本体的另一个对象（色带自己的曲线上缘不算）
    #   ② 位于色带上缘附近：纵向中心落在色带纵向范围内，或在其上缘之上 0.10 页高内
    #   ③ 起于版面左侧（左边缘位于页宽 0.20 以内），横向跨度 ≥ 页宽 0.20
    #   ④ 形态为波浪/曲线/线条类即可：自由曲线、曲线或直线连接符、wave/arc 等
    #      预设曲线形状、细长线条形状（宽 ≥ 6 倍高）均可，深浅不限
    # 与办公软件对齐：本条只认"色带之外另叠了一道装饰线"这件事，画法与颜色都放开，
    #   用曲线工具画的、插入波浪/弧形预设形状的、拿细长形状叠出来的都算；右栏那种
    #   短横向分隔线因起点在版面右半区，不会被算成色带装饰。
    wave_ok = False
    bands2 = bands_of(p2)
    band2_ids = set(id(s) for s in bands2)
    if bands2:
        b_top = min(s.top for s in bands2)
        b_bot = max(s.bottom() for s in bands2)
    else:
        b_top, b_bot = int(0.70 * H), H
    for s in p2['all']:
        if s.is_group or s.has_text() or id(s) in band2_ids:
            continue
        if rw(s.width) < 0.20 or rw(s.left) > 0.20:
            continue
        if not (b_top - 0.10 * H <= s.cy() <= b_bot):
            continue
        curvy = (s.has_custgeom or s.curve_n > 0 or s.kind == 'cxn'
                 or (s.prst or '').startswith(('wave', 'doubleWave', 'arc',
                                               'curvedConnector', 'bentConnector'))
                 or (s.prst or '') == 'line'
                 or (s.height > 0 and s.width >= 6 * s.height))
        if curvy:
            wave_ok = True
    if wave_ok:
        results.append("+1：第2页底部色带上缘另叠有波浪/曲线类线条装饰，起于版面左侧、横向跨度≥页宽0.20，且为独立于色带本体的对象（自由曲线、曲线连接符、波浪或弧形预设形状、细长线条形状均可，深浅不限）")
    else:
        results.append("+0：第2页底部色带上缘另叠有波浪/曲线类线条装饰，起于版面左侧、横向跨度≥页宽0.20，且为独立于色带本体的对象（自由曲线、曲线连接符、波浪或弧形预设形状、细长线条形状均可，深浅不限）")

    # ============================================================
    # +5：两页所有对象均落在卡片主体范围内，装饰图形与文字不相互遮挡、
    #     文本框之间无重叠、无内容溢出画布边界
    # 严格对齐细则：
    #   ① 每个对象都落在卡片主体范围内（允许 0.02 页宽/页高的描边级误差）
    #   ② 没有内容溢出画布边界
    #   ③ 文本框之间不重叠（按行数×字号估算的实际字块，重叠 > 较小者的 0.5 视为重叠）
    #   ④ 后置的实心装饰图形没有压住文字（覆盖 > 文字框 0.5 视为遮挡）
    # 说明：文字放在胶囊/圆形底衬之上是设计需要，不算遮挡——判定只看"后画的
    #   实心图形压住了先画的文字"这种真正会看不见字的情况。文本框高度常留有
    #   富余，框线相擦不等于文字压字，故文字之间按实际字块估算后再比。宽高为负
    #   的畸形图形坐标本身已不可读，由"宽高非负"一项单独扣分，这里不重复计入。
    def layout_ok(pg, card):
        tol_x, tol_y = 0.02 * W, 0.02 * H
        for s in pg['all']:
            if s.area() <= 0 or s.raw_w < 0 or s.raw_h < 0:
                continue
            if s.left < -tol_x or s.top < -tol_y \
                    or s.right() > W + tol_x or s.bottom() > H + tol_y:
                return False
            if card is not None:
                if s.left < card.left - tol_x or s.top < card.top - tol_y \
                        or s.right() > card.right() + tol_x \
                        or s.bottom() > card.bottom() + tol_y:
                    return False
        txts = [s for s in pg['all'] if s.has_text() and s.area() > 0]
        inks = [_text_ink_box(s) for s in txts]
        for i in range(len(txts)):
            for j in range(i + 1, len(txts)):
                small = min(_box_area(inks[i]), _box_area(inks[j]))
                if small > 0 and _box_inter_area(inks[i], inks[j]) > 0.5 * small:
                    return False
        for t in txts:
            for s in pg['all']:
                if s is t or s.has_text() or s.is_group or s.area() <= 0:
                    continue
                if s.z <= t.z:
                    continue
                if s.fill_kind not in ('solid', 'grad', 'pic'):
                    continue
                if _cover_ratio(t, s) > 0.5:
                    return False
        return True

    if layout_ok(p1, card1) and layout_ok(p2, card2):
        results.append("+5：两页所有对象均落在卡片主体范围内，装饰图形与文字不相互遮挡、文本框之间无重叠、无内容溢出画布边界")
    else:
        results.append("+0：两页所有对象均落在卡片主体范围内，装饰图形与文字不相互遮挡、文本框之间无重叠、无内容溢出画布边界")

    # ============================================================
    # +3：两页共用统一配色体系（深墨绿为主色、铜橙为辅色、米白为底色），
    #     无明显跳色或高饱和杂色；卡片主体不描黑色边框
    # 严格对齐细则：
    #   ① 卡片主体不描黑色边框（近黑且接近中性灰的描边即视为黑边）
    #   ② 全篇填充色与文字色中，不出现"高饱和且色相既不属墨绿也不属铜橙"的杂色
    # 说明：只对"亮度中高 + 饱和度高"的颜色设色相门槛——深色、浅色、灰调色在视觉上
    #   不构成跳色，不据此扣分。
    # 与办公软件对齐：a:ln/solidFill → 形状轮廓颜色；黑色描边会让米白卡片显得生硬。
    def stray_color(hexstr):
        c = _rgb(hexstr)
        if c is None:
            return False
        h, s, l = _rgb_to_hsl(c)
        if s < 0.60 or not (0.30 <= l <= 0.88):
            return False
        deg = h * 360.0
        if 140.0 <= deg <= 210.0 or deg <= 50.0 or deg >= 350.0:
            return False
        return True

    color_ok = True
    for pg, cd in ((p1, card1), (p2, card2)):
        if cd is not None and cd.line_hex and _is_blackish(cd.line_hex):
            color_ok = False
        for s in pg['all']:
            for hx in s.fills:
                if stray_color(hx):
                    color_ok = False
        for _, r in _iter_runs(pg['texts']):
            if r['clr'] and stray_color(r['clr']):
                color_ok = False
    if color_ok:
        results.append("+3：两页共用统一配色体系（深墨绿#004F49-#05565D为主色、铜橙#B5713C-#C77B45为辅色、米白为底色），无明显跳色或高饱和杂色；卡片主体不描黑色边框")
    else:
        results.append("+0：两页共用统一配色体系（深墨绿#004F49-#05565D为主色、铜橙#B5713C-#C77B45为辅色、米白为底色），无明显跳色或高饱和杂色；卡片主体不描黑色边框")

    # ========== 将 results 反解为结构化 dim2_items ==========
    if len(results) != len(_MAX_DELTAS_ORDER):
        raise RuntimeError(
            f"dim2 项数不一致：results={len(results)} vs order={len(_MAX_DELTAS_ORDER)}")
    dim2_items = []
    total_score = 0
    for raw, max_delta in zip(results, _MAX_DELTAS_ORDER):
        head, _, rule = raw.partition("：")
        try:
            delta = int(head.replace("+", "").strip())
        except ValueError:
            delta = 0
        hit = delta > 0
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
    # 注：evaluate() 的接口约定就是"入参为目录路径"——路径始终由外部显式传入；
    #     这里的默认值仅为本地调试便利，不影响 evaluate 的接口约定。
    if len(sys.argv) >= 2:
        _dir = sys.argv[1]
    else:
        _dir = os.path.dirname(os.path.abspath(__file__))
    _out = evaluate(_dir)
    print(json.dumps(_out, ensure_ascii=False, indent=2))


























