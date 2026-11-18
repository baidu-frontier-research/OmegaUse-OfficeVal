#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# pyright: reportUnknownMemberType=false, reportUnknownVariableType=false, reportUnknownArgumentType=false, reportUnknownParameterType=false, reportMissingParameterType=false, reportPrivateUsage=false, reportUnknownLambdaType=false, reportGeneralTypeIssues=false

import os
import re
import io
import sys
import json
import math
import hashlib
import traceback
from pptx import Presentation
from pptx.util import Emu
try:
    from PIL import Image  # Pillow：用于按"图片内容"识别被迁移的两张配图与 5 张老师卡片
except ImportError:
    Image = None

# 说明：按接口统一约定，本模块不修改全局 sys.stdout，也不 print 主结果。
# 主结果统一由 evaluate(dir_path) -> dict 返回；入参为"脚本所在目录的路径"，
# 由脚本自己在该目录内定位并打开被评估的 .pptx。本文件仅在 __main__ 下作为
# 本地调试入口，才把 dict 序列化为 JSON 打印。

SCRIPT_ID = "061"
# 维度二所有加分项 max_delta 之和；用于维度一未通过 / 脚本异常时兜底
MAX_SCORE_TOTAL = 72

# 维度二所有评分项按"结果追加顺序"的 max_delta 列表；末端把 results 里的
# "+N：..." / "-N：..." 字符串反解为结构化 dim2_items 时用。顺序须与
# _evaluate_impl 中 results.append(...) 的顺序严格一致；
# 共 33 项 = 20 个加分项（累加 = MAX_SCORE_TOTAL）+ 13 个减分项（累加 = -23）。
_MAX_DELTAS_ORDER = [
    5,   # 01 首页右侧两张圆角矩形图片：内容/等尺寸/并排
    3,   # 02 两图宽高比 1.55-1.65、单张 4.0-6.6 × 2.5-4.1cm
    3,   # 03 两图为独立图片对象、四段文字为 4 个独立文本框
    5,   # 04 四段文字的分组、上下排列与相对图片的位置
    3,   # 05 四段文字字号 20 磅加粗 / 14 磅且单行完整显示
    3,   # 06 整体区域与参考示意图红框位置一致
    1,   # 07 整体区域完全在右半侧且与首页四处文字互不压盖
    3,   # 08 全篇不出现标题为"品牌宣言"的页面
    5,   # 09 5 张老师卡片为 5 个独立图片对象（非整幅位图/GIF）
    5,   # 10 5 张橙白配色卡片且人像互不相同
    3,   # 11 全篇不出现"学习支持团队介绍"页及其文本框与说明段
    5,   # 12 连续 5 张"教学团队介绍"页且非卡片元素逐页一致
    5,   # 13 5 张卡片"居中最大、向两侧递减"的三级高度
    3,   # 14 5 张卡片宽高比 0.54-0.58
    3,   # 15 5 张卡片共用水平中心线、纵向 56%-62%、居中卡片 48%-53%
    3,   # 16 相邻卡片重叠且靠中者压在靠外者之上、居中卡片最上层
    5,   # 17 5 页卡片排列逐页循环左移一位
    5,   # 18 第 2-5 页为平滑/变体（Morph）切换且时长 1.5-2.5 秒
    3,   # 19 5 页均自动换片、换片时间 0 秒
    1,   # 20 5 页位于"团队介绍"之后、"品牌文化"之前
    -3,  # 21 页面尺寸不为 33.87 × 19.05cm
    -3,  # 22 首页署名文字不为"星市静默回声教育"
    -3,  # 23 "01.品牌"栏首行不为"星市静默回声教育科技有限公司"
    -1,  # 24 "01.品牌"栏第二行不符
    -1,  # 25 "02.概述"栏正文不符
    -1,  # 26 卡片页缺白色圆角底板 / 菱形箭头 / 三角装饰
    -1,  # 27 首页四处文字未完整出现
    -1,  # 28 首页左侧装饰图形缺失
    -3,  # 29 任一页面出现红色参考标注框线
    -3,  # 30 未同时出现"01""品牌概述"(2页)"协作架构""团队介绍"页
    -1,  # 31 "品牌文化/价值观/使命/愿景"四页缺失或顺序不符
    -1,  # 32 出现水印 / 注释性文字 / 日期文本框
    -1,  # 33 文字超出文本框可视区域被截断
]
assert sum(d for d in _MAX_DELTAS_ORDER if d > 0) == MAX_SCORE_TOTAL
assert len(_MAX_DELTAS_ORDER) == 33

# ---------------- 全局常量 ----------------
EMU_PER_CM = 360000.0
PT_PER_CM = 1.0 / 0.0352778          # 1 磅 = 0.0352778cm
A_NS = '{http://schemas.openxmlformats.org/drawingml/2006/main}'
R_NS = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'
P14_NS = '{http://schemas.microsoft.com/office/powerpoint/2010/main}'
MSO_GROUP = 6                        # MSO_SHAPE_TYPE.GROUP
MSO_PICTURE = 13                     # MSO_SHAPE_TYPE.PICTURE

# 页面尺寸：33.87cm × 19.05cm（16:9 横版），即 12192000 × 6858000 EMU
PAGE_W_CM = 33.87
PAGE_H_CM = 19.05
PAGE_SIZE_TOL_CM = 0.05

# 首页迁入的四段文字（细则原文）
TXT_BIG_L = "持续精进"
TXT_SUB_L = "学习成长与迭代创新"
TXT_BIG_R = "实践见效"
TXT_SUB_R = "把方法落实到课堂与服务"
MIGRATED_TEXTS = (TXT_BIG_L, TXT_SUB_L, TXT_BIG_R, TXT_SUB_R)

# 首页四处仅作位置参照的文字
COVER_TEXTS = ("2026", "公司简介", "品牌介绍手册", "simpleatmospherecompanyprofile")

# 页面标题（按"页内字号最大的短文本"解析，见 _title_of）
T_TEAM_PAGE = "教学团队介绍"
T_SUPPORT_PAGE = "学习支持团队介绍"
T_DECLARE_PAGE = "品牌宣言"
T_TEAM_SECTION = "团队介绍"
T_CULTURE = "品牌文化"

# 图片内容指纹：8×8 RGB 缩略图（LANCZOS）的原始字节，取自输入原稿中的对应图片。
#   · LEFT_BOOK   = 原"品牌宣言"页左图：翻开的书本与发光灯泡
#   · RIGHT_CLASS = 原"品牌宣言"页右图：教室内学生举手的课堂场景
#   · CARD_1..5   = 原第 7、8 页的 5 张橙白配色老师简历卡片（5 位老师人像各不相同）
# 判定口径：两图 8×8 RGB 的平均绝对差 ≤ SIG_MATCH_MAE 视为"同一张图片内容"。
#   实测同图不同编码（重压缩 / 缩放到 804×503）差值 ≤ 0.64；不同照片差值 ≥ 2.66，
#   5 张卡片相互差值 ≥ 28，故阈值取 1.5 有足够裕量。
SIG_MATCH_MAE = 1.5

_REF_SIG_HEX = {
    "LEFT_BOOK": (
        "221b0f1310091f13072b1a0a2a1a0a241607563e20514430352a0635250531170471481b6b441c40"
        "21084b2e10432d131f1702482f09582e0ca76c2b9a6328633a1b42230d2311012d1d0e55331867360d"
        "bb6f18b16b1c532d0c4221093b1d0643240b7a4b1e996c3ffed586f5c978a77f55533013240e004120"
        "065b34133b1d0d96734d8d64375c381f764c27532e12391d052913012a1201441d003b1900230d0051"
        "2e108b5726230f01361a044622065f330e58300d4a260c673c16502f12"
    ),
    "RIGHT_CLASS": (
        "cbbdafbeb09eac987ca9987eb7a38cb9a48dad967b9a846beee8e5e6e6e0b9b195a99477beb6acb5ab"
        "a2a79f999e9894e6e1e0d3d2be93946a978671c7c1bb88808099969aaba9ac7b6b5d8c79676a5e5666"
        "4e417c705c80858b867b7a7865527c66556854495047467c6b671d1811767169766255635947947f6e"
        "5d46353224187c7970283e4078706897796157524d3d3b3d463b36291c123e444239595b898b907369"
        "6617141131313429292d32302f282729242829393b3b3f3e40191a19"
    ),
    "CARD_1": (
        "fcb16ffde5ceffe9d4d7b9a1d6beabfffffffff8f1fbf5f0fcd0aafdffffbdc4c489716297715e8696"
        "94eeefecfefaf7fbe2cbfdf5e88f918d6b696b9988885f6e71f3f4f0fcf7f2fbe2cafee8d3facca5e9"
        "bc96efc5a0ffcea4fff3e8fbf9f6fcf7f2fcead9f8f9f9fbfaf8faf8f6fafaf9fdfdfdfcf8f4fcf4ed"
        "fde1c7f9f5f1f6ede6f8f3effaf9f9fbfcfcfcf9f6fcf0e5fee8d4fcf8f3f6f5f5f6f7f8f5f6f8f9f7"
        "f6fcf3ebfbcba0fcb475fcc798facea8facea8fbc698fcb06dfcca9e"
    ),
    "CARD_2": (
        "f9b487f9893ddfc0a4877c728a7464beb4aacdc7bfe3e3e1feae6eb97b469389799f7659bf8e71c4bc"
        "b4cfcdc9ebf1f2ffa967bf8350977658ac7e57bb906bbc926cb2987fccbfb2facdaefcc59ae9d0bcdd"
        "b696ddb18ed3ac8cc5ac98cac9c7fbfefffee7d5fef4edfcf3ecfcf5f0fdfefefffffffffffffbf8f5"
        "fdf4edf3f7f9f0f2f4f4f6f8f8f9fafcfaf8f9f8f7fbf8f6fdf0e5f8f4f0f6f0ecf9f4f0fbf7f4fef7"
        "f1fbf9f8faf7f4faeee4faefe6fbefe5faeee4faf0e7faece1faf8f7"
    ),
    "CARD_3": (
        "fd9330ff9d41fadfc2b4a29579604dd7c7b9cec7b8d6c7b5f3c293eed6bef3f6f5c9a790ad724ea39d"
        "9044533d877855dcc2abac937ab5663ddea583dfa8839d5e44c6a07affe6c5f8eadde9d0bdd49a79ee"
        "c7a7f3caa7f2d4c0fffffffbf3edfef6effcd8b8fff0defcf5edf9e8dafaf0e6faf7f3fcf3ebfbeddf"
        "fbf3ecf7e8ddfcfffffcf8f3f9ece0f9efe6fcf5effcefe2fcf5eff9eadcfae9dbfbecddfbf5f0fbf5"
        "effcefe3fbe7d4fce9d7fadbbef7caa4f9e1cafce1c9fcc898fbaa61"
    ),
    "CARD_4": (
        "f2b989e5d6cddcc7b3b79073a6856bdcb99ed7b9a0ee9549e8b588d0e0e3aa9f927a503590664aaaa3"
        "9ecbd0d7de9d69e78f45b3b1a2bfab9a947b67ceb3a1ab907b8f8576cf884bf9a763f7e1cdf8dabff1"
        "cdafe9d7c8e4d4c8ccc1b4dea570fbb67efbfaf7f7eee7f9f4effbf9f8fffffffffffffacfa7fab179"
        "fdf6eef8f5f3f0f2f3f5f5f4f9f7f6fbfdfcf6bc8dfab780fdfaf7faf7f3f7f5f4f9f8f7fbf9f7fdff"
        "fff6bc8dfab57af5dbc5f6d7bef7d7bdf6d8bff7dac2f5ddc9f6b37c"
    ),
    "CARD_5": (
        "fad1b5f9751fecbe9eb29789ad9284e8e4e0dad2ccf1edeafbc4a0efa26ed4e2eaa78372805744b0b1"
        "b3d5d0cdece9e6fbcaa8ceaa923f4650372f2e373131211d226a6560e9e3ddfcd8bee4aa84a49b97dd"
        "9a6ece9775aaa2a0a49a93e9dfd8fcfdfdfdfffffffdf8fbf2ebfdf3ecfffefafffffff8f5f3fcfaf8"
        "fbdcc7f4e1d4f5ece5f5f8f9f7f8f9fbfbfbf7f4f2fcfcfbfceadcf9ece1f9eee5f9f1eafbf4edfcf8"
        "f4f7f5f4fbe6d7f7bf9af7d0b5f8d7c1f7d5bef7cbadf5ba92f6e6db"
    ),
}
_REF_SIG = {k: bytes.fromhex(v) for k, v in _REF_SIG_HEX.items()}
_CARD_KEYS = ("CARD_1", "CARD_2", "CARD_3", "CARD_4", "CARD_5")


# ---------------- 通用辅助函数 ----------------
def _cm(emu):
    """EMU → 厘米。"""
    return (0.0 if emu is None else float(emu)) / EMU_PER_CM


def _norm(text):
    """去掉所有空白（含全角空格/制表符/换行），用于文字内容比对。"""
    return re.sub(r'\s+', '', text or '')


def _norm_q(text):
    """在 _norm 基础上把中英文引号统一，避免"直/弯引号"差异造成误判。"""
    t = _norm(text)
    for ch in ('“', '”', '‘', '’', '「', '」'):
        t = t.replace(ch, '"')
    return t


def _open_presentation(file_path):
    """打开 .pptx；损坏 / 需修复 / 非 pptx 会抛异常，由维度一门禁捕获。"""
    return Presentation(file_path)


def _group_xfrm(grp):
    """取组合的 (off, ext, chOff, chExt)，缺任一项返回 None。"""
    gpr = grp._element.find(
        '{http://schemas.openxmlformats.org/presentationml/2006/main}grpSpPr')
    if gpr is None:
        return None
    xfrm = gpr.find(f'{A_NS}xfrm')
    if xfrm is None:
        return None
    vals = []
    for tag, ax, ay in ((f'{A_NS}off', 'x', 'y'), (f'{A_NS}ext', 'cx', 'cy'),
                        (f'{A_NS}chOff', 'x', 'y'), (f'{A_NS}chExt', 'cx', 'cy')):
        el = xfrm.find(tag)
        if el is None or el.get(ax) is None or el.get(ay) is None:
            return None
        vals.append((float(el.get(ax)), float(el.get(ay))))
    return vals


def _walk_shapes(shapes, tf=(0.0, 1.0, 0.0, 1.0), depth=0):
    """深度优先遍历（含组合内的子形状），产出 (shape, rect_cm, depth)。
    rect_cm = (x, y, w, h)，单位厘米，且已按组合的 off/ext/chOff/chExt 换算成
    幻灯片绝对坐标 —— 与在 PowerPoint 里看到的位置一致（python-pptx 对组合内
    子形状返回的是子坐标系原始值，必须做这层换算）。"""
    ax, sx, ay, sy = tf
    for sh in shapes:
        rect = None
        if sh.left is not None and sh.top is not None and sh.width is not None:
            rect = (_cm(ax + sx * sh.left), _cm(ay + sy * sh.top),
                    _cm(sx * sh.width), _cm(sy * (sh.height or 0)))
        yield sh, rect, depth
        if sh.shape_type == MSO_GROUP:
            g = _group_xfrm(sh)
            ntf = tf
            if g is not None and g[3][0] and g[3][1]:
                (ox, oy), (ex, ey), (cx0, cy0), (cex, cey) = g
                nsx, nsy = sx * ex / cex, sy * ey / cey
                ntf = (ax + sx * ox - nsx * cx0, nsx, ay + sy * oy - nsy * cy0, nsy)
            for item in _walk_shapes(sh.shapes, ntf, depth + 1):
                yield item


def _img_blob(shape, slide):
    """取形状承载的位图字节：图片对象取 image.blob；自动图形若用图片填充
    （a:blipFill/a:blip@r:embed，例如"圆角矩形+图片填充"）则顺关系取出。
    与办公软件对齐：PowerPoint 中"插入图片再裁成圆角"与"给圆角矩形填图片"
    看起来一样，两者都应算作"放了一张图"，故两种承载方式都要支持。"""
    try:
        if shape.shape_type == MSO_PICTURE:
            return shape.image.blob
    except Exception:
        pass
    try:
        blip = shape._element.find(f'.//{A_NS}blipFill/{A_NS}blip')
        if blip is None:
            return None
        rid = blip.get(f'{R_NS}embed')
        if not rid:
            return None
        return slide.part.related_part(rid).blob
    except Exception:
        return None


def _img_sig(blob):
    """图片内容指纹：8×8 RGB（LANCZOS）缩略图的原始字节（192 字节）。
    GIF/多帧图取第 1 帧。Pillow 缺失或解码失败返回 None。"""
    if Image is None or not blob:
        return None
    try:
        im = Image.open(io.BytesIO(blob))
        try:
            im.seek(0)
        except Exception:
            pass
        return im.convert('RGB').resize((8, 8), Image.LANCZOS).tobytes()
    except Exception:
        return None


def _img_frames(blob):
    """位图帧数：用于识别"整幅动图（GIF）"这类把 5 张卡片压成一个对象的做法。"""
    if Image is None or not blob:
        return 1
    try:
        im = Image.open(io.BytesIO(blob))
        return int(getattr(im, 'n_frames', 1) or 1)
    except Exception:
        return 1


def _sig_mae(a, b):
    """两个 8×8 RGB 指纹的平均绝对差；任一为空返回 +∞。"""
    if not a or not b or len(a) != len(b):
        return float('inf')
    return sum(abs(x - y) for x, y in zip(a, b)) / float(len(a))


def _match_ref(sig, keys=None):
    """把指纹匹配到参考图片名；返回 (key, mae) 或 (None, inf)。"""
    best, best_d = None, float('inf')
    for k in (keys or _REF_SIG.keys()):
        d = _sig_mae(sig, _REF_SIG[k])
        if d < best_d:
            best, best_d = k, d
    if best_d <= SIG_MATCH_MAE:
        return best, best_d
    return None, best_d


def _prst(shape):
    """取自动图形的预设几何名（prstGeom@prst），无则返回 ''。"""
    try:
        g = shape._element.find(f'.//{A_NS}prstGeom')
        return '' if g is None else (g.get('prst') or '')
    except Exception:
        return ''


def _run_size(shape, para, run):
    """解析 run 的实际字号（磅）：run 显式 sz → 段落 pPr/defRPr → 形状
    lstStyle/lvl1pPr/defRPr → 缺省 18 磅。
    与办公软件对齐：PowerPoint 里未显式设字号的文字会继承段落/占位符样式，
    只读 run.font.size 会全是 None，必须逐级回溯。"""
    try:
        if run.font.size is not None:
            return float(run.font.size.pt)
    except Exception:
        pass
    pPr = para._p.find(f'{A_NS}pPr')
    if pPr is not None:
        d = pPr.find(f'{A_NS}defRPr')
        if d is not None and d.get('sz'):
            return int(d.get('sz')) / 100.0
    try:
        lst = shape.text_frame._txBody.find(f'{A_NS}lstStyle')
    except Exception:
        lst = None
    if lst is not None:
        lvl = lst.find(f'{A_NS}lvl1pPr')
        if lvl is not None:
            d = lvl.find(f'{A_NS}defRPr')
            if d is not None and d.get('sz'):
                return int(d.get('sz')) / 100.0
    return 18.0


def _run_bold(shape, para, run):
    """解析 run 是否加粗，回溯顺序与 _run_size 一致；未设置视为不加粗。"""
    try:
        if run.font.bold is not None:
            return bool(run.font.bold)
    except Exception:
        pass
    pPr = para._p.find(f'{A_NS}pPr')
    if pPr is not None:
        d = pPr.find(f'{A_NS}defRPr')
        if d is not None and d.get('b') is not None:
            return d.get('b') in ('1', 'true')
    try:
        lst = shape.text_frame._txBody.find(f'{A_NS}lstStyle')
    except Exception:
        lst = None
    if lst is not None:
        lvl = lst.find(f'{A_NS}lvl1pPr')
        if lvl is not None:
            d = lvl.find(f'{A_NS}defRPr')
            if d is not None and d.get('b') is not None:
                return d.get('b') in ('1', 'true')
    return False


def _paragraphs(slide):
    """产出全页（含组合内）所有非空段落：(shape, rect_cm, text, max_pt, all_bold)。"""
    out = []
    for sh, rect, _d in _walk_shapes(slide.shapes):
        if not getattr(sh, 'has_text_frame', False):
            continue
        for para in sh.text_frame.paragraphs:
            txt = _norm(para.text)
            if not txt:
                continue
            runs = [r for r in para.runs if _norm(r.text)]
            if not runs:
                continue
            sizes = [_run_size(sh, para, r) for r in runs]
            bolds = [_run_bold(sh, para, r) for r in runs]
            out.append((sh, rect, txt, max(sizes), all(bolds)))
    return out


def _title_of(slide, paras=None):
    """页面标题 = 页内"字号最大的短段落"（去空白后 ≤20 字），并列取更靠上者。
    与办公软件对齐：这些页面并不使用标题占位符，人眼判断"这页叫什么"就是看
    最大的那行字；目录页里"01""团队介绍"等小字不会被误当成标题。"""
    best_key, best_txt = None, None
    for sh, rect, txt, size, _b in (_paragraphs(slide) if paras is None else paras):
        if len(txt) > 20:
            continue
        top = 0.0 if rect is None else rect[1]
        key = (round(size, 2), -top)
        if best_key is None or key > best_key:
            best_key, best_txt = key, txt
    return best_txt


# 单字符估宽（单位 em）：中日韩全角字符 1.0；拉丁字母约 0.55；数字约 0.52；
# 半角标点/空格约 0.34。用于"单行完整显示""文字未被截断"这类判定。
def _text_width_cm(text, size_pt):
    """按字符类别估算一行文字的宽度（厘米）。"""
    em = 0.0
    for ch in (text or ''):
        o = ord(ch)
        if o > 0x2E7F:          # CJK / 全角标点 / 假名等
            em += 1.0
        elif ch.isdigit():
            em += 0.52
        elif ch.isalpha():
            em += 0.55
        else:
            em += 0.34
    return em * size_pt * 0.0352778


def _body_ins(shape):
    """取文本框左右内边距（厘米），未显式设置时用 PowerPoint 默认 0.25cm。"""
    lins = rins = 91440.0
    try:
        bp = shape.text_frame._txBody.find(f'{A_NS}bodyPr')
        if bp is not None:
            if bp.get('lIns') is not None:
                lins = float(bp.get('lIns'))
            if bp.get('rIns') is not None:
                rins = float(bp.get('rIns'))
    except Exception:
        pass
    return _cm(lins), _cm(rins)


def _wrap_off(shape):
    """文本框是否关闭自动换行（bodyPr@wrap="none"）。"""
    try:
        bp = shape.text_frame._txBody.find(f'{A_NS}bodyPr')
        return bp is not None and bp.get('wrap') == 'none'
    except Exception:
        return False


def _para_align(shape, para):
    """段落水平对齐：'l' / 'ctr' / 'r' / 'just'；未显式设置按 'l'（左对齐）。"""
    pPr = para._p.find(f'{A_NS}pPr')
    if pPr is not None and pPr.get('algn'):
        return pPr.get('algn')
    try:
        lst = shape.text_frame._txBody.find(f'{A_NS}lstStyle')
        if lst is not None:
            lvl = lst.find(f'{A_NS}lvl1pPr')
            if lvl is not None and lvl.get('algn'):
                return lvl.get('algn')
    except Exception:
        pass
    return 'l'


def _ink_rect(shape, rect):
    """估算文本形状里"墨迹"（真正有字的那块）的矩形，横向按对齐方式定位。
    与办公软件对齐：PowerPoint 的文本框常常远宽于文字本身（例如整行居中的
    标题框跨满整页），若直接用文本框外框判断"是否压盖"会把肉眼完全不相碰的
    两处判成压盖，故横向改用文字估宽。纵向仍用文本框范围（更保守）。"""
    if rect is None:
        return None
    x, y, w, h = rect
    lins, rins = _body_ins(shape)
    avail = max(0.0, w - lins - rins)
    best_w, best_algn = 0.0, 'l'
    for para in shape.text_frame.paragraphs:
        txt = _norm(para.text)
        if not txt:
            continue
        runs = [r for r in para.runs if _norm(r.text)]
        size = max([_run_size(shape, para, r) for r in runs]) if runs else 18.0
        pw = _text_width_cm(txt, size)
        if not _wrap_off(shape):
            pw = min(pw, avail)
        if pw > best_w:
            best_w, best_algn = pw, _para_align(shape, para)
    if best_w <= 0:
        return rect
    if best_algn == 'ctr':
        nx = x + lins + (avail - best_w) / 2.0
    elif best_algn == 'r':
        nx = x + w - rins - best_w
    else:
        nx = x + lins
    return (nx, y, best_w, h)


def _overlap(r1, r2, tol=0.0):
    """两矩形的重叠宽×高（厘米）；tol 为容差（各边向内收缩 tol 后再判）。"""
    if r1 is None or r2 is None:
        return 0.0, 0.0
    ox = min(r1[0] + r1[2], r2[0] + r2[2]) - max(r1[0], r2[0]) - 2 * tol
    oy = min(r1[1] + r1[3], r2[1] + r2[3]) - max(r1[1], r2[1]) - 2 * tol
    return max(0.0, ox), max(0.0, oy)


def _fill_desc(shape):
    """形状填充摘要：'srgb:RRGGBB' / 'scheme:name' / 'blip' / 'grad' / 'noFill' / '-'。"""
    el = shape._element
    sp = el.find(f'.//{A_NS}solidFill')
    if sp is not None:
        c = sp.find(f'{A_NS}srgbClr')
        if c is not None:
            return 'srgb:' + (c.get('val') or '')
        c = sp.find(f'{A_NS}schemeClr')
        if c is not None:
            return 'scheme:' + (c.get('val') or '')
    if el.find(f'.//{A_NS}blipFill') is not None:
        return 'blip'
    if el.find(f'.//{A_NS}gradFill') is not None:
        return 'grad'
    if el.find(f'.//{A_NS}noFill') is not None:
        return 'noFill'
    return '-'


def _is_white_fill(shape):
    """是否白色填充：srgbClr FFFFFF/FEFEFE 之类的近白，或主题色 bg1/lt1。"""
    d = _fill_desc(shape)
    if d.startswith('scheme:'):
        return d.split(':', 1)[1] in ('bg1', 'lt1', 'bg2', 'lt2')
    if d.startswith('srgb:') and len(d) == 11:
        try:
            r, g, b = int(d[5:7], 16), int(d[7:9], 16), int(d[9:11], 16)
            return min(r, g, b) >= 235
        except ValueError:
            return False
    return False


def _red_line_frame(shape):
    """是否"红色矩形框线"式的参考标注：矩形/圆角矩形 + 红色描边 + 无填充。
    橙色主题色（如 FD691E / FEA741）绿分量偏高，不会被判成红色。"""
    if _prst(shape) not in ('rect', 'roundRect'):
        return False
    if _fill_desc(shape) not in ('noFill', '-'):
        return False
    ln = shape._element.find(f'.//{A_NS}ln')
    if ln is None:
        return False
    c = ln.find(f'{A_NS}solidFill/{A_NS}srgbClr')
    if c is None or not c.get('val') or len(c.get('val')) != 6:
        return False
    try:
        r, g, b = (int(c.get('val')[i:i + 2], 16) for i in (0, 2, 4))
    except ValueError:
        return False
    return r >= 150 and g <= 100 and b <= 100 and (r - max(g, b)) >= 60


def _transition_info(slide):
    """解析该页的"切换效果"：返回 {'has','morph','dur_ms','adv_tm','adv_click'}。
    与办公软件对齐：PowerPoint 把"平滑（Morph）"写在 mc:AlternateContent 的
    mc:Choice 分支里（p159:morph / p14:morph），并在 mc:Fallback 里放一个淡出等
    退化效果；因此同一页可能出现多个 <p:transition>，须优先取带 morph 的那个。"""
    info = {'has': False, 'morph': False, 'dur_ms': None,
            'adv_tm': None, 'adv_click': None}
    try:
        xml = slide._element.xml
    except Exception:
        return info
    blocks = re.findall(r'<p:transition\b.*?(?:/>|</p:transition>)', xml, re.S)
    if not blocks:
        return info
    info['has'] = True
    pick = None
    for b in blocks:
        if 'morph' in b:
            pick = b
            info['morph'] = True
            break
    if pick is None:
        pick = blocks[0]
    m = re.search(r'\bdur="(\d+)"', pick)
    if m:
        info['dur_ms'] = int(m.group(1))
    for b in blocks:                      # advTm/advClick 在各分支上一致，取首个出现者
        if info['adv_tm'] is None:
            m = re.search(r'\badvTm="(\d+)"', b)
            if m:
                info['adv_tm'] = int(m.group(1))
        if info['adv_click'] is None:
            m = re.search(r'\badvClick="([^"]+)"', b)
            if m:
                info['adv_click'] = m.group(1)
    return info


def _element_sig(shape, rect, text):
    """页面元素指纹：用于"5 页除卡片外其余元素逐页完全一致"的比对。
    取（形状类别、预设几何、填充摘要、文字、位置与大小四舍五入到 0.01cm）。
    不含形状名/ID：复制页时 PowerPoint 会重编 ID，肉眼看不出差别，不应算差异。"""
    r = ('-', '-', '-', '-') if rect is None else tuple(round(v, 2) for v in rect)
    return (str(shape.shape_type), _prst(shape), _fill_desc(shape), text, r)


def _is_top_level(shape):
    """是否幻灯片顶层形状（不在组合内）—— 对应"可单独选中与移动"。"""
    try:
        return shape._element.getparent().tag.endswith('}spTree')
    except Exception:
        return True


def _scan_pages(prs):
    """预扫描全篇：每页收集形状、段落、位图（含内容指纹）、标题与切换效果。"""
    pages = []
    for no, sl in enumerate(prs.slides, 1):
        shapes = list(_walk_shapes(sl.shapes))
        paras = _paragraphs(sl)
        pics = []
        for z, (sh, rect, depth) in enumerate(shapes):
            blob = _img_blob(sh, sl)
            if not blob:
                continue
            key, mae = _match_ref(_img_sig(blob))
            pics.append({'z': z, 'sh': sh, 'rect': rect, 'key': key, 'mae': mae,
                         'frames': _img_frames(blob), 'depth': depth,
                         'is_pic': sh.shape_type == MSO_PICTURE})
        pages.append({'no': no, 'sl': sl, 'shapes': shapes, 'paras': paras,
                      'texts': [p[2] for p in paras], 'pics': pics,
                      'cards': [p for p in pics if p['key'] in _CARD_KEYS],
                      'title': _title_of(sl, paras), 'trans': _transition_info(sl)})
    return pages


def _add(results, delta, fired, text, detail=""):
    """追加一条评分结果，格式为 "+N：规则原文" / "-N：规则原文"。
    delta > 0 为加分项，fired 表示达标；delta < 0 为减分项，fired 表示触发扣分。
    未命中时分别写作"+0：…"/"-0：…"，便于末端反解为结构化 dim2_items。

    detail 是判定时收集的实测值，只用于本地调试，不写入返回信息：命中与未命中
    登记的规则原文完全一致，返回的 rule 不随判定结果变化。"""
    if fired:
        head = ("+%d" % delta) if delta > 0 else ("%d" % delta)
    else:
        head = "+0" if delta > 0 else "-0"
    results.append(f"{head}：{text}")


def _column_lines(page, head):
    """定位页面上以 head 开头的栏目（如"01.品牌""02.概述"）并返回其正文各行。
    与办公软件对齐：原稿中栏目标题与栏目正文是两个并列文本框，且两栏的标题框
    坐标完全重合（放映时靠动画依次出现），按坐标"就近取下方"会错配，
    因此改按文档顺序取标题框之后紧邻的那个文本框；若标题框自身含多段则取第 2 段起。
    找不到该栏目返回 None（区别于"栏目存在但没有正文"的 []）。"""
    seq = [sh for sh, _r, _d in page['shapes']
           if getattr(sh, 'has_text_frame', False) and _norm(sh.text_frame.text)]
    for k, sh in enumerate(seq):
        if not _norm(sh.text_frame.text).startswith(head):
            continue
        lines = [_norm(p.text) for p in sh.text_frame.paragraphs if _norm(p.text)]
        if len(lines) >= 2:
            return lines[1:]
        if k + 1 < len(seq):
            return [_norm(p.text) for p in seq[k + 1].text_frame.paragraphs
                    if _norm(p.text)]
        return []
    return None


def _dim1_fail(result, reason):
    """维度一未通过：完成度总分记 0，不再逐条罗列（未真正逐条评过）。"""
    result["dim1_pass"] = False
    result["dim1_reason"] = reason or "维度一未通过（未给出具体原因）"
    result["dim2_items"] = []
    result["total_score"] = 0
    result["max_score"] = MAX_SCORE_TOTAL
    return result


def _evaluate_impl(file_path, result):
    # ---------------- 维度一：文件可用性（门禁） ----------------
    # 交付 1 个 .pptx 格式的演示文稿，文件存在，能被 PowerPoint 或兼容工具正常
    # 打开，不提示文件损坏或需要修复，全部幻灯片均可正常解析与显示。
    # 说明：文件已在目录内找到、但包结构损坏或需要修复而打不开，属于维度一未通过
    # （status 仍为 ok，由 _dim1_fail 记 0 分并写错误摘要），不是"没评上"；此处再
    # 逐页触碰形状与文字，捕获"能打开但某页解析失败"的情况。
    try:
        prs = _open_presentation(file_path)
    except Exception as exc:
        return _dim1_fail(
            result, f"交付文件无法正常打开（损坏或需修复）：{type(exc).__name__}: {exc}")
    slides = list(prs.slides)
    if not slides:
        return _dim1_fail(result, "演示文稿中不含任何幻灯片")
    for no, sl in enumerate(slides, 1):
        try:
            for sh, _r, _d in _walk_shapes(sl.shapes):
                if getattr(sh, 'has_text_frame', False):
                    _ = sh.text_frame.text
        except Exception as exc:
            return _dim1_fail(
                result, f"第 {no} 页无法正常解析：{type(exc).__name__}: {exc}")
    result["dim1_pass"] = True
    # 约定：维度一通过时 dim1_reason 保持空串，只有未通过才写错误摘要。
    result["dim1_reason"] = ""

    # ---------------- 维度二：完成度评分 ----------------
    pw, ph = _cm(prs.slide_width), _cm(prs.slide_height)
    pages = _scan_pages(prs)
    cover = pages[0]
    results = []

    # 首页迁入的两张配图：只按"图片内容指纹"认定，不依赖形状名或替代文字
    # （原稿中右图的替代文字与画面内容并不一致，名称同样不可靠）。
    cand_l = [p for p in cover['pics'] if p['key'] == 'LEFT_BOOK' and p['rect']]
    cand_r = [p for p in cover['pics'] if p['key'] == 'RIGHT_CLASS' and p['rect']]
    img_l = min(cand_l, key=lambda p: p['mae']) if cand_l else None
    img_r = min(cand_r, key=lambda p: p['mae']) if cand_r else None

    # 首页迁入的四段文字：文本 → 所在形状、矩形、字号、加粗
    mig = {}
    for sh, rect, txt, size, bold in cover['paras']:
        if txt in MIGRATED_TEXTS and txt not in mig:
            mig[txt] = {'sh': sh, 'rect': rect, 'size': size, 'bold': bold}

    # 首页四处参照文字所在形状（英文串按忽略大小写比对：原稿用 cap="all"
    # 让小写字符串显示为全大写，直接按大写比对会判不出来）。
    ref_shapes = {}
    for sh, rect, txt, _s, _b in cover['paras']:
        low = txt.lower()
        for want in COVER_TEXTS:
            if want.lower() == low and want not in ref_shapes:
                ref_shapes[want] = (sh, rect)

    # "整体区域" = 两张图片 + 四段文字的外接矩形
    region_rects = [p['rect'] for p in (img_l, img_r) if p and p['rect']]
    region_rects += [m['rect'] for m in mig.values() if m['rect']]
    if region_rects:
        rx0 = min(r[0] for r in region_rects)
        ry0 = min(r[1] for r in region_rects)
        rx1 = max(r[0] + r[2] for r in region_rects)
        ry1 = max(r[1] + r[3] for r in region_rects)
        region = (rx0, ry0, rx1 - rx0, ry1 - ry0)
    else:
        region = None

    # 卡片页：标题为"教学团队介绍"的页面；卡片按水平中心从左至右排序
    card_pages = [pg for pg in pages if pg['title'] == T_TEAM_PAGE]
    for pg in card_pages:
        pg['cards_lr'] = sorted([c for c in pg['cards'] if c['rect']],
                                key=lambda c: c['rect'][0] + c['rect'][2] / 2.0)
    # 承载卡片的页面（用于白色底板等装饰的检查）：优先按标题，其次按"页面上
    # 出现了老师卡片图片"兜底，两者都没有时该减分项视为不适用。
    board_pages = card_pages or [pg for pg in pages if pg['cards']]

    # ---------------- 01（+5）首页右侧两张圆角矩形图片 ----------------
    # ①两张图片按"图片内容"认定：左图=翻开的书本与发光灯泡，右图=教室内学生
    #   举手的课堂场景（8×8 RGB 缩略图指纹，容差 SIG_MATCH_MAE）；
    # ②承载形式为圆角矩形（prstGeom=roundRect）或图片对象；
    # ③宽度差 ≤0.1cm 且高度差 ≤0.1cm；④左右并排（左图右边界不越过右图左边界）；
    # ⑤两者上边缘垂直差 ≤0.15cm；⑥位于首页右侧（两图水平中心均在页面右半侧）。
    # 与办公软件对齐：PowerPoint 中"插入图片再裁成圆角"与"给圆角矩形填充图片"
    # 观感完全相同，两种承载都算"圆角矩形图片"。
    if img_l and img_r and img_l['sh'] is not img_r['sh']:
        rl, rr = img_l['rect'], img_r['rect']
        d_w, d_h = abs(rl[2] - rr[2]), abs(rl[3] - rr[3])
        d_top = abs(rl[1] - rr[1])
        shape_ok = all(_prst(p['sh']) == 'roundRect' or p['sh'].shape_type == MSO_PICTURE
                       for p in (img_l, img_r))
        side_by_side = rl[0] + rl[2] <= rr[0] + 0.05
        right_half = min(rl[0] + rl[2] / 2.0, rr[0] + rr[2] / 2.0) > pw / 2.0
        ok01 = bool(shape_ok and d_w <= 0.1 and d_h <= 0.1 and d_top <= 0.15
                    and side_by_side and right_half)
        det01 = (f"左图{_prst(img_l['sh']) or img_l['sh'].shape_type}"
                 f"({rl[0]:.2f},{rl[1]:.2f}) {rl[2]:.2f}×{rl[3]:.2f}cm，"
                 f"右图({rr[0]:.2f},{rr[1]:.2f}) {rr[2]:.2f}×{rr[3]:.2f}cm，"
                 f"宽差{d_w:.2f}cm、高差{d_h:.2f}cm、上边缘差{d_top:.2f}cm，"
                 f"并排={'是' if side_by_side else '否'}")
    else:
        ok01, det01 = False, ("首页未按图片内容找到"
                             + ("左图（书本与灯泡）" if not img_l else "")
                             + ("右图（教室举手）" if not img_r else ""))
    _add(results, 5, ok01,
         "首页右侧出现两张圆角矩形图片，左图内容为翻开的书本与发光灯泡，右图内容为教室内学生举手的课堂场景；"
         "两张图片的宽度差与高度差均不超过0.1cm，左右并排且两者上边缘的垂直差不超过0.15cm", det01)

    # ---------------- 02（+3）两图宽高比与单张尺寸 ----------------
    # ①两张图片宽高比均落在 1.55–1.65（原图 1.60，允许 ±0.05 的等比缩放误差）；
    # ②单张宽度 4.0–6.6cm、高度 2.5–4.1cm。
    if img_l and img_r:
        ars = [(p['rect'][2] / p['rect'][3] if p['rect'][3] else 0.0) for p in (img_l, img_r)]
        ws = [p['rect'][2] for p in (img_l, img_r)]
        hs = [p['rect'][3] for p in (img_l, img_r)]
        ok02 = (all(1.55 <= a <= 1.65 for a in ars)
                and all(4.0 <= w <= 6.6 for w in ws)
                and all(2.5 <= h <= 4.1 for h in hs))
        det02 = (f"宽高比{ars[0]:.3f}/{ars[1]:.3f}，宽{ws[0]:.2f}/{ws[1]:.2f}cm，"
                 f"高{hs[0]:.2f}/{hs[1]:.2f}cm")
    else:
        ok02, det02 = False, "两张图片未同时找到，无法核对比例与尺寸"
    _add(results, 3, ok02,
         "首页右侧两张图片的宽高比均为1.55–1.65（与原“品牌宣言”页两图1.60的宽高比保持一致，未被单向拉伸变形），"
         "单张宽度4.0–6.6cm、高度2.5–4.1cm", det02)

    # ---------------- 03（+3）两图与四段文字均为独立可编辑对象 ----------------
    # ①两张图片是两个不同的形状对象，且均为幻灯片顶层形状（可单独选中与移动）；
    # ②四段文字分别落在 4 个不同的文本形状里，同样位于顶层；
    # ③四段文字仍是可编辑文字（能取到 text_frame），未被拼合成位图或整体截图。
    # 与办公软件对齐：若把这块内容"复制为图片"再粘贴，只会得到 1 个图片对象，
    # 四段文字也就取不到文本，②③随之不成立。
    mig_shapes = [m['sh'] for m in mig.values()]
    uniq_txt_shapes = {id(s) for s in mig_shapes}
    ok03 = bool(img_l and img_r and img_l['sh'] is not img_r['sh']
                and _is_top_level(img_l['sh']) and _is_top_level(img_r['sh'])
                and len(mig) == 4 and len(uniq_txt_shapes) == 4
                and all(_is_top_level(s) for s in mig_shapes))
    det03 = (f"图片对象{2 if (img_l and img_r and img_l['sh'] is not img_r['sh']) else len([p for p in (img_l, img_r) if p])}个、"
             f"独立文本框{len(uniq_txt_shapes)}个（四段文字命中{len(mig)}段）")
    _add(results, 3, ok03,
         "首页右侧的两张图片为两个各自独立、可单独选中与移动的图片对象（图片形状或图片填充的圆角矩形形状），"
         "其下方四段文字为4个各自独立、可编辑修改的文本框，未被拼合为单张位图，也未以整体截图粘贴的方式呈现", det03)

    # ---------------- 04（+5）四段文字的分组、上下排列与相对图片的位置 ----------------
    # ①四段文字均出现在首页；②"持续精进"+"学习成长与迭代创新"归左图，
    #   "实践见效"+"把方法落实到课堂与服务"归右图，且每组内两段自上而下排列；
    # ③每段文字的水平中心与其上方图片的水平中心偏差 ≤0.8cm；
    # ④两段大字（"持续精进""实践见效"）顶部位于对应图片下边缘以下 0.2–2.2cm。
    # 与办公软件对齐：文本框常远宽于文字本身，但因四段文字都是居中对齐，
    # 文本框水平中心即文字水平中心，故此处用文本框中心比对。
    pair_specs = ((TXT_BIG_L, TXT_SUB_L, img_l, "左"), (TXT_BIG_R, TXT_SUB_R, img_r, "右"))
    ok04, det04_parts = True, []
    for big, sub, img, tag in pair_specs:
        mb, ms = mig.get(big), mig.get(sub)
        if not (mb and ms and mb['rect'] and ms['rect'] and img and img['rect']):
            ok04 = False
            det04_parts.append(f"{tag}组缺失（{big}/{sub}或对应图片未找到）")
            continue
        ir = img['rect']
        icx, ibot = ir[0] + ir[2] / 2.0, ir[1] + ir[3]
        d_big = abs(mb['rect'][0] + mb['rect'][2] / 2.0 - icx)
        d_sub = abs(ms['rect'][0] + ms['rect'][2] / 2.0 - icx)
        gap = mb['rect'][1] - ibot
        order_ok = ms['rect'][1] > mb['rect'][1]
        if not (d_big <= 0.8 and d_sub <= 0.8 and 0.2 <= gap <= 2.2 and order_ok):
            ok04 = False
        det04_parts.append(f"{tag}组中心偏差{d_big:.2f}/{d_sub:.2f}cm、"
                           f"大字顶部距图片下边缘{gap:.2f}cm、上下顺序{'正确' if order_ok else '颠倒'}")
    _add(results, 5, ok04,
         "首页出现“持续精进”“学习成长与迭代创新”“实践见效”“把方法落实到课堂与服务”四段文字："
         "“持续精进”“学习成长与迭代创新”位于左图正下方并自上而下排列，"
         "“实践见效”“把方法落实到课堂与服务”位于右图正下方并自上而下排列；"
         "每段文字的水平中心与其上方图片的水平中心偏差不超过0.8cm，"
         "“持续精进”“实践见效”的顶部位于对应图片下边缘以下0.2–2.2cm", "；".join(det04_parts))

    # ---------------- 05（+3）四段文字的字号、加粗与单行完整显示 ----------------
    # ①"持续精进""实践见效"字号 20 磅且加粗；②两段小字字号 14 磅；
    # ③四段文字均在一行内完整显示：按字符类别估算的行宽不超过文本框可用宽度
    #   （宽度减左右内边距），或该文本框已关闭自动换行且未被边界截断。
    # 允许 ±0.5 磅误差，吸收 pt↔EMU 往返换算带来的尾差。
    ok05, det05_parts = True, []
    for txt, want in ((TXT_BIG_L, 20.0), (TXT_BIG_R, 20.0),
                      (TXT_SUB_L, 14.0), (TXT_SUB_R, 14.0)):
        m = mig.get(txt)
        if not m or not m['rect']:
            ok05 = False
            det05_parts.append(f"“{txt}”未找到")
            continue
        size_ok = abs(m['size'] - want) <= 0.5
        bold_ok = (want < 20.0) or m['bold']
        lins, rins = _body_ins(m['sh'])
        avail = m['rect'][2] - lins - rins
        need = _text_width_cm(txt, m['size'])
        one_line = _wrap_off(m['sh']) or need <= avail + 0.05
        if not (size_ok and bold_ok and one_line):
            ok05 = False
        det05_parts.append(f"“{txt}”{m['size']:.2f}磅"
                           f"{'加粗' if m['bold'] else '未加粗'}、"
                           f"估算行宽{need:.2f}cm/可用{avail:.2f}cm")
    _add(results, 3, ok05,
         "首页“持续精进”“实践见效”字号为20磅且加粗，“学习成长与迭代创新”“把方法落实到课堂与服务”字号为14磅，"
         "且四段文字均在一行内完整显示，未出现折行或被文本框边界截断", "；".join(det05_parts))

    # ---------------- 06（+3）整体区域与参考示意图红框位置一致 ----------------
    # 逐项换算成页面百分比：①左图左边界 58%–65%；②右图右边界 93%–100%；
    # ③两图上边缘均 36%–44%；④两图水平间距为单张图片宽度的 15%–55%；
    # ⑤整体区域（两图 + 四段文字的外接矩形）右边界不超过 100%；
    # ⑥整体区域垂直中心 52%–63%。
    if img_l and img_r and region:
        rl, rr = img_l['rect'], img_r['rect']
        p_left = rl[0] / pw * 100.0
        p_right = (rr[0] + rr[2]) / pw * 100.0
        p_top_l, p_top_r = rl[1] / ph * 100.0, rr[1] / ph * 100.0
        gap_cm = rr[0] - (rl[0] + rl[2])
        avg_w = (rl[2] + rr[2]) / 2.0
        p_gap = (gap_cm / avg_w * 100.0) if avg_w else 0.0
        p_reg_right = (region[0] + region[2]) / pw * 100.0
        p_reg_vc = (region[1] + region[3] / 2.0) / ph * 100.0
        ok06 = (58.0 <= p_left <= 65.0 and 93.0 <= p_right <= 100.0
                and 36.0 <= p_top_l <= 44.0 and 36.0 <= p_top_r <= 44.0
                and 15.0 <= p_gap <= 55.0 and p_reg_right <= 100.0
                and 52.0 <= p_reg_vc <= 63.0)
        det06 = (f"左图左边界{p_left:.1f}%、右图右边界{p_right:.1f}%、上边缘{p_top_l:.1f}%/"
                 f"{p_top_r:.1f}%、图间距{gap_cm:.2f}cm={p_gap:.1f}%单图宽、"
                 f"区域右边界{p_reg_right:.1f}%、垂直中心{p_reg_vc:.1f}%")
    else:
        ok06, det06 = False, "两张图片或四段文字未齐备，无法核对整体区域位置"
    _add(results, 3, ok06,
         "首页右侧两张图片与四段文字构成的整体区域，其位置与输入参考示意图中红框标注的位置一致："
         "左图左边界位于页面从左至右58%–65%、右图右边界位于93%–100%，两张图片的上边缘均位于页面从上至下36%–44%，"
         "两图之间的水平间距为单张图片宽度的15%–55%，整体区域右边界不超过100%、垂直中心位于页面从上至下52%–63%", det06)

    # ---------------- 07（+1）整体区域在右半侧且与首页四处文字互不压盖 ----------------
    # ①整体区域左边界不小于页面一半宽度；②区域内每个元素与首页四处参照文字的
    #   "文字墨迹矩形"均无重叠（重叠判定要求横竖两向同时相交）。
    # 与办公软件对齐："公司简介"这类标题的文本框横跨半页，但真正有笔画的只有
    # 左端几个字，用文本框外框判压盖会把肉眼毫不相碰的两处判成压盖，
    # 故此处对参照文字改用按字号估算的墨迹矩形（见 _ink_rect）。
    if region and img_l and img_r:
        in_right = region[0] >= pw / 2.0 - 0.01
        clash = []
        for want, (sh, rect) in ref_shapes.items():
            ink = _ink_rect(sh, rect)
            for er in region_rects:
                ox, oy = _overlap(ink, er)
                if ox > 0.0 and oy > 0.0:
                    clash.append(f"{want}重叠{ox:.2f}×{oy:.2f}cm")
                    break
        ok07 = bool(in_right and not clash and len(ref_shapes) == len(COVER_TEXTS))
        det07 = (f"区域左边界{region[0] / pw * 100.0:.1f}%（页面右半侧起点50%），"
                 + ("与四处参照文字均无压盖" if not clash else "、".join(clash)))
    else:
        ok07, det07 = False, "整体区域未成形，无法核对"
    _add(results, 1, ok07,
         "首页该整体区域完全位于页面右半侧，与页面左侧“2026”“公司简介”“品牌介绍手册”"
         "“SIMPLE ATMOSPHERE COMPANY PROFILE”四处文字（此处仅作位置参照）在视觉上互不压盖，"
         "文字笔画未被图片或文字遮挡", det07)

    # ---------------- 08（+3）全篇不出现"品牌宣言"页 ----------------
    # 原稿的"品牌宣言"页是两图与四段文字的来源页，按指令须整页删除。
    dec_pages = [pg['no'] for pg in pages if pg['title'] == T_DECLARE_PAGE]
    _add(results, 3, not dec_pages,
         "全篇不出现标题为“品牌宣言”的页面",
         "未出现" if not dec_pages else "仍出现于第 " + "、".join(str(n) for n in dec_pages) + " 页")

    # ---------------- 09（+5）5 张卡片为 5 个独立图片对象 ----------------
    # ①存在标题为"教学团队介绍"的页面；②每页恰好 5 张按内容认出的老师卡片；
    # ③每张都是独立的图片对象（PICTURE）且位于幻灯片顶层，可单独选中与移动；
    # ④页面上不存在多帧位图（GIF 动图），即未用一张动图承载全部卡片内容。
    ok09, det09_parts = bool(card_pages), []
    for pg in card_pages:
        cards = pg['cards']
        ids = {id(c['sh']) for c in cards}
        gif = [p for p in pg['pics'] if p['frames'] > 1]
        page_ok = (len(cards) == 5 and len(ids) == 5
                   and all(c['is_pic'] and _is_top_level(c['sh']) for c in cards)
                   and all(c['frames'] == 1 for c in cards) and not gif)
        if not page_ok:
            ok09 = False
        det09_parts.append(f"第{pg['no']}页卡片图片对象{len(cards)}个"
                           + (f"、含{len(gif)}张多帧动图" if gif else ""))
    _add(results, 5, ok09,
         "标题为“教学团队介绍”的每一页上，5位老师的简历卡片均为5个各自独立、可单独选中与移动的图片对象，"
         "未以单张整幅位图或单张GIF动图承载全部卡片内容",
         "；".join(det09_parts) if det09_parts else "全篇无标题为“教学团队介绍”的页面")

    # ---------------- 10（+5）5 张橙白配色卡片且人像互不相同 ----------------
    # 按图片内容指纹把每张卡片映射到 5 位老师之一（CARD_1..CARD_5），
    # 要求每页 5 张卡片恰好覆盖 5 位不同老师（互不重复即人像互不相同）。
    ok10, det10_parts = bool(card_pages), []
    for pg in card_pages:
        keys = [c['key'] for c in pg['cards']]
        if sorted(keys) != sorted(_CARD_KEYS):
            ok10 = False
        det10_parts.append(f"第{pg['no']}页命中老师{len(set(keys))}位")
    _add(results, 5, ok10,
         "标题为“教学团队介绍”的每一页上均出现5张橙白配色的老师简历卡片，5张卡片的人像互不相同"
         "（蓝衬衫戴眼镜男士、浅色西装长发女士、棕色毛衣男士、深色衬衫戴眼镜男士、橙色外套卷发女士）",
         "；".join(det10_parts) if det10_parts else "全篇无标题为“教学团队介绍”的页面")

    # ---------------- 11（+3）全篇不出现"学习支持团队介绍"页及其文字 ----------------
    # ①无标题为"学习支持团队介绍"的页面；②全篇无"服务亮点"文本框；
    # ③全篇无该页的说明段（按去空白并统一引号后比对，容忍直/弯引号差异）。
    sup_desc = ("学习支持团队负责排课协调、进度跟踪与家校沟通，围绕学习目标建立服务台账，"
                "及时汇总反馈并推动改进。通过标准流程与灵活响应，保障每位学习者获得稳定陪伴。")
    sup_pages = [pg['no'] for pg in pages if pg['title'] == T_SUPPORT_PAGE]
    hl_pages = [pg['no'] for pg in pages if any(t == "服务亮点" for t in pg['texts'])]
    desc_key = _norm_q(sup_desc)
    desc_pages = [pg['no'] for pg in pages
                  if any(desc_key in _norm_q(t) for t in pg['texts'])]
    ok11 = not (sup_pages or hl_pages or desc_pages)
    det11 = ("均未出现" if ok11 else
             "；".join(x for x in (
                 ("页面残留于第 " + "、".join(map(str, sup_pages)) + " 页") if sup_pages else "",
                 ("“服务亮点”残留于第 " + "、".join(map(str, hl_pages)) + " 页") if hl_pages else "",
                 ("说明段残留于第 " + "、".join(map(str, desc_pages)) + " 页") if desc_pages else "",
             ) if x))
    _add(results, 3, ok11,
         "全篇不出现标题为“学习支持团队介绍”的页面，也不出现“服务亮点”文本框及“" + sup_desc + "”说明段", det11)

    # ---------------- 12（+5）连续 5 张"教学团队介绍"页且非卡片元素逐页一致 ----------------
    # ①恰好 5 张标题为"教学团队介绍"的页面且页码连续；
    # ②除 5 张老师卡片外，其余元素的种类、预设几何、填充、文字、位置与大小
    #   逐页完全一致（元素指纹见 _element_sig，忽略会被 PowerPoint 重编的形状名/ID）。
    if len(card_pages) == 5 and [pg['no'] for pg in card_pages] == list(
            range(card_pages[0]['no'], card_pages[0]['no'] + 5)):
        sig_sets = []
        for pg in card_pages:
            card_ids = {id(c['sh']) for c in pg['cards']}
            sigs = []
            for sh, rect, _d in pg['shapes']:
                if id(sh) in card_ids:
                    continue
                txt = _norm(sh.text_frame.text) if getattr(sh, 'has_text_frame', False) else ''
                sigs.append(_element_sig(sh, rect, txt))
            sig_sets.append(sorted(sigs))
        same = all(s == sig_sets[0] for s in sig_sets[1:])
        ok12 = same
        diff_pages = [card_pages[k]['no'] for k, s in enumerate(sig_sets) if s != sig_sets[0]]
        det12 = (f"第 {card_pages[0]['no']}–{card_pages[-1]['no']} 页连续，"
                 f"每页非卡片元素{len(sig_sets[0])}个，"
                 + ("逐页一致" if same else "与首页不一致的页：" + "、".join(map(str, diff_pages))))
    else:
        ok12 = False
        det12 = (f"标题为“教学团队介绍”的页面共{len(card_pages)}页"
                 + (f"（第 {'、'.join(str(pg['no']) for pg in card_pages)} 页）" if card_pages else "")
                 + "，未构成连续 5 页")
    _add(results, 5, ok12,
         "全篇出现连续5张标题均为“教学团队介绍”的页面，5页之间除5张老师卡片的位置与大小不同外，"
         "页面其余所有元素的种类、内容、位置与大小逐页完全一致", det12)

    # ---------------- 13（+5）卡片"居中最大、向两侧递减"的三级高度 ----------------
    # 按水平中心从左至右取 5 张卡片：居中 1 张 15.0–17.0cm，左右紧邻各 1 张
    # 13.0–15.0cm，最外侧各 1 张 11.0–13.0cm；三级依次递减，同级两张高度相同
    # （允许 0.1cm 误差，吸收缩放取整尾差）。
    ok13, det13_parts = bool(card_pages), []
    for pg in card_pages:
        cs = pg['cards_lr']
        if len(cs) != 5:
            ok13 = False
            det13_parts.append(f"第{pg['no']}页卡片{len(cs)}张")
            continue
        hs = [c['rect'][3] for c in cs]
        tier_ok = (15.0 <= hs[2] <= 17.0
                   and all(13.0 <= h <= 15.0 for h in (hs[1], hs[3]))
                   and all(11.0 <= h <= 13.0 for h in (hs[0], hs[4]))
                   and hs[2] > hs[1] and hs[1] > hs[0]
                   and abs(hs[1] - hs[3]) <= 0.1 and abs(hs[0] - hs[4]) <= 0.1)
        if not tier_ok:
            ok13 = False
        det13_parts.append(f"第{pg['no']}页高度" + "/".join(f"{h:.2f}" for h in hs) + "cm")
    _add(results, 5, ok13,
         "每张“教学团队介绍”页上5张卡片呈“居中最大、向两侧递减”的三级尺寸："
         "居中1张高度15.0–17.0cm，其左右紧邻各1张高度13.0–15.0cm，最外侧各1张高度11.0–13.0cm，"
         "三级高度依次递减且同级两张高度相同",
         "；".join(det13_parts) if det13_parts else "全篇无标题为“教学团队介绍”的页面")

    # ---------------- 14（+3）5 张卡片宽高比 0.54–0.58 ----------------
    ok14, det14_parts = bool(card_pages), []
    for pg in card_pages:
        cs = pg['cards_lr']
        if len(cs) != 5:
            ok14 = False
            det14_parts.append(f"第{pg['no']}页卡片{len(cs)}张")
            continue
        ars = [(c['rect'][2] / c['rect'][3] if c['rect'][3] else 0.0) for c in cs]
        if not all(0.54 <= a <= 0.58 for a in ars):
            ok14 = False
        det14_parts.append(f"第{pg['no']}页宽高比" + "/".join(f"{a:.3f}" for a in ars))
    _add(results, 3, ok14,
         "每张“教学团队介绍”页上5张卡片的宽高比均为0.54–0.58",
         "；".join(det14_parts) if det14_parts else "全篇无标题为“教学团队介绍”的页面")

    # ---------------- 15（+3）共用水平中心线、纵向 56%–62%、居中卡片 48%–53% ----------------
    ok15, det15_parts = bool(card_pages), []
    for pg in card_pages:
        cs = pg['cards_lr']
        if len(cs) != 5:
            ok15 = False
            det15_parts.append(f"第{pg['no']}页卡片{len(cs)}张")
            continue
        cys = [c['rect'][1] + c['rect'][3] / 2.0 for c in cs]
        spread = max(cys) - min(cys)
        pcts = [cy / ph * 100.0 for cy in cys]
        mid_cx_pct = (cs[2]['rect'][0] + cs[2]['rect'][2] / 2.0) / pw * 100.0
        page_ok = (spread <= ph * 0.03 and all(56.0 <= p <= 62.0 for p in pcts)
                   and 48.0 <= mid_cx_pct <= 53.0)
        if not page_ok:
            ok15 = False
        det15_parts.append(f"第{pg['no']}页垂直中心极差{spread:.2f}cm、纵向"
                           f"{min(pcts):.1f}%–{max(pcts):.1f}%、居中卡片水平中心{mid_cx_pct:.1f}%")
    _add(results, 3, ok15,
         "每张“教学团队介绍”页上5张卡片共用同一水平中心线（各卡片垂直中心相互差值不超过页面高度的3%），"
         "且各卡片垂直中心均位于页面从上至下56%–62%，其中居中卡片的水平中心位于页面从左至右48%–53%",
         "；".join(det15_parts) if det15_parts else "全篇无标题为“教学团队介绍”的页面")

    # ---------------- 16（+3）相邻卡片重叠且靠中者压在靠外者之上 ----------------
    # ①相邻两张卡片横向有重叠（同时纵向相交，才是肉眼可见的层叠）；
    # ②每对重叠卡片中更靠中间的一张叠放层级更高；③居中卡片位于 5 张的最上层。
    # 与办公软件对齐：PowerPoint 的叠放次序就是 XML 中的文档顺序，越靠后越在上层。
    ok16, det16_parts = bool(card_pages), []
    for pg in card_pages:
        cs = pg['cards_lr']
        if len(cs) != 5:
            ok16 = False
            det16_parts.append(f"第{pg['no']}页卡片{len(cs)}张")
            continue
        ov, zo = [], []
        for k in range(4):
            ox, oy = _overlap(cs[k]['rect'], cs[k + 1]['rect'])
            ov.append(ox if oy > 0 else 0.0)
            inner = k if abs(k - 2) < abs(k + 1 - 2) else k + 1
            outer = k + 1 if inner == k else k
            zo.append(cs[inner]['z'] > cs[outer]['z'])
        top_ok = cs[2]['z'] == max(c['z'] for c in cs)
        page_ok = all(o > 0.0 for o in ov) and all(zo) and top_ok
        if not page_ok:
            ok16 = False
        det16_parts.append(f"第{pg['no']}页相邻重叠" + "/".join(f"{o:.2f}" for o in ov)
                           + f"cm、层级正确{sum(zo)}/4、居中最上层{'是' if top_ok else '否'}")
    _add(results, 3, ok16,
         "每张“教学团队介绍”页上相邻卡片左右相互重叠，且每两张相互重叠的卡片中更靠中间的一张压在更靠外的一张之上，"
         "居中卡片位于该页5张卡片的最上层",
         "；".join(det16_parts) if det16_parts else "全篇无标题为“教学团队介绍”的页面")

    # ---------------- 17（+5）5 页卡片排列逐页循环左移一位 ----------------
    # 每页把 5 张卡片按水平中心从左至右排成一个老师序列，要求后一页的序列等于
    # 前一页整体左移一位（首位挪到末位），且 5 页的居中老师互不相同、
    # 第 5 页再左移一位即回到第 1 页的排列。
    seqs = []
    if len(card_pages) == 5 and all(len(pg['cards_lr']) == 5 for pg in card_pages):
        seqs = [[c['key'] for c in pg['cards_lr']] for pg in card_pages]
    if seqs and all(None not in s for s in seqs):
        shift_ok = all(seqs[k] == seqs[k - 1][1:] + seqs[k - 1][:1] for k in range(1, 5))
        wrap_ok = seqs[0] == seqs[4][1:] + seqs[4][:1]
        mids = [s[2] for s in seqs]
        mid_ok = len(set(mids)) == 5
        ok17 = bool(shift_ok and wrap_ok and mid_ok)
        det17 = ("逐页排列 " + " → ".join("".join(k[-1] for k in s) for s in seqs)
                 + f"（1=第1位老师…5=第5位老师），左移正确{'是' if shift_ok else '否'}、"
                 f"首尾衔接{'是' if wrap_ok else '否'}、居中老师{len(set(mids))}位不同")
    else:
        ok17, det17 = False, "未取得 5 页 × 5 张可识别卡片的排列，无法核对循环左移"
    _add(results, 5, ok17,
         "5张“教学团队介绍”页中，5张卡片自左至右的排列顺序按页依次循环左移一位，"
         "每页由不同一位老师的卡片占据居中放大位置，第5页之后回到第1页的排列", det17)

    # ---------------- 18（+5）第 2–5 页为平滑/变体（Morph）切换且时长 1.5–2.5 秒 ----------------
    # 与办公软件对齐：PowerPoint 把"平滑"写在 mc:AlternateContent 的 mc:Choice
    # 分支（p159:morph / p14:morph）里，并在 mc:Fallback 里放淡出等退化效果；
    # 时长取 p14:dur（毫秒）。
    if len(card_pages) == 5:
        infos = [pg['trans'] for pg in card_pages[1:]]
        ok18 = all(t['morph'] and t['dur_ms'] is not None and 1500 <= t['dur_ms'] <= 2500
                   for t in infos)
        det18 = "；".join(f"第{card_pages[k + 1]['no']}页"
                          f"{'平滑' if t['morph'] else ('其他切换' if t['has'] else '无切换')}"
                          f"、时长{'—' if t['dur_ms'] is None else '%.2f秒' % (t['dur_ms'] / 1000.0)}"
                          for k, t in enumerate(infos))
    else:
        ok18, det18 = False, "不足 5 张“教学团队介绍”页，无法核对第 2–5 页切换效果"
    _add(results, 5, ok18,
         "5张“教学团队介绍”页中第2至第5页的幻灯片切换效果为“平滑／变体（Morph）”，切换时长1.5–2.5秒", det18)

    # ---------------- 19（+3）5 页均自动换片、换片时间 0 秒 ----------------
    # advTm=0 即"设置自动换片时间 00:00"；放映时无需点击即连续播放。
    if len(card_pages) == 5:
        tms = [pg['trans']['adv_tm'] for pg in card_pages]
        ok19 = all(t is not None and t == 0 for t in tms)
        det19 = "；".join(f"第{card_pages[k]['no']}页换片时间"
                          + ("未设置自动换片" if t is None else f"{t / 1000.0:.2f}秒")
                          for k, t in enumerate(tms))
    else:
        ok19, det19 = False, "不足 5 张“教学团队介绍”页，无法核对换片设置"
    _add(results, 3, ok19,
         "5张“教学团队介绍”页均设置为自动换片，换片时间为0秒，放映时无需任何操作即可连续播放出卡片整排左移的效果", det19)

    # ---------------- 20（+1）5 页位于"团队介绍"之后、"品牌文化"之前 ----------------
    team_no = next((pg['no'] for pg in pages if pg['title'] == T_TEAM_SECTION), None)
    culture_no = next((pg['no'] for pg in pages if pg['title'] == T_CULTURE), None)
    if len(card_pages) == 5 and team_no and culture_no:
        ok20 = team_no < card_pages[0]['no'] and card_pages[-1]['no'] < culture_no
        det20 = (f"“团队介绍”第{team_no}页、卡片页第{card_pages[0]['no']}–{card_pages[-1]['no']}页、"
                 f"“品牌文化”第{culture_no}页")
    else:
        ok20 = False
        det20 = (f"卡片页{len(card_pages)}页、“团队介绍”页{team_no or '缺失'}、"
                 f"“品牌文化”页{culture_no or '缺失'}")
    _add(results, 1, ok20,
         "5张“教学团队介绍”页位于标题为“团队介绍”的页面之后、标题为“品牌文化”的页面之前", det20)

    # ================= 以下为减分项：指令未要求改动的部分若被破坏则扣分 =================

    # ---------------- 21（-3）页面尺寸 ----------------
    # 33.87cm×19.05cm 即 12192000×6858000 EMU；换算回厘米为 33.8667cm，
    # 故容差取 0.05cm。与办公软件对齐：.pptx 的页面尺寸是全篇唯一的
    # （sldSz 记在 presentation.xml），不存在"各页不一致"的物理可能，
    # 该分句只在文件被拆分/合并出多套尺寸时才有意义，这里按全篇尺寸判定。
    size_bad = (abs(pw - PAGE_W_CM) > PAGE_SIZE_TOL_CM
                or abs(ph - PAGE_H_CM) > PAGE_SIZE_TOL_CM)
    _add(results, -3, size_bad,
         "页面尺寸不为33.87cm×19.05cm（16:9横版），或各页页面尺寸不一致",
         f"全篇页面尺寸 {pw:.2f}cm×{ph:.2f}cm")

    # ---------------- 22（-3）首页署名文字 ----------------
    # 取首页左下角区域（水平中心在左半侧、垂直中心在页面 60% 以下）的文字，
    # 其中任一段等于"星市静默回声教育"即视为署名保留。
    sign_txts = [txt for _sh, rect, txt, _s, _b in cover['paras']
                 if rect and (rect[0] + rect[2] / 2.0) < pw / 2.0
                 and (rect[1] + rect[3] / 2.0) > ph * 0.6]
    sign_bad = "星市静默回声教育" not in sign_txts
    _add(results, -3, sign_bad,
         "首页左下方带人形图标的署名文字不为“星市静默回声教育”",
         "实测左下方文字：" + ("、".join(sign_txts) if sign_txts else "未找到"))

    # ---------------- 23/24/25（-3/-1/-1）"品牌概述"页两栏文字 ----------------
    ov_pages = [pg for pg in pages if pg['title'] == "品牌概述"]
    col_brand = col_intro = None
    for pg in ov_pages:
        if col_brand is None:
            col_brand = _column_lines(pg, "01.品牌")
        if col_intro is None:
            col_intro = _column_lines(pg, "02.概述")
    line1 = col_brand[0] if col_brand else ""
    line2 = col_brand[1] if col_brand and len(col_brand) > 1 else ""
    body2 = "".join(col_intro) if col_intro else ""

    _add(results, -3, line1 != "星市静默回声教育科技有限公司",
         "标题为“品牌概述”的页面中“01.品牌”栏首行文字不为“星市静默回声教育科技有限公司”",
         "实测首行：" + (line1 or "未找到该栏"))

    want_line2 = "2024年春季启动  学习服务创新品牌  校社协同实践者"
    _add(results, -1, _norm(line2) != _norm(want_line2),
         "标题为“品牌概述”的页面中“01.品牌”栏第二行文字不为“" + want_line2 + "”",
         "实测第二行：" + (line2 or "未找到该栏"))

    want_body2 = ("专注于中小学阶段学习支持与综合素养提升，秉承“以学促长，实践见效”的核心理念，"
                  "围绕学习诊断、课程辅导与成长陪伴，提供灵活的教学方案，"
                  "帮助学生建立稳定的学习习惯与解决问题能力。")
    _add(results, -1, _norm_q(body2) != _norm_q(want_body2),
         "标题为“品牌概述”的页面中“02.概述”栏正文不为“" + want_body2 + "”",
         "实测正文" + (f"共{len(body2)}字，" + ("一致" if _norm_q(body2) == _norm_q(want_body2)
                                             else "与原文不一致") if body2 else "：未找到该栏"))

    # ---------------- 26（-1）卡片页的白色底板与两处橙色装饰 ----------------
    # ①白色底板：白色填充、31.64cm×17.09cm（容差 0.5cm）、左右 3.3%–96.7%、
    #   上下 4.1%–93.9%（容差 1.5 个百分点）；
    # ②左上角图标：位于页面左上（水平中心 <20%、垂直中心 <25%）的 0.6–2.5cm 小图形；
    # ③右下角装饰：位于页面右下（水平中心 >75%、垂直中心 >70%）的 ≥2cm 图形。
    # 与办公软件对齐：底板在原稿中是"矩形"预设几何但视觉上带圆角，故只按
    # 白色填充 + 尺寸 + 位置认定，不强求 prstGeom=roundRect。
    board_bad, board_parts = False, []
    for pg in board_pages:
        has_board = has_icon = has_deco = False
        for sh, rect, _d in pg['shapes']:
            if rect is None:
                continue
            x, y, w, h = rect
            cx, cy = x + w / 2.0, y + h / 2.0
            if (_is_white_fill(sh) and abs(w - 31.64) <= 0.5 and abs(h - 17.09) <= 0.5
                    and abs(x / pw * 100.0 - 3.3) <= 1.5
                    and abs((x + w) / pw * 100.0 - 96.7) <= 1.5
                    and abs(y / ph * 100.0 - 4.1) <= 1.5
                    and abs((y + h) / ph * 100.0 - 93.9) <= 1.5):
                has_board = True
            if (0.6 <= w <= 2.5 and 0.6 <= h <= 2.5
                    and cx < pw * 0.20 and cy < ph * 0.25):
                has_icon = True
            if w >= 2.0 and h >= 2.0 and cx > pw * 0.75 and cy > ph * 0.70:
                has_deco = True
        if not (has_board and has_icon and has_deco):
            board_bad = True
        board_parts.append(f"第{pg['no']}页白色底板{'有' if has_board else '无'}、"
                           f"左上角图标{'有' if has_icon else '无'}、"
                           f"右下角装饰{'有' if has_deco else '无'}")
    _add(results, -1, board_bad,
         "承载5张老师卡片的页面上未出现31.64cm×17.09cm的白色圆角底板"
         "（位于页面从左至右3.3%–96.7%、从上至下4.1%–93.9%），或未出现左上角橙色菱形箭头图标与右下角橙色三角装饰",
         "；".join(board_parts) if board_parts else "全篇未找到承载老师卡片的页面，该项不适用")

    # ---------------- 27（-1）首页四处参照文字完整出现 ----------------
    # 英文串按忽略大小写比对（原稿用 cap="all" 让小写文本显示为全大写）。
    missing_cover = [t for t in COVER_TEXTS if t not in ref_shapes]
    _add(results, -1, bool(missing_cover),
         "首页未完整出现“2026”“公司简介”“品牌介绍手册”“SIMPLE ATMOSPHERE COMPANY PROFILE”四处文字",
         "四处均在" if not missing_cover else "缺失：" + "、".join(missing_cover))

    # ---------------- 28（-1）首页左侧装饰图形 ----------------
    # 逐类清点首页左半侧的装饰：椭圆（圆环轮廓）、三角形（下三角与渐变箭头群）、
    # 任意多边形（三角形色块、人形图标、右侧菱形等自由图形）。
    # 与办公软件对齐：这些装饰是原稿自带的组合，逐一比对形状名不稳定（复制后会
    # 重命名），故按"类别数量下限"判断是否被整体删除。
    ell = tri = free = 0
    for sh, rect, _d in cover['shapes']:
        if rect is None or (rect[0] + rect[2] / 2.0) >= pw / 2.0:
            continue
        p = _prst(sh)
        if p == 'ellipse':
            ell += 1
        elif p == 'triangle':
            tri += 1
        if sh.shape_type == 5:                     # MSO_SHAPE_TYPE.FREEFORM
            free += 1
    deco_bad = not (ell >= 2 and tri >= 8 and free >= 3)
    _add(results, -1, deco_bad,
         "首页左侧未出现橙色三角形色块、橙色渐变箭头群、两个圆环轮廓、顶部橙色下三角与右侧橙色菱形等装饰图形",
         f"左半侧实测椭圆{ell}个、三角形{tri}个、任意多边形{free}个")

    # ---------------- 29（-3）参考标注红框 ----------------
    red_pages = sorted({pg['no'] for pg in pages
                        for sh, _r, _d in pg['shapes'] if _red_line_frame(sh)})
    _add(results, -3, bool(red_pages),
         "任一页面出现红色矩形框线等参考标注图形（将参考示意图中仅用于标示位置的红框画入成品）",
         "全篇未出现" if not red_pages else "出现于第 " + "、".join(map(str, red_pages)) + " 页")

    # ---------------- 30（-3）原有结构页是否齐备 ----------------
    # "品牌概述"须有 2 页（章节页 + 内容页）；"01"在原稿中是目录页里的栏目编号，
    # 并非独立页面的标题，故允许以"页面上存在独立的『01』文字"认定其仍在。
    titles = [pg['title'] for pg in pages]
    has_01 = ("01" in titles) or any("01" in pg['texts'] for pg in pages)
    n_overview = titles.count("品牌概述")
    has_coop = "协作架构" in titles
    has_team = T_TEAM_SECTION in titles
    struct_bad = not (has_01 and n_overview >= 2 and has_coop and has_team)
    _add(results, -3, struct_bad,
         "全篇未同时出现标题为“01”“品牌概述”（2页）“协作架构”“团队介绍”的页面",
         f"“01”{'在' if has_01 else '缺'}、“品牌概述”{n_overview}页、"
         f"“协作架构”{'在' if has_coop else '缺'}、“团队介绍”{'在' if has_team else '缺'}")

    # ---------------- 31（-1）"品牌文化/价值观/使命/愿景"四页及其顺序 ----------------
    culture_seq = ("品牌文化", "品牌价值观", "品牌使命", "品牌愿景")
    nos = [next((pg['no'] for pg in pages if pg['title'] == t), None) for t in culture_seq]
    culture_bad = (any(n is None for n in nos)
                   or any(nos[k] >= nos[k + 1] for k in range(3)))
    _add(results, -1, culture_bad,
         "全篇未依次出现标题为“品牌文化”“品牌价值观”“品牌使命”“品牌愿景”的四页，或四页相互顺序不为该排列",
         "、".join(f"{t}第{n}页" if n else f"{t}缺失" for t, n in zip(culture_seq, nos)))

    # ---------------- 32（-1）水印 / 注释性文字 / 日期文本框 ----------------
    # 关键词命中即扣；日期另按"20xx年x月x日""20xx-x-x""20xx/x/x"以及
    # PowerPoint 日期占位符（a:fld type="datetime*"）识别。
    # 注意"2024年春季启动"这类原稿文案不含月日，不会被日期正则误伤。
    note_kws = ("修改说明", "动画说明", "由AI生成", "AI生成", "水印")
    date_re = re.compile(r'20\d{2}\s*年\s*\d{1,2}\s*月\s*\d{1,2}\s*日'
                         r'|20\d{2}[-/.]\d{1,2}[-/.]\d{1,2}')
    note_hits = []
    for pg in pages:
        hit = [k for k in note_kws if any(k in t for t in pg['texts'])]
        if any(date_re.search(t) for t in pg['texts']):
            hit.append("日期文字")
        for sh, _r, _d in pg['shapes']:
            if not getattr(sh, 'has_text_frame', False):
                continue
            for fld in sh.text_frame._txBody.iter(f'{A_NS}fld'):
                if (fld.get('type') or '').startswith('datetime'):
                    hit.append("日期占位符")
                    break
        if hit:
            note_hits.append(f"第{pg['no']}页{'/'.join(sorted(set(hit)))}")
    _add(results, -1, bool(note_hits),
         "任一页面出现水印文字，或出现“修改说明”“动画说明”“由AI生成”等注释性文字，或出现日期文本框",
         "全篇未出现" if not note_hits else "；".join(note_hits))

    # ---------------- 33（-1）文字超出文本框被截断 ----------------
    # 仅当文本框既未设"根据文字调整形状大小"（spAutoFit）也未设"溢出时缩排文字"
    # （normAutofit）时，才可能出现肉眼可见的截断；此时按字符估宽算出折行数与
    # 所需高度，超过框高 1.3 倍（留足行距与估算误差）才判截断。
    # 首页迁入的四段文字已由 05 项单独考核，此处按细则排除。
    exempt = {id(m['sh']) for m in mig.values()}
    trunc = []
    for pg in pages:
        for sh, rect, _d in pg['shapes']:
            if rect is None or not getattr(sh, 'has_text_frame', False) or id(sh) in exempt:
                continue
            if not _norm(sh.text_frame.text):
                continue
            bp = sh.text_frame._txBody.find(f'{A_NS}bodyPr')
            kinds = [c.tag.split('}')[-1] for c in (bp if bp is not None else [])]
            if 'spAutoFit' in kinds or 'normAutofit' in kinds:
                continue
            lins, rins = _body_ins(sh)
            avail = rect[2] - lins - rins
            if avail <= 0.1:
                continue
            need_h = 0.0
            for para in sh.text_frame.paragraphs:
                txt = _norm(para.text)
                if not txt:
                    continue
                runs = [r for r in para.runs if _norm(r.text)]
                size = max([_run_size(sh, para, r) for r in runs]) if runs else 18.0
                lines = (1 if _wrap_off(sh)
                         else max(1, int(math.ceil(_text_width_cm(txt, size) / avail - 1e-6))))
                need_h += lines * size * 1.2 * 0.0352778
            if need_h > rect[3] * 1.3 + 0.1:
                trunc.append(f"第{pg['no']}页“{_norm(sh.text_frame.text)[:10]}…”"
                             f"需{need_h:.2f}cm/框高{rect[3]:.2f}cm")
    _add(results, -1, bool(trunc),
         "除首页迁入的两张图片与四段文字、以及“教学团队介绍”页的5张卡片外，"
         "任一页面出现文字超出所属文本框可视区域而被截断、显示不全",
         "未发现截断" if not trunc else "；".join(trunc[:4]))

    # ========== 将 results 反解为结构化 dim2_items ==========
    # 逐项按追加顺序对齐 _MAX_DELTAS_ORDER：解析每条 "+N：说明" / "-N：说明" 的 N。
    # 加分项未命中写作 "+0："，减分项未命中写作 "-0："，两者的 delta 均为 0。
    if len(results) != len(_MAX_DELTAS_ORDER):
        # 顺序表与实际条目不一致：视为脚本内部错误，交给外层兜底
        raise RuntimeError(
            f"dim2 项数不一致：results={len(results)} vs order={len(_MAX_DELTAS_ORDER)}"
        )
    dim2_items = []
    total_score = 0
    for raw, max_delta in zip(results, _MAX_DELTAS_ORDER):
        head, _, rule = raw.partition("：")
        try:
            delta = int(head.replace("+", "").strip())
        except ValueError:
            delta = 0
        hit = delta != 0
        dim2_items.append({
            "rule": rule,
            "max_delta": max_delta,
            "delta": max_delta if hit else 0,
            "hit": hit,
        })
        total_score += (max_delta if hit else 0)

    result["dim2_items"] = dim2_items
    # 扣分项可能把总分压到负数；总分不低于 0。
    result["total_score"] = max(0, total_score)
    result["max_score"] = MAX_SCORE_TOTAL
    return result


def evaluate(dir_path):
    """按 §2.2 约定，返回结构化 dict；不 print、不改 sys.stdout、不 sys.exit。
    入参 dir_path 为"脚本所在目录的路径"，脚本自己在该目录里定位并打开被评估的
    .pptx（忽略以 ~$ 开头的 PowerPoint 临时锁文件）。

    约定：status='error' 表示这份交付根本没能进入评分（目录不存在、目录中无
    .pptx、脚本自身异常），此时 dim1_pass 为空串（维度一未判定）、dim1_reason 为
    空串、dim2_items 为空表、total_score 与 max_score 一律为 0。dim1_pass 为
    False 时 dim1_reason 必须带错误摘要；为 True 时 dim1_reason 为空串。"""
    result = {
        "id": SCRIPT_ID,
        "file_name": "",
        "status": "ok",
        "error": None,
        "dim1_pass": "",
        "dim1_reason": "",
        "dim2_items": [],
        "total_score": 0,
        "max_score": 0,
    }
    try:
        if not dir_path or not os.path.isdir(dir_path):
            result["status"] = "error"
            result["error"] = f"目录不存在：{dir_path}"
            return result
        cands = sorted(
            f for f in os.listdir(dir_path)
            if f.lower().endswith('.pptx') and not f.startswith('~$')
        )
        if not cands:
            result["status"] = "error"
            result["error"] = f"目录中未找到 .pptx 文件：{dir_path}"
            return result
        file_path = os.path.join(dir_path, cands[0])
        result["file_name"] = os.path.basename(file_path)
        return _evaluate_impl(file_path, result)
    except Exception as exc:
        # 脚本异常同样视为"没评上"：维度一未判定，两项分数归零。
        result["status"] = "error"
        result["error"] = f"{type(exc).__name__}: {exc}"
        result["dim1_pass"] = ""
        result["dim1_reason"] = ""
        result["dim2_items"] = []
        result["total_score"] = 0
        result["max_score"] = 0
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

















