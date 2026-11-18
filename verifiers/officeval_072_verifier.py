#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# pyright: reportUnknownMemberType=false, reportUnknownVariableType=false, reportUnknownArgumentType=false, reportUnknownParameterType=false, reportMissingParameterType=false, reportPrivateUsage=false, reportUnknownLambdaType=false, reportGeneralTypeIssues=false

import os, re, sys, json, math, zipfile, colorsys, traceback
from pptx import Presentation
from pptx.util import Emu

# 说明：按接口统一约定，本模块不修改全局 sys.stdout，也不 print 主结果。
# 主结果统一由 evaluate(dir_path) -> dict 返回；入参为“脚本所在目录的路径”，
# 由脚本自己在该目录内定位并打开被评估的 .pptx。本文件仅在 __main__ 下作为
# 本地调试入口，才把 dict 序列化为 JSON 打印。
#
# 与 008（Word）脚本的差异：本细则同时包含加分项与减分项，因此 _MAX_DELTAS_ORDER
# 中允许出现负值。负值项的语义为“该条描述的呈现效果必须原样保持”，命中（hit=True）
# 表示破坏已发生、按该条分值扣分；协议串相应写作 "-3：…"（命中）/ "-0：…"（未命中）。
#
# 判定策略（两种截然不同的实现路线都要能评）：
#   路线A「整页背景图 + 叠加文本框」：页面用 <p:bg> 图片或铺满页面的图片承载全部
#           装饰图形，页面上只有文本框。此时图形类要求视为“已按原样保留”，
#           图形类减分项不触发；文字类要求仍逐条检查。
#   路线B「全矢量重绘」：页面用形状重新画出装饰。此时逐条比对形状特征
#           （几何、颜色、旋转、数量、位置关系）。
# 文本折行/溢出统一用字宽模型估算：中日韩全角字符按 1.0 em、西文按 0.52 em、
# 行高按 字号*1.2*行距系数；带 <a:spAutoFit/> 的文本框由制作程序按文本高度
# 自动定高，不判溢出。

SCRIPT_ID = "072"
MAX_SCORE_TOTAL = 86     # 20 条加分项满分之和
MIN_SCORE_TOTAL = -71    # 29 条减分项扣分之和

_MAX_DELTAS_ORDER = [
    5,   # P01 12页/16:9/页序
    5,   # P02 全部文字可编辑
    5,   # P03 中文黑体类字体
    5,   # P04 第1页四处文字
    5,   # P05 第2-11页顶部长条/序号/标题/副标题/页码
    5,   # P06 第2页卡片与图标
    3,   # P07 第3页两张大卡与底部胶囊
    5,   # P08 第4页四角卡与中心圆环
    3,   # P09 第5页阶段卡与图标
    5,   # P10 第6页四等分圆环
    5,   # P11 第6页四张编号卡与运行原则
    3,   # P12 第7页顶部长条/岗位卡/齿轮/底部胶囊
    5,   # P13 第8页四卡环绕中心圆
    5,   # P14 第9页六边形蜂巢
    3,   # P15 第9页右侧三卡与底部单行
    3,   # P16 第10页左上大卡/左下两卡/右侧两卡
    5,   # P17 第11页三处顶部图标
    3,   # P18 第11页卡内文字与底部胶囊
    3,   # P19 第12页结束页文字
    5,   # P20 背景装饰逐页保留
    -3,  # N01 第1页笑脸五角星
    -3,  # N02 第1页黄色半圆山丘与白条纹
    -3,  # N03 第1页白色弧形波浪线
    -1,  # N04 第1页主标题紧凑排列
    -1,  # N05 第1页右下角被裁切的青绿丘形
    -3,  # N06 第2页星轨图案
    -1,  # N07 第3页浅色底条目单行
    -3,  # N08 第5页白圆＋彩环图标
    -1,  # N09 第5页底部单行
    -3,  # N10 无描边纯白圆角卡
    -1,  # N11 第2-11页页码单行横排
    -5,  # N12 第4页四角卡说明不溢出
    -3,  # N13 第4页卡标题与中心圆文字单行
    -5,  # N14 第6页四等分圆环未退化
    -3,  # N15 第4、7、11页底部胶囊形横条
    -3,  # N16 第8页四段橙黄弧形箭头
    -3,  # N17 第8页卡片说明不溢出
    -5,  # N18 第9页六边形紧邻蜂巢排布
    -3,  # N19 第9页六边形标签横排单行
    -3,  # N20 第9页右侧卡片说明不溢出
    -3,  # N21 第11页三处顶部图标造型
    -3,  # N22 第11页卡内说明不溢出
    -3,  # N23 第12页上下两组星座连线
    -1,  # N24 第2页带班信念文案
    -1,  # N25 第6页能量积分文案
    -1,  # N26 第7页底部文案
    -1,  # N27 第9页搭建方式文案
    -1,  # N28 第10页顶部副标题文案
    -1,  # N29 第11页第三张卡影响文案
]
assert len(_MAX_DELTAS_ORDER) == 49
assert sum(d for d in _MAX_DELTAS_ORDER if d > 0) == MAX_SCORE_TOTAL
assert sum(d for d in _MAX_DELTAS_ORDER if d < 0) == MIN_SCORE_TOTAL

# 与《评测细则.txt》维度2逐条一一对应的细则原文（去掉行首分值前缀）。
_RULES = [
    # ---- 加分项 P01-P20 ----
    "交付演示文稿共12页，页面为16:9宽屏（33.5-34.2cm×18.8-19.3cm / 12192000×6858000EMU），页序依次为封面页、01班级初印象、02精准画像、03核心理念、04发展目标、05创意实践、06第一步、07第二步、08第三步、09第四步、10特色成效、结束页",
    "12页中的全部文字（含标题、副标题、卡片标题与说明、图形内标签、编号、页码）均为可单独点选、双击即可改字的文本框或形状文字，无任何一页把成段文字并入图片",
    "中文文字在编辑与放映时显示为黑体类无衬线字体（微软雅黑、思源黑体等），即主题东亚字体（a:ea）或各段落东亚字体已显式指定为黑体类字体",
    "第1页含“以光为航 · 以心育人”主标题（深墨绿色#15424C-#15505A、加粗，置于白底描橙金细边的圆角框内）、“晨河班成长共创方案”副标题（白色加粗字，置于橙黄色#EBAE46-#F0B13E圆角胶囊内）、“班级文化与学生发展路径”与“让安静的努力被看见，让微小的改变被珍惜”两行底部文字（均为白色居中，上行加粗且字号明显大于下行），四处文字齐全、分行方式一致",
    "第2至第11页每页顶部有一条白底描橙金细边的圆角长条，条内含深墨绿圆角序号标签（依次为01-10）、其右侧的黑色加粗页标题、右端的灰色小字副标题（依次为“从观察到行动”“温柔激活与持续托举”“班级运行的底层逻辑”“从看见到自我驱动”“把理念转成班级日常”“让“我能做”变得具体”“把积极行为转化为正向反馈”“把单点光芒连成群星”“把阶段成果变成真实体验”“看见改变，沉淀习惯”）；每页右下角有两位数页码（依次为02-11），单行横排、字号11-13磅",
    "第2页左半为白色大卡（内含深墨绿色缩略图块、深青加粗“观察印象”、两行灰色说明），右半自上而下为“现状扫描”“带班信念”“核心目标”三张白卡，每张含左侧图标（深青圆形内白色叶形、橙黄心形、绿色五角星）、深青加粗标题与两行灰色说明",
    "第3页左右两张白色大卡为深青加粗标题“待唤醒的星火”“最坚实的土壤”，各含3条浅色底条目（左卡浅青底、以“学习：”“纪律：”“心理：”起首；右卡浅黄底、以“性格：”“氛围：”“潜能：”起首），每条单行完整显示；底部深墨绿圆角胶囊内为白色加粗单行“破局关键：把“被动安静”转化为“愿意尝试”，把“偶然发光”沉淀为“持续成长”。”",
    "第4页四角各有一张白底描橙金细边的圆角卡（发现微光、及时回应、伙伴互助、持续生长），每张为左侧图标＋深青加粗标题＋两行灰色说明；中心为深青实心圆＋一圈橙金细环＋两圈绿色细环，并有四条细青线自中心分别连向四角卡片",
    "第5页三张白卡沿左下至右上阶梯排列，各配白色圆形＋彩色圆环图标（绿环内树形、深青环内叶形、橙环内五角星），卡内为彩色小字阶段名（绿“起步阶段”、青“连接阶段”、橙“跃迁阶段”）＋深色加粗名称（看见微光/形成光束/自我点亮）＋两行灰色说明，相邻卡片间有青金双色斜条相连；底部为深色加粗单行“目标不是让每个孩子立刻耀眼，而是让他们逐渐拥有主动生长的力量。”",
    "第6页中部为四等分圆环，扇区颜色为上红、右深青、左绿、下橙黄，环心为带凹口的白色圆形并内含“成长的”与深青加粗“内驱力”两行，红色扇区内有白色心形、橙色扇区内有白色五角星，四个扇区分别标注“心愿花园”“责任剧场”“伙伴联盟”“能量积分”",
    "第6页左侧四张白卡依次为数字标签1（深青）、2（橙黄）、3（绿）、4（红）＋深色加粗标题“一、责任剧场”“二、能量积分”“三、伙伴联盟”“四、心愿花园”＋灰色单行说明“把班级事务改造成可参与的小角色”“把正向行为转化为持续反馈”“让互助从偶然变成一种结构”“用阶段成果承接情感期待”；右侧白底描橙金边卡内为深色加粗“运行原则”＋自上而下四个白字圆角胶囊：“低门槛参与”（深青）、“高频率反馈”（橙黄）、“小步子迭代”（绿）、“真实感表达”（红）",
    "第7页顶部白底描橙金边圆角长条内为深色加粗单行“把班级日常拆成若干小岗位：每一次认真完成，都会成为被记录的光点。”；右侧三张白卡为深青加粗标题“岗位认领”“微光观察员”“自我提醒官”＋灰色单行说明，左侧配深青与绿色齿轮图标及细连接线；底部深青圆角胶囊内为白色加粗“成效期待：学生在小任务中积累可感知的责任感，班级运行更顺畅。”",
    "第8页四张卡片（描边依次为深青“获得光点”、红“复盘升级”、橙黄“存入光册”、绿“兑换体验”，卡内各配同色图标）环绕中心深青实心圆排布，圆内为白色“能量”“循环”两行；底部橙色圆角胶囊内为白色加粗“规则底色：积分不是排名工具，而是帮助学生看见“我正在变好”。”",
    "第9页左侧为六色平顶正六边形蜂巢（红“支持”、蓝“创意”、深青“表达”、橙黄“整理”、绿“观察”、紫“记录”）环绕中心白色圆形（外描橙金环，内含橙黄五角星与深青“同盟”），六个白色标签均为横排单行",
    "第9页右侧三张白卡为深青加粗标题“搭建方式”“核心机制”“温柔约定”＋两行灰色说明；底部为深色加粗单行“合作的目的不是把每个人变成一样，而是让不同的光互相照亮。”",
    "第10页左上为深墨绿圆角大卡，卡内上部横排五枚彩色五角星（橙黄、青、红、橙黄、青）、下部居中为浅黄色加粗“心愿花园”；左下两张白卡为彩色五角星＋深青加粗“我敢开口了”“我能坚持了”＋灰色单行说明；右侧两张白卡分别描橙金边与深青边，为深青加粗“升级方式”“情感闭环”＋两行灰色说明",
    "第11页三张白卡的顶部图标依次为橙黄圆形内嵌白色喇叭（圆形右侧带缺口）、绿色中空圆角方框（只有描边、内部不填色）、深青实心圆形内嵌白色心形，卡片描边依次为橙金、绿、深青",
    "第11页三张卡内为深青加粗居中标题“自信在发声”“习惯在迁移”“氛围在变暖”＋“变化：…”“影响：…”两段灰色说明；底部深青圆角胶囊内为白色加粗“班级建设的成果，最终会落在孩子的眼神、语言和行动里。”",
    "第12页白底描橙金边圆角框内为深墨绿加粗“教育是一场温柔的点亮”＋“这份答案沉默而坚定，”“正被一颗颗愿意尝试的心慢慢写出来。”两行深色加粗文字，且白色大圆下部有浅黄色（#FFE591系）“感谢聆听与同行”",
    "背景装饰逐页对应保留——第2至第11页为统一浅色底＋左上淡绿色块、右上青绿色大圆、左侧橙黄大圆、右下绿色大圆、中部白色大圆及散落橙黄小圆点；第1、12页为深墨绿夜空底＋左上黄色大圆＋中部白色大圆＋右下青绿色圆＋散落的白色与浅黄星点",
    # ---- 减分项 N01-N29（描述的是必须原样保持的呈现效果，未满足即扣分）----
    "第1页的黄色五角星不满足为带两只深色圆眼睛与一条深色弧形笑嘴的完整笑脸造型，或页面上出现游离于星形之外的单独弧线/符号文本",
    "第1页五角星下方不满足有一座黄色半圆形山丘，且丘面上有多条白色横向条纹",
    "第1页底部深青色水面不满足为约9-10条等距平行、两端上翘的白色弧形波浪线（不得改为向一侧汇聚于同一点的直线束）",
    "第1页主标题不满足为“以光为航 · 以心育人”的紧凑排列（间隔号两侧不得插入多余空格使字距明显拉开）",
    "第1页右下角不满足为一块被页面右缘裁切的青绿色丘形（不得改为完整的圆形）",
    "第2页左侧大卡内的深墨绿色缩略图块不满足为多条浅青色同心椭圆弧线＋散落黄色圆点组成的星轨图案（不得简化为单条椭圆轮廓）",
    "第3页两张大卡内的6条浅色底条目不满足每条单行完整显示（不得折行越出所在的浅色底条）",
    "第5页三处阶段图标不满足为完整的白色圆形＋外圈彩色圆环（依次为绿、深青、橙金）包裹（不得出现圆环残缺、断口或丢失外环）",
    "第5页底部“目标不是让每个孩子立刻耀眼，而是让他们逐渐拥有主动生长的力量。”不满足单行显示（不得折成两行）",
    "第2页左侧大卡与右侧三张卡片、第5页三张阶段卡、第6页左侧四张编号卡、第7页三张岗位卡、第9页右侧三张卡不满足为无描边的纯白圆角卡（仅带柔和投影，不得为其添加橙金/橙黄色描边）",
    "第2至第11页右下角的两位数页码不满足单行横排显示（不得因文本框过窄而竖向折成上下两行，如“0”“6”分列两行）",
    "第4页四角卡片的两行灰色说明不满足完整位于卡片白色区域之内（不得溢出卡片下沿，压在背景圆环与连接线上）",
    "第4页四角卡片标题不满足单行完整显示（不得折行为“发现微”“光”两行），中心深青圆内的“微光汇聚”“向阳生长”不满足各自单行显示于圆形范围之内（不得折成四行并越出圆形）",
    "第6页中部不满足为红、深青、绿、橙黄四等分圆环（不得退化为绿色方块＋橙色扇形色块并丢失其余扇区）",
    "第4、7、11页底部的横条不满足为两端全圆角的胶囊形（不得改为直角矩形）",
    "第8页四张卡片之间不满足有四段橙黄色弧形箭头构成顺时针循环指向（不得替换为深色直线段并丢失箭头）",
    "第8页四张卡片的灰色说明不满足完整位于卡片内部（不得越出卡片下边界）",
    "第9页六个彩色六边形不满足为平顶朝上、彼此紧邻的蜂巢排布（不得旋转为尖顶朝上并在六边形之间留出缝隙）",
    "第9页六边形内的“支持”“创意”“表达”“整理”“观察”“记录”不满足横排单行且完整可见（不得逐字竖排换行，“表达”不得被中心白色圆形遮挡）",
    "第9页右侧三张卡片的两行说明不满足完整位于卡片内部（不得末行越出卡片下边界落到卡外）",
    "第11页三处顶部图标不满足为橙黄圆形内嵌白色喇叭、绿色中空圆角方框、深青实心圆内嵌白色心形（不得简化为圆形内白色小三角、绿色实心圆内白色方块、无圆形底的单个心形）",
    "第11页三张卡内“变化：”“影响：”两段说明不满足完整位于卡片内部（不得折行溢出卡片下沿）",
    "第12页白色大圆内不满足有上、下两组橙黄色星座连线（上方为多颗五角星连成的折线、下方为多颗五角星连成的锯齿折线，不得只保留其中一组）",
    "第2页“带班信念”说明不满足为“每个孩子都有独特光点，教育要做的是帮他们找到被看见的入口。”",
    "第6页“二、能量积分”的说明不满足为“把正向行为转化为持续反馈”",
    "第7页底部文字不满足为“成效期待：学生在小任务中积累可感知的责任感，班级运行更顺畅。”（不得漏掉句末句号）",
    "第9页“搭建方式”说明不满足为“根据性格与能力进行混合搭配，让每个小队都拥有表达者、倾听者、组织者和创意者。”",
    "第10页顶部右端的灰色副标题不满足为“把阶段成果变成真实体验”",
    "第11页第三张卡的“影响”不满足为“互助关系更柔和，团队归属感增强。”",
]
assert len(_RULES) == 49

# ----------------------------------------------------------------- 基础常量
NS_A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
NS_P = "{http://schemas.openxmlformats.org/presentationml/2006/main}"
EMU_CM = 360000.0
PT_CM = 0.03528                 # 1磅 ≈ 0.03528cm
DEF_SZ = 18.0                   # presentation.xml defaultTextStyle lvl1 sz=1800
DEF_LINS = 0.254                # bodyPr 默认 lIns/rIns = 91440EMU
DEF_TINS = 0.127                # bodyPr 默认 tIns/bIns = 45720EMU
OVF_TOL = 0.35                  # 溢出容差（按行高倍数）

GOLD = ("orange", "yellow")     # 橙金 / 橙黄 / 浅黄
TEAL = ("teal",)                # 深青 / 深墨绿 / 浅青
GREENS = ("green",)

_SANS_EA = ("微软雅黑", "雅黑", "黑体", "思源黑", "苹方", "等线", "普惠", "鸿蒙",
            "microsoft yahei", "msyh", "simhei", "heiti", "source han sans",
            "sourcehansans", "noto sans", "pingfang", "dengxian", "harmonyos sans",
            "alibaba puhuiti", "yahei", "hiragino sans", "sans-serif", "sans")
_SERIF_EA = ("宋体", "simsun", "楷体", "kaiti", "仿宋", "fangsong", "隶书", "华文中宋",
             "times", "georgia", "serif", "songti", "mingliu", "ming", "batang")

_PUNCT = "，。、；：？！“”‘’（）《》〈〉【】—…·,.;:?!\"'()<>[]{}~-_＋+＝= 　"


def _cm(v):
    return float(v or 0) / EMU_CM


def _norm(s):
    return re.sub(r"\s+", "", s or "")


def _key(s):
    """去掉空白与全部标点后的比较键，用于容忍标点写法差异。"""
    return "".join(ch for ch in _norm(s) if ch not in _PUNCT)


def _i(v, dflt=None):
    try:
        return int(v)
    except (TypeError, ValueError):
        return dflt

# ----------------------------------------------------------------- 颜色工具
def _rgb(hx):
    hx = (hx or "").strip().lstrip("#")
    if len(hx) != 6:
        return None
    try:
        return int(hx[0:2], 16), int(hx[2:4], 16), int(hx[4:6], 16)
    except ValueError:
        return None


def _hsl(hx):
    t = _rgb(hx)
    if t is None:
        return None
    h, l, s = colorsys.rgb_to_hls(t[0] / 255.0, t[1] / 255.0, t[2] / 255.0)
    return h * 360.0, s, l


def _fam(hx):
    """把颜色归入色系；细则中的“深青/深墨绿/浅青”统一落到 teal。"""
    v = _hsl(hx)
    if v is None:
        return None
    H, S, L = v
    if S < 0.12 or L > 0.955:
        if L > 0.90:
            return "white"
        if L < 0.18:
            return "black"
        return "gray"
    if L < 0.12:
        return "black"
    if H < 15 or H >= 345:
        return "red"
    if H < 45:
        return "orange"
    if H < 72:
        return "yellow"
    if H < 160:
        return "green"
    if H < 205:
        return "teal"
    if H < 258:
        return "blue"
    if H < 300:
        return "purple"
    return "pink"


def _is_white(hx):
    v = _hsl(hx)
    return v is not None and v[2] > 0.90 and v[1] < 0.30

def _is_pale(hx):
    """浅色（如浅黄 #FFE591/#FEE695、浅青 #62C2BD 系）：明度偏高即可。"""
    v = _hsl(hx)
    return v is not None and v[2] > 0.75


def _is_dark(hx):
    v = _hsl(hx)
    return v is not None and v[2] < 0.45


def _apply_mods(hx, node):
    """按 lumMod/lumOff/shade/tint/satMod 近似修正颜色。"""
    v = _hsl(hx)
    if v is None or node is None:
        return hx
    H, S, L = v

    def g(tag):
        e = node.find(NS_A + tag)
        if e is None:
            return None
        n = _i(e.get("val"))
        return None if n is None else n / 100000.0

    lm, lo, sh, ti, sm = g("lumMod"), g("lumOff"), g("shade"), g("tint"), g("satMod")
    if lm is not None:
        L *= lm
    if lo is not None:
        L += lo
    if sm is not None:
        S *= sm
    if sh is not None:
        L *= sh
    if ti is not None:
        L = L * ti + (1.0 - ti)
    L = max(0.0, min(1.0, L))
    S = max(0.0, min(1.0, S))
    r, g2, b = colorsys.hls_to_rgb(H / 360.0, L, S)
    return "%02X%02X%02X" % (int(round(r * 255)), int(round(g2 * 255)), int(round(b * 255)))


def _scheme_hex(name, ctx):
    if not name or name == "phClr":
        return None
    return ctx["scheme"].get(ctx["clrmap"].get(name, name))

def _solid_hex(node, ctx):
    """node 为任意含颜色子元素的容器（solidFill / gs / rPr 等）。"""
    if node is None:
        return None
    s = node.find(NS_A + "srgbClr")
    if s is not None:
        return _apply_mods((s.get("val") or "").upper(), s)
    s = node.find(NS_A + "schemeClr")
    if s is not None:
        hx = _scheme_hex(s.get("val"), ctx)
        return _apply_mods(hx, s) if hx else None
    s = node.find(NS_A + "sysClr")
    if s is not None:
        return (s.get("lastClr") or "").upper() or None
    return None


def _fill_of(spPr, style, ctx):
    """返回 16 进制色 / 'none' / 'blip' / 'grad' / None（继承且无法判定）。"""
    if spPr is not None:
        if spPr.find(NS_A + "noFill") is not None:
            return "none"
        sf = spPr.find(NS_A + "solidFill")
        if sf is not None:
            return _solid_hex(sf, ctx)
        gf = spPr.find(NS_A + "gradFill")
        if gf is not None:
            gs = gf.find(NS_A + "gsLst/" + NS_A + "gs")
            return _solid_hex(gs, ctx) or "grad"
        if spPr.find(NS_A + "blipFill") is not None:
            return "blip"
        if spPr.find(NS_A + "pattFill") is not None:
            return "patt"
    if style is not None:
        fr = style.find(NS_A + "fillRef")
        if fr is not None:
            if fr.get("idx") == "0":
                return "none"
            return _solid_hex(fr, ctx)
    return None


def _line_of(spPr, style, ctx):
    """返回 (描边色 / 'none' / None, 线宽cm)。"""
    if spPr is not None:
        ln = spPr.find(NS_A + "ln")
        if ln is not None:
            w = _cm(_i(ln.get("w"), 0))
            if ln.find(NS_A + "noFill") is not None:
                return "none", w
            c = _solid_hex(ln.find(NS_A + "solidFill"), ctx)
            if c:
                return c, w
            if ln.find(NS_A + "gradFill") is not None:
                return "grad", w
    if style is not None:
        lr = style.find(NS_A + "lnRef")
        if lr is not None and lr.get("idx") != "0":
            return _solid_hex(lr, ctx), 0.0
    return None, 0.0

# ------------------------------------------------------------- 字宽/折行模型
_WIDE_RANGES = ((0x1100, 0x115F), (0x2E80, 0x303E), (0x3041, 0x33FF),
                (0x3400, 0x4DBF), (0x4E00, 0x9FFF), (0xA000, 0xA4CF),
                (0xAC00, 0xD7A3), (0xF900, 0xFAFF), (0xFE30, 0xFE4F),
                (0xFF00, 0xFF60), (0xFFE0, 0xFFE6))
_WIDE_EXTRA = (0x00B7, 0x2013, 0x2014, 0x2018, 0x2019, 0x201C, 0x201D, 0x2026, 0x2030)


def _adv(ch):
    """单字符推进量（em）：中日韩全角 1.0，西文 0.52。"""
    o = ord(ch)
    if o in (0x09, 0x0A, 0x0B, 0x0D):
        return 0.0
    if o in _WIDE_EXTRA:
        return 1.0
    for a, b in _WIDE_RANGES:
        if a <= o <= b:
            return 1.0
    return 0.52


def _tw(s, pt):
    return sum(_adv(c) for c in s) * (pt or DEF_SZ) * PT_CM


def _seg_lines(seg, pt, avail):
    if not seg:
        return 1
    return max(1, int(math.ceil(_tw(seg, pt) / max(avail, 0.08) - 1e-6)))


def _avail(d, width=None):
    w = (d["W"] if width is None else width) - d["ins"][0] - d["ins"][1]
    return max(w, 0.08)


def _nlines(d, width=None):
    av = _avail(d, width)
    n = 0
    for p in d["paras"]:
        for seg in re.split(r"[\n\x0b]", p["text"]):
            n += _seg_lines(seg, p["sz"], av)
    return n


def _nsegs(d):
    return sum(len(re.split(r"[\n\x0b]", p["text"])) for p in d["paras"])


def _wraps(d, width=None):
    """文本是否发生了非人工的折行（实际行数多于显式分行数）。"""
    return _nlines(d, width) > _nsegs(d)

def _est_h(d, width=None):
    av = _avail(d, width)
    h = d["ins"][2] + d["ins"][3]
    for p in d["paras"]:
        pt = p["sz"] or DEF_SZ
        n = 0
        for seg in re.split(r"[\n\x0b]", p["text"]):
            n += _seg_lines(seg, pt, av)
        h += n * pt * 1.2 * (p["lnspc"] or 1.0) * PT_CM
    return h


def _lh(d):
    pt = (d["paras"][0]["sz"] if d["paras"] else DEF_SZ) or DEF_SZ
    return pt * 1.2 * PT_CM


def _ovf(d, card=None):
    """文本是否溢出自身文本框（或所在白卡）。

    spAutoFit 由制作程序按文本定高，故不判文本框自身；但白卡尺寸是固定的，
    文本框自动增高只会把文字推到卡外，因此卡片边界仍需判定。
    """
    eh = _est_h(d)
    if d["autofit"] != "sp" and eh > d["H"] + OVF_TOL * _lh(d):
        return True
    if card is not None and d["T"] + max(d["H"], eh) > card["T"] + card["H"] - 0.05:
        return True
    return False

# ------------------------------------------------------------- 文本与形状解析
def _fontnames(rPr, dfr):
    """取（东亚字体, 西文字体）；两级：run 级 rPr 优先，其次段落 defRPr。"""
    ea = lat = None
    for src in (rPr, dfr):
        if src is None:
            continue
        if ea is None:
            e = src.find(NS_A + "ea")
            if e is not None and (e.get("typeface") or "").strip():
                ea = e.get("typeface").strip()
        if lat is None:
            e = src.find(NS_A + "latin")
            if e is not None and (e.get("typeface") or "").strip():
                lat = e.get("typeface").strip()
    return ea, lat


def _paras(el, ctx):
    tx = el.find(NS_P + "txBody")
    if tx is None:
        tx = el.find(".//" + NS_P + "txBody")
    if tx is None:
        return [], {}
    body = tx.find(NS_A + "bodyPr")
    out = []
    for p in tx.findall(NS_A + "p"):
        pPr = p.find(NS_A + "pPr")
        dfr = pPr.find(NS_A + "defRPr") if pPr is not None else None
        lnspc = 1.0
        algn = None
        if pPr is not None:
            algn = pPr.get("algn")
            sp = pPr.find(NS_A + "lnSpc/" + NS_A + "spcPct")
            if sp is not None:
                lnspc = (_i(sp.get("val"), 100000) or 100000) / 100000.0
        txt = []
        sz = b = col = ea = lat = None
        for node in p:
            if node.tag in (NS_A + "r", NS_A + "fld"):
                rPr = node.find(NS_A + "rPr")
                t = node.find(NS_A + "t")
                txt.append((t.text or "") if t is not None else "")
                if sz is None:
                    sz = _i((rPr.get("sz") if rPr is not None else None) or
                            (dfr.get("sz") if dfr is not None else None))
                    sz = None if sz is None else sz / 100.0
                if b is None:
                    bv = (rPr.get("b") if rPr is not None else None)
                    if bv is None and dfr is not None:
                        bv = dfr.get("b")
                    b = None if bv is None else (bv in ("1", "true"))
                if col is None:
                    col = _solid_hex(rPr.find(NS_A + "solidFill") if rPr is not None else None, ctx)
                    if col is None and dfr is not None:
                        col = _solid_hex(dfr.find(NS_A + "solidFill"), ctx)
                if ea is None and lat is None:
                    ea, lat = _fontnames(rPr, dfr)
            elif node.tag == NS_A + "br":
                txt.append("\x0b")
        out.append({"text": "".join(txt), "sz": sz, "b": b, "color": col,
                    "algn": algn, "lnspc": lnspc, "ea": ea, "latin": lat})
    return out, _bodypr(body)

def _bodypr(body):
    d = {"autofit": "none", "vert": None, "wrap": None,
         "ins": (DEF_LINS, DEF_LINS, DEF_TINS, DEF_TINS)}
    if body is None:
        return d
    if body.find(NS_A + "spAutoFit") is not None:
        d["autofit"] = "sp"
    elif body.find(NS_A + "normAutofit") is not None:
        d["autofit"] = "norm"
        na = body.find(NS_A + "normAutofit")
        d["fontScale"] = (_i(na.get("fontScale")) or 100000) / 100000.0
    d["vert"] = body.get("vert")
    d["wrap"] = body.get("wrap")
    ins = []
    for k, dv in (("lIns", DEF_LINS), ("rIns", DEF_LINS), ("tIns", DEF_TINS), ("bIns", DEF_TINS)):
        v = _i(body.get(k))
        ins.append(dv if v is None else _cm(v))
    d["ins"] = tuple(ins)
    return d


def _shape_info(el, ox, oy, sx, sy, ctx):
    spPr = el.find(NS_P + "spPr")
    if spPr is None:
        spPr = el.find(NS_P + "grpSpPr")
    style = el.find(NS_P + "style")
    xf = spPr.find(NS_A + "xfrm") if spPr is not None else None
    if xf is None:
        xf = el.find(".//" + NS_A + "xfrm")
    L = T = W = H = 0.0
    rot = 0.0
    fh = fv = False
    if xf is not None:
        off, ext = xf.find(NS_A + "off"), xf.find(NS_A + "ext")
        if off is not None:
            L = ox + sx * float(_i(off.get("x"), 0) or 0)
            T = oy + sy * float(_i(off.get("y"), 0) or 0)
        if ext is not None:
            W = sx * float(_i(ext.get("cx"), 0) or 0)
            H = sy * float(_i(ext.get("cy"), 0) or 0)
        rot = (_i(xf.get("rot"), 0) or 0) / 60000.0
        fh, fv = xf.get("flipH") == "1", xf.get("flipV") == "1"
    geom, adj = None, {}
    if spPr is not None:
        g = spPr.find(NS_A + "prstGeom")
        if g is not None:
            geom = g.get("prst")
            for gd in g.findall(NS_A + "avLst/" + NS_A + "gd"):
                adj[gd.get("name")] = _i((gd.get("fmla") or "").replace("val ", "").strip())
        elif spPr.find(NS_A + "custGeom") is not None:
            geom = "custom"
    paras, bp = _paras(el, ctx)
    fill = _fill_of(spPr, style, ctx)
    line, lnw = _line_of(spPr, style, ctx)
    d = {"tag": el.tag.split("}")[-1], "geom": geom, "adj": adj,
         "fill": fill, "line": line, "lnw": lnw,
         "L": L / EMU_CM, "T": T / EMU_CM, "W": W / EMU_CM, "H": H / EMU_CM,
         "rot": rot % 360.0, "flipH": fh, "flipV": fv,
         "paras": paras, "text": "\n".join(p["text"] for p in paras)}
    d.update(bp)
    d["cx"], d["cy"] = d["L"] + d["W"] / 2.0, d["T"] + d["H"] / 2.0
    return d

def _walk(tree, ctx, ox=0.0, oy=0.0, sx=1.0, sy=1.0, out=None):
    if out is None:
        out = []
    for el in tree:
        if el.tag == NS_P + "grpSp":
            nox, noy, nsx, nsy = ox, oy, sx, sy
            gx = el.find(NS_P + "grpSpPr/" + NS_A + "xfrm")
            if gx is not None:
                off, ext = gx.find(NS_A + "off"), gx.find(NS_A + "ext")
                cho, che = gx.find(NS_A + "chOff"), gx.find(NS_A + "chExt")
                if None not in (off, ext, cho, che):
                    cw = float(_i(che.get("cx"), 1) or 1) or 1.0
                    ch = float(_i(che.get("cy"), 1) or 1) or 1.0
                    nsx = sx * (float(_i(ext.get("cx"), 0) or 0) / cw)
                    nsy = sy * (float(_i(ext.get("cy"), 0) or 0) / ch)
                    nox = ox + sx * float(_i(off.get("x"), 0) or 0) - nsx * float(_i(cho.get("x"), 0) or 0)
                    noy = oy + sy * float(_i(off.get("y"), 0) or 0) - nsy * float(_i(cho.get("y"), 0) or 0)
            _walk(el, ctx, nox, noy, nsx, nsy, out)
        elif el.tag in (NS_P + "sp", NS_P + "pic", NS_P + "cxnSp", NS_P + "graphicFrame"):
            out.append(_shape_info(el, ox, oy, sx, sy, ctx))
    return out


def _load_ctx(path):
    ctx = {"scheme": {}, "clrmap": {}, "ea": "", "hans": "", "media": [], "bg": set()}
    try:
        zf = zipfile.ZipFile(path)
    except Exception:
        return ctx
    try:
        names = zf.namelist()
        ctx["media"] = [n for n in names if n.startswith("ppt/media/")]
        tn = sorted(n for n in names if re.match(r"ppt/theme/theme\d+\.xml$", n))
        if tn:
            t = zf.read(tn[0]).decode("utf-8", "ignore")
            m = re.search(r"<a:clrScheme.*?</a:clrScheme>", t, re.S)
            if m:
                for name, body in re.findall(r"<a:(\w+)>(.*?)</a:\1>", m.group(0), re.S):
                    mm = (re.search(r'srgbClr val="([0-9A-Fa-f]{6})"', body) or
                          re.search(r'lastClr="([0-9A-Fa-f]{6})"', body))
                    if mm:
                        ctx["scheme"][name] = mm.group(1).upper()
            m = re.search(r"<a:minorFont>(.*?)</a:minorFont>", t, re.S)
            if m:
                mm = re.search(r'<a:ea typeface="([^"]*)"', m.group(1))
                ctx["ea"] = (mm.group(1) if mm else "").strip()
                mm = re.search(r'<a:font script="Hans" typeface="([^"]*)"', m.group(1))
                ctx["hans"] = (mm.group(1) if mm else "").strip()
        mn = sorted(n for n in names if re.match(r"ppt/slideMasters/slideMaster\d+\.xml$", n))
        if mn:
            mm = re.search(r"<p:clrMap([^>]*)/>", zf.read(mn[0]).decode("utf-8", "ignore"))
            if mm:
                for k, v in re.findall(r'(\w+)="(\w+)"', mm.group(1)):
                    ctx["clrmap"][k] = v
    finally:
        zf.close()
    return ctx

# --------------------------------------------------------------- 文档模型构建
def _build(path):
    prs = Presentation(path)
    ctx = _load_ctx(path)
    D = {"path": path, "ctx": ctx, "pw": _cm(prs.slide_width), "ph": _cm(prs.slide_height),
         "slides": []}
    for i, s in enumerate(prs.slides, 1):
        shapes = _walk(s.shapes._spTree, ctx)
        bg = s.element.find(NS_P + "cSld/" + NS_P + "bg")
        bgpic = bg is not None and bg.find(".//" + NS_A + "blip") is not None
        pics = [d for d in shapes if d["tag"] == "pic"]
        fullpic = any(d["W"] >= 0.90 * D["pw"] and d["H"] >= 0.90 * D["ph"] for d in pics)
        txt = [d for d in shapes if _norm(d["text"])]
        allt = "\n".join(d["text"] for d in txt)
        D["slides"].append({
            "i": i, "shapes": shapes, "txt": txt, "pics": pics,
            "bgimg": bool(bgpic or fullpic), "vec": not (bgpic or fullpic),
            "text": allt, "ntext": _norm(allt), "ktext": _key(allt),
        })
    return D


def _S(D, i):
    """取第 i 页（1 基）；越界返回空页模型，避免越界异常。"""
    if 1 <= i <= len(D["slides"]):
        return D["slides"][i - 1]
    return {"i": i, "shapes": [], "txt": [], "pics": [], "bgimg": False, "vec": False,
            "text": "", "ntext": "", "ktext": ""}


def _has(D, i, sub):
    return _norm(sub) in _S(D, i)["ntext"]


def _hask(D, i, sub):
    return _key(sub) in _S(D, i)["ktext"]


def _find(D, i, sub):
    n = _norm(sub)
    return [d for d in _S(D, i)["txt"] if n in _norm(d["text"])]


def _one(D, i, sub):
    r = _find(D, i, sub)
    return r[0] if r else None

def _sel(sl, geom=None, fill=None, line=None, nofill=False,
         minw=0.0, minh=0.0, maxw=1e9, maxh=1e9, cx=None, cy=None, text=None):
    """按几何/填充色系/描边色系/尺寸/中心位置筛形状。"""
    out = []
    for d in sl["shapes"]:
        if geom is not None and d["geom"] not in geom:
            continue
        if fill is not None and _fam(d["fill"]) not in fill:
            continue
        if nofill and d["fill"] not in ("none", None):
            continue
        if line is not None and _fam(d["line"]) not in line:
            continue
        if not (minw <= d["W"] <= maxw and minh <= d["H"] <= maxh):
            continue
        if cx is not None and not (cx[0] <= d["cx"] <= cx[1]):
            continue
        if cy is not None and not (cy[0] <= d["cy"] <= cy[1]):
            continue
        if text is False and _norm(d["text"]):
            continue
        out.append(d)
    return out


def _card_of(sl, d, fams=("white",)):
    """找包住文本框 d 的最小圆角卡片。"""
    best = None
    for c in sl["shapes"]:
        if c is d or c["W"] <= 0 or c["H"] <= 0:
            continue
        if c["geom"] not in ("rect", "roundRect", "round2SameRect"):
            continue
        if _fam(c["fill"]) not in fams:
            continue
        if (c["L"] - 0.06 <= d["L"] and c["T"] - 0.06 <= d["T"]
                and c["L"] + c["W"] + 0.06 >= d["L"] + d["W"]):
            if best is None or c["W"] * c["H"] < best["W"] * best["H"]:
                best = c
    return best


def _below(sl, anchor, maxgap=5.0):
    """取锚点文本框下方、横向有交叠且最近的另一个文本框（说明文字）。"""
    best = None
    for d in sl["txt"]:
        if d is anchor:
            continue
        if d["T"] < anchor["T"] + max(anchor["H"] * 0.4, 0.15):
            continue
        if d["T"] > anchor["T"] + maxgap:
            continue
        ox = min(d["L"] + d["W"], anchor["L"] + anchor["W"]) - max(d["L"], anchor["L"])
        if ox <= 0.2:
            continue
        if best is None or d["T"] < best["T"]:
            best = d
    return best


# --------------------------------------------------- 段落级取用（同框标题＋说明）
# 两份输出文件的组织方式不同：一份把“卡标题”与“说明”放在同一个多段文本框里，
# 另一份拆成上下两个文本框。下列函数统一这两种写法，避免把相邻卡片误当作说明。
def _pidx(d, sub):
    """文本框 d 中包含关键词 sub 的段落下标；找不到返回 -1。"""
    k = _key(sub)
    if k:
        for j, p in enumerate(d["paras"]):
            if k in _key(p["text"]):
                return j
    return -1


def _tp(d, sub=None):
    """取锚点文本框中承载 sub 的段落（默认首段）。"""
    j = 0 if sub is None else max(_pidx(d, sub), 0)
    return d["paras"][j] if d["paras"] else {"text": "", "sz": None, "b": None,
                                             "color": None, "algn": None, "lnspc": 1.0}


def _desc(sl, anchor, sub=None, maxgap=5.0):
    """卡片说明文字：优先同框内标题段之后的非空段落，其次取下方最近的文本框。

    返回 (承载说明的文本框, 说明段列表)；都找不到返回 None。
    """
    j = 0 if sub is None else max(_pidx(anchor, sub), 0)
    rest = [p for p in anchor["paras"][j + 1:] if _norm(p["text"])]
    if rest:
        return anchor, rest
    d = _below(sl, anchor, maxgap)
    if d is None:
        return None
    return d, [p for p in d["paras"] if _norm(p["text"])]


def _pwraps(d, p, slack=1.06):
    """单个段落是否被动折行（留 6% 余量，避免临界宽度误判）。"""
    av = _avail(d) * slack
    segs = re.split(r"[\n\x0b]", p["text"])
    return any(_tw(seg, p["sz"] or DEF_SZ) > av for seg in segs)


def _grayish(col):
    """说明文字“灰色”判定：灰/黑/近黑，或未显式着色（继承深色正文色）。"""
    return col is None or _fam(col) in ("gray", "black")


def _emph(p, big=16.0):
    """强调：加粗，或字号已明显大于正文（部分制作程序用大字号代替加粗）。"""
    return bool(p["b"]) or (p["sz"] or DEF_SZ) >= big


def _lcs(a, b):
    prev = [0] * (len(b) + 1)
    for x in a:
        cur = [0]
        for j, y in enumerate(b):
            cur.append(prev[j] + 1 if x == y else max(cur[j], prev[j + 1]))
        prev = cur
    return prev[-1]


def _near(a, b, maxdiff=2):
    """两段文案是否仅有个别字差异（用于加分项，个别错字不整条判负）。"""
    a, b = _key(a), _key(b)
    if not a or not b:
        return False
    n = _lcs(a, b)
    return len(a) - n <= maxdiff and len(b) - n <= maxdiff


def _pnear(D, i, want, maxdiff=2):
    """第 i 页是否有与 want 相同或仅个别字差异的文案。"""
    if _key(want) in _S(D, i)["ktext"]:
        return True
    for d in _S(D, i)["txt"]:
        for p in d["paras"]:
            if _near(p["text"], want, maxdiff):
                return True
    return False


def _around(pts):
    """四个点是否环绕其几何中心分居四个方向（正交/对角两种基准任一成立）。"""
    if len(pts) < 4:
        return False
    cx = sum(p[0] for p in pts) / len(pts)
    cy = sum(p[1] for p in pts) / len(pts)
    for base in (0.0, 45.0):
        bins = set()
        for x, y in pts:
            a = (math.degrees(math.atan2(y - cy, x - cx)) - base) % 360.0
            bins.add(int(a // 90.0))
        if len(bins) >= 4:
            return True
    return False

# ------------------------------------------------------------------- 页面常量
# 页序锚点：每页任取其一命中即认为该页内容到位（避免依赖单一标题写法）。
_ORDER = [("以光为航",), ("班级初印象", "观察印象"), ("精准画像", "待唤醒的星火"),
          ("核心理念", "发现微光"), ("发展目标", "看见微光"), ("创意实践", "内驱力"),
          ("第一步", "岗位认领"), ("第二步", "获得光点"), ("第三步", "同盟"),
          ("第四步", "我敢开口了"), ("特色成效", "自信在发声"),
          ("感谢聆听", "教育是一场温柔的点亮")]
_SUBS = ["从观察到行动", "温柔激活与持续托举", "班级运行的底层逻辑", "从看见到自我驱动",
         "把理念转成班级日常", "让“我能做”变得具体", "把积极行为转化为正向反馈",
         "把单点光芒连成群星", "把阶段成果变成真实体验", "看见改变，沉淀习惯"]


def _vec(D, i):
    return _S(D, i)["vec"]


def _cap(d):
    """圆角矩形是否为两端全圆角的胶囊形。"""
    if d["geom"] not in ("roundRect", "round2SameRect", "pill"):
        return False
    a = d["adj"].get("adj")
    if a is None:
        a = d["adj"].get("adj1")
    if a is None:
        a = 16667                      # roundRect 默认圆角，远小于胶囊
    return a >= 40000


def _fmt(bad, ok_msg):
    return (not bad), ("；".join(bad[:6]) if bad else ok_msg)


# --------------------------------------------------------------- 加分项 P01-P20
def p01(D):
    bad = []
    n = len(D["slides"])
    if n != 12:
        bad.append("共%d页" % n)
    if not (33.5 <= D["pw"] <= 34.2 and 18.8 <= D["ph"] <= 19.3):
        bad.append("页面%.2f×%.2fcm" % (D["pw"], D["ph"]))
    for k, keys in enumerate(_ORDER, 1):
        if not any(_hask(D, k, s) for s in keys):
            bad.append("第%d页缺%s" % (k, "/".join(keys)))
    return _fmt(bad, "12页、16:9、页序与细则一致")


def p02(D):
    bad = []
    for sl in D["slides"]:
        need = 3 if 2 <= sl["i"] <= 11 else 2
        if len(sl["txt"]) < need:
            bad.append("第%d页仅%d个可编辑文本" % (sl["i"], len(sl["txt"])))
    return _fmt(bad, "各页文字均为独立可编辑文本框，共%d处"
                % sum(len(s["txt"]) for s in D["slides"]))

def _sans(name):
    s = (name or "").strip().lower()
    if not s:
        return False
    if any(w in s for w in _SERIF_EA):
        return False
    return any(w in s for w in _SANS_EA)


def p03(D):
    ea, hans = D["ctx"]["ea"], D["ctx"]["hans"]
    if _sans(ea):
        return True, "主题东亚字体 a:ea=%s" % ea
    if not ea and _sans(hans):
        return True, "主题 a:ea 为空但 Hans 字体映射=%s" % hans
    tot = miss = 0
    sample = ""
    for sl in D["slides"]:
        for d in sl["txt"]:
            for p in d["paras"]:
                if not re.search(r"[一-鿿]", p["text"]):
                    continue
                tot += 1
                if not _sans(p["ea"]):
                    miss += 1
                    if not sample:
                        sample = "如第%d页“%s”a:ea=%s" % (
                            sl["i"], p["text"][:8], p["ea"] or "未指定")
    if tot and not miss:
        return True, "全部%d个中文段落均显式指定黑体类东亚字体" % tot
    return False, ("主题 a:ea=%r / Hans=%r 非黑体类，且%d/%d个中文段落未指定东亚字体（%s）"
                   % (ea, hans, miss, tot, sample))


def p04(D):
    sl, bad = _S(D, 1), []
    t = _one(D, 1, "以光为航")
    if t is None:
        bad.append("缺主标题“以光为航 · 以心育人”")
    else:
        p0 = t["paras"][0]
        if not _emph(p0, 26.0) or _fam(p0["color"]) != "teal" or not _is_dark(p0["color"]):
            bad.append("主标题非深墨绿加粗(色=%s,b=%s)" % (p0["color"], p0["b"]))
        if _nsegs(t) > 1:
            bad.append("主标题被拆成%d行" % _nsegs(t))
    s = _one(D, 1, "晨河班成长共创方案")
    if s is None:
        bad.append("缺副标题“晨河班成长共创方案”")
    elif not (_emph(s["paras"][0], 20.0) and _is_white(s["paras"][0]["color"])):
        bad.append("副标题非白色加粗")
    return _fmt(bad + _p04_bottom(D, sl), "第1页四处文字齐全、颜色与分行方式一致")

def _p04_bottom(D, sl):
    bad = []
    a = _one(D, 1, "班级文化与学生发展路径")
    b = _one(D, 1, "让安静的努力被看见")
    if a is None:
        bad.append("缺“班级文化与学生发展路径”")
    if b is None:
        bad.append("缺“让安静的努力被看见，让微小的改变被珍惜”")
    if a is not None and b is not None:
        pa, pb = a["paras"][0], b["paras"][0]
        if not _is_white(pa["color"]) or not _is_white(pb["color"]):
            bad.append("底部两行非白色(%s/%s)" % (pa["color"], pb["color"]))
        if not _emph(pa, 20.0):
            bad.append("底部上行未加粗")
        sa, sb = pa["sz"] or DEF_SZ, pb["sz"] or DEF_SZ
        if sa <= sb + 1.0:
            bad.append("底部上行字号%.1f未明显大于下行%.1f" % (sa, sb))
        if (a["paras"][0]["algn"] or "ctr") != "ctr" or (b["paras"][0]["algn"] or "ctr") != "ctr":
            bad.append("底部两行未居中")
    if sl["vec"]:
        if not _sel(sl, geom=("roundRect",), fill=("white",), line=GOLD, minw=8.0):
            bad.append("主标题缺白底描橙金细边圆角框")
        if not _sel(sl, geom=("roundRect",), fill=GOLD, minw=4.0, maxh=3.0):
            bad.append("副标题缺橙黄圆角胶囊")
    return bad


def p05(D):
    bad = []
    for i in range(2, 12):
        sl = _S(D, i)
        num, pg = "%02d" % (i - 1), "%02d" % i
        if not [d for d in sl["txt"] if _norm(d["text"]) == num and d["T"] < 4.0]:
            bad.append("第%d页顶部缺序号%s" % (i, num))
        if not _pnear(D, i, _SUBS[i - 2], 2):
            bad.append("第%d页缺副标题“%s”" % (i, _SUBS[i - 2]))
        pgs = [d for d in sl["txt"] if _norm(d["text"]) == pg and d["T"] > D["ph"] * 0.70]
        if not pgs:
            bad.append("第%d页右下角缺页码%s" % (i, pg))
        else:
            sz = pgs[0]["paras"][0]["sz"] or DEF_SZ
            if not (10.5 <= sz <= 13.5):
                bad.append("第%d页页码字号%.1f磅" % (i, sz))
        if sl["vec"] and not _sel(sl, geom=("roundRect",), fill=("white",), line=GOLD,
                                  minw=12.0, maxh=4.0, cy=(0.0, 4.0)):
            bad.append("第%d页缺顶部白底描橙金细边圆角长条" % i)
    return _fmt(bad, "第2-11页顶部长条/序号01-10/副标题/页码02-11齐备")

def p06(D):
    sl, bad = _S(D, 2), []
    for k in ("观察印象", "现状扫描", "带班信念", "核心目标"):
        if not _hask(D, 2, k):
            bad.append("缺“%s”" % k)
    t = _one(D, 2, "观察印象")
    if t is not None:
        p0 = t["paras"][0]
        if not p0["b"] or _fam(p0["color"]) != "teal":
            bad.append("“观察印象”非深青加粗(色=%s,b=%s)" % (p0["color"], p0["b"]))
    for k in ("现状扫描", "带班信念", "核心目标"):
        a = _one(D, 2, k)
        if a is None:
            continue
        got = _desc(sl, a, k)
        if got is None:
            bad.append("“%s”下方缺灰色说明" % k)
        elif not _grayish(got[1][0]["color"]):
            bad.append("“%s”说明非灰色(色=%s)" % (k, got[1][0]["color"]))
    if sl["vec"]:
        if not _sel(sl, geom=("ellipse",), fill=TEAL, maxw=2.4, maxh=2.4):
            bad.append("缺深青圆形图标")
        if not _sel(sl, geom=("heart",), fill=GOLD):
            bad.append("缺橙黄心形图标")
        if not _sel(sl, geom=("star5",), fill=GREENS):
            bad.append("缺绿色五角星图标")
    return _fmt(bad, "第2页左侧大卡与右侧三卡（图标＋深青加粗标题＋灰色说明）齐备")


_S3_ITEMS = ("学习：", "纪律：", "心理：", "性格：", "氛围：", "潜能：")


def p07(D):
    sl, bad = _S(D, 3), []
    for k in ("待唤醒的星火", "最坚实的土壤"):
        a = _one(D, 3, k)
        if a is None:
            bad.append("缺“%s”" % k)
        elif not a["paras"][0]["b"] or _fam(a["paras"][0]["color"]) != "teal":
            bad.append("“%s”非深青加粗" % k)
    for k in _S3_ITEMS:
        got = _find(D, 3, k)
        if not got:
            bad.append("缺以“%s”起首的条目" % k)
        elif _wraps(got[0]):
            bad.append("条目“%s”折行" % k)
    if not _hask(D, 3, "破局关键"):
        bad.append("缺底部“破局关键：…”胶囊文字")
    if sl["vec"] and not _sel(sl, geom=("roundRect",), fill=TEAL, minw=12.0, maxh=3.0,
                              cy=(D["ph"] * 0.62, D["ph"])):
        bad.append("缺底部深墨绿圆角胶囊")
    return _fmt(bad, "第3页两张大卡6条浅底条目单行显示，底部胶囊文字完整")

_S4 = ("发现微光", "及时回应", "伙伴互助", "持续生长")


def p08(D):
    sl, bad = _S(D, 4), []
    pos = []
    for k in _S4:
        a = _one(D, 4, k)
        if a is None:
            bad.append("缺“%s”卡" % k)
            continue
        if not _tp(a, k)["b"] or _fam(_tp(a, k)["color"]) != "teal":
            bad.append("“%s”非深青加粗" % k)
        pos.append((a["cx"], a["cy"]))
        got = _desc(sl, a, k)
        if got is None:
            bad.append("“%s”下方缺灰色说明" % k)
        elif not _grayish(got[1][0]["color"]):
            bad.append("“%s”说明非灰色(色=%s)" % (k, got[1][0]["color"]))
    if len(pos) == 4 and not _around(pos):
        bad.append("四张卡未分居四角")
    if sl["vec"]:
        cxr = (D["pw"] * 0.33, D["pw"] * 0.67)
        cyr = (D["ph"] * 0.25, D["ph"] * 0.80)
        if not _sel(sl, geom=("ellipse",), fill=TEAL, cx=cxr, cy=cyr):
            bad.append("中心缺深青实心圆")
        rings = _sel(sl, geom=("ellipse", "donut", "circle"), nofill=True, cx=cxr, cy=cyr)
        if not [r for r in rings if _fam(r["line"]) in GOLD]:
            bad.append("中心缺一圈橙金细环")
        if len([r for r in rings if _fam(r["line"]) in GREENS]) < 2:
            bad.append("中心缺两圈绿色细环")
        if len(_sel(sl, geom=("line", "straightConnector1"), line=TEAL)) < 4:
            bad.append("缺四条自中心连向四角的细青线")
        if len(_sel(sl, geom=("roundRect",), fill=("white",), line=GOLD, minw=5.0)) < 4:
            bad.append("四角卡未描橙金细边")
    return _fmt(bad, "第4页四角卡＋中心圆环＋四条连接线齐备")


_S5 = (("起步阶段", "看见微光", ("green",)), ("连接阶段", "形成光束", ("teal",)),
       ("跃迁阶段", "自我点亮", GOLD))

def p09(D):
    sl, bad, pos = _S(D, 5), [], []
    for stage, name, fams in _S5:
        a = _one(D, 5, stage)
        if a is None:
            bad.append("缺阶段名“%s”" % stage)
        elif _fam(_tp(a, stage)["color"]) not in fams:
            bad.append("阶段名“%s”颜色%s不符" % (stage, _tp(a, stage)["color"]))
        b = _one(D, 5, name)
        if b is None:
            bad.append("缺名称“%s”" % name)
        else:
            if not _emph(_tp(b, name), 16.0):
                bad.append("“%s”未加粗" % name)
            pos.append((b["cx"], b["cy"]))
            if _desc(sl, b, name) is None:
                bad.append("“%s”下方缺灰色说明" % name)
    if len(pos) == 3:
        pos.sort()
        if not (pos[0][1] > pos[1][1] > pos[2][1]):
            bad.append("三张卡未呈左下至右上阶梯排列")
    if not _hask(D, 5, "目标不是让每个孩子立刻耀眼"):
        bad.append("缺底部“目标不是让每个孩子立刻耀眼…”单行文字")
    if sl["vec"]:
        if len(_sel(sl, geom=("ellipse",), fill=("white",), maxw=3.6, maxh=3.6)) < 3:
            bad.append("缺三处白色圆形图标底")
        rings = [d for d in sl["shapes"] if d["geom"] in ("ellipse", "donut", "circle")
                 and d["W"] <= 4.2 and (_fam(d["line"]) in ("green", "teal") + GOLD
                                        or _fam(d["fill"]) in ("green", "teal") + GOLD)]
        if len(rings) < 3:
            bad.append("缺绿/深青/橙三色圆环（仅%d处）" % len(rings))
        slant = [d for d in sl["shapes"] if not _norm(d["text"]) and d["H"] < 1.8
                 and d["W"] > 0.8 and 5.0 < (d["rot"] % 180.0) < 175.0]
        if len(slant) < 2:
            bad.append("相邻卡片间缺青金双色斜条")
    return _fmt(bad, "第5页三张阶段卡阶梯排列、图标与底部单行齐备")

def p10(D):
    sl, bad = _S(D, 6), []
    for k in ("成长的", "内驱力", "心愿花园", "责任剧场", "伙伴联盟", "能量积分"):
        if not _hask(D, 6, k):
            bad.append("缺“%s”" % k)
    a = _one(D, 6, "内驱力")
    if a is not None and (not a["paras"][0]["b"] or _fam(a["paras"][0]["color"]) != "teal"):
        bad.append("“内驱力”非深青加粗")
    if sl["vec"]:
        w = _sel(sl, geom=("pie", "pieWedge", "blockArc", "arc", "chord", "donut"))
        fams = set(_fam(d["fill"]) for d in w)
        if len(w) < 4:
            bad.append("中部扇区仅%d个" % len(w))
        if not {"red", "teal", "green"}.issubset(fams) or not (fams & set(GOLD)):
            bad.append("扇区色系%s与上红/右深青/左绿/下橙黄不符"
                       % sorted(f for f in fams if f))
        if not _sel(sl, geom=("ellipse", "pie", "moon", "chord"), fill=("white",),
                    minw=2.0, maxw=9.0):
            bad.append("环心缺带凹口的白色圆形")
        if not _sel(sl, geom=("heart",), fill=("white",)):
            bad.append("红色扇区内缺白色心形")
        if not _sel(sl, geom=("star5",), fill=("white",)):
            bad.append("橙色扇区内缺白色五角星")
    return _fmt(bad, "第6页四等分圆环、环心文字与四个扇区标注齐备")


_S6C = (("一、责任剧场", "把班级事务改造成可参与的小角色"),
        ("二、能量积分", "把正向行为转化为持续反馈"),
        ("三、伙伴联盟", "让互助从偶然变成一种结构"),
        ("四、心愿花园", "用阶段成果承接情感期待"))
_S6P = ("低门槛参与", "高频率反馈", "小步子迭代", "真实感表达")


def p11(D):
    sl, bad = _S(D, 6), []
    # 说明文案的逐字准确性由减分项 N25 单独考核，此处只判“有灰色单行说明”，
    # 避免同一处文案差异在加分项与减分项中被重复计分。
    for j, (t, _dsc) in enumerate(_S6C):
        a = _one(D, 6, t)
        if a is None:
            bad.append("缺标题“%s”" % t)
            continue
        if not _emph(_tp(a, t), 16.0):
            bad.append("标题“%s”未加粗" % t)
        got = _desc(sl, a, t)
        if got is None:
            bad.append("“%s”缺灰色说明" % t)
        elif not _grayish(got[1][0]["color"]):
            bad.append("“%s”说明非灰色(色=%s)" % (t, got[1][0]["color"]))
        elif _pwraps(got[0], got[1][0]):
            bad.append("“%s”说明非单行" % t)
    for j, n in enumerate("1234"):
        cn = "一二三四"[j]
        if [d for d in sl["txt"] if _norm(d["text"]) in (n, "0" + n)]:
            continue
        if _hask(D, 6, cn + "、"):  # 序号已并入标题“一、责任剧场”写法
            continue
        bad.append("缺数字标签%s" % n)
    if not _hask(D, 6, "运行原则"):
        bad.append("缺“运行原则”")
    for k in _S6P:
        got = _find(D, 6, k)
        if not got:
            bad.append("缺胶囊“%s”" % k)
        elif not _is_white(got[0]["paras"][0]["color"]):
            bad.append("胶囊“%s”文字非白色" % k)
    if sl["vec"]:
        if len([d for d in _sel(sl, geom=("roundRect",), minw=2.5, maxh=1.8) if _cap(d)]) < 4:
            bad.append("右侧四个白字圆角胶囊不全")
        if not _sel(sl, geom=("roundRect",), fill=("white",), line=GOLD, minw=4.0, minh=3.0):
            bad.append("右侧缺白底描橙金边卡")
    return _fmt(bad, "第6页四张编号卡与“运行原则”四胶囊齐备")

_S7 = ("岗位认领", "微光观察员", "自我提醒官")


def p12(D):
    sl, bad = _S(D, 7), []
    top = _one(D, 7, "把班级日常拆成若干小岗位")
    if top is None:
        bad.append("顶部缺“把班级日常拆成若干小岗位…”")
    elif _wraps(top):
        bad.append("顶部长条文字折成%d行" % _nlines(top))
    for k in _S7:
        a = _one(D, 7, k)
        if a is None:
            bad.append("缺“%s”卡" % k)
            continue
        if not _tp(a, k)["b"] or _fam(_tp(a, k)["color"]) != "teal":
            bad.append("“%s”非深青加粗" % k)
        got = _desc(sl, a, k)
        if got is None:
            bad.append("“%s”下方缺灰色说明" % k)
        elif _pwraps(got[0], got[1][0]):
            bad.append("“%s”说明非单行" % k)
    if not _hask(D, 7, "成效期待"):
        bad.append("缺底部“成效期待：…”胶囊文字")
    if sl["vec"]:
        gears = _sel(sl, geom=("gear6", "gear9", "sun", "star16", "star24"))
        if len(gears) < 2:
            bad.append("左侧缺深青与绿色齿轮图标")
        if not _sel(sl, geom=("roundRect",), fill=("white",), line=GOLD, minw=12.0,
                    maxh=4.0, cy=(0.0, 6.0)):
            bad.append("缺顶部白底描橙金边圆角长条")
        if not _sel(sl, geom=("roundRect",), fill=TEAL, minw=10.0, maxh=3.0,
                    cy=(D["ph"] * 0.62, D["ph"])):
            bad.append("缺底部深青圆角胶囊")
    return _fmt(bad, "第7页顶部长条、三张岗位卡、齿轮与底部胶囊齐备")


_S8 = ("获得光点", "复盘升级", "存入光册", "兑换体验")

def p13(D):
    sl, bad = _S(D, 8), []
    pos = []
    for k in _S8:
        a = _one(D, 8, k)
        if a is None:
            bad.append("缺“%s”卡" % k)
            continue
        pos.append((a["cx"], a["cy"]))
    if len(pos) == 4 and not _around(pos):
        bad.append("四张卡未环绕中心排布")
    if not (_hask(D, 8, "能量") and _hask(D, 8, "循环")):
        bad.append("中心圆内缺“能量”“循环”两行")
    if not _hask(D, 8, "规则底色"):
        bad.append("缺底部“规则底色：…”胶囊文字")
    if sl["vec"]:
        if not _sel(sl, geom=("ellipse",), fill=TEAL, minw=2.0,
                    cx=(D["pw"] * 0.30, D["pw"] * 0.70)):
            bad.append("中心缺深青实心圆")
        want = [TEAL, ("red",), GOLD, GREENS]
        lines = [_fam(d["line"]) for d in _sel(sl, geom=("roundRect", "rect"),
                                               fill=("white",), minw=4.0, minh=2.5)]
        for fams, k in zip(want, _S8):
            if not [f for f in lines if f in fams]:
                bad.append("“%s”卡描边色不符" % k)
        if not _sel(sl, geom=("roundRect",), fill=GOLD, minw=10.0, maxh=3.0,
                    cy=(D["ph"] * 0.62, D["ph"])):
            bad.append("缺底部橙色圆角胶囊")
    return _fmt(bad, "第8页四卡环绕中心深青圆、底部胶囊文字齐备")


_S9H = ("支持", "创意", "表达", "整理", "观察", "记录")
_S9HF = {"支持": ("red",), "创意": ("blue", "teal"), "表达": ("teal",),
         "整理": GOLD, "观察": GREENS, "记录": ("purple", "pink", "blue")}

def p14(D):
    sl, bad = _S(D, 9), []
    for k in _S9H:
        got = [d for d in sl["txt"] if _norm(d["text"]) == k] or _find(D, 9, k)
        if not got:
            bad.append("缺六边形标签“%s”" % k)
        elif not _is_white(got[0]["paras"][0]["color"]):
            bad.append("标签“%s”非白色" % k)
        elif _wraps(got[0]) or got[0]["vert"] in ("vert", "eaVert", "vert270"):
            bad.append("标签“%s”非横排单行" % k)
    if not _hask(D, 9, "同盟"):
        bad.append("中心缺“同盟”")
    if sl["vec"]:
        hexes = _sel(sl, geom=("hexagon", "hexagon2"))
        fams = set(_fam(d["fill"]) for d in hexes)
        if len(hexes) < 6:
            bad.append("平顶正六边形仅%d个" % len(hexes))
        elif len([f for f in fams if f]) < 5:
            bad.append("六边形色系仅%d种：%s" % (len(fams), sorted(f for f in fams if f)))
        if not _sel(sl, geom=("ellipse",), fill=("white",), line=GOLD, minw=1.5):
            bad.append("中心缺白色圆形（外描橙金环）")
        if not _sel(sl, geom=("star5",), fill=GOLD):
            bad.append("中心缺橙黄五角星")
    return _fmt(bad, "第9页六色平顶六边形蜂巢＋中心白圆齐备")


_S9C = ("搭建方式", "核心机制", "温柔约定")


def p15(D):
    sl, bad = _S(D, 9), []
    for k in _S9C:
        a = _one(D, 9, k)
        if a is None:
            bad.append("缺“%s”卡" % k)
            continue
        if not _tp(a, k)["b"] or _fam(_tp(a, k)["color"]) != "teal":
            bad.append("“%s”非深青加粗" % k)
        got = _desc(sl, a, k)
        if got is None:
            bad.append("“%s”下方缺两行灰色说明" % k)
        elif not _grayish(got[1][0]["color"]):
            bad.append("“%s”说明非灰色(色=%s)" % (k, got[1][0]["color"]))
    b = _one(D, 9, "合作的目的不是把每个人变成一样")
    if b is None:
        bad.append("缺底部“合作的目的不是把每个人变成一样…”")
    elif _wraps(b):
        bad.append("底部文字折成%d行" % _nlines(b))
    return _fmt(bad, "第9页右侧三卡与底部单行齐备")

def p16(D):
    sl, bad = _S(D, 10), []
    # 页标题“第四步：心愿花园与成长展”同样含该关键词，优先取整段即为“心愿花园”的文本框。
    exact = [d for d in sl["txt"] if _norm(d["text"]) == "心愿花园"]
    a = exact[0] if exact else _one(D, 10, "心愿花园")
    if a is None:
        bad.append("左上大卡内缺“心愿花园”")
    else:
        p0 = _tp(a, "心愿花园")
        if not _emph(p0, 20.0) or _fam(p0["color"]) not in GOLD or not _is_pale(p0["color"]):
            bad.append("“心愿花园”非浅黄色加粗(色=%s,b=%s)" % (p0["color"], p0["b"]))
    for k in ("我敢开口了", "我能坚持了"):
        b = _one(D, 10, k)
        if b is None:
            bad.append("缺“%s”" % k)
        elif not _tp(b, k)["b"] or _fam(_tp(b, k)["color"]) != "teal":
            bad.append("“%s”非深青加粗" % k)
        elif _desc(sl, b, k) is None:
            bad.append("“%s”缺灰色单行说明" % k)
    for k in ("升级方式", "情感闭环"):
        c = _one(D, 10, k)
        if c is None:
            bad.append("缺“%s”卡" % k)
        elif _desc(sl, c, k) is None:
            bad.append("“%s”下方缺两行灰色说明" % k)
    if sl["vec"]:
        if not _sel(sl, geom=("roundRect",), fill=TEAL, minw=6.0, minh=4.0,
                    cx=(0.0, D["pw"] * 0.55), cy=(0.0, D["ph"] * 0.60)):
            bad.append("左上缺深墨绿圆角大卡")
        if len(_sel(sl, geom=("star5",), cy=(0.0, D["ph"] * 0.55))) < 5:
            bad.append("大卡内横排五枚彩色五角星不全")
        golds = _sel(sl, geom=("roundRect", "rect"), fill=("white",), line=GOLD,
                     minw=4.0, cx=(D["pw"] * 0.45, D["pw"]))
        teals = _sel(sl, geom=("roundRect", "rect"), fill=("white",), line=TEAL,
                     minw=4.0, cx=(D["pw"] * 0.45, D["pw"]))
        if not golds or not teals:
            bad.append("右侧两卡未分别描橙金边与深青边")
    return _fmt(bad, "第10页左上大卡、左下两卡与右侧两卡齐备")


def p17(D):
    sl = _S(D, 11)
    if not sl["vec"]:
        return True, "第11页图形由整页背景图承载，三处顶部图标按原样保留"
    bad = []
    if not _sel(sl, geom=("ellipse",), fill=GOLD, maxw=3.6, maxh=3.6):
        bad.append("缺橙黄圆形（内嵌白色喇叭）")
    if not [d for d in _sel(sl, geom=("roundRect", "rect"), maxw=3.6, maxh=3.6)
            if d["fill"] in ("none", None) and _fam(d["line"]) in GREENS]:
        bad.append("缺绿色中空圆角方框（只描边不填色）")
    if not _sel(sl, geom=("ellipse",), fill=TEAL, maxw=3.6, maxh=3.6):
        bad.append("缺深青实心圆")
    if not _sel(sl, geom=("heart",), fill=("white",)):
        bad.append("深青圆内缺白色心形")
    order = [_fam(d["line"]) for d in sorted(_sel(sl, geom=("roundRect", "rect"),
                                                  fill=("white",), minw=6.0, minh=5.0),
                                             key=lambda d: d["L"])]
    if len(order) >= 3 and not (order[0] in GOLD and order[1] in GREENS
                                and order[2] in TEAL):
        bad.append("三张卡描边顺序%s与橙金/绿/深青不符" % order[:3])
    return _fmt(bad, "第11页三处顶部图标造型与卡片描边齐备")

_S11 = ("自信在发声", "习惯在迁移", "氛围在变暖")


def p18(D):
    sl, bad = _S(D, 11), []
    for k in _S11:
        a = _one(D, 11, k)
        if a is None:
            bad.append("缺标题“%s”" % k)
            continue
        p0 = _tp(a, k)
        if not p0["b"] or _fam(p0["color"]) != "teal":
            bad.append("“%s”非深青加粗" % k)
        if (p0["algn"] or "ctr") != "ctr":
            bad.append("“%s”未居中" % k)
    if len(_find(D, 11, "变化：")) < 3:
        bad.append("“变化：…”段落不足3处")
    if len(_find(D, 11, "影响：")) < 3:
        bad.append("“影响：…”段落不足3处")
    if not _hask(D, 11, "班级建设的成果"):
        bad.append("缺底部“班级建设的成果…”胶囊文字")
    if sl["vec"] and not _sel(sl, geom=("roundRect",), fill=TEAL, minw=10.0, maxh=3.0,
                              cy=(D["ph"] * 0.62, D["ph"])):
        bad.append("缺底部深青圆角胶囊")
    return _fmt(bad, "第11页三卡标题、变化/影响两段说明与底部胶囊齐备")


def p19(D):
    sl, bad = _S(D, 12), []
    a = _one(D, 12, "教育是一场温柔的点亮")
    if a is None:
        bad.append("缺“教育是一场温柔的点亮”")
    else:
        p0 = _tp(a, "教育是一场温柔的点亮")
        if not _emph(p0, 24.0) or _fam(p0["color"]) != "teal" or not _is_dark(p0["color"]):
            bad.append("主句非深墨绿加粗(色=%s,b=%s)" % (p0["color"], p0["b"]))
    for k in ("这份答案沉默而坚定", "正被一颗颗愿意尝试的心慢慢写出来"):
        b = _one(D, 12, k)
        if b is None:
            bad.append("缺“%s”" % k)
        elif not _emph(_tp(b, k), 20.0):
            bad.append("“%s”未加粗" % k)
    c = None
    for d in sl["txt"]:
        if _pidx(d, "感谢聆听与同行") >= 0:
            c = d
            break
    if c is None:
        bad.append("缺“感谢聆听与同行”")
    else:
        col = _tp(c, "感谢聆听与同行")["color"]
        if _fam(col) not in GOLD or not _is_pale(col):
            bad.append("“感谢聆听与同行”非浅黄色(色=%s)" % col)
    if sl["vec"] and not _sel(sl, geom=("roundRect",), fill=("white",), line=GOLD, minw=8.0):
        bad.append("缺白底描橙金边圆角框")
    return _fmt(bad, "第12页结束页四处文字齐备且颜色一致")

def p20(D):
    bad = []
    kept = 0
    for i in range(1, 13):
        sl = _S(D, i)
        if sl["bgimg"]:
            kept += 1
            continue
        big = _sel(sl, geom=("ellipse",), minw=3.0, minh=3.0)
        fams = set(_fam(d["fill"]) for d in big)
        dots = _sel(sl, geom=("ellipse", "star5", "star4"), minw=0.05, maxw=0.9, maxh=0.9)
        if i in (1, 12):
            if not (fams & set(GOLD)) or "white" not in fams or not (fams & set(TEAL)):
                bad.append("第%d页夜空底缺黄/白/青绿大圆(%s)"
                           % (i, sorted(f for f in fams if f)))
            if len(dots) < 6:
                bad.append("第%d页缺散落星点(%d处)" % (i, len(dots)))
        else:
            if not {"green", "teal", "white"}.issubset(fams) or not (fams & set(GOLD)):
                bad.append("第%d页缺淡绿/青绿/橙黄/绿/白装饰(%s)"
                           % (i, sorted(f for f in fams if f)))
            if len(dots) < 3:
                bad.append("第%d页缺散落橙黄小圆点(%d处)" % (i, len(dots)))
    return _fmt(bad, "背景装饰逐页保留（其中%d页以整页背景图原样承载）" % kept)


# --------------------------------------------------------------- 减分项 N01-N29
_BGKEEP = "该页装饰由整页背景图原样承载，未被改动"


def n01(D):
    sl = _S(D, 1)
    if not sl["vec"]:
        return False, _BGKEEP
    bad = []
    stray = [d for d in sl["txt"]
             if re.match(r"^[\s -⯿〰︐-︟＿~^_\-]+$", d["text"] or "")]
    stars = _sel(sl, geom=("star5",), minw=1.2)
    if not stars:
        bad.append("第1页未见五角星形")
    else:
        s = max(stars, key=lambda d: d["W"] * d["H"])

        def _in(d):
            return (s["L"] <= d["cx"] <= s["L"] + s["W"]
                    and s["T"] <= d["cy"] <= s["T"] + s["H"])
        eyes = [d for d in _sel(sl, geom=("ellipse",), maxw=1.0, maxh=1.0)
                if _in(d) and _is_dark(d["fill"])]
        mouth = [d for d in sl["shapes"]
                 if d["geom"] in ("arc", "curvedDown", "curvedUp", "chord", "blockArc", "custom")
                 and _in(d) and (_is_dark(d["line"]) or _is_dark(d["fill"]))]
        if len(eyes) < 2:
            bad.append("星形内深色圆眼睛仅%d个" % len(eyes))
        if not mouth:
            bad.append("星形内缺深色弧形笑嘴")
    if stray:
        bad.append("出现游离于星形之外的符号文本%r" % stray[0]["text"][:6])
    return (bool(bad), "；".join(bad[:4]) if bad else "笑脸五角星造型完整保留")

def n02(D):
    sl = _S(D, 1)
    if not sl["vec"]:
        return False, _BGKEEP
    hills = [d for d in sl["shapes"] if _fam(d["fill"]) in GOLD and d["W"] >= 2.0
             and d["geom"] in ("pie", "chord", "arc", "halfFrame", "ellipse",
                               "blockArc", "moon", "custom")]
    if not hills:
        return True, "五角星下方未见黄色半圆形山丘"
    h = max(hills, key=lambda d: d["W"])
    stripes = [d for d in sl["shapes"] if _is_white(d["fill"]) and d["H"] <= 0.7
               and d["W"] >= 0.7
               and h["T"] - 0.3 <= d["cy"] <= h["T"] + h["H"] + 0.3
               and h["L"] - 0.6 <= d["cx"] <= h["L"] + h["W"] + 0.6]
    if len(stripes) < 2:
        return True, "山丘上白色横向条纹仅%d条" % len(stripes)
    return False, "黄色半圆山丘＋%d条白色横纹保留" % len(stripes)


def n03(D):
    sl = _S(D, 1)
    if not sl["vec"]:
        return False, _BGKEEP
    low = D["ph"] * 0.45
    waves = [d for d in sl["shapes"] if d["T"] > low
             and (_is_white(d["line"]) or _is_white(d["fill"]))
             and d["geom"] in ("arc", "curvedDown", "curvedUp", "custom", "blockArc", "chord")]
    seg = [d for d in sl["shapes"] if d["T"] > low and _is_white(d["line"])
           and d["geom"] in ("line", "straightConnector1")]
    if len(waves) >= 8:
        return False, "底部保留%d条白色弧形波浪线" % len(waves)
    return True, "底部白色弧形波浪线仅%d条（另有%d条白色直线段）" % (len(waves), len(seg))


def n04(D):
    t = _one(D, 1, "以光为航")
    if t is None:
        return True, "未找到第1页主标题文本"
    raw = (t["text"] or "").replace("\x0b", " ")
    tight = re.sub(r"\s+", " ", raw).strip()
    if tight == "以光为航 · 以心育人":
        return False, "主标题为“以光为航 · 以心育人”紧凑排列"
    if _norm(raw) == "以光为航·以心育人" and (len(raw) - len(_norm(raw))) <= 2:
        return False, "主标题为紧凑排列（%r）" % tight
    return True, "主标题字距被拉开：%r" % raw

def n05(D):
    sl = _S(D, 1)
    if not sl["vec"]:
        return False, _BGKEEP
    cands = [d for d in sl["shapes"] if _fam(d["fill"]) in ("teal", "green")
             and d["W"] >= 2.5 and d["cy"] > D["ph"] * 0.5 and d["cx"] > D["pw"] * 0.5]
    if not cands:
        return True, "第1页右下角未见青绿色丘形"
    if [d for d in cands if d["L"] + d["W"] > D["pw"] + 0.05]:
        return False, "右下角青绿丘形仍被页面右缘裁切"
    d = max(cands, key=lambda x: x["W"])
    return True, "右下角青绿形状(%s，右缘%.2fcm<页宽%.2fcm)未被裁切，已改为完整形状" % (
        d["geom"], d["L"] + d["W"], D["pw"])


def n06(D):
    sl = _S(D, 2)
    if not sl["vec"]:
        return False, _BGKEEP
    blocks = [d for d in sl["shapes"] if _fam(d["fill"]) == "teal" and _is_dark(d["fill"])
              and d["W"] >= 3.5 and d["geom"] in ("rect", "roundRect")]
    if not blocks:
        return True, "左侧大卡内未见深墨绿色缩略图块"
    b = max(blocks, key=lambda d: d["W"] * d["H"])

    def _in(d):
        return (b["L"] - 0.2 <= d["cx"] <= b["L"] + b["W"] + 0.2
                and b["T"] - 0.2 <= d["cy"] <= b["T"] + b["H"] + 0.2)
    arcs = [d for d in sl["shapes"] if d is not b and _in(d)
            and d["geom"] in ("arc", "ellipse", "blockArc", "chord", "custom")
            and (d["fill"] in ("none", None) or _is_pale(d["fill"]))]
    dots = [d for d in sl["shapes"] if d is not b and _in(d) and d["geom"] == "ellipse"
            and d["W"] <= 0.8 and _fam(d["fill"]) in GOLD]
    if len(arcs) >= 3 and len(dots) >= 3:
        return False, "星轨图案保留（%d条浅青弧线＋%d个黄点）" % (len(arcs), len(dots))
    return True, "星轨被简化：浅青同心弧线%d条、黄色圆点%d个" % (len(arcs), len(dots))


def n07(D):
    miss = [k for k in _S3_ITEMS if not _find(D, 3, k)]
    if miss:
        return True, "第3页缺条目：%s" % "、".join(miss)
    fold = [k for k in _S3_ITEMS if _wraps(_find(D, 3, k)[0])]
    if fold:
        return True, "折行越出浅色底条的条目：%s" % "、".join(fold)
    return False, "6条浅色底条目均单行完整显示"

def n08(D):
    sl = _S(D, 5)
    if not sl["vec"]:
        return False, _BGKEEP
    circles = _sel(sl, geom=("ellipse",), fill=("white",), minw=0.8, maxw=3.6, maxh=3.6)
    rings = []
    for d in sl["shapes"]:
        # 外环：足够大且以描边（或环形填充）呈现的彩色圆，散落小圆点不计。
        if d["geom"] not in ("ellipse", "donut", "circle") or not (0.8 <= d["W"] <= 4.4):
            continue
        c = d["line"] if d["line"] not in ("none", None) else d["fill"]
        if _fam(c) in ("green", "teal") + GOLD:
            rings.append(d)
    part = [d for d in sl["shapes"] if d["geom"] in ("arc", "blockArc", "pie", "chord")
            and d["W"] <= 4.4 and _fam(d["line"] if d["line"] not in ("none", None)
                                      else d["fill"]) in ("green", "teal") + GOLD]
    if len(circles) >= 3 and len(rings) >= 3 and not part:
        return False, "三处图标均为白色圆形＋完整彩色外环"
    return True, "阶段图标不完整：白圆%d个、完整彩环%d个、残缺弧形环%d个" % (
        len(circles), len(rings), len(part))


def n09(D):
    d = _one(D, 5, "目标不是让每个孩子立刻耀眼")
    if d is None:
        return True, "第5页底部文字缺失"
    if _wraps(d):
        return True, "底部文字折成%d行（估算行宽%.2fcm>可用%.2fcm）" % (
            _nlines(d), _tw(d["text"], d["paras"][0]["sz"]), _avail(d))
    return False, "第5页底部文字单行显示"


_N10_PAGES = ((2, None), (5, None), (6, (0.0, 0.52)), (7, (0.28, 1.0)), (9, (0.45, 1.0)))


def n10(D):
    hits, checked = [], 0
    for i, xr in _N10_PAGES:
        sl = _S(D, i)
        if not sl["vec"]:
            continue
        checked += 1
        for d in _sel(sl, geom=("rect", "roundRect"), fill=("white",), minw=3.5, minh=1.6):
            if d["cy"] < 3.6:
                continue                       # 顶部长条本应描橙金边
            if xr and not (xr[0] * D["pw"] <= d["cx"] <= xr[1] * D["pw"]):
                continue
            if d["line"] not in ("none", None) and _fam(d["line"]) in GOLD:
                hits.append("第%d页(%.1f,%.1f)描边%s" % (i, d["L"], d["T"], d["line"]))
    if not checked:
        return False, "相关页面卡片由整页背景图承载，纯白无描边样式未被改动"
    if hits:
        return True, "被加上橙金/橙黄描边的白卡：%s" % "；".join(hits[:4])
    return False, "指定各页白卡均为无描边纯白圆角卡"

def n11(D):
    bad = []
    for i in range(2, 12):
        pg = "%02d" % i
        got = [d for d in _S(D, i)["txt"] if _norm(d["text"]) == pg]
        if not got:
            bad.append("第%d页无独立页码文本" % i)
            continue
        d = got[0]
        if d["vert"] in ("vert", "eaVert", "vert270"):
            bad.append("第%d页页码为竖排" % i)
        elif _wraps(d):
            bad.append("第%d页页码折成%d行（框宽%.2fcm/需%.2fcm）" % (
                i, _nlines(d), d["W"], _tw(pg, d["paras"][0]["sz"]) + 2 * DEF_LINS))
    if bad:
        return True, "；".join(bad[:5])
    return False, "第2-11页两位数页码均单行横排"


def _ovf_group(D, page, titles, label):
    """通用：卡片标题下方说明是否溢出卡片。"""
    sl, bad = _S(D, page), []
    for k in titles:
        a = _one(D, page, k)
        if a is None:
            bad.append("缺“%s”卡" % k)
            continue
        got = _desc(sl, a, k)
        if got is None:
            bad.append("“%s”下方无说明文字" % k)
            continue
        d = got[0]
        card = _card_of(sl, d) if sl["vec"] else None
        if _ovf(d, card):
            bad.append("“%s”说明溢出（估高%.2fcm/框高%.2fcm%s）" % (
                k, _est_h(d), d["H"],
                "" if card is None else "/卡底%.2fcm" % (card["T"] + card["H"])))
    if bad:
        return True, "；".join(bad[:4])
    return False, "%s说明均完整位于卡片内部" % label


def n12(D):
    return _ovf_group(D, 4, _S4, "第4页四角卡两行灰色")


def n13(D):
    sl, bad = _S(D, 4), []
    for k in _S4:
        a = _one(D, 4, k)
        if a is None:
            bad.append("缺卡标题“%s”" % k)
        elif _pwraps(a, _tp(a, k)):
            bad.append("卡标题“%s”折行（需%.2fcm/可用%.2fcm）" % (
                k, _tw(k, _tp(a, k)["sz"] or DEF_SZ), _avail(a)))
    for k in ("微光汇聚", "向阳生长"):
        a = _one(D, 4, k)
        if a is None:
            bad.append("中心圆内缺“%s”" % k)
        elif _pwraps(a, _tp(a, k)):
            bad.append("中心圆内“%s”折行" % k)
    if bad:
        return True, "；".join(bad[:4])
    return False, "第4页卡标题与中心圆文字均单行显示"

def n14(D):
    sl = _S(D, 6)
    if not sl["vec"]:
        return False, _BGKEEP
    w = _sel(sl, geom=("pie", "pieWedge", "blockArc", "arc", "chord", "donut"))
    keys = set((round(d["L"], 1), round(d["T"], 1), round(d["W"], 1), round(d["H"], 1),
                round(d["rot"]), tuple(sorted(d["adj"].items()))) for d in w)
    fams = set(_fam(d["fill"]) for d in w)
    if (len(w) >= 4 and len(keys) >= 3 and {"red", "green", "teal"}.issubset(fams)
            and (fams & set(GOLD))):
        return False, "四等分圆环完整（%d个扇区、%d组独立几何）" % (len(w), len(keys))
    return True, "圆环退化：扇区%d个、独立几何仅%d组、色系%s" % (
        len(w), len(keys), sorted(f for f in fams if f))


def n15(D):
    pages = [i for i in (4, 7, 11) if _S(D, i)["vec"]]
    if not pages:
        return False, "第4、7、11页底部横条由整页背景图承载，胶囊形未被改动"
    bad = []
    for i in pages:
        sl = _S(D, i)
        bars = [d for d in sl["shapes"] if d["W"] >= 9.0 and d["H"] <= 3.2
                and d["cy"] > D["ph"] * 0.62 and d["fill"] not in ("none", None)
                and _fam(d["fill"]) != "white"]
        if not bars:
            bad.append("第%d页底部未见横条" % i)
            continue
        b = max(bars, key=lambda d: d["W"])
        if not _cap(b):
            bad.append("第%d页底部横条为%s(adj=%s)非两端全圆角胶囊" % (
                i, b["geom"], b["adj"].get("adj", b["adj"].get("adj1", "默认16667"))))
    if bad:
        return True, "；".join(bad)
    return False, "第4、7、11页底部横条均为两端全圆角胶囊形"


def n16(D):
    sl = _S(D, 8)
    if not sl["vec"]:
        return False, _BGKEEP
    arrows, straight = [], []
    for d in sl["shapes"]:
        g = d["geom"] or ""
        c = d["fill"] if d["fill"] not in ("none", None) else d["line"]
        if g in ("line", "straightConnector1"):
            straight.append(d)
        if _fam(c) not in GOLD:
            continue
        if ("Arrow" in g or "arrow" in g
                or g in ("arc", "curvedDown", "curvedUp", "curvedConnector3",
                         "bentConnector3", "blockArc", "custom")):
            arrows.append(d)
    if len(arrows) >= 4:
        return False, "四段橙黄弧形箭头保留（%d段）" % len(arrows)
    return True, "橙黄弧形箭头仅%d段（另有%d条直线段）" % (len(arrows), len(straight))


def n17(D):
    return _ovf_group(D, 8, _S8, "第8页四张卡片灰色")

def n18(D):
    sl = _S(D, 9)
    if not sl["vec"]:
        return False, _BGKEEP
    hexes = _sel(sl, geom=("hexagon", "hexagon2"))
    if len(hexes) != 6:
        return True, "六边形数量为%d（应为6个）" % len(hexes)
    turned = [h for h in hexes if not (h["rot"] % 180.0 < 8.0 or h["rot"] % 180.0 > 172.0)]
    if turned:
        return True, "%d个六边形被旋转（如%.0f°），成尖顶朝上" % (len(turned), turned[0]["rot"])
    cx = sum(h["cx"] for h in hexes) / 6.0
    cy = sum(h["cy"] for h in hexes) / 6.0
    pitch = sum(h["H"] for h in hexes) / 6.0
    rs = [math.hypot(h["cx"] - cx, h["cy"] - cy) for h in hexes]
    if [r for r in rs if abs(r - pitch) > 0.08 * pitch]:
        return True, "环半径%.2f-%.2fcm与紧邻蜂巢应有的%.2fcm不符，六边形间留出缝隙" % (
            min(rs), max(rs), pitch)
    return False, "六个六边形平顶朝上且紧邻排布（环半径≈%.2fcm）" % pitch


def n19(D):
    sl, bad = _S(D, 9), []
    for k in _S9H:
        got = [d for d in sl["txt"] if _norm(d["text"]) == k] or _find(D, 9, k)
        if not got:
            bad.append("缺标签“%s”" % k)
            continue
        d = got[0]
        if d["vert"] in ("vert", "eaVert", "vert270"):
            bad.append("标签“%s”被竖排" % k)
        elif _wraps(d):
            bad.append("标签“%s”折成%d行" % (k, _nlines(d)))
    if sl["vec"]:
        t = [d for d in sl["txt"] if _norm(d["text"]) == "表达"]
        circ = _sel(sl, geom=("ellipse",), fill=("white",), minw=1.5)
        if t and circ:
            c = max(circ, key=lambda d: d["W"])
            if math.hypot(t[0]["cx"] - c["cx"], t[0]["cy"] - c["cy"]) < c["W"] / 2.0 * 0.85:
                bad.append("“表达”被中心白色圆形遮挡")
    if bad:
        return True, "；".join(bad[:4])
    return False, "六个六边形标签均横排单行完整可见"


def n20(D):
    return _ovf_group(D, 9, _S9C, "第9页右侧三卡两行")

def n21(D):
    sl = _S(D, 11)
    if not sl["vec"]:
        return False, _BGKEEP
    bad = []
    if not _sel(sl, geom=("ellipse",), fill=GOLD, maxw=3.8, maxh=3.8):
        bad.append("缺橙黄圆形喇叭底")
    if not [d for d in _sel(sl, geom=("roundRect", "rect"), maxw=3.8, maxh=3.8)
            if d["fill"] in ("none", None) and _fam(d["line"]) in GREENS]:
        bad.append("缺绿色中空圆角方框（只描边、内部不填色）")
    if not _sel(sl, geom=("ellipse",), fill=TEAL, maxw=3.8, maxh=3.8):
        bad.append("缺深青实心圆")
    elif not _sel(sl, geom=("heart",), fill=("white",)):
        bad.append("深青圆内缺白色心形")
    if bad:
        return True, "；".join(bad)
    return False, "第11页三处顶部图标造型保留"


def n22(D):
    sl = _S(D, 11)
    got = [d for d in sl["txt"] if "变化" in d["text"] or "影响" in d["text"]]
    if not got:
        return True, "第11页未见“变化：/影响：”两段说明"
    bad = []
    for d in got:
        card = _card_of(sl, d) if sl["vec"] else None
        if _ovf(d, card):
            bad.append("(%.1f,%.1f)处说明溢出（估高%.2fcm/框高%.2fcm）" % (
                d["L"], d["T"], _est_h(d), d["H"]))
    if bad:
        return True, "；".join(bad[:4])
    return False, "第11页“变化/影响”两段说明均完整位于卡内"


def n23(D):
    sl = _S(D, 12)
    if not sl["vec"]:
        return False, _BGKEEP
    stars = _sel(sl, geom=("star5",))
    if len(stars) < 6:
        return True, "第12页五角星仅%d颗，星座连线不完整" % len(stars)
    ys = sorted(s["cy"] for s in stars)
    bands = [[ys[0]]]
    for y in ys[1:]:
        if y - bands[-1][-1] > 1.5:
            bands.append([])
        bands[-1].append(y)
    big = [b for b in bands if len(b) >= 3]
    conn = [d for d in sl["shapes"]
            if d["geom"] in ("line", "straightConnector1", "bentConnector3", "custom")
            and _fam(d["line"] if d["line"] not in ("none", None) else d["fill"]) in GOLD]
    if len(big) >= 2 and len(conn) >= 2:
        return False, "上下两组星座连线保留（%d组星带、%d段橙黄连线）" % (len(big), len(conn))
    return True, "星座连线不完整：星带%d组、橙黄连线%d段" % (len(big), len(conn))

def _ntxt(page, want, strict=False):
    """文案类减分项：命中（缺失/被改写）返回 True。"""
    def _fn(D):
        got = (_norm(want) in _S(D, page)["ntext"] if strict
               else _key(want) in _S(D, page)["ktext"])
        if got:
            return False, "第%d页文案与细则一致：“%s”" % (page, want)
        return True, "第%d页未找到“%s”" % (page, want)
    return _fn


n24 = _ntxt(2, "每个孩子都有独特光点，教育要做的是帮他们找到被看见的入口。")
n25 = _ntxt(6, "把正向行为转化为持续反馈")
n26 = _ntxt(7, "成效期待：学生在小任务中积累可感知的责任感，班级运行更顺畅。", True)
n27 = _ntxt(9, "根据性格与能力进行混合搭配，让每个小队都拥有表达者、倾听者、组织者和创意者。")
n28 = _ntxt(10, "把阶段成果变成真实体验")
n29 = _ntxt(11, "互助关系更柔和，团队归属感增强。")

_CHECKS = [p01, p02, p03, p04, p05, p06, p07, p08, p09, p10,
           p11, p12, p13, p14, p15, p16, p17, p18, p19, p20,
           n01, n02, n03, n04, n05, n06, n07, n08, n09, n10,
           n11, n12, n13, n14, n15, n16, n17, n18, n19, n20,
           n21, n22, n23, n24, n25, n26, n27, n28, n29]
assert len(_CHECKS) == len(_RULES) == len(_MAX_DELTAS_ORDER)


# ------------------------------------------------------------------- 维度1判定
# 维度1 只保留“交付文件为1个可正常打开的 .pptx”这一条：文件能被解析打开即通过。
# 页数不足、整页位图不可编辑、版面错乱等原先的一票否决条件已按细则删除，
# 相关表现全部交由维度2逐条加/减分体现。文件损坏打不开时 _build 会抛异常，
# 由 evaluate 捕获为 status=error、total_score=0。
def _dim1(D):
    return True, "文件为可正常打开的 .pptx，共%d页，%d处可单独点选编辑的文字" % (
        len(D["slides"]), sum(len(s["txt"]) for s in D["slides"]))

# --------------------------------------------------------------------- 主流程
def _evaluate_impl(path):
    D = _build(path)
    dim1_pass, dim1_reason = _dim1(D)
    results, details = [], []
    for k, (fn, md) in enumerate(zip(_CHECKS, _MAX_DELTAS_ORDER)):
        try:
            hit, detail = fn(D)
        except Exception as exc:                 # 单条检查异常不影响整体评测
            hit, detail = False, "检查异常：%s: %s" % (type(exc).__name__, exc)
        sign = "+" if md > 0 else "-"
        results.append("%s%d：%s" % (sign, abs(md) if hit else 0, _RULES[k]))
        details.append(detail)
    return dim1_pass, dim1_reason, results, details


class _TargetMissing(Exception):
    """目录不存在或目录中没有 .pptx：连待评测对象都没有，不进入维度判定。"""


def evaluate(dir_path):
    """入参为脚本所在目录；返回统一格式的评测结果字典。"""
    out = {
        "id": SCRIPT_ID, "file_name": "", "status": "ok", "error": "",
        "dim1_pass": False, "dim1_reason": "", "dim2_items": [],
        "total_score": 0, "max_score": MAX_SCORE_TOTAL,
    }
    try:
        if not dir_path or not os.path.isdir(dir_path):
            raise _TargetMissing("目录不存在：%s" % dir_path)
        files = sorted(f for f in os.listdir(dir_path)
                       if f.lower().endswith(".pptx") and not f.startswith("~"))
        if not files:
            raise _TargetMissing("目录中未找到 .pptx 文件：%s" % dir_path)
        out["file_name"] = files[0]
        p1, r1, results, details = _evaluate_impl(os.path.join(dir_path, files[0]))
        out["dim1_pass"], out["dim1_reason"] = p1, r1
        bonus = penalty = 0
        for raw, max_delta in zip(results, _MAX_DELTAS_ORDER):
            head, _, rule = raw.partition("：")
            delta = _i(head.replace("+", "").strip(), 0) or 0
            hit = delta != 0
            d = max_delta if hit else 0
            if max_delta > 0:
                bonus += d
            else:
                penalty += -d
            out["dim2_items"].append({"rule": rule, "max_delta": max_delta, "delta": d,
                                      "hit": hit})
        # 总分不做下限截断：扣分多于加分即为负分，取值范围 [MIN_SCORE_TOTAL, MAX_SCORE_TOTAL]。
        # 维度1 未通过时加分作废，但已命中的减分项照扣。
        out["total_score"] = (bonus - penalty) if p1 else -penalty
    except _TargetMissing as exc:
        # 找不到待评测文件：维度1 未参与判定，dim1_pass / dim1_reason 一律留空，满分也置 0
        out["file_name"] = ""
        out["status"] = "error"
        out["error"] = str(exc)
        out["dim1_pass"] = ""
        out["dim1_reason"] = ""
        out["dim2_items"] = []
        out["total_score"] = 0
        out["max_score"] = 0
    except Exception as exc:
        # 文件存在但打不开 / 解析失败：维度1 判不通过，dim1_reason 给出错误摘要
        out["status"] = "error"
        out["error"] = "%s: %s" % (type(exc).__name__, exc)
        out["dim1_pass"] = False
        out["dim1_reason"] = "文件无法解析或无法打开：%s: %s" % (type(exc).__name__, exc)
        out["dim2_items"] = []
        out["total_score"] = 0
        out["traceback"] = traceback.format_exc()
    return out


if __name__ == "__main__":
    _dir = sys.argv[1] if len(sys.argv) > 1 else os.path.dirname(os.path.abspath(__file__))
    _out = evaluate(_dir)
    print(json.dumps(_out, ensure_ascii=False, indent=2))
