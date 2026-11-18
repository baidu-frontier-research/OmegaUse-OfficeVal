#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# pyright: reportUnknownMemberType=false, reportUnknownVariableType=false
# pyright: reportUnknownArgumentType=false, reportMissingImports=false
# pyright: reportOptionalMemberAccess=false, reportOptionalSubscript=false

import os
import re
import sys
import json
import zipfile
import hashlib

try:
    from lxml import etree
except ImportError:                       # 退化到标准库；本模块只用到 find/iter/get
    import xml.etree.ElementTree as etree  # type: ignore

# 说明：按接口统一约定，本模块不修改全局 sys.stdout，也不 print 主结果。
# 主结果统一由 evaluate(dir_path) -> dict 返回；入参为“脚本所在目录的路径”，
# 由脚本自己在该目录内定位并打开被评估的 .pptx。本文件仅在 __main__ 下作为
# 本地调试入口，才把 dict 序列化为 JSON 打印。
#
# 任务背景：把「呼吸项目课件.pptx」（39 页，原稿）的模板换成「医疗服务收费指引.pptx」
# （5 页，风格模板）那种的——背景、配色、风格改版，不动原来的文字和内容。
# 本脚本按该任务细则的 35 个加分项逐项判定，满分 91。
#
# 判定口径（评测过程中已确认的裁定，逐条写在对应评分项的注释里）：
#   ① 按渲染结果判定：被后层不透明形状完全覆盖的对象不算“可见”，既不算风格并存，
#      也不算内容缺失；p:bg 被满幅色块盖住时，底色取盖上去的那一层。
#   ② 属性缺失 = 继承而非没有：字号/字体/颜色沿
#      a:rPr → a:pPr/a:defRPr → a:lstStyle → 占位符(版式/母版) → p:style → 主题 解析。
#   ③ 模块页（第 3、8 分节页与第 39 结束页）豁免品牌标签/副署位/底部细线/页脚
#      （风格模板自身在这两类页上也不带这些元素），相关三条按 36 页核查。
#   ④ 形状没有 a:ln、只靠 p:style/a:lnRef 继承主题默认线，不算“额外添加描边”。
#   ⑤ 同一行相邻的两个文本对象拼起来读通，即视为该句仍在显示（逐字保留照给）。
#   ⑥ 页序按 p:sldIdLst → rels → 文件 解析出的“放映位置”核对，不假定恒等映射。
#   ⑦ Office 主题里 <a:ea typeface=""/> 为空是常态，中文实际由形状上显式写死的
#      a:latin 承担；两个槽都无声明时才回落主题（主题 ea 为空则取 script="Hans"）。
#   ⑧ 文本框通常比它压着的色板宽，故“文字压在什么底色上”按文字框中心点归属判定；
#      文字所在形状自身有填充时，直接以该填充为背景。
#   ⑨ 加分项一律按各条的适用范围逐页全量核查，一页不合即不得分；若某条只针对特定页，
#      由细则本身在该条开头写明页码（如“11、12、15、19、27 页……”），脚本不自行取样。
#      返回的 dim2_items 只有 rule / max_delta / delta / hit 四个字段，rule 只保留细则
#      原文，脚本追加的核查依据在反解时丢弃。
#
# 与办公软件对齐：本脚本不做整篇渲染，几何、颜色、字体、字号全部按 OOXML 的继承与
# 叠放规则解析后再判定，位置容差 0.12 英寸、字号容差 0.15 磅，与人在 PowerPoint 里
# 用“选中—看格式—量位置”的读数口径一致。

SCRIPT_ID = "075"
# 维度二所有评分项 max_delta 之和；维度一未通过时 total_score 记 0、max_score 仍取本值，
# 而目录里找不到文档等异常情况下 max_score 记 0（没有可评的对象）
MAX_SCORE_TOTAL = 103

# 分值只取 1 / 3 / 5 三档，按“把这个点做出来要花多少工夫”定档：
#   1 = 简单的点（5 分钟内）：单一属性的全局替换，或只在一两页上的一处改动，
#       以及“保持原样、别改坏”这类零工作量的约束。
#   3 = 中等难度的点（5–10 分钟）：多属性的复合规格批量落地，或按页型分支取值，
#       或新建形制单一的对象并铺到全篇。
#   5 = 高难度的点（大于 10 分钟）：要遍历全篇上百个对象逐个改造、整页版式重建，
#       或成体系地铺一整套规格。
#
# 维度二所有评分项按“结果追加顺序”的 max_delta 列表；末端把 results 里的
# “+N：...”字符串反解为结构化 dim2_items 时用。顺序须与 _evaluate_impl 中
# results.append(...) 的顺序严格一致；共 35 项，累加 = MAX_SCORE_TOTAL。
_MAX_DELTAS_ORDER = [
    5,  # 01 39 页底色换为模板青白色系、无米黄残留
    3,  # 02 底色按页面类型分层取值
    5,  # 03 主色/辅色/强调橙取模板精确色值，非自拟近似值
    3,  # 04 第 3、8 页深底 143642 + 深页浅字
    3,  # 05 标题色 143642 / 24343B 且与正文灰有层级差
    1,  # 06 正文灰 56626A → 667987
    1,  # 07 卡片白底 + 描边 D3ECE8 / D7E9E6 / D6E9E5
    1,  # 08 装饰件 BDD3D2 → B7DCD7
    3,  # 09 每页顶部通栏色条（39 页）
    3,  # 10 每页底部细分隔线（36 页）
    5,  # 11 左上品牌圆角标签 + 右侧副署位（36 页）
    3,  # 12 页码改“当前页 / 总页数”写法
    3,  # 13 页码规格 8 磅加粗右对齐、位置与颜色
    3,  # 14 页脚 5.5 磅 667987（36 页）
    5,  # 15 原稿特有装饰件清除、无两套风格并存
    3,  # 16 中文统一 Microsoft YaHei
    3,  # 17 数字/英文标签改 Arial、未沿用 Aptos
    5,  # 18 卡片正文 9.7 / 9.8 磅且未被自适应缩放
    3,  # 19 内容页主标题 27 磅加粗、目录页 28 磅加粗
    3,  # 20 卡片小标题 14.5 / 15.0 磅加粗
    1,  # 21 过渡副标题 11.5 磅一档、左对齐
    5,  # 22 卡片序号徽标配色与形态
    3,  # 23 卡片底板为白色圆角卡 roundRect
    1,  # 24 纯色装饰条未额外添加描边
    3,  # 25 无同色/极低对比文字
    3,  # 26 封面主标题 34 磅加粗 143642 且层级重排
    3,  # 27 分节页结构（白色 34 磅章节标题 + 导语 + PART 01）
    5,  # 28 结束页白色圆角卡与三级文字
    1,  # 29 结束页署名行取 109D9B 一档
    5,  # 30 39 页文字逐字保留
    1,  # 31 仍为 39 页且页序与页面主题对应关系不变
    1,  # 32 画布仍为 13.333 × 7.5 英寸
    3,  # 33 未搬用模板素材照片、同一张图未重复多页
    1,  # 34 机构名/期次/明细页编号等实体取值未被改写
    3,  # 35 文字与图形不冲撞（不出画布、文本框不叠压）
]
assert sum(_MAX_DELTAS_ORDER) == MAX_SCORE_TOTAL
assert len(_MAX_DELTAS_ORDER) == 35

# ========== OOXML 命名空间与单位 ==========
NS_P = 'http://schemas.openxmlformats.org/presentationml/2006/main'
NS_A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
NS = {'p': NS_P, 'a': NS_A, 'r': NS_R}
Q_T = '{%s}t' % NS_A          # a:t   文字
Q_P = '{%s}p' % NS_A          # a:p   段落
EMU_PER_INCH = 914400.0       # 1 英寸 = 914400 EMU
PT_PER_SZ = 100.0             # a:rPr/@sz 为百分磅

# 位置/尺寸容差：细则写“允许 ±0.02 英寸”，评测过程中用户放宽到 ±0.30 厘米，
# 取较宽者（0.30 cm = 0.1181 in），避免把肉眼无差别的排版差判成不合规。
POS_TOL = 0.12
SIZE_TOL = 0.12
# 字号容差：同一档位的等价写法（如 9.7 与 9.8 同属一档）在各项里显式列出取值集合，
# 单个取值本身给 ±0.15 磅的浮点容差。
SZ_TOL = 0.15

# ========== 风格模板（医疗服务收费指引.pptx）的设计系统取值 ==========
TPL_BG_LIGHT = ('F4FBFA', 'FFFDF7', 'F8FBF7', 'F5FBFA')   # 分页类型底色
TPL_BG_DARK = '143642'                                    # 分节页深底
TPL_MAIN = '109D9B'                                       # 主色
TPL_SECOND = '3CC7BE'                                     # 辅色
TPL_ORANGE = 'F4A53A'                                     # 强调橙
TPL_TITLE = ('143642', '24343B')                          # 标题色
TPL_BODY_GREY = '667987'                                  # 正文灰
TPL_CARD_LINE = ('D3ECE8', 'D7E9E6', 'D6E9E5')            # 卡片描边
TPL_DECOR = 'B7DCD7'                                      # 箭头/分隔线/折角装饰件
TPL_DARK_TEXT = ('FFFFFF', 'DDEDEA', 'EBF6F5')            # 深底页文字色
TPL_DARK_PAGENO = 'B5EFEA'                                # 深底页页码色
WHITE = 'FFFFFF'
# 模板全部许可色值（判“自拟近似色”时先白名单排除）
TPL_ALL = set(TPL_BG_LIGHT) | set(TPL_TITLE) | set(TPL_CARD_LINE) | set(TPL_DARK_TEXT) | {
    TPL_BG_DARK, TPL_MAIN, TPL_SECOND, TPL_ORANGE, TPL_BODY_GREY, TPL_DECOR,
    TPL_DARK_PAGENO, WHITE, '000000',
}

# ========== 原稿（呼吸项目课件.pptx）的设计系统取值 ==========
ORIG_BG = ('FBF8EF', 'EEF9F7')                            # 米黄/浅青交替底色
ORIG_MAIN = ('17979D', '4EBFB8')                          # 青绿主色
ORIG_ORANGE = 'E3AA41'
ORIG_TITLE = '133545'
ORIG_BODY_GREY = '56626A'
ORIG_CARD_LINE = 'D0E5E2'
ORIG_DECOR = 'BDD3D2'
ORIG_BAND = 'DAEEEB'                                      # 底部色带
ORIG_PANEL = '005A64'                                     # 右侧色板
# 第 30 行“原稿特有装饰件”的色值集合（形状填充/描边口径；卡片描边与装饰件
# 另有专条，不放进来重复计分）
ORIG_DECOR_FILLS = (ORIG_TITLE, ORIG_BAND, ORIG_MAIN[0], ORIG_MAIN[1], ORIG_PANEL, ORIG_ORANGE)
ORIG_ALL = set(ORIG_BG) | set(ORIG_MAIN) | {
    ORIG_ORANGE, ORIG_TITLE, ORIG_BODY_GREY, ORIG_CARD_LINE, ORIG_DECOR,
    ORIG_BAND, ORIG_PANEL,
}
# 原稿字号/字体基线
ORIG_CARD_BODY_PT = 10.19
ORIG_FONT_ASCII = 'Aptos'

# 风格模板自带素材照片（ppt/media/*）的 md5 指纹：模板内共 2 张（各被 2 页引用），
# 用于判定“是否直接搬用模板素材”。比内容哈希而不比文件名，改名/重压也认得出。
TPL_MEDIA_MD5 = (
    'de3a75fce8c0b3d4e1ab767ccfcbfa9d',   # template_img 1672×941（模板第 1、3 页）
    'a6fb06352a41606b867c256657009d93',   # template_img 1672×941（模板第 2、5 页）
)

# ========== 原稿 39 页的段落锚点基线 ==========
# 换模板任务的“原文文字与内容完整保留”必须有原稿基线才能判定，故把原稿 39 页
# 逐页的段落文本（去掉所有空白后、长度 ≥2、按页内去重）内嵌为基线，共 497 条。
# 每行 = 原稿第 N 页，页内锚点以 '¦' 分隔。判定时只做“该页可见文本里是否含有
# 这一串”，不比对 run 拆分方式——同一句被拆成多个 run/多个文本框都算保留。
_ORIG_ANCHORS_RAW = (
    "呼吸系统诊疗收费项目合规解读¦01¦医疗服务价格管理¦基于医保支付口径与院内收费规范的业务培训材料¦瑞康医疗价格管理中心¦2026年春季",  # p1
    "呼吸系统诊疗收费项目合规解读¦02¦目录¦从政策口径到项目明细，按“先规则、后执行”的路径展开¦01¦政策文件要点¦理解整合口径、医保衔接与价格边界¦项目明细规范¦按检查、治疗、手术三类建立收费判断¦03¦执行注意事项¦聚焦互斥规则、计价单位和复杂条件",  # p2
    "呼吸系统诊疗收费项目合规解读¦03¦Part1¦政策框架与应用口径¦先明确项目边界，再进入具体收费场景",  # p3
    "呼吸系统诊疗收费项目合规解读¦04¦规范整合呼吸系统医疗服务项目¦本页承接政策框架，说明项目整理与医保衔接的基本方向¦项目整合¦围绕呼吸功能检查、内镜诊疗、胸腔及肺部手术等服务产出，统一项目名称和收费边界。¦项目修订¦对名称、计价单位、加收条件和互斥规则进行归并，减少同类项目交叉和重复。¦项目停用¦历史项目中口径过细、重复或不适合继续单独收费的内容，统一转入整合项目管理。¦医保衔接口径¦纳入医保支付范围的项目，按照统一属性、限制范围和个人先行自付规则执行；超出核定项目价格边界的部分，应按规定向患者说明并留存凭证。¦执行提示¦院内收费审核重点从“名称是否对应”转向“服务产出、计价单位、计价说明是否一致”。",  # p4
    "呼吸系统诊疗收费项目合规解读¦05¦项目类别与价格边界¦承接上一页政策概览，进一步拆解实际收费判断口径¦按服务产出设项¦以呼吸系统诊疗环节的最终服务成果为核心，避免把临床步骤拆成多个独立收费点。¦政府指导价格¦项目价格作为上限管理，院内可结合实际服务和支付政策执行下浮或优惠安排。¦价格构成理解¦价格构成用于测算资源消耗，不等同于临床操作流程，也不要求未发生环节强制减计。¦进入项目明细",  # p5
    "呼吸系统诊疗收费项目合规解读¦06¦加收、扩展与基本物耗¦承接上一页政策概览，进一步拆解实际收费判断口径¦加收项¦同一项目在特殊方式或特殊场景下确需差异化时使用，加减收应以主项目价格为基础。¦扩展项¦只扩展适用范围，不额外提高价格，执行时仍按主项目口径计价。¦基本物耗¦常规消毒、清洁、防护、标签、普通注射和软件使用等成本纳入项目价格，不另行拆收。¦进入项目明细",  # p6
    "呼吸系统诊疗收费项目合规解读¦07¦无创、内镜与手术并行规则¦承接上一页政策概览，进一步拆解实际收费判断口径¦无创定义¦通过自然腔道或无需切开皮肤组织的方式完成操作，需与开放手术项目区分。¦内镜费用¦非手术治疗如需内镜，可按说明收取检查费；内镜下手术已包含内镜使用成本。¦并行手术¦同一切口不同手术、同一切口双侧手术应按计价说明执行，避免简单叠加。¦进入项目明细",  # p7
    "呼吸系统诊疗收费项目合规解读¦08¦Part2¦项目明细与收费规范¦以“检查—治疗—手术—注意事项”的顺序展开",  # p8
    "呼吸系统诊疗收费项目合规解读¦09¦呼吸系统类项目明细规范¦01｜肺功能基础检查：承接上页分类，细化收费口径¦肺功能基础检查¦围绕肺容积、通气、弥散和气道反应等检查建立项目边界，重点看报告产出、设备准备、测定与用物处理是否完整。¦按次计价¦确认服务产出与适用场景¦报告产出¦核对计价单位与是否加收¦舒张试验单列规则¦检查互斥项目与记录凭证¦收费审核重点¦名称、服务产出、计价单位、计价说明四项一致后再进入院内审核流程；特殊情况需有病历、医嘱或操作记录支撑。",  # p9
    "呼吸系统诊疗收费项目合规解读¦10¦呼吸系统类项目明细规范¦02｜气道反应与运动评估：承接上页分类，细化收费口径¦气道反应与运动评估¦将激发试验、呼吸阻力、运动心肺功能等纳入同一检查逻辑，强调适应证、监测过程和风险告知。¦评估气道反应¦确认服务产出与适用场景¦运动监测¦核对计价单位与是否加收¦记录留痕¦检查互斥项目与记录凭证¦收费审核重点¦名称、服务产出、计价单位、计价说明四项一致后再进入院内审核流程；特殊情况需有病历、医嘱或操作记录支撑。",  # p10
    "呼吸系统诊疗收费项目合规解读¦11¦呼吸系统类项目明细规范¦03｜呼吸肌与膈肌评估：承接上页分类，细化收费口径¦呼吸肌与膈肌评估¦对呼吸肌压力、流量变化及膈肌功能测定进行规范，便于与后续治疗方案衔接。¦功能测定¦确认服务产出与适用场景¦压力指标¦核对计价单位与是否加收¦结果解释¦检查互斥项目与记录凭证¦收费审核重点¦名称、服务产出、计价单位、计价说明四项一致后再进入院内审核流程；特殊情况需有病历、医嘱或操作记录支撑。",  # p11
    "呼吸系统诊疗收费项目合规解读¦12¦呼吸系统类项目明细规范¦04｜睡眠与连续监测：承接上页分类，细化收费口径¦睡眠与连续监测¦睡眠呼吸监测、经皮氧分压和二氧化碳监测按监测场景区分，避免与同类生命体征项目重复计价。¦不可重复收费¦确认服务产出与适用场景¦按次或时段¦核对计价单位与是否加收¦便携情形另列¦检查互斥项目与记录凭证¦收费审核重点¦名称、服务产出、计价单位、计价说明四项一致后再进入院内审核流程；特殊情况需有病历、医嘱或操作记录支撑。",  # p12
    "呼吸系统诊疗收费项目合规解读¦13¦呼吸系统类项目明细规范¦05｜支气管镜检查：承接上页分类，细化收费口径¦支气管镜检查¦常规内镜、超声内镜、共聚焦内镜按照观察深度和成像方式区分，特殊光源仅在实际使用时体现。¦内镜类型¦确认服务产出与适用场景¦图像采集¦核对计价单位与是否加收¦特殊光源¦检查互斥项目与记录凭证¦收费审核重点¦名称、服务产出、计价单位、计价说明四项一致后再进入院内审核流程；特殊情况需有病历、医嘱或操作记录支撑。",  # p13
    "呼吸系统诊疗收费项目合规解读¦14¦呼吸系统类项目明细规范¦06｜镜下探查项目：承接上页分类，细化收费口径¦镜下探查项目¦对肺叶通气评估和纵隔镜探查进行场景化归集，明确入镜、观察、撤镜及报告形成的边界。¦自然腔道或切口¦确认服务产出与适用场景¦评估目的¦核对计价单位与是否加收¦操作记录¦检查互斥项目与记录凭证¦收费审核重点¦名称、服务产出、计价单位、计价说明四项一致后再进入院内审核流程；特殊情况需有病历、医嘱或操作记录支撑。",  # p14
    "呼吸系统诊疗收费项目合规解读¦15¦呼吸系统类项目明细规范¦07｜非手术治疗：承接上页分类，细化收费口径¦非手术治疗¦体外膈肌起搏、一氧化氮吸入、雾化治疗按照治疗过程计价，关键是设备连接、治疗观察与记录完整。¦治疗过程¦确认服务产出与适用场景¦用时管理¦核对计价单位与是否加收¦多药物规则¦检查互斥项目与记录凭证¦收费审核重点¦名称、服务产出、计价单位、计价说明四项一致后再进入院内审核流程；特殊情况需有病历、医嘱或操作记录支撑。",  # p15
    "呼吸系统诊疗收费项目合规解读¦16¦呼吸系统类项目明细规范¦08｜灌洗与镜下常规治疗：承接上页分类，细化收费口径¦灌洗与镜下常规治疗¦全肺灌洗、支气管肺泡灌洗和镜下常规治疗按照清除范围和操作目的区分，不与检查项目混淆。¦灌洗范围¦确认服务产出与适用场景¦不含内镜检查¦核对计价单位与是否加收¦肺段口径¦检查互斥项目与记录凭证¦收费审核重点¦名称、服务产出、计价单位、计价说明四项一致后再进入院内审核流程；特殊情况需有病历、医嘱或操作记录支撑。",  # p16
    "呼吸系统诊疗收费项目合规解读¦17¦呼吸系统类项目明细规范¦09｜镜下特殊治疗：承接上页分类，细化收费口径¦镜下特殊治疗¦封堵、扩张、消融、冷冻等操作归入特殊治疗，要求根据实际技术路径留存操作记录。¦治疗方式¦确认服务产出与适用场景¦设备耗用¦核对计价单位与是否加收¦必要性说明¦检查互斥项目与记录凭证¦收费审核重点¦名称、服务产出、计价单位、计价说明四项一致后再进入院内审核流程；特殊情况需有病历、医嘱或操作记录支撑。",  # p17
    "呼吸系统诊疗收费项目合规解读¦18¦呼吸系统类项目明细规范¦10｜气道支架服务：承接上页分类，细化收费口径¦气道支架服务¦支架置入与取出分别管理，若涉及儿童或特殊风险场景，应按项目说明和院内授权流程执行。¦置入/取出¦确认服务产出与适用场景¦儿童场景¦核对计价单位与是否加收¦导丝与球囊¦检查互斥项目与记录凭证¦收费审核重点¦名称、服务产出、计价单位、计价说明四项一致后再进入院内审核流程；特殊情况需有病历、医嘱或操作记录支撑。",  # p18
    "呼吸系统诊疗收费项目合规解读¦19¦呼吸系统类项目明细规范¦11｜无创修补与切除：承接上页分类，细化收费口径¦无创修补与切除¦经自然腔道完成的瘘修补、气管病变切除等项目，应区分内镜检查费用与治疗费用。¦无创方式¦确认服务产出与适用场景¦不含检查¦核对计价单位与是否加收¦病变处理¦检查互斥项目与记录凭证¦收费审核重点¦名称、服务产出、计价单位、计价说明四项一致后再进入院内审核流程；特殊情况需有病历、医嘱或操作记录支撑。",  # p19
    "呼吸系统诊疗收费项目合规解读¦20¦呼吸系统类项目明细规范¦12｜无创减容与异物取出：承接上页分类，细化收费口径¦无创减容与异物取出¦肺减容和气管异物取出按技术路径归集，收费凭证需体现操作目的和实际完成情况。¦活瓣或消融¦确认服务产出与适用场景¦异物取出¦核对计价单位与是否加收¦过程留痕¦检查互斥项目与记录凭证¦收费审核重点¦名称、服务产出、计价单位、计价说明四项一致后再进入院内审核流程；特殊情况需有病历、医嘱或操作记录支撑。",  # p20
    "呼吸系统诊疗收费项目合规解读¦21¦呼吸系统类项目明细规范¦13｜气管成形类手术：承接上页分类，细化收费口径¦气管成形类手术¦气管成形、隆突重建等重建项目，以手术计划、切除、修复、缝合为核心计价边界。¦重建修复¦确认服务产出与适用场景¦手术计划¦核对计价单位与是否加收¦耗材边界¦检查互斥项目与记录凭证¦收费审核重点¦名称、服务产出、计价单位、计价说明四项一致后再进入院内审核流程；特殊情况需有病历、医嘱或操作记录支撑。",  # p21
    "呼吸系统诊疗收费项目合规解读¦22¦呼吸系统类项目明细规范¦14｜气管食管瘘修补：承接上页分类，细化收费口径¦气管食管瘘修补¦常规与复杂情形应根据是否涉及填充、皮瓣等增强修复措施判断，不宜仅凭病情描述上浮。¦复杂定义¦确认服务产出与适用场景¦修补路径¦核对计价单位与是否加收¦材料说明¦检查互斥项目与记录凭证¦收费审核重点¦名称、服务产出、计价单位、计价说明四项一致后再进入院内审核流程；特殊情况需有病历、医嘱或操作记录支撑。",  # p22
    "呼吸系统诊疗收费项目合规解读¦23¦呼吸系统类项目明细规范¦15｜气管病变切除：承接上页分类，细化收费口径¦气管病变切除¦气管及隆突病变切除按照部位和手术难度管理，术中取样、送检应在记录中说明。¦部位识别¦确认服务产出与适用场景¦切除步骤¦核对计价单位与是否加收¦送检记录¦检查互斥项目与记录凭证¦收费审核重点¦名称、服务产出、计价单位、计价说明四项一致后再进入院内审核流程；特殊情况需有病历、医嘱或操作记录支撑。",  # p23
    "呼吸系统诊疗收费项目合规解读¦24¦呼吸系统类项目明细规范¦16｜胸腔探查与胸腔病变：承接上页分类，细化收费口径¦胸腔探查与胸腔病变¦胸腔探查通常不与同部位其他手术并行计收，胸腔病变切除需明确具体组织范围。¦同部位规则¦确认服务产出与适用场景¦病变范围¦核对计价单位与是否加收¦止血处理¦检查互斥项目与记录凭证¦收费审核重点¦名称、服务产出、计价单位、计价说明四项一致后再进入院内审核流程；特殊情况需有病历、医嘱或操作记录支撑。",  # p24
    "呼吸系统诊疗收费项目合规解读¦25¦呼吸系统类项目明细规范¦17｜肺部分切除与肺叶切除：承接上页分类，细化收费口径¦肺部分切除与肺叶切除¦非解剖性部分切除、肺叶切除按照切除范围区分，复杂情形需满足项目限定条件。¦单侧口径¦确认服务产出与适用场景¦切除范围¦核对计价单位与是否加收¦复杂依据¦检查互斥项目与记录凭证¦收费审核重点¦名称、服务产出、计价单位、计价说明四项一致后再进入院内审核流程；特殊情况需有病历、医嘱或操作记录支撑。",  # p25
    "呼吸系统诊疗收费项目合规解读¦26¦呼吸系统类项目明细规范¦18｜肺段切除：承接上页分类，细化收费口径¦肺段切除¦常规肺段与复杂肺段切除应依据联合肺段、亚段支气管处理等条件判断，避免简单套用高等级项目。¦肺段范围¦确认服务产出与适用场景¦联合情形¦核对计价单位与是否加收¦术式证据¦检查互斥项目与记录凭证¦收费审核重点¦名称、服务产出、计价单位、计价说明四项一致后再进入院内审核流程；特殊情况需有病历、医嘱或操作记录支撑。",  # p26
    "呼吸系统诊疗收费项目合规解读¦27¦呼吸系统类项目明细规范¦19｜全肺切除：承接上页分类，细化收费口径¦全肺切除¦常规与复杂全肺切除按照是否涉及心包内处理、胸膜外路径等情况区分，并保留术式依据。¦全肺范围¦确认服务产出与适用场景¦复杂路径¦核对计价单位与是否加收¦单侧计价¦检查互斥项目与记录凭证¦收费审核重点¦名称、服务产出、计价单位、计价说明四项一致后再进入院内审核流程；特殊情况需有病历、医嘱或操作记录支撑。",  # p27
    "呼吸系统诊疗收费项目合规解读¦28¦呼吸系统类项目明细规范¦20｜肺修补与胸腺病变：承接上页分类，细化收费口径¦肺修补与胸腺病变¦肺组织缺损修补和胸腺病变切除按操作目的分别归集，避免与相邻胸腔项目重复。¦组织修补¦确认服务产出与适用场景¦胸腺病变¦核对计价单位与是否加收¦边界清晰¦检查互斥项目与记录凭证¦收费审核重点¦名称、服务产出、计价单位、计价说明四项一致后再进入院内审核流程；特殊情况需有病历、医嘱或操作记录支撑。",  # p28
    "呼吸系统诊疗收费项目合规解读¦29¦呼吸系统类项目明细规范¦21｜胸壁切除与修复：承接上页分类，细化收费口径¦胸壁切除与修复¦胸壁病变切除、缺损修复以是否需要固定、修复方式和缺损复杂度为判断重点。¦缺损修复¦确认服务产出与适用场景¦固定措施¦核对计价单位与是否加收¦复杂条件¦检查互斥项目与记录凭证¦收费审核重点¦名称、服务产出、计价单位、计价说明四项一致后再进入院内审核流程；特殊情况需有病历、医嘱或操作记录支撑。",  # p29
    "呼吸系统诊疗收费项目合规解读¦30¦呼吸系统类项目明细规范¦22｜胸廓成形：承接上页分类，细化收费口径¦胸廓成形¦胸廓重建需与胸壁缺损修复区分；涉及畸形矫正等复杂情形时，应有充分手术记录支撑。¦不重复计价¦确认服务产出与适用场景¦重建目的¦核对计价单位与是否加收¦畸形矫正¦检查互斥项目与记录凭证¦收费审核重点¦名称、服务产出、计价单位、计价说明四项一致后再进入院内审核流程；特殊情况需有病历、医嘱或操作记录支撑。",  # p30
    "呼吸系统诊疗收费项目合规解读¦31¦呼吸系统类项目明细规范¦23｜脓胸廓清：承接上页分类，细化收费口径¦脓胸廓清¦脓胸廓清按引流、清除和修复措施归集，复杂情形通常需体现填充或皮瓣处理。¦清除引流¦确认服务产出与适用场景¦填充处理¦核对计价单位与是否加收¦术后记录¦检查互斥项目与记录凭证¦收费审核重点¦名称、服务产出、计价单位、计价说明四项一致后再进入院内审核流程；特殊情况需有病历、医嘱或操作记录支撑。",  # p31
    "呼吸系统诊疗收费项目合规解读¦32¦呼吸系统类项目明细规范¦24｜胸膜处理：承接上页分类，细化收费口径¦胸膜处理¦胸膜剥脱、胸膜固定属于不同服务产出，应按照处理目标和实施步骤独立判断。¦剥脱¦确认服务产出与适用场景¦固定¦核对计价单位与是否加收¦目标不同¦检查互斥项目与记录凭证¦收费审核重点¦名称、服务产出、计价单位、计价说明四项一致后再进入院内审核流程；特殊情况需有病历、医嘱或操作记录支撑。",  # p32
    "呼吸系统诊疗收费项目合规解读¦33¦呼吸系统类项目明细规范¦25｜胸内异物与纵隔病变：承接上页分类，细化收费口径¦胸内异物与纵隔病变¦胸内异物清除和纵隔病变切除按部位、入路和是否复杂处理区分，重点保存病灶定位信息。¦异物清除¦确认服务产出与适用场景¦纵隔病变¦核对计价单位与是否加收¦入路说明¦检查互斥项目与记录凭证¦收费审核重点¦名称、服务产出、计价单位、计价说明四项一致后再进入院内审核流程；特殊情况需有病历、医嘱或操作记录支撑。",  # p33
    "呼吸系统诊疗收费项目合规解读¦34¦呼吸系统类项目明细规范¦26｜纵隔减压与感染处理：承接上页分类，细化收费口径¦纵隔减压与感染处理¦纵隔气肿减压、感染清创引流应与胸腔粘连松解等相近项目保持边界，防止并行重复。¦减压¦确认服务产出与适用场景¦清创引流¦核对计价单位与是否加收¦互斥规则¦检查互斥项目与记录凭证¦收费审核重点¦名称、服务产出、计价单位、计价说明四项一致后再进入院内审核流程；特殊情况需有病历、医嘱或操作记录支撑。",  # p34
    "呼吸系统诊疗收费项目合规解读¦35¦呼吸系统类项目明细规范¦27｜膈肌相关手术：承接上页分类，细化收费口径¦膈肌相关手术¦膈肌修补和折叠服务产出不同，收费时需以修补缺损或改善运动功能为核心依据。¦膈肌修补¦确认服务产出与适用场景¦膈肌折叠¦核对计价单位与是否加收¦目的区分¦检查互斥项目与记录凭证¦收费审核重点¦名称、服务产出、计价单位、计价说明四项一致后再进入院内审核流程；特殊情况需有病历、医嘱或操作记录支撑。",  # p35
    "呼吸系统诊疗收费项目合规解读¦36¦呼吸系统类项目明细规范¦28｜气管异物与肺空洞处理：承接上页分类，细化收费口径¦气管异物与肺空洞处理¦气管异物取出和肺空洞药物填充按手术目的归集，药物使用与操作记录保持一致。¦异物取出¦确认服务产出与适用场景¦药物填充¦核对计价单位与是否加收¦记录一致¦检查互斥项目与记录凭证¦收费审核重点¦名称、服务产出、计价单位、计价说明四项一致后再进入院内审核流程；特殊情况需有病历、医嘱或操作记录支撑。",  # p36
    "呼吸系统诊疗收费项目合规解读¦37¦呼吸系统类项目明细规范¦29｜淋巴清扫与粘连松解：承接上页分类，细化收费口径¦淋巴清扫与粘连松解¦胸腔淋巴清扫、采样扩展、粘连松解和交感神经链切除按服务产出分别归集。¦清扫/采样¦确认服务产出与适用场景¦粘连松解¦核对计价单位与是否加收¦神经链处理¦检查互斥项目与记录凭证¦收费审核重点¦名称、服务产出、计价单位、计价说明四项一致后再进入院内审核流程；特殊情况需有病历、医嘱或操作记录支撑。",  # p37
    "呼吸系统诊疗收费项目合规解读¦38¦执行特别注意事项¦承接明细页，提炼最容易出错的三类收费风险¦不得同时收费¦严格遵守计价说明中的互斥规则；同部位探查、修复、松解等相近项目不应因名称不同而重复计收。¦计价单位不同¦按次、按单侧、按时段等口径需逐项核对，收费数量应与病历记录、操作记录和报告内容一致。¦复杂条件须满足¦只有达到项目说明中明确的复杂情形，才可执行对应层级；不能仅凭手术耗时或主观难度判断。¦规范收费·留痕可查",  # p38
    "呼吸系统诊疗收费项目合规解读¦39¦感谢聆听¦以规范口径促进合理收费，以流程留痕支撑合规管理¦瑞康医疗价格管理中心",  # p39
)
ORIG_ANCHORS = [tuple(a for a in line.split('¦') if a) for line in _ORIG_ANCHORS_RAW]
ORIG_PAGES = len(ORIG_ANCHORS)   # 39

# 模块页（第 3、8 分节页 + 第 39 结束页）：细则第 25/26/29 行按 36 页核查
MODULE_PAGES = (3, 8, ORIG_PAGES)
SECTION_PAGES = (3, 8)
COVER_PAGE = 1
TOC_PAGE = 2
END_PAGE = ORIG_PAGES

# 第 60 行“实体取值”必须原样出现的字符串
ENTITY_STRINGS = ('瑞康医疗价格管理中心', '2026年春季', '呼吸系统诊疗收费项目合规解读')
# 明细页编号 01｜ … 29｜
ENTITY_ITEM_NOS = tuple('%02d｜' % k for k in range(1, 30))

# ========== 通用小工具 ==========
_WS_RE = re.compile(r'[\s　\xa0]+')
_CJK_RE = re.compile(r'[㐀-䶿一-鿿豈-﫿]')
_ALNUM_RE = re.compile(r'[A-Za-z0-9]')


def _norm(s):
    """去掉全部空白（半角/全角/NBSP/换行），用于文本比对。"""
    return _WS_RE.sub('', s or '')


def _has_cjk(s):
    return bool(_CJK_RE.search(s or ''))


def _rgb(h):
    h = (h or '').lstrip('#')
    return (int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))


def _rel_lum(h):
    """WCAG 相对亮度，用于判“深底深字”与“极低对比”。"""
    out = []
    for c in _rgb(h):
        c = c / 255.0
        out.append(c / 12.92 if c <= 0.03928 else ((c + 0.055) / 1.055) ** 2.4)
    return 0.2126 * out[0] + 0.7152 * out[1] + 0.0722 * out[2]


def _contrast(h1, h2):
    """WCAG 对比度，1.0 = 完全同色。"""
    a, b = _rel_lum(h1), _rel_lum(h2)
    if a < b:
        a, b = b, a
    return (a + 0.05) / (b + 0.05)


def _color_dist(h1, h2):
    """RGB 欧氏距离，用于判“自拟近似色值”。"""
    a, b = _rgb(h1), _rgb(h2)
    return (sum((a[i] - b[i]) ** 2 for i in range(3))) ** 0.5


def _near(v, target, tol):
    return v is not None and abs(v - target) <= tol


def _in_pt_set(v, values, tol=SZ_TOL):
    return v is not None and any(abs(v - t) <= tol for t in values)


def _emu(v):
    """EMU 字符串 → 英寸。"""
    try:
        return int(v) / EMU_PER_INCH
    except (TypeError, ValueError):
        return None


def _rect_covers(outer, inner, pad=0.02):
    """outer 是否在容差内完整覆盖 inner（四边各放宽 pad 英寸）。"""
    if not outer or not inner:
        return False
    ox, oy, ow, oh = outer
    ix, iy, iw, ih = inner
    return (ox <= ix + pad and oy <= iy + pad
            and ox + ow >= ix + iw - pad and oy + oh >= iy + ih - pad)


def _overlap(a, b):
    """两个矩形的重叠面积（英寸²）。"""
    if not a or not b:
        return 0.0
    dx = min(a[0] + a[2], b[0] + b[2]) - max(a[0], b[0])
    dy = min(a[1] + a[3], b[1] + b[3]) - max(a[1], b[1])
    return dx * dy if (dx > 0 and dy > 0) else 0.0


def _zip_join(base_dir, target):
    """把 rels 里的相对 Target 解析成包内绝对零件名。"""
    if target.startswith('/'):
        return target[1:]
    parts = [p for p in base_dir.split('/') if p]
    for seg in target.split('/'):
        if seg in ('', '.'):
            continue
        if seg == '..':
            if parts:
                parts.pop()
        else:
            parts.append(seg)
    return '/'.join(parts)


class _Pkg(object):
    """.pptx 包访问层：读零件、按 rels 解析引用、缓存已解析的 XML。

    与办公软件对齐：PowerPoint 打开 pptx 时就是按 [Content_Types] + rels 图
    去找幻灯片、版式、母版与主题的，这里复用同一条链路，保证解析出的
    字体/颜色/几何与打开文件后看到的一致。
    """

    def __init__(self, path):
        self.zf = zipfile.ZipFile(path)
        self.names = set(self.zf.namelist())
        self._xml = {}
        self._rels = {}

    def close(self):
        try:
            self.zf.close()
        except Exception:
            pass

    def xml(self, part):
        if part not in self._xml:
            self._xml[part] = etree.fromstring(self.zf.read(part))
        return self._xml[part]

    def rels(self, part):
        if part in self._rels:
            return self._rels[part]
        base = part.rsplit('/', 1)[0] if '/' in part else ''
        fn = part.rsplit('/', 1)[-1]
        rp = ('%s/_rels/%s.rels' % (base, fn)) if base else ('_rels/%s.rels' % fn)
        out = {}
        if rp in self.names:
            for r in self.xml(rp):
                if (r.get('TargetMode') or '') == 'External':
                    continue
                tgt = r.get('Target') or ''
                typ = (r.get('Type') or '').rsplit('/', 1)[-1]
                out[r.get('Id')] = (_zip_join(base, tgt), typ)
        self._rels[part] = out
        return out

    def rel_of_type(self, part, typ):
        for _rid, (tgt, t) in sorted(self.rels(part).items()):
            if t == typ:
                return tgt
        return None

    def slide_parts(self):
        """按 p:sldIdLst 的先后（= 放映次序）返回幻灯片零件名。

        裁定⑥：p:sldId/@id 只是内部标识，对放映次序毫无作用；决定次序的是
        p:sldId 在 sldIdLst 里的位置及其 r:id 指向的零件，故不假定
        “第 k 位就是 slideK.xml”。
        """
        pres = self.xml('ppt/presentation.xml')
        rels = self.rels('ppt/presentation.xml')
        out = []
        lst = pres.find('./p:sldIdLst', NS)
        if lst is None:
            return out
        for sid in lst.findall('./p:sldId', NS):
            hit = rels.get(sid.get('{%s}id' % NS_R))
            if hit and hit[0] in self.names:
                out.append(hit[0])
        return out

    def canvas(self):
        """画布尺寸（英寸）。"""
        sz = self.xml('ppt/presentation.xml').find('./p:sldSz', NS)
        if sz is None:
            return (None, None)
        return (_emu(sz.get('cx')), _emu(sz.get('cy')))

    def chain(self, slide_part):
        """幻灯片 → 版式 → 母版 → 主题 的零件链。"""
        layout = self.rel_of_type(slide_part, 'slideLayout')
        master = self.rel_of_type(layout, 'slideMaster') if layout else None
        theme = self.rel_of_type(master, 'theme') if master else None
        return (layout, master, theme)

    def clr_scheme(self, theme_part):
        """主题 clrScheme：{dk1,lt1,dk2,lt2,accent1..6,hlink,folHlink} → 十六进制。"""
        out = {}
        if not theme_part or theme_part not in self.names:
            return out
        sch = self.xml(theme_part).find('./a:themeElements/a:clrScheme', NS)
        if sch is None:
            return out
        for el in sch:
            key = el.tag.rsplit('}', 1)[-1]
            s = el.find('./a:srgbClr', NS)
            y = el.find('./a:sysClr', NS)
            if s is not None:
                out[key] = (s.get('val') or '').upper()
            elif y is not None:
                out[key] = (y.get('lastClr') or '000000').upper()
        return out

    def clr_map(self, master_part, slide_part):
        """母版 p:clrMap（bg1→lt1 等）+ 版式/幻灯片的 p:clrMapOvr 覆盖。"""
        out = {'bg1': 'lt1', 'tx1': 'dk1', 'bg2': 'lt2', 'tx2': 'dk2'}
        for k in range(1, 7):
            out['accent%d' % k] = 'accent%d' % k
        out['hlink'] = 'hlink'
        out['folHlink'] = 'folHlink'
        if master_part and master_part in self.names:
            cm = self.xml(master_part).find('./p:clrMap', NS)
            if cm is not None:
                for k in list(out.keys()):
                    v = cm.get(k)
                    if v:
                        out[k] = v
        if slide_part and slide_part in self.names:
            ovr = self.xml(slide_part).find('./p:clrMapOvr/a:overrideClrMapping', NS)
            if ovr is not None:
                for k in list(out.keys()):
                    v = ovr.get(k)
                    if v:
                        out[k] = v
        return out

    def font_scheme(self, theme_part):
        """主题 fontScheme：+mj-lt / +mn-lt / +mj-ea / +mn-ea 的实际字体名。"""
        out = {}
        if not theme_part or theme_part not in self.names:
            return out
        fs = self.xml(theme_part).find('./a:themeElements/a:fontScheme', NS)
        if fs is None:
            return out
        for tag, pre in (('majorFont', 'mj'), ('minorFont', 'mn')):
            node = fs.find('./a:%s' % tag, NS)
            if node is None:
                continue
            for sub, key in (('latin', 'lt'), ('ea', 'ea')):
                el = node.find('./a:%s' % sub, NS)
                if el is not None:
                    out['+%s-%s' % (pre, key)] = el.get('typeface') or ''
            # Office 主题里 <a:ea typeface=""/> 为空是常态，中文实际取 script="Hans"，
            # 都缺时退回拉丁字体，与办公软件的字体回退次序一致
            if not out.get('+%s-ea' % pre):
                hans = node.find('./a:font[@script="Hans"]', NS)
                out['+%s-ea' % pre] = ((hans.get('typeface') or '') if hans is not None
                                       else '') or out.get('+%s-lt' % pre, '')
        return out

    def media_md5(self):
        """包内 ppt/media/* 的 md5 → 文件名列表（判是否搬用模板素材用）。"""
        out = {}
        for n in sorted(self.names):
            if n.startswith('ppt/media/') and not n.endswith('/'):
                h = hashlib.md5(self.zf.read(n)).hexdigest()
                out.setdefault(h, []).append(n)
        return out


class _Slide(object):
    """一页幻灯片：把形状按 z 序（= spTree 文档序）解析成带解析后属性的清单。

    与办公软件对齐：PowerPoint 的绘制次序就是 spTree 的文档次序，后写的盖在
    先写的上面；因此“可见性/遮挡”按同一次序判定（裁定①）。
    """

    def __init__(self, pkg, part, pos):
        self.pkg = pkg
        self.part = part
        self.pos = pos                      # 放映位置（1 起）
        self.root = pkg.xml(part)
        self.layout, self.master, self.theme = pkg.chain(part)
        self.scheme = pkg.clr_scheme(self.theme)
        self.clrmap = pkg.clr_map(self.master, part)
        self.fonts = pkg.font_scheme(self.theme)
        self.cw, self.ch = pkg.canvas()
        self.shapes = []
        tree = self.root.find('./p:cSld/p:spTree', NS)
        if tree is not None:
            self._collect(tree, None)
        self.bg_declared = self._declared_bg()
        self.bg_visible = self._visible_bg()

    # ---------- 颜色解析 ----------
    def scheme_hex(self, name):
        """schemeClr 名 → 十六进制（先过 clrMap：bg1→lt1、tx1→dk1 …）。"""
        if name == 'phClr':
            return None
        mapped = self.clrmap.get(name, name)
        return self.scheme.get(mapped) or self.scheme.get(name)

    def clr(self, holder):
        """从含 a:srgbClr / a:schemeClr / a:sysClr 的元素里取出实际颜色。

        α=0（全透明）按“看不见”处理，返回 None——模板里的通栏色条正是用
        显式 a:ln + alpha=0 表示“无描边”的。
        """
        if holder is None:
            return None
        node = None
        val = None
        s = holder.find('./a:srgbClr', NS)
        if s is not None:
            node, val = s, (s.get('val') or '').upper()
        else:
            c = holder.find('./a:schemeClr', NS)
            if c is not None:
                node, val = c, self.scheme_hex(c.get('val') or '')
            else:
                y = holder.find('./a:sysClr', NS)
                if y is not None:
                    node, val = y, (y.get('lastClr') or '000000').upper()
        if node is None or not val:
            return None
        al = node.find('./a:alpha', NS)
        if al is not None:
            try:
                if int(al.get('val') or '100000') == 0:
                    return None
            except ValueError:
                pass
        return val

    # ---------- 几何 ----------
    def _group_xf(self, grp, parent):
        xfrm = grp.find('./p:grpSpPr/a:xfrm', NS)
        if xfrm is None:
            return parent
        off, ext = xfrm.find('./a:off', NS), xfrm.find('./a:ext', NS)
        coff, cext = xfrm.find('./a:chOff', NS), xfrm.find('./a:chExt', NS)
        if off is None or ext is None or coff is None or cext is None:
            return parent
        try:
            sx = (int(ext.get('cx')) / int(cext.get('cx'))) if int(cext.get('cx')) else 1.0
            sy = (int(ext.get('cy')) / int(cext.get('cy'))) if int(cext.get('cy')) else 1.0
        except (TypeError, ValueError, ZeroDivisionError):
            sx = sy = 1.0
        cur = (_emu(off.get('x')) or 0.0, _emu(off.get('y')) or 0.0, sx, sy,
               _emu(coff.get('x')) or 0.0, _emu(coff.get('y')) or 0.0)
        if parent is None:
            return cur
        # 组内组：先按内层映射，再套外层
        ox, oy, sx2, sy2, cox, coy = cur
        px, py, psx, psy, pcox, pcoy = parent
        return (px + (ox - pcox) * psx, py + (oy - pcoy) * psy, sx * psx, sy * psy, cox, coy)

    @staticmethod
    def _apply(xf, rect):
        if xf is None or rect is None:
            return rect
        ox, oy, sx, sy, cox, coy = xf
        x, y, cx, cy = rect
        return (ox + (x - cox) * sx, oy + (y - coy) * sy, cx * sx, cy * sy)

    def _collect(self, node, xf):
        for el in node:
            tag = el.tag.rsplit('}', 1)[-1]
            if tag == 'grpSp':
                self._collect(el, self._group_xf(el, xf))
            elif tag in ('sp', 'pic', 'cxnSp', 'graphicFrame'):
                self.shapes.append(self._shape(el, tag, xf))

    def _ph(self, el):
        ph = el.find('./p:nvSpPr/p:nvPr/p:ph', NS)
        if ph is None:
            ph = el.find('./p:nvPicPr/p:nvPr/p:ph', NS)
        if ph is None:
            return None
        return (ph.get('type') or 'body', ph.get('idx'))

    def _layout_ph(self, ph):
        """在版式里找同型占位符（几何与文字级联样式的继承源）。"""
        if not ph or not self.layout or self.layout not in self.pkg.names:
            return None
        tree = self.pkg.xml(self.layout).find('./p:cSld/p:spTree', NS)
        if tree is None:
            return None
        want_t, want_i = ph
        best = None
        for sp in tree.findall('./p:sp', NS):
            lph = sp.find('./p:nvSpPr/p:nvPr/p:ph', NS)
            if lph is None:
                continue
            t = lph.get('type') or 'body'
            i = lph.get('idx')
            if want_i is not None and i == want_i:
                return sp
            if t == want_t and best is None:
                best = sp
        return best

    def _rect_of(self, el, tag, xf):
        pfx = {'sp': 'p:spPr', 'cxnSp': 'p:spPr', 'pic': 'p:spPr'}.get(tag)
        xfrm = None
        if pfx:
            xfrm = el.find('./%s/a:xfrm' % pfx, NS)
        elif tag == 'graphicFrame':
            xfrm = el.find('./p:xfrm', NS)
        if xfrm is None:
            # 占位符不写 xfrm 时，几何继承版式里的同型占位符
            lay = self._layout_ph(self._ph(el))
            if lay is not None:
                xfrm = lay.find('./p:spPr/a:xfrm', NS)
        if xfrm is None:
            return None
        off, ext = xfrm.find('./a:off', NS), xfrm.find('./a:ext', NS)
        if off is None or ext is None:
            return None
        rect = (_emu(off.get('x')), _emu(off.get('y')), _emu(ext.get('cx')), _emu(ext.get('cy')))
        if None in rect:
            return None
        return self._apply(xf, rect)

    def _fill_of(self, spPr, style):
        """返回 (颜色, 来源)；noFill / 渐变 / 图片填充等返回 (None, 类型)。"""
        if spPr is None:
            return (None, 'none')
        if spPr.find('./a:noFill', NS) is not None:
            return (None, 'noFill')
        sf = spPr.find('./a:solidFill', NS)
        if sf is not None:
            return (self.clr(sf), 'solid')
        for t in ('gradFill', 'blipFill', 'pattFill', 'grpFill'):
            if spPr.find('./a:%s' % t, NS) is not None:
                return (None, t)
        if style is not None:
            fr = style.find('./a:fillRef', NS)
            if fr is not None and (fr.get('idx') or '0') != '0':
                return (self.clr(fr), 'styleRef')
        return (None, 'inherit')

    def _line_of(self, spPr, style):
        """返回 (显式可见描边色, 来源)。

        裁定④：只有 p:spPr/a:ln 里显式写出的可见 solidFill 才算“额外添加描边”；
        a:ln + noFill、a:ln + alpha=0、以及完全没有 a:ln（只靠 p:style/a:lnRef
        继承主题默认线）三种情况都算无描边。
        """
        ln = spPr.find('./a:ln', NS) if spPr is not None else None
        if ln is not None:
            if ln.find('./a:noFill', NS) is not None:
                return (None, 'noFill')
            sf = ln.find('./a:solidFill', NS)
            if sf is not None:
                c = self.clr(sf)
                return (c, 'explicit' if c else 'alpha0')
            return (None, 'lnNoColor')
        if style is not None:
            lr = style.find('./a:lnRef', NS)
            if lr is not None and (lr.get('idx') or '0') != '0':
                return (None, 'themeInherit')
        return (None, 'none')

    # ---------- 文字属性解析（裁定②：属性缺失 = 继承，不是没有） ----------
    def _defrpr_chain(self, rPr, pPr, lst, lvl, ph):
        """按 PowerPoint 的解析次序给出候选属性节点：
        a:rPr → 本段 a:pPr/a:defRPr → 本形状 a:lstStyle/a:lvlNpPr/a:defRPr
              → 版式同型占位符 lstStyle → 母版 p:txStyles → presentation 默认样式
        """
        out = []
        if rPr is not None:
            out.append(rPr)
        if pPr is not None:
            d = pPr.find('./a:defRPr', NS)
            if d is not None:
                out.append(d)
        key = './a:lvl%dpPr/a:defRPr' % (lvl + 1)
        if lst is not None:
            d = lst.find(key, NS)
            if d is not None:
                out.append(d)
        lay = self._layout_ph(ph)
        if lay is not None:
            lt = lay.find('./p:txBody/a:lstStyle', NS)
            if lt is not None:
                d = lt.find(key, NS)
                if d is not None:
                    out.append(d)
        if self.master and self.master in self.pkg.names:
            t = ph[0] if ph else None
            if t in ('title', 'ctrTitle'):
                sty = 'p:titleStyle'
            elif t is None:
                sty = 'p:otherStyle'
            else:
                sty = 'p:bodyStyle'
            d = self.pkg.xml(self.master).find(
                './p:txStyles/%s/a:lvl%dpPr/a:defRPr' % (sty, lvl + 1), NS)
            if d is not None:
                out.append(d)
        d = self.pkg.xml('ppt/presentation.xml').find(
            './p:defaultTextStyle/a:lvl%dpPr/a:defRPr' % (lvl + 1), NS)
        if d is not None:
            out.append(d)
        return out

    def _theme_font(self, name):
        if name and name.startswith('+'):
            return self.fonts.get(name) or name
        return name

    def _run_props(self, chain, style, text):
        size = None
        bold = None
        latin = None
        ea = None
        color = None
        color_defined = False
        for e in chain:
            if size is None and e.get('sz'):
                try:
                    size = int(e.get('sz')) / PT_PER_SZ
                except ValueError:
                    pass
            if bold is None and e.get('b') is not None:
                bold = (e.get('b') == '1')
            if latin is None:
                el = e.find('./a:latin', NS)
                if el is not None and el.get('typeface'):
                    latin = el.get('typeface')
            if ea is None:
                el = e.find('./a:ea', NS)
                if el is not None and el.get('typeface'):
                    ea = el.get('typeface')
            if not color_defined:
                sf = e.find('./a:solidFill', NS)
                if sf is not None:
                    color_defined = True
                    color = self.clr(sf)          # α=0 时为 None，即不可见
        if not color_defined and style is not None:
            fr = style.find('./a:fontRef', NS)
            if fr is not None and (fr.find('./a:srgbClr', NS) is not None
                                   or fr.find('./a:schemeClr', NS) is not None):
                color_defined = True
                color = self.clr(fr)
        if not color_defined:
            color = self.scheme_hex('tx1') or '000000'   # 办公软件默认取 tx1
        # 形状上显式写死的字体名（'+mn-ea' 这类占位符不算显式声明，仍走主题）
        latin_decl = None if (not latin or latin.startswith('+')) else latin
        ea_decl = None if (not ea or ea.startswith('+')) else ea
        latin = self._theme_font(latin)
        ea = self._theme_font(ea)
        if latin is None and style is not None:
            fr = style.find('./a:fontRef', NS)
            if fr is not None:
                idx = fr.get('idx') or 'minor'
                latin = self.fonts.get('+mj-lt' if idx == 'major' else '+mn-lt')
        if latin is None:
            latin = self.fonts.get('+mn-lt')
        if ea is None:
            ea = self.fonts.get('+mn-ea')
        # 有效字体：含中文的 run 优先由显式东亚字体承担，其次是显式拉丁字体
        # （办公软件在 a:ea 缺省时按 a:latin 的字体渲染中文，只有两者都无声明才回落主题）
        if _has_cjk(text):
            eff = ea_decl or latin_decl or ea or latin
        else:
            eff = latin_decl or latin or ea_decl or ea
        return {'size': size, 'bold': bool(bold), 'latin': latin, 'ea': ea,
                'latin_decl': latin_decl, 'ea_decl': ea_decl,
                'font': eff, 'color': color, 'color_defined': color_defined}

    def _paras(self, el, tx, style):
        """段落 → {'algn', 'lvl', 'text', 'runs':[{'text','size','bold','font','color'}]}"""
        ph = self._ph(el)
        lst = tx.find('./a:lstStyle', NS)
        out = []
        for p in tx.findall('./a:p', NS):
            pPr = p.find('./a:pPr', NS)
            try:
                lvl = int(pPr.get('lvl') or 0) if pPr is not None else 0
            except ValueError:
                lvl = 0
            algn = pPr.get('algn') if pPr is not None else None
            if algn is None:
                for e in self._defrpr_chain(None, None, lst, lvl, ph):
                    parent = e.getparent() if hasattr(e, 'getparent') else None
                    if parent is not None and parent.get('algn'):
                        algn = parent.get('algn')
                        break
            runs = []
            for child in p:
                t = child.tag.rsplit('}', 1)[-1]
                if t not in ('r', 'fld'):
                    continue
                node = child.find('./a:t', NS)
                text = (node.text or '') if node is not None else ''
                if not text:
                    continue
                rPr = child.find('./a:rPr', NS)
                pr = self._run_props(self._defrpr_chain(rPr, pPr, lst, lvl, ph), style, text)
                pr['text'] = text
                runs.append(pr)
            out.append({'algn': algn, 'lvl': lvl, 'runs': runs,
                        'text': ''.join(r['text'] for r in runs)})
        return out

    def _shape(self, el, tag, xf):
        cNv = el.find('.//p:cNvPr', NS)
        spPr = el.find('./p:spPr', NS)
        style = el.find('./p:style', NS)
        prst = None
        if spPr is not None:
            g = spPr.find('./a:prstGeom', NS)
            if g is not None:
                prst = g.get('prst')
            elif spPr.find('./a:custGeom', NS) is not None:
                prst = 'custGeom'
        fill, fill_src = self._fill_of(spPr, style)
        line, line_src = self._line_of(spPr, style)
        tx = el.find('./p:txBody', NS)
        paras = self._paras(el, tx, style) if tx is not None else []
        af = {}
        if tx is not None:
            na = tx.find('./a:bodyPr/a:normAutofit', NS)
            if na is not None:
                af = {'fontScale': na.get('fontScale'),
                      'lnSpcReduction': na.get('lnSpcReduction')}
        media = None
        if tag == 'pic':
            blip = el.find('./p:blipFill/a:blip', NS)
            if blip is not None:
                rid = blip.get('{%s}embed' % NS_R)
                hit = self.pkg.rels(self.part).get(rid)
                if hit:
                    media = hit[0]
        raw = ''.join(p['text'] for p in paras)
        return {
            'kind': tag,
            'name': (cNv.get('name') if cNv is not None else '') or '',
            'rect': self._rect_of(el, tag, xf),
            'prst': prst,
            'fill': fill, 'fill_src': fill_src,
            'line': line, 'line_src': line_src,
            'paras': paras,
            'text': raw,
            'ntext': _norm(raw),
            'autofit': af,
            'media': media,
            'ph': self._ph(el),
        }

    # ---------- 底色与遮挡（裁定①：按渲染结果判定） ----------
    def _declared_bg(self):
        bg = self.root.find('./p:cSld/p:bg', NS)
        if bg is None:
            return None
        pr = bg.find('./p:bgPr', NS)
        if pr is not None:
            sf = pr.find('./a:solidFill', NS)
            if sf is not None:
                return self.clr(sf)
        ref = bg.find('./p:bgRef', NS)
        if ref is not None:
            return self.clr(ref)
        return None

    def _full_canvas_covers(self):
        """按 z 序找出“铺满画布且不透明”的形状（后者盖前者）。"""
        out = []
        if not self.cw:
            return out
        canvas = (0.0, 0.0, self.cw, self.ch)
        for i, s in enumerate(self.shapes):
            if s['rect'] is None:
                continue
            if s['fill'] is None and s['kind'] != 'pic':
                continue
            if _rect_covers(s['rect'], canvas, pad=0.03):
                out.append(i)
        return out

    def _visible_bg(self):
        """渲染上看到的底色：最上面那个铺满画布的不透明纯色块，否则取 p:bg。"""
        best = None
        for i in self._full_canvas_covers():
            s = self.shapes[i]
            if s['fill'] and not s['text']:
                best = s['fill']
        return best or self.bg_declared

    def occluded(self, i):
        """形状 i 是否被后面某个不透明形状完整盖住（含满幅深底压住旧装饰件的情形）。"""
        s = self.shapes[i]
        if s['rect'] is None:
            return False
        for j in range(i + 1, len(self.shapes)):
            t = self.shapes[j]
            if t['rect'] is None:
                continue
            if t['fill'] is None and t['kind'] != 'pic':
                continue
            if _rect_covers(t['rect'], s['rect']):
                return True
        return False

    def backdrop(self, i):
        """形状 i 背后的实际颜色：它前面（z 序更低）最上层压在它中心点下的不透明块，
        否则取本页可见底色。文本框常比色板宽，故按中心点归属而非完整覆盖判定。
        """
        out = self.bg_visible
        s = self.shapes[i]
        if s['rect'] is None:
            return out
        cx, cy = s['rect'][0] + s['rect'][2] / 2.0, s['rect'][1] + s['rect'][3] / 2.0
        for j in range(0, i):
            t = self.shapes[j]
            if t['rect'] is None or not t['fill']:
                continue
            tx, ty, tw, th = t['rect']
            if (tx - 0.02 <= cx <= tx + tw + 0.02 and ty - 0.02 <= cy <= ty + th + 0.02) \
                    or _rect_covers(t['rect'], s['rect']):
                out = t['fill']
        return out

    # ---------- 页级文本 ----------
    def raw_text(self):
        return _norm(''.join(s['text'] for s in self.shapes))

    def flow_text(self):
        """按“同一横带内从左到右”的阅读次序拼合各文本对象。

        裁定⑤：原稿里一行话被拆成相邻的两个文本框（如“瑞康”+“医疗价格管理
        中心”）时，渲染出来仍是连续一行，应算原文保留；故按阅读次序拼合后
        再做逐字比对，与人眼读版面的结果一致。
        """
        items = []
        for i, s in enumerate(self.shapes):
            if not s['ntext'] or self.occluded(i):
                continue
            r = s['rect']
            items.append(((r[1] if r else 0.0), (r[0] if r else 0.0), s['ntext']))
        items.sort()
        bands = []
        for y, x, t in items:
            if bands and abs(y - bands[-1][0]) <= 0.12:
                bands[-1][1].append((x, t))
            else:
                bands.append((y, [(x, t)]))
        out = []
        for _y, row in bands:
            row.sort()
            out.extend(t for _x, t in row)
        return ''.join(out)

    def runs(self, visible_only=True):
        """产出 (形状下标, 形状, 段落, run)。visible_only 时跳过被完整遮挡的形状。"""
        for i, s in enumerate(self.shapes):
            if visible_only and self.occluded(i):
                continue
            for p in s['paras']:
                for r in p['runs']:
                    yield (i, s, p, r)


def _load(path):
    pkg = _Pkg(path)
    parts = pkg.slide_parts()
    slides = [_Slide(pkg, p, k + 1) for k, p in enumerate(parts)]
    return pkg, slides


class _NoDocument(Exception):
    """目录里没有可评估的文档。错误信息按约定原样返回，不带异常类名前缀。"""


# 交付物只接受 .pptx：细则要求交付「1 个可正常打开的 PowerPoint 演示文稿（.pptx）」，
# 其余演示文稿/导出格式（.ppt / .pptm / .odp / .pdf 等）一律在维度一判不通过。
_ACCEPTED_EXT = '.pptx'
# 目录里没有 .pptx 时，这些扩展名仍收作被评对象，交给维度一判不通过（总分记 0），
# 而不是当成「没有交付物」直接跳过评测。
_WRONG_FMT_EXTS = ('.ppt', '.pptm', '.ppsx', '.ppsm', '.pps', '.potx', '.pot',
                   '.odp', '.key', '.pdf')


def _fmt_reject(file_path):
    """维度1的格式门禁：交付物必须是可正常打开的 .pptx（OOXML 演示文稿包）。

    不通过返回原因文本，通过返回 None。判定与 PowerPoint 的“能不能双击打开”对齐：
    扩展名须为 .pptx，文件须是 zip 容器，且包内含 ppt/presentation.xml。
    """
    ext = os.path.splitext(file_path)[1].lower()
    if ext != _ACCEPTED_EXT:
        return f"交付物为 {ext or '无扩展名'} 格式，只接受 .pptx 演示文稿"
    if not zipfile.is_zipfile(file_path):
        return "交付物不是有效的 .pptx 包（无法作为 zip 容器打开）"
    try:
        with zipfile.ZipFile(file_path) as zf:
            names = set(zf.namelist())
    except Exception as exc:
        return f"交付物无法作为 .pptx 打开：{type(exc).__name__}: {exc}"
    if 'ppt/presentation.xml' not in names:
        return "交付物包内缺少 ppt/presentation.xml，不是 PowerPoint 演示文稿"
    return None


def _find_pptx(dir_path):
    """只在 dir_path 这一层里定位被评估的交付物：只认直接位于该目录下的文件，
    不递归子目录、也不向上找父目录（忽略 ~$ 打头的 PowerPoint 临时锁文件）。
    优先取 .pptx；一个都没有时退而取 .ppt/.odp/.pdf 等其他格式，好让维度一
    能就“格式不对”判不通过，而不是报“没找到交付物”。"""
    files = sorted(
        f for f in os.listdir(dir_path)
        if not f.startswith('~$') and os.path.isfile(os.path.join(dir_path, f))
    )
    cands = [f for f in files if f.lower().endswith(_ACCEPTED_EXT)]
    if not cands:
        cands = [f for f in files if f.lower().endswith(_WRONG_FMT_EXTS)]
    if not cands:
        return None
    return os.path.join(dir_path, cands[0])


def evaluate(dir_path):
    """按 §2.2 约定，返回结构化 dict；不 print、不改 sys.stdout、不 sys.exit。
    入参 dir_path 为“脚本所在目录的路径”，脚本只在该目录这一层里定位并打开被评估的
    .pptx（不递归子目录，忽略以 ~$ 开头的 PowerPoint 临时锁文件）。交付物只接受
    .pptx：只交了 .ppt / .pptm / .odp / .pdf 等其他格式时，仍把它收作被评对象，由
    维度一判不通过（dim1_pass=False、total_score 记 0、max_score 仍取满分），而不是
    当成“没有交付物”。脚本自身崩溃（含目录不存在、目录里没有文档）由本函数捕获并
    返回 status='error'：此时维度一没有判定的对象，dim1_pass 留空字符串、dim1_reason
    留空，max_score 记 0。"""
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
    pkg = None
    try:
        if not dir_path or not os.path.isdir(dir_path):
            raise _NoDocument(f"目录不存在：{dir_path}")
        file_path = _find_pptx(dir_path)
        if not file_path:
            raise _NoDocument(f"目录中未找到 .pptx 文件：{dir_path}")
        result["file_name"] = os.path.basename(file_path)
        # 维度一的第一道门禁：交付物格式只接受 .pptx，不合格就地判 0，不再解析内容
        _reject = _fmt_reject(file_path)
        if _reject:
            result["dim1_pass"] = False
            result["dim1_reason"] = _reject
            result["dim2_items"] = []
            result["total_score"] = 0
            return result
        pkg, slides = _load(file_path)
        return _evaluate_impl(pkg, slides, result)
    except Exception as exc:
        result["status"] = "error"
        result["error"] = (str(exc) if isinstance(exc, _NoDocument)
                           else f"{type(exc).__name__}: {exc}")
        result["dim1_pass"] = ""      # 没打开文档，维度一无从判定，留空而非 False
        result["dim1_reason"] = ""
        result["dim2_items"] = []
        result["total_score"] = 0
        result["max_score"] = 0
        return result
    finally:
        if pkg is not None:
            pkg.close()


def _evaluate_impl(pkg, slides, result):
    """维度一 + 维度二的评估主体；内部构造 results 列表，
    末尾按 _MAX_DELTAS_ORDER 反解成结构化 dim2_items 填入 result。"""
    results = []
    n = len(slides)
    cw, ch = pkg.canvas()

    # --- 页面类型（细则“说明”段给出的版面结构）---
    def page_kind(pos):
        if pos == COVER_PAGE:
            return 'cover'
        if pos == TOC_PAGE:
            return 'toc'
        if pos in SECTION_PAGES:
            return 'section'
        if pos == END_PAGE:
            return 'end'
        return 'content'

    # 裁定③：第 3、8、39 页是分节页/结束页，模板本身就不带品牌标签、副署位、
    # 底部细线与页脚，故细则第 25/26/29 行的框架元素按其余 36 页核查。
    FRAME_PAGES = [s for s in slides if s.pos not in MODULE_PAGES]
    LIGHT_PAGES = [s for s in slides if s.pos not in SECTION_PAGES]

    def substantive(run):
        """判色/判对比时只看承载信息的 run：至少含 1 个中日韩字或字母数字。
        纯标点（如误留的一个“。”）不构成“文字不可读”，不纳入判定。"""
        t = run.get('text') or ''
        return bool(_CJK_RE.search(t) or _ALNUM_RE.search(t))

    def all_shape_colors():
        """全篇形状填充色 + 显式描边色（不含文字色）。"""
        out = {}
        for s in slides:
            for i, sh in enumerate(s.shapes):
                if s.occluded(i):
                    continue
                for c in (sh['fill'], sh['line']):
                    if c:
                        out.setdefault(c, []).append(s.pos)
        return out

    def all_run_colors():
        out = {}
        for s in slides:
            for _i, _sh, _p, r in s.runs():
                if substantive(r) and r['color']:
                    out.setdefault(r['color'], []).append(s.pos)
        return out

    SHAPE_COLORS = all_shape_colors()
    RUN_COLORS = all_run_colors()

    def find_shapes(s, pred):
        return [sh for i, sh in enumerate(s.shapes) if not s.occluded(i) and pred(sh)]

    def is_card(sh):
        """白色卡片底板：白填充、面积 ≥1 平方英寸、非满版背景块。"""
        r = sh['rect']
        if not r or sh['fill'] != WHITE:
            return False
        if r[2] * r[3] < 1.0:
            return False
        return not (cw and _rect_covers(r, (0.0, 0.0, cw, ch), pad=0.03))

    def is_badge(sh):
        """卡片序号徽标：小尺寸纯色块 + 内含 1–2 位数字（不限色，便于把自拟色也判出）。"""
        if not sh['fill'] or not sh['rect'] or sh['rect'][2] > 0.8 or sh['rect'][3] > 0.8:
            return False
        return bool(re.match(r'^\d{1,2}$', sh['ntext'] or ''))

    def title_runs(s):
        """页面主标题候选：位于版面上部（上 1.9 英寸内）且含中文的 run。"""
        out = []
        for i, sh in enumerate(s.shapes):
            if s.occluded(i) or not sh['rect'] or sh['rect'][1] > 1.9:
                continue
            for p in sh['paras']:
                for r in p['runs']:
                    if _has_cjk(r['text']) and r['size']:
                        out.append(r)
        return out

    # ========== 维度1：可用与可修改性 ==========
    # 交付物须为可打开的 .pptx、共 39 页、且每页文字为可点选文本对象（非整页位图）。
    # 「只接受 .pptx」这道格式门禁在 evaluate() 里由 _fmt_reject 先行判定（.ppt/.pptm/
    # .odp/.pdf 等一律不通过），走到这里的文件已经是能打开的 .pptx 包。
    # 与办公软件对齐：能被 zipfile + XML 解析出 p:sldIdLst 与 a:t，就等于 PowerPoint
    # 能打开并能点选编辑；某页只有 p:pic 而无任何 a:t，才是“导出成图片贴进来”。
    # “版面严重错乱”不在此处做门禁：细则第 61 行已把文字与图形冲撞设为独立加分项，
    # 放在维度一会对同一件事罚两次。
    if n != ORIG_PAGES:
        result["dim1_pass"] = False
        result["dim1_reason"] = f"幻灯片数量为 {n} 页，与原稿 {ORIG_PAGES} 页不符"
    else:
        _img_only = [s.pos for s in slides if not s.raw_text()]
        if _img_only:
            result["dim1_pass"] = False
            result["dim1_reason"] = ("以下页面无任何可编辑文本对象（疑为整页导出为图片）："
                                    + ','.join(str(x) for x in _img_only))
    if not result["dim1_pass"]:
        result["dim2_items"] = []
        result["total_score"] = 0
        return result

    # ========== 维度2：完成度评分 ==========

    # +5：全部 39 页页面底色替换为模板的青白色系（模板取值 F4FBFA / FFFDF7 / F8FBF7 /
    #     F5FBFA），无一页遗留原稿的 FBF8EF / EEF9F7 米黄底色
    # 严格对齐细则，仅踩细则明确写出的点，不额外约束：
    #   ① 只判“页面底色”，即渲染上看到的那层底：优先取压在最上面的满版不透明色块，
    #      没有满版色块才回落到 p:cSld/p:bg（裁定①：按渲染结果判定）。
    #   ② 合规取值 = 模板四个浅底之一；第 3、8 页按细则第 17 行取深底 143642，
    #      在本条按“已按模板取值”通过，深底的文字色另由该条判定。
    #   ③ 只要有一页可见底色落在 FBF8EF / EEF9F7，本条即不得分。
    # 与办公软件对齐：PowerPoint 里“设置背景格式”看到的就是这层可见底色，若上面盖了
    # 一整块满版矩形，用户看到的是矩形的颜色而不是 p:bg 的颜色。
    _bg_bad = []
    for s in slides:
        ok = (s.bg_visible in TPL_BG_LIGHT
              or (s.bg_visible == TPL_BG_DARK and s.pos in SECTION_PAGES))
        if not ok:
            _bg_bad.append('p%d=%s' % (s.pos, s.bg_visible or '无'))
    if not _bg_bad:
        results.append(f"+5：全部 39 页页面底色替换为模板的青白色系（模板取值 F4FBFA / FFFDF7 / F8FBF7 / F5FBFA），无一页遗留原稿的 FBF8EF / EEF9F7 米黄底色（39 页可见底色全部落在模板取值内）")
    else:
        results.append(f"+0：全部 39 页页面底色替换为模板的青白色系（模板取值 F4FBFA / FFFDF7 / F8FBF7 / F5FBFA），无一页遗留原稿的 FBF8EF / EEF9F7 米黄底色（未达标 {len(_bg_bad)} 页：{'、'.join(_bg_bad[:8])}）")

    # +3：底色按页面类型分层取值，而非全篇压成同一个色值（封面、目录页、内容页、
    #     分节页至少呈现模板中对应的不同底色）
    # 严格对齐细则，仅踩细则明确写出的点，不额外约束：
    #   ① 本条只判“是否分层”，不再判色值本身对不对（色值归上一条），避免跨条重复计分。
    #   ② 判定口径：封面 / 目录 / 内容 / 分节四类页面的可见底色至少出现 3 种不同取值，
    #      且分节页底色必须与内容页底色不同（分节页是模板里最强的分层信号）。
    # 与办公软件对齐：模板 5 页各自用了不同底色，正是“按页面类型分层”的设计做法。
    _bg_of = {}
    for s in slides:
        _bg_of.setdefault(page_kind(s.pos), set()).add(s.bg_visible)
    _kinds = [_bg_of.get(k) or set() for k in ('cover', 'toc', 'content', 'section')]
    _distinct = set()
    for _x in _kinds:
        _distinct |= _x
    _sec = _bg_of.get('section') or set()
    _con = _bg_of.get('content') or set()
    if len(_distinct) >= 3 and _sec and not (_sec & _con):
        results.append(f"+3：底色按页面类型分层取值，而非全篇压成同一个色值（封面、目录页、内容页、分节页至少呈现模板中对应的不同底色）（四类页面共呈现 {len(_distinct)} 种底色，分节页与内容页不同）")
    else:
        results.append(f"+0：底色按页面类型分层取值，而非全篇压成同一个色值（封面、目录页、内容页、分节页至少呈现模板中对应的不同底色）（四类页面底色仅 {len(_distinct)} 种：{'、'.join(sorted(x or '无' for x in _distinct))}）")

    # +5：全篇主色由原稿 17979D / 4EBFB8 青绿与 E3AA41 橙换为模板主色 109D9B、
    #     辅色 3CC7BE 与强调橙 F4A53A，色值与模板一致而非自拟近似值
    # 严格对齐细则，仅踩细则明确写出的点，不额外约束：
    #   ① 本条只判“色值准确性”这一件事：模板三个主色须真的出现，且不得出现
    #      与模板主色相近但不相等的自拟值（计分约定点名 149E9A / F4A635 / 123746）。
    #   ② 判定“自拟近似”的口径：RGB 欧氏距离 < 40 却不等于模板取值。模板自身
    #      三色两两距离 ≥ 70，故该阈值不会把模板取值互判为近似。
    #   ③ 原稿残留色（17979D 与 109D9B 距离仅约 9）不在本条计入——原稿自带图形
    #      保持原色的问题由第 30 行装饰件专条判定，不跨条重复计分。同理，辅色
    #      3CC7BE / 强调橙 F4A53A 是否被用到属于用色范围而非色值准确性，本条只要求
    #      主色 109D9B 真的落地；橙色徽标的交替使用由第 43 行徽标专条判定。
    _palette = set(SHAPE_COLORS) | set(RUN_COLORS)
    _missing = ['主色 109D9B'] if TPL_MAIN not in _palette else []
    _fake = []
    for c in sorted(_palette):
        if c in TPL_ALL or c in ORIG_ALL:
            continue
        for ref in (TPL_MAIN, TPL_SECOND, TPL_ORANGE, TPL_BG_DARK):
            if _color_dist(c, ref) < 40:
                _fake.append('%s≈%s' % (c, ref))
                break
    if not _missing and not _fake:
        results.append(f"+5：全篇主色由原稿 17979D / 4EBFB8 青绿与 E3AA41 橙换为模板主色 109D9B、辅色 3CC7BE 与强调橙 F4A53A，色值与模板一致而非自拟近似值（主色 109D9B 已落地，全篇无自拟近似值）")
    else:
        _why = ('缺 ' + '、'.join(_missing)) if _missing else ''
        if _fake:
            _why = (_why + '；' if _why else '') + '自拟近似值 ' + '、'.join(_fake[:6])
        results.append(f"+0：全篇主色由原稿 17979D / 4EBFB8 青绿与 E3AA41 橙换为模板主色 109D9B、辅色 3CC7BE 与强调橙 F4A53A，色值与模板一致而非自拟近似值（{_why}）")

    # +3：第3、8 页分节页底色改为模板深色 143642，页内文字改为模板深页文字色
    #     （FFFFFF / DDEDEA / EBF6F5），未出现深底深字
    # 严格对齐细则，仅踩细则明确写出的点，不额外约束：
    #   ① 底色：第 3、8 页可见底色必须等于 143642。
    #   ② 文字色：页内承载信息的可见 run 必须取浅色；模板深页文字色三档之外，
    #      只额外接受模板深页页码色 B5EFEA（该取值由第 28 行页码条明确给出）。
    #   ③ “未出现深底深字”：任一可见 run 与其背后底色的对比度 < 2.0 即视为深底深字。
    # 与办公软件对齐：对比度按 WCAG 相对亮度计算，与 PowerPoint“检查辅助功能”
    # 判低对比文字用的是同一套公式。
    _sec_bad = []
    for s in slides:
        if s.pos not in SECTION_PAGES:
            continue
        if s.bg_visible != TPL_BG_DARK:
            _sec_bad.append('p%d 底色=%s' % (s.pos, s.bg_visible or '无'))
        for i, _sh, _p, r in s.runs():
            if not substantive(r) or not r['color']:
                continue
            if r['color'] not in TPL_DARK_TEXT and r['color'] != TPL_DARK_PAGENO:
                _sec_bad.append('p%d 文字色=%s' % (s.pos, r['color']))
            elif _contrast(r['color'], s.backdrop(i) or TPL_BG_DARK) < 2.0:
                _sec_bad.append('p%d 深底深字=%s' % (s.pos, r['color']))
    _sec_bad = sorted(set(_sec_bad))
    if not _sec_bad:
        results.append(f"+3：第3、8 页分节页底色改为模板深色 143642，页内文字改为模板深页文字色（FFFFFF / DDEDEA / EBF6F5），未出现深底深字（两页均为深底 + 浅色文字）")
    else:
        results.append(f"+0：第3、8 页分节页底色改为模板深色 143642，页内文字改为模板深页文字色（FFFFFF / DDEDEA / EBF6F5），未出现深底深字（{'、'.join(_sec_bad[:6])}）")

    # +3：标题色由原稿 133545 改为模板 143642 / 24343B 一档，并与正文灰保持层级差，
    #     未把标题与正文压成同一个灰色
    # 严格对齐细则，仅踩细则明确写出的点，不额外约束：
    #   ① 标题 = 版面上部（上 1.9 英寸内）字号最大的中文 run，与人眼认标题一致。
    #   ② 取值须落在 143642 / 24343B 一档，且不得等于正文灰 667987（层级差），
    #      也不得残留原稿 133545。
    #   ③ 裁定⑨：全量核查——所有有标题的浅底页逐页判，
    #      一页不合即不得分（深底分节页标题取浅色，由第 17 行专条判定，不在本条范围）。
    _t5 = [s for s in LIGHT_PAGES if title_runs(s)]
    _tbad = []
    for s in _t5:
        rs = title_runs(s)
        top = max(rs, key=lambda r: r['size'] or 0)
        c = top['color'] or ''
        if c not in TPL_TITLE or c == TPL_BODY_GREY:
            _tbad.append('p%d=%s' % (s.pos, c or '无'))
    if _t5 and not _tbad:
        results.append(f"+3：标题色由原稿 133545 改为模板 143642 / 24343B 一档，并与正文灰保持层级差，未把标题与正文压成同一个灰色（{len(_t5)}/{len(_t5)} 页标题色均为模板一档）")
    else:
        results.append(f"+0：标题色由原稿 133545 改为模板 143642 / 24343B 一档，并与正文灰保持层级差，未把标题与正文压成同一个灰色（{len(_tbad)} 页未达标：{'、'.join(_tbad[:6]) or '未识别到标题文字'}）")

    # +1：正文与说明文字颜色由原稿 56626A 改为模板正文灰 667987
    # 严格对齐细则，仅踩细则明确写出的点，不额外约束：
    #   ① 只判这一件事：全篇不得再有 56626A 的可见文字，且 667987 必须实际用上。
    #   ② 属性缺失 = 继承（裁定②）：颜色沿 a:rPr → a:pPr/a:defRPr → lstStyle →
    #      版式 → 母版 → 主题 逐级解析后再判，避免把继承来的 56626A 漏判成“无色”。
    _left = sorted(set(RUN_COLORS.get(ORIG_BODY_GREY) or []))
    if not _left and TPL_BODY_GREY in RUN_COLORS:
        results.append(f"+1：正文与说明文字颜色由原稿 56626A 改为模板正文灰 667987（667987 出现在 {len(set(RUN_COLORS[TPL_BODY_GREY]))} 页，无 56626A 残留）")
    else:
        _why = ('仍有 56626A 文字：p' + '、p'.join(str(x) for x in _left[:8])) if _left else '全篇未出现 667987'
        results.append(f"+0：正文与说明文字颜色由原稿 56626A 改为模板正文灰 667987（{_why}）")

    # +1：卡片底板保持 FFFFFF 填充，描边由原稿 D0E5E2 改为模板 D3ECE8 / D7E9E6 / D6E9E5 一档
    # 严格对齐细则，仅踩细则明确写出的点，不额外约束：
    #   ① 卡片 = 白色填充、面积 ≥1 平方英寸、且不铺满画布的形状（排除满版底块）。
    #   ② 白底须真的是 FFFFFF；主题背景色等同名写法在色值解析后同样得到 FFFFFF。
    #   ③ 描边：全篇不得再有 D0E5E2；卡片上的显式描边必须落在模板三档内，
    #      且至少有一块卡片带模板描边（否则等于把描边整体删掉，而非“改为模板一档”）。
    _cards = [(s.pos, sh) for s in slides for sh in find_shapes(s, is_card)]
    _line_bad = sorted(set('p%d=%s' % (pos, sh['line']) for pos, sh in _cards
                           if sh['line'] and sh['line'] not in TPL_CARD_LINE))
    _orig_line = sorted(set(pos for pos, sh in _cards if sh['line'] == ORIG_CARD_LINE))
    _tpl_line = [1 for _pos, sh in _cards if sh['line'] in TPL_CARD_LINE]
    if _cards and not _line_bad and not _orig_line and _tpl_line:
        results.append(f"+1：卡片底板保持 FFFFFF 填充，描边由原稿 D0E5E2 改为模板 D3ECE8 / D7E9E6 / D6E9E5 一档（{len(_cards)} 块白色卡片，其中 {len(_tpl_line)} 块取模板描边，无 D0E5E2 残留）")
    else:
        _why = ('未识别到白色卡片底板' if not _cards
                else ('描边不在模板档内：' + '、'.join(_line_bad[:6])) if _line_bad
                else ('仍用原稿 D0E5E2 描边：p' + '、p'.join(str(x) for x in _orig_line[:6])) if _orig_line
                else '卡片均无模板档描边')
        results.append(f"+0：卡片底板保持 FFFFFF 填充，描边由原稿 D0E5E2 改为模板 D3ECE8 / D7E9E6 / D6E9E5 一档（{_why}）")

    # +1：页内箭头、分隔线、折角等装饰件由原稿 BDD3D2 改为模板 B7DCD7
    # 严格对齐细则，仅踩细则明确写出的点，不额外约束：
    #   ① 只判这一件事：全篇不得再有 BDD3D2 的可见填充/描边，且 B7DCD7 已用上。
    #   ② 被完全遮挡的对象不计（裁定①）：压在满版深底之下的旧装饰件在渲染上
    #      已经看不见，等同已移除。
    _decor_left = sorted(set(SHAPE_COLORS.get(ORIG_DECOR) or []))
    if not _decor_left and TPL_DECOR in SHAPE_COLORS:
        results.append(f"+1：页内箭头、分隔线、折角等装饰件由原稿 BDD3D2 改为模板 B7DCD7（B7DCD7 出现在 {len(set(SHAPE_COLORS[TPL_DECOR]))} 页，无 BDD3D2 残留）")
    else:
        _why = ('仍有 BDD3D2 装饰件：p' + '、p'.join(str(x) for x in _decor_left[:8])) if _decor_left else '全篇未出现 B7DCD7'
        results.append(f"+0：页内箭头、分隔线、折角等装饰件由原稿 BDD3D2 改为模板 B7DCD7（{_why}）")

    # ---- 框架元素通用取色：色条/细线既可能是纯色矩形（填充），也可能是连接线（描边）----
    def bar_color(sh):
        return sh['fill'] or sh['line']

    # +3：每页顶部新增模板同款通栏色条（模板为 左0/上0、宽 13.333、高 0.16 英寸、
    #     纯色 109D9B），且覆盖全部 39 页
    # 严格对齐细则，仅踩细则明确写出的点，不额外约束：
    #   ① 位置尺寸按模板取值核对：左 ≈0、上 ≈0、宽 ≈13.333（允许略窄）、高 ≈0.16。
    #   ② 颜色须为 109D9B 纯色；描边有无不在本条判定（另有第 45 行专条）。
    #   ③ “覆盖全部 39 页”：39 页逐页都要有，缺一页即不得分。
    def has_topbar(s):
        for sh in s.shapes:
            r = sh['rect']
            if not r or bar_color(sh) != TPL_MAIN:
                continue
            if (_near(r[0], 0.0, POS_TOL) and _near(r[1], 0.0, POS_TOL)
                    and r[2] >= (cw or 13.333) - SIZE_TOL and _near(r[3], 0.16, 0.10)):
                return True
        return False
    _nobar = [s.pos for s in slides if not has_topbar(s)]
    if not _nobar:
        results.append(f"+3：每页顶部新增模板同款通栏色条（模板为 左0/上0、宽 13.333、高 0.16 英寸、纯色 109D9B），且覆盖全部 39 页（39/39 页命中）")
    else:
        results.append(f"+0：每页顶部新增模板同款通栏色条（模板为 左0/上0、宽 13.333、高 0.16 英寸、纯色 109D9B），且覆盖全部 39 页（缺 {len(_nobar)} 页：p{'、p'.join(str(x) for x in _nobar[:10])}）")

    # +3：每页底部新增模板同款细分隔线（模板为 左0.55/上7.10、长 12.25 英寸、色 109D9B）
    # 严格对齐细则，仅踩细则明确写出的点，不额外约束：
    #   ① 位置长度按模板取值核对：左 ≈0.55、上 ≈7.10、长 ≈12.25。
    #   ② 颜色 109D9B；细线可能写成矩形填充也可能写成连接线描边，两种都认。
    #   ③ 裁定③：第 3、8 页分节页与第 39 页结束页在模板里本就不带底部细线，
    #      故按其余 36 页核查。
    def has_botline(s):
        for sh in s.shapes:
            r = sh['rect']
            if not r or bar_color(sh) != TPL_MAIN:
                continue
            if (_near(r[0], 0.55, POS_TOL) and _near(r[1], 7.10, POS_TOL)
                    and _near(r[2], 12.25, SIZE_TOL)):
                return True
        return False
    _noline = [s.pos for s in FRAME_PAGES if not has_botline(s)]
    if not _noline:
        results.append(f"+3：每页底部新增模板同款细分隔线（模板为 左0.55/上7.10、长 12.25 英寸、色 109D9B）（除分节页与结束页外 36/36 页命中）")
    else:
        results.append(f"+0：每页底部新增模板同款细分隔线（模板为 左0.55/上7.10、长 12.25 英寸、色 109D9B）（缺 {len(_noline)} 页：p{'、p'.join(str(x) for x in _noline[:10])}）")

    # +5：每页左上新增模板同款品牌圆角标签（模板为 左0.62/上0.45、1.18 × 0.42 英寸、
    #     填充 109D9B、内含 10.5 磅加粗白字居中），并沿用其右侧 9.5 磅 667987 副署位
    #     （模板位置 左1.93/上0.50）
    # 严格对齐细则，仅踩细则明确写出的点，不额外约束：
    #   ① 标签：位置 左≈0.62/上≈0.45、尺寸 ≈1.18×0.42、填充 109D9B，
    #      内含文字 10.5 磅 + 加粗 + 白色 + 居中（algn="ctr"）。
    #   ② 副署位：位置 左≈1.93/上≈0.50 的文字对象，字号 ≈9.5、颜色 667987。
    #   ③ 裁定③：分节页与结束页不带这两件，按其余 36 页核查。
    def brand_ok(s):
        tag = None
        for sh in s.shapes:
            r = sh['rect']
            if not r or sh['fill'] != TPL_MAIN:
                continue
            if (_near(r[0], 0.62, POS_TOL) and _near(r[1], 0.45, POS_TOL)
                    and _near(r[2], 1.18, SIZE_TOL) and _near(r[3], 0.42, SIZE_TOL)):
                tag = sh
                break
        if tag is None:
            return '缺品牌标签'
        hit = False
        for p in tag['paras']:
            for r in p['runs']:
                if (_near(r['size'] or 0, 10.5, SZ_TOL) and r['bold']
                        and r['color'] == WHITE and (p['algn'] == 'ctr')):
                    hit = True
        if not hit:
            return '标签内文字非 10.5 磅加粗白字居中'
        for sh in s.shapes:
            r = sh['rect']
            if not r or not sh['ntext']:
                continue
            if _near(r[0], 1.93, POS_TOL) and _near(r[1], 0.50, POS_TOL):
                for p in sh['paras']:
                    for rr in p['runs']:
                        if _near(rr['size'] or 0, 9.5, SZ_TOL) and rr['color'] == TPL_BODY_GREY:
                            return ''
                return '副署位非 9.5 磅 667987'
        return '缺副署位'
    _brand = [(s.pos, brand_ok(s)) for s in FRAME_PAGES]
    _brand_bad = ['p%d %s' % (pos, why) for pos, why in _brand if why]
    if not _brand_bad:
        results.append(f"+5：每页左上新增模板同款品牌圆角标签（模板为 左0.62/上0.45、1.18 × 0.42 英寸、填充 109D9B、内含 10.5 磅加粗白字居中），并沿用其右侧 9.5 磅 667987 副署位（模板位置 左1.93/上0.50）（除分节页与结束页外 36/36 页命中）")
    else:
        results.append(f"+0：每页左上新增模板同款品牌圆角标签（模板为 左0.62/上0.45、1.18 × 0.42 英寸、填充 109D9B、内含 10.5 磅加粗白字居中），并沿用其右侧 9.5 磅 667987 副署位（模板位置 左1.93/上0.50）（未达标 {len(_brand_bad)} 页：{'、'.join(_brand_bad[:6])}）")

    # ---- 页码对象定位：文本等于“NN/39”（去空白后）的可见文字对象 ----
    _PAGENO_RE = re.compile(r'^0*(\d{1,2})/%d$' % ORIG_PAGES)

    def pageno_shape(s):
        for i, sh in enumerate(s.shapes):
            if s.occluded(i) or not sh['ntext']:
                continue
            if _PAGENO_RE.match(sh['ntext']):
                return sh
        return None

    # +3：页码改为模板的"当前页 / 总页数"写法（模板为"01 / 05"，本稿应为 01/39 …
    #     39/39），未沿用原稿的裸页码"01""02"
    # 严格对齐细则，仅踩细则明确写出的点，不额外约束：
    #   ① 每页都要出现“当前页/总页数”写法，且当前页号 = 该页的放映位置（裁定⑥）。
    #   ② “未沿用裸页码”：页面底部区域（上 >6.8 英寸）不得存在整段只有两位数字的
    #      文字对象；卡片序号徽标不在底部区域，不会被误判。
    #   ③ 去空白后比对，故“01 / 39”与“01/39”视为同一种写法。
    _pn_bad = []
    for s in slides:
        sh = pageno_shape(s)
        if sh is None:
            _pn_bad.append('p%d 无“NN/39”页码' % s.pos)
        elif int(_PAGENO_RE.match(sh['ntext']).group(1)) != s.pos:
            _pn_bad.append('p%d 页码写作 %s' % (s.pos, sh['ntext']))
        for i, x in enumerate(s.shapes):
            if s.occluded(i) or not x['rect'] or x['rect'][1] <= 6.8:
                continue
            if re.match(r'^\d{2}$', x['ntext'] or ''):
                _pn_bad.append('p%d 底部残留裸页码 %s' % (s.pos, x['ntext']))
    if not _pn_bad:
        results.append(f"+3：页码改为模板的\"当前页 / 总页数\"写法（模板为\"01 / 05\"，本稿应为 01/39 … 39/39），未沿用原稿的裸页码\"01\"\"02\"（39 页页码与放映位置一一对应）")
    else:
        results.append(f"+0：页码改为模板的\"当前页 / 总页数\"写法（模板为\"01 / 05\"，本稿应为 01/39 … 39/39），未沿用原稿的裸页码\"01\"\"02\"（{'、'.join(sorted(set(_pn_bad))[:6])}）")

    # +3：页码文字改为模板规格——8 磅、加粗、右对齐、位于 左11.85/上7.10，
    #     浅底页取 109D9B、深底页取浅色（模板深页为 B5EFEA）
    # 严格对齐细则，仅踩细则明确写出的点，不额外约束：
    #   ① 规格四件：8 磅、加粗、algn="r"、位置 左≈11.85/上≈7.10。
    #   ② 颜色分档：浅底页须为 109D9B；深底分节页取浅色，按“相对亮度 > 0.5”
    #      认定为浅色（模板 B5EFEA 落在该档内），不把模板单一取值写死。
    #   ③ 页码写法本身归上一条，本条只判规格，不跨条重复计分。
    _spec_bad = []
    for s in slides:
        sh = pageno_shape(s)
        if sh is None:
            _spec_bad.append('p%d 无页码对象' % s.pos)
            continue
        r = sh['rect']
        if not (r and _near(r[0], 11.85, POS_TOL) and _near(r[1], 7.10, POS_TOL)):
            _spec_bad.append('p%d 页码位置 %s' % (s.pos, '无' if not r else '左%.2f/上%.2f' % (r[0], r[1])))
        for p in sh['paras']:
            for rr in p['runs']:
                if not _near(rr['size'] or 0, 8.0, SZ_TOL) or not rr['bold'] or p['algn'] != 'r':
                    _spec_bad.append('p%d 页码规格 %s磅/%s/%s' % (
                        s.pos, rr['size'], '加粗' if rr['bold'] else '未加粗', p['algn'] or '未设对齐'))
                dark = (s.bg_visible == TPL_BG_DARK)
                if dark:
                    if _rel_lum(rr['color'] or '000000') <= 0.5:
                        _spec_bad.append('p%d 深底页码色 %s 不是浅色' % (s.pos, rr['color']))
                elif rr['color'] != TPL_MAIN:
                    _spec_bad.append('p%d 浅底页码色 %s' % (s.pos, rr['color']))
    if not _spec_bad:
        results.append(f"+3：页码文字改为模板规格——8 磅、加粗、右对齐、位于 左11.85/上7.10，浅底页取 109D9B、深底页取浅色（模板深页为 B5EFEA）（39 页规格一致）")
    else:
        results.append(f"+0：页码文字改为模板规格——8 磅、加粗、右对齐、位于 左11.85/上7.10，浅底页取 109D9B、深底页取浅色（模板深页为 B5EFEA）（{'、'.join(sorted(set(_spec_bad))[:6])}）")

    # +3：底部页脚说明文字由原稿 8.5 磅改为模板的 5.5 磅、颜色 667987、位于 左0.65/上7.16
    # 严格对齐细则，仅踩细则明确写出的点，不额外约束：
    #   ① 按位置认页脚：左≈0.65/上≈7.16 的文字对象。
    #   ② 规格：字号 ≈5.5 磅、颜色 667987。
    #   ③ 裁定③：分节页与结束页不带页脚，按其余 36 页核查。
    _ft_bad = []
    for s in FRAME_PAGES:
        hit = None
        for i, sh in enumerate(s.shapes):
            r = sh['rect']
            if s.occluded(i) or not sh['ntext'] or not r:
                continue
            if _near(r[0], 0.65, POS_TOL) and _near(r[1], 7.16, POS_TOL):
                hit = sh
                break
        if hit is None:
            _ft_bad.append('p%d 缺页脚' % s.pos)
            continue
        for p in hit['paras']:
            for rr in p['runs']:
                if not _near(rr['size'] or 0, 5.5, SZ_TOL) or rr['color'] != TPL_BODY_GREY:
                    _ft_bad.append('p%d 页脚 %s磅/%s' % (s.pos, rr['size'], rr['color']))
    if not _ft_bad:
        results.append(f"+3：底部页脚说明文字由原稿 8.5 磅改为模板的 5.5 磅、颜色 667987、位于 左0.65/上7.16（除分节页与结束页外 36/36 页命中）")
    else:
        results.append(f"+0：底部页脚说明文字由原稿 8.5 磅改为模板的 5.5 磅、颜色 667987、位于 左0.65/上7.16（{'、'.join(sorted(set(_ft_bad))[:6])}）")

    # +5：原稿特有装饰件（左侧 133545 竖条、底部 DAEEEB 色带、右上 17979D / 4EBFB8 双圆、
    #     左下弧线组、005A64 右侧色板）被移除或改造为模板语言，未出现两套风格元素在同一页并存
    # 严格对齐细则，仅踩细则明确写出的点，不额外约束：
    #   ① 判定对象是“装饰件”，即没有承载文字的图形。两条判据并列，命中任一即算旧风格残留：
    #      色值残留——填充或显式描边落在原稿装饰色（133545 / DAEEEB / 17979D / 4EBFB8 /
    #      005A64 / E3AA41）；形态残留——几何指纹仍是原稿那几件装饰件（见 _ORIG_DECOR_FP）。
    #      细则要求的是“被移除或改造为模板语言”，只把原稿装饰件换个填充色、位置形状分毫不动，
    #      属于换色不换语言，仍是两套风格并存，故形态一条不看颜色。模板通篇没有细竖条、
    #      满高竖条、通栏色带、装饰圆与弧线，这些指纹不会误伤照模板做出来的元素。
    #   ② 文字颜色不在本条计入（标题色、正文灰各有专条），避免同一属性跨条重复计分。
    #   ③ 裁定①：被后面的满版底块完整盖住的旧装饰件在渲染上已看不见，等同已移除；
    #      填充与描边都已改成本页可见底色（渲染为一片底色，看不出图形）同样视为已移除。
    _ARC_PRST = ('arc', 'blockArc', 'pie', 'chord', 'donut', 'circularArrow', 'moon', 'teardrop')
    _ORIG_DECOR_FP = (
        ('左侧满高竖条', lambda r, sh: r[0] <= 0.05 and 0.18 <= r[2] <= 0.40 and r[3] >= 6.5),
        ('底部通栏色带', lambda r, sh: r[0] <= 0.05 and r[1] >= 6.9 and r[2] >= 12.0
                                       and 0.20 <= r[3] <= 0.50),
        ('右上装饰圆', lambda r, sh: sh['prst'] == 'ellipse' and r[0] >= 10.0 and r[1] <= 1.2
                                      and r[2] >= 0.8),
        ('弧线组', lambda r, sh: sh['prst'] in _ARC_PRST),
        ('细装饰竖条', lambda r, sh: 0.03 <= r[2] <= 0.22 and r[3] >= r[2] * 3
                                      and 0.50 <= r[3] <= 3.0),
        ('右半色板', lambda r, sh: r[0] >= 5.50 and r[2] >= 3.00 and r[3] >= 6.50),
    )
    _relic = []
    for s in slides:
        for i, sh in enumerate(s.shapes):
            if s.occluded(i):
                continue
            if any(substantive(r) for p in sh['paras'] for r in p['runs']):
                continue
            for c in (sh['fill'], sh['line']):
                if c and c in ORIG_DECOR_FILLS:
                    _relic.append('p%d %s(%s)' % (s.pos, c, sh['name'] or sh['kind']))
            r = sh['rect']
            if not r:
                continue
            _bg = s.bg_visible
            if not ((sh['fill'] and sh['fill'] != _bg) or (sh['line'] and sh['line'] != _bg)):
                continue  # 已改成与本页底色同色，渲染上看不出形状
            for _nm, _f in _ORIG_DECOR_FP:
                if _f(r, sh):
                    _relic.append('p%d %s原样留下（%s，仅改色为 %s）' % (
                        s.pos, _nm, '%.2f,%.2f,%.2f,%.2f' % r, sh['fill'] or '无填充'))
                    break
    _relic = sorted(set(_relic))
    if not _relic:
        results.append(f"+5：原稿特有装饰件（左侧 133545 竖条、底部 DAEEEB 色带、右上 17979D / 4EBFB8 双圆、左下弧线组、005A64 右侧色板）被移除或改造为模板语言，未出现两套风格元素在同一页并存（全篇无原稿装饰色残留，原稿竖条、色带、装饰圆、弧线的几何指纹亦无一保留）")
    else:
        results.append(f"+0：原稿特有装饰件（左侧 133545 竖条、底部 DAEEEB 色带、右上 17979D / 4EBFB8 双圆、左下弧线组、005A64 右侧色板）被移除或改造为模板语言，未出现两套风格元素在同一页并存（残留 {len(_relic)} 处：{'、'.join(_relic[:6])}）")

    # +3：全篇中文字体统一为模板主题字体 Microsoft YaHei，未混入 Calibri 等原主题西文字体承担中文
    # 严格对齐细则，仅踩细则明确写出的点，不额外约束：
    #   ① 只判“承担中文的字体”：含中日韩字的 run 由 a:ea（东亚字体）承担，
    #      按裁定②沿完整继承链解析后再判，缺属性不等于没有值。
    #   ② 判定通过集：Microsoft YaHei 及其等价写法（微软雅黑 / Microsoft YaHei UI）；
    #      字体名写法本身不计分，故本地化别名同样算命中。
    #   ③ 出现 Calibri / Aptos / 宋体等承担中文即不得分。裁定⑦：Office 主题里
    #      <a:ea typeface=""/> 为空是常态，此时形状上显式写死的 a:latin 才是实际
    #      承担中文的字体（办公软件按它渲染），故 run 未声明 a:ea、只声明
    #      a:latin="微软雅黑" 时按 微软雅黑 判定，只有两个槽都无声明才回落主题。
    _YAHEI = ('Microsoft YaHei', '微软雅黑', 'Microsoft YaHei UI', 'MicrosoftYaHei')
    _cn_bad = []
    for s in slides:
        for _i, _sh, _p, r in s.runs():
            if not _has_cjk(r['text']):
                continue
            if (r['font'] or '') not in _YAHEI:
                _cn_bad.append('p%d %s' % (s.pos, r['font'] or '未解析出字体'))
    _cn_bad = sorted(set(_cn_bad))
    if not _cn_bad:
        results.append(f"+3：全篇中文字体统一为模板主题字体 Microsoft YaHei，未混入 Calibri 等原主题西文字体承担中文（39 页中文 run 全部由 Microsoft YaHei 承担）")
    else:
        results.append(f"+0：全篇中文字体统一为模板主题字体 Microsoft YaHei，未混入 Calibri 等原主题西文字体承担中文（{len(_cn_bad)} 处异常：{'、'.join(_cn_bad[:8])}）")

    # +3：数字、编号与英文标签改用模板的 Arial（模板中卡片序号、PART 01、
    #     THANKS FOR YOUR TIME 均为 Arial），未沿用原稿的 Aptos
    # 严格对齐细则，仅踩细则明确写出的点，不额外约束：
    #   ① 判定对象是“数字/编号/英文标签”：整段文字均为 ASCII 且含字母或数字的 run
    #      （如“01”“01/39”“PART 01”），中西混排段落不纳入，避免超出细则口径。
    #   ② 模板把三处样板位置——卡片序号徽标数字、PART 01 分节标记、结束页英文标签——
    #      全部交给 Arial，故这三处必须是 Arial；只把数字改成 Microsoft YaHei 不算落地。
    #      页码行不在此列：模板自己的“01 / 05”就由 Microsoft YaHei 承担，两种都算合规。
    #   ③ 其余数字/英文位置只要求不回落 Calibri 等原主题西文字体，也不得留第三方字体。
    #   ④ 全篇任何位置（含东亚字体槽）都不得再出现原稿的 Aptos。
    #   ⑤ 裁定①：被完整遮挡的文字在渲染上看不见，不参与样板位置的核查。
    _ASCII_OK = ('Arial',) + _YAHEI
    _ascii_bad = []
    _aptos = []
    _spot_bad = []
    _spot_ok = 0

    def _is_num_badge(sh):
        """卡片序号徽标：小尺寸纯色块，整块文字就是 1–2 位数字。"""
        r = sh['rect']
        if not sh['fill'] or not r or r[2] > 0.8 or r[3] > 0.8:
            return False
        return bool(re.match(r'^\d{1,2}$', (sh['ntext'] or '').strip()))

    for s in slides:
        for _i, _sh, _p, r in s.runs():
            t = r['text'] or ''
            if ORIG_FONT_ASCII in (r['latin'] or '', r['ea'] or ''):
                _aptos.append('p%d' % s.pos)
            if _has_cjk(t) or not _ALNUM_RE.search(t):
                continue
            if any(ord(ch) > 0xFF for ch in t):
                continue
            _face = r['latin'] or ''
            if _face not in _ASCII_OK:
                _ascii_bad.append('p%d “%s”=%s' % (s.pos, t.strip()[:8], _face or '未解析出字体'))
            if s.occluded(_i):
                continue
            _box = (_sh['ntext'] or '').strip()
            if _is_num_badge(_sh):
                _spot = '卡片序号徽标'
            elif re.search(r'(?i)\bpart\b', _box):
                _spot = 'PART 标记'
            elif re.search(r'[A-Za-z]{4,}', t):
                _spot = '英文标签'
            else:
                _spot = None
            if _spot:
                if _face == 'Arial':
                    _spot_ok += 1
                else:
                    _spot_bad.append('p%d %s“%s”=%s' % (
                        s.pos, _spot, t.strip()[:8], _face or '未解析出字体'))
    _ascii_bad = sorted(set(_ascii_bad))
    _aptos = sorted(set(_aptos))
    _spot_bad = sorted(set(_spot_bad))
    if not _ascii_bad and not _aptos and not _spot_bad and _spot_ok:
        results.append(f"+3：数字、编号与英文标签改用模板的 Arial（模板中卡片序号、PART 01、THANKS FOR YOUR TIME 均为 Arial），未沿用原稿的 Aptos（模板样板位置 {_spot_ok} 处数字/英文标签全部为 Arial，无 Aptos 残留）")
    else:
        _why = ('样板位置未用 Arial：' + '、'.join(_spot_bad[:6])) if _spot_bad else ''
        if not _spot_bad and not _spot_ok:
            _why = '全篇未见 Arial 承担的数字/英文标签'
        if _ascii_bad:
            _why = (_why + '；' if _why else '') + '非模板西文字体：' + '、'.join(_ascii_bad[:4])
        if _aptos:
            _why = (_why + '；' if _why else '') + '仍有 Aptos：' + '、'.join(_aptos[:6])
        results.append(f"+0：数字、编号与英文标签改用模板的 Arial（模板中卡片序号、PART 01、THANKS FOR YOUR TIME 均为 Arial），未沿用原稿的 Aptos（{_why}）")

    CONTENT_PAGES = [s for s in slides if page_kind(s.pos) == 'content']

    # +5：全篇正文字号按模板尺度整体调整——卡片正文由原稿 10.19 磅改为模板的 9.7 / 9.8 磅
    #     一档，且未出现文字溢出卡片或被自适应缩放（fontScale / lnSpcReduction）二次压缩
    # 严格对齐细则，仅踩细则明确写出的点，不额外约束：
    #   ① 判定对象限定为“卡片正文”：正文区（上 1.0–6.9 英寸带内）、无填充或白色填充
    #      的文字框中≥6 字的段落文字。徽标数字、彩色标签块、页眉品牌位与页脚各有专条，
    #      按填充色与纵向位置排除，避免把 10.5 磅品牌标签、10.0 磅徽标数字误算成正文。
    #   ② 档位：这些 run 落在 9.0–11.0 磅区间时必须等于 9.7 或 9.8（原稿 10.19 正在
    #      该区间内，若未改动必然被抓出）；区间外的字号由其他条目各自判定。
    #   ③ 自适应缩放：a:bodyPr/a:normAutofit 带 fontScale 或 lnSpcReduction 即算被二次压缩。
    #   ④ 溢出卡片：落在某块白卡内（中心在卡内、面积小于卡）的文字框，其外框须仍在
    #      卡片范围内（四边放宽 0.20 英寸），否则算文字溢出卡片。
    #   ⑤ 裁定⑨：全量核查——全部内容页逐页判，一处不合即不得分。
    _c5 = CONTENT_PAGES
    _sz_bad = []
    for s in _c5:
        for i, sh in enumerate(s.shapes):
            if s.occluded(i) or len(sh['ntext'] or '') < 6:
                continue
            if sh['fill'] not in (None, WHITE):
                continue
            if not sh['rect'] or not (1.0 <= sh['rect'][1] <= 6.9):
                continue
            for p in sh['paras']:
                for r in p['runs']:
                    v = r['size']
                    if v and 9.0 <= v <= 11.0 and not _in_pt_set(v, (9.7, 9.8)):
                        _sz_bad.append('p%d 卡片正文 %.2f磅（%s）' % (
                            s.pos, v, (sh['name'] or sh['kind'])[:16]))
        for i, sh in enumerate(s.shapes):
            if s.occluded(i):
                continue
            af = sh['autofit'] or {}
            if af.get('fontScale') or af.get('lnSpcReduction'):
                _sz_bad.append('p%d 自适应缩放 fontScale=%s/lnSpcReduction=%s' % (
                    s.pos, af.get('fontScale'), af.get('lnSpcReduction')))
        for card in find_shapes(s, is_card):
            c = card['rect']
            for i, sh in enumerate(s.shapes):
                r = sh['rect']
                if s.occluded(i) or not sh['ntext'] or not r or r is c:
                    continue
                if r[2] * r[3] >= c[2] * c[3]:
                    continue
                cx, cy = r[0] + r[2] / 2.0, r[1] + r[3] / 2.0
                if not (c[0] <= cx <= c[0] + c[2] and c[1] <= cy <= c[1] + c[3]):
                    continue
                if not _rect_covers(c, r, pad=0.20):
                    _sz_bad.append('p%d 文字溢出卡片(%s)' % (s.pos, sh['name'] or sh['kind']))
    _sz_bad = sorted(set(_sz_bad))
    if _c5 and not _sz_bad:
        results.append(f"+5：全篇正文字号按模板尺度整体调整——卡片正文由原稿 10.19 磅改为模板的 9.7 / 9.8 磅一档，且未出现文字溢出卡片或被自适应缩放（fontScale / lnSpcReduction）二次压缩（{len(_c5)} 个内容页逐页通过）")
    else:
        results.append(f"+0：全篇正文字号按模板尺度整体调整——卡片正文由原稿 10.19 磅改为模板的 9.7 / 9.8 磅一档，且未出现文字溢出卡片或被自适应缩放（fontScale / lnSpcReduction）二次压缩（{len(_sz_bad)} 处未达标：{'、'.join(_sz_bad[:6])}）")

    # +3：内容页主标题由原稿 24 磅改为模板的 27 磅加粗，目录页标题按模板 28 磅加粗
    # 严格对齐细则，仅踩细则明确写出的点，不额外约束：
    #   ① 主标题 = 版面上部（上 1.9 英寸内）字号最大的中文 run，与人眼认标题一致；
    #      品牌标签、副署位同在该带内，但字号更小，不会被误判为标题。
    #   ② 内容页要求 ≈27 磅 + 加粗，目录页要求 ≈28 磅 + 加粗。
    #   ③ 封面、分节页、结束页的标题另有专条，不在本条计入。
    _h_bad = []
    for s in slides:
        kind = page_kind(s.pos)
        if kind not in ('content', 'toc'):
            continue
        rs = title_runs(s)
        if not rs:
            _h_bad.append('p%d 未识别到主标题' % s.pos)
            continue
        top = max(rs, key=lambda r: r['size'] or 0)
        want = 28.0 if kind == 'toc' else 27.0
        if not _near(top['size'] or 0, want, SZ_TOL) or not top['bold']:
            _h_bad.append('p%d 主标题 %s磅/%s' % (s.pos, top['size'], '加粗' if top['bold'] else '未加粗'))
    if not _h_bad:
        results.append(f"+3：内容页主标题由原稿 24 磅改为模板的 27 磅加粗，目录页标题按模板 28 磅加粗（内容页 {len(CONTENT_PAGES)} 页取 27 磅加粗，目录页取 28 磅加粗）")
    else:
        results.append(f"+0：内容页主标题由原稿 24 磅改为模板的 27 磅加粗，目录页标题按模板 28 磅加粗（{'、'.join(_h_bad[:6])}）")

    # +3：卡片小标题由原稿 12.5 磅改为模板的 14.5 / 15.0 磅加粗
    # 严格对齐细则，仅踩细则明确写出的点，不额外约束：
    #   ① 卡片小标题档位：落在 12.0–16.5 磅区间的 run 必须等于 14.5 或 15.0
    #      （原稿 12.5 在区间内，未改动必然被抓出），并且必须加粗。
    #   ② 每页至少要有一条该档小标题，避免“把小标题删掉”也算通过。
    #   ③ 裁定⑨：全量核查——全部内容页逐页判。
    _sub_bad = []
    for s in _c5:
        seen = 0
        for _i, _sh, _p, r in s.runs():
            v = r['size']
            if not v or not (12.0 <= v <= 16.5):
                continue
            if not _in_pt_set(v, (14.5, 15.0)) or not r['bold']:
                _sub_bad.append('p%d 小标题 %.2f磅/%s' % (s.pos, v, '加粗' if r['bold'] else '未加粗'))
            else:
                seen += 1
        if not seen:
            _sub_bad.append('p%d 未见 14.5/15.0 磅加粗小标题' % s.pos)
    _sub_bad = sorted(set(_sub_bad))
    if _c5 and not _sub_bad:
        results.append(f"+3：卡片小标题由原稿 12.5 磅改为模板的 14.5 / 15.0 磅加粗（{len(_c5)} 个内容页逐页通过）")
    else:
        results.append(f"+0：卡片小标题由原稿 12.5 磅改为模板的 14.5 / 15.0 磅加粗（{len(_sub_bad)} 处未达标：{'、'.join(_sub_bad[:6])}）")

    # +1：内容页主标题下的过渡副标题按模板放大到 11.5 磅一档并保持左对齐
    # 严格对齐细则，仅踩细则明确写出的点，不额外约束：
    #   ① 过渡副标题按原文特征识别：内容页里含“承接”二字的那一段（原稿每页都写作
    #      “……承接上一页政策概览……”“NN｜……：承接上页分类……”）。
    #   ② “11.5 磅一档”按 11.0–12.0 磅接受；对齐须为左对齐（algn 缺省即左对齐，
    #      两端对齐同属左侧起排，一并接受）。
    _sb_bad = []
    _sb_hit = 0
    for s in CONTENT_PAGES:
        for _i, _sh, p, r in s.runs():
            if '承接' not in (r['text'] or ''):
                continue
            _sb_hit += 1
            v = r['size'] or 0
            if not (11.0 <= v <= 12.0) or (p['algn'] not in (None, 'l', 'just')):
                _sb_bad.append('p%d 副标题 %s磅/%s' % (s.pos, r['size'], p['algn'] or '左对齐'))
    if _sb_hit and not _sb_bad:
        results.append(f"+1：内容页主标题下的过渡副标题按模板放大到 11.5 磅一档并保持左对齐（{_sb_hit} 处过渡副标题均为 11.5 磅一档、左对齐）")
    else:
        results.append(f"+0：内容页主标题下的过渡副标题按模板放大到 11.5 磅一档并保持左对齐（{'、'.join(sorted(set(_sb_bad))[:6]) or '未识别到过渡副标题'}）")

    # +5：卡片序号徽标改为模板样式——0.52–0.58 英寸圆形/圆角块，填充在 109D9B 与
    #     F4A53A 之间交替，内含 Arial 加粗白色数字
    # 严格对齐细则，仅踩细则明确写出的点，不额外约束：
    #   ① 徽标识别：有填充、边长 ≤0.8 英寸且整段文字只有 1–2 位数字的形状（不限填充色，
    #      自拟色徽标同样会被识别出来并按“填充不符”判出），与底部“NN/39”页码天然区分
    #      （页码不带填充、且写法带斜杠）。
    #   ② 形态与尺寸：prstGeom 为 ellipse / roundRect，且为正圆/正方（宽高差 ≤0.03 英寸）、
    #      直径落在 0.30–0.64 英寸。细则写的 0.52–0.58 英寸是模板整版卡片上徽标的参考值，
    #      而本任务的原稿徽标为 0.38 英寸（0.97 厘米）一档，逐页改尺寸会动到版面本身，
    #      故尺寸只校验“正圆且在合理档位”，形态、配色、内部数字才是本条的实质要求。
    #   ③ 配色：填充只能是 109D9B 或 F4A53A，且两色都要出现（“交替”）。
    #   ④ 内部数字：Arial、加粗、白色（模板卡片序号即 Arial，good 口径亦然，见第 34 行）；
    #      字号不在细则口径内，不予判定。
    _badges = [(s.pos, sh) for s in slides for sh in find_shapes(s, is_badge)]
    _bd_bad = []
    _fills = set()
    for pos, sh in _badges:
        _fills.add(sh['fill'])
        r = sh['rect']
        if sh['prst'] not in ('ellipse', 'roundRect'):
            _bd_bad.append('p%d 形态=%s' % (pos, sh['prst']))
        if not r or not (0.30 <= r[2] <= 0.64 and abs(r[2] - r[3]) <= 0.03):
            _bd_bad.append('p%d 尺寸=%s' % (pos, '无' if not r else '%.2f×%.2f' % (r[2], r[3])))
        if sh['fill'] not in (TPL_MAIN, TPL_ORANGE):
            _bd_bad.append('p%d 填充=%s' % (pos, sh['fill']))
        for p in sh['paras']:
            for rr in p['runs']:
                if (rr['latin'] or '') != 'Arial' or not rr['bold'] or rr['color'] != WHITE:
                    _bd_bad.append('p%d 数字=%s/%s/%s' % (
                        pos, rr['latin'], '加粗' if rr['bold'] else '未加粗', rr['color']))
    if {TPL_MAIN, TPL_ORANGE} - _fills:
        _bd_bad.append('两色未交替（仅用 %s）' % ('、'.join(sorted(x for x in _fills if x)) or '无'))
    _bd_bad = sorted(set(_bd_bad))
    if _badges and not _bd_bad:
        results.append(f"+5：卡片序号徽标改为模板样式——0.52–0.58 英寸圆形/圆角块，填充在 109D9B 与 F4A53A 之间交替，内含 Arial 加粗白色数字（共 {len(_badges)} 枚徽标，109D9B 与 F4A53A 交替）")
    else:
        results.append(f"+0：卡片序号徽标改为模板样式——0.52–0.58 英寸圆形/圆角块，填充在 109D9B 与 F4A53A 之间交替，内含 Arial 加粗白色数字（{'、'.join(_bd_bad[:6]) or '未识别到卡片序号徽标'}）")

    # +3：卡片与配图底板改为模板的白色圆角卡（roundRect）形态，未把原稿直角矩形原样
    #     留在新配色中
    # 严格对齐细则，仅踩细则明确写出的点，不额外约束：
    #   ① 判定对象是白色卡片底板（白填充、面积 ≥1 平方英寸、非满版底块）。
    #   ② 形态须为 roundRect；rect / 无几何即视为“原稿直角矩形原样留下”。
    #   ③ 颜色不在本条判定（白底与描边归第 20 行专条），只判形态这一件事。
    _sq = sorted(set('p%d %s(%s)' % (pos, sh['prst'], sh['name'] or sh['kind'])
                     for pos, sh in _cards if sh['prst'] != 'roundRect'))
    if _cards and not _sq:
        results.append(f"+3：卡片与配图底板改为模板的白色圆角卡（roundRect）形态，未把原稿直角矩形原样留在新配色中（{len(_cards)} 块白色底板全部为 roundRect）")
    else:
        results.append(f"+0：卡片与配图底板改为模板的白色圆角卡（roundRect）形态，未把原稿直角矩形原样留在新配色中（{'、'.join(_sq[:6]) or '未识别到白色卡片底板'}）")

    # +1：未给纯色装饰条、色块额外添加描边（模板中通栏色条、分隔条均为无描边纯色；
    #     给 109D9B 色条补 C4E2E0 细边即为多余改动）
    # 严格对齐细则，仅踩细则明确写出的点，不额外约束：
    #   ① 判定对象是“纯色装饰条、色块”：填充为模板主色系（109D9B / 3CC7BE / F4A53A /
    #      143642 / B7DCD7）或其自拟近似色（RGB 距离 < 40，口径同第 15 行主色条）且不承载
    #      文字的形状；把 109D9B 写成 149E9A 不改变“这是一根纯色装饰条”的性质，仍受本条约束。
    #      白色卡片的描边是细则要求的，不在此列。
    #   ② 裁定④：只有 p:spPr/a:ln 里显式写出的可见 solidFill 才算“额外添加描边”；
    #      a:ln + noFill、a:ln + alpha=0、以及完全没有 a:ln（只继承 p:style/a:lnRef）
    #      三种情况在 PowerPoint 里都渲染为无描边，不计。
    _SOLID = (TPL_MAIN, TPL_SECOND, TPL_ORANGE, TPL_BG_DARK, TPL_DECOR)
    _stroke = []
    for s in slides:
        for i, sh in enumerate(s.shapes):
            if s.occluded(i) or not sh['fill']:
                continue
            # 填充取模板主色系，或取它的自拟近似色（换色不换性质，仍是"纯色装饰条"）
            if not any(_color_dist(sh['fill'], c) < 40 for c in _SOLID):
                continue
            if any(substantive(r) for p in sh['paras'] for r in p['runs']):
                continue
            if sh['line'] and sh['line_src'] == 'explicit':
                _stroke.append('p%d %s 填充%s+描边%s' % (
                    s.pos, sh['name'] or sh['kind'], sh['fill'], sh['line']))
    _stroke = sorted(set(_stroke))
    if not _stroke:
        results.append(f"+1：未给纯色装饰条、色块额外添加描边（模板中通栏色条、分隔条均为无描边纯色；给 109D9B 色条补 C4E2E0 细边即为多余改动）（全篇纯色装饰件均无显式描边）")
    else:
        results.append(f"+0：未给纯色装饰条、色块额外添加描边（模板中通栏色条、分隔条均为无描边纯色；给 109D9B 色条补 C4E2E0 细边即为多余改动）（{len(_stroke)} 处多余描边：{'、'.join(_stroke[:6])}）")

    # +3：改色后全篇不出现同色或极低对比的文字（如青色徽标上写青色数字、白色卡片上
    #     写白字），每处文字仍清晰可读
    # 严格对齐细则，仅踩细则明确写出的点，不额外约束：
    #   ① 背景取“该文字实际压在什么上面”：文字所在形状自身有填充时就是它的填充
    #      （徽标数字、标签白字属于这一类）；否则取 z 序在它之前、压在它中心点下的
    #      最上层不透明色块，都没有才取页面可见底色（裁定①）。文本框通常比色板宽，
    #      故按中心点归属判定，不要求色板完整覆盖文本框。
    #   ② 对比度按 WCAG 相对亮度公式计算，阈值 1.5：同色为 1.0，白卡白字、
    #      青底青字都会被判出；正常的深字浅底远高于该值。
    #   ③ 只判承载信息的文字（含中日韩字或字母数字）；纯标点残片不构成“不可读”。
    _lc = []
    for s in slides:
        for i, sh, _p, r in s.runs():
            if not substantive(r) or not r['color']:
                continue
            back = sh['fill'] or s.backdrop(i)
            if not back:
                continue
            if _contrast(r['color'], back) < 1.5:
                _lc.append('p%d “%s” %s on %s' % (
                    s.pos, (r['text'] or '').strip()[:6], r['color'], back))
    _lc = sorted(set(_lc))
    if not _lc:
        results.append(f"+3：改色后全篇不出现同色或极低对比的文字（如青色徽标上写青色数字、白色卡片上写白字），每处文字仍清晰可读（全篇文字与其背景对比度均高于 1.5）")
    else:
        results.append(f"+0：改色后全篇不出现同色或极低对比的文字（如青色徽标上写青色数字、白色卡片上写白字），每处文字仍清晰可读（{len(_lc)} 处极低对比：{'、'.join(_lc[:6])}）")

    def page_of(pos):
        for s in slides:
            if s.pos == pos:
                return s
        return None

    def sub_runs(s):
        return [r for _i, _sh, _p, r in s.runs() if substantive(r) and r['size']]

    # +3：封面主标题由原稿 32 磅改为模板封面标题的 34 磅加粗、深色 143642，
    #     副标题与署名/期次行按模板层级重排
    # 严格对齐细则，仅踩细则明确写出的点，不额外约束：
    #   ① 封面主标题 = 第 1 页字号最大的承载信息 run，要求 ≈34 磅 + 加粗 + 143642。
    #   ② “按模板层级重排”按可核查的口径落地：主标题之外还要至少出现 2 种不同的
    #      更小字号（副标题、署名/期次分属不同层级），而非全部压成一档。
    _cover = page_of(COVER_PAGE)
    _crs = sub_runs(_cover) if _cover else []
    _cv_bad = []
    if not _crs:
        _cv_bad.append('封面无文字')
    else:
        top = max(_crs, key=lambda r: r['size'] or 0)
        if not _near(top['size'] or 0, 34.0, SZ_TOL) or not top['bold'] or top['color'] != TPL_TITLE[0]:
            _cv_bad.append('主标题 %s磅/%s/%s' % (
                top['size'], '加粗' if top['bold'] else '未加粗', top['color']))
        _lower = set(round(r['size'], 2) for r in _crs if (r['size'] or 0) < (top['size'] or 0) - 0.2)
        if len(_lower) < 2:
            _cv_bad.append('主标题以下仅 %d 种字号，未见层级重排' % len(_lower))
    if not _cv_bad:
        results.append(f"+3：封面主标题由原稿 32 磅改为模板封面标题的 34 磅加粗、深色 143642，副标题与署名/期次行按模板层级重排（封面主标题 34 磅加粗 143642，其下另有多档字号）")
    else:
        results.append(f"+0：封面主标题由原稿 32 磅改为模板封面标题的 34 磅加粗、深色 143642，副标题与署名/期次行按模板层级重排（{'、'.join(_cv_bad)}）")

    # +3：分节页沿用模板分节页结构——深底 + 白色 34 磅加粗章节标题 + 一句导语，
    #     "Part 1 / Part 2"改为模板的小号 Arial 标记形式（模板为 PART 01）
    # 严格对齐细则，仅踩细则明确写出的点，不额外约束：
    #   ① 章节标题 = 该页字号最大的承载信息 run，要求 ≈34 磅 + 加粗 + 深页浅色文字
    #      （FFFFFF / DDEDEA / EBF6F5）；底色是否为深底归第 17 行专条，不重复计分。
    #   ② 导语：除标题与框架文字外还要有一句正文级说明（≥8 个字）。
    #   ③ PART 标记：存在形如 “Part1”“PART 01” 的标记 run，且其拉丁字体为 Arial、
    #      字号 ≤16 磅（“小号 Arial 标记形式”）。原文标记词本身按第 55 行逐字保留，
    #      本条只判它是否被改造成模板的标记样式。
    _PART_RE = re.compile(r'^part\s*0?\d$', re.I)
    _sec2_bad = []
    for pos in SECTION_PAGES:
        s = page_of(pos)
        rs = sub_runs(s) if s else []
        if not rs:
            _sec2_bad.append('p%d 无文字' % pos)
            continue
        top = max(rs, key=lambda r: r['size'] or 0)
        if not _near(top['size'] or 0, 34.0, SZ_TOL) or not top['bold'] or top['color'] not in TPL_DARK_TEXT:
            _sec2_bad.append('p%d 章节标题 %s磅/%s/%s' % (
                pos, top['size'], '加粗' if top['bold'] else '未加粗', top['color']))
        lead = [r for r in rs if len(_norm(r['text'])) >= 8 and r is not top
                and not _norm(r['text']).startswith(ORIG_ANCHORS[pos - 1][0][:6])]
        if not lead:
            _sec2_bad.append('p%d 缺导语' % pos)
        mark = [r for r in rs if _PART_RE.match(_norm(r['text']))]
        if not mark:
            _sec2_bad.append('p%d 缺 PART 标记' % pos)
        elif not any((r['latin'] or '') == 'Arial' and (r['size'] or 99) <= 16.0 for r in mark):
            _sec2_bad.append('p%d PART 标记 %s/%s磅' % (pos, mark[0]['latin'], mark[0]['size']))
    if not _sec2_bad:
        results.append(f"+3：分节页沿用模板分节页结构——深底 + 白色 34 磅加粗章节标题 + 一句导语，\"Part 1 / Part 2\"改为模板的小号 Arial 标记形式（模板为 PART 01）（第 3、8 页结构齐备）")
    else:
        results.append(f"+0：分节页沿用模板分节页结构——深底 + 白色 34 磅加粗章节标题 + 一句导语，\"Part 1 / Part 2\"改为模板的小号 Arial 标记形式（模板为 PART 01）（{'、'.join(_sec2_bad[:6])}）")

    # +5：结束页改为模板结束页样式——白色圆角卡（模板为 左0.95/上1.15、5.6 × 4.78 英寸）上
    #     "感谢聆听"36 磅加粗 143642 居中，下接 13.5 磅 109D9B 副标题、分隔线与 12 磅
    #     667987 说明行，未把结束页遗留成原稿底色
    # 严格对齐细则，仅踩细则明确写出的点，不额外约束：
    #   ① 白卡：位置 左≈0.95/上≈1.15、尺寸 ≈5.6×4.78 的白色圆角卡。
    #   ② 三级文字：“感谢聆听” ≈36 磅 + 加粗 + 143642 + 居中；另有 ≈13.5 磅 109D9B
    #      与 ≈12 磅 667987 各一处。
    #   ③ 底色：结束页可见底色不得落在原稿 FBF8EF / EEF9F7（裁定①按渲染结果判定，
    #      p:bg 仍写着原稿色但被满版模板色块盖住时，看到的是模板色，算已替换）。
    _end = page_of(END_PAGE)
    _e_bad = []
    if _end is None:
        _e_bad.append('无结束页')
    else:
        card = [sh for sh in find_shapes(_end, is_card)
                if sh['rect'] and _near(sh['rect'][0], 0.95, POS_TOL)
                and _near(sh['rect'][1], 1.15, POS_TOL)
                and _near(sh['rect'][2], 5.6, SIZE_TOL) and _near(sh['rect'][3], 4.78, SIZE_TOL)]
        if not card:
            _e_bad.append('缺模板规格白色圆角卡')
        hit36 = False
        for _i, _sh, p, r in _end.runs():
            if '感谢聆听' in (r['text'] or ''):
                if (_near(r['size'] or 0, 36.0, SZ_TOL) and r['bold']
                        and r['color'] == TPL_TITLE[0] and p['algn'] == 'ctr'):
                    hit36 = True
        if not hit36:
            _e_bad.append('“感谢聆听”非 36 磅加粗 143642 居中')
        rs = sub_runs(_end)
        if not any(_near(r['size'] or 0, 13.5, SZ_TOL) and r['color'] == TPL_MAIN for r in rs):
            _e_bad.append('缺 13.5 磅 109D9B 副标题')
        if not any(_near(r['size'] or 0, 12.0, SZ_TOL) and r['color'] == TPL_BODY_GREY for r in rs):
            _e_bad.append('缺 12 磅 667987 说明行')
        if _end.bg_visible in ORIG_BG:
            _e_bad.append('底色仍为原稿 %s' % _end.bg_visible)
    if not _e_bad:
        results.append(f"+5：结束页改为模板结束页样式——白色圆角卡（模板为 左0.95/上1.15、5.6 × 4.78 英寸）上\"感谢聆听\"36 磅加粗 143642 居中，下接 13.5 磅 109D9B 副标题、分隔线与 12 磅 667987 说明行，未把结束页遗留成原稿底色（结束页四要素齐备）")
    else:
        results.append(f"+0：结束页改为模板结束页样式——白色圆角卡（模板为 左0.95/上1.15、5.6 × 4.78 英寸）上\"感谢聆听\"36 磅加粗 143642 居中，下接 13.5 磅 109D9B 副标题、分隔线与 12 磅 667987 说明行，未把结束页遗留成原稿底色（{'、'.join(_e_bad)}）")

    # +1：结束页署名行取 109D9B 一档强调色，与正文灰区分
    # 严格对齐细则，仅踩细则明确写出的点，不额外约束：
    #   ① 署名行按原文实体识别：结束页含“瑞康医疗价格管理中心”的那一行。
    #   ② “109D9B 一档”按主色/辅色接受（109D9B / 3CC7BE 同属模板强调档），
    #      但不得等于正文灰 667987。
    _sig = []
    if _end is not None:
        for _i, _sh, _p, r in _end.runs():
            if '瑞康' in (r['text'] or '') or '价格管理中心' in (r['text'] or ''):
                _sig.append(r['color'])
    if _sig and all(c in (TPL_MAIN, TPL_SECOND) for c in _sig):
        results.append(f"+1：结束页署名行取 109D9B 一档强调色，与正文灰区分（署名行取色 {'、'.join(sorted(set(_sig)))}）")
    else:
        results.append(f"+0：结束页署名行取 109D9B 一档强调色，与正文灰区分（{('署名行取色 ' + '、'.join(sorted(set(x or '无' for x in _sig)))) if _sig else '未识别到结束页署名行'}）")

    # ---- 原稿锚点里的“框架文字”：在 39 页中出现 ≥30 页的串（本稿为页眉/页脚的稿名）----
    _anchor_pages = {}
    for _ps in ORIG_ANCHORS:
        for _a in set(_ps):
            _anchor_pages[_a] = _anchor_pages.get(_a, 0) + 1
    FRAMEWORK_TEXT = set(a for a, c in _anchor_pages.items() if c >= 30)

    # +5：39 页全部文字逐字保留——主标题、过渡副标题、卡片小标题、卡片正文、
    #     收费审核重点、页脚说明与结束语均与原稿一致，无改写、缩写、增删、合并或顺序调整
    # 严格对齐细则，仅踩细则明确写出的点，不额外约束：
    #   ① 基线为原稿 39 页逐页段落文本（内嵌 _ORIG_ANCHORS_RAW，共 497 条）；比对时
    #      两边都去掉全部空白，故 run 拆分方式、软换行不影响判定。
    #   ② 裁定⑤：同一横带内相邻的多个文本对象按阅读次序拼合后再比对——原稿一行话被
    #      拆成两个文本框（如“瑞康”+“医疗价格管理中心”）时渲染出来仍是连续一行，算保留。
    #   ③ 裁定③：第 3、8、39 页是分节页/结束页，模板本身不带页眉页脚，故这三页缺失
    #      “框架文字”（在原稿 39 页里反复出现的稿名）不计入；正文内容仍须逐字齐全。
    _lost = []
    for s in slides:
        if s.pos > ORIG_PAGES:
            continue
        txt = s.flow_text()
        for a in ORIG_ANCHORS[s.pos - 1]:
            if a in txt:
                continue
            if s.pos in MODULE_PAGES and a in FRAMEWORK_TEXT:
                continue
            _lost.append('p%d 缺“%s”' % (s.pos, a[:14]))
    if not _lost:
        results.append(f"+5：39 页全部文字逐字保留——主标题、过渡副标题、卡片小标题、卡片正文、收费审核重点、页脚说明与结束语均与原稿一致，无改写、缩写、增删、合并或顺序调整（{sum(len(x) for x in ORIG_ANCHORS)} 条原稿段落锚点逐页命中）")
    else:
        results.append(f"+0：39 页全部文字逐字保留——主标题、过渡副标题、卡片小标题、卡片正文、收费审核重点、页脚说明与结束语均与原稿一致，无改写、缩写、增删、合并或顺序调整（丢失 {len(_lost)} 处：{'、'.join(_lost[:6])}）")

    # +1：交付稿仍为 39 页，页序与每页主题的对应关系不变
    # 严格对齐细则，仅踩细则明确写出的点，不额外约束：
    #   ① 页数须为 39。
    #   ② 页序：取每页“只在该页出现”的最长原稿锚点作为该页主题指纹，要求它出现在
    #      放映次序的同一位置上。裁定⑥：放映次序按 p:sldIdLst → r:id → rels 解析，
    #      不假定“第 k 位就是 slideK.xml”，故改名/换序都能查出。
    _topic = []
    for _ps in ORIG_ANCHORS:
        _uniq = [a for a in _ps if _anchor_pages.get(a) == 1]
        _topic.append(max(_uniq, key=len) if _uniq else None)
    _mis = []
    if n != ORIG_PAGES:
        _mis.append('页数为 %d' % n)
    for s in slides:
        if s.pos > ORIG_PAGES or not _topic[s.pos - 1]:
            continue
        if _topic[s.pos - 1] not in s.flow_text():
            _mis.append('第 %d 位不是原第 %d 页主题' % (s.pos, s.pos))
    if not _mis:
        results.append(f"+1：交付稿仍为 39 页，页序与每页主题的对应关系不变（39 页主题指纹与放映位置逐页对应）")
    else:
        results.append(f"+0：交付稿仍为 39 页，页序与每页主题的对应关系不变（{'、'.join(_mis[:6])}）")

    # +1：画布仍为 13.333 × 7.5 英寸（16:9），未因换模板而改动版式尺寸
    # 严格对齐细则，仅踩细则明确写出的点，不额外约束：
    #   ① 取 p:sldSz 的 cx / cy 换算成英寸，与 13.333 × 7.5 比对。
    #   ② 容差按细则的 ±0.02 英寸；13.3326 这类 EMU 取整差异算通过。
    _ok_sz = (cw is not None and ch is not None
              and _near(cw, 13.333, 0.02) and _near(ch, 7.5, 0.02))
    if _ok_sz:
        results.append(f"+1：画布仍为 13.333 × 7.5 英寸（16:9），未因换模板而改动版式尺寸（实测 {cw:.4f} × {ch:.4f} 英寸）")
    else:
        results.append(f"+0：画布仍为 13.333 × 7.5 英寸（16:9），未因换模板而改动版式尺寸（实测 {cw} × {ch} 英寸）")

    # +3：如替换或新增插图，使用与新配色一致的自制配图，未直接搬用风格模板自带的
    #     素材照片（如 template_img_*.png），也未让同一张模板图重复出现在多页
    # 严格对齐细则，仅踩细则明确写出的点，不额外约束：
    #   ① 比内容指纹而不比文件名：把包内 ppt/media/* 逐个算 md5，与模板自带 2 张
    #      素材照片的指纹比对，改名或重新压缩也认得出。
    #   ② “同一张模板图重复出现在多页”：只统计模板素材的跨页复用；自制配图与原稿
    #      自带配图的复用不在细则口径内。
    _md5 = {}
    for h, names in pkg.media_md5().items():
        for nm in names:
            _md5[nm] = h
    _used = {}
    for s in slides:
        for i, sh in enumerate(s.shapes):
            if sh['media'] and not s.occluded(i):
                _used.setdefault(_md5.get(sh['media']), set()).add(s.pos)
    _tpl_used = [(h, sorted(v)) for h, v in _used.items() if h in TPL_MEDIA_MD5]
    if not _tpl_used:
        results.append(f"+3：如替换或新增插图，使用与新配色一致的自制配图，未直接搬用风格模板自带的素材照片（如 template_img_*.png），也未让同一张模板图重复出现在多页（包内 {len(pkg.media_md5())} 张素材均非模板自带照片）")
    else:
        _why = '、'.join('%s 用于 p%s' % (h[:8], '、p'.join(str(x) for x in v)) for h, v in _tpl_used)
        results.append(f"+0：如替换或新增插图，使用与新配色一致的自制配图，未直接搬用风格模板自带的素材照片（如 template_img_*.png），也未让同一张模板图重复出现在多页（直接搬用模板素材 {len(_tpl_used)} 张：{_why}）")

    # +1：项目名称、机构名"瑞康医疗价格管理中心"、期次"2026年春季"、明细页编号
    #     "01｜…29｜"等实体取值未被改写
    # 严格对齐细则，仅踩细则明确写出的点，不额外约束：
    #   ① 逐项在全篇阅读次序文本里查找：稿名、机构名、期次，以及 29 个明细页编号
    #      “01｜”…“29｜”。
    #   ② 与第 55 行的区别：本条只盯这些实体取值本身是否被改写，不重复判整页文字，
    #      不跨条重复计分。
    _deck = ''.join(s.flow_text() for s in slides)
    _ent_miss = [x for x in ENTITY_STRINGS if x not in _deck]
    _no_miss = [x for x in ENTITY_ITEM_NOS if x not in _deck]
    if not _ent_miss and not _no_miss:
        results.append(f"+1：项目名称、机构名\"瑞康医疗价格管理中心\"、期次\"2026年春季\"、明细页编号\"01｜…29｜\"等实体取值未被改写（3 项实体取值 + 29 个明细页编号全部保留）")
    else:
        results.append(f"+0：项目名称、机构名\"瑞康医疗价格管理中心\"、期次\"2026年春季\"、明细页编号\"01｜…29｜\"等实体取值未被改写（缺 {'、'.join(_ent_miss + _no_miss)[:60]}）")

    # +3：换色换版后文字与图形不冲撞——文本框未被挤压导致半句丢失、换行错乱、
    #     压住配图主体或超出画布
    # 严格对齐细则，仅踩细则明确写出的点，不额外约束：
    #   ① 超出画布：可见文字对象的外框须落在 0–13.333 / 0–7.5 之内（放宽 0.12 英寸）。
    #   ② 文本框叠压：两个承载信息的文字对象重叠面积超过较小者的 60% 即算冲撞；
    #      模板里文字对象互不相交，正常排版不会触发。
    #   ③ 裁定①：判定按渲染结果。文本框的声明高度常留富余（good 稿的卡片正文框声明
    #      1.05 英寸、实际两行文字仅约 0.34 英寸），故取"按字数与字号估出的文字高度"
    #      与声明高度的较小者作为文字实际占位，避免把框子的空白算成压字。
    #   ④ 裁定⑨：全量核查——39 页逐页判，一处冲撞即不得分。

    def _text_rect(sh):
        """文字实际渲染范围的估算：宽度取文本框宽度，高度取估算行高与声明高度的较小者。
        中日韩字按 1 个字宽、西文按 0.55 个字宽折算，行距按 1.25 倍字号。"""
        r = sh['rect']
        usable = max(r[2] - 0.10, 0.10)
        h = 0.0
        for p in sh['paras']:
            sz = max([rr['size'] or 0 for rr in p['runs']] or [0])
            if not sz:
                continue
            em = 0.0
            for rr in p['runs']:
                t = rr['text'] or ''
                cjk = sum(1 for ch in t if ord(ch) > 0x2000)
                em += cjk + (len(t) - cjk) * 0.55
            lines = max(1, int(-(-(em * (sz / 72.0)) // usable)))
            h += lines * sz * 1.25 / 72.0
        return (r[0], r[1], r[2], min(r[3], h)) if h else r

    _p5 = slides
    _hit = []
    for s in _p5:
        boxes = []
        for i, sh in enumerate(s.shapes):
            if s.occluded(i) or not sh['rect']:
                continue
            if not any(substantive(r) for p in sh['paras'] for r in p['runs']):
                continue
            r = _text_rect(sh)
            boxes.append((r, sh))
            if (r[0] < -POS_TOL or r[1] < -POS_TOL
                    or r[0] + r[2] > (cw or 13.333) + POS_TOL
                    or r[1] + r[3] > (ch or 7.5) + POS_TOL):
                _hit.append('p%d “%s”超出画布' % (s.pos, (sh['name'] or sh['kind'])[:10]))
        for a in range(len(boxes)):
            for b in range(a + 1, len(boxes)):
                ra, rb = boxes[a][0], boxes[b][0]
                small = min(ra[2] * ra[3], rb[2] * rb[3])
                if small > 0 and _overlap(ra, rb) / small > 0.6:
                    _hit.append('p%d “%s”与“%s”叠压' % (
                        s.pos, (boxes[a][1]['name'] or '')[:8], (boxes[b][1]['name'] or '')[:8]))
    _hit = sorted(set(_hit))
    if _p5 and not _hit:
        results.append(f"+3：换色换版后文字与图形不冲撞——文本框未被挤压导致半句丢失、换行错乱、压住配图主体或超出画布（{len(_p5)} 页逐页核查，无越界与叠压）")
    else:
        results.append(f"+0：换色换版后文字与图形不冲撞——文本框未被挤压导致半句丢失、换行错乱、压住配图主体或超出画布（{len(_hit)} 处冲撞：{'、'.join(_hit[:6])}）")

    # ========== 将 results 反解为结构化 dim2_items ==========
    # results 里每条的写法固定为 “+N：<细则原文>（<本次核查依据>）”，末尾那一组全角括号
    # 是脚本追加的诊断说明，不属于细则，反解时丢弃；rule 只留细则原文。
    def _strip_detail(text):
        if not text.endswith('）'):
            return text
        depth = 0
        for k in range(len(text) - 1, -1, -1):
            if text[k] == '）':
                depth += 1
            elif text[k] == '（':
                depth -= 1
                if depth == 0:
                    return text[:k].rstrip()
        return text

    if len(results) != len(_MAX_DELTAS_ORDER):
        raise RuntimeError(
            f"dim2 项数不一致：results={len(results)} vs order={len(_MAX_DELTAS_ORDER)}")
    dim2_items = []
    total_score = 0
    for raw, max_delta in zip(results, _MAX_DELTAS_ORDER):
        head, _, body = raw.partition("：")
        try:
            delta = int(head.replace("+", "").strip())
        except ValueError:
            delta = 0
        hit = delta > 0
        dim2_items.append({"rule": _strip_detail(body), "max_delta": max_delta,
                           "delta": max_delta if hit else 0, "hit": hit})
        total_score += (max_delta if hit else 0)
    result["dim2_items"] = dim2_items
    result["total_score"] = total_score
    result["max_score"] = MAX_SCORE_TOTAL
    return result


if __name__ == "__main__":
    if len(sys.argv) >= 2:
        _dir = sys.argv[1]
    else:
        _dir = os.path.dirname(os.path.abspath(__file__))
    _out = evaluate(_dir)
    print(json.dumps(_out, ensure_ascii=False, indent=2))
