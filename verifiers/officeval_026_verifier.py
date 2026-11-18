#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# pyright: reportUnknownMemberType=false, reportUnknownVariableType=false, reportUnknownArgumentType=false, reportUnknownParameterType=false, reportMissingParameterType=false, reportMissingTypeArgument=false, reportPrivateUsage=false, reportUnknownLambdaType=false, reportGeneralTypeIssues=false
"""
三线表（表5-4 至 表5-12 统计表）排版格式自动评估工具（officeval_026）

统一入口：
    evaluate(dir_path: str) -> dict
    - dir_path: 输出文件所在目录（内含待评估的 .docx，忽略以 ~$ 开头的 Word 临时锁文件）
    - 返回字段：id / file_name / status / error / dim1_pass / dim1_reason /
                dim2_items（命中与未命中项都返回）/ total_score / max_score
    - dim2_items 每项只含 rule / max_delta / delta / hit 四个键

规范约束：
    * 只暴露 evaluate 一个函数供批量运行器调用
    * 不修改 sys.stdout、不打印主结果、不使用 sys.exit
    * 不硬编码文档路径，全部由 dir_path 参数驱动
    * status='error' 表示这份提交没被真正评过（目录不存在、目录内无可评估文档、
      脚本自身异常），此时 dim1_pass 为空串（维度一未判定）、
      total_score 与 max_score 一律为 0；
      文档能定位到但不达标一律 status='ok'，由 dim1_pass / dim2_items 承载
    * dim1_pass 为 False 时 dim1_reason 必须给出未通过的错误摘要；通过时为空串
    * 本文件仅在 __main__ 下作为本地调试入口，才把 dict 序列化为 JSON 输出
"""

import os
import re
import sys
import json
from docx import Document
from docx.oxml.ns import qn

SCRIPT_ID = "026"
# 维度二所有加分项 max_delta 之和；用于维度一未通过 / 脚本异常时兜底
MAX_SCORE_TOTAL = 116

# ---------------------------------------------------------------------------
# 常量
# ---------------------------------------------------------------------------
# Word 边框线宽 w:sz 单位为 1/8 磅：1.5磅 = 12；0.75磅 = 6。
SZ_1_5PT = 12
SZ_0_75PT = 6
# 字号 w:sz 单位为 half-point：小五(9磅) = 18；五号(10.5磅) = 21。
SZ_XIAOWU = 18
SZ_WUHAO = 21
# 1cm = 567 twips
TWIPS_PER_CM = 567.0

# 9 张统计表的键（文档顺序）与展示用标签
ORDER = ['5-4', '5-5', '5-6', '5-7', '5-8', '5-9', '5-10', '5-11', '5-12']
def LAB(k):
    return '表' + k

# 含 “Levene检验”“t 检验” 两个分组标题的 7 列表
SEVEN = ['5-4', '5-5', '5-8', '5-9', '5-12']
# 仅含 “t 检验” 一个分组标题的 5 列表
FIVE = ['5-6', '5-7', '5-10', '5-11']
# 维度列需要纵向合并的表（表5-12 第1列为“组别”，无维度列）
DIM_COL = ['5-4', '5-5', '5-6', '5-7', '5-8', '5-9', '5-10', '5-11']

# 细则规定的行数、列数
SHAPE = {'5-4': (6, 7), '5-5': (6, 7), '5-6': (6, 5), '5-7': (6, 5), '5-8': (12, 7),
         '5-9': (12, 7), '5-10': (12, 5), '5-11': (12, 5), '5-12': (4, 7)}

# 9 条表题全文（细则逐字列出，检测时按“去空格”归一化后比较）
CAPS = {
    '5-4': '表5-4 干预组与参照组任务动机前测结果',
    '5-5': '表5-5 干预组与参照组任务动机后测结果',
    '5-6': '表5-6 干预组任务动机前后测比较',
    '5-7': '表5-7 参照组任务动机前后测比较',
    '5-8': '表5-8 干预组与参照组科学探究兴趣前测结果',
    '5-9': '表5-9 干预组与参照组科学探究兴趣后测结果',
    '5-10': '表5-10 干预组科学探究兴趣前后测结果',
    '5-11': '表5-11 参照组科学探究兴趣前后测结果',
    '5-12': '表5-12 干预组与参照组科学探究技能考核结果',
}
CAP_SET = set(CAPS.values())
# 文档中允许存在的唯一提示段
SAMPLE_HINT = '例子：'

# 各表维度名称
DIMS = {
    '5-4': ['任务价值', '学习投入'],
    '5-5': ['任务价值', '学习投入'],
    '5-6': ['任务价值', '学习投入'],
    '5-7': ['任务价值', '学习投入'],
    '5-8': ['探究兴趣', '学习顾虑', '课堂参与度', '课外实践参与', '学习资源关注'],
}
for _k in ('5-9', '5-10', '5-11'):
    DIMS[_k] = DIMS['5-8']
DIMS['5-12'] = []

# 各表组别 / 测试时间名称
GRP = {_k: ['干预组', '参照组'] for _k in ('5-4', '5-5', '5-8', '5-9', '5-12')}
for _k in ('5-6', '5-7', '5-10', '5-11'):
    GRP[_k] = ['干预前', '干预后']

# 表头必需文字；元组表示“三者之一即可”
HDR_REQ = {}
for _k in ('5-4', '5-5', '5-8', '5-9'):
    HDR_REQ[_k] = ['Levene检验', 't检验', '维度', '组别', 'X±S', 'F值', 'Sig.', 't值', 'Sig(双侧)']
HDR_REQ['5-12'] = ['Levene检验', 't检验', '组别', 'N', 'X±S', 'F值', 'Sig.', 't值', 'Sig(双侧)']
for _k in ('5-6', '5-7', '5-10', '5-11'):
    HDR_REQ[_k] = ['t检验', '维度', ('测试时间', '时间', '组别'), 'X±S', 't值', 'Sig(双侧)']

# 各表必须保留的数值（细则逐字列出，检测时按“去空格 + 归一化负号”后比较）
VALS = {
    '5-4': ['23.72±4.86', '23.45±4.93', '0.091', '0.764', '0.274', '0.785',
            '15.26±3.41', '15.44±3.36', '0.128', '0.722', '-0.265', '0.792'],
    '5-5': ['27.36±4.28', '24.68±4.51', '0.037', '0.848', '3.054', '0.003',
            '18.02±3.57', '15.98±3.49', '0.072', '0.789', '2.889', '0.005'],
    '5-6': ['23.72±4.86', '27.36±4.28', '-4.126', '0.001',
            '15.26±3.41', '18.02±3.57', '-3.184', '0.003'],
    '5-7': ['23.45±4.93', '24.68±4.51', '-1.248', '0.219',
            '15.44±3.36', '15.98±3.49', '-0.876', '0.386'],
    '5-8': ['20.18±4.32', '19.94±4.40', '0.316', '0.576', '0.273', '0.786',
            '26.74±5.08', '26.96±4.97', '0.104', '0.748', '-0.219', '0.827',
            '19.22±4.06', '19.08±4.13', '0.193', '0.662', '0.171', '0.865',
            '12.34±3.04', '12.18±3.11', '0.076', '0.784', '0.260', '0.796',
            '10.62±2.73', '10.48±2.69', '0.142', '0.708', '0.258', '0.797'],
    '5-9': ['25.08±5.21', '20.42±5.16', '0.232', '0.632', '4.492', '0.001',
            '22.46±6.38', '26.18±6.71', '0.063', '0.803', '-2.842', '0.006',
            '23.14±4.12', '19.66±4.24', '0.115', '0.736', '4.158', '0.001',
            '14.02±3.18', '12.76±3.09', '0.087', '0.769', '2.009', '0.048',
            '11.36±2.82', '10.74±2.76', '0.176', '0.676', '1.112', '0.270'],
    '5-10': ['20.18±4.32', '25.08±5.21', '-5.308', '0.001',
             '26.74±5.08', '22.46±6.38', '3.014', '0.004',
             '19.22±4.06', '23.14±4.12', '-4.806', '0.001',
             '12.34±3.04', '14.02±3.18', '-2.316', '0.025',
             '10.62±2.73', '11.36±2.82', '-1.248', '0.218'],
    '5-11': ['19.94±4.40', '20.42±5.16', '-0.524', '0.603',
             '26.96±4.97', '26.18±6.71', '0.462', '0.646',
             '19.08±4.13', '19.66±4.24', '-0.712', '0.480',
             '12.18±3.11', '12.76±3.09', '-0.638', '0.527',
             '10.48±2.69', '10.74±2.76', '-0.421', '0.676'],
    '5-12': ['40', '85.16±6.42', '78.34±6.96', '0.718', '0.399', '4.558', '0.001'],
}

# ---------------------------------------------------------------------------
# 维度二条目定义
#   加分项按“一个三线表满足此条 +N 分”逐表拆分为独立条目，使每个条目都满足
#   delta == max_delta（命中）/ 0（未命中）的接口约定；上限分 = 每表分值 × 表数。
#   减分项 max_delta 为负值，命中即扣。
# ---------------------------------------------------------------------------
_PLUS_SPECS = [
    (3, 'A1', ORDER,
     '表5-4至表5-12表格外框线：外框线只保留上下外框线，表格上框线和下框线为1.5磅黑色实线，'
     '左右两侧无外框线；且该设置在表格整体边框上完成，表格整体边框中的左框线、右框线、'
     '内部横框线、内部竖框线均为无线条'),
    (3, 'A2', SEVEN,
     '表5-4、表5-5、表5-8、表5-9和表5-12分组内框线：仅有“Levene检验”、“t检验”两个分组标题'
     '下方有下框线，其余位置无内框线，框线类型为0.75磅黑色实线'),
    (3, 'A3', FIVE,
     '表5-6、表5-7、表5-10和表5-11分组表头内框线：仅有“t检验”标题下方有下框线，'
     '其余位置无内框线，框线类型为0.75磅黑色实线'),
    (1, 'A4', ORDER,
     '表5-4至表5-12表格行列数：表5-4、表5-5为6行7列，表5-6、表5-7为6行5列，'
     '表5-8、表5-9为12行7列，表5-10、表5-11为12行5列，表5-12为4行7列'),
    (1, 'A5', SEVEN,
     '表5-4、表5-5、表5-8、表5-9和表5-12分组表头横向合并：第1行前3个单元格无文字且不合并，'
     '第4列单元格横向合并2列，第6列单元格横向合并2列，第2行7个单元格均不横向合并'),
    (1, 'A6', FIVE,
     '表5-6、表5-7、表5-10和表5-11分组表头横向合并：第1行前3个单元格无文字且不合并，'
     '第4列单元格横向合并2列，第2行5个单元格均不横向合并'),
    (1, 'A7', DIM_COL,
     '表5-4至表5-11维度列纵向合并：第1列每个维度名称占2个数据行并纵向合并为1个单元格，'
     '维度名称在合并区内只出现1次，第2个数据行对应位置无重复文字'),
    (1, 'A8', ORDER,
     '表5-4至表5-12统计量列纵向合并：“F值”“Sig.”“t值”“Sig(双侧)”各列在同属一个维度的'
     '2个数据行内纵向合并为1个单元格，同一数值只出现1次'),
    (1, 'A9', ORDER,
     '表5-4至表5-12表格宽度：表格宽度设置为按百分比自适应正文宽度、取值约为正文宽度的100%，'
     '表格实际总宽度不大于14.7cm，右边缘不越出右侧页边距'),
    (1, 'A10', ORDER,
     '表5-4至表5-12数据行行高：表5-4至表5-7的4个数据行中第1、3个数据行，'
     '表5-8至表5-11的10个数据行中第1、3、5、7、9个数据行，表5-12的2个数据行中第1个数据行，'
     '均设置0.69cm及以上的最小行高'),
    (1, 'A11', ORDER,
     '表5-4至表5-12表内字体与对齐：表内中文字体为宋体、西文与数字字体为Times New Roman，'
     '且表内所有有文字的单元格均为水平居中、垂直居中、首行缩进为0'),
]

_MINUS_SPECS = [
    (-1, 'D1', '纸张尺寸不为A4（21.0cm×29.7cm），或页边距不为上下2.54cm、左右约3.17cm'),
    (-3, 'D2', '“表5-4 干预组与参照组任务动机前测结果”“表5-5 干预组与参照组任务动机后测结果”'
     '“表5-6 干预组任务动机前后测比较”“表5-7 参照组任务动机前后测比较”'
     '“表5-8 干预组与参照组科学探究兴趣前测结果”“表5-9 干预组与参照组科学探究兴趣后测结果”'
     '“表5-10 干预组科学探究兴趣前后测结果”“表5-11 参照组科学探究兴趣前后测结果”'
     '“表5-12 干预组与参照组科学探究技能考核结果”9条表题中任一条被删除或文字被改写'),
    (-1, 'D3', '任一表题段落不为居中对齐，或带有首行缩进'),
    (-1, 'D4', '任一表题段落文字不为宋体（西文Times New Roman）、10.5磅（五号）、加粗'),
    (-3, 'D5', '表5-4至表5-12中任一表题段之后未紧跟1张独立的Word表格，'
     '表题与表格之间被插入空段落，或相邻两张统计表被合并为同一张表'),
    (-1, 'D6', '表5-4至表5-12中任一表内文字的字号不为9磅（小五）'),
    (-1, 'D7', '表5-4、表5-5、表5-8、表5-9、表5-12中“Levene检验”“t 检验”“X±S”“F值”“Sig.”'
     '“t值”“Sig(双侧)”任一表头文字被删除或改写，或表5-6、表5-7、表5-10、表5-11中'
     '“t 检验”“X±S”“t值”“Sig(双侧)”任一表头文字被删除或改写（表5-4至表5-11第1列列名为“维度”，'
     '表5-4、表5-5、表5-8、表5-9第2列列名为“组别”，表5-6、表5-7、表5-10、表5-11第2列列名为'
     '“测试时间”“时间”或“组别”，表5-12第1、2列列名为“组别”“N”）'),
    (-1, 'D8', '表5-4至表5-7中“任务价值”“学习投入”任一维度名称被删除或改写，'
     '或表5-8至表5-11中“探究兴趣”“学习顾虑”“课堂参与度”“课外实践参与”“学习资源关注”'
     '任一维度名称被删除或改写'),
    (-1, 'D9', '表5-4、表5-5、表5-8、表5-9、表5-12中“干预组”“参照组”两个组别名称任一被删除或改写，'
     '或表5-6、表5-7、表5-10、表5-11中“干预前”“干预后”两个测试时间名称任一被删除或改写'),
    (-3, 'V5-4', '表5-4中以下数值任一被删除或改写（任务价值维度：干预组“X±S”23.72±4.86、'
     '参照组“X±S”23.45±4.93、“F值”0.091、“Sig.”0.764、“t值”0.274、“Sig(双侧)”0.785；'
     '学习投入维度：干预组“X±S”15.26±3.41、参照组“X±S”15.44±3.36、“F值”0.128、“Sig.”0.722、'
     '“t值”-0.265、“Sig(双侧)”0.792）'),
    (-3, 'V5-5', '表5-5中以下数值任一被删除或改写（任务价值维度：干预组“X±S”27.36±4.28、'
     '参照组“X±S”24.68±4.51、“F值”0.037、“Sig.”0.848、“t值”3.054、“Sig(双侧)”0.003；'
     '学习投入维度：干预组“X±S”18.02±3.57、参照组“X±S”15.98±3.49、“F值”0.072、“Sig.”0.789、'
     '“t值”2.889、“Sig(双侧)”0.005）'),
    (-1, 'V5-6', '表5-6中以下数值任一被删除或改写（任务价值维度：干预前“X±S”23.72±4.86、'
     '干预后“X±S”27.36±4.28、“t值”-4.126、“Sig(双侧)”0.001；学习投入维度：干预前“X±S”15.26±3.41、'
     '干预后“X±S”18.02±3.57、“t值”-3.184、“Sig(双侧)”0.003）'),
    (-1, 'V5-7', '表5-7中以下数值任一被删除或改写（任务价值维度：干预前“X±S”23.45±4.93、'
     '干预后“X±S”24.68±4.51、“t值”-1.248、“Sig(双侧)”0.219；学习投入维度：干预前“X±S”15.44±3.36、'
     '干预后“X±S”15.98±3.49、“t值”-0.876、“Sig(双侧)”0.386）'),
    (-3, 'V5-8', '表5-8中以下数值任一被删除或改写（探究兴趣维度：干预组“X±S”20.18±4.32、'
     '参照组“X±S”19.94±4.40、“F值”0.316、“Sig.”0.576、“t值”0.273、“Sig(双侧)”0.786；'
     '学习顾虑维度：干预组“X±S”26.74±5.08、参照组“X±S”26.96±4.97、“F值”0.104、“Sig.”0.748、'
     '“t值”-0.219、“Sig(双侧)”0.827；课堂参与度维度：干预组“X±S”19.22±4.06、参照组“X±S”19.08±4.13、'
     '“F值”0.193、“Sig.”0.662、“t值”0.171、“Sig(双侧)”0.865；课外实践参与维度：'
     '干预组“X±S”12.34±3.04、参照组“X±S”12.18±3.11、“F值”0.076、“Sig.”0.784、“t值”0.260、'
     '“Sig(双侧)”0.796；学习资源关注维度：干预组“X±S”10.62±2.73、参照组“X±S”10.48±2.69、'
     '“F值”0.142、“Sig.”0.708、“t值”0.258、“Sig(双侧)”0.797）'),
    (-3, 'V5-9', '表5-9中以下数值任一被删除或改写（探究兴趣维度：干预组“X±S”25.08±5.21、'
     '参照组“X±S”20.42±5.16、“F值”0.232、“Sig.”0.632、“t值”4.492、“Sig(双侧)”0.001；'
     '学习顾虑维度：干预组“X±S”22.46±6.38、参照组“X±S”26.18±6.71、“F值”0.063、“Sig.”0.803、'
     '“t值”-2.842、“Sig(双侧)”0.006；课堂参与度维度：干预组“X±S”23.14±4.12、参照组“X±S”19.66±4.24、'
     '“F值”0.115、“Sig.”0.736、“t值”4.158、“Sig(双侧)”0.001；课外实践参与维度：'
     '干预组“X±S”14.02±3.18、参照组“X±S”12.76±3.09、“F值”0.087、“Sig.”0.769、“t值”2.009、'
     '“Sig(双侧)”0.048；学习资源关注维度：干预组“X±S”11.36±2.82、参照组“X±S”10.74±2.76、'
     '“F值”0.176、“Sig.”0.676、“t值”1.112、“Sig(双侧)”0.270）'),
    (-3, 'V5-10', '表5-10中以下数值任一被删除或改写（探究兴趣维度：干预前“X±S”20.18±4.32、'
     '干预后“X±S”25.08±5.21、“t值”-5.308、“Sig(双侧)”0.001；学习顾虑维度：干预前“X±S”26.74±5.08、'
     '干预后“X±S”22.46±6.38、“t值”3.014、“Sig(双侧)”0.004；课堂参与度维度：干预前“X±S”19.22±4.06、'
     '干预后“X±S”23.14±4.12、“t值”-4.806、“Sig(双侧)”0.001；课外实践参与维度：干预前“X±S”12.34±3.04、'
     '干预后“X±S”14.02±3.18、“t值”-2.316、“Sig(双侧)”0.025；学习资源关注维度：干预前“X±S”10.62±2.73、'
     '干预后“X±S”11.36±2.82、“t值”-1.248、“Sig(双侧)”0.218）'),
    (-3, 'V5-11', '表5-11中以下数值任一被删除或改写（探究兴趣维度：干预前“X±S”19.94±4.40、'
     '干预后“X±S”20.42±5.16、“t值”-0.524、“Sig(双侧)”0.603；学习顾虑维度：干预前“X±S”26.96±4.97、'
     '干预后“X±S”26.18±6.71、“t值”0.462、“Sig(双侧)”0.646；课堂参与度维度：干预前“X±S”19.08±4.13、'
     '干预后“X±S”19.66±4.24、“t值”-0.712、“Sig(双侧)”0.480；课外实践参与维度：干预前“X±S”12.18±3.11、'
     '干预后“X±S”12.76±3.09、“t值”-0.638、“Sig(双侧)”0.527；学习资源关注维度：干预前“X±S”10.48±2.69、'
     '干预后“X±S”10.74±2.76、“t值”-0.421、“Sig(双侧)”0.676）'),
    (-1, 'V5-12', '表5-12中以下数值任一被删除或改写（干预组行“N”40、“X±S”85.16±6.42；'
     '参照组行“N”40、“X±S”78.34±6.96；“F值”0.718、“Sig.”0.399、“t值”4.558、“Sig(双侧)”0.001）'),
    (-1, 'D19', '任一统计表单元格内文字被设置为加粗、倾斜或下划线'),
    (-1, 'D20', '任一统计表单元格被设置底纹或填充颜色'),
    (-1, 'D21', '任一统计表单元格文字颜色不为自动色或黑色'),
    (-3, 'D22', '任一统计表中出现整行单元格均无文字的空数据行'),
    (-3, 'D23', '任一统计表的数据区出现“维度+组别+X±S”完全相同的重复数据行'),
    (-1, 'D24', '文档被添加页眉或页脚内容'),
    (-1, 'D25', '文档中出现除9条表题与“例子：”提示段以外的新增文字段落'
     '（如修改说明、生成备注、版本标注等）'),
    (-1, 'D26', '文档中被插入分页符，或分节符多于1处，统计表被强制分页'),
]

# 维度二条目的固定顺序：先 11 条加分项（逐表拆分），再 26 条减分项
_ITEMS = []
for _per, _key, _labels, _rule in _PLUS_SPECS:
    for _lab in _labels:
        _ITEMS.append((_key, _lab, _per, '%s（%s）' % (_rule, LAB(_lab))))
for _delta, _key, _rule in _MINUS_SPECS:
    _ITEMS.append((_key, None, _delta, _rule))

_MAX_DELTAS_ORDER = [_it[2] for _it in _ITEMS]
assert sum(x for x in _MAX_DELTAS_ORDER if x > 0) == MAX_SCORE_TOTAL
assert len(_MAX_DELTAS_ORDER) == 80 + 26


# ---------------------------------------------------------------------------
# 底层工具
# ---------------------------------------------------------------------------
def _nz(s):
    """归一化：去掉半角/全角空格，全角括号转半角，各类连字符统一为 ASCII 减号。"""
    if not s:
        return ''
    s = s.replace(' ', '').replace('　', '').replace('\t', '')
    s = s.replace('（', '(').replace('）', ')')
    for ch in ('−', '–', '—', '－'):
        s = s.replace(ch, '-')
    return s


def _txt_of(el):
    """取元素下所有 w:t 的拼接文本。"""
    return ''.join(n.text or '' for n in el.iter(qn('w:t')))


def _border(el, name):
    """读取 w:tblBorders / w:tcBorders 下某一方向的 (val, sz, color)。"""
    if el is None:
        return None
    e = el.find(qn('w:' + name))
    if e is None:
        return None
    try:
        sz = int(e.get(qn('w:sz')) or 0)
    except ValueError:
        sz = 0
    return (e.get(qn('w:val')), sz, (e.get(qn('w:color')) or 'auto').lower())


def _line(b):
    """是否为“有线条”。val 为 none/nil 或宽度为 0 均视为无线条。"""
    return b is not None and b[0] not in (None, 'none', 'nil') and b[1] > 0


def _solid_black(b, sz):
    """是否为指定磅值的黑色实线。"""
    return (_line(b) and b[0] == 'single' and b[1] == sz
            and b[2] in ('auto', '000000', '000'))


def _cells_of(tr):
    """解析一行的所有 w:tc；col 为该单元格起始网格列号（累加 gridSpan）。"""
    out = []
    col = 0
    for tc in tr.findall(qn('w:tc')):
        tcPr = tc.find(qn('w:tcPr'))
        gs = 1
        vm = None
        va = None
        tb = {}
        if tcPr is not None:
            g = tcPr.find(qn('w:gridSpan'))
            if g is not None:
                gs = int(g.get(qn('w:val')))
            e = tcPr.find(qn('w:vMerge'))
            if e is not None:
                vm = e.get(qn('w:val')) or 'continue'
            a = tcPr.find(qn('w:vAlign'))
            if a is not None:
                va = a.get(qn('w:val'))
            bd = tcPr.find(qn('w:tcBorders'))
            if bd is not None:
                for name in ('top', 'bottom', 'left', 'right', 'insideH', 'insideV'):
                    b = _border(bd, name)
                    if b is not None:
                        tb[name] = b
        out.append({'tc': tc, 'col': col, 'gs': gs, 'txt': _txt_of(tc).strip(),
                    'vm': vm, 'va': va, 'tb': tb})
        col += gs
    return out


def _hlines(rows):
    """收集所有内部横线，按“边界序号 + 网格列号”去重。

    同一条视觉横线既可写成第 i 行单元格的下框线，也可写成第 i+1 行单元格的上框线，
    因此两种写法都要归并到同一个 (边界序号, 列号) 上。返回 (线集合, 是否全为0.75磅黑色实线)。
    """
    got = set()
    ok_sz = True
    for i in range(len(rows) - 1):
        for c in rows[i]:
            b = c['tb'].get('bottom')
            if _line(b):
                for x in range(c['col'], c['col'] + c['gs']):
                    got.add((i, x))
                if not _solid_black(b, SZ_0_75PT):
                    ok_sz = False
        for c in rows[i + 1]:
            b = c['tb'].get('top')
            if _line(b):
                for x in range(c['col'], c['col'] + c['gs']):
                    got.add((i, x))
                if not _solid_black(b, SZ_0_75PT):
                    ok_sz = False
    return got, ok_sz


def _vlines(rows):
    """统计内部竖线条数（行首左框线、行尾右框线属于外框，不计入）。"""
    n = 0
    for r in rows:
        for i, c in enumerate(r):
            if i > 0 and _line(c['tb'].get('left')):
                n += 1
            if i < len(r) - 1 and _line(c['tb'].get('right')):
                n += 1
    return n


def _table_info(tbl):
    """把一张 w:tbl 解析成检测所需的字典。"""
    rows = [_cells_of(tr) for tr in tbl.findall(qn('w:tr'))]
    pr = tbl.find(qn('w:tblPr'))
    bd = pr.find(qn('w:tblBorders')) if pr is not None else None
    tb = {}
    for name in ('top', 'bottom', 'left', 'right', 'insideH', 'insideV'):
        b = _border(bd, name)
        if b is not None:
            tb[name] = b
    wel = pr.find(qn('w:tblW')) if pr is not None else None
    grid = tbl.find(qn('w:tblGrid'))
    cols = [int(g.get(qn('w:w')) or 0) for g in grid.findall(qn('w:gridCol'))] if grid is not None else []
    hs = []
    for tr in tbl.findall(qn('w:tr')):
        trPr = tr.find(qn('w:trPr'))
        h = None
        if trPr is not None:
            he = trPr.find(qn('w:trHeight'))
            if he is not None:
                h = (int(he.get(qn('w:val')) or 0) / TWIPS_PER_CM, he.get(qn('w:hRule')))
        hs.append(h)
    try:
        wval = int(wel.get(qn('w:w')) or 0) if wel is not None else 0
    except ValueError:
        wval = 0
    return {'el': tbl, 'rows': rows, 'tb': tb, 'hs': hs, 'grid': cols, 'ncol': len(cols),
            'wtype': wel.get(qn('w:type')) if wel is not None else None, 'wval': wval}


def _map_labeled_tables(body):
    """按“表题段落 + 紧随其后的表格”把 9 张统计表映射到表号。

    表题与表格之间允许夹入空段落（该情形由减分项 D5 单独判定），此处只负责定位。
    """
    T = {}
    pending = None
    for ch in body.iterchildren():
        if ch.tag == qn('w:p'):
            t = _nz(_txt_of(ch).strip())
            for k, cap in CAPS.items():
                if t == _nz(cap):
                    pending = k
                    break
            else:
                if t:
                    pending = None
        elif ch.tag == qn('w:tbl'):
            if pending is not None and pending not in T:
                T[pending] = _table_info(ch)
            pending = None
    return T


# ---------------------------------------------------------------------------
# 有效格式解析：run 直接格式 -> 段落样式链（basedOn） -> docDefaults
# ---------------------------------------------------------------------------
def _style_maps(doc):
    st_el = doc.styles.element
    by_id = {}
    for st in st_el.findall(qn('w:style')):
        by_id[st.get(qn('w:styleId'))] = st
    dd_r = dd_p = None
    dd = st_el.find(qn('w:docDefaults'))
    if dd is not None:
        rd = dd.find(qn('w:rPrDefault'))
        if rd is not None:
            dd_r = rd.find(qn('w:rPr'))
        pd = dd.find(qn('w:pPrDefault'))
        if pd is not None:
            dd_p = pd.find(qn('w:pPr'))
    return (by_id, dd_r, dd_p)


def _para_style_id(p):
    pPr = p.find(qn('w:pPr')) if p is not None else None
    if pPr is None:
        return None
    ps = pPr.find(qn('w:pStyle'))
    return ps.get(qn('w:val')) if ps is not None else None


def _walk_styles(getter, sid, by_id, dflt):
    seen = set()
    while sid and sid not in seen:
        seen.add(sid)
        st = by_id.get(sid)
        if st is None:
            break
        v = getter(st)
        if v is not None:
            return v
        b = st.find(qn('w:basedOn'))
        sid = b.get(qn('w:val')) if b is not None else None
    return dflt


def _resolve_run(getter, r, p, sm):
    by_id, dd_r, _ = sm
    v = getter(r.find(qn('w:rPr'))) if r is not None else None
    if v is not None:
        return v
    v = _walk_styles(lambda st: getter(st.find(qn('w:rPr'))), _para_style_id(p), by_id, None)
    if v is not None:
        return v
    return getter(dd_r)


def _resolve_para(getter, p, sm):
    by_id, _, dd_p = sm
    v = getter(p.find(qn('w:pPr'))) if p is not None else None
    if v is not None:
        return v
    v = _walk_styles(lambda st: getter(st.find(qn('w:pPr'))), _para_style_id(p), by_id, None)
    if v is not None:
        return v
    return getter(dd_p)


def _g_eastasia(rPr):
    if rPr is None:
        return None
    rf = rPr.find(qn('w:rFonts'))
    return rf.get(qn('w:eastAsia')) if rf is not None else None


def _g_ascii(rPr):
    if rPr is None:
        return None
    rf = rPr.find(qn('w:rFonts'))
    if rf is None:
        return None
    return rf.get(qn('w:ascii')) or rf.get(qn('w:hAnsi'))


def _g_sz(rPr):
    if rPr is None:
        return None
    e = rPr.find(qn('w:sz'))
    if e is None:
        return None
    try:
        return int(e.get(qn('w:val')))
    except (TypeError, ValueError):
        return None


def _g_bold(rPr):
    if rPr is None:
        return None
    e = rPr.find(qn('w:b'))
    if e is None:
        return None
    return e.get(qn('w:val')) not in ('0', 'false', 'off')


def _g_jc(pPr):
    if pPr is None:
        return None
    e = pPr.find(qn('w:jc'))
    return e.get(qn('w:val')) if e is not None else None


def _g_firstline(pPr):
    """返回首行缩进（twips 或字符数），未设置返回 None。"""
    if pPr is None:
        return None
    e = pPr.find(qn('w:ind'))
    if e is None:
        return None
    for attr in ('w:firstLine', 'w:firstLineChars'):
        v = e.get(qn(attr))
        if v is not None:
            try:
                return float(v)
            except ValueError:
                return None
    return None


def _parent_para(el):
    """向上找到 run 所属的 w:p。"""
    node = el.getparent()
    while node is not None and node.tag != qn('w:p'):
        node = node.getparent()
    return node


# ---------------------------------------------------------------------------
# 加分项 A1~A11 的单表判定
# ---------------------------------------------------------------------------
def _chk_a1_outer(t):
    """A1 外框线：整体边框仅上下 1.5 磅黑色实线，左右与内部横竖框线均无线条。"""
    tb = t['tb']
    bad = []
    if not _solid_black(tb.get('top'), SZ_1_5PT):
        bad.append('表格整体上框线不是1.5磅黑色实线')
    if not _solid_black(tb.get('bottom'), SZ_1_5PT):
        bad.append('表格整体下框线不是1.5磅黑色实线')
    for name, zh in (('left', '左框线'), ('right', '右框线'),
                     ('insideH', '内部横框线'), ('insideV', '内部竖框线')):
        if _line(tb.get(name)):
            bad.append('表格整体边框的%s有线条' % zh)
    return (not bad), '；'.join(bad)


def _chk_group_line(k, t):
    """A2/A3 分组标题下方内框线：仅第1行分组标题下方有 0.75 磅黑色实线，其余无内框线。"""
    ncol = SHAPE[k][1]
    if (len(t['rows']), t['ncol']) != SHAPE[k]:
        return False, '行列数不符（实际%d行%d列），无法判定分组内框线' % (len(t['rows']), t['ncol'])
    got, ok_sz = _hlines(t['rows'])
    expect = {(0, c) for c in range(3, ncol)}
    bad = []
    miss = sorted(expect - got)
    extra = sorted(got - expect)
    if miss:
        bad.append('分组标题下方缺下框线（边界/列号 %s）' % miss)
    if extra:
        bad.append('存在多余内部横框线（边界/列号 %s）' % extra)
    if not ok_sz:
        bad.append('内框线不是0.75磅黑色实线')
    nv = _vlines(t['rows'])
    if nv:
        bad.append('存在%d处内部竖框线' % nv)
    return (not bad), '；'.join(bad)


def _chk_a4_shape(k, t):
    """A4 行列数。"""
    want = SHAPE[k]
    got = (len(t['rows']), t['ncol'])
    if got == want:
        return True, ''
    return False, '要求%d行%d列，实际%d行%d列' % (want[0], want[1], got[0], got[1])


def _chk_header_merge(k, t):
    """A5/A6 分组表头横向合并：第1行前3个单元格不合并且无文字，第4列（7列表另加第6列）横向合并2列，第2行均不合并。"""
    ncol = SHAPE[k][1]
    rows = t['rows']
    if (len(rows), t['ncol']) != SHAPE[k]:
        return False, '行列数不符（实际%d行%d列），无法判定表头合并' % (len(rows), t['ncol'])
    r0, r1 = rows[0], rows[1]
    need_cells = 5 if ncol == 7 else 4
    bad = []
    if len(r0) != need_cells:
        bad.append('第1行应为%d个单元格，实际%d个' % (need_cells, len(r0)))
    else:
        for i in range(3):
            if r0[i]['gs'] != 1:
                bad.append('第1行第%d个单元格被横向合并' % (i + 1))
            if r0[i]['txt']:
                bad.append('第1行第%d个单元格有文字“%s”' % (i + 1, r0[i]['txt']))
        if not (r0[3]['col'] == 3 and r0[3]['gs'] == 2):
            bad.append('第4列单元格未横向合并2列')
        if ncol == 7 and not (r0[4]['col'] == 5 and r0[4]['gs'] == 2):
            bad.append('第6列单元格未横向合并2列')
    if len(r1) != ncol or any(c['gs'] != 1 for c in r1):
        bad.append('第2行未保持%d个不合并的单元格' % ncol)
    return (not bad), '；'.join(bad)


def _chk_a7_dim_merge(k, t):
    """A7 维度列纵向合并：第1列每个维度占2个数据行合并为1格，第2个数据行无重复文字。"""
    rows = t['rows']
    if (len(rows), t['ncol']) != SHAPE[k]:
        return False, '行列数不符（实际%d行%d列），无法判定维度列合并' % (len(rows), t['ncol'])
    n = len(rows)
    seq, txts = [], []
    for i in range(2, n):
        c0 = rows[i][0] if rows[i] and rows[i][0]['col'] == 0 else None
        seq.append(c0['vm'] if c0 else 'MISSING')
        txts.append(c0['txt'] if c0 else '')
    want = ['restart', 'continue'] * ((n - 2) // 2)
    bad = []
    if seq != want:
        bad.append('第1列纵向合并序列不符（实际 %s）' % '/'.join(str(x) for x in seq))
    dup = [txts[i] for i in range(1, len(txts), 2) if txts[i]]
    if dup:
        bad.append('维度名称在第2个数据行重复出现：%s' % '、'.join(dup))
    return (not bad), '；'.join(bad)


def _chk_a8_stat_merge(k, t):
    """A8 统计量列纵向合并：第4列起各统计量列在同一维度的2个数据行内合并为1格。"""
    rows = t['rows']
    if (len(rows), t['ncol']) != SHAPE[k]:
        return False, '行列数不符（实际%d行%d列），无法判定统计量列合并' % (len(rows), t['ncol'])
    n = len(rows)
    want = ['restart', 'continue'] * ((n - 2) // 2)
    bad = []
    for ci in range(3, SHAPE[k][1]):
        seq = []
        for i in range(2, n):
            found = None
            for c in rows[i]:
                if c['col'] <= ci < c['col'] + c['gs']:
                    found = c
            seq.append(found['vm'] if found else 'MISSING')
        if seq != want:
            bad.append('第%d列纵向合并序列不符（实际 %s）'
                       % (ci + 1, '/'.join(str(x) for x in seq)))
    return (not bad), '；'.join(bad)


def _chk_a9_width(t, text_width_cm):
    """A9 表格宽度：按百分比自适应正文宽度、约 100%；此时实际宽度等于正文宽度，需不大于 14.7cm。"""
    if t['wtype'] != 'pct':
        return False, '表格宽度未设置为按百分比自适应（实际类型 %s，值 %d）' % (t['wtype'], t['wval'])
    if not (4900 <= t['wval'] <= 5100):
        return False, '百分比宽度不在100%%附近（实际 %.1f%%）' % (t['wval'] / 50.0)
    if text_width_cm > 14.7 + 1e-6:
        return False, '正文宽度%.2fcm已超过14.7cm' % text_width_cm
    return True, ''


def _chk_a10_height(k, t):
    """A10 数据行行高：每个维度的第1个数据行须设置 0.69cm 及以上的最小行高。"""
    rows = t['rows']
    ndata = len(rows) - 2
    if ndata != SHAPE[k][0] - 2:
        return False, '数据行数不符（要求%d行，实际%d行）' % (SHAPE[k][0] - 2, ndata)
    bad = []
    for j in range(0, ndata, 2):
        h = t['hs'][2 + j]
        if h is None:
            bad.append('第%d个数据行未设置行高' % (j + 1))
        elif h[1] != 'atLeast':
            bad.append('第%d个数据行行高规则不是“最小值”' % (j + 1))
        elif h[0] < 0.69 - 0.005:
            bad.append('第%d个数据行最小行高仅%.2fcm' % (j + 1, h[0]))
    return (not bad), '；'.join(bad)


def _join(items, limit=6):
    """去重并截断明细，避免 detail 过长。"""
    out = []
    for x in items:
        if x not in out:
            out.append(x)
    if len(out) > limit:
        return '；'.join(out[:limit]) + '；等%d项' % len(out)
    return '；'.join(out)


def _chk_a11_font_align(t, sm):
    """A11 表内字体与对齐：中文宋体、西文与数字 Times New Roman，有文字单元格水平垂直居中且首行缩进 0。"""
    bad = []
    for r in t['el'].iter(qn('w:r')):
        if not _txt_of(r).strip():
            continue
        p = _parent_para(r)
        ea = _resolve_run(_g_eastasia, r, p, sm)
        asc = _resolve_run(_g_ascii, r, p, sm)
        if ea != '宋体':
            bad.append('中文字体为“%s”' % (ea or '未设置'))
        if asc != 'Times New Roman':
            bad.append('西文字体为“%s”' % (asc or '未设置'))
    for row in t['rows']:
        for c in row:
            if not c['txt']:
                continue
            tag = c['txt'][:8]
            if c['va'] != 'center':
                bad.append('单元格“%s”未垂直居中' % tag)
            for p in c['tc'].iter(qn('w:p')):
                if not _txt_of(p).strip():
                    continue
                if _resolve_para(_g_jc, p, sm) != 'center':
                    bad.append('单元格“%s”未水平居中' % tag)
                fl = _resolve_para(_g_firstline, p, sm)
                if fl is not None and fl > 0:
                    bad.append('单元格“%s”首行缩进不为0' % tag)
    return (not bad), _join(bad)


# ---------------------------------------------------------------------------
# 减分项 D1~D26 判定
# ---------------------------------------------------------------------------
def _chk_d1_page(doc):
    sec = doc.sections[0]
    want = [('纸张宽度', sec.page_width, 21.0), ('纸张高度', sec.page_height, 29.7),
            ('上边距', sec.top_margin, 2.54), ('下边距', sec.bottom_margin, 2.54),
            ('左边距', sec.left_margin, 3.17), ('右边距', sec.right_margin, 3.17)]
    bad = []
    for name, val, exp in want:
        cm = (val.cm if val is not None else 0.0)
        if abs(cm - exp) > 0.05:
            bad.append('%s为%.2fcm（应为%.2fcm）' % (name, cm, exp))
    return bool(bad), _join(bad)


def _chk_d2_caps(body_paras_norm):
    miss = [CAPS[k] for k in ORDER if _nz(CAPS[k]) not in body_paras_norm]
    return bool(miss), ('缺失或被改写的表题：' + '、'.join(miss)) if miss else ''


def _chk_d3_cap_align(cap_paras, sm):
    bad = []
    for k, p in cap_paras:
        if _resolve_para(_g_jc, p, sm) != 'center':
            bad.append('%s表题未居中' % LAB(k))
        fl = _resolve_para(_g_firstline, p, sm)
        if fl is not None and fl > 0:
            bad.append('%s表题有首行缩进' % LAB(k))
    return bool(bad), _join(bad)


def _chk_d4_cap_font(cap_paras, sm):
    bad = []
    for k, p in cap_paras:
        for r in p.iter(qn('w:r')):
            if not _txt_of(r).strip():
                continue
            ea = _resolve_run(_g_eastasia, r, p, sm)
            asc = _resolve_run(_g_ascii, r, p, sm)
            sz = _resolve_run(_g_sz, r, p, sm)
            bold = _resolve_run(_g_bold, r, p, sm)
            if ea != '宋体':
                bad.append('%s表题中文字体为“%s”' % (LAB(k), ea or '未设置'))
            if asc != 'Times New Roman':
                bad.append('%s表题西文字体为“%s”' % (LAB(k), asc or '未设置'))
            if sz != SZ_WUHAO:
                bad.append('%s表题字号为%.1f磅' % (LAB(k), (sz or 0) / 2.0))
            if not bold:
                bad.append('%s表题未加粗' % LAB(k))
    return bool(bad), _join(bad)


def _chk_d5_follow(seq, T):
    """表题段之后必须紧跟 1 张独立表格；不得夹入空段落，不得两表合并。"""
    bad = []
    for i, (kind, txt) in enumerate(seq):
        if kind != 'P':
            continue
        for k in ORDER:
            if _nz(CAPS[k]) == _nz(txt):
                if i + 1 >= len(seq) or seq[i + 1][0] != 'T':
                    bad.append('%s表题后未紧跟表格' % LAB(k))
                break
    missing = [LAB(k) for k in ORDER if k not in T]
    if missing:
        bad.append('未定位到独立表格：' + '、'.join(missing))
    ids = [id(T[k]['el']) for k in ORDER if k in T]
    if len(ids) != len(set(ids)):
        bad.append('存在多条表题共用同一张表格（统计表被合并）')
    return bool(bad), _join(bad)


def _chk_d6_size(T, sm):
    bad = []
    for k in ORDER:
        t = T.get(k)
        if t is None:
            continue
        for r in t['el'].iter(qn('w:r')):
            if not _txt_of(r).strip():
                continue
            sz = _resolve_run(_g_sz, r, _parent_para(r), sm)
            if sz != SZ_XIAOWU:
                bad.append('%s表内字号为%.1f磅' % (LAB(k), (sz or 0) / 2.0))
    return bool(bad), _join(bad)


def _chk_presence(T, need_map, prefix):
    """按“文字是否仍存在于该表中”判定删除/改写；文本先归一化。"""
    bad = []
    for k in ORDER:
        needs = need_map.get(k, [])
        if not needs:
            continue
        t = T.get(k)
        if t is None:
            bad.append('%s未定位到表格' % LAB(k))
            continue
        full = _nz(_txt_of(t['el']))
        for need in needs:
            if isinstance(need, tuple):
                if not any(_nz(x) in full for x in need):
                    bad.append('%s缺“%s”' % (LAB(k), '/'.join(need)))
            elif _nz(need) not in full:
                bad.append('%s缺“%s”' % (LAB(k), need))
    return bool(bad), (prefix + _join(bad, 8) if bad else '')


def _chk_values(k, T):
    """某张表内规定数值是否仍完整存在。"""
    t = T.get(k)
    if t is None:
        return True, '%s未定位到表格，规定数值全部缺失' % LAB(k)
    full = _nz(_txt_of(t['el']))
    miss = [v for v in VALS[k] if _nz(v) not in full]
    if not miss:
        return False, ''
    return True, '缺失或被改写的数值：' + _join(miss, 8)



def _chk_d19_d20_d21(T):
    """D19 加粗/倾斜/下划线、D20 底纹、D21 文字颜色。"""
    b19, b20, b21 = [], [], []
    for k in ORDER:
        t = T.get(k)
        if t is None:
            continue
        for r in t['el'].iter(qn('w:r')):
            rPr = r.find(qn('w:rPr'))
            if rPr is None:
                continue
            for tag, zh in (('b', '加粗'), ('i', '倾斜')):
                e = rPr.find(qn('w:' + tag))
                if e is not None and e.get(qn('w:val')) not in ('0', 'false', 'off'):
                    b19.append('%s存在%s文字' % (LAB(k), zh))
            u = rPr.find(qn('w:u'))
            if u is not None and u.get(qn('w:val')) not in (None, 'none'):
                b19.append('%s存在下划线文字' % LAB(k))
            c = rPr.find(qn('w:color'))
            if c is not None and (c.get(qn('w:val')) or 'auto').lower() not in ('auto', '000000', '000'):
                b21.append('%s文字颜色为%s' % (LAB(k), c.get(qn('w:val'))))
        for s in t['el'].iter(qn('w:shd')):
            fill = (s.get(qn('w:fill')) or 'auto').lower()
            if fill not in ('auto', 'ffffff') or (s.get(qn('w:val')) or 'clear') not in ('clear', 'nil'):
                b20.append('%s存在底纹/填充（fill=%s）' % (LAB(k), s.get(qn('w:fill'))))
    return ((bool(b19), _join(b19)), (bool(b20), _join(b20)), (bool(b21), _join(b21)))


def _chk_d22_d23(T):
    """D22 空数据行、D23 重复数据行。"""
    b22, b23 = [], []
    for k in ORDER:
        t = T.get(k)
        if t is None:
            continue
        rows = t['rows']
        for i in range(2, len(rows)):
            if rows[i] and all(c['txt'] == '' and c['vm'] is None for c in rows[i]):
                b22.append('%s第%d个数据行整行无文字' % (LAB(k), i - 1))
        keys = [tuple(c['txt'] for c in rows[i]) for i in range(2, len(rows))]
        if len(keys) != len(set(keys)):
            b23.append('%s数据区存在完全相同的重复数据行' % LAB(k))
    return (bool(b22), _join(b22)), (bool(b23), _join(b23))


def _chk_d24_hf(doc):
    """D24 页眉页脚内容。"""
    found = []
    for si, sec in enumerate(doc.sections):
        for attr, zh in (('header', '页眉'), ('footer', '页脚'),
                         ('first_page_header', '首页页眉'), ('first_page_footer', '首页页脚'),
                         ('even_page_header', '偶数页页眉'), ('even_page_footer', '偶数页页脚')):
            part = getattr(sec, attr, None)
            if part is None:
                continue
            try:
                txt = ''.join(p.text for p in part.paragraphs).strip()
            except Exception:
                txt = ''
            if txt:
                found.append('第%d节%s含文字“%s”' % (si + 1, zh, txt[:20]))
    return bool(found), _join(found)


def _chk_d25_extra(seq):
    """D25 除 9 条表题与“例子：”提示段以外的新增文字段落。"""
    allow = {_nz(c) for c in CAP_SET}
    allow.add(_nz(SAMPLE_HINT))
    extra = []
    for kind, txt in seq:
        if kind != 'P':
            continue
        t = txt.strip()
        if t and _nz(t) not in allow:
            extra.append(t[:24])
    return bool(extra), ('新增文字段落：' + _join(extra, 6)) if extra else ''


def _chk_d26_break(body):
    """D26 分页符，或分节符多于 1 处。"""
    pgbr = [b for b in body.iter(qn('w:br')) if b.get(qn('w:type')) == 'page']
    pgbr += [e for e in body.iter(qn('w:pageBreakBefore'))
             if e.get(qn('w:val')) not in ('0', 'false', 'off')]
    sects = body.findall('.//' + qn('w:sectPr'))
    msg = []
    if pgbr:
        msg.append('存在%d处分页符' % len(pgbr))
    if len(sects) > 1:
        msg.append('存在%d处分节符' % len(sects))
    return bool(msg), '；'.join(msg)


# ---------------------------------------------------------------------------
# 对外接口
# ---------------------------------------------------------------------------
def _open_word_document(file_path):
    """打开 .docx；损坏或非法包会抛出异常，由调用方转为维度一未通过。"""
    return Document(file_path)


def _build_dim2_items(hits):
    """按 _ITEMS 固定顺序生成严格二值的评分明细，返回 (items, total_score)。

    hits 为 {条目键: (是否命中, 判定明细)}；条目键在加分项为 (key, 表号)、
    减分项为 key。缺项直接抛错，以免条目表与判定点悄悄漂移后仍算出一个
    看似正常的分数。各 _chk_* 函数给出的明细只用于内部校验，不再输出：
    每项只含 rule / max_delta / delta / hit 四个键，判定说明由 rule 原文承载。
    """
    if len(_ITEMS) != len(_MAX_DELTAS_ORDER):
        raise RuntimeError('dim2 项数不一致：items=%d vs order=%d'
                           % (len(_ITEMS), len(_MAX_DELTAS_ORDER)))
    items = []
    total_score = 0
    for (key, lab, max_delta, rule), order_delta in zip(_ITEMS, _MAX_DELTAS_ORDER):
        if max_delta != order_delta:
            raise RuntimeError('dim2 分值与顺序表不一致：%s 记为 %d，顺序表为 %d'
                               % (key, max_delta, order_delta))
        slot = (key, lab) if lab is not None else key
        if slot not in hits:
            raise RuntimeError('dim2 条目缺少判定结果：%s' % (slot,))
        hit, _detail = hits[slot]
        hit = bool(hit)
        delta = max_delta if hit else 0
        items.append({'rule': rule, 'max_delta': max_delta, 'delta': delta,
                      'hit': hit})
        total_score += delta
    return items, total_score


def evaluate(dir_path: str) -> dict:
    """统一入口：接收输出文件所在目录，定位并评估目录内的三线表文档。

    评分口径：维度一为门槛项，未通过则 dim2_items 置空、total_score 记 0，
    max_score 仍为加分项满分之和（MAX_SCORE_TOTAL）；维度二逐条二值判定，
    total_score 为加分项得分减去减分项扣分。

    status='error' 只在“这份提交没被真正评过”时出现——目录不存在、目录内没有
    可评估的 Word 文档、或脚本自身抛异常；此时 dim1_pass 为空串（未判定）、
    dim1_reason 为空串、dim2_items 为空表、total_score 与 max_score 一律为 0。
    文档能定位到但打不开（损坏包），仍按 status='ok' + 维度一未通过处理，
    此时 dim1_pass=False 且 dim1_reason 带有未通过的错误摘要。

    参数:
        dir_path: 输出文件所在目录路径。
    返回:
        dict，字段对齐统一接口约定 §2.2。
    """
    result: dict[str, object] = {
        'id': SCRIPT_ID,
        'file_name': '',
        'status': 'ok',
        'error': None,
        # 空串表示“维度一尚未判定”；status='error' 的三条出口都保持空串，
        # 只有真正跑过维度一门槛判定才会写入 True / False。
        'dim1_pass': '',
        'dim1_reason': '',
        'dim2_items': [],
        'total_score': 0,
        'max_score': 0,
    }

    try:
        if not os.path.isdir(dir_path):
            result['status'] = 'error'
            result['error'] = f'目录不存在：{dir_path}'
            return result

        docxs = [f for f in sorted(os.listdir(dir_path))
                 if f.lower().endswith('.docx') and not f.startswith('~$')
                 and os.path.isfile(os.path.join(dir_path, f))]
        if not docxs:
            result['status'] = 'error'
            result['error'] = f'目录中未找到 .docx 文件：{dir_path}'
            return result

        file_path = os.path.join(dir_path, docxs[0])
        result['file_name'] = os.path.basename(file_path)

        # 维度一：交付文件能否被当作 .docx 正常打开。打不开是“评了但不达标”，
        # 按门槛短路约定返回 status='ok' + dim1_pass=False + dim2_items=[]，
        # 且 dim1_reason 必须带上未通过的错误摘要。
        try:
            doc = _open_word_document(file_path)
        except Exception as exc:
            result['dim1_pass'] = False
            result['dim1_reason'] = '.docx 文件损坏无法打开（%s: %s）' % (type(exc).__name__, exc)
            result['dim2_items'] = []
            result['total_score'] = 0
            result['max_score'] = MAX_SCORE_TOTAL
            return result

        # 维度一通过：按统一约定 dim1_reason 保持空串（只有未通过才写原因）。
        result['dim1_pass'] = True
        result['dim1_reason'] = ''
        return _evaluate_impl(doc, result)

    except Exception as exc:
        # _evaluate_impl 是就地写 result 的，异常可能发生在它写了一半之后，
        # 因此这里把所有承载评分结论的字段显式复位，保证 error 的返回值是干净的全 0；
        # dim1_pass 复位为空串而不是 False——异常意味着维度一的结论作废，不是判定为不通过。
        result['status'] = 'error'
        result['error'] = f'{type(exc).__name__}: {exc}'
        result['dim1_pass'] = ''
        result['dim1_reason'] = ''
        result['dim2_items'] = []
        result['total_score'] = 0
        result['max_score'] = 0
        return result


def _evaluate_impl(doc, result):
    body = doc.element.body
    sm = _style_maps(doc)
    T = _map_labeled_tables(body)

    # 正文（body）直接子级的段落 / 表格顺序
    seq = []
    cap_paras = []
    for ch in body.iterchildren():
        if ch.tag == qn('w:p'):
            txt = _txt_of(ch).strip()
            seq.append(('P', txt))
            for k in ORDER:
                if _nz(CAPS[k]) == _nz(txt):
                    cap_paras.append((k, ch))
                    break
        elif ch.tag == qn('w:tbl'):
            seq.append(('T', ''))
    paras_norm = {_nz(t) for kind, t in seq if kind == 'P'}

    sec = doc.sections[0]
    try:
        text_width_cm = sec.page_width.cm - sec.left_margin.cm - sec.right_margin.cm
    except Exception:
        text_width_cm = 0.0

    # --- 加分项：逐表判定 -------------------------------------------------
    hits = {}
    for key, lab, _md, _rule in _ITEMS:
        if lab is None:
            continue
        t = T.get(lab)
        if t is None:
            hits[(key, lab)] = (False, '未在文档中定位到%s（表题缺失或表题后无独立表格）' % LAB(lab))
            continue
        if key == 'A1':
            hits[(key, lab)] = _chk_a1_outer(t)
        elif key in ('A2', 'A3'):
            hits[(key, lab)] = _chk_group_line(lab, t)
        elif key == 'A4':
            hits[(key, lab)] = _chk_a4_shape(lab, t)
        elif key in ('A5', 'A6'):
            hits[(key, lab)] = _chk_header_merge(lab, t)
        elif key == 'A7':
            hits[(key, lab)] = _chk_a7_dim_merge(lab, t)
        elif key == 'A8':
            hits[(key, lab)] = _chk_a8_stat_merge(lab, t)
        elif key == 'A9':
            hits[(key, lab)] = _chk_a9_width(t, text_width_cm)
        elif key == 'A10':
            hits[(key, lab)] = _chk_a10_height(lab, t)
        elif key == 'A11':
            hits[(key, lab)] = _chk_a11_font_align(t, sm)

    # --- 减分项 -----------------------------------------------------------
    hits['D1'] = _chk_d1_page(doc)
    hits['D2'] = _chk_d2_caps(paras_norm)
    hits['D3'] = _chk_d3_cap_align(cap_paras, sm)
    hits['D4'] = _chk_d4_cap_font(cap_paras, sm)
    hits['D5'] = _chk_d5_follow(seq, T)
    hits['D6'] = _chk_d6_size(T, sm)
    hits['D7'] = _chk_presence(T, HDR_REQ, '表头文字缺失或被改写：')
    hits['D8'] = _chk_presence(T, DIMS, '维度名称缺失或被改写：')
    hits['D9'] = _chk_presence(T, GRP, '组别/测试时间名称缺失或被改写：')
    for k in ORDER:
        hits['V' + k] = _chk_values(k, T)
    hits['D19'], hits['D20'], hits['D21'] = _chk_d19_d20_d21(T)
    hits['D22'], hits['D23'] = _chk_d22_d23(T)
    hits['D24'] = _chk_d24_hf(doc)
    hits['D25'] = _chk_d25_extra(seq)
    hits['D26'] = _chk_d26_break(body)

    # ========== 生成结构化 dim2_items ==========
    # 按 _ITEMS 固定顺序（80 条逐表拆分的加分项 + 26 条减分项）逐项取判定结果。
    dim2_items, total_score = _build_dim2_items(hits)

    result['dim2_items'] = dim2_items
    result['total_score'] = total_score
    result['max_score'] = MAX_SCORE_TOTAL
    return result


if __name__ == '__main__':
    # 本地调试入口：命令行第 1 个参数可覆盖目录路径；未指定时使用脚本所在目录。
    # 直接写 stdout 的二进制缓冲区并显式指定 utf-8，避免受控制台默认编码影响。
    _dir = sys.argv[1] if len(sys.argv) > 1 else os.path.dirname(os.path.abspath(__file__))
    _output = json.dumps(evaluate(_dir), ensure_ascii=False, indent=2) + '\n'
    sys.stdout.buffer.write(_output.encode('utf-8', errors='replace'))
