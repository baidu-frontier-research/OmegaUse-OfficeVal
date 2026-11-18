#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# pyright: reportUnknownMemberType=false, reportUnknownVariableType=false, reportUnknownArgumentType=false, reportUnknownParameterType=false, reportMissingParameterType=false, reportPrivateUsage=false, reportUnknownLambdaType=false, reportGeneralTypeIssues=false

import os
import re
import json
import math
import sys
import traceback
from collections import Counter, defaultdict
# PyMuPDF 已从本项目移除（AGPL 与本项目许可证不兼容），运行环境不安装 fitz。
# 改用同目录的 pymupdf_compat：用 pdfplumber + pypdfium2 实现所需的 PyMuPDF 子集。
# verifier 由 core/worker.py 按文件路径加载（其所在目录已被插入 sys.path），
# 下面两行在作为包导入时补齐目录，两种加载方式都成立。
_VERIFIER_DIR = os.path.dirname(os.path.abspath(__file__))
if _VERIFIER_DIR not in sys.path:
    sys.path.insert(0, _VERIFIER_DIR)
try:
    import pymupdf_compat as fitz
except ImportError:
    fitz = None

# 说明：按接口统一约定，本模块不修改全局标准输出，也不 print 主结果。
# 主结果统一由 evaluate() -> dict 返回；评估目标固定为"本脚本所在的那个目录"，
# 交付物只接受该目录下可正常打开的 .pdf（交了 .docx / .doc 等其它格式时，仍收作
# 被评对象，但在维度一直接判不通过），不进入子目录、不检索上级目录、
# 也不接受命令行或入参指定的其它路径。
# 本文件仅在 __main__ 下作为本地调试入口，才把 dict 序列化为 JSON 打印。
#
# 判定口径：细则的每一条都落在"版面"上（分栏、配对、分页、页码、字号、边距），
# 因此统一按交付 PDF 的 span/line/drawing 级几何量测，不做任何格式转换或渲染。

SCRIPT_ID = "034"
# 评估目标锁定为本脚本所在目录：评测时把本脚本与被评交付物放在同一个目录下即可。
_SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
# 维度二加分项 max_delta 之和（= 细则"加分项满分"）；用于维度一未通过 /
# 脚本异常时兜底。减分项只在命中时产生负 delta，不计入 max_score。
MAX_SCORE_TOTAL = 43

# 维度二加分项按"结果追加顺序"的 max_delta 列表；末端把 plus_results 里的
# "+N：..." 字符串反解为结构化 dim2_items 时用。顺序须与 _evaluate_impl 中
# plus_results.append(...) 的顺序严格一致；共 19 项，累加 = MAX_SCORE_TOTAL。
_MAX_DELTAS_ORDER = [
    5,  # 01 十六条条款正文全部两栏表格化，左栏中文右栏英文
    3,  # 02 每个条款表格首行为"中文 / English"表头行
    3,  # 03 子条款中英成对同行，编号双栏保留且一致
    1,  # 04 条款表格可见完整边框网格
    1,  # 05 左右栏宽大致均等、栏内左对齐、无越栏压线
    3,  # 06 条款题序标题置于表格上方，中前英后且与正文有区分
    1,  # 07 文档标题与条款题序标题中英两行字重一致（均加粗）
    3,  # 08 文档标题中文行在前、英文行紧随其后
    3,  # 09 签署前言与双方信息逐行中前英后
    3,  # 10 签署页逐行中前英后且未套用两栏表格
    1,  # 11 全文中英顺序统一为中文在前
    5,  # 12 无跨页拆分的条款【前提：条款已两栏表格化】
    3,  # 13 多条条款同页排布【前提：同上】
    1,  # 14 条款之间保留可辨空白间距【前提：同上】
    1,  # 15 填空下划线未被折行拆断【前提：同上】
    1,  # 16 子条款编号与正文同行行首【前提：同上】
    3,  # 17 全部页面阿拉伯数字连续页码
    1,  # 18 页码水平居中
    1,  # 19 页码垂直位于页脚区域中部
]
assert sum(_MAX_DELTAS_ORDER) == MAX_SCORE_TOTAL

# 维度二减分项按"结果追加顺序"的 max_delta（负值）列表；顺序须与
# _evaluate_impl 中 minus_results.append(...) 的顺序严格一致；共 18 项。
_MIN_DELTAS_ORDER = [
    -1,  # 20 第16.5条原文未逐字保留，或其中英排布被单独改为英文在前
    -3,  # 21 条款正文、标题、签署栏文字被删除或改写
    -3,  # 22 条款/子条款/签署前言/签署页内容缺失
    -1,  # 23 条款编号、附件编号或交叉引用数值被改动
    -1,  # 24 擅自新增页眉/水印/目录/封面/占位文字
    -1,  # 25 正文冒号被替换为全角"："
    -1,  # 26 句内逗号被替换为全角"，"
    -1,  # 27 文档标题字号偏离约15.5磅
    -1,  # 28 条款题序标题字号被放大至14磅以上
    -1,  # 29 签署页字号被单独放大为标题级
    -1,  # 30 页面左右边距被改为宽边距
    -1,  # 31 纸张非A4纵向
    -1,  # 32 页码格式非纯阿拉伯数字或页脚被加料
    -1,  # 33 非条款内容被误套用两栏表格
    -1,  # 34 中文被设为非中文字体导致字形异常
    -1,  # 35 英文单词被硬拆断或空格粘连
    -1,  # 36 填空下划线被压缩或长短不一
    -1,  # 37 表格行高/内边距压缩或中英两栏起始行错位
]

# ========== 交付内容的基准文本（仅作"未被改动"的比对基准，不含任何版式信息） ==========
# 十六条条款题序标题：中文题序 / 英文题序（用于条款定位与内容缺失判定）
_CLAUSE_TITLES = [
    (1, "第一条", "合同标的及设备清单", "Clause 1", "Contract Subject and Equipment List"),
    (2, "第二条", "服务期限与交付", "Clause 2", "Service Period and Delivery"),
    (3, "第三条", "费用及付款安排", "Clause 3", "Fees and Payment Arrangement"),
    (4, "第四条", "所有权及使用限制", "Clause 4", "Ownership and Use Restrictions"),
    (5, "第五条", "安装、网络环境及访问", "Clause 5", "Installation, Network Environment and Access"),
    (6, "第六条", "维护及服务水平", "Clause 6", "Maintenance and Service Level"),
    (7, "第七条", "数据处理与保密", "Clause 7", "Data Handling and Confidentiality"),
    (8, "第八条", "安全措施", "Clause 8", "Security Measures"),
    (9, "第九条", "变更请求及额外工作", "Clause 9", "Change Requests and Additional Work"),
    (10, "第十条", "遗失、损坏及保险", "Clause 10", "Loss, Damage and Insurance"),
    (11, "第十一条", "审核、记录及盘点", "Clause 11", "Audit, Records and Inventory Check"),
    (12, "第十二条", "暂停与终止", "Clause 12", "Suspension and Termination"),
    (13, "第十三条", "返还及最终结算", "Clause 13", "Return and Final Settlement"),
    (14, "第十四条", "责任限制", "Clause 14", "Liability Limitation"),
    (15, "第十五条", "争议解决及适用法律", "Clause 15", "Dispute Resolution and Governing Law"),
    (16, "第十六条", "其他", "Clause 16", "Miscellaneous"),
]

# 每条条款的子条款个数（1.1 起连续编号）；合计 63 个子条款
_SUBCLAUSE_COUNTS = {
    1: 3, 2: 4, 3: 4, 4: 3, 5: 4, 6: 5, 7: 5, 8: 4,
    9: 4, 10: 4, 11: 4, 12: 4, 13: 4, 14: 3, 15: 3, 16: 5,
}
_CANON_SUBCLAUSES = [
    "%d.%d" % (c, i)
    for c in range(1, 17)
    for i in range(1, _SUBCLAUSE_COUNTS[c] + 1)
]

# 第16.5条的中英表述（细则 -1 项单独占用的判定点，不参与"文字被改写"的比对）
_C1605_CN = "英文文本列于对应中文文本之上. 如解释存在差异, 应以最符合商业目的的文本为准."
_C1605_EN = ("The English text appears above the corresponding Chinese text. "
             "If interpretation differs, the language most consistent with the "
             "commercial purpose shall prevail.")

# 签署前言与双方信息：中文行关键字 -> 紧随其后的英文行关键字
_PREFACE_PAIRS = [
    ("本合同由以下双方于", "This Contract is entered into on"),
    ("甲方/服务提供方", "Party A / Service Provider"),
    ("授权代表", "Authorized Representative"),
    ("行政联系", "Administrative Contact"),
    ("乙方/客户方", "Party B / Client"),
    ("甲方与乙方以下合称为", "hereinafter collectively referred to"),
]
# 签署页：中文行关键字 -> 紧随其后的英文行关键字
_SIGN_PAIRS = [
    ("签署页", "Signature Page"),
    ("甲方/服务提供方", "Party A / Service Provider"),
    ("授权代表签字", "Authorized Representative Signature"),
    ("姓名", "Name"),
    ("日期", "Date"),
    ("乙方/客户方", "Party B / Client"),
]
_PREFACE_LABELS = ["本合同由以下双方于", "甲方/服务提供方", "统一识别号", "授权代表",
                   "地址", "行政联系", "乙方/客户方", "甲方与乙方以下合称为"]
_SIGN_LABELS = ["签署页", "SignaturePage", "授权代表签字", "姓名", "日期", "Name", "Date"]
_TITLE_CN = "智能办公设备租赁及支持服务合同"
_TITLE_EN = "Managed Workspace Equipment Leasing and Support Contract"
# 子条款正文基准（编号前缀不参与比对，因编号与正文之间的空格属版式差异）；
# 比对前统一做归一化：去空白、去下划线填空、全角标点折半、引号折半。
# 第16.5条不在此表内——它由减分项 20 单独判定，避免同一处内容重复扣分。
_SUB_TEXT = {
    "1.1": ("甲方向乙方出租一套智能办公设备组合, 仅供乙方内部办公使用.",
            "Party A shall lease to Party B a managed workspace equipment package for internal office use only."),
    "1.2": ("设备组合包括____台轻量终端、____台网络打印机、____台安全路由器及____套机柜组件.",
            "The equipment package includes ____ thin-client terminals, ____ network printers, ____ secure routers, and ____ cabinet set."),
    "1.3": ("资产编码、序列号及状态说明载于附件A, 并构成本合同的组成部分.",
            "The asset codes, serial numbers, and condition notes are recorded in Schedule A and shall form part of this Contract."),
    "2.1": ("初始服务期限为____个月, 自____起至____止.",
            "The initial service period is ____ months, commencing on ____ and ending on ____."),
    "2.2": ("甲方应于____或之前将设备交付乙方, 并在移交前完成基础配置.",
            "Party A shall deliver the equipment to Party B on or before ____ and complete basic configuration before handover."),
    "2.3": ("乙方应在交付后____个工作日内验收设备组合, 如无重大缺陷则签署移交记录.",
            "Party B shall inspect the package within ____ business days after delivery and sign the handover record if no material defect is found."),
    "2.4": ("乙方在验收期限内未提出书面意见的, 视为设备已通过验收.",
            "If Party B does not provide written comments within the inspection period, the equipment shall be deemed accepted."),
    "3.1": ("月度服务费为新加坡元____元, 包含设备租赁、远程监测、预防性维护及标准服务台支持.",
            "The monthly service fee is S$____, including equipment rental, remote monitoring, preventive maintenance, and standard helpdesk support."),
    "3.2": ("乙方应在签署本合同后____个工作日内支付可退还保证金新加坡元____元.",
            "Party B shall pay a refundable security deposit of S$____ within ____ business days after signing this Contract."),
    "3.3": ("月度服务费应于每月第____个自然日前支付至甲方指定账户.",
            "The monthly service fee shall be paid before the ____ calendar day of each month to the account designated by Party A."),
    "3.4": ("因乙方付款渠道产生的银行费用由乙方承担.",
            "Any bank charge generated by Party B's payment channel shall be borne by Party B."),
    "4.1": ("除非双方另行签署书面转让文件, 所有租赁设备的所有权均归甲方所有.",
            "Ownership of all leased equipment remains with Party A unless the Parties sign a separate written transfer document."),
    "4.2": ("乙方不得转让、质押、转租、拆解、更换标签或在租赁设备上设立任何权利负担.",
            "Party B shall not sell, pledge, sublease, dismantle, replace labels, or create any security interest over the leased equipment."),
    "4.3": ("乙方仅可在约定地点使用设备, 未经甲方书面同意不得迁移至其他地点.",
            "Party B shall use the equipment only at the agreed site and shall not move it to another location without Party A's written consent."),
    "5.1": ("乙方应提供安装所需的稳定电力、网络布线、机柜空间及合理现场访问条件.",
            "Party B shall provide stable power supply, network cabling, rack space, and reasonable site access required for installation."),
    "5.2": ("甲方应按照上线清单配置管理员账户、设备策略及监测代理程序.",
            "Party A shall configure administrator accounts, device policies, and monitoring agents in accordance with the onboarding checklist."),
    "5.3": ("甲方人员在发出合理通知后方可进入现场进行安装或维护.",
            "Party A's personnel may access the site for installation or maintenance only after giving reasonable notice to Party B."),
    "5.4": ("当设备故障对业务连续性产生重大影响时, 甲方可申请紧急访问.",
            "Emergency access may be requested where equipment failure materially affects business continuity."),
    "6.1": ("甲方应在新加坡时间工作日____至____提供远程监测服务.",
            "Party A shall provide remote monitoring on business days from ____ to ____ Singapore time."),
    "6.2": ("对于工作日____前报告的优先级事件, 甲方应在____个工作小时内开始诊断.",
            "For priority incidents reported before ____ on a business day, Party A shall begin diagnosis within ____ business hours."),
    "6.3": ("如终端无法远程恢复, 且库存可用, 甲方应在____个工作日内安排替换设备.",
            "If a terminal cannot be restored remotely, Party A shall arrange a replacement unit within ____ business days where stock is available."),
    "6.4": ("耗材、用户造成的损坏、楼宇网络故障及第三方应用故障不属于标准服务范围.",
            "Consumables, user-caused damage, building network faults, and third-party application faults are outside the standard service scope."),
    "6.5": ("工作时间外的可选现场支持按每名工程师每小时新加坡元____元计费, 并须事先获得书面批准.",
            "Optional on-site support outside business hours shall be charged at S$____ per engineer hour, subject to prior written approval."),
    "7.1": ("双方应对从对方获知的非公开商业、技术、财务、运营及人员信息承担保密义务.",
            "Each Party shall keep confidential all non-public business, technical, financial, operational, and personnel information received from the other Party."),
    "7.2": ("除为提供支持所必需且经乙方书面授权外, 甲方不得查看乙方用户文件.",
            "Party A shall not review Party B's user files unless access is necessary for support and Party B has given written authorization."),
    "7.3": ("诊断日志可用于事件处理、容量复核、安全审计及服务改进.",
            "Diagnostic logs may be collected for incident resolution, capacity review, security audit, and service improvement."),
    "7.4": ("除法律要求或未结争议另有需要外, 诊断日志保存期限不超过____天.",
            "Diagnostic logs shall be retained for no longer than ____ days unless a longer period is required by law or an ongoing dispute."),
    "7.5": ("保密义务在本合同终止后继续有效____年.",
            "Confidentiality obligations shall continue for ____ years after termination of this Contract."),
    "8.1": ("甲方应为受管终端启用设备加密、终端防护及管理员密码轮换.",
            "Party A shall enable device encryption, endpoint protection, and administrative password rotation for managed terminals."),
    "8.2": ("乙方应确保其用户不绕过安全控制、不安装未批准软件且不共享管理员凭证.",
            "Party B shall ensure that its users do not bypass security controls, install unapproved software, or share administrative credentials."),
    "8.3": ("一方初步确认疑似安全事件后, 应在____小时内通知另一方.",
            "A suspected security incident shall be notified to the other Party within ____ hours after initial confirmation."),
    "8.4": ("双方应就控制影响、保全证据及修复措施进行合理配合.",
            "Each Party shall reasonably cooperate in containment, evidence preservation, and remediation."),
    "9.1": ("对数量、地点、配置、访问策略或服务范围的变更应以书面方式提出.",
            "A change to quantity, location, configuration, access policy, or service scope shall be submitted in writing."),
    "9.2": ("甲方应提供影响评估, 包括时间安排、费用、停机时间及运营风险.",
            "Party A shall provide an impact assessment covering timeline, cost, downtime, and operational risk."),
    "9.3": ("未经双方授权代表签署或以其他书面方式确认的变更, 对双方均无约束力.",
            "No change shall be binding unless signed or otherwise confirmed in writing by authorized representatives of both Parties."),
    "9.4": ("可在____名工程师分钟内完成的轻微配置调整, 甲方可酌情纳入标准支持范围.",
            "Minor configuration adjustments completed within ____ engineer minutes may be included in standard support at Party A's discretion."),
    "10.1": ("自交付至返还期间, 乙方应对设备尽到合理保管义务.",
             "Party B shall exercise reasonable care over the equipment from delivery until return."),
    "10.2": ("因乙方员工、承包商、访客或现场条件造成的遗失或物理损坏, 由乙方赔偿.",
             "Loss or physical damage caused by Party B's staff, contractors, visitors, or site conditions shall be compensated by Party B."),
    "10.3": ("每项资产的重置价值应按照附件B的折旧表确定.",
             "The replacement value for each asset shall be based on the depreciation schedule in Schedule B."),
    "10.4": ("乙方可自费投保办公设备保险, 在可行情况下将甲方列为权益相关方.",
             "Party B may maintain office equipment insurance at its own cost and shall list Party A as an interested party where practicable."),
    "11.1": ("甲方应维护资产登记册, 记录交付日期、资产编码、配置档案及服务历史.",
             "Party A shall maintain an asset register covering delivery date, asset code, configuration profile, and service history."),
    "11.2": ("乙方应在收到____个工作日通知后, 每____个月配合一次远程或现场盘点.",
             "Party B shall cooperate with a remote or on-site inventory check once every ____ months after receiving ____ business days notice."),
    "11.3": ("任何设备缺失、标签变更、未经授权迁移或未报告损坏均应记入盘点说明.",
             "Any missing equipment, altered label, unauthorized relocation, or unreported damage shall be recorded in the inventory note."),
    "11.4": ("根据本条形成的记录可用于结算、保险索赔及争议解决.",
             "Records created under this Clause may be used for settlement, insurance claim, and dispute resolution."),
    "12.1": ("乙方在书面通知后超过____个自然日仍未支付无争议账单的, 甲方可暂停非关键支持服务.",
             "Party A may suspend non-critical support if Party B fails to pay an undisputed invoice for more than ____ calendar days after written notice."),
    "12.2": ("一方发生重大违约且未能在____个工作日内补救的, 另一方有权终止本合同.",
             "Either Party may terminate this Contract if the other Party commits a material breach and fails to cure it within ____ business days."),
    "12.3": ("合同终止不影响已发生的付款义务、保密义务或终止前产生的权利.",
             "Termination shall not affect accrued payment obligations, confidentiality obligations, or rights arising before termination."),
    "12.4": ("乙方无故提前终止的, 应支付初始服务期剩余服务费的____%.",
             "Early termination by Party B for convenience requires payment of ____% of the remaining service fees for the initial service period."),
    "13.1": ("合同届满或终止时, 乙方应返还全部设备且设备处于良好工作状态, 合理损耗除外.",
             "Upon expiry or termination, Party B shall return all equipment in good working condition, fair wear and tear excepted."),
    "13.2": ("甲方应在返还后____个工作日内检查设备并出具返还验收说明.",
             "Party A shall inspect returned equipment within ____ business days and issue a return inspection note."),
    "13.3": ("保证金可用于抵扣未付账单、维修费用、缺失配件及双方约定的终止费用.",
             "The security deposit may be applied against unpaid invoices, repair costs, missing accessories, and agreed termination charges."),
    "13.4": ("最终结算后仍有保证金余额的, 甲方应在____个自然日内退还乙方.",
             "Any remaining deposit balance shall be refunded to Party B within ____ calendar days after final settlement."),
    "14.1": ("任一方均不对间接损失、利润损失、商誉损失或后果性损害承担责任.",
             "Neither Party shall be liable for indirect loss, loss of profit, loss of goodwill, or consequential damages."),
    "14.2": ("除第14.3条另有规定外, 任一方的累计责任不超过此前____个月已支付服务费总额.",
             "Subject to Clause 14.3, each Party's aggregate liability shall not exceed the total service fees paid in the preceding ____ months."),
    "14.3": ("责任限制不适用于欺诈、故意不当行为、违反保密义务、数据滥用或未付费用.",
             "The limitation shall not apply to fraud, wilful misconduct, confidentiality breach, data misuse, or unpaid fees."),
    "15.1": ("除强制性法律另有规定外, 本合同适用新加坡法律.",
             "This Contract shall be governed by the laws of Singapore unless mandatory law provides otherwise."),
    "15.2": ("争议应首先提交双方运营负责人进行善意协商.",
             "A dispute shall first be escalated to the operational managers of both Parties for good faith negotiation."),
    "15.3": ("如____个自然日内未达成解决方案, 任一方可将争议提交____有管辖权的法院.",
             "If no settlement is reached within ____ calendar days, either Party may submit the dispute to the competent courts of ____."),
    "16.1": ("本合同连同附件A及附件B构成双方就本合同标的达成的完整协议.",
             "This Contract, together with Schedule A and Schedule B, constitutes the entire agreement between the Parties regarding the subject matter."),
    "16.2": ("任何修改均应以书面形式作出, 并由双方签署或以其他书面方式确认.",
             "Any amendment shall be made in writing and signed or otherwise confirmed by both Parties."),
    "16.3": ("通知可通过专人递送、快递、挂号邮件或附件A记载的行政联系渠道送达.",
             "Notices may be delivered by hand, courier, registered mail, or the administrative contacts recorded in Schedule A."),
    "16.4": ("本合同一式两份, 双方各执一份, 具有同等法律效力.",
             "This Contract is executed in two counterparts, and each Party shall hold one counterpart with equal legal effect."),
}

_CJK_RE = re.compile(r"[一-鿿]")
_LATIN_RE = re.compile(r"[A-Za-z]")
_SUBNO_RE = re.compile(r"^(\d{1,2})\s*\.\s*(\d{1,2})(?!\d)")
_CLAUSE_CN_RE = re.compile(r"^第\s*(一|二|三|四|五|六|七|八|九|十|十一|十二|十三|十四|十五|十六)\s*条")
_CLAUSE_EN_RE = re.compile(r"^Clause\s*(\d{1,2})\s*[:：]")
_CN_NUMERALS = {"一": 1, "二": 2, "三": 3, "四": 4, "五": 5, "六": 6, "七": 7, "八": 8,
                "九": 9, "十": 10, "十一": 11, "十二": 12, "十三": 13, "十四": 14,
                "十五": 15, "十六": 16}
# 纯拉丁字形字体族：若中文字符落在这些字体上，中文会缺字或显示为方框
_LATIN_ONLY_FONTS = ("helvetica", "arial", "times", "courier", "calibri", "verdana",
                     "tahoma", "georgia", "cambria", "garamond", "consolas")
_PUNCT_MAP = {
    "，": ",", "。": ".", "：": ":", "；": ";", "！": "!", "？": "?",
    "（": "(", "）": ")", "、": "、",
    "“": '"', "”": '"', "‘": "'", "’": "'", "＄": "$", "％": "%",
    "‘": "'", "’": "'", "“": '"', "”": '"',
    "％": "%", "＄": "$", " ": "",
}


def _norm(text):
    """内容比对用归一化：全角标点折半、去引号差异、去所有空白、去下划线填空。
    这样"改版式"不会被误判成"改文字"，而真正的增删改写仍会暴露。"""
    if not text:
        return ""
    out = []
    for ch in text:
        out.append(_PUNCT_MAP.get(ch, ch))
    s = "".join(out)
    s = re.sub(r"_+", "", s)
    s = re.sub(r"\s+", "", s)
    return s


def _sentence_lost(want, stream):
    """句子是否被删改：容忍句末标点（填空线后的收尾句号常在重排时丢失），
    其余文字必须原样出现在同语言文本流中。"""
    key = _norm(want).rstrip(".,;:!?")
    return bool(key) and key not in stream


def _is_bold(span):
    """加粗判定：优先取 PDF 字体描述符的 bold 标志位，再退回字体名关键字。"""
    if span.get("flags", 0) & (1 << 4):
        return True
    name = (span.get("font") or "").lower()
    return bool(re.search(r"bold|black|heavy|semibold|demibold", name))


def _lang_of(text):
    """行语种：含汉字记 'C'，只含拉丁字母记 'E'，其余（纯填空、纯标点）记 None。"""
    if _CJK_RE.search(text):
        return "C"
    if _LATIN_RE.search(text):
        return "E"
    return None


# 交付物只接受 .pdf：细则要求交的是可正常打开的 PDF，其余版式/源文件格式
# （.docx / .doc / .rtf / .pptx 等）一律在维度一判不通过。
_ACCEPTED_EXT = ".pdf"
# 目录里没有 .pdf 时，这些扩展名仍收作被评对象，交给维度一判不通过（完成度总分记 0），
# 而不是当成「没有交付物」直接跳过评测。
_WRONG_FMT_EXTS = (".docx", ".doc", ".docm", ".dotx", ".dot", ".rtf", ".odt", ".wps",
                   ".pages", ".pptx", ".ppt", ".pptm", ".odp", ".key")


def _fmt_reject(file_path):
    """维度1的格式门禁：交付物必须是可正常打开的 .pdf。

    不通过返回原因文本，通过返回 None。判定与阅读器的「能不能双击打开」对齐：
    扩展名须为 .pdf，文件头须含 %PDF- 标识；PDF 解析后端可用时再校验能否打开、
    是否至少有 1 页、是否需要口令。
    """
    ext = os.path.splitext(file_path)[1].lower()
    if ext != _ACCEPTED_EXT:
        return "交付物为 %s 格式，只接受 .pdf 文档" % (ext or "无扩展名")
    try:
        with open(file_path, "rb") as fh:
            head = fh.read(1024)
    except OSError as exc:
        return "交付物无法读取：%s: %s" % (type(exc).__name__, exc)
    if b"%PDF-" not in head:
        return "交付物不是有效的 PDF（文件头缺少 %PDF- 标识）"
    if fitz is None:
        return None
    try:
        doc = fitz.open(file_path)
    except Exception as exc:
        return "交付物无法作为 PDF 打开：%s: %s" % (type(exc).__name__, exc)
    try:
        if doc.needs_pass:
            return "交付物 PDF 已加密，需要口令才能打开"
        if doc.page_count < 1:
            return "交付物 PDF 内没有任何页面"
    finally:
        doc.close()
    return None


def _locate_document(dir_path):
    """仅在给定目录本身定位被评估的交付文档：优先 .pdf；目录内没有 .pdf 时，
    退一步收下 .docx / .doc 等其他格式，交由维度一的格式门禁判不通过（而不是
    当成「没有交付物」跳过评测）。忽略 ~$ 打头的 Office 临时锁文件。
    不进入任何子目录，也不检索上级目录。"""
    try:
        names = sorted(os.listdir(dir_path))
    except OSError:
        names = []
    files = [f for f in names
             if not f.startswith("~$")
             and os.path.isfile(os.path.join(dir_path, f))]
    cands = [f for f in files if f.lower().endswith(_ACCEPTED_EXT)]
    if not cands:
        cands = [f for f in files if f.lower().endswith(_WRONG_FMT_EXTS)]
    if not cands:
        raise FileNotFoundError("目录中未找到 .pdf 文件：%s" % dir_path)
    return os.path.join(dir_path, cands[0])


def _raw_lines(page):
    """取页面上的原始文本行（span 级信息保留），过滤空行。"""
    out = []
    for block in page.get_text("dict").get("blocks", []):
        if block.get("type") != 0:
            continue
        for ln in block.get("lines", []):
            spans = [s for s in ln.get("spans", []) if s.get("text", "").strip()]
            if not spans:
                continue
            # PDF 文本流中的 span 顺序不保证等于视觉阅读顺序；例如“风险”可能
            # 以“险”“风”的对象顺序存储。按横坐标还原，避免内容比对假报改写。
            spans.sort(key=lambda s: (s.get("bbox") or (0, 0, 0, 0))[0])
            text = "".join(s["text"] for s in spans)
            if not text.strip():
                continue
            x0, y0, x1, y1 = ln["bbox"]
            out.append({
                "x0": x0, "y0": y0, "x1": x1, "y1": y1,
                "text": text,
                "size": max(s["size"] for s in spans),
                "bold": all(_is_bold(s) for s in spans),
                "any_bold": any(_is_bold(s) for s in spans),
                "fonts": [s.get("font") or "" for s in spans],
                "spans": spans,
            })
    out.sort(key=lambda r: (round(r["y0"], 1), r["x0"]))
    return out


def _merge_by_baseline(rows, ytol=4.0):
    """把同一视觉行被切成多段的文本行合并回一行（PDF 里填空、换字体都会断段）。
    合并只在同一个"区域"内进行（页面正文区 / 表格左栏 / 表格右栏 各自独立），
    因此不会把左右两栏的文字拼到一起。合并后按 x 从左到右还原阅读顺序。"""
    merged = []
    for row in sorted(rows, key=lambda r: (round(r["y0"], 1), r["x0"])):
        placed = False
        for m in merged:
            if abs(m["y0"] - row["y0"]) <= ytol or abs(m["yc"] - (row["y0"] + row["y1"]) / 2) <= ytol:
                m["parts"].append(row)
                m["y0"] = min(m["y0"], row["y0"])
                m["y1"] = max(m["y1"], row["y1"])
                m["yc"] = (m["y0"] + m["y1"]) / 2
                placed = True
                break
        if not placed:
            merged.append({
                "y0": row["y0"], "y1": row["y1"],
                "yc": (row["y0"] + row["y1"]) / 2,
                "parts": [row],
            })
    out = []
    for m in merged:
        parts = sorted(m["parts"], key=lambda r: r["x0"])
        # 一个 raw line 自身也可能包含横坐标交错的 span，不能只排序 line。
        # 将全部 span 摊平后按视觉横坐标重排，才能还原“风”“险”等被 PDF
        # 对象顺序打乱的相邻文字。
        spans = sorted(
            [s for p in parts for s in p["spans"]],
            key=lambda s: (s.get("bbox") or (0, 0, 0, 0))[0],
        )
        text = " ".join(s["text"] for s in spans)
        out.append({
            "x0": min(p["x0"] for p in parts),
            "x1": max(p["x1"] for p in parts),
            "y0": m["y0"], "y1": m["y1"],
            "text": text,
            "size": max(p["size"] for p in parts),
            "bold": all(p["bold"] for p in parts),
            "any_bold": any(p["any_bold"] for p in parts),
            "fonts": [f for p in parts for f in p["fonts"]],
            "spans": spans,
        })
    out.sort(key=lambda r: r["y0"])
    return out


def _page_segments(page):
    """取页面上的水平/垂直线段（表格边框线与填空下划线都在其中）。
    水平段记 (x_left, x_right, y)，垂直段记 (y_top, y_bottom, x)。"""
    segs_h, segs_v = [], []
    for d in page.get_drawings():
        for item in d.get("items", []):
            if item[0] == "l":
                p1, p2 = item[1], item[2]
                if abs(p1.y - p2.y) < 0.8 and abs(p1.x - p2.x) >= 0.8:
                    segs_h.append((min(p1.x, p2.x), max(p1.x, p2.x), (p1.y + p2.y) / 2))
                elif abs(p1.x - p2.x) < 0.8 and abs(p1.y - p2.y) >= 0.8:
                    segs_v.append((min(p1.y, p2.y), max(p1.y, p2.y), (p1.x + p2.x) / 2))
            elif item[0] == "re":
                r = item[1]
                if r.height < 1.5 <= r.width:
                    segs_h.append((r.x0, r.x1, (r.y0 + r.y1) / 2))
                elif r.width < 1.5 <= r.height:
                    segs_v.append((r.y0, r.y1, (r.x0 + r.x1) / 2))
    return segs_h, segs_v


def _build_table(page, bbox, raw_rows, segs_h, segs_v):
    """按表格外框把页面文本行分到左右两栏，并量出分栏线、行分隔线等几何量。
    分栏线取"跨表格高度过半、且落在表格中段"的垂直线中最靠近外框中线的一条；
    找不到画线时退回外框中线。"""
    x0, y0, x1, y1 = bbox
    w, h = x1 - x0, y1 - y0
    inner = [r for r in raw_rows
             if r["x0"] >= x0 - 3 and r["x1"] <= x1 + 3
             and r["y0"] >= y0 - 3 and r["y1"] <= y1 + 8]
    cand = [s for s in segs_v
            if min(s[1], y1) - max(s[0], y0) > h * 0.5
            and x0 + w * 0.2 <= s[2] <= x1 - w * 0.2]
    center = (x0 + x1) / 2
    divider = min((s[2] for s in cand), key=lambda v: abs(v - center)) if cand else center

    def span_row(spans):
        spans = sorted(spans, key=lambda s: (s.get("bbox") or (0, 0, 0, 0))[0])
        boxes = [s["bbox"] for s in spans]
        return {
            "x0": min(b[0] for b in boxes), "y0": min(b[1] for b in boxes),
            "x1": max(b[2] for b in boxes), "y1": max(b[3] for b in boxes),
            "text": "".join(s["text"] for s in spans),
            "size": max(s["size"] for s in spans),
            "bold": all(_is_bold(s) for s in spans),
            "any_bold": any(_is_bold(s) for s in spans),
            "fonts": [s.get("font") or "" for s in spans],
            "spans": spans,
        }

    # 兼容后端有时会把同一基线上的左右栏 span 合成一个 raw line。必须先按
    # 分栏线拆回两侧，再做行合并；用整行 bbox 判越栏会把正常双栏误报为压线。
    rows_l, rows_r, straddle = [], [], []
    for row in inner:
        left_spans, right_spans = [], []
        for span in row.get("spans", []):
            sx0, _sy0, sx1, _sy1 = span["bbox"]
            if sx0 < divider - 1.5 and sx1 > divider + 1.5:
                straddle.append(span)
            target = left_spans if (sx0 + sx1) / 2 < divider else right_spans
            target.append(span)
        if left_spans:
            rows_l.append(span_row(left_spans))
        if right_spans:
            rows_r.append(span_row(right_spans))
    # 行分隔线：跨表宽 85% 以上的水平线（外框上下沿也在其中）
    grid_y = sorted({round(s[2], 1) for s in segs_h
                     if s[1] - s[0] > w * 0.85 and y0 - 2 <= s[2] <= y1 + 2})
    row_seps = [y for y in grid_y if y0 + 2 < y < y1 - 2]
    has_top = any(abs(y - y0) <= 2.5 for y in grid_y)
    has_bottom = any(abs(y - y1) <= 2.5 for y in grid_y)
    tall_v = [s[2] for s in segs_v if min(s[1], y1) - max(s[0], y0) > h * 0.5]
    has_left = any(abs(v - x0) <= 2.5 for v in tall_v)
    has_right = any(abs(v - x1) <= 2.5 for v in tall_v)
    has_divider = bool(cand)
    return {
        "bbox": (x0, y0, x1, y1),
        "divider": divider,
        "left": _merge_by_baseline(rows_l),
        "right": _merge_by_baseline(rows_r),
        "raw_left": rows_l,
        "raw_right": rows_r,
        "straddle": straddle,
        "row_seps": row_seps,
        "box_ok": has_top and has_bottom and has_left and has_right and has_divider,
    }
def _inside_boxes(row, boxes):
    """判断一个文本行是否落在任一表格外框内。"""
    for bx0, by0, bx1, by1 in boxes:
        if (row["x0"] >= bx0 - 3 and row["x1"] <= bx1 + 3
                and row["y0"] >= by0 - 3 and row["y1"] <= by1 + 8):
            return True
    return False


def _build_pages(doc):
    """把整篇文档解析成"逐页版面模型"：原始行、线段、两栏表格、表格外正文行。"""
    pages = []
    for pno in range(doc.page_count):
        page = doc[pno]
        raw = _raw_lines(page)
        segs_h, segs_v = _page_segments(page)
        try:
            found = page.find_tables()
            found_tabs = list(found.tables)
        except Exception:
            found_tabs = []
        boxes, cols, tables = [], [], []
        for t in found_tabs:
            bbox = (float(t.bbox[0]), float(t.bbox[1]), float(t.bbox[2]), float(t.bbox[3]))
            boxes.append(bbox)
            cols.append(int(t.col_count))
            if int(t.col_count) == 2:
                tables.append(_build_table(page, bbox, raw, segs_h, segs_v))
        outside = [r for r in raw if not _inside_boxes(r, boxes)]
        pages.append({
            "no": pno + 1,
            "width": float(page.rect.width),
            "height": float(page.rect.height),
            "raw": raw,
            "segs_h": segs_h,
            "segs_v": segs_v,
            "tables": tables,
            "table_boxes": boxes,
            "table_cols": cols,
            "raw_outside": outside,
            "body": _merge_by_baseline(outside),
        })
    return pages
def _cjk_count(text):
    return len(_CJK_RE.findall(text or ""))


def _column_langs(table):
    """判断两栏各自的主体语言，返回 ("C","E") 之类；用于拼接中英文内容流。"""
    lt = "".join(r["text"] for r in table["left"])
    rt = "".join(r["text"] for r in table["right"])
    lc, rc = _cjk_count(lt), _cjk_count(rt)
    if lc == rc:
        return (None, None)
    return ("C", "E") if lc > rc else ("E", "C")


def _find_footer(page_model):
    """找页脚行：位于页面底部、明显与正文拉开距离、内容短且含数字的那一行。"""
    rows = sorted(page_model["body"] + [r for t in page_model["tables"]
                                        for r in t["left"] + t["right"]],
                  key=lambda r: r["y0"])
    if not rows:
        return None
    last = rows[-1]
    text = last["text"].strip()
    if last["y0"] < page_model["height"] * 0.82:
        return None
    if len(text) > 12 or not re.search(r"\d", text):
        return None
    if len(rows) >= 2 and last["y0"] - rows[-2]["y1"] < 8.0:
        return None
    return last


def _build_model(doc):
    """在逐页版面模型之上，汇总条款题序、中英文内容流、页脚等文档级信息。"""
    pages = _build_pages(doc)
    clause_cn, clause_en, cn_parts, en_parts = [], [], [], []
    for pg in pages:
        for r in pg["body"]:
            text = r["text"].strip()
            m = _CLAUSE_CN_RE.match(text)
            if m and _CN_NUMERALS.get(m.group(1)):
                clause_cn.append({"n": _CN_NUMERALS[m.group(1)], "page": pg["no"], "row": r})
            m = _CLAUSE_EN_RE.match(text)
            if m:
                clause_en.append({"n": int(m.group(1)), "page": pg["no"], "row": r})
        items = [(r["y0"], 0, r) for r in pg["body"]]
        items += [(t["bbox"][1], 1, t) for t in pg["tables"]]
        items.sort(key=lambda it: (it[0], it[1]))
        for _y, kind, obj in items:
            if kind == 0:
                lang = _lang_of(obj["text"])
                if lang != "E":
                    cn_parts.append(obj["text"])
                if lang != "C":
                    en_parts.append(obj["text"])
            else:
                lang_l, lang_r = _column_langs(obj)
                left = " ".join(r["text"] for r in obj["left"])
                right = " ".join(r["text"] for r in obj["right"])
                if lang_l == "E" and lang_r == "C":
                    left, right = right, left
                cn_parts.append(left)
                en_parts.append(right)
    footers = [_find_footer(pg) for pg in pages]
    return {
        "pages": pages,
        "clause_cn": clause_cn,
        "clause_en": clause_en,
        "cn_stream": _norm(" ".join(cn_parts)),
        "en_stream": _norm(" ".join(en_parts)),
        "full_text": "\n".join(doc[i].get_text() for i in range(doc.page_count)),
        "footers": footers,
    }
def _clause_map(model):
    """条款号 -> {"page","cn","en","table"}：把条款题序标题与其下方的两栏表格关联。"""
    info = {}
    for c in model["clause_cn"]:
        info.setdefault(c["n"], {})["cn"] = c
    for c in model["clause_en"]:
        info.setdefault(c["n"], {})["en"] = c
    for _n, d in info.items():
        anchor = d.get("cn") or d.get("en")
        pg = next(p for p in model["pages"] if p["no"] == anchor["page"])
        ytop = min(x["row"]["y0"] for x in (d.get("cn"), d.get("en")) if x)
        cands = [t for t in pg["tables"] if t["bbox"][1] >= ytop - 2]
        d["page"] = pg
        d["table"] = min(cands, key=lambda t: t["bbox"][1]) if cands else None
    return info


def _row_subno(row):
    """取一行行首的子条款编号（"1.1"/"16.5"），行首没有编号则返回 None。"""
    m = _SUBNO_RE.match(row["text"].strip())
    if not m:
        return None
    return "%d.%d" % (int(m.group(1)), int(m.group(2)))


def _body_size(model):
    """正文字号：按字符数加权的众数（表头、标题占比小，不会抢到众数）。"""
    tally = Counter()
    for pg in model["pages"]:
        for r in pg["raw"]:
            for s in r["spans"]:
                text = (s.get("text") or "").strip()
                if text:
                    tally[round(s["size"], 1)] += len(text)
    return tally.most_common(1)[0][0] if tally else 0.0


def _all_rows(model):
    """全文合并后的可视行（表格内按左右栏各自合并），带页码与所属区域标记。"""
    out = []
    for pg in model["pages"]:
        for r in pg["body"]:
            out.append((pg, None, None, r))
        for t in pg["tables"]:
            for r in t["left"]:
                out.append((pg, t, "L", r))
            for r in t["right"]:
                out.append((pg, t, "R", r))
    return out


def _find_rows(model, keyword, in_table=None):
    """按归一化关键字检索可视行；in_table=False 只找表格外的行。"""
    key = _norm(keyword).lower()
    hits = []
    for pg, tab, side, r in _all_rows(model):
        if in_table is False and tab is not None:
            continue
        if in_table is True and tab is None:
            continue
        if key and key in _norm(r["text"]).lower():
            hits.append({"page": pg["no"], "table": tab, "side": side, "row": r})
    return hits


def _pair_cn_before_en(model, pairs, in_table=None):
    """逐对检查"中文行在前、英文行紧随其后"：返回 (成对且顺序正确的对数, 总对数, 首个反例)。"""
    ok, bad = 0, ""
    for cn_key, en_key in pairs:
        cn_hits = _find_rows(model, cn_key, in_table=in_table)
        en_hits = _find_rows(model, en_key, in_table=in_table)
        if not cn_hits or not en_hits:
            if not bad:
                bad = "%s / %s 未成对出现" % (cn_key, en_key)
            continue
        good = 0
        for ch in cn_hits:
            after = [eh for eh in en_hits
                     if (eh["page"], eh["row"]["y0"]) > (ch["page"], ch["row"]["y0"])]
            before = [eh for eh in en_hits
                      if (eh["page"], eh["row"]["y0"]) < (ch["page"], ch["row"]["y0"])]
            near_after = min((abs(eh["row"]["y0"] - ch["row"]["y0"])
                              for eh in after if eh["page"] == ch["page"]), default=1e9)
            near_before = min((abs(eh["row"]["y0"] - ch["row"]["y0"])
                               for eh in before if eh["page"] == ch["page"]), default=1e9)
            if near_after <= near_before:
                good += 1
        if good == len(cn_hits):
            ok += 1
        elif not bad:
            bad = "%s 的英文行出现在中文行之前" % cn_key
    return ok, len(pairs), bad
def _tag(hit, delta):
    return ("+%d：" % delta) if hit else "+0："


def _mtag(hit, delta):
    return ("-%d：" % delta) if hit else "-0："


def _left_aligned(rows, border):
    """栏内左对齐：以 12 磅为一档统计行首 x 偏移，众数档须紧贴栏边界。
    正文换行后的续行会缩进，占少数档，不影响判定。"""
    if not rows:
        return False
    buckets = Counter(max(0, int((r["x0"] - border) // 12)) for r in rows)
    return buckets.most_common(1)[0][0] * 12 <= 14


def _score_plus(model, results):
    """维度二加分项（19 项）：命中记满分、未命中记 0；顺序须与 _MAX_DELTAS_ORDER 一致。"""
    clauses = _clause_map(model)
    tables = [d["table"] for _n, d in sorted(clauses.items()) if d.get("table")]
    body_sz = _body_size(model)
    reformatted = len(tables) >= 8

    # 01 十六条条款全部两栏表格化，左栏中文右栏英文
    all16 = all(clauses.get(n, {}).get("table") for n in range(1, 17))
    lr_ok = bool(tables) and all(_column_langs(t) == ("C", "E") for t in tables)
    hit01 = all16 and lr_ok
    results.append(_tag(hit01, 5) + "第一条至第十六条共16条条款的子条款正文全部改为两栏表格呈现，"
                                    "左栏中文、右栏英文，中英文分列不混排")

    # 02 每个条款表格首行为"中文 / English"表头行
    hit02 = all16
    for t in tables:
        if not t["left"] or not t["right"]:
            hit02 = False
            break
        if "中文" not in _norm(t["left"][0]["text"]) or "english" not in _norm(t["right"][0]["text"]).lower():
            hit02 = False
            break
    results.append(_tag(hit02, 3) + '每个条款表格首行为表头行，左栏文字为"中文"、右栏文字为"English"')

    # 03 每一子条款中英成对占同一表格行，编号双栏均保留且一致
    expect = paired = aligned = 0
    for n, d in sorted(clauses.items()):
        need = _SUBCLAUSE_COUNTS.get(n, 0)
        expect += need
        t = d.get("table")
        if not t:
            continue
        lmap, rmap = {}, {}
        for r in t["left"]:
            k = _row_subno(r)
            if k and k not in lmap:
                lmap[k] = r
        for r in t["right"]:
            k = _row_subno(r)
            if k and k not in rmap:
                rmap[k] = r
        for i in range(1, need + 1):
            k = "%d.%d" % (n, i)
            if k in lmap and k in rmap:
                paired += 1
                dy = abs((lmap[k]["y0"] + lmap[k]["y1"]) / 2 - (rmap[k]["y0"] + rmap[k]["y1"]) / 2)
                if dy <= 20.0:
                    aligned += 1
    hit03 = expect > 0 and paired >= expect * 0.9 and aligned >= expect * 0.85
    results.append(_tag(hit03, 3) + "每一子条款（1.1、1.2、2.1……16.5）的中文与英文成对占同一表格行，"
                                    "左右一一对应，编号在中英两栏中均保留且一致")

    # 04 条款表格具有可见的完整边框网格
    box_ok = bool(tables) and all(t["box_ok"] for t in tables)
    with_sep = sum(1 for t in tables if t["row_seps"])
    hit04 = all16 and box_ok and with_sep >= len(tables) * 0.8
    results.append(_tag(hit04, 1) + "条款表格具有可见的完整边框网格（外边框与单元格分隔线齐全）")

    # 05 左右栏宽大致均等、栏内左对齐、无越栏压线
    hit05 = bool(tables)
    for t in tables:
        x0, _y0, x1, _y1 = t["bbox"]
        lw, rw = t["divider"] - x0, x1 - t["divider"]
        if lw <= 0 or rw <= 0 or not (0.6 <= lw / rw <= 1 / 0.6):
            hit05 = False
        if t["straddle"]:
            hit05 = False
        if not _left_aligned(t["left"], x0) or not _left_aligned(t["right"], t["divider"]):
            hit05 = False
    results.append(_tag(hit05, 1) + "左右两栏宽度大致均等，栏内文字各自左对齐，无文字越栏或压线")
    # 06 条款题序标题置于两栏表格上方，中前英后，且与表格内正文有字号或加粗区分
    hit06 = all16 and len(clauses) == 16
    for n, d in sorted(clauses.items()):
        cn, en, t = d.get("cn"), d.get("en"), d.get("table")
        if not cn or not en or not t:
            hit06 = False
            break
        if cn["page"] != en["page"] or cn["row"]["y0"] >= en["row"]["y0"]:
            hit06 = False
            break
        if en["row"]["y1"] > t["bbox"][1] + 3:
            hit06 = False
            break
        distinct = ((cn["row"]["any_bold"] and en["row"]["any_bold"])
                    or min(cn["row"]["size"], en["row"]["size"]) >= body_sz + 1.0)
        if not distinct:
            hit06 = False
            break
    results.append(_tag(hit06, 3) + '条款题序及标题（"第X条 ……"与"Clause X: ……"）置于该条款两栏表格上方，'
                                    "中文行在前、英文行紧随其后，且与表格内正文有明显字号或加粗区分")

    # 07 文档标题与各条款题序标题的中英两行字重一致（均加粗）
    title_cn = _find_rows(model, _TITLE_CN, in_table=False)
    title_en = _find_rows(model, _TITLE_EN, in_table=False)
    clause_rows = [d[k]["row"] for _n, d in sorted(clauses.items())
                   for k in ("cn", "en") if d.get(k)]
    checked = [h["row"] for h in title_cn[:1] + title_en[:1]] + clause_rows
    hit07 = (bool(title_cn) and bool(title_en) and len(clause_rows) == 32
             and all(r["bold"] for r in checked))
    results.append(_tag(hit07, 1) + "文档标题与各条款题序标题的中文行与英文行字重一致（中英文两行均为加粗）")

    # 08 文档标题为中文行在前、英文行紧随其后的两行式排布
    hit08 = False
    if title_cn and title_en:
        a, b = title_cn[0], title_en[0]
        hit08 = (a["page"] == b["page"] and a["row"]["y0"] < b["row"]["y0"]
                 and b["row"]["y0"] - a["row"]["y1"] <= 30.0)
    results.append(_tag(hit08, 3) + "文档标题为中文标题行在前、英文标题行紧随其后的两行式排布")

    # 09 签署前言与双方信息逐行中文在前、英文在后
    ok09, tot09, _bad09 = _pair_cn_before_en(model, _PREFACE_PAIRS, in_table=False)
    hit09 = ok09 == tot09
    results.append(_tag(hit09, 3) + '签署前言与双方信息部分（"本合同由以下双方于……签署"、"甲方/服务提供方"、'
                                    '"统一识别号"、"授权代表"、"地址"、"行政联系"、"乙方/客户方"、'
                                    '"甲方与乙方以下合称为\'双方\'"）逐行中文在前、英文在后成对排列')

    # 10 签署页逐行中前英后，且未被套用两栏表格
    ok10, tot10, _bad10 = _pair_cn_before_en(model, _SIGN_PAIRS, in_table=False)
    sign_hits = _find_rows(model, "签署页")
    hit10 = False
    if ok10 == tot10 and sign_hits:
        anchor = sign_hits[0]
        below = [tab for pg, tab, _s, r in _all_rows(model)
                 if pg["no"] > anchor["page"]
                 or (pg["no"] == anchor["page"] and r["y0"] >= anchor["row"]["y0"] - 2)]
        hit10 = anchor["table"] is None and not any(tab is not None for tab in below)
    results.append(_tag(hit10, 3) + '签署页部分（"签署页/Signature Page"、"甲方/服务提供方"、"授权代表签字"、'
                                    '"姓名"、"日期"、"乙方/客户方"等）逐行中文在前、英文在后成对排列，'
                                    "且未被套用两栏表格")

    # 11 全文中英顺序统一为"中文在前、英文在后"
    hit11 = hit06 and hit08 and hit09 and hit10 and lr_ok
    results.append(_tag(hit11, 1) + '全文中英文顺序统一为"中文在前、英文在后"，不存在中英颠倒的段落或表格行')
    # ---- 条款分页布局组：以"条款已完成两栏表格化"为计分前提 ----
    # 12 任一条款（含题序标题与整张表格）完整落在同一页内
    hit12 = reformatted and all16 and len(clauses) == 16
    if hit12:
        for n, d in sorted(clauses.items()):
            cn, en, t = d.get("cn"), d.get("en"), d.get("table")
            if not cn or not en or not t or cn["page"] != en["page"]:
                hit12 = False
                break
            found = {_row_subno(r) for r in t["left"]} | {_row_subno(r) for r in t["right"]}
            need = {"%d.%d" % (n, i) for i in range(1, _SUBCLAUSE_COUNTS[n] + 1)}
            if not need.issubset(found):
                hit12 = False
                break
    results.append(_tag(hit12, 5) + "任一条款（含其题序标题与整张两栏表格）均完整落在同一页内，全文无跨页拆分的条款")

    # 13 多条条款同页排布，未出现"一条一页"造成的大面积空白
    per_page = Counter((d.get("cn") or d.get("en"))["page"] for d in clauses.values())
    multi = sum(1 for v in per_page.values() if v >= 2)
    hit13 = reformatted and multi >= 2 and len(per_page) <= 12
    results.append(_tag(hit13, 3) + '实现多条条款同页排布（除首页与签署页外，正文页普遍容纳2条及以上条款），'
                                    '未出现"一条一页"造成的大面积空白')

    # 14 条款与条款之间保留可辨的空白间距
    ordered = sorted(((d.get("cn") or d.get("en"))["page"],
                      (d.get("cn") or d.get("en"))["row"]["y0"], n)
                     for n, d in clauses.items())
    hit14, seen_gap, prev = reformatted, False, None
    for pno, y0, n in ordered:
        if prev and prev[0] == pno and prev[1] is not None:
            seen_gap = True
            if y0 - prev[1] < 5.0:
                hit14 = False
        t = clauses[n].get("table")
        prev = (pno, t["bbox"][3] if t else None)
    hit14 = hit14 and seen_gap
    results.append(_tag(hit14, 1) + "条款与条款之间保留可辨的空白间距，后一条题序标题未与前一条表格紧贴或压行")

    # 15 单元格内每处填空下划线保持在同一行内连续显示，未被折行拆断
    lines = model["full_text"].split("\n")
    broken = sum(1 for i in range(len(lines) - 1)
                 if lines[i].rstrip().endswith("_") and lines[i + 1].lstrip().startswith("_"))
    hit15 = reformatted and broken == 0
    results.append(_tag(hit15, 1) + "条款表格单元格内每处填空下划线保持在同一行内连续显示，未被折行拆断为上下两段")

    # 16 各子条款编号与其正文位于同一行行首，未出现编号单独成行
    alone = 0
    for _pg, _tab, _side, r in _all_rows(model):
        text = r["text"].strip()
        m = _SUBNO_RE.match(text)
        if m and not _norm(text[m.end():]):
            alone += 1
    hit16 = reformatted and alone == 0
    results.append(_tag(hit16, 1) + "各子条款编号（如6.5、7.5、15.1）与其正文位于同一行行首，"
                                    "未出现编号单独成行、正文另起一行的情况")

    # ---- 页码组 ----
    footers = model["footers"]
    nums = []
    for f in footers:
        s = _norm(f["text"]) if f else ""
        nums.append(int(s) if s.isdigit() else None)
    all_num = bool(nums) and all(v is not None for v in nums)
    hit17 = all_num and all(nums[i + 1] == nums[i] + 1 for i in range(len(nums) - 1))
    results.append(_tag(hit17, 3) + "全部页面添加页脚页码，为阿拉伯数字且连续编号")

    hit18 = all_num and all(abs((f["x0"] + f["x1"]) / 2 - pg["width"] / 2) <= 20.0
                            for pg, f in zip(model["pages"], footers))
    results.append(_tag(hit18, 1) + "页码水平居中于页面左右中线")

    # 兼容 PDF 后端的字体 bbox 顶边相对真实墨迹约有 1–2pt 波动，保留 2pt
    # 容差；7pt 左右的页码顶边仍明显处于页脚带内，不能据此判为贴底。
    hit19 = all_num and all(pg["height"] - f["y0"] >= 6.0 and f["y0"] >= pg["height"] - 80.0
                            for pg, f in zip(model["pages"], footers))
    results.append(_tag(hit19, 1) + "页码垂直位置位于页脚区域中部，既不紧贴页面下边缘，也不与正文最后一行相接")
def _score_minus(model, results):
    """维度二减分项（18 项）：均为指令未要求、却被改动的多余修改；命中记负分。
    顺序须与 _MIN_DELTAS_ORDER 一致。"""
    cn = model["cn_stream"]
    en = model["en_stream"].lower()
    full = model["full_text"]
    clauses = _clause_map(model)
    tables = [d["table"] for _n, d in sorted(clauses.items()) if d.get("table")]

    # 20 第16.5条原文未逐字保留，或其中英排布被单独改为"英文在前"
    # 一次比对同时覆盖两个判定面：cn_stream / en_stream 是按"中文栏 / 英文栏"分别拼出来的，
    # 所以①改写原文 → 原句不在对应语言流中；②只把 16.5 这一行的中英位置调换（整栏语言
    # 多数仍是左中右英，_column_langs 不会整体交换）→ 英文句落进 cn_stream，中文句落进
    # en_stream；③整句删除 → 两侧都找不到。三种情形均命中。
    hit20 = not (_norm(_C1605_CN) in cn and _norm(_C1605_EN).lower() in en)
    results.append(_mtag(hit20, 1) + '第16.5条原文未逐字保留，或其中英排布被单独改为"英文在前"——须保持'
                                     '"16.5 英文文本列于对应中文文本之上. 如解释存在差异, 应以最符合商业'
                                     '目的的文本为准."与"16.5 The English text appears above the corresponding '
                                     'Chinese text. If interpretation differs, the language most consistent '
                                     'with the commercial purpose shall prevail."，且该子条款仍与全文一致地'
                                     '中文在前、英文在后（不得为使条款描述与新版面自洽而改写文字或调换中英'
                                     '位置；改写、互换、删改均计）')

    # 21 条款正文/标题/签署栏文字被删除或改写；22 条款、子条款、前言或签署页内容缺失
    missing, changed = [], []
    for key in _CANON_SUBCLAUSES:
        if key not in _SUB_TEXT:
            continue
        c_txt, e_txt = _SUB_TEXT[key]
        if key not in cn or key not in en:
            missing.append(key)
        elif _sentence_lost(c_txt, cn) or _sentence_lost(e_txt.lower(), en):
            changed.append(key)
    for _n, _cn_no, cn_title, _en_no, en_title in _CLAUSE_TITLES:
        if _norm(cn_title) not in cn:
            missing.append(cn_title)
        if _norm(en_title).lower() not in en:
            missing.append(en_title)
    both = (cn + en).lower()
    for lab in _PREFACE_LABELS + _SIGN_LABELS:
        if _norm(lab) and _norm(lab).lower() not in both:
            missing.append(lab)
    results.append(_mtag(bool(changed), 3) + "删除或改写条款正文、标题、签署栏文字（增删、替换非格式性的文字内容）")
    results.append(_mtag(bool(missing), 3) + "第一条至第十六条条款、子条款、签署前言或签署页内容缺失"
                                             "（删除任一条款、子条款、双方信息行或签署栏均计）")

    # 23 条款编号、附件编号或交叉引用数值被改动
    ref_bad = []
    for token in ("附件A", "附件B", "第14.3条"):
        if _norm(token) not in cn:
            ref_bad.append(token)
    for token in ("Schedule A", "Schedule B", "Clause 14.3"):
        if _norm(token).lower() not in en:
            ref_bad.append(token)
    for _n, cn_no, _ct, en_no, _et in _CLAUSE_TITLES:
        if _norm(cn_no) not in cn:
            ref_bad.append(cn_no)
        if _norm(en_no).lower() not in en:
            ref_bad.append(en_no)
    results.append(_mtag(bool(ref_bad), 1) + '条款编号、附件编号（附件A、附件B / Schedule A、Schedule B）'
                                             '或交叉引用（如"第14.3条"、"Clause 14.3"）的数值被改动')

    # 24 擅自新增指令未要求的内容
    extra = []
    for kw in ("水印", "目录", "封面", "以下空白", "本页以下无正文", "此页无正文", "备注",
               "免责声明", "Watermark", "Table of Contents", "Disclaimer", "Note:", "草稿", "Draft"):
        if _norm(kw).lower() in both:
            extra.append(kw)
    tops = Counter()
    for pg in model["pages"]:
        rows = sorted(pg["body"] + [r for t in pg["tables"] for r in t["left"] + t["right"]],
                      key=lambda r: r["y0"])
        if rows:
            tops[_norm(rows[0]["text"])] += 1
    for text, cnt in tops.items():
        if text and cnt >= 3:
            extra.append("每页重复页眉：" + text[:20])
    results.append(_mtag(bool(extra), 1) + '擅自新增指令未要求的内容（新增页眉文字、水印、目录、封面、条款说明、'
                                           '"以下空白"等占位文字）')

    # 25 正文冒号被替换为全角"："；26 句内逗号被替换为全角"，"
    results.append(_mtag(full.count("：") > 0, 1) + '正文冒号不满足半角":"形式——"甲方/服务提供方:"、"统一识别号:"、'
                                                   '"授权代表:"、"地址:"、"行政联系:"、"乙方/客户方:"、"姓名:"、'
                                                   '"日期:"等被替换为全角"："')
    results.append(_mtag(full.count("，") > 0, 1) + '条款与前言句内逗号不满足半角","形式——如"1.1 甲方向乙方出租一套'
                                                   '智能办公设备组合, 仅供乙方内部办公使用"、"2.2 甲方应于……或之前将'
                                                   '设备交付乙方, 并在移交前完成基础配置"等处逗号被替换为全角"，"')
    # 27 文档标题字号偏离约15.5磅
    title_rows = [h["row"] for h in _find_rows(model, _TITLE_CN, in_table=False)[:1]]
    title_rows += [h["row"] for h in _find_rows(model, _TITLE_EN, in_table=False)[:1]]
    t_sz = [r["size"] for r in title_rows]
    hit27 = (not t_sz) or any(s > 16.0 or s < 14.0 for s in t_sz)
    results.append(_mtag(hit27, 1) + '文档标题"智能办公设备租赁及支持服务合同"与"Managed Workspace Equipment '
                                     'Leasing and Support Contract"字号不满足约15.5磅（被放大至16磅以上或'
                                     '缩小至14磅以下均计）')

    # 28 条款题序标题字号被放大至14磅以上
    c_sz = [d[k]["row"]["size"] for _n, d in sorted(clauses.items())
            for k in ("cn", "en") if d.get(k)]
    hit28 = bool(c_sz) and max(c_sz) > 14.0
    results.append(_mtag(hit28, 1) + '各条款题序及标题（"第X条 ……"、"Clause X: ……"）字号不满足约12.5磅'
                                     '（被放大至14磅以上均计）')

    # 29 签署页字号被单独放大为标题级
    sign_hits = _find_rows(model, "签署页")
    s_sz = []
    if sign_hits:
        anchor = sign_hits[0]
        for pg, _tab, _side, r in _all_rows(model):
            if (pg["no"] > anchor["page"]
                    or (pg["no"] == anchor["page"] and r["y0"] >= anchor["row"]["y0"] - 2)):
                s_sz.append(r["size"])
    hit29 = bool(s_sz) and max(s_sz) > 12.5
    results.append(_mtag(hit29, 1) + '签署页"签署页"、"Signature Page"及其下方各签署信息行字号不满足与正文一致的'
                                     '约10.8磅（被单独放大为标题级字号均计）')

    # 30 页面左右边距被改为宽边距
    xs0 = [r["x0"] for pg in model["pages"] for r in pg["raw"]]
    xs1 = [r["x1"] for pg in model["pages"] for r in pg["raw"]]
    page_w = model["pages"][0]["width"] if model["pages"] else 0.0
    hit30 = bool(xs0) and (min(xs0) > 40.0 or (page_w - max(xs1)) > 40.0)
    results.append(_mtag(hit30, 1) + "页面左右边距不满足约18磅（约0.65cm）的窄边距设置，被改为2cm级宽边距"
                                     "而使每行可容纳内容明显减少")

    # 31 纸张非A4纵向
    hit31 = any(not (abs(pg["width"] - 595.28) <= 6.0 and abs(pg["height"] - 841.89) <= 8.0)
                for pg in model["pages"])
    results.append(_mtag(hit31, 1) + "纸张不满足A4（宽210mm、高297mm）纵向（被改为Letter、A3等其他尺寸或"
                                     "改为横向均计）")

    # 32 页码格式非纯阿拉伯数字，或页脚被额外加料
    footers = [f for f in model["footers"] if f]
    hit32 = bool(footers) and not all(_norm(f["text"]).isdigit() for f in footers)
    marker = re.compile(r"^(第\d{1,3}页|-\d{1,3}-|\d{1,3}/\d{1,3}|page\d{1,3})", re.I)
    for pg in model["pages"]:
        rows = sorted(pg["body"] + [r for t in pg["tables"] for r in t["left"] + t["right"]],
                      key=lambda r: r["y0"])
        if rows and marker.match(_norm(rows[-1]["text"])):
            hit32 = True
    results.append(_mtag(hit32, 1) + '页码格式被改为"第X页 / 共Y页"、"- X -"等非纯阿拉伯数字形式，或在页脚额外'
                                     "添加文件名、日期、公司名等文字")

    # 33 非条款内容被误套用两栏表格
    hit33 = False
    if clauses:
        first = min(((d.get("cn") or d.get("en"))["page"],
                     (d.get("cn") or d.get("en"))["row"]["y0"]) for d in clauses.values())
        for pg, tab, _side, r in _all_rows(model):
            if tab is not None and (pg["no"], r["y0"]) < first:
                hit33 = True
    if sign_hits:
        anchor = sign_hits[0]
        for pg, tab, _side, r in _all_rows(model):
            if tab is not None and (pg["no"], r["y0"]) >= (anchor["page"], anchor["row"]["y0"] - 2):
                hit33 = True
    results.append(_mtag(hit33, 1) + '签署前言、"甲方与乙方以下合称为\'双方\'"或签署页等非条款内容被误套用'
                                     "两栏表格呈现")

    # 34 中文被设为非中文字体导致字形异常
    hit34 = False
    for pg in model["pages"]:
        for r in pg["raw"]:
            for s in r["spans"]:
                if _CJK_RE.search(s.get("text") or ""):
                    name = (s.get("font") or "").lower()
                    if any(k in name for k in _LATIN_ONLY_FONTS):
                        hit34 = True
    results.append(_mtag(hit34, 1) + "中文文字被设置为非中文字体（如英文字体或日文/韩文字体）导致字形异常、"
                                     "缺字或显示为方框")

    # 35 英文单词被硬拆断（无连字符处插入连字符）或空格粘连
    canon_en = _norm(" ".join(e for _c, e in _SUB_TEXT.values()) + " " + _C1605_EN).lower()
    hit35 = len(re.findall(r"[a-z]-[a-z]", en)) > len(re.findall(r"[a-z]-[a-z]", canon_en))
    results.append(_mtag(hit35, 1) + "英文单词在无连字符处被硬拆断换行，或英文单词间空格被删除导致单词粘连")

    # 36 填空下划线被压缩或长短不一
    runs = re.findall(r"_{3,}", full)
    hit36 = len({len(r) for r in runs}) > 6
    results.append(_mtag(hit36, 1) + "填空下划线长度被大幅压缩或长短不一，同类填空空位宽度明显不统一")

    # 37 表格行高/内边距压缩至文字碰边框，或中英两栏同一子条款起始行明显错位
    hit37 = False
    for t in tables:
        for sep in t["row_seps"]:
            for r in t["left"] + t["right"]:
                if r["y0"] + 1.5 < sep < r["y1"] - 1.5:
                    hit37 = True
        lmap, rmap = {}, {}
        for r in t["left"]:
            k = _row_subno(r)
            if k and k not in lmap:
                lmap[k] = r
        for r in t["right"]:
            k = _row_subno(r)
            if k and k not in rmap:
                rmap[k] = r
        for key in set(lmap) & set(rmap):
            if abs(lmap[key]["y0"] - rmap[key]["y0"]) > 30.0:
                hit37 = True
    results.append(_mtag(hit37, 1) + "表格行高、单元格内边距被压缩至文字与边框相碰，或中英两栏行数差异导致"
                                     "同一子条款中英文起始行明显错位")


def _check_dim1(doc, model):

    """维度一：交付文件格式正确、可用可修改；版面严重错乱则整体不得分。

    「必须是可正常打开的 .pdf」这一格式门禁已由 _fmt_reject 在 evaluate() 里先判过，
    这里只查版面与文字层。
    """
    if doc.page_count < 1:
        return False, "文件无有效页面"
    if len(_norm(model["full_text"])) < 500:
        return False, "文件缺少可提取的文字层（疑为扫描件或渲染失败），无法二次编辑"
    over = 0
    for pg in model["pages"]:
        for r in pg["raw"]:
            if (r["x0"] < -8 or r["x1"] > pg["width"] + 8
                    or r["y0"] < -8 or r["y1"] > pg["height"] + 8):
                over += 1
    if over > doc.page_count:
        return False, f"有 {over} 行文字超出纸张范围，版面溢出严重"
    cols = [c for pg in model["pages"] for c in pg["table_cols"]]
    if cols and sum(1 for c in cols if c == 2) * 2 < len(cols):
        return False, f"检出 {len(cols)} 个表格，其中仅 {sum(1 for c in cols if c == 2)} 个为两栏结构，表格结构错乱"
    return True, "文件可正常打开、文字层完整、版面无溢出，可直接二次编辑"


def _evaluate_pdf(pdf_path, result):
    doc = fitz.open(pdf_path)
    try:
        model = _build_model(doc)
        ok, reason = _check_dim1(doc, model)
        result["dim1_pass"] = bool(ok)
        # 维度一未通过时，dim1_reason 必须带上错误摘要，便于直接定位原因
        result["dim1_reason"] = reason if ok else ("维度一未通过：%s，完成度总分记 0" % reason)
        if not ok:
            result["dim2_items"] = []
            result["total_score"] = 0
            return
        plus_results, minus_results = [], []
        _score_plus(model, plus_results)
        _score_minus(model, minus_results)
        if len(plus_results) != len(_MAX_DELTAS_ORDER):
            raise RuntimeError(
                f"加分项数不一致：results={len(plus_results)} vs order={len(_MAX_DELTAS_ORDER)}")
        if len(minus_results) != len(_MIN_DELTAS_ORDER):
            raise RuntimeError(
                f"减分项数不一致：results={len(minus_results)} vs order={len(_MIN_DELTAS_ORDER)}")
        dim2_items, total = [], 0
        for raw, max_delta in zip(plus_results, _MAX_DELTAS_ORDER):
            head, _, rule = raw.partition("：")
            delta = int(head.replace("+", "").strip())
            hit = delta > 0
            total += max_delta if hit else 0
            dim2_items.append({"rule": rule, "max_delta": max_delta,
                               "delta": max_delta if hit else 0,
                               "hit": hit})
        for raw, max_delta in zip(minus_results, _MIN_DELTAS_ORDER):
            head, _, rule = raw.partition("：")
            delta = int(head.replace("+", "").strip())
            hit = delta < 0
            total += max_delta if hit else 0
            dim2_items.append({"rule": rule, "max_delta": max_delta,
                               "delta": max_delta if hit else 0,
                               "hit": hit})
        result["dim2_items"] = dim2_items
        result["total_score"] = max(0, total)
    finally:
        doc.close()


def evaluate(dir_path=None):
    """评估本脚本所在目录下的唯一交付物。

    dir_path 参数仅为保持统一接口而保留：无论传入什么路径，实际评估目标一律锁定为
    本脚本所在目录（_SCRIPT_DIR），既不进入子目录，也不会去别的目录找文件。

    - 交付物只接受 .pdf：扩展名不是 .pdf、或文件打不开（缺 %PDF- 头、加密、零页），
      一律在维度一判不通过（dim1_pass=False、total_score=0），而不是当成没有交付物。
      目录里没有 .pdf 但有 .docx/.doc 等其他格式时，仍收作被评对象走这条判负路径。
    - 定位不到交付文件时：error 只回纯文本原因（不带异常类名），
      dim1_pass 与 dim1_reason 一律留空（空字符串，表示维度一根本没进入判定），
      dim2_items=[]、total_score=max_score=0。
    - 维度一未通过时：dim1_pass=False，且 dim1_reason 带出错误摘要。
    """
    dir_path = dir_path if dir_path and os.path.isdir(dir_path) else _SCRIPT_DIR
    result = {
        "id": SCRIPT_ID,
        "file_name": "",
        "status": "ok",
        "error": "",
        "dim1_pass": False,
        "dim1_reason": "",
        "dim2_items": [],
        "total_score": 0,
        "max_score": MAX_SCORE_TOTAL,
    }
    try:
        if not os.path.isdir(dir_path):
            raise FileNotFoundError("目录不存在或不是目录：%s" % dir_path)
        file_path = _locate_document(dir_path)
        result["file_name"] = os.path.basename(file_path)
        # 维度一的第一道门禁：交付物格式只接受 .pdf，不合格即判 0，不再解析版面
        reject = _fmt_reject(file_path)
        if reject:
            result["dim1_pass"] = False
            result["dim1_reason"] = "维度一未通过：%s，完成度总分记 0" % reject
            result["dim2_items"] = []
            result["total_score"] = 0
            return result
        if fitz is None:
            raise RuntimeError("缺少 PDF 解析后端依赖，无法量测交付 PDF 的版面")
        # 门禁已保证走到这里的一定是 .pdf，直接按交付文件本身量测版面
        _evaluate_pdf(file_path, result)
    except FileNotFoundError as exc:
        # 找不到被评估文件：error 只回传纯文本原因，不带异常类名；维度一根本没进入
        # 判定，dim1_pass 与 dim1_reason 一律留空（空字符串，而非 False），满分归零
        result["status"] = "error"
        result["error"] = str(exc)
        result["file_name"] = ""
        result["dim1_pass"] = ""
        result["dim1_reason"] = ""
        result["dim2_items"] = []
        result["total_score"] = 0
        result["max_score"] = 0
    except Exception as exc:
        # 文件在但打不开/解析失败：维度一未通过，dim1_reason 带上错误摘要便于定位
        result["status"] = "error"
        result["error"] = f"脚本运行异常：{type(exc).__name__}: {exc}"
        result["dim1_pass"] = False
        result["dim1_reason"] = f"文档无法打开或解析：{type(exc).__name__}: {exc}"
        result["dim2_items"] = []
        result["total_score"] = 0
    return result


if __name__ == "__main__":
    # 不接受命令行目录参数：评估目标固定为本脚本所在目录
    print(json.dumps(evaluate(), ensure_ascii=False, indent=2))
