#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# pyright: reportUnknownMemberType=false, reportUnknownVariableType=false, reportUnknownArgumentType=false, reportUnknownParameterType=false, reportMissingTypeArgument=false, reportImplicitStringConcatenation=false, reportDeprecated=false, reportExplicitAny=false, reportGeneralTypeIssues=false, reportArgumentType=false, reportAttributeAccessIssue=false, reportPrivateUsage=false, reportOptionalMemberAccess=false, reportOptionalSubscript=false, reportOptionalIterable=false, reportOptionalCall=false
# 说明：本模块统一对外接口约定，不修改全局 sys.stdout，也不 print 主结果。
# 主结果统一由 evaluate(dir_path) -> dict 返回，入参为"脚本所在目录的路径"，
# 由脚本自己在该目录内定位并打开被评估的 .docx（忽略以 ~$ 开头的 Word 临时
# 锁文件）。本文件仅在 __main__ 下作为本地调试入口，把 dict 序列化为 JSON 打印。
#
# 评估逻辑（保持不变）：
#   维度一（可用与可修改性）为门槛：不通过直接判 0 分，不再检查维度二。
#   维度二分为得分点（+）与扣分点（-）：得分点满足该细则每一点才累加其分值；
#   扣分点满足任意一点即扣其分值。最终得分 = 命中得分点之和 - 命中扣分点之和。
#
# 主结果由 evaluate(dir_path) -> dict 返回，字段：
#   id / file_name / status / error / dim1_pass / dim1_reason /
#   dim2_items / total_score / max_score

import os
import re
import sys
import json
import zipfile
import contextlib
import xml.etree.ElementTree as ET
from docx import Document
from docx.oxml.ns import qn
from docx.shared import Emu


SCRIPT_ID = "016"
# 满分（max_score）= 所有得分点分值之和（不含扣分点）；其构成见文末 _DIM2_RULES，
# 并在该表定义处以 assert 校验一致。
MAX_SCORE_TOTAL = 96


W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
EMU_PER_CM = 360000.0
EMU_PER_PT = 12700.0

# 字号（磅）对照表，用于把 w:sz（半磅）换算并与细则中的中文字号名对应
PT_XIAOCHU = 36.0   # 小初
PT_YIHAO = 26.0     # 一号
PT_XIAOYI = 24.0    # 小一
PT_ERHAO = 22.0     # 二号
PT_XIAOER = 18.0    # 小二
PT_SANHAO = 16.0    # 三号
PT_SIHAO = 14.0     # 四号
PT_XIAOSI = 12.0    # 小四
PT_WUHAO = 10.5     # 五号


class WordDocumentEvaluator:
    """成教本科毕业论文评估器"""

    def __init__(self, file_path):
        self.file_path = file_path
        self.open_path = file_path
        self.doc = None
        # 穿透 w:sdt 的扁平文档索引（在维度二开始时构建）
        self._flat_paragraphs = []
        self._flat_para_elements = []
        self._flat_tables = []
        self._flat_sections = []
        self._para_section_index = []
        # 关键页定位（扁平段落索引）
        self.cover_title_idx = -1     # "成教本科毕业论文"
        self.decl_title_idx = -1      # "原创性声明及使用授权声明"
        self.toc_title_idx = -1       # "目录"
        self.body_title_idx = -1      # 正文首页论文标题
        self.abstract_idx = -1        # "摘 要："
        self.keywords_idx = -1        # "关键词："
        self.results = {
            'dimension1_passed': True,
            'dimension1_failures': [],
            'dimension2_scores': [],
            'dimension2_deductions': [],
            'total_score': 0,
            'hit_points': [],
            # 按 _DIM2_RULES 顺序记录每条规则的评估结果：{hit, delta, detail}
            'dim2_outcomes': [],
        }

    # ================= 主流程 =================
    def evaluate(self):
        print("=" * 60)
        print("Word文档自动评估报告")
        print("=" * 60)
        print(f"文件: {os.path.basename(self.file_path)}")
        print()

        if not self._evaluate_dimension1():
            self.results['dimension1_passed'] = False
            self._print_results()
            return

        self._evaluate_dimension2()
        self._calculate_total_score()
        self._print_results()

    # ================= 维度一 =================
    def _evaluate_dimension1(self):
        print("-" * 40)
        print("维度一：可用与可修改性")
        print("-" * 40)

        if os.path.splitext(self.file_path)[1].lower() != '.docx':
            self.results['dimension1_failures'].append("文件格式不是 .docx")
            return False
        print("[PASS] 文件格式：.docx")

        # 可正常打开（不报"需要修复"）：能被 zip 解出且 python-docx 能加载
        try:
            with zipfile.ZipFile(self.open_path) as zf:
                bad = zf.testzip()
                if bad is not None:
                    self.results['dimension1_failures'].append(
                        f"压缩包损坏，无法打开：{bad}")
                    return False
                # 关键部件存在
                names = set(zf.namelist())
                if 'word/document.xml' not in names:
                    self.results['dimension1_failures'].append(
                        "缺少 word/document.xml，文档结构损坏")
                    return False
                # document.xml 可被 XML 解析（乱码/损坏会在此暴露）
                ET.fromstring(zf.read('word/document.xml'))
        except Exception as e:
            self.results['dimension1_failures'].append(f"文件无法正常打开：{e}")
            return False

        try:
            self.doc = Document(self.open_path)
        except Exception as e:
            self.results['dimension1_failures'].append(f"python-docx 无法加载：{e}")
            return False

        # 有可编辑正文内容
        text = "".join(p.text for p in self.doc.paragraphs)
        if len(text.strip()) < 50:
            self.results['dimension1_failures'].append("文档缺少可编辑正文内容")
            return False

        print("[PASS] 文件可正常打开、可编辑，无乱码/损坏")
        print()
        return True

    # ================= 维度二 =================
    def _evaluate_dimension2(self):
        print("-" * 40)
        print("维度二：完成度")
        print("-" * 40)

        self._build_flat_document_index()
        self._locate_key_paragraphs()

        # ---- 得分点 ----
        self._check_cover_structure()            # +5
        self._check_cover_title_font()           # +1
        self._check_cover_images()               # +3
        self._check_cover_field_rows_font()       # +3
        self._check_cover_date()                 # +1
        self._check_cover_single_page()          # +1
        self._check_decl_page()                  # +5
        self._check_decl_title_font()            # +1
        self._check_decl_paragraphs()            # +3
        self._check_decl_paragraphs_font()       # +1
        self._check_decl_signature()             # +1
        self._check_toc_page()                   # +5
        self._check_toc_auto()                   # +3
        self._check_body_title()                 # +3
        self._check_abstract()                   # +3
        self._check_keywords()                   # +1
        self._check_body_font()                  # +5
        self._check_body_first_line_indent()     # +3
        self._check_body_line_spacing()          # +3
        self._check_body_justify()               # +3
        self._check_h1_style()                   # +3
        self._check_h1_numbering()               # +3
        self._check_h1_font()                    # +3
        self._check_h2_style()                   # +3
        self._check_h2_numbering()               # +5
        self._check_h2_font()                    # +3
        self._check_page_margins()               # +3
        self._check_footer_page_number()         # +3
        self._check_footer_page_number_format()  # +1
        self._check_body_first_page_number()     # +3 (逐页递增/首页第1页)
        self._check_sections_distinct()          # +3
        self._check_tail_pages_new_page()        # +1
        self._check_ref_ack_title_style()        # +3
        self._check_ref_bracket_number()         # +3
        self._check_ref_font()                   # +3
        self._check_ref_hanging_indent()         # +1

        # ---- 扣分点 ----（顺序须与 _DIM2_RULES 扣分段严格一致）
        self._check_extra_full_page_content()    # -3 结构外整页内容(评审表等)
        self._check_body_extra_images()          # -3 封面两张之外的图片
        self._check_revisions_comments()         # -1 批注/修订
        self._check_color_highlight_strike()     # -1 彩色/高亮/删除线
        self._check_header_unrelated()           # -1 页眉出现文字或横线
        self._check_blank_page()                 # -1 空白整页
        self._check_consecutive_blank_lines()    # -1 连续3行空行
        self._check_missing_chapters()           # -3 缺少应有章节
        self._check_body_text_tamper()           # -3 正文占位/截断/重复
        self._check_reference_content()          # -3 参考文献内容缺损
        self._check_garbled_chars()              # -1 乱码/方块字
        self._check_extra_structure_content()    # -1 结构外多余内容
        self._check_page_size_orientation()      # -1 非A4/横向
        self._check_hyperlink_footnote()         # -1 无关超链接/脚注尾注

    # ================= 扁平文档索引 =================
    def _build_flat_document_index(self):
        """按文档流顺序收集顶层段落/表格/节，穿透 w:sdt/w:sdtContent。"""
        from docx.text.paragraph import Paragraph
        from docx.table import Table
        from docx.section import Section

        self._flat_paragraphs = []
        self._flat_para_elements = []
        self._flat_tables = []
        self._flat_sections = []
        self._para_section_index = []

        sect_idx = [0]
        part = self.doc.part

        def walk(container):
            for child in container.iterchildren():
                tag = child.tag.split('}')[-1]
                if tag == 'p':
                    self._flat_paragraphs.append(Paragraph(child, self.doc))
                    self._flat_para_elements.append(child)
                    self._para_section_index.append(sect_idx[0])
                    pPr = child.find(qn('w:pPr'))
                    if pPr is not None:
                        sp = pPr.find(qn('w:sectPr'))
                        if sp is not None:
                            self._flat_sections.append(Section(sp, part))
                            sect_idx[0] += 1
                elif tag == 'tbl':
                    self._flat_tables.append(Table(child, self.doc))
                elif tag == 'sdt':
                    sdt_content = child.find(qn('w:sdtContent'))
                    if sdt_content is not None:
                        walk(sdt_content)
                elif tag == 'sectPr':
                    self._flat_sections.append(Section(child, part))
                    sect_idx[0] += 1

        walk(self.doc.element.body)

    def _locate_key_paragraphs(self):
        """定位封面标题、声明标题、目录标题、正文标题、摘要、关键词的段落。"""
        for i, p in enumerate(self._flat_paragraphs):
            norm = re.sub(r'\s+', '', p.text)
            if not norm:
                continue
            if self.cover_title_idx == -1 and norm == '成教本科毕业论文':
                self.cover_title_idx = i
            elif self.decl_title_idx == -1 and '原创性声明' in norm and '声明' in norm:
                self.decl_title_idx = i
            elif self.toc_title_idx == -1 and norm in ('目录', '目次'):
                self.toc_title_idx = i
            elif self.abstract_idx == -1 and norm.startswith('摘要'):
                self.abstract_idx = i
            elif self.keywords_idx == -1 and norm.startswith('关键词'):
                self.keywords_idx = i
        # 正文标题：目录之后、摘要之前，第一段非空且不含制表符/页码的居中大标题
        start = self.toc_title_idx + 1 if self.toc_title_idx >= 0 else 0
        end = self.abstract_idx if self.abstract_idx >= 0 else len(self._flat_paragraphs)
        for i in range(start, end):
            t = self._flat_paragraphs[i].text.strip()
            if t and '\t' not in t and '…' not in t:
                self.body_title_idx = i
                break

    # ================= 通用解析辅助 =================
    def _para_section(self, idx):
        if idx < 0 or idx >= len(self._para_section_index):
            return -1
        return self._para_section_index[idx]

    def _run_east_asia_font(self, run, paragraph):
        """解析 run 生效的中文字体：run 直接设置 → 段落样式链。"""
        rPr = run._element.rPr
        if rPr is not None:
            rFonts = rPr.find(qn('w:rFonts'))
            if rFonts is not None:
                ea = rFonts.get(qn('w:eastAsia'))
                if ea:
                    return ea
        return self._style_east_asia_font(paragraph)

    def _style_east_asia_font(self, paragraph):
        try:
            style = paragraph.style
        except Exception:
            return None
        visited = set()
        while style is not None and id(style) not in visited:
            visited.add(id(style))
            try:
                sr = style.element.find(qn('w:rPr'))
            except Exception:
                sr = None
            if sr is not None:
                rFonts = sr.find(qn('w:rFonts'))
                if rFonts is not None:
                    ea = rFonts.get(qn('w:eastAsia'))
                    if ea:
                        return ea
            try:
                style = style.base_style
            except Exception:
                break
        return None

    def _run_size_pt(self, run, paragraph):
        """解析 run 生效字号（磅）：run → 段落样式链。"""
        try:
            if run.font.size is not None:
                return run.font.size.pt
        except Exception:
            pass
        try:
            style = paragraph.style
        except Exception:
            style = None
        visited = set()
        while style is not None and id(style) not in visited:
            visited.add(id(style))
            try:
                sz = style.font.size
            except Exception:
                sz = None
            if sz is not None:
                return sz.pt
            try:
                style = style.base_style
            except Exception:
                break
        return None

    def _run_bold(self, run, paragraph):
        try:
            if run.font.bold is not None:
                return bool(run.font.bold)
        except Exception:
            pass
        try:
            style = paragraph.style
        except Exception:
            style = None
        visited = set()
        while style is not None and id(style) not in visited:
            visited.add(id(style))
            try:
                b = style.font.bold
            except Exception:
                b = None
            if b is not None:
                return bool(b)
            try:
                style = style.base_style
            except Exception:
                break
        return False

    def _resolve_alignment(self, paragraph):
        align = paragraph.paragraph_format.alignment
        if align is not None:
            return align
        try:
            style = paragraph.style
        except Exception:
            return None
        visited = set()
        while style is not None and id(style) not in visited:
            visited.add(id(style))
            try:
                pf = style.paragraph_format
            except Exception:
                pf = None
            if pf is not None and pf.alignment is not None:
                return pf.alignment
            try:
                style = style.base_style
            except Exception:
                break
        return None

    def _text_runs(self, paragraph):
        """返回段落中有可见文字的 run（跳过纯空白/纯制表符）。"""
        return [r for r in paragraph.runs if r.text and r.text.strip()]

    def _is_center(self, paragraph):
        from docx.enum.text import WD_ALIGN_PARAGRAPH
        return self._resolve_alignment(paragraph) == WD_ALIGN_PARAGRAPH.CENTER

    def _is_font_hei(self, font):
        return bool(font) and ('黑体' in font or 'SimHei' in font)

    def _is_font_song(self, font):
        return bool(font) and ('宋体' in font or 'SimSun' in font)

    def _para_has_page_break_before(self, paragraph):
        pPr = paragraph._element.pPr
        if pPr is None:
            return False
        pbb = pPr.find(qn('w:pageBreakBefore'))
        if pbb is None:
            return False
        return pbb.get(qn('w:val')) in (None, '1', 'on', 'true')

    def _para_has_manual_break(self, paragraph):
        x = paragraph._element.xml
        return '<w:br' in x and 'w:type="page"' in x

    def _starts_new_page(self, idx):
        """判断扁平段落 idx 是否另起一页（段前分页 / 其前有分页符或分节）。"""
        if idx < 0 or idx >= len(self._flat_paragraphs):
            return False
        p = self._flat_paragraphs[idx]
        if self._para_has_page_break_before(p):
            return True
        # 其前一个非空段落带手动分页符，或它与前段分属不同节
        for j in range(idx - 1, -1, -1):
            if self._para_has_manual_break(self._flat_paragraphs[j]):
                return True
            if self._flat_paragraphs[j].text.strip():
                break
        if idx > 0 and self._para_section(idx) != self._para_section(idx - 1):
            return True
        return False

    def _line_spacing_is_1_5(self, paragraph):
        """1.5 倍行距：line_spacing_rule 为 ONE_POINT_FIVE，或 line=360 且 auto。"""
        from docx.enum.text import WD_LINE_SPACING
        pf = paragraph.paragraph_format
        try:
            if pf.line_spacing_rule == WD_LINE_SPACING.ONE_POINT_FIVE:
                return True
            if pf.line_spacing_rule == WD_LINE_SPACING.MULTIPLE and pf.line_spacing:
                return abs(float(pf.line_spacing) - 1.5) < 0.05
        except Exception:
            pass
        try:
            if pf.line_spacing and abs(float(pf.line_spacing) - 1.5) < 0.05:
                return True
        except Exception:
            pass
        return False

    # ---- 记分工具 ----
    # 说明：维度二各检查方法按 _DIM2_RULES 的顺序依次调用，且每个检查方法
    # 恰好调用下列四个记分函数之一，因此 dim2_outcomes 会与 _DIM2_RULES 一一对应。
    def _award(self, score, msg):
        self.results['dimension2_scores'].append(score)
        self.results['hit_points'].append(f"+{score}：{msg}")
        self.results['dim2_outcomes'].append({'hit': True, 'delta': score, 'detail': msg})
        print(f"[+{score}] {msg}")

    def _miss(self, msg):
        self.results['dim2_outcomes'].append({'hit': False, 'delta': 0, 'detail': msg})
        print(f"[+0] 未得分：{msg}")

    def _deduct(self, score, msg):
        self.results['dimension2_deductions'].append(score)
        self.results['hit_points'].append(f"-{score}：{msg}")
        self.results['dim2_outcomes'].append({'hit': True, 'delta': -score, 'detail': msg})
        print(f"[-{score}] {msg}")

    def _no_deduct(self, msg):
        self.results['dim2_outcomes'].append({'hit': False, 'delta': 0, 'detail': msg})
        print(f"[PASS] {msg}")

    # ================= 得分点：封面页 =================
    def _check_cover_structure(self):
        """+5：第1页为封面页，有"成教本科毕业论文"一行，其下依次排列
        题目/学号/姓名/专业/指导老师五个字段行。"""
        msg = ('第1页为封面页，有"成教本科毕业论文"一行，其下依次排列'
               '"题目""学号""姓名""专业""指导老师"五个字段行')
        if self.cover_title_idx < 0:
            self._miss(msg)
            return
        # 封面标题应在文档很靠前（第1页）
        if self._para_section(self.cover_title_idx) != 0 or self.cover_title_idx > 8:
            self._miss(msg + "（未出现在首页顶部）")
            return
        fields = ['题目', '题　目', '学号', '学　号', '姓名', '姓　名',
                  '专业', '专　业', '指导老师', '指导教师']
        # 按顺序在标题之后查找五个字段
        wanted = [('题', '目'), ('学', '号'), ('姓', '名'), ('专', '业'), ('指导',)]
        seq_ok = True
        pos = self.cover_title_idx + 1
        limit = self.cover_title_idx + 40
        for grp in wanted:
            found = -1
            for i in range(pos, min(limit, len(self._flat_paragraphs))):
                norm = re.sub(r'\s+', '', self._flat_paragraphs[i].text)
                if all(k in norm for k in grp) and len(norm) <= 12:
                    found = i
                    break
            if found < 0:
                seq_ok = False
                break
            pos = found + 1
        if seq_ok:
            self._award(5, msg)
        else:
            self._miss(msg + "（五个字段行未按序齐备）")

    def _check_cover_title_font(self):
        """+1：封面页"成教本科毕业论文"为黑体、小初、加粗、居中。"""
        msg = '封面页"成教本科毕业论文"一行为黑体、小初、加粗、居中'
        if self.cover_title_idx < 0:
            self._miss(msg)
            return
        p = self._flat_paragraphs[self.cover_title_idx]
        runs = self._text_runs(p)
        if not runs:
            self._miss(msg)
            return
        ok = self._is_center(p)
        for r in runs:
            font = self._run_east_asia_font(r, p)
            size = self._run_size_pt(r, p)
            bold = self._run_bold(r, p)
            if not self._is_font_hei(font):
                ok = False
            if size is None or abs(size - PT_XIAOCHU) > 0.6:
                ok = False
            if not bold:
                ok = False
        if ok:
            self._award(1, msg)
        else:
            self._miss(msg)

    def _check_cover_images(self):
        """+3：封面页"成教本科毕业论文"上方居中排有两张图片，
        尺寸分别为 2.8×2.8 与 11.5×2.9 厘米，误差 0.3 厘米内。"""
        msg = ('封面页"成教本科毕业论文"上方居中排有两张图片，尺寸分别为'
               '2.8cm×2.8cm 与 11.5cm×2.9cm（误差0.3cm内）')
        try:
            with zipfile.ZipFile(self.open_path) as zf:
                doc = zf.read('word/document.xml').decode('utf-8')
        except Exception:
            self._miss(msg)
            return
        # 收集所有图片 extent（cx,cy 单位 EMU→cm）
        sizes = []
        for m in re.finditer(r'<wp:extent cx="(\d+)" cy="(\d+)"', doc):
            cx = int(m.group(1)) / EMU_PER_CM
            cy = int(m.group(2)) / EMU_PER_CM
            sizes.append((cx, cy))

        def match(sz, tw, th, tol=0.3):
            return abs(sz[0] - tw) <= tol and abs(sz[1] - th) <= tol

        has_small = any(match(s, 2.8, 2.8) for s in sizes)
        has_big = any(match(s, 11.5, 2.9) for s in sizes)
        if has_small and has_big:
            self._award(3, msg)
        else:
            self._miss(msg + f"（检测到图片尺寸：{[(round(a,2),round(b,2)) for a,b in sizes]}）")

    def _check_cover_field_rows_font(self):
        """+3：五个字段行均为黑体、三号、加粗、居中；字段值处带下划线，
        横线长度约 8~15 字符宽，五行大体一致。"""
        msg = ('封面页五个字段行均为黑体、三号、加粗、居中，字段值处带下划线，'
               '横线长度约8~15字符宽且五行大体一致')
        if self.cover_title_idx < 0:
            self._miss(msg)
            return
        # 定位五个字段行段落
        wanted = [('题', '目'), ('学', '号'), ('姓', '名'), ('专', '业'), ('指导',)]
        rows = []
        pos = self.cover_title_idx + 1
        limit = self.cover_title_idx + 40
        for grp in wanted:
            found = -1
            for i in range(pos, min(limit, len(self._flat_paragraphs))):
                norm = re.sub(r'\s+', '', self._flat_paragraphs[i].text)
                if all(k in norm for k in grp) and len(norm) <= 12:
                    found = i
                    break
            if found < 0:
                break
            rows.append(found)
            pos = found + 1
        if len(rows) != 5:
            self._miss(msg + "（未齐备五个字段行）")
            return

        default_tab = self._default_tab_stop_cm()
        ok = True
        underline_widths = []
        underline_groups = []
        for idx in rows:
            p = self._flat_paragraphs[idx]
            if not self._is_center(p):
                ok = False
            label_runs = self._text_runs(p)
            for r in label_runs:
                font = self._run_east_asia_font(r, p)
                size = self._run_size_pt(r, p)
                bold = self._run_bold(r, p)
                if not self._is_font_hei(font):
                    ok = False
                if size is None or abs(size - PT_SANHAO) > 0.6:
                    ok = False
                if not bold:
                    ok = False
            # 字段值处带下划线：统计下划线宽度与分组
            w = self._underline_char_width(p, default_tab)
            underline_widths.append(w)
            underline_groups.append(self._count_underline_groups(p))
        # 字段值处带下划线：每行都需存在下划线内容
        if any(g < 1 or w <= 0 for g, w in zip(underline_groups, underline_widths)):
            ok = False
        # 横线长度大体一致：五行下划线宽度极差不超过 3 字符
        # （制表符渲染的下划线精确宽度依赖排版引擎，这里以"存在且五行一致"
        #  近似满足"横线长度约8~15字符宽、五行大体一致"的评估意图）
        if underline_widths and (max(underline_widths) - min(underline_widths) > 3):
            ok = False
        if ok:
            self._award(3, msg)
        else:
            self._miss(msg + f"（下划线宽度估算：{[round(w,1) for w in underline_widths]}字符）")

    def _default_tab_stop_cm(self):
        try:
            with zipfile.ZipFile(self.open_path) as zf:
                st = zf.read('word/settings.xml').decode('utf-8')
            m = re.search(r'<w:defaultTabStop w:val="(\d+)"', st)
            if m:
                return int(m.group(1)) / 1440.0 * 2.54
        except Exception:
            pass
        return 0.74  # Word 默认 0.74cm

    def _underline_char_width(self, paragraph, default_tab_cm):
        """估算段落中带下划线内容的横向宽度（换算为中文字符数，1字符≈五号字宽0.37cm）。

        字段值常用制表符+下划线表达：统计带下划线 run 的宽度。
          - 制表符：按段落自定义制表位间距，退回默认制表位宽度
          - 普通字符：中文按约 0.37cm/字（10.5pt），ASCII 折半
        """
        # 段落自定义左对齐制表位位置（cm），用于估算 tab 推进宽度
        tab_positions = []
        try:
            for t in paragraph.paragraph_format.tab_stops:
                if t.position is not None:
                    tab_positions.append(t.position.cm)
        except Exception:
            pass
        tab_positions = sorted(tab_positions)

        char_cm = 0.37  # 五号字全角字符宽
        total_cm = 0.0
        # 追踪当前水平位置以估算 tab 推进（仅在有自定义 tab 时用）
        cursor = 0.0
        for r in paragraph.runs:
            u = bool(r.underline)
            for ch in r.text:
                if ch == '\t':
                    # tab 推进到下一制表位
                    nxt = None
                    for tp in tab_positions:
                        if tp > cursor + 1e-6:
                            nxt = tp
                            break
                    if nxt is None:
                        # 退回默认制表位网格
                        step = default_tab_cm if default_tab_cm > 0 else 0.74
                        advance = step - (cursor % step if step else 0)
                        if advance <= 1e-6:
                            advance = step
                    else:
                        advance = nxt - cursor
                    cursor += advance
                    if u:
                        total_cm += advance
                else:
                    w = char_cm if ord(ch) > 0x2E80 else char_cm * 0.5
                    cursor += w
                    if u:
                        total_cm += w
        return total_cm / char_cm

    def _check_cover_date(self):
        """+1：封面页底部有一行日期（如"2026 年 3 月"），黑体、四号、加粗、居中。"""
        msg = '封面页底部有一行日期（如"2026年3月"），黑体、四号、加粗、居中'
        if self.cover_title_idx < 0:
            self._miss(msg)
            return
        end = self.decl_title_idx if self.decl_title_idx > 0 else len(self._flat_paragraphs)
        date_re = re.compile(r'(19|20)\d{2}\s*年\s*\d{1,2}\s*月')
        found = None
        for i in range(self.cover_title_idx + 1, min(end, len(self._flat_paragraphs))):
            if date_re.search(re.sub(r'\s+', ' ', self._flat_paragraphs[i].text)):
                found = i
        if found is None:
            self._miss(msg)
            return
        p = self._flat_paragraphs[found]
        ok = self._is_center(p)
        for r in self._text_runs(p):
            font = self._run_east_asia_font(r, p)
            size = self._run_size_pt(r, p)
            bold = self._run_bold(r, p)
            # 日期中数字常为西文字体，中文字（年/月）应为黑体，字号四号
            if re.search(r'[年月]', r.text) and not self._is_font_hei(font):
                ok = False
            if size is None or abs(size - PT_SIHAO) > 0.6:
                ok = False
            if not bold:
                ok = False
        if ok:
            self._award(1, msg)
        else:
            self._miss(msg)

    def _check_cover_single_page(self):
        """+1：封面页单独占一页。"""
        msg = '封面页单独占一页'
        if self.cover_title_idx < 0 or self.decl_title_idx < 0:
            self._miss(msg)
            return
        # 声明页另起一页（分节或分页），且封面与声明之间无正文实质内容
        if self._starts_new_page(self.decl_title_idx):
            self._award(1, msg)
        else:
            self._miss(msg + "（声明页未与封面分页）")

    # ================= 得分点：原创性声明页 =================
    def _check_decl_page(self):
        """+5：封面页之后有独立的原创性声明页，标题为"原创性声明及使用授权声明"。"""
        msg = '封面页之后有独立的原创性声明页，标题为"原创性声明及使用授权声明"'
        if self.decl_title_idx < 0:
            self._miss(msg)
            return
        norm = re.sub(r'\s+', '', self._flat_paragraphs[self.decl_title_idx].text)
        # 独立：位于封面之后，且另起一页
        if (self.cover_title_idx < self.decl_title_idx
                and norm == '原创性声明及使用授权声明'
                and self._starts_new_page(self.decl_title_idx)):
            self._award(5, msg)
        else:
            self._miss(msg)

    def _check_decl_title_font(self):
        """+1：原创性声明页标题行为宋体、二号、加粗、居中。"""
        msg = '原创性声明页标题行为宋体、二号、加粗、居中'
        if self.decl_title_idx < 0:
            self._miss(msg)
            return
        p = self._flat_paragraphs[self.decl_title_idx]
        ok = self._is_center(p)
        for r in self._text_runs(p):
            font = self._run_east_asia_font(r, p)
            size = self._run_size_pt(r, p)
            bold = self._run_bold(r, p)
            if not self._is_font_song(font):
                ok = False
            if size is None or abs(size - PT_ERHAO) > 0.6:
                ok = False
            if not bold:
                ok = False
        if ok:
            self._award(1, msg)
        else:
            self._miss(msg)

    def _decl_body_paragraphs(self):
        """返回声明页三段声明正文段落（以三种约定开头之一起始）。"""
        if self.decl_title_idx < 0:
            return []
        starts = ['本人郑重声明', '本人了解学校', '本人同意学校保存']
        end = self.toc_title_idx if self.toc_title_idx > 0 else len(self._flat_paragraphs)
        found = []
        for prefix in starts:
            for i in range(self.decl_title_idx + 1, min(end, len(self._flat_paragraphs))):
                norm = re.sub(r'\s+', '', self._flat_paragraphs[i].text)
                if norm.startswith(prefix):
                    found.append(i)
                    break
        return found

    def _check_decl_paragraphs(self):
        """+3：声明页有三段声明正文，分别以三种约定开头。"""
        msg = ('原创性声明页有三段声明正文，分别以"本人郑重声明……"'
               '"本人了解学校……""本人同意学校保存……"开头')
        found = self._decl_body_paragraphs()
        if len(found) == 3:
            self._award(3, msg)
        else:
            self._miss(msg + f"（命中开头段落数：{len(found)}）")

    def _check_decl_paragraphs_font(self):
        """+1：三段声明正文均为宋体、四号、1.5倍行距、首行缩进两字符、两端对齐。"""
        from docx.enum.text import WD_ALIGN_PARAGRAPH
        msg = ('原创性声明页三段声明正文均为宋体、四号、1.5倍行距、'
               '首行缩进两字符、两端对齐')
        found = self._decl_body_paragraphs()
        if len(found) != 3:
            self._miss(msg)
            return
        ok = True
        for idx in found:
            p = self._flat_paragraphs[idx]
            for r in self._text_runs(p):
                font = self._run_east_asia_font(r, p)
                size = self._run_size_pt(r, p)
                if not self._is_font_song(font):
                    ok = False
                if size is None or abs(size - PT_SIHAO) > 0.6:
                    ok = False
            if not self._line_spacing_is_1_5(p):
                ok = False
            if not self._first_line_indent_two_chars(p):
                ok = False
            if self._resolve_alignment(p) != WD_ALIGN_PARAGRAPH.JUSTIFY:
                ok = False
        if ok:
            self._award(1, msg)
        else:
            self._miss(msg)

    def _first_line_indent_two_chars(self, paragraph):
        """首行缩进两字符：优先看 firstLineChars=200；否则按四号≈14pt判断约2字符宽。"""
        pPr = paragraph._element.pPr
        if pPr is not None:
            ind = pPr.find(qn('w:ind'))
            if ind is not None:
                fchars = ind.get(qn('w:firstLineChars'))
                if fchars is not None:
                    try:
                        return abs(int(fchars) - 200) <= 30
                    except ValueError:
                        pass
                first = ind.get(qn('w:firstLine'))
                if first is not None:
                    try:
                        twips = int(first)
                    except ValueError:
                        twips = 0
                    # 段落首个可见run的字号，两字符宽 = 2*字号
                    runs = self._text_runs(paragraph)
                    size = self._run_size_pt(runs[0], paragraph) if runs else 12.0
                    size = size or 12.0
                    two_char_twips = 2 * size * 20  # 1pt=20twips
                    return abs(twips - two_char_twips) <= 0.6 * size * 20
        # 通过 python-docx 属性兜底
        pf = paragraph.paragraph_format
        try:
            if pf.first_line_indent is not None:
                emu = pf.first_line_indent
                runs = self._text_runs(paragraph)
                size = self._run_size_pt(runs[0], paragraph) if runs else 12.0
                size = size or 12.0
                two_char_emu = 2 * size * EMU_PER_PT
                return abs(int(emu) - two_char_emu) <= 0.6 * size * EMU_PER_PT
        except Exception:
            pass
        return False

    def _check_decl_signature(self):
        """+1：三段正文之下有一行"论文作者签名：__ 日 期：__"，宋体、小四，
        签名与日期后各带一段下划线，横线长度约 4~10 字符宽。"""
        msg = ('原创性声明页三段正文之下有"论文作者签名：__ 日 期：__"一行，'
               '宋体、小四，签名与日期后各带下划线，横线约4~10字符宽')
        found = self._decl_body_paragraphs()
        if not found:
            self._miss(msg)
            return
        end = self.toc_title_idx if self.toc_title_idx > 0 else len(self._flat_paragraphs)
        sign_idx = -1
        for i in range(max(found) + 1, min(end, len(self._flat_paragraphs))):
            norm = re.sub(r'\s+', '', self._flat_paragraphs[i].text)
            if '论文作者签名' in norm and ('日期' in norm or '日　期' in self._flat_paragraphs[i].text):
                sign_idx = i
                break
            if '作者签名' in norm and '日期' in norm:
                sign_idx = i
                break
        if sign_idx < 0:
            self._miss(msg)
            return
        p = self._flat_paragraphs[sign_idx]
        ok = True
        for r in self._text_runs(p):
            font = self._run_east_asia_font(r, p)
            size = self._run_size_pt(r, p)
            if not self._is_font_song(font):
                ok = False
            if size is None or abs(size - PT_XIAOSI) > 0.6:
                ok = False
        # 下划线：至少两处带下划线内容（签名后、日期后）
        default_tab = self._default_tab_stop_cm()
        underline_groups = self._count_underline_groups(p)
        total_w = self._underline_char_width(p, default_tab)
        if underline_groups < 2:
            ok = False
        # 两段下划线合计约 8~20 字符（每段 4~10）
        if not (6 <= total_w <= 24):
            ok = False
        if ok:
            self._award(1, msg)
        else:
            self._miss(msg + f"（下划线段数={underline_groups}, 合计宽≈{round(total_w,1)}字符）")

    def _count_underline_groups(self, paragraph):
        """统计段落中"连续带下划线"的分组数量。"""
        groups = 0
        prev = False
        for r in paragraph.runs:
            cur = bool(r.underline) and len(r.text) > 0
            if cur and not prev:
                groups += 1
            prev = cur
        return groups

    # ================= 得分点：目录页 =================
    def _check_toc_page(self):
        """+5：声明页之后、正文之前有独立目录页，"目录"二字为宋体、二号、居中。"""
        msg = '原创性声明页之后、正文之前有独立的目录页，"目录"二字为宋体、二号、居中'
        if self.toc_title_idx < 0:
            self._miss(msg)
            return
        # 顺序：声明 < 目录 < 正文标题
        order_ok = (self.decl_title_idx < self.toc_title_idx
                    and (self.body_title_idx < 0 or self.toc_title_idx < self.body_title_idx))
        p = self._flat_paragraphs[self.toc_title_idx]
        ok = order_ok and self._is_center(p) and self._starts_new_page(self.toc_title_idx)
        for r in self._text_runs(p):
            font = self._run_east_asia_font(r, p)
            size = self._run_size_pt(r, p)
            if not self._is_font_song(font):
                ok = False
            if size is None or abs(size - PT_ERHAO) > 0.6:
                ok = False
        if ok:
            self._award(5, msg)
        else:
            self._miss(msg)

    def _check_toc_auto(self):
        """+3：目录页为自动生成目录，条目按标题层级缩进，页码右对齐。"""
        msg = '目录页为自动生成目录，条目按标题层级缩进，页码右对齐'
        try:
            with zipfile.ZipFile(self.open_path) as zf:
                doc = zf.read('word/document.xml').decode('utf-8')
        except Exception:
            self._miss(msg)
            return
        has_toc_field = bool(re.search(r'<w:instrText[^>]*>[^<]*TOC[^<]*</w:instrText>', doc)) \
            or 'PAGEREF' in doc
        # 层级缩进：目录条目段落使用不同样式(TOC1/TOC2)或不同缩进
        has_levels = bool(re.search(r'w:val="(TOC|toc)\s*\d', doc)) or doc.count('PAGEREF') >= 2
        # 页码右对齐：目录使用右对齐制表位
        has_right_tab = 'w:val="right"' in doc
        if has_toc_field and (has_levels or has_right_tab) and has_right_tab:
            self._award(3, msg)
        else:
            self._miss(msg + f"（TOC域={has_toc_field}, 右对齐制表位={has_right_tab}）")

    # ================= 正文段落 / 标题识别 =================
    def _body_start_idx(self):
        """正文起始扁平段落索引：正文标题所在段（无则目录后第一段）。"""
        if self.body_title_idx >= 0:
            return self.body_title_idx
        if self.toc_title_idx >= 0:
            return self.toc_title_idx + 1
        return 0

    def _is_h1(self, paragraph):
        """一级标题：样式为 Heading 1/标题1，或大纲级别0，或中文数字序号章标题、
        结论/参考文献/致谢。"""
        try:
            name = (paragraph.style.name or '').lower()
        except Exception:
            name = ''
        if name in ('heading 1', '标题 1') or name.replace(' ', '') in ('heading1', '标题1'):
            return True
        text = paragraph.text.strip()
        norm = re.sub(r'\s+', '', text)
        if not norm:
            return False
        if '\t' in paragraph.text or '…' in text:
            return False
        if norm in ('结论', '结束语', '结语', '参考文献', '致谢', '致　谢'):
            return True
        # "一 绪论" 中文数字 + 空格/顿号 + 标题
        if re.match(r'^[一二三四五六七八九十百]+[\s、][^\s]', text):
            return True
        return False

    def _is_h2(self, paragraph):
        """二级标题：样式为 Heading 2/标题2，或全角括号中文数字序号开头。"""
        try:
            name = (paragraph.style.name or '').lower()
        except Exception:
            name = ''
        if name in ('heading 2', '标题 2') or name.replace(' ', '') in ('heading2', '标题2'):
            return True
        text = paragraph.text.strip()
        # （一）xxx 或 (一)xxx
        if re.match(r'^[（(][一二三四五六七八九十]+[）)]', text):
            return True
        return False

    def _is_heading(self, paragraph):
        return self._is_h1(paragraph) or self._is_h2(paragraph)

    def _is_body_text_paragraph(self, idx):
        """正文正文段落：位于正文区、非空、非标题、非摘要/关键词标签行、
        非图注表注、非目录条目。"""
        if idx < self._body_start_idx():
            return False
        p = self._flat_paragraphs[idx]
        text = p.text.strip()
        if not text:
            return False
        if '\t' in p.text or '…' in text:
            return False
        if self._is_heading(p):
            return False
        if idx == self.abstract_idx or idx == self.keywords_idx:
            return False
        if idx == self.body_title_idx:
            return False
        norm = re.sub(r'\s+', '', text)
        if norm.startswith('摘要') or norm.startswith('关键词') or norm.lower().startswith('abstract') or norm.lower().startswith('keywords'):
            return False
        if re.match(r'^图\s*\d', text) or re.match(r'^表\s*\d', text):
            return False
        # 参考文献条目单独由专门规则判定，这里排除以免误判正文格式
        if re.match(r'^[\[【]\s*\d+\s*[\]】]', text):
            return False
        return True

    def _body_text_indices(self):
        return [i for i in range(self._body_start_idx(), len(self._flat_paragraphs))
                if self._is_body_text_paragraph(i)]

    def _h1_indices(self):
        return [i for i in range(self._body_start_idx(), len(self._flat_paragraphs))
                if self._is_h1(self._flat_paragraphs[i])]

    def _h2_indices(self):
        return [i for i in range(self._body_start_idx(), len(self._flat_paragraphs))
                if self._is_h2(self._flat_paragraphs[i]) and not self._is_h1(self._flat_paragraphs[i])]

    # ================= 得分点：正文首页 =================
    def _check_body_title(self):
        """+3：正文首页顶部有论文标题一行，小二、加粗、居中，标题前不带"标题："前缀。"""
        msg = ('正文首页顶部有论文标题一行"小微企业财务数据管理规范化问题及改进路径研究"，'
               '小二、加粗、居中，标题前不带"标题："等前缀')
        if self.body_title_idx < 0:
            self._miss(msg)
            return
        p = self._flat_paragraphs[self.body_title_idx]
        text = p.text.strip()
        ok = self._is_center(p)
        if text.startswith('标题') or text.startswith('题目') or '：' in text[:4]:
            ok = False
        for r in self._text_runs(p):
            size = self._run_size_pt(r, p)
            bold = self._run_bold(r, p)
            if size is None or abs(size - PT_XIAOER) > 0.6:
                ok = False
            if not bold:
                ok = False
        if ok:
            self._award(3, msg)
        else:
            self._miss(msg + f"（标题="\
                       f"{text[:20]!r}）")

    def _check_abstract(self):
        """+3：摘要为一整段，以"摘 要："起头直接接内容，"摘 要："为黑体、四号、加粗，
        其后内容为宋体、小四。"""
        msg = ('正文首页摘要为一整段，以"摘　要："起头直接接内容，'
               '"摘　要："为黑体、四号、加粗，其后内容为宋体、小四')
        if self.abstract_idx < 0:
            self._miss(msg)
            return
        p = self._flat_paragraphs[self.abstract_idx]
        runs = self._text_runs(p)
        if not runs:
            self._miss(msg)
            return
        norm = re.sub(r'\s+', '', p.text)
        # 一整段且标签后直接接内容
        if not norm.startswith('摘要') or len(norm) < 12:
            self._miss(msg + "（摘要标签后无内容或非整段）")
            return
        ok = True
        # 首个 run 通常是"摘 要："标签
        label_done = False
        for r in runs:
            font = self._run_east_asia_font(r, p)
            size = self._run_size_pt(r, p)
            bold = self._run_bold(r, p)
            rnorm = re.sub(r'\s+', '', r.text)
            if not label_done and '摘要' in rnorm:
                # 标签：黑体四号加粗
                if not self._is_font_hei(font):
                    ok = False
                if size is None or abs(size - PT_SIHAO) > 0.6:
                    ok = False
                if not bold:
                    ok = False
                label_done = True
            else:
                # 内容：宋体小四
                if not self._is_font_song(font):
                    ok = False
                if size is None or abs(size - PT_XIAOSI) > 0.6:
                    ok = False
        if ok:
            self._award(3, msg)
        else:
            self._miss(msg)

    def _check_keywords(self):
        """+1：关键词为一段，紧接摘要段之后，以"关键词："起头，该三字为黑体、四号、加粗，
        其后内容为宋体、小四。"""
        msg = ('正文首页关键词为一段，紧接摘要段之后，以"关键词："起头，'
               '该三字为黑体、四号、加粗，其后内容为宋体、小四')
        if self.keywords_idx < 0 or self.abstract_idx < 0:
            self._miss(msg)
            return
        # 紧接摘要段之后（允许中间没有实质段落）
        adjacent = self.keywords_idx > self.abstract_idx
        for j in range(self.abstract_idx + 1, self.keywords_idx):
            if self._flat_paragraphs[j].text.strip():
                adjacent = False
                break
        p = self._flat_paragraphs[self.keywords_idx]
        ok = adjacent
        label_done = False
        for r in self._text_runs(p):
            font = self._run_east_asia_font(r, p)
            size = self._run_size_pt(r, p)
            bold = self._run_bold(r, p)
            rnorm = re.sub(r'\s+', '', r.text)
            if not label_done and '关键词' in rnorm:
                if not self._is_font_hei(font):
                    ok = False
                if size is None or abs(size - PT_SIHAO) > 0.6:
                    ok = False
                if not bold:
                    ok = False
                label_done = True
            else:
                if not self._is_font_song(font):
                    ok = False
                if size is None or abs(size - PT_XIAOSI) > 0.6:
                    ok = False
        if ok:
            self._award(1, msg)
        else:
            self._miss(msg)

    # ================= 得分点：正文段落格式 =================
    def _check_body_font(self):
        """+5：除标题/摘要关键词标签/图注表注/表格文字外的所有正文段落均为宋体、小四。"""
        msg = '正文各页中，除标题、摘要关键词标签、图注表注、表格文字外的所有正文段落均为宋体、小四'
        idxs = self._body_text_indices()
        if not idxs:
            self._miss(msg + "（未识别到正文段落）")
            return
        bad = None
        for i in idxs:
            p = self._flat_paragraphs[i]
            for r in self._text_runs(p):
                font = self._run_east_asia_font(r, p)
                size = self._run_size_pt(r, p)
                # 允许西文数字/字母使用西文字体，只校验中文字体为宋体
                has_cjk = any(ord(c) > 0x2E80 for c in r.text)
                if has_cjk and not self._is_font_song(font):
                    bad = (i, f"字体{font}", p.text[:16])
                    break
                if size is None or abs(size - PT_XIAOSI) > 0.6:
                    bad = (i, f"字号{size}", p.text[:16])
                    break
            if bad:
                break
        if bad is None:
            self._award(5, msg)
        else:
            self._miss(msg + f"（段落「{bad[2]}」{bad[1]}）")

    def _check_body_first_line_indent(self):
        """+3：所有正文段落首行缩进两字符。"""
        msg = '正文各页中，所有正文段落首行缩进两字符'
        idxs = self._body_text_indices()
        if not idxs:
            self._miss(msg)
            return
        bad = None
        for i in idxs:
            if not self._first_line_indent_two_chars(self._flat_paragraphs[i]):
                bad = self._flat_paragraphs[i].text[:16]
                break
        if bad is None:
            self._award(3, msg)
        else:
            self._miss(msg + f"（段落「{bad}」首行缩进不符）")

    def _check_body_line_spacing(self):
        """+3：所有正文段落为 1.5 倍行距。"""
        msg = '正文各页中，所有正文段落为1.5倍行距'
        idxs = self._body_text_indices()
        if not idxs:
            self._miss(msg)
            return
        bad = None
        for i in idxs:
            if not self._line_spacing_is_1_5(self._flat_paragraphs[i]):
                bad = self._flat_paragraphs[i].text[:16]
                break
        if bad is None:
            self._award(3, msg)
        else:
            self._miss(msg + f"（段落「{bad}」行距不符）")

    def _check_body_justify(self):
        """+3：所有正文段落两端对齐。"""
        from docx.enum.text import WD_ALIGN_PARAGRAPH
        msg = '正文各页中，所有正文段落两端对齐'
        idxs = self._body_text_indices()
        if not idxs:
            self._miss(msg)
            return
        bad = None
        for i in idxs:
            align = self._resolve_alignment(self._flat_paragraphs[i])
            # 两端对齐；未显式设置但样式为两端对齐(Normal jc=both)也算
            if align != WD_ALIGN_PARAGRAPH.JUSTIFY:
                bad = self._flat_paragraphs[i].text[:16]
                break
        if bad is None:
            self._award(3, msg)
        else:
            self._miss(msg + f"（段落「{bad}」对齐={bad}）")

    # ================= 得分点：标题样式 =================
    def _check_h1_style(self):
        """+3：所有一级标题段落样式为"标题 1"。"""
        msg = '正文各页中，所有一级标题段落样式为"标题 1"'
        idxs = self._h1_indices()
        if not idxs:
            self._miss(msg)
            return
        bad = None
        for i in idxs:
            name = (self._flat_paragraphs[i].style.name or '').lower().replace(' ', '')
            if name not in ('heading1', '标题1'):
                bad = self._flat_paragraphs[i].text[:16]
                break
        if bad is None:
            self._award(3, msg)
        else:
            self._miss(msg + f"（一级标题「{bad}」样式不是标题1）")

    def _check_h1_numbering(self):
        """+3：带序号一级标题序号为"中文数字＋空格＋标题"，如"一 绪论"，不带顿号，
        从"一"起连续不跳号；结论/参考文献/致谢不要求带序号。"""
        msg = ('正文各页中，带序号的一级标题序号样式为"中文数字＋空格＋标题"'
               '（如"一 绪论"，不带顿号，从"一"起连续不跳号；结论/参考文献/致谢除外）')
        idxs = self._h1_indices()
        if not idxs:
            self._miss(msg)
            return
        cn = '一二三四五六七八九十'
        exempt = ('结论', '结束语', '结语', '参考文献', '致谢')
        expected = 0
        ok = True
        for i in idxs:
            text = self._flat_paragraphs[i].text.strip()
            norm = re.sub(r'\s+', '', text)
            if norm in exempt:
                continue
            # 不带顿号：序号后应为空格而非顿号
            if re.match(r'^[一二三四五六七八九十]+、', text):
                ok = False
                break
            m = re.match(r'^([一二三四五六七八九十]+)\s+\S', text)
            if not m:
                ok = False
                break
            num = m.group(1)
            # 连续不跳号：第 expected+1 个中文数字
            if len(num) == 1:
                val = cn.index(num) + 1 if num in cn else -1
            else:
                val = -1  # 十以上暂不精确，容忍
            if val != -1:
                if val != expected + 1:
                    ok = False
                    break
                expected = val
            else:
                expected += 1
        if ok and expected >= 1:
            self._award(3, msg)
        else:
            self._miss(msg)

    def _check_h1_font(self):
        """+3：所有一级标题均为黑体、三号、居中。"""
        msg = '正文各页中，所有一级标题均为黑体、三号、居中'
        idxs = self._h1_indices()
        if not idxs:
            self._miss(msg)
            return
        bad = None
        for i in idxs:
            p = self._flat_paragraphs[i]
            if not self._is_center(p):
                bad = (p.text[:16], '非居中')
                break
            for r in self._text_runs(p):
                font = self._run_east_asia_font(r, p)
                size = self._run_size_pt(r, p)
                if not self._is_font_hei(font):
                    bad = (p.text[:16], f'字体{font}')
                    break
                if size is None or abs(size - PT_SANHAO) > 0.6:
                    bad = (p.text[:16], f'字号{size}')
                    break
            if bad:
                break
        if bad is None:
            self._award(3, msg)
        else:
            self._miss(msg + f"（一级标题「{bad[0]}」{bad[1]}）")

    def _check_h2_style(self):
        """+3：所有二级标题段落样式为"标题 2"。"""
        msg = '正文各页中，所有二级标题段落样式为"标题 2"'
        idxs = self._h2_indices()
        if not idxs:
            self._miss(msg)
            return
        bad = None
        for i in idxs:
            name = (self._flat_paragraphs[i].style.name or '').lower().replace(' ', '')
            if name not in ('heading2', '标题2'):
                bad = self._flat_paragraphs[i].text[:16]
                break
        if bad is None:
            self._award(3, msg)
        else:
            self._miss(msg + f"（二级标题「{bad}」样式不是标题2）")

    def _check_h2_numbering(self):
        """+5：二级标题序号为全角括号中文数字，如"（一）研究背景"；
        写成"1.1"、半角"(一)"、"一."均不得分。"""
        msg = ('正文各页中，二级标题序号样式为全角括号中文数字（如"（一）研究背景"）；'
               '写成"1.1 研究背景"、半角"(一)"、"一."均不得分')
        idxs = self._h2_indices()
        if not idxs:
            self._miss(msg)
            return
        ok = True
        for i in idxs:
            text = self._flat_paragraphs[i].text.strip()
            # 必须为全角括号：（一）
            if not re.match(r'^（[一二三四五六七八九十]+）', text):
                ok = False
                break
        if ok:
            self._award(5, msg)
        else:
            self._miss(msg + f"（存在非全角括号中文数字序号，如「{self._flat_paragraphs[idxs[0]].text[:12]}」）")

    def _check_h2_font(self):
        """+3：所有二级标题均为黑体、四号、左对齐。"""
        from docx.enum.text import WD_ALIGN_PARAGRAPH
        msg = '正文各页中，所有二级标题均为黑体、四号、左对齐'
        idxs = self._h2_indices()
        if not idxs:
            self._miss(msg)
            return
        bad = None
        for i in idxs:
            p = self._flat_paragraphs[i]
            align = self._resolve_alignment(p)
            # 左对齐：None(默认左) 或 LEFT
            if align not in (None, WD_ALIGN_PARAGRAPH.LEFT):
                bad = (p.text[:16], f'对齐{align}')
                break
            for r in self._text_runs(p):
                font = self._run_east_asia_font(r, p)
                size = self._run_size_pt(r, p)
                if not self._is_font_hei(font):
                    bad = (p.text[:16], f'字体{font}')
                    break
                if size is None or abs(size - PT_SIHAO) > 0.6:
                    bad = (p.text[:16], f'字号{size}')
                    break
            if bad:
                break
        if bad is None:
            self._award(3, msg)
        else:
            self._miss(msg + f"（二级标题「{bad[0]}」{bad[1]}）")

    # ================= 得分点：页面/页脚 =================
    def _check_page_margins(self):
        """+3：全文各页页边距为上、下各 2.54cm，左、右各 3.0cm。"""
        msg = '全文各页页边距为上、下各2.54cm，左、右各3.0cm'
        ok = True
        detail = []
        for s in self.doc.sections:
            try:
                top = s.top_margin.cm
                bottom = s.bottom_margin.cm
                left = s.left_margin.cm
                right = s.right_margin.cm
            except Exception:
                ok = False
                break
            detail.append((round(top, 2), round(bottom, 2), round(left, 2), round(right, 2)))
            if not (abs(top - 2.54) <= 0.1 and abs(bottom - 2.54) <= 0.1
                    and abs(left - 3.0) <= 0.1 and abs(right - 3.0) <= 0.1):
                ok = False
        if ok and detail:
            self._award(3, msg)
        else:
            self._miss(msg + f"（各节边距cm：{detail}）")

    def _footer_xml(self, footer):
        """返回页脚（含文本框 txbxContent）的完整 XML 与所有 w:t 文本拼接。"""
        if footer is None:
            return '', ''
        try:
            xml = footer._element.xml
        except Exception:
            return '', ''
        texts = re.findall(r'<w:t[^>]*>([^<]*)</w:t>', xml)
        return xml, ''.join(texts)

    def _footer_has_page_field(self, footer):
        """页脚是否含 PAGE 域或"第 X 页"页码文本（穿透文本框）。"""
        xml, text = self._footer_xml(footer)
        if not xml:
            return False
        if re.search(r'<w:instrText[^>]*>[^<]*PAGE[^<]*</w:instrText>', xml):
            return True
        if re.search(r'<w:fldSimple[^>]*PAGE', xml):
            return True
        if re.search(r'第\s*\d*\s*页', re.sub(r'\s+', ' ', text)):
            return True
        return False

    def _footer_shows_di_x_ye(self, footer):
        """页脚是否显示"第 X 页"样式（穿透文本框；PAGE 域夹在"第""页"之间也算）。"""
        xml, text = self._footer_xml(footer)
        if not xml:
            return False
        norm = re.sub(r'\s+', ' ', text)
        if re.search(r'第\s*\d*\s*页', norm):
            return True
        if '第' in text and '页' in text and 'PAGE' in xml:
            return True
        return False

    def _effective_footer(self, section_idx, kind='default'):
        """沿链接回溯得到实际生效页脚对象。"""
        secs = list(self.doc.sections)
        if section_idx < 0 or section_idx >= len(secs):
            return None
        section = secs[section_idx]
        try:
            footer = section.footer
        except Exception:
            return None
        if footer is None:
            return None
        try:
            linked = footer.is_linked_to_previous
        except Exception:
            linked = False
        if linked and section_idx > 0:
            return self._effective_footer(section_idx - 1, kind)
        return footer

    def _body_content_section(self):
        """正文实质内容（章节/正文段落）所在节索引。

        正文首页标题/摘要可能仍在前置节，真正的正文章节（如"一 绪论"）常在
        新节。取正文区第一处一级标题（或第一段正文文字）所在节。
        """
        for i in range(self._body_start_idx(), len(self._flat_paragraphs)):
            p = self._flat_paragraphs[i]
            if self._is_h1(p) or self._is_body_text_paragraph(i):
                return self._para_section(i)
        return self._para_section(self._body_start_idx())

    def _check_footer_page_number(self):
        """+3：正文各页页脚均有自动页码，显示如"第 1 页"。"""
        msg = '正文各页页脚均有自动页码，显示如"第 1 页"'
        body_sec = self._body_content_section()
        secs = list(self.doc.sections)
        footer = None
        if 0 <= body_sec < len(secs):
            footer = self._effective_footer(body_sec)
        if footer is None or not self._footer_shows_di_x_ye(footer):
            # 兜底：任一节页脚含"第 X 页"
            footer = self._effective_footer(len(secs) - 1)
        if footer is not None and self._footer_shows_di_x_ye(footer):
            self._award(3, msg)
        else:
            self._miss(msg)

    def _check_footer_page_number_format(self):
        """+1：正文各页页脚页码居中，宋体，五号或小四。

        页码可能位于文本框（txbxContent）中，直接在页脚 XML 层面定位含
        "第X页"/PAGE 域的段落块，校验其居中与字体字号。
        """
        msg = '正文各页页脚页码居中，宋体，五号或小四'
        body_sec = self._body_content_section()
        secs = list(self.doc.sections)
        footer = self._effective_footer(body_sec if 0 <= body_sec < len(secs) else len(secs) - 1)
        if footer is None:
            self._miss(msg)
            return
        xml, _ = self._footer_xml(footer)
        if not xml:
            self._miss(msg)
            return
        # 逐个 <w:p> 段落块，找含页码（PAGE 域或"第""页"）者
        ok = False
        for pm in re.finditer(r'<w:p\b.*?</w:p>', xml, re.S):
            blk = pm.group(0)
            texts = ''.join(re.findall(r'<w:t[^>]*>([^<]*)</w:t>', blk))
            has_page = ('PAGE' in blk) or ('第' in texts and '页' in texts)
            if not has_page:
                continue
            centered = 'w:val="center"' in blk and '<w:jc' in blk
            # 中文字体：所有 eastAsia 均为宋体
            fonts = re.findall(r'<w:rFonts[^>]*w:eastAsia="([^"]*)"', blk)
            font_ok = all(self._is_font_song(f) for f in fonts) if fonts else True
            # 字号：出现的 w:sz 至少有一个是五号(21半磅)或小四(24半磅)，且不出现异常大字号
            szs = [int(s) / 2.0 for s in re.findall(r'<w:sz w:val="(\d+)"', blk)]
            size_ok = True
            if szs:
                size_ok = all(abs(s - PT_WUHAO) <= 0.6 or abs(s - PT_XIAOSI) <= 0.6 for s in szs)
            if centered and font_ok and size_ok:
                ok = True
                break
        if ok:
            self._award(1, msg)
        else:
            self._miss(msg)

    def _footer_has_page_field_para(self, paragraph):
        x = paragraph._element.xml
        if re.search(r'PAGE', x) and ('instrText' in x or 'fldSimple' in x):
            return True
        if re.search(r'第\s*\d*\s*页', paragraph.text):
            return True
        return False

    def _check_body_first_page_number(self):
        """+1：正文首页页脚页码显示"第 1 页"，其后逐页递增。
        通过节的 pgNumType start=1 与页脚 PAGE 域自动递增判定。"""
        msg = '正文首页页脚页码显示"第 1 页"，其后逐页递增'
        try:
            with zipfile.ZipFile(self.open_path) as zf:
                doc = zf.read('word/document.xml').decode('utf-8')
        except Exception:
            self._miss(msg)
            return
        # 正文节 pgNumType start=1，且页脚用 PAGE 域（自动逐页递增）
        has_start1 = bool(re.search(r'<w:pgNumType[^>]*w:start="1"', doc))
        body_sec = self._body_content_section()
        secs = list(self.doc.sections)
        footer = self._effective_footer(body_sec if 0 <= body_sec < len(secs) else len(secs) - 1)
        auto = footer is not None and self._footer_has_page_field(footer)
        if has_start1 and auto:
            self._award(1, msg)
        else:
            self._miss(msg + f"（起始页码=1:{has_start1}, 自动页码:{auto}）")

    def _check_sections_distinct(self):
        """+3：封面页、声明页、目录页与正文页分属不同节，正文首页页码从1重新起算。"""
        msg = '封面页、声明页、目录页与正文页分属不同节，正文首页页码从1重新起算'
        # 前置(封面/声明/目录)与正文分属不同节
        front_sec = self._para_section(self.toc_title_idx) if self.toc_title_idx >= 0 else 0
        body_sec = self._body_content_section()
        distinct = body_sec > front_sec
        try:
            with zipfile.ZipFile(self.open_path) as zf:
                doc = zf.read('word/document.xml').decode('utf-8')
        except Exception:
            doc = ''
        restart = bool(re.search(r'<w:pgNumType[^>]*w:start="1"', doc))
        if distinct and restart:
            self._award(3, msg)
        else:
            self._miss(msg + f"（前置节={front_sec}, 正文节={body_sec}, 重起页码={restart}）")

    def _check_tail_pages_new_page(self):
        """+1：结论页、参考文献页、致谢页三者各自另起一页。"""
        msg = '结论页、参考文献页、致谢页三者各自另起一页'
        targets = {'结论': -1, '参考文献': -1, '致谢': -1}
        for i in range(self._body_start_idx(), len(self._flat_paragraphs)):
            norm = re.sub(r'\s+', '', self._flat_paragraphs[i].text)
            for key in targets:
                if targets[key] < 0 and norm == key and self._is_h1(self._flat_paragraphs[i]):
                    targets[key] = i
        found = {k: v for k, v in targets.items() if v >= 0}
        if not found:
            self._miss(msg)
            return
        page_state = {k: self._starts_new_page(v) for k, v in found.items()}
        ok = all(page_state.values())
        if ok and len(found) >= 1:
            self._award(1, msg)
        else:
            self._miss(msg + f"（各标题另起页情况：{page_state}）")

    # ================= 得分点：参考文献 / 致谢 =================
    def _find_h1_by_names(self, names):
        """返回名称匹配的一级标题扁平索引（norm 完全等于给定名之一）。"""
        for i in range(self._body_start_idx(), len(self._flat_paragraphs)):
            norm = re.sub(r'\s+', '', self._flat_paragraphs[i].text)
            if norm in names and self._is_h1(self._flat_paragraphs[i]):
                return i
        return -1

    def _ref_title_idx(self):
        return self._find_h1_by_names(('参考文献',))

    def _ack_title_idx(self):
        return self._find_h1_by_names(('致谢', '致　谢'))

    def _ref_entry_indices(self):
        """参考文献条目段落：参考文献标题之后、致谢/文末之前的非空非标题段落。"""
        ref = self._ref_title_idx()
        if ref < 0:
            return []
        ack = self._ack_title_idx()
        end = ack if ack > ref else len(self._flat_paragraphs)
        out = []
        for i in range(ref + 1, end):
            if self._flat_paragraphs[i].text.strip():
                out.append(i)
        return out

    def _check_ref_ack_title_style(self):
        """+3：参考文献、致谢标题样式与一级标题一致（黑体、三号、居中，标题 1 样式），
        但不编入自动序号。"""
        msg = ('参考文献、致谢的标题样式与一级标题一致（黑体、三号、居中，'
               '"标题 1"样式），但不编入自动序号')
        ref = self._ref_title_idx()
        ack = self._ack_title_idx()
        targets = [x for x in (ref, ack) if x >= 0]
        if not targets:
            self._miss(msg)
            return
        ok = True
        for idx in targets:
            p = self._flat_paragraphs[idx]
            name = (p.style.name or '').lower().replace(' ', '')
            if name not in ('heading1', '标题1'):
                ok = False
            if not self._is_center(p):
                ok = False
            for r in self._text_runs(p):
                font = self._run_east_asia_font(r, p)
                size = self._run_size_pt(r, p)
                if not self._is_font_hei(font):
                    ok = False
                if size is None or abs(size - PT_SANHAO) > 0.6:
                    ok = False
            # 不编入自动序号：标题文本前不应有"一二三..．"章节数字
            txt = p.text.strip()
            if re.match(r'^[一二三四五六七八九十]+[\s、]', txt) or re.match(r'^\d+[\.\s]', txt):
                ok = False
        if ok:
            self._award(3, msg)
        else:
            self._miss(msg)

    def _check_ref_bracket_number(self):
        """+3：每条参考文献以"[序号]"开头，序号从 1 起连续编号。"""
        msg = '参考文献每条以"[序号]"开头，序号从1起连续编号'
        entries = self._ref_entry_indices()
        if not entries:
            self._miss(msg)
            return
        expected = 1
        ok = True
        counted = 0
        for i in entries:
            text = self._flat_paragraphs[i].text.strip()
            m = re.match(r'^[\[【]\s*(\d+)\s*[\]】]', text)
            if not m:
                # 允许存在非条目的说明行？参考文献区应全部为条目，缺失即失败
                ok = False
                break
            if int(m.group(1)) != expected:
                ok = False
                break
            expected += 1
            counted += 1
        if ok and counted >= 1:
            self._award(3, msg)
        else:
            self._miss(msg + f"（连续编号校验失败，识别条目数={counted}）")

    def _check_ref_font(self):
        """+3：参考文献条目为宋体、五号。"""
        msg = '参考文献条目为宋体、五号'
        entries = self._ref_entry_indices()
        if not entries:
            self._miss(msg)
            return
        bad = None
        for i in entries:
            p = self._flat_paragraphs[i]
            for r in self._text_runs(p):
                font = self._run_east_asia_font(r, p)
                size = self._run_size_pt(r, p)
                has_cjk = any(ord(c) > 0x2E80 for c in r.text)
                if has_cjk and not self._is_font_song(font):
                    bad = (p.text[:16], f'字体{font}')
                    break
                if size is None or abs(size - PT_WUHAO) > 0.6:
                    bad = (p.text[:16], f'字号{size}')
                    break
            if bad:
                break
        if bad is None:
            self._award(3, msg)
        else:
            self._miss(msg + f"（条目「{bad[0]}」{bad[1]}）")

    def _check_ref_hanging_indent(self):
        """+1：参考文献条目为悬挂缩进（回行与序号后文字对齐）。"""
        msg = '参考文献条目采用悬挂缩进（回行与序号后文字对齐）'
        entries = self._ref_entry_indices()
        if not entries:
            self._miss(msg)
            return
        ok = True
        any_checked = False
        for i in entries:
            p = self._flat_paragraphs[i]
            pPr = p._element.pPr
            hanging = None
            if pPr is not None:
                ind = pPr.find(qn('w:ind'))
                if ind is not None:
                    hanging = ind.get(qn('w:hanging')) or ind.get(qn('w:hangingChars'))
            if hanging is None:
                # 通过 python-docx：first_line_indent 为负表示悬挂
                try:
                    fli = p.paragraph_format.first_line_indent
                    if fli is not None and int(fli) < 0:
                        hanging = str(-int(fli))
                except Exception:
                    hanging = None
            any_checked = True
            try:
                if hanging is None or int(hanging) <= 0:
                    ok = False
                    break
            except ValueError:
                ok = False
                break
        if ok and any_checked:
            self._award(1, msg)
        else:
            self._miss(msg)

    # ================= 扣分点 =================
    def _check_body_extra_images(self):
        """-3：正文各页中出现封面页两张之外的图片。"""
        msg = '正文各页中出现封面页两张之外的图片'
        try:
            with zipfile.ZipFile(self.open_path) as zf:
                names = zf.namelist()
                media = [n for n in names if n.startswith('word/media/')]
        except Exception:
            self._no_deduct("无法读取媒体资源，视为无多余图片")
            return
        # 封面预期两张图；正文一般无图。图片总数明显多于 2 视为多余
        # 更稳健：统计 document.xml 中 drawing 数量，超过封面 2 张即判多余
        try:
            with zipfile.ZipFile(self.open_path) as zf:
                doc = zf.read('word/document.xml').decode('utf-8')
            draw_count = len(re.findall(r'<w:drawing[ >]', doc)) or doc.count('<pic:pic')
        except Exception:
            draw_count = len(media)
        if draw_count > 2:
            self._deduct(3, msg + f"（检测到图片/绘图数量={draw_count}，超过封面2张）")
        else:
            self._no_deduct(f"正文无多余图片（图片/绘图数量={draw_count}）")

    def _check_revisions_comments(self):
        """-1：文档保留修订标记或批注。"""
        msg = '文档保留了修订标记或批注'
        try:
            with zipfile.ZipFile(self.open_path) as zf:
                names = set(zf.namelist())
                doc = zf.read('word/document.xml').decode('utf-8')
        except Exception:
            self._no_deduct("无法读取，视为无修订/批注")
            return
        has_ins = bool(re.search(r'<w:ins[ >]', doc))
        has_del = bool(re.search(r'<w:del[ >]', doc))
        has_moves = bool(re.search(r'<w:moveFrom[ >]', doc)) or bool(re.search(r'<w:moveTo[ >]', doc))
        has_comment_ref = bool(re.search(r'<w:commentReference[ >]', doc))
        has_comments_part = 'word/comments.xml' in names
        if has_ins or has_del or has_moves or has_comment_ref or has_comments_part:
            self._deduct(1, msg + f"（修订ins={has_ins}, del={has_del}, 批注={has_comment_ref or has_comments_part}）")
        else:
            self._no_deduct("无修订标记与批注")

    def _check_color_highlight_strike(self):
        """-1：正文出现彩色字体、突出显示（高亮）或删除线等异常格式。"""
        msg = '正文出现彩色字体、突出显示（高亮）或删除线等异常格式'
        try:
            with zipfile.ZipFile(self.open_path) as zf:
                doc = zf.read('word/document.xml').decode('utf-8')
        except Exception:
            self._no_deduct("无法读取，视为无异常格式")
            return
        hit = []
        # 高亮
        if re.search(r'<w:highlight\s+w:val="(?!none)', doc):
            hit.append('高亮')
        # 删除线
        if re.search(r'<w:strike\b(?![^>]*w:val="(0|false|none)")', doc) or '<w:dstrike' in doc:
            hit.append('删除线')
        # 彩色字体：w:color val 非 auto/000000/黑
        for m in re.finditer(r'<w:color\s+w:val="([0-9A-Fa-f]{6}|auto)"', doc):
            val = m.group(1).lower()
            if val not in ('auto', '000000'):
                hit.append(f'彩色字体#{val}')
                break
        if hit:
            self._deduct(1, msg + f"（{', '.join(hit)}）")
        else:
            self._no_deduct("无彩色/高亮/删除线")

    def _check_header_unrelated(self):
        """-1：任意页页眉中出现文字或横线（本文档不应有任何页眉内容）。"""
        msg = '任意页页眉中出现文字或横线（本文档不应有任何页眉内容）'
        try:
            with zipfile.ZipFile(self.open_path) as zf:
                names = [n for n in zf.namelist() if re.match(r'word/header\d*\.xml', n)]
                hit = None
                for n in names:
                    x = zf.read(n).decode('utf-8')
                    texts = re.findall(r'<w:t[^>]*>([^<]*)</w:t>', x)
                    joined = re.sub(r'\s+', '', ''.join(texts))
                    if joined:
                        hit = f'页眉文字「{joined[:16]}」'
                        break
                    # 横线：段落边框（pBdr）或页眉中的图形/线条/图片
                    if re.search(r'<w:pBdr\b', x):
                        hit = '页眉横线（段落边框）'
                        break
                    if '<v:line' in x or '<w:drawing' in x or '<w:pict' in x:
                        hit = '页眉图形/线条'
                        break
        except Exception:
            self._no_deduct("无法读取页眉，视为无内容")
            return
        if hit:
            self._deduct(1, msg + f"（{hit}）")
        else:
            self._no_deduct("页眉无文字或横线")

    def _check_blank_page(self):
        """-1：文档中存在完全空白的整页。

        通过重建"页"模型判定：按显式分页符 / 段前分页 / 分节边界切分页面，
        统计每页可见文字量；某页可见文字为 0 即为空白页。
        （内容页末尾夹一个空段带分页符属正常排版，不算空白页。）
        """
        msg = '文档中存在完全空白的整页'
        pages = []           # 每页可见字符数
        current = 0
        prev_sec = None
        for i, p in enumerate(self._flat_paragraphs):
            sec = self._para_section(i)
            sec_changed = prev_sec is not None and sec != prev_sec
            prev_sec = sec
            # 段前分页 / 分节：先结束上一页
            if self._para_has_page_break_before(p) or sec_changed:
                pages.append(current)
                current = 0
            # 段内手动分页符：文本按符号切段，逐个断页
            xml = p._element.xml
            n_breaks = len(re.findall(r'<w:br[^>]*w:type="page"', xml))
            vis = len(re.sub(r'\s+', '', p.text))
            if n_breaks == 0:
                current += vis
            else:
                # 近似：本段可见文字计入当前页，然后逐个断页
                current += vis
                for _ in range(n_breaks):
                    pages.append(current)
                    current = 0
        pages.append(current)
        blank_pages = sum(1 for c in pages if c == 0)
        # 允许文末可能出现的 1 个空尾页（排版收尾），≥1 个中间空白页才扣分
        # 更稳健：仅当存在"夹在两个分页之间且无任何可见内容"的页
        if blank_pages >= 1 and any(c == 0 for c in pages[:-1]):
            self._deduct(1, msg + f"（重建页数={len(pages)}, 空白页数={blank_pages}）")
        else:
            self._no_deduct(f"无完全空白整页（重建页数={len(pages)}）")

    def _check_consecutive_blank_lines(self):
        """-1：正文中存在连续 3 个及以上的空行（多余空段落）。"""
        msg = '正文中存在连续3个及以上的空行'
        max_run = 0
        run = 0
        start = self._body_start_idx()
        for i in range(start, len(self._flat_paragraphs)):
            if not self._flat_paragraphs[i].text.strip():
                run += 1
                max_run = max(max_run, run)
            else:
                run = 0
        if max_run >= 3:
            self._deduct(1, msg + f"（最大连续空行={max_run}）")
        else:
            self._no_deduct(f"无3行及以上连续空行（最大连续空行={max_run}）")

    # ---- 新增扣分点：内容完整性 / 结构 / 额外破坏 ----
    _FORM_KEYWORDS = ['论文评审表', '评阅表', '评阅意见', '评审意见', '答辩记录',
                      '答辩登记', '答辩情况', '答辩表', '开题报告', '成绩评定',
                      '评分表', '指导记录', '中期检查', '任务书']

    def _check_extra_full_page_content(self):
        """-3：文档中出现本细则所列结构之外的整页内容（评审表、答辩记录表、
        开题报告、空白模板页等）。以典型表单关键词在段落/表格中的出现判定。"""
        msg = ('文档中出现本细则所列结构之外的整页内容'
               '（如论文评审表、答辩记录表、开题报告、空白模板页）')
        hit = None
        for p in self._flat_paragraphs:
            norm = re.sub(r'\s+', '', p.text)
            for k in self._FORM_KEYWORDS:
                if k in norm:
                    hit = k
                    break
            if hit:
                break
        if hit is None:
            for t in self._flat_tables:
                for row in t.rows:
                    for c in row.cells:
                        norm = re.sub(r'\s+', '', c.text)
                        for k in self._FORM_KEYWORDS:
                            if k in norm:
                                hit = k
                                break
                        if hit:
                            break
                    if hit:
                        break
                if hit:
                    break
        if hit:
            self._deduct(3, msg + f"（检出「{hit}」）")
        else:
            self._no_deduct("无结构外整页内容")

    def _check_missing_chapters(self):
        """-3：正文缺少应有章节。应含六个一级章节（绪论/相关理论概述/现状及问题/
        原因/改进策略/结论）及"参考文献""致谢"；任一部分标题缺失或标题下无正文段落即扣。"""
        msg = ('正文缺少应有的章节（六个一级章节及参考文献、致谢，'
               '标题缺失或标题下无正文段落）')
        h1 = self._h1_indices()
        heading_texts = {i: re.sub(r'\s+', '', self._flat_paragraphs[i].text) for i in h1}
        expected = [
            ('绪论', ('绪论', '引言')),
            ('相关理论概述', ('理论概述', '相关理论', '理论基础')),
            ('现状及存在的问题', ('现状',)),
            ('原因', ('原因',)),
            ('改进策略', ('改进策略', '改进', '对策', '规范化改进')),
            ('结论', ('结论', '结语', '结束语')),
            ('参考文献', ('参考文献',)),
            ('致谢', ('致谢',)),
        ]
        missing = []
        present = {}
        for label, kws in expected:
            found = -1
            for i in h1:
                if any(kw in heading_texts[i] for kw in kws):
                    found = i
                    break
            if found < 0:
                missing.append(label)
            else:
                present[label] = found
        # 标题下有无正文段落
        empty_sections = []
        for label, idx in present.items():
            nxt = len(self._flat_paragraphs)
            for i in h1:
                if i > idx:
                    nxt = i
                    break
            has_body = False
            for j in range(idx + 1, nxt):
                pj = self._flat_paragraphs[j]
                if pj.text.strip() and not self._is_heading(pj):
                    has_body = True
                    break
            if not has_body:
                empty_sections.append(label)
        if missing or empty_sections:
            self._deduct(3, msg + f"（缺失章节={missing}，空标题={empty_sections}）")
        else:
            self._no_deduct("六个一级章节及参考文献、致谢结构完整")

    def _check_body_text_tamper(self):
        """-3：抽检正文段落，出现占位/提示文字、句中截断或整段重复粘贴。
        （无原文可比对，仅做内部一致性判定。）"""
        msg = '正文段落出现占位/提示文字、句子中间截断或整段重复粘贴'
        placeholders = ['loremipsum', '此处为正文', '正文内容', '待补充', '占位',
                        'todo', 'xxxx', '请在此', 'tobefilled', '内容待定']
        hit = None
        prev_norm = None
        for i in self._body_text_indices():
            t = self._flat_paragraphs[i].text.strip()
            norm = re.sub(r'\s+', '', t).lower()
            for ph in placeholders:
                if ph in norm:
                    hit = f'占位符「{t[:16]}」'
                    break
            if hit:
                break
            # 句中截断：长段以逗号/顿号/分号结尾
            if len(norm) >= 20 and t[-1] in '，、；,;':
                hit = f'疑似截断「…{t[-12:]}」'
                break
            # 段内重复粘贴：长段前后半完全相同
            half = len(norm) // 2
            if len(norm) >= 40 and norm[:half] == norm[half:half * 2]:
                hit = f'段内重复粘贴「{t[:16]}」'
                break
            # 相邻段落完全相同
            if prev_norm is not None and len(norm) >= 20 and norm == prev_norm:
                hit = f'相邻段落重复「{t[:16]}」'
                break
            prev_norm = norm
        if hit:
            self._deduct(3, msg + f"（{hit}）")
        else:
            self._no_deduct("正文段落无占位/截断/重复异常")

    def _check_reference_content(self):
        """-3：参考文献条目不足六条，或作者/篇名/出处残缺、为空、出现乱码。
        以固定的六位作者作为自包含比对基准。"""
        msg = '参考文献条目不足六条，或作者/篇名/出处残缺、为空、出现乱码'
        entries = self._ref_entry_indices()
        if not entries:
            self._deduct(3, msg + "（未找到参考文献条目）")
            return
        texts = [self._flat_paragraphs[i].text.strip() for i in entries]
        reason = None
        if len(entries) < 6:
            reason = f'条目数={len(entries)}<6'
        if reason is None:
            for t in texts:
                n = re.sub(r'\s+', '', t)
                body = re.sub(r'^[\[【]?\s*\d+\s*[\]】\.、]?', '', n)
                if len(body) < 4:
                    reason = f'条目过短/为空「{t[:16]}」'
                    break
                if '�' in t or '口口口' in t or '□' in t:
                    reason = f'条目含乱码「{t[:16]}」'
                    break
        if reason is None:
            expected_authors = ['许澜', '段知衡', '江映初', '马清禾', '赵予安', '林若弦']
            block = ''.join(texts)
            cover = sum(1 for a in expected_authors if a in block)
            if cover < 6:
                reason = f'应有六位作者仅匹配{cover}位（作者被改动/删减）'
        if reason:
            self._deduct(3, msg + f"（{reason}）")
        else:
            self._no_deduct("参考文献六条完整")

    def _check_garbled_chars(self):
        """-1：文档任意位置出现乱码、方块字，或问号、口口口等异常字符。"""
        msg = '文档出现乱码、方块字，或显示为问号、口口口等异常字符'
        text = ''.join(p.text for p in self._flat_paragraphs)
        hit = None
        if '�' in text:
            hit = '替换字符\\uFFFD'
        elif '口口口' in text:
            hit = '口口口'
        elif re.search(r'[□■]{2,}', text):
            hit = '连续方块字'
        elif any(0xE000 <= ord(c) <= 0xF8FF for c in text):
            hit = '私用区异常字符'
        if hit:
            self._deduct(1, msg + f"（{hit}）")
        else:
            self._no_deduct("无乱码/异常字符")

    def _check_extra_structure_content(self):
        """-1：出现本细则结构之外、篇幅不足整页的多余内容：第七个及以后的一级章节、
        六条之外的多余参考文献条目、英文摘要（Abstract）等。"""
        msg = '出现结构之外的多余内容（多余章节/多余参考文献/英文摘要等）'
        cn = '一二三四五六七八九十'
        hit = []
        for i in self._h1_indices():
            t = self._flat_paragraphs[i].text.strip()
            m = re.match(r'^([一二三四五六七八九十]+)[\s、]', t)
            if m:
                num = m.group(1)
                if (len(num) == 1 and num in cn and cn.index(num) + 1 >= 7) or len(num) > 1 or num.startswith('十'):
                    hit.append(f'多余章节「{t[:12]}」')
        if len(self._ref_entry_indices()) > 6:
            hit.append(f'参考文献超过六条(={len(self._ref_entry_indices())})')
        for p in self._flat_paragraphs:
            norm = re.sub(r'\s+', '', p.text).lower()
            if norm.startswith('abstract') or norm.startswith('keywords'):
                hit.append('英文摘要/关键词')
                break
        if hit:
            self._deduct(1, msg + f"（{hit[:3]}）")
        else:
            self._no_deduct("无结构外多余内容")

    def _check_page_size_orientation(self):
        """-1：全文纸张被改为非 A4 尺寸，或出现横向（landscape）页面。"""
        from docx.enum.section import WD_ORIENT
        msg = '全文纸张被改为非A4尺寸，或出现横向页面'
        bad = []
        for si, s in enumerate(self.doc.sections):
            try:
                w = s.page_width.cm
                h = s.page_height.cm
            except Exception:
                continue
            try:
                orient = s.orientation
            except Exception:
                orient = None
            if orient == WD_ORIENT.LANDSCAPE or w > h + 0.1:
                bad.append(f'节{si}横向({round(w,1)}×{round(h,1)})')
                continue
            dims = sorted([w, h])
            if not (abs(dims[0] - 21.0) <= 0.5 and abs(dims[1] - 29.7) <= 0.5):
                bad.append(f'节{si}非A4({round(w,1)}×{round(h,1)})')
        if bad:
            self._deduct(1, msg + f"（{bad[:3]}）")
        else:
            self._no_deduct("全文A4纵向")

    def _check_hyperlink_footnote(self):
        """-1：正文文字中添加了与论文内容无关的外部超链接，或出现脚注、尾注。
        （目录自动生成的内部锚点超链接不计。）"""
        msg = '正文出现与内容无关的外部超链接，或出现脚注、尾注'
        try:
            with zipfile.ZipFile(self.open_path) as zf:
                names = set(zf.namelist())
                doc = zf.read('word/document.xml').decode('utf-8')
                rels = ''
                if 'word/_rels/document.xml.rels' in names:
                    rels = zf.read('word/_rels/document.xml.rels').decode('utf-8')
        except Exception:
            self._no_deduct("无法读取，视为无超链接/脚注")
            return
        hit = []
        for m in re.finditer(r'<Relationship\b[^>]*>', rels):
            tag = m.group(0)
            if 'hyperlink' in tag.lower() and 'External' in tag:
                hit.append('外部超链接')
                break
        if re.search(r'<w:footnoteReference', doc) or re.search(r'<w:endnoteReference', doc):
            hit.append('脚注/尾注')
        if hit:
            self._deduct(1, msg + f"（{'、'.join(hit)}）")
        else:
            self._no_deduct("无外部超链接与脚注尾注")

    # ================= 计分与输出 =================
    def _calculate_total_score(self):
        total = sum(self.results['dimension2_scores']) - sum(self.results['dimension2_deductions'])
        self.results['total_score'] = total

    def _print_results(self):
        print()
        print("=" * 60)
        print("评估结果汇总")
        print("=" * 60)
        if not self.results['dimension1_passed']:
            print("维度一未通过，直接判 0 分：")
            for f in self.results['dimension1_failures']:
                print(f"  - {f}")
            print()
            print("最终得分：0")
            return
        print("命中的评分点：")
        if self.results['hit_points']:
            for hp in self.results['hit_points']:
                print(f"  {hp}")
        else:
            print("  （无）")
        print()
        print(f"得分点合计：+{sum(self.results['dimension2_scores'])}")
        print(f"扣分点合计：-{sum(self.results['dimension2_deductions'])}")
        print(f"最终得分：{self.results['total_score']}")


# ============================================================
# 统一对外接口（§2.2 约定，对齐 officeval_008_verifier.py）
# ============================================================
# 维度二规则元数据：(kind, rule_desc, max_delta)
#   kind: 'bonus'（得分点，max_delta 为正）/ 'penalty'（扣分点，max_delta 为负）
# 命中得分点累加其正分、命中扣分点累加其负分；本表顺序与 _evaluate_dimension2
# 中各检查方法的调用顺序严格一致，据此把评估结果解析为 dim2_items。
_DIM2_RULES = [
    # ---- 得分点 ----
    ('bonus', '封面页结构（成教本科毕业论文+五个字段行）', 5),
    ('bonus', '封面标题黑体小初加粗居中', 1),
    ('bonus', '封面两张图片尺寸', 3),
    ('bonus', '封面五字段行黑体三号加粗居中+下划线', 3),
    ('bonus', '封面底部日期黑体四号加粗居中', 1),
    ('bonus', '封面单独占一页', 1),
    ('bonus', '独立原创性声明页', 5),
    ('bonus', '声明页标题宋体二号加粗居中', 1),
    ('bonus', '声明页三段声明正文', 3),
    ('bonus', '声明正文宋体四号1.5倍首行缩进两端对齐', 1),
    ('bonus', '声明页签名行', 1),
    ('bonus', '独立目录页（宋体二号居中）', 5),
    ('bonus', '自动生成目录（层级缩进+页码右对齐）', 3),
    ('bonus', '正文首页论文标题小二加粗居中', 3),
    ('bonus', '摘要整段（黑体四号标签+宋体小四内容）', 3),
    ('bonus', '关键词段（黑体四号标签+宋体小四内容）', 1),
    ('bonus', '正文段落宋体小四', 5),
    ('bonus', '正文段落首行缩进两字符', 3),
    ('bonus', '正文段落1.5倍行距', 3),
    ('bonus', '正文段落两端对齐', 3),
    ('bonus', '一级标题样式为标题1', 3),
    ('bonus', '一级标题中文数字序号连续', 3),
    ('bonus', '一级标题黑体三号居中', 3),
    ('bonus', '二级标题样式为标题2', 3),
    ('bonus', '二级标题全角括号中文数字序号', 5),
    ('bonus', '二级标题黑体四号左对齐', 3),
    ('bonus', '页边距上下2.54cm左右3.0cm', 3),
    ('bonus', '正文页脚自动页码"第X页"', 3),
    ('bonus', '页脚页码居中宋体五号/小四', 1),
    ('bonus', '正文首页页码第1页并逐页递增', 1),
    ('bonus', '封面/声明/目录/正文分节且正文重起页码', 3),
    ('bonus', '结论/参考文献/致谢各自另起一页', 1),
    ('bonus', '参考文献/致谢标题与一级标题一致不编号', 3),
    ('bonus', '参考文献[序号]连续编号', 3),
    ('bonus', '参考文献条目宋体五号', 3),
    ('bonus', '参考文献条目悬挂缩进', 1),
    # ---- 扣分点 ----（顺序须与 _evaluate_dimension2 中扣分检查调用顺序严格一致）
    ('penalty', '结构外整页内容(评审表/答辩记录/开题报告/空白模板页)', -3),
    ('penalty', '正文封面两张之外的图片', -3),
    ('penalty', '保留批注框或修订标记', -1),
    ('penalty', '彩色字体/高亮/删除线', -1),
    ('penalty', '页眉出现文字或横线', -1),
    ('penalty', '完全空白整页', -1),
    ('penalty', '连续3个及以上空行', -1),
    ('penalty', '缺少应有章节(六章+参考文献+致谢)', -3),
    ('penalty', '正文段落占位/截断/重复粘贴', -3),
    ('penalty', '参考文献不足六条或内容残缺', -3),
    ('penalty', '乱码/方块字/异常字符', -1),
    ('penalty', '结构外多余内容(多余章节/文献/英文摘要)', -1),
    ('penalty', '纸张非A4或横向页面', -1),
    ('penalty', '无关外部超链接或脚注尾注', -1),
]


# 满分（max_score）= 所有得分点分值之和（不含扣分点）
assert MAX_SCORE_TOTAL == sum(d for kind, _, d in _DIM2_RULES if kind == 'bonus')


def evaluate(dir_path):
    """按 §2.2 约定返回结果 dict；不 print、不改全局 sys.stdout、不 sys.exit。

    入参 dir_path 为"脚本所在目录的路径"，脚本自己在该目录内定位并打开被评估
    的 .docx（忽略以 ~$ 开头的 Word 临时锁文件）。脚本自身崩溃（含目录不存在、
    未发现文档等），由本函数捕获并返回 status='error'。
    """
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
            raise FileNotFoundError(f"目录不存在：{dir_path}")
        cands = sorted(
            f for f in os.listdir(dir_path)
            if f.lower().endswith('.docx') and not f.startswith('~$')
        )
        if not cands:
            raise FileNotFoundError(f"目录内未找到 .docx 文件：{dir_path}")
        file_path = os.path.join(dir_path, cands[0])
        result["file_name"] = os.path.basename(file_path)
        return _evaluate_impl(file_path, result)
    except Exception as exc:
        # status 为 error 时未能完成判定，dim1_pass 留空，dim1_reason 也为空
        result["status"] = "error"
        result["error"] = f"{type(exc).__name__}: {exc}"
        result["dim1_pass"] = ""
        result["dim1_reason"] = ""
        result["dim2_items"] = []
        result["total_score"] = 0
        return result


def _evaluate_impl(file_path, result):
    """维度一 + 维度二评估主体：沿用类内既有评分逻辑（WordDocumentEvaluator），
    末尾把评估器结果解析为结构化 dim2_items 填入 result。评分逻辑保持不变。"""
    evaluator = WordDocumentEvaluator(file_path)
    # 评估过程中的调试打印重定向到 devnull，保持返回值纯净（不改全局 sys.stdout）
    devnull = open(os.devnull, 'w', encoding='utf-8')
    try:
        with contextlib.redirect_stdout(devnull), contextlib.redirect_stderr(devnull):
            evaluator.evaluate()
    finally:
        devnull.close()

    res = evaluator.results
    result["dim1_pass"] = bool(res['dimension1_passed'])
    if not res['dimension1_passed']:
        result["dim1_reason"] = '；'.join(res['dimension1_failures']) or '维度一未通过'
        result["dim2_items"] = []
        result["total_score"] = 0
        return result

    result["dim1_reason"] = ""
    result["dim2_items"] = _build_dim2_items(res)
    result["total_score"] = res['total_score']
    result["max_score"] = MAX_SCORE_TOTAL
    return result


def _build_dim2_items(res):
    """根据评估器结果重建 dim2_items 列表。

    每项：{rule, max_delta, delta, hit}
      - rule: 规则标签（取自 _DIM2_RULES）
      - max_delta: 该规则分值（得分点为正、扣分点为负）
      - delta: 实际计入分值（命中得分点=正分；命中扣分点=负分；未命中=0）
      - hit: 是否命中

    维度二各检查方法按 _DIM2_RULES 顺序依次执行，且每个方法恰好产生一条
    dim2_outcomes 记录，因此二者按位置一一对应。
    """
    outcomes = res.get('dim2_outcomes', [])
    items = []
    for i, (kind, rule_desc, max_delta) in enumerate(_DIM2_RULES):
        if i < len(outcomes):
            oc = outcomes[i]
            items.append({
                'rule': rule_desc,
                'max_delta': max_delta,
                'delta': oc['delta'],
                'hit': oc['hit'],
            })
        else:
            items.append({
                'rule': rule_desc,
                'max_delta': max_delta,
                'delta': 0,
                'hit': False,
            })
    return items


if __name__ == "__main__":
    # 本地调试入口：evaluate 只接收"脚本所在目录的路径"，由脚本自己在该目录内
    # 定位并打开被评估的 .docx。命令行第 1 个参数可覆盖该目录，未指定时使用
    # 脚本自身所在目录。
    if len(sys.argv) >= 2:
        _dir = sys.argv[1]
    else:
        _dir = os.path.dirname(os.path.abspath(__file__))
    _out = evaluate(_dir)
    print(json.dumps(_out, ensure_ascii=False, indent=2))
