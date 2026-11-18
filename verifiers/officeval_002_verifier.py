#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# pyright: reportUnknownMemberType=false, reportUnknownVariableType=false, reportUnknownArgumentType=false, reportUnknownParameterType=false, reportMissingParameterType=false, reportMissingTypeArgument=false, reportImplicitStringConcatenation=false, reportGeneralTypeIssues=false, reportArgumentType=false, reportAttributeAccessIssue=false
"""
页眉横线调整任务 —— 完成度自动评分脚本

对外唯一入口(见 "脚本接口差异与统一建议.md" §2.1/§2.2):
    def evaluate(file_path: str) -> dict
        - file_path 既可是任务目录(自动定位其中的 .docx), 也可是单个 .docx;
        - 返回规范化 dict(id / file_name / status / error / dim1_pass /
          dim1_reason / dim2_items / total_score / max_score);
        - 主结果只走 return, 不 print、不改 sys.stdout、不 sys.exit、不硬编码路径。

    if __name__ == "__main__": 仅用于本地调试, 调用 evaluate(sys.argv[1]) 后
    把规范化 dict 以 JSON 打印到 stdout。

    内部评分逻辑(Docx 解析 / score / report / 各校验项判定语义)保持不变,
    仅在文件末尾新增统一对外接口层。

评分依据《细则.docx》:
  维度一 可用与可修改性: 仅核验文件可否正常打开(不可打开则总分0)
  维度二 完成度:
    加分项(来自指令任务, 满分 13):
      +5 摘要及之后页眉横线 = 与PDF一致(thinThickMediumGap, sz=18/2.25pt 双线)
      +3 摘要之前页眉横线全部删除(none / 无 pBdr)
      +3 在摘要前正确断开分节(摘要前后分属不同节)
      +1 横线分界点恰在摘要页(前无线、摘要页起有线)
      +1 单双数页页眉横线一致(关闭奇偶页页眉或 even 同步)
    减分项(指令只要求改页眉横线; 凡改到横线以外即为多余修改):
      共 55 条, 与《细则.docx》"二、减分项"逐条对应, 分 11 组:
        中文封面 6 / 英文封面 5 / 声明页 6 / 摘要与 Abstract 6 /
        目录与前置内容控件 5 / 正文各级标题 4 /
        正文段落与参考文献、附录、致谢 5 / 图表与题注 5 /
        页面设置 3 / 页眉文字与页脚页码 5 / 通用破坏 5
      每条判定都直接比对细则写明的绝对格式值(字体/字号/加粗/下划线/对齐/
      行距/缩进/段前段后/纸张/页边距/图片尺寸/表格行列/字数等), 不需要原始
      输入文档作参照; 命中一条扣对应分, 可累计。
      字体名按同义归并(SimSun≡宋体, SimHei≡黑体), 避免不同 Word 语言版本
      存名方式不同造成误判。
      首行缩进按原始输入文档的记录方式比对: 原文以 w:firstLineChars 记录的段落
      (如声明、摘要、Abstract)按"N 字符"比对; 原文以 w:firstLine 绝对值记录的
      段落(如正文各章、参考文献、附录、致谢, 值为 480 twips = 0.85 厘米)按厘米
      比对。两种单位互相等值改写(如 WPS 另存会把绝对值归一为字符数)排版结果
      相同, 均视为未改动。

无第三方依赖(仅用标准库解析 docx=zip+xml)。
"""
import sys, zipfile, re, json
import xml.etree.ElementTree as ET

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
def w(tag): return f'{{{W}}}{tag}'

# PDF 参照基准(细则量化基准): 粗线约2.25pt, 双线细+粗
PDF_SZ = 18            # 2.25pt = 8*2.25
PDF_STYLE = 'thinThickMediumGap'

# ---- 细则写明的绝对量 ----
TEXTLEN = 12525            # 可见文字字数
TEXT_TOL = 0.01            # 允许偏差 ±1%
N_IMG = 4                  # 内嵌图片张数
TBL_SHAPE = [(5, [1] * 5), (3, [1] * 3), (5, [3] * 5), (6, [3] * 6)]
N_PAGEFLD = 8              # 页脚 PAGE 域数量
HDR_TEXT = '澄湖理工大学硕士学位论文'
FTR_TEXT = '城市社区微气候韧性评价与公共空间更新研究'
FTR_ALIGN = [('right', 'right', 240), ('right', 'right', 120),
             ('left', 'left', 180), ('right', 'right', 300)]
PG_SZ = (11906, 16838)                       # A4
PG_MAR_FRONT = dict(top=1440, bottom=1440, left=1803, right=1803)
PG_MAR_MAIN = dict(top=1440, bottom=1440, left=1701, right=1587)
PG_HDR_FTR = (850, 992)                      # 页眉 15mm / 页脚 17.5mm
IMG_CM = [(7.13, 2.07), (10.56, 6.41), (13.80, 8.38), (12.70, 8.38)]
KEEP_VALUES = ['唐知遥', '许怀砚 教授', '工学硕士', '75264', '2024AQ0926',
               'Tang Zhi Yao', 'Professor Xu Huai yan', 'Master of Engineering']


class Docx:
    """轻量 docx 读取: 只解析评分需要的部分。"""
    def __init__(self, path):
        self.path = path
        self.ok = False
        self.err = None
        try:
            self.zf = zipfile.ZipFile(path)
            self.document = self._xml('word/document.xml')
            self.settings = self._xml_opt('word/settings.xml')
            self.rels = self._read('word/_rels/document.xml.rels')
            self.ok = True
        except Exception as e:
            self.err = str(e)

    def _read(self, name):
        try:
            return self.zf.read(name).decode('utf-8', 'ignore')
        except KeyError:
            return ''

    def _xml(self, name):
        return ET.fromstring(self.zf.read(name))

    def _xml_opt(self, name):
        try:
            return ET.fromstring(self.zf.read(name))
        except KeyError:
            return None

    # --- rId -> header 文件 映射 ---
    def rid_to_header(self):
        m = {}
        for rid, tgt in re.findall(r'Id="([^"]+)"[^>]*Target="(header\d+\.xml)"', self.rels):
            m[rid] = 'word/' + tgt
        for tgt, rid in re.findall(r'Target="(header\d+\.xml)"[^>]*Id="([^"]+)"', self.rels):
            m[rid] = 'word/' + tgt
        return m

    def header_bottom(self, headerfile):
        """返回该页眉段落底边框 dict 或 None(无横线)。"""
        xml = self._read(headerfile)
        mb = re.search(r'<w:bottom\b[^>]*/>', xml)
        if not mb:
            return None
        seg = mb.group(0)
        val = re.search(r'w:val="([^"]*)"', seg)
        sz = re.search(r'w:sz="([^"]*)"', seg)
        val = val.group(1) if val else ''
        sz = int(sz.group(1)) if sz else 0
        if val == 'none' or sz == 0:
            return None
        return {'val': val, 'sz': sz}

    def sections(self):
        """按出现顺序返回各 sectPr 的 (字符位置, {type:rid}). document 末尾 sectPr 也计入。"""
        raw = self.zf.read('word/document.xml').decode('utf-8', 'ignore')
        secs = []
        for m in re.finditer(r'<w:sectPr\b.*?</w:sectPr>|<w:sectPr\b[^>]*/>', raw, re.S):
            refs = {}
            for hr in re.finditer(r'<w:headerReference\b[^>]*/>', m.group(0)):
                t = re.search(r'w:type="([^"]*)"', hr.group(0))
                i = re.search(r'r:id="([^"]*)"', hr.group(0))
                if t and i:
                    refs[t.group(1)] = i.group(1)
            secs.append((m.start(), refs))
        return raw, secs

    def abstract_pos(self):
        raw = self.zf.read('word/document.xml').decode('utf-8', 'ignore')
        for kw in ['摘  要', '摘 要', '摘要']:
            i = raw.find(kw)
            if i >= 0:
                return raw, i
        return raw, -1

    def abstract_idx(self):
        """"摘  要"标题段的段落序号(与 sects_ranged 同一序号空间); 找不到返回 -1。

        判定分节点是否紧邻摘要必须用段落序号: sectPr 在 XML 里的字符位置只能
        说明"分节符在摘要之前", 无法区分"紧接摘要之前"与"提前若干段"。
        """
        body = self.document.find(w('body'))
        idx, found, loose = [0], [-1], [-1]

        def walk(node):
            for ch in list(node):
                if ch.tag == w('p'):
                    if found[0] < 0:
                        txt = re.sub(r'\s+', '', ''.join(
                            t.text or '' for t in ch.iter(w('t'))))
                        if txt == '摘要':
                            found[0] = idx[0]
                        elif loose[0] < 0 and '摘要' in txt and len(txt) <= 8:
                            loose[0] = idx[0]
                    idx[0] += 1
                elif ch.tag == w('sectPr'):
                    continue
                else:
                    walk(ch)

        walk(body)
        return found[0] if found[0] >= 0 else loose[0]

    def even_odd_headers(self):
        if self.settings is None:
            return False
        return self.settings.find(w('evenAndOddHeaders')) is not None

    def body_text(self):
        parts = []
        body = self.document.find(w('body'))
        for t in body.iter(w('t')):
            parts.append(t.text or '')
        return ''.join(parts)

    def counts(self):
        raw = self.zf.read('word/document.xml').decode('utf-8', 'ignore')
        return {
            'sdt': raw.count('<w:sdt>'),
            'drawing': raw.count('<w:drawing>'),
            'tbl': raw.count('<w:tbl>'),
        }

    def footers_sig(self):
        sig = []
        for name in sorted(self.zf.namelist()):
            if re.match(r'word/footer\d+\.xml', name):
                sig.append((name.split('/')[-1], len(self._read(name))))
        return sig

    # ---------- 自包含完整性核查(无需原始输入文件) ----------
    def _all_raw_parts(self):
        """document 及所有 header/footer 的原始 xml 文本, 供跨部件核查。"""
        parts = {'word/document.xml':
                 self.zf.read('word/document.xml').decode('utf-8', 'ignore')}
        for n in self.zf.namelist():
            if re.match(r'word/(header|footer)\d+\.xml', n):
                parts[n] = self._read(n)
        return parts

    def rels_map(self):
        """document.xml 的 rId -> (类型末段, Target)。"""
        m = {}
        for tag in re.findall(r'<Relationship\b[^>]*?/>', self.rels):
            rid = re.search(r'Id="([^"]+)"', tag)
            typ = re.search(r'Type="([^"]*)"', tag)
            tgt = re.search(r'Target="([^"]*)"', tag)
            if rid:
                t = typ.group(1).rsplit('/', 1)[-1] if typ else ''
                m[rid.group(1)] = (t, tgt.group(1) if tgt else '')
        return m

    def _part_rels(self, part):
        """任意部件(如 word/header2.xml)对应的 rId -> Target。"""
        folder, _, fname = part.rpartition('/')
        txt = self._read(f'{folder}/_rels/{fname}.rels')
        m = {}
        for tag in re.findall(r'<Relationship\b[^>]*?/>', txt):
            rid = re.search(r'Id="([^"]+)"', tag)
            tgt = re.search(r'Target="([^"]*)"', tag)
            if rid:
                m[rid.group(1)] = tgt.group(1) if tgt else ''
        return m

    # ---------- 绝对量核查所需的结构信息 ----------
    def text_len(self):
        return len(self.body_text())

    def dangling(self):
        """document.xml 内引用了但 rels 中不存在的 rId(悬挂引用)。"""
        raw = self.zf.read('word/document.xml').decode('utf-8', 'ignore')
        rels = self.rels_map()
        used = set(re.findall(r'r:(?:id|embed|link)="([^"]+)"', raw))
        return sorted(u for u in used if u not in rels)

    def sdt_info(self):
        """(内容控件总数, 内容为空的个数)。"""
        total = broken = 0
        body = self.document.find(w('body'))
        for sdt in body.iter(w('sdt')):
            total += 1
            content = sdt.find(w('sdtContent'))
            if content is None or len(list(content)) == 0:
                broken += 1
        return total, broken

    def footer_info(self):
        """页脚: 缺失引用 / 域配对 / PAGE 域数量 / 各页脚(对齐, 缩进, 字号, 文字)。"""
        raw = self.zf.read('word/document.xml').decode('utf-8', 'ignore')
        names = set(self.zf.namelist())
        rels = self.rels_map()
        missing = []
        for rid in sorted(set(re.findall(r'<w:footerReference\b[^>]*r:id="([^"]+)"', raw))):
            tgt = rels.get(rid, ('', ''))[1]
            if not tgt or ('word/' + tgt) not in names:
                missing.append(rid)
        begin = sep = end = page = 0
        parts = []
        for n in sorted(names):
            if not re.match(r'word/footer\d+\.xml', n):
                continue
            x = self._read(n)
            begin += len(re.findall(r'w:fldCharType="begin"', x))
            sep += len(re.findall(r'w:fldCharType="separate"', x))
            end += len(re.findall(r'w:fldCharType="end"', x))
            page += len(re.findall(r'\bPAGE\b', x))
            jc = re.findall(r'<w:jc w:val="(\w+)"/>', x)
            ind = re.search(r'<w:ind[^/]*/>', x)
            li = ri = None
            if ind:
                mi = re.search(r'w:left="(-?\d+)"', ind.group(0))
                mr = re.search(r'w:right="(-?\d+)"', ind.group(0))
                li = int(mi.group(1)) if mi else None
                ri = int(mr.group(1)) if mr else None
            szs = sorted(set(int(v) for v in re.findall(r'<w:sz w:val="(\d+)"/>', x)))
            txt = ''.join(re.findall(r'<w:t[^>]*>([^<]*)</w:t>', x))
            parts.append({'name': n.split('/')[-1], 'jc': jc[0] if jc else None,
                          'left': li, 'right': ri, 'sz': szs, 'text': txt})
        return {'missing': missing, 'begin': begin, 'separate': sep, 'end': end,
                'page': page, 'parts': parts}

    def header_parts(self):
        """各 header 部件: (文件名, 文字, jc, sz集合, 字体集合, 是否加粗)。"""
        out = []
        for n in sorted(self.zf.namelist()):
            if not re.match(r'word/header\d+\.xml', n):
                continue
            x = self._read(n)
            txt = ''.join(re.findall(r'<w:t[^>]*>([^<]*)</w:t>', x))
            jc = re.findall(r'<w:jc w:val="(\w+)"/>', x)
            szs = sorted(set(int(v) for v in re.findall(r'<w:sz w:val="(\d+)"/>', x)))
            ea = sorted(set(re.findall(r'w:eastAsia="([^"]+)"', x)))
            asc = sorted(set(re.findall(r'w:ascii="([^"]+)"', x)))
            line = re.findall(r'<w:spacing[^/]*w:line="(\d+)"', x)
            bold = bool(re.search(r'<w:b(?: w:val="(?!0|false|off)[^"]*")?/>', x))
            out.append({'name': n.split('/')[-1], 'text': txt,
                        'jc': jc[0] if jc else None, 'sz': szs,
                        'ea': ea, 'ascii': asc, 'bold': bold,
                        'line': int(line[0]) if line else None})
        return out

    def img_info(self):
        """(图片处数, 无法解析的 embed, [(宽cm, 高cm)])。"""
        names = set(self.zf.namelist())
        bad, sizes = [], []
        n_img = 0
        for part, x in self._all_raw_parts().items():
            prels = self._part_rels(part)
            for rid in re.findall(r'r:embed="([^"]+)"', x):
                n_img += 1
                tgt = prels.get(rid)
                resolved = ('word/' + tgt.lstrip('/').replace('../', '')) if tgt else None
                if not resolved or resolved not in names:
                    bad.append((part.split('/')[-1], rid))
        raw = self.zf.read('word/document.xml').decode('utf-8', 'ignore')
        for m in re.finditer(r'<wp:extent cx="(\d+)" cy="(\d+)"', raw):
            sizes.append((round(int(m.group(1)) / 360000.0, 2),
                          round(int(m.group(2)) / 360000.0, 2)))
        return n_img, bad, sizes

    def tbl_shapes(self):
        shapes = []
        body = self.document.find(w('body'))
        for tbl in body.iter(w('tbl')):
            rows = tbl.findall(w('tr'))
            shapes.append((len(rows), [len(r.findall(w('tc'))) for r in rows]))
        return shapes

    def tbl_borders(self, idx):
        """第 idx 个表格的 tblBorders 各边 val(缺省视为 nil)。"""
        body = self.document.find(w('body'))
        tbls = list(body.iter(w('tbl')))
        if idx >= len(tbls):
            return {}
        bd = tbls[idx].find(w('tblPr'))
        bd = bd.find(w('tblBorders')) if bd is not None else None
        out = {}
        if bd is not None:
            for e in list(bd):
                out[e.tag.split('}')[1]] = e.get(w('val'))
        return out

    def _sect_dict(self, sp):
        d = {'pgSz': None, 'pgMar': {}, 'fmt': None, 'refs': {}}
        sz = sp.find(w('pgSz'))
        if sz is not None:
            d['pgSz'] = (int(sz.get(w('w')) or 0), int(sz.get(w('h')) or 0))
        mar = sp.find(w('pgMar'))
        if mar is not None:
            for k in ('top', 'bottom', 'left', 'right', 'header', 'footer'):
                v = mar.get(w(k))
                if v is not None:
                    d['pgMar'][k] = int(v)
        num = sp.find(w('pgNumType'))
        if num is not None:
            d['fmt'] = num.get(w('fmt'))
        for hr in sp.findall(w('headerReference')):
            d['refs'][hr.get(w('type'))] = hr.get(f'{{{R}}}id')
        return d

    def sects_ranged(self):
        """各节: 覆盖的段落序号区间 [start, end] + 页面属性(段落序号与 Fmt.paras 一致)。"""
        body = self.document.find(w('body'))
        order = []
        pi = [0]

        def walk(node):
            for ch in list(node):
                if ch.tag == w('p'):
                    idx = pi[0]
                    pi[0] += 1
                    pPr = ch.find(w('pPr'))
                    sp = pPr.find(w('sectPr')) if pPr is not None else None
                    if sp is not None:
                        order.append((idx, sp))
                elif ch.tag == w('sectPr'):
                    order.append((max(pi[0] - 1, 0), ch))
                else:
                    walk(ch)

        walk(body)
        out, start = [], 0
        for end, sp in order:
            d = self._sect_dict(sp)
            d.update({'start': start, 'end': end})
            out.append(d)
            start = end + 1
        return out




def effective_default_headers_by_region(dx):
    """把各节的 default 页眉按 '摘要之前 / 摘要及之后' 分类, 并判定分节点位置。

    两点关键判据:
    1) 节的归属按其覆盖的**段落区间**判定, 不按 sectPr 自身的字符位置。若某节
       起点在摘要之前、终点在摘要之后(跨过摘要), 说明摘要之前还有若干页沿用该
       节的页眉, 其横线同时作用于两部分, 故 before 与 after 都要计入。
    2) 未写 headerReference 的节沿用前一节的页眉(OOXML 继承规则), 需先解析出
       每节实际生效的 default 引用, 否则会漏判这些节的横线。

    返回 (before, after, i_abs, secs, split_ok, why)。
    """
    i_abs = dx.abstract_idx()
    secs = dx.sects_ranged()
    rid2h = dx.rid_to_header()

    eff, inherited = [], {}
    for s in secs:
        refs = dict(inherited)
        refs.update({k: v for k, v in s['refs'].items() if v})
        eff.append(refs)
        inherited = refs

    before, after, straddle = [], [], False
    for s, refs in zip(secs, eff):
        rid = refs.get('default')
        if not rid:
            continue
        hb = dx.header_bottom(rid2h.get(rid, ''))
        if i_abs < 0 or s['start'] >= i_abs:
            after.append(hb)
        elif s['end'] < i_abs:
            before.append(hb)
        else:
            straddle = True
            before.append(hb)
            after.append(hb)

    # 分节点是否紧接"摘  要"之前: 摘要须是其所在节的第一段(其前只允许空段)
    if i_abs < 0:
        split_ok, why = False, '未定位到"摘  要"标题段'
    else:
        host = next((s for s in secs if s['start'] <= i_abs <= s['end']), None)
        if host is None:
            split_ok, why = False, '未解析到摘要所在的节'
        elif host is secs[0]:
            split_ok, why = False, '摘要仍与封面、声明页同属第一节, 摘要之前未断开分节'
        elif straddle:
            n = i_abs - host['start']
            split_ok = False
            why = (f'分节符提前了 {n} 段: 摘要所在节从第 {host["start"]} 段起, '
                   f'把声明页等前置内容一并并入了"摘要及之后"的节')
        else:
            prev = secs[secs.index(host) - 1]
            ty = prev.get('type') or 'nextPage'
            if ty not in ('nextPage', 'oddPage', 'evenPage'):
                split_ok, why = False, f'摘要前的分节符类型为 {ty}, 摘要未另起一页'
            else:
                split_ok = True
                why = f'分节符紧接"摘  要"之前(摘要为第 {i_abs} 段, 亦为该节首段)'
    return before, after, i_abs, secs, split_ok, why


# ==================== 生效格式解析引擎(段落/文字最终格式) ====================
PKEYS = ('jc', 'line', 'lineRule', 'before', 'after', 'beforeLines', 'afterLines',
         'firstLine', 'firstLineChars', 'left', 'leftChars', 'right', 'rightChars', 'hanging')
RKEYS = ('ea', 'ascii', 'sz', 'b', 'u')

# 不同 Word 语言版本对同一字体存名不同, 归并后比对以免误判
FONT_ALIAS = {
    'SimSun': '宋体', '宋体': '宋体', 'NSimSun': '宋体',
    'SimHei': '黑体', '黑体': '黑体',
    'FangSong': '仿宋', '仿宋': '仿宋', 'KaiTi': '楷体', '楷体': '楷体',
    'FZXiaoBiaoSong-B05S': '方正小标宋简体', '方正小标宋简体': '方正小标宋简体',
    'Times New Roman': 'Times New Roman', 'TimesNewRoman': 'Times New Roman',
}


def fnorm(name):
    if not name:
        return None
    return FONT_ALIAS.get(name.strip(), name.strip())


def _A(e, n):
    return e.get(w(n))


def pPr_dict(pPr):
    d = {}
    if pPr is None:
        return d
    jc = pPr.find(w('jc'))
    if jc is not None:
        d['jc'] = _A(jc, 'val')
    sp = pPr.find(w('spacing'))
    if sp is not None:
        for k in ('line', 'lineRule', 'before', 'after', 'beforeLines', 'afterLines'):
            v = _A(sp, k)
            if v is not None:
                d[k] = v
    ind = pPr.find(w('ind'))
    if ind is not None:
        for k in ('firstLine', 'firstLineChars', 'left', 'leftChars',
                  'right', 'rightChars', 'hanging'):
            v = _A(ind, k)
            if v is not None:
                d[k] = v
    return d
def rPr_dict(rPr):
    d = {}
    if rPr is None:
        return d
    rf = rPr.find(w('rFonts'))
    if rf is not None:
        if _A(rf, 'eastAsia'):
            d['ea'] = _A(rf, 'eastAsia')
        if _A(rf, 'ascii'):
            d['ascii'] = _A(rf, 'ascii')
    sz = rPr.find(w('sz'))
    if sz is not None:
        d['sz'] = _A(sz, 'val')
    b = rPr.find(w('b'))
    if b is not None:
        d['b'] = _A(b, 'val') not in ('0', 'false', 'off')
    u = rPr.find(w('u'))
    if u is not None:
        d['u'] = _A(u, 'val')
    return d


class Fmt:
    """按 docDefaults → 样式 basedOn 链 → 直接 pPr/rPr 的顺序解析生效格式。"""
    def __init__(self, zf, document):
        self.zf = zf
        self.doc = document
        try:
            self.sx = ET.fromstring(zf.read('word/styles.xml'))
        except Exception:
            self.sx = None
        self.styles = {}
        self.dflt_p, self.dflt_r = {}, {}
        self.default_pstyle = None
        if self.sx is None:
            return
        dd = self.sx.find(w('docDefaults'))
        if dd is not None:
            pd = dd.find(w('pPrDefault'))
            if pd is not None:
                self.dflt_p = pPr_dict(pd.find(w('pPr')))
            rd = dd.find(w('rPrDefault'))
            if rd is not None:
                self.dflt_r = rPr_dict(rd.find(w('rPr')))
        for st in self.sx.findall(w('style')):
            sid = _A(st, 'styleId')
            bo = st.find(w('basedOn'))
            self.styles[sid] = {'type': _A(st, 'type'),
                                'basedOn': _A(bo, 'val') if bo is not None else None,
                                'pPr': pPr_dict(st.find(w('pPr'))),
                                'rPr': rPr_dict(st.find(w('rPr')))}
            if _A(st, 'default') in ('1', 'true') and _A(st, 'type') == 'paragraph':
                self.default_pstyle = sid
    def chain(self, sid):
        out, seen = [], set()
        while sid and sid in self.styles and sid not in seen:
            seen.add(sid)
            out.append(sid)
            sid = self.styles[sid]['basedOn']
        return list(reversed(out))

    def eff_p(self, pPr):
        ps = pPr.find(w('pStyle')) if pPr is not None else None
        sid = _A(ps, 'val') if ps is not None else self.default_pstyle
        d = dict(self.dflt_p)
        for s in self.chain(sid):
            d.update(self.styles[s]['pPr'])
        d.update(pPr_dict(pPr))
        return d, sid

    def eff_r(self, sid, rPr):
        d = dict(self.dflt_r)
        for s in self.chain(sid):
            d.update(self.styles[s]['rPr'])
        if rPr is not None:
            rs = rPr.find(w('rStyle'))
            if rs is not None:
                for s in self.chain(_A(rs, 'val')):
                    d.update(self.styles[s]['rPr'])
            d.update(rPr_dict(rPr))
        return d

    def paras(self, root=None):
        """按文档顺序返回段落记录(会进入 w:sdt 与表格)。"""
        root = self.doc.find(w('body')) if root is None else root
        out = []
        counter = [0]

        def walk(node, tbl):
            for ch in list(node):
                if ch.tag == w('tbl'):
                    counter[0] += 1
                    walk(ch, counter[0] - 1)
                elif ch.tag == w('p'):
                    pPr = ch.find(w('pPr'))
                    ep, sid = self.eff_p(pPr)
                    runs = []
                    for r in ch.iter(w('r')):
                        t = ''.join(x.text or '' for x in r.iter(w('t')))
                        if not t.strip():
                            continue
                        runs.append((t, self.eff_r(sid, r.find(w('rPr')))))
                    txt = ''.join(x.text or '' for x in ch.iter(w('t')))
                    out.append({'i': len(out), 't': txt, 'p': ep, 'sid': sid,
                                'runs': runs, 'tbl': tbl,
                                'img': next(ch.iter(w('drawing')), None) is not None})
                else:
                    walk(ch, tbl)

        walk(root, None)
        return out

    def part(self, name):
        """解析 header/footer 等部件内的段落(样式继承同正文)。"""
        try:
            return self.paras(ET.fromstring(self.zf.read(name)))
        except Exception:
            return []
# ==================== 绝对格式比对助手 ====================
TWIPS_CM = 1440 / 2.54     # 1 厘米 = 566.93 twips


def gjc(p):
    return p.get('jc') or 'left'


def gi(d, k, dflt=0):
    try:
        return int(d.get(k) or dflt)
    except (TypeError, ValueError):
        return dflt


def _sp_tw(p, key):
    """段前/段后的生效磅值(twips); 以行为单位的 xxxLines 非 0 时优先。"""
    lines = gi(p, key + 'Lines', 0)
    if lines:
        return int(lines * 240 / 100)
    return gi(p, key, 0)


def _first_chars(p, sz_half):
    """首行缩进的生效值(字符); 只写 w:firstLine 绝对值时按字号折算。"""
    fc = p.get('firstLineChars')
    if fc is not None:
        return int(fc) / 100.0
    per = max((sz_half / 2.0) * 20.0, 1.0)
    return gi(p, 'firstLine', 0) / per


def _first_tw(p, sz_half):
    """首行缩进的生效值(twips); 以字符为单位的 firstLineChars 按字号折算。

    原始文档正文一类段落的缩进本就是绝对值(w:firstLine="480" = 0.85 厘米),
    细则对这些段落写明的也是厘米值, 故按 twips 比对; 若模型改成等值的
    w:firstLineChars(如 WPS 另存会做此归一), 折算后数值相同, 同样视为未改动。
    """
    fc = p.get('firstLineChars')
    if fc is not None:
        return int(fc) / 100.0 * (sz_half / 2.0) * 20.0
    return gi(p, 'firstLine', 0)


def _left_tw(p, sz_half):
    lc = p.get('leftChars')
    if lc is not None:
        return int(lc) / 100.0 * (sz_half / 2.0) * 20.0
    return gi(p, 'left', 0)


def pbad(rec, jc=None, line=None, before=None, after=None, first=None,
         first_cm=None, left=None, right=None, sz_half=24):
    """返回该段落与期望绝对值不符之处的说明列表。

    first    期望首行缩进为 N 字符(文档以 w:firstLineChars 记录的段落);
    first_cm 期望首行缩进为 N 厘米(文档以 w:firstLine 绝对值记录的段落)。
    """
    p, bad = rec['p'], []
    tag = repr(rec['t'].strip()[:10])
    if jc is not None and gjc(p) != jc:
        bad.append(f'{tag} 对齐={gjc(p)}(应{jc})')
    if line is not None:
        rule, val = (p.get('lineRule') or 'auto'), gi(p, 'line', 240)
        if rule != line[0] or abs(val - line[1]) > 4:
            bad.append(f'{tag} 行距={rule}/{val}(应{line[0]}/{line[1]})')
    if before is not None and abs(_sp_tw(p, 'before') - before) > 12:
        bad.append(f"{tag} 段前={_sp_tw(p, 'before')}(应{before} twips)")
    if after is not None and abs(_sp_tw(p, 'after') - after) > 12:
        bad.append(f"{tag} 段后={_sp_tw(p, 'after')}(应{after} twips)")
    if first is not None and abs(_first_chars(p, sz_half) - first) > 0.2:
        bad.append(f'{tag} 首行缩进={_first_chars(p, sz_half):.2f}字符(应{first}字符)')
    if first_cm is not None and abs(_first_tw(p, sz_half) - first_cm * TWIPS_CM) > 24:
        bad.append(f'{tag} 首行缩进={_first_tw(p, sz_half) / TWIPS_CM:.2f}厘米(应{first_cm}厘米)')
    if left is not None and abs(_left_tw(p, sz_half) - left) > 24:
        bad.append(f'{tag} 左缩进={_left_tw(p, sz_half):.0f}(应{left} twips)')
    if right is not None and abs(gi(p, 'right', 0) - right) > 24:
        bad.append(f"{tag} 右缩进={gi(p, 'right', 0)}(应{right} twips)")
    return bad
def rbad(rec, ea=None, asc=None, sz=None, b=None, u=None, pick=None):
    """返回该段落内文字(run)与期望绝对值不符之处的说明列表。"""
    bad = []
    for t, r in rec['runs']:
        if pick is not None and not pick(t):
            continue
        tag = repr(t.strip()[:10])
        if ea is not None and re.search(r'[一-鿿]', t):
            got = fnorm(r.get('ea'))
            if got != fnorm(ea):
                bad.append(f'{tag} 中文字体={got}(应{ea})')
        if asc is not None and re.search(r'[A-Za-z]', t):
            got = fnorm(r.get('ascii'))
            if got != fnorm(asc):
                bad.append(f'{tag} 西文字体={got}(应{asc})')
        if sz is not None and gi(r, 'sz', 20) != sz:
            bad.append(f"{tag} 字号={gi(r, 'sz', 20) / 2:.1f}磅(应{sz / 2:.1f}磅)")
        if b is not None and bool(r.get('b')) != b:
            bad.append(f'{tag} 加粗={bool(r.get("b"))}(应{b})')
        if u is not None:
            got = r.get('u') or 'none'
            if got != u:
                bad.append(f'{tag} 下划线={got}(应{u})')
    return bad


def verdict(label, targets, bad, want_n=None):
    """三态判定: 找不到对象→None(未核查); 有不符→False(扣分); 否则→True。"""
    if not targets:
        return None, f'{label}: 未找到对应段落/对象, 未核查(内容缺失由通用破坏项判定)'
    if want_n is not None and len(targets) != want_n:
        bad = [f'共 {len(targets)} 处(应 {want_n} 处)'] + list(bad)
    if bad:
        more = f' 等共 {len(bad)} 处不符' if len(bad) > 4 else ''
        return False, f'{label}: ' + '; '.join(bad[:4]) + more
    return True, f'{label}: {len(targets)} 处均满足细则写明的绝对格式'


RX_CH = r'^[一二三四]、\S'
RX_SEC = r'^\d\.\d(?!\.)\s'
RX_SUB = r'^\d\.\d\.\d\s'
RX_FIG = r'^图\d-\d'
RX_TBL = r'^表\d-\d'
RX_REF = r'^\[\d+\]'
RX_THESIS = r'^Master.s Degree Thesis$'
CH_TITLES = ['一、研究基础与背景', '二、社区微气候韧性评价体系构建',
             '三、典型社区样本实证分析', '四、更新路径与结论']
# 必须存在的各部分: (名称, 定位正则)
SECT_MUST = ([('中文封面', r'^硕士学位论文$'), ('英文封面', RX_THESIS),
              ('学术成果原创性声明', r'^学术成果原创性声明$'),
              ('学术成果使用授权声明', r'^学术成果使用授权声明$'),
              ('摘要', r'^摘\s*要$'), ('Abstract', r'^Abstract$'), ('目录', r'^目\s*录$')]
             + [(t, '^' + t + '$') for t in CH_TITLES]
             + [('结语', r'^结\s*语$'), ('参考文献', r'^参考文献$'),
                ('附录', r'^附\s*录$'), ('致谢', r'^致\s*谢$')])
class Layout:
    """按细则的分区(封面/声明/摘要/目录/正文/文后)组织段落, 供各减分项定位对象。"""
    def __init__(self, dx):
        self.dx = dx
        self.fmt = Fmt(dx.zf, dx.document)
        self.ps = self.fmt.paras()
        self.txt = [r['t'].strip() for r in self.ps]
        self.i_abs = self.idx(r'^摘\s*要$')
        self.i_kw = self.idx(r'^关键词')
        self.i_en = self.idx(r'^Abstract$')
        self.i_kws = self.idx(r'^Keywords')
        self.i_toc = self.idx(r'^目\s*录$')
        self.i_ch1 = self.idx('^' + CH_TITLES[0] + '$')
        self.i_end = self.idx(r'^结\s*语$')
        self.i_ref = self.idx(r'^参考文献$')
        self.i_apx = self.idx(r'^附\s*录$')
        self.i_thx = self.idx(r'^致\s*谢$')
        self.i_grant = self.idx(r'^学术成果使用授权声明$')

    def idx(self, pat, start=0):
        for i in range(start, len(self.ps)):
            if re.search(pat, self.txt[i]):
                return i
        return -1

    def sel(self, pat, lo=0, hi=None, tbl='any'):
        hi = len(self.ps) if hi is None else hi
        if lo < 0 or hi < 0:
            return []
        out = []
        for r in self.ps[lo:hi]:
            if tbl == 'none' and r['tbl'] is not None:
                continue
            if tbl == 'only' and r['tbl'] is None:
                continue
            if re.search(pat, r['t'].strip()):
                out.append(r)
        return out

    def tbl(self, i):
        return [r for r in self.ps if r['tbl'] == i]

    def tbl_rows(self, i):
        """把第 i 个表格的单元格段落按行分组(依据行内单元格数由 shape 给出)。"""
        cells = self.tbl(i)
        shapes = self.dx.tbl_shapes()
        if i >= len(shapes):
            return []
        rows, k = [], 0
        for ncol in shapes[i][1]:
            rows.append(cells[k:k + ncol])
            k += ncol
        return rows

    def headers(self):
        """被各节引用的页眉: [(类型, 部件名, 段落记录, 所属节)]。"""
        rid2h = self.dx.rid_to_header()
        out = []
        for s in self.dx.sects_ranged():
            for ty, rid in s['refs'].items():
                part = rid2h.get(rid)
                if part:
                    out.append((ty, part, self.fmt.part(part), s))
        return out

    def sect_of(self, i):
        """覆盖段落序号 i 的节。"""
        for s in self.dx.sects_ranged():
            if s['start'] <= i <= s['end']:
                return s
        return None

    def sect_kinds(self):
        """各节 + 其所属部分: margin(front/main) 与页码格式(roman/arabic)。

        以节的起始段落落在哪一部分判定, 与文档实际分节数量无关: 起始段落在
        "Abstract"之前的节属封面部分(左右各 31.8 毫米), 其余属正文部分。
        """
        out = []
        for s in self.dx.sects_ranged():
            mar = 'front' if (self.i_en < 0 or s['start'] < self.i_en) else 'main'
            num = 'roman' if (self.i_ch1 >= 0 and s['start'] < self.i_ch1) else 'arabic'
            out.append((s, mar, num))
        return out

    def find_tbl(self, first_cell):
        """按表头首格文字定位表格序号(表格被增删时仍能对上)。"""
        for i in range(len(self.dx.tbl_shapes())):
            cells = self.tbl(i)
            if cells and cells[0]['t'].strip() == first_cell:
                return i
        return None

    def between(self, lo, hi):
        """区间内的非空普通段落(不含表格单元格)。"""
        if lo < 0 or hi < 0 or hi <= lo:
            return []
        return [r for r in self.ps[lo + 1:hi] if r['t'].strip() and r['tbl'] is None]

    def body_paras(self):
        """正文各章的普通段落(排除标题、图表题注、表格单元格)。"""
        if self.i_ch1 < 0:
            return []
        hi = self.i_end if self.i_end > 0 else len(self.ps)
        out = []
        for r in self.ps[self.i_ch1:hi]:
            t = r['t'].strip()
            if r['tbl'] is not None or not t:
                continue
            if re.search(RX_CH, t) or re.search(RX_SEC, t) or re.search(RX_SUB, t):
                continue
            if re.search(RX_FIG, t) or re.search(RX_TBL, t):
                continue
            out.append(r)
        return out
# ==================== 减分项(55 条, 与《细则.docx》"二、减分项"逐条对应) ====================
RULES = []


def rule(code, pts):
    def deco(fn):
        RULES.append((code, pts, fn))
        return fn
    return deco


def has_u(recs):
    """每段是否都有加单下划线的取值 run。"""
    bad = []
    for r in recs:
        if not any((x[1].get('u') or 'none') == 'single' for x in r['runs']):
            bad.append(f"{r['t'].strip()[:10]!r} 取值未加单下划线")
    return bad


def near(a, b, tol):
    return a is not None and b is not None and abs(a - b) <= tol


# ---------- 【中文封面】 ----------
@rule('-1 中文封面"密级/分类号/机构代码/编号"四行格式被改', -1)
def d01(dx, L):
    t = L.sel(r'^(密\s*级|分\s*类\s*号|机构代码|编\s*号)\s*[：:]', tbl='none')
    bad = []
    for r in t:
        bad += pbad(r, jc='left', line=('auto', 360), first=1.5, sz_half=24)
        bad += rbad(r, ea='宋体', asc='宋体', sz=24, b=False)
    return verdict('中文封面密级等四行', t, bad, want_n=4)


@rule('-1 中文封面"硕士学位论文"格式被改', -1)
def d02(dx, L):
    t = L.sel(r'^硕士学位论文$')
    bad = []
    for r in t:
        bad += pbad(r, jc='left', line=('auto', 360), first=1, sz_half=72)
        bad += rbad(r, ea='黑体', sz=72, b=False)
    return verdict('中文封面"硕士学位论文"', t, bad)


@rule('-1 中文封面题名格式被改', -1)
def d03(dx, L):
    t = L.sel(r'^城市社区微气候韧性评价与公共空间更新研究$')
    bad = []
    for r in t:
        bad += pbad(r, jc='left', line=('auto', 360), sz_half=52)
        bad += rbad(r, ea='方正小标宋简体', sz=52, b=False)
    return verdict('中文封面题名', t, bad)
@rule('-1 中文封面校名图形被删除/替换/缩放', -1)
def d04(dx, L):
    n, bad_embed, sizes = dx.img_info()
    if not sizes:
        return False, '未检出任何内嵌图形, 封面校名图形已被删除'
    got = sizes[0]
    if bad_embed:
        return False, f'图片关系引用失效 {bad_embed[:2]}, 图形无法显示'
    if not (near(got[0], IMG_CM[0][0], 0.2) and near(got[1], IMG_CM[0][1], 0.2)):
        return False, (f'封面首个图形 {got[0]}×{got[1]} 厘米, '
                       f'与 {IMG_CM[0][0]}×{IMG_CM[0][1]} 厘米不符, 已被缩放或替换')
    return True, f'封面校名图形 {got[0]}×{got[1]} 厘米, 嵌入式版式与尺寸均满足'


@rule('-1 中文封面申请人/指导教师/学位类别三行格式被改', -1)
def d05(dx, L):
    t = L.sel(r'^(学位申请人|指\s*导\s*教\s*师|学\s*位\s*类\s*别)\s*[：:]')
    bad = []
    for r in t:
        bad += pbad(r, jc='both', line=('auto', 360), first=3, sz_half=28)
        bad += rbad(r, ea='黑体', asc='Times New Roman', sz=28, b=False)
    bad += has_u(t)
    return verdict('中文封面申请人信息三行', t, bad, want_n=3)


@rule('-1 中文封面 5 行 1 列无框线表格结构被改变', -1)
def d06(dx, L):
    shapes = dx.tbl_shapes()
    if not shapes:
        return False, '文档内已无表格, 承载申请人信息的 5 行 1 列表格被删除或拆为普通段落'
    if shapes[0] != TBL_SHAPE[0]:
        return False, f'第 1 个表格为 {shapes[0]}, 与 5 行 1 列不符'
    bd = dx.tbl_borders(0)
    vis = {k: v for k, v in bd.items() if v not in (None, 'nil', 'none')}
    if vis:
        return False, f'该表被加上可见框线 {vis}'
    return True, '中文封面 5 行 1 列无框线表格结构完整、无可见框线'


# ---------- 【英文封面】 ----------
@rule('-1 英文封面 Classified Index 等四行格式被改', -1)
def d07(dx, L):
    t = L.sel(r'^(Classified Index|U\.D\.C\.|CODE|No\.)')
    bad = []
    for r in t:
        bad += pbad(r, jc='left', line=('auto', 360), sz_half=24)
        bad += rbad(r, asc='Times New Roman', sz=24, b=True)
    return verdict('英文封面 Classified Index 等四行', t, bad, want_n=4)
@rule("-1 英文封面 Master's Degree Thesis 格式被改", -1)
def d08(dx, L):
    t = L.sel(RX_THESIS)
    bad = []
    for r in t:
        bad += pbad(r, line=('auto', 191), before=204, left=814, sz_half=71)
        bad += rbad(r, asc='Times New Roman', sz=71, b=True)
    return verdict("英文封面 Master's Degree Thesis", t, bad)


@rule('-1 英文封面题名格式被改', -1)
def d09(dx, L):
    t = L.sel(r'^Research on Microclimate Resilience')
    bad = []
    for r in t:
        bad += pbad(r, jc='center', line=('auto', 240), before=150, sz_half=52)
        bad += rbad(r, asc='Times New Roman', sz=52, b=True)
    return verdict('英文封面题名', t, bad)


@rule('-1 英文封面 Applicant/Supervisor/Degree Category 三行格式被改', -1)
def d10(dx, L):
    t = L.sel(r'^(Applicant|Supervisor|Degree Category)\s*[：:]')
    bad = []
    for r in t:
        bad += pbad(r, jc='both', line=('auto', 360), sz_half=28)
        bad += rbad(r, asc='Times New Roman', sz=28, b=True)
    bad += has_u(t)
    return verdict('英文封面申请人信息三行', t, bad, want_n=3)


@rule('-1 英文封面三行首行缩进或 3 行 1 列无框线表格结构被改变', -1)
def d11(dx, L):
    t = L.sel(r'^(Applicant|Supervisor|Degree Category)\s*[：:]')
    bad = []
    for r in t:
        want = 1 if r['t'].strip().startswith('Degree Category') else 4
        bad += pbad(r, first=want, sz_half=28)
    shapes = dx.tbl_shapes()
    if len(shapes) < 2:
        bad.append('3 行 1 列表格缺失或已被拆为普通段落')
    elif shapes[1] != TBL_SHAPE[1]:
        bad.append(f'第 2 个表格为 {shapes[1]}, 与 3 行 1 列不符')
    else:
        vis = {k: v for k, v in dx.tbl_borders(1).items()
               if v not in (None, 'nil', 'none')}
        if vis:
            bad.append(f'该表被加上可见框线 {vis}')
    return verdict('英文封面三行缩进与 3 行 1 列表格', t, bad, want_n=3)
# ---------- 【学术成果原创性声明与使用授权声明页】 ----------
@rule('-1 声明页"澄湖理工大学"格式被改', -1)
def d12(dx, L):
    t = L.sel(r'^澄湖理工大学$')
    bad = []
    for r in t:
        bad += pbad(r, jc='center', line=('auto', 480), first=2, sz_half=36)
        bad += rbad(r, ea='宋体', asc='Times New Roman', sz=36, b=True)
    return verdict('声明页"澄湖理工大学"', t, bad)


@rule('-1 "学术成果原创性声明"标题格式被改', -1)
def d13(dx, L):
    t = L.sel(r'^学术成果原创性声明$')
    bad = []
    for r in t:
        bad += pbad(r, jc='center', line=('auto', 480), sz_half=32)
        bad += rbad(r, sz=32, b=True)
    return verdict('"学术成果原创性声明"标题', t, bad)


@rule('-1 原创性声明正文格式被改', -1)
def d14(dx, L):
    t = L.sel(r'^本人郑重确认')
    bad = []
    for r in t:
        bad += pbad(r, jc='both', line=('auto', 360), first=2, sz_half=24)
        bad += rbad(r, ea='宋体', asc='Times New Roman', sz=24, b=False)
    return verdict('原创性声明正文', t, bad)


@rule('-1 声明页签名行格式被改', -1)
def d15(dx, L):
    t = L.sel(r'^(撰写人签名|指导人签名)')
    bad = []
    for r in t:
        i = r['i']
        sz = 21 if (L.i_grant < 0 or i < L.i_grant) else 28
        bad += pbad(r, line=('auto', 480), first=0, left=0, sz_half=sz)
        bad += rbad(r, ea='宋体', sz=sz, b=False,
                    pick=lambda s: s.strip() not in ('年', '月', '日'))
    return verdict('声明页签名行', t, bad, want_n=3)


@rule('-1 "学术成果使用授权声明"标题格式被改', -1)
def d16(dx, L):
    t = L.sel(r'^学术成果使用授权声明$')
    bad = []
    for r in t:
        bad += pbad(r, jc='center', line=('exact', 460), sz_half=32)
        bad += rbad(r, sz=32, b=True)
    return verdict('"学术成果使用授权声明"标题', t, bad)


@rule('-1 授权声明两段正文格式被改', -1)
def d17(dx, L):
    t = L.sel(r'^(本人了解学校关于|本论文可在规定范围内)')
    bad = []
    for r in t:
        bad += pbad(r, line=('auto', 480), first_cm=0.85, sz_half=24)
        bad += rbad(r, ea='宋体', sz=24, b=False)
    return verdict('授权声明两段正文', t, bad, want_n=2)
# ---------- 【摘要与 Abstract】 ----------
@rule('-1 "摘  要"标题格式或字间两空格被改', -1)
def d18(dx, L):
    t = L.sel(r'^摘\s*要$')
    bad = []
    for r in t:
        if r['t'].strip() != '摘  要':
            bad.append(f"标题为 {r['t'].strip()!r}, 未保留\"摘\"\"要\"之间的两个空格")
        bad += pbad(r, jc='center', sz_half=36)
        bad += rbad(r, ea='黑体', asc='Times New Roman', sz=36, b=False)
    return verdict('"摘  要"标题', t, bad)


@rule('-1 摘要两段正文格式被改', -1)
def d19(dx, L):
    t = L.between(L.i_abs, L.i_kw)
    bad = []
    for r in t:
        bad += pbad(r, jc='both', line=('auto', 360), first=2, sz_half=24)
        bad += rbad(r, ea='宋体', asc='Times New Roman', sz=24, b=False)
    return verdict('摘要两段正文', t, bad, want_n=2)


@rule('-1 中文关键词行格式或词条间两空格被改', -1)
def d20(dx, L):
    t = L.sel(r'^关键词')
    bad = []
    for r in t:
        bad += pbad(r, jc='both', sz_half=24)
        bad += rbad(r, sz=24, b=False)
        bad += rbad(r, ea='黑体', pick=lambda s: s.strip().startswith('关键词'))
        bad += rbad(r, ea='宋体', pick=lambda s: not s.strip().startswith('关键词'))
        if len(re.findall(r'\S {2}\S', r['t'])) < 4:
            bad.append('词条之间未保留两个空格')
    return verdict('中文关键词行', t, bad)


@rule('-1 "Abstract"标题格式被改', -1)
def d21(dx, L):
    t = L.sel(r'^Abstract$')
    bad = []
    for r in t:
        bad += pbad(r, jc='center', sz_half=36)
        bad += rbad(r, asc='Times New Roman', sz=36, b=True)
    return verdict('"Abstract"标题', t, bad)


@rule('-1 Abstract 三段正文格式被改', -1)
def d22(dx, L):
    t = L.between(L.i_en, L.i_kws)
    bad = []
    for r in t:
        bad += pbad(r, jc='both', line=('auto', 360), first=2, sz_half=24)
        bad += rbad(r, asc='Times New Roman', sz=24, b=False)
    return verdict('Abstract 三段正文', t, bad, want_n=3)


@rule('-1 Keywords 行格式或词条间两空格被改', -1)
def d23(dx, L):
    t = L.sel(r'^Keywords')
    bad = []
    for r in t:
        bad += rbad(r, asc='Times New Roman', sz=24)
        bad += rbad(r, b=True, pick=lambda s: s.strip().startswith('Keywords'))
        bad += rbad(r, b=False, pick=lambda s: not s.strip().startswith('Keywords'))
        if len(re.findall(r'\S {2}\S', r['t'])) < 4:
            bad.append('词条之间未保留两个空格')
    return verdict('Keywords 行', t, bad)
# ---------- 【目录与前置内容控件】 ----------
RX_TOC1 = r'^(摘\s+要|Abstract|[一二三四]、|结\s+语|参考文献|附\s+录|致\s+谢)'


@rule('-1 "目  录"标题格式或字间两空格被改', -1)
def d24(dx, L):
    t = L.sel(r'^目\s*录$')
    bad = []
    for r in t:
        if r['t'].strip() != '目  录':
            bad.append(f"标题为 {r['t'].strip()!r}, 未保留\"目\"\"录\"之间的两个空格")
        bad += pbad(r, jc='center', sz_half=32)
        bad += rbad(r, ea='黑体', asc='Times New Roman', sz=32, b=True)
    return verdict('"目  录"标题', t, bad)


@rule('-1 目录 10 条一级条目格式被改', -1)
def d25(dx, L):
    t = L.sel(RX_TOC1, L.i_toc + 1, L.i_ch1, tbl='none') if L.i_toc >= 0 else []
    bad = []
    for r in t:
        bad += pbad(r, jc='both', line=('auto', 360), left=0, sz_half=24)
        bad += rbad(r, ea='宋体', asc='Times New Roman', sz=24, b=True,
                    pick=lambda s: not s.strip().isdigit())
    return verdict('目录一级条目', t, bad, want_n=10)


@rule('-1 目录 12 条二级条目格式被改', -1)
def d26(dx, L):
    t = L.sel(RX_SEC, L.i_toc + 1, L.i_ch1, tbl='none') if L.i_toc >= 0 else []
    bad = []
    for r in t:
        bad += pbad(r, jc='both', line=('auto', 360), left=420, sz_half=24)
        bad += rbad(r, ea='宋体', sz=24, b=False)
    return verdict('目录二级条目', t, bad, want_n=12)


@rule('-1 目录条目数量不足 22 行或条目文字/页码数字被改写', -1)
def d27(dx, L):
    if L.i_toc < 0 or L.i_ch1 < 0:
        return None, '未定位到目录区, 未核查'
    t = [r for r in L.ps[L.i_toc + 1:L.i_ch1]
         if r['t'].strip() and r['tbl'] is None]
    bad = []
    if len(t) != 22:
        bad.append(f'目录共 {len(t)} 行(应 22 行: 10 条一级 + 12 条二级)')
    nopage = [r['t'].strip()[:14] for r in t if not re.search(r'\d\s*$', r['t'].strip())]
    if nopage:
        bad.append(f'条目末尾页码数字缺失: {nopage[:3]}')
    return verdict('目录条目完整性', t, bad)


@rule('-1 前置内容未整体保留在 1 个内容控件(SDT)中', -1)
def d28(dx, L):
    total, broken = dx.sdt_info()
    if total < 1:
        return False, '内容控件(SDT)已被删除或解除为普通段落, 前置版式被拆散'
    if broken:
        return False, f'{total} 个内容控件中 {broken} 个内容为空, 前置版式被破坏'
    return True, f'前置内容仍完整保留在 {total} 个内容控件中'
# ---------- 【正文各级标题】 ----------
@rule('-1 4 个章标题格式被改', -1)
def d29(dx, L):
    t = L.sel(RX_CH, L.i_ch1, L.i_end if L.i_end > 0 else None, tbl='none')
    bad = []
    for r in t:
        bad += pbad(r, jc='center', line=('auto', 360), before=240, after=240, sz_half=32)
        bad += rbad(r, ea='黑体', asc='Times New Roman', sz=32, b=True)
    return verdict('章标题', t, bad, want_n=4)


@rule('-1 12 个节标题格式被改', -1)
def d30(dx, L):
    t = L.sel(RX_SEC, L.i_ch1, L.i_end if L.i_end > 0 else None, tbl='none')
    bad = []
    for r in t:
        bad += pbad(r, jc='left', line=('auto', 360), before=120, after=120,
                    first=0, sz_half=28)
        bad += rbad(r, ea='黑体', sz=28, b=True)
    return verdict('节标题', t, bad, want_n=12)


@rule('-1 17 个三级标题格式被改', -1)
def d31(dx, L):
    t = L.sel(RX_SUB, L.i_ch1, L.i_end if L.i_end > 0 else None, tbl='none')
    bad = []
    for r in t:
        bad += pbad(r, jc='left', line=('auto', 360), before=60, after=60,
                    first=0, sz_half=24)
        bad += rbad(r, ea='黑体', sz=24, b=True)
    return verdict('三级标题', t, bad, want_n=17)


@rule('-1 "结  语""参考文献""附  录""致  谢"四个标题格式被改', -1)
def d32(dx, L):
    t = []
    for pat, want in ((r'^结\s*语$', '结  语'), (r'^参考文献$', '参考文献'),
                      (r'^附\s*录$', '附  录'), (r'^致\s*谢$', '致  谢')):
        t += [(r, want) for r in L.sel(pat, L.i_end if L.i_end > 0 else 0, None, 'none')]
    bad = []
    for r, want in t:
        if r['t'].strip() != want:
            bad.append(f"标题为 {r['t'].strip()!r}, 未保留原有的两个空格")
        bad += pbad(r, jc='center', line=('auto', 360), before=240, after=240,
                    first=0, sz_half=32)
        bad += rbad(r, ea='黑体', sz=32, b=True)
    return verdict('文后四个标题', t, bad, want_n=4)
# ---------- 【正文段落与参考文献、附录、致谢】 ----------
@rule('-1 正文各章段落格式被改', -1)
def d33(dx, L):
    t = L.body_paras()
    bad = []
    for r in t:
        bad += pbad(r, jc='both', line=('auto', 360), first_cm=0.85,
                    before=0, after=0, sz_half=24)
        bad += rbad(r, ea='宋体', asc='Times New Roman', sz=24, b=False)
    return verdict('正文各章段落', t, bad)


@rule('-1 参考文献 18 条格式被改或序号被重排', -1)
def d34(dx, L):
    t = L.sel(RX_REF, L.i_ref, L.i_apx if L.i_apx > 0 else None, tbl='none')
    bad = []
    nums = [int(re.match(r'^\[(\d+)\]', r['t'].strip()).group(1)) for r in t]
    if nums != list(range(1, 19)):
        bad.append(f'序号为 {nums[:6]}…, 与 [1]~[18] 连续编号不符')
    for r in t:
        bad += pbad(r, jc='both', line=('auto', 360), first_cm=0.85, sz_half=24)
        bad += rbad(r, ea='宋体', asc='Times New Roman', sz=24, b=False)
    return verdict('参考文献条目', t, bad, want_n=18)


@rule('-1 "附录A""附录B"两行标题格式被改', -1)
def d35(dx, L):
    t = L.sel(r'^附录[AB]')
    bad = []
    for r in t:
        bad += pbad(r, jc='both', line=('auto', 360), first_cm=0.85,
                    before=0, after=0, sz_half=24)
        bad += rbad(r, ea='宋体', sz=24, b=True)
    return verdict('附录 A/B 标题行', t, bad, want_n=2)


@rule('-1 附录 A、附录 B 之下的说明段落格式被改', -1)
def d36(dx, L):
    hi = L.i_thx if L.i_thx > 0 else len(L.ps)
    t = [r for r in L.between(L.i_apx, hi) if not re.match(r'^附录[AB]', r['t'].strip())]
    bad = []
    for r in t:
        bad += pbad(r, jc='both', line=('auto', 360), first_cm=0.85, sz_half=24)
        bad += rbad(r, ea='宋体', sz=24, b=False)
    return verdict('附录说明段落', t, bad, want_n=2)


@rule('-1 致谢段落格式被改', -1)
def d37(dx, L):
    t = L.between(L.i_thx, len(L.ps))
    bad = []
    for r in t:
        bad += pbad(r, jc='both', line=('auto', 360), first_cm=0.85, sz_half=24)
        bad += rbad(r, ea='宋体', sz=24, b=False)
    return verdict('致谢段落', t, bad, want_n=2)
# ---------- 【图表与题注】 ----------
@rule('-1 三条图题格式/位置/序名间两空格被改', -1)
def d38(dx, L):
    t = L.sel(RX_FIG, tbl='none')
    bad = []
    for r in t:
        if not re.match(r'^图\d-\d {2}\S', r['t'].strip()):
            bad.append(f"{r['t'].strip()[:12]!r} 图序与图名之间未空两格")
        prev = [x for x in L.ps[:r['i']] if x['t'].strip() or x['img']]
        if not (prev and prev[-1]['img']):
            bad.append(f"{r['t'].strip()[:12]!r} 未位于所属图形的下方")
        bad += pbad(r, jc='center', sz_half=21)
        bad += rbad(r, ea='宋体', asc='Times New Roman', sz=21)
    return verdict('三条图题', t, bad, want_n=3)


@rule('-1 两条表题格式/位置/序名间两空格被改', -1)
def d39(dx, L):
    t = L.sel(RX_TBL, tbl='none')
    bad = []
    for r in t:
        if not re.match(r'^表\d-\d {2}\S', r['t'].strip()):
            bad.append(f"{r['t'].strip()[:12]!r} 表序与表名之间未空两格")
        nxt = [x for x in L.ps[r['i'] + 1:] if x['t'].strip() or x['tbl'] is not None]
        if not (nxt and nxt[0]['tbl'] is not None):
            bad.append(f"{r['t'].strip()[:12]!r} 未位于所属表格的正上方")
        bad += pbad(r, jc='center', sz_half=21)
        bad += rbad(r, ea='宋体', sz=21)
    return verdict('两条表题', t, bad, want_n=2)


@rule('-1 三幅正文插图被删除/替换/缩放或引用失效', -1)
def d40(dx, L):
    n, bad_embed, sizes = dx.img_info()
    if len(sizes) < 4:
        return False, f'仅检出 {len(sizes)} 处图形(应 4 处: 封面 1 幅 + 正文 3 幅), 插图被删除'
    if bad_embed:
        return False, f'图片关系引用失效 {bad_embed[:2]}, 插图显示为红叉'
    bad = []
    for k in (1, 2, 3):
        want, got = IMG_CM[k], sizes[k]
        if not (near(got[0], want[0], 0.2) and near(got[1], want[1], 0.2)):
            bad.append(f'第 {k} 幅正文插图 {got[0]}×{got[1]} 厘米(应 {want[0]}×{want[1]} 厘米)')
    if bad:
        return False, '正文插图尺寸被改: ' + '; '.join(bad)
    return True, '三幅正文插图尺寸与引用均完好'
GRID = ('top', 'left', 'bottom', 'right', 'insideH', 'insideV')
GRID_FULL = dict.fromkeys(GRID, 'single')
GRID_TB = {'top': 'single', 'bottom': 'single', 'left': 'none',
           'right': 'none', 'insideH': 'none', 'insideV': 'none'}
NAME_CN = {'top': '上框线', 'left': '左框线', 'bottom': '下框线',
           'right': '右框线', 'insideH': '内部横线', 'insideV': '内部竖线'}


def grid_bad(dx, ti, want=None):
    """表格框线与细则写明的绝对值(单线/无线)逐条比对。"""
    want, bd, out = want or GRID_FULL, dx.tbl_borders(ti), []
    for k in GRID:
        got = bd.get(k) or 'none'
        if got != want[k]:
            out.append(f'{NAME_CN[k]}={got}'
                       f"(应 {'single 单线' if want[k] == 'single' else 'none 无线'})")
    return out


def cells_bad(rows, sz, line, before=None, after=None, ea='宋体', asc=None):
    bad = []
    for ri, row in enumerate(rows):
        for c in row:
            bad += pbad(c, jc='center' if ri == 0 else 'left', line=line,
                        before=before, after=after, sz_half=sz)
            bad += rbad(c, ea=ea, asc=asc, sz=sz, b=(ri == 0))
    return bad


@rule('-1 "表1-1"表格结构或表内文字格式被改', -1)
def d41(dx, L):
    ti = L.find_tbl('阶段')
    if ti is None:
        return None, '未定位到"表1-1"(表头首格"阶段"), 未核查'
    rows = L.tbl_rows(ti)
    bad = []
    if dx.tbl_shapes()[ti] != TBL_SHAPE[2]:
        bad.append(f'表格为 {dx.tbl_shapes()[ti]}(应 5 行 3 列)')
    hdr = [c['t'].strip() for c in rows[0]] if rows else []
    if hdr != ['阶段', '主要任务', '阶段成果']:
        bad.append(f'表头为 {hdr}(应 阶段/主要任务/阶段成果)')
    bad += grid_bad(dx, ti)
    bad += cells_bad(rows, 21, ('auto', 312), ea='宋体', asc='Times New Roman')
    return verdict('"表1-1"表格', rows, bad, want_n=5)


@rule('-1 "表2-1"表格结构、权重取值或表内文字格式被改', -1)
def d42(dx, L):
    ti = L.find_tbl('评价指标')
    if ti is None:
        return None, '未定位到"表2-1"(表头首格"评价指标"), 未核查'
    rows = L.tbl_rows(ti)
    bad = []
    if dx.tbl_shapes()[ti] != TBL_SHAPE[3]:
        bad.append(f'表格为 {dx.tbl_shapes()[ti]}(应 6 行 3 列)')
    hdr = [c['t'].strip() for c in rows[0]] if rows else []
    if hdr != ['评价指标', '观测与判定要点', '建议权重']:
        bad.append(f'表头为 {hdr}(应 评价指标/观测与判定要点/建议权重)')
    wts = [row[2]['t'].strip() for row in rows[1:] if len(row) > 2]
    if wts != ['25%', '20%', '15%', '20%', '20%']:
        bad.append(f'建议权重列为 {wts}(应 25%/20%/15%/20%/20%)')
    bad += grid_bad(dx, ti, GRID_TB)
    bad += cells_bad(rows, 19, ('auto', 240), before=0, after=0, ea='宋体')
    return verdict('"表2-1"表格', rows, bad, want_n=6)
# ---------- 【页面设置】 ----------
@rule('-1 纸张不再是 A4', -1)
def d43(dx, L):
    secs = dx.sects_ranged()
    bad = [f"第 {i + 1} 节 pgSz={s['pgSz']}(应 {PG_SZ} twips 即 210×297 毫米)"
           for i, s in enumerate(secs)
           if not (s['pgSz'] and near(s['pgSz'][0], PG_SZ[0], 20)
                   and near(s['pgSz'][1], PG_SZ[1], 20))]
    return verdict('纸张', secs, bad)


@rule('-1 页边距被改', -1)
def d44(dx, L):
    kinds = L.sect_kinds()
    bad = []
    for i, (s, mar, _) in enumerate(kinds):
        want = PG_MAR_FRONT if mar == 'front' else PG_MAR_MAIN
        got = s['pgMar']
        for k in ('top', 'bottom', 'left', 'right'):
            if not near(got.get(k), want[k], 12):
                bad.append(f'第 {i + 1} 节'
                           f'({"封面至中文摘要" if mar == "front" else "英文摘要及之后"})'
                           f' {k}={got.get(k)}(应 {want[k]} twips)')
    return verdict('页边距', kinds, bad)


@rule('-1 页眉/页脚距边界被改', -1)
def d45(dx, L):
    secs = dx.sects_ranged()
    bad = []
    for i, s in enumerate(secs):
        if not near(s['pgMar'].get('header'), PG_HDR_FTR[0], 12):
            bad.append(f"第 {i + 1} 节 页眉距边界={s['pgMar'].get('header')}"
                       f'(应 {PG_HDR_FTR[0]} twips 即 15 毫米)')
        if not near(s['pgMar'].get('footer'), PG_HDR_FTR[1], 12):
            bad.append(f"第 {i + 1} 节 页脚距边界={s['pgMar'].get('footer')}"
                       f'(应 {PG_HDR_FTR[1]} twips 即 17.5 毫米)')
    return verdict('页眉/页脚距边界', secs, bad)


# ---------- 【页眉文字与页脚页码】 ----------
@rule('-1 页眉文字被改写/删除/加粗/改对齐', -1)
def d46(dx, L):
    hs = [(ty, part, ps) for ty, part, ps, s in L.headers()
          if ty != 'first' and any(r['t'].strip() for r in ps)]
    bad = []
    for ty, part, ps in hs:
        for r in ps:
            if not r['t'].strip():
                continue
            if r['t'].strip() != HDR_TEXT:
                bad.append(f"{part.split('/')[-1]} 页眉文字={r['t'].strip()[:16]!r}(应{HDR_TEXT!r})")
            bad += pbad(r, jc='center', sz_half=21)
            bad += rbad(r, ea='宋体', asc='Times New Roman', b=False)
    return verdict('页眉文字', hs, bad)
@rule('-1 封面首页页眉出现文字(该页页眉应保持空白)', -1)
def d47(dx, L):
    secs = dx.sects_ranged()
    if not secs:
        return None, '未解析到分节信息, 未核查'
    rid2h = dx.rid_to_header()
    rid = secs[0]['refs'].get('first')
    if not rid:
        return False, '首节未设置"首页不同"页眉, 封面首页将显示与其他页相同的页眉文字'
    ps = L.fmt.part(rid2h.get(rid, ''))
    txt = ''.join(r['t'] for r in ps).strip()
    if txt:
        return False, f'封面首页页眉出现文字 {txt[:20]!r}, 该页页眉应保持空白'
    return True, '封面首页页眉保持空白(横线形态由加分项判定)'


@rule('-1 页脚文字/对齐/缩进被改', -1)
def d48(dx, L):
    fi = dx.footer_info()
    parts = [p for p in fi['parts'] if p['text'].strip()]
    bad = []
    for p in parts:
        if FTR_TEXT not in p['text']:
            bad.append(f"{p['name']} 页脚文字={p['text'][:18]!r}(应含{FTR_TEXT!r})")
        if 18 not in p['sz']:
            bad.append(f"{p['name']} 字号={[s / 2 for s in p['sz']]}磅(应含 9 磅)")
    got = sorted((p['jc'] or 'left', 'left' if p['left'] else 'right',
                  p['left'] or p['right'] or 0) for p in parts)
    want = sorted(FTR_ALIGN)
    if got != want:
        bad.append(f'四个页脚的对齐/缩进为 {got}(应 {want})')
    return verdict('页脚文字与对齐缩进', parts, bad, want_n=4)


@rule('-1 页脚页码不再由 PAGE 域自动生成', -1)
def d49(dx, L):
    fi = dx.footer_info()
    if fi['missing']:
        return False, f"引用的页脚部件 {fi['missing']} 缺失, 页脚被删除"
    if fi['begin'] != fi['end']:
        return False, (f"域标记不配对(begin={fi['begin']}, end={fi['end']}), "
                       f'页码将显示为错误代码')
    if fi['page'] != N_PAGEFLD:
        return False, f"PAGE 域 {fi['page']} 个(应 {N_PAGEFLD} 个), 页码被删除或改为静态数字"
    return True, f"PAGE 域 {fi['page']} 个, begin/separate/end 标记配对完好"


@rule('-1 页码编号格式(罗马数字/阿拉伯数字)被改', -1)
def d50(dx, L):
    kinds = L.sect_kinds()
    bad, seen = [], None
    for i, (s, _, num) in enumerate(kinds):
        got = s['fmt'] or seen
        want = 'upperRoman' if num == 'roman' else 'decimal'
        if got is not None and got != want:
            bad.append(f'第 {i + 1} 节编号={got}(应 {want})')
        seen = got
    return verdict('页码编号格式', kinds, bad)
# ---------- 【通用破坏】 ----------
@rule('-3 文字内容被增删或改写', -3)
def d51(dx, L):
    n = dx.text_len()
    tol = max(1, int(TEXTLEN * TEXT_TOL))
    if abs(n - TEXTLEN) > tol:
        return False, (f'可见文字 {n} 字, 超出 {TEXTLEN}±{tol} 字范围, '
                       f'属非格式性的内容增删或改写')
    return True, f'可见文字 {n} 字, 在 {TEXTLEN}±{tol} 字范围内, 内容未被增删改写'


@rule('-3 图片、表格等对象被删除或破坏', -3)
def d52(dx, L):
    n_img, bad_embed, sizes = dx.img_info()
    shapes = dx.tbl_shapes()
    bad = []
    if bad_embed:
        bad.append(f'图片关系引用失效 {bad_embed[:2]}')
    if n_img != N_IMG:
        bad.append(f'内嵌图片 {n_img} 张(应 {N_IMG} 张)')
    if shapes != TBL_SHAPE:
        bad.append(f'表格行列结构 {shapes}(应 {TBL_SHAPE})')
    if bad:
        return False, '对象被删除或破坏: ' + '; '.join(bad)
    return True, f'内嵌图片 {n_img} 张、表格 {len(shapes)} 个行列结构均完好'


@rule('-3 段落或整节内容被删除导致内容缺失', -3)
def d53(dx, L):
    bad = []
    miss = [name for name, pat in SECT_MUST if L.idx(pat) < 0]
    if miss:
        bad.append(f'缺失部分: {miss}')
    hi = L.i_end if L.i_end > 0 else None
    for pat, want, label in ((RX_CH, 4, '章标题'), (RX_SEC, 12, '节标题'),
                             (RX_SUB, 17, '三级标题')):
        got = len(L.sel(pat, L.i_ch1, hi, tbl='none')) if L.i_ch1 >= 0 else 0
        if got != want:
            bad.append(f'{label} {got} 个(应 {want} 个)')
    if bad:
        return False, '内容缺失: ' + '; '.join(bad)
    return True, '各部分齐全, 章 4 个、节 12 个、三级标题 17 个数量未变'


@rule('-1 姓名、编号、日期等信息的数值被改动', -1)
def d54(dx, L):
    text = dx.body_text()
    miss = [v for v in KEEP_VALUES if v.replace(' ', '') not in text.replace(' ', '')]
    if miss:
        return False, f'以下信息数值已被改动或删除: {miss}'
    return True, f'{len(KEEP_VALUES)} 项姓名/编号等信息数值均与细则写明的取值一致'


@rule('-1 文档出现指向缺失部件的悬挂引用', -1)
def d55(dx, L):
    dang = dx.dangling()
    fi = dx.footer_info()
    bad = []
    if dang:
        bad.append(f'document.xml 引用了不存在的关系 {dang[:4]}')
    if fi['missing']:
        bad.append(f"页脚部件缺失 {fi['missing']}")
    rid2h = dx.rid_to_header()
    names = set(dx.zf.namelist())
    for s in dx.sects_ranged():
        for ty, rid in s['refs'].items():
            part = rid2h.get(rid)
            if not part or part not in names:
                bad.append(f'页眉引用 {rid}({ty}) 指向缺失部件')
    if bad:
        return False, '关系断裂, Word 打开时可能提示需要修复: ' + '; '.join(bad)
    return True, '图片/页眉/页脚等关系引用均可解析, 无悬挂引用'



def score(target_path, input_path=None):
    """input_path 仅为兼容旧调用而保留: 减分项已改为直接比对细则写明的绝对值。"""
    dx = Docx(target_path)
    result = {'file': target_path, 'items': [], 'add': 0, 'sub': 0, 'total': 0,
              'openable': dx.ok, 'max_add': 13}

    def rec(code, pts, ok, note):
        result['items'].append({'code': code, 'points': pts, 'passed': ok, 'note': note})

    if not dx.ok:
        result['note'] = f'文件无法打开: {dx.err}; 维度一不通过, 总分 0'
        return result

    result['note_ref'] = '减分项判据: 《细则.docx》"二、减分项"写明的绝对格式值(共 %d 条)' % len(RULES)

    before, after, abs_i, secs, split_ok, split_why = \
        effective_default_headers_by_region(dx)

    # +5 摘要及之后 与 PDF 一致
    after_lines = [b for b in after if b is not None]
    ok5 = len(after_lines) > 0 and all(
        b['val'] == PDF_STYLE and b['sz'] == PDF_SZ for b in after_lines)
    detail5 = ', '.join(f"{b['val']}/sz{b['sz']}" for b in after_lines) or '无横线'
    rec('+5 摘要后横线与PDF一致', 5, ok5, f'摘要及之后生效横线: {detail5}')

    # +3 摘要之前 全部删除
    ok3a = len(before) > 0 and all(b is None for b in before)
    if ok3a:
        note3a = '全部为无'
    else:
        got = ', '.join(f"{b['val']}/sz{b['sz']}" if b else '无' for b in before)
        note3a = (got or '无对应节') + '(其中任一处有线即不得分)'
    rec('+3 摘要前横线删除', 3, ok3a, '摘要之前生效横线: ' + note3a)

    # +3 摘要前正确分节: 分节符须紧接"摘  要"之前, 前置内容不得并入后半部分
    rec('+3 摘要前正确分节', 3, split_ok, f'共 {len(secs)} 节; {split_why}')

    # +1 分界点恰在摘要页 (摘要前节均无线, 摘要所在节起有线, 且分节点紧邻摘要)
    ok1a = ok3a and split_ok and (after[0] is not None if after else False)
    rec('+1 分界点在摘要页', 1, ok1a,
        '摘要前无线且摘要所在节起即有线' if ok1a else '分界点不在摘要页')

    # +1 单双数页一致
    #    前提: 横线已按要求处理(摘要之前无线、摘要及之后有线);
    #    页眉下方本就没有横线、或前后两部分未按要求区分时不得分。
    even_on = dx.even_odd_headers()
    has_line_after = len(after_lines) > 0
    if not (ok3a and has_line_after):
        ok1b = False
        why = ('摘要及之后各页页眉下方无横线' if not has_line_after
               else '摘要之前各页页眉下方仍有横线')
        note1b = f'横线尚未按要求处理({why}), 不具备"单双数页横线一致"的前提, 不得分'
    elif not even_on:
        ok1b = True
        note1b = '已关闭奇偶页页眉, 单双页共用同一横线'
    else:
        # even 开启: 需 even 与 default 每节底边框一致
        rid2h = dx.rid_to_header()
        consistent = True
        for s in secs:
            d = s['refs'].get('default'); e = s['refs'].get('even')
            if e:
                bd = dx.header_bottom(rid2h.get(d, '')) if d else None
                be = dx.header_bottom(rid2h.get(e, ''))
                if bd != be:
                    consistent = False
        ok1b = consistent
        note1b = '奇偶页页眉开启且已同步' if consistent else '奇偶页页眉开启但不一致'
    rec('+1 单双数页横线一致', 1, ok1b, note1b)

    result['add'] = sum(i['points'] for i in result['items'] if i['points'] > 0 and i['passed'])

    # ---------- 减分项: 逐条比对细则写明的绝对格式值(无需原始输入作参照) ----------
    #   passed=True  → 已核查, 与细则写明的绝对值一致, 未触发扣分(实得0)
    #   passed=False → 检出指令未要求的多余改动, 触发扣分
    #   passed=None  → 对象缺失等无法核查的情形(内容缺失另由通用破坏项判定)
    try:
        L = Layout(dx)
    except Exception as e:
        L = None
        result['note_ref'] = f'格式解析失败({e}), 减分项未核查'
    if L is not None:
        for code, pts, fn in RULES:
            try:
                ok, note = fn(dx, L)
            except Exception as e:
                ok, note = None, f'该项核查异常({type(e).__name__}: {e}), 未计分'
            rec(code, pts, ok, note)

    result['sub'] = sum(i['points'] for i in result['items']
                        if i['points'] < 0 and i['passed'] is False)
    # 不设下限: 加分项不得分而减分项累计时, 总分可为负数(如实反映多余改动的代价)
    result['total'] = result['add'] + result['sub']
    return result


def report(res):
    lines = []
    lines.append(f"评测文件: {res['file']}")
    if not res['openable']:
        lines.append(f"  [维度一] 文件无法打开 → 总分 0")
        lines.append(f"  {res.get('note','')}")
        return '\n'.join(lines)
    lines.append(f"  [维度一] 可正常打开 ✓")
    lines.append(f"  [维度二] 逐条判定:")
    for it in res['items']:
        if it['points'] > 0:
            # 加分项
            if it['passed']:
                verdict = f"得分 ✓  得 +{it['points']}"
                got = it['points']
            else:
                verdict = f"不得分 ✗  得 0 (满分+{it['points']})"
                got = 0
        else:
            # 减分项: 三态 —— passed=False 触发扣分; True 已核查未触发; None 未核查
            if it['passed'] is False:
                verdict = f"触发扣分 ✗  扣 {it['points']}"
                got = it['points']
            elif it['passed'] is True:
                verdict = f"已核查未触发 ✓  实得 0 (风险 {it['points']})"
                got = 0
            else:
                verdict = f"未核查 —  实得 0 (风险 {it['points']})"
                got = 0
        lines.append(f"    - {it['code']}")
        lines.append(f"        判定: {verdict}   实得 {got:+d}")
        lines.append(f"        依据: {it['note']}")
    if res.get('note_ref'):
        lines.append(f"  注: {res['note_ref']}")
    lines.append(f"  ---")
    lines.append(f"  加分合计 {res['add']} / {res['max_add']}   减分合计 {res['sub']}")
    lines.append(f"  >>> 完成度总分 = {res['total']} / {res['max_add']}")
    return '\n'.join(lines)


def auto_discover(dirpath):
    """扫描目录下的 .docx, 区分'输入文件'与'待评测文件'。
    含'输入'关键字(或文件名以 input 开头)的视为原始输入。"""
    import os
    docs = [f for f in os.listdir(dirpath)
            if f.lower().endswith('.docx') and not f.startswith('~$')]
    inputs, targets = [], []
    for f in docs:
        low = f.lower()
        if '输入' in f or low.startswith('input') or '原始' in f:
            inputs.append(f)
        else:
            targets.append(f)
    inp = os.path.join(dirpath, inputs[0]) if inputs else None
    targets = [os.path.join(dirpath, t) for t in targets]
    return targets, inp, docs


# ---- 统一对外接口(见 "脚本接口差异与统一建议.md" §2.1 / §2.2) -----------------
# 说明: 内部评分逻辑(Docx / score / report / 各校验项判定语义)保持不变;
# 此处新增模块级 evaluate(file_path) 作为唯一对外入口, 把 score() 结果规范成
# §2.2 约定的 dict; 主结果只走 return, 不 print、不改 sys.stdout、不 sys.exit。

SCRIPT_ID = "002"


def _locate_office_file(dir_path):
    """在任务目录中定位待评测的 DOCX(排除'输入/原始'与临时文件)。"""
    import os, glob
    candidates = sorted(
        p for p in glob.glob(os.path.join(dir_path, "*.docx"))
        if not os.path.basename(p).startswith("~$")
    )
    if not candidates:
        raise FileNotFoundError(f"目录中未找到 .docx 文件：{dir_path}")
    targets = [
        p for p in candidates
        if not ('输入' in os.path.basename(p)
                or os.path.basename(p).lower().startswith('input')
                or '原始' in os.path.basename(p))
    ]
    return (targets or candidates)[0]


def _find_input_beside(target_path):
    """在待评测文件同目录内寻找原始输入文档(用于减分核查); 找不到返回 None。"""
    import os, glob
    dir_path = os.path.dirname(os.path.abspath(target_path))
    for p in sorted(glob.glob(os.path.join(dir_path, "*.docx"))):
        base = os.path.basename(p)
        if base.startswith("~$"):
            continue
        if ('输入' in base or base.lower().startswith('input') or '原始' in base):
            return p
    return None


def _resolve_office_file(path):
    """兼容批量运行器目录入口和本地调试的文件入口。"""
    import os
    if os.path.isdir(path):
        return _locate_office_file(path)
    if os.path.isfile(path):
        ext = os.path.splitext(path)[1].lower()
        if ext == '.docx' and not os.path.basename(path).startswith("~$"):
            return path
        raise FileNotFoundError(f"不是可评测的 .docx 文件：{path}")
    raise FileNotFoundError(f"路径不存在：{path}")


def evaluate(file_path):
    """脚本对外唯一入口, 接收任务目录或单个 .docx, 返回规范化结果 dict。

    - 不改动内部评估逻辑与参数, 仍复用 Docx / score()。
    - 主结果仅通过 return 传出; 任何未捕获异常一律折成 status="error",
      以区分"评估 0 分"与"脚本自身崩溃"。
    - 定位不到 .docx 等前置失败时, dim1_pass 与 dim1_reason 一律留空(空字符串,
      表示维度一根本没进入判定)、dim2_items=[]、total_score=max_score=0,
      error 仅给出可读原因(不带异常类名前缀)。
    - 维度一未通过时 dim1_pass=False, 且 dim1_reason 带出错误摘要。
    - dim2_items 命中项与未命中项都列出, 便于对齐评分项矩阵。
    """
    import os
    dim2_items = []
    result = {
        "id": SCRIPT_ID,
        "file_name": "",
        "status": "ok",
        "error": None,
        "dim1_pass": False,
        "dim1_reason": "",
        "dim2_items": dim2_items,
        "total_score": 0,
        "max_score": 0,
    }

    try:
        target = _resolve_office_file(file_path)
        result["file_name"] = os.path.basename(target)
        inp = _find_input_beside(target)
        res = score(target, inp)
    except FileNotFoundError as e:
        # 找不到被评估文件：error 只回传纯文本原因(不带异常类名)；维度一根本没进入
        # 判定，dim1_pass 与 dim1_reason 一律留空(空字符串,而非 False)，dim2 不出项，
        # 总分与满分一并归零。
        result["status"] = "error"
        result["error"] = str(e)
        result["file_name"] = ""
        result["dim1_pass"] = ""
        result["dim1_reason"] = ""
        result["dim2_items"] = []
        result["total_score"] = 0
        result["max_score"] = 0
        return result
    except Exception as e:
        result["status"] = "error"
        result["error"] = f"脚本运行异常：{type(e).__name__}: {e}"
        return result

    # 维度一: 文件可否正常打开(不可打开则维度一未通过, 总分 0)
    if not res.get("openable", False):
        # 维度一未通过: dim1_pass 保持 False, 且 dim1_reason 必须带上错误摘要,
        # 便于直接定位原因(note 里已含底层解析报错, 缺失时给出兜底说明)
        note = (res.get("note") or "").strip()
        result["dim1_reason"] = note or "文件无法打开或无法解析(未取得具体报错), 维度一不通过, 总分 0"
        result["total_score"] = 0
        result["max_score"] = int(res.get("max_add", 13))
        return result
    result["dim1_pass"] = True

    # 维度二: 逐项映射
    #   加分项(points>0): hit=passed; 命中得 +points
    #   减分项(points<0): 内部 passed=False 表示"检出与细则写明的绝对值不符的改动"→触发扣分;
    #                     hit=(passed is False); 命中扣 points(本身为负)
    max_score = 0
    for it in res.get("items", []):
        pts = it["points"]
        passed = it["passed"]
        if pts > 0:
            hit = bool(passed)
            max_delta = pts
            delta = pts if hit else 0
            max_score += pts
        else:
            hit = (passed is False)
            max_delta = pts
            delta = pts if hit else 0
        dim2_items.append({
            "rule": it["code"],
            "max_delta": max_delta,
            "delta": delta,
            "hit": hit,
        })

    result["max_score"] = int(res.get("max_add", max_score))
    result["total_score"] = int(res.get("total", 0))
    return result


def main():
    """本地调试入口: evaluate(sys.argv[1]) 后打印规范化 JSON 到 stdout。

    无参数时兜底扫描脚本同目录下的 .docx 用于自测。
    """
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass

    import os
    if len(sys.argv) >= 2:
        target = sys.argv[1]
    else:
        target = os.path.dirname(os.path.abspath(__file__))

    print(json.dumps(evaluate(target), ensure_ascii=False, indent=2))


if __name__ == '__main__':
    main()
