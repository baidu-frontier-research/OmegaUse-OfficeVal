# -*- coding: utf-8 -*-
"""
自动评估脚本：对一个 PDF 按"打分细则"打分。

评估逻辑
--------
维度1（可用与可修改性）：硬性门槛。任何一条不满足 -> 总分 0，且不再检查维度2。
维度2（完成度）：在维度1通过后逐条检测加分点/扣分点并累加。
  - 加分细则：必须满足该细则内"每一个点"才计该正分。
  - 扣分细则：满足该细则内"任意一点"即计该负分。

因为目标 PDF 为纯图像页（无可提取文本层），所有文字/表格/题目相关的检测
均在渲染后的位图（pypdfium2 经 pdf_backend 适配层渲染）上用图像处理（OpenCV）
以"灵活变通"的方式实现，逼近评分意图。

用法:
    python officeval_027_verifier.py ["脚本所在目录路径"]
不传参数时默认使用脚本自身所在目录；目录内待评估的 PDF 由 evaluate() 自动定位。
"""
import sys, os, math, glob, json

_DEPENDENCY_ERRORS = []
try:
    import numpy as np
except ImportError:
    np = None
    _DEPENDENCY_ERRORS.append("numpy")
try:
    try:
        import pdf_backend
    except ImportError:
        from verifiers import pdf_backend
    import pypdfium2  # noqa: F401 —— 渲染依赖，缺失时提前报告
except ImportError:
    pdf_backend = None
    _DEPENDENCY_ERRORS.append("pdfplumber/pypdfium2")
try:
    import cv2
except ImportError:
    cv2 = None
    _DEPENDENCY_ERRORS.append("opencv-python")

# ----------------------------- 常量 -----------------------------
SCRIPT_ID = "027"
A4_W_PT, A4_H_PT = 595.276, 841.890   # A4 纵向，单位 point (1pt = 1/72 inch)
A4_RATIO = A4_H_PT / A4_W_PT          # ≈ 1.414
RENDER_DPI = 150
PT_PER_CM = 72.0 / 2.54               # 1cm ≈ 28.35pt


# ===================================================================
#  文档/页面句柄（适配层包装）与渲染
# ===================================================================
class _Page:
    """页面句柄：尺寸/旋转/渲染/嵌入图像/注释计数（替代 fitz.Page）。"""

    def __init__(self, doc: "pdf_backend.PdfDocument", index: int):
        self._doc = doc
        self.index = index
        w, h = doc.page_size(index)
        self.rect = pdf_backend.PdfRect(0.0, 0.0, w, h)
        self.rotation = doc.page_rotation(index)

    def render(self, dpi):
        return self._doc.render_page(self.index, dpi=dpi)

    def images(self):
        return self._doc.extract_images(self.index)

    def annots_count(self):
        return self._doc.count_annotations(self.index)


class _Doc:
    """文档句柄：页数/页访问/关闭（替代 fitz.Document）。"""

    def __init__(self, path: str):
        self._doc = pdf_backend.open_pdf(path)
        self.page_count = self._doc.page_count
        self._pages = [_Page(self._doc, i) for i in range(self.page_count)]

    def __getitem__(self, index: int) -> _Page:
        return self._pages[index]

    def close(self):
        self._doc.close()


def render_page(page, dpi=RENDER_DPI):
    """渲染单页为 RGB ndarray (H,W,3)。考虑页面旋转(rotation)后的视觉外观。"""
    return page.render(dpi)


def ink_mask(gray, thresh=200):
    """深色像素(墨迹/内容)的布尔掩码。"""
    return gray < thresh


def content_bbox(gray, thresh=200, min_run=3):
    """
    返回内容外接框 (top, bottom, left, right)，单位像素。
    用每行/每列的墨迹像素数过滤掉极少量噪点(min_run)。
    无内容时返回 None。
    """
    m = ink_mask(gray, thresh)
    row_cnt = m.sum(axis=1)
    col_cnt = m.sum(axis=0)
    rows = np.where(row_cnt >= min_run)[0]
    cols = np.where(col_cnt >= min_run)[0]
    if len(rows) == 0 or len(cols) == 0:
        return None
    return int(rows[0]), int(rows[-1]), int(cols[0]), int(cols[-1])


def page_analysis(page, dpi=RENDER_DPI):
    """对单页做一次性渲染+测量，返回一个 dict 供各条细则复用。"""
    img = render_page(page, dpi)
    h, w = img.shape[:2]
    gray = cv2.cvtColor(img, cv2.COLOR_RGB2GRAY)
    mask = ink_mask(gray)
    bbox = content_bbox(gray)
    rgb = img.astype(np.int16)

    # 红色像素(红框/批注/裁剪线/水印常见红色)
    red = (rgb[:, :, 0] > 150) & (rgb[:, :, 1] < 110) & (rgb[:, :, 2] < 110)

    info = {
        "img": img, "gray": gray, "mask": mask, "h": h, "w": w,
        "bbox": bbox,
        "ink_frac": float(mask.mean()),
        "red_count": int(red.sum()),
        "rect_w": page.rect.width, "rect_h": page.rect.height,
        "rotation": page.rotation,
        "px_per_cm": dpi / 2.54,
    }
    return info


# ===================================================================
#  几何/倾斜/行结构 辅助函数
# ===================================================================
def detect_text_rows(gray):
    """
    基于横向墨迹投影，找出文字行的 y 区间列表 [(y0,y1),...] 与对应间距。
    用于评估题目间垂直间距、行高。
    """
    m = ink_mask(gray)
    proj = m.sum(axis=1)
    if proj.max() == 0:
        return [], []
    thr = max(3, proj.max() * 0.03)
    on = proj > thr
    rows = []
    start = None
    for y, v in enumerate(on):
        if v and start is None:
            start = y
        elif not v and start is not None:
            rows.append((start, y - 1))
            start = None
    if start is not None:
        rows.append((start, len(on) - 1))
    # 过滤极薄的噪声行
    rows = [(a, b) for a, b in rows if b - a >= 2]
    heights = [b - a + 1 for a, b in rows]
    return rows, heights


def text_baseline_angles(gray):
    """
    返回 (angles, unresolved)：
      angles     —— 每一"可靠测得"的正文文字行的基线斜率角度列表(度)。
                    正值向下倾斜，负值向上扬，0 表示水平。
      unresolved —— 检测到疑似正文行但**无法可靠测出角度**的行数。
                    调用方应把 unresolved>0 视作"不确定 → 无法证明合格"，
                    按细则"倾斜≤±0.5°"该 +5 项定性为失败——见 plus_5_skew.

    评估对象：细则点1/点2 针对"题干正文基线"，仅评估单行正文带；
    表格 / 统计图等被投影合并成的"厚块"不属于此项（由点3的横线规则单独判定），
    此类厚块从计数中静默排除，**不**计入 unresolved.

    做法：先取文字行区间，仅保留高度接近单行正文的带；对每行取每列墨迹的
    最低点(基线/字底)拟合直线得到该行基线斜率，用 RMS 判定拟合可信度.

    不确定 → unresolved 的两种情形（原实现是直接跳过、导致真倾斜行漏进 bad 列表）：
      · 横向跨度不足页宽 15%（原代码用 25% 直接跳过，与"整行"定义脱节，
        把窄栏正文/居中标题一律丢弃，可能漏检真正上扬/下斜）；
      · 拟合残差 RMS>1.5px（原代码同样直接跳过；改为记为"不可靠 → unresolved"，
        使调用方能反映到 bad 列表）.
    """
    rows, heights = detect_text_rows(gray)
    if not heights:
        return [], 0
    line_h = float(np.median(heights))
    m = ink_mask(gray)
    angles = []
    unresolved = 0
    w = gray.shape[1]
    min_span_px = max(60, int(w * 0.15))  # 从原 25% 放宽到 15%，覆盖窄栏/居中行
    for (y0, y1) in rows:
        band_h = y1 - y0 + 1
        # 只评估单行正文带：高度不超过约 1.8 倍中位行高，排除表格/统计图厚块.
        # 厚块不属于"文字行"，故不计入 unresolved（其水平度由点3单独判定）.
        if band_h > max(line_h * 1.8, line_h + 6):
            continue
        band = m[y0:y1 + 1, :]
        ys, xs = np.where(band)
        # 完全无墨迹：非文字行，不算不确定
        if len(xs) < 20:
            continue
        cols = np.unique(xs)
        if len(cols) < 10:
            continue
        span = int(cols[-1] - cols[0])
        # 以每列墨迹的最低点(字底)作为基线采样点
        cx, cy = [], []
        for c in cols:
            yy = ys[xs == c]
            cx.append(c)
            cy.append(yy.max())
        cx = np.array(cx, dtype=float)
        cy = np.array(cy, dtype=float)
        slope, inter = np.polyfit(cx, cy, 1)
        resid = cy - (slope * cx + inter)
        rms = float(np.sqrt((resid ** 2).mean()))
        ang = math.degrees(math.atan(slope))
        # 只有当拟合足够"像一条直线" 且 横向跨度足够 → 记为可信角度；
        # 否则记为 unresolved（不再静默跳过），由 plus_5_skew 把这类页也归入 bad.
        if rms <= 1.5 and span >= min_span_px:
            angles.append(ang)
        else:
            unresolved += 1
    return angles, unresolved


def line_skew_angles_per_page(gray):
    """
    返回该页"水平线"（表格横线/统计图横轴）的角度列表(度)，用于 ±0.5° 判定.

    实现要点（相较原版）：
      · 细则针对"表格横线"和"统计图横轴"，二者长度不一定达页宽 40%——
        窄表、并排统计图的横轴常在 15%~30% 页宽区间.
        原实现要求长度 ≥ 页宽 40%（klen=0.40w、minLineLength=0.40w），
        会漏掉此类短横线，遗漏真倾斜.
      · 现放宽为 ≥ 页宽 10%（klen=0.10w、minLineLength=0.10w，Hough 阈值 0.08w）,
        并保留形态学开运算 + 角度绝对值 ≤5° 过滤，以避免文字笔画/斜向装饰误判.
    """
    h, w = gray.shape
    bw = cv2.adaptiveThreshold(gray, 255, cv2.ADAPTIVE_THRESH_MEAN_C,
                               cv2.THRESH_BINARY_INV, 31, 15)
    # 形态学只保留水平线；结构元长度从 40%→10% 页宽，以覆盖短横线/统计图横轴
    klen = max(20, int(w * 0.10))
    hkernel = cv2.getStructuringElement(cv2.MORPH_RECT, (klen, 1))
    horiz = cv2.morphologyEx(bw, cv2.MORPH_OPEN, hkernel)
    if horiz.sum() == 0:
        return []
    lines = cv2.HoughLinesP(horiz, 1, np.pi / 1800,
                            threshold=max(20, int(w * 0.08)),
                            minLineLength=max(20, int(w * 0.10)),
                            maxLineGap=10)
    angs = []
    if lines is not None:
        # OpenCV 版本差异：HoughLinesP 返回 (N,1,4) 或 (N,4)，统一 reshape.
        for x1, y1, x2, y2 in lines.reshape(-1, 4):
            a = math.degrees(math.atan2(y2 - y1, x2 - x1))
            if abs(a) <= 5:
                angs.append(a)
    return angs


# ===================================================================
#  维度1：可用与可修改性（硬性门槛）
# ===================================================================
def check_dimension_1(doc, pages):
    """
    返回 (passed: bool, reasons: list[str], details: list[str])
    details 记录每个子项的通过情况，便于打印。

    说明（本次按用户要求删减）：
      以下维度一子项已删除，不再作为硬性门槛：
        · 文件可"翻页、缩放和打印"
        · "所有题干/选项/表格/统计图/题号清晰可辨，无大面积模糊/裁切/重叠/黑屏"
        · "页面统一为 A4 纵向、阅读方向正确、不需手动旋转"
        · "若文件无法打开、题目大量缺失、清晰度严重下降或页面方向错误则维度1=0"
      现维度一仅保留最基本的"文件为 PDF 且能正常打开"这一项。
    """
    reasons = []      # 失败原因（导致维度1=0）
    details = []      # 逐项说明
    del pages         # 现维度一仅凭"能否打开"判定，不再用逐页图像信号

    # 文件为 .pdf 且能正常打开（能 open 并渲染即视为可打开）
    if doc.page_count == 0:
        reasons.append("PDF 无任何页面，无法打开")
    else:
        details.append(f"[通过] 文件为 PDF 且能正常打开，共 {doc.page_count} 页")

    passed = len(reasons) == 0
    return passed, reasons, details


# ===================================================================
#  维度2：完成度  —— 加分点
# ===================================================================
def plus_5_skew(doc, pages):
    """
    +5：全部页文字区域：完成倾斜校正，题干正文基线保持水平；倾斜角度控制在
        水平线 ±0.5 度以内，不出现肉眼可见的整行上扬或下斜。
        全部表格横线和统计图横轴与页面上边线基本平行，倾斜角度不超过 ±0.5 度。

    细则逐点对应（须每一点都满足才计 +5）：
      点1【全部页 文字区域 完成倾斜校正，题干正文基线保持水平，倾斜≤±0.5°】
          —— 对每一页的每一文字行拟合基线斜率，要求所有行 |角度| ≤ 0.5°。
      点2【不出现肉眼可见的整行上扬或下斜】
          —— "肉眼可见"以 ±0.5° 为界：任一行 |角度| > 0.5° 即视为可见上扬/下斜。
             （由点1的"全部行≤±0.5°"覆盖；单独显式判定并报告。）
      点3【全部表格横线和统计图横轴 与页面上边线基本平行，倾斜≤±0.5°】
          —— 提取每页的长横线（表格横线/统计图横轴），要求所有 |角度| ≤ 0.5°；
             "与页面上边线平行"即与水平方向夹角，故同样用 ±0.5° 衡量。
    """
    TOL = 0.5
    bad_text_pages = []      # 点1/点2：文字基线不水平 / 出现可见上扬下斜 / 无法可靠测定
    bad_line_pages = []      # 点3：横线/横轴不平行
    for i, info in enumerate(pages):
        gray = info["gray"]

        # 点1 & 点2：逐行文字基线角度
        baseline_angs, text_unresolved = text_baseline_angles(gray)
        worst_text = max((abs(a) for a in baseline_angs), default=0.0)
        if worst_text > TOL:
            bad_text_pages.append((i + 1, round(worst_text, 2)))
        elif text_unresolved > 0:
            # 检测到文字行但无法可靠测出基线角度 → 不确定，无法证明"已校正≤±0.5°"，
            # 按细则该 +5 项定性为失败（不给"无法证明的合格"）。
            bad_text_pages.append((i + 1, f"{text_unresolved}行无法测定"))

        # 点3：表格横线 / 统计图横轴角度（与页面上边线即水平方向的夹角）
        line_angs = line_skew_angles_per_page(gray)
        worst_line = max((abs(a) for a in line_angs), default=0.0)
        if worst_line > TOL:
            bad_line_pages.append((i + 1, round(worst_line, 2)))

    p1p2_ok = (len(bad_text_pages) == 0)   # 文字基线水平 + 无可见整行上扬/下斜 + 全部可测
    p3_ok = (len(bad_line_pages) == 0)     # 横线/横轴平行
    ok = p1p2_ok and p3_ok

    parts = []
    parts.append(
        "文字基线全部≤±0.5°（已校正、无可见整行上扬/下斜）"
        if p1p2_ok else
        "文字基线超±0.5°或无法测定的页: " + ", ".join(f"第{p}页({a})" for p, a in bad_text_pages[:6])
    )
    parts.append(
        "表格横线/统计图横轴全部≤±0.5°（与上边线基本平行）"
        if p3_ok else
        "横线/横轴超±0.5°的页: " + ", ".join(f"第{p}页({a}°)" for p, a in bad_line_pages[:6])
    )
    msg = "；".join(parts)
    return ok, msg


def plus_3_top_blank(doc, pages):
    """
    +3：全部页顶部区域：A4 页面顶部未出现超过页面 10% 的空白。

    细则逐点对应（须每一点都满足才计 +3）：
      点1【全部页】          —— 文档每一页都要满足，任一页不满足即不计分。
      点2【顶部区域/A4页面顶部】—— 评估对象是"页面上边线到该页首个有效内容"
                                之间的顶部空白带（不看左右或下方留白）。
      点3【未出现超过页面10%的空白】
                            —— 顶部空白带高度 ≤ A4 页面高度的 10% 即合格；
                               > 10% 即视为"出现超过页面10%的空白"。
    """
    THRESH = 10.0   # 百分比上限：未"超过"10% => ≤10% 合格
    bad = []
    for i, info in enumerate(pages):
        page_h = info["h"]                      # A4 页面高度(像素)
        bbox = info["bbox"]                     # 内容外接框 (top,bot,left,right)
        if bbox is None:
            # 整页无有效内容 => 顶部空白视为占满整页(100%)
            bad.append((i + 1, 100.0))
            continue
        top_blank = bbox[0]                     # 页顶到首个有效内容的空白高度(像素)
        frac = top_blank / page_h * 100.0       # 顶部空白占 A4 页高比例
        if frac > THRESH:
            bad.append((i + 1, round(frac, 1)))
    ok = (len(bad) == 0)
    if ok:
        msg = "全部页顶部空白均未超过 A4 页面高度的 10%"
    else:
        msg = "顶部空白超过页面10%的页: " + ", ".join(f"第{p}页({f}%)" for p, f in bad[:6])
    return ok, msg


def plus_5_spacing(doc, pages):
    """
    +5：相邻两个题目区域：保留约一行正文高度的垂直间距，题目之间既不紧贴
        也不出现大段空白.

    原实现的问题（用户反馈）:
      · 用 sep_thresh = normal_gap + 0.5×line_h 作为"是否为题目分隔"的门槛,
        只评估 g > sep_thresh 的间距 —— 真正紧贴的相邻题目 (g ≤ sep_thresh)
        被静默跳过, 且原 tight 分支 (ratio<0.5 才紧贴) 与 sep_thresh 冲突,
        实际永远无法命中 tight, 会把紧贴误判为通过.

    改造后的评估思路:
      ① 直接识别"题目边界" —— 相邻题目边界 = 相邻两个"题号所在行"之间的空白.
         image-only PDF 无文本层, 用 left-margin heuristic:
         · 每 row 的最左墨迹列 left_x;
         · 题号顶格时, 题头行 left_x 显著小于段落缩进行(即左边距行);
         · 取 left_x 分布 15% 分位 P15, 满足 left_x ≤ P15 + 0.5×line_h 的
           row 作为题头候选. 首行天然是页首题头, 单独加入.
      ② 无法识别题号(候选数<2 或占比>80%, 即无双簇) → 回退:
         把该页所有行间空白从大到小排序, 用 Otsu-style 双簇分割 (最大差比法)
         把 gap 划为"段内 vs 段间"两簇, 段间簇的每一个 gap 均视为题目分隔.
         该回退不再用 sep_thresh 硬门槛跳过.
      ③ 对每一对相邻题目边界, 直接测量 gap 与 line_h 的比值, 无跳过.

    "约一行"的明确依据 (容差):
      · 一行正文高度 = line_h (全文档所有文字行高的中位数);
      · "约" 取 ±50% → [LO, HI] × line_h = [0.5, 1.5] × line_h;
        · <0.5×line_h : 与段内换行(≈line_h)不可分, 视觉上紧贴;
        · >1.5×line_h : 空白高度超过一个整行, 属于大段空白;
        · [0.5, 1.5]  : 约一行的可视间距.
      每一对相邻题目边界的间距 ratio = gap / line_h 都必须落在 [0.5, 1.5],
      任一对越界即不给 +5.
    """
    LO, HI = 0.5, 1.5     # "约一行" = 1×line_h 的 ±50% 容差带

    # ---- 全文档基准: line_h ----
    all_heights = []
    for info in pages:
        _, hs = detect_text_rows(info["gray"])
        all_heights += hs
    if not all_heights:
        return False, "未能检测到文字行, 无法评估题目间距"
    line_h = float(np.median(all_heights))

    def _row_left_x(m, y0, y1):
        band = m[y0:y1 + 1, :]
        cols_any = band.any(axis=0)
        idx = np.where(cols_any)[0]
        return int(idx[0]) if len(idx) else None

    def _detect_boundaries(rows, gray):
        """
        返回该页所有"相邻题目边界"的 gap 列表: gap = rows[i+1].y0 - rows[i].y1,
        其中 rows[i+1] 是被识别为下一题的题头行, rows[i] 是紧邻其上一行.
        无相邻题目对时返回 [].
        """
        if len(rows) < 2:
            return []
        m = ink_mask(gray)
        left_xs = [_row_left_x(m, y0, y1) for (y0, y1) in rows]
        valid_xs = [x for x in left_xs if x is not None]
        header_idx = []
        used_fallback = False
        if len(valid_xs) >= 3:
            p15 = float(np.percentile(valid_xs, 15))
            tol = max(line_h * 0.5, 5.0)
            header_idx = [
                i for i, x in enumerate(left_xs)
                if x is not None and x <= p15 + tol
            ]
            # 首行天然是题头, 若未包含则补入
            if 0 not in header_idx:
                header_idx = [0] + header_idx
            # 双簇不成立(过多/过少) → 走 gap 回退
            if len(header_idx) < 2 or len(header_idx) > len(rows) * 0.8:
                header_idx = []
                used_fallback = True
        else:
            used_fallback = True

        if not header_idx:
            # 回退: gap 分布 2 簇分割. 用排序后相邻差比找断点, 大簇即题目分隔.
            gaps_all = [(k, rows[k + 1][0] - rows[k][1])
                        for k in range(len(rows) - 1)]
            if not gaps_all:
                return []
            sorted_gaps = sorted(gaps_all, key=lambda x: x[1])
            values = [g for _, g in sorted_gaps]
            # 找最大相对跃升位置作为断点; 若 max 与 median 之比不足 1.8, 认为
            # 全页段间/段内几乎同质 → 保守: 只把最大 gap 作为唯一分隔候选,
            # 避免把普通行间距误判为题目分隔.
            med = float(np.median(values)) if values else 0.0
            top = values[-1] if values else 0.0
            if med > 0 and top / med >= 1.8:
                # 找 sorted_gaps 中相邻比值最大的位置作为断点
                best_ratio, cut = 0.0, len(values) - 1
                for k in range(1, len(values)):
                    if values[k - 1] <= 0:
                        continue
                    r = values[k] / values[k - 1]
                    if r > best_ratio:
                        best_ratio, cut = r, k
                sep_gap_indices = {sg[0] for sg in sorted_gaps[cut:]}
            else:
                sep_gap_indices = {sorted_gaps[-1][0]} if sorted_gaps else set()
            return [(rows[k + 1][0] - rows[k][1]) for k in sorted(sep_gap_indices)]

        # 主路径: 用相邻题头对生成 gap
        header_idx = sorted(set(header_idx))
        result = []
        for j in range(1, len(header_idx)):
            hi = header_idx[j]
            if hi == 0:
                continue
            prev = rows[hi - 1]
            cur = rows[hi]
            result.append(cur[0] - prev[1])
        return result

    tight = []   # 紧贴: (页, ratio)
    big = []     # 大段空白: (页, ratio)
    total_pairs = 0
    for i, info in enumerate(pages):
        rows, _ = detect_text_rows(info["gray"])
        gaps = _detect_boundaries(rows, info["gray"])
        for g in gaps:
            total_pairs += 1
            ratio = g / line_h if line_h > 0 else 0.0
            if ratio < LO:
                tight.append((i + 1, round(ratio, 2)))
            elif ratio > HI:
                big.append((i + 1, round(ratio, 2)))

    if total_pairs == 0:
        # 全篇未能识别到相邻题目对(如整页单题) — 无法证明合格, 保守判失败
        return False, "未能识别到相邻题目边界, 无法证明题目间距达到约一行"

    p3_ok = (len(tight) == 0)
    p4_ok = (len(big) == 0)
    ok = p3_ok and p4_ok

    parts = [f"一行正文高度 line_h≈{line_h:.0f}px, 容差 [{LO},{HI}]×line_h; "
             f"共检测 {total_pairs} 对相邻题目边界"]
    if ok:
        parts.append("所有相邻题目间距均 ≈ 一行高 (不紧贴、无大段空白)")
    else:
        if not p3_ok:
            parts.append("紧贴(<0.5行)处: " +
                         ", ".join(f"第{p}页({r}×)" for p, r in tight[:6]))
        if not p4_ok:
            parts.append("大段空白(>1.5行)处: " +
                         ", ".join(f"第{p}页({r}×)" for p, r in big[:6]))
    return ok, "；".join(parts)


# ===================================================================
#  维度2：完成度  —— 扣分点
# ===================================================================
def minus_5_not_a4(doc, pages):
    """
    -5：任意三页以上不是 A4 纵向页面。

    细则逐点对应（满足任一点即判该页"不是A4纵向"，命中页数≥3 即扣分）：
      点1【A4】   —— 页面的视觉尺寸须为 A4：宽≈210mm、高≈297mm。
                     以 point 计 A4 纵向为 595.28×841.89pt，允许 ±5% 容差；
                     仅靠宽高比不足以确认是 A4（A5/A3 比例相同），故同时校验绝对尺寸。
      点2【纵向】 —— 页面视觉方向须为纵向（高 ≥ 宽）。横向(宽>高)即不满足。
      "视觉"尺寸/方向均已计入页面 rotation（旋转 90/270 时交换宽高）。
      "任意三页以上" —— 不是 A4纵向 的页数 ≥ 3 时触发扣分。
    """
    TOL = 0.05   # A4 绝对尺寸容差 ±5%
    bad = []
    for i, info in enumerate(pages):
        rot = info["rotation"]
        # 计入旋转后的视觉宽高(point)
        if rot in (90, 270):
            vis_w, vis_h = info["rect_h"], info["rect_w"]
        else:
            vis_w, vis_h = info["rect_w"], info["rect_h"]

        # 点2：纵向（高≥宽）
        is_portrait = vis_h >= vis_w
        # 点1：A4 绝对尺寸（纵向取向下，短边≈595.28pt、长边≈841.89pt）
        short, long = (vis_w, vis_h) if vis_h >= vis_w else (vis_h, vis_w)
        is_a4_size = (abs(short - A4_W_PT) <= A4_W_PT * TOL and
                      abs(long - A4_H_PT) <= A4_H_PT * TOL)

        if (not is_a4_size) or (not is_portrait):
            why = []
            if not is_a4_size:
                why.append(f"非A4尺寸({vis_w:.0f}x{vis_h:.0f}pt)")
            if not is_portrait:
                why.append("非纵向(横向)")
            bad.append((i + 1, "/".join(why)))
    trigger = len(bad) >= 3
    if bad:
        msg = f"不是A4纵向的页共 {len(bad)} 页: " + ", ".join(
            f"第{p}页({w})" for p, w in bad[:8])
    else:
        msg = "全部页面均为 A4 纵向"
    return trigger, msg


def minus_3_near_edge(doc, pages):
    """
    -3：任意三页以上内容贴近页面边缘，打印时存在被裁掉的风险。

    细则逐点对应（须每一点都满足才触发该 -3）：
      点1【内容贴近页面边缘】
          —— "内容"以该页内容外接框衡量；"贴近页面边缘"指外接框任一侧
             与对应页面物理边缘的留白过小（安全边距 < ~0.5cm，即常见打印机
             不可打印区/裁切误差范围内），任一侧贴近即视为该页"内容贴近边缘"。
      点2【打印时存在被裁掉的风险】
          —— 当内容落入页面边缘的安全边距以内时，打印/裁切时该侧内容即可能
             被裁掉，构成"被裁掉的风险"。此风险与点1同源，由同一安全边距界定。
      点3【任意三页以上】
          —— 满足"内容贴近边缘(有被裁风险)"的页数 ≥ 3 时触发扣分。
    """
    SAFE_CM = 0.5    # 安全边距：留白 < 0.5cm 视为贴近边缘、打印存在被裁风险
    bad = []
    for i, info in enumerate(pages):
        bbox = info["bbox"]
        if bbox is None:
            continue
        top, bot, left, right = bbox
        m = SAFE_CM * info["px_per_cm"]
        # 点1+点2：内容外接框任一侧留白 < 安全边距 => 贴近边缘、打印存在被裁掉风险
        near = (top < m or left < m or
                (info["h"] - 1 - bot) < m or (info["w"] - 1 - right) < m)
        if near:
            bad.append(i + 1)
    trigger = len(bad) >= 3    # 点3：达到三页以上才扣分
    msg = (f"内容贴近边缘、打印存在被裁掉风险的页共 {len(bad)} 页(<{SAFE_CM}cm): {bad[:8]}" if bad
           else f"全部页面四周留白均 ≥ {SAFE_CM}cm，无内容贴近边缘、无打印被裁掉风险")
    return trigger, msg


def minus_3_annotations(doc, pages):
    """
    -3：PDF 中出现红色标注框、批注、水印、裁剪线或临时说明文字。

    细则逐点对应（出现其中任意一类即触发该 -3）：
      点1【红色标注框】
          —— 渲染图中出现显著的红色像素块（红框/红色标注），以红色像素数量
             与占比阈值判定，避免极少量红色噪点误判。
      点2【批注】
          —— PDF 的注释对象(annotations)，即 page.annots() 返回的批注/标注。
      点3【水印】
          —— 水印通常以注释或叠加图层形式出现；落入 PDF 注释对象(点2)或
             显著红色叠加(点1)时一并被覆盖检出。
      点4【裁剪线】
          —— 裁剪线常以红色细线/标记呈现，由点1的显著红色检测覆盖；
             若以 PDF 注释形式存在，则由点2覆盖。
      点5【临时说明文字】
          —— 临时说明文字常以批注/注释或红色提示文字形式存在，
             分别由点2(PDF注释)与点1(显著红色)覆盖。
      "出现…或…"：以上任意一类（任一点命中）即触发扣分。
    """
    ann_pages = []
    for i in range(doc.page_count):
        page = doc[i]
        # 点2/点3/点5：PDF 注释对象(批注/水印/临时说明文字常以注释形式存在)
        try:
            n_annots = page.annots_count()
        except Exception:
            n_annots = 0
        if n_annots:
            ann_pages.append((i + 1, f"{n_annots}个PDF注释/批注"))
            continue
        # 点1/点4/点3/点5：显著红色像素块(红色标注框/裁剪线/红色提示文字)
        red = pages[i]["red_count"]
        # 显著红色（相对页面像素的占比阈值，避免少量红色噪点误判由阈值控制）
        red_frac = red / (pages[i]["h"] * pages[i]["w"])
        if red > 800 and red_frac > 0.0005:
            ann_pages.append((i + 1, f"红色标注/裁剪线{red}px"))
    trigger = len(ann_pages) >= 1   # "出现…或…"：任一页命中任一类即触发
    if trigger:
        msg = "发现红色标注框/批注/水印/裁剪线/临时说明文字: " + ", ".join(f"第{p}页({w})" for p, w in ann_pages[:6])
    else:
        msg = "未发现红色标注框/批注/水印/裁剪线/临时说明文字"
    return trigger, msg


def minus_5_blank_black_inverted(doc, pages):
    """
    -5：任意一页出现超过页面 30% 的空白页、黑页、页面方向倒置或需手动旋转的页面。

    细则逐点对应（出现其中任意一类即触发该 -5）：
      点1【超过页面30%的空白页】
          —— "超过页面30%"修饰空白：该页空白(近白像素)占比 > 30% 即满足。
      点2【超过页面30%的黑页】
          —— "超过页面30%"修饰黑：该页黑(近黑像素)占比 > 30% 即满足。
      点3【页面方向倒置】
          —— 方向倒置即页面旋转 180°(rotation == 180)，阅读方向上下颠倒。
      点4【需手动旋转的页面】
          —— 阅读方向不正、需手动旋转，即页面旋转 90° 或 270°(横置)。
      "任意一页 出现…或…"：任一页命中以上任意一类即触发扣分。

    说明：细则中"超过页面30%"仅修饰"空白页/黑页"；"页面方向倒置/需手动旋转"
    是独立成立的条件，不受 30% 比例约束，故不对其附加比例阈值。
    """
    bad = []
    for i, info in enumerate(pages):
        gray = info["gray"]
        white_frac = float((gray >= 245).mean())
        black_frac = float((gray < 50).mean())
        rot = info["rotation"]
        # 点1：超过页面30%的空白页（近白像素占比 > 30%）
        blank_over_30 = white_frac > 0.30
        # 点2：超过页面30%的黑页（近黑像素占比 > 30%）
        black_over_30 = black_frac > 0.30
        # 点3：页面方向倒置（旋转 180°）
        inverted = (rot == 180)
        # 点4：需手动旋转的页面（旋转 90°/270°，横置需转正）
        need_rotate = rot in (90, 270)
        if blank_over_30 or black_over_30 or inverted or need_rotate:
            why = []
            if blank_over_30: why.append(f"空白{white_frac*100:.0f}%(>30%)")
            if black_over_30: why.append(f"黑页{black_frac*100:.0f}%(>30%)")
            if inverted: why.append("方向倒置(180°)")
            if need_rotate: why.append(f"需手动旋转({rot}°)")
            bad.append((i + 1, "/".join(why)))
    trigger = len(bad) >= 1   # 任意一页命中即触发
    if trigger:
        msg = "存在 >30%空白/>30%黑页/方向倒置/需手动旋转的页: " + ", ".join(f"第{p}页({w})" for p, w in bad[:6])
    else:
        msg = "无超过页面30%的空白页/黑页，无页面方向倒置或需手动旋转的页面"
    return trigger, msg


# ===================================================================
#  主流程
# ===================================================================
def _find_pdf(dir_path):
    """在给定目录内定位待评估的 PDF：优先取唯一的 .pdf；多个时取最近修改的一个。"""
    candidates = sorted(glob.glob(os.path.join(dir_path, "*.pdf")))
    if not candidates:
        return None
    if len(candidates) == 1:
        return candidates[0]
    return max(candidates, key=os.path.getmtime)


def evaluate(dir_path: str) -> dict:
    """
    对 dir_path 目录内的 PDF 按打分细则打分。

    参数：
      dir_path：脚本所在目录的路径；脚本自己在该目录内定位并打开被评估的 PDF。

    返回：结构化字典（含维度一通过与否、维度二逐项得分、总分），不 print 主结果、
    不改 sys.stdout、不 sys.exit。
    """
    try:
        if _DEPENDENCY_ERRORS:
            return {
                "id": SCRIPT_ID, "file_name": None, "status": "error",
                "error": "缺少依赖: " + ", ".join(_DEPENDENCY_ERRORS),
                "dim1_pass": False, "dim1_reason": "",
                "dim2_items": [], "total_score": 0, "max_score": 0,
            }
        pdf_path = _find_pdf(dir_path)
        if pdf_path is None:
            return {
                "id": SCRIPT_ID, "file_name": None, "status": "error",
                "error": f"目录内未找到 .pdf 文件：{dir_path}",
                "dim1_pass": False, "dim1_reason": "",
                "dim2_items": [], "total_score": 0, "max_score": 0,
            }

        file_name = os.path.basename(pdf_path)

        try:
            doc = _Doc(pdf_path)
        except Exception as e:
            return {
                "id": SCRIPT_ID, "file_name": file_name, "status": "error",
                "error": f"文件无法打开：{e}",
                "dim1_pass": False, "dim1_reason": "",
                "dim2_items": [], "total_score": 0, "max_score": 0,
            }

        if not pdf_path.lower().endswith(".pdf"):
            doc.close()
            return {
                "id": SCRIPT_ID, "file_name": file_name, "status": "ok",
                "error": None,
                "dim1_pass": False, "dim1_reason": "交付文件不是 .pdf 格式",
                "dim2_items": [], "total_score": 0, "max_score": 0,
            }

        # 预渲染所有页（供两个维度复用）
        pages = [page_analysis(doc[i]) for i in range(doc.page_count)]

        # ---------- 维度1 ----------
        d1_pass, d1_reasons, d1_details = check_dimension_1(doc, pages)
        if not d1_pass:
            doc.close()
            return {
                "id": SCRIPT_ID, "file_name": file_name, "status": "ok",
                "error": None,
                "dim1_pass": False, "dim1_reason": "；".join(d1_reasons),
                "dim2_items": [], "total_score": 0, "max_score": 0,
            }

        # ---------- 维度2 ----------
        plus_rules = [
            (5, "全部页文字区域：完成倾斜校正，题干正文基线保持水平；倾斜角度控制在水平线±0.5度以内，不出现肉眼可见的整行上扬或下斜。全部表格横线和统计图横轴与页面上边线基本平行，倾斜角度不超过±0.5度。", plus_5_skew),
            (3, "全部页顶部区域：A4页面顶部未出现超过页面10%的空白。", plus_3_top_blank),
            (5, "相邻两个题目区域：保留约一行正文高度的垂直间距，题目之间既不紧贴也不出现大段空白。", plus_5_spacing),
        ]
        minus_rules = [
            (-5, "任意三页以上不是A4纵向页面。", minus_5_not_a4),
            (-3, "任意三页以上内容贴近页面边缘，打印时存在被裁掉的风险。", minus_3_near_edge),
            (-3, "PDF中出现红色标注框、批注、水印、裁剪线或临时说明文字。", minus_3_annotations),
            (-5, "任意一页出现超过页面30%的空白页、黑页、页面方向倒置或需手动旋转的页面。", minus_5_blank_black_inverted),
        ]

        score = 0
        max_score = sum(pts for pts, _, _ in plus_rules)
        dim2_items = []

        for pts, name, fn in plus_rules:
            ok, msg = fn(doc, pages)
            delta = pts if ok else 0
            score += delta
            dim2_items.append({
                "rule": name, "max_delta": pts, "delta": delta,
                "hit": ok, "detail": "",
            })

        for pts, name, fn in minus_rules:
            triggered, msg = fn(doc, pages)
            delta = pts if triggered else 0
            score += delta
            dim2_items.append({
                "rule": name, "max_delta": pts, "delta": delta,
                "hit": triggered, "detail": "",
            })

        doc.close()
        return {
            "id": SCRIPT_ID, "file_name": file_name, "status": "ok",
            "error": None,
            "dim1_pass": True, "dim1_reason": "",
            "dim2_items": dim2_items,
            "total_score": score, "max_score": max_score,
        }
    except Exception as e:
        return {
            "id": SCRIPT_ID, "file_name": None, "status": "error",
            "error": str(e),
            "dim1_pass": False, "dim1_reason": "",
            "dim2_items": [], "total_score": 0, "max_score": 0,
        }


if __name__ == "__main__":
    # 仅用于本地调试：传入"脚本所在目录的路径"，默认为脚本自身所在目录。
    target_dir = sys.argv[1] if len(sys.argv) > 1 else os.path.dirname(os.path.abspath(__file__))
    result = evaluate(target_dir)
    print(json.dumps(result, ensure_ascii=False, indent=2))
