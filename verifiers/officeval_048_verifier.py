#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# pyright: reportUnknownMemberType=false, reportUnknownVariableType=false, reportUnknownArgumentType=false, reportUnknownParameterType=false, reportMissingParameterType=false, reportPrivateUsage=false, reportUnknownLambdaType=false, reportGeneralTypeIssues=false

import os
import re
import sys
import json
import wave
import zipfile
import traceback
from xml.etree import ElementTree as ET

# 说明：接口统一约定，本模块不修改全局 sys.stdout，也不 print 主结果。
# 主结果统一由 evaluate(dir_path) -> dict 返回，入参为"脚本所在目录的路径"，
# 由脚本自己在该目录内定位并打开被评估的 .pptx。本文件仅在 __main__ 下作为
# 本地调试入口，把该 dict 序列化为 JSON 打印。

SCRIPT_ID = "048"
# 维度二所有加分项 max_delta 之和，用于维度一未通过 / 脚本异常时兜底
MAX_SCORE_TOTAL = 59

# 维度二加分项按"结果追加顺序"的 max_delta 列表，用于把末端 plus_results 里的
# "+N：..." 字符串拆解为结构化 dim2_items。顺序须与 _evaluate_impl 中
# plus_results.append(...) 的顺序严格一致，共 19 项，累加 = MAX_SCORE_TOTAL。
_MAX_DELTAS_ORDER = [
    3,  # 01 进入第1页时8条案例全部隐藏，画面静止等待
    5,  # 02 第一条单击触发，其余7条自动接着播，全程点1次
    3,  # 03 出现顺序与时间轴一致
    3,  # 04 单条案例动画持续 0.1-0.5 秒
    3,  # 05 相邻两条开始时间相差 0.15-1 秒
    1,  # 06 8条全部播完总时长 1.5-6 秒
    1,  # 07 带位移或速度感的进入效果，非缓慢淡入非随机
    3,  # 08 前3条都设置了进入动画
    3,  # 09 中间3条都设置了进入动画
    3,  # 10 最后2条都设置了进入动画
    3,  # 11 每条案例4个部件同组同时开始
    1,  # 12 年份色块和引线一起出现或完全不参与动画
    5,  # 13 真正的进入类动画效果，不是只标记为显示的空动作
    5,  # 14 播放前隐藏、由隐藏变显示、都在动画列表里
    5,  # 15 无退出类动画，无重新隐藏或还原到起点
    5,  # 16 有音效、能解码、非静音、音量非零
    3,  # 17 音效挂在最后一条案例上
    1,  # 18 音效与最后一条案例同时开始，无额外延迟
    3,  # 19 音效嵌入文件内部，不是外部链接
]
assert sum(_MAX_DELTAS_ORDER) == MAX_SCORE_TOTAL

# 维度二减分项按"结果追加顺序"的 max_delta 列表（均为负值），共 14 项。
# 与 minus_results.append(...) 的顺序严格一致。减分项不计入 max_score。
_MIN_DELTAS_ORDER = [
    -5,  # 01 第1页8条案例文字内容
    -3,  # 02 第1页案例卡片布局与尺寸
    -3,  # 03 第1页案例配色与字号
    -3,  # 04 第1页非案例元素齐备
    -1,  # 05 第1页页脚齐备
    -3,  # 06 非案例元素被加上动画
    -5,  # 07 共16页且第2-16页标题
    -5,  # 08 第2-16页正文抽检
    -3,  # 09 第2-16页页脚页码与形状边界
    -3,  # 10 页面尺寸16比9
    -3,  # 11 形状里文字实际套用的字体为微软雅黑
    -3,  # 12 第2-16页无动画且全16页无切换
    -3,  # 13 无自动换片时间且无循环放映
    -3,  # 14 页面上无喇叭图标播放控件音频占位形状
]

NS = {
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "ct": "http://schemas.openxmlformats.org/package/2006/content-types",
    "pr": "http://schemas.openxmlformats.org/package/2006/relationships",
}

EMU_PER_CM = 360000.0
R_EMBED = "{%s}embed" % NS["r"]
R_LINK = "{%s}link" % NS["r"]
R_ID = "{%s}id" % NS["r"]

# 带位移或速度感的进入效果白名单（与办公软件"动画"库对齐）：
# 2=飞入 13=劈裂 那类擦除式 22=擦除 23=缩放 36=中心旋转 53=弹跳。
# 10=淡入（无位移，属于缓慢淡入）与 0（随机效果）不在白名单内。
_MOTION_PRESET_IDS = {2, 3, 4, 5, 6, 8, 12, 13, 14, 16, 18, 22, 23, 25, 26, 27,
                      28, 30, 31, 34, 35, 36, 37, 38, 39, 40, 41, 42, 47, 48,
                      49, 50, 51, 52, 53}

# 第1页8条案例的年份标签文字（用于出现顺序校验，破折号为全角长横）
_YEAR_TEXTS = ["1948", "1956", "1967", "1980—1988", "1990—1991", "2003",
               "2011以来", "2023以来"]

# 第2至16页的页面标题文字（逐字对照，标点为全角）
_TITLES_2_16 = [
    "解锁战争背后的",
    "本课的三个解码方向",
    "半球与纬度：位置先决定自然环境",
    "“一湾两洋三洲四通道五海”的枢纽格局",
    "解码一",
    "黑色黄金：资源集中带来的安全焦虑",
    "读图任务：石油分布与出口路线",
    "资源带来的不同声音",
    "石油资源的两面性",
    "生命之源",
    "水资源为何会成为冲突导火索",
    "气候图里的水资源线索",
    "应对水资源匮乏的路径",
    "三把钥匙，看懂区域冲突",
    "谢谢观看",
]


def _txt(el):
    """取出一个形状/段落节点下的全部文字，按 a:t 顺序拼接。"""
    if el is None:
        return ""
    return "".join(t.text or "" for t in el.iter("{%s}t" % NS["a"]))


def _num(v, default=None):
    try:
        return int(v)
    except (TypeError, ValueError):
        return default


def _delay_ms(cTn):
    """读取一个 p:cTn 的开始延迟（毫秒）。返回 (延迟, 是否需要单击)。"""
    if cTn is None:
        return 0, False
    for cond in cTn.findall("./p:stCondLst/p:cond", NS):
        d = cond.get("delay")
        if d == "indefinite":
            return 0, True
        n = _num(d)
        if n is not None:
            return n, False
    return 0, False


class Shape(object):
    """幻灯片上一个形状/图片的几何与文字摘要。"""

    __slots__ = ("sid", "name", "x", "y", "cx", "cy", "text", "sizes",
                 "fills", "kind", "prst", "el")

    def __init__(self, sid, name, x, y, cx, cy, text, sizes, fills, kind,
                 prst, el):
        self.sid = sid
        self.name = name
        self.x = x
        self.y = y
        self.cx = cx
        self.cy = cy
        self.text = text
        self.sizes = sizes
        self.fills = fills
        self.kind = kind
        self.prst = prst
        self.el = el

    # 允许 cy 为负（翻转的装饰线），取包围盒时统一用 min/max
    @property
    def x0(self):
        return min(self.x, self.x + self.cx)

    @property
    def x1(self):
        return max(self.x, self.x + self.cx)

    @property
    def y0(self):
        return min(self.y, self.y + self.cy)

    @property
    def y1(self):
        return max(self.y, self.y + self.cy)

    @property
    def xc(self):
        return (self.x0 + self.x1) / 2.0

    @property
    def yc(self):
        return (self.y0 + self.y1) / 2.0


def _collect_shapes(spTree):
    """遍历 spTree（含组合内部），产出 Shape 列表。"""
    out = []
    if spTree is None:
        return out
    tags = {
        "{%s}sp" % NS["p"]: "sp",
        "{%s}pic" % NS["p"]: "pic",
        "{%s}graphicFrame" % NS["p"]: "frame",
        "{%s}cxnSp" % NS["p"]: "cxn",
        "{%s}grpSp" % NS["p"]: "grp",
    }
    for el in spTree.iter():
        kind = tags.get(el.tag)
        if kind is None or kind == "grp":
            continue
        nv = el.find(".//p:cNvPr", NS)
        sid = _num(nv.get("id")) if nv is not None else None
        name = nv.get("name", "") if nv is not None else ""
        off = el.find("./p:spPr/a:xfrm/a:off", NS)
        ext = el.find("./p:spPr/a:xfrm/a:ext", NS)
        x = _num(off.get("x"), 0) if off is not None else 0
        y = _num(off.get("y"), 0) if off is not None else 0
        cx = _num(ext.get("cx"), 0) if ext is not None else 0
        cy = _num(ext.get("cy"), 0) if ext is not None else 0
        prstEl = el.find("./p:spPr/a:prstGeom", NS)
        prst = prstEl.get("prst") if prstEl is not None else None
        sizes = sorted({_num(r.get("sz")) for r in el.iter("{%s}rPr" % NS["a"])
                        if _num(r.get("sz")) is not None})
        fills = sorted({c.get("val", "").upper()
                        for c in el.iter("{%s}srgbClr" % NS["a"])
                        if c.get("val")})
        out.append(Shape(sid, name, x, y, cx, cy, _txt(el), sizes, fills,
                         kind, prst, el))
    return out


_BEH_TAGS = {("{%s}" % NS["p"]) + t for t in
             ("anim", "animClr", "animEffect", "animMotion", "animRot",
              "animScale", "set", "cmd")}
_CONTAINER_TAGS = {("{%s}" % NS["p"]) + t for t in ("par", "seq", "excl")}


class Timing(object):
    """一页幻灯片的动画时间轴解析结果（时间单位统一为毫秒）。"""

    def __init__(self):
        self.groups = []        # 顶层单击组：{"need_click","start","end","idx"}
        self.behaviors = []     # 全部动作：dict(spid,start,dur,preset*,grp,gidx,tag,to)
        self.audios = []        # 音效节点：dict(embed,link,vol,start,gidx)
        self.bld_spids = set()  # 出现在 p:bldLst / p:bldP 里的形状编号
        self.node_start = {}    # p:cTn 的 id -> 绝对开始时间

    @property
    def anim_spids(self):
        return {b["spid"] for b in self.behaviors if b["spid"] is not None}


def _walk_cTn(cTn, base, gidx, preset, tm):
    """递归累加延迟，落到动作节点上时记录绝对开始时间。"""
    nid = cTn.get("id")
    if nid is not None and nid not in tm.node_start:
        tm.node_start[nid] = base
    if cTn.get("presetID") is not None or cTn.get("presetClass") is not None:
        preset = (_num(cTn.get("presetID")), cTn.get("presetClass"),
                  _num(cTn.get("presetSubtype")), cTn.get("grpId"))
    # 动画节点可能挂在 p:childTnLst 下，音效等媒体节点常挂在 p:subTnLst 下，两者都要走
    children = []
    for tag in ("childTnLst", "subTnLst"):
        c = cTn.find("./p:%s" % tag, NS)
        if c is not None:
            children.extend(list(c))
    if not children:
        return
    for child in children:
        if child.tag in _CONTAINER_TAGS:
            sub = child.find("./p:cTn", NS)
            if sub is None:
                continue
            d, _ = _delay_ms(sub)
            _walk_cTn(sub, base + d, gidx, preset, tm)
        elif child.tag in _BEH_TAGS:
            bc = child.find("./p:cBhvr/p:cTn", NS)
            bd, _ = _delay_ms(bc)
            dur = _num(bc.get("dur"), 0) if bc is not None else 0
            spt = child.find("./p:cBhvr/p:tgtEl/p:spTgt", NS)
            attr = child.find("./p:cBhvr/p:attrNameLst/p:attrName", NS)
            to = child.find("./p:to/p:strVal", NS)
            tm.behaviors.append({
                "spid": _num(spt.get("spid")) if spt is not None else None,
                "start": base + bd,
                "dur": dur or 0,
                "presetID": preset[0],
                "presetClass": preset[1],
                "presetSubtype": preset[2],
                "grp": preset[3],
                "gidx": gidx,
                "tag": child.tag.rsplit("}", 1)[-1],
                "attr": (attr.text or "") if attr is not None else "",
                "to": to.get("val") if to is not None else None,
            })
        elif child.tag == "{%s}audio" % NS["p"]:
            mc = child.find("./p:cMediaNode", NS)
            bc = mc.find("./p:cTn", NS) if mc is not None else None
            bd, _ = _delay_ms(bc)
            snd = mc.find("./p:tgtEl/p:sndTgt", NS) if mc is not None else None
            tm.audios.append({
                "embed": snd.get(R_EMBED) if snd is not None else None,
                "link": snd.get(R_LINK) if snd is not None else None,
                "vol": (mc.get("vol") if mc is not None else None),
                "start": base + bd,
                "bd": bd,
                "gidx": gidx,
                "tn": _tn_ref(bc),
                # 办公软件保存的动画音效，媒体节点上必然能读到播放长度：
                # 写了 dur（含 indefinite），或者给了结束条件
                # （PowerPoint / WPS 写的都是 evt="onStopAudio"）。
                # 两者在 schema 里都是可选的，缺席不算违反规范，
                # 只是说明这段 XML 不是办公软件生成的。
                "dur_raw": (bc.get("dur") if bc is not None else None),
                "has_end": (bc is not None
                            and bc.find("./p:endCondLst/p:cond", NS) is not None),
            })


def _tn_ref(cTn):
    """音效等节点可能写成"跟随某个动画节点开始"，取出被引用的节点 id。"""
    if cTn is None:
        return None
    for cond in cTn.findall("./p:stCondLst/p:cond", NS):
        tn = cond.find("./p:tn", NS)
        if tn is not None and tn.get("val"):
            return tn.get("val")
    return None


def _parse_timing(slide_root):
    """把一页的 p:timing 解析成 Timing。没有动画时返回空的 Timing。"""
    tm = Timing()
    if slide_root is None:
        return tm
    timing = slide_root.find("./p:timing", NS)
    if timing is None:
        return tm
    for bp in timing.iter("{%s}bldP" % NS["p"]):
        v = _num(bp.get("spid"))
        if v is not None:
            tm.bld_spids.add(v)
    for bg in timing.iter("{%s}bldGraphic" % NS["p"]):
        v = _num(bg.get("spid"))
        if v is not None:
            tm.bld_spids.add(v)
    seq = None
    for s in timing.iter("{%s}seq" % NS["p"]):
        seq = s
        break
    if seq is None:
        # 没有主序列时，仍尝试从整棵时间树上收集动作，保证不漏判
        root = timing.find("./p:tnLst/p:par/p:cTn", NS)
        if root is not None:
            _walk_cTn(root, 0, 0, (None, None, None, None), tm)
        return tm
    seq_cTn = seq.find("./p:cTn", NS)
    if seq_cTn is None:
        return tm
    lst = seq_cTn.find("./p:childTnLst", NS)
    if lst is None:
        return tm
    prev_start = 0
    prev_end = 0
    for idx, par in enumerate(lst.findall("./p:par", NS)):
        cTn = par.find("./p:cTn", NS)
        if cTn is None:
            continue
        d, click = _delay_ms(cTn)
        node_type = cTn.get("nodeType") or ""
        if click:
            # 需要观众单击才开始，单击时刻即上一组结束时刻
            start = prev_end
        elif node_type == "withEffect":
            start = prev_start + d
        else:
            # afterEffect / 其余：上一组播完再等 d 毫秒
            start = prev_end + d
        before = len(tm.behaviors)
        _walk_cTn(cTn, start, idx, (None, None, None, None), tm)
        mine = tm.behaviors[before:]
        end = max([b["start"] + b["dur"] for b in mine] + [start])
        tm.groups.append({"idx": idx, "need_click": click, "start": start,
                          "end": end, "nodeType": node_type})
        prev_start, prev_end = start, end
    return tm


class Deck(object):
    """打开 .pptx 包，按放映顺序整理出各页的 XML、关系与媒体清单。"""

    def __init__(self, path):
        self.zf = zipfile.ZipFile(path, "r")
        self.names = set(self.zf.namelist())
        self.pres = self._xml("ppt/presentation.xml")
        self.pres_rels = self._rels("ppt/presentation.xml")
        self.slide_parts = []
        sldIdLst = self.pres.find("./p:sldIdLst", NS) if self.pres is not None else None
        if sldIdLst is not None:
            for sid in sldIdLst.findall("./p:sldId", NS):
                tgt = self.pres_rels.get(sid.get(R_ID), {}).get("target")
                if tgt:
                    self.slide_parts.append(self._abs("ppt", tgt))
        self.slides = [self._xml(p) for p in self.slide_parts]
        self.slide_rels = [self._rels(p) for p in self.slide_parts]
        sz = self.pres.find("./p:sldSz", NS) if self.pres is not None else None
        self.page_cx = _num(sz.get("cx"), 0) if sz is not None else 0
        self.page_cy = _num(sz.get("cy"), 0) if sz is not None else 0
        self._shapes = {}

    def close(self):
        try:
            self.zf.close()
        except Exception:
            pass

    def _xml(self, name):
        if name not in self.names:
            return None
        try:
            return ET.fromstring(self.zf.read(name))
        except ET.ParseError:
            return None

    @staticmethod
    def _abs(base_dir, target):
        t = target.replace("\\", "/")
        if t.startswith("/"):
            return t.lstrip("/")
        parts = (base_dir + "/" + t).split("/")
        out = []
        for p in parts:
            if p == "..":
                if out:
                    out.pop()
            elif p and p != ".":
                out.append(p)
        return "/".join(out)

    def _rels(self, part):
        d, f = os.path.split(part)
        rp = (d + "/_rels/" + f + ".rels") if d else ("_rels/" + f + ".rels")
        root = self._xml(rp)
        out = {}
        if root is None:
            return out
        for rel in root:
            out[rel.get("Id")] = {
                "target": rel.get("Target", ""),
                "type": rel.get("Type", ""),
                "mode": rel.get("TargetMode", ""),
            }
        return out

    def shapes(self, i):
        """第 i 页（0 基）的形状列表，带缓存。"""
        if i not in self._shapes:
            root = self.slides[i] if 0 <= i < len(self.slides) else None
            tree = root.find("./p:cSld/p:spTree", NS) if root is not None else None
            self._shapes[i] = _collect_shapes(tree)
        return self._shapes[i]

    def raw(self, i):
        try:
            return self.zf.read(self.slide_parts[i]).decode("utf-8", "ignore")
        except Exception:
            return ""

    @property
    def n_slides(self):
        return len(self.slides)


def _probe_audio(zf, name):
    """解码音频，返回 (是否可解码, 时长秒, 峰值比例, 说明)。"""
    try:
        data = zf.read(name)
    except Exception as exc:
        return False, 0.0, 0.0, "读取失败 %s" % exc
    if len(data) < 512:
        return False, 0.0, 0.0, "音频数据过小"
    if not name.lower().endswith(".wav"):
        # 标准库只能可靠解码 wav；其它格式退化为"存在且有实际数据"
        return True, 0.0, 1.0, "非 wav 格式，按有效数据处理"
    import io
    try:
        wf = wave.open(io.BytesIO(data), "rb")
    except Exception as exc:
        return False, 0.0, 0.0, "无法解码 %s" % exc
    try:
        sw = wf.getsampwidth()
        fr = wf.getframerate() or 1
        nf = wf.getnframes()
        dur = nf / float(fr)
        raw = wf.readframes(min(nf, 4000000))
    finally:
        wf.close()
    if sw <= 0 or not raw:
        return False, dur, 0.0, "无采样数据"
    full = float((1 << (8 * sw - 1)) - 1)
    peak = 0
    step = sw
    if sw == 1:
        for i in range(0, len(raw), step):
            v = abs(raw[i] - 128)
            if v > peak:
                peak = v
        full = 127.0
    else:
        for i in range(0, len(raw) - sw + 1, step):
            v = int.from_bytes(raw[i:i + sw], "little", signed=True)
            v = -v if v < 0 else v
            if v > peak:
                peak = v
    ratio = peak / full if full else 0.0
    return True, dur, ratio, "时长 %.3f 秒，峰值比例 %.4f" % (dur, ratio)


_YEAR_RE = re.compile(r"^\s*\d{4}\s*(—\s*\d{4})?\s*(以来)?\s*$")


def _page1_cases(shapes, page_cx, page_cy):
    """把第1页切成 8 条案例：年份色块 + 引线 + 卡片区部件。按横向位置排序。"""
    badges = [s for s in shapes if _YEAR_RE.match(s.text or "")]
    badges.sort(key=lambda s: s.xc)
    if len(badges) < 2:
        return badges, []
    xs = [s.xc for s in badges]
    pitch = min([xs[i + 1] - xs[i] for i in range(len(xs) - 1)] or [page_cx / 8.0])
    half = max(pitch * 0.55, page_cx / 24.0)
    cases = []
    for b in badges:
        cluster = []
        for s in shapes:
            if s is b:
                continue
            if abs(s.xc - b.xc) > half:
                continue
            if abs(s.cx) > page_cx * 0.5 or abs(s.cy) > page_cy * 0.5:
                continue
            # 只收时间轴上下相邻的卡片带，避开页面标题区与页脚区
            if abs(s.yc - b.yc) > page_cy * 0.22:
                continue
            cluster.append(s)
        leaders = [s for s in cluster if s.cx == 0 or s.cy == 0]
        parts = [s for s in cluster if s not in leaders]
        cases.append({"badge": b, "leaders": leaders, "parts": parts,
                      "year": re.sub(r"\s+", "", b.text or "")})
    return badges, cases


# 第2至16页正文抽检锚点：页序号（1 基） -> 必须出现的文字片段
_BODY_ANCHORS = {
    2: ["地理密码", "空间、资源与民生逻辑"],
    3: ["咽喉之地", "为什么同一片区域会反复成为国际关注焦点"],
    4: ["区域整体处在东半球与北半球", "跨越中低纬度地带"],
    5: ["霍尔木兹海峡、曼德海峡、苏伊士运河、土耳其海峡", "阿拉伯海、红海、地中海、黑海、里海"],
    6: ["为地理位置而战", "控制通行就会影响供应链"],
    7: ["海湾危机的地理视角", "资源越集中，运输通道越窄"],
    8: ["区域能源格局图", "港口—海峡—市场"],
    9: ["同一资源既能创造财富，也可能放大风险", "远洋采购经理林澜"],
    10: ["有利影响", "资源优势只有转化为多元产业与稳定治理"],
    11: ["为水资源而战", "更是城市、农业与边界安全的底线"],
    12: ["以约旦河流域争端为例", "水源分布不均"],
    13: ["降水稀少，蒸发旺盛", "取水工程容易触动邻国利益"],
    14: ["海水淡化", "需要技术、制度与区域互信共同支撑"],
    15: ["三把钥匙", "理解复杂问题中最稳定的空间线索"],
    16: ["地理位置、石油资源与水资源，共同塑造了区域发展的机会与压力"],
}

# 页脚文字片段（每页都应齐备）
_FOOTER_PARTS = ["第九章", "东半球其他地区与国家", "第一节", "第1课时"]

# 第1页允许出现的字号（单位为百分之一磅）
_PAGE1_SIZES = {660, 760, 880, 900, 1000, 1100, 1150, 1450, 3000}

# 允许的字体名（微软雅黑），以 + 开头的主题字体占位不参与判定
_ALLOWED_FONTS = {"microsoft yahei", "微软雅黑", "microsoft yahei ui"}


def _fill_color(el):
    c = el.find("./p:spPr/a:solidFill/a:srgbClr", NS)
    return c.get("val", "").upper() if c is not None else None


def _norm(s):
    return re.sub(r"\s+", "", s or "")


def _check_dim1(file_path, deck_box):
    """维度一：可用或可修改性。返回 (是否通过, 不通过的原因)，通过时原因为空串。"""
    if not file_path.lower().endswith(".pptx"):
        return False, "不是 .pptx 格式的文件"
    try:
        deck = Deck(file_path)
    except zipfile.BadZipFile:
        return False, "文件无法打开，不是有效的演示文稿包（可能已加密或损坏）"
    deck_box.append(deck)
    if "[Content_Types].xml" not in deck.names:
        return False, "内部结构不完整，缺少内容类型部件"
    if deck.pres is None:
        return False, "内部结构不完整，缺少或无法解析演示文稿主部件"
    pres_raw = deck.zf.read("ppt/presentation.xml").decode("utf-8", "ignore")
    if not re.search(r"<p:sldMasterId\s", pres_raw):
        return False, "内部结构不完整，缺少幻灯片母版引用"
    if deck.n_slides < 1:
        return False, "内部结构不完整，没有任何幻灯片"
    for i, part in enumerate(deck.slide_parts):
        if part not in deck.names or deck.slides[i] is None:
            return False, "第 %d 页部件缺失或无法解析" % (i + 1)
    if re.search(r"<p:modifyVerifier", pres_raw):
        return False, "文件被设置为只读或有修改密码，无法编辑"
    for i in range(deck.n_slides):
        shapes = deck.shapes(i)
        if not any(_norm(s.text) for s in shapes):
            pics = [s for s in shapes if s.kind == "pic"]
            if pics:
                return False, "第 %d 页只有图片没有文字，属于导出的图片式页面" % (i + 1)
            return False, "第 %d 页是空白页，没有任何文字" % (i + 1)
    for i in range(deck.n_slides):
        for rid, rel in deck.slide_rels[i].items():
            t = rel["type"].rsplit("/", 1)[-1]
            if t in ("image", "audio", "video", "media") and rel["mode"] == "External":
                return False, "第 %d 页的图片或音频是外部链接，没有保存在文件内部" % (i + 1)
            if t in ("image", "audio", "video", "media"):
                tgt = Deck._abs(os.path.dirname(deck.slide_parts[i]), rel["target"])
                if tgt not in deck.names:
                    return False, "第 %d 页引用的图片或音频在文件内部找不到" % (i + 1)
    # 通过时不需要说明理由，返回空串
    return True, ""


def evaluate(dir_path):
    """统一入口：在 dir_path 内定位唯一的 .pptx 并给出评测结果字典。"""
    result = {
        "id": SCRIPT_ID,
        "file_name": "",
        "status": "ok",
        "error": None,
        "dim1_pass": False,
        "dim1_reason": "",
        "dim2_items": [],
        "total_score": 0,
        "max_score": MAX_SCORE_TOTAL,
    }
    try:
        if not dir_path or not os.path.isdir(dir_path):
            raise FileNotFoundError("目录不存在：%s" % dir_path)
        cands = [f for f in os.listdir(dir_path)
                 if f.lower().endswith(".pptx") and not f.startswith("~$")]
        if not cands:
            raise FileNotFoundError("目录内未找到 .pptx 文件：%s" % dir_path)
        file_name = sorted(cands)[0]
        result["file_name"] = file_name
        _evaluate_impl(os.path.join(dir_path, file_name), result)
    except Exception as exc:  # noqa: BLE001
        # status 为 error 时未能完成判定，dim1_pass 留空，dim1_reason 也为空
        result["status"] = "error"
        result["error"] = "%s: %s" % (type(exc).__name__, exc)
        result["dim1_pass"] = ""
        result["dim1_reason"] = ""
        result["dim2_items"] = []
        result["total_score"] = 0
        result["max_score"] = MAX_SCORE_TOTAL
    return result


def _pack(result, plus_results, minus_results):
    """把 "+N：..." / "-N：..." 结果串拆成结构化 dim2_items 并累计得分。"""
    if len(plus_results) != len(_MAX_DELTAS_ORDER):
        raise RuntimeError("加分项数量与 _MAX_DELTAS_ORDER 不一致：%d vs %d"
                           % (len(plus_results), len(_MAX_DELTAS_ORDER)))
    if len(minus_results) != len(_MIN_DELTAS_ORDER):
        raise RuntimeError("减分项数量与 _MIN_DELTAS_ORDER 不一致：%d vs %d"
                           % (len(minus_results), len(_MIN_DELTAS_ORDER)))
    items = []
    total = 0
    for text, max_delta in list(zip(plus_results, _MAX_DELTAS_ORDER)) + \
            list(zip(minus_results, _MIN_DELTAS_ORDER)):
        head, _, body = text.partition("：")
        delta = int(head)
        items.append({
            "rule": body.split("｜", 1)[0],
            "max_delta": max_delta,
            "delta": delta,
            "hit": delta != 0,
        })
        total += delta
    result["dim2_items"] = items
    result["total_score"] = max(0, total)
    result["max_score"] = MAX_SCORE_TOTAL
    return result


def _evaluate_impl(file_path, result):
    """维度一先做门槛判定，通过后才逐条计算维度二的加分项与减分项。"""
    deck_box = []
    ok1, reason1 = _check_dim1(file_path, deck_box)
    result["dim1_pass"] = ok1
    result["dim1_reason"] = reason1
    deck = deck_box[0] if deck_box else None
    try:
        if not ok1:
            # 维度一任一条不通过：维度二不再计分，完成度按 0 处理
            result["dim2_items"] = []
            result["total_score"] = 0
            result["max_score"] = MAX_SCORE_TOTAL
            return result
        plus_results, minus_results = _score_dim2(deck)
        _pack(result, plus_results, minus_results)
    finally:
        if deck is not None:
            deck.close()
    return result


def _score_dim2(deck):
    """返回 (plus_results, minus_results) 两个字符串列表。"""
    page_cx = deck.page_cx or 1
    page_cy = deck.page_cy or 1
    sh1 = deck.shapes(0)
    tm = _parse_timing(deck.slides[0])
    badges, cases = _page1_cases(sh1, page_cx, page_cy)
    n_case = len(cases)

    beh_by_spid = {}
    for b in tm.behaviors:
        beh_by_spid.setdefault(b["spid"], []).append(b)

    def _behs(shapes):
        out = []
        for s in shapes:
            out.extend(beh_by_spid.get(s.sid, []))
        return out

    cb = [_behs(c["parts"]) for c in cases]
    starts = [min([b["start"] for b in x]) if x else None for x in cb]
    ends = [max([b["start"] + b["dur"] for b in x]) if x else None for x in cb]
    durs = [max([b["dur"] for b in x]) if x else None for x in cb]
    entr = [[b for b in x if b["presetClass"] == "entr"] for x in cb]
    covered = [i for i in range(n_case) if cb[i]]

    # ---- 门槛一：是否是真正的进入类动画效果（决定前 15 条加分是否作废）----
    # 与办公软件对齐：办公软件在"动画"里写下的进入效果，会在时间轴节点上留下
    # 进入类标记与效果编号；只把形状标记为"显示"的空动作没有这两样东西。
    anim_ok = (n_case == 8 and len(covered) == 8
               and all(entr[i] for i in range(8))
               and all(any(b["tag"] != "set" for b in cb[i]) for i in range(8)))

    # ---- 门槛二：音效是否有效（决定后 4 条加分是否作废）----
    a_hit = False
    a_start = None
    a_detail = "第 1 页没有找到任何音效节点"
    a_embed = False
    a_int = False
    if tm.audios:
        a = tm.audios[0]
        a_start = a["start"]
        if a["tn"] and a["tn"] in tm.node_start:
            a_start = tm.node_start[a["tn"]] + a["bd"]
        rid = a["embed"] or a["link"]
        rel = deck.slide_rels[0].get(rid) if rid else None
        a_embed = bool(a["embed"]) and not a["link"]
        if rel is None:
            a_detail = "音效节点没有指向任何音频文件"
        elif rel["mode"] == "External":
            a_detail = "音效是外部链接，文件没有保存在演示文稿内部"
        else:
            tgt = Deck._abs(os.path.dirname(deck.slide_parts[0]), rel["target"])
            if tgt not in deck.names:
                a_detail = "音效指向的音频文件在演示文稿内部找不到"
            else:
                a_int = True
                dec, dur, peak, msg = _probe_audio(deck.zf, tgt)
                vol = a["vol"]
                vol_ok = (vol is None) or (_num(vol, 100000) > 0)
                # dur 与 endCondLst 在 CT_TimeNode 里都是可选的，两者都缺席并不
                # 违反 schema；但办公软件保存动画音效时必写其中之一，所以按第 16
                # 条的文本要求，媒体节点上必须能读到播放长度。
                play_ok = bool(a.get("dur_raw")) or bool(a.get("has_end"))
                if not dec:
                    a_detail = "音频无法解码：%s" % msg
                elif peak < 0.02:
                    a_detail = "音频接近静音，%s" % msg
                elif not vol_ok:
                    a_detail = "音效音量被设为零"
                elif not play_ok:
                    a_detail = "音效节点既没有写播放时长也没有结束条件"
                else:
                    a_hit = True
                    a_detail = "音效嵌入在 %s，%s" % (tgt.rsplit("/", 1)[-1], msg)
    audio_ok = a_hit

    plus = []

    def _a(lst, delta, hit, rule, detail):
        lst.append("%+d：%s｜%s" % (delta if hit else 0, rule, detail))

    last = n_case - 1
    n_click = sum(1 for g in tm.groups if g["need_click"])

    # 01 单击前画面静止：第一个顶层组必须等观众单击，且 8 条案例都在动画列表里
    c1 = bool(tm.groups) and tm.groups[0]["need_click"] and len(covered) == 8
    _a(plus, 3, anim_ok and c1, "进入第1页时8条案例全部隐藏，画面静止等待",
       "顶层动画组 %d 个，首组需单击=%s，被动画覆盖的案例 %d 条"
       % (len(tm.groups), bool(tm.groups) and tm.groups[0]["need_click"], len(covered)))

    # 02 全程只点 1 次：需要单击的顶层组恰好 1 个，其余组都是数值延迟自动接续
    c2 = (n_click == 1 and all(not g["need_click"] for g in tm.groups[1:])
          and len(covered) == 8)
    _a(plus, 5, anim_ok and c2, "第一条单击触发，其余7条自动接着播，全程只点1次",
       "需要单击的动画组 %d 个，自动接续组 %d 个" % (n_click, max(0, len(tm.groups) - n_click)))

    # 03 出现顺序与时间轴一致：按横向位置排出的 8 条案例，开始时间必须不递减
    c3 = (len(covered) == 8
          and all(starts[i] <= starts[i + 1] for i in range(n_case - 1)))
    _a(plus, 3, anim_ok and c3, "8条案例的出现顺序与时间轴从左到右一致",
       "案例顺序 %s，开始时间 %s" % ("、".join(c["year"] for c in cases),
                                 [starts[i] for i in range(n_case)]))

    # 04 单条动画持续 0.1 到 0.5 秒：取该条案例各部件动作时长的最大值
    c4 = len(covered) == 8 and all(100 <= (durs[i] or 0) <= 500 for i in range(n_case))
    _a(plus, 3, anim_ok and c4, "单条案例的动画持续时间在0.1到0.5秒之间",
       "各条时长毫秒 %s" % [durs[i] for i in range(n_case)])

    gaps = [starts[i + 1] - starts[i] for i in range(n_case - 1)] \
        if len(covered) == 8 else []
    # 05 相邻两条开始时间相差 0.15 到 1 秒
    c5 = bool(gaps) and all(150 <= g <= 1000 for g in gaps)
    _a(plus, 3, anim_ok and c5, "相邻两条案例的开始时间相差0.15到1秒",
       "相邻间隔毫秒 %s" % gaps)

    span = (max([e for e in ends if e is not None]) -
            min([s for s in starts if s is not None])) if covered else 0
    # 06 8 条全部播完的总时长 1.5 到 6 秒
    c6 = len(covered) == 8 and 1500 <= span <= 6000
    _a(plus, 1, anim_ok and c6, "8条案例全部播完的总时长在1.5到6秒之间",
       "总时长 %d 毫秒" % span)

    pids = sorted({b["presetID"] for i in range(n_case) for b in entr[i]
                   if b["presetID"] is not None})
    # 07 带位移或速度感的进入效果：效果编号必须落在有位移感的白名单内
    c7 = bool(pids) and all(p in _MOTION_PRESET_IDS for p in pids)
    _a(plus, 1, anim_ok and c7, "用的是带位移或速度感的进入效果，不是缓慢淡入也不是随机效果",
       "进入效果编号 %s" % pids)

    # 08 到 10 覆盖检查：前 3 条 / 中间 3 条 / 最后 2 条都必须设置了进入动画
    g1 = [i for i in (0, 1, 2) if i < n_case and entr[i]]
    _a(plus, 3, anim_ok and len(g1) == 3, "前3条案例都设置了进入动画",
       "前3条中有进入动画的 %d 条" % len(g1))
    g2 = [i for i in (3, 4, 5) if i < n_case and entr[i]]
    _a(plus, 3, anim_ok and len(g2) == 3, "中间3条案例都设置了进入动画",
       "中间3条中有进入动画的 %d 条" % len(g2))
    g3 = [i for i in (6, 7) if i < n_case and entr[i]]
    _a(plus, 3, anim_ok and len(g3) == 2, "最后2条案例都设置了进入动画",
       "最后2条中有进入动画的 %d 条" % len(g3))

    def _same_group(i):
        sids = [s.sid for s in cases[i]["parts"]]
        if len(sids) < 3:
            return False
        hit_sids = {b["spid"] for b in cb[i]}
        if not set(sids) <= hit_sids:
            return False
        st = [b["start"] for b in cb[i]]
        if max(st) - min(st) > 10:
            return False
        # 办公软件只给同一"组"里的第 2 个及之后的效果写分组编号，空值不参与比对
        grps = {b["grp"] for b in cb[i] if b["grp"] is not None}
        return len(grps) <= 1

    # 11 每条案例的 4 个部件同组同时开始：同一分组编号且开始时间一致
    c11 = n_case == 8 and all(_same_group(i) for i in range(n_case))
    _a(plus, 3, anim_ok and c11, "每条案例的4个部件在同一个动画组里同时开始",
       "各条部件数 %s" % [len(c["parts"]) for c in cases])

    def _badge_ok(i):
        extra = [cases[i]["badge"]] + cases[i]["leaders"]
        moved = [s for s in extra if beh_by_spid.get(s.sid)]
        if not moved:
            return True
        if len(moved) != len(extra) or starts[i] is None:
            return False
        bl = _behs(extra)
        return max(abs(b["start"] - starts[i]) for b in bl) <= 10

    # 12 年份色块和引线：要么跟本条案例一起出现，要么完全不参与动画
    c12 = n_case == 8 and all(_badge_ok(i) for i in range(n_case))
    _a(plus, 1, anim_ok and c12, "年份色块和引线跟案例一起出现，或者完全不参与动画",
       "参与动画的年份色块与引线 %d 个"
       % sum(1 for c in cases for s in [c["badge"]] + c["leaders"]
             if beh_by_spid.get(s.sid)))

    # 13 门槛：必须是真正的进入类动画，不能只是把形状标记为显示的空动作
    _a(plus, 5, anim_ok, "案例用的是真正的进入类动画效果，不是只把形状标记为显示的空动作",
       "带进入标记的案例 %d 条，进入类动作 %d 个"
       % (sum(1 for i in range(n_case) if entr[i]),
          sum(len(entr[i]) for i in range(n_case))))

    def _vis_ok():
        if n_case != 8:
            return False
        for i in range(n_case):
            for s in cases[i]["parts"]:
                bl = beh_by_spid.get(s.sid, [])
                if not bl:
                    return False
                if not any(b["tag"] == "set" and "visibility" in b["attr"]
                           and (b["to"] or "").lower() == "visible" for b in bl):
                    return False
        return True

    # 14 播放前隐藏、由隐藏变显示、并且都在动画列表里
    c14 = _vis_ok()
    _a(plus, 5, anim_ok and c14, "案例部件播放前隐藏，由隐藏变为显示，并且都在动画列表里",
       "可见性动作 %d 个，动画构建列表中的形状 %d 个"
       % (sum(1 for b in tm.behaviors if b["tag"] == "set"
              and "visibility" in b["attr"]), len(tm.bld_spids)))

    raw1 = deck.raw(0)
    # 15 没有退出类动画，也没有重新隐藏或还原到起点的动作
    c15 = (not any(b["presetClass"] == "exit" for b in tm.behaviors)
           and not any(b["tag"] == "set" and "visibility" in b["attr"]
                       and (b["to"] or "").lower() == "hidden"
                       for b in tm.behaviors)
           and 'transition="out"' not in raw1)
    _a(plus, 5, anim_ok and c15, "没有退出类动画，也没有重新隐藏或还原到起点的动作",
       "退出类动作 %d 个，重新隐藏动作 %d 个"
       % (sum(1 for b in tm.behaviors if b["presetClass"] == "exit"),
          sum(1 for b in tm.behaviors if b["tag"] == "set"
              and "visibility" in b["attr"] and (b["to"] or "").lower() == "hidden")))

    # 16 门槛：音效存在、能解码、不是静音、音量不为零，且节点写明了播放长度
    r16 = "第1页有一个能正常解码、不是静音、音量不为零的音效，且音效节点写明了播放时长或结束条件"
    _a(plus, 5, audio_ok, r16, a_detail)

    # 17 挂载位置：全页只有一个音效，且挂在最后一条案例上
    c17 = (len(tm.audios) == 1 and last >= 0 and starts[last] is not None
           and a_start is not None and a_start >= starts[last] - 60)
    _a(plus, 3, audio_ok and c17, "音效挂在最后一条案例上，全页没有别的音效",
       "音效节点 %d 个，音效开始 %s 毫秒，最后一条案例开始 %s 毫秒"
       % (len(tm.audios), a_start, starts[last] if last >= 0 else None))

    # 18 时间对齐：音效与最后一条案例同时开始，没有额外延迟
    c18 = (last >= 0 and starts[last] is not None and a_start is not None
           and abs(a_start - starts[last]) <= 60)
    _a(plus, 1, audio_ok and c18, "音效与最后一条案例同时开始，没有额外延迟",
       "时间差 %s 毫秒"
       % (abs(a_start - starts[last]) if (a_start is not None and last >= 0
                                          and starts[last] is not None) else None))

    # 19 音效嵌入在文件内部，不是外部链接
    _a(plus, 3, audio_ok and a_embed and a_int, "音效嵌入在演示文稿文件内部，不是外部链接",
       "嵌入引用=%s，包内存在=%s" % (a_embed, a_int))

    minus = []
    events = ["阿以战争", "苏伊士运河危机", "六日战争", "两伊战争", "海湾战争",
              "伊拉克战争", "叙利亚冲突", "加沙冲突"]
    tags = ["位置", "通道", "位置", "石油", "石油", "资源", "水源", "民生"]

    def ptext(i):
        return _norm("".join(s.text for s in deck.shapes(i)))

    years = [c["year"] for c in cases]
    case_texts = [_norm("".join(s.text for s in c["parts"])) for c in cases]

    # 减 01 案例文字内容：条数、年份、事件名称、分类标签逐条比对
    bad = []
    if n_case != 8:
        bad.append("案例数量为 %d 条" % n_case)
    else:
        if years != _YEAR_TEXTS:
            bad.append("年份文字为 %s" % "、".join(years))
        for i in range(8):
            if events[i] not in case_texts[i]:
                bad.append("第 %d 条缺少事件名称 %s" % (i + 1, events[i]))
            if tags[i] not in case_texts[i]:
                bad.append("第 %d 条缺少分类标签 %s" % (i + 1, tags[i]))
            if sum(1 for s in cases[i]["parts"] if _norm(s.text)) < 2:
                bad.append("第 %d 条文字部件不足" % (i + 1))
    _a(minus, -5, bool(bad), "第1页8条案例的文字内容齐备且未被改动",
       "；".join(bad[:6]) if bad else "8条案例的年份、事件名称与分类标签均齐备")

    # 减 02 卡片布局与尺寸：部件数、卡片尺寸一致、上下交替、同侧上边缘对齐
    bad = []
    if n_case != 8:
        bad.append("案例数量为 %d 条" % n_case)
    else:
        cards = []
        for i in range(8):
            if len(cases[i]["parts"]) != 4:
                bad.append("第 %d 条部件数为 %d" % (i + 1, len(cases[i]["parts"])))
            cards.append(max(cases[i]["parts"],
                             key=lambda s: abs(s.cx) * abs(s.cy)))
        w = [abs(s.cx) for s in cards]
        h = [abs(s.cy) for s in cards]
        if max(w) - min(w) > max(w) * 0.01 or max(h) - min(h) > max(h) * 0.01:
            bad.append("8个卡片尺寸不一致")
        side = [1 if s.yc > page_cy / 2.0 else -1 for s in cards]
        if any(side[i] == side[i + 1] for i in range(7)):
            bad.append("卡片没有沿时间轴上下交替")
        for sgn in (1, -1):
            tops = [s.y0 for s, v in zip(cards, side) if v == sgn]
            if tops and max(tops) - min(tops) > page_cy * 0.01:
                bad.append("同一侧卡片上边缘没有对齐")
    _a(minus, -3, bool(bad), "第1页案例卡片的布局与尺寸保持原样",
       "；".join(bad[:6]) if bad else "8个卡片等尺寸、上下交替、同侧上边缘对齐")

    # 减 03 配色与字号：年份色块两色交替，全页字号在既有档位内
    bad = []
    bcolors = [_fill_color(c["badge"].el) for c in cases]
    uniq = [x for x in dict.fromkeys([c for c in bcolors if c])]
    if n_case != 8 or len(uniq) != 2:
        bad.append("年份色块配色为 %s" % ("、".join(uniq) if uniq else "无"))
    elif any(bcolors[i] == bcolors[i + 1] for i in range(7)):
        bad.append("年份色块颜色没有交替")
    sizes1 = sorted({z for s in sh1 for z in s.sizes})
    extra_sz = [z for z in sizes1 if z not in _PAGE1_SIZES]
    if extra_sz:
        bad.append("出现了额外字号 %s" % ["%.2f磅" % (z / 100.0) for z in extra_sz])
    _a(minus, -3, bool(bad), "第1页案例的配色与字号保持原样",
       "；".join(bad[:6]) if bad else "年份色块两色交替，字号档位未被改动")

    # 减 04 非案例元素齐备：主轴线、30磅主标题、副标题、轴标题、背景图片
    bad = []
    axis = [s for s in sh1 if abs(s.cx) > page_cx * 0.6 and s.cy == 0]
    if not axis:
        bad.append("缺少贯穿页面的时间轴主线")
    if not any(3000 in s.sizes for s in sh1):
        bad.append("缺少30磅的页面主标题")
    t1 = ptext(0)
    for frag in ("从冲突看区域地理", "一条时间线，串联位置、石油与水资源", "历史事件轴"):
        if _norm(frag) not in t1:
            bad.append("缺少文字 %s" % frag)
    if not any(s.kind == "pic" for s in sh1):
        bad.append("缺少页面配图")
    _a(minus, -3, bool(bad), "第1页的非案例元素齐备",
       "；".join(bad[:6]) if bad else "主轴线、主标题、副标题、轴标题与配图均在")

    def footer_shapes(i):
        return [s for s in deck.shapes(i) if s.y0 > page_cy * 0.85 and _norm(s.text)]

    # 减 05 第1页页脚齐备：细分隔线、左下三段文字、右下页码、页脚字号
    bad = []
    if not [s for s in sh1 if abs(s.cx) > page_cx * 0.5 and s.y0 > page_cy * 0.85]:
        bad.append("缺少页脚细分隔线")
    fsh = footer_shapes(0)
    ftxt = _norm("".join(s.text for s in fsh))
    for frag in _FOOTER_PARTS:
        if frag not in ftxt:
            bad.append("页脚缺少文字 %s" % frag)
    if not any(_norm(s.text) == "01" and s.xc > page_cx * 0.5 for s in fsh):
        bad.append("右下缺少页码 01")
    fsz = {z for s in fsh for z in s.sizes}
    if not {880, 900} <= fsz:
        bad.append("页脚字号为 %s" % ["%.2f磅" % (z / 100.0) for z in sorted(fsz)])
    _a(minus, -1, bool(bad), "第1页页脚齐备",
       "；".join(bad[:6]) if bad else "分隔线、左下三段文字、右下页码与字号均齐备")

    # 减 06 非案例元素被加上动画：第1页动画目标必须全在案例元素范围内
    case_ids = set()
    for c in cases:
        for s in c["parts"] + c["leaders"] + [c["badge"]]:
            if s.sid is not None:
                case_ids.add(s.sid)
    stray = sorted(x for x in tm.anim_spids if x not in case_ids)
    _a(minus, -3, bool(stray), "第1页只有案例元素带动画，其它元素没有被加动画",
       "被加动画的非案例形状 %d 个" % len(stray) if stray else "动画只作用在案例元素上")

    # 减 07 共16页且第2-16页标题：页数与逐页标题文字比对
    bad = []
    if deck.n_slides != 16:
        bad.append("共 %d 页" % deck.n_slides)
    for idx, title in enumerate(_TITLES_2_16, start=2):
        if idx - 1 >= deck.n_slides:
            break
        if _norm(title) not in ptext(idx - 1):
            bad.append("第 %d 页缺少标题 %s" % (idx, title))
    _a(minus, -5, bool(bad), "共16页且第2到16页的标题保持原样",
       "；".join(bad[:6]) if bad else "页数为16页，第2到16页标题逐页一致")

    # 减 08 第2-16页正文抽检：每页抽取关键文字片段核对
    bad = []
    for idx, anchors in sorted(_BODY_ANCHORS.items()):
        if idx - 1 >= deck.n_slides:
            bad.append("第 %d 页缺失" % idx)
            continue
        t = ptext(idx - 1)
        for anchor in anchors:
            if _norm(anchor) not in t:
                bad.append("第 %d 页缺少正文 %s" % (idx, anchor))
    _a(minus, -5, bool(bad), "第2到16页的正文内容保持原样",
       "；".join(bad[:6]) if bad else "逐页抽检的正文片段全部命中")

    # 减 09 第2-16页页脚页码与形状边界：页码、页脚文字、形状不出界
    bad = []
    for i in range(1, deck.n_slides):
        fs = footer_shapes(i)
        ft = _norm("".join(s.text for s in fs))
        if not any(_norm(s.text) == "%02d" % (i + 1) for s in fs):
            bad.append("第 %d 页缺少页码 %02d" % (i + 1, i + 1))
        if not all(frag in ft for frag in _FOOTER_PARTS):
            bad.append("第 %d 页页脚文字不全" % (i + 1))
    tolx, toly = page_cx * 0.05, page_cy * 0.05
    for i in range(deck.n_slides):
        for s in deck.shapes(i):
            if s.x0 < -tolx or s.y0 < -toly or \
                    s.x1 > page_cx + tolx or s.y1 > page_cy + toly:
                bad.append("第 %d 页有形状超出页面边界" % (i + 1))
                break
    _a(minus, -3, bool(bad), "第2到16页的页脚页码齐备，所有形状都在页面范围内",
       "；".join(bad[:6]) if bad else "页码与页脚齐备，没有形状超出页面")

    # 减 10 页面尺寸 16 比 9，允许千分之五以内偏差
    ratio = page_cx / float(page_cy) if page_cy else 0.0
    off = abs(ratio - 16.0 / 9.0) / (16.0 / 9.0)
    _a(minus, -3, off > 0.005, "页面尺寸为16比9",
       "页面宽高比 %.5f，偏差 %.4f%%" % (ratio, off * 100.0))

    # 减 11 字体：形状里文字实际套用的字体必须是微软雅黑
    fonts = set()
    for i in range(deck.n_slides):
        root = deck.slides[i]
        if root is None:
            continue
        for tag in ("latin", "ea", "cs"):
            for el in root.iter("{%s}%s" % (NS["a"], tag)):
                tf = (el.get("typeface") or "").strip()
                if tf and not tf.startswith("+"):
                    fonts.add(tf)
    wrong = sorted(f for f in fonts if f.lower() not in _ALLOWED_FONTS)
    _a(minus, -3, bool(wrong), "形状里文字实际套用的字体是微软雅黑",
       "出现了其它字体 %s" % "、".join(wrong[:5]) if wrong
       else "全部形状使用微软雅黑")

    # 减 12 第2-16页无动画，且全部16页都没有切换效果
    bad = []
    for i in range(1, deck.n_slides):
        t = _parse_timing(deck.slides[i])
        if t.behaviors or t.audios:
            bad.append("第 %d 页被加了动画" % (i + 1))
    for i in range(deck.n_slides):
        if deck.slides[i] is not None and \
                deck.slides[i].find("./p:transition", NS) is not None:
            bad.append("第 %d 页被加了切换效果" % (i + 1))
    _a(minus, -3, bool(bad), "第2到16页没有动画，全部16页都没有切换效果",
       "；".join(bad[:6]) if bad else "只有第1页有动画，全篇没有切换效果")

    # 减 13 没有自动换片时间，也没有设置循环放映
    bad = []
    for i in range(deck.n_slides):
        if re.search(r"advTm=", deck.raw(i)):
            bad.append("第 %d 页设置了自动换片时间" % (i + 1))
    pres_raw = deck.zf.read("ppt/presentation.xml").decode("utf-8", "ignore")
    if re.search(r'<p:showPr[^>]*loop="(1|true)"', pres_raw):
        bad.append("设置了循环放映")
    _a(minus, -3, bool(bad), "没有自动换片时间，也没有设置循环放映",
       "；".join(bad[:6]) if bad else "换片仍由单击控制，没有循环放映")

    # 减 14 页面上没有喇叭图标、播放控件或音频占位形状
    bad = []
    for i in range(deck.n_slides):
        raw = deck.raw(i)
        if re.search(r"<a:audioFile|<a:videoFile|<p:media\b", raw):
            bad.append("第 %d 页放了音视频占位形状" % (i + 1))
        for s in deck.shapes(i):
            nm = (s.name or "").lower()
            if any(k in nm for k in ("speaker", "audio", "media", "sound")) or \
                    any(k in (s.name or "") for k in ("扬声器", "喇叭", "声音", "媒体")):
                bad.append("第 %d 页有名为 %s 的音频形状" % (i + 1, s.name))
    _a(minus, -3, bool(bad), "页面上没有喇叭图标、播放控件或音频占位形状",
       "；".join(bad[:6]) if bad else "音效只挂在动画上，页面上没有音频形状")

    return plus, minus


if __name__ == "__main__":
    _dir = sys.argv[1] if len(sys.argv) > 1 else os.path.dirname(os.path.abspath(__file__))
    try:
        _out = evaluate(_dir)
    except Exception:  # noqa: BLE001
        _out = {
            "id": SCRIPT_ID,
            "file_name": "",
            "status": "error",
            "error": traceback.format_exc(limit=3),
            "dim1_pass": "",
            "dim1_reason": "",
            "dim2_items": [],
            "total_score": 0,
            "max_score": MAX_SCORE_TOTAL,
        }
    print(json.dumps(_out, ensure_ascii=False, indent=2))









