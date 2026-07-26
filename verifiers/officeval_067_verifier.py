# -*- coding: utf-8 -*-
"""
对 “更新版_论文预答辩_智慧园区再生水管网_切换动画版.pptx” 进行自动评估。

评估逻辑:
  维度1 (交付文件为 .pptx 格式, 能够正常打开): 若不满足, 总分直接 0 分, 不再检查维度 2.
  维度2 (完成度):           加分项满足"全部要点"才加分; 扣分项满足"任意一点"即扣分.
"""

import json

SCRIPT_ID = "067"
import os
import re
import sys
import zipfile

# ---------- 常量 ----------
EXPECTED_PAGES = 35

NS = {
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
}

# 速度文本映射 -> 估计秒数 (PowerPoint 默认: slow~1s, med~0.75s, fast~0.5s)
SPD_TO_SEC = {"slow": 1.0, "med": 0.75, "fast": 0.5}

# ====== 可在真实办公软件中真正启动播放的切换效果白名单 ======
# 说明: 评估"动画能否在实际办公软件里看到并启动播放", 不能只看到 <p:transition>
#       就算数. 必须满足:
#         1) transition 内含有一个"标准 OOXML 切换效果标签" (在白名单内);
#         2) 该效果所需的方向/子类型属性合法 (不是空标签、不是非法值);
#         3) 该效果在所选目标软件中确实会被渲染播放 (见 OFFICE_TARGET).
#       否则打开放映时该页不会播放任何切换 (静默忽略或退化),
#       这种"解析到有 transition 但实际不播放"的情况不计为有效切换动画.
#
# 口径选择 (用户要求: 动画必须能在实际办公软件中真正看到并启动播放):
#   OFFICE_TARGET = "both"  -> 只有 PowerPoint 和 WPS 两者都能播放的效果才计为有效.
#   实测背景 (以用户的 WPS 实际放映为准):
#     - 经用户在 WPS 中逐页确认: 本 PPT 中只有第18-21页的 uncover/揭开 效果
#       在 WPS 放映时看不到切换动画; 其余效果 (cover/comb/randomBar/wheel 等)
#       在该 WPS 中均能正常播放.
#     - 因此 wps 兼容性仅将 uncover 标为不渲染, 其他效果按可播放处理.
#       若后续在其他 WPS 版本发现更多不渲染效果, 再据实更新对应 "wps" 标记.
#
# 每个效果映射到:
#   - "dir_set":   dir 属性允许取值集合 (None = 不需要 dir; 若提供则必须在集合内)
#   - "orient_set": orient 属性允许取值 (split 用)
#   - "needs_dir": True 表示该效果必须带合法 dir 才会真正按方向播放
#   - "ppt":  PowerPoint 是否渲染播放
#   - "wps":  WPS 是否渲染播放
OFFICE_TARGET: str = "both"   # both | ppt | wps

PLAYABLE_TRANSITIONS = {
    # 无方向类
    "fade":      {"dir_set": None,                      "needs_dir": False, "ppt": True, "wps": True},
    "circle":    {"dir_set": None,                      "needs_dir": False, "ppt": True, "wps": True},
    "diamond":   {"dir_set": None,                      "needs_dir": False, "ppt": True, "wps": True},
    "dissolve":  {"dir_set": None,                      "needs_dir": False, "ppt": True, "wps": True},
    "random":    {"dir_set": None,                      "needs_dir": False, "ppt": True, "wps": True},
    "cut":       {"dir_set": None,                      "needs_dir": False, "ppt": True, "wps": True},
    "plus":      {"dir_set": None,                      "needs_dir": False, "ppt": True, "wps": True},
    "newsflash": {"dir_set": None,                      "needs_dir": False, "ppt": True, "wps": True},
    # 四方向类: 必须带合法 dir
    "push":      {"dir_set": {"l", "r", "u", "d"},      "needs_dir": True,  "ppt": True, "wps": True},
    "wipe":      {"dir_set": {"l", "r", "u", "d"},      "needs_dir": True,  "ppt": True, "wps": True},
    "cover":     {"dir_set": {"l", "r", "u", "d",
                              "lu", "ru", "ld", "rd"},  "needs_dir": True,  "ppt": True, "wps": True},
    "pull":      {"dir_set": {"l", "r", "u", "d",
                              "lu", "ru", "ld", "rd"},  "needs_dir": True,  "ppt": True, "wps": True},
    # uncover/揭开 — 经用户确认在该 WPS 中不渲染 (第18-21页看不到切换)
    "uncover":   {"dir_set": {"l", "r", "u", "d",
                              "lu", "ru", "ld", "rd"},  "needs_dir": True,  "ppt": True, "wps": False},
    "strips":    {"dir_set": {"lu", "ru", "ld", "rd"},  "needs_dir": True,  "ppt": True, "wps": True},
    # 横/纵向类: dir 取 horz/vert
    "blinds":    {"dir_set": {"horz", "vert"},          "needs_dir": True,  "ppt": True, "wps": True},
    "checker":   {"dir_set": {"horz", "vert"},          "needs_dir": True,  "ppt": True, "wps": True},
    "comb":      {"dir_set": {"horz", "vert"},          "needs_dir": True,  "ppt": True, "wps": True},
    "randomBar": {"dir_set": {"horz", "vert"},          "needs_dir": True,  "ppt": True, "wps": True},
    # split: 需要 orient(horz/vert) + dir(in/out)
    "split":     {"orient_set": {"horz", "vert"},
                  "dir_set": {"in", "out"},             "needs_dir": True,  "ppt": True, "wps": True},
    "wheel":     {"spokes": True,                       "needs_dir": False, "ppt": True, "wps": True},
    "zoom":      {"dir_set": {"in", "out"},             "needs_dir": False, "ppt": True, "wps": True},
    # morph/平滑 — PowerPoint 2016+ 的平滑切换 (Morph). 通常写在
    #   <mc:AlternateContent> 的 <mc:Choice> 内 (命名空间前缀多为 p159/p14),
    #   并配 <mc:Fallback> 里的标准切换 (如 fade) 供老版本/WPS 回退播放.
    #   属于一种真实可播放的切换动画.
    "morph":     {"dir_set": None,                      "needs_dir": False, "ppt": True, "wps": True},
    # ---- PowerPoint 2010+/2013+ 扩展切换效果 (命名空间 p14/p15) ----
    #   这些不在最初的 OOXML 标准 (p:) 命名空间里, 但在真实 PowerPoint 中
    #   均为可正常放映的切换. 经用户确认本 PPT 中这些页放映时都能看到切换动画,
    #   故全部按可播放处理; dir 均可选 (缺省时按默认方向播放).
    "wedge":     {"dir_set": None,                      "needs_dir": False, "ppt": True, "wps": True},
    "ripple":    {"dir_set": None,                      "needs_dir": False, "ppt": True, "wps": True},
    "prism":     {"dir_set": None,                      "needs_dir": False, "ppt": True, "wps": True},
    "doors":     {"dir_set": {"horz", "vert"},          "needs_dir": False, "ppt": True, "wps": True},
    "honeycomb": {"dir_set": None,                      "needs_dir": False, "ppt": True, "wps": True},
    "glitter":   {"dir_set": None,                      "needs_dir": False, "ppt": True, "wps": True},
    "flash":     {"dir_set": None,                      "needs_dir": False, "ppt": True, "wps": True},
    "shred":     {"dir_set": None,                      "needs_dir": False, "ppt": True, "wps": True},
    "reveal":    {"dir_set": None,                      "needs_dir": False, "ppt": True, "wps": True},
    "vortex":    {"dir_set": None,                      "needs_dir": False, "ppt": True, "wps": True},
    "switch":    {"dir_set": None,                      "needs_dir": False, "ppt": True, "wps": True},
    "flip":      {"dir_set": None,                      "needs_dir": False, "ppt": True, "wps": True},
    "gallery":   {"dir_set": None,                      "needs_dir": False, "ppt": True, "wps": True},
    "cube":      {"dir_set": None,                      "needs_dir": False, "ppt": True, "wps": True},
    "box":       {"dir_set": None,                      "needs_dir": False, "ppt": True, "wps": True},
    "rotate":    {"dir_set": None,                      "needs_dir": False, "ppt": True, "wps": True},
    "window":    {"dir_set": None,                      "needs_dir": False, "ppt": True, "wps": True},
    "ferris":    {"dir_set": None,                      "needs_dir": False, "ppt": True, "wps": True},
    "conveyor":  {"dir_set": None,                      "needs_dir": False, "ppt": True, "wps": True},
    "pan":       {"dir_set": None,                      "needs_dir": False, "ppt": True, "wps": True},
    "fracture":  {"dir_set": None,                      "needs_dir": False, "ppt": True, "wps": True},
    "warp":      {"dir_set": None,                      "needs_dir": False, "ppt": True, "wps": True},
    "flythrough":{"dir_set": None,                      "needs_dir": False, "ppt": True, "wps": True},
    "orbit":     {"dir_set": None,                      "needs_dir": False, "ppt": True, "wps": True},
    # prstTrans: PowerPoint 2013+ 预置切换容器, 具体效果在 prst 属性
    #   (如 pageCurlDouble/airplane/peelOff/origami/...). 真实可放映.
    "prstTrans": {"dir_set": None,                      "needs_dir": False, "ppt": True, "wps": True},
}


def _renders_in_target(spec):
    """根据 OFFICE_TARGET 判断该效果在目标办公软件中是否会被渲染播放."""
    if OFFICE_TARGET == "ppt":
        return spec.get("ppt", False)
    if OFFICE_TARGET == "wps":
        return spec.get("wps", False)
    # both: 两者都必须能播
    return spec.get("ppt", False) and spec.get("wps", False)


def is_playable(info):
    """
    判定一页的切换动画是否能在真实办公软件中启动并播放出来 (按 OFFICE_TARGET).

    返回 (ok: bool, reason: str).
    不通过 (ok=False) 的典型情况:
      - 根本没有 <p:transition>;
      - <p:transition> 为空, 没有任何切换效果子标签 (如 <p:transition/>);
      - 切换效果标签不在白名单内 (办公软件无法识别 -> 放映时不播放);
      - 该效果缺少必需的方向属性, 或方向/子类型取值非法 (放映时退化/不播放);
      - wheel 的 spokes 不是正整数; split 缺少合法 orient;
      - 该效果在目标软件中不渲染 (如 OFFICE_TARGET=both 时的 uncover, WPS 看不到).
    """
    if not info["present"]:
        return False, "无 <p:transition> 节点"
    eff = info["effect"]
    if not eff:
        return False, "transition 内无切换效果标签 (空切换, 放映时不播放)"
    if eff not in PLAYABLE_TRANSITIONS:
        return False, f"效果 '{eff}' 非标准 OOXML 切换, 办公软件放映时不识别/不播放"

    spec = PLAYABLE_TRANSITIONS[eff]
    # 目标软件是否渲染该效果 (如 OFFICE_TARGET=both 时, uncover 在 WPS 看不到)
    if not _renders_in_target(spec):
        miss = "WPS" if OFFICE_TARGET in ("both", "wps") and not spec.get("wps") else "PowerPoint"
        return False, f"效果 '{eff}' 在 {miss} 中不渲染播放 (放映时该页看不到切换动画)"

    # 解析该效果标签的属性
    attrs = {}
    if info["subtype"]:
        for kv in info["subtype"].split(";"):
            if "=" in kv:
                k, v = kv.split("=", 1)
                attrs[k] = v

    # wheel: spokes 必须是正整数
    if spec.get("spokes"):
        sp = attrs.get("spokes")
        if sp is not None:
            if not sp.isdigit() or int(sp) <= 0:
                return False, f"wheel spokes='{sp}' 非法, 放映时不播放"
        return True, "可播放"

    # split: 需要合法 orient + dir
    #   用户确认: 本 PPT 中缺 orient/dir 的 split 页放映时仍能看到切换动画,
    #   PowerPoint 会按默认 (horz + out) 播放. 故缺失时不判死, 仅在写了
    #   非法取值时才判不播.
    if "orient_set" in spec:
        orient = attrs.get("orient")
        d = attrs.get("dir")
        if orient is not None and orient not in spec["orient_set"]:
            return False, f"split orient='{orient}' 非法, 放映时不播放"
        if d is not None and d not in spec["dir_set"]:
            return False, f"split dir='{d}' 非法, 放映时不播放"
        return True, "可播放"

    # 方向类校验
    #   用户确认: 本 PPT 中缺 dir 的方向类效果 (strips/wipe/blinds/comb/pull/
    #   checker 等) 放映时仍能看到切换 (PowerPoint 用该效果的默认方向播放).
    #   故缺 dir 不再判死, 仅在写了"非法 dir 值"时才判不播.
    d = attrs.get("dir")
    if spec["dir_set"] is None:
        # 无方向效果: 无需 dir; 若误写了 dir 也不致命 (会被忽略)
        return True, "可播放"
    # 有方向的效果: 缺 dir 按默认方向可播; 写了非法 dir 才判不播
    if d is not None and d not in spec["dir_set"]:
        return False, f"效果 '{eff}' 方向 dir='{d}' 非法, 放映时不播放"
    return True, "可播放"


# ---------- 工具函数 ----------
def list_slides(z):
    """返回按页码排序的 slideN.xml 列表"""
    names = [n for n in z.namelist()
             if re.match(r"ppt/slides/slide\d+\.xml$", n)]
    return sorted(names, key=lambda x: int(re.search(r"slide(\d+)", x).group(1)))


def parse_transition(xml_bytes):
    """
    解析切换动画, 返回字典:
      {
        'present': bool,           # 是否存在 <p:transition>
        'effect': str,             # 效果类型 (fade/push/wipe/...)
        'subtype': str,            # 方向或子类型, 如 'l','r','horz-in', 'spokes=2'
        'speed': str,              # 'slow' / 'med' / 'fast'
        'duration': str,           # dur 属性 (ms), 若 PPT2010+ 标准
        'advClick': bool,          # 是否保留单击换页
        'advTm':   str | None,     # 自动换页毫秒数, None 表示未开启
        'sound':   bool,           # 是否带切换声音
        'raw':     str,            # 原始 transition XML
      }
    """
    info = {
        "present": False, "effect": "", "subtype": "",
        "speed": "", "duration": "",
        "advClick": True, "advTm": None,
        "sound": False, "raw": "",
    }
    xml = xml_bytes.decode("utf-8", errors="ignore")
    m = re.search(r"<p:transition\b[^>]*(?:/>|>.*?</p:transition>)", xml, re.S)
    if not m:
        return info
    info["present"] = True
    raw = m.group(0)
    info["raw"] = raw

    # 解析属性 (spd, dur, advClick, advTm)
    head = re.match(r"<p:transition\b([^>]*?)(/?)>", raw).group(1)
    for k, v in re.findall(r'(\w+)="([^"]*)"', head):
        if k == "spd":
            info["speed"] = v
        elif k == "dur":
            info["duration"] = v
        elif k == "advClick":
            info["advClick"] = (v != "0" and v.lower() != "false")
        elif k == "advTm":
            info["advTm"] = v

    # 效果类型 + 方向/子类型
    #   注意: 切换效果标签的命名空间前缀不一定是 "p:".
    #   例如平滑切换 (Morph) 常写作 <p159:morph .../> 或 <p14:...>,
    #   放在 <mc:AlternateContent>/<mc:Choice> 内. 因此这里匹配"任意前缀:标签"
    #   (可含冒号), 再剥掉前缀取本地名, 避免把带命名空间前缀的合法效果漏判成空切换.
    inner_tags = re.findall(r"<([A-Za-z_][\w.]*(?::\w+)?)([^/>]*)/?>", raw)
    # 过滤掉 transition / sndAc / sound 本身
    effect = None
    for tag, attrs in inner_tags:
        local = tag.split(":")[-1]  # 去掉命名空间前缀
        if local in ("transition", "sndAc", "stSnd", "endSnd", "snd"):
            continue
        if effect is None:
            effect = (local, attrs)
            break
    if effect:
        info["effect"] = effect[0]
        attrs = dict(re.findall(r'(\w+)="([^"]*)"', effect[1]))
        # 把所有子属性按 key 排序拼接, 形成稳定的 subtype
        info["subtype"] = ";".join(f"{k}={attrs[k]}" for k in sorted(attrs))

    # 声音
    if re.search(r"<p:sndAc>|<p:snd\b", raw):
        info["sound"] = True

    return info


def estimated_duration_sec(info):
    """估算切换持续时间(秒). 优先用 dur(ms), 其次用 spd."""
    if info["duration"]:
        try:
            return int(info["duration"]) / 1000.0
        except ValueError:
            pass
    return SPD_TO_SEC.get(info["speed"] or "med", 0.75)


def transition_signature(info):
    """切换效果组合 = 类型 + 方向/子类型"""
    return f"{info['effect']}|{info['subtype']}"


def transition_full_signature(info):
    """更严格: 类型+方向+持续时间+触发方式 (用于比较第1页与第35页)"""
    return (
        info["effect"], info["subtype"],
        info["speed"], info["duration"],
        info["advClick"], info["advTm"],
    )


def count_shapes(slide_xml):
    """统计 slide 中的形状/图片/表格/图表/页码占位符数量, 作为'原有内容'的近似量"""
    counts = {"sp": 0, "pic": 0, "tbl": 0, "chart": 0, "sldNum": 0}
    counts["sp"]    = len(re.findall(r"<p:sp\b", slide_xml))
    counts["pic"]   = len(re.findall(r"<p:pic\b", slide_xml))
    counts["tbl"]   = len(re.findall(r"<a:tbl\b", slide_xml))
    counts["chart"] = len(re.findall(r"c:chart\b", slide_xml))
    counts["sldNum"] = len(re.findall(r'type="sldNum"', slide_xml))
    return counts


# ---------- 维度1 ----------
def check_dim1(path):
    """
    维度1: 交付文件为 .pptx 格式, 能够正常打开.
      - 扩展名必须为 .pptx
      - 能作为 OOXML (zip) 正常解析打开
      - 至少存在一页幻灯片
    返回 (ok: bool, reasons: list[str], z: ZipFile|None, slides: list[str])
    """
    reasons: list[str] = []
    ext = os.path.splitext(path)[1].lower()
    if ext != ".pptx":
        return False, [f"扩展名 {ext} 不是 .pptx"], None, []

    if not os.path.exists(path):
        return False, ["文件不存在"], None, []

    try:
        z = zipfile.ZipFile(path)
    except zipfile.BadZipFile:
        return False, ["文件不是有效的 pptx 压缩包, 无法打开"], None, []

    slides = list_slides(z)
    if len(slides) == 0:
        return False, ["未发现任何幻灯片页面, 无法正常打开"], z, slides

    return (len(reasons) == 0), reasons, z, slides


# ---------- 维度2 ----------

# 评分细则原文 (用于命中时按字面打印)
RULE = {
    "+1_cover":   "第1页封面：设置一种切换动画。",
    "+1_end":     "第35页结束页：切换动画的类型、方向、持续时间和触发方式与第1页完全相同。",
    "+1_each":    "第2-34页皆设置切换动画（一页有切换动画+1分，两页有切换动画则+2分,以此类推）",
    "+5_unique":  ("第2至第34页切换唯一性：33页的“动画类型+方向或子类型”组合均不完全相同。"
                   "第2至第34页均不使用第1页、第35页共同采用的切换效果组合。"),
    "+5_trigger": ("全部35页切换触发：均保留“单击鼠标后换页”功能。未启用固定时间自动换页，"
                   "讲解时不会自行跳转。持续时间约0.5至1.5秒，速度适中，不出现过慢等待。"),
}


def check_dim2(z, slides):
    """返回 (score, hits: list[tuple[int, str]])"""
    hits: list[tuple[int, str]] = []
    total = 0
    n_pages = len(slides)

    # 预先解析每页 transition + xml
    trans = []
    slide_xmls = []
    for n in slides:
        xml = z.read(n).decode("utf-8", errors="ignore")
        slide_xmls.append(xml)
        trans.append(parse_transition(xml.encode("utf-8")))

    # 预先判定每页切换动画"能否在真实办公软件中启动播放"
    #   playable[i] = True  表示 PowerPoint/WPS 打开放映时该页会真正播放出切换动画;
    #   playable[i] = False 表示虽然 XML 里写了 <p:transition>, 但放映时不会播放
    #                       (空切换 / 非标准效果 / 缺失或非法方向等).
    # 后续所有"该页是否设置了切换动画"的判定, 一律以 playable 为准,
    # 不再用"只要解析到 transition 就算数"的宽松口径.
    playable = []
    play_reasons = []
    for t in trans:
        ok, why = is_playable(t)
        playable.append(ok)
        play_reasons.append(why)

    # ===== 加分项 =====

    # +1  第1页封面: 设置一种切换动画
    #     细则要点: 是"第1页封面" + "设置" + "一种切换动画"
    #       - 第1页存在 <p:transition>
    #       - 切换效果有且仅有一种 (transition 子元素中切换效果标签恰好 1 个)
    #       - 且该效果能在真实办公软件中启动播放 (playable)
    #     注意 <mc:AlternateContent> 兼容写法: 同一种切换 (如平滑 Morph) 会同时
    #       写在 <mc:Choice> (新版效果) 和 <mc:Fallback> (老版回退效果) 内,
    #       XML 里因此出现 2 个 <p:transition> 块, 但放映时只播其中一个,
    #       仍属于"一种切换动画". 故这里若检测到 AlternateContent 包裹, 则视为一种;
    #       否则要求 transition 块恰好 1 个.
    slide1_xml = slide_xmls[0]
    s1_trans_blocks = re.findall(
        r"<p:transition\b[^>]*(?:/>|>.*?</p:transition>)", slide1_xml, re.S)
    s1_alt_content = bool(re.search(r"<mc:AlternateContent\b", slide1_xml))
    s1_effect_tags = []
    if s1_trans_blocks:
        # 统计 transition 内部的"切换效果"标签 (排除声音相关 sndAc/snd/stSnd/endSnd),
        # 兼容任意命名空间前缀 (如 p159:morph)
        inner = re.findall(r"<([A-Za-z_][\w.]*(?::\w+)?)\b[^/>]*/?>", s1_trans_blocks[0])
        s1_effect_tags = [t for t in inner
                          if t.split(":")[-1] not in
                          ("transition", "sndAc", "snd", "stSnd", "endSnd")]
    # "一种切换": 要么只有 1 个 transition 块; 要么是 AlternateContent 兼容写法
    # (多个块但表述的是同一种切换).
    one_transition = (len(s1_trans_blocks) == 1) or s1_alt_content
    if (one_transition
            and len(s1_effect_tags) == 1
            and trans[0]["present"]
            and trans[0]["effect"]
            and playable[0]):
        total += 1
        hits.append((+1, RULE["+1_cover"]))

    # +1  第35页结束页: 切换动画的"类型、方向、持续时间、触发方式"与第1页完全相同
    #     细则要点 (4 个点必须全部相同):
    #       1) 类型     -> transition 效果标签名 (effect)
    #       2) 方向     -> 该效果的方向/子类型属性 (subtype)
    #       3) 持续时间 -> dur (ms) 或 spd (slow/med/fast)
    #       4) 触发方式 -> advClick (是否单击) + advTm (是否自动换页时长)
    if n_pages >= 35:
        t1, t35 = trans[0], trans[34]
        same_type     = (t1["effect"]   == t35["effect"])
        same_dir      = (t1["subtype"]  == t35["subtype"])
        same_duration = (t1["duration"] == t35["duration"]
                         and t1["speed"] == t35["speed"])
        same_trigger  = (t1["advClick"] == t35["advClick"]
                         and t1["advTm"]   == t35["advTm"])
        # 两页都必须能在真实办公软件中启动播放, 否则"相同"也无意义
        if (same_type and same_dir and same_duration and same_trigger
                and playable[0] and playable[34]):
            total += 1
            hits.append((+1, RULE["+1_end"]))

    # +1  第2-34页皆设置切换动画
    #     细则要点: 检查第2-34页中"总共有几页出现了切换动画", 有几页就加几分
    #              (一页+1分, 两页+2分, 以此类推)
    #     口径: "设置切换动画" = 该页切换动画能在真实办公软件 (PowerPoint/WPS) 中
    #           启动并播放出来 (playable=True). 仅在 XML 里解析到 <p:transition>
    #           但放映时不会播放的, 不计入.
    pages_no_anim: list[tuple[int, str]] = []
    cnt = 0
    for idx, t in enumerate(trans[1:34], start=2):
        if playable[idx - 1]:
            cnt += 1
        else:
            pages_no_anim.append((idx, play_reasons[idx - 1]))
    if cnt > 0:
        awarded = min(cnt, 33)
        total += awarded
        hits.append((+awarded, RULE["+1_each"]))

    # +5  第2至第34页切换唯一性
    #     前置条件 (必须先全部满足, 否则不给分, 也不再检查后两点):
    #       (*) 第2-34页这 33 页每一页都设置了切换效果 (transition 存在且 effect 非空)
    #           即 33 页全有切换效果, 数量正好 33 个.
    #     细则两个要点 (前置满足后, 两点同时满足才 +5):
    #       (a) 33 页的"动画类型 + 方向或子类型"组合均不完全相同
    #       (b) 第2-34页均不使用"第1页、第35页采用的切换效果组合"
    #           严格口径: 第1页用的组合、第35页用的组合, 两者任一都不能在第2-34页出现
    #           (即使第1页与第35页用的是不同组合, 两个组合也都禁用)
    if n_pages >= 35:
        mid = trans[1:34]  # 第2..第34, 共 33 页

        # (*) 前置: 33 页必须全部"能在真实办公软件中启动播放"的切换效果
        all_have_effect = all(playable[i] for i in range(1, 34))

        if all_have_effect:
            sigs = [transition_signature(t) for t in mid]

            # (a) 33 个组合两两不完全相同
            unique_ok = (len(set(sigs)) == 33)

            # (b) 第1页与第35页采用的组合, 两者都禁止在第2-34页出现
            forbidden: set[str] = set()
            if trans[0]["present"]:
                forbidden.add(transition_signature(trans[0]))
            if trans[34]["present"]:
                forbidden.add(transition_signature(trans[34]))
            no_reuse_ok = forbidden.isdisjoint(sigs)

            if unique_ok and no_reuse_ok:
                total += 5
                hits.append((+5, RULE["+5_unique"]))

    # +5  全部35页切换触发
    #     细则三个要点 (全部满足才 +5):
    #       (1) 均保留"单击鼠标后换页"功能          -> 每页 advClick=True
    #       (2) 未启用固定时间自动换页, 讲解时不会自行跳转
    #                                              -> 每页 advTm 未设置 (None/空)
    #       (3) 持续时间约 0.5 至 1.5 秒, 速度适中, 不出现过慢等待
    #                                              -> 每页 estimated_duration_sec() ∈ [0.5, 1.5]
    #     注: 仅对"能在真实办公软件中启动播放"的页 (playable) 判定持续时间/触发,
    #         因为无法播放的页其速度设置没有实际意义.
    pages_click_off  = [i for i, t in enumerate(trans, 1)
                        if not (playable[i - 1] and t["advClick"])]
    pages_auto_adv   = [(i, t["advTm"]) for i, t in enumerate(trans, 1)
                        if t["advTm"] not in (None, "")]
    pages_dur_bad    = [(i, round(estimated_duration_sec(t), 3))
                        for i, t in enumerate(trans, 1)
                        if playable[i - 1] and not (0.5 <= estimated_duration_sec(t) <= 1.5)]

    click_ok    = (len(pages_click_off) == 0)
    autotime_ok = (len(pages_auto_adv)  == 0)
    duration_ok = (len(pages_dur_bad)   == 0)

    if click_ok and autotime_ok and duration_ok:
        total += 5
        hits.append((+5, RULE["+5_trigger"]))

    return total, hits


# ---------- 主流程 ----------
# 维度二加分/扣分项规则清单 (与 check_dim2 里的判定一一对应).
# 用于在返回结构里同时列出"命中"与"未命中"项, 供批量运行时对齐评分矩阵.
DIM2_RULES = [
    {"key": "+1_cover",   "max_delta": +1, "rule": RULE["+1_cover"]},
    {"key": "+1_end",     "max_delta": +1, "rule": RULE["+1_end"]},
    {"key": "+1_each",    "max_delta": +33, "rule": RULE["+1_each"]},
    {"key": "+5_unique",  "max_delta": +5, "rule": RULE["+5_unique"]},
    {"key": "+5_trigger", "max_delta": +5, "rule": RULE["+5_trigger"]},
]


def _find_target_file(dir_path: str) -> str:
    """在给定目录内定位被评估的 .pptx / .ppt 文档.

    约定: 目录内应只放一个待评估的 PPT 文件. 若存在多个, 取第一个 (按名称排序);
    若没有, 返回空字符串.
    """
    if not os.path.isdir(dir_path):
        return ""
    candidates = [
        n for n in os.listdir(dir_path)
        if os.path.splitext(n)[1].lower() in (".pptx", ".ppt")
        and not n.startswith("~$")  # 排除临时文件
    ]
    candidates.sort()
    return os.path.join(dir_path, candidates[0]) if candidates else ""


def evaluate(dir_path: str) -> dict[str, object]:
    """批量运行入口: 接收脚本所在目录的路径, 在该目录内定位并评估 PPT 文档.

    返回结构见 “脚本接口差异与统一建议.md §2.2”.
    """
    max_score = sum(int(r["max_delta"]) for r in DIM2_RULES if int(r["max_delta"]) > 0)
    result: dict[str, object] = {
        "id": "067",
        "file_name": "",
        "status": "ok",
        "error": None,
        "dim1_pass": False,
        "dim1_reason": "",
        "dim2_items": [],
        "total_score": 0,
        "max_score": max_score,
    }

    try:
        file_path = _find_target_file(dir_path)
        if not file_path:
            result["status"] = "error"
            result["error"] = f"目录 {dir_path} 内未找到 .pptx/.ppt 文件"
            return result
        result["file_name"] = os.path.basename(file_path)

        ok, reasons, z, slides = check_dim1(file_path)
        result["dim1_pass"] = ok
        if not ok:
            result["dim1_reason"] = "; ".join(reasons)
            result["total_score"] = 0
            return result

        score, hits = check_dim2(z, slides)
        # 把命中项按规则 key 归位; 未命中项 delta=0, hit=False
        hit_map: dict[str, tuple[int, str]] = {}
        # hits 是 (delta, rule_text) 列表, 需要按 rule_text 反查 key
        text_to_key = {RULE[k]: k for k in RULE}
        for delta, rule_text in hits:
            key = text_to_key.get(rule_text, rule_text)
            hit_map[key] = (delta, rule_text)

        items = []
        for r in DIM2_RULES:
            key = r["key"]
            if key == "+1_each":
                awarded_count = max(0, min(int(hit_map.get(key, (0, ""))[0]), 33))
                for unit_index in range(1, 34):
                    hit = unit_index <= awarded_count
                    items.append({
                        "rule": f"{r['rule']}（计数项 {unit_index}/33）",
                        "max_delta": 1,
                        "delta": 1 if hit else 0,
                        "hit": hit,
                        "detail": "",
                    })
                continue
            if key in hit_map:
                delta, _ = hit_map[key]
                items.append({
                    "rule": r["rule"],
                    "max_delta": r["max_delta"],
                    "delta": delta,
                    "hit": True,
                    "detail": "",
                })
            else:
                items.append({
                    "rule": r["rule"],
                    "max_delta": r["max_delta"],
                    "delta": 0,
                    "hit": False,
                    "detail": "",
                })
        result["dim2_items"] = items
        result["total_score"] = score
        return result
    except Exception as e:  # noqa: BLE001 — 顶层兜底, 交给批量运行器识别
        result["status"] = "error"
        result["error"] = f"{type(e).__name__}: {e}"
        return result


if __name__ == "__main__":
    # 本地调试: 传入脚本所在目录的路径 (默认当前脚本所在目录).
    _dir = sys.argv[1] if len(sys.argv) > 1 else os.path.dirname(os.path.abspath(__file__))
    print(json.dumps(evaluate(_dir), ensure_ascii=False, indent=2))
