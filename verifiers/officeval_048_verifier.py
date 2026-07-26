# -*- coding: utf-8 -*-
"""自动评估 PPT 文件评分脚本

对外只暴露一个函数 :func:`evaluate`，签名统一为::

    def evaluate(dir_path: str) -> dict

其中 ``dir_path`` 是脚本所在目录（由 runner 传入），脚本自身负责在该目录里
定位并打开被评估的 ``.pptx`` 文档。返回结构见《脚本接口差异与统一建议》§2.2。
"""
import sys, os, json, zipfile, xml.etree.ElementTree as ET

SCRIPT_ID = '048'
DOC_EXTS = ('.pptx',)
NS = {
    'p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
    'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
}

# 8条案例文字关键词（按顺序）
CASE_KEYWORDS = ['阿以战争', '苏伊士运河', '六日战争', '两伊战争', '海湾战争', '伊拉克战争', '叙利亚', '加沙']
CASE_YEAR_MARKERS = ['1948', '1956', '1967', '1980', '1990', '2003', '2011', '2023']


def get_all_text(root):
    return [''.join(t.text or '' for t in sp.findall('.//a:t', NS)).strip()
            for sp in root.findall('.//p:sp', NS)]


def dim1_check(file_name: str) -> list[str]:
    """维度1：可用与可修改性"""
    issues: list[str] = []
    # 1a. 交付文件为 .pptx 格式
    if not file_name.lower().endswith(DOC_EXTS):
        issues.append('文件格式不是 .pptx')
    return issues


def dim2_check(z):
    scores = []  # [(delta, label, hit)]

    slide1_xml = z.read('ppt/slides/slide1.xml')
    root1 = ET.fromstring(slide1_xml)
    all_texts_1 = get_all_text(root1)
    all_text_joined = ' '.join(all_texts_1)

    # ── +5: 第1页初始放映状态 ──────────────────────────────────────────
    # 细则逐条：
    # 1. 进入第1页时，标题、背景、底部课程信息和时间轴基础线可先显示
    # 2. 8条案例内容进入时未全部同时显示（初始隐藏）
    # 3. 放映模式下点击鼠标1次后，8条案例开始按顺序自动连续出现
    # 4. 不需要再次点击，动画全部通过"上一动画之后"或等效自动触发方式连续播放
    # 5. 放映顺序：1948阿以战争→1956苏伊士运河危机→1967六日战争→
    #              1980两伊战争→1990海湾战争→2003伊拉克战争→2011叙利亚→2023加沙
    s5_ok = True
    s5_reasons = []

    # 构建 spid -> text 映射（后续多处复用）
    spid_text = {}
    for sp in root1.findall('.//p:sp', NS):
        cNvPr = sp.find('.//p:cNvPr', NS)
        if cNvPr is None:
            continue
        sid = cNvPr.get('id')
        txt = ''.join(t.text or '' for t in sp.findall('.//a:t', NS)).strip()
        if txt:
            spid_text[sid] = txt

    timing = root1.find('.//p:timing', NS)

    # 收集所有初始隐藏 shape（动画中被 set visibility=visible 的）
    initially_hidden_spids = set()
    if timing is not None:
        for s in timing.findall('.//p:set', NS):
            attr = s.find('.//p:attrName', NS)
            to_val = s.find('.//p:to/p:strVal', NS)
            tgt = s.find('.//p:spTgt', NS)
            if (attr is not None and 'visibility' in (attr.text or '') and
                    to_val is not None and to_val.get('val') == 'visible' and
                    tgt is not None):
                initially_hidden_spids.add(tgt.get('spid'))

    # 从 mainSeq 按文档顺序提取动画组
    mainSeq = None
    if timing is not None:
        for seq in timing.findall('.//p:seq', NS):
            cTn = seq.find('p:cTn', NS)
            if cTn is not None and cTn.get('nodeType') == 'mainSeq':
                mainSeq = seq
                break

    # 新增全局判定规则：
    # - 细则里凡是出现”动画”要求，必须是办公软件动画窗格可见的真实动画。
    #   OOXML 中动画窗格可见动画：animEffect / anim / animMotion / animScale / animClr / animRot；
    #   纯 set style.visibility 只在XML层切换可见性，不在动画窗格显示，不计为真实动画。
    # - 若该评分细则没有指定动画持续时间，则默认持续时间不能小于0.5秒（500ms）。
    PANE_VISIBLE_ANIM_TAGS = {'animEffect', 'anim', 'animMotion', 'animScale', 'animClr', 'animRot'}
    DEFAULT_MIN_DUR_MS = 500  # 细则未指定时长时的最低要求

    click_groups = []   # nodeType=clickEffect
    after_groups = []   # nodeType=afterEffect（按文档顺序）
    if mainSeq is not None:
        children_node = mainSeq.find('p:cTn/p:childTnLst', NS)
        if children_node is not None:
            for child in children_node:
                cTn = child.find('p:cTn', NS)
                if cTn is None:
                    continue
                nt = cTn.get('nodeType', '')
                if nt == 'clickEffect':
                    click_groups.append(child)
                elif nt == 'afterEffect':
                    after_groups.append(child)

    # ── 细则点1+2：标题/背景等先显示，8条案例初始未全同时显示 ──────────
    # 验证方式：8条案例的核心文字 shape 均在 initially_hidden_spids 中
    CASE_NAMES = ['阿以战争', '苏伊士运河危机', '六日战争', '两伊战争',
                  '海湾战争', '伊拉克战争', '叙利亚冲突', '加沙冲突']
    # 找到每条案例对应的 spid
    case_spids_ordered = []
    for name in CASE_NAMES:
        for sid, txt in spid_text.items():
            if name in txt or (name == '苏伊士运河危机' and '苏伊士运河' in txt):
                case_spids_ordered.append(sid)
                break
        else:
            case_spids_ordered.append(None)

    all_cases_found = all(sid is not None for sid in case_spids_ordered)
    if not all_cases_found:
        missing = [CASE_NAMES[i] for i, sid in enumerate(case_spids_ordered) if sid is None]
        s5_ok = False
        s5_reasons.append(f'第1页缺少案例文字: {missing}')
    else:
        s5_reasons.append(f'8条案例文字均存在于第1页')

    # 8条案例 shape 初始均隐藏（未全部同时显示）
    cases_initially_hidden = [sid in initially_hidden_spids
                               for sid in case_spids_ordered if sid is not None]
    if all(cases_initially_hidden):
        s5_reasons.append('8条案例内容进入时均初始隐藏，不同时显示 ✓')
    else:
        visible_cases = [CASE_NAMES[i] for i, hidden in enumerate(cases_initially_hidden) if not hidden]
        s5_ok = False
        s5_reasons.append(f'以下案例进入时已直接显示（未隐藏）: {visible_cases}')

    # ── 细则点1+补充：动画必须是动画窗格可见的真实动画，且无时长规定时默认≥500ms ──
    # 检查8组动画是否含真实动画（PANE_VISIBLE_ANIM_TAGS）；
    # 若仅有 set visibility，则不计为动画窗格可见动画，不满足"动画"要求。
    groups_with_real_anim = 0
    groups_min_dur_ok = 0
    for par in (click_groups + after_groups):
        real_anims = [elem for elem in par.iter()
                      if elem.tag.split('}')[-1] in PANE_VISIBLE_ANIM_TAGS]
        if real_anims:
            groups_with_real_anim += 1
            # 细则对+5未指定每条动画持续时间，用 DEFAULT_MIN_DUR_MS
            durs = [int(ctn.get('dur')) for ctn in par.findall('.//p:cTn', NS)
                    if ctn.get('dur') and ctn.get('dur') != 'indefinite']
            if durs and max(durs) >= DEFAULT_MIN_DUR_MS:
                groups_min_dur_ok += 1
    if groups_with_real_anim == 0:
        s5_ok = False
        s5_reasons.append('8组动画均为纯 set visibility，不在动画窗格显示，不计为真实动画')
    else:
        s5_reasons.append(f'{groups_with_real_anim}组含动画窗格可见动画')
        if groups_min_dur_ok < groups_with_real_anim:
            s5_ok = False
            s5_reasons.append(f'部分动画持续时间 < {DEFAULT_MIN_DUR_MS}ms（细则未指定时长，默认最低{DEFAULT_MIN_DUR_MS}ms）')
        visible_cases = [CASE_NAMES[i] for i, hidden in enumerate(cases_initially_hidden) if not hidden]
        s5_ok = False
        s5_reasons.append(f'以下案例进入时已直接显示（未隐藏）: {visible_cases}')

    # ── 细则点3：点击鼠标1次后8条案例开始出现 ──────────────────────────
    # 恰好有1个 clickEffect 组，且该组触发了第1条案例
    if len(click_groups) != 1:
        s5_ok = False
        s5_reasons.append(f'clickEffect 动画组数量为 {len(click_groups)}，应为1个（单击1次触发）')
    else:
        # 检查 clickEffect 组是否包含第1条案例（阿以战争）的 shape
        first_click_spids = {sp.get('spid') for sp in click_groups[0].findall('.//p:spTgt', NS)}
        if case_spids_ordered[0] in first_click_spids:
            s5_reasons.append('点击1次触发第1条案例（阿以战争）✓')
        else:
            s5_ok = False
            s5_reasons.append(f'clickEffect 组未包含第1条案例 shape（spid={case_spids_ordered[0]}）')

    # ── 细则点4：不需要再次点击，后续通过"上一动画之后"自动触发 ─────────
    # 所有后续组均为 afterEffect（delay为数值，非 indefinite）
    non_auto = []
    for i, par in enumerate(after_groups):
        cTn = par.find('p:cTn', NS)
        if cTn is None:
            continue
        cond = cTn.find('.//p:stCondLst/p:cond', NS)
        delay = cond.get('delay', '') if cond is not None else ''
        if delay == 'indefinite' or delay == '':
            non_auto.append(i + 2)  # 第几条案例
    if non_auto:
        s5_ok = False
        s5_reasons.append(f'第{non_auto}组 afterEffect 延迟为 indefinite，需要额外点击')
    else:
        s5_reasons.append(f'{len(after_groups)} 个后续组均为自动触发（afterEffect，delay为数值）✓')

    # ── 细则点5：放映顺序 1948→1956→1967→1980—1988→1990—1991→2003→2011→2023 ────
    # 按文档顺序：click_groups[0] 是第1条，after_groups[0..6] 是第2-8条
    # 每条要求：年份（含范围时两端年份均出现）+ 关键词 同时命中，缺一不可。
    expected_keywords = [
        (('1948',),         '阿以战争'),
        (('1956',),         '苏伊士运河'),
        (('1967',),         '六日战争'),
        (('1980', '1988'),  '两伊战争'),
        (('1990', '1991'),  '海湾战争'),
        (('2003',),         '伊拉克战争'),
        (('2011',),         '叙利亚'),
        (('2023',),         '加沙'),
    ]
    all_anim_groups = click_groups + after_groups  # 共8组，按顺序
    order_ok = True
    order_details = []
    for i, (years, keyword) in enumerate(expected_keywords):
        expected_label = '—'.join(years) + keyword
        if i >= len(all_anim_groups):
            order_ok = False
            order_details.append(f'第{i+1}组（{expected_label}）缺失')
            continue
        group_spids = {sp.get('spid') for sp in all_anim_groups[i].findall('.//p:spTgt', NS)}
        group_texts = ' '.join(spid_text.get(sid, '') for sid in group_spids)
        missing_years = [y for y in years if y not in group_texts]
        keyword_ok = keyword in group_texts
        if not missing_years and keyword_ok:
            order_details.append(f'第{i+1}条={expected_label} ✓')
        else:
            order_ok = False
            miss_parts = []
            if missing_years:
                miss_parts.append(f'年份缺失{missing_years}')
            if not keyword_ok:
                miss_parts.append(f'关键词"{keyword}"未出现')
            order_details.append(
                f'第{i+1}条期望{expected_label}，{"；".join(miss_parts)}，实际文字={group_texts[:30]!r}'
            )
    if order_ok:
        s5_reasons.append(
            '8条案例放映顺序正确: '
            + '、'.join('—'.join(ys) + k for ys, k in expected_keywords)
        )
    else:
        s5_ok = False
        s5_reasons.append('案例放映顺序存在问题:')
        s5_reasons.extend(f'  {d}' for d in order_details)

    scores.append((5, '+5 第1页初始放映状态', s5_ok, s5_reasons))

    # ── +3: 第1页动画速度 ──────────────────────────────────────────────
    # 细则逐条：
    # 1. 每条案例出现动画持续时间 0.1—0.5秒（100—500ms）
    # 2. 上一条结束后 0—0.5秒（0—500ms）内下一条自动开始
    # 3. 8条案例总时长 2—5秒（2000—5000ms）
    # 4. 动画效果为淡入/擦除/飞入/缩放/强调闪现之一，整体风格统一
    # 5. 节奏由前到后保持快速，或后半段速度略加快（不变慢）
    s3a_ok = True
    s3a_reasons = []

    # 8组动画（click_groups + after_groups 按顺序）
    all_8_groups = click_groups + after_groups  # 共8组

    # 判断每组的动画效果类型和持续时间
    # OOXML 动画效果映射（允许的类型）:
    #   animEffect filter="fade*" -> 淡入
    #   animEffect filter="wipe*"/"strips*" -> 擦除
    #   animEffect filter="fly*"/"blinds*" 等 -> 飞入
    #   animScale -> 缩放
    #   set style.visibility (dur≈1ms, 即瞬间出现) -> 强调闪现
    #   anim (强调类) -> 强调
    ALLOWED_EFFECT_TAGS = {'animEffect', 'animScale', 'anim', 'animMotion'}
    FLASH_THRESHOLD_MS = 50  # dur ≤ 50ms 认定为闪现效果

    group_dur_list = []    # 每组有效动画持续时间（ms）
    group_delay_list = []  # 每组与上一组的间隔（ms）
    group_effects = []     # 每组效果类型描述

    for par in all_8_groups:
        cTn_top = par.find('p:cTn', NS)
        # 组间延迟
        cond = cTn_top.find('.//p:stCondLst/p:cond', NS) if cTn_top is not None else None
        try:
            delay_val = int(cond.get('delay', 0)) if cond is not None else 0
        except ValueError:
            delay_val = 0
        group_delay_list.append(delay_val)

        # 动画效果类型
        effect_tags = {elem.tag.split('}')[-1] for elem in par.iter()
                       if elem.tag.split('}')[-1] in ALLOWED_EFFECT_TAGS}
        # 所有 cTn dur 值（非 indefinite）
        durs = [int(ctn.get('dur')) for ctn in par.findall('.//p:cTn', NS)
                if ctn.get('dur') and ctn.get('dur') != 'indefinite']
        max_dur = max(durs) if durs else 0

        # 判断是否为闪现（纯 set visibility，dur 极小）
        has_set_visibility = any(
            elem.tag.split('}')[-1] == 'set' and
            (elem.find('.//p:attrName', NS) is not None) and
            'visibility' in (elem.find('.//p:attrName', NS).text or '')
            for elem in par.iter()
        )
        is_flash = has_set_visibility and not effect_tags and max_dur <= FLASH_THRESHOLD_MS

        if is_flash:
            group_effects.append('闪现')
            # 闪现效果本身瞬间完成，有效动画时长视为 max_dur（1ms）
            group_dur_list.append(max_dur)
        elif effect_tags:
            group_effects.append('+'.join(sorted(effect_tags)))
            group_dur_list.append(max_dur)
        else:
            group_effects.append('未知')
            group_dur_list.append(max_dur)

    # ── 细则点1：每条案例动画持续时间 100—500ms ────────────────────────
    # 闪现效果（set visibility dur≈1ms）本身即为瞬间出现，时长不适用常规动画范围。
    # 但细则要求 0.1-0.5s，闪现不满足此条。
    bad_durs = []
    for i, (d, ef) in enumerate(zip(group_dur_list, group_effects)):
        if not (100 <= d <= 500):
            bad_durs.append(f'第{i+1}条:{d}ms({ef})')
    if bad_durs:
        s3a_ok = False
        s3a_reasons.append(f'以下案例动画持续时间不在100—500ms: {bad_durs}')
    else:
        s3a_reasons.append(f'各案例动画持续时间 {group_dur_list} ms，均在100—500ms ✓')

    # ── 细则点2：组间延迟 0—500ms ──────────────────────────────────────
    # 第1组（clickEffect）本身无"上一条"，从第2组开始检查
    inter_delays = group_delay_list[1:]  # after_groups 的 delay
    bad_delays = [d for d in inter_delays if not (0 <= d <= 500)]
    if bad_delays:
        s3a_ok = False
        s3a_reasons.append(f'以下组间延迟不在0—500ms: {bad_delays}')
    else:
        s3a_reasons.append(f'各组间延迟 {inter_delays} ms，均在0—500ms ✓')

    # ── 细则点3：8条总时长 2—5秒 ────────────────────────────────────────
    # 总时长 = 第1组dur + Σ(后7组的 delay + dur)
    total_ms = group_dur_list[0] + sum(
        group_delay_list[i] + group_dur_list[i] for i in range(1, len(all_8_groups))
    ) if all_8_groups else 0
    if 2000 <= total_ms <= 5000:
        s3a_reasons.append(f'8条案例总时长约 {total_ms}ms，在2—5秒范围内 ✓')
    else:
        s3a_ok = False
        s3a_reasons.append(f'8条案例总时长约 {total_ms}ms，不在2—5秒(2000—5000ms)范围内')

    # ── 细则点4：动画效果为允许类型之一，整体风格统一 ──────────────────
    # rubric 明确要求“整体风格统一”，因此 8 条动画必须收敛为同一种效果；
    # 均为允许类型但混用多种（如淡入+缩放），也判为不统一，扣分。
    ALLOWED_EFFECT_NAMES = {'淡入', '擦除', '飞入', '缩放', '闪现',
                             'animEffect', 'animScale', 'anim', 'animMotion'}
    unique_effects = set(group_effects)
    all_allowed = all(
        ef in ALLOWED_EFFECT_NAMES or any(a in ef for a in ALLOWED_EFFECT_NAMES)
        for ef in unique_effects
    )
    if all_allowed and len(unique_effects) == 1:
        s3a_reasons.append(f'动画效果统一，均为"{next(iter(unique_effects))}" ✓')
    elif all_allowed:
        s3a_ok = False
        s3a_reasons.append(f'动画效果均为允许类型，但存在多种，未做到整体风格统一: {unique_effects}')
    else:
        s3a_ok = False
        s3a_reasons.append(f'存在不允许的动画效果类型: {unique_effects - ALLOWED_EFFECT_NAMES}')

    # ── 细则点5：节奏不变慢（后半段 delay+dur 均值 ≤ 前半段 * 1.2） ────
    if len(all_8_groups) >= 8:
        # 每组"出现间隔" = 本组delay + 本组dur（从第2组起）
        intervals = [group_delay_list[i] + group_dur_list[i] for i in range(1, 8)]
        first_half_avg = sum(intervals[:3]) / 3   # 第2-4条
        second_half_avg = sum(intervals[4:7]) / 3  # 第6-8条
        if second_half_avg <= first_half_avg * 1.2:
            s3a_reasons.append(
                f'节奏未变慢（前半均值{first_half_avg:.0f}ms，后半均值{second_half_avg:.0f}ms）✓')
        else:
            s3a_ok = False
            s3a_reasons.append(
                f'后半段节奏明显变慢（前半均值{first_half_avg:.0f}ms，后半均值{second_half_avg:.0f}ms）')

    scores.append((3, '+3 第1页动画速度', s3a_ok, s3a_reasons))

    # ── +5: 第1页最后案例音效 ────────────────────────────────────────────
    # 细则逐条：
    # 1. 第8条"2023以来 加沙冲突"出现时同步播放战争爆炸音效，音效与最后案例动画绑定
    # 2. 音效嵌入PPT或随PPT打包保存（Internal 关系且包内存在），换电脑不丢失
    # 3. 爆炸音效时长约1—4秒，不循环播放
    # 4. 音频图标可隐藏或位于页面边缘，不遮挡时间线/案例文字/标题/底部信息
    s5b_ok = True
    s5b_reasons = []

    import struct

    REL_NS_ID = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'
    REL_PKG_NS = 'http://schemas.openxmlformats.org/package/2006/relationships'

    # 第8组 = after_groups[-1]，必须对应 2023/加沙
    last_group = after_groups[-1] if after_groups else None
    last_group_texts = ''
    if last_group is not None:
        group_spids = {sp.get('spid') for sp in last_group.findall('.//p:spTgt', NS)}
        last_group_texts = ' '.join(spid_text.get(sid, '') for sid in group_spids)
    if last_group is None or ('加沙' not in last_group_texts and '2023' not in last_group_texts):
        s5b_ok = False
        s5b_reasons.append('最后一个动画组不是第8条"2023以来 加沙冲突"')

    # a) 与第8条动画绑定的音效：同 par 节点内的 <p:sndTgt r:embed=?/>
    #    以及配套的 <p:audio> 媒体节点（用于判 loop）
    sound_embeds = []           # [(rid, snd_elem), ...]
    audio_media_nodes = []      # [<p:audio> ...]
    if last_group is not None:
        for elem in last_group.iter():
            tag = elem.tag.split('}')[-1]
            if tag == 'sndTgt':
                rid = elem.get(REL_NS_ID + 'embed')
                if rid:
                    sound_embeds.append((rid, elem))
            elif tag == 'audio':
                audio_media_nodes.append(elem)

    # 载入 slide1 音频关系
    audio_rels = {}
    try:
        rels_xml = z.read('ppt/slides/_rels/slide1.xml.rels')
        rels_root = ET.fromstring(rels_xml)
        for rel in rels_root.findall(f'{{{REL_PKG_NS}}}Relationship'):
            if rel.get('Type', '').endswith('/audio'):
                rid = rel.get('Id')
                target = rel.get('Target', '')
                norm = 'ppt/media/' + target.split('/')[-1] if target.startswith('../') else target
                audio_rels[rid] = (target, norm, rel.get('TargetMode', 'Internal'))
    except KeyError:
        pass

    # a) 是否存在有效嵌入音效
    bound_audio_rid = None
    bound_audio_path = None
    for rid, _ in sound_embeds:
        if rid in audio_rels:
            _t, norm_path, mode = audio_rels[rid]
            if mode != 'External' and norm_path in z.namelist():
                bound_audio_rid = rid
                bound_audio_path = norm_path
                break

    if bound_audio_rid is None:
        s5b_ok = False
        if sound_embeds:
            s5b_reasons.append('第8条动画组绑定的音效为外链或包内缺失，换电脑会丢失媒体')
        else:
            s5b_reasons.append('第8条动画组未绑定任何战争爆炸音效对象')
    else:
        s5b_reasons.append(f'第8条动画组已绑定嵌入音效：{bound_audio_path} ✓')

        # b) 时长 1—4 秒（仅 WAV 直接解析头信息；非 WAV 无法解析时不扣分，仅提示）
        duration_s = None
        try:
            data = z.read(bound_audio_path)
            if bound_audio_path.lower().endswith('.wav') and len(data) > 44:
                chunk_size = struct.unpack_from('<I', data, 4)[0]
                byte_rate = struct.unpack_from('<I', data, 28)[0]
                if byte_rate > 0:
                    duration_s = (chunk_size - 36) / byte_rate
        except Exception:
            duration_s = None
        if duration_s is None:
            s5b_reasons.append(f'音效时长未能从 {os.path.basename(bound_audio_path)} 头信息中解析（非 WAV 或头异常），跳过时长判定')
        elif 1.0 <= duration_s <= 4.0:
            s5b_reasons.append(f'音效时长 {duration_s:.2f}s，在 1—4s ✓')
        else:
            s5b_ok = False
            s5b_reasons.append(f'音效时长 {duration_s:.2f}s，不在 1—4s 范围')

        # c) 不循环播放：<p:audio>/<p:cMediaNode>/<p:cTn repeatCount=...> 与 sndTgt@loop
        loop_found = False
        loop_evidence = ''
        for audio_node in audio_media_nodes:
            for cTn in audio_node.findall('.//p:cTn', NS):
                rc = cTn.get('repeatCount')
                if rc == 'indefinite':
                    loop_found = True; loop_evidence = 'repeatCount=indefinite'; break
                try:
                    # OOXML repeatCount 单位 1000 = 1 次；>1000 视为循环
                    if rc is not None and int(rc) > 1000:
                        loop_found = True; loop_evidence = f'repeatCount={rc}'; break
                except ValueError:
                    pass
            if loop_found:
                break
        for _rid, snd_elem in sound_embeds:
            if snd_elem.get('loop') == '1':
                loop_found = True
                loop_evidence = loop_evidence or 'sndTgt@loop=1'
                break
        if loop_found:
            s5b_ok = False
            s5b_reasons.append(f'音效被设置为循环播放（{loop_evidence}）')
        else:
            s5b_reasons.append('音效未设置循环播放 ✓')

        # d) 音频图标位置与遮挡：在 slide1 找到引用同一 rId 的 <p:pic>
        icon_pic = None
        for pic in root1.findall('.//p:pic', NS):
            for af in pic.iter():
                if af.tag.split('}')[-1] == 'audioFile':
                    link_rid = af.get(REL_NS_ID + 'link') or af.get(REL_NS_ID + 'embed')
                    if link_rid == bound_audio_rid:
                        icon_pic = pic
                        break
            if icon_pic is not None:
                break

        if icon_pic is None:
            s5b_reasons.append('未在页面上找到音频图标（等同隐藏）✓')
        else:
            cnv = icon_pic.find('.//p:nvPicPr/p:cNvPr', NS)
            hidden = (cnv is not None and cnv.get('hidden') == '1')
            xfrm = icon_pic.find('.//p:spPr/a:xfrm', NS)
            off_el = xfrm.find('a:off', NS) if xfrm is not None else None
            ext_el = xfrm.find('a:ext', NS) if xfrm is not None else None
            x = int(off_el.get('x', 0)) if off_el is not None else 0
            y = int(off_el.get('y', 0)) if off_el is not None else 0
            cx = int(ext_el.get('cx', 0)) if ext_el is not None else 0
            cy = int(ext_el.get('cy', 0)) if ext_el is not None else 0
            slide_w, slide_h = 12192000, 6858000  # 16:9 幻灯片默认尺寸（EMU）
            edge_margin = 914400  # 1 英寸 ≈ 2.54cm，作为"页面边缘"阈值
            at_edge = (
                x <= edge_margin or y <= edge_margin or
                x + cx >= slide_w - edge_margin or y + cy >= slide_h - edge_margin
            )

            # 与其他 sp 的重叠：占图标面积 ≥25% 视为实质遮挡
            overlap_names = []
            icon_area = max(cx * cy, 1)
            for sp in root1.findall('.//p:sp', NS):
                sxfrm = sp.find('.//p:spPr/a:xfrm', NS)
                if sxfrm is None:
                    continue
                so = sxfrm.find('a:off', NS)
                se = sxfrm.find('a:ext', NS)
                if so is None or se is None:
                    continue
                sx, sy = int(so.get('x', 0)), int(so.get('y', 0))
                scx, scy = int(se.get('cx', 0)), int(se.get('cy', 0))
                if x + cx <= sx or sx + scx <= x or y + cy <= sy or sy + scy <= y:
                    continue
                ox = max(0, min(x+cx, sx+scx) - max(x, sx))
                oy = max(0, min(y+cy, sy+scy) - max(y, sy))
                if ox * oy / icon_area >= 0.25:
                    name_el = sp.find('.//p:nvSpPr/p:cNvPr', NS)
                    overlap_names.append(name_el.get('name', '') if name_el is not None else '')

            if hidden:
                s5b_reasons.append('音频图标已设为隐藏 ✓')
            elif at_edge and not overlap_names:
                s5b_reasons.append('音频图标位于页面边缘且未遮挡其他内容 ✓')
            else:
                s5b_ok = False
                if not hidden and not at_edge:
                    s5b_reasons.append('音频图标既未隐藏也不在页面边缘')
                if overlap_names:
                    s5b_reasons.append(f'音频图标与其他形状实质重叠：{overlap_names}')

    scores.append((5, '+5 第1页最后案例音效', s5b_ok, s5b_reasons))

    return scores


def _find_target_file(dir_path):
    """在给定目录中定位被评估的 .pptx 文档，返回绝对路径。"""
    if not os.path.isdir(dir_path):
        return None
    candidates = [
        name for name in os.listdir(dir_path)
        if name.lower().endswith(DOC_EXTS) and not name.startswith('~$')
    ]
    if not candidates:
        return None
    # 若目录内存在多个候选文件，取字母序第一个，保证确定性
    candidates.sort()
    return os.path.join(dir_path, candidates[0])


def _build_dim2_items(score_items):
    """把 dim2_check 返回的 (delta, label, hit, reasons) 转成统一 dim2_items。

    返回 (items, max_score, total_score)：
      - items 是对外的 dim2_items 列表
      - max_score 是所有加分项（正向 delta）的满分之和
      - total_score 是实际得分之和（含扣分项）
    """
    items = []
    max_score: int = 0
    total_score: int = 0
    for delta, label, hit, _reasons in score_items:
        delta_i: int = int(delta)
        rule = label.lstrip('+-0123456789').strip()
        max_delta = delta_i
        actual = max_delta if hit else 0
        if delta_i > 0:
            max_score += max_delta
        total_score += actual
        items.append({
            'rule': rule,
            'max_delta': max_delta,
            'delta': actual,
            'hit': bool(hit),
            'detail': '',
        })
    return items, max_score, total_score


def evaluate(dir_path: str) -> dict:
    """统一入口。

    参数 ``dir_path`` 为脚本所在目录路径，脚本自行在该目录中定位并打开
    被评估的 PPT 文档。返回值结构见《脚本接口差异与统一建议》§2.2。
    """
    result = {
        'id': SCRIPT_ID,
        'file_name': '',
        'status': 'ok',
        'error': None,
        'dim1_pass': False,
        'dim1_reason': '',
        'dim2_items': [],
        'total_score': 0,
        'max_score': 0,
    }

    try:
        target = _find_target_file(dir_path)
        if target is None:
            result['status'] = 'error'
            result['error'] = f'目录 {dir_path} 下未找到 .pptx 文件'
            return result
        result['file_name'] = os.path.basename(target)

        z = zipfile.ZipFile(target)

        try:
            d1_issues = dim1_check(result['file_name'])
            if d1_issues:
                result['dim1_pass'] = False
                result['dim1_reason'] = '；'.join(d1_issues)
                return result

            result['dim1_pass'] = True
            score_items = dim2_check(z)
        finally:
            z.close()

        dim2_items, max_score, total_score = _build_dim2_items(score_items)
        result['dim2_items'] = dim2_items
        # max_score 仅统计加分项（正向 delta）的满分之和；扣分项不计入满分基准
        result['max_score'] = max_score
        result['total_score'] = total_score
        return result
    except Exception as e:
        result['status'] = 'error'
        result['error'] = f'{type(e).__name__}: {e}'
        return result


if __name__ == '__main__':
    _dir = sys.argv[1] if len(sys.argv) > 1 else os.path.dirname(os.path.abspath(__file__))
    print(json.dumps(evaluate(_dir), ensure_ascii=False, indent=2))
