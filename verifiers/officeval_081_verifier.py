#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# pyright: reportUnknownMemberType=false, reportUnknownVariableType=false, reportUnknownArgumentType=false, reportUnknownParameterType=false, reportMissingParameterType=false, reportPrivateUsage=false, reportUnknownLambdaType=false, reportGeneralTypeIssues=false

import os
import re
import sys
import json
import zipfile
import openpyxl

# VBA 宏源码抽取统一走本文件内置的 OLE 复合文档解析 + MS-OVBA 解压实现。
# 不引入 oletools —— 其强制依赖 pcodedmp 为 GPL-3.0-or-later，与本项目
# Apache-2.0 许可不兼容；内置实现已能在零第三方依赖下读取 .xlsm 宏源码。
VBA_Parser = None

# 说明：所有接口统一约定，本模块不修改全局 sys.stdout，也不 print 主结果。
# 主结果统一由 evaluate(dir_path) -> dict 返回，入参为"脚本所在目录的路径"，
# 由脚本自己在该目录内定位并打开被评估的 Excel（.xlsx/.xlsm）。本文件仅在 __main__ 下作为
# 本地调试入口，把结果 dict 序列化为 JSON 打印。

SCRIPT_ID = "081"
# 维度二各评分项 max_delta 之和，用于维度一未通过 / 脚本异常时兜底
MAX_SCORE_TOTAL = 14

# 维度二各评分项按"结果追加顺序"的 max_delta 列表，用于把末尾 results 里的
# "+N：..." 字符串反解为结构化 dim2_items 时取值。顺序须与 _evaluate_impl 中
# results.append(...) 的顺序严格一致（共 4 项，累加 = MAX_SCORE_TOTAL）。
_MAX_DELTAS_ORDER = [
    5,  # 01 自动跳转功能真实生效
    5,  # 02 三列蛇形顺序（奇偶方向都正确）
    3,  # 03 行末换行正确
    1,  # 04 录入结构为三列
]
assert sum(_MAX_DELTAS_ORDER) == MAX_SCORE_TOTAL

# 维度二评分项文案（不含 "+N：" 前缀），与 _MAX_DELTAS_ORDER 一一对应
_RULE_TEXTS = [
    "自动跳转功能真实生效：输入数字并按回车后，光标自动跳到下一格（宏可实际运行，而非仅以文字/截图形式给出代码）",
    "跳转按三列蛇形顺序：奇数行左→右、偶数行右→左，两个方向都正确",
    "行末换行正确：一行走到尽头跳到下一行同侧首格，无错位",
    "录入结构为三列",
]
assert len(_RULE_TEXTS) == len(_MAX_DELTAS_ORDER)

def _rd(buf, off, n):
    """小端整数读取。"""
    return int.from_bytes(buf[off:off + n], "little")


def _col_num(letters):
    """把列字母（A、C、AA...）转成列号；非法输入返回 0。"""
    n = 0
    for ch in letters.upper():
        if not ("A" <= ch <= "Z"):
            return 0
        n = n * 26 + (ord(ch) - 64)
    return n


# ---------------------------------------------------------------------------
# 自包含 OLE 复合文档解析 + MS-OVBA 解压（olevba 不可用时的回退路径）
# ---------------------------------------------------------------------------
def _ole_extract_streams(data):
    """解析 OLE 复合文档字节，返回 [(stream_name, stream_bytes), ...]。"""
    if len(data) < 512 or data[:8] != b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1":
        return []
    ssz = 1 << _rd(data, 30, 2)          # 扇区大小
    mssz = 1 << _rd(data, 32, 2)         # 迷你扇区大小
    first_dir = _rd(data, 48, 4)
    mini_cutoff = _rd(data, 56, 4)
    first_minifat = _rd(data, 60, 4)
    first_difat = _rd(data, 68, 4)
    n_difat = _rd(data, 72, 4)
    ENDOFCHAIN, FREESECT = 0xFFFFFFFE, 0xFFFFFFFF

    def sect_bytes(sid):
        o = (sid + 1) * ssz
        return data[o:o + ssz]

    # 文件实际拥有的扇区数：损坏文件的头部字段可能是任意值，
    # 所有扇区号都必须落在该范围内，否则视为无效并停止跟链，
    # 避免 DIFAT 链自环或 n_difat 被写成超大值时陷入近似死循环。
    max_sect = max(0, (len(data) - 512) // ssz)

    # DIFAT：头部 109 项 + 后续 DIFAT 扇区链
    difat = []
    for k in range(109):
        v = _rd(data, 76 + k * 4, 4)
        if v not in (FREESECT, ENDOFCHAIN) and v < max_sect:
            difat.append(v)
    ds, cnt, seen_difat = first_difat, 0, set()
    while (ds not in (ENDOFCHAIN, FREESECT) and cnt <= n_difat + 1
           and ds < max_sect and ds not in seen_difat):
        seen_difat.add(ds)
        sb = sect_bytes(ds)
        for k in range((ssz // 4) - 1):
            v = _rd(sb, k * 4, 4)
            if v not in (FREESECT, ENDOFCHAIN) and v < max_sect:
                difat.append(v)
        ds = _rd(sb, ssz - 4, 4)
        cnt += 1
    # FAT
    fat = []
    for fs in difat:
        sb = sect_bytes(fs)
        for k in range(ssz // 4):
            fat.append(_rd(sb, k * 4, 4))

    def chain(sid):
        out, seen = [], set()
        while sid not in (ENDOFCHAIN, FREESECT) and sid < len(fat) and sid not in seen:
            seen.add(sid)
            out.append(sid)
            sid = fat[sid]
        return out

    def read_fat(start, size):
        b = bytearray()
        for sid in chain(start):
            b += sect_bytes(sid)
        return bytes(b[:size]) if size else bytes(b)

    # 目录流
    dir_data = read_fat(first_dir, 0)
    entries = []
    for i in range(0, len(dir_data) - 127, 128):
        e = dir_data[i:i + 128]
        nlen = _rd(e, 64, 2)
        etype = e[66]
        if etype not in (1, 2, 5):
            continue
        name = e[0:max(0, nlen - 2)].decode("utf-16-le", "replace") if nlen >= 2 else ""
        entries.append((name, etype, _rd(e, 116, 4), _rd(e, 120, 8)))

    root = next((x for x in entries if x[1] == 5), None)
    mini_stream = read_fat(root[2], root[3]) if root else b""

    # miniFAT
    minifat = []
    for sid in chain(first_minifat):
        sb = sect_bytes(sid)
        for k in range(ssz // 4):
            minifat.append(_rd(sb, k * 4, 4))

    def read_mini(start, size):
        b, sid, seen = bytearray(), start, set()
        while sid not in (ENDOFCHAIN, FREESECT) and sid < len(minifat) and sid not in seen:
            seen.add(sid)
            o = sid * mssz
            b += mini_stream[o:o + mssz]
            sid = minifat[sid]
        return bytes(b[:size])

    streams = []
    for name, etype, start, size in entries:
        if etype != 2:
            continue
        streams.append((name, read_mini(start, size) if size < mini_cutoff else read_fat(start, size)))
    return streams
def _ovba_decompress(comp, start):
    """MS-OVBA CompressedContainer 解压（comp[start] 必须为 0x01 签名）。"""
    if start >= len(comp) or comp[start] != 0x01:
        return None
    out = bytearray()
    i, n = start + 1, len(comp)
    while i + 1 < n:
        header = comp[i] | (comp[i + 1] << 8)
        i += 2
        size = (header & 0x0FFF) + 1
        compressed = (header & 0x8000) != 0
        if ((header >> 12) & 0x07) != 0b011:      # 非法块头 → 停止
            break
        chunk_end = min(i + size, n)
        if not compressed:
            out += comp[i:i + 4096]
            i += 4096
            continue
        dstart = len(out)
        while i < chunk_end:
            flags = comp[i]
            i += 1
            for bit in range(8):
                if i >= chunk_end:
                    break
                if not ((flags >> bit) & 1):
                    out.append(comp[i])
                    i += 1
                else:
                    if i + 1 >= n:
                        return bytes(out)
                    token = comp[i] | (comp[i + 1] << 8)
                    i += 2
                    diff = len(out) - dstart
                    if diff <= 0:
                        return bytes(out)
                    bc = min(12, max(4, (diff - 1).bit_length()))
                    lmask = 0xFFFF >> bc
                    length = (token & lmask) + 3
                    offset = ((token & (~lmask & 0xFFFF)) >> (16 - bc)) + 1
                    src = len(out) - offset
                    if src < 0:
                        return bytes(out)
                    for _ in range(length):
                        out.append(out[src])
                        src += 1
    return bytes(out)


def _extract_module_from_stream(sbytes):
    """在单个流里扫描出 VBA 模块源码（含 'Attribute VB_Name' 才认定有效）。"""
    n, idx, tries = len(sbytes), 0, 0
    while tries < 4000:
        p = sbytes.find(b"\x01", idx)
        if p < 0 or p + 3 > n:
            return None
        header = sbytes[p + 1] | (sbytes[p + 2] << 8)
        if ((header >> 12) & 0x07) == 0b011:
            tries += 1
            try:
                dec = _ovba_decompress(sbytes, p)
            except Exception:
                dec = None
            if dec and b"Attribute VB_Name" in dec:
                return dec
        idx = p + 1
    return None
# ---------------------------------------------------------------------------
# Excel 打开性 / 结构辅助
# ---------------------------------------------------------------------------
def _zip_names(path):
    try:
        with zipfile.ZipFile(path) as z:
            return set(z.namelist())
    except Exception:
        return set()


# 维度一门禁接受的交付物扩展名：仅 .xlsx / .xlsm
_ACCEPTED_EXTS = (".xlsx", ".xlsm")


def _is_openable_excel(path):
    """维度一门禁：是否为能正常打开的 Excel 文件（仅接受 .xlsx/.xlsm）。

    返回 (bool, reason)：reason 仅在未通过时给出失败原因摘要，通过时为空字符串。
    """
    ext = os.path.splitext(path)[1].lower()
    try:
        if ext in _ACCEPTED_EXTS:
            if not zipfile.is_zipfile(path):
                return (False, "文件不是有效的 Excel（OOXML/zip）结构，无法作为 Excel 打开")
            if "xl/workbook.xml" not in _zip_names(path):
                return (False, "缺少 xl/workbook.xml，非有效 Excel 工作簿")
            return (True, "")
        return (False, "交付物不是受支持的 Excel 文件类型（仅接受 .xlsx/.xlsm）")
    except Exception as exc:
        return (False, f"打开 Excel 失败：{type(exc).__name__}: {exc}")


def _read_vba_bin(path):
    try:
        with zipfile.ZipFile(path) as z:
            if "xl/vbaProject.bin" in z.namelist():
                return z.read("xl/vbaProject.bin")
    except Exception:
        pass
    return b""


def _extract_vba_code(path):
    """抽取 .xlsm 中的全部 VBA 源码：优先 olevba，失败回退自包含解析。"""
    parts = []
    if VBA_Parser is not None:
        try:
            vp = VBA_Parser(path)
            try:
                for _f, _s, _vfn, code in vp.extract_macros():
                    if code:
                        parts.append(code)
            finally:
                try:
                    vp.close()
                except Exception:
                    pass
        except Exception:
            parts = []
    if parts:
        return "\n".join(parts)
    binb = _read_vba_bin(path)
    if not binb:
        return ""
    try:
        streams = _ole_extract_streams(binb)
    except Exception:
        return ""
    for _name, sb in streams:
        try:
            mod = _extract_module_from_stream(sb)
        except Exception:
            mod = None
        if mod:
            parts.append(mod.decode("latin-1", "replace"))
    return "\n".join(parts)
def _sheet_has_3col(path):
    """回退结构判断：某工作表存在连续 3 列、且至少 3 行同时非空的数据块。"""
    try:
        wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    except Exception:
        return False
    try:
        for ws in wb.worksheets:
            rows = []
            for i, row in enumerate(ws.iter_rows(values_only=True)):
                rows.append(row)
                if i >= 199:
                    break
            maxc = max((len(r) for r in rows), default=0)
            if maxc < 3:
                continue
            for c in range(0, maxc - 2):
                cnt = 0
                for r in rows:
                    v0 = r[c] if c < len(r) else None
                    v1 = r[c + 1] if c + 1 < len(r) else None
                    v2 = r[c + 2] if c + 2 < len(r) else None
                    if v0 not in (None, "") and v1 not in (None, "") and v2 not in (None, ""):
                        cnt += 1
                if cnt >= 3:
                    return True
    finally:
        try:
            wb.close()
        except Exception:
            pass
    return False


# ---------------------------------------------------------------------------
# 维度二四项判定（对 VBA 源码做静态分析 + 结构兜底）
# ---------------------------------------------------------------------------
def _check_functional(code, has_vba_bin):
    """+5 自动跳转功能真实生效：存在真正可运行的 Worksheet_Change 宏并移动光标。"""
    if not has_vba_bin or not code:
        return False          # .xlsx 里"以文字给出的代码"没有 vbaProject.bin，判否
    if not re.search(r"Worksheet_Change", code, re.I):
        return False
    return bool(
        re.search(r"\.\s*Select\b", code, re.I)
        or re.search(r"Application\s*\.\s*Goto", code, re.I)
        or re.search(r"\.\s*Activate\b", code, re.I)
    )


def _check_serpentine(code):
    """+5 三列蛇形顺序：奇偶行两个方向都正确（既有右移又有左移，或含反向公式）。"""
    if not code or not re.search(r"Mod\s*2", code, re.I):
        return False
    reversal = bool(re.search(r"Columns\.Count\s*-\s*\w+\s*\+\s*1", code, re.I)) or \
        bool(re.search(r"-\s*(pos|col\w*)\s*\+\s*1", code, re.I))
    plus = bool(re.search(r"(nxtC|nextcol|c)\s*=\s*\w+\s*\+\s*1", code, re.I))
    minus = bool(re.search(r"(nxtC|nextcol|c)\s*=\s*\w+\s*-\s*1", code, re.I))
    return reversal or (plus and minus)


def _check_rowfold(code):
    """+3 行末换行正确：走到列边界时跳到下一行（存在换行 + 边界条件）。"""
    if not code:
        return False
    row_next = bool(re.search(r"(nxtR|nextrow|r)\s*=\s*\w+\s*\+\s*1", code, re.I)) or \
        bool(re.search(r"Row\s*\+\s*1", code, re.I))
    boundary = bool(re.search(r"[<>=]\s*3\b", code)) or bool(re.search(r"[<>=]\s*1\b", code)) or \
        bool(re.search(r"Columns\.Count", code, re.I)) or bool(re.search(r"Int\s*\(", code, re.I))
    return row_next and boundary


def _check_threecol(code, file_path):
    """+1 录入结构为三列：VBA 引用了跨 3 列的区域，或工作表存在 3 列数据块。"""
    if code:
        for m in re.finditer(r"\$?([A-Za-z]{1,3})\$?\d*\s*:\s*\$?([A-Za-z]{1,3})\$?\d*", code):
            a, b = _col_num(m.group(1)), _col_num(m.group(2))
            if a and b and (b - a + 1) == 3:
                return True
        if re.search(r"Mod\s*2", code, re.I) and re.search(r"=\s*3\b", code) and re.search(r"=\s*1\b", code):
            return True
    return _sheet_has_3col(file_path)
# ---------------------------------------------------------------------------
# 主评估逻辑
# ---------------------------------------------------------------------------
def _evaluate_impl(file_path, result):
    results = []

    # 维度一：能否作为 Excel 正常打开（门禁）
    dim1_ok, dim1_reason = _is_openable_excel(file_path)
    result["dim1_pass"] = dim1_ok
    # dim1_reason 只在维度一未通过时给出失败原因摘要，通过时保持为空
    result["dim1_reason"] = "" if dim1_ok else dim1_reason

    if not dim1_ok:
        hits = [False, False, False, False]
    else:
        has_vba_bin = "xl/vbaProject.bin" in _zip_names(file_path)
        code = _extract_vba_code(file_path) if has_vba_bin else ""
        hits = [
            _check_functional(code, has_vba_bin),
            _check_serpentine(code),
            _check_rowfold(code),
            _check_threecol(code, file_path),
        ]

    # 按 _MAX_DELTAS_ORDER 顺序拼装 "+N：rule" 字符串
    for hit, pts, rule in zip(hits, _MAX_DELTAS_ORDER, _RULE_TEXTS):
        results.append(("+%d：" % pts if hit else "+0：") + rule)

    if len(results) != len(_MAX_DELTAS_ORDER):
        raise RuntimeError("维度二评分项数量与 _MAX_DELTAS_ORDER 不一致")

    # 反解为结构化 dim2_items
    dim2_items = []
    total = 0
    for raw, max_delta in zip(results, _MAX_DELTAS_ORDER):
        head, _, rule = raw.partition("：")
        try:
            delta = int(head.replace("+", "").strip())
        except ValueError:
            delta = 0
        hit = delta > 0
        dim2_items.append({
            "rule": rule,
            "max_delta": max_delta,
            "delta": max_delta if hit else 0,
            "hit": hit,
        })
        total += max_delta if hit else 0

    result["dim2_items"] = dim2_items
    result["total_score"] = total
    result["max_score"] = MAX_SCORE_TOTAL
    return result


def evaluate(dir_path):
    """入口：在 dir_path 中定位被评估的 Excel 文件（仅接受 .xlsx/.xlsm）并给出评分。"""
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
            raise FileNotFoundError("目录不存在：%s" % dir_path)
        names = [f for f in sorted(os.listdir(dir_path)) if not f.startswith("~$")]
        cands = [f for f in names if f.lower().endswith(_ACCEPTED_EXTS)]
        if not cands:
            raise FileNotFoundError("目录内未找到 .xlsx/.xlsm 文件：%s" % dir_path)
        file_path = os.path.join(dir_path, cands[0])
        result["file_name"] = os.path.basename(file_path)
        return _evaluate_impl(file_path, result)
    except Exception as exc:
        # status 为 error 时未能完成判定，dim1_pass 留空，dim1_reason 也为空
        result["status"] = "error"
        result["error"] = "%s: %s" % (type(exc).__name__, exc)
        result["dim1_pass"] = ""
        result["dim1_reason"] = ""
        result["dim2_items"] = []
        result["total_score"] = 0
        return result


if __name__ == "__main__":
    _target = sys.argv[1] if len(sys.argv) > 1 else os.path.dirname(os.path.abspath(__file__))
    _out = evaluate(_target)
    print(json.dumps(_out, ensure_ascii=False, indent=2))
