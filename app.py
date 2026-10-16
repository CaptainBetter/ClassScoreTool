# -*- coding: utf-8 -*-
"""
班级"二课分"统计小工具（Streamlit 本地网页版）—— 批量版
=========================================================
流程：① 上传班级总名单(Excel)
      ② 反复添加活动：填活动名 + 粘贴微信名单 -> 点"添加到待统计列表"
        （活动会暂存在 st.session_state 里，中途改阈值/翻看名单都不会丢）
      ③ 一键生成汇总 Excel：
         依次对每个活动做正则清洗 + 姓名提取 + 模糊匹配(阈值可调)
         低分名字统一列出来人工核对
         一个人参加多个活动时，横向拆成多个格子（绝不拼进同一格）：
             活动1=观影活动，活动2=歌唱比赛，没参加的格子留空
      ④ 网页预览 -> 下载汇总 Excel（纯数据、无任何样式）

运行：streamlit run app.py
"""

from __future__ import annotations

import datetime as _dt
import inspect
import io
import re
from collections import defaultdict

import pandas as pd
import streamlit as st

# =============================================================================
# 0. 模糊匹配后端：优先 rapidfuzz，其次 thefuzz，最后退化为标准库 difflib
#    （三个都不装也能跑，只是速度慢一点、精度略低）
# =============================================================================
try:
    from rapidfuzz import fuzz as _rf_fuzz
    from rapidfuzz import process as _rf_process

    FUZZ_BACKEND = "rapidfuzz"
except Exception:  # pragma: no cover
    _rf_fuzz = _rf_process = None
    try:
        from thefuzz import fuzz as _tf_fuzz
        from thefuzz import process as _tf_process

        FUZZ_BACKEND = "thefuzz"
    except Exception:  # pragma: no cover
        _tf_fuzz = _tf_process = None
        FUZZ_BACKEND = "difflib"

import difflib  # noqa: E402  (作为最后兜底)


def rank_candidates(
    query: str, choices: list[str], limit: int = 3, scorer: str = "ratio"
) -> list[tuple[str, float]]:
    """返回 [(候选姓名, 相似度0~100), ...]，按相似度从高到低，最多 limit 个。

    scorer: "ratio" = 严格按整体相似度（推荐，不易误匹配）
            "wratio" = 宽松（能容忍错别字/漏字，但更容易误匹配）
    """
    if not choices:
        return []

    # --- rapidfuzz（推荐，速度快）---
    if _rf_process is not None:
        _scorer = _rf_fuzz.ratio if scorer == "ratio" else _rf_fuzz.WRatio
        raw = _rf_process.extract(query, choices, scorer=_scorer, limit=limit)
        out = []
        for item in raw:
            # rapidfuzz>=3 返回 (choice, score, index)，旧版返回 (choice, score)
            if len(item) >= 2:
                out.append((item[0], float(item[1])))
        return out

    # --- thefuzz（fuzzywuzzy 的维护版）---
    if _tf_process is not None:
        _scorer = _tf_fuzz.ratio if scorer == "ratio" else _tf_fuzz.WRatio
        raw = _tf_process.extract(query, choices, scorer=_scorer, limit=limit)
        return [(item[0], float(item[1])) for item in raw]

    # --- 标准库兜底（不支持 wratio，等价于 ratio）---
    scored = sorted(
        ((difflib.SequenceMatcher(None, query, c).ratio() * 100.0, c) for c in choices),
        key=lambda x: -x[0],
    )[:limit]
    return [(c, s) for s, c in scored]


# =============================================================================
# 1. 文本清洗 与 姓名提取
# =============================================================================
# 需要从粘贴文本里剔除的"废话"（按长度从长到短依次替换，避免短词先吃掉长词）
NOISE_WORDS = [
    "以上同学", "以下同学", "以上名单", "以下名单", "参与名单", "活动名单",
    "名单如下", "本次活动", "活动名称", "第二课堂", "参与人员", "同学们",
    "名单", "如下", "附件", "备注", "说明", "请注意", "注意",
    "同学", "各位", "大家", "本次", "参加", "参与", "人员", "报名",
    "签到", "打卡", "统计", "核对", "确认", "负责", "请假", "缺席",
    "迟到", "未到", "已到", "收到", "谢谢", "辛苦", "合计", "总计",
    "共计", "序号", "姓名", "学号", "班级", "学院", "专业", "年级",
    "时间", "地点", "加分", "二课分", "二课", "综测", "活动", "以及",
    "还有", "@所有人", "所有人",
    # 常见职务备注（多出现在括号里，如"张三（班长）"）
    "副班长", "班长", "团支书", "学习委员", "生活委员", "体育委员", "文艺委员",
    "委员", "学委", "组长", "队长", "部长", "干事", "负责人", "老师",
]
NOISE_WORDS = sorted(set(NOISE_WORDS), key=len, reverse=True)

# 中文并列连词：只在"需要拆开一长串汉字"时当作分隔符用
CONJ = "和与跟及暨"

# 这些字几乎只出现在句子里，不会出现在姓名中。
# 用来判断一段超长汉字到底是"连写的名字"还是"一句通知"，避免把
# “请于周三到体育馆集合”切成 请于/周三/到体/育馆/集合 这种假名字。
SENTENCE_MARKERS = set("的了请是在我你他她们这那哪什么就都也很把被给对从向之其该此每些位会能要于到去来谢谢")

# 段落分隔符：标点 + 换行/制表符。
# 注意：普通空格**不算**段落分隔符，它只在段落内部作为"分词"用，
# 这样"李 四"才有机会被还原成"李四"。
_SEP_CHARS = "、,，;；:：/\\|·•~～—–-_=+*#`^&%$@!！?？。.…“”\"'‘’()[]【】{}<>《》"
_SEG_SPLIT_RE = re.compile("[" + re.escape(_SEP_CHARS) + r"\t\n\r\f\v]+")

_CJK = r"\u4e00-\u9fff\u3400-\u4dbf"


def _is_cjk_only(token: str) -> bool:
    return bool(token) and re.fullmatch("[" + _CJK + r"]+", token) is not None


def clean_raw_text(raw: str) -> str:
    """去掉废话词、表情符号、无关符号，只留下汉字/字母/数字/分隔符/空格。"""
    text = str(raw)
    text = text.replace("\u3000", " ").replace("\xa0", " ")  # 全角空格 / 不换行空格
    for word in NOISE_WORDS:  # 剔除"以上同学""名单"之类的噪声
        if word in text:
            text = text.replace(word, " ")
    # 保留：汉字、英文、数字、常见分隔符、空白；其余（emoji、日文假名等）当空格
    text = re.sub(
        r"[^" + _CJK + r"A-Za-z0-9" + re.escape(_SEP_CHARS) + r"\s]",
        " ",
        text,
    )
    return text


def _plan_chunks(n: int, max_len: int) -> list[int]:
    """把总长度 n（> max_len）拆成若干片段长度。

    中文姓名以 2 字为主、3 字次之，所以默认按 2 字切，
    长度为奇数时把最后一段凑成 3 字。
    例：6 -> [2,2,2]；5 -> [2,3]；7 -> [2,2,3]；9 -> [2,2,2,3]
    """
    sizes = [2] * (n // 2)
    if n % 2:
        if sizes:
            sizes[-1] += 1  # 奇数长度：最后一段补成 3 字
        else:
            sizes = [n]
    if sizes and max(sizes) > max_len:  # 例如限制"最长 2 字"，退化成按上限切
        base = max(int(max_len), 1)
        sizes = [base] * (n // base)
        if n % base:
            sizes[-1] += n % base
    return sizes


def clean_name(value) -> str:
    """把名单里的一个姓名单元格洗干净：去掉所有空白/全角空格，空值统一成空串。"""
    if value is None:
        return ""
    text = re.sub(r"[\s\u3000]+", "", str(value))
    return "" if text.lower() in ("nan", "none") else text


def _split_long_run(run: str, max_len: int) -> tuple[list[str] | None, str]:
    """尝试把一长串连写的汉字切成名字。

    返回 (片段列表 或 None, 给用户看的提示)。返回 None 表示"这更像一句话，建议丢掉"。

    策略：
      1) 中间有并列连词（张三和李四）-> 按连词切；
      2) 出现句子专用字（请/的/了/在/是…）-> 判定为句子，丢掉；
      3) 否则按 2 字一组切开（中文姓名以 2 字为主）。
    """
    # 策略一：连词切分，切出来每一段都像名字才采用
    parts = [p for p in re.split("[" + CONJ + "]", run) if p]
    if len(parts) > 1 and all(2 <= len(p) <= max_len for p in parts):
        return parts, f"“{run}”中间没有分隔符，已按名字切开：{' / '.join(parts)}，请人工核对"

    # 策略二：像句子 -> 直接丢掉，不要制造假名字
    hit = sorted({ch for ch in run if ch in SENTENCE_MARKERS})
    if hit:
        return None, (f"“{run}”读起来像一句通知（含“{''.join(hit)}”），已跳过。"
                      f"如果这确实是连写的名字，请用顿号或空格隔开再粘贴。")

    # 策略三：按 2 字一组切分
    pieces, pos = [], 0
    for size in _plan_chunks(len(run), max_len):
        pieces.append(run[pos:pos + size])
        pos += size
    pieces = [p for p in pieces if p]
    return pieces, f"“{run}”中间没有分隔符，已自动切成 {' / '.join(pieces)}，请人工核对"


def extract_names(
    raw: str,
    min_len: int = 2,
    max_len: int = 4,
    space_mode: str = "auto",
    keep_latin: bool = False,
    split_long: bool = True,
) -> tuple[list[str], list[str]]:
    """从杂乱文本里提取姓名。

    space_mode: auto  = 智能（"李 四"合成"李四"，"张三 李四"拆成两个人）
                split = 空格一律当分隔符
                merge = 空格一律忽略（整段合成名字）
    返回 (姓名列表(去重保序), 警告信息列表)
    """
    text = clean_raw_text(raw)
    names: list[str] = []
    warnings: list[str] = []

    def emit(run: str) -> None:
        """提交一段长度 >= min_len 的汉字串（必要时自动切分）。"""
        if len(run) < min_len:
            return
        if len(run) <= max_len:
            names.append(run)
            return
        if not split_long:
            warnings.append(f"“{run}”超过 {max_len} 字且没有分隔符，已跳过（可勾选“自动切分超长汉字串”）")
            return
        pieces, message = _split_long_run(run, max_len)
        warnings.append(message)
        if pieces:
            names.extend(pieces)

    for segment in _SEG_SPLIT_RE.split(text):
        segment = segment.strip()
        if not segment:
            continue

        # ---- merge 模式：整段去掉空格当一个整体 ----
        if space_mode == "merge":
            compact = re.sub(r"\s+", "", segment)
            if keep_latin or _is_cjk_only(compact):
                emit(compact)
            continue

        tokens = [t for t in segment.split() if t]
        pending: list[str] = []  # 等待配对的单字（处理"李 四"这种被空格拆开的名字）

        def flush_pending() -> None:
            """把相邻单字两两配对成名字（李+四 -> 李四）。"""
            i = 0
            while i + 1 < len(pending):
                names.append(pending[i] + pending[i + 1])
                i += 2
            if i < len(pending):
                warnings.append(f"“{pending[i]}”只剩一个字，可能是错别字或被拆散的名字，已忽略")
            pending.clear()

        for token in tokens:
            if _is_cjk_only(token):
                lead = len(token) - len(token.lstrip(CONJ))
                core = token.strip(CONJ)
                trail = len(token) - len(token.rstrip(CONJ))
                if lead:
                    flush_pending()
                if core:
                    if len(core) == 1:
                        if min_len <= 1:
                            names.append(core)
                        elif space_mode == "split":
                            warnings.append(f"“{core}”只有一个字，已忽略")
                        else:
                            pending.append(core)
                    else:
                        flush_pending()
                        emit(core)
                if trail:
                    flush_pending()
            elif keep_latin and re.fullmatch(r"[A-Za-z][A-Za-z0-9.'\-]*", token):
                flush_pending()
                if len(token) >= 2:
                    names.append(token)

        flush_pending()

    # 去重并保持原顺序
    unique = list(dict.fromkeys(n for n in names if n))
    return unique, warnings


# =============================================================================
# 2. 模糊匹配
# =============================================================================
IGNORE_OPTION = "（忽略此名字）"


def match_names(
    roster_names: list[str],
    extracted: list[str],
    threshold: float,
    scorer: str = "ratio",
):
    """把提取出的名字匹配到班级名单（单个活动）。

    返回:
      matched_map : {名单行号: [(原始名字, 相似度, 匹配方式), ...]}
      details     : 明细行（用于导出"匹配明细"表）
      review      : 需要人工核对的名字 [{raw, candidates}]
    """
    # 名单里可能有重名 -> 一个名字对应多行
    name_to_rows: dict[str, list[int]] = defaultdict(list)
    for i, n in enumerate(roster_names):
        name_to_rows[n].append(i)

    matched_map: dict[int, list[tuple[str, float, str]]] = defaultdict(list)
    details: list[dict] = []
    review: list[dict] = []

    for raw in extracted:
        # ---- 精确匹配优先 ----
        if raw in name_to_rows:
            rows = name_to_rows[raw]
            how = "精确" if len(rows) == 1 else "精确(名单重名)"
            for idx in rows:
                matched_map[idx].append((raw, 100.0, how))
            details.append({"原始名字": raw, "匹配到的姓名": raw, "相似度": 100.0, "匹配方式": how,
                            "说明": "" if len(rows) == 1 else f"名单中有 {len(rows)} 个同名同学，均已计入"})
            continue

        # ---- 模糊匹配 ----
        candidates = rank_candidates(raw, roster_names, limit=3, scorer=scorer)
        if candidates and candidates[0][1] >= threshold:
            best_name, best_score = candidates[0]
            rows = name_to_rows[best_name]
            how = "模糊" if len(rows) == 1 else "模糊(名单重名)"
            for idx in rows:
                matched_map[idx].append((raw, best_score, how))
            details.append({"原始名字": raw, "匹配到的姓名": best_name,
                            "相似度": round(best_score, 1), "匹配方式": how,
                            "说明": f"由“{raw}”模糊匹配而来"})
        else:
            # 低于阈值 -> 交给人工核对
            review.append({"raw": raw, "candidates": candidates})
            details.append({"原始名字": raw, "匹配到的姓名": "", "匹配方式": "待核对",
                            "相似度": round(candidates[0][1], 1) if candidates else 0.0,
                            "说明": "相似度低于阈值，未自动匹配"})

    return matched_map, details, review


def match_activities(
    roster_names: list[str],
    activities: list[dict],
    threshold: float,
    extract_kwargs: dict,
    scorer: str = "ratio",
) -> dict:
    """★ 批量核心：依次对每个暂存活动做 提取 + 模糊匹配。

    参数
    ----
    activities : [{"name": 活动名称, "raw": 名单原文}, ...]（就是 session_state 里那份）
    extract_kwargs : 传给 extract_names 的参数（侧边栏那几个开关）

    返回
    ----
    {
      "extracted":  [[该活动提取到的名字], ...]        与 activities 一一对应
      "warnings":   [[该活动的提示信息], ...]          与 activities 一一对应
      "stats":      [{"活动名称", "提取到", "自动匹配", "待核对"}, ...]
      "person_activities": {名单行号: [活动名称, ...]}  按添加顺序，自动去重
      "details":    [匹配明细行, ...]（含"活动名称"列）
      "review":     [{"activity_index", "activity_name", "raw", "candidates"}, ...]
    }
    """
    extracted_all: list[list[str]] = []
    warnings_all: list[list[str]] = []
    stats: list[dict] = []
    # 用 dict 当"有序集合"：同一个活动名对同一个人只记一次，且保持添加顺序
    person_activities: dict[int, list[str]] = defaultdict(list)
    details: list[dict] = []
    review: list[dict] = []

    for act_index, act in enumerate(activities):
        act_name = act["name"]
        extracted, warns = extract_names(act["raw"], **extract_kwargs)
        extracted_all.append(extracted)
        warnings_all.append(warns)

        matched_map, act_details, act_review = match_names(
            roster_names, extracted, threshold, scorer=scorer
        )

        # 把"谁匹配上了"翻译成"这个活动名记到这几行上"
        for idx, hits in matched_map.items():
            if hits and act_name not in person_activities[idx]:
                person_activities[idx].append(act_name)

        for row in act_details:
            details.append({"活动名称": act_name, **row})
        for item in act_review:
            review.append({**item, "activity_index": act_index, "activity_name": act_name})

        stats.append({
            "活动名称": act_name,
            "提取到": len(extracted),
            "自动匹配": len(matched_map),
            "待核对": len(act_review),
        })

    return {
        "extracted": extracted_all,
        "warnings": warnings_all,
        "stats": stats,
        "person_activities": person_activities,
        "details": details,
        "review": review,
    }


# -----------------------------------------------------------------------------
# 汇总表的两种排法
#   · "按活动对齐"（默认）：第 i 列永远对应"第 i 个添加的活动"，
#       活动1 这一列的格子里要么是第 1 个活动的名字，要么是空。
#   · "按人依次填"：每行从 活动1 开始，把他参加的活动依次填进后面的格子。
# 两种排法都遵守同一件事：一个活动一个格子，没参加就空着，绝不合并单元格。
# -----------------------------------------------------------------------------
LAYOUT_ALIGNED = "aligned"
LAYOUT_PACKED = "packed"
LAYOUT_LABELS = {
    LAYOUT_ALIGNED: "按活动对齐（活动1 固定是第 1 个添加的活动）",
    LAYOUT_PACKED: "按人依次填（每行从 活动1 开始，依次写他参加的活动）",
}

# 上一次统计留在表里的活动列：旧版是"活动名称"（逗号拼接），新版是"活动1/活动2…"
LEGACY_ACTIVITY_COL = "活动名称"
_ACTIVITY_COL_RE = re.compile(r"^活动\s*(\d+)$")


def is_activity_column(col) -> bool:
    """判断某一列是不是"统计结果"那种活动列（旧版"活动名称"或新版"活动1/活动2…"）。"""
    text = str(col).strip()
    return text == LEGACY_ACTIVITY_COL or _ACTIVITY_COL_RE.match(text) is not None


def activity_column_names(count: int) -> list[str]:
    """活动列的表头：活动1、活动2、活动3……"""
    return [f"活动{i}" for i in range(1, count + 1)]


def cell_to_text(value) -> str:
    """把表格里的一个单元格转成字符串：空值(None/NaN) -> ""。

    这样导出的空格子里既没有 "nan"，也没有逗号、0 之类的填充物。
    """
    if value is None:
        return ""
    try:
        if pd.isna(value):
            return ""
    except (TypeError, ValueError):  # 传进来的是列表等特殊值，直接当普通值处理
        pass
    text = str(value).strip()
    return "" if text.lower() in {"nan", "none", "nat"} else text


def build_summary_df(
    roster_df: pd.DataFrame,
    person_activities: dict[int, list[str]],
    activity_order: list[str] | None = None,
    layout: str = LAYOUT_ALIGNED,
) -> pd.DataFrame:
    """一行一个学生 + 每个活动单独占一个格子。

    规则（就是需求本身）：
      · 基础信息列（姓名、学号、班级……）原样保留，列名和列顺序一个都不动；
      · 活动列动态追加在最后，表头是 活动1、活动2、活动3……；
      · 参加了活动 -> 格子里直接写活动名称；没参加 -> 空字符串；
      · 一个活动只占一个格子，绝不把多个活动拼进同一格，也绝不合并单元格。

    参数
    ----
    person_activities : {名单行号: [活动名称, ...]}
    activity_order    : 本次添加的活动名（按添加顺序），决定"活动N"是哪一列
    layout            : LAYOUT_ALIGNED（按活动对齐）/ LAYOUT_PACKED（按人依次填）
    """
    activity_order = [str(n).strip() for n in (activity_order or []) if str(n or "").strip()]

    # 只留基础信息列：旧版的"活动名称"和旧版导出的"活动N"都属于上一次的统计结果，
    # 本次按新格式重新生成，避免把逗号拼接的老格式又带出来。
    base = roster_df[[c for c in roster_df.columns if not is_activity_column(c)]].copy()

    # 每个学生参加了哪些活动：去重、丢掉空名字，并按"添加顺序"排好
    add_order = {name: i for i, name in enumerate(activity_order)}
    per_row: list[list[str]] = []
    for pos in range(len(base)):
        names = [n for n in dict.fromkeys(person_activities.get(pos, [])) if n]
        names.sort(key=lambda n: add_order.get(n, len(add_order)))
        per_row.append(names)

    if layout == LAYOUT_PACKED:
        # 每行从"活动1"开始依次填：参加几个活动就占几个格子，剩下的全空
        width = max((len(names) for names in per_row), default=0)
        columns = activity_column_names(width)
        data = {
            columns[i]: [names[i] if i < len(names) else "" for names in per_row]
            for i in range(width)
        }
    else:
        # 默认：第 i 列永远是第 i 个添加的活动，没参加就留空
        columns = activity_column_names(len(activity_order))
        row_sets = [set(names) for names in per_row]
        data = {}
        for i, act_name in enumerate(activity_order):
            data[columns[i]] = [
                act_name if act_name in row_sets[pos] else ""
                for pos in range(len(base))
            ]

    out = base
    for col_name in columns:
        # 空格子统一写空字符串，绝不落成 nan / 逗号 / 0
        out[col_name] = data[col_name]
    return out


def finalize_summary_df(
    edited_df: pd.DataFrame,
    activity_cols: list[str],
) -> tuple[pd.DataFrame, pd.DataFrame]:
    """把预览表格里（可能被手动改过）的内容整理成最终导出的表。

    返回 (整理后的完整表, 至少参加了一个活动的那些行)：
      · 活动格子里的 None / nan 统一成空字符串，绝不留 "nan"、逗号、0；
      · "参加了活动" = 至少有一个活动格子不是空的。
    """
    out = edited_df.copy()
    present = [c for c in activity_cols if c in out.columns]
    for col in present:
        out[col] = out[col].map(cell_to_text)

    if present:
        hit_rows = out[out[present].ne("").any(axis=1)]
    else:
        hit_rows = out.iloc[0:0]
    return out, hit_rows


# =============================================================================
# 3. Excel 导出（纯数据，不带任何样式）
# =============================================================================
_UNNAMED_RE = re.compile(r"^Unnamed: \d+$")


def _blank_unnamed_headers(df: pd.DataFrame) -> pd.DataFrame:
    """原表里没有表头的列，pandas 会读成 "Unnamed: 0" 这种名字。

    导出时把它还原成真正的空白表头，避免凭空多出 "Unnamed: 0" 这种字。
    """
    blank = {c: "" for c in df.columns if isinstance(c, str) and _UNNAMED_RE.match(c)}
    return df.rename(columns=blank) if blank else df


def dataframes_to_xlsx_bytes(sheets: dict[str, pd.DataFrame]) -> bytes:
    """把若干 DataFrame 写成一个多 sheet 的 xlsx。

    刻意不做任何格式处理：
      · 不加背景填充色（表头不会被刷成蓝色）、不加粗、不设字体颜色
      · 不加边框、不冻结首行、不加自动筛选、不改列宽、不设行高
      · index=False —— 绝不写出 pandas 的索引列，不会多出 0/1/2… 那一列
    导出的表格就是纯白底 + 和原表一致的列结构，只有单元格里的值。
    """
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        for sheet_name, df in sheets.items():
            _blank_unnamed_headers(df).to_excel(
                writer, sheet_name=sheet_name[:31], index=False
            )
    return buf.getvalue()


def safe_filename(text: str, fallback: str = "统计结果") -> str:
    """把活动名变成安全的文件名片段。"""
    name = re.sub(r'[\\/:*?"<>|\s]+', "_", str(text or "")).strip("_")
    return name[:40] or fallback


def stretch_kwargs(func) -> dict:
    """让表格占满整行，同时兼容不同版本的 Streamlit。

    新版用 width="stretch"，老版本用 use_container_width=True。
    （Streamlit 正在弃用 use_container_width，这里做一下自动适配，免得以后报错。）
    """
    try:
        params = inspect.signature(func).parameters
        if "width" in params:
            return {"width": "stretch"}
        if "use_container_width" in params:
            return {"use_container_width": True}
    except (TypeError, ValueError):
        pass
    return {}


# =============================================================================
# 4. 待统计活动的暂存 —— st.session_state 用法详解
# =============================================================================
# ▼▼▼ 这一节是"批量录入"能成立的关键，先把 st.session_state 讲清楚 ▼▼▼
#
# ① 为什么非用它不可？
#    Streamlit 的机制是：**只要动一下控件（点按钮、改滑块、上传文件），
#    整个脚本就会从头到尾重新执行一遍**（rerun）。
#    普通的局部变量（比如 activities = []）在每次重跑时都会被重新创建，
#    所以"上一次点按钮加进去的活动"在下一次重跑就消失得干干净净。
#    st.session_state 是一个**跟着当前浏览器会话走**的字典，
#    重跑多少次它都在，正好用来存这类需要跨重跑保留的数据。
#
# ② 它的生命周期（按需求只放在内存里）
#    一个浏览器标签页 = 一个会话 = 一份独立的 session_state。
#    · 同一次运行中，反复点击"添加活动"、切换阈值，数据都在；
#    · 按 F5 刷新页面 / 关掉标签页 -> session_state 会被清空。
#      （这是 Streamlit 的设计，不是 bug；真要永久保存就得写文件/数据库。）
#
# ③ 本项目用到的三个 key
#    "activities" -> [{"name": 活动名称, "raw": 名单原文}, ...]  已添加的活动
#    "generated"  -> bool   是否点过"一键生成汇总 Excel"
#    "flash"      -> (提示类型, 提示文字)   操作完成后显示一次的提示
#
# ④ 三种基本写法
#    读：  st.session_state.get("activities", [])   取不到就给默认值，最安全
#    写：  st.session_state["activities"] = [...]   直接赋值即可
#    删：  st.session_state.pop("flash", None)      删除某个 key
#    判断：if "activities" not in st.session_state: st.session_state["activities"] = []
#
# ⑤ 一个容易踩的坑
#    控件自己也有 key（例如 key="summary_editor"），控件的值同样存在
#    session_state 里。**某个控件一旦被实例化，就不能再改它对应的 key**
#    （会抛 StreamlitAPIException）。所以下面"添加活动"用的是 st.form +
#    clear_on_submit=True，让 Streamlit 自己把输入框清空，
#    而不是我们手动去改 text_input 的值。
# ▲▲▲
SS_ACTIVITIES = "activities"
SS_GENERATED = "generated"
SS_FLASH = "flash"


def get_activities() -> list[dict]:
    """读出暂存的活动列表；第一次访问时先初始化成空列表。"""
    if SS_ACTIVITIES not in st.session_state:
        st.session_state[SS_ACTIVITIES] = []
    return st.session_state[SS_ACTIVITIES]


def set_flash(kind: str, message: str) -> None:
    """记一条提示，等这次页面重跑时显示一次（一般紧接着 st.rerun() 用）。"""
    st.session_state[SS_FLASH] = (kind, message)


def show_flash() -> None:
    """显示并清除上一次操作留下的提示（显示过就删，避免一直挂在页面上）。"""
    item = st.session_state.pop(SS_FLASH, None)
    if not item:
        return
    kind, message = item
    {"success": st.success, "warning": st.warning,
     "error": st.error, "info": st.info}.get(kind, st.info)(message)


def add_activity(name: str, raw: str) -> str:
    """把 (活动名称, 名单原文) 存进待统计列表，返回给用户看的结果文字。

    同名活动不会重复添加，而是把这次粘的名单并进原来那一条里
    （方便把一份很长的名单分两次粘贴）。
    """
    activities = get_activities()
    name, raw = name.strip(), (raw or "").strip()
    for act in activities:
        if act["name"] == name:
            act["raw"] = act["raw"].rstrip() + "\n" + raw
            return f"活动《{name}》之前加过了，已把这次粘的名单并进那一条。"
    activities.append({"name": name, "raw": raw})
    return f"✅ 已添加活动《{name}》，当前共 {len(activities)} 个活动待统计。"


# =============================================================================
# 5. 网页界面
# =============================================================================
def main() -> None:
    st.set_page_config(page_title="班级二课分统计", page_icon="📊", layout="wide")
    st.title("📊 班级二课分统计小工具（批量版）")
    st.caption("一次上传总名单 → 逐个添加活动 → 一键生成汇总 Excel（每个活动单独占一列，不再用逗号拼接）")

    # ---------------------------- 侧边栏：参数 ----------------------------
    with st.sidebar:
        st.header("⚙️ 匹配参数")
        threshold = st.slider(
            "相似度阈值（低于它的名字要人工核对）", min_value=50, max_value=100, value=85, step=1,
            help="推荐 80~90。阈值越高越严格，越不容易误匹配，但需要人工核对的名字会变多。",
        )
        st.divider()
        st.subheader("姓名提取")
        min_len, max_len = st.columns(2)
        min_name_len = min_len.number_input("最短字数", 1, 4, 2)
        max_name_len = max_len.number_input("最长字数", 2, 6, 4)
        space_mode_label = st.radio(
            "空格怎么处理",
            ["智能识别（推荐）", "空格就是分隔符", "忽略空格合并"],
            index=0,
            help="智能识别：能把“李 四”还原成“李四”，也能把“张三 李四”拆成两个人。",
        )
        space_mode = {"智能识别（推荐）": "auto", "空格就是分隔符": "split", "忽略空格合并": "merge"}[space_mode_label]
        keep_latin = st.checkbox("保留英文/拼音名字", value=False)
        split_long = st.checkbox("自动切分超长汉字串", value=True,
                                 help="例如“张三李四王五”会自动切成 张三/李四/王五（会给出提示）。")

        st.divider()
        scorer_label = st.radio(
            "相似度算法",
            ["严格 ratio（推荐）", "宽松 WRatio（容忍错别字）"],
            index=0,
            help=(
                "严格：整体相似度，不易误匹配。\n\n"
                "宽松：能容忍漏字/错别字，但更容易把不相干的名字匹配上，务必检查“匹配明细”。"
            ),
        )
        scorer = "ratio" if scorer_label.startswith("严格") else "wratio"
        st.caption(
            "小提示：名字只有 2~3 个字时，错一个字通常只有 60~75 分，"
            "会被请去人工核对——这是刻意的保护，避免张冠李戴。"
        )
        st.caption(f"模糊匹配引擎：**{FUZZ_BACKEND}**")
        if FUZZ_BACKEND != "rapidfuzz":
            st.caption("建议 `pip install rapidfuzz` 以获得更快更准的匹配。")

    # 提取参数打包好，后面 add / generate 都用同一份
    extract_kwargs = {
        "min_len": int(min_name_len),
        "max_len": int(max_name_len),
        "space_mode": space_mode,
        "keep_latin": keep_latin,
        "split_long": split_long,
    }

    # 从 session_state 取暂存的活动（跨重跑保留，详见第 4 节注释）
    activities = get_activities()

    # ---------------------------- ① 上传总名单 ----------------------------
    st.subheader("① 上传班级总名单（只需上传一次）")
    roster_file = st.file_uploader(
        "支持 .xlsx / .xlsm（必须有“姓名”列）",
        type=["xlsx", "xlsm", "xls"],
        accept_multiple_files=False,
    )
    show_flash()

    if roster_file is None:
        st.info("👆 先上传班级总名单 Excel，然后再往下添加活动。")
        _show_format_help()
        return

    # ---------------------------- 读取名单 ----------------------------
    if roster_file.name.lower().endswith(".xls"):
        st.error("不支持旧版 .xls 格式，请在 Excel 里“另存为” .xlsx 后再上传。")
        return

    try:
        sheet_dict = pd.read_excel(
            io.BytesIO(roster_file.getvalue()), sheet_name=None, engine="openpyxl", dtype=object
        )
    except Exception as exc:  # noqa: BLE001
        st.error(f"读取 Excel 失败：{exc}")
        return

    if not sheet_dict:
        st.error("这个 Excel 里没有工作表。")
        return

    c1, c2 = st.columns([1, 2])
    sheet_name = c1.selectbox("选择工作表", list(sheet_dict.keys()), index=0)
    roster_df = sheet_dict[sheet_name].dropna(how="all").reset_index(drop=True)

    if roster_df.empty:
        st.error(f"工作表《{sheet_name}》是空的。")
        return

    # 找姓名列：优先叫“姓名”的列，找不到就让用户自己选
    columns = [str(c) for c in roster_df.columns]
    default_col = "姓名" if "姓名" in columns else None
    for cand in ("名字", "学生姓名", "姓 名", "name", "Name"):
        if default_col is None and cand in columns:
            default_col = cand
    name_col = c2.selectbox(
        "哪一列是姓名？",
        columns,
        index=columns.index(default_col) if default_col in columns else 0,
    )
    if name_col != "姓名":
        st.warning(f"使用的姓名列是《{name_col}》。导出时基础信息列（姓名、学号……）原样保留，"
                   f"活动会作为新列 活动1、活动2…… 追加在最后。")

    # 总名单里如果还留着上一次统计的活动列，这里提醒一声：本次会按新格式重新生成
    legacy_cols = [c for c in roster_df.columns if is_activity_column(c)]
    if legacy_cols:
        st.info("检测到总名单里已经有活动列（" + "、".join(str(c) for c in legacy_cols) + "）。"
                "它们属于上一次的统计结果，本次导出会丢掉这些旧列、按新格式重新生成活动列，"
                "姓名/学号等基础信息不受影响。")

    # 清洗名单里的姓名（去空格、去全角空格、统一成字符串）
    # 注意：这里用 Python 自己的 re 逐元素处理，不要用 Series.str.replace，
    # 因为 pandas 3 + pyarrow 的字符串后端用 RE2，不认 \u 这种转义。
    roster_df[name_col] = roster_df[name_col].map(clean_name)
    roster_df = roster_df[roster_df[name_col] != ""].reset_index(drop=True)
    roster_names = roster_df[name_col].tolist()

    with st.expander(f"📋 查看班级总名单（共 {len(roster_names)} 人）", expanded=False):
        st.dataframe(roster_df, height=260, **stretch_kwargs(st.dataframe))

    # 重名提醒
    dup = pd.Series(roster_names).value_counts()
    dup = dup[dup > 1]
    if not dup.empty:
        st.warning("名单里存在重名同学：" + "、".join(f"{n}({c}人)" for n, c in dup.items())
                   + "。如果活动名单里出现这些名字，会被同时计入所有同名同学，请注意核对。")

    # ---------------------------- ② 添加活动 ----------------------------
    st.divider()
    st.subheader("② 添加活动（可以反复添加多个）")
    st.caption("填好活动名称 + 粘贴该活动的微信名单，点下面的按钮加入待统计列表；"
               "输入框会自动清空，接着粘下一个活动即可。")

    # 表单版本号：每次成功提交后 +1，下一轮渲染出来的是全新（空）的输入框。
    # 这样"添加上一个活动"的名单绝不会残留到下一个活动的输入框里。
    form_ver = int(st.session_state.get("form_ver", 0))

    with st.form(f"add_activity_form_{form_ver}", clear_on_submit=True):
        act_name = st.text_input(
            "活动名称",
            key=f"act_name_{form_ver}",
            placeholder="例如：青年大学习 / 运动会志愿者 / 班级团建",
        )
        act_raw = st.text_area(
            "该活动的参与名单（微信聊天记录直接粘贴）",
            key=f"act_raw_{form_ver}",
            height=150,
            placeholder="例：\n张三、李 四、王五\n以上同学请于周三下午到体育馆集合\n赵六，孙七；周八",
        )
        submitted = st.form_submit_button("➕ 添加到待统计列表", type="primary")

    if submitted:
        # 表单提交后：下面把 form_ver 加一，下一轮渲染出来的是全新的输入框（必定是空的），
        # 这样无论 clear_on_submit 有没有生效，都不会把上一个活动的名单粘到下一个活动上。
        st.session_state["form_ver"] = form_ver + 1
        name = (act_name or "").strip()
        raw = (act_raw or "").strip()
        if not name:
            set_flash("error", "请先填「活动名称」再添加。")
        elif not raw:
            set_flash("error", f"《{name}》还没有粘贴名单，请把参与名单粘进去再添加。")
        else:
            preview, warns = extract_names(raw, **extract_kwargs)
            if not preview:
                set_flash("error",
                          f"没能从《{name}》的文本里提取出任何名字。"
                          f"请检查文本内容，或在左侧侧边栏调整“最短/最长字数”。")
            else:
                message = add_activity(name, raw)
                message += f"（这条名单提取到 {len(preview)} 个名字）"
                if warns:
                    message += f"；另有 {len(warns)} 处提示，生成汇总时会列出来。"
                set_flash("success", message)
        st.rerun()

    # ---------------------------- ③ 待统计列表 ----------------------------
    if not activities:
        st.info("还没有添加任何活动。填好上面的活动名称和名单，点「➕ 添加到待统计列表」即可。")
        _show_format_help()
        return

    # 每次重跑都按当前参数重新算一遍（数据量很小，界面能实时跟着阈值变化）
    analysis = match_activities(roster_names, activities, threshold, extract_kwargs, scorer=scorer)

    st.divider()
    st.subheader(f"③ 待统计活动列表（已添加 {len(activities)} 个）")
    st.caption("核对一下活动名称和人数有没有问题；确认无误后点下面的「一键生成汇总 Excel」。")

    overview_df = pd.DataFrame(analysis["stats"])
    overview_df.insert(0, "序号", range(1, len(overview_df) + 1))
    st.dataframe(overview_df, height=min(120 + 36 * len(overview_df), 420),
                 hide_index=True, **stretch_kwargs(st.dataframe))

    with st.expander("🗑 删除某一个活动"):
        for i, act in enumerate(list(activities)):
            col_a, col_b = st.columns([5, 1])
            col_a.write(f"**{i + 1}. {act['name']}**")
            if col_b.button("删除", key=f"del_act_{i}"):
                activities.pop(i)
                set_flash("info", f"已删除活动《{act['name']}》。")
                st.rerun()

    btn_clear, btn_generate = st.columns([1, 1])
    if btn_clear.button("🧹 清空所有待统计活动", **stretch_kwargs(st.button)):
        activities.clear()                       # 直接清空 session_state 里那个列表
        st.session_state[SS_GENERATED] = False
        set_flash("info", "已清空全部待统计活动。")
        st.rerun()
    if btn_generate.button("🚀 一键生成汇总 Excel", type="primary", **stretch_kwargs(st.button)):
        # 只是把开关打开；下面这一段会在**同一次重跑**里立刻执行
        st.session_state[SS_GENERATED] = True

    if not st.session_state.get(SS_GENERATED):
        st.info("点「🚀 一键生成汇总 Excel」后，这里会显示匹配结果和下载按钮。")
        return

    # ---------------------------- ④ 人工核对 ----------------------------
    review = analysis["review"]
    # 人工核对的结果：{名单行号: [活动名称, ...]}
    manual_activities: dict[int, list[str]] = defaultdict(list)

    st.divider()
    if review:
        st.subheader(f"④ 需要人工核对的名字（{len(review)} 个）")
        st.caption(
            "这些名字与名单里任何人的相似度都低于阈值，默认**不计入**。"
            "下拉框里的候选已按相似度排序，最像的排在最前；确实不在名单里就保持“忽略”。"
        )
        auto_pick = st.checkbox(
            "偷懒模式：低分名字默认采用最接近的候选（选完后请检查下面的汇总表格）", value=False
        )
        roster_options = [IGNORE_OPTION] + sorted(roster_names)

        # 按活动分组显示，方便你知道这个错名字是哪个活动里的
        grouped: dict[int, list[dict]] = defaultdict(list)
        for item in review:
            grouped[item["activity_index"]].append(item)

        for act_index in sorted(grouped):
            st.markdown(f"**活动：{activities[act_index]['name']}**（{len(grouped[act_index])} 个待核对）")
            for j, item in enumerate(grouped[act_index]):
                raw = item["raw"]
                candidates = item["candidates"]
                hint = "、".join(f"{n} {s:.0f}%" for n, s in candidates) or "名单为空"
                cc1, cc2 = st.columns([1, 2])
                cc1.markdown(f"**{raw}**  \n<small>最接近：{hint}</small>", unsafe_allow_html=True)
                default_idx = 0
                if auto_pick and candidates and candidates[0][0] in roster_options:
                    default_idx = roster_options.index(candidates[0][0])
                # key 由"活动序号 + 组内序号 + 原始名字"组成，保证每次重跑都能对上
                choice = cc2.selectbox(
                    f"“{raw}”对应名单里的谁？",
                    roster_options,
                    index=default_idx,
                    key=f"rv_{act_index}_{j}_{raw}",
                    label_visibility="collapsed",
                )
                if choice != IGNORE_OPTION:
                    for idx in [k for k, n in enumerate(roster_names) if n == choice]:
                        if item["activity_name"] not in manual_activities[idx]:
                            manual_activities[idx].append(item["activity_name"])
                    for row in analysis["details"]:
                        if (row["活动名称"] == item["activity_name"]
                                and row["原始名字"] == raw and not row["匹配到的姓名"]):
                            row.update({"匹配到的姓名": choice, "相似度": 100.0,
                                        "匹配方式": "人工指定", "说明": "人工核对确认"})
        if manual_activities:
            st.info(f"已人工指定 {len(manual_activities)} 人。")
    else:
        st.success("✅ 所有活动的名字都自动匹配上了，无需人工核对。")

    # 自动匹配 + 人工指定 合并
    person_activities: dict[int, list[str]] = defaultdict(list)
    for idx, names in analysis["person_activities"].items():
        person_activities[idx].extend(names)
    for idx, names in manual_activities.items():
        for name in names:
            if name not in person_activities[idx]:
                person_activities[idx].append(name)

    # ---------------------------- ⑤ 汇总结果 ----------------------------
    st.divider()
    st.subheader("⑤ 汇总结果（活动格子可以直接在表格里改）")
    st.caption("一行一个学生；每个活动单独占一个格子，填活动名称，没参加的格子留空。"
               "一个活动绝不会和别的活动挤在同一格里，也不会合并单元格。")
    layout_label = st.radio(
        "活动列怎么排？",
        list(LAYOUT_LABELS.values()),
        index=0,
        horizontal=True,
        help=(
            "按活动对齐：第 1 列永远是第 1 个添加的活动，参加了就写活动名，没参加就空着"
            "（每一列都是同一个活动，方便按活动统计人数）。\n\n"
            "按人依次填：每行从 活动1 开始，依次写他参加的活动，参加几个就占几个格子。"
        ),
    )
    layout = LAYOUT_ALIGNED if layout_label == LAYOUT_LABELS[LAYOUT_ALIGNED] else LAYOUT_PACKED

    activity_order = [str(a["name"]) for a in activities]
    summary_df = build_summary_df(roster_df, person_activities, activity_order, layout=layout)
    activity_cols = [c for c in summary_df.columns if is_activity_column(c)]
    edited_df = st.data_editor(
        summary_df,
        height=380,
        num_rows="fixed",
        disabled=[c for c in summary_df.columns if c not in activity_cols],
        # key 里带上排法和列数：换了排法或增删了活动，表格会干净重来，不会残留上一次的编辑
        key=f"summary_editor_{layout}_{len(activity_cols)}",
        **stretch_kwargs(st.data_editor),
    )

    # 手动改过的内容照样算数；空格子统一成 ""（不留 nan、逗号、0）
    edited_df, hit_rows = finalize_summary_df(edited_df, activity_cols)

    k1, k2, k3 = st.columns(3)
    k1.metric("本次统计活动", f"{len(activities)} 个")
    k2.metric("汇总计入人数", f"{len(hit_rows)} 人")
    k3.metric("待人工核对", f"{len(review)} 个")

    details_df = pd.DataFrame(
        analysis["details"],
        columns=["活动名称", "原始名字", "匹配到的姓名", "相似度", "匹配方式", "说明"],
    )

    # ---------------------------- 下载 ----------------------------
    xlsx_bytes = dataframes_to_xlsx_bytes(
        {"统计结果": edited_df, "匹配明细": details_df}
    )
    stamp = _dt.datetime.now().strftime("%Y%m%d_%H%M")
    if len(activities) == 1:
        tag = safe_filename(activities[0]["name"], fallback="1个活动")
    else:
        tag = f"{len(activities)}个活动"
    file_name = f"二课分汇总_{tag}_{stamp}.xlsx"

    d1, d2 = st.columns([1, 2])
    d1.download_button(
        "⬇️ 下载汇总 Excel",
        data=xlsx_bytes,
        file_name=file_name,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
    )
    d2.caption(
        f"共 {len(activities)} 个活动汇总在一起，文件名：`{file_name}`。"
        "文件包含两个工作表：**统计结果**（完整名单 + 活动1/活动2…… 每个活动单独一格）和 "
        "**匹配明细**（每个活动里每个名字的匹配情况，便于复查）。"
    )

    # 提取过程中的提示（比如把一句通知整段跳过了）
    all_warnings = [(activities[i]["name"], w)
                    for i, warns in enumerate(analysis["warnings"]) for w in warns]
    if all_warnings:
        with st.expander(f"⚠️ 姓名提取提示（{len(all_warnings)} 条）"):
            for act_name, w in all_warnings:
                st.warning(f"【{act_name}】{w}")

    with st.expander("🔍 匹配明细"):
        st.dataframe(details_df, height=300, hide_index=True, **stretch_kwargs(st.dataframe))


def _show_format_help() -> None:
    with st.expander("📌 使用说明 / 班级总名单应该长什么样？"):
        st.markdown(
            """
**三步走：**

1. 上传班级总名单（`.xlsx`，必须有一列 **姓名**；列名不叫"姓名"也可以，上传后手动选一下）。
2. 每有一个活动，就填一次活动名称 + 粘贴该活动的名单，点「➕ 添加到待统计列表」。
   想统计几个活动就加几次，输入框会自动清空。
3. 核对无误后点「🚀 一键生成汇总 Excel」，检查预览表格，最后下载。

| 姓名 | 学号 | 备注 |
| --- | --- | --- |
| 张三 | 20230101 | |
| 李四 | 20230102 | |

- 其它列（学号、班级……）会原样保留，导出时在最后**横向追加活动列**：`活动1`、`活动2`、`活动3`……
- 一个活动只占一个格子，格子里直接写活动名称；**没参加的格子彻底留空**，
  绝不把多个活动拼进同一格（不再有"观影活动, 歌唱比赛"这种写法），也绝不合并单元格。
- 表格长这样（张三参加了两个活动，李四只参加了一个）：

  | 姓名 | 学号 | 活动1 | 活动2 |
  | --- | --- | --- | --- |
  | 张三 | 20230101 | 观影活动 | 歌唱比赛 |
  | 李四 | 20230102 | 观影活动 | |
- 如果总名单里**已经有"活动名称"列或"活动1/活动2…"列**（上次导出的结果），
  本次会丢掉这些旧活动列、按新格式重新生成，姓名/学号等基础信息不受影响。
- 添加的活动只存在浏览器内存里：**刷新页面(F5)会清空**，生成前请不要刷新。
            """
        )


if __name__ == "__main__":
    main()
