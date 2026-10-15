# -*- coding: utf-8 -*-
"""规则层：业务词汇、排版容忍度阈值、评分卡权重全部外置，企业可定制不改代码。
加载顺序：显式路径 > 环境变量 EXCEL_AGENT_RULES > 项目根 rules.json > 内置默认。
定制方法：python -m engine.rules --dump > rules.json 后按需改。
纪律：改规则不改代码后必跑 regression.py；规则文件与黄金快照要一起入库。"""
import copy
import json
import os
import re

RULES_DEFAULT = {
    # 业务词汇：不同企业用语不同（最常需要定制的部分）
    "vocab": {
        "summary_prefixes": ["合计", "总计", "小计", "汇总"],
        "summary_exact": ["累计"],
        "ratio_keywords": ["率", "占比", "人均", "平均", "均价", "单价"],
        "diagonal_symbols": ["/", "\\", "︿", "丿"],
        # 装饰串字符类（正则 class 体）：REPT 进度条/手画分隔线不算标签也不算数值
        "decor_class": r"|\-—_=·*\s\d%．.",
        "note_marker_chars": ["："],
        # 表级说明词：含这些词的备注是整表口径说明，不挂到具体行
        "table_note_markers": ["本表", "以上", "注：", "备注：", "数据截至", "来源"],
    },
    # 分割层排版容忍度
    "segment": {
        "annotation": {"max_cells": 2, "min_text_len": 6, "dense_neighbor": 3,
                       "header_rep_frac": 0.5, "text_row_min_cells": 5,
                       "row_note_min_len": 6},
        "bridge": {"max_gap_rows": 2, "min_overlap": 0.7},
        "attach": {"max_gap_rows": 2, "min_overlap": 0.5, "min_table_rows": 2},
        "classify": {
            "note_max_cells": 6, "note_text_ratio": 0.8,
            "note2_avg_len": 20, "note2_text_ratio": 0.7, "note2_max_cells": 80,
            "table_num_ratio": 0.25,
            "ledger_min_cells": 8, "ledger_first_row_cov": 0.6, "ledger_min_rows": 3,
        },
    },
    # 代码置信评分卡（与模型分独立，供双信号合议；权重可按企业校准）
    "score": {
        "base": 0.40, "table": 0.15,
        "border_closed_at": 0.8, "border_closed": 0.20,
        "border_partial_at": 0.5, "border_partial": 0.10,
        "num_rich_at": 0.5, "num_rich": 0.10,
        "neat_first_row_at": 0.9, "neat_first_row": 0.10,
        "note_text_at": 0.9, "note_text": 0.15,
        "merge_cover_at": 0.5, "merge_cover_pen": 0.15,
        "sparse_fill_at": 0.35, "sparse_fill_pen": 0.15,
    },
    # 表头层判定阈值
    "headers": {
        "text_ratio": 0.6, "max_header_rows": 6, "scan_window": 8,
        "title_max_cells": 3,
        "float_scan_rows": 5, "float_min_top_cov": 4, "float_cov_ratio": 0.6,
        "ledger_min_rows": 2,
        "key_ratio": 0.8, "key_int_max": 999999,
        "label_max_len": 30,
    },
    # 抽取层
    "extract": {"fast_values_min_mb": 1},  # ≥此大小走 calamine 快值路径（实测快约一个数量级）
    # 勾稽/问答容差
    "query": {
            "tie_rel_tol": 0.005, "tie_abs_tol": 0.05,
            "naive_rel_tol": 0.005, "naive_abs_tol": 0.5, "naive_min_rows": 2,
            "lookup_show_rows": 8, "analyze_rows": 25, "analyze_stats_cols": 12,
        },
        "serve": {"max_rows": 400, "max_cols": 48},
}

_RULES_CACHE: dict[str, dict] = {}


def _deep_merge(base: dict, override: dict) -> dict:
    out = copy.deepcopy(base)
    for k, v in (override or {}).items():
        if isinstance(v, dict) and isinstance(out.get(k), dict):
            out[k] = _deep_merge(out[k], v)
        else:
            out[k] = v
    return out


def _compile(rules: dict) -> dict:
    v = rules["vocab"]
    rules["_summary_prefix_re"] = re.compile(
        "^(" + "|".join(re.escape(x) for x in v["summary_prefixes"]) + ")")
    rules["_summary_exact_set"] = set(v["summary_exact"])
    rules["_ratio_re"] = re.compile("|".join(re.escape(x) for x in v["ratio_keywords"]))
    rules["_diagonal_re"] = re.compile(
        "[" + "".join(re.escape(x) for x in v["diagonal_symbols"]) + "]")
    rules["_decor_re"] = re.compile("[" + v["decor_class"] + "]")
    return rules


def _default_path() -> str | None:
    p = os.environ.get("EXCEL_AGENT_RULES")
    if p and os.path.exists(p):
        return p
    p = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "rules.json")
    return p if os.path.exists(p) else None


def _find_unknown(base: dict, override: dict, prefix: str = "") -> list:
    """收集 override 里 base 不存在的键路径——拼错的定制会静默失效，必须告警。"""
    unknown = []
    for k, v in (override or {}).items():
        if k not in base:
            unknown.append(f"{prefix}{k}")
        elif isinstance(v, dict) and isinstance(base.get(k), dict):
            unknown.extend(_find_unknown(base[k], v, prefix=f"{prefix}{k}."))
    return unknown


def get_rules(path: str | None = None) -> dict:
    """加载规则（带缓存）。显式 path 优先；否则环境变量/项目根 rules.json/内置默认。"""
    resolved = path or _default_path() or "<default>"
    if resolved in _RULES_CACHE:
        return _RULES_CACHE[resolved]
    if resolved == "<default>":
        rules = _compile(copy.deepcopy(RULES_DEFAULT))
    else:
        try:
            with open(resolved, encoding="utf-8") as f:
                override = json.load(f)
        except json.JSONDecodeError as e:
            raise SystemExit(f"规则文件 JSON 语法错误 {resolved}: 第{e.lineno}行 {e.msg}")
        unknown = _find_unknown(RULES_DEFAULT, override)
        if unknown:
            import sys as _sys
            print(f"⚠ 规则文件存在未知键（拼错即不生效，请核对）: {', '.join(unknown[:8])}",
                  file=_sys.stderr)
        rules = _compile(_deep_merge(RULES_DEFAULT, override))
    _RULES_CACHE[resolved] = rules
    return rules


def is_summary_label(s: str, rules: dict) -> bool:
    sv = s.strip()
    return bool(rules["_summary_prefix_re"].match(sv)) or sv in rules["_summary_exact_set"]


def is_ratio_column(name: str, rules: dict) -> bool:
    return bool(rules["_ratio_re"].search(name))


if __name__ == "__main__":
    import sys
    sys.stdout.reconfigure(encoding="utf-8")
    if "--dump" in sys.argv:
        print(json.dumps(RULES_DEFAULT, ensure_ascii=False, indent=2))
