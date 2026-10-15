# -*- coding: utf-8 -*-
"""表头树层：对每个 table 区块构建 列树（多层主次表头）+ 行维度（左侧键列）。
纯几何规则；语义歧义只打标（同名/斜表头/汇总行），不猜业务含义。
判定阈值与业务词汇见 rules.json（vocab/headers 段）。"""
import re

from .extract import _has_value
from .rules import get_rules, is_summary_label


def _val(info):
    v = info["cv"] if info["cv"] is not None else info["v"]
    return v


def _is_text(info) -> bool:
    v = _val(info)
    return isinstance(v, str) and v.strip() != ""


def _is_num(info) -> bool:
    v = _val(info)
    return isinstance(v, (int, float)) and not isinstance(v, bool)


def build_header_tree(sheet, block, rules=None) -> dict:
    rules = rules or get_rules()
    rh = rules["headers"]
    cells = sheet["cells"]
    merged = sheet["merged"]
    r1, c1, r2, c2 = block["bbox"]

    def info(r, c):
        return cells.get((r, c))

    def value(r, c):
        i = info(r, c)
        return _val(i) if i else None

    # 合并区索引：格子 → 其合并区
    mindex = {}
    for mr1, mc1, mr2, mc2 in merged:
        for r in range(mr1, mr2 + 1):
            for c in range(mc1, mc2 + 1):
                mindex[(r, c)] = (mr1, mc1, mr2, mc2)

    rows = list(range(r1, r2 + 1))
    cols = list(range(c1, c2 + 1))

    # --- 行画像：每行 文本数/数值数/公式数 ---
    profile = {}
    for r in rows:
        t = n = f = 0
        for c in cols:
            i = info(r, c)
            if not i or not _has_value(_val(i)):
                continue
            if _is_text(i):
                t += 1
            elif _is_num(i):
                n += 1
            if i["f"]:
                f += 1
        profile[r] = {"t": t, "n": n, "f": f}

    # --- 表头行判定：从上往下，文本主导 或 含文本型跨列合并 的行视为表头；连续两个数据行则停止 ---
    # 先找浮动草稿区：表头上方列覆盖率显著偏低的稀疏行（漂着的计算/说明），不是表头
    cover = {r: len({c for c in cols if _has_value(value(r, c))}) for r in rows}
    max_top_cov = max((cover[r] for r in rows[:5]), default=0)
    floating_rows = []
    header_start_idx = 0
    if max_top_cov >= rh["float_min_top_cov"]:  # 小表没有浮动区可言
        for idx, r in enumerate(rows[:rh["float_scan_rows"]]):
            if cover[r] >= rh["float_cov_ratio"] * max_top_cov:
                header_start_idx = idx
                break
            floating_rows.append(r)
    body_rows = rows[header_start_idx:]

    # 文本型跨列合并才算分组表头；REPT("|"…)&"86%" 单元格内进度条之类的装饰串不算
    def looks_label(v) -> bool:
        if not isinstance(v, str):
            return False
        s = v.strip()
        core = rules["_decor_re"].sub("", s)
        return bool(core) and len(core) <= rh["label_max_len"]

    def col_merge_text_at(r):
        for m in merged:
            if m[0] == r and m[3] > m[1] and r1 <= m[0] and m[2] <= r2 \
                    and c1 <= m[1] and m[3] <= c2:
                if looks_label(value(m[0], m[1])):
                    return True
        return False

    # 纯文本块（字典/台账）：无任何数值与公式 → 首行即表头，其余为数据
    all_text = all(profile[r]["n"] == 0 and profile[r]["f"] == 0 for r in body_rows) \
        and any(profile[r]["t"] > 0 for r in body_rows)
    if all_text and len(body_rows) >= rh["ledger_min_rows"]:
        header_rows = [body_rows[0]]
        data_rows = body_rows[1:]
    else:
        header_rows = []
        data_rows = []
        for idx, r in enumerate(body_rows):
            p = profile[r]
            total = p["t"] + p["n"]
            has_col_merge = col_merge_text_at(r)
            is_headerish = total > 0 and (p["n"] == 0 or has_col_merge) and \
                           (p["t"] / max(total, 1) >= rh["text_ratio"] or has_col_merge)
            if is_headerish and len(header_rows) < rh["max_header_rows"] \
                    and idx < rh["scan_window"]:
                header_rows.append(r)
            else:
                data_rows.append(r)
                if len(header_rows) > 0 and idx >= len(header_rows):
                    nxt = body_rows[idx + 1] if idx + 1 < len(body_rows) else None
                    if nxt is not None:
                        break  # 表头结束后进入数据区，剩余行全部视为数据
        if not data_rows:
            data_rows = [r for r in body_rows if r not in header_rows] or body_rows
            header_rows = [r for r in body_rows if r not in data_rows]

    # 标题行：首层表头内容格 ≤N 且下方还有表头层 → 剥离为区块标题
    title_row = None
    if header_rows and len(header_rows) >= 2:
        first = header_rows[0]
        n_cells = len({c for c in cols if _has_value(value(first, c))})
        if n_cells <= rh["title_max_cells"]:
            title_row = header_rows.pop(0)
            title_text = " / ".join(str(value(title_row, c)).strip()
                                    for c in cols if _has_value(value(title_row, c)))
        else:
            title_text = None
    else:
        title_text = None

    header_depth = len(header_rows)
    data_zone = [r for r in rows
                 if r not in header_rows and r not in floating_rows and r != title_row]

    # --- 叶子列：数据区有内容的列 ---
    leaf_cols = [c for c in cols
                 if any(info(r, c) and _has_value(_val(info(r, c))) for r in data_zone)]

    # --- 列树：每个叶子列在每层表头行的标签（合并区取左上值） ---
    def label_at(r, c):
        v = value(r, c)
        if v is None and (r, c) in mindex:
            mr1, mc1, _, _ = mindex[(r, c)]
            v = value(mr1, mc1)
        return str(v).strip() if _has_value(v) else None

    col_tree = []
    for c in leaf_cols:
        path = []
        for hr in header_rows:
            lab = label_at(hr, c)
            if lab and (not path or path[-1] != lab):
                path.append(lab)
        col_tree.append({"col": c, "path": path, "label": path[-1] if path else None})

    # --- 左侧键列（行维度）：数据区前段连续的 键候选 列（文本/日期串/短序号整数） ---
    def key_eligible(i) -> bool:
        v = _val(i)
        if isinstance(v, str):
            return bool(v.strip())
        return isinstance(v, int) and not isinstance(v, bool) \
            and 0 <= v <= rh["key_int_max"]

    key_cols = []
    for c in leaf_cols:
        elig = sum(1 for r in data_zone if info(r, c) and key_eligible(info(r, c)))
        vals = sum(1 for r in data_zone if info(r, c) and _has_value(_val(info(r, c))))
        if vals and elig / vals >= rh["key_ratio"]:
            key_cols.append(c)
        else:
            break
    # 垂直合并 → 行分组树
    row_groups = []
    for c in key_cols:
        groups = []
        for r in data_zone:
            m = mindex.get((r, c))
            if m and m[2] > m[0]:
                groups.append({"rows": [m[0], m[2]], "label": str(value(m[0], m[1]) or "").strip()})
        if groups:
            row_groups.append({"col": c, "groups": groups})

    # --- 交叉表判定：多层表头 + 键列存在垂直合并分组（行维度）才是交叉表，
    #     普通台账（单层表头+多文本键列）不算 ---
    top_left = value(header_rows[0], cols[0]) if header_rows else None
    diagonal = False
    if top_left is not None and isinstance(top_left, str):
        diagonal = ("\n" in top_left) or bool(rules["_diagonal_re"].search(top_left[:8]))
    cross_tab = bool(key_cols) and header_depth >= 2 and bool(row_groups)

    # --- 保护标记：汇总行 / 同名列 / 单元格内换行备注 ---
    # 汇总行 = 键列标签(合计/总计/…) 或 SUM 区间【纵向跨行、位于当前行上方、落在本块内】的聚合行；
    # 行内横向合计列（=SUM(B3:M3) 单行区间）与跨表取数不算
    EXACT_SUM_PAT = re.compile(r"=SUM\(|=SUBTOTAL\(", re.I)
    sum_range_pat = re.compile(r"\$?([A-Z]{1,3})\$?(\d+):\$?[A-Z]{1,3}\$?(\d+)")

    def intra_block_col_sum(formula: str, cur_r: int) -> bool:
        # 只看最后一个跨表引用符之后的本地部分，避免 =SUM(其他表!B3:B9) 的行号撞进本块
        local = formula.rsplit("!", 1)[-1]
        for _, r_a, r_b in sum_range_pat.findall(local):
            ra, rb = int(r_a), int(r_b)
            if rb > ra and rb < cur_r and r1 <= ra and rb <= r2:
                return True
        return False

    summary_rows = []
    for r in data_zone:
        hit = False
        for c in key_cols or cols:
            v = value(r, c)
            if isinstance(v, str) and is_summary_label(v, rules):
                hit = True
                break
        if not hit:
            for c in cols:
                i = info(r, c)
                if i and isinstance(i["v"], str) and EXACT_SUM_PAT.search(i["v"]) \
                        and intra_block_col_sum(i["v"], r):
                    hit = True
                    break
        if hit:
            summary_rows.append(r)

    seen, dup_names = {}, []
    for node in col_tree:
        lab = node["label"]
        if lab:
            if lab in seen and lab not in dup_names:
                dup_names.append(lab)
            seen[lab] = True

    nl_notes = [{"pos": [r, c], "text": str(_val(i))[:60]}
                for (r, c), i in cells.items()
                if r1 <= r <= r2 and c1 <= c <= c2 and i["nl"]]

    return {
        "header_rows": header_rows,
        "header_depth": header_depth,
        "title_row": title_row,
        "title_text": title_text,
        "floating_rows": floating_rows,
        "data_rows": [data_zone[0], data_zone[-1]] if data_zone else [],
        "leaf_cols": leaf_cols,
        "col_tree": col_tree,
        "key_cols": key_cols,
        "row_groups": row_groups,
        "cross_tab": cross_tab,
        "diagonal_header": diagonal,
        "summary_rows": summary_rows,
        "dup_header_names": dup_names,
        "newline_notes": nl_notes[:20],
    }
