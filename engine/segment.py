# -*- coding: utf-8 -*-
"""分割层：多线索区块分割（纯代码规则，零模型）。
线索：1) 全空行/列切带  2) 带内空列左右分表  3) 合并区连续性（行列双向不断裂）
      4) 批注行剥离（稀疏文本行不当表格粘合剂）  5) 边框闭合加分
      6) 文本/数值密度分类  7) 合计行回贴（被空行切断的表尾汇总）
输出区块 + 类型 + 代码置信分（0~1，与模型分独立，供双信号合议）。"""
from .extract import _has_value
from .rules import get_rules, is_summary_label

NOTE, TABLE, MIXED = "note", "table", "mixed"
TYPE_LABEL = {TABLE: "数据表", NOTE: "文本说明", MIXED: "混合/待复核"}


def _is_num(v) -> bool:
    if isinstance(v, bool):
        return False
    return isinstance(v, (int, float))


def _val(info):
    return info["cv"] if info["cv"] is not None else info["v"]


def _bands(rows_with_content):
    """全空行切带：返回 [(r1,r2), ...]，行号含端点。"""
    if not rows_with_content:
        return []
    bands, start, prev = [], rows_with_content[0], rows_with_content[0]
    for r in rows_with_content[1:]:
        if r > prev + 1:
            bands.append((start, prev))
            start = r
        prev = r
    bands.append((start, prev))
    return bands


def _split_cols(colidx):
    if not colidx:
        return []
    segs, start, prev = [], colidx[0], colidx[0]
    for c in colidx[1:]:
        if c > prev + 1:
            segs.append((start, prev))
            start = c
        prev = c
    segs.append((start, prev))
    return segs


def _merge_cover(r1, c1, r2, c2, merged):
    area = (r2 - r1 + 1) * (c2 - c1 + 1)
    cov = 0
    for mr1, mc1, mr2, mc2 in merged:
        if mr1 >= r1 and mr2 <= r2 and mc1 >= c1 and mc2 <= c2:
            cov += (mr2 - mr1 + 1) * (mc2 - mc1 + 1)
    return cov / area if area else 0.0


def _perimeter_border(r1, c1, r2, c2, cells) -> float:
    per = []
    for c in range(c1, c2 + 1):
        per.append((r1, c))
        if r2 > r1:
            per.append((r2, c))
    for r in range(r1 + 1, r2):
        per.append((r, c1))
        if c2 > c1:
            per.append((r, c2))
    if not per:
        return 0.0
    hit = sum(1 for pos in per if cells.get(pos, {}).get("bd"))
    return hit / len(per)


def _carve_annotations(cells, content, rules):
    """批注剥离：a) 整行仅 1~2 个长文本格（如"更新日期：…"）会把上下两张表粘成一块；
    b) 整列仅 1 个内容格的孤点（游离说明）。仅当邻近有密集内容时剥离，避免误杀竖排备注列。
    直接在 content 上剔除。阈值见 rules.json segment.annotation。"""
    ra = rules["segment"]["annotation"]
    markers = rules["vocab"]["note_marker_chars"]
    rows = sorted({r for r, _ in content})
    max_c = max(c for _, c in content)
    row_cells = {r: [c for c in range(1, max_c + 1) if (r, c) in content] for r in rows}
    col_count = {}
    for r, c in content:
        col_count[c] = col_count.get(c, 0) + 1
    # 表头词汇行：同一行里重复出现的字符串标签占比高（如"基准/处理/复合/评价"×N 列
    # 的多级表头叶子行，其下无数据的列整列只有这一个格）。真批注（"更新日期：…"、
    # "小结"）在同一家不会复现，占不上这个比例。
    rep_frac_min = ra.get("header_rep_frac", 0.5)
    row_min = ra.get("text_row_min_cells", 5)
    def _header_vocab_row(r, cs, vals):
        if len(cs) < row_min:
            return False
        cnt = {}
        for v in vals:
            if isinstance(v, str):
                cnt[v] = cnt.get(v, 0) + 1
        rep = sum(n for n in cnt.values() if n >= 2)
        return rep >= rep_frac_min * len(cs)
    content_backup = set(content)
    annotations = {}
    for i, r in enumerate(rows):
        cs = row_cells[r]
        vals = [_val(cells[(r, c)]) for c in cs]
        is_row_note = (1 <= len(cs) <= ra["max_cells"]
                       and all(isinstance(v, str) for v in vals)
                       and any(len(v) >= ra["min_text_len"] or any(m in v for m in markers)
                               for v in vals))
        strays = [] if _header_vocab_row(r, cs, vals) else [
            c for c in cs if col_count[c] == 1 and len(cs) > 1]
        if strays:
            # 只剥"夹在中间"的孤点列：全表视角左右都还有实内容列才剥离，
            # 右边缘的空字段表头列（如 仅表头无数据 的人均列）不是孤点
            solid_all = [c for c, n in col_count.items() if n > 1]
            strays = [c for c in strays
                      if any(sc < c for sc in solid_all) and any(sc > c for sc in solid_all)]
        if not (is_row_note or strays):
            continue
        dense_above = i > 0 and len(row_cells[rows[i - 1]]) >= ra["dense_neighbor"]
        dense_below = i + 1 < len(rows) and len(row_cells[rows[i + 1]]) >= ra["dense_neighbor"]
        if is_row_note and (dense_above or dense_below):
            annotations[r] = cs
        elif strays:
            annotations.setdefault(r, []).extend(strays)
        else:
            continue
        for c in annotations[r]:
            content.discard((r, c))
    if not content:  # 全被剥光的退化保护
        content.update(content_backup)
        return {}
    return annotations


def _merge_continuity(content, merged):
    """纵向合并的空身体行不是分隔（合并体非空）→ 作为幻影行并入切带依据。
    横向合并的列连续性在 segment_sheet 内按带计算，不做全表集合。"""
    phantom_rows = set()
    content_rows = {r for r, _ in content}
    for mr1, mc1, mr2, mc2 in merged:
        if mr2 > mr1:
            for r in range(mr1, mr2 + 1):
                if r not in content_rows:
                    phantom_rows.add(r)
    return phantom_rows


def _classify(block_cells, full_cells, merged, r1, c1, r2, c2, rules):
    rc = rules["segment"]["classify"]
    sc = rules["score"]
    stats = {"n": 0, "text": 0, "num": 0, "formula": 0, "date": 0, "text_len_sum": 0}
    for info in block_cells.values():
        v = _val(info)
        if not _has_value(v):
            continue
        stats["n"] += 1
        if info["f"]:
            stats["formula"] += 1
        if isinstance(v, str):
            stats["text"] += 1
            stats["text_len_sum"] += len(v)
        elif _is_num(v):
            stats["num"] += 1
        else:
            stats["date"] += 1
    n = max(stats["n"], 1)
    text_r = stats["text"] / n
    num_r = (stats["num"] + stats["date"]) / n
    formula_r = stats["formula"] / n
    avg_len = stats["text_len_sum"] / max(stats["text"], 1)
    mcov = _merge_cover(r1, c1, r2, c2, merged)
    pb = _perimeter_border(r1, c1, r2, c2, full_cells)
    area = (r2 - r1 + 1) * (c2 - c1 + 1)
    # 有效密度：分母用"有内容的行×有内容的列"，bbox 面积会被合并占位/剥离区稀释，
    # 导致稀疏可选列（备注/渠道）把正常台账置信分拖进人工复核区
    row_n = len({p[0] for p in block_cells})
    col_n = len({p[1] for p in block_cells})
    fill = stats["n"] / max(row_n * col_n, 1)

    # 首行覆盖率：判定文本台账 / 标题行
    row_cols = {p[1] for p in block_cells}
    first_row_cols = {p[1] for p in block_cells if p[0] == r1}
    first_row_cov = len(first_row_cols) / max(len(row_cols), 1)

    if n <= rc["note_max_cells"] and text_r >= rc["note_text_ratio"]:
        btype = NOTE
    elif avg_len >= rc["note2_avg_len"] and text_r >= rc["note2_text_ratio"] \
            and n < rc["note2_max_cells"]:
        btype = NOTE
    elif num_r + formula_r >= rc["table_num_ratio"]:
        btype = TABLE
    elif stats["n"] >= rc["ledger_min_cells"] and first_row_cov >= rc["ledger_first_row_cov"] \
            and n_rows_of(r1, r2) >= rc["ledger_min_rows"] and (c2 - c1) >= 1:
        btype = TABLE  # 文本台账型（字典表）：首行整齐表头 + 规整填充
    else:
        btype = MIXED

    score = sc["base"]
    if btype == TABLE:
        score += sc["table"]
    if pb >= sc["border_closed_at"]:
        score += sc["border_closed"]
    elif pb >= sc["border_partial_at"]:
        score += sc["border_partial"]
    if btype == TABLE and num_r + formula_r >= sc["num_rich_at"]:
        score += sc["num_rich"]
    if btype == TABLE and first_row_cov >= sc["neat_first_row_at"]:
        score += sc["neat_first_row"]
    if btype == NOTE and text_r >= sc["note_text_at"]:
        score += sc["note_text"]
    if mcov > sc["merge_cover_at"]:
        score -= sc["merge_cover_pen"]
    if fill < sc["sparse_fill_at"]:
        score -= sc["sparse_fill_pen"]
    return btype, min(max(score, 0.0), 1.0), {
        "cells": stats["n"], "text_r": round(text_r, 2), "num_r": round(num_r, 2),
        "formula_r": round(formula_r, 2), "avg_text_len": round(avg_len, 1),
        "merge_cover": round(mcov, 2), "perimeter_border": round(pb, 2),
        "first_row_cov": round(first_row_cov, 2), "fill": round(fill, 2),
    }


def n_rows_of(r1, r2):
    return r2 - r1 + 1


def segment_sheet(sheet, rules=None) -> dict:
    rules = rules or get_rules()
    cells = sheet["cells"]
    merged = sheet["merged"]
    content = {pos for pos, info in cells.items() if _has_value(_val(info))}
    if not content:
        return {"blocks": [], "empty": True}

    annotations = _carve_annotations(cells, content, rules)
    phantom_rows = _merge_continuity(content, merged)

    rows_effective = sorted({r for r, _ in content} | phantom_rows)
    blocks = []
    for b_r1, b_r2 in _bands(rows_effective):
        band_content_cols = {c for r, c in content if b_r1 <= r <= b_r2}
        # 合并区列连续性：只并入【与本带相交的合并区自己覆盖的列】，不能全表混入，
        # 否则带内永不留缝、左右并排表被粘死
        cover_cols = set()
        for mr1, mc1, mr2, mc2 in merged:
            if mc2 > mc1 and not (mr2 < b_r1 or mr1 > b_r2):
                cover_cols.update(range(mc1, mc2 + 1))
        band_cols = sorted(band_content_cols | cover_cols)
        for c1, c2 in _split_cols(band_cols):
            full = {p: info for p, info in cells.items()
                    if b_r1 <= p[0] <= b_r2 and c1 <= p[1] <= c2}
            block_content = {p: full[p] for p in content
                             if b_r1 <= p[0] <= b_r2 and c1 <= p[1] <= c2}
            btype, score, sig = _classify(block_content, full, merged, b_r1, c1, b_r2, c2, rules)
            blocks.append({
                "id": 0, "bbox": [b_r1, c1, b_r2, c2], "type": btype,
                "type_label": TYPE_LABEL[btype], "score": round(score, 2),
                "signals": sig,
                "n_rows": b_r2 - b_r1 + 1, "n_cols": c2 - c1 + 1,
            })

    blocks = _bridge_border(blocks, cells, rules)
    blocks = _attach_summary_tail(blocks, cells, content, rules)

    # 批注行 → 独立 note 区块（不参与表格几何）
    for r in sorted(annotations):
        cs = annotations[r]
        blocks.append({
            "id": 0, "bbox": [r, min(cs), r, max(cs)], "type": NOTE,
            "type_label": TYPE_LABEL[NOTE], "score": 0.9,
            "signals": {"cells": len(cs), "annotation": True},
            "n_rows": 1, "n_cols": len(cs), "annotation": True,
        })

    blocks.sort(key=lambda b: (b["bbox"][0], b["bbox"][1]))
    for i, b in enumerate(blocks):
        b["id"] = i
    return {"blocks": blocks, "empty": False}


def _bridge_border(blocks, cells, rules):
    """边框桥接：相邻带列对齐、中间空行全带边框 → 同一张表被空行断开，合回。"""
    br = rules["segment"]["bridge"]
    if len(blocks) < 2:
        return blocks
    out = [blocks[0]]
    for b in blocks[1:]:
        top = out[-1]
        if not (top["type"] == TABLE and b["type"] == TABLE):
            out.append(b)
            continue
        r_gap = b["bbox"][0] - top["bbox"][2] - 1
        c_overlap = min(top["bbox"][3], b["bbox"][3]) - max(top["bbox"][1], b["bbox"][1]) + 1
        c_min = min(top["bbox"][3] - top["bbox"][1], b["bbox"][3] - b["bbox"][1]) + 1
        if 1 <= r_gap <= br["max_gap_rows"] and c_min > 0 \
                and c_overlap / c_min >= br["min_overlap"]:
            gap_rows = range(top["bbox"][2] + 1, b["bbox"][0])
            span = range(max(top["bbox"][1], b["bbox"][1]),
                         min(top["bbox"][3], b["bbox"][3]) + 1)
            if all(cells.get((r, c), {}).get("bd") for r in gap_rows for c in span):
                top["bbox"] = [top["bbox"][0], min(top["bbox"][1], b["bbox"][1]),
                               b["bbox"][2], max(top["bbox"][3], b["bbox"][3])]
                top["n_rows"] = top["bbox"][2] - top["bbox"][0] + 1
                top["n_cols"] = top["bbox"][3] - top["bbox"][1] + 1
                top["signals"]["cells"] += b["signals"]["cells"]
                continue
        out.append(b)
    return out


def _attach_summary_tail(blocks, cells, content, rules):
    """合计行回贴：单行区块（含被列缝切碎的同批碎片）+ 与上方表格列重叠达标 + 空档有限
    → 判定为该表被空行切断的表尾汇总行，合回并标记。回溯时跳过批注等单行小块。"""
    import re
    at = rules["segment"]["attach"]
    exact_sum = re.compile(r"=SUM\(|=SUBTOTAL\(", re.I)
    out = []
    for b in blocks:
        if b["n_rows"] == 1 and b["type"] in (TABLE, NOTE):
            target = None
            for prev in reversed(out):
                gap = b["bbox"][0] - prev["bbox"][2] - 1
                if gap > at["max_gap_rows"]:
                    break
                if prev["type"] == TABLE and prev["n_rows"] >= at["min_table_rows"]:
                    overlap = min(prev["bbox"][3], b["bbox"][3]) \
                        - max(prev["bbox"][1], b["bbox"][1]) + 1
                    c_min = min(prev["bbox"][3] - prev["bbox"][1],
                                b["bbox"][3] - b["bbox"][1]) + 1
                    if overlap / max(c_min, 1) >= at["min_overlap"]:
                        target = prev
                        break
                if prev["n_rows"] > 1:
                    break  # 隔着多行块（另一张表），不回贴
            if target is not None:
                vals = [(_val(cells[p]), cells[p]) for p in content
                        if b["bbox"][0] <= p[0] <= b["bbox"][2]
                        and b["bbox"][1] <= p[1] <= b["bbox"][3]]
                has_sum_formula = any(i["f"] and isinstance(i["v"], str) and exact_sum.search(i["v"])
                                      for _, i in vals)
                has_sum_label = any(isinstance(v, str) and is_summary_label(v, rules)
                                    for v, _ in vals)
                if has_sum_formula or has_sum_label:
                    target["bbox"] = [target["bbox"][0],
                                      min(target["bbox"][1], b["bbox"][1]),
                                      b["bbox"][2],
                                      max(target["bbox"][3], b["bbox"][3])]
                    target["n_rows"] = target["bbox"][2] - target["bbox"][0] + 1
                    target["n_cols"] = target["bbox"][3] - target["bbox"][1] + 1
                    target["summary_tail"] = True
                    continue
        out.append(b)
    return out
