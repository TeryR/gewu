# -*- coding: utf-8 -*-
"""报告层：汇总 抽取→分割→表头树 结果，输出 JSON（全量）与 Markdown（人读）。"""
import json
import os

from .extract import extract_workbook
from .headers import build_header_tree
from .segment import segment_sheet


def analyze_file(path: str, rules=None) -> dict:
    wb = extract_workbook(path)
    for sheet in wb["sheets"]:
        seg = segment_sheet(sheet, rules)
        table_blocks = [b for b in seg["blocks"] if b["type"] == "table"]
        for b in seg["blocks"]:
            if b["type"] == "table":
                b["header"] = build_header_tree(sheet, b, rules)
        sheet["blocks"] = seg["blocks"]
        sheet["sheet_flags"] = {
            "multi_table": len(table_blocks) > 1,
            "n_table_blocks": len(table_blocks),
            "n_note_blocks": len([b for b in seg["blocks"] if b["type"] == "note"]),
            "n_mixed_blocks": len([b for b in seg["blocks"] if b["type"] == "mixed"]),
            "has_summary_rows": any(b.get("header", {}).get("summary_rows") for b in table_blocks),
            "has_dup_headers": any(b.get("header", {}).get("dup_header_names") for b in table_blocks),
            "has_diagonal": any(b.get("header", {}).get("diagonal_header") for b in table_blocks),
        }
    return wb


def to_markdown(wb: dict) -> str:
    lines = [f"# 解析报告：{os.path.basename(wb['file'])}", ""]
    if wb["sanitized"]:
        lines.append("> ⚠ 原文件含非标样式（已自动归一化修复后解析，原文件未改动）")
        lines.append("")
    if wb.get("values_degraded"):
        lines.append("> ⚠⚠ 缓存值加载失败已降级：公式格数值退化为公式串，问答数值不可用，几何结构不受影响")
        lines.append("")
    lines.append(f"- Sheet 数：{len(wb['sheets'])}")
    n_all_t = sum(s["sheet_flags"]["n_table_blocks"] for s in wb["sheets"])
    n_all_mixed = sum(s["sheet_flags"]["n_mixed_blocks"] for s in wb["sheets"])
    lines.append(f"- 全簿数据表区块：{n_all_t}，混合/待复核区块：{n_all_mixed}")
    for s in wb["sheets"]:
        vis = "" if s["visible"] else "（隐藏页）"
        lines.append(f"\n## [{s['name']}]{vis}  {s['declared_dims'][0]}行x{s['declared_dims'][1]}列  "
                     f"合并{len(s['merged'])}处  "
                     f"图片{s['n_images']} 图表{s['n_charts']}")
        f = s["sheet_flags"]
        flags = []
        if f["multi_table"]:
            flags.append(f"⚠同Sheet多表({f['n_table_blocks']}张)")
        if f["has_summary_rows"]:
            flags.append("含汇总行")
        if f["has_dup_headers"]:
            flags.append("⚠同名表头")
        if f["has_diagonal"]:
            flags.append("⚠斜表头标记")
        if s["hidden_rows"]:
            flags.append(f"隐藏行{len(s['hidden_rows'])}")
        if s["hidden_cols"]:
            flags.append(f"隐藏列{len(s['hidden_cols'])}")
        lines.append(f"   {'  '.join(flags) if flags else '—'}")
        lines.append("")
        lines.append("| # | 位置(行x列) | 类型 | 置信 | 行x列 | 要点 |")
        lines.append("|---|---|---|---|---|---|")
        for b in s["blocks"]:
            r1, c1, r2, c2 = b["bbox"]
            icon = {"table": "🟦", "note": "🟨", "mixed": "🟥"}.get(b["type"], "")
            pts = []
            hr = [r for r in s["hidden_rows"] if r1 <= r <= r2]
            hc = [c for c in s["hidden_cols"] if c1 <= c <= c2]
            if hr:
                pts.append(f"隐藏行x{len(hr)}")
            if hc:
                pts.append(f"隐藏列x{len(hc)}")
            h = b.get("header")
            if b.get("annotation"):
                pts.append("📝批注行")
            if h:
                if h.get("title_text"):
                    pts.append(f"标题「{h['title_text'][:12]}」")
                if h["header_depth"] > 1:
                    pts.append(f"{h['header_depth']}层表头")
                if h.get("floating_rows"):
                    pts.append(f"⚠浮动草稿区x{len(h['floating_rows'])}行")
                if h["cross_tab"]:
                    pts.append("交叉表")
                if h["diagonal_header"]:
                    pts.append("斜表头?")
                if h["summary_rows"]:
                    pts.append(f"汇总行x{len(h['summary_rows'])}")
                if b.get("summary_tail"):
                    pts.append("⚠空行切断的尾合计(已回贴)")
                if h["dup_header_names"]:
                    pts.append(f"同名列{','.join(h['dup_header_names'][:3])}")
                if h["key_cols"]:
                    pts.append(f"键列{len(h['key_cols'])}")
                leaves = len(h["leaf_cols"])
                pts.append(f"字段{leaves}")
            if not pts:
                pts.append(f"文本占比{b['signals']['text_r']}")
            lines.append(f"| {b['id']} | r{r1}-r{r2} x c{c1}-c{c2} | {icon}{b['type_label']} "
                         f"| {b['score']} | {b['n_rows']}x{b['n_cols']} | {'；'.join(pts)} |")
    return "\n".join(lines)


def save_reports(wb: dict, out_dir: str, stem: str):
    os.makedirs(out_dir, exist_ok=True)
    jpath = os.path.join(out_dir, f"{stem}.json")
    mpath = os.path.join(out_dir, f"{stem}.md")
    # JSON 里去掉巨型 cells 原始字典，报告层只保留统计与结构
    slim = {k: v for k, v in wb.items() if k != "sheets"}
    slim["sheets"] = []
    for s in wb["sheets"]:
        s2 = {k: v for k, v in s.items() if k != "cells"}
        s2["n_cells"] = len(s["cells"])
        slim["sheets"].append(s2)
    with open(jpath, "w", encoding="utf-8") as f:
        json.dump(slim, f, ensure_ascii=False, indent=1, default=str)
    with open(mpath, "w", encoding="utf-8") as f:
        f.write(to_markdown(wb))
    return jpath, mpath
