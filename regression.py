# -*- coding: utf-8 -*-
"""回归门禁：快照对比 + 结构不变量断言。
用法：
  python regression.py --update   # 首次生成/有意变更后更新黄金快照
  python regression.py            # 门禁：任何解析行为变化都会红
纪律：动 engine/ 任何解析逻辑前先跑一次确认绿，改完再跑，差异必须能用人话说清楚。"""
import glob
import json
import math
import os
import sys

sys.stdout.reconfigure(encoding="utf-8")

from engine.report import analyze_file

GOLDEN_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "tests", "golden")


def fingerprint(wb: dict) -> dict:
    """结构指纹：只取语义相关字段，快照稳定不抖动。"""
    fp = {"sanitized": wb["sanitized"], "sheets": []}
    for s in wb["sheets"]:
        sf = {"name": s["name"], "visible": s["visible"],
              "n_merged": len(s["merged"]), "blocks": []}
        for b in s["blocks"]:
            h = b.get("header") or {}
            sf["blocks"].append({
                "bbox": b["bbox"], "type": b["type"], "score": b["score"],
                "n_rows": b["n_rows"], "n_cols": b["n_cols"],
                "annotation": bool(b.get("annotation")),
                "summary_tail": bool(b.get("summary_tail")),
                "header": {
                    "rows": h.get("header_rows"), "title": h.get("title_row"),
                    "floating": h.get("floating_rows"),
                    "depth": h.get("header_depth"),
                    "leaves": len(h.get("leaf_cols") or []),
                    "keys": h.get("key_cols"),
                    "cross": h.get("cross_tab"),
                    "summary": h.get("summary_rows"),
                    "dups": h.get("dup_header_names"),
                    "data_zone": h.get("data_rows"),
                } if h else None,
            })
        fp["sheets"].append(sf)
    return fp


def check_invariants(wb: dict) -> list:
    errs = []
    for s in wb["sheets"]:
        mr, mc = s["declared_dims"]
        for b in s["blocks"]:
            r1, c1, r2, c2 = b["bbox"]
            if not (1 <= r1 <= r2 and 1 <= c1 <= c2):
                errs.append(f"[{s['name']}] block{b['id']} bbox 非法 {b['bbox']}")
            if mr and r2 > mr + 1 or mc and c2 > mc + 1:
                errs.append(f"[{s['name']}] block{b['id']} 越界 {b['bbox']} 超 {mr}x{mc}")
            if b["type"] == "table":
                if not b.get("header"):
                    errs.append(f"[{s['name']}] table block{b['id']} 无表头树")
                else:
                    for sr in b["header"].get("summary_rows") or []:
                        if not (r1 <= sr <= r2):
                            errs.append(f"[{s['name']}] block{b['id']} 汇总行 r{sr} 越界")
            if b["score"] is None or (isinstance(b["score"], float) and math.isnan(b["score"])):
                errs.append(f"[{s['name']}] block{b['id']} 置信分非法")
        ids = [b["id"] for b in s["blocks"]]
        if ids != list(range(len(ids))):
            errs.append(f"[{s['name']}] 区块编号不连续 {ids}")
    return errs


def main():
    update = "--update" in sys.argv
    os.makedirs(GOLDEN_DIR, exist_ok=True)
    root = os.path.dirname(os.path.abspath(__file__))
    # 挑战集（T1~T8，程序生成的虚构数据，开源安全）——开源后的主要回归源
    challenge_dir = os.path.join(os.path.dirname(root), "复杂Excel挑战集")
    samples = []
    if os.path.isdir(challenge_dir):
        samples = sorted(glob.glob(os.path.join(challenge_dir, "*.xlsx")))
    # 真实样本（含真实业务数据，.gitignore 排除——本机跑有就加）
    samples += sorted(glob.glob(os.path.join(root, "samples", "*.xlsx")))
    # 真实布料文件（多级表头/竖合并/转置矩阵的代表，本机路径）
    fabric = r"C:\Users\Administrator\Desktop\后整理不后整理布 浸胶、贴合 物性对比20261001.xlsx"
    if os.path.isfile(fabric):
        samples.append(fabric)
    assert samples, "无样本文件（挑战集目录或 samples/ 至少需要一处有 .xlsx）"
    failed = 0
    for path in samples:
        stem = os.path.splitext(os.path.basename(path))[0]
        wb = analyze_file(path)
        errs = check_invariants(wb)
        fp = fingerprint(wb)
        gpath = os.path.join(GOLDEN_DIR, f"{stem}.json")
        if errs:
            print(f"✗ {stem} 不变量违例:")
            for e in errs:
                print(f"    {e}")
            failed += 1
            continue
        if update or not os.path.exists(gpath):
            with open(gpath, "w", encoding="utf-8") as f:
                json.dump(fp, f, ensure_ascii=False, indent=1, sort_keys=True)
            print(f"● {stem} 快照已写入")
            continue
        with open(gpath, encoding="utf-8") as f:
            golden = json.load(f)
        if fp == golden:
            n = sum(len(sf["blocks"]) for sf in fp["sheets"])
            print(f"✓ {stem} 一致（{n} 区块）")
        else:
            print(f"✗ {stem} 与快照有差异：")
            for gs, cs in zip(golden["sheets"], fp["sheets"]):
                gb = {(tuple(b["bbox"])): b for b in gs["blocks"]}
                cb = {(tuple(b["bbox"])): b for b in cs["blocks"]}
                for k in sorted(set(gb) - set(cb)):
                    print(f"    [{gs['name']}] 消失区块 {list(k)} {gb[k]['type']}")
                for k in sorted(set(cb) - set(gb)):
                    print(f"    [{cs['name']}] 新增区块 {list(k)} {cb[k]['type']}")
                for k in sorted(set(gb) & set(cb)):
                    if gb[k] != cb[k]:
                        diff = {kk for kk in gb[k] if gb[k][kk] != cb[k].get(kk)}
                        print(f"    [{gs['name']}] 区块 {list(k)} 字段变化 {sorted(diff)}")
            failed += 1
    if failed:
        print(f"\n回归未通过：{failed} 个文件")
        sys.exit(1)
    print("\n回归全部通过")


if __name__ == "__main__":
    main()
