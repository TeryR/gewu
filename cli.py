# -*- coding: utf-8 -*-
"""CLI 入口：python cli.py <xlsx...> [--out reports]"""
import argparse
import glob
import os
import sys
import time

sys.stdout.reconfigure(encoding="utf-8")

from engine.report import analyze_file, save_reports, to_markdown
from engine.rules import get_rules


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("files", nargs="+")
    ap.add_argument("--out", default="reports")
    ap.add_argument("--rules", default=None, help="规则文件路径（缺省用项目根 rules.json）")
    args = ap.parse_args()

    rules = get_rules(args.rules)

    paths = []
    for f in args.files:
        paths.extend(sorted(glob.glob(f)) if any(ch in f for ch in "*?") else [f])

    for p in paths:
        t0 = time.time()
        print(f"\n{'=' * 90}\n解析: {os.path.basename(p)}")
        wb = analyze_file(p, rules)
        stem = os.path.splitext(os.path.basename(p))[0]
        jpath, mpath = save_reports(wb, args.out, stem)
        print(to_markdown(wb))
        print(f"\n⏱ {time.time() - t0:.1f}s  JSON: {jpath}")


if __name__ == "__main__":
    main()
