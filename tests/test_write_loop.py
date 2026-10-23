# -*- coding: utf-8 -*-
"""写回链路沙箱实测：在副本上跑全流程，绝不碰原始样本。
场景：合法写 / 公式格拒绝 / 汇总行拒绝 / 表头拒绝 / 合并身体格拒绝 /
      无区块拒绝 / 数值列写文本警告 / 写后校验 / 撤销精确恢复。"""
import json
import os
import shutil
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
sys.stdout.reconfigure(encoding="utf-8")

import openpyxl

from engine.extract import extract_workbook
from engine.mirror import build_mirror, freeze
from engine.writer import execute, post_check, undo_last_batch, validate_plan

SRC = os.path.join("samples", "示例_集团费用归并.xlsx")
SB_DIR = os.path.join("tests", "sandbox")
os.makedirs(SB_DIR, exist_ok=True)
SB = os.path.join(SB_DIR, "write_test_费用归并.xlsx")
shutil.copy2(SRC, SB)  # 每次全新副本
DB = os.path.join(SB_DIR, "write_test.db")
if os.path.exists(DB):
    os.remove(DB)

wid = build_mirror(SB, DB)
freeze(wid, DB)
print(f"镜像+冻结完成 wid={wid}\n")

# 合法写目标：销售部 r3c3 金额（手填数值，先程序性确认非公式）
src = extract_workbook(SB)
target = next(s for s in src["sheets"] if s["name"] == "销售部")["cells"][(3, 3)]
assert not target["f"], "测试前提失败：r3c3 应为手填数值"
old_val = target["cv"]
print(f"合法写目标：销售部 r3c3 金额 = {old_val}（非公式✓）\n")

cases = [
    ("合法：明细数值更新 r3c3 → 20000",
     [{"sheet": "销售部", "row": 3, "col": 3, "value": 20000}], "放行"),
    ("拒绝：公式格 汇总视图 r4c3（跨表取数）",
     [{"sheet": "汇总视图", "row": 4, "col": 3, "value": 131}], "拒绝"),
    ("拒绝：汇总行 销售部 r12c3",
     [{"sheet": "销售部", "row": 12, "col": 3, "value": 999}], "拒绝"),
    ("拒绝：表头区 销售部 r2c1",
     [{"sheet": "销售部", "row": 2, "col": 1, "value": "改表头"}], "拒绝"),
    ("拒绝：合并身体格 人事部 r4c1（A3:A5 合并的身体格）",
     [{"sheet": "人事部", "row": 4, "col": 1, "value": "x"}], "拒绝"),
    ("拒绝：无区块空档 勾稽校验 r10c4",
     [{"sheet": "勾稽校验", "row": 10, "col": 4, "value": 1}], "拒绝"),
    ("警告+放行：数值列写文本 销售部 r4c3",
     [{"sheet": "销售部", "row": 4, "col": 3, "value": "七千"}], "放行"),
]
ops_legal = None
for name, ops, expect in cases:
    v = validate_plan(ops, SB, DB, wid)
    got = "放行" if v["ok"] else "拒绝"
    mark = "✓" if got == expect else "✗期望" + expect
    warn = f"  ⚠警告: {v['warnings'][0]}" if v["warnings"] else ""
    print(f"{mark} [{got}] {name}{warn}")
    if not v["ok"]:
        print(f"         理由: {v['rejections'][0]}")
    assert got == expect, f"{name} 期望{expect}实际{got}: {v}"
    if name.startswith("合法") and v["ok"]:
        ops_legal = ops

print("\n--- 执行合法写 ---")
res = execute(ops_legal, SB, DB, wid)
print(f"engine={res['engine']} applied={res['applied']} backup={os.path.basename(res['backup'])}")
pc = post_check(ops_legal, SB, DB, wid)
print(f"写后校验: {'✓ 值与结构均符' if pc['ok'] else '✗ ' + str(pc['problems'])}")
assert pc["ok"], pc["problems"]

print("\n--- 写后续用（hash 已滚动：重解析复用 wid 冻结态，镜像记录已同步）---")
wid2 = build_mirror(SB, DB)
assert wid2 == wid, "自己写数后 hash 滚动，重开应走同 hash 分支复用 wid"
v = validate_plan(ops_legal, SB, DB, wid2)
print(f"{'✓ 冻结保持，写入仍放行' if v['ok'] else '✗ ' + str(v['rejections'])}")
from engine.query import Mirror as _M
_m = _M(DB, wid2)
_b = next(x for x in _m.blocks("销售部", tables_only=True)
          if any("金额" in "/".join(n["path"])
                 for n in json.loads(x["header_json"]).get("col_tree") or []))
_r = next(r for r in _m._records(_b) if r["_r"] == ops_legal[0]["row"])
assert _r.get("金额（元）") == ops_legal[0]["value"], "apply_ops_to_mirror 应同步记录值"
print(f"  镜像记录已同步: 金额（元）={_r.get('金额（元）')}")
_m.close()

print("\n--- 撤销最近一批（恢复原值）---")
u = undo_last_batch(DB, SB, workbook_id=wid)
print(f"undo: engine={u.get('engine')} restored={u.get('restored')}")
wbv = openpyxl.load_workbook(SB, data_only=True)
after_val = wbv["销售部"].cell(row=3, column=3).value
print(f"撤销后 r3c3 = {after_val}（原值 {old_val} → "
      f"{'✓精确恢复' if after_val == old_val else '✗未恢复'}）")
assert after_val == old_val, "撤销未恢复原值"
