# -*- coding: utf-8 -*-
"""行键匹配回归：精确匹配守卫（"1月"不得命中"11月"——曾被子串匹配引入的正确性 bug）。"""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
sys.stdout.reconfigure(encoding="utf-8")

import openpyxl

from engine import llm
from engine.mirror import build_mirror, freeze
from engine.query import Mirror

SB = os.path.join("tests", "sandbox")
p = os.path.join(SB, "month_keys.xlsx")
wb = openpyxl.Workbook()
ws = wb.active
ws.title = "台账"
ws.append(["月份", "项目", "套数"])
ws.append(["1月", "云澜居", 6])
ws.append(["11月", "云澜居", 66])
ws.append(["12月", "云澜居", 88])
ws.append(["1月", "回购房", 150])
wb.save(p)
db = os.path.join(SB, "month_keys.db")
if os.path.exists(db):
    os.remove(db)
wid = build_mirror(p, db)
freeze(wid, db)
mx = Mirror(db)


def resolve(keys):
    return llm.resolve_write_plan(
        {"find": {"sheet": "台账", "keys": keys}, "set": {"column": "套数", "value": 7}}, mx)


r = resolve({"月份": "1月", "项目": "云澜居"})
assert r["ok"] and r["op"]["row"] == 2, f"1月应唯一命中 r2: {r}"
r = resolve({"月份": "11月"})
assert r["ok"] and r["op"]["row"] == 3, f"11月应唯一命中 r3: {r}"
r = resolve({"月份": "1月"})
assert r.get("ambiguous"), f"仅1月命中多项目应歧义拦截: {r}"
r = resolve({"月份": "2月"})
assert not r["ok"] and not r.get("ambiguous"), f"无命中应报 clarify: {r}"
assert "近似命中" in r.get("clarify", "") and "12月" in r.get("clarify", ""), \
    f"应提示 12月 近似候选供确认而非自动写入: {r}"
print("行键精确匹配回归：4/4 通过（子串仅提示，永不自动写入）")
