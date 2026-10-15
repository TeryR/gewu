# -*- coding: utf-8 -*-
"""应用内打开：_switch_workbook 切换工作簿（镜像重建+缓存清理+副本标记）。"""
import os
import shutil
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
sys.stdout.reconfigure(encoding="utf-8")

from openpyxl import Workbook

import serve

SB = os.path.join("tests", "sandbox")
os.makedirs(SB, exist_ok=True)
A = os.path.join(SB, "open_a.xlsx")
B = os.path.join(SB, "open_b.xlsx")
for p, tag in ((A, "甲"), (B, "乙")):
    wb = Workbook(); ws = wb.active; ws.title = tag
    ws["A1"] = "k"; ws["B1"] = "v"; ws["A2"] = tag; ws["B2"] = 1
    wb.save(p)

DB = os.path.join(SB, "open_test.db")
if os.path.exists(DB):
    os.remove(DB)
serve.STATE["db"] = DB

print("== 切到 A ==")
serve._switch_workbook(A)
assert serve.STATE["wb_path"] == A and serve.STATE["wid"]
assert serve.STATE["copy"] is False
sheets_a = [s["name"] for s in serve._get_wb(A)["sheets"]]
print("  sheets:", sheets_a)

print("== 切到 B（缓存清理+镜像切换） ==")
serve._switch_workbook(B, copy=True)
assert serve.STATE["wb_path"] == B and serve.STATE["copy"] is True
assert serve._SHEET_CACHE == {}, "sheet 缓存必须清空"
sheets_b = [s["name"] for s in serve._get_wb(B)["sheets"]]
assert sheets_a != sheets_b, "两个文件的 sheet 应不同"
print("  sheets:", sheets_b, "副本标记:", serve.STATE["copy"])

print("== 回 A：镜像复用（同文件同 wid） ==")
wid_a = serve.STATE["wid"]
serve._switch_workbook(B)          # 先到 B
serve._switch_workbook(A)          # 再回 A
con_check = serve.STATE["wid"]
serve._switch_workbook(A)          # 重复打开同文件（幂等）
assert serve.STATE["wid"] == con_check, "同文件重复打开 wid 稳定"
print("  ✓ 幂等")

print("\n全部通过 ✓")
