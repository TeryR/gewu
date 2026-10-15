# -*- coding: utf-8 -*-
"""区块类型切换的下游实效：确认页的承诺必须与行为一致。

切为 文本/批注|忽略 → 退出问答（blocks tables_only 查不到）+ 改数被拒；
切为 混合待复核 → 改数被拒（不定型不许动）；
引擎判为文本的块切"数据表" → 无列结构，改数仍拒。"""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
sys.stdout.reconfigure(encoding="utf-8")

from openpyxl import Workbook

from engine.mirror import build_mirror, freeze
from engine.query import Mirror
from engine.writer import validate_plan

SB = os.path.join("tests", "sandbox")
os.makedirs(SB, exist_ok=True)
PATH = os.path.join(SB, "type_effect.xlsx")
DB = os.path.join(SB, "type_effect.db")


def _build():
    wb = Workbook()
    ws = wb.active
    ws.title = "台账"
    ws["A1"] = "项目"; ws["B1"] = "数量"
    for i, (p, q) in enumerate([("甲", 10), ("乙", 20), ("丙", 30)]):
        ws.cell(row=2 + i, column=1, value=p)
        ws.cell(row=2 + i, column=2, value=q)
    wb.save(PATH)


def _switch(wid, db, sheet, to):
    """模拟确认页：把该 sheet 唯一区块切成 to 并冻结。"""
    import sqlite3
    con = sqlite3.connect(db)
    bid = con.execute("SELECT id FROM blocks WHERE workbook_id=? AND sheet_name=?",
                      (wid, sheet)).fetchone()[0]
    con.close()
    freeze(wid, db, edits=[{"block_db_id": bid, "user_type": to}])
    return bid


_build()
if os.path.exists(DB):
    os.remove(DB)
wid = build_mirror(PATH, DB)
op = {"sheet": "台账", "row": 2, "col": 2, "value": 99}

print("== 基线：table 态，问答可用、写入放行 ==")
freeze(wid, DB)
m = Mirror(DB, wid)
assert len(m.blocks("台账", tables_only=True)) == 1
assert m.aggregate(m.blocks("台账", tables_only=True)[0], "数量", "sum")["value"] == 60
assert validate_plan([op], PATH, DB, wid)["ok"]

for to, desc in [("note", "文本/批注"), ("ignore", "忽略"), ("mixed", "混合待复核")]:
    print(f"\n== 切为 {desc} → 问答消失 + 改数被拒 ==")
    _switch(wid, DB, "台账", to)
    m = Mirror(DB, wid)
    tabs = m.blocks("台账", tables_only=True)
    assert not tabs, f"{to}: 切走后不应再出现在问答里"
    v = validate_plan([op], PATH, DB, wid)
    assert not v["ok"], f"{to}: 写入应被拒绝"
    print(f"  ✓ 问答区块=0；拒写理由: {v['rejections'][0][:44]}…")
    # 切回 table 恢复
    _switch(wid, DB, "台账", "table")
    m2 = Mirror(DB, wid)
    assert len(m2.blocks("台账", tables_only=True)) == 1
    assert validate_plan([op], PATH, DB, wid)["ok"]
    print(f"  ✓ 切回数据表后恢复（问答/写入均可用）")

print("\n== 引擎判为文本的块切'数据表'（user_type=table）→ 仍不进问答（无列结构） ==")
import sqlite3
wb2 = Workbook(); ws2 = wb2.active; ws2.title = "备注页"
ws2["A1"] = "说明：本页仅文字，无表格结构"
wb2.save(PATH)  # 重存文件追加 sheet（hash 变化走接管，标记迁移）
wid = build_mirror(PATH, DB)  # 结构变化 → parsed；重新冻结
freeze(wid, DB)
con = sqlite3.connect(DB)
con.execute("UPDATE blocks SET user_type='table' WHERE workbook_id=? AND type='note'",
            (wid,))
con.commit(); con.close()
m = Mirror(DB, wid)
note_sheets = [b["sheet_name"] for b in m.blocks(tables_only=True)]
assert "备注页" not in note_sheets, "文本块即便人工标成 table，无列结构也不得进问答"
print("  ✓ 备注页未进入问答区块（tables_only 只认引擎表）")

print("\n全部通过 ✓")
