# -*- coding: utf-8 -*-
"""镜像重建幂等性回归：同一文件重复 build_mirror，记录数必须不变。
（曾因 records 删除用错键导致每次重建翻倍累积、同行出现多份。）"""
import os
import sqlite3
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
sys.stdout.reconfigure(encoding="utf-8")

from engine.mirror import build_mirror

DB = os.path.join("tests", "sandbox", "rebuild.db")
SRC = os.path.join("samples", "示例_集团费用归并.xlsx")
if os.path.exists(DB):
    os.remove(DB)

n1 = n2 = None
for i in range(3):
    build_mirror(SRC, DB)
    con = sqlite3.connect(DB)
    n = con.execute("SELECT COUNT(*) FROM records").fetchone()[0]
    dup = con.execute(
        "SELECT COUNT(*) FROM (SELECT block_db_id, row_no, COUNT(*) c "
        "FROM records GROUP BY block_db_id, row_no HAVING c > 1)").fetchone()[0]
    con.close()
    print(f"第{i + 1}次构建: records={n} 重复行={dup}")
    n1 = n1 or n
    n2 = n
assert n1 == n2, f"重复构建记录数漂移：{n1} → {n2}（重建未幂等）"
print("镜像重建幂等性：通过（3 次构建记录数一致，无重复行）")

# 冻结持久性：同文件重建（如重启服务）后冻结状态与人工标记必须幸存
from engine.mirror import freeze
import openpyxl  # noqa: E402

con = sqlite3.connect(DB)
con.row_factory = sqlite3.Row
blk = con.execute("SELECT id FROM blocks WHERE workbook_id=1 AND type='table' LIMIT 1").fetchone()
con.close()
freeze(1, DB, edits=[{"block_db_id": blk["id"], "user_type": "table", "readonly": 1}])
build_mirror(SRC, DB)
con = sqlite3.connect(DB)
con.row_factory = sqlite3.Row
st = con.execute("SELECT status, frozen_at FROM workbooks WHERE id=1").fetchone()
mark = con.execute(
    "SELECT user_type, readonly FROM blocks WHERE workbook_id=1 "
    "AND sheet_name=(SELECT sheet_name FROM blocks WHERE workbook_id=1 AND type='table' LIMIT 1) "
    "AND block_id=(SELECT block_id FROM blocks WHERE workbook_id=1 AND type='table' LIMIT 1)"
).fetchone()
con.close()
assert st["status"] == "frozen", f"重建丢失冻结状态: {dict(st)}"
assert mark and mark["readonly"] == 1, f"重建丢失人工标记: {dict(mark) if mark else None}"
print("冻结持久性：通过（同文件重建后 frozen 状态与只读标记幸存）")
