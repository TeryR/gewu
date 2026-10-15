# -*- coding: utf-8 -*-
"""口径词典闭环测试：保存（覆盖式）→ glossary() 读取 → 注入 LLM 上下文。"""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
sys.stdout.reconfigure(encoding="utf-8")

from engine import llm
from engine.mirror import add_glossary, build_mirror
from engine.query import Mirror

DB = os.path.join("tests", "sandbox", "glossary.db")
if os.path.exists(DB):
    os.remove(DB)
wid = build_mirror(os.path.join("samples", "总部经营数据源2026.xlsx"), DB)
m = Mirror(DB, workbook_id=wid)

add_glossary(wid, term="套数", meaning="月度网签套数（非年度累计）",
             db=DB, col_path="套数", sheet_name="代理")
add_glossary(wid, term="套数", meaning="月度网签套数（非年度累计）",  # 重复保存=覆盖
             db=DB, col_path="套数", sheet_name="代理")
add_glossary(wid, term="销售额", meaning="单位：元（表内另有万元口径列）",
             db=DB, col_path="销售额", sheet_name="代理")

gl = m.glossary()
assert len(gl) == 2, f"重复保存应覆盖不堆积，实际 {len(gl)} 条"
print(f"① 词典保存（覆盖式）：{len(gl)} 条")

ctx = llm.mirror_context(m)
assert isinstance(ctx, dict) and "field_notes" in ctx, "有词典时上下文应含 field_notes"
notes = ctx["field_notes"]
assert any(n["column"] == "套数" and "月度网签" in n["口径"] for n in notes)
assert len(ctx["tables"]) > 0, "表上下文仍应保留"
print(f"② mirror_context 注入：field_notes={len(notes)} 条置顶，tables={len(ctx['tables'])} 块")

qctx = m._llm_context()
assert "field_notes" in qctx and len(qctx["field_notes"]) == 2
print("③ 问答上下文（_llm_context）同样注入 ✓")
print("\n口径词典闭环：全部通过")
