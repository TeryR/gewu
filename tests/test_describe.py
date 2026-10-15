# -*- coding: utf-8 -*-
"""理解型问答（op=describe）：『这个表格是做什么的』这类问题不再打回 unknown。

链路：LLM 判 describe → 引擎组描述上下文（标题/列/样例/口径/行备注）→
LLM 自然语言回答（零算术）；LLM 挂了 → 代码结构化兜底；
无 LLM 的规则版也直接给结构化描述。"""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
sys.stdout.reconfigure(encoding="utf-8")

from openpyxl import Workbook

from engine import llm
from engine.mirror import build_mirror
from engine.query import Mirror

SB = os.path.join("tests", "sandbox")
os.makedirs(SB, exist_ok=True)
PATH = os.path.join(SB, "describe.xlsx")
DB = os.path.join(SB, "describe.db")


def _build():
    wb = Workbook()
    ws = wb.active
    ws.title = "销售台账"
    ws["A1"] = "项目"; ws["B1"] = "月份"; ws["C1"] = "成交套数"; ws["D1"] = "创收(万)"; ws["E1"] = "完成率"
    for i, row in enumerate([("云澜居", "1月", 3, 85.3, "=D2/100"),
                             ("云澜居", "2月", 5, 142.0, "=D3/100"),
                             ("青枫苑", "1月", 2, 60.5, "=D4/100")]):
        for j, v in enumerate(row):
            ws.cell(row=2 + i, column=1 + j, value=v)
    ws["A6"] = "合计"; ws["C6"] = 10; ws["D6"] = 287.8
    ws2 = wb.create_sheet("月度汇总")
    ws2["A1"] = "月份"; ws2["B1"] = "创收(万)"
    ws2["A2"] = "1月"; ws2["B2"] = 85.3; ws2["A3"] = "2月"; ws2["B3"] = 142.0
    wb.save(PATH)


print("== 构建：销售台账（含汇总行） ==")
_build()
if os.path.exists(DB):
    os.remove(DB)
wid = build_mirror(PATH, DB)
m = Mirror(DB, wid)

print("\n== 链路1：LLM 判 describe → describe_block 自然语言回答 ==")
llm.chat = lambda messages, **kw: '{"op":"describe","block_hint":"销售台账"}'
orig_describe = llm.describe_block
seen = {}

def fake_describe(question, ctx, fn=None):
    seen["q"] = question
    seen["ctx"] = ctx
    return "这是各项目按月记录成交套数与创收的台账。"

llm.describe_block = fake_describe
ans = m.answer("这个表格是做什么的", llm=llm.query_intent)
print(f"  回答: {ans}")
assert ans == "这是各项目按月记录成交套数与创收的台账。"
assert seen["q"] == "这个表格是做什么的"
ctx = seen["ctx"]
assert ctx["sheet"] == "销售台账" and "成交套数" in ctx["columns"]
assert ctx["sample_rows"] and ctx["sample_rows"][0]["项目"] == "云澜居"
print(f"  上下文: sheet={ctx['sheet']} 列={ctx['columns']} 样例行数={len(ctx['sample_rows'])}")

print("\n== 链路1b：公式进上下文（完成率列带 =D2/100） ==")
assert ctx.get("列公式样例") and "D2/100" in ctx["列公式样例"].get("完成率", ""), ctx.get("列公式样例")
print(f"  列公式样例: {ctx['列公式样例']}")

print("\n== 链路1c：多轮历史透传（『那/它』有处安放） ==")
seen_hist = {}

def fake_describe_hist(question, ctx, history=None):
    seen_hist["history"] = history
    return "按上一轮说的算。"

llm.describe_block = fake_describe_hist
llm.chat = lambda messages, **kw: '{"op":"describe","block_hint":"销售台账"}'
hist = [{"role": "user", "content": "完成率是什么"},
        {"role": "assistant", "content": "完成率=创收/100"}]
m.answer("那为什么这么算", llm=llm.query_intent, history=hist)
assert seen_hist["history"] == hist, seen_hist
print("  ✓ describe_block 收到 history")

print("\n== 链路2：sheet_hint 生效（LLM 没给 block_hint 时用当前页） ==")
llm.describe_block = lambda q, c, history=None: "这是销售台账。"
llm.chat = lambda messages, **kw: '{"op":"describe"}'
ans = m.answer("这个表格是做什么的", llm=llm.query_intent, sheet_hint="销售台账")
assert "台账" in ans
print(f"  ✓ {ans[:40]}")

print("\n== 链路3：describe 调用挂了 → 结构化兜底（不哑） ==")
def boom(question, ctx, fn=None):
    raise SystemExit("API 超时")

llm.describe_block = boom
ans = m.answer("这个表格是做什么的", llm=llm.query_intent)
assert "[销售台账]" in ans and "成交套数" in ans and "字段" in ans
print("  ✓ 兜底输出含 sheet/列/行数")

print("\n== 链路4：无 LLM 规则版也能结构化描述 ==")
llm.describe_block = orig_describe
ans = m.answer("这个表格是做什么的", llm=None)
assert "[销售台账]" in ans and "成交套数" in ans
print(f"  ✓ {ans.splitlines()[0]}")

print("\n== 链路5：全簿模式（block_hint=*）整个文件是做什么的 ==")
llm.chat = lambda messages, **kw: '{"op":"describe","block_hint":"*"}'
seen5 = {}

def fake_wb(question, ctx):
    seen5["ctx"] = ctx
    return "这是销售明细加月度汇总的文件。"

llm.describe_block = fake_wb
ans = m.answer("整个文件是做什么的", llm=llm.query_intent)
assert ans == "这是销售明细加月度汇总的文件。"
ctx5 = seen5["ctx"]
assert ctx5.get("workbook") is True and ctx5["n_sheets"] == 2
assert {s["sheet"] for s in ctx5["sheets"]} == {"销售台账", "月度汇总"}
print(f"  ✓ 全簿上下文: {ctx5['n_sheets']} sheets, tables={[(s['sheet'], s['n_rows']) for s in ctx5['sheets']]}")

# sheet_hint 不得覆盖显式的 "*"
ans = m.answer("整个文件是做什么的", llm=llm.query_intent, sheet_hint="销售台账")
assert ans == "这是销售明细加月度汇总的文件。"

def boom_wb(question, ctx):
    raise SystemExit("API 超时")

llm.describe_block = boom_wb
ans = m.answer("整个文件是做什么的", llm=llm.query_intent)
assert "全簿共 2 个 sheet" in ans and "月度汇总" in ans
print("  ✓ 全簿兜底（LLM 挂了）输出 sheet 清单")

print("\n全部通过 ✓")
