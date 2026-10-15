# -*- coding: utf-8 -*-
"""LLM 适配层 mock 全链测试：假 LLM 只吐 JSON 计划，
验证 查询意图→聚合 与 写指令→语义解析→坐标定位→硬校验 两条链路。"""
import os
import shutil
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
sys.stdout.reconfigure(encoding="utf-8")

from engine import llm
from engine.mirror import build_mirror, freeze
from engine.query import Mirror
from engine.writer import validate_plan

SB = os.path.join("tests", "sandbox")
os.makedirs(SB, exist_ok=True)


def fake_chat(responses):
    def _chat(messages, **kw):
        return responses.pop(0)
    return _chat


print("== 链路1：问答意图 → 聚合 ==")
# 自建沙箱镜像：不依赖共享 mirror.db 的"最新工作簿"（用户拖入新文件会换主）
DB1 = os.path.join(SB, "llm_read.db")
if os.path.exists(DB1):
    os.remove(DB1)
build_mirror("samples/总部经营数据源2026.xlsx", DB1)
m = Mirror(DB1)
llm.chat = fake_chat(['{"op":"max","column":"年创收","block_hint":"代理"}'])
plan = llm.query_intent("哪个项目年创收最高", m._llm_context())
ans = m.answer("哪个项目年创收最高", llm=lambda q, context: plan)
print(" ", ans.splitlines()[0])

llm.chat = fake_chat(['{"op":"sum","column":"1月","block_hint":"指标分解",'
                      '"filters":{"部门":"示例A部"}}'])
plan = llm.query_intent("示例A部1月合计多少", m._llm_context())
b = m._pick_block(plan.get("block_hint"))
r = m.aggregate(b, plan["column"], plan.get("op", "sum"), filters=plan.get("filters"))
print(f"  示例A部 1 月 = {r['value']:.2f}（期望 85.34）")

print("\n== 链路2：写指令 → 语义解析 → 坐标 → 硬校验 ==")
SBX = os.path.join(SB, "llm_write.xlsx")
shutil.copy2("samples/总部经营数据源2026.xlsx", SBX)
DBX = os.path.join(SB, "llm_write.db")
if os.path.exists(DBX):
    os.remove(DBX)
wid = build_mirror(SBX, DBX)
freeze(wid, DBX)
mx = Mirror(DBX)

llm.chat = fake_chat(['{"find":{"sheet":"示例B部经营数据明细表","keys":{"项目":"云澜居","月份":"1月"}},'
                      '"set":{"column":"成交套数","value":7}}'])
plan = llm.write_intent("把1月云澜居的成交套数改成7", llm.mirror_context(mx))
res = llm.resolve_write_plan(plan, mx)
print(f"  解析: ok={res['ok']} via={res.get('via')} 定位={res['op']['sheet']}!r{res['op']['row']}c{res['op']['col']}")
v = validate_plan([res["op"]], SBX, DBX, wid)
print(f"  硬校验: {'✓放行' if v['ok'] else '✗' + str(v['rejections'])}")

print("\n== 链路2b：歧义指令（仅项目名，月度台账命中多行）必须被拦截 ==")
res2 = llm.resolve_write_plan(
    {"find": {"sheet": "示例B部", "keys": {"项目": "云澜居"}},
     "set": {"column": "成交套数", "value": 7}}, mx)
print(f"  歧义检测: ok={res2['ok']} ambiguous={res2.get('ambiguous')} "
      f"提示={str(res2.get('clarify'))[:50]}")

print("\n== 链路2c：need_clarify 透传 ==")
llm.chat = fake_chat(['{"need_clarify":"要改哪个月份的？"}'])
plan = llm.write_intent("改套数", llm.mirror_context(mx))
res3 = llm.resolve_write_plan(plan, mx)
print(f"  透传: ok={res3['ok']} clarify={res3.get('clarify')}")

print("\n== 未配置 key 的诚实失败 ==")
llm.chat = llm.__dict__.get("_orig_chat", llm.chat)
import importlib
importlib.reload(llm)
saved = {k: os.environ.pop(k, None) for k in list(os.environ)
         if k.startswith("EXCEL_AGENT_LLM")}
try:
    llm.chat([{"role": "user", "content": "hi"}])
except SystemExit as e:
    print(f"  SystemExit: {str(e)[:40]}...")
finally:
    for k, v in saved.items():
        if v:
            os.environ[k] = v
