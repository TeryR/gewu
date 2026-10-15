# -*- coding: utf-8 -*-
"""提示词守卫：三段 SYSTEM 提示词的关键条款必须存在。
改提示词是允许的（且应真 LLM 实测），但数值纪律/重名列/need_clarify 这些
用血泪换来的条款不许无声消失。"""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
sys.stdout.reconfigure(encoding="utf-8")

from engine import llm

MUST_HAVE = {
    "QUERY_SYSTEM": [
        "完整路径",          # 重名列（原样×N）必须带组头路径
        "区分词",            # 近亲 sheet（汇总表变体）要带规格型号
        "describe",
        "need_clarify",      # unknown 可带追问
        "只输出 JSON",
        "*",                 # 全簿模式：整个文件 → block_hint="*"（v3 新增）
        "整个文件",
        "当前页",            # v3：这个表格/当前表格 → 当前页（v4 修正指代）
        "回指",              # v5：对话历史中的"它/为什么这么算"按历史理解
       "analyze",           # v6：数据解读型问题（评价/趋势/哪个更好）
        "设计合理",           # v7：analyze 全簿模式（这个文件设计合理吗）
    ],
    "ANALYZE_SYSTEM": [
        "不执行算术",        # v2 措辞改写（数值纪律不变）
        "不编造",            # v2 措辞改写
        "统计量",            # 必须说明统计量来自引擎
        "So What",          # v9：麦肯锡 So What 测试
        "金字塔",           # v9：结论先行
        "假设驱动",         # v9：先假设再验证
        "帕累托",           # v9：抓关键少数
        "宁少勿滥",         # v9：质量优先于数量
    ],
    "WRITE_SYSTEM": [
        "你不负责任何算术",   # 数值纪律地基
        "add_pct",            # 百分数 mode
        "逐字一致",           # set 值禁止换算
        "禁止输出坐标",       # LLM 无权决定坐标
        "need_clarify",
        "完整路径",           # 重名列用完整路径（v2 新增）
        "多行",               # 批量指令不许猜（v2 新增）
        "示例",               # few-shot 必须保留（v2 新增）
    ],
    "DESCRIBE_SYSTEM": [
        "不执行算术",        # v2 措辞改写（数值纪律不变）
        "不编造",            # v2 措辞改写
        "百分比",            # v2：百分比列说明
        "「xx列合计」",       # v2：计算问题引导改问法
        "全簿模式",           # v3：sheets 列表时概括各表关系
        "SCQA",            # v9：情境-冲突-问题-答案
        "So What",         # v9：每个发现必须回答"那又怎样"
        "洞察优先",         # v9：多子问题时高价值先答
    ],
}

for name, clauses in MUST_HAVE.items():
    text = getattr(llm, name)
    for c in clauses:
        assert c in text, f"{name} 丢失关键条款「{c}」"
    print(f"  {name}: {len(clauses)} 条款齐全 ✓")

# 上下文契约：describe 上下文必须带行数与百分比列（DESCRIBE_SYSTEM 承诺了它们存在）
from engine.query import Mirror
import json
from openpyxl import Workbook

SB = os.path.join("tests", "sandbox")
os.makedirs(SB, exist_ok=True)
p = os.path.join(SB, "prompt_ctx.xlsx")
wb = Workbook(); ws = wb.active; ws.title = "t"
ws["A1"] = "项目"; ws["B1"] = "完成率"
ws["A2"] = "甲"; ws["B2"] = 0.5; ws["B2"].number_format = "0%"
ws["A3"] = "乙"; ws["B3"] = 0.2; ws["B3"].number_format = "0%"
wb.save(p)
from engine.mirror import build_mirror
db = os.path.join(SB, "prompt_ctx.db")
if os.path.exists(db):
    os.remove(db)
wid = build_mirror(p, db)
m = Mirror(db, wid)
b = m.blocks("t", tables_only=True)[0]
ctx = m._describe_ctx(b)
assert ctx["n_rows"] == 2, ctx
assert ctx["pct_columns"] == ["完成率"], ctx
print("  describe 上下文契约（n_rows/pct_columns）✓")
m.close()

print("\n全部通过 ✓")
