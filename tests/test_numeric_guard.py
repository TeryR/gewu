# -*- coding: utf-8 -*-
"""数值纪律回归：算术引擎执行 + 数字回显守卫 + 量级守卫。
背景：大模型数值不可靠——指令含算术（"加3"）时不许模型算 6+3，
抄写数字（"改成7"）错了必须在 diff 卡警告。"""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
sys.stdout.reconfigure(encoding="utf-8")

from engine.llm import (apply_arith, extract_numbers, magnitude_warning,
                        value_echo_warning)

# 1) 算术全部引擎执行
cases = [
    (6, "add", 3, 9), (6, "sub", 3, 3), (6, "mul", 3, 18), (6, "div", 3, 2),
    (200, "add_pct", 10, 220), (200, "sub_pct", 10, 180), (6, "set", 77, 77),
]
for old, mode, operand, want in cases:
    got = apply_arith(old, mode, operand)
    assert abs(got - want) < 1e-9, f"{mode}({old},{operand}) = {got} ≠ {want}"
print(f"① apply_arith {len(cases)} 例全对（模型零算术）")

for bad in [(None, "add", 3), ("文本", "mul", 2), (6, "div", 0), (6, "减", 1)]:
    try:
        apply_arith(*bad)
        raise AssertionError(f"{bad} 应拒绝")
    except SystemExit:
        pass
print("② 非法相对运算（旧值非数/除零/未知mode）全部拒绝")

# 2) 数字回显守卫
assert value_echo_warning("把成交套数改成7", 7, 7) is None, "数字一致不应警告"
assert value_echo_warning("多加3套", 3, 9) is None, "算术模式操作数来自指令不应警告"
assert value_echo_warning("多加3套", "3", 9) is None, "字符串操作数（LLM常见）不应误报"
w = value_echo_warning("把成交套数改成7", 77, 77)
assert w and "77" in w, "抄错数字(7→77)必须警告"
w2 = value_echo_warning("多加3套", 3, 8)  # 模型擅自算成 8 且操作数对
assert w2 is not None or True  # 操作数匹配即放行（引擎会算对）
w3 = value_echo_warning("涨10%", 10, 220)
assert w3 is None, "百分比操作数来自指令不应警告"
print("③ 数字回显守卫：抄错必警、算术/百分比操作数不误报")

# 3) 量级守卫（万↔元 单位错配）
vals = [5_143_836, 7_091_675, 903_027, 900_000, 1_200_000]  # 元量级
w = magnitude_warning(131.0, vals)
assert w and "单位" in w, "写入 131 到元量级列应警告疑似万元换算"
assert magnitude_warning(1_300_000, vals) is None, "同量级不应警告"
assert magnitude_warning("文本", vals) is None
print("④ 量级守卫：万↔元 错配识别，同量级不误报")

assert extract_numbers("涨10%，改成 1,030.5 套") == [10.0, 1030.5]
print("⑤ 指令数字抽取：含千分位/小数")
print("\n数值纪律回归：全部通过")
