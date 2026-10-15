# -*- coding: utf-8 -*-
"""格物 Gewu · 产品演示脚本（自动驱动全部核心功能，带旁白解说）

用法：
  ① 先启动格物（拖文件到 bat 或 python serve.py 文件.xlsx）
  ② 浏览器打开 http://127.0.0.1:8765（投影给观众看）
  ③ 新开终端运行：python demo.py
  ④ 脚本自动按步骤调用 API，浏览器同步刷新看到效果

每步有旁白解说（终端打印），可设 delay 控制节奏。
"""
import json
import os
import sys
import time
import urllib.request
import urllib.error

sys.stdout.reconfigure(encoding="utf-8")
BASE = "http://127.0.0.1:8765"
DELAY = 3  # 每步间隔秒（给观众看清楚）

# 演示用挑战集文件（程序生成的虚构数据，无真实业务信息）
CHALLENGE = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                         "samples")
SAMPLE = os.path.join(CHALLENGE, "示例_集团费用归并.xlsx")


def say(msg, step=""):
    bar = "─" * 56
    print(f"\n{bar}")
    if step:
        print(f"  ▶ {step}")
        print(f"{bar}")
    for line in msg.split("\n"):
        print(f"  {line}")
    time.sleep(DELAY)


def api(path, payload=None, method=None):
    if payload is not None or method == "POST":
        data = json.dumps(payload or {}).encode("utf-8")
        req = urllib.request.Request(BASE + path, data=data,
                                     headers={"Content-Type": "application/json"})
    else:
        req = urllib.request.Request(BASE + path)
    resp = urllib.request.urlopen(req, timeout=60)
    return json.load(resp)


def run_demo():
    print("=" * 56)
    print("  格物 Gewu · 产品演示")
    print("  复杂表格理解工作台 · 自动化演示脚本")
    print("=" * 56)
    print(f"\n  服务地址: {BASE}")
    print(f"  演示文件: {os.path.basename(SAMPLE)}")
    print(f"  步骤间隔: {DELAY} 秒（可在脚本头部调整）")

    # ===== 第 1 步：引擎状态 =====
    say("大家好，我是格物（Gewu），一个能读懂复杂 Excel 的 AI 工作台。\n"
        "现在演示的是：打开一个集团费用归并表（虚构数据），\n"
        "从结构识别到安全改数的完整流程。",
        "第 1 步 · 引擎状态检查")

    boot = api("/api/boot")
    say(f"引擎状态：{'就绪' if boot['ready'] else '启动中'}\n"
        f"当前工作簿：{api('/api/status').get('status', '未知')}",
        "第 2 步 · 当前状态")

    # ===== 第 3 步：打开文件 =====
    say("首先，通过 API 打开一个真实的地产集团经营数据表。\n"
        "这个文件有 10 个 sheet：5 个部门明细表 + 汇总视图 + 勾稽校验。\n"
        "格物会在 1~2 秒内完成结构解析（快启缓存命中时 0.5 秒）。",
        "第 3 步 · 打开 Excel 文件（API 自动化）")
    r = api("/api/open", {"path": SAMPLE}, "POST")
    say(f"打开结果：{'成功' if r.get('ok') else '失败'}\n"
        f"文件：T1_结构迷宫_集团费用归并.xlsx（虚构演示数据）",
        "第 4 步 · 文件已打开")

    # ===== 第 5 步：查看 sheet =====
    sheets = api("/api/sheets")
    say(f"格物自动识别出 {len(sheets)} 个 sheet：\n"
        + "\n".join(f"  · {s['name']}" for s in sheets[:5])
        + ("\n  · ..." if len(sheets) > 5 else ""),
        "第 5 步 · 自动识别 Sheet 结构")

    # ===== 第 6 步：结构解读 =====
    say("现在问格物：'这个 excel 是做什么的？'\n"
        "格物会分析全部 sheet 的结构、列名、数据模式，\n"
        "用麦肯锡金字塔原理给出结论先行的回答。",
        "第 6 步 · AI 结构解读（金字塔原理）")
    r = api("/api/ask", {"question": "这个excel是做什么的", "sheet": None, "scope": None}, "POST")
    say(f"AI 回答：\n{r.get('answer', '...')[:500]}",
        "第 7 步 · 结构解读结果")

    # ===== 第 8 步：数据分析 =====
    say("接下来问一个更有深度的问题：\n"
        "'从部门数据来看你能看出来哪些业务问题？'\n"
        "格物的分析引擎会先算好统计量（均值/最值/异常检测），\n"
        "然后 AI 引用这些数字做业务解读（数值纪律：AI 永不算数）。",
        "第 8 步 · AI 数据分析（引擎统计 + AI 解读）")
    r = api("/api/ask", {"question": "从部门数据来看你能看出来哪些业务问题",
                         "sheet": "指标分解", "scope": None}, "POST")
    say(f"AI 分析结果：\n{r.get('answer', '...')[:600]}",
        "第 9 步 · 业务洞察")

    # ===== 第 10 步：冻结 =====
    say("在改数之前，必须先'冻结'语义镜像。\n"
        "这是格物的安全机制：人工确认结构判定后，\n"
        "才能打开写入的总闸门。不冻结 = 只读模式。",
        "第 10 步 · 冻结语义镜像（安全闸门）")
    r = api("/api/freeze", {"edits": []}, "POST")
    say(f"冻结结果：{'成功' if r.get('ok') else '失败'}\n"
        f"已应用：{r.get('applied', 0)} 个区块\n"
        f"从现在起可以安全改数了。",
        "第 11 步 · 已冻结")

    # ===== 第 12 步：生成写计划 =====
    say("现在演示核心能力：自然语言改数。\n"
        "指令：'把销售部1月的金额改成1000'\n"
        "格物会：LLM 解析意图 → 引擎定位坐标 → 八条硬校验 → diff 预览。\n"
        "注意：AI 只选模式和抄操作数，不自己做算术。",
        "第 12 步 · 自然语言改数（LLM→引擎→校验→diff）")
    r = api("/api/plan_write", {"mode": "nl",
                                "text": "把销售部1月的金额改成1000"}, "POST")
    if r.get("ok"):
        op = r.get("op", {})
        say(f"写计划生成成功！\n"
            f"  位置：{op.get('sheet')}!r{op.get('row')}c{op.get('col')}\n"
            f"  旧值：{r.get('old')} → 新值：{op.get('value')}\n"
            f"  命中行：{json.dumps(r.get('hit', {}), ensure_ascii=False)[:100]}\n"
            f"  校验：{'通过' if r.get('validation', {}).get('ok') else '拒绝'}",
            "第 13 步 · Diff 预览")
    else:
        say(f"写计划：{r.get('clarify', '...')}",
            "第 13 步 · 写计划结果")

    # ===== 第 14 步：执行写入 =====
    if r.get("ok"):
        say("人工确认 diff 无误后，执行写入。\n"
            "写入走 Excel COM 接口（和真人打开 Excel 改格子的系统通道相同），\n"
            "自动备份到 .cache/backup/，写后校验值和结构。",
            "第 14 步 · 执行写入（COM + 自动备份）")
        r2 = api("/api/execute_write", {"op": op}, "POST")
        say(f"写入结果：{'成功' if r2.get('ok') else '失败'}\n"
            f"  引擎：{r2.get('engine', '...')}\n"
            f"  备份：{r2.get('backup', '...')}\n"
            f"  写后校验：{'通过' if r2.get('post_check', {}).get('ok') else '异常'}",
            "第 15 步 · 写入完成")

        # ===== 第 16 步：撤销 =====
        say("如果改错了？一键撤销。\n"
            "格物记录了写入前的旧值，撤销精确恢复。",
            "第 16 步 · 一键撤销")
        r3 = api("/api/undo", {}, "POST")
        say(f"撤销结果：{'成功' if r3.get('ok') else r3.get('reason', '失败')}\n"
            f"  恢复了 {r3.get('restored', 0)} 个单元格",
            "第 17 步 · 撤销完成")

    # ===== 结束 =====
    print("\n" + "=" * 56)
    print("  演示结束")
    print("=" * 56)
    print("""
  核心能力总结：
  ① 结构识别 — 自动解析 9 个 sheet 的表头/区块/汇总行
  ② 安全问答 — 数值 100% 引擎计算，AI 只做业务解读
  ③ 结构解读 — 麦肯锡金字塔原理，结论先行
  ④ 数据分析 — So What 测试 + 假设驱动
  ⑤ 安全改数 — 八条硬校验 + diff 预览 + 自动备份 + 撤销
  ⑥ API 全覆盖 — 31 个端点，可被任何 Agent 自动化
  ⑦ 本地优先 — 数据不出的机器，LLM 走自己的 key
""")


if __name__ == "__main__":
    try:
        run_demo()
    except KeyboardInterrupt:
        print("\n\n  演示中断（Ctrl+C）")
    except Exception as e:
        print(f"\n  ✗ 演示出错：{e}")
        print("  请确认格物服务已启动（http://127.0.0.1:8765）")
