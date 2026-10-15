# -*- coding: utf-8 -*-
"""LLM 设置：多协议分派 / 本地免钥 / 保存保留旧钥。"""
import io
import json
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
sys.stdout.reconfigure(encoding="utf-8")

from engine import llm

SB = os.path.join("tests", "sandbox")
os.makedirs(SB, exist_ok=True)
CFG = os.path.join(SB, "llm_test.json")

print("== ① save_config：保存与保留旧钥 ==")
if os.path.exists(CFG):
    os.remove(CFG)
llm.save_config("openai", "https://api.siliconflow.cn/v1", "sk-aaaaaaaaaaaaaaaaaaaa",
                "m1", path=CFG)
d = json.load(io.open(CFG, encoding="utf-8"))
assert d["model"] == "m1" and d["api_key"].startswith("sk-")
# 只换模型不填 key → 旧钥保留
llm.save_config("openai", "https://api.siliconflow.cn/v1", "", "m2", path=CFG)
d2 = json.load(io.open(CFG, encoding="utf-8"))
assert d2["model"] == "m2" and d2["api_key"].startswith("sk-"), "同地址改模型应保留旧钥"
# 换地址不填 key → 不保留（防把 A 家钥匙发给 B 家）
llm.save_config("anthropic", "https://api.anthropic.com", "", "claude-x", path=CFG)
d3 = json.load(io.open(CFG, encoding="utf-8"))
assert d3["api_key"] == "" and d3["protocol"] == "anthropic"
print("  ✓ 同址留钥 / 换址清钥 / 协议写入")

print("== ② 协议分派：anthropic 走 /v1/messages + x-api-key ==")
captured = {}


def fake_urlopen(req, timeout=None):
    captured["url"] = req.full_url
    captured["headers"] = dict(req.header_items())
    captured["body"] = json.loads(req.data.decode("utf-8"))

    class R:
        def read(self):
            return b""

        def __enter__(self):
            return self

        def __exit__(self, *a):
            return False
    return R()


def fake_load():
    return {"protocol": "anthropic", "base_url": "https://api.anthropic.com",
            "api_key": "sk-ant-xxx", "model": "claude-sonnet-4-5"}


# anthropic 响应结构（fake_urlopen 返回的 R 已可复用）
orig_load, orig_open, orig_json = llm.load_config, llm.urllib.request.urlopen, llm.json.load
llm.load_config = fake_load
llm.urllib.request.urlopen = fake_urlopen
llm.json.load = lambda r: {"content": [{"type": "text", "text": "在线"}]}
out = llm.chat([{"role": "system", "content": "S"}, {"role": "user", "content": "hi"}])
assert out == "在线"
assert captured["url"].endswith("/v1/messages")
assert captured["headers"].get("X-api-key") == "sk-ant-xxx"
assert captured["body"]["system"] == "S" and captured["body"]["messages"] == [
    {"role": "user", "content": "hi"}]
print("  ✓ URL/头/体均按 Anthropic 协议；system 正确抽出")

print("== ③ 协议分派：openai 走 /chat/completions + Bearer ==")
def fake_load2():
    return {"protocol": "openai", "base_url": "https://api.siliconflow.cn/v1",
            "api_key": "sk-sf", "model": "m"}
llm.load_config = fake_load2
llm.json.load = lambda r: {"choices": [{"message": {"content": "ok"}}]}
out2 = llm.chat([{"role": "user", "content": "hi"}])
assert out2 == "ok" and captured["url"].endswith("/chat/completions")
assert captured["headers"].get("Authorization") == "Bearer sk-sf"
print("  ✓ OpenAI 兼容路径不变")

print("== ④ 本地免钥：localhost 无 key 可调用 ==")
def fake_load3():
    return {"protocol": "openai", "base_url": "http://localhost:11434/v1",
            "api_key": "", "model": "qwen3:8b"}
llm.load_config = fake_load3
llm.json.load = lambda r: {"choices": [{"message": {"content": "local"}}]}
assert llm.chat([{"role": "user", "content": "hi"}]) == "local"
assert captured["headers"].get("Authorization") == "Bearer local"
print("  ✓ Ollama 场景免钥放行")

llm.load_config, llm.urllib.request.urlopen, llm.json.load = orig_load, orig_open, orig_json
print("\n全部通过 ✓")
