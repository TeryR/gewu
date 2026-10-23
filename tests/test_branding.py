# -*- coding: utf-8 -*-
"""品牌/属主指纹/双语字典完整性（前端已拆 static/ 三文件）。

① I18N 字典 zh/en 键集合一致（漏译会静默回退中文，测试兜住）；
② 页面三处属主水印注入点：HTML注释+meta 模板占位（serve 静态服务时注入）+ HTTP头；
③ branding.json 可解析，watermark 计算确定。"""
import json
import os
import re
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
sys.stdout.reconfigure(encoding="utf-8")

import serve

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
HTML = open(os.path.join(ROOT, "static", "index.html"), encoding="utf-8").read()
JS = open(os.path.join(ROOT, "static", "app.js"), encoding="utf-8").read()
PAGE = HTML + "\n" + JS  # 前端拆离前的整页视角

print("== ① I18N 字典 zh/en 键集合一致 ==")
m = re.search(r"const I18N=\{\r?\n(.*?)\r?\n\};\r?\nlet LANG", PAGE, re.S)
assert m, "I18N 字典未找到"
block = m.group(1)
# 提取两段键值（zh: {...}, "en": {...}）——按顶层键切（兼容 CRLF 行尾）
zh_m = re.search(r'"zh":\{(.*?)\r?\n \},', block, re.S)
en_m = re.search(r'"en":\{(.*?)\r?\n \}', block, re.S)
assert zh_m and en_m, "zh/en 段未找到"
keys = lambda s: set(re.findall(r'^\s*"?([a-z_0-9]+)"?:', s, re.M))
zh_keys, en_keys = keys(zh_m.group(1)), keys(en_m.group(1))
missing = zh_keys - en_keys
extra = en_keys - zh_keys - {"help_full"}   # help_full 的 zh 从 DOM 原文捕获，字典只存 en
assert not missing, f"英文版漏译: {sorted(missing)}"
assert not extra, f"英文版多出的键: {sorted(extra)}"
print(f"  ✓ 双语各 {len(zh_keys)} 键，完全对齐")

print("\n== ② 三处水印注入点（模板占位 + serve 注入 + HTTP 头）==")
assert "gewu-owner-fingerprint: pending" in HTML   # 静态模板占位符
assert 'content="Gewu/1.0 (wm:pending)"' in HTML
_serve_src = open(os.path.join(ROOT, "serve.py"), encoding="utf-8").read()
assert 'X-Gewu-Owner", BRANDING["watermark"]' in _serve_src  # HTTP 头
assert '"wm:pending", f\'wm:{BRANDING["watermark"]}\'' in _serve_src  # 静态注入
print("  ✓ HTML注释占位 / meta 占位 / serve 静态注入 / X-Gewu-Owner 响应头")

print("\n== ③ branding 加载与指纹 ==")
b = serve.BRANDING
assert b["owner"] and b["watermark"] and len(b["watermark"]) == 16, b
print(f"  属主={b['owner']}  水印={b['watermark']}  产品={b['product']}·{b['product_en']}")
cfg = json.load(open(os.path.join(ROOT, "branding.json"), encoding="utf-8"))
assert cfg["product"] == b["product"]

print("\n== ④ 修复守护：i18n 不得毁容器子元素 / 帮助面板双语 / 邮箱进关于面板 ==")
# 蒙版/浓度的 label 不得直接挂 data-i18n（textContent 覆写会删掉里面的 input）
assert '<span data-i18n="mask">' in PAGE and '<span data-i18n="opacity">' in PAGE
assert 'label class="hint" style="display:flex;gap:2px;align-items:center" data-i18n' not in PAGE
# 帮助面板挂了 help_full，且英文版有词条
assert 'data-i18n-html="help_full"' in PAGE
assert 'help_full:' in PAGE and 'What freezing means' in PAGE
# 邮箱参与指纹 + 关于面板渲染 EMAIL（serve 静态注入填充常量）
assert cfg.get("email"), 'branding.json 缺 email'
assert 'const EMAIL="";' in JS                      # 模板占位（serve 注入真实值）
assert JS.count("📧 ${EMAIL}") == 2, "关于面板中英文两处应渲染 EMAIL"
print("  ✓ 滑条/开关安全 · 帮助双语 · email=", cfg["email"])

assert 'id="aboutBox"' in PAGE and 'showAbout' in PAGE
print("  ✓ 关于面板存在")

print("\n全部通过 ✓")
