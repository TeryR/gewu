# -*- coding: utf-8 -*-
"""LLM 适配层：只做意图解析与语义定位——绝不产数值、绝不产坐标。
输出契约：严格 JSON；写指令是"find 行键 → set 列值"的语义计划，
由 resolve_write_plan 查冻结镜像解析成具体单元格，再经 writer 八条硬校验。
配置（BYOK）：环境变量 EXCEL_AGENT_LLM_BASE_URL / API_KEY / MODEL，或项目根 llm.json。
隐私纪律：上下文只送 表/列名+采样行，结构坐标不出网。"""
import json
import os
import re
import urllib.error
import urllib.request

_PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


_PLACEHOLDER_KEYS = {"sk-...", "你的key"}


def load_config():
    """顶层字段优先；顶层为空时自动认领 _常用配置 里真实填过的块（模板陷阱防御）。
    JSON 语法错误必须响亮报错——静默跳过会让用户以为配置生效了。"""
    cfg = {}
    for env_key, key in [("EXCEL_AGENT_LLM_BASE_URL", "base_url"),
                         ("EXCEL_AGENT_LLM_API_KEY", "api_key"),
                         ("EXCEL_AGENT_LLM_MODEL", "model")]:
        if os.environ.get(env_key):
            cfg[key] = os.environ[env_key]
    p = os.path.join(_PROJECT_ROOT, "llm.json")
    if os.path.exists(p):
        try:
            data = json.load(open(p, encoding="utf-8"))
        except json.JSONDecodeError as e:
            raise SystemExit(f"llm.json JSON 语法错误：第{e.lineno}行 {e.msg}")
        merged = {k: v for k, v in data.items()
                  if v and not k.startswith("_") and k not in cfg}
        profiles = data.get("_常用配置") or data.get("_三家常用配置") or {}
        if not merged.get("api_key"):
            for name, prof in profiles.items():
                if not isinstance(prof, dict):
                    continue
                key_v = str(prof.get("api_key") or "")
                if prof.get("model") and len(key_v) >= 16 and key_v not in _PLACEHOLDER_KEYS:
                    merged = {**prof, **{k: v for k, v in merged.items() if v},
                              "_profile": name}
                    break
        cfg.update({k: v for k, v in merged.items() if v})
    if cfg and not cfg.get("protocol"):
        cfg["protocol"] = "anthropic" if "anthropic.com" in (cfg.get("base_url") or "") \
            else "openai"
    return cfg or None


def llm_ready(cfg=None):
    """serve 层门槛：model 配了即算就绪（key 是否必需由 chat 按 本地/云端 判断）。"""
    cfg = cfg or load_config()
    return bool(cfg and cfg.get("model"))


def _is_local(base_url):
    h = (base_url or "").lower()
    return any(x in h for x in ("localhost", "127.0.0.1", "[::1]", "0.0.0.0"))


def save_config(protocol="openai", base_url="", api_key="", model="", path=None):
    """设置面板保存：写 llm.json。api_key 留空且未换地址时保留旧钥（只改模型的常见路径）。"""
    p = path or os.path.join(_PROJECT_ROOT, "llm.json")
    old = {}
    if os.path.exists(p):
        try:
            old = json.load(open(p, encoding="utf-8"))
        except (OSError, json.JSONDecodeError):
            old = {}
    if not api_key and old.get("base_url") == base_url and old.get("api_key"):
        api_key = old["api_key"]
    data = dict(old)                       # 保留用户手写的其他顶层键（如 _常用配置）
    data.update({"protocol": protocol, "base_url": base_url,
                 "api_key": api_key, "model": model,
                 "_note": "由格物设置面板写入；本地模型（Ollama 等）api_key 可留空"})
    json.dump(data, open(p, "w", encoding="utf-8"), ensure_ascii=False, indent=2)
    return data


def chat(messages, timeout=60, cfg=None):
    """多协议：OpenAI 兼容（云端/本地 Ollama 等）+ Anthropic 兼容，标准库零依赖。
    cfg 可显式传入（测试连接用）——不 monkeypatch 全局 loader，避免并发竞态。"""
    cfg = cfg or load_config()
    local = _is_local(cfg.get("base_url")) if cfg else False
    if not cfg or not (cfg.get("api_key") or local):
        raise SystemExit("LLM 未配置：点页眉 ⚙️ 打开设置，或编辑 llm.json"
                         "（本地模型 api_key 可留空）")
    model = cfg.get("model")
    if not model:
        # 不猜默认模型名——猜错会得到含混的 API 报错，明确要求配置更诚实
        raise SystemExit("LLM 未配置 model：在 llm.json 或 EXCEL_AGENT_LLM_MODEL 指定"
                         "（如 glm-4.6 / deepseek-chat / qwen-plus）")
    base = (cfg.get("base_url") or "https://open.bigmodel.cn/api/paas/v4").rstrip("/")
    protocol = (cfg.get("protocol") or "openai").lower()
    if protocol == "anthropic":
        system = "\n".join(m["content"] for m in messages if m["role"] == "system")
        msgs = [m for m in messages if m["role"] != "system"]
        req = urllib.request.Request(
            base + "/v1/messages",
            data=json.dumps({"model": model, "max_tokens": 2048, "temperature": 0,
                             "system": system or None, "messages": msgs}).encode("utf-8"),
            headers={"Content-Type": "application/json",
                     "x-api-key": cfg.get("api_key") or "",
                     "anthropic-version": "2023-06-01"})
    else:
        req = urllib.request.Request(
            base + "/chat/completions",
            data=json.dumps({"model": model,
                             "messages": messages, "temperature": 0}).encode("utf-8"),
            headers={"Content-Type": "application/json",
                     "Authorization": f"Bearer {cfg.get('api_key') or 'local'}"})
    try:
        with urllib.request.urlopen(req, timeout=timeout) as r:
            out = json.load(r)
    except urllib.error.HTTPError as e:
        raise SystemExit(f"LLM API HTTP {e.code}（检查 api_key/model/base_url；"
                         f"网关返回：{e.read()[:120].decode('utf-8', 'ignore')}）")
    except urllib.error.URLError as e:
        raise SystemExit(f"LLM 网络错误：{e.reason}（检查 base_url 或网络）")
    except (TimeoutError, json.JSONDecodeError, KeyError, IndexError) as e:
        # 超时/响应非 JSON/缺字段：不接住会砸掉 HTTP 线程（表现为"连接被重置"）
        raise SystemExit(f"LLM 响应异常：{type(e).__name__} {e}")
    if protocol == "anthropic":
        try:
            return out["content"][0]["text"]
        except (KeyError, IndexError, TypeError) as e:
            raise SystemExit(f"Anthropic 响应异常：{type(e).__name__} {e}")
    return out["choices"][0]["message"]["content"]


def ask_json(system, user, retries=1, fn=None, history=None):
    """强制 JSON：剥代码围栏、截取首尾花括号、失败重试一次。fn 注入点（测试 mock 用）。"""
    call = fn or chat  # 调用时解析，支持 monkeypatch llm.chat
    for attempt in range(retries + 1):
        msgs = [{"role": "system", "content": system},
                *_hist_messages(history),
                {"role": "user",
                 "content": user + ("\n（重试）只输出 JSON，不要任何其他文字。"
                                    if attempt else "\n只输出 JSON。")}]
        txt = call(msgs)
        s = txt.strip()
        for fence in ("```json", "```"):
            if s.startswith(fence):
                s = s[len(fence):]
            if s.endswith("```"):
                s = s[:-3]
        s = s.strip()
        try:
            return json.loads(s)
        except json.JSONDecodeError:
            start, end = s.find("{"), s.rfind("}")
            if start >= 0 and end > start:
                try:
                    return json.loads(s[start:end + 1])  # 模型带了解说文字时截取 JSON 体
                except json.JSONDecodeError:
                    pass
    raise SystemExit("LLM 输出无法解析为 JSON（已重试一次）")


QUERY_SYSTEM = """你是表格查询意图解析器。把用户问题转为查询计划 JSON：
{"op":"sum|avg|count|max|min|lookup|ties|describe|analyze","column":"列名（必须取自上下文真实列）",
 "block_hint":"sheet名或表名关键词","filters":{"列名":"值"}}
选列与选表规则：
- 列名在上下文里重复出现时（如多个分组下同名“金额”），column 必须用完整路径（含组头，如“销售部/金额”），
  必要时配合 block_hint/filters 锁定是哪一组。
- sheet 名相近时（如多个“汇总表”变体），block_hint 要带上问题里的区分词（规格/型号/对象，如“汇总表规格TR-4020”）。
- describe 用于理解型问题（这个表是做什么的/有哪些字段/某字段什么口径/备注说了什么），
  此时只需 block_hint 不需要 column。
- analyze 用于数据解读型问题（有什么评价/哪个更好/趋势如何/数据说明什么/有无异常），
  此时只需 block_hint 不需要 column。
- 用户问整个文件/全簿/所有表（"这个文件是做什么的""各表什么关系"）时，block_hint 用 "*"
  （describe 讲结构与关系；analyze 评价整体设计/组织合理性，如"这个文件设计合理吗"）。
- 上下文带「当前页」时，"这个表格/当前表格"指它，block_hint 用当前页 sheet 名。
- 对话历史里出现过的表/列，用户的"它/这个/为什么这么算"等回指按历史理解：
  默认延续历史中最近讨论的那张表/列（describe/analyze 时 block_hint 直接用它），不要追问；仅在历史无线索时才 need_clarify。
只输出 JSON。无法确定时输出 {"op":"unknown"}；
知道缺什么信息就带上 need_clarify，如 {"op":"unknown","need_clarify":"要看哪个规格的汇总表？"}。"""

ANALYZE_SYSTEM = """你是一位受过麦肯锡方法论训练的数据分析师。基于提供的表格数据回答用户问题。

IMPORTANT: You MUST respond in the SAME language as the user's question.
If the user asks in English, respond entirely in English. If in Chinese, respond in Chinese.
Do NOT mix languages. Column names may be in any language — quote as-is.

思考方法（内部执行，不输出思考过程）：
1. 假设驱动：先形成假设（"数据可能说明X"），再用上下文中的统计量证明或推翻
2. MECE 检查：支撑论点是否相互独立、完全穷尽？有无遗漏维度？
3. So What 测试：每个发现问自己"那又怎样？对决策有什么影响？"——答不上来就砍掉
4. 帕累托：抓影响最大的 2~3 个发现，不平均用力

输出结构（金字塔原理：结论先行）：
**[一句话结论——回答"所以呢？"]**
- [发现1：具体数字 + 业务含义（So What）]
- [发现2：同上]
- [发现3：同上]
> [行动建议：基于上述发现，下一步该做什么]

质量标准：
- 每条发现必须同时包含“数字”和“含义”：“差旅费 4.8 万 vs 运输费 1.9 万”是数字，
  “差旅费占总费用六成、集中在 1 月”才是含义
- 结论回答“所以呢”：不说“存在差异”，说“应优先管控差旅费报销节奏”
- 行动建议回答"那怎么办"：引用具体列名和数字，给可执行的下一步
- 异常值是机会：**负值**、**极端值**、**突变** 通常藏着重磅发现，优先指出
- 最多 3 个发现 + 1 条建议。宁少勿滥，每条必须有 So What。

纪律：不执行算术——统计量由引擎算好在上下文里，你只引用；
数字必须与上下文一致，禁止计算或换算；没有的不编造。
直接输出回答文本，不要 JSON。"""


def analyze_block(question, ctx, history=None):
    """数据解读（op=analyze）：引擎算统计量，LLM 只做业务解读，不碰算术。"""
    return chat([{"role": "system", "content": ANALYZE_SYSTEM},
                 *_hist_messages(history),
                 {"role": "user",
                  "content": f"表格数据：{json.dumps(ctx, ensure_ascii=False, default=str)}\n"
                             f"用户问题：{question}"}])

DESCRIBE_SYSTEM = """你是一位业务分析顾问，擅长用 SCQA 框架快速理解表格。

IMPORTANT: You MUST respond in the SAME language as the user's question.
If the user asks in English, respond entirely in English. If in Chinese, respond in Chinese.
Do NOT mix languages. Column names from the table may be in any language — quote them as-is
(e.g. `部门` or `Department`), but your analysis text must match the user's language.

思考方法（内部执行，不输出）：
- Situation：这张表/文件在什么业务情境下存在？
- Complication：数据里有什么值得注意的（异常、差异、模式）？
- Question：使用者真正想了解什么？
- Answer：用一句话回答，然后给支撑论据

用户问题含多个子问题时，按"洞察优先"排序：
先给业务发现和数据问题（高价值），再给结构描述（低价值，压缩到 1~2 条）。

输出结构（金字塔原理：结论先行）：
**[一句话：这张表是干什么的、为什么存在]**
- [关键发现1：具体数字/列名 + 业务含义]
- [关键发现2：同上（如有）]
> [使用建议：看完这些数据后该注意什么/做什么]

质量标准：
- 结论具体到可行动：说“这是对比 3 个部门费率的费用归并表”（可指导下一步），不说“这是测试数据表”
- 每条发现必须通过 So What 测试："这对使用者意味着什么？"
- 列名用反引号：-\`成交套数\` 为月度口径
- 百分比列按百分比说（-0.11 → -11%）
- 公式问题：引用原文，解释引用了什么、做了什么运算
- 合计/最值类问题：说"请用「xx列合计」问法"

全簿模式（含 sheets 列表）时：
- **结论**：文件的业务用途 + 表间关系（按什么维度展开、各表分工）
- 发现列表：跨表规律（如同名列口径差异、数据冗余/矛盾）
- 不超过 6 行

纪律：不执行算术；数字只照抄上下文；没有的不编造；不知道明说。
直接输出回答文本，不要 JSON。"""


def describe_block(question, ctx, history=None):
    """理解型问答（op=describe）：LLM 只做描述，不碰数值计算。
    可测试性：测试里直接替换 llm.chat。"""
    return chat([{"role": "system", "content": DESCRIBE_SYSTEM},
                 *_hist_messages(history),
                 {"role": "user",
                  "content": f"表格上下文：{json.dumps(ctx, ensure_ascii=False, default=str)}\n"
                             f"用户问题：{question}"}])


def _hist_messages(history):
    """多轮历史（[{role,content}]）转 chat 消息，截最近 6 条控 token。"""
    return [dict(m) for m in (history or [])][-6:]


def query_intent(question, context, history=None, fn=None):
    return ask_json(QUERY_SYSTEM,
                    f"表格上下文：{json.dumps(context, ensure_ascii=False)}\n用户问题：{question}",
                    fn=fn, history=history)


WRITE_SYSTEM = """你是 Excel 修改指令解析器。把指令转为语义写计划 JSON：
{"find":{"sheet":"sheet名关键词","keys":{"列名":"值", ...可多键交集}},
 "set":{"column":"目标列名","mode":"set|add|sub|mul|div|add_pct|sub_pct","value":"操作数"}}
铁律（数值纪律）：
- 你不负责任何算术。指令含加减乘除或百分比（"加3""涨10%"）时，只选 mode 并抄写操作数，
  由引擎用旧值计算：add/sub/mul/div 对操作数运算，add_pct/sub_pct 的操作数是百分数（10 表示 10%）。
- 纯赋值（"改成7""涨到10"）用 mode=set，value 必须与指令中的数字逐字一致，禁止换算或改写。
- keys 用足够唯一定位一行的键（如 项目+月份）；列名必须取自上下文，重名列用完整路径；
  不确定就输出 {"need_clarify":"具体疑问"}。禁止输出坐标/行列号/计算结果。
- 指令波及多行（"都改成""全部加2"）或多张表时不要猜，输出 {"need_clarify":"要改哪一行/哪张表？"}。
示例（列名以实际上下文为准）：
“把1月差旅费的金额改成20000” → {"find":{"sheet":"销售部","keys":{"月份":"1","费用项目":"差旅费"}},"set":{"column":"金额（元）","mode":"set","value":"20000"}}
“2月运输费金额加3000” → {"find":{"sheet":"销售部","keys":{"月份":"2","费用项目":"运输费"}},"set":{"column":"金额（元）","mode":"add","value":"3000"}}
“1月差旅费金额涨10%” → {"find":{"sheet":"销售部","keys":{"月份":"1","费用项目":"差旅费"}},"set":{"column":"金额（元）","mode":"add_pct","value":"10"}}
“1月差旅费金额涨到20000” → {"find":{"sheet":"销售部","keys":{"月份":"1","费用项目":"差旅费"}},"set":{"column":"金额（元）","mode":"set","value":"20000"}}
“所有费用都改成5000” → {"need_clarify":"指令涉及多行，要改哪一行？"}"""


def write_intent(instruction, context, fn=None):
    return ask_json(WRITE_SYSTEM,
                    f"冻结镜像上下文：{json.dumps(context, ensure_ascii=False)}\n指令：{instruction}",
                    fn=fn)


# ---------------- 语义计划 → 具体单元格（纯代码解析，LLM 无权决定坐标） ----------------

def _col_of(header, needle):
    for node in header.get("col_tree") or []:
        if node.get("path"):
            p = "/".join(node["path"])
            if needle == p or needle == p.split("/")[-1] or needle in p:
                return node["col"], p
    return None, None


def _key_match(rec_value, want) -> bool:
    """行键精确匹配：禁止子串——"1月" 绝不能命中 "11月"。数值按值比较。"""
    if isinstance(rec_value, (int, float)) and isinstance(want, (int, float)) \
            and not isinstance(rec_value, bool) and not isinstance(want, bool):
        return float(rec_value) == float(want)
    return str(rec_value).strip() == str(want).strip()


# ---------------- 数值纪律：算术全在引擎，模型只选操作符抄操作数 ----------------

ARITH_MODES = ("set", "add", "sub", "mul", "div", "add_pct", "sub_pct")


def apply_arith(old, mode, operand):
    """引擎侧执行写操作：相对运算基于旧值计算，模型零算术。
    LLM 可能把操作数输出成字符串（"3"），算术模式统一数值化。"""
    if mode not in ARITH_MODES:
        raise SystemExit(f"未知写操作 mode「{mode}」（可选 {'/'.join(ARITH_MODES)}）")
    if mode == "set":
        return operand
    if isinstance(operand, str):
        try:
            operand = float(operand.strip().rstrip("%"))
        except ValueError:
            raise SystemExit(f"相对运算的操作数「{operand}」不是数字")
        operand = int(operand) if operand == int(operand) else operand
    if not isinstance(operand, (int, float)) or isinstance(operand, bool):
        raise SystemExit(f"相对运算的操作数「{operand}」不是数字")
    if not isinstance(old, (int, float)) or isinstance(old, bool):
        raise SystemExit("相对运算（加减乘除/百分比）要求目标格已有数值，"
                         f"当前旧值为 {old!r}；纯赋值请用 mode=set")
    if mode in ("add_pct", "sub_pct"):
        return old * (1 + (operand if mode == "add_pct" else -operand) / 100)
    if mode == "div" and operand == 0:
        raise SystemExit("除数不能为 0")
    return {"add": old + operand, "sub": old - operand,
            "mul": old * operand, "div": old / operand}[mode]


_NUM_PAT = re.compile(r"[-+]?[\d,]+(?:\.\d+)?")


def extract_numbers(text):
    out = []
    for m in _NUM_PAT.findall(text or ""):
        try:
            out.append(float(m.replace(",", "")))
        except ValueError:
            pass
    return out


def value_echo_warning(text, operand, new_value):
    """数字回显守卫：指令里的数字既不在操作数也不在结果里 → 模型可能自行计算/抄错。
    操作数可能是字符串数字（LLM 常把 3 输出成 "3"），比对前先数值化。"""
    nums = extract_numbers(text)
    if not nums:
        return None

    def close(a):
        if isinstance(a, bool) or a is None:
            return False
        if isinstance(a, str):
            try:
                a = float(a.rstrip("%").replace(",", ""))
            except ValueError:
                return False
        if not isinstance(a, (int, float)):
            return False
        return any(abs(a - n) <= max(abs(n) * 1e-6, 1e-9) for n in nums)

    if close(operand) or close(new_value):
        return None
    shown = [n if n == int(n) else n for n in nums[:5]]
    return (f"⚠ 指令中的数字 {shown} 与写入值 {new_value} 都对不上——"
            f"模型可能自行计算或抄错了数字，请人工核对")


def magnitude_warning(new_value, col_values, lo=1e-3, hi=1e3):
    """量级守卫：写入值与该列常见量级差 3 个数量级以上 → 疑似万↔元单位换算错误。"""
    if not isinstance(new_value, (int, float)) or isinstance(new_value, bool) or new_value == 0:
        return None
    vals = sorted(abs(v) for v in (col_values or [])
                  if isinstance(v, (int, float)) and not isinstance(v, bool) and v != 0)
    if len(vals) < 5:
        return None
    med = vals[len(vals) // 2]
    ratio = abs(new_value) / med
    if lo <= ratio <= hi:
        return None
    return (f"⚠ 写入值 {new_value:,.6g} 与该列常见量级（中位数 ≈{med:,.6g}）相差 "
            f"{ratio:,.4g} 倍，疑似单位换算（万↔元）或量纲错误")


def column_values(mirror, op):
    """取目标列的明细值（供量级守卫），只读语义镜像。"""
    for b in mirror.blocks(sheet=op["sheet"], tables_only=True):
        if b["r1"] <= op["row"] <= b["r2"]:
            h = json.loads(b["header_json"] or "{}")
            path = next(("/".join(n["path"]) for n in (h.get("col_tree") or [])
                         if n.get("col") == op["col"]), None)
            if path:
                return [r.get(path) for r in mirror._records(b) if not r.get("_summary")]
    return []


def resolve_write_plan(plan, mirror):
    """find(行键，可多键交集) → set(列) 解析为 {sheet,row,col,value}。
    精确匹配优先；全部 sheet 无命中才回退子串匹配（并标注）；多行命中报歧义。"""
    if "need_clarify" in plan:
        return {"ok": False, "clarify": plan["need_clarify"]}
    find, set_ = plan.get("find") or {}, plan.get("set") or {}
    keys = find.get("keys") or ({find["column"]: find["equals"]}
                                if find.get("column") else {})
    if not keys:
        return {"ok": False, "clarify": "缺少行键（find.keys）"}
    matched_sheets = [s for s in {b["sheet_name"] for b in mirror.blocks(tables_only=True)}
                      if find.get("sheet", "") in s]
    if not matched_sheets:
        return {"ok": False, "clarify": f"找不到 sheet「{find.get('sheet')}」"}

    def scan(matcher):
        cands = []
        for sheet in matched_sheets:
            for b in mirror.blocks(sheet=sheet, tables_only=True):
                header = json.loads(b["header_json"] or "{}")
                key_paths = {}
                for k in keys:
                    col, path = _col_of(header, k)
                    if col is not None:
                        key_paths[k] = path
                if not key_paths:
                    continue
                scol, spath = _col_of(header, set_.get("column", ""))
                if scol is None:
                    continue
                for rec in mirror._records(b):
                    if not all(matcher(rec.get(p, ""), keys[k]) for k, p in key_paths.items()):
                        continue
                    via = " & ".join(f"「{p}」={rec.get(p)}" for p in key_paths.values())
                    cands.append({"sheet": b["sheet_name"], "row": rec["_r"], "col": scol,
                                  "mode": set_.get("mode", "set"),
                                  "operand": set_.get("value"),
                                  "value": set_.get("value"),
                                  "via": f"{via} → 「{spath}」",
                                  "summary_row": bool(rec.get("_summary"))})
        return cands

    cands = scan(_key_match)
    if not cands:
        # 子串只做近似提示、绝不自动解析——"2月" 子串命中 "12月" 的教训：
        # 近似候选必须回问用户确认，写入只认精确唯一命中
        near = scan(lambda rv, w: str(w) in str(rv))
        hint = ""
        if near:
            hint = f"（近似命中 {len(near)} 行，如：{'；'.join(c['via'] for c in near[:2])}）"
        return {"ok": False,
                "clarify": "行键无命中：" + json.dumps(keys, ensure_ascii=False) + hint}
    if len(cands) > 1:
        return {"ok": False, "ambiguous": True,
                "clarify": f"命中 {len(cands)} 行，请补充行键使唯一（如月份/部门）",
                "candidates": cands[:8]}
    op = cands[0]
    # LLM 常把数字输出成字符串（"12345"）：规范数字串归一为数值，
    # 消除"向数值列写入文本"的误警告；"007"/"12.50" 等非规范串保持原样（可能是编号）
    v = op.get("value")
    if isinstance(v, str):
        t = v.strip()
        for cast in (int, float):
            try:
                if str(cast(t)) == t:
                    op["value"] = op["operand"] = cast(t)
                break
            except ValueError:
                continue
    return {"ok": True, "op": op}


def mirror_context(mirror, sample_rows=3, max_blocks=40):
    """上下文纪律：sheet/区块/列名 + 少量采样行；坐标与全量数据不出网。
    区块数封顶防 token 爆炸（大工作簿优先覆盖前面的表）。
    口径词典置顶注入——同名不同口径的列（如月/年套数）靠它消歧。"""
    notes = [{"sheet": g["sheet_name"], "column": g["col_path"] or g["term"],
              "口径": g["meaning"]} for g in mirror.glossary()]
    ctx = []
    for b in mirror.blocks(tables_only=True)[:max_blocks]:
        h = json.loads(b["header_json"] or "{}")
        cols = ["/".join(n["path"]) for n in (h.get("col_tree") or []) if n.get("path")]
        recs = mirror._records(b)
        sample = [{k: v for k, v in r.items() if not k.startswith("_") and v not in (None, "")}
                  for r in recs[:sample_rows] if not r.get("_summary")]
        ctx.append({"sheet": b["sheet_name"], "block": b["block_id"],
                    "title": h.get("title_text"), "columns": cols[:40], "sample": sample})
    rules = mirror.biz_rules() if hasattr(mirror, "biz_rules") else []
    out = {"tables": ctx}
    if notes:
        out["field_notes"] = notes
    if rules:
        out["业务规则"] = [{"sheet": r["sheet_name"], "column": r["col_path"],
                            "范围": [r["rmin"], r["rmax"]], "severity": r["severity"]}
                           for r in rules[:20]]
    return out


if __name__ == "__main__":
    import sys
    sys.stdout.reconfigure(encoding="utf-8")
    q = " ".join(sys.argv[1:]) or "配置检查"
    if q == "配置检查":
        cfg = load_config()
        print("LLM 已配置:", bool(cfg and cfg.get("api_key")),
              "| model:", (cfg or {}).get("model", "-"),
              "| base:", (cfg or {}).get("base_url", "默认(GLM)"))
    else:
        print(chat([{"role": "user", "content": q}]))
