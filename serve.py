# -*- coding: utf-8 -*-
"""确认页 Web 原型（冻结层）：底图渲染 + 彩色蒙版区块 + 点击审阅 + 冻结写镜像库。
用法：python serve.py <xlsx> [端口]   然后浏览器打开打印出的地址。
纯标准库实现；产品化阶段此页迁入 Univer+Tauri 壳，冻结写库协议保持不变。"""
import json
import os
import sqlite3
import sys
import threading
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from urllib.parse import urlparse, parse_qs

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
sys.stdout.reconfigure(encoding="utf-8")

from engine.extract import extract_workbook, _has_value
from engine.headers import build_header_tree
from engine.mirror import build_mirror, freeze, DEFAULT_DB
from engine.query import Mirror
from engine import llm as llm_mod
from engine import writer as writer_mod
from engine.segment import segment_sheet

PORT_DEFAULT = 8765
MAX_R, MAX_C = 400, 48  # 渲染上限：区块并集 + 余量

TYPE_STYLE = {
    "table": ("#2f6fdb", "🟦 数据表"),
    "note": ("#d9a406", "🟨 文本/批注"),
    "mixed": ("#d43b3b", "🟥 混合待复核"),
    "ignore": ("#888888", "⬜ 忽略"),
}

STATE = {"wb_path": None, "wid": None, "db": DEFAULT_DB, "boot": None, "copy": False}
_WB_CACHE = {}      # 进程内缓存：切 Sheet 不重复全量解析
_SHEET_CACHE = {}


def _load_branding():
    """品牌与属主指纹：产品名/副题/属主。owner 留空则取系统用户名。
    隐藏指纹 = sha256(product+owner+salt) 截断，随页面注释/meta/HTTP 头三处埋入。"""
    import getpass
    import hashlib
    try:
        b = json.load(open(os.path.join(os.path.dirname(os.path.abspath(__file__)),
                                        "branding.json"), encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        b = {}
    owner = (b.get("owner") or getpass.getuser() or "owner").strip()
    b["owner"] = owner
    email = (b.get("email") or "").strip()
    raw = f'{b.get("product_en","Gewu")}|{owner}|{email}|{b.get("watermark_salt","gewu")}'
    b["watermark"] = hashlib.sha256(raw.encode("utf-8")).hexdigest()[:16]
    return b


BRANDING = _load_branding()

API_DOCS = {
    "name": "格物 Gewu · 复杂表格理解工作台 API",
    "version": "1.0",
    "description": "本地 Excel 解析与安全改数服务。所有功能均通过 HTTP JSON API 暴露，"
                   "Web UI 只是这些 API 的消费者。外部 Agent 可直接调用完成全流程。",
    "base_url": "http://127.0.0.1:8765",
    "principles": {
        "数值纪律": "所有数值计算 100% 由引擎执行，LLM 只做意图解析和业务解读",
        "写入安全": "必须先冻结（/api/freeze）才能改数；写入前有八条硬校验+diff预览+自动备份+撤销",
        "本地优先": "全部数据在本机，通信走 127.0.0.1 回环"
    },
    "workflow": [
        "1. GET /api/boot → 确认引擎就绪",
        "2. POST /api/open {path} → 打开 Excel 文件",
        "3. GET /api/sheets → 查看有哪些 sheet",
        "4. GET /api/state?sheet=X → 获取某 sheet 的区块/蒙版/单元格数据",
        "5. POST /api/freeze {edits:[]} → 人工确认结构后冻结（打开改数闸门）",
        "6. POST /api/ask {question} → 中文问答（数值/查行/解读/分析四类）",
        "7. POST /api/plan_write {mode:'nl', text:'把X改成7'} → 生成写计划（含diff预览）",
        "8. POST /api/execute_write {op} → 执行写入（自动备份）",
        "9. POST /api/undo {} → 撤销最近一批写入",
        "10. GET /api/export → 下载改后文件"
    ],
    "endpoints": {
        "GET /api/boot": {
            "desc": "引擎启动状态",
            "returns": {"ready": "bool", "phase": "str", "error": "str|null", "fast": "bool"}
        },
        "GET /api/status": {
            "desc": "当前工作簿状态",
            "returns": {"status": "parsed|frozen", "frozen_at": "str|null",
                        "copy": "bool", "change": "object|null"}
        },
        "POST /api/open": {
            "desc": "打开 Excel 文件（路径方式，写回原文件）",
            "body": {"path": "绝对路径.xlsx"},
            "returns": {"ok": "bool"}
        },
        "POST /api/upload": {
            "desc": "上传文件副本（分析用，写入只动副本）",
            "body": "原始文件字节（Content-Type 无要求，?name=文件名.xlsx）",
            "returns": {"ok": "bool", "file": "str"}
        },
        "GET /api/sheets": {
            "desc": "列出当前工作簿全部 sheet",
            "returns": [{"name": "str", "visible": "bool"}]
        },
        "GET /api/state": {
            "desc": "某 sheet 的区块判定+单元格数据+蒙版信息",
            "query": "sheet=sheet名",
            "returns": {"blocks": "list", "cells": "dict", "merges": "list"}
        },
        "POST /api/freeze": {
            "desc": "冻结语义镜像（打开改数总闸门）",
            "body": {"edits": [{"block_db_id": "int", "user_type": "str", "readonly": "int"}]},
            "returns": {"ok": "bool", "applied": "int", "missed": "int"}
        },
        "POST /api/ask": {
            "desc": "中文问答（支持四类：数值计算/查行/结构解读/数据分析）",
            "body": {"question": "str", "sheet": "str|null", "scope": "list|null",
                     "session_id": "int|null"},
            "returns": {"answer": "str", "source": "llm|rules|rules-fallback",
                        "session_id": "int"}
        },
        "POST /api/plan_write": {
            "desc": "生成写计划（含diff预览，不执行）",
            "body": {"mode": "nl|plan", "text": "自然语言或JSON计划"},
            "returns": {"ok": "bool", "op": "object", "old": "any", "hit": "dict",
                        "deps": "list", "validation": "object"}
        },
        "POST /api/execute_write": {
            "desc": "执行已校验的写计划（COM写入，自动备份）",
            "body": {"op": "plan_write 返回的 op 对象"},
            "returns": {"ok": "bool", "engine": "str", "backup": "str",
                        "post_check": "object", "deps_changed": "list"}
        },
        "POST /api/undo": {
            "desc": "撤销最近一批写入",
            "body": {},
            "returns": {"ok": "bool", "restored": "int", "engine": "str"}
        },
        "GET /api/export": {
            "desc": "下载当前工作簿文件（改后成品）",
            "returns": "xlsx 文件字节流"
        },
        "GET /api/write_log": {
            "desc": "最近写入流水",
            "returns": [{"ts": "str", "sheet": "str", "row": "int", "col": "int",
                         "old": "str", "new": "str", "undone": "bool"}]
        },
        "GET /api/rules": {
            "desc": "业务规则列表（值域/格式）",
            "query": "sheet=sheet名（可选过滤）",
            "returns": [{"sheet_name": "str", "col_path": "str", "rmin": "num|null",
                         "rmax": "num|null", "pattern": "str|null", "severity": "block|warn"}]
        },
        "POST /api/rules_save": {
            "desc": "保存业务规则（覆盖式）",
            "body": {"sheet": "str", "column": "str", "min": "num|null",
                     "max": "num|null", "pattern": "str|null", "severity": "block|warn"},
            "returns": {"ok": "bool"}
        },
        "POST /api/rules_check": {
            "desc": "全表规则体检（对现有数据跑一遍规则）",
            "returns": [{"sheet": "str", "row": "int", "column": "str",
                         "value": "any", "问题": "str", "severity": "str"}]
        },
        "GET /api/glossary": {
            "desc": "字段口径词典",
            "returns": [{"term": "str", "meaning": "str", "sheet_name": "str"}]
        },
        "POST /api/glossary_save": {
            "desc": "保存字段口径",
            "body": {"column": "str", "meaning": "str", "sheet": "str"},
            "returns": {"ok": "bool"}
        },
        "GET /api/sessions": {
            "desc": "当前文件的会话列表",
            "returns": [{"id": "int", "title": "str", "scope": "list|null",
                         "updated_at": "str"}]
        },
        "POST /api/ask → session": "会话通过 /api/ask 自动创建/续写",
        "GET /api/search_history": {
            "desc": "跨文件历史搜索",
            "query": "q=关键词",
            "returns": [{"session_id": "int", "title": "str", "file": "str", "snippet": "str"}]
        },
        "GET /api/llm_config": {
            "desc": "当前 LLM 配置（脱敏）",
            "returns": {"protocol": "str", "base_url": "str", "model": "str", "has_key": "bool"}
        },
        "POST /api/llm_test": {
            "desc": "测试 LLM 连接",
            "body": {"protocol": "str", "base_url": "str", "model": "str", "api_key": "str"},
            "returns": {"ok": "bool", "ms": "int", "reply": "str"}
        },
        "GET /api/app_settings": {
            "desc": "应用设置（托盘等）",
            "returns": {"close_to_tray": "bool"}
        }
    },
    "agent_quickstart": {
        "open_and_analyze": [
            "POST /api/open {path: 'C:/path/file.xlsx'}",
            "GET /api/sheets",
            "POST /api/ask {question: '这个表格是做什么的'}",
            "POST /api/ask {question: '分析一下数据质量'}"
        ],
        "safe_write": [
            "POST /api/freeze {edits: []}",
            "POST /api/plan_write {mode: 'nl', text: '把1月云澜居的成交套数改成7'}",
            "POST /api/execute_write {op: <上一步返回的op>}",
            "POST /api/undo {}  ← 如果需要撤销"
        ],
        "data_quality": [
            "POST /api/rules_save {sheet: '明细', column: '数量', min: 0, max: 100}",
            "POST /api/rules_check"
        ]
    }
}



def S_SHEET(state):
    return state.get("sheet") or ""


def _get_wb(path):
    if path not in _WB_CACHE:
        _WB_CACHE[path] = extract_workbook(path)
    return _WB_CACHE[path]


def sheet_state(sheet_name):
    """渲染数据：区块并集范围内的 格值/合并/填充 + 镜像区块（带 db 主键）。
    缓存键含 wid：镜像重建会换 db 主键，旧缓存里的 db_id 必须失效。"""
    cache_key = (STATE["wb_path"], STATE["wid"], sheet_name)
    if cache_key in _SHEET_CACHE:
        return _SHEET_CACHE[cache_key]
    wb = _get_wb(STATE["wb_path"])
    s = next(x for x in wb["sheets"] if x["name"] == sheet_name)
    seg = segment_sheet(s)
    con = sqlite3.connect(STATE["db"])
    db_ids = {(sname, bid): sid for sid, sname, bid in
              con.execute("SELECT id, sheet_name, block_id FROM blocks WHERE workbook_id=?",
                          (STATE["wid"],))}
    # 行内备注挂接（note 引擎id → {表引擎id, 行, 文本}）——侧栏展示与解除入口用
    attaches = {nb: {"table_block_id": tb, "row_no": r, "text": t} for nb, tb, r, t in
                con.execute("SELECT note_block_id, table_block_id, row_no, text "
                            "FROM row_notes WHERE workbook_id=? AND sheet_name=?",
                            (STATE["wid"], sheet_name))}
    con.close()
    blocks = []
    for b in seg["blocks"]:
        if b["type"] == "table":
            b["header"] = build_header_tree(s, b)
        blocks.append({k: b.get(k) for k in
                       ("id", "bbox", "type", "type_label", "score", "signals",
                        "annotation", "summary_tail", "header")}
                      | {"db_id": db_ids.get((sheet_name, b["id"])),
                         "attach": attaches.get(b["id"])})
    if blocks:
        r1 = max(1, min(b["bbox"][0] for b in blocks) - 2)
        r2 = min(MAX_R, max(b["bbox"][2] for b in blocks) + 2)
        c1 = max(1, min(b["bbox"][1] for b in blocks) - 1)
        c2 = min(MAX_C, max(b["bbox"][3] for b in blocks) + 1)
    else:
        r1, r2, c1, c2 = 1, 30, 1, 20
    cells = {}
    for (r, c), info in s["cells"].items():
        if r1 <= r <= r2 and c1 <= c <= c2:
            v = info["cv"] if info["cv"] is not None else info["v"]
            if _has_value(v):
                cells[f"{r},{c}"] = {
                    "v": (str(v)[:60] if not isinstance(v, (int, float))
                          else round(v, 4) if isinstance(v, float) else v),
                    "f": info["f"], "fm": info["fm"],
                    "pct": bool(info.get("nf") and "%" in info["nf"]),
                }
    merges = [m for m in s["merged"]
              if not (m[2] < r1 or m[0] > r2 or m[3] < c1 or m[1] > c2)]
    state = {
        "sheet": sheet_name, "range": [r1, c1, r2, c2],
        "cells": cells, "merges": merges,
        "hidden_rows": [r for r in s["hidden_rows"] if r1 <= r <= r2],
        "hidden_cols": s["hidden_cols"],
        "blocks": blocks,
        "workbook": {"id": STATE["wid"], "path": os.path.basename(STATE["wb_path"])},
    }
    _SHEET_CACHE[cache_key] = state
    return state


class Handler(BaseHTTPRequestHandler):
    def log_message(self, *a):
        pass

    def _busy(self):
        """启动门控：镜像未就绪时镜像类端点一律 503+boot 状态（前端轮询 /api/boot）。"""
        if not STATE.get("ready"):
            self._json({"ok": False, "reason": "引擎解析中",
                        "boot": {"ready": False, "phase": STATE.get("boot_phase"),
                                 "error": STATE.get("boot_error")}}, 503)
            return True
        return False

    def _json(self, obj, code=200):
        body = json.dumps(obj, ensure_ascii=False, default=str).encode("utf-8")
        self.send_response(code)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.send_header("X-Gewu-Owner", BRANDING["watermark"])  # 隐藏属主指纹
        self.end_headers()
        self.wfile.write(body)

    def do_GET(self):
        u = urlparse(self.path)
        if u.path == "/":
            body = (PAGE.replace("__WATERMARK__", BRANDING["watermark"])
                        .replace("__OWNER__", BRANDING["owner"])
                        .replace("__EMAIL__", BRANDING.get("email", ""))
                        .replace("__WATERMARK__", BRANDING["watermark"])).encode("utf-8")
            self.send_response(200)
            self.send_header("Content-Type", "text/html; charset=utf-8")
            self.send_header("X-Gewu-Owner", BRANDING["watermark"])
            self.send_header("Content-Length", str(len(body)))
            self.end_headers()
            self.wfile.write(body)
        elif u.path == "/api/sheets":
            if self._busy(): return
            wb = _get_wb(STATE["wb_path"])
            self._json([{"name": s["name"], "visible": s["visible"]} for s in wb["sheets"]])
        elif u.path == "/api/glossary":
            if self._busy(): return
            q = parse_qs(u.query)
            sheet = (q.get("sheet") or [None])[0]
            rows = self._mirror().glossary()
            if sheet:
                rows = [r for r in rows if not r["sheet_name"] or sheet in str(r["sheet_name"])]
            self._json(rows)
        elif u.path == "/api/state":
            if self._busy(): return
            q = parse_qs(u.query)
            self._json(sheet_state(q["sheet"][0]))
        elif u.path == "/api/status":
            con = sqlite3.connect(STATE["db"])
            row = con.execute(
                "SELECT status, frozen_at, change_json FROM workbooks WHERE id=?",
                (STATE["wid"],)).fetchone()
            change = json.loads(row[2]) if row and row[2] else None
            self._json({"status": row[0], "frozen_at": row[1], "boot": STATE["boot"],
                        "change": change, "copy": bool(STATE.get("copy"))})
        elif u.path == "/api/export":
            import urllib.parse as _up
            path = STATE["wb_path"]
            if not path or not os.path.isfile(path):
                self._json({"ok": False, "reason": "当前没有打开的文件"}, 404)
                return
            with open(path, "rb") as f:
                body = f.read()
            name = os.path.basename(path)
            cd = "attachment; filename*=UTF-8''" + _up.quote(name)
            self.send_response(200)
            self.send_header("Content-Type",
                             "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            self.send_header("Content-Disposition", cd)
            self.send_header("X-Gewu-Owner", BRANDING["watermark"])
            self.send_header("Content-Length", str(len(body)))
            self.end_headers()
            self.wfile.write(body)
        elif u.path == "/api/rules":
            if self._busy(): return
            from engine.mirror import list_biz_rules
            sheet = (parse_qs(u.query).get("sheet") or [None])[0]
            self._json(list_biz_rules(STATE["wid"], sheet, STATE["db"]))
        elif u.path == "/api/rules_check":
            if self._busy(): return
            self._json(self._mirror().rules_check())
        elif u.path == "/api/docs":
            self._json(API_DOCS)
        elif u.path == "/api/agent_context":
            # Agent 快照：当前状态 + 可用动作（可直接做 system prompt 的一部分）
            m = self._mirror() if STATE.get("ready") else None
            self._json({
                "ready": bool(STATE.get("ready")),
                "workbook": STATE.get("wb_path"),
                "status": "frozen" if _is_frozen() else "parsed",
                "sheets": ([{"name": x["name"], "visible": x["visible"]}
                            for x in _get_wb(STATE["wb_path"])["sheets"]]
                           if STATE.get("wb_path") and os.path.isfile(STATE["wb_path"]) else []),
                "api_docs_url": "/api/docs",
                "can_write": _is_frozen(),
                "hint": "未冻结 → 先 POST /api/freeze 才能改数" if not _is_frozen()
                        else "已冻结 → 可 POST /api/plan_write 生成写计划"
            })
        elif u.path == "/api/boot":
            self._json({"ready": bool(STATE.get("ready")),
                        "phase": STATE.get("boot_phase"),
                        "error": STATE.get("boot_error"),
                        "fast": bool(STATE.get("fast_boot"))})
        elif u.path == "/api/app_settings":
            try:
                b = json.load(open(os.path.join(
                    os.path.dirname(os.path.abspath(__file__)), "branding.json"),
                    encoding="utf-8"))
                self._json({"close_to_tray": bool(b.get("close_to_tray"))})
            except (OSError, json.JSONDecodeError):
                self._json({"close_to_tray": False})
        elif u.path == "/api/llm_config":
            cfg = llm_mod.load_config()
            if cfg:
                self._json({"protocol": cfg.get("protocol"), "base_url": cfg.get("base_url"),
                            "model": cfg.get("model"), "has_key": bool(cfg.get("api_key"))})
            else:
                self._json(None)
        elif u.path == "/api/sessions":
            if self._busy(): return
            from engine.mirror import chat_session_list
            self._json(chat_session_list(STATE["wid"], STATE["db"]))
        elif u.path == "/api/search_history":
            from engine.mirror import chat_history_search
            q = (parse_qs(u.query).get("q") or [""])[0]
            self._json(chat_history_search(q, STATE["db"]))
        elif u.path == "/api/write_log":
            con = sqlite3.connect(STATE["db"])
            rows = con.execute(
                "SELECT ts, sheet, row_no, col_no, old_repr, new_repr, undone "
                "FROM write_ops WHERE workbook_id=? ORDER BY id DESC LIMIT 30",
                (STATE["wid"],)).fetchall()
            con.close()
            self._json([{"ts": r[0], "sheet": r[1], "row": r[2], "col": r[3],
                         "old": r[4], "new": r[5], "undone": bool(r[6])} for r in rows])

    def _mirror(self):
        m = Mirror(STATE["db"], workbook_id=STATE["wid"])
        return m

    def do_POST(self):
        u = urlparse(self.path)
        n = int(self.headers.get("Content-Length", 0))
        raw = self.rfile.read(n)
        if u.path == "/api/upload":     # 二进制体：不做 JSON 解码，走上传分支
            payload = {}
        else:
            try:
                payload = json.loads(raw or b"{}")
            except (json.JSONDecodeError, UnicodeDecodeError):
                self._json({"ok": False, "reason": "请求体不是合法 UTF-8 JSON"}, 400)
                return
        if u.path == "/api/freeze":
            if self._busy(): return
            result = freeze(STATE["wid"], STATE["db"], edits=payload.get("edits", []))
            result["ok"] = True
            if result["missed"]:
                result["warning"] = (f"{result['missed']} 条修改未命中区块"
                                     f"（镜像已重建，请刷新页面后重试）")
            self._json(result)
        elif u.path == "/api/ask":
            if self._busy(): return
            q = (payload.get("question") or "").strip()
            sheet = (payload.get("sheet") or "").strip() or None
            scope = payload.get("scope") or None
            lang = payload.get("lang") or "zh"
            if isinstance(scope, list):
                scope = [s for s in scope if s] or None
            if not q:
                self._json({"answer": "请输入问题"})
                return
            from engine import mirror as mirror_mod
            sid = payload.get("session_id")
            if not sid:   # 首问自动建会话，标题=首问
                sid = mirror_mod.chat_session_create(
                    STATE["wid"], STATE["db"], title=q, scope=scope)
            hist = mirror_mod.chat_session_load(sid, STATE["db"])[-8:]
            m = self._mirror()
            cfg = llm_mod.load_config()
            if llm_mod.llm_ready(cfg):
                try:
                    ans = m.answer(q, llm=llm_mod.query_intent, sheet_hint=sheet,
                                   history=hist, scope=scope, lang=lang)
                    mirror_mod.chat_session_touch(sid, scope, STATE["db"])
                    mirror_mod.chat_session_append(sid, "user", q)
                    mirror_mod.chat_session_append(sid, "assistant", str(ans))
                    self._json({"answer": ans, "source": "llm", "session_id": sid})
                    return
                except SystemExit as e:  # API 挂了回退规则引擎，不让工作台哑掉
                    self._json({"answer": f"（LLM 调用失败，已回退规则引擎：{e}）\n"
                                           + m.answer(q, sheet_hint=sheet, scope=scope),
                                "source": "rules-fallback", "session_id": sid})
                    return
            self._json({"answer": m.answer(q, sheet_hint=sheet, scope=scope),
                        "source": "rules", "session_id": sid})
        elif u.path == "/api/glossary_save":
            from engine.mirror import add_glossary
            add_glossary(STATE["wid"],
                         term=payload.get("column") or "",
                         meaning=(payload.get("meaning") or "").strip(),
                         db=STATE["db"],
                         col_path=payload.get("column") or "",
                         sheet_name=payload.get("sheet") or None)
            self._json({"ok": True})
        elif u.path == "/api/open":
            path = (payload.get("path") or "").strip().strip('"')
            if not path.lower().endswith(".xlsx") or not os.path.isfile(path):
                self._json({"ok": False, "reason": "不是存在的 .xlsx 文件路径"}, 400)
                return
            try:
                _switch_workbook(os.path.abspath(path))
            except SystemExit as e:
                self._json({"ok": False, "reason": str(e)[:200]}, 400)
                return
            self._json({"ok": True})
        elif u.path == "/api/upload":
            if not (u.query and "name=" in u.query):
                self._json({"ok": False, "reason": "缺文件名"}, 400)
                return
            name = os.path.basename(parse_qs(u.query)["name"][0])
            if not name.lower().endswith(".xlsx"):
                self._json({"ok": False, "reason": "只支持 .xlsx"}, 400)
                return
            updir = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                                 ".cache", "uploads")
            os.makedirs(updir, exist_ok=True)
            dst = os.path.join(updir, name)
            with open(dst, "wb") as f:
                f.write(raw)
            try:
                _switch_workbook(dst, copy=True)
            except SystemExit as e:
                self._json({"ok": False, "reason": str(e)[:200]}, 400)
                return
            self._json({"ok": True, "file": name})
        elif u.path == "/api/rules":
            from engine.mirror import list_biz_rules
            sheet = (parse_qs(u.query).get("sheet") or [None])[0]
            self._json(list_biz_rules(STATE["wid"], sheet, STATE["db"]))
        elif u.path == "/api/rules_save":
            if self._busy(): return
            from engine.mirror import add_biz_rule
            col = (payload.get("column") or "").strip()
            sheet = (payload.get("sheet") or S_SHEET(STATE)).strip()
            if not col:
                self._json({"ok": False, "reason": "缺列名"}, 400)
                return
            def _num(x):
                try:
                    return float(x) if x not in (None, "") else None
                except (TypeError, ValueError):
                    return None
            add_biz_rule(STATE["wid"], sheet, col, rmin=_num(payload.get("min")),
                         rmax=_num(payload.get("max")),
                         pattern=(payload.get("pattern") or "").strip() or None,
                         severity=payload.get("severity") or "block", db=STATE["db"])
            self._json({"ok": True})
        elif u.path == "/api/rules_delete":
            from engine.mirror import del_biz_rule
            del_biz_rule(STATE["wid"], (payload.get("sheet") or "").strip(),
                         (payload.get("column") or "").strip(), STATE["db"])
            self._json({"ok": True})
        elif u.path == "/api/rules_check":
            self._json(self._mirror().rules_check())
        elif u.path == "/api/llm_save":
            data = {k: payload.get(k) or "" for k in
                    ("protocol", "base_url", "model", "api_key")}
            if not data["base_url"] or not data["model"]:
                self._json({"ok": False, "reason": "接口地址与模型名必填"}, 400)
                return
            llm_mod.save_config(**data)
            self._json({"ok": True})
        elif u.path == "/api/llm_test":
            import time as _time
            test = {k: payload.get(k) or "" for k in
                    ("protocol", "base_url", "model", "api_key")}
            cur = llm_mod.load_config() or {}
            # 临时生效被测配置（不动磁盘文件）：写环境变量风格——直接走临时 llm.json 切换太重，
            # 复用 chat 的配置装载：临时把参数注入进程内单例
            tmp = {"protocol": test["protocol"] or cur.get("protocol") or "openai",
                   "base_url": test["base_url"] or cur.get("base_url"),
                   "model": test["model"] or cur.get("model"),
                   "api_key": test["api_key"] or cur.get("api_key") or ""}
            if not tmp["base_url"] or not tmp["model"]:
                self._json({"ok": False, "reason": "接口地址与模型名必填"})
                return
            t0 = _time.time()
            try:
                reply = llm_mod.chat([{"role": "user", "content": "回复两个字：在线"}],
                                     timeout=30, cfg=tmp)
                self._json({"ok": True, "ms": int((_time.time() - t0) * 1000), "reply": str(reply)[:40]})
            except SystemExit as e:
                self._json({"ok": False, "reason": str(e)[:200]})
        elif u.path == "/api/app_settings_save":
            cfg_path = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                                    "branding.json")
            try:
                b = json.load(open(cfg_path, encoding="utf-8"))
            except (OSError, json.JSONDecodeError):
                b = {}
            b["close_to_tray"] = bool(payload.get("close_to_tray"))
            json.dump(b, open(cfg_path, "w", encoding="utf-8"),
                      ensure_ascii=False, indent=2)
            self._json({"ok": True})
        elif u.path == "/api/session_new":
            if self._busy(): return
            from engine.mirror import chat_session_create
            sid = chat_session_create(STATE["wid"], STATE["db"],
                                      scope=payload.get("scope"))
            self._json({"ok": True, "session_id": sid})
        elif u.path == "/api/session_switch":
            if self._busy(): return
            from engine.mirror import chat_session_load, chat_session_list
            sid = int(payload.get("id") or 0)
            mine = next((s for s in chat_session_list(STATE["wid"], STATE["db"])
                         if s["id"] == sid), None)
            if not mine:
                self._json({"ok": False, "reason": "会话不存在"}, 404)
                return
            self._json({"ok": True, "session": mine,
                        "messages": chat_session_load(sid, STATE["db"])})
        elif u.path == "/api/session_delete":
            if self._busy(): return
            from engine.mirror import chat_session_delete, chat_session_list
            sid = int(payload.get("id") or 0)
            if not any(s["id"] == sid for s in
                       chat_session_list(STATE["wid"], STATE["db"])):
                self._json({"ok": False, "reason": "会话不属于当前文件"}, 404)
                return
            chat_session_delete(sid, STATE["db"])
            self._json({"ok": True})
        elif u.path == "/api/sessions_clear":
            if self._busy(): return
            from engine.mirror import chat_sessions_clear
            chat_sessions_clear(STATE["wid"], STATE["db"])
            self._json({"ok": True})
        elif u.path == "/api/note_detach":
            if self._busy(): return
            # 人工取消行内备注挂接：黄块保留，只是不再算作那行的信息
            from engine.mirror import detach_note
            sheet = payload.get("sheet") or ""
            block_id = payload.get("block_id")
            if not sheet or block_id is None:
                self._json({"ok": False, "reason": "缺 sheet/block_id"}, 400)
                return
            res = detach_note(STATE["wid"], sheet, int(block_id), db=STATE["db"])
            _SHEET_CACHE.pop((STATE["wb_path"], STATE["wid"], sheet), None)
            self._json({"ok": True, **res})
        elif u.path == "/api/plan_write":
            if self._busy(): return
            self._json(self._plan_write(payload))
        elif u.path == "/api/execute_write":
            if self._busy(): return
            self._json(self._execute_write(payload.get("op") or {}))
        elif u.path == "/api/undo":
            if self._busy(): return
            res = writer_mod.undo_last_batch(STATE["db"], STATE["wb_path"],
                                             workbook_id=STATE["wid"])
            if res.get("ok"):
                _WB_CACHE.pop(STATE["wb_path"], None)
                _SHEET_CACHE.clear()
            self._json(res)
        else:
            self._json({"ok": False}, 404)

    def _plan_write(self, payload):
        """写指令 → 语义计划 → 坐标解析 → 硬校验 → diff 预览（不执行）。"""
        mode, text = payload.get("mode", "plan"), (payload.get("text") or "").strip()
        if not text:
            return {"ok": False, "clarify": "指令为空"}
        try:
            if mode == "nl":
                plan = llm_mod.write_intent(text, llm_mod.mirror_context(self._mirror()))
            else:
                plan = json.loads(text)
        except SystemExit as e:
            return {"ok": False, "clarify": str(e)}
        except json.JSONDecodeError as e:
            return {"ok": False, "clarify": f"计划 JSON 语法错误：{e}"}
        res = llm_mod.resolve_write_plan(plan, self._mirror())
        if not res.get("ok"):
            out = {"ok": False, "clarify": res.get("clarify", "无法解析")}
            if res.get("candidates"):
                out["candidates"] = res["candidates"][:8]
            return out
        op = res["op"]
        mode = op.pop("mode", "set")
        operand = op.pop("operand", op.get("value"))
        sh = next((s for s in _get_wb(STATE["wb_path"])["sheets"]
                   if s["name"] == op["sheet"]), None)
        old = (sh["cells"].get((op["row"], op["col"])) or {}).get("cv") if sh else None
        try:
            # 数值纪律：算术由引擎基于旧值执行，模型只选 mode 抄操作数
            op["value"] = llm_mod.apply_arith(old, mode, operand)
        except SystemExit as e:
            return {"ok": False, "clarify": str(e)}
        v = writer_mod.validate_plan([op], STATE["wb_path"], STATE["db"], STATE["wid"])
        echo = llm_mod.value_echo_warning(text, operand if mode != "set" else op["value"],
                                          op["value"])
        if echo:
            v["warnings"].append(echo)
        mag = llm_mod.magnitude_warning(op["value"],
                                        llm_mod.column_values(self._mirror(), op))
        if mag:
            v["warnings"].append(mag)
        # #1 命中行键值徽章：确认的不止坐标，还有"是哪一行"
        hit = {}
        try:
            mm = self._mirror()
            blk = next((b for b in mm.blocks(op["sheet"], tables_only=True)
                        if b["r1"] <= op["row"] <= b["r2"]
                        and b["c1"] <= op["col"] <= b["c2"]), None)
            if blk:
                rec = next((r for r in mm._records(blk) if r["_r"] == op["row"]), None)
                if rec:
                    hit = {k: v for k, v in list(rec.items())[1:7]
                           if not k.startswith("_") and v not in (None, "")}
        except Exception:
            pass
        # #4 公式依赖预测：谁引用了目标格（写入后 Excel 自动重算）
        deps = _predict_deps(op)
        return {"ok": True, "op": op, "via": res.get("via"), "old": old,
                "mode": mode, "operand": operand, "validation": v,
                "hit": hit, "deps": deps}

    def _execute_write(self, op):
        """执行写入：服务端重新校验（不信任客户端传来的校验结果）。"""
        if not op or "sheet" not in op:
            return {"ok": False, "reason": "缺少操作"}
        v = writer_mod.validate_plan([op], STATE["wb_path"], STATE["db"], STATE["wid"])
        if not v["ok"]:
            return {"ok": False, "reason": "校验未通过", "validation": v}
        deps = _predict_deps(op)
        dep_olds = {}
        if deps:
            _sh = next((x for x in _get_wb(STATE["wb_path"])["sheets"]
                        if x["name"] == op["sheet"]), None)
            if _sh:
                dep_olds = {(d["row"], d["col"]):
                            (_sh["cells"].get((d["row"], d["col"])) or {}).get("cv")
                            for d in deps}
        try:
            res = writer_mod.execute([op], STATE["wb_path"], STATE["db"], STATE["wid"])
        except writer_mod.Rejection as e:
            return {"ok": False, "reason": str(e)}
        pc = writer_mod.post_check([op], STATE["wb_path"], STATE["db"], STATE["wid"])
        # 联动真验证：重读文件比对依赖格缓存值（COM 路径 Excel 已重算）
        deps_changed = []
        if dep_olds:
            _after = next((x for x in extract_workbook(STATE["wb_path"])["sheets"]
                           if x["name"] == op["sheet"]), None)
            if _after:
                for (r, c), old_v in dep_olds.items():
                    new_v = (_after["cells"].get((r, c)) or {}).get("cv")
                    if new_v != old_v:
                        deps_changed.append({"row": r, "col": c})
        _WB_CACHE.pop(STATE["wb_path"], None)
        _SHEET_CACHE.clear()
        return {"ok": res["ok"] and pc["ok"], "engine": res["engine"],
                "backup": os.path.basename(res["backup"]), "post_check": pc,
                "deps_predicted": len(dep_olds), "deps_changed": deps_changed}


PAGE = r"""<!doctype html><html lang="zh"><head><meta charset="utf-8">
<title>格物 Gewu · 复杂表格理解工作台</title>
<meta name="generator" content="Gewu/1.0 (wm:__WATERMARK__)">
<!-- gewu-owner-fingerprint: __WATERMARK__ --><style>
body{font:13px/1.5 system-ui;margin:0;display:flex;flex-direction:column;height:100vh}
header{display:flex;gap:12px;align-items:center;padding:8px 16px;
        background:linear-gradient(90deg,#0f172a,#1e3a5f 70%,#1f2937);color:#fff;
        box-shadow:0 1px 6px rgba(15,23,42,.35);position:relative;z-index:10}
header b{font-size:14px}
#status{padding:2px 10px;border-radius:10px;background:#b45309}
#status.frozen{background:#15803d}
main{flex:1;display:flex;min-height:0}
#appBody{flex:1;display:flex;min-height:0}
#stage{flex:1;min-width:0}
#panelHandle{width:5px;cursor:col-resize;background:#e2e8f0}
#panelHandle:hover{background:#93c5fd}
#panelExpand{border:0;background:#dbeafe;font-size:15px;cursor:pointer;padding:0 3px}
#sidePanel{width:430px;min-width:300px;max-width:70vw;display:flex;flex-direction:column;background:#fff;border-left:1px solid #d1d5db}
#ptabs{display:flex;gap:2px;background:#f1f5f9;border-bottom:1px solid #d1d5db;padding:4px;align-items:center}
.ptab{border:0;background:transparent;cursor:pointer;font-size:12.5px;padding:5px 10px;border-radius:6px;color:#475569}
.ptab.on{background:#2563eb;color:#fff}
#panelCollapse{margin-left:auto;border:0;background:transparent;cursor:pointer;font-size:15px;color:#64748b}
.pane{display:none;flex:1;min-height:0;flex-direction:column;gap:6px;padding:8px 10px;overflow:auto}
.pane.on{display:flex}
#chatToolbar{display:flex;align-items:center;gap:4px;padding:4px 6px;background:#f8fafc;
  border:1px solid #e2e8f0;border-radius:8px;font-size:12px}
#scopeBtn{border:1px solid #cbd5e1;background:#fff;border-radius:14px;padding:2px 12px;
  cursor:pointer;font-size:12px;color:#1e293b;white-space:nowrap}
#scopeBtn:hover{border-color:#2563eb;color:#2563eb}
.toolBtn{border:1px solid #e2e8f0;background:#fff;border-radius:6px;padding:2px 8px;
  cursor:pointer;font-size:14px;color:#475569;line-height:1.2}
.toolBtn:hover{border-color:#2563eb;color:#2563eb;background:#eff6ff}
#scopePop{display:none;position:absolute;left:8px;right:8px;top:38px;z-index:25;
  background:#fff;border:1px solid #cbd5e1;border-radius:10px;padding:8px 10px;
  box-shadow:0 6px 24px rgba(0,0,0,.12)}
.scopeChip{cursor:pointer;border:1px solid #cbd5e1;border-radius:12px;padding:2px 10px;font-size:12px;background:#fff;user-select:none}
.scopeChip.on{background:#2563eb;color:#fff;border-color:#2563eb}
#tabBar .ftab{display:inline-flex;align-items:center;gap:4px;padding:3px 10px;border-radius:6px 6px 0 0;
  cursor:pointer;font-size:12px;background:#f1f5f9;color:#475569;border:1px solid #cbd5e1;border-bottom:0;max-width:220px}
#tabBar .ftab.on{background:#fff;color:#1e293b;font-weight:600;border-bottom:2px solid #2563eb}
#tabBar .ftab .tName{overflow:hidden;text-overflow:ellipsis;white-space:nowrap;max-width:160px}
#tabBar .ftab .tClose{border:0;background:transparent;cursor:pointer;color:#94a3b8;font-size:11px;padding:0 2px}
#tabBar .ftab .tClose:hover{color:#dc2626}
#tabBar .ftab .tBadge{font-size:10px;color:#d97706}
.sRow{display:flex;align-items:center;gap:4px}
.sRow button{border:1px solid #cbd5e1;background:#fff;border-radius:6px;cursor:pointer;font-size:13px;padding:2px 8px}
#sessPop{display:none;position:absolute;right:10px;top:86px;z-index:30;width:250px;max-height:300px;overflow:auto;
  background:#fff;border:1px solid #cbd5e1;border-radius:8px;padding:6px;box-shadow:0 6px 24px rgba(0,0,0,.12)}
#sessPop .sItem{padding:5px 7px;border-radius:6px;cursor:pointer;font-size:12.5px;margin-top:3px;position:relative}
#sessPop .sItem.on{background:#dbeafe}
#sessPop .sItem .sDel{position:absolute;right:4px;top:3px;border:0;background:transparent;cursor:pointer;color:#94a3b8;font-size:11px}
#sessPop .sItem .sDel:hover{color:#dc2626}
.askRow{display:flex;gap:6px;padding-top:6px}
.pane textarea{font:12.5px monospace;border:1px solid #cbd5e1;border-radius:6px;padding:6px}
#p-chat{position:relative;overflow:hidden;gap:0!important;padding-top:6px!important}
#p-chat #chatLog{flex:1}
#p-write{overflow:auto}
#diffCard{margin-top:2px}
#stage{flex:1;overflow:auto;position:relative;background:#eef2f7}
#gridWrap{position:relative;display:inline-block;min-width:100%;padding:0 0 30px 0}
table{border-collapse:separate;border-spacing:0;background:#fff;table-layout:fixed}
td{border-right:1px solid #edf1f6;border-bottom:1px solid #edf1f6;padding:3px 7px;
    min-width:52px;max-width:220px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;
    font-size:12px;height:20px}
tbody tr:nth-child(even) td:not(.hrow):not(.hcol){background:#fbfdff}
tbody tr:hover td:not(.hrow):not(.hcol){background:#eef6ff}
td.formula{color:#1d4ed8}
/* sticky 列头/行号：滚动不丢参照（separate 边框下才可靠） */
td.hcol{position:sticky;top:0;z-index:2;background:#e8eef6;font-weight:600;color:#475569;
        border-bottom:2px solid #cbd5e1}
td.hrow{position:sticky;left:0;z-index:1;background:#e8eef6;color:#475569;min-width:34px;text-align:right}
td.hcol.hrow{z-index:3}
#stage::-webkit-scrollbar,#chatLog::-webkit-scrollbar,.pane::-webkit-scrollbar,
#writeLog::-webkit-scrollbar,#histList::-webkit-scrollbar,#histResults::-webkit-scrollbar,
#sessPop::-webkit-scrollbar{width:9px;height:9px}
#stage::-webkit-scrollbar-thumb,#chatLog::-webkit-scrollbar-thumb,.pane::-webkit-scrollbar-thumb,
#writeLog::-webkit-scrollbar-thumb,#histList::-webkit-scrollbar-thumb,#histResults::-webkit-scrollbar-thumb,
#sessPop::-webkit-scrollbar-thumb{background:#c3cdd9;border-radius:5px;border:2px solid transparent;background-clip:content-box}
#stage::-webkit-scrollbar-track{background:transparent}
.blk{position:absolute;border:2px solid;border-radius:4px;pointer-events:auto;
     cursor:pointer;transition:background .15s}
.blk .tag{position:absolute;top:-11px;left:4px;font-size:11px;background:#fff;
          padding:0 5px;border-radius:8px;border:1px solid #ccc;white-space:nowrap}
.blk.sel{border-width:3px}
@keyframes cellFlash{0%,100%{outline:0 solid transparent}25%{outline:3px solid #f59e0b;background:#fef3c7 !important}60%{outline:3px solid #f59e0b;background:#fef3c7 !important}}
td.flashCell{animation:cellFlash 2.2s ease 2}
td.flashDone{animation:cellFlash 1.2s ease 2;outline-color:#16a34a !important}
#side h3{margin:4px 0}
.kv{display:grid;grid-template-columns:90px 1fr;gap:2px 8px;font-size:12px;margin:8px 0}
.kv b{color:#475569}
header button,#guide button{font:12px system-ui;padding:4px 10px;margin:2px;cursor:pointer;border-radius:6px;
       border:1px solid #cbd5e1;background:#fff}
button.primary{background:#1d4ed8;color:#fff;border-color:#1d4ed8}
#guide{background:#eff6ff;border-bottom:1px solid #bfdbfe;padding:6px 14px;font-size:12px;
       display:flex;gap:10px;align-items:center}
#guide b{color:#1e3a8a}
input[type=range]{vertical-align:middle}
.hint{color:#64748b;font-size:11px}
ul{padding-left:16px;margin:4px 0}
#chatLog{flex:1;overflow:auto;border:1px solid #e2e8f0;border-radius:8px;padding:8px;font-size:14px;line-height:1.65;background:#fafbfc}
#chatLog .u{color:#fff;background:linear-gradient(135deg,#2563eb,#3b82f6);display:block;width:fit-content;
   margin:8px 0 2px auto;padding:4px 12px;border-radius:12px 12px 3px 12px;font-size:13px;
   box-shadow:0 1px 3px rgba(37,99,235,.3)}
/* ===== 聊天区布局 ===== */
#chatLog{display:flex;flex-direction:column;gap:2px}
#chatLog .a{color:#1e293b;white-space:pre-wrap;margin:2px 0 6px;padding:6px 10px;
  border-left:3px solid #93c5fd;border-radius:0 8px 8px 0;background:linear-gradient(90deg,#f8fafc,#fff);
  line-height:1.75;font-size:13.5px;box-shadow:0 1px 2px rgba(0,0,0,.04)}
#chatLog .a b.mdBold{color:#0f172a;font-weight:700}
#chatLog .a code.mdCode{background:#e0e7ff;border-radius:4px;padding:1px 6px;
  font:12.5px Consolas,monospace;color:#3730a3;border:1px solid #c7d2fe}
/* 公式（Excel 公式原文）：等宽+蓝色调+浅底 */
.mdFormula{background:#eff6ff;border:1px solid #bfdbfe;border-radius:4px;padding:1px 6px;
  font:12.5px Consolas,'Courier New',monospace;color:#1d4ed8;white-space:nowrap;display:inline-block}
/* 代码块：深底 */
pre.mdPre{background:#1e293b;color:#e2e8f0;border-radius:8px;padding:8px 12px;margin:4px 0;
  font:12px Consolas,monospace;overflow-x:auto;white-space:pre;line-height:1.5}
pre.mdPre code{color:#e2e8f0}
/* 列表 */
.mdLi{display:block;padding-left:14px;margin:1px 0;color:#334155}
.mdLi::before{content:"●";font-size:8px;color:#3b82f6;margin-right:6px;vertical-align:2px}
.mdOl{display:block;padding-left:8px;margin:1px 0;color:#334155}
/* 引用 */
.mdQuote{display:block;border-left:3px solid #fbbf24;padding:2px 10px;margin:2px 0;
  background:#fffbeb;border-radius:0 4px 4px 0;color:#78350f;font-style:italic}
/* 表格行 */
.mdTr{display:block;font-family:Consolas,monospace;font-size:12px;color:#475569;
  background:#f1f5f9;padding:1px 6px;margin:1px 0;border-radius:3px}
#diffCard{border:1px solid #f59e0b;background:#fffbeb;border-radius:8px;padding:6px 8px;font-size:12px;max-height:110px;overflow:auto}
#diffCard .bad{border-color:#dc2626;background:#fef2f2}
#diffCard .ok{border-color:#16a34a;background:#f0fdf4}
</style></head><body>
<header><b id="brandLogo" style="font-size:16px;letter-spacing:2px">格物 <span style="background:linear-gradient(90deg,#60a5fa,#93c5fd);-webkit-background-clip:text;background-clip:text;color:transparent">Gewu</span></b>
  <span class="hint" id="brandTag" style="font-size:11px">复杂表格理解工作台</span>
  <span id="ownerChip" onclick="showAbout()" style="cursor:pointer;font-size:11px;color:#64748b;border:1px solid #cbd5e1;border-radius:10px;padding:1px 8px">◈ __OWNER__</span>
  <button id="langBtn" title="中文 / English" style="font-size:11px;padding:2px 8px;cursor:pointer"
          onclick="toggleLang()">中/EN</button>
  <span id="wbName"></span>
  <select id="sheetSel"></select>
  <span id="status">解析</span>
  <label class="hint" style="display:flex;gap:2px;align-items:center">
    <input type="checkbox" id="showMask" checked onchange="drawOverlays()"><span data-i18n="mask">蒙版</span></label>
  <label class="hint" style="display:flex;gap:2px;align-items:center"><span data-i18n="opacity">浓度</span>
    <input type="range" id="opSlider" min="0" max="50" value="14"
           oninput="paintAll()" style="width:70px"></label>
  <span id="pendingChip" style="display:none;color:#b45309;background:#fef3c7;border:1px solid #fcd34d;border-radius:10px;padding:2px 8px;font-size:12px"></span>
  <span style="flex:1"></span>
  <button id="freezeBtn" class="primary" onclick="doFreeze()" data-i18n="freeze">✔ 检查无误，冻结</button>
  <button onclick="showHelp()" title="注意事项" data-i18n-t="help_t" style="font-weight:700">❓</button>
  <button onclick="showAbout()" title="关于格物" data-i18n-t="about_t" style="font-weight:700">ⓘ</button>
  <a id="exportBtn" href="/api/export" download style="display:none;text-decoration:none;color:#fff;
     background:#d97706;border-radius:6px;padding:3px 10px;font-size:12px;font-weight:700"
     data-i18n="export_btn" data-i18n-t="export_t"
     title="下载改后的文件（副本模式的改动只有导出才是成品）">📥 导出改后文件</a>
  <button onclick="showOpen()" title="打开文件" data-i18n-t="open_t" style="font-weight:700">📂</button>
  <button onclick="showLLM()" title="LLM 设置" data-i18n-t="llm_t" style="font-weight:700">⚙️</button>
</header>
<div id="helpBox" style="display:none;position:fixed;inset:0;background:rgba(15,23,42,.45);z-index:50" onclick="if(event.target===this)this.style.display='none'">
 <div data-i18n-html="help_full" style="position:absolute;top:50%;left:50%;transform:translate(-50%,-50%);background:#fff;border-radius:12px;max-width:660px;width:92%;max-height:82vh;overflow:auto;padding:20px 26px;font-size:13.5px;line-height:1.75;box-shadow:0 10px 40px rgba(0,0,0,.25)">
  <div style="display:flex;align-items:center;border-bottom:2px solid #e2e8f0;padding-bottom:8px;margin-bottom:4px">
    <b style="font-size:16.5px">使用注意事项</b>
    <button onclick="document.getElementById('helpBox').style.display='none'" style="margin-left:auto">关闭 ✕</button>
  </div>

  <div style="margin:12px 0 4px;font-weight:700;color:#1d4ed8">① 冻结是什么</div>
  <div style="padding-left:10px">
    冻结 = 你确认"机器切的框对了"，同时落地确认页上的手动修正，并打开<b>改数</b>的总闸门。<br>
    <b>不冻结 = 只读模式</b>：问答 / 表格解读 / 勾稽核对随便用，改数一律被拦，文件绝对安全。
  </div>

  <div style="margin:14px 0 4px;font-weight:700;color:#1d4ed8">② 问答怎么问</div>
  <div style="padding-left:10px">
    "这个表格 / 当前表格" 指你现在看的 sheet；"整个文件 / 这些表什么关系" 问全簿。<br>
    追问（如"那为什么这么算"）自动带最近 4 轮上下文；点 <b>🗑 新对话</b> 清空重来。
  </div>

  <div style="margin:14px 0 4px;font-weight:700;color:#1d4ed8">③ 改数怎么写</div>
  <div style="padding-left:10px">
    改数<b>不认"当前页面"</b>，按 sheet 关键词 + 行键 + 列名在全簿定位——指令里带上 sheet 名最稳<br>
    （例："把<b>汇总表晨光韩国</b>的克重250加3"）。命中多行会被拦下让你说清。<br>
    流程：语义解析 → 硬校验 → <b>diff 预览（旧值→新值）</b> → 你确认 → 写入（自动备份，可撤销）。上传副本模式下写入只进副本，成品需点页眉「📥 导出改后文件」下载。
  </div>

  <div style="margin:14px 0 4px;font-weight:700;color:#1d4ed8">④ 区块按钮的效果（冻结后生效）</div>
  <table style="width:100%;border-collapse:collapse;font-size:13px;margin:4px 0">
    <tr style="background:#f1f5f9"><th style="text-align:left;padding:4px 8px;border:1px solid #e2e8f0">按钮</th><th style="text-align:left;padding:4px 8px;border:1px solid #e2e8f0">效果</th></tr>
    <tr><td style="padding:4px 8px;border:1px solid #e2e8f0">切为 🟨文本/批注</td><td style="padding:4px 8px;border:1px solid #e2e8f0">这块退出问答与改数（数据仍在文件里，只是不当表用）</td></tr>
    <tr><td style="padding:4px 8px;border:1px solid #e2e8f0">切为 ⬜忽略</td><td style="padding:4px 8px;border:1px solid #e2e8f0">同上：退出问答 + 拒绝改数</td></tr>
    <tr><td style="padding:4px 8px;border:1px solid #e2e8f0">切为 🟥混合待复核</td><td style="padding:4px 8px;border:1px solid #e2e8f0">隔离：禁改数、不进问答，直到定型成数据表或文本</td></tr>
    <tr><td style="padding:4px 8px;border:1px solid #e2e8f0">🔒只读</td><td style="padding:4px 8px;border:1px solid #e2e8f0"><b>仅本应用内只读</b>：工具禁止改这块（可正常问答）。<b>文件本身不变</b>——Excel/WPS 里仍可编辑，文件属性也不变</td></tr>
    <tr><td style="padding:4px 8px;border:1px solid #e2e8f0">🔓可写</td><td style="padding:4px 8px;border:1px solid #e2e8f0">解除上面的应用内只读</td></tr>
  </table>
  <div style="padding-left:10px;color:#475569;font-size:12.5px">
    误触保护：修改点冻结才生效（顶部有"待提交 N 处"提示）；冻结前刷新页面 = 全部撤销；每块可 <b>↩ 撤销本块修改</b>；冻结后切错了，再切一次并重新冻结即可。<br>
    注：引擎判为文本的黄块切"数据表"仅作视觉标记（无列结构，不参与问答）。
  </div>

  <div style="margin:14px 0 4px;font-weight:700;color:#1d4ed8">⑤ 文件被改动后</div>
  <div style="padding-left:10px">
    自己用工具改的数 → 重开无缝续用（冻结保持）。<br>
    别人改了数值 → 重新拖入，绿条提示"数值已刷新"，冻结仍有效，不用重确认。<br>
    结构变了（增删表/改表头） → 重新拖入，橙条列出变化清单，未变块的标记自动保留，复核后重新冻结。
  </div>

  <div style="margin:14px 0 4px;font-weight:700;color:#1d4ed8">⑥ 打开文件与副本</div>
  <div style="padding-left:10px">
    三种打开：<b>📂 路径</b>（Shift 右键复制文件地址粘贴——写回原文件，推荐）；
    <b>上传副本</b>（分析体验——改动只写副本，页眉出现〔副本〕徽记）；
    拖文件到桌面应用图标 / 拖入桌面窗口（=写回原文件）；网页页面里拖入（=上传副本）。<br>
    <b>副本模式的成品 = 点页眉「📥 导出改后文件」下载的文件</b>——原上传的文件永远不含你的改动，别把它发给别人。
  </div>
  <div style="margin-top:14px;padding:8px 10px;background:#f8fafc;border-left:3px solid #94a3b8;color:#64748b;font-size:12.5px">
    改数前会重过一遍全部校验；表头、汇总行、公式格、忽略区块永远不会被写入。所有人工标记（类型/只读/词典/行内备注挂接）只存在于本工具的镜像里，<b>永不修改你的文件</b>（改数指令除外）。
  </div>
 </div>
</div>
<div id="guide" data-i18n-html="guide">
  <b>怎么用：</b><b>蓝框=表格</b>（问答和改数都发生在蓝框里，只需扫一眼框得对不对）；
  <b>黄虚线=标题/备注</b>（不当表格处理，<u>不用管它</u>；标了「表内批注」的是表格中间夹的一行备注，正常现象）；
  <b>红框=机器拿不准</b>（最好点开看一眼）。
  没问题 → 点右上「检查无误，冻结」→ 去底部工作台。框不对 → 点色块改类型。
  <button onclick="this.parentNode.remove()" style="margin-left:auto">知道了 ✕</button>
</div>
<div id="llmBox" style="display:none;position:fixed;inset:0;background:rgba(15,23,42,.45);z-index:50" onclick="if(event.target===this)this.style.display='none'">
 <div style="position:absolute;top:50%;left:50%;transform:translate(-50%,-50%);background:#fff;border-radius:12px;max-width:560px;width:92%;max-height:86vh;overflow:auto;padding:22px 26px;font-size:13px;line-height:1.7;box-shadow:0 10px 40px rgba(0,0,0,.25)">
  <div style="display:flex;align-items:center"><b style="font-size:15px" data-i18n="llm_h">⚙️ LLM 设置</b>
    <button onclick="document.getElementById('llmBox').style.display='none'" style="margin-left:auto" data-i18n="llm_close">关闭 ✕</button></div>
  <div style="margin:10px 0 4px"><b data-i18n="llm_preset_h">预设</b>
    <select id="llmPreset" onchange="applyPreset()" style="width:100%;font-size:13px;padding:5px;margin-top:3px">
      <option value="">— 选择预设自动填充 —</option>
      <option value="siliconflow">SiliconFlow（硅基流动）</option>
      <option value="openai">OpenAI</option>
      <option value="deepseek">DeepSeek 官方</option>
      <option value="glm">智谱 GLM</option>
      <option value="anthropic">Anthropic Claude</option>
      <option value="ollama">Ollama（本地）</option>
      <option value="lmstudio">LM Studio（本地）</option>
      <option value="custom">自定义（OpenAI 兼容）</option>
    </select></div>
  <div style="display:grid;grid-template-columns:110px 1fr;gap:6px 8px;margin-top:8px;align-items:center">
    <b data-i18n="llm_proto">协议</b>
    <select id="llmProto" style="font-size:13px;padding:5px">
      <option value="openai">OpenAI 兼容</option>
      <option value="anthropic">Anthropic 兼容</option>
    </select>
    <b data-i18n="llm_base">接口地址</b>
    <input id="llmBase" placeholder="https://api.siliconflow.cn/v1" style="font-size:13px;padding:5px 8px;border:1px solid #cbd5e1;border-radius:6px">
    <b data-i18n="llm_model">模型名</b>
    <input id="llmModel" placeholder="deepseek-ai/DeepSeek-V4-Flash" style="font-size:13px;padding:5px 8px;border:1px solid #cbd5e1;border-radius:6px">
    <b data-i18n="llm_key">API Key</b>
    <input id="llmKey" type="password" placeholder="sk-...（本地模型可留空）" style="font-size:13px;padding:5px 8px;border:1px solid #cbd5e1;border-radius:6px">
  </div>
  <div class="hint" style="font-size:11.5px;margin-top:6px" data-i18n="llm_privacy">密钥只保存在本机 llm.json，请求直连你所选的服务商；不配置则问答走内置纯代码引擎。</div>
  <div style="display:flex;align-items:center;gap:6px;margin-top:8px;padding:6px 8px;background:#f8fafc;border:1px solid #e2e8f0;border-radius:6px">
    <input type="checkbox" id="trayToggle" onchange="saveTraySetting(this.checked)" style="cursor:pointer">
    <span style="font-size:12.5px" data-i18n="tray_label">关闭窗口时缩到系统托盘（引擎保持运行，重开秒进；取消勾选=直接退出）</span>
  </div>
  <div class="hint" style="font-size:11px;margin-top:2px;margin-left:8px" data-i18n="tray_note">仅在桌面应用中生效；改后需重启桌面应用。网页版无此概念。</div>
  <div style="display:flex;gap:6px;margin-top:10px;align-items:center">
    <button class="primary" onclick="llmSave()" data-i18n="llm_save">保存</button>
    <button onclick="llmTest()" data-i18n="llm_test">测试连接</button>
    <span id="llmState" class="hint" style="font-size:12px"></span>
  </div>
 </div>
</div>
<div id="openBox" style="display:none;position:fixed;inset:0;background:rgba(15,23,42,.45);z-index:50" onclick="if(event.target===this)this.style.display='none'">
 <div style="position:absolute;top:50%;left:50%;transform:translate(-50%,-50%);background:#fff;border-radius:12px;max-width:520px;width:90%;padding:22px 26px;font-size:13.5px;box-shadow:0 10px 40px rgba(0,0,0,.25)">
  <div style="display:flex;align-items:center"><b style="font-size:15px" data-i18n="open_title">📂 打开文件</b>
    <button onclick="document.getElementById('openBox').style.display='none'" style="margin-left:auto">关闭 ✕</button></div>
  <div style="margin:12px 0 4px"><b data-i18n="open_path_h">① 文件路径</b><span class="hint">（写回原文件，推荐）</span>
    <div class="hint" style="font-size:11.5px"><span data-i18n="open_path_hint">在资源管理器按住 Shift 右键文件 →「复制文件地址」粘贴到这。改数直接写回这个原文件。</span></div></div>
  <div style="display:flex;gap:6px">
    <input id="openPath" data-i18n-ph="open_ph" placeholder="C:\Users\...\表.xlsx" style="flex:1;padding:6px 10px;font-size:13px;border:1px solid #cbd5e1;border-radius:6px">
    <button class="primary" onclick="doOpenPath()" data-i18n="open_do">打开</button>
  </div>
  <div id="openErr" style="color:#dc2626;font-size:12px;margin-top:4px"></div>
  <div style="margin:14px 0 4px"><b data-i18n="open_copy_h">② 上传副本</b><span class="hint" data-i18n="open_copy_note">（分析体验用；改数只写副本——成品须点页眉「📥 导出改后文件」下载，别把原上传文件当成品发人）</span>
    <div style="display:flex;gap:6px;margin-top:6px;align-items:center">
      <input type="file" id="openFile" accept=".xlsx" style="flex:1;font-size:12.5px">
      <button onclick="doOpenUpload()" data-i18n="open_up">上传并打开</button>
    </div>
    <div id="openUpState" class="hint" style="font-size:12px;margin-top:4px"></div></div>
 </div>
</div>
<div id="aboutBox" style="display:none;position:fixed;inset:0;background:rgba(15,23,42,.45);z-index:50" onclick="if(event.target===this)this.style.display='none'">
 <div style="position:absolute;top:50%;left:50%;transform:translate(-50%,-50%);background:#fff;border-radius:14px;max-width:520px;width:90%;padding:26px 30px;font-size:13.5px;line-height:1.8;box-shadow:0 10px 40px rgba(0,0,0,.25)" id="aboutBody"></div>
</div>
<div id="tabBar" style="display:none;gap:2px;padding:3px 8px;background:#e2e8f0;border-bottom:1px solid #cbd5e1;overflow-x:auto;white-space:nowrap"></div>
<div id="changeBar" style="display:none;border-bottom:1px solid #fdba74;background:#fff7ed;padding:6px 14px;font-size:12.5px"></div>
<div id="appBody">
<div id="stage"><div id="gridWrap"><table id="grid"></table>
  <div id="overlays"></div></div></div>
<button id="panelExpand" style="display:none" title="展开面板">»</button>
<div id="panelHandle" title="拖动调宽窄"></div>
<div id="sidePanel">
  <div id="ptabs">
    <button class="ptab on" data-p="review" data-i18n="tab_review">📋 审阅</button>
    <button class="ptab" data-p="chat" data-i18n="tab_chat">💬 问答</button>
    <button class="ptab" data-p="write" data-i18n="tab_write">✏️ 改数</button>
    <button class="ptab" data-p="history" data-i18n="tab_history">🕘 历史</button>
    <button id="panelCollapse" title="收起面板（表格全宽）" data-i18n-t="collapse">«</button>
  </div>
  <div id="p-review" class="pane on">
    <h3 style="margin:4px 0" data-i18n="review_h">区块审阅</h3><div class="hint" data-i18n="review_hint">点色块查看详情；不确定怎么操作就点右上 ❓——大多数文件直接冻结即可。</div>
    <div id="detail" class="kv"></div>
    <details id="glossaryBox" style="display:none;margin-top:10px;border-top:1px solid #e2e8f0;padding-top:6px">
    <summary style="cursor:pointer;font-weight:600" data-i18n="glossary">📖 字段口径词典</summary>
    <div id="glossaryList" class="hint">选中数据表区块后可给字段加口径注释</div>
    </details>
    <details id="rulesBox" style="display:none;margin-top:10px;border-top:1px solid #e2e8f0;padding-top:6px">
    <summary style="cursor:pointer;font-weight:600" data-i18n="rules_h">📏 业务规则（值域）</summary>
  <div class="hint" style="font-size:11.5px" data-i18n="rules_note">给数值列设上下限：block=越界拒写，warn=仅警告。规则由引擎硬执行，LLM 只知情。</div>
  <div style="display:flex;gap:4px;flex-wrap:wrap;margin-top:4px;align-items:center">
    <select id="ruleCol" style="font-size:12px;max-width:120px"></select>
    <input id="ruleMin" placeholder="≥" style="width:44px;font-size:12px" data-i18n-ph="rules_min_ph">
    <input id="ruleMax" placeholder="≤" style="width:44px;font-size:12px" data-i18n-ph="rules_max_ph">
    <input id="rulePat" placeholder="正则(选填)" style="width:76px;font-size:12px" data-i18n-ph="rules_pat_ph">
    <select id="ruleSev" style="font-size:12px"><option value="block">block</option><option value="warn">warn</option></select>
    <button style="font-size:11px;padding:2px 8px" onclick="saveRule()" data-i18n="rules_save">保存</button>
    <button style="font-size:11px;padding:2px 8px" onclick="delRule()" data-i18n="rules_del">删除</button>
  </div>
  <div style="display:flex;gap:4px;margin-top:4px;align-items:center">
    <button style="font-size:11px;padding:2px 8px" onclick="runRulesCheck()">▶ <span data-i18n="rules_check">全表体检</span></button>
    <span class="hint" style="font-size:11px" data-i18n="rules_check_note">按现有数据跑一遍规则，列出违规行</span>
  </div>
  <div id="rulesList" class="hint" style="font-size:11.5px">（选中数据表后可配置）</div>
  <div id="rulesCheckOut" style="font-size:11.5px"></div>
  </details>
  </div>
  <div id="p-chat" class="pane">
    <div id="chatToolbar">
      <button id="scopeBtn" onclick="toggleScopePop()" title="选择问答范围">
        <span data-i18n="scope">范围</span>: <span id="scopeLabel">整个文件</span> ▾
      </button>
      <span style="flex:1"></span>
      <span id="curTitle" class="hint" style="max-width:40%;overflow:hidden;text-overflow:ellipsis;white-space:nowrap">（提问后自动建会话）</span>
      <button class="toolBtn" title="历史会话" onclick="toggleSessPop()" data-i18n-t="sess_hist">🕘</button>
      <button class="toolBtn" title="新对话" data-i18n-t="sess_new" onclick="newChat()">＋</button>
    </div>
    <div id="scopePop">
      <div class="hint" style="font-size:11px;margin-bottom:4px">点击选择范围（可多选 sheet）</div>
      <div id="scopeChips" style="display:flex;flex-wrap:wrap;gap:3px;max-height:200px;overflow:auto"></div>
    </div>
    <div id="sessPop"><div id="sessionList" class="hint">（尚无会话）</div></div>
    <div id="chatLog"><div class="a hint" data-i18n="chat_hint">问数值（合计/最高的）→ 引擎计算；问理解（这表是做什么的/为什么这么算）→ LLM 解读；问评价（分析一下）→ 数据解读。追问自动带最近 4 轮上下文。</div></div>
    <div class="askRow">
      <input id="askBox" data-i18n-ph="ask_ph" placeholder="问：销售额合计 / 这表是做什么的 / 分析评价一下 （回车发送）" style="flex:1;min-width:0;padding:6px 10px;font-size:13px" onkeydown="if(event.key==='Enter'){event.preventDefault();doAsk()}">
      <button class="primary" onclick="doAsk()">问</button>
    </div>
  </div>
  <div id="p-write" class="pane">
    <div style="font-size:12px;color:#475569">✏️ <b data-i18n="write_h">改数（写回 Excel）</b>：说一句话定位到某行某列 → 机器算新值 → 看差异 → 确认才写入（自动备份，可撤销）。改数<b>不认对话范围</b>，按指令语义全簿定位。</div>
    <textarea id="writeBox" style="height:56px" data-i18n-ph='write_ph' placeholder='自然语言：把1月云澜居的成交套数改成7&#10;或手写计划 JSON：{"find":{"keys":{"列":"值"}},"set":{"column":"列","value":7}}' onkeydown="if(event.key==='Enter'&&(event.ctrlKey||event.metaKey)){event.preventDefault();doPlanWrite(/[\u4e00-\u9fff]/.test(this.value)&&!this.value.trim().startsWith('{')?'nl':'plan')}"></textarea>
    <div style="display:flex;gap:4px;flex-wrap:wrap">
      <button class="primary" onclick="doPlanWrite('nl')" data-i18n-t="btn_nl_t" data-i18n="btn_nl">💬 自然语言</button>
      <button onclick="doPlanWrite('plan')" data-i18n="btn_plan">📋 计划 JSON</button>
      <button onclick="doUndo()" style="margin-left:auto" data-i18n="btn_undo">↩ 撤销一批</button>
    </div>
    <div id="diffCard" class="hint" data-i18n="diff_hint">语义解析 → 硬校验（冻结/表头/汇总行等全不许碰）→ diff 预览 → 确认执行</div>
    <div style="margin-top:8px"><b style="font-size:12.5px" data-i18n="write_log">最近写入</b>
      <div id="writeLog" class="hint" style="border:1px solid #e2e8f0;border-radius:8px;padding:6px 8px;flex:1;overflow:auto;font-size:12px">（切到本页签时加载）</div></div>
  </div>
  <div id="p-history" class="pane">
    <div style="display:flex;align-items:center;gap:8px">
      <b style="font-size:13px" data-i18n="hist_h">本文件的会话</b>
      <button style="margin-left:auto;font:11px system-ui;padding:2px 8px" onclick="clearSessions();loadHistList()" data-i18n="hist_clear">清空全部</button>
    </div>
    <div id="histList" class="hint" style="border:1px solid #e2e8f0;border-radius:8px;padding:6px 8px;max-height:180px;overflow:auto;font-size:12.5px">（切到本页签时加载）</div>
    <div style="margin-top:10px"><b style="font-size:13px" data-i18n="hist_search_h">🔎 搜全部文件的历史</b>
      <div style="display:flex;gap:4px;margin-top:4px">
        <input id="histSearch" data-i18n-ph="hist_ph" placeholder="关键词（跨工作簿检索，只读参考）" style="flex:1;min-width:0;padding:5px 10px;font-size:13px" onkeydown="if(event.key==='Enter'){event.preventDefault();doHistSearch()}">
        <button onclick="doHistSearch()">搜</button>
      </div>
      <div id="histResults" class="hint" style="margin-top:6px;border:1px solid #e2e8f0;border-radius:8px;padding:6px 8px;flex:1;overflow:auto;font-size:12.5px"></div>
    </div>
  </div>
</div>
</div>
<script>
let S=null, edits={}, sel=null;
const COLORS={table:"#2f6fdb",note:"#d9a406",mixed:"#d43b3b",ignore:"#888888"};
const LABELS={table:"🟦数据表",note:"🟨文本/批注",mixed:"🟥混合待复核",ignore:"⬜忽略"};

async function loadSheets(){
  const sheets=await (await fetch('/api/sheets')).json();
  SHEETS=sheets;
  renderScope();loadSessions();refreshLLMCfg();
  const sel=document.getElementById('sheetSel');
  sel.innerHTML=sheets.map(s=>`<option value="${esc(s.name)}">${s.visible?'':'[隐] '}${esc(s.name)}</option>`).join('');
  sel.onchange=()=>loadState(sel.value);
  loadState(sheets[0].name);
  refreshStatus();
  applyLang();
  loadTabs();setTimeout(syncTabs,600);
}
async function refreshStatus(){
  const st=await (await fetch('/api/status')).json();
  const el=document.getElementById('status');
  el.textContent=st.status==='frozen'?t('status_frozen_pre')+st.frozen_at:t('status_parsed');
  el.className=st.status==='frozen'?'frozen':'';
  window.COPY=!!st.copy;
  document.getElementById('wbName').textContent=(S?S.workbook.path:'')+(st.copy?'〔副本〕':'');
  document.getElementById('exportBtn').style.display=st.copy?'':'none';
  const bar=document.getElementById('changeBar');
  const ch=st.change;
  if(ch&&ch.type==='structure'){
    bar.style.display='';
    bar.innerHTML=`${t('chg_struct_pre')}${ch.added.length}${t('chg_struct_mid')}${ch.removed.length}${t('chg_struct_mid2')}${ch.changed.length}${t('chg_struct_mid3')}`+
      (ch.added.length?`${t('chg_added')}${ch.added.join('、')}。`:'')+
      (ch.removed.length?`${t('chg_removed')}${ch.removed.join('、')}。`:'')+
      `${t('chg_tail')}${ch.marks_transferred}${t('chg_tail2')}`;
  }else if(ch&&ch.type==='value-only'){
    bar.style.display='';
    bar.innerHTML=t('chg_value');
  }else{bar.style.display='none'}
}
async function loadState(name){
  S=await (await fetch('/api/state?sheet='+encodeURIComponent(name))).json();
  document.getElementById('wbName').textContent=S.workbook.path;
  render();
}
function render(){
  const [r1,c1,r2,c2]=S.range, g=document.getElementById('grid');
  const covered=new Set(), topleft={};
  S.merges.forEach(m=>{topleft[`${m[0]},${m[1]}`]=m;
    for(let r=m[0];r<=m[2];r++)for(let c=m[1];c<=m[3];c++)if(!(r==m[0]&&c==m[1]))covered.add(`${r},${c}`)});
  const hr=new Set(S.hidden_rows), hc=new Set(S.hidden_cols);
  let html='<tr><td class="hcol hrow"></td>';
  for(let c=c1;c<=c2;c++)html+=`<td class="hcol">${hc.has(c)?'⟨隐⟩':''}${colName(c)}</td>`;
  html+='</tr>';
  for(let r=r1;r<=r2;r++){
    html+=`<tr><td class="hrow">${hr.has(r)?'⟨隐⟩':''}${r}</td>`;
    for(let c=c1;c<=c2;c++){
      if(covered.has(`${r},${c}`))continue;
      const cell=S.cells[`${r},${c}`]; const m=topleft[`${r},${c}`];
      const span=m?` rowspan="${m[2]-m[0]+1}" colspan="${m[3]-m[1]+1}"`:'';
      let style='';
      if(cell&&cell.fm&&!String(cell.fm).startsWith('theme'))style=` style="background:#${cell.fm.slice(2)}55"`;
      let txt='';
      if(cell){txt=cell.v;if(cell.pct&&typeof txt==='number')txt=(txt*100).toFixed(1)+'%';
        if(cell.f)txt='ƒ'+txt}
      html+=`<td data-r="${r}" data-c="${c}"${span}${style} title="${esc(String(txt).slice(0,120))}">${esc(String(txt))}</td>`;
    }
    html+='</tr>';
  }
  g.innerHTML=html;
  autoFitCols();
  drawOverlays();
}
function esc(s){return s.replace(/&/g,'&amp;').replace(/</g,'&lt;')}
function colName(c){let s='';while(c>0){s=String.fromCharCode(65+(c-1)%26)+s;c=Math.floor((c-1)/26)}return s}
function hexA(hex,a){a=Math.min(Math.max(a,0),1);
  return hex+Math.round(a*255).toString(16).padStart(2,'0')}
function paintAll(){const base=(+document.getElementById('opSlider').value)/100;
  document.querySelectorAll('.blk').forEach(d=>{
    const quiet=d.dataset.quiet==='1';
    const a=base*(quiet?0.25:1)+(d.dataset.boost==='hover'?0.18:0)
      +(d.classList.contains('sel')?0.25:0);
    d.style.background=hexA(d.dataset.color,a)})}
function drawOverlays(){
  const ov=document.getElementById('overlays');ov.innerHTML='';
  ov.style.display=document.getElementById('showMask').checked?'':'none';
  // 表内批注：黄框几何上落在某个蓝表范围内 → 标注出来，避免"蓝框套黄框"看不懂
  const insideOf={};
  S.blocks.forEach(a=>{
    if(typeOf(a)!=='table')return;
    S.blocks.forEach(b=>{
      if(a===b||typeOf(b)!=='note')return;
      const [ar1,ac1,ar2,ac2]=a.bbox,[br1,bc1,br2,bc2]=b.bbox;
      if(br1>=ar1&&br2<=ar2&&bc1>=ac1&&bc2<=ac2)insideOf[b.id]=a.id;
    });
  });
  const wrap=document.getElementById('gridWrap'), wr=wrap.getBoundingClientRect();
  S.blocks.forEach(b=>{
    const t=typeOf(b), a=anchor(b.bbox), [r1,c1,r2,c2]=b.bbox;
    if(!a)return;
    const d=document.createElement('div');d.className='blk'+(sel===b.id?' sel':'');
    d.dataset.color=COLORS[t];
    if(t==='note'||t==='ignore'){d.dataset.quiet='1';d.style.borderStyle='dashed'}
    const e1=anchorCell(r1,c1), e2=anchorCell(r2,c2)||a;
    if(!e1)return;
    const p=e1.getBoundingClientRect(), q=e2.getBoundingClientRect();
    d.style.cssText+=`left:${p.left-wr.left}px;top:${p.top-wr.top}px;width:${q.right-p.left}px;
      height:${q.bottom-p.top}px;border-color:${COLORS[t]}`;
    const nested=insideOf[b.id]!==undefined;
    d.innerHTML=`<span class="tag" style="color:${COLORS[t]}">${LABELS[t]} #${b.id} ${b.score??''}${nested?' · 表内批注':''}</span>`;
    d.onmouseenter=()=>{d.dataset.boost='hover';paintAll()};
    d.onmouseleave=()=>{d.dataset.boost='';paintAll()};
    d.onclick=()=>{sel=b.id;showDetail(b);drawOverlays()};
    ov.appendChild(d);
  });
  paintAll();
}
function typeOf(b){return (edits[b.id]&&edits[b.id].user_type)||b.type}
function anchor(bbox){return anchorCell(bbox[0],bbox[1])}
function anchorCell(r,c){
  // 精确格 → 若被合并区盖住（这些位置不渲染 td），用合并区左上格（其矩形覆盖该位置）
  const exact=document.querySelector(`td[data-r="${r}"][data-c="${c}"]`);
  if(exact)return exact;
  const m=(S.merges||[]).find(m=>m[0]<=r&&r<=m[2]&&m[1]<=c&&c<=m[3]);
  if(m)return document.querySelector(`td[data-r="${m[0]}"][data-c="${m[1]}"]`);
  const rowTd=document.querySelector(`td[data-r="${r}"][data-c="${c}"]`)
    ||document.querySelector(`td[data-r="${r}"]`);
  if(rowTd)return rowTd;
  // 行超渲染范围（serve.max_rows 限制）→ 钳制到最后渲染行，取该列（或最后格）
  const all=document.querySelectorAll('#grid td[data-r]');
  if(!all.length)return null;
  const lastRow=all[all.length-1].dataset.r;
  return document.querySelector(`td[data-r="${lastRow}"][data-c="${c}"]`)||all[all.length-1];
}
function showHelp(){document.getElementById('helpBox').style.display=''}
const EMAIL="__EMAIL__";


/* 网页内拖入 xlsx → 上传副本打开（桌面壳里由 Rust 拦截为"打开原文件"，此 handler 不触发） */
window.addEventListener('dragover',e=>{e.preventDefault()});
window.addEventListener('drop',async e=>{
  e.preventDefault();
  const f=[...e.dataTransfer.files].find(f=>f.name.toLowerCase().endsWith('.xlsx'));
  if(!f)return;
  const bar=document.getElementById('dropBar')||document.body.appendChild(Object.assign(
    document.createElement('div'),{id:'dropBar'}));
  Object.assign(bar.style,{position:'fixed',left:'50%',transform:'translateX(-50%)',bottom:'14px',
    zIndex:99,background:'#fff7ed',border:'1px solid #fdba74',borderRadius:'8px',
    padding:'8px 14px',fontSize:'12.5px',color:'#92400e'});
  bar.textContent=t('drop_pre')+f.name+t('drop_suf');
  const buf=await f.arrayBuffer();
  const r=await (await fetch('/api/upload?name='+encodeURIComponent(f.name),{method:'POST',body:buf})).json();
  if(!r.ok){bar.textContent='上传失败：'+(r.reason||'');return}
  bar.remove();await reloadForNewFile();syncTabs();
});

const LLM_PRESETS={
  siliconflow:{proto:'openai',base:'https://api.siliconflow.cn/v1',model:'',key:''},
  openai:{proto:'openai',base:'https://api.openai.com/v1',model:'gpt-4o-mini',key:''},
  deepseek:{proto:'openai',base:'https://api.deepseek.com/v1',model:'deepseek-chat',key:''},
  glm:{proto:'openai',base:'https://open.bigmodel.cn/api/paas/v4',model:'glm-4.6',key:''},
  anthropic:{proto:'anthropic',base:'https://api.anthropic.com',model:'claude-sonnet-4-5',key:''},
  ollama:{proto:'openai',base:'http://localhost:11434/v1',model:'qwen3:8b',key:''},
  lmstudio:{proto:'openai',base:'http://localhost:1234/v1',model:'',key:''},
  custom:{proto:'openai',base:'',model:'',key:''}};
function showLLM(){
  loadTraySetting();
  document.getElementById('llmBox').style.display='';
  const st=document.getElementById('llmState');
  st.textContent=(window.__llmCfg&&window.__llmCfg.model)?(window.__llmCfg.model+' @ '+(window.__llmCfg.base_url||'')):t('llm_none');
  if(window.__llmCfg&&window.__llmCfg.model){
    document.getElementById('llmProto').value=window.__llmCfg.protocol||'openai';
    document.getElementById('llmBase').value=window.__llmCfg.base_url||'';
    document.getElementById('llmModel').value=window.__llmCfg.model||'';
    document.getElementById('llmKey').value='';   // 不回显密钥
  }
}
function applyPreset(){
  const p=LLM_PRESETS[document.getElementById('llmPreset').value];
  if(!p)return;
  document.getElementById('llmProto').value=p.proto;
  document.getElementById('llmBase').value=p.base;
  document.getElementById('llmModel').value=p.model;
}

async function loadTraySetting(){
  try{
    const r=await (await fetch('/api/app_settings')).json();
    document.getElementById('trayToggle').checked=!!r.close_to_tray;
  }catch(e){}
}
async function saveTraySetting(v){
  await fetch('/api/app_settings_save',{method:'POST',headers:{'Content-Type':'application/json'},
    body:JSON.stringify({close_to_tray:v})});
}
async function refreshLLMCfg(){
  try{window.__llmCfg=await (await fetch('/api/llm_config')).json()}catch(e){window.__llmCfg=null}
}
async function llmSave(){
  const r=await (await fetch('/api/llm_save',{method:'POST',headers:{'Content-Type':'application/json'},
    body:JSON.stringify({protocol:document.getElementById('llmProto').value,
      base_url:document.getElementById('llmBase').value.trim(),
      model:document.getElementById('llmModel').value.trim(),
      api_key:document.getElementById('llmKey').value.trim()})})).json();
  document.getElementById('llmState').textContent=r.ok?t('llm_saved'):(r.reason||'');
  refreshLLMCfg();
}
async function llmTest(){
  const st=document.getElementById('llmState');st.textContent=t('llm_testing');
  const r=await (await fetch('/api/llm_test',{method:'POST',headers:{'Content-Type':'application/json'},
    body:JSON.stringify({protocol:document.getElementById('llmProto').value,
      base_url:document.getElementById('llmBase').value.trim(),
      model:document.getElementById('llmModel').value.trim(),
      api_key:document.getElementById('llmKey').value.trim()})})).json();
  st.innerHTML=r.ok?`<span style="color:#15803d">✓ ${r.ms}ms · ${esc(String(r.reply).slice(0,30))}</span>`
    :`<span style="color:#dc2626">✗ ${esc(String(r.reason).slice(0,90))}</span>`;
}
function showOpen(){document.getElementById('openBox').style.display='';
  document.getElementById('openErr').textContent='';
  document.getElementById('openPath').value=S?S.workbook.path:''}
async function doOpenPath(){
  const p=document.getElementById('openPath').value.trim().replace(/^"+|"+$/g,'');
  const err=document.getElementById('openErr');
  if(!p){err.textContent=t('open_err_empty');return}
  err.textContent=t('open_parsing');
  const r=await (await fetch('/api/open',{method:'POST',headers:{'Content-Type':'application/json'},
    body:JSON.stringify({path:p})})).json();
  if(!r.ok){err.textContent=r.reason||'打开失败';return}
  document.getElementById('openBox').style.display='none';
  await reloadForNewFile();syncTabs();
}
async function doOpenUpload(){
  const f=document.getElementById('openFile').files[0];if(!f)return;
  const st=document.getElementById('openUpState');st.textContent='上传中…';
  const buf=await f.arrayBuffer();
  const r=await (await fetch('/api/upload?name='+encodeURIComponent(f.name),{method:'POST',
    body:buf})).json();
  if(!r.ok){st.textContent=r.reason||'上传失败';return}
  st.textContent='解析中…';
  await reloadForNewFile();syncTabs();
}
function showAbout(){
  const zh=LANG==='zh';
  document.getElementById('aboutBody').innerHTML=zh?`
   <div style="display:flex;align-items:baseline;gap:10px;border-bottom:2px solid #e2e8f0;padding-bottom:10px;margin-bottom:12px">
     <b style="font-size:22px;letter-spacing:2px">格物 <span style="color:#2563eb">Gewu</span></b>
     <span class="hint" style="font-size:12px">复杂表格理解工作台</span>
     <button onclick="document.getElementById('aboutBox').style.display='none'" style="margin-left:auto">关闭 ✕</button></div>
   <p style="margin:6px 0">格物致知——穷究每一格，理解整张表。</p>
   <p style="margin:6px 0;color:#475569">格物是<b>本地优先</b>的复杂表格理解工作台：机器解析结构，人工一次确认，随后即可问答、解读与安全改数。除显式改数外<b>永不修改原文件</b>；语义镜像、口径词典与对话会话全部保存在你自己的电脑上。</p>
   <div style="margin:14px 0 4px;padding:10px 12px;background:#f8fafc;border-left:3px solid #2563eb;border-radius:6px">
     <b>识玥 shyue</b><br>
     <span style="color:#64748b">📧 __EMAIL__</span><br>
     <span class="hint" style="font-size:11px">Gewu/1.0</span>
   </div>`:`
   <div style="display:flex;align-items:baseline;gap:10px;border-bottom:2px solid #e2e8f0;padding-bottom:10px;margin-bottom:12px">
     <b style="font-size:22px;letter-spacing:2px">格物 <span style="color:#2563eb">Gewu</span></b>
     <span class="hint" style="font-size:12px">Intelligence workbench for complex spreadsheets</span>
     <button onclick="document.getElementById('aboutBox').style.display='none'" style="margin-left:auto">Close ✕</button></div>
   <p style="margin:6px 0">Investigate every cell; understand the whole sheet.</p>
   <p style="margin:6px 0;color:#475569">Gewu is a <b>local-first</b> workbench for understanding complex spreadsheets: the machine parses structure, you confirm once, then ask, interpret and safely edit. It <b>never touches your file</b> except explicit edit commands; the semantic mirror, glossary and chat sessions all stay on your own machine.</p>
   <div style="margin:14px 0 4px;padding:10px 12px;background:#f8fafc;border-left:3px solid #2563eb;border-radius:6px">
     <b>识玥 shyue</b><br>
     <span style="color:#64748b">📧 __EMAIL__</span><br>
     <span class="hint" style="font-size:11px">Gewu/1.0</span>
   </div>`;
  document.getElementById('aboutBox').style.display='';
}
function showDetail(b){
  const t=typeOf(b), d=document.getElementById('detail');
  const h=b.header||{};
  let rows=`
   <b>类型</b><span>${LABELS[t]}</span>
   <b>位置</b><span>r${b.bbox[0]}-r${b.bbox[2]} × c${b.bbox[1]}-c${b.bbox[3]}</span>
   <b>代码置信</b><span>${b.score}</span>`;
  if(h.header_rows)rows+=`<b>表头行</b><span>r${h.header_rows.join(',r')}</span>`;
  if(h.title_text)rows+=`<b>标题</b><span>${esc(h.title_text)}</span>`;
  if(h.floating_rows&&h.floating_rows.length)rows+=`<b>浮动草稿区</b><span>r${h.floating_rows.join(',r')}</span>`;
  if(h.summary_rows&&h.summary_rows.length)rows+=`<b>⚠汇总行</b><span>r${h.summary_rows.join(',r')}（写保护）</span>`;
  if(h.dup_header_names&&h.dup_header_names.length)rows+=`<b>⚠同名列</b><span>${h.dup_header_names.join('、')}</span>`;
  if(h.cross_tab)rows+=`<b>交叉表</b><span>是</span>`;
  if(h.col_tree)rows+=`<b>字段</b><span>${h.col_tree.filter(n=>n.path&&n.path.length).length} 个叶子列</span>`;
  if(b.signals&&!b.annotation)rows+=`<b>信号</b><span>文本${b.signals.text_r} 数值${b.signals.num_r} 边框${b.signals.perimeter_border}</span>`;
  if(b.annotation)rows+=`<b>批注</b><span>是（已从表格剥离）</span>`;
  if(b.attach){
    rows+=`<b>行内备注</b><span>挂到 表#${b.attach.table_block_id} 第${b.attach.row_no}行<br>
      <span class="hint">${esc((b.attach.text||'').slice(0,80))}</span><br>
      <button onclick="doDetach(${b.id})">取消挂接（不算这行的信息）</button></span>`;
  }
  const switchTitle=(t==='note')
    ?'该块引擎判定为文本（无列结构）：切为数据表仅作视觉标记，不会进入问答'
    :'把这块按另一种类型对待；点冻结才生效，可反复切换';
  rows+=`<b>操作</b><span>
    <button onclick="setEdit(${b.id},'user_type','${nextType(t)}')" title="${switchTitle}">切为${LABELS[nextType(t)]}</button>
    <button onclick="setEdit(${b.id},'readonly',1)" title="应用内只读：本工具禁止改这块的数（问答不受影响）；文件本身不变，Excel/WPS 里仍可编辑">🔒只读</button>
    <button onclick="setEdit(${b.id},'readonly',0)" title="解除只读">🔓可写</button></span>`;
  if(edits[b.id]){
    const e=edits[b.id];const parts=[];
    if(e.user_type)parts.push('切为'+LABELS[e.user_type]);
    if(e.readonly!==undefined)parts.push(e.readonly?'设为只读🔒':'设为可写🔓');
    rows+=`<b>已改(未冻结)</b><span>${parts.join('，')} <button onclick="revertEdit(${b.id})" title="撤销这块的全部未冻结修改">↩ 撤销本块修改</button></span>`;
  }
  d.innerHTML=rows;
  if(t==='table'){renderGlossary(b);renderRules()}
}

async function renderRules(){
  const box=document.getElementById('rulesBox');
  const cols=[...new Set((sel!==null&&S.blocks.find(b=>b.id===sel&&b.type!=='note')?.header?.col_tree||[])
    .filter(n=>n.path&&n.path.length).map(n=>n.path[n.path.length-1]))].slice(0,30);
  if(!cols.length){box.style.display='none';return}
  box.style.display='';
  document.getElementById('ruleCol').innerHTML=cols.map(c=>`<option>${esc(c)}</option>`).join('');
  const rs=await (await fetch('/api/rules?sheet='+encodeURIComponent(S.sheet))).json();
  document.getElementById('rulesList').innerHTML=rs.length?rs.map(r=>
    `· <b>${esc(r.col_path.split('/').pop())}</b> [${r.rmin??'-∞'},${r.rmax??'+∞'}] ${r.severity}`).join('<br>')
    :'（本 sheet 尚无规则）';
}
async function saveRule(){
  await fetch('/api/rules_save',{method:'POST',headers:{'Content-Type':'application/json'},
    body:JSON.stringify({sheet:S.sheet,column:document.getElementById('ruleCol').value,
      min:document.getElementById('ruleMin').value,max:document.getElementById('ruleMax').value,
      pattern:document.getElementById('rulePat').value.trim(),
      severity:document.getElementById('ruleSev').value})});
  renderRules();
}
async function delRule(){
  await fetch('/api/rules_delete',{method:'POST',headers:{'Content-Type':'application/json'},
    body:JSON.stringify({sheet:S.sheet,column:document.getElementById('ruleCol').value})});
  renderRules();
}
async function runRulesCheck(){
  const out=await (await fetch('/api/rules_check')).json();
  document.getElementById('rulesCheckOut').innerHTML=out.length?
    `<b style="color:#b91c1c">✗ ${out.length} 处违规：</b><br>`+out.slice(0,20).map(x=>
      `· [${esc(x.sheet)}] r${x.row} 「${esc(x.column)}」=${x.value} ${esc(x.问题)}(${x.severity})`).join('<br>')
    :'<span style="color:#15803d">✓ 现有数据全部符合规则</span>';
}
async function renderGlossary(b){
  const h=b.header||{};
  const cols=[...new Set((h.col_tree||[]).filter(n=>n.path&&n.path.length)
    .map(n=>n.path[n.path.length-1]))].slice(0,25);
  let gl=[];
  try{gl=await (await fetch('/api/glossary?sheet='+encodeURIComponent(S.sheet))).json();}catch(e){}
  const box=document.getElementById('glossaryBox');
  if(!cols.length){box.style.display='none';return}
  box.style.display='';
  const byCol={};gl.forEach(g=>byCol[g.term]=g.meaning);
  box.querySelector('#glossaryList').innerHTML=cols.map(c=>`
    <div style="display:flex;gap:3px;align-items:center;margin:2px 0">
      <span style="flex:0 0 38%;overflow:hidden;text-overflow:ellipsis;white-space:nowrap"
            title="${esc(c)}">${esc(c)}</span>
      <input id="gl_${b.id}_${cols.indexOf(c)}" style="flex:1;font:11px system-ui;padding:2px 4px"
             placeholder="口径注释…" value="${esc(byCol[c]||'')}">
      <button onclick="saveGlossary('${esc(c)}',${b.id},${cols.indexOf(c)})">存</button>
    </div>`).join('')+
    `<div class="hint" style="margin-top:4px">口径注释会自动注入问答与写指令的模型上下文（同名不同口径的列靠它消歧）</div>`;
}
async function saveGlossary(col,bid,idx){
  const v=document.getElementById(`gl_${bid}_${idx}`).value.trim();
  if(!v)return;
  const r=await (await fetch('/api/glossary_save',{method:'POST',
    headers:{'Content-Type':'application/json'},
    body:JSON.stringify({sheet:S.sheet,column:col,meaning:v})})).json();
  if(r.ok)document.getElementById(`gl_${bid}_${idx}`).style.background='#dcfce7';
}
function nextType(t){return {table:'note',note:'mixed',mixed:'ignore',ignore:'table'}[t]}
function setEdit(id,k,v){
  edits[id]=edits[id]||{};
  if(!edits[id]._orig){const b0=S.blocks.find(x=>x.id===id);edits[id]._orig={type:b0?b0.type:'table'}}
  edits[id][k]=v;
  if(k==='user_type'){const b=S.blocks.find(x=>x.id===id);if(b){b.type=v}}
  drawOverlays();updatePending();
  const b=S.blocks.find(x=>x.id===id);if(b)showDetail(b);
}
function revertEdit(id){          // 误触保护：撤销本块全部未冻结修改
  const e=edits[id];if(!e)return;
  const b=S.blocks.find(x=>x.id===id);
  if(b&&e._orig)b.type=e._orig.type;
  delete edits[id];
  drawOverlays();updatePending();
  if(b)showDetail(b);
}
function updatePending(){
  const n=Object.keys(edits).length;
  const chip=document.getElementById('pendingChip');
  chip.style.display=n?'':'none';
  chip.textContent=`${t('pending_pre')}${n}${t('pending_suf')}`;
  document.getElementById('freezeBtn').textContent=n?`${t('freeze_n_pre')}${n}${t('freeze_n_suf')}`:t('freeze');
}
async function doFreeze(){
  const list=[];
  S.blocks.forEach(b=>{if(edits[b.id]&&b.db_id)
    list.push({block_db_id:b.db_id,...edits[b.id]})});
  const r=await (await fetch('/api/freeze',{method:'POST',headers:{'Content-Type':'application/json'},
    body:JSON.stringify({edits:list,sheet:S.sheet})})).json();
  if(r.ok){edits={};updatePending();}
  alert(r.ok?t('frozen_ok'):t('frozen_fail'));refreshStatus();
}
async function doDetach(blockId){
  const r=await (await fetch('/api/note_detach',{method:'POST',headers:{'Content-Type':'application/json'},
    body:JSON.stringify({sheet:S.sheet,block_id:blockId})})).json();
  if(r.ok){S=await (await fetch('/api/state?sheet='+encodeURIComponent(S.sheet))).json();
    sel=blockId;const b=S.blocks.find(x=>x.id===blockId);if(b)showDetail(b);drawOverlays();}
  else alert('取消失败：'+(r.reason||''));
}
function esc(s){return String(s).replace(/[&<>"']/g,m=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[m]))}

/* ===== 列宽：初次内容自适应 + 表头右缘拖拽 + 双击复位 ===== */
const _measure=(()=>{const c=document.createElement('canvas').getContext('2d');
  c.font='12px system-ui';return t=>Math.ceil(c.measureText(String(t)).width)})();
function colWKey(c){return (S&&S.workbook?S.workbook.path.slice(-40):'wb')+'|'+(S?S.sheet:'')+'|c'+c}
function loadColW(){try{return JSON.parse(localStorage.getItem('gewu-colw')||'{}')}catch(e){return{}}}
function saveColW(o){localStorage.setItem('gewu-colw',JSON.stringify(o))}
function autoFitCols(){
  const g=document.getElementById('grid');
  if(!g.rows.length)return;
  const [r1,c1,r2,c2]=S.range, hr=new Set(S.hidden_cols);
  const ov=loadColW();
  const widths={};let total=0;
  for(let c=c1;c<=c2;c++){
    if(hr.has(c)){widths[c]=40;total+=40;continue}
    let w=54;                                  // 列头宽底
    for(const tr of g.rows){for(const td of tr.cells){
      if(td.cellIndex-1===c-c1 && !td.classList.contains('hcol') && !td.classList.contains('hrow')){
        const t=td.textContent;if(!t)continue;
        w=Math.max(w,_measure(t)+18);          // 内容自适应（含 padding）
      }}}
    w=Math.min(w,320);
    const k=colWKey(c);
    if(ov[k])w=ov[k];                          // 手动调过的列尊重用户
    widths[c]=w;total+=w;
  }
  const tmp=document.createElement('table');    // colgroup 需在 table 上下文里解析
  tmp.innerHTML=`<colgroup><col style="width:44px">`+
    Array.from({length:c2-c1+1},(_,i)=>`<col data-c="${c1+i}" style="width:${widths[c1+i]}px">`).join('')+`</colgroup>`;
  g.insertBefore(tmp.firstChild,g.firstChild);
  g.querySelectorAll('td.hrow').forEach(td=>{td.style.width='44px'});
  g.style.width=(44+total)+'px';
  // 表头挂拖拽/双击复位
  g.querySelectorAll('td.hcol:not(.hrow)').forEach(td=>{
    td.style.cursor='col-resize';
    td.onmousedown=e=>{ if(e.offsetX>=td.offsetWidth-6) dragCol(td,e) };
    td.ondblclick=()=>{ const c=c1+td.cellIndex-1,k=colWKey(c),ov=loadColW();
      delete ov[k];saveColW(ov);autoFitCols();drawOverlays(); };
  });
}
function dragCol(td,e0){
  const c=S.range[1]+td.cellIndex-1,k=colWKey(c);
  const x0=e0.clientX,w0=td.offsetWidth;
  const mv=e=>{const w=Math.max(40,Math.min(600,w0+e.clientX-x0));
    document.querySelector(`#grid colgroup col[data-c="${c}"]`).style.width=w+'px';
    document.getElementById('grid').style.width=(parseInt(document.getElementById('grid').style.width)-w0+w)+'px';_w=w};
  let _w=w0;
  const up=()=>{window.removeEventListener('mousemove',mv);window.removeEventListener('mouseup',up);
    const ov=loadColW();ov[k]=_w;saveColW(ov);drawOverlays()};
  window.addEventListener('mousemove',mv);window.addEventListener('mouseup',up);
  e0.preventDefault();
}

function mdRich(s){
  /* 富文本渲染：先转义再逐级替换（安全：输入先 esc）。
     支持：**加粗** · `行内代码` · ```代码块``` · =公式= · > 引用 · - 列表 · 数字列表 · 表格行 | */
  let t=esc(s);
  // 代码块（三反引号）
  t=t.replace(/```(\w*)\n?([\s\S]*?)```/g,(m,lang,code)=>{
    return `<pre class="mdPre"><code>${code.trim()}</code></pre>`});
  // Excel 公式：=SUM(...) / =(A1+B2)/C3 / ==SUM(A1:A10)——等宽蓝底
  t=t.replace(/={1,2}\(?[A-Z][A-Z0-9]*(?:\([^)]*\))?[^<\s=]{0,40}\)?/g,
    m=>{ if(m.length<3||m.length>100)return m;
         return `<span class="mdFormula">${m}</span>`});
  // 行内代码
  t=t.replace(/`([^`\n]+)`/g,'<code class="mdCode">$1</code>');
  // 加粗
  t=t.replace(/\*\*([^*\n]+)\*\*/g,'<b class="mdBold">$1</b>');
  // 列表行（- 开头）→ 圆点
  t=t.replace(/^[\s]*[-·]\s+(.+)$/gm,'<span class="mdLi">· $1</span>');
  // 数字列表行（1. 2. 开头）
  t=t.replace(/^[\s]*(\d+)[.、]\s+(.+)$/gm,'<span class="mdOl"><b>$1.</b> $2</span>');
  // 引用行（> 开头；esc 后 &gt;）
  t=t.replace(/^[\s]*(?:>|&gt;)\s*(.+)$/gm,'<span class="mdQuote">$1</span>');
  // 表格行（| 分隔）→ 简单表格样式
  t=t.replace(/^[\s]*(\|[^|\n]+\|[^|\n]*)$/gm,'<span class="mdTr">$1</span>');
  // 换行保留
  t=t.replace(/\n/g,'<br>');
  return t;
}
// 兼容旧调用
function mdLite(s){return mdRich(s)}

/* ===== 双语字典（zh 为默认原文；en 为英文版）===== */
const I18N={
 "zh":{
  brand_tag:"复杂表格理解工作台", mask:"蒙版", opacity:"浓度",
  freeze:"✔ 检查无误，冻结", help_t:"注意事项", about_t:"关于格物", open_t:"打开文件",
  export_btn:"📥 导出改后文件", export_t:"下载改后的文件（副本模式的改动只有导出才是成品）",
  open_title:"📂 打开文件", open_path_h:"① 文件路径",
  open_path_hint:"在资源管理器按住 Shift 右键文件 →「复制文件地址」粘贴到这。改数直接写回这个原文件。",
  open_ph:"C:\\Users\\...\\表.xlsx", open_do:"打开",
  open_copy_h:"② 上传副本",
  open_copy_note:"（分析体验用；改数只写副本——成品须点页眉「📥 导出改后文件」下载，别把原上传文件当成品发人）",
  open_up:"上传并打开", open_err_empty:"请输入文件路径", open_parsing:"解析中…（大文件可能十几秒）",
  drop_pre:"📥 收到 ", drop_suf:"：网页拖入=上传副本（改动只进副本，成品用页眉「导出改后文件」）；要写回原文件请用 📂 路径打开。上传解析中…", collapse:"收起面板（表格全宽）",
  tab_review:"📋 审阅", tab_chat:"💬 问答", tab_write:"✏️ 改数", tab_history:"🕘 历史",
  review_h:"区块审阅", review_hint:"点色块查看详情；不确定怎么操作就点右上 ❓——大多数文件直接冻结即可。",
  glossary:"📖 字段口径词典", scope:"范围", whole_file:"整个文件",
  sess_hist:"历史会话列表", sess_new:"新对话", cur_title_empty:"（提问后自动建会话）",
  chat_hint:"问数值（合计/最高的）→ 引擎计算；问理解（这表是做什么的/为什么这么算）→ LLM 解读；问评价（分析一下）→ 数据解读。追问自动带最近 4 轮上下文。",
  ask_ph:"问：销售额合计 / 这表是做什么的 / 分析评价一下 （回车发送）", ask_btn:"问",
  write_h:"改数（写回 Excel）", write_note:"说一句话定位到某行某列 → 机器算新值 → 看差异 → 确认才写入（自动备份，可撤销）。改数不认对话范围，按指令语义全簿定位。",
  write_ph:"自然语言：把1月云澜居的成交套数改成7\n或手写计划 JSON：{\"find\":{\"keys\":{\"列\":\"值\"}},\"set\":{\"column\":\"列\",\"value\":7}}",
  btn_nl:"💬 自然语言", btn_plan:"📋 计划 JSON", btn_undo:"↩ 撤销一批",
  diff_hint:"语义解析 → 硬校验（冻结/表头/汇总行等全不许碰）→ diff 预览 → 确认执行",
  write_log:"最近写入", write_log_empty:"（切到本页签时加载）", write_log_none:"（本文件还没有写入记录）",
  hist_h:"本文件的会话", hist_clear:"清空全部", hist_clear_confirm:"确定清空本文件全部会话？此操作不可撤销。", hist_search_h:"🔎 搜全部文件的历史",
  hist_ph:"关键词（跨工作簿检索，只读参考）", hist_none:"（本文件尚无会话）", hist_empty:"（尚无会话，提问后自动创建）",
  status_parsed:"未冻结（解析态）", status_frozen_pre:"已冻结 ", undone:"（已撤销）",
  pending_pre:"✏️ 待提交修改 ", pending_suf:" 处（点冻结才生效；刷新页面=全部撤销）",
  freeze_n_pre:"✔ 检查无误，冻结（含 ", freeze_n_suf:" 处修改）",
  frozen_ok:"已冻结：语义镜像已写入 mirror.db，下游问答/写回将只认此版本", frozen_fail:"失败",
  new_chat:"已开新对话。", sess_deleted:"会话已删除。", sess_empty:"（空会话）",
  chg_struct_pre:"⚠ <b>文件内容有变化（结构级）</b>：新增 ", chg_struct_mid:" 块、消失 ", chg_struct_mid2:" 块、结构变动 ", chg_struct_mid3:" 块。",
  chg_added:" 新增：", chg_removed:" 消失：", chg_tail:" 未变化块的人工标记已自动保留（", chg_tail2:" 处），<b>请复核变化后重新冻结</b>；冻结前改数仍被拦截。",
  chg_value:"✅ 检测到文件数值更新（结构无变化）：冻结的表结构判定仍然有效，数值已自动刷新。",
  hit_row:"命中行：", dep_pre:"本次写入将联动重算 ", dep_suf:" 个公式格（Excel 自动重算）",
  dep_done_pre:"联动公式格已重算：", dep_done_suf:"（切回表格可见新值）",
  rules_h:"📏 业务规则（值域）", rules_note:"给数值列设上下限：block=越界拒写，warn=仅警告。规则由引擎硬执行，LLM 只知情。",
  rules_save:"保存", rules_del:"删除", rules_check:"全表体检",
  rules_check_note:"按现有数据跑一遍规则，列出违规行",
  rules_min_ph:"≥", rules_max_ph:"≤",
  btn_nl_t:"需在 ⚙️ 配置 LLM（本地模型亦可）", dep_unchanged:"⚠ 预测的联动格缓存值未变化——若预期应重算，请打开 Excel 确认", rules_pat_ph:"正则(选填)",
  llm_t:"LLM 设置", llm_h:"⚙️ LLM 设置", llm_close:"关闭 ✕", llm_preset_h:"预设",
  llm_proto:"协议", llm_base:"接口地址", llm_model:"模型名", llm_key:"API Key",
  llm_privacy:"密钥只保存在本机 llm.json，请求直连你所选的服务商；不配置则问答走内置纯代码引擎。",
  llm_save:"保存", llm_test:"测试连接", llm_testing:"测试中…",
  llm_saved:"已保存（新对话生效）", llm_none:"（未配置：问答走内置规则引擎）",
  tray_label:"关闭窗口时缩到系统托盘（引擎保持运行，重开秒进；取消勾选=直接退出）",
  tray_note:"仅桌面应用生效，改后重启应用即可。", copy_badge:"副本",
  guide:'<b>怎么用：</b><b>蓝框=表格</b>（问答和改数都发生在蓝框里，只需扫一眼框得对不对）；<b>黄虚线=标题/备注</b>（不当表格处理，<u>不用管它</u>；标了「表内批注」的是表格中间夹的一行备注，正常现象）；<b>红框=机器拿不准</b>（最好点开看一眼）。没问题 → 点右上「检查无误，冻结」→ 右侧面板开始问答与改数。框不对 → 点色块改类型。<button onclick="this.parentNode.remove()" style="margin-left:auto">知道了 ✕</button>'
 },
 "en":{
  brand_tag:"Intelligence workbench for complex spreadsheets", mask:"Mask", opacity:"Opacity",
  freeze:"✔ All good — Freeze", help_t:"Notes", about_t:"About Gewu", open_t:"Open file",
  export_btn:"📥 Export edited file", export_t:"Download the edited file (in copy mode, only the export is the finished product)",
  open_title:"📂 Open file", open_path_h:"1. File path",
  open_path_hint:"Shift + right-click the file in Explorer → “Copy as path”, then paste here. Edits write back to this original file.",
  open_ph:"C:\\Users\\...\\file.xlsx", open_do:"Open",
  open_copy_h:"2. Upload a copy",
  open_copy_note:"(for analysis; edits go to the copy only — the finished product must be downloaded via “📥 Export edited file”; never send the originally uploaded file as your result)",
  open_up:"Upload & open", open_err_empty:"Please enter a file path", open_parsing:"Parsing… (large files may take a while)",
  drop_pre:"📥 Received ", drop_suf:": dropping onto the web page uploads a COPY (edits stay in the copy; use “Export edited file” for the result). To edit the original, open it by path via 📂. Uploading…", collapse:"Collapse panel (full-width grid)",
  tab_review:"📋 Review", tab_chat:"💬 Chat", tab_write:"✏️ Edit", tab_history:"🕘 History",
  review_h:"Block review", review_hint:"Click a colored block for details; most files can be frozen directly — see ❓ if unsure.",
  glossary:"📖 Field glossary", scope:"Scope", whole_file:"Whole file",
  sess_hist:"Session history", sess_new:"New chat", cur_title_empty:"(session starts on first question)",
  chat_hint:"Numbers (sum / top) → engine computes · Understanding (what is this table / why this formula) → LLM explains · Evaluation (analyze this) → data interpretation. Follow-ups carry the last 4 turns.",
  ask_ph:"Ask: total of sales / what is this table / analyze this  (Enter to send)", ask_btn:"Ask",
  write_h:"Edit cells (writes back to Excel)", write_note:"Say which row & column → engine computes the new value → you review the diff → confirm to write (auto-backup, undoable). Edits ignore chat scope and locate semantically across the workbook.",
  write_ph:"Natural language: set Sales of Project X in Jan to 7\nOr plan JSON: {\"find\":{\"keys\":{\"col\":\"value\"}},\"set\":{\"column\":\"col\",\"value\":7}}",
  btn_nl:"💬 Natural language", btn_plan:"📋 Plan JSON", btn_undo:"↩ Undo batch",
  diff_hint:"Semantic parse → hard validation (frozen/header/summary rows are protected) → diff preview → confirm to execute",
  write_log:"Recent writes", write_log_empty:"(loads when this tab opens)", write_log_none:"(no writes yet for this file)",
  hist_h:"Sessions of this file", hist_clear:"Clear all", hist_clear_confirm:"Clear ALL sessions of this file? This cannot be undone.", hist_search_h:"🔎 Search history across all files",
  hist_ph:"Keyword (cross-workbook, read-only)", hist_none:"(no sessions yet for this file)", hist_empty:"(no sessions yet — created on first question)",
  status_parsed:"Not frozen (parsed)", status_frozen_pre:"Frozen ", undone:" (undone)",
  pending_pre:"✏️ Pending edits: ", pending_suf:" (applied on freeze; refresh discards all)",
  freeze_n_pre:"✔ All good — Freeze (", freeze_n_suf:" edits)",
  frozen_ok:"Frozen: the semantic mirror is now the single source of truth for Q&A and write-back", frozen_fail:"Failed",
  new_chat:"New conversation started.", sess_deleted:"Session deleted.", sess_empty:"(empty session)",
  chg_struct_pre:"⚠ <b>File changed (structural)</b>: ", chg_struct_mid:" added, ", chg_struct_mid2:" removed, ", chg_struct_mid3:" changed blocks.",
  chg_added:" Added: ", chg_removed:" Removed: ", chg_tail:" Manual marks on unchanged blocks were kept (", chg_tail2:"). <b>Please review and re-freeze</b>; edits stay blocked until then.",
  chg_value:"✅ Values updated (no structural change): frozen structure remains valid; numbers refreshed automatically.",
  hit_row:"Matched row:", dep_pre:"This edit will recalculate ", dep_suf:" dependent formula cells (Excel recalcs automatically)",
  dep_done_pre:"Dependent formulas recalculated: ", dep_done_suf:" (switch to the grid to see new values)",
  rules_h:"📏 Business rules (value ranges)", rules_note:"Set min/max for numeric columns: block = reject out-of-range writes, warn = warn only. Rules are engine-enforced; the LLM is only informed.",
  rules_save:"Save", rules_del:"Delete", rules_check:"Check all data",
  rules_check_note:"Run rules against existing data and list violations",
  rules_min_ph:"min", rules_max_ph:"max",
  btn_nl_t:"Requires LLM in ⚙️ (local models also work)", dep_unchanged:"⚠ Predicted dependent cells' cached values unchanged — open Excel to verify if recalc was expected", rules_pat_ph:"regex (optional)",
  llm_t:"LLM settings", llm_h:"⚙️ LLM Settings", llm_close:"Close ✕", llm_preset_h:"Preset",
  llm_proto:"Protocol", llm_base:"Base URL", llm_model:"Model", llm_key:"API Key",
  llm_privacy:"The key is stored only in the local llm.json; requests go directly to your chosen provider. Without a key, Q&A runs on the built-in rules engine.",
  llm_save:"Save", llm_test:"Test connection", llm_testing:"Testing…",
  llm_saved:"Saved (takes effect on next question)", llm_none:"(not configured: Q&A uses the built-in rules engine)",
  tray_label:"Minimize to system tray on close (engine keeps running, instant reopen; uncheck = quit directly)",
  tray_note:"Desktop app only. Restart the app to apply.", copy_badge:"copy",
  help_full:`<div style="display:flex;align-items:center;border-bottom:2px solid #e2e8f0;padding-bottom:8px;margin-bottom:4px"><b style="font-size:16.5px">Notes</b><button onclick="document.getElementById('helpBox').style.display='none'" style="margin-left:auto">Close ✕</button></div>
  <div style="margin:12px 0 4px;font-weight:700;color:#1d4ed8">1. What freezing means</div>
  <div style="padding-left:10px">Freeze = you confirm "the machine's frames are right", commit your manual fixes, and open the master switch for <b>cell edits</b>.<br><b>Not frozen = read-only mode</b>: Q&amp;A / table interpretation / tie-checking all work; every edit is blocked and the file is absolutely safe.</div>
  <div style="margin:14px 0 4px;font-weight:700;color:#1d4ed8">2. How to ask</div>
  <div style="padding-left:10px">"This table / current table" refers to the sheet you are viewing; "the whole file / how do these tables relate" asks about the whole workbook.<br>Follow-ups (e.g. "then why is it computed that way") automatically carry the last 4 turns; click <b>🗑 New chat</b> to start over.</div>
  <div style="margin:14px 0 4px;font-weight:700;color:#1d4ed8">3. How to edit cells</div>
  <div style="padding-left:10px">Edits <b>do not follow the chat scope</b> — they locate semantically across the whole workbook by sheet keyword + row keys + column name, so include the sheet name ("add 3 to the 250 weight of <b>Chenhong-Korea summary</b>"). Multiple hits are blocked until clarified.<br>Flow: semantic parse → hard validation → <b>diff preview (old → new)</b> → your confirmation → write (auto-backup, undoable). In uploaded-copy mode edits go to the copy only — download the finished file via “📥 Export edited file” in the header.</div>
  <div style="margin:14px 0 4px;font-weight:700;color:#1d4ed8">4. Block buttons (effective after freeze)</div>
  <table style="width:100%;border-collapse:collapse;font-size:13px;margin:4px 0"><tr style="background:#f1f5f9"><th style="text-align:left;padding:4px 8px;border:1px solid #e2e8f0">Button</th><th style="text-align:left;padding:4px 8px;border:1px solid #e2e8f0">Effect</th></tr><tr><td style="padding:4px 8px;border:1px solid #e2e8f0">→ 🟨 Text/note</td><td style="padding:4px 8px;border:1px solid #e2e8f0">The block leaves Q&amp;A and editing (data stays in the file; just not treated as a table)</td></tr><tr><td style="padding:4px 8px;border:1px solid #e2e8f0">→ ⬜ Ignore</td><td style="padding:4px 8px;border:1px solid #e2e8f0">Same as above: out of Q&amp;A + edits rejected</td></tr><tr><td style="padding:4px 8px;border:1px solid #e2e8f0">→ 🟥 Mixed review</td><td style="padding:4px 8px;border:1px solid #e2e8f0">Isolated: no edits, no Q&amp;A, until re-typed as table or text</td></tr><tr><td style="padding:4px 8px;border:1px solid #e2e8f0">🔒 Read-only</td><td style="padding:4px 8px;border:1px solid #e2e8f0"><b>In-app only</b>: this tool refuses to edit the block (Q&amp;A unaffected). <b>The file itself is untouched</b> — still editable in Excel/WPS, file attributes unchanged</td></tr><tr><td style="padding:4px 8px;border:1px solid #e2e8f0">🔓 Writable</td><td style="padding:4px 8px;border:1px solid #e2e8f0">Clears the in-app read-only mark</td></tr></table>
  <div style="padding-left:10px;color:#475569;font-size:12.5px">Accidental clicks: edits apply only on freeze (a "pending N" chip shows at the top); refreshing before freeze discards everything; each block has "↩ revert"; if you froze something wrong, switch it again and re-freeze.<br>Note: switching an engine-detected text block to "table" is a visual mark only (no column structure, never enters Q&amp;A).</div>
  <div style="margin:14px 0 4px;font-weight:700;color:#1d4ed8">5. After the file changes</div>
  <div style="padding-left:10px">Edits made by this tool → reopening continues seamlessly (freeze kept).<br>Someone changed values → re-drop the file: green bar "values refreshed", freeze still valid, no re-confirmation.<br>Structure changed (tables added/removed, headers edited) → re-drop: orange bar lists the changes; marks on unchanged blocks are kept automatically; review then re-freeze.</div>
  <div style="margin-top:14px;padding:8px 10px;background:#f8fafc;border-left:3px solid #94a3b8;color:#64748b;font-size:12.5px">Every edit is re-validated before writing; headers, summary rows, formula cells and ignored blocks are never written.   <div style="margin:14px 0 4px;font-weight:700;color:#1d4ed8">6. Opening files & copies</div>
  <div style="padding-left:10px">
    Three ways to open: <b>📂 path</b> (Shift+right-click → Copy as path — writes back to the original, recommended);
    <b>upload a copy</b> (for analysis — edits go to the copy only, a〔copy〕badge appears in the header);
    drop a file onto the desktop app icon / into the desktop window (= original); drop onto the web page (= uploads a copy).<br>
    <b>In copy mode the finished file = the one downloaded via “📥 Export edited file”</b> — the originally uploaded file never contains your edits; don't send it to anyone.
  </div>
  <div style="margin-top:14px;padding:8px 10px;background:#f8fafc;border-left:3px solid #94a3b8;color:#64748b;font-size:12.5px">
    Every edit is re-validated before writing; headers, summary rows, formula cells and ignored blocks are never written. All manual marks (types/read-only/glossary/row notes) live only in this tool's mirror and <b>never modify your file</b> (except explicit edit commands).</div>`,
   guide:'<b>How to use:</b> <b>Blue = table</b> (Q&A and edits happen inside blue frames — just glance they are right); <b>dashed yellow = title/notes</b> (not treated as a table; blocks tagged “in-table note” are remark rows inside a table — normal); <b>red = machine unsure</b> (worth a look). All good → click “Freeze” top-right → start in the right panel. Wrong frame → click the block and switch its type. <button onclick="this.parentNode.remove()" style="margin-left:auto">Got it ✕</button>'
 }
};
let LANG=localStorage.getItem('gewu-lang')||'zh';
function t(k){const d=I18N[LANG]&&I18N[LANG][k];return d!==undefined?d:(I18N.zh[k]??k)}
function applyLang(){
  document.documentElement.lang=LANG==='en'?'en':'zh';
  document.querySelectorAll('[data-i18n]').forEach(el=>{el.textContent=t(el.dataset.i18n)});
  document.querySelectorAll('[data-i18n-html]').forEach(el=>{const k=el.dataset.i18nHtml;
    if(!el.dataset.orig)el.dataset.orig=el.innerHTML;      // zh=DOM 原文，避免长文重复维护
    el.innerHTML=(LANG==='zh')?el.dataset.orig:(I18N.en[k]??el.dataset.orig);});
  document.querySelectorAll('[data-i18n-t]').forEach(el=>{el.title=t(el.dataset.i18nT)});
  document.querySelectorAll('[data-i18n-ph]').forEach(el=>{el.placeholder=t(el.dataset.i18nPh)});
  document.getElementById('brandTag').textContent=t('brand_tag');
  refreshStatus();
}
function toggleLang(){LANG=LANG==='zh'?'en':'zh';localStorage.setItem('gewu-lang',LANG);applyLang();loadSessions();if(MODE==='write')loadWriteLog();if(MODE==='history')loadHistList()}

let MODE='review', SHEETS=[], SCOPE=null, CUR_SESSION=null;
function showPane(name){
  MODE=name;
  document.querySelectorAll('.ptab').forEach(b=>b.classList.toggle('on',b.dataset.p===name));
  document.querySelectorAll('.pane').forEach(m=>m.classList.toggle('on',m.id==='p-'+name));
  if(name==='write')loadWriteLog();
  if(name==='history')loadHistList();
}
document.querySelectorAll('.ptab').forEach(btn=>{btn.onclick=()=>showPane(btn.dataset.p)});
document.getElementById('panelCollapse').onclick=()=>{
  document.getElementById('sidePanel').style.display='none';
  document.getElementById('panelHandle').style.display='none';
  document.getElementById('panelExpand').style.display='';};
document.getElementById('panelExpand').onclick=()=>{
  document.getElementById('sidePanel').style.display='';
  document.getElementById('panelHandle').style.display='';
  document.getElementById('panelExpand').style.display='none';};
(function(){  // 拖动调面板宽窄
  const h=document.getElementById('panelHandle'),sp=document.getElementById('sidePanel');
  let drag=false;
  h.onmousedown=e=>{drag=true;e.preventDefault()};
  window.addEventListener('mousemove',e=>{if(!drag)return;
    const w=window.innerWidth-e.clientX;
    sp.style.width=Math.min(Math.max(w,300),window.innerWidth*0.7)+'px';});
  window.addEventListener('mouseup',()=>drag=false);
})();
function toggleSessPop(){
  const p=document.getElementById('sessPop');
  // CSS 初始 display:none + JS 空串切换会永远隐藏——用显式 block/none
  p.style.display=p.style.display==='block'?'none':'block';
}
function renderScope(){
  const box=document.getElementById('scopeChips');
  let html=`<span class="scopeChip${SCOPE===null?' on':''}" onclick="setScope(null)">${t('whole_file')}</span>`;
  html+=SHEETS.filter(s=>s.visible).map(s=>
    `<span class="scopeChip${SCOPE&&SCOPE.includes(s.name)?' on':''}" onclick="toggleScope('${esc(s.name)}')">${esc(s.name)}</span>`).join('');
  box.innerHTML=html;
  // 折叠按钮上的标签
  const lbl=document.getElementById('scopeLabel');
  if(!lbl)return;
  if(SCOPE===null)lbl.textContent=t('whole_file');
  else if(SCOPE.length===1)lbl.textContent=SCOPE[0].slice(0,18);
  else lbl.textContent=SCOPE.length+' sheets';
}
function toggleScopePop(){
  const p=document.getElementById('scopePop');
  p.style.display=p.style.display==='block'?'none':'block';
}
function setScope(v){SCOPE=v;renderScope()}
function toggleScope(name){
  if(SCOPE===null)SCOPE=[name];
  else if(SCOPE.includes(name))SCOPE=SCOPE.filter(x=>x!==name);
  else SCOPE=SCOPE.concat([name]);
  if(!SCOPE.length)SCOPE=null;
  renderScope();
}
async function newChat(){
  const r=await (await fetch('/api/session_new',{method:'POST',headers:{'Content-Type':'application/json'},
    body:JSON.stringify({scope:SCOPE})})).json();
  CUR_SESSION=r.session_id;
  document.getElementById('chatLog').innerHTML=`<div class="a hint">${t('new_chat')}</div>`;
  loadSessions();
}
async function loadSessions(){
  const list=await (await fetch('/api/sessions')).json();
  const el=document.getElementById('sessionList');
  const cur=list.find(s=>s.id===CUR_SESSION);
  document.getElementById('curTitle').textContent=cur?(cur.title||'(untitled)'):t('cur_title_empty');
  if(!list.length){el.innerHTML=`<div class="hint">${t('hist_empty')}</div>`;return}
  el.innerHTML=list.map(s=>`<div class="sItem${s.id===CUR_SESSION?' on':''}" onclick="switchSession(${s.id})">
    <div style="font-weight:600;overflow:hidden;text-overflow:ellipsis;white-space:nowrap">${esc(s.title||'(未命名)')}</div>
    <div class="hint" style="font-size:11px">${s.updated_at||''}${s.scope?' · '+s.scope.length+'表':''}</div>
    <button class="sDel" title="删除本会话" onclick="event.stopPropagation();delSession(${s.id})">✕</button></div>`).join('');
}
async function switchSession(id){
  const r=await (await fetch('/api/session_switch',{method:'POST',headers:{'Content-Type':'application/json'},
    body:JSON.stringify({id})})).json();
  CUR_SESSION=id;
  SCOPE=r.session.scope||null;renderScope();
  const log=document.getElementById('chatLog');
  log.innerHTML=(r.messages.length?'':`<div class="a hint">${t('sess_empty')}</div>`)+
    r.messages.map(m=>m.role==='user'?`<div class="u">问：${esc(m.content)}</div>`
      :`<div class="a">${mdRich(m.content)}</div>`).join('');
  log.scrollTop=log.scrollHeight;
  loadSessions();
}
async function delSession(id){
  await fetch('/api/session_delete',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({id})});
  if(CUR_SESSION===id){CUR_SESSION=null;document.getElementById('chatLog').innerHTML=`<div class="a hint">${t('sess_deleted')}</div>`}
  loadSessions();
}
async function clearSessions(){
  if(!confirm(t('hist_clear_confirm')))return;
  await fetch('/api/sessions_clear',{method:'POST',headers:{'Content-Type':'application/json'},body:'{}'});
  CUR_SESSION=null;loadSessions();loadHistList();
}
async function loadHistList(){
  const list=await (await fetch('/api/sessions')).json();
  document.getElementById('histList').innerHTML=list.length?list.map(s=>
    `<div style="display:flex;gap:6px;padding:4px 2px;border-bottom:1px solid #f1f5f9">
      <div style="flex:1;cursor:pointer" onclick="showPane('chat');switchSession(${s.id})">
      <b>${esc(s.title||'(未命名)')}</b> <span class="hint">${s.updated_at||''}</span></div>
      <button style="border:0;background:transparent;color:#dc2626;cursor:pointer" onclick="delSession(${s.id});loadHistList()">✕</button></div>`).join('')
    :t('hist_none');
}
async function doHistSearch(){
  const q=document.getElementById('histSearch').value.trim();if(!q)return;
  const r=await (await fetch('/api/search_history?q='+encodeURIComponent(q))).json();
  document.getElementById('histResults').innerHTML=r.length?r.map(h=>
    `<div style="padding:5px 2px;border-bottom:1px solid #f1f5f9"><b>[${esc(h.file)}]</b> ${esc(h.title)}<br>
     <span class="hint">…${esc(h.snippet)}…</span></div>`).join(''):'（没有命中）';
}
async function loadWriteLog(){
  const r=await (await fetch('/api/write_log')).json();
  document.getElementById('writeLog').innerHTML=r.length?r.map(w=>
    `<div style="padding:3px 2px;border-bottom:1px solid #f1f5f9">${w.undone?'<s>':''}
     <span class="hint">${w.ts}</span> [${esc(w.sheet)}] r${w.row}c${w.col}：${esc(w.old)} → <b>${esc(w.new)}</b>${w.undone?t('undone'):''}${w.undone?'</s>':''}</div>`).join('')
    :t('write_log_none');
}
async function doAsk(){
  const box=document.getElementById('askBox');
  const q=box.value.trim();if(!q)return;
  box.value='';box.focus();  // 发送瞬间清空并聚焦，不等回答
  const log=document.getElementById('chatLog');
  log.insertAdjacentHTML('beforeend',`<div class="u">问：${esc(q)}</div>`);
  log.scrollTop=log.scrollHeight;
  const r=await (await fetch('/api/ask',{method:'POST',headers:{'Content-Type':'application/json'},
    body:JSON.stringify({question:q,sheet:S?S.sheet:null,scope:SCOPE,session_id:CUR_SESSION,lang:LANG})})).json();
  if(r.session_id)CUR_SESSION=r.session_id;
  const tag=r.source==='llm'?'🤖':(r.source==='rules-fallback'?'⚙️(LLM挂,规则兜底)':'⚙️');
  log.insertAdjacentHTML('beforeend',`<div class="a">${tag} ${mdRich(String(r.answer??r).slice(0,1200))}</div>`);
  log.scrollTop=log.scrollHeight;
  loadSessions();
}
let pendingOp=null,pendingOpSaved=null;
async function doPlanWrite(mode){
  const text=document.getElementById('writeBox').value.trim();if(!text)return;
  const card=document.getElementById('diffCard');
  card.className='';card.textContent='解析中…';
  const r=await (await fetch('/api/plan_write',{method:'POST',headers:{'Content-Type':'application/json'},
    body:JSON.stringify({mode,text})})).json();
  if(!r.ok){
    pendingOp=null;
    card.className='bad';
    card.innerHTML=`<b>未生成写计划</b><br>${esc(r.clarify||JSON.stringify(r).slice(0,200))}`+
      (r.candidates?`<br>候选：<br>${r.candidates.map(c=>'· '+esc(c.via)+' @r'+c.row).join('<br>')}`:'');
    return;
  }
  pendingOp=r.op;pendingOpSaved=r.op;window.__lastDeps=r.deps||[];
  const v=r.validation;
  card.className=v.ok?'ok':'bad';
  const modeZh={set:'改为',add:'加',sub:'减',mul:'乘',div:'除',add_pct:'涨%',sub_pct:'降%'};
  card.innerHTML=
   (window.COPY?`<div style="background:#fef3c7;border:1px solid #f59e0b;border-radius:6px;
      padding:4px 8px;margin-bottom:6px;font-size:12px;color:#92400e">
      ⚠ <b>副本模式</b>：本次写入只进副本。要发给别人的成品，写完点页眉<b>「📥 导出改后文件」</b>下载，别发原上传的文件。</div>`:'')+
   (r.hit&&Object.keys(r.hit).length?`<div style="display:flex;align-items:center;gap:4px;flex-wrap:wrap;margin:4px 0">
     <span class="hint" style="font-size:11px">${t('hit_row')}</span>`+
     Object.entries(r.hit).map(([k,vv])=>`<span style="background:#eef2ff;border:1px solid #c7d2fe;border-radius:10px;padding:0 8px;font-size:11px">${esc(k)}=${esc(String(vv).slice(0,14))}</span>`).join('')+`</div>`:'')+
   `<div style="display:flex;align-items:center;gap:6px;flex-wrap:wrap">
     <b>✏️ 写计划</b><span class="hint">${esc(r.via||'')}</span>
     ${r.op.mode&&modeZh[r.op.mode]?`<span style="background:#e0e7ff;border-radius:8px;padding:0 8px;font-size:11.5px">${modeZh[r.op.mode]}</span>`:''}
     <button style="margin-left:auto;font-size:11px" onclick="locateCell('${esc(r.op.sheet)}',${r.op.row},${r.op.col})">📍 定位到表格</button></div>
   <div style="display:flex;align-items:center;gap:8px;margin:6px 0;flex-wrap:wrap">
     <span class="hint">${esc(r.op.sheet)}!r${r.op.row}c${r.op.col}</span>
     <span style="background:#fee2e2;color:#b91c1c;border-radius:6px;padding:2px 10px;font-weight:600">${esc(String(r.old??'∅'))}</span>
     <span style="color:#94a3b8;font-size:16px">→</span>
     <span style="background:#dcfce7;color:#15803d;border-radius:6px;padding:2px 10px;font-weight:700">${esc(String(r.op.value))}</span>
   </div>`+
   (r.deps&&r.deps.length?`<div style="color:#7c3aed;font-size:12px">🔗 ${t('dep_pre')}${r.deps.length}${t('dep_suf')}（r${r.deps[0].row}c${r.deps[0].col} ${esc(r.deps[0].formula)}）</div>`:'')+
   (v.ok?`<div style="color:#15803d">✓ 硬校验全部通过</div>`:
        `<div style="color:#b91c1c">✗ 拒绝：${v.rejections.map(esc).join('；')}</div>`)+
   (v.warnings.length?`<div style="color:#b45309">⚠ ${v.warnings.map(esc).join('；')}</div>`:'')+
   (v.ok?`<div style="display:flex;gap:6px;margin-top:8px">
      <button class="primary" onclick="doExec()">确认写入（COM，自动备份）</button>
      <button onclick="pendingOp=null;document.getElementById('diffCard').innerText='已取消'">取消</button></div>`:'');
  locateCell(r.op.sheet,r.op.row,r.op.col);
}
async function locateCell(sheet,row,col){
  if(S.sheet!==sheet){
    document.getElementById('sheetSel').value=sheet;
    await loadState(sheet);
  }
  const td=anchorCell(row,col);
  if(!td)return;
  td.scrollIntoView({block:'center',inline:'center'});
  td.classList.remove('flashCell');void td.offsetWidth;td.classList.add('flashCell');
}
async function doExec(){
  if(!pendingOp)return;
  const r=await (await fetch('/api/execute_write',{method:'POST',headers:{'Content-Type':'application/json'},
    body:JSON.stringify({op:pendingOp})})).json();
  const card=document.getElementById('diffCard');
  card.className=r.ok?'ok':'bad';
  card.innerHTML=r.ok
    ?`✓ 已写入（${esc(r.engine)}，备份 ${esc(r.backup)}）写后校验通过。`+
      (window.__lastDeps&&window.__lastDeps.length?
        (r.deps_changed&&r.deps_changed.length?
          `<br><span style="color:#7c3aed;font-size:12px">🔗 ${t('dep_done_pre')}r`+
          r.deps_changed.slice(0,4).map(d=>`${d.row}c${d.col}`).join('、')+
          `${t('dep_done_suf')}</span>`:
          `<br><span style="color:#b45309;font-size:12px">${t('dep_unchanged')}</span>`):'')+
      (window.COPY?`<br><b style="color:#92400e">改动在副本上——点此拿到成品：</b>
        <a href="/api/export" download style="color:#d97706;font-weight:700">📥 导出改后文件</a>`:'')+
      `<br>出错了可点「撤销最近写入」。`
    :`✗ 执行失败：${esc(r.reason||JSON.stringify(r).slice(0,200))}`;
  pendingOp=null;loadState(document.getElementById('sheetSel').value);
  if(r.ok){const op=r.op||pendingOpSaved; if(op){locateCell(op.sheet,op.row,op.col).then(()=>{
    const td=anchorCell(op.row,op.col);if(td){td.classList.remove('flashCell');
      void td.offsetWidth;td.classList.add('flashDone')}})}}
}
async function doUndo(){
  const r=await (await fetch('/api/undo',{method:'POST',headers:{'Content-Type':'application/json'},body:'{}'})).json();
  const log=document.getElementById('chatLog');
  log.insertAdjacentHTML('beforeend',
    `<div class="a">${r.ok?'↩ 已撤销 '+r.restored+' 格（'+esc(r.engine)+'）':esc(r.reason||'无可撤销')}</div>`);
  loadState(document.getElementById('sheetSel').value);
}
window.addEventListener('resize',()=>S&&drawOverlays());
bootWait();
/* ===== 文件标签页 ===== */
let TABS=[], ACTIVE_TAB=-1;
function loadTabs(){try{TABS=JSON.parse(localStorage.getItem('gewu-tabs')||'[]')}catch(e){TABS=[]}
  renderTabs()}
function saveTabs(){localStorage.setItem('gewu-tabs',JSON.stringify(TABS))}
function renderTabs(){
  const bar=document.getElementById('tabBar');
  if(TABS.length<=1){bar.style.display='none';return}
  bar.style.display='flex';
  bar.innerHTML=TABS.map((tab,i)=>{
    const nm=tab.name||'?';
    return `<div class="ftab${i===ACTIVE_TAB?' on':''}" onclick="switchToTab(${i})" title="${esc(tab.path)}">
      <span class="tName">${esc(nm)}</span>
      ${tab.copy?`<span class="tBadge">〔${t('copy_badge')}〕</span>`:''}
      <button class="tClose" onclick="event.stopPropagation();closeTab(${i})">✕</button>
    </div>`}).join('')+
    `<button style="border:0;background:transparent;cursor:pointer;font-size:14px;padding:3px 8px;color:#64748b"
      onclick="showOpen()">＋</button>`;
}
async function syncTabs(){
  if(!S||!S.workbook)return;
  const cur=S.workbook.path;
  const i=TABS.findIndex(t=>t.path===cur);
  if(i>=0){ACTIVE_TAB=i;TABS[i].sheet=S.sheet;TABS[i].scope=SCOPE;
    TABS[i].sessionId=CUR_SESSION;TABS[i].copy=window.COPY}
  else{TABS.push({path:cur,name:cur.split(/[\\/]/).pop(),sheet:S.sheet,
    scope:SCOPE,sessionId:CUR_SESSION,copy:window.COPY});
    ACTIVE_TAB=TABS.length-1}
  saveTabs();renderTabs();
}
async function switchToTab(i){
  if(i===ACTIVE_TAB||!TABS[i])return;
  if(TABS[ACTIVE_TAB]){TABS[ACTIVE_TAB].sheet=S.sheet;TABS[ACTIVE_TAB].scope=SCOPE;
    TABS[ACTIVE_TAB].sessionId=CUR_SESSION}
  const tgt=TABS[i];ACTIVE_TAB=i;
  const r=await (await fetch('/api/open',{method:'POST',headers:{'Content-Type':'application/json'},
    body:JSON.stringify({path:tgt.path})})).json();
  if(!r.ok){alert('切换失败: '+(r.reason||''));return}
  await reloadForNewFile();
  if(tgt.sheet){const sel=document.getElementById('sheetSel');
    const opt=[...sel.options].find(o=>o.textContent.includes(tgt.sheet));
    if(opt){sel.value=opt.value;await loadState(tgt.sheet)}}
  if(tgt.scope!==undefined&&tgt.scope!==null){SCOPE=tgt.scope;renderScope()}
  else{SCOPE=null;renderScope()}
  saveTabs();renderTabs();
}
async function closeTab(i){
  if(TABS.length<=1)return;
  TABS.splice(i,1);
  if(i===ACTIVE_TAB){ACTIVE_TAB=Math.min(i,TABS.length-1);
    if(TABS[ACTIVE_TAB])await switchToTab(ACTIVE_TAB)}
  else if(i<ACTIVE_TAB)ACTIVE_TAB--;
  saveTabs();renderTabs();
}
async function reloadForNewFile(){
  edits={};sel=null;CUR_SESSION=null;SCOPE=null;
  await loadSheets();loadSessions();refreshStatus();
}
async function bootWait(){
  const ov=document.createElement('div');
  ov.id='bootOverlay';
  ov.style.cssText='position:fixed;inset:0;z-index:200;background:linear-gradient(135deg,#0f172a,#1e3a5f);color:#e2e8f0;display:flex;flex-direction:column;align-items:center;justify-content:center;font:14px/1.8 system-ui';
  ov.innerHTML='<div style="font-size:26px;font-weight:700;letter-spacing:4px">格物 <span style="background:linear-gradient(90deg,#60a5fa,#93c5fd);-webkit-background-clip:text;background-clip:text;color:transparent">Gewu</span></div><div id="bootMsg" style="color:#94a3b8;margin-top:10px">…</div><div style="width:30px;height:30px;border:3px solid #334155;border-top-color:#60a5fa;border-radius:50%;animation:gwspin .9s linear infinite;margin-top:16px"></div><style>@keyframes gwspin{to{transform:rotate(360deg)}}</style>';
  document.body.appendChild(ov);
  for(let i=0;i<300;i++){
    let b=null;
    try{b=await (await fetch('/api/boot')).json()}catch(e){}
    if(b&&b.ready){ov.remove();loadSheets();return}
    if(b&&b.phase==='failed'){
      ov.querySelector('#bootMsg').innerHTML='⚠ 解析失败：'+(b.error||'未知原因')+'<br><span style="font-size:12px;color:#64748b">请确认文件是有效的 .xlsx（.xls/.csv 暂不支持），或换一个文件重试。</span>';
      ov.lastElementChild?.remove?.(); ov.querySelector('div[style*="gwspin"],div[style*="rotate"]')?.remove?.();
      return;
    }
    if(b&&b.phase)document.getElementById('bootMsg').textContent=b.phase;
    await new Promise(r=>setTimeout(r,800));
  }
}
//
</script></body></html>"""



def _start_parent_watchdog():
    """桌面壳模式下（GEWU_PARENT_PID 注入）：父进程死了自己退出，绝不做孤儿。
    覆盖强杀场景——Tauri 的窗口销毁回调在 taskkill /F 下不会触发。"""
    pid = os.environ.get("GEWU_PARENT_PID")
    if not pid:
        return
    import ctypes
    import time as time_mod

    def _watch():
        k32 = ctypes.windll.kernel32
        PROCESS_QUERY_LIMITED_INFORMATION = 0x1000
        STILL_ACTIVE = 259
        target = int(pid)
        while True:
            h = k32.OpenProcess(PROCESS_QUERY_LIMITED_INFORMATION, False, target)
            if not h:
                os._exit(0)   # 打不开句柄：父进程大概率已消失（含复用竞态兜底）
            code = ctypes.c_ulong()
            ok = k32.GetExitCodeProcess(h, ctypes.byref(code))
            k32.CloseHandle(h)
            if not ok or code.value != STILL_ACTIVE:
                os._exit(0)
            time_mod.sleep(2.0)

    import threading
    threading.Thread(target=_watch, daemon=True).start()


def _col_letter(c):
    n = c
    out = ""
    while n:
        n, r = divmod(n - 1, 26)
        out = chr(65 + r) + out
    return out


def _predict_deps(op):
    """#4 同 sheet 公式依赖预测：谁引用了目标格（v1 不含跨 sheet 引用）。"""
    import re as _re2
    wb = _get_wb(STATE["wb_path"])
    sh = next((x for x in wb["sheets"] if x["name"] == op["sheet"]), None)
    if not sh:
        return []
    ref = f"{_col_letter(op['col'])}{op['row']}"
    pat = _re2.compile(rf"(?<![A-Z0-9!$]){ref}(?![0-9])")
    deps = []
    for (r, c), info in sh["cells"].items():
        if info.get("f") and isinstance(info["v"], str) and pat.search(info["v"]):
            deps.append({"sheet": op["sheet"], "row": r, "col": c,
                         "formula": info["v"][:40]})
    return deps[:8]



def _is_frozen():
    import sqlite3 as _sq
    try:
        con = _sq.connect(STATE["db"])
        row = con.execute("SELECT status FROM workbooks WHERE id=?",
                          (STATE["wid"],)).fetchone()
        con.close()
        return bool(row and row[0] == "frozen")
    except Exception:
        return False

def _switch_workbook(path, copy=False):
    """应用内打开：切换 STATE 到新工作簿（快启优先/失败全量重建），失败抛 SystemExit。"""
    import datetime as _dt
    from engine.mirror import try_fast_boot
    STATE["wb_path"] = path
    wid = try_fast_boot(path, STATE["db"])
    if wid:
        STATE["wid"] = wid
        STATE["fast_boot"] = True
    else:
        STATE["wid"] = build_mirror(path, STATE["db"])
    STATE["boot"] = _dt.datetime.now().isoformat(timespec="seconds")
    STATE["copy"] = copy
    _WB_CACHE.pop(path, None)
    _SHEET_CACHE.clear()

def main():
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(1)
    path = os.path.abspath(sys.argv[1])
    port = int(sys.argv[2]) if len(sys.argv) > 2 else PORT_DEFAULT
    STATE["wb_path"] = path
    STATE["db"] = DEFAULT_DB
    STATE["ready"] = False
    STATE["boot_phase"] = "启动中"

    # ① 端口先立起来：前端/exe 立刻可见"解析中"，端口存在性不再与解析成败绑死
    class _Server(ThreadingHTTPServer):
        allow_reuse_address = False
    try:
        httpd = _Server(("127.0.0.1", port), Handler)
    except OSError:
        print(f"✗ 端口 {port} 已被占用——可能已有一个确认页在运行。\n"
              f"  请先关闭旧窗口（或换端口：python serve.py 文件.xlsx {port + 1}）。")
        sys.exit(1)
    threading.Thread(target=httpd.serve_forever, daemon=True).start()
    url = f"http://127.0.0.1:{port}"
    _start_parent_watchdog()

    # ② 快启：同文件指纹 + 同引擎签名 → 直接载镜像（重开秒级）；否则后台解析
    from engine.mirror import try_fast_boot
    wid = try_fast_boot(path, STATE["db"])
    if wid:
        STATE["wid"] = wid
        STATE["ready"] = True
        STATE["boot"] = __import__("datetime").datetime.now().isoformat(timespec="seconds")
        STATE["boot_phase"] = "ready"
        STATE["fast_boot"] = True
        print(f"确认页已启动（快启）：{url}  （工作簿 id={wid} 复用现有镜像）")
    else:
        STATE["boot_phase"] = "正在解析工作簿…（大文件约需十几秒）"

        def _parse():
            try:
                STATE["wid"] = build_mirror(path, STATE["db"])
                STATE["ready"] = True
                STATE["boot"] = __import__("datetime").datetime.now().isoformat(timespec="seconds")
                STATE["boot_phase"] = "ready"
                print(f"解析完成：{url}  （工作簿 id={STATE['wid']}）")
            except SystemExit as e:
                STATE["boot_phase"] = "failed"
                STATE["boot_error"] = str(e)
                print(f"✗ 解析失败：{e}")
        threading.Thread(target=_parse, daemon=True).start()
    import webbrowser
    webbrowser.open(url)
    try:
        threading.Event().wait()
    except KeyboardInterrupt:
        print("\n已退出")


if __name__ == "__main__":
    main()
