# -*- coding: utf-8 -*-
"""格物 Gewu · Web 服务：确认页 + 工作台 + API。
用法：python serve.py <xlsx> [端口]   然后浏览器打开打印出的地址。
前端在 static/ 目录（index.html + style.css + app.js）。"""
import json
import logging
import os
import sqlite3
import sys
import threading
import time as _time_mod
import uuid
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from urllib.parse import urlparse, parse_qs

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
sys.stdout.reconfigure(encoding="utf-8")

# ── 日志：结构化输出到文件 + 控制台 ──
LOG_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "logs")
os.makedirs(LOG_DIR, exist_ok=True)
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s %(levelname)-5s [%(name)s] %(message)s",
    handlers=[
        logging.FileHandler(os.path.join(LOG_DIR, "gewu.log"), encoding="utf-8"),
        logging.StreamHandler(sys.stdout),
    ],
)
logger = logging.getLogger("gewu")
req_logger = logging.getLogger("gewu.request")

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
    """每个请求分配 request_id，记录方法/路径/状态/耗时。"""
    _req_start = None
    _request_id = None

    def log_message(self, fmt, *args):
        pass  # 覆写默认 stderr 日志，改用 req_logger

    def _begin_request(self):
        self._req_start = _time_mod.time()
        self._request_id = uuid.uuid4().hex[:8]

    def _end_request(self, method, path, status):
        if self._req_start is None:
            return
        elapsed_ms = int((_time_mod.time() - self._req_start) * 1000)
        req_logger.info(f"[{self._request_id}] {method} {path} {status} {elapsed_ms}ms")
        self._req_start = None

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
        self._begin_request()
        self._serve_request("GET")

    def do_POST(self):
        self._begin_request()
        self._serve_request("POST")

    def _serve_request(self, method):
        u = urlparse(self.path)
        try:
            if method == "GET":
                self._route_get(u)
            else:
                self._route_post(u)
        except BrokenPipeError:
            logger.warning(f"[{self._request_id}] {method} {u.path} client disconnected")
        except Exception as e:
            logger.error(f"[{self._request_id}] {method} {u.path} {type(e).__name__}: {e}")
            try:
                self._json({"ok": False, "reason": f"内部错误: {type(e).__name__}"}, 500)
            except Exception:
                pass
        finally:
            self._end_request(method, u.path, getattr(self, "_status_code", 200))

    def _route_get(self, u):
        if u.path.startswith("/static/"):
            fname = u.path[len("/static/"):]
            fpath = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                                 "static", fname)
            if os.path.isfile(fpath):
                ext = os.path.splitext(fname)[1]
                mime = {".css": "text/css", ".js": "application/javascript",
                        ".html": "text/html; charset=utf-8",
                        ".png": "image/png", ".ico": "image/x-icon"}.get(ext, "application/octet-stream")
                with open(fpath, "rb") as f:
                    body = f.read()
                self.send_response(200)
                self.send_header("Content-Type", mime)
                self.send_header("Content-Length", str(len(body)))
                self.end_headers()
                self.wfile.write(body)
            else:
                self.send_response(404)
                self.end_headers()
        elif u.path == "/":
            with open(PAGE_FILE, encoding="utf-8") as f:
                body = f.read().encode("utf-8")
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
        elif u.path == "/api/branding":
            self._json({"product": BRANDING.get("product", ""),
                        "product_en": BRANDING.get("product_en", ""),
                        "tagline": BRANDING.get("tagline", ""),
                        "owner": BRANDING.get("owner", ""),
                        "owner_tip": BRANDING.get("owner_tip", ""),
                        "watermark": BRANDING.get("watermark", ""),
                        "email": BRANDING.get("email", "")})
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
            try:
                rows = con.execute(
                    "SELECT ts, sheet, row_no, col_no, old_repr, new_repr, undone "
                    "FROM write_ops WHERE workbook_id=? ORDER BY id DESC LIMIT 30",
                    (STATE["wid"],)).fetchall()
            except sqlite3.OperationalError:
                rows = []  # write_ops 表在首次写入时才创建——空表正常
            finally:
                con.close()
            self._json([{"ts": r[0], "sheet": r[1], "row": r[2], "col": r[3],
                         "old": r[4], "new": r[5], "undone": bool(r[6])} for r in rows])

    def _mirror(self):
        m = Mirror(STATE["db"], workbook_id=STATE["wid"])
        return m

    def _route_post(self, u):
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


PAGE_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "static", "index.html")



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
