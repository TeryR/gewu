# -*- coding: utf-8 -*-
"""语义镜像层：把解析结果物化为自包含的 SQLite 冻结镜像（PRD 核心资产）。
下游 LLM 问答、写回校验只认镜像，不回头碰原始文件。
表：workbooks / sheets / blocks / records / glossary。"""
import hashlib
import json
import os
import sqlite3
import datetime as _dt

from .extract import extract_workbook, _has_value
from .headers import build_header_tree
from .segment import segment_sheet

DEFAULT_DB = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
                          "mirror.db")


def _now():
    return _dt.datetime.now().isoformat(timespec="seconds")


def _file_hash(path):
    h = hashlib.sha1()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()[:16]


def _init(db):
    con = sqlite3.connect(db)
    con.row_factory = sqlite3.Row
    con.executescript("""
    CREATE TABLE IF NOT EXISTS workbooks(
        id INTEGER PRIMARY KEY, path TEXT, file_hash TEXT UNIQUE, sanitized INTEGER,
        created_at TEXT, status TEXT DEFAULT 'parsed', frozen_at TEXT);
    CREATE TABLE IF NOT EXISTS sheets(
        id INTEGER PRIMARY KEY, workbook_id INTEGER, idx INTEGER, name TEXT, visible INTEGER);
    CREATE TABLE IF NOT EXISTS blocks(
        id INTEGER PRIMARY KEY, workbook_id INTEGER, sheet_name TEXT, block_id INTEGER,
        r1 INTEGER, c1 INTEGER, r2 INTEGER, c2 INTEGER,
        type TEXT, score REAL, flags_json TEXT, header_json TEXT,
        summary_formulas_json TEXT,
        user_type TEXT, readonly INTEGER DEFAULT 0);
    CREATE TABLE IF NOT EXISTS records(
        id INTEGER PRIMARY KEY, block_db_id INTEGER, row_no INTEGER,
        is_summary INTEGER DEFAULT 0, data_json TEXT);
    CREATE TABLE IF NOT EXISTS glossary(
        id INTEGER PRIMARY KEY, workbook_id INTEGER, term TEXT, meaning TEXT,
        col_path TEXT, sheet_name TEXT, created_at TEXT);
    CREATE TABLE IF NOT EXISTS row_notes(
        id INTEGER PRIMARY KEY, workbook_id INTEGER, sheet_name TEXT,
        note_block_id INTEGER, table_block_id INTEGER,   -- 引擎确定性编号
        note_db_id INTEGER, table_db_id INTEGER,         -- 本次构建的库主键
        row_no INTEGER, text TEXT);
    CREATE TABLE IF NOT EXISTS note_attach_off(          -- 人工取消挂接（重建幸存）
        workbook_id INTEGER, sheet_name TEXT, block_id INTEGER,
        PRIMARY KEY(workbook_id, sheet_name, block_id));
    CREATE TABLE IF NOT EXISTS chat_sessions(
        id INTEGER PRIMARY KEY, workbook_id INTEGER, title TEXT,
        scope_json TEXT, created_at TEXT, updated_at TEXT);
    CREATE TABLE IF NOT EXISTS chat_messages(
        id INTEGER PRIMARY KEY, session_id INTEGER, role TEXT, content TEXT, ts TEXT);
    CREATE TABLE IF NOT EXISTS biz_rules(            -- 业务值域/格式规则（人配置，引擎硬执行）
        id INTEGER PRIMARY KEY, workbook_id INTEGER, sheet_name TEXT, col_path TEXT,
        rmin REAL, rmax REAL, pattern TEXT, severity TEXT DEFAULT 'block', created_at TEXT);
    """)
    # 旧库迁移：变更摘要列 / 引擎版本列（快启判定用）
    _cols = {r[1] for r in con.execute("PRAGMA table_info(workbooks)")}
    if "change_json" not in _cols:
        con.execute("ALTER TABLE workbooks ADD COLUMN change_json TEXT")
    if "engine_ver" not in _cols:
        con.execute("ALTER TABLE workbooks ADD COLUMN engine_ver TEXT")
    con.commit()
    return con


def engine_sig():
    """引擎签名：engine/*.py 与 rules.json 的（大小, mtime）指纹。
    解析逻辑或规则一变签名即变，快启缓存自动失效——不会拿旧结构的镜像冒充新结果。"""
    import hashlib
    root = os.path.dirname(os.path.abspath(__file__))
    h = hashlib.sha1()
    for name in sorted(os.listdir(root)):
        if name.endswith(".py"):
            st = os.stat(os.path.join(root, name))
            h.update(f"{name}:{st.st_size}:{int(st.st_mtime)};".encode())
    rules = os.path.join(os.path.dirname(root), "rules.json")
    if os.path.exists(rules):
        st = os.stat(rules)
        h.update(f"rules.json:{st.st_size}:{int(st.st_mtime)};".encode())
    return h.hexdigest()[:12]


def try_fast_boot(path, db=DEFAULT_DB):
    """快启：文件指纹相同 且 引擎签名相同 且 镜像非空 → 直接复用镜像（跳过解析）。
    返回 wid 或 None（None = 调用方走全量 build_mirror）。"""
    fhash = _file_hash(path)
    sig = engine_sig()
    if not os.path.exists(db):
        return None
    con = _init(db)
    row = con.execute("SELECT id FROM workbooks WHERE file_hash=?", (fhash,)).fetchone()
    if not row:
        con.close()
        return None
    wid = row[0]
    ver, n_blocks = con.execute(
        "SELECT engine_ver, (SELECT COUNT(*) FROM blocks WHERE workbook_id=workbooks.id) "
        "FROM workbooks WHERE id=?", (wid,)).fetchone()
    if ver != sig or not n_blocks:
        con.close()
        return None
    con.execute("UPDATE workbooks SET path=? WHERE id=?", (path, wid))
    con.commit()
    con.close()
    return wid


def _col_paths(header):
    """叶子列 → 全路径标签（同名全路径追加 #2）。"""
    seen = {}
    out = []
    for node in header.get("col_tree") or []:
        p = "/".join(node["path"]) if node["path"] else f"col{node['col']}"
        if p in seen:
            seen[p] += 1
            p = f"{p}#{seen[p]}"
        else:
            seen[p] = 1
        out.append({"col": node["col"], "path": p})
    return out


def _build_records(cells, block):
    h = block.get("header")
    if not h:
        return [], [], {}
    r1, c1, r2, c2 = block["bbox"]
    skip_rows = set(h["header_rows"]) | set(h.get("floating_rows") or []) \
        | ({h["title_row"]} if h.get("title_row") else set())
    summary = set(h.get("summary_rows") or [])
    cols = _col_paths(h)
    col_of_path = {c["path"]: c["col"] for c in cols}
    col_meta = {}
    rows = []
    summary_formulas = {}
    col_formulas = {}   # 列 → 首个公式原文（回答"这列为什么这么算"，描述用，不参与计算）
    for r in range(r1, r2 + 1):
        if r in skip_rows:
            continue
        rec = {"_r": r}
        for cinfo in cols:
            c = cinfo["col"]
            info = cells.get((r, c))
            if not info:
                continue
            v = info["cv"] if info["cv"] is not None else info["v"]
            if _has_value(v):
                rec[cinfo["path"]] = v
                if info.get("nf") and "%" in info["nf"]:
                    col_meta.setdefault(cinfo["path"], {"pct": True})
            if info.get("f") and isinstance(info["v"], str):
                # f 是"有公式"标记，公式原文在原始 v（如 "==经营数据分析!B4/B19"）；
                # 上面合并出的 v 此刻已是缓存数值，不能用作公式
                path = cinfo["path"]
                if r in summary:
                    summary_formulas[f"{r}|{path}"] = info["v"]
                elif path not in col_formulas and len(col_formulas) < 12:
                    col_formulas[path] = info["v"]
        if len(rec) > 1:
            rec["_summary"] = r in summary
            rows.append(rec)
    pct_cols = [p for p, m in col_meta.items() if m.get("pct")]
    return rows, [{"path": p, "pct": bool(m.get("pct"))} for p, m in
                  {c["path"]: col_meta.get(c["path"], {}) for c in cols}.items()], \
        {"formulas": summary_formulas, "col_of_path": col_of_path,
         "pct_cols": pct_cols, "col_formulas": col_formulas}


def _attach_row_notes(cells, seg_blocks, off_keys, sheet_name, rules):
    """行内备注挂接（保守推导）：单行黄块，与表的数据行重叠、且该行还有表内容
    → 挂为该行批注（records._note + row_notes）。
    备注列被剥开常把表劈成两半 → 取列距最近的表（平手取左）。
    跨行/表级说明词/人工解除的保持独立黄块，不猜。
    返回 {(表引擎id, 行号): 备注文本, ...} 与挂接明细列表。"""
    markers = (rules.get("vocab") or {}).get("table_note_markers") or []
    min_len = ((rules.get("segment") or {}).get("annotation") or {}).get("row_note_min_len", 6)
    content = {p for p, i in cells.items()
               if i.get("v") is not None and str(i["v"]).strip()}
    tabs = [b for b in seg_blocks if b["type"] == "table"]
    notes = [b for b in seg_blocks if b["type"] == "note"]
    # 同一行所有黄块的列区间并集（排除时用，避免把同表内容误算成备注自己的）
    note_span = {}
    for nb in notes:
        r1, c1, r2, c2 = nb["bbox"]
        if r1 == r2:
            note_span.setdefault(r1, []).append((c1, c2))
    attached, detail = {}, []
    for nb in notes:
        r1, c1, r2, c2 = nb["bbox"]
        if r1 != r2:                      # 跨行备注归属说不清 → 不挂
            continue
        text = "；".join(str(cells[(r1, c)]["v"]).strip()
                        for c in range(c1, c2 + 1) if (r1, c) in content)
        if not text or text.startswith("="):   # 空文本/公式串（降级模式的公式文本）不是人写备注
            continue
        if len(text) < min_len:                # 短标签（如"销售部"节标签）当不了行备注
            continue
        if any(mk in text for mk in markers):   # 表级说明（本表数据截至…）→ 不挂到某一行
            continue
        if (sheet_name, nb["id"]) in off_keys:   # 人工解除过 → 不挂
            continue
        cands = []
        for t in tabs:
            tr1, tc1, tr2, tc2 = t["bbox"]
            if not (tr1 <= r1 <= tr2):
                continue
            h = t.get("header") or {}
            if r1 in (h.get("header_rows") or []) or r1 == h.get("title_row") \
                    or r1 in (h.get("floating_rows") or []):
                continue                        # 表头区里的说明不是行数据
            spans = note_span.get(r1, [])
            has = any((r1, c) in content and tc1 <= c <= tc2
                      and not any(a <= c <= b for a, b in spans)
                      for c in range(tc1, tc2 + 1))
            if has:
                cands.append(t)
        if cands:
            # 备注列被剥开常把一张表劈成左右两半，两边同行都有内容——
            # 取列距最近的表（平手取左：行身份/键列通常在左半）。挂错了人工可解除。
            def gap(t):
                _, a1, _, a2 = t["bbox"]
                return (c1 - a2) if a2 < c1 else (a1 - c2) if a1 > c2 else -1
            t = min(cands, key=lambda t: (gap(t), t["bbox"][1]))
            key = (t["id"], r1)
            attached[key] = f"{attached[key]}；{text}" if key in attached else text
            detail.append({"note_block_id": nb["id"], "table_block_id": t["id"],
                           "row_no": r1, "text": text})
    return attached, detail


def build_mirror(path: str, db: str = DEFAULT_DB, rules=None) -> int:
    """解析并物化镜像；同文件重复构建会替换旧镜像。返回 workbook_id。"""
    from .rules import get_rules
    rules = rules or get_rules()
    wb = extract_workbook(path)
    con = _init(db)
    fhash = _file_hash(path)
    row = con.execute("SELECT id FROM workbooks WHERE file_hash=?", (fhash,)).fetchone()
    takeover = False      # 同路径但 hash 变了：文件被改过（自己写数后重开/外部编辑）
    if row:
        wid = row[0]
    else:
        t = con.execute("SELECT id FROM workbooks WHERE path=? ORDER BY id DESC LIMIT 1",
                        (path,)).fetchone()
        if t:
            wid, takeover = t[0], True
    if row or takeover:
        # 重建（同 hash）或接管（同路径新版本）：冻结状态与人工标记必须幸存——
        # 重启服务不得清空人工确认成果（PRD：下游只认冻结镜像）
        prev = con.execute("SELECT status, frozen_at FROM workbooks WHERE id=?",
                           (wid,)).fetchone()
        prev = {"status": prev[0], "frozen_at": prev[1]} if prev else None
        old_blocks = [dict(r) for r in con.execute(
            "SELECT sheet_name, block_id, r1, c1, r2, c2, type, user_type, "
            "readonly, header_json FROM blocks WHERE workbook_id=?", (wid,))]
        marks = {(b["sheet_name"], b["block_id"]): (b["user_type"], b["readonly"])
                 for b in old_blocks if b["user_type"] or b["readonly"]}
        old_offs = [(s, bid) for s, bid in con.execute(
            "SELECT sheet_name, block_id FROM note_attach_off WHERE workbook_id=?",
            (wid,))]
        # 先删子表再删父表；records 必须经由 blocks 反查——block 主键与 wid 是两套数字，
        # 直接 WHERE block_db_id=wid 会漏删，重复构建时记录翻倍累积（曾致同行出现3份）
        con.execute("DELETE FROM records WHERE block_db_id IN "
                    "(SELECT id FROM blocks WHERE workbook_id=?)", (wid,))
        con.execute("DELETE FROM blocks WHERE workbook_id=?", (wid,))
        con.execute("DELETE FROM sheets WHERE workbook_id=?", (wid,))
        # row_notes 是构建时确定性推导，随重建重算；note_attach_off 接管时按签名重映射
        con.execute("DELETE FROM row_notes WHERE workbook_id=?", (wid,))
        if takeover:
            con.execute("UPDATE workbooks SET path=?, file_hash=?, sanitized=?, "
                        "created_at=?, change_json=NULL, engine_ver=? WHERE id=?",
                        (path, fhash, int(wb["sanitized"]), _now(), engine_sig(), wid))
        else:
            con.execute("UPDATE workbooks SET path=?, sanitized=?, created_at=?, "
                        "change_json=NULL, engine_ver=? WHERE id=?",
                        (path, int(wb["sanitized"]), _now(), engine_sig(), wid))
    else:
        prev, marks, old_blocks, old_offs = None, {}, [], []
        cur = con.execute(
            "INSERT INTO workbooks(path, file_hash, sanitized, created_at, engine_ver) "
            "VALUES(?,?,?,?,?)",
            (path, fhash, int(wb["sanitized"]), _now(), engine_sig()))
        wid = cur.lastrowid

    # 人工解除挂接名单（重建幸存；block_id 由引擎确定性生成，同文件同 id）
    off_keys = set(old_offs) if (row or takeover) else set()

    new_key6_id, new_key6_paths = {}, {}   # 接管比对：结构键 → 新引擎 id / 列路径签名
    for idx, s in enumerate(wb["sheets"]):
        con.execute("INSERT INTO sheets(workbook_id, idx, name, visible) VALUES(?,?,?,?)",
                    (wid, idx, s["name"], int(s["visible"])))
        seg = segment_sheet(s, rules)
        for b in seg["blocks"]:               # 先备表头树（挂接判定要看表头区）
            if b["type"] == "table":
                b["header"] = build_header_tree(s, b, rules)
        attach_map, attach_detail = _attach_row_notes(
            s["cells"], seg["blocks"], off_keys, s["name"], rules)
        dbid_of = {}                          # 引擎块id → 库主键（表和黄块都要）
        for b in seg["blocks"]:
            flags = {k: b.get(k) for k in ("annotation", "summary_tail")}
            header_json = json.dumps(b.get("header") or {}, ensure_ascii=False)  # 含 col_tree
            rows, summary_meta = [], {"formulas": {}, "col_of_path": {}}
            if b["type"] == "table":
                rows, _, summary_meta = _build_records(s["cells"], b)
            cur = con.execute(
                "INSERT INTO blocks(workbook_id, sheet_name, block_id, r1,c1,r2,c2,"
                "type, score, flags_json, header_json, summary_formulas_json) "
                "VALUES(?,?,?,?,?,?,?,?,?,?,?,?)",
                (wid, s["name"], b["id"], *b["bbox"], b["type"], b["score"],
                 json.dumps(flags, ensure_ascii=False), header_json,
                 json.dumps(summary_meta, ensure_ascii=False)))
            dbid_of[b["id"]] = cur.lastrowid
            # 接管比对用：结构键（sheet+几何+类型）→ 新引擎 id；表另记列路径签名
            k6 = (s["name"], b["bbox"][0], b["bbox"][1], b["bbox"][2], b["bbox"][3], b["type"])
            new_key6_id[k6] = b["id"]
            if b["type"] == "table":
                new_key6_paths[k6] = tuple(
                    "/".join(n["path"]) for n in ((b.get("header") or {}).get("col_tree") or [])
                    if n.get("path"))
            for rec in rows:
                # 行内备注进入记录（_ 前缀，不会与列路径冲突；聚合/勾稽天然忽略）
                note = attach_map.get((b["id"], rec["_r"]))
                if note:
                    rec["_note"] = note
                con.execute(
                    "INSERT INTO records(block_db_id, row_no, is_summary, data_json) "
                    "VALUES(?,?,?,?)",
                    (cur.lastrowid, rec["_r"], int(rec.get("_summary", False)),
                     json.dumps(rec, ensure_ascii=False, default=str)))
        for a in attach_detail:
            con.execute(
                "INSERT INTO row_notes(workbook_id, sheet_name, note_block_id, "
                "table_block_id, note_db_id, table_db_id, row_no, text) "
                "VALUES(?,?,?,?,?,?,?,?)",
                (wid, s["name"], a["note_block_id"], a["table_block_id"],
                 dbid_of.get(a["note_block_id"]), dbid_of.get(a["table_block_id"]),
                 a["row_no"], a["text"]))
    marks_apply, change = marks, None
    if takeover:
        def _key6(b):
            return (b["sheet_name"], b["r1"], b["c1"], b["r2"], b["c2"], b["type"])

        def _paths(b):
            if b["type"] != "table":
                return ()
            h = json.loads(b["header_json"] or "{}")
            return tuple("/".join(n["path"]) for n in (h.get("col_tree") or []) if n.get("path"))

        old_by_id = {b["block_id"]: b for b in old_blocks}
        old_key6 = {_key6(b): b for b in old_blocks}
        added = [k for k in new_key6_id if k not in old_key6]
        removed = [k for k in old_key6 if k not in new_key6_id]
        changed = [k for k in new_key6_id
                   if k in old_key6 and k[5] == "table"
                   and _paths(old_key6[k]) != new_key6_paths.get(k)]
        fmt = lambda ks: [f"{k[0]}!r{k[1]}-r{k[3]}×c{k[2]}-c{k[4]}" for k in ks][:10]
        # 人工标记迁移：结构键+列路径都没变的块才带过去（变了的需要人重新看过）
        marks_apply, transferred = {}, 0
        for (sname, bid), (ut, ro) in marks.items():
            ob = old_by_id.get(bid)
            if not ob:
                continue
            k6 = _key6(ob)
            if k6 not in new_key6_id or (k6[5] == "table"
                                         and _paths(ob) != new_key6_paths.get(k6)):
                continue
            if ut is not None or ro:
                marks_apply[(sname, new_key6_id[k6])] = (ut, ro)
                transferred += 1
        # 行内备注"人工解除"按黄块签名重映射（引擎 id 重新编号了）
        new_offs = [(sname, new_key6_id[_key6(ob)])
                    for sname, bid in old_offs
                    if (ob := old_by_id.get(bid)) and _key6(ob) in new_key6_id]
        con.execute("DELETE FROM note_attach_off WHERE workbook_id=?", (wid,))
        con.executemany("INSERT OR IGNORE INTO note_attach_off VALUES(?,?,?)",
                        [(wid, s, b) for s, b in new_offs])
        structural = bool(added or removed or changed)
        was_frozen = prev and prev["status"] == "frozen"
        if was_frozen and not structural:
            # 冻结保护的是结构判定；纯数值变化结构不动 → 冻结自动保持，数值随重建刷新
            change = {"type": "value-only", "kept_frozen": True}
        elif was_frozen:
            change = {"type": "structure", "added": fmt(added), "removed": fmt(removed),
                      "changed": fmt(changed), "marks_transferred": transferred,
                      "hint": "结构有变化，请复核变更块后重新冻结"}
        else:
            change = {"type": "fresh", "marks_transferred": transferred} if transferred else None
    # 同文件重建：回填冻结状态与人工标记（block_id 由引擎确定性生成，同文件同 id）
    keep_frozen = prev and prev["status"] == "frozen" and \
        (not takeover or not change or change.get("type") == "value-only")
    if keep_frozen:
        con.execute("UPDATE workbooks SET status='frozen', frozen_at=? WHERE id=?",
                    (prev["frozen_at"] or _now(), wid))
    elif takeover and prev and prev["status"] == "frozen":
        # 结构变化：冻结必须主动降级为待复核（行上的旧状态不能沿用）
        con.execute("UPDATE workbooks SET status='parsed', frozen_at=NULL WHERE id=?",
                    (wid,))
    for (sname, bid), (ut, ro) in marks_apply.items():
        if ut is not None or ro:
            con.execute(
                "UPDATE blocks SET user_type=COALESCE(?, user_type), readonly=COALESCE(?, readonly) "
                "WHERE workbook_id=? AND sheet_name=? AND block_id=?",
                (ut, ro or None, wid, sname, bid))
    if change is not None or row is None:
        con.execute("UPDATE workbooks SET change_json=? WHERE id=?",
                    (json.dumps(change, ensure_ascii=False) if change else None, wid))
    con.commit()
    con.close()
    return wid


def apply_ops_to_mirror(workbook_id, db, path, ops):
    """自己写数后的镜像同步：①file_hash 滚到写入后的新文件（重开/重拖即同 hash 续用冻结）
    ②只改受影响记录的对应字段值（结构/冻结/标记不动）。
    撤销时同样调用（ops 传恢复的旧值）。"""
    con = _init(db)
    con.execute("UPDATE workbooks SET file_hash=? WHERE id=?",
                (_file_hash(path), workbook_id))
    for op in ops:
        blk = con.execute(
            "SELECT id, header_json FROM blocks WHERE workbook_id=? AND sheet_name=? "
            "AND type='table' AND r1<=? AND r2>=? AND c1<=? AND c2>=?",
            (workbook_id, op["sheet"], op["row"], op["row"], op["col"], op["col"])
        ).fetchone()
        if not blk:
            continue
        h = json.loads(blk["header_json"] or "{}")
        col_path = next(("/".join(n["path"]) for n in (h.get("col_tree") or [])
                         if n.get("col") == op["col"] and n.get("path")), None)
        if not col_path:
            continue
        rec = con.execute("SELECT id, data_json FROM records WHERE block_db_id=? AND row_no=?",
                          (blk["id"], op["row"])).fetchone()
        if not rec:
            continue
        d = json.loads(rec["data_json"])
        if op.get("value") is None:
            d.pop(col_path, None)
        else:
            d[col_path] = op["value"]
        con.execute("UPDATE records SET data_json=? WHERE id=?",
                    (json.dumps(d, ensure_ascii=False, default=str), rec["id"]))
    con.commit()
    con.close()


def freeze(workbook_id: int, db: str = DEFAULT_DB, edits: list | None = None):
    """人工确认后冻结：edits=[{block_db_id, user_type?, readonly?}]。
    返回 {applied, missed}：镜像若已重建，旧 db 主键会落空，missed>0 必须提示用户刷新。"""
    con = _init(db)
    applied = missed = 0
    for e in edits or []:
        sets, vals = [], []
        if "user_type" in e:
            sets.append("user_type=?")
            vals.append(e["user_type"])
        if "readonly" in e:
            sets.append("readonly=?")
            vals.append(int(e["readonly"]))
        if sets:
            vals.append(e["block_db_id"])
            cur = con.execute(f"UPDATE blocks SET {', '.join(sets)} WHERE id=?", vals)
            if cur.rowcount:
                applied += 1
            else:
                missed += 1
    # 重新冻结 = 变更已复核完毕：清掉接管横幅，问答/写入闸门恢复
    con.execute("UPDATE workbooks SET status='frozen', frozen_at=?, change_json=NULL "
                "WHERE id=?", (_now(), workbook_id))
    con.commit()
    con.close()
    return {"applied": applied, "missed": missed}


def add_glossary(workbook_id: int, term: str, meaning: str, db: str = DEFAULT_DB,
                 col_path: str = None, sheet_name: str = None):
    """企业口径词典：人工确认的语义结论即存即用（PRD 积累断层补丁）。
    同字段重复保存覆盖旧释义，不堆积。"""
    con = _init(db)
    con.execute("DELETE FROM glossary WHERE workbook_id=? AND term=? AND "
                "IFNULL(col_path,'')=IFNULL(?,'') AND IFNULL(sheet_name,'')=IFNULL(?,'')",
                (workbook_id, term, col_path, sheet_name))
    con.execute("INSERT INTO glossary(workbook_id, term, meaning, col_path, sheet_name, "
                "created_at) VALUES(?,?,?,?,?,?)",
                (workbook_id, term, meaning, col_path, sheet_name, _now()))
    con.commit()
    con.close()


def detach_note(workbook_id: int, sheet_name: str, block_id: int,
                db: str = DEFAULT_DB):
    """人工取消一条行内备注挂接：记入 note_attach_off（同文件重建幸存），
    并即时清掉挂接行与记录里的 _note。无挂接时幂等成功。"""
    con = _init(db)
    row = con.execute(
        "SELECT table_db_id, row_no FROM row_notes WHERE workbook_id=? "
        "AND sheet_name=? AND note_block_id=?",
        (workbook_id, sheet_name, block_id)).fetchone()
    con.execute("INSERT OR IGNORE INTO note_attach_off(workbook_id, sheet_name, block_id) "
                "VALUES(?,?,?)", (workbook_id, sheet_name, block_id))
    con.execute("DELETE FROM row_notes WHERE workbook_id=? AND sheet_name=? "
                "AND note_block_id=?", (workbook_id, sheet_name, block_id))
    if row and row[0]:
        rec = con.execute("SELECT id, data_json FROM records "
                          "WHERE block_db_id=? AND row_no=?", (row[0], row[1])).fetchone()
        if rec:
            d = json.loads(rec[1])
            d.pop("_note", None)
            con.execute("UPDATE records SET data_json=? WHERE id=?",
                        (json.dumps(d, ensure_ascii=False, default=str), rec[0]))
    con.commit()
    con.close()
    return {"detached": 1 if row else 0}


# ---------------- 对话会话（按文件谱系持久化；跨文件检索） ----------------

def chat_session_create(workbook_id, db=DEFAULT_DB, title="", scope=None):
    con = _init(db)
    cur = con.execute(
        "INSERT INTO chat_sessions(workbook_id, title, scope_json, created_at, updated_at) "
        "VALUES(?,?,?,?,?)",
        (workbook_id, (title or "")[:40],
         json.dumps(scope, ensure_ascii=False) if scope else None, _now(), _now()))
    con.commit()
    sid = cur.lastrowid
    con.close()
    return sid


def chat_session_list(workbook_id, db=DEFAULT_DB):
    con = _init(db)
    rows = con.execute(
        "SELECT id, title, scope_json, created_at, updated_at FROM chat_sessions "
        "WHERE workbook_id=? ORDER BY id DESC", (workbook_id,)).fetchall()
    con.close()
    return [{"id": r[0], "title": r[1],
             "scope": json.loads(r[2]) if r[2] else None,
             "created_at": r[3], "updated_at": r[4]} for r in rows]


def chat_session_load(session_id, db=DEFAULT_DB):
    con = _init(db)
    msgs = con.execute("SELECT role, content FROM chat_messages WHERE session_id=? "
                       "ORDER BY id", (session_id,)).fetchall()
    con.close()
    return [{"role": r[0], "content": r[1]} for r in msgs]


def chat_session_append(session_id, role, content, db=DEFAULT_DB):
    con = _init(db)
    con.execute("INSERT INTO chat_messages(session_id, role, content, ts) VALUES(?,?,?,?)",
                (session_id, role, str(content)[:4000], _now()))
    con.execute("UPDATE chat_sessions SET updated_at=? WHERE id=?", (_now(), session_id))
    con.commit()
    con.close()


def chat_session_delete(session_id, db=DEFAULT_DB):
    con = _init(db)
    con.execute("DELETE FROM chat_messages WHERE session_id=?", (session_id,))
    con.execute("DELETE FROM chat_sessions WHERE id=?", (session_id,))
    con.commit()
    con.close()


def chat_sessions_clear(workbook_id, db=DEFAULT_DB):
    con = _init(db)
    con.execute("DELETE FROM chat_messages WHERE session_id IN "
                "(SELECT id FROM chat_sessions WHERE workbook_id=?)", (workbook_id,))
    con.execute("DELETE FROM chat_sessions WHERE workbook_id=?", (workbook_id,))
    con.commit()
    con.close()


def chat_history_search(q, db=DEFAULT_DB, limit=30):
    """跨工作簿检索会话（标题+消息命中），只读参考：返回文件名/标题/命中片段。"""
    if not q or not q.strip():
        return []
    con = _init(db)
    like = f"%{q.strip()}%"
    rows = con.execute(
        "SELECT m.content, s.id, s.title, w.path FROM chat_messages m "
        "JOIN chat_sessions s ON s.id=m.session_id JOIN workbooks w ON w.id=s.workbook_id "
        "WHERE m.content LIKE ? ORDER BY m.id DESC LIMIT ?", (like, limit)).fetchall()
    out = []
    for r in rows:
        text = r[0]
        i = text.find(q.strip())
        snippet = text[max(0, i - 20):i + 80].replace(chr(10), " ")
        out.append({"session_id": r[1], "title": r[2] or "(未命名)",
                    "file": os.path.basename(r[3] or ""), "snippet": snippet})
    con.close()
    return out

def chat_session_touch(session_id, scope=None, db=DEFAULT_DB):
    """会话中途改了范围：同步到会话记录（切回会话时恢复的是最新范围）。"""
    con = _init(db)
    con.execute("UPDATE chat_sessions SET scope_json=?, updated_at=? WHERE id=?",
                (json.dumps(scope, ensure_ascii=False) if scope else None,
                 _now(), session_id))
    con.commit()
    con.close()


# ---------------- 业务规则（值域/格式；block=拒写 warn=警告） ----------------

def add_biz_rule(workbook_id, sheet, col_path, rmin=None, rmax=None,
                 pattern=None, severity="block", db=DEFAULT_DB):
    """覆盖式保存：同(工作簿,sheet,列)重复保存替换旧规则。"""
    con = _init(db)
    con.execute("DELETE FROM biz_rules WHERE workbook_id=? AND sheet_name=? AND col_path=?",
                (workbook_id, sheet, col_path))
    con.execute("INSERT INTO biz_rules(workbook_id, sheet_name, col_path, rmin, rmax, "
                "pattern, severity, created_at) VALUES(?,?,?,?,?,?,?,?)",
                (workbook_id, sheet, col_path, rmin, rmax, pattern, severity, _now()))
    con.commit()
    con.close()


def del_biz_rule(workbook_id, sheet, col_path, db=DEFAULT_DB):
    con = _init(db)
    con.execute("DELETE FROM biz_rules WHERE workbook_id=? AND sheet_name=? AND col_path=?",
                (workbook_id, sheet, col_path))
    con.commit()
    con.close()


def list_biz_rules(workbook_id, sheet=None, db=DEFAULT_DB):
    con = _init(db)
    q = ("SELECT sheet_name, col_path, rmin, rmax, pattern, severity FROM biz_rules "
         "WHERE workbook_id=?")
    args = [workbook_id]
    if sheet:
        q += " AND sheet_name=?"
        args.append(sheet)
    rows = con.execute(q + " ORDER BY sheet_name, col_path", args).fetchall()
    con.close()
    return [dict(r) for r in rows]
