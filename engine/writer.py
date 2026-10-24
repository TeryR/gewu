# -*- coding: utf-8 -*-
"""写回层：操作计划 → 写前硬校验 → COM 定点写入 → 写后校验 → 撤销日志。
纪律（PRD 4.9）：LLM 只产计划 JSON；校验与 IO 全部代码执行；宁可拒绝不可脏写。
写入路径优先级：Excel COM → WPS COM（均保真）→ openpyxl（会丢图表/图片，须显式允许）。
原文件写入前自动备份到 .cache/backup/。"""
import json
import os
import shutil
import sqlite3
import datetime as _dt

from .extract import extract_workbook, _has_value

BACKUP_DIR = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
                          ".cache", "backup")


def _now():
    return _dt.datetime.now().isoformat(timespec="seconds")


class Rejection(Exception):
    pass


# ---------------- 写前硬校验 ----------------

def validate_plan(ops, path, db, workbook_id) -> dict:
    """ops=[{sheet,row,col,value}]，逐条对照冻结镜像与磁盘几何。
    返回 {ok, rejections:[], warnings:[]}；存在任一 rejection 则整批拒绝。"""
    con = sqlite3.connect(db)
    con.row_factory = sqlite3.Row
    wbrow = con.execute("SELECT status FROM workbooks WHERE id=?",
                        (workbook_id,)).fetchone()
    rejections, warnings = [], []
    if not wbrow:
        return {"ok": False, "rejections": ["镜像中无此工作簿"], "warnings": []}
    if wbrow["status"] != "frozen":
        return {"ok": False,
                "rejections": ["工作簿未冻结：写回只认人工冻结后的语义镜像（PRD 4.9）"],
                "warnings": []}
    biz = [dict(r) for r in con.execute(
        "SELECT * FROM biz_rules WHERE workbook_id=?", (workbook_id,))]

    blocks = [dict(r) for r in con.execute(
        "SELECT * FROM blocks WHERE workbook_id=?", (workbook_id,))]
    con.close()
    src = extract_workbook(path)
    sheets = {s["name"]: s for s in src["sheets"]}

    for i, op in enumerate(ops):
        tag = f"op{i}[{op.get('sheet')}!r{op.get('row')}c{op.get('col')}]"
        sh = sheets.get(op["sheet"])
        if not sh:
            rejections.append(f"{tag} 工作表不存在")
            continue
        r, c = op["row"], op["col"]

        blk = next((b for b in blocks if b["sheet_name"] == op["sheet"]
                    and b["r1"] <= r <= b["r2"] and b["c1"] <= c <= b["c2"]), None)
        if not blk:
            rejections.append(f"{tag} 不在任何已冻结数据表区块内（防止跨区块污染）")
            continue
        eff = blk.get("user_type") or blk["type"]   # 有效类型：人工修正优先
        if eff == "ignore":
            rejections.append(f"{tag} 区块已被人工标记为忽略")
            continue
        if eff == "note":
            rejections.append(f"{tag} 区块已被人工标记为文本/批注（非数据表，不参与改数）")
            continue
        if eff == "mixed":
            rejections.append(f"{tag} 区块处于待复核状态，请先在确认页定型后再写")
            continue
        if blk["type"] != "table":
            rejections.append(f"{tag} 区块无列结构（引擎判定为文本），暂不支持写入")
            continue
        if blk.get("readonly"):
            rejections.append(f"{tag} 区块已标记只读🔒")
            continue

        h = json.loads(blk["header_json"] or "{}")
        prot_rows = set(h.get("header_rows") or []) | set(h.get("floating_rows") or [])
        if h.get("title_row"):
            prot_rows.add(h["title_row"])
        if r in prot_rows:
            rejections.append(f"{tag} 位于表头/标题/浮动区，禁止写入")
            continue
        if r in (h.get("summary_rows") or []):
            rejections.append(f"{tag} 位于汇总行（写保护，需显式解锁流程）")
            continue

        info = sh["cells"].get((r, c))
        if info and info.get("f"):
            rejections.append(f"{tag} 目标格是公式格（默认保护，防止破坏取数逻辑）")
            continue
        for mr1, mc1, mr2, mc2 in sh["merged"]:
            if mr1 <= r <= mr2 and mc1 <= c <= mc2 and not (r == mr1 and c == mc1):
                rejections.append(f"{tag} 是合并区非左上格（写入无效，请写左上 {mr1},{mc1}）")
                break

        # 列类型一致性：数值列写文本给警告（可带 confirm 强制）
        col_leaves = [n for n in (h.get("col_tree") or []) if n.get("col") == c]
        if col_leaves and not isinstance(op["value"], (int, float)) \
                and isinstance(op["value"], str) and op["value"].strip():
            nums = tot = 0
            for rr in range(blk["r1"], blk["r2"] + 1):
                v = sh["cells"].get((rr, c), {}).get("cv")
                if isinstance(v, (int, float)) and not isinstance(v, bool):
                    nums += 1
                if _has_value(v):
                    tot += 1
            if tot and nums / tot >= 0.8:
                if not op.get("confirm"):
                    warnings.append(f"{tag} 向数值列写入文本「{op['value'][:12]}」")
        # 业务规则（人配置，引擎硬执行）：数值走值域，文本走正则；block 拒写 / warn 警告
        leaf_paths = ["/".join(n["path"]) for n in
                      (json.loads(blk["header_json"] or "{}").get("col_tree") or [])
                      if n.get("col") == c and n.get("path")]
        for rule in biz:
            if rule["sheet_name"] != op["sheet"]:
                continue
            rp = rule["col_path"]
            if rp not in leaf_paths and not any(
                    p_.split("/")[-1] == rp.split("/")[-1] for p_ in leaf_paths):
                continue
            v = op["value"]
            msg = None
            if isinstance(v, (int, float)) and not isinstance(v, bool):
                if rule["rmin"] is not None and v < rule["rmin"]:
                    msg = f"值 {v} 低于下限 {rule['rmin']}"
                if rule["rmax"] is not None and v > rule["rmax"]:
                    msg = (msg + "；" if msg else "") + f"值 {v} 超过上限 {rule['rmax']}"
            elif isinstance(v, str) and v.strip() and rule.get("pattern"):
                import re as _re
                if not _re.search(rule["pattern"], v):
                    msg = f"值「{v[:16]}」不匹配格式 {rule['pattern']}"
            if msg:
                full = f"{tag} 违反业务规则[{rp}]: {msg}"
                if rule["severity"] == "warn":
                    warnings.append(full)
                else:
                    rejections.append(full)
    return {"ok": not rejections, "rejections": rejections, "warnings": warnings}


# ---------------- 执行：备份 + COM 写入 ----------------

def _backup(path):
    os.makedirs(BACKUP_DIR, exist_ok=True)
    stamp = _dt.datetime.now().strftime("%Y%m%d_%H%M%S")
    dst = os.path.join(BACKUP_DIR, f"{os.path.splitext(os.path.basename(path))[0]}_{stamp}.xlsx")
    shutil.copy2(path, dst)
    return dst


def _com_app(engine_pref=None, retries=1):
    """带重试与错误详情的 COM 启动：同进程内刚 Quit 再 Dispatch 有瞬态失败先例，
    吞异常返回 None 会把可诊断问题变成'COM 不可用'黑盒。"""
    import time
    import pythoncom
    import win32com.client as wc
    last_err = "未尝试"
    for attempt in range(retries + 1):
        try:
            pythoncom.CoInitialize()  # HTTP 工作线程不会自动初始化 COM（CoInitialize 尚未调用）
        except Exception:
            pass
        for progid in ((engine_pref,) if engine_pref
                       else ("Excel.Application", "Ket.Application")):
            try:
                return wc.DispatchEx(progid), progid
            except Exception as e:
                last_err = f"{progid}: {type(e).__name__} {e}"
        if attempt < retries:
            time.sleep(1.5)
    raise Rejection(f"COM 启动失败（{last_err}）")


def execute(ops, path, db, workbook_id, allow_openpyxl_fallback=False,
            engine_pref=None) -> dict:
    """执行已通过校验的写入。返回 {ok, engine, backup, applied, journal_ids}。
    写前护栏：文件 hash 必须与镜像登记一致——被外部改过而未重新解析的文件禁止写入。"""
    from .mirror import _file_hash, apply_ops_to_mirror
    con = sqlite3.connect(db)
    stored = con.execute("SELECT file_hash FROM workbooks WHERE id=?",
                         (workbook_id,)).fetchone()
    con.close()
    if stored and stored[0] and stored[0] != _file_hash(path):
        raise Rejection("文件在冻结后被外部修改（镜像已过期）：请重新拖入核对结构后再写")

    src = extract_workbook(path)
    olds = {}
    for op in ops:
        sh = next(s for s in src["sheets"] if s["name"] == op["sheet"])
        info = sh["cells"].get((op["row"], op["col"]))
        if not info:
            olds[(op["sheet"], op["row"], op["col"])] = None
        elif info.get("f"):
            # 公式格撤销必须恢复公式原文，只存缓存值会让公式永久丢失
            olds[(op["sheet"], op["row"], op["col"])] = info["v"]
        else:
            olds[(op["sheet"], op["row"], op["col"])] = \
                info["cv"] if info["cv"] is not None else info["v"]

    backup = _backup(path)
    try:
        app, progid = _com_app(engine_pref)
    except Rejection:
        if not allow_openpyxl_fallback:
            raise
        import openpyxl
        wb = openpyxl.load_workbook(path)
        for op in ops:
            wb[op["sheet"]].cell(row=op["row"], column=op["col"]).value = op["value"]
        wb.save(path)
        engine = "openpyxl(降级)"
    else:
        try:
            app.Visible = False
            app.DisplayAlerts = False
            wb = app.Workbooks.Open(os.path.abspath(path))
            for op in ops:
                wb.Worksheets(op["sheet"]).Cells(op["row"], op["col"]).Value = op["value"]
            wb.Save()
            wb.Close(SaveChanges=False)
        finally:
            app.Quit()
        engine = progid

    ids = journal(ops, olds, db, workbook_id, backup)
    # 镜像同步：hash 滚到写后的文件 + 只更新受影响记录的值（冻结与标记不动）
    apply_ops_to_mirror(workbook_id, db, path, ops)
    return {"ok": True, "engine": engine, "backup": backup,
            "applied": len(ops), "journal_ids": ids}


def journal(ops, olds, db, workbook_id, backup) -> list:
    con = sqlite3.connect(db)
    con.executescript("""
    CREATE TABLE IF NOT EXISTS write_ops(
        id INTEGER PRIMARY KEY, workbook_id INTEGER, ts TEXT, batch TEXT,
        sheet TEXT, row_no INTEGER, col_no INTEGER,
        old_repr TEXT, new_repr TEXT, backup_path TEXT, undone INTEGER DEFAULT 0);""")
    batch = _now() + f"#{len(ops)}"
    ids = []
    for op in ops:
        old = olds.get((op["sheet"], op["row"], op["col"]))
        cur = con.execute(
            "INSERT INTO write_ops(workbook_id, ts, batch, sheet, row_no, col_no, "
            "old_repr, new_repr, backup_path) VALUES(?,?,?,?,?,?,?,?,?)",
            (workbook_id, _now(), batch, op["sheet"], op["row"], op["col"],
             json.dumps(old, ensure_ascii=False, default=str),
             json.dumps(op["value"], ensure_ascii=False, default=str), backup))
        ids.append(cur.lastrowid)
    con.commit()
    con.close()
    return ids


def undo_last_batch(db, path, workbook_id=None) -> dict:
    """撤销最近一批写入：把 old 值写回（恢复原状绕过保护校验）。
    同样过 hash 护栏：文件被外部改过而未重新解析时，把旧值写回去可能砸掉别人的改动。"""
    from .mirror import apply_ops_to_mirror, _file_hash
    con = sqlite3.connect(db)
    con.row_factory = sqlite3.Row
    # hash 护栏和批次选择都必须限定当前工作簿：多文件镜像库下
    # 按 id 全局取最新会拿错文件/撤错别人的写入
    if workbook_id is not None:
        wbrow = con.execute("SELECT id, file_hash FROM workbooks WHERE id=?",
                            (workbook_id,)).fetchone()
    else:
        wbrow = con.execute("SELECT id, file_hash FROM workbooks ORDER BY id DESC").fetchone()
    if wbrow and wbrow["file_hash"] and wbrow["file_hash"] != _file_hash(path):
        con.close()
        return {"ok": False, "reason": "文件在冻结后被外部修改（镜像已过期）：请重新拖入核对后再撤销"}
    con.row_factory = sqlite3.Row
    if workbook_id is not None:
        row = con.execute("SELECT batch FROM write_ops WHERE undone=0 AND workbook_id=? "
                          "ORDER BY id DESC LIMIT 1", (workbook_id,)).fetchone()
    else:
        row = con.execute("SELECT batch FROM write_ops WHERE undone=0 "
                          "ORDER BY id DESC LIMIT 1").fetchone()
    if not row:
        con.close()
        return {"ok": False, "reason": "无可撤销的写入"}
    batch_rows = con.execute("SELECT * FROM write_ops WHERE batch=? AND undone=0",
                             (row["batch"],)).fetchall()
    con.close()
    ops = [{"sheet": x["sheet"], "row": x["row_no"], "col": x["col_no"],
            "value": json.loads(x["old_repr"])} for x in batch_rows]
    try:
        app, progid = _com_app()
    except Rejection as e:
        return {"ok": False, "reason": f"{e}；可从备份恢复：" + batch_rows[0]["backup_path"]}
    try:
        app.Visible = False
        app.DisplayAlerts = False
        wb = app.Workbooks.Open(os.path.abspath(path))
        for op in ops:
            v = op["value"]
            wb.Worksheets(op["sheet"]).Cells(op["row"], op["col"]).Value = v
        wb.Save()
        wb.Close(SaveChanges=False)
    finally:
        app.Quit()
    con = sqlite3.connect(db)
    con.execute("UPDATE write_ops SET undone=1 WHERE batch=?", (row["batch"],))
    con.commit()
    con.close()
    # 撤销也是一次写入：hash 滚动 + 记录值恢复为旧值
    if workbook_id is None:
        w = sqlite3.connect(db).execute(
            "SELECT id FROM workbooks ORDER BY id DESC LIMIT 1").fetchone()
        workbook_id = w[0] if w else None
    if workbook_id:
        apply_ops_to_mirror(workbook_id, db, path, ops)
    return {"ok": True, "engine": progid, "restored": len(ops), "batch": row["batch"]}


# ---------------- 写后校验 ----------------

def post_check(ops, path, db, workbook_id) -> dict:
    """重读文件：①目标格值=意图值 ②同 Sheet 区块结构未漂移。"""
    from .segment import segment_sheet
    after = extract_workbook(path)
    sheets = {s["name"]: s for s in after["sheets"]}
    problems = []
    for op in ops:
        sh = sheets.get(op["sheet"])
        if not sh:
            problems.append(f"[{op['sheet']}] 消失")
            continue
        info = sh["cells"].get((op["row"], op["col"])) or {}
        got = info.get("cv", info.get("v"))
        want = op["value"]
        if isinstance(want, float) and isinstance(got, (int, float)):
            match = abs(got - want) < 1e-9
        else:
            match = str(got) == str(want)
        if not match:
            problems.append(f"[{op['sheet']}!r{op['row']}c{op['col']}] "
                            f"期望 {want!r} 实读 {got!r}")
    con = sqlite3.connect(db)
    con.row_factory = sqlite3.Row
    # before = 冻结镜像里的生效几何（含人工调整 user_bbox）
    before_blocks, user_geoms = set(), {}
    for r in con.execute(
            "SELECT * FROM blocks WHERE workbook_id=? AND type='table'", (workbook_id,)):
        eff = json.loads(r["user_bbox"]) if r["user_bbox"] else             [r["r1"], r["c1"], r["r2"], r["c2"]]
        before_blocks.add((r["sheet_name"], tuple(eff)))
        if r["user_bbox"]:
            # 引擎 block_id 同文件确定性生成：重分割后套用同一人工范围再比对
            user_geoms[(r["sheet_name"], r["block_id"])] = eff
    con.close()
    for name in {op["sheet"] for op in ops}:
        seg = segment_sheet(sheets[name])
        after_blocks = set()
        for b in seg["blocks"]:
            if b["type"] != "table":
                continue
            eff = user_geoms.get((name, b["id"]), b["bbox"])
            after_blocks.add((name, tuple(eff)))
        gone = {k for k in before_blocks if k[0] == name} - after_blocks
        if gone:
            problems.append(f"[{name}] 区块结构漂移：丢失 {sorted(gone)}")
    return {"ok": not problems, "problems": problems}
