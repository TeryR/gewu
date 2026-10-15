# -*- coding: utf-8 -*-
"""只读问答引擎：基于冻结镜像的查询/聚合/勾稽校验。
纪律（PRD 4.6）：所有数值运算 100% 代码执行，LLM 只做意图解析（可插拔，缺省走规则）。
用法：python -m engine.query "问题" [--db mirror.db] [--wb N]"""
import json
import re
import sqlite3

from .rules import get_rules, is_ratio_column

DEFAULT_DB = "mirror.db"

OPS = ("sum", "avg", "count", "max", "min")


class Mirror:
    def __init__(self, db=DEFAULT_DB, workbook_id=None, rules=None):
        self.con = sqlite3.connect(db)
        self.con.row_factory = sqlite3.Row
        self.rules = rules or get_rules()
        if workbook_id is None:
            row = self.con.execute(
                "SELECT id FROM workbooks ORDER BY id DESC LIMIT 1").fetchone()
            if not row:
                raise SystemExit("镜像库为空，先运行 build_mirror")
            workbook_id = row["id"]
        self.wid = workbook_id

    # ---------- 基础访问 ----------
    def close(self):
        self.con.close()

    def glossary(self):
        """企业口径词典：人工确认的字段口径，注入 LLM 上下文。"""
        rows = self.con.execute(
            "SELECT term, meaning, sheet_name, col_path, created_at FROM glossary "
            "WHERE workbook_id=? ORDER BY id", (self.wid,)).fetchall()
        return [dict(r) for r in rows]

    def biz_rules(self, sheet=None):
        """业务规则（值域）：人配置、引擎硬执行；注入 LLM 上下文与写校验。"""
        q = ("SELECT sheet_name, col_path, rmin, rmax, pattern, severity "
             "FROM biz_rules WHERE workbook_id=?")
        args = [self.wid]
        if sheet:
            q += " AND sheet_name=?"
            args.append(sheet)
        return [dict(r) for r in self.con.execute(q, args).fetchall()]

    def rules_check(self, sheet=None):
        """全表规则体检：对镜像全部明细记录跑值域规则，返回违规清单（纯代码）。"""
        out = []
        rules = self.biz_rules(sheet)
        if not rules:
            return out
        by_sheet = {}
        for r in rules:
            by_sheet.setdefault(r["sheet_name"], []).append(r)
        for b in self.blocks(tables_only=True):
            for rule in by_sheet.get(b["sheet_name"], []):
                path = self._match_column(b, rule["col_path"])
                for rec in self._records(b):
                    v = rec.get(path)
                    if rec.get("_summary") or v in (None, ""):
                        continue
                    bad = None
                    if isinstance(v, (int, float)) and not isinstance(v, bool):
                        if rule["rmin"] is not None and v < rule["rmin"]:
                            bad = f"低于下限 {rule['rmin']}"
                        if rule["rmax"] is not None and v > rule["rmax"]:
                            bad = (bad + "；" if bad else "") + f"超过上限 {rule['rmax']}"
                    elif isinstance(v, str) and rule.get("pattern"):
                        if not re.search(rule["pattern"], v):
                            bad = f"不匹配格式 {rule['pattern']}"
                    if bad:
                        out.append({"sheet": b["sheet_name"], "row": rec["_r"],
                                    "column": path, "value": v, "问题": bad,
                                    "severity": rule["severity"]})
        return out

    def row_notes(self, sheet=None):
        """行内备注（剥离黄块挂回表行的批注）：LLM 上下文与行查询可见。"""
        q = ("SELECT sheet_name, note_block_id, table_block_id, row_no, text "
             "FROM row_notes WHERE workbook_id=?")
        args = [self.wid]
        if sheet:
            q += " AND sheet_name=?"
            args.append(sheet)
        rows = self.con.execute(q + " ORDER BY sheet_name, row_no", args).fetchall()
        return [dict(r) for r in rows]

    def blocks(self, sheet=None, tables_only=False, scope=None):
        q = "SELECT * FROM blocks WHERE workbook_id=?"
        args = [self.wid]
        if sheet:
            q += " AND sheet_name=?"
            args.append(sheet)
        if scope:   # 对话范围：只看圈定的 sheet（None=全簿）
            q += " AND sheet_name IN (%s)" % ",".join("?" * len(scope))
            args += list(scope)
        if tables_only:
            # 人工切过类型的块按有效类型对待：切走（note/ignore/mixed）的表
            # 不得继续出现在问答/勾稽里——确认页的承诺必须与下游行为一致
            q += " AND type='table' AND IFNULL(user_type,'table')='table'"
        return self.con.execute(q + " ORDER BY sheet_name, r1, c1", args).fetchall()

    def _records(self, block_row):
        rows = self.con.execute(
            "SELECT row_no, is_summary, data_json FROM records WHERE block_db_id=? "
            "ORDER BY row_no", (block_row["id"],)).fetchall()
        out = []
        for r in rows:
            d = json.loads(r["data_json"])
            d["_summary"] = bool(r["is_summary"])
            out.append(d)
        return out

    def _disp(self, v, pct):
        # 百分比列：任何数值（含整数，如缓存值恰为 1 表示 100%）一律按比率放大显示
        if pct and isinstance(v, (int, float)) and not isinstance(v, bool):
            return f"{v * 100:.2f}%"
        if isinstance(v, float):
            return f"{v:,.4g}"
        return v

    # ---------- 聚合与查询 ----------
    def _pct_cols(self, block):
        try:
            meta = json.loads(block["summary_formulas_json"] or "{}")
            return set(meta.get("pct_cols") or [])
        except (KeyError, TypeError):
            return set()

    def aggregate(self, block, column, op="sum", filters=None, include_summary=False):
        recs = self._records(block)
        if filters:
            recs = [r for r in recs
                    if not r.get("_summary") or include_summary
                    if all(str(r.get(k, "")) == str(v) or v in str(r.get(k, ""))
                           for k, v in filters.items())]
        else:
            recs = [r for r in recs if include_summary or not r.get("_summary")]
        col = self._match_column(block, column)
        vals = [r[col] for r in recs
                if col in r and isinstance(r[col], (int, float)) and not isinstance(r[col], bool)]
        if not vals:
            return None
        f = {"sum": sum, "avg": lambda x: sum(x) / len(x),
             "count": len, "max": max, "min": min}[op]
        return {"op": op, "column": col, "n": len(vals),
                "value": f(vals) if op != "count" else len(vals),
                "provenance": f"[{block['sheet_name']}] r{block['r1']}-r{block['r2']} "
                              f"明细 {len(vals)} 行（已剔除汇总行）"}

    def lookup(self, block, filters, limit=None):
        """命中全部返回（limit 仅用于显示截断）。调用方展示前几行并报真实总数。"""
        recs = self._records(block)
        hits = []
        for r in recs:
            ok = True
            for k, v in filters.items():
                col = self._match_column(block, k)
                if col not in r or str(v) not in str(r[col]):
                    ok = False
                    break
            if ok:
                hits.append(r)
        return hits if limit is None else hits[:limit]

    def check_ties(self, block):
        """公式感知勾稽（纯代码）：只重算汇总格自己的 =SUM(区间)[/常数]，
        比值/跨表/其他口径归类跳过；另检查汇总区间是否漏行（公式陈旧）。"""
        meta = json.loads(block["summary_formulas_json"]
                          or '{"formulas": {}, "col_of_path": {}}')
        formulas = meta.get("formulas") or {}
        path_of_col = {col: path for path, col in (meta.get("col_of_path") or {}).items()}
        recs = {r["_r"]: r for r in self._records(block)}
        detail_rows = sorted(r for r, d in recs.items() if not d.get("_summary"))
        sum_pat = re.compile(
            r"=SUM\(\$?([A-Z]{1,3})\$?(\d+):\$?([A-Z]{1,3})\$?(\d+)\)(?:\s*/\s*([\d.]+))?",
            re.I)

        def col_idx(letters):
            n = 0
            for ch in letters.upper():
                n = n * 26 + ord(ch) - 64
            return n

        results = []
        for key, formula in formulas.items():
            r_s, path = key.split("|", 1)
            r_s = int(r_s)
            stored = recs.get(r_s, {}).get(path)
            if not isinstance(stored, (int, float)) or isinstance(stored, bool):
                continue
            m = sum_pat.search(formula)
            if not m:
                if "!" in formula:
                    kind = "skip_cross"      # 跨表取数，非本块勾稽
                elif "/" in formula or "*" in formula:
                    kind = "skip_ratio"      # 比值/率/人均口径，不可加总
                else:
                    kind = "skip_other"
                results.append({"row": r_s, "column": path, "stored": stored,
                                "kind": kind, "formula": formula[:60]})
                continue
            ca, ra, cb, rb, divisor = m.groups()
            if col_idx(ca) != col_idx(cb):
                results.append({"row": r_s, "column": path, "stored": stored,
                                "kind": "skip_rowwise", "formula": formula[:60]})
                continue
            target_path = path_of_col.get(col_idx(ca))
            total, n = 0.0, 0
            for rr in range(int(ra), int(rb) + 1):
                v = recs.get(rr, {}).get(target_path)
                if isinstance(v, (int, float)) and not isinstance(v, bool):
                    total += v
                    n += 1
            if divisor:
                total /= float(divisor)
            tol = max(abs(total) * self.rules["query"]["tie_rel_tol"],
                      self.rules["query"]["tie_abs_tol"])
            kind = "ok" if abs(total - stored) <= tol else "bad"
            # 结构检查：公式区间终点是否落后于其实际明细末行（插行后公式陈旧）
            stale = None
            below = [r for r in detail_rows if int(rb) < r < r_s]
            if below:
                stale = f"区间止于 r{rb}，其后仍有明细行 r{below[0]}..r{below[-1]} 未纳入"
            results.append({"row": r_s, "column": path, "stored": stored,
                            "computed": round(total, 4), "n": n,
                            "kind": kind, "formula": formula[:60], "stale": stale})
        # 兜底：无公式的硬编码汇总格——过滤率/占比/人均列后做上方明细求和比对。
        # 无公式佐证，确定性低于公式重算，单独标 naive_bad（存疑）不算实锤偏差
        qcfg = self.rules["query"]
        processed = {(x["row"], x["column"]) for x in results}
        for sr in sorted(r for r, d in recs.items() if d.get("_summary")):
            for path, v in recs[sr].items():
                if path.startswith("_") or not isinstance(v, (int, float)) \
                        or isinstance(v, bool) or (sr, path) in processed:
                    continue
                if is_ratio_column(path, self.rules):
                    continue
                nums = [recs[r][path] for r in recs
                        if r < sr and not recs[r].get("_summary")
                        and isinstance(recs[r].get(path), (int, float))
                        and not isinstance(recs[r][path], bool)]
                if len(nums) < qcfg["naive_min_rows"]:
                    continue
                calc = sum(nums)
                tol = max(abs(calc) * qcfg["naive_rel_tol"], qcfg["naive_abs_tol"])
                if abs(calc - v) > tol:
                    results.append({"row": sr, "column": path, "stored": v,
                                    "computed": round(calc, 4), "n": len(nums),
                                    "kind": "naive_bad"})
        return results

    def _match_column(self, block, needle):
        h = json.loads(block["header_json"] or "{}")
        paths = []
        for node in h.get("col_tree") or []:
            p = "/".join(node["path"]) if node["path"] else f"col{node['col']}"
            paths.append(p)
        needle = str(needle).strip()
        for p in paths:
            if needle == p:
                return p
        for p in paths:
            if needle and needle in p:
                return p
        for p in paths:
            if needle and needle in p.split("/")[-1]:
                return p
        return needle  # 原样返回，允许直接用记录里的键

    def _pick_block_for_column(self, needle, scope=None):
        """按列名反查区块：同名列在不同表里内容形态可能完全不同
        （如[代理]的"年完成率"列装的是 REPT 进度条字符串），
        优先选该列真有数值行的表，数值行数为主序、面积为次序。"""
        bs = self.blocks(tables_only=True, scope=scope)
        if not bs:
            raise SystemExit("镜像中无数据表区块（或范围内没有表）")

        def area(b):
            return (b["r2"] - b["r1"]) * (b["c2"] - b["c1"])

        hits = []
        for b in bs:
            h = json.loads(b["header_json"] or "{}")
            paths = ["/".join(n["path"]) for n in (h.get("col_tree") or []) if n["path"]]
            if any(needle in p or p.split("/")[-1] == needle for p in paths):
                hits.append(b)
        if not hits:
            return max(bs, key=area)

        def numeric_rows(b):
            col = self._match_column(b, needle)
            return sum(1 for r in self._records(b)
                       if not r.get("_summary")
                       and isinstance(r.get(col), (int, float))
                       and not isinstance(r.get(col), bool))

        return max(hits, key=lambda b: (numeric_rows(b), area(b)))

    def _pick_block(self, needle=None, scope=None):
        bs = self.blocks(tables_only=True, scope=scope)
        if not bs:
            raise SystemExit("镜像中无数据表区块")
        if needle:
            for b in bs:
                if needle in b["sheet_name"]:
                    return b
            for b in bs:
                h = json.loads(b["header_json"] or "{}")
                if h.get("title_text") and needle in h["title_text"]:
                    return b
        return max(bs, key=lambda b: (b["r2"] - b["r1"]) * (b["c2"] - b["c1"]))

    # ---------- NL 前端（规则版，LLM 可插拔替换） ----------
    def answer(self, question: str, llm=None, sheet_hint=None, history=None, scope=None, lang=None):
        q = question.strip()
        # 语言指令：UI 中/EN 开关驱动，解决中文模型对英文提问仍回中文的偏向
        _lang_note = "[Respond in English]" if lang == "en" else ""
        if _lang_note:
            q = q + _lang_note
        if llm:  # 结构化升级路径：LLM 只输出查询 JSON，不输出数值
            kw = {"history": history} if history else {}
            plan = llm(q, context=self._llm_context(sheet_hint, scope), **kw)
            if plan.get("op") in ("describe", "analyze") and not plan.get("block_hint"):
                # "这个表格"指代优先级：范围唯一 sheet > 当前页（且在范围内）
                if scope and len(scope) == 1:
                    plan["block_hint"] = scope[0]
                elif plan.get("op") == "analyze" and not scope:
                    plan["block_hint"] = "*"   # 评价类无明确目标 → 全簿（"设计合理吗"问的是整体）
                elif sheet_hint and (not scope or sheet_hint in scope):
                    plan["block_hint"] = sheet_hint
            return self._execute_plan(plan, question=q, llm=llm, history=history, scope=scope)
        if re.search(r"勾稽|核对|对不上|验", q):
            return self._answer_ties(q)
        if re.search(r"是做什么|是干嘛|什么用途|介绍|有哪些字段|哪些列|字段.*口径|口径", q):
            return self._describe_fallback(self._pick_block(None), q)
        if re.search(r"评价|分析|解读|趋势|哪个更好|怎么看", q):
            return self._analyze_fallback(self._pick_block(sheet_hint, scope))
        m = re.search(r"(最高|最大|最低|最小)", q)
        if m and not re.search(r"合计|总和", q):
            col = self._guess_column(q) or "数量"
            return self._answer_extreme(q, col, m.group(1))
        if re.search(r"合计|总和|一共|总共", q):
            col = self._guess_column(q)
            return self._answer_sum(q, col)
        return self._answer_help()

    def _llm_context(self, sheet_hint=None, scope=None):
        notes = [{"sheet": g["sheet_name"], "column": g["col_path"] or g["term"],
                  "口径": g["meaning"]} for g in self.glossary()]
        remarks = [{"sheet": r["sheet_name"], "block": r["table_block_id"],
                    "row": r["row_no"], "备注": r["text"]}
                   for r in self.row_notes()[:30]]
        ctx = []
        for b in self.blocks(tables_only=True, scope=scope):
            h = json.loads(b["header_json"] or "{}")
            paths = ["/".join(n["path"]) for n in (h.get("col_tree") or []) if n["path"]]
            # 多 sheet 文件瘦身：非当前页只带前 12 列（够 LLM 选表），当前页带全量选列
            cap = 30 if (sheet_hint and sheet_hint in b["sheet_name"]) else 12
            ctx.append({"sheet": b["sheet_name"], "block": b["block_id"],
                        "title": h.get("title_text"), "columns": paths[:cap]})
        out = {"tables": ctx}
        if sheet_hint:
            out["当前页"] = sheet_hint   # 让"这个表格/当前表格"有明确指代
        if notes:
            out["field_notes"] = notes
        if remarks:
            out["row_remarks"] = remarks
        return out

    def _describe_ctx(self, block):
        """理解型问答的上下文：标题/列/样例行/行数/百分比列/列公式/口径/行内备注。"""
        h = json.loads(block["header_json"] or "{}")
        meta = json.loads(block["summary_formulas_json"] or "{}")
        paths = ["/".join(n["path"]) for n in (h.get("col_tree") or []) if n["path"]]
        recs = [r for r in self._records(block) if not r.get("_summary")]
        samples = [{k: v for k, v in r.items() if not k.startswith("_")}
                   for r in recs[:3]]
        gl = {g["term"]: g["meaning"] for g in self.glossary()
              if not g["sheet_name"] or g["sheet_name"] == block["sheet_name"]}
        remarks = [r["text"] for r in self.row_notes(block["sheet_name"])
                   if r["table_block_id"] == block["block_id"]][:10]
        return {"sheet": block["sheet_name"], "title": h.get("title_text"),
                "columns": paths[:30], "n_rows": len(recs),
                "pct_columns": sorted(self._pct_cols(block)) or None,
                "列公式样例": (meta.get("col_formulas") or None),
                "sample_rows": samples,
                "字段口径": gl or None, "行内备注": remarks or None}

    def _describe_workbook_ctx(self):
        """全簿模式上下文：各 sheet 的表/行数/关键列，回答"整个文件是做什么的"。"""
        sheets = []
        for b in self.blocks(tables_only=True):
            h = json.loads(b["header_json"] or "{}")
            paths = ["/".join(n["path"]) for n in (h.get("col_tree") or []) if n["path"]]
            n = len([r for r in self._records(b) if not r.get("_summary")])
            sheets.append({"sheet": b["sheet_name"], "block": b["block_id"],
                           "title": h.get("title_text"), "n_rows": n,
                           "columns": paths[:8]})
        return {"workbook": True,
                "n_sheets": len({s["sheet"] for s in sheets}), "sheets": sheets,
                "字段口径条数": len(self.glossary()) or None,
                "行内备注条数": len(self.row_notes()) or None}

    def _describe_workbook_fallback(self):
        ctx = self._describe_workbook_ctx()
        out = [f"全簿共 {ctx['n_sheets']} 个 sheet、{len(ctx['sheets'])} 张数据表："]
        for s in ctx["sheets"]:
            cols = "、".join(s["columns"][:4]) + ("…" if len(s["columns"]) > 4 else "")
            out.append(f"  · {s['sheet']}：{s['n_rows']} 行明细，列如 {cols}")
        out.append("（结构描述由代码生成；配置 LLM key 后可获得业务语义解读）")
        return "\n".join(out)

    def _describe_fallback(self, block, q):
        """无 LLM / LLM 挂了时的结构化描述（纯代码，无编造空间）。"""
        ctx = self._describe_ctx(block)
        head = f"[{ctx['sheet']}] {ctx.get('title') or ''}".rstrip()
        cols = "、".join(ctx["columns"][:12]) + ("…" if len(ctx["columns"]) > 12 else "")
        n = len(self._records(block))
        out = [f"{head}", f"共 {n} 行明细，{len(ctx['columns'])} 个字段：{cols}"]
        if ctx["字段口径"]:
            out.append("字段口径：" + "；".join(f"{k}={v}" for k, v in
                                              list(ctx["字段口径"].items())[:5]))
        if ctx["行内备注"]:
            out.append(f"行内备注 {len(ctx['行内备注'])} 条，如：{ctx['行内备注'][0][:50]}")
        out.append("（结构描述由代码生成；配置 LLM key 后可获得业务语义解读）")
        return "\n".join(out)

    def _analyze_ctx(self, block):
        """数据解读上下文：明细行（百分比按%格式化）+ 引擎算好的各列统计量。
        数值纪律：统计量由代码计算，LLM 只引用不计算。"""
        pct = self._pct_cols(block)
        recs = [r for r in self._records(block) if not r.get("_summary")]
        rows = []
        for r in recs[:int(self.rules['query'].get('analyze_rows',25))]:
            d = {}
            for k, v in r.items():
                if k.startswith("_") or v in (None, ""):
                    continue
                d[k] = self._disp(v, k in pct) if isinstance(v, (int, float)) \
                    and not isinstance(v, bool) else v
            rows.append(d)
        stats = {}
        for r in recs:
            for k, v in r.items():
                if k.startswith("_") or not isinstance(v, (int, float)) \
                        or isinstance(v, bool):
                    continue
                s = stats.setdefault(k, {"n": 0, "min": v, "max": v, "sum": 0})
                s["n"] += 1
                s["min"] = min(s["min"], v)
                s["max"] = max(s["max"], v)
                s["sum"] += v
        for k, s in stats.items():
            s["avg"] = round(s.pop("sum") / s["n"], 4)
            s["min"] = self._disp(s["min"], k in pct)
            s["max"] = self._disp(s["max"], k in pct)
            s["avg"] = self._disp(s["avg"], k in pct)
        base = self._describe_ctx(block)
        base.update({"records": rows, "列统计": dict(list(stats.items())[:int(self.rules['query'].get('analyze_stats_cols',12))])})
        return base

    def _analyze_workbook_ctx(self, scope=None):
        """全簿设计评价上下文：逐表结构画像（引擎事实）+ 跨表观察，LLM 只解读。"""
        tables = []
        for b in self.blocks(tables_only=True, scope=scope)[:20]:
            h = json.loads(b["header_json"] or "{}")
            paths = ["/".join(n["path"]) for n in (h.get("col_tree") or []) if n["path"]]
            leaves = [p.split("/")[-1] for p in paths]
            dups = sorted({x for x in leaves if leaves.count(x) > 1})
            meta = json.loads(b["summary_formulas_json"] or "{}")
            tables.append({
                "sheet": b["sheet_name"], "n_rows": len(self._records(b)),
                "n_cols": len(paths), "columns": paths[:14],
                "重复列名": dups or None,
                "公式列数": len(meta.get("col_formulas") or {}),
                "百分比列数": len(meta.get("pct_cols") or [])})
        sheets = [s["name"] for s in self.con.execute(
            "SELECT name FROM sheets WHERE workbook_id=? AND visible=1", (self.wid,))]
        prefix = {}
        for name in sheets:                      # 近亲 sheet 分组（如 15 个"汇总表"变体）
            key = name.split(" ")[0].split("(")[0].strip()
            prefix[key] = prefix.get(key, 0) + 1
        return {"workbook": True, "n_sheets": len(sheets), "sheets": sheets,
                "表结构画像": tables,
                "近亲sheet分组": {k: v for k, v in prefix.items() if v > 1} or None}

    def _analyze_workbook_fallback(self, scope=None):
        ctx = self._analyze_workbook_ctx(scope)
        out = [f"全簿 {ctx['n_sheets']} 个 sheet、{len(ctx['表结构画像'])} 张数据表（引擎结构画像）："]
        for t in ctx["表结构画像"]:
            out.append(f"  {t['sheet']}: {t['n_rows']}行×{t['n_cols']}列"
                       + (f"，重复列名{'、'.join(t['重复列名'][:4])}" if t["重复列名"] else ""))
        out.append("（结构画像由代码生成；配置 LLM key 后可获得设计评价）")
        return chr(10).join(out)

    def _analyze_fallback(self, block):
        ctx = self._analyze_ctx(block)
        out = [f"[{ctx['sheet']}] 引擎统计（共 {ctx['n_rows']} 行明细）："]
        for col, s in (ctx.get("列统计") or {}).items():
            out.append(f"  {col}: n={s['n']} 最小={s['min']} 最大={s['max']} 均值={s['avg']}")
        out.append("（数值由代码计算；配置 LLM key 后可获得业务解读与评价）")
        return "\n".join(out)

    def _execute_plan(self, plan, question=None, llm=None, history=None, scope=None):
        op = plan.get("op", "sum")
        if op == "unknown" or (not plan.get("column") and op not in ("describe", "analyze")):
            return plan.get("need_clarify") or \
                "无法确定查询意图。试试「xx列合计」「xx最高的」「这个表格是做什么的」「分析评价一下这表」或「勾稽核对」。"
        kw = {"history": history} if history else {}
        if op == "analyze":  # 数据解读：引擎算统计，LLM 只解读
            hint = plan.get("block_hint") or ""
            if hint == "*":  # 全簿：评价整体设计/组织
                try:
                    from .llm import analyze_block
                    return analyze_block(question or "评价这个文件的设计",
                                         self._analyze_workbook_ctx(scope), **kw)
                except SystemExit:
                    return self._analyze_workbook_fallback(scope)
            b = self._pick_block(hint or None, scope)
            try:
                from .llm import analyze_block
                return analyze_block(question or "分析评价这份数据",
                                     self._analyze_ctx(b), **kw)
            except SystemExit:
                return self._analyze_fallback(b)
        if op == "describe":
            hint = plan.get("block_hint") or ""
            if hint == "*":  # 全簿模式：用户问的是整个文件
                try:
                    from .llm import describe_block
                    return describe_block(question or "这个文件是做什么的",
                                          self._describe_workbook_ctx(), **kw)
                except SystemExit:
                    return self._describe_workbook_fallback()
            b = self._pick_block(hint or None, scope)
            try:
                from .llm import describe_block
                return describe_block(question or "这个表格是做什么的",
                                      self._describe_ctx(b), **kw)
            except SystemExit:  # LLM 挂了 → 结构化兜底，不让工作台哑掉
                return self._describe_fallback(b, question or "")
        if op == "ties":
            return self._answer_ties("")
        if op == "lookup":
            b = self._pick_block(plan.get("block_hint"), scope)
            hits = self.lookup(b, plan.get("filters") or {})
            if not hits:
                return "没有命中记录"
            pctset = self._pct_cols(b)
            lines = []
            shown = hits[:int(self.rules['query'].get('lookup_show_rows', 8))]
            for h_row in shown:
                cells = {k: (self._disp(v, k in pctset)
                             if isinstance(v, (int, float)) and not isinstance(v, bool) else v)
                         for k, v in h_row.items()
                         if not k.startswith("_") and v not in (None, "")}
                line = "  " + json.dumps(cells, ensure_ascii=False, default=str)[:200]
                if h_row.get("_note"):
                    line += f"  备注:{str(h_row['_note'])[:80]}"
                lines.append(line)
            # 行数多时自动附带汇总（用户要的是结论不是数据罗列）
            head = f"命中 {len(hits)} 行" + \
                (f"（显示前 {len(shown)} 行）" if len(hits) > len(shown) else "") + "："
            if len(hits) > len(shown):
                # 对数值列自动算汇总统计
                num_cols = {}
                for r in hits:
                    for k, v in r.items():
                        if not k.startswith("_") and isinstance(v, (int, float)) \
                                and not isinstance(v, bool) and not r.get("_summary"):
                            num_cols.setdefault(k, []).append(v)
                if num_cols:
                    head += "\n**汇总**："
                    for col, vals in list(num_cols.items())[:4]:
                        pct = col in pctset
                        head += (f"\n- `{col}`: 合计 {self._disp(sum(vals), pct)}"
                                 f" · 均值 {self._disp(sum(vals)/len(vals), pct)}"
                                 f" · 最大 {self._disp(max(vals), pct)}"
                                 f" · 最小 {self._disp(min(vals), pct)}")
            return head + "\n" + "\n".join(lines)
        b = self._pick_block_for_column(plan["column"], scope)
        r = self.aggregate(b, plan["column"], op if op in OPS else "sum",
                           filters=plan.get("filters"))
        return self._fmt_agg(r, b) if r else f"列「{plan['column']}」没有可计算的数值"

    def _guess_column(self, q):
        cands = set()
        for b in self.blocks(tables_only=True):
            h = json.loads(b["header_json"] or "{}")
            for node in h.get("col_tree") or []:
                if node["path"]:
                    cands.add(node["path"][-1])
        for c in sorted(cands, key=len, reverse=True):
            if c and c in q:
                return c
        return None

    def _answer_sum(self, q, col):
        col = col or self._guess_column(q)
        if not col:
            return "没猜到要合计哪一列，试试「xx列合计」"
        b = self._pick_block_for_column(col)
        r = self.aggregate(b, col, "sum")
        return self._fmt_agg(r, b) if r else f"列「{col}」没有可合计的数值"

    def _answer_extreme(self, q, col, word):
        op = "max" if word in ("最高", "最大") else "min"
        col = self._match_column_probe(col) or col
        b = self._pick_block_for_column(col)
        r = self.aggregate(b, col, op)
        if not r:
            return f"列「{col}」无数值"
        recs = [x for x in self._records(b) if not x.get("_summary")]
        key = r["column"]
        best = max(recs, key=lambda x: x.get(key, -1e18) if isinstance(x.get(key), (int, float))
                   else -1e18) if op == "max" else \
               min(recs, key=lambda x: x.get(key, 1e18) if isinstance(x.get(key), (int, float))
                   else 1e18)
        who = {k: v for k, v in best.items() if not k.startswith("_") and v not in (None, "")}
        pct = key in self._pct_cols(b)
        return (f"{word}的「{key}」= {self._disp(r['value'], pct)}"
                f"（{json.dumps(who, ensure_ascii=False, default=str)[:200]}）\n{r['provenance']}")

    def _match_column_probe(self, needle):
        """把口语列名解析成镜像里真实存在的列（找不到返回 None）。"""
        if not needle:
            return None
        for b in self.blocks(tables_only=True):
            h = json.loads(b["header_json"] or "{}")
            for node in h.get("col_tree") or []:
                if node["path"]:
                    p = "/".join(node["path"])
                    if needle in p or needle == p.split("/")[-1]:
                        return p.split("/")[-1]
        return None

    def _answer_ties(self, q):
        skip_names = {"skip_ratio": "比值/率口径", "skip_cross": "跨表取数",
                      "skip_other": "其他公式", "skip_rowwise": "行内合计列"}
        out = ["公式感知勾稽（只重算 =SUM 区间，纯代码零模型）："]
        n_bad = n_ok = 0
        for b in self.blocks(tables_only=True):
            res = self.check_ties(b)
            bad = [x for x in res if x["kind"] == "bad"]
            naive = [x for x in res if x["kind"] == "naive_bad"]
            stale = [x for x in res if x.get("stale")]
            skips = {}
            for x in res:
                if x["kind"].startswith("skip"):
                    skips[skip_names.get(x["kind"], x["kind"])] = \
                        skips.get(skip_names.get(x["kind"], x["kind"]), 0) + 1
            if not res:
                out.append(f"  [{b['sheet_name']}] 区块{b['block_id']}：无汇总行")
                continue
            marks = []
            if bad:
                marks.append(f"✗ 偏差{len(bad)}")
            if stale:
                marks.append(f"⚠ 漏行{len(stale)}")
            if naive:
                marks.append(f"⚠ 存疑{len(naive)}(硬编码值,朴素比对)")
            tag = " ".join(marks) if marks else "✓"
            skip_s = f"，跳过{','.join(f'{k}x{v}' for k, v in skips.items())}" if skips else ""
            out.append(f"  [{b['sheet_name']}] 区块{b['block_id']} r{b['r1']}-r{b['r2']}："
                       f"{tag}{skip_s}")
            for x in bad[:4]:
                out.append(f"      ✗ r{x['row']} 「{x['column']}」 表内={x['stored']} "
                           f"公式重算={x['computed']}（{x['formula']}）")
            for x in naive[:3]:
                out.append(f"      ⚠ r{x['row']} 「{x['column']}」 表内={x['stored']} "
                           f"朴素重算={x['computed']}（无公式，仅供参考）")
            for x in stale[:4]:
                out.append(f"      ⚠ r{x['row']} 「{x['column']}」 {x['stale']}")
            n_bad += len(bad)
            n_ok += sum(1 for x in res if x["kind"] == "ok")
        out.append(f"\nSUM 重算 {n_ok} 处相符、{n_bad} 处实锤偏差；"
                   "硬编码值仅朴素比对标存疑；偏差与漏行的业务原因解读留给 LLM，数值以本结果为准。")
        return "\n".join(out)

    def _fmt_agg(self, r, block=None):
        pct = bool(block) and r["column"] in self._pct_cols(block)
        return (f"{r['op']}({r['column']}) = {self._disp(r['value'], pct)} "
                f"（{r['n']} 行明细）\n{r['provenance']}")

    def _answer_help(self):
        return ("当前是纯代码问答引擎（LLM 未接入），支持：\n"
                "  · 「xx列合计」 —— 明细求和（自动剔除汇总行）\n"
                "  · 「xx最高的」 —— 极值定位\n"
                "  · 「勾稽核对」 —— 全簿汇总行代码重算比对\n"
                "接入 LLM 后自然语言能力由模型解析、数值仍由本引擎计算。")


if __name__ == "__main__":
    import sys
    sys.stdout.reconfigure(encoding="utf-8")
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    db = "mirror.db"
    for i, a in enumerate(sys.argv):
        if a == "--db":
            db = sys.argv[i + 1]
    m = Mirror(db)
    print(m.answer(args[0] if args else "勾稽核对"))
