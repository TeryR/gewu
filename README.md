# Gewu (格物) — Excel Understanding Agent

**An open-source agent that reads real-world spreadsheets the way a careful analyst does — including the messy ones nobody admits to making.**

Merged cells. Multi-level headers. Floating meta rows. Annotation lines. A second table hiding at row 23 with restarted serial numbers. Every tool you've tried reads that as garbage. Gewu segments it, builds a header tree, and lets an LLM answer questions over the actual structure — then, if you own the file, writes results back through real Excel COM.

- **Language**: Python (engine) · Tauri (desktop shell)
- **License**: AGPL-3.0 (commercial licensing available)
- **Status**: preparing public release · challenge corpus included

> 中文说明：[README.zh-CN.md](README.zh-CN.md)

---

## Why Gewu

Real business spreadsheets are not tables. They are a title row, a units line, a three-level merged header, twenty data rows with category columns merged vertically, a subtotal, an annotation merged across the full width, and then a second table that starts without warning.

LLMs misread that. So do most parsers.

Gewu's engine solves it with **zero model calls** — pure structural rules:

1. **Table segmentation** — multi-clue block carving: empty-band cutting, merge continuity, perimeter borders, annotation-row stripping, summary-row re-attachment. A sheet containing two tables plus a note comes out as three clean blocks.
2. **Header tree** — multi-level merged headers become a column tree with paths (`本月发生 / 金额`), floating meta rows identified separately, cross-tab and key columns detected, duplicate names disambiguated.
3. **LLM-ready transcription** — the deconstructed structure rendered as text an LLM can actually reason over: hierarchy, empty-value semantics (a column that's 3/3 empty because it updates weekly is *labeled as such*, not silently dropped), and data rows.
4. **Write-back** — results go back into the real .xlsx through Excel COM, with diff location. Your original file is never modified.

## The Challenge Corpus

`tests/golden/` ships with a structured challenge set built from real-world wild tables:

| Series | What breaks naive parsers |
|---|---|
| T1 结构迷宫 Structure Maze | merged fee categories across row groups |
| T2 间接寻址 Indirect Addressing | dynamic region lookups |
| T3 循环求解 Circular Solve | tax and bonus pool interlocks |
| T4 埋错狩猎 Buried Error Hunt | deliberately planted inconsistencies |
| T5 口径炼狱 Caliber Purgatory | same metric, five definitions |
| T6 野路子结构 Wild Structure | the kitchen-sink ledger (see our test walkthrough) |
| T7 碎片拼图 Fragment Assembly | tables split across scattered ranges |
| T8 格式密码 Format Cipher | full-width digits, comma styles, text numbers |

Every golden file is a regression test: if a change alters how any challenge is understood, the suite fails before you ship.

## Quick Start

```bash
git clone https://github.com/TeryR/gewu.git
cd gewu
pip install -r requirements.txt
cp llm.json.example llm.json   # add your API key (OpenAI-compatible endpoints work)
python cli.py
```

> Windows + local Excel required for write-back (COM). Reading and analysis work anywhere Python runs.

## Architecture

```
engine/
├── extract.py    dual-load (formulas + cached values), borders, merges, hidden rows
├── segment.py    multi-clue table segmentation, annotation carving, summary re-attachment
├── headers.py    column tree from merged multi-level headers
├── normalize.py  cleanup and sanitization
├── rules.py      thresholds and vocabulary (rules.json)
├── query.py      LLM Q&A over the structured representation
├── writer.py     write-back via Excel COM
├── mirror.py     file watching and lineage
└── llm.py        provider abstraction (BYOK, OpenAI-compatible)
```

## Roadmap

- [ ] Public release with full challenge corpus
- [ ] English UI
- [ ] macOS / LibreOffice write-back path
- [ ] Hosted version (no install)

## License

AGPL-3.0 — use it, fork it, ship it (open your modifications). Closed-source commercial licensing: open an issue or contact the maintainer.

---

中文文档与说明：[README.zh-CN.md](README.zh-CN.md)
