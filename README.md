# Gewu (格物) — Excel Understanding Agent

[English](README.md) | [简体中文](README.zh-CN.md)

**An open-source agent that reads real-world spreadsheets the way a careful analyst does — including the messy ones nobody admits to making.**

[![License: AGPL v3](https://img.shields.io/badge/License-AGPL_v3-blue.svg)](LICENSE)
![Python 3.12+](https://img.shields.io/badge/Python-3.12%2B-3776AB?logo=python&logoColor=white)
![Platform](https://img.shields.io/badge/Platform-Windows-0078D6?logo=windows&logoColor=white)
![Tauri 2](https://img.shields.io/badge/Desktop-Tauri_2-FFC131?logo=tauri&logoColor=black)

Merged cells. Multi-level headers. Floating meta rows. Annotation lines. A second table hiding at row 23 with restarted serial numbers. Every tool you've tried reads that as garbage. Gewu segments it, builds a header tree, and lets an LLM answer questions over the actual structure — then, if you own the file, writes results back through real Excel COM.

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

Three fictional example files ship in `samples/` (safe to share, no real data).
A local challenge set of 24 additional files (8 structural categories) can be used for extended testing — see `regression.py` for setup.

| Series | What breaks naive parsers |
|---|---|
| T1 结构迷宫 Structure Maze | merged fee categories across row groups |
| T2 间接寻址 Indirect Addressing | dynamic region lookups |
| T3 循环求解 Circular Solve | tax and bonus pool interlocks |
| T4 埋错狩猎 Buried Error Hunt | deliberately planted inconsistencies |
| T5 口径炼狱 Caliber Purgatory | same metric, five definitions |
| T6 野路子结构 Wild Structure | the kitchen-sink ledger |
| T7 碎片拼图 Fragment Assembly | tables split across scattered ranges |
| T8 格式密码 Format Cipher | full-width digits, comma styles, text numbers |

Every golden file is a regression test: if a change alters how any challenge is understood, the suite fails before you ship.

## Quick Start

```bash
git clone https://github.com/TeryR/gewu.git
cd gewu
pip install -r requirements.txt
cp llm.json.example llm.json   # optional — or just set your provider in the app's Settings panel
python cli.py
```

> Windows + local Excel required for write-back (COM). Reading and analysis work anywhere Python runs.

## Key Features

- **Safe write-back** — natural language edit instructions, 8 hard validations, diff preview, auto-backup, one-click undo. The LLM never does arithmetic; it only selects the operation mode and copies the operand.
- **Human-in-the-loop fixes** — wrong block type? Switch it. Wrong frame (too small / split)? Select the block and drag its edges — headers and records are re-detected inside your range. Fixes survive reopens and file takeovers.
- **Bilingual AI output** — switch the UI between 中文/English; AI responses follow the selected language.
- **File tabs** — open multiple workbooks in one window, 0.05s tab switching with fingerprint validation.
- **Business rules** — set value ranges per column (e.g. `quantity 0–100`), violations blocked or flagged, one-click full-data health check.
- **Field glossary** — teach the tool your company's terminology; definitions auto-injected into LLM context.
- **LLM-agnostic** — works with any OpenAI-compatible API (SiliconFlow, DeepSeek, GLM, Qwen), Anthropic-compatible API, or local models (Ollama, no key required). Built-in settings panel with presets and connection testing.
- **Full API** — 31 HTTP endpoints cover every UI feature. `GET /api/docs` returns a self-describing API spec, so external agents can automate the entire workflow.
- **Desktop app** — Tauri 2 native window, drag-and-drop file opening, optional system tray, zero-orphan-process shutdown.

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

## License

AGPL-3.0 — use it, fork it, ship it (open your modifications). Closed-source commercial licensing: open an issue or contact the maintainer.
