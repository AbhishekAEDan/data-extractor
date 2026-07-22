<div align="center">

<pre>
 ____                 _      ___      ____
    | __ ) _ __ ___ _ __ | |_   ( _ )    / ___|___
       |  _ \| '__/ _ \ '_ \| __|  / _ \/\ | |   / _ \ _
        | |_) | | |  __/ | | | |_  | (_>  < | |__| (_) |_|
     |____/|_|  \___|_| |_|\__|  \___/\/  \____\___/
        D A T A   E X T R A C T O R
</pre>

### 📚 Turns ELA lesson `.docx` files into clean CSV spreadsheets — fast, verbatim, no hallucination.

Made with ❤️ by **[AbhishekAEDan](https://github.com/AbhishekAEDan)**

[![Latest Release](https://img.shields.io/github/v/release/AbhishekAEDan/data-extractor?label=latest&color=brightgreen)](https://github.com/AbhishekAEDan/data-extractor/releases/latest)
![Python](https://img.shields.io/badge/python-3.9%2B-blue)
![Platform](https://img.shields.io/badge/platform-Windows-lightgrey)

</div>

---

## ✨ What it does

Feed it a folder of lesson documents. Get back **12 ready-to-import CSV
spreadsheets** — Main Lesson, Mini Practice, Vocabulary Vault, Mistake
Spotter, Unit Covers, Comprehension Texts, Portfolios and more.

- 🔒 **Verbatim by construction** — document text is *sliced, never
  rewritten*. What's in the doc is what lands in the cell.
- 🤖 **AI only as a librarian** — a local model (Ollama/Qwen) or the Gemini
  API is used *only* to classify section headings the parser doesn't
  recognise. It never touches your content.
- 🖱️ **Drag & drop** — drop a file or a whole Term folder onto `run.bat`.
- 🔔 **Self-updating** — the app checks this repo's
  [Releases](https://github.com/AbhishekAEDan/data-extractor/releases) on
  startup and offers new versions automatically.

## 🚀 Quick start

1. **[Grab the latest release](https://github.com/AbhishekAEDan/data-extractor/releases/latest)**
   and unzip it anywhere.
2. Put your `.docx` lessons in `documents/` (subfolders like `Term/Unit`
   are fine) — or drag them straight onto `run.bat`.
3. Double-click **`run.bat`** (or run `python main.py`).
4. Pick **`1) Run extraction`**. ☕ Sip something. CSVs land in `output/`.

> 🐍 Needs Python 3.9+. Missing packages auto-install on first run.

## 🔑 API key (optional)

You only need a key if you pick the **Gemini** engine — the **Ollama/Qwen**
engine is local and free.

| How | Where |
|---|---|
| In-app | menu → `3) Change Gemini API key` (saved to a git-ignored `.env`) |
| Manual | copy `.env.example` → `.env`, paste your key |
| Env var | set `GEMINI_API_KEY` in your environment |

Get a key at <https://aistudio.google.com/apikey>. The real key **never**
lands in this repo.

## 🗂️ What comes out

```
output/
├── ELA Main Lesson.csv
├── ELA Mini Practice.csv
├── ELA Bridge Back & Vocabulary Vault.csv
├── ELA Mistake Spotter.csv
├── ELA Unit Cover (Goals & By End Of Lesson Content).csv
├── ELA Individual Sub Unit My Checklist, Quick Recap, Reflect.csv
├── _ELA Main Comprehension Text.csv
├── Unit Intro Page.csv
├── ELA End Of Unit Portfolio Doc.csv
├── ELA  Overall Portfolio Project.csv
├── ELA Check Your Understanding Term 2.csv
└── wide.csv          ← everything, one row per document (debugging)
```

## 🧰 For tinkerers

| File | Job |
|---|---|
| `main.py` | console UI, menu, config |
| `parser_core.py` | the deterministic .docx parser |
| `writers.py` | shapes parsed rows into the 12 CSVs |
| `judge.py` | AI heading classifier (+ deterministic hard rules) |
| `updater.py` | GitHub release update check |
| `bootstrap.py` / `checks.py` / `logger.py` | auto-install, env checks, run logs |

---

<div align="center">

*No school content ships in this repository — bring your own documents.* 🎒

</div>
