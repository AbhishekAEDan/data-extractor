<div align="center">

<pre>
 ____                 _      ___      ____
   | __ ) _ __ ___ _ __ | |_   ( _ )    / ___|___
      |  _ \| '__/ _ \ '_ \| __|  / _ \/\ | |   / _ \ _
       | |_) | | |  __/ | | | |_  | (_>  < | |__| (_) |_|
    |____/|_|  \___|_| |_|\__|  \___/\/  \____\___/
       D A T A   E X T R A C T O R
</pre>

### 📚 Turns ELA **and IT** lesson `.docx` files into clean CSV spreadsheets — fast, verbatim, no hallucination.

Made with ❤️ by **[AbhishekAEDan](https://github.com/AbhishekAEDan)**

[![Latest Release](https://img.shields.io/github/v/release/AbhishekAEDan/data-extractor?label=latest&color=brightgreen)](https://github.com/AbhishekAEDan/data-extractor/releases/latest)
![Python](https://img.shields.io/badge/python-3.9%2B-blue)
![Platform](https://img.shields.io/badge/platform-Windows-lightgrey)

</div>

---

## ✨ What it does

Feed it a folder of lesson documents. Get back a set of ready-to-import CSV
spreadsheets that mirror the official templates one-to-one.

- 🎓 **Two subjects, one app** — pick **ELA** or **IT**, or let the app
  work it out for itself from the documents you fed it. The detected
  subject is shown right in the header next to the engine.
- 🔒 **Verbatim by construction** — document text is *sliced, never
  rewritten*. What's in the doc is what lands in the cell.
- ✏️ **Formatting kept** — bullets, numbered lists (`1.` `a)` `i.`) and
  nested indentation are rebuilt from the document's own list definitions,
  not thrown away.
- 🖼️ **Images tracked** — every sheet has an `Image` column; diagram,
  screenshot and video markers, plus genuinely embedded pictures, are
  pulled out of the section they belong to.
- 🤖 **AI only as a librarian** — a local model (Ollama/Qwen) or the Gemini
  API is used *only* to classify section headings the ELA parser doesn't
  recognise. It never touches your content. **IT runs are fully
  deterministic and need no AI at all.**
- 🖱️ **Drag & drop** — drop a file, a whole Term folder, or a **.zip full
  of lessons** onto `run.bat` (zips in `documents/` work too — they're
  extracted automatically).
- 🔄 **Live folder watch** — add or remove documents *while a run is going*;
  the run rescans at the end and reconciles automatically.
- 🔔 **Self-updating** — the app checks this repo's
  [Releases](https://github.com/AbhishekAEDan/data-extractor/releases) on
  startup and offers new versions automatically.

## 🚀 Quick start

1. **[Grab the latest release](https://github.com/AbhishekAEDan/data-extractor/releases/latest)**
   and unzip it anywhere.
2. Put your `.docx` lessons in `documents/` (subfolders like `Term/Unit`
   are fine) — or drag them straight onto `run.bat`.
3. Double-click **`run.bat`** (or run `python main.py`).
4. Check the **Subject** line in the header. It auto-detects; override it
   any time with **`2) Select subject`**.
5. Pick **`1) Run extraction`**. ☕ Sip something. CSVs land in `output/`.

> 🐍 Needs Python 3.9+. Missing packages auto-install on first run.

## 🔑 API key (optional)

You only need a key if you pick the **Gemini** engine — the **Ollama/Qwen**
engine is local and free, and **IT extraction needs no engine at all**.

| How | Where |
|---|---|
| In-app | menu → `4) Change Gemini API key` (saved to a git-ignored `.env`) |
| Manual | copy `.env.example` → `.env`, paste your key |
| Env var | set `GEMINI_API_KEY` in your environment |

Get a key at <https://aistudio.google.com/apikey>. The real key **never**
lands in this repo.

## 🗂️ What comes out

Which sheets you get depends on the subject.

<table>
<tr><th>ELA</th><th>IT</th></tr>
<tr valign="top"><td>

```
Unit Cover Page.csv
Vocab Vault.csv
Mistake Spotter.csv
Mini Practice.csv
Check Your Understanding.csv
In Sub Unit Portfolio Proj.csv
Sub Unit Recap.csv
Unit Introduction.csv
End of Unit Recap.csv
End of Unit Portfolio.csv
full_extract.csv
```

</td><td>

```
Unit cover page & Learning Objective.csv
Introduction Topic Overview.csv
Key Concepts (Main Content).csv
Practical Activity (Hands-On).csv
Guided Practice (With Support).csv
Independent Practice.csv
Challenge Higher-Order Thinking.csv
Knowledge Check (Quick Assessment).csv
Common Mistakes & Tips.csv
Real-World Application.csv
Summary Key Takeaways.csv
Answer Solution.csv
Suggested Interactive Moments.csv
full_extract.csv
```

</td></tr>
</table>

`full_extract.csv` is the debug sheet — every parsed field, one row per
document.

Layouts mirror the official spreadsheet templates one-to-one (on the ELA
side the Main Lesson sheet is deliberately left to a human).

## 🧰 For tinkerers

| File | Job |
|---|---|
| `main.py` | console UI, menu, config, subject routing |
| `parser_core.py` | the deterministic ELA .docx parser |
| `writers.py` | shapes parsed ELA rows into the output CSVs |
| `parser_it.py` | the deterministic IT .docx parser (+ subject detection) |
| `writers_it.py` | shapes parsed IT rows into the output CSVs |
| `judge.py` | AI heading classifier (+ deterministic hard rules) |
| `updater.py` | GitHub release update check |
| `bootstrap.py` / `checks.py` / `logger.py` | auto-install, env checks, run logs |

---

<div align="center">

*No school content ships in this repository — bring your own documents.* 🎒

</div>
