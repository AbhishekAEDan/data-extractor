# Data Extractor

Console app (`main.py`) that parses lesson `.docx` files into .xlsx
spreadsheets matching official templates. Verbatim by construction: text is
sliced from the document, never rewritten. AI (Ollama/Qwen local, or Gemini)
is used only to classify a handful of unrecognised ELA heading labels — it
never touches document content. On the IT path AI is optional and only
locates split points in text the parser already extracted (see "AI use in
the IT path"); with no engine configured, IT output is fully deterministic.
On the current corpus a live qwen3:8b run makes only 2 model calls and
finishes in ~8s, and produces byte-identical output to the ai=None run.

Version: **3.3.1** (bump lives in `updater.py:__version__`, imported by
`main.py` for the banner). Repo: `AbhishekAEDan/data-extractor` on GitHub;
`gh` is authed as AbhishekAEDan. Releases are flat `git archive` zips
(no `repo-sha/` prefix) so the in-app self-updater can find `main.py` at
the root and copy it in place without touching `documents/`, `output/`,
`logs/`, `.env`, or `config.json` (see `updater.py:PROTECTED`).

## Two subjects, one app

- **ELA** — original layout. Parser: `parser_core.py`. Writer: `writers.py`.
  Sections matched by a bold-label whitelist (`LABELS`). Unknown headings go
  through the AI judge (`judge.py`) to be mapped onto canonical columns.
- **IT** — added in v3.0.0. Parser: `parser_it.py`. Writer: `writers_it.py`.
  14 numbered sections matched by NAME regardless of numbering style or
  Word Heading level (docs vary: bold-only, Heading 1/2/3, or Word list
  numbering with no typed digits at all). Fully deterministic — no AI judge.

Subject is picked in `main.py`'s menu (`2) Select subject`: Auto/ELA/IT),
saved in `config.json`. Auto-detect (`parser_it.py:detect_subject`) scores
each doc's section markers against both layouts and picks the winner;
shown in the banner next to Engine, e.g. `IT  (auto-detected)`.

## IT parser specifics (parser_it.py) — the fragile part

Documents in this corpus are NOT uniform. Section headings appear as:
- Word Heading 1/2/3 styles
- Fully-bold plain-style paragraphs, numbered ("1 Learning Objectives") or not
- Word list-numbered paragraphs where the visible "1." is rendered by
  `numbering.xml`, not typed into the text — `python-docx` `.text` doesn't
  include it. Handled via `ListFormatter` (reads `w:numFmt` per abstract
  numbering level, tracks per-numId per-ilvl counters, renders `•` / `1.` /
  `a)` / `i.` markers and re-attaches them to the line during parsing).

Known traps already fixed once — don't reintroduce:
- Treating "list item -> never a heading" broke docs whose section
  headings ARE Word-numbered list items (e.g. File Maintenance.docx). A
  bold/styled list item that names a known section IS a heading.
- A colon-terminated unnumbered bold line ("Steps:", "Correct state:")
  is a sub-label inside the current section, NOT a new section — UNLESS
  it matches no known section name at all shouldn't have blocked real
  section matches either (e.g. "Challenge / Higher-Order Thinking:" with a
  colon must still resolve to that section).
- Inside `Answers & Solutions`, plain-bold restatements of earlier section
  names ("Guided Practice Answers", "5. Guided Practice:", "Section 5:
  Practical Activity") are answer sub-blocks, not new sections — only a
  real Word Heading style may switch `current` away from Answers &
  Solutions.
- A bare "Challenge" heading matches exactly (`n.rstrip(" :") == "challenge"`)
  so content sub-titles like "Quick Break Challenge" don't hijack the
  section.
- `[DIAGRAM: ...]` / `[IMAGE: ...]` marker lines must never be picked as
  the module/lesson title candidate.

Known genuine (not bugs) gaps in the current 20-doc Term 1-3 corpus:
- `Netiquette.docx` has no Learning Objectives / Introduction section —
  the document literally starts at "3. Key Concepts". Fix in the doc.
- `Cyberbullying.docx`'s Common Mistakes & Tips heading has no body text.
- `Saving and Protecting Data...docx` has no top-level Challenge section;
  its challenge answers live inside Answers & Solutions instead.
- `Computer Care.docx` and `Keyboarding.docx` each contain two full copies
  of the lesson back to back (a draft plus a revised copy). Both copies are
  still parsed; the writer drops the draft ONLY when the two copies are
  demonstrably the same content (see "Duplicate-content removal").
  - `Keyboarding.docx`'s two copies are near-identical, so all 14 sections
    dedupe cleanly.
  - `Computer Care.docx`'s "revised" copy is not just reworded, it is
    substantially CUT DOWN: the draft carries material the revised copy
    dropped entirely (the whole 3.1 physical-harm walkthrough with its
    four hazards and `[SCREENSHOT:]` markers, the shut-down/unplugging
    explanation, 3.3 defragmentation + the HDD-vs-SSD note, the 7-step
    Practical Activity with its checklist, all five Guided Practice
    questions with hints, the Knowledge Check MCQ options, and ~79 lines
    of Answers & Solutions). The similarity guard detects that the halves
    differ and therefore KEEPS BOTH copies — 11 of its 14 sections skip
    dedupe and print `! dedupe SKIPPED (halves not similar enough)`.
    That is deliberate: nothing may be deleted when the halves do not
    match. The SOURCE DOCUMENT needs human cleanup to pick one version;
    until then Computer Care's rows carry duplicated text and the dynamic
    sheets are correspondingly wider.

## Formatting + images (IT path, v3.0.0+)

- List markers (bullets, `1.`/`a)`/`i.` counters, nested indent via
  leading tabs) are reconstructed from `numbering.xml` via `ListFormatter`
  and kept in the extracted text — not flattened to bare lines.
- Every IT sheet has a trailing `Image` column. `writers_it.py:split_images`
  pulls `[DIAGRAM:...]` / `[IMAGE:...]` / `[SCREENSHOT:...]` / video marker
  lines out of a section's body. The parser also detects genuinely
  embedded Word pictures (`_para_images`, reads `w:drawing` + `wp:docPr`
  alt text) and inserts an `[IMAGE: ...]` placeholder at their position —
  none of the current 20 docs have real embedded pictures, only text
  markers, but the code path is exercised and ready.
- Matching helpers (`_pa_bucket`, `_MISTAKE_RE`, `_NUMLINE`, etc.) run
  against `_bare(line)` (strips list markers/tabs) so matching logic
  never sees the new markers; the actual cell content keeps them.

## Dynamic column sizing (v3.1.0, uncommitted as of last session)

Three IT sheets used to have FIXED slot counts with overflow silently
merged into the last slot — this quietly lost structure (Plagiarism doc's
9 Key Concepts sub-sections crammed 3.6-3.9 into one "Paragraph 5" cell).
Fixed by computing slot counts from the actual parsed data every run
(two-pass: parse everything, take the max count across all docs, pad
shorter rows with `""`, never truncate or merge):

- **Key Concepts** — `Heading N`/`Paragraph N` pairs, N = max N.N
  sub-block count across docs (was fixed at 5, corpus needs 9).
- **Introduction / Topic Overview** — `Paragraph N` columns, N = max
  paragraph count (was fixed at 5, corpus needs 8).
- **Answer Solution** — four independently-sized categories (Guided
  Practice / Independent Practice / Higher-Order Thinking / Knowledge
  Check), each sized to its own max (was fixed 10/10/5/10; corpus needs
  21/33/22/21).

`Image` column stays last on every sheet regardless of dynamic width.
This means column counts will shift if new documents introduce more
sub-sections/paragraphs/answers than the current max — that's intended
behavior, not a bug, but it means re-running extraction on a bigger
corpus can add columns to sheets that were already delivered. Mention
this to the user if it comes up.

## Tables + XLSX output (uncommitted, this session 2026-07-30)

Both ELA and IT output are .xlsx (openpyxl), same base filenames; only
full_extract.csv stays CSV. `xlsx_out.py` holds the shared writing
machinery (`write_xlsx`, `_clean_cell`, widths, yellow fill, optional
image embedding) imported by both `writers.py` and `writers_it.py`;
`writers.write_csv` is gone. Parser walks body in document order
(`_iter_body`), so Word tables are captured. Per table, two markers at
its position: `[TABLE: r1c1 | r1c2 || r2c1 ...]` (debug only — survives
in full_extract.csv, stripped from ALL xlsx cells; there is NO Table
column, user removed it) and `[TABLE-REQ: TABLE REQUIRED – see
<doc>.docx (Term N)]` (stays in the paragraph flow; writer strips the
wrapper and yellow-fills (FFFF00) any cell containing "TABLE REQUIRED").
Corpus: 12 tables in 8 docs → 12 placeholder occurrences in 11 cells.

Image markers no longer collect into a trailing Image column. Each
marker rides its slot and lands in an image column IMMEDIATELY AFTER the
content column it occurred under: `_C` cell class + `_finalize` in
writers_it.py expand each content column into content + `<col> Image`
(plain `Image` if only one), created ONLY when some row has markers
there. `pad_cells` keeps pairs aligned (`pad` now unused but kept).
Matching helpers (`_pa_bucket`, `_fan`, etc.) guard via `_is_img`/
`_is_treq` so markers are never counted as slots/answers. Embedded
pictures extracted to `<out>/_images/` and anchored into the adjacent
image cell (~200px); corpus has none, path tested with synthetic doc.
`="3.10"` CSV hack replaced by text-format cells. bootstrap/requirements
gained openpyxl + pillow.

Known quirks: table/image markers inside Answers & Solutions ride the
preceding answer (never their own slot) so answer columns don't shift;
Cyberbullying's placeholder is alone in Mistake (section has no body
text); two image-ish lines NOT matched by split_images and left as
paragraph text (deliberate): "[Suggestion: A video showing...]"
(Definitions of IT) and "[Use of a graphic organizer for this list]"
(Cyberbullying) — widen regex only if user asks.

## Question fanning + wide practice sheets (v3.3.0)

- `writers_it.py:split_questions(text, body_re, ai=None)` is the ONE
  implementation of question fanning (`_split_top` + `_fan_q`). Both the
  four practice sheets and `_questions()` (Answer Solution) go through it,
  so question boundaries always agree across sheets. `_NUMLINE_Q` accepts
  bare digits ("1 Explain ...") and dotted "7.1"; a multi-line question
  (scenario + `•` sub-bullets, e.g. 2.9's Q4) stays whole in one cell.
- The four practice sheets are now DYNAMIC WIDE, same two-pass philosophy
  as Key Concepts/Intro/Answers: intro columns unchanged, the old
  questions-blob column replaced by `<name> N` columns, N = that sheet's
  max question count. Guided → `Practice Scenarios N`, Independent →
  `Independent Practice N` (old trailing-space header quirk dropped),
  Knowledge Check → `Questions N`, Challenge → `Question N` with NO intro
  column (the challenge section has no top-paragraph regex; its old
  "Section Content" name was generic). Cells are `_C`, so per-column
  `Image` columns keep working.
- Zero-question safety net inside `split_questions`: body text but 0
  questions → retry `_fan_q_retry` over the FULL text ignoring the top
  split (used when it yields ≥2 items, with empty top) → else, if `ai`
  given, ask it for the verbatim line the first question starts on.
  Non-verbatim reply / NONE / None / raising callable → deterministic
  result stands. `_fan_q_retry` is `_fan_q` plus `_QLABEL_LINE`: a
  question-type label ("Multiple Choice:", "True or False:", "Short
  Answer:", "Fill in the blank(s):", "Matching:") opens an item as well as
  a number. It is used ONLY on this path — widening `_KC_BODY` and friends
  globally would re-split documents that already fan correctly.
  Corpus triggers: `Computer Care.docx` Knowledge Check (post-dedupe; the
  revised copy is label-numbered, not digit-numbered) → retry ACCEPTED,
  4 questions; `Tables and Images.docx` Guided Practice (task-based, no
  numbering and no labels) → retry REJECTED, its text stays whole in the
  Top paragraph cell. Nothing is ever lost either way.
- `Common Mistakes & Tips` is likewise one row per doc with `Mistake N`/
  `Tip N` column pairs (was one row per pair).

## Duplicate-content removal (v3.3.0)

`writers_it.dedupe_rows(rows, ai)` runs in `write_all_it` BEFORE any
writer, on shallow row copies, so every sheet (and full_extract.csv) sees
single copies. Per section, in order of preference:

1. **Re-entry offset (primary).** `parser_it.parse_it_docx` records
   `row["_refilled"] = {section: line_index}` whenever it switches INTO a
   section it had already filled — the signature of a second copy. The
   writer keeps everything from that line onward, verbatim, PROVIDED the
   similarity guard below agrees the two parts really are copies. The
   flag is set at the
   real section switch, after the Answers & Solutions restatement guard,
   so "Guided Practice Answers" sub-blocks never fake a re-entry.
   `_refilled` is in `_PRIVATE_KEYS` — never a sheet column.
2. **Deterministic halves.** Body splits into two near-identical halves
   (normalised via `_bare` + whitespace + case, split point slides ±2,
   ≤10% of lines may differ, each half ≥3 lines or ≥150 chars) → keep the
   second. For docs the parser did not flag.
3. **AI fallback.** Only with `ai`: body >600 chars AND >40% of lines
   repeated (the repeat gate is dropped for `_refilled` sections). Asks
   for the verbatim line the second copy starts on; verbatim-substring
   check, else unchanged.

**Similarity guard (mandatory on paths 1 and 3).** The offset and AI paths
only LOCATE a boundary — they never prove the two parts say the same thing,
so a wrong boundary would delete unique content. Before any removal on
those paths, `_guard`/`_similarity` compare the part about to be dropped
against the part being kept, on `_norm_line`-normalised text: a
`difflib.SequenceMatcher` ratio ≥ 0.50, OR ≥ 60% of dropped lines having a
`get_close_matches` counterpart (cutoff 0.75) in the kept lines. Fail →
the FULL original text is kept and `! dedupe SKIPPED (halves not similar
enough): <doc> / <section> -- ratio X, line share Y` is printed and logged.
Path 2 (deterministic halves) already proves similarity and is exempt.
`SequenceMatcher` MUST be built with `autojunk=False`: the default
heuristic treats frequent characters as junk on inputs over 200 chars and
scores near-identical texts at ~0.02 (this skipped 28/28 removals once).

Every removal prints `! duplicate content removed (<how>): <doc> /
<section> -- N chars`, goes to the run log, and has its dropped text
written VERBATIM to `logs/dedupe_removed_<timestamp>.txt` (one file per
run, header per entry: doc / section / how / chars) so any removal is
auditable and recoverable. Corpus: 17 removals — Keyboarding all 14
sections (ratios 0.66–1.00), Computer Care only 3 (Learning Objectives,
Practical Activity, Suggested Interactive Moments); Computer Care's other
11 skip the guard, see the Computer Care note above. No other document
loses a single character. Dedupe shrinks some dynamic widths — expected,
the doubled copies were inflating the maxima.

## AI use in the IT path (v3.3.0)

`main.py:make_ai(cfg)` builds a `prompt -> reply | None` callable from the
configured engine (ollama, default model **qwen3:8b**, or gemini) and
passes it as `write_all_it(rows, out_dir, ai=...)`. It returns None when no
engine is usable, so the default path stays fully deterministic. The model
is only ever asked to LOCATE a split point (Mistake/Tip boundary, first
question line, second-copy start); replies are accepted only when they are
verbatim substrings of the parsed text, so model words can never become
cell content. Extraction stays verbatim by construction.

The first real run with Ollama LOOKED like a hang: qwen3 is a thinking
model, every call could burn the whole 30s `judge.AI_DECIDE_TIMEOUT`, the
same section was asked about up to five times (four practice writers plus
`_questions`), and nothing was printed while it blocked. Three fixes, all
in place — do not undo them:

- `judge._ai_ollama` sends `"think": false` plus
  `options {"temperature": 0, "num_predict": 200}`. It retries ONCE without
  `think` on an `HTTPError`, for older Ollama builds that reject the field.
  Timeout stays 30s. Locate-a-line replies now come back in ~3s.
- `writers_it._ai_ask(ai, prompt, ctx, purpose)` is the single funnel for
  every AI call (`_ai_locate`, `_ai_split_tip`). It memoises on the prompt
  in the module-level `_AI_CACHE` (cleared per run by `reset_ai_cache()`
  from `write_all_it`; `ai_stats()` exposes call/hit counts for tests), so
  a section is never sent twice.
- It also prints `  [AI] deciding: <doc> / <section> / <purpose>...` before
  blocking and ` ok` / ` no answer (using deterministic result)` /
  ` failed (...)` after, plus `  [AI] cached: ...` on a hit. `ctx` is
  threaded from the call sites (`split_questions(..., ctx=)`,
  `_write_practice`, `_questions`, `write_mistakes`, `dedupe_rows`).

`_AI_QSPLIT_PROMPT` used to have no `{text}` placeholder, so `_ai_locate`
sent the model instructions with no document text — that safety net never
actually worked. Fixed (it now ends with `Text:\n{text}` like
`_AI_DUP_PROMPT`). Live on `Tables and Images.docx` Guided Practice the
model answers either `NONE` or the "Task 1:" line; both are safe (a
non-verbatim/NONE reply leaves the deterministic result standing).

## Working conventions for this project

- **User's style is caveman-terse.** Keep prose replies short/fragment-y;
  code, commits, and technical content stay normal/full quality — the
  terseness is about conversational filler, not about explanations that
  need precision.
- Verify claims by actually running the parser/writer against the real
  `documents/` corpus (20 `.docx` files across Term 1/2/3) before reporting
  a fix works — this corpus has repeatedly surfaced edge cases that
  synthetic tests miss.
- `documents/*.docx` and `output/*` are gitignored (school content must
  never be public) — only code/docs get committed. Double-check `git
  status`/`git add -An` before committing to make sure no `.docx` leaked in.
- Never add a `Co-Authored-By` trailer unless asked — the user explicitly
  said not to for the v3.0.0 commit; treat that as standing preference
  unless told otherwise.
- When delegating a Data Extractor implementation task to a subagent
  worktree, the worktree can silently diverge from the main checkout
  (stale copies of files not in scope for that task). After merging the
  agent's target file back, re-run the verification yourself against the
  real main-checkout files, not inside the worktree.
- `main.py`'s menu numbering has shifted before (subject selector inserted
  as item 2, pushing Gemini key/checks/output/documents down by one) — if
  referencing menu numbers in README or elsewhere, re-check against
  `main.py`'s actual `print()` calls rather than assuming.
