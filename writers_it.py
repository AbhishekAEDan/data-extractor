"""
writers_it.py

Writes parsed IT lesson rows to formatted .xlsx workbooks. Layouts mirror the
Form 1 IT spreadsheet templates plus the "AI Sample Sheets/content
sheets.xlsx" examples (headers copied exactly, including their spacing/typo
quirks -- "Acitvity Intro ", "Module Title 1", "Think about it " etc.):

    Unit cover page & Learning Objective.xlsx
    Introduction Topic Overview.xlsx
    Key Concepts (Main Content).xlsx
    Practical Activity (Hands-On).xlsx
    Guided Practice (With Support).xlsx
    Independent Practice.xlsx
    Challenge Higher-Order Thinking.xlsx
    Knowledge Check (Quick Assessment).xlsx
    Common Mistakes & Tips.xlsx        (one row per Mistake/Tip pair)
    Real-World Application.xlsx
    Summary Key Takeaways.xlsx
    Answer Solution.xlsx               (answers fanned into numbered slots)
    Suggested Interactive Moments.xlsx (one row per suggested moment)
    full_extract.csv                   (debug: every parsed field, stays CSV)

Both subject paths are xlsx now; the formatting machinery (`write_xlsx` and
friends) lives in `xlsx_out.py` and is shared with `writers.py` (ELA).
full_extract.csv stays CSV on both paths.

All values are verbatim slices from the parser -- no rewriting happens here;
the only text the writer removes is its own internal `[TABLE-REQ: ...]` /
`[TABLE: ...]` wrappers.

Author: AbhishekAEDan
"""
__author__ = "AbhishekAEDan"

import csv
import difflib
import os
import re
from datetime import datetime

from logger import LOGS_DIR, log
from parser_it import IT_ORDER
from writers import unit_sort_key, join_nonempty
from xlsx_out import (_TBL_LINE, _TREQ_LINE, _index_images,  # noqa: F401
                      write_xlsx)


# ---------- AI call plumbing: cache + console feedback ----------
#
# The model is slow (a thinking model can spend the whole 30s timeout on one
# call) and the SAME section is handed to `split_questions` up to five times
# per run (four practice sheets plus `_questions`). Without a cache and
# without console output the run looks like a hang. Both are fixed here, in
# the one place every AI call goes through.

_AI_CACHE = {}
_AI_STATS = {"calls": 0, "hits": 0}


def reset_ai_cache():
    """Fresh cache/counters. Called at the start of every `write_all_it` run;
    also handy for tests that want to count real calls."""
    _AI_CACHE.clear()
    _AI_STATS["calls"] = 0
    _AI_STATS["hits"] = 0


def ai_stats():
    return dict(_AI_STATS)


def _ai_ask(ai, prompt, ctx="", purpose=""):
    """Run one AI query with memoisation and console feedback.

    Returns the reply text or None. Identical prompts are answered from the
    module-level cache, so a section is never sent to the model twice in a
    run. Every real call announces itself before it blocks and reports the
    outcome after, so a slow model reads as slow rather than as a hang."""
    label = " / ".join(x for x in (ctx, purpose) if x) or "?"
    if prompt in _AI_CACHE:
        _AI_STATS["hits"] += 1
        print(f"  [AI] cached: {label}")
        return _AI_CACHE[prompt]
    print(f"  [AI] deciding: {label}...", end="", flush=True)
    _AI_STATS["calls"] += 1
    try:
        reply = ai(prompt)
    except Exception as e:
        print(" failed (using deterministic result)")
        log(f"AI call failed for {label}: {e}")
        _AI_CACHE[prompt] = None
        return None
    if reply:
        print(" ok")
    else:
        print(" no answer (using deterministic result)")
    _AI_CACHE[prompt] = reply
    return reply


def _sorted_rows(rows):
    out = [r for r in rows if r.get("_doc_type") == "it_lesson"]
    return sorted(out, key=lambda r: unit_sort_key(r.get("Module Number", "")))


# internal, non-textual row keys that must never become a sheet column
_PRIVATE_KEYS = {"_image_files", "_refilled"}


def order_cols(rows):
    cols = [c for c in IT_ORDER if any(c in r for r in rows)]
    for r in rows:
        for k in r:
            if k not in cols and k not in _PRIVATE_KEYS:
                cols.append(k)
    return cols


def write_full_extract(rows, out_dir):
    cols = order_cols(rows)
    path = os.path.join(out_dir, "full_extract.csv")
    with open(path, "w", newline="", encoding="utf-8-sig") as fh:
        w = csv.DictWriter(fh, fieldnames=cols)
        w.writeheader()
        for r in rows:
            w.writerow({k: r.get(k, "") for k in cols})
    return path


# image / diagram / video marker lines authored in the documents, plus the
# "[IMAGE...]" placeholders the parser inserts for embedded pictures
_IMG_LINE = re.compile(
    r'^\s*\[\s*(?:diagram|image|screenshot|photo|picture|video|sim\b|'
    r'animation|gif|use of a video)', re.I)


def split_images(text):
    """Split a section's verbatim text into (body, images). Image lines are
    the bracketed marker lines; a marker whose ']' falls on a later line is
    consumed up to that line."""
    body, imgs = [], []
    in_img = False
    for ln in (text or "").split("\n"):
        s = ln.strip()
        if in_img:
            imgs[-1] += "\n" + s
            if "]" in s:
                in_img = False
            continue
        if _IMG_LINE.match(s):
            imgs.append(s)
            in_img = "]" not in s
        else:
            body.append(ln)
    return "\n".join(body).strip(), "\n".join(imgs).strip()


def split_tables(text):
    """Split a section's verbatim text into (body, tables). Table marker
    lines are pulled out and rendered for a spreadsheet cell: one row per
    line, cells joined with ' | '. Several tables in one section are
    separated by a blank line.

    NOTE: no sheet writes the table text any more -- the sheets carry only
    the yellow "TABLE REQUIRED" placeholder. This is kept because the parser
    still emits the [TABLE: ...] marker (it reaches full_extract.csv for
    debugging) and `_body` needs to drop it from every sheet cell."""
    body, tables = [], []
    for ln in (text or "").split("\n"):
        m = _TBL_LINE.match(ln)
        if m:
            tables.append("\n".join(r.strip() for r in m.group(1).split(" || ")))
        else:
            body.append(ln)
    return "\n".join(body).strip(), "\n\n".join(tables).strip()


def _is_img(line):
    """True for an image/diagram/screenshot/video marker line. Slot-splitting
    helpers use this so a marker is never counted as a paragraph, an answer,
    a mistake or a moment of its own -- it rides with the slot it sat in and
    is moved to that slot's adjacent image column at write time."""
    return bool(_IMG_LINE.match((line or "").strip()))


def _body(r, section):
    """One section's verbatim text with the [TABLE: ...] markers dropped.

    Image markers and the [TABLE-REQ: ...] placeholder deliberately STAY in
    the flow: the placeholder belongs in the paragraph cell, and each image
    marker has to survive slot assignment so it can be routed to the image
    column sitting immediately after its own slot."""
    return split_tables(r.get(section, ""))[0].strip()


class _C:
    """A content cell that owns the image markers found inside it.

    `text` is the cell's content with marker lines removed; `img` is the
    marker lines, newline-joined. `_finalize` turns each `_C` into its
    content column plus an image column placed immediately after it."""
    __slots__ = ("text", "img")

    def __init__(self, text):
        self.text, self.img = split_images(text or "")

    def __bool__(self):
        return bool(self.text or self.img)


def _finalize(path, header, rows):
    """Write a sheet whose row entries are plain strings or `_C` cells.

    Every `_C` position expands to its content column followed by an image
    column -- but only when at least one row actually carries markers there,
    the same dynamic philosophy the slot columns already use. A sheet that
    ends up with exactly one image column names it plain "Image"; with
    several, each is named "<content column> Image" so it is obvious which
    paragraph it belongs to. There is no trailing catch-all Image column."""
    keep = {i for i in range(len(header))
            if any(isinstance(r[i], _C) and r[i].img for r in rows)}
    single = len(keep) == 1
    out_h, out_rows = [], [[] for _ in rows]
    for i, name in enumerate(header):
        out_h.append(name)
        for j, r in enumerate(rows):
            v = r[i]
            out_rows[j].append(v.text if isinstance(v, _C) else v)
        if i in keep:
            out_h.append("Image" if single else f"{str(name).strip()} Image")
            for j, r in enumerate(rows):
                v = r[i]
                out_rows[j].append(v.img if isinstance(v, _C) else "")
    write_xlsx(path, out_h, out_rows)
    return path


def _is_treq(line):
    """True for the parser's table placeholder line. Matching helpers use it
    to make sure the placeholder is never mistaken for a heading, a mistake,
    a tip or an answer -- it is plain content that rides where it sat."""
    return bool(_TREQ_LINE.match((line or "").strip()))


def _bare(s):
    """Copy of a line with list markers/indent removed -- for MATCHING only;
    cell content keeps the markers."""
    return s.lstrip("\t •·●▪-–—")


def _mod(r):
    return r.get("Module Number", "")


def _title(r):
    return r.get("Module Title", "")


# ---------- duplicate-content removal ----------
#
# Some source documents (Computer Care.docx is the known case) contain the
# whole lesson twice back to back: an early draft followed by a revised copy.
# Both get captured by the parser, so every sheet shows doubled text. The
# clean-up runs ONCE per section, on the parsed row, before any writer sees
# it -- so all sheets benefit and none of them has to know about it.

# the parsed section keys a lesson body can live in (everything between the
# first and last content section of IT_ORDER); metadata keys are skipped
_SECTION_KEYS = IT_ORDER[IT_ORDER.index("Learning Objectives"):
                         IT_ORDER.index("Developer Notes") + 1]

_WS_RE = re.compile(r'\s+')

# a half must be substantial before it may be treated as a duplicated copy --
# short legitimate repeats (a repeated instruction line, "True or False:" ...)
# must never be deduped
_DUP_MIN_LINES = 3
_DUP_MIN_CHARS = 150
_DUP_SLIDE = 2          # how far the split point may slide either way
_DUP_TOL = 0.10         # share of lines the two halves may differ in

_AI_DUP_MIN_CHARS = 600
_AI_DUP_MIN_REPEAT = 0.40

_AI_DUP_PROMPT = """The text below was extracted from a school lesson \
document. Some of these documents accidentally contain the same content twice: \
an early draft immediately followed by a revised copy of the same material.

Decide whether that has happened here. If it has, reply with ONLY the exact \
line of the text on which the SECOND copy begins, copied character for \
character. No quotes, no explanation, no commentary. If the text is not the \
same content repeated twice, reply NONE.

Text:
{text}"""


def _norm_line(s):
    """Line reduced for COMPARISON only: list markers/indent stripped,
    whitespace collapsed, case folded. Kept text is always verbatim."""
    return _WS_RE.sub(" ", _bare(s or "").strip()).lower()


def _halves_match(a, b):
    """True when line lists `a` and `b` are near-identical copies."""
    if not a or not b:
        return False
    longest = max(len(a), len(b))
    diff = abs(len(a) - len(b))
    for x, y in zip(a, b):
        if _norm_line(x) != _norm_line(y):
            diff += 1
    return diff <= _DUP_TOL * longest


def _dup_half_ok(lines):
    return (len(lines) >= _DUP_MIN_LINES
            or len("\n".join(lines)) >= _DUP_MIN_CHARS)


def _dedupe_deterministic(text):
    """The second of two near-identical halves, verbatim, or None."""
    lines = [ln for ln in text.split("\n") if ln.strip()]
    n = len(lines)
    if n < 2 * _DUP_MIN_LINES and len(text) < 2 * _DUP_MIN_CHARS:
        return None
    mid = n // 2
    for s in sorted(range(max(1, mid - _DUP_SLIDE), min(n, mid + _DUP_SLIDE + 1)),
                    key=lambda x: abs(x - mid)):
        a, b = lines[:s], lines[s:]
        if not (_dup_half_ok(a) and _dup_half_ok(b)):
            continue
        if _halves_match(a, b):
            return "\n".join(b).strip()
    return None


def _looks_repeated(text):
    """True when enough of the body's lines occur two or more times for the
    model to be worth asking. Cheap guard: the AI call is skipped otherwise."""
    if len(text) <= _AI_DUP_MIN_CHARS:
        return False
    norms = [_norm_line(ln) for ln in text.split("\n") if ln.strip()]
    if not norms:
        return False
    counts = {}
    for x in norms:
        counts[x] = counts.get(x, 0) + 1
    repeated = sum(1 for x in norms if counts[x] > 1)
    return repeated > _AI_DUP_MIN_REPEAT * len(norms)


def _ai_locate(text, prompt, ai, ctx="", purpose=""):
    """Ask `ai` to point at a line of `text`; return its index in `text`, or
    None. The reply is accepted ONLY when it is a verbatim substring, so the
    model can locate a split point but can never introduce content."""
    reply = _ai_ask(ai, prompt.format(text=text), ctx, purpose)
    if not reply:
        return None
    reply = reply.strip().strip('"').strip()
    if not reply or reply.upper().startswith("NONE"):
        return None
    idx = text.find(reply)
    return idx if idx > 0 else None


def _dedupe_offset(text, offset):
    """The section text from the parser's recorded re-entry line onward.

    `offset` is the line index at which the parser re-entered a section it had
    already filled -- i.e. exactly where the second copy of the lesson starts.
    No similarity test is needed, so a REWRITTEN second copy (Computer Care's
    revised draft) is handled just as well as an identical one."""
    if not offset:
        return None
    lines = text.split("\n")
    if not 0 < offset < len(lines):
        return None
    return "\n".join(lines[offset:]).strip() or None


# ---------- similarity guard ----------
#
# The offset and AI paths locate a split point WITHOUT proving that the two
# parts say the same thing. If the locator is wrong, the "duplicate" being
# dropped is unique content. Before any removal on those paths the dropped
# part must actually look like a copy of the kept part.

_SIM_RATIO = 0.50        # whole-text SequenceMatcher ratio on normalised text
_SIM_LINE_CUTOFF = 0.75  # per-line close-match cutoff
_SIM_LINE_SHARE = 0.60   # share of dropped lines needing a counterpart


def _norm_lines(text):
    return [n for n in (_norm_line(ln) for ln in (text or "").split("\n")) if n]


def _similarity(dropped, kept):
    """(ratio, line_share) for a candidate removal, both on normalised text."""
    d_lines, k_lines = _norm_lines(dropped), _norm_lines(kept)
    if not d_lines or not k_lines:
        return 0.0, 0.0
    # autojunk MUST stay off: on character sequences longer than 200 items
    # difflib's default heuristic treats every frequent character (i.e. most
    # of the alphabet) as junk and reports ~0.0 for texts that are in fact
    # near-identical.
    ratio = difflib.SequenceMatcher(
        None, " ".join(d_lines), " ".join(k_lines), autojunk=False).ratio()
    matched = sum(1 for ln in d_lines
                  if difflib.get_close_matches(ln, k_lines, n=1,
                                               cutoff=_SIM_LINE_CUTOFF))
    return ratio, matched / len(d_lines)


def _similar_enough(dropped, kept):
    ratio, share = _similarity(dropped, kept)
    return (ratio >= _SIM_RATIO or share >= _SIM_LINE_SHARE), ratio, share


def _guard(dropped, kept, how, ctx):
    """True when `dropped` may be removed. Otherwise prints/logs the skip."""
    ok, ratio, share = _similar_enough(dropped, kept)
    if ok:
        log(f"dedupe similarity ok ({how}): {ctx} -- "
            f"ratio {ratio:.2f}, line share {share:.2f}")
        return True
    msg = (f"dedupe SKIPPED (halves not similar enough): {ctx} "
           f"-- ratio {ratio:.2f}, line share {share:.2f} ({how})")
    print(f"  ! {msg}")
    log(msg)
    return False


def _dedupe_body(text, ai=None, offset=None, flagged=False, ctx=""):
    """(cleaned_text, how, dropped_text) for one section body. `how` is None
    when the body was left untouched. Nothing is ever rewritten -- the result
    is always a verbatim tail slice of the input.

    `offset`/`flagged` come from the parser's `_refilled` record: `flagged`
    marks a section the parser re-entered (a second copy of the lesson),
    `offset` is the line the re-entry began on. The offset path is preferred
    over every heuristic below; the AI is a fallback for a flagged section
    whose offset is unusable, and there the repeated-lines gate is dropped
    because the copies may be fully rewritten.

    Both of those paths only LOCATE a boundary, so the part about to be
    dropped is checked against the part being kept first (`_guard`); the
    deterministic-halves path already proved the two halves match."""
    if not text or not text.strip():
        return text, None, ""
    out = _dedupe_offset(text, offset)
    if out is not None:
        dropped = "\n".join(text.split("\n")[:offset]).strip()
        if _guard(dropped, out, "re-entry offset", ctx):
            return out, "re-entry offset", dropped
        return text, None, ""
    det = _dedupe_deterministic(text)
    if det is not None:
        return det, "deterministic", text[:len(text) - len(det)].strip()
    if ai and (_looks_repeated(text)
               or (flagged and len(text) > _AI_DUP_MIN_CHARS)):
        idx = _ai_locate(text, _AI_DUP_PROMPT, ai, ctx, "duplicate copy start")
        if idx:
            kept, dropped = text[idx:].strip(), text[:idx].strip()
            if _guard(dropped, kept, "ai", ctx):
                return kept, "ai", dropped
    return text, None, ""


# ---------- removal audit trail ----------

def _removal_log(state):
    """Lazily open this run's removal file and return its handle (or None)."""
    if state.get("fh") is not None or state.get("failed"):
        return state.get("fh")
    try:
        os.makedirs(LOGS_DIR, exist_ok=True)
        stamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
        path = os.path.join(LOGS_DIR, f"dedupe_removed_{stamp}.txt")
        fh = open(path, "w", encoding="utf-8")
        fh.write(f"=== Data Extractor -- duplicate content removed, "
                 f"{stamp} ===\n")
        fh.write("Verbatim text dropped by writers_it.dedupe_rows.\n")
        state["fh"] = fh
        state["path"] = path
    except Exception as e:      # auditing must never break a run
        log(f"could not open dedupe removal log: {e}")
        state["failed"] = True
    return state.get("fh")


def dedupe_rows(rows, ai=None):
    """Copies of `rows` with duplicated lesson copies removed, per section.

    Every removal is printed once, written to the run log (document, section,
    characters removed), and its dropped text is written verbatim to
    logs/dedupe_removed_<timestamp>.txt -- one file per run -- so nothing is
    lost silently and any removal can be recovered."""
    out = []
    state = {}
    for r in rows:
        new = dict(r)
        refilled = r.get("_refilled") or {}
        doc = r.get("_file", "?")
        for k in _SECTION_KEYS:
            v = new.get(k)
            if not isinstance(v, str) or not v.strip():
                continue
            cleaned, how, dropped = _dedupe_body(
                v, ai, refilled.get(k), k in refilled, ctx=f"{doc} / {k}")
            if how:
                new[k] = cleaned
                removed = len(v) - len(cleaned)
                msg = (f"duplicate content removed ({how}): "
                       f"{doc} / {k} -- {removed} chars "
                       f"({len(v)} -> {len(cleaned)})")
                print(f"  ! {msg}")
                log(msg)
                fh = _removal_log(state)
                if fh:
                    fh.write(f"\n{'-' * 70}\n"
                             f"doc     : {doc}\n"
                             f"section : {k}\n"
                             f"how     : {how}\n"
                             f"chars   : {removed}\n"
                             f"{'-' * 70}\n{dropped}\n")
                    fh.flush()
        out.append(new)
    fh = state.get("fh")
    if fh:
        fh.close()
        print(f"  Removed duplicate text saved to: {state['path']}")
        log(f"dedupe removals written to {state['path']}")
    return out


# ---------- simple one-block sheets ----------

def write_cover(rows, out_dir):
    header = ["Module Number", "Module Title 1", "Module Descriptive Title",
              "Learning Objectives"]
    data = []
    for r in _sorted_rows(rows):
        obj = _C(_body(r, "Learning Objectives"))
        if obj:
            data.append([_mod(r), _title(r),
                         r.get("Module Descriptive Title", ""), obj])
    return _finalize(os.path.join(
        out_dir, "Unit cover page & Learning Objective.xlsx"), header, data)


def pad(lst, n):
    """Return lst padded with empty strings to length n."""
    return list(lst) + [""] * (n - len(lst))


def pad_cells(lst, n):
    """pad() for slots holding `_C` cells -- pads with empty ones so the
    content/image column pair stays aligned on short rows."""
    return list(lst) + [_C("") for _ in range(n - len(lst))]


# a list item line: nested-indent tabs from ListFormatter, a bullet glyph,
# or a "1." / "a)" / "iv." style counter rendered by ListFormatter. Such a
# line continues the paragraph above it instead of opening a new slot.
_LIST_LINE = re.compile(
    r'^(?:\t+'                      # nested list indent
    r'|[ \t]*(?:[•▪‣◦·⁃-]'   # bullet glyphs / dash
    r'|\d{1,2}[.)]'                 # 1.  2)
    r'|[a-z][.)]'                   # a)  b.
    r'|[ivxlc]{1,5}[.)]'            # i.  iv)
    r')\s)')


def write_intro(rows, out_dir):
    """Introduction paragraphs fan into one slot each. The number of
    Paragraph columns is the maximum paragraph count across all documents;
    shorter rows pad with ''. An image marker is not a paragraph -- it joins
    the paragraph above it and surfaces in that slot's image column. A
    bullet / numbered-list / indented line is not a paragraph either: list
    items belong to the paragraph that introduced them."""
    parsed = []
    for r in _sorted_rows(rows):
        text = _body(r, "Introduction / Topic Overview")
        if not text:
            continue
        paras = []                      # each entry = [lines of one paragraph]
        for ln in text.split("\n"):
            if not ln.strip():
                continue
            if paras and (_is_img(ln) or _LIST_LINE.match(ln)):
                paras[-1].append(ln)
            else:
                paras.append([ln])
        parsed.append((r, [_C("\n".join(p)) for p in paras]))
    n = max([len(p) for _, p in parsed] + [1])
    header = (["Module Number", "Module Title 1"]
              + [f"Paragraph {i}" for i in range(1, n + 1)])
    data = [[_mod(r), _title(r)] + pad_cells(paras, n)
            for r, paras in parsed]
    return _finalize(os.path.join(
        out_dir, "Introduction Topic Overview.xlsx"), header, data)


_SUBHEAD_RE = re.compile(r'^\d{1,2}\.\d{1,2}\b')


def write_key_concepts(rows, out_dir):
    """Intro text before the first N.N sub-heading goes to 'Key concepts
    intro'; each N.N block splits into its Heading slot (the '3.3 Title'
    line) and its Paragraph slot (the text that follows the heading). The
    number of Heading/Paragraph pairs is the maximum block count across all
    documents; shorter rows pad with ''.

    An image marker inside an N.N block belongs to that block's paragraph, so
    it lands in the image column immediately after that Paragraph N."""
    parsed = []
    for r in _sorted_rows(rows):
        text = _body(r, "Key Concepts")
        if not text:
            continue
        intro, blocks = [], []          # block = [heading, [body lines]]
        for ln in text.split("\n"):
            # a marker line is never a sub-heading; _SUBHEAD_RE cannot match
            # one, so it simply accumulates into the block it sat in
            if _SUBHEAD_RE.match(_bare(ln).strip()):
                blocks.append([ln.strip(), []])
            elif blocks:
                blocks[-1][1].append(ln)
            else:
                intro.append(ln)
        parsed.append((r, _C("\n".join(intro)),
                       [(h, _C("\n".join(b))) for h, b in blocks]))
    n = max([len(b) for _, _, b in parsed] + [1])
    header = ["Lesson No.", "Key concepts intro"]
    for i in range(1, n + 1):
        header += [f"Heading {i}", f"Paragraph {i}"]
    data = []
    for r, intro, blocks in parsed:
        cells = []
        for h, b in blocks:
            cells += [h, b]
        # pad in Heading/Paragraph pairs so the Paragraph slots keep holding
        # `_C` cells and their image columns stay aligned
        while len(cells) < n * 2:
            cells += ["", _C("")]
        data.append([_mod(r), intro] + cells)
    return _finalize(os.path.join(
        out_dir, "Key Concepts (Main Content).xlsx"), header, data)


# ---------- practical activity ----------

# bold sub-labels inside the Practical Activity section -> output column
_PA_BUCKETS = [
    (re.compile(r'^materials\b', re.I),                 "steps"),
    (re.compile(r'^instructions?\b', re.I),             "steps"),
    (re.compile(r'^(guide|steps?)\b\s*[:/]?', re.I),    "steps"),
    (re.compile(r'^step\s*\d', re.I),                   "steps"),
    (re.compile(r'^success criteria\b', re.I),          "success"),
    (re.compile(r'^step\s*\d+\s*[:.]?\s*review\b', re.I), "success"),
    (re.compile(r'^discussion\b', re.I),                "discussion"),
    (re.compile(r'^troubleshooting\b', re.I),           "trouble"),
    (re.compile(r'^.{0,20}rubric\b', re.I),             "rubric"),
]

_PA_REVIEW = re.compile(r'^step\s*\d+\s*[:.]?\s*review\b', re.I)


def _pa_bucket(line):
    # a marker line is content, never a bucket label -- guarded explicitly
    # because the loose "^.{0,20}rubric" pattern could otherwise be tripped
    # by something like "[SCREENSHOT: A rubric ...]"
    if _is_img(line) or _is_treq(line):
        return None
    if _PA_REVIEW.match(line):
        return "success"
    for rx, b in _PA_BUCKETS:
        if rx.match(line):
            return b
    return None


def write_practical(rows, out_dir):
    """Each bold sub-label starts a bucket; every bucket gets its own image
    column immediately after it when any document has markers there."""
    header = ["Module Number", "Module Title 1", "Practical Title",
              "Acitvity Intro ", "Guide/ Steps", "Success Criteria",
              "Discussion", "Troubleshooting", "Rubric "]
    data = []
    for r in _sorted_rows(rows):
        text = _body(r, "Practical Activity")
        if not text:
            continue
        buckets = {"intro": [], "steps": [], "success": [],
                   "discussion": [], "trouble": [], "rubric": []}
        title = r.get("Practical Title", "")
        cur = "intro"
        for ln in text.split("\n"):
            s = ln.strip()
            if not s:
                continue
            # "Task: ..." lead names the activity when the heading didn't
            if not title and re.match(r'^task\s*[:—–-]', _bare(s), re.I):
                title = re.split(r'[:—–-]', s, 1)[1].strip().split(". ")[0]
            # a marker never opens a bucket -- _pa_bucket cannot match one, so
            # it stays with the bucket it appeared in
            b = _pa_bucket(_bare(s))
            if b:
                cur = b
            buckets[cur].append(s)
        data.append([_mod(r), _title(r), title]
                    + [_C("\n".join(buckets[k])) for k in
                       ("intro", "steps", "success", "discussion",
                        "trouble", "rubric")])
    return _finalize(os.path.join(
        out_dir, "Practical Activity (Hands-On).xlsx"), header, data)


# ---------- practice sheets (top paragraph + body) ----------

def _split_top(text, body_re):
    """Leading lines before the first body line (scenario/question/numbered)
    form the top paragraph; the rest is the body, verbatim. A marker line can
    never start the body (body_re cannot match one), so it rides with whichever
    part it sat in and surfaces in that part's own image column."""
    top, body = [], []
    for ln in text.split("\n"):
        s = ln.strip()
        if not s:
            continue
        if not body and not body_re.match(_bare(s)):
            top.append(s)
        else:
            body.append(s)
    return "\n".join(top).strip(), "\n".join(body).strip()


# Question numbering is not uniform across the corpus: most docs type "1." or
# "1)", but some type the digit with no separator at all ("1 Explain in your
# own words why ..."). The separator is therefore optional -- `_KC_BODY` below
# already allowed this; GP/IP were the odd ones out and silently swallowed a
# whole question list into the top paragraph.
_GP_BODY = re.compile(r'^(\d{1,2}[.)]?\s|scenario\b|situation\b)', re.I)
_IP_BODY = re.compile(r'^(\d{1,2}[.)]?\s|[a-e][.)]\s|[A-E]\.\s)', re.I)
_KC_BODY = re.compile(r'^\d{1,2}[.)]?\s|^what\b|^which\b|^why\b|^who\b|^how\b',
                      re.I)


# The practice sheets and the Answer Solution sheet MUST agree on where one
# question ends and the next begins, so both go through `split_questions`
# below -- there is exactly one implementation of question fanning.

# A question-type label opening a line ("Multiple Choice: ...", "True or
# False: ...") numbers a question just as definitely as "1." does. It is used
# ONLY in the safety-net retry below, never in normal fanning: widening the
# section body regexes would re-split documents that already fan correctly.
_QLABEL_LINE = re.compile(
    r'^(multiple\s+choice|true\s+or\s+false|short\s+answer'
    r'|fill\s+in\s+the\s+blanks?|matching)\s*[:—–-]', re.I)


def _fan_q_retry(lines):
    """`_fan_q` for the zero-question retry: a question-type label opens a new
    item as well as a number. Everything else behaves identically."""
    items = []
    for ln in lines:
        if not ln.strip():
            continue
        bare = _bare(ln).strip()
        m = _NUMLINE_Q.match(bare)
        if m:
            items.append([m.group(2).strip()])
        elif _QLABEL_LINE.match(bare):
            items.append([ln.strip()])
        elif items:
            items[-1].append(ln.strip())
        else:
            items.append([ln.strip()])
    return ["\n".join(x).strip() for x in items]


_AI_QSPLIT_PROMPT = """The text below is one practice section from a school \
lesson document. It normally opens with an instruction or scenario paragraph \
and is then followed by the questions students must answer.

Reply with ONLY the exact line of the text on which the FIRST question starts, \
copied character for character. No quotes, no explanation, no commentary. If \
the text contains no questions at all, reply NONE.

Text:
{text}"""


def split_questions(text, body_re, ai=None, ctx=""):
    """(top paragraph, [question, ...]) for one practice section.

    `body_re` is the section's body regex (None for a section with no top
    paragraph, e.g. the challenge). The deterministic path is `_split_top`
    followed by `_fan_q`, exactly as before.

    Safety net: a section that has body text but fans into NO questions has
    had everything swallowed by the top-paragraph split. It is retried over
    the full section text with the top split ignored, and with question-type
    labels ("Multiple Choice:", "True or False:") allowed to open an item as
    well as numbers; if that yields two or more items they are used with an
    empty top paragraph. If that still fails
    and `ai` was given, the model is asked to point at the line the first
    question starts on -- a reply is accepted only when it is a verbatim
    substring, so the model can locate a boundary but never write content."""
    if not text or not text.strip():
        return "", []
    top, body = _split_top(text, body_re) if body_re else ("", text)
    items = _fan_q(body.split("\n")) if body.strip() else []
    if items:
        return top, items
    retry = _fan_q_retry(text.split("\n"))
    if len(retry) >= 2:
        return "", retry
    if ai:
        idx = _ai_locate(text, _AI_QSPLIT_PROMPT, ai, ctx, "first question line")
        if idx:
            rest = text[idx:]
            fanned = _fan_q(rest.split("\n"))
            if fanned:
                return text[:idx].strip(), fanned
    return top, items


def _q_cols(name, n):
    """Header names for n fanned question columns."""
    return [f"{name} {i}" for i in range(1, n + 1)]


def _write_practice(rows, out_dir, filename, section, body_re, lead_cols,
                    q_name, ai=None):
    """Shared body of the four practice sheets: the leading columns, then one
    numbered column per question. The column count is the maximum question
    count across all documents (two-pass); shorter rows pad with ''."""
    parsed = []
    for r in _sorted_rows(rows):
        text = _body(r, section)
        if not text:
            continue
        top, items = split_questions(
            text, body_re, ai, ctx=f"{r.get('_file', '?')} / {section}")
        parsed.append((r, top, [_C(x) for x in items]))
    n = max([len(q) for _, _, q in parsed] + [1])
    header = ["Module Number", "Module Title 1"] + lead_cols + _q_cols(q_name, n)
    data = [[_mod(r), _title(r)]
            + ([_C(top)] if lead_cols else [])
            + pad_cells(q, n)
            for r, top, q in parsed]
    return _finalize(os.path.join(out_dir, filename), header, data)


def write_guided(rows, out_dir, ai=None):
    return _write_practice(
        rows, out_dir, "Guided Practice (With Support).xlsx",
        "Guided Practice", _GP_BODY, ["Top paragraph"],
        "Practice Scenarios", ai)


def write_independent(rows, out_dir, ai=None):
    return _write_practice(
        rows, out_dir, "Independent Practice.xlsx",
        "Independent Practice", _IP_BODY, ["Top paragraph"],
        "Independent Practice", ai)


def write_challenge(rows, out_dir, ai=None):
    # the challenge section has no top-paragraph split (see `_questions`), so
    # the sheet has no intro column; "Section Content" was a generic name, so
    # the fanned columns are plain "Question N"
    return _write_practice(
        rows, out_dir, "Challenge Higher-Order Thinking.xlsx",
        "Challenge / Higher-Order Thinking", None, [], "Question", ai)


def write_knowledge_check(rows, out_dir, ai=None):
    return _write_practice(
        rows, out_dir, "Knowledge Check (Quick Assessment).xlsx",
        "Knowledge Check", _KC_BODY, ["Instruction"], "Questions", ai)


# ---------- common mistakes ----------

_MISTAKE_RE = re.compile(r'^mistake\s*[:—–-]?\s*', re.I)
_TIP_RE = re.compile(r'^tip\s*[:—–-]?\s*', re.I)

# Some documents write a whole Mistake/Tip pair as ONE paragraph:
#   "• Mistake: <text>. Tip: <text>"
# The line-start regexes above then never see the Tip. This finds a tag that
# starts mid-line: it must be a standalone word (not "multip:"...) and must be
# followed by a colon or dash, so the bare word "tip" inside prose never
# triggers a split.
_INLINE_TAG = re.compile(r'(?<![\w-])(mistake|tip)\s*[:—–-]\s', re.I)


def _split_inline_tags(line):
    """Break one line into segments, each opening with its own Mistake/Tip tag.

    The tag text is kept (original casing) and the first segment keeps whatever
    bullet/indent the line started with. A line with no mid-line tag comes back
    unchanged, as a single-element list."""
    if _is_img(line) or _is_treq(line):
        return [line]
    cuts = [m.start() for m in _INLINE_TAG.finditer(line)]
    # a tag that opens the line (ignoring bullet/indent) is handled by the
    # existing line-start match -- only mid-line tags force a cut
    lead = len(line) - len(_bare(line))
    cuts = [c for c in cuts if c > lead]
    if not cuts:
        return [line]
    segs, prev = [], 0
    for c in cuts + [len(line)]:
        part = line[prev:c].strip()
        if part:
            segs.append(part if prev else line[prev:c].rstrip())
        prev = c
    return segs or [line]


# an over-long Mistake cell with no Tip that still mentions a tip somewhere is
# the case worth asking the model about (deterministic parsing already failed)
_AI_MIN_LEN = 400

_AI_TIP_PROMPT = """The text below is one "common mistake" note from a school \
lesson document. It should be two parts: the mistake itself, and a tip that \
follows it. Find where the tip begins.

Reply with ONLY the exact substring of the text that starts the tip, copied \
character for character, starting at the first word of the tip. No quotes, no \
explanation, no commentary. If there is no tip, reply NONE.

Text:
{text}"""


def _ai_split_tip(cell, ai, ctx=""):
    """(mistake, tip) for an under-split Mistake cell, using `ai` only to LOCATE
    the split point. Returns None when no confident verbatim split is found, so
    the deterministic result stands. `ai` is a callable taking a prompt and
    returning the reply text (or None)."""
    reply = _ai_ask(ai, _AI_TIP_PROMPT.format(text=cell), ctx, "mistake/tip split")
    if not reply:
        return None
    reply = reply.strip().strip('"').strip()
    if not reply or reply.upper().startswith("NONE"):
        return None
    # verbatim only: the reply must occur in the cell exactly as written
    idx = cell.find(reply)
    if idx <= 0:
        return None
    return cell[:idx].rstrip(), cell[idx:].strip()


def write_mistakes(rows, out_dir, ai=None):
    """One row per document, with each Mistake/Tip pair fanned into its own
    "Mistake N"/"Tip N" column pair -- the number of pairs is the maximum
    pair count across all documents; shorter rows pad with ''. Text that
    carries no explicit tags is kept whole in the Mistake column so nothing
    is dropped. An image marker is never a Mistake or a Tip of its own -- it
    rides with the part it sat in and lands in that column's own adjacent
    image column.

    `ai`, when given, is a callable prompt -> reply text (or None). It is used
    ONLY as a last-resort locator for a Tip that deterministic parsing could
    not separate; the split is always a verbatim slice of the parsed text."""
    parsed = []
    for r in _sorted_rows(rows):
        text = _body(r, "Common Mistakes & Tips")
        if not text.strip():
            continue
        pairs = [["", ""]]
        cur = None
        tagged = False
        # (segment, opens a line that was split on an inline tag)
        lines = []
        for ln in text.split("\n"):
            segs = _split_inline_tags(ln)
            lines.extend((s, len(segs) > 1 and i == 0)
                         for i, s in enumerate(segs))
        for ln, split_head in lines:
            s = ln.strip()
            if not s:
                continue
            if _MISTAKE_RE.match(_bare(s)):
                tagged = True
                if pairs[-1][0]:
                    pairs.append(["", ""])
                pairs[-1][0] = s
                cur = 0
            elif _TIP_RE.match(_bare(s)):
                tagged = True
                if pairs[-1][1]:      # a second Tip means a new pair started
                    pairs.append(["", ""])
                pairs[-1][1] = s
                cur = 1
            elif split_head:
                # an untagged mistake that carries its Tip inline on the same
                # source line -- the line break, not a tag, delimits the pair
                tagged = True
                if pairs[-1][0] or pairs[-1][1]:
                    pairs.append(["", ""])
                pairs[-1][0] = s
                cur = 0
            elif cur is not None:
                pairs[-1][cur] = (pairs[-1][cur] + "\n" + s).strip()
            else:
                pairs[-1][0] = join_nonempty(pairs[-1][0], s)
        if not tagged and pairs == [["", ""]]:
            pairs = [[text.strip(), ""]]
        cells = []
        for m, t in pairs:
            if ai and not t and len(m) > _AI_MIN_LEN and "tip" in m.lower():
                split = _ai_split_tip(
                    m, ai, ctx=f"{r.get('_file', '?')} / Common Mistakes & Tips")
                if split:
                    m, t = split
            if m or t:
                cells += [_C(m), _C(t)]
        if cells:
            parsed.append((r, cells))
    n = max([len(cells) // 2 for _, cells in parsed] + [1])
    header = ["Module Number", "Module Title 1"]
    for i in range(1, n + 1):
        header += [f"Mistake {i}", f"Tip {i}"]
    data = [[_mod(r), _title(r)] + pad_cells(cells, n * 2)
            for r, cells in parsed]
    return _finalize(os.path.join(
        out_dir, "Common Mistakes & Tips.xlsx"), header, data)


def write_real_world(rows, out_dir):
    header = ["Module Number", "Module Title 1", "Intro/ paragraph", "Points",
              "Think about it "]
    data = []
    for r in _sorted_rows(rows):
        text = _body(r, "Real-World Application")
        if not text:
            continue
        # a trailing "Think about it"/question block splits off when marked
        think = ""
        m = re.search(r'^[\s•·●▪\t-]*think about it\b.*$', text, re.I | re.M)
        if m:
            text, think = text[:m.start()].rstrip(), text[m.start():].strip()
        data.append([_mod(r), _title(r), _C(text), "", _C(think)])
    return _finalize(os.path.join(
        out_dir, "Real-World Application.xlsx"), header, data)


def write_summary(rows, out_dir):
    header = ["Module Number", "Module Title 1", "Summary Content "]
    data = []
    for r in _sorted_rows(rows):
        obj = _C(_body(r, "Summary / Key Takeaways"))
        if obj:
            data.append([_mod(r), _title(r), obj])
    return _finalize(os.path.join(
        out_dir, "Summary Key Takeaways.xlsx"), header, data)


# ---------- answers & solutions ----------

# sub-blocks inside the Answers & Solutions section, matched by name
_ANS_BLOCKS = [
    (re.compile(r'^guided practice\b.*[:]?', re.I),      "gp"),
    (re.compile(r'^independent practice\b.*[:]?', re.I), "ip"),
    (re.compile(r'^(challenge|higher-order|evaluate a scenario)\b.*', re.I), "hot"),
    (re.compile(r'^knowledge check\b.*[:]?', re.I),      "kc"),
]

_NUMLINE = re.compile(r'^(\d{1,2})[.)]\s*(.*)$')

# `_NUMLINE` for QUESTION fanning only, covering the two numbering styles the
# strict form misses: no separator at all ("1 Explain ...") and section-style
# sub-numbering ("7.1 Mixed scenario reasoning" -- which `_NUMLINE` matched as
# "7." + a body of "1 Mixed scenario reasoning", mangling the text). It insists
# on real text after the number so a bare "12" line is never eaten.
_NUMLINE_Q = re.compile(r'^(\d{1,2}(?:\.\d{1,2})*)[.)]?[ \t]+(\S.*)$')

# docs label their answer sub-blocks with a numbering prefix as often as not:
# "5. Guided Practice:", "Section 6: Guided Practice". Stripped before match.
_ANS_LABEL_NUM = re.compile(r'^(?:section\s+)?\d{1,2}(?:\.\d{1,2})*\s*[.):]\s*',
                            re.I)
# a trailing parenthetical aside on a label line ("(Note: Minor variations
# ... are acceptable)") is commentary, not content -- ignored when judging
# whether the line is short enough to be a label rather than an answer
_ANS_LABEL_PAREN = re.compile(r'\s*\([^)]*\)\s*[.:;]?\s*$')


def _ans_label(s):
    """The answer sub-block this line labels, or None if it is not a label.

    Only the category NAME has to lead the line -- surrounding numbering and
    a trailing parenthetical note are stripped first. The length guard then
    runs on what is left, so a genuine label stays a label however it is
    decorated, while a long answer that merely opens with a category name
    still fails it."""
    core = _ANS_LABEL_PAREN.sub("", _ANS_LABEL_NUM.sub("", _bare(s).strip()))
    if not core or len(core) > 60:
        return None
    for rx, b in _ANS_BLOCKS:
        if rx.match(core):
            return b
    return None


# a sub-point bullet under the answer above it -- never an answer of its own
_ANS_SUBPOINT = re.compile(r'^[\t ]*[•▪‣◦·⁃]')

# "Question 3" as a bare line: a worked-solution label. It numbers its item
# the way "3." does, but the answer text sits on the lines BELOW it (and the
# numbers need not run in sequence -- docs work only selected questions).
_ANS_QLABEL = re.compile(r'^question\s+\d{1,2}\s*[:.]?\s*$', re.I)

# a sub-heading introducing a run of worked solutions ("Step-by-step
# solutions for Guided Practice", "Selected step-by-step solutions for
# Independent Practice"). It labels what follows instead of answering
# anything, so it is carried onto the next item rather than becoming one.
_ANS_SUBHEAD = re.compile(
    r'^(?:selected\s+)?(?:step[-\s]by[-\s]step\s+)?(?:model\s+)?solutions?\s+for\b',
    re.I)


# the answer to a "Match the term to its definition" question is a run of
# short "Term → letter" lines. They belong to ONE question, so they must
# collapse into one answer instead of each claiming a slot of its own.
_MAPLINE = re.compile(r'^.{1,60}\s*(?:→|->|=>)\s*\S')


def _is_mapline(bare):
    return len(bare) <= 80 and bool(_MAPLINE.match(bare))


# "a) ..." / "B. ..." -- a lettered option or step. The whole alphabet, not just
# a-e: a worked answer can list steps A. through H. and every one of them
# belongs to the same answer.
_LETTER_LINE = re.compile(r'^[a-z][).]\s', re.I)


def _fan(lines):
    """Fan answer lines into items. Explicit '1.' numbering wins; unnumbered
    lines each start a new item. No slot cap -- callers size and pad their
    own columns."""
    items = []
    head = []       # pending sub-heading lines, carried onto the next item
    below = False   # current item is a "Question N" label -- text follows it
    mapping = False  # current item is a run of "term → letter" mapping lines
    for ln in lines:
        # the table placeholder and image markers are content, never answers
        # of their own -- they ride with the answer they followed so the
        # numbered slots don't shift. The marker is separated out into the
        # slot's adjacent image column later, by `_C`.
        if (_is_treq(ln) or _is_img(ln)) and items:
            items[-1].append(ln)
            continue
        bare = _bare(ln).strip()
        m = _NUMLINE.match(bare)
        if _ANS_SUBHEAD.match(bare):
            head.append(ln.strip())
            continue
        mapline = _is_mapline(m.group(2).strip() if m else bare)
        if mapline and mapping and items:
            # continue the matching-exercise answer above -- the whole run of
            # "term → letter" lines answers ONE question
            items[-1].append(m.group(2).strip() if m else ln.strip())
            continue
        if m:
            items.append(head + [m.group(2).strip()])
            head, below, mapping = [], False, mapline
        elif _ANS_QLABEL.match(bare):
            # keep the label itself -- it says WHICH question this solves
            items.append(head + [ln.strip()])
            head, below, mapping = [], True, False
        elif items and _LETTER_LINE.match(_bare(ln).strip()) \
                and not _LETTER_LINE.match(_bare(items[-1][0])):
            # lettered option/step lines ride with the item above them --
            # unless the previous item is itself a lettered answer
            # (knowledge-check style: each "b) ..." line is its own answer)
            items[-1].append(ln)
        elif items and _ANS_SUBPOINT.match(ln) \
                and not _ANS_SUBPOINT.match(items[-1][0]):
            # a bulleted sub-point elaborates the answer above it -- unless
            # that answer is itself a bullet, in which case this block uses
            # bullets as its answer form and each one stands alone
            items[-1].append(ln)
        elif items and _bare(ln).strip().lower().startswith(("reason", "fix", "explanation")):
            items[-1].append(ln)
        elif below and items:
            # the body of a "Question N" worked solution
            items[-1].append(ln.strip())
        elif ln.strip():
            items.append(head + [ln.strip()])
            head, mapping = [], mapline
    if head:                       # a heading with nothing after it -- keep it
        items.append(head)
    return ["\n".join(x).strip() for x in items]


def _fan_q(lines):
    """Fan QUESTION lines into items -- stricter than `_fan`.

    Question bodies are not reliably numbered the way answers are, and a
    scenario routinely runs over several unnumbered lines ("Situation: ...",
    "Question: ...", "Hint: ..."). Only an explicit "1." / "2)" line may open
    a new item here; every other line rides with the item above it. A body
    with no numbered line at all is therefore one single question.

    The number's trailing "." / ")" is optional here (see `_GP_BODY`): docs that
    write "1 Explain ..." number their questions just as definitely as docs
    that write "1. Explain ...". Answers keep the strict `_NUMLINE` -- there the
    separator is what distinguishes a counter from an answer that happens to
    open with a figure."""
    items = []
    for ln in lines:
        if not ln.strip():
            continue
        m = _NUMLINE_Q.match(_bare(ln).strip())
        if m:
            items.append([m.group(2).strip()])
        elif items:
            items[-1].append(ln.strip())
        else:
            items.append([ln.strip()])
    return ["\n".join(x).strip() for x in items]


def _questions(r, ai=None):
    """The question items for each answer category, pulled from the ORIGINAL
    practice sections of the same document, so question i can be paired
    positionally with answer i. Uses the same `split_questions` helper as the
    practice sheets, so the two always agree on question boundaries."""
    def fan_body(section, body_re):
        return split_questions(
            _body(r, section), body_re, ai,
            ctx=f"{r.get('_file', '?')} / {section}")[1]

    return {
        "gp": fan_body("Guided Practice", _GP_BODY),
        "ip": fan_body("Independent Practice", _IP_BODY),
        "hot": fan_body("Challenge / Higher-Order Thinking", None),
        "kc": fan_body("Knowledge Check", _KC_BODY),
    }


def write_answers(rows, out_dir, ai=None):
    """Question/answer pairs per category: each answer slot is preceded by the
    matching question slot from the source practice section. Every slot gets
    its own adjacent image column when any document has a marker in it."""
    parsed = []
    for r in _sorted_rows(rows):
        text = _body(r, "Answers & Solutions")
        if not text.strip():
            continue
        buckets = {"gp": [], "ip": [], "hot": [], "kc": [], "": []}
        cur = ""
        for ln in text.split("\n"):
            s = ln.strip()
            if not s:
                continue
            hit = _ans_label(s)
            if hit:
                # "Evaluate a Scenario:" is a sub-part of the challenge block
                if hit == "hot" and cur == "hot":
                    buckets["hot"].append(s)
                else:
                    cur = hit
                continue
            buckets[cur].append(s)
        # answers before any recognised block label ride on Guided Practice
        if buckets[""]:
            buckets["gp"] = buckets[""] + buckets["gp"]
        qs = _questions(r, ai)
        parsed.append((r,
                       {k: [_C(x) for x in _fan(buckets[k])]
                        for k in ("gp", "ip", "hot", "kc")},
                       {k: [_C(x) for x in qs[k]]
                        for k in ("gp", "ip", "hot", "kc")}))
    cats = [("gp", "Guided Practice"), ("ip", "Independent Practice"),
            ("hot", "Higher-Order Thinking"), ("kc", "Knowledge Check")]
    # a category is as wide as its longest side in any doc -- questions and
    # answers are never truncated to match each other, the short side pads
    widths = {k: max([len(f[k]) for _, f, _ in parsed]
                     + [len(q[k]) for _, _, q in parsed] + [1])
              for k, _ in cats}
    header = ["Module Number", "Module Title 1"]
    for k, label in cats:
        for i in range(1, widths[k] + 1):
            header += [f"{label} Question {i}", f"{label} Answer {i}"]
    data = []
    for r, fans, ques in parsed:
        cells = []
        for k, _ in cats:
            q = pad_cells(ques[k], widths[k])
            a = pad_cells(fans[k], widths[k])
            for i in range(widths[k]):
                cells += [q[i], a[i]]
        data.append([_mod(r), _title(r)] + cells)
    return _finalize(os.path.join(
        out_dir, "Answer Solution.xlsx"), header, data)


# ---------- suggested interactive moments ----------

_MOMENT_START = re.compile(r'^(moment\s*\d+|location\s*[:—–-]|where it fits\s*[:—–-])',
                           re.I)
_SECTION_LINE = re.compile(r'^(location|where it fits)\s*[:—–-]\s*(.*)$', re.I)


def write_moments(rows, out_dir):
    """One row per suggested moment. Groups start at a 'Moment N' or
    'Location:'/'Where it fits:' line; the Section column carries the
    location, the Suggestion column the whole group verbatim."""
    header = ["Module title ", "Section ", "Suggestion"]
    data = []
    for r in _sorted_rows(rows):
        text = _body(r, "Suggested Interactive Moments")
        if not text.strip():
            continue
        groups = []
        for ln in text.split("\n"):
            s = ln.strip()
            if not s:
                continue
            # a new group starts at "Moment N", or at a Location line when
            # the current group already has one (bullet-list style docs)
            # a marker line never opens a group -- it belongs to the moment
            # it was written under
            if _is_img(s) and groups:
                groups[-1].append(s)
                continue
            is_loc = bool(_SECTION_LINE.match(_bare(s)))
            starts_new = bool(re.match(r'^moment\s*\d+', _bare(s), re.I)) or (
                is_loc and groups and any(_SECTION_LINE.match(_bare(x))
                                          for x in groups[-1]))
            if starts_new or not groups:
                groups.append([s])
            else:
                groups[-1].append(s)
        mod_title = join_nonempty(_mod(r), _title(r), sep=" ")
        for g in groups:
            section = ""
            for x in g:
                m = _SECTION_LINE.match(_bare(x))
                if m:
                    section = m.group(2).strip()
                    break
            data.append([mod_title, section, _C("\n".join(g))])
    return _finalize(os.path.join(
        out_dir, "Suggested Interactive Moments.xlsx"), header, data)


def write_all_it(rows, out_dir, ai=None):
    """`ai` is an optional callable prompt -> reply text (or None), used only
    to locate split points the deterministic parsers could not find. With
    ai=None (the default) the output is fully deterministic, as before."""
    os.makedirs(out_dir, exist_ok=True)
    # one cache per run: the same section is handed to `split_questions` by up
    # to five writers, and the model must not be asked about it five times
    reset_ai_cache()
    # documents that contain the whole lesson twice (draft + revised copy) are
    # reduced to the revised copy ONCE here, before any sheet is written
    rows = dedupe_rows(rows, ai=ai)
    # embedded pictures land on disk once, then get anchored into whichever
    # sheet's Image cell carries their marker
    _index_images(rows, out_dir)
    paths = [
        write_full_extract(rows, out_dir),
        write_cover(rows, out_dir),
        write_intro(rows, out_dir),
        write_key_concepts(rows, out_dir),
        write_practical(rows, out_dir),
        write_guided(rows, out_dir, ai=ai),
        write_independent(rows, out_dir, ai=ai),
        write_challenge(rows, out_dir, ai=ai),
        write_knowledge_check(rows, out_dir, ai=ai),
        write_mistakes(rows, out_dir, ai=ai),
        write_real_world(rows, out_dir),
        write_summary(rows, out_dir),
        write_answers(rows, out_dir, ai=ai),
        write_moments(rows, out_dir),
    ]
    return paths
