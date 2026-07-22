"""
writers.py

Writes the parsed rows to CSV files. Output = the seven "Spread Sheets 2"
export layouts plus wide.csv (everything, one row per doc, for debugging):

  wide.csv                                       (all fields, every document)
  ELA Mini Practice.csv                          (per lesson)
  ELA Individual Sub Unit My Checklist, Quick Recap, Reflect.csv (per lesson)
  _ELA Main Comprehension Text.csv               (per lesson that has a text)
  Unit Intro Page.csv                            (per Unit N Introduction)
  ELA End Of Unit Portfolio Doc.csv              (per End of Unit N)
  ELA  Overall Portfolio Project.csv             (per End of Unit N Portfolio)
  ELA Check Your Understanding Term 2.csv         (Term 2 docs only; header-only
                                                  when no such source is present)

plus the four top-level "Corrected OUTPUT" layouts (restored 2026-07-20):

  ELA Main Lesson.csv                            (per lesson; Main Lesson split
                                                  into Top Content / Step 1-4)
  ELA Bridge Back & Vocabulary Vault.csv         (per lesson; Vocab Vault split
                                                  into Key Word/Definition pairs)
  ELA Mistake Spotter.csv                        (per lesson; Key Tag/content pairs)
  ELA Unit Cover (Goals & By End Of Lesson Content).csv  (per lesson)

All values are verbatim slices from the parser -- no rewriting happens here.

Author: AbhishekAEDan
"""
__author__ = "AbhishekAEDan"

import csv
import os
import re

from parser_core import ORDER


# ---------- helpers ----------

def order_cols(rows):
    cols = [c for c in ORDER if any(c in r for r in rows)]
    for r in rows:
        for k in r:
            if k not in cols:
                cols.append(k)
    return cols


def write_csv(path, header, data_rows):
    with open(path, "w", newline="", encoding="utf-8-sig") as fh:
        w = csv.writer(fh)
        w.writerow(header)
        for r in data_rows:
            w.writerow(r)


def unit_sort_key(unit):
    """Sort '1', '1.1', '10.2' numerically so sheets read in unit order."""
    parts = re.findall(r'\d+', str(unit or ""))
    return tuple(int(p) for p in parts) if parts else (9999,)


def _rows_of_type(rows, *doc_types):
    out = [r for r in rows if r.get("_doc_type") in doc_types]
    return sorted(out, key=lambda r: unit_sort_key(r.get("Unit Number", "")))


def join_nonempty(*parts, sep="\n"):
    return sep.join(p for p in parts if p and p.strip()).strip()


# ---------- lesson-level sheets ----------

def write_wide(rows, out_dir):
    cols = order_cols(rows)
    path = os.path.join(out_dir, "wide.csv")
    with open(path, "w", newline="", encoding="utf-8-sig") as fh:
        w = csv.DictWriter(fh, fieldnames=cols)
        w.writeheader()
        for r in rows:
            w.writerow({k: r.get(k, "") for k in cols})
    return path


def write_mini_practice(rows, out_dir):
    header = ["Unit Number", "Main Practice Page Title", "Page Top Content",
              "Practice A", "Practice B", "Practice C", "What to look for",
              "Answer Key"]
    data = []
    for r in _rows_of_type(rows, "lesson"):
        data.append([
            r.get("Unit Number", ""),
            r.get("Mini Practice Title", ""),
            r.get("Mini Practice Intro", ""),
            r.get("Practice A", ""), r.get("Practice B", ""),
            r.get("Practice C", ""), r.get("What to look for", ""),
            r.get("Mini Practice Answer Key", ""),
        ])
    path = os.path.join(out_dir, "ELA Mini Practice.csv")
    write_csv(path, header, data)
    return path


def write_sub_unit(rows, out_dir):
    header = ["Unit Number", "My Learning Checklist", "Quick Recap", "Reflect"]
    data = [[r.get("Unit Number", ""), r.get("My Learning Checklist", ""),
             r.get("Quick Recap", ""), r.get("Reflect", "")]
            for r in _rows_of_type(rows, "lesson")]
    path = os.path.join(out_dir,
                        "ELA Individual Sub Unit My Checklist, Quick Recap, Reflect.csv")
    write_csv(path, header, data)
    return path


def write_comprehension(rows, out_dir):
    header = ["Unit Number", "Comprehension Title", "Before Reading / Listening",
              "Main Text", "After Reading & Listening (Questions)", "Answer Key"]
    data = []
    for r in _rows_of_type(rows, "lesson"):
        cq = r.get("Comprehension Questions", "")
        after = join_nonempty(r.get("After Reading", ""),
                              ("Comprehension Questions\n" + cq) if cq else "")
        answer = r.get("Comprehension Questions Answer Key", "") \
            or r.get("After Reading Answer Key", "")
        title = r.get("Comprehension Title", "")
        main = r.get("Main Text", "")
        before = r.get("Before Reading", "")
        # some lessons title the passage with a bold "Prose Text:/Poem:/Drama:"
        # heading (kept as an unknown '?' column) instead of "Main Comprehension
        # Text" -- use that label as the title when the canonical one is empty.
        if not title:
            for k in r:
                if k.startswith("?") and re.search(
                        r'\b(prose|poem|drama|play|comprehension) ', k, re.I):
                    title = k[1:]
                    break
        # only lessons that actually carry a comprehension text
        if not any([title, before, main, after]):
            continue
        data.append([r.get("Unit Number", ""), title, before, main, after, answer])
    path = os.path.join(out_dir, "_ELA Main Comprehension Text.csv")
    write_csv(path, header, data)
    return path


# ---------- restored "Corrected OUTPUT" top-level sheets ----------

_STEP_RE = re.compile(r'^step\s*(\d+)\b', re.I)


def _split_main_lesson(text):
    """Split the verbatim Main Lesson text into
    (title, top_content, {step_no: step_text}). The first line is the lesson
    subtitle when it isn't already a Step heading; lines before 'Step 1' are
    Top Content; each 'Step N ...' heading starts that step's block."""
    title, top, steps = "", [], {}
    cur = None                       # None -> top content, else step number
    lines = [l for l in (text or "").split("\n")]
    if lines and lines[0].strip() and not _STEP_RE.match(lines[0].strip()):
        title = lines[0].strip()
        lines = lines[1:]
    for ln in lines:
        m = _STEP_RE.match(ln.strip())
        if m:
            cur = int(m.group(1))
            steps.setdefault(cur, []).append(ln)
        elif cur is None:
            top.append(ln)
        else:
            steps[cur].append(ln)
    return (title, "\n".join(top).strip(),
            {k: "\n".join(v).strip() for k, v in steps.items()})


def write_main_lesson(rows, out_dir):
    header = ["Unit Number", "Title", "Top Content", "Step 1", "Step  2",
              "Step 3", "Step 3 Image", "Step 4", "Mini Practice ",
              "A", "B", "C", "Answer Key"]
    data = []
    for r in _rows_of_type(rows, "lesson"):
        title, top, steps = _split_main_lesson(r.get("Main Lesson", ""))
        mini = join_nonempty(r.get("Mini Practice Title", ""),
                             r.get("Mini Practice Intro", ""))
        if not any([title, top, steps, mini]):
            continue
        # steps 5+ (rare) are appended to Step 4 so nothing is dropped
        extra = join_nonempty(*[steps[k] for k in sorted(steps) if k > 4])
        step4 = join_nonempty(steps.get(4, ""), extra)
        data.append([
            r.get("Unit Number", ""), title, top,
            steps.get(1, ""), steps.get(2, ""), steps.get(3, ""),
            "", step4, mini,
            r.get("Practice A", ""), r.get("Practice B", ""),
            r.get("Practice C", ""), r.get("Mini Practice Answer Key", ""),
        ])
    path = os.path.join(out_dir, "ELA Main Lesson.csv")
    write_csv(path, header, data)
    return path


_VOCAB_ENTRY_RE = re.compile(r'^([^:]{1,60}?):\s*(.*)$')


def write_bridge_vocab(rows, out_dir):
    header = (["Unit Number"]
              + [c for i in range(1, 8)
                 for c in (f"Key Word {i}",
                           "Definition  1" if i == 1 else f"Definition {i}")]
              + [f"Vocabulary Vault {i}" for i in range(8, 12)])
    data = []
    for r in _rows_of_type(rows, "lesson"):
        entries = []                 # [(word, definition_lines)]
        for ln in (r.get("Vocab Vault", "") or "").split("\n"):
            m = _VOCAB_ENTRY_RE.match(ln.strip())
            if m:
                entries.append((m.group(1).strip(), [m.group(2).strip()]))
            elif entries and ln.strip():
                entries[-1][1].append(ln.strip())
        if not entries:
            continue
        row = [r.get("Unit Number", "")]
        for i in range(7):
            if i < len(entries):
                row += [entries[i][0], "\n".join(entries[i][1]).strip()]
            else:
                row += ["", ""]
        overflow = [f"{w}: " + "\n".join(d).strip() for w, d in entries[7:]]
        for i in range(4):
            if i == 3 and len(overflow) > 4:
                row.append("\n".join(overflow[3:]))   # never drop entries
            else:
                row.append(overflow[i] if i < len(overflow) else "")
        data.append(row)
    path = os.path.join(out_dir, "ELA Bridge Back & Vocabulary Vault.csv")
    write_csv(path, header, data)
    return path


# a Mistake Spotter tag line: short capitalised lead ending in ':'
# (e.g. "Common mistake:", "How to improve it:") -- quotes/sentences excluded
_TAG_RE = re.compile(r'^([A-Z][^:.!?"“”]{2,45}):\s*(.*)$')


def write_mistake_spotter(rows, out_dir):
    header = ["Unit Number "]
    for i in range(1, 11):
        tail = " " if i in (6, 8) else ""
        header += [f"Key Tag {i}", f"Mistake Spotter {i}{tail}"]
    data = []
    for r in _rows_of_type(rows, "lesson"):
        pairs = []                   # [(tag, content_lines)]
        for ln in (r.get("Mistake Spotter", "") or "").split("\n"):
            m = _TAG_RE.match(ln.strip())
            if m:
                pairs.append((m.group(1).strip() + ":", [m.group(2).strip()]))
            elif pairs and ln.strip():
                pairs[-1][1].append(ln.strip())
            elif ln.strip():         # content before any tag
                pairs.append(("", [ln.strip()]))
        if not pairs:
            continue
        if len(pairs) > 10:          # never drop content: merge extras into #10
            rest = pairs[9:]
            merged = "\n".join(
                (f"{t} " if t else "") + "\n".join(c) for t, c in rest[1:])
            pairs = pairs[:9] + [(rest[0][0], rest[0][1] + [merged])]
        row = [r.get("Unit Number", "")]
        for i in range(10):
            if i < len(pairs):
                row += [pairs[i][0], "\n".join(pairs[i][1]).strip()]
            else:
                row += ["", ""]
        data.append(row)
    path = os.path.join(out_dir, "ELA Mistake Spotter.csv")
    write_csv(path, header, data)
    return path


def write_unit_cover(rows, out_dir):
    header = ["Unit Number", "Sub Unit Title", "My Goal Information",
              "By end of lesson content", "Bridge Back Content"]
    data = [[r.get("Unit Number", ""), r.get("Lesson Title", ""),
             r.get("My Goal", ""), r.get("Lesson Objectives", ""),
             r.get("Bridge Back", "")]
            for r in _rows_of_type(rows, "lesson")]
    path = os.path.join(out_dir,
                        "ELA Unit Cover (Goals & By End Of Lesson Content).csv")
    write_csv(path, header, data)
    return path


# ---------- unit-level sheets ----------

def write_unit_intro(rows, out_dir):
    header = ["Unit Number", "Title", "Top Content", "At the End of Unit check list"]
    data = [[r.get("Unit Number", ""), r.get("Unit Title", ""),
             r.get("Top Content", ""), r.get("At End Of Unit Checklist", "")]
            for r in _rows_of_type(rows, "unit_intro")]
    path = os.path.join(out_dir, "Unit Intro Page.csv")
    write_csv(path, header, data)
    return path


def write_end_of_unit(rows, out_dir):
    header = ["Unit Number", "My Unit Learning Checklist", "Unit Quick Recap",
              "Reflect"]
    data = [[r.get("Unit Number", ""), r.get("Unit Learning Checklist", ""),
             r.get("Unit Quick Recap", ""), r.get("Unit Reflect", "")]
            for r in _rows_of_type(rows, "end_of_unit")]
    path = os.path.join(out_dir, "ELA End Of Unit Portfolio Doc.csv")
    write_csv(path, header, data)
    return path


def write_portfolio(rows, out_dir):
    header = ["Unit Number", "Project Title", "Purpose", "Portfolio Link",
              "What you will create", "Materials", "Steps",
              "Template / How to fill it", "Example", "Common mistakes & fixes",
              "Submission Checklist", "Optional extension", "End-of-Unit Inventory",
              "Motivation + Reflection Skill Badge Earned", "Reflect"]
    field = ["Project Title", "Purpose", "Portfolio Link", "What You Will Create",
             "Materials", "Steps", "Template How To Fill", "Example",
             "Common Mistakes And Fixes", "Submission Checklist",
             "Optional Extension", "End-of-Unit Inventory", "Motivation Reflection",
             "Unit Reflect"]
    data = []
    for r in _rows_of_type(rows, "end_of_unit_portfolio"):
        data.append([r.get("Unit Number", "")] + [r.get(f, "") for f in field])
    path = os.path.join(out_dir, "ELA  Overall Portfolio Project.csv")
    write_csv(path, header, data)
    return path


_NUMBERED_RE = re.compile(r'^(\d{1,2})[.)]\s*(.*)$')


def _fan_numbered(text, slots=5):
    """Split verbatim text into (preface, [item1..item5]). Lines starting
    '1.' / '2)' begin numbered items; earlier lines are the preface;
    continuation lines stay with their item. Items past `slots` are appended
    to the last slot so nothing is dropped."""
    preface, items = [], []
    for ln in (text or "").split("\n"):
        m = _NUMBERED_RE.match(ln.strip())
        if m:
            items.append([m.group(2).strip()])
        elif items:
            items[-1].append(ln)
        elif ln.strip():
            preface.append(ln)
    # docs numbered via Word list formatting carry no visible "1." -- there
    # every non-blank line is its own question/answer
    if not items and preface:
        items = []
        for ln in preface:
            # a)/b)/c) choice lines belong to the question above them
            if items and re.match(r'^[a-eA-E][).]\s', ln.strip()):
                items[-1].append(ln)
            else:
                items.append([ln])
        preface = []
    flat = ["\n".join(x).strip() for x in items]
    if len(flat) > slots:
        flat = flat[:slots - 1] + ["\n".join(flat[slots - 1:])]
    flat += [""] * (slots - len(flat))
    return "\n".join(preface).strip(), flat


def write_check_understanding(rows, out_dir):
    """One row per lesson that carries a 'Check for Understanding' block.
    The block's numbered lines are fanned into Question 1-5; its Answer Key
    into Answer 1-5; any lead-in text becomes the Preface."""
    header = ["Unit Number", "Preface", "Question 1", "Question 2", "Question 3",
              "Question 4", "Question 5", "Answer 1", "Answer 2", "Answer 3",
              "Answer 4", "Answer 5"]
    data = []
    for r in _rows_of_type(rows, "lesson", "end_of_unit"):
        cfu = r.get("Check for Understanding", "")
        ak = r.get("Check for Understanding Answer Key", "")
        if not (cfu or ak):
            continue
        # some docs keep the answers inside the block after a
        # "Check your answers:" line instead of a bold Answer Key section
        if not ak:
            m = re.search(r'^\s*check your answers\b.*$', cfu,
                          re.I | re.M)
            if m:
                cfu, ak = cfu[:m.start()].rstrip(), cfu[m.end():].strip()
        preface, qs = _fan_numbered(cfu)
        ak_preface, ans = _fan_numbered(ak)
        # unnumbered blocks can't tell a lead-in instruction from question 1;
        # when there is exactly one more question than answers, the first
        # "question" is really the preface
        nq = sum(1 for q in qs if q)
        na = sum(1 for a in ans if a)
        unnumbered = not any(_NUMBERED_RE.match(l.strip())
                             for l in (cfu or "").split("\n"))
        if unnumbered and not preface and na and nq == na + 1:
            preface, qs = qs[0], qs[1:] + [""]
            nq -= 1
        # a single answer means one question (e.g. a numbered word bank is
        # options for one question, not five questions) -- and vice versa
        if na == 1 and nq > 1:
            qs = [join_nonempty(preface, *qs)] + [""] * 4
            preface = ""
        elif nq == 1 and na > 1:
            ans = [join_nonempty(*ans)] + [""] * 4
        # an answer key with no numbered lines still belongs in Answer 1
        if ak_preface and not any(ans):
            ans[0] = ak_preface
        data.append([r.get("Unit Number", ""), preface] + qs + ans)
    path = os.path.join(out_dir, "ELA Check Your Understanding Term 2.csv")
    write_csv(path, header, data)
    return path


def write_all(rows, out_dir):
    os.makedirs(out_dir, exist_ok=True)
    paths = [
        write_wide(rows, out_dir),
        write_mini_practice(rows, out_dir),
        write_sub_unit(rows, out_dir),
        write_comprehension(rows, out_dir),
        write_unit_intro(rows, out_dir),
        write_end_of_unit(rows, out_dir),
        write_portfolio(rows, out_dir),
        write_check_understanding(rows, out_dir),
        write_main_lesson(rows, out_dir),
        write_bridge_vocab(rows, out_dir),
        write_mistake_spotter(rows, out_dir),
        write_unit_cover(rows, out_dir),
    ]
    return paths
