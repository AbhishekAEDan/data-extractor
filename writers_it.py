"""
writers_it.py

Writes parsed IT lesson rows to CSV files. Layouts mirror the Form 1 IT
spreadsheet templates plus the "AI Sample Sheets/content sheets.xlsx"
examples (headers copied exactly, including their spacing/typo quirks --
"Acitvity Intro ", "Module Title 1", "Think about it " etc.):

    Unit cover page & Learning Objective.csv
    Introduction Topic Overview.csv
    Key Concepts (Main Content).csv
    Practical Activity (Hands-On).csv
    Guided Practice (With Support).csv
    Independent Practice.csv
    Challenge Higher-Order Thinking.csv
    Knowledge Check (Quick Assessment).csv
    Common Mistakes & Tips.csv        (one row per Mistake/Tip pair)
    Real-World Application.csv
    Summary Key Takeaways.csv
    Answer Solution.csv               (answers fanned into numbered slots)
    Suggested Interactive Moments.csv (one row per suggested moment)
    full_extract.csv                  (debug: every parsed field)

All values are verbatim slices from the parser -- no rewriting happens here.

Author: AbhishekAEDan
"""
__author__ = "AbhishekAEDan"

import csv
import os
import re

from parser_it import IT_ORDER
from writers import write_csv, unit_sort_key, join_nonempty


def _sorted_rows(rows):
    out = [r for r in rows if r.get("_doc_type") == "it_lesson"]
    return sorted(out, key=lambda r: unit_sort_key(r.get("Module Number", "")))


def order_cols(rows):
    cols = [c for c in IT_ORDER if any(c in r for r in rows)]
    for r in rows:
        for k in r:
            if k not in cols:
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


def _body_img(r, section):
    return split_images(r.get(section, ""))


def _bare(s):
    """Copy of a line with list markers/indent removed -- for MATCHING only;
    cell content keeps the markers."""
    return s.lstrip("\t •·●▪-–—")


def _mod(r):
    return r.get("Module Number", "")


def _title(r):
    return r.get("Module Title", "")


# ---------- simple one-block sheets ----------

def write_cover(rows, out_dir):
    header = ["Module Number", "Module Title 1", "Module Descriptive Title",
              "Learning Objectives", "Image"]
    data = []
    for r in _sorted_rows(rows):
        body, img = _body_img(r, "Learning Objectives")
        if body or img:
            data.append([_mod(r), _title(r),
                         r.get("Module Descriptive Title", ""), body, img])
    path = os.path.join(out_dir, "Unit cover page & Learning Objective.csv")
    write_csv(path, header, data)
    return path


def write_intro(rows, out_dir):
    """Introduction paragraphs fan into up to 5 slots (one per paragraph;
    extras merge into the last slot so nothing is dropped)."""
    header = ["Module Number", "Module Title 1", "Paragraph 1", "Paragraph 2",
              "Paragraph 3", "Paragraph 4", "Paragraph 5", "Image"]
    data = []
    for r in _sorted_rows(rows):
        text, img = _body_img(r, "Introduction / Topic Overview")
        if not (text or img):
            continue
        paras = [ln for ln in text.split("\n") if ln.strip()]
        if len(paras) > 5:
            paras = paras[:4] + ["\n".join(paras[4:])]
        paras += [""] * (5 - len(paras))
        data.append([_mod(r), _title(r)] + paras + [img])
    path = os.path.join(out_dir, "Introduction Topic Overview.csv")
    write_csv(path, header, data)
    return path


_SUBHEAD_RE = re.compile(r'^\d{1,2}\.\d{1,2}\b')


_KC_SLOTS = 5


def write_key_concepts(rows, out_dir):
    """Intro text before the first N.N sub-heading goes to 'Key concepts
    intro'; each N.N block splits into its Heading slot (the '3.3 Title'
    line) and its Paragraph slot (the text that follows the heading)."""
    header = ["Lesson No.", "Key concepts intro"]
    for i in range(1, _KC_SLOTS + 1):
        header += [f"Heading {i}", f"Paragraph {i}"]
    header += ["Image"]
    data = []
    for r in _sorted_rows(rows):
        text, img = _body_img(r, "Key Concepts")
        if not (text or img):
            continue
        intro, blocks = [], []          # block = [heading, [body lines]]
        for ln in text.split("\n"):
            if _SUBHEAD_RE.match(_bare(ln).strip()):
                blocks.append([ln.strip(), []])
            elif blocks:
                blocks[-1][1].append(ln)
            else:
                intro.append(ln)
        if len(blocks) > _KC_SLOTS:     # never drop content: merge overflow
            head, tail = blocks[:_KC_SLOTS - 1], blocks[_KC_SLOTS - 1:]
            merged_body = []
            for h, b in tail[1:]:
                merged_body += [h] + b
            blocks = head + [[tail[0][0], tail[0][1] + merged_body]]
        cells = []
        for h, b in blocks:
            cells += [h, "\n".join(b).strip()]
        cells += [""] * (_KC_SLOTS * 2 - len(cells))
        data.append([_mod(r), "\n".join(intro).strip()] + cells + [img])
    path = os.path.join(out_dir, "Key Concepts (Main Content).csv")
    write_csv(path, header, data)
    return path


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
    if _PA_REVIEW.match(line):
        return "success"
    for rx, b in _PA_BUCKETS:
        if rx.match(line):
            return b
    return None


def write_practical(rows, out_dir):
    header = ["Module Number", "Module Title 1", "Practical Title",
              "Acitvity Intro ", "Guide/ Steps", "Success Criteria",
              "Discussion", "Troubleshooting", "Rubric ", "Image"]
    data = []
    for r in _sorted_rows(rows):
        text, img = _body_img(r, "Practical Activity")
        if not (text or img):
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
            b = _pa_bucket(_bare(s))
            if b:
                cur = b
            buckets[cur].append(s)
        data.append([_mod(r), _title(r), title,
                     "\n".join(buckets["intro"]).strip(),
                     "\n".join(buckets["steps"]).strip(),
                     "\n".join(buckets["success"]).strip(),
                     "\n".join(buckets["discussion"]).strip(),
                     "\n".join(buckets["trouble"]).strip(),
                     "\n".join(buckets["rubric"]).strip(), img])
    path = os.path.join(out_dir, "Practical Activity (Hands-On).csv")
    write_csv(path, header, data)
    return path


# ---------- practice sheets (top paragraph + body) ----------

def _split_top(text, body_re):
    """Leading lines before the first body line (scenario/question/numbered)
    form the top paragraph; the rest is the body, verbatim."""
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


_GP_BODY = re.compile(r'^(\d{1,2}[.)]\s|scenario\b|situation\b)', re.I)
_IP_BODY = re.compile(r'^(\d{1,2}[.)]\s|[a-e][.)]\s|[A-E]\.\s)', re.I)
_KC_BODY = re.compile(r'^\d{1,2}[.)]?\s|^what\b|^which\b|^why\b|^who\b|^how\b',
                      re.I)


def write_guided(rows, out_dir):
    header = ["Module Number", "Module Title 1", "Top paragraph",
              "Practice Scenarios", "Image"]
    data = []
    for r in _sorted_rows(rows):
        text, img = _body_img(r, "Guided Practice")
        if not (text or img):
            continue
        top, body = _split_top(text, _GP_BODY)
        data.append([_mod(r), _title(r), top, body, img])
    path = os.path.join(out_dir, "Guided Practice (With Support).csv")
    write_csv(path, header, data)
    return path


def write_independent(rows, out_dir):
    header = ["Module Number", "Module Title 1", "Top paragraph",
              "Independent Practice ", "Image"]
    data = []
    for r in _sorted_rows(rows):
        text, img = _body_img(r, "Independent Practice")
        if not (text or img):
            continue
        top, body = _split_top(text, _IP_BODY)
        data.append([_mod(r), _title(r), top, body, img])
    path = os.path.join(out_dir, "Independent Practice.csv")
    write_csv(path, header, data)
    return path


def write_challenge(rows, out_dir):
    header = ["Module Number", "Module Title 1", "Section Content", "Image"]
    data = []
    for r in _sorted_rows(rows):
        body, img = _body_img(r, "Challenge / Higher-Order Thinking")
        if body or img:
            data.append([_mod(r), _title(r), body, img])
    path = os.path.join(out_dir, "Challenge Higher-Order Thinking.csv")
    write_csv(path, header, data)
    return path


def write_knowledge_check(rows, out_dir):
    header = ["Module Number", "Module Title 1", "Instruction", "Questions",
              "Image"]
    data = []
    for r in _sorted_rows(rows):
        text, img = _body_img(r, "Knowledge Check")
        if not (text or img):
            continue
        top, body = _split_top(text, _KC_BODY)
        data.append([_mod(r), _title(r), top, body, img])
    path = os.path.join(out_dir, "Knowledge Check (Quick Assessment).csv")
    write_csv(path, header, data)
    return path


# ---------- common mistakes ----------

_MISTAKE_RE = re.compile(r'^mistake\s*[:—–-]?\s*', re.I)
_TIP_RE = re.compile(r'^tip\s*[:—–-]?\s*', re.I)


def write_mistakes(rows, out_dir):
    """One row per Mistake/Tip pair. Text that carries no explicit tags is
    kept whole in the Mistake column so nothing is dropped."""
    header = ["Module Number", "Module Title 1", "Mistake ", "Tip", "Image"]
    data = []
    for r in _sorted_rows(rows):
        text, img = _body_img(r, "Common Mistakes & Tips")
        if not (text.strip() or img):
            continue
        pairs = [["", ""]]
        cur = None
        tagged = False
        for ln in text.split("\n"):
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
                pairs[-1][1] = s
                cur = 1
            elif cur is not None:
                pairs[-1][cur] = (pairs[-1][cur] + "\n" + s).strip()
            else:
                pairs[-1][0] = join_nonempty(pairs[-1][0], s)
        if not tagged and pairs == [["", ""]]:
            pairs = [[text.strip(), ""]]
        first = True
        for m, t in pairs:
            if m or t:
                # section-level images ride on the module's first row
                data.append([_mod(r), _title(r), m, t, img if first else ""])
                first = False
        if first and img:               # images but no mistake text
            data.append([_mod(r), _title(r), "", "", img])
    path = os.path.join(out_dir, "Common Mistakes & Tips.csv")
    write_csv(path, header, data)
    return path


def write_real_world(rows, out_dir):
    header = ["Module Number", "Module Title 1", "Intro/ paragraph", "Points",
              "Think about it ", "Image"]
    data = []
    for r in _sorted_rows(rows):
        text, img = _body_img(r, "Real-World Application")
        if not (text or img):
            continue
        # a trailing "Think about it"/question block splits off when marked
        think = ""
        m = re.search(r'^[\s•·●▪\t-]*think about it\b.*$', text, re.I | re.M)
        if m:
            text, think = text[:m.start()].rstrip(), text[m.start():].strip()
        data.append([_mod(r), _title(r), text, "", think, img])
    path = os.path.join(out_dir, "Real-World Application.csv")
    write_csv(path, header, data)
    return path


def write_summary(rows, out_dir):
    header = ["Module Number", "Module Title 1", "Summary Content ", "Image"]
    data = []
    for r in _sorted_rows(rows):
        body, img = _body_img(r, "Summary / Key Takeaways")
        if body or img:
            data.append([_mod(r), _title(r), body, img])
    path = os.path.join(out_dir, "Summary Key Takeaways.csv")
    write_csv(path, header, data)
    return path


# ---------- answers & solutions ----------

# sub-blocks inside the Answers & Solutions section, matched by name
_ANS_BLOCKS = [
    (re.compile(r'^guided practice\b.*[:]?', re.I),      "gp"),
    (re.compile(r'^independent practice\b.*[:]?', re.I), "ip"),
    (re.compile(r'^(challenge|higher-order|evaluate a scenario)\b.*', re.I), "hot"),
    (re.compile(r'^knowledge check\b.*[:]?', re.I),      "kc"),
]

_NUMLINE = re.compile(r'^(\d{1,2})[.)]\s*(.*)$')


def _fan(lines, slots):
    """Fan answer lines into numbered slots. Explicit '1.' numbering wins;
    unnumbered lines each take the next slot. Overflow merges into the last
    slot so nothing is dropped."""
    items = []
    for ln in lines:
        m = _NUMLINE.match(_bare(ln).strip())
        if m:
            items.append([m.group(2).strip()])
        elif items and re.match(r'^[a-e][).]\s', _bare(ln).strip(), re.I) \
                and not re.match(r'^[a-e][).]\s', _bare(items[-1][0]), re.I):
            # a)-e) option lines ride with the question above them -- unless
            # the previous item is itself a lettered answer (knowledge-check
            # style: each "b) ..." line is its own answer)
            items[-1].append(ln)
        elif items and _bare(ln).strip().lower().startswith(("reason", "fix", "explanation")):
            items[-1].append(ln)
        elif ln.strip():
            items.append([ln.strip()])
    flat = ["\n".join(x).strip() for x in items]
    if len(flat) > slots:
        flat = flat[:slots - 1] + ["\n".join(flat[slots - 1:])]
    flat += [""] * (slots - len(flat))
    return flat


def write_answers(rows, out_dir):
    header = (["Module Number", "Module Title 1"]
              + [f"Guided Practice {i}" for i in range(1, 11)]
              + [f"Independent Practice {i}" for i in range(1, 11)]
              + [f"Higher-Order Thinking {i}" for i in range(1, 6)]
              + [f"Knowledge Check {i}" for i in range(1, 11)]
              + ["Image"])
    data = []
    for r in _sorted_rows(rows):
        text, img = _body_img(r, "Answers & Solutions")
        if not (text.strip() or img):
            continue
        buckets = {"gp": [], "ip": [], "hot": [], "kc": [], "": []}
        cur = ""
        for ln in text.split("\n"):
            s = ln.strip()
            if not s:
                continue
            hit = None
            for rx, b in _ANS_BLOCKS:
                if rx.match(_bare(s)) and len(s) <= 60:
                    hit = b
                    break
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
        data.append([_mod(r), _title(r)]
                    + _fan(buckets["gp"], 10) + _fan(buckets["ip"], 10)
                    + _fan(buckets["hot"], 5) + _fan(buckets["kc"], 10)
                    + [img])
    path = os.path.join(out_dir, "Answer Solution.csv")
    write_csv(path, header, data)
    return path


# ---------- suggested interactive moments ----------

_MOMENT_START = re.compile(r'^(moment\s*\d+|location\s*[:—–-]|where it fits\s*[:—–-])',
                           re.I)
_SECTION_LINE = re.compile(r'^(location|where it fits)\s*[:—–-]\s*(.*)$', re.I)


def write_moments(rows, out_dir):
    """One row per suggested moment. Groups start at a 'Moment N' or
    'Location:'/'Where it fits:' line; the Section column carries the
    location, the Suggestion column the whole group verbatim."""
    header = ["Module title ", "Section ", "Suggestion", "Image"]
    data = []
    for r in _sorted_rows(rows):
        text, img = _body_img(r, "Suggested Interactive Moments")
        if not (text.strip() or img):
            continue
        groups = []
        for ln in text.split("\n"):
            s = ln.strip()
            if not s:
                continue
            # a new group starts at "Moment N", or at a Location line when
            # the current group already has one (bullet-list style docs)
            is_loc = bool(_SECTION_LINE.match(_bare(s)))
            starts_new = bool(re.match(r'^moment\s*\d+', _bare(s), re.I)) or (
                is_loc and groups and any(_SECTION_LINE.match(_bare(x))
                                          for x in groups[-1]))
            if starts_new or not groups:
                groups.append([s])
            else:
                groups[-1].append(s)
        mod_title = join_nonempty(_mod(r), _title(r), sep=" ")
        first = True
        for g in groups:
            section = ""
            for x in g:
                m = _SECTION_LINE.match(_bare(x))
                if m:
                    section = m.group(2).strip()
                    break
            data.append([mod_title, section, "\n".join(g).strip(),
                         img if first else ""])
            first = False
        if first and img:
            data.append([mod_title, "", "", img])
    path = os.path.join(out_dir, "Suggested Interactive Moments.csv")
    write_csv(path, header, data)
    return path


def write_all_it(rows, out_dir):
    os.makedirs(out_dir, exist_ok=True)
    paths = [
        write_full_extract(rows, out_dir),
        write_cover(rows, out_dir),
        write_intro(rows, out_dir),
        write_key_concepts(rows, out_dir),
        write_practical(rows, out_dir),
        write_guided(rows, out_dir),
        write_independent(rows, out_dir),
        write_challenge(rows, out_dir),
        write_knowledge_check(rows, out_dir),
        write_mistakes(rows, out_dir),
        write_real_world(rows, out_dir),
        write_summary(rows, out_dir),
        write_answers(rows, out_dir),
        write_moments(rows, out_dir),
    ]
    return paths
