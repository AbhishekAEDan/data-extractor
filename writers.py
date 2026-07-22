"""
writers.py

Writes the parsed rows to CSV files. Output layouts mirror the official
"Updated Spread Sheets" workbooks (one CSV per sheet, headers copied exactly
from the templates -- including their spacing/typo quirks, e.g. "Vocab
Vault  1", "Unit Number " and "Porject Title"):

  main doc sheets (one row per sub-unit lesson):
    Unit Cover Page.csv
    Vocab Vault.csv
    Mistake Spotter.csv
    Mini Practice.csv
    Check Your Understanding.csv
    In Sub Unit Portfolio Proj.csv
    Sub Unit Recap.csv
    (Main Lesson sheet is intentionally NOT produced -- filled by a human)

  unit-level sheets:
    Unit Introduction.csv          (per Unit N Introduction doc)
    End of Unit Recap.csv          (per End of Unit N doc)
    End of Unit Portfolio.csv      (per End of Unit N Portfolio doc)

  extras:
    full_extract.csv                  (debug: every parsed field, one row
                                       per document; formerly wide.csv --
                                       comprehension fields live here too)

All values are verbatim slices from the parser -- no rewriting happens here.

Author: AbhishekAEDan
"""
__author__ = "AbhishekAEDan"

import csv
import os
import re

from parser_core import ORDER, _match_portfolio_section, strip_bullet


# ---------- helpers ----------

def order_cols(rows):
    cols = [c for c in ORDER if any(c in r for r in rows)]
    for r in rows:
        for k in r:
            if k not in cols:
                cols.append(k)
    return cols


def write_csv(path, header, data_rows):
    """Write one template sheet. The first column is always Unit Number:
    it is wrapped as ="3.10" so Excel keeps it as text -- otherwise Excel
    coerces 3.10 to the number 3.1 and the trailing zero is lost."""
    with open(path, "w", newline="", encoding="utf-8-sig") as fh:
        w = csv.writer(fh)
        w.writerow(header)
        for r in data_rows:
            r = list(r)
            if r and r[0] not in ("", None):
                r[0] = f'="{r[0]}"'
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


# ---------- debug sheet ----------

def write_full_extract(rows, out_dir):
    cols = order_cols(rows)
    path = os.path.join(out_dir, "full_extract.csv")
    with open(path, "w", newline="", encoding="utf-8-sig") as fh:
        w = csv.DictWriter(fh, fieldnames=cols)
        w.writeheader()
        for r in rows:
            w.writerow({k: r.get(k, "") for k in cols})
    return path


# ---------- main doc sheets (per sub-unit lesson) ----------

def write_unit_cover(rows, out_dir):
    header = ["Unit Number", "Sub Unit Title", "My Goal",
              "By the End of This Lesson", "Bridge Back"]
    data = [[r.get("Unit Number", ""), r.get("Lesson Title", ""),
             r.get("My Goal", ""), r.get("Lesson Objectives", ""),
             r.get("Bridge Back", "")]
            for r in _rows_of_type(rows, "lesson")]
    path = os.path.join(out_dir, "Unit Cover Page.csv")
    write_csv(path, header, data)
    return path


_VOCAB_ENTRY_RE = re.compile(r'^([^:]{1,60}?):\s*(.*)$')


def write_vocab_vault(rows, out_dir):
    # template header has a double space: "Vocab Vault  1"
    header = ["Unit Number"] + [f"Vocab Vault  {i}" for i in range(1, 11)]
    data = []
    for r in _rows_of_type(rows, "lesson"):
        entries = []                 # each entry = list of verbatim lines
        for ln in (r.get("Vocab Vault", "") or "").split("\n"):
            if _VOCAB_ENTRY_RE.match(ln.strip()):
                entries.append([ln.strip()])
            elif entries and ln.strip():
                entries[-1].append(ln.strip())
            elif ln.strip():
                entries.append([ln.strip()])
        if not entries:
            continue
        cells = ["\n".join(e).strip() for e in entries]
        if len(cells) > 10:          # never drop entries: merge extras into #10
            cells = cells[:9] + ["\n".join(cells[9:])]
        cells += [""] * (10 - len(cells))
        data.append([r.get("Unit Number", "")] + cells)
    path = os.path.join(out_dir, "Vocab Vault.csv")
    write_csv(path, header, data)
    return path


# the four fixed Mistake Spotter tags used in the lesson docs
_MS_TAGS = [
    ("common mistake",          "Common mistake"),
    ("what this may look like", "What this may look like"),
    ("why this needs revision", "Why this needs revision"),
    ("how to improve it",       "How to improve it"),
]
_MS_COLS = [c for _, c in _MS_TAGS]


def _ms_tag(line):
    low = line.lower()
    for key, col in _MS_TAGS:
        if low.startswith(key):
            rest = line[len(key):].lstrip(" :–—-").strip()
            return col, rest
    return None


def write_mistake_spotter(rows, out_dir):
    """One row per mistake group. A repeated 'Common mistake' tag starts a
    new group (some lessons spot more than one mistake)."""
    header = ["Unit Number"] + _MS_COLS
    data = []
    for r in _rows_of_type(rows, "lesson"):
        text = r.get("Mistake Spotter", "") or ""
        if not text.strip():
            continue
        groups = [{}]
        cur = None
        for ln in text.split("\n"):
            s = ln.strip()
            if not s:
                continue
            hit = _ms_tag(s)
            if hit:
                col, rest = hit
                if col in groups[-1]:            # tag repeats -> new group
                    groups.append({})
                cur = col
                groups[-1][col] = rest
            elif cur:
                g = groups[-1]
                g[cur] = (g[cur] + "\n" + s).strip() if g[cur] else s
            else:                                # content before any tag
                g = groups[-1]
                c = _MS_COLS[0]
                g[c] = (g.get(c, "") + "\n" + s).strip() if g.get(c) else s
                cur = c
        for g in groups:
            if any(g.get(c) for c in _MS_COLS):
                data.append([r.get("Unit Number", "")]
                            + [g.get(c, "") for c in _MS_COLS])
    path = os.path.join(out_dir, "Mistake Spotter.csv")
    write_csv(path, header, data)
    return path


def write_mini_practice(rows, out_dir):
    # template quirk: "Unit Number " has a trailing space on this sheet
    header = ["Unit Number ", "Mini Practice Title", "Top Paragraph",
              "Top What to look for", "Part A", "Part B", "Part C"]
    data = []
    for r in _rows_of_type(rows, "lesson"):
        # the template has no Answer Key column -- keep the key verbatim at
        # the end of Part C so nothing is dropped
        part_c = r.get("Practice C", "")
        ak = r.get("Mini Practice Answer Key", "")
        if ak:
            part_c = join_nonempty(part_c, "Answer Key:\n" + ak)
        if not any([r.get("Mini Practice Title"), r.get("Mini Practice Intro"),
                    r.get("Practice A"), r.get("Practice B"), part_c]):
            continue
        data.append([
            r.get("Unit Number", ""),
            r.get("Mini Practice Title", ""),
            r.get("Mini Practice Intro", ""),
            r.get("What to look for", ""),
            r.get("Practice A", ""), r.get("Practice B", ""), part_c,
        ])
    path = os.path.join(out_dir, "Mini Practice.csv")
    write_csv(path, header, data)
    return path


_NUMBERED_RE = re.compile(r'^(\d{1,2})[.)]\s*(.*)$')


def _fan_numbered(text, slots=10):
    """Split verbatim text into (preface, [item1..itemN]). Lines starting
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
    """One row per doc that carries a 'Check for Understanding' block.
    Numbered lines fan into Question 1-10, the Answer Key into Answer 1-10.
    The template has no Preface column, so any lead-in text stays verbatim
    at the top of Question 1."""
    header = (["Unit Number"]
              + [f"Question {i}" for i in range(1, 11)]
              + [f"Answer {i}" for i in range(1, 11)])
    data = []
    for r in _rows_of_type(rows, "lesson", "end_of_unit"):
        cfu = r.get("Check for Understanding", "")
        ak = r.get("Check for Understanding Answer Key", "")
        if not (cfu or ak):
            continue
        # some docs keep the answers inside the block after a
        # "Check your answers:" line instead of a bold Answer Key section
        if not ak:
            m = re.search(r'^\s*check your answers\b.*$', cfu, re.I | re.M)
            if m:
                cfu, ak = cfu[:m.start()].rstrip(), cfu[m.end():].strip()
        preface, qs = _fan_numbered(cfu)
        ak_preface, ans = _fan_numbered(ak)
        # unnumbered blocks can't tell a lead-in instruction from question 1;
        # when there is exactly one more question than answers, the first
        # "question" is really the lead-in
        nq = sum(1 for q in qs if q)
        na = sum(1 for a in ans if a)
        unnumbered = not any(_NUMBERED_RE.match(l.strip())
                             for l in (cfu or "").split("\n"))
        if unnumbered and not preface and na and nq == na + 1:
            preface, qs = qs[0], qs[1:] + [""]
            nq -= 1
        # a single answer means one question (e.g. a numbered word bank is
        # options for one question, not many questions) -- and vice versa
        if na == 1 and nq > 1:
            qs = [join_nonempty(preface, *qs)] + [""] * 9
            preface = ""
        elif nq == 1 and na > 1:
            ans = [join_nonempty(*ans)] + [""] * 9
        # an answer key with no numbered lines still belongs in Answer 1
        if ak_preface and not any(ans):
            ans[0] = ak_preface
        # no Preface column in the template -- lead-in rides on Question 1
        if preface:
            qs[0] = join_nonempty(preface, qs[0])
        data.append([r.get("Unit Number", "")] + qs + ans)
    path = os.path.join(out_dir, "Check Your Understanding.csv")
    write_csv(path, header, data)
    return path


def _parse_portfolio_text(text):
    """Run the portfolio section matcher over a lesson's verbatim
    'Portfolio Project' block. Returns {column: content}."""
    out = {}
    cur, buf = None, []

    def flush():
        if cur and buf:
            content = "\n".join(buf).strip()
            if content and not out.get(cur):
                out[cur] = content

    for raw in (text or "").split("\n"):
        s = strip_bullet(raw.strip())
        if not s:
            continue
        hit = _match_portfolio_section(s)
        if hit:
            flush()
            cur, buf = hit[0], ([hit[1]] if hit[1] else [])
        elif cur:
            buf.append(s)
        else:
            # text before any recognised label (usually the project title
            # line) -- keep it, never drop content
            cur, buf = "Project Title", [s]
    flush()
    return out


def write_sub_unit_portfolio(rows, out_dir):
    # header copied exactly from the template sheet (incl. the unit-1-specific
    # "Example: My Netiquette Pledge" column title)
    header = ["Unit Number", "Project Title", "Project Name", "Project link",
              "Portfolio Link", "What you will create", "Materials",
              "Numbered Steps", "Templates/organisers + how to fill",
              "Example: My Netiquette Pledge", "Common mistakes + fixes",
              "Submission checklist", "Optional extension",
              "Motivation + reflection", "End-of-Unit Inventory"]
    field = ["Project Title", "Project Name", "Project Link", "Portfolio Link",
             "What You Will Create", "Materials", "Steps",
             "Template How To Fill", "Example", "Common Mistakes And Fixes",
             "Submission Checklist", "Optional Extension",
             "Motivation Reflection", "End-of-Unit Inventory"]
    data = []
    for r in _rows_of_type(rows, "lesson"):
        p = _parse_portfolio_text(r.get("Portfolio Project", ""))
        if not p:
            continue
        # the template has no Purpose column -- keep the purpose text
        # verbatim with 'What you will create' so nothing is dropped
        if p.get("Purpose"):
            p["What You Will Create"] = join_nonempty(
                "Purpose: " + p["Purpose"], p.get("What You Will Create", ""))
        data.append([r.get("Unit Number", "")] + [p.get(f, "") for f in field])
    path = os.path.join(out_dir, "In Sub Unit Portfolio Proj.csv")
    write_csv(path, header, data)
    return path


def write_sub_unit_recap(rows, out_dir):
    # template quirks: trailing spaces on the checklist/recap column titles
    header = ["Unit Number", "My Learning Checklist ", "Quick Recap ", "Reflect"]
    data = [[r.get("Unit Number", ""), r.get("My Learning Checklist", ""),
             r.get("Quick Recap", ""), r.get("Reflect", "")]
            for r in _rows_of_type(rows, "lesson")
            if any([r.get("My Learning Checklist"), r.get("Quick Recap"),
                    r.get("Reflect")])]
    path = os.path.join(out_dir, "Sub Unit Recap.csv")
    write_csv(path, header, data)
    return path


# ---------- unit-level sheets ----------

def write_unit_intro(rows, out_dir):
    # template quirk: "Unit Number " has a trailing space on this sheet
    header = ["Unit Number ", "Title", "Big Idea", "Essential Question",
              "Role Call: Master Communicators", "At the End of Unit"]
    data = [[r.get("Unit Number", ""), r.get("Unit Title", ""),
             r.get("Big Idea", ""), r.get("Essential Question", ""),
             r.get("Role Call", ""), r.get("At End Of Unit Checklist", "")]
            for r in _rows_of_type(rows, "unit_intro")]
    path = os.path.join(out_dir, "Unit Introduction.csv")
    write_csv(path, header, data)
    return path


def write_end_of_unit_recap(rows, out_dir):
    # header copied exactly from the template (incl. "Unit 1" wording and the
    # trailing space on "Reflect ")
    header = ["Unit Number", "My Unit 1 Learning Checklist", "Quick Recap",
              "Reflect "]
    data = [[r.get("Unit Number", ""), r.get("Unit Learning Checklist", ""),
             r.get("Unit Quick Recap", ""), r.get("Unit Reflect", "")]
            for r in _rows_of_type(rows, "end_of_unit")]
    path = os.path.join(out_dir, "End of Unit Recap.csv")
    write_csv(path, header, data)
    return path


def write_end_of_unit_portfolio(rows, out_dir):
    # header copied exactly from the template sheet, quirks and all
    # ("Porject Title", "Project Name " trailing space, the long inventory
    # column title, unit-1-specific "Example Central Community Newsletter")
    header = ["Unit Number", "Porject Title", "Project Name ", "Project Link",
              "Portfolio Link", "What You Will Create", "Materials",
              "Numbered Steps", "Templates/Organisers + How to Fill ",
              "Example Central Community Newsletter", "Common Mistakes + Fixes",
              "Submission Checklist", "Optional Extension",
              "End-of-Unit Inventory By th[11.1]e end of this unit, your "
              "portfolio should contain: Unit 1 Artefact Pack",
              "Motivation + Reflection Skill Badge Earned", "Reflect"]
    field = ["Project Title", "Project Name", "Project Link", "Portfolio Link",
             "What You Will Create", "Materials", "Steps",
             "Template How To Fill", "Example", "Common Mistakes And Fixes",
             "Submission Checklist", "Optional Extension",
             "End-of-Unit Inventory", "Motivation Reflection", "Unit Reflect"]
    data = []
    for r in _rows_of_type(rows, "end_of_unit_portfolio"):
        data.append([r.get("Unit Number", "")] + [r.get(f, "") for f in field])
    path = os.path.join(out_dir, "End of Unit Portfolio.csv")
    write_csv(path, header, data)
    return path


# NOTE: the comprehension sheet was removed (2026-07-22, user request) --
# comprehension text belongs to the human-filled Main Lesson sheet. The
# parsed fields (Comprehension Title / Before Reading / Main Text / ...)
# still land in full_extract.csv, so nothing is lost.

def write_all(rows, out_dir):
    os.makedirs(out_dir, exist_ok=True)
    paths = [
        write_full_extract(rows, out_dir),
        write_unit_cover(rows, out_dir),
        write_vocab_vault(rows, out_dir),
        write_mistake_spotter(rows, out_dir),
        write_mini_practice(rows, out_dir),
        write_check_understanding(rows, out_dir),
        write_sub_unit_portfolio(rows, out_dir),
        write_sub_unit_recap(rows, out_dir),
        write_unit_intro(rows, out_dir),
        write_end_of_unit_recap(rows, out_dir),
        write_end_of_unit_portfolio(rows, out_dir),
    ]
    return paths
