"""
judge.py

LLM "judge" -- used ONLY to map unknown bold section labels onto canonical
columns. The judge never sees or rewrites document content, so extraction
stays verbatim regardless of which engine is used.

Engines:
  - ollama : local model via http://localhost:11434 (default qwen3:4b-instruct)
  - gemini : gemini flash-lite via google-genai SDK

Author: AbhishekAEDan
"""
__author__ = "AbhishekAEDan"

import json
import re
import urllib.request

from parser_core import CANONICAL_COLUMNS

OLLAMA_URL = "http://localhost:11434"

# Deterministic label rules (user-confirmed placements, 2026-07-20).
# Checked BEFORE the AI judge -- these headings always land in the same
# column, no model call needed. Targets may be any wide-row column, not
# just CANONICAL_COLUMNS (e.g. the fanned Mini Practice fields).
HARD_RULES = [
    (re.compile(r'word class crew', re.I),            "Bridge Back"),
    (re.compile(r'^student page copy', re.I),         "Interactive / Simulation"),
    (re.compile(r'descriptive wheel', re.I),          "Vocab Vault"),
    (re.compile(r'labelled diagram', re.I),           "Main Lesson"),
    (re.compile(r'two-?column chart', re.I),          "Mini Practice Intro"),
    (re.compile(r'missing dictionary incident', re.I), "Mini Practice Intro"),
    # a genre-titled passage heading ("Drama Josiah's Promise", "Poem ...")
    # is the comprehension passage title
    (re.compile(r'^(drama|poem|prose|play)\b', re.I), "Comprehension Title"),
]


def apply_hard_rules(labels):
    """Split labels into ({label: target} for rule hits, [remaining])."""
    mapped, rest = {}, []
    for l in labels:
        for rx, target in HARD_RULES:
            if rx.search(l):
                mapped[l] = target
                break
        else:
            rest.append(l)
    return mapped, rest

JUDGE_PROMPT = """You are labeling section headings from school lesson documents.

Canonical section names:
{canon}

For each heading below, answer with the canonical name it belongs to,
or "KEEP" if it is a genuinely new/different section, or "IGNORE" if it is
not a real section (e.g. page furniture, decoration).

Headings:
{headings}

Return ONLY a JSON object mapping each heading exactly as given to its answer.
No commentary."""


def _build_prompt(labels):
    return JUDGE_PROMPT.format(
        canon="\n".join(f"- {c}" for c in CANONICAL_COLUMNS),
        headings="\n".join(f'- "{l}"' for l in labels),
    )


def _parse_judge_json(raw, labels):
    raw = raw.strip()
    m = re.search(r'\{.*\}', raw, re.S)
    if not m:
        return {}
    try:
        data = json.loads(m.group(0))
    except Exception:
        return {}
    out = {}
    for label in labels:
        ans = str(data.get(label, "KEEP")).strip()
        if ans == "IGNORE":
            out[label] = "IGNORE"
        elif ans in CANONICAL_COLUMNS:
            out[label] = ans
        else:
            out[label] = None  # KEEP -> keep as its own column
    return out


def judge_ollama(labels, model):
    body = json.dumps({
        "model": model,
        "prompt": _build_prompt(labels),
        "stream": False,
        "options": {"temperature": 0},
        "format": "json",
    }).encode()
    req = urllib.request.Request(f"{OLLAMA_URL}/api/generate", data=body,
                                 headers={"Content-Type": "application/json"})
    with urllib.request.urlopen(req, timeout=300) as resp:
        data = json.loads(resp.read().decode())
    return _parse_judge_json(data.get("response", ""), labels)


def judge_gemini(labels, api_key, model="gemini-3.1-flash-lite"):
    from google import genai
    from google.genai import types
    client = genai.Client(api_key=api_key)
    resp = client.models.generate_content(
        model=model,
        contents=_build_prompt(labels),
        config=types.GenerateContentConfig(
            response_mime_type="application/json", temperature=0),
    )
    return _parse_judge_json(resp.text or "", labels)


def run_judge(labels, engine, ollama_model=None, gemini_key=None, gemini_model=None):
    """Returns {label: canonical|'IGNORE'|None}. On any failure returns {}
    (labels are then kept as their own columns -- nothing lost)."""
    labels = sorted(labels)
    if not labels:
        return {}
    mapped, rest = apply_hard_rules(labels)
    if not rest:
        return mapped
    try:
        if engine == "ollama":
            ai = judge_ollama(rest, ollama_model or "qwen3:4b-instruct")
        else:
            ai = judge_gemini(rest, gemini_key, gemini_model or "gemini-3.1-flash-lite")
    except Exception as e:
        print(f"  [!] Judge failed ({e}); unknown headings kept as their own columns.")
        ai = {}
    mapped.update(ai)
    return mapped
