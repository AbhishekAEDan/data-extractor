#!/usr/bin/env python3
"""
Lesson Table Extractor -- interactive console app.

Put .docx lesson files in ./documents, run run.bat (or `python main.py`),
pick an engine (Qwen via Ollama, or Gemini) for the judge step, run the
extraction. Output CSVs land in ./output.

Extraction itself is a deterministic verbatim parser; the selected AI engine
is ONLY used to classify unrecognised section headings -- it never rewrites
document text.

Author: AbhishekAEDan
"""
__author__ = "AbhishekAEDan"

import glob
import json
import os
import subprocess
import sys
import time

BASE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, BASE)

DOCS_DIR = os.path.join(BASE, "documents")
OUT_DIR = os.path.join(BASE, "output")
CONFIG_PATH = os.path.join(BASE, "config.json")
ENV_PATH = os.path.join(BASE, ".env")

DEFAULT_CONFIG = {
    "engine": "ollama",                       # "ollama" | "gemini"
    "ollama_model": "qwen3:4b-instruct",
    "gemini_model": "gemini-3.1-flash-lite",
}


# ---------- config / env ----------

def load_config():
    cfg = dict(DEFAULT_CONFIG)
    if os.path.exists(CONFIG_PATH):
        try:
            with open(CONFIG_PATH, "r", encoding="utf-8") as f:
                cfg.update(json.load(f))
        except Exception:
            pass
    return cfg


def save_config(cfg):
    with open(CONFIG_PATH, "w", encoding="utf-8") as f:
        json.dump(cfg, f, indent=2)


def load_api_key():
    if os.path.exists(ENV_PATH):
        with open(ENV_PATH, "r", encoding="utf-8") as f:
            for line in f:
                if line.strip().startswith("GEMINI_API_KEY="):
                    return line.strip().split("=", 1)[1]
    return os.getenv("GEMINI_API_KEY", "")


def save_api_key(key):
    with open(ENV_PATH, "w", encoding="utf-8") as f:
        f.write(f"GEMINI_API_KEY={key}\n")


# ---------- ui helpers ----------

APP_NAME = "Brent & Co. Data Extractor"
AUTHOR = "AbhishekAEDan"

LOGO_BIG = r"""
 ____                 _      ___      ____
| __ ) _ __ ___ _ __ | |_   ( _ )    / ___|___
|  _ \| '__/ _ \ '_ \| __|  / _ \/\ | |   / _ \ _
| |_) | | |  __/ | | | |_  | (_>  < | |__| (_) |_|
|____/|_|  \___|_| |_|\__|  \___/\/  \____\___/
        D A T A   E X T R A C T O R
"""

LOGO_SMALL = r"""
+------------------------------------+
|   BRENT & CO.  DATA EXTRACTOR      |
+------------------------------------+
"""


def term_width():
    try:
        import shutil
        return shutil.get_terminal_size().columns
    except Exception:
        return 80


def print_logo(animate=False):
    """Logo scales to the window: big art centered on wide consoles,
    compact box on narrow ones."""
    w = term_width()
    logo = LOGO_BIG if w >= 60 else LOGO_SMALL
    pad = " " * max((w - max(len(l) for l in logo.splitlines())) // 2, 0)
    for line in logo.splitlines():
        print(pad + line)
        if animate:
            time.sleep(0.05)
    # small maker credit, centered under the title graphic
    from updater import __version__
    credit = f"made by {AUTHOR}  .  v{__version__}"
    print(" " * max((w - len(credit)) // 2, 0) + credit)


def type_out(msg, delay=0.012, center=True):
    w = term_width()
    pad = " " * max((w - len(msg)) // 2, 0) if center else ""
    sys.stdout.write(pad)
    for ch in msg:
        sys.stdout.write(ch)
        sys.stdout.flush()
        time.sleep(delay)
    print()


def intro_animation():
    clear()
    print_logo(animate=True)
    w = term_width()
    bar = "=" * min(48, w - 4)
    print(" " * max((w - len(bar)) // 2, 0) + bar)
    type_out("verbatim . fast . no hallucination")
    print()
    time.sleep(0.3)


def progress_bar(i, n, width=28):
    done = int(width * i / max(n, 1))
    return "[" + "#" * done + "-" * (width - done) + f"] {i}/{n}"


class Spinner:
    """Console spinner for blocking waits (AI judge, downloads)."""
    FRAMES = "|/-\\"

    def __init__(self, label):
        self.label = label
        self._stop = False
        self._t = None

    def __enter__(self):
        import threading
        def spin():
            k = 0
            while not self._stop:
                sys.stdout.write(f"\r  {self.FRAMES[k % 4]} {self.label} ")
                sys.stdout.flush()
                k += 1
                time.sleep(0.12)
        self._t = threading.Thread(target=spin, daemon=True)
        self._t.start()
        return self

    def __exit__(self, *a):
        self._stop = True
        if self._t:
            self._t.join(timeout=1)
        sys.stdout.write("\r" + " " * (len(self.label) + 8) + "\r")
        sys.stdout.flush()


def clear():
    os.system("cls" if os.name == "nt" else "clear")


def pause():
    input("\n  Press Enter to go back...")


def find_docx(sources):
    """Collect .docx from files/folders (recursive -- Term/Unit subfolders ok)."""
    out = []
    for src in sources:
        if os.path.isfile(src) and src.lower().endswith(".docx"):
            out.append(src)
        elif os.path.isdir(src):
            out.extend(glob.glob(os.path.join(src, "**", "*.docx"), recursive=True))
    out = [f for f in out if not os.path.basename(f).startswith("~$")]
    return sorted(dict.fromkeys(out))


def banner(cfg, key, sources):
    n = len(find_docx(sources))
    if sources == [DOCS_DIR]:
        docs_msg = f"{n} .docx found in documents\\ (subfolders included)"
    else:
        docs_msg = f"{n} .docx from dragged-in path(s)"
    engine = cfg["engine"]
    model = cfg["ollama_model"] if engine == "ollama" else cfg["gemini_model"]
    print_logo()
    w = min(62, term_width() - 2)
    print("=" * w)
    print(f"  Documents : {docs_msg}")
    print(f"  Engine    : {engine.upper()}  ({model})")
    print(f"  Gemini key: {'set (...' + key[-6:] + ')' if key else 'NOT SET'}")
    print(f"  Output    : {OUT_DIR}")
    print("=" * w)


# ---------- menu actions ----------

def menu_engine(cfg):
    while True:
        clear()
        print("--- Select engine (used only to judge odd headings) ---\n")
        print("  1) Ollama / Qwen (local, private, free)")
        print("  2) Gemini API (smarter, needs key + internet)")
        print("  3) Change Ollama model name  [current: %s]" % cfg["ollama_model"])
        print("  4) Change Gemini model name  [current: %s]" % cfg["gemini_model"])
        print("  0) Back")
        c = input("\n  Choice: ").strip()
        if c == "1":
            cfg["engine"] = "ollama"
            save_config(cfg)
            print("\n  Engine set to OLLAMA.")
            time.sleep(1)
            return
        if c == "2":
            cfg["engine"] = "gemini"
            save_config(cfg)
            print("\n  Engine set to GEMINI.")
            time.sleep(1)
            return
        if c == "3":
            m = input("  Ollama model (e.g. qwen3:4b-instruct, hermes3:8b): ").strip()
            if m:
                cfg["ollama_model"] = m
                save_config(cfg)
        elif c == "4":
            m = input("  Gemini model (e.g. gemini-3.1-flash-lite): ").strip()
            if m:
                cfg["gemini_model"] = m
                save_config(cfg)
        elif c == "0":
            return


def menu_api_key():
    clear()
    print("--- Change Gemini API key ---\n")
    cur = load_api_key()
    print(f"  Current: {'...' + cur[-6:] if cur else '(none)'}")
    key = input("  New key (blank = cancel): ").strip()
    if key:
        save_api_key(key)
        print("  Saved to .env")
        time.sleep(1)


def run_checks(cfg, sources, verbose=True):
    """Returns True when the selected engine + parser can run."""
    from checks import (check_python_docx, check_genai, check_ollama_server,
                        check_ollama_model, check_gemini_key)
    results = []
    results.append(check_python_docx())
    n = len(find_docx(sources))
    results.append((n > 0, f"{n} .docx file(s) ready" if n else
                    "no .docx files found (documents\\ or dragged path)"))
    if cfg["engine"] == "ollama":
        ok_srv, msg = check_ollama_server()
        results.append((ok_srv, msg))
        if ok_srv:
            results.append(check_ollama_model(cfg["ollama_model"]))
    else:
        results.append(check_genai())
        results.append(check_gemini_key(load_api_key()))

    all_ok = all(ok for ok, _ in results)
    if verbose:
        print()
        for ok, msg in results:
            print(f"  [{'OK ' if ok else 'FAIL'}] {msg}")
        print()
    return all_ok


def try_start_ollama():
    """Best-effort: start the Ollama app if the server is down."""
    from checks import check_ollama_server
    if check_ollama_server()[0]:
        return True
    print("  Ollama not running -- attempting to start it...")
    try:
        subprocess.Popen(["ollama", "serve"],
                         stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    except FileNotFoundError:
        print("  'ollama' command not found. Install/start Ollama manually.")
        return False
    for _ in range(15):
        time.sleep(1)
        if check_ollama_server()[0]:
            print("  Ollama started.")
            return True
    print("  Ollama did not come up in time.")
    return False


def handle_unknown_headings(cfg, label_files):
    """Show unknown headings readably, let the user decide what happens.
    Returns a mapping for remap_unknown()."""
    from judge import run_judge
    from logger import log
    labels = sorted(label_files)
    print("\n" + "-" * 58)
    print(f"  {len(labels)} UNKNOWN HEADING(S) FOUND")
    print("  (sections in the documents that don't match any known column)")
    print("-" * 58)
    for i, lb in enumerate(labels, 1):
        fs = label_files[lb]
        shown = ", ".join(fs[:3]) + (f" +{len(fs)-3} more" if len(fs) > 3 else "")
        print(f"   {i}. \"{lb}\"")
        print(f"      found in: {shown}")
    print("-" * 58)
    print("\n  What should happen with these?")
    print("   1) Ask AI to decide which known column each one belongs to")
    print("   2) Ignore them -- keep known columns only")
    print("      (their text still saved as extra columns in wide.csv)")
    choice = input("\n  Choice [1/2]: ").strip()
    log(f"unknown headings: {labels}; user choice={choice}")

    if choice != "1":
        print("  Skipping AI -- unknown headings kept as extra columns in wide.csv.")
        return {}   # remap keeps them as their own visible columns

    print("\n  Which AI decides?")
    print(f"   1) Ollama local  ({cfg['ollama_model']})")
    print(f"   2) Gemini API    ({cfg['gemini_model']})")
    eng = "gemini" if input("\n  Choice [1/2]: ").strip() == "2" else "ollama"
    print()
    log(f"judge engine={eng}")
    with Spinner(f"Asking {eng.upper()} to classify {len(labels)} heading(s)..."):
        mapping = run_judge(set(labels), eng,
                            ollama_model=cfg["ollama_model"],
                            gemini_key=load_api_key(),
                            gemini_model=cfg["gemini_model"])
    print("\n  AI decisions:")
    for lb in labels:
        target = mapping.get(lb)
        if target == "IGNORE":
            verdict = "AI says not real content -> kept as own column anyway (nothing discarded)"
        elif target:
            verdict = f"goes into column '{target}'"
        else:
            verdict = "new section -> kept as its own column"
        print(f"   - \"{lb}\"")
        print(f"       {verdict}")
        log(f"judge: '{lb}' -> {target}")
    return mapping


def run_extraction(cfg, sources):
    clear()
    print("--- Run extraction ---")
    from bootstrap import auto_setup
    if not auto_setup(cfg):
        print("  Prerequisites could not be installed. See messages above.")
        pause()
        return
    if not run_checks(cfg, sources):
        print("  Fix the FAIL items above, then retry.")
        pause()
        return

    from parser_core import parse_docx, remap_unknown
    from writers import write_all

    files = find_docx(sources)
    root = sources[0] if len(sources) == 1 and os.path.isdir(sources[0]) else None
    print(f"  Parsing {len(files)} document(s)...\n")
    t0 = time.time()

    from logger import log
    rows, label_files = [], {}   # label -> [filenames it appears in]
    for i, path in enumerate(files, 1):
        name = os.path.relpath(path, root) if root else os.path.basename(path)
        try:
            found = set()
            row = parse_docx(path, found)
            row["_folder"] = os.path.basename(os.path.dirname(path))
            rows.append(row)
            for lb in found:
                label_files.setdefault(lb, []).append(os.path.basename(path))
            flags = []
            if row.get("_misc"):
                flags.append("_misc text")
            if row.get("has_image") == "yes":
                flags.append("image")
            extra = f"   ({', '.join(flags)})" if flags else ""
            print(f"  {progress_bar(i, len(files))}  {name}{extra}")
            log(f"parsed {name}: {len(row)} fields, unknown={sorted(found)}")
        except Exception as e:
            print(f"  {progress_bar(i, len(files))}  {name}  FAILED: {e}")
            log(f"PARSE FAILED {name}: {e}")

    mapping = handle_unknown_headings(cfg, label_files) if label_files else {}
    rows = [remap_unknown(r, mapping) for r in rows]

    paths = write_all(rows, OUT_DIR)
    dt = time.time() - t0
    print(f"\n  Done: {len(rows)} row(s) in {dt:.1f}s. Output files:")
    for p in paths:
        print(f"    {p}")
    log(f"run complete: {len(rows)} rows in {dt:.1f}s -> {paths}")
    print()
    type_out("*  A L L   D O N E  *", delay=0.02)

    misc_rows = [r["_file"] for r in rows if r.get("_misc")]
    if misc_rows:
        print(f"\n  [!] Review needed -- unplaced text (_misc) in: {', '.join(misc_rows)}")
    pause()


def open_output():
    os.makedirs(OUT_DIR, exist_ok=True)
    os.startfile(OUT_DIR)  # noqa -- Windows only


def main():
    os.makedirs(DOCS_DIR, exist_ok=True)
    os.makedirs(OUT_DIR, exist_ok=True)
    cfg = load_config()

    # paths dragged onto run.bat (files or whole Term folders) become the source
    dragged = [a for a in sys.argv[1:] if os.path.exists(a)]
    sources = dragged if dragged else [DOCS_DIR]

    from logger import setup_logging, log
    logpath = setup_logging()
    log(f"session start; sources={sources}; engine={cfg['engine']}")

    intro_animation()
    print(f"  Log file: {logpath}\n")
    from updater import check_for_update
    if check_for_update():
        log("update check: newer release available on GitHub")
    from bootstrap import auto_setup
    auto_setup(cfg)
    time.sleep(1)
    while True:
        clear()
        banner(cfg, load_api_key(), sources)
        print("\n  1) Run extraction")
        print("  2) Select engine (Qwen local / Gemini API)")
        print("  3) Change Gemini API key")
        print("  4) Check / auto-install requirements")
        print("  5) Open output folder")
        print("  6) Open documents folder")
        print("  0) Exit")
        c = input("\n  Choice: ").strip()
        if c == "1":
            run_extraction(cfg, sources)
            cfg = load_config()
        elif c == "2":
            menu_engine(cfg)
        elif c == "3":
            menu_api_key()
        elif c == "4":
            clear()
            print("--- Requirement checks ---")
            auto_setup(cfg)
            run_checks(cfg, sources)
            pause()
        elif c == "5":
            open_output()
        elif c == "6":
            os.startfile(DOCS_DIR)  # noqa
        elif c == "0":
            return


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\nBye.")
