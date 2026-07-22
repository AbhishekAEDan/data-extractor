"""
updater.py

Checks GitHub Releases for a newer version of the app. On startup the app
calls check_for_update(); if the latest release tag on GitHub is newer than
__version__, the user is prompted and (on yes) the update is DOWNLOADED and
INSTALLED in place, then the app restarts itself. User data (documents/,
output/, logs/, .env, config.json) is never touched. If the in-place
install fails for any reason, the releases page opens in the browser as a
fallback. Fully silent on any network failure -- an offline machine never
sees an error.

Author: AbhishekAEDan
"""
__author__ = "AbhishekAEDan"

import io
import json
import os
import re
import shutil
import sys
import tempfile
import urllib.request
import webbrowser
import zipfile

__version__ = "2.2.0"

REPO = "AbhishekAEDan/data-extractor"
RELEASES_PAGE = f"https://github.com/{REPO}/releases/latest"
API_URL = f"https://api.github.com/repos/{REPO}/releases/latest"

BASE = os.path.dirname(os.path.abspath(__file__))

# never overwrite the user's own data/config during an update
PROTECTED = {"documents", "output", "logs", ".env", "config.json"}


def _ver_tuple(s):
    """'v1.2.10' -> (1, 2, 10). Non-numeric junk is ignored."""
    return tuple(int(p) for p in re.findall(r'\d+', s or ""))


def _http_get(url, timeout=15):
    req = urllib.request.Request(
        url, headers={"Accept": "application/vnd.github+json",
                      "User-Agent": "data-extractor-update-check"})
    with urllib.request.urlopen(req, timeout=timeout) as resp:
        return resp.read()


def get_latest_release(timeout=4):
    """Return the latest-release JSON dict, or {} on any failure."""
    try:
        return json.loads(_http_get(API_URL, timeout).decode())
    except Exception:
        return {}


def get_latest_version(timeout=4):
    """Return the latest release tag (e.g. 'v1.1.0') or '' on any failure."""
    return get_latest_release(timeout).get("tag_name", "") or ""


def _zip_url(release):
    """Prefer an uploaded .zip asset (git archive -- no folder prefix);
    fall back to GitHub's auto zipball (repo-sha/ prefixed)."""
    for a in release.get("assets", []):
        if a.get("name", "").lower().endswith(".zip"):
            return a.get("browser_download_url", "")
    return release.get("zipball_url", "")


def _find_root(tree):
    """The folder inside the extracted zip that holds main.py (zipballs
    nest everything under repo-name-sha/; asset zips don't)."""
    if os.path.exists(os.path.join(tree, "main.py")):
        return tree
    for name in os.listdir(tree):
        cand = os.path.join(tree, name)
        if os.path.isdir(cand) and os.path.exists(os.path.join(cand, "main.py")):
            return cand
    return ""


def download_and_install(release, echo=print):
    """Download the release zip and copy it over this install.
    Returns True on success."""
    url = _zip_url(release)
    if not url:
        return False
    echo("      Downloading update...")
    data = _http_get(url, timeout=60)
    tmp = tempfile.mkdtemp(prefix="data-extractor-update-")
    try:
        with zipfile.ZipFile(io.BytesIO(data)) as z:
            z.extractall(tmp)
        root = _find_root(tmp)
        if not root:
            return False
        echo("      Installing...")
        for dirpath, dirnames, filenames in os.walk(root):
            rel = os.path.relpath(dirpath, root)
            top = rel.split(os.sep)[0] if rel != "." else ""
            if top in PROTECTED:
                continue
            for fn in filenames:
                if rel == "." and fn in PROTECTED:
                    continue
                dst_dir = os.path.join(BASE, rel) if rel != "." else BASE
                os.makedirs(dst_dir, exist_ok=True)
                shutil.copy2(os.path.join(dirpath, fn),
                             os.path.join(dst_dir, fn))
        return True
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


def _restart():
    """Relaunch the app with the freshly installed code."""
    main_py = os.path.join(BASE, "main.py")
    os.execl(sys.executable, sys.executable, main_py)


def check_for_update(prompt=input, echo=print):
    """When GitHub has a newer version: offer to download + install it in
    place and restart. Falls back to opening the releases page if the
    in-place install fails. Returns True if an update was offered."""
    release = get_latest_release()
    latest = release.get("tag_name", "") or ""
    if not latest or _ver_tuple(latest) <= _ver_tuple(__version__):
        return False
    echo()
    echo(f"  [*] Update available!  v{__version__}  ->  {latest}")
    try:
        ans = prompt("      Download and install it now? (y/n): ").strip().lower()
    except (EOFError, KeyboardInterrupt):
        return True
    if ans not in ("y", "yes"):
        return True
    try:
        if download_and_install(release, echo=echo):
            echo(f"      Updated to {latest}. Restarting...")
            echo("      (your documents, output, API key and settings are kept)")
            _restart()          # never returns on success
        raise RuntimeError("no usable zip in the release")
    except SystemExit:
        raise
    except Exception as e:
        echo(f"      Auto-update failed ({e}); opening the download page instead.")
        echo(f"      {RELEASES_PAGE}")
        webbrowser.open(RELEASES_PAGE)
        echo("      Download the new version, then copy your .env (API key)")
        echo("      into the new folder.")
    return True
