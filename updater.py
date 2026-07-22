"""
updater.py

Checks GitHub Releases for a newer version of the app. On startup the app
calls check_for_update(); if the latest release tag on GitHub is newer than
__version__, the user is prompted and (on yes) the releases page opens in
their browser. Fully silent on any network failure -- an offline machine
never sees an error.

Author: AbhishekAEDan
"""
__author__ = "AbhishekAEDan"

import json
import re
import urllib.request
import webbrowser

__version__ = "1.0.1"

REPO = "AbhishekAEDan/data-extractor"
RELEASES_PAGE = f"https://github.com/{REPO}/releases/latest"
API_URL = f"https://api.github.com/repos/{REPO}/releases/latest"


def _ver_tuple(s):
    """'v1.2.10' -> (1, 2, 10). Non-numeric junk is ignored."""
    return tuple(int(p) for p in re.findall(r'\d+', s or ""))


def get_latest_version(timeout=4):
    """Return the latest release tag (e.g. 'v1.1.0') or '' on any failure."""
    try:
        req = urllib.request.Request(
            API_URL, headers={"Accept": "application/vnd.github+json",
                              "User-Agent": "data-extractor-update-check"})
        with urllib.request.urlopen(req, timeout=timeout) as resp:
            data = json.loads(resp.read().decode())
        return data.get("tag_name", "") or ""
    except Exception:
        return ""


def check_for_update(prompt=input, echo=print):
    """Prompt to open the releases page when GitHub has a newer version.
    Returns True if an update was offered."""
    latest = get_latest_version()
    if not latest or _ver_tuple(latest) <= _ver_tuple(__version__):
        return False
    echo()
    echo(f"  [*] Update available!  v{__version__}  ->  {latest}")
    echo(f"      {RELEASES_PAGE}")
    try:
        ans = prompt("      Open the download page now? (y/n): ").strip().lower()
    except (EOFError, KeyboardInterrupt):
        return True
    if ans in ("y", "yes"):
        webbrowser.open(RELEASES_PAGE)
        echo("      Opened in your browser. Download the new version, then")
        echo("      copy your .env (API key) into the new folder.")
    return True
