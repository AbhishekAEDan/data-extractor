"""
bootstrap.py -- auto-install prerequisites.

Called at startup and before every run: anything missing gets installed
(pip packages, Ollama itself via winget/installer, the judge model via
`ollama pull`). Everything installs per-user -- no admin rights needed.

Author: AbhishekAEDan
"""
__author__ = "AbhishekAEDan"

import importlib
import os
import shutil
import subprocess
import sys
import tempfile
import time
import urllib.request

PIP_PACKAGES = [("docx", "python-docx"), ("google.genai", "google-genai")]
OLLAMA_SETUP_URL = "https://ollama.com/download/OllamaSetup.exe"


def _run(cmd, **kw):
    return subprocess.run(cmd, **kw).returncode == 0


def ensure_python_packages():
    for module, pip_name in PIP_PACKAGES:
        try:
            importlib.import_module(module)
        except ImportError:
            print(f"  [setup] Installing {pip_name} ...")
            if not _run([sys.executable, "-m", "pip", "install", pip_name]):
                print(f"  [setup] FAILED to install {pip_name} -- install manually:")
                print(f"          {sys.executable} -m pip install {pip_name}")
                return False
    return True


def ollama_cli_present():
    return shutil.which("ollama") is not None


def ensure_ollama_installed():
    if ollama_cli_present():
        return True
    print("  [setup] Ollama not installed. Installing...")
    # 1) try winget (present on Win10/11)
    if shutil.which("winget"):
        ok = _run(["winget", "install", "--id", "Ollama.Ollama", "--silent",
                   "--accept-package-agreements", "--accept-source-agreements"])
        if ok and _refresh_path_and_check():
            print("  [setup] Ollama installed via winget.")
            return True
    # 2) fallback: download official installer, silent install
    try:
        print("  [setup] Downloading Ollama installer (may take a few minutes)...")
        setup_path = os.path.join(tempfile.gettempdir(), "OllamaSetup.exe")
        urllib.request.urlretrieve(OLLAMA_SETUP_URL, setup_path)
        print("  [setup] Running installer silently...")
        _run([setup_path, "/VERYSILENT", "/NORESTART"])
        if _refresh_path_and_check():
            print("  [setup] Ollama installed.")
            return True
    except Exception as e:
        print(f"  [setup] Installer failed: {e}")
    print("  [setup] Could not install Ollama automatically.")
    print("          Install manually from https://ollama.com/download")
    return False


def _refresh_path_and_check():
    """New installs land in %LOCALAPPDATA%\\Programs\\Ollama -- current
    process PATH may not know it yet."""
    candidates = [
        os.path.join(os.environ.get("LOCALAPPDATA", ""), "Programs", "Ollama"),
        r"C:\Program Files\Ollama",
    ]
    for c in candidates:
        if os.path.exists(os.path.join(c, "ollama.exe")):
            os.environ["PATH"] = c + os.pathsep + os.environ.get("PATH", "")
    return ollama_cli_present()


def ensure_ollama_running():
    from checks import check_ollama_server
    if check_ollama_server()[0]:
        return True
    print("  [setup] Starting Ollama server...")
    try:
        # DETACHED_PROCESS: don't tie ollama to this console window --
        # otherwise the window can't close on exit until ollama is killed
        flags = getattr(subprocess, "DETACHED_PROCESS", 0) | \
            getattr(subprocess, "CREATE_NEW_PROCESS_GROUP", 0)
        subprocess.Popen(["ollama", "serve"],
                         stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
                         stdin=subprocess.DEVNULL, creationflags=flags)
    except FileNotFoundError:
        return False
    for _ in range(20):
        time.sleep(1)
        if check_ollama_server()[0]:
            print("  [setup] Ollama server up.")
            return True
    print("  [setup] Ollama server did not start in time.")
    return False


def ensure_ollama_model(model):
    from checks import ollama_models
    models = ollama_models()
    base = model.split(":")[0]
    if model in models or any(m.startswith(base) for m in models):
        return True
    print(f"  [setup] Model '{model}' not found locally. Pulling (one-time download)...")
    # inherit console so the user sees ollama's own progress bar
    ok = _run(["ollama", "pull", model])
    if ok:
        print(f"  [setup] Model '{model}' ready.")
        return True
    print(f"  [setup] FAILED to pull '{model}'. Check the model name / internet.")
    return False


def auto_setup(cfg):
    """Install/start everything the selected engine needs.
    Returns True when ready to run."""
    print("\n--- Prerequisite check & auto-install ---")
    ok = ensure_python_packages()
    if cfg["engine"] == "ollama":
        ok = ensure_ollama_installed() and ok
        ok = ollama_cli_present() and ensure_ollama_running() and ok
        if ok:
            ok = ensure_ollama_model(cfg["ollama_model"])
    print("--- Setup " + ("complete" if ok else "INCOMPLETE -- see messages above") + " ---\n")
    return ok
