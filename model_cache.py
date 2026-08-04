"""
model_cache.py -- portable local model cache.

Mirrors Ollama's on-disk store (`manifests/...` + `blobs/sha256-...`) into a
sibling `..\\AI Model` folder next to the app, so a fresh computer can import
the model verbatim instead of re-downloading it. All operations are best-effort
and non-fatal: AI is optional, so any failure prints a warning and returns
False -- the app still runs.

Store layout (both the real Ollama store and this cache use it):
  manifests/registry.ollama.ai/library/<model>/<tag>   (JSON manifest)
  blobs/sha256-<hex>                                    (config + layer blobs)

Author: AbhishekAEDan
"""
__author__ = "AbhishekAEDan"

import json
import os
import shutil
import subprocess
import time
import urllib.request

BASE = os.path.dirname(os.path.abspath(__file__))
# sibling folder next to the app dir, created on demand
CACHE_DIR = os.path.join(os.path.dirname(BASE), "AI Model")

OLLAMA_URL = "http://localhost:11434"


def _split_name(name):
    """`qwen3:8b` -> ('qwen3', '8b'); `qwen3` -> ('qwen3', 'latest')."""
    if ":" in name:
        model, tag = name.split(":", 1)
    else:
        model, tag = name, "latest"
    return model, tag


def _manifest_relpath(name):
    """Relative path (under a store root) to this model tag's manifest."""
    model, tag = _split_name(name)
    return os.path.join("manifests", "registry.ollama.ai", "library", model, tag)


def _blob_filename(digest):
    """`sha256:<hex>` -> `sha256-<hex>` (the on-disk blob file name)."""
    return digest.replace(":", "-", 1)


def _referenced_digests(manifest_path):
    """Every blob digest a manifest JSON points at: config + all layers."""
    with open(manifest_path, "r", encoding="utf-8") as f:
        data = json.load(f)
    digests = []
    cfg = data.get("config") or {}
    if cfg.get("digest"):
        digests.append(cfg["digest"])
    for layer in data.get("layers") or []:
        if layer.get("digest"):
            digests.append(layer["digest"])
    return digests


def ollama_models_dir():
    """Where Ollama keeps its store: $OLLAMA_MODELS or %USERPROFILE%\\.ollama\\models."""
    env = os.environ.get("OLLAMA_MODELS")
    if env:
        return env
    return os.path.join(os.environ.get("USERPROFILE", os.path.expanduser("~")),
                        ".ollama", "models")


def model_in_ollama(name):
    """True if Ollama's /api/tags lists this model tag. Silent False on error."""
    try:
        with urllib.request.urlopen(f"{OLLAMA_URL}/api/tags", timeout=4) as r:
            data = json.loads(r.read().decode())
    except Exception:
        return False
    names = [m.get("name", "") for m in data.get("models", [])]
    if name in names:
        return True
    # a bare name (tag latest) may be listed as `name:latest`
    model, tag = _split_name(name)
    return f"{model}:{tag}" in names


def model_in_cache(name):
    """True if the cache holds this tag's manifest AND every blob it references."""
    manifest = os.path.join(CACHE_DIR, _manifest_relpath(name))
    if not os.path.isfile(manifest):
        return False
    try:
        digests = _referenced_digests(manifest)
    except Exception:
        return False
    for d in digests:
        if not os.path.isfile(os.path.join(CACHE_DIR, "blobs", _blob_filename(d))):
            return False
    return True


def _copy_model(name, src_root, dst_root):
    """Copy manifest + all referenced blobs from src_root store to dst_root store.
    Creates dirs; skips blobs already present; never overwrites existing blobs."""
    rel = _manifest_relpath(name)
    src_manifest = os.path.join(src_root, rel)
    if not os.path.isfile(src_manifest):
        return False
    digests = _referenced_digests(src_manifest)

    # blobs first, so the manifest never points at a missing blob mid-copy
    for d in digests:
        fn = _blob_filename(d)
        src_blob = os.path.join(src_root, "blobs", fn)
        dst_blob = os.path.join(dst_root, "blobs", fn)
        if os.path.isfile(dst_blob):
            continue  # never overwrite an existing blob
        if not os.path.isfile(src_blob):
            return False
        os.makedirs(os.path.dirname(dst_blob), exist_ok=True)
        shutil.copy2(src_blob, dst_blob)

    dst_manifest = os.path.join(dst_root, rel)
    os.makedirs(os.path.dirname(dst_manifest), exist_ok=True)
    if not os.path.isfile(dst_manifest):
        shutil.copy2(src_manifest, dst_manifest)
    return True


def import_from_cache(name):
    """Copy the model from the cache into Ollama's store. Return True/False."""
    try:
        return _copy_model(name, CACHE_DIR, ollama_models_dir())
    except Exception as e:
        print(f"  [MODEL] import failed: {e}")
        return False


def export_to_cache(name):
    """Copy the model from Ollama's store into the cache. Return True/False."""
    try:
        os.makedirs(CACHE_DIR, exist_ok=True)
        return _copy_model(name, ollama_models_dir(), CACHE_DIR)
    except Exception as e:
        print(f"  [MODEL] export failed: {e}")
        return False


def _restart_ollama():
    """Restart the ollama service so it re-reads the store and lists the imported
    model in /api/tags. Uses the app's detached-spawn pattern (DETACHED_PROCESS)
    so the console window can still close on exit (see bootstrap.py note)."""
    try:
        flags = getattr(subprocess, "DETACHED_PROCESS", 0) | \
            getattr(subprocess, "CREATE_NEW_PROCESS_GROUP", 0)
        subprocess.Popen(["ollama", "serve"],
                         stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
                         stdin=subprocess.DEVNULL, creationflags=flags)
    except Exception:
        return
    for _ in range(20):
        time.sleep(1)
        if model_in_ollama_server_up():
            return


def model_in_ollama_server_up():
    try:
        with urllib.request.urlopen(f"{OLLAMA_URL}/api/version", timeout=4):
            return True
    except Exception:
        return False


def ensure_model(name):
    """Make `name` available to Ollama, using the portable cache to avoid
    re-downloads. Non-fatal: on any failure prints a warning, returns False.

      - in ollama              -> export to cache if not yet cached; done.
      - not in ollama, cached  -> import from cache (restart ollama if needed).
      - in neither             -> `ollama pull`, then export to cache.
    """
    try:
        if model_in_ollama(name):
            if not model_in_cache(name):
                print(f"  [MODEL] backing up '{name}' to AI Model cache...")
                if export_to_cache(name):
                    print("  [MODEL] backed up (future computers skip the download)")
                else:
                    print("  [MODEL] backup skipped (could not copy from Ollama store)")
            return True

        if model_in_cache(name):
            print("  [MODEL] found in AI Model cache -- importing (no download needed)")
            if not import_from_cache(name):
                print("  [MODEL] import failed -- falling back to download")
            else:
                if model_in_ollama(name):
                    print("  [MODEL] imported and ready")
                    return True
                print("  [MODEL] restarting Ollama so it picks up the imported model...")
                _restart_ollama()
                if model_in_ollama(name):
                    print("  [MODEL] imported and ready")
                    return True
                print("  [MODEL] Ollama did not list the model after import -- downloading")

        print(f"  [MODEL] '{name}' not local and not cached -- downloading (one-time)...")
        ok = subprocess.run(["ollama", "pull", name]).returncode == 0
        if not ok:
            print(f"  [MODEL] download of '{name}' failed")
            return False
        print("  [MODEL] downloaded -- backing up to AI Model cache...")
        if export_to_cache(name):
            print("  [MODEL] backed up (future computers skip the download)")
        else:
            print("  [MODEL] backup skipped (could not copy from Ollama store)")
        return True
    except Exception as e:
        print(f"  [MODEL] cache step failed ({e}) -- AI stays optional, app continues")
        return False
