"""
checks.py -- environment/requirements checks shown in the menu.

Author: AbhishekAEDan
"""
__author__ = "AbhishekAEDan"
import json
import os
import urllib.request

OLLAMA_URL = "http://localhost:11434"


def check_python_docx():
    try:
        import docx  # noqa: F401
        return True, "python-docx installed"
    except ImportError:
        return False, "python-docx MISSING -> pip install python-docx"


def check_genai():
    try:
        import google.genai  # noqa: F401
        return True, "google-genai installed"
    except ImportError:
        return False, "google-genai MISSING -> pip install google-genai"


def check_ollama_server():
    try:
        with urllib.request.urlopen(f"{OLLAMA_URL}/api/version", timeout=3) as r:
            v = json.loads(r.read().decode()).get("version", "?")
        return True, f"Ollama server running (v{v})"
    except Exception:
        return False, "Ollama server NOT running -> start Ollama app or `ollama serve`"


def ollama_models():
    try:
        with urllib.request.urlopen(f"{OLLAMA_URL}/api/tags", timeout=3) as r:
            data = json.loads(r.read().decode())
        return [m["name"] for m in data.get("models", [])]
    except Exception:
        return []


def check_ollama_model(model):
    models = ollama_models()
    if not models:
        return False, "cannot list Ollama models (server down?)"
    base = model.split(":")[0]
    if model in models or any(m.startswith(base) for m in models):
        return True, f"model '{model}' available"
    return False, f"model '{model}' NOT pulled -> ollama pull {model}"


def check_gemini_key(api_key):
    if not api_key:
        return False, "no GEMINI_API_KEY set"
    try:
        from google import genai
        client = genai.Client(api_key=api_key)
        # cheap validation: list models (no generation cost)
        next(iter(client.models.list()), None)
        return True, "Gemini API key valid"
    except Exception as e:
        return False, f"Gemini key check failed: {str(e)[:80]}"


def check_docs_folder(docs_dir):
    if not os.path.isdir(docs_dir):
        return False, f"documents folder missing: {docs_dir}"
    n = len([f for f in os.listdir(docs_dir)
             if f.lower().endswith(".docx") and not f.startswith("~$")])
    if n == 0:
        return False, f"no .docx files in {docs_dir}"
    return True, f"{n} .docx file(s) ready"
