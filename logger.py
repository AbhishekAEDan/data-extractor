"""
logger.py -- tee console output into dated log files under ./logs.

Everything printed (and typed prompts' surrounding output) lands in
logs/run_YYYY-MM-DD_HH-MM-SS.txt, flushed on every write so the file
updates live while the process runs.

Author: AbhishekAEDan
"""
__author__ = "AbhishekAEDan"

import os
import sys
from datetime import datetime

LOGS_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "logs")


class Tee:
    def __init__(self, stream, logfile):
        self.stream = stream
        self.logfile = logfile

    def write(self, data):
        self.stream.write(data)
        try:
            self.logfile.write(data)
            self.logfile.flush()
        except Exception:
            pass

    def flush(self):
        self.stream.flush()
        try:
            self.logfile.flush()
        except Exception:
            pass

    def __getattr__(self, name):
        return getattr(self.stream, name)


_logfile = None


def setup_logging():
    """Redirect stdout/stderr through a tee into a new dated log file.
    Returns the log file path."""
    global _logfile
    os.makedirs(LOGS_DIR, exist_ok=True)
    stamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    path = os.path.join(LOGS_DIR, f"run_{stamp}.txt")
    _logfile = open(path, "a", encoding="utf-8")
    _logfile.write(f"=== Brent & Co. Data Extractor log -- {stamp} ===\n")
    _logfile.flush()
    sys.stdout = Tee(sys.stdout, _logfile)
    sys.stderr = Tee(sys.stderr, _logfile)
    return path


def log(msg):
    """Backend-only detail: goes to the log file, not the console."""
    if _logfile:
        try:
            stamp = datetime.now().strftime("%H:%M:%S")
            _logfile.write(f"[{stamp}] {msg}\n")
            _logfile.flush()
        except Exception:
            pass
