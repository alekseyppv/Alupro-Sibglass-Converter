from __future__ import annotations

from pathlib import Path
import sys


def _base_dir() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parents[1]


BASE_DIR = _base_dir()
CONFIG_DIR = BASE_DIR / "config"
DATA_DIR = BASE_DIR / "data"
LOG_FILE = BASE_DIR / "errors.log"
SETTINGS_FILE = CONFIG_DIR / "settings.json"
GLASS_FILE = DATA_DIR / "glass.txt"
AUTOSAVE_FILE = DATA_DIR / "autosave.tmp"


for directory in (CONFIG_DIR, DATA_DIR):
    directory.mkdir(parents=True, exist_ok=True)
