# file_utils.py
import os
import re
import json

from settings import LOG_FILE, PREFS_FILE


def load_processed_log():
    """Loads the list of already processed photo fingerprints."""
    if os.path.exists(LOG_FILE):
        try:
            with open(LOG_FILE, "r") as f:
                return json.load(f)
        except (json.JSONDecodeError, OSError):
            return []
    return []


def save_processed_log(log_list):
    """Saves the log safely (write to a temp file, then swap it in)."""
    tmp_path = LOG_FILE + ".tmp"
    with open(tmp_path, "w") as f:
        json.dump(log_list, f)
    os.replace(tmp_path, LOG_FILE)


def load_prefs():
    """The remembered choices (last photo folder, last save folder...) as a dict.
    Returns {} if nothing was saved yet or the file can't be read."""
    if os.path.exists(PREFS_FILE):
        try:
            with open(PREFS_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
            return data if isinstance(data, dict) else {}
        except (json.JSONDecodeError, OSError):
            return {}
    return {}


def save_prefs(updates):
    """Adds `updates` (a dict) to the remembered choices and saves them safely
    (temp file, then swap it in). Never raises: failing to remember a choice
    must not stop a batch."""
    prefs = load_prefs()
    prefs.update(updates)
    tmp_path = PREFS_FILE + ".tmp"
    try:
        with open(tmp_path, "w", encoding="utf-8") as f:
            json.dump(prefs, f, indent=2)
        os.replace(tmp_path, PREFS_FILE)
    except OSError:
        pass


def sanitize_batch_filename(name):
    """Turns what the person typed into a safe .xlsx filename, or None if
    nothing usable was entered."""
    name = (name or "").strip()
    if not name:
        return None
    name = re.sub(r'[\\/:*?"<>|]', "_", name)  # characters not allowed in Windows/SharePoint filenames
    name = name.rstrip(". ")  # trailing dots/spaces aren't allowed either
    name = name[:150]
    if not name:
        return None
    return name if name.lower().endswith(".xlsx") else f"{name}.xlsx"
