# file_utils.py
import os
import re
import json

from settings import LOG_FILE


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
