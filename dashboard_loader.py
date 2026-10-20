# dashboard_loader.py
"""Downloads the batch Excel files from SharePoint and keeps them, already
parsed, in memory for the Analytics Dashboard.

The important part is the background warm-up: the moment anyone opens the app
(in any view), start_warmup_if_due() begins reading every Excel file in a
background thread (including nested subfolders up to 2 levels deep). By the time
someone clicks over to the dashboard the data is usually already there."""
import os
import tempfile
import threading
import time
import traceback
from concurrent.futures import ThreadPoolExecutor

import sharepoint_client
from dashboard_data import parse_workbook
from settings import ENABLE_SHAREPOINT_SYNC

# What the Folder column shows for files that sit directly in the landing folder
MAIN_FOLDER_COLUMN = "(Main folder)"

# Bump this whenever dashboard_data.py parsing rules change so in-memory cached files re-parse automatically
PARSER_VERSION = 5

# A new background check starts at most this often, however many clicks happen
WARMUP_MIN_GAP_SECONDS = 60
# Files read at the same time during the warm-up (kept small to be kind to SharePoint)
WARMUP_WORKERS = 4

# (subfolder, filename, modified, PARSER_VERSION) -> (rows, error_text)
_store = {}
_store_lock = threading.Lock()
_key_locks = {}

_warm_state = {"running": False, "last_started": 0.0}
_warm_lock = threading.Lock()


# ---------------------------------------------------------------------------
# The in-memory store
# ---------------------------------------------------------------------------

def is_loaded(subfolder, filename, modified):
    """True if this version of the file is already in memory."""
    with _store_lock:
        return (subfolder, filename, modified, PARSER_VERSION) in _store


def _download_and_parse(subfolder, filename):
    """Returns (rows, error_text). A network failure raises, so it is never stored as an answer."""
    with tempfile.TemporaryDirectory() as tmp_dir:
        local_path = os.path.join(tmp_dir, "batch.xlsx")
        if not sharepoint_client.download_file(filename, local_path, subfolder):
            return [], "it is no longer in SharePoint"
        try:
            return parse_workbook(local_path, subfolder or MAIN_FOLDER_COLUMN, filename), ""
        except Exception as error:
            return [], f"it could not be read ({error})"


def get_file_rows(subfolder, filename, modified):
    """The defect rows of one batch file, as (rows, error_text)."""
    key = (subfolder, filename, modified, PARSER_VERSION)
    with _store_lock:
        if key in _store:
            return _store[key]
        key_lock = _key_locks.setdefault(key, threading.Lock())

    with key_lock:
        with _store_lock:
            if key in _store:
                return _store[key]
        result = _download_and_parse(subfolder, filename)
        with _store_lock:
            _store[key] = result
        return result


# ---------------------------------------------------------------------------
# Background warm-up
# ---------------------------------------------------------------------------

def _load_one(key):
    try:
        get_file_rows(*key)
    except Exception:
        traceback.print_exc()


def _warm_up():
    """Reads every Excel file in the landing folder and all of its nested subfolders,
    then forgets the stored copies of files that were changed or deleted."""
    try:
        keys = []
        for folder in [""] + sharepoint_client.list_all_output_folders():
            for item in sharepoint_client.list_xlsx_files(folder):
                keys.append((item["subfolder"], item["name"], item["modified"]))

        with ThreadPoolExecutor(max_workers=WARMUP_WORKERS) as pool:
            list(pool.map(_load_one, keys))

        current = {(sub, name, mod, PARSER_VERSION) for sub, name, mod in keys}
        with _store_lock:
            for stale in [k for k in _store if k not in current]:
                del _store[stale]
                _key_locks.pop(stale, None)
    except Exception:
        traceback.print_exc()
    finally:
        with _warm_lock:
            _warm_state["running"] = False


def start_warmup_if_due():
    """Starts the background read of all Excel files, unless one is already running
    or one started less than WARMUP_MIN_GAP_SECONDS ago."""
    if not ENABLE_SHAREPOINT_SYNC:
        return False
    with _warm_lock:
        now = time.time()
        if _warm_state["running"] or now - _warm_state["last_started"] < WARMUP_MIN_GAP_SECONDS:
            return False
        _warm_state["running"] = True
        _warm_state["last_started"] = now
    threading.Thread(target=_warm_up, name="dashboard-warmup", daemon=True).start()
    return True
