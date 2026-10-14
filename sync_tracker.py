# sync_tracker.py
"""Keeps track of which local batch files have reached SharePoint, so a file
whose upload failed is never silently forgotten.

How it decides what is "pending": after every successful upload it saves the
file's last-modified time in sync_status.json. A batch file in local_batches/ is
pending when it has no saved time (never uploaded) or was modified after it
(changed since the last upload)."""
import os
import json
import datetime
import traceback

import sharepoint_client
from settings import LOCAL_BATCH_DIR, SYNC_STATUS_FILE


# ---------------------------------------------------------------------------
# The status file
# ---------------------------------------------------------------------------

def _key(local_path):
    """The file's path inside local_batches/, e.g. 'Folder/Batch_1.xlsx'."""
    return os.path.relpath(local_path, LOCAL_BATCH_DIR).replace(os.sep, "/")


def _now():
    return datetime.datetime.now().strftime("%Y-%m-%d %H:%M")


def _load():
    if os.path.exists(SYNC_STATUS_FILE):
        try:
            with open(SYNC_STATUS_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except (json.JSONDecodeError, OSError):
            return {}
    return {}


def _save(status):
    """Saves safely (write to a temp file, then swap it in)."""
    tmp_path = SYNC_STATUS_FILE + ".tmp"
    with open(tmp_path, "w", encoding="utf-8") as f:
        json.dump(status, f, indent=2)
    os.replace(tmp_path, SYNC_STATUS_FILE)


def mark_synced(local_path):
    """Records that the file, as it is right now, matches SharePoint."""
    status = _load()
    status[_key(local_path)] = {
        "synced_mtime": os.path.getmtime(local_path),
        "synced_at": _now(),
        "last_error": "",
    }
    _save(status)


def mark_failed(local_path, error):
    """Records a failed upload. The last successful upload time (if any) is kept,
    so the file still counts as pending."""
    status = _load()
    record = status.get(_key(local_path), {})
    record["last_error"] = error
    record["last_attempt"] = _now()
    status[_key(local_path)] = record
    _save(status)


def forget(local_path):
    """Drops the file's record (used when its local copy is deleted)."""
    status = _load()
    if status.pop(_key(local_path), None) is not None:
        _save(status)


# ---------------------------------------------------------------------------
# Uploading
# ---------------------------------------------------------------------------

def _describe_error(error):
    response = getattr(error, "response", None)
    status_code = getattr(response, "status_code", None)
    if status_code:
        return f"SharePoint returned HTTP {status_code}"
    return f"{type(error).__name__}: {error}"[:200]


def sync_file(local_path, subfolder=""):
    """Uploads one batch file to SharePoint and records the result.
    Returns (True, "") on success or (False, reason) on failure. Never raises."""
    try:
        sharepoint_client.upload_file(local_path, os.path.basename(local_path), subfolder)
    except Exception as e:
        traceback.print_exc()
        reason = _describe_error(e)
        mark_failed(local_path, reason)
        return False, reason
    mark_synced(local_path)
    return True, ""


# ---------------------------------------------------------------------------
# Finding stranded files
# ---------------------------------------------------------------------------

def pending_files():
    """Batch files in local_batches/ that are not up to date in SharePoint.
    Each item: path, key, filename, subfolder, reason, last_error, last_attempt."""
    status = _load()
    pending = []
    for folder, _subdirs, files in os.walk(LOCAL_BATCH_DIR):
        for name in files:
            if not name.lower().endswith(".xlsx"):
                continue
            path = os.path.join(folder, name)
            key = _key(path)
            record = status.get(key, {})

            synced_mtime = record.get("synced_mtime")
            if synced_mtime is None:
                reason = "Never uploaded to SharePoint"
            elif os.path.getmtime(path) > synced_mtime:
                reason = "Changed since the last upload to SharePoint"
            else:
                continue

            pending.append({
                "path": path,
                "key": key,
                "filename": name,
                "subfolder": os.path.dirname(key),
                "reason": reason,
                "last_error": record.get("last_error", ""),
                "last_attempt": record.get("last_attempt", ""),
            })
    return sorted(pending, key=lambda item: item["key"].lower())
