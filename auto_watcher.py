# auto_watcher.py
"""Auto-extract: watches the "Whiteboard Photos" folder in SharePoint and adds every new
photo to the right Excel file, with nobody clicking anything.

    Whiteboard Photos / Project Name / [Tower #] / Elevation / Drop / photo.jpg
        -> <Excel landing folder> / Project Name / Project Name - Tower # - Elevation - Drop.xlsx
    (the naming rule lives in auto_naming.py)

How it runs: one background thread, started by app.py (start_if_due), checks the folder
every AUTO_WATCH_INTERVAL_SECONDS. It lives inside the Streamlit process, so it only runs
while the app is awake; after a sleep it simply catches up on everything uploaded meanwhile.

Safety rules (each one answers a way data could otherwise be lost or doubled):
  * The Excel file is pulled fresh from SharePoint right before each photo is added, so
    corrections people made in Excel are never overwritten by an old local copy.
  * A photo counts as processed only AFTER the upload of its Excel file succeeded. If the
    upload fails, the local copy is thrown away and the photo is tried again next check.
  * Which photos are done is kept in SharePoint (auto_watch_state.json), not on the app's
    disk, so a restart never re-adds old photos.
  * A photo modified in the last AUTO_WATCH_SETTLE_SECONDS is skipped until the next check.
  * A manual batch and the watcher never write the same Excel file at once (write_lock.py).

The Pause switch and the "retry failed photos" request come from the app's status card
(ui_auto_watch.py) through auto_watch_control.json, so only one side writes each file."""
import collections
import datetime
import hashlib
import importlib
import json
import os
import tempfile
import threading
import time
import traceback
import uuid

import requests

import sharepoint_client
import sync_tracker
import settings as _settings_mod
from auto_naming import target_for_folder
from excel_writer import write_parsed_data_to_excel
from extraction import extract_data
from file_utils import load_processed_log, save_processed_log, sanitize_batch_filename
from write_lock import lock_for

# Hot-reload safeguard if Streamlit Cloud cached an older settings.py in memory
if not hasattr(_settings_mod, "ENABLE_AUTO_WATCH"):
    importlib.reload(_settings_mod)

ENABLE_SHAREPOINT_SYNC = getattr(_settings_mod, "ENABLE_SHAREPOINT_SYNC", True)
ENABLE_AUTO_WATCH = getattr(_settings_mod, "ENABLE_AUTO_WATCH", True)
AUTO_WATCH_PHOTO_ROOT = getattr(
    _settings_mod,
    "AUTO_WATCH_PHOTO_ROOT",
    ["07 - Employee Folders (to be reviewed)", "Lorenz", "Whiteboard Photos"],
)
AUTO_WATCH_INTERVAL_SECONDS = getattr(_settings_mod, "AUTO_WATCH_INTERVAL_SECONDS", 120)
AUTO_WATCH_SETTLE_SECONDS = getattr(_settings_mod, "AUTO_WATCH_SETTLE_SECONDS", 120)
AUTO_WATCH_MAX_FOLDERS = getattr(_settings_mod, "AUTO_WATCH_MAX_FOLDERS", 2000)
AUTO_WATCH_MAX_ATTEMPTS = getattr(_settings_mod, "AUTO_WATCH_MAX_ATTEMPTS", 3)
AUTO_WATCH_STATE_FILE = getattr(_settings_mod, "AUTO_WATCH_STATE_FILE", "auto_watch_state.json")
AUTO_WATCH_CONTROL_FILE = getattr(_settings_mod, "AUTO_WATCH_CONTROL_FILE", "auto_watch_control.json")
LOCAL_BATCH_DIR = getattr(_settings_mod, "LOCAL_BATCH_DIR", "local_batches")

THREAD_NAME = "auto-watcher"
_LOCAL_TZ = datetime.timezone(datetime.timedelta(hours=8))   # times shown in the app are Manila time
_SELECT = "$select=id,name,size,lastModifiedDateTime,file,folder&$top=200"
_RETRYABLE_STATUS = {423, 429, 503, 504}

# ---------------------------------------------------------------------------
# What the status card shows
# ---------------------------------------------------------------------------

_status_lock = threading.Lock()
_status = {
    "state": "starting",        # starting | idle | scanning | working | paused | error
    "paused": False,
    "last_check": "",           # text, Manila time
    "waiting": 0,               # photos seen but still too fresh to read
    "processed_session": 0,     # photos added since the app started
    "failed": 0,                # photos set aside after repeated failures
    "error": "",                # why the last check could not run, if it could not
    "current_index": 0,         # 1-based index of the photo currently being processed in this cycle
    "current_total": 0,         # total settled photos queued in this cycle
    "current_photo": "",        # filename of the photo currently being processed
    "current_file": "",         # target .xlsx filename currently being updated
}
_events = collections.deque(maxlen=12)   # newest last: {"time", "level", "text"}


def _now_local_text():
    return datetime.datetime.now(_LOCAL_TZ).strftime("%b %d, %H:%M")


def _event(level, text):
    with _status_lock:
        _events.append({"time": _now_local_text(), "level": level, "text": text})


def _set_status(**fields):
    with _status_lock:
        _status.update(fields)


def get_status():
    """A copy of the current status plus the latest events (newest first), for the app."""
    with _status_lock:
        snapshot = dict(_status)
        snapshot["events"] = list(reversed(_events))
    return snapshot


# ---------------------------------------------------------------------------
# Small JSON files kept in the Excel landing folder in SharePoint
# ---------------------------------------------------------------------------

def _json_download(remote_name):
    """The JSON stored in the landing folder under this name, or None if it is not there
    yet. Raises when SharePoint cannot be reached, so a network hiccup is never mistaken
    for "no file" (which would make the watcher start over)."""
    with tempfile.TemporaryDirectory() as tmp:
        path = os.path.join(tmp, "data.json")
        if not sharepoint_client.download_file(remote_name, path, ""):
            return None
        with open(path, "r", encoding="utf-8") as f:
            return json.load(f)


def _json_upload(remote_name, data):
    """Saves JSON into the landing folder, replacing the previous version. Retries the
    short "file is locked / busy" answers SharePoint gives right after a write."""
    payload = json.dumps(data, ensure_ascii=False).encode("utf-8")
    response = None
    for attempt in range(sharepoint_client.MAX_UPLOAD_ATTEMPTS):
        token = sharepoint_client._get_access_token()
        site_id = sharepoint_client._get_site_id()
        item_path = sharepoint_client._encode_item_path(sharepoint_client._remote_folder(""), remote_name)
        url = f"{sharepoint_client.GRAPH_BASE}/sites/{site_id}/drive/root:/{item_path}:/content"
        response = requests.put(
            url,
            headers={"Authorization": f"Bearer {token}", "Content-Type": "application/json"},
            data=payload,
            timeout=60,
        )
        if response.status_code in _RETRYABLE_STATUS and attempt < sharepoint_client.MAX_UPLOAD_ATTEMPTS - 1:
            time.sleep(sharepoint_client._retry_wait_seconds(response, attempt))
            continue
        response.raise_for_status()
        return
    response.raise_for_status()


def _empty_state():
    return {"version": 1, "items": {}}


def _state_load():
    data = _json_download(AUTO_WATCH_STATE_FILE)
    if data is None:
        return _empty_state()
    if not isinstance(data, dict) or not isinstance(data.get("items"), dict):
        raise ValueError(f"{AUTO_WATCH_STATE_FILE} in SharePoint is not in the expected format")
    return data


def _state_save(state):
    """Saved after every photo. Retried a few times: if it still fails, the photo that was
    just added could be added again on the next check, so the caller stops the check."""
    last_error = None
    for attempt in range(3):
        try:
            _json_upload(AUTO_WATCH_STATE_FILE, state)
            return
        except Exception as error:
            last_error = error
            time.sleep(2 + attempt * 2)
    raise last_error


def _control_load():
    """{"paused", "retry_failed"}. If SharePoint cannot be reached, the last known pause
    state is kept instead of guessing."""
    try:
        data = _json_download(AUTO_WATCH_CONTROL_FILE)
    except Exception:
        traceback.print_exc()
        return {"paused": _status["paused"], "retry_failed": False}
    data = data if isinstance(data, dict) else {}
    return {"paused": bool(data.get("paused", False)), "retry_failed": bool(data.get("retry_failed", False))}


def set_paused(paused):
    """Called by the app's Pause switch. Raises if SharePoint cannot be reached."""
    control = _control_load()
    control["paused"] = bool(paused)
    _json_upload(AUTO_WATCH_CONTROL_FILE, control)
    _set_status(paused=bool(paused), state="paused" if paused else "idle")


def request_retry_failed():
    """Called by the app's "Retry failed photos" button; the next check picks it up."""
    control = _control_load()
    control["retry_failed"] = True
    _json_upload(AUTO_WATCH_CONTROL_FILE, control)


# ---------------------------------------------------------------------------
# Looking at the photo folder
# ---------------------------------------------------------------------------

_root_checked = {"done": False}
_ensured_folders = set()


def _ensure_photo_root():
    """Creates "Whiteboard Photos" if it is not there yet, so people can find it."""
    if _root_checked["done"]:
        return
    parent, name = list(AUTO_WATCH_PHOTO_ROOT[:-1]), AUTO_WATCH_PHOTO_ROOT[-1]
    sharepoint_client._create_folder_if_missing(parent, name)
    _root_checked["done"] = True


def _scan_photos():
    """Every photo below the photo root. Returns (photos, hit_folder_limit); each photo is
    {"id", "name", "relative" (folder names below the root), "modified"}."""
    root = list(AUTO_WATCH_PHOTO_ROOT)
    queue, walked, photos, capped = [root], 0, [], False
    while queue:
        current = queue.pop(0)
        walked += 1
        for item in sharepoint_client._iter_children(current, _SELECT):
            name = item.get("name", "")
            if "folder" in item:
                if walked + len(queue) < AUTO_WATCH_MAX_FOLDERS:
                    queue.append(current + [name])
                else:
                    capped = True
                continue
            if "file" not in item or name.startswith("~$"):
                continue
            if not name.lower().endswith(sharepoint_client.IMAGE_EXTENSIONS):
                continue
            if not item.get("size"):
                continue  # an empty file is an upload that has only just started
            photos.append({
                "id": item["id"],
                "name": name,
                "relative": current[len(root):],
                "modified": item.get("lastModifiedDateTime", ""),
            })
    return photos, capped


def _is_settled(modified_iso):
    """False while the photo is newer than AUTO_WATCH_SETTLE_SECONDS (maybe still uploading)."""
    if not modified_iso:
        return True
    try:
        modified = datetime.datetime.fromisoformat(modified_iso.replace("Z", "+00:00"))
    except ValueError:
        return True
    age = (datetime.datetime.now(datetime.timezone.utc) - modified).total_seconds()
    return age >= AUTO_WATCH_SETTLE_SECONDS


def _group_by_excel_file(photos):
    """{(subfolder, filename) lowercased: {"subfolder", "filename", "photos"}}. Capitals are
    ignored when deciding which photos share a file, and the spelling used is the one of
    the oldest photo, so a file never gets renamed because of a later photo."""
    groups = {}
    for photo in photos:
        subfolder, filename = target_for_folder(
            photo["relative"], sanitize_batch_filename, sharepoint_client.sanitize_folder_name
        )
        key = (subfolder.casefold(), filename.casefold())
        group = groups.get(key)
        if group is None:
            groups[key] = {"subfolder": subfolder, "filename": filename,
                           "photos": [photo], "oldest": photo["modified"]}
            continue
        group["photos"].append(photo)
        if photo["modified"] < group["oldest"]:
            group["subfolder"], group["filename"], group["oldest"] = subfolder, filename, photo["modified"]
    return groups


# ---------------------------------------------------------------------------
# Adding one photo to its Excel file
# ---------------------------------------------------------------------------

def _now_iso():
    return datetime.datetime.now(datetime.timezone.utc).isoformat(timespec="seconds")


def _record(state, item_id, **fields):
    entry = state["items"].setdefault(item_id, {})
    entry.update(fields)
    entry["updated"] = _now_iso()
    _state_save(state)


def _record_failure(state, photo, error_text):
    attempts = int(state["items"].get(photo["id"], {}).get("attempts", 0)) + 1
    gave_up = attempts >= AUTO_WATCH_MAX_ATTEMPTS
    _record(state, photo["id"], status="gave_up" if gave_up else "failed",
            attempts=attempts, name=photo["name"], error=str(error_text)[:300])
    if gave_up:
        _event("error", f"{photo['name']}: set aside after {attempts} failed tries ({error_text})")
    else:
        _event("warning", f"{photo['name']}: failed ({error_text}); it will be tried again")


def _fresh_local_copy(local_path, filename, subfolder):
    """Makes local_path the current SharePoint copy of the Excel file (or removes it when
    the file does not exist there yet, so a brand-new workbook is started)."""
    if os.path.exists(local_path):
        os.remove(local_path)
    sharepoint_client.download_file(filename, local_path, subfolder)


def _remove_local(local_path):
    if os.path.exists(local_path):
        os.remove(local_path)
    sync_tracker.forget(local_path)


def _process_photo(photo, group, local_path, state, known_hashes, local_log):
    """Returns "done", "duplicate", "failed", "retry" (a network problem: nothing was
    recorded) or "upload_failed" (the caller must stop working on this Excel file)."""
    filename, subfolder = group["filename"], group["subfolder"]

    try:
        data = sharepoint_client.download_item_to_bytes(photo["id"])
    except Exception as error:
        _event("warning", f"{photo['name']}: could not be downloaded right now ({error}); will retry")
        return "retry"
    if data is None:
        return "retry"  # deleted meanwhile

    file_hash = hashlib.sha256(data).hexdigest()
    if file_hash in known_hashes or file_hash in local_log:
        _record(state, photo["id"], status="duplicate", hash=file_hash, name=photo["name"])
        _event("info", f"{photo['name']}: already processed before (same photo), skipped")
        return "duplicate"

    extension = os.path.splitext(photo["name"])[1].lower() or ".jpg"
    temp_path = os.path.join(tempfile.gettempdir(), f"auto_{uuid.uuid4().hex}{extension}")
    with open(temp_path, "wb") as f:
        f.write(data)

    try:
        # The slow part (Gemini) comes first; the Excel file is pulled fresh only after it,
        # so the time between "latest copy" and "uploaded" is a few seconds.
        try:
            parsed = extract_data(temp_path)
        except Exception as error:
            traceback.print_exc()
            _record_failure(state, photo, f"reading the photo failed: {error}")
            return "failed"

        try:
            _fresh_local_copy(local_path, filename, subfolder)
        except Exception as error:
            traceback.print_exc()
            _event("warning", f"{filename}: could not load the current copy from SharePoint ({error}); will retry")
            return "retry"

        try:
            write_parsed_data_to_excel(parsed, temp_path, local_path)
        except Exception as error:
            traceback.print_exc()
            _record_failure(state, photo, f"writing to Excel failed: {error}")
            return "failed"

        ok, reason = sync_tracker.sync_file(local_path, subfolder)
        if not ok:
            _event("error", f"{filename}: upload to SharePoint failed ({reason}); {photo['name']} will be retried")
            return "upload_failed"

        notes = parsed.get("Review Notes", []) or []
        flagged = bool(parsed.get("Flagged Materials"))
        _record(state, photo["id"], status="done", hash=file_hash, name=photo["name"],
                file=f"{subfolder}/{filename}" if subfolder else filename, flagged=flagged)
        known_hashes.add(file_hash)
        local_log.append(file_hash)
        save_processed_log(local_log)   # so the manual tools also skip this photo

        if flagged:
            _event("warning", f"Added {photo['name']} to {filename}. Some rows are highlighted red for review.")
        elif notes:
            _event("info", f"Added {photo['name']} to {filename}. Note: {notes[0]}")
        else:
            _event("success", f"Added {photo['name']} to {filename}")
        return "done"
    finally:
        if os.path.exists(temp_path):
            os.remove(temp_path)


def _ensure_output_folder(subfolder):
    if not subfolder or subfolder.casefold() in _ensured_folders:
        return
    sharepoint_client.ensure_subfolder(subfolder)
    _ensured_folders.add(subfolder.casefold())


def _process_group(group, pending, state, known_hashes, local_log, progress_counter):
    """Adds the pending photos of one Excel file. Returns how many were added."""
    subfolder, filename = group["subfolder"], group["filename"]
    local_dir = os.path.join(LOCAL_BATCH_DIR, subfolder) if subfolder else LOCAL_BATCH_DIR
    os.makedirs(local_dir, exist_ok=True)
    local_path = os.path.join(local_dir, filename)

    lock = lock_for(local_path)
    if not lock.acquire(blocking=False):
        _event("info", f"{filename} is being updated by a manual batch; its new photos wait for the next check")
        progress_counter["done"] += len(pending)
        return 0
    added = 0
    try:
        _ensure_output_folder(subfolder)
        for photo in pending:
            progress_counter["done"] += 1
            _set_status(
                state="working",
                current_index=progress_counter["done"],
                current_total=progress_counter["total"],
                current_photo=photo["name"],
                current_file=filename,
            )
            result = _process_photo(photo, group, local_path, state, known_hashes, local_log)
            if result == "done":
                added += 1
                with _status_lock:
                    _status["processed_session"] += 1
            elif result == "upload_failed":
                break   # later photos would carry this one's rows into the next upload
    finally:
        _remove_local(local_path)
        lock.release()
    return added


# ---------------------------------------------------------------------------
# One check, and the thread that repeats it
# ---------------------------------------------------------------------------

def _cycle():
    _ensure_photo_root()
    control = _control_load()
    _set_status(paused=control["paused"])
    if control["paused"]:
        _set_status(
            state="paused", error="", last_check=_now_local_text(),
            current_index=0, current_total=0, current_photo="", current_file="",
        )
        return

    state = _state_load()
    if control["retry_failed"]:
        stuck = [i for i, e in state["items"].items() if e.get("status") in ("failed", "gave_up")]
        for item_id in stuck:
            del state["items"][item_id]
        if stuck:
            _state_save(state)
            _event("info", f"Retrying {len(stuck)} photo(s) that had failed")
        control["retry_failed"] = False
        _json_upload(AUTO_WATCH_CONTROL_FILE, control)

    _set_status(
        state="scanning", error="",
        current_index=0, current_total=0, current_photo="", current_file="",
    )
    photos, capped = _scan_photos()
    if capped:
        _event("warning", f"The photo folder has more than {AUTO_WATCH_MAX_FOLDERS} folders; the deepest ones were not checked")

    todo_ids, waiting = set(), 0
    for photo in photos:
        entry = state["items"].get(photo["id"], {})
        if entry.get("status") in ("done", "duplicate", "gave_up"):
            continue
        if not _is_settled(photo["modified"]):
            waiting += 1
            continue
        todo_ids.add(photo["id"])

    known_hashes = {e["hash"] for e in state["items"].values() if e.get("hash")}
    local_log = list(load_processed_log())

    total_todo = len(todo_ids)
    _set_status(
        state="working" if total_todo > 0 else "idle",
        waiting=waiting,
        current_index=0,
        current_total=total_todo,
        current_photo="",
        current_file="",
    )

    progress_counter = {"done": 0, "total": total_todo}
    groups = _group_by_excel_file(photos)
    for index, key in enumerate(sorted(groups)):
        group = groups[key]
        pending = sorted((p for p in group["photos"] if p["id"] in todo_ids),
                         key=lambda p: (p["modified"], p["name"]))
        if not pending:
            continue
        if index > 0 and _control_load()["paused"]:
            break   # paused while this check was running
        _process_group(group, pending, state, known_hashes, local_log, progress_counter)

    failed = sum(1 for e in state["items"].values() if e.get("status") in ("failed", "gave_up"))
    _set_status(
        state="paused" if _control_load()["paused"] else "idle",
        waiting=waiting,
        failed=failed,
        last_check=_now_local_text(),
        current_index=0,
        current_total=0,
        current_photo="",
        current_file="",
    )


def _loop():
    while True:
        try:
            _cycle()
        except Exception as error:
            traceback.print_exc()
            _set_status(
                state="error",
                error=f"{type(error).__name__}: {error}"[:300],
                last_check=_now_local_text(),
                current_index=0,
                current_total=0,
                current_photo="",
                current_file="",
            )
        time.sleep(AUTO_WATCH_INTERVAL_SECONDS)


_start_lock = threading.Lock()


def start_if_due():
    """Starts the background thread unless it is already running. Cheap to call on every
    page run. Returns True if a new thread was started."""
    if not (ENABLE_SHAREPOINT_SYNC and ENABLE_AUTO_WATCH):
        return False
    with _start_lock:
        if any(t.name == THREAD_NAME and t.is_alive() for t in threading.enumerate()):
            return False
        threading.Thread(target=_loop, name=THREAD_NAME, daemon=True).start()
        return True