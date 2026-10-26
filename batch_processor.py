# batch_processor.py
"""Processes a batch of whiteboard photos, sorts them chronologically by Date,
groups them by Day into chunks of up to 6 photos, and generates a PowerPoint (.pptx) report.

A "photo" is any object with a .name and a .getvalue() that returns the image
bytes: a Streamlit UploadedFile, or app.py's RemoteFile (which downloads the
photo from SharePoint at the moment getvalue() is called).
"""
import datetime
import hashlib
import os
import re
import traceback
import uuid
from collections import OrderedDict
from dataclasses import dataclass, field

import sharepoint_client
import sync_tracker
from extraction import extract_data
from file_utils import load_processed_log, save_processed_log
from pptx_writer import format_defect_text, write_chunks_to_pptx
from settings import ENABLE_SHAREPOINT_SYNC

MAX_PHOTOS_PER_CHUNK = 6

_DATE_FORMATS = (
    "%B %d, %Y",
    "%b %d, %Y",
    "%B %d %Y",
    "%b %d %Y",
    "%d %B %Y",
    "%d %b %Y",
    "%d %B, %Y",
    "%d %b, %Y",
    "%Y-%m-%d",
    "%Y/%m/%d",
    "%m/%d/%Y",
    "%d/%m/%Y",
    "%m-%d-%Y",
    "%d-%m-%Y",
    "%m/%d/%y",
    "%d/%m/%y",
    "%Y.%m.%d",
)


@dataclass
class BatchOutcome:
    messages: list = field(default_factory=list)  # (kind, text) tuples
    processed: int = 0
    skipped: int = 0
    created_folder: bool = False
    # One dict per photo (name, status, reason, details); status is "added", "review",
    # "skipped" or "failed". The Batch results panel draws its rows from this list.
    files: list = field(default_factory=list)
    # Messages that are not about one photo (folder created, slides generated, ...)
    notices: list = field(default_factory=list)

    def add_file(self, name, status, reason, details=""):
        self.files.append({"name": name, "status": status, "reason": reason, "details": details})

    def notice(self, kind, text):
        """A general message: kept in `messages` as before, and also listed in `notices`."""
        self.messages.append((kind, text))
        self.notices.append((kind, text))


def _parse_date(date_raw):
    """Parses the AI-extracted Date string into (sort_date, day_key, display_date).
    Valid dates sort chronologically and group by calendar day; missing/unparseable
    dates sort after valid dates."""
    raw = str(date_raw or "").strip()
    if not raw:
        return datetime.date.max, "9999-12-31", "Unknown Date"

    # Strip ordinal suffixes (1st, 2nd, 3rd, 4th) and normalize whitespace
    cleaned = re.sub(r"(\d+)(?:st|nd|rd|th)\b", r"\1", raw, flags=re.IGNORECASE)
    cleaned = re.sub(r"\s+", " ", cleaned).strip(" ,.-")

    for fmt in _DATE_FORMATS:
        try:
            dt = datetime.datetime.strptime(cleaned, fmt).date()
            display = f"{dt.strftime('%B')} {dt.day}, {dt.year}"
            return dt, dt.isoformat(), display
        except ValueError:
            continue

    return datetime.date.max, f"9999-12-31:{raw.lower()}", raw


def sort_and_chunk_photos(photo_records, max_per_chunk=MAX_PHOTOS_PER_CHUNK):
    """Sorts processed photo records chronologically by their AI-extracted Date,
    groups them by Date (Day), and splits each day's photos into chunks of up to
    `max_per_chunk` (default 6) photos per slide."""
    sorted_records = sorted(
        photo_records,
        key=lambda item: (item["sort_date"], item["day_key"], item["order_index"]),
    )

    grouped_by_day = OrderedDict()
    for record in sorted_records:
        day_key = record["day_key"]
        grouped_by_day.setdefault(day_key, []).append(record)

    daily_chunks = []
    for day_key, day_photos in grouped_by_day.items():
        day_label = day_photos[0]["date_display"]
        for start in range(0, len(day_photos), max_per_chunk):
            chunk_slice = day_photos[start : start + max_per_chunk]
            daily_chunks.append({
                "day_key": day_key,
                "date": day_label,
                "photos": chunk_slice,
            })

    return daily_chunks


def _create_folder(subfolder, outcome):
    """Creates the new SharePoint folder first, so it exists before uploading the presentation."""
    try:
        sharepoint_client.ensure_subfolder(subfolder)
        outcome.created_folder = True
        outcome.notice("status", f'Created the folder "{subfolder}" in SharePoint.')
    except Exception:
        traceback.print_exc()
        outcome.notice(
            "warning",
            f'Could not create the folder "{subfolder}" in SharePoint just now. '
            "The photos are still being processed and saved locally; syncing will be tried when the presentation is ready.",
        )


def _review_message(filename, parsed_data):
    """The warning shown for a photo whose automatic checks found something, or None."""
    notes = parsed_data.get("Review Notes", [])
    if not notes:
        return None
    shown = "; ".join(notes[:3])
    extra = f" (+{len(notes) - 3} more)" if len(notes) > 3 else ""
    if parsed_data.get("Flagged Materials"):
        return (
            "warning",
            f"'{filename}' was added to the PowerPoint report, but needs manual review: {shown}{extra}",
        )
    return (
        "warning",
        f"'{filename}' was added to the PowerPoint report. FYI: {shown}{extra}",
    )


def _build_file_status(name, parsed_data):
    """Builds the per-photo status dict for the Batch results UI."""
    notes = parsed_data.get("Review Notes", [])
    flagged = bool(notes) and bool(parsed_data.get("Flagged Materials"))
    extra = f" (+{len(notes) - 1} more)" if len(notes) > 1 else ""
    if flagged:
        status, reason = "review", f"Flagged for review in presentation: {notes[0]}{extra}"
    elif notes:
        status, reason = "added", f"Added to presentation. FYI: {notes[0]}{extra}"
    else:
        status, reason = "added", "Added to the PowerPoint report"
    details = "\n".join(notes)
    return {"name": name, "status": status, "reason": reason, "details": details}


def _extract_photo_record(uploaded_file, file_bytes, file_hash, order_index, outcome):
    """Runs AI extraction on one photo and returns its structured record, or None on error."""
    extension = os.path.splitext(uploaded_file.name)[1].lower() or ".jpg"
    temp_path = f"temp_{uuid.uuid4().hex}{extension}"
    with open(temp_path, "wb") as f:
        f.write(file_bytes)

    try:
        parsed_data = extract_data(temp_path)
        sort_date, day_key, date_display = _parse_date(parsed_data.get("Date", ""))
        drop_val = str(parsed_data.get("Drop", "") or "").strip()
        floor_val = str(parsed_data.get("Floor", "") or "").strip()
        defect_val = format_defect_text(parsed_data)

        return {
            "name": uploaded_file.name,
            "file_hash": file_hash,
            "order_index": order_index,
            "image_bytes": file_bytes,
            "parsed_data": parsed_data,
            "sort_date": sort_date,
            "day_key": day_key,
            "date": str(parsed_data.get("Date", "") or "").strip(),
            "date_display": date_display,
            "drop": drop_val,
            "floor": floor_val,
            "defect": defect_val,
        }
    except Exception as e:
        traceback.print_exc()
        outcome.messages.append(("error", f"An error occurred while processing {uploaded_file.name}: {e}"))
        short = str(e).strip().splitlines()[0][:90] if str(e).strip() else type(e).__name__
        outcome.add_file(uploaded_file.name, "failed", f"Extraction error: {short}", str(e))
        return None
    finally:
        if os.path.exists(temp_path):
            os.remove(temp_path)


def run_batch(uploaded_files, output_path, subfolder="", creating_new_folder=False, on_progress=None, title_block_info=None):
    """Extracts every photo in `uploaded_files`, sorts them chronologically by Date,
    groups by Day into chunks of max 6 photos, writes the PowerPoint presentation
    to `output_path`, and syncs to SharePoint."""
    outcome = BatchOutcome()
    total = len(uploaded_files)

    def notify(i, name, finished):
        if on_progress:
            on_progress(i, total, name, finished)

    if ENABLE_SHAREPOINT_SYNC and creating_new_folder:
        _create_folder(subfolder, outcome)
    elif ENABLE_SHAREPOINT_SYNC:
        # Make sure the landing folder ("Image sorter") and the chosen subfolder exist before
        # the presentation is uploaded. Existing folders are left untouched.
        try:
            sharepoint_client.ensure_subfolder(subfolder)
        except Exception:
            traceback.print_exc()

    processed_log = load_processed_log()
    seen_hashes = set()
    extracted_records = []

    for i, uploaded_file in enumerate(uploaded_files):
        notify(i, uploaded_file.name, finished=False)

        try:
            file_bytes = uploaded_file.getvalue()
        except Exception as e:
            traceback.print_exc()
            outcome.messages.append(("error", f"Could not read '{uploaded_file.name}': {e}. It was skipped."))
            outcome.add_file(uploaded_file.name, "failed", "Could not read the photo", str(e))
            notify(i, uploaded_file.name, finished=True)
            continue

        file_hash = hashlib.sha256(file_bytes).hexdigest()

        if file_hash in processed_log or file_hash in seen_hashes:
            outcome.messages.append(("warning", f"Skipping '{uploaded_file.name}' - already processed."))
            outcome.add_file(
                uploaded_file.name,
                "skipped",
                "Already processed",
                "This photo's content matches one processed before (duplicates are matched by content, not by file name).",
            )
            outcome.skipped += 1
            notify(i, uploaded_file.name, finished=True)
            continue
        seen_hashes.add(file_hash)

        record = _extract_photo_record(uploaded_file, file_bytes, file_hash, i, outcome)
        if record is not None:
            extracted_records.append(record)

        notify(i, uploaded_file.name, finished=True)

    if not extracted_records:
        return outcome

    # Sort chronologically by Date, group by Day, and split into chunks of max 6 photos
    daily_chunks = sort_and_chunk_photos(extracted_records, max_per_chunk=MAX_PHOTOS_PER_CHUNK)
    write_chunks_to_pptx(daily_chunks, output_path, title_block_info=title_block_info)

    sync_failed = False
    if ENABLE_SHAREPOINT_SYNC:
        synced, _reason = sync_tracker.sync_file(output_path, subfolder)
        if not synced:
            sync_failed = True
            outcome.notice(
                "warning",
                f"'{os.path.basename(output_path)}' was saved locally, but syncing to SharePoint failed "
                "just now. You can push it from 'Pending syncs' at the bottom of the page.",
            )

    for record in extracted_records:
        parsed_data = record["parsed_data"]
        message = _review_message(record["name"], parsed_data)
        if message:
            outcome.messages.append(message)

        file_row = _build_file_status(record["name"], parsed_data)
        if sync_failed:
            file_row["reason"] += " (SharePoint sync pending)"
            file_row["details"] = (file_row["details"] + "\n" if file_row["details"] else "") + (
                "Saved locally, but syncing to SharePoint failed just now. You can push it from 'Pending syncs'."
            )
        outcome.add_file(
            file_row["name"],
            file_row["status"],
            file_row["reason"],
            file_row["details"],
        )
        processed_log.append(record["file_hash"])
        outcome.processed += 1

    save_processed_log(processed_log)
    outcome.notice(
        "status",
        f"Generated {len(daily_chunks)} slide(s) across {len({c['day_key'] for c in daily_chunks})} day(s).",
    )
    return outcome
