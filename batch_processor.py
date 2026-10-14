# batch_processor.py
"""Processes a batch of photos into one master Excel file.
Returns messages instead of drawing them, so app.py decides how to show them.

A "photo" is any object with a .name and a .getvalue() that returns the image
bytes: a Streamlit UploadedFile, or app.py's RemoteFile (which downloads the
photo from SharePoint at the moment getvalue() is called)."""
import os
import uuid
import hashlib
import traceback
from dataclasses import dataclass, field

import sharepoint_client
import sync_tracker
from extraction import extract_data
from excel_writer import write_parsed_data_to_excel
from file_utils import load_processed_log, save_processed_log
from settings import ENABLE_SHAREPOINT_SYNC


@dataclass
class BatchOutcome:
    messages: list = field(default_factory=list)  # (kind, text) tuples
    processed: int = 0
    skipped: int = 0
    created_folder: bool = False
    # One dict per photo (name, status, reason, details); status is "added", "review",
    # "skipped" or "failed". The Batch results panel draws its rows from this list.
    files: list = field(default_factory=list)
    # Messages that are not about one photo (folder created, existing file loaded, ...)
    notices: list = field(default_factory=list)

    def add_file(self, name, status, reason, details=""):
        self.files.append({"name": name, "status": status, "reason": reason, "details": details})

    def notice(self, kind, text):
        """A general message: kept in `messages` as before, and also listed in `notices`."""
        self.messages.append((kind, text))
        self.notices.append((kind, text))


def _create_folder(subfolder, outcome):
    """Creates the new SharePoint folder first, so it exists before the first upload."""
    try:
        sharepoint_client.ensure_subfolder(subfolder)
        outcome.created_folder = True
        outcome.notice("status", f'Created the folder "{subfolder}" in SharePoint.')
    except Exception:
        traceback.print_exc()
        outcome.notice(
            "warning",
            f'Could not create the folder "{subfolder}" in SharePoint just now. '
            "The photos are still being processed and saved in the app; syncing will be tried again after each photo."
        )


def _download_existing(remote_filename, output_path, subfolder, outcome):
    """The local disk is temporary, SharePoint persists: pull down any existing
    master file first so new entries are appended instead of overwriting it."""
    try:
        if sharepoint_client.download_file(remote_filename, output_path, subfolder):
            sync_tracker.mark_synced(output_path)  # identical to SharePoint right now
            outcome.notice("status", "Found an existing master file in SharePoint and loaded it before continuing.")
    except Exception:
        traceback.print_exc()
        outcome.notice(
            "warning",
            "Could not check SharePoint for an existing master file before starting. "
            "If one already exists there, please verify afterward that no data was lost."
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
            f"'{filename}' was added to the Excel file, but its data row is "
            f"highlighted in RED for manual validation: {shown}{extra}"
        )
    return (
        "warning",
        f"'{filename}' was added to the Excel file. FYI only, nothing is "
        f"highlighted red: {shown}{extra}"
    )


def _record_added(outcome, name, parsed_data, sync_failed):
    """Adds the row for a photo that reached the Excel file: "review" when its data row was
    highlighted red for manual checking, otherwise "added" (with the FYI note, if any)."""
    notes = parsed_data.get("Review Notes", [])
    flagged = bool(notes) and bool(parsed_data.get("Flagged Materials"))
    extra = f" (+{len(notes) - 1} more)" if len(notes) > 1 else ""
    if flagged:
        status, reason = "review", f"Row highlighted red in Excel: {notes[0]}{extra}"
    elif notes:
        status, reason = "added", f"Added. FYI: {notes[0]}{extra}"
    else:
        status, reason = "added", "Added to the Excel file"
    details = "\n".join(notes)
    if sync_failed:
        reason += " (SharePoint sync pending)"
        details = (details + "\n" if details else "") + (
            "Saved locally, but syncing to SharePoint failed just now. It is retried after the next "
            "photo, or you can push it from 'Pending syncs'.")
    outcome.add_file(name, status, reason, details)


def _process_photo(uploaded_file, file_bytes, output_path, subfolder, outcome):
    """Extracts one photo, writes it to Excel, then syncs. Returns True on success."""
    extension = os.path.splitext(uploaded_file.name)[1].lower() or ".jpg"
    temp_path = f"temp_{uuid.uuid4().hex}{extension}"
    with open(temp_path, "wb") as f:
        f.write(file_bytes)

    try:
        parsed_data = extract_data(temp_path)
        write_parsed_data_to_excel(parsed_data, temp_path, output_path)

        # Best-effort: the data is already saved locally, so a SharePoint hiccup
        # never loses or blocks it. A failed upload is recorded, and the file then
        # shows up under "Pending syncs" in the app until it is pushed.
        sync_failed = False
        if ENABLE_SHAREPOINT_SYNC:
            synced, _reason = sync_tracker.sync_file(output_path, subfolder)
            if not synced:
                sync_failed = True
                outcome.messages.append((
                    "warning",
                    f"'{uploaded_file.name}' was saved locally, but syncing to SharePoint failed "
                    "just now. It is retried after the next photo, or you can push it yourself "
                    "from 'Pending syncs' at the bottom of the page."
                ))

        message = _review_message(uploaded_file.name, parsed_data)
        if message:
            outcome.messages.append(message)
        _record_added(outcome, uploaded_file.name, parsed_data, sync_failed)
        return True

    except Exception as e:
        # Full traceback goes to the app logs; the person sees the short message.
        traceback.print_exc()
        outcome.messages.append(("error", f"An error occurred while processing {uploaded_file.name}: {e}"))
        short = str(e).strip().splitlines()[0][:90] if str(e).strip() else type(e).__name__
        outcome.add_file(uploaded_file.name, "failed", f"Extraction error: {short}", str(e))
        return False
    finally:
        if os.path.exists(temp_path):
            os.remove(temp_path)


def run_batch(uploaded_files, output_path, subfolder="", creating_new_folder=False, on_progress=None):
    """Processes every photo into output_path.

    on_progress(index, total, filename, finished) is called before (finished=False)
    and after (finished=True) each photo, so the UI can update its progress bar."""
    outcome = BatchOutcome()
    remote_filename = os.path.basename(output_path)
    total = len(uploaded_files)

    def notify(i, name, finished):
        if on_progress:
            on_progress(i, total, name, finished)

    if ENABLE_SHAREPOINT_SYNC and creating_new_folder:
        _create_folder(subfolder, outcome)
    if ENABLE_SHAREPOINT_SYNC and not os.path.exists(output_path):
        _download_existing(remote_filename, output_path, subfolder, outcome)

    # Photos are matched by content, not filename, so a renamed re-upload is still a duplicate
    processed_log = load_processed_log()
    seen_hashes = set()

    for i, uploaded_file in enumerate(uploaded_files):
        notify(i, uploaded_file.name, finished=False)

        # For photos coming from SharePoint this is where the download happens, one photo
        # at a time. If it fails, that photo is reported and the batch carries on.
        try:
            file_bytes = uploaded_file.getvalue()
        except Exception as e:
            traceback.print_exc()
            outcome.messages.append((
                "error", f"Could not read '{uploaded_file.name}': {e}. It was skipped."
            ))
            outcome.add_file(uploaded_file.name, "failed", "Could not read the photo", str(e))
            notify(i, uploaded_file.name, finished=True)
            continue

        file_hash = hashlib.sha256(file_bytes).hexdigest()

        if file_hash in processed_log or file_hash in seen_hashes:
            outcome.messages.append(("warning", f"Skipping '{uploaded_file.name}' - already processed."))
            outcome.add_file(
                uploaded_file.name, "skipped", "Already processed",
                "This photo's content matches one processed before (duplicates are matched by content, not by file name).")
            outcome.skipped += 1
            notify(i, uploaded_file.name, finished=True)
            continue
        seen_hashes.add(file_hash)

        if _process_photo(uploaded_file, file_bytes, output_path, subfolder, outcome):
            processed_log.append(file_hash)
            save_processed_log(processed_log)
            outcome.processed += 1

        notify(i, uploaded_file.name, finished=True)

    return outcome
