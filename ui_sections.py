"""Reusable Streamlit sections for the Whiteboard app.
(Header and CSS stay in ui_components.py.)"""
import io
import os
import html
import hashlib
import traceback
import ui_components
import streamlit as st
from PIL import Image, ImageOps

import sharepoint_client
import sync_tracker
from settings import ENABLE_SHAREPOINT_SYNC, MAIN_FOLDER_LABEL, NEW_FOLDER_LABEL

# The one final action, whichever photo source is used
PROCESS_BUTTON_LABEL = "Extract data"

_SESSION_DEFAULTS = {
    "uploader_key": 0,             # changing it resets the uploader (clears all photos)
    "batch_results": [],
    "batch_report": None,          # per-photo rows of the last batch (see show_batch_results)
    "confirm_reset": False,
    "last_output_path": None,
    "last_output_label": None,
    "last_output_folder": "",      # "" = the main landing folder
    "last_folder_choice": MAIN_FOLDER_LABEL,
    "sp_folders": None,            # None = not loaded yet
    "sp_folders_error": False,
}


def init_session_state():
    for key, default in _SESSION_DEFAULTS.items():
        if key not in st.session_state:
            st.session_state[key] = list(default) if isinstance(default, list) else default


# With more messages than this, the old plain message list goes into a scrollable box
RESULTS_BOX_MIN_MESSAGES = 5
RESULTS_BOX_HEIGHT_PX = 320
# app.py always ends a batch with two messages ("Batch Complete." and the summary
# line); they stay outside the box so the outcome is visible without scrolling.
RESULTS_PINNED_MESSAGES = 2

# The compact result rows scroll inside a box once there are more than this many
RESULT_ROWS_BOX_MIN = 8
RESULT_ROWS_BOX_HEIGHT_PX = 380

# Short icons for the result rows: green check, orange alert, gray skip, red cross
_ROW_ICONS = {
    "added": '<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="9"/><path d="m8.5 12.3 2.4 2.4 4.6-5"/></svg>',
    "review": '<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"><path d="M12 4 2.8 19.5h18.4L12 4Z"/><path d="M12 10v4.2M12 17h.01"/></svg>',
    "skipped": '<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="9"/><path d="M8 12h8"/></svg>',
    "failed": '<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="9"/><path d="m9 9 6 6m0-6-6 6"/></svg>',
}
_ROW_LABELS = {"added": "Added", "review": "Needs review", "skipped": "Skipped", "failed": "Failed"}
# Rows that need attention come first
_ROW_ORDER = {"failed": 0, "review": 1, "skipped": 2, "added": 3}


def _show_message(kind, text):
    if kind == "warning":
        st.warning(text)
    elif kind == "error":
        st.error(text)
    elif kind == "success":
        st.success(text)
    else:
        st.markdown(f'<p class="status-msg">{text}</p>', unsafe_allow_html=True)


def _short_name(name, limit=32):
    """The file name without its folder path, shortened in the middle if it is long
    (the extension stays visible)."""
    base = os.path.basename(str(name).replace("\\", "/"))
    if len(base) <= limit:
        return base
    stem, ext = os.path.splitext(base)
    keep = max(limit - len(ext) - 1, 8)
    return f"{stem[:keep]}\u2026{ext}"


def _summary_html(files):
    """"16 processed - 15 added - 1 needs review", with the counts coloured."""
    count = {status: 0 for status in _ROW_LABELS}
    for item in files:
        count[item["status"]] = count.get(item["status"], 0) + 1
    parts = [f'<span class="rs-total">{count["added"] + count["review"]} processed</span>']
    if count["added"]:
        parts.append(f'<span class="rs-added">{count["added"]} added</span>')
    if count["review"]:
        parts.append(f'<span class="rs-review">{count["review"]} needs review</span>')
    if count["skipped"]:
        parts.append(f'<span class="rs-skipped">{count["skipped"]} duplicate{"s" if count["skipped"] != 1 else ""} skipped</span>')
    if count["failed"]:
        parts.append(f'<span class="rs-failed">{count["failed"]} failed</span>')
    return '<div class="rs-summary">' + '<span class="rs-sep">&middot;</span>'.join(parts) + '</div>'


def _row_html(item):
    """One compact row: icon, short file name, result chip, one-line reason. Rows with more
    to say (everything except a plain "Added") open to show the full details."""
    status = item["status"] if item["status"] in _ROW_LABELS else "added"
    full_name = str(item["name"])
    short = _short_name(full_name)
    reason = str(item.get("reason") or "")
    head = (
        f'<span class="rs-icon rs-{status}">{_ROW_ICONS[status]}</span>'
        f'<span class="rs-name" title="{html.escape(full_name)}">{html.escape(short)}</span>'
        f'<span class="rs-chip rs-chip-{status}">{_ROW_LABELS[status]}</span>'
        f'<span class="rs-reason">{html.escape(reason)}</span>'
    )
    details = [line for line in str(item.get("details") or "").splitlines() if line.strip()]
    if status == "added" and not details and short == full_name:
        return f'<div class="rs-row rs-row-{status}"><div class="rs-head">{head}</div></div>'
    body = f'<div class="rs-line"><b>File:</b> {html.escape(full_name)}</div>'
    if reason:
        body += f'<div class="rs-line">{html.escape(reason)}</div>'
    if details:
        body += '<ul>' + ''.join(f'<li>{html.escape(line)}</li>' for line in details) + '</ul>'
    return (
        f'<details class="rs-row rs-row-{status}"><summary class="rs-head">{head}</summary>'
        f'<div class="rs-details">{body}</div></details>'
    )


def _show_report(report):
    """The compact version of a batch's results: a summary line, any general notices, then
    one row per photo (the ones that need attention first)."""
    files = report.get("files") or []
    if files:
        st.markdown(_summary_html(files), unsafe_allow_html=True)
    for kind, text in report.get("notices") or []:
        if kind in ("warning", "error"):
            _show_message(kind, text)
        else:
            st.caption(text)
    if not files:
        return
    rows = sorted(files, key=lambda item: _ROW_ORDER.get(item["status"], 9))
    rows_html = '<div class="rs-list">' + ''.join(_row_html(item) for item in rows) + '</div>'
    if len(rows) > RESULT_ROWS_BOX_MIN:
        with st.container(height=RESULT_ROWS_BOX_HEIGHT_PX, border=False, key="batch_files_box"):
            st.markdown(rows_html, unsafe_allow_html=True)
    else:
        st.markdown(rows_html, unsafe_allow_html=True)


def show_batch_results():
    """Shows the last batch (kept even after the uploader is cleared): a compact summary
    and one row per photo when the batch recorded them, otherwise the plain message list
    (a long one goes inside a scrollable box)."""
    results = st.session_state.batch_results
    if not results:
        return
    report = st.session_state.get("batch_report")
    if report and (report.get("files") or report.get("notices")):
        _show_report(report)
        return

    if len(results) < RESULTS_BOX_MIN_MESSAGES:
        for kind, text in results:
            _show_message(kind, text)
        return

    scrolling = results[:-RESULTS_PINNED_MESSAGES]
    pinned = results[-RESULTS_PINNED_MESSAGES:]
    with st.container(height=RESULTS_BOX_HEIGHT_PX, border=True, key="batch_results_box"):
        for kind, text in scrolling:
            _show_message(kind, text)
    for kind, text in pinned:
        _show_message(kind, text)


def status_message(text):
    st.markdown(f'<p class="status-msg">{text}</p>', unsafe_allow_html=True)


# ---------------------------------------------------------------------------
# Uploader: hide the empty-state chrome once photos are chosen
# ---------------------------------------------------------------------------

_UPLOADER_FILLED_CSS = (
    "<style>"
    "[data-testid='stFileUploaderDropzone']::before, "
    "[data-testid='stFileUploadDropzone']::before, "
    "[data-testid='stFileUploaderDropzone']::after, "
    "[data-testid='stFileUploadDropzone']::after, "
    "[data-testid='stFileUploaderDropzoneInstructions'] "
    "{ display: none !important; } "
    "[data-testid='stFileUploaderDropzone'], [data-testid='stFileUploadDropzone'] "
    "{ min-height: 0 !important; padding: 12px 14px !important; flex-direction: column !important; align-items: flex-start !important; cursor: default !important; } "
    "div[data-testid='stFileUploader'] [data-testid='stFileUploaderDropzone'] > button[data-testid='stBaseButton-secondary'], "
    "div[data-testid='stFileUploader'] [data-testid='stFileUploaderDropzone'] > button "
    "{ position: static !important; inset: auto !important; width: auto !important; height: 34px !important; padding: 0 16px !important; "
    "opacity: 1 !important; font-size: 0 !important; background: #ffffff !important; border: 1px solid #cbd5e1 !important; border-radius: 7px !important; box-shadow: none !important; } "
    "div[data-testid='stFileUploader'] [data-testid='stFileUploaderDropzone'] > button * { display: none !important; } "
    "div[data-testid='stFileUploader'] [data-testid='stFileUploaderDropzone'] > button::after "
    "{ content: '+ Add more'; font-size: 13px; font-weight: 700; color: #1e3a8a; } "
    "</style>"
)


def hide_uploader_chrome_if_filled(uploaded_files):
    """Call right after st.file_uploader(...). Does nothing for an empty uploader."""
    if uploaded_files:
        st.markdown(_UPLOADER_FILLED_CSS, unsafe_allow_html=True)


# ---------------------------------------------------------------------------
# Photo previews
# ---------------------------------------------------------------------------

PREVIEW_SCROLL_HEIGHT = 290  # px: more than one row of thumbnails scrolls inside a box this tall
PREVIEW_COLUMNS = 6       # thumbnails per row
PREVIEW_PIXELS = 240      # longest side of a thumbnail


@st.cache_data(show_spinner=False, max_entries=300)
def _thumbnail_bytes(name, size, file_id, _data):
    """A small JPEG of one uploaded photo, or None if it can't be read. Phone
    photos are large, so only this small copy is sent to the browser."""
    try:
        with Image.open(io.BytesIO(_data)) as img:
            img.draft("RGB", (PREVIEW_PIXELS * 2, PREVIEW_PIXELS * 2))  # fast, rough decode for JPEGs
            img = ImageOps.exif_transpose(img)  # phones store "rotated" photos with a flag
            img.thumbnail((PREVIEW_PIXELS, PREVIEW_PIXELS))
            out = io.BytesIO()
            img.convert("RGB").save(out, format="JPEG", quality=80)
            return out.getvalue()
    except Exception:
        return None


def _format_size(num_bytes):
    """12 345 678 bytes -> "11.8 MB"."""
    size = float(num_bytes)
    for unit in ("B", "KB", "MB"):
        if size < 1024:
            return f"{size:.0f} {unit}" if unit == "B" else f"{size:.1f} {unit}"
        size /= 1024
    return f"{size:.1f} GB"


def render_photo_previews(uploaded_files):
    """A count and a grid of small previews of the selected photos."""
    if not uploaded_files:
        return

    photos = list(uploaded_files)
    total_bytes = sum(getattr(photo, "size", 0) or 0 for photo in photos)
    
    st.markdown(
        f'<div style="color:#16865b; font-size:12px; font-weight:700; margin-bottom:8px;">'
        f'<span style="display:inline-block;width:7px;height:7px;border-radius:50%;background:#16865b;margin-right:6px;"></span>'
        f'{len(photos)} photo{"s" if len(photos) != 1 else ""} selected &middot; {_format_size(total_bytes)}</div>',
        unsafe_allow_html=True
    )

    def draw_grid():
        for start in range(0, len(photos), PREVIEW_COLUMNS):
            columns = st.columns(PREVIEW_COLUMNS)
            for column, photo in zip(columns, photos[start:start + PREVIEW_COLUMNS]):
                thumbnail = _thumbnail_bytes(photo.name, photo.size, getattr(photo, "file_id", ""), photo.getvalue())
                with column:
                    if thumbnail:
                        st.image(thumbnail)
                    else:
                        st.caption("(no preview)")

    if len(photos) > PREVIEW_COLUMNS:
        with st.container(height=PREVIEW_SCROLL_HEIGHT, border=False, key="thumbnail_scroll_gallery"):
            draw_grid()
    else:
        draw_grid()


# ---------------------------------------------------------------------------
# SharePoint folder picker
# ---------------------------------------------------------------------------

def refresh_folder_list():
    """Reads the current subfolders from SharePoint. On failure the app carries
    on with an empty list (main folder + creating a new folder still work)."""
    try:
        st.session_state.sp_folders = sharepoint_client.list_subfolders()
        st.session_state.sp_folders_error = False
    except Exception:
        traceback.print_exc()
        st.session_state.sp_folders = []
        st.session_state.sp_folders_error = True


def render_folder_picker(source_chosen=True, photo_count=0):
    """Returns (folder_choice, new_folder_raw).
    Mutes settings area if source isn't chosen yet (Progressive disclosure)."""
    if not ENABLE_SHAREPOINT_SYNC:
        return MAIN_FOLDER_LABEL, ""

    disabled_state = not source_chosen and not photo_count
    
    with st.container(key="save_settings_block"):
        # We manually render the title because the container controls the inner padding
        ui_components.card_heading("OUTPUT PREFERENCES", "Save settings")
        
        if disabled_state:
            st.markdown(
                '<p style="color:#64748b; font-size:12.5px; margin: 4px 0 0 0;">'
                'Please select photos or a SharePoint source folder above to unlock destination settings.</p>',
                unsafe_allow_html=True
            )
            return MAIN_FOLDER_LABEL, ""

        if st.session_state.sp_folders is None:
            with st.spinner("Loading SharePoint folders..."):
                refresh_folder_list()
        if st.session_state.sp_folders_error:
            st.warning(
                "Could not load the folder list from SharePoint right now. You can still save to "
                "the main folder, or create a new folder."
            )

        col1, col2 = st.columns(2)
        options = [MAIN_FOLDER_LABEL] + list(st.session_state.sp_folders) + [NEW_FOLDER_LABEL]
        default = st.session_state.last_folder_choice
        
        with col1:
            folder_choice = st.selectbox(
                "Save into which SharePoint folder?",
                options,
                index=options.index(default) if default in options else 0,
                key="folder_choice_input",
                disabled=disabled_state,
            )

        new_folder_raw = ""
        with col2:
            if folder_choice == NEW_FOLDER_LABEL:
                new_folder_raw = st.text_input("Name for the new folder", key="new_folder_name_input", disabled=disabled_state)
            else:
                st.text_input("Batch name", value="Auto-generated based on date", disabled=True)

        return folder_choice, new_folder_raw


def resolve_subfolder(folder_choice, new_folder_raw):
    """Turns the picker's answer into (subfolder, creating_new_folder, error_message)."""
    if not ENABLE_SHAREPOINT_SYNC:
        return "", False, None

    if folder_choice == NEW_FOLDER_LABEL:
        typed = sharepoint_client.sanitize_folder_name(new_folder_raw)
        if not typed:
            return "", False, "Please enter a name for the new folder before extracting."
        # A name matching an existing folder (ignoring capitals) just uses that folder
        match = next((f for f in st.session_state.sp_folders if f.lower() == typed.lower()), None)
        return (match or typed), match is None, None

    if folder_choice != MAIN_FOLDER_LABEL:
        return folder_choice, False, None
    return "", False, None


def remember_new_folder(subfolder):
    """Adds a just-created folder to the dropdown list without re-reading SharePoint."""
    folders = st.session_state.sp_folders or []
    if subfolder.lower() not in [f.lower() for f in folders]:
        st.session_state.sp_folders = sorted(folders + [subfolder], key=str.lower)


# ---------------------------------------------------------------------------
# Buttons and file management
# ---------------------------------------------------------------------------

def process_blockers(photo_count, source_chosen, folder_choice, new_folder_raw, batch_name_raw, remote):
    """Returns (photos_missing, destination_missing): a short sentence for each thing that is
    still missing, or "" when it is fine. The Process button stays disabled until both are ""."""
    photos_missing = ""
    if remote and not source_chosen:
        photos_missing = "Select a SharePoint folder to continue"
    elif not remote and not photo_count:
        photos_missing = "Select local photos to continue"
    destination_missing = ""
    if folder_choice == NEW_FOLDER_LABEL and not (new_folder_raw or "").strip():
        destination_missing = "Enter a name for the new folder"
    elif not (batch_name_raw or "").strip():
        destination_missing = "Enter a batch name"
    return photos_missing, destination_missing


def _process_label(photo_count, remote, blocked_reason=""):
    if blocked_reason:
        return "Select a folder" if remote else "Select photos"
    if remote:
        return f"Extract {photo_count} photos" if photo_count else "Extract SharePoint photos"
    if photo_count:
        return f"Extract {photo_count} photo{'s' if photo_count != 1 else ''}"
    return "Extract data"


def render_action_buttons(photo_count=0, blocked_reason="", remote=False):
    """The sticky action footer (style.css pins it to the bottom of the left panel): the
    Process button, whose text carries the context ("Extract 16 photos"). 
    The requirement text sits directly beside the disabled button."""
    with st.container(key="action_buttons"):
        label = _process_label(photo_count, remote, blocked_reason)
        
        # Align buttons right, hint left
        col_hint, col_btn_clear, col_btn_primary = st.columns([1, 0.4, 0.4])
        
        with col_hint:
            if blocked_reason:
                st.markdown(
                    f'<p style="color:#d85f0d; font-size:12.5px; font-weight:600; padding-top:12px;">'
                    f'{html.escape(blocked_reason)}</p>', 
                    unsafe_allow_html=True
                )
        
        with col_btn_clear:
            if not remote:
                clear_clicked = st.button("Clear All", key="clear_photos", use_container_width=True)
            else:
                clear_clicked = False
                
        with col_btn_primary:
            if remote:
                process_clicked = st.button(label, type="primary", key="fetch_folder", disabled=bool(blocked_reason), use_container_width=True)
            else:
                process_clicked = st.button(label, type="primary", key="extract_data", disabled=bool(blocked_reason), use_container_width=True)
                
    return process_clicked, clear_clicked


def file_management_title_html(label, folder):
    """The heading of the file card: a small label over "Generated workbook", then the
    batch name and folder (which can be long) in lighter gray."""
    folder_note = f' in folder "{html.escape(str(folder))}"' if folder else ""
    return (
        '<p class="eyebrow">File management</p>'
        '<div class="card-title">Generated workbook</div>'
        f'<div class="file-sub">"{html.escape(str(label))}"{folder_note}</div>'
    )


def render_file_management():
    """Download / delete controls for the file of the last batch, in a card."""
    path = st.session_state.last_output_path
    if not (path and os.path.exists(path)):
        return

    with st.container(border=True, key="file_card"):
        st.markdown(
            file_management_title_html(st.session_state.last_output_label, st.session_state.last_output_folder),
            unsafe_allow_html=True,
        )

        col1, col2 = st.columns(2)
        with col1:
            try:
                with open(path, "rb") as f:
                    xlsx_bytes = f.read()
            except FileNotFoundError:
                xlsx_bytes = None
            if xlsx_bytes is not None:
                st.download_button(
                    label="Download master log",
                    data=xlsx_bytes,
                    file_name=os.path.basename(path),
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary",
                )
        with col2:
            if st.button("Delete batch file", type="primary", key="start_fresh"):
                st.session_state.confirm_reset = True
                st.rerun()

        if st.session_state.confirm_reset:
            st.warning(
                f'This will permanently delete the local copy of "{os.path.basename(path)}". It does not '
                "touch the processed-photos duplicate log, and it does not delete any copy already "
                "synced to SharePoint, you can remove that there directly if needed. Are you sure?"
            )
            yes_col, cancel_col = st.columns(2)
            with yes_col:
                if st.button("Yes, delete this file", type="primary", key="confirm_delete"):
                    if os.path.exists(path):
                        os.remove(path)
                    sync_tracker.forget(path)
                    st.session_state.last_output_path = None
                    st.session_state.last_output_label = None
                    st.session_state.last_output_folder = ""
                    st.session_state.confirm_reset = False
                    st.rerun()
            with cancel_col:
                if st.button("Cancel", type="primary", key="cancel_delete"):
                    st.session_state.confirm_reset = False
                    st.rerun()


# ---------------------------------------------------------------------------
# Pending Syncs: batch files that never reached SharePoint
# ---------------------------------------------------------------------------

def _push_pending(item):
    """Uploads one stranded file and leaves a message for the next screen draw."""
    with st.spinner(f"Pushing {item['filename']} to SharePoint..."):
        ok, reason = sync_tracker.sync_file(item["path"], item["subfolder"])
    if ok:
        st.session_state.sync_notice = ("success", f"'{item['filename']}' is now uploaded to SharePoint.")
    else:
        st.session_state.sync_notice = (
            "error", f"Could not push '{item['filename']}' to SharePoint: {reason}. You can try again."
        )
    st.rerun()


def render_pending_syncs():
    """Lists batch files that are not up to date in SharePoint, each with a
    Push button. Draws nothing when everything is synced."""
    if not ENABLE_SHAREPOINT_SYNC:
        return

    notice = st.session_state.pop("sync_notice", None)
    if notice:
        (st.success if notice[0] == "success" else st.error)(notice[1])

    pending = sync_tracker.pending_files()
    if not pending:
        return

    st.markdown("<br>", unsafe_allow_html=True)
    st.subheader("Pending syncs")
    st.warning(
        f"{len(pending)} file(s) are not up to date in SharePoint. Pushing a file replaces the "
        "SharePoint copy with the one stored here. These files are kept only on the app's "
        "temporary storage and are lost if the app restarts, so download any you want to keep."
    )

    for item in pending:
        folder_note = f' in folder "{item["subfolder"]}"' if item["subfolder"] else ""
        detail = item["reason"]
        if item["last_error"]:
            detail += f' (last attempt {item["last_attempt"]}: {item["last_error"]})'

        info_col, push_col, download_col = st.columns([3, 1.3, 1.3])
        with info_col:
            st.markdown(
                f'<p class="sync-name">{html.escape(item["filename"])}{html.escape(folder_note)}</p>'
                f'<p class="sync-detail">{html.escape(detail)}</p>',
                unsafe_allow_html=True,
            )
        widget_id = hashlib.md5(item["key"].encode("utf-8")).hexdigest()[:10]
        with push_col:
            if st.button("Push to SharePoint", type="primary", key=f"push_{widget_id}"):
                _push_pending(item)
        with download_col:
            try:
                with open(item["path"], "rb") as f:
                    file_bytes = f.read()
            except OSError:
                file_bytes = None
            if file_bytes is not None:
                st.download_button(
                    label="Download",
                    data=file_bytes,
                    file_name=item["filename"],
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary",
                    key=f"download_{widget_id}",
                )
