# ui_sections.py
"""Reusable Streamlit sections for the Whiteboard app.
(Header and CSS stay in ui_components.py.)"""
import io
import os
import html
import hashlib
import traceback

import streamlit as st
from PIL import Image, ImageOps

import sharepoint_client
import sync_tracker
from settings import ENABLE_SHAREPOINT_SYNC, MAIN_FOLDER_LABEL, NEW_FOLDER_LABEL

_SESSION_DEFAULTS = {
    "uploader_key": 0,             # changing it resets the uploader (clears all photos)
    "batch_results": [],
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


# With more messages than this, the list goes into a scrollable box
RESULTS_BOX_MIN_MESSAGES = 5
RESULTS_BOX_HEIGHT_PX = 320
# app.py always ends a batch with two messages ("Batch Complete." and the summary
# line); they stay outside the box so the outcome is visible without scrolling.
RESULTS_PINNED_MESSAGES = 2


def _show_message(kind, text):
    if kind == "warning":
        st.warning(text)
    elif kind == "error":
        st.error(text)
    elif kind == "success":
        st.success(text)
    else:
        st.markdown(f'<p class="status-msg">{text}</p>', unsafe_allow_html=True)


def show_batch_results():
    """Shows the messages from the last batch (kept even after the uploader is cleared).
    A long list is shown inside a scrollable box, so it doesn't push the form and the
    Extract button far down the page."""
    results = st.session_state.batch_results
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
# Photo previews
# ---------------------------------------------------------------------------

PREVIEW_MAX_PHOTOS = 12   # a big batch shows the first few, then "+ N more"
PREVIEW_COLUMNS = 6       # thumbnails per row
PREVIEW_PIXELS = 240      # longest side of a thumbnail


@st.cache_data(show_spinner=False, max_entries=300)
def _thumbnail_bytes(name, size, file_id, _data):
    """A small JPEG of one uploaded photo, or None if it can't be read. Phone
    photos are large, so only this small copy is sent to the browser. It is made
    once per photo: the cache is keyed by name, size and id; the photo's own bytes
    (_data, underscore = not part of the key) are only used the first time."""
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


def render_photo_previews(uploaded_files):
    """A count and a grid of small previews of the selected photos, so people can
    check they picked the right ones before extracting."""
    if not uploaded_files:
        return

    total = len(uploaded_files)
    st.caption(f"{total} photo(s) selected")
    shown = list(uploaded_files)[:PREVIEW_MAX_PHOTOS]
    for start in range(0, len(shown), PREVIEW_COLUMNS):
        columns = st.columns(PREVIEW_COLUMNS)
        for column, photo in zip(columns, shown[start:start + PREVIEW_COLUMNS]):
            thumbnail = _thumbnail_bytes(photo.name, photo.size, getattr(photo, "file_id", ""), photo.getvalue())
            with column:
                if thumbnail:
                    st.image(thumbnail)
                else:
                    st.caption("(no preview)")
    if total > len(shown):
        st.caption(f"+ {total - len(shown)} more photo(s) not shown here")


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


def render_folder_picker():
    """Returns (folder_choice, new_folder_raw)."""
    if not ENABLE_SHAREPOINT_SYNC:
        return MAIN_FOLDER_LABEL, ""

    if st.session_state.sp_folders is None:
        with st.spinner("Loading SharePoint folders..."):
            refresh_folder_list()
    if st.session_state.sp_folders_error:
        st.warning(
            "Could not load the folder list from SharePoint right now. You can still save to "
            "the main folder, or create a new folder."
        )

    options = [MAIN_FOLDER_LABEL] + list(st.session_state.sp_folders) + [NEW_FOLDER_LABEL]
    default = st.session_state.last_folder_choice
    folder_choice = st.selectbox(
        "Save into which SharePoint folder?",
        options,
        index=options.index(default) if default in options else 0,
        key="folder_choice_input",
    )

    new_folder_raw = ""
    if folder_choice == NEW_FOLDER_LABEL:
        new_folder_raw = st.text_input("Name for the new folder", key="new_folder_name_input")
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

def render_action_buttons():
    """Returns (extract_clicked, clear_clicked)."""
    with st.container(key="action_buttons"):
        # The first (empty) column is a spacer that pushes both buttons to the right edge
        _spacer, col2, col3 = st.columns([0.6, 1.5, 1], gap="small")
        with col2:
            extract_clicked = st.button("Extract Data & Update Excel", type="primary")
        with col3:
            clear_clicked = st.button("Clear All Photos", type="primary", key="clear_photos")
    return extract_clicked, clear_clicked


def file_management_title_html(label, folder):
    """The heading of the file card: a small label over "Generated workbook", then the
    batch name and folder (which can be long) in lighter gray. Both are escaped,
    since they come from what people typed."""
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
            if st.button("Delete This Batch File", type="primary", key="start_fresh"):
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
# Pending syncs: batch files that never reached SharePoint
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
