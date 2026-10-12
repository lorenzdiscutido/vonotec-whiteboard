# app.py
"""Main flow of the Whiteboard AI web app. Everything else lives in:
    settings.py          switches and constants
    file_utils.py        processed-photo log, filename cleanup
    extraction.py        Gemini reading + review pass
    excel_writer.py      master Excel file (layout, units, repair formulas)
    batch_processor.py   the per-photo batch loop
    ui_sections.py       session state, folder picker, file management, pending syncs
    ui_source_browser.py cascading folder browser for "Process from SharePoint"
    sync_tracker.py      remembers which batch files reached SharePoint
    ui_components.py     CSS and header
    ui_dashboard.py      the Analytics Dashboard view (charts from the SharePoint Excel files)
    dashboard_data.py    reads the batch Excel files into a table for the dashboard
    dashboard_loader.py  downloads those files in the background so the dashboard is ready

Photos can come from two places: uploaded from this computer, or already sitting in any
SharePoint folder ("Process from SharePoint"). Both end up in the same batch loop.
"""
import os
import datetime

import streamlit as st

import sharepoint_client
import ui_components
from settings import LOCAL_BATCH_DIR, MAIN_FOLDER_LABEL, ENABLE_SHAREPOINT_SYNC
from file_utils import sanitize_batch_filename, load_prefs, save_prefs
from batch_processor import run_batch
from ui_sections import (
    init_session_state, show_batch_results, render_photo_previews,
    render_folder_picker, resolve_subfolder, remember_new_folder,
    render_action_buttons, render_file_management, render_pending_syncs,
)
from ui_source_browser import render_source_browser
from ui_dashboard import render_analytics_dashboard, render_dashboard_controls
from dashboard_loader import start_warmup_if_due

VIEW_EXTRACTION = "Whiteboard Extraction"
VIEW_DASHBOARD = "Analytics Dashboard"

SOURCE_LOCAL = "Upload Local Photos"
SOURCE_SHAREPOINT = "Process from SharePoint"

# --- PAGE SETUP ---
st.set_page_config(page_title="Vonotec Whiteboard Extractor", layout="wide")
ui_components.load_local_css("style.css")
ui_components.render_header()

# --- DASHBOARD DATA ---
# Start reading the Excel files for the dashboard in the background right away, in
# whichever view the person opened, so the charts are ready when they click over.
# It never blocks the page, and does nothing if one is already running or ran recently.
start_warmup_if_due()

# --- VIEW SWITCH ---
# Streamlit runs the code of every st.tabs() tab on each click, which would
# make the dashboard read SharePoint while someone is extracting photos. A
# radio that only runs the chosen view avoids that.
with st.container(key="view_switch"):
    view = st.radio(
        "View", [VIEW_EXTRACTION, VIEW_DASHBOARD],
        horizontal=True, label_visibility="collapsed", key="active_view",
    )

# Streamlit forgets the value of any widget that is not drawn in a run. While the
# dashboard is open the extraction widgets are not drawn, so without this they would
# reset (the photo source back to "Upload Local Photos", the folder choices, the batch
# name...) when the person comes back. Writing a widget's value back into session state
# on every run keeps it. This must run before those widgets are drawn.
_KEEP_ALIVE_PREFIXES = (
    "photo_source", "batch_name_input", "folder_choice_input",
    "new_folder_name_input", "src_recursive", "src_lv_",
)
for _key in list(st.session_state.keys()):
    if isinstance(_key, str) and _key.startswith(_KEEP_ALIVE_PREFIXES):
        st.session_state[_key] = st.session_state[_key]

if view == VIEW_DASHBOARD:
    # The controls (which files, Refresh, filters) go in the sidebar; the numbers
    # and charts fill the main area and refresh by themselves.
    with st.sidebar:
        render_dashboard_controls()
    with st.container(key="analytics"):
        render_analytics_dashboard()
    st.stop()  # nothing below belongs to the dashboard view

ui_components.render_instructions()
init_session_state()

# Remember where the person left off: on the first run of a session, the "Save into which
# folder" dropdown starts on the folder used for the last batch (kept in user_prefs.json,
# so it survives a page refresh). The photo-folder selectors read the same file themselves.
if not st.session_state.get("prefs_loaded"):
    remembered_folder = load_prefs().get("save_folder")
    if isinstance(remembered_folder, str) and remembered_folder:
        st.session_state.last_folder_choice = remembered_folder
    st.session_state.prefs_loaded = True


class RemoteFile:
    """A photo that lives in SharePoint, shaped like a Streamlit UploadedFile
    (it has a .name and a .getvalue()). The download happens only when
    getvalue() is called, so the batch loop fetches one photo at a time and
    never holds the whole folder in memory. `name` is the file name with its
    subfolder path, so photos with the same name in different folders are told apart."""

    def __init__(self, name, item_id):
        self.name = name
        self.item_id = item_id

    def getvalue(self):
        data = sharepoint_client.download_item_to_bytes(self.item_id)
        if data is None:
            raise FileNotFoundError("it is no longer in SharePoint")
        return data


def _prepare_output(batch_name_raw, folder_choice, new_folder_raw):
    """Checks the batch name and folder. Returns (filename, subfolder,
    creating_new_folder, error_message); error_message is None when all is well."""
    filename = sanitize_batch_filename(batch_name_raw)
    if not filename:
        return None, "", False, "Please enter a name for this batch before extracting."
    subfolder, creating_new_folder, error = resolve_subfolder(folder_choice, new_folder_raw)
    if error:
        return None, "", False, error
    return filename, subfolder, creating_new_folder, None


def _run_and_finish(files, batch_name_raw, filename, subfolder, creating_new_folder, intro=None):
    """Runs the batch for `files` (uploaded or from SharePoint), then stores the
    messages and resets the uploader. `intro` is a list of (kind, text) messages
    to show before the batch's own messages."""
    local_dir = os.path.join(LOCAL_BATCH_DIR, subfolder) if subfolder else LOCAL_BATCH_DIR
    os.makedirs(local_dir, exist_ok=True)
    output_path = os.path.join(local_dir, filename)

    st.session_state.last_output_path = output_path
    st.session_state.last_output_label = batch_name_raw.strip()
    st.session_state.last_output_folder = subfolder
    st.session_state.last_folder_choice = subfolder if subfolder else MAIN_FOLDER_LABEL
    save_prefs({"save_folder": st.session_state.last_folder_choice})

    progress_bar = st.progress(0)
    status_text = st.empty()

    def on_progress(index, total, name, finished):
        if finished:
            progress_bar.progress((index + 1) / total)
        else:
            status_text.markdown(
                f'<p class="status-msg">Processing image {index + 1} of {total}: {name}</p>',
                unsafe_allow_html=True,
            )

    outcome = run_batch(files, output_path, subfolder, creating_new_folder, on_progress)

    if outcome.created_folder:
        remember_new_folder(subfolder)

    results = list(intro or []) + outcome.messages
    results.append(("status", "Batch Complete."))
    results.append((
        "success",
        f"Successfully added {outcome.processed} new file(s). Skipped {outcome.skipped} duplicate(s).",
    ))
    st.session_state.batch_results = results
    st.session_state.uploader_key += 1  # clears the uploader and suggests a fresh batch name
    st.rerun()


def start_batch(files, batch_name_raw, folder_choice, new_folder_raw):
    """Local photos: validates the form, then runs the batch."""
    if not files:
        st.error("Please select at least one photo before extracting.")
        return

    filename, subfolder, creating_new_folder, error = _prepare_output(
        batch_name_raw, folder_choice, new_folder_raw
    )
    if error:
        st.error(error)
        return

    _run_and_finish(files, batch_name_raw, filename, subfolder, creating_new_folder)


def start_remote_batch(batch_name_raw, folder_choice, new_folder_raw, source_path, include_subfolders):
    """SharePoint photos: lists the photos in the folder opened in the browser
    (and below it, if asked), then runs the batch over them. Each photo is
    downloaded as the loop reaches it. The Excel file goes to the "Save settings" folder."""
    filename, subfolder, creating_new_folder, error = _prepare_output(
        batch_name_raw, folder_choice, new_folder_raw
    )
    if error:
        st.error(error)
        return

    if not source_path:
        st.error("Please choose the SharePoint folder to read photos from.")
        return

    where = " / ".join(source_path)
    try:
        with st.spinner("Looking for photos in SharePoint..."):
            images = sharepoint_client.list_images_at(source_path, recursive=include_subfolders)
    except Exception as e:
        st.error(f"Could not read that SharePoint folder right now: {e}")
        return

    if not images:
        extra = "" if include_subfolders else " Turn on 'Include photos from all folders inside this one' to also look in its subfolders."
        st.warning(f"No photos (.jpg, .jpeg or .png) were found in {where}.{extra}")
        return

    # Remember this location so the selectors open on it next time
    save_prefs({"photo_source_path": list(source_path), "include_subfolders": bool(include_subfolders)})

    files = [RemoteFile(item["display_name"], item["id"]) for item in images]
    intro = [("status", f"Found {len(files)} photo(s) in {where}"
                        f"{' and the folders below it' if include_subfolders else ''}.")]
    _run_and_finish(files, batch_name_raw, filename, subfolder, creating_new_folder, intro)


# Two columns, like the dashboard uses the full width: the form on the left, and
# everything that comes out of it (messages, download, sync status) on the right,
# so a long list of messages never pushes the form's buttons down the page.
form_col, results_col = st.columns([3, 2], gap="large")

# --- LEFT: PHOTO SOURCE, UPLOADER AND BATCH FORM ---
with form_col:
    st.subheader("Whiteboard Photos")

    # A radio rather than st.tabs: only the chosen source's code runs on each click.
    # Without SharePoint sync there is nothing to read from, so only uploading is offered.
    if ENABLE_SHAREPOINT_SYNC:
        with st.container(key="source_switch"):
            source = st.radio(
                "Photo source", [SOURCE_LOCAL, SOURCE_SHAREPOINT],
                horizontal=True, label_visibility="collapsed", key="photo_source",
            )
    else:
        source = SOURCE_LOCAL
    remote_mode = source == SOURCE_SHAREPOINT

    uploaded_files = []
    source_path, include_subfolders = [], False
    if remote_mode:
        st.caption(
            "Pick the SharePoint folder that holds the photos, one level at a time: each choice "
            "opens a selector for the folders inside it. The photos are read directly from "
            "SharePoint, one at a time. The Excel file is saved to the folder chosen under "
            "Save settings. The last folder you used is remembered."
        )
        source_path, include_subfolders = render_source_browser()
    else:
        uploaded_files = st.file_uploader(
            "Choose whiteboard images...",
            type=["jpg", "jpeg", "png"],
            accept_multiple_files=True,
            label_visibility="collapsed",
            key=f"uploader_{st.session_state.uploader_key}",
        )
        render_photo_previews(uploaded_files)

    # Suggest a fresh default name each time a new round of photos is selected,
    # without overwriting what the person is typing in the meantime. It is also
    # refilled when the Batch name box was dropped (for example after the person
    # visited the Analytics Dashboard and came back), because Streamlit forgets
    # the state of widgets that were not drawn in a run.
    name_was_reset = (
        st.session_state.get("batch_name_round") != st.session_state.uploader_key
        or "batch_name_input" not in st.session_state
    )
    if name_was_reset:
        st.session_state.batch_name_round = st.session_state.uploader_key
        st.session_state.batch_name_input = f"Batch_{datetime.datetime.now().strftime('%Y-%m-%d_%H%M')}"
        # Re-read the folder list each round, so folders other people created show up
        st.session_state.sp_folders = None

    # "Process from SharePoint": the batch name defaults to the name of the chosen folder.
    # It is filled in when a folder is chosen (or a new round starts, or the box was
    # refilled above) and then left alone, so whatever the person types over it is kept
    # until they pick a different folder.
    if remote_mode and source_path:
        folder_marker = (st.session_state.uploader_key, "/".join(source_path))
        if name_was_reset or st.session_state.get("batch_name_folder_marker") != folder_marker:
            st.session_state.batch_name_folder_marker = folder_marker
            st.session_state.batch_name_input = source_path[-1]

    # Where the file goes and what it is called: grouped in a box, side by side
    with st.container(border=True, key="save_settings"):
        st.markdown('<p class="panel-title">Save settings</p>', unsafe_allow_html=True)
        folder_column, name_column = st.columns(2, gap="medium")
        with folder_column:
            folder_choice, new_folder_raw = render_folder_picker()
        with name_column:
            batch_name_raw = st.text_input(
                "Batch name",
                key="batch_name_input",
                help="This becomes the Excel file name, both here and in SharePoint. "
                     "With 'Process from SharePoint' it starts as the chosen folder's name, "
                     "and you can change it.",
            )

    if remote_mode:
        with st.container(key="action_buttons"):
            fetch_clicked = st.button("Fetch and Process Folder", type="primary", key="fetch_folder")
        extract_clicked, clear_clicked = False, False
    else:
        extract_clicked, clear_clicked = render_action_buttons()
        fetch_clicked = False

if clear_clicked:
    st.session_state.uploader_key += 1
    st.session_state.batch_results = []
    st.rerun()

# --- RIGHT: RESULTS, DOWNLOAD / DELETE, FILES THAT NEVER REACHED SHAREPOINT ---
with results_col:
    st.subheader("Batch Results")
    st.caption("Messages, downloads and sync status appear here once you extract a batch.")

    # As soon as new photos are selected, drop the previous batch's messages
    if uploaded_files:
        st.session_state.batch_results = []
    show_batch_results()

    if extract_clicked:
        start_batch(uploaded_files, batch_name_raw, folder_choice, new_folder_raw)
    if fetch_clicked:
        start_remote_batch(batch_name_raw, folder_choice, new_folder_raw, source_path, include_subfolders)

    render_file_management()
    render_pending_syncs()
