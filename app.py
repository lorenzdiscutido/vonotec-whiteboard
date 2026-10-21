# app.py
"""Main flow of the Whiteboard AI web app."""
import os
import datetime
import importlib
import traceback
import ui_polish
import streamlit as st

import auto_watcher
import sharepoint_client
import ui_components
import ui_auto_watch
from settings import LOCAL_BATCH_DIR, MAIN_FOLDER_LABEL, ENABLE_SHAREPOINT_SYNC
from file_utils import sanitize_batch_filename, load_prefs, save_prefs
from batch_processor import run_batch, BatchOutcome
from ui_sections import (
    init_session_state, show_batch_results, render_photo_previews,
    render_folder_picker, resolve_subfolder, remember_new_folder,
    render_action_buttons, render_file_management, render_pending_syncs,
    hide_uploader_chrome_if_filled, process_blockers,
)
from ui_source_browser import render_source_browser
from ui_dashboard import render_dashboard_view
from dashboard_loader import start_warmup_if_due

# Hot-reload safeguards so Streamlit Cloud never fails on cached modules
if not hasattr(ui_auto_watch, "render_auto_watch_strip"):
    importlib.reload(ui_auto_watch)
if not hasattr(ui_components, "render_locked_save_settings"):
    importlib.reload(ui_components)

VIEW_EXTRACTION = "Extraction"
VIEW_DASHBOARD = "Dashboard"

SOURCE_LOCAL = "Device upload"
SOURCE_SHAREPOINT = "Import from SharePoint"

ui_polish.configure_page()        # must be the first st.* call; delete your existing st.set_page_config
# ... your existing CSS injection stays here ...
ui_polish.load_polish_css()       # right after it

# Normalize legacy session value if user switches from an older session
if st.session_state.get("photo_source") in ("Upload Local Photos", "Upload from device"):
    st.session_state["photo_source"] = SOURCE_LOCAL

# --- PAGE SETUP ---
st.set_page_config(page_title="VONOTEC Whiteboard Extractor", layout="wide")
ui_components.load_local_css("style.css")

# --- HEADER AND VIEW SWITCH ---
view = ui_components.render_header([VIEW_EXTRACTION, VIEW_DASHBOARD])

# --- DASHBOARD DATA ---
start_warmup_if_due()

# --- AUTO-EXTRACT ---
auto_watcher.start_if_due()

_KEEP_ALIVE_PREFIXES = (
    "photo_source", "batch_name_input", "folder_choice_input",
    "new_folder_name_input", "src_recursive", "src_lv_", "subfolder_choice_lv", "new_subfolder_name_lv",
)
for _key in list(st.session_state.keys()):
    if isinstance(_key, str) and _key.startswith(_KEEP_ALIVE_PREFIXES):
        st.session_state[_key] = st.session_state[_key]

if view == VIEW_DASHBOARD:
    st.session_state.batch_running = False
    render_dashboard_view()
    st.stop()

init_session_state()

if not st.session_state.get("prefs_loaded"):
    remembered_folder = load_prefs().get("save_folder")
    if isinstance(remembered_folder, str) and remembered_folder:
        st.session_state.last_folder_choice = remembered_folder
    st.session_state.prefs_loaded = True


class RemoteFile:
    """A photo that lives in SharePoint, shaped like a Streamlit UploadedFile."""

    def __init__(self, name, item_id):
        self.name = name
        self.item_id = item_id

    def getvalue(self):
        data = sharepoint_client.download_item_to_bytes(self.item_id)
        if data is None:
            raise FileNotFoundError("it is no longer in SharePoint")
        return data


def _prepare_output(batch_name_raw, folder_choice, new_folder_raw):
    filename = sanitize_batch_filename(batch_name_raw)
    if not filename:
        return None, "", False, "Please enter a name for this batch before extracting."
    subfolder, creating_new_folder, error = resolve_subfolder(folder_choice, new_folder_raw)
    if error:
        return None, "", False, error
    return filename, subfolder, creating_new_folder, None


def _run_and_finish(files, batch_name_raw, filename, subfolder, creating_new_folder, intro=None):
    local_dir = os.path.join(LOCAL_BATCH_DIR, subfolder) if subfolder else LOCAL_BATCH_DIR
    os.makedirs(local_dir, exist_ok=True)
    output_path = os.path.join(local_dir, filename)

    st.session_state.last_output_path = output_path
    st.session_state.last_output_label = batch_name_raw.strip()
    st.session_state.last_output_folder = subfolder
    st.session_state.last_folder_choice = subfolder if subfolder else MAIN_FOLDER_LABEL
    st.session_state.confirm_reset = False
    save_prefs({"save_folder": st.session_state.last_folder_choice})

    progress_bar = st.progress(0)
    status_text = st.empty()

    def on_progress(index, total, name, finished):
        if finished:
            progress_bar.progress((index + 1) / total)
        else:
            status_text.markdown(
                f'<p class="status-msg">Extracting {index + 1} of {total}: {name}</p>',
                unsafe_allow_html=True,
            )

    try:
        outcome = run_batch(files, output_path, subfolder, creating_new_folder, on_progress)
    except Exception as e:
        traceback.print_exc()
        stopped = ("error", f"The batch stopped unexpectedly: {e}")
        outcome = BatchOutcome(messages=[stopped], notices=[stopped])

    if outcome.created_folder:
        remember_new_folder(subfolder)

    results = list(intro or []) + outcome.messages
    results.append(("status", "Batch Complete."))
    results.append((
        "success",
        f"Successfully added {outcome.processed} new file(s). Skipped {outcome.skipped} duplicate(s).",
    ))
    st.session_state.batch_results = results
    st.session_state.batch_report = {"files": outcome.files, "notices": list(intro or []) + outcome.notices}

    st.session_state.uploader_key += 1
    st.rerun()


def _fail(kind, text):
    st.session_state.start_problem = (kind, text)
    st.rerun()


def start_batch(files, batch_name_raw, folder_choice, new_folder_raw):
    if not files:
        _fail("error", "Please select at least one photo before extracting.")
        return

    filename, subfolder, creating_new_folder, error = _prepare_output(
        batch_name_raw, folder_choice, new_folder_raw
    )
    if error:
        _fail("error", error)
        return

    _run_and_finish(files, batch_name_raw, filename, subfolder, creating_new_folder)


def start_remote_batch(batch_name_raw, folder_choice, new_folder_raw, source_path, include_subfolders):
    filename, subfolder, creating_new_folder, error = _prepare_output(
        batch_name_raw, folder_choice, new_folder_raw
    )
    if error:
        _fail("error", error)
        return

    if not source_path:
        _fail("error", "Please choose the SharePoint folder to read photos from.")
        return

    where = " / ".join(source_path)
    try:
        with st.spinner("Looking for photos in SharePoint..."):
            images = sharepoint_client.list_images_at(source_path, recursive=include_subfolders)
    except Exception as e:
        _fail("error", f"Could not read that SharePoint folder right now: {e}")
        return

    if not images:
        extra = "" if include_subfolders else " Turn on 'Include photos from all folders inside this one' to also look in its subfolders."
        _fail("warning", f"No photos (.jpg, .jpeg or .png) were found in {where}.{extra}")
        return

    save_prefs({"photo_source_path": list(source_path), "include_subfolders": bool(include_subfolders)})
    st.session_state.last_remote_path = list(source_path)

    files = [RemoteFile(item["display_name"], item["id"]) for item in images]
    intro = [("status", f"Found {len(files)} photo(s) in {where}"
                        f"{' and the folders below it' if include_subfolders else ''}.")]
    _run_and_finish(files, batch_name_raw, filename, subfolder, creating_new_folder, intro)


# 1. Global 3-step workflow stepper slot
stepper_slot = st.empty()

# 2. Dedicated SharePoint Auto-Extract status strip above the workspace
if hasattr(ui_auto_watch, "render_auto_watch_strip"):
    ui_auto_watch.render_auto_watch_strip()
elif hasattr(ui_auto_watch, "render_auto_watch_card"):
    ui_auto_watch.render_auto_watch_card()

# 3. 55/45 split workspace
form_col, results_col = st.columns([11, 9], gap="large")

# --- LEFT: PHOTO SOURCE, UPLOADER AND PROGRESSIVE SAVE SETTINGS ---
with form_col:
    with st.container(key="form_panel"):
        ui_components.section_heading(
            "Input source",
            "Choose where to import photos from",
            description="Select photos from your device or import them from an existing SharePoint folder.",
        )

        if ENABLE_SHAREPOINT_SYNC:
            with st.container(key="source_mode_switcher"):
                source = st.radio(
                    "Photo source", [SOURCE_LOCAL, SOURCE_SHAREPOINT],
                    horizontal=True, label_visibility="collapsed", key="photo_source",
                )
        else:
            source = SOURCE_LOCAL
        remote_mode = source == SOURCE_SHAREPOINT

        uploaded_files = []
        source_path, include_subfolders, remote_photo_count = [], False, 0
        if remote_mode:
            source_path, include_subfolders, remote_photo_count = render_source_browser()
        else:
            ui_components.subsection_heading(
                "Upload whiteboard photos",
                "Drop whiteboard photos here or click Browse files. Supports JPG and PNG up to 200 MB per file.",
            )
            uploaded_files = st.file_uploader(
                "Choose whiteboard images...",
                type=["jpg", "jpeg", "png"],
                accept_multiple_files=True,
                label_visibility="collapsed",
                key=f"uploader_{st.session_state.uploader_key}",
            )
            hide_uploader_chrome_if_filled(uploaded_files)
            render_photo_previews(uploaded_files)

        name_was_reset = (
            st.session_state.get("batch_name_round") != st.session_state.uploader_key
            or "batch_name_input" not in st.session_state
        )
        if name_was_reset:
            st.session_state.batch_name_round = st.session_state.uploader_key
            st.session_state.batch_name_input = f"Batch_{datetime.datetime.now().strftime('%Y-%m-%d_%H%M')}"
            st.session_state.sp_folders = None

        if remote_mode and source_path:
            folder_marker = (st.session_state.uploader_key, "/".join(source_path))
            if name_was_reset or st.session_state.get("batch_name_folder_marker") != folder_marker:
                st.session_state.batch_name_folder_marker = folder_marker
                st.session_state.batch_name_input = source_path[-1]

        # Progressive disclosure: Save settings stays in a compact locked card until Step 1 has photos
        active_photo_count = remote_photo_count if remote_mode else len(uploaded_files or [])
        photos_gate_open = (bool(source_path) and remote_photo_count > 0) if remote_mode else bool(uploaded_files)

        if photos_gate_open:
            with st.container(border=True, key="save_settings"):
                ui_components.card_heading("Output preferences", "Save settings")
                folder_column, name_column = st.columns(2, gap="medium")
                with folder_column:
                    folder_choice, new_folder_raw = render_folder_picker(disabled=False)
                with name_column:
                    batch_name_raw = st.text_input(
                        "Batch name",
                        key="batch_name_input",
                        help="This becomes the Excel file name, both here and in SharePoint. "
                             "With 'Import from SharePoint' it starts as the chosen folder's name, "
                             "and you can change it.",
                    )
        else:
            if hasattr(ui_components, "render_locked_save_settings"):
                ui_components.render_locked_save_settings()
            folder_choice = st.session_state.get("last_folder_choice", MAIN_FOLDER_LABEL)
            new_folder_raw = ""
            batch_name_raw = st.session_state.get("batch_name_input", "")

        photos_missing, destination_missing = process_blockers(
            active_photo_count, bool(source_path), folder_choice, new_folder_raw, batch_name_raw, remote_mode,
        )
        running = bool(st.session_state.get("batch_running", False))
        if remote_mode:
            batch_done = bool(st.session_state.batch_results) and list(source_path) == st.session_state.get("last_remote_path") and bool(source_path)
        else:
            batch_done = bool(st.session_state.batch_results) and not uploaded_files

        process_clicked, clear_clicked = render_action_buttons(
            photo_count=active_photo_count,
            source_chosen=bool(source_path) if remote_mode else bool(uploaded_files),
            blocked_reason=photos_missing or destination_missing,
            remote=remote_mode,
            running=running,
            done=batch_done,
        )
        if process_clicked:
            st.session_state.batch_running = True
            st.rerun()
        extract_clicked = running and not remote_mode
        fetch_clicked = running and remote_mode

# Fill the stepper now that the form's state is known
_done = set()
if batch_done:
    _done = {1, 2, 3}
    _current = 3
else:
    if not photos_missing:
        _done.add(1)
    if 1 in _done and not destination_missing:
        _done.add(2)
    _current = 1 if 1 not in _done else (2 if 2 not in _done else 3)
with stepper_slot.container():
    ui_components.render_stepper(_done, _current)

if clear_clicked:
    st.session_state.uploader_key += 1
    st.session_state.batch_results = []
    st.session_state.batch_report = None
    st.rerun()

# --- RIGHT: RESULTS, DOWNLOAD / DELETE, PENDING SYNCS ---
with results_col:
    with st.container(key="results_panel"):
        if uploaded_files:
            st.session_state.batch_results = []
            st.session_state.batch_report = None

        if extract_clicked or fetch_clicked:
            badge = ("Extracting", "processing")
        elif st.session_state.batch_results:
            badge = ("Complete", "complete")
        elif photos_gate_open:
            badge = ("Ready", "ready")
        else:
            badge = ("Waiting for photos", "waiting")
        ui_components.section_heading("Output", "Batch results", badge)

        last_file = st.session_state.last_output_path
        nothing_to_show = (
            not st.session_state.batch_results
            and not extract_clicked
            and not fetch_clicked
            and not (last_file and os.path.exists(last_file))
        )
        if nothing_to_show:
            ui_components.render_empty_results()

        problem = st.session_state.pop("start_problem", None)
        if problem:
            (st.error if problem[0] == "error" else st.warning)(problem[1])

        show_batch_results()

        try:
            if extract_clicked:
                start_batch(uploaded_files, batch_name_raw, folder_choice, new_folder_raw)
            if fetch_clicked:
                start_remote_batch(batch_name_raw, folder_choice, new_folder_raw, source_path, include_subfolders)
        finally:
            st.session_state.batch_running = False

        render_file_management()
        render_pending_syncs()

        report_files = (st.session_state.get("batch_report") or {}).get("files") or []
        review_count = sum(1 for f in report_files if f.get("status") == "review")
        try:
            ui_components.render_tip(
                completed=bool(st.session_state.batch_results),
                review_count=review_count,
            )
        except TypeError:
            ui_components.render_tip()
