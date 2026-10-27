"""Main flow of the Whiteboard Photo Sorter & PowerPoint Generator web app. Everything else lives in:

    settings.py          switches and constants
    file_utils.py        processed-photo log, filename cleanup
    extraction.py        Gemini reading + review pass
    pptx_writer.py       PowerPoint (.pptx) generator and in-memory image optimizer
    batch_processor.py   chronological sorting, daily chunking (max 4), and batch pipeline
    ui_sections.py       session state, folder picker, file management, pending syncs
    ui_source_browser.py folder card for "Import from SharePoint"
    sync_tracker.py      remembers which batch files reached SharePoint
    ui_components.py     CSS and header

Photos can come from two places: uploaded from this device, or already sitting
in any SharePoint folder ("Import from SharePoint"). Both end up in the same
batch loop.
"""
import datetime
import importlib
import os
import traceback
from pathlib import Path

from PIL import Image
import streamlit as st

import sharepoint_client
import ui_components
from batch_processor import BatchOutcome, run_batch
from file_utils import load_prefs, sanitize_batch_filename, save_prefs
from gemini_client import get_key_pool_stats
from settings import ENABLE_SHAREPOINT_SYNC, LOCAL_BATCH_DIR, MAIN_FOLDER_LABEL
from ui_sections import (
    hide_uploader_chrome_if_filled,
    init_session_state,
    process_blockers,
    remember_new_folder,
    render_action_buttons,
    render_file_management,
    render_folder_picker,
    render_pending_syncs,
    render_photo_previews,
    resolve_subfolder,
    show_batch_results,
)
from ui_source_browser import render_source_browser

# Parallel wording: both options describe the same stage (bringing photos in)
SOURCE_LOCAL = "Upload from device"
SOURCE_SHAREPOINT = "Import from SharePoint"


# --- PAGE SETUP ---
icon_path = Path(__file__).with_name("vonotec_icon.png")
page_icon = Image.open(icon_path) if icon_path.exists() else "📊"

st.set_page_config(
    page_title="VONOTEC Photo Sorter & PowerPoint Generator",
    page_icon=page_icon,
    layout="wide",
)
if "dark_mode" not in st.session_state:
    st.session_state.dark_mode = bool(load_prefs().get("dark_mode", False))

ui_components.load_local_css("style.css")
ui_components.render_header()

_KEEP_ALIVE_PREFIXES = (
    "photo_source",
    "batch_name_input",
    "folder_choice_input",
    "new_folder_name_input",
    "tb_title_input",
    "tb_service_input",
    "src_recursive",
    "src_lv_",
    "dark_mode",
)
for _key in list(st.session_state.keys()):
    if isinstance(_key, str) and _key.startswith(_KEEP_ALIVE_PREFIXES):
        st.session_state[_key] = st.session_state[_key]

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
    """Checks the batch name and folder. Returns (filename, subfolder,
    creating_new_folder, error_message); error_message is None when all is well."""
    filename = sanitize_batch_filename(batch_name_raw)
    if not filename:
        return None, "", False, "Please enter a name for this batch before extracting."
    subfolder, creating_new_folder, error = resolve_subfolder(folder_choice, new_folder_raw)
    if error:
        return None, "", False, error
    return filename, subfolder, creating_new_folder, None


def _run_and_finish(files, batch_name_raw, filename, subfolder, creating_new_folder, intro=None, title_block_info=None):
    """Runs the batch for `files`, then stores the messages and resets the uploader."""
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
        outcome = run_batch(files, output_path, subfolder, creating_new_folder, on_progress, title_block_info=title_block_info)
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
        f"Successfully added {outcome.processed} new file(s) to the PowerPoint report. Skipped {outcome.skipped} duplicate(s).",
    ))
    st.session_state.batch_results = results
    st.session_state.batch_report = {"files": outcome.files, "notices": list(intro or []) + outcome.notices}

    st.session_state.uploader_key += 1
    st.rerun()


def _fail(kind, text):
    st.session_state.start_problem = (kind, text)
    st.rerun()


def start_batch(files, batch_name_raw, folder_choice, new_folder_raw, title_block_info=None):
    """Local photos: validates the form, then runs the batch."""
    if not files:
        _fail("error", "Please select at least one photo before processing.")
        return

    filename, subfolder, creating_new_folder, error = _prepare_output(
        batch_name_raw, folder_choice, new_folder_raw
    )
    if error:
        _fail("error", error)
        return

    _run_and_finish(files, batch_name_raw, filename, subfolder, creating_new_folder, title_block_info=title_block_info)


def start_remote_batch(batch_name_raw, folder_choice, new_folder_raw, source_path, include_subfolders, title_block_info=None):
    """SharePoint photos: lists the photos in the folder opened in the browser, then runs the batch."""
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
    intro = [(
        "status",
        f"Found {len(files)} photo(s) in {where}"
        f"{' and the folders below it' if include_subfolders else ''}.",
    )]
    _run_and_finish(files, batch_name_raw, filename, subfolder, creating_new_folder, intro, title_block_info=title_block_info)


# 55/45 split gives the results and review list plenty of horizontal room
form_col, results_col = st.columns([11, 9], gap="large")

# --- LEFT: PHOTO SOURCE, UPLOADER AND BATCH FORM ---
with form_col:
    with st.container(key="form_panel"):
        stepper_slot = st.empty()
        ui_components.section_heading(
            "",
            "Choose where to import photos from",
            description="Select photos from your device or import them from an existing SharePoint folder.",
        )

        pool_stats = get_key_pool_stats()
        if pool_stats["total_count"] > 1:
            st.markdown(
                f'<div style="display: inline-flex; align-items: center; gap: 6px; padding: 4px 12px; '
                f'border-radius: 9999px; background: rgba(30, 58, 138, 0.08); font-size: 12px; '
                f'font-weight: 600; color: #1e3a8a; margin: 2px 0 10px 0;">'
                f'<span>⚡</span> Gemini Key Pool: {pool_stats["active_count"]} active keys (~{pool_stats["daily_quota"]:,} daily requests)'
                f'</div>',
                unsafe_allow_html=True,
            )
        elif pool_stats["total_count"] == 0:
            st.warning("⚠️ No Gemini API key detected. Please configure GEMINI_API_KEY in secrets.")

        if ENABLE_SHAREPOINT_SYNC:
            with st.container(key="source_mode_switcher"):
                source = st.radio(
                    "Photo source",
                    [SOURCE_LOCAL, SOURCE_SHAREPOINT],
                    horizontal=True,
                    label_visibility="collapsed",
                    key="photo_source",
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
                "Drop photos here or browse your device. JPG and PNG are supported.",
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

        # Step 1 strict gate: Save settings is locked until photos are uploaded or a SharePoint folder with photos is chosen
        active_photo_count = remote_photo_count if remote_mode else len(uploaded_files or [])
        photos_gate_open = (bool(source_path) and remote_photo_count > 0) if remote_mode else bool(uploaded_files)

        with st.container(border=True, key="save_settings" if photos_gate_open else "save_settings_muted"):
            ui_components.card_heading("Output preferences", "Save settings")
            folder_column, name_column = st.columns(2, gap="medium")
            with folder_column:
                folder_choice, new_folder_raw = render_folder_picker(disabled=not photos_gate_open)
            with name_column:
                batch_name_raw = st.text_input(
                    "Batch name",
                    key="batch_name_input",
                    disabled=not photos_gate_open,
                    help="This becomes the PowerPoint (.pptx) file name, both here and in SharePoint. "
                         "With 'Import from SharePoint' it starts as the chosen folder's name, "
                         "and you can change it.",
                )

            photos_missing, destination_missing = process_blockers(
                active_photo_count,
                bool(source_path),
                folder_choice,
                new_folder_raw,
                batch_name_raw,
                remote_mode,
            )

            st.markdown("<hr style='margin: 16px 0 14px 0; border: none; border-top: 1px solid var(--border-subtle);' />", unsafe_allow_html=True)
            ui_components.card_heading("Slide title block details (optional)", "Slide information")

            tb_col1, tb_col2 = st.columns(2, gap="medium")
            with tb_col1:
                project_title = st.text_input(
                    "Title",
                    key="tb_title_input",
                    placeholder="e.g. Project Name / Location",
                    disabled=not photos_gate_open,
                    help="Custom title placed in Field 2 of each slide's bottom title block. Leave blank to keep empty.",
                )
                service_offered = st.text_input(
                    "Service offered",
                    key="tb_service_input",
                    placeholder="e.g. Facade Inspection / Repair Log",
                    disabled=not photos_gate_open,
                    help="Service type placed in Field 3 of each slide's bottom title block. Leave blank to keep empty.",
                )
            with tb_col2:
                st.markdown(
                    "<div style='font-size: 13px; font-weight: 600; color: var(--text-primary); margin-bottom: 4px;'>"
                    "Location logo (Field 4)"
                    "</div>",
                    unsafe_allow_html=True,
                )
                with st.container(key="compact_logo_uploader"):
                    loc_logo_file = st.file_uploader(
                        "Location logo (optional)",
                        type=["png", "jpg", "jpeg"],
                        key=f"tb_logo_{st.session_state.uploader_key}",
                        disabled=not photos_gate_open,
                        label_visibility="collapsed",
                        help="Optional logo of the location placed in Field 4 of each slide's bottom title block.",
                    )
                if loc_logo_file is not None:
                    st.image(loc_logo_file, caption="Field 4 Logo Preview", width=110)

            loc_logo_bytes = loc_logo_file.getvalue() if loc_logo_file is not None else None
            title_block_info = {
                "title": project_title.strip() if project_title else "",
                "service_offered": service_offered.strip() if service_offered else "",
                "location_logo_bytes": loc_logo_bytes,
            }

            running = bool(st.session_state.get("batch_running", False))
            if remote_mode:
                batch_done = (
                    bool(st.session_state.batch_results)
                    and list(source_path) == st.session_state.get("last_remote_path")
                    and bool(source_path)
                )
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

# Fill the stepper now that the form's state is known:
# Once a batch has completed (batch_done), all 3 steps show as completed (green checkmarks).
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
    st.session_state.tb_title_input = ""
    st.session_state.tb_service_input = ""
    st.rerun()


# --- RIGHT: RESULTS, DOWNLOAD / DELETE, FILES THAT NEVER REACHED SHAREPOINT ---
with results_col:
    with st.container(key="results_panel"):
        if uploaded_files:
            st.session_state.batch_results = []
            st.session_state.batch_report = None

        if extract_clicked or fetch_clicked:
            badge = ("Extracting", "processing")
        elif st.session_state.batch_results:
            badge = ("Complete", "complete")
        else:
            badge = ("Ready", "ready")
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
                start_batch(uploaded_files, batch_name_raw, folder_choice, new_folder_raw, title_block_info=title_block_info)
            if fetch_clicked:
                start_remote_batch(batch_name_raw, folder_choice, new_folder_raw, source_path, include_subfolders, title_block_info=title_block_info)
        finally:
            st.session_state.batch_running = False

        render_file_management()
        render_pending_syncs()

        # Contextual footer tip
        report_files = (st.session_state.get("batch_report") or {}).get("files") or []
        review_count = sum(1 for f in report_files if f.get("status") == "review")
        try:
            ui_components.render_tip(
                completed=bool(st.session_state.batch_results),
                review_count=review_count,
            )
        except TypeError:
            importlib.reload(ui_components)
            try:
                ui_components.render_tip(
                    completed=bool(st.session_state.batch_results),
                    review_count=review_count,
                )
            except TypeError:
                ui_components.render_tip()
