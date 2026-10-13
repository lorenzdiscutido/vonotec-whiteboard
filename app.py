# app.py
"""Main flow of the Whiteboard AI web app. Everything else lives in:
    settings.py          switches and constants
    file_utils.py        processed-photo log, filename cleanup
    extraction.py        Gemini reading + review pass
    excel_writer.py      master Excel file (layout, units, repair formulas)
    batch_processor.py   the per-photo batch loop
    ui_sections.py       session state, folder picker, file management, pending syncs
    ui_source_browser.py folder card for "Import from SharePoint"
    sync_tracker.py      remembers which batch files reached SharePoint
    ui_components.py     CSS and header
    ui_dashboard.py      the Analytics Dashboard view (charts from the SharePoint Excel files)
    dashboard_data.py    reads the batch Excel files into a table for the dashboard
    dashboard_loader.py  downloads those files in the background so the dashboard is ready

Photos can come from two places: uploaded from this device, or already sitting in any
SharePoint folder ("Import from SharePoint"). Both end up in the same batch loop.
"""
import os
import datetime
import traceback

import streamlit as st

import sharepoint_client
import ui_components
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

VIEW_EXTRACTION = "Extraction"
VIEW_DASHBOARD = "Dashboard"

# Parallel wording: both options describe the same stage (bringing photos in)
SOURCE_LOCAL = "Upload from device"
SOURCE_SHAREPOINT = "Import from SharePoint"

# --- PAGE SETUP ---
st.set_page_config(page_title="Vonotec Whiteboard Extractor", layout="wide")
ui_components.load_local_css("style.css")

# --- HEADER AND VIEW SWITCH ---
# One header bar: the brand on the left, the Extraction / Dashboard switch in the middle
# and the tagline on the right. Only the chosen view's code runs on each click (a radio
# rather than st.tabs, which would run every tab and make the dashboard read SharePoint
# while someone is extracting photos).
view = ui_components.render_header([VIEW_EXTRACTION, VIEW_DASHBOARD])

# --- DASHBOARD DATA ---
# Start reading the Excel files for the dashboard in the background right away, in
# whichever view the person opened, so the charts are ready when they click over.
# It never blocks the page, and does nothing if one is already running or ran recently.
start_warmup_if_due()

# Streamlit forgets the value of any widget that is not drawn in a run. While the
# dashboard is open the extraction widgets are not drawn, so without this they would
# reset (the photo source, the folder choices, the batch name...) when the person
# comes back. Writing a widget's value back into session state on every run keeps it.
# This must run before those widgets are drawn.
_KEEP_ALIVE_PREFIXES = (
    "photo_source", "batch_name_input", "folder_choice_input",
    "new_folder_name_input", "src_recursive", "src_lv_",
)
for _key in list(st.session_state.keys()):
    if isinstance(_key, str) and _key.startswith(_KEEP_ALIVE_PREFIXES):
        st.session_state[_key] = st.session_state[_key]

if view == VIEW_DASHBOARD:
    # One connected card: the controls (which files, Refresh, filters) on the left, the
    # numbers and charts on the right, which refresh by themselves.
    render_dashboard_view()
    st.stop()  # nothing below belongs to the dashboard view

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
    to show before the batch's own messages.

    Everything the right-hand panel needs (messages, the generated workbook's path,
    label and folder) lives in st.session_state, so it survives the st.rerun() at
    the end. Bumping uploader_key gives st.file_uploader a brand-new key, which
    empties it and the photo previews."""
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
                f'<p class="status-msg">Processing image {index + 1} of {total}: {name}</p>',
                unsafe_allow_html=True,
            )

    # Even if something unexpected goes wrong, the uploader is still reset below
    # and the person sees the error in the Batch results panel.
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
    # The compact Batch results list is drawn from these rows and notices
    st.session_state.batch_report = {"files": outcome.files, "notices": list(intro or []) + outcome.notices}

    # Reset the uploader: a new key means a fresh, empty st.file_uploader (and no previews),
    # and it also suggests a fresh batch name. The right panel keeps showing the results.
    st.session_state.uploader_key += 1
    st.rerun()


def start_batch(files, batch_name_raw, folder_choice, new_folder_raw):
    """Local photos: validates the form, then runs the batch."""
    if not files:
        st.error("Please select at least one photo before processing.")
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
# The workflow stepper sits above both columns. It is filled in further down, once the
# form's values are known (a placeholder lets it be drawn at the top after the fact).
stepper_slot = st.empty()

form_col, results_col = st.columns([3, 2], gap="large")

# --- LEFT: PHOTO SOURCE, UPLOADER AND BATCH FORM ---
with form_col:
    with st.container(key="form_panel"):
        # Stable heading first: what this step is, then the choice, then the content for that choice.
        ui_components.section_heading(
            "Input source", "Choose where to import photos from",
            description="Select photos from your device or import them from an existing SharePoint folder.",
        )

        # A radio rather than st.tabs: only the chosen source's code runs on each click.
        # Without SharePoint sync there is nothing to read from, so only uploading is offered.
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
        source_path, include_subfolders = [], False
        if remote_mode:
            # The folder card has its own title ("Choose a SharePoint folder"), status,
            # selectors and photo count
            source_path, include_subfolders = render_source_browser()
        else:
            ui_components.subsection_heading(
                "Upload whiteboard photos", "Drop photos here or browse your device. JPG and PNG are supported.",
            )
            uploaded_files = st.file_uploader(
                "Choose whiteboard images...",
                type=["jpg", "jpeg", "png"],
                accept_multiple_files=True,
                label_visibility="collapsed",
                key=f"uploader_{st.session_state.uploader_key}",
            )
            # With 1+ photos chosen, hide the round upload icon, the helper text and the
            # browse overlay, so only the scrollable file chips remain in the dashed box.
            hide_uploader_chrome_if_filled(uploaded_files)
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

        # "Import from SharePoint": the batch name defaults to the name of the chosen folder.
        # It is filled in when a folder is chosen (or a new round starts, or the box was
        # refilled above) and then left alone, so whatever the person types over it is kept
        # until they pick a different folder.
        if remote_mode and source_path:
            folder_marker = (st.session_state.uploader_key, "/".join(source_path))
            if name_was_reset or st.session_state.get("batch_name_folder_marker") != folder_marker:
                st.session_state.batch_name_folder_marker = folder_marker
                st.session_state.batch_name_input = source_path[-1]

        # Where the file goes and what it is called: grouped in a box, side by side
        # Step 1 gate: the card is muted (not locked) until photos are chosen
        photos_gate_open = bool(source_path) if remote_mode else bool(uploaded_files)
        with st.container(border=True, key="save_settings" if photos_gate_open else "save_settings_muted"):
            ui_components.card_heading("Output preferences", "Save settings")
            folder_column, name_column = st.columns(2, gap="medium")
            with folder_column:
                folder_choice, new_folder_raw = render_folder_picker()
            with name_column:
                batch_name_raw = st.text_input(
                    "Batch name",
                    key="batch_name_input",
                    help="This becomes the Excel file name, both here and in SharePoint. "
                         "With 'Import from SharePoint' it starts as the chosen folder's name, "
                         "and you can change it.",
                )

        # The sticky footer: the same final action whichever source was used. Its text carries
        # the context ("Process 16 photos") and it stays disabled until everything is valid.
        photos_missing, destination_missing = process_blockers(
            len(uploaded_files or []), bool(source_path), folder_choice, new_folder_raw, batch_name_raw, remote_mode,
        )
        process_clicked, clear_clicked = render_action_buttons(
            photo_count=len(uploaded_files or []),
            blocked_reason=photos_missing or destination_missing,
            remote=remote_mode,
        )
        extract_clicked = process_clicked and not remote_mode
        fetch_clicked = process_clicked and remote_mode

# Fill the stepper now that the form's state is known: finished steps green, the one in
# progress blue, the rest muted.
_done = set()  # steps complete so far; a later step is never done before an earlier one
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
    st.rerun()

# --- RIGHT: RESULTS, DOWNLOAD / DELETE, FILES THAT NEVER REACHED SHAREPOINT ---
with results_col:
    with st.container(key="results_panel"):
        # As soon as new photos are selected, drop the previous batch's messages
        if uploaded_files:
            st.session_state.batch_results = []

        badge = ("Complete", "complete") if st.session_state.batch_results else ("Ready", "ready")
        ui_components.section_heading("Output", "Batch results", badge)

        # Before the first batch there is nothing to show: say so in a card
        last_file = st.session_state.last_output_path
        nothing_to_show = (
            not st.session_state.batch_results
            and not extract_clicked
            and not fetch_clicked
            and not (last_file and os.path.exists(last_file))
        )
        if nothing_to_show:
            ui_components.render_empty_results()

        show_batch_results()

        if extract_clicked:
            start_batch(uploaded_files, batch_name_raw, folder_choice, new_folder_raw)
        if fetch_clicked:
            start_remote_batch(batch_name_raw, folder_choice, new_folder_raw, source_path, include_subfolders)

        render_file_management()
        render_pending_syncs()
        ui_components.render_tip()
