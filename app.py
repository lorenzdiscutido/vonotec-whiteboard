# app.py
"""Main flow of the Whiteboard AI web app. Everything else lives in:
    settings.py         switches and constants
    file_utils.py       processed-photo log, filename cleanup
    extraction.py       Gemini reading + review pass
    excel_writer.py     master Excel file (layout, units, repair formulas)
    batch_processor.py  the per-photo batch loop
    ui_sections.py      session state, folder picker, file management, pending syncs
    sync_tracker.py     remembers which batch files reached SharePoint
    ui_components.py    CSS and header
    ui_dashboard.py     the Analytics Dashboard view (charts from the SharePoint Excel files)
    dashboard_data.py   reads the batch Excel files into a table for the dashboard
    dashboard_loader.py downloads those files in the background so the dashboard is ready
"""
import os
import datetime

import streamlit as st

import ui_components
from settings import LOCAL_BATCH_DIR, MAIN_FOLDER_LABEL
from file_utils import sanitize_batch_filename
from batch_processor import run_batch
from ui_sections import (
    init_session_state, show_batch_results, status_message,
    render_folder_picker, resolve_subfolder, remember_new_folder,
    render_action_buttons, render_file_management, render_pending_syncs,
)
from ui_dashboard import render_analytics_dashboard, render_dashboard_controls
from dashboard_loader import start_warmup_if_due

VIEW_EXTRACTION = "Whiteboard Extraction"
VIEW_DASHBOARD = "Analytics Dashboard"

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


def start_batch(batch_name_raw, folder_choice, new_folder_raw):
    """Validates the form, runs the batch, then stores the messages and resets the uploader."""
    if not uploaded_files:
        st.error("Please select at least one photo before extracting.")
        return

    filename = sanitize_batch_filename(batch_name_raw)
    if not filename:
        st.error("Please enter a name for this batch before extracting.")
        return

    subfolder, creating_new_folder, error = resolve_subfolder(folder_choice, new_folder_raw)
    if error:
        st.error(error)
        return

    local_dir = os.path.join(LOCAL_BATCH_DIR, subfolder) if subfolder else LOCAL_BATCH_DIR
    os.makedirs(local_dir, exist_ok=True)
    output_path = os.path.join(local_dir, filename)

    st.session_state.last_output_path = output_path
    st.session_state.last_output_label = batch_name_raw.strip()
    st.session_state.last_output_folder = subfolder
    st.session_state.last_folder_choice = subfolder if subfolder else MAIN_FOLDER_LABEL

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

    outcome = run_batch(uploaded_files, output_path, subfolder, creating_new_folder, on_progress)

    if outcome.created_folder:
        remember_new_folder(subfolder)

    results = outcome.messages
    results.append(("status", "Batch Complete."))
    results.append((
        "success",
        f"Successfully added {outcome.processed} new file(s). Skipped {outcome.skipped} duplicate(s).",
    ))
    st.session_state.batch_results = results
    st.session_state.uploader_key += 1  # clears the uploader
    st.rerun()


# Two columns, like the dashboard uses the full width: the form on the left, and
# everything that comes out of it (messages, download, sync status) on the right,
# so a long list of messages never pushes the form's buttons down the page.
form_col, results_col = st.columns([3, 2], gap="large")

# --- LEFT: UPLOADER AND BATCH FORM ---
with form_col:
    st.subheader("Upload Whiteboard Photos")
    uploaded_files = st.file_uploader(
        "Choose whiteboard images...",
        type=["jpg", "jpeg", "png"],
        accept_multiple_files=True,
        label_visibility="collapsed",
        key=f"uploader_{st.session_state.uploader_key}",
    )
    if uploaded_files:
        status_message(f"{len(uploaded_files)} image(s) selected.")

    # Suggest a fresh default name each time a new round of photos is selected,
    # without overwriting what the person is typing in the meantime.
    if st.session_state.get("batch_name_round") != st.session_state.uploader_key:
        st.session_state.batch_name_round = st.session_state.uploader_key
        st.session_state.batch_name_input = f"Batch_{datetime.datetime.now().strftime('%Y-%m-%d_%H%M')}"
        # Re-read the folder list each round, so folders other people created show up
        st.session_state.sp_folders = None

    folder_choice, new_folder_raw = render_folder_picker()
    batch_name_raw = st.text_input(
        "Name this batch (this becomes the Excel file name, locally and in SharePoint)",
        key="batch_name_input",
    )

    extract_clicked, clear_clicked = render_action_buttons()

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
        start_batch(batch_name_raw, folder_choice, new_folder_raw)

    render_file_management()
    render_pending_syncs()
