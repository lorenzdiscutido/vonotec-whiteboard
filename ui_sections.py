# ui_sections.py
"""Reusable Streamlit sections for the Whiteboard app.
(Header and CSS stay in ui_components.py.)"""
import os
import traceback

import streamlit as st

import sharepoint_client
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


def show_batch_results():
    """Shows the messages from the last batch (kept even after the uploader is cleared)."""
    for kind, text in st.session_state.batch_results:
        if kind == "warning":
            st.warning(text)
        elif kind == "error":
            st.error(text)
        elif kind == "success":
            st.success(text)
        else:
            st.markdown(f'<p class="status-msg">{text}</p>', unsafe_allow_html=True)


def status_message(text):
    st.markdown(f'<p class="status-msg">{text}</p>', unsafe_allow_html=True)


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
        col1, col2 = st.columns(2, gap="small")
        with col1:
            extract_clicked = st.button("Extract Data & Update Excel", type="primary")
        with col2:
            clear_clicked = st.button("Clear All Photos", type="primary", key="clear_photos")
    return extract_clicked, clear_clicked


def render_file_management():
    """Download / delete controls for the file of the last batch."""
    path = st.session_state.last_output_path
    if not (path and os.path.exists(path)):
        return

    st.markdown("<br>", unsafe_allow_html=True)
    folder = st.session_state.last_output_folder
    folder_note = f' in folder "{folder}"' if folder else ""
    st.subheader(f'File Management — "{st.session_state.last_output_label}"{folder_note}')

    col1, col2 = st.columns(2)
    with col1:
        try:
            with open(path, "rb") as f:
                xlsx_bytes = f.read()
        except FileNotFoundError:
            xlsx_bytes = None
        if xlsx_bytes is not None:
            st.download_button(
                label=f"Download {os.path.basename(path)}",
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
                st.session_state.last_output_path = None
                st.session_state.last_output_label = None
                st.session_state.last_output_folder = ""
                st.session_state.confirm_reset = False
                st.rerun()
        with cancel_col:
            if st.button("Cancel", type="primary", key="cancel_delete"):
                st.session_state.confirm_reset = False
                st.rerun()
