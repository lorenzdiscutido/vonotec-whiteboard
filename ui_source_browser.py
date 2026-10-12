# ui_source_browser.py
"""The folder browser used by "Process from SharePoint": lets the person walk
through the SharePoint document library, folder by folder, to the one that
holds the photos they want to process. It is separate from the "Save settings"
folder, which is only where the Excel file goes."""
import html

import streamlit as st

import sharepoint_client
from settings import PHOTO_BROWSE_ROOT

PATH_KEY = "src_path"          # the folder currently open, as a list of folder names
OPEN_KEY = "src_open"          # the "Open a subfolder" dropdown
RECURSIVE_KEY = "src_recursive"
STAY_LABEL = "(stay in this folder)"


def _root():
    return [s for s in PHOTO_BROWSE_ROOT.split("/") if s]


def _current_path():
    """The open folder; falls back to the starting folder if nothing is open yet."""
    root = _root()
    path = st.session_state.get(PATH_KEY)
    if path is None or list(path[:len(root)]) != root:
        path = root
        st.session_state[PATH_KEY] = path
    return list(path)


@st.cache_data(ttl=60, show_spinner=False)
def _cached_folders(path):
    """Subfolders of a folder, kept for a minute so clicking around stays quick."""
    return sharepoint_client.list_folders_at(list(path))


def _open_subfolder():
    """Runs when a subfolder is picked in the dropdown: step into it."""
    choice = st.session_state.get(OPEN_KEY)
    if choice and choice != STAY_LABEL:
        st.session_state[PATH_KEY] = _current_path() + [choice]
    st.session_state[OPEN_KEY] = STAY_LABEL


def _go_up():
    path = _current_path()
    if len(path) > len(_root()):
        st.session_state[PATH_KEY] = path[:-1]


def render_source_browser():
    """Draws the browser. Returns (folder_path, include_subfolders): the open
    folder as a list of folder names from the top of the library, and whether
    photos in the folders below it should be included too."""
    path = _current_path()
    shown = " / ".join(path) if path else "(top of the document library)"
    st.markdown(
        f'<p class="sync-name">Photos are read from: {html.escape(shown)}</p>',
        unsafe_allow_html=True,
    )

    try:
        subfolders = _cached_folders(tuple(path))
    except Exception as error:
        st.warning(f"Could not list the folders here right now: {error}")
        subfolders = []

    st.selectbox(
        "Open a subfolder",
        [STAY_LABEL] + list(subfolders),
        key=OPEN_KEY,
        on_change=_open_subfolder,
    )
    if not subfolders:
        st.caption("There are no subfolders in this folder.")

    if len(path) > len(_root()):
        st.button("Go up one folder", type="primary", key="src_up", on_click=_go_up)

    include_subfolders = st.toggle(
        "Include photos from all folders inside this one",
        key=RECURSIVE_KEY,
        help="Also reads the photos in every subfolder below the open folder "
             "(up to 300 folders).",
    )
    return path, include_subfolders
