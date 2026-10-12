# ui_source_browser.py
"""The folder browser used by "Process from SharePoint": a cascading drill-down
that works like a file explorer. The first selector lists the top-level folders;
choosing one opens a second selector with its subfolders, and so on, as deep as
the folders go. The last folder chosen is where the photos are read from.

It is separate from the "Save settings" folder, which is only where the Excel
file goes. The last location used is remembered in user_prefs.json (saved by
app.py when a batch starts) and pre-fills the selectors the next time."""
import hashlib
import html

import streamlit as st

import sharepoint_client
from file_utils import load_prefs
from settings import PHOTO_BROWSE_ROOT

RECURSIVE_KEY = "src_recursive"
PLACEHOLDER_TOP = "Select a folder..."
PLACEHOLDER_HERE = "(this folder)"
LEVELS_PER_ROW = 3     # selectors side by side before starting a new row
MAX_LEVELS = 12        # safety stop for very deep folder trees


def _root():
    return [s for s in PHOTO_BROWSE_ROOT.split("/") if s]


@st.cache_data(ttl=60, show_spinner=False)
def _cached_folders(path):
    """Subfolders of a folder, kept for a minute so clicking around stays quick."""
    return sharepoint_client.list_folders_at(list(path))


def _saved_path():
    """The last photo folder used, from user_prefs.json (a list of folder names), or []."""
    saved = load_prefs().get("photo_source_path")
    if not isinstance(saved, list):
        return []
    return [part for part in saved if isinstance(part, str)]


def _level_key(level, current_path):
    """A selector's widget key includes the folders above it, so when a higher
    selector changes, every selector below it starts fresh instead of keeping a
    choice that belongs to another folder."""
    digest = hashlib.md5("/".join(current_path).encode("utf-8")).hexdigest()[:8]
    return f"src_lv_{level}_{digest}"


def render_source_browser():
    """Draws the cascading selectors. Returns (folder_path, include_subfolders):
    the chosen folder as a list of folder names from the top of the document
    library ([] if none is chosen yet), and whether photos in the folders below
    it should be included too."""
    prefs = load_prefs()
    root = _root()
    saved = _saved_path()
    saved_below_root = saved[len(root):] if saved[:len(root)] == root else []

    selections = []
    columns = None
    for level in range(MAX_LEVELS):
        current = root + selections
        try:
            folders = _cached_folders(tuple(current))
        except Exception as error:
            st.warning(f"Could not list the folders here right now: {error}")
            break
        if not folders:
            if level == 0:
                st.caption("No folders were found to choose from.")
            break

        placeholder = PLACEHOLDER_TOP if level == 0 else PLACEHOLDER_HERE
        options = [placeholder] + list(folders)

        # Pre-fill from the remembered location, but only while it still matches
        # what is chosen above this selector and the folder still exists.
        default_index = 0
        if (saved_below_root[:level] == selections
                and level < len(saved_below_root)
                and saved_below_root[level] in folders):
            default_index = options.index(saved_below_root[level])

        if level % LEVELS_PER_ROW == 0:
            columns = st.columns(LEVELS_PER_ROW)
        label = "Top-level folder" if level == 0 else f"Subfolder (level {level})"
        with columns[level % LEVELS_PER_ROW]:
            choice = st.selectbox(
                label, options, index=default_index, key=_level_key(level, current),
            )

        if choice == placeholder:
            break
        selections.append(choice)

    path = root + selections
    if selections:
        st.markdown(
            f'<p class="sync-name">Photos are read from: {html.escape(" / ".join(path))}</p>',
            unsafe_allow_html=True,
        )
    else:
        st.caption("Choose a folder above. Each choice opens a selector for the folders inside it.")

    include_subfolders = st.toggle(
        "Include photos from all folders inside this one",
        value=bool(prefs.get("include_subfolders", False)),
        key=RECURSIVE_KEY,
        help="Also reads the photos in every subfolder below the chosen folder "
             "(up to 300 folders).",
    )
    return path, include_subfolders
