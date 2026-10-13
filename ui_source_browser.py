# ui_source_browser.py
"""The folder browser used by "Process from SharePoint", drawn as one card:
a header (folder icon, title, Connected status), the SharePoint site and a
cascading "Photo folder" drill-down (each choice opens a selector for the
folders inside it, as deep as the folders go), and a footer that says how many
photos the chosen folder holds and when that was last checked.

It is separate from the "Save settings" folder, which is only where the Excel
file goes. The last location used is remembered in user_prefs.json (saved by
app.py when a batch starts) and pre-fills the selectors the next time."""
import hashlib
import html
import time

import streamlit as st

import sharepoint_client
from config import SHAREPOINT_SITE_NAME
from file_utils import load_prefs
from settings import PHOTO_BROWSE_ROOT

RECURSIVE_KEY = "src_recursive"
PLACEHOLDER_TOP = "Select a folder..."
PLACEHOLDER_HERE = "(this folder)"
LEVELS_PER_ROW = 2     # selectors side by side (the site box takes the first slot)
MAX_LEVELS = 12        # safety stop for very deep folder trees
COUNT_CACHE_SECONDS = 180

_ICON_FOLDER = (
    '<svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" '
    'stroke-linecap="round" stroke-linejoin="round"><path d="M3.5 6.5h6l2-2h9v15h-17v-13Z"/></svg>'
)
_ICON_IMAGE = (
    '<svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" '
    'stroke-linecap="round" stroke-linejoin="round"><path d="M4 5h16v14H4V5Zm0 11 4.5-4.5 3.2 3.2 2.3-2.3 6 5.6M15.5 9a1 1 0 1 0 0-2 1 1 0 0 0 0 2Z"/></svg>'
)


def _root():
    return [s for s in PHOTO_BROWSE_ROOT.split("/") if s]


@st.cache_data(ttl=60, show_spinner=False)
def _cached_folders(path):
    """Subfolders of a folder, kept for a minute so clicking around stays quick."""
    return sharepoint_client.list_folders_at(list(path))


@st.cache_data(ttl=COUNT_CACHE_SECONDS, show_spinner=False)
def _cached_photo_count(path, recursive):
    """(number of photos, time it was counted) for a folder."""
    images = sharepoint_client.list_images_at(list(path), recursive=recursive)
    return len(images), time.time()


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


def _ago(timestamp):
    seconds = max(int(time.time() - timestamp), 0)
    if seconds < 60:
        return "just now"
    minutes = seconds // 60
    if minutes < 60:
        return f"{minutes} minute{'s' if minutes != 1 else ''} ago"
    hours = minutes // 60
    return f"{hours} hour{'s' if hours != 1 else ''} ago"


def _header_html(connected):
    status = (
        '<span class="source-status"><span></span>Connected</span>' if connected
        else '<span class="source-status error"><span></span>Not connected</span>'
    )
    return (
        '<div class="source-head"><div class="source-head-left">'
        f'<span class="source-icon">{_ICON_FOLDER}</span>'
        '<div><div class="source-title">Choose a SharePoint folder</div>'
        '<div class="source-sub">Select the folder containing the whiteboard photos.</div></div>'
        f'</div>{status}</div>'
    )


def _site_box_html():
    return (
        '<div class="source-field"><div class="source-field-label">SharePoint site</div>'
        f'<div class="source-field-box">{html.escape(SHAREPOINT_SITE_NAME)}</div></div>'
    )


def _footer_html(path, has_selection, recursive):
    if not has_selection:
        return (
            '<div class="source-footer"><span class="source-count">'
            f'{_ICON_IMAGE}Choose a folder above to see how many photos it holds.</span></div>'
        )
    try:
        with st.spinner("Counting photos..."):
            count, counted_at = _cached_photo_count(tuple(path), bool(recursive))
    except Exception:
        return (
            '<div class="source-footer"><span class="source-count">'
            f'{_ICON_IMAGE}Could not count the photos right now.</span></div>'
        )

    if count == 0:
        label = "<b>No photos found</b> in the selected folder"
    else:
        label = f"<b>{count:,} photo{'s' if count != 1 else ''} found</b> in the selected folder"
    return (
        '<div class="source-footer">'
        f'<span class="source-count">{_ICON_IMAGE}{label}</span>'
        f'<span class="source-updated">Updated {_ago(counted_at)}</span></div>'
    )


def render_source_browser():
    """Draws the folder card. Returns (folder_path, include_subfolders): the chosen
    folder as a list of folder names from the top of the document library ([] if
    none is chosen yet), and whether photos in the folders below it should be
    included too."""
    prefs = load_prefs()
    root = _root()
    saved = _saved_path()
    saved_below_root = saved[len(root):] if saved[:len(root)] == root else []
    include_subfolders = bool(prefs.get("include_subfolders", False))

    # Reaching the top-level folder list is what "Connected" means
    connection_error = None
    try:
        _cached_folders(tuple(root))
    except Exception as error:
        connection_error = error

    selections = []
    with st.container(border=True, key="source_card"):
        st.markdown(_header_html(connection_error is None), unsafe_allow_html=True)

        if connection_error is not None:
            st.warning(f"Could not reach SharePoint right now: {connection_error}")
            return [], include_subfolders

        columns = st.columns(LEVELS_PER_ROW, gap="medium")
        with columns[0]:
            st.markdown(_site_box_html(), unsafe_allow_html=True)

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

            slot = level + 1  # slot 0 is the SharePoint site box
            if slot % LEVELS_PER_ROW == 0:
                columns = st.columns(LEVELS_PER_ROW, gap="medium")
            label = "Photo folder" if level == 0 else f"Subfolder (level {level})"
            with columns[slot % LEVELS_PER_ROW]:
                choice = st.selectbox(
                    label, options, index=default_index, key=_level_key(level, current),
                )

            if choice == placeholder:
                break
            selections.append(choice)

        path = root + selections
        if selections:
            st.markdown(
                f'<p class="source-path">Photos are read from: {html.escape(" / ".join(path))}</p>',
                unsafe_allow_html=True,
            )

        include_subfolders = st.toggle(
            "Include photos from all folders inside this one",
            value=include_subfolders,
            key=RECURSIVE_KEY,
            help="Also reads the photos in every subfolder below the chosen folder "
                 "(up to 300 folders).",
        )

        st.markdown(_footer_html(path, bool(selections), include_subfolders), unsafe_allow_html=True)

    return path, include_subfolders
