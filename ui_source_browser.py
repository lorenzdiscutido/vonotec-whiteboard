# ui_source_browser.py
"""The folder browser used by "Import from SharePoint", drawn as one card:
a header (folder icon, title, Connected status), a quick "Search for a folder"
box (to jump directly to any matching folder in SharePoint), or the cascading
"Photo folder" drill-down (each choice opens a selector for the folders inside
it, as deep as the folders go), and a footer that says how many photos the
chosen folder holds and when that was last checked.

Starts with no folder selected ("Select a folder...") so Step 1 -> Step 2
progressive disclosure is strictly respected."""
import hashlib
import html
import time

import streamlit as st

import sharepoint_client
from config import SHAREPOINT_SITE_NAME
from file_utils import load_prefs
from settings import PHOTO_BROWSE_ROOT

RECURSIVE_KEY = "src_recursive"
SEARCH_QUERY_KEY = "src_folder_search"
SEARCH_RESULT_KEY = "src_folder_search_pick"
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
    'stroke-linecap="round" stroke-linejoin="round"><path d="M4 5h16v14H4V5Zm0 11 4.5-4.5 3.2 3.2 2.3-2.3 6 5.6M15.5 9a1 1 0 1 0 0-2 1 1 0 1 0 0 2Z"/></svg>'
)


def _root():
    return [s for s in PHOTO_BROWSE_ROOT.split("/") if s]


@st.cache_data(ttl=60, show_spinner=False)
def _cached_folders(path):
    """Subfolders of a folder, kept for a minute so clicking around stays quick."""
    return sharepoint_client.list_folders_at(list(path))


@st.cache_data(ttl=60, show_spinner=False)
def _cached_folder_search(root_path, query_text):
    """Folders matching `query_text` anywhere under `root_path`, cached for 60s."""
    return sharepoint_client.search_folders_at(list(root_path), query_text)


@st.cache_data(ttl=COUNT_CACHE_SECONDS, show_spinner=False)
def _cached_photo_count(path, recursive):
    """(number of photos, number of folders that hold them, time it was counted)."""
    images = sharepoint_client.list_images_at(list(path), recursive=recursive)
    return len(images), len({image["folder"] for image in images}), time.time()


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
        '<div class="source-sub">Search by folder name or browse through the folders below.</div></div>'
        f'</div>{status}</div>'
    )


def _site_box_html():
    return (
        '<div class="source-field"><div class="source-field-label">SharePoint site</div>'
        f'<div class="source-field-box">{html.escape(SHAREPOINT_SITE_NAME.upper())}</div></div>'
    )


def _footer_html_and_count(path, has_selection, recursive):
    """Returns (footer_html, photo_count)."""
    if not has_selection:
        return (
            '<div class="source-footer"><span class="source-count">'
            f'{_ICON_IMAGE}Choose a folder above to see how many photos it holds.</span></div>',
            0,
        )
    try:
        with st.spinner("Counting photos..."):
            count, folder_count, counted_at = _cached_photo_count(tuple(path), bool(recursive))
    except Exception:
        return (
            '<div class="source-footer"><span class="source-count">'
            f'{_ICON_IMAGE}Could not count the photos right now.</span></div>',
            0,
        )

    if count == 0:
        label = "<b>No photos found</b> in the selected folder"
    else:
        label = f"<b>{count:,} photo{'s' if count != 1 else ''} found</b>"
        if recursive and folder_count > 1:
            label += f" across {folder_count:,} folders"
    return (
        '<div class="source-footer">'
        f'<span class="source-count">{_ICON_IMAGE}{label}</span>'
        f'<span class="source-updated">Last updated {_ago(counted_at)}</span></div>',
        count,
    )


def render_source_browser():
    """Draws the folder card. Returns (folder_path, include_subfolders, photo_count):
    the chosen folder as a list of folder names from the top of the document library
    ([] if none is chosen yet), whether photos in subfolders are included, and how many
    photos were found in the selection."""
    prefs = load_prefs()
    root = _root()
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
            return [], include_subfolders, 0

        # Top row: SharePoint site on the left, quick folder search box on the right
        top_cols = st.columns(LEVELS_PER_ROW, gap="medium")
        with top_cols[0]:
            st.markdown(_site_box_html(), unsafe_allow_html=True)
        with top_cols[1]:
            search_query = st.text_input(
                "Search folder by name",
                placeholder="Type a folder name (e.g. Tower 1, Drop 7)...",
                key=SEARCH_QUERY_KEY,
                help="Type part of a folder name to search across SharePoint, or leave blank to browse level by level.",
            )

        cleaned_search = (search_query or "").strip()
        if cleaned_search:
            # Search mode: show all matching SharePoint folders in a single searchable dropdown
            try:
                with st.spinner("Searching SharePoint folders..."):
                    matches = _cached_folder_search(tuple(root), cleaned_search)
            except Exception as error:
                st.warning(f"Could not search SharePoint folders right now: {error}")
                matches = []

            if not matches:
                st.caption(f'No folders matching "{cleaned_search}" were found.')
                path = []
            else:
                label_to_segments = {" / ".join(segs[len(root):] or segs): segs for segs in matches}
                search_options = [PLACEHOLDER_TOP] + list(label_to_segments.keys())
                picked_label = st.selectbox(
                    f"Matching folders ({len(matches)})",
                    search_options,
                    index=0,
                    key=f"{SEARCH_RESULT_KEY}_{hashlib.md5(cleaned_search.encode('utf-8')).hexdigest()[:8]}",
                )
                if picked_label != PLACEHOLDER_TOP and picked_label in label_to_segments:
                    path = label_to_segments[picked_label]
                    selections = path[len(root):]
                else:
                    path = []
        else:
            # Standard drill-down mode: cascading level-by-level dropdowns
            # Note: Streamlit's selectbox also supports typing directly inside it to filter that level!
            columns = st.columns(LEVELS_PER_ROW, gap="medium")
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

                if level > 0 and level % LEVELS_PER_ROW == 0:
                    columns = st.columns(LEVELS_PER_ROW, gap="medium")
                label = "Photo folder" if level == 0 else f"Subfolder (level {level})"
                with columns[level % LEVELS_PER_ROW]:
                    choice = st.selectbox(
                        label, options, index=0, key=_level_key(level, current),
                    )

                if choice == placeholder:
                    break
                selections.append(choice)

            path = (root + selections) if selections else []

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

        footer_html, photo_count = _footer_html_and_count(path, bool(selections), include_subfolders)
        st.markdown(footer_html, unsafe_allow_html=True)

    return path, include_subfolders, photo_count