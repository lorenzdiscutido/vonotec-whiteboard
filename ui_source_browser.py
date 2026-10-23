"""The folder browser used by "Import from SharePoint", drawn as one card:
a header (folder icon, title, Connected status), the SharePoint site and a
cascading "Photo folder" drill-down where every field supports instant
type-to-search (via index=None + placeholder), and a footer that says how many
photos the chosen folder holds and when that was last checked.

Pre-fills "01 - Projects > 04 - Execution" by default on first open, while keeping
every level editable so users can switch to any other folder."""
import hashlib
import html
import time

import streamlit as st

import sharepoint_client
from config import SHAREPOINT_SITE_NAME
from file_utils import load_prefs
from settings import PHOTO_BROWSE_ROOT

RECURSIVE_KEY = "src_recursive"
PLACEHOLDER_TOP = "Type to search or select a folder..."
PLACEHOLDER_SUB = "Type to search subfolder (optional)..."
LEVELS_PER_ROW = 2     # selectors side by side
MAX_LEVELS = 12        # safety stop for very deep folder trees
COUNT_CACHE_SECONDS = 180

# Usual starting path where whiteboard photos are stored
DEFAULT_SOURCE_PATH = ("01 - Projects", "04 - Execution")

# Uniform Lucide 24x24 folder & image icons (stroke-width="2")
_ICON_FOLDER = (
    '<svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" '
    'stroke-linecap="round" stroke-linejoin="round">'
    '<path d="M20 20a2 2 0 0 0 2-2V8a2 2 0 0 0-2-2h-7.9a2 2 0 0 1-1.69-.9L9.6 3.9A2 2 0 0 0 7.93 3H4a2 2 0 0 0-2 2v13a2 2 0 0 0 2 2Z"/></svg>'
)
_ICON_IMAGE = (
    '<svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" '
    'stroke-linecap="round" stroke-linejoin="round">'
    '<rect width="18" height="18" x="3" y="3" rx="2" ry="2"/>'
    '<circle cx="9" cy="9" r="2"/>'
    '<path d="m21 15-3.086-3.086a2 2 0 0 0-2.828 0L6 21"/></svg>'
)


def _root():
    return [s for s in PHOTO_BROWSE_ROOT.split("/") if s]


@st.cache_data(ttl=60, show_spinner=False)
def _cached_folders(path):
    """Subfolders of a folder, kept for a minute so clicking around stays quick."""
    return sharepoint_client.list_folders_at(list(path))


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
        f'<span class="source-status"><span></span>Connected to {html.escape(SHAREPOINT_SITE_NAME.upper())}</span>' if connected
        else '<span class="source-status error"><span></span>Not connected</span>'
    )
    return (
        '<div class="source-head"><div class="source-head-left">'
        f'<span class="source-icon">{_ICON_FOLDER}</span>'
        '<div><div class="source-title">Choose a SharePoint folder</div>'
        '<div class="source-sub">Click any folder field and type to search, or pick from the list.</div></div>'
        f'</div>{status}</div>'
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

            slot = level
            if slot % LEVELS_PER_ROW == 0:
                columns = st.columns(LEVELS_PER_ROW, gap="medium")

            label = "Photo folder" if level == 0 else f"Subfolder (level {level})"
            placeholder = PLACEHOLDER_TOP if level == 0 else PLACEHOLDER_SUB

            # Pre-fill "01 - Projects" -> "04 - Execution" on initial load while still
            # allowing the user to clear (x) or change any level freely.
            default_index = None
            if (
                level < len(DEFAULT_SOURCE_PATH)
                and selections == list(DEFAULT_SOURCE_PATH[:level])
                and DEFAULT_SOURCE_PATH[level] in folders
            ):
                default_index = folders.index(DEFAULT_SOURCE_PATH[level])

            with columns[slot % LEVELS_PER_ROW]:
                choice = st.selectbox(
                    label,
                    list(folders),
                    index=default_index,
                    placeholder=placeholder,
                    key=_level_key(level, current),
                )

            if not choice:
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
