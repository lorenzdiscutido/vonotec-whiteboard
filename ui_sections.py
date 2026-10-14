# ui_sections.py
"""Reusable Streamlit sections for the Whiteboard app.
(Header and CSS stay in ui_components.py.)"""
import io
import os
import html
import hashlib
import traceback

import streamlit as st
from PIL import Image, ImageOps

import sharepoint_client
import sync_tracker
from settings import ENABLE_SHAREPOINT_SYNC, MAIN_FOLDER_LABEL, NEW_FOLDER_LABEL

# The one final action, whichever photo source is used
PROCESS_BUTTON_LABEL = "Extract photos"

# Labels for the second-level output subfolder dropdown (Option 2 cascading picker)
NEW_SUBFOLDER_LABEL = "+ Create a new subfolder..."

_SESSION_DEFAULTS = {
    "uploader_key": 0,             # changing it resets the uploader (clears all photos)
    "batch_results": [],
    "batch_report": None,          # per-photo rows of the last batch (see show_batch_results)
    "confirm_reset": False,
    "last_output_path": None,
    "last_output_label": None,
    "last_output_folder": "",      # "" = the main landing folder, or "Parent" / "Parent/Child"
    "last_folder_choice": MAIN_FOLDER_LABEL,
    "sp_folders": None,            # None = not loaded yet
    "sp_folders_error": False,
    "sp_nested_folders": {},       # parent_folder -> list of child folder names
}


def init_session_state():
    for key, default in _SESSION_DEFAULTS.items():
        if key not in st.session_state:
            if isinstance(default, list):
                st.session_state[key] = list(default)
            elif isinstance(default, dict):
                st.session_state[key] = dict(default)
            else:
                st.session_state[key] = default


# With more messages than this, the old plain message list goes into a scrollable box
RESULTS_BOX_MIN_MESSAGES = 5
RESULTS_BOX_HEIGHT_PX = 320
RESULTS_PINNED_MESSAGES = 2

# The compact result rows scroll inside a box once there are more than this many
RESULT_ROWS_BOX_MIN = 8
RESULT_ROWS_BOX_HEIGHT_PX = 390

# Short icons for the result rows: green check, orange alert, gray skip, red cross
_ROW_ICONS = {
    "added": '<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="9"/><path d="m8.5 12.3 2.4 2.4 4.6-5"/></svg>',
    "review": '<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"><path d="M12 4 2.8 19.5h18.4L12 4Z"/><path d="M12 10v4.2M12 17h.01"/></svg>',
    "skipped": '<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="9"/><path d="M8 12h8"/></svg>',
    "failed": '<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="9"/><path d="m9 9 6 6m0-6-6 6"/></svg>',
}
_ICON_CHEVRON = (
    '<svg class="rs-chevron" width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" '
    'stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"><path d="m9 6 6 6-6 6"/></svg>'
)
_ICON_FILE_BADGE = (
    '<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" '
    'stroke-linecap="round" stroke-linejoin="round"><path d="M7 3.5h6l4 4V20H7V3.5Zm6 0v4h4"/></svg>'
)
_ICON_FOLDER_BADGE = (
    '<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" '
    'stroke-linecap="round" stroke-linejoin="round"><path d="M3.5 6.5h6l2-2h9v15h-17v-13Z"/></svg>'
)

_ROW_LABELS = {"added": "Added", "review": "Needs review", "skipped": "Skipped", "failed": "Failed"}
_ROW_ORDER = {"failed": 0, "review": 1, "skipped": 2, "added": 3}


def _show_message(kind, text):
    if kind == "warning":
        st.warning(text)
    elif kind == "error":
        st.error(text)
    elif kind == "success":
        st.success(text)
    else:
        st.markdown(f'<p class="status-msg">{text}</p>', unsafe_allow_html=True)


def _middle_truncate(name, limit=46):
    """Middle-truncates long filenames so both the beginning and unique ending/extension stay visible."""
    base = os.path.basename(str(name).replace("\\", "/"))
    if len(base) <= limit:
        return base
    stem, ext = os.path.splitext(base)
    tail_len = min(10, max(len(stem) // 3, 4))
    head_len = max(limit - len(ext) - tail_len - 1, 8)
    return f"{stem[:head_len]}\u2026{stem[-tail_len:]}{ext}"


def _summary_html(files):
    """Renders status metric cards: Processed, Added, Need review, Skipped, Failed."""
    count = {status: 0 for status in _ROW_LABELS}
    for item in files:
        count[item["status"]] = count.get(item["status"], 0) + 1

    processed_total = count["added"] + count["review"]
    cards = [
        f'<div class="rs-metric rs-metric-neutral"><span class="rs-metric-val">{processed_total}</span>'
        '<span class="rs-metric-lbl">Processed</span></div>'
    ]
    if count["added"]:
        cards.append(
            f'<div class="rs-metric rs-metric-added"><span class="rs-metric-val">{count["added"]}</span>'
            '<span class="rs-metric-lbl">Added</span></div>'
        )
    if count["review"]:
        cards.append(
            f'<div class="rs-metric rs-metric-review"><span class="rs-metric-val">{count["review"]}</span>'
            '<span class="rs-metric-lbl">Need review</span></div>'
        )
    if count["skipped"]:
        cards.append(
            f'<div class="rs-metric rs-metric-neutral"><span class="rs-metric-val">{count["skipped"]}</span>'
            '<span class="rs-metric-lbl">Skipped</span></div>'
        )
    if count["failed"]:
        cards.append(
            f'<div class="rs-metric rs-metric-failed"><span class="rs-metric-val">{count["failed"]}</span>'
            '<span class="rs-metric-lbl">Failed</span></div>'
        )
    return '<div class="rs-metrics-bar">' + "".join(cards) + '</div>'


def _row_html(item):
    """Two-line result row with expandable details."""
    status = item["status"] if item["status"] in _ROW_LABELS else "added"
    full_name = str(item["name"])
    short = _middle_truncate(full_name)
    reason = str(item.get("reason") or "")
    details = [line for line in str(item.get("details") or "").splitlines() if line.strip()]
    expandable = bool(details) or status in ("review", "failed", "skipped") or short != full_name

    action_cue = (
        f'<span class="rs-action-cue">{"Review" if status == "review" else "Details"}{_ICON_CHEVRON}</span>'
        if expandable else ""
    )

    head = (
        f'<div class="rs-top-line">'
        f'<span class="rs-icon rs-{status}">{_ROW_ICONS[status]}</span>'
        f'<span class="rs-name" title="{html.escape(full_name)}">{html.escape(short)}</span>'
        f'<span class="rs-chip rs-chip-{status}">{_ROW_LABELS[status]}</span>'
        f'{action_cue}'
        f'</div>'
        f'<div class="rs-sub-line">{html.escape(reason)}</div>'
    )

    if not expandable:
        return f'<div class="rs-row rs-row-{status}"><div class="rs-head">{head}</div></div>'

    body = f'<div class="rs-line"><b>Full file name:</b> {html.escape(full_name)}</div>'
    if details:
        body += "<ul>" + "".join(f"<li>{html.escape(line)}</li>" for line in details) + "</ul>"
    elif reason:
        body += f'<div class="rs-line">{html.escape(reason)}</div>'

    return (
        f'<details class="rs-row rs-row-{status}"><summary class="rs-head">{head}</summary>'
        f'<div class="rs-details">{body}</div></details>'
    )


def _show_report(report):
    """Shows the metric summary pills, general notices, and a two-line row per photo."""
    files = report.get("files") or []
    if files:
        st.markdown(_summary_html(files), unsafe_allow_html=True)
    for kind, text in report.get("notices") or []:
        if kind in ("warning", "error"):
            _show_message(kind, text)
        else:
            st.caption(text)
    if not files:
        return

    rows = sorted(files, key=lambda item: _ROW_ORDER.get(item["status"], 9))
    flagged_count = sum(1 for item in rows if item["status"] in ("review", "failed"))
    if flagged_count > 0:
        st.markdown(
            f'<div class="rs-group-label">{flagged_count} item{"s" if flagged_count != 1 else ""} '
            'requiring attention (click a row to inspect)</div>',
            unsafe_allow_html=True,
        )

    rows_html = '<div class="rs-list">' + "".join(_row_html(item) for item in rows) + "</div>"
    if len(rows) > RESULT_ROWS_BOX_MIN:
        with st.container(height=RESULT_ROWS_BOX_HEIGHT_PX, border=False, key="batch_files_box"):
            st.markdown(rows_html, unsafe_allow_html=True)
    else:
        st.markdown(rows_html, unsafe_allow_html=True)


def show_batch_results():
    """Shows the last batch: status metric cards and two-line expandable rows."""
    results = st.session_state.batch_results
    if not results:
        return
    report = st.session_state.get("batch_report")
    if report and (report.get("files") or report.get("notices")):
        _show_report(report)
        return

    if len(results) < RESULTS_BOX_MIN_MESSAGES:
        for kind, text in results:
            _show_message(kind, text)
        return

    scrolling = results[:-RESULTS_PINNED_MESSAGES]
    pinned = results[-RESULTS_PINNED_MESSAGES:]
    with st.container(height=RESULTS_BOX_HEIGHT_PX, border=True, key="batch_results_box"):
        for kind, text in scrolling:
            _show_message(kind, text)
    for kind, text in pinned:
        _show_message(kind, text)


def status_message(text):
    st.markdown(f'<p class="status-msg">{text}</p>', unsafe_allow_html=True)


# ---------------------------------------------------------------------------
# Uploader: hide the empty-state chrome once photos are chosen
# ---------------------------------------------------------------------------

_UPLOADER_FILLED_CSS = (
    "<style>"
    "[data-testid='stFileUploaderDropzone']::before, "
    "[data-testid='stFileUploadDropzone']::before, "
    "[data-testid='stFileUploaderDropzone']::after, "
    "[data-testid='stFileUploadDropzone']::after, "
    "[data-testid='stFileUploaderDropzoneInstructions'] "
    "{ display: none !important; } "
    "[data-testid='stFileUploaderDropzone'], [data-testid='stFileUploadDropzone'] "
    "{ min-height: 0 !important; padding: 12px 14px !important; flex-direction: column !important; align-items: flex-start !important; cursor: default !important; } "
    "div[data-testid='stFileUploader'] [data-testid='stFileUploaderDropzone'] > button[data-testid='stBaseButton-secondary'], "
    "div[data-testid='stFileUploader'] [data-testid='stFileUploaderDropzone'] > button "
    "{ position: static !important; inset: auto !important; width: auto !important; height: 34px !important; padding: 0 16px !important; "
    "opacity: 1 !important; font-size: 0 !important; background: #ffffff !important; border: 1px solid #cbd5e1 !important; border-radius: 7px !important; box-shadow: none !important; } "
    "div[data-testid='stFileUploader'] [data-testid='stFileUploaderDropzone'] > button * { display: none !important; } "
    "div[data-testid='stFileUploader'] [data-testid='stFileUploaderDropzone'] > button::after "
    "{ content: '+ Add more'; font-size: 13px; font-weight: 700; color: #1e3a8a; } "
    "</style>"
)


def hide_uploader_chrome_if_filled(uploaded_files):
    """Call right after st.file_uploader(...). Does nothing for an empty uploader."""
    if uploaded_files:
        st.markdown(_UPLOADER_FILLED_CSS, unsafe_allow_html=True)


# ---------------------------------------------------------------------------
# Photo previews
# ---------------------------------------------------------------------------

PREVIEW_SCROLL_HEIGHT = 290
PREVIEW_COLUMNS = 6
PREVIEW_PIXELS = 240


@st.cache_data(show_spinner=False, max_entries=300)
def _thumbnail_bytes(name, size, file_id, _data):
    """A small JPEG of one uploaded photo, or None if it can't be read."""
    try:
        with Image.open(io.BytesIO(_data)) as img:
            img.draft("RGB", (PREVIEW_PIXELS * 2, PREVIEW_PIXELS * 2))
            img = ImageOps.exif_transpose(img)
            img.thumbnail((PREVIEW_PIXELS, PREVIEW_PIXELS))
            out = io.BytesIO()
            img.convert("RGB").save(out, format="JPEG", quality=80)
            return out.getvalue()
    except Exception:
        return None


def _format_size(num_bytes):
    """12 345 678 bytes -> "11.8 MB"."""
    size = float(num_bytes)
    for unit in ("B", "KB", "MB"):
        if size < 1024:
            return f"{size:.0f} {unit}" if unit == "B" else f"{size:.1f} {unit}"
        size /= 1024
    return f"{size:.1f} GB"


def render_photo_previews(uploaded_files):
    """A count and a grid of small previews of the selected photos."""
    if not uploaded_files:
        return

    photos = list(uploaded_files)
    total_bytes = sum(getattr(photo, "size", 0) or 0 for photo in photos)
    st.caption(f"{len(photos)} photo{'s' if len(photos) != 1 else ''} selected \u00b7 {_format_size(total_bytes)}")

    def draw_grid():
        for start in range(0, len(photos), PREVIEW_COLUMNS):
            columns = st.columns(PREVIEW_COLUMNS)
            for column, photo in zip(columns, photos[start:start + PREVIEW_COLUMNS]):
                thumbnail = _thumbnail_bytes(photo.name, photo.size, getattr(photo, "file_id", ""), photo.getvalue())
                with column:
                    if thumbnail:
                        st.image(thumbnail)
                    else:
                        st.caption("(no preview)")

    if len(photos) > PREVIEW_COLUMNS:
        with st.container(height=PREVIEW_SCROLL_HEIGHT, border=False, key="thumbnail_scroll_gallery"):
            draw_grid()
    else:
        draw_grid()


# ---------------------------------------------------------------------------
# Cascading SharePoint output folder & subfolder picker (Option 2)
# ---------------------------------------------------------------------------

def refresh_folder_list():
    """Reads the current top-level subfolders from SharePoint."""
    try:
        st.session_state.sp_folders = sharepoint_client.list_subfolders()
        st.session_state.sp_folders_error = False
    except Exception:
        traceback.print_exc()
        st.session_state.sp_folders = []
        st.session_state.sp_folders_error = True


def _get_nested_folders(parent_folder):
    """Returns the cached list of subfolders inside `parent_folder`, fetching once per session."""
    nested_map = st.session_state.setdefault("sp_nested_folders", {})
    if parent_folder not in nested_map:
        try:
            nested_map[parent_folder] = sharepoint_client.list_subfolders_of(parent_folder)
        except Exception:
            nested_map[parent_folder] = []
    return nested_map[parent_folder]


def _direct_parent_label(parent_name):
    return f"(Save directly in {parent_name})"


def render_folder_picker(disabled=False):
    """Cascading 2-level output folder picker:
    1. Top-level folder ('Main folder', existing folders, or '+ Create a new folder...')
    2. When an existing folder (or new folder) is chosen, a second dropdown lets the user
       save directly in that folder, pick an existing subfolder, or '+ Create a new subfolder...'.
    Returns (folder_choice, new_folder_raw)."""
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

    top_folders = list(st.session_state.sp_folders or [])
    options = [MAIN_FOLDER_LABEL] + top_folders + [NEW_FOLDER_LABEL]

    # Split remembered "Parent/Child" choice into parent and child defaults
    remembered = str(st.session_state.last_folder_choice or MAIN_FOLDER_LABEL)
    remembered_parts = [p for p in remembered.split("/") if p] if remembered != MAIN_FOLDER_LABEL else []
    default_parent = remembered_parts[0] if remembered_parts else MAIN_FOLDER_LABEL
    default_child = remembered_parts[1] if len(remembered_parts) > 1 else ""

    parent_choice = st.selectbox(
        "Save into which SharePoint folder?",
        options,
        index=options.index(default_parent) if default_parent in options else 0,
        key="folder_choice_input",
        disabled=disabled,
    )

    # Case 1: Saving straight into the Main landing folder
    if parent_choice == MAIN_FOLDER_LABEL:
        return MAIN_FOLDER_LABEL, ""

    # Case 2: Creating a brand-new top-level folder (with optional new subfolder inside it)
    if parent_choice == NEW_FOLDER_LABEL:
        new_parent_raw = st.text_input(
            "Name for the new folder",
            key="new_folder_name_input",
            disabled=disabled,
        )
        new_sub_raw = st.text_input(
            "Subfolder inside the new folder (optional)",
            placeholder="Leave blank to save directly in the new folder",
            key="new_subfolder_under_new_input",
            disabled=disabled,
        )
        combined_new = f"{new_parent_raw.strip()}/{new_sub_raw.strip()}" if (new_sub_raw or "").strip() else new_parent_raw
        return NEW_FOLDER_LABEL, combined_new

    # Case 3: Existing top-level folder selected -> show cascading Subfolder dropdown
    child_folders = _get_nested_folders(parent_choice) if not disabled else []
    direct_label = _direct_parent_label(parent_choice)
    sub_options = [direct_label] + list(child_folders) + [NEW_SUBFOLDER_LABEL]
    default_sub_index = sub_options.index(default_child) if default_child in sub_options else 0

    parent_hash = hashlib.md5(parent_choice.encode("utf-8")).hexdigest()[:8]
    sub_choice = st.selectbox(
        f'Subfolder inside "{parent_choice}"',
        sub_options,
        index=default_sub_index,
        key=f"subfolder_choice_input_{parent_hash}",
        disabled=disabled,
        help="Save directly in this folder, choose an existing subfolder, or create a new subfolder inside it.",
    )

    if sub_choice == direct_label:
        return parent_choice, ""
    if sub_choice == NEW_SUBFOLDER_LABEL:
        new_sub_raw = st.text_input(
            f'Name for the new subfolder inside "{parent_choice}"',
            key=f"new_subfolder_name_input_{parent_hash}",
            disabled=disabled,
        )
        return f"{parent_choice}/{NEW_SUBFOLDER_LABEL}", new_sub_raw
    return f"{parent_choice}/{sub_choice}", ""


def resolve_subfolder(folder_choice, new_folder_raw):
    """Turns the cascading picker's answer into (subfolder_path, creating_new_folder, error_message)."""
    if not ENABLE_SHAREPOINT_SYNC:
        return "", False, None

    top_folders = list(st.session_state.sp_folders or [])

    # Creating a new top-level folder (and optionally a nested subfolder inside it)
    if folder_choice == NEW_FOLDER_LABEL:
        raw_parts = [p.strip() for p in (new_folder_raw or "").split("/") if p.strip()]
        if not raw_parts:
            return "", False, "Please enter a name for the new folder before extracting."
        clean_parent = sharepoint_client.sanitize_folder_name(raw_parts[0])
        if not clean_parent:
            return "", False, "Please enter a valid name for the new folder before extracting."
        matched_parent = next((f for f in top_folders if f.lower() == clean_parent.lower()), None)
        final_parent = matched_parent or clean_parent

        if len(raw_parts) > 1:
            clean_child = sharepoint_client.sanitize_folder_name(raw_parts[1])
            if not clean_child:
                return "", False, "Please enter a valid name for the new subfolder before extracting."
            return f"{final_parent}/{clean_child}", True, None
        return final_parent, matched_parent is None, None

    # Creating a new subfolder inside an existing parent folder
    if isinstance(folder_choice, str) and folder_choice.endswith(f"/{NEW_SUBFOLDER_LABEL}"):
        parent_folder = folder_choice[: -len(f"/{NEW_SUBFOLDER_LABEL}")]
        clean_child = sharepoint_client.sanitize_folder_name(new_folder_raw)
        if not clean_child:
            return "", False, f'Please enter a name for the new subfolder inside "{parent_folder}".'
        existing_children = _get_nested_folders(parent_folder)
        matched_child = next((c for c in existing_children if c.lower() == clean_child.lower()), None)
        final_child = matched_child or clean_child
        return f"{parent_folder}/{final_child}", matched_child is None, None

    if folder_choice != MAIN_FOLDER_LABEL:
        return folder_choice, False, None
    return "", False, None


def remember_new_folder(subfolder):
    """Adds a just-created top-level or nested subfolder to the session dropdown lists."""
    parts = [p for p in (subfolder or "").split("/") if p]
    if not parts:
        return
    parent = parts[0]
    folders = st.session_state.sp_folders or []
    if parent.lower() not in [f.lower() for f in folders]:
        st.session_state.sp_folders = sorted(folders + [parent], key=str.lower)

    if len(parts) > 1:
        child = parts[1]
        nested_map = st.session_state.setdefault("sp_nested_folders", {})
        children = list(nested_map.get(parent) or [])
        if child.lower() not in [c.lower() for c in children]:
            nested_map[parent] = sorted(children + [child], key=str.lower)


# ---------------------------------------------------------------------------
# Buttons and file management
# ---------------------------------------------------------------------------

PROCESS_VERB = "Extract"


def process_blockers(photo_count, source_chosen, folder_choice, new_folder_raw, batch_name_raw, remote):
    """Returns (photos_missing, destination_missing): a short sentence for each thing that is
    still missing, or "" when it is fine. The Extract button stays disabled until both are ""."""
    photos_missing = ""
    if remote:
        if not source_chosen:
            photos_missing = "Select a SharePoint folder to continue."
        elif photo_count <= 0:
            photos_missing = "No whiteboard photos found in the selected folder."
    else:
        if photo_count <= 0:
            photos_missing = "Upload at least one whiteboard photo to continue."

    destination_missing = ""
    if folder_choice == NEW_FOLDER_LABEL:
        raw_parent = (new_folder_raw or "").split("/")[0].strip()
        if not raw_parent:
            destination_missing = "Enter a name for the new folder."
    elif isinstance(folder_choice, str) and folder_choice.endswith(f"/{NEW_SUBFOLDER_LABEL}"):
        if not (new_folder_raw or "").strip():
            destination_missing = "Enter a name for the new subfolder."
    if not destination_missing and not (batch_name_raw or "").strip():
        destination_missing = "Enter a batch name to continue."
    return photos_missing, destination_missing


def _process_label(photo_count, source_chosen, remote, running=False, done=False):
    """Dynamic CTA button label following workflow progress."""
    if running:
        if photo_count:
            return f"Extracting {photo_count:,} photo{'s' if photo_count != 1 else ''}\u2026"
        return "Extracting\u2026"
    if done:
        return f"Re-extract {photo_count:,} photo{'s' if photo_count != 1 else ''}" if photo_count > 0 else "Start new extraction"
    if remote and not source_chosen:
        return "Select a folder to continue"
    if not remote and photo_count <= 0:
        return "Upload photos to continue"
    if photo_count > 0:
        return f"{PROCESS_VERB} {photo_count:,} photo{'s' if photo_count != 1 else ''}"
    return PROCESS_BUTTON_LABEL


def render_action_buttons(photo_count=0, source_chosen=False, blocked_reason="", remote=False, running=False, done=False):
    """The sticky action footer: primary action aligned to the right edge with requirement text above it."""
    with st.container(key="action_buttons"):
        label = _process_label(photo_count, source_chosen, remote, running, done)
        process_disabled = bool(blocked_reason) or running
        if blocked_reason:
            st.markdown(f'<p class="footer-hint">{html.escape(blocked_reason)}</p>', unsafe_allow_html=True)
        if remote:
            process_clicked = st.button(
                label, type="primary", key="fetch_folder_running" if running else "fetch_folder",
                disabled=process_disabled,
            )
            clear_clicked = False
        else:
            clear_col, process_col = st.columns(2, gap="small")
            with clear_col:
                clear_clicked = st.button(
                    "Clear All Photos",
                    type="primary",
                    key="clear_photos",
                    disabled=running or photo_count <= 0,
                )
            with process_col:
                process_clicked = st.button(
                    label, type="primary", key="extract_data_running" if running else "extract_data",
                    disabled=process_disabled,
                )
    return process_clicked, clear_clicked


def file_management_title_html(filename, folder):
    """The heading of the file card with clear file and folder badges."""
    folder_display = folder.replace("/", " / ") if folder else "Main folder"
    return (
        '<p class="eyebrow">File management</p>'
        '<div class="card-title">Generated workbook</div>'
        '<div class="file-meta-badges">'
        f'<span class="file-meta-pill" title="{html.escape(str(filename))}">{_ICON_FILE_BADGE}{html.escape(str(filename))}</span>'
        f'<span class="file-meta-pill file-meta-folder" title="{html.escape(str(folder_display))}">{_ICON_FOLDER_BADGE}{html.escape(str(folder_display))}</span>'
        '</div>'
    )


def render_file_management():
    """Download / Start new extraction / quiet Delete controls for the file of the last batch."""
    path = st.session_state.last_output_path
    if not (path and os.path.exists(path)):
        return

    with st.container(border=True, key="file_card"):
        st.markdown(
            file_management_title_html(os.path.basename(path), st.session_state.last_output_folder),
            unsafe_allow_html=True,
        )

        col1, col2 = st.columns(2, gap="small")
        with col1:
            try:
                with open(path, "rb") as f:
                    xlsx_bytes = f.read()
            except FileNotFoundError:
                xlsx_bytes = None
            if xlsx_bytes is not None:
                st.download_button(
                    label="Download master log",
                    data=xlsx_bytes,
                    file_name=os.path.basename(path),
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary",
                )
        with col2:
            if st.button("Start new extraction", type="primary", key="start_new_extraction"):
                st.session_state.uploader_key += 1
                st.session_state.batch_results = []
                st.session_state.batch_report = None
                st.session_state.last_remote_path = None
                st.session_state.confirm_reset = False
                st.rerun()

        with st.expander("More options", expanded=bool(st.session_state.confirm_reset)):
            st.caption("Remove the local working copy of this workbook from the app.")
            if not st.session_state.confirm_reset:
                if st.button("Delete local batch file", type="primary", key="start_fresh"):
                    st.session_state.confirm_reset = True
                    st.rerun()
            else:
                st.warning(
                    f'This will permanently delete the local copy of "{os.path.basename(path)}". It does not '
                    "touch the processed-photos duplicate log, and it does not delete any copy already "
                    "synced to SharePoint. Are you sure?"
                )
                yes_col, cancel_col = st.columns(2, gap="small")
                with yes_col:
                    if st.button("Yes, delete file", type="primary", key="confirm_delete"):
                        if os.path.exists(path):
                            os.remove(path)
                        sync_tracker.forget(path)
                        st.session_state.last_output_path = None
                        st.session_state.last_output_label = None
                        st.session_state.last_output_folder = ""
                        st.session_state.confirm_reset = False
                        st.rerun()
                with cancel_col:
                    if st.button("Cancel", type="primary", key="cancel_delete"):
                        st.session_state.confirm_reset = False
                        st.rerun()


# ---------------------------------------------------------------------------
# Pending Syncs: batch files that never reached SharePoint
# ---------------------------------------------------------------------------

def _push_pending(item):
    """Uploads one stranded file and leaves a message for the next screen draw."""
    with st.spinner(f"Pushing {item['filename']} to SharePoint..."):
        ok, reason = sync_tracker.sync_file(item["path"], item["subfolder"])
    if ok:
        st.session_state.sync_notice = ("success", f"'{item['filename']}' is now uploaded to SharePoint.")
    else:
        st.session_state.sync_notice = (
            "error", f"Could not push '{item['filename']}' to SharePoint: {reason}. You can try again."
        )
    st.rerun()


def render_pending_syncs():
    """Lists batch files that are not up to date in SharePoint, each with a
    Push button. Draws nothing when everything is synced."""
    if not ENABLE_SHAREPOINT_SYNC:
        return

    notice = st.session_state.pop("sync_notice", None)
    if notice:
        (st.success if notice[0] == "success" else st.error)(notice[1])

    pending = sync_tracker.pending_files()
    if not pending:
        return

    st.markdown("<br>", unsafe_allow_html=True)
    st.subheader("Pending syncs")
    st.warning(
        f"{len(pending)} file(s) are not up to date in SharePoint. Pushing a file replaces the "
        "SharePoint copy with the one stored here. These files are kept only on the app's "
        "temporary storage and are lost if the app restarts, so download any you want to keep."
    )

    for item in pending:
        folder_note = f' in folder "{item["subfolder"]}"' if item["subfolder"] else ""
        detail = item["reason"]
        if item["last_error"]:
            detail += f' (last attempt {item["last_attempt"]}: {item["last_error"]})'

        info_col, push_col, download_col = st.columns([3, 1.3, 1.3])
        with info_col:
            st.markdown(
                f'<p class="sync-name">{html.escape(item["filename"])}{html.escape(folder_note)}</p>'
                f'<p class="sync-detail">{html.escape(detail)}</p>',
                unsafe_allow_html=True,
            )
        widget_id = hashlib.md5(item["key"].encode("utf-8")).hexdigest()[:10]
        with push_col:
            if st.button("Push to SharePoint", type="primary", key=f"push_{widget_id}"):
                _push_pending(item)
        with download_col:
            try:
                with open(item["path"], "rb") as f:
                    file_bytes = f.read()
            except OSError:
                file_bytes = None
            if file_bytes is not None:
                st.download_button(
                    label="Download",
                    data=file_bytes,
                    file_name=item["filename"],
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary",
                    key=f"download_{widget_id}",
                )