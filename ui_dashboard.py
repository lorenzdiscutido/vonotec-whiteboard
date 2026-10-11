# ui_dashboard.py
"""The "Analytics Dashboard" tab: charts of defect frequency by Tower, Floor
and Material, built by reading the batch Excel files straight from SharePoint.

The Excel files are the source of truth (people correct them there), so the
charts always reflect the checked data. Each file is downloaded once and kept
in memory until SharePoint says it changed, so clicking around the filters
doesn't re-download anything.

The dashboard re-checks SharePoint on its own every DASHBOARD_REFRESH_SECONDS
(see settings.py). Only the dashboard section reloads, and only files that are
new or changed are downloaded, so it is quick and doesn't disturb the rest of
the page.
(The sheet parsing itself lives in dashboard_data.py.)"""
import os
import tempfile

import streamlit as st

import sharepoint_client
from dashboard_data import (
    parse_workbook, rows_to_dataframe, apply_filters, count_by, natural_key,
)
from settings import ENABLE_SHAREPOINT_SYNC, MAIN_FOLDER_LABEL, DASHBOARD_REFRESH_SECONDS

# Dropdown choice that reads the whole landing folder: its own files plus every subfolder's
ALL_FILES_LABEL = "Whiteboard Excel Files"
# What the Folder column shows for files that sit directly in the landing folder
MAIN_FOLDER_COLUMN = "(Main folder)"

CHART_COLOR = "#FF9933"        # light orange, matching the Vonotec logo
CHART_BACKGROUND = "#FFFFFF"   # set explicitly: the app's blue theme would otherwise show through
CHART_TEXT = "#1E293B"
CHART_GRID = "#E2E8F0"
# Folder/file lists are kept for just under one refresh interval, so every
# automatic check really asks SharePoint instead of reusing the previous answer.
LIST_CACHE_SECONDS = max(DASHBOARD_REFRESH_SECONDS - 1, 1)

# Files already downloaded in this run of the app, as (folder, name, modified).
# Used to show the "Reading Excel files..." bar only when something is really
# being downloaded, not on every quiet automatic check.
_loaded_keys = set()


# ---------------------------------------------------------------------------
# Cached SharePoint reads
# ---------------------------------------------------------------------------

@st.cache_data(ttl=LIST_CACHE_SECONDS, show_spinner=False)
def _cached_subfolders():
    return sharepoint_client.list_subfolders()


@st.cache_data(ttl=LIST_CACHE_SECONDS, show_spinner=False)
def _cached_file_list(subfolder):
    return sharepoint_client.list_xlsx_files(subfolder)


@st.cache_data(show_spinner=False, max_entries=1000)
def _cached_file_rows(subfolder, filename, modified):
    """Downloads one batch file and returns (rows, error_text). `modified` is
    SharePoint's last-modified time: it isn't used below, but it is part of the
    cache key, so a file that was edited is downloaded again while an unchanged
    one never is. A network failure raises (and is therefore not cached)."""
    with tempfile.TemporaryDirectory() as tmp_dir:
        local_path = os.path.join(tmp_dir, "batch.xlsx")
        if not sharepoint_client.download_file(filename, local_path, subfolder):
            return [], "it is no longer in SharePoint"
        try:
            return parse_workbook(local_path, subfolder or MAIN_FOLDER_COLUMN, filename), ""
        except Exception as error:
            return [], f"it could not be read ({error})"


def _refresh_lists():
    """Forgets the folder and file lists so new files show up. Files that were
    already downloaded stay cached; they are only re-downloaded if they changed."""
    _cached_subfolders.clear()
    _cached_file_list.clear()


# ---------------------------------------------------------------------------
# Loading the chosen scope
# ---------------------------------------------------------------------------

def _folders_for_scope(scope, subfolders):
    """The folders to read, where "" is the landing folder itself."""
    if scope == ALL_FILES_LABEL:
        return [""] + list(subfolders)
    if scope == MAIN_FOLDER_LABEL:
        return [""]
    return [scope]


def _load_scope(scope, subfolders):
    """Returns (rows, files_read, skipped), where skipped is a list of
    (filename, reason) for files that could not be used."""
    files = []
    for folder in _folders_for_scope(scope, subfolders):
        files += _cached_file_list(folder)

    def key_of(item):
        return (item["subfolder"], item["name"], item["modified"])

    # Only show a progress bar when at least one file is new or changed
    to_download = [item for item in files if key_of(item) not in _loaded_keys]
    progress = st.progress(0.0, text="Reading Excel files from SharePoint...") if to_download else None

    rows, skipped, files_read = [], [], 0
    for index, item in enumerate(files):
        try:
            file_rows, error = _cached_file_rows(item["subfolder"], item["name"], item["modified"])
            _loaded_keys.add(key_of(item))
        except Exception as network_error:
            file_rows, error = [], f"it could not be downloaded ({network_error})"

        label = f'{item["subfolder"]}/{item["name"]}' if item["subfolder"] else item["name"]
        if error:
            skipped.append((label, error))
        else:
            rows += file_rows
            files_read += 1
        if progress is not None:
            progress.progress((index + 1) / len(files), text=f"Reading Excel files... {index + 1} of {len(files)}")
    if progress is not None:
        progress.empty()
    return rows, files_read, skipped


# ---------------------------------------------------------------------------
# Charts
# ---------------------------------------------------------------------------

def _bar_chart(counts, column, title):
    """A bar chart of counts (from dashboard_data.count_by), keeping the order
    the table is already in."""
    import altair as alt  # installed together with Streamlit

    chart = (
        alt.Chart(counts)
        .mark_bar(color=CHART_COLOR, cornerRadiusTopLeft=3, cornerRadiusTopRight=3)
        .encode(
            # Padding leaves a gap between bars (and keeps a lone bar from filling the chart)
            x=alt.X(f"{column}:N", sort=list(counts[column]), title=column,
                    scale=alt.Scale(paddingInner=0.25, paddingOuter=0.25),
                    axis=alt.Axis(labelAngle=0, labelLimit=120)),
            y=alt.Y("Defects:Q", title="Number of defects", axis=alt.Axis(tickMinStep=1)),
            tooltip=[alt.Tooltip(f"{column}:N", title=column), alt.Tooltip("Defects:Q", title="Defects")],
        )
        .properties(height=280, width="container", background=CHART_BACKGROUND)
        .configure_axis(
            labelColor=CHART_TEXT, titleColor=CHART_TEXT,
            gridColor=CHART_GRID, domainColor=CHART_TEXT, tickColor=CHART_TEXT,
        )
        .configure_view(stroke=None)
    )
    st.markdown(f"#### {title}")
    st.altair_chart(chart, theme=None)  # theme=None keeps dark axis text on the light page


# ---------------------------------------------------------------------------
# The tab
# ---------------------------------------------------------------------------

def _sorted_options(series):
    return sorted(series.unique(), key=natural_key)


@st.fragment(run_every=DASHBOARD_REFRESH_SECONDS)
def render_analytics_dashboard():
    """A fragment: it re-runs by itself every DASHBOARD_REFRESH_SECONDS without
    reloading the rest of the page, and also when someone changes a widget
    inside it (the folder choice or a filter)."""
    st.subheader("Analytics Dashboard")

    if not ENABLE_SHAREPOINT_SYNC:
        st.info("The dashboard reads the Excel files stored in SharePoint, and SharePoint sync is switched off.")
        return

    # --- Which files to analyze ---
    try:
        subfolders = _cached_subfolders()
    except Exception as error:
        st.error(f"Could not load the SharePoint folders right now: {error}. It will try again automatically.")
        if st.button("Try again", type="primary", key="analytics_retry"):
            _refresh_lists()
            st.rerun()
        return

    options = [ALL_FILES_LABEL, MAIN_FOLDER_LABEL] + list(subfolders)
    pick_col, refresh_col = st.columns([3, 1], vertical_alignment="bottom")
    with pick_col:
        scope = st.selectbox(
            "Which files do you want to analyze?",
            options,
            key="analytics_scope",
            help=f'"{ALL_FILES_LABEL}" combines every Excel file in the landing folder and all of its folders.',
        )
    with refresh_col:
        if st.button("Refresh data", type="primary", key="analytics_refresh"):
            _refresh_lists()
            st.rerun()
    st.caption(
        f"Updates automatically: new or changed Excel files are picked up within "
        f"{DASHBOARD_REFRESH_SECONDS} seconds."
    )

    try:
        rows, files_read, skipped = _load_scope(scope, subfolders)
    except Exception as error:
        st.error(f"Could not read the file list from SharePoint right now: {error}. It will try again automatically.")
        return

    if skipped:
        names = "\n".join(f"- {name}: {reason}" for name, reason in skipped)
        st.warning(f"{len(skipped)} file(s) were left out of the charts because:\n\n{names}")

    if files_read == 0:
        st.info("There are no Excel files to analyze in this selection yet.")
        return

    df = rows_to_dataframe(rows)
    if df.empty:
        st.info(f"Read {files_read} file(s), but none of them contain any defects yet.")
        return

    # --- Filters (the keys include the choice, so switching folders resets them) ---
    st.markdown("**Filters** (leave empty to include everything)")
    folders_in_data = _sorted_options(df["Folder"])
    show_folder_filter = scope == ALL_FILES_LABEL and len(folders_in_data) > 1
    filter_cols = st.columns(4 if show_folder_filter else 3)
    towers = filter_cols[0].multiselect("Tower", _sorted_options(df["Tower"]), key=f"flt_tower_{scope}")
    floors = filter_cols[1].multiselect("Floor", _sorted_options(df["Floor"]), key=f"flt_floor_{scope}")
    materials = filter_cols[2].multiselect("Material", _sorted_options(df["Material"]), key=f"flt_material_{scope}")
    if show_folder_filter:
        folders = filter_cols[3].multiselect("Folder", folders_in_data, key=f"flt_folder_{scope}")
        if folders:
            df = df[df["Folder"].isin(folders)]
    df = apply_filters(df, towers, floors, materials)

    # --- Summary ---
    total_col, files_col, folders_col = st.columns(3)
    total_col.metric("Defects", f"{len(df):,}")
    files_col.metric("Excel files read", f"{files_read:,}")
    folders_col.metric("Folders", f"{df['Folder'].nunique():,}")

    if df.empty:
        st.info("No defects match these filters.")
        return

    # --- Charts ---
    _bar_chart(count_by(df, "Tower"), "Tower", "Defects by Tower")
    _bar_chart(count_by(df, "Floor"), "Floor", "Defects by Floor")
    _bar_chart(count_by(df, "Material"), "Material", "Defects by Material")

    with st.expander("View the data behind the charts"):
        st.dataframe(df, hide_index=True)
