# ui_dashboard.py
"""The "Analytics Dashboard" tab: charts of defect frequency by Tower, Floor
and Material, built by reading the batch Excel files straight from SharePoint.

The Excel files are the source of truth (people correct them there), so the
charts always reflect the checked data. Each file is downloaded once and kept
in memory until SharePoint says it changed, so clicking around the filters
doesn't re-download anything. app.py starts reading all the files in the
background as soon as the app is opened, so the data is normally ready before
anyone clicks over to this view (see dashboard_loader.py).

The page has two parts, drawn by app.py:
  render_dashboard_controls()   the sidebar: which files, Refresh, the filters
  render_analytics_dashboard()  the main area: the numbers and charts. It is a
                                "fragment", so it re-checks SharePoint on its own
                                every DASHBOARD_REFRESH_SECONDS (see settings.py)
                                without reloading the rest of the page. Only new
                                or changed files are downloaded, so that is quick.
They are separate because Streamlit does not let an auto-refreshing fragment draw
widgets in the sidebar. The two share the widgets' values through st.session_state.
(The sheet parsing itself lives in dashboard_data.py.)"""
import streamlit as st

import sharepoint_client
import ui_components
from dashboard_data import rows_to_dataframe, apply_filters, count_by, natural_key, UNKNOWN
from dashboard_loader import get_file_rows, is_loaded
from settings import ENABLE_SHAREPOINT_SYNC, MAIN_FOLDER_LABEL, DASHBOARD_REFRESH_SECONDS

# Dropdown choice that reads the whole landing folder: its own files plus every subfolder's
ALL_FILES_LABEL = "Whiteboard Excel Files"

# Chart look: Vonotec-blue bars on a white card. The card's own background is set in
# style.css (.st-key-analytics ... stVegaLiteChart) - keep CHART_BACKGROUND the same.
CHART_COLOR = "#2F5597"
ACCENT_COLOR = "#F97316"       # orange: marks the "Unknown" bars, so missing data stands out
CHART_BACKGROUND = "#FFFFFF"
CHART_TEXT = "#64748B"         # axis labels
CHART_GRID = "#EDF0F4"         # soft gridlines behind the bars
CHART_FONT = "DM Sans"

# With more bars than this, the x-axis labels are tilted so they don't run into each other
STRAIGHT_LABELS_MAX_BARS = 6

# One line under each chart title (the design's "Distribution across active levels" style)
CHART_SUBTITLES = {
    "Tower": "Distribution across towers",
    "Material": "Top classifications",
    "Floor": "Distribution across active levels",
}

# Column order of the data table: the useful ones first, the long text ones last
TABLE_COLUMNS = [
    "Tower", "Floor", "Material", "Defect Code", "Defect Type", "Dimension", "Unit",
    "Date", "Submitter", "Elevation", "Drop", "Folder", "File",
]

# The data table is drawn as HTML (see _data_table_html); this many rows at most
DATA_TABLE_MAX_ROWS = 1000

# Keys of the sidebar widgets, read again by the main area
SCOPE_KEY = "analytics_scope"
HIDE_UNKNOWN_KEY = "analytics_hide_unknown"
# Folder/file lists are kept for just under one refresh interval, so every
# automatic check really asks SharePoint instead of reusing the previous answer.
LIST_CACHE_SECONDS = max(DASHBOARD_REFRESH_SECONDS - 1, 1)


# ---------------------------------------------------------------------------
# Cached SharePoint reads
# ---------------------------------------------------------------------------

@st.cache_data(ttl=LIST_CACHE_SECONDS, show_spinner=False)
def _cached_subfolders():
    return sharepoint_client.list_subfolders()


@st.cache_data(ttl=LIST_CACHE_SECONDS, show_spinner=False)
def _cached_file_list(subfolder):
    return sharepoint_client.list_xlsx_files(subfolder)


def _refresh_lists():
    """Forgets the folder and file lists so new files show up. Files that were
    already downloaded stay in memory; they are only re-downloaded if they changed."""
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

    # Only show a progress bar when at least one file still has to be read
    # (not on quiet checks, and not when the background warm-up got there first)
    to_download = [i for i in files if not is_loaded(i["subfolder"], i["name"], i["modified"])]
    progress = st.progress(0.0, text="Reading Excel files from SharePoint...") if to_download else None

    rows, skipped, files_read = [], [], 0
    for index, item in enumerate(files):
        try:
            file_rows, error = get_file_rows(item["subfolder"], item["name"], item["modified"])
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

def _bar_chart(counts, column):
    """A bar chart of counts (from dashboard_data.count_by), keeping the order
    the table is already in."""
    import altair as alt  # installed together with Streamlit

    # Few bars: labels stay straight. Many bars: tilt them, and let Vega drop a label
    # only if it would still collide with its neighbour ("greedy").
    tilted = len(counts) > STRAIGHT_LABELS_MAX_BARS
    x_axis = alt.Axis(
        labelAngle=-45 if tilted else 0,
        labelAlign="right" if tilted else "center",
        labelBaseline="top",
        labelOverlap="greedy",
        labelLimit=160,
        labelPadding=6,
    )

    chart = (
        alt.Chart(counts)
        .mark_bar(cornerRadiusTopLeft=3, cornerRadiusTopRight=3)
        .encode(
            # Blue for known values, orange for the "Unknown" bar
            color=alt.condition(
                f"datum['{column}'] == '{UNKNOWN}'",
                alt.value(ACCENT_COLOR),
                alt.value(CHART_COLOR),
            ),
            # Padding leaves a gap between bars (and keeps a lone bar from filling the chart)
            x=alt.X(f"{column}:N", sort=list(counts[column]), title=None,
                    scale=alt.Scale(paddingInner=0.25, paddingOuter=0.25),
                    axis=x_axis),
            y=alt.Y("Defects:Q", title=None, axis=alt.Axis(tickMinStep=1)),
            tooltip=[alt.Tooltip(f"{column}:N", title=column), alt.Tooltip("Defects:Q", title="Number of defects")],
        )
        # padding: breathing room between the chart and the rounded card, so nothing touches the edge
        .properties(
            height=240, width="container", background=CHART_BACKGROUND,
            padding={"left": 4, "right": 12, "top": 10, "bottom": 4},
        )
        .configure(font=CHART_FONT)
        .configure_axis(
            labelColor=CHART_TEXT, labelFontSize=10,
            gridColor=CHART_GRID, domainColor="#DBE2EA", tickColor="#DBE2EA",
        )
        .configure_view(stroke=None)
    )
    st.altair_chart(chart, theme=None)  # theme=None keeps the colors above instead of the app theme's


def _chart_section(df, column, title, hide_unknown):
    """A white card with a title, a one-line subtitle, the total, and the bar chart of
    the defects per value of `column`. With hide_unknown, the "Unknown" bar is left out
    (and a note says how many defects that hides), so the real distribution isn't
    flattened by it."""
    counts = count_by(df, column)
    hidden = 0
    if hide_unknown:
        is_unknown = counts[column] == UNKNOWN
        hidden = int(counts.loc[is_unknown, "Defects"].sum())
        counts = counts[~is_unknown]

    total = int(counts["Defects"].sum()) if not counts.empty else 0
    with st.container(border=True, key=f"chart_{column.lower()}"):
        st.markdown(
            '<div class="chart-title"><div>'
            f'<div class="chart-name">{title}</div>'
            f'<div class="chart-sub">{CHART_SUBTITLES.get(column, "")}</div>'
            f'</div><span class="chart-total">{total:,} total</span></div>',
            unsafe_allow_html=True,
        )
        if counts.empty:
            st.caption(f"Nothing to show: every defect here has an unknown {column.lower()}.")
            return
        _bar_chart(counts, column)
        if (counts[column] == UNKNOWN).any():
            st.caption(f"The orange bar is 'Unknown': defects whose {column.lower()} wasn't on the board or in the file name.")
        if hidden:
            st.caption(f"{hidden:,} defect(s) with an unknown {column.lower()} are hidden.")


# ---------------------------------------------------------------------------
# Sidebar: which files, refresh, filters
# ---------------------------------------------------------------------------

def _filter_key(name, scope):
    """Filter widgets are keyed by the folder choice, so switching folders resets them."""
    return f"flt_{name}_{scope}"


def _sorted_options(series):
    return sorted(series.unique(), key=natural_key)


def _data_table_html(df):
    """The data table as plain HTML: a bold dark header and light rows (see .data-table
    in style.css). st.dataframe draws on a canvas that takes the app theme's dark blue
    and cannot be recolored with CSS, so this is the way to get a light table."""
    ordered = [c for c in TABLE_COLUMNS if c in df.columns] + [c for c in df.columns if c not in TABLE_COLUMNS]
    shown = df[ordered].head(DATA_TABLE_MAX_ROWS)
    table = shown.to_html(
        index=False, border=0, na_rep="", classes="data-table", escape=True,
        float_format=lambda value: f"{value:g}",
    )
    table = "".join(line.strip() for line in table.splitlines())  # no indentation, so Markdown leaves it alone
    return f'<div class="data-table-wrap">{table}</div>'


def render_dashboard_controls():
    """The dashboard's controls. Call it inside `with st.sidebar:`. Not a
    fragment: any change here reruns the whole page, so the charts follow."""
    st.markdown('<p class="side-title">Analytics controls</p>', unsafe_allow_html=True)

    if not ENABLE_SHAREPOINT_SYNC:
        st.caption("SharePoint sync is switched off, so there is nothing to analyze.")
        return

    try:
        subfolders = _cached_subfolders()
    except Exception as error:
        st.error(f"Could not load the SharePoint folders right now: {error}")
        if st.button("Try again", type="primary", key="analytics_retry"):
            _refresh_lists()
            st.rerun()
        return

    options = [ALL_FILES_LABEL, MAIN_FOLDER_LABEL] + list(subfolders)
    scope = st.selectbox(
        "Which files do you want to analyze?",
        options,
        key=SCOPE_KEY,
        help=f'"{ALL_FILES_LABEL}" combines every Excel file in the landing folder and all of its folders.',
    )
    if st.button("Refresh data", type="primary", key="analytics_refresh"):
        _refresh_lists()
        st.rerun()

    st.toggle(
        "Exclude 'Unknown' locations",
        key=HIDE_UNKNOWN_KEY,
        help="Hides the 'Unknown' bar in each chart (defects whose tower, floor or material "
             "couldn't be found), so the known ones are easier to compare.",
    )

    # The filter lists come from the data in the chosen folder(s). It is already in
    # memory, so this is quick; the charts re-check SharePoint on their own.
    try:
        rows, files_read, _skipped = _load_scope(scope, subfolders)
    except Exception:
        return  # the main area explains what went wrong
    df = rows_to_dataframe(rows)
    if df.empty:
        return

    st.divider()  # separates the global controls above from the filters below
    st.markdown('<p class="side-title">Filters</p>', unsafe_allow_html=True)
    st.caption("Leave a filter empty to include everything.")
    st.multiselect("Tower", _sorted_options(df["Tower"]), key=_filter_key("tower", scope))
    st.multiselect("Floor", _sorted_options(df["Floor"]), key=_filter_key("floor", scope))
    st.multiselect("Material", _sorted_options(df["Material"]), key=_filter_key("material", scope))
    folders_in_data = _sorted_options(df["Folder"])
    if scope == ALL_FILES_LABEL and len(folders_in_data) > 1:
        st.multiselect("Folder", folders_in_data, key=_filter_key("folder", scope))


# ---------------------------------------------------------------------------
# Main area: numbers and charts
# ---------------------------------------------------------------------------

@st.fragment(run_every=DASHBOARD_REFRESH_SECONDS)
def render_analytics_dashboard():
    """A fragment: it re-runs by itself every DASHBOARD_REFRESH_SECONDS without
    reloading the rest of the page. It reads the sidebar's choices from st.session_state."""
    ui_components.section_heading("Portfolio overview", "Extraction analytics")

    if not ENABLE_SHAREPOINT_SYNC:
        st.info("The dashboard reads the Excel files stored in SharePoint, and SharePoint sync is switched off.")
        return

    try:
        subfolders = _cached_subfolders()
    except Exception as error:
        st.error(f"Could not load the SharePoint folders right now: {error}. It will try again automatically.")
        return

    scope = st.session_state.get(SCOPE_KEY, ALL_FILES_LABEL)
    if scope not in [ALL_FILES_LABEL, MAIN_FOLDER_LABEL] + list(subfolders):
        scope = ALL_FILES_LABEL  # the chosen folder no longer exists
    hide_unknown = bool(st.session_state.get(HIDE_UNKNOWN_KEY, False))

    st.caption(
        f"Showing: {scope}. Updates automatically: new or changed Excel files are picked up "
        f"within {DASHBOARD_REFRESH_SECONDS} seconds."
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

    # --- Apply the sidebar filters ---
    folders = st.session_state.get(_filter_key("folder", scope)) or []
    if folders:
        df = df[df["Folder"].isin(folders)]
    df = apply_filters(
        df,
        st.session_state.get(_filter_key("tower", scope)),
        st.session_state.get(_filter_key("floor", scope)),
        st.session_state.get(_filter_key("material", scope)),
    )

    # --- Summary ---
    total_col, files_col, folders_col = st.columns(3)
    total_col.metric("Defects", f"{len(df):,}")
    files_col.metric("Excel files read", f"{files_read:,}")
    folders_col.metric("Folders", f"{df['Folder'].nunique():,}")

    if df.empty:
        st.info("No defects match these filters.")
        return

    # --- Charts: Tower beside Material, then Floor (the widest) across the page ---
    tower_col, material_col = st.columns(2, gap="large")
    with tower_col:
        _chart_section(df, "Tower", "Defects by tower", hide_unknown)
    with material_col:
        _chart_section(df, "Material", "Defects by material", hide_unknown)
    _chart_section(df, "Floor", "Defects by floor", hide_unknown)

    with st.expander("View the data behind the charts"):
        st.markdown(_data_table_html(df), unsafe_allow_html=True)
        if len(df) > DATA_TABLE_MAX_ROWS:
            st.caption(f"Showing the first {DATA_TABLE_MAX_ROWS:,} of {len(df):,} defects.")
