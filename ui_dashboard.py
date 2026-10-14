# ui_dashboard.py
"""The "Analytics Dashboard" tab: charts of defect frequency by Tower, Floor
and Material (with the exact value on each bar), and a table of the total dimension and
total repair dimension per defect, built by reading the batch Excel files straight from SharePoint.

The Excel files are the source of truth (people correct them there), so the
charts always reflect the checked data. Each file is downloaded once and kept
in memory until SharePoint says it changed, so clicking around the filters
doesn't re-download anything. app.py starts reading all the files in the
background as soon as the app is opened, so the data is normally ready before
anyone clicks over to this view (see dashboard_loader.py).

The page has two parts, side by side in one card (see render_dashboard_view):
  render_dashboard_controls()   the controls panel on the left: which files, Refresh, the filters
  render_analytics_dashboard()  the main area: the numbers and charts. It is a
                                "fragment", so it re-checks SharePoint on its own
                                every DASHBOARD_REFRESH_SECONDS (see settings.py)
                                without reloading the rest of the page. Only new
                                or changed files are downloaded, so that is quick.
They are separate because Streamlit does not let an auto-refreshing fragment draw
widgets outside itself. The two share the widgets' values through st.session_state.
(The sheet parsing itself lives in dashboard_data.py.)"""
import io
import html

import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill
import streamlit as st

import sharepoint_client
import ui_components
from dashboard_data import rows_to_dataframe, apply_filters, count_by, dimension_by_defect, natural_key, UNKNOWN
from dashboard_loader import get_file_rows, is_loaded
from defect_rules import BOARD_UNITS
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
    "Repair Dimension", "Repair Unit", "Date", "Submitter", "Elevation", "Drop", "Folder", "File",
]

# Columns shown in the "Records needing review (missing dimensions)" queue table
REVIEW_QUEUE_COLUMNS = [
    "Folder", "File", "Tower", "Floor", "Drop", "Elevation",
    "Material", "Defect Code", "Defect Type", "Date", "Submitter",
]

# The data table is drawn as HTML (see _data_table_html); this many rows at most
DATA_TABLE_MAX_ROWS = 1000

# Keys of the sidebar widgets, read again by the main area
SCOPE_KEY = "analytics_scope"
HIDE_UNKNOWN_KEY = "analytics_hide_unknown"
# Folder/file lists are kept for just under one refresh interval, so every
# automatic check really asks SharePoint instead of reusing the previous answer.
LIST_CACHE_SECONDS = max(DASHBOARD_REFRESH_SECONDS - 1, 1)

# Inline SVG icons for the 4 KPI cards
_KPI_ICON_DEFECTS = (
    '<svg width="19" height="19" viewBox="0 0 24 24" fill="none" stroke="currentColor" '
    'stroke-width="1.9" stroke-linecap="round" stroke-linejoin="round">'
    '<path d="M5 20V10m7 10V4m7 16v-7"/></svg>'
)
_KPI_ICON_LINEAR = (
    '<svg width="19" height="19" viewBox="0 0 24 24" fill="none" stroke="currentColor" '
    'stroke-width="1.9" stroke-linecap="round" stroke-linejoin="round">'
    '<path d="M3 12h18M3 12l3.5-3.5M3 12l3.5 3.5M21 12l-3.5-3.5M21 12l-3.5 3.5"/></svg>'
)
_KPI_ICON_AREA = (
    '<svg width="19" height="19" viewBox="0 0 24 24" fill="none" stroke="currentColor" '
    'stroke-width="1.9" stroke-linecap="round" stroke-linejoin="round">'
    '<rect x="4" y="4" width="16" height="16" rx="2"/><path d="M4 12h16M12 4v16"/></svg>'
)
_KPI_ICON_WARN = (
    '<svg width="19" height="19" viewBox="0 0 24 24" fill="none" stroke="currentColor" '
    'stroke-width="1.9" stroke-linecap="round" stroke-linejoin="round">'
    '<path d="M12 4 2.8 19.5h18.4L12 4Z"/><path d="M12 10v4.2M12 17h.01"/></svg>'
)
_KPI_ICON_OK = (
    '<svg width="19" height="19" viewBox="0 0 24 24" fill="none" stroke="currentColor" '
    'stroke-width="1.9" stroke-linecap="round" stroke-linejoin="round">'
    '<circle cx="12" cy="12" r="9"/><path d="m8.5 12.3 2.4 2.4 4.6-5"/></svg>'
)


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
# KPI Cards & Missing-Dimension Review Queue
# ---------------------------------------------------------------------------

def _base_defect_code(code):
    parts = str(code or "").split()
    return parts[0] if parts else ""


def _missing_dimension_rows(df):
    """Returns the subset of `df` where a defect requires a dimension on the board
    (per BOARD_UNITS in config.py) but its Dimension is missing/NaN."""
    if df.empty:
        return df
    requires_dim = df["Defect Code"].apply(lambda c: bool(BOARD_UNITS.get(_base_defect_code(c), "")))
    return df[requires_dim & df["Dimension"].isna()]


def _kpi_card_html(label, value_text, unit_text, sub_text, icon_svg, tone="blue"):
    """Renders one enterprise KPI card with an icon tile, label, 24px bold value + unit,
    and a secondary context line underneath."""
    unit_html = f'<span class="kpi-unit">{html.escape(unit_text)}</span>' if unit_text else ""
    return (
        f'<div class="kpi-card kpi-{html.escape(tone)}">'
        f'<span class="kpi-icon kpi-icon-{html.escape(tone)}">{icon_svg}</span>'
        '<div class="kpi-body">'
        f'<div class="kpi-label">{html.escape(label)}</div>'
        f'<div class="kpi-value"><span>{html.escape(value_text)}</span>{unit_html}</div>'
        f'<div class="kpi-sub">{html.escape(sub_text)}</div>'
        '</div></div>'
    )


def _render_kpi_cards(df, summary, files_read):
    """Draws the 4 Option B KPI cards across the top of the dashboard:
      1. Total defects
      2. Linear repair total (LM) + board LM subtitle
      3. Area repair total (SQ LM) + board SQ LM subtitle
      4. Missing dimensions (Needs review)"""
    linear_rows = summary[summary["Repair Unit"] == "LM"] if not summary.empty else summary
    area_rows = summary[summary["Repair Unit"] == "SQ LM"] if not summary.empty else summary

    linear_repair = float(linear_rows["Total Repair Dimension"].dropna().sum()) if not linear_rows.empty else 0.0
    linear_board = float(linear_rows["Total Dimension"].dropna().sum()) if not linear_rows.empty else 0.0

    area_repair = float(area_rows["Total Repair Dimension"].dropna().sum()) if not area_rows.empty else 0.0
    area_board = float(area_rows["Total Dimension"].dropna().sum()) if not area_rows.empty else 0.0

    missing_count = int(summary["Missing Dimension"].sum()) if not summary.empty else 0
    folder_count = int(df["Folder"].nunique()) if not df.empty else 0

    col1, col2, col3, col4 = st.columns(4, gap="small")
    with col1:
        st.markdown(
            _kpi_card_html(
                "Total defects",
                f"{len(df):,}",
                "",
                f"Across {files_read:,} file{'s' if files_read != 1 else ''} \u00b7 {folder_count:,} folder{'s' if folder_count != 1 else ''}",
                _KPI_ICON_DEFECTS,
                tone="blue",
            ),
            unsafe_allow_html=True,
        )
    with col2:
        st.markdown(
            _kpi_card_html(
                "Linear repair total",
                f"{linear_repair:,.2f}",
                "LM",
                f"Board total: {linear_board:,.2f} LM",
                _KPI_ICON_LINEAR,
                tone="blue",
            ),
            unsafe_allow_html=True,
        )
    with col3:
        st.markdown(
            _kpi_card_html(
                "Area repair total",
                f"{area_repair:,.2f}",
                "SQ LM",
                f"Board total: {area_board:,.2f} SQ LM",
                _KPI_ICON_AREA,
                tone="blue",
            ),
            unsafe_allow_html=True,
        )
    with col4:
        warn_tone = "orange" if missing_count > 0 else "green"
        warn_icon = _KPI_ICON_WARN if missing_count > 0 else _KPI_ICON_OK
        warn_sub = "Needs review in Excel" if missing_count > 0 else "All dimensions complete"
        st.markdown(
            _kpi_card_html(
                "Missing dimensions",
                f"{missing_count:,}",
                "",
                warn_sub,
                warn_icon,
                tone=warn_tone,
            ),
            unsafe_allow_html=True,
        )


def _render_missing_dimensions_queue(df):
    """Best Practice #1: When any defects are missing a required board dimension,
    shows a compact expandable review queue with the exact SharePoint Folder, File,
    Tower, Floor, Drop, Material, and Defect Code so engineers can jump straight to
    that Excel workbook and fill in the missing number."""
    missing_df = _missing_dimension_rows(df)
    if missing_df.empty:
        return

    count = len(missing_df)
    label = f"Records needing review: {count:,} defect{'s' if count != 1 else ''} missing a dimension in Excel"
    with st.expander(label, expanded=False):
        st.caption(
            "Open the corresponding Excel file in SharePoint to check the embedded whiteboard photo "
            "and fill in the missing dimension. The dashboard updates automatically once saved."
        )
        cols = [c for c in REVIEW_QUEUE_COLUMNS if c in missing_df.columns]
        shown = missing_df[cols].head(DATA_TABLE_MAX_ROWS)
        table = shown.to_html(index=False, border=0, na_rep="", classes="data-table", escape=True)
        table = "".join(line.strip() for line in table.splitlines())
        st.markdown(f'<div class="data-table-wrap">{table}</div>', unsafe_allow_html=True)


# ---------------------------------------------------------------------------
# Charts
# ---------------------------------------------------------------------------

# Bars are at most this wide (pixels) when a chart has this many bars or fewer
BAR_MAX_WIDTH = 48
FEW_BARS_MAX = 8

# The value written above each bar; with more bars than MANY_BARS_MIN it gets smaller so the numbers fit
VALUE_LABEL_COLOR = "#1E293B"
VALUE_LABEL_SIZE_FEW = 12
VALUE_LABEL_SIZE_MANY = 9.5
MANY_BARS_MIN = 20


def _bar_chart(counts, column):
    """A bar chart of counts (from dashboard_data.count_by), keeping the order the table is
    already in. Every bar carries its exact value above it, so nobody has to hover to read it."""
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
    # Padding leaves a gap between bars (and keeps a lone bar from filling the chart)
    x = alt.X(f"{column}:N", sort=list(counts[column]), title=None,
              scale=alt.Scale(paddingInner=0.25, paddingOuter=0.25), axis=x_axis)
    y = alt.Y("Defects:Q", title=None, axis=alt.Axis(tickMinStep=1))

    bars = (
        alt.Chart(counts)
        # With only a few categories the bars would stretch into huge blocks, so cap their width
        .mark_bar(cornerRadiusTopLeft=3, cornerRadiusTopRight=3, **({"size": BAR_MAX_WIDTH} if len(counts) <= FEW_BARS_MAX else {}))
        .encode(
            # Blue for known values, orange for the "Unknown" bar
            color=alt.condition(
                f"datum['{column}'] == '{UNKNOWN}'",
                alt.value(ACCENT_COLOR),
                alt.value(CHART_COLOR),
            ),
            x=x,
            y=y,
            tooltip=[alt.Tooltip(f"{column}:N", title=column), alt.Tooltip("Defects:Q", title="Number of defects")],
        )
    )
    # The exact value, just above each bar (smaller when there are very many bars)
    labels = (
        alt.Chart(counts)
        .mark_text(
            align="center", baseline="bottom", dy=-4,
            fontSize=VALUE_LABEL_SIZE_FEW if len(counts) <= MANY_BARS_MIN else VALUE_LABEL_SIZE_MANY,
            fontWeight=600, color=VALUE_LABEL_COLOR,
        )
        .encode(x=x, y=y, text=alt.Text("Defects:Q", format=",d"))
    )

    chart = (
        alt.layer(bars, labels)
        # padding: breathing room between the chart and the rounded card, so nothing touches the
        # edge; the top one is larger so the label above the tallest bar is not cut off
        .properties(
            height=220, width="container", background=CHART_BACKGROUND,
            padding={"left": 4, "right": 12, "top": 22, "bottom": 4},
        )
        .configure(font=CHART_FONT)
        .configure_axis(
            labelColor=CHART_TEXT, labelFontSize=12,  # Upgraded to 12px for legibility
            titleFontSize=13,
            gridColor=CHART_GRID, domainColor="#DBE2EA", tickColor="#DBE2EA",
        )
        .configure_view(stroke=None)
    )
    try:
        st.altair_chart(chart, theme=None, use_container_width=True)  # theme=None keeps the colors above instead of the app theme's
    except TypeError:  # newer Streamlit without use_container_width: the chart already has width="container"
        st.altair_chart(chart, theme=None)


# ---------------------------------------------------------------------------
# Dimensions by defect: total dimension and total repair dimension per defect code
# ---------------------------------------------------------------------------

def _number_text(value, decimals=2, trim=False):
    """12345.5 -> '12,345.50' (or '12,345.5' with trim). '' for a missing value."""
    if value is None or value != value:
        return ""
    text = f"{value:,.{decimals}f}"
    if trim and "." in text:
        text = text.rstrip("0").rstrip(".")
    return text


def _dimension_cell(value, unit, maximum, css_class, decimals, trim):
    """A total with its unit and a thin bar under it, as long as the total is compared with the
    largest one of the same unit (different units are never compared with each other)."""
    if value is None or value != value:
        return '<span class="dim-na">&mdash;</span>'
    width = max(value / maximum * 100, 2) if maximum else 0
    return (
        f'<div class="dim-value"><b>{_number_text(value, decimals, trim)}</b>'
        f'<span class="dim-unit">{html.escape(unit)}</span></div>'
        f'<div class="dim-track"><div class="dim-bar {css_class}" style="width:{width:.1f}%"></div></div>'
    )


def _largest_by_unit(records, value_key, unit_key):
    largest = {}
    for record in records:
        value = record[value_key]
        if value is not None and value == value:
            largest[record[unit_key]] = max(largest.get(record[unit_key], 0), value)
    return largest


def _totals_lines(records, value_key, unit_key, decimals, trim):
    """Returns a list of ('2,645.34 LM', '216.10 SQ LM') strings, one per unit, so each unit
    can be displayed on its own separate line."""
    totals = {}
    for record in records:
        value = record[value_key]
        unit = record[unit_key]
        if value is not None and value == value and unit:
            totals[unit] = totals.get(unit, 0) + value
    return [f"{_number_text(total, decimals, trim)} {unit}".strip() for unit, total in totals.items()]


def _summary_to_excel_bytes(summary):
    """Builds a clean, formatted Excel (.xlsx) file in memory from the Dimensions by defect
    summary table, including the grand total rows for LM and SQ LM at the bottom."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Dimensions by Defect"

    headers = [
        "Defect Code", "Defect Type", "Defects",
        "Total Dimension", "Unit",
        "Total Repair Dimension", "Repair Unit",
        "Missing Dimensions",
    ]
    ws.append(headers)

    header_fill = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid")
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    for col_idx in range(1, len(headers) + 1):
        cell = ws.cell(row=1, column=col_idx)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 24

    records = summary.to_dict("records")
    for record in records:
        tot_dim = record["Total Dimension"]
        rep_dim = record["Total Repair Dimension"]
        ws.append([
            record["Defect Code"],
            record["Defect Type"],
            int(record["Defects"]),
            round(float(tot_dim), 2) if tot_dim is not None and tot_dim == tot_dim else "",
            record["Board Unit"] or "",
            round(float(rep_dim), 2) if rep_dim is not None and rep_dim == rep_dim else "",
            record["Repair Unit"] or "",
            int(record["Missing Dimension"]),
        ])
        row_num = ws.max_row
        ws.cell(row=row_num, column=4).number_format = "#,##0.00"
        ws.cell(row=row_num, column=6).number_format = "#,##0.00"

    # Append Grand Total rows per unit (LM and SQ LM)
    ws.append([])
    total_fill = PatternFill(start_color="F1F5F9", end_color="F1F5F9", fill_type="solid")
    total_font = Font(name="Calibri", size=11, bold=True, color="0F172A")

    for unit in ("LM", "SQ LM"):
        unit_rows = [r for r in records if r["Repair Unit"] == unit or r["Board Unit"] == unit]
        if not unit_rows:
            continue
        board_sum = sum(
            float(r["Total Dimension"]) for r in unit_rows
            if r["Board Unit"] == unit and r["Total Dimension"] is not None and r["Total Dimension"] == r["Total Dimension"]
        )
        repair_sum = sum(
            float(r["Total Repair Dimension"]) for r in unit_rows
            if r["Repair Unit"] == unit and r["Total Repair Dimension"] is not None and r["Total Repair Dimension"] == r["Total Repair Dimension"]
        )
        ws.append([
            f"TOTAL ({unit})",
            "",
            sum(int(r["Defects"]) for r in unit_rows),
            round(board_sum, 2),
            unit,
            round(repair_sum, 2),
            unit,
            sum(int(r["Missing Dimension"]) for r in unit_rows),
        ])
        row_num = ws.max_row
        ws.cell(row=row_num, column=4).number_format = "#,##0.00"
        ws.cell(row=row_num, column=6).number_format = "#,##0.00"
        for col_idx in range(1, len(headers) + 1):
            c = ws.cell(row=row_num, column=col_idx)
            c.fill = total_fill
            c.font = total_font

    widths = {"A": 16, "B": 28, "C": 12, "D": 18, "E": 10, "F": 22, "G": 12, "H": 18}
    for letter, width in widths.items():
        ws.column_dimensions[letter].width = width

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def _dimension_table_html(records):
    """The per-defect table as plain HTML (see .dim-table in style.css), written without
    indentation so Markdown leaves it alone."""
    board_largest = _largest_by_unit(records, "Total Dimension", "Board Unit")
    repair_largest = _largest_by_unit(records, "Total Repair Dimension", "Repair Unit")
    body = ""
    for record in records:
        body += (
            "<tr>"
            f'<td class="dim-code">{html.escape(str(record["Defect Code"]))}</td>'
            f'<td>{html.escape(str(record["Defect Type"]))}</td>'
            f'<td class="dim-count">{record["Defects"]:,}</td>'
            '<td class="dim-cell">'
            + _dimension_cell(record["Total Dimension"], record["Board Unit"],
                              board_largest.get(record["Board Unit"]), "dim-bar-board", 2, False)
            + '</td><td class="dim-cell">'
            + _dimension_cell(record["Total Repair Dimension"], record["Repair Unit"],
                              repair_largest.get(record["Repair Unit"]), "dim-bar-repair", 2, False)
            + "</td></tr>"
        )
    return (
        '<div class="dim-table-wrap"><table class="dim-table"><thead><tr>'
        '<th>Defect</th><th>Defect type</th><th class="dim-count">Defects</th>'
        '<th>Total dimension</th><th>Total repair dimension</th>'
        f"</tr></thead><tbody>{body}</tbody></table></div>"
    )


def _dimension_section(summary, scope="All"):
    """A card with the total dimension and the total repair dimension of every defect code,
    plus a Download button to export the table as an Excel (.xlsx) workbook."""
    records = summary.to_dict("records")
    repair_lines = _totals_lines(records, "Total Repair Dimension", "Repair Unit", 2, False)
    repair_html = "".join(
        f'<div class="chart-total-line">Repair total: {html.escape(line)}</div>'
        for line in repair_lines
    )

    with st.container(border=True, key="chart_dimensions"):
        st.markdown(
            '<div class="chart-title"><div>'
            '<div class="chart-name">Dimensions by defect</div>'
            '<div class="chart-sub">Total dimension and total repair dimension for each defect code (in LM and SQ LM)</div>'
            f'</div><div class="chart-totals-stack">{repair_html}</div></div>',
            unsafe_allow_html=True,
        )
        if not records:
            st.caption("Nothing to show yet.")
            return
        st.markdown(_dimension_table_html(records), unsafe_allow_html=True)

        notes = []
        missing = int(summary["Missing Dimension"].sum())
        if missing:
            notes.append(
                f"{missing:,} defect(s) were written without a dimension, so they add nothing to the totals."
            )
        notes.append(
            "Board dimensions are converted from CM and CM² to LM and SQ LM. "
            "Repair dimensions are read from the Excel files (or calculated with the same formulas)."
        )

        # Bottom row inside the card: caption on the left, Download Excel button on the right
        foot_left, foot_right = st.columns([0.75, 0.25], gap="small")
        with foot_left:
            st.caption(" ".join(notes))
        with foot_right:
            safe_scope = "".join(ch if ch.isalnum() or ch in ("-", "_") else "_" for ch in str(scope)).strip("_") or "Summary"
            st.download_button(
                label="Download summary (.xlsx)",
                data=_summary_to_excel_bytes(summary),
                file_name=f"Dimensions_by_Defect_{safe_scope}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
                key="download_dimensions_summary",
            )


def _chart_notes(df, hide_unknown):
    """ONE short line drawn above the chart filters area or as a general note."""
    columns = ("Tower", "Floor", "Material")
    if hide_unknown:
        parts = []
        for column in columns:
            counts = count_by(df, column)
            hidden = int(counts.loc[counts[column] == UNKNOWN, "Defects"].sum())
            if hidden:
                parts.append(f"{hidden:,} defect(s) with an unknown {column.lower()} are hidden.")
        return " ".join(parts)
    return ""


def _chart_section(df, column, title, hide_unknown):
    """A white card with a title, a one-line subtitle, the total, and the bar chart of
    the defects per value of `column`. With hide_unknown, the "Unknown" bar is left out,
    so the real distribution isn't flattened by it."""
    counts = count_by(df, column)
    if hide_unknown:
        counts = counts[counts[column] != UNKNOWN]
    
    # Check if this chart has an unknown bar
    has_unknown = not hide_unknown and not counts[counts[column] == UNKNOWN].empty

    total = int(counts["Defects"].sum()) if not counts.empty else 0
    with st.container(border=True, key=f"chart_{column.lower()}"):
        # Include the in-chart legend if 'Unknown' is present
        legend_html = '<span style="display:inline-block;width:8px;height:8px;border-radius:2px;background:#f97316;margin-right:5px;"></span>Unknown classification' if has_unknown else ''
        
        st.markdown(
            '<div class="chart-title"><div>'
            f'<div class="chart-name">{title}</div>'
            f'<div class="chart-sub">{CHART_SUBTITLES.get(column, "")}</div>'
            f'</div>'
            f'<div style="text-align:right;">'
            f'<span class="chart-total">{total:,} total</span>'
            f'<div style="color:#64748b;font-size:11.5px;font-weight:500;margin-top:2px;">{legend_html}</div>'
            f'</div></div>',
            unsafe_allow_html=True,
        )
        if counts.empty:
            st.caption(f"Nothing to show: every defect here has an unknown {column.lower()}.")
            return
        _bar_chart(counts, column)


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
    """The dashboard's controls (the left panel). Not a
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

    # --- Render Active Filter Chips ---
    active_filters = []
    
    selected_towers = st.session_state.get(_filter_key("tower", scope)) or []
    if selected_towers:
        active_filters.append(f"Tower: {', '.join(selected_towers)}")
        
    selected_floors = st.session_state.get(_filter_key("floor", scope)) or []
    if selected_floors:
        active_filters.append(f"Floor: {', '.join(selected_floors)}")
        
    selected_materials = st.session_state.get(_filter_key("material", scope)) or []
    if selected_materials:
        active_filters.append(f"Material: {', '.join(selected_materials)}")
        
    folders = st.session_state.get(_filter_key("folder", scope)) or []
    if folders:
        active_filters.append(f"Folder: {', '.join(folders)}")

    if hide_unknown:
        active_filters.append("Exclude unknown")

    if active_filters:
        chips_html = "".join([f'<span class="filter-chip">{html.escape(f)} &times;</span>' for f in active_filters])
        st.markdown(
            f'<div class="filter-chips-container">{chips_html}</div>', 
            unsafe_allow_html=True
        )

        # Clear button (requires a rerun to actually clear session state)
        if st.button("Clear all filters", key="clear_all_filters", help="Reset all applied filters"):
            st.session_state[_filter_key("tower", scope)] = []
            st.session_state[_filter_key("floor", scope)] = []
            st.session_state[_filter_key("material", scope)] = []
            st.session_state[_filter_key("folder", scope)] = []
            st.session_state[HIDE_UNKNOWN_KEY] = False
            st.rerun()

    # --- Apply the sidebar filters to data ---
    if folders:
        df = df[df["Folder"].isin(folders)]
    df = apply_filters(
        df,
        selected_towers,
        selected_floors,
        selected_materials,
    )

    summary = dimension_by_defect(df)

    # --- 4 Engineering KPI Cards (Option B) ---
    _render_kpi_cards(df, summary, files_read)

    if df.empty:
        st.info("No defects match these filters.")
        return

    # --- Best Practice #1: Review Queue for Defects Missing a Required Dimension ---
    _render_missing_dimensions_queue(df)

    # --- Charts: Tower beside Material, then Floor (the widest) across the page ---
    with st.container(key="charts_grid"):
        tower_col, material_col = st.columns([1, 1], gap="small")  # equal halves; style.css sets the 14px gap
        with tower_col:
            _chart_section(df, "Tower", "Defects by tower", hide_unknown)
        with material_col:
            _chart_section(df, "Material", "Defects by material", hide_unknown)
        _chart_section(df, "Floor", "Defects by floor", hide_unknown)
        _dimension_section(summary, scope=scope)

    # Note about hidden unknown defects
    note = _chart_notes(df, hide_unknown)
    if note:
        st.markdown(f'<div class="chart-legend">{html.escape(note)}</div>', unsafe_allow_html=True)

    with st.expander("View the data behind the charts"):
        st.markdown(_data_table_html(df), unsafe_allow_html=True)
        if len(df) > DATA_TABLE_MAX_ROWS:
            st.caption(f"Showing the first {DATA_TABLE_MAX_ROWS:,} of {len(df):,} defects.")


def render_dashboard_view():
    """The whole Dashboard view: the controls panel and the charts, side by side in
    one connected card under the header. The CSS in style.css draws the card
    around the two columns (it finds them through the two container keys)."""
    col_controls, col_main = st.columns([0.22, 0.78], gap="small")
    with col_controls:
        with st.container(key="dashboard_sidebar"):
            render_dashboard_controls()
    with col_main:
        with st.container(key="dashboard_main"):
            # "analytics" stays its own container: the KPI card icons in style.css are keyed to it
            with st.container(key="analytics"):
                render_analytics_dashboard()