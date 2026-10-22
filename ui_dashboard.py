"""The "Analytics Dashboard" tab, structured around the Engineering Overview +
Repair Quantity Summary + Location Analysis layout:

  1. ENGINEERING OVERVIEW ("Facade defect dashboard")
     - 4 KPI cards: Total defects, Linear repair total (LM), Area repair total (SQM), Missing dimensions
     - Collapsible warm-amber review banner ("X records need review") when dimensions are missing

  2. REPAIR QUANTITY SUMMARY ("Dimensions by defect")
     - Header with "Download summary XLSX" button on the top right
     - Table ordered with all Linear Meter (LM) rows first, then Square Meter (SQM) rows,
       with pill badges for Defect Code, Description, Count, Board Dimensions,
       Linear Repair (LM), Area Repair (SQM), Data Health, and a Portfolio Total footer row

  3. LOCATION ANALYSIS ("Where are the defects located?")
     - Defects by material beside Defects by tower, then Defects by floor across the width

  4. RAW AUDIT LOG
     - Expandable drawer at the bottom showing the underlying defect records
"""
import io
import html

import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill
import streamlit as st

import sharepoint_client
import ui_components
from dashboard_data import (
    rows_to_dataframe, apply_filters, count_by, dimension_by_defect,
    natural_key, UNKNOWN,
)
from dashboard_loader import get_file_rows, is_loaded
from defect_rules import BOARD_UNITS
from settings import ENABLE_SHAREPOINT_SYNC, MAIN_FOLDER_LABEL, DASHBOARD_REFRESH_SECONDS

# Dropdown choice that reads the whole landing folder: its own files plus every subfolder's
ALL_FILES_LABEL = "Whiteboard Excel Files"

# Strict canonical defect options per substrate for the sidebar sub-filter
SUBSTRATE_DEFECT_OPTIONS = {
    "Concrete": ["C-", "C+", "BH", "US"],
    "Sealant": [
        "DS CC", "DS CF", "DS GG", "DS FG", "DS FF",
        "MS CC", "MS CF", "MS GG", "MS FG", "MS FF",
    ],
    "Paint": ["DP", "FP", "BP", "NP"],
    "Gasket": ["DG"],
    "Broken Glass": ["BG"],
    "With Film": ["GLASS-FRAME"],
}

# Chart look: Vonotec-blue bars on a card, dynamically adapted for Dark Mode.
CHART_COLOR = "#1E40AF"
ACCENT_COLOR = "#EA580C"       # orange: marks the "Unknown" bars, so missing data stands out
CHART_BACKGROUND = "rgba(0,0,0,0)"  # transparent, so the card behind the chart shows through
CHART_TEXT = "#64748B"         # axis labels
CHART_GRID = "#EDF0F4"         # soft gridlines behind the bars
CHART_FONT = "DM Sans"

# With more bars than this, the x-axis labels are tilted so they don't run into each other
STRAIGHT_LABELS_MAX_BARS = 6

CHART_SUBTITLES = {
    "Material": "Top facade classifications",
    "Tower": "Distribution across project towers",
    "Floor": "Concentration across active levels",
    "Cause": "Why each defect most likely happened",
}

# Column order of the raw audit log table
TABLE_COLUMNS = [
    "Tower", "Floor", "Material", "Defect Code", "Defect Type", "Dimension", "Unit",
    "Repair Dimension", "Repair Unit", "Date", "Submitter", "Elevation", "Drop", "Folder", "File",
]

# Columns shown in the "Records needing review (missing dimensions)" queue table
REVIEW_QUEUE_COLUMNS = [
    "Folder", "File", "Tower", "Floor", "Drop", "Elevation",
    "Material", "Defect Code", "Defect Type", "Date", "Submitter",
]

DATA_TABLE_MAX_ROWS = 1000

SCOPE_KEY = "analytics_scope"
HIDE_UNKNOWN_KEY = "analytics_hide_unknown"
LIST_CACHE_SECONDS = max(DASHBOARD_REFRESH_SECONDS - 1, 1)

# Uniform Lucide 24x24 SVG icons (stroke-width="2", round caps & joins) for KPI cards, banners & drawers
_KPI_ICON_DEFECTS = (
    '<svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" '
    'stroke-width="2" stroke-linecap="round" stroke-linejoin="round">'
    '<path d="M3 3v16a2 2 0 0 0 2 2h16"/><path d="M18 17V9"/><path d="M13 17V5"/><path d="M8 17v-3"/></svg>'
)
_KPI_ICON_LINEAR = (
    '<svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" '
    'stroke-width="2" stroke-linecap="round" stroke-linejoin="round">'
    '<path d="M21.3 15.3a2.4 2.4 0 0 1 0 3.4l-2.6 2.6a2.4 2.4 0 0 1-3.4 0L2.7 8.7a2.41 2.41 0 0 1 0-3.4l2.6-2.6a2.41 2.41 0 0 1 3.4 0Z"/>'
    '<path d="m14.5 12.5 2-2"/><path d="m11.5 9.5 2-2"/><path d="m8.5 6.5 2-2"/><path d="m17.5 15.5 2-2"/></svg>'
)
_KPI_ICON_AREA = (
    '<svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" '
    'stroke-width="2" stroke-linecap="round" stroke-linejoin="round">'
    '<rect width="7" height="7" x="3" y="3" rx="1"/><rect width="7" height="7" x="14" y="3" rx="1"/>'
    '<rect width="7" height="7" x="14" y="14" rx="1"/><rect width="7" height="7" x="3" y="14" rx="1"/></svg>'
)
_KPI_ICON_WARN = (
    '<svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" '
    'stroke-width="2" stroke-linecap="round" stroke-linejoin="round">'
    '<path d="m21.73 18-8-14a2 2 0 0 0-3.48 0l-8 14A2 2 0 0 0 4 21h16a2 2 0 0 0 1.73-3"/>'
    '<path d="M12 9v4"/><path d="M12 17h.01"/></svg>'
)
_KPI_ICON_OK = (
    '<svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" '
    'stroke-width="2" stroke-linecap="round" stroke-linejoin="round">'
    '<circle cx="12" cy="12" r="10"/><path d="m9 12 2 2 4-4"/></svg>'
)
_ICON_AUDIT_TABLE = (
    '<svg width="19" height="19" viewBox="0 0 24 24" fill="none" stroke="currentColor" '
    'stroke-width="2" stroke-linecap="round" stroke-linejoin="round">'
    '<path d="M9 3H5a2 2 0 0 0-2 2v4m6-6h10a2 2 0 0 1 2 2v4M9 3v18m0 0h10a2 2 0 0 0 2-2V9M9 21H5a2 2 0 0 1-2-2V9m0 0h18"/></svg>'
)
_ICON_CHECK_SM = (
    '<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" '
    'stroke-width="2" stroke-linecap="round" stroke-linejoin="round">'
    '<circle cx="12" cy="12" r="10"/><path d="m9 12 2 2 4-4"/></svg>'
)
_ICON_WARN_SM = (
    '<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" '
    'stroke-width="2" stroke-linecap="round" stroke-linejoin="round">'
    '<path d="m21.73 18-8-14a2 2 0 0 0-3.48 0l-8 14A2 2 0 0 0 4 21h16a2 2 0 0 0 1.73-3"/>'
    '<path d="M12 9v4"/><path d="M12 17h.01"/></svg>'
)
_ICON_CHEVRON_DOWN = (
    '<svg class="boq-chevron" width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" '
    'stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="m6 9 6 6 6-6"/></svg>'
)


# ---------------------------------------------------------------------------
# Cached SharePoint reads (supports nested output folders)
# ---------------------------------------------------------------------------

@st.cache_data(ttl=LIST_CACHE_SECONDS, show_spinner=False)
def _cached_subfolders():
    """Returns all output folders and nested subfolders inside the landing folder."""
    return sharepoint_client.list_all_output_folders()


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
    """The folders to read, where "" is the landing folder itself.
    Selecting a parent folder (e.g. 'Cerca Nuveo') automatically includes any
    nested subfolders inside it (e.g. 'Cerca Nuveo/Tower 1')."""
    if scope == ALL_FILES_LABEL:
        return [""] + list(subfolders)
    if scope == MAIN_FOLDER_LABEL:
        return [""]
    prefix = f"{scope}/"
    children = [f for f in subfolders if f.startswith(prefix)]
    return [scope] + children


def _load_scope(scope, subfolders):
    """Returns (rows, files_read, skipped), where skipped is a list of
    (filename, reason) for files that could not be used."""
    files = []
    for folder in _folders_for_scope(scope, subfolders):
        files += _cached_file_list(folder)

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
            link = item.get("web_url", "")
            rows += [{**row, "File Link": link} for row in file_rows]
            files_read += 1
        if progress is not None:
            progress.progress((index + 1) / len(files), text=f"Reading Excel files... {index + 1} of {len(files)}")
    if progress is not None:
        progress.empty()
    return rows, files_read, skipped


# ---------------------------------------------------------------------------
# Section 1: KPI Cards & Warm-Amber Review Banner
# ---------------------------------------------------------------------------

SEALANT_LOCATION_NAMES = {
    "CC": "Concrete to Concrete",
    "CF": "Concrete to Frame",
    "GG": "Glass to Glass",
    "FG": "Frame to Glass",
    "FF": "Frame to Frame",
}

_DASHBOARD_DESC_OVERRIDES = {
    "BH": "Bughole",
    "US": "Uneven Surface",
}


def _defect_description(code, defect_type):
    """The Description of a summary row: uses dashboard-specific overrides for BH and US,
    plus the surfaces for a sealant code."""
    clean_code = str(code or "").strip()
    if clean_code in _DASHBOARD_DESC_OVERRIDES:
        return _DASHBOARD_DESC_OVERRIDES[clean_code]
    parts = clean_code.split()
    if len(parts) > 1 and parts[1] in SEALANT_LOCATION_NAMES:
        return f"{defect_type} ({SEALANT_LOCATION_NAMES[parts[1]]})"
    return str(defect_type)


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


def _kpi_card_html(label, value_text, unit_text, sub_html, icon_svg, tone="blue"):
    """Renders one enterprise KPI card matching the Engineering Overview design."""
    unit_html = f'<span class="kpi-unit">{html.escape(unit_text)}</span>' if unit_text else ""
    return (
        f'<div class="kpi-card kpi-{html.escape(tone)}">'
        f'<span class="kpi-icon kpi-icon-{html.escape(tone)}">{icon_svg}</span>'
        '<div class="kpi-body">'
        f'<div class="kpi-label">{html.escape(label)}</div>'
        f'<div class="kpi-value"><span>{html.escape(value_text)}</span>{unit_html}</div>'
        f'<div class="kpi-sub">{sub_html}</div>'
        '</div></div>'
    )


def _render_kpi_cards(df, summary, files_read):
    """Draws the 4 Engineering KPI cards across the top of the dashboard:
      1. Total defects
      2. Linear repair total (LM)
      3. Area repair total (SQM)
      4. Missing dimensions"""
    linear_rows = summary[summary["Repair Unit"] == "LM"] if not summary.empty else summary
    area_rows = summary[summary["Repair Unit"] == "SQM"] if not summary.empty else summary

    linear_repair = float(linear_rows["Total Repair Dimension"].dropna().sum()) if not linear_rows.empty else 0.0
    linear_board = float(linear_rows["Total Dimension"].dropna().sum()) if not linear_rows.empty else 0.0
    linear_measured = int((linear_rows["Defects"] - linear_rows["Missing Dimension"]).clip(lower=0).sum()) if not linear_rows.empty else 0

    area_repair = float(area_rows["Total Repair Dimension"].dropna().sum()) if not area_rows.empty else 0.0
    area_board = float(area_rows["Total Dimension"].dropna().sum()) if not area_rows.empty else 0.0
    area_measured = int((area_rows["Defects"] - area_rows["Missing Dimension"]).clip(lower=0).sum()) if not area_rows.empty else 0

    missing_count = int(summary["Missing Dimension"].sum()) if not summary.empty else 0
    missing_cats = int((summary["Missing Dimension"] > 0).sum()) if not summary.empty else 0
    folder_count = int(df["Folder"].nunique()) if not df.empty else 0

    col1, col2, col3, col4 = st.columns(4, gap="small")
    with col1:
        sub1 = (
            f'<b class="kpi-sub-green">{files_read:,} file{"s" if files_read != 1 else ""}</b> '
            f'across {folder_count:,} folder{"s" if folder_count != 1 else ""}'
        )
        st.markdown(
            _kpi_card_html("Total defects", f"{len(df):,}", "", sub1, _KPI_ICON_DEFECTS, tone="blue"),
            unsafe_allow_html=True,
        )
    with col2:
        sub2 = f"Board: {linear_board:,.2f} LM \u00b7 {linear_measured:,} measured"
        st.markdown(
            _kpi_card_html("Linear repair total", f"{linear_repair:,.2f}", "LM", html.escape(sub2), _KPI_ICON_LINEAR, tone="teal"),
            unsafe_allow_html=True,
        )
    with col3:
        sub3 = f"Board: {area_board:,.2f} SQM \u00b7 {area_measured:,} measured"
        st.markdown(
            _kpi_card_html("Area repair total", f"{area_repair:,.2f}", "SQM", html.escape(sub3), _KPI_ICON_AREA, tone="green"),
            unsafe_allow_html=True,
        )
    with col4:
        warn_tone = "orange" if missing_count > 0 else "green"
        warn_icon = _KPI_ICON_WARN if missing_count > 0 else _KPI_ICON_OK
        warn_sub = (
            f"Across {missing_cats:,} defect categor{'ies' if missing_cats != 1 else 'y'}"
            if missing_count > 0 else "All dimensions complete"
        )
        st.markdown(
            _kpi_card_html("Missing dimensions", f"{missing_count:,}", "", html.escape(warn_sub), warn_icon, tone=warn_tone),
            unsafe_allow_html=True,
        )


def _data_table_html(shown, float_format=None, link_files=False):
    """The HTML for a plain data table. Every cell is escaped. With link_files=True the File column
    becomes a link that opens that Excel file in SharePoint (only people with access to the site can open it)."""
    table_df = shown.copy()
    links = table_df.pop("File Link") if "File Link" in table_df.columns else None
    for column in table_df.columns:
        table_df[column] = table_df[column].map(lambda value: html.escape(value) if isinstance(value, str) else value)
    if link_files and links is not None and "File" in table_df.columns:
        def _file_cell(name, link):
            link = link if isinstance(link, str) else ""
            if not link.startswith("https://"):
                return name
            return f'<a class="file-link" href="{html.escape(link, quote=True)}" target="_blank" rel="noopener noreferrer">{name}</a>'
        table_df["File"] = [_file_cell(name, link) for name, link in zip(table_df["File"], links)]
    kwargs = {"float_format": float_format} if float_format else {}
    table = table_df.to_html(index=False, border=0, na_rep="", classes="data-table", escape=False, **kwargs)
    return "".join(line.strip() for line in table.splitlines())


def _render_missing_dimensions_queue(df):
    """Draws the warm-amber collapsible review banner ('X records need review') right below
    the KPI cards when any defects are missing a required dimension in Excel."""
    missing_df = _missing_dimension_rows(df)
    if missing_df.empty:
        return

    count = len(missing_df)
    cols = [c for c in REVIEW_QUEUE_COLUMNS if c in missing_df.columns]
    shown = missing_df[cols + (["File Link"] if "File Link" in missing_df.columns else [])].head(DATA_TABLE_MAX_ROWS)
    table = _data_table_html(shown, link_files=True)

    banner_html = (
        '<details class="review-banner">'
        '<summary class="review-banner-head">'
        '<div class="review-banner-left">'
        f'<span class="review-banner-icon">{_KPI_ICON_WARN}</span>'
        '<div>'
        f'<div class="review-banner-title">{count:,} record{"s" if count != 1 else ""} need{"" if count != 1 else "s"} review</div>'
        '<div class="review-banner-sub">Dimensions are missing in the Excel Master Log</div>'
        '</div></div>'
        f'<div class="review-banner-cta">Review records {_ICON_CHEVRON_DOWN}</div>'
        '</summary>'
        '<div class="review-banner-body">'
        '<p class="review-banner-note">Open the corresponding Excel file in SharePoint to check the embedded whiteboard photo and fill in the missing dimension. The dashboard updates automatically once saved.</p>'
        f'<div class="data-table-wrap">{table}</div>'
        '</div></details>'
    )
    st.markdown(banner_html, unsafe_allow_html=True)


# ---------------------------------------------------------------------------
# Section 2: Repair Quantity Summary ("Dimensions by defect" table + Excel export)
# ---------------------------------------------------------------------------

def _number_text(value, decimals=2, trim=False):
    """12345.5 -> '12,345.50' (or '12,345.5' with trim). '' for a missing value."""
    if value is None or value != value:
        return ""
    text = f"{value:,.{decimals}f}"
    if trim and "." in text:
        text = text.rstrip("0").rstrip(".")
    return text


def _summary_to_excel_bytes(summary):
    """Builds a clean, formatted Excel (.xlsx) file in memory from the Dimensions by defect
    summary table, matching the Repair Quantity Summary columns and Portfolio total footer."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Repair Quantity Summary"

    headers = [
        "Defect Code", "Description", "Count",
        "Board Dimensions", "Board Unit",
        "Linear Repair (LM)", "Area Repair (SQM)",
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
    total_defects = 0
    total_linear_repair = 0.0
    total_area_repair = 0.0
    total_missing = 0

    group_fill = PatternFill(start_color="E8EEF8", end_color="E8EEF8", fill_type="solid")
    group_font = Font(name="Calibri", size=11, bold=True, color="1E3A8A")
    material_counts = {}
    for record in records:
        material_counts[record["Material"]] = material_counts.get(record["Material"], 0) + int(record["Defects"])

    current_material = None
    for record in records:
        if record["Material"] != current_material:
            current_material = record["Material"]
            ws.append([str(current_material), "", material_counts[current_material], "", "", "", "", ""])
            for col_idx in range(1, len(headers) + 1):
                group_cell = ws.cell(row=ws.max_row, column=col_idx)
                group_cell.fill = group_fill
                group_cell.font = group_font

        tot_dim = record["Total Dimension"]
        rep_dim = record["Total Repair Dimension"]
        rep_unit = record["Repair Unit"] or ""
        cnt = int(record["Defects"])
        miss = int(record["Missing Dimension"])

        has_rep = rep_dim is not None and rep_dim == rep_dim
        lin_val = round(float(rep_dim), 2) if has_rep and rep_unit == "LM" else ""
        area_val = round(float(rep_dim), 2) if has_rep and rep_unit == "SQM" else ""

        total_defects += cnt
        if isinstance(lin_val, float):
            total_linear_repair += lin_val
        if isinstance(area_val, float):
            total_area_repair += area_val
        total_missing += miss

        ws.append([
            record["Defect Code"],
            _defect_description(record["Defect Code"], record["Defect Type"]),
            cnt,
            round(float(tot_dim), 2) if tot_dim is not None and tot_dim == tot_dim else "",
            record["Board Unit"] or "",
            lin_val,
            area_val,
            miss,
        ])
        row_num = ws.max_row
        ws.cell(row=row_num, column=4).number_format = "#,##0.00"
        ws.cell(row=row_num, column=6).number_format = "#,##0.00"
        ws.cell(row=row_num, column=7).number_format = "#,##0.00"

    # Portfolio Total Footer Row
    total_fill = PatternFill(start_color="F1F5F9", end_color="F1F5F9", fill_type="solid")
    total_font = Font(name="Calibri", size=11, bold=True, color="0F172A")
    ws.append([
        "Portfolio total",
        "",
        total_defects,
        "",
        "",
        round(total_linear_repair, 2),
        round(total_area_repair, 2),
        total_missing,
    ])
    row_num = ws.max_row
    ws.cell(row=row_num, column=6).number_format = "#,##0.00"
    ws.cell(row=row_num, column=7).number_format = "#,##0.00"
    for col_idx in range(1, len(headers) + 1):
        c = ws.cell(row=row_num, column=col_idx)
        c.fill = total_fill
        c.font = total_font

    widths = {"A": 16, "B": 30, "C": 12, "D": 18, "E": 12, "F": 20, "G": 22, "H": 18}
    for letter, width in widths.items():
        ws.column_dimensions[letter].width = width

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def _data_health_html(record):
    """Renders the DATA HEALTH cell for one defect row."""
    code = str(record["Defect Code"])
    base = _base_defect_code(code)
    missing = int(record["Missing Dimension"] or 0)
    if not BOARD_UNITS.get(base, ""):
        return '<span class="boq-health boq-health-na">No dimension</span>'
    if missing > 0:
        return f'<span class="boq-health boq-health-warn">{_ICON_WARN_SM}{missing:,} missing</span>'
    return f'<span class="boq-health boq-health-ok">{_ICON_CHECK_SM}Complete</span>'


def _dimension_table_html(records):
    """Builds the Repair Quantity Summary HTML table."""
    body = ""
    total_defects = 0
    total_linear = 0.0
    total_area = 0.0
    total_missing = 0

    material_counts = {}
    for record in records:
        material_counts[record["Material"]] = material_counts.get(record["Material"], 0) + int(record["Defects"] or 0)

    current_material = None
    for record in records:
        if record["Material"] != current_material:
            current_material = record["Material"]
            group_count = material_counts[current_material]
            body += (
                '<tr class="boq-group"><td colspan="7">'
                f'{html.escape(str(current_material))} '
                f'<span class="boq-group-meta">&middot; {group_count:,} defect{"s" if group_count != 1 else ""}</span>'
                "</td></tr>"
            )

        code = str(record["Defect Code"])
        desc = _defect_description(code, record["Defect Type"])
        cnt = int(record["Defects"] or 0)
        tot_dim = record["Total Dimension"]
        board_unit = str(record["Board Unit"] or "")
        rep_dim = record["Total Repair Dimension"]
        rep_unit = str(record["Repair Unit"] or "")
        miss = int(record["Missing Dimension"] or 0)

        total_defects += cnt
        total_missing += miss

        if tot_dim is not None and tot_dim == tot_dim and board_unit:
            board_html = f'{_number_text(tot_dim, 2, False)} <span class="boq-unit-inline">{html.escape(board_unit)}</span>'
        else:
            board_html = '<span class="dim-na">&mdash;</span>'

        has_rep = rep_dim is not None and rep_dim == rep_dim
        if has_rep and rep_unit == "LM":
            total_linear += float(rep_dim)
            linear_html = f'<span class="boq-val-repair">{_number_text(rep_dim, 2, False)}</span>'
        else:
            linear_html = '<span class="dim-na">&mdash;</span>'

        if has_rep and rep_unit == "SQM":
            total_area += float(rep_dim)
            area_html = f'<span class="boq-val-repair">{_number_text(rep_dim, 2, False)}</span>'
        else:
            area_html = '<span class="dim-na">&mdash;</span>'

        body += (
            "<tr>"
            f'<td><span class="boq-code-pill">{html.escape(code)}</span></td>'
            f'<td class="boq-desc">{html.escape(desc)}</td>'
            f'<td class="boq-num">{cnt:,}</td>'
            f'<td class="boq-board">{board_html}</td>'
            f'<td class="boq-repair">{linear_html}</td>'
            f'<td class="boq-repair">{area_html}</td>'
            f'<td>{_data_health_html(record)}</td>'
            "</tr>"
        )

    footer_health = (
        f'<span class="boq-health boq-health-warn">{total_missing:,} missing</span>'
        if total_missing > 0
        else f'<span class="boq-health boq-health-ok">{_ICON_CHECK_SM}All complete</span>'
    )
    tfoot = (
        "<tfoot><tr>"
        '<td colspan="2" class="boq-foot-label">Portfolio total</td>'
        f'<td class="boq-num">{total_defects:,}</td>'
        '<td class="boq-board"><span class="dim-na">&mdash;</span></td>'
        f'<td class="boq-foot-val">{_number_text(total_linear, 2, False)} <span class="boq-unit-inline">LM</span></td>'
        f'<td class="boq-foot-val">{_number_text(total_area, 2, False)} <span class="boq-unit-inline">SQM</span></td>'
        f"<td>{footer_health}</td>"
        "</tr></tfoot>"
    )

    return (
        '<div class="dim-table-wrap"><table class="dim-table"><thead><tr>'
        '<th>DEFECT</th>'
        '<th>DESCRIPTION</th>'
        '<th class="boq-num">COUNT</th>'
        '<th>BOARD DIMENSIONS <span class="th-unit">LM / SQM</span></th>'
        '<th>LINEAR REPAIR <span class="th-unit">LM</span></th>'
        '<th>AREA REPAIR <span class="th-unit">SQM</span></th>'
        '<th>DATA HEALTH</th>'
        f"</tr></thead><tbody>{body}</tbody>{tfoot}</table></div>"
    )


def _dimension_section(summary, scope="All"):
    """Draws the Repair Quantity Summary ('Dimensions by defect') card with the Download summary
    button in the top-right header and the full engineering table below."""
    records = summary.to_dict("records")
    safe_scope = "".join(ch if ch.isalnum() or ch in ("-", "_") else "_" for ch in str(scope)).strip("_") or "Summary"

    with st.container(border=True, key="chart_dimensions"):
        head_left, head_right = st.columns([0.72, 0.28], gap="small")
        with head_left:
            st.markdown(
                '<div class="boq-header">'
                '<p class="eyebrow">Repair quantity summary</p>'
                '<div class="chart-name">Dimensions by defect</div>'
                '<div class="chart-sub">Repair quantities grouped by material, then by defect code</div>'
                '</div>',
                unsafe_allow_html=True,
            )
        with head_right:
            if records:
                st.download_button(
                    label="Download summary  XLSX",
                    data=_summary_to_excel_bytes(summary),
                    file_name=f"Repair_Quantity_Summary_{safe_scope}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary",
                    key="download_dimensions_summary",
                )

        if not records:
            st.caption("Nothing to show yet.")
            return
        st.markdown(_dimension_table_html(records), unsafe_allow_html=True)


# ---------------------------------------------------------------------------
# Section 3: Location Analysis Charts
# ---------------------------------------------------------------------------

BAR_MAX_WIDTH = 44
FEW_BARS_MAX = 10

VALUE_LABEL_COLOR = "#1E293B"
VALUE_LABEL_SIZE_FEW = 12
VALUE_LABEL_SIZE_MANY = 9.5
MANY_BARS_MIN = 20


def _chart_theme():
    """Returns (chart_color, accent_color, text_color, grid_color, domain_color, value_label_color)
    adapted to st.session_state.dark_mode."""
    if bool(st.session_state.get("dark_mode", False)):
        return "#3B82F6", "#F97316", "#94A3B8", "#334155", "#475569", "#F8FAFC"
    return CHART_COLOR, ACCENT_COLOR, CHART_TEXT, CHART_GRID, "#DBE2EA", VALUE_LABEL_COLOR


def _bar_chart(counts, column, horizontal=False):
    """A bar chart of counts (from dashboard_data.count_by), keeping the order the table is
    already in. Every bar carries its exact value above it."""
    import altair as alt

    if horizontal:
        _horizontal_bar_chart(counts, column)
        return

    chart_color, accent_color, text_color, grid_color, domain_color, value_label_color = _chart_theme()
    tilted = len(counts) > STRAIGHT_LABELS_MAX_BARS
    x_axis = alt.Axis(
        labelAngle=-45 if tilted else 0,
        labelAlign="right" if tilted else "center",
        labelBaseline="top",
        labelOverlap="greedy",
        labelLimit=160,
        labelPadding=6,
    )
    x = alt.X(f"{column}:N", sort=list(counts[column]), title=None,
              scale=alt.Scale(paddingInner=0.28, paddingOuter=0.25), axis=x_axis)
    y = alt.Y("Defects:Q", title=None, axis=alt.Axis(tickMinStep=1))

    bars = (
        alt.Chart(counts)
        .mark_bar(cornerRadiusTopLeft=4, cornerRadiusTopRight=4, **({"size": BAR_MAX_WIDTH} if len(counts) <= FEW_BARS_MAX else {}))
        .encode(
            color=alt.condition(
                f"datum['{column}'] == '{UNKNOWN}'",
                alt.value(accent_color),
                alt.value(chart_color),
            ),
            x=x,
            y=y,
            tooltip=[alt.Tooltip(f"{column}:N", title=column), alt.Tooltip("Defects:Q", title="Number of defects")],
        )
    )
    labels = (
        alt.Chart(counts)
        .mark_text(
            align="center", baseline="bottom", dy=-4,
            fontSize=VALUE_LABEL_SIZE_FEW if len(counts) <= MANY_BARS_MIN else VALUE_LABEL_SIZE_MANY,
            fontWeight=600, color=value_label_color,
        )
        .encode(x=x, y=y, text=alt.Text("Defects:Q", format=",d"))
    )

    chart = (
        alt.layer(bars, labels)
        .properties(
            height=220, width="container", background=CHART_BACKGROUND,
            padding={"left": 4, "right": 12, "top": 22, "bottom": 4},
        )
        .configure(font=CHART_FONT)
        .configure_axis(
            labelColor=text_color, labelFontSize=12,
            titleFontSize=13,
            gridColor=grid_color, domainColor=domain_color, tickColor=domain_color,
        )
        .configure_view(stroke=None)
    )
    try:
        st.altair_chart(chart, theme=None, use_container_width=True)
    except TypeError:
        st.altair_chart(chart, theme=None)


def _horizontal_bar_chart(counts, column):
    """Horizontal bars for categories with long names (the likely causes): the names sit in full on
    the left, one bar per row, every bar with its exact value at the end."""
    import altair as alt

    chart_color, accent_color, text_color, grid_color, domain_color, value_label_color = _chart_theme()
    y = alt.Y(f"{column}:N", sort=list(counts[column]), title=None,
              scale=alt.Scale(paddingInner=0.35, paddingOuter=0.2),
              axis=alt.Axis(labelLimit=320, labelPadding=10, labelFontSize=13, ticks=False, domain=False))
    x = alt.X("Defects:Q", title=None, axis=alt.Axis(tickMinStep=1))

    bars = (
        alt.Chart(counts)
        .mark_bar(cornerRadiusTopRight=4, cornerRadiusBottomRight=4, size=26)
        .encode(
            color=alt.condition(
                f"datum['{column}'] == '{UNKNOWN}'",
                alt.value(accent_color),
                alt.value(chart_color),
            ),
            y=y,
            x=x,
            tooltip=[alt.Tooltip(f"{column}:N", title=column), alt.Tooltip("Defects:Q", title="Number of defects")],
        )
    )
    labels = (
        alt.Chart(counts)
        .mark_text(align="left", baseline="middle", dx=6, fontSize=VALUE_LABEL_SIZE_FEW, fontWeight=600,
                   color=value_label_color)
        .encode(y=y, x=x, text=alt.Text("Defects:Q", format=",d"))
    )

    chart = (
        alt.layer(bars, labels)
        .properties(
            height=max(120, 52 * len(counts) + 16), width="container", background=CHART_BACKGROUND,
            padding={"left": 4, "right": 40, "top": 8, "bottom": 4},
        )
        .configure(font=CHART_FONT)
        .configure_axis(
            labelColor=text_color, labelFontSize=12,
            titleFontSize=13,
            gridColor=grid_color, domainColor=domain_color, tickColor=domain_color,
        )
        .configure_view(stroke=None)
    )
    try:
        st.altair_chart(chart, theme=None, use_container_width=True)
    except TypeError:
        st.altair_chart(chart, theme=None)


def _chart_badge_text(counts, column):
    """Contextual right-hand badge for each chart card."""
    if counts.empty:
        return "0 total"
    if column == "Tower":
        known = counts[counts["Tower"] != UNKNOWN]
        n = len(known) if not known.empty else len(counts)
        return f"{n:,} tower{'s' if n != 1 else ''}"
    if column == "Floor":
        known = counts[counts["Floor"] != UNKNOWN]
        target = known if not known.empty else counts
        top_row = target.sort_values("Defects", ascending=False).iloc[0]
        return f"Floor {top_row['Floor']} is highest"
    if column == "Cause":
        known = counts[counts["Cause"] != UNKNOWN]
        target = known if not known.empty else counts
        top_row = target.sort_values("Defects", ascending=False).iloc[0]
        return f"Mostly {top_row['Cause']}"
    total = int(counts["Defects"].sum())
    return f"{total:,} total"


def _chart_section(df, column, title, hide_unknown):
    """A card with a title, a one-line subtitle, a contextual right badge, and the bar chart."""
    counts = count_by(df, column)
    if hide_unknown:
        counts = counts[counts[column] != UNKNOWN]

    has_unknown = not hide_unknown and not counts[counts[column] == UNKNOWN].empty
    badge_text = _chart_badge_text(counts, column)

    with st.container(border=True, key=f"chart_{column.lower()}"):
        legend_html = (
            '<span style="display:inline-block;width:8px;height:8px;border-radius:2px;'
            'background:#f97316;margin-right:5px;"></span>Unknown classification'
            if has_unknown else ""
        )
        st.markdown(
            '<div class="chart-title"><div>'
            f'<div class="chart-name">{html.escape(title)}</div>'
            f'<div class="chart-sub">{html.escape(CHART_SUBTITLES.get(column, ""))}</div>'
            '</div>'
            '<div style="text-align:right;">'
            f'<span class="chart-total">{html.escape(badge_text)}</span>'
            f'<div style="color:var(--text-muted);font-size:11.5px;font-weight:500;margin-top:2px;">{legend_html}</div>'
            '</div></div>',
            unsafe_allow_html=True,
        )
        if counts.empty:
            st.caption(f"Nothing to show: every defect here has an unknown {column.lower()}.")
            return
        _bar_chart(counts, column, horizontal=(column == "Cause"))


def _chart_notes(df, hide_unknown):
    """Short note when 'Exclude Unknown locations' is switched on."""
    columns = ("Tower", "Material", "Cause")
    if hide_unknown:
        parts = []
        for column in columns:
            counts = count_by(df, column)
            hidden = int(counts.loc[counts[column] == UNKNOWN, "Defects"].sum())
            if hidden:
                parts.append(f"{hidden:,} defect(s) with an unknown {column.lower()} are hidden.")
        return " ".join(parts)
    return ""


def _render_raw_audit_log(df):
    """Draws the bottom 'Raw audit log' card."""
    ordered = [c for c in TABLE_COLUMNS if c in df.columns] + [
        c for c in df.columns if c not in TABLE_COLUMNS and c != "File Link"
    ] + (["File Link"] if "File Link" in df.columns else [])
    shown = df[ordered].head(DATA_TABLE_MAX_ROWS)
    table = _data_table_html(shown, float_format=lambda value: f"{value:g}")
    cap_note = (
        f'<p class="audit-cap-note">Showing the first {DATA_TABLE_MAX_ROWS:,} of {len(df):,} defects.</p>'
        if len(df) > DATA_TABLE_MAX_ROWS else ""
    )

    audit_html = (
        '<details class="audit-drawer">'
        '<summary class="audit-drawer-head">'
        '<div class="audit-drawer-left">'
        f'<span class="audit-drawer-icon">{_ICON_AUDIT_TABLE}</span>'
        '<div>'
        '<div class="audit-drawer-title">Raw audit log</div>'
        '<div class="audit-drawer-sub">View the records behind these quantities and charts</div>'
        '</div></div>'
        f'<div class="audit-drawer-cta">View data {_ICON_CHEVRON_DOWN}</div>'
        '</summary>'
        '<div class="audit-drawer-body">'
        f'<div class="data-table-wrap">{table}</div>{cap_note}'
        '</div></details>'
    )
    st.markdown(audit_html, unsafe_allow_html=True)


# ---------------------------------------------------------------------------
# Sidebar: which files, refresh, filters (including strict Defect sub-filter per Substrate)
# ---------------------------------------------------------------------------

def _filter_key(name, scope):
    return f"flt_{name}_{scope}"


def _sorted_options(series):
    return sorted(series.unique(), key=natural_key)


def _defect_options_for_substrates(selected_substrates):
    """Returns the strict canonical defect codes for the selected substrates in exact order:
      - Concrete: C-, C+, BH, US
      - Sealant:  DS CC, DS CF, DS GG, DS FG, DS FF, MS CC, MS CF, MS GG, MS FG, MS FF
      - Paint:    DP, FP, BP, NP
      - Gasket:   DG
    """
    substrates = selected_substrates if selected_substrates else list(SUBSTRATE_DEFECT_OPTIONS.keys())
    options = []
    for sub in substrates:
        for code in SUBSTRATE_DEFECT_OPTIONS.get(sub, []):
            if code not in options:
                options.append(code)
    return options


def render_dashboard_controls():
    """The dashboard's controls (the left panel)."""
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
        help=f'"{ALL_FILES_LABEL}" combines every Excel file in the landing folder and all of its folders. '
             'Selecting a parent folder automatically includes its subfolders.',
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

    try:
        rows, files_read, _skipped = _load_scope(scope, subfolders)
    except Exception:
        return
    df = rows_to_dataframe(rows)
    if df.empty:
        return

    st.divider()
    st.markdown('<p class="side-title">Filters</p>', unsafe_allow_html=True)
    st.caption("Leave a filter empty to include everything.")
    st.multiselect("Tower", _sorted_options(df["Tower"]), key=_filter_key("tower", scope))
    st.multiselect("Floor", _sorted_options(df["Floor"]), key=_filter_key("floor", scope))
    selected_substrates = st.multiselect(
        "Substrate",
        _sorted_options(df["Material"]),
        key=_filter_key("material", scope),
    )

    # Sub-filter for Defect Codes belonging to the selected Substrate(s)
    defect_key = _filter_key("defect", scope)
    available_defects = _defect_options_for_substrates(selected_substrates)
    current_defects = st.session_state.get(defect_key) or []
    valid_current = [d for d in current_defects if d in available_defects]
    if valid_current != current_defects:
        st.session_state[defect_key] = valid_current

    defect_label = (
        f"Defect ({', '.join(selected_substrates)})"
        if selected_substrates and len(selected_substrates) <= 2
        else ("Defect (selected substrates)" if selected_substrates else "Defect")
    )
    st.multiselect(
        defect_label,
        available_defects,
        key=defect_key,
        help="Choose specific defect codes for the selected substrate(s).",
    )

    folders_in_data = _sorted_options(df["Folder"])
    if len(folders_in_data) > 1:
        st.multiselect("Folder", folders_in_data, key=_filter_key("folder", scope))


# ---------------------------------------------------------------------------
# Main area: Engineering Overview -> Repair Quantity Summary -> Location Analysis -> Raw Audit Log
# ---------------------------------------------------------------------------

@st.fragment(run_every=DASHBOARD_REFRESH_SECONDS)
def render_analytics_dashboard():
    """Re-runs every DASHBOARD_REFRESH_SECONDS without reloading the rest of the page."""
    ui_components.section_heading("Engineering overview", "Facade defect dashboard")

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
        scope = ALL_FILES_LABEL
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
        active_filters.append(f"Substrate: {', '.join(selected_materials)}")
    selected_defects = st.session_state.get(_filter_key("defect", scope)) or []
    if selected_defects:
        active_filters.append(f"Defect: {', '.join(selected_defects)}")
    folders = st.session_state.get(_filter_key("folder", scope)) or []
    if folders:
        active_filters.append(f"Folder: {', '.join(folders)}")
    if hide_unknown:
        active_filters.append("Exclude unknown")

    if active_filters:
        chips_html = "".join([f'<span class="filter-chip">{html.escape(f)} &times;</span>' for f in active_filters])
        st.markdown(f'<div class="filter-chips-container">{chips_html}</div>', unsafe_allow_html=True)
        if st.button("Clear all filters", key="clear_all_filters", help="Reset all applied filters"):
            st.session_state[_filter_key("tower", scope)] = []
            st.session_state[_filter_key("floor", scope)] = []
            st.session_state[_filter_key("material", scope)] = []
            st.session_state[_filter_key("defect", scope)] = []
            st.session_state[_filter_key("folder", scope)] = []
            st.session_state[HIDE_UNKNOWN_KEY] = False
            st.rerun()

    # --- Apply sidebar filters ---
    if folders:
        df = df[df["Folder"].isin(folders)]
    df = apply_filters(df, selected_towers, selected_floors, selected_materials, selected_defects)

    summary = dimension_by_defect(df, expand_all_sealant=not bool(selected_defects))

    # 1. Top Row: 4 Engineering KPI Cards
    _render_kpi_cards(df, summary, files_read)

    if df.empty:
        st.info("No defects match these filters.")
        return

    # 2. Warm-Amber Collapsible Review Banner (when dimensions are missing)
    _render_missing_dimensions_queue(df)

    # 3. Primary Centerpiece: Repair Quantity Summary ("Dimensions by defect")
    _dimension_section(summary, scope=scope)

    # 4. Secondary Section: Location Analysis ("Where are the defects located?")
    known_towers = df[df["Tower"] != UNKNOWN]["Tower"].nunique() or df["Tower"].nunique()
    known_floors = df[df["Floor"] != UNKNOWN]["Floor"].nunique() or df["Floor"].nunique()
    loc_summary = (
        f"{len(df):,} defect{'s' if len(df) != 1 else ''} across "
        f"{known_towers:,} tower{'s' if known_towers != 1 else ''} and "
        f"{known_floors:,} active floor{'s' if known_floors != 1 else ''}"
    )
    st.markdown(
        '<div class="loc-section-head"><div>'
        '<p class="eyebrow">Location analysis</p>'
        '<h2>Where are the defects located?</h2>'
        f'</div><span class="loc-section-meta">{html.escape(loc_summary)}</span></div>',
        unsafe_allow_html=True,
    )

    with st.container(key="charts_grid"):
        material_col, tower_col = st.columns([1, 1], gap="small")
        with material_col:
            _chart_section(df, "Material", "Defects by substrate", hide_unknown)
        with tower_col:
            _chart_section(df, "Tower", "Defects by tower", hide_unknown)
        _chart_section(df, "Cause", "Defects by likely cause", hide_unknown)

    note = _chart_notes(df, hide_unknown)
    if note:
        st.markdown(f'<div class="chart-legend">{html.escape(note)}</div>', unsafe_allow_html=True)

    # 5. Bottom Drawer: Raw Audit Log
    _render_raw_audit_log(df)


def render_dashboard_view():
    """The whole Dashboard view: the controls panel and the main analytics area side by side."""
    col_controls, col_main = st.columns([0.22, 0.78], gap="small")
    with col_controls:
        with st.container(key="dashboard_sidebar"):
            render_dashboard_controls()
    with col_main:
        with st.container(key="dashboard_main"):
            with st.container(key="analytics"):
                render_analytics_dashboard()
