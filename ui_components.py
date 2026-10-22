import base64
import html
from pathlib import Path
import streamlit as st

from icons import icon

# Icons come from icons.py (Lucide), so they all share one style
_ICON_LAYERS = icon("layers", 23)
_ICON_SPARK = icon("sparkles", 18)

# Gray file icon on the right of the "Save settings" card heading
_ICON_FILE = icon("file-text", 18, stroke=1.75, color="#9ba8b9", hidden=True)

# The "?" icon of the Help link in the header
_ICON_HELP = icon("circle-question-mark", 15, stroke=2, hidden=True)

# The how-to text that opens from the Help link.
_HELP_HTML = (
    '<div class="help-title">How to use the whiteboard extractor</div>'
    '<p>Add whiteboard photos from your device or from a SharePoint folder. The AI extracts the data '
    'and appends it to the Master Excel file.</p>'
    '<ol>'
    '<li>Choose where to import from: <b>Upload from device</b> or <b>Import from SharePoint</b>, then select your photos or folder.</li>'
    '<li>Choose the SharePoint folder to save into and give the batch a name. The name becomes the Excel file name.</li>'
    '<li>Click the <b>Extract</b> button.</li>'
    '<li>Follow the progress, results and download on the right. The <b>Dashboard</b> picks up the new data by itself.</li>'
    '</ol>'
)


def load_local_css(file_name="style.css"):
    """Injects the external CSS file into the Streamlit app layout and strips non-breaking spaces."""
    css_path = Path(__file__).with_name(file_name)
    if not css_path.exists():
        css_path = Path(file_name)
    try:
        css = css_path.read_text(encoding="utf-8").replace("\u00a0", " ")
        st.markdown(f"<style>{css}</style>", unsafe_allow_html=True)
    except OSError:
        pass


def _brand_html():
    """The logo, a thin divider and the product name (left of the header bar)."""
    logo_path = Path(__file__).with_name("vonotec.png")
    if not logo_path.exists():
        logo_path = Path("vonotec.png")
    try:
        logo_base64 = base64.b64encode(logo_path.read_bytes()).decode()
        logo = f'<img src="data:image/png;base64,{logo_base64}" width="138" alt="VONOTEC">'
    except OSError:
        logo = '<span class="brand-text">VONOTEC</span>'
    return f'<div class="brand">{logo}<span class="product-name">Automated Defect Logger</span></div>'


def render_header(views):
    """One white header card with three equal columns: the brand on the left,
    the view switch (a radio, drawn as two pills) in the middle and the tagline on the right."""
    with st.container(key="top_app_header"):
        try:
            col_left, col_center, col_right = st.columns([1, 1, 1], vertical_alignment="center")
        except TypeError:
            col_left, col_center, col_right = st.columns([1, 1, 1])

        with col_left:
            st.markdown(_brand_html(), unsafe_allow_html=True)
        with col_center:
            view = st.radio(
                "View",
                views,
                horizontal=True,
                label_visibility="collapsed",
                key="nav_view_switcher",
            )
        with col_right:
            st.markdown(
                '<div class="header-right">'
                '<span class="header-tagline">Defect Extraction & Repair Analytics</span>'
                f'<details class="help-pop"><summary>{_ICON_HELP}Help</summary>'
                f'<div class="help-panel">{_HELP_HTML}</div></details>'
                '</div>',
                unsafe_allow_html=True,
            )
    return view


WORKFLOW_STEPS = ("Select photos", "Save destination", "Extract data")


def render_stepper(done, current):
    """Renders a clean, transparent '1. Select photos > 2. Save destination > 3. Extract data' line
    directly where the eyebrow sits, with no white background or circles."""
    parts = []
    for number, label in enumerate(WORKFLOW_STEPS, start=1):
        state = "done" if number in done else ("current" if number == current else "todo")
        prefix = f"✓ {number}." if state == "done" else f"{number}."
        if state == "current":
            color = "#1e3a8a"
            weight = "700"
        elif state == "done":
            color = "#15803d"
            weight = "600"
        else:
            color = "#64748b"
            weight = "500"
        parts.append(
            f'<span style="color:{color};font-weight:{weight};background:transparent !important;'
            f'padding:0 !important;border:none !important;box-shadow:none !important;">'
            f'{prefix} {html.escape(label)}</span>'
        )
    sep = '<span style="color:#94a3b8;margin:0 8px;font-weight:500;background:transparent !important;">&gt;</span>'
    st.markdown(
        '<style>'
        '.st-key-top_app_header { margin-bottom: 18px !important; }'
        '</style>'
        '<div style="display:flex;align-items:center;flex-wrap:wrap;font-size:13px;'
        'line-height:1.4;background:transparent !important;border:none !important;'
        'box-shadow:none !important;padding:0 !important;margin:0 0 4px 0 !important;">'
        + sep.join(parts)
        + '</div>',
        unsafe_allow_html=True,
    )


def section_heading(eyebrow, title, badge=None, step=None, description=None):
    """A small orange label over a large title. On the right there is either a status badge,
    badge = (text, kind) with kind "ready", "processing" or "complete", or a plain step label."""
    right = ""
    if badge:
        text, kind = badge
        right = f'<span class="status-badge {html.escape(kind)}"><span></span>{html.escape(text)}</span>'
    elif step:
        right = f'<span class="step-label">{html.escape(step)}</span>'
    eyebrow_html = f'<p class="eyebrow">{html.escape(eyebrow)}</p>' if eyebrow else ""
    desc_html = f'<p class="section-desc">{html.escape(description)}</p>' if description else ""
    st.markdown(
        '<div class="section-heading"><div>'
        f'{eyebrow_html}'
        f'<h2>{html.escape(title)}</h2>'
        f'</div>{right}</div>{desc_html}',
        unsafe_allow_html=True,
    )


def subsection_heading(title, description=None):
    """The smaller heading of the source-specific content under the source switch."""
    desc_html = f'<div class="source-subhead-desc">{html.escape(description)}</div>' if description else ""
    st.markdown(
        f'<div class="source-subhead"><div class="source-subhead-title">{html.escape(title)}</div>{desc_html}</div>',
        unsafe_allow_html=True,
    )


def card_heading(eyebrow, title, icon=_ICON_FILE, step=None):
    """The same label-over-title pair, smaller, for the top of a card."""
    step_html = f'<span class="step-label">{html.escape(step)}</span>' if step else ""
    st.markdown(
        '<div class="card-heading"><div>'
        f'<p class="eyebrow">{html.escape(eyebrow)}</p>'
        f'<div class="card-title">{html.escape(title)}</div>'
        f'</div><div class="card-heading-right">{step_html}{icon}</div></div>',
        unsafe_allow_html=True,
    )


# Small icons for the "After extraction, you'll see" preview
_ICON_PREVIEW_STATUS = icon("circle-check", 18)
_ICON_PREVIEW_REVIEW = icon("triangle-alert", 18)
_ICON_PREVIEW_FILE = icon("file-spreadsheet", 18)


def render_empty_results():
    """Shown in the Batch results column until the first batch is run."""
    st.markdown(
        '<div class="sync-card">'
        f'<span class="sync-icon">{_ICON_LAYERS}</span>'
        '<div><strong>Ready to process a new batch</strong>'
        '<p>Messages, downloads and sync status appear here once you extract a batch.</p></div>'
        '</div>'
        '<div class="preview-card">'
        '<p class="preview-title">After extraction, you\'ll see:</p>'
        '<ul class="preview-list">'
        f'<li><span class="preview-icon">{_ICON_PREVIEW_STATUS}</span>'
        '<div><b>File-level status</b><span>Added, skipped or failed, one row per photo</span></div></li>'
        f'<li><span class="preview-icon preview-icon-warn">{_ICON_PREVIEW_REVIEW}</span>'
        '<div><b>Records needing review</b><span>Rows highlighted red in Excel, with the reason</span></div></li>'
        f'<li><span class="preview-icon">{_ICON_PREVIEW_FILE}</span>'
        '<div><b>Generated Excel workbook</b><span>Download it, or find it in SharePoint</span></div></li>'
        '</ul></div>',
        unsafe_allow_html=True,
    )


def render_tip(completed=False, review_count=0):
    """Contextual footer card in the results panel:
    - Before extraction: photo quality tip
    - After extraction with flagged rows: explains how to review red rows in Excel
    - After extraction with 0 flagged rows: confirms all rows passed validation"""
    if completed and review_count > 0:
        st.markdown(
            '<div class="tip-card">'
            f'<span class="tip-icon">{_ICON_PREVIEW_REVIEW}</span>'
            f'<div><strong>How to review {review_count} flagged item{"s" if review_count != 1 else ""}</strong>'
            '<p>Flagged items were added to the Excel workbook and highlighted in <b>red</b>. '
            'Click any row above to see what was missing or shared, then open the workbook in SharePoint '
            'or download it above to check the embedded photo and fill in the cell.</p></div>'
            '</div>',
            unsafe_allow_html=True,
        )
    elif completed:
        st.markdown(
            '<div class="tip-card tip-card-ok">'
            f'<span class="tip-icon tip-icon-ok">{_ICON_PREVIEW_STATUS}</span>'
            '<div><strong>All extracted rows passed validation</strong>'
            '<p>Every processed whiteboard had complete defect codes and dimensions. '
            'The workbook is synced to SharePoint and ready in the Analytics Dashboard.</p></div>'
            '</div>',
            unsafe_allow_html=True,
        )
    else:
        st.markdown(
            '<div class="tip-card">'
            f'<span class="tip-icon">{_ICON_SPARK}</span>'
            '<div><strong>Extraction tip</strong>'
            '<p>For best results, use clear photos with the full whiteboard in frame.</p></div>'
            '</div>',
            unsafe_allow_html=True,
        )


def render_instructions():
    """The how-to text above the uploader (Whiteboard Extraction view only),
    folded away by default so it doesn't use up the top of the screen."""
    with st.expander("How to use the whiteboard extractor", expanded=False):
        st.markdown(
            """
            <div class="instruction-text">
            Add whiteboard photos from your device or from a SharePoint folder. The AI will
            automatically extract the data and append it to the Master Excel file.
            <ol>
                <li>Choose where to import from: <b>Upload from device</b> or <b>Import from SharePoint</b>, then select your photos or folder.</li>
                <li>Choose the SharePoint folder to save into and give the batch a name. The name becomes the Excel file name.</li>
                <li>Click <b>Extract photos</b>.</li>
                <li>Follow the progress, messages and download on the right. The <b>Analytics Dashboard</b> picks up the new data by itself.</li>
            </ol>
            </div>
            """,
            unsafe_allow_html=True,
        )